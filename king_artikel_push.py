"""
KING Artikel Push — directe push naar KING Webservices API.
Slaat elke push op in een SQLite auditlog (push_history.db).

Volgorde per artikel:
  1. Artikel_ToevoegenWijzigen       — basisvelden
  2. Artikel_Omschrijving_ToevoegenWijzigen — lange omschrijving per taalcode
  3. Artikel_VrijeRubriek_Wijzigen   — alle VR_ART_* velden
  4. Artikel_Leverancier_ToevoegenWijzigen  — als Leveranciernummer gevuld
  5. Artikel_Ean_ToevoegenWijzigen   — als EanCode gevuld
"""
from __future__ import annotations

import getpass
import json
import sqlite3
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import requests

# ── Veld-routing ──────────────────────────────────────────────────────────────

# (bron-kolom, KING-veldnaam) voor Artikel_ToevoegenWijzigen
BASIS_MAP: list[tuple[str, str]] = [
    ("Artikelnummer", "ArtikelNummer"),
    ("Zoekcode",      "ZoekCode"),
    ("Opbrengstgroep","OpbrengstGroepNummer"),
    ("WebArtikel",    "WebArtikel"),
    ("AfbeeldingKlein","AfbeeldingKlein"),
    ("AfbeeldingGroot","AfbeeldingGroot"),
]

# Lange omschrijving per taalcode: bron-kolom → TaalCode
TAAL_MAP: dict[str, str] = {
    "1NL_FOV_NL_LANGE_OMSCHRIJVING": "NLD",
}

# Vrije rubriek prefixen — VR_ART_* plus de 1N*-tekstvelden (zijn vrije rubrieken in KING)
# 1NL_ is uitgesloten: die gaat via Artikel_Omschrijving_ToevoegenWijzigen (stap 2)
VR_PREFIXES = ("VR_ART_", "1NT_", "1NF_", "1NH_", "1N1_", "1N2_")

# Velden die NIET via Webservices gepushed worden
SKIP_VELDEN: set[str] = {
    # Geen eigen endpoint of alleen relevant voor Excel-import
    "Eenheid", "TekstOpFactuur",
    "Leveranciernaam", "ArtikelOmschrijvingLeverancier", "ArtikelNummerBijLeverancier",
    "Omschrijving",           # gaat als OmschrijvingKort (max 40 tekens) mee in basis
    # Controle-/confidence-kolommen
    "Attribuutset_Confidence_%", "Attribuutset_Label",
    "Webcategorie_Confidence_%",
}


# ── Configuratie ──────────────────────────────────────────────────────────────

@dataclass
class KingConfig:
    protocol: str
    host: str
    port: str
    administratie: str
    access_token: str
    sql_timeout: str = ""
    verify_ssl: bool = False

    def endpoint(self, operatie: str) -> str:
        host = self.host.split("://")[-1].rstrip("/")
        hp = host if ":" in host else f"{host}:{self.port}"
        return f"{self.protocol}://{hp}/{self.administratie}/{operatie}"

    @classmethod
    def from_dict(cls, d: dict) -> "KingConfig":
        return cls(
            protocol=d.get("protocol", "http").lower(),
            host=d.get("host", ""),
            port=str(d.get("port", "8082")),
            administratie=d.get("administratie", ""),
            access_token=d.get("access_token", ""),
            sql_timeout=d.get("sql_timeout", ""),
            verify_ssl=bool(d.get("verify_ssl", False)),
        )

    def is_geldig(self) -> bool:
        return bool(self.host and self.access_token and self.administratie)


# ── SQLite auditlog ───────────────────────────────────────────────────────────

_SCHEMA = """
CREATE TABLE IF NOT EXISTS push_sessie (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    tijdstip    TEXT NOT NULL,
    gebruiker   TEXT NOT NULL,
    dry_run     INTEGER DEFAULT 0,
    n_artikelen INTEGER DEFAULT 0,
    n_gelukt    INTEGER DEFAULT 0,
    n_mislukt   INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS push_artikel (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    sessie_id    INTEGER NOT NULL REFERENCES push_sessie(id),
    artikelnummer TEXT NOT NULL,
    omschrijving  TEXT,
    tijdstip      TEXT NOT NULL,
    status        TEXT NOT NULL,
    n_calls       INTEGER DEFAULT 0,
    n_fouten      INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS push_call (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    artikel_id   INTEGER NOT NULL REFERENCES push_artikel(id),
    tijdstip     TEXT NOT NULL,
    operatie     TEXT NOT NULL,
    veld_naam    TEXT,
    nieuwe_waarde TEXT,
    oude_waarde   TEXT,
    http_status   TEXT,
    king_status   INTEGER,
    foutcode      TEXT,
    foutmelding   TEXT
);
"""


class PushDB:
    def __init__(self, pad: Path):
        self.pad = pad
        self.pad.parent.mkdir(parents=True, exist_ok=True)
        with self._con() as con:
            con.executescript(_SCHEMA)
            # Migratie: dry_run kolom toevoegen aan bestaande DB's
            cols = {r[1] for r in con.execute("PRAGMA table_info(push_sessie)")}
            if "dry_run" not in cols:
                con.execute("ALTER TABLE push_sessie ADD COLUMN dry_run INTEGER DEFAULT 0")

    def _con(self) -> sqlite3.Connection:
        con = sqlite3.connect(self.pad)
        con.row_factory = sqlite3.Row
        return con

    # ── Sessie ──
    def nieuwe_sessie(self, gebruiker: str, dry_run: bool = False) -> int:
        with self._con() as con:
            cur = con.execute(
                "INSERT INTO push_sessie (tijdstip, gebruiker, dry_run) VALUES (?, ?, ?)",
                (_nu(), gebruiker, 1 if dry_run else 0),
            )
            return cur.lastrowid

    def update_sessie(self, sessie_id: int, n_artikelen: int, n_gelukt: int, n_mislukt: int):
        with self._con() as con:
            con.execute(
                "UPDATE push_sessie SET n_artikelen=?, n_gelukt=?, n_mislukt=? WHERE id=?",
                (n_artikelen, n_gelukt, n_mislukt, sessie_id),
            )

    # ── Artikel ──
    def nieuw_artikel(self, sessie_id: int, artikelnummer: str, omschrijving: str) -> int:
        with self._con() as con:
            cur = con.execute(
                "INSERT INTO push_artikel (sessie_id, artikelnummer, omschrijving, tijdstip, status)"
                " VALUES (?, ?, ?, ?, 'bezig')",
                (sessie_id, artikelnummer, omschrijving, _nu()),
            )
            return cur.lastrowid

    def update_artikel(self, artikel_id: int, status: str, n_calls: int, n_fouten: int):
        with self._con() as con:
            con.execute(
                "UPDATE push_artikel SET status=?, n_calls=?, n_fouten=? WHERE id=?",
                (status, n_calls, n_fouten, artikel_id),
            )

    # ── Call ──
    def log_call(
        self, artikel_id: int, operatie: str, veld_naam: str | None,
        nieuwe_waarde: str, oude_waarde: str | None,
        http_status: str, king_status: int | None,
        foutcode: str | None, foutmelding: str | None,
    ):
        with self._con() as con:
            con.execute(
                "INSERT INTO push_call"
                " (artikel_id, tijdstip, operatie, veld_naam, nieuwe_waarde, oude_waarde,"
                "  http_status, king_status, foutcode, foutmelding)"
                " VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (artikel_id, _nu(), operatie, veld_naam, nieuwe_waarde, oude_waarde,
                 http_status, king_status, foutcode, foutmelding),
            )

    def haal_vorige_waarde(
        self, artikelnummer: str, operatie: str, veld_naam: str | None
    ) -> str | None:
        """Laatste succesvolle waarde voor dit artikel+veld — wordt 'oude waarde' in de diff."""
        with self._con() as con:
            if veld_naam is None:
                cur = con.execute(
                    "SELECT pc.nieuwe_waarde FROM push_call pc"
                    " JOIN push_artikel pa ON pa.id = pc.artikel_id"
                    " WHERE pa.artikelnummer = ? AND pc.operatie = ? AND pc.veld_naam IS NULL"
                    "   AND pc.king_status = 0 ORDER BY pc.id DESC LIMIT 1",
                    (artikelnummer, operatie),
                )
            else:
                cur = con.execute(
                    "SELECT pc.nieuwe_waarde FROM push_call pc"
                    " JOIN push_artikel pa ON pa.id = pc.artikel_id"
                    " WHERE pa.artikelnummer = ? AND pc.operatie = ? AND pc.veld_naam = ?"
                    "   AND pc.king_status = 0 ORDER BY pc.id DESC LIMIT 1",
                    (artikelnummer, operatie, veld_naam),
                )
            row = cur.fetchone()
            return row[0] if row else None

    # ── Queries voor overzicht ──
    def haal_sessies(self, limit: int = 100) -> list[dict]:
        with self._con() as con:
            cur = con.execute(
                "SELECT * FROM push_sessie ORDER BY id DESC LIMIT ?", (limit,)
            )
            return [dict(r) for r in cur.fetchall()]

    def haal_artikelen(self, sessie_id: int) -> list[dict]:
        with self._con() as con:
            cur = con.execute(
                "SELECT * FROM push_artikel WHERE sessie_id = ? ORDER BY id", (sessie_id,)
            )
            return [dict(r) for r in cur.fetchall()]

    def haal_calls(self, artikel_id: int) -> list[dict]:
        with self._con() as con:
            cur = con.execute(
                "SELECT * FROM push_call WHERE artikel_id = ? ORDER BY id", (artikel_id,)
            )
            return [dict(r) for r in cur.fetchall()]

    def haal_artikel_detail(self, artikelnummer: str, limit: int = 20) -> list[dict]:
        """Alle push-sessies voor een specifiek artikelnummer, recentste eerst."""
        with self._con() as con:
            cur = con.execute(
                "SELECT pa.*, ps.tijdstip AS sessie_tijd, ps.gebruiker"
                " FROM push_artikel pa JOIN push_sessie ps ON ps.id = pa.sessie_id"
                " WHERE pa.artikelnummer = ? ORDER BY pa.id DESC LIMIT ?",
                (artikelnummer, limit),
            )
            return [dict(r) for r in cur.fetchall()]


# ── HTTP helper ───────────────────────────────────────────────────────────────

def _king_post(
    cfg: KingConfig, operatie: str, payload: dict
) -> tuple[int | None, str | None, str, str]:
    """POST naar KING. Returns (king_status, foutcode, foutmelding, http_status)."""
    headers = {
        "ACCESS-TOKEN": cfg.access_token,
        "Content-Type": "application/json",
    }
    if cfg.sql_timeout:
        headers["SQL-TIMEOUT"] = cfg.sql_timeout

    try:
        resp = requests.post(
            cfg.endpoint(operatie),
            headers=headers,
            json=payload,
            timeout=30,
            verify=cfg.verify_ssl,
        )
        http_status = str(resp.status_code)
        try:
            data = resp.json()
        except ValueError:
            return None, None, resp.text[:500], http_status

        if isinstance(data, list):
            data = data[0] if data else {}
        if not isinstance(data, dict):
            return None, None, str(data)[:500], http_status

        return (
            data.get("Status"),
            str(data.get("FoutCode") or ""),
            str(data.get("FoutMelding") or ""),
            http_status,
        )
    except requests.RequestException as e:
        return None, None, str(e)[:500], "—"


# ── Push logica ───────────────────────────────────────────────────────────────

def push_artikel(
    waarden: dict[str, Any],
    cfg: KingConfig,
    db: PushDB,
    sessie_id: int,
    log_func=None,
    dry_run: bool = False,
) -> tuple[bool, int, int]:
    """
    Push één artikel naar KING. Returns (volledig_gelukt, n_calls, n_fouten).
    Stap 1 (basis aanmaken) is blokkerend: bij fout worden de vervolgstappen overgeslagen.
    Bij dry_run=True worden payloads opgebouwd en gelogd maar niet verstuurd.
    """
    def _log(msg: str):
        if log_func:
            log_func(msg)

    def _post(operatie: str, payload: dict) -> tuple[int | None, str | None, str, str]:
        if dry_run:
            return 0, None, "", "dry-run"
        return _king_post(cfg, operatie, payload)

    artikelnummer = str(waarden.get("Artikelnummer", "") or "").strip()
    omschrijving  = str(waarden.get("Omschrijving", "") or "").strip()
    if not artikelnummer:
        return False, 0, 0

    artikel_id = db.nieuw_artikel(sessie_id, artikelnummer, omschrijving)
    n_calls = 0
    n_fouten = 0

    # ── Stap 1: basisartikel ──────────────────────────────────────
    operatie = "Artikel_ToevoegenWijzigen"
    payload: dict[str, Any] = {"BestellenViaDirecteInkoop": 0}

    if omschrijving:
        payload["OmschrijvingKort"] = omschrijving[:40]

    for bron, king_naam in BASIS_MAP:
        waarde = str(waarden.get(bron, "") or "").strip()
        if waarde:
            payload[king_naam] = waarde

    basis_str = json.dumps(payload, ensure_ascii=False)
    oude_basis = db.haal_vorige_waarde(artikelnummer, operatie, None)

    king_st, fc, fm, hs = _post(operatie, payload)
    n_calls += 1
    db.log_call(artikel_id, operatie, None, basis_str, oude_basis, hs, king_st, fc, fm)

    if king_st != 0:
        _log(f"  {artikelnummer}: FOUT basis — {fm}")
        n_fouten += 1
        db.update_artikel(artikel_id, "mislukt", n_calls, n_fouten)
        return False, n_calls, n_fouten

    _log(f"  {artikelnummer}: basisartikel aangemaakt/bijgewerkt")

    # ── Stap 2: taalcode omschrijvingen ──────────────────────────
    for veld, taalcode in TAAL_MAP.items():
        tekst = str(waarden.get(veld, "") or "").strip()
        if not tekst:
            continue

        op_taal = "Artikel_Omschrijving_ToevoegenWijzigen"
        pl_taal = {"ArtikelNummer": artikelnummer, "TaalCode": taalcode, "Omschrijving": tekst}
        oude_taal = db.haal_vorige_waarde(artikelnummer, op_taal, taalcode)

        ks_t, fc_t, fm_t, hs_t = _post(op_taal, pl_taal)
        n_calls += 1
        db.log_call(artikel_id, op_taal, taalcode, tekst, oude_taal, hs_t, ks_t, fc_t, fm_t)

        if ks_t != 0:
            _log(f"  {artikelnummer}: FOUT omschrijving {taalcode} — {fm_t}")
            n_fouten += 1
        else:
            _log(f"  {artikelnummer}: omschrijving {taalcode} bijgewerkt")

    # ── Stap 3: vrije rubrieken ───────────────────────────────────
    for veld, waarde in waarden.items():
        if veld in SKIP_VELDEN:
            continue
        if not any(veld.startswith(p) for p in VR_PREFIXES):
            continue
        waarde_str = str(waarde or "").strip()
        if not waarde_str:
            continue

        op_vr = "Artikel_VrijeRubriek_Wijzigen"
        pl_vr = {
            "ArtikelNummer":      artikelnummer,
            "RubriekOmschrijving": veld[:40],
            "RubriekInhoud":      waarde_str,
        }
        oude_vr = db.haal_vorige_waarde(artikelnummer, op_vr, veld)

        ks_vr, fc_vr, fm_vr, hs_vr = _post(op_vr, pl_vr)
        n_calls += 1
        db.log_call(artikel_id, op_vr, veld, waarde_str, oude_vr, hs_vr, ks_vr, fc_vr, fm_vr)

        if ks_vr != 0:
            _log(f"  {artikelnummer}: FOUT rubriek {veld} — {fm_vr}")
            n_fouten += 1

    # ── Stap 4: leverancier ───────────────────────────────────────
    leveranciernummer = str(waarden.get("Leveranciernummer", "") or "").strip()
    if leveranciernummer:
        op_lev = "Artikel_Leverancier_ToevoegenWijzigen"
        pl_lev = {
            "ArtikelNummer":    artikelnummer,
            "LeverancierNummer": leveranciernummer,
            "DefaultLeverancier": 1,
        }
        oude_lev = db.haal_vorige_waarde(artikelnummer, op_lev, "LeverancierNummer")

        ks_lev, fc_lev, fm_lev, hs_lev = _post(op_lev, pl_lev)
        n_calls += 1
        db.log_call(artikel_id, op_lev, "LeverancierNummer", leveranciernummer, oude_lev,
                    hs_lev, ks_lev, fc_lev, fm_lev)

        if ks_lev != 0:
            _log(f"  {artikelnummer}: FOUT leverancier — {fm_lev}")
            n_fouten += 1
        else:
            _log(f"  {artikelnummer}: leverancier {leveranciernummer} gekoppeld")

    # ── Stap 5: EAN-code ──────────────────────────────────────────
    ean = str(waarden.get("EanCode", "") or "").strip()
    if ean:
        op_ean = "Artikel_Ean_ToevoegenWijzigen"
        pl_ean = {"ArtikelNummer": artikelnummer, "EanCode": ean, "DefaultEanCode": 1}
        oude_ean = db.haal_vorige_waarde(artikelnummer, op_ean, "EanCode")

        ks_ean, fc_ean, fm_ean, hs_ean = _post(op_ean, pl_ean)
        n_calls += 1
        db.log_call(artikel_id, op_ean, "EanCode", ean, oude_ean,
                    hs_ean, ks_ean, fc_ean, fm_ean)

        if ks_ean != 0:
            _log(f"  {artikelnummer}: FOUT EAN — {fm_ean}")
            n_fouten += 1
        else:
            _log(f"  {artikelnummer}: EAN {ean} gekoppeld")

    # ── Eindstatus ────────────────────────────────────────────────
    if n_fouten == 0:
        status = "gelukt"
    elif n_fouten < n_calls:
        status = "gedeeltelijk"
    else:
        status = "mislukt"

    db.update_artikel(artikel_id, status, n_calls, n_fouten)
    return n_fouten == 0, n_calls, n_fouten


def push_batch(
    alle_waarden: list[dict],
    cfg: KingConfig,
    db: PushDB,
    gebruiker: str,
    log_func=None,
    progress_func=None,
    dry_run: bool = False,
) -> dict:
    """
    Push een lijst artikelen naar KING en retourneert een samenvatting.
    log_func(msg, type) — zelfde interface als in app.py
    progress_func(huidig, totaal)
    """
    sessie_id = db.nieuwe_sessie(gebruiker, dry_run=dry_run)
    n_gelukt = n_mislukt = 0

    def _log(msg: str, typ: str = "info"):
        if log_func:
            log_func(msg, typ)

    if dry_run:
        _log("DRY-RUN — payloads worden opgebouwd en gelogd, niets wordt naar KING verstuurd", "warning")

    for i, waarden in enumerate(alle_waarden):
        if progress_func:
            progress_func(i + 1, len(alle_waarden))

        artnr = str(waarden.get("Artikelnummer", f"#{i+1}"))
        _log(f"[{i+1}/{len(alle_waarden)}] {artnr}", "info")

        try:
            gelukt, n_calls, n_fouten = push_artikel(
                waarden, cfg, db, sessie_id,
                log_func=lambda m: _log(m, "info"),
                dry_run=dry_run,
            )
        except Exception as e:
            _log(f"  {artnr}: onverwachte fout — {e}", "error")
            gelukt = False
            n_fouten = 1

        if gelukt:
            n_gelukt += 1
            _log(f"  {artnr}: gereed", "success")
        else:
            n_mislukt += 1
            _log(f"  {artnr}: mislukt of gedeeltelijk", "warning")

    db.update_sessie(sessie_id, len(alle_waarden), n_gelukt, n_mislukt)

    _log(
        f"Push klaar — {n_gelukt}/{len(alle_waarden)} artikelen volledig gelukt"
        + (f", {n_mislukt} met fouten" if n_mislukt else ""),
        "success" if n_mislukt == 0 else "warning",
    )
    return {"sessie_id": sessie_id, "n_gelukt": n_gelukt, "n_mislukt": n_mislukt}


# ── Util ──────────────────────────────────────────────────────────────────────

def _nu() -> str:
    return datetime.now().isoformat(timespec="seconds")


def huidig_gebruiker() -> str:
    try:
        return getpass.getuser()
    except Exception:
        return "onbekend"
