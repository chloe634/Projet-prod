# common/storage.py
from __future__ import annotations
import json, os, tempfile, shutil
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple, Any
import pandas as pd

STATE_DIR  = Path(".streamlit")
STATE_PATH = STATE_DIR / "saved_productions.json"
MAX_SLOTS  = 4

def _ensure_dir():
    STATE_DIR.mkdir(parents=True, exist_ok=True)
    if not STATE_PATH.exists():
        STATE_PATH.write_text("[]", encoding="utf-8")

def _read_all() -> List[Dict[str, Any]]:
    _ensure_dir()
    try:
        return json.loads(STATE_PATH.read_text(encoding="utf-8") or "[]")
    except Exception:
        return []

def _atomic_write(text: str):
    _ensure_dir()
    fd, tmp = tempfile.mkstemp(dir=str(STATE_DIR), prefix="sp_", suffix=".json")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(text)
    shutil.move(tmp, STATE_PATH)

def _encode_sp(sp: Dict[str, Any]) -> Dict[str, Any]:
    def _df(x):
        return x.to_json(orient="split") if isinstance(x, pd.DataFrame) else None
    return {
        "semaine_du": sp.get("semaine_du"),
        "ddm": sp.get("ddm"),
        "gouts": list(sp.get("gouts", [])),
        "df_min": _df(sp.get("df_min")),
        "df_calc": _df(sp.get("df_calc")),
    }

def _decode_sp(obj: Dict[str, Any]) -> Dict[str, Any]:
    def _df(s):
        return pd.read_json(s, orient="split") if isinstance(s, str) and s.strip() else None
    return {
        "semaine_du": obj.get("semaine_du"),
        "ddm": obj.get("ddm"),
        "gouts": obj.get("gouts") or [],
        "df_min": _df(obj.get("df_min")),
        "df_calc": _df(obj.get("df_calc")),
    }

def list_saved() -> List[Dict[str, Any]]:
    """Retourne [{name, ts, meta...}] triés du plus récent au plus ancien."""
    data = _read_all()
    data.sort(key=lambda x: x.get("ts",""), reverse=True)
    out = []
    for it in data:
        p = it.get("payload", {})
        out.append({
            "name": it.get("name"),
            "ts": it.get("ts"),
            "gouts": (p.get("gouts") or [])[:],
            "semaine_du": p.get("semaine_du"),
        })
    return out

def save_snapshot(name: str, sp: Dict[str, Any]) -> Tuple[bool, str]:
    """Crée ou remplace une proposition. Limite MAX_SLOTS si nouveau nom."""
    name = (name or "").strip()
    if not name:
        return False, "Nom vide."
    data = _read_all()
    # remplace si même nom
    idx = next((i for i, it in enumerate(data) if it.get("name")==name), None)
    entry = {
        "name": name,
        "ts": datetime.utcnow().isoformat(timespec="seconds") + "Z",
        "payload": _encode_sp(sp)
    }
    if idx is not None:
        data[idx] = entry
        _atomic_write(json.dumps(data, ensure_ascii=False, indent=2))
        return True, "Proposition mise à jour."
    # nouveau nom: respect limite
    if len(data) >= MAX_SLOTS:
        return False, f"Limite atteinte ({MAX_SLOTS}). Supprime ou renomme une entrée."
    data.append(entry)
    _atomic_write(json.dumps(data, ensure_ascii=False, indent=2))
    return True, "Proposition enregistrée."

def load_snapshot(name: str) -> Dict[str, Any] | None:
    data = _read_all()
    it = next((it for it in data if it.get("name")==name), None)
    return _decode_sp(it.get("payload", {})) if it else None

def delete_snapshot(name: str) -> bool:
    data = _read_all()
    new = [it for it in data if it.get("name") != name]
    if len(new) == len(data):
        return False
    _atomic_write(json.dumps(new, ensure_ascii=False, indent=2))
    return True

def rename_snapshot(old: str, new: str) -> Tuple[bool, str]:
    new = (new or "").strip()
    if not new:
        return False, "Nouveau nom vide."
    data = _read_all()
    if any(it.get("name")==new for it in data):
        return False, "Ce nom existe déjà."
    it = next((it for it in data if it.get("name")==old), None)
    if not it:
        return False, "Entrée introuvable."
    it["name"] = new
    _atomic_write(json.dumps(data, ensure_ascii=False, indent=2))
    return True, "Renommée."
