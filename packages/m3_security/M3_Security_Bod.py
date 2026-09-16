"""
M3_Security_Bod - the SyncSecurityRoleMaster event document.

Deleting a security role in M3 is not something MNS405MI will do, so the
removal is published as an event document instead: the same EventData that
M3 itself sends for CMNROL, addressed to the SecurityRoleMaster smart rule.

Per-tenant settings (owner/CHID and the running sequence number) live in
M3_Security_Bod.json beside this file. Nothing here is mandatory - a document
can be built from the tenant and the role name alone.

Reconstructed from M3_Security_Bod.cpython-312.pyc after the source was lost.
Logic matches the bytecode; comments and formatting are not original.
"""
from __future__ import annotations

import argparse
import json
import logging
import re
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape

log = logging.getLogger("M3_Security_Bod")

BASE_DIR = Path(__file__).parent.resolve()
CONFIG_PATH = BASE_DIR / "M3_Security_Bod.json"
DEFAULT_BOD_DIR = BASE_DIR / "output" / "m3_security" / "bods"

# Fixed parts of the envelope, exactly as M3 sends them.
PUBLISHER = "M3BODProcessor"
DOCUMENT_NAME = "SyncSecurityRoleMaster"
BOD_NOUN = "SecurityRoleMaster"
BOD_VERB = "Sync"
ORIGIN_PUBLISHER = "M3"
ORIGIN_DOCUMENT = "CMNROL"
FIND_DIVI = "NOLOOKUP"
CORRELATION_TYPE = "NONE"

# Which smart rule the event is addressed to. {op} is the rule operation
# (DELETE / UPDATE / ...), which is separate from the <Operation> element.
PROCESSOR_TEMPLATE = (
    "com.infor.event.analytics.smartrules.M3BODs_SecurityRoleMaster."
    "CMNROL_{op}_SecurityRoleMaster"
)

# <Operation> stays UPDATE - that is what M3 puts on the wire even when the
# rule being triggered is the delete one.
DEFAULT_OPERATION = "UPDATE"
DEFAULT_CURRENT_PROGRAM = "MNS405"
DEFAULT_START_PROGRAM = "MNS405"
API_CURRENT_PROGRAM = "MNS405Fnc"
API_START_PROGRAM = "MNS405MI"
DEFAULT_VERSION = "16"
DEFAULT_APPLICATION = "foundation"
DEFAULT_ENVIRONMENT = "foundation"
DEFAULT_STATUS_CODE = "Deleted"

# CMNROL.ROLL is a 10-character field; the key value is padded to it.
ROLL_WIDTH = 10


class BodConfigError(RuntimeError):
    """Something needed for the event document is missing."""


def load_bod_config(path: str | Path = CONFIG_PATH) -> dict:
    p = Path(path)
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text())
    except Exception as exc:
        raise BodConfigError(f"{p.name} is not valid JSON: {exc}") from exc


def save_bod_config(cfg: dict, path: str | Path = CONFIG_PATH) -> Path:
    p = Path(path)
    p.write_text(json.dumps(cfg, indent=2, sort_keys=True) + "\n")
    return p


def tenant_bod_config(tenant: str, path: str | Path = CONFIG_PATH) -> dict:
    """
    Settings for one tenant, all with defaults.

    Nothing here is mandatory - the event document can be built from the tenant
    and the role alone - so this never raises for a missing value.
    """
    cfg = load_bod_config(path).get(tenant) or {}
    return {
        "owner": (cfg.get("owner") or "").strip(),
        "sequence": int(cfg.get("sequence") or 0),
        "version": (cfg.get("version") or DEFAULT_VERSION).strip(),
        "application": (cfg.get("application") or DEFAULT_APPLICATION).strip(),
        "environment": (cfg.get("environment") or DEFAULT_ENVIRONMENT).strip(),
        "operation": (cfg.get("operation") or DEFAULT_OPERATION).strip(),
        "current_program": (cfg.get("current_program")
                            or DEFAULT_CURRENT_PROGRAM).strip(),
        "start_program": (cfg.get("start_program")
                          or DEFAULT_START_PROGRAM).strip(),
        "old_values": bool(cfg.get("old_values", False)),
    }


def next_sequence(tenant: str, path: str | Path = CONFIG_PATH) -> int:
    """Take the next sequence number for a tenant and store it."""
    cfg = load_bod_config(path)
    entry = cfg.setdefault(tenant, {})
    seq = int(entry.get("sequence") or 0) + 1
    entry["sequence"] = seq
    save_bod_config(cfg, path)
    return seq


def sent_timestamp(when: datetime | None = None) -> str:
    """2026-08-18T14:49:50.668Z"""
    when = (when or datetime.now(timezone.utc)).astimezone(timezone.utc)
    return when.strftime("%Y-%m-%dT%H:%M:%S.") + f"{when.microsecond // 1000:03d}Z"


def key_value(role: str) -> str:
    """
    The CMNROL key as the event carries it.

    "AP ADMIN" is 8 characters, padded to the 10-character ROLL field and
    written with spaces as "+", giving "KRROLL,AP+ADMIN++".
    """
    padded = (role or "")[:ROLL_WIDTH].ljust(ROLL_WIDTH)
    return "KRROLL," + padded.replace(" ", "+")


def build_event(
    role: str,
    tenant_id: str,
    tx15: str = "",
    tx40: str = "",
    owner: str = "",
    sequence: int | None = None,
    operation: str = DEFAULT_OPERATION,
    status_code: str = DEFAULT_STATUS_CODE,
    version: str = DEFAULT_VERSION,
    application: str = DEFAULT_APPLICATION,
    environment: str = DEFAULT_ENVIRONMENT,
    rgdt: str = "",
    rgtm: str = "",
    current_program: str = DEFAULT_CURRENT_PROGRAM,
    start_program: str = DEFAULT_START_PROGRAM,
    old_values: bool = False,
    rule_operation: str = "DELETE",
    when: datetime | None = None,
) -> dict:
    """Build one EventData document. Returns the xml plus what went into it."""
    role = (role or "").strip()
    if not role:
        raise BodConfigError("A role name is required.")
    if not (tenant_id or "").strip():
        raise BodConfigError("A tenant is required.")

    when = (when or datetime.now(timezone.utc)).astimezone(timezone.utc)
    tracking_id = str(uuid.uuid4())
    event_id = str(uuid.uuid4())
    lmts = str(int(when.timestamp() * 1000))
    rgdt = (rgdt or "").strip() or when.strftime("%Y%m%d")
    rgtm = (rgtm or "").strip() or when.strftime("%H%M%S")
    kv = key_value(role)

    def old(v: str) -> str | None:
        # The smart rule only wants OldValue when the config asks for it.
        return v if old_values else None

    elements = [
        ("ROLL", role, old(role)),
        ("CHID", owner, old(owner)),
        ("LMTS", lmts, old(lmts)),
        ("RGDT", rgdt, old(rgdt)),
        ("RGTM", rgtm, old(rgtm)),
        ("TX15", tx15, old(tx15)),
        ("TX40", tx40, old(tx40)),
        ("application", application, None),
        ("currentProgram", current_program, None),
        ("environment", environment, None),
        ("owner", owner, None),
        ("startProgram", start_program, None),
        ("version", version, None),
        ("BODNoun", BOD_NOUN, None),
        ("BODVerb", BOD_VERB, None),
        ("findDIVI", FIND_DIVI, None),
        ("SecurityRoleMasterStatusCode", status_code, None),
        ("keyValue", kv, None),
        ("originKeyValue", kv, None),
        ("originPublisher", ORIGIN_PUBLISHER, None),
        ("originDocument", ORIGIN_DOCUMENT, None),
        ("*processorName", PROCESSOR_TEMPLATE.format(op=rule_operation), None),
        ("*correlationType", CORRELATION_TYPE, None),
    ]

    parts = []
    for name, value, old in elements:
        el = f"<ElementData><Name>{escape(name)}</Name><Value>{escape(value)}</Value>"
        if old is not None:
            el += f"<OldValue>{escape(old)}</OldValue>"
        parts.append(el + "</ElementData>")

    xml = (
        f"<?xml version='1.0' encoding='UTF-8'?><EventData>"
        f"<TenantId>{escape(tenant_id)}</TenantId>"
        f"<Publisher>{PUBLISHER}</Publisher>"
        f"<DocumentName>{DOCUMENT_NAME}</DocumentName>"
        f"<Operation>{escape(operation)}</Operation>"
        f"<TrackingId>{tracking_id}</TrackingId>"
        f"<EventId>{event_id}</EventId>"
        f"<SentTimestamp>{sent_timestamp(when)}</SentTimestamp>"
        + (f"<Sequence>{sequence}</Sequence>" if sequence is not None else "")
        + "<Document>" + "".join(parts) + "</Document></EventData>"
    )

    return {
        "role": role,
        "xml": xml,
        "operation": operation,
        "rule_operation": rule_operation,
        "tenant_id": tenant_id,
        "tracking_id": tracking_id,
        "event_id": event_id,
        "sequence": sequence,
        "sent_timestamp": sent_timestamp(when),
        "key_value": kv,
        "status_code": status_code,
        "owner": owner,
        "tx15": tx15,
        "tx40": tx40,
    }


_SAFE = re.compile("[^A-Za-z0-9._-]+")


def bod_file_name(role: str, operation: str = "DELETE", when: datetime | None = None) -> str:
    when = (when or datetime.now(timezone.utc)).astimezone(timezone.utc)
    safe = _SAFE.sub("_", (role or "").strip()) or "role"
    return (f"SyncSecurityRoleMaster_{operation}_{safe}_"
            f"{when.strftime('%Y%m%d_%H%M%S_%f')}.xml")


def write_event(event: dict, out_dir: str | Path = DEFAULT_BOD_DIR,
                when: datetime | None = None) -> Path:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / bod_file_name(event["role"],
                                   event.get("rule_operation")
                                   or event["operation"],
                                   when)
    path.write_text(event["xml"], encoding="utf-8")
    return path


def build_and_write(
    roles: list[dict] | list[str],
    tenant: str,
    out_dir: str | Path = DEFAULT_BOD_DIR,
    config_path: str | Path = CONFIG_PATH,
    operation: str = "DELETE",
    status_code: str = DEFAULT_STATUS_CODE,
) -> list[dict]:
    """
    One event document per role, written to out_dir.

    roles may be plain names or dicts carrying roll / tx15 / tx40 / rgdt / rgtm,
    so a caller that already has the role record does not make us re-read it.
    Each entry gets 'file' and 'error'; a role that could not be written is
    reported rather than dropped.
    """
    cfg = tenant_bod_config(tenant, config_path)
    out = []
    for item in roles:
        rec = {"roll": item} if isinstance(item, str) else dict(item)
        role = (rec.get("roll") or "").strip()
        try:
            event = build_event(
                role, tenant,
                tx15=(rec.get("tx15") or "").strip(),
                tx40=(rec.get("tx40") or "").strip(),
                owner=cfg["owner"],
                sequence=next_sequence(tenant, config_path),
                operation=cfg["operation"],
                status_code=status_code,
                version=cfg["version"],
                application=cfg["application"],
                environment=cfg["environment"],
                rgdt=rec.get("rgdt") or "",
                rgtm=rec.get("rgtm") or "",
                current_program=cfg["current_program"],
                start_program=cfg["start_program"],
                old_values=cfg["old_values"],
                rule_operation=operation,
            )
            path = write_event(event, out_dir)
            out.append({**{k: v for k, v in event.items() if k != "xml"},
                        "file": path.name,
                        "path": str(path),
                        "error": ""})
        except Exception as exc:
            log.error("event for %s: %s", role, exc)
            out.append({"role": role, "file": "", "path": "", "error": str(exc)})
    return out


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="Build the SyncSecurityRoleMaster event document for a role.")
    ap.add_argument("--tenant")
    ap.add_argument("--role", nargs="*", default=[])
    ap.add_argument("--operation", default="DELETE",
                    help="Which smart rule the event names: CMNROL_<OP>_... "
                         "The <Operation> element follows the config and stays "
                         "UPDATE, as M3 sends it")
    ap.add_argument("--status", default=DEFAULT_STATUS_CODE)
    ap.add_argument("--tx15", default="")
    ap.add_argument("--tx40", default="")
    ap.add_argument("--out-dir", default=str(DEFAULT_BOD_DIR))
    ap.add_argument("--config", default=str(CONFIG_PATH))
    ap.add_argument("--show-config", action="store_true")
    ap.add_argument("--set-owner", help="CHID / owner to stamp on the events")
    ap.add_argument("--set-sequence", type=int,
                    help="Last used sequence; the next document takes +1")
    ap.add_argument("--print", action="store_true",
                    help="Print the XML instead of writing files")
    args = ap.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(message)s",
                        handlers=[logging.StreamHandler(sys.stdout)])

    if args.show_config:
        cfg = load_bod_config(args.config)
        print(json.dumps(cfg, indent=2, sort_keys=True) if cfg
              else f"{Path(args.config).name} is empty or missing.")
        return 0

    if args.set_owner is not None or args.set_sequence is not None:
        if not args.tenant:
            ap.error("--set-owner / --set-sequence need --tenant")
        cfg = load_bod_config(args.config)
        entry = cfg.setdefault(args.tenant, {})
        if args.set_owner is not None:
            entry["owner"] = args.set_owner
        if args.set_sequence is not None:
            entry["sequence"] = args.set_sequence
        print(f"Saved to {save_bod_config(cfg, args.config)}")
        return 0

    if not args.tenant or not args.role:
        ap.error("--tenant and at least one --role are required")

    if args.print:
        c = tenant_bod_config(args.tenant, args.config)
        for role in args.role:
            print(build_event(role, args.tenant, args.tx15, args.tx40,
                              owner=c["owner"],
                              sequence=c["sequence"] + 1,
                              operation=c["operation"],
                              status_code=args.status,
                              version=c["version"],
                              application=c["application"],
                              environment=c["environment"],
                              current_program=c["current_program"],
                              start_program=c["start_program"],
                              old_values=c["old_values"],
                              rule_operation=args.operation)["xml"])
        return 0

    rc = 0
    for r in build_and_write(
        [{"roll": x, "tx15": args.tx15, "tx40": args.tx40} for x in args.role],
        args.tenant, args.out_dir, args.config,
        args.operation, args.status,
    ):
        if r["error"]:
            log.error("  %-12s %s", r["role"], r["error"])
            rc = 1
            continue
        log.info("  %-12s seq %-8s -> %s", r["role"], r["sequence"], r["path"])
    return rc


if __name__ == "__main__":
    raise SystemExit(main())
