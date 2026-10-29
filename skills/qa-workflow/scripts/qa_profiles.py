#!/usr/bin/env python3
r"""Manage per-tester template profiles for the qa-workflow skill.

A profile remembers one tester's Excel template and how to read it (sheet,
header row, column mapping, required fields, metadata defaults), so many
testers can reuse the workflow with different templates and no re-detection.

Profiles live outside the skill folder, so reinstalling the skill keeps them:
  project scope: <project>\.kiro\qa-workflow\        (commit it to share with the team)
  user scope:    %USERPROFILE%\.kiro\qa-workflow\    (personal; or set QA_WORKFLOW_HOME)
Lookup order: --home, %QA_WORKFLOW_HOME%, <project>\.kiro\qa-workflow,
%USERPROFILE%\.kiro\qa-workflow.

Drop-in templates (no commands for testers): put .xlsx files in
  <project>\.kiro\qa-workflow\inbox\         (project; shared with the team)
  %USERPROFILE%\.kiro\qa-workflow\inbox\     (personal; all projects)
and run "auto": every new or changed template becomes a profile and the
newest one is reported as active.

Usage:
  python qa_profiles.py auto [--template FILE ...] [--json]
  python qa_profiles.py add NAME --template FILE [--owner TESTER] [--scope user|project]
                        [--default] [--sheet S] [--header-row N] [--data-start-row N]
                        [--col ROLE=COLUMN] [--meta "LABEL=VALUE"] [--required id,steps,expected]
  python qa_profiles.py list
  python qa_profiles.py show NAME
  python qa_profiles.py set-default NAME [--scope user|project]
  python qa_profiles.py remove NAME --yes

Exit codes: 0 = ok, 1 = refused, 2 = usage or file error.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sys
import unicodedata
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from validate_workbook import (  # noqa: E402
    Findings, UsageError, check_secrets, col_label, locate, open_workbook,
)
from openpyxl.utils import get_column_letter  # noqa: E402

# Kept under .kiro so templates show up in the Kiro explorer and stay with the
# tester's Kiro data; the power's own folder is replaced on every update.
HOME_DIR = Path(".kiro") / "qa-workflow"
INBOX = "inbox"
NAME_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$")
LAYOUT_KEYS = ("sheet", "header_row", "data_start_row", "col")
TEMPLATE_SUFFIXES = (".xlsx", ".xlsm")
BUNDLED_TEMPLATE = Path(__file__).resolve().parent.parent / "assets" / "template.xlsx"


class Refused(Exception):
    pass


def slugify(text) -> str:
    """'Template Ngân (API).xlsx' -> 'template-ngan-api'."""
    text = unicodedata.normalize("NFKD", str(text).replace("đ", "d").replace("Đ", "D"))
    text = "".join(ch for ch in text if not unicodedata.combining(ch)).lower()
    text = re.sub(r"[^a-z0-9]+", "-", text).strip("-")
    return (text or "template")[:64]


def sha256(path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def layout_summary(layout) -> str:
    cols = "; ".join(f"{r}={col_label(layout, r)}" for r in sorted(layout["cols"], key=layout["cols"].get))
    return (f"sheet {layout['sheet'].ws.title} | header row {layout['header_row']} | "
            f"data from row {layout['data_start']} | {cols}")


def save_profile(name, src, home, overrides=None, owner="", description="", metadata=None,
                 required=None, copy_template=True, make_default=False, allow_secrets=False):
    """Detect the layout of src and store it as profile `name` in `home`.

    `overrides` (sheet/header_row/data_start_row/col) are kept in the profile so
    re-detection after a template change re-applies them.
    """
    overrides = {k: v for k, v in (overrides or {}).items() if v not in (None, [])}
    src = Path(src)
    wb = open_workbook(src, data_only=False)
    layout = locate(wb, None, argparse.Namespace(**{k: overrides.get(k) for k in LAYOUT_KEYS}))
    if layout is None:
        raise Refused("could not find the Test Case header row; the tester needs to name the "
                      "sheet/header row/columns once (--sheet/--header-row/--col)")
    secrets = check_secrets(wb, Findings())
    if secrets and not allow_secrets:
        raise Refused("possible credentials in " + ", ".join(secrets[:10]) +
                      " (values not shown); replace them with placeholders such as <TOKEN>")
    if copy_template:
        dest = home / "templates" / f"{name}{src.suffix.lower()}"
        dest.parent.mkdir(parents=True, exist_ok=True)
        if dest.resolve() != src.resolve():
            shutil.copy2(src, dest)
        template_ref = f"templates/{dest.name}"
    else:
        template_ref = str(src.resolve())
    ws = layout["sheet"].ws
    profile = {
        "name": name,
        "owner": owner or "",
        "description": description or "",
        "template": template_ref,
        "source_template": str(src.resolve()),
        "source_sha256": sha256(src),
        "layout": {
            "sheet": ws.title,
            "header_row": layout["header_row"],
            "data_start_row": layout["data_start"],
            "col": [f"{role}={get_column_letter(c)}" for role, c in sorted(layout["cols"].items(), key=lambda x: x[1])],
        },
        "overrides": overrides,
        "required": required or "id,steps,expected",
        "metadata": metadata or {},
        "updated": date.today().isoformat(),
    }
    write_json(home / "profiles" / f"{name}.json", profile)
    if make_default:
        cfg_path = home / "config.json"
        cfg = read_json(cfg_path) if cfg_path.is_file() else {}
        cfg["default"] = name
        write_json(cfg_path, cfg)
    return profile, layout


def candidate_homes(home=None, project_dir=None):
    """Profile folders in lookup order, without duplicates."""
    paths = []
    if home:
        paths.append(Path(home).expanduser())
    if os.environ.get("QA_WORKFLOW_HOME"):
        paths.append(Path(os.environ["QA_WORKFLOW_HOME"]).expanduser())
    paths.append(Path(project_dir or Path.cwd()) / HOME_DIR)
    paths.append(Path.home() / HOME_DIR)
    seen, result = set(), []
    for p in paths:
        key = str(p.resolve())
        if key not in seen:
            seen.add(key)
            result.append(p)
    return result


def write_home(scope, home=None, project_dir=None) -> Path:
    if home:
        return Path(home).expanduser()
    if scope == "project":
        return Path(project_dir or Path.cwd()) / HOME_DIR
    if os.environ.get("QA_WORKFLOW_HOME"):
        return Path(os.environ["QA_WORKFLOW_HOME"]).expanduser()
    return Path.home() / HOME_DIR


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def write_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def load_profile(name=None, home=None, project_dir=None) -> dict:
    """Return a profile with an absolute "template_path". name None/'default' = default profile."""
    homes = candidate_homes(home, project_dir)
    if not name or name == "default":
        for h in homes:
            cfg = h / "config.json"
            if cfg.is_file() and read_json(cfg).get("default"):
                name = read_json(cfg)["default"]
                break
        else:
            raise UsageError("No default profile set. Run qa_profiles.py list / set-default.")
    for h in homes:
        path = h / "profiles" / f"{name}.json"
        if path.is_file():
            profile = read_json(path)
            template = Path(profile["template"])
            profile["template_path"] = str(template if template.is_absolute() else h / template)
            profile["home"] = str(h)
            if not Path(profile["template_path"]).is_file():
                raise UsageError(f"Profile {name!r}: template not found at {profile['template_path']}")
            return profile
    raise UsageError(f"Profile {name!r} not found in: " + ", ".join(str(h) for h in homes))


def apply_profile(args, profile):
    """Fill unset CLI options from a profile; explicit options always win."""
    layout = profile.get("layout", {})
    for key in LAYOUT_KEYS:
        if getattr(args, key, None) in (None, []):
            setattr(args, key, layout.get(key))
    if hasattr(args, "required") and not args.required:
        args.required = profile.get("required")
    if hasattr(args, "meta"):
        merged = dict(profile.get("metadata") or {})
        merged.update(dict(item.split("=", 1) for item in (args.meta or [])))
        args.meta = [f"{k}={v}" for k, v in merged.items()]


def iter_profiles(home=None, project_dir=None):
    seen = set()
    for h in candidate_homes(home, project_dir):
        default = None
        if (h / "config.json").is_file():
            default = read_json(h / "config.json").get("default")
        for path in sorted((h / "profiles").glob("*.json")) if (h / "profiles").is_dir() else []:
            name = path.stem
            data = read_json(path)
            yield h, name, data, name == default, name in seen
            seen.add(name)


def cmd_add(args) -> int:
    if not NAME_RE.match(args.name):
        raise UsageError("Profile name: letters, digits, '.', '_' or '-' (max 64), e.g. ngan-api")
    for item in args.meta or []:
        if "=" not in item:
            raise UsageError(f"--meta expects LABEL=VALUE, got {item!r}")
    home = write_home(args.scope, args.home, args.project_dir)
    profile_path = home / "profiles" / f"{args.name}.json"
    if profile_path.exists() and not args.force:
        print(f"Refused: profile {args.name!r} already exists at {profile_path}; pass --force to replace it.")
        return 1
    overrides = {"sheet": args.sheet, "header_row": args.header_row,
                 "data_start_row": args.data_start_row, "col": args.col}
    try:
        profile, layout = save_profile(
            args.name, args.template, home, overrides=overrides, owner=args.owner,
            description=args.description, metadata=dict(item.split("=", 1) for item in (args.meta or [])),
            required=args.required, copy_template=not args.no_copy, make_default=args.default,
            allow_secrets=args.allow_secrets)
    except Refused as exc:
        print(f"Refused: {exc}.")
        return 1
    print(f"Saved profile {args.name!r} -> {profile_path}")
    print(f"Template: {profile['template']}" + (" (linked, not copied)" if args.no_copy else " (copy)"))
    print(layout_summary(layout))
    if layout["tc_rows"]:
        print(f"Note: the template already has {len(layout['tc_rows'])} test case row(s); "
              "exports will need --existing clear or append.")
    return 0


def user_home_dir(home=None) -> Path:
    if home:
        return Path(home).expanduser()
    if os.environ.get("QA_WORKFLOW_HOME"):
        return Path(os.environ["QA_WORKFLOW_HOME"]).expanduser()
    return Path.home() / HOME_DIR


def template_files(folder):
    if not folder.is_dir():
        return []
    return sorted(p for p in folder.iterdir()
                  if p.is_file() and p.suffix.lower() in TEMPLATE_SUFFIXES
                  and not p.name.startswith(("~$", ".")))  # skip Excel lock files


def cmd_auto(args) -> int:
    """Turn every dropped-in template into a profile and report which one to use."""
    project = Path(args.project_dir or Path.cwd())
    user_home = user_home_dir(args.home)
    inboxes = [("project", project / HOME_DIR / INBOX, project / HOME_DIR),
               ("user", user_home / INBOX, user_home)]
    try:
        (user_home / INBOX).mkdir(parents=True, exist_ok=True)  # our own folder; shows testers where to drop files
    except OSError:
        pass

    items = [(scope, f, home, False) for scope, inbox, home in inboxes for f in template_files(inbox)]
    explicit_home = project / HOME_DIR if args.scope == "project" else user_home
    for t in args.template or []:
        if not Path(t).is_file():
            raise UsageError(f"Template not found: {t}")
        items.append((args.scope, Path(t), explicit_home, True))

    results = []
    for scope, src, home, explicit in items:
        name = slugify(src.stem)
        path = home / "profiles" / f"{name}.json"
        existing = read_json(path) if path.is_file() else None
        st = src.stat()
        # Copying keeps the old modified time on Windows; creation time marks when it was dropped in.
        dropped = max(st.st_mtime, getattr(st, "st_birthtime", 0) or 0)
        entry = {"file": str(src), "profile": name, "scope": scope, "explicit": explicit, "mtime": dropped}
        if (existing and existing.get("source_sha256") == sha256(src)
                and (home / existing["template"]).is_file()):
            lay = existing["layout"]
            entry.update(status="unchanged", layout=f"sheet {lay['sheet']} | header row {lay['header_row']} | "
                                                    f"data from row {lay['data_start_row']} | {'; '.join(lay['col'])}")
        else:
            old = existing or {}
            try:
                _, layout = save_profile(name, src, home, overrides=old.get("overrides"), owner=old.get("owner"),
                                         description=old.get("description"), metadata=old.get("metadata"),
                                         required=old.get("required"))
                entry.update(status="updated" if existing else "new", layout=layout_summary(layout))
                if layout["tc_rows"]:
                    entry["sample_rows"] = len(layout["tc_rows"])
            except Refused as exc:
                entry.update(status="refused", reason=str(exc))
            except UsageError as exc:
                entry.update(status="unreadable", reason=str(exc))
        results.append(entry)

    usable = [e for e in results if e["status"] in ("new", "updated", "unchanged")]
    explicit_ok = [e for e in usable if e["explicit"]]
    chosen = explicit_ok[-1] if explicit_ok else (max(usable, key=lambda e: e["mtime"]) if usable else None)
    if chosen:
        active = {"profile": chosen["profile"], "template": chosen["file"], "source": chosen["scope"] + " template"}
        if chosen["scope"] == "user" and chosen["status"] != "unchanged":
            cfg_path = user_home / "config.json"
            cfg = read_json(cfg_path) if cfg_path.is_file() else {}
            cfg["default"] = chosen["profile"]
            write_json(cfg_path, cfg)
    else:
        try:
            p = load_profile(None, args.home, args.project_dir)
            active = {"profile": p["name"], "template": p["template_path"], "source": "default profile"}
        except UsageError:
            active = {"profile": None, "template": str(BUNDLED_TEMPLATE), "source": "bundled default template"}

    report = {
        "inboxes": [{"scope": s, "path": str(i), "exists": i.is_dir()} for s, i, _ in inboxes],
        "templates": [{k: v for k, v in e.items() if k != "mtime"} for e in results],
        "active": active,
        "choose_needed": len(usable) > 1 and not explicit_ok,
    }
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        for box in report["inboxes"]:
            print(f"Inbox ({box['scope']}): {box['path']}" + ("" if box["exists"] else " (not created)"))
        for e in report["templates"]:
            detail = e.get("layout") or e.get("reason", "")
            extra = f" | {e['sample_rows']} sample row(s)" if e.get("sample_rows") else ""
            print(f"- {Path(e['file']).name} -> profile {e['profile']} [{e['status']}] {detail}{extra}")
        print(f"Active: {active['profile'] or '(none)'} | {active['template']} ({active['source']})")
        if report["choose_needed"]:
            print("Several templates are available; ask the tester which one to use (newest is shown as active).")
    return 0


def cmd_list(args) -> int:
    rows = list(iter_profiles(args.home, args.project_dir))
    if not rows:
        print("No profiles yet. Create one with: qa_profiles.py add NAME --template FILE")
        return 0
    for home, name, data, is_default, shadowed in rows:
        flags = (" [default]" if is_default else "") + (" [shadowed]" if shadowed else "")
        owner = f" owner={data.get('owner')}" if data.get("owner") else ""
        print(f"{name}{flags}{owner} | {data.get('template')} | {home}"
              + (f" | {data['description']}" if data.get("description") else ""))
    return 0


def cmd_show(args) -> int:
    profile = load_profile(args.name, args.home, args.project_dir)
    print(json.dumps(profile, ensure_ascii=False, indent=2))
    return 0


def cmd_set_default(args) -> int:
    load_profile(args.name, args.home, args.project_dir)
    home = write_home(args.scope, args.home, args.project_dir)
    cfg_path = home / "config.json"
    cfg = read_json(cfg_path) if cfg_path.is_file() else {}
    cfg["default"] = args.name
    write_json(cfg_path, cfg)
    print(f"Default profile in {home}: {args.name}")
    return 0


def cmd_remove(args) -> int:
    profile = load_profile(args.name, args.home, args.project_dir)
    if not args.yes:
        print(f"Refused: pass --yes to delete profile {args.name!r} in {profile['home']}.")
        return 1
    home = Path(profile["home"])
    (home / "profiles" / f"{args.name}.json").unlink()
    template = Path(profile["template_path"])
    if not Path(profile["template"]).is_absolute() and template.is_file():
        template.unlink()  # only the copy we own; linked templates are left alone
    cfg_path = home / "config.json"
    if cfg_path.is_file() and read_json(cfg_path).get("default") == args.name:
        cfg = read_json(cfg_path)
        cfg.pop("default")
        write_json(cfg_path, cfg)
    print(f"Removed profile {args.name!r} from {home}")
    return 0


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
    p = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    p.add_argument("--home", help="profile folder to use instead of the lookup order")
    p.add_argument("--project-dir", help="project root for project-scope profiles (default: current dir)")
    sub = p.add_subparsers(dest="command", required=True)

    a = sub.add_parser("add", help="create or replace a profile from a template")
    a.add_argument("name")
    a.add_argument("--template", required=True)
    a.add_argument("--owner", help="tester name or account")
    a.add_argument("--description")
    a.add_argument("--scope", choices=("user", "project"), default="user")
    a.add_argument("--default", action="store_true", help="make it the default profile in that scope")
    a.add_argument("--sheet")
    a.add_argument("--header-row", type=int)
    a.add_argument("--data-start-row", type=int)
    a.add_argument("--col", action="append", metavar="ROLE=COLUMN")
    a.add_argument("--required", help="roles that must be filled (default id,steps,expected)")
    a.add_argument("--meta", action="append", metavar="LABEL=VALUE",
                   help='workbook field to fill on export, by label text ("Originator:=ngannt23") '
                        'or cell ("Cover!E13=ngannt23")')
    a.add_argument("--no-copy", action="store_true", help="link to the template instead of copying it")
    a.add_argument("--allow-secrets", action="store_true", help="store even if credentials are detected")
    a.add_argument("--force", action="store_true", help="replace an existing profile")

    au = sub.add_parser("auto", help="import dropped-in templates and report the active one")
    au.add_argument("--template", action="append", metavar="FILE",
                    help="template the tester attached or named in chat (wins over inbox files)")
    au.add_argument("--scope", choices=("user", "project"), default="user",
                    help="where to store profiles for --template files (default: user)")
    au.add_argument("--json", action="store_true")

    sub.add_parser("list", help="list profiles from all lookup locations")
    s = sub.add_parser("show", help="print one profile")
    s.add_argument("name")
    d = sub.add_parser("set-default", help="set the default profile")
    d.add_argument("name")
    d.add_argument("--scope", choices=("user", "project"), default="user")
    r = sub.add_parser("remove", help="delete a profile (and its copied template)")
    r.add_argument("name")
    r.add_argument("--yes", action="store_true")

    args = p.parse_args(argv)
    handlers = {"auto": cmd_auto, "add": cmd_add, "list": cmd_list, "show": cmd_show,
                "set-default": cmd_set_default, "remove": cmd_remove}
    try:
        return handlers[args.command](args)
    except UsageError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
