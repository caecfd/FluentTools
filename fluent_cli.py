# Apache-2.0 (see LICENSE in the repository root).
"""Standalone CLI for the PyFluent-MCP tool surface — no MCP server required.

Each MCP tool of ``ansys-fluent-mcp`` becomes a CLI subcommand that emits a
JSON document — to stdout by default, or to a file with ``--json-out``. The
heavy lifting (launch validation, sandboxed
``run_code``, introspection, mesh/quality parsing, file comparison, offline
API search) is reused verbatim from the package's backend layer
(``ansys.fluent.mcp.solve.backends.pyfluent`` and ``solve.lib``), so CLI
results are byte-for-byte the same envelopes the MCP tools return.

Session model
-------------
MCP holds the PyFluent session in one long-lived process; a CLI script is a
fresh process per call. The bridge is a small session file:

* ``connect`` launches (or attaches to) Fluent, then persists the live
  session's ``connection_properties`` (ip/port/password) to the session file.
  Launch mode uses ``cleanup_on_exit=False`` so Fluent outlives the script.
* Every other command reads the session file and attaches to the running
  Fluent over gRPC within seconds — the Fluent process itself is untouched.
* ``disconnect`` attaches and shuts Fluent down, then deletes the file.

Default session file: ``~/.fluent_tools/session.json``. Override with
``--session-file`` or the ``FLUENT_TOOLS_SESSION`` environment variable.

Exit codes: 0 success, 1 runtime/tool error, 2 usage error,
3 no live session (stale session file or Fluent unreachable).
"""

from __future__ import annotations

import argparse
import asyncio
import base64
import dataclasses
import json
import os
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional

# ---------------------------------------------------------------------------
# Path bootstrap: prefer this checkout's sources, fall back to an installed
# ``ansys-fluent-mcp`` package. Keeps the CLI usable both from a repo clone
# and after ``pip install ansys-fluent-mcp``.
# ---------------------------------------------------------------------------
_TOOLS_DIR = Path(__file__).resolve().parent
_REPO_SRC = _TOOLS_DIR.parent / "src"
if (_REPO_SRC / "ansys" / "fluent" / "mcp").is_dir():
    sys.path.insert(0, str(_REPO_SRC))

DEFAULT_SESSION_FILE = Path.home() / ".fluent_tools" / "session.json"

# Set in main() from ``--json-out``; when non-None, ``_emit`` writes the JSON
# envelope to this file instead of stdout. Module-level so ``_fail`` (which has
# no access to ``args``) honors it too.
_JSON_OUT: Optional[Path] = None


def _fail(message: str, *, exit_code: int = 1, **extra: Any) -> None:
    """Print a JSON error envelope and exit."""
    _emit({"error": message, **extra}, pretty=False)
    raise SystemExit(exit_code)


def _emit(payload: Any, *, pretty: bool) -> None:
    """Serialize ``payload`` as JSON to stdout, or to ``--json-out`` when set."""
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    text = json.dumps(
        _to_jsonable(payload),
        ensure_ascii=False,
        indent=2 if pretty else None,
        separators=None if pretty else (",", ":"),
        default=str,
    )
    if _JSON_OUT is not None:
        try:
            _JSON_OUT.parent.mkdir(parents=True, exist_ok=True)
            _JSON_OUT.write_text(text + "\n", encoding="utf-8")
        except OSError as exc:
            print(f"could not write result to {_JSON_OUT}: {exc}", file=sys.stderr)
            raise SystemExit(1)
        print(f"result written to {_JSON_OUT}", file=sys.stderr)
        return
    print(text)


def _to_jsonable(obj: Any) -> Any:
    """Convert pydantic models / dataclasses / paths into JSON-safe values."""
    if obj is None or isinstance(obj, (bool, int, float, str)):
        return obj
    if isinstance(obj, Path):
        return str(obj)
    model_dump = getattr(obj, "model_dump", None)
    if callable(model_dump):
        return model_dump(mode="json")
    if dataclasses.is_dataclass(obj) and not isinstance(obj, type):
        return dataclasses.asdict(obj)
    if isinstance(obj, dict):
        return {str(k): _to_jsonable(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple, set)):
        return [_to_jsonable(v) for v in obj]
    return obj


def _session_file(args: argparse.Namespace) -> Path:
    """Resolve the session file from CLI/env/default precedence."""
    raw = getattr(args, "session_file", None) or os.environ.get("FLUENT_TOOLS_SESSION")
    return Path(raw) if raw else DEFAULT_SESSION_FILE


def _load_session(args: argparse.Namespace) -> Optional[dict[str, Any]]:
    """Return the persisted session mapping, or ``None`` when absent."""
    path = _session_file(args)
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        _fail(
            f"session file {path} is unreadable: {exc}; delete it or run `connect` again",
            exit_code=3,
        )
    if not isinstance(data, dict) or "port" not in data:
        return None
    return data


def _save_session(args: argparse.Namespace, session: dict[str, Any]) -> Path:
    """Persist the session mapping and return the file path."""
    path = _session_file(args)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(session, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return path


def _delete_session(args: argparse.Namespace) -> None:
    """Remove the session file if present."""
    path = _session_file(args)
    try:
        path.unlink(missing_ok=True)
    except OSError:
        pass


def _new_backend() -> Any:
    """Build a fresh ``PyFluentBackend`` (no session attached yet)."""
    from ansys.fluent.mcp.solve.backends.pyfluent import PyFluentBackend

    return PyFluentBackend()


async def _attach_from_session(args: argparse.Namespace) -> Any:
    """Attach a backend to the Fluent session recorded in the session file.

    Exits with code 3 when no session file exists or the recorded Fluent is
    unreachable (stale session). The stale file is kept so ``--force``
    diagnostics and the user can still inspect it; `connect` overwrites it.
    """
    session = _load_session(args)
    if session is None:
        _fail(
            "no active session: run `connect` first (or pass --session-file / "
            "set FLUENT_TOOLS_SESSION)",
            exit_code=3,
        )
    backend = _new_backend()
    result = await backend.connect(
        ip=session.get("ip") or "127.0.0.1",
        port=session.get("port"),
        password=session.get("password"),
    )
    if getattr(result, "status", "") != "ok":
        _fail(
            "could not attach to the recorded Fluent session "
            f"({session.get('ip')}:{session.get('port')}): "
            f"{getattr(result, 'message', '') or 'connection failed'}. "
            "It may have exited; run `connect` again.",
            exit_code=3,
            error_code="stale_session",
        )
    return backend


def _require_backend(coro_factory: Any, args: argparse.Namespace) -> Any:
    """Run an async backend operation after attaching from the session file."""
    async def _runner() -> Any:
        backend = await _attach_from_session(args)
        return await coro_factory(backend)

    return asyncio.run(_runner())


# ---------------------------------------------------------------------------
# connect / disconnect / session_status
# ---------------------------------------------------------------------------

def _gpu_value(raw: Optional[str]) -> Optional[Any]:
    """Normalize ``--gpu`` CLI text into a bool or list of device ids."""
    if raw is None:
        return None
    text = raw.strip().lower()
    if text in {"true", "yes", "on"}:
        return True
    if text in {"false", "no", "off"}:
        return False
    try:
        return [int(part) for part in text.split(",") if part.strip()]
    except ValueError as exc:
        raise SystemExit(f"--gpu must be true/false or a comma-separated id list: {exc}")


def cmd_connect(args: argparse.Namespace) -> Any:
    """Launch or attach Fluent and persist the session for later commands."""
    attach_kwargs = {
        "ip": args.ip,
        "port": args.port,
        "password": args.password,
        "server_info_file": args.server_info_file,
    }
    explicit_attach = any(v is not None for v in attach_kwargs.values())

    # Guard against double launches: if a recorded session is still alive,
    # report it instead of starting a second Fluent.
    existing = _load_session(args)
    if existing is not None and not args.force and not explicit_attach:
        try:
            asyncio.run(_attach_from_session(args))
        except SystemExit:
            pass  # stale — fall through to a fresh launch
        else:
            return {
                "status": "ok",
                "already_connected": True,
                "endpoint": f"{existing.get('ip')}:{existing.get('port')}",
                "mode": existing.get("mode"),
                "message": (
                    "Session file records a live Fluent session; no new launch. "
                    "Use --force to relaunch (after `disconnect`)."
                ),
            }

    async def _run() -> Any:
        backend = _new_backend()
        if explicit_attach:
            result = await backend.connect(**{k: v for k, v in attach_kwargs.items() if v is not None})
            if getattr(result, "status", "") != "ok":
                return {"error": getattr(result, "message", "attach failed"),
                        "error_code": getattr(result, "error_code", None)}
            props = backend._solver.connection_properties
            session = {
                "mode": "attach",
                "ip": getattr(props, "ip", None) or args.ip or "127.0.0.1",
                "port": getattr(props, "port", None) or args.port,
                "password": getattr(props, "password", None) or args.password,
                "attached_at": datetime.now(timezone.utc).isoformat(),
            }
            path = _save_session(args, session)
            return {
                "status": "ok",
                "already_connected": False,
                "mode": "attach",
                "endpoint": f"{session['ip']}:{session['port']}",
                "session_file": str(path),
                "message": getattr(result, "message", None),
            }

        # Route the Fluent child process console output to a log file next
        # to the session file; otherwise Fluent writes errors straight to
        # the terminal and pollutes the JSON stdout contract of later calls.
        import ansys.fluent.core as _pfl

        log_path = _session_file(args).parent / "fluent_launch.log"
        log_handle = open(log_path, "a", encoding="utf-8", buffering=1)
        prev_stdout = _pfl.config.launch_fluent_stdout
        prev_stderr = _pfl.config.launch_fluent_stderr
        _pfl.config.launch_fluent_stdout = log_handle
        _pfl.config.launch_fluent_stderr = log_handle
        try:
            result = await backend.connect(
                precision=args.precision,
                processor_count=args.processor_count,
                ui_mode=args.ui_mode,
                product_version=args.product_version,
                dimension=args.dimension,
                mode=args.mode,
                gpu=_gpu_value(args.gpu),
                journal_file_names=args.journal,
                case_file_name=args.case_file,
                case_data_file_name=args.case_data,
                cwd=args.cwd,
                fluent_path=args.fluent_path,
                graphics_driver=args.graphics_driver,
                start_timeout=args.start_timeout,
                cleanup_on_exit=False,  # Fluent must outlive this short-lived script
                additional_arguments=args.additional_arguments,
            )
        finally:
            _pfl.config.launch_fluent_stdout = prev_stdout
            _pfl.config.launch_fluent_stderr = prev_stderr
            log_handle.close()
        if getattr(result, "status", "") != "ok":
            return {"error": getattr(result, "message", "launch failed"),
                    "error_code": getattr(result, "error_code", None),
                    "hint": f"see {log_path} for the Fluent launch log"}
        props = backend._solver.connection_properties
        session = {
            "mode": "launch",
            "ip": getattr(props, "ip", None) or "127.0.0.1",
            "port": getattr(props, "port", None),
            "password": getattr(props, "password", None),
            "ui_mode": args.ui_mode,
            "processor_count": args.processor_count,
            "dimension": args.dimension,
            "launched_at": datetime.now(timezone.utc).isoformat(),
        }
        if session["port"] is None:
            _delete_session(args)
            return {"error": "launched Fluent but could not read connection "
                             "properties; session not persisted"}
        path = _save_session(args, session)
        return {
            "status": "ok",
            "already_connected": False,
            "mode": "launch",
            "endpoint": f"{session['ip']}:{session['port']}",
            "session_file": str(path),
            "message": getattr(result, "message", None),
        }

    return asyncio.run(_run())


def cmd_disconnect(args: argparse.Namespace) -> Any:
    """Shut the recorded Fluent session down and clear the session file.

    Attaches with ``cleanup_on_exit=True`` (PyFluent only kills the server
    when the *exiting* client asks for cleanup) and exits with a timeout +
    wait, so the solver process is verified gone instead of orphaned.
    """
    session = _load_session(args)
    if session is None:
        return {"status": "ok", "message": "no active session"}
    endpoint = f"{session.get('ip')}:{session.get('port')}"

    from ansys.fluent.core import connect_to_fluent

    notes: list[str] = []
    try:
        solver = connect_to_fluent(
            ip=session.get("ip") or "127.0.0.1",
            port=session.get("port"),
            password=session.get("password"),
            cleanup_on_exit=True,
        )
        solver.exit(timeout=args.timeout, wait=True)
    except Exception as exc:  # noqa: BLE001 — session may already be dead
        notes.append(
            f"shutdown path reported: {type(exc).__name__}: {exc} "
            "(the session may have already exited)"
        )
    _delete_session(args)
    result: dict[str, Any] = {"status": "ok", "endpoint": endpoint}
    if notes:
        result["notes"] = notes
    return result


def cmd_session_status(args: argparse.Namespace) -> Any:
    """Report session-file state and probe whether Fluent is still alive."""
    session = _load_session(args)
    if session is None:
        return {
            "connected": False,
            "session_file": None,
            "notes": ["No session file. Run `connect` to launch or attach Fluent."],
        }
    info: dict[str, Any] = {
        "connected": False,
        "session_file": str(_session_file(args)),
        "mode": session.get("mode"),
        "endpoint": f"{session.get('ip')}:{session.get('port')}",
        "launched_at": session.get("launched_at") or session.get("attached_at"),
        "ui_mode": session.get("ui_mode"),
        "processor_count": session.get("processor_count"),
        "dimension": session.get("dimension"),
    }
    if args.no_ping:
        info["notes"] = ["Liveness probe skipped (--no-ping)."]
        return info
    try:
        asyncio.run(_attach_from_session(args))
    except SystemExit:
        info["notes"] = [
            "Recorded Fluent is unreachable; the session is stale. Run `connect` again."
        ]
        return info
    info["connected"] = True
    info["notes"] = ["Live session verified via gRPC attach."]
    return info


# ---------------------------------------------------------------------------
# Schema / introspection tools
# ---------------------------------------------------------------------------

def cmd_find_api(args: argparse.Namespace) -> Any:
    """MCP ``find_api``: lexical search over the bundled Fluent API catalog.

    Works offline (no live session needed) — same as the MCP tool.
    """
    async def _run() -> Any:
        return await _new_backend().find_api(
            args.query,
            top_k=args.top_k,
            kinds=args.kind,
            under=args.under,
        )

    hits = asyncio.run(_run())
    if not args.compact:
        return hits
    slim: list[dict[str, Any]] = []
    for hit in hits:
        desc = hit.get("docstring") or hit.get("description") or ""
        one_line = desc.strip().split("\n", 1)[0][:160] if isinstance(desc, str) else ""
        slim.append({
            "path": hit.get("path"),
            "kind": hit.get("kind"),
            "score": hit.get("score"),
            "summary": one_line,
        })
    return slim


def cmd_get_state(args: argparse.Namespace) -> Any:
    """MCP ``get_state``: read live settings state for dotted paths."""
    paths = args.path
    if args.key is not None:
        if not paths or len(paths) != 1:
            raise SystemExit("--key requires exactly one collection path in --path")
        base = paths[0].rstrip(".")
        if base.endswith("]"):
            raise SystemExit("--path already indexes a named object; drop --key")
        if any(ch in args.key for ch in "\"'[]"):
            raise SystemExit("--key contains invalid characters")
        paths = [f"{base}[{args.key}]"]

    def _call(backend: Any) -> Any:
        return backend.get_state(paths=paths)

    return _require_backend(_call, args)


def cmd_list_named_objects(args: argparse.Namespace) -> Any:
    """MCP ``list_named_objects``: enumerate named-object collections."""
    def _call(backend: Any) -> Any:
        return backend.list_named_objects()

    mapping = _require_backend(_call, args)
    if args.limit is None and not args.offset:
        return mapping
    from ansys.fluent.mcp.common.errors import InvalidArgumentsError

    if args.limit is not None and args.limit < 1:
        raise InvalidArgumentsError("limit must be >= 1")
    sliced: dict[str, Any] = {}
    totals: dict[str, int] = {}
    for coll, names in (mapping or {}).items():
        lst = list(names or [])
        totals[coll] = len(lst)
        end = args.offset + args.limit if args.limit is not None else None
        sliced[coll] = lst[args.offset:end]
    sliced["_pagination"] = {
        "offset": args.offset,
        "limit": args.limit,
        "totals": totals,
        "truncated": any(
            (args.limit is not None and totals[c] > args.offset + args.limit) or args.offset > 0
            for c in totals
        ),
    }
    return sliced


def cmd_select_named_objects(args: argparse.Namespace) -> Any:
    """MCP ``select_named_objects``: glob-expand a named-object collection."""
    from ansys.fluent.mcp.common.base import select_named_objects_from_mapping

    def _call(backend: Any) -> Any:
        return backend.list_named_objects()

    mapping = _require_backend(_call, args)
    return select_named_objects_from_mapping(
        mapping,
        collection=args.collection,
        pattern=args.pattern,
        include_shadows=not args.no_include_shadows,
        exclude=args.exclude,
    )


# ---------------------------------------------------------------------------
# Code execution
# ---------------------------------------------------------------------------

def _read_code(args: argparse.Namespace) -> str:
    """Read ``--code`` or ``--code-file`` input."""
    if args.code and args.code_file:
        raise SystemExit("pass either --code or --code-file, not both")
    if args.code_file:
        path = Path(args.code_file)
        if not path.is_file():
            raise SystemExit(f"code file not found: {path}")
        code = path.read_text(encoding="utf-8")
    else:
        code = args.code or ""
    if not code.strip():
        raise SystemExit("code must be a non-empty string")
    return code


def cmd_run_code(args: argparse.Namespace) -> Any:
    """MCP ``run_code``: sandboxed Python against the live session.

    ``solver`` (alias ``session``) is pre-injected. Imports are restricted to
    the allowlist (math/json/itertools/functools/collections/dataclasses/
    typing/ansys.fluent.core); reflection writes are blocked.
    """
    code = _read_code(args)

    def _call(backend: Any) -> Any:
        return backend.run_code(code, filename=args.filename)

    result = _require_backend(_call, args)
    # Backend invalidates its own caches per call; nothing else to do here.
    return result


def cmd_validate_code(args: argparse.Namespace) -> Any:
    """MCP ``validate_code``: dry-run syntax/safety checks, no side effects.

    Offline when no session is recorded; with a live session the backend
    also runs the live-compile check and intent guards.
    """
    code = _read_code(args)

    async def _run() -> Any:
        backend = _new_backend()
        session = _load_session(args)
        if session is not None:
            try:
                result = await backend.connect(
                    ip=session.get("ip") or "127.0.0.1",
                    port=session.get("port"),
                    password=session.get("password"),
                )
            except Exception:
                result = None
            if getattr(result, "status", "") != "ok":
                backend = _new_backend()  # fall back to offline validation
        return await backend.validate_code(code)

    return asyncio.run(_run())


# ---------------------------------------------------------------------------
# Reporting / visualization
# ---------------------------------------------------------------------------

def cmd_screenshot(args: argparse.Namespace) -> Any:
    """MCP ``screenshot``: capture the current model view as a PNG file.

    Writes the decoded image to ``--out`` (default ``./screenshot.png``)
    instead of returning base64 in the JSON envelope.
    """
    out_path = Path(args.out or "screenshot.png").resolve()

    def _call(backend: Any) -> Any:
        return backend.screenshot(view=args.view)

    payload = _require_backend(_call, args)
    data = payload.get("data") if isinstance(payload, dict) else None
    if not data:
        return {"error": "screenshot returned no image data", "raw": payload}
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_bytes(base64.b64decode(data))
    return {
        "format": "png",
        "path": str(out_path),
        "size_bytes": out_path.stat().st_size,
        "view": args.view,
    }


def cmd_summarize_setup(args: argparse.Namespace) -> Any:
    """MCP ``summarize_setup``: Fluent Report > Summary in one JSON call."""
    async def _run() -> Any:
        backend = await _attach_from_session(args)
        tmp_file = tempfile.NamedTemporaryFile(suffix=".txt", delete=False)
        tmp_file.close()
        tmp = str(Path(tmp_file.name).resolve())
        tmp_fluent = tmp.replace("\\", "/")
        snippet = (
            "session.settings.results.report.summary("
            f"write_to_file=True, file_name={tmp_fluent!r})"
        )
        result = await backend.run_code(snippet)
        if getattr(result, "status", "") != "ok":
            return {
                "error": (
                    getattr(result, "stderr", None)
                    or getattr(result, "message", None)
                    or "summary command failed"
                )
            }
        fp = Path(tmp)
        content = ""
        try:
            if fp.exists():
                content = fp.read_text(encoding="utf-8", errors="replace")
                fp.unlink(missing_ok=True)
        except OSError:
            pass
        return {"summary": content or "(no output)"}

    return asyncio.run(_run())


def cmd_simulation_report(args: argparse.Namespace) -> Any:
    """MCP ``simulation_report``: generate / export / list simulation reports."""
    action = (args.action or "list").strip().lower()
    valid = {"generate", "export_html", "export_pdf", "export_pptx", "list"}
    if action not in valid:
        return {"error": f"invalid action {action!r}", "valid_actions": sorted(valid)}

    async def _run() -> Any:
        backend = await _attach_from_session(args)
        report_name = args.report_name
        output_path = args.output_path

        if action == "list":
            snippet = (
                "session.settings.results.report"
                ".simulation_reports.list_simulation_reports()"
            )
            result = await backend.run_code(snippet)
            if getattr(result, "status", "") != "ok":
                return {"error": getattr(result, "stderr", None)
                        or "list_simulation_reports failed"}
            return {
                "reports": getattr(result, "return_value", None),
                "stdout": getattr(result, "stdout", "") or None,
            }

        if action == "generate":
            snippet = (
                "session.settings.results.report"
                ".simulation_reports"
                ".generate_simulation_report("
                f"report_name={report_name!r})"
            )
            out = None
        else:
            out_path: Optional[str] = output_path
            if out_path is None:
                suffix = {"export_html": ".html", "export_pdf": ".pdf",
                          "export_pptx": ".pptx"}[action]
                out_file = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
                out_file.close()
                out_path = out_file.name
            out = str(Path(out_path).resolve()).replace("\\", "/")
            method = {
                "export_html": "export_simulation_report_as_html",
                "export_pdf": "export_simulation_report_as_pdf",
                "export_pptx": "export_simulation_report_as_pptx",
            }[action]
            if action == "export_html":
                args_part = f"report_name={report_name!r}, output_dir={out!r}"
            else:
                args_part = f"report_name={report_name!r}, file_name={out!r}"
            snippet = (
                "session.settings.results.report"
                ".simulation_reports"
                f".{method}({args_part})"
            )

        result = await backend.run_code(snippet)
        if getattr(result, "status", "") != "ok":
            return {"error": getattr(result, "stderr", None) or f"{action} failed"}
        return {
            "action": action,
            "report_name": report_name,
            "output_path": out,
            "stdout": getattr(result, "stdout", "") or None,
            "note": f"Simulation report {action} completed.",
        }

    return asyncio.run(_run())


# ---------------------------------------------------------------------------
# Domain tools (solve/lib implementations)
# ---------------------------------------------------------------------------

def _domain_call(impl: Any, args: argparse.Namespace, **kwargs: Any) -> Any:
    """Run a ``solve.lib`` domain-tool impl against the attached backend."""
    async def _run() -> Any:
        backend = await _attach_from_session(args)
        return await impl(backend, **kwargs)

    return asyncio.run(_run())


def cmd_mesh_quality(args: argparse.Namespace) -> Any:
    """MCP ``mesh_quality``: live skewness / orthogonal quality / counts."""
    from ansys.fluent.mcp.solve.lib.mesh_tools import mesh_quality_impl

    return _domain_call(mesh_quality_impl, args, include_check=args.include_check)


def cmd_list_fields(args: argparse.Namespace) -> Any:
    """MCP ``list_fields``: enumerate scalar/vector fields in the case."""
    from ansys.fluent.mcp.solve.lib.discovery_tools import list_fields_impl

    return _domain_call(list_fields_impl, args, scope=args.scope)


def cmd_compare_files(args: argparse.Namespace) -> Any:
    """MCP ``compare_files``: diff two case/mesh files in ephemeral sessions."""
    from ansys.fluent.mcp.solve.lib.compare_tools import compare_files_impl

    async def _run() -> Any:
        backend = await _attach_from_session(args)
        return await compare_files_impl(backend, path_a=args.a, path_b=args.b)

    return asyncio.run(_run())


def cmd_probe_path(args: argparse.Namespace) -> Any:
    """MCP ``probe_path``: batch pre-flight probe for settings paths."""
    from ansys.fluent.mcp.solve.lib.schema_probe_tools import probe_path_impl

    return _domain_call(probe_path_impl, args, paths=args.path)


def cmd_get_active_status(args: argparse.Namespace) -> Any:
    """MCP ``get_active_status``: batch active-status probe."""
    from ansys.fluent.mcp.solve.lib.schema_probe_tools import get_active_status_impl

    return _domain_call(get_active_status_impl, args, paths=args.path)


def cmd_get_allowed_values(args: argparse.Namespace) -> Any:
    """MCP ``get_allowed_values``: batch enum allowed-values probe."""
    from ansys.fluent.mcp.solve.lib.schema_probe_tools import get_allowed_values_impl

    return _domain_call(get_allowed_values_impl, args, paths=args.path)


def cmd_describe_named_object_template(args: argparse.Namespace) -> Any:
    """MCP ``describe_named_object_template``: field shape of a fresh child."""
    from ansys.fluent.mcp.solve.lib.schema_probe_tools import (
        describe_named_object_template_impl,
    )

    return _domain_call(describe_named_object_template_impl, args, path=args.path)


def cmd_describe_path(args: argparse.Namespace) -> Any:
    """MCP ``describe_path``: unified per-path descriptor (probe+values+template)."""
    from ansys.fluent.mcp.solve.lib.schema_probe_tools import describe_path_impl

    return _domain_call(
        describe_path_impl,
        args,
        paths=args.path,
        include_template=not args.no_template,
    )


def cmd_solver_status(args: argparse.Namespace) -> Any:
    """MCP ``solver_status``: initialized / iterations / residuals summary."""
    def _call(backend: Any) -> Any:
        return backend.solver_status()

    return _require_backend(_call, args)


# ---------------------------------------------------------------------------
# Parser assembly
# ---------------------------------------------------------------------------

def _add_session_arg(p: argparse.ArgumentParser) -> None:
    """Add the shared ``--session-file`` option."""
    p.add_argument(
        "--session-file",
        default=None,
        help=(
            "Session file path (default: "
            f"{DEFAULT_SESSION_FILE}; env FLUENT_TOOLS_SESSION)"
        ),
    )


def _add_pretty_arg(p: argparse.ArgumentParser) -> None:
    """Add the shared ``--pretty`` output option."""
    p.add_argument("--pretty", action="store_true",
                   help="Pretty-print the JSON envelope (stdout or --json-out)")


def _paths_arg(p: argparse.ArgumentParser, *, required: bool = True) -> None:
    """Add a repeatable ``--path`` option used by introspection commands."""
    p.add_argument(
        "--path",
        action="append",
        required=required,
        help="Dotted Fluent settings path (repeatable for batch calls)",
    )


def build_parser() -> argparse.ArgumentParser:
    """Build the argument parser with one subcommand per MCP tool."""
    parser = argparse.ArgumentParser(
        prog="fluent_cli.py",
        description=(
            "PyFluent-MCP tools as a standalone CLI. Each subcommand maps 1:1 "
            "to an ansys-fluent-mcp MCP tool and prints its JSON envelope to "
            "stdout. Run `connect` once to launch/attach Fluent; the other "
            "commands attach to that session automatically."
        ),
    )
    sub = parser.add_subparsers(dest="command", required=True, metavar="COMMAND")

    # -- session lifecycle -------------------------------------------------
    p = sub.add_parser(
        "connect",
        help="Launch (or attach to) a Fluent session and persist it",
        description=(
            "Launch Fluent headless (default) or attach to an existing "
            "session with --ip/--port/--password or --server-info-file. "
            "The live session's ip/port/password are persisted to the "
            "session file so all other commands can attach. If a live "
            "session is already recorded, reports it instead of launching "
            "again unless --force is given."
        ),
    )
    _add_session_arg(p)
    _add_pretty_arg(p)
    p.add_argument("--ip", default=None, help="Attach mode: Fluent host IP")
    p.add_argument("--port", type=int, default=None, help="Attach mode: Fluent gRPC port")
    p.add_argument("--password", default=None, help="Attach mode: connection password")
    p.add_argument("--server-info-file", default=None,
                   help="Attach mode: server-info file written by Fluent")
    p.add_argument("--force", action="store_true",
                   help="Launch even if a session file already exists")
    p.add_argument("--processor-count", type=int, default=1,
                   help="Launch mode: parallel CPU cores (default 1)")
    p.add_argument("--precision", choices=["single", "double"], default="double")
    p.add_argument("--ui-mode", default="hidden_gui",
                   help="Launch mode: gui | hidden_gui | no_gui | no_gui_or_graphics "
                        "(default hidden_gui: headless but screenshots work; "
                        "no_gui is lighter but cannot export pictures)")
    p.add_argument("--dimension", type=int, choices=[2, 3], default=None,
                   help="Launch mode: 2D or 3D (auto-detected when omitted)")
    p.add_argument("--mode", choices=["solver", "meshing"], default=None,
                   help="Launch mode: solver (default) or meshing session")
    p.add_argument("--gpu", default=None,
                   help="Launch mode: true/false or comma-separated GPU ids")
    p.add_argument("--product-version", default=None,
                   help="Launch mode: e.g. 261 for 2026R1")
    p.add_argument("--journal", action="append", default=None,
                   help="Launch mode: journal file to replay (repeatable)")
    p.add_argument("--case-file", default=None,
                   help="Launch mode: case file to read at startup")
    p.add_argument("--case-data", default=None,
                   help="Launch mode: data file to read at startup")
    p.add_argument("--cwd", default=None, help="Launch mode: Fluent working directory")
    p.add_argument("--fluent-path", default=None, help="Launch mode: custom fluent.exe path")
    p.add_argument("--graphics-driver", default=None,
                   help="Launch mode: graphics driver override")
    p.add_argument("--start-timeout", type=int, default=300,
                   help="Launch mode: startup timeout seconds (default 300)")
    p.add_argument("--additional-arguments", default=None,
                   help="Launch mode: extra raw arguments passed to fluent.exe")
    p.set_defaults(func=cmd_connect)

    p = sub.add_parser("disconnect", help="Shut down the recorded Fluent session",
                       description=("Attaches with cleanup_on_exit and exits Fluent, "
                                    "then deletes the session file."))
    _add_session_arg(p)
    _add_pretty_arg(p)
    p.add_argument("--timeout", type=int, default=30,
                   help="Seconds to wait for the exit request before force-terminating "
                        "(default 30)")
    p.set_defaults(func=cmd_disconnect)

    p = sub.add_parser("session_status", help="Report session state and liveness",
                       description=("Reads the session file and (unless --no-ping) "
                                    "verifies the recorded Fluent answers over gRPC."))
    _add_session_arg(p)
    _add_pretty_arg(p)
    p.add_argument("--no-ping", action="store_true", help="Skip the liveness probe")
    p.set_defaults(func=cmd_session_status)

    # -- schema discovery --------------------------------------------------
    p = sub.add_parser("find_api", help="Lexical search over the Fluent API catalog")
    _add_session_arg(p)
    _add_pretty_arg(p)
    p.add_argument("--query", required=True, help="Search text, e.g. 'temperature wall'")
    p.add_argument("--top-k", type=int, default=10)
    p.add_argument("--kind", action="append", default=None,
                   help="Filter: Parameter | Command | Object | Group (repeatable)")
    p.add_argument("--under", default=None, help="Restrict search to a path prefix")
    p.add_argument("--compact", action="store_true",
                   help="Return slim hits (path/kind/score/summary only)")
    p.set_defaults(func=cmd_find_api)

    p = sub.add_parser("get_help", help="Docstring, children, allowed values for a path")
    _add_session_arg(p)
    _add_pretty_arg(p)
    p.add_argument("--path", required=True, help="Dotted Fluent settings path")
    p.set_defaults(func=lambda a: _require_backend(
        lambda b: b.get_help(a.path), a))

    p = sub.add_parser("get_state", help="Read live settings state for paths")
    _add_session_arg(p)
    _add_pretty_arg(p)
    _paths_arg(p, required=False)
    p.add_argument("--key", default=None,
                   help="Fetch one named object: --path <collection> --key <name>")
    p.set_defaults(func=cmd_get_state)

    p = sub.add_parser("get_targeted_context",
                       help="Batched context: active+state+named+children+allowed")
    _add_session_arg(p)
    _add_pretty_arg(p)
    _paths_arg(p)
    p.add_argument("--named-object-type", action="append", default=None,
                   help="Named-object family to include (repeatable)")
    p.add_argument("--instance-state-fetch", action="append", default=None,
                   help="Named-object instance path whose state to fetch (repeatable)")
    p.set_defaults(func=lambda a: _require_backend(
        lambda b: b.get_targeted_context(
            paths_to_check=a.path,
            named_object_types=a.named_object_type or [],
            instance_state_fetch=a.instance_state_fetch or [],
        ), a))

    # -- named objects -----------------------------------------------------
    p = sub.add_parser("list_named_objects", help="List named-object collections")
    _add_session_arg(p)
    _add_pretty_arg(p)
    p.add_argument("--limit", type=int, default=None,
                   help="Pagination: max names per collection")
    p.add_argument("--offset", type=int, default=0)
    p.set_defaults(func=cmd_list_named_objects)

    p = sub.add_parser("find_named_object",
                       help="Resolve a symbolic name across all collections")
    _add_session_arg(p)
    _add_pretty_arg(p)
    p.add_argument("--name", required=True, help="e.g. 'inlet-1'")
    p.set_defaults(func=lambda a: _require_backend(
        lambda b: b.find_named_object(a.name), a))

    p = sub.add_parser("select_named_objects",
                       help="Glob-expand a named-object collection")
    _add_session_arg(p)
    _add_pretty_arg(p)
    p.add_argument("--collection", required=True,
                   help="e.g. setup.boundary_conditions.wall")
    p.add_argument("--pattern", default="*", help="Unix-shell glob (default *)")
    p.add_argument("--no-include-shadows", action="store_true",
                   help="Drop '-shadow' walls from the result")
    p.add_argument("--exclude", action="append", default=None,
                   help="Glob patterns to subtract (repeatable)")
    p.set_defaults(func=cmd_select_named_objects)

    # -- code execution ----------------------------------------------------
    p = sub.add_parser("run_code", help="Sandboxed Python against the live session")
    _add_session_arg(p)
    _add_pretty_arg(p)
    p.add_argument("--code", default=None, help="Python source (use --code-file for long code)")
    p.add_argument("--code-file", default=None, help="Path to a .py file to execute")
    p.add_argument("--filename", default="<fluent_cli>",
                   help="Label used in tracebacks/telemetry")
    p.set_defaults(func=cmd_run_code)

    p = sub.add_parser("validate_code", help="Dry-run code checks without side effects")
    _add_session_arg(p)
    _add_pretty_arg(p)
    p.add_argument("--code", default=None)
    p.add_argument("--code-file", default=None)
    p.set_defaults(func=cmd_validate_code)

    # -- reporting / visualization ----------------------------------------
    p = sub.add_parser("screenshot", help="Save the current model view as PNG")
    _add_session_arg(p)
    _add_pretty_arg(p)
    p.add_argument("--out", default=None, help="Output PNG path (default ./screenshot.png)")
    p.add_argument("--view", default=None, help="Optional view/camera preset")
    p.set_defaults(func=cmd_screenshot)

    p = sub.add_parser("summarize_setup",
                       help="Full setup summary (models/materials/BCs/solver)")
    _add_session_arg(p)
    _add_pretty_arg(p)
    p.set_defaults(func=cmd_summarize_setup)

    p = sub.add_parser("simulation_report", help="Generate/export simulation reports")
    _add_session_arg(p)
    _add_pretty_arg(p)
    p.add_argument("--action", default="list",
                   choices=["generate", "export_html", "export_pdf", "export_pptx", "list"])
    p.add_argument("--report-name", default="default-report")
    p.add_argument("--output-path", default=None,
                   help="Output file/dir; defaults to a temp path")
    p.set_defaults(func=cmd_simulation_report)

    # -- domain tools ------------------------------------------------------
    p = sub.add_parser("mesh_quality", help="Mesh counts + skewness/orthogonal quality")
    _add_session_arg(p)
    _add_pretty_arg(p)
    p.add_argument("--include-check", action="store_true",
                   help="Also embed the structured mesh.check() payload")
    p.set_defaults(func=cmd_mesh_quality)

    p = sub.add_parser("list_fields", help="Enumerate scalar/vector fields")
    _add_session_arg(p)
    _add_pretty_arg(p)
    p.add_argument("--scope", default="any",
                   choices=["any", "cell", "node", "face"],
                   help="Field domain filter (default any)")
    p.set_defaults(func=cmd_list_fields)

    p = sub.add_parser("compare_files", help="Diff two case/mesh files")
    _add_session_arg(p)
    _add_pretty_arg(p)
    p.add_argument("--a", required=True, help="First case/mesh file")
    p.add_argument("--b", required=True, help="Second case/mesh file")
    p.set_defaults(func=cmd_compare_files)

    p = sub.add_parser("probe_path", help="Batch probe: exists/active/creatable/kind")
    _add_session_arg(p)
    _add_pretty_arg(p)
    _paths_arg(p)
    p.set_defaults(func=cmd_probe_path)

    p = sub.add_parser("get_active_status", help="Batch {path: is_active} probe")
    _add_session_arg(p)
    _add_pretty_arg(p)
    _paths_arg(p)
    p.set_defaults(func=cmd_get_active_status)

    p = sub.add_parser("get_allowed_values", help="Batch {path: [allowed values]} probe")
    _add_session_arg(p)
    _add_pretty_arg(p)
    _paths_arg(p)
    p.set_defaults(func=cmd_get_allowed_values)

    p = sub.add_parser("describe_named_object_template",
                       help="Field shape of a fresh child under a collection")
    _add_session_arg(p)
    _add_pretty_arg(p)
    p.add_argument("--path", required=True, help="NamedObject collection path")
    p.set_defaults(func=cmd_describe_named_object_template)

    p = sub.add_parser("describe_path",
                       help="Unified descriptor: probe + values + template + command args")
    _add_session_arg(p)
    _add_pretty_arg(p)
    _paths_arg(p)
    p.add_argument("--no-template", action="store_true",
                   help="Skip the named-object template probe")
    p.set_defaults(func=cmd_describe_path)

    # -- solver ------------------------------------------------------------
    p = sub.add_parser("solver_status",
                       help="Initialized/iterations/residuals/convergence summary")
    _add_session_arg(p)
    _add_pretty_arg(p)
    p.set_defaults(func=cmd_solver_status)

    # Shared --json-out: every subcommand can write its JSON envelope to a
    # file instead of stdout. Added in one loop so future subcommands inherit
    # it automatically; screenshot's --out keeps meaning the PNG path.
    for cmd_parser in sub.choices.values():
        cmd_parser.add_argument(
            "--json-out",
            default=None,
            help="Write the JSON envelope to this file (UTF-8) instead of "
                 "stdout; error envelopes are written here too",
        )

    return parser


def main(argv: Optional[list[str]] = None) -> None:
    """CLI entry point."""
    parser = build_parser()
    args = parser.parse_args(argv)

    # Set before dispatch so envelopes emitted via _fail (e.g. exit code 3,
    # no live session) also land in the file.
    global _JSON_OUT
    if getattr(args, "json_out", None):
        _JSON_OUT = Path(args.json_out)

    debug = bool(os.environ.get("FLUENT_CLI_DEBUG"))

    # Backend library logs are noise on stderr; keep them behind FLUENT_CLI_DEBUG.
    import logging

    logging.basicConfig(level=logging.DEBUG if debug else logging.WARNING,
                        stream=sys.stderr)

    # PyFluent's transcript stream registers a `print` callback, so Fluent
    # chatter (scheme errors, command echoes) would otherwise land on real
    # stdout and corrupt the JSON contract. run_code needs that callback to
    # stay alive — it captures transcript output via an inner
    # redirect_stdout — so instead of stopping the stream we swallow any
    # writes that happen outside a run_code capture and print the JSON
    # envelope to the restored stdout afterwards.
    import contextlib
    import io as _io

    try:
        if debug:
            result = args.func(args)
        else:
            with contextlib.redirect_stdout(_io.StringIO()):
                result = args.func(args)
    except SystemExit:
        raise
    except ImportError as exc:
        _fail(
            f"missing dependency: {exc}. Use the Python environment that has "
            "`ansys-fluent-core` installed (e.g. the conda env used by "
            "ansys-fluent-mcp), or `pip install ansys-fluent-core pydantic`."
        )
    except Exception as exc:  # noqa: BLE001 — CLI boundary converts to JSON
        if debug:
            raise
        _fail(f"{type(exc).__name__}: {exc}")
    _emit(result, pretty=getattr(args, "pretty", False))
    # Shell-friendly failure signal: an error envelope on stdout also
    # produces a non-zero exit code (except where noted in the README).
    if isinstance(result, dict) and ("error" in result or result.get("status") == "error"):
        raise SystemExit(1)
    if getattr(result, "status", "") == "error":
        raise SystemExit(1)


if __name__ == "__main__":
    main()
