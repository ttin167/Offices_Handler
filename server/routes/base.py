# -*- coding: utf-8 -*-
"""
server.routes.base — Modular Route Base, Helpers & Registry for Antigravity Web Studio.
"""

from __future__ import annotations

import json
import os
import re
import tempfile
import time
import urllib.parse
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple, Union


class Router:
    """Singleton Route Registry and Request Dispatcher for Python http.server."""

    _instance: Optional[Router] = None

    def __init__(self) -> None:
        # Exact match: (METHOD, exact_path) -> handler_func
        self._exact_routes: Dict[Tuple[str, str], Callable] = {}
        # Pattern match: (METHOD, regex_pattern, param_names) -> handler_func
        self._pattern_routes: List[Tuple[str, re.Pattern, List[str], Callable]] = []

    @classmethod
    def get_instance(cls) -> Router:
        if cls._instance is None:
            cls._instance = Router()
        return cls._instance

    @classmethod
    def reset(cls) -> None:
        cls._instance = None

    def register(self, method: str, path: str, func: Callable) -> None:
        """Registers a route for an HTTP method and path."""
        method_upper = method.upper()
        if "<" in path and ">" in path:
            # Parametric route, e.g. /api/download/<filename>
            param_names = []

            def repl(m: re.Match) -> str:
                param_names.append(m.group(1))
                return r"([^/]+)"

            pattern_str = "^" + re.sub(r"<([a-zA-Z0-9_]+)>", repl, path) + "$"
            regex = re.compile(pattern_str)
            self._pattern_routes.append((method_upper, regex, param_names, func))
        else:
            self._exact_routes[(method_upper, path)] = func

    def get(self, path: str) -> Callable:
        def decorator(func: Callable) -> Callable:
            self.register("GET", path, func)
            return func
        return decorator

    def post(self, path: str) -> Callable:
        def decorator(func: Callable) -> Callable:
            self.register("POST", path, func)
            return func
        return decorator

    def dispatch(self, handler: Any, method: str, path: str) -> bool:
        """
        Attempts to dispatch request to registered routes.
        Returns True if handled, False otherwise.
        """
        method_upper = method.upper()

        # 1. Exact match
        if (method_upper, path) in self._exact_routes:
            func = self._exact_routes[(method_upper, path)]
            func(handler)
            return True

        # 2. Pattern match
        for r_method, regex, param_names, func in self._pattern_routes:
            if r_method == method_upper:
                mo = regex.match(path)
                if mo:
                    params = dict(zip(param_names, mo.groups()))
                    func(handler, **params)
                    return True

        return False


# Global decorator shortcuts
router = Router.get_instance()


def send_json(handler: Any, data: Any, status: int = 200) -> None:
    """Helper to send JSON response with CORS headers."""
    if hasattr(handler, "_send_json"):
        handler._send_json(data, status=status)
    else:
        encoded = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
        handler.send_response(status)
        if hasattr(handler, "_send_cors_headers"):
            handler._send_cors_headers()
        handler.send_header("Content-Type", "application/json; charset=utf-8")
        handler.send_header("Content-Length", str(len(encoded)))
        handler.end_headers()
        handler.wfile.write(encoded)


def send_error(handler: Any, error_msg: str, status: int = 400) -> None:
    """Helper to send standardized error JSON response."""
    send_json(handler, {"success": False, "error": error_msg}, status=status)


def read_json_body(handler: Any) -> Dict[str, Any]:
    """Helper to read and parse JSON body from request."""
    if hasattr(handler, "_read_body_json"):
        return handler._read_body_json()

    content_length = int(handler.headers.get("Content-Length", 0))
    if content_length <= 0:
        return {}
    raw_body = handler.rfile.read(content_length).decode("utf-8", errors="replace")
    return json.loads(raw_body) if raw_body else {}


def get_base_dir() -> str:
    """Return workspace root directory."""
    return str(Path(__file__).resolve().parent.parent.parent)


def get_upload_dir() -> str:
    """Ensure and return `.web_uploads` directory."""
    d = os.path.join(get_base_dir(), ".web_uploads")
    os.makedirs(d, exist_ok=True)
    return d


def get_output_dir() -> str:
    """Ensure and return `.web_outputs` directory."""
    d = os.path.join(get_base_dir(), ".web_outputs")
    os.makedirs(d, exist_ok=True)
    return d


def resolve_workspace_file(path_str: Optional[str]) -> Optional[str]:
    """Safely resolve a submitted file path within allowed sandbox directories."""
    if not path_str or not isinstance(path_str, str):
        return None
    cleaned = path_str.strip()
    if not cleaned:
        return None

    base_dir = get_base_dir()
    upload_dir = get_upload_dir()
    output_dir = get_output_dir()
    allowed_roots = [
        os.path.abspath(base_dir),
        os.path.abspath(upload_dir),
        os.path.abspath(output_dir),
        os.path.abspath(os.path.join(base_dir, "uploads")),
        os.path.abspath(tempfile.gettempdir()),
    ]

    def _is_safe(candidate_abs: str) -> bool:
        norm_cand = os.path.normpath(candidate_abs)
        for root in allowed_roots:
            norm_root = os.path.normpath(root)
            try:
                if os.path.commonpath([norm_root, norm_cand]) == norm_root:
                    return True
            except (ValueError, OSError):
                continue
        return False

    candidates = [
        cleaned,
        os.path.join(upload_dir, os.path.basename(cleaned)),
        os.path.join(output_dir, os.path.basename(cleaned)),
        os.path.join(base_dir, "uploads", os.path.basename(cleaned)),
        os.path.join(base_dir, cleaned),
    ]
    for cand in candidates:
        if os.path.isfile(cand):
            abs_p = os.path.abspath(cand)
            if _is_safe(abs_p):
                return abs_p
    return None


def build_output_filepath(src_path: str, suffix: str, ext: Optional[str] = None) -> Tuple[str, str]:
    """Generate unique output filename and path in `.web_outputs` plus timestamp."""
    base_stem = os.path.splitext(os.path.basename(src_path))[0]
    safe_stem = re.sub(r"[^a-zA-Z0-9_\-\.]", "_", base_stem)[:50] or "output"
    target_ext = (ext or os.path.splitext(src_path)[1] or ".xlsx")
    if not target_ext.startswith("."):
        target_ext = f".{target_ext}"
    filename = f"{safe_stem}_{suffix}_{int(time.time())}{target_ext}"
    return filename, os.path.join(get_output_dir(), filename)


def download_url_for(filename: str) -> str:
    """Return encoded `/api/download/<filename>` URL."""
    return f"/api/download/{urllib.parse.quote(filename)}"


