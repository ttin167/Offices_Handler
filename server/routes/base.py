# -*- coding: utf-8 -*-
"""
server.routes.base — Modular Route Base, Helpers & Registry for Antigravity Web Studio.
"""

from __future__ import annotations

import json
import re
import urllib.parse
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

