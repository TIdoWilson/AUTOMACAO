#!/usr/bin/env python3
"""
Espelha o conteudo publico de um site para uso offline.

Importante:
- Este script baixa apenas o que esta publicamente acessivel pelo navegador.
- Ele nao consegue obter backend, banco de dados, APIs privadas ou codigo do servidor.
- Use somente em sites que voce tem permissao para copiar.
"""

from __future__ import annotations

import argparse
import hashlib
import html
import json
import mimetypes
import os
import posixpath
import re
import shutil
import sys
import time
from collections import deque
from contextlib import suppress
from http.cookiejar import Cookie, CookieJar
from html.parser import HTMLParser
from pathlib import Path, PurePosixPath
from typing import Iterable
from urllib.error import HTTPError, URLError
from urllib.parse import urljoin, urlsplit, urlunsplit
from urllib.request import HTTPCookieProcessor, Request, build_opener


HTML_EXTENSIONS = {
    "",
    ".html",
    ".htm",
    ".php",
    ".asp",
    ".aspx",
    ".jsp",
    ".jspx",
    ".cfm",
    ".cgi",
    ".pl",
    ".shtml",
}

URL_ATTRS = {
    ("a", "href"),
    ("link", "href"),
    ("img", "src"),
    ("script", "src"),
    ("iframe", "src"),
    ("source", "src"),
    ("audio", "src"),
    ("video", "src"),
    ("embed", "src"),
    ("track", "src"),
    ("object", "data"),
    ("form", "action"),
    ("input", "src"),
    ("input", "formaction"),
    ("video", "poster"),
}

INVALID_SEGMENT_CHARS = re.compile(r'[<>:"/\\|?*]+')
CSS_URL_RE = re.compile(r"url\(\s*(['\"]?)([^)'\"]+)\1\s*\)", re.IGNORECASE)
CSS_IMPORT_RE = re.compile(
    r"@import\s+(?:url\(\s*)?(['\"])([^'\"]+)\1\s*\)?",
    re.IGNORECASE,
)
ABSOLUTE_URL_RE = re.compile(r"https?://[A-Za-z0-9._~:/?#\[\]@!$&'()*+,;=%-]+", re.IGNORECASE)
ROOT_PATH_RE = re.compile(r"(?P<quote>['\"])(?P<path>/(?!/)[^'\"\s<>]+)(?P=quote)")
EDGE_SKIP_DIRS = {
    "Cache",
    "Code Cache",
    "GPUCache",
    "ShaderCache",
    "GrShaderCache",
    "GraphiteDawnCache",
    "DawnGraphiteCache",
    "DawnWebGPUCache",
    "Crashpad",
    "Service Worker",
    "Session Storage",
    "Safe Browsing",
    "Safe Browsing Network",
}
EDGE_SKIP_FILES = {
    "LOCK",
    "LOCKfile",
    "SingletonCookie",
    "SingletonLock",
    "SingletonSocket",
    "Current Session",
    "Current Tabs",
    "Last Session",
    "Last Tabs",
    "Cookies",
    "Cookies-journal",
}


def log(message: str) -> None:
    print(message, flush=True)


def normalize_url(url: str) -> str:
    parsed = urlsplit(url.strip())
    scheme = parsed.scheme.lower()
    netloc = parsed.netloc.lower()
    path = parsed.path or "/"
    if not path.startswith("/"):
        path = "/" + path
    normalized = urlunsplit((scheme, netloc, path, parsed.query, ""))
    return normalized


def is_supported_url(url: str) -> bool:
    return url.startswith("http://") or url.startswith("https://")


def clean_reference(value: str | None) -> str | None:
    if value is None:
        return None
    candidate = value.strip()
    if not candidate:
        return None
    lowered = candidate.lower()
    if lowered.startswith(("javascript:", "mailto:", "tel:", "data:", "#")):
        return None
    return candidate


def sanitize_segment(segment: str) -> str:
    cleaned = INVALID_SEGMENT_CHARS.sub("_", segment).strip(" .")
    return cleaned or "_"


def append_query_suffix(filename: str, query: str) -> str:
    if not query:
        return filename
    suffix = hashlib.sha1(query.encode("utf-8")).hexdigest()[:10]
    path = PurePosixPath(filename)
    stem = path.stem
    ext = path.suffix
    return f"{stem}__q_{suffix}{ext}"


def guess_extension_from_content_type(content_type: str | None) -> str:
    if not content_type:
        return ".bin"
    base_type = content_type.split(";", 1)[0].strip().lower()
    return mimetypes.guess_extension(base_type) or ".bin"


def content_type_is_html(content_type: str | None) -> bool:
    if not content_type:
        return False
    base_type = content_type.split(";", 1)[0].strip().lower()
    return base_type in {"text/html", "application/xhtml+xml"}


def decode_bytes(data: bytes, content_type: str | None) -> str:
    candidates = []
    if content_type and "charset=" in content_type.lower():
        charset = content_type.split("charset=", 1)[1].split(";", 1)[0].strip()
        if charset:
            candidates.append(charset.strip("\"'"))
    candidates.extend(["utf-8", "latin-1", "cp1252"])
    for encoding in candidates:
        try:
            return data.decode(encoding)
        except (LookupError, UnicodeDecodeError):
            continue
    return data.decode("utf-8", errors="replace")


def rebuild_tag(tag: str, attrs: list[tuple[str, str | None]], closed: bool = False) -> str:
    parts = [f"<{tag}"]
    for name, value in attrs:
        if value is None:
            parts.append(f" {name}")
        else:
            escaped = html.escape(value, quote=True)
            parts.append(f' {name}="{escaped}"')
    parts.append(" />" if closed else ">")
    return "".join(parts)


class DownloadError(Exception):
    pass


def browser_cookie_to_cookiejar(cookie_data: dict[str, object]) -> Cookie:
    domain = str(cookie_data.get("domain") or "")
    if not domain:
        raise ValueError("Cookie sem dominio")

    path = str(cookie_data.get("path") or "/")
    name = str(cookie_data.get("name") or "")
    value = str(cookie_data.get("value") or "")
    secure = bool(cookie_data.get("secure", False))
    expires_raw = cookie_data.get("expirationDate")
    expires = int(float(expires_raw)) if expires_raw is not None else None

    initial_dot = domain.startswith(".")
    host_only = bool(cookie_data.get("hostOnly", False))
    domain_specified = not host_only
    domain_initial_dot = initial_dot and domain_specified

    return Cookie(
        version=0,
        name=name,
        value=value,
        port=None,
        port_specified=False,
        domain=domain,
        domain_specified=domain_specified,
        domain_initial_dot=domain_initial_dot,
        path=path,
        path_specified=True,
        secure=secure,
        expires=expires,
        discard=expires is None,
        comment=None,
        comment_url=None,
        rest={"HttpOnly": cookie_data.get("httpOnly", False)},
        rfc2109=False,
    )


def load_browser_cookie_json(raw_text: str) -> CookieJar:
    try:
        parsed = json.loads(raw_text)
    except json.JSONDecodeError as error:
        raise RuntimeError("JSON de cookies invalido.") from error

    if not isinstance(parsed, list):
        raise RuntimeError("O JSON de cookies deve ser uma lista de objetos.")

    jar = CookieJar()
    total = 0
    for item in parsed:
        if not isinstance(item, dict):
            continue
        name = str(item.get("name") or "")
        if not name:
            continue
        try:
            jar.set_cookie(browser_cookie_to_cookiejar(item))
            total += 1
        except ValueError:
            continue

    if total == 0:
        raise RuntimeError("Nenhum cookie valido foi encontrado no JSON informado.")
    return jar


class UrllibFetcher:
    def __init__(self, user_agent: str, timeout: int) -> None:
        self.user_agent = user_agent
        self.timeout = timeout
        self.opener = build_opener()

    def _request(self, url: str) -> tuple[bytes, str | None]:
        request = Request(
            url,
            headers={
                "User-Agent": self.user_agent,
                "Accept": "*/*",
            },
        )
        try:
            with self.opener.open(request, timeout=self.timeout) as response:
                data = response.read()
                content_type = response.headers.get("Content-Type")
                return data, content_type
        except (HTTPError, URLError, TimeoutError) as error:
            raise DownloadError(str(error)) from error

    def fetch_page(self, url: str) -> tuple[bytes, str | None]:
        return self._request(url)

    def fetch_asset(self, url: str) -> tuple[bytes, str | None]:
        return self._request(url)

    def close(self) -> None:
        return None


class CookieUrllibFetcher(UrllibFetcher):
    def __init__(self, user_agent: str, timeout: int, cookie_jar: CookieJar) -> None:
        self.cookie_jar = cookie_jar
        super().__init__(user_agent=user_agent, timeout=timeout)
        self.opener = build_opener(HTTPCookieProcessor(self.cookie_jar))


class PlaywrightFetcher:
    def __init__(
        self,
        start_url: str,
        login_url: str,
        user_agent: str,
        timeout: int,
        state_file: Path,
        user_data_dir: Path,
        profile_directory: str | None,
        using_existing_profile: bool,
        force_login: bool,
    ) -> None:
        self.start_url = start_url
        self.login_url = login_url
        self.user_agent = user_agent
        self.timeout_ms = max(1, timeout) * 1000
        self.idle_timeout_ms = min(5000, self.timeout_ms)
        self.state_file = state_file
        self.user_data_dir = user_data_dir
        self.profile_directory = profile_directory
        self.using_existing_profile = using_existing_profile
        self.force_login = force_login

        try:
            from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
            from playwright.sync_api import sync_playwright
        except ImportError as error:
            raise RuntimeError(
                "Playwright nao esta instalado. Rode: pip install playwright"
            ) from error

        self._playwright_timeout_error = PlaywrightTimeoutError
        self._sync_playwright = sync_playwright
        self._manager = self._sync_playwright().start()
        self.user_data_dir = self._prepare_user_data_dir()
        launch_kwargs: dict[str, object] = {
            "channel": "msedge",
            "headless": False,
            "user_agent": self.user_agent,
        }
        if self.profile_directory:
            launch_kwargs["args"] = [f"--profile-directory={self.profile_directory}"]
        try:
            self.context = self._manager.chromium.launch_persistent_context(
                str(self.user_data_dir),
                **launch_kwargs,
            )
        except Exception as error:
            with suppress(Exception):
                self._manager.stop()
            raise RuntimeError(
                "Nao foi possivel abrir o Microsoft Edge pelo Playwright. "
                "Se voce estiver usando um perfil existente do Edge, tente novamente com todas as janelas do Edge fechadas. "
                "Se o Edge nao estiver disponivel, tente: playwright install msedge"
            ) from error
        self.context.set_default_timeout(self.timeout_ms)
        if self.context.pages:
            self.page = self.context.pages[0]
        else:
            self.page = self.context.new_page()

        if self.force_login or not self.state_file.exists():
            self.capture_login()
        else:
            log(f"Reutilizando perfil persistente do Edge em: {self.user_data_dir}")
            if self.profile_directory:
                log(f"Perfil selecionado: {self.profile_directory}")
            log(f"Sessao salva em: {self.state_file}")

    def _prepare_user_data_dir(self) -> Path:
        if not self.using_existing_profile:
            self.user_data_dir.mkdir(parents=True, exist_ok=True)
            return self.user_data_dir

        if not self.profile_directory:
            raise RuntimeError("Informe um perfil do Edge, como Default ou 'Profile 1'.")

        source_root = self.user_data_dir
        source_profile_dir = source_root / self.profile_directory
        if not source_profile_dir.exists():
            raise RuntimeError(
                f"O perfil do Edge '{self.profile_directory}' nao foi encontrado em {source_root}"
            )

        clone_root = self.state_file.parent / "edge_existing_profile_clone"
        clone_profile_dir = clone_root / self.profile_directory

        if clone_root.exists():
            shutil.rmtree(clone_root, ignore_errors=True)
        clone_root.mkdir(parents=True, exist_ok=True)

        local_state = source_root / "Local State"
        if local_state.exists():
            shutil.copy2(local_state, clone_root / "Local State")
        first_run = source_root / "First Run"
        if first_run.exists():
            shutil.copy2(first_run, clone_root / "First Run")

        log(f"Clonando perfil existente do Edge '{self.profile_directory}'...")
        try:
            skipped = self._copy_edge_profile(source_profile_dir, clone_profile_dir)
        except Exception as error:
            raise RuntimeError(
                "Nao foi possivel copiar o perfil existente do Edge. "
                "Feche todas as janelas do Edge e tente novamente."
            ) from error
        log(f"Perfil clonado para: {clone_root}")
        if skipped:
            log(f"Arquivos ignorados na clonagem: {skipped}")
        return clone_root

    def _copy_edge_profile(self, source_dir: Path, target_dir: Path) -> int:
        skipped = 0
        for root, dirnames, filenames in os.walk(source_dir):
            root_path = Path(root)
            relative_root = root_path.relative_to(source_dir)

            dirnames[:] = [name for name in dirnames if name not in EDGE_SKIP_DIRS]
            destination_root = target_dir / relative_root
            destination_root.mkdir(parents=True, exist_ok=True)

            for filename in filenames:
                if filename in EDGE_SKIP_FILES:
                    skipped += 1
                    continue

                source_file = root_path / filename
                destination_file = destination_root / filename
                try:
                    shutil.copy2(source_file, destination_file)
                except OSError:
                    skipped += 1
        return skipped

    def capture_login(self) -> None:
        log("Abrindo o Microsoft Edge para login manual...")
        if self.using_existing_profile:
            log("Modo perfil existente ativo. Feche antes as outras janelas do Edge se houver travamento.")
        if self.profile_directory:
            log(f"Usando o perfil do Edge: {self.profile_directory}")
        self.page.goto(self.login_url, wait_until="domcontentloaded", timeout=self.timeout_ms)
        log("Entre com sua conta Google no Edge aberto.")
        log("Esse perfil do Edge fica salvo para lembrar sua conta nas proximas execucoes.")
        log("Quando o site desejado estiver carregado e autenticado, volte ao terminal.")
        input("Pressione Enter para continuar com a copia autenticada...")
        with suppress(self._playwright_timeout_error):
            self.page.goto(self.start_url, wait_until="domcontentloaded", timeout=self.timeout_ms)
            self._wait_for_idle()
        self.state_file.parent.mkdir(parents=True, exist_ok=True)
        self.context.storage_state(path=str(self.state_file), indexed_db=True)
        log(f"Sessao autenticada salva em: {self.state_file}")

    def _wait_for_idle(self) -> None:
        with suppress(self._playwright_timeout_error):
            self.page.wait_for_load_state("networkidle", timeout=self.idle_timeout_ms)

    def fetch_page(self, url: str) -> tuple[bytes, str | None]:
        try:
            response = self.page.goto(url, wait_until="domcontentloaded", timeout=self.timeout_ms)
            self._wait_for_idle()
            content_type = response.headers.get("content-type") if response else None
            if content_type and not content_type_is_html(content_type):
                return self.fetch_asset(url)
            html_text = self.page.content()
            return html_text.encode("utf-8"), content_type or "text/html; charset=utf-8"
        except self._playwright_timeout_error as error:
            raise DownloadError(f"Timeout ao abrir {url}") from error
        except Exception as error:
            raise DownloadError(str(error)) from error

    def fetch_asset(self, url: str) -> tuple[bytes, str | None]:
        response = None
        try:
            response = self.context.request.get(
                url,
                timeout=self.timeout_ms,
                fail_on_status_code=False,
                headers={"User-Agent": self.user_agent},
            )
            status = response.status
            content_type = response.headers.get("content-type")
            data = response.body()
            if status >= 400:
                raise DownloadError(f"HTTP {status} ao baixar {url}")
            return data, content_type
        except self._playwright_timeout_error as error:
            raise DownloadError(f"Timeout ao baixar recurso {url}") from error
        except DownloadError:
            raise
        except Exception as error:
            raise DownloadError(str(error)) from error
        finally:
            if response is not None:
                with suppress(Exception):
                    response.dispose()

    def close(self) -> None:
        with suppress(Exception):
            self.context.close()
        with suppress(Exception):
            self._manager.stop()


class HtmlRewriter(HTMLParser):
    def __init__(self, mirror: "SiteMirror", page_url: str, page_local_path: Path) -> None:
        super().__init__(convert_charrefs=False)
        self.mirror = mirror
        self.page_url = page_url
        self.page_local_path = page_local_path
        self.output: list[str] = []
        self.found_pages: set[str] = set()
        self.found_assets: set[str] = set()

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self.output.append(self._render_tag(tag, attrs, closed=False))

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self.output.append(self._render_tag(tag, attrs, closed=True))

    def handle_endtag(self, tag: str) -> None:
        self.output.append(f"</{tag}>")

    def handle_data(self, data: str) -> None:
        self.output.append(data)

    def handle_comment(self, data: str) -> None:
        self.output.append(f"<!--{data}-->")

    def handle_decl(self, decl: str) -> None:
        self.output.append(f"<!{decl}>")

    def handle_pi(self, data: str) -> None:
        self.output.append(f"<?{data}>")

    def handle_entityref(self, name: str) -> None:
        self.output.append(f"&{name};")

    def handle_charref(self, name: str) -> None:
        self.output.append(f"&#{name};")

    def get_html(self) -> str:
        return "".join(self.output)

    def _render_tag(self, tag: str, attrs: list[tuple[str, str | None]], closed: bool) -> str:
        attr_map = {name.lower(): value for name, value in attrs}
        new_attrs: list[tuple[str, str | None]] = []
        for name, value in attrs:
            lowered_name = name.lower()
            rewritten = value
            if (tag.lower(), lowered_name) in URL_ATTRS:
                rewritten = self._rewrite_url(tag.lower(), lowered_name, attr_map, value)
            elif lowered_name == "srcset" and value:
                rewritten = self._rewrite_srcset(value)
            elif lowered_name == "style" and value:
                rewritten = self.mirror.rewrite_css_text(
                    value,
                    base_url=self.page_url,
                    current_local_path=self.page_local_path,
                )
            elif tag.lower() == "meta" and lowered_name == "content" and value:
                http_equiv = (attr_map.get("http-equiv") or "").lower()
                if http_equiv == "refresh":
                    rewritten = self._rewrite_meta_refresh(value)
            new_attrs.append((name, rewritten))
        return rebuild_tag(tag, new_attrs, closed=closed)

    def _rewrite_url(
        self,
        tag: str,
        attr_name: str,
        attr_map: dict[str, str | None],
        raw_value: str | None,
    ) -> str | None:
        reference = clean_reference(raw_value)
        if reference is None:
            return raw_value

        absolute_url = normalize_url(urljoin(self.page_url, reference))
        if not self.mirror.should_download(absolute_url):
            return raw_value

        if self._is_page_reference(tag, attr_name, attr_map, absolute_url):
            self.found_pages.add(absolute_url)
            local_target = self.mirror.local_html_path(absolute_url)
        else:
            self.found_assets.add(absolute_url)
            local_target = self.mirror.local_asset_path(absolute_url, content_type=None)

        relative = self.mirror.relative_reference(self.page_local_path, local_target)
        return relative

    def _is_page_reference(
        self,
        tag: str,
        attr_name: str,
        attr_map: dict[str, str | None],
        absolute_url: str,
    ) -> bool:
        if tag in {"a", "form", "iframe"}:
            return True
        if tag == "meta" and attr_name == "content":
            return True
        if tag == "link":
            rel = (attr_map.get("rel") or "").lower()
            asset_rels = ("stylesheet", "icon", "preload", "prefetch", "manifest", "mask-icon")
            if any(item in rel for item in asset_rels):
                return False
            return self.mirror.is_html_candidate(absolute_url)
        if attr_name in {"poster"}:
            return False
        return False

    def _rewrite_srcset(self, srcset_value: str) -> str:
        parts = []
        for item in srcset_value.split(","):
            candidate = item.strip()
            if not candidate:
                continue
            pieces = candidate.split()
            source = pieces[0]
            descriptor = " ".join(pieces[1:])
            rewritten = self._rewrite_url("source", "srcset", {}, source) or source
            if descriptor:
                parts.append(f"{rewritten} {descriptor}")
            else:
                parts.append(rewritten)
        return ", ".join(parts)

    def _rewrite_meta_refresh(self, content: str) -> str:
        match = re.search(r"(url\s*=\s*)(.+)$", content, re.IGNORECASE)
        if not match:
            return content
        prefix = match.group(1)
        raw_url = match.group(2).strip(" '\"")
        rewritten = self._rewrite_url("meta", "content", {"http-equiv": "refresh"}, raw_url) or raw_url
        return content[: match.start(2)] + rewritten


class SiteMirror:
    def __init__(
        self,
        start_url: str,
        destination: Path,
        max_pages: int,
        delay: float,
        timeout: int,
        allow_subdomains: bool,
        fetcher: UrllibFetcher | PlaywrightFetcher,
    ) -> None:
        self.start_url = normalize_url(start_url)
        self.destination = destination
        self.max_pages = max_pages
        self.delay = delay
        self.timeout = timeout
        self.allow_subdomains = allow_subdomains
        self.fetcher = fetcher

        start_parts = urlsplit(self.start_url)
        self.base_scheme = start_parts.scheme
        self.base_host = start_parts.hostname or ""
        self.base_netloc = start_parts.netloc.lower()
        self.root_output = self.destination / sanitize_segment(self.base_netloc)

        self.visited_pages: set[str] = set()
        self.downloaded_assets: set[str] = set()
        self.queued_pages: set[str] = set()
        self.pending_pages: set[str] = set()

    def should_download(self, url: str) -> bool:
        if not is_supported_url(url):
            return False
        parts = urlsplit(url)
        if parts.scheme.lower() != self.base_scheme:
            return False
        host = parts.hostname or ""
        if self.allow_subdomains:
            return host == self.base_host or host.endswith("." + self.base_host)
        return parts.netloc.lower() == self.base_netloc

    def is_html_candidate(self, url: str, content_type: str | None = None) -> bool:
        if content_type:
            if content_type_is_html(content_type):
                return True
            base_type = content_type.split(";", 1)[0].strip().lower()
            if base_type and not base_type.startswith("text/html"):
                return False

        path = urlsplit(url).path
        suffix = PurePosixPath(path).suffix.lower()
        return suffix in HTML_EXTENSIONS

    def fetch_page(self, url: str) -> tuple[bytes, str | None]:
        return self.fetcher.fetch_page(url)

    def fetch_asset(self, url: str) -> tuple[bytes, str | None]:
        return self.fetcher.fetch_asset(url)

    def local_html_path(self, url: str) -> Path:
        return self._build_local_path(url, is_html=True, content_type="text/html")

    def local_asset_path(self, url: str, content_type: str | None) -> Path:
        return self._build_local_path(url, is_html=False, content_type=content_type)

    def relative_reference(self, current_path: Path, target_path: Path) -> str:
        relative = posixpath.relpath(
            target_path.as_posix(),
            start=current_path.parent.as_posix(),
        )
        return relative.replace("\\", "/")

    def _build_local_path(self, url: str, is_html: bool, content_type: str | None) -> Path:
        parts = urlsplit(url)
        url_path = parts.path or "/"
        segments = [sanitize_segment(seg) for seg in PurePosixPath(url_path).parts if seg not in {"/", ""}]

        if is_html:
            if url_path.endswith("/") or not PurePosixPath(url_path).suffix:
                relative_dir = Path(*segments) if segments else Path()
                filename = "index.html"
            else:
                relative_dir = Path(*segments[:-1]) if len(segments) > 1 else Path()
                filename = segments[-1]
            filename = append_query_suffix(filename, parts.query)
            return self.root_output / relative_dir / filename

        if url_path.endswith("/") or not segments:
            relative_dir = Path(*segments) if segments else Path()
            ext = guess_extension_from_content_type(content_type)
            filename = append_query_suffix(f"index{ext}", parts.query)
        else:
            relative_dir = Path(*segments[:-1]) if len(segments) > 1 else Path()
            filename = append_query_suffix(segments[-1], parts.query)
            if "." not in Path(filename).name:
                filename += guess_extension_from_content_type(content_type)

        return self.root_output / relative_dir / filename

    def save_file(self, local_path: Path, data: bytes) -> None:
        local_path.parent.mkdir(parents=True, exist_ok=True)
        local_path.write_bytes(data)

    def enqueue_page(self, queue: deque[str], url: str) -> None:
        if url in self.visited_pages or url in self.queued_pages:
            return
        self.queued_pages.add(url)
        queue.append(url)

    def enqueue_pending_pages(self, queue: deque[str]) -> None:
        for url in sorted(self.pending_pages):
            self.enqueue_page(queue, url)
        self.pending_pages.clear()

    def is_text_discovery_candidate(self, local_path: Path, content_type: str | None) -> bool:
        base_type = (content_type or "").split(";", 1)[0].strip().lower()
        if base_type in {
            "application/javascript",
            "text/javascript",
            "application/json",
            "application/manifest+json",
            "text/plain",
            "text/html",
        }:
            return True
        return local_path.suffix.lower() in {".js", ".json", ".webmanifest", ".html", ".txt"}

    def classify_discovered_url(self, url: str) -> str:
        parts = urlsplit(url)
        path = parts.path or "/"
        lowered_path = path.lower()
        suffix = PurePosixPath(path).suffix.lower()

        if lowered_path.startswith(("/_framework/", "/_content/", "/assets/", "/icons/")):
            return "asset"
        if lowered_path.startswith("/api") or "/api/" in lowered_path:
            return "asset"
        if lowered_path.startswith("/_blazor") or "signalr" in lowered_path or lowered_path.endswith("/hub"):
            return "asset"
        if suffix and suffix not in HTML_EXTENSIONS:
            return "asset"
        return "page"

    def discover_text_references(self, text: str, base_url: str) -> tuple[set[str], set[str]]:
        pages: set[str] = set()
        assets: set[str] = set()

        def add_candidate(candidate: str) -> None:
            reference = clean_reference(candidate)
            if reference is None:
                return
            absolute = normalize_url(urljoin(base_url, reference))
            if not self.should_download(absolute):
                return
            if self.classify_discovered_url(absolute) == "page":
                pages.add(absolute)
            else:
                assets.add(absolute)

        for match in ABSOLUTE_URL_RE.finditer(text):
            add_candidate(match.group(0))
        for match in ROOT_PATH_RE.finditer(text):
            add_candidate(match.group("path"))
        return pages, assets

    def process_text_discoveries(self, text: str, base_url: str) -> None:
        pages, assets = self.discover_text_references(text, base_url=base_url)
        for page_url in pages:
            if page_url not in self.visited_pages and page_url not in self.queued_pages:
                self.pending_pages.add(page_url)
        for asset_url in assets:
            self.download_asset(asset_url)

    def rewrite_css_text(self, css_text: str, base_url: str, current_local_path: Path) -> str:
        def replace_url(match: re.Match[str]) -> str:
            original = match.group(0)
            raw_target = match.group(2).strip()
            reference = clean_reference(raw_target)
            if reference is None:
                return original

            absolute = normalize_url(urljoin(base_url, reference))
            if not self.should_download(absolute):
                return original

            local_target = self.local_asset_path(absolute, content_type=None)
            self.download_asset(absolute)
            relative = self.relative_reference(current_local_path, local_target)
            quote = match.group(1) or ""
            return f"url({quote}{relative}{quote})"

        def replace_import(match: re.Match[str]) -> str:
            raw_target = match.group(2).strip()
            reference = clean_reference(raw_target)
            if reference is None:
                return match.group(0)

            absolute = normalize_url(urljoin(base_url, reference))
            if not self.should_download(absolute):
                return match.group(0)

            local_target = self.local_asset_path(absolute, content_type="text/css")
            self.download_asset(absolute)
            relative = self.relative_reference(current_local_path, local_target)
            quote = match.group(1)
            return f"@import {quote}{relative}{quote}"

        css_text = CSS_URL_RE.sub(replace_url, css_text)
        css_text = CSS_IMPORT_RE.sub(replace_import, css_text)
        return css_text

    def download_asset(self, url: str) -> None:
        if url in self.downloaded_assets or url in self.visited_pages:
            return
        if not self.should_download(url):
            return

        try:
            data, content_type = self.fetch_asset(url)
        except DownloadError as error:
            log(f"[aviso] Falha ao baixar recurso {url}: {error}")
            return

        if self.is_html_candidate(url, content_type):
            return

        local_path = self.local_asset_path(url, content_type)
        output = data

        base_type = (content_type or "").split(";", 1)[0].strip().lower()
        if base_type == "text/css" or local_path.suffix.lower() == ".css":
            css_text = decode_bytes(data, content_type)
            css_text = self.rewrite_css_text(
                css_text,
                base_url=url,
                current_local_path=local_path,
            )
            output = css_text.encode("utf-8")
        elif self.is_text_discovery_candidate(local_path, content_type):
            self.process_text_discoveries(
                decode_bytes(data, content_type),
                base_url=url,
            )

        self.save_file(local_path, output)
        self.downloaded_assets.add(url)
        log(f"[asset] {url} -> {local_path}")
        time.sleep(self.delay)

    def mirror(self) -> Path:
        queue: deque[str] = deque()
        self.enqueue_page(queue, self.start_url)

        while queue and len(self.visited_pages) < self.max_pages:
            url = queue.popleft()
            self.queued_pages.discard(url)

            if url in self.visited_pages:
                continue

            try:
                data, content_type = self.fetch_page(url)
            except DownloadError as error:
                log(f"[aviso] Falha ao baixar pagina {url}: {error}")
                continue

            if not self.is_html_candidate(url, content_type):
                local_path = self.local_asset_path(url, content_type)
                if self.is_text_discovery_candidate(local_path, content_type):
                    self.process_text_discoveries(
                        decode_bytes(data, content_type),
                        base_url=url,
                    )
                self.save_file(local_path, data)
                self.downloaded_assets.add(url)
                log(f"[asset] {url} -> {local_path}")
                self.enqueue_pending_pages(queue)
                continue

            html_text = decode_bytes(data, content_type)
            local_path = self.local_html_path(url)
            parser = HtmlRewriter(self, url, local_path)
            parser.feed(html_text)
            parser.close()

            rewritten_html = parser.get_html()
            self.save_file(local_path, rewritten_html.encode("utf-8"))
            self.visited_pages.add(url)
            log(f"[page]  {url} -> {local_path}")

            self.process_text_discoveries(html_text, base_url=url)

            for asset_url in sorted(parser.found_assets):
                self.download_asset(asset_url)

            for page_url in sorted(parser.found_pages):
                if self.should_download(page_url):
                    self.enqueue_page(queue, page_url)

            self.enqueue_pending_pages(queue)

            time.sleep(self.delay)

        if queue:
            log(f"[aviso] Limite de {self.max_pages} paginas atingido.")

        return self.local_html_path(self.start_url)


def default_state_file(start_url: str) -> Path:
    host = sanitize_segment(urlsplit(start_url).netloc.lower())
    return Path.home() / ".site_mirror_auth" / host / "edge_state.json"


def default_user_data_dir(start_url: str) -> Path:
    host = sanitize_segment(urlsplit(start_url).netloc.lower())
    return Path.home() / ".site_mirror_auth" / host / "edge_profile"


def installed_edge_user_data_dir() -> Path:
    local_appdata = Path.home() / "AppData" / "Local"
    return local_appdata / "Microsoft" / "Edge" / "User Data"


def load_cookie_jar_from_args(args: argparse.Namespace) -> CookieJar | None:
    if not args.cookies_json and not args.cookies_stdin:
        return None

    if args.cookies_json and args.cookies_stdin:
        raise RuntimeError("Use apenas uma opcao de cookies: --cookies-json ou --cookies-stdin.")

    if args.cookies_json:
        cookie_path = Path(args.cookies_json).expanduser().resolve()
        raw_text = cookie_path.read_text(encoding="utf-8")
    else:
        raw_text = sys.stdin.read()

    return load_browser_cookie_json(raw_text)


def parse_args(argv: Iterable[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Baixa o codigo-fonte publico de um site para uma pasta local.",
    )
    parser.add_argument("url", help="URL inicial do site, ex.: https://exemplo.com")
    parser.add_argument(
        "-d",
        "--destino",
        default="site_espelhado",
        help="Pasta onde a copia sera salva (padrao: site_espelhado)",
    )
    parser.add_argument(
        "--max-paginas",
        type=int,
        default=200,
        help="Numero maximo de paginas HTML internas para baixar (padrao: 200)",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=0.25,
        help="Pausa entre requisicoes em segundos (padrao: 0.25)",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=20,
        help="Timeout por requisicao em segundos (padrao: 20)",
    )
    parser.add_argument(
        "--permitir-subdominios",
        action="store_true",
        help="Inclui subdominios do mesmo host base",
    )
    parser.add_argument(
        "--user-agent",
        default="SiteMirrorOffline/1.0 (+uso interno autorizado)",
        help="User-Agent enviado nas requisicoes",
    )
    parser.add_argument(
        "--login-edge",
        action="store_true",
        help="Abre o Microsoft Edge com perfil persistente para login manual e reutiliza a sessao autenticada",
    )
    parser.add_argument(
        "--login-url",
        help="URL inicial do login manual no Edge; se omitida, usa a propria URL informada",
    )
    parser.add_argument(
        "--estado-sessao",
        help="Arquivo JSON para salvar os cookies/sessao autenticada do Edge",
    )
    parser.add_argument(
        "--edge-user-data-dir",
        help="Pasta do perfil persistente do Edge; ela guarda a conta Google reconhecida",
    )
    parser.add_argument(
        "--usar-perfil-edge-existente",
        action="store_true",
        help="Clona o perfil ja existente do Edge do Windows e abre essa copia para reaproveitar o login",
    )
    parser.add_argument(
        "--edge-profile-directory",
        default="Default",
        help="Nome do perfil dentro do Edge User Data, ex.: Default ou 'Profile 1'",
    )
    parser.add_argument(
        "--forcar-login",
        action="store_true",
        help="Ignora a sessao salva e pede novo login manual no Edge",
    )
    parser.add_argument(
        "--cookies-json",
        help="Arquivo JSON com cookies exportados do navegador",
    )
    parser.add_argument(
        "--cookies-stdin",
        action="store_true",
        help="Le o JSON de cookies pela entrada padrao",
    )
    return parser.parse_args(list(argv))


def main(argv: Iterable[str]) -> int:
    args = parse_args(argv)
    start_url = normalize_url(args.url)
    if not is_supported_url(start_url):
        print("Informe uma URL http:// ou https://", file=sys.stderr)
        return 1

    destination = Path(args.destino).resolve()
    state_file = (
        Path(args.estado_sessao).expanduser().resolve()
        if args.estado_sessao
        else default_state_file(start_url)
    )
    if args.edge_user_data_dir:
        user_data_dir = Path(args.edge_user_data_dir).expanduser().resolve()
    elif args.usar_perfil_edge_existente:
        user_data_dir = installed_edge_user_data_dir()
    else:
        user_data_dir = default_user_data_dir(start_url)

    if args.usar_perfil_edge_existente and not user_data_dir.exists():
        print(
            f"Perfil padrao do Edge nao encontrado em: {user_data_dir}",
            file=sys.stderr,
        )
        return 1

    try:
        cookie_jar = load_cookie_jar_from_args(args)
        if args.login_edge and cookie_jar is not None:
            raise RuntimeError("Nao combine --login-edge com cookies; use apenas um metodo de autenticacao.")

        if cookie_jar is not None:
            fetcher = CookieUrllibFetcher(
                user_agent=args.user_agent,
                timeout=max(1, args.timeout),
                cookie_jar=cookie_jar,
            )
        elif args.login_edge:
            fetcher = PlaywrightFetcher(
                start_url=start_url,
                login_url=normalize_url(args.login_url) if args.login_url else start_url,
                user_agent=args.user_agent,
                timeout=max(1, args.timeout),
                state_file=state_file,
                user_data_dir=user_data_dir,
                profile_directory=args.edge_profile_directory if args.login_edge else None,
                using_existing_profile=args.usar_perfil_edge_existente,
                force_login=args.forcar_login,
            )
        else:
            fetcher = UrllibFetcher(
                user_agent=args.user_agent,
                timeout=max(1, args.timeout),
            )
    except RuntimeError as error:
        print(str(error), file=sys.stderr)
        return 1

    mirror = SiteMirror(
        start_url=start_url,
        destination=destination,
        max_pages=max(1, args.max_paginas),
        delay=max(0.0, args.delay),
        timeout=max(1, args.timeout),
        allow_subdomains=args.permitir_subdominios,
        fetcher=fetcher,
    )

    try:
        if cookie_jar is not None:
            log("Iniciando espelhamento com cookies autenticados...")
        elif args.login_edge:
            log("Iniciando espelhamento com sessao autenticada do Edge...")
        else:
            log("Iniciando espelhamento do conteudo publico do site...")
        first_page = mirror.mirror()
        log("")
        log("Concluido.")
        log(f"Pagina inicial local: {first_page}")
        return 0
    finally:
        fetcher.close()


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
