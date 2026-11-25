"""Leitura da árvore de documentos na janela Visualização de Autos do SAJ."""

from __future__ import annotations

from dataclasses import dataclass
from hashlib import blake2b
from io import BytesIO
import re
import subprocess
import sys
import time
import unicodedata
from typing import Iterator

from PIL import Image, ImageGrab, ImageOps


_AR = re.compile(r"^(aviso(?:s)? de recebimento\b|ar(?: digital)?(?=$|[(:,-]))")
_MANDADO = re.compile(r"^(mandados?\b|devolucao d[eo] mandado\b)")
_EDITAL = re.compile(r"^(edital\b|editais\b|publicacao de edital\b)")
_CERTIDAO = re.compile(r"^(certid(?:ao|oes)|certificac(?:ao|oes))\b")
_PAGE = re.compile(r"^pagina\s*:?\s*(\d+)\b")
_CAPTION_PAGE = re.compile(r"\bpagina\s*[:.,;]?\s*(\d+)\b")
_CAPTION_CATEGORY = re.compile(r"\bcategoria\s*[:.,;]?\s*([^,]+)")
_CAPTURE_REGION = (0.19, 0.35, 0.99, 0.94)
_CAPTION_REGION = (0.198, 0.29, 0.80, 0.351)
_DOCUMENT_REGION = (0.19, 0.264, 0.982, 0.94)


class SajError(RuntimeError):
    """A consulta ou captura do SAJ não pôde ser confirmada."""


@dataclass(frozen=True)
class DocumentCapture:
    page: str
    image: Image.Image | None
    document_index: int = 1
    error: str | None = None
    kind: str = "AR"
    complete: bool = True
    text: str | None = None


ReceiptCapture = DocumentCapture


def _normalize(value: str) -> str:
    ascii_text = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode()
    return " ".join(ascii_text.lower().split())


def _digits(value: str) -> str:
    return re.sub(r"\D", "", value)


def _document_kind(value: str) -> str | None:
    label = _normalize(value)
    for kind, pattern in (
        ("AR", _AR), ("MANDADO", _MANDADO), ("EDITAL", _EDITAL), ("CERTIDAO", _CERTIDAO),
    ):
        if pattern.match(label):
            return kind
    return None


def _process_matches(text: str, digits: str) -> bool:
    return bool(re.fullmatch(r"[\s\d.-]+", text)) and _digits(text) == digits


def _process_parts(value: str) -> dict[str, str]:
    digits = _digits(value)
    if len(digits) != 20 or not _process_matches(value, digits):
        raise SajError("O número do processo precisa ter 20 dígitos no padrão CNJ.")
    prefix = f"{digits[:7]}-{digits[7:9]}.{digits[9:13]}"
    suffix = f"{digits[13]}.{digits[14:16]}.{digits[16:]}"
    return {"cnj": f"{prefix}.{suffix}", "prefix": prefix, "suffix": suffix, "digits": digits}


def _relative_region(config: dict, name: str, default: tuple) -> tuple:
    try:
        left, top, right, bottom = map(float, config.get(name, default))
    except (TypeError, ValueError) as exc:
        raise SajError(f"{name} deve conter quatro números.") from exc
    if not (0 <= left < right <= 1 and 0 <= top < bottom <= 1):
        raise SajError(f"{name} deve delimitar uma região dentro da janela de 0 a 1.")
    return left, top, right, bottom


class SajClient:
    def __init__(self, config: dict):
        self.config = config
        backend = config.get("backend", "auto")
        if backend not in ("auto", "win32", "uia"):
            raise SajError("O backend do SAJ deve ser auto, win32 ou uia.")
        if backend == "auto":
            self.backends = ("uia", "win32") if config.get("document_text_control") else ("win32", "uia")
        else:
            self.backends = (backend,)
        self.capture_region = _relative_region(config, "capture_region", _CAPTURE_REGION)
        self.caption_region = _relative_region(config, "caption_region", _CAPTION_REGION)
        self.document_region = _relative_region(config, "document_region", _DOCUMENT_REGION)
        self.timeout = float(config.get("confirm_timeout_seconds", 12))
        self.settle = float(config.get("settle_seconds", 0.7))
        if not 0 < self.settle < self.timeout:
            raise SajError("O tempo de confirmação deve ser maior que o tempo de estabilização.")
        self._desktop = None
        self._backend = None
        self._window = None
        self._tree_control = None
        self._process = None

    def _windows(self):
        if sys.platform != "win32":
            raise SajError("A consulta automática ao SAJ precisa da interface aberta no Windows.")
        try:
            from pywinauto import Desktop
        except ImportError as exc:
            raise SajError("Instale as dependências para acessar a janela do SAJ.") from exc

        for backend in self.backends:
            desktop = Desktop(backend=backend)
            windows = desktop.windows(title_re=self.config.get("window_title", "^Visualização de Autos$"))
            if len(windows) > 1:
                raise SajError("Há mais de uma janela de Visualização de Autos. Deixe aberta a janela da consulta.")
            if windows:
                yield backend, desktop, windows[0]

    def _select(self, selector: dict):
        return self._desktop.window(handle=self._window.handle).child_window(**selector).wrapper_object()

    def _tree(self):
        if self._tree_control is not None:
            return self._tree_control
        selector = self.config.get("tree")
        if selector:
            tree = self._select(selector)
            if not callable(getattr(tree, "roots", None)):
                raise SajError("O seletor tree precisa apontar para a árvore de documentos.")
            return tree
        candidates = [
            control for control in self._window.descendants()
            if callable(getattr(control, "roots", None)) and control.is_visible()
        ]
        process_trees = [
            tree for tree in candidates
            if any(
                _process_matches(text, _digits(text)) and len(_digits(text)) == 20
                for text in (self._item_text(node) for node in tree.roots())
            )
        ]
        if len(process_trees) == 1:
            return process_trees[0]
        raise SajError("A árvore de documentos do processo não pôde ser localizada. Confira a aba Documentos.")

    def _connect(self) -> None:
        if self._window is not None:
            return
        errors = []
        for backend, desktop, window in self._windows():
            self._backend, self._desktop, self._window = backend, desktop, window
            try:
                self._tree_control = self._tree()
            except Exception as exc:
                errors.append(f"{backend}: {exc}")
                self._window = None
                continue
            return
        context = " | ".join(errors) if errors else "Abra a Visualização de Autos em um processo."
        raise SajError(context)

    def ready(self) -> None:
        """Confirma que a janela e uma árvore de processo estão acessíveis."""
        try:
            self._connect()
            current, _ = self._current_root()
            self._wait_for_case(current)
        except SajError:
            raise
        except Exception as exc:
            raise SajError(f"A verificação inicial da janela do SAJ falhou. {exc}") from exc

    def inspect(self) -> str:
        """Lista seletores e posições locais para calibrar a interface."""
        lines = []
        for backend, _, window in self._windows():
            rect = window.rectangle()
            lines.append(f"{backend} | janela {window.window_text()!r} | {rect.width()} × {rect.height()} px")
            for control in window.descendants():
                if not control.is_visible():
                    continue
                info = control.element_info
                bounds = control.rectangle()
                coords = (
                    (bounds.left - rect.left) / rect.width(),
                    (bounds.top - rect.top) / rect.height(),
                    (bounds.right - rect.left) / rect.width(),
                    (bounds.bottom - rect.top) / rect.height(),
                )
                selectors = (
                    f"classe={info.class_name!r}, tipo={getattr(info, 'control_type', '')!r}, "
                    f"auto_id={getattr(info, 'automation_id', '')!r}, "
                    f"control_id={getattr(info, 'control_id', None)!r}"
                )
                label = self._control_text(control).replace("\n", " ")[:130]
                lines.append(f"  {selectors} | {tuple(round(n, 3) for n in coords)} | {label!r}")
        if not lines:
            raise SajError("Abra a janela Visualização de Autos antes de executar a inspeção.")
        return "\n".join(lines)

    @staticmethod
    def _control_text(control) -> str:
        if callable(getattr(control, "get_value", None)):
            try:
                value = control.get_value()
            except Exception:
                value = None
            if isinstance(value, str):
                return value
        return control.window_text()

    @staticmethod
    def _item_text(node) -> str:
        return node.text() if callable(getattr(node, "text", None)) else node.window_text()

    def _current_root(self):
        matches = []
        for node in self._tree().roots():
            label = self._item_text(node)
            digits = _digits(label)
            if len(digits) == 20 and _process_matches(label, digits):
                matches.append((digits, node))
        if len(matches) != 1:
            raise SajError("O número do processo exibido na árvore não pôde ser confirmado.")
        return matches[0]

    def _children(self, node) -> list:
        if self._backend == "win32":
            children = node.children()
            if not children and node.item().cChildren:
                node.expand()
                children = node.children()
                if not children:
                    raise SajError("A árvore do SAJ indica documentos que o controle Win32 não expôs.")
            return children
        from pywinauto.uia_defines import NoPatternInterfaceError

        try:
            collapsed = node.is_collapsed()
        except NoPatternInterfaceError:
            collapsed = False
        if collapsed:
            node.expand()
        children = node.children(control_type="TreeItem")
        if collapsed and not children:
            raise SajError("A árvore do SAJ indica documentos que o controle UIA não expôs.")
        return children

    def _fields(self, current: str, requested: dict) -> list:
        override = self.config.get("process_fields")
        if override:
            return [(self._select(field["selector"]), field["format"].format_map(requested)) for field in override]

        controls = [
            (control, self._control_text(control))
            for control in self._window.descendants()
            if callable(getattr(control, "set_edit_text", None)) and control.is_visible()
        ]
        current_parts = _process_parts(current)
        for keys in (("cnj",), ("prefix", "suffix")):
            matching_fields = [
                [
                    control for control, label in controls
                    if _process_matches(label, _digits(current_parts[key]))
                ]
                for key in keys
            ]
            if all(len(matches) == 1 for matches in matching_fields):
                return [(matches[0], requested[key]) for matches, key in zip(matching_fields, keys)]
        raise SajError("Os campos do número do processo não foram reconhecidos. Calibre process_fields.")

    def _wait_for_case(self, digits: str) -> None:
        deadline = time.monotonic() + self.timeout
        seen = None
        stable_since = 0.0
        last_error = None
        while time.monotonic() < deadline:
            try:
                current, root = self._current_root()
                children = self._children(root) if current == digits else []
                if children:
                    state = len(children), self._item_text(children[0]), self._item_text(children[-1])
                    now = time.monotonic()
                    if seen == state and now - stable_since >= self.settle:
                        return
                    if seen != state:
                        seen, stable_since = state, now
                else:
                    seen = None
            except Exception as exc:
                last_error = exc
                self._tree_control = None
                seen = None
            time.sleep(0.2)
        raise SajError("O SAJ não confirmou o número solicitado e a lista de documentos carregada.") from last_error

    def open_case(self, process_number: str) -> None:
        requested = _process_parts(process_number)
        self._process = None
        self._connect()
        try:
            self._window.set_focus()
            current, _ = self._current_root()
            if current != requested["digits"]:
                fields = self._fields(current, requested)
                submit_from = None
                for field, value in fields:
                    if _process_matches(self._control_text(field), _digits(value)):
                        continue
                    if not field.is_enabled():
                        raise SajError("Um campo necessário para abrir o processo está desabilitado.")
                    field.set_focus()
                    field.set_edit_text(value)
                    if not _process_matches(self._control_text(field), _digits(value)):
                        field.type_keys("^a{BACKSPACE}" + value, pause=0.03)
                    if not _process_matches(self._control_text(field), _digits(value)):
                        raise SajError("O número digitado no campo do SAJ não pôde ser confirmado.")
                    submit_from = field

                submit = self.config.get("submit_control")
                if submit:
                    self._select(submit).click_input()
                else:
                    if submit_from is None:
                        submit_from = next((field for field, _ in fields if field.is_enabled()), None)
                    if submit_from is None:
                        raise SajError("Nenhum campo disponível para confirmar a consulta no SAJ.")
                    submit_from.type_keys("{ENTER}")
            self._wait_for_case(requested["digits"])
        except SajError:
            raise
        except Exception as exc:
            raise SajError(f"A consulta do processo falhou na interface do SAJ. {exc}") from exc
        self._process = requested["digits"]

    def _check_foreground(self) -> None:
        from win32gui import GetForegroundWindow

        if GetForegroundWindow() != self._window.handle:
            raise SajError("A Visualização de Autos precisa ficar em primeiro plano e sem janelas sobrepostas.")

    @staticmethod
    def _crop(image: Image.Image, region: tuple) -> Image.Image:
        left, top, right, bottom = region
        return image.crop((
            round(image.width * left),
            round(image.height * top),
            round(image.width * right),
            round(image.height * bottom),
        ))

    def _snapshot(self) -> Image.Image:
        self._check_foreground()
        rect = self._window.rectangle()
        bounds = rect.left, rect.top, rect.right, rect.bottom
        snapshot = ImageGrab.grab(bbox=bounds, all_screens=True).convert("RGB")
        self._check_foreground()
        return snapshot

    def snapshot_window(self) -> Image.Image:
        if self._window is not None:
            self._window.set_focus()
            return self._snapshot()
        for _, _, window in self._windows():
            window.set_focus()
            self._window = window
            try:
                return self._snapshot()
            finally:
                self._window = None
        raise SajError("Abra a janela Visualização de Autos para capturar a imagem de calibração.")

    def _caption(self, image: Image.Image) -> str:
        selector = self.config.get("caption_control")
        if selector:
            return _normalize(self._control_text(self._select(selector)))
        gray = ImageOps.autocontrast(ImageOps.grayscale(image))
        enlarged = gray.resize((gray.width * 2, gray.height * 2), Image.Resampling.LANCZOS)
        with BytesIO() as buffer:
            enlarged.save(buffer, format="PNG")
            try:
                result = subprocess.run(
                    [
                        self.config.get("tesseract_cmd", "tesseract"), "stdin", "stdout",
                        "-l", self.config.get("tesseract_lang", "por"), "--psm", "7", "--dpi", "300",
                    ],
                    input=buffer.getvalue(),
                    capture_output=True,
                    timeout=3,
                    check=True,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
            except subprocess.CalledProcessError as exc:
                detail = exc.stderr.decode("utf-8", errors="replace").strip()
                raise SajError(f"Falha no OCR do cabeçalho do SAJ. {detail}") from exc
            except (OSError, subprocess.TimeoutExpired) as exc:
                raise SajError("Falha ao executar o Tesseract para conferir o cabeçalho do SAJ.") from exc
        return _normalize(result.stdout.decode("utf-8", errors="replace"))

    def _caption_confirms(self, caption: str, page: str, kind: str) -> bool:
        number = _CAPTION_PAGE.search(caption)
        if number:
            matched = int(number.group(1)) == int(page)
        else:
            matched = caption.isdigit() and int(caption) == int(page)
        category = _CAPTION_CATEGORY.search(caption)
        category_kind = _document_kind(category.group(1)) if category else None
        related = kind in ("MANDADO", "CERTIDAO") and category_kind in ("MANDADO", "CERTIDAO")
        return matched and (
            not self.config.get("require_category", True)
            or category_kind == kind
            or related
        )

    def _capture_page(self, node, page: str, kind: str) -> Image.Image:
        region = self.capture_region if kind == "AR" else self.document_region
        previous = self._snapshot()
        old_header = self._crop(previous, self.caption_region)
        uses_control = bool(self.config.get("caption_control"))
        old_caption = self._caption(old_header) if uses_control else None
        if self._backend == "uia":
            from pywinauto.uia_defines import NoPatternInterfaceError

            try:
                node.ensure_visible()
            except NoPatternInterfaceError:
                if not node.is_visible():
                    raise SajError("A página do AR está fora da área visível da árvore do SAJ.")
        self._check_foreground()
        node.click_input()
        fit_page = self.config.get("fit_page_control")
        if fit_page and kind != "AR":
            self._select(fit_page).click_input()
        deadline = time.monotonic() + self.timeout
        previous_checksum = self._content_checksum(previous, region)
        last_checksum = None
        header_checksum = None
        caption = ""
        same_page = None
        stable_since = 0.0
        while time.monotonic() < deadline:
            snapshot = self._snapshot()
            capture = self._crop(snapshot, region)
            header = self._crop(snapshot, self.caption_region)
            new_header_checksum = blake2b(header.tobytes()).digest()
            if uses_control or header_checksum != new_header_checksum:
                caption = self._caption(header)
                header_checksum = new_header_checksum
            checksum = self._content_checksum(snapshot, region)
            if checksum != last_checksum:
                last_checksum = checksum
                stable_since = time.monotonic()
            if self._caption_confirms(caption, page, kind) and time.monotonic() - stable_since >= self.settle:
                if checksum != previous_checksum:
                    return capture
                if same_page is None:
                    if old_caption is None:
                        old_caption = self._caption(old_header)
                    same_page = self._caption_confirms(old_caption, page, kind)
                if same_page:
                    return capture
            time.sleep(0.15)
        raise SajError(f"O painel não confirmou a imagem da página {page}.")

    def _content_checksum(self, snapshot: Image.Image, region: tuple) -> bytes:
        left, _, right, bottom = region
        top = max(region[1], self.caption_region[3])
        content = self._crop(snapshot, (left, top, right, bottom)) if top < bottom else self._crop(snapshot, region)
        return blake2b(content.tobytes()).digest()

    def _document_text(self) -> str | None:
        selector = self.config.get("document_text_control")
        if not selector or self._backend != "uia":
            return None
        from pywinauto.uia_defines import NoPatternInterfaceError

        try:
            control = self._select(selector)
            return control.iface_text.DocumentRange.GetText(-1).strip() or None
        except NoPatternInterfaceError:
            return None

    @staticmethod
    def _whole_sheet(image: Image.Image) -> Image.Image | None:
        # Page edges need to be visible in the gray viewer. A page touching
        # the viewport or connected to the metadata banner may be clipped.
        import cv2
        import numpy as np

        gray = np.asarray(ImageOps.grayscale(image))
        _, paper = cv2.threshold(gray, 248, 255, cv2.THRESH_BINARY)
        contours, _ = cv2.findContours(paper, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if not contours:
            return None
        sheet = max(contours, key=cv2.contourArea)
        x, y, width, height = cv2.boundingRect(sheet)
        margin = max(3, min(image.size) // 100)
        if (
            x < margin or y < margin
            or x + width > image.width - margin or y + height > image.height - margin
            or width * height < 0.12 * image.width * image.height
            or not 0.5 <= width / height <= 2
            or cv2.contourArea(sheet) < 0.86 * width * height
        ):
            return None
        return image.crop((x, y, x + width, y + height))

    def _document_nodes(self, root) -> Iterator[tuple[str, list, str | None]]:
        stack = [(node, None) for node in reversed(self._children(root))]
        while stack:
            node, parent_kind = stack.pop()
            kind = _document_kind(self._item_text(node)) or parent_kind
            try:
                children = self._children(node)
            except Exception as exc:
                if kind is None:
                    raise
                yield kind, [], f"Falha ao expandir o documento no SAJ. {exc}"
                continue

            pages = []
            groups = []
            for child in children:
                page = _PAGE.match(_normalize(self._item_text(child)))
                if page:
                    pages.append((page.group(1), child))
                else:
                    groups.append(child)
            if kind and (pages or not groups):
                yield kind, pages, None
            stack.extend((child, kind) for child in reversed(groups))

    def iter_documents(self, *, only: str | None = None) -> Iterator[DocumentCapture]:
        if self._process is None:
            raise SajError("Abra o processo antes de percorrer os documentos.")
        try:
            current, root = self._current_root()
            if current != self._process:
                raise SajError("O processo exibido mudou antes da consulta dos documentos.")
            documents = [
                (kind, pages, error) for kind, pages, error in self._document_nodes(root)
                if only is None or kind == only
            ]
        except SajError:
            raise
        except Exception as exc:
            raise SajError(f"O inventário de documentos do processo falhou. {exc}") from exc

        for index, (kind, pages, error) in enumerate(documents, 1):
            if not pages:
                yield DocumentCapture(
                    "", None, index, error or "Documento sem página identificável na árvore.", kind, False,
                )
                continue

            for page, node in pages:
                try:
                    image = self._capture_page(node, page, kind)
                    text = self._document_text() if kind != "AR" else None
                    if kind != "AR" and not text:
                        sheet = self._whole_sheet(image)
                        if sheet is None:
                            yield DocumentCapture(
                                page, image, index, "A página inteira não ficou visível. Ajuste o zoom do leitor.",
                                kind, False,
                            )
                            continue
                        image = sheet
                except Exception as exc:
                    yield DocumentCapture(page, None, index, str(exc), kind, False)
                    continue
                yield DocumentCapture(page, image, index, kind=kind, text=text)

        try:
            current, _ = self._current_root()
            if current != self._process:
                raise SajError("O processo exibido mudou durante a consulta dos documentos.")
        except SajError:
            raise
        except Exception as exc:
            raise SajError(f"Falha ao conferir o processo após a consulta dos documentos. {exc}") from exc

    def iter_receipts(self) -> Iterator[ReceiptCapture]:
        return self.iter_documents(only="AR")
