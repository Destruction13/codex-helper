import base64
import ctypes
import html
import imaplib
import os
import re
import select
import shutil
import socket
import socketserver
import subprocess
import sys
import threading
import time
import traceback
import webbrowser
from ctypes import wintypes
from email import message_from_bytes
from email.header import decode_header, make_header
from pathlib import Path
from typing import Any

import keyboard
import paramiko
from PySide6.QtCore import QObject, QSettings, QTimer, Signal
from PySide6.QtGui import QKeySequence
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QDialogButtonBox,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QKeySequenceEdit,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QSizePolicy,
    QStyle,
    QSystemTrayIcon,
    QVBoxLayout,
    QWidget,
)


ENV_FILE = Path(__file__).resolve().with_name(".env")
LOG_FILE = Path(__file__).resolve().with_name("codex_helper.log")


def load_env_file(path: Path) -> None:
    if not path.exists():
        return

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        normalized_key = key.strip()
        if not normalized_key:
            continue

        normalized_value = value.strip()
        if (
            len(normalized_value) >= 2
            and normalized_value[0] == normalized_value[-1]
            and normalized_value[0] in {'"', "'"}
        ):
            normalized_value = normalized_value[1:-1]

        os.environ.setdefault(normalized_key, normalized_value)


load_env_file(ENV_FILE)


APP_NAME = "Codex helper"
SETTINGS_ORG = "OpenCode"
SETTINGS_APP = "CodexHelperLite"
ZOHO_IMAP_PORT = 993
IMAP_POLL_INTERVAL_SECONDS = 5
CHATGPT_URL = "https://chatgpt.com"
CHATGPT_SIGNUP_BUTTON_TEXT = "Зарегистрироваться бесплатно"
CHATGPT_CREATE_PASSWORD_TEXTS = (
    "создать пароль",
    "create a password",
    "create your password",
)
CHATGPT_PASSWORD_LABEL_TEXTS = (
    "пароль",
    "password",
)
CHATGPT_CHECK_EMAIL_TEXTS = (
    "проверьте свою почту",
    "check your email",
)
CHATGPT_INVALID_CODE_TEXTS = (
    "неправильный код",
    "неверный код",
    "invalid code",
    "incorrect code",
)
CHATGPT_BUTTON_TIMEOUT_SECONDS = 15.0
CHATGPT_START_DELAY_SECONDS = 2.0
CHATGPT_EMAIL_STEP_DELAY_SECONDS = 1.0
CHATGPT_REGISTRATION_WATCH_TIMEOUT_SECONDS = 180.0
CHATGPT_REGISTRATION_POLL_INTERVAL_SECONDS = 1.0
CHATGPT_PROFILE_STEP_DELAY_SECONDS = 2.0
CHATGPT_STARTUP_GUARD_TIMEOUT_SECONDS = 15.0
CHATGPT_FINISH_ACCOUNT_BUTTON_TEXTS = (
    "Завершить создание учетной записи",
    "Finish creating account",
)
OMNIROUTE_PAGE_PATH = "/dashboard/providers/codex"
OMNIROUTE_ADD_BUTTON_TEXTS = (
    "+Add",
    "Add",
)
OMNIROUTE_PASSWORD_STEP_TEXTS = (
    "введите ваш пароль",
    "enter your password",
)
OMNIROUTE_CHECK_EMAIL_TEXTS = CHATGPT_CHECK_EMAIL_TEXTS
OMNIROUTE_CONTINUE_BUTTON_TEXTS = (
    "Продолжить",
    "Continue",
)
OMNIROUTE_INVALID_CODE_TEXTS = (
    "неверный код",
    "неверно",
    "invalid code",
    "incorrect code",
)
OMNIROUTE_REJECT_TEXTS = (
    "давайте подтвердим ваш возраст",
    "требуется номер телефона",
)
OMNIROUTE_PAGE_LOAD_DELAY_SECONDS = 0.3
OMNIROUTE_STEP_TIMEOUT_SECONDS = 180.0
OMNIROUTE_ADD_TIMEOUT_SECONDS = 20.0
OMNIROUTE_POST_PASSWORD_TIMEOUT_SECONDS = 20.0
CODE_WAIT_TIMEOUT_SECONDS = 30.0
MAX_CODE_RESEND_ATTEMPTS = 2
INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004
VK_TAB = 0x09
VK_SHIFT = 0x10
VK_CONTROL = 0x11
VK_BACK = 0x08
VK_MENU = 0x12
VK_RETURN = 0x0D
VK_TAB = 0x09

user32 = ctypes.WinDLL("user32", use_last_error=True)


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(wintypes.ULONG)),
    ]


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", wintypes.LONG),
        ("dy", wintypes.LONG),
        ("mouseData", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(wintypes.ULONG)),
    ]


class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ("uMsg", wintypes.DWORD),
        ("wParamL", wintypes.WORD),
        ("wParamH", wintypes.WORD),
    ]


class _INPUT_UNION(ctypes.Union):
    _fields_ = [
        ("ki", KEYBDINPUT),
        ("mi", MOUSEINPUT),
        ("hi", HARDWAREINPUT),
    ]


class INPUT(ctypes.Structure):
    _anonymous_ = ("u",)
    _fields_ = [("type", wintypes.DWORD), ("u", _INPUT_UNION)]


user32.SendInput.argtypes = (wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int)
user32.SendInput.restype = wintypes.UINT


def build_alias_email(email: str, next_index: int) -> str:
    normalized = email.strip()
    if not normalized:
        raise ValueError("Enter a Gmail address first.")
    if "@" not in normalized:
        raise ValueError("Email must contain '@'.")

    local_part, domain = normalized.split("@", 1)
    if not local_part or not domain:
        raise ValueError("Email must include both local part and domain.")

    base_local = local_part.split("+", 1)[0]
    if next_index <= 0:
        return f"{base_local}@{domain}"
    return f"{base_local}+max{next_index}@{domain}"


def extract_alias_index(email: str) -> int:
    normalized = email.strip()
    if "@" not in normalized:
        return 0
    local_part = normalized.split("@", 1)[0]
    if "+" not in local_part:
        return 0
    suffix = local_part.rsplit("+", 1)[1]
    if suffix.isdigit():
        return int(suffix)
    match = re.fullmatch(r"max(\d+)", suffix, flags=re.IGNORECASE)
    return int(match.group(1)) if match else 0


def decode_message_header(value: str) -> str:
    if not value:
        return ""
    return str(make_header(decode_header(value)))


def html_to_text(value: str) -> str:
    no_breaks = re.sub(r"<br\s*/?>", "\n", value, flags=re.IGNORECASE)
    without_tags = re.sub(r"<[^>]+>", " ", no_breaks)
    return html.unescape(without_tags)


def extract_text_from_message(raw_message: bytes) -> str:
    message = message_from_bytes(raw_message)
    parts: list[str] = []

    if message.is_multipart():
        for part in message.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if part.get("Content-Disposition", "").lower().startswith("attachment"):
                continue
            content_type = part.get_content_type()
            raw_payload = part.get_payload(decode=True)
            if not isinstance(raw_payload, bytes):
                continue
            charset = part.get_content_charset() or "utf-8"
            decoded = raw_payload.decode(charset, errors="replace")
            if content_type == "text/html":
                decoded = html_to_text(decoded)
            parts.append(decoded)
    else:
        raw_payload = message.get_payload(decode=True)
        if isinstance(raw_payload, bytes):
            charset = message.get_content_charset() or "utf-8"
            decoded = raw_payload.decode(charset, errors="replace")
            if message.get_content_type() == "text/html":
                decoded = html_to_text(decoded)
            parts.append(decoded)

    headers = " ".join(
        [
            decode_message_header(message.get("From", "")),
            decode_message_header(message.get("Subject", "")),
        ]
    )
    combined = "\n".join(part.strip() for part in parts if part.strip())
    return f"{headers}\n{combined}".strip()


def extract_openai_code(text: str) -> str:
    strict_match = re.search(r"\bCode:\s*(\d{4,8})\b", text, flags=re.IGNORECASE)
    if strict_match:
        return strict_match.group(1)

    localized_match = re.search(
        r"\b(?:ваш\s+код\s+chatgpt|chatgpt\s+code|verification\s+code)\D*(\d{4,8})\b",
        text,
        flags=re.IGNORECASE,
    )
    if localized_match:
        return localized_match.group(1)

    if "openai" not in text.lower():
        return ""

    fallback_match = re.search(r"\b(\d{6})\b", text)
    return fallback_match.group(1) if fallback_match else ""


def fetch_latest_uid(mailbox: imaplib.IMAP4_SSL) -> int:
    status, data = mailbox.uid("SEARCH", "ALL")
    if status != "OK" or not data or not data[0]:
        return 0
    values = [int(raw_uid) for raw_uid in data[0].split() if raw_uid.isdigit()]
    return values[-1] if values else 0


def resolve_zoho_imap_host(email: str) -> str:
    normalized = email.strip().lower()
    if normalized.endswith(".eu"):
        return "imap.zoho.eu"
    return "imap.zoho.com"


def normalize_zoho_password(password: str) -> str:
    stripped = password.strip()
    if any(char.isspace() for char in stripped):
        collapsed = re.sub(r"\s+", "", stripped)
        if len(collapsed) >= 12:
            return collapsed
    return stripped


def list_imap_mailboxes(mailbox: imaplib.IMAP4_SSL) -> list[str]:
    status, data = mailbox.list()
    if status != "OK" or not data:
        return ["INBOX"]

    folders: list[str] = []
    for raw_item in data:
        if not raw_item:
            continue
        decoded = (
            raw_item.decode(errors="replace")
            if isinstance(raw_item, bytes)
            else str(raw_item)
        )
        quoted_match = re.search(r'"([^"]+)"\s*$', decoded)
        if quoted_match:
            folders.append(quoted_match.group(1))
            continue
        atom_match = re.search(r'\s([^\s"]+)$', decoded)
        if atom_match:
            folders.append(atom_match.group(1))

    unique_folders = list(dict.fromkeys(folder for folder in folders if folder))
    if "INBOX" in unique_folders:
        unique_folders.remove("INBOX")
        unique_folders.insert(0, "INBOX")
    return unique_folders or ["INBOX"]


def verify_zoho_imap_credentials(email: str, password: str) -> None:
    mailbox: imaplib.IMAP4_SSL | None = None
    try:
        mailbox = imaplib.IMAP4_SSL(resolve_zoho_imap_host(email), ZOHO_IMAP_PORT)
        mailbox.login(email, password)
        select_status, _ = mailbox.select("INBOX")
        if select_status != "OK":
            raise RuntimeError("Failed to open Zoho inbox.")
    finally:
        if mailbox is not None:
            try:
                mailbox.logout()
            except Exception:
                pass


def extract_message_bytes(fetch_data: list[tuple[bytes, bytes] | bytes]) -> bytes:
    for item in fetch_data:
        if isinstance(item, tuple) and len(item) > 1 and isinstance(item[1], bytes):
            return item[1]
    return b""


def fetch_latest_openai_code(
    mailbox: imaplib.IMAP4_SSL, min_uid: int
) -> tuple[int, str]:
    status, data = mailbox.uid("SEARCH", "ALL")
    if status != "OK" or not data or not data[0]:
        return min_uid, ""

    new_uids = [int(raw_uid) for raw_uid in data[0].split() if raw_uid.isdigit()]
    new_uids = [uid for uid in new_uids if uid > min_uid]
    if not new_uids:
        return min_uid, ""

    latest_uid = new_uids[-1]
    for uid in reversed(new_uids):
        fetch_status, fetch_data = mailbox.uid("FETCH", str(uid), "(RFC822)")
        if fetch_status != "OK" or not fetch_data:
            continue
        raw_message = extract_message_bytes(fetch_data)
        if not raw_message:
            continue
        code = extract_openai_code(extract_text_from_message(raw_message))
        if code:
            return latest_uid, code

    return latest_uid, ""


def fetch_unseen_openai_code_from_all_folders(
    mailbox: imaplib.IMAP4_SSL,
    processed_message_ids: set[str],
) -> tuple[set[str], str]:
    latest_code = ""
    discovered_ids: set[str] = set()

    for folder in list_imap_mailboxes(mailbox):
        select_status, _ = mailbox.select(folder, readonly=True)
        if select_status != "OK":
            continue

        search_status, search_data = mailbox.uid("SEARCH", "UNSEEN")
        if search_status != "OK" or not search_data or not search_data[0]:
            continue

        unseen_uids = [
            int(raw_uid) for raw_uid in search_data[0].split() if raw_uid.isdigit()
        ]
        for uid in reversed(unseen_uids):
            message_id = f"{folder}:{uid}"
            discovered_ids.add(message_id)
            if message_id in processed_message_ids:
                continue

            fetch_status, fetch_data = mailbox.uid("FETCH", str(uid), "(BODY.PEEK[])")
            if fetch_status != "OK" or not fetch_data:
                continue
            raw_message = extract_message_bytes(fetch_data)
            if not raw_message:
                continue
            code = extract_openai_code(extract_text_from_message(raw_message))
            if code:
                latest_code = code

    return discovered_ids, latest_code


def send_virtual_key(vk_code: int, *, key_up: bool = False) -> None:
    extra = ctypes.pointer(wintypes.ULONG(0))
    key_input = KEYBDINPUT(
        wVk=vk_code,
        wScan=0,
        dwFlags=KEYEVENTF_KEYUP if key_up else 0,
        time=0,
        dwExtraInfo=extra,
    )
    input_record = INPUT(type=INPUT_KEYBOARD, ki=key_input)
    result = user32.SendInput(1, ctypes.byref(input_record), ctypes.sizeof(INPUT))
    if result != 1:
        raise OSError(ctypes.get_last_error(), "SendInput failed for virtual key.")


def send_unicode_char(char: str) -> None:
    extra = ctypes.pointer(wintypes.ULONG(0))
    code_point = ord(char)

    key_down = KEYBDINPUT(
        wVk=0,
        wScan=code_point,
        dwFlags=KEYEVENTF_UNICODE,
        time=0,
        dwExtraInfo=extra,
    )
    key_up = KEYBDINPUT(
        wVk=0,
        wScan=code_point,
        dwFlags=KEYEVENTF_UNICODE | KEYEVENTF_KEYUP,
        time=0,
        dwExtraInfo=extra,
    )

    down_record = INPUT(type=INPUT_KEYBOARD, ki=key_down)
    up_record = INPUT(type=INPUT_KEYBOARD, ki=key_up)

    result = user32.SendInput(1, ctypes.byref(down_record), ctypes.sizeof(INPUT))
    if result != 1:
        raise OSError(ctypes.get_last_error(), "SendInput failed for key down.")

    result = user32.SendInput(1, ctypes.byref(up_record), ctypes.sizeof(INPUT))
    if result != 1:
        raise OSError(ctypes.get_last_error(), "SendInput failed for key up.")


def send_unicode_text(text: str) -> None:
    normalized_text = text.replace("\r\n", "\n").replace("\r", "\n")
    for char in normalized_text:
        if char == "\n":
            send_virtual_key(VK_RETURN, key_up=False)
            send_virtual_key(VK_RETURN, key_up=True)
            continue
        send_unicode_char(char)


def press_virtual_key(vk_code: int) -> None:
    send_virtual_key(vk_code, key_up=False)
    send_virtual_key(vk_code, key_up=True)


def release_modifier_keys() -> None:
    for vk_code in (VK_MENU, VK_CONTROL, VK_SHIFT):
        user32.keybd_event(vk_code, 0, KEYEVENTF_KEYUP, 0)


def run_powershell_script(
    script: str, *, timeout_seconds: float
) -> subprocess.CompletedProcess[str]:
    encoded_script = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
    return subprocess.run(
        ["powershell", "-NoProfile", "-EncodedCommand", encoded_script],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        timeout=max(5, int(timeout_seconds) + 5),
    )


def completed_process_output(completed: subprocess.CompletedProcess[str]) -> str:
    stdout = completed.stdout or ""
    stderr = completed.stderr or ""
    combined = stderr if stderr.strip() else stdout
    return combined.strip()


def require_env(name: str) -> str:
    value = os.getenv(name, "")
    if not value.strip():
        raise RuntimeError(f"Missing {name} in .env. Use .env.example as a template.")
    return value.strip()


def append_traceback_to_log(
    context: str, exc_info: tuple[type[BaseException], BaseException, Any]
) -> None:
    timestamp = time.strftime("%H:%M:%S")
    formatted = "".join(traceback.format_exception(*exc_info)).rstrip()
    try:
        LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
        with LOG_FILE.open("a", encoding="utf-8") as log_file:
            log_file.write(f"[{timestamp}] {context}\n{formatted}\n")
    except Exception:
        pass


def install_global_exception_logging() -> None:
    def _sys_excepthook(
        exc_type: type[BaseException], exc_value: BaseException, exc_traceback: Any
    ) -> None:
        append_traceback_to_log(
            "UNHANDLED EXCEPTION",
            (exc_type, exc_value, exc_traceback),
        )
        sys.__excepthook__(exc_type, exc_value, exc_traceback)

    def _thread_excepthook(args: threading.ExceptHookArgs) -> None:
        append_traceback_to_log(
            f"UNHANDLED THREAD EXCEPTION: {getattr(args.thread, 'name', 'unknown-thread')}",
            (args.exc_type, args.exc_value, args.exc_traceback),
        )

    sys.excepthook = _sys_excepthook
    threading.excepthook = _thread_excepthook


class ThreadingForwardServer(socketserver.ThreadingTCPServer):
    allow_reuse_address = True
    daemon_threads = True


def build_forward_handler(
    transport: paramiko.Transport, remote_host: str, remote_port: int
) -> type[socketserver.BaseRequestHandler]:
    class ForwardHandler(socketserver.BaseRequestHandler):
        def handle(self) -> None:
            channel = None
            try:
                channel = transport.open_channel(
                    "direct-tcpip",
                    (remote_host, remote_port),
                    self.request.getpeername(),
                )
            except Exception:
                return

            if channel is None:
                return

            try:
                while True:
                    ready_to_read, _, _ = select.select([self.request, channel], [], [])
                    if self.request in ready_to_read:
                        data = self.request.recv(4096)
                        if not data:
                            break
                        channel.sendall(data)
                    if channel in ready_to_read:
                        data = channel.recv(4096)
                        if not data:
                            break
                        self.request.sendall(data)
            finally:
                try:
                    channel.close()
                except Exception:
                    pass
                self.request.close()

    return ForwardHandler


class SshTunnelManager:
    def __init__(self) -> None:
        self.client: paramiko.SSHClient | None = None
        self.servers: list[ThreadingForwardServer] = []
        self.server_threads: list[threading.Thread] = []

    def is_active(self) -> bool:
        transport = self.client.get_transport() if self.client is not None else None
        return bool(transport and transport.is_active() and self.servers)

    def start(
        self,
        *,
        host: str,
        port: int,
        username: str,
        password: str,
        forwards: list[tuple[int, str, int]],
    ) -> None:
        if self.is_active():
            return

        self.stop()
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(
            hostname=host,
            port=port,
            username=username,
            password=password,
            allow_agent=False,
            look_for_keys=False,
            timeout=15,
            auth_timeout=15,
            banner_timeout=15,
        )

        transport = client.get_transport()
        if transport is None or not transport.is_active():
            client.close()
            raise RuntimeError("SSH transport did not become ready.")

        created_servers: list[ThreadingForwardServer] = []
        created_threads: list[threading.Thread] = []
        try:
            for local_port, remote_host, remote_port in forwards:
                server = ThreadingForwardServer(
                    ("127.0.0.1", local_port),
                    build_forward_handler(transport, remote_host, remote_port),
                )
                thread = threading.Thread(target=server.serve_forever, daemon=True)
                thread.start()
                created_servers.append(server)
                created_threads.append(thread)
        except Exception:
            for server in created_servers:
                server.shutdown()
                server.server_close()
            client.close()
            raise

        self.client = client
        self.servers = created_servers
        self.server_threads = created_threads

    def stop(self) -> None:
        for server in self.servers:
            try:
                server.shutdown()
            except Exception:
                pass
            try:
                server.server_close()
            except Exception:
                pass

        self.servers = []
        self.server_threads = []

        if self.client is not None:
            try:
                self.client.close()
            except Exception:
                pass
            self.client = None


def key_sequence_to_hotkey(sequence: QKeySequence) -> str:
    portable_text = sequence.toString(QKeySequence.SequenceFormat.PortableText).strip()
    if not portable_text:
        return ""

    normalized_parts: list[str] = []
    for part in portable_text.split("+"):
        value = part.strip()
        if not value:
            continue
        lowered = value.lower()
        if lowered == "ctrl":
            normalized_parts.append("ctrl")
        elif lowered == "alt":
            normalized_parts.append("alt")
        elif lowered == "shift":
            normalized_parts.append("shift")
        elif lowered in {"meta", "win"}:
            normalized_parts.append("windows")
        else:
            normalized_parts.append(lowered)
    return "+".join(normalized_parts)


def hotkey_label(sequence: QKeySequence) -> str:
    return sequence.toString(QKeySequence.SequenceFormat.NativeText).strip()


class KBindRow:
    def __init__(self) -> None:
        self.handle: Any | None = None
        self.widget = QWidget()
        self.layout = QHBoxLayout(self.widget)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(8)

        self.hotkey_edit = QKeySequenceEdit()
        self.hotkey_edit.setMaximumSequenceLength(1)
        self.hotkey_edit.setFixedWidth(160)

        self.text_input = QLineEdit()
        self.text_input.setPlaceholderText("Text to insert")

        self.delete_button = QPushButton("Delete")
        self.delete_button.setFixedWidth(90)

        self.layout.addWidget(self.hotkey_edit)
        self.layout.addWidget(self.text_input, 1)
        self.layout.addWidget(self.delete_button)


class ZohoLoginDialog(QDialog):
    def __init__(self, parent: QWidget | None = None, *, email: str = "") -> None:
        super().__init__(parent)
        self.setWindowTitle("Zoho Login")
        self.setModal(True)
        self.resize(420, 150)

        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(16, 16, 16, 16)
        root_layout.setSpacing(10)

        hint = QLabel(
            "Enter your Zoho Mail address and password. The app will save this account and reuse it until you log out."
        )
        hint.setWordWrap(True)

        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("ghoulgpt@zohomail.eu")
        self.email_input.setText(email)

        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Zoho password")
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        root_layout.addWidget(hint)
        root_layout.addWidget(self.email_input)
        root_layout.addWidget(self.password_input)
        root_layout.addWidget(buttons)

    def credentials(self) -> tuple[str, str]:
        return self.email_input.text().strip(), self.password_input.text()


class WorkerSignals(QObject):
    text_bind_requested = Signal(object)
    email_copy_requested = Signal()
    code_insert_requested = Signal()
    inspect_window_requested = Signal()
    text_bind_error = Signal(str)
    text_bind_success = Signal(str)
    imap_code_received = Signal(str)
    registration_password_step_detected = Signal()
    registration_email_check_step_detected = Signal()
    imap_status = Signal(str)
    imap_error = Signal(str)
    server_connected = Signal(str)
    server_connection_failed = Signal(str)
    omniroute_failed = Signal(str)
    omniroute_status = Signal(str)
    omniroute_finished = Signal()
    registration_flow_completed = Signal()
    reg_cycle_finished = Signal()
    registration_post_submit = Signal()
    code_wait_timeout = Signal(str, int)


class CodexApp(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.settings = QSettings(SETTINGS_ORG, SETTINGS_APP)
        self.signals = WorkerSignals()
        self.bind_rows: list[KBindRow] = []
        self.email_hotkey_handle: Any | None = None
        self.code_hotkey_handle: Any | None = None
        self.inspect_window_hotkey_handle: Any | None = None
        self.imap_stop_event = threading.Event()
        self.imap_thread: threading.Thread | None = None
        self.registration_watch_stop_event = threading.Event()
        self.registration_watch_thread: threading.Thread | None = None
        self.awaiting_registration_code = False
        self.registration_code_baseline = ""
        self.pending_code_flow = ""
        self.pending_code_process_id = 0
        self.server_tunnel = SshTunnelManager()
        self.server_connection_in_progress = False
        self.omniroute_in_progress = False
        self.reg_in_progress = False
        self.reg_automation_active = False
        self.chatgpt_edge_process_id = 0
        self.reuse_omniroute_browser = False
        self.omniroute_restart_requested = False
        self.code_submit_in_progress = False
        self.code_wait_refocus_needed = False
        self.code_wait_clear_before_submit = False
        self.code_wait_generation = 0
        self.code_wait_flow = ""
        self.code_wait_process_id = 0
        self.code_wait_attempts = 0
        self.retry_hotkey_in_progress = False

        self.email_input = QLineEdit()
        self.email_hotkey_edit = QKeySequenceEdit()
        self.plus_button = QPushButton("+1")
        self.imap_account_label = QLabel("Not signed in")
        self.code_value_input = QLineEdit()
        self.code_hotkey_edit = QKeySequenceEdit()
        self.login_imap_button = QPushButton("Login")
        self.logout_imap_button = QPushButton("Logout")
        self.reg_button = QPushButton("reg")
        self.stop_button = QPushButton("Stop")
        self.connect_server_button = QPushButton("Connect to server")
        self.open_logs_button = QPushButton("Open logs")
        self.bind_rows_widget = QWidget()
        self.bind_rows_layout = QVBoxLayout(self.bind_rows_widget)
        self.add_bind_button = QPushButton("Add K-Bind")
        self.save_binds_button = QPushButton("Save Binds")
        self.log_output: QPlainTextEdit | None = None
        self.tray_icon = QSystemTrayIcon(self)

        self.setup_ui()
        self.restore_state()
        self.connect_signals()
        try:
            self.register_text_hotkeys()
        except Exception as exc:
            self.append_log(f"Failed to restore K-Bind: {exc}")
        self.start_imap_monitor(log_startup=False)
        self.refresh_button_state()

    def setup_ui(self) -> None:
        self.setWindowTitle(APP_NAME)
        self.resize(720, 360)
        self.tray_icon.setIcon(
            self.style().standardIcon(QStyle.StandardPixmap.SP_MessageBoxInformation)
        )
        self.tray_icon.setToolTip(APP_NAME)
        self.tray_icon.show()

        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(16, 16, 16, 16)
        root_layout.setSpacing(12)

        account_group = QGroupBox("Аккаунт")
        account_layout = QHBoxLayout(account_group)
        account_layout.setContentsMargins(12, 12, 12, 12)
        self.email_input.setPlaceholderText("example@gmail.com")
        account_layout.addWidget(self.email_input, 1)

        imap_group = QGroupBox("Zoho IMAP")
        imap_layout = QVBoxLayout(imap_group)
        imap_layout.setContentsMargins(12, 12, 12, 12)
        imap_layout.setSpacing(8)

        imap_account_row = QWidget()
        imap_account_layout = QHBoxLayout(imap_account_row)
        imap_account_layout.setContentsMargins(0, 0, 0, 0)
        imap_account_layout.setSpacing(8)
        self.login_imap_button.setFixedWidth(44)
        self.logout_imap_button.setFixedWidth(44)
        self.login_imap_button.setText("↪")
        self.logout_imap_button.setText("⎋")
        self.login_imap_button.setStyleSheet(
            "QPushButton { background-color: #22c55e; color: white; font-weight: 700; border-radius: 6px; padding: 6px; }"
        )
        self.logout_imap_button.setStyleSheet(
            "QPushButton { background-color: #ef4444; color: white; font-weight: 700; border-radius: 6px; padding: 6px; }"
        )
        self.imap_account_label.setWordWrap(True)
        imap_account_layout.addWidget(QLabel("Account:"))
        imap_account_layout.addWidget(self.imap_account_label, 1)
        imap_account_layout.addWidget(self.login_imap_button)
        imap_account_layout.addWidget(self.logout_imap_button)

        imap_layout.addWidget(imap_account_row)

        server_group = QGroupBox("Сервер")
        server_layout = QHBoxLayout(server_group)
        server_layout.setContentsMargins(12, 12, 12, 12)
        self.connect_server_button.setMinimumHeight(36)
        self.connect_server_button.setText("Подключиться к серверу")
        server_layout.addWidget(self.connect_server_button)

        registration_group = QGroupBox("Регистрация")
        registration_layout = QHBoxLayout(registration_group)
        registration_layout.setContentsMargins(12, 12, 12, 12)
        self.reg_button.setText("Старт")
        self.stop_button.setText("Стоп")
        self.reg_button.setMinimumHeight(36)
        self.stop_button.setMinimumHeight(36)
        registration_layout.addWidget(self.reg_button)
        registration_layout.addWidget(self.stop_button)

        logs_group = QGroupBox("Логи")
        logs_layout = QHBoxLayout(logs_group)
        logs_layout.setContentsMargins(12, 12, 12, 12)
        self.open_logs_button.setMinimumHeight(36)
        self.open_logs_button.setText("Открыть логи")
        logs_layout.addWidget(self.open_logs_button)

        root_layout.addWidget(account_group)
        root_layout.addWidget(imap_group)
        root_layout.addWidget(server_group)
        root_layout.addWidget(registration_group)
        root_layout.addWidget(logs_group)

    def connect_signals(self) -> None:
        self.plus_button.clicked.connect(self.increment_email)
        self.login_imap_button.clicked.connect(self.open_imap_login_dialog)
        self.logout_imap_button.clicked.connect(self.logout_imap_account)
        self.reg_button.clicked.connect(self.handle_reg_button_clicked)
        self.stop_button.clicked.connect(self.handle_stop_button_clicked)
        self.connect_server_button.clicked.connect(self.handle_server_button_clicked)
        self.open_logs_button.clicked.connect(self.open_logs_file)
        self.add_bind_button.clicked.connect(self.add_bind_row)
        self.save_binds_button.clicked.connect(self.save_text_binds)
        self.email_input.textChanged.connect(self.handle_email_changed)
        self.signals.email_copy_requested.connect(self.copy_current_email_to_clipboard)
        self.signals.code_insert_requested.connect(self.insert_current_code)
        self.signals.inspect_window_requested.connect(self.inspect_current_window)
        self.signals.text_bind_requested.connect(self.trigger_text_bind)
        self.signals.text_bind_error.connect(self.handle_text_bind_error)
        self.signals.text_bind_success.connect(self.handle_text_bind_success)
        self.signals.imap_code_received.connect(self.handle_imap_code_received)
        self.signals.registration_password_step_detected.connect(
            self.handle_registration_password_step_detected
        )
        self.signals.registration_email_check_step_detected.connect(
            self.handle_registration_email_check_step_detected
        )
        self.signals.imap_status.connect(self.append_log)
        self.signals.imap_error.connect(self.handle_imap_error)
        self.signals.server_connected.connect(self.handle_server_connected)
        self.signals.server_connection_failed.connect(
            self.handle_server_connection_failed
        )
        self.signals.omniroute_failed.connect(self.handle_omniroute_failed)
        self.signals.omniroute_status.connect(self.append_log)
        self.signals.omniroute_finished.connect(self.handle_omniroute_finished)
        self.signals.registration_flow_completed.connect(
            self.handle_registration_flow_completed
        )
        self.signals.reg_cycle_finished.connect(self.handle_reg_cycle_finished)
        self.signals.registration_post_submit.connect(
            self.handle_registration_post_submit
        )
        self.signals.code_wait_timeout.connect(self.handle_code_wait_timeout)

    def restore_state(self) -> None:
        saved_email = str(self.settings.value("email", "")).strip()
        saved_email_hotkey = str(self.settings.value("email_hotkey", "")).strip()
        saved_imap_email = str(self.settings.value("imap_email", "")).strip()
        saved_imap_password = str(self.settings.value("imap_password", ""))
        saved_code_hotkey = str(self.settings.value("code_hotkey", "")).strip()

        if saved_email:
            self.email_input.setText(saved_email)
        if saved_email_hotkey:
            self.email_hotkey_edit.setKeySequence(QKeySequence(saved_email_hotkey))
        if saved_code_hotkey:
            self.code_hotkey_edit.setKeySequence(QKeySequence(saved_code_hotkey))

        self.update_imap_account_label(saved_imap_email)

        size = self.settings.beginReadArray("text_binds")
        for index in range(size):
            self.settings.setArrayIndex(index)
            hotkey = str(self.settings.value("hotkey", "")).strip()
            text = str(self.settings.value("text", ""))
            self.add_bind_row(hotkey=hotkey, text=text)
        self.settings.endArray()

        if not self.bind_rows:
            self.add_bind_row()

    def closeEvent(self, event) -> None:  # noqa: N802
        self.settings.setValue("email", self.current_email())
        self.persist_imap_settings()
        self.persist_text_bind_settings()
        self.stop_registration_watch()
        self.stop_imap_monitor()
        self.server_tunnel.stop()
        self.unregister_text_hotkeys()
        super().closeEvent(event)

    def append_log(self, message: str) -> None:
        timestamp = time.strftime("%H:%M:%S")
        log_line = f"[{timestamp}] {message}"
        try:
            LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
            with LOG_FILE.open("a", encoding="utf-8") as log_file:
                log_file.write(log_line + "\n")
        except Exception:
            pass

        if self.log_output is not None:
            self.log_output.appendPlainText(log_line)

    def open_logs_file(self) -> None:
        try:
            LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
            if not LOG_FILE.exists():
                LOG_FILE.touch()
            os.startfile(str(LOG_FILE))
        except Exception as exc:
            self.show_error("Open logs failed", str(exc))

    def show_error(self, title: str, message: str) -> None:
        QMessageBox.critical(self, title, message)

    def clipboard_text(self, text: str) -> None:
        clipboard = QApplication.clipboard()
        if clipboard is None:
            raise RuntimeError("Clipboard is not available.")
        clipboard.setText(text)

    def current_email(self) -> str:
        return self.email_input.text().strip()

    def current_code(self) -> str:
        return self.code_value_input.text().strip()

    def handle_email_changed(self, value: str) -> None:
        self.settings.setValue("email", value.strip())
        self.refresh_button_state()

    def refresh_button_state(self) -> None:
        self.plus_button.setEnabled(bool(self.current_email()))
        has_imap_session = bool(self.stored_imap_email()) and bool(
            self.stored_imap_password()
        )
        self.logout_imap_button.setEnabled(has_imap_session)
        self.reg_button.setEnabled(not self.reg_in_progress)
        self.stop_button.setEnabled(self.reg_automation_active or self.reg_in_progress)
        self.connect_server_button.setEnabled(not self.server_connection_in_progress)
        if self.server_connection_in_progress:
            self.connect_server_button.setText("Connecting...")
        elif self.server_tunnel.is_active():
            self.connect_server_button.setText("Open server")
        else:
            self.connect_server_button.setText("Connect to server")

        if self.reg_in_progress:
            self.reg_button.setText("reg...")
        else:
            self.reg_button.setText("reg")

        if self.reg_automation_active or self.reg_in_progress:
            self.stop_button.setText("Stop")
        else:
            self.stop_button.setText("Stop")

    def persist_imap_settings(self) -> None:
        self.settings.setValue(
            "code_hotkey",
            self.code_hotkey_edit.keySequence().toString(
                QKeySequence.SequenceFormat.PortableText
            ),
        )

    def save_imap_credentials(self, email: str, password: str) -> None:
        self.settings.setValue("imap_email", email.strip())
        self.settings.setValue("imap_password", normalize_zoho_password(password))
        self.update_imap_account_label(email)
        self.refresh_button_state()

    def clear_imap_credentials(self) -> None:
        self.settings.remove("imap_email")
        self.settings.remove("imap_password")
        self.update_imap_account_label("")
        self.code_value_input.clear()
        self.refresh_button_state()

    def stored_imap_email(self) -> str:
        return str(self.settings.value("imap_email", "")).strip()

    def stored_imap_password(self) -> str:
        return normalize_zoho_password(str(self.settings.value("imap_password", "")))

    def update_imap_account_label(self, email: str) -> None:
        normalized = email.strip()
        self.imap_account_label.setText(normalized or "Not signed in")

    def persist_text_bind_settings(self) -> None:
        self.settings.setValue(
            "email_hotkey",
            self.email_hotkey_edit.keySequence().toString(
                QKeySequence.SequenceFormat.PortableText
            ),
        )
        self.settings.remove("text_binds")
        self.settings.beginWriteArray("text_binds", len(self.bind_rows))
        for index, bind_row in enumerate(self.bind_rows):
            self.settings.setArrayIndex(index)
            sequence = bind_row.hotkey_edit.keySequence()
            self.settings.setValue(
                "hotkey",
                sequence.toString(QKeySequence.SequenceFormat.PortableText),
            )
            self.settings.setValue("text", bind_row.text_input.text())
        self.settings.endArray()

    def add_bind_row(
        self, _checked: bool = False, *, hotkey: str = "", text: str = ""
    ) -> None:
        bind_row = KBindRow()
        if hotkey:
            bind_row.hotkey_edit.setKeySequence(QKeySequence(hotkey))
        if text:
            bind_row.text_input.setText(text)
        bind_row.delete_button.clicked.connect(
            lambda _=False, row=bind_row: self.remove_bind_row(row)
        )
        self.bind_rows.append(bind_row)
        self.bind_rows_layout.addWidget(bind_row.widget)

    def remove_bind_row(self, bind_row: KBindRow) -> None:
        if bind_row not in self.bind_rows:
            return
        if bind_row.handle is not None:
            keyboard.remove_hotkey(bind_row.handle)
            bind_row.handle = None
        self.bind_rows.remove(bind_row)
        bind_row.widget.deleteLater()
        self.persist_text_bind_settings()
        try:
            self.register_text_hotkeys()
        except Exception as exc:
            self.unregister_text_hotkeys()
            self.show_error("Failed to register hotkeys", str(exc))
            return
        self.append_log("Removed K-Bind.")

    def unregister_text_hotkeys(self) -> None:
        if self.email_hotkey_handle is not None:
            keyboard.remove_hotkey(self.email_hotkey_handle)
            self.email_hotkey_handle = None
        if self.code_hotkey_handle is not None:
            keyboard.remove_hotkey(self.code_hotkey_handle)
            self.code_hotkey_handle = None
        if self.inspect_window_hotkey_handle is not None:
            keyboard.remove_hotkey(self.inspect_window_hotkey_handle)
            self.inspect_window_hotkey_handle = None
        for bind_row in self.bind_rows:
            if bind_row.handle is not None:
                keyboard.remove_hotkey(bind_row.handle)
                bind_row.handle = None

    def register_text_hotkeys(self) -> None:
        self.unregister_text_hotkeys()
        email_hotkey = key_sequence_to_hotkey(self.email_hotkey_edit.keySequence())
        if email_hotkey and self.current_email():
            self.email_hotkey_handle = keyboard.add_hotkey(
                email_hotkey,
                self.signals.email_copy_requested.emit,
                suppress=True,
            )
        code_hotkey = key_sequence_to_hotkey(self.code_hotkey_edit.keySequence())
        if code_hotkey:
            self.code_hotkey_handle = keyboard.add_hotkey(
                code_hotkey,
                self.signals.code_insert_requested.emit,
                suppress=True,
            )
        self.inspect_window_hotkey_handle = keyboard.add_hotkey(
            "down",
            self.signals.inspect_window_requested.emit,
            suppress=True,
        )
        for bind_row in self.bind_rows:
            hotkey = key_sequence_to_hotkey(bind_row.hotkey_edit.keySequence())
            text = bind_row.text_input.text()
            if not hotkey or not text.strip():
                continue
            bind_row.handle = keyboard.add_hotkey(
                hotkey,
                lambda row=bind_row: self.signals.text_bind_requested.emit(row),
                suppress=True,
            )

    def save_text_binds(self) -> None:
        hotkeys: dict[str, str] = {}
        active_labels: list[str] = []
        email_hotkey = key_sequence_to_hotkey(self.email_hotkey_edit.keySequence())
        code_hotkey = key_sequence_to_hotkey(self.code_hotkey_edit.keySequence())
        if email_hotkey:
            if not self.current_email():
                self.show_error(
                    "Email hotkey is invalid",
                    "Enter an email before saving an email hotkey.",
                )
                return
            hotkeys[email_hotkey] = "email"
            active_labels.append(hotkey_label(self.email_hotkey_edit.keySequence()))

        if code_hotkey:
            if code_hotkey in hotkeys:
                self.show_error(
                    "Duplicate hotkey",
                    "Email hotkey, code hotkey, and K-Bind rows must all use unique hotkeys.",
                )
                return
            hotkeys[code_hotkey] = "code"
            active_labels.append(hotkey_label(self.code_hotkey_edit.keySequence()))

        for bind_row in self.bind_rows:
            hotkey = key_sequence_to_hotkey(bind_row.hotkey_edit.keySequence())
            text = bind_row.text_input.text()
            if bool(hotkey) != bool(text.strip()):
                self.show_error(
                    "Incomplete K-Bind",
                    "Each K-Bind row needs both a hotkey and text, or both fields cleared.",
                )
                return
            if hotkey:
                if hotkey in hotkeys:
                    self.show_error(
                        "Duplicate hotkey",
                        "Email hotkey, code hotkey, and K-Bind rows must all use unique hotkeys.",
                    )
                    return
                hotkeys[hotkey] = "text"
                active_labels.append(hotkey_label(bind_row.hotkey_edit.keySequence()))

        try:
            self.register_text_hotkeys()
        except Exception as exc:
            self.unregister_text_hotkeys()
            self.show_error("Failed to register hotkeys", str(exc))
            return

        self.persist_text_bind_settings()
        self.persist_imap_settings()
        if active_labels:
            self.append_log(f"Saved K-Binds: {', '.join(active_labels)}.")
        else:
            self.append_log("Cleared all K-Binds.")

    def trigger_text_bind(self, bind_row: object) -> None:
        if isinstance(bind_row, KBindRow):
            text = bind_row.text_input.text()
            hotkey_name = hotkey_label(bind_row.hotkey_edit.keySequence()) or "K-Bind"
        else:
            return

        if not text.strip():
            return

        threading.Thread(
            target=self.text_bind_worker,
            args=(text, hotkey_name),
            daemon=True,
        ).start()

    def text_bind_worker(self, text: str, hotkey_name: str) -> None:
        try:
            release_modifier_keys()
            time.sleep(0.05)
            send_unicode_text(text)
        except Exception as exc:
            self.signals.text_bind_error.emit(f"{hotkey_name}: {exc}")
            return

        self.signals.text_bind_success.emit(hotkey_name)

    def handle_text_bind_error(self, message: str) -> None:
        self.append_log(f"Text insert failed: {message}")
        self.show_error("Text insert failed", message)

    def handle_text_bind_success(self, hotkey_name: str) -> None:
        self.append_log(f"Inserted text from {hotkey_name}.")

    def handle_registration_password_step_detected(self) -> None:
        def insert_password_and_submit() -> None:
            self.text_bind_worker(
                require_env("CHATGPT_ACCOUNT_PASSWORD"), "Registration password"
            )
            time.sleep(0.15)
            press_virtual_key(VK_RETURN)

        self.awaiting_registration_code = True
        self.registration_code_baseline = self.current_code()
        self.pending_code_flow = "chatgpt"
        self.pending_code_process_id = self.chatgpt_edge_process_id
        self.start_code_wait("chatgpt", self.chatgpt_edge_process_id)
        threading.Thread(target=insert_password_and_submit, daemon=True).start()
        self.append_log(
            "Detected password step. Inserted password, submitted it, and waiting for a new verification code."
        )

    def handle_registration_email_check_step_detected(self) -> None:
        self.awaiting_registration_code = True
        self.registration_code_baseline = self.current_code()
        self.pending_code_flow = "chatgpt"
        self.pending_code_process_id = self.chatgpt_edge_process_id
        self.start_code_wait("chatgpt", self.chatgpt_edge_process_id)
        self.append_log(
            "Detected email verification step. Waiting for a new mail code and copying it when it arrives."
        )

    def complete_profile_after_code(self) -> None:
        time.sleep(CHATGPT_PROFILE_STEP_DELAY_SECONDS)
        self.text_bind_worker(require_env("CHATGPT_ACCOUNT_NAME"), "Registration name")
        time.sleep(0.2)
        press_virtual_key(VK_TAB)
        time.sleep(0.1)

        try:
            has_yyyy = self.edge_window_contains_text("ГГГГ", timeout_seconds=2.0)
            has_year_2026 = self.edge_window_contains_text("2026", timeout_seconds=2.0)
        except Exception as exc:
            self.signals.imap_error.emit(f"Profile step watcher failed: {exc}")
            return

        if has_yyyy:
            self.text_bind_worker(
                require_env("CHATGPT_ACCOUNT_BIRTH_DAY"), "Registration birth day"
            )
            time.sleep(0.1)
            press_virtual_key(VK_TAB)
            time.sleep(0.1)
            self.text_bind_worker(
                require_env("CHATGPT_ACCOUNT_BIRTH_MONTH"),
                "Registration birth month",
            )
            time.sleep(0.1)
            press_virtual_key(VK_TAB)
            time.sleep(0.1)
            self.text_bind_worker(
                require_env("CHATGPT_ACCOUNT_BIRTH_YEAR"), "Registration birth year"
            )
            self.append_log("Detected YYYY date step. Inserted 13, 04, and 2000.")
        elif has_year_2026:
            press_virtual_key(VK_TAB)
            time.sleep(0.1)
            press_virtual_key(VK_TAB)
            time.sleep(0.1)
            self.text_bind_worker(
                require_env("CHATGPT_ACCOUNT_BIRTH_YEAR"), "Registration birth year"
            )
            self.append_log(
                "Detected year-picker step with 2026. Inserted birth year 2000 after two Tabs."
            )
        else:
            self.text_bind_worker(
                require_env("CHATGPT_ACCOUNT_AGE"), "Registration age"
            )
            self.append_log("Inserted age 26 on the profile step.")

        time.sleep(0.15)
        try:
            matched_name = self.activate_finish_account_button()
        except Exception:
            press_virtual_key(VK_RETURN)
        else:
            self.append_log(f"Activated '{matched_name}' on the profile step.")
        self.append_log("Submitted the profile step.")
        self.signals.registration_post_submit.emit()

    def wait_for_chatgpt_initial_next_step(self, process_id: int) -> str:
        deadline = time.monotonic() + CHATGPT_STARTUP_GUARD_TIMEOUT_SECONDS
        while time.monotonic() < deadline:
            page_text = self.read_edge_window_text(process_id, timeout_seconds=1.0)
            normalized_text = page_text.casefold()
            if any(text in normalized_text for text in CHATGPT_CREATE_PASSWORD_TEXTS):
                return "password"
            if any(text in normalized_text for text in CHATGPT_CHECK_EMAIL_TEXTS):
                return "email_check"
            time.sleep(0.5)
        raise RuntimeError("Timed out waiting for the next ChatGPT registration step.")

    def wait_for_chatgpt_post_code_state(self, process_id: int) -> str:
        deadline = time.monotonic() + CHATGPT_REGISTRATION_WATCH_TIMEOUT_SECONDS
        while time.monotonic() < deadline:
            try:
                page_text = self.read_edge_window_text(
                    process_id,
                    timeout_seconds=CHATGPT_REGISTRATION_POLL_INTERVAL_SECONDS,
                )
            except Exception as exc:
                raise RuntimeError(f"ChatGPT post-code watcher failed: {exc}") from exc

            normalized_text = page_text.casefold()
            if any(text in normalized_text for text in CHATGPT_INVALID_CODE_TEXTS):
                return "invalid_code"
            if not any(text in normalized_text for text in CHATGPT_CHECK_EMAIL_TEXTS):
                return "continue"

            time.sleep(CHATGPT_REGISTRATION_POLL_INTERVAL_SECONDS)

        raise RuntimeError("Timed out waiting for ChatGPT post-code state.")

    def handle_chatgpt_invalid_code(self, process_id: int) -> None:
        self.append_log(
            "ChatGPT: detected invalid code; clearing field and waiting for a new code"
        )
        self.focus_and_clear_chatgpt_code_field(process_id)
        self.awaiting_registration_code = True
        self.registration_code_baseline = self.current_code()
        self.pending_code_flow = "chatgpt"
        self.pending_code_process_id = process_id
        self.start_code_wait("chatgpt", process_id, clear_before_submit=True)

    def focus_and_clear_chatgpt_code_field(self, process_id: int) -> None:
        self.activate_edge_element_by_name(
            process_id,
            "Код",
            timeout_seconds=CHATGPT_BUTTON_TIMEOUT_SECONDS,
            allow_tab_fallback=False,
        )
        release_modifier_keys()
        time.sleep(0.05)
        keyboard.press("ctrl")
        keyboard.press_and_release("a")
        keyboard.release("ctrl")
        time.sleep(0.1)
        for _ in range(6):
            press_virtual_key(VK_BACK)
            time.sleep(0.06)

    def activate_finish_account_button(self) -> str:
        last_error = ""
        for button_text in CHATGPT_FINISH_ACCOUNT_BUTTON_TEXTS:
            try:
                self.activate_edge_element_by_name(
                    0,
                    button_text,
                    timeout_seconds=3.0,
                    allow_tab_fallback=False,
                )
                return button_text
            except Exception as exc:
                last_error = str(exc)

        raise RuntimeError(last_error or "Finish account button was not found.")

    def copy_current_email_to_clipboard(self) -> None:
        email = self.current_email()
        if not email:
            return
        try:
            self.clipboard_text(email)
        except Exception as exc:
            self.show_error("Email copy failed", str(exc))
            return
        hotkey_name = hotkey_label(self.email_hotkey_edit.keySequence()) or "Email"
        self.append_log(f"Copied current email with {hotkey_name}: {email}")

    def insert_current_code(self) -> None:
        code = self.current_code()
        if not code:
            return
        hotkey_name = hotkey_label(self.code_hotkey_edit.keySequence()) or "Code"
        threading.Thread(
            target=self.text_bind_worker,
            args=(code, hotkey_name),
            daemon=True,
        ).start()

    def inspect_current_window(self) -> None:
        if self.retry_hotkey_in_progress:
            return
        threading.Thread(target=self.inspect_current_window_worker, daemon=True).start()

    def inspect_current_window_worker(self) -> None:
        self.retry_hotkey_in_progress = True
        try:
            script = r"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class NativeWindowHelpers
{
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
}
"@

$foreground = [NativeWindowHelpers]::GetForegroundWindow()
if ($foreground -eq [IntPtr]::Zero) {
    Write-Error 'Foreground window was not found.'
    exit 1
}

$root = [System.Windows.Automation.AutomationElement]::FromHandle($foreground)
if (-not $root) {
    Write-Error 'Automation root for foreground window was not found.'
    exit 1
}

Write-Output '=== WINDOW ==='
Write-Output ('Name: ' + $root.Current.Name)
Write-Output ('ClassName: ' + $root.Current.ClassName)
Write-Output ('AutomationId: ' + $root.Current.AutomationId)

$allNodes = $root.FindAll(
    [System.Windows.Automation.TreeScope]::Descendants,
    [System.Windows.Automation.Condition]::TrueCondition
)

Write-Output '=== TEXT NODES ==='
for ($index = 0; $index -lt $allNodes.Count; $index++) {
    $node = $allNodes.Item($index)
    $name = $node.Current.Name
    if (-not [string]::IsNullOrWhiteSpace($name)) {
        try {
            $controlType = $node.Current.ControlType.ProgrammaticName
            $className = $node.Current.ClassName
            $automationId = $node.Current.AutomationId
            $bounds = $node.Current.BoundingRectangle
            Write-Output ("TEXT|Name=" + $name + "|Type=" + $controlType + "|Class=" + $className + "|AutomationId=" + $automationId + "|Top=" + [int]$bounds.Top + "|Left=" + [int]$bounds.Left + "|Width=" + [int]$bounds.Width + "|Height=" + [int]$bounds.Height)
        } catch {
        }
    }
}

$edits = $root.FindAll(
    [System.Windows.Automation.TreeScope]::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Edit
    ))
)

Write-Output '=== EDIT CANDIDATES ==='
if ($edits.Count -eq 0) {
    Write-Output 'No edit fields found.'
} else {
    for ($index = 0; $index -lt $edits.Count; $index++) {
        $edit = $edits.Item($index)
        try {
            $name = $edit.Current.Name
            $className = $edit.Current.ClassName
            $automationId = $edit.Current.AutomationId
            $bounds = $edit.Current.BoundingRectangle
            $hasValuePattern = $false
            try {
                $valuePattern = $null
                $hasValuePattern = $edit.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$valuePattern)
            } catch {
            }
            Write-Output ("EDIT|Index=" + $index + "|Name=" + $name + "|Class=" + $className + "|AutomationId=" + $automationId + "|Top=" + [int]$bounds.Top + "|Left=" + [int]$bounds.Left + "|Width=" + [int]$bounds.Width + "|Height=" + [int]$bounds.Height + "|HasValuePattern=" + $hasValuePattern)
        } catch {
        }
    }
}

Write-Output '=== END ==='
exit 0
"""
            completed = run_powershell_script(script, timeout_seconds=20.0)
            output = ((completed.stdout or "") + (completed.stderr or "")).strip()
            if completed.returncode != 0:
                self.append_log(f"Inspect hotkey failed: {output}")
                return
            try:
                self.clipboard_text(output)
            except Exception as exc:
                self.append_log(
                    f"Inspect hotkey: failed to copy inspection to clipboard: {exc}"
                )
            self.append_log("=== INSPECT OUTPUT START ===")
            for line in output.splitlines():
                self.append_log(line)
            self.append_log("=== INSPECT OUTPUT END ===")
            self.append_log(
                "Inspect hotkey: inspection was written to logs and copied to clipboard if possible."
            )
        finally:
            self.retry_hotkey_in_progress = False

    def focus_first_edit_in_foreground_window(self) -> None:
        script = r"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class NativeWindowHelpers
{
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
}
"@

$foreground = [NativeWindowHelpers]::GetForegroundWindow()
if ($foreground -eq [IntPtr]::Zero) {
    Write-Error 'Foreground window was not found.'
    exit 1
}

$root = [System.Windows.Automation.AutomationElement]::FromHandle($foreground)
if (-not $root) {
    Write-Error 'Automation root for foreground window was not found.'
    exit 1
}

$edits = $root.FindAll(
    [System.Windows.Automation.TreeScope]::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Edit
    ))
)

if ($edits.Count -lt 1) {
    Write-Error 'No edit field was found in the foreground window.'
    exit 1
}

$targetEdit = $null
for ($index = 0; $index -lt $edits.Count; $index++) {
    $candidate = $edits.Item($index)
    try {
        $bounds = $candidate.Current.BoundingRectangle
        if ($bounds.Top -gt 120) {
            $targetEdit = $candidate
            break
        }
    } catch {
    }
}

if (-not $targetEdit) {
    $targetEdit = $edits.Item(0)
}

$targetEdit.SetFocus()
Start-Sleep -Milliseconds 150
Write-Output 'focused'
exit 0
"""
        completed = run_powershell_script(script, timeout_seconds=10.0)
        if completed.returncode != 0:
            raise RuntimeError(
                completed_process_output(completed)
                or "Failed to focus the first edit field in the foreground window."
            )

    def open_chatgpt_in_edge(self) -> None:
        threading.Thread(target=self.open_chatgpt_in_edge_worker, daemon=True).start()

    def handle_reg_button_clicked(self) -> None:
        if self.reg_in_progress:
            return

        self.reg_automation_active = True
        self.reg_in_progress = True
        self.refresh_button_state()
        self.append_log("reg: starting ChatGPT registration flow")
        self.open_chatgpt_in_edge()

    def start_next_reg_cycle(self, *, reason: str) -> None:
        if not self.reg_automation_active:
            return
        if self.reg_in_progress:
            return

        self.reg_in_progress = True
        self.refresh_button_state()
        self.append_log(reason)
        self.open_chatgpt_in_edge()

    def handle_stop_button_clicked(self) -> None:
        self.reg_automation_active = False
        self.reg_in_progress = False
        self.refresh_button_state()
        self.append_log("reg: stop requested; no new cycle will be started")

    def fail_current_iteration(self, reason: str) -> None:
        self.append_log(f"reg: iteration failed: {reason}")
        self.reset_pending_code_state()
        self.omniroute_in_progress = False
        self.server_connection_in_progress = False
        self.reuse_omniroute_browser = False
        self.omniroute_restart_requested = False
        self.code_submit_in_progress = False

        if self.reg_automation_active and self.reg_in_progress:
            self.signals.reg_cycle_finished.emit()
            self.reg_in_progress = False
            self.refresh_button_state()
            QTimer.singleShot(
                0,
                lambda: self.start_next_reg_cycle(
                    reason=f"reg: continuing after failure ({reason})"
                ),
            )
            return

        self.reg_in_progress = False
        self.refresh_button_state()

    def handle_server_button_clicked(self) -> None:
        if self.server_connection_in_progress:
            return

        if self.server_tunnel.is_active():
            try:
                browser_url = self.server_browser_url()
            except Exception as exc:
                self.show_error("Server URL is invalid", str(exc))
                return
            webbrowser.open_new_tab(browser_url)
            self.append_log(f"Opened server in browser: {browser_url}")
            return

        self.server_connection_in_progress = True
        self.refresh_button_state()
        threading.Thread(target=self.connect_server_worker, daemon=True).start()

    def handle_omniroute_button_clicked(self) -> None:
        if self.omniroute_in_progress:
            return

        self.omniroute_in_progress = True
        self.refresh_button_state()
        threading.Thread(target=self.run_omniroute_worker, daemon=True).start()

    def handle_registration_flow_completed(self) -> None:
        self.append_log("reg: ChatGPT flow completed, starting Omniroute")
        if self.omniroute_in_progress:
            return
        self.omniroute_in_progress = True
        self.refresh_button_state()
        threading.Thread(target=self.run_omniroute_worker, daemon=True).start()

    def handle_registration_post_submit(self) -> None:
        self.append_log("reg: waiting 2 seconds before closing ChatGPT Edge")
        QTimer.singleShot(2000, self.finish_registration_and_start_omniroute)

    def finish_registration_and_start_omniroute(self) -> None:
        try:
            self.close_chatgpt_edge_window()
        except Exception as exc:
            self.signals.imap_error.emit(f"Failed to close ChatGPT Edge window: {exc}")
            self.fail_current_iteration(f"Failed to close ChatGPT Edge window: {exc}")
            return

        self.signals.registration_flow_completed.emit()

    def close_chatgpt_edge_window(self) -> None:
        script = """
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$matches = Get-Process msedge -ErrorAction SilentlyContinue |
    Where-Object {
        $_.MainWindowHandle -ne 0 -and
        -not [string]::IsNullOrWhiteSpace($_.MainWindowTitle) -and
        $_.MainWindowTitle -match 'OpenAI'
    }

if (-not $matches) {
    Write-Error 'OpenAI Edge window was not found.'
    exit 1
}

foreach ($process in $matches) {
    [void]$process.CloseMainWindow()
}

Start-Sleep -Seconds 2

$stillOpen = Get-Process msedge -ErrorAction SilentlyContinue |
    Where-Object {
        $_.MainWindowHandle -ne 0 -and
        -not [string]::IsNullOrWhiteSpace($_.MainWindowTitle) -and
        $_.MainWindowTitle -match 'OpenAI'
    }

if ($stillOpen) {
    Write-Error 'OpenAI Edge window is still open after close request.'
    exit 1
}

Write-Output 'closed'
exit 0
"""

        try:
            completed = run_powershell_script(script, timeout_seconds=15.0)
            if completed.returncode != 0:
                raise RuntimeError(
                    completed_process_output(completed)
                    or "Failed to close OpenAI Edge window."
                )
        except Exception as exc:
            self.append_log(f"Failed to close ChatGPT Edge window: {exc}")
            raise
        else:
            self.append_log("Closed ChatGPT Edge window.")
        finally:
            self.chatgpt_edge_process_id = 0

    def handle_reg_cycle_finished(self) -> None:
        self.increment_email()
        self.append_log("reg: incremented alias for the next run")

    def server_browser_url(self) -> str:
        browser_url = os.getenv("CODEX_SERVER_BROWSER_URL", "").strip()
        if not browser_url:
            raise RuntimeError(
                "Missing CODEX_SERVER_BROWSER_URL in .env. Use .env.example as a template."
            )
        return browser_url

    def omniroute_url(self) -> str:
        return f"{self.server_browser_url().rstrip('/')}{OMNIROUTE_PAGE_PATH}"

    def server_connection_config(
        self,
    ) -> tuple[str, int, str, str, list[tuple[int, str, int]]]:
        host = os.getenv("CODEX_SERVER_SSH_HOST", "").strip()
        username = os.getenv("CODEX_SERVER_SSH_USER", "").strip()
        password = os.getenv("CODEX_SERVER_SSH_PASSWORD", "")
        if not host or not username or not password:
            raise RuntimeError(
                "Missing SSH settings in .env. Fill CODEX_SERVER_SSH_HOST, CODEX_SERVER_SSH_USER, and CODEX_SERVER_SSH_PASSWORD."
            )

        ssh_port_raw = os.getenv("CODEX_SERVER_SSH_PORT", "22").strip() or "22"
        first_local_port_raw = os.getenv("CODEX_SERVER_TUNNEL_1_LOCAL_PORT", "").strip()
        first_remote_host = (
            os.getenv("CODEX_SERVER_TUNNEL_1_REMOTE_HOST", "127.0.0.1").strip()
            or "127.0.0.1"
        )
        first_remote_port_raw = os.getenv(
            "CODEX_SERVER_TUNNEL_1_REMOTE_PORT", ""
        ).strip()
        second_local_port_raw = os.getenv(
            "CODEX_SERVER_TUNNEL_2_LOCAL_PORT", ""
        ).strip()
        second_remote_host = (
            os.getenv("CODEX_SERVER_TUNNEL_2_REMOTE_HOST", "127.0.0.1").strip()
            or "127.0.0.1"
        )
        second_remote_port_raw = os.getenv(
            "CODEX_SERVER_TUNNEL_2_REMOTE_PORT", ""
        ).strip()

        required_port_values = {
            "CODEX_SERVER_TUNNEL_1_LOCAL_PORT": first_local_port_raw,
            "CODEX_SERVER_TUNNEL_1_REMOTE_PORT": first_remote_port_raw,
            "CODEX_SERVER_TUNNEL_2_LOCAL_PORT": second_local_port_raw,
            "CODEX_SERVER_TUNNEL_2_REMOTE_PORT": second_remote_port_raw,
        }
        missing_port_keys = [
            key for key, value in required_port_values.items() if not value.strip()
        ]
        if missing_port_keys:
            raise RuntimeError(
                f"Missing tunnel settings in .env: {', '.join(missing_port_keys)}"
            )

        try:
            ssh_port = int(ssh_port_raw)
            first_local_port = int(first_local_port_raw)
            first_remote_port = int(first_remote_port_raw)
            second_local_port = int(second_local_port_raw)
            second_remote_port = int(second_remote_port_raw)
        except ValueError as exc:
            raise RuntimeError("SSH ports in .env must be valid integers.") from exc

        forwards = [
            (first_local_port, first_remote_host, first_remote_port),
            (second_local_port, second_remote_host, second_remote_port),
        ]
        return host, ssh_port, username, password, forwards

    def connect_server_worker(self) -> None:
        try:
            host, ssh_port, username, password, forwards = (
                self.server_connection_config()
            )
            browser_url = self.server_browser_url()
            self.server_tunnel.start(
                host=host,
                port=ssh_port,
                username=username,
                password=password,
                forwards=forwards,
            )
        except Exception as exc:
            self.signals.server_connection_failed.emit(str(exc))
            return

        self.signals.server_connected.emit(browser_url)

    def ensure_server_connected(self) -> None:
        host, ssh_port, username, password, forwards = self.server_connection_config()
        self.server_tunnel.start(
            host=host,
            port=ssh_port,
            username=username,
            password=password,
            forwards=forwards,
        )

    def reset_pending_code_state(self) -> None:
        self.awaiting_registration_code = False
        self.registration_code_baseline = ""
        self.pending_code_flow = ""
        self.pending_code_process_id = 0
        self.code_submit_in_progress = False
        self.code_wait_generation += 1
        self.code_wait_flow = ""
        self.code_wait_process_id = 0
        self.code_wait_attempts = 0
        self.code_wait_refocus_needed = False
        self.code_wait_clear_before_submit = False

    def start_code_wait(
        self,
        flow: str,
        process_id: int,
        *,
        clear_before_submit: bool = False,
        reset_attempts: bool = True,
    ) -> None:
        self.code_wait_generation += 1
        generation = self.code_wait_generation
        self.code_wait_flow = flow
        self.code_wait_process_id = process_id
        if reset_attempts:
            self.code_wait_attempts = 0
        self.code_wait_refocus_needed = False
        self.code_wait_clear_before_submit = clear_before_submit

        def wait_worker() -> None:
            while (
                generation == self.code_wait_generation
                and self.awaiting_registration_code
            ):
                time.sleep(CODE_WAIT_TIMEOUT_SECONDS)
                if (
                    generation != self.code_wait_generation
                    or not self.awaiting_registration_code
                ):
                    return
                self.signals.code_wait_timeout.emit(flow, generation)
                return

        threading.Thread(target=wait_worker, daemon=True).start()

    def handle_code_wait_timeout(self, flow: str, generation: int) -> None:
        if generation != self.code_wait_generation:
            return
        if not self.awaiting_registration_code:
            return
        if flow != self.code_wait_flow:
            return

        self.code_wait_attempts += 1
        current_attempt = self.code_wait_attempts
        process_id = self.code_wait_process_id
        self.append_log(
            f"{flow}: no code received after {int(CODE_WAIT_TIMEOUT_SECONDS)} seconds (attempt {current_attempt}/{MAX_CODE_RESEND_ATTEMPTS})"
        )

        try:
            current_page_state = self.inspect_current_code_wait_page(flow, process_id)
        except Exception as exc:
            self.append_log(f"{flow}: failed to inspect current code-wait page: {exc}")
            current_page_state = "unknown"

        if current_page_state == "reject":
            self.awaiting_registration_code = False
            self.append_log(f"{flow}: page changed to reject state before resend.")
            if flow == "omniroute":
                try:
                    self.handle_omniroute_reject(process_id)
                except Exception as exc:
                    self.signals.omniroute_failed.emit(str(exc))
                    return
                self.signals.omniroute_finished.emit()
            return

        if current_page_state == "continue":
            self.awaiting_registration_code = False
            self.append_log(
                f"{flow}: page already reached continue state before resend."
            )
            if flow == "omniroute":
                self.signals.omniroute_finished.emit()
            return

        if current_page_state == "other":
            self.awaiting_registration_code = False
            self.append_log(
                f"{flow}: page left the code step before resend. Continuing without another wait cycle."
            )
            if flow == "omniroute":
                self.signals.omniroute_failed.emit(
                    "OmniRoute left the code step unexpectedly."
                )
            return

        if current_attempt <= MAX_CODE_RESEND_ATTEMPTS:
            try:
                self.click_resend_email_button(flow, process_id)
            except Exception as exc:
                self.append_log(f"{flow}: resend click failed: {exc}")
            else:
                self.append_log(f"{flow}: requested a new code by clicking resend")
            self.code_wait_refocus_needed = True
            self.start_code_wait(flow, process_id, reset_attempts=False)
            self.code_wait_refocus_needed = True
            return

        self.awaiting_registration_code = False
        self.append_log(
            f"{flow}: no code received after resend limit; continuing with failover"
        )
        if flow == "omniroute":
            try:
                self.handle_omniroute_reject(process_id)
            except Exception as exc:
                self.signals.omniroute_failed.emit(str(exc))
                return
            self.signals.omniroute_finished.emit()
            return

        if flow == "chatgpt":
            try:
                self.close_chatgpt_edge_window()
            except Exception as exc:
                self.append_log(
                    f"chatgpt: failed to close Edge after resend limit: {exc}"
                )
            if self.reg_automation_active:
                self.fail_current_iteration(
                    "ChatGPT did not receive a code after resend limit"
                )
            return

    def click_resend_email_button(self, flow: str, process_id: int) -> None:
        if flow == "chatgpt":
            self.activate_edge_element_by_name(
                process_id,
                "Отправить электронное письмо повторно",
                timeout_seconds=CHATGPT_BUTTON_TIMEOUT_SECONDS,
                allow_tab_fallback=False,
            )
            return

        process_filter = "$window = Get-Process browser -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Sort-Object StartTime -Descending | Select-Object -First 1"
        not_found_message = "Yandex browser window was not found for resend."

        script = """
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

{process_filter}

if (-not $window) {{
    Write-Error '{not_found_message}'
    exit 1
}}

[void][Microsoft.VisualBasic.Interaction]::AppActivate($window.Id)
Start-Sleep -Milliseconds 300

$root = [System.Windows.Automation.AutomationElement]::FromHandle($window.MainWindowHandle)
if (-not $root) {
    Write-Error 'Automation root was not found for resend.'
    exit 1
}

$nodes = $root.FindAll(
    [System.Windows.Automation.TreeScope]::Descendants,
    [System.Windows.Automation.Condition]::TrueCondition
)

for ($index = 0; $index -lt $nodes.Count; $index++) {
    $node = $nodes.Item($index)
    $name = $node.Current.Name
    if (-not [string]::IsNullOrWhiteSpace($name) -and $name.ToLower().Contains('повторно')) {
        try {
            $node.SetFocus()
            Start-Sleep -Milliseconds 150
            [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
            Write-Output "clicked:$name"
            exit 0
        } catch {
        }
    }
}

Write-Error 'Resend button containing "повторно" was not found.'
exit 1
""".format(process_filter=process_filter, not_found_message=not_found_message)
        completed = run_powershell_script(script, timeout_seconds=15.0)
        if completed.returncode != 0:
            raise RuntimeError(
                completed_process_output(completed) or "Failed to click resend button."
            )

    def prepare_omniroute_code_wait(
        self, process_id: int, *, clear_before_submit: bool = False
    ) -> None:
        self.awaiting_registration_code = True
        self.registration_code_baseline = self.current_code()
        self.pending_code_flow = "omniroute"
        self.pending_code_process_id = process_id
        self.start_code_wait(
            "omniroute", process_id, clear_before_submit=clear_before_submit
        )
        self.signals.omniroute_status.emit("Omniroute: ожидаем новый код из почты")
        self.signals.omniroute_status.emit(
            "Omniroute is waiting for a new mail code and will submit it automatically."
        )

    def open_omniroute_in_browser(self) -> int:
        browser_executable = self.resolve_yandex_browser_executable()
        omniroute_process = subprocess.Popen(
            [browser_executable, "--start-maximized", self.omniroute_url()],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        time.sleep(OMNIROUTE_PAGE_LOAD_DELAY_SECONDS)
        return omniroute_process.pid

    def reuse_omniroute_browser_window(self) -> int:
        script = """
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$browser = Get-Process browser -ErrorAction SilentlyContinue |
    Where-Object {
        $_.MainWindowHandle -ne 0 -and
        -not [string]::IsNullOrWhiteSpace($_.MainWindowTitle) -and
        $_.MainWindowTitle -match 'OmniRoute'
    } |
    Sort-Object StartTime -Descending |
    Select-Object -First 1

if (-not $browser) {
    Write-Error 'Active OmniRoute browser window was not found.'
    exit 1
}

$shell = New-Object -ComObject WScript.Shell
[void]$shell.AppActivate($browser.Id)
Start-Sleep -Milliseconds 300
Write-Output $browser.Id
exit 0
"""
        completed = run_powershell_script(script, timeout_seconds=10.0)
        if completed.returncode != 0:
            raise RuntimeError(
                completed_process_output(completed)
                or "Failed to activate existing OmniRoute browser window."
            )

        output = (completed.stdout or "").strip()
        try:
            return int(output)
        except ValueError as exc:
            raise RuntimeError("Failed to read OmniRoute browser process id.") from exc

    def activate_edge_element_by_names(
        self,
        process_id: int,
        element_names: tuple[str, ...],
        *,
        timeout_seconds: float,
        allow_tab_fallback: bool = True,
    ) -> str:
        last_error = ""
        per_name_timeout = max(2.0, timeout_seconds / max(1, len(element_names)))
        for element_name in element_names:
            try:
                self.activate_edge_element_by_name(
                    process_id,
                    element_name,
                    timeout_seconds=per_name_timeout,
                    allow_tab_fallback=allow_tab_fallback,
                )
                return element_name
            except Exception as exc:
                last_error = str(exc)

        raise RuntimeError(last_error or "Failed to activate a matching UI element.")

    def detect_omniroute_step(
        self,
        process_id: int,
        *,
        timeout_seconds: float,
        allow_continue: bool = True,
    ) -> str:
        deadline = time.monotonic() + timeout_seconds
        last_debug_signature = ""

        while time.monotonic() < deadline:
            try:
                page_text = self.read_edge_window_text(
                    process_id,
                    timeout_seconds=CHATGPT_REGISTRATION_POLL_INTERVAL_SECONDS,
                )
            except Exception as exc:
                raise RuntimeError(f"Omniroute watcher failed: {exc}") from exc

            if page_text:
                debug_signature = page_text[:300]
                if debug_signature != last_debug_signature:
                    last_debug_signature = debug_signature
                    preview = page_text.replace("\n", " | ")[:300]
                    self.signals.omniroute_status.emit(
                        f"Omniroute watcher sees: {preview}"
                    )

            normalized_text = page_text.casefold()
            if any(text in normalized_text for text in OMNIROUTE_PASSWORD_STEP_TEXTS):
                return "password"
            if any(text in normalized_text for text in OMNIROUTE_CHECK_EMAIL_TEXTS):
                return "email_check"
            if any(text in normalized_text for text in OMNIROUTE_REJECT_TEXTS):
                return "reject"
            if allow_continue and any(
                text.casefold() in normalized_text
                for text in OMNIROUTE_CONTINUE_BUTTON_TEXTS
            ):
                return "continue"

            time.sleep(CHATGPT_REGISTRATION_POLL_INTERVAL_SECONDS)

        raise RuntimeError("Timed out waiting for the Omniroute page state.")

    def detect_omniroute_post_password_step(self, process_id: int) -> str:
        deadline = time.monotonic() + OMNIROUTE_POST_PASSWORD_TIMEOUT_SECONDS
        seen_password_state = False

        while time.monotonic() < deadline:
            page_text = self.read_edge_window_text(process_id, timeout_seconds=1.0)
            normalized_text = page_text.casefold()

            if any(text in normalized_text for text in OMNIROUTE_PASSWORD_STEP_TEXTS):
                seen_password_state = True
                time.sleep(0.5)
                continue

            if seen_password_state:
                if any(text in normalized_text for text in OMNIROUTE_CHECK_EMAIL_TEXTS):
                    return "email_check"
                if any(text in normalized_text for text in OMNIROUTE_REJECT_TEXTS):
                    return "reject"
                if any(
                    text.casefold() in normalized_text
                    for text in OMNIROUTE_CONTINUE_BUTTON_TEXTS
                ) and not any(
                    text in normalized_text for text in OMNIROUTE_CHECK_EMAIL_TEXTS
                ):
                    return "continue"

            time.sleep(0.5)

        raise RuntimeError("Timed out waiting for OmniRoute post-password transition.")

    def inspect_current_code_wait_page(self, flow: str, process_id: int) -> str:
        page_text = self.read_edge_window_text(process_id, timeout_seconds=2.0)
        normalized_text = page_text.casefold()

        if flow == "omniroute":
            if any(text in normalized_text for text in OMNIROUTE_INVALID_CODE_TEXTS):
                return "invalid_code"
            if any(text in normalized_text for text in OMNIROUTE_CHECK_EMAIL_TEXTS):
                return "email_check"
            if any(text in normalized_text for text in OMNIROUTE_REJECT_TEXTS):
                return "reject"
            if any(
                text.casefold() in normalized_text
                for text in OMNIROUTE_CONTINUE_BUTTON_TEXTS
            ) and not any(
                text in normalized_text for text in OMNIROUTE_CHECK_EMAIL_TEXTS
            ):
                return "continue"
            return "other"

        if any(text in normalized_text for text in CHATGPT_INVALID_CODE_TEXTS):
            return "invalid_code"
        if any(text in normalized_text for text in CHATGPT_CHECK_EMAIL_TEXTS):
            return "email_check"
        return "other"

    def submit_omniroute_password(self, process_id: int) -> None:
        self.text_bind_worker(
            require_env("CHATGPT_ACCOUNT_PASSWORD"), "Omniroute password"
        )
        time.sleep(0.15)
        press_virtual_key(VK_RETURN)
        self.signals.omniroute_status.emit("Omniroute: вставлен пароль")
        self.signals.omniroute_status.emit(
            "Omniroute watcher matched the password step and submitted the password."
        )

    def complete_omniroute_after_code(self, process_id: int) -> None:
        self.signals.omniroute_status.emit("Omniroute: вставлен код")
        matched_name = self.wait_for_omniroute_continue(process_id)
        if matched_name == "reject":
            self.handle_omniroute_reject(process_id)
            return
        if matched_name == "invalid_code":
            self.handle_omniroute_invalid_code(process_id)
            return
        self.signals.omniroute_status.emit(
            'Omniroute: найден финальный шаг "Продолжить"'
        )
        matched_name = self.activate_edge_element_by_names(
            process_id,
            OMNIROUTE_CONTINUE_BUTTON_TEXTS,
            timeout_seconds=CHATGPT_BUTTON_TIMEOUT_SECONDS,
            allow_tab_fallback=False,
        )
        self.signals.omniroute_status.emit('Omniroute: нажата "Продолжить"')
        self.signals.omniroute_status.emit(
            f"Omniroute activated '{matched_name}' after code entry."
        )

    def wait_for_omniroute_continue(self, process_id: int) -> str:
        deadline = time.monotonic() + OMNIROUTE_STEP_TIMEOUT_SECONDS
        last_debug_signature = ""

        while time.monotonic() < deadline:
            try:
                page_text = self.read_edge_window_text(
                    process_id,
                    timeout_seconds=CHATGPT_REGISTRATION_POLL_INTERVAL_SECONDS,
                )
            except Exception as exc:
                raise RuntimeError(f"Omniroute final watcher failed: {exc}") from exc

            if page_text:
                debug_signature = page_text[:300]
                if debug_signature != last_debug_signature:
                    last_debug_signature = debug_signature
                    preview = page_text.replace("\n", " | ")[:300]
                    self.signals.omniroute_status.emit(
                        f"Omniroute final watcher sees: {preview}"
                    )

            normalized_text = page_text.casefold()
            if any(text in normalized_text for text in OMNIROUTE_INVALID_CODE_TEXTS):
                return "invalid_code"
            if any(text in normalized_text for text in OMNIROUTE_REJECT_TEXTS):
                return "reject"
            if any(
                text.casefold() in normalized_text
                for text in OMNIROUTE_CONTINUE_BUTTON_TEXTS
            ) and not any(
                text in normalized_text for text in OMNIROUTE_CHECK_EMAIL_TEXTS
            ):
                return "continue"

            time.sleep(CHATGPT_REGISTRATION_POLL_INTERVAL_SECONDS)

        raise RuntimeError("Timed out waiting for the final continue step.")

    def handle_omniroute_invalid_code(self, process_id: int) -> None:
        self.signals.omniroute_status.emit(
            "Omniroute: detected invalid code; clearing field and waiting for a new code"
        )
        self.focus_and_clear_active_code_field("omniroute")
        self.prepare_omniroute_code_wait(process_id, clear_before_submit=True)

    def build_edge_process_filter(self, process_id: int) -> str:
        return f"$window = Get-Process -Id {process_id} -ErrorAction SilentlyContinue | Where-Object {{ $_.MainWindowHandle -ne 0 }} | Select-Object -First 1"

    def wait_for_omniroute_provider_page_ready(self, process_id: int) -> None:
        deadline = time.monotonic() + 15.0
        while time.monotonic() < deadline:
            page_text = self.read_edge_window_text(process_id, timeout_seconds=1.0)
            normalized_text = page_text.casefold()
            if "openai codex" in normalized_text or "connections" in normalized_text:
                self.signals.omniroute_status.emit(
                    "Omniroute: provider page is ready for Add search"
                )
                return
            time.sleep(0.2)
        raise RuntimeError("Timed out waiting for OmniRoute provider page readiness.")

    def focus_and_clear_active_code_field(self, flow: str) -> None:
        if flow == "chatgpt":
            process_filter = self.build_edge_process_filter(self.code_wait_process_id)
            not_found_message = (
                "Edge window was not found while clearing the code field."
            )
        else:
            process_filter = "$window = Get-Process browser -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Sort-Object StartTime -Descending | Select-Object -First 1"
            not_found_message = (
                "Yandex browser window was not found while clearing the code field."
            )

        script = """
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

{process_filter}

if (-not $window) {{
    Write-Error '{not_found_message}'
    exit 1
}}

[void][Microsoft.VisualBasic.Interaction]::AppActivate($window.Id)
Start-Sleep -Milliseconds 300

$root = [System.Windows.Automation.AutomationElement]::FromHandle($window.MainWindowHandle)
if (-not $root) {
    Write-Error 'Automation root was not found while clearing the code field.'
    exit 1
}

$edits = $root.FindAll(
    [System.Windows.Automation.TreeScope]::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Edit
    ))
)

if ($edits.Count -lt 1) {
    Write-Error 'Edit field was not found while clearing the code field.'
    exit 1
}

$targetEdit = $edits.Item(0)
try {
    $targetEdit.SetFocus()
    Start-Sleep -Milliseconds 150
} catch {
    Write-Error 'Failed to focus the code field.'
    exit 1
}

try {
    $valuePattern = $null
    if ($targetEdit.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$valuePattern)) {
        $valuePattern.SetValue('')
        Start-Sleep -Milliseconds 150
        Write-Output 'cleared:value-pattern'
        exit 0
    }
} catch {
}

[System.Windows.Forms.SendKeys]::SendWait('{{END}}')
Start-Sleep -Milliseconds 100
for ($index = 0; $index -lt 6; $index++) {{
    [System.Windows.Forms.SendKeys]::SendWait('{{BACKSPACE}}')
    Start-Sleep -Milliseconds 60
}}

Write-Output 'cleared:fallback'
exit 0
""".format(process_filter=process_filter, not_found_message=not_found_message)
        completed = run_powershell_script(script, timeout_seconds=15.0)
        if completed.returncode != 0:
            raise RuntimeError(
                completed_process_output(completed)
                or "Failed to focus and clear the active code field."
            )

    def focus_active_code_field(self, flow: str) -> None:
        if flow == "chatgpt":
            if self.code_wait_process_id:
                process_filter = self.build_edge_process_filter(
                    self.code_wait_process_id
                )
            else:
                process_filter = "$window = Get-Process msedge -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 -and -not [string]::IsNullOrWhiteSpace($_.MainWindowTitle) -and $_.MainWindowTitle -match 'OpenAI' } | Sort-Object StartTime -Descending | Select-Object -First 1"
            not_found_message = (
                "Edge window was not found while focusing the code field."
            )
        else:
            process_filter = "$window = Get-Process browser -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Sort-Object StartTime -Descending | Select-Object -First 1"
            not_found_message = (
                "Browser window was not found while focusing the code field."
            )

        script = """
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName Microsoft.VisualBasic

{process_filter}

if (-not $window) {{
    Write-Error '{not_found_message}'
    exit 1
}}

[void][Microsoft.VisualBasic.Interaction]::AppActivate($window.Id)
Start-Sleep -Milliseconds 300

$root = [System.Windows.Automation.AutomationElement]::FromHandle($window.MainWindowHandle)
if (-not $root) {
    Write-Error 'Automation root was not found while focusing the code field.'
    exit 1
}

$edits = $root.FindAll(
    [System.Windows.Automation.TreeScope]::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Edit
    ))
)

if ($edits.Count -lt 1) {
    Write-Error 'Edit field was not found while focusing the code field.'
    exit 1
}

$targetEdit = $null
for ($index = 0; $index -lt $edits.Count; $index++) {{
    $candidate = $edits.Item($index)
    try {{
        $name = $candidate.Current.Name
        $automationId = $candidate.Current.AutomationId
        if ($name -eq 'Код' -or $automationId.EndsWith('-code')) {{
            $targetEdit = $candidate
            break
        }}
    }} catch {{
    }}
}}

if (-not $targetEdit) {{
    $targetEdit = $edits.Item(0)
}}

$targetEdit.SetFocus()
Start-Sleep -Milliseconds 150
Write-Output 'focused'
exit 0
""".format(process_filter=process_filter, not_found_message=not_found_message)
        completed = run_powershell_script(script, timeout_seconds=15.0)
        if completed.returncode != 0:
            raise RuntimeError(
                completed_process_output(completed)
                or "Failed to focus the active code field."
            )

    def handle_omniroute_reject(self, process_id: int) -> None:
        self.signals.omniroute_status.emit(
            "Omniroute: detected age/phone verification requirement; restarting cycle"
        )
        self.omniroute_restart_requested = True
        self.try_close_active_yandex_tab(
            "Yandex provider tab",
            expect_title_fragment="OpenAI",
        )
        self.try_close_active_yandex_tab(
            "OmniRoute base tab",
            expect_title_fragment="OmniRoute",
        )
        self.reuse_omniroute_browser = False

    def try_close_active_yandex_tab(
        self, label: str, *, expect_title_fragment: str
    ) -> None:
        try:
            self.close_active_yandex_tab(
                label, expect_title_fragment=expect_title_fragment
            )
        except Exception as exc:
            self.signals.omniroute_status.emit(
                f"Omniroute: failed to close {label}: {exc}. Continuing with a fresh cycle."
            )

    def close_active_yandex_tab(
        self, label: str, *, expect_title_fragment: str
    ) -> None:
        expected_title = expect_title_fragment.replace("'", "''")
        script = f"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class NativeKeyboard
{{
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
}}
"@
$target = Get-Process browser -ErrorAction SilentlyContinue |
    Where-Object {{ $_.MainWindowHandle -ne 0 }} |
    Sort-Object StartTime -Descending |
    Select-Object -First 1

if (-not $target) {{
    Write-Error 'Active Yandex browser window was not found.'
    exit 1
}}

$beforeTitle = $target.MainWindowTitle
if ([string]::IsNullOrWhiteSpace($beforeTitle) -or $beforeTitle -notmatch '{expected_title}') {{
    Write-Error "Unexpected active Yandex tab title: $beforeTitle"
    exit 1
}}

[void][Microsoft.VisualBasic.Interaction]::AppActivate($target.Id)
Start-Sleep -Milliseconds 300
Add-Type -AssemblyName System.Windows.Forms
[NativeKeyboard]::keybd_event(0x11, 0, 0, 0)
Start-Sleep -Milliseconds 50
[NativeKeyboard]::keybd_event(0x57, 0, 0, 0)
Start-Sleep -Milliseconds 50
[NativeKeyboard]::keybd_event(0x57, 0, 2, 0)
Start-Sleep -Milliseconds 50
[NativeKeyboard]::keybd_event(0x11, 0, 2, 0)
Start-Sleep -Seconds 2

$after = Get-Process browser -ErrorAction SilentlyContinue |
    Where-Object {{ $_.MainWindowHandle -ne 0 }} |
    Sort-Object StartTime -Descending |
    Select-Object -First 1

if ($after -and $after.MainWindowTitle -eq $beforeTitle) {{
    Write-Error 'Yandex tab title did not change after Ctrl+W.'
    exit 1
}}

Write-Output 'closed-tab'
exit 0
"""
        completed = run_powershell_script(script, timeout_seconds=10.0)
        if completed.returncode != 0:
            raise RuntimeError(
                completed_process_output(completed) or f"Failed to close {label}."
            )
        self.signals.omniroute_status.emit(f"Omniroute: closed {label}")

    def run_omniroute_worker(self) -> None:
        try:
            self.signals.omniroute_status.emit("Omniroute: запуск команды")
            email = self.current_email()
            if not email:
                raise ValueError("Enter an email before starting Omniroute.")

            self.ensure_server_connected()
            if self.reuse_omniroute_browser and self.reg_automation_active:
                self.signals.omniroute_status.emit(
                    "Omniroute: reusing existing browser tab"
                )
                try:
                    process_id = self.reuse_omniroute_browser_window()
                except Exception as exc:
                    self.reuse_omniroute_browser = False
                    self.signals.omniroute_status.emit(
                        f"Omniroute: reuse failed ({exc}). Opening a fresh OmniRoute tab."
                    )
                    process_id = self.open_omniroute_in_browser()
            else:
                process_id = self.open_omniroute_in_browser()
                try:
                    self.wait_for_omniroute_provider_page_ready(process_id)
                except Exception as exc:
                    self.signals.omniroute_status.emit(
                        f"Omniroute: provider readiness pre-check failed ({exc}). Reopening a fresh OmniRoute tab before Add search."
                    )
                    process_id = self.open_omniroute_in_browser()
                    try:
                        self.wait_for_omniroute_provider_page_ready(process_id)
                    except Exception:
                        pass
            self.signals.omniroute_status.emit("Omniroute: начался поиск Add")
            try:
                matched_add_button = self.activate_edge_element_by_names(
                    process_id,
                    OMNIROUTE_ADD_BUTTON_TEXTS,
                    timeout_seconds=OMNIROUTE_ADD_TIMEOUT_SECONDS,
                    allow_tab_fallback=False,
                )
            except Exception:
                if self.reg_automation_active and not self.reuse_omniroute_browser:
                    self.signals.omniroute_status.emit(
                        "Omniroute: first Add attempt failed, reopening a fresh tab and retrying once."
                    )
                    process_id = self.open_omniroute_in_browser()
                    try:
                        self.wait_for_omniroute_provider_page_ready(process_id)
                    except Exception:
                        pass
                    matched_add_button = self.activate_edge_element_by_names(
                        process_id,
                        OMNIROUTE_ADD_BUTTON_TEXTS,
                        timeout_seconds=OMNIROUTE_ADD_TIMEOUT_SECONDS,
                        allow_tab_fallback=False,
                    )
                else:
                    raise
            self.signals.omniroute_status.emit(
                f"Omniroute: найдена кнопка {matched_add_button}"
            )
            self.signals.omniroute_status.emit(
                f"Omniroute activated '{matched_add_button}' on the provider page."
            )
            self.signals.omniroute_status.emit(
                f"Omniroute: нажата кнопка {matched_add_button}"
            )
            time.sleep(5.0)
            release_modifier_keys()
            send_unicode_text(email)
            press_virtual_key(VK_RETURN)
            self.signals.omniroute_status.emit("Omniroute: вставлена почта")
            self.signals.omniroute_status.emit(
                "Omniroute inserted the current email and submitted it."
            )

            current_step = self.detect_omniroute_step(
                process_id,
                timeout_seconds=OMNIROUTE_STEP_TIMEOUT_SECONDS,
                allow_continue=False,
            )
            if current_step == "password":
                self.signals.omniroute_status.emit("Omniroute: найден шаг пароля")
                self.submit_omniroute_password(process_id)
                next_step = self.detect_omniroute_post_password_step(process_id)
                if next_step == "email_check":
                    self.prepare_omniroute_code_wait(process_id)
                    return
                if next_step == "reject":
                    self.handle_omniroute_reject(process_id)
                    return
                if next_step == "continue":
                    matched_continue_button = self.activate_edge_element_by_names(
                        process_id,
                        OMNIROUTE_CONTINUE_BUTTON_TEXTS,
                        timeout_seconds=CHATGPT_BUTTON_TIMEOUT_SECONDS,
                        allow_tab_fallback=False,
                    )
                    self.signals.omniroute_status.emit(
                        f"Omniroute activated '{matched_continue_button}'."
                    )
                    return
                raise RuntimeError(
                    f"Unsupported OmniRoute state after password submit: {next_step}"
                )

            if current_step == "email_check":
                self.signals.omniroute_status.emit("Omniroute: найден шаг кода")
                self.prepare_omniroute_code_wait(process_id)
                return

            if current_step == "reject":
                self.handle_omniroute_reject(process_id)
                return

            if current_step == "continue":
                matched_continue_button = self.activate_edge_element_by_names(
                    process_id,
                    OMNIROUTE_CONTINUE_BUTTON_TEXTS,
                    timeout_seconds=CHATGPT_BUTTON_TIMEOUT_SECONDS,
                    allow_tab_fallback=False,
                )
                self.signals.omniroute_status.emit(
                    f"Omniroute activated '{matched_continue_button}'."
                )
                return

            raise RuntimeError(f"Unsupported Omniroute state: {current_step}")
        except Exception as exc:
            self.signals.omniroute_failed.emit(str(exc))
            return
        finally:
            if not self.awaiting_registration_code:
                self.signals.omniroute_finished.emit()

    def handle_server_connected(self, browser_url: str) -> None:
        self.server_connection_in_progress = False
        self.refresh_button_state()
        self.append_log("SSH tunnel connected. Opening the local server in a browser.")
        webbrowser.open_new_tab(browser_url)

    def handle_server_connection_failed(self, message: str) -> None:
        self.server_connection_in_progress = False
        self.server_tunnel.stop()
        self.refresh_button_state()
        self.append_log(f"Server connection failed: {message}")
        self.show_error("Server connection failed", message)

    def handle_omniroute_failed(self, message: str) -> None:
        self.append_log(f"Omniroute failed: {message}")
        if self.reg_automation_active:
            self.fail_current_iteration(f"Omniroute failed: {message}")
            return
        self.reset_pending_code_state()
        self.omniroute_in_progress = False
        self.reg_in_progress = False
        self.reuse_omniroute_browser = False
        self.omniroute_restart_requested = False
        self.refresh_button_state()
        self.show_error("Omniroute failed", message)

    def handle_omniroute_finished(self) -> None:
        self.omniroute_in_progress = False
        should_increment_alias = self.reg_in_progress
        restart_after_reject = (
            self.reg_automation_active and self.omniroute_restart_requested
        )
        should_repeat_cycle = (
            self.reg_automation_active
            and self.reg_in_progress
            and not self.omniroute_restart_requested
        )
        self.reuse_omniroute_browser = not restart_after_reject
        self.reg_in_progress = False
        self.refresh_button_state()
        if should_increment_alias:
            self.signals.reg_cycle_finished.emit()
        if restart_after_reject:
            QTimer.singleShot(
                0,
                lambda: self.start_next_reg_cycle(
                    reason="reg: restarting cycle after Omniroute rejection"
                ),
            )
        if should_repeat_cycle:
            QTimer.singleShot(
                0,
                lambda: self.start_next_reg_cycle(reason="reg: starting next cycle"),
            )
        self.omniroute_restart_requested = False

    def open_chatgpt_in_edge_worker(self) -> None:
        try:
            email = self.current_email()
            if not email:
                raise ValueError("Enter an email before opening ChatGPT.")

            edge_executable = self.resolve_edge_executable()
            attempt = 0
            while True:
                attempt += 1
                edge_process = subprocess.Popen(
                    [edge_executable, "--inprivate", "--start-maximized", CHATGPT_URL],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
                self.chatgpt_edge_process_id = edge_process.pid

                try:
                    time.sleep(CHATGPT_START_DELAY_SECONDS)
                    time.sleep(4.0)
                    self.activate_edge_element_by_name(
                        edge_process.pid,
                        CHATGPT_SIGNUP_BUTTON_TEXT,
                        timeout_seconds=CHATGPT_STARTUP_GUARD_TIMEOUT_SECONDS,
                    )
                    time.sleep(CHATGPT_EMAIL_STEP_DELAY_SECONDS)
                    release_modifier_keys()
                    send_unicode_text(email)
                    press_virtual_key(VK_RETURN)
                    self.wait_for_chatgpt_initial_next_step(edge_process.pid)
                    self.start_registration_watch(edge_process.pid)
                    break
                except Exception as exc:
                    try:
                        self.close_chatgpt_edge_window()
                    except Exception as close_exc:
                        self.append_log(
                            f"ChatGPT startup guard: failed to close Edge after startup failure: {close_exc}"
                        )
                    if attempt >= 3:
                        raise RuntimeError(
                            f"ChatGPT startup guard failed after {attempt} attempts: {exc}"
                        ) from exc
                    self.append_log(
                        f"ChatGPT startup guard retry {attempt}/3: {exc}. Retrying with the same email."
                    )
                    time.sleep(2.0)
        except Exception as exc:
            self.signals.imap_error.emit(f"ChatGPT launch failed: {exc}")
            if self.reg_automation_active:
                QTimer.singleShot(
                    0,
                    lambda exc_text=str(exc): self.fail_current_iteration(
                        f"ChatGPT launch failed: {exc_text}"
                    ),
                )
            return

        self.signals.imap_status.emit(
            f"Opened ChatGPT in Edge InPrivate, activated '{CHATGPT_SIGNUP_BUTTON_TEXT}', inserted email, submitted it, and started watching the registration step."
        )

    def start_registration_watch(self, process_id: int) -> None:
        self.stop_registration_watch()
        self.registration_watch_stop_event = threading.Event()
        self.registration_watch_thread = threading.Thread(
            target=self.registration_watch_worker,
            args=(process_id, self.registration_watch_stop_event),
            daemon=True,
        )
        self.registration_watch_thread.start()

    def stop_registration_watch(self) -> None:
        self.registration_watch_stop_event.set()
        self.registration_watch_thread = None
        self.awaiting_registration_code = False
        self.registration_code_baseline = ""
        self.pending_code_flow = ""
        self.pending_code_process_id = 0

    def registration_watch_worker(
        self, process_id: int, stop_event: threading.Event
    ) -> None:
        seen_states: set[str] = set()
        deadline = time.monotonic() + CHATGPT_REGISTRATION_WATCH_TIMEOUT_SECONDS
        last_debug_signature = ""

        while not stop_event.is_set() and time.monotonic() < deadline:
            try:
                page_text = self.read_edge_window_text(
                    process_id,
                    timeout_seconds=CHATGPT_REGISTRATION_POLL_INTERVAL_SECONDS,
                )
            except Exception as exc:
                self.signals.imap_error.emit(f"Registration watcher failed: {exc}")
                return

            if page_text:
                debug_signature = page_text[:300]
                if debug_signature != last_debug_signature:
                    last_debug_signature = debug_signature
                    preview = page_text.replace("\n", " | ")[:300]
                    self.signals.imap_status.emit(
                        f"Registration watcher sees: {preview}"
                    )

            text_lines = [
                line.strip().casefold()
                for line in page_text.splitlines()
                if line.strip()
            ]
            normalized_text = "\n".join(text_lines)
            has_password_title = any(
                text in normalized_text for text in CHATGPT_CREATE_PASSWORD_TEXTS
            )
            has_password_label = any(
                line in CHATGPT_PASSWORD_LABEL_TEXTS for line in text_lines
            )

            if has_password_title and has_password_label:
                if "password" not in seen_states:
                    seen_states.add("password")
                    self.signals.registration_password_step_detected.emit()
                    self.signals.imap_status.emit(
                        "Registration watcher matched the password step."
                    )
                    stop_event.set()
                    return
            elif any(text in normalized_text for text in CHATGPT_CHECK_EMAIL_TEXTS):
                if "email_check" not in seen_states:
                    seen_states.add("email_check")
                    self.signals.registration_email_check_step_detected.emit()
                    self.signals.imap_status.emit(
                        "Registration watcher matched the email verification step."
                    )
                    stop_event.set()
                    return
            elif "email_check" in seen_states and not any(
                text in normalized_text for text in CHATGPT_CHECK_EMAIL_TEXTS
            ):
                self.awaiting_registration_code = False

            stop_event.wait(CHATGPT_REGISTRATION_POLL_INTERVAL_SECONDS)

        if not stop_event.is_set():
            self.signals.imap_status.emit(
                "Stopped watching ChatGPT registration state after timeout."
            )

    def activate_edge_element_by_name(
        self,
        process_id: int,
        element_name: str,
        *,
        timeout_seconds: float,
        allow_tab_fallback: bool = True,
    ) -> None:
        encoded_name = element_name.replace("'", "''")
        tab_fallback_block = """
        [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
        Start-Sleep -Milliseconds 30

        $focusedAfterTab = [System.Windows.Automation.AutomationElement]::FocusedElement
        if ($focusedAfterTab -ne $null -and $focusedAfterTab.Current.Name -eq $elementName) {
            [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
            Write-Output \"activated:$elementName\"
            exit 0
        }
"""
        if not allow_tab_fallback:
            tab_fallback_block = ""
        script = f"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class NativeWindowHelpers
{{
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}}
"@

$processId = {process_id}
$elementName = '{encoded_name}'
$timeoutSeconds = {timeout_seconds}
$deadline = [DateTime]::UtcNow.AddSeconds($timeoutSeconds)
$showMaximized = 3

function Get-EdgeWindowElement {{
    param([int]$TargetProcessId)

    $process = Get-Process -Id $TargetProcessId -ErrorAction SilentlyContinue
    if ($process -and $process.MainWindowHandle -ne 0) {{
        return [System.Windows.Automation.AutomationElement]::FromHandle($process.MainWindowHandle)
    }}

    $processNames = @('msedge', 'browser')
    foreach ($name in $processNames) {{
        $edgeProcess = Get-Process $name -ErrorAction SilentlyContinue |
            Where-Object {{ $_.MainWindowHandle -ne 0 }} |
            Sort-Object StartTime -Descending |
            Select-Object -First 1
        if ($edgeProcess) {{
            return [System.Windows.Automation.AutomationElement]::FromHandle($edgeProcess.MainWindowHandle)
        }}
    }}

    $matchedHandle = [IntPtr]::Zero
    $callback = [NativeWindowHelpers+EnumWindowsProc]{{
        param([IntPtr]$hWnd, [IntPtr]$lParam)

        if (-not [NativeWindowHelpers]::IsWindowVisible($hWnd)) {{
            return $true
        }}

        $textLength = [NativeWindowHelpers]::GetWindowTextLength($hWnd)
        if ($textLength -le 0) {{
            return $true
        }}

        $builder = New-Object System.Text.StringBuilder ($textLength + 1)
        [void][NativeWindowHelpers]::GetWindowText($hWnd, $builder, $builder.Capacity)

        [uint32]$windowProcessId = 0
        [void][NativeWindowHelpers]::GetWindowThreadProcessId($hWnd, [ref]$windowProcessId)
        try {{
            $windowProcess = Get-Process -Id ([int]$windowProcessId) -ErrorAction Stop
        }} catch {{
            return $true
        }}

        if ($windowProcess.ProcessName -in $processNames) {{
            $script:matchedHandle = $hWnd
            return $false
        }}

        return $true
    }}
    [void][NativeWindowHelpers]::EnumWindows($callback, [IntPtr]::Zero)
    if ($matchedHandle -ne [IntPtr]::Zero) {{
        return [System.Windows.Automation.AutomationElement]::FromHandle($matchedHandle)
    }}

    return $null
}}

while ([DateTime]::UtcNow -lt $deadline) {{
    $window = Get-EdgeWindowElement -TargetProcessId $processId
    if ($window -ne $null) {{
        $handle = [IntPtr]$window.Current.NativeWindowHandle
        if ($handle -ne [IntPtr]::Zero) {{
            [NativeWindowHelpers]::ShowWindowAsync($handle, $showMaximized) | Out-Null
            [NativeWindowHelpers]::SetForegroundWindow($handle) | Out-Null
        }}

        $focused = [System.Windows.Automation.AutomationElement]::FocusedElement
        if ($focused -ne $null -and $focused.Current.Name -eq $elementName) {{
            [System.Windows.Forms.SendKeys]::SendWait('{{ENTER}}')
            Write-Output "activated:$elementName"
            exit 0
        }}

        $target = $window.FindFirst(
            [System.Windows.Automation.TreeScope]::Descendants,
            (New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::NameProperty,
                $elementName
            ))
        )
        if ($target -ne $null) {{
            try {{
                $target.SetFocus()
                Start-Sleep -Milliseconds 30
                try {{
                    $invokePattern = $null
                    if ($target.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$invokePattern)) {{
                        $invokePattern.Invoke()
                        Write-Output "activated:$elementName"
                        exit 0
                    }}
                }} catch {{
                }}

                try {{
                    $legacyPattern = $null
                    if ($target.TryGetCurrentPattern([System.Windows.Automation.LegacyIAccessiblePattern]::Pattern, [ref]$legacyPattern)) {{
                        $legacyPattern.DoDefaultAction()
                        Write-Output "activated:$elementName"
                        exit 0
                    }}
                }} catch {{
                }}

                $focusedAfterSetFocus = [System.Windows.Automation.AutomationElement]::FocusedElement
                if ($focusedAfterSetFocus -ne $null -and $focusedAfterSetFocus.Current.Name -eq $elementName) {{
                    [System.Windows.Forms.SendKeys]::SendWait('{{ENTER}}')
                    Write-Output "activated:$elementName"
                    exit 0
                }}
            }} catch {{
            }}
        }}

{tab_fallback_block}
    }}

    Start-Sleep -Milliseconds 50
}}

Write-Error "Timed out waiting for UI element: $elementName"
exit 1
"""
        completed = run_powershell_script(script, timeout_seconds=timeout_seconds)
        if completed.returncode != 0:
            error_output = completed_process_output(completed)
            raise RuntimeError(error_output or f"Failed to activate '{element_name}'.")

    def focus_edge_password_field(
        self, *, timeout_seconds: float, process_id: int = 0
    ) -> None:
        script = f"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class NativeWindowHelpers
{{
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
}}
"@

$processId = {process_id}
$deadline = [DateTime]::UtcNow.AddSeconds({timeout_seconds})
$showMaximized = 3
$passwordNames = @('Пароль', 'Password')

function Get-BrowserWindowElement {{
    if ($processId -ne 0) {{
        $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
        if ($process -and $process.MainWindowHandle -ne 0) {{
            return [System.Windows.Automation.AutomationElement]::FromHandle($process.MainWindowHandle)
        }}
    }}

    $processNames = @('msedge', 'browser')

    foreach ($name in $processNames) {{
        $candidate = Get-Process $name -ErrorAction SilentlyContinue |
            Where-Object {{ $_.MainWindowHandle -ne 0 }} |
            Sort-Object StartTime -Descending |
            Select-Object -First 1
        if ($candidate) {{
            return [System.Windows.Automation.AutomationElement]::FromHandle($candidate.MainWindowHandle)
        }}
    }}

    $matchedHandle = [IntPtr]::Zero
    $callback = [NativeWindowHelpers+EnumWindowsProc]{{
        param([IntPtr]$hWnd, [IntPtr]$lParam)

        if (-not [NativeWindowHelpers]::IsWindowVisible($hWnd)) {{
            return $true
        }}

        $textLength = [NativeWindowHelpers]::GetWindowTextLength($hWnd)
        if ($textLength -le 0) {{
            return $true
        }}

        [uint32]$windowProcessId = 0
        [void][NativeWindowHelpers]::GetWindowThreadProcessId($hWnd, [ref]$windowProcessId)
        try {{
            $windowProcess = Get-Process -Id ([int]$windowProcessId) -ErrorAction Stop
        }} catch {{
            return $true
        }}

        if ($windowProcess.ProcessName -in @('msedge', 'browser')) {{
            $script:matchedHandle = $hWnd
            return $false
        }}

        return $true
    }}
    [void][NativeWindowHelpers]::EnumWindows($callback, [IntPtr]::Zero)
    if ($matchedHandle -ne [IntPtr]::Zero) {{
        return [System.Windows.Automation.AutomationElement]::FromHandle($matchedHandle)
    }}

    return $null
}}

while ([DateTime]::UtcNow -lt $deadline) {{
    $window = Get-BrowserWindowElement
    if ($window -ne $null) {{
        $handle = [IntPtr]$window.Current.NativeWindowHandle
        if ($handle -ne [IntPtr]::Zero) {{
            [NativeWindowHelpers]::ShowWindowAsync($handle, $showMaximized) | Out-Null
            [NativeWindowHelpers]::SetForegroundWindow($handle) | Out-Null
        }}

        $passwordLabel = $null
        foreach ($passwordName in $passwordNames) {{
            $passwordLabel = $window.FindFirst(
                [System.Windows.Automation.TreeScope]::Descendants,
                (New-Object System.Windows.Automation.PropertyCondition(
                    [System.Windows.Automation.AutomationElement]::NameProperty,
                    $passwordName
                ))
            )
            if ($passwordLabel -ne $null) {{
                break
            }}
        }}

        if ($passwordLabel -ne $null) {{
            try {{
                $passwordLabel.SetFocus()
                Start-Sleep -Milliseconds 80
                [System.Windows.Forms.SendKeys]::SendWait('{{TAB}}')
                Start-Sleep -Milliseconds 80
                $focused = [System.Windows.Automation.AutomationElement]::FocusedElement
                if ($focused -ne $null) {{
                    Write-Output 'focused:password'
                    exit 0
                }}
            }} catch {{
            }}
        }}

        $edits = $window.FindAll(
            [System.Windows.Automation.TreeScope]::Descendants,
            (New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                [System.Windows.Automation.ControlType]::Edit
            ))
        )

        if ($edits.Count -gt 0) {{
            try {{
                $edits.Item(0).SetFocus()
                Start-Sleep -Milliseconds 80
                Write-Output 'focused:fallback-edit'
                exit 0
            }} catch {{
            }}
        }}
    }}

    Start-Sleep -Milliseconds 100
}}

Write-Error 'Timed out waiting for password field.'
exit 1
"""
        completed = run_powershell_script(script, timeout_seconds=timeout_seconds)
        if completed.returncode != 0:
            error_output = completed_process_output(completed)
            raise RuntimeError(error_output or "Failed to focus password field.")

    def read_edge_window_text(self, process_id: int, *, timeout_seconds: float) -> str:
        script = f"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class NativeWindowHelpers
{{
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}}
"@

$processId = {process_id}
$deadline = [DateTime]::UtcNow.AddSeconds({timeout_seconds})

function Get-EdgeWindowElement {{
    param([int]$TargetProcessId)

    $process = Get-Process -Id $TargetProcessId -ErrorAction SilentlyContinue
    if ($process -and $process.MainWindowHandle -ne 0) {{
        return [System.Windows.Automation.AutomationElement]::FromHandle($process.MainWindowHandle)
    }}

    $processNames = @('msedge', 'browser')
    foreach ($name in $processNames) {{
        $edgeProcess = Get-Process $name -ErrorAction SilentlyContinue |
            Where-Object {{ $_.MainWindowHandle -ne 0 }} |
            Sort-Object StartTime -Descending |
            Select-Object -First 1
        if ($edgeProcess) {{
            return [System.Windows.Automation.AutomationElement]::FromHandle($edgeProcess.MainWindowHandle)
        }}
    }}

    $matchedHandle = [IntPtr]::Zero
    $callback = [NativeWindowHelpers+EnumWindowsProc]{{
        param([IntPtr]$hWnd, [IntPtr]$lParam)

        if (-not [NativeWindowHelpers]::IsWindowVisible($hWnd)) {{
            return $true
        }}

        $textLength = [NativeWindowHelpers]::GetWindowTextLength($hWnd)
        if ($textLength -le 0) {{
            return $true
        }}

        [uint32]$windowProcessId = 0
        [void][NativeWindowHelpers]::GetWindowThreadProcessId($hWnd, [ref]$windowProcessId)
        try {{
            $windowProcess = Get-Process -Id ([int]$windowProcessId) -ErrorAction Stop
        }} catch {{
            return $true
        }}

        if ($windowProcess.ProcessName -in $processNames) {{
            $script:matchedHandle = $hWnd
            return $false
        }}

        return $true
    }}
    [void][NativeWindowHelpers]::EnumWindows($callback, [IntPtr]::Zero)
    if ($matchedHandle -ne [IntPtr]::Zero) {{
        return [System.Windows.Automation.AutomationElement]::FromHandle($matchedHandle)
    }}

    return $null
}}

while ([DateTime]::UtcNow -lt $deadline) {{
    $window = Get-EdgeWindowElement -TargetProcessId $processId
    if ($window -ne $null) {{
        $nodes = $window.FindAll(
            [System.Windows.Automation.TreeScope]::Descendants,
            [System.Windows.Automation.Condition]::TrueCondition
        )

        $parts = New-Object System.Collections.Generic.List[string]
        $windowName = $window.Current.Name
        if (-not [string]::IsNullOrWhiteSpace($windowName)) {{
            $parts.Add($windowName)
        }}

        $focused = [System.Windows.Automation.AutomationElement]::FocusedElement
        if ($focused -ne $null) {{
            $focusedName = $focused.Current.Name
            if (-not [string]::IsNullOrWhiteSpace($focusedName)) {{
                $parts.Add($focusedName)
            }}
        }}

        for ($index = 0; $index -lt $nodes.Count; $index++) {{
            $name = $nodes.Item($index).Current.Name
            if (-not [string]::IsNullOrWhiteSpace($name)) {{
                $parts.Add($name)
            }}
        }}

        if ($parts.Count -gt 0) {{
            Write-Output ($parts -join "`n")
            exit 0
        }}
    }}

    Start-Sleep -Milliseconds 100
}}

Write-Error "Timed out waiting for Edge window text."
exit 1
"""
        completed = run_powershell_script(script, timeout_seconds=timeout_seconds)
        if completed.returncode != 0:
            error_output = completed_process_output(completed)
            raise RuntimeError(error_output or "Failed to read Edge window text.")
        return (completed.stdout or "").strip()

    def edge_window_contains_text(self, text: str, *, timeout_seconds: float) -> bool:
        page_text = self.read_edge_window_text(0, timeout_seconds=timeout_seconds)
        return text.casefold() in page_text.casefold()

    def resolve_edge_executable(self) -> str:
        edge_path = shutil.which("msedge")
        if edge_path:
            return edge_path

        candidate_paths = [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        ]
        for candidate in candidate_paths:
            if Path(candidate).exists():
                return candidate

        raise FileNotFoundError("Microsoft Edge executable was not found.")

    def resolve_yandex_browser_executable(self) -> str:
        browser_path = shutil.which("browser")
        if browser_path:
            return browser_path

        candidate_paths = [
            os.path.expandvars(
                r"%LOCALAPPDATA%\Yandex\YandexBrowser\Application\browser.exe"
            ),
            r"C:\Program Files\Yandex\YandexBrowser\Application\browser.exe",
            r"C:\Program Files (x86)\Yandex\YandexBrowser\Application\browser.exe",
        ]
        for candidate in candidate_paths:
            if Path(candidate).exists():
                return candidate

        raise FileNotFoundError("Yandex Browser executable was not found.")

    def open_imap_login_dialog(self) -> None:
        dialog = ZohoLoginDialog(self, email=self.stored_imap_email())
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        imap_email, imap_password = dialog.credentials()
        normalized_password = normalize_zoho_password(imap_password)
        if not imap_email or not imap_password:
            self.show_error(
                "Invalid IMAP settings",
                "Enter both a Zoho Mail address and a password.",
            )
            return

        try:
            verify_zoho_imap_credentials(imap_email, normalized_password)
        except imaplib.IMAP4.error:
            self.show_error(
                "Zoho login failed",
                "Zoho rejected the credentials. If you pasted an App Password, regenerate it and paste it without spaces.",
            )
            return
        except Exception as exc:
            self.show_error("Zoho login failed", str(exc))
            return

        self.save_imap_credentials(imap_email, normalized_password)
        try:
            self.register_text_hotkeys()
        except Exception as exc:
            self.unregister_text_hotkeys()
            self.show_error("Failed to register hotkeys", str(exc))
            return
        self.start_imap_monitor(log_startup=True)

    def logout_imap_account(self) -> None:
        if not self.stored_imap_email():
            return
        self.stop_imap_monitor()
        self.clear_imap_credentials()
        self.append_log("Signed out from Zoho IMAP.")

    def stop_imap_monitor(self) -> None:
        self.imap_stop_event.set()
        self.imap_thread = None

    def start_imap_monitor(self, *, log_startup: bool) -> None:
        self.stop_imap_monitor()
        self.imap_stop_event = threading.Event()
        self.refresh_button_state()

        imap_email = self.stored_imap_email()
        imap_password = self.stored_imap_password()
        if not imap_email or not imap_password:
            return

        self.imap_thread = threading.Thread(
            target=self.imap_monitor_worker,
            args=(imap_email, imap_password, self.imap_stop_event, log_startup),
            daemon=True,
        )
        self.imap_thread.start()

    def imap_monitor_worker(
        self,
        imap_email: str,
        imap_password: str,
        stop_event: threading.Event,
        log_startup: bool,
    ) -> None:
        announced_ready = False
        processed_message_ids: set[str] = set()

        while not stop_event.is_set():
            mailbox: imaplib.IMAP4_SSL | None = None
            try:
                mailbox = imaplib.IMAP4_SSL(
                    resolve_zoho_imap_host(imap_email), ZOHO_IMAP_PORT
                )
                mailbox.login(imap_email, imap_password)
                select_status, _ = mailbox.select("INBOX", readonly=True)
                if select_status != "OK":
                    raise RuntimeError("Failed to open Zoho inbox.")

                discovered_ids, code = fetch_unseen_openai_code_from_all_folders(
                    mailbox, processed_message_ids
                )
                processed_message_ids.update(discovered_ids)
                if code:
                    self.signals.imap_code_received.emit(code)

                if log_startup and not announced_ready:
                    self.signals.imap_status.emit(
                        f"Zoho IMAP connected for {imap_email}. Watching unread mail in all folders for OpenAI codes."
                    )
                    announced_ready = True
            except Exception as exc:
                self.signals.imap_error.emit(str(exc))
            finally:
                if mailbox is not None:
                    try:
                        mailbox.logout()
                    except Exception:
                        pass

            stop_event.wait(IMAP_POLL_INTERVAL_SECONDS)

    def handle_imap_code_received(self, code: str) -> None:
        self.code_value_input.setText(code)

        if not self.awaiting_registration_code:
            self.append_log(
                f"Ignored OpenAI code because no verification step is active yet: {code}"
            )
            return

        try:
            self.clipboard_text(code)
        except Exception as exc:
            self.append_log(f"Failed to copy new OpenAI code: {exc}")
        else:
            self.append_log(f"Copied new OpenAI code: {code}")

        if self.tray_icon.isVisible():
            self.tray_icon.showMessage(
                APP_NAME,
                f"New OpenAI code: {code}",
                QSystemTrayIcon.MessageIcon.Information,
                5000,
            )
        QApplication.beep()
        self.append_log(f"Received new OpenAI code: {code}")

        if self.awaiting_registration_code and code != self.registration_code_baseline:
            if self.code_submit_in_progress:
                self.append_log(
                    f"Ignored duplicate registration code while another code submit is in progress: {code}"
                )
                return

            pending_flow = self.pending_code_flow
            process_id = self.pending_code_process_id
            refocus_needed = self.code_wait_refocus_needed
            clear_before_submit = self.code_wait_clear_before_submit
            self.awaiting_registration_code = False
            self.registration_code_baseline = code
            self.pending_code_flow = ""
            self.pending_code_process_id = 0
            self.code_submit_in_progress = True
            self.code_wait_refocus_needed = False
            self.code_wait_clear_before_submit = False

            def insert_code_and_submit() -> None:
                try:
                    if pending_flow == "omniroute" and clear_before_submit:
                        self.focus_and_clear_active_code_field("omniroute")
                    elif pending_flow == "chatgpt" and clear_before_submit:
                        self.focus_and_clear_chatgpt_code_field(process_id)
                    elif refocus_needed:
                        if pending_flow == "omniroute":
                            self.focus_active_code_field("omniroute")
                        elif pending_flow == "chatgpt":
                            self.focus_active_code_field("chatgpt")
                    self.text_bind_worker(code, "Registration code")
                    time.sleep(0.15)
                    press_virtual_key(VK_RETURN)
                    if pending_flow == "chatgpt":
                        chatgpt_post_code_state = self.wait_for_chatgpt_post_code_state(
                            process_id
                        )
                        if chatgpt_post_code_state == "invalid_code":
                            self.handle_chatgpt_invalid_code(process_id)
                            return
                        self.complete_profile_after_code()
                    elif pending_flow == "omniroute":
                        self.complete_omniroute_after_code(process_id)
                except Exception as exc:
                    if pending_flow == "omniroute":
                        self.signals.omniroute_failed.emit(str(exc))
                        return
                    self.signals.imap_error.emit(f"Code submit failed: {exc}")
                    if self.reg_automation_active:
                        QTimer.singleShot(
                            0,
                            lambda exc_text=str(exc): self.fail_current_iteration(
                                f"Code submit failed: {exc_text}"
                            ),
                        )
                    return
                finally:
                    self.code_submit_in_progress = False
                    if pending_flow == "omniroute":
                        self.signals.omniroute_finished.emit()

            threading.Thread(target=insert_code_and_submit, daemon=True).start()
            self.append_log(
                "Inserted the latest verification code into the active field and submitted it."
            )

    def handle_imap_error(self, message: str) -> None:
        self.append_log(f"Zoho IMAP error: {message}")

    def increment_email(self) -> None:
        try:
            next_index = extract_alias_index(self.current_email()) + 1
            next_email = build_alias_email(self.current_email(), next_index)
        except ValueError as exc:
            self.show_error("Invalid email", str(exc))
            return

        self.email_input.setText(next_email)
        self.append_log(f"Prepared alias {next_email}.")


def main() -> int:
    install_global_exception_logging()
    application = QApplication(sys.argv)
    window = CodexApp()
    window.show()
    return application.exec()


if __name__ == "__main__":
    raise SystemExit(main())
