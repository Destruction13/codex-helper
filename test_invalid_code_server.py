from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


HOST = "127.0.0.1"
PORT = 8877
HTML_PATH = Path(__file__).resolve().with_name("test_invalid_code_flow.html")
FOCUS_HTML_PATH = (
    Path(__file__).resolve().with_name("test_invalid_code_focus_flow.html")
)
RESEND_HTML_PATH = Path(__file__).resolve().with_name("test_resend_code_flow.html")


class InvalidCodeHandler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:  # noqa: N802
        if self.path in {"/", "/test-invalid-code", "/test-invalid-code/"}:
            content = HTML_PATH.read_bytes()
        elif self.path in {"/test-invalid-code-focus", "/test-invalid-code-focus/"}:
            content = FOCUS_HTML_PATH.read_bytes()
        elif self.path in {"/test-resend-code", "/test-resend-code/"}:
            content = RESEND_HTML_PATH.read_bytes()
        else:
            self.send_response(404)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.end_headers()
            self.wfile.write("Not Found".encode("utf-8"))
            return

        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def log_message(self, format: str, *args) -> None:  # noqa: A003
        return


def main() -> int:
    server = ThreadingHTTPServer((HOST, PORT), InvalidCodeHandler)
    print(
        f"Test invalid-code harnesses are running on http://{HOST}:{PORT}/test-invalid-code, http://{HOST}:{PORT}/test-invalid-code-focus and http://{HOST}:{PORT}/test-resend-code"
    )
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
