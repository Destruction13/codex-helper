import shutil
import subprocess
import sys
import time
from pathlib import Path


TEST_PROVIDER_URL = (
    "data:text/html,<title>OpenAI Test Tab</title><h1>OpenAI Test Tab</h1>"
)
TEST_BASE_URL = (
    "data:text/html,<title>OmniRoute Test Tab</title><h1>OmniRoute Test Tab</h1>"
)


def resolve_yandex_browser_executable() -> str:
    browser_path = shutil.which("browser")
    if browser_path:
        return browser_path

    candidate_paths = [
        str(
            Path.home()
            / "AppData"
            / "Local"
            / "Yandex"
            / "YandexBrowser"
            / "Application"
            / "browser.exe"
        ),
        r"C:\Program Files\Yandex\YandexBrowser\Application\browser.exe",
        r"C:\Program Files (x86)\Yandex\YandexBrowser\Application\browser.exe",
    ]
    for candidate in candidate_paths:
        if Path(candidate).exists():
            return candidate

    raise FileNotFoundError("Yandex Browser executable was not found.")


def run_powershell_script(
    script: str, timeout_seconds: float = 15.0
) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            script,
        ],
        capture_output=True,
        text=True,
        timeout=timeout_seconds,
        check=False,
    )


def close_active_yandex_tab(
    expect_title_fragment: str,
) -> subprocess.CompletedProcess[str]:
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
Write-Output "before:$beforeTitle"
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

if ($after) {{
    Write-Output "after:$($after.MainWindowTitle)"
    if ($after.MainWindowTitle -eq $beforeTitle) {{
        Write-Error 'Yandex tab title did not change after close attempt.'
        exit 1
    }}
}} else {{
    Write-Output 'after:<no-window>'
}}

Write-Output 'closed-tab'
exit 0
"""
    return run_powershell_script(script, timeout_seconds=10.0)


def main() -> int:
    browser_executable = resolve_yandex_browser_executable()
    print(f"Using browser: {browser_executable}")

    existing_window = run_powershell_script(
        """
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$target = Get-Process browser -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowHandle -ne 0 } |
    Sort-Object StartTime -Descending |
    Select-Object -First 1

if (-not $target) {
    Write-Error 'Open any Yandex Browser window before running this test.'
    exit 1
}

Write-Output $target.Id
exit 0
""",
        timeout_seconds=10.0,
    )
    if existing_window.returncode != 0:
        print(existing_window.stderr or existing_window.stdout, file=sys.stderr)
        return 1

    subprocess.Popen(
        [browser_executable, TEST_BASE_URL],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    time.sleep(2)
    subprocess.Popen(
        [browser_executable, TEST_PROVIDER_URL],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    time.sleep(2)

    print("Closing provider-like tab...")
    provider_result = close_active_yandex_tab("OpenAI")
    print(provider_result.stdout)
    if provider_result.stderr:
        print(provider_result.stderr, file=sys.stderr)

    print("Closing base-like tab...")
    base_result = close_active_yandex_tab("OmniRoute")
    print(base_result.stdout)
    if base_result.stderr:
        print(base_result.stderr, file=sys.stderr)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
