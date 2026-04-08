import shutil
import subprocess
import sys
import time
from pathlib import Path


TEST_URL = "http://127.0.0.1:8877/test-resend-code"
WINDOW_TITLE_PATTERN = "Resend Code Flow Test"


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
        encoding="utf-8",
        timeout=timeout_seconds,
        check=False,
    )


def open_test_tab() -> None:
    browser_executable = resolve_yandex_browser_executable()
    subprocess.Popen(
        [browser_executable, "--new-window", "--start-maximized", TEST_URL],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    time.sleep(3)


def read_page_state() -> str:
    script = rf"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$window = Get-Process browser -ErrorAction SilentlyContinue |
    Where-Object {{ $_.MainWindowHandle -ne 0 }} |
    Sort-Object StartTime -Descending |
    Select-Object -First 1

if (-not $window) {{
    Write-Error 'Resend test browser tab was not found.'
    exit 1
}}

$root = [System.Windows.Automation.AutomationElement]::FromHandle($window.MainWindowHandle)
if (-not $root) {{
    Write-Error 'Automation root was not found.'
    exit 1
}}

$nodes = $root.FindAll(
    [System.Windows.Automation.TreeScope]::Descendants,
    [System.Windows.Automation.Condition]::TrueCondition
)

$parts = New-Object System.Collections.Generic.List[string]
for ($index = 0; $index -lt $nodes.Count; $index++) {{
    $name = $nodes.Item($index).Current.Name
    if (-not [string]::IsNullOrWhiteSpace($name)) {{
        $parts.Add($name)
    }}
}}

Write-Output ($parts -join "`n")
exit 0
"""
    result = run_powershell_script(script, timeout_seconds=15.0)
    if result.returncode != 0:
        raise RuntimeError(
            result.stderr or result.stdout or "Failed to read resend page state."
        )
    return result.stdout or ""


def click_button_by_index(index: int, label: str) -> None:
    script = rf"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

$window = Get-Process browser -ErrorAction SilentlyContinue |
    Where-Object {{ $_.MainWindowHandle -ne 0 }} |
    Sort-Object StartTime -Descending |
    Select-Object -First 1

if (-not $window) {{
    Write-Error 'Resend test browser tab was not found.'
    exit 1
}}

[void][Microsoft.VisualBasic.Interaction]::AppActivate($window.Id)
Start-Sleep -Milliseconds 300

$root = [System.Windows.Automation.AutomationElement]::FromHandle($window.MainWindowHandle)
$nodes = $root.FindAll(
    [System.Windows.Automation.TreeScope]::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Button
    ))
)

$targetIndex = {index}
if ($nodes.Count -le $targetIndex) {{
    Write-Error 'Requested button index was not found.'
    exit 1
}}

$node = $nodes.Item($targetIndex)
$name = $node.Current.Name
try {{
    $node.SetFocus()
    Start-Sleep -Milliseconds 150
    [System.Windows.Forms.SendKeys]::SendWait('{{ENTER}}')
    Write-Output "clicked:$name"
    exit 0
}} catch {{
}}

Write-Error 'Requested button could not be activated.'
exit 1
"""
    result = run_powershell_script(script, timeout_seconds=15.0)
    if result.returncode != 0:
        raise RuntimeError(
            result.stderr or result.stdout or f"Failed to click button {label}."
        )
    print(result.stdout)


def focus_and_clear_code_field() -> None:
    script = rf"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

$window = Get-Process browser -ErrorAction SilentlyContinue |
    Where-Object {{ $_.MainWindowHandle -ne 0 }} |
    Sort-Object StartTime -Descending |
    Select-Object -First 1

if (-not $window) {{
    Write-Error 'Resend test browser tab was not found.'
    exit 1
}}

[void][Microsoft.VisualBasic.Interaction]::AppActivate($window.Id)
Start-Sleep -Milliseconds 300
$root = [System.Windows.Automation.AutomationElement]::FromHandle($window.MainWindowHandle)
$edits = $root.FindAll(
    [System.Windows.Automation.TreeScope]::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Edit
    ))
)

if ($edits.Count -lt 1) {{
    Write-Error 'Edit field was not found.'
    exit 1
}}

$targetEdit = $edits.Item(0)
$targetEdit.SetFocus()
Start-Sleep -Milliseconds 150

try {{
    $valuePattern = $null
    if ($targetEdit.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$valuePattern)) {{
        $valuePattern.SetValue('')
        Start-Sleep -Milliseconds 150
        Write-Output 'cleared:value-pattern'
        exit 0
    }}
}} catch {{
}}

[System.Windows.Forms.SendKeys]::SendWait('{{END}}')
Start-Sleep -Milliseconds 100
for ($index = 0; $index -lt 6; $index++) {{
    [System.Windows.Forms.SendKeys]::SendWait('{{BACKSPACE}}')
    Start-Sleep -Milliseconds 60
}}

Write-Output 'cleared:fallback'
exit 0
"""
    result = run_powershell_script(script, timeout_seconds=15.0)
    if result.returncode != 0:
        raise RuntimeError(
            result.stderr or result.stdout or "Failed to clear code field."
        )
    print(result.stdout)


def type_code(code: str) -> None:
    escaped = code.replace("'", "''")
    script = rf"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

$window = Get-Process browser -ErrorAction SilentlyContinue |
    Where-Object {{
        $_.MainWindowHandle -ne 0 -and
        $_.MainWindowTitle -match '{WINDOW_TITLE_PATTERN}'
    }} |
    Sort-Object StartTime -Descending |
    Select-Object -First 1

if (-not $window) {{
    Write-Error 'Resend test browser tab was not found.'
    exit 1
}}

[void][Microsoft.VisualBasic.Interaction]::AppActivate($window.Id)
Start-Sleep -Milliseconds 300
[System.Windows.Forms.SendKeys]::SendWait('{escaped}')
Start-Sleep -Milliseconds 100
[System.Windows.Forms.SendKeys]::SendWait('{{ENTER}}')
Write-Output 'typed'
exit 0
"""
    result = run_powershell_script(script, timeout_seconds=10.0)
    if result.returncode != 0:
        raise RuntimeError(result.stderr or result.stdout or "Failed to type code.")
    print(result.stdout)


def extract_marker(page_text: str, prefix: str) -> str:
    for line in page_text.splitlines():
        if prefix in line:
            return line.split(prefix, 1)[1].strip()
    raise RuntimeError(f"Marker {prefix!r} was not found in page state.")


def main() -> int:
    print(f"Opening resend test: {TEST_URL}")
    open_test_tab()

    print("Initial state:")
    print(read_page_state())

    print("Waiting 30 seconds before resend, as in the target scenario...")
    time.sleep(30)

    print("Clicking resend after the wait...")
    click_button_by_index(1, "resend")
    time.sleep(1)
    page_text = read_page_state()
    print(page_text)

    resend_count = extract_marker(page_text, "RESEND_COUNT: ")
    latest_code = extract_marker(page_text, "LATEST_CODE: ")
    if resend_count != "1":
        raise RuntimeError(f"Expected RESEND_COUNT 1, got {resend_count!r}")
    if latest_code == "NONE":
        raise RuntimeError("Expected a delivered code after the resend.")

    print("Returning focus to the code field and clearing it...")
    focus_and_clear_code_field()

    print(f"Typing delivered code: {latest_code}")
    type_code(latest_code)
    time.sleep(1)
    final_state = read_page_state()
    print(final_state)
    if "Код принят" not in final_state:
        raise RuntimeError("Resend retry scenario did not finish successfully.")

    print("Resend retry scenario succeeded.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
