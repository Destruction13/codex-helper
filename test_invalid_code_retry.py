import shutil
import subprocess
import sys
import time
from pathlib import Path


TEST_URL = "http://127.0.0.1:8877/test-invalid-code-focus"
WINDOW_TITLE_PATTERN = "Invalid Code Focus Test"


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
        [browser_executable, TEST_URL],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    time.sleep(2)


def focus_input_and_clear() -> subprocess.CompletedProcess[str]:
    script = r"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

$window = Get-Process browser -ErrorAction SilentlyContinue |
    Where-Object {
        $_.MainWindowHandle -ne 0 -and
        $_.MainWindowTitle -match 'Invalid Code Focus Test'
    } |
    Sort-Object StartTime -Descending |
    Select-Object -First 1

if (-not $window) {
    Write-Error 'Focus test browser tab was not found.'
    exit 1
}

[void][Microsoft.VisualBasic.Interaction]::AppActivate($window.Id)
Start-Sleep -Milliseconds 300

$root = [System.Windows.Automation.AutomationElement]::FromHandle($window.MainWindowHandle)
if (-not $root) {
    Write-Error 'Automation root was not found.'
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
    Write-Error 'Edit field was not found on the focus test page.'
    exit 1
}

$targetEdit = $edits.Item(0)
try {
    $targetEdit.SetFocus()
    Start-Sleep -Milliseconds 150
} catch {
    Write-Error 'Failed to focus the edit field.'
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

[System.Windows.Forms.SendKeys]::SendWait('{END}')
Start-Sleep -Milliseconds 100
for ($index = 0; $index -lt 6; $index++) {
    [System.Windows.Forms.SendKeys]::SendWait('{BACKSPACE}')
    Start-Sleep -Milliseconds 60
}
Write-Output 'cleared'
exit 0
"""
    return run_powershell_script(script, timeout_seconds=15.0)


def focus_code_field() -> subprocess.CompletedProcess[str]:
    script = r"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

$window = Get-Process browser -ErrorAction SilentlyContinue |
    Where-Object {
        $_.MainWindowHandle -ne 0 -and
        $_.MainWindowTitle -match 'Invalid Code Focus Test'
    } |
    Sort-Object StartTime -Descending |
    Select-Object -First 1

if (-not $window) {
    Write-Error 'Focus test browser tab was not found.'
    exit 1
}

[void][Microsoft.VisualBasic.Interaction]::AppActivate($window.Id)
Start-Sleep -Milliseconds 300

$root = [System.Windows.Automation.AutomationElement]::FromHandle($window.MainWindowHandle)
if (-not $root) {
    Write-Error 'Automation root was not found.'
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
    Write-Error 'Edit field was not found on the focus test page.'
    exit 1
}

$targetEdit = $edits.Item(0)
try {
    $targetEdit.SetFocus()
    Start-Sleep -Milliseconds 150
    Write-Output 'focused'
    exit 0
} catch {
}

Write-Error 'Failed to focus the code field.'
exit 1
"""
    return run_powershell_script(script, timeout_seconds=15.0)


def read_page_state() -> subprocess.CompletedProcess[str]:
    script = r"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$window = Get-Process browser -ErrorAction SilentlyContinue |
    Where-Object {
        $_.MainWindowHandle -ne 0 -and
        $_.MainWindowTitle -match 'Invalid Code Focus Test'
    } |
    Sort-Object StartTime -Descending |
    Select-Object -First 1

if (-not $window) {
    Write-Error 'Focus test browser tab was not found.'
    exit 1
}

$root = [System.Windows.Automation.AutomationElement]::FromHandle($window.MainWindowHandle)
if (-not $root) {
    Write-Error 'Automation root was not found.'
    exit 1
}

$nodes = $root.FindAll(
    [System.Windows.Automation.TreeScope]::Descendants,
    [System.Windows.Automation.Condition]::TrueCondition
)

$parts = New-Object System.Collections.Generic.List[string]
for ($index = 0; $index -lt $nodes.Count; $index++) {
    $name = $nodes.Item($index).Current.Name
    if (-not [string]::IsNullOrWhiteSpace($name)) {
        $parts.Add($name)
    }
}

Write-Output ($parts -join "`n")
exit 0
"""
    return run_powershell_script(script, timeout_seconds=15.0)


def type_text(text: str, submit: bool = False) -> subprocess.CompletedProcess[str]:
    escaped = text.replace("'", "''")
    submit_block = ""
    if submit:
        submit_block = "[System.Windows.Forms.SendKeys]::SendWait('{ENTER}')"
    script = f"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

$window = Get-Process browser -ErrorAction SilentlyContinue |
    Where-Object {{
        $_.MainWindowHandle -ne 0 -and
        $_.MainWindowTitle -match 'Invalid Code Focus Test'
    }} |
    Sort-Object StartTime -Descending |
    Select-Object -First 1

if (-not $window) {{
    Write-Error 'Focus test browser tab was not found.'
    exit 1
}}

[void][Microsoft.VisualBasic.Interaction]::AppActivate($window.Id)
Start-Sleep -Milliseconds 300
[System.Windows.Forms.SendKeys]::SendWait('{escaped}')
Start-Sleep -Milliseconds 100
{submit_block}
Write-Output 'typed'
exit 0
"""
    return run_powershell_script(script, timeout_seconds=10.0)


def extract_new_code(page_text: str) -> str:
    marker = "LATEST_CODE: "
    for line in page_text.splitlines():
        if marker in line:
            return line.split(marker, 1)[1].strip()
    raise RuntimeError("Failed to extract a newly generated code from the page log.")


def has_state(page_text: str, state: str) -> bool:
    return f"STATE: {state}" in page_text


def main() -> int:
    print(f"Opening focus test: {TEST_URL}")
    open_test_tab()

    print("Focusing the code field...")
    focus_result = focus_code_field()
    print(focus_result.stdout)
    if focus_result.stderr:
        print(focus_result.stderr, file=sys.stderr)

    print("Typing an invalid code...")
    invalid_result = type_text("111111", submit=True)
    print(invalid_result.stdout)
    if invalid_result.stderr:
        print(invalid_result.stderr, file=sys.stderr)

    time.sleep(1)
    page_result = read_page_state()
    page_text = page_result.stdout or ""
    print("Current page state captured.")
    if not has_state(page_text, "INVALID_CODE"):
        print(page_text)
        raise RuntimeError("The focus test page did not enter INVALID_CODE state.")

    print("Refocusing the code field and clearing it...")
    clear_result = focus_input_and_clear()
    print(clear_result.stdout)
    if clear_result.stderr:
        print(clear_result.stderr, file=sys.stderr)

    print("Waiting for a new generated code...")
    time.sleep(3)
    page_result = read_page_state()
    page_text = page_result.stdout or ""
    new_code = extract_new_code(page_text)
    print(f"Detected new code: {new_code}")

    print("Refocusing the code field before retry...")
    focus_retry_result = focus_code_field()
    print(focus_retry_result.stdout)
    if focus_retry_result.stderr:
        print(focus_retry_result.stderr, file=sys.stderr)

    print("Clearing the field again before retry...")
    clear_retry_result = focus_input_and_clear()
    print(clear_retry_result.stdout)
    if clear_retry_result.stderr:
        print(clear_retry_result.stderr, file=sys.stderr)

    print("Typing the new code...")
    retry_result = type_text(new_code, submit=True)
    print(retry_result.stdout)
    if retry_result.stderr:
        print(retry_result.stderr, file=sys.stderr)

    time.sleep(1)
    final_state = read_page_state().stdout or ""
    if not has_state(final_state, "SUCCESS"):
        print(final_state)
        raise RuntimeError("Retry scenario did not finish successfully.")

    print("Retry scenario succeeded.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
