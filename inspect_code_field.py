import subprocess
from pathlib import Path


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


def main() -> int:
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
    output = (completed.stdout or "") + (completed.stderr or "")
    print(output.strip())
    return completed.returncode


if __name__ == "__main__":
    raise SystemExit(main())
