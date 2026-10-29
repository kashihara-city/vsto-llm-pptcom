<#
  Drives the real add-in in PowerPoint from outside, through its automation object (Pptcom.AddinAutomation),
  for tests that need it as the user runs it: the conversation across requests, hand edits in between, the task
  pane. Needs 他のプログラムからの操作を許可する turned on in the task pane's settings (off by default); every call is refused
  otherwise. Prints first which build answered ({"addin": "0.0.4 5efe257+ Debug"}).

  Each step is one of:
    any text              send it as a request and wait until it ends (prints the add-in's LastRun as one JSON line)
    /new                  新しい会話
    /continue             続ける: carry on the request that stopped at the turn limit, and wait until it ends
    /cancel-after N text  send the request and press stop after N seconds
    /attach path          attach a document as if dropped on the task pane (waits until it is converted)
    /ps code              run PowerShell code with $ppt (PowerPoint.Application) and $pres (the active presentation),
                          e.g. a hand edit: "/ps $pres.Slides(1).Shapes(1).TextFrame.TextRange.Text = 'edited'"

  Uses the PowerPoint that is running, or starts one (minimized) with a new presentation. Windows PowerShell 5.1
  (late-bound COM).

  usage:
    powershell -NoProfile -ExecutionPolicy Bypass -File scripts\test\drive_addin.ps1 -NewDeck -StepsFile steps.txt
    (from Git Bash set MSYS_NO_PATHCONV=1, or it turns "/new" into a path)
#>
param(
    [string[]]$Steps = @(),
    # One step per line (UTF-8), for callers that cannot pass an array (powershell -File takes one string per parameter).
    [string]$StepsFile = "",
    [int]$TimeoutSec = 900,
    [string]$Open = "",
    # Work in a new presentation instead of the one the user has open (the conversation is per presentation).
    [switch]$NewDeck
)
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ($StepsFile) { $Steps += @(Get-Content -LiteralPath $StepsFile -Encoding UTF8 | Where-Object { $_.Trim() -ne "" }) }
if ($Steps.Count -eq 0) { throw "No steps: pass -Steps or -StepsFile." }

try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application") }
catch {
    # Not New-Object -ComObject: setting Visible on that one failed with TYPE_E_ELEMENTNOTFOUND (0x8002802B) here.
    $ppt = [Activator]::CreateInstance([type]::GetTypeFromProgID("PowerPoint.Application"))
    $ppt.Visible = -1          # msoTrue: the add-in's task pane and ActivePresentation need a window
    $ppt.WindowState = 2       # ppWindowMinimized
}
if ($Open) { $deck = $ppt.Presentations.Open((Resolve-Path $Open).Path) }
elseif ($NewDeck -or $ppt.Presentations.Count -eq 0) { $deck = $ppt.Presentations.Add(-1) }
else { $deck = $ppt.ActivePresentation }

$addin = $ppt.COMAddIns.Item("pptcom")
if (-not $addin.Connect) { throw "pptcom (vsto-llm-pptcom) is not loaded in this PowerPoint." }
$ai = $addin.Object
if ($null -eq $ai) { throw "pptcom gave no automation object." }
# A refused call comes back as an empty value through late-bound COM, not as an error.
if ($null -eq $ai.IsBusy) { throw "他のプログラムからの操作がオフです。作業ウィンドウの接続設定で「他のプログラムからの操作を許可する」をオンにしてください。" }
# Which build answered (version, commit, Debug): a Debug build and an installed one register under the same name.
Write-Output (@{ addin = $ai.Version } | ConvertTo-Json -Compress)

function Wait-Idle([int]$limit) {
    $start = Get-Date
    while ($ai.IsBusy) {
        if (((Get-Date) - $start).TotalSeconds -gt $limit) { $ai.Cancel(); throw "request still running after $limit s; stopped it" }
        Start-Sleep -Milliseconds 500
    }
}

foreach ($step in $Steps) {
    $pres = $deck
    if ($step -eq "/new") {
        $ai.NewConversation()
        Write-Output (@{ step = $step; done = $true } | ConvertTo-Json -Compress)
    }
    elseif ($step -match '^/cancel-after\s+(\d+)\s+(.+)$') {
        $ai.Send($Matches[2])
        Start-Sleep -Seconds ([int]$Matches[1])
        $ai.Cancel()
        Wait-Idle $TimeoutSec
        Write-Output $ai.LastRun
    }
    elseif ($step -match '^/attach\s+(.+)$') {
        # Attach a file as if dropped on the task pane, and wait until the pane has converted it.
        $path = (Resolve-Path $Matches[1].Trim()).Path
        $before = ($ai.Documents | ConvertFrom-Json).failed.Count
        $ai.Attach($path)
        $start = Get-Date
        do {
            Start-Sleep -Milliseconds 300
            $state = $ai.Documents | ConvertFrom-Json
            if (((Get-Date) - $start).TotalSeconds -gt 120) { throw "attach of $path still converting after 120 s" }
        } while ($state.converting.Count -gt 0)
        if ($state.failed.Count -gt $before) {
            # the later steps would run without the document: stop here
            Write-Output (@{ step = $step; failed = $state.failed[-1] } | ConvertTo-Json -Compress)
            exit 1
        }
        else { Write-Output ($state.documents[-1] | ConvertTo-Json -Compress) }
    }
    elseif ($step -eq "/continue") {
        $ai.ContinueStopped()
        Wait-Idle $TimeoutSec
        Write-Output $ai.LastRun
    }
    elseif ($step -match '^/ps\s+(.+)$') {
        $null = Invoke-Expression $Matches[1]
        Write-Output (@{ step = $step; done = $true } | ConvertTo-Json -Compress)
    }
    else {
        $ai.Send($step)
        Wait-Idle $TimeoutSec
        Write-Output $ai.LastRun
    }
}
