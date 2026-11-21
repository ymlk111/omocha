# screenshot.ps1 - Lightweight screenshot tool (single file)
# Run: powershell -ExecutionPolicy Bypass -File .\screenshot.ps1

# ============================================================
# ASSEMBLIES
# ============================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================
# NATIVE APIs
# ============================================================
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Text;

public class Win32 {
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("dwmapi.dll")]
    public static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out RECT pvAttribute, int cbAttribute);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left; public int Top; public int Right; public int Bottom;
    }

    public static Rectangle GetWindowBounds(IntPtr hWnd) {
        RECT rect;
        int hr = DwmGetWindowAttribute(hWnd, 9, out rect, Marshal.SizeOf(typeof(RECT)));
        if (hr != 0) { GetWindowRect(hWnd, out rect); }
        return new Rectangle(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);
    }

    public static string GetWindowTitle(IntPtr hWnd) {
        int len = GetWindowTextLength(hWnd);
        if (len == 0) return "";
        var sb = new StringBuilder(len + 1);
        GetWindowText(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }
}
"@ -ReferencedAssemblies System.Drawing

Add-Type -TypeDefinition @"
using System;
using System.Windows.Forms;

public class HotkeyMessageFilter : IMessageFilter {
    public event EventHandler<int> HotkeyPressed;
    public bool PreFilterMessage(ref Message m) {
        if (m.Msg == 0x0312) {
            int id = m.WParam.ToInt32();
            if (HotkeyPressed != null) HotkeyPressed(this, id);
            return true;
        }
        return false;
    }
}
"@ -ReferencedAssemblies System.Windows.Forms

# ============================================================
# GLOBALS
# ============================================================
$script:ScriptRoot = $PSScriptRoot
if (-not $script:ScriptRoot) { $script:ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition }

$script:ConfigPath = Join-Path $script:ScriptRoot "config.json"
$script:Config = $null
$script:RegisteredHotkeys = @{}
$script:NextHotkeyId = 1
$script:LastCaptureTime = [DateTime]::MinValue
$script:MainForm = $null
$script:HistoryPanel = $null
$script:CapturedHotkey = ""

# ============================================================
# CONFIG
# ============================================================
function Get-DefaultConfig {
    return [PSCustomObject]@{
        hotkey_capture = "Ctrl+Shift+1"
        save_path      = ".\history"
        window_x       = -1
        window_y       = -1
        cooldown_ms    = 500
    }
}

function Initialize-Config {
    if (Test-Path $script:ConfigPath) {
        $json = Get-Content $script:ConfigPath -Raw -Encoding UTF8
        $script:Config = $json | ConvertFrom-Json
    }
    else {
        $script:Config = Get-DefaultConfig
        Save-Config
    }
    $defaults = Get-DefaultConfig
    $changed = $false
    foreach ($prop in $defaults.PSObject.Properties) {
        if ($script:Config.PSObject.Properties.Name -notcontains $prop.Name) {
            $script:Config | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value
            $changed = $true
        }
    }
    if ($changed) { Save-Config }
}

function Save-Config {
    $script:Config | ConvertTo-Json -Depth 5 | Set-Content $script:ConfigPath -Encoding UTF8
}

function Set-ConfigValue {
    param([string]$Key, $Value)
    if ($script:Config.PSObject.Properties.Name -contains $Key) { $script:Config.$Key = $Value }
    else { $script:Config | Add-Member -NotePropertyName $Key -NotePropertyValue $Value }
    Save-Config
}

function Resolve-SavePath {
    $path = $script:Config.save_path
    if (-not [System.IO.Path]::IsPathRooted($path)) { $path = Join-Path $script:ScriptRoot $path }
    $path = [System.IO.Path]::GetFullPath($path)
    if (-not (Test-Path $path)) { New-Item -ItemType Directory -Path $path -Force | Out-Null }
    return $path
}

function Parse-HotkeyString {
    param([string]$HotkeyStr)
    $MOD_ALT = 0x0001; $MOD_CONTROL = 0x0002; $MOD_SHIFT = 0x0004; $MOD_WIN = 0x0008
    $parts = $HotkeyStr -split '\+'
    $modifiers = 0; $keyName = ""
    foreach ($part in $parts) {
        $p = $part.Trim()
        switch ($p.ToLower()) {
            "ctrl"    { $modifiers = $modifiers -bor $MOD_CONTROL }
            "control" { $modifiers = $modifiers -bor $MOD_CONTROL }
            "shift"   { $modifiers = $modifiers -bor $MOD_SHIFT }
            "alt"     { $modifiers = $modifiers -bor $MOD_ALT }
            "win"     { $modifiers = $modifiers -bor $MOD_WIN }
            default   { $keyName = $p }
        }
    }
    $keyCode = [System.Windows.Forms.Keys]::None
    if ($keyName -match '^\d$') { $keyCode = [System.Windows.Forms.Keys]"D$keyName" }
    elseif ($keyName -match '^[Ff]\d+$') { $keyCode = [System.Windows.Forms.Keys]($keyName.ToUpper()) }
    elseif ($keyName -match '^[A-Za-z]$') { $keyCode = [System.Windows.Forms.Keys]($keyName.ToUpper()) }
    elseif ($keyName.ToLower() -eq "printscreen" -or $keyName.ToLower() -eq "prtsc") { $keyCode = [System.Windows.Forms.Keys]::PrintScreen }
    elseif ($keyName.ToLower() -eq "space") { $keyCode = [System.Windows.Forms.Keys]::Space }
    elseif ($keyName.ToLower() -eq "enter") { $keyCode = [System.Windows.Forms.Keys]::Enter }
    elseif ($keyName.ToLower() -eq "escape") { $keyCode = [System.Windows.Forms.Keys]::Escape }
    return @{ Modifiers = $modifiers; KeyCode = $keyCode }
}

# ============================================================
# HOTKEY
# ============================================================
function Register-GlobalHotkey {
    param([IntPtr]$WindowHandle, [int]$Modifiers, [System.Windows.Forms.Keys]$KeyCode, [string]$Name)
    $id = $script:NextHotkeyId; $script:NextHotkeyId++
    $result = [Win32]::RegisterHotKey($WindowHandle, $id, [uint32]$Modifiers, [uint32]$KeyCode)
    if ($result) {
        $script:RegisteredHotkeys[$id] = $Name
        return $id
    }
    else { Write-Warning "Failed to register hotkey: $Name"; return -1 }
}

function Unregister-AllHotkeys {
    param([IntPtr]$WindowHandle)
    foreach ($id in @($script:RegisteredHotkeys.Keys)) {
        [Win32]::UnregisterHotKey($WindowHandle, $id) | Out-Null
    }
    $script:RegisteredHotkeys.Clear(); $script:NextHotkeyId = 1
}

function Register-ConfiguredHotkeys {
    param([System.Windows.Forms.Form]$Form)
    Unregister-AllHotkeys -WindowHandle $Form.Handle
    $hk = Parse-HotkeyString -HotkeyStr $script:Config.hotkey_capture
    if ($hk.KeyCode -ne [System.Windows.Forms.Keys]::None) {
        Register-GlobalHotkey -WindowHandle $Form.Handle -Modifiers $hk.Modifiers -KeyCode $hk.KeyCode -Name "Capture" | Out-Null
    }
}

# ============================================================
# CAPTURE
# ============================================================
function Invoke-Capture {
    param([string]$SaveDir)
    try {
        $hwnd = [Win32]::GetForegroundWindow()
        $bounds = [Win32]::GetWindowBounds($hwnd)
        if ($bounds.Width -le 0 -or $bounds.Height -le 0) { return $null }

        $bitmap = New-Object System.Drawing.Bitmap($bounds.Width, $bounds.Height)
        $g = [System.Drawing.Graphics]::FromImage($bitmap)
        $g.CopyFromScreen($bounds.Left, $bounds.Top, 0, 0, $bitmap.Size)
        $g.Dispose()

        [System.Windows.Forms.Clipboard]::SetImage($bitmap)

        if (-not (Test-Path $SaveDir)) { New-Item -ItemType Directory -Path $SaveDir -Force | Out-Null }
        $ts = Get-Date -Format "yyyyMMdd_HHmmss_fff"
        $filePath = Join-Path $SaveDir "${ts}.png"
        $bitmap.Save($filePath, [System.Drawing.Imaging.ImageFormat]::Png)

        return @{ FilePath = $filePath; Bitmap = $bitmap }
    }
    catch { Write-Warning "Capture failed: $_"; return $null }
}

# ============================================================
# HISTORY
# ============================================================
function Get-HistoryItems {
    param([string]$SaveDir, [int]$MaxItems = 30)
    if (-not (Test-Path $SaveDir)) { return @() }
    return @(Get-ChildItem -Path $SaveDir -Filter "*.png" -File | Sort-Object LastWriteTime -Descending | Select-Object -First $MaxItems)
}

function New-Thumbnail {
    param([string]$FilePath, [int]$MaxSize = 85)
    try {
        $orig = [System.Drawing.Image]::FromFile($FilePath)
        $ratio = [Math]::Min($MaxSize / $orig.Width, $MaxSize / $orig.Height)
        $w = [Math]::Max(1, [int]($orig.Width * $ratio))
        $h = [Math]::Max(1, [int]($orig.Height * $ratio))
        $thumb = New-Object System.Drawing.Bitmap($w, $h)
        $g = [System.Drawing.Graphics]::FromImage($thumb)
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.DrawImage($orig, 0, 0, $w, $h)
        $g.Dispose(); $orig.Dispose()
        return $thumb
    }
    catch { return $null }
}

function Copy-HistoryToClipboard {
    param([string]$FilePath)
    try {
        $img = [System.Drawing.Image]::FromFile($FilePath)
        [System.Windows.Forms.Clipboard]::SetImage($img)
        $img.Dispose()
    }
    catch {}
}

function Update-HistoryPanel {
    if ($script:HistoryPanel -eq $null) { return }
    $script:HistoryPanel.Controls.Clear()
    $saveDir = Resolve-SavePath
    $items = Get-HistoryItems -SaveDir $saveDir
    foreach ($item in $items) {
        $pb = New-Object System.Windows.Forms.PictureBox
        $pb.Size = New-Object System.Drawing.Size(85, 65)
        $pb.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
        $pb.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
        $pb.Cursor = [System.Windows.Forms.Cursors]::Hand
        $pb.Tag = $item.FullName
        $pb.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $thumb = New-Thumbnail -FilePath $item.FullName
        if ($thumb -ne $null) { $pb.Image = $thumb }
        $tt = New-Object System.Windows.Forms.ToolTip
        $tt.SetToolTip($pb, ($item.Name + " - " + $item.LastWriteTime.ToString("yyyy/MM/dd HH:mm:ss")))
        $pb.Add_Click({ param($sender, $e); Copy-HistoryToClipboard -FilePath $sender.Tag })
        $script:HistoryPanel.Controls.Add($pb)
    }
}

# ============================================================
# GUI
# ============================================================
function Build-GUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "ScreenShot Tool"
    $form.Size = New-Object System.Drawing.Size(320, 380)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
    $form.TopMost = $true
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $form.ForeColor = [System.Drawing.Color]::White
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $script:MainForm = $form

    if ($script:Config.window_x -ge 0 -and $script:Config.window_y -ge 0) {
        $form.Location = New-Object System.Drawing.Point($script:Config.window_x, $script:Config.window_y)
    }
    else {
        $scr = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $form.Location = New-Object System.Drawing.Point(($scr.Right - 340), ($scr.Bottom - 400))
    }

    $y = 10

    # Capture button
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = "Capture"
    $btn.Location = New-Object System.Drawing.Point(10, $y)
    $btn.Size = New-Object System.Drawing.Size(290, 40)
    $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btn.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btn.Add_Click({
        $script:MainForm.Visible = $false
        Start-Sleep -Milliseconds 200
        $result = Invoke-Capture -SaveDir (Resolve-SavePath)
        $script:MainForm.Visible = $true
        if ($result -ne $null) { Update-HistoryPanel; $result.Bitmap.Dispose() }
    })
    $form.Controls.Add($btn)
    $y += 50

    # Separator
    $sep = New-Object System.Windows.Forms.Label; $sep.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $sep.Location = New-Object System.Drawing.Point(10, $y); $sep.Size = New-Object System.Drawing.Size(290, 2)
    $form.Controls.Add($sep); $y += 10

    # Settings header
    $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "Settings"
    $lbl.Location = New-Object System.Drawing.Point(10, $y); $lbl.Size = New-Object System.Drawing.Size(290, 20)
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lbl); $y += 22

    # Hotkey row
    $lblH = New-Object System.Windows.Forms.Label; $lblH.Text = "Hotkey:"
    $lblH.Location = New-Object System.Drawing.Point(10, $y); $lblH.Size = New-Object System.Drawing.Size(55, 22)
    $form.Controls.Add($lblH)

    $script:TxtHotkey = New-Object System.Windows.Forms.TextBox
    $script:TxtHotkey.Text = $script:Config.hotkey_capture
    $script:TxtHotkey.Location = New-Object System.Drawing.Point(70, $y)
    $script:TxtHotkey.Size = New-Object System.Drawing.Size(140, 22)
    $script:TxtHotkey.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
    $script:TxtHotkey.ForeColor = [System.Drawing.Color]::White
    $script:TxtHotkey.ReadOnly = $true
    $form.Controls.Add($script:TxtHotkey)

    $btnHk = New-Object System.Windows.Forms.Button; $btnHk.Text = "Set"
    $btnHk.Location = New-Object System.Drawing.Point(215, $y); $btnHk.Size = New-Object System.Drawing.Size(85, 22)
    $btnHk.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnHk.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $btnHk.ForeColor = [System.Drawing.Color]::White
    $btnHk.Add_Click({
        $newKey = Show-HotkeyDialog -CurrentKey $script:TxtHotkey.Text
        if ($newKey -ne $null) {
            $script:TxtHotkey.Text = $newKey
            Set-ConfigValue -Key "hotkey_capture" -Value $newKey
            Register-ConfiguredHotkeys -Form $script:MainForm
        }
    })
    $form.Controls.Add($btnHk); $y += 28

    # Save path row
    $lblS = New-Object System.Windows.Forms.Label; $lblS.Text = "Save to:"
    $lblS.Location = New-Object System.Drawing.Point(10, $y); $lblS.Size = New-Object System.Drawing.Size(55, 22)
    $form.Controls.Add($lblS)

    $script:TxtPath = New-Object System.Windows.Forms.TextBox
    $script:TxtPath.Text = $script:Config.save_path
    $script:TxtPath.Location = New-Object System.Drawing.Point(70, $y)
    $script:TxtPath.Size = New-Object System.Drawing.Size(140, 22)
    $script:TxtPath.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
    $script:TxtPath.ForeColor = [System.Drawing.Color]::White
    $script:TxtPath.ReadOnly = $true
    $form.Controls.Add($script:TxtPath)

    $btnBr = New-Object System.Windows.Forms.Button; $btnBr.Text = "Browse..."
    $btnBr.Location = New-Object System.Drawing.Point(215, $y); $btnBr.Size = New-Object System.Drawing.Size(85, 22)
    $btnBr.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnBr.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $btnBr.ForeColor = [System.Drawing.Color]::White
    $btnBr.Add_Click({
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.SelectedPath = Resolve-SavePath
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:TxtPath.Text = $fbd.SelectedPath
            Set-ConfigValue -Key "save_path" -Value $fbd.SelectedPath
            Update-HistoryPanel
        }
    })
    $form.Controls.Add($btnBr); $y += 35

    # Separator
    $sep2 = New-Object System.Windows.Forms.Label; $sep2.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $sep2.Location = New-Object System.Drawing.Point(10, $y); $sep2.Size = New-Object System.Drawing.Size(290, 2)
    $form.Controls.Add($sep2); $y += 10

    # History header + open folder
    $lblHist = New-Object System.Windows.Forms.Label; $lblHist.Text = "History (click to copy)"
    $lblHist.Location = New-Object System.Drawing.Point(10, $y); $lblHist.Size = New-Object System.Drawing.Size(200, 20)
    $lblHist.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblHist)

    $btnOpen = New-Object System.Windows.Forms.Button; $btnOpen.Text = "Open Folder"
    $btnOpen.Location = New-Object System.Drawing.Point(210, $y); $btnOpen.Size = New-Object System.Drawing.Size(90, 20)
    $btnOpen.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOpen.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $btnOpen.ForeColor = [System.Drawing.Color]::White
    $btnOpen.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $btnOpen.Add_Click({ Start-Process explorer.exe (Resolve-SavePath) })
    $form.Controls.Add($btnOpen); $y += 25

    # History panel
    $script:HistoryPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $script:HistoryPanel.Location = New-Object System.Drawing.Point(10, $y)
    $script:HistoryPanel.Size = New-Object System.Drawing.Size(290, 150)
    $script:HistoryPanel.AutoScroll = $true
    $script:HistoryPanel.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $script:HistoryPanel.WrapContents = $true
    $script:HistoryPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $form.Controls.Add($script:HistoryPanel)

    Update-HistoryPanel

    # Save position on move
    $form.Add_LocationChanged({
        Set-ConfigValue -Key "window_x" -Value $script:MainForm.Location.X
        Set-ConfigValue -Key "window_y" -Value $script:MainForm.Location.Y
    })

    return $form
}

function Show-HotkeyDialog {
    param([string]$CurrentKey)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Set Hotkey"; $dlg.Size = New-Object System.Drawing.Size(320, 160)
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $dlg.ForeColor = [System.Drawing.Color]::White
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false; $dlg.TopMost = $true

    $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "Press desired key combination..."
    $lbl.Location = New-Object System.Drawing.Point(10, 15); $lbl.Size = New-Object System.Drawing.Size(290, 20)
    $dlg.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox; $txt.Text = $CurrentKey
    $txt.Location = New-Object System.Drawing.Point(10, 45); $txt.Size = New-Object System.Drawing.Size(285, 25)
    $txt.ReadOnly = $true
    $txt.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
    $txt.ForeColor = [System.Drawing.Color]::LimeGreen
    $txt.Font = New-Object System.Drawing.Font("Consolas", 12)
    $txt.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
    $dlg.Controls.Add($txt)

    $script:CapturedHotkey = $CurrentKey

    $txt.Add_KeyDown({
        param($sender, $e)
        $e.SuppressKeyPress = $true
        $parts = New-Object System.Collections.ArrayList
        if ($e.Control) { [void]$parts.Add("Ctrl") }
        if ($e.Shift)   { [void]$parts.Add("Shift") }
        if ($e.Alt)     { [void]$parts.Add("Alt") }
        $k = $e.KeyCode.ToString()
        if ($k -ne "ControlKey" -and $k -ne "ShiftKey" -and $k -ne "Menu") {
            if ($k -match '^D(\d)$') { $k = $Matches[1] }
            [void]$parts.Add($k)
            $combo = $parts -join "+"
            $sender.Text = $combo
            $script:CapturedHotkey = $combo
        }
    })

    $btnOk = New-Object System.Windows.Forms.Button; $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(80, 85); $btnOk.Size = New-Object System.Drawing.Size(70, 28)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnOk.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOk.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnOk.ForeColor = [System.Drawing.Color]::White
    $dlg.Controls.Add($btnOk)

    $btnCa = New-Object System.Windows.Forms.Button; $btnCa.Text = "Cancel"
    $btnCa.Location = New-Object System.Drawing.Point(160, 85); $btnCa.Size = New-Object System.Drawing.Size(70, 28)
    $btnCa.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $btnCa.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCa.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $btnCa.ForeColor = [System.Drawing.Color]::White
    $dlg.Controls.Add($btnCa)

    $dlg.AcceptButton = $btnOk; $dlg.CancelButton = $btnCa
    $txt.Select()

    $dr = $dlg.ShowDialog()
    $key = $script:CapturedHotkey
    $dlg.Dispose()

    if ($dr -eq [System.Windows.Forms.DialogResult]::OK) { return $key }
    return $null
}

# ============================================================
# MAIN
# ============================================================
Initialize-Config
$form = Build-GUI
Register-ConfiguredHotkeys -Form $form

$msgFilter = New-Object HotkeyMessageFilter
$msgFilter.Add_HotkeyPressed({
    param($sender, $hotkeyId)

    # Cooldown to prevent repeated captures on key hold
    $cooldown = [int]$script:Config.cooldown_ms
    $now = [DateTime]::Now
    if (($now - $script:LastCaptureTime).TotalMilliseconds -lt $cooldown) { return }
    $script:LastCaptureTime = $now

    # Only respond to our registered hotkey
    if (-not $script:RegisteredHotkeys.ContainsKey($hotkeyId)) { return }

    $script:MainForm.Visible = $false
    Start-Sleep -Milliseconds 150

    $result = Invoke-Capture -SaveDir (Resolve-SavePath)

    $script:MainForm.Visible = $true

    if ($result -ne $null) { Update-HistoryPanel; $result.Bitmap.Dispose() }
})

[System.Windows.Forms.Application]::AddMessageFilter($msgFilter)

$form.Add_FormClosing({
    Unregister-AllHotkeys -WindowHandle $script:MainForm.Handle
    [System.Windows.Forms.Application]::RemoveMessageFilter($msgFilter)
})

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($form)
