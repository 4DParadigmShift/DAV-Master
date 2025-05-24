Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing

# App/Config/data location logic
$ConfigFolder = "$env:APPDATA\\DAV_Transport"
$LocationFile = "$ConfigFolder\\location.txt"
if (-not (Test-Path $ConfigFolder)) { New-Item -ItemType Directory -Path $ConfigFolder -Force | Out-Null }
function Get-DataDir {
    if (Test-Path $LocationFile) {
        $path = Get-Content $LocationFile -Raw
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path $path)) { return $path }
    }
    return "$env:LOCALAPPDATA\\DAV_Transport"
}
function Set-DataDir($newPath) {
    $newPath = $newPath.Trim()
    if (-not (Test-Path $newPath)) { New-Item -ItemType Directory -Path $newPath -Force | Out-Null }
    Set-Content $LocationFile $newPath
}

$AppDir = Get-DataDir

# GUARANTEE DATA DIR EXISTS FOR ALL MODES (local or network)
if (-not (Test-Path $AppDir)) {
    New-Item -ItemType Directory -Path $AppDir -Force | Out-Null
}

$DriversFile = "$AppDir\\drivers.csv"
$ClientsFile = "$AppDir\\clients.csv"
$DestFile    = "$AppDir\\destinations.csv"
$ApptFile    = "$AppDir\\appts.csv"
$LogFile     = "$AppDir\\dispatch.log"

# Create the files if they do not exist yet
if (-not (Test-Path $DriversFile)) {
    [PSCustomObject]@{Name=''; Address=''; Phone=''; State='FL'; Email=''} | Export-Csv $DriversFile -NoTypeInformation -Encoding utf8
}
if (-not (Test-Path $ClientsFile)) {
    [PSCustomObject]@{Name=''; Address=''; Phone=''; State='FL'; Email=''} | Export-Csv $ClientsFile -NoTypeInformation -Encoding utf8
}
if (-not (Test-Path $DestFile))    { "Name,Address,State,Notes"       | Out-File $DestFile -Encoding utf8 }
if (-not (Test-Path $ApptFile))    { "Driver,Client,Destination,Date,Time,State,Notes" | Out-File $ApptFile -Encoding utf8 }
if (-not (Test-Path $LogFile))     { New-Item -Path $LogFile -ItemType File | Out-Null }

$AppName = "DAV Dispatch"
$PrimaryColor = "#5b5c5a"
$AccentColor  = "#bed12b"
$FontFamily   = "Arial"
$HeaderFont   = 42
$LabelFont    = 26
$ButtonFont   = 32
$StateList = @("AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN","IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN","TX","UT","VT","VA","WA","WV","WI","WY")

function Get-Drivers {
    if (-not (Test-Path $DriversFile)) {
        [PSCustomObject]@{Name=''; Address=''; Phone=''; State='FL'; Email=''} | Export-Csv $DriversFile -NoTypeInformation -Encoding utf8
    }
    Import-Csv $DriversFile | Where-Object { $_.Name -ne "" }
}
function Set-Drivers($x) {
    if ($x.Count -eq 0) {
        [PSCustomObject]@{Name=''; Address=''; Phone=''; State='FL'; Email=''} | Export-Csv $DriversFile -NoTypeInformation -Encoding utf8
    } else {
        $x | Export-Csv $DriversFile -NoTypeInformation -Encoding utf8
    }
}
function Get-Clients {
    if (-not (Test-Path $ClientsFile)) {
        [PSCustomObject]@{Name=''; Address=''; Phone=''; State='FL'; Email=''} | Export-Csv $ClientsFile -NoTypeInformation -Encoding utf8
    }
    Import-Csv $ClientsFile | Where-Object { $_.Name -ne "" }
}
function Set-Clients($x) {
    if ($x.Count -eq 0) {
        [PSCustomObject]@{Name=''; Address=''; Phone=''; State='FL'; Email=''} | Export-Csv $ClientsFile -NoTypeInformation -Encoding utf8
    } else {
        $x | Export-Csv $ClientsFile -NoTypeInformation -Encoding utf8
    }
}
function Get-Dest {
    if (-not (Test-Path $DestFile)) {
        "Name,Address,State,Notes" | Out-File $DestFile -Encoding utf8
    }
    Import-Csv $DestFile | Where-Object { $_.Name -ne "" }
}
function Set-Dest($x) {
    if ($x.Count -eq 0) {
        "Name,Address,State,Notes" | Out-File $DestFile -Encoding utf8
    } else {
        $x | Export-Csv $DestFile -NoTypeInformation -Encoding utf8
    }
}
function Get-Appts {
    if (-not (Test-Path $ApptFile)) {
        "Driver,Client,Destination,Date,Time,State,Notes" | Out-File $ApptFile -Encoding utf8
    }
    Import-Csv $ApptFile | Where-Object { $_.Client -ne "" }
}
function Set-Appts($x) {
    if ($x.Count -eq 0) {
        "Driver,Client,Destination,Date,Time,State,Notes" | Out-File $ApptFile -Encoding utf8
    } else {
        $x | Export-Csv $ApptFile -NoTypeInformation -Encoding utf8
    }
}

function Read-StylizedInput($prompt, $title="Input") {
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.WindowState = 'Maximized'
    $form.TopMost = $true
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml($PrimaryColor)
    $form.FormBorderStyle = 'None'

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Width = 900
    $panel.Height = 330
    $panel.Anchor = 'None'
    $form.Controls.Add($panel)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $prompt
    $lbl.Font = New-Object System.Drawing.Font($FontFamily, 34)
    $lbl.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lbl.Width = 800
    $lbl.Height = 70
    $lbl.Left = [math]::Max(0, ($panel.Width - $lbl.Width) / 2)
    $lbl.Top = 45
    $lbl.TextAlign = 'MiddleCenter'
    $panel.Controls.Add($lbl)

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Font = New-Object System.Drawing.Font($FontFamily, 34)
    $tb.Width = 800
    $tb.Height = 56
    $tb.Left = [math]::Max(0, ($panel.Width - $tb.Width) / 2)
    $tb.Top = 120
    $panel.Controls.Add($tb)

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = "OK"
    $btn.Font = New-Object System.Drawing.Font($FontFamily, 32)
    $btn.Width = 220
    $btn.Height = 70
    $btn.Left = [math]::Max(0, ($panel.Width - $btn.Width) / 2)
    $btn.Top = 200
    $btn.BackColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $btn.ForeColor = "Black"
    $btn.FlatStyle = "Flat"
    $btn.Add_Click({ $form.Tag = $tb.Text; $form.Close() })
    $panel.Controls.Add($btn)

    $tb.Add_KeyDown({
        if ($_.KeyCode -eq 'Enter') {
            $form.Tag = $tb.Text
            $form.Close()
        }
    })

    function Center-Panel {
        $panel.Left = [Math]::Max(0, ($form.ClientSize.Width - $panel.Width) / 2)
        $panel.Top  = [Math]::Max(0, ($form.ClientSize.Height - $panel.Height) / 2)
    }
    $form.Add_Resize({ Center-Panel })
    Center-Panel

    $form.Add_FormClosing({
        if ($form.Tag -eq $null) { $form.Tag = "" }
    })
    $form.ShowDialog() | Out-Null
    return $form.Tag
}

function Show-StylizedMessage($msg, $title = $AppName) {
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.WindowState = 'Maximized'
    $form.TopMost = $true
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml($PrimaryColor)
    $form.FormBorderStyle = 'None'

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Width = 1100; $panel.Height = 370
    $panel.Anchor = 'None'
    $form.Controls.Add($panel)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $msg
    $lbl.Font = New-Object System.Drawing.Font($FontFamily, 40)
    $lbl.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lbl.Width = 1000
    $lbl.Height = 180
    $lbl.Left = [math]::Max(0, ($panel.Width - $lbl.Width) / 2)
    $lbl.Top = 70
    $lbl.AutoSize = $false
    $lbl.TextAlign = 'MiddleCenter'
    $panel.Controls.Add($lbl)

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = "OK"
    $btn.Font = New-Object System.Drawing.Font($FontFamily, 34)
    $btn.Width = 250; $btn.Height = 82
    $btn.Left = [math]::Max(0, ($panel.Width - $btn.Width) / 2)
    $btn.Top = 255
    $btn.BackColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $btn.ForeColor = "Black"
    $btn.FlatStyle = "Flat"
    $btn.Add_Click({ $form.Close() })
    $panel.Controls.Add($btn)

    function Center-Panel {
        $panel.Left = [Math]::Max(0, ($form.ClientSize.Width - $panel.Width) / 2)
        $panel.Top  = [Math]::Max(0, ($form.ClientSize.Height - $panel.Height) / 2)
    }
    $form.Add_Resize({ Center-Panel })
    Center-Panel

    $form.ShowDialog() | Out-Null
}

function Read-AddressDialog($title="Enter Address") {
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.WindowState = 'Maximized'
    $form.TopMost = $true
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml($PrimaryColor)
    $form.FormBorderStyle = 'None'

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Width = 1100
    $panel.Height = 700
    $panel.Anchor = 'None'
    $form.Controls.Add($panel)

    $fields = @("Street","City","Zip")
    $tbs = @{ }
    $y = 70
    foreach ($f in $fields) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = "${f}:"
        $lbl.Font = New-Object System.Drawing.Font($FontFamily, 36)
        $lbl.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
        $lbl.Width = 280
        $lbl.Height = 56
        $lbl.Left = 60
        $lbl.Top = $y
        $panel.Controls.Add($lbl)
        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Font = New-Object System.Drawing.Font($FontFamily, 36)
        $tb.Width = 660
        $tb.Height = 56
        $tb.Left = 340
        $tb.Top = $y
        $panel.Controls.Add($tb)
        $tbs[$f] = $tb
        $y += 105
    }
    # State dropdown
    $lblState = New-Object System.Windows.Forms.Label
    $lblState.Text = "State:"
    $lblState.Font = New-Object System.Drawing.Font($FontFamily, 36)
    $lblState.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lblState.Width = 280
    $lblState.Height = 56
    $lblState.Left = 60
    $lblState.Top = $y
    $panel.Controls.Add($lblState)
    $cbState = New-Object System.Windows.Forms.ComboBox
    $cbState.Font = New-Object System.Drawing.Font($FontFamily, 36)
    $cbState.Width = 400
    $cbState.Height = 56
    $cbState.Left = 340
    $cbState.Top = $y
    foreach ($s in $StateList) { $cbState.Items.Add($s) | Out-Null }
    $cbState.SelectedItem = "FL"
    $panel.Controls.Add($cbState)
    $y += 105

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = "OK"
    $btn.Font = New-Object System.Drawing.Font($FontFamily, 38)
    $btn.Width = 250
    $btn.Height = 82
    $btn.Left = [math]::Max(0, ($panel.Width - $btn.Width) / 2)
    $btn.Top = $y + 10
    $btn.BackColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $btn.ForeColor = "Black"
    $btn.FlatStyle = "Flat"
    $btn.Add_Click({
        $address = "$($tbs['Street'].Text), $($tbs['City'].Text), $($cbState.Text) $($tbs['Zip'].Text)"
        $form.Tag = $address + "||" + $cbState.Text
        $form.Close()
    })
    $panel.Controls.Add($btn)

    foreach ($tb in $tbs.Values) {
        $tb.Add_KeyDown({
            if ($_.KeyCode -eq 'Enter') {
                $address = "$($tbs['Street'].Text), $($tbs['City'].Text), $($cbState.Text) $($tbs['Zip'].Text)"
                $form.Tag = $address + "||" + $cbState.Text
                $form.Close()
            }
        })
    }
    $cbState.Add_KeyDown({
        if ($_.KeyCode -eq 'Enter') {
            $address = "$($tbs['Street'].Text), $($tbs['City'].Text), $($cbState.Text) $($tbs['Zip'].Text)"
            $form.Tag = $address + "||" + $cbState.Text
            $form.Close()
        }
    })

    function Center-Panel {
        $panel.Left = [Math]::Max(0, ($form.ClientSize.Width - $panel.Width) / 2)
        $panel.Top  = [Math]::Max(0, ($form.ClientSize.Height - $panel.Height) / 2)
    }
    $form.Add_Resize({ Center-Panel })
    Center-Panel

    $form.Add_FormClosing({
        if ($form.Tag -eq $null) { $form.Tag = "" }
    })
    $form.ShowDialog() | Out-Null
    return $form.Tag
}

function Read-StateDialog($prompt="Select State:", $default="FL", $title="Pick State") {
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.WindowState = 'Maximized'
    $form.TopMost = $true
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml($PrimaryColor)
    $form.FormBorderStyle = 'None'

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Width = 900
    $panel.Height = 330
    $panel.Anchor = 'None'
    $form.Controls.Add($panel)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $prompt
    $lbl.Font = New-Object System.Drawing.Font($FontFamily, 34)
    $lbl.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lbl.Width = 800
    $lbl.Height = 70
    $lbl.Left = [math]::Max(0, ($panel.Width - $lbl.Width) / 2)
    $lbl.Top = 45
    $lbl.TextAlign = 'MiddleCenter'
    $panel.Controls.Add($lbl)

    $cbState = New-Object System.Windows.Forms.ComboBox
    $cbState.Font = New-Object System.Drawing.Font($FontFamily, 34)
    $cbState.Width = 400
    $cbState.Height = 56
    $cbState.Left = [math]::Max(0, ($panel.Width - $cbState.Width) / 2)
    $cbState.Top = 120
    foreach ($s in $StateList) { $cbState.Items.Add($s) | Out-Null }
    $cbState.SelectedItem = $default
    $panel.Controls.Add($cbState)

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = "OK"
    $btn.Font = New-Object System.Drawing.Font($FontFamily, 32)
    $btn.Width = 220
    $btn.Height = 70
    $btn.Left = [math]::Max(0, ($panel.Width - $btn.Width) / 2)
    $btn.Top = 200
    $btn.BackColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $btn.ForeColor = "Black"
    $btn.FlatStyle = "Flat"
    $btn.Add_Click({ $form.Tag = $cbState.Text; $form.Close() })
    $panel.Controls.Add($btn)

    $cbState.Add_KeyDown({
        if ($_.KeyCode -eq 'Enter') {
            $form.Tag = $cbState.Text
            $form.Close()
        }
    })

    function Center-Panel {
        $panel.Left = [Math]::Max(0, ($form.ClientSize.Width - $panel.Width) / 2)
        $panel.Top  = [Math]::Max(0, ($form.ClientSize.Height - $panel.Height) / 2)
    }
    $form.Add_Resize({ Center-Panel })
    Center-Panel

    $form.Add_FormClosing({
        if ($form.Tag -eq $null) { $form.Tag = "" }
    })
    $form.ShowDialog() | Out-Null
    return $form.Tag
}

function Show-ChangeDataLocation {
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Change Data Location"
    $form.WindowState = 'Maximized'
    $form.TopMost = $true
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml($PrimaryColor)
    $form.FormBorderStyle = 'None'
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Width = 1000
    $panel.Height = 300
    $form.Controls.Add($panel)
    $panel.Left = [math]::Max(0, ($form.ClientSize.Width - $panel.Width) / 2)
    $panel.Top  = [math]::Max(0, ($form.ClientSize.Height - $panel.Height) / 2)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Select or enter a new shared/network folder for data:"
    $lbl.Font = New-Object System.Drawing.Font($FontFamily, 28)
    $lbl.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lbl.Left = 20; $lbl.Top = 20; $lbl.Width = 960; $lbl.Height = 56
    $panel.Controls.Add($lbl)

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Font = New-Object System.Drawing.Font($FontFamily, 24)
    $tb.Left = 40; $tb.Top = 80; $tb.Width = 800; $tb.Height = 50
    $tb.Text = Get-DataDir
    $panel.Controls.Add($tb)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "Browse"
    $btnBrowse.Font = New-Object System.Drawing.Font($FontFamily, 24)
    $btnBrowse.Width = 120
    $btnBrowse.Height = 50
    $btnBrowse.Left = 860
    $btnBrowse.Top = 80
    $btnBrowse.BackColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $btnBrowse.ForeColor = "Black"
    $btnBrowse.FlatStyle = "Flat"
    $btnBrowse.Add_Click({
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.SelectedPath = $tb.Text
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $tb.Text = $fbd.SelectedPath
        }
    })
    $panel.Controls.Add($btnBrowse)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Save"
    $btnSave.Font = New-Object System.Drawing.Font($FontFamily, 26)
    $btnSave.Width = 240
    $btnSave.Height = 60
    $btnSave.Left = 180
    $btnSave.Top = 180
    $btnSave.BackColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $btnSave.ForeColor = "Black"
    $btnSave.FlatStyle = "Flat"
    $btnSave.Add_Click({
        $newPath = $tb.Text.Trim()
        if (-not (Test-Path $newPath)) {
            try { New-Item -ItemType Directory -Path $newPath -Force | Out-Null }
            catch { Show-StylizedMessage "Could not create/access $newPath."; return }
        }
        Set-DataDir $newPath
        Show-StylizedMessage "Data location changed to:`n$newPath`nRestarting app..."
        Start-Sleep -Seconds 1
        # Relaunch (EXE-safe): use $PSScriptRoot for .exe, else PS1
        $exe = [System.IO.Path]::ChangeExtension($PSCommandPath, 'exe')
        if (Test-Path $exe) {
            Start-Process $exe
        } else {
            Start-Process powershell -ArgumentList "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        }
        exit
    })
    $panel.Controls.Add($btnSave)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = "Back"
    $btnBack.Font = New-Object System.Drawing.Font($FontFamily, 24)
    $btnBack.Width = 160
    $btnBack.Height = 60
    $btnBack.Left = 480
    $btnBack.Top = 180
    $btnBack.BackColor = "Gray"
    $btnBack.ForeColor = "White"
    $btnBack.FlatStyle = "Flat"
    $btnBack.Add_Click({ $form.Close(); Show-ManageMenu })
    $panel.Controls.Add($btnBack)

    $form.ShowDialog() | Out-Null
}

function Show-ManageMenu {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Manage Menu"
    $form.WindowState = 'Maximized'
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml($PrimaryColor)
    $form.FormBorderStyle = 'None'
    $form.TopMost = $true

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Width = 1100
    $panel.Height = 1000
    $form.Controls.Add($panel)
    $panel.Left = [math]::Max(0, ($form.ClientSize.Width - $panel.Width) / 2)
    $panel.Top  = [math]::Max(0, ($form.ClientSize.Height - $panel.Height) / 2)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Management"
    $lbl.Font = New-Object System.Drawing.Font($FontFamily, 66, [System.Drawing.FontStyle]::Bold)
    $lbl.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lbl.Left = 0; $lbl.Top = 60
    $lbl.Width = 1100
    $lbl.Height = 120
    $lbl.TextAlign = 'MiddleCenter'
    $panel.Controls.Add($lbl)

    $y = 240
    foreach ($item in @(
        @{Name="Change Data Location"; Action={ $form.Close(); Show-ChangeDataLocation }},
        @{Name="Back"; Action={ $form.Close(); Show-MainMenu }}
    )) {
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = $item.Name
        $btn.Font = New-Object System.Drawing.Font($FontFamily, 50)
        $btn.Width = 800; $btn.Height = 130
        $btn.Left = 150; $btn.Top = $y
        
	# Need to switch this to an IF/ELSE, ternary operators not supported in PS
	#$btn.BackColor = $item.Name -eq "Back" ? "Gray" : [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
        #$btn.ForeColor = $item.Name -eq "Back" ? "White" : "Black"
        
	$btn.FlatStyle = "Flat"
        $btn.Add_Click($item.Action)
        $panel.Controls.Add($btn)
        $y += 180
    }
    $form.ShowDialog()
}

function Show-MainMenu {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $AppName
    $form.WindowState = 'Maximized'
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml($PrimaryColor)
    $form.FormBorderStyle = 'None'
    $form.TopMost = $true

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Width = 1400
    $panel.Height = 1400
    $panel.Anchor = 'None'
    $form.Controls.Add($panel)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $AppName
    $lbl.Font = New-Object System.Drawing.Font($FontFamily, 88, [System.Drawing.FontStyle]::Bold)
    $lbl.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lbl.Width = 1400
    $lbl.Height = 144
    $lbl.Left = 0
    $lbl.Top = 32
    $lbl.TextAlign = 'MiddleCenter'
    $panel.Controls.Add($lbl)

    $buttonY = 220
    $buttonWidth = 1100
    $buttonHeight = 132

    foreach ($pair in @(
        @{txt="Drivers"; fn={ $form.Close(); Show-DriverMgmt }},
        @{txt="Clients"; fn={ $form.Close(); Show-ClientMgmt }},
        @{txt="Destinations"; fn={ $form.Close(); Show-DestMgmt }},
        @{txt="Schedule Appointment"; fn={ $form.Close(); Show-Scheduler }},
        @{txt="Dispatch/Appointment History"; fn={ $form.Close(); Show-DispatchHistory }},
        @{txt="Manage"; fn={ $form.Close(); Show-ManageMenu }},
        @{txt="Quit"; fn={ $form.Close(); [System.Windows.Forms.Application]::Exit() }}
    )) {
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = $pair.txt
        $btn.Font = New-Object System.Drawing.Font($FontFamily, 52)
        $btn.Width = $buttonWidth; $btn.Height = $buttonHeight
        $btn.Left = [math]::Max(0, ($panel.Width - $buttonWidth) / 2)
        $btn.Top = $buttonY

# Need to switch this to an IF/ELSE, ternary operators not supported in PS
        #$btn.BackColor = $pair.txt -eq "Quit" ? "DarkRed" : [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
        #$btn.ForeColor = $pair.txt -eq "Quit" ? "White" : "Black"


        $btn.FlatStyle = "Flat"
        $btn.Anchor = 'None'
        $btn.Add_Click($pair.fn)
        $panel.Controls.Add($btn)
        $buttonY += 155
    }

    function Center-Panel {
        $panel.Left = [Math]::Max(0, ($form.ClientSize.Width - $panel.Width) / 2)
        $panel.Top  = [Math]::Max(0, ($form.ClientSize.Height - $panel.Height) / 2)
    }
    $form.Add_Resize({ Center-Panel })
    Center-Panel
    $form.ShowDialog()
}

function Show-DriverMgmt {
    $drivers = @(Get-Drivers)
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Driver Management"
    $form.WindowState = 'Maximized'
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml($PrimaryColor)
    $form.FormBorderStyle = 'None'
    $form.TopMost = $true

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Width = 1500
    $panel.Height = 950
    $panel.Anchor = 'None'
    $form.Controls.Add($panel)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Drivers"
    $lbl.Font = New-Object System.Drawing.Font($FontFamily, $HeaderFont)
    $lbl.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lbl.Width = 520
    $lbl.Height = 70
    $lbl.Left = [math]::Max(0, ($panel.Width - $lbl.Width) / 2)
    $lbl.Top = 24
    $panel.Controls.Add($lbl)

    $lv = New-Object System.Windows.Forms.ListView
    $lv.View = 'Details'
    $lv.FullRowSelect = $true
    $lv.Width = 1330
    $lv.Height = 540
    $lv.Left = [math]::Max(0, ($panel.Width - $lv.Width) / 2)
    $lv.Top = 110
    $lv.Font = New-Object System.Drawing.Font($FontFamily, 28)
    $lv.GridLines = $true
    $lv.MultiSelect = $false
    $lv.Columns.Add("Name", 220) | Out-Null
    $lv.Columns.Add("Address", 380) | Out-Null
    $lv.Columns.Add("State", 120) | Out-Null
    $lv.Columns.Add("Phone", 200) | Out-Null
    $lv.Columns.Add("Email", 340) | Out-Null
    $lv.Items.Clear()
    foreach ($d in $drivers) {
        if ($null -eq $d -or $d -is [array] -or $d -is [int] -or -not ($d -is [psobject])) { continue }
        


# Need to switch this to an IF/ELSE, ternary operators not supported in PS
#$state = ($d.State -ne $null -and $d.State -ne "") ? $d.State : "FL"


        $item = New-Object System.Windows.Forms.ListViewItem([string]$d.Name)
        $item.SubItems.Add(([string]$d.Address).Split('||')[0]) | Out-Null
        $item.SubItems.Add($state) | Out-Null
        $item.SubItems.Add([string]$d.Phone)   | Out-Null
        $item.SubItems.Add([string]$d.Email)   | Out-Null
        $lv.Items.Add($item) | Out-Null
    }
    $panel.Controls.Add($lv)

    $buttonY = 730
    $buttonWidth = 320
    $buttonHeight = 100
    $xPad = [math]::Max(0, ($panel.Width - (3 * $buttonWidth) - 2 * 30) / 2)

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = "Add Driver"
    $btnAdd.Font = New-Object System.Drawing.Font($FontFamily, $ButtonFont)
    $btnAdd.Width = $buttonWidth; $btnAdd.Height = $buttonHeight
    $btnAdd.Left = $xPad; $btnAdd.Top = $buttonY
    $btnAdd.BackColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $btnAdd.ForeColor = "Black"
    $btnAdd.FlatStyle = "Flat"
    $btnAdd.Anchor = 'None'
    $btnAdd.Add_Click({
        $form.Hide()
        $name = Read-StylizedInput "Driver Name:" "Add Driver"
        if (-not $name) { Show-StylizedMessage "Name required."; $form.Show(); return }
        $adstate = Read-AddressDialog "Driver Address"
        if (-not $adstate -or $adstate -notmatch '\|\|') { Show-StylizedMessage "Full address required."; $form.Show(); return }
        $address, $state = $adstate -split '\|\|'
        if (-not $state) { $state = "FL" }
        $phone = Read-StylizedInput "Phone:" "Add Driver"
        $email = Read-StylizedInput "Email:" "Add Driver"
        $drivers = @(Get-Drivers)
        $newDriver = [pscustomobject]@{
            Name    = "$name"
            Address = "$address"
            State   = "$state"
            Phone   = "$phone"
            Email   = "$email"
        }
        $drivers += $newDriver
        Set-Drivers $drivers
        $form.Close(); Show-DriverMgmt
    })
    $panel.Controls.Add($btnAdd)

    $btnDel = New-Object System.Windows.Forms.Button
    $btnDel.Text = "Delete Driver"
    $btnDel.Font = New-Object System.Drawing.Font($FontFamily, $ButtonFont)
    $btnDel.Width = $buttonWidth; $btnDel.Height = $buttonHeight
    $btnDel.Left = $xPad + $buttonWidth + 30; $btnDel.Top = $buttonY
    $btnDel.BackColor = "Gray"
    $btnDel.ForeColor = "White"
    $btnDel.FlatStyle = "Flat"
    $btnDel.Anchor = 'None'
    $btnDel.Add_Click({
        if (-not $lv.SelectedItems.Count) { Show-StylizedMessage "Select a driver to delete."; return }
        $idx = $lv.SelectedItems[0].Index
        $drivers = @(Get-Drivers)
        $drivers = $drivers | Where-Object { $drivers.IndexOf($_) -ne $idx }
        Set-Drivers $drivers
        $form.Close(); Show-DriverMgmt
    })
    $panel.Controls.Add($btnDel)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = "Back"
    $btnBack.Font = New-Object System.Drawing.Font($FontFamily, $ButtonFont)
    $btnBack.Width = $buttonWidth; $btnBack.Height = $buttonHeight
    $btnBack.Left = $xPad + 2*($buttonWidth + 30); $btnBack.Top = $buttonY
    $btnBack.BackColor = "Gray"
    $btnBack.ForeColor = "White"
    $btnBack.FlatStyle = "Flat"
    $btnBack.Anchor = 'None'
    $btnBack.Add_Click({ $form.Close(); Show-MainMenu })
    $panel.Controls.Add($btnBack)

    function Center-Panel {
        $panel.Left = [Math]::Max(0, ($form.ClientSize.Width - $panel.Width) / 2)
        $panel.Top  = [Math]::Max(0, ($form.ClientSize.Height - $panel.Height) / 2)
    }
    $form.Add_Resize({ Center-Panel })
    Center-Panel
    $form.ShowDialog()
}

function Show-ClientMgmt {
    $clients = @(Get-Clients)
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Client Management"
    $form.WindowState = 'Maximized'
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml($PrimaryColor)
    $form.FormBorderStyle = 'None'
    $form.TopMost = $true

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Width = 1500
    $panel.Height = 950
    $panel.Anchor = 'None'
    $form.Controls.Add($panel)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Clients"
    $lbl.Font = New-Object System.Drawing.Font($FontFamily, $HeaderFont)
    $lbl.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lbl.Width = 520
    $lbl.Height = 70
    $lbl.Left = [math]::Max(0, ($panel.Width - $lbl.Width) / 2)
    $lbl.Top = 24
    $panel.Controls.Add($lbl)

    $lv = New-Object System.Windows.Forms.ListView
    $lv.View = 'Details'
    $lv.FullRowSelect = $true
    $lv.Width = 1330
    $lv.Height = 540
    $lv.Left = [math]::Max(0, ($panel.Width - $lv.Width) / 2)
    $lv.Top = 110
    $lv.Font = New-Object System.Drawing.Font($FontFamily, 28)
    $lv.GridLines = $true
    $lv.MultiSelect = $false
    $lv.Columns.Add("Name", 220) | Out-Null
    $lv.Columns.Add("Address", 380) | Out-Null
    $lv.Columns.Add("State", 120) | Out-Null
    $lv.Columns.Add("Phone", 200) | Out-Null
    $lv.Columns.Add("Email", 340) | Out-Null
    $lv.Items.Clear()
    foreach ($c in $clients) {
        if ($null -eq $c -or $c -is [array] -or $c -is [int] -or -not ($c -is [psobject])) { continue }
        
# Need to switch this to an IF/ELSE, ternary operators not supported in PS
#$state = ($c.State -ne $null -and $c.State -ne "") ? $c.State : "FL"

        $item = New-Object System.Windows.Forms.ListViewItem([string]$c.Name)
        $item.SubItems.Add(([string]$c.Address).Split('||')[0]) | Out-Null
        $item.SubItems.Add($state) | Out-Null
        $item.SubItems.Add([string]$c.Phone)   | Out-Null
        $item.SubItems.Add([string]$c.Email)   | Out-Null
        $lv.Items.Add($item) | Out-Null
    }
    $panel.Controls.Add($lv)

    $buttonY = 730
    $buttonWidth = 320
    $buttonHeight = 100
    $xPad = [math]::Max(0, ($panel.Width - (3 * $buttonWidth) - 2 * 30) / 2)

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = "Add Client"
    $btnAdd.Font = New-Object System.Drawing.Font($FontFamily, $ButtonFont)
    $btnAdd.Width = $buttonWidth; $btnAdd.Height = $buttonHeight
    $btnAdd.Left = $xPad; $btnAdd.Top = $buttonY
    $btnAdd.BackColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $btnAdd.ForeColor = "Black"
    $btnAdd.FlatStyle = "Flat"
    $btnAdd.Anchor = 'None'
    $btnAdd.Add_Click({
        $form.Hide()
        $name = Read-StylizedInput "Client Name:" "Add Client"
        if (-not $name) { Show-StylizedMessage "Name required."; $form.Show(); return }
        $adstate = Read-AddressDialog "Client Address"
        if (-not $adstate -or $adstate -notmatch '\|\|') { Show-StylizedMessage "Full address required."; $form.Show(); return }
        $address, $state = $adstate -split '\|\|'
        if (-not $state) { $state = "FL" }
        $phone = Read-StylizedInput "Phone:" "Add Client"
        $email = Read-StylizedInput "Email:" "Add Client"
        $clients = @(Get-Clients)
        $newClient = [pscustomobject]@{
            Name    = "$name"
            Address = "$address"
            State   = "$state"
            Phone   = "$phone"
            Email   = "$email"
        }
        $clients += $newClient
        Set-Clients $clients
        $form.Close(); Show-ClientMgmt
    })
    $panel.Controls.Add($btnAdd)

    $btnDel = New-Object System.Windows.Forms.Button
    $btnDel.Text = "Delete Client"
    $btnDel.Font = New-Object System.Drawing.Font($FontFamily, $ButtonFont)
    $btnDel.Width = $buttonWidth; $btnDel.Height = $buttonHeight
    $btnDel.Left = $xPad + $buttonWidth + 30; $btnDel.Top = $buttonY
    $btnDel.BackColor = "Gray"
    $btnDel.ForeColor = "White"
    $btnDel.FlatStyle = "Flat"
    $btnDel.Anchor = 'None'
    $btnDel.Add_Click({
        if (-not $lv.SelectedItems.Count) { Show-StylizedMessage "Select a client to delete."; return }
        $idx = $lv.SelectedItems[0].Index
        $clients = @(Get-Clients)
        $clients = $clients | Where-Object { $clients.IndexOf($_) -ne $idx }
        Set-Clients $clients
        $form.Close(); Show-ClientMgmt
    })
    $panel.Controls.Add($btnDel)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = "Back"
    $btnBack.Font = New-Object System.Drawing.Font($FontFamily, $ButtonFont)
    $btnBack.Width = $buttonWidth; $btnBack.Height = $buttonHeight
    $btnBack.Left = $xPad + 2*($buttonWidth + 30); $btnBack.Top = $buttonY
    $btnBack.BackColor = "Gray"
    $btnBack.ForeColor = "White"
    $btnBack.FlatStyle = "Flat"
    $btnBack.Anchor = 'None'
    $btnBack.Add_Click({ $form.Close(); Show-MainMenu })
    $panel.Controls.Add($btnBack)

    function Center-Panel {
        $panel.Left = [Math]::Max(0, ($form.ClientSize.Width - $panel.Width) / 2)
        $panel.Top  = [Math]::Max(0, ($form.ClientSize.Height - $panel.Height) / 2)
    }
    $form.Add_Resize({ Center-Panel })
    Center-Panel
    $form.ShowDialog()
}

function Show-DestMgmt {
    $dests = @(Get-Dest)
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Destination Management"
    $form.WindowState = 'Maximized'
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml($PrimaryColor)
    $form.FormBorderStyle = 'None'
    $form.TopMost = $true

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Width = 1500
    $panel.Height = 950
    $panel.Anchor = 'None'
    $form.Controls.Add($panel)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Destinations"
    $lbl.Font = New-Object System.Drawing.Font($FontFamily, $HeaderFont)
    $lbl.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lbl.Width = 520
    $lbl.Height = 70
    $lbl.Left = [math]::Max(0, ($panel.Width - $lbl.Width) / 2)
    $lbl.Top = 24
    $panel.Controls.Add($lbl)

    $lv = New-Object System.Windows.Forms.ListView
    $lv.View = 'Details'
    $lv.FullRowSelect = $true
    $lv.Width = 1330
    $lv.Height = 540
    $lv.Left = [math]::Max(0, ($panel.Width - $lv.Width) / 2)
    $lv.Top = 110
    $lv.Font = New-Object System.Drawing.Font($FontFamily, 28)
    $lv.GridLines = $true
    $lv.MultiSelect = $false
    $lv.Columns.Add("Name", 220) | Out-Null
    $lv.Columns.Add("Address", 380) | Out-Null
    $lv.Columns.Add("State", 120) | Out-Null
    $lv.Columns.Add("Notes", 560) | Out-Null
    $lv.Items.Clear()
    foreach ($d in $dests) {
        if ($null -eq $d -or $d -is [array] -or $d -is [int] -or -not ($d -is [psobject])) { continue }
        
# Need to switch this to an IF/ELSE, ternary operators not supported in PS
#$state = ($d.State -ne $null -and $d.State -ne "") ? $d.State : "FL"

        $item = New-Object System.Windows.Forms.ListViewItem([string]$d.Name)
        $item.SubItems.Add(([string]$d.Address).Split('||')[0]) | Out-Null
        $item.SubItems.Add($state) | Out-Null
        $item.SubItems.Add([string]$d.Notes)   | Out-Null
        $lv.Items.Add($item) | Out-Null
    }
    $panel.Controls.Add($lv)

    $buttonY = 730
    $buttonWidth = 320
    $buttonHeight = 100
    $xPad = [math]::Max(0, ($panel.Width - (3 * $buttonWidth) - 2 * 30) / 2)

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = "Add Destination"
    $btnAdd.Font = New-Object System.Drawing.Font($FontFamily, $ButtonFont)
    $btnAdd.Width = $buttonWidth; $btnAdd.Height = $buttonHeight
    $btnAdd.Left = $xPad; $btnAdd.Top = $buttonY
    $btnAdd.BackColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $btnAdd.ForeColor = "Black"
    $btnAdd.FlatStyle = "Flat"
    $btnAdd.Anchor = 'None'
    $btnAdd.Add_Click({
        $form.Hide()
        $name = Read-StylizedInput "Destination Name:" "Add Destination"
        if (-not $name) { Show-StylizedMessage "Name required."; $form.Show(); return }
        $adstate = Read-AddressDialog "Destination Address"
        if (-not $adstate -or $adstate -notmatch '\|\|') { Show-StylizedMessage "Full address required."; $form.Show(); return }
        $address, $state = $adstate -split '\|\|'
        if (-not $state) { $state = "FL" }
        $notes = Read-StylizedInput "Notes:" "Add Destination"
        $dests = @(Get-Dest)
        $newDest = [pscustomobject]@{
            Name    = "$name"
            Address = "$address"
            State   = "$state"
            Notes   = "$notes"
        }
        $dests += $newDest
        Set-Dest $dests
        $form.Close(); Show-DestMgmt
    })
    $panel.Controls.Add($btnAdd)

    $btnDel = New-Object System.Windows.Forms.Button
    $btnDel.Text = "Delete Destination"
    $btnDel.Font = New-Object System.Drawing.Font($FontFamily, $ButtonFont)
    $btnDel.Width = $buttonWidth; $btnDel.Height = $buttonHeight
    $btnDel.Left = $xPad + $buttonWidth + 30; $btnDel.Top = $buttonY
    $btnDel.BackColor = "Gray"
    $btnDel.ForeColor = "White"
    $btnDel.FlatStyle = "Flat"
    $btnDel.Anchor = 'None'
    $btnDel.Add_Click({
        if (-not $lv.SelectedItems.Count) { Show-StylizedMessage "Select a destination to delete."; return }
        $idx = $lv.SelectedItems[0].Index
        $dests = @(Get-Dest)
        $dests = $dests | Where-Object { $dests.IndexOf($_) -ne $idx }
        Set-Dest $dests
        $form.Close(); Show-DestMgmt
    })
    $panel.Controls.Add($btnDel)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = "Back"
    $btnBack.Font = New-Object System.Drawing.Font($FontFamily, $ButtonFont)
    $btnBack.Width = $buttonWidth; $btnBack.Height = $buttonHeight
    $btnBack.Left = $xPad + 2*($buttonWidth + 30); $btnBack.Top = $buttonY
    $btnBack.BackColor = "Gray"
    $btnBack.ForeColor = "White"
    $btnBack.FlatStyle = "Flat"
    $btnBack.Anchor = 'None'
    $btnBack.Add_Click({ $form.Close(); Show-MainMenu })
    $panel.Controls.Add($btnBack)

    function Center-Panel {
        $panel.Left = [Math]::Max(0, ($form.ClientSize.Width - $panel.Width) / 2)
        $panel.Top  = [Math]::Max(0, ($form.ClientSize.Height - $panel.Height) / 2)
    }
    $form.Add_Resize({ Center-Panel })
    Center-Panel
    $form.ShowDialog()
}

function Show-Scheduler {
    $drivers = @(Get-Drivers)
    $clients = @(Get-Clients)
    $dests   = @(Get-Dest)
    $stateList = $StateList

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Schedule Appointment"
    $form.WindowState = 'Maximized'
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml($PrimaryColor)
    $form.FormBorderStyle = 'None'
    $form.TopMost = $true

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Width = 1550
    $panel.Height = 1100
    $panel.Anchor = 'None'
    $form.Controls.Add($panel)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Schedule New Appointment"
    $lbl.Font = New-Object System.Drawing.Font($FontFamily, 38)
    $lbl.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lbl.Width = 1200
    $lbl.Height = 65
    $lbl.Left = [math]::Max(0, ($panel.Width - $lbl.Width) / 2)
    $lbl.Top = 32
    $panel.Controls.Add($lbl)

    $fieldLabelW = 350
    $inputLeft = 480
    $fieldW = 400
    $y = 120

    # Driver
    $lblDriver = New-Object System.Windows.Forms.Label
    $lblDriver.Text = "Driver:"
    $lblDriver.Font = New-Object System.Drawing.Font($FontFamily, $LabelFont)
    $lblDriver.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lblDriver.Left = 80; $lblDriver.Top = $y; $lblDriver.Width = $fieldLabelW
    $lblDriver.Height = 40
    $panel.Controls.Add($lblDriver)
    $cbDriver = New-Object System.Windows.Forms.ComboBox
    $cbDriver.Font = New-Object System.Drawing.Font($FontFamily, $LabelFont)
    $cbDriver.Width = $fieldW
    $cbDriver.Left = $inputLeft; $cbDriver.Top = $y
    $cbDriver.Items.Clear()
    foreach ($d in $drivers) { if ($d.Name -and $d.Name -is [string]) { $cbDriver.Items.Add([string]$d.Name) | Out-Null } }
    if ($cbDriver.Items.Count -gt 0) { $cbDriver.SelectedIndex = 0 }
    $panel.Controls.Add($cbDriver)
    $y += 70

    # Client
    $lblClient = New-Object System.Windows.Forms.Label
    $lblClient.Text = "Client:"
    $lblClient.Font = New-Object System.Drawing.Font($FontFamily, $LabelFont)
    $lblClient.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lblClient.Left = 80; $lblClient.Top = $y; $lblClient.Width = $fieldLabelW
    $lblClient.Height = 40
    $panel.Controls.Add($lblClient)
    $cbClient = New-Object System.Windows.Forms.ComboBox
    $cbClient.Font = New-Object System.Drawing.Font($FontFamily, $LabelFont)
    $cbClient.Width = $fieldW
    $cbClient.Left = $inputLeft; $cbClient.Top = $y
    $cbClient.Items.Clear()
    foreach ($c in $clients) { if ($c.Name -and $c.Name -is [string]) { $cbClient.Items.Add([string]$c.Name) | Out-Null } }
    if ($cbClient.Items.Count -gt 0) { $cbClient.SelectedIndex = 0 }
    $panel.Controls.Add($cbClient)
    $y += 70

    # Destination
    $lblDest = New-Object System.Windows.Forms.Label
    $lblDest.Text = "Destination:"
    $lblDest.Font = New-Object System.Drawing.Font($FontFamily, $LabelFont)
    $lblDest.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lblDest.Left = 80; $lblDest.Top = $y; $lblDest.Width = $fieldLabelW
    $lblDest.Height = 40
    $panel.Controls.Add($lblDest)
    $cbDest = New-Object System.Windows.Forms.ComboBox
    $cbDest.Font = New-Object System.Drawing.Font($FontFamily, $LabelFont)
    $cbDest.Width = $fieldW
    $cbDest.Left = $inputLeft; $cbDest.Top = $y
    $cbDest.Items.Clear()
    foreach ($d in $dests) { if ($d.Name -and $d.Name -is [string]) { $cbDest.Items.Add([string]$d.Name) | Out-Null } }
    if ($cbDest.Items.Count -gt 0) { $cbDest.SelectedIndex = 0 }
    $panel.Controls.Add($cbDest)
    $y += 70

    # Date (DateTimePicker)
    $lblDate = New-Object System.Windows.Forms.Label
    $lblDate.Text = "Date:"
    $lblDate.Font = New-Object System.Drawing.Font($FontFamily, $LabelFont)
    $lblDate.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lblDate.Left = 80; $lblDate.Top = $y; $lblDate.Width = $fieldLabelW
    $lblDate.Height = 40
    $panel.Controls.Add($lblDate)
    $dtPicker = New-Object System.Windows.Forms.DateTimePicker
    $dtPicker.Font = New-Object System.Drawing.Font($FontFamily, $LabelFont)
    $dtPicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $dtPicker.Width = 270; $dtPicker.Left = $inputLeft; $dtPicker.Top = $y
    $dtPicker.Value = [datetime]::Now
    $panel.Controls.Add($dtPicker)
    $y += 70

    # Time
    $lblTime = New-Object System.Windows.Forms.Label
    $lblTime.Text = "Time (hh:mm AM/PM):"
    $lblTime.Font = New-Object System.Drawing.Font($FontFamily, $LabelFont)
    $lblTime.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lblTime.Left = 80; $lblTime.Top = $y; $lblTime.Width = $fieldLabelW
    $lblTime.Height = 40
    $panel.Controls.Add($lblTime)
    $tbTime = New-Object System.Windows.Forms.TextBox
    $tbTime.Font = New-Object System.Drawing.Font($FontFamily, $LabelFont)
    $tbTime.Left = $inputLeft; $tbTime.Top = $y; $tbTime.Width = 270; $tbTime.Height = 44
    $tbTime.Text = (Get-Date -Format "hh:mm tt")
    $panel.Controls.Add($tbTime)
    $y += 70

    # State dropdown
    $lblState = New-Object System.Windows.Forms.Label
    $lblState.Text = "State:"
    $lblState.Font = New-Object System.Drawing.Font($FontFamily, $LabelFont)
    $lblState.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lblState.Left = 80; $lblState.Top = $y; $lblState.Width = 120
    $lblState.Height = 40
    $panel.Controls.Add($lblState)
    $cbState = New-Object System.Windows.Forms.ComboBox
    $cbState.Font = New-Object System.Drawing.Font($FontFamily, $LabelFont)
    $cbState.Width = 160
    $cbState.Left = $inputLeft; $cbState.Top = $y
    $cbState.Items.Clear()
    foreach ($s in $stateList) { $cbState.Items.Add($s) | Out-Null }
    $cbState.SelectedItem = "FL"
    $panel.Controls.Add($cbState)
    $y += 70

    # Notes
    $lblNotes = New-Object System.Windows.Forms.Label
    $lblNotes.Text = "Notes:"
    $lblNotes.Font = New-Object System.Drawing.Font($FontFamily, $LabelFont)
    $lblNotes.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lblNotes.Left = 80; $lblNotes.Top = $y; $lblNotes.Width = $fieldLabelW
    $lblNotes.Height = 40
    $panel.Controls.Add($lblNotes)
    $tbNotes = New-Object System.Windows.Forms.TextBox
    $tbNotes.Font = New-Object System.Drawing.Font($FontFamily, $LabelFont)
    $tbNotes.Left = $inputLeft; $tbNotes.Top = $y; $tbNotes.Width = 700; $tbNotes.Height = 44
    $panel.Controls.Add($tbNotes)
    $y += 90

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = "Add Appointment"
    $btnAdd.Font = New-Object System.Drawing.Font($FontFamily, $ButtonFont)
    $btnAdd.Width = 340; $btnAdd.Height = 90
    $btnAdd.Left = [math]::Max(0, ($panel.Width - $btnAdd.Width - 250) / 2); $btnAdd.Top = $y
    $btnAdd.BackColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $btnAdd.ForeColor = "Black"
    $btnAdd.FlatStyle = "Flat"
    $btnAdd.Anchor = 'None'
    $btnAdd.Add_Click({
        $form.Hide()
        $driver = $cbDriver.Text; if (-not $driver) { Show-StylizedMessage "Driver required."; $form.Show(); return }
        $client = $cbClient.Text; if (-not $client) { Show-StylizedMessage "Client required."; $form.Show(); return }
        $dest = $cbDest.Text;     if (-not $dest) { Show-StylizedMessage "Destination required."; $form.Show(); return }
        $date = $dtPicker.Value.ToString("MM/dd/yyyy"); if (-not $date) { Show-StylizedMessage "Date required."; $form.Show(); return }
        $time = $tbTime.Text;     if (-not $time) { Show-StylizedMessage "Time required."; $form.Show(); return }
        $state = $cbState.Text;   if (-not $state) { Show-StylizedMessage "State required."; $form.Show(); return }
        $notes = $tbNotes.Text
        $appts = @(Get-Appts)
        $appt = [pscustomobject]@{
            Driver      = "$driver"
            Client      = "$client"
            Destination = "$dest"
            Date        = "$date"
            Time        = "$time"
            State       = "$state"
            Notes       = "$notes"
        }
        $appts += $appt
        Set-Appts $appts
        Show-StylizedMessage "Appointment Added!"
        $form.Close(); Show-MainMenu
    })
    $panel.Controls.Add($btnAdd)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = "Back"
    $btnBack.Font = New-Object System.Drawing.Font($FontFamily, $ButtonFont)
    $btnBack.Width = 220; $btnBack.Height = 90
    $btnBack.Left = [math]::Max(0, ($panel.Width - $btnBack.Width + 250) / 2); $btnBack.Top = $y
    $btnBack.BackColor = "Gray"
    $btnBack.ForeColor = "White"
    $btnBack.FlatStyle = "Flat"
    $btnBack.Anchor = 'None'
    $btnBack.Add_Click({ $form.Close(); Show-MainMenu })
    $panel.Controls.Add($btnBack)

    function Center-Panel {
        $panel.Left = [Math]::Max(0, ($form.ClientSize.Width - $panel.Width) / 2)
        $panel.Top  = [Math]::Max(0, ($form.ClientSize.Height - $panel.Height) / 2)
    }
    $form.Add_Resize({ Center-Panel })
    Center-Panel
    $form.ShowDialog()
}

function Show-DispatchHistory {
    $drivers = @(Get-Drivers)
    $dests = @(Get-Dest)
    $appts = @(Get-Appts)
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Dispatch History"
    $form.WindowState = 'Maximized'
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml($PrimaryColor)
    $form.FormBorderStyle = 'None'
    $form.TopMost = $true

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Width = 1900
    $panel.Height = 1200
    $panel.Anchor = 'None'
    $form.Controls.Add($panel)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Dispatch/Appointment History"
    $lbl.Font = New-Object System.Drawing.Font($FontFamily, 38)
    $lbl.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lbl.Width = 1700
    $lbl.Height = 70
    $lbl.Left = [math]::Max(0, ($panel.Width - $lbl.Width) / 2)
    $lbl.Top = 20
    $panel.Controls.Add($lbl)

    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Text = "Search:"
    $lblSearch.Font = New-Object System.Drawing.Font($FontFamily, 26)
    $lblSearch.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lblSearch.Left = 60; $lblSearch.Top = 110; $lblSearch.Width = 120; $lblSearch.Height = 40
    $panel.Controls.Add($lblSearch)
    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Font = New-Object System.Drawing.Font($FontFamily, 26)
    $txtSearch.Left = 180; $txtSearch.Top = 110; $txtSearch.Width = 370; $txtSearch.Height = 44
    $panel.Controls.Add($txtSearch)

    $lblDriver = New-Object System.Windows.Forms.Label
    $lblDriver.Text = "Filter by Driver:"
    $lblDriver.Font = New-Object System.Drawing.Font($FontFamily, 26)
    $lblDriver.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lblDriver.Left = 600; $lblDriver.Top = 110; $lblDriver.Width = 260; $lblDriver.Height = 40
    $panel.Controls.Add($lblDriver)
    $cbDriver = New-Object System.Windows.Forms.ComboBox
    $cbDriver.Font = New-Object System.Drawing.Font($FontFamily, 26)
    $cbDriver.Width = 280; $cbDriver.Left = 870; $cbDriver.Top = 110
    $cbDriver.Items.Clear()
    $cbDriver.Items.Add("All") | Out-Null
    foreach ($d in $drivers) { if ($d.Name -and $d.Name -is [string]) { $cbDriver.Items.Add([string]$d.Name) | Out-Null } }
    $cbDriver.SelectedIndex = 0
    $panel.Controls.Add($cbDriver)

    $lblDate = New-Object System.Windows.Forms.Label
    $lblDate.Text = "Filter by Date:"
    $lblDate.Font = New-Object System.Drawing.Font($FontFamily, 26)
    $lblDate.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($AccentColor)
    $lblDate.Left = 1200; $lblDate.Top = 110; $lblDate.Width = 220; $lblDate.Height = 40
    $panel.Controls.Add($lblDate)
    $cbDate = New-Object System.Windows.Forms.ComboBox
    $cbDate.Font = New-Object System.Drawing.Font($FontFamily, 26)
    $cbDate.Width = 280; $cbDate.Left = 1430; $cbDate.Top = 110
    $cbDate.Items.Clear()
    $cbDate.Items.Add("All") | Out-Null
    $allDates = $appts | ForEach-Object { $_.Date }
    foreach ($d in ($allDates | Sort-Object -Unique)) { if ($d -and $d -is [string]) { $cbDate.Items.Add([string]$d) | Out-Null } }
    $cbDate.SelectedIndex = 0
    $panel.Controls.Add($cbDate)

    $lv = New-Object System.Windows.Forms.ListView
    $lv.View = 'Details'
    $lv.FullRowSelect = $true
    $lv.Width = 1830; $lv.Height = 700
    $lv.Left = [math]::Max(0, ($panel.Width - $lv.Width) / 2)
    $lv.Top = 250
    $lv.Font = New-Object System.Drawing.Font($FontFamily, 24)
    $lv.GridLines = $true
    $lv.MultiSelect = $false
    $lv.Columns.Add("Driver", 180) | Out-Null
    $lv.Columns.Add("Client", 180) | Out-Null
    $lv.Columns.Add("Destination Address", 500) | Out-Null
    $lv.Columns.Add("Date", 140) | Out-Null
    $lv.Columns.Add("Time", 120) | Out-Null
    $lv.Columns.Add("State", 100) | Out-Null
    $lv.Columns.Add("Notes", 520) | Out-Null
    $panel.Controls.Add($lv)

    function GetDestinationAddress($destName) {
        $found = $dests | Where-Object { $_.Name -eq $destName }
        if ($found -and $found.Address) { return $found.Address }
        return $destName
    }

    function GetFiltered {
        $fDriver = $cbDriver.Text
        $fDate = $cbDate.Text
        $search = $txtSearch.Text.ToLower()
        $filtered = $appts | Where-Object {
            ($fDriver -eq "All" -or $_.Driver -eq $fDriver) -and
            ($fDate -eq "All" -or $_.Date -eq $fDate) -and
            (
                $search -eq "" -or
                ($_.Driver -and $_.Driver.ToLower().Contains($search)) -or
                ($_.Client -and $_.Client.ToLower().Contains($search)) -or
                ($_.Destination -and $_.Destination.ToLower().Contains($search)) -or
                (GetDestinationAddress($_.Destination).ToLower().Contains($search)) -or
                ($_.Date -and $_.Date.ToLower().Contains($search)) -or
                ($_.Time -and $_.Time.ToLower().Contains($search)) -or
                ($_.State -and $_.State.ToLower().Contains($search)) -or
                ($_.Notes -and $_.Notes.ToLower().Contains($search))
            )
        }
        return ,$filtered
    }

    function Load-Schedule {
        $lv.Items.Clear()
        $filtered = GetFiltered
        foreach ($appt in $filtered) {
            $addr = GetDestinationAddress $appt.Destination
            $item = New-Object System.Windows.Forms.ListViewItem([string]$appt.Driver)
            $item.SubItems.Add([string]$appt.Client)         | Out-Null
            $item.SubItems.Add([string]$addr)                | Out-Null
            $item.SubItems.Add([string]$appt.Date)           | Out-Null
            $item.SubItems.Add([string]$appt.Time)           | Out-Null

# Need to switch this to an IF/ELSE, ternary operators not supported in PS            
#$item.SubItems.Add(($appt.State -ne $null -and $appt.State -ne "") ? $appt.State : "FL") | Out-Null

            $item.SubItems.Add([string]$appt.Notes)          | Out-Null
            $lv.Items.Add($item) | Out-Null
        }
        if ($lv.Items.Count -eq 0) {
            $item = New-Object System.Windows.Forms.ListViewItem("No results found.")
            for ($i=0; $i -lt 6; $i++) { $item.SubItems.Add("") | Out-Null }
            $lv.Items.Add($item) | Out-Null
        }
        foreach ($i in 0..($lv.Columns.Count-1)) { $lv.Columns[$i].Width = -2 }
    }

    $cbDriver.Add_SelectedIndexChanged({ Load-Schedule })
    $cbDate.Add_SelectedIndexChanged({ Load-Schedule })
    $txtSearch.Add_TextChanged({ Load-Schedule })

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = "Back"
    $btnBack.Font = New-Object System.Drawing.Font($FontFamily, $ButtonFont)
    $btnBack.Width = 400
    $btnBack.Height = 90
    $btnBack.Left = [math]::Max(0, ($panel.Width - 400) / 2)
    $btnBack.Top = 1000
    $btnBack.BackColor = "Gray"
    $btnBack.ForeColor = "White"
    $btnBack.FlatStyle = "Flat"
    $btnBack.Anchor = 'None'
    $btnBack.Add_Click({ $form.Close(); Show-MainMenu })
    $panel.Controls.Add($btnBack)

    function Center-Panel {
        $panel.Left = [Math]::Max(0, ($form.ClientSize.Width - $panel.Width) / 2)
        $panel.Top  = [Math]::Max(0, ($form.ClientSize.Height - $panel.Height) / 2)
    }
    $form.Add_Resize({ Center-Panel })
    Center-Panel
    Load-Schedule
    $form.ShowDialog()
}

try {
    Show-MainMenu
} catch {
    [System.Windows.Forms.MessageBox]::Show("Fatal Error: $($_.Exception.Message)")
}
