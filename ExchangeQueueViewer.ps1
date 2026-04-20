# Exchange Queue Manager PRO v6.2
# Compatibility: PowerShell 5.1 & PowerShell 7+
# Author: Rexxer777

function Load-ExchangeTools {
    if (Get-Command Get-Queue -ErrorAction SilentlyContinue) { return $true }
    if ($PSVersionTable.PSVersion.Major -lt 6) {
        $snapin = 'Microsoft.Exchange.Management.PowerShell.SnapIn'
        if (Get-PSSnapin -Registered $snapin -ErrorAction SilentlyContinue) {
            Add-PSSnapin $snapin -ErrorAction SilentlyContinue
        }
    }
    if (-not (Get-Command Get-Queue -ErrorAction SilentlyContinue)) {
        Import-Module Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue
    }
    $check = Get-Command Get-Queue -ErrorAction SilentlyContinue
    if ($check) { return $true } else { return $false }
}

$ToolsLoaded = Load-ExchangeTools
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

$colorBg = [Drawing.Color]::FromArgb(245, 246, 250)
$colorAccent = [Drawing.Color]::FromArgb(0, 120, 212)
$colorDanger = [Drawing.Color]::FromArgb(196, 43, 28)
$fontBold = New-Object Drawing.Font('Segoe UI', 9, [Drawing.FontStyle]::Bold)

$form = New-Object Windows.Forms.Form
$form.Text = 'Exchange Queue Manager PRO v6.2'
$form.Size = New-Object Drawing.Size(1200, 720)
$form.StartPosition = 'CenterScreen'
$form.BackColor = $colorBg

$toolbar = New-Object Windows.Forms.Panel
$toolbar.Dock = 'Top'; $toolbar.Height = 48; $toolbar.BackColor = $colorAccent
$form.Controls.Add($toolbar)

$lblTitle = New-Object Windows.Forms.Label
$lblTitle.Text = 'Exchange Queue Manager PRO v6.2'
$lblTitle.ForeColor = [Drawing.Color]::White
$lblTitle.Font = New-Object Drawing.Font('Segoe UI', 12, [Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true; $lblTitle.Location = New-Object Drawing.Point(14, 12)
$toolbar.Controls.Add($lblTitle)

$filterPanel = New-Object Windows.Forms.Panel
$filterPanel.Dock = 'Top'; $filterPanel.Height = 50; $filterPanel.BackColor = [Drawing.Color]::White
$form.Controls.Add($filterPanel)

$filterText = New-Object Windows.Forms.TextBox
$filterText.Location = New-Object Drawing.Point(10, 22); $filterText.Size = New-Object Drawing.Size(140, 24)
$filterText.BorderStyle = 'FixedSingle'
$filterPanel.Controls.Add($filterText)

$statusFilter = New-Object Windows.Forms.ComboBox
$statusFilter.Location = New-Object Drawing.Point(162, 22); $statusFilter.Items.AddRange(@('(All)','Active','Retry','Suspended'))
$statusFilter.SelectedIndex = 0; $statusFilter.DropDownStyle = 'DropDownList'
$filterPanel.Controls.Add($statusFilter)

$grid = New-Object Windows.Forms.DataGridView
$grid.Location = New-Object Drawing.Point(12, 110); $grid.Size = New-Object Drawing.Size(630, 350)
$grid.SelectionMode = 'FullRowSelect'; $grid.ReadOnly = $true; $grid.RowHeadersVisible = $false
$grid.AutoSizeColumnsMode = 'Fill'; $grid.MultiSelect = $true
$form.Controls.Add($grid)

$msgGrid = New-Object Windows.Forms.DataGridView
$msgGrid.Location = New-Object Drawing.Point(652, 110); $msgGrid.Size = New-Object Drawing.Size(518, 350)
$msgGrid.SelectionMode = 'FullRowSelect'; $msgGrid.ReadOnly = $true; $msgGrid.RowHeadersVisible = $false
$msgGrid.AutoSizeColumnsMode = 'Fill'; $msgGrid.MultiSelect = $true
$form.Controls.Add($msgGrid)

$btnPanel = New-Object Windows.Forms.Panel
$btnPanel.Location = New-Object Drawing.Point(12, 468); $btnPanel.Size = New-Object Drawing.Size(1168, 44)
$form.Controls.Add($btnPanel)

function New-Btn($txt, $x, $clr, $p, $w) {
    $b = New-Object Windows.Forms.Button; $b.Text = $txt; $b.Location = New-Object Drawing.Point($x, 4)
    $b.Size = New-Object Drawing.Size($w, 34); $b.FlatStyle = 'Flat'; $b.BackColor = $clr
    $b.ForeColor = [Drawing.Color]::White; $b.Font = $fontBold; $p.Controls.Add($b); return $b
}

$btnDelM = New-Btn 'Delete Selected Msg' 0 $colorDanger $btnPanel 160
$btnDelQ = New-Btn 'Clear Selected Queue' 170 $colorDanger $btnPanel 200
$btnDelA = New-Btn 'Clear ALL Filtered' 380 ([Drawing.Color]::DarkRed) $btnPanel 180

$statusBox = New-Object Windows.Forms.TextBox
$statusBox.Location = New-Object Drawing.Point(12, 540); $statusBox.Size = New-Object Drawing.Size(1168, 100)
$statusBox.Multiline = $true; $statusBox.ReadOnly = $true; $statusBox.ScrollBars = 'Vertical'
$statusBox.Font = New-Object Drawing.Font('Consolas', 9); $statusBox.Text = 'Status: Ready'
$form.Controls.Add($statusBox)

$ScriptBlock = {
    param($ParamHash)
    if (-not (Get-Command Get-Queue -ErrorAction SilentlyContinue)) {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue
    }
    try {
        $mode = $ParamHash['mode']
        if ($mode -eq 'refresh') {
            $qs = Get-Queue -ErrorAction Stop
            if ($ParamHash['fText']) { $wild = '*' + $ParamHash['fText'] + '*'; $qs = $qs | Where-Object { $_.Identity -like $wild } }
            if ($ParamHash['fStatus'] -and $ParamHash['fStatus'] -ne '(All)') { $qs = $qs | Where-Object { $_.Status -eq $ParamHash['fStatus'] } }
            $dt = New-Object System.Data.DataTable
            $null = $dt.Columns.Add('Identity'); $null = $dt.Columns.Add('Messages', [int]); $null = $dt.Columns.Add('Status'); $null = $dt.Columns.Add('NextHopDomain')
            foreach ($q in $qs) { $null = $dt.Rows.Add($q.Identity.ToString(), [int]$q.MessageCount, $q.Status.ToString(), $q.NextHopDomain.ToString()) }
            return @{ type = 'refresh'; data = $dt }
        }
        elseif ($mode -eq 'loadmsg') {
            $ms = Get-Message -Queue $ParamHash['queue'] -ResultSize 150 -ErrorAction Stop
            $dt = New-Object System.Data.DataTable
            $null = $dt.Columns.Add('Identity'); $null = $dt.Columns.Add('From'); $null = $dt.Columns.Add('Subject'); $null = $dt.Columns.Add('Status')
            foreach ($m in $ms) { $null = $dt.Rows.Add($m.Identity.ToString(), $m.FromAddress.ToString(), $m.Subject, $m.Status.ToString()) }
            return @{ type = 'msgs'; data = $dt; queue = $ParamHash['queue'] }
        }
        elseif ($mode -eq 'deleteMsgs') {
            $c = 0; foreach ($id in $ParamHash['ids']) { try { Remove-Message -Identity $id -Confirm:$false -ErrorAction Stop; $c++ } catch {} }
            return @{ type = 'done' }
        }
        elseif ($mode -eq 'deleteQueues') {
            $c = 0; foreach ($qi in $ParamHash['queues']) { try { Remove-Message -Queue $qi -Confirm:$false -ErrorAction Stop; $c++ } catch {} }
            return @{ type = 'done' }
        }
    } catch { return @{ type = 'error'; msg = ('WORKER ERROR: ' + $_.Exception.Message) } }
}

$SyncTimer = New-Object Windows.Forms.Timer; $SyncTimer.Interval = 400
function Start-Task($params) {
    if ($Global:PSInstance -and -not $Global:JobHandle.IsCompleted) { return }
    $PS = [PowerShell]::Create().AddScript($ScriptBlock).AddArgument($params)
    $Global:JobHandle = $PS.BeginInvoke(); $Global:PSInstance = $PS; $SyncTimer.Start()
}

function Run-Refresh { $p = @{}; $p['mode'] = 'refresh'; $p['fText'] = $filterText.Text; $p['fStatus'] = $statusFilter.Text; Start-Task $p }

$grid.Add_SelectionChanged({ if ($grid.SelectedRows.Count -eq 1) { $id = $grid.SelectedRows[0].Cells['Identity'].Value; if ($id) { $p = @{}; $p['mode'] = 'loadmsg'; $p['queue'] = $id; Start-Task $p } } })
$filterText.Add_TextChanged({ Run-Refresh })
$statusFilter.Add_SelectedIndexChanged({ Run-Refresh })

$btnDelM.Add_Click({
    if ($msgGrid.SelectedRows.Count -eq 0) { return }
    $ids = @(); foreach ($row in $msgGrid.SelectedRows) { $ids += $row.Cells['Identity'].Value }
    if ([Windows.Forms.MessageBox]::Show(('Delete ' + $ids.Count + ' items?'), 'Confirm', 4) -eq 'Yes') { $p = @{}; $p['mode'] = 'deleteMsgs'; $p['ids'] = $ids; Start-Task $p }
})

$btnDelQ.Add_Click({
    if ($grid.SelectedRows.Count -eq 0) { return }
    $qs = @(); foreach ($row in $grid.SelectedRows) { $qs += $row.Cells['Identity'].Value }
    if ([Windows.Forms.MessageBox]::Show(('Clear ' + $qs.Count + ' queues?'), 'Confirm', 4) -eq 'Yes') { $p = @{}; $p['mode'] = 'deleteQueues'; $p['queues'] = $qs; Start-Task $p }
})

$btnDelA.Add_Click({
    if ($grid.Rows.Count -eq 0) { return }
    $qs = @(); foreach ($row in $grid.Rows) { $qs += $row.Cells['Identity'].Value }
    if ([Windows.Forms.MessageBox]::Show(('Clear ALL ' + $qs.Count + ' queues?'), 'Confirm', 4, 16) -eq 'Yes') { $p = @{}; $p['mode'] = 'deleteQueues'; $p['queues'] = $qs; Start-Task $p }
})

$SyncTimer.Add_Tick({
    if ($Global:JobHandle -and $Global:JobHandle.IsCompleted) {
        $SyncTimer.Stop(); $localPS = $Global:PSInstance; $localHandle = $Global:JobHandle; $Global:PSInstance = $null; $Global:JobHandle = $null; $triggerRefresh = $false
        try {
            if ($localPS.InvocationStateInfo.State -eq 'Failed') { $statusBox.Text = ('ENGINE ERROR: ' + $localPS.InvocationStateInfo.Reason.Message) }
            else {
                $resArr = $localPS.EndInvoke($localHandle); if ($resArr.Count -gt 0) {
                    $res = $resArr[0]
                    if ($res['type'] -eq 'refresh') { $grid.DataSource = $res['data']; $statusBox.Text = 'Queues updated.' }
                    elseif ($res['type'] -eq 'msgs') { $msgGrid.DataSource = $res['data']; $statusBox.Text = 'Messages loaded.' }
                    elseif ($res['type'] -eq 'done') { $statusBox.Text = 'Success: Action processed.'; $triggerRefresh = $true }
                    elseif ($res['type'] -eq 'error') { $statusBox.Text = $res['msg'] }
                }
            }
        } catch { $statusBox.Text = ('SYSTEM ERROR: ' + $_.Exception.Message) }
        finally { if ($localPS) { $localPS.Dispose() }; if ($triggerRefresh) { Run-Refresh } }
    }
})

$form.Add_Shown({ if (-not $ToolsLoaded) { $statusBox.Text = 'FATAL ERROR: Exchange tools not found.'; $statusBox.ForeColor = [Drawing.Color]::Red } else { $statusBox.Text = 'Status: Loading...'; Run-Refresh } })
[Windows.Forms.Application]::Run($form)
