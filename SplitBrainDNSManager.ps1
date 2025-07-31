<#
.SYNOPSIS
    This script helps to manage Split-Brain DNS on Windows Server

.DESCRIPTION
    This script creates a GUI to manage Split-Brain DNS.  It allows you to view/create/delete Zone Scopes, Client Subnets, and Query Resolution Policies, as well as view/create/delete Records within the Zone Scopes.

.EXAMPLE
    Example usage of the script:
    PS> .\SplitBrainDNSManager.ps1

.NOTES
    Author: Jessie Slight
    Date: 2025/07/30
    Version: 1.0
    Initial release.
    Version: 1.1
    Added Import and Export buttons

.LINK
    GitHub: https://github.com/jslight90

#>

# Check if the script is running as Administrator
If (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    # Relaunch the script with elevated privileges
    $arguments = "& '" + $MyInvocation.MyCommand.Definition + "'"
    Start-Process powershell -ArgumentList $arguments -Verb RunAs
    Exit
}

Add-Type -AssemblyName System.Windows.Forms, System.Drawing, System.Data

# --- Helper Functions ------------------------------------------------------

function Show-Error($msg) {
    [System.Windows.Forms.MessageBox]::Show($msg, 'Error', 'OK', 'Error') | Out-Null
}
function Show-Info($msg) {
    [System.Windows.Forms.MessageBox]::Show($msg, 'Info', 'OK', 'Information') | Out-Null
}

# --- Variables ---------------------------------------------------------

$script:recordsZone = $null
$script:recordsScope = $null

# --- Functions ---------------------------------------------------------

function Get-Zones {
    try {
        $zones = Get-DnsServerZone | Where-Object { $_.ZoneName -notmatch 'in-addr\.arpa$|ip6\.arpa$|^TrustAnchors$' } | Select-Object -ExpandProperty ZoneName

        return $zones
    }
    catch {
        Show-Error "Failed loading zones: $_"
    }
    
}

function Get-ZoneScopes {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $ZoneName
    )

    try {
        $zoneScopes = Get-DnsServerZoneScope -ZoneName $ZoneName | Where-Object { $_.ZoneScope -ne $_.ZoneName }

        return $zoneScopes
    }
    catch {
        Show-Error "Failed loading zone scopes for Zone '$ZoneName': $_"
    }
}

function Get-ClientSubnets {
    try {
        $subnets = Get-DnsServerClientSubnet

        return $subnets
    }
    catch {
        Show-Error "Failed loading client subnets: $_"
    }
}

function Update-ZoneScopes {
    try {
        $all = @()

        foreach ($zone in Get-Zones) {
            try {
                foreach ($scope in Get-ZoneScopes -ZoneName $zone) {
                    $all += [PSCustomObject]@{
                        Name = $scope.ZoneScope
                        Zone = $zone
                    }
                }
            } catch {
                # if a zone genuinely doesn't support scopes, skip quietly
            }
        }

        $dtScopes.Rows.Clear() | Out-Null
        foreach ($row in $all) {
            $dtScopes.Rows.Add($row.Name, $row.Zone) | Out-Null
        }
        $dgScopes.Refresh()
    } catch {
        Show-Error "Failed loading scopes: $_"
    }
}

function Update-ClientSubnets {
    try {
        $dtSubnets.Rows.Clear() | Out-Null
        foreach ($row in Get-ClientSubnets) {
            $dtSubnets.Rows.Add($row.Name, $row.IPV4Subnet -join ', ', $row.IPv6Subnet -join ', ') | Out-Null
        }
        $dgSubnets.Refresh()
    }
    catch {
        Show-Error "Failed loading client subnets: $_"
    }
    
}

function Update-QueryResolutionPolicies {
    try {
        $all = @()

        foreach ($zone in Get-Zones) {
            try {
                Get-DnsServerQueryResolutionPolicy -ZoneName $zone |
                ForEach-Object {
                    $all += [PSCustomObject]@{
                        Name = $_.Name
                        Zone = $zone
                        Scope = $_.Content.ScopeName
                        Subnet = $_.Criteria.Criteria.Substring(3)
                        Action = $_.Action
                        Enabled = $_.IsEnabled
                    }
                }
            }
            catch {
                # if a zone genuinely doesn't support scopes, skip quietly
            }
        }

        $dtPolicies.Rows.Clear() | Out-Null
        foreach ($row in $all) {
            $dtPolicies.Rows.Add($row.Name, $row.Zone, $row.Scope, $row.Subnet, $row.Action, $row.Enabled) | Out-Null
        }
        $dgPolicies.Refresh()
    }
    catch {
        Show-Error "Failed loading query resolution polcies: $_"
    }
}

function Update-Records {
    try {
        $all = @()

        foreach ($record in Get-DnsServerResourceRecord -ZoneName $script:recordsZone -ZoneScope $script:recordsScope) {
            $type = $record.RecordType
            $data = $record.RecordData
            $all += [PSCustomObject]@{
                Name = $record.HostName
                Type = $type
                Data = switch ($type) {
                    'A'     { $data.IPv4Address }
                    'AAAA'  { $data.IPv6Address }
                    'CNAME' { $data.HostNameAlias }
                    'MX'    { "[$($data.Preference)] $($data.MailExchange)" }
                    'TXT'   { $data.DescriptiveText }
                    'SRV'   { "[$($data.Priority)][$($data.Weight)][$($data.Port)] $($data.DomainName)" }
                    'PTR'   { $data.PtrDomainName }
                    'NS'    { $data.NameServer }
                    'SOA'   { "$($data.ResponsiblePerson -replace '^(.*?)\.','$1@') [$($data.PrimaryServer)]" }
                    default { $data }
                }
            }
        }

        $dtRecords.Rows.Clear() | Out-Null
        foreach ($row in $all) {
            $dtRecords.Rows.Add($row.Name, $row.Type, $row.Data) | Out-Null
        }
        $dgRecords.Refresh()
    }
    catch {
        Show-Error "Failed loading resource records: $_"
    }
}

function Add-Record {
    [CmdletBinding()]
    param (
        [string]$Name,
        [String]$ZoneName,
        [string]$ZoneScope,
        [string]$RecordType,
        [string]$RecordData
    )
    
    switch ($RecordType) {
        'A'     {
            Add-DnsServerResourceRecordA -ZoneName $ZoneName -ZoneScope $ZoneScope -Name $Name -IPv4Address $RecordData
        }
        'CNAME' {
            Add-DnsServerResourceRecordCName -ZoneName $ZoneName -ZoneScope $ZoneScope -Name $Name -HostNameAlias $RecordData
        }
        'TXT'   {
            Add-DnsServerResourceRecord -Txt -ZoneName $ZoneName -ZoneScope $ZoneScope -Name $Name -DescriptiveText $RecordData
        }
        'PTR'   {
            Add-DnsServerResourceRecordPtr -ZoneName $ZoneName -ZoneScope $ZoneScope -Name $Name -PtrDomainName $RecordData
        }
        Default {
            # unsupported record type, silently continue
            Write-Host "Unable to add $RecordType record."
        }
    }
}

function Import-FromCsv {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $openFileDialog.Title = "Open CSV Import"

    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $openPath = $openFileDialog.FileName
        try {
            $data = Import-Csv -Path $openPath
            return $data
        }
        catch {
            Show-Error "Error during import: $_"
        }
    } else {
        Write-host "Import cancelled by user."
    }
}

function Export-GridViewToCsv {
    [CmdletBinding()]
    param (
        [System.Windows.Forms.DataGridView]$DataGridView
    )
    
    $date = Get-Date -Format "yyyy-MM-dd"
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $saveFileDialog.Title = "Select CSV Export Location"
    $saveFileDialog.FileName = "SplitBrain_Export_$date.csv" # default filename

    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $savePath = $saveFileDialog.FileName

        try {
            $DataGridView.Rows | Select-Object -ExpandProperty DataBoundItem | Export-Csv -Path $savePath -NoTypeInformation
            Show-Info "Export saved to $savePath."
        }
        catch {
            Show-Error "Error during export: $_"
        }
    } else {
        Write-Host "Export cancelled by user."
    }
}

# --- Main Window ---------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Split-Brain DNS Manager'
$form.Size = New-Object System.Drawing.Size(900,600)
$form.StartPosition = 'CenterScreen'

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = 'Fill'
# Create each tab page
$tabs = @{
    'Zone Scopes and Records'   = New-Object System.Windows.Forms.TabPage('Zone Scopes and Records')
    'DNS Client Subnets'        = New-Object System.Windows.Forms.TabPage('DNS Client Subnets')
    'Query Resolution Policies' = New-Object System.Windows.Forms.TabPage('Query Resolution Policies')
}
$tabControl.TabPages.Add($tabs['Zone Scopes and Records'])
$tabControl.TabPages.Add($tabs['DNS Client Subnets'])
$tabControl.TabPages.Add($tabs['Query Resolution Policies'])
$form.Controls.Add($tabControl)

# --- Records Window ---------------------------------------------------------

$records = New-Object System.Windows.Forms.Form
$records.Text = 'Resource Records'
$records.Size = New-Object System.Drawing.Size(1200,800)
$records.StartPosition = 'CenterScreen'

$dtRecords = New-Object System.Data.DataTable
$dtRecords.Columns.Add('Name') | Out-Null
$dtRecords.Columns.Add('Type') | Out-Null
$dtRecords.Columns.Add('Data') | Out-Null

$dgRecords = New-Object System.Windows.Forms.DataGridView
$dgRecords.Dock = 'Top'
$dgRecords.Height = 690
$dgRecords.SelectionMode = 'FullRowSelect'
$dgRecords.AutoSizeColumnsMode = 'AllCells'
$dgRecords.EditMode = 'EditProgrammatically'
$dgRecords.MultiSelect = $false;
$dgRecords.AllowUserToAddRows = $false
$dgRecords.AllowUserToDeleteRows = $false
$dgRecords.DataSource = $dtRecords
$records.Controls.Add($dgRecords)

$btnRefeshRecords          = New-Object System.Windows.Forms.Button
$btnRefeshRecords.Text     = 'Refresh'
$btnRefeshRecords.Location = '10,700'
$btnRefeshRecords.Size     = '100,30'

$btnAddRecord              = New-Object System.Windows.Forms.Button
$btnAddRecord.Text         = 'Add'
$btnAddRecord.Location     = '120,700'
$btnAddRecord.Size         = '100,30'

$btnRemoveRecord           = New-Object System.Windows.Forms.Button
$btnRemoveRecord.Text      = 'Remove'
$btnRemoveRecord.Location  = '230,700'
$btnRemoveRecord.Size      = '100,30'

$btnImportRecords          = New-Object System.Windows.Forms.Button
$btnImportRecords.Text     = 'Import'
$btnImportRecords.Location = '970,700'
$btnImportRecords.Size     = '100,30'

$btnExportRecords          = New-Object System.Windows.Forms.Button
$btnExportRecords.Text     = 'Export'
$btnExportRecords.Location = '1080,700'
$btnExportRecords.Size     = '100,30'

$records.Controls.AddRange(@($btnRefeshRecords,$btnAddRecord,$btnRemoveRecord,$btnImportRecords,$btnExportRecords))

$btnRefeshRecords.Add_Click({ Update-Records })

$btnAddRecord.Add_Click({
    $prompt               = New-Object System.Windows.Forms.Form
    $prompt.Text          = 'New Record'
    $prompt.Size          = New-Object System.Drawing.Size(300, 300)
    $prompt.StartPosition = 'CenterScreen'

    $labelName          = New-Object System.Windows.Forms.Label
    $labelName.Location = New-Object System.Drawing.Point(10, 10)
    $labelName.Size     = New-Object System.Drawing.Size(260, 20)
    $labelName.Text     = 'Record Name (w/o Zone Name)'
    $prompt.Controls.Add($labelName)
    $fieldName          = New-Object System.Windows.Forms.TextBox
    $fieldName.Location = New-Object System.Drawing.Point(10, 30)
    $fieldName.Size     = New-Object System.Drawing.Size(260, 20)
    $prompt.Controls.Add($fieldName)

    $labelType          = New-Object System.Windows.Forms.Label
    $labelType.Location = New-Object System.Drawing.Point(10, 70)
    $labelType.Size     = New-Object System.Drawing.Size(260, 20)
    $labelType.Text     = 'Type'
    $prompt.Controls.Add($labelType)
    $fieldType          = New-Object System.Windows.Forms.ComboBox
    $fieldType.Location = New-Object System.Drawing.Point(10, 90)
    $fieldType.Size     = New-Object System.Drawing.Size(260, 20)
    $fieldType.Items.Clear()
    $fieldType.Items.AddRange(@('A','CNAME','TXT','PTR'))
    $prompt.Controls.Add($fieldType)

    $labelData          = New-Object System.Windows.Forms.Label
    $labelData.Location = New-Object System.Drawing.Point(10, 130)
    $labelData.Size     = New-Object System.Drawing.Size(260, 20)
    $labelData.Text     = 'Data'
    $prompt.Controls.Add($labelData)
    $fieldData          = New-Object System.Windows.Forms.TextBox
    $fieldData.Location = New-Object System.Drawing.Point(10, 150)
    $fieldData.Size     = New-Object System.Drawing.Size(260, 20)
    $prompt.Controls.Add($fieldData)

    # Create the Yes button and its properties
    $yesButton = New-Object System.Windows.Forms.Button
    $yesButton.Location = New-Object System.Drawing.Point(60, 220)
    $yesButton.Size = New-Object System.Drawing.Size(75, 23)
    $yesButton.Text = 'Create'
    $yesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $prompt.AcceptButton = $yesButton
    $prompt.Controls.Add($yesButton)
    # Create the No button and its properties
    $noButton = New-Object System.Windows.Forms.Button
    $noButton.Location = New-Object System.Drawing.Point(165, 220)
    $noButton.Size = New-Object System.Drawing.Size(75, 23)
    $noButton.Text = 'Cancel'
    $noButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $prompt.CancelButton = $noButton
    $prompt.Controls.Add($noButton)

    $prompt.TopMost = $true

    if ($prompt.ShowDialog() -eq [System.Windows.Forms.DialogResult]::Yes) {
        $zoneName = $script:recordsZone
        $scopeName = $script:recordsScope
        $recordName = $fieldName.Text
        $recordType = $fieldType.SelectedItem
        $recordData = $fieldData.Text

        if ($recordName -and $recordType -and $recordData) {
            try {
                Add-Record -Name $recordName -ZoneName $zoneName -ZoneScope $scopeName -RecordType $recordType -RecordData $recordData
                Show-Info "$recordType Record '$recordName' added for Zone '$zoneName' and Scope '$scopeName'."
                Update-Records
            }
            catch {
                Show-Error $_
            }
        } else {
            Show-Error "Please provide a value for all fields."
        }
    }
})

$btnRemoveRecord.Add_Click({
    if ($dgRecords.SelectedRows.Count -eq 0) { return }
    $row = $dgRecords.SelectedRows[0]

    $recordName = $row.Cells['Name'].Value
    $recordType = $row.Cells['Type'].Value
    $recordData = $row.Cells['Data'].Value
    $zoneName = $script:recordsZone
    $scopeName = $script:recordsScope

    $prompt               = New-Object System.Windows.Forms.Form
    $prompt.Text          = 'Remove Record'
    $prompt.Size          = New-Object System.Drawing.Size(300, 200)
    $prompt.StartPosition = 'CenterScreen'

    $promptDesc          = New-Object System.Windows.Forms.Label
    $promptDesc.Location = New-Object System.Drawing.Point(10, 10)
    $promptDesc.Size     = New-Object System.Drawing.Size(260, 80)
    $promptDesc.Text     = "Are you sure you want to remove the $recordType Record '$recordName' for Zone '$zoneName' and Scope '$scopeName'?"
    $prompt.Controls.Add($promptDesc)

    # Create the Yes button and its properties
    $yesButton = New-Object System.Windows.Forms.Button
    $yesButton.Location = New-Object System.Drawing.Point(60, 120)
    $yesButton.Size = New-Object System.Drawing.Size(75, 23)
    $yesButton.Text = 'Remove'
    $yesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $prompt.AcceptButton = $yesButton
    $prompt.Controls.Add($yesButton)
    # Create the No button and its properties
    $noButton = New-Object System.Windows.Forms.Button
    $noButton.Location = New-Object System.Drawing.Point(165, 120)
    $noButton.Size = New-Object System.Drawing.Size(75, 23)
    $noButton.Text = 'Cancel'
    $noButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $prompt.CancelButton = $noButton
    $prompt.Controls.Add($noButton)

    $prompt.TopMost = $true

    if ($prompt.ShowDialog() -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            if ($recordType -in @('A','AAAA','CNAME','TXT','PTR','NS')) {
                Remove-DnsServerResourceRecord -Name $recordName -RRType $recordType -ZoneName $zoneName -ZoneScope $scopeName -RecordData $recordData -Force
            } else {
                Remove-DnsServerResourceRecord -Name $recordName -RRType $recordType -ZoneName $zoneName -ZoneScope $scopeName -Force
            }
            Show-Info "$recordType Record '$recordName' removed for Zone '$zoneName' and Scope '$scopeName'."
            Update-Records
        }
        catch {
            Show-Error $_
        }
    }
})

$btnImportRecords.Add_Click({
    $import = Import-FromCsv
    if ($import -and @($import).Count -gt 0) {
        foreach ($row in $import) {
            try {
                Add-Record -Name $row.Name -ZoneName $script:recordsZone -ZoneScope $script:recordsScope -RecordType $row.Type -RecordData $row.Data
            }
            catch {
                Write-Host "Unable to add $($row.Type) record '$($row.Name)' with data '$($row.Data)':`n$_"
            }
        }
        Show-Info "Imported records."
        Update-Records
    }  
})

$btnExportRecords.Add_Click({ Export-GridViewToCsv -DataGridView $dgRecords })

# --- Tab 1: Zone Scopes and Records ---------------------------------------------------

$dtScopes = New-Object System.Data.DataTable
$dtScopes.Columns.Add('Scope Name') | Out-Null
$dtScopes.Columns.Add('Zone') | Out-Null

$dgScopes = New-Object System.Windows.Forms.DataGridView
$dgScopes.Dock = 'Top'
$dgScopes.Height = 490
$dgScopes.SelectionMode = 'FullRowSelect'
$dgScopes.AutoSizeColumnsMode = 'AllCells'
$dgScopes.EditMode = 'EditProgrammatically'
$dgScopes.MultiSelect = $false;
$dgScopes.AllowUserToAddRows = $false
$dgScopes.AllowUserToDeleteRows = $false
$dgScopes.DataSource = $dtScopes
$tabs['Zone Scopes and Records'].Controls.Add($dgScopes)

$btnRefreshScopes          = New-Object System.Windows.Forms.Button
$btnRefreshScopes.Text     = 'Refresh'
$btnRefreshScopes.Location = '10,500'
$btnRefreshScopes.Size     = '100,30'

$btnAddScope               = New-Object System.Windows.Forms.Button
$btnAddScope.Text          = 'Add'
$btnAddScope.Location      = '120,500'
$btnAddScope.Size          = '100,30'

$btnEditRecords            = New-Object System.Windows.Forms.Button
$btnEditRecords.Text       = 'Edit Records'
$btnEditRecords.Location   = '230,500'
$btnEditRecords.Size       = '100,30'

$btnRemoveScope            = New-Object System.Windows.Forms.Button
$btnRemoveScope.Text       = 'Remove'
$btnRemoveScope.Location   = '340,500'
$btnRemoveScope.Size       = '100,30'

$btnImportScopes           = New-Object System.Windows.Forms.Button
$btnImportScopes.Text      = 'Import'
$btnImportScopes.Location  = '670,500'
$btnImportScopes.Size      = '100,30'

$btnExportScopes           = New-Object System.Windows.Forms.Button
$btnExportScopes.Text      = 'Export'
$btnExportScopes.Location  = '780,500'
$btnExportScopes.Size      = '100,30'

$tabs['Zone Scopes and Records'].Controls.AddRange(@($btnRefreshScopes,$btnAddScope,$btnEditRecords,$btnRemoveScope,$btnImportScopes,$btnExportScopes))

$btnRefreshScopes.Add_Click({ Update-ZoneScopes })
Update-ZoneScopes

$btnAddScope.Add_Click({
    $zones = Get-Zones

    $prompt               = New-Object System.Windows.Forms.Form
    $prompt.Text          = 'New Scope'
    $prompt.Size          = New-Object System.Drawing.Size(300, 200)
    $prompt.StartPosition = 'CenterScreen'

    $labelName          = New-Object System.Windows.Forms.Label
    $labelName.Location = New-Object System.Drawing.Point(10, 10)
    $labelName.Size     = New-Object System.Drawing.Size(260, 20)
    $labelName.Text     = 'Scope Name'
    $prompt.Controls.Add($labelName)
    $fieldName          = New-Object System.Windows.Forms.TextBox
    $fieldName.Location = New-Object System.Drawing.Point(10, 30)
    $fieldName.Size     = New-Object System.Drawing.Size(260, 20)
    $prompt.Controls.Add($fieldName)

    $labelZone          = New-Object System.Windows.Forms.Label
    $labelZone.Location = New-Object System.Drawing.Point(10, 70)
    $labelZone.Size     = New-Object System.Drawing.Size(260, 20)
    $labelZone.Text     = 'Zone'
    $prompt.Controls.Add($labelZone)
    $fieldZone          = New-Object System.Windows.Forms.ComboBox
    $fieldZone.Location = New-Object System.Drawing.Point(10, 90)
    $fieldZone.Size     = New-Object System.Drawing.Size(260, 20)
    $fieldZone.Items.Clear()
    $fieldZone.Items.AddRange($zones)
    $prompt.Controls.Add($fieldZone)

    # Create the Yes button and its properties
    $yesButton = New-Object System.Windows.Forms.Button
    $yesButton.Location = New-Object System.Drawing.Point(60, 120)
    $yesButton.Size = New-Object System.Drawing.Size(75, 23)
    $yesButton.Text = 'Create'
    $yesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $prompt.AcceptButton = $yesButton
    $prompt.Controls.Add($yesButton)
    # Create the No button and its properties
    $noButton = New-Object System.Windows.Forms.Button
    $noButton.Location = New-Object System.Drawing.Point(165, 120)
    $noButton.Size = New-Object System.Drawing.Size(75, 23)
    $noButton.Text = 'Cancel'
    $noButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $prompt.CancelButton = $noButton
    $prompt.Controls.Add($noButton)

    $prompt.TopMost = $true

    if ($prompt.ShowDialog() -eq [System.Windows.Forms.DialogResult]::Yes) {
        $scopeName = $fieldName.Text
        $zoneName = $fieldZone.SelectedItem
        if ($scopeName -and $zoneName) {
            try {
                Add-DnsServerZoneScope -ZoneName $zoneName -Name $scopeName
                Show-Info "Scope '$scopeName' added for Zone '$zoneName'."
                Update-ZoneScopes
            }
            catch {
                Show-Error $_
            }
        } else {
            Show-Error "Please provide a value for both Name and Zone."
        }
    }
})

$btnEditRecords.Add_Click({
    if ($dgScopes.SelectedRows.Count -eq 0) { return }
    $row = $dgScopes.SelectedRows[0]

    $zoneName = $row.Cells['Zone'].Value
    $scopeName = $row.Cells['Scope Name'].Value

    $script:recordsZone = $zoneName
    $script:recordsScope = $scopeName
    Update-Records
    $records.ShowDialog()
})

$btnRemoveScope.Add_Click({
    if ($dgScopes.SelectedRows.Count -eq 0) { return }
    $row = $dgScopes.SelectedRows[0]

    $zoneName = $row.Cells['Zone'].Value
    $scopeName = $row.Cells['Scope Name'].Value

    $prompt               = New-Object System.Windows.Forms.Form
    $prompt.Text          = 'Remove Scope'
    $prompt.Size          = New-Object System.Drawing.Size(300, 200)
    $prompt.StartPosition = 'CenterScreen'

    $promptDesc          = New-Object System.Windows.Forms.Label
    $promptDesc.Location = New-Object System.Drawing.Point(10, 10)
    $promptDesc.Size     = New-Object System.Drawing.Size(260, 80)
    $promptDesc.Text     = "Are you sure you want to remove the Scope '$scopeName' for Zone '$zoneName'?"
    $prompt.Controls.Add($promptDesc)

    # Create the Yes button and its properties
    $yesButton = New-Object System.Windows.Forms.Button
    $yesButton.Location = New-Object System.Drawing.Point(60, 120)
    $yesButton.Size = New-Object System.Drawing.Size(75, 23)
    $yesButton.Text = 'Remove'
    $yesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $prompt.AcceptButton = $yesButton
    $prompt.Controls.Add($yesButton)
    # Create the No button and its properties
    $noButton = New-Object System.Windows.Forms.Button
    $noButton.Location = New-Object System.Drawing.Point(165, 120)
    $noButton.Size = New-Object System.Drawing.Size(75, 23)
    $noButton.Text = 'Cancel'
    $noButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $prompt.CancelButton = $noButton
    $prompt.Controls.Add($noButton)

    $prompt.TopMost = $true

    if ($prompt.ShowDialog() -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            Remove-DnsServerZoneScope -ZoneName $zoneName -Name $scopeName -Force
            Show-Info "Scope '$scopeName' removed for Zone '$zoneName'."
            Update-ZoneScopes
        }
        catch {
            Show-Error $_
        }
    }
})

$btnImportScopes.Add_Click({
    $import = Import-FromCsv
    if ($import -and @($import).Count -gt 0) {
        foreach ($row in $import) {
            try {
                Add-DnsServerZoneScope -ZoneName $row.Zone -Name $row.'Scope Name'
            }
            catch {
                Write-Host "Unable to add Zone Scope '$($row.'Scope Name')' for Zone '$($row.Zone)':`n$_"
            }
        }
        Show-Info "Imported Zone Scopes."
        Update-ZoneScopes
    }
})

$btnExportScopes.Add_Click({ Export-GridViewToCsv -DataGridView $dgScopes })

# --- Tab 2: DNS Client Subnets ---------------------------------------------------

$dtSubnets = New-Object System.Data.DataTable
$dtSubnets.Columns.Add('Subnet Name') | Out-Null
$dtSubnets.Columns.Add('IPv4 Address(es)') | Out-Null
$dtSubnets.Columns.Add('IPv6 Address(es)') | Out-Null

$dgSubnets = New-Object System.Windows.Forms.DataGridView
$dgSubnets.Dock = 'Top'
$dgSubnets.Height = 490
$dgSubnets.SelectionMode = 'FullRowSelect'
$dgSubnets.AutoSizeColumnsMode = 'AllCells'
$dgSubnets.EditMode = 'EditProgrammatically'
$dgSubnets.MultiSelect = $false;
$dgSubnets.AllowUserToAddRows = $false
$dgSubnets.AllowUserToDeleteRows = $false
$dgSubnets.DataSource = $dtSubnets
$tabs['DNS Client Subnets'].Controls.Add($dgSubnets)

$btnRefreshSubnets          = New-Object System.Windows.Forms.Button
$btnRefreshSubnets.Text     = 'Refresh'
$btnRefreshSubnets.Location = '10,500'
$btnRefreshSubnets.Size     = '100,30'

$btnAddSubnet               = New-Object System.Windows.Forms.Button
$btnAddSubnet.Text          = 'Add'
$btnAddSubnet.Location      = '120,500'
$btnAddSubnet.Size          = '100,30'

$btnRemoveSubnet            = New-Object System.Windows.Forms.Button
$btnRemoveSubnet.Text       = 'Remove'
$btnRemoveSubnet.Location   = '230,500'
$btnRemoveSubnet.Size       = '100,30'

$btnImportSubnets           = New-Object System.Windows.Forms.Button
$btnImportSubnets.Text      = 'Import'
$btnImportSubnets.Location  = '670,500'
$btnImportSubnets.Size      = '100,30'

$btnExportSubnets           = New-Object System.Windows.Forms.Button
$btnExportSubnets.Text      = 'Export'
$btnExportSubnets.Location  = '780,500'
$btnExportSubnets.Size      = '100,30'

$tabs['DNS Client Subnets'].Controls.AddRange(@($btnRefreshSubnets,$btnAddSubnet,$btnRemoveSubnet,$btnImportSubnets,$btnExportSubnets))

$btnRefreshSubnets.Add_Click({ Update-ClientSubnets })
Update-ClientSubnets

$btnAddSubnet.Add_Click({
    $prompt               = New-Object System.Windows.Forms.Form
    $prompt.Text          = 'New DNS Client Subnet'
    $prompt.Size          = New-Object System.Drawing.Size(300, 200)
    $prompt.StartPosition = 'CenterScreen'

    $labelName          = New-Object System.Windows.Forms.Label
    $labelName.Location = New-Object System.Drawing.Point(10, 10)
    $labelName.Size     = New-Object System.Drawing.Size(260, 20)
    $labelName.Text     = 'Client Subnet Name'
    $prompt.Controls.Add($labelName)
    $fieldName          = New-Object System.Windows.Forms.TextBox
    $fieldName.Location = New-Object System.Drawing.Point(10, 30)
    $fieldName.Size     = New-Object System.Drawing.Size(260, 20)
    $prompt.Controls.Add($fieldName)

    $labelSubnets          = New-Object System.Windows.Forms.Label
    $labelSubnets.Location = New-Object System.Drawing.Point(10, 70)
    $labelSubnets.Size     = New-Object System.Drawing.Size(260, 20)
    $labelSubnets.Text     = 'IPv4 Subnet(s) (CIDR, comma-separated)'
    $prompt.Controls.Add($labelSubnets)
    $fieldSubnets          = New-Object System.Windows.Forms.TextBox
    $fieldSubnets.Location = New-Object System.Drawing.Point(10, 90)
    $fieldSubnets.Size     = New-Object System.Drawing.Size(260, 20)
    $prompt.Controls.Add($fieldSubnets)

    # Create the Yes button and its properties
    $yesButton = New-Object System.Windows.Forms.Button
    $yesButton.Location = New-Object System.Drawing.Point(60, 120)
    $yesButton.Size = New-Object System.Drawing.Size(75, 23)
    $yesButton.Text = 'Create'
    $yesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $prompt.AcceptButton = $yesButton
    $prompt.Controls.Add($yesButton)
    # Create the No button and its properties
    $noButton = New-Object System.Windows.Forms.Button
    $noButton.Location = New-Object System.Drawing.Point(165, 120)
    $noButton.Size = New-Object System.Drawing.Size(75, 23)
    $noButton.Text = 'Cancel'
    $noButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $prompt.CancelButton = $noButton
    $prompt.Controls.Add($noButton)

    $prompt.TopMost = $true

    if ($prompt.ShowDialog() -eq [System.Windows.Forms.DialogResult]::Yes) {
        $subnetName = $fieldName.Text
        $subnets = $fieldSubnets.Text
        if ($subnetName -and $subnets) {
            try {
                Add-DnsServerClientSubnet -Name $subnetName -IPv4Subnet $subnets.Split(',')
                Show-Info "Client Subnet '$subnetName' created."
                Update-ClientSubnets
            }
            catch {
                Show-Error $_
            }
        } else {
            Show-Error "Please provide a value for both Name and Zone."
        }
    }
})

$btnRemoveSubnet.Add_Click({
    if ($dgSubnets.SelectedRows.Count -eq 0) { return }
    $row = $dgSubnets.SelectedRows[0]

    $subnetName = $row.Cells['Subnet Name'].Value

    $prompt               = New-Object System.Windows.Forms.Form
    $prompt.Text          = 'Remove DNS Client Subnet'
    $prompt.Size          = New-Object System.Drawing.Size(300, 200)
    $prompt.StartPosition = 'CenterScreen'

    $promptDesc          = New-Object System.Windows.Forms.Label
    $promptDesc.Location = New-Object System.Drawing.Point(10, 10)
    $promptDesc.Size     = New-Object System.Drawing.Size(260, 80)
    $promptDesc.Text     = "Are you sure you want to remove the DNS Client Subnet '$subnetName'?"
    $prompt.Controls.Add($promptDesc)

    # Create the Yes button and its properties
    $yesButton = New-Object System.Windows.Forms.Button
    $yesButton.Location = New-Object System.Drawing.Point(60, 120)
    $yesButton.Size = New-Object System.Drawing.Size(75, 23)
    $yesButton.Text = 'Remove'
    $yesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $prompt.AcceptButton = $yesButton
    $prompt.Controls.Add($yesButton)
    # Create the No button and its properties
    $noButton = New-Object System.Windows.Forms.Button
    $noButton.Location = New-Object System.Drawing.Point(165, 120)
    $noButton.Size = New-Object System.Drawing.Size(75, 23)
    $noButton.Text = 'Cancel'
    $noButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $prompt.CancelButton = $noButton
    $prompt.Controls.Add($noButton)

    $prompt.TopMost = $true

    if ($prompt.ShowDialog() -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            Remove-DnsServerClientSubnet -Name $subnetName -Force
            Show-Info "DNS Client Subnet '$subnetName' removed."
            Update-ClientSubnets
        }
        catch {
            Show-Error $_
        }
    }
})

$btnImportSubnets.Add_Click({
    $import = Import-FromCsv
    if ($import -and @($import).Count -gt 0) {
        foreach ($row in $import) {
            try {
                Add-DnsServerClientSubnet -Name $row.'Subnet Name' -IPv4Subnet $row.'IPv4 Address(es)'.Split(',')
            }
            catch {
                Write-Host "Unable to add DNS Client Subnet '$($row.'Subnet Name')':`n$_"
            }
        }
        Show-Info "Imported DNS Client Subnets."
        Update-ClientSubnets
    }
})

$btnExportSubnets.Add_Click({ Export-GridViewToCsv -DataGridView $dgSubnets })

# --- Tab 3: Query Resolution Policies ---------------------------------------------------

$dtPolicies = New-Object System.Data.DataTable
$dtPolicies.Columns.Add('Policy Name') | Out-Null
$dtPolicies.Columns.Add('Zone') | Out-Null
$dtPolicies.Columns.Add('Scope Name') | Out-Null
$dtPolicies.Columns.Add('Subnet Name') | Out-Null
$dtPolicies.Columns.Add('Action') | Out-Null
$dtPolicies.Columns.Add('Enabled') | Out-Null

$dgPolicies = New-Object System.Windows.Forms.DataGridView
$dgPolicies.Dock = 'Top'
$dgPolicies.Height = 490
$dgPolicies.SelectionMode = 'FullRowSelect'
$dgPolicies.AutoSizeColumnsMode = 'AllCells'
$dgPolicies.EditMode = 'EditProgrammatically'
$dgPolicies.MultiSelect = $false;
$dgPolicies.AllowUserToAddRows = $false
$dgPolicies.AllowUserToDeleteRows = $false
$dgPolicies.DataSource = $dtPolicies
$tabs['Query Resolution Policies'].Controls.Add($dgPolicies)

$btnRefreshPolicies          = New-Object System.Windows.Forms.Button
$btnRefreshPolicies.Text     = 'Refresh'
$btnRefreshPolicies.Location = '10,500'
$btnRefreshPolicies.Size     = '100,30'

$btnAddPolicy               = New-Object System.Windows.Forms.Button
$btnAddPolicy.Text          = 'Add'
$btnAddPolicy.Location      = '120,500'
$btnAddPolicy.Size          = '100,30'

$btnRemovePolicy            = New-Object System.Windows.Forms.Button
$btnRemovePolicy.Text       = 'Remove'
$btnRemovePolicy.Location   = '230,500'
$btnRemovePolicy.Size       = '100,30'

$btnImportPolicies           = New-Object System.Windows.Forms.Button
$btnImportPolicies.Text      = 'Import'
$btnImportPolicies.Location  = '670,500'
$btnImportPolicies.Size      = '100,30'

$btnExportPolicies           = New-Object System.Windows.Forms.Button
$btnExportPolicies.Text      = 'Export'
$btnExportPolicies.Location  = '780,500'
$btnExportPolicies.Size      = '100,30'

$tabs['Query Resolution Policies'].Controls.AddRange(@($btnRefreshPolicies,$btnAddPolicy,$btnRemovePolicy,$btnImportPolicies,$btnExportPolicies))

$btnRefreshPolicies.Add_Click({ Update-QueryResolutionPolicies })
Update-QueryResolutionPolicies

$btnAddPolicy.Add_Click({
    $zones = Get-Zones
    $subnets = Get-ClientSubnets | Select-Object -ExpandProperty Name

    $prompt               = New-Object System.Windows.Forms.Form
    $prompt.Text          = 'New Query Resolution Policy'
    $prompt.Size          = New-Object System.Drawing.Size(300, 400)
    $prompt.StartPosition = 'CenterScreen'

    $labelName          = New-Object System.Windows.Forms.Label
    $labelName.Location = New-Object System.Drawing.Point(10, 10)
    $labelName.Size     = New-Object System.Drawing.Size(260, 20)
    $labelName.Text     = 'Policy Name'
    $prompt.Controls.Add($labelName)
    $fieldName          = New-Object System.Windows.Forms.TextBox
    $fieldName.Location = New-Object System.Drawing.Point(10, 30)
    $fieldName.Size     = New-Object System.Drawing.Size(260, 20)
    $prompt.Controls.Add($fieldName)

    $labelZone          = New-Object System.Windows.Forms.Label
    $labelZone.Location = New-Object System.Drawing.Point(10, 70)
    $labelZone.Size     = New-Object System.Drawing.Size(260, 20)
    $labelZone.Text     = 'Zone'
    $prompt.Controls.Add($labelZone)
    $fieldZone          = New-Object System.Windows.Forms.ComboBox
    $fieldZone.Location = New-Object System.Drawing.Point(10, 90)
    $fieldZone.Size     = New-Object System.Drawing.Size(260, 20)
    $fieldZone.Items.Clear()
    $fieldZone.Items.AddRange($zones)
    $prompt.Controls.Add($fieldZone)

    $labelScope          = New-Object System.Windows.Forms.Label
    $labelScope.Location = New-Object System.Drawing.Point(10, 130)
    $labelScope.Size     = New-Object System.Drawing.Size(260, 20)
    $labelScope.Text     = 'Zone Scope'
    $prompt.Controls.Add($labelScope)
    $fieldScope          = New-Object System.Windows.Forms.ComboBox
    $fieldScope.Location = New-Object System.Drawing.Point(10, 150)
    $fieldScope.Size     = New-Object System.Drawing.Size(260, 20)
    $fieldScope.Items.Clear()
    $prompt.Controls.Add($fieldScope)

    $labelSubnet          = New-Object System.Windows.Forms.Label
    $labelSubnet.Location = New-Object System.Drawing.Point(10, 190)
    $labelSubnet.Size     = New-Object System.Drawing.Size(260, 20)
    $labelSubnet.Text     = 'Subnet Name'
    $prompt.Controls.Add($labelSubnet)
    $fieldSubnet          = New-Object System.Windows.Forms.ComboBox
    $fieldSubnet.Location = New-Object System.Drawing.Point(10, 210)
    $fieldSubnet.Size     = New-Object System.Drawing.Size(260, 20)
    $fieldSubnet.Items.Clear()
    $fieldSubnet.Items.AddRange($subnets)
    $prompt.Controls.Add($fieldSubnet)

    $labelAction          = New-Object System.Windows.Forms.Label
    $labelAction.Location = New-Object System.Drawing.Point(10, 250)
    $labelAction.Size     = New-Object System.Drawing.Size(260, 20)
    $labelAction.Text     = 'Action'
    $prompt.Controls.Add($labelAction)
    $fieldAction          = New-Object System.Windows.Forms.ComboBox
    $fieldAction.Location = New-Object System.Drawing.Point(10, 270)
    $fieldAction.Size     = New-Object System.Drawing.Size(260, 20)
    $fieldAction.Items.Clear()
    $fieldAction.Items.AddRange(@('Allow', 'Block', 'Override'))
    $prompt.Controls.Add($fieldAction)

    $fieldZone.Add_SelectedIndexChanged({
        $scopes = Get-ZoneScopes -ZoneName $fieldZone.SelectedItem | Select-Object -ExpandProperty ZoneScope
        $fieldScope.Items.Clear()
        if ($scopes.Count -gt 0) {
            $fieldScope.Items.AddRange($scopes)
        }
    })

    # Create the Yes button and its properties
    $yesButton = New-Object System.Windows.Forms.Button
    $yesButton.Location = New-Object System.Drawing.Point(60, 320)
    $yesButton.Size = New-Object System.Drawing.Size(75, 23)
    $yesButton.Text = 'Create'
    $yesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $prompt.AcceptButton = $yesButton
    $prompt.Controls.Add($yesButton)
    # Create the No button and its properties
    $noButton = New-Object System.Windows.Forms.Button
    $noButton.Location = New-Object System.Drawing.Point(165, 320)
    $noButton.Size = New-Object System.Drawing.Size(75, 23)
    $noButton.Text = 'Cancel'
    $noButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $prompt.CancelButton = $noButton
    $prompt.Controls.Add($noButton)

    $prompt.TopMost = $true

    if ($prompt.ShowDialog() -eq [System.Windows.Forms.DialogResult]::Yes) {
        $policyName = $fieldName.Text
        $zoneName = $fieldZone.SelectedItem
        $scopeName = $fieldScope.SelectedItem
        $subnetName = $fieldSubnet.SelectedItem
        $action = $fieldAction.SelectedItem
        if ($policyName -and $zoneName -and $scopeName -and $subnetName -and $action) {
            try {
                Add-DnsServerQueryResolutionPolicy -Name $policyName -ZoneName $zoneName -ZoneScope $scopeName -ClientSubnet "EQ,$subnetName" -Action $action
                Show-Info "Query Resolution Policy '$policyName' created."
                Update-QueryResolutionPolicies
            }
            catch {
                Show-Error $_
            }
        } else {
            Show-Error "Please provide values for all the fields."
        }
    }
})

$btnRemovePolicy.Add_Click({
    if ($dgPolicies.SelectedRows.Count -eq 0) { return }
    $row = $dgPolicies.SelectedRows[0]

    $policyName = $row.Cells['Policy Name'].Value
    $zoneName = $row.Cells['Zone'].Value

    $prompt               = New-Object System.Windows.Forms.Form
    $prompt.Text          = 'Remove Query Resolution Policy'
    $prompt.Size          = New-Object System.Drawing.Size(300, 200)
    $prompt.StartPosition = 'CenterScreen'

    $promptDesc          = New-Object System.Windows.Forms.Label
    $promptDesc.Location = New-Object System.Drawing.Point(10, 10)
    $promptDesc.Size     = New-Object System.Drawing.Size(260, 80)
    $promptDesc.Text     = "Are you sure you want to remove the Query Resolution Policy '$policyName' for Zone '$zoneName'?"
    $prompt.Controls.Add($promptDesc)

    # Create the Yes button and its properties
    $yesButton = New-Object System.Windows.Forms.Button
    $yesButton.Location = New-Object System.Drawing.Point(60, 120)
    $yesButton.Size = New-Object System.Drawing.Size(75, 23)
    $yesButton.Text = 'Remove'
    $yesButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $prompt.AcceptButton = $yesButton
    $prompt.Controls.Add($yesButton)
    # Create the No button and its properties
    $noButton = New-Object System.Windows.Forms.Button
    $noButton.Location = New-Object System.Drawing.Point(165, 120)
    $noButton.Size = New-Object System.Drawing.Size(75, 23)
    $noButton.Text = 'Cancel'
    $noButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $prompt.CancelButton = $noButton
    $prompt.Controls.Add($noButton)

    $prompt.TopMost = $true

    if ($prompt.ShowDialog() -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            Remove-DnsServerQueryResolutionPolicy -Name $policyName -ZoneName $zoneName -Force
            Show-Info "Query Resolution Policy '$policyName' for Zone '$zoneName' removed."
            Update-QueryResolutionPolicies
        }
        catch {
            Show-Error $_
        }
    }
})

$btnImportPolicies.Add_Click({
    $import = Import-FromCsv
    if ($import -and @($import).Count -gt 0) {
        foreach ($row in $import) {
            try {
                Add-DnsServerQueryResolutionPolicy -Name $row.'Policy Name' -ZoneName $row.Zone -ZoneScope $row.'Scope Name' -ClientSubnet "EQ,$($row.'Subnet Name')" -Action $row.Action
            }
            catch {
                Write-Host "Unable to add Query Resolution Policy '$($row.'Polcy Name')' for Zone '$($row.Zone)' and Scope '$($row.'Scope Name')':`n$_"
            }
        }
        Show-Info "Imported Query Resolution Policies."
        Update-QueryResolutionPolicies
    }
})

$btnExportPolicies.Add_Click({ Export-GridViewToCsv -DataGridView $dgPolicies })

# --- Start GUI -------------------------------------------------------------
[void] $form.ShowDialog()
