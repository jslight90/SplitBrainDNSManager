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

.LINK
    GitHub: https://github.com/jslight90

#>

Add-Type -AssemblyName System.Windows.Forms, System.Drawing

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

        $dgScopes.Rows.Clear()
        foreach ($row in $all) {
            $index = $dgScopes.Rows.Add()
            $dgScopes.Rows[$index].Cells[0].Value = $row.Name
            $dgScopes.Rows[$index].Cells[1].Value = $row.Zone
        }
    } catch {
        Show-Error "Failed loading scopes: $_"
    }
}

function Update-ClientSubnets {
    try {
        $dgSubnets.Rows.Clear()
        foreach ($row in Get-ClientSubnets) {
            $index = $dgSubnets.Rows.Add()
            $dgSubnets.Rows[$index].Cells[0].Value = $row.Name
            $dgSubnets.Rows[$index].Cells[1].Value = $row.IPV4Subnet -join ", "
            $dgSubnets.Rows[$index].Cells[2].Value = $row.IPv6Subnet -join ", "
        }
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

        $dgPolicies.Rows.Clear()
        foreach ($row in $all) {
            $index = $dgPolicies.Rows.Add()
            $dgPolicies.Rows[$index].Cells[0].Value = $row.Name
            $dgPolicies.Rows[$index].Cells[1].Value = $row.Zone
            $dgPolicies.Rows[$index].Cells[2].Value = $row.Scope
            $dgPolicies.Rows[$index].Cells[3].Value = $row.Subnet
            $dgPolicies.Rows[$index].Cells[4].Value = $row.Action
            $dgPolicies.Rows[$index].Cells[5].Value = $row.Enabled
        }
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

        $dgRecords.Rows.Clear()
        foreach ($row in $all) {
            $index = $dgRecords.Rows.Add()
            $dgRecords.Rows[$index].Cells[0].Value = $row.Name
            $dgRecords.Rows[$index].Cells[1].Value = $row.Type
            $dgRecords.Rows[$index].Cells[2].Value = $row.Data
        }
    }
    catch {
        Show-Error "Failed loading resource records: $_"
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

$dgRecords = New-Object System.Windows.Forms.DataGridView
$dgRecords.Dock = 'Top'
$dgRecords.Height = 690
$dgRecords.SelectionMode = 'FullRowSelect'
$dgRecords.AutoSizeColumnsMode = 'AllCells'
$dgRecords.EditMode = 'EditProgrammatically'
$dgRecords.MultiSelect = $false;
$dgRecords.AllowUserToAddRows = $false
$dgRecords.AllowUserToDeleteRows = $false
$dgRecords.ColumnCount = 3
$dgRecords.Columns[0].Name = "Name"
$dgRecords.Columns[1].Name = "Type"
$dgRecords.Columns[2].Name = "Data"
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

$records.Controls.AddRange(@($btnRefeshRecords,$btnAddRecord,$btnRemoveRecord))

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
                switch ($recordType) {
                    'A'     {
                        Add-DnsServerResourceRecordA -ZoneName $zoneName -ZoneScope $scopeName -Name $recordName -IPv4Address $recordData
                    }
                    'CNAME' {
                        Add-DnsServerResourceRecordCName -ZoneName $zoneName -ZoneScope $scopeName -Name $recordName -HostNameAlias $recordData
                    }
                    'TXT'   {
                        Add-DnsServerResourceRecord -Txt -ZoneName $zoneName -ZoneScope $scopeName -Name $recordName -DescriptiveText $recordData
                    }
                    'PTR'   {
                        Add-DnsServerResourceRecordPtr -ZoneName $zoneName -ZoneScope $scopeName -Name $recordName -PtrDomainName $recordData
                    }
                    Default {
                        Show-Error "Unsuppored record type: $recordType"
                    }
                }
                Show-Info "$recordType Record '$scopeName' added for Zone '$zoneName' and Scope '$scopeName'."
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


# --- Tab 1: Zone Scopes and Records ---------------------------------------------------

$dgScopes = New-Object System.Windows.Forms.DataGridView
$dgScopes.Dock = 'Top'
$dgScopes.Height = 490
$dgScopes.SelectionMode = 'FullRowSelect'
$dgScopes.AutoSizeColumnsMode = 'AllCells'
$dgScopes.EditMode = 'EditProgrammatically'
$dgScopes.MultiSelect = $false;
$dgScopes.AllowUserToAddRows = $false
$dgScopes.AllowUserToDeleteRows = $false
$dgScopes.ColumnCount = 2
$dgScopes.Columns[0].Name = "Scope Name"
$dgScopes.Columns[1].Name = "Zone"
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

$tabs['Zone Scopes and Records'].Controls.AddRange(@($btnRefreshScopes,$btnAddScope,$btnEditRecords,$btnRemoveScope))

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

# --- Tab 2: DNS Client Subnets ---------------------------------------------------

$dgSubnets = New-Object System.Windows.Forms.DataGridView
$dgSubnets.Dock = 'Top'
$dgSubnets.Height = 490
$dgSubnets.SelectionMode = 'FullRowSelect'
$dgSubnets.AutoSizeColumnsMode = 'AllCells'
$dgSubnets.EditMode = 'EditProgrammatically'
$dgSubnets.MultiSelect = $false;
$dgSubnets.AllowUserToAddRows = $false
$dgSubnets.AllowUserToDeleteRows = $false
$dgSubnets.ColumnCount = 3
$dgSubnets.Columns[0].Name = "Subnet Name"
$dgSubnets.Columns[1].Name = "IPv4 Address(es)"
$dgSubnets.Columns[2].Name = "IPv6 Address(es)"
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

$tabs['DNS Client Subnets'].Controls.AddRange(@($btnRefreshSubnets,$btnAddSubnet,$btnRemoveSubnet))

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

# --- Tab 3: Query Resolution Policies ---------------------------------------------------

$dgPolicies = New-Object System.Windows.Forms.DataGridView
$dgPolicies.Dock = 'Top'
$dgPolicies.Height = 490
$dgPolicies.SelectionMode = 'FullRowSelect'
$dgPolicies.AutoSizeColumnsMode = 'AllCells'
$dgPolicies.EditMode = 'EditProgrammatically'
$dgPolicies.MultiSelect = $false;
$dgPolicies.AllowUserToAddRows = $false
$dgPolicies.AllowUserToDeleteRows = $false
$dgPolicies.ColumnCount = 6
$dgPolicies.Columns[0].Name = "Policy Name"
$dgPolicies.Columns[1].Name = "Zone"
$dgPolicies.Columns[2].Name = "Scope Name"
$dgPolicies.Columns[3].Name = "Subnet Name"
$dgPolicies.Columns[4].Name = "Action"
$dgPolicies.Columns[5].Name = "Enabled"
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

$tabs['Query Resolution Policies'].Controls.AddRange(@($btnRefreshPolicies,$btnAddPolicy,$btnRemovePolicy))

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

# --- Start GUI -------------------------------------------------------------
[void] $form.ShowDialog()
