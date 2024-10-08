param(
    # Whatever name you chose will be the name of the GPO
    [Parameter(Mandatory)]
    [string] $gpoName,

    # Name of the computer that you want to gather services from
    [Parameter(Mandatory)]
    [string] $serverName,

    # Name of the group to which access should be granted
    [Parameter(Mandatory)]
    [string] $AdGroup,

    # FQDN of the domain e.g. domain.local
    # Otional, if no domain is given it will retrieve it from the running computer
    [string] $domainName,

    # Link new policy to closest OU of the server
    [switch]$AutoLinkGPO
)

# All columns to be sorted by clicking the header
function SortListView {
    Param(
        [System.Windows.Forms.ListView]$sender,
        $column
    )
    $temp = $sender.Items | Foreach-Object { $_ }
    $Script:SortingDescending = !$Script:SortingDescending
    $sender.Items.Clear()
    $sender.ShowGroups = $false
    $sender.Sorting = 'none'
    $sender.Items.AddRange(($temp | Sort-Object -Descending:$script:SortingDescending -Property @{ Expression={ $_.SubItems[$column].Text } }))
}

Function Get-Domain {
   $CompInfo = Get-WmiObject -Namespace root\cimv2 -Class Win32_ComputerSystem | Select Name, Domain
   Return $CompInfo.Domain.ToUpper()
}


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($domainFQDN))
{
    $domainName = Get-Domain
}

$testName = Get-ADComputer $serverName

$adServiceGroup = Get-ADGroup $AdGroup

Invoke-Command -ComputerName $testName.DNSHostName {gpupdate /force}

# Get services from remote computer
$serviceList = Invoke-Command -ComputerName $testName.DNSHostName {get-service | select name, DisplayName} | select name, DisplayName

# Create form with list view of services
$form = New-Object System.Windows.Forms.Form
$form.Text = "Services List"
$form.Size = New-Object System.Drawing.Size(410,600)
$form.StartPosition = "CenterScreen"

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75,500)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150,500)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.Controls.Add($CancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = "Please select services to be locked down:"
$form.Controls.Add($label)

$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10,40)
$listView.Size = New-Object System.Drawing.Size(380,20)
$listView.add_ColumnClick({SortListView $this $_.Column})
$listView.View = "Details"

$test = New-Object System.Windows.Forms.ListView

$listView.Columns.Add("Name") | Out-Null
$listView.Columns[0].Width = -1
$listView.Columns.Add("DisplayName") | Out-Null
$listView.Columns[1].Width = -1

foreach($service in $serviceList)
{
    $listViewItem = [System.Windows.Forms.ListViewItem]::new("$($service.Name)")
    $listViewSubItem = [System.Windows.Forms.ListViewItem+ListViewSubItem]::new($listViewItem,"$($service.DisplayName)")
    $listViewItem.SubItems.Add($listViewSubItem) | Out-Null
    $listView.Items.Add($listViewItem) | Out-Null
}

$listView.Height = 450
$form.Controls.Add($listView)
$form.TopMost = $true

# Display the form
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $services = $listView.SelectedItems

    # Create txt for the GPO inf file
    $text = "[Unicode]
    Unicode=yes
    [Version]
    signature=`"`$CHICAGO$`"
    Revision=1
    [Service General Setting]"

    foreach ($item in $services)
    {
        $text += "`n`"$($item.Text)`",2,`"D:AR(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;RPWPDTRC;;;$($adServiceGroup.SID))S:(AU;FA;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;WD)`""
    }

    # Create GPO
    $dc = Get-ADDomainController
    $newGpo = New-GPO -Name "$gpoName" -Server $dc

    # Add settings
    New-Item "\\$domainName\SYSVOL\$domainName\Policies\{$($newGpo.id)}\Machine\Microsoft\Windows NT\SecEdit\" -ItemType directory
    New-Item "\\$domainName\SYSVOL\$domainName\Policies\{$($newGpo.id)}\Machine\Scripts\Shutdown\" -ItemType directory
    New-Item "\\$domainName\SYSVOL\$domainName\Policies\{$($newGpo.id)}\Machine\Scripts\Startup\" -ItemType directory

    # Output inf file
    $text | Out-File "\\$domainName\SYSVOL\$domainName\Policies\{$($newGpo.id)}\Machine\Microsoft\Windows NT\SecEdit\GptTmpl.inf"

    $domainName1 = ($domainName.Split("."))[0]
    $domainName2 = ($domainName.Split("."))[1]

    # Set the required GUIDs for Client Side Extensions
    Set-ADObject "CN={$($newGpo.id)},CN=Policies,CN=System,DC=$domainName1,DC=$domainName2" -Replace @{gPCMachineExtensionNames="[{827D319E-6EAC-11D2-A4EA-00C04F79F83A}{803E14A0-B4FB-11D0-A0D0-00A0C90F574B}]"}

    # Force AD to process updates
    $newGpo | Set-GPRegistryValue -Key HKLM\SOFTWARE -ValueName "Default" -Value "" -Type String -Server $dc
    $newGpo | Remove-GPRegistryValue -Key HKLM\SOFTWARE -ValueName "Default" -Server $dc

    if ($AutoLinkGPO)
    {
        $ou = $testName.DistinguishedName
        $ou = $ou.Remove(0, $ou.IndexOf(",") + 1)

        $newGpo | New-GPLink -Target $ou -LinkEnabled Yes -Server $dc

        Invoke-Command -ComputerName $testName.DNSHostName {gpupdate /force}
    }
}
