<#
.SYNOPSIS
  <Overview of script>
.DESCRIPTION
  <Brief description of script>
.PARAMETER ParametersFile
  <Brief description of parameter input required. Repeat this attribute if required>
.INPUTS
  <Inputs if any, otherwise state None>
.OUTPUTS
  <Outputs if any, otherwise state None - example: Log file stored in C:\Windows\Temp\<name>.log>
.NOTES
  Created by: Bipul Raman @BipulRaman
  Modified by: Bipul Raman @BipulRaman
  Modified: 12/22/2018 02:26 PM IST
  Purpose/Change: Initial script development
.EXAMPLE
  <Example goes here. Repeat this attribute for more than one example>
#>

#-------------------[PARAMETERS]-------------------

Param(
  [string] $VMDetailsInputFile = 'Extend_RDC_Backup.Input.csv',
  [string] $InputRDCBackupFile = 'InputRDCBackup.rdb',
  [string] $RDCGroupName = 'CAT',
  [string] $RDCUserAccountName = 'CAT-VMs',
  [string] $OutputRDCBackupFile = 'OutputRDCBackup.rdb'
)

#-------------------[INITIALISATIONS]-------------------

$VMDetailsInputPath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($PSScriptRoot, $VMDetailsInputFile))
$RDCBackupPath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($PSScriptRoot, $InputRDCBackupFile))
$OutputRDCBackupFilePath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($PSScriptRoot, $OutputRDCBackupFile))

$VMDetailsInput = Import-Csv -Path $VMDetailsInputPath
$RDCBackup = (Get-Content -LiteralPath $RDCBackupPath) -join "`n" | ConvertFrom-Json

#-------------------[EXECUTION]-------------------

$GroupId = ($RDCBackup.Groups | Where-Object -FilterScript { $_.Name -eq $RDCGroupName })[0].PersistentModelId
$CredentialId = ($RDCBackup.Credentials | Where-Object -FilterScript { $_.FriendlyName -eq $RDCUserAccountName })[0].PersistentModelId
$Connections = ($RDCBackup.Connections | Where-Object -FilterScript { ($_.GroupId -eq $GroupId) -and ($_.CredentialsId -eq $CredentialId) })
$ConnectionSample = ($RDCBackup.Connections | Where-Object -FilterScript { ($_.GroupId -eq $GroupId) -and ($_.CredentialsId -eq $CredentialId) })[0]

foreach ($eachVM in $VMDetailsInput) {
  Write-Host "Checking $($eachVM.VMName)"
  $ExistingConnection = ($Connections | Where-Object -FilterScript { ($_.HostName -eq ($eachVM.IP)) -or ($_.FriendlyName -eq ($eachVM.VMName)) })
  if ($null -eq $ExistingConnection) {  
    $NewConnection = $ConnectionSample.PSObject.Copy()
    $NewConnection.HostName = $eachVM.IP
    $NewConnection.FriendlyName = $eachVM.VMName
    $NewConnection.PersistentModelId = [guid]::NewGuid()
    $RDCBackup.Connections += $NewConnection
  }
  else {
    Write-Host "Skipping this one because connection already exist with same IP - $($eachVM.IP) / VMName - $($eachVM.VMName)"
  }
}

$RDCBackup | ConvertTo-Json -Depth 50 | Out-File -FilePath $OutputRDCBackupFilePath