<#
.SYNOPSIS
  A solution to add VMs in bulk to Microsoft Remote Desktop Windows Store App.
.DESCRIPTION
  Microsoft Remote Desktop (From Windows Store App) currently do not support bulk addition of VMs. However there is a way to restore VMs from Backup. This PowerShell script will read the current backup and extend the backup with additional required VM entries from CSV input file.
  Please refer to ReadMe file at https://github.com/BipulRaman/Extend-RDC-Backup for more details about execution steps, inputs and output.
.NOTES
  Scripted with ♥ in India by Bipul Raman @BipulRaman
#>

#-------------------[PARAMETERS]-------------------

Param(
  [string] $VMDetailsInputFile = 'Extend_RDC_Backup.Input.csv',
  [string] $InputRDCBackupFile = 'InputRDCBackup.rdb',
  [string] $RDCGroupName = 'Sample',
  [string] $RDCUserAccountName = 'SampleAccount',
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