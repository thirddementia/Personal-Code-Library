# 
   Name: InstallNewCert.ps1
   Date: 9/3/2021
   Purpose: Replace a certificate on a server
#>
Param ($FullFileName)

Add-Type -AssemblyName System.Windows.Forms
Import-Module -Name WebAdministration

#Get Computer Certificates Currently Installed
$CertsResults = Get-ChildItem -Path Cert:\LocalMachine\My
$fileCount = 0
ForEach ($Item in $CertsResults) {
	$fileCount = $fileCount + 1
	Write-Host $fileCount " " $Item.Subject
}
$fileCount = $fileCount + 1
Write-Host $fileCount "  Install new certificate"
$fileCount = $fileCount + 1
Write-Host $fileCount "  QUIT!"
Write-Host ""
$FolderChoice = $null
while (!$FolderChoice) {
	$FolderChoice = Read-Host -Prompt "Which Certificate would you like to Replace?" 
}
if ($FolderChoice -eq $fileCount) {
	Write-Host "Script aborted."
	Break
}
ElseIf ($FolderChoice -eq $fileCount-1 ) {
	$Replace = $false
}
Else {
	$Replace = $True
	$OldCertSubject = $CertsResults[$FolderChoice - 1].Subject
	$OldCertThumbprint = $CertsResults[$FolderChoice - 1].Thumbprint
}
#Write-Host $FolderChoice
#Write-Host $Replace

#Get New Certificate
if (!$FullFileName) {
	$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
	$null = $FileBrowser.ShowDialog()
	$FullFileName = $FileBrowser.FileName
}
#Write-Host $FullFileName
Write-Host ""
Write-Host ""

$mypwd = Get-Credential -UserName 'Enter Certificate Password' -Message 'Enter Certificate Password'

Try {
	Import-PfxCertificate -FilePath $FullFileName -CertStoreLocation Cert:\LocalMachine\My -Password $mypwd.Password
}
Catch {
	$MyErrorString = "ERROR: " + $_.Exception.Message + " occured at Import New Certificate"
	Write-Host $MyErrorString
}

Write-Host ""
Write-Host ""

#If replacing then Unbind old certificate
if ($Replace) {
	#Get Ports Currently Bound
	$CertsResults = Get-ChildItem -Path IIS:SSLBindings
	$fileCount = 0
	ForEach ($Item in $CertsResults) {
		$fileCount = $fileCount + 1
		Write-Host $fileCount " " $Item.Port
	}
	$fileCount = $fileCount + 1
	Write-Host $fileCount "  QUIT!"
	Write-Host ""
	$FolderChoice = $null
	while (!$FolderChoice) {
		$FolderChoice = Read-Host -Prompt "Which port does it need to be bound to?" 
	}
	if ($FolderChoice -eq $fileCount) {
		Write-Host "Script aborted."
		Break
	}
	Else {
		$Port = $CertsResults[$FolderChoice - 1].Port
	}
	#Write-Host $Port
}

#Delete Old Certificate
If ($Replace) {
	Remove-Item -path "IIS:\SslBindings\0.0.0.0!$Port"
}

#Get new cert thumbprint
$CertsResults = Get-ChildItem -Path Cert:\LocalMachine\My
$fileCount = 0
ForEach ($Item in $CertsResults) {
	$fileCount = $fileCount + 1
}
$NewCertThumbprint = $CertsResults[$fileCount-1].Thumbprint
#Write-Host $fileCount
#Write-Host $NewCertThumbprint

#Bind new certificate
Try {
	get-item cert:\LocalMachine\MY\1CD953E1640521F03E00FA1686454C5441BEB59D | new-item IIS:\SslBindings\0.0.0.0!$Port
}
Catch {
	$MyErrorString = "ERROR: " + $_.Exception.Message + " occured at Bind New Certificate"
	Write-Host $MyErrorString
}

#Delete Old certificate
Try {
	Remove-Item -path Cert:\LocalMachine\My\$OldCertThumbprint
}
Catch {
	$MyErrorString = "ERROR: " + $_.Exception.Message + " occured at Delete Old Certificate"
	Write-Host $MyErrorString
}

Write-Host "Verify binding set by running: netsh http show sslcert"
