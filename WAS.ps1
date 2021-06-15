##########################
### Andreys AUTOMATION ###
##########################

function Show-Menu
{
    param (
        [string]$Title = 'Project startup-kit'
    )
    Clear-Host
    Write-Host "================ $Title ================"
    
    Write-Host "1: Press '1' for Architactural survey template."
    Write-Host "2: Press '2' for Firewall rules checks template."
    Write-Host "3: Press '3' for Windows Hardening survey template."
    Write-Host "4: Press '4' for RHEL Hardening survey template."
    Write-Host "Q: Press 'Q' to quit."
}

# Variables
$project_name = read-host "Enter Project name"
$client_firm_name = read-host "Enter client firm name"
$templates_path = read-host "Enter reports templates path"
$year = '2021'
$folder_already_exists = 'folder already exists'
$file_already_exists = 'file already exists'

do
 {
     Show-Menu
     $selection = Read-Host "Please make a selection"
     switch ($selection)
     {
         '1' {
			 if ( -not (Test-Path -Path $project_name -PathType Container)) {
					New-Item -Path $project_name -Type Directory
				}	else {
				    Write-Host $project_name $folder_already_exists -ForegroundColor Green
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_ARCH' -PathType Container)) {
					New-Item -Path $project_name\$project_name'_ARCH' -Type Directory
				}	else {
				    Write-Host $project_name\$project_name'_ARCH' $folder_already_exists -ForegroundColor Green
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_ARCH\'$client_firm_name'_'$project_name'_ARCH_'$year'_01.vsdx' -PathType Leaf)) {
					New-Item -Path $project_name\$project_name'_ARCH\'$client_firm_name'_'$project_name'_ARCH_'$year'_01.vsdx' -ItemType File
				}	else {
				    Write-Host $project_name\$project_name'_ARCH\'$client_firm_name'_'$project_name'_ARCH_'$year'_01.vsdx' $file_already_exists -ForegroundColor Green
				}
			 if ( -not (Test-Path -Path $project_name\$client_firm_name'_'$project_name'_ARCH_'$year'_01.docx' -PathType Leaf)) {
					Copy-Item $templates_path'ARCH.docx' $project_name\$client_firm_name'_'$project_name'_ARCH_'$year'_01.docx'
				}	else {
				    Write-Host $project_name\$client_firm_name'_'$project_name'_ARCH_'$year'_01.docx' $file_already_exists -ForegroundColor Green
				}
         }
		 
		 '2' {
			 if ( -not (Test-Path -Path $project_name -PathType Container)) {
					New-Item -Path $project_name -Type Directory
				}	else {
				    Write-Host $project_name $folder_already_exists -ForegroundColor Green
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_FW' -PathType Container)) {
					New-Item -Path $project_name\$project_name'_FW' -Type Directory
				}	else {
				    Write-Host $project_name\$project_name'_FW' $folder_already_exists -ForegroundColor Green
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_FW\Findings' -PathType Container)) {
					New-Item -Path $project_name\$project_name'_FW\Findings' -Type Directory
				}	else {
				    Write-Host $project_name\$project_name'_FW\Findings' $folder_already_exists -ForegroundColor Green
				}
			 if ( -not (Test-Path -Path $project_name\'Servers_INFO.txt' -PathType Leaf)) {
					New-Item -Path $project_name\'Servers_INFO.txt' -ItemType File
				}	else {
				    Write-Host $project_name\'Servers_INFO.txt' $file_already_exists -ForegroundColor Green
				}
			 if ( -not (Test-Path -Path $project_name\$client_firm_name'_'$project_name'_FW_'$year'_01.docx' -PathType Leaf)) {
					Copy-Item $templates_path'FW.docx' $project_name\$client_firm_name'_'$project_name'_FW_'$year'_01.docx'
				}	else {
				    Write-Host $project_name\$client_firm_name'_'$project_name'_FW_'$year'_01.docx' $file_already_exists -ForegroundColor Green
				}
         }
		 
		 '3' {
			 if ( -not (Test-Path -Path $project_name -PathType Container)) {
					New-Item -Path $project_name -Type Directory
				}	else {
				    Write-Host $project_name $folder_already_exists -ForegroundColor Green
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_HRD_WIN' -PathType Container)) {
					New-Item -Path $project_name\$project_name'_HRD_WIN' -Type Directory
				}	else {
				    Write-Host $project_name\$project_name'_HRD_WIN' $folder_already_exists -ForegroundColor Green
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_HRD_WIN\Findings' -PathType Container)) {
					New-Item -Path $project_name\$project_name'_HRD_WIN\Findings' -Type Directory
				}	else {
				    Write-Host $project_name\$project_name'_HRD_WIN\Findings' $folder_already_exists -ForegroundColor Green
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_HRD_WIN\Recieved_Script_Outputs' -PathType Container)) {
					New-Item -Path $project_name\$project_name'_HRD_WIN\Recieved_Script_Outputs' -Type Directory
				}	else {
				    Write-Host $project_name\$project_name'_HRD_WIN\Recieved_Script_Outputs' $folder_already_exists -ForegroundColor Green
				}
			 if ( -not (Test-Path -Path $project_name\$client_firm_name'_'$project_name'_HRD_WIN_'$year'_01.docx' -PathType Leaf)) {
					Copy-Item $templates_path'HRD_WIN.docx' $project_name\$client_firm_name'_'$project_name'_HRD_WIN_'$year'_01.docx'
				}	else {
				    Write-Host $project_name\$client_firm_name'_'$project_name'_HRD_WIN_'$year'_01.docx' $file_already_exists -ForegroundColor Green
				}
		 }
		 '4' {
			 if ( -not (Test-Path -Path $project_name -PathType Container)) {
					New-Item -Path $project_name -Type Directory
				}	else {
				    Write-Host $project_name $folder_already_exists -ForegroundColor Green
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_HRD_RHEL' -PathType Container)) {
					New-Item -Path $project_name\$project_name'_HRD_RHEL' -Type Directory
				}	else {
				    Write-Host $project_name\$project_name'_HRD_RHEL' $folder_already_exists -ForegroundColor Green
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_HRD_RHEL\Findings' -PathType Container)) {
					New-Item -Path $project_name\$project_name'_HRD_RHEL\Findings' -Type Directory
				}	else {
				    Write-Host $project_name\$project_name'_HRD_RHEL\Findings' $folder_already_exists -ForegroundColor Green
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_HRD_RHEL\Recieved_Script_Outputs' -PathType Container)) {
					New-Item -Path $project_name\$project_name'_HRD_RHEL\Recieved_Script_Outputs' -Type Directory
				}	else {
				    Write-Host $project_name\$project_name'_HRD_RHEL\Recieved_Script_Outputs' $folder_already_exists -ForegroundColor Green
				}
			 if ( -not (Test-Path -Path $project_name\$client_firm_name'_'$project_name'_HRD_RHEL_'$year'_01.docx' -PathType Leaf)) {
					Copy-Item $templates_path'HRD_RHEL.docx' $project_name\$client_firm_name'_'$project_name'_HRD_RHEL_'$year'_01.docx'
				}	else {
				    Write-Host $project_name\$client_firm_name'_'$project_name'_HRD_RHEL_'$year'_01.docx' $file_already_exists -ForegroundColor Green
				}
         }
     }
     pause
 }
 until ($selection -eq 'q')
