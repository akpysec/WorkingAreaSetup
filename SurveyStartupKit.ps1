###########################################
### Andriy Kolihanovs Survey AUTOMATION ###
###########################################

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
    Write-Host "U: Press 'U' for Updating values inside a survey template."
    Write-Host "Q: Press 'Q' to quit."
}

# Variables
$project_name = read-host "Enter Project name"
$where_to_create = read-host "Enter path where to create the Project" $project_name
$client_firm_name = read-host "Enter client firm name"
Write-Host "##########################" -ForegroundColor Magenta
Write-Host "Available Report Paths - "  -ForegroundColor Magenta
Write-Host "##########################" -ForegroundColor Magenta
Get-Childitem ".\report_templates" -recurse -Directory | % {Write-Host $_.FullName -ForegroundColor Green}
Write-Host "----------------------------------------------------------" -ForegroundColor Magenta
$templates_path = read-host "Enter reports templates path"
$year = '2021'
$folder_already_exists = 'folder already exists'
$file_already_exists = 'file already exists'
$arch_vsdx_file_path = Get-Childitem ".\report_templates" -File

do
 {
     Show-Menu
     $selection = Read-Host "Please make a selection"
     switch ($selection)
     {
         '1' {
			 if ( -not (Test-Path -Path $where_to_create\$project_name -PathType Container)) {
					Write-Host 'Created Project Directory - ' $project_name -ForegroundColor Green
					New-Item -Path $where_to_create\$project_name -Type Directory | out-null
				}	else {
				    Write-Host $where_to_create\$project_name $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\$project_name'_ARCH' -PathType Container)) {
					Write-Host 'Created Sub Directory in Projects Folder for ARCH' -ForegroundColor Green
					New-Item -Path $where_to_create\$project_name\$project_name'_ARCH' -Type Directory | out-null
				}	else {
				    Write-Host $where_to_create\$project_name\$project_name'_ARCH' $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\'Servers_INFO.txt' -PathType Leaf)) {
					Write-Host 'Created Servers INFO file in Project Directory' -ForegroundColor Green
					New-Item -Path $where_to_create\$project_name\'Servers_INFO.txt' -ItemType File | out-null
					Add-Content -Path $where_to_create\$project_name\'Servers_INFO.txt' -Value "ADD HERE: HOSTNAME and IP addresses of the servers for convenience"
				}	else {
				    Write-Host $where_to_create\$project_name\'Servers_INFO.txt' $file_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\$project_name'_ARCH\'$client_firm_name'_'$project_name'_ARCH_'$year'_01.vsdx' -PathType Leaf)) {
					Write-Host 'Created Visio empty file at Directory ARCH' -ForegroundColor Green
					Copy-Item -Path $arch_vsdx_file_path.FullName $where_to_create\$project_name\$project_name'_ARCH\'$client_firm_name'_'$project_name'_ARCH_'$year'_01.vsdx' | out-null
				}	else {
				    Write-Host $where_to_create\$project_name\$project_name'_ARCH\'$client_firm_name'_'$project_name'_ARCH_'$year'_01.vsdx' $file_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\$client_firm_name'_'$project_name'_ARCH_'$year'_01.docx' -PathType Leaf)) {
					Write-Host 'Created Template Report in ' $project_name 'Directory' -ForegroundColor Green
					Copy-Item $templates_path'\ARCH.docx' $where_to_create\$project_name\$client_firm_name'_'$project_name'_ARCH_'$year'_01.docx' | out-null
				}	else {
				    Write-Host $where_to_create\$project_name\$client_firm_name'_'$project_name'_ARCH_'$year'_01.docx' $file_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\$client_firm_name'_'$project_name'_HIGH_ALERT_specify-type_ARCH_'$year'_01.docx' -PathType Leaf)) {
					Write-Host 'Copied template file High Alert in Project Directory' -ForegroundColor Green
					Copy-Item $templates_path'\HIGH_ALERT.docx' $where_to_create\$project_name\$client_firm_name'_'$project_name'_HIGH_ALERT_specify-type_ARCH_'$year'_01.docx' | out-null
				}	else {
				    Write-Host $where_to_create\$project_name\$client_firm_name'_'$project_name'_HIGH_ALERT_specify-type_ARCH_'$year'_01.docx' $file_already_exists -ForegroundColor Magenta
				}
         }
		 '2' {
			 if ( -not (Test-Path -Path $where_to_create\$project_name -PathType Container)) {
					Write-Host 'Created Project Directory - ' $project_name -ForegroundColor Green
					New-Item -Path $where_to_create\$project_name -Type Directory | out-null
				}	else {
				    Write-Host $where_to_create\$project_name $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\$project_name'_FW' -PathType Container)) {
					Write-Host 'Created Sub Directory in Projects Folder for FW' -ForegroundColor Green
					New-Item -Path $where_to_create\$project_name\$project_name'_FW' -Type Directory | out-null
				}	else {
				    Write-Host $where_to_create\$project_name\$project_name'_FW' $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\$project_name'_FW\Evidence' -PathType Container)) {
					Write-Host 'Created Evidence folder in Projects Sub Directory' -ForegroundColor Green
					New-Item -Path $where_to_create\$project_name\$project_name'_FW\Evidence' -Type Directory | out-null
				}	else {
				    Write-Host $where_to_create\$project_name\$project_name'_FW\Evidence' $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\'Servers_INFO.txt' -PathType Leaf)) {
					Write-Host 'Created Servers INFO file in Project Directory' -ForegroundColor Green
					New-Item -Path $where_to_create\$project_name\'Servers_INFO.txt' -ItemType File | out-null
					Add-Content -Path $where_to_create\$project_name\'Servers_INFO.txt' -Value "ADD HERE: HOSTNAME and IP addresses of the servers for convenience"
				}	else {
				    Write-Host $where_to_create\$project_name\'Servers_INFO.txt' $file_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\$client_firm_name'_'$project_name'_FW_'$year'_01.docx' -PathType Leaf)) {
					Write-Host 'Copied template file in Project Directory' -ForegroundColor Green
					Copy-Item $templates_path'\FW.docx' $where_to_create\$project_name\$client_firm_name'_'$project_name'_FW_'$year'_01.docx' | out-null
				}	else {
				    Write-Host $where_to_create\$project_name\$client_firm_name'_'$project_name'_FW_'$year'_01.docx' $file_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\$client_firm_name'_'$project_name'_HIGH_ALERT_specify-type_FW_'$year'_01.docx' -PathType Leaf)) {
					Write-Host 'Copied template file High Alert in Project Directory' -ForegroundColor Green
					Copy-Item $templates_path'\HIGH_ALERT.docx' $where_to_create\$project_name\$client_firm_name'_'$project_name'_HIGH_ALERT_specify-type_FW_'$year'_01.docx' | out-null
				}	else {
				    Write-Host $where_to_create\$project_name\$client_firm_name'_'$project_name'_HIGH_ALERT_specify-type_FW_'$year'_01.docx' $file_already_exists -ForegroundColor Magenta
				}
         } 
		 '3' {
			 if ( -not (Test-Path -Path $where_to_create\$project_name -PathType Container)) {
					Write-Host 'Created Project Directory - ' $project_name -ForegroundColor Green
					New-Item -Path $where_to_create\$project_name -Type Directory | out-null
				}	else {
				    Write-Host $where_to_create\$project_name $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\$project_name'_HRD_WIN' -PathType Container)) {
					Write-Host 'Created Sub Directory in Projects Folder for HRD-WIN' -ForegroundColor Green
					New-Item -Path $where_to_create\$project_name\$project_name'_HRD_WIN' -Type Directory | out-null
				}	else {
				    Write-Host $where_to_create\$project_name\$project_name'_HRD_WIN' $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\$project_name'_HRD_WIN\Evidence' -PathType Container)) {
					Write-Host 'Created Evidence folder in Projects Sub Directory' -ForegroundColor Green
					New-Item -Path $where_to_create\$project_name\$project_name'_HRD_WIN\Evidence' -Type Directory | out-null
				}	else {
				    Write-Host $where_to_create\$project_name\$project_name'_HRD_WIN\Evidence' $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\$project_name'_HRD_WIN\Recieved_Script_Outputs' -PathType Container)) {
					Write-Host 'Created folder for Recieved Script Outputs in Projects Sub Directory' -ForegroundColor Green
					New-Item -Path $where_to_create\$project_name\$project_name'_HRD_WIN\Recieved_Script_Outputs' -Type Directory | out-null
				}	else {
				    Write-Host $where_to_create\$project_name\$project_name'_HRD_WIN\Recieved_Script_Outputs' $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\$client_firm_name'_'$project_name'_HRD_WIN_'$year'_01.docx' -PathType Leaf)) {
					Write-Host 'Copied template file in Project Directory' -ForegroundColor Green
					Copy-Item $templates_path'\HRD_WIN.docx' $where_to_create\$project_name\$client_firm_name'_'$project_name'_HRD_WIN_'$year'_01.docx' | out-null
				}	else {
				    Write-Host $where_to_create\$project_name\$client_firm_name'_'$project_name'_HRD_WIN_'$year'_01.docx' $file_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\$client_firm_name'_'$project_name'_HIGH_ALERT_specify-type_HRD_WIN_'$year'_01.docx' -PathType Leaf)) {
					Write-Host 'Copied template file High Alert in Project Directory' -ForegroundColor Green
					Copy-Item $templates_path'\HIGH_ALERT.docx' $where_to_create\$project_name\$client_firm_name'_'$project_name'_HIGH_ALERT_specify-type_HRD_WIN_'$year'_01.docx' | out-null
				}	else {
				    Write-Host $where_to_create\$project_name\$client_firm_name'_'$project_name'_HIGH_ALERT_specify-type_HRD_WIN_'$year'_01.docx' $file_already_exists -ForegroundColor Magenta
				}
		 }
		 '4' {
			 if ( -not (Test-Path -Path $where_to_create\$project_name -PathType Container)) {
					Write-Host 'Created Project Directory - ' $project_name -ForegroundColor Green
					New-Item -Path $where_to_create\$project_name -Type Directory | out-null
				}	else {
				    Write-Host $where_to_create\$project_name $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\$project_name'_HRD_RHEL' -PathType Container)) {
					Write-Host 'Created Sub Directory in Projects Folder for HRD-RHEL' -ForegroundColor Green
					New-Item -Path $where_to_create\$project_name\$project_name'_HRD_RHEL' -Type Directory | out-null
				}	else {
				    Write-Host $where_to_create\$project_name\$project_name'_HRD_RHEL' $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\$project_name'_HRD_RHEL\Evidence' -PathType Container)) {
					Write-Host 'Created Evidence folder in Projects Sub Directory' -ForegroundColor Green
					New-Item -Path $where_to_create\$project_name\$project_name'_HRD_RHEL\Evidence' -Type Directory | out-null
				}	else {
				    Write-Host $where_to_create\$project_name\$project_name'_HRD_RHEL\Evidence' $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\$project_name'_HRD_RHEL\Recieved_Script_Outputs' -PathType Container)) {
					Write-Host 'Created folder for Recieved Script Outputs in Projects Sub Directory' -ForegroundColor Green
					New-Item -Path $where_to_create\$project_name\$project_name'_HRD_RHEL\Recieved_Script_Outputs' -Type Directory | out-null
				}	else {
				    Write-Host $where_to_create\$project_name\$project_name'_HRD_RHEL\Recieved_Script_Outputs' $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\$client_firm_name'_'$project_name'_HRD_RHEL_'$year'_01.docx' -PathType Leaf)) {
					Write-Host 'Copied template file in Project Directory' -ForegroundColor Green
					Copy-Item $templates_path'\HRD_RHEL.docx' $where_to_create\$project_name\$client_firm_name'_'$project_name'_HRD_RHEL_'$year'_01.docx' | out-null
				}	else {
				    Write-Host $where_to_create\$project_name\$client_firm_name'_'$project_name'_HRD_RHEL_'$year'_01.docx' $file_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $where_to_create\$project_name\$client_firm_name'_'$project_name'_HIGH_ALERT_specify-type_HRD_RHEL_'$year'_01.docx' -PathType Leaf)) {
					Write-Host 'Copied template file High Alert in Project Directory' -ForegroundColor Green
					Copy-Item $templates_path'\HIGH_ALERT.docx' $where_to_create\$project_name\$client_firm_name'_'$project_name'_HIGH_ALERT_specify-type_HRD_RHEL_'$year'_01.docx' | out-null
				}	else {
				    Write-Host $where_to_create\$project_name\$client_firm_name'_'$project_name'_HIGH_ALERT_specify-type_HRD_RHEL_'$year'_01.docx' $file_already_exists -ForegroundColor Magenta
				}
         }		 
		 'U' {
				$date_today = read-host "Enter Date you want to be displayed in the template report"
				$month = read-host "Enter Month you want to be displayed in the template report (he)"
				$client_firm_name_in_template = read-host "Enter Clients name you want to be displayed in the template report (he)"
				$consultant_name = read-host "Enter a Consultant Name you want to be displayed in the template report (he)"
				$project_name_t = read-host "Enter a Project Name you want to be displayed in the template report (he)"
				$system_manager_name = read-host "Enter a Systems Manager Name you want to be displayed in the template report (he)"
				#$templates_copied_to_dst_path = $where_to_create\$project_name # multi-folders: "C:\fso1*", "C:\fso2*"
				$fileType = "*.docx"           # *.docx will take all .doc* files

				$textToReplace = @{
				# "TextToFind" = "TextToReplaceWith"
				"DATE" = $date_today
				"CONSULTANT_NAME" = $consultant_name
				"PROJECT_NAME" = $project_name_t
				"YEAR" = $year
				"MONTH" = $month
				"SYSTEM_MANAGER" = $system_manager_name
				"CLIENT_NAME" = $client_firm_name_in_template
				}

				$word = New-Object -ComObject Word.Application
				$word.Visible = $false

				#region Find/Replace parameters
				$matchCase = $true
				$matchWholeWord = $true
				$matchWildcards = $false
				$matchSoundsLike = $false
				$matchAllWordForms = $false
				$forward = $true
				$findWrap = [Microsoft.Office.Interop.Word.WdFindWrap]::wdFindContinue
				$format = $false
				$replace = [Microsoft.Office.Interop.Word.WdReplace]::wdReplaceOne
				#endregion

				$countf = 0 #count files
				$countr = 0 #count replacements per file
				$counta = 0 #count all replacements

				Function findAndReplace($objFind, $FindText, $ReplaceWith) {
					#simple Find and Replace to execute on a Find object
					#we let the function return (True/False) to count the replacements
					$objFind.Execute($FindText, $matchCase, $matchWholeWord, $matchWildCards, $matchSoundsLike, $matchAllWordForms, $forward, $findWrap, $format, $ReplaceWith, $replace) #> $null
				}

				Function findAndReplaceAll($objFind, $FindText, $ReplaceWith) {
					#make sure we replace all occurrences (while we find a match)
					$count = 0
					$count += findAndReplace $objFind $FindText $ReplaceWith
					While ($objFind.Found) {
						$count += findAndReplace $objFind $FindText $ReplaceWith
					}
					return $count
				}

				Function findAndReplaceMultiple($objFind, $lookupTable) {
					#apply multiple Find and Replace on the same Find object
					$count = 0
					$lookupTable.GetEnumerator() | ForEach-Object {
						$count += findAndReplaceAll $objFind $_.Key $_.Value
					}
					return $count
				}

				Function findAndReplaceWholeDoc($Document, $lookupTable) {
					$count = 0
					# Loop through each StoryRange
					ForEach ($storyRge in $Document.StoryRanges) {
						Do {
							$count += findAndReplaceMultiple $storyRge.Find $lookupTable
							#check for linked Ranges
							$storyRge = $storyRge.NextStoryRange
						} Until (!$storyRge) #null is False

					}

					$shapes = $Document.Sections.Item(1).Headers.Item(1).Shapes
					If ($shapes.Count) {
						ForEach ($shape in $shapes | Where {[bool]$_.TextFrame.HasText}) {
							$count += findAndReplaceMultiple $shape.TextFrame.TextRange.Find $lookupTable
						}
					}
					return $count
				}

				Function processDoc {
					$doc = $word.Documents.Open($_.FullName)
					$count = findAndReplaceWholeDoc $doc $textToReplace
					$doc.Close([ref]$true)
					return $count
				}

				$sw = [Diagnostics.Stopwatch]::StartNew()
				Get-ChildItem -Path $where_to_create\$project_name -Recurse -Filter $fileType | ForEach-Object { 
				  Write-Host "Processing \`"$($_.Name)\`"..." # - ForegroundColor Yellow
				  $countr = processDoc
				  Write-Host "$countr replacements made." # - ForegroundColor Cyan
				  $counta += $countr
				  $countf++
				}
				$sw.Stop()
				$elapsed = $sw.Elapsed.toString()
				Write-Host "`nDone. $countf files processed in $elapsed" # - ForegroundColor Green
				Write-Host "$counta replacements made in total." # - ForegroundColor Green

				$word.Quit()
				$word = $null
				[gc]::collect() 
				[gc]::WaitForPendingFinalizers()
         }
     }
     pause
 }
 until ($selection -eq 'q')
