#################################
### Andreys Survey AUTOMATION ###
#################################

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
					Write-Host 'Created Project name' $project_name 'Directory' -ForegroundColor Green
					New-Item -Path $project_name -Type Directory | out-null
				}	else {
				    Write-Host $project_name $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_ARCH' -PathType Container)) {
					Write-Host 'Created Sub Directory ARCH' -ForegroundColor Green
					New-Item -Path $project_name\$project_name'_ARCH' -Type Directory | out-null
				}	else {
				    Write-Host $project_name\$project_name'_ARCH' $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_ARCH\'$client_firm_name'_'$project_name'_ARCH_'$year'_01.vsdx' -PathType Leaf)) {
					Write-Host 'Created Visio empty file at Directory ARCH' -ForegroundColor Green
					New-Item -Path $project_name\$project_name'_ARCH\'$client_firm_name'_'$project_name'_ARCH_'$year'_01.vsdx' -ItemType File | out-null
				}	else {
				    Write-Host $project_name\$project_name'_ARCH\'$client_firm_name'_'$project_name'_ARCH_'$year'_01.vsdx' $file_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $project_name\$client_firm_name'_'$project_name'_ARCH_'$year'_01.docx' -PathType Leaf)) {
					Write-Host 'Created Template Report in ' $project_name 'Directory' -ForegroundColor Green
					Copy-Item $templates_path'ARCH.docx' $project_name\$client_firm_name'_'$project_name'_ARCH_'$year'_01.docx' | out-null
				}	else {
				    Write-Host $project_name\$client_firm_name'_'$project_name'_ARCH_'$year'_01.docx' $file_already_exists -ForegroundColor Magenta
				}
         }
		 
		 '2' {
			 if ( -not (Test-Path -Path $project_name -PathType Container)) {
					New-Item -Path $project_name -Type Directory | out-null
				}	else {
				    Write-Host $project_name $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_FW' -PathType Container)) {
					New-Item -Path $project_name\$project_name'_FW' -Type Directory | out-null
				}	else {
				    Write-Host $project_name\$project_name'_FW' $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_FW\Evidence' -PathType Container)) {
					New-Item -Path $project_name\$project_name'_FW\Evidence' -Type Directory | out-null
				}	else {
				    Write-Host $project_name\$project_name'_FW\Evidence' $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $project_name\'Servers_INFO.txt' -PathType Leaf)) {
					New-Item -Path $project_name\'Servers_INFO.txt' -ItemType File | out-null
				}	else {
				    Write-Host $project_name\'Servers_INFO.txt' $file_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $project_name\$client_firm_name'_'$project_name'_FW_'$year'_01.docx' -PathType Leaf)) {
					Copy-Item $templates_path'FW.docx' $project_name\$client_firm_name'_'$project_name'_FW_'$year'_01.docx' | out-null
				}	else {
				    Write-Host $project_name\$client_firm_name'_'$project_name'_FW_'$year'_01.docx' $file_already_exists -ForegroundColor Magenta
				}
         }
		 
		 '3' {
			 if ( -not (Test-Path -Path $project_name -PathType Container)) {
					New-Item -Path $project_name -Type Directory | out-null
				}	else {
				    Write-Host $project_name $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_HRD_WIN' -PathType Container)) {
					New-Item -Path $project_name\$project_name'_HRD_WIN' -Type Directory | out-null
				}	else {
				    Write-Host $project_name\$project_name'_HRD_WIN' $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_HRD_WIN\Evidence' -PathType Container)) {
					New-Item -Path $project_name\$project_name'_HRD_WIN\Evidence' -Type Directory | out-null
				}	else {
				    Write-Host $project_name\$project_name'_HRD_WIN\Evidence' $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_HRD_WIN\Recieved_Script_Outputs' -PathType Container)) {
					New-Item -Path $project_name\$project_name'_HRD_WIN\Recieved_Script_Outputs' -Type Directory | out-null
				}	else {
				    Write-Host $project_name\$project_name'_HRD_WIN\Recieved_Script_Outputs' $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $project_name\$client_firm_name'_'$project_name'_HRD_WIN_'$year'_01.docx' -PathType Leaf)) {
					Copy-Item $templates_path'HRD_WIN.docx' $project_name\$client_firm_name'_'$project_name'_HRD_WIN_'$year'_01.docx' | out-null
				}	else {
				    Write-Host $project_name\$client_firm_name'_'$project_name'_HRD_WIN_'$year'_01.docx' $file_already_exists -ForegroundColor Magenta
				}
		 }
		 '4' {
			 if ( -not (Test-Path -Path $project_name -PathType Container)) {
					New-Item -Path $project_name -Type Directory | out-null
				}	else {
				    Write-Host $project_name $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_HRD_RHEL' -PathType Container)) {
					New-Item -Path $project_name\$project_name'_HRD_RHEL' -Type Directory | out-null
				}	else {
				    Write-Host $project_name\$project_name'_HRD_RHEL' $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_HRD_RHEL\Evidence' -PathType Container)) {
					New-Item -Path $project_name\$project_name'_HRD_RHEL\Evidence' -Type Directory | out-null
				}	else {
				    Write-Host $project_name\$project_name'_HRD_RHEL\Evidence' $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $project_name\$project_name'_HRD_RHEL\Recieved_Script_Outputs' -PathType Container)) {
					New-Item -Path $project_name\$project_name'_HRD_RHEL\Recieved_Script_Outputs' -Type Directory | out-null
				}	else {
				    Write-Host $project_name\$project_name'_HRD_RHEL\Recieved_Script_Outputs' $folder_already_exists -ForegroundColor Magenta
				}
			 if ( -not (Test-Path -Path $project_name\$client_firm_name'_'$project_name'_HRD_RHEL_'$year'_01.docx' -PathType Leaf)) {
					Copy-Item $templates_path'HRD_RHEL.docx' $project_name\$client_firm_name'_'$project_name'_HRD_RHEL_'$year'_01.docx' | out-null
				}	else {
				    Write-Host $project_name\$client_firm_name'_'$project_name'_HRD_RHEL_'$year'_01.docx' $file_already_exists -ForegroundColor Magenta
				}
         }		 
		 'U' {
				$date_today = read-host "Enter Date you want to be displayed in the template report"
				$consultant_name = read-host "Enter a Consultant Name you want to be displayed in the template report"
				$system_manager_name = read-host "Enter a Systems Manager Name you want to be displayed in the template report"
				$templates_copied_to_dst_path = $project_name # multi-folders: "C:\fso1*", "C:\fso2*"
				$fileType = "*.docx"           # *.docx will take all .doc* files

				$textToReplace = @{
				# "TextToFind" = "TextToReplaceWith"
				"DATE" = $date_today
				"CONSULTANT_NAME" = $consultant_name
				"PROJECT_NAME" = $project_name
				"YEAR" = $year
				"SYSTEM_MANAGER" = $system_manager_name
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
				Get-ChildItem -Path $templates_copied_to_dst_path -Recurse -Filter $fileType | ForEach-Object { 
				  Write-Host "Processing \`"$($_.Name)\`"..."
				  $countr = processDoc
				  Write-Host "$countr replacements made."
				  $counta += $countr
				  $countf++
				}
				$sw.Stop()
				$elapsed = $sw.Elapsed.toString()
				Write-Host "`nDone. $countf files processed in $elapsed"
				Write-Host "$counta replacements made in total."

				$word.Quit()
				$word = $null
				[gc]::collect() 
				[gc]::WaitForPendingFinalizers()
         }
     }
     pause
 }
 until ($selection -eq 'q')
