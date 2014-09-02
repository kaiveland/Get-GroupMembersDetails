#--------------------------------
# Input parameters block
#--------------------------------
Param
(
	[Parameter(ParameterSetName="Identity",Mandatory=$true,HelpMessage="You must specify Active Directory group")]
		[ValidateNotNull()]
        [string]$Identity,
		
		[switch]$ConvertToExcel
)
#--------------------------------

#----------------------
#
#   Functions section
#
#----------------------

Function Convert-CsvToXls 
{
	param ($csv_filepath, $xls_filepath)
	
	$csv_entries = Import-Csv -Path $csv_filepath
	$num_of_lines = $csv_entries.Count
	$xl = New-Object -comobject Excel.Application
	$xl.DisplayAlerts = $False
	# Create a workbook
	$wb = $xl.Workbooks.open($xls_filepath)
	$num_sheets = $xl.sheetsInNewWorkbook
}

Function Get-ADGroupMembersType 
{
	param ($groupName)
	
	$groupMembers = Get-ADGroupMember -Identity $groupName
	$i = 1
	foreach ($groupMember in $groupMembers)
	{
		switch ($groupMember.objectClass)
			{
				'user' { $Ans[$i] = 1 }
				'group' { $Ans[$i] = 2 }
				default { $Ans[$i] = 0 }
			}
		$i++
	}
}

#-----END-OF-SECTION---

#--------------------------------
# Welcome info message
#--------------------------------
Write-Host "Use -ConvertToExcel switch to convert CSV files to one XLSX file." -foregroundcolor cyan -backgroundcolor black
#--------------------------------

#--------------------------------
# Initiate variables
#--------------------------------
$work_base_dir = $PSScriptRoot
$work_dir = $work_base_dir + '\work_folder'
$collection_path = $work_dir+'\*'
#--------------------------------

#--------------------------------
# Test working directory existence
#--------------------------------
if (!(Test-Path $work_dir)) { New-Item -ItemType directory -Path $work_dir }
#--------------------------------

#--------------------------------
# Delete all previous data
#--------------------------------
Remove-Item -Recurse -Force $collection_path
#--------------------------------

$source_group = $Identity
Write-Verbose -Message "Getting subordinate groups..."
$GroupMembers = Get-ADGroupMember -Identity $source_group
if ($GroupMembers -eq $Null)
{
	Write-Host "No such group in Active Directory"
	Exit
}
else
{ 
	Write-Verbose -Message "[done]"
}

foreach ($group_member in $GroupMembers)
{
	if ($group_member.objectClass -eq 'group')
	{
		$name_group_member = $group_member.Name
		Write-Verbose -Message "Processing $name_group_member"
		$Members = Get-ADGroupMember -Identity $group_member -Recursive
		$Properties = @('Name','SamAccountName','Company','Department','Title','Mobile','DisplayName','mail','msDS-cloudExtensionAttribute10')
		$filename = $group_member.Name+'_members.csv'
		$pathToCsv = $work_dir+'\'+$filename
	 	$temporary_fileName = '\z_'+$pid+'_processing.csv'
		$pathToTemp = $work_dir + $temporary_fileName
		foreach ($member in $Members)
		{
			Get-ADUser -Identity $member -Properties $Properties | select $Properties | ConvertTo-Csv | select -Skip 2 | Out-File $pathToTemp
			[System.IO.File]::ReadAllText($pathToTemp) | Out-File $pathToCsv -Append
		}
	
		if (Test-Path $pathToCsv)
		{
			$content = Get-Content -Path $pathToCsv
			$content -notmatch '(^[\s]*$)' | Set-Content -Path $pathToCsv
			$header ='"Name","SamAccountName","Company","Department","Title","Mobile","DisplayName","Mail","cloudExtensionAttribute10"'
			$final_content = Get-Content -Path $pathToCsv
			Set-Content -Path $pathToCsv -value $header, $final_content
			(Get-Content -Path $pathToCsv) | Set-Content -Encoding UTF8 -Path $pathToCsv # convert to utf8
			Write-Verbose -Message "[done]" #
		}
		else { Write-Verbose -Message "[skipped]" }  # skip empty groups
	}
	else { Write-Host $group_member.Name "was skipped, not a group" -foregroundcolor yellow -backgroundcolor black }
}

if ($pathToTemp -eq $Null) { Exit }
if (Test-Path $pathToTemp)
{
	Remove-Item $pathToTemp 
	Write-Host "CSV files was prepared in working folder."
}

if (!($ConvertToExcel)) { Exit } 

$collection = Get-ChildItem $collection_path -include *.csv
$csv_count=$collection.Count
Write-Verbose -Message "Detected the following CSV files: ($csv_count)"
$temp_filename = '\'+$Identity+'.xlsx'
$outputfilename = $work_dir + $temp_filename
$excelapp = new-object -comobject excel.application
$excelapp.DisplayAlerts = $False 
$excelapp.sheetsInNewWorkbook = $csv_count
#$excelapp.Visible = $true 					# if you need to control visually 
$xlsx = $excelapp.Workbooks.Add()
$sheet=1

foreach ($item in $collection)
{
	$currentFileName = $item.Name
	Write-Verbose -Message "Converting $currentFileName"
	$currentCsv = $work_dir + '\' + $currentFileName
	$processes = Import-Csv -Path $currentCsv
	$num_item = $processes.Count
	$worksheet = $xlsx.Worksheets.Item($sheet)
	$worksheet.Name = $currentFileName
	$worksheet.cells.item(1,1) = "Name" 
	$worksheet.cells.item(1,2) = "SamAccountName" 
	$worksheet.cells.item(1,3) = "Company" 
	$worksheet.cells.item(1,4) = "Department" 
	$worksheet.cells.item(1,5) = "Title"
	$worksheet.cells.item(1,6) = "Mobile" 
	$worksheet.cells.item(1,7) = "DisplayName" 
	$worksheet.cells.item(1,8) = "Mail"
	$worksheet.cells.item(1,9) = "cloudExtensionAttribute10"
	$i = 2
	foreach($process in $processes)
	{
		$j = $i - 1
		$worksheet.cells.item($i,1) = $process.Name 
		$worksheet.cells.item($i,2) = $process.SamAccountName
		$worksheet.cells.item($i,3) = $process.Company 
		$worksheet.cells.item($i,4) = $process.Department 
		$worksheet.cells.item($i,5) = $process.Title
		$worksheet.cells.item($i,6) = $process.Mobile 
		$worksheet.cells.item($i,7) = $process.DisplayName 
		$worksheet.cells.item($i,8) = $process.Mail
		$worksheet.cells.item($i,9) = $process.cloudExtensionAttribute10
		Write-Progress -Activity "Converting file $currentFileName" -status "processing line $j of $num_item" -percentComplete ($j / $num_item * 100)
		$i++
	}
	$sheet++
	Write-Verbose -Message "[done]"
}
$xlsx.SaveAs($outputfilename)
$excelapp.quit()
Write-Host "Done!" -foregroundcolor green -backgroundcolor black