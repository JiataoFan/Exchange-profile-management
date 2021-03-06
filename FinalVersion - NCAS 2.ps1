#"!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
#"!!!!!!!  THIS IS NOT A MICROSOFT SUPPORTED SCRIPT.  !!!!!!!!"
#"!!!!!!!      TEST IN A LAB FOR DESIRED OUTCOME      !!!!!!!!"
#"!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
#
# 11/08/2009	    Removed steps 2, 3, 4, 5, 6, 12, 13 and 14 (Simon)
# 11/23/2009	Added steps 8a and 8b (Simon)
# 11/25/2009	Added functions ToProperCase and CreateMailbox (Simon)
# 11/26/2009	Added functions ShowForm1 and ShowForm2 (Simon)
# 11/27/2009	Added fucntion Validate (Simon)
# 12/01/2009	Added ForeColor property to message boxes (Simon)
# 12/10/2009	Added Office ComboBox to forms (Simon)
# 12/14/2009	Added Phone/Extension TextBox to forms (Simon)
# 01/05/2010	Changed newbiiz to superbiiz (Simon)
# 01/09/2010	Checked $UserPhone and $UserExtension (Simon)
# 01/17/2010	Added functions ShowForm (Simon)
# 01/18/2010	Added Password and E-mail labels  (Simon)
# 02/29/2010	Updated Phone to nnn-nnn-nnnn format and validation of User Logon Name (Simon)
# 02/25/2010	Added $Email label and textbox (Simon)
# 02/26/2010	Made E-mail and Office optional (Simon)
# 02/29/2010	Added Mailbox Features checkboxes (Simon)
# 03/13/2010	Checked $UserAlias in function Validate (Simon)
# 03/19/2010	Added OutlookAnywhere checkbox (Simon)
# 03/20/2010	Added alternate smtp email address for malabs.com (Simon)
# 04/03/2010	Added function IsMember (Simon)
# 04/03/2010	Added logging (Simon)
# 04/06/2010	Checked for existing email address in Mail Contact (Simon)
# 04/05/2010	Added distribution group (Simon)
# 05/06/2010	Added group boxes to Form (Simon)
# 06/29/2010	Checked for existing name in Disabled OU (Simon)
# 07/23/2010	Added code to create distribution group for existing group (Simon)
# 07/24/2010	Checked user with no first or last name (Simon)
# 08/06/2010	Added e-mail notification for CreateMailbox (Simon)
# 09/03/2010	Added e-mail notification for DisableMailbox (Simon)
# 09/04/2010	Added e-mail notification for Create-Group (Simon)
# 09/09/2010	Added distribution group to CreateMailbox (Simon)
# 10/26/2010	Converted e-mail notification to HTML format (Simon)
# 11/18/2010	Added mail contact to the menu (Simon)
#

# Quick Convert Tip from malabs to malabs version
# 1.Replace all the malabs by malabs
# 2.Change the $DC to "dc1.ma.local"
# 3.1/10 Change the database and password (database: Delete mailserver Password: Convert the security back)
# 4.3/10 Change the company offline address book to default offline address book
# 5.8/10 Change the Address List
# 6.9/10 Update - Offline Address Book to the default one
# 7.10/10 Change the user property to match right case (LDAP CN pramater)
# 8.Change the default company from "Ma Labs" to "malabs"
# 9.Change the DC to "MA Labs" during the step 10  


<##############################
	Library files
################################>

. ".\launch form.ps1"
. ".\mailbox form.ps1"
. ".\UI config.ps1"
. ".\tree view.ps1"


[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][reflection.assembly]::LoadWithPartialName("'Microsoft.VisualBasic")


<##############################
	Configurations
################################>

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {

	Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit

}

$ErrorActionPreference = "stop"
#$ErrorActionPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"

$logarchive = "\\MLSQL\F$\Microsoft\Exchange\Logs"
# $logpath = "\\NETAPP1\MIS$\IT\Scripts\Logs\Exchange\"
# $logpath = "\\VASTO\IT$\Scripts\Logs\Exchange\"
$logpath = "C:\UserLog\"
$Forest = "malabs.com"
$NTDomain = "malabs"
#$DC = (${env:logonserver}).Substring(2)
$DC = "dc1.ma.local"
$title = "Exchange Manager"
$title_newdg = "Create a new distribution group"
$title_newuser = "Create a new mailbox"
$title_newcontact = "Create a new mail contact"
$title_disableuser = "Disable user"
$title_existinguser = "Create mailbox for an existing User"
$title_existingdg = "Create distribution group for an existing group"

$emailFrom = "Administrator@malabs.com"
$emailTo = "IT@malabs.com"
$emailToHR = "sjhr@malabs.com"
$emailToRuddy = "eric.xu@malabs.com"
$emailCc = ""
$smtpServer = "smtp.malabs.com"
$smtp = New-Object Net.Mail.SmtpClient ($smtpServer)

$context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext ([System.DirectoryServices.ActiveDirectory.DirectoryContextType]::DirectoryServer,$DC)
$domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($context)
$root = $domain.psbase.GetDirectoryEntry()

function ToProperCase ([string]$name) {

	(Get-Culture).TextInfo.ToTitleCase($name)

}

function IsMember ([string]$company, [string]$sam) {

	if (Get-DistributionGroupMember $company | Where-Object { $_.SamAccountName -eq $sam }) { $True } else { $False }

}

function ValidateEmailAddress {

	param([string]$email)
	[Microsoft.Exchange.Data.SmtpAddress]::IsValidSmtpAddress($email)

}

function Autofill {

	if ($new -and $firstNameLastNameTextBox.Text -ne "" -and $organizationalUnitComboBox.Text -ne "" -and $companyComboBox -ne "" -and $officeComboBox -ne "") {

		if (Validate-Fullname) {

			$userLogonNameTextBox.Enabled = $True
			#$domainComboBox.Enabled = $True
			$passwordTextBox.Enabled = $True
			$serverComboBox.Enabled = $True


			$databaseCheckedList.Enabled = $True
			$emailTextBox.Enabled = $True
			#$emailDomainComboBox.Enabled = $True
			$distributionGroupCheckedListBox.Enabled = $True
			#$phoneTextBox.Enabled = $True
			$extensionTextBox.Enabled = $True
			$OWACheckBox.Enabled = $True
			$activeSyncCheckBox.Enabled = $True
			$DFSCheckedListBox.Enabled = $True
			$fileShareCheckedListBox.Enabled = $True
			$websenseDepartmentCheckedListBox.Enabled = $True
			$websenseUserDefineCheckedListBox.Enabled = $True

			$UserName.Text = $UserFirst.substring(0, 1).ToUpper() + $UserFirst.substring(1).ToLower() + " " + $UserLast.substring(0,1).ToUpper() + $UserLast.substring(1).ToLower()

			$distributionGroupCheckedListBox.Items.Clear()
			$DFSCheckedListBox.Items.Clear()
			$fileShareCheckedListBox.Items.Clear()
			$websenseDepartmentCheckedListBox.Items.Clear()
			$websenseUserDefineCheckedListBox.Items.Clear()

			$MainCompanyName = $companyComboBox.SelectedItem.ToString()
			$Officename = $officeComboBox.SelectedItem.ToString()
			$DepartmentName = $departmentComboBox.SelectedItem.ToString()

			if ($PMGroupComboBox.visible -eq $True) {

				if ($PMGroupComboBox.SelectedItem -eq $null) {

					$PmgroupName = ''

				} else {

					$PmgroupName = $PMGroupComboBox.SelectedItem.ToString()

				}

			}

			#The group filtering of NDFS server
			$DFShash = @{

				#Hash-table for MA Labs
				"MA Labs/San Jose/ACCT" = "ACCT-HQ";
				"MA Labs/San Jose/AP" = "AP-HQ";
				"MA Labs/San Jose/AR" = "AR-HQ";
				"MA Labs/San Jose/Credit" = "Credit-HQ";
				"MA Labs/San Jose/Marketing" = "Marketing-HQ";
				"MA Labs/San Jose/Shipping" = "Shipping";
				"MA Labs/San Jose/MIS" = "MIS-HQ";
				"MA Labs/Wuhan/AP" = "AP-WH";
				"MA Labs/Wuhan/AR" = "AR-WH";
				"MA Labs/Wuhan/Credit" = "Credit-WH";
				"MA Labs/Wuhan/Inventory" = "Inv-WH";
				"MA Labs/Wuhan/Marketing" = "Marketing-WH";

				"MA Labs/San Jose/Purchasing/HDD" = "HDD-HQ";
				"MA Labs/San Jose/Purchasing/Memory" = "Memory-HQ";
				"MA Labs/San Jose/Purchasing/Microsoft" = "Microsoft-HQ";
				"MA Labs/San Jose/Purchasing/Monitor" = "Monitor-HQ";
				"MA Labs/San Jose/Purchasing/Motherboard" = "Motherboard-HQ";
				"MA Labs/San Jose/Purchasing/Networking" = "Networking-HQ";
				"MA Labs/San Jose/Purchasing/Notebook" = "Notebook-HQ";
				"MA Labs/San Jose/Purchasing/VGA" = "VGA-HQ";

				"MA Labs/Wuhan/PM/HDD" = "HDD-WH";
				"MA Labs/Wuhan/PM/Microsoft" = "Microsoft-WH";
				"MA Labs/Wuhan/PM/Monitor" = "Monitor-WH";
				"MA Labs/Wuhan/PM/Motherboard" = "Motherboard-WH";
				"MA Labs/Wuhan/PM/Networking" = "Networking-WH";
				"MA Labs/Wuhan/PM/Notebook" = "Notebook-WH";

				#Hash-table for Superbiiz
				"Superbiiz/San Jose/Accounting" = "SBZ-Accounting";
				"Superbiiz/Wuhan/Accounting" = "SBZ-Accounting-WH";

				#Hash-table for Supertalent
				"Supertalent/Wuhan/Sales" = "STT-Sales-WH";

			}

			$searcher_G = New-Object system.directoryservices.directorysearcher;
			$grp_G = New-Object system.directoryservices.directoryentry;
			$PathDirection1 = "OU = DFS, OU = Groups, DC = malabs, DC = com"
			$root_G = [adsi]("LDAP://" + $PathDirection1)
			Write-Host $root_G
			$searcher_G.SearchRoot = $root_G
			$searcher_G.SearchScope = "Onelevel"

			if ($PmgroupName -eq $null) {

				$Keyword = $DFShash.Item("$MainCompanyName/$Officename/$DepartmentName")

			} else {

				$Keyword = $DFShash.Item("$MainCompanyName/$Officename/$DepartmentName/$PmgroupName")

			}

			$searcher_G.FindAll() | ForEach-Object {

				if ($Keyword -eq $null) {

				} else {

					$SecurityGroupName = $_.properties.Name

					if ($_.properties.info -like "*$Keyword*") {

						[void]$DFSCheckedListBox.Items.Add("$SecurityGroupName")

					}

					for ($i = 0; $i -lt $DFSCheckedListBox.Items.Count; $i++) {

						$DFSCheckedListBox.SetItemChecked($i,$true)

					}

				}

			}

			#The group filtering of File Share server
			$FShash = @{

				#Hash-Table of Malabs
				"MA Labs/San Jose/ACCT" = "ACCT-HQ";
				"MA Labs/San Jose/AP" = "AP-HQ";
				"MA Labs/San Jose/AR" = "AR-HQ";
				"MA Labs/San Jose/Credit" = "Credit-HQ";
				"MA Labs/San Jose/HR" = "HR";
				"MA Labs/San Jose/Payroll" = "Payroll";
				"MA Labs/San Jose/Marketing" = "Marketing";

				#Hash-Table of Supertalent
				"Supertalent/San Jose/Accounting" = "STT-Accounting";
				"Supertalent/San Jose/Engineering" = "STT-Engineer";
				"Supertalent/San Jose/Sales" = "STT-Sales";

			}

			#
			$searcher_G = New-Object system.directoryservices.directorysearcher;
			$grp_G = New-Object system.directoryservices.directoryentry;
			$PathDirection1 = "OU=File Share,OU=Groups,DC=malabs,DC=com"
			$root_G = [adsi]("LDAP://" + $PathDirection1)
			$searcher_G.SearchRoot = $root_G
			$searcher_G.SearchScope = "Onelevel"

			#Chose the patten from the hash table
			$Keyword = $FShash.Item("$MainCompanyName/$Officename/$DepartmentName")

			#Achieve the name of the second level group
			$searcher_G.FindAll() | ForEach-Object {

				if ($Keyword -eq $null) {

				} else {

					$SecurityGroupName = $_.properties.Name
					if ($_.properties.info -like "*$Keyword*") {

						[void]$fileShareCheckedListBox.Items.Add("$SecurityGroupName")

					}

					for ($i = 0; $i -lt $fileShareCheckedListBox.Items.Count; $i++) {

						$fileShareCheckedListBox.SetItemChecked($i,$true)

					}

				}

			}

			#Add the websense department group
			$WebsenseDepthash = @{

				"MA Labs/Chicago/Accounting" = "Branch-ACCT";
				"MA Labs/Chicago/Sales" = "Sales-CHG";
				"MA Labs/Geogia/Accounting" = "Branch-ACCT";
				"MA Labs/Geogia/Sales" = "Sales-GA";
				"MA Labs/Geogia/Warehouse" = "Warehouse-GA";
				"MA Labs/Los Angles/Accounting" = "Branch-ACCT";
				"MA Labs/Los Angles/HR" = "HR-LA";
				"MA Labs/Los Angles/RMA" = "RMA-LA";
				"MA Labs/Los Angles/Sales" = "Sales-LA";
				"MA Labs/Los Angles/Warehouse" = "Warehouse-LA";
				"MA Labs/Miami/Accounting" = "Branch-ACCT";
				"MA Labs/Miami/Purchasing" = "Purchasing-MIA";
				"MA Labs/Miami/Sales" = "Sales-MIA";
				"MA Labs/Miami/Warehouse" = "Warehouse-MIA";
				"MA Labs/New Jersey/Accounting" = "Branch-ACCT";
				"MA Labs/New Jersey/Purchasing" = "Purchasing-NJ";
				"MA Labs/New Jersey/Sales" = "Sales-NJ";

				# Websense department for headquater
				"MA Labs/San Jose/AP" = "AP-HQ";
				"MA Labs/San Jose/AR" = "AR-HQ";
				"MA Labs/San Jose/Assembly" = "Assembly-HQ";
				"MA Labs/San Jose/Credit" = "Credit-HQ";
				"MA Labs/San Jose/Data Entry" = "Data Entry-HQ";
				"MA Labs/San Jose/HR" = "HR-HQ";
				"MA Labs/San Jose/Inventory" = "Inventory-HQ";
				"MA Labs/San Jose/IT" = "Information Tech-HQ";
				"MA Labs/San Jose/Marketing" = "Marketing-HQ";
				"MA Labs/San Jose/MIS" = "MIS-HQ";
				"MA Labs/San Jose/Purchasing" = "Purchasing-HQ";
				"MA Labs/San Jose/RMA" = "RMA-HQ";
				"MA Labs/San Jose/Sales" = "Sales-HQ";
				"MA Labs/San Jose/Shipping" = "Shipping-HQ";
				"MA Labs/San Jose/Tech Support" = "Tech Support-HQ";

				# Websense department for Superbiiz
				"Superbiiz/San Jose/Accounting" = "SBZ-ACCT";
				"Superbiiz/San Jose/Customer Service" = "SBZ-Customer Service";
				"Superbiiz/San Jose/Marketing" = "SBZ-Marketing";
				"Superbiiz/San Jose/Sales" = "SBZ-Sales";

				# Websense department for Supertalent
				"SuperTalent/San Jose/Accounting" = "STT-ACCT";
				"SuperTalent/San Jose/Engineering" = "STT-Engineering";
				"SuperTalent/San Jose/Marketing" = "STT-Marketing";
				"SuperTalent/San Jose/Sales" = "STT-Sales";
				"SuperTalent/San Jose/Tech Support" = "STT-Tech Support";

			}

			$searcher_G = New-Object system.directoryservices.directorysearcher
			$grp_G = New-Object system.directoryservices.directoryentry
			$PathDirection1 = "OU=Department,OU=Groups,DC=malabs,DC=com"
			$root_G = [adsi]("LDAP://" + $PathDirection1)
			$searcher_G.SearchRoot = $root_G
			$searcher_G.SearchScope = "Onelevel"
			#Chose the patten from the hash table
			$Keyword = $WebsenseDepthash.Item("$MainCompanyName/$Officename/$DepartmentName")
			#Achieve the name of the second level group
			$searcher_G.FindAll() | ForEach-Object {

				if ($Keyword -eq $null) {

				} else {

					$SecurityGroupName = $_.properties.Name

					if ($_.properties.info -like "*$Keyword*") {

						[void]$websenseDepartmentCheckedListBox.Items.Add("$SecurityGroupName")

					}

					for ($i = 0; $i -lt $websenseDepartmentCheckedListBox.Items.Count; $i++) {

						$websenseDepartmentCheckedListBox.SetItemChecked($i, $true)

					}

				}

			}

			# Add the websense list group
			$WebsenseListhash = @{

				# Websense List for superbiiz
				"Superbiiz/San Jose/Customer Service" = "SBZ-Customer Service";
				"Superbiiz/San Jose/Accounting" = "SBZ-ACCT";

				#Websense List for MA Labs
				"MA Labs/San Jose/Purchasing" = "Purchasing-HQ";
				"MA Labs/San Jose/RMA" = "RMA-HQ";
				"MA Labs/Los Angles/Purchasing" = "Purchasing-LA";
				"MA Labs/Los Angles/RMA" = "RMA-LA";

			}

			$searcher_G = New-Object system.directoryservices.directorysearcher;
			$grp_G = New-Object system.directoryservices.directoryentry;
			$PathDirection1 = "OU=Websense, OU=Groups, DC=malabs, DC=com"
			$root_G = [adsi]("LDAP://" + $PathDirection1)
			$searcher_G.SearchRoot = $root_G
			$searcher_G.SearchScope = "Onelevel"

			#Chose the patten from the hash table
			$Keyword = $WebsenseListhash.Item("$MainCompanyName/$Officename/$DepartmentName")

			#Achieve the name of the second level group
			$searcher_G.FindAll() | ForEach-Object {

				if ($Keyword -eq $null) {

				} else {

					$SecurityGroupName = $_.properties.Name
					if ($_.properties.info -like "*$Keyword*") {

						[void]$websenseUserDefineCheckedListBox.Items.Add("$SecurityGroupName")

					}

					for ($i = 0; $i -lt $websenseUserDefineCheckedListBox.Items.Count; $i++) {

						$websenseUserDefineCheckedListBox.SetItemChecked($i,$true)

					}

				}

			}

			$distributionGroupCheckedListBox.Items.Clear()

			Get-DistributionGroup -Filter "(CustomAttribute1 -eq '$MainCompanyName' -and CustomAttribute4 -eq '$Officename' -and CustomAttribute5 -like '*Office*')
			-or (CustomAttribute2 -eq '$MainCompanyName' -and CustomAttribute4 -eq '$Officename' -and CustomAttribute5 -like '*Office*')
			-or (CustomAttribute3 -eq '$MainCompanyName' -and CustomAttribute4 -eq '$Officename' -and CustomAttribute5 -like '*Office*')
			-or (CustomAttribute1 -eq '$MainCompanyName' -and CustomAttribute4 -like '*$Officename*' -and CustomAttribute5 -like '*<$DepartmentName>*')
			-or (CustomAttribute2 -eq '$MainCompanyName' -and CustomAttribute4 -like '*$Officename*' -and CustomAttribute5 -like '*<$DepartmentName>*')
			-or (CustomAttribute3 -eq '$MainCompanyName' -and CustomAttribute4 -like '*$Officename*' -and CustomAttribute5 -like '*<$DepartmentName>*')" | ForEach-Object {

				[void]$distributionGroupCheckedListBox.Items.Add($_.Name)

			}

			switch ($UserOU) {

				# Auto generated for Georgia
				{ ($_ -like "*Georgia*") -or ($_ -like "*GA/*") } {

					$Headletter = $UserLast.substring(0,1)
					$LogonName.Text = "g_" + $UserFirst.ToLower() + $Headletter
					$UserAlias = $LogonName.Text
					$Server.Text = "Raonic"
					$Keyword = "GA "

					$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"
					if ($tmpuser) {

						$Headletter = $UserFirst.substring(0,1)
						$LogonName.Clear()
						$LogonName.Text = "g_" + $Headletter + $UserLast.ToLower()
						$UserAlias = $LogonName.Text
						$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"

						if ($tmpuser) {

							$ErrorMsg.Text = "WARNING: Need type the logon name manually"
							$LogonName.Clear()

						}

					}

					$Email.Text = $UserFirst.ToLower() + "." + $UserLast.ToLower()
					$RandomNumber = Get-Random -Minimum 100 -Maximum 999
					$Password.Text = "test$RandomNumber"

				}

				# Auto generated for San Jose
				{ ($_ -like "*San Jose*") -or ($_ -like "*SJ/*") } {

					$Headletter = $UserLast.substring(0,1)
					$LogonName.Text = $UserFirst.ToLower() + $Headletter
					$UserAlias = $LogonName.Text
					$Keyword = "SJ "
					$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"

					switch ($UserOU) {

						# -or ($_ -like "*/AR*") -or ($_ -like "*/Credit*") -or ($_ -like "*/Sales*")
						{ ($_ -like "*/AP*") } {

							$Server.Text = "Sharapova"
							Set-Mailbox $UserAlias -CustomAttribute11 "ACC" -domaincontroller $DC

						}

						{ ($_ -like "*/AR*") } {

							$Server.Text = "Sharapova"
							Set-Mailbox $UserAlias -CustomAttribute11 "ACC" -domaincontroller $DC

						}

						{ ($_ -like "*/Credit*") } {

							$Server.Text = "Sharapova"
							Set-Mailbox $UserAlias -CustomAttribute11 "CRD" -domaincontroller $DC

						}

						{ ($_ -like "*/Sales*") } {

							$Server.Text = "Isner"
							Set-Mailbox $UserAlias -CustomAttribute11 "SAL" -domaincontroller $DC

						}

						#-or ($_ -like "*/Warehouse*") -or ($_ -like "*/RMA*")
						{ ($_ -like "*/Tech Support*") } {

							$Server.Text = "Wawrinka"
							Set-Mailbox $UserAlias -CustomAttribute11 "TECH" -domaincontroller $DC

						}

						{ ($_ -like "*/Warehouse*") } {

							$Server.Text = "Wawrinka"
							Set-Mailbox $UserAlias -CustomAttribute11 "WHS" -domaincontroller $DC

						}

						{ ($_ -like "*/RMA*") } {

							$Server.Text = "Wawrinka"
							Set-Mailbox $UserAlias -CustomAttribute11 "RMA" -domaincontroller $DC

						}

						{ ($_ -like "*/Shipping*") } {

							$Server.Text = "Wawrinka"

						}

						# -or ($_ -like "*/Marketing*") -or ($_ -like "*/MIS*") -or ($_ -like "*/MGMT*") -or ($_ -like "*/IT*") -or ($_ -like "*/Purchasing*")
						{ ($_ -like "*/HR*") } {

							$Server.Text = "Djokovic"
							Set-Mailbox $UserAlias -CustomAttribute11 "HR" -domaincontroller $DC

						}

						{ ($_ -like "*/Marketing*") } {

							$Server.Text = "Djokovic"
							Set-Mailbox $UserAlias -CustomAttribute11 "MKT" -domaincontroller $DC

						}

						{ ($_ -like "*/MIS*") } {

							$Server.Text = "Djokovic"
							Set-Mailbox $UserAlias -CustomAttribute11 "MIS" -domaincontroller $DC

						}

						{ ($_ -like "*/MGMT*") } {

							$Server.Text = "Orangutan"

						}

						{ ($_ -like "*/IT*") } {

							$Server.Text = "Djokovic"
							Set-Mailbox $UserAlias -CustomAttribute11 "IT" -domaincontroller $DC

						}

						{ ($_ -like "*/Purchasing*") } {

							$Server.Text = "Berdych"
							Set-Mailbox $UserAlias -CustomAttribute11 "PUR" -domaincontroller $DC

						}

						{ ($_ -like "*/ACCT*") } {

							$Server.Text = "Sharapova"
							Set-Mailbox $UserAlias -CustomAttribute11 "ACC" -domaincontroller $DC

						}

						{ ($_ -like "*/Payroll*") } {

							$Server.Text = "Orangutan"

						}

						{ ($_ -like "*/Data Entry*") } {

							$Server.Text = "Tarpaulin"
							Set-Mailbox $UserAlias -CustomAttribute11 "DEN" -domaincontroller $DC

						}

					}

					if ($tmpuser) {

						$Headletter = $UserFirst.substring(0,1)
						$LogonName.Clear()
						$LogonName.Text = $Headletter + $UserLast.ToLower()
						$UserAlias = $LogonName.Text
						$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"

						if ($tmpuser) {

							$ErrorMsg.Text = "WARNING: Need type the logon name manually"
							$LogonName.Clear()

						}

					}

					$Email.Text = $UserFirst.ToLower() + "." + $UserLast.ToLower()
					$RandomNumber = Get-Random -Minimum 100 -Maximum 999
					$Password.Text = "test$RandomNumber"

				}

				# Auto generated for Los Angles
				{ ($_ -like "*Los Angles*") -or ($_ -like "*LA/*") } {

					$Headletter = $UserLast.substring(0,1)
					$LogonName.Text = "c_" + $UserFirst.ToLower() + $Headletter
					$UserAlias = $LogonName.Text
					$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"
					$Server.Text = "Wawrinka"
					$Keyword = "LA "

					if ($tmpuser) {

						$Headletter = $UserFirst.substring(0,1)
						$LogonName.Clear()
						$LogonName.Text = "c_" + $Headletter + $UserLast.ToLower()
						$UserAlias = $LogonName.Text
						$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"

						if ($tmpuser) {

							$ErrorMsg.Text = "WARNING: Need type the logon name manually"
							$LogonName.Clear()

						}

					}

					$Email.Text = $UserFirst.ToLower() + "." + $UserLast.ToLower()
					$RandomNumber = Get-Random -Minimum 100 -Maximum 999
					$Password.Text = "test$RandomNumber"

				}

				# Auto generated for Chicago
				{ ($_ -like "*Chicago*") -or ($_ -like "*IL/*") } {

					$Headletter = $UserLast.substring(0,1)
					$LogonName.Text = "i_" + $UserFirst.ToLower() + $Headletter
					$UserAlias = $LogonName.Text
					$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"
					$Server.Text = "Raonic"
					$Keyword = "CH "

					if ($tmpuser) {

						$Headletter = $UserFirst.substring(0,1)
						$LogonName.Clear()
						$LogonName.Text = "i_" + $Headletter + $UserLast.ToLower()
						$UserAlias = $LogonName.Text
						$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"

						if ($tmpuser) {

							$ErrorMsg.Text = "WARNING: Need type the logon name manually"
							$LogonName.Clear()

						}

					}

					$Email.Text = $UserFirst.ToLower() + "." + $UserLast.ToLower()
					$RandomNumber = Get-Random -Minimum 100 -Maximum 999
					$Password.Text = "test$RandomNumber"

				}

				# Auto generated for New Jersey
				{ ($_ -like "*New Jersey*") -or ($_ -like "*NJ/*") } {

					$Headletter = $UserLast.substring(0,1)
					$LogonName.Text = "n_" + $UserFirst.ToLower() + $Headletter
					$UserAlias = $LogonName.Text
					$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"
					$Server.Text = "Raonic"
					$Keyword = "NJ "

					if ($tmpuser) {

						$Headletter = $UserFirst.substring(0,1)
						$LogonName.Clear()
						$LogonName.Text = "n_" + $Headletter + $UserLast.ToLower()
						$UserAlias = $LogonName.Text
						$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"

						if ($tmpuser) {

							$ErrorMsg.Text = "WARNING: Need type the logon name manually"
							$LogonName.Clear()

						}

					}

					$Email.Text = $UserFirst.ToLower() + "." + $UserLast.ToLower()
					$RandomNumber = Get-Random -Minimum 100 -Maximum 999
					$Password.Text = "test$RandomNumber"

				}

				# Auto generated for Miami
				{ $_ -like "*Miami*" -or ($_ -like "*MI/*") } {

					$Headletter = $UserLast.substring(0,1)
					$LogonName.Text = "m_" + $UserFirst.ToLower() + $Headletter
					$UserAlias = $LogonName.Text
					$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"
					$Server.Text = "Raonic"
					$Keyword = "MI "

					if ($tmpuser) {

						$Headletter = $UserFirst.substring(0,1)
						$LogonName.Clear()
						$LogonName.Text = "m_" + $Headletter + $UserLast.ToLower()
						$UserAlias = $LogonName.Text
						$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"

						if ($tmpuser) {

							$ErrorMsg.Text = "WARNING: Need type the logon name manually"
							$LogonName.Clear()

						}

					}

					$Email.Text = $UserFirst.ToLower() + "." + $UserLast.ToLower()
					$RandomNumber = Get-Random -Minimum 100 -Maximum 999
					$Password.Text = "test$RandomNumber"

				}

				# Auto generated for Ma Labs/China
				{ $_ -like "*MA Labs/China*" } {

					$Headletter = $UserLast.substring(0,1)
					$LogonName.Text = "w_" + $UserFirst.ToLower() + $Headletter
					$UserAlias = $LogonName.Text
					$Server.Text = "Sprintbok"
					$Keyword = "wh "

					$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"
					if ($tmpuser) {

						$Headletter = $UserFirst.substring(0,1)
						$LogonName.Clear()
						$LogonName.Text = "w_" + $Headletter + $UserLast.ToLower()
						$UserAlias = $LogonName.Text
						$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"

						if ($tmpuser) {

							$ErrorMsg.Text = "WARNING: Need type the logon name manually"
							$LogonName.Clear()

						}

					}

					$Email.Text = $UserFirst.ToLower() + "." + $UserLast.ToLower()
					$RandomNumber = Get-Random -Minimum 100 -Maximum 999
					$Password.Text = "test$RandomNumber"

				}

				# Auto generated for Superbiiz
				{ $_ -like "*Superbiiz*" } {

					$Headletter = $UserLast.substring(0,1)
					$LogonName.Text = $UserFirst.ToLower() + $Headletter
					$UserAlias = $LogonName.Text
					$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"
					$Server.Text = "Gasquet"
					$Keyword = "Superbiiz "

					if ($tmpuser) {

						$Headletter = $UserFirst.substring(0,1)
						$LogonName.Clear()
						$LogonName.Text = $Headletter + $UserLast.ToLower()
						$UserAlias = $LogonName.Text
						$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"

						if ($tmpuser) {

							$ErrorMsg.Text = "WARNING: Need type the logon name manually"
							$LogonName.Clear()

						}

					}

					$Email.Text = $UserFirst.ToLower() + "." + $UserLast.ToLower()
					$RandomNumber = Get-Random -Minimum 100 -Maximum 999
					$Password.Text = "test$RandomNumber"

				}

				# Auto generated for Superbiiz/China 
				{ $_ -like "*Superbiiz/China*" } {

					# Avoid the DG List repeat 
					$Headletter = $UserLast.substring(0,1)
					$LogonName.Text = "w_" + $UserFirst.ToLower() + $Headletter
					$UserAlias = $LogonName.Text
					$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"
					$Server.Text = "Gasquet"
					$Keyword = "Superbiiz "

					if ($tmpuser) {

						$Headletter = $UserFirst.substring(0,1)
						$LogonName.Clear()
						$LogonName.Text = "w_" + $Headletter + $UserLast.ToLower()
						$UserAlias = $LogonName.Text
						$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"

						if ($tmpuser) {

							$ErrorMsg.Text = "WARNING: Need type the logon name manually"
							$LogonName.Clear()

						}

					}

					$Email.Text = $UserFirst.ToLower() + "." + $UserLast.ToLower()
					$RandomNumber = Get-Random -Minimum 100 -Maximum 999
					$Password.Text = "test$RandomNumber"

				}

				# Auto generated for Supertalent
				{ $_ -like "*Supertalent*" } {

					$Headletter = $UserLast.substring(0,1)
					$LogonName.Text = $UserFirst.ToLower() + $Headletter
					$UserAlias = $LogonName.Text
					$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"
					$Server.Text = "Jaguar"
					$Keyword = "Supertalent "

					if ($tmpuser) {

						$Headletter = $UserFirst.substring(0, 1)
						$LogonName.Clear()
						$LogonName.Text = $Headletter + $UserLast.ToLower()
						$UserAlias = $LogonName.Text
						$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"

						if ($tmpuser) {

							$ErrorMsg.Text = "WARNING: Need type the logon name manually"
							$LogonName.Clear()

						}

					}

					$Email.Text = $UserFirst.ToLower() + "." + $UserLast.ToLower()
					$RandomNumber = Get-Random -Minimum 100 -Maximum 999
					$Password.Text = "test$RandomNumber"

				}

				# Auto generated for Supertalent/China
				{ $_ -like "*Supertalent/China*" } {

					$Headletter = $UserLast.substring(0, 1)
					$LogonName.Text = "w_" + $UserFirst.ToLower() + $Headletter
					$UserAlias = $LogonName.Text
					$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"
					$Server.Text = "Jaguar"
					$Keyword = "Supertalent "

					if ($tmpuser) {

						$Headletter = $UserFirst.substring(0, 1)
						$LogonName.Clear()
						$LogonName.Text = "w_" + $Headletter + $UserLast.ToLower()
						$UserAlias = $LogonName.Text
						$tmpuser = Get-User -Filter "SamAccountName -eq '$UserAlias'"

						if ($tmpuser) {

							$ErrorMsg.Text = "WARNING: Need type the logon name manually"
							$LogonName.Clear()

						}

					}

					$Email.Text = $UserFirst.ToLower() + "." + $UserLast.ToLower()
					$RandomNumber = Get-Random -Minimum 100 -Maximum 999
					$Password.Text = "test$RandomNumber"

				}

			}

			for ($i = 0; $i -lt $distributionGroupCheckedListBox.Items.Count; $i++) {

				$distributionGroupCheckedListBox.SetItemChecked($i, $true)

			}

		}

	}

}

<#
	To check if the fullname already exists in the organizational unit
#>
function Validate-Fullname {

	$UserFirst = ToProperCase ($UserName.Text.Split()[0].Trim())
	$UserLast = ToProperCase ($UserName.Text.Split()[1].Trim())
	$UserOU = $OU.SelectedItem

	if ($new) {

		$UserFull = $UserFirst + " " + $UserLast
		$tmpName = Get-User -Filter "Name -eq '$UserFull'" -OrganizationalUnit "$Forest$UserOU" -domaincontroller $DC

		if ($tmpName) {

			$ErrorMsg.Text = "WARNING: Name `"$UserFull`" already exists in OU."
			$FirstMsg.Text = "*"

			return $False

		}

	}

	return $True

}


function Validate {

	$Error = $False
	$OUMsg.Text = $null
	$FirstMsg.Text = $null
	$LogonMsg.Text = $null

	############ $contact?
	if (!$contact) {

		############ $PasswordMsg?
		$PasswordMsg.Text = $null

	}

	if ($contact) {

		$EmailMsg.Text = "*"

	} else {

		$EmailMsg.Text = $null

	}

	if (!$contact) {

		$ServerMsg.Text = $null

	}

	$PhoneMsg.Text = $null
	$ErrorMsg.Text = $null
	$tmpEmail = $null
	$tmpMailbox = $null
	$tmpName = $null
	$tmpUser = $null
	$tmpOU = $null

	##### $new ???
	if ($new -or $contact) {

		if ($UserAlias -match "^([a-zA-Z]([_]*[0-9a-zA-Z])*)$") {

			$LogonMsg.Text = $Null

		} else {

			$LogonMsg.Text = "*";
			$Error = $True

		}

	} else {

		if ($UserAlias) {

			$LogonMsg.Text = $Null

		} else {

			$LogonMsg.Text = "*";
			$Error = $True

		}

	}

	if ($new -or $contact) {

		if ($UserOU) {

			$OUMsg.Text = $Null

		} else {

			$OUMsg.Text = "*";
			$Error = $True

		}

		if (($UserFirst -match "^([a-zA-Z][0-9a-zA-Z ]*)$") -and ($UserLast -match "^([a-zA-Z][0-9a-zA-Z ]*)$")) {

			$FirstMsg.Text = $Null

		} else {

			$FirstMsg.Text = "*";
			$Error = $True

		}

		if ($UserDomain) {

			$LogonMsg.Text = $Null

		} else {

			$LogonMsg.Text = "*";
			$Error = $True

		}

	}

	if ($new) {

		if ($UserPassword -and $UserPassword.Length -ge 5) {

			$PasswordMsg.Text = $Null

		} else {

			$PasswordMsg.Text = "*";
			$Error = $True

		}

	}

	if ($new -or $found) {

		if ($UserEmail) {

			if (!($UserEmail -match "^([0-9a-zA-Z]([-.\w]*[0-9a-zA-Z])*)$")) {

				$EmailMsg.Text = "*";
				$Error = $True

			}

		}

	}

	if ($contact) {

		$Error = $True

		if ($UserEmail) {

			if (ValidateEmailAddress $UserEmail) {

				$EmailMsg.Text = $Null;
				$Error = $False

			}

		}

	}

	if ($new -or $contact -or $found) {

		if (!$disable -and !$contact) {

			if ($MailboxServer -and $Database) {

				$ServerMsg.Text = $Null

			} else {

				$ServerMsg.Text = "*";
				$Error = $True

			}

		}

		if ($UserPhone) {

			if (!(($UserPhone -match "^\d{3}-\d{3}-\d{4}$") -or ($UserPhone -match "^\d{3}-\d{3}-\d{3}-\d{4}$"))) {

				$PhoneMsg.Text = "*";
				$Error = $True

			}

		}

		if ($UserExtension) {

			if (!(($UserExtension -match "\d{3}") -or ($UserExtension -match "\d{4}"))) {

				$PhoneMsg.Text = "*";
				$Error = $True

			}

		}

	}

	if ($Error) {

		$ErrorMsg.Text = "WARNING: Incomplete or invalid data in field(s).";

		return $False

	}

	if ($contact) {

		$tmpContact = Get-MailContact -Filter "Alias -eq '$UserAlias'" -domaincontroller $DC

		if ($tmpContact) {

			$ErrorMsg.Text = "WARNING: Mail contact `"$UserAlias`" already exists."
			$LogonMsg.Text = "*"
			return $False

		}

		$UserFull = $UserFirst + " " + $UserLast
		$tmpName = Get-MailContact -Filter "DisplayName -eq '$UserFull'" -OrganizationalUnit "$Forest$UserOU" -domaincontroller $DC

		if ($tmpName) {

			$ErrorMsg.Text = "WARNING: Name `"$UserFull`" already exists in OU."
			$FirstMsg.Text = "*"
			return $False

		}

	} else {

		$tmpUser = Get-User -Filter "SamAccountName -eq '$UserAlias'" -domaincontroller $DC
		$tmpOU = ([string]$tmpUser.Identity.Parent)
		#$tmpOU = ([string]$tmpUser.Identity).Substring($Forest.Length, ($tmpUser.Identity.Length - $tmpUser.Name.Length-$Forest.Length -1))

		if ($new) {

			if ($tmpUser) {

				$ErrorMsg.Text = "WARNING: User `"$UserAlias`" already exists."
				$LogonMsg.Text = "*"
				return $False

			}

			$UserFull = $UserFirst + " " + $UserLast
			$tmpName = Get-User -Filter "Name -eq '$UserFull'" -OrganizationalUnit "$Forest$UserOU" -domaincontroller $DC
			if ($tmpName) {

				$ErrorMsg.Text = "WARNING: Name `"$UserFull`" already exists in OU."
				$FirstMsg.Text = "*"
				return $False

			}

		} else {

			if (!$tmpUser) {

				$ErrorMsg.Text = "WARNING: User `"$UserAlias`" does not exist."
				$LogonMsg.Text = "*"
				return $False

			}

			if (!$disable) {

				$tmpMailbox = Get-Mailbox -Filter "SAMAccountName -eq '$UserAlias'" -domaincontroller $DC

				if ($tmpMailbox) {

					$ErrorMsg.Text = "WARNING: Mailbox `"$UserAlias`" already exists."
					$LogonMsg.Text = "*"
					return $False

				}

			}

			if (!($tmpOU -match "Users$")) {

				$ErrorMsg.Text = "WARNING: User `"$UserAlias`" not in an Users OU."
				$OUMsg.Text = "*"
				return $False

			}

		}

	}

	if ($new -or $contact -or $found) {

		$emailaddr = $UserEmail + $UserEmailDomain
		$tmpEmail = Get-Mailbox -Filter "EmailAddresses -eq '$emailaddr'" -domaincontroller $DC

		if ($tmpEmail -and !$disable) {

			$ErrorMsg.Text = "WARNING: E-mail is already being used by mailbox `"" + $tmpEmail.Name + "`"."
			$EmailMsg.Text = "*"
			return $False

		}

		$tmpContact = Get-MailContact -Filter "EmailAddresses -eq '$emailaddr'" -domaincontroller $DC
		if ($tmpContact -and !$disable) {

			$ErrorMsg.Text = "WARNING: E-mail is already being used by mail contact `"" + $tmpContact.Name + "`"."
			$EmailMsg.Text = "*"
			return $False

		}

	}

	# If user already disabled
	if ($disable -and ($tmpUser.RecipientTypeDetails -eq "DisabledUser")) {

		$ErrorMsg.Text = "WARNING: User `"$UserAlias`" already disabled."
		$LogonMsg.Text = "*"
		return $False

	}

	# If name already exists in Disabled OU
	if ($tmpUser.Company) {

		$company = $tmpUser.Company

	} else {

		$company = "MA Labs"

	}

	$name = $tmpUser.Name
	# $tmpName = Get-User -Filter {DisplayName -eq $name} -OrganizationalUnit "$Forest/Disabled/$company/Users" -DomainController $DC
	if ($disable -and $tmpName) {

		$ErrorMsg.Text = "WARNING: Name `"$name`" already exists in /Disabled/$company/Users OU."
		$LogonMsg.Text = "*"
		return $False

	}

	return $True

}

function ValidateGroup {

	$Error = $False
	$OUMsg.Text = $null
	$NameMsg.Text = $null
	$LogonMsg.Text = $null
	$EmailMsg.Text = $null
	$ErrorMsg.Text = $null
	$tmpEmail = $null
	$tmpGroup = $null
	$tmpName = $null

	if ($GroupAlias -match "^([a-zA-Z][0-9a-zA-Z]*)$") {

		$LogonMsg.Text = $Null

	} else {

		$LogonMsg.Text = "*";
		$Error = $True

	}

	if ($newgroup) {

		if ($GroupOU) {

			$OUMsg.Text = $Null

		} else {

			$OUMsg.Text = "*"
			$Error = $True

		}

		if ($GroupName -match "^([a-zA-Z][0-9a-zA-Z ]*)$") {

			$NameMsg.Text = $Null

		} else {

			$NameMsg.Text = "*";
			$Error = $True

		}

		if ($GroupEmail) {

			if (!($GroupEmail -match "^([0-9a-zA-Z]([-.\w]*[0-9a-zA-Z])*)$")) {

				$EmailMsg.Text = "*";
				$Error = $True

			}

		}

	}

	if ($Error) {

		$ErrorMsg.Text = "WARNING: Incomplete or invalid data in field(s).";
		return $False

	}

	$tmpGroup = Get-Group -Filter { SamAccountName -EQ $GroupAlias } -domaincontroller $DC
	$tmpOU = ([string]$tmpGroup.Identity.Parent)

	if ($newgroup) {

		if ($tmpGroup) {

			$ErrorMsg.Text = "WARNING: Group `"$GroupAlias`" already exists."
			$LogonMsg.Text = "*"
			return $False

		}

		$tmpGroupName = Get-Group -Filter { DisplayName -EQ $GroupName } -OrganizationalUnit "$Forest$GroupOU" -domaincontroller $DC
		if ($tmpGroupName) {

			$ErrorMsg.Text = "WARNING: Name `"$GroupName`" already exists."
			$NameMsg.Text = "*"
			return $False

		}

	} else {

		if (!$tmpGroup) {

			$ErrorMsg.Text = "WARNING: Group `"$GroupAlias`" does not exist."
			$LogonMsg.Text = "*"
			return $False

		}

		$tmpDG = Get-DistributionGroup -Filter { SamAccountName -EQ $GroupAlias } -domaincontroller $DC
		if ($tmpDG) {

			$ErrorMsg.Text = "WARNING: Distribution group `"$GroupAlias`" already exists."
			$LogonMsg.Text = "*"
			return $False

		}

	}

	if ($newgroup -or $found) {

		$emailaddr = $GroupEmail + $GroupDomain
		$tmpEmail = Get-Group -Filter { WindowsEmailAddress -EQ $emailaddr } -domaincontroller $DC
		if ($tmpEmail -ne $null) {

			$ErrorMsg.Text = "WARNING: E-mail `"$emailaddr`" already exists."
			$EmailMsg.Text = "*"
			return $False

		}

	}

	return $True

}

function CreateMailbox {

	$log = $logpath + $logfile
	if (Test-Path $log) {

		$oldlog = $logpath + $UserAlias + ".old.log"
		Move-Item $log $oldlog -Force

	}

	switch ($UserDomain) {

		"@superbiiz.com" {

			$UserCompany = "Superbiiz"

		}

		"@supertalent.com" {

			$UserCompany = "Supertalent"

		}

		default {

			$UserCompany = "MA Labs"

		}

	}

	# Convert password to secured password
	if ($new) {

		$SecurePassword = New-Object System.Security.SecureString
		foreach ($char in $PassWord.Text.ToCharArray()) {

			$SecurePassword.AppendChar($char)

		}

	}

	# Hide buttons
	$backButton.Enabled = $False
	$cancelButton.Enabled = $False
	$OKButton.Enabled = $False

	# 1 of 10, creating mailbox
	$errorMessage1.Text = "[1/10] Creating mailbox `"$UserAlias`"...please wait."
	$Form2.Refresh()
	if ($new) {

		$result = New-Mailbox -Name "$UserFirst $UserLast" -Alias $UserAlias -OrganizationalUnit "$Forest$UserOU" -SamAccountName $UserAlias -UserPrincipalName "$UserAlias$UserDomain" -FirstName $UserFirst -LastName $UserLast -Password $SecurePassword -Database "$Database" -domaincontroller $DC

	} else {

		$result = Enable-Mailbox -Identity "$NTDomain\$UserAlias" -Database "$MailboxServer\$Database" -domaincontroller $DC

	}

	"1. Creating mailbox `"$UserAlias`"." > $log
	Start-Sleep -s 1

	if ($result -ne $null) {

		# 2 of 10, setting custom attribute
		switch ($UserDomain) {

			"@superbiiz.com" {

				$customattr = "CustomAttribute1"

			}

			"@supertalent.com" {

				$customattr = "CustomAttribute3"

			}

			default {

				$customattr = "CustomAttribute2"

			}

		}

		$errorMessage2.Text = "[2/10] Setting $customattr attribute to `"$UserCompany`"."
		$Form2.Refresh()
		switch ($UserDomain) {

			"@superbiiz.com" {

				$customattr = "CustomAttribute1";
				Set-Mailbox $userAlias -CustomAttribute1 "$userCompany" -UserPrincipalName "$UserAlias$UserDomain" -domaincontroller $DC -EmailAddressPolicyEnabled $False -EmailAddresses "SMTP:$UserEmail$UserEmailDomain"

			}

			"@supertalent.com" {

				$customattr = "CustomAttribute3";
				Set-Mailbox $userAlias -CustomAttribute3 "$userCompany" -UserPrincipalName "$UserAlias$UserDomain" -domaincontroller $DC -EmailAddressPolicyEnabled $False -EmailAddresses "SMTP:$UserEmail$UserEmailDomain"

			}

			default {

				$customattr = "CustomAttribute2";
				Set-Mailbox $userAlias -CustomAttribute2 "$userCompany" -UserPrincipalName "$UserAlias$UserDomain" -domaincontroller $DC -EmailAddressPolicyEnabled $False -EmailAddresses "SMTP:$UserEmail$UserEmailDomain"

			}

		}

		"" >> $log
		"2. Setting $customattr attribute to `"$UserCompany`"." >> $log
		Start-Sleep -s 1

		# 3 of 10, setting msExchUseOAB
		$errorMessage3.Text = "[3/10] Setting offline address book to `"$UserCompany OAB`"."
		$Form2.Refresh()
		Set-Mailbox $userAlias -offlineaddressbook "$userCompany OAB" -domaincontroller $DC
		"" >> $log
		"3. Setting offline address book to `"$UserCompany OAB`"." >> $log
		Start-Sleep -s 1

		# 4 of 10, disabling mailbox features
		$errorMessage4.Text = "[4/10] Disabling mailbox features."
		$Form2.Refresh()
		$BlockOutlookAnywhere = !$UserOutlookAnywhere
		Set-CASMailbox -Identity "$NTDomain\$UserAlias" -ActiveSyncEnabled $UserActiveSync -OWAEnabled $UserOWA -MAPIBlockOutlookRpcHttp $BlockOutlookAnywhere -domaincontroller $DC
		"" >> $log
		"4. Disabling mailbox features." >> $log
		Start-Sleep -s 1

		# 5 of 10, setting Organization and Phone
		$errorMessage5.Text = "[5/10] Setting Organization and Phone."
		$Form2.Refresh()
		Set-User $UserAlias -Company "$UserCompany" -Phone ("$UserPhone $UserExtension".Trim()) -Office "$UserOffice" -domaincontroller $DC
		"" >> $log
		"5. Setting Organization and Phone." >> $log
		Start-Sleep -s 1

		# 6 of 10, adding user to company
		if (IsMember "$UserCompany Company" $UserAlias) {

			$errorMessage6.ForeColor = "Red"
			$errorMessage6.Text = "[6/10] WARNING: User already a member of `"$UserCompany Company`"."
			"" >> $log
			"6. WARNING: User already a member of `"$UserCompany Company`"." >> $log

		} else {

			$errorMessage6.Text = "[6/10] Adding user to `"$UserCompany Company`"."
			Add-DistributionGroupMember "$UserCompany Company" -member $UserAlias -domaincontroller $DC -BypassSecurityGroupManagerCheck
			"" >> $log
			"6. Adding user to `"$UserCompany Company`"." >> $log

		}
		$Form2.Refresh()
		Start-Sleep -s 1

		# 7 of 10, adding user to distribution group

		if ($GroupName_DG) {

			$SuccessDG = @()

			for ($i = 0; $i -lt $GroupName_DG.Length; $i++) {

				$UserDG = $GroupName_DG[$i]
				if (IsMember $UserDG $UserAlias) {

					$errorMessage16.ForeColor = "Red"
					$errorMessage16.Text = "[7/10] WARNING: User already a member of distribution group `"$UserDG`"."
					"" >> $log
					"7. WARNING: User already a member of distribution group `"$UserDG`"." >> $log
					cotinue

				} else {

					$SuccessDG += $UserDG
					$errorMessage7.Text = "[7/10] Adding user to distribution group `"$SuccessDG`"."
					Add-DistributionGroupMember "$UserDG" -member $UserAlias -domaincontroller $DC -BypassSecurityGroupManagerCheck
					"" >> $log
					"7. Adding user to distribution group `"$UserDG`"." >> $log
					Start-Sleep -s 1

				}

			}

		} else {

			$errorMessage7.Text = "[7/10] Skipped...no distribution group specified."
			"" >> $log
			"7. Skipped...no distribution group specified." >> $log

		}

		$Form2.Refresh()
		Start-Sleep -s 1

		# 8 of 10, updating Address List
		$errorMessage8.ForeColor = "Green"
		$errorMessage8.Text = "[8/10] Updating address list `"$UserCompany Address List`"."
		$Form2.Refresh()
		Update-AddressList "$UserCompany Address List" -domaincontroller $DC
		"" >> $log
		"8. Updating address list `"$UserCompany Address List`"." >> $log
		Start-Sleep -s 1

		# 9 of 10, updating Offline Address Book
		$errorMessage9.Text = "[9/10] Updating offline address book `"$UserCompany OAB`"."
		$Form2.Refresh()
		Update-OfflineAddressBook "$UserCompany OAB" -domaincontroller $DC
		"" >> $log
		"9. Updating offline address book `"$UserCompany OAB`"." >> $log
		Start-Sleep -s 1

		# To Bind
		$tmpUser = Get-Mailbox $UserAlias -domaincontroller $DC
		Write-Host $tmpUser
		$user = ([adsi]("LDAP://" + $tmpUser.distinguishedname)).psbase
		Write-Host $user
		$distinguishedName = $tmpUser.distinguishedname
		Write-Host $distinguishedName
		Start-Sleep -s 1
		$comp = Get-AdUser -Filter * -SearchBase "$distinguishedName"
		#$user.ObjectSecurity.Owner = "MALABS\Domain Admins"
		# To Modifys:
		# 10 of 10, setting msExchQueryBaseDN (Uncheck the option, may not use in exchange 2013)
		#Change ownership of AD object
		Write-Host $comp
		$comppath = "AD:$($comp.DistinguishedName.ToString())"
		$acl = Get-Acl -Path $comppath
		$objUser = New-Object System.Security.Principal.NTAccount ("malabs","domain admins")
		$acl.SetOwner($objUser)
		Set-Acl -Path $comppath -AclObject $acl

		#Set mail MAPIBlockOutlookRpcHttp to false

		Set-CASMailbox $UserAlias -MAPIBlockOutlookRpcHttp $false

		#############################

		$errorMessage10.Text = "[10/10] Setting msExchQueryBaseDN attribute."
		$Form2.Refresh()
		#$user.MsExchSeachBase = "CN=$UserCompany Address List,CN=All Address Lists,CN=Address Lists Container,CN=MA Labs,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=MA Labs,DC=local"
		"" >> $log
		"10. Setting msExchQueryBaseDN attribute." >> $log
		Start-Sleep -s 1
		[void]$user.CommitChanges()
		$Pass = $Password.Text
		$emailToBranch = $EmailCheckLabel3.Text
		Write-Host $emailToBranch
		$OKButton.visible = $False
		$finishButton.TabIndex = 0
		$finishButton.visible = $True
		$errorMessage11.Text = "Mailbox `"$UserFirst $UserLast ($UserAlias)`" created sucessfully. Check log for errors."
		"" >> $log
		"Mailbox `"$UserFirst $UserLast ($UserAlias)`" created sucessfully." >> $log
		Start-Sleep -s 1

	} else {

		$errorMessage11.Text = "ERROR: Unable to create mailbox `"$UserFirst $UserLast ($UserAlias)`". Check the command syntax."
		"" >> $log
		"ERROR: Unable to create mailbox `"$UserFirst $UserLast ($UserAlias)`". Check the command syntax." >> $log
		Start-Sleep -s 1

	}

	#Copy-Item $log $logarchive
	if ($CheckBranch) {

		$subject = "Mailbox $UserFirst $UserLast ($UserAlias) Created"
		$body = "<font face=Arial size=2>THIS IS A SYSTEM GENERATED MESSAGE. PLEASE DO NOT REPLY."
		$body = $body + "<p>A new mailbox has been created for:"
		$body = $body + "<p>Name: $UserFirst $UserLast"
		$body = $body + "<br>User Logon Name: $UserAlias"
		$body = $body + "<br>Password: $Pass"
		$body = $body + "<br>E-mail: <a href=`"mailto:$UserEmail$UserEmailDomain`">$UserEmail$UserEmailDomain</a>"
		$body = $body + "<br>Phone: $UserPhone $UserExtension"
		$body = $body + "<p>Exchange Administrator<br>Ma Lab, Inc.</font>"
		$message = New-Object Net.Mail.MailMessage ($emailFrom,$emailToBranch,$subject,$body)
		$message.cc.Add("ray.jiang@malabs.com")
		$message.IsBodyHTML = $True
		$smtp.Send($message)

	}

	if ($CheckHR) {

		# The email for HR
		$subject = "Mailbox $UserFirst $UserLast ($UserAlias) Created"
		$body = "<font face=Arial size=2>THIS IS A SYSTEM GENERATED MESSAGE. PLEASE DO NOT REPLY."
		$body = $body + "<p>A new mailbox has been created for:"
		$body = $body + "<p>Name: $UserFirst $UserLast"
		$body = $body + "<br>User Logon Name: $UserAlias"
		$body = $body + "<br>E-mail: <a href=`"mailto:$UserEmail$UserEmailDomain`">$UserEmail$UserEmailDomain</a>"
		$body = $body + "<br>Distribution Group: $SuccessDG"
		$body = $body + "<br>Office: $UserOffice"
		$body = $body + "<br>Phone: $UserPhone $UserExtension"
		$body = $body + "<p>Exchange Administrator<br>Ma Lab, Inc.</font>"
		$message = New-Object Net.Mail.MailMessage ($emailFrom,$emailToHR,$subject,$body)
		$message.cc.Add("ray.jiang@malabs.com")
		$message.IsBodyHTML = $True
		$smtp.Send($message)

	}

	if ($CheckIT) {

		$subject = "Mailbox $UserFirst $UserLast ($UserAlias) Created"
		$body = "<font face=Arial size=2>THIS IS A SYSTEM GENERATED MESSAGE. PLEASE DO NOT REPLY."
		$body = $body + "<p>A new mailbox has been created for:"
		$body = $body + "<p>Name: $UserFirst $UserLast"
		$body = $body + "<br>User Logon Name: $UserAlias"
		$body = $body + "<br>Password: $Pass"
		$body = $body + "<br>E-mail: <a href=`"mailto:$UserEmail$UserEmailDomain`">$UserEmail$UserEmailDomain</a>"
		$body = $body + "<br>Organizational Unit: $UserOU"
		$body = $body + "<br>Server: $MailboxServer"
		$body = $body + "<br>Database: $Database"
		$body = $body + "<br>Distribution Group: $SuccessDG"
		$body = $body + "<br>Office: $UserOffice"
		$body = $body + "<br>Phone: $UserPhone $UserExtension"
		$body = $body + "<p>Exchange Administrator<br>Ma Lab, Inc.</font>"
		$message = New-Object Net.Mail.MailMessage ($emailFrom,$emailTo,$subject,$body)
		$message.IsBodyHTML = $True
		$smtp.Send($message)

	}

}


function CreateContact {

	$log = $logpath + $logfile

	if (Test-Path $log) {

		$oldlog = $logpath + $UserAlias + ".old.log"
		Move-Item $log $oldlog -Force

	}

	switch ($UserDomain) {

		"@superbiiz.com" {

			$UserCompany = "Superbiiz"

		}

		"@supertalent.com" {

			$UserCompany = "Supertalent"

		}

		default {

			$UserCompany = "malabs"

		}

	}

	# Hide buttons
	$BackButton.Enabled = $False
	$CancelButton.Enabled = $False
	$OKButton.Enabled = $False

	# 1 of 7, creating mail contact
	$ErrorMsg2.Text = "[1/7] Creating mail contact `"$UserAlias`"...please wait."
	$Form6.Refresh()
	$result = New-MailContact -Name "$UserFirst $UserLast" -Alias $UserAlias -ExternalEmailAddress $UserEmail -OrganizationalUnit "$Forest$UserOU" -FirstName $UserFirst -LastName $UserLast -domaincontroller $DC

	"1. Creating mail contact `"$UserAlias`"." > $log
	Start-Sleep -s 1

	if ($result -ne $null) {

		# 2 of 7, setting custom attribute
		switch ($UserDomain) {

			"@superbiiz.com" {

				$customattr = "CustomAttribute1"

			}

			"@supertalent.com" {

				$customattr = "CustomAttribute3"

			}

			default {

				$customattr = "CustomAttribute2"

			}

		}

		$ErrorMsg3.Text = "[2/7] Setting $customattr attribute to `"$UserCompany`"."
		$Form6.Refresh()
		switch ($UserDomain) {

			"@superbiiz.com" {

				$customattr = "CustomAttribute1"; 
				Set-MailContact "$userAlias" -CustomAttribute1 "$userCompany" -domaincontroller $DC -EmailAddressPolicyEnabled $False -EmailAddresses "SMTP:$UserEmail$UserEmailDomain"

			}

			"@supertalent.com" {

				$customattr = "CustomAttribute3";
				Set-MailContact "$userAlias" -CustomAttribute3 "$userCompany" -domaincontroller $DC -EmailAddressPolicyEnabled $False -EmailAddresses "SMTP:$UserEmail$UserEmailDomain"

			}

			default {

				$customattr = "CustomAttribute2";

				Set-MailContact "$userAlias" -CustomAttribute2 "$userCompany" -domaincontroller $DC -EmailAddressPolicyEnabled $False -EmailAddresses "SMTP:$UserEmail$UserEmailDomain"

			}

		}

		"" >> $log
		"2. Setting $customattr attribute to `"$UserCompany`"." >> $log
		Start-Sleep -s 1

		# 3 of 7, setting Organization and Phone
		$ErrorMsg4.Text = "[3/7] Setting Organization and Phone."
		$Form6.Refresh()
		Set-Contact $UserAlias -Company "$UserCompany" -Phone ("$UserPhone $UserExtension".Trim()) -Office "$UserOffice" -domaincontroller $DC
		"" >> $log
		"3. Setting Organization and Phone." >> $log
		Start-Sleep -s 1

		# 4 of 7, adding user to company
		if (IsMember "$UserCompany Company" $UserAlias) {

			$ErrorMsg5.ForeColor = "Red"
			$ErrorMsg5.Text = "[4/7] WARNING: User already a member of `"$UserCompany Company`"."
			"" >> $log
			"4. WARNING: User already a member of `"$UserCompany Company`"." >> $log

		} else {

			$ErrorMsg5.Text = "[4/7] Adding user to `"$UserCompany Company`"."
			Add-DistributionGroupMember -Identity "$UserCompany Company" -member $UserAlias -domaincontroller $DC
			"" >> $log
			"4. Adding user to `"$UserCompany Company`"." >> $log

		}
		$Form6.Refresh()
		Start-Sleep -s 1

		# 5 of 7, adding user to distribution group
		if ($UserDG) {

			$ErrorMsg6.Text = "[5/7] Adding user to distribution group `"$UserDG`"."
			Add-DistributionGroupMember -Identity "$UserDG" -member $UserAlias -domaincontroller $DC
			"" >> $log
			"5. Adding user to distribution group `"$UserDG`"." >> $log

		} else {

			$ErrorMsg6.Text = "[5/7] Skipped...no distribution group specified."
			"" >> $log
			"7. Skipped...no distribution group specified." >> $log

		}

		$Form6.Refresh()
		Start-Sleep -s 1

		# 6 of 7, updating Address List
		$errorMessage6.ForeColor = "Green"
		$errorMessage6.Text = "[6/7] Updating address list `"$UserCompany Address List`"."
		$Form6.Refresh()
		Update-AddressList "$UserCompany Address List" -domaincontroller $DC
		"" >> $log
		"6. Updating address list `"$UserCompany Address List`"." >> $log
		Start-Sleep -s 1

		# 7 of 7, updating Offline Address Book
		$ErrorMsg8.Text = "[7/7] Updating offline address book `"$UserCompany OAB`"."
		$Form6.Refresh()
		Update-OfflineAddressBook "$UserCompany OAB" -domaincontroller $DC
		"" >> $log
		"7. Updating offline address book `"$UserCompany OAB`"." >> $log
		Start-Sleep -s 1

		$OKButton.visible = $False
		$FinishButton.TabIndex = 0
		$FinishButton.visible = $True
		$ErrorMsg9.Text = "Mail contact `"$UserFirst $UserLast ($UserAlias)`" created sucessfully. Check log for errors."
		"" >> $log
		"Mail contact `"$UserFirst $UserLast ($UserAlias)`" created sucessfully." >> $log

	} else {

		$ErrorMsg9.Text = "ERROR: Unable to create mail contact `"$UserFirst $UserLast ($UserAlias)`". Check the command syntax."
		"" >> $log
		"ERROR: Unable to create mail contact `"$UserFirst $UserLast ($UserAlias)`". Check the command syntax." >> $log

	}

	#Copy-Item $log $logarchive
	$subject = "Mail Contact $UserFirst $UserLast ($UserAlias) Created"
	$body = "<font face=Arial size=2>THIS IS A SYSTEM GENERATED MESSAGE. PLEASE DO NOT REPLY."
	$body = $body + "<p>A new mail contact has been created for:"
	$body = $body + "<p>Name: $UserFirst $UserLast"
	$body = $body + "<br>User Logon Name: $UserAlias"
	$body = $body + "<br>E-mail: <a href=`"mailto:$UserEmail$UserEmailDomain`">$UserEmail$UserEmailDomain</a>"
	$body = $body + "<br>Organizational Unit: $UserOU"
	$body = $body + "<br>Distribution Group: $UserDG"
	$body = $body + "<br>Office: $UserOffice"
	$body = $body + "<br>Phone: $UserPhone $UserExtension"
	$body = $body + "<p>Exchange Administrator<br>Ma Lab, Inc.</font>"
	$message = New-Object Net.Mail.MailMessage ($emailFrom,$emailTo,$subject,$body)
	$message.IsBodyHTML = $True
	$smtp.Send($message)

}

function DisableUser {

	$log = $logpath + $logfile

	if ($tmpUser.RecipientType -eq "UserMailbox") {

		$mailenabled = $True

	} else {

		$mailenabled = $False

	}

	if ($tmpUser.Company) {

		$company = $tmpUser.Company

	} else {

		$company = "malabs"

	}

	# Hide buttons
	$backButton.Enabled = $False
	$cancelButton.Enabled = $False
	$OKButton.Enabled = $False

	# 1 of 8, resetting custom attributes
	if ($mailenabled) {

		$errorMessage1.Text = "[1/8] Resetting custom attributes and hidden from Exchange Address lists."
		$Form2.Refresh()
		Set-Mailbox $userAlias -CustomAttribute1 "" -CustomAttribute2 "" -CustomAttribute3 "" -HiddenFromAddressListsEnabled $True -domaincontroller $DC
		"1. Resetting custom attributes and hidden from Exchange Address lists." > $log

	} else {

		$errorMessage1.Text = "[1/8] Skipped."
		$Form2.Refresh()
		"1. Skipped resetting custom attributes and hidden from Exchange Address lists." > $log

	}

	Start-Sleep -s 1

	# 2 of 8, disabling mailbox features
	if ($mailenabled) {

		$errorMessage2.Text = "[2/8] Disabling mailbox features."
		$Form2.Refresh()
		Set-CASMailbox -Identity $UserAlias -ActiveSyncEnabled $False -OWAEnabled $False -MAPIBlockOutlookRpcHttp $True -Confirm:$False -domaincontroller $DC
		"" >> $log
		"2. Disabling mailbox features." >> $log

	} else {

		$errorMessage2.Text = "[2/8] Skipped."
		$Form2.Refresh()
		"" >> $log
		"2. Skipped disabling mailbox features." >> $log

	}

	Start-Sleep -s 1

	# 3 of 8, removing group membership
	$errorMessage3.Text = "[3/8] Removing group membership."
	$Form2.Refresh()
	#To Bind:
	$userdn = $tmpUser.distinguishedname
	$user = [adsi]"LDAP://$userdn"
	foreach ($group in $user.memberof) {

		([adsi]"LDAP://$group").Remove("LDAP://$userdn")

	}

	"" >> $log
	"3. Removing group membership." >> $log
	Start-Sleep -s 1

	# 4 of 8, updating Address List
	if ($mailenabled) {

		$errorMessage4.Text = "[4/8] Updating address list `"$Company Address List`"."
		$Form2.Refresh()
		Update-AddressList "$company Address List" -domaincontroller $DC
		"" >> $log
		"4. Updating address list `"$Company Address List`"." >> $log

	} else {

		$errorMessage4.Text = "[4/8] Skipped."
		$Form2.Refresh()
		"" >> $log
		"4. Skipped updating address list `"$Company Address List`"." >> $log

	}
	Start-Sleep -s 1

	# 5 of 8, updating Offline Address Book
	if ($mailenabled) {

		$errorMessage5.Text = "[5/8] Updating offline address book `"$company OAB`"."
		$Form2.Refresh()
		Update-OfflineAddressBook "$company OAB" -domaincontroller $DC
		"" >> $log
		"5. Updating offline address book `"$company OAB`"." >> $log

	} else {

		$errorMessage5.Text = "[5/8] Skipped."
		$Form2.Refresh()
		"" >> $log
		"5. Skipped updating offline address book `"$company OAB`"." >> $log

	}
	Start-Sleep -s 1

	# 6 of 8, moving user
	$errorMessage6.Text = "[6/8] Moving user to `"/Disabled/$company/Users`" OU."
	$Form2.Refresh()
	$to = [adsi]"LDAP://OU=Users,OU=$company,OU=Disabled,DC=malabs,DC=com"
	$user.psbase.MoveTo($to)
	"" >> $log
	"6. Moving user to `"/Disabled/$company/Users`" OU." >> $log
	Start-Sleep -s 1

	# 7 of 8, removing home directory
	$userhome = $user.homeDirectory
	$archive = "\\mlfs3\userdirs\Archive"
	if ($userhome) {

		if (Test-Path $userhome) {

			if (@( Get-ChildItem $userhome -Name).Count -eq 0) {

				$errorMessage7.Text = "[7/8] Removing user home directory `"$userhome`"."
				$Form2.Refresh()
				Remove-Item $userhome -Force -Recurse
				"" >> $log
				"7. Removing user home directory `"$userhome`"." >> $log

			} else {

				$errorMessage7.Text = "[7/8] Archiving user home directory `"$userhome`"."
				$Form2.Refresh()
				Move-Item $userhome $archive
				"" >> $log
				"7. Archiving user home directory `"$userhome`"." >> $log

			}

		}

	} else {

		$errorMessage7.Text = "[7/8] Skipped."
		$Form2.Refresh()
		"" >> $logarchive
		"7. Skipped archiving user home directory." >> $log

	}
	Start-Sleep -s 1

	# 8 of 8, disabling user
	$errorMessage8.Text = "[8/8] Disabling user."
	$Form2.Refresh()
	$user.psbase.invokeset("AccountDisabled", "True")
	$user.setinfo()
	if ($user.psbase.invokeget("AccountDisabled") -eq $False) {

		$errorMessage8.Text = "ERROR: Unable to disable user `"$UserFirst $UserLast ($UserAlias)`". Check the command syntax."

	}
	"" >> $log
	"8. Disabling user." >> $log
	Start-Sleep -s 1

	$OKButton.visible = $False
	$finishButton.TabIndex = 0
	$finishButton.TabStop = $True
	$finishButton.visible = $True
	$errorMessage9.Text = "User `"$UserFirst $UserLast ($UserAlias)`" disabled sucessfully. Check log for errors."
	"" >> $log
	"User `"$userFirst $UserLast ($UserAlias)`" disabled sucessfully." >> $log

	#	Copy-Item $log $logarchive
	$subject = "User $UserFirst $UserLast ($UserAlias) Disabled"
	$body = "<font face=Arial size=2>THIS IS A SYSTEM GENERATED MESSAGE. PLEASE DO NOT REPLY."
	$body = $body + "<p>The following user has been disabled:"
	$body = $body + "<p>Name: $UserFirst $UserLast"
	$body = $body + "<br>User Logon Name: $UserAlias"
	$body = $body + "<br>E-mail: <a href=`"mailto:$UserEmail$UserEmailDomain`">$UserEmail$UserEmailDomain</a>"
	$body = $body + "<br>Organizational Unit: $UserOU"
	$body = $body + "<br>Server: $MailboxServer"
	$body = $body + "<br>Database: $Database"
	$body = $body + "<p>Exchange Administrator<br>Ma Lab, Inc.</font>"
	$message = New-Object Net.Mail.MailMessage ($emailFrom, $emailTo, $subject, $body)
	$message.IsBodyHTML = $True
	$smtp.Send($message)

}

function Create-Group {

	# Hide buttons
	$BackButton4.Enabled = $False
	$CancelButton4.Enabled = $False
	$OKButton4.Enabled = $False

	# 1 of 4, creating group
	$log = $logpath + $GroupAlias + "-created.log"
	$ErrorMsg2.Text = "[1/4] Creating group...please wait."
	$Form4.Refresh()
	$result = New-DistributionGroup -Name "$GroupName" -Type Distribution -SamAccountName "$GroupAlias" -OrganizationalUnit "$Forest$GroupOU" -domaincontroller $DC
	"1. Creating group." > $log
	Start-Sleep -s 1

	if ($result -ne $null) {

		# 2 of 4, setting custom attribute
		switch ($GroupDomain) {

			"@superbiiz.com" {

				$customattr = "CustomAttribute1";
				$GroupCompany = "Superbiiz"

			}

			"@supertalent.com" {

				$customattr = "CustomAttribute3";
				$GroupCompany = "Supertalent"

			}

			default {

				$customattr = "CustomAttribute2";
				$GroupCompany = "malabs"

			}

		}

		$ErrorMsg3.Text = "[2/4] Setting $customattr attribute to `"$GroupCompany`"."
		$Form4.Refresh()
		switch ($GroupDomain) {

			"@superbiiz.com" {

				Set-DistributionGroup "$GroupAlias" -CustomAttribute1 "$GroupCompany" -domaincontroller $DC -RequireSenderAuthenticationEnabled $GroupRSA -EmailAddressPolicyEnabled $False -EmailAddresses "SMTP:$GroupEmail$GroupDomain"

			}

			"@supertalent.com" {

				Set-DistributionGroup "$GroupAlias" -CustomAttribute3 "$GroupCompany" -domaincontroller $DC -RequireSenderAuthenticationEnabled $GroupRSA -EmailAddressPolicyEnabled $False -EmailAddresses "SMTP:$GroupEmail$GroupDomain"

			}

			default {

				Set-DistributionGroup "$GroupAlias" -CustomAttribute2 "$GroupCompany" -domaincontroller $DC -RequireSenderAuthenticationEnabled $GroupRSA -EmailAddressPolicyEnabled $False -EmailAddresses "SMTP:$GroupEmail$GroupDomain","smtp:$GroupEmail@exchange.malabs.com"

			}

		}
		"" >> $log
		"2. Setting $customattr attribute to `"$GroupCompany`"." >> $log
		Start-Sleep -s 1

		# 3 of 4, updating Address List
		$ErrorMsg4.Text = "[3/4] Updating address list `"$GroupCompany Address List`"."
		$Form4.Refresh()
		Update-AddressList "$GroupCompany Address List" -domaincontroller $DC
		"" >> $log
		"3. Updating address list `"$GroupCompany Address List`"." >> $log
		Start-Sleep -s 1

		# 4 of 4, updating Offline Address Book
		$ErrorMsg5.Text = "[4/4] Updating offline address book `"$GroupCompany OAB`"."
		$Form4.Refresh()
		Update-OfflineAddressBook "$GroupCompany OAB" -domaincontroller $DC
		"" >> $log
		"4. Updating offline address book `"$GroupCompany OAB`"." >> $log
		Start-Sleep -s 1

		$OKButton4.visible = $False
		$FinishButton4.TabIndex = 0
		$FinishButton4.visible = $True
		$ErrorMsg6.Text = "Group `"$GroupName ($GroupAlias)`" created sucessfully. Check the log for errors."
		"" >> $log
		"Group `"$GroupName ($GroupAlias)`" created sucessfully." >> $log

	} else {

		$ErrorMsg6.Text = "ERROR: Unable to create group `"$GroupName [$GroupEmail]`". Check the command syntax."
		"" >> $log
		"ERROR: Unable to create group `"$GroupName [$GroupEmail]`". Check the command syntax." >> $log

	}

}

# Enable Lync Account 
function Enable-LyncAccount {

	Enable-CsUser -Identity "$AccountName" -RegistrarPool "$registrarPool" -SipAddress "$SipAddress"

	$CsAdUser = Get-CsAdUser

}

function Check-IsGroupMember {

	param($user,$grp)

	$strFilter = "(&(objectClass=Group)(name=" + $grp + "))"
	$objDomain = New-Object System.DirectoryServices.DirectoryEntry
	$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
	$objSearcher.SearchRoot = $objDomain
	$objSearcher.PageSize = 1000
	$objSearcher.Filter = $strFilter

	$objSearcher.SearchScope = "Subtree"
	$colResults = $objSearcher.FindOne()
	$objItem = $colResults.properties
	([string]$objItem.member).contains($user)

}

function Add-Group {

	#Get the information of the user
	$tmpUser = Get-User -Filter "SamAccountName -eq '$UserAlias'" -domaincontroller $DC

	if ($tmpUser) {

		$UserDSName = $tmpUser.distinguishedname
		$UserFullName = $tmpUser.Name
		$Userpath = "LDAP://" + $UserDSName

		#Get the the path of the group
		$searcher_G = New-Object system.directoryservices.directorysearcher;
		$grp_G = New-Object system.directoryservices.directoryentry;

		#Network file share (FS) drive successful group
		$SuccessfulGroup = @()

		$ErrorMsg13.Text = "[1/2]Adding `"$UserFirst $UserLast ($UserAlias)`" to FS...please wait"
		$ErrorMsg13.ForeColor = "Green"
		"[1/2]Adding `"$UserFirst $UserLast ($UserAlias)`" to FS" >> $log

		Start-Sleep -s 1

		if ($GroupName_Vasto.Length -eq 0) {

			$ErrorMsg14.Text = "[2/2]No Target name selected for FS !"
			$ErrorMsg14.ForeColor = "Red"
			"[2/2]No Target name selected for FS !" >> $log

		}

		#Loop the all the group in a list
		for ($j = 0; $j -lt $GroupName_Vasto.Length; $j++) {

			$SecurityGroupName = $GroupName_Vasto[$j]
			$PathDirection1 = "OU=File Share, OU=Groups, DC=malabs, DC=com"
			$root_G = [adsi]("LDAP://" + $PathDirection1)
			$searcher_G.SearchRoot = $root_G
			$result_G = $searcher_G.FindAll() | Where-Object { $_.properties.Name -eq "$SecurityGroupName" }
			$grp_g = $result_g.GetDirectoryEntry()
			$checkUser = Check-IsGroupMember $UserFullName $SecurityGroupName
			#check the exist user

			if ($checkUser) {

				$ErrorMsg14.Text = "[2/2]$UserFullName already exsist in $SecurityGroupName"
				$ErrorMsg14.ForeColor = "Red"

				continue

				Start-Sleep -s 1

			} else {

				#add user to that Group
				$grp_g.psbase.Invoke("Add", $Userpath)
				$grp_g.psbase.CommitChanges()
				$SuccessfulGroup += $SecurityGroupName
				$ErrorMsg14.Text = "[2/2]$UserFullName insert into $SecurityGroupName successfully "
				$ErrorMsg14.ForeColor = "Green"

				Start-Sleep -s 1

			}

		}

		$ErrorMsg15.Text = "[1/2]Adding `"$UserFirst $UserLast ($UserAlias)`" to NDFS...please wait"
		$ErrorMsg15.ForeColor = "Green"
		"[1/2]Adding `"$UserFirst $UserLast ($UserAlias)`" to NDFS" >> $log

		Start-Sleep -s 1

		if ($GroupName_DFS.Length -eq 0) {

			$ErrorMsg16.Text = "[2/2]No Target name selected for NDFS !"
			$ErrorMsg16.ForeColor = "Red"
			"[2/2]No Target name selected for NDFS !" >> $log

		}

		for ($j = 0; $j -lt $GroupName_DFS.Length; $j++) {

			$SecurityGroupName = $GroupName_DFS[$j]
			$PathDirection1 = "OU=DFS,OU=Groups,DC=malabs,DC=com"
			$root_G = [adsi]("LDAP://" + $PathDirection1)
			$searcher_G.SearchRoot = $root_G
			$result_G = $searcher_G.FindAll() | Where-Object { $_.properties.Name -eq "$SecurityGroupName" }
			$grp_g = $result_g.GetDirectoryEntry()
			$checkUser = Check-IsGroupMember $UserFullName $SecurityGroupName
			#check the exist user

			if ($checkUser) {

				$ErrorMsg16.Text = "[2/2]$UserFullName already exsist in $SecurityGroupName"
				$ErrorMsg16.ForeColor = "Red"

				continue

				Start-Sleep -s 1

			} else {

				#add user to that Group
				$grp_g.psbase.Invoke("Add", $Userpath)
				$grp_g.psbase.CommitChanges()
				$SuccessfulGroup += $SecurityGroupName
				$ErrorMsg16.Text = "[2/2]$UserFullName insert into $SecurityGroupName successfully "
				$ErrorMsg16.ForeColor = "Green"

				Start-Sleep -s 1

			}

		}
		
		#Websense List group
		$SuccessfulWsGroup = @()

		#Add the user into the websense department list   
		$ErrorMsg17.Text = "[1/2]Adding `"$UserFirst $UserLast ($UserAlias)`" to Websense Department...please wait"
		$ErrorMsg17.ForeColor = "Green"
		"[1/2]Adding `"$UserFirst $UserLast ($UserAlias)`" to Websense Department" >> $log

		Start-Sleep -s 1

		if ($GroupName_Dept.Length -eq 0) {

			$ErrorMsg18.Text = "[2/2]No Target name selected for Websense Department !"
			$ErrorMsg18.ForeColor = "Red"
			"[2/2]No Target name selected for Websense Department !" >> $log

		}
		
		for ($j = 0; $j -lt $GroupName_Dept.Length; $j++) {

			$SecurityGroupName = $GroupName_Dept[$j]
			$PathDirection1 = "OU=Department,OU=Groups,DC=malabs,DC=com"
			$root_G = [adsi]("LDAP://" + $PathDirection1)
			$searcher_G.SearchRoot = $root_G
			$result_G = $searcher_G.FindAll() | Where-Object { $_.properties.Name -eq "$SecurityGroupName" }
			$grp_g = $result_g.GetDirectoryEntry()
			$checkUser = Check-IsGroupMember $UserFullName $SecurityGroupName
			#check the exist user

			if ($checkUser) {

				$ErrorMsg18.Text = "[2/2]$UserFullName already exsist in $SecurityGroupName"
				$ErrorMsg18.ForeColor = "Red"

				continue

				Start-Sleep -s 1

			} else {

				#add user to that Group
				$grp_g.psbase.Invoke("Add",$Userpath)
				$grp_g.psbase.CommitChanges()
				$SuccessfulWsGroup += $SecurityGroupName
				$ErrorMsg18.Text = "[2/2]$UserFullName insert into $SecurityGroupName successfully "
				$ErrorMsg18.ForeColor = "Green"

				Start-Sleep -s 1

			}

		}

		#Add the user into the websense department list
		$ErrorMsg19.Text = "[1/2]Adding `"$UserFirst $UserLast ($UserAlias)`" to Websense List...please wait"
		$ErrorMsg19.ForeColor = "Green"
		"[1/2]Adding `"$UserFirst $UserLast ($UserAlias)`" to Websense List" >> $log

		Start-Sleep -s 1

		if ($GroupName_List.Length -eq 0) {

			$ErrorMsg20.Text = "[2/2]No Target name selected for Websense List!"
			$ErrorMsg20.ForeColor = "Red"
			"[2/2]No Target name selected for Websense list !" >> $log

		}

		for ($j = 0; $j -lt $GroupName_List.Length; $j++) {

			$SecurityGroupName = $GroupName_List[$j]
			$PathDirection1 = "OU=Websense,OU=Groups,DC=malabs,DC=com"
			$root_G = [adsi]("LDAP://" + $PathDirection1)
			$searcher_G.SearchRoot = $root_G
			$result_G = $searcher_G.FindAll() | Where-Object { $_.properties.Name -eq "$SecurityGroupName" }
			$grp_g = $result_g.GetDirectoryEntry()
			$checkUser = Check-IsGroupMember $UserFullName $SecurityGroupName

			#check the exist user
			if ($checkUser) {

				$ErrorMsg20.Text = "[2/2]$UserFullName already exsist in $SecurityGroupName"
				$ErrorMsg20.ForeColor = "Red"

				continue

				Start-Sleep -s 1

			} else {

				#add user to that Group
				$grp_g.psbase.Invoke("Add",$Userpath)
				$grp_g.psbase.CommitChanges()
				$SuccessfulWsGroup += $SecurityGroupName
				$ErrorMsg20.Text = "[2/2]$UserFullName insert into $SecurityGroupName successfully "
				$ErrorMsg20.ForeColor = "Green"

				Start-Sleep -s 1

			}

		}

		Start-Sleep -s 1


		if ($CheckBranch) {

			$subject = "Network Drive for $UserFirst $UserLast ($UserAlias) Created"
			$body = "<font face=Arial size=2>THIS IS A SYSTEM GENERATED MESSAGE. PLEASE DO NOT REPLY."
			$body = $body + "<p>A new mailbox has been created for:"
			$body = $body + "<p>Name: $UserFirst $UserLast"
			$body = $body + "<br>User Logon Name: $UserAlias"
			$body = $body + "<br>Network Share Drive: $SuccessfulGroup"
			$body = $body + "<br>Websense Seurity Group: $SuccessfulWsGroup"
			$body = $body + "<p>Exchange Administrator<br>Ma Lab, Inc.</font>"
			$message = New-Object Net.Mail.MailMessage ($emailFrom,$emailToBranch,$subject,$body)
			$message.cc.Add("ray.jiang@malabs.com")
			$message.IsBodyHTML = $True
			$smtp.Send($message)

		}

		if ($CheckIT) {

			$subject = "Network Drive for $UserFirst $UserLast ($UserAlias) Created"
			$body = "<font face=Arial size=2>THIS IS A SYSTEM GENERATED MESSAGE. PLEASE DO NOT REPLY."
			$body = $body + "<p>A new mailbox has been created for:"
			$body = $body + "<p>Name: $UserFirst $UserLast"
			$body = $body + "<br>User Logon Name: $UserAlias"
			$body = $body + "<br>Network Share Drive: $SuccessfulGroup"
			$body = $body + "<br>Websense Seurity Group: $SuccessfulWsGroup"
			$body = $body + "<p>Exchange Administrator<br>Ma Lab, Inc.</font>"
			$message = New-Object Net.Mail.MailMessage ($emailFrom,$emailTo,$subject,$body)
			$message.IsBodyHTML = $True
			$smtp.Send($message)

		}

	} else {

		$ErrorMsg13.Text = "The account is not created normally"
		$ErrorMsg13.ForeColor = "Red"

	}

}

################################################################################
# Form 3
################################################################################
function ShowForm3 {

	$launchForm.visible = $False

	$Form3 = New-Object System.Windows.Forms.Form
	$Form3.Text = $title
	$Form3.Size = New-Object System.Drawing.Size (500, 400)
	$Form3.StartPosition = "CenterScreen"

	$Form3.KeyPreview = $True
	$Form3.Add_KeyDown({ 

		if ($_.KeyCode -eq "Enter") {

			$GroupAlias = $LogonName.Text.Trim()
			if ($newgroup -or $found) {

				$GroupName = ToProperCase $Name.Text.Trim()
				$GroupOU = $OU.SelectedItem
					if (!$Email.Text -and $GroupName) {
						$GroupEmail = $GroupName.Replace(" ","")
					} else {
						$GroupEmail = $Email.Text.Trim().ToLower()
					}
					if ($Domain.SelectedItem) {
						$GroupDomain = $Domain.SelectedItem
					} else {
						$GroupDomain = "@ma.local"
					}
					$GroupRSA = $RSA.Checked

				if (ValidateGroup) {

					ShowForm4

				}

			} else {
					if (ValidateGroup) {
						$found = $True
						$tmpGroup = Get-Group -Filter { SamAccountName -EQ $GroupAlias } -domaincontroller $DC
						$OU.SelectedItem = $([string]$tmpGroup.Identity.Parent).substring($Forest.Length)
						$Name.Text = $tmpGroup.Name
						if ($tmpUser.WindowsEmailAddress.Domain) { $Domain.SelectedItem = "@" + $tmpUser.WindowsEmailAddress.Domain }
						$Email.Text = $tmpGroup.WindowsEmailAddress.Local
						$Domain.Enabled = $True
						$Email.Enabled = $True
						$RSA.Enabled = $True
						$Form3.Refresh()
					}
				}
			} })
	$Form3.Add_KeyDown({ if ($_.KeyCode -eq "Escape") { $Form3.close() } })

	$NextButton3 = New-Object System.Windows.Forms.Button
	$NextButton3.Location = New-Object System.Drawing.Size (280,300)
	$NextButton3.Size = New-Object System.Drawing.Size (75,23)
	$NextButton3.Text = "Next >"
	$NextButton3.TabIndex = 7
	$NextButton3.Add_Click({
			$GroupAlias = $LogonName.Text.Trim()
			if ($newgroup -or $found) {
				$GroupName = ToProperCase $Name.Text.Trim()
				$GroupOU = $OU.SelectedItem
				if (!$Email.Text -and $GroupName) {
					$GroupEmail = $GroupName.Replace(" ","")
				} else {
					$GroupEmail = $Email.Text.Trim().ToLower()
				}
				if ($Domain.SelectedItem) {
					$GroupDomain = $Domain.SelectedItem
				} else {
					$GroupDomain = "@ma.local"
				}
				$GroupRSA = $RSA.Checked
				if (ValidateGroup) { ShowForm4 }
			} else {
				if (ValidateGroup) {
					$found = $True
					$tmpGroup = Get-Group -Filter { SamAccountName -EQ $GroupAlias } -domaincontroller $DC
					$OU.SelectedItem = $([string]$tmpGroup.Identity.Parent).substring($Forest.Length)
					$Name.Text = $tmpGroup.Name
					if ($tmpUser.WindowsEmailAddress.Domain) { $Domain.SelectedItem = "@" + $tmpUser.WindowsEmailAddress.Domain }
					$Email.Text = $tmpGroup.WindowsEmailAddress.Local
					$Domain.Enabled = $True
					$Email.Enabled = $True
					$RSA.Enabled = $True
					$Form3.Refresh()
				}
			}
		})
	$Form3.Controls.Add($NextButton3)

	$CancelButton3 = New-Object System.Windows.Forms.Button
	$CancelButton3.Location = New-Object System.Drawing.Size (365,300)
	$CancelButton3.Size = New-Object System.Drawing.Size (75,23)
	$CancelButton3.Text = "Cancel"
	$CancelButton3.TabIndex = 8
	$CancelButton3.Add_Click({ $Form3.close() })
	$Form3.Controls.Add($CancelButton3)

	$BackButton3 = New-Object System.Windows.Forms.Button
	$BackButton3.Location = New-Object System.Drawing.Size (205,300)
	$BackButton3.Size = New-Object System.Drawing.Size (75,23)
	$BackButton3.Text = "< Back"
	$BackButton3.TabIndex = 9
	$BackButton3.Add_Click({ $Form3.visible = $False; $launchForm.visible = $True })
	$Form3.Controls.Add($BackButton3)

	# Error Message Box
	$ErrorMsg = New-Object System.Windows.Forms.Label
	$ErrorMsg.Location = New-Object System.Drawing.Size (20,10)
	$ErrorMsg.Size = New-Object System.Drawing.Size (500,20)
	$ErrorMsg.Text = $null
	$ErrorMsg.ForeColor = "Red"
	$Form3.Controls.Add($ErrorMsg)

	# Title Label
	$TitleLabel = New-Object System.Windows.Forms.Label
	$TitleLabel.Location = New-Object System.Drawing.Size (20,40)
	$TitleLabel.Size = New-Object System.Drawing.Size (400,20)
	$TitleLabel.Text = "$Forest"
	$TitleLabel.Font = New-Object System.Drawing.Font ("Arial",8,([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)),[System.Drawing.GraphicsUnit]::Point,([System.Byte](0)))
	$Form3.Controls.Add($TitleLabel)

	# OU Label
	$OULabel = New-Object System.Windows.Forms.Label
	$OULabel.Location = New-Object System.Drawing.Size (20,70)
	$OULabel.Size = New-Object System.Drawing.Size (120,20)
	$OULabel.Text = "Organizational Unit:"
	$Form3.Controls.Add($OULabel)

	# Name Label
	$NameLabel = New-Object System.Windows.Forms.Label
	$NameLabel.Location = New-Object System.Drawing.Size (20,100)
	$NameLabel.Size = New-Object System.Drawing.Size (130,20)
	$NameLabel.Text = "Name:"
	$Form3.Controls.Add($NameLabel)

	# Alias Label
	$LogonLabel = New-Object System.Windows.Forms.Label
	$LogonLabel.Location = New-Object System.Drawing.Size (20,130)
	$LogonLabel.Size = New-Object System.Drawing.Size (120,20)
	$LogonLabel.Text = "Alias:"
	$Form3.Controls.Add($LogonLabel)

	# Email Label
	$EmailLabel = New-Object System.Windows.Forms.Label
	$EmailLabel.Location = New-Object System.Drawing.Size (20,160)
	$EmailLabel.Size = New-Object System.Drawing.Size (130,20)
	$EmailLabel.Text = "E-mail:"
	$Form3.Controls.Add($EmailLabel)

	# Mailflow Settings Label
	$MailflowLabel = New-Object System.Windows.Forms.Label
	$MailflowLabel.Location = New-Object System.Drawing.Size (20,190)
	$MailflowLabel.Size = New-Object System.Drawing.Size (120,20)
	$MailflowLabel.Text = "Mailflow Settings:"
	$Form3.Controls.Add($MailflowLabel)

	# Require Sender Authentication Label
	$RSALabel = New-Object System.Windows.Forms.Label
	$RSALabel.Location = New-Object System.Drawing.Size (165,190)
	$RSALabel.Size = New-Object System.Drawing.Size (200,20)
	$RSALabel.Text = "Require all senders are authenticated"
	$Form3.Controls.Add($RSALabel)

	<#
		OU Combo Box
	#>
	$OUhash = @{}
	$OU = New-Object System.Windows.Forms.ComboBox
	$OU.Location = New-Object System.Drawing.Size (150, 70)
	$OU.Size = New-Object System.Drawing.Size (290, 20)
	[void]$OU.Items.Add("/Users")
	#$OUhash.Add("/Users", "Users")
	$root = [adsi]''
	$searcher = New-Object System.DirectoryServices.DirectorySearcher ($root)
	$searcher.Filter = "(&(objectClass=organizationalUnit)(name=Users))"
	[void]$searcher.PropertiesToLoad.Add("canonicalName")
	[void]$searcher.PropertiesToLoad.Add("Name")
	$searcherall = $searcher.FindAll()
	foreach ($person in $searcherall) {

		[string]$ent = $person.properties.canonicalname
		$OUhash.Add($ent.substring($ent.IndexOf("/"),$ent.Length - $ent.IndexOf("/")),$ent)
		[void]$OU.Items.Add($ent.substring($ent.IndexOf("/"),$ent.Length - $ent.IndexOf("/")))

	}
	$OU.DropDownStyle = 2
	$OU.Sorted = $True
	$OU.TabIndex = 0
	if (!$newgroup) {

		$OU.Enabled = $False 

	}
	$Form3.Controls.Add($OU)

	# OU Message Box
	$OUMsg = New-Object System.Windows.Forms.Label
	$OUMsg.Location = New-Object System.Drawing.Size (450, 75)
	$OUMsg.Size = New-Object System.Drawing.Size (20, 20)
	if ($newgroup) { $OUMsg.Text = "*" }
	$OUMsg.ForeColor = "Red"
	$Form3.Controls.Add($OUMsg)

	# Name Text Box
	$Name = New-Object System.Windows.Forms.TextBox
	$Name.Location = New-Object System.Drawing.Size (150, 100)
	$Name.Size = New-Object System.Drawing.Size (140, 20)
	$Name.TabIndex = 2
	$Name.MaxLength = 40
	if (!$newgroup) { $Name.Enabled = $False }
	$Form3.Controls.Add($Name)

	# Name Message Box
	$NameMsg = New-Object System.Windows.Forms.Label
	$NameMsg.Location = New-Object System.Drawing.Size (450, 105)
	$NameMsg.Size = New-Object System.Drawing.Size (20, 20)
	if ($newgroup) { $NameMsg.Text = "*" }
	$NameMsg.ForeColor = "Red"
	$Form3.Controls.Add($NameMsg)

	# Alias Text Box
	$LogonName = New-Object System.Windows.Forms.TextBox
	$LogonName.Location = New-Object System.Drawing.Size (150, 130)
	$LogonName.Size = New-Object System.Drawing.Size (140, 20)
	$LogonName.MaxLength = 40
	$LogonName.TabIndex = 3
	#$LogonName.CharacterCasing = 2
	$Form3.Controls.Add($LogonName)

	# Alias Message Box
	$LogonMsg = New-Object System.Windows.Forms.Label
	$LogonMsg.Location = New-Object System.Drawing.Size (450, 135)
	$LogonMsg.Size = New-Object System.Drawing.Size (20, 20)
	$LogonMsg.Text = "*"
	$LogonMsg.ForeColor = "Red"
	$Form3.Controls.Add($LogonMsg)

	# Email Text Box
	$Email = New-Object System.Windows.Forms.TextBox
	$Email.Location = New-Object System.Drawing.Size (150, 160)
	$Email.Size = New-Object System.Drawing.Size (140, 20)
	$Email.MaxLength = 100
	$Email.TabIndex = 4
	if (!$newgroup) { $Email.Enabled = $False }
	$Email.Text = $null
	$Form3.Controls.Add($Email)

	# Domain Combo Box
	$Domain = New-Object System.Windows.Forms.ComboBox
	$Domain.Location = New-Object System.Drawing.Size (300, 160)
	$Domain.Size = New-Object System.Drawing.Size (140, 20)
	[void]$Domain.Items.Add("@ma.local")
	[void]$Domain.Items.Add("@superbiiz.com")
	[void]$Domain.Items.Add("@supertalent.com")
	$Domain.DropDownStyle = 2
	$Domain.Sorted = $True
	$Domain.TabIndex = 5
	if (!$newgroup) { $Domain.Enabled = $False }
	$Form3.Controls.Add($Domain)

	# Email Message Box
	$EmailMsg = New-Object System.Windows.Forms.Label
	$EmailMsg.Location = New-Object System.Drawing.Size (450, 165)
	$EmailMsg.Size = New-Object System.Drawing.Size (20, 20)
	$EmailMsg.Text = $null
	$EmailMsg.ForeColor = "Red"
	$Form3.Controls.Add($EmailMsg)

	# Require Sender Authentication Check Box
	$RSA = New-Object System.Windows.Forms.CheckBox
	$RSA.Location = New-Object System.Drawing.Size (150, 188)
	$RSA.Size = New-Object System.Drawing.Size (140, 20)
	$RSA.Checked = $False
	$RSA.TabIndex = 6
	if (!$newgroup) {

		$RSA.Enabled = $False
	
	}

	$Form3.Controls.Add($RSA)

	$Form3.Topmost = $True
	$Form3.Add_Shown({ $Form3.Activate() })
	[void]$Form3.ShowDialog()

}

################################################################################
# Form 4
################################################################################
function ShowForm4 {

	$Form3.visible = $False

	$Form4 = New-Object System.Windows.Forms.Form
	$Form4.Text = $title
	$Form4.Size = New-Object System.Drawing.Size (500,400)
	$Form4.StartPosition = "CenterScreen"

	$Form4.KeyPreview = $True
	$Form4.Add_KeyDown({

		if ($_.KeyCode -eq "Enter") {

			if ($OKButton4.visible) {

				Create-Group

			}

			if ($FinishButton4.visible) { 

				$Form3.close();
				$Form4.close();
				$launchForm.visible = $True
				
			}

		}
	
	})
	$Form4.Add_KeyDown({
	
		if ($_.KeyCode -eq "Escape") { $Form3.close(); $Form4.close() } })

	$BackButton4 = New-Object System.Windows.Forms.Button
	$BackButton4.Location = New-Object System.Drawing.Size (205,300)
	$BackButton4.Size = New-Object System.Drawing.Size (75,23)
	$BackButton4.Text = "< Back"
	$BackButton4.TabIndex = 4
	$BackButton4.Add_Click({ $Form4.visible = $False; $Form3.visible = $True })
	$Form4.Controls.Add($BackButton4)

	$OKButton4 = New-Object System.Windows.Forms.Button
	$OKButton4.Location = New-Object System.Drawing.Size (280,300)
	$OKButton4.Size = New-Object System.Drawing.Size (75,23)
	$OKButton4.Text = "OK"
	$OKButton4.TabIndex = 1
	$OKButton4.Add_Click({ Create-Group })
	$Form4.Controls.Add($OKButton4)

	$FinishButton4 = New-Object System.Windows.Forms.Button
	$FinishButton4.Location = New-Object System.Drawing.Size (280,300)
	$FinishButton4.Size = New-Object System.Drawing.Size (75,23)
	$FinishButton4.Text = "Finish"
	$FinishButton4.TabIndex = 2
	$FinishButton4.visible = $False
	$FinishButton4.Add_Click({ $Form3.close(); $Form4.close(); $launchForm.visible = $True })
	$Form4.Controls.Add($FinishButton4)

	$CancelButton4 = New-Object System.Windows.Forms.Button
	$CancelButton4.Location = New-Object System.Drawing.Size (365,300)
	$CancelButton4.Size = New-Object System.Drawing.Size (75,23)
	$CancelButton4.Text = "Cancel"
	$CancelButton4.TabIndex = 3
	$CancelButton4.Add_Click({ $Form3.close(); $Form4.close() })
	$Form4.Controls.Add($CancelButton4)

	# Error Message Box
	$ErrorMsg2 = New-Object System.Windows.Forms.Label
	$ErrorMsg2.Location = New-Object System.Drawing.Size (20,10)
	$ErrorMsg2.Size = New-Object System.Drawing.Size (500,20)
	$ErrorMsg2.Text = $null
	$ErrorMsg2.ForeColor = "Green"
	$Form4.Controls.Add($ErrorMsg2)

	$TitleLabel = New-Object System.Windows.Forms.Label
	$TitleLabel.Location = New-Object System.Drawing.Size (20,40)
	$TitleLabel.Size = New-Object System.Drawing.Size (400,20)
	$TitleLabel.Text = "Click OK to create the distribution group."
	$Form4.Controls.Add($TitleLabel)

	# OU Label
	$OULabel = New-Object System.Windows.Forms.Label
	$OULabel.Location = New-Object System.Drawing.Size (20,70)
	$OULabel.Size = New-Object System.Drawing.Size (120,20)
	$OULabel.Text = "Organizational Unit:"
	$Form4.Controls.Add($OULabel)

	$OULabel = New-Object System.Windows.Forms.Label
	$OULabel.Location = New-Object System.Drawing.Size (150,70)
	$OULabel.Size = New-Object System.Drawing.Size (400,20)
	$OULabel.Text = $GroupOU
	$Form4.Controls.Add($OULabel)

	# Name Label
	$NameLabel = New-Object System.Windows.Forms.Label
	$NameLabel.Location = New-Object System.Drawing.Size (20,100)
	$NameLabel.Size = New-Object System.Drawing.Size (130,20)
	$NameLabel.Text = "Name:"
	$Form4.Controls.Add($NameLabel)

	$NameLabel = New-Object System.Windows.Forms.Label
	$NameLabel.Location = New-Object System.Drawing.Size (150,100)
	$NameLabel.Size = New-Object System.Drawing.Size (200,20)
	$NameLabel.Text = "$GroupName"
	$Form4.Controls.Add($NameLabel)

	# User Logon Name Label
	$LogonLabel = New-Object System.Windows.Forms.Label
	$LogonLabel.Location = New-Object System.Drawing.Size (20,130)
	$LogonLabel.Size = New-Object System.Drawing.Size (120,20)
	$LogonLabel.Font = New-Object System.Drawing.Font ("ArialNarrow",8)
	$LogonLabel.Text = "Alias:"
	$Form4.Controls.Add($LogonLabel)

	$LogonLabel = New-Object System.Windows.Forms.Label
	$LogonLabel.Location = New-Object System.Drawing.Size (150,130)
	$LogonLabel.Size = New-Object System.Drawing.Size (350,20)
	$LogonLabel.Text = "$GroupAlias"
	$Form4.Controls.Add($LogonLabel)

	# E-mail Label
	$EmailLabel = New-Object System.Windows.Forms.Label
	$EmailLabel.Location = New-Object System.Drawing.Size (20,160)
	$EmailLabel.Size = New-Object System.Drawing.Size (120,20)
	$EmailLabel.Text = "E-mail:"
	$Form4.Controls.Add($EmailLabel)

	$EmailLabel = New-Object System.Windows.Forms.Label
	$EmailLabel.Location = New-Object System.Drawing.Size (150,160)
	$EmailLabel.Size = New-Object System.Drawing.Size (200,20)
	$EmailLabel.Text = "$GroupEmail$GroupDomain"
	$Form4.Controls.Add($EmailLabel)

	# Mailflow Settings Label
	$MailflowLabel = New-Object System.Windows.Forms.Label
	$MailflowLabel.Location = New-Object System.Drawing.Size (20,190)
	$MailflowLabel.Size = New-Object System.Drawing.Size (120,20)
	$MailflowLabel.Text = "Mailflow Settings:"
	$Form4.Controls.Add($MailflowLabel)

	# Require Sender Authentication Label
	$RSALabel = New-Object System.Windows.Forms.Label
	$RSALabel.Location = New-Object System.Drawing.Size (150,190)
	$RSALabel.Size = New-Object System.Drawing.Size (400,20)
	if ($GroupRSA) {
		$RSALabel.Text = "Require all senders are authenticated: True"
	} else {
		$RSALabel.Text = "Require all senders are authenticated: False"
	}
	$Form4.Controls.Add($RSALabel)

	$Form4.Topmost = $True
	$Form4.Add_Shown({ $Form4.Activate() })
	[void]$Form4.ShowDialog()
}

################################################################################
# Form 5
################################################################################
function ShowForm5
{
	$launchForm.visible = $False

	$Form5 = New-Object System.Windows.Forms.Form
	$Form5.Text = $title
	$Form5.Size = New-Object System.Drawing.Size (500,510)
	$Form5.StartPosition = "CenterScreen"

	$Form5.KeyPreview = $True
	$Form5.Add_KeyDown({ if ($_.KeyCode -eq "Enter") {
				$UserAlias = $LogonName.Text.Trim()

				$UserFirst = ToProperCase ($FirstName.Text.Trim())
				$UserLast = ToProperCase ($LastName.Text.Trim())
				$UserOU = $OU.SelectedItem
				$UserDomain = $Domain.SelectedItem
				$UserDG = $DG.SelectedItem
				$UserOffice = $Office.SelectedItem
				$UserPhone = $Phone.Text.Trim()
				$UserExtension = $Extension.Text.Trim()
				$UserEmail = $Email.Text.Trim()
				if (Validate) { ShowForm6 }
			} })
	$Form5.Add_KeyDown({ if ($_.KeyCode -eq "Escape") { $Form5.close() } })

	$NextButton = New-Object System.Windows.Forms.Button
	$NextButton.Location = New-Object System.Drawing.Size (280,420)
	$NextButton.Size = New-Object System.Drawing.Size (75,23)
	$NextButton.Text = "Next >"
	$NextButton.TabIndex = 10
	$NextButton.Add_Click({
			$UserAlias = $LogonName.Text.Trim()
			$UserFirst = ToProperCase ($FirstName.Text.Trim())
			$UserLast = ToProperCase ($LastName.Text.Trim())
			$UserOU = $OU.SelectedItem
			$UserDomain = $Domain.SelectedItem
			$UserDG = $DG.SelectedItem
			$UserOffice = $Office.SelectedItem
			$UserPhone = $Phone.Text.Trim()
			$UserExtension = $Extension.Text.Trim()
			$UserEmail = $Email.Text.Trim()
			if (Validate) { ShowForm6 }
		})
	$Form5.Controls.Add($NextButton)

	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size (365,420)
	$CancelButton.Size = New-Object System.Drawing.Size (75,23)
	$CancelButton.Text = "Cancel"
	$CancelButton.TabIndex = 11
	$CancelButton.Add_Click({ $Form5.close() })
	$Form5.Controls.Add($CancelButton)

	$BackButton = New-Object System.Windows.Forms.Button
	$BackButton.Location = New-Object System.Drawing.Size (205,420)
	$BackButton.Size = New-Object System.Drawing.Size (75,23)
	$BackButton.Text = "< Back"
	$BackButton.TabIndex = 12
	$BackButton.Add_Click({ $Form5.visible = $False; $launchForm.visible = $True })
	$Form5.Controls.Add($BackButton)

	# Error Message Box
	$ErrorMsg = New-Object System.Windows.Forms.Label
	$ErrorMsg.Location = New-Object System.Drawing.Size (20,10)
	$ErrorMsg.Size = New-Object System.Drawing.Size (500,20)
	$ErrorMsg.Text = $null
	$ErrorMsg.ForeColor = "Red"
	$Form5.Controls.Add($ErrorMsg)

	# Title Label
	$TitleLabel = New-Object System.Windows.Forms.Label
	$TitleLabel.Location = New-Object System.Drawing.Size (20,40)
	$TitleLabel.Size = New-Object System.Drawing.Size (450,20)
	$TitleLabel.Text = "$Forest"
	$TitleLabel.Font = New-Object System.Drawing.Font ("Arial",8,([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)),[System.Drawing.GraphicsUnit]::Point,([System.Byte](0)))
	$Form5.Controls.Add($TitleLabel)

	# OU Label
	$OULabel = New-Object System.Windows.Forms.Label
	$OULabel.Location = New-Object System.Drawing.Size (20,70)
	$OULabel.Size = New-Object System.Drawing.Size (120,20)
	$OULabel.Text = "Organizational Unit:"
	$Form5.Controls.Add($OULabel)

	# First Name Label
	$FirstLabel = New-Object System.Windows.Forms.Label
	$FirstLabel.Location = New-Object System.Drawing.Size (20,100)
	$FirstLabel.Size = New-Object System.Drawing.Size (130,20)
	$FirstLabel.Text = "First Name / Last Name:"
	$Form5.Controls.Add($FirstLabel)

	# User Logon Name Label
	$LogonLabel = New-Object System.Windows.Forms.Label
	$LogonLabel.Location = New-Object System.Drawing.Size (20,130)
	$LogonLabel.Size = New-Object System.Drawing.Size (120,20)
	$LogonLabel.Text = "Alias:"
	$Form5.Controls.Add($LogonLabel)

	# E-mail Label
	$EmailLabel = New-Object System.Windows.Forms.Label
	$EmailLabel.Location = New-Object System.Drawing.Size (20,160)
	$EmailLabel.Size = New-Object System.Drawing.Size (120,20)
	$EmailLabel.Text = "E-mail:"
	$Form5.Controls.Add($EmailLabel)

	# Distribution Group Label
	$DGLabel = New-Object System.Windows.Forms.Label
	$DGLabel.Location = New-Object System.Drawing.Size (20,190)
	$DGLabel.Size = New-Object System.Drawing.Size (120,20)
	$DGLabel.Text = "Distribution Group:"
	$Form5.Controls.Add($DGLabel)

	# Office Label
	$OfficeLabel = New-Object System.Windows.Forms.Label
	$OfficeLabel.Location = New-Object System.Drawing.Size (20,220)
	$OfficeLabel.Size = New-Object System.Drawing.Size (120,20)
	$OfficeLabel.Text = "Office:"
	$Form5.Controls.Add($OfficeLabel)

	# Phone Label
	$PhoneLabel = New-Object System.Windows.Forms.Label
	$PhoneLabel.Location = New-Object System.Drawing.Size (20,250)
	$PhoneLabel.Size = New-Object System.Drawing.Size (120,15)
	$PhoneLabel.Text = "Phone / Extension:"
	$Form5.Controls.Add($PhoneLabel)

	$Phone2Label = New-Object System.Windows.Forms.Label
	$Phone2Label.Location = New-Object System.Drawing.Size (150,270)
	$Phone2Label.Size = New-Object System.Drawing.Size (150,20)
	$Phone2Label.Font = New-Object System.Drawing.Font ("Arial",7)
	$Phone2Label.Text = "e.g. 408-941-0808"
	$Form5.Controls.Add($Phone2Label)

	$Phone3Label = New-Object System.Windows.Forms.Label
	$Phone3Label.Location = New-Object System.Drawing.Size (300,270)
	$Phone3Label.Size = New-Object System.Drawing.Size (140,20)
	$Phone3Label.Font = New-Object System.Drawing.Font ("Arial",7)
	$Phone3Label.Text = "e.g. 808"
	$Form5.Controls.Add($Phone3Label)

	# OU Combo Box
	$OUhash = @{}
	$OU = New-Object System.Windows.Forms.ComboBox
	$OU.Location = New-Object System.Drawing.Size (150,70)
	$OU.Size = New-Object System.Drawing.Size (290,20)
	[void]$OU.Items.Add("/Users")
	#$OUhash.Add("/Users","Users")
	$root = [adsi]''
	$searcher = New-Object System.DirectoryServices.DirectorySearcher ($root)
	$searcher.Filter = "(&(objectClass=organizationalUnit)(name=Users))"
	[void]$searcher.PropertiesToLoad.Add("canonicalName")
	[void]$searcher.PropertiesToLoad.Add("Name")
	$searcherall = $searcher.FindAll()
	foreach ($person in $searcherall) {
		[string]$ent = $person.properties.canonicalname
		$OUhash.Add($ent.substring($ent.IndexOf("/"),$ent.Length - $ent.IndexOf("/")),$ent)
		[void]$OU.Items.Add($ent.substring($ent.IndexOf("/"),$ent.Length - $ent.IndexOf("/")))
	}
	$OU.DropDownStyle = 2
	$OU.Sorted = $True
	$OU.TabIndex = 0
	$Form5.Controls.Add($OU)

	# OU Message Box
	$OUMsg = New-Object System.Windows.Forms.Label
	$OUMsg.Location = New-Object System.Drawing.Size (450,75)
	$OUMsg.Size = New-Object System.Drawing.Size (20,20)
	$OUMsg.Text = "*"
	$OUMsg.ForeColor = "Red"
	$Form5.Controls.Add($OUMsg)

	# First Name Text Box
	$FirstName = New-Object System.Windows.Forms.TextBox
	$FirstName.Location = New-Object System.Drawing.Size (150,100)
	$FirstName.Size = New-Object System.Drawing.Size (140,20)
	$FirstName.TabIndex = 1
	$FirstName.MaxLength = 15
	$Form5.Controls.Add($FirstName)

	# First Name Message Box
	$FirstMsg = New-Object System.Windows.Forms.Label
	$FirstMsg.Location = New-Object System.Drawing.Size (450,105)
	$FirstMsg.Size = New-Object System.Drawing.Size (20,20)
	$FirstMsg.Text = "*"
	$FirstMsg.ForeColor = "Red"
	$Form5.Controls.Add($FirstMsg)

	# Last Name Text Box
	$LastName = New-Object System.Windows.Forms.TextBox
	$LastName.Location = New-Object System.Drawing.Size (300,100)
	$LastName.Size = New-Object System.Drawing.Size (140,20)
	$LastName.MaxLength = 15
	$LastName.TabIndex = 2
	$Form5.Controls.Add($LastName)

	# Logon Name Text Box
	$LogonName = New-Object System.Windows.Forms.TextBox
	$LogonName.Location = New-Object System.Drawing.Size (150,130)
	$LogonName.Size = New-Object System.Drawing.Size (140,20)
	$LogonName.MaxLength = 12
	$LogonName.TabIndex = 3
	$LogonName.CharacterCasing = 2
	$Form5.Controls.Add($LogonName)

	# Domain Combo Box
	$Domain = New-Object System.Windows.Forms.ComboBox
	$Domain.Location = New-Object System.Drawing.Size (300,130)
	$Domain.Size = New-Object System.Drawing.Size (140,20)
	[void]$Domain.Items.Add("@ma.local")
	[void]$Domain.Items.Add("@superbiiz.com")
	[void]$Domain.Items.Add("@supertalent.com")
	$Domain.DropDownStyle = 2
	$Domain.Sorted = $True
	$Domain.TabIndex = 4
	$Form5.Controls.Add($Domain)

	# Logon Name Message Box
	$LogonMsg = New-Object System.Windows.Forms.Label
	$LogonMsg.Location = New-Object System.Drawing.Size (450,135)
	$LogonMsg.Size = New-Object System.Drawing.Size (20,20)
	$LogonMsg.Text = "*"
	$LogonMsg.ForeColor = "Red"
	$Form5.Controls.Add($LogonMsg)

	# E-mail Text Box
	$Email = New-Object System.Windows.Forms.TextBox
	$Email.Location = New-Object System.Drawing.Size (150,160)
	$Email.Size = New-Object System.Drawing.Size (140,20)
	$Email.MaxLength = 40
	$Email.TabIndex = 5
	$Form5.Controls.Add($Email)

	# E-mail Message Box
	$EmailMsg = New-Object System.Windows.Forms.Label
	$EmailMsg.Location = New-Object System.Drawing.Size (450,165)
	$EmailMsg.Size = New-Object System.Drawing.Size (20,20)
	$EmailMsg.Text = "*"
	$EmailMsg.ForeColor = "Red"
	$Form5.Controls.Add($EmailMsg)

	# Distribution Group Combo Box
	$DG = New-Object System.Windows.Forms.ComboBox
	$DG.Location = New-Object System.Drawing.Size (150,190)
	$DG.Size = New-Object System.Drawing.Size (140,20)
	[void]$DG.Items.Add("")
	[void]$DG.Items.Add("Chicago Sales")
	[void]$DG.Items.Add("Chicago Others")
	[void]$DG.Items.Add("GA Sales")
	[void]$DG.Items.Add("GA Others")
	[void]$DG.Items.Add("LA Sales")
	[void]$DG.Items.Add("LA Others")
	[void]$DG.Items.Add("Miami Sales")
	[void]$DG.Items.Add("Miami Others")
	[void]$DG.Items.Add("NJ Sales")
	[void]$DG.Items.Add("NJ Others")
	[void]$DG.Items.Add("Production Line")
	[void]$DG.Items.Add("SJ Accounting")
	[void]$DG.Items.Add("SJ AP")
	[void]$DG.Items.Add("SJ AR")
	[void]$DG.Items.Add("SJ Credit")
	[void]$DG.Items.Add("SJ Data Entry")
	[void]$DG.Items.Add("SJ HR")
	[void]$DG.Items.Add("SJ Inventory Control")
	[void]$DG.Items.Add("SJ Management")
	[void]$DG.Items.Add("SJ Marketing")
	[void]$DG.Items.Add("SJ MIS")
	[void]$DG.Items.Add("SJ OEM Sales")
	[void]$DG.Items.Add("SJ PM")
	[void]$DG.Items.Add("SJ RMA")
	[void]$DG.Items.Add("SJ Sales")
	[void]$DG.Items.Add("SJ Shipping")
	[void]$DG.Items.Add("SJ Tech")
	[void]$DG.Items.Add("SJ Warehouse")
	[void]$DG.Items.Add("Wuhan AR")
	[void]$DG.Items.Add("Wuhan AP")
	[void]$DG.Items.Add("Wuhan Credit")
	[void]$DG.Items.Add("Wuhan Inventory")
	[void]$DG.Items.Add("Wuhan RMA")
	[void]$DG.Items.Add("Wuhan Sales")
	$DG.DropDownStyle = 2
	$DG.Sorted = $True
	$DG.TabIndex = 6
	$Form5.Controls.Add($DG)

	# Office Combo Box
	$Office = New-Object System.Windows.Forms.ComboBox
	$Office.Location = New-Object System.Drawing.Size (150,220)
	$Office.Size = New-Object System.Drawing.Size (140,20)
	[void]$Office.Items.Add("")
	[void]$Office.Items.Add("Chicago")
	[void]$Office.Items.Add("China - Guangzhou")
	[void]$Office.Items.Add("China - Shenzhen")
	[void]$Office.Items.Add("China - Wuhan")
	[void]$Office.Items.Add("LA")
	[void]$Office.Items.Add("Georgia")
	[void]$Office.Items.Add("Korea")
	[void]$Office.Items.Add("Miami")
	[void]$Office.Items.Add("New Jersey")
	[void]$Office.Items.Add("San Jose")
	[void]$Office.Items.Add("Other")
	$Office.DropDownStyle = 2
	$Office.Sorted = $True
	$Office.TabIndex = 7
	$Form5.Controls.Add($Office)

	# Phone Text Box
	$Phone = New-Object System.Windows.Forms.TextBox
	$Phone.Location = New-Object System.Drawing.Size (150,250)
	$Phone.Size = New-Object System.Drawing.Size (140,20)
	$Phone.MaxLength = 12
	$Phone.TabIndex = 8
	$Phone.Text = $null
	$Phone.Enabled = $True
	$Form5.Controls.Add($Phone)

	# Extension Text Box
	$Extension = New-Object System.Windows.Forms.TextBox
	$Extension.Location = New-Object System.Drawing.Size (300,250)
	$Extension.Size = New-Object System.Drawing.Size (140,20)
	$Extension.MaxLength = 3
	$Extension.TabIndex = 9
	$Extension.Text = $null
	$Extension.Enabled = $True
	$Form5.Controls.Add($Extension)

	# Phone Message Box
	$PhoneMsg = New-Object System.Windows.Forms.Label
	$PhoneMsg.Location = New-Object System.Drawing.Size (450,255)
	$PhoneMsg.Size = New-Object System.Drawing.Size (20,20)
	$PhoneMsg.Text = $null
	$PhoneMsg.ForeColor = "Red"
	$Form5.Controls.Add($PhoneMsg)

	$Form5.Topmost = $True
	$Form5.Add_Shown({ $Form5.Activate() })
	[void]$Form5.ShowDialog()
}

################################################################################
# Form 6
################################################################################
function ShowForm6
{
	$Form5.visible = $False

	$Form6 = New-Object System.Windows.Forms.Form
	$Form6.Text = $title
	$Form6.Size = New-Object System.Drawing.Size (500,510)
	$Form6.StartPosition = "CenterScreen"

	$Form6.KeyPreview = $True
	$Form6.Add_KeyDown({ if ($_.KeyCode -eq "Enter") {
				if ($OKButton.visible) {
					$logfile = $UserAlias + "-created.log"
					CreateContact
				}
				if ($FinishButton.visible) { $Form5.close(); $Form6.close(); $launchForm.visible = $True }
			} })
	$Form6.Add_KeyDown({ if ($_.KeyCode -eq "Escape") { $launchForm.close(); $Form5.close(); $Form6.close() } })

	if ($UserExtension) { $UserExtension = "x" + $UserExtension }

	$OKButton = New-Object System.Windows.Forms.Button
	$OKButton.Location = New-Object System.Drawing.Size (280,420)
	$OKButton.Size = New-Object System.Drawing.Size (75,23)
	$OKButton.Text = "OK"
	$OKButton.TabIndex = 2
	$OKButton.Add_Click({
			$logfile = $UserAlias + "-created.log"
			CreateContact
		})
	$Form6.Controls.Add($OKButton)

	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size (365,420)
	$CancelButton.Size = New-Object System.Drawing.Size (75,23)
	$CancelButton.Text = "Cancel"
	$CancelButton.TabIndex = 3
	$CancelButton.Add_Click({ $launchForm.close(); $Form5.close(); $Form6.close() })
	$Form6.Controls.Add($CancelButton)

	$BackButton = New-Object System.Windows.Forms.Button
	$BackButton.Location = New-Object System.Drawing.Size (205,420)
	$BackButton.Size = New-Object System.Drawing.Size (75,23)
	$BackButton.Text = "< Back"
	$BackButton.TabIndex = 4
	$BackButton.Add_Click({ $Form6.visible = $False; $Form5.visible = $True })
	$Form6.Controls.Add($BackButton)

	$FinishButton = New-Object System.Windows.Forms.Button
	$FinishButton.Location = New-Object System.Drawing.Size (280,420)
	$FinishButton.Size = New-Object System.Drawing.Size (75,23)
	$FinishButton.Text = "Finish"
	$FinishButton.TabIndex = 1
	$FinishButton.visible = $False
	$FinishButton.Add_Click({ $Form5.close(); $Form6.close(); $launchForm.visible = $True })
	$Form6.Controls.Add($FinishButton)

	# Error Message Box
	$ErrorMsg = New-Object System.Windows.Forms.Label
	$ErrorMsg.Location = New-Object System.Drawing.Size (20,10)
	$ErrorMsg.Size = New-Object System.Drawing.Size (500,20)
	$ErrorMsg.Text = $null
	$ErrorMsg.ForeColor = "Green"
	$Form6.Controls.Add($ErrorMsg)

	$TitleLabel = New-Object System.Windows.Forms.Label
	$TitleLabel.Location = New-Object System.Drawing.Size (20,40)
	$TitleLabel.Size = New-Object System.Drawing.Size (500,20)
	$TitleLabel.Text = "Click OK to " + $title.ToLower() + "."
	$TitleLabel.Font = New-Object System.Drawing.Font ("Arial",8,([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)),[System.Drawing.GraphicsUnit]::Point,([System.Byte](0)))
	$Form6.Controls.Add($TitleLabel)

	# OU Label
	$OULabel = New-Object System.Windows.Forms.Label
	$OULabel.Location = New-Object System.Drawing.Size (20,70)
	$OULabel.Size = New-Object System.Drawing.Size (120,20)
	$OULabel.Text = "Organizational Unit:"
	$Form6.Controls.Add($OULabel)

	$OULabel = New-Object System.Windows.Forms.Label
	$OULabel.Location = New-Object System.Drawing.Size (150,70)
	$OULabel.Size = New-Object System.Drawing.Size (350,20)
	$OULabel.Text = $UserOU
	$Form6.Controls.Add($OULabel)

	# Name Label
	$NameLabel = New-Object System.Windows.Forms.Label
	$NameLabel.Location = New-Object System.Drawing.Size (20,100)
	$NameLabel.Size = New-Object System.Drawing.Size (120,20)
	$NameLabel.Text = "Name:"
	$Form6.Controls.Add($NameLabel)

	$NameLabel = New-Object System.Windows.Forms.Label
	$NameLabel.Location = New-Object System.Drawing.Size (150,100)
	$NameLabel.Size = New-Object System.Drawing.Size (350,20)
	$NameLabel.Font = New-Object System.Drawing.Font ("ArialNarrow",8)
	$NameLabel.Text = "$UserFirst $UserLast"
	$Form6.Controls.Add($NameLabel)

	# User Logon Name Label
	$LogonLabel = New-Object System.Windows.Forms.Label
	$LogonLabel.Location = New-Object System.Drawing.Size (20,130)
	$LogonLabel.Size = New-Object System.Drawing.Size (120,20)
	$LogonLabel.Font = New-Object System.Drawing.Font ("ArialNarrow",8)
	$LogonLabel.Text = "Alias:"
	$Form6.Controls.Add($LogonLabel)

	$LogonLabel = New-Object System.Windows.Forms.Label
	$LogonLabel.Location = New-Object System.Drawing.Size (150,130)
	$LogonLabel.Size = New-Object System.Drawing.Size (350,20)
	$LogonLabel.Text = "$UserAlias$UserDomain"
	$Form6.Controls.Add($LogonLabel)

	# Email Label
	$EmailLabel = New-Object System.Windows.Forms.Label
	$EmailLabel.Location = New-Object System.Drawing.Size (20,160)
	$EmailLabel.Size = New-Object System.Drawing.Size (130,20)
	$EmailLabel.Text = "E-mail:"
	$Form6.Controls.Add($EmailLabel)

	$EmailLabel = New-Object System.Windows.Forms.Label
	$EmailLabel.Location = New-Object System.Drawing.Size (150,160)
	$EmailLabel.Size = New-Object System.Drawing.Size (350,20)
	$EmailLabel.Font = New-Object System.Drawing.Font ("ArialNarrow",8)
	$EmailLabel.Text = "$UserEmail$UserEmailDomain"
	$Form6.Controls.Add($EmailLabel)

	# Distribution Group Label
	$DGLabel = New-Object System.Windows.Forms.Label
	$DGLabel.Location = New-Object System.Drawing.Size (20,190)
	$DGLabel.Size = New-Object System.Drawing.Size (120,20)
	$DGLabel.Text = "Distribution Group:"
	$Form6.Controls.Add($DGLabel)

	$DGLabel = New-Object System.Windows.Forms.Label
	$DGLabel.Location = New-Object System.Drawing.Size (150,190)
	$DGLabel.Size = New-Object System.Drawing.Size (350,20)
	$DGLabel.Text = "$UserDG"
	$Form6.Controls.Add($DGLabel)

	# Office Label
	$OfficeLabel = New-Object System.Windows.Forms.Label
	$OfficeLabel.Location = New-Object System.Drawing.Size (20,220)
	$OfficeLabel.Size = New-Object System.Drawing.Size (120,20)
	$OfficeLabel.Text = "Office:"
	$Form6.Controls.Add($OfficeLabel)

	$OfficeLabel = New-Object System.Windows.Forms.Label
	$OfficeLabel.Location = New-Object System.Drawing.Size (150,220)
	$OfficeLabel.Size = New-Object System.Drawing.Size (350,20)
	$OfficeLabel.Text = "$UserOffice"
	$Form6.Controls.Add($OfficeLabel)

	# Phone Label
	$PhoneLabel = New-Object System.Windows.Forms.Label
	$PhoneLabel.Location = New-Object System.Drawing.Size (20,250)
	$PhoneLabel.Size = New-Object System.Drawing.Size (120,20)
	$PhoneLabel.Text = "Phone:"
	$Form6.Controls.Add($PhoneLabel)

	$PhoneLabel = New-Object System.Windows.Forms.Label
	$PhoneLabel.Location = New-Object System.Drawing.Size (150,250)
	$PhoneLabel.Size = New-Object System.Drawing.Size (350,20)
	$PhoneLabel.Text = "$UserPhone $UserExtension".Trim()
	$Form6.Controls.Add($PhoneLabel)

	$Form6.Topmost = $True
	$Form6.Add_Shown({ $Form6.Activate() })
	[void]$Form6.ShowDialog()
}

function ShowFormMembership
{

	$launchForm = New-Object System.Windows.Forms.Form
	$launchForm.Text = "Test"
	$launchForm.Size = New-Object System.Drawing.Size (1000,950)
	$launchForm.StartPosition = "CenterScreen"
	$launchForm.Topmost = $True

	# The Name label
	$CompanyLabel = New-Object System.Windows.Forms.Label
	$CompanyLabel.Location = New-Object System.Drawing.Size (95,63)
	$CompanyLabel.Size = New-Object System.Drawing.Size (100,20)
	$CompanyLabel.Text = "Company/Office: "
	$launchForm.Controls.Add($CompanyLabel)

	# Need two more section

	# The Name label
	$NameLabel = New-Object System.Windows.Forms.Label
	$NameLabel.Location = New-Object System.Drawing.Size (95,83)
	$NameLabel.Size = New-Object System.Drawing.Size (40,20)
	$NameLabel.Text = "Name: "
	$launchForm.Controls.Add($NameLabel)

	# The form label
	$DisableNameLabel = New-Object System.Windows.Forms.Label
	$DisableNameLabel.Location = New-Object System.Drawing.Size (95,213)
	$DisableNameLabel.Size = New-Object System.Drawing.Size (40,20)
	$DisableNameLabel.Text = "Alias: "
	$launchForm.Controls.Add($DisableNameLabel)

	# Logon Name Text Box
	$DisableName = New-Object System.Windows.Forms.TextBox
	$DisableName.Location = New-Object System.Drawing.Size (135,210)
	$DisableName.Size = New-Object System.Drawing.Size (140,20)
	$DisableName.AutoCompleteSource = 'CustomSource'
	$DisableName.AutoCompleteMode = 'SuggestAppend'
	$DisableName.AutoCompleteCustomSource = $autocomplete
	Get-User | ForEach-Object { $DisableName.AutoCompleteCustomSource.Add($_.SamAccountName) }
	$launchForm.Controls.Add($DisableName)

	# The Detail Information table
	$dataGridView1 = New-Object System.Windows.Forms.DataGridView
	$dataGridView1.Location = New-Object System.Drawing.Size (95,240)
	$dataGridView1.Size = New-Object System.Drawing.Size (800,400)
	$dataGridView1.DataBindings.DefaultDataSourceUpdateMode = 0
	$dataGridView1.Name = "dataGrid1"
	$dataGridView1.DataMember = ""
	$dataGridView1.ReadOnly = $True
	$dataGridView1.TabIndex = 3
	$dataGridView1.AllowUserToAddRows = $False
	$launchForm.Controls.Add($dataGridView1)

	# Error Message Box
	$ErrorMsg = New-Object System.Windows.Forms.Label
	$ErrorMsg.Location = New-Object System.Drawing.Size (20,160)
	$ErrorMsg.Size = New-Object System.Drawing.Size (500,20)
	$ErrorMsg.Text = $null
	$ErrorMsg.ForeColor = "Red"
	$launchForm.Controls.Add($ErrorMsg)

	# The data grid for the user information
	$dataGridView1.ColumnCount = 4
	$dataGridView1.ColumnHeadersVisible = $true
	$dataGridView1.Columns[0].Name = "Name"
	$dataGridView1.Columns[1].Name = "Organization Unit"
	$dataGridView1.Columns[2].Name = "Alias Name"
	$dataGridView1.Columns[3].Name = "Email Address"
	$dataGridView1.Columns[0].Width = "200"
	$dataGridView1.Columns[1].Width = "200"
	$dataGridView1.Columns[2].Width = "200"
	$dataGridView1.Columns[3].Width = "200"

	# Add the user Account information in to the detail table
	$AddButton = New-Object System.Windows.Forms.Button
	$AddButton.Location = New-Object System.Drawing.Size (315,210)
	$AddButton.Size = New-Object System.Drawing.Size (75,23)
	$AddButton.Text = "Add"
	$AddButton.TabIndex = 1
	$AddButton.Add_Click({

			# Check exsisting user in the table
			function Check
			{
				$checkname = $Disablename.Text
				$checkrownumber = $dataGridView1.RowCount.ToString()
				for ($j = 1; $j -le $checkrownumber; $j++)
				{
					$existname = $dataGridView1.Rows[$j - 1].Cells['Alias Name'].Value
					if ($checkname -eq $existname)
					{
						return $True
					}
				}
				return $False
			}

			# The control statement for the account add
			if ($Disablename.Text -eq "")
			{
				$ErrorMsg.Text = "Please enter the information!"

			}
			elseif (Check)
			{
				$Disablename.Clear()
				$ErrorMsg.Text = "Account already exsit in the table"
			}
			else {
				$tmpname = $Disablename.Text
				$tmpMailbox = Get-Mailbox -Filter "Alias -eq '$tmpname'"

				if ($tmpMailbox) {
					Get-Mailbox -Identity $Disablename.Text | ForEach-Object {

						$dataGridView1.Rows.Add($_.Name,$_.OrganizationalUnit,$_.Alias,$_.PrimarySmtpAddress) | Out-Null

					}
				}
				else {
					$ErrorMsg.Text = "No Record Found!"
				}

				$DisableName.Clear()
			}

			# Check for the button visiblity
			$rownumber = $dataGridView1.RowCount.ToString()
			if ($rownumber -ge 1) {
				$RemoveButton.Enabled = $True
			} else {
				$RemoveButton.Enabled = $False
			}

			$launchForm.Refresh()
		})
	$launchForm.Controls.Add($AddButton)

	# Add the function to remove the user account
	$RemoveButton = New-Object System.Windows.Forms.Button
	$RemoveButton.Location = New-Object System.Drawing.Size (405,210)
	$RemoveButton.Size = New-Object System.Drawing.Size (75,23)
	$RemoveButton.Text = "Remove"
	$RemoveButton.TabIndex = 2
	$RemoveButton.Enabled = $False

	# Check for the button visiblity
	# Move the user from the datagridview

	$RemoveButton.Add_Click({

			if ($dataGridView1.SelectedRows.Count -eq $dataGridView1.Rows.Count)
			{
				$dataGridView1.Rows.Clear();
			}
			foreach ($row in $dataGridView1.SelectedRows)
			{
				$dataGridView1.Rows.Remove($row);
			}

			$rownumber = $dataGridView1.RowCount.ToString()
			if ($rownumber -ge 1) {

				$RemoveButton.Enabled = $True


			} else {
				$RemoveButton.Enabled = $False

			}

			$launchForm.Refresh()
		})

	$launchForm.Controls.Add($RemoveButton)


	$BrowseButton = New-Object System.Windows.Forms.Button
	$BrowseButton.Location = New-Object System.Drawing.Size (490,210)
	$BrowseButton.Size = New-Object System.Drawing.Size (75,23)
	$BrowseButton.Text = "Browse"
	$BrowseButton.Add_Click({

			$myDialog = New-Object System.Windows.Forms.OpenFileDialog
			$myDialog.Title = "Please select a file"
			$myDialog.Filter = "All Files (*.*)|*.*"
			$result = $myDialog.ShowDialog()
			if ($result -eq "OK") {

				if ($myDialog.FileName -like "*.csv*") {

					$Disablename.Text = $myDialog.FileName

				} else {

					$ErrorMsg.Text = "Please load the csv format list"

				}

				# Continue working with file
			} else {

				Write-Host ?Cancelled by user?

			}
		})

	$launchForm.Controls.Add($BrowseButton)

	$ImportButton = New-Object System.Windows.Forms.Button
	$ImportButton.Location = New-Object System.Drawing.Size (575,210)
	$ImportButton.Size = New-Object System.Drawing.Size (75,23)
	$ImportButton.Text = "Import"

	$ImportButton.Add_Click({

			$Filepath = $Disablename.Text
			Import-Csv $Filepath | ForEach-Object {
				$UserListName = $_."Username"
				$CheckUser = Get-Mailbox -Filter "DisplayName -eq '$UserListName'"

				if ($CheckUser)
				{
					$dataGridView1.Rows.Add($CheckUser.Name,$CheckUser.OrganizationalUnit,$CheckUser.Alias,$CheckUser.PrimarySmtpAddress) | Out-Null
				}

				else
				{
					$dataGridView1.Rows.Add($UserListName,"not found") | Out-Null
				}
			}

			$Disablename.Text = $null
			$rownumber = $dataGridView1.RowCount.ToString()
			if ($rownumber -ge 1) {
				$RemoveButton.Enabled = $True
			} else {
				$RemoveButton.Enabled = $False
			}

			$launchForm.Refresh()
		})

	#Move User to Selected Group
	$launchForm.Controls.Add($ImportButton)

	$label1 = New-Object System.Windows.Forms.Label
	$label1.Location = New-Object System.Drawing.Size (95,655)
	$label1.Size = New-Object System.Drawing.Size (85,20)
	$label1.Text = "Selected OU: "
	$label1.Font = New-Object System.Drawing.Font ("Arial",8,([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)),[System.Drawing.GraphicsUnit]::Point,([System.Byte](0)))
	$label1.TabStop = $False
	$launchForm.Controls.Add($label1)

	$label2 = New-Object System.Windows.Forms.Label
	$label2.Location = New-Object System.Drawing.Size (95,655)
	$label2.Size = New-Object System.Drawing.Size (85,20)
	$label2.Text = "Group: "
	$label2.Font = New-Object System.Drawing.Font ("Arial",8,([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)),[System.Drawing.GraphicsUnit]::Point,([System.Byte](0)))
	$label2.TabStop = $False
	$launchForm.Controls.Add($label2)

	$Group = New-Object System.Windows.Forms.ComboBox
	$Group.Location = New-Object System.Drawing.Size (185,650)
	$Group.Size = New-Object System.Drawing.Size (140,20)
	[void]$Group.Items.Add("DFS")
	[void]$Group.Items.Add("File Share")
	$Group.Add_SelectedValueChanged({
			$MoveButton.Enabled = $True
			$MoveAllButton.Enabled = $True
			$objListbox.Items.Clear()
			$searcher_G = New-Object system.directoryservices.directorysearcher;
			$grp_G = New-Object system.directoryservices.directoryentry;
			$Groupname = $Group.SelectedItem.ToString()
			$PathDirection1 = "OU=$Groupname,OU=Groups,DC=malabs,DC=com"
			$root_G = [adsi]("LDAP://" + $PathDirection1)
			$searcher_G.SearchRoot = $root_G
			$searcher_G.SearchScope = "Onelevel"
			$searcher_G.FindAll() | ForEach-Object {
				$SecurityGroupName = $_.properties.Name
				[void]$objListbox.Items.Add("$SecurityGroupName")
			}

		})
	$launchForm.Controls.Add($Group)

	$objListbox = New-Object System.Windows.Forms.Listbox
	$objListbox.Location = New-Object System.Drawing.Size (185,675)
	$objListbox.Size = New-Object System.Drawing.Size (100,20)
	$objListbox.SelectionMode = "MultiExtended"
	$objListbox.Height = 200
	$objListbox.Width = 140
	$objListbox.HorizontalScrollbar = $True
	$launchForm.Controls.Add($objListbox)

	function Check-IsGroupMember {
		param($user,$grp)
		$strFilter = "(&(objectClass=Group)(name=" + $grp + "))"
		$objDomain = New-Object System.DirectoryServices.DirectoryEntry
		$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
		$objSearcher.SearchRoot = $objDomain
		$objSearcher.PageSize = 1000
		$objSearcher.Filter = $strFilter
		$objSearcher.SearchScope = "Subtree"
		$colResults = $objSearcher.FindOne()
		$objItem = $colResults.properties
		([string]$objItem.member).contains($user)
	}

	$MoveButton = New-Object System.Windows.Forms.Button
	$MoveButton.Location = New-Object System.Drawing.Size (600,650)
	$MoveButton.Size = New-Object System.Drawing.Size (75,23)
	$MoveButton.Text = "Move"
	$MoveButton.Enabled = $False

	$MoveButton.Add_Click({
			#collect the group name from the objlistbox
			$x = @()
			foreach ($objItem in $objListbox.SelectedItems)
			{
				$x += $objItem
			}

			ShowFormMembership2

		})

	$launchForm.Controls.Add($MoveButton)

	$MoveAllButton = New-Object System.Windows.Forms.Button
	$MoveAllButton.Location = New-Object System.Drawing.Size (680,650)
	$MoveAllButton.Size = New-Object System.Drawing.Size (75,23)
	$MoveAllButton.Text = "Move All"
	$MoveAllButton.Enabled = $False

	$MoveAllButton.Add_Click({
			#collect the group name from the objlistbox
			$x = @()
			foreach ($objItem in $objListbox.SelectedItems)
			{
				$x += $objItem
			}

			ShowFormMembership2

		})

	$launchForm.Controls.Add($MoveAllButton)

	$CancelButton2 = New-Object System.Windows.Forms.Button
	$CancelButton2.Location = New-Object System.Drawing.Size (765,650)
	$CancelButton2.Size = New-Object System.Drawing.Size (75,23)
	$CancelButton2.Text = "Cancel"
	$CancelButton2.TabIndex = 3
	$CancelButton2.Add_Click({ $launchForm.close(); $Form7.close() })
	$launchForm.Controls.Add($CancelButton2)

	#The disable buttons for disable the users on the list

	$DisableButton = New-Object System.Windows.Forms.Button
	$DisableButton.Location = New-Object System.Drawing.Size (600,680)
	$DisableButton.Size = New-Object System.Drawing.Size (75,23)
	$DisableButton.Text = "Disable"
	$DisableButton.TabIndex = 4
	$DisableButton.Add_Click({

			$CheckDisable = $True
			$x = @()
			foreach ($objItem in $objListbox.SelectedItems)
			{
				$x += $objItem
			}

			ShowFormMembership2

		})

	$launchForm.Controls.Add($DisableButton)

	#The diasble all buttons for disable all the users who append in the datagridview 
	$DisableAllButton = New-Object System.windows.Forms.Button
	$DisableAllButton.Location = New-Object System.Drawing.Size (680,680)
	$DisableAllButton.Size = New-Object System.Drawing.Size (75,23)
	$DisableAllButton.Text = "Disable All"
	$DisableAllButton.TabIndex = 5

	$DisableAllButton.Add_Click({

		$CheckDisable = $True
		$x = @()
		foreach ($objItem in $objListbox.SelectedItems)
		{
			$x += $objItem
		}

		ShowFormMembership2

	})

	$launchForm.Controls.Add($DisableAllButton)

	$launchForm.Add_Shown({

		$launchForm.Activate()

	})
	[void]$Form.ShowDialog()

}

#The process status window console
function ShowFormMembership1 {

	$launchForm.visible = $True
	$Form1 = New-Object System.Windows.Forms.Form
	$Form1.Text = "Process Status"
	$Form1.Size = New-Object System.Drawing.Size (900,410)
	$Form1.StartPosition = "CenterScreen"

	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size (400,330)
	$CancelButton.Size = New-Object System.Drawing.Size (75,23)
	$CancelButton.Text = "Cancel"
	$CancelButton.Add_Click({

		$Form1.close()

	})
	$Form1.Controls.Add($CancelButton)

	$FinishButton = New-Object System.Windows.Forms.Button
	$FinishButton.Location = New-Object System.Drawing.Size (320,330)
	$FinishButton.Size = New-Object System.Drawing.Size (75,23)
	$FinishButton.Text = "Finish"
	$FinishButton.visible = $False
	$FinishButton.Add_Click({

		$launchForm.Refresh()
		Start-Sleep -s 1
		$Form1.close()

	})

	$Form1.Controls.Add($FinishButton)

	$OKButton = New-Object System.Windows.Forms.Button
	$OKButton.Location = New-Object System.Drawing.Size (320,330)
	$OKButton.Size = New-Object System.Drawing.Size (75,23)
	$OKButton.Text = "Start"
	$OKButton.Add_Click({

			if ($CheckDisabled = $True) {


				Disable


			}

			$m = 0
			$ErrorMsg3 = New-Object System.Windows.Forms.Label
			$ErrorMsg3.Location = New-Object System.Drawing.Size (20,40)
			$ErrorMsg3.Size = New-Object System.Drawing.Size (500,20)
			$ErrorMsg3.Text = $null
			$ErrorMsg3.ForeColor = "Red"
			$Form1.Controls.Add($ErrorMsg3)

			$rownumber = $dataGridView1.RowCount.ToString()

			if ($rownumber -eq 0) {

				$ErrorMsg3.Text = "Please insert user into the list"
				$ErrorMsg3.ForeColor = "Red"

				Start-Sleep -s 1

			} else {

				$ErrorMsg3.Text = "Start moving the user..."
				$ErrorMsg3.ForeColor = "Green"

				Start-Sleep -s 1

				# The record number in the datagrid
				#Get the User's LDAP path
				$dataGridView1.SelectedRows | ForEach-Object {

					$RecordNum = $m + 1
					$p = $m % 10

					if ($p -eq 0) {

						$q += 1

					}

					$ErrorMsg2 = New-Object System.Windows.Forms.Label
					$ErrorMsg2.Location = New-Object System.Drawing.Size ((20 + ($q - 1) * 400),(60 + $p * 20))
					$ErrorMsg2.Size = New-Object System.Drawing.Size (400,20)
					$Form1.Controls.Add($ErrorMsg2)

					$user = $dataGridView1.Rows[$_.Index].Cells['Name'].Value
					$tmpUser = Get-User -Filter "Name -eq '$user'" -domaincontroller $DC
					if ($tmpUser) {

						$UserFullName = $tmpUser.Name
						$UserDSName = $tmpUser.distinguishedname
						$Member = $tmpUser.memberof
						$Userpath = "LDAP://" + $UserDSName
						$Member = $Userpath.memberof
						#Get the the path of the group
						$searcher_G = New-Object system.directoryservices.directorysearcher;
						$grp_G = New-Object system.directoryservices.directoryentry;
						for ($j = 0; $j -lt $x.Length; $j++) {

							$rootname = $Group.Text
							$SecurityGroupName = $x[$j]
							$root_G = [adsi]("LDAP:// OU=$rootname,OU=Groups,DC=malabs,DC=com")
							$searcher_G.SearchRoot = $root_G
							$result_G = $searcher_G.FindAll() | Where-Object { $_.properties.Name -eq "$SecurityGroupName" }
							$grp_g = $result_g.GetDirectoryEntry()
							$checkUser = Check-IsGroupMember $UserFullName $SecurityGroupName

							if ($checkUser) {

								Start-Sleep -s 1

								$ErrorMsg2.Text = "$RecordNum. $UserFullName already exsist in assigned group"
								$ErrorMsg2.ForeColor = "Red"

								$OKButton.visible = $False
								$FinishButton.visible = $True
								$Form1.Refresh()

							} else {

								#add user to that Group
								$grp_g.psbase.Invoke("Add",$Userpath);
								$grp_g.psbase.CommitChanges();

								Start-Sleep -s 1

								$ErrorMsg2.Text = "$RecordNum. $UserFullName is moved successfully!"
								$ErrorMsg2.ForeColor = "Green"

							}

						}
						#Loop the group list stop
					} else {

						$ErrorMsg2.Text = "$RecordNum. $User moved failed"
						$ErrorMsg2.ForeColor = "Red"

						Start-Sleep -s 1

					}

					$m += 1

				}

			}

			Start-Sleep -s 1

			$OKButton.visible = $False
			$FinishButton.visible = $True
			$ErrorMsg3.Text = "$RecordNum Job is done"
			$Form1.Refresh()

		})
	$Form1.Controls.Add($OKButton)


	# Office Label
	$Label = New-Object System.Windows.Forms.Label
	$Label.Location = New-Object System.Drawing.Size (20, 20)
	$Label.Size = New-Object System.Drawing.Size (120, 20)
	$Label.Text = "Status:"
	$Form1.Controls.Add($Label)

	$Form1.Topmost = $True
	$Form1.Add_Shown({

		$Form1.Activate()

	})
	[void]$Form1.ShowDialog()

}

function ShowFormMembership2 {

	$launchForm.visible = $True
	$Form2 = New-Object System.Windows.Forms.Form
	$Form2.Text = "Process Status"
	$Form2.Size = New-Object System.Drawing.Size (900, 410)
	$Form2.StartPosition = "CenterScreen"

	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size (400, 330)
	$CancelButton.Size = New-Object System.Drawing.Size (75, 23)
	$CancelButton.Text = "Cancel"
	$CancelButton.Add_Click({ $Form2.close() })
	$Form2.Controls.Add($CancelButton)

	$FinishButton = New-Object System.Windows.Forms.Button
	$FinishButton.Location = New-Object System.Drawing.Size (320, 330)
	$FinishButton.Size = New-Object System.Drawing.Size (75, 23)
	$FinishButton.Text = "Finish"
	$FinishButton.visible = $False
	$FinishButton.Add_Click({
			$launchForm.Refresh()
			Start-Sleep -s 1
			$Form2.close() })
	$Form2.Controls.Add($FinishButton)

	$OKButton = New-Object System.Windows.Forms.Button
	$OKButton.Location = New-Object System.Drawing.Size (320, 330)
	$OKButton.Size = New-Object System.Drawing.Size (75, 23)
	$OKButton.Text = "Start"
	$OKButton.Add_Click({

		$m = 0
		$q = 0
		$ErrorMsg3 = New-Object System.Windows.Forms.Label
		$ErrorMsg3.Location = New-Object System.Drawing.Size (20, 40)
		$ErrorMsg3.Size = New-Object System.Drawing.Size (500, 20)
		$ErrorMsg3.Text = $null
		$ErrorMsg3.ForeColor = "Red"
		$Form2.Controls.Add($ErrorMsg3)

		# The record number in the datagrid
		$rownumber = $dataGridView1.RowCount.ToString()

		if ($rownumber -eq 0) {

			$ErrorMsg3.Text = "Please insert user into the list"
			$ErrorMsg3.ForeColor = "Red"

			Start-Sleep -s 1

		} else {

			$ErrorMsg3.Text = "Start moving users..."
			$ErrorMsg3.ForeColor = "Green"

			Start-Sleep -s 1

			for ($i = 0; $i -lt $rownumber; $i++) {

				$recordnumber = $i + 1
				$n = $i % 10

				if ($n -eq 0) {

					$m += 1

				}

				$ErrorMsg2 = New-Object System.Windows.Forms.Label
				$ErrorMsg2.Location = New-Object System.Drawing.Size ((20 + 400 * ($m - 1)),(60 + $n * 20))
				$ErrorMsg2.Size = New-Object System.Drawing.Size (400, 20)
				$Form2.Controls.Add($ErrorMsg2)
				# The record number in the datagrid
				#Get the User's LDAP path
				$user = $dataGridView1.Rows[$i].Cells['Name'].Value
				$tmpUser = Get-User -Filter "Name -eq '$user'" -domaincontroller $DC

				if ($tmpUser) {

					$UserDSName = $tmpUser.distinguishedname
					$UserFullName = $tmpUser.Name
					$Userpath = "LDAP://" + $UserDSName
					#Get the the path of the group
					$searcher_G = New-Object system.directoryservices.directorysearcher;
					$grp_G = New-Object system.directoryservices.directoryentry;

					#Loop the all the group in a list
					for ($j = 0; $j -lt $x.Length; $j++) {

						$rootname = $Group.Text
						$SecurityGroupName = $x[$j]
						$PathDirection1 = "OU=$rootname,OU=Groups,DC=malabs,DC=com"
						$root_G = [adsi]("LDAP://" + $PathDirection1)
						$searcher_G.SearchRoot = $root_G
						$result_G = $searcher_G.FindAll() | Where-Object { $_.properties.Name -eq "$SecurityGroupName" }
						$grp_g = $result_g.GetDirectoryEntry()
						$checkUser = Check-IsGroupMember $UserFullName $SecurityGroupName
						#check the exist user

						if ($checkUser) {

							$ErrorMsg2.Text = "$recordnumber.$UserFullName already exsist in assigned group"
							$ErrorMsg2.ForeColor = "Red"

							Start-Sleep -s 1

						} else {

							#add user to that Group
							$grp_g.psbase.Invoke("Add",$Userpath);
							$grp_g.psbase.CommitChanges();

							$ErrorMsg2.Text = "$UserFullName is moved successfully!"
							$ErrorMsg2.ForeColor = "Green"

							Start-Sleep -s 1

						}

					}

					#Group List loop stop
				} else {

					$ErrorMsg2.Text = "$recordnumber.$user moved failed"
					$ErrorMsg2.ForeColor = "Red"

					Start-Sleep -s 1

				}
			}
		}

		Start-Sleep -s 1

		$OKButton.visible = $False
		$FinishButton.visible = $True

		$ErrorMsg3.Text = "$recordnumber Job is done"

		$Form2.Refresh()

	})

	$Form2.Controls.Add($OKButton)

	# Office Label
	$Label = New-Object System.Windows.Forms.Label
	$Label.Location = New-Object System.Drawing.Size (20, 20)
	$Label.Size = New-Object System.Drawing.Size (120, 20)
	$Label.Text = "Status:"
	$Form2.Controls.Add($Label)

	$Form2.Topmost = $True
	$Form2.Add_Shown({

		$Form2.Activate()

	})
	[void]$Form2.ShowDialog()

}

Write-Host "This window will close when the Powershell script exits."
$launchForm = Show-LaunchForm