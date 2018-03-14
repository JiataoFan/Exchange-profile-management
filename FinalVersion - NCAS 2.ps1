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

	############## $new ?
	if ($new -and $UserName.Text -ne "" -and $OU.Text -ne "" -and $MainCompany -ne "" -and $Office -ne "") {

		if (Validate-Fullname) {

			$LogonName.Enabled = $True
			#$Domain.Enabled = $True
			$Password.Enabled = $True
			$Server.Enabled = $True


			$MB.Enabled = $True
			$Email.Enabled = $True
			#$EmailDomain.Enabled = $True
			$DG.Enabled = $True
			#$Phone.Enabled = $True
			$Extension.Enabled = $True
			$OWA.Enabled = $True
			$ActiveSync.Enabled = $True
			$objListbox_DFS.Enabled = $True
			$objListbox_fileshare.Enabled = $True
			$objListbox_WebsenseDept.Enabled = $True
			$objListbox_WebsenseList.Enabled = $True

			$UserName.Text = $UserFirst.substring(0,1).ToUpper() + $UserFirst.substring(1).ToLower() + " " + $UserLast.substring(0,1).ToUpper() + $UserLast.substring(1).ToLower()
			#$Lastname.Text = $UserLast.substring(0, 1).toupper() + $UserLast.substring(1).tolower()

			$DG.Items.Clear()
			$objListbox_DFS.Items.Clear()
			$objListbox_fileshare.Items.Clear()
			$objListbox_WebsenseDept.Items.Clear()
			$objListbox_WebsenseList.Items.Clear()
			$MainCompanyName = $MainCompany.SelectedItem.ToString()
			$Officename = $Office.SelectedItem.ToString()
			$DepartmentName = $Department.SelectedItem.ToString()

			##################pmgroup? 
			if ($Pmgroup.visible -eq $false) {

			} else {

				if ($Pmgroup.SelectedItem -eq $null) {

					$PmgroupName = ''

				} else {

					$PmgroupName = $Pmgroup.SelectedItem.ToString()

				}

			}

			#Network drive adding

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

						[void]$objListbox_DFS.Items.Add("$SecurityGroupName")

					}

					for ($i = 0; $i -lt $objListbox_DFS.Items.Count; $i++) {

						$objListbox_DFS.SetItemChecked($i,$true)

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

						[void]$objListbox_fileshare.Items.Add("$SecurityGroupName")

					}

					for ($i = 0; $i -lt $objListbox_fileshare.Items.Count; $i++) {

						$objListbox_fileshare.SetItemChecked($i,$true)

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

			$searcher_G = New-Object system.directoryservices.directorysearcher;
			$grp_G = New-Object system.directoryservices.directoryentry;
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

						[void]$objListbox_WebsenseDept.Items.Add("$SecurityGroupName")

					}

					for ($i = 0; $i -lt $objListbox_WebsenseDept.Items.Count; $i++) {

						$objListbox_WebsenseDept.SetItemChecked($i,$true)

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

						[void]$objListbox_WebsenseList.Items.Add("$SecurityGroupName")

					}

					for ($i = 0; $i -lt $objListbox_WebsenseList.Items.Count; $i++) {

						$objListbox_WebsenseList.SetItemChecked($i,$true)

					}

				}

			}

			$DG.Items.Clear()

			Get-DistributionGroup -Filter "(CustomAttribute1 -eq '$MainCompanyName' -and CustomAttribute4 -eq '$Officename' -and CustomAttribute5 -like '*Office*')
			-or (CustomAttribute2 -eq '$MainCompanyName' -and CustomAttribute4 -eq '$Officename' -and CustomAttribute5 -like '*Office*')
			-or (CustomAttribute3 -eq '$MainCompanyName' -and CustomAttribute4 -eq '$Officename' -and CustomAttribute5 -like '*Office*')
			-or (CustomAttribute1 -eq '$MainCompanyName' -and CustomAttribute4 -like '*$Officename*' -and CustomAttribute5 -like '*<$DepartmentName>*')
			-or (CustomAttribute2 -eq '$MainCompanyName' -and CustomAttribute4 -like '*$Officename*' -and CustomAttribute5 -like '*<$DepartmentName>*')
			-or (CustomAttribute3 -eq '$MainCompanyName' -and CustomAttribute4 -like '*$Officename*' -and CustomAttribute5 -like '*<$DepartmentName>*')" | ForEach-Object {

				[void]$DG.Items.Add($_.Name)

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

			for ($i = 0; $i -lt $DG.Items.Count; $i++) {

				$DG.SetItemChecked($i, $true)

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

			$OUMsg.Text = "*";
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
	$BackButton2.Enabled = $False
	$CancelButton2.Enabled = $False
	$OKButton2.Enabled = $False

	# 1 of 10, creating mailbox
	$ErrorMsg2.Text = "[1/10] Creating mailbox `"$UserAlias`"...please wait."
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

		$ErrorMsg3.Text = "[2/10] Setting $customattr attribute to `"$UserCompany`"."
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
		$ErrorMsg4.Text = "[3/10] Setting offline address book to `"$UserCompany OAB`"."
		$Form2.Refresh()
		Set-Mailbox $userAlias -offlineaddressbook "$userCompany OAB" -domaincontroller $DC
		"" >> $log
		"3. Setting offline address book to `"$UserCompany OAB`"." >> $log
		Start-Sleep -s 1

		# 4 of 10, disabling mailbox features
		$ErrorMsg5.Text = "[4/10] Disabling mailbox features."
		$Form2.Refresh()
		$BlockOutlookAnywhere = !$UserOutlookAnywhere
		Set-CASMailbox -Identity "$NTDomain\$UserAlias" -ActiveSyncEnabled $UserActiveSync -OWAEnabled $UserOWA -MAPIBlockOutlookRpcHttp $BlockOutlookAnywhere -domaincontroller $DC
		"" >> $log
		"4. Disabling mailbox features." >> $log
		Start-Sleep -s 1

		# 5 of 10, setting Organization and Phone
		$ErrorMsg6.Text = "[5/10] Setting Organization and Phone."
		$Form2.Refresh()
		Set-User $UserAlias -Company "$UserCompany" -Phone ("$UserPhone $UserExtension".Trim()) -Office "$UserOffice" -domaincontroller $DC
		"" >> $log
		"5. Setting Organization and Phone." >> $log
		Start-Sleep -s 1

		# 6 of 10, adding user to company
		if (IsMember "$UserCompany Company" $UserAlias) {

			$ErrorMsg7.ForeColor = "Red"
			$ErrorMsg7.Text = "[6/10] WARNING: User already a member of `"$UserCompany Company`"."
			"" >> $log
			"6. WARNING: User already a member of `"$UserCompany Company`"." >> $log

		} else {

			$ErrorMsg7.Text = "[6/10] Adding user to `"$UserCompany Company`"."
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

					$ErrorMsg17.ForeColor = "Red"
					$ErrorMsg17.Text = "[7/10] WARNING: User already a member of distribution group `"$UserDG`"."
					"" >> $log
					"7. WARNING: User already a member of distribution group `"$UserDG`"." >> $log
					cotinue

				} else {

					$SuccessDG += $UserDG
					$ErrorMsg8.Text = "[7/10] Adding user to distribution group `"$SuccessDG`"."
					Add-DistributionGroupMember "$UserDG" -member $UserAlias -domaincontroller $DC -BypassSecurityGroupManagerCheck
					"" >> $log
					"7. Adding user to distribution group `"$UserDG`"." >> $log
					Start-Sleep -s 1

				}

			}

		} else {

			$ErrorMsg8.Text = "[7/10] Skipped...no distribution group specified."
			"" >> $log
			"7. Skipped...no distribution group specified." >> $log

		}

		$Form2.Refresh()
		Start-Sleep -s 1

		# 8 of 10, updating Address List
		$ErrorMsg9.ForeColor = "Green"
		$ErrorMsg9.Text = "[8/10] Updating address list `"$UserCompany Address List`"."
		$Form2.Refresh()
		Update-AddressList "$UserCompany Address List" -domaincontroller $DC
		"" >> $log
		"8. Updating address list `"$UserCompany Address List`"." >> $log
		Start-Sleep -s 1

		# 9 of 10, updating Offline Address Book
		$ErrorMsg10.Text = "[9/10] Updating offline address book `"$UserCompany OAB`"."
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

		$ErrorMsg11.Text = "[10/10] Setting msExchQueryBaseDN attribute."
		$Form2.Refresh()
		#$user.MsExchSeachBase = "CN=$UserCompany Address List,CN=All Address Lists,CN=Address Lists Container,CN=MA Labs,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=MA Labs,DC=local"
		"" >> $log
		"10. Setting msExchQueryBaseDN attribute." >> $log
		Start-Sleep -s 1
		[void]$user.CommitChanges()
		$Pass = $Password.Text
		$emailToBranch = $EmailCheckLabel3.Text
		Write-Host $emailToBranch
		$OKButton2.visible = $False
		$FinishButton2.TabIndex = 0
		$FinishButton2.visible = $True
		$ErrorMsg12.Text = "Mailbox `"$UserFirst $UserLast ($UserAlias)`" created sucessfully. Check log for errors."
		"" >> $log
		"Mailbox `"$UserFirst $UserLast ($UserAlias)`" created sucessfully." >> $log
		Start-Sleep -s 1

	} else {

		$ErrorMsg12.Text = "ERROR: Unable to create mailbox `"$UserFirst $UserLast ($UserAlias)`". Check the command syntax."
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
		$ErrorMsg7.ForeColor = "Green"
		$ErrorMsg7.Text = "[6/7] Updating address list `"$UserCompany Address List`"."
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
	$BackButton2.Enabled = $False
	$CancelButton2.Enabled = $False
	$OKButton2.Enabled = $False

	# 1 of 8, resetting custom attributes
	if ($mailenabled) {

		$ErrorMsg2.Text = "[1/8] Resetting custom attributes and hidden from Exchange Address lists."
		$Form2.Refresh()
		Set-Mailbox $userAlias -CustomAttribute1 "" -CustomAttribute2 "" -CustomAttribute3 "" -HiddenFromAddressListsEnabled $True -domaincontroller $DC
		"1. Resetting custom attributes and hidden from Exchange Address lists." > $log

	} else {

		$ErrorMsg2.Text = "[1/8] Skipped."
		$Form2.Refresh()
		"1. Skipped resetting custom attributes and hidden from Exchange Address lists." > $log

	}

	Start-Sleep -s 1

	# 2 of 8, disabling mailbox features
	if ($mailenabled) {

		$ErrorMsg3.Text = "[2/8] Disabling mailbox features."
		$Form2.Refresh()
		Set-CASMailbox -Identity $UserAlias -ActiveSyncEnabled $False -OWAEnabled $False -MAPIBlockOutlookRpcHttp $True -Confirm:$False -domaincontroller $DC
		"" >> $log
		"2. Disabling mailbox features." >> $log

	} else {

		$ErrorMsg3.Text = "[2/8] Skipped."
		$Form2.Refresh()
		"" >> $log
		"2. Skipped disabling mailbox features." >> $log

	}

	Start-Sleep -s 1

	# 3 of 8, removing group membership
	$ErrorMsg4.Text = "[3/8] Removing group membership."
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

		$ErrorMsg5.Text = "[4/8] Updating address list `"$Company Address List`"."
		$Form2.Refresh()
		Update-AddressList "$company Address List" -domaincontroller $DC
		"" >> $log
		"4. Updating address list `"$Company Address List`"." >> $log

	} else {

		$ErrorMsg5.Text = "[4/8] Skipped."
		$Form2.Refresh()
		"" >> $log
		"4. Skipped updating address list `"$Company Address List`"." >> $log

	}
	Start-Sleep -s 1

	# 5 of 8, updating Offline Address Book
	if ($mailenabled) {

		$ErrorMsg6.Text = "[5/8] Updating offline address book `"$company OAB`"."
		$Form2.Refresh()
		Update-OfflineAddressBook "$company OAB" -domaincontroller $DC
		"" >> $log
		"5. Updating offline address book `"$company OAB`"." >> $log

	} else {

		$ErrorMsg6.Text = "[5/8] Skipped."
		$Form2.Refresh()
		"" >> $log
		"5. Skipped updating offline address book `"$company OAB`"." >> $log

	}
	Start-Sleep -s 1

	# 6 of 8, moving user
	$ErrorMsg7.Text = "[6/8] Moving user to `"/Disabled/$company/Users`" OU."
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

				$ErrorMsg8.Text = "[7/8] Removing user home directory `"$userhome`"."
				$Form2.Refresh()
				Remove-Item $userhome -Force -Recurse
				"" >> $log
				"7. Removing user home directory `"$userhome`"." >> $log

			} else {

				$ErrorMsg8.Text = "[7/8] Archiving user home directory `"$userhome`"."
				$Form2.Refresh()
				Move-Item $userhome $archive
				"" >> $log
				"7. Archiving user home directory `"$userhome`"." >> $log

			}

		}

	} else {

		$ErrorMsg8.Text = "[7/8] Skipped."
		$Form2.Refresh()
		"" >> $logarchive
		"7. Skipped archiving user home directory." >> $log

	}
	Start-Sleep -s 1

	# 8 of 8, disabling user
	$ErrorMsg9.Text = "[8/8] Disabling user."
	$Form2.Refresh()
	$user.psbase.invokeset("AccountDisabled", "True")
	$user.setinfo()
	if ($user.psbase.invokeget("AccountDisabled") -eq $False) {

		$ErrorMsg9.Text = "ERROR: Unable to disable user `"$UserFirst $UserLast ($UserAlias)`". Check the command syntax."

	}
	"" >> $log
	"8. Disabling user." >> $log
	Start-Sleep -s 1

	$OKButton2.visible = $False
	$FinishButton2.TabIndex = 0
	$FinishButton2.TabStop = $True
	$FinishButton2.visible = $True
	$ErrorMsg10.Text = "User `"$UserFirst $UserLast ($UserAlias)`" disabled sucessfully. Check log for errors."
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

<##############################
	Tree view
################################>
function Add-Node {

	param($selectedNode, $name)

	$newNode = New-Object System.Windows.Forms.TreeNode
	$newNode.Name = $name
	$newNode.Text = $name
	$selectedNode.Nodes.Add($newNode) | Out-Null

	return $newNode

}

function Get-NextLevel {

	param($selectedNode, $dn)

	$path = [adsi]("LDAP://" + $dn)
	$OU = New-Object System.DirectoryServices.DirectorySearcher ($path)
	$OU.SearchScope = "onelevel"
	$OU.Filter = "(&(objectClass=organizationalUnit))"
	$OUs = $OU.FindAll()

	if ($OUs -eq $null) {

		$node = Add-Node $selectedNode $path

	} else {

		foreach ($person in $OUs) {

			[string]$ent = $person.properties.Name
			$node = Add-Node $selectedNode $ent
			[string]$dn = $person.properties.distinguishedname
			Get-NextLevel $node $dn

		}

	}

}

function Build-TreeView {

	if ($treeNodes) {

		$treeview1.Nodes.Remove($treeNodes)
		$form1.Refresh()

	}

	$treeNodes = New-Object System.Windows.Forms.TreeNode
	$treeNodes.Text = "Active Directory Hierarchy"
	$treeNodes.Name = "Active Directory Hierarchy"
	$treeNodes.Tag = "root"
	$treeView1.Nodes.Add($treeNodes) | Out-Null

	if ($new) {

		$treeView1.add_AfterSelect({

				[string]$fullpath = $this.SelectedNode.FullPath
				if ($fullpath -eq "Active Directory Hierarchy") {

					$ldappath = ""

				} else {

					[string]$ldappath = $fullpath.substring($fullpath.IndexOf("\"), $fullpath.Length - $fullpath.IndexOf("\"))
					$ldappath = $ldappath.Replace('\', '/')
					$textbox1.Text = $ldappath
					[void]$OU.Items.Add("$ldappath")
					$OU.Text = $ldappath
					$Office.Items.Clear()
					$MainCompany.SelectedIndex = 0

				}

			})

	}

	#Generate Module nodes 
	$OUs = Get-NextLevel $treeNodes $strDomainDN

	$treeNodes.Expand()

}

<##############################
	Form
################################>

<#
	Mailbox form
#>
function ShowForm1 {

	$launchForm.visible = $False

	$Form1 = New-Object System.Windows.Forms.Form
	$Form1.Text = $title
	$Form1.Size = New-Object System.Drawing.Size (900, 810)
	$Form1.WindowState = "Maximized"
	$Form1.StartPosition = "CenterScreen"

	$Form1.KeyPreview = $True

	<#
		"Enter" key behavior
	#>
	$Form1.Add_KeyDown({

		if ($_.KeyCode -eq "Enter") {

			$UserAlias = $LogonName.Text.Trim()
			if ($new -or $found) {

				$UserFirst = ToProperCase ($FirstName.Text.Trim())
				$UserLast = ToProperCase ($LastName.Text.Trim())
				$UserPassword = $Password.Text
				$UserOU = $OU.SelectedItem
				$MailboxServer = $Server.SelectedItem
				$DatabaseName = $MB.SelectedItem.Split(" ")
				$Database = $DatabaseName[0]
				$UserDomain = $Domain.SelectedItem
				$UserDG = $DG.SelectedItem
				$UserOffice = $Office.SelectedItem
				$UserPhone = $Phone.Text.Trim()
				$UserExtension = $Extension.Text.Trim()
				$UserEmail = $Email.Text.Trim()

				if (!$Email.Text) {

					if ($UserFirst -and $UserLast) {

						$UserEmail = $UserFirst + "." + $UserLast

					} else {

						$UserEmail = $UserFirst + $UserLast

					}

					$UserEmail = $UserEmail.Replace(" ", "")

				}

				$UserEmailDomain = $EmailDomain.SelectedItem
				$UserOutlookAnywhere = $OutlookAnywhere.Checked
				$UserOWA = $OWA.Checked
				$UserActiveSync = $ActiveSync.Checked

				if (Validate) {

					ShowForm2

				}

			} else {

				if (Validate) {

					$found = $True
					$tmpUser = Get-User -Filter "SamAccountName -eq '$UserAlias'" -domaincontroller $DC

					if ($disable) {

						$UserOU = ([string]$tmpUser.Identity.Parent).substring($Forest.Length)
						#$UserOU = ([string]$tmpUser.Identity).Substring($Forest.Length, ($tmpUser.Identity.Length - $tmpUser.Name.Length - $Forest.Length - 1))
						$UserFirst = $tmpUser.FirstName
						$UserLast = $tmpUser.LastName
						$UserDomain = $tmpUser.UserPrincipalName.substring($UserAlias.Length)
						$UserPassword = "********"
						$UserOffice = $tmpUser.Office
						$UserEmail = $tmpUser.WindowsEmailAddress.Local

						if ($tmpUser.WindowsEmailAddress.Domain) {

							$UserEmailDomain = "@" + $tmpUser.WindowsEmailAddress.Domain

						}

						if ($tmpUser.RecipientType -eq "UserMailbox") {

							$tmpMailbox = Get-Mailbox $UserAlias -domaincontroller $DC
							$MailboxServer = $tmpMailbox.Database.Parent.Parent.Parent.Name
							$Database = $tmpMailbox.Database.Name

						}

						if ($tmpUser.Phone) {

							$telephone = ($tmpUser.Phone).Split('x')

							switch ($telephone.Count) {

								1 {

									$UserPhone = $telephone[0].Trim()

								}
								
								2 {

									$UserPhone = $telephone[0].Trim();
									$UserExtension = $telephone[1].Trim()

								}

								default {}

							}

						}

					} else {

						$OU.SelectedItem = $([string]$tmpUser.Identity.Parent).substring($Forest.Length)
						$FirstName.Text = $tmpUser.FirstName
						$LastName.Text = $tmpUser.LastName
						$Domain.SelectedItem = $tmpUser.UserPrincipalName.substring($UserAlias.Length)
						$Password.Text = "********"
						$Office.SelectedItem = $tmpUser.Office
						$Email.Text = $tmpUser.WindowsEmailAddress.Local

						if ($tmpUser.WindowsEmailAddress.Domain) {

							$EmailDomain.Text = "@" + $tmpUser.WindowsEmailAddress.Domain

						}

						if ($tmpUser.Phone) {

							$telephone = ($tmpUser.Phone).Split('x')

							switch ($telephone.Count) {

								1 {

									$Phone.Text = $telephone[0].Trim()

								}

								2 {

									$Phone.Text = $telephone[0].Trim();
									$Extension.Text = $telephone[1].Trim()

								}

								default {}

							}

						}

						$Domain.Enabled = $True
						$Email.Enabled = $True
						$EmailDomain.Enabled = $True
						$Server.Enabled = $True
						$MB.Enabled = $True
						$DG.Enabled = $True
						$Office.Enabled = $True
						$Phone.Enabled = $True
						$Extension.Enabled = $True
						$OutlookAnywhere.Enabled = $True
						$OWA.Enabled = $True
						$ActiveSync.Enabled = $True

					}

					if ($disable) {

						ShowForm2

					} else {

						$Form1.Refresh()

					}

				}

			}

		}

	})

	<#
		"Escape" key behavior
	#>
	$Form1.Add_KeyDown({ 

		if ($_.KeyCode -eq "Escape") {

			$Form1.close()

		}

	})

	<#
		Config next button and behaviors
	#>
	$NextButton = New-Object System.Windows.Forms.Button
	$NextButton.Location = New-Object System.Drawing.Size (960, 770)
	$NextButton.Size = New-Object System.Drawing.Size (105, 43)
	$NextButton.Text = "Next >"
	$NextButton.TabIndex = 17
	$NextButton.Add_Click({

			$UserAlias = $LogonName.Text.Trim()

			if ($new -or $found) {

				$UserFirst = ToProperCase ($UserName.Text.Split()[0].Trim())
				$UserLast = ToProperCase ($UserName.Text.Split()[1].Trim())
				$UserPassword = $Password.Text
				$UserOU = $OU.Text
				$MailboxServer = $Server.SelectedItem
				$Database = $MB.SelectedItem.Split("==")[0].Trim()
				$UserDomain = $Domain.SelectedItem
				$UserOffice = $Office.SelectedItem
				$UserPhone = $Phone.Text.Trim()
				$UserExtension = $Extension.Text.Trim()
				$UserEmail = $Email.Text.Trim()
				#Group of DG names
				$GroupName_DG = @()
				foreach ($objItem in $DG.CheckedItems) {

					$GroupName_DG += $objItem

				}
				#The Data Store from the group membership selection list
				$GroupName_DFS = @()
				$GroupName_Vasto = @()
				$GroupName_Dept = @()
				$GroupName_List = @()
				foreach ($objItem in $objListbox_DFS.CheckedItems) {

					$GroupName_DFS += $objItem

				}

				foreach ($objItem in $objListbox_FileShare.CheckedItems) {

					$GroupName_Vasto += $objItem

				}

				foreach ($objItem in $objListbox_WebsenseDept.CheckedItems) {

					$GroupName_Dept += $objItem

				}

				foreach ($objItem in $objListbox_WebsenseList.CheckedItems) {

					$GroupName_List += $objItem

				}

				if (!$Email.Text) {

					if ($UserFirst -and $UserLast) {

						$UserEmail = $UserFirst + "." + $UserLast

					} else {

						$UserEmail = $UserFirst + $UserLast

					}

					$UserEmail = $UserEmail.Replace(" ","")

				}

				$UserEmailDomain = $EmailDomain.SelectedItem
				$UserOutlookAnywhere = $OutlookAnywhere.Checked
				$UserOWA = $OWA.Checked
				$UserActiveSync = $ActiveSync.Checked
				if (Validate) {

					ShowForm2

				}

			} else {

				if (Validate) {

					$found = $True
					$tmpUser = Get-User -Filter "SamAccountName -eq '$UserAlias'" -domaincontroller $DC

					if ($disable) {

						$UserOU = ([string]$tmpUser.Identity.Parent).substring($Forest.Length)
						#$UserOU = ([string]$tmpUser.Identity).Substring($Forest.Length,($tmpUser.Identity.Length - $tmpUser.Name.Length-$Forest.Length -1))
						$UserFirst = $tmpUser.FirstName
						$UserLast = $tmpUser.LastName
						$UserDomain = $tmpUser.UserPrincipalName.substring($UserAlias.Length)
						$UserPassword = "********"
						$UserOffice = $tmpUser.Office
						$UserEmail = $tmpUser.WindowsEmailAddress.Local
						if ($tmpUser.WindowsEmailAddress.Domain) {

							$UserEmailDomain = "@" + $tmpUser.WindowsEmailAddress.Domain

						}

						if ($tmpUser.RecipientType -eq "UserMailbox") {

							$tmpMailbox = Get-Mailbox $UserAlias -domaincontroller $DC
							$MailboxServer = $tmpMailbox.Database.Parent.Parent.Parent.Name
							$Database = $tmpMailbox.Database.Name

						}

						if ($tmpUser.Phone) {

							$telephone = ($tmpUser.Phone).Split('x')

							switch ($telephone.Count) {

								1 {

									$UserPhone = $telephone[0].Trim()

								}

								2 { 

									$UserPhone = $telephone[0].Trim(); 
									$UserExtension = $telephone[1].Trim()

								}

								default {}

							}

						}

					} else {

						$OU.SelectedItem = $([string]$tmpUser.Identity.Parent).substring($Forest.Length)
						$FirstName.Text = $tmpUser.FirstName
						$LastName.Text = $tmpUser.LastName
						$Domain.SelectedItem = $tmpUser.UserPrincipalName.substring($UserAlias.Length)
						$Password.Text = "********"
						$Office.SelectedItem = $tmpUser.Office
						$Email.Text = $tmpUser.WindowsEmailAddress.Local

						if ($tmpUser.WindowsEmailAddress.Domain) {

							$EmailDomain.Text = "@" + $tmpUser.WindowsEmailAddress.Domain

						}
						
						if ($tmpUser.Phone) {

							$telephone = ($tmpUser.Phone).Split('x')

							switch ($telephone.Count) {

								1 { 

									$Phone.Text = $telephone[0].Trim()

								}

								2 {

									$Phone.Text = $telephone[0].Trim(); 
									$Extension.Text = $telephone[1].Trim()

								}

								default {}

							}

						}

						$Domain.Enabled = $True
						$Email.Enabled = $True
						$EmailDomain.Enabled = $True
						$Server.Enabled = $True
						$MB.Enabled = $True
						$DG.Enabled = $True
						$Office.Enabled = $True
						$Phone.Enabled = $True
						$Extension.Enabled = $True
						$OutlookAnywhere.Enabled = $True
						$OWA.Enabled = $True
						$ActiveSync.Enabled = $True

					}

					if ($disable) { 

						ShowForm2

					} else {					

						$Form1.Refresh() 

					}

				}

			}

		})

	$Form1.Controls.Add($NextButton)

	<#
		Config "Next" button and its behaviors
	#>
	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size (1065, 770)
	$CancelButton.Size = New-Object System.Drawing.Size (105, 43)
	$CancelButton.Text = "Cancel"
	$CancelButton.TabIndex = 18
	$CancelButton.Add_Click({ 

		$Form1.close()

	})
	$Form1.Controls.Add($CancelButton)

	<#
		Config "Back" button and its behaviors
	#>
	$BackButton = New-Object System.Windows.Forms.Button
	$BackButton.Location = New-Object System.Drawing.Size (855, 770)
	$BackButton.Size = New-Object System.Drawing.Size (105, 43)
	$BackButton.Text = "< Back"
	$BackButton.TabIndex = 19
	$BackButton.Add_Click({

		$Form1.visible = $False; $launchForm.visible = $True

	})
	$Form1.Controls.Add($BackButton)

	<#
		Config error message box
	#>
	$ErrorMsg = New-Object System.Windows.Forms.Label
	$ErrorMsg.Location = New-Object System.Drawing.Size (20, 10)
	$ErrorMsg.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg.Text = $null
	$ErrorMsg.ForeColor = "Red"
	$Form1.Controls.Add($ErrorMsg)

	<#
		Config "Company/Office/Dept.:" label
	#>
	$OfficeLabel = New-Object System.Windows.Forms.Label
	$OfficeLabel.Location = New-Object System.Drawing.Size (20, 70)
	$OfficeLabel.Size = New-Object System.Drawing.Size (140, 20)
	$OfficeLabel.Text = "Company/Office/Dept.:"
	$OfficeLabel.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$Form1.Controls.Add($OfficeLabel)

	<# 
		Config title label
	#>
	$TitleLabel = New-Object System.Windows.Forms.Label
	$TitleLabel.Location = New-Object System.Drawing.Size (20, 40)
	$TitleLabel.Size = New-Object System.Drawing.Size (450, 20)
	$TitleLabel.Text = ""
	$TitleLabel.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$Form1.Controls.Add($TitleLabel)

	<#
		Config "Organizational Unit:" label
	#>
	$OULabel = New-Object System.Windows.Forms.Label
	$OULabel.Location = New-Object System.Drawing.Size (20, 100)
	$OULabel.Size = New-Object System.Drawing.Size (120, 20)
	$OULabel.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$OULabel.Text = "Organizational Unit:"
	$Form1.Controls.Add($OULabel)

	<#
		Config "First Name/Last Name:" label
	#>
	$FirstLabel = New-Object System.Windows.Forms.Label
	$FirstLabel.Location = New-Object System.Drawing.Size (20, 130)
	$FirstLabel.Size = New-Object System.Drawing.Size (130, 20)
	$FirstLabel.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$FirstLabel.Text = "First Name/Last Name:"
	$Form1.Controls.Add($FirstLabel)

	<#
		Config "User Logon Name:" label
	#>
	$LogonLabel = New-Object System.Windows.Forms.Label
	$LogonLabel.Location = New-Object System.Drawing.Size (20, 160)
	$LogonLabel.Size = New-Object System.Drawing.Size (120, 20)
	$LogonLabel.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$LogonLabel.Text = "User Logon Name:"
	$Form1.Controls.Add($LogonLabel)

	<#
		Config "Password: " label
	#>
	$PasswordLabel = New-Object System.Windows.Forms.Label
	$PasswordLabel.Location = New-Object System.Drawing.Size (20, 190)
	$PasswordLabel.Size = New-Object System.Drawing.Size (120, 20)
	$PasswordLabel.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$PasswordLabel.Text = "Password: "
	$Form1.Controls.Add($PasswordLabel)

	$PasswordLabel = New-Object System.Windows.Forms.Label
	$PasswordLabel.Location = New-Object System.Drawing.Size (340, 195)
	$PasswordLabel.Size = New-Object System.Drawing.Size (140, 20)
	$PasswordLabel.Font = New-Object System.Drawing.Font ("Arial", 7)
	$PasswordLabel.Text = "(minimum 5 characters)"
	$Form1.Controls.Add($PasswordLabel)

	<#
		Config "Server / Info Store:" label
	#>
	$ServerLabel = New-Object System.Windows.Forms.Label
	$ServerLabel.Location = New-Object System.Drawing.Size (20, 220)
	$ServerLabel.Size = New-Object System.Drawing.Size (120, 20)
	$ServerLabel.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$ServerLabel.Text = "Server / Info Store:"
	$Form1.Controls.Add($ServerLabel)

	<#
		Config "E-mail:" label
	#>
	$EmailLabel = New-Object System.Windows.Forms.Label
	$EmailLabel.Location = New-Object System.Drawing.Size (20, 540)
	$EmailLabel.Size = New-Object System.Drawing.Size (120, 20)
	$EmailLabel.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$EmailLabel.Text = "E-mail:"
	$Form1.Controls.Add($EmailLabel)

	<#
		Config "Distribution Group:" label
	#>
	$DGLabel = New-Object System.Windows.Forms.Label
	$DGLabel.Location = New-Object System.Drawing.Size (595, 60)
	$DGLabel.Size = New-Object System.Drawing.Size (120, 20)
	$DGLabel.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$DGLabel.Text = "Distribution Group:"
	$Form1.Controls.Add($DGLabel)

	<#
		Config "Phone / Extension:" label
	#>
	$PhoneLabel = New-Object System.Windows.Forms.Label
	$PhoneLabel.Location = New-Object System.Drawing.Size (20, 570)
	$PhoneLabel.Size = New-Object System.Drawing.Size (120, 15)
	$PhoneLabel.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$PhoneLabel.Text = "Phone / Extension:"
	$Form1.Controls.Add($PhoneLabel)

	$Phone2Label = New-Object System.Windows.Forms.Label
	$Phone2Label.Location = New-Object System.Drawing.Size (190, 590)
	$Phone2Label.Size = New-Object System.Drawing.Size (150, 20)
	$Phone2Label.Font = New-Object System.Drawing.Font ("Arial", 7)
	$Phone2Label.Text = "e.g. 408-941-0808"
	$Form1.Controls.Add($Phone2Label)

	$Phone3Label = New-Object System.Windows.Forms.Label
	$Phone3Label.Location = New-Object System.Drawing.Size (340, 590)
	$Phone3Label.Size = New-Object System.Drawing.Size (140, 20)
	$Phone3Label.Font = New-Object System.Drawing.Font ("Arial", 7)
	$Phone3Label.Text = "e.g. 808"
	$Form1.Controls.Add($Phone3Label)

	<#
		Config "Mailbox Features:" label
	#>
	$FeaturesLabel = New-Object System.Windows.Forms.Label
	$FeaturesLabel.Location = New-Object System.Drawing.Size (420, 370)
	$FeaturesLabel.Size = New-Object System.Drawing.Size (120, 15)
	$FeaturesLabel.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$FeaturesLabel.Text = "Mailbox Features:"
	#$Form1.Controls.Add($FeaturesLabel)

	<#
		Config "( Approve Needed )" label
	#>
	$ApproveLabel = New-Object System.Windows.Forms.Label
	$ApproveLabel.Location = New-Object System.Drawing.Size (420, 390)
	$ApproveLabel.Size = New-Object System.Drawing.Size (120, 15)
	$ApproveLabel.Text = "( Approve Needed )"
	#$Form1.Controls.Add($ApproveLabel)


	<#
		Config "Outlook Anywhere" label
	#>
	$OutlookAnywhereLabel = New-Object System.Windows.Forms.Label
	$OutlookAnywhereLabel.Location = New-Object System.Drawing.Size (565, 350)
	$OutlookAnywhereLabel.Size = New-Object System.Drawing.Size (150, 20)
	$OutlookAnywhereLabel.Text = "Outlook Anywhere"
	#$Form1.Controls.Add($OutlookAnywhereLabel)

	<#
		Config "Outlook Web Access (OWA)" label
	#>
	$OWALabel = New-Object System.Windows.Forms.Label
	$OWALabel.Location = New-Object System.Drawing.Size (565, 370)
	$OWALabel.Size = New-Object System.Drawing.Size (150, 20)
	$OWALabel.Text = "Outlook Web Access (OWA)"
	#$Form1.Controls.Add($OWALabel)

	<#
		"Exchange ActiveSync" label
	#>
	$ActiveSyncLabel = New-Object System.Windows.Forms.Label
	$ActiveSyncLabel.Location = New-Object System.Drawing.Size (565, 390)
	$ActiveSyncLabel.Size = New-Object System.Drawing.Size (120, 20)
	$ActiveSyncLabel.Text = "Exchange ActiveSync"
	#$Form1.Controls.Add($ActiveSyncLabel)

	<#
		Config "Company" combo box and its behaviors
	#>
	$MainCompany = New-Object System.Windows.Forms.ComboBox
	$MainCompany.Location = New-Object System.Drawing.Size (190, 70)
	$MainCompany.Size = New-Object System.Drawing.Size (90, 20)

	[void]$MainCompany.Items.Add("")
	[void]$MainCompany.Items.Add("MA Labs")
	[void]$MainCompany.Items.Add("Superbiiz")
	[void]$MainCompany.Items.Add("Supertalent")
	$MainCompany.Add_SelectedValueChanged({

		$LogonName.Enabled = $False
		$Domain.Enabled = $False
		$Password.Enabled = $False
		$Server.Enabled = $False
		$MB.Enabled = $False
		$Email.Enabled = $False
		$EmailDomain.Enabled = $False
		$DG.Enabled = $False
		$Phone.Enabled = $False
		$Extension.Enabled = $False
		$objListbox_DFS.Enabled = $False
		$objListbox_fileshare.Enabled = $False
		$Domain.Items.Clear()
		[void]$Domain.Items.Add("@ma.local")
		[void]$Domain.Items.Add("@superbiiz.com")
		[void]$Domain.Items.Add("@supertalent.com")
		$EmailDomain.Items.Clear()
		[void]$EmailDomain.Items.Add("@ma.local")
		[void]$EmailDomain.Items.Add("@superbiiz.com")
		[void]$EmailDomain.Items.Add("@supertalent.com")
		$OU.Items.Clear()
		$OU.Text = $null
		$Office.Items.Clear()
		$Office.Text = $null
		$Pmgroup.visible = $False
		$Pmgroup.Items.Clear()
		$Email.Text = $null
		$Password.Text = $null
		$LogonName.Text = $null
		$Phone.Text = $null
		$Extension.Text = $null

		if ($MainCompany.SelectedItem.ToString() -eq "") {

			$MainCompanyName = ''
			$TitleLabel.Text = ''

		} else {

			$MainCompanyName = $MainCompany.SelectedItem.ToString()
			$TitleLabel.Text = $MainCompanyName

		}

		switch ($MainCompanyName) {

			{ $_ -eq "MA Labs" } {

				$Domain.SelectedIndex = 0
				$Office.Items.Add("Chicago")
				$Office.Items.Add("Georgia")
				$Office.Items.Add("Los Angles")
				$Office.Items.Add("New Jersey")
				$Office.Items.Add("Miami")
				$Office.Items.Add("San Jose")
				$Office.Items.Add("Wuhan")

			}

			{ $_ -eq "Superbiiz" } {

				$Domain.SelectedIndex = 1
				$Office.Items.Add("San Jose")
				$Office.Items.Add("Wuhan")

			}

			{ $_ -eq "Supertalent" } {

				$Domain.SelectedIndex = 2
				$Office.Items.Add("San Jose")
				$Office.Items.Add("Wuhan")

			}

		}

	})

	$MainCompany.DropDownStyle = 2
	$MainCompany.Sorted = $True
	$MainCompany.TabIndex = 0
	if (!$new) {

		$MainCompany.Enabled = $False

	}
	$Form1.Controls.Add($MainCompany)


	<#
		Config "Office" combo box and its behaviors
	#>
	$Office = New-Object System.Windows.Forms.ComboBox
	$Office.Location = New-Object System.Drawing.Size (290, 70)
	$Office.Size = New-Object System.Drawing.Size (90, 20)

	<#
		Office lead the Data Directory
	#>
	$Office.Add_SelectedValueChanged({

		$LogonName.Enabled = $False
		$Domain.Enabled = $False
		$Password.Enabled = $False
		$Server.Enabled = $False
		$MB.Enabled = $False
		$Email.Enabled = $False
		$EmailDomain.Enabled = $False
		$DG.Enabled = $False
		$Phone.Enabled = $False
		$Extension.Enabled = $False
		$objListbox_DFS.Enabled = $False
		$objListbox_fileshare.Enabled = $False

		#$OWA.Enabled = $False
		#$ActiveSync.Enabled = $False
		$OU.Items.Clear()
		$OU.Text = $null
		$Email.Text = $null
		$Password.Text = $null
		$LogonName.Text = $null
		$Phone.Text = $null
		$Extension.Text = $null
		$root = [adsi]''
		$searcher = New-Object System.DirectoryServices.DirectorySearcher ($root)

		if ($Office.SelectedItem.ToString() -eq "") {

			$Officename = ''

		} else {

			$Officename = $Office.SelectedItem.ToString()
			$MainCompanyName = $MainCompany.SelectedItem.ToString()
			$NameCombine = "$MainCompanyName/$Officename"

			$Department.Items.Clear()
			$Pmgroup.visible = $False
			$Pmgroup.Items.Clear()

			switch ($NameCombine) {

				{ $_ -eq "MA Labs/Georgia" } {

					$Department.Items.Add('Accounting');
					$Department.Items.Add('HR');
					$Department.Items.Add('Others');
					$Department.Items.Add('RMA');
					$Department.Items.Add('Sales');

				}

				{ $_ -eq "MA Labs/San Jose" } {

					$Department.Items.Add('AP');
					$Department.Items.Add('AR');
					$Department.Items.Add('IT');
					$Department.Items.Add('Sales');
					$Department.Items.Add('Marketing');
					$Department.Items.Add('HR');
					$Department.Items.Add('Payroll');
					$Department.Items.Add('MIS');
					$Department.Items.Add('Credit');
					$Department.Items.Add('Data Entry');
					$Department.Items.Add('Purchasing');
					$Department.Items.Add('Shipping');
					$Department.Items.Add('Tech Support');
					$Department.Items.Add('Warehouse');
					$Department.Items.Add('ACCT');
					$Department.Items.Add('RMA') }

				{ $_ -eq "MA Labs/Los Angles" } {

					$Department.Items.Add('Accounting');
					$Department.Items.Add('Others');
					$Department.Items.Add('RMA');
					$Department.Items.Add('Sales');
					$Department.Items.Add('Warehouse');
					$Department.Items.Add('HR');

				}

				{ $_ -eq "MA Labs/New Jersey" } {

					$Department.Items.Add('Others');
					$Department.Items.Add('RMA')
					$Department.Items.Add('Sales');
					$Department.Items.Add('Warehouse');
					$Department.Items.Add('Accounting');
					$Department.Items.Add('Purchasing');

				}

				{ $_ -eq "MA Labs/Chicago" } {

					$Department.Items.Add('Accounting');
					$Department.Items.Add('Others');
					$Department.Items.Add('Sales');

				}

				{ $_ -eq "MA Labs/Miami" } {

					$Department.Items.Add('Accounting');
					$Department.Items.Add('Others');
					$Department.Items.Add('Sales');
					$Department.Items.Add('Purchasing');
					$Department.Items.Add('Warehouse');

				}

				{ $_ -eq "MA Labs/Wuhan" } {

					$Department.Items.Add('AP');
					$Department.Items.Add('AR');
					$Department.Items.Add('Credit');
					$Department.Items.Add('Inventory');
					$Department.Items.Add('Marketing');
					$Department.Items.Add('PM');
					$Department.Items.Add('HR');
					$Department.Items.Add('Sales');

				}

				{ $_ -eq "Supertalent/San Jose" } {

					$Department.Items.Add('Accounting');
					$Department.Items.Add('Engineering')
					$Department.Items.Add('HR');
					$Department.Items.Add('Sales');
					$Department.Items.Add('Marketing');
					$Department.Items.Add('Tech Support');
					$Department.Items.Add('RMA');

				}

				{ $_ -eq "Superbiiz/San Jose" } {

					$Department.Items.Add('Marketing');
					$Department.Items.Add('Customer Service');
					$Department.Items.Add('Sales');
					$Department.Items.Add('Accounting');
					$Department.Items.Add('MIS')

				}

				{ $_ -eq "Superbiiz/Wuhan" } {

					$Department.Items.Add('Users');
					$Department.Items.Add('Accounting');

				}

				{ $_ -eq "Supertalent/Wuhan" } {

					$Department.Items.Add('Users');
					$Department.Items.Add('Sales')

				}

			}

		}

		#The Group Note Search to detect the distribution group
		#Get-Group -Filter "Notes -eq '$Officename'"| ForEach-Object{[void] $DG.Items.Add($_.Name)}

	})

	$Office.DropDownStyle = 2
	$Office.Sorted = $True
	$Office.TabIndex = 1
	if (!$new) {

		$Office.Enabled = $False

	}
	$Form1.Controls.Add($Office)

	<#
		Config "Department" combo box
	#>
	$Department = New-Object System.Windows.Forms.ComboBox
	$Department.Location = New-Object System.Drawing.Size (390, 70)
	$Department.Size = New-Object System.Drawing.Size (90, 20)
	$Department.Add_SelectedValueChanged({

		$Pmgroup.visible = $False
		$Pmgroup.Items.Clear()
		$PmgroupName = $Null

		if ($Department.SelectedItem.ToString() -eq "") {

			$DepartmentName = ''

		} else {

			$DepartmentName = $Department.SelectedItem.ToString()

		}

		$root = [adsi]''
		$searcher = New-Object System.DirectoryServices.DirectorySearcher ($root)

		if ($Office.SelectedItem.ToString() -eq "") {

			$Officename = ''

		} else {

			$MainCompanyName = $MainCompany.SelectedItem.ToString()
			$Officename = $Office.SelectedItem.ToString()
			$NameCombine = "$MainCompanyName/$Officename"

			#Control the subgroup of the purchasing
			if (($Office.Text -eq "San Jose") -and ($Department.Text -eq "Purchasing")) {
				$Pmgroup.visible = $True
				$Pmgroup.Items.Add("HDD")
				$Pmgroup.Items.Add("Memory")
				$Pmgroup.Items.Add("Microsoft")
				$Pmgroup.Items.Add("Monitor")
				$Pmgroup.Items.Add("Motherboard")
				$Pmgroup.Items.Add("Networking")
				$Pmgroup.Items.Add("Notebook")
				$Pmgroup.Items.Add("VGA")
			}

			if (($Office.Text -eq "Wuhan") -and ($Department.Text -eq "PM")) {

				$Pmgroup.visible = $True
				$Pmgroup.Items.Add("HDD")
				$Pmgroup.Items.Add("Microsoft")
				$Pmgroup.Items.Add("Monitor")
				$Pmgroup.Items.Add("Motherboard")
				$Pmgroup.Items.Add("Networking")
				$Pmgroup.Items.Add("Notebook")

			}

			#The Group Note Search to detect the distribution group
			#Get-Group -Filter "Notes -eq '$Officename'"| ForEach-Object{[void] $DG.Items.Add($_.Name)}
			switch ($NameCombine) {

				{ $_ -eq "MA Labs/Georgia" } {

					$Phone.Text = "770-209-6600";
					$Extension.MaxLength = 3;
					$OUKeyword = "MA Labs/GA"

				}

				{ $_ -eq "MA Labs/San Jose" } {

					$Phone.Text = "408-941-0808";
					$Extension.MaxLength = 3;
					$OUKeyword = "MA Labs/San Jose/$DepartmentName"

				}

				{ $_ -eq "MA Labs/Los Angles" } {

					$Phone.Text = "626-820-8988";
					$Extension.MaxLength = 3;
					$OUKeyword = "MA Labs/LA"

				}

				{ $_ -eq "MA Labs/New Jersey" } {

					$Phone.Text = "732-661-3388";
					$Extension.MaxLength = 3;
					$OUKeyword = "MA Labs/NJ"

				}

				{ $_ -eq "MA Labs/Chicago" } {

					$Phone.Text = "630-893-2323";
					$Extension.MaxLength = 3;
					$OUKeyword = "MA Labs/Chicago"

				}

				{ $_ -eq "MA Labs/Miami" } {

					$Phone.Text = "305-594-8700";
					$Extension.MaxLength = 3;
					$OUKeyword = "MA Labs/Miami"

				}

				{ $_ -eq "MA Labs/Wuhan" } {

					$Phone.Text = "086-275-973-1208";
					$Extension.MaxLength = 4;
					$OUKeyword = "MA Labs/China"

				}

				{ $_ -eq "Supertalent/San Jose" } {

					$Phone.Text = "408-934-2560";
					$Extension.MaxLength = 3;
					$OUKeyword = "Supertalent/$DepartmentName"

				}

				{ $_ -eq "Superbiiz/San Jose" } {

					$Phone.Text = "408-934-2500";
					$Extension.MaxLength = 3;
					$OUKeyword = "Superbiiz/$DepartmentName"

				}

				{ $_ -eq "Superbiiz/Wuhan" } {

					$Phone.Text = "086-275-973-1208";
					$Extension.MaxLength = 4;
					$OUKeyword = "Superbiiz/China/"

				}

				{ $_ -eq "Supertalent/Wuhan" } {

					$Phone.Text = "086-275-973-1208";
					$Extension.MaxLength = 4;
					$OUKeyword = "Supertalent/China/"

				}

			}

			$OU.Items.Clear()
			$searcher.Filter = "(&(objectClass=organizationalUnit)(name=Users))"
			[void]$searcher.PropertiesToLoad.Add("canonicalName")
			[void]$searcher.PropertiesToLoad.Add("Name")
			$searcherall = $searcher.FindAll()
			$findTag = 0
			foreach ($person in $searcherall) {

				[string]$ent = $person.properties.canonicalname
				$OUTarget = $ent.substring($ent.IndexOf("/"), $ent.Length - $ent.IndexOf("/"))

				if (($OUTarget -like "*$OUKeyword*") -and !($OUTarget -like "*Disabled*")) {

					[void]$OU.Items.Add($OUTarget)
					$OU.SelectedIndex = 0
					$findTag = 1

				}

				if (($NameCombine -eq "Supertalent/San Jose") -and ($OUTarget -like "*Supertalent/China*")) {

					[void]$OU.Items.Remove($OUTarget)

				}

				if (($NameCombine -eq "Superbiiz/San Jose") -and ($OUTarget -like "*Superbiiz/China*")) {

					[void]$OU.Items.Remove($OUTarget)

				}

			}

		}

	})

	$Department.DropDownStyle = 2
	$Department.Sorted = $True
	$Department.DropDownHeight = 750
	$Department.TabIndex = 1
	if (!$new) {

		$Department.Enabled = $False

	}
	$Form1.Controls.Add($Department)

	<# 
		Subgroup for PM (Specially for different selling target)
	#>
	$Pmgroup = New-Object System.Windows.Forms.ComboBox
	$Pmgroup.Location = New-Object System.Drawing.Size (490, 70)
	$Pmgroup.Size = New-Object System.Drawing.Size (90, 20)
	$Pmgroup.DropDownStyle = 2
	$Pmgroup.Sorted = $True
	$Pmgroup.visible = $False
	$Pmgroup.Add_SelectedValueChanged({

		if ($Pmgroup.SelectedItem.ToString() -eq "") {

			$PmgroupName = ''

		} else {

			$PmgroupName = $Pmgroup.SelectedItem.ToString() 

		}

	})
	$Form1.Controls.Add($Pmgroup)

	<#
		Config "Organizational Unit" combo box
	#>
	$OUhash = @{}
	$OU = New-Object System.Windows.Forms.ComboBox
	$OU.Location = New-Object System.Drawing.Size (190, 100)
	$OU.Size = New-Object System.Drawing.Size (290, 20)
	#$OUhash.Add("/Users","Users")
	[void]$OU.Items.Add("/Users")
	$OU.DropDownStyle = 2
	$OU.DropDownHeight = 750
	$OU.Sorted = $True
	$OU.TabIndex = 2
	if (!$new) {

		$OU.Enabled = $False

	}

	$Form1.Controls.Add($OU)

	<#
		Config "Organizational Unit" message box
	#>
	$OUMsg = New-Object System.Windows.Forms.Label
	$OUMsg.Location = New-Object System.Drawing.Size (450, 105)
	$OUMsg.Size = New-Object System.Drawing.Size (20, 20)
	if ($new) {

		$OUMsg.Text = "*"

	}
	$OUMsg.ForeColor = "Red"
	$Form1.Controls.Add($OUMsg)

	$UserName = New-Object System.Windows.Forms.TextBox
	$UserName.Location = New-Object System.Drawing.Size (190, 130)
	$UserName.Size = New-Object System.Drawing.Size (240, 20)
	$UserName.TabIndex = 3
	#$FirstName.MaxLength = 30
	if (!$new) {

		$UserName.Enabled = $False

	}
	$Form1.Controls.Add($UserName)

	<#
		First Name Text Box
	#>
	$FirstName = New-Object System.Windows.Forms.TextBox
	$FirstName.Location = New-Object System.Drawing.Size (190, 130)
	$FirstName.Size = New-Object System.Drawing.Size (120, 20)
	$FirstName.TabIndex = 3
	#$FirstName.MaxLength = 30
	if (!$new) { $FirstName.Enabled = $False }
	#$Form1.Controls.Add($FirstName)

	<#
		First Name Message Box
	#>
	$FirstMsg = New-Object System.Windows.Forms.Label
	$FirstMsg.Location = New-Object System.Drawing.Size (490, 135)
	$FirstMsg.Size = New-Object System.Drawing.Size (20, 20)
	if ($new) { $FirstMsg.Text = "*" }
	$FirstMsg.ForeColor = "Red"
	$Form1.Controls.Add($FirstMsg)

	<#
		Last Name Text Box
	#>
	$LastName = New-Object System.Windows.Forms.TextBox
	$LastName.Location = New-Object System.Drawing.Size (315, 130)
	$LastName.Size = New-Object System.Drawing.Size (120, 20)
	#$LastName.MaxLength = 30
	$LastName.TabIndex = 4
	if (!$new) {

		$LastName.Enabled = $False

	}
	#$Form1.Controls.Add($LastName)

	<#
		User Autofill Button
	#>
	$Autofill = New-Object System.Windows.Forms.Button
	$Autofill.Location = New-Object System.Drawing.Size (440, 130)
	$Autofill.Size = New-Object System.Drawing.Size (40, 20)
	$Autofill.Text = "Auto Fill"
	$Autofill.TabIndex = 5
	$Autofill.Add_Click({

			$UserFirst = ToProperCase ($UserName.Text.Split()[0].Trim())
			$UserLast = ToProperCase ($UserName.Text.Split()[1].Trim())
			$UserOU = $OU.SelectedItem
			if (Autofill) {

				ShowForm2

			}

	})

	if (!$new) {

		$Autofill.Enabled = $False

	}

	$Form1.Controls.Add($Autofill)

	<#
		Logon Name Text Box
	#>
	$LogonName = New-Object System.Windows.Forms.TextBox
	$LogonName.Location = New-Object System.Drawing.Size (190, 160)
	$LogonName.Size = New-Object System.Drawing.Size (140, 20)
	if ($new) {

		$LogonName.MaxLength = 22
		$Logonname.Enabled = $False

	} else {

		$LogonName.AutoCompleteSource = 'CustomSource'
		$LogonName.AutoCompleteMode = 'SuggestAppend'
		$LogonName.AutoCompleteCustomSource = $autocomplete
		Get-User | ForEach-Object {

			$LogonName.AutoCompleteCustomSource.Add($_.SamAccountName)

		}

	}

	$LogonName.TabIndex = 6
	$LogonName.CharacterCasing = 2
	$Form1.Controls.Add($LogonName)

	<#
		Logon Name Message Box
	#>
	$LogonMsg = New-Object System.Windows.Forms.Label
	$LogonMsg.Location = New-Object System.Drawing.Size (490, 165)
	$LogonMsg.Size = New-Object System.Drawing.Size (20, 20)
	$LogonMsg.Text = "*"
	$LogonMsg.ForeColor = "Red"
	$Form1.Controls.Add($LogonMsg)

	<#
		Domain Combo Box
	#>
	$Domain = New-Object System.Windows.Forms.ComboBox
	$Domain.Location = New-Object System.Drawing.Size (340, 160)
	$Domain.Size = New-Object System.Drawing.Size (140, 20)
	[void]$Domain.Items.Add("@ma.local")
	[void]$Domain.Items.Add("@superbiiz.com")
	[void]$Domain.Items.Add("@supertalent.com")
	$Domain.Add_SelectedValueChanged({ 

		$EmailDomain.SelectedItem = $Domain.SelectedItem 

	})
	$Domain.DropDownStyle = 2
	$Domain.Sorted = $True
	$Domain.TabIndex = 7
	$Domain.Enabled = $False
	$Form1.Controls.Add($Domain)

	<#
		Password Text Box
	#>
	$Password = New-Object System.Windows.Forms.TextBox
	$Password.Location = New-Object System.Drawing.Size (190, 190)
	$Password.Size = New-Object System.Drawing.Size (140, 20)
	$Password.MaxLength = 20
	$Password.UseSystemPasswordChar = $False
	$Password.TabIndex = 8
	$Password.Enabled = $False
	$Form1.Controls.Add($Password)

	<#
		Password Message Box
	#>
	$PasswordMsg = New-Object System.Windows.Forms.Label
	$PasswordMsg.Location = New-Object System.Drawing.Size (490, 195)
	$PasswordMsg.Size = New-Object System.Drawing.Size (20, 20)
	if ($new) {

		$PasswordMsg.Text = "*"

	}
	$PasswordMsg.ForeColor = "Red"
	$Form1.Controls.Add($PasswordMsg)

	<#
		Server Combo Box
	#>
	$Server = New-Object System.Windows.Forms.ComboBox
	$Server.Location = New-Object System.Drawing.Size (190, 220)
	$Server.Size = New-Object System.Drawing.Size (140, 20)
	Get-MailboxServer | ForEach-Object {

		[void]$Server.Items.Add($_.Name)
		if ($_.Name -eq "PI") {

			$Server.Items.Remove("PI")

		}

	}

	$Server.Add_SelectedValueChanged({

		$MBhash = @{}
		$MB.Items.Clear()
		$MB.Text = $null
		Get-MailboxDatabase -Status -Server $Server.SelectedItem.ToString() | ForEach-Object {

			$Name = $_.Name
			$ServerName = $_.server
			$Filepath1 = $_.EdbFilePath
			$Fullpath2 = "`\`\" + $_.server + "`\" + $_.EdbFilePath.DriveName.Remove(1).ToString() + "$" + $_.EdbFilePath.PathName.Remove(0, 2)
			$Sizeinfo = ((Get-ChildItem $Fullpath2).Length) / 1048576KB
			$Size = [math]::Round($Sizeinfo, 2)

			$MBInfo = "$Name == $Size GB"
			if ($Size -lt 40) {

				$MB.ForeColor = "Black"

			} else {

				$MB.ForeColor = "Red"

			}

			[void]$MB.Items.Add($MBInfo)
			$MBhash.Add($_.Name,$_.ServerName + "\" + $_.StorageGroup.Name + "\" + $_.Name)

		}

	})

	$Server.DropDownStyle = 2
	$Server.Sorted = $True
	$Server.TabIndex = 9
	$Server.Enabled = $False
	$Form1.Controls.Add($Server)

	<#
		Server Name Message Box
	#>
	$ServerMsg = New-Object System.Windows.Forms.Label
	$ServerMsg.Location = New-Object System.Drawing.Size (490, 225)
	$ServerMsg.Size = New-Object System.Drawing.Size (20, 20)
	if ($new) {

		$ServerMsg.Text = "*"

	}
	$ServerMsg.ForeColor = "Red"
	$Form1.Controls.Add($ServerMsg)

	<#
		Database Combo Box
	#>
	$MB = New-Object System.Windows.Forms.CheckedListbox
	$MB.Location = New-Object System.Drawing.Size (190, 250)
	$MB.Size = New-Object System.Drawing.Size (290, 20)
	$MB.Height = 280
	$MB.Width = 290
	$MB.CheckOnClick = $True
	$MB.HorizontalScrollbar = $True
	$MB.Enabled = $False
	$Form1.Controls.Add($MB)

	<#
		E-mail Text Box
	#>
	$Email = New-Object System.Windows.Forms.TextBox
	$Email.Location = New-Object System.Drawing.Size (190, 540)
	$Email.Size = New-Object System.Drawing.Size (140, 20)
	$Email.MaxLength = 40
	$Email.TabIndex = 11
	$Email.Enabled = $False
	$Form1.Controls.Add($Email)

	<#
		E-mail Combo Box
	#>
	$EmailDomain = New-Object System.Windows.Forms.ComboBox
	$EmailDomain.Location = New-Object System.Drawing.Size (340, 540)
	$EmailDomain.Size = New-Object System.Drawing.Size (140, 20)
	[void]$EmailDomain.Items.Add("@ma.local")
	[void]$EmailDomain.Items.Add("@superbiiz.com")
	[void]$EmailDomain.Items.Add("@supertalent.com")
	$EmailDomain.DropDownStyle = 2
	$EmailDomain.Sorted = $True
	$EmailDomain.TabIndex = 12
	$EmailDomain.Enabled = $False
	$Form1.Controls.Add($EmailDomain)

	<#
		E-mail Message Box
	#>
	$EmailMsg = New-Object System.Windows.Forms.Label
	$EmailMsg.Location = New-Object System.Drawing.Size (450, 545)
	$EmailMsg.Size = New-Object System.Drawing.Size (20, 20)
	$EmailMsg.Text = $null
	$EmailMsg.ForeColor = "Red"
	$Form1.Controls.Add($EmailMsg)

	<#
		Distribution Group Combo Box
	#>
	$DG = New-Object System.Windows.Forms.CheckedListbox
	$DG.Location = New-Object System.Drawing.Size (595, 80)
	$DG.Size = New-Object System.Drawing.Size (180, 700)
	$DG.Height = 520
	$DG.Width = 160
	$DG.HorizontalScrollbar = $True
	$DG.Enabled = $False
	$Form1.Controls.Add($DG)

	#The DGTextbox for the real name of the distribution group
	#$DGdisplay = new-object System.Windows.Forms.TextBox
	#$DGdisplay.Location = New-Object System.Drawing.Size(700,280)
	#$DGdisplay.Size = New-Object System.Drawing.Size(140,20)
	#$DGdisplay.Enabled = $False
	#$Form1.Controls.Add($DGdisplay)

	<#
		Phone Text Box
	#>
	$Phone = New-Object System.Windows.Forms.TextBox
	$Phone.Location = New-Object System.Drawing.Size (190, 570)
	$Phone.Size = New-Object System.Drawing.Size (140, 20)
	$Phone.MaxLength = 12
	$Phone.TabIndex = 14
	$Phone.Text = $null
	$Phone.Enabled = $True
	$Phone.Enabled = $False
	$Form1.Controls.Add($Phone)

	<#
		Extension Text Box
	#>
	$Extension = New-Object System.Windows.Forms.TextBox
	$Extension.Location = New-Object System.Drawing.Size (340, 570)
	$Extension.Size = New-Object System.Drawing.Size (140, 20)
	$Extension.TabIndex = 15
	$Extension.Text = $null
	$Extension.Enabled = $True
	$Extension.Enabled = $False
	$Form1.Controls.Add($Extension)

	<#
		Phone Message Box
	#>
	$PhoneMsg = New-Object System.Windows.Forms.Label
	$PhoneMsg.Location = New-Object System.Drawing.Size (450, 575)
	$PhoneMsg.Size = New-Object System.Drawing.Size (20, 20)
	$PhoneMsg.Text = $null
	$PhoneMsg.ForeColor = "Red"
	$Form1.Controls.Add($PhoneMsg)

	<#
		Outlook Anywhere Check Box
	#>
	$OutlookAnywhere = New-Object System.Windows.Forms.CheckBox
	$OutlookAnywhere.Location = New-Object System.Drawing.Size (550, 348)
	$OutlookAnywhere.Size = New-Object System.Drawing.Size (140, 20)
	$OutlookAnywhere.Checked = $False
	$OutlookAnywhere.TabIndex = 14
	$OutlookAnywhere.Enabled = $True
	if (!$new) {

		$OutlookAnywhere.Enabled = $False

	}
	# $Form1.Controls.Add($OutlookAnywhere)

	<#
		OWA Check Box
	#>
	$OWA = New-Object System.Windows.Forms.CheckBox
	$OWA.Location = New-Object System.Drawing.Size (550, 368)
	$OWA.Size = New-Object System.Drawing.Size (140, 20)
	$OWA.Checked = $False
	$OWA.TabIndex = 16
	$OWA.Enabled = $True
	$OWA.Enabled = $False
	#$Form1.Controls.Add($OWA)

	<#
		Exchange ActiveSync Check Box
	#>
	$ActiveSync = New-Object System.Windows.Forms.CheckBox
	$ActiveSync.Location = New-Object System.Drawing.Size (550, 388)
	$ActiveSync.Size = New-Object System.Drawing.Size (140, 20)
	$ActiveSync.Checked = $False
	$ActiveSync.TabIndex = 16
	$ActiveSync.Enabled = $False
	#$Form1.Controls.Add($ActiveSync)

	<#
		The Lync label
	#>
	$Lync = New-Object System.Windows.Forms.Label
	$Lync.Location = New-Object System.Drawing.Size (20, 615)
	$Lync.Size = New-Object System.Drawing.Size (140, 20)
	$Lync.Text = "MS Lync Account:"
	$Lync.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	#$Form1.Controls.Add($Lync)

	<#
		The Lync Check box
	#>
	$LyncCheck = New-Object System.Windows.Forms.CheckBox
	$LyncCheck.Location = New-Object System.Drawing.Size (190, 615)
	$LyncCheck.Size = New-Object System.Drawing.Size (140, 20)
	$LyncCheck.Checked = $False
	$LyncCheck.TabIndex = 18
	$LyncCheck.Enabled = $True
	#$Form1.Controls.Add($LyncCheck)

	<#
		Add Group Membership
	#>
	$label1 = New-Object System.Windows.Forms.Label
	$label1.Location = New-Object System.Drawing.Size (775, 60)
	$label1.Size = New-Object System.Drawing.Size (85, 20)
	$label1.Text = "Network Drive: "
	$label1.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$label1.TabStop = $False
	$Form1.Controls.Add($label1)

	$label2 = New-Object System.Windows.Forms.Label
	$label2.Location = New-Object System.Drawing.Size (775,80)
	$label2.Size = New-Object System.Drawing.Size (85,20)
	$label2.Text = "NDFS: "
	$label2.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$label2.TabStop = $False
	$Form1.Controls.Add($label2)

	$label3 = New-Object System.Windows.Forms.Label
	$label3.Location = New-Object System.Drawing.Size (775, 350)
	$label3.Size = New-Object System.Drawing.Size (85, 20)
	$label3.Text = "FS: "
	$label3.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$label3.TabStop = $False
	$Form1.Controls.Add($label3)

	$label4 = New-Object System.Windows.Forms.Label
	$label4.Location = New-Object System.Drawing.Size (1000,80)
	$label4.Size = New-Object System.Drawing.Size (165,20)
	$label4.Text = "Websense Department: "
	$label4.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$label4.TabStop = $False
	$Form1.Controls.Add($label4)

	$label5 = New-Object System.Windows.Forms.Label
	$label5.Location = New-Object System.Drawing.Size (1000,350)
	$label5.Size = New-Object System.Drawing.Size (185,20)
	$label5.Text = "Websense User Define: "
	$label5.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$label5.TabStop = $False
	$Form1.Controls.Add($label5)

	#$Group = new-object System.Windows.Forms.ComboBox
	#$Group.Location = new-object System.Drawing.Size(20,135)
	#$Group.Size = new-object System.Drawing.Size(180,20)
	#[void] $Group.Items.Add("DFS")
	#[void] $Group.Items.Add("File Share")
	#$Group.Add_SelectedValueChanged({
	#$objListbox.Items.Clear()
	#$searcher_G = new-object system.directoryservices.directorysearcher;
	#$grp_G = New-Object system.directoryservices.directoryentry;
	#$Groupname = $Group.SelectedItem.ToString()
	#$PathDirection1 = "OU=$Groupname,OU=Groups,DC=malabs,DC=com"
	#$root_G = [ADSI]("LDAP://" + $PathDirection1 )
	#$searcher_G.SearchRoot = $root_G
	#$searcher_G.SearchScope = "Onelevel"
	#$searcher_G.FindAll() | foreach {
	#$SecurityGroupName = $_.properties.name
	#[void] $objListbox.Items.Add("$SecurityGroupName")
	#}

	#})
	#$Form1.Controls.Add($Group)

	<#
		The list box of group name under DFS
	#>
	$objListbox_DFS = New-Object System.Windows.Forms.CheckedListbox
	$objListbox_DFS.Location = New-Object System.Drawing.Size (775, 100)
	$objListbox_DFS.Size = New-Object System.Drawing.Size (180, 20)
	$objListbox_DFS.Height = 230
	$objListbox_DFS.Width = 200
	$objListbox_DFS.CheckOnClick = $True
	$objListbox_DFS.HorizontalScrollbar = $True
	$objListbox_DFS.Enabled = $False
	$Form1.Controls.Add($objListbox_DFS)

	# The list box of group name under vasto
	$objListbox_fileshare = New-Object System.Windows.Forms.CheckedListbox
	$objListbox_fileshare.Location = New-Object System.Drawing.Size (775, 380)
	$objListbox_fileshare.Size = New-Object System.Drawing.Size (180, 20)
	$objListbox_fileshare.Height = 220
	$objListbox_fileshare.Width = 200
	$objListbox_fileshare.CheckOnClick = $True
	$objListbox_fileshare.HorizontalScrollbar = $True
	$objListbox_fileshare.Enabled = $False
	$Form1.Controls.Add($objListbox_fileshare)

	# The list box of group name for websense department
	$objListbox_WebsenseDept = New-Object System.Windows.Forms.CheckedListbox
	$objListbox_WebsenseDept.Location = New-Object System.Drawing.Size (1000, 100)
	$objListbox_WebsenseDept.Size = New-Object System.Drawing.Size (180, 20)
	$objListbox_WebsenseDept.Height = 230
	$objListbox_WebsenseDept.Width = 200
	$objListbox_WebsenseDept.CheckOnClick = $True
	$objListbox_WebsenseDept.HorizontalScrollbar = $True
	$objListbox_WebsenseDept.Enabled = $False

	$Form1.Controls.Add($objListbox_WebsenseDept)

	# The list box of group name for websense list
	$objListbox_WebsenseList = New-Object System.Windows.Forms.CheckedListbox
	$objListbox_WebsenseList.Location = New-Object System.Drawing.Size (1000, 380)
	$objListbox_WebsenseList.Size = New-Object System.Drawing.Size (180, 20)
	$objListbox_WebsenseList.Height = 220
	$objListbox_WebsenseList.Width = 200
	$objListbox_WebsenseList.CheckOnClick = $True
	$objListbox_WebsenseList.HorizontalScrollbar = $True
	$objListbox_WebsenseList.Enabled = $False
	$Form1.Controls.Add($objListbox_WebsenseList)


	# AD Tree View

	$objIPProperties = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties()
	$strDNSDomain = $objIPProperties.DomainName.ToLower()
	$strDomainDN = $strDNSDomain.ToString().Split('.'); 
	foreach ($strVal in $strDomainDN) { 
	
		$strTemp += "dc=$strVal," 
	
	}; 
	$strDomainDN = $strTemp.TrimEnd(",").ToLower()
	#endregion 

	#region Generated Form Objects 
	$treeView1 = New-Object System.Windows.Forms.TreeView
	$label1 = New-Object System.Windows.Forms.Label
	$textbox1 = New-Object System.Windows.Forms.TextBox
	$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
	#endregion Generated Form Objects 

	#---------------------------------------------- 
	#Generated Event Script Blocks 
	#---------------------------------------------- 
	#$button1_OnClick=  
	#{ 
	#$form1.Close() 

	#} 

	$OnLoadForm_StateCorrection = {

		Build-TreeView

	}

	#---------------------------------------------- 
	#region Generated Form Code 

	#$treeView1.Size = New-Object System.Drawing.Size(260,270)
	#$treeView1.Name = "treeView1" 
	#$treeView1.Location = New-Object System.Drawing.Size(15,15)
	#$treeView1.DataBindings.DefaultDataSourceUpdateMode = 0 
	#$treeView1.TabStop = $False
	#$form1.Controls.Add($treeView1)

	$label1.Name = "label1"
	$label1.Location = New-Object System.Drawing.Size (15, 650)
	$label1.Size = New-Object System.Drawing.Size (50, 20)
	$label1.Text = "Berdych: "
	$label1.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$label1.TabStop = $False
	#$form1.Controls.Add($label1)

	$label2 = New-Object System.Windows.Forms.Label
	$label2.Location = New-Object System.Drawing.Size (65, 650)
	$label2.Size = New-Object System.Drawing.Size (100, 20)
	$label2.Text = "HQ-PM"
	$label2.TabStop = $False
	#$form1.Controls.Add($label2)

	$label3 = New-Object System.Windows.Forms.Label
	$label3.Location = New-Object System.Drawing.Size (15, 770)
	$label3.Size = New-Object System.Drawing.Size (50, 20)
	$label3.Text = "Isner: "
	$label3.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$label3.TabStop = $False
	$form1.Controls.Add($label3)

	$label4 = New-Object System.Windows.Forms.Label
	$label4.Location = New-Object System.Drawing.Size (70, 770)
	$label4.Size = New-Object System.Drawing.Size (70, 20)
	$label4.Text = "HQ-Sales"
	$label4.TabStop = $False
	$form1.Controls.Add($label4)

	$label5 = New-Object System.Windows.Forms.Label
	$label5.Location = New-Object System.Drawing.Size (140, 770)
	$label5.Size = New-Object System.Drawing.Size (70, 20)
	$label5.Text = "Djokovic: "
	$label5.TabStop = $False
	$label5.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$form1.Controls.Add($label5)

	$label6 = New-Object System.Windows.Forms.Label
	$label6.Location = New-Object System.Drawing.Size (210, 770)
	$label6.Size = New-Object System.Drawing.Size (201, 20)
	$label6.Text = "HR, Markerting, MIS, MGMT,IT, Payroll"
	$label6.TabStop = $False
	$form1.Controls.Add($label6)

	$label7 = New-Object System.Windows.Forms.Label
	$label7.Location = New-Object System.Drawing.Size (410, 770)
	$label7.Size = New-Object System.Drawing.Size (90, 20)
	$label7.Text = "Wawrinka: "
	$label7.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$label7.TabStop = $False
	$form1.Controls.Add($label7)

	$label8 = New-Object System.Windows.Forms.Label
	$label8.Location = New-Object System.Drawing.Size (470, 770)
	$label8.Size = New-Object System.Drawing.Size (280, 20)
	$label8.Text = "Tech Support, Warehouse, RMA, LA, Data Entry"
	$label8.TabStop = $False
	$form1.Controls.Add($label8)

	$label9 = New-Object System.Windows.Forms.Label
	$label9.Location = New-Object System.Drawing.Size (15, 790)
	$label9.Size = New-Object System.Drawing.Size (60, 20)
	$label9.Text = "Raonic: "
	$label9.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$label9.TabStop = $False
	$form1.Controls.Add($label9)

	$label10 = New-Object System.Windows.Forms.Label
	$label10.Location = New-Object System.Drawing.Size (73, 790)
	$label10.Size = New-Object System.Drawing.Size (100, 20)
	$label10.Text = "CH, GA, MI, NJ"
	$label10.TabStop = $False
	$form1.Controls.Add($label10)

	$label11 = New-Object System.Windows.Forms.Label
	$label11.Location = New-Object System.Drawing.Size (173, 790)
	$label11.Size = New-Object System.Drawing.Size (53, 20)
	$label11.Text = "Gasquet: "
	$label11.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$label11.TabStop = $False
	$form1.Controls.Add($label11)

	$label12 = New-Object System.Windows.Forms.Label
	$label12.Location = New-Object System.Drawing.Size (230, 790)
	$label12.Size = New-Object System.Drawing.Size (30, 20)
	$label12.Text = "SBZ"
	$label12.TabStop = $False
	$form1.Controls.Add($label12)

	$label13 = New-Object System.Windows.Forms.Label
	$label13.Location = New-Object System.Drawing.Size (273, 790)
	$label13.Size = New-Object System.Drawing.Size (60, 20)
	$label13.Text = "Jarguar: "
	$label13.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$label13.TabStop = $False
	$form1.Controls.Add($label13)

	$label14 = New-Object System.Windows.Forms.Label
	$label14.Location = New-Object System.Drawing.Size (333, 790)
	$label14.Size = New-Object System.Drawing.Size (50, 20)
	$label14.Text = "STT"
	$label14.TabStop = $False
	$form1.Controls.Add($label14)

	$label15 = New-Object System.Windows.Forms.Label
	$label15.Location = New-Object System.Drawing.Size (383, 790)
	$label15.Size = New-Object System.Drawing.Size (65, 20)
	$label15.Text = "Sharapova: "
	$label15.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$label15.TabStop = $False
	$form1.Controls.Add($label15)

	$label16 = New-Object System.Windows.Forms.Label
	$label16.Location = New-Object System.Drawing.Size (453, 790)
	$label16.Size = New-Object System.Drawing.Size (70, 20)
	$label16.Text = "AP, AR, CR"
	$label16.TabStop = $False
	$form1.Controls.Add($label16)

	$label17 = New-Object System.Windows.Forms.Label
	$label17.Location = New-Object System.Drawing.Size (523, 790)
	$label17.Size = New-Object System.Drawing.Size (65, 20)
	$label17.Text = "Sprintbok: "
	$label17.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$label17.TabStop = $False
	$form1.Controls.Add($label17)

	$label16 = New-Object System.Windows.Forms.Label
	$label16.Location = New-Object System.Drawing.Size (588, 790)
	$label16.Size = New-Object System.Drawing.Size (100, 20)
	$label16.Text = "Wuhan"
	$label16.TabStop = $False
	$form1.Controls.Add($label16)

	#endregion Generated Form Code 

	#Save the initial state of the form 
	$InitialFormWindowState = $form1.WindowState
	#Init the OnLoad event to correct the initial state of the form 
	$Form1.add_Load($OnLoadForm_StateCorrection)

	$Form1.Topmost = $True
	$Form1.Add_Shown({

		$Form1.Activate()

	})
	[void]$Form1.ShowDialog()

}

################################################################################
# Form 2
################################################################################
function ShowForm2 {

	$Form1.visible = $False

	$Form2 = New-Object System.Windows.Forms.Form
	$Form2.Text = $title
	$Form2.Size = New-Object System.Drawing.Size (900, 710)
	$Form2.StartPosition = "CenterScreen"
	$Form2.KeyPreview = $True
	$Form2.Add_KeyDown({

		if ($_.KeyCode -eq "Enter") {

			if ($OKButton2.visible) {

				if ($disable) {

					$logfile = $UserAlias + "-disabled.log"
					DisableUser

				} else {

					$logfile = $UserAlias + "-created.log"
					CreateMailbox

				}

			}

			if ($FinishButton2.visible) {

				$Form1.close(); 
				$Form2.close(); 
				$launchForm.visible = $True

			}

		}

	})

	$Form2.Add_KeyDown({

		if ($_.KeyCode -eq "Escape") {

			$launchForm.close();
			$Form1.close();
			$Form2.close()
		
		}

	})

	if ($UserExtension) {

		$UserExtension = "x" + $UserExtension

	}

	$OKButton2 = New-Object System.Windows.Forms.Button
	$OKButton2.Location = New-Object System.Drawing.Size (680, 500)
	$OKButton2.Size = New-Object System.Drawing.Size (75, 23)
	$OKButton2.Text = "OK"
	$OKButton2.TabIndex = 2
	$OKButton2.Add_Click({

		if ($disable) {

			$logfile = $UserAlias + "-disabled.log"
			DisableUser

		} else {

			$logfile = $UserAlias + "-created.log"
			$CheckIT = $ITemailcheck.Checked
			$CheckHR = $HRemailcheck.Checked
			$CheckBranch = $Branchemailcheck.Checked
			CreateMailbox

			Start-Sleep -s 1

			if ($GroupName_DFS.Length -ne 0 -or $GroupName_Vasto.Length -ne 0 -or $GroupName_Dept.Length -ne 0 -or $GroupName_List.Length -ne 0) {

				Add-Group

			}

		}

	})

	$Form2.Controls.Add($OKButton2)

	$CancelButton2 = New-Object System.Windows.Forms.Button
	$CancelButton2.Location = New-Object System.Drawing.Size (765, 500)
	$CancelButton2.Size = New-Object System.Drawing.Size (75, 23)
	$CancelButton2.Text = "Cancel"
	$CancelButton2.TabIndex = 3
	$CancelButton2.Add_Click({

		$launchForm.close();
		$Form1.close();
		$Form2.close()

	})
	$Form2.Controls.Add($CancelButton2)

	$BackButton2 = New-Object System.Windows.Forms.Button
	$BackButton2.Location = New-Object System.Drawing.Size (605, 500)
	$BackButton2.Size = New-Object System.Drawing.Size (75, 23)
	$BackButton2.Text = "< Back"
	$BackButton2.TabIndex = 4
	$BackButton2.Add_Click({

		$Form2.visible = $False; 

		if ($disable) {

			$launchForm.visible = $True

		} else {

			$Form1.visible = $True

		}

	})
	$Form2.Controls.Add($BackButton2)

	$FinishButton2 = New-Object System.Windows.Forms.Button
	$FinishButton2.Location = New-Object System.Drawing.Size (680, 500)
	$FinishButton2.Size = New-Object System.Drawing.Size (75, 23)
	$FinishButton2.Text = "Finish"
	$FinishButton2.TabIndex = 1
	$FinishButton2.visible = $False
	$FinishButton2.Add_Click({

			$Form2.Refresh()
			Start-Sleep -s 1
			$Form1.close();
			$Form2.close();
			$launchForm.visible = $True

	})
	$Form2.Controls.Add($FinishButton2)

	# Job Status Label
	$JobStatusLabel = New-Object System.Windows.Forms.Label
	$JobStatusLabel.Location = New-Object System.Drawing.Size (420, 40)
	$JobStatusLabel.Size = New-Object System.Drawing.Size (120, 20)
	$JobStatusLabel.Text = "Job Status:"
	$JobStatusLabel.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$Form2.Controls.Add($JobStatusLabel)

	# Error Message Box
	$ErrorMsg2 = New-Object System.Windows.Forms.Label
	$ErrorMsg2.Location = New-Object System.Drawing.Size (420, 70)
	$ErrorMsg2.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg2.Text = $null
	$ErrorMsg2.ForeColor = "Green"
	$Form2.Controls.Add($ErrorMsg2)

	$ErrorMsg3 = New-Object System.Windows.Forms.Label
	$ErrorMsg3.Location = New-Object System.Drawing.Size (420, 90)
	$ErrorMsg3.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg3.Text = $null
	$ErrorMsg3.ForeColor = "Green"
	$Form2.Controls.Add($ErrorMsg3)

	$ErrorMsg4 = New-Object System.Windows.Forms.Label
	$ErrorMsg4.Location = New-Object System.Drawing.Size (420, 110)
	$ErrorMsg4.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg4.Text = $null
	$ErrorMsg4.ForeColor = "Green"
	$Form2.Controls.Add($ErrorMsg4)

	$ErrorMsg5 = New-Object System.Windows.Forms.Label
	$ErrorMsg5.Location = New-Object System.Drawing.Size (420, 130)
	$ErrorMsg5.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg5.Text = $null
	$ErrorMsg5.ForeColor = "Green"
	$Form2.Controls.Add($ErrorMsg5)

	$ErrorMsg6 = New-Object System.Windows.Forms.Label
	$ErrorMsg6.Location = New-Object System.Drawing.Size (420, 150)
	$ErrorMsg6.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg6.Text = $null
	$ErrorMsg6.ForeColor = "Green"
	$Form2.Controls.Add($ErrorMsg6)

	$ErrorMsg7 = New-Object System.Windows.Forms.Label
	$ErrorMsg7.Location = New-Object System.Drawing.Size (420, 170)
	$ErrorMsg7.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg7.Text = $null
	$ErrorMsg7.ForeColor = "Green"
	$Form2.Controls.Add($ErrorMsg7)

	$ErrorMsg8 = New-Object System.Windows.Forms.Label
	$ErrorMsg8.Location = New-Object System.Drawing.Size (420, 190)
	$ErrorMsg8.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg8.Text = $null
	$ErrorMsg8.ForeColor = "Green"
	$Form2.Controls.Add($ErrorMsg8)

	$ErrorMsg9 = New-Object System.Windows.Forms.Label
	$ErrorMsg9.Location = New-Object System.Drawing.Size (420, 210)
	$ErrorMsg9.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg9.Text = $null
	$ErrorMsg9.ForeColor = "Green"
	$Form2.Controls.Add($ErrorMsg9)

	$ErrorMsg10 = New-Object System.Windows.Forms.Label
	$ErrorMsg10.Location = New-Object System.Drawing.Size (420, 230)
	$ErrorMsg10.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg10.Text = $null
	$ErrorMsg10.ForeColor = "Green"
	$Form2.Controls.Add($ErrorMsg10)

	$ErrorMsg11 = New-Object System.Windows.Forms.Label
	$ErrorMsg11.Location = New-Object System.Drawing.Size (420, 250)
	$ErrorMsg11.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg11.Text = $null
	$ErrorMsg11.ForeColor = "Green"
	$Form2.Controls.Add($ErrorMsg11)

	$ErrorMsg12 = New-Object System.Windows.Forms.Label
	$ErrorMsg12.Location = New-Object System.Drawing.Size (420, 270)
	$ErrorMsg12.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg12.Text = $null
	$ErrorMsg12.ForeColor = "Green"
	$Form2.Controls.Add($ErrorMsg12)

	$ErrorMsg13 = New-Object System.Windows.Forms.Label
	$ErrorMsg13.Location = New-Object System.Drawing.Size (420, 290)
	$ErrorMsg13.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg13.Text = $null
	$ErrorMsg13.ForeColor = "Green"
	$Form2.Controls.Add($ErrorMsg13)

	$ErrorMsg14 = New-Object System.Windows.Forms.Label
	$ErrorMsg14.Location = New-Object System.Drawing.Size (420, 310)
	$ErrorMsg14.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg14.Text = $null
	$ErrorMsg14.ForeColor = "Green"
	$Form2.Controls.Add($ErrorMsg14)

	$ErrorMsg15 = New-Object System.Windows.Forms.Label
	$ErrorMsg15.Location = New-Object System.Drawing.Size (420, 330)
	$ErrorMsg15.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg15.Text = $null
	$ErrorMsg15.ForeColor = "Green"
	$Form2.Controls.Add($ErrorMsg15)

	$ErrorMsg16 = New-Object System.Windows.Forms.Label
	$ErrorMsg16.Location = New-Object System.Drawing.Size (420, 350)
	$ErrorMsg16.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg16.Text = $null
	$ErrorMsg16.ForeColor = "Green"
	$Form2.Controls.Add($ErrorMsg16)

	$ErrorMsg17 = New-Object System.Windows.Forms.Label
	$ErrorMsg17.Location = New-Object System.Drawing.Size (420, 370)
	$ErrorMsg17.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg17.Text = $null
	$ErrorMsg17.ForeColor = "Green"
	$Form2.Controls.Add($ErrorMsg17)

	$ErrorMsg18 = New-Object System.Windows.Forms.Label
	$ErrorMsg18.Location = New-Object System.Drawing.Size (420, 390)
	$ErrorMsg18.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg18.Text = $null
	$ErrorMsg18.ForeColor = "Green"
	$Form2.Controls.Add($ErrorMsg18)

	$ErrorMsg19 = New-Object System.Windows.Forms.Label
	$ErrorMsg19.Location = New-Object System.Drawing.Size (420, 410)
	$ErrorMsg19.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg19.Text = $null
	$ErrorMsg19.ForeColor = "Green"
	$Form2.Controls.Add($ErrorMsg19)

	$ErrorMsg20 = New-Object System.Windows.Forms.Label
	$ErrorMsg20.Location = New-Object System.Drawing.Size (420, 430)
	$ErrorMsg20.Size = New-Object System.Drawing.Size (500, 20)
	$ErrorMsg20.Text = $null
	$ErrorMsg20.ForeColor = "Green"
	$Form2.Controls.Add($ErrorMsg20)

	$TitleLabel = New-Object System.Windows.Forms.Label
	$TitleLabel.Location = New-Object System.Drawing.Size (20, 40)
	$TitleLabel.Size = New-Object System.Drawing.Size (500, 20)
	$TitleLabel.Text = "Click OK to " + $title.ToLower() + "."
	$TitleLabel.Font = New-Object System.Drawing.Font ("Arial", 8, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)), [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
	$Form2.Controls.Add($TitleLabel)

	# OU Label
	$OULabel = New-Object System.Windows.Forms.Label
	$OULabel.Location = New-Object System.Drawing.Size (20, 70)
	$OULabel.Size = New-Object System.Drawing.Size (120, 20)
	$OULabel.Text = "Organizational Unit:"
	$Form2.Controls.Add($OULabel)

	$OULabel = New-Object System.Windows.Forms.Label
	$OULabel.Location = New-Object System.Drawing.Size (150, 70)
	$OULabel.Size = New-Object System.Drawing.Size (350, 20)
	$OULabel.Text = $UserOU
	$Form2.Controls.Add($OULabel)

	# Name Label
	$NameLabel = New-Object System.Windows.Forms.Label
	$NameLabel.Location = New-Object System.Drawing.Size (20, 100)
	$NameLabel.Size = New-Object System.Drawing.Size (120, 20)
	$NameLabel.Text = "Name:"
	$Form2.Controls.Add($NameLabel)

	$NameLabel = New-Object System.Windows.Forms.Label
	$NameLabel.Location = New-Object System.Drawing.Size (150, 100)
	$NameLabel.Size = New-Object System.Drawing.Size (350, 20)
	$NameLabel.Font = New-Object System.Drawing.Font ("ArialNarrow", 8)
	$NameLabel.Text = "$UserFirst $UserLast"
	$Form2.Controls.Add($NameLabel)

	# User Logon Name Label
	$LogonLabel = New-Object System.Windows.Forms.Label
	$LogonLabel.Location = New-Object System.Drawing.Size (20,130)
	$LogonLabel.Size = New-Object System.Drawing.Size (120,20)
	$LogonLabel.Font = New-Object System.Drawing.Font ("ArialNarrow",8)
	$LogonLabel.Text = "User Logon Name:"
	$Form2.Controls.Add($LogonLabel)

	$LogonLabel = New-Object System.Windows.Forms.Label
	$LogonLabel.Location = New-Object System.Drawing.Size (150,130)
	$LogonLabel.Size = New-Object System.Drawing.Size (350,20)
	$LogonLabel.Text = "$UserAlias$UserDomain"
	$Form2.Controls.Add($LogonLabel)

	# Password Label
	$PasswordLabel = New-Object System.Windows.Forms.Label
	$PasswordLabel.Location = New-Object System.Drawing.Size (20,160)
	$PasswordLabel.Size = New-Object System.Drawing.Size (120,20)
	$PasswordLabel.Text = "Password:"
	$Form2.Controls.Add($PasswordLabel)

	$PasswordLabel = New-Object System.Windows.Forms.Label
	$PasswordLabel.Location = New-Object System.Drawing.Size (150,160)
	$PasswordLabel.Size = New-Object System.Drawing.Size (200,20)
	$PasswordLabel.Text = "$UserPassword"
	$Form2.Controls.Add($PasswordLabel)

	# Server Label
	$ServerLabel = New-Object System.Windows.Forms.Label
	$ServerLabel.Location = New-Object System.Drawing.Size (20,190)
	$ServerLabel.Size = New-Object System.Drawing.Size (120,20)
	$ServerLabel.Text = "Server / Database:"
	$Form2.Controls.Add($ServerLabel)

	$ServerLabel = New-Object System.Windows.Forms.Label
	$ServerLabel.Location = New-Object System.Drawing.Size (150,190)
	$ServerLabel.Size = New-Object System.Drawing.Size (350,20)
	if ($MailboxServer -and $Database) { 

		$ServerLabel.Text = "$MailboxServer/$Database"

	}
	$Form2.Controls.Add($ServerLabel)

	# Email Label
	$EmailLabel = New-Object System.Windows.Forms.Label
	$EmailLabel.Location = New-Object System.Drawing.Size (20,220)
	$EmailLabel.Size = New-Object System.Drawing.Size (130,20)
	$EmailLabel.Text = "E-mail:"
	$Form2.Controls.Add($EmailLabel)

	$EmailLabel = New-Object System.Windows.Forms.Label
	$EmailLabel.Location = New-Object System.Drawing.Size (150,220)
	$EmailLabel.Size = New-Object System.Drawing.Size (350,20)
	$EmailLabel.Font = New-Object System.Drawing.Font ("ArialNarrow",8)
	$EmailLabel.Text = "$UserEmail$UserEmailDomain"
	$Form2.Controls.Add($EmailLabel)

	# Distribution Group Label
	$DGLabel = New-Object System.Windows.Forms.Label
	$DGLabel.Location = New-Object System.Drawing.Size (20,250)
	$DGLabel.Size = New-Object System.Drawing.Size (120,20)
	$DGLabel.Text = "Distribution Group:"
	$Form2.Controls.Add($DGLabel)

	$DGLabel = New-Object System.Windows.Forms.Label
	$DGLabel.Location = New-Object System.Drawing.Size (150,250)
	$DGLabel.Size = New-Object System.Drawing.Size (350,20)
	$DGLabel.Text = "$GroupName_DG"
	$Form2.Controls.Add($DGLabel)

	# Office Label
	$OfficeLabel = New-Object System.Windows.Forms.Label
	$OfficeLabel.Location = New-Object System.Drawing.Size (20,280)
	$OfficeLabel.Size = New-Object System.Drawing.Size (120,20)
	$OfficeLabel.Text = "Office:"
	$Form2.Controls.Add($OfficeLabel)

	$OfficeLabel = New-Object System.Windows.Forms.Label
	$OfficeLabel.Location = New-Object System.Drawing.Size (150,280)
	$OfficeLabel.Size = New-Object System.Drawing.Size (350,20)
	$OfficeLabel.Text = "$UserOffice"
	$Form2.Controls.Add($OfficeLabel)

	# Phone Label
	$PhoneLabel = New-Object System.Windows.Forms.Label
	$PhoneLabel.Location = New-Object System.Drawing.Size (20,310)
	$PhoneLabel.Size = New-Object System.Drawing.Size (120,20)
	$PhoneLabel.Text = "Phone:"
	$Form2.Controls.Add($PhoneLabel)

	# Phone Extension
	$PhoneLabel = New-Object System.Windows.Forms.Label
	$PhoneLabel.Location = New-Object System.Drawing.Size (150,310)
	$PhoneLabel.Size = New-Object System.Drawing.Size (350,20)
	$PhoneLabel.Text = "$UserPhone $UserExtension".Trim()
	$Form2.Controls.Add($PhoneLabel)

	# Group Membership DFS
	$DFSLabel = New-Object System.Windows.Forms.Label
	$DFSLabel.Location = New-Object System.Drawing.Size (20,340)
	$DFSLabel.Size = New-Object System.Drawing.Size (120,20)
	$DFSLabel.Text = "DFS Group:"
	$Form2.Controls.Add($DFSLabel)

	# Group Membership DFS
	$DFSLabel = New-Object System.Windows.Forms.Label
	$DFSLabel.Location = New-Object System.Drawing.Size (150,340)
	$DFSLabel.Size = New-Object System.Drawing.Size (350,20)
	$DFSLabel.Text = "$GroupName_DFS".Trim()
	$Form2.Controls.Add($DFSLabel)

	# Group Membership Vasto
	$DFSLabel = New-Object System.Windows.Forms.Label
	$DFSLabel.Location = New-Object System.Drawing.Size (20,370)
	$DFSLabel.Size = New-Object System.Drawing.Size (120,20)
	$DFSLabel.Text = "Vasto Group:"
	$Form2.Controls.Add($DFSLabel)

	# Group Membership Vasto
	$DFSLabel = New-Object System.Windows.Forms.Label
	$DFSLabel.Location = New-Object System.Drawing.Size (150,370)
	$DFSLabel.Size = New-Object System.Drawing.Size (350,20)
	$DFSLabel.Text = "$GroupName_Vasto".Trim()
	$Form2.Controls.Add($DFSLabel)

	# Group Membership Department
	$DeptLabel = New-Object System.Windows.Forms.Label
	$DeptLabel.Location = New-Object System.Drawing.Size (20,400)
	$DeptLabel.Size = New-Object System.Drawing.Size (130,20)
	$DeptLabel.Text = "Websense Department:"
	$Form2.Controls.Add($DeptLabel)

	# Group Membership Department
	$DeptLabel = New-Object System.Windows.Forms.Label
	$DeptLabel.Location = New-Object System.Drawing.Size (150,400)
	$DeptLabel.Size = New-Object System.Drawing.Size (550,20)
	$DeptLabel.Text = "$GroupName_Dept".Trim()
	$Form2.Controls.Add($DeptLabel)

	# Group Membership Vasto
	$WebsenseLabel = New-Object System.Windows.Forms.Label
	$WebsenseLabel.Location = New-Object System.Drawing.Size (20,430)
	$WebsenseLabel.Size = New-Object System.Drawing.Size (130,20)
	$WebsenseLabel.Text = "Websense List:"
	$Form2.Controls.Add($WebsenseLabel)

	# Group Membership Vasto
	$WebsenseLabel = New-Object System.Windows.Forms.Label
	$WebsenseLabel.Location = New-Object System.Drawing.Size (150,430)
	$WebsenseLabel.Size = New-Object System.Drawing.Size (350,20)
	$WebsenseLabel.Text = "$GroupName_List".Trim()
	$Form2.Controls.Add($WebsenseLabel)

	# Mailbox Features Label
	#$FeaturesLabel = New-Object System.Windows.Forms.Label
	#$FeaturesLabel.Location = New-Object System.Drawing.Size(20,340)
	#$FeaturesLabel.Size = New-Object System.Drawing.Size(120,20) 
	#$FeaturesLabel.Text = "Mailbox Features:"
	#$Form2.Controls.Add($FeaturesLabel)

	#$FeaturesLabel = New-Object System.Windows.Forms.Label
	#$FeaturesLabel.Location = New-Object System.Drawing.Size(150,340)
	#$FeaturesLabel.Size = New-Object System.Drawing.Size(350,20)
	#if ($UserOutlookAnywhere) {
	#	$FeaturesLabel.Text = "Outlook Anywhere: Enabled"
	#} else {
	#	$FeaturesLabel.Text = "Outlook Anywhere: Disabled"
	#}
	#$Form2.Controls.Add($FeaturesLabel)

	#$FeaturesLabel = New-Object System.Windows.Forms.Label
	#$FeaturesLabel.Location = New-Object System.Drawing.Size(150,360)
	#$FeaturesLabel.Size = New-Object System.Drawing.Size(350,20)
	#if ($UserOWA) {
	#	$FeaturesLabel.Text = "Outlook Web Access: Enabled"
	#} else {
	#	$FeaturesLabel.Text = "Outlook Web Access: Disabled"
	#}
	#$Form2.Controls.Add($FeaturesLabel)

	#$FeaturesLabel = New-Object System.Windows.Forms.Label
	#$FeaturesLabel.Location = New-Object System.Drawing.Size(150,380)
	#$FeaturesLabel.Size = New-Object System.Drawing.Size(350,20)
	#if ($UserActiveSync) {
	#	$FeaturesLabel.Text = "Exchange ActiveSync: Enabled"
	#} else {
	#	$FeaturesLabel.Text = "Exchange ActiveSync: Disabled"
	#}
	#$Form2.Controls.Add($FeaturesLabel)

	# The check function Label
	$CheckLabel1 = New-Object System.Windows.Forms.Label
	$CheckLabel1.Location = New-Object System.Drawing.Size (20,473)
	$CheckLabel1.Size = New-Object System.Drawing.Size (60,20)
	$CheckLabel1.Text = "Email To:"
	$Form2.Controls.Add($CheckLabel1)

	# EmailToIT Label
	$EmailCheckLabel1 = New-Object System.Windows.Forms.Label
	$EmailCheckLabel1.Location = New-Object System.Drawing.Size (95,473)
	$EmailCheckLabel1.Size = New-Object System.Drawing.Size (50,20)
	$EmailCheckLabel1.Text = "IT Admin"
	$Form2.Controls.Add($EmailCheckLabel1)

	# IT Email Check Box
	$ITemailcheck = New-Object System.Windows.Forms.CheckBox
	$ITemailcheck.Location = New-Object System.Drawing.Size (80,470)
	$ITemailcheck.Size = New-Object System.Drawing.Size (40,20)
	$ITemailcheck.Checked = $True
	$ITemailcheck.TabIndex = 14
	$ITemailcheck.Enabled = $True
	$Form2.Controls.Add($ITemailcheck)

	# EmailToHR Label
	$EmailCheckLabel2 = New-Object System.Windows.Forms.Label
	$EmailCheckLabel2.Location = New-Object System.Drawing.Size (165,473)
	$EmailCheckLabel2.Size = New-Object System.Drawing.Size (30,20)
	$EmailCheckLabel2.Text = "HR"
	$Form2.Controls.Add($EmailCheckLabel2)

	# HR Email Check Box
	$HRemailcheck = New-Object System.Windows.Forms.CheckBox
	$HRemailcheck.Location = New-Object System.Drawing.Size (150,470)
	$HRemailcheck.Size = New-Object System.Drawing.Size (40,20)
	$HRemailcheck.TabIndex = 14
	$Form2.Controls.Add($HRemailcheck)

	if ($UserAlias -like "*w_*") {

		$EmailCheckLabel2.Enabled = $True
		$HRemailcheck.Checked = $True
		$HRemailcheck.Enabled = $True

	} else {

		$EmailCheckLabel2.Enabled = $False
		$HRemailcheck.Checked = $False
		$HRemailcheck.Enabled = $False

	}

	$Branchemailcheck = New-Object System.Windows.Forms.CheckBox
	$Branchemailcheck.Location = New-Object System.Drawing.Size (80,495)
	$Branchemailcheck.Size = New-Object System.Drawing.Size (15,20)
	$Branchemailcheck.Enabled = $True
	$Form2.Controls.Add($Branchemailcheck)

	switch ($UserAlias) {

		{ ($_ -like "*w_*") } { 

			$emailToBranch = "admin.wh@newbiiz.com"
			$Branchemailcheck.Checked = $True

		}
		
		{ ($_ -like "*c_*") } { 

			$emailToBranch = "Brian.Li@malabs.com"
			$Branchemailcheck.Checked = $True

		}

		{ ($_ -like "*g_*") } { 
		
			$emailToBranch = "Christina.Tay@malabs.com"
			$Branchemailcheck.Checked = $True

		}

		{ ($_ -like "*i_*") } { 

			$emailToBranch = "Keith.Yarbrough@malabs.com"
			$Branchemailcheck.Checked = $True

		}

		{ ($_ -like "*n_*") } {

			$emailToBranch = "davidc@malabs.com"
			$Branchemailcheck.Checked = $True

		}

		{ ($_ -like "*m_*") } { 

			$emailToBranch = "fari@malabs.com"
			$Branchemailcheck.Checked = $True

		}

		{ ($_ -notlike '*_*') } {

			$Branchemailcheck.Enabled = $False
			$Branchemailcheck.Checked = $False

		}

	}

	# EmailToWuHan Label
	$EmailCheckLabel3 = New-Object System.Windows.Forms.Label
	$EmailCheckLabel3.Location = New-Object System.Drawing.Size (95,498)
	$EmailCheckLabel3.Size = New-Object System.Drawing.Size (220,20)
	$EmailCheckLabel3.Text = "$emailToBranch"
	$EmailCheckLabel3.Enabled = $True
	$Form2.Controls.Add($EmailCheckLabel3)

	$Form2.Topmost = $True
	$Form2.Add_Shown({ 
	
		$Form2.Activate()
	
	})
	[void]$Form2.ShowDialog()

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