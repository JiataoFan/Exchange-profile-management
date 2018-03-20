<#
	Mailbox form
#>
function ShowForm1 {

	param([bool]$new, [string]$title, [bool]$disable, [bool]$found)

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

									$UserPhone = $telephone[0].Trim()
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

									$Phone.Text = $telephone[0].Trim()
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
		Config "Next" button and behaviors
	#>
	$nextButton = Config-Button -horizontalPosition 960 -verticalPosition 770 -width 105 -height 43 -text "Next >" -tabIndex 17

	$nextButton.Add_Click({

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

	$Form1.Controls.Add($nextButton)

	<#
		Config "Cancel" button and its behaviors
	#>
	$cancelButton = Config-Button -horizontalPosition 1065 -verticalPosition 770 -width 105 -height 43 -text "Cancel" -tabIndex 18
	$cancelButton.Add_Click({

		$Form1.close()

	})
	$Form1.Controls.Add($cancelButton)

	<#
		Config "Back" button and its behaviors
	#>	
	$backButton = Config-Button -horizontalPosition 855 -verticalPosition 770 -width 105 -height 43 -text "< Back" -tabIndex 19
	$backButton.Add_Click({

		$Form1.visible = $False
		$launchForm.visible = $True

	})
	$Form1.Controls.Add($backButton)

	<#
		Config error message
	#>
	$errorMessage = Config-ErrorMessage -horizontalPosition 20 -verticalPosition 10 -width 500 -height 20
	$Form1.Controls.Add($errorMessage)

	<#
		Config "Company/Office/Dept.:" label
	#>
	$officeLabel = Config-Label -horizontalPosition 20 -verticalPosition 70 -width 140 -height 20 -text	"Company/Office/Dept.:" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$Form1.Controls.Add($officeLabel)

	<# 
		Config title label
	#>
	$titleLabel = Config-Label -horizontalPosition 20 -verticalPosition 40 -width 450 -height 20 -text	"" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$Form1.Controls.Add($titleLabel)

	<#
		Config "Organizational Unit:" label
	#>
	$organizationalUnitLabel = Config-Label -horizontalPosition 20 -verticalPosition 100 -width 120 -height 20 -text "Organizational Unit:" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$Form1.Controls.Add($organizationalUnitLabel)

	<#
		Config "First Name/Last Name:" label
	#>
	$firstNameLastNameLabel = Config-Label -horizontalPosition 20 -verticalPosition 130 -width 130 -height 20 -text "First Name/Last Name:" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$Form1.Controls.Add($firstNameLastNameLabel)

	<#
		Config "User Logon Name:" label
	#>
	$logonNameLabel = Config-Label -horizontalPosition 20 -verticalPosition 160 -width 120 -height 20 -text "User Logon Name:" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$Form1.Controls.Add($logonNameLabel)

	<#
		Config "Password: " label
	#>
	$passwordLabel = Config-Label -horizontalPosition 20 -verticalPosition 190 -width 120 -height 20 -text "Password: " -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$Form1.Controls.Add($passwordLabel)

	$passwordLimitLabel = Config-Label -horizontalPosition 340 -verticalPosition 195 -width 140 -height 20 -text "(minimum 5 characters)" -fontFamily "Arial" -fontSize 7
	$Form1.Controls.Add($passwordLimitLabel)

	<#
		Config "Server / Info Store:" label
	#>
	$serverInfoStoreLabel = Config-Label -horizontalPosition 20 -verticalPosition 220 -width 120 -height 20 -text "Server / Info Store:" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$Form1.Controls.Add($serverInfoStoreLabel)

	<#
		Config "E-mail:" label
	#>
	$emailLabel = Config-Label -horizontalPosition 20 -verticalPosition 540 -width 120 -height 20 -text "E-mail:" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$Form1.Controls.Add($emailLabel)

	<#
		Config "Distribution Group:" label
	#>
	$distributionGroupLabel = Config-Label -horizontalPosition 595 -verticalPosition 60 -width 120 -height 20 -text "Distribution Group:" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$Form1.Controls.Add($distributionGroupLabel)

	<#
		Config "Phone / Extension:" label
	#>
	$phoneExtensionLabel1 = Config-Label -horizontalPosition 20 -verticalPosition 570 -width 120 -height 15 -text "Phone / Extension:" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$Form1.Controls.Add($phoneExtensionLabel1)

	$phoneExtensionLabel2 = Config-Label -horizontalPosition 190 -verticalPosition 590 -width 150 -height 20 -text "e.g. 408-941-0808" -fontFamily "Arial" -fontSize 7
	$Form1.Controls.Add($phoneExtensionLabel2)

	$phoneExtensionLabel3 = Config-Label -horizontalPosition 340 -verticalPosition 590 -width 140 -height 20 -text "e.g. 808 " -fontFamily "Arial" -fontSize 7
	$Form1.Controls.Add($phoneExtensionLabel3)

	<#
		Config "Mailbox Features:" label
	#>
	$mailboxFeaturesLabel = Config-Label -horizontalPosition 420 -verticalPosition 370 -width 120 -height 15 -text "Mailbox Features:" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	#$Form1.Controls.Add($mailboxFeaturesLabel)

	<#
		Config "( Approve Needed )" label
	#>
	$approveNeededLabel = Config-Label -horizontalPosition 420 -verticalPosition 390 -width 120 -height 15 -text "( Approve Needed )"
	#$Form1.Controls.Add($approveNeededLabel)

	<#
		Config "Outlook Anywhere" label
	#>
	$outlookAnywhereLabel = Config-Label -horizontalPosition 565 -verticalPosition 350 -width 150 -height 20 -text "Outlook Anywhere"
	#$Form1.Controls.Add($OutlookAnywhereLabel)

	<#
		Config "Outlook Web Access (OWA)" label
	#>
	$outlookWebAccess = Config-Label -horizontalPosition 565 -verticalPosition 370 -width 150 -height 20 -text "Outlook Web Access (OWA)"
	#$Form1.Controls.Add($outlookWebAccess)

	<#
		"Exchange ActiveSync" label
	#>
	$exchangeActiveSyncLabel = Config-Label -horizontalPosition 565 -verticalPosition 390 -width 120 -height 20 -text "Exchange ActiveSync"
	#$Form1.Controls.Add($exchangeActiveSyncLabel)

	<#
		Config "Company" combo box and its behaviors
	#>
	$MainCompany = New-Object System.Windows.Forms.ComboBox
	$MainCompany.Location = New-Object System.Drawing.Size (190, 70)
	$MainCompany.Size = New-Object System.Drawing.Size (90, 20)

	$MainCompany.Items.Add("")
	$MainCompany.Items.Add("MA Labs")
	$MainCompany.Items.Add("Superbiiz")
	$MainCompany.Items.Add("Supertalent")
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
	$organizationalUnitMessageLabel = Config-Label -horizontalPosition 450 -verticalPosition 105 -width 20 -height 20
	if ($new) {

		$organizationalUnitMessageLabel.Text = "*"

	}
	$organizationalUnitMessageLabel.ForeColor = "Red"
	$Form1.Controls.Add($organizationalUnitMessageLabel)

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
	if (!$new) {

		$FirstName.Enabled = $False

	}
	#$Form1.Controls.Add($FirstName)

	<#
		First Name Message Box
	#>
	$firstNameMessageLabel = Config-Label -horizontalPosition 490 -verticalPosition 135 -width 20 -height 20
	if ($new) {

		$firstNameMessageLabel.Text = "*"

	}
	$firstNameMessageLabel.ForeColor = "Red"
	$Form1.Controls.Add($firstNameMessageLabel)

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
		Logon Name Message label
	#>
	$logonNameMessageLabel = Config-Label -horizontalPosition 490 -verticalPosition 165 -width 20 -height 20 -text "*"
	$logonNameMessageLabel.ForeColor = "Red"
	$Form1.Controls.Add($logonNameMessageLabel)

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
		Password Message label
	#>
	$passwordMessageLabel = Config-Label -horizontalPosition 490 -verticalPosition 195 -width 20 -height 20
	if ($new) {

		$passwordMessageLabel.Text = "*"

	}
	$passwordMessageLabel.ForeColor = "Red"
	$Form1.Controls.Add($passwordMessageLabel)

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
		Server Name Message label
	#>	
	$serverNameMessage = Config-Label -horizontalPosition 490 -verticalPosition 225 -width 20 -height 20
	if ($new) {

		$serverNameMessage.Text = "*"

	}
	$serverNameMessage.ForeColor = "Red"
	$Form1.Controls.Add($serverNameMessage)

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
	$EmailDomain.Items.Add("@ma.local")
	$EmailDomain.Items.Add("@superbiiz.com")
	$EmailDomain.Items.Add("@supertalent.com")
	$EmailDomain.DropDownStyle = 2
	$EmailDomain.Sorted = $True
	$EmailDomain.TabIndex = 12
	$EmailDomain.Enabled = $False
	$Form1.Controls.Add($EmailDomain)

	<#
		E-mail Message Box
	#>
	$emailMessageLabel = Config-Label -horizontalPosition 450 -verticalPosition 545 -width 20 -height 20
	$emailMessageLabel.ForeColor = "Red"
	$Form1.Controls.Add($emailMessageLabel)

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
		Phone Message label
	#>
	$phoneMessageLabel = Config-Label -horizontalPosition 450 -verticalPosition 575 -width 20 -height 20 -text "MS Lync Account:"	
	$phoneMessageLabel.ForeColor = "Red"
	$Form1.Controls.Add($phoneMessageLabel)

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
		Lync label
	#>
	$lyncLabel = Config-Label -horizontalPosition 20 -verticalPosition 615 -width 140 -height 20 -text "MS Lync Account:"
	#$Form1.Controls.Add($lyncLabel)

	<#
		Lync Check box
	#>
	$LyncCheck = New-Object System.Windows.Forms.CheckBox
	$LyncCheck.Location = New-Object System.Drawing.Size (190, 615)
	$LyncCheck.Size = New-Object System.Drawing.Size (140, 20)
	$LyncCheck.Checked = $False
	$LyncCheck.TabIndex = 18
	$LyncCheck.Enabled = $True
	#$Form1.Controls.Add($LyncCheck)

	<#
		Add Group Membership labels
	#>
	$networkDriveLabel = Config-Label -horizontalPosition 775 -verticalPosition 60 -width 85 -height 20 -text "Network Drive: "
	$Form1.Controls.Add($networkDriveLabel)

	$NDFSLabel = Config-Label -horizontalPosition 775 -verticalPosition 80 -width 85 -height 20 -text "NDFS: "
	$Form1.Controls.Add($NDFSLabel)

	$FSLabel = Config-Label -horizontalPosition 775 -verticalPosition 350 -width 85 -height 20 -text "FS: "
	$Form1.Controls.Add($FSLabel)

	$websenseDepartmentLabel = Config-Label -horizontalPosition 1000 -verticalPosition 80 -width 165 -height 20 -text "Websense Department: "
	$Form1.Controls.Add($websenseDepartmentLabel)

	$websenseUserDefineLabel = Config-Label -horizontalPosition 1000 -verticalPosition 350 -width 185 -height 20 -text "Websense User Define: "
	$Form1.Controls.Add($websenseUserDefineLabel)

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
	$strDomainDN = $strDNSDomain.ToString().Split('.')
	foreach ($strVal in $strDomainDN) { 
	
		$strTemp += "dc=$strVal," 
	
	}
	$strDomainDN = $strTemp.TrimEnd(",").ToLower()
	#endregion 

	#region Generated Form Objects 
	$treeView1 = New-Object System.Windows.Forms.TreeView
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

	<#
		"Berdych" server label
	#>
	$BerdychServerLabel = Config-Label -horizontalPosition 15 -verticalPosition 650 -width 50 -height 20 -text "Berdych: " -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	#$form1.Controls.Add($BerdychServerLabel)

	$HQPMDepartmentLabel = Config-Label -horizontalPosition 65 -verticalPosition 650 -width 100 -height 20 -text "HQ-PM"
	#$form1.Controls.Add($HQPMDepartmentLabel)

	$IsnerServerLabel = Config-Label -horizontalPosition 15 -verticalPosition 770 -width 50 -height 20 -text "Isner: " -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$form1.Controls.Add($IsnerServerLabel)

	$HQSalesDepartmentLabel = Config-Label -horizontalPosition 70 -verticalPosition 770 -width 70 -height 20 -text "HQ-Sales: "
	$form1.Controls.Add($HQSalesDepartmentLabels)

	$DjokovicServerLabel = Config-Label -horizontalPosition 140 -verticalPosition 770 -width 70 -height 20 -text "Djokovic: " -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$form1.Controls.Add($DjokovicServerLabel)

	$HRMarketingETCDepartmentLabel = Config-Label -horizontalPosition 210 -verticalPosition 770 -width 201 -height 20 -text "HR, Markerting, MIS, MGMT,IT, Payroll"
	$form1.Controls.Add($HRMarketingETCDepartmentLabel)

	$WawrinkaServerLabel = Config-Label -horizontalPosition 410 -verticalPosition 770 -width 90 -height 20 -text "Wawrinka: " -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$form1.Controls.Add($WawrinkaServerLabel)

	$TechSupportWarehouseETCDepartmentLabel = Config-Label -horizontalPosition 470 -verticalPosition 770 -width 280 -height 20 -text "Tech Support, Warehouse, RMA, LA, Data Entry"
	$form1.Controls.Add($TechSupportWarehouseETCDepartmentLabel)

	$RaonicServerLabel = Config-Label -horizontalPosition 15 -verticalPosition 790 -width 60 -height 20 -text "Wawrinka: " -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$form1.Controls.Add($RaonicServerLabel)

	$CHGAETCLocationLabel = Config-Label -horizontalPosition 73 -verticalPosition 790 -width 100 -height 20 -text "CH, GA, MI, NJ"
	$form1.Controls.Add($CHGAETCLocationLabel)

	$GasquetServerLabel = Config-Label -horizontalPosition 173 -verticalPosition 790 -width 53 -height 20 -text "Gasquet: " -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$form1.Controls.Add($GasquetServerLabel)

	$SBZLabel = Config-Label -horizontalPosition 230 -verticalPosition 790 -width 30 -height 20 -text "SBZ"
	$form1.Controls.Add($SBZLabel)

	$JarguarServerLabel = Config-Label -horizontalPosition 273 -verticalPosition 790 -width 60 -height 20 -text "Jarguar: " -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$form1.Controls.Add($JarguarServerLabel)

	$STTLabel = Config-Label -horizontalPosition 333 -verticalPosition 790 -width 50 -height 20 -text "STT"
	$form1.Controls.Add($STTLabel)

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

	#$Form1.Topmost = $True
	$Form1.Add_Shown({

		$Form1.Activate()

	})
	$Form1.ShowDialog()

}