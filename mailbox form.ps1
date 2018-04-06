<################################################################################
# New Mailbox form
################################################################################>


function Show-MailboxForm {

	param([string]$title, [bool]$new, [bool]$disable, [bool]$found)

	$launchForm.visible = $False

	<#
		Config New Mailbox form
	#>
	$newMailBoxForm = Config-Form -width 900 -height 810 -text $title -windowState "Maximized" -startPosition "CenterScreen" -keyPreview $True

	<#
		"Enter" key behavior
	#>
	$newMailBoxForm.Add_KeyDown({

		if ($_.KeyCode -eq "Enter") {

			 Handle-Enters

		} elseif ($_.KeyCode -eq "Escape") {

			$newMailBoxForm.close()

		}

	})

	<#
		Config "Next" button and behaviors
	#>
	$nextButton = Config-Button -horizontalPosition 960 -verticalPosition 770 -width 105 -height 43 -text "Next >" -tabIndex 17
	$nextButton.Add_Click({

			 Handle-Enters

		})
	$newMailBoxForm.Controls.Add($nextButton)

	<#
		Config "Cancel" button and its behaviors
	#>
	$cancelButton = Config-Button -horizontalPosition 1065 -verticalPosition 770 -width 105 -height 43 -text "Cancel" -tabIndex 18
	$cancelButton.Add_Click({

		$newMailBoxForm.close()

	})
	$newMailBoxForm.Controls.Add($cancelButton)

	<#
		Config "Back" button and its behaviors
	#>	
	$backButton = Config-Button -horizontalPosition 855 -verticalPosition 770 -width 105 -height 43 -text "< Back" -tabIndex 19
	$backButton.Add_Click({

		$newMailBoxForm.visible = $False
		$launchForm.visible = $True

	})
	$newMailBoxForm.Controls.Add($backButton)

	<#
		Config error message
	#>
	$errorMessage = Config-ErrorMessage -horizontalPosition 20 -verticalPosition 10 -width 500 -height 20
	$newMailBoxForm.Controls.Add($errorMessage)

	<#
		Config "Company/Office/Dept.:" label
	#>
	$officeLabel = Config-Label -horizontalPosition 20 -verticalPosition 70 -width 140 -height 20 -text	"Company/Office/Dept.:" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$newMailBoxForm.Controls.Add($officeLabel)

	<# 
		Config title label
	#>
	$titleLabel = Config-Label -horizontalPosition 20 -verticalPosition 40 -width 450 -height 20 -text	"" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$newMailBoxForm.Controls.Add($titleLabel)

	<#
		Config "Organizational Unit:" label
	#>
	$organizationalUnitLabel = Config-Label -horizontalPosition 20 -verticalPosition 100 -width 120 -height 20 -text "Organizational Unit:" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$newMailBoxForm.Controls.Add($organizationalUnitLabel)

	<#
		Config "First Name/Last Name:" label
	#>
	$firstNameLastNameLabel = Config-Label -horizontalPosition 20 -verticalPosition 130 -width 130 -height 20 -text "First Name/Last Name:" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$newMailBoxForm.Controls.Add($firstNameLastNameLabel)

	<#
		Config "User Logon Name:" label
	#>
	$logonNameLabel = Config-Label -horizontalPosition 20 -verticalPosition 160 -width 120 -height 20 -text "User Logon Name:" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$newMailBoxForm.Controls.Add($logonNameLabel)

	<#
		Config "Password: " label
	#>
	$passwordLabel = Config-Label -horizontalPosition 20 -verticalPosition 190 -width 120 -height 20 -text "Password: " -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$newMailBoxForm.Controls.Add($passwordLabel)

	$passwordLimitLabel = Config-Label -horizontalPosition 340 -verticalPosition 195 -width 140 -height 20 -text "(minimum 5 characters)" -fontFamily "Arial" -fontSize 7
	$newMailBoxForm.Controls.Add($passwordLimitLabel)

	<#
		Config "Server / Info Store:" label
	#>
	$serverInfoStoreLabel = Config-Label -horizontalPosition 20 -verticalPosition 220 -width 120 -height 20 -text "Server / Info Store:" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$newMailBoxForm.Controls.Add($serverInfoStoreLabel)

	<#
		Config "E-mail:" label
	#>
	$emailLabel = Config-Label -horizontalPosition 20 -verticalPosition 540 -width 120 -height 20 -text "E-mail:" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$newMailBoxForm.Controls.Add($emailLabel)

	<#
		Config "Distribution Group:" label
	#>
	$distributionGroupLabel = Config-Label -horizontalPosition 595 -verticalPosition 60 -width 120 -height 20 -text "Distribution Group:" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$newMailBoxForm.Controls.Add($distributionGroupLabel)

	<#
		Config "Phone / Extension:" label
	#>
	$phoneExtensionLabel1 = Config-Label -horizontalPosition 20 -verticalPosition 570 -width 120 -height 15 -text "Phone / Extension:" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$newMailBoxForm.Controls.Add($phoneExtensionLabel1)

	$phoneExtensionLabel2 = Config-Label -horizontalPosition 190 -verticalPosition 590 -width 150 -height 20 -text "e.g. 408-941-0808" -fontFamily "Arial" -fontSize 7
	$newMailBoxForm.Controls.Add($phoneExtensionLabel2)

	$phoneExtensionLabel3 = Config-Label -horizontalPosition 340 -verticalPosition 590 -width 140 -height 20 -text "e.g. 808 " -fontFamily "Arial" -fontSize 7
	$newMailBoxForm.Controls.Add($phoneExtensionLabel3)

	<#
		Config "Mailbox Features:" label
	#>
	$mailboxFeaturesLabel = Config-Label -horizontalPosition 420 -verticalPosition 370 -width 120 -height 15 -text "Mailbox Features:" -fontFamily "Arial" -fontSize 8 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	#$newMailBoxForm.Controls.Add($mailboxFeaturesLabel)

	<#
		Config "( Approve Needed )" label
	#>
	$approveNeededLabel = Config-Label -horizontalPosition 420 -verticalPosition 390 -width 120 -height 15 -text "( Approve Needed )"
	#$newMailBoxForm.Controls.Add($approveNeededLabel)

	<#
		Config "Outlook Anywhere" label
	#>
	$outlookAnywhereLabel = Config-Label -horizontalPosition 565 -verticalPosition 350 -width 150 -height 20 -text "Outlook Anywhere"
	#$newMailBoxForm.Controls.Add($outlookAnywhereLabel)

	<#
		Config "Outlook Web Access (OWA)" label
	#>
	$outlookWebAccess = Config-Label -horizontalPosition 565 -verticalPosition 370 -width 150 -height 20 -text "Outlook Web Access (OWA)"
	#$newMailBoxForm.Controls.Add($outlookWebAccess)

	<#
		"Exchange ActiveSync" label
	#>
	$exchangeActiveSyncLabel = Config-Label -horizontalPosition 565 -verticalPosition 390 -width 120 -height 20 -text "Exchange ActiveSync"
	#$newMailBoxForm.Controls.Add($exchangeActiveSyncLabel)

	<#
		Config "Company" combo box and its behaviors
	#>
	$companyComboBox = Config-ComboBox -horizontalPosition 190 -verticalPosition 70 -width 90 -height 20
	$companyComboBox.Items.AddRange(("", "MA Labs", "Superbiiz", "Supertalent"))
	$companyComboBox.DropDownStyle = 2
	$companyComboBox.Sorted = $True
	$companyComboBox.TabIndex = 0
	if (!$new) {

		$companyComboBox.Enabled = $False

	}
	$companyComboBox.Add_SelectedValueChanged({

		Handle-CompanyComboBoxChanged

	})
	$newMailBoxForm.Controls.Add($companyComboBox)

	<#
		Config "Office" combo box and its behaviors
	#>
	$officeComboBox = Config-ComboBox -horizontalPosition 290 -verticalPosition 70 -width 90 -height 20
	$officeComboBox.DropDownStyle = 2
	$officeComboBox.Sorted = $True
	$officeComboBox.TabIndex = 1
	if (!$new) {

		$officeComboBox.Enabled = $False

	}
	$officeComboBox.Add_SelectedValueChanged({

		Handle-OfficeComboBoxChanged

	})
	$newMailBoxForm.Controls.Add($officeComboBox)

	<#
		Config "Department" combo box and its behaviors
	#>
	$departmentComboBox = Config-ComboBox -horizontalPosition 390 -verticalPosition 70 -width 90 -height 20
	$departmentComboBox.DropDownStyle = 2
	$departmentComboBox.Sorted = $True
	$departmentComboBox.DropDownHeight = 750
	$departmentComboBox.TabIndex = 1
	if (!$new) {

		$departmentComboBox.Enabled = $False

	}

	$departmentComboBox.Add_SelectedValueChanged({

		Handle-DepartmentComboBoxChanged

	})
	$newMailBoxForm.Controls.Add($departmentComboBox)

	<# 
		Subgroup for PM (Specially for different selling target)
	#>
	$PMGroupComboBox = $departmentComboBox = Config-ComboBox -horizontalPosition 490 -verticalPosition 70 -width 90 -height 20
	$PMGroupComboBox.DropDownStyle = 2
	$PMGroupComboBox.Sorted = $True
	$PMGroupComboBox.visible = $False
	$PMGroupComboBox.Add_SelectedValueChanged({

		Handle-PMGroupComboBoxChanged

	})
	$newMailBoxForm.Controls.Add($PMGroupComboBox)

	<#
		Config "Organizational Unit" combo box
	#>
	$organizationalUnitComboBox = Config-ComboBox -horizontalPosition 190 -verticalPosition 100 -width 290 -height 20
	$organizationalUnitComboBox.Items.Add("/Users")
	$organizationalUnitComboBox.DropDownStyle = 2
	$organizationalUnitComboBox.DropDownHeight = 750
	$organizationalUnitComboBox.Sorted = $True
	$organizationalUnitComboBox.TabIndex = 2
	if (!$new) {

		$organizationalUnitComboBox.Enabled = $False

	}
	$newMailBoxForm.Controls.Add($organizationalUnitComboBox)

	<#
		Config "Organizational Unit" message label
	#>
	$organizationalUnitMessageLabel = Config-Label -horizontalPosition 450 -verticalPosition 105 -width 20 -height 20
	if ($new) {

		$organizationalUnitMessageLabel.Text = "*"

	}
	$organizationalUnitMessageLabel.ForeColor = "Red"
	$newMailBoxForm.Controls.Add($organizationalUnitMessageLabel)

	<#
		First Name/Last Name text box
	#>
	$firstNameLastNameTextBox = Config-TextBox -horizontalPosition 190 -verticalPosition 130 -width 240 -height 20 -tabIndex 3
	if (!$new) {

		$firstNameLastNameTextBox.Enabled = $False

	}
	$newMailBoxForm.Controls.Add($firstNameLastNameTextBox)

	<#
		First Name text box
	#>
	$firstNameTextBox = Config-TextBox -horizontalPosition 190 -verticalPosition 130 -width 120 -height 20 -tabIndex 3
	if (!$new) {

		$firstNameTextBox.Enabled = $False

	}
	#$newMailBoxForm.Controls.Add($firstNameTextBox)

	<#
		First Name message label
	#>
	$firstNameMessageLabel = Config-Label -horizontalPosition 490 -verticalPosition 135 -width 20 -height 20
	if ($new) {

		$firstNameMessageLabel.Text = "*"

	}
	$firstNameMessageLabel.ForeColor = "Red"
	$newMailBoxForm.Controls.Add($firstNameMessageLabel)

	<#
		Last Name text box
	#>
	$lastNameTextBox = Config-TextBox -horizontalPosition 315 -verticalPosition 130 -width 120 -height 20 -tabIndex 4
	if (!$new) {

		$lastNameTextBox.Enabled = $False

	}
	#$newMailBoxForm.Controls.Add($lastNameTextBox)


	<#
		User Autofill button
	#>	
	$autofillButton = Config-Button -horizontalPosition 440 -verticalPosition 130 -width 40 -height 20 -text "Auto Fill" -tabIndex 5
	$autofillButton.Add_Click({

		Handle-AutofillComboBoxClicked

	})
	if (!$new) {

		$autofillButton.Enabled = $False

	}
	$newMailBoxForm.Controls.Add($autofillButton)


	<#
		User Logon Name text box
	#>
	$userLogonNameTextBox = Config-TextBox -horizontalPosition 190 -verticalPosition 160 -width 140 -height 20 -tabIndex 6
	if ($new) {

		$userLogonNameTextBox.Enabled = $False

	} else {

		$userLogonNameTextBox.AutoCompleteSource = 'CustomSource'
		$userLogonNameTextBox.AutoCompleteMode = 'SuggestAppend'
		$userLogonNameTextBox.AutoCompleteCustomSource = $autocomplete
		Get-User | ForEach-Object {

			$userLogonNameTextBox.AutoCompleteCustomSource.Add($_.SamAccountName)

		}

	}
	$userLogonNameTextBox.CharacterCasing = 2
	$newMailBoxForm.Controls.Add($userLogonNameTextBox)

	<#
		Logon Name message label
	#>
	$logonNameMessageLabel = Config-Label -horizontalPosition 490 -verticalPosition 165 -width 20 -height 20 -text "*"
	$logonNameMessageLabel.ForeColor = "Red"
	$newMailBoxForm.Controls.Add($logonNameMessageLabel)

	<#
		Domain combo box
	#>
	$domainComboBox = Config-ComboBox -horizontalPosition 340 -verticalPosition 160 -width 140 -height 20
	$domainComboBox.Items.AddRange(("@ma.local", "@superbiiz.com", "@supertalent.com"))
	$domainComboBox.Add_SelectedValueChanged({

		$emailDomainComboBox.SelectedItem = $domainComboBox.SelectedItem

	})
	$domainComboBox.DropDownStyle = 2
	$domainComboBox.Sorted = $True
	$domainComboBox.TabIndex = 7
	$domainComboBox.Enabled = $False
	$newMailBoxForm.Controls.Add($domainComboBox)

	<#
		Password text box
	#>
	$passwordTextBox = Config-TextBox -horizontalPosition 190 -verticalPosition 190 -width 140 -height 20 -tabIndex 8 -maxLength 20
	$passwordTextBox.UseSystemPasswordChar = $False
	$passwordTextBox.Enabled = $False
	$newMailBoxForm.Controls.Add($passwordTextBox)

	<#
		Password message label
	#>
	$passwordMessageLabel = Config-Label -horizontalPosition 490 -verticalPosition 195 -width 20 -height 20
	if ($new) {

		$passwordMessageLabel.Text = "*"

	}
	$passwordMessageLabel.ForeColor = "Red"
	$newMailBoxForm.Controls.Add($passwordMessageLabel)

	<#
		Server combo box
	#>
	$serverComboBox = Config-ComboBox -horizontalPosition 190 -verticalPosition 220 -width 140 -height 20
	Get-MailboxServer | ForEach-Object {

		if ($_.Name -ne "PI") {

			$serverComboBox.Items.Add($_.Name)

		}

	}	
	$serverComboBox.DropDownStyle = 2
	$serverComboBox.Sorted = $True
	$serverComboBox.TabIndex = 9
	$serverComboBox.Enabled = $False

	$serverComboBox.Add_SelectedValueChanged({

		Hanle-ServerComboBoxChanged

	})
	$newMailBoxForm.Controls.Add($serverComboBox)

	<#
		Server Name message label
	#>	
	$serverNameMessage = Config-Label -horizontalPosition 490 -verticalPosition 225 -width 20 -height 20
	if ($new) {

		$serverNameMessage.Text = "*"

	}
	$serverNameMessage.ForeColor = "Red"
	$newMailBoxForm.Controls.Add($serverNameMessage)

	<#
		Database checked list box
	#>
	$databaseCheckedList = Config-CheckedListbox -horizontalPosition 190 -verticalPosition 250 -width 290 -height 280 -checkOnClick $True -horizontalScrollbar $True -enabled $False
	$newMailBoxForm.Controls.Add($databaseCheckedList)

	<#
		E-mail text box
	#>
	$emailTextBox = Config-TextBox -horizontalPosition 190 -verticalPosition 540 -width 140 -height 20 -tabIndex 11
	$emailTextBox.Enabled = $False
	$newMailBoxForm.Controls.Add($emailTextBox)

	<#
		E-mail Combo Box
	#>
	$emailDomainComboBox = Config-ComboBox -horizontalPosition 340 -verticalPosition 540 -width 140 -height 20
	$emailDomainComboBox.Items.AddRange(("@ma.local", "@superbiiz.com", "@supertalent.com"))
	$emailDomainComboBox.DropDownStyle = 2
	$emailDomainComboBox.Sorted = $True
	$emailDomainComboBox.TabIndex = 12
	$emailDomainComboBox.Enabled = $False
	$newMailBoxForm.Controls.Add($emailDomainComboBox)

	<#
		E-mail message label
	#>
	$emailMessageLabel = Config-Label -horizontalPosition 450 -verticalPosition 545 -width 20 -height 20
	$emailMessageLabel.ForeColor = "Red"
	$newMailBoxForm.Controls.Add($emailMessageLabel)

	<#
		Distribution Group checked list box
	#>
	$distributionGroupCheckedListBox = Config-CheckedListbox -horizontalPosition 595 -verticalPosition 80 -width 160 -height 520 -horizontalScrollbar $True -enabled $False
	$newMailBoxForm.Controls.Add($distributionGroup)

	<#
		Distribution Group text box
	#>
	<#
		The distributionGroupTextBox for the real name of the distribution group
	#>
	<#
	$distributionGroupTextBox = Config-TextBox -horizontalPosition 700 -verticalPosition 280 -width 140 -height 20 -tabIndex 14 -maxLength 12
	$distributionGroupTextBox.Enabled = $False
	$newMailBoxForm.Controls.Add($distributionGroupTextBox)
	#>

	<#
		Phone text box
	#>
	$phoneTextBox = Config-TextBox -horizontalPosition 190 -verticalPosition 570 -width 140 -height 20 -tabIndex 14 -maxLength 12
	$phoneTextBox.Enabled = $False
	$newMailBoxForm.Controls.Add($phoneTextBox)

	<#
		Extension Text Box
	#>
	$extensionTextBox = Config-TextBox -horizontalPosition 340 -verticalPosition 570 -width 140 -height 20 -tabIndex 15
	$extensionTextBox.Enabled = $False
	$newMailBoxForm.Controls.Add($extensionTextBox)

	<#
		Phone Message label
	#>
	$phoneMessageLabel = Config-Label -horizontalPosition 450 -verticalPosition 575 -width 20 -height 20 -text "MS Lync Account:"	
	$phoneMessageLabel.ForeColor = "Red"
	$newMailBoxForm.Controls.Add($phoneMessageLabel)

	<#
		Outlook Anywhere Check Box
	#>
	$outlookAnywhereCheckBox = Config-CheckBox -horizontalPosition 550 -verticalPosition 348 -width 140 -height 20 -tabIndex 14
	if (!$new) {

		$outlookAnywhereCheckBox.Enabled = $False

	}
	# $newMailBoxForm.Controls.Add($outlookAnywhereCheckBox)

	<#
		OWA Check Box
	#>
	$OWACheckBox = Config-CheckBox -horizontalPosition 550 -verticalPosition 368 -width 140 -height 20 -tabIndex 16 -enabled $False
	#$newMailBoxForm.Controls.Add($OWACheckBox)

	<#
		Exchange ActiveSync Check Box
	#>
	$activeSyncCheckBox = Config-CheckBox -horizontalPosition 550 -verticalPosition 388 -width 140 -height 20 -tabIndex 16 -enabled $False
	#$newMailBoxForm.Controls.Add($activeSyncCheckBox)


	<#
		Lync label
	#>
	$lyncLabel = Config-Label -horizontalPosition 20 -verticalPosition 615 -width 140 -height 20 -text "MS Lync Account:"
	#$newMailBoxForm.Controls.Add($lyncLabel)

	<#
		Lync Check box
	#>
	$lyncCheckBox = Config-CheckBox -horizontalPosition 190 -verticalPosition 615 -width 140 -height 20 -tabIndex 18

	<#
		Add Group Membership labels
	#>
	$networkDriveLabel = Config-Label -horizontalPosition 775 -verticalPosition 60 -width 85 -height 20 -text "Network Drive: "
	$newMailBoxForm.Controls.Add($networkDriveLabel)

	$NDFSLabel = Config-Label -horizontalPosition 775 -verticalPosition 80 -width 85 -height 20 -text "NDFS: "
	$newMailBoxForm.Controls.Add($NDFSLabel)

	$FSLabel = Config-Label -horizontalPosition 775 -verticalPosition 350 -width 85 -height 20 -text "FS: "
	$newMailBoxForm.Controls.Add($FSLabel)

	$websenseDepartmentLabel = Config-Label -horizontalPosition 1000 -verticalPosition 80 -width 165 -height 20 -text "Websense Department: "
	$newMailBoxForm.Controls.Add($websenseDepartmentLabel)

	$websenseUserDefineLabel = Config-Label -horizontalPosition 1000 -verticalPosition 350 -width 185 -height 20 -text "Websense User Define: "
	$newMailBoxForm.Controls.Add($websenseUserDefineLabel)

	<#
		The list box of group name under DFS
	#>
	$DFSCheckedListBox = Config-CheckedListbox -horizontalPosition 775 -verticalPosition 100 -width 200 -height 230 -checkOnClick $True -horizontalScrollbar $True -enabled $False
	$newMailBoxForm.Controls.Add($DFSCheckedListBox)

	<# 
		The list box of group name under vasto
	#>
	$fileShareCheckedListBox = Config-CheckedListbox -horizontalPosition 775 -verticalPosition 380 -width 200 -height 220 -checkOnClick $True -horizontalScrollbar $True -enabled $False
	$newMailBoxForm.Controls.Add($fileShareCheckedListBox)

	<# 
		The list box of group name for websense department
	#>
	$websenseDepartmentCheckedListBox = Config-CheckedListbox -horizontalPosition 1000 -verticalPosition 100 -width 200 -height 230 -checkOnClick $True -horizontalScrollbar $True -enabled $False
	$newMailBoxForm.Controls.Add($websenseDepartmentCheckedListBox)

	<#
		The list box of group name for websense user define
	#>
	$websenseUserDefineCheckedListBox = Config-CheckedListbox -horizontalPosition 1000 -verticalPosition 380 -width 200 -height 220 -checkOnClick $True -horizontalScrollbar $True -enabled $False
	$newMailBoxForm.Controls.Add($websenseUserDefineCheckedListBox)

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

	$OnLoadForm_StateCorrection = {

		Build-TreeView

	}

	#$treeView1.Size = New-Object System.Drawing.Size(260, 270)
	#$treeView1.Name = "treeView1" 
	#$treeView1.Location = New-Object System.Drawing.Size(15, 15)
	#$treeView1.DataBindings.DefaultDataSourceUpdateMode = 0 
	#$treeView1.TabStop = $False
	#$newMailBoxForm.Controls.Add($treeView1)

	<#
		Department server reference labels
	#>
	$BerdychServerLabel = Config-Label -horizontalPosition 15 -verticalPosition 650 -width 50 -height 20 -text "Berdych: " -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	#$newMailBoxForm.Controls.Add($BerdychServerLabel)

	$HQPMDepartmentLabel = Config-Label -horizontalPosition 65 -verticalPosition 650 -width 100 -height 20 -text "HQ-PM"
	#$newMailBoxForm.Controls.Add($HQPMDepartmentLabel)

	$IsnerServerLabel = Config-Label -horizontalPosition 15 -verticalPosition 770 -width 50 -height 20 -text "Isner: " -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$newMailBoxForm.Controls.Add($IsnerServerLabel)

	$HQSalesDepartmentLabel = Config-Label -horizontalPosition 70 -verticalPosition 770 -width 70 -height 20 -text "HQ-Sales: "
	$newMailBoxForm.Controls.Add($HQSalesDepartmentLabels)

	$DjokovicServerLabel = Config-Label -horizontalPosition 140 -verticalPosition 770 -width 70 -height 20 -text "Djokovic: " -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$newMailBoxForm.Controls.Add($DjokovicServerLabel)

	$HRMarketingETCDepartmentLabel = Config-Label -horizontalPosition 210 -verticalPosition 770 -width 201 -height 20 -text "HR, Markerting, MIS, MGMT,IT, Payroll"
	$newMailBoxForm.Controls.Add($HRMarketingETCDepartmentLabel)

	$WawrinkaServerLabel = Config-Label -horizontalPosition 410 -verticalPosition 770 -width 90 -height 20 -text "Wawrinka: " -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$newMailBoxForm.Controls.Add($WawrinkaServerLabel)

	$TechSupportWarehouseETCDepartmentLabel = Config-Label -horizontalPosition 470 -verticalPosition 770 -width 280 -height 20 -text "Tech Support, Warehouse, RMA, LA, Data Entry"
	$newMailBoxForm.Controls.Add($TechSupportWarehouseETCDepartmentLabel)

	$RaonicServerLabel = Config-Label -horizontalPosition 15 -verticalPosition 790 -width 60 -height 20 -text "Wawrinka: " -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$newMailBoxForm.Controls.Add($RaonicServerLabel)

	$CHGAETCLocationLabel = Config-Label -horizontalPosition 73 -verticalPosition 790 -width 100 -height 20 -text "CH, GA, MI, NJ"
	$newMailBoxForm.Controls.Add($CHGAETCLocationLabel)

	$GasquetServerLabel = Config-Label -horizontalPosition 173 -verticalPosition 790 -width 53 -height 20 -text "Gasquet: " -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$newMailBoxForm.Controls.Add($GasquetServerLabel)

	$SBZLabel = Config-Label -horizontalPosition 230 -verticalPosition 790 -width 30 -height 20 -text "SBZ"
	$newMailBoxForm.Controls.Add($SBZLabel)

	$JarguarServerLabel = Config-Label -horizontalPosition 273 -verticalPosition 790 -width 60 -height 20 -text "Jarguar: " -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$newMailBoxForm.Controls.Add($JarguarServerLabel)

	$STTLabel = Config-Label -horizontalPosition 333 -verticalPosition 790 -width 50 -height 20 -text "STT"
	$newMailBoxForm.Controls.Add($STTLabel)

	$SharapovaServerLabel = Config-Label -horizontalPosition 383 -verticalPosition 790 -width 65 -height 20 -text "Sharapova: " -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$newMailBoxForm.Controls.Add($SharapovaServerLabel)

	$APARCRLabel = Config-Label -horizontalPosition 453 -verticalPosition 790 -width 70 -height 20 -text "AP, AR, CR"
	$newMailBoxForm.Controls.Add($APARCRLabel)

	$SprintbokServerLabel = Config-Label -horizontalPosition 523 -verticalPosition 790 -width 65 -height 20 -text "Sprintbok: " -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
	$newMailBoxForm.Controls.Add($SprintbokServerLabel)

	$WuhanLabel = Config-Label -horizontalPosition 588 -verticalPosition 790 -width 100 -height 20 -text "Wuhan"
	$newMailBoxForm.Controls.Add($WuhanLabel)

	#endregion Generated Form Code

	#Save the initial state of the form
	$InitialFormWindowState = $newMailBoxForm.WindowState

	#Init the OnLoad event to correct the initial state of the form
	$newMailBoxForm.add_Load($OnLoadForm_StateCorrection)

	#$newMailBoxForm.Topmost = $True
	$newMailBoxForm.Add_Shown({

		$newMailBoxForm.Activate()

	})
	$newMailBoxForm.ShowDialog()

}

function Handle-Enters {

	$UserAlias = $userLogonNameTextBox.Text.Trim()

	if ($new -or $found) {

		$UserFirst = ToProperCase ($firstNameLastNameTextBox.Text.Split()[0].Trim())
		$UserLast = ToProperCase ($firstNameLastNameTextBox.Text.Split()[1].Trim())
		$UserPassword = $passwordTextBox.Text
		$UserOU = $organizationalUnitComboBox.Text
		$MailboxServer = $serverComboBox.SelectedItem
		$Database = $databaseCheckedList.SelectedItem.Split("==")[0].Trim()
		$UserDomain = $domainComboBox.SelectedItem
		$UserOffice = $officeComboBox.SelectedItem
		$UserPhone = $phoneTextBox.Text.Trim()
		$UserExtension = $extensionTextBox.Text.Trim()
		$UserEmail = $emailTextBox.Text.Trim()

		#Group of DG names
		$GroupName_DG = @()
		foreach ($objItem in $distributionGroupCheckedListBox.CheckedItems) {

			$GroupName_DG += $objItem

		}
		#The Data Store from the group membership selection list
		$GroupName_DFS = @()
		$GroupName_Vasto = @()
		$GroupName_Dept = @()
		$GroupName_List = @()
		foreach ($objItem in $DFSCheckedListBox.CheckedItems) {

			$GroupName_DFS += $objItem

		}

		foreach ($objItem in $objListbox_FileShare.CheckedItems) {

			$GroupName_Vasto += $objItem

		}

		foreach ($objItem in $websenseDepartmentCheckedListBox.CheckedItems) {

			$GroupName_Dept += $objItem

		}

		foreach ($objItem in $websenseUserDefineCheckedListBox.CheckedItems) {

			$GroupName_List += $objItem

		}

		if (!$emailTextBox.Text) {

			if ($UserFirst -and $UserLast) {

				$UserEmail = $UserFirst + "." + $UserLast

			} else {

				$UserEmail = $UserFirst + $UserLast

			}

			$UserEmail = $UserEmail.Replace(" ", "")

		}

		$UserEmailDomain = $emailDomainComboBox.SelectedItem
		$UserOutlookAnywhere = $outlookAnywhereCheckBox.Checked
		$UserOWA = $OWACheckBox.Checked
		$UserActiveSync = $activeSyncCheckBox.Checked
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

				$organizationalUnitComboBox.SelectedItem = $([string]$tmpUser.Identity.Parent).substring($Forest.Length)
				$firstNameTextBox.Text = $tmpUser.FirstName
				$lastNameTextBox.Text = $tmpUser.LastName
				$domainComboBox.SelectedItem = $tmpUser.UserPrincipalName.substring($UserAlias.Length)
				$passwordTextBox.Text = "********"
				$officeComboBox.SelectedItem = $tmpUser.Office
				$emailTextBox.Text = $tmpUser.WindowsEmailAddress.Local

				if ($tmpUser.WindowsEmailAddress.domainComboBox) {

					$emailDomainComboBox.Text = "@" + $tmpUser.WindowsEmailAddress.domainComboBox

				}
				
				if ($tmpUser.Phone) {

					$telephone = ($tmpUser.Phone).Split('x')

					switch ($telephone.Count) {

						1 { 

							$phoneTextBox.Text = $telephone[0].Trim()

						}

						2 {

							$phoneTextBox.Text = $telephone[0].Trim(); 
							$extensionTextBox.Text = $telephone[1].Trim()

						}

						default {}

					}

				}

				$domainComboBox.Enabled = $True
				$emailTextBox.Enabled = $True
				$emailDomainComboBox.Enabled = $True
				$serverComboBox.Enabled = $True
				$databaseCheckedList.Enabled = $True
				$distributionGroupCheckedListBox.Enabled = $True
				$officeComboBox.Enabled = $True
				$phoneTextBox.Enabled = $True
				$extensionTextBox.Enabled = $True
				$outlookAnywhereCheckBox.Enabled = $True
				$OWACheckBox.Enabled = $True
				$activeSyncCheckBox.Enabled = $True

			}

			if ($disable) { 

				ShowForm2

			} else {					

				$newMailBoxForm.Refresh() 

			}

		}

	}

}

function Handle-OfficeComboBoxChanged {

	$userLogonNameTextBox.Enabled = $False
		$domainComboBox.Enabled = $False
		$passwordTextBox.Enabled = $False
		$serverComboBox.Enabled = $False
		$databaseCheckedList.Enabled = $False
		$emailTextBox.Enabled = $False
		$emailDomainComboBox.Enabled = $False
		$distributionGroupCheckedListBox.Enabled = $False
		$phoneTextBox.Enabled = $False
		$extensionTextBox.Enabled = $False
		$DFSCheckedListBox.Enabled = $False
		$fileShareCheckedListBox.Enabled = $False

		#$OWACheckBox.Enabled = $False
		#$activeSyncCheckBox.Enabled = $False
		$organizationalUnitComboBox.Items.Clear()
		$organizationalUnitComboBox.Text = $null
		$emailTextBox.Text = $null
		$passwordTextBox.Text = $null
		$userLogonNameTextBox.Text = $null
		$phoneTextBox.Text = $null
		$extensionTextBox.Text = $null

		$root = [adsi]''
		$searcher = New-Object System.DirectoryServices.DirectorySearcher ($root)

		if ($officeComboBox.SelectedItem.ToString() -eq "") {

			$Officename = ''

		} else {

			$Officename = $officeComboBox.SelectedItem.ToString()
			$MainCompanyName = $companyComboBox.SelectedItem.ToString()
			$NameCombine = "$MainCompanyName/$Officename"

			$departmentComboBox.Items.Clear()
			$PMGroupComboBox.visible = $False
			$PMGroupComboBox.Items.Clear()

			switch ($NameCombine) {

				{ $_ -eq "MA Labs/Georgia" } {

					$departmentComboBox.Items.AddRange(("Accounting", "HR", "Others", "RMA", "Sales"))

				}

				{ $_ -eq "MA Labs/San Jose" } {

					$departmentComboBox.Items.AddRange(("AP", "AR", "IT", "Sales", "Marketing", "HR", "Payroll", "MIS", "Credit", "Data Entry", "Purchasing", "Shipping", "Tech Support", "Warehouse", "ACCT", "RMA"))

				}

				{ $_ -eq "MA Labs/Los Angles" } {

					$departmentComboBox.Items.AddRange(("Accounting", "Others", "RMA", "Sales", "Warehouse", "HR"))

				}

				{ $_ -eq "MA Labs/New Jersey" } {

					$departmentComboBox.Items.AddRange(("Others", "RMA", "Sales", "Warehouse", "Accounting", "Purchasing"))

				}

				{ $_ -eq "MA Labs/Chicago" } {

					$departmentComboBox.Items.AddRange(("Accounting", "Others", "Sales"))

				}

				{ $_ -eq "MA Labs/Miami" } {

					$departmentComboBox.Items.AddRange(("Accounting", "Others", "Sales", "Purchasing", "Warehouse"))

				}

				{ $_ -eq "MA Labs/Wuhan" } {

					$departmentComboBox.Items.AddRange(("AP", "AR", "Credit", "Inventory", "Marketing", "PM", "HR", "Sales"))

				}

				{ $_ -eq "Supertalent/San Jose" } {

					$departmentComboBox.Items.AddRange(("Accounting", "Engineering", "HR", "Sales", "Marketing", "Tech Support", "RMA"))

				}

				{ $_ -eq "Superbiiz/San Jose" } {

					$departmentComboBox.Items.AddRange(("Marketing", "Customer Service", "Sales", "Accounting", "MIS"))

				}

				{ $_ -eq "Superbiiz/Wuhan" } {

					$departmentComboBox.Items.AddRange(("Users", "Accounting"))

				}

				{ $_ -eq "Supertalent/Wuhan" } {

					$departmentComboBox.Items.AddRange(("Users", "Sales"))

				}

			}

		}

		#The Group Note Search to detect the distribution group
		#Get-Group -Filter "Notes -eq '$Officename'"| ForEach-Object{[void] $distributionGroupCheckedListBox.Items.Add($_.Name)}


}

function Handle-CompanyComboBoxChanged {

	$userLogonNameTextBox.Enabled = $False
	$domainComboBox.Enabled = $False
	$passwordTextBox.Enabled = $False
	$serverComboBox.Enabled = $False
	$databaseCheckedList.Enabled = $False
	$emailTextBox.Enabled = $False
	$emailDomainComboBox.Enabled = $False
	$distributionGroupCheckedListBox.Enabled = $False
	$phoneTextBox.Enabled = $False
	$extensionTextBox.Enabled = $False
	$DFSCheckedListBox.Enabled = $False
	$fileShareCheckedListBox.Enabled = $False

	$domainComboBox.Items.Clear()
	$domainComboBox.Items.AddRange(("@ma.local", "@superbiiz.com", "@supertalent.com"))

	$emailDomainComboBox.Items.Clear()
	$emailDomainComboBox.Items.AddRange(("@ma.local", "@superbiiz.com", "@supertalent.com"))

	$organizationalUnitComboBox.Items.Clear()
	$organizationalUnitComboBox.Text = $null

	$officeComboBox.Items.Clear()
	$officeComboBox.Text = $null
	$PMGroupComboBox.visible = $False
	$PMGroupComboBox.Items.Clear()
	$emailTextBox.Text = $null
	$passwordTextBox.Text = $null
	$userLogonNameTextBox.Text = $null
	$phoneTextBox.Text = $null
	$extensionTextBox.Text = $null

	if ($companyComboBox.SelectedItem.ToString() -eq "") {

		$MainCompanyName = ''
		$TitleLabel.Text = ''

	} else {

		$MainCompanyName = $companyComboBox.SelectedItem.ToString()
		$TitleLabel.Text = $MainCompanyName

	}

	switch ($MainCompanyName) {

		{ $_ -eq "MA Labs" } {

			$domainComboBox.SelectedIndex = 0
			$officeComboBox.Items.Add(("Chicago", "Georgia", "Los Angles", "New Jersey", "Miami", "San Jose", "Wuhan"))

		}

		{ $_ -eq "Superbiiz" } {

			$domainComboBox.SelectedIndex = 1
			$officeComboBox.Items.Add(("San Jose", "Wuhan"))

		}

		{ $_ -eq "Supertalent" } {

			$domainComboBox.SelectedIndex = 2
			$officeComboBox.Items.Add(("San Jose", "Wuhan"))

		}

	}

}

function Handle-DepartmentComboBoxChanged {
	
	$PMGroupComboBox.visible = $False
	$PMGroupComboBox.Items.Clear()
	$PmgroupName = $Null

	if ($departmentComboBox.SelectedItem.ToString() -eq "") {

		$DepartmentName = ''

	} else {

		$DepartmentName = $departmentComboBox.SelectedItem.ToString()

	}

	$root = [adsi]''
	$searcher = New-Object System.DirectoryServices.DirectorySearcher ($root)

	if ($officeComboBox.SelectedItem.ToString() -eq "") {

		$Officename = ''

	} else {

		$MainCompanyName = $companyComboBox.SelectedItem.ToString()
		$Officename = $officeComboBox.SelectedItem.ToString()
		$NameCombine = "$MainCompanyName/$Officename"

		#Control the subgroup of the purchasing
		if (($officeComboBox.Text -eq "San Jose") -and ($departmentComboBox.Text -eq "Purchasing")) {

			$PMGroupComboBox.visible = $True
			$PMGroupComboBox.Items.AddRange(("HDD", "Memory", "Microsoft", "Monitor", "Motherboard", "Networking", "Notebook", "VGA"))

		}

		if (($officeComboBox.Text -eq "Wuhan") -and ($departmentComboBox.Text -eq "PM")) {

			$PMGroupComboBox.visible = $True
			$PMGroupComboBox.Items.AddRange(("HDD", "Microsoft", "Monitor", "Motherboard", "Networking", "Notebook"))

		}

		#The Group Note Search to detect the distribution group
		#Get-Group -Filter "Notes -eq '$Officename'"| ForEach-Object{[void] $distributionGroupCheckedListBox.Items.Add($_.Name)}
		switch ($NameCombine) {

			{ $_ -eq "MA Labs/Georgia" } {

				$phoneTextBox.Text = "770-209-6600"
				$extensionTextBox.MaxLength = 3
				$OUKeyword = "MA Labs/GA"

			}

			{ $_ -eq "MA Labs/San Jose" } {

				$phoneTextBox.Text = "408-941-0808"
				$extensionTextBox.MaxLength = 3
				$OUKeyword = "MA Labs/San Jose/$DepartmentName"

			}

			{ $_ -eq "MA Labs/Los Angles" } {

				$phoneTextBox.Text = "626-820-8988"
				$extensionTextBox.MaxLength = 3
				$OUKeyword = "MA Labs/LA"

			}

			{ $_ -eq "MA Labs/New Jersey" } {

				$phoneTextBox.Text = "732-661-3388"
				$extensionTextBox.MaxLength = 3
				$OUKeyword = "MA Labs/NJ"

			}

			{ $_ -eq "MA Labs/Chicago" } {

				$phoneTextBox.Text = "630-893-2323"
				$extensionTextBox.MaxLength = 3
				$OUKeyword = "MA Labs/Chicago"

			}

			{ $_ -eq "MA Labs/Miami" } {

				$phoneTextBox.Text = "305-594-8700"
				$extensionTextBox.MaxLength = 3
				$OUKeyword = "MA Labs/Miami"

			}

			{ $_ -eq "MA Labs/Wuhan" } {

				$phoneTextBox.Text = "086-275-973-1208"
				$extensionTextBox.MaxLength = 4
				$OUKeyword = "MA Labs/China"

			}

			{ $_ -eq "Supertalent/San Jose" } {

				$phoneTextBox.Text = "408-934-2560"
				$extensionTextBox.MaxLength = 3
				$OUKeyword = "Supertalent/$DepartmentName"

			}

			{ $_ -eq "Superbiiz/San Jose" } {

				$phoneTextBox.Text = "408-934-2500"
				$extensionTextBox.MaxLength = 3
				$OUKeyword = "Superbiiz/$DepartmentName"

			}

			{ $_ -eq "Superbiiz/Wuhan" } {

				$phoneTextBox.Text = "086-275-973-1208"
				$extensionTextBox.MaxLength = 4
				$OUKeyword = "Superbiiz/China/"

			}

			{ $_ -eq "Supertalent/Wuhan" } {

				$phoneTextBox.Text = "086-275-973-1208"
				$extensionTextBox.MaxLength = 4
				$OUKeyword = "Supertalent/China/"

			}

		}

		$organizationalUnitComboBox.Items.Clear()
		$searcher.Filter = "(&(objectClass=organizationalUnit)(name=Users))"
		$searcher.PropertiesToLoad.Add("canonicalName")
		$searcher.PropertiesToLoad.Add("Name")
		$searcherall = $searcher.FindAll()
		$findTag = 0
		foreach ($person in $searcherall) {

			[string]$ent = $person.properties.canonicalname
			$OUTarget = $ent.substring($ent.IndexOf("/"), $ent.Length - $ent.IndexOf("/"))

			if (($OUTarget -like "*$OUKeyword*") -and !($OUTarget -like "*Disabled*")) {

				$organizationalUnitComboBox.Items.Add($OUTarget)
				$organizationalUnitComboBox.SelectedIndex = 0
				$findTag = 1

			}

			if (($NameCombine -eq "Supertalent/San Jose") -and ($OUTarget -like "*Supertalent/China*")) {

				$organizationalUnitComboBox.Items.Remove($OUTarget)

			}

			if (($NameCombine -eq "Superbiiz/San Jose") -and ($OUTarget -like "*Superbiiz/China*")) {

				$organizationalUnitComboBox.Items.Remove($OUTarget)

			}

		}

	}

}

function Handle-PMGroupComboBoxChanged {

	if ($PMGroupComboBox.SelectedItem.ToString() -eq "") {

		$PmgroupName = ''

	} else {

		$PmgroupName = $PMGroupComboBox.SelectedItem.ToString()

	}

}

function Handle-AutofillComboBoxClicked {

	$UserFirst = ToProperCase ($firstNameLastNameTextBox.Text.Split()[0].Trim())
	$UserLast = ToProperCase ($firstNameLastNameTextBox.Text.Split()[1].Trim())
	$UserOU = $organizationalUnitComboBox.SelectedItem

	if (Autofill) {

		ShowForm2

	}

}

function Hanle-ServerComboBoxChanged {

	$MBhash = @{}
	$databaseCheckedList.Items.Clear()
	$databaseCheckedList.Text = $null
	Get-MailboxDatabase -Status -Server $serverComboBox.SelectedItem.ToString() | ForEach-Object {

		$Name = $_.Name
		$ServerName = $_.server
		$Filepath1 = $_.EdbFilePath
		$Fullpath2 = "`\`\" + $_.server + "`\" + $_.EdbFilePath.DriveName.Remove(1).ToString() + "$" + $_.EdbFilePath.PathName.Remove(0, 2)
		$Sizeinfo = ((Get-ChildItem $Fullpath2).Length) / 1048576KB
		$Size = [math]::Round($Sizeinfo, 2)

		$MBInfo = "$Name == $Size GB"
		if ($Size -lt 40) {

			$databaseCheckedList.ForeColor = "Black"

		} else {

			$databaseCheckedList.ForeColor = "Red"

		}

		$databaseCheckedList.Items.Add($MBInfo)
		$MBhash.Add($_.Name, $_.ServerName + "\" + $_.StorageGroup.Name + "\" + $_.Name)

	}

}