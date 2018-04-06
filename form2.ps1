<################################################################################
# Form 2
################################################################################>


function ShowForm2 {

	$newMailBoxForm.visible = $False

	$Form2 = Config-Form -width 900 -height 710 -text $title
	$Form2.Add_KeyDown({

		Handle-KeyDown

	})

	if ($UserExtension) {

		$UserExtension = "x" + $UserExtension

	}

	$OKButton = Config-Button -horizontalPosition 680 -verticalPosition 500 -width 75 -height 23 -text "OK" -tabIndex 2
	$OKButton.Add_Click({

		Handle-OKButtonClick

	})
	$Form2.Controls.Add($OKButton)

	$cancelButton = Config-Button -horizontalPosition 765 -verticalPosition 500 -width 75 -height 23 -text "Cancel" -tabIndex 3
	$cancelButton.Add_Click({

		$launchForm.close()
		$newMailBoxForm.close()
		$Form2.close()

	})
	$Form2.Controls.Add($cancelButton)

	$backButton = Config-Button -horizontalPosition 605 -verticalPosition 500 -width 75 -height 23 -text "< Back" -tabIndex 4
	$backButton.Add_Click({

		$Form2.visible = $False; 

		if ($disable) {

			$launchForm.visible = $True

		} else {

			$newMailBoxForm.visible = $True

		}

	})
	$Form2.Controls.Add($backButton)

	$finishButton = Config-Button -horizontalPosition -verticalPosition -width -height -text -tabIndex
	$finishButton.Add_Click({

			$Form2.Refresh()
			Start-Sleep -s 1
			$newMailBoxForm.close();
			$Form2.close();
			$launchForm.visible = $True

	})
	$Form2.Controls.Add($finishButton)

	# Job Status label
	$jobStatusLabel = Config-Label -horizontalPosition 420 -verticalPosition 40 -width 120 -height 20 -text "Job Status:" -fontStyle [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)
	$Form2.Controls.Add($jobStatusLabel)

	# Error Message label
	$errorMessage1 = Config-Label -horizontalPosition 420 -verticalPosition 70 -width 500 -height 20 -text $null
	$errorMessage1.ForeColor = "Green"
	$Form2.Controls.Add($errorMessage1)

	$errorMessage2 = Config-Label -horizontalPosition 420 -verticalPosition 90 -width 500 -height 20 -text $null
	$errorMessage2.ForeColor = "Green"
	$Form2.Controls.Add($errorMessage2)
	
	$errorMessage3 = Config-Label -horizontalPosition 420 -verticalPosition 110 -width 500 -height 20 -text $null
	$errorMessage3.ForeColor = "Green"
	$Form2.Controls.Add($errorMessage3)

	$errorMessage4 = Config-Label -horizontalPosition 420 -verticalPosition 130 -width 500 -height 20 -text $null
	$errorMessage4.ForeColor = "Green"
	$Form2.Controls.Add($errorMessage4)

	$errorMessage5 = Config-Label -horizontalPosition 420 -verticalPosition 150 -width 500 -height 20 -text $null
	$errorMessage5.ForeColor = "Green"
	$Form2.Controls.Add($errorMessage5)

	$errorMessage6 = Config-Label -horizontalPosition 420 -verticalPosition 170 -width 500 -height 20 -text $null
	$errorMessage6.ForeColor = "Green"
	$Form2.Controls.Add($errorMessage6)

	$errorMessage7 = Config-Label -horizontalPosition 420 -verticalPosition 190 -width 500 -height 20 -text $null
	$errorMessage7.ForeColor = "Green"
	$Form2.Controls.Add($errorMessage7)

	$errorMessage8 = Config-Label -horizontalPosition 420 -verticalPosition 210 -width 500 -height 20 -text $null
	$errorMessage8.ForeColor = "Green"
	$Form2.Controls.Add($errorMessage8)

	$errorMessage9 = Config-Label -horizontalPosition 420 -verticalPosition 230 -width 500 -height 20 -text $null
	$errorMessage9.ForeColor = "Green"
	$Form2.Controls.Add($errorMessage9)

	$errorMessage10 = Config-Label -horizontalPosition 420 -verticalPosition 250 -width 500 -height 20 -text $null
	$errorMessage10.ForeColor = "Green"
	$Form2.Controls.Add($errorMessage10)

	$errorMessage11 = Config-Label -horizontalPosition 420 -verticalPosition 270 -width 500 -height 20 -text $null
	$errorMessage11.ForeColor = "Green"
	$Form2.Controls.Add($errorMessage11)

	$errorMessage12 = Config-Label -horizontalPosition 420 -verticalPosition 290 -width 500 -height 20 -text $null
	$errorMessage12.ForeColor = "Green"
	$Form2.Controls.Add($errorMessage12)

	$errorMessage13 = Config-Label -horizontalPosition 420 -verticalPosition 310 -width 500 -height 20 -text $null
	$errorMessage13.ForeColor = "Green"
	$Form2.Controls.Add($errorMessage13)

	$errorMessage14 = Config-Label -horizontalPosition 420 -verticalPosition 330 -width 500 -height 20 -text $null
	$errorMessage14.ForeColor = "Green"
	$Form2.Controls.Add($errorMessage14)

	$errorMessage15 = Config-Label -horizontalPosition 420 -verticalPosition 350 -width 500 -height 20 -text $null
	$errorMessage15.ForeColor = "Green"
	$Form2.Controls.Add($errorMessage15)

	$errorMessage16 = Config-Label -horizontalPosition 420 -verticalPosition 370 -width 500 -height 20 -text $null
	$errorMessage16.ForeColor = "Green"
	$Form2.Controls.Add($errorMessage16)

	$errorMessage17 = Config-Label -horizontalPosition 420 -verticalPosition 390 -width 500 -height 20 -text $null
	$errorMessage17.ForeColor = "Green"
	$Form2.Controls.Add($errorMessage17)

	$errorMessage18 = Config-Label -horizontalPosition 420 -verticalPosition 410 -width 500 -height 20 -text $null
	$errorMessage18.ForeColor = "Green"
	$Form2.Controls.Add($errorMessage18)

	$errorMessage19 = Config-Label -horizontalPosition 420 -verticalPosition 430 -width 500 -height 20 -text $null
	$errorMessage19.ForeColor = "Green"
	$Form2.Controls.Add($errorMessage19)

	#Title label
	$titleLabel = Config-Label -horizontalPosition 20 -verticalPosition 40 -width 500 -height 20 -text "Click OK to " + $title.ToLower() + "." -fontStyle [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold)
	$Form2.Controls.Add($titleLabel)

	# Organizational unit label
	$organizationalUnitLabel = Config-Label -horizontalPosition 20 -verticalPosition 70 -width 120 -height 20 -text "Organizational Unit:"
	$Form2.Controls.Add($organizationalUnitLabel)

	$userOrganizationalUnitLabel = Config-Label -horizontalPosition 150 -verticalPosition 70 -width 350 -height 20 -text $UserOU
	$Form2.Controls.Add($userOrganizationalUnitLabel)

	# Name label
	$nameLabel = Config-Label -horizontalPosition 20 -verticalPosition 100 -width 120 -height 20 -text "Name:"
	$Form2.Controls.Add($nameLabel)

	$userNameLabel = Config-Label -horizontalPosition 150 -verticalPosition 100 -width 350 -height 20 -text "$UserFirst $UserLast" -fontFamily "ArialNarrow"
	$Form2.Controls.Add($userNameLabel)

	# User Logon Name Label
	$logonNameLabel = Config-Label -horizontalPosition 20 -verticalPosition 130 -width 120 -height 20 -text "User Logon Name:" -fontFamily "ArialNarrow"
	$Form2.Controls.Add($logonNameLabel)

	$userLogonNameLabel = Config-Label -horizontalPosition 150 -verticalPosition 130 -width 350 -height 20 -text "$UserAlias$UserDomain"
	$Form2.Controls.Add($userLogonNameLabel)

	# Password Label
	$passwordLabel = Config-Label -horizontalPosition 20 -verticalPosition 160 -width 120 -height 20 -text "$Password:"
	$Form2.Controls.Add($passwordLabel)

	$userPasswordLabel = Config-Label -horizontalPosition 150 -verticalPosition 160 -width 200 -height 20 -text "$UserPassword:"
	$Form2.Controls.Add($userPasswordLabel)

	# Server label
	$serverLabel = Config-Label -horizontalPosition 20 -verticalPosition 190 -width 120 -height 20 -text "Server / Database:"
	$Form2.Controls.Add($serverLabel)

	$userServerLabel = Config-Label -horizontalPosition 150 -verticalPosition 190 -width 350 -height 20 -text ""
	if ($MailboxServer -and $Database) {

		$userServerLabel.Text = "$MailboxServer/$Database"

	}
	$Form2.Controls.Add($userServerLabel)

	# Email label
	$emailLabel = Config-Label -horizontalPosition 20 -verticalPosition 220 -width 130 -height 20 -text "E-mail:"
	$Form2.Controls.Add($emailLabel)

	$userEmailLabel = Config-Label -horizontalPosition 150 -verticalPosition 220 -width 350 -height 20 -text "$UserEmail$UserEmailDomain" -fontFamily "ArialNarrow"
	$Form2.Controls.Add($userEmailLabel)

	# Distribution Group label
	$distributionGroupLabel = Config-Label -horizontalPosition 20 -verticalPosition 250 -width 120 -height 20 -text "Distribution Group:"
	$Form2.Controls.Add($distributionGroupLabel)

	$userDistributionGroupLabel = Config-Label -horizontalPosition 150 -verticalPosition 250 -width 350 -height 20 -text "$GroupName_DG"
	$Form2.Controls.Add($userDistributionGroupLabel)

	# Office label
	$officeLabel = Config-Label -horizontalPosition 20 -verticalPosition 280 -width 120 -height 20 -text "Office:"
	$Form2.Controls.Add($officeLabel)

	$userOfficeLabel = Config-Label -horizontalPosition 150 -verticalPosition 280 -width 350 -height 20 -text "$UserOffice"
	$Form2.Controls.Add($userOfficeLabel)

	# Phone label
	$phoneLabel = Config-Label -horizontalPosition 20 -verticalPosition 310 -width 120 -height 20 -text "Phone:"
	$Form2.Controls.Add($phoneLabel)

	$userPhoneLabel = Config-Label -horizontalPosition 150 -verticalPosition 310 -width 350 -height 20 -text "$UserPhone $UserExtension".Trim()
	$Form2.Controls.Add($userPhoneLabel)

	# DFS label
	$DFSLabel = Config-Label -horizontalPosition 20 -verticalPosition 340 -width 120 -height 20 -text "DFS Group:"
	$Form2.Controls.Add($DFSLabel)

	$userDFSLabel = Config-Label -horizontalPosition 150 -verticalPosition 340 -width 350 -height 20 -text "$GroupName_DFS".Trim()
	$Form2.Controls.Add($userDFSLabel)

	# Vasto label
	$FSLabel = Config-Label -horizontalPosition 20 -verticalPosition 370 -width 120 -height 20 -text "Vasto Group:"
	$Form2.Controls.Add($FSLabel)

	$userFSLabel = Config-Label -horizontalPosition 150 -verticalPosition 370 -width 350 -height 20 -text "$GroupName_Vasto".Trim()
	$Form2.Controls.Add($userFSLabel)

	# Websense department label
	$websenseDepartmentLabel = Config-Label -horizontalPosition 20 -verticalPosition 400 -width 130 -height 20 -text "Websense Department:"
	$Form2.Controls.Add($websenseDepartmentLabel)

	$userWebsenseDepartmentLabel = Config-Label -horizontalPosition 150 -verticalPosition 400 -width 550 -height 20 -text "$GroupName_Dept".Trim()
	$Form2.Controls.Add($userWebsenseDepartmentLabel)

	# Websense label
	$websenseLabel = Config-Label -horizontalPosition 20 -verticalPosition 430 -width 130 -height 20 -text "Websense List:"
	$Form2.Controls.Add($websenseLabel)

	$userWebsenseLabel = Config-Label -horizontalPosition 150 -verticalPosition 430 -width 350 -height 20 -text "$GroupName_List".Trim()
	$Form2.Controls.Add($userWebsenseLabel)

	# Email to label
	$emailToLabel = Config-Label -horizontalPosition 20 -verticalPosition 473 -width 60 -height 20 -text "Email To:"
	$Form2.Controls.Add($emailToLabel)

	# Email to IT Admin label
	$emailToITAdminLabel = Config-Label -horizontalPosition 95 -verticalPosition 473 -width 50 -height 20 -text "IT Admin"
	$Form2.Controls.Add($emailToITAdminLabel)

	# IT Email Check Box
	$emailToITAdminCheckBox = Config-CheckBox -horizontalPosition 80 -verticalPosition 470 -width 40 -height 20 -checked $True -tabIndex 14
	$Form2.Controls.Add($emailToITAdminCheckBox)

	# EmailToHR Label
	$emailToHRLabel = Config-Label -horizontalPosition 165 -verticalPosition 473 -width 30 -height 20 -text "HR"
	$Form2.Controls.Add($emailToHRLabel)

	# HR Email Check Box
	$emailToHRCheckBox = Config-CheckBox -horizontalPosition 150 -verticalPosition 470 -width 40 -height 20 -tabIndex 14
	$Form2.Controls.Add($emailToHRCheckBox)
	if ($UserAlias -like "*w_*") {

		$emailToHRLabel.Enabled = $True
		$emailToHRCheckBox.Checked = $True
		$emailToHRCheckBox.Enabled = $True

	} else {

		$emailToHRLabel.Enabled = $False
		$emailToHRCheckBox.Checked = $False
		$emailToHRCheckBox.Enabled = $False

	}

	# EmailToWuHan Label
	switch ($UserAlias) {

		{ ($_ -like "*w_*") } { 

			$emailToBranch = "admin.wh@newbiiz.com"

		}

		{ ($_ -like "*c_*") } { 

			$emailToBranch = "Brian.Li@malabs.com"

		}

		{ ($_ -like "*g_*") } { 

			$emailToBranch = "Christina.Tay@malabs.com"

		}

		{ ($_ -like "*i_*") } { 

			$emailToBranch = "Keith.Yarbrough@malabs.com"

		}

		{ ($_ -like "*n_*") } {

			$emailToBranch = "davidc@malabs.com"

		}

		{ ($_ -like "*m_*") } { 

			$emailToBranch = "fari@malabs.com"

		}

		{ ($_ -notlike '*_*') } {

			$branchEmailCheckBox.Enabled = $False
			$branchEmailCheckBox.Checked = $False

		}

	}

	$emailToBranchLabel = Config-Label -horizontalPosition 95 -verticalPosition 498 -width 220 -height 20 -text "$emailToBranch"
	$Form2.Controls.Add($emailToBranchLabel)

	$branchEmailCheckBox = Config-CheckBox -horizontalPosition 80 -verticalPosition 495 -width 15 -height 20 -checked $True
	$Form2.Controls.Add($branchEmailCheckBox)

	$Form2.Topmost = $True
	$Form2.Add_Shown({ 

		$Form2.Activate()

	})
	$Form2.ShowDialog()

}

function Handle-KeyDown {

	if ($_.KeyCode -eq "Enter") {

		if ($OKButton.visible) {

			if ($disable) {

				$logfile = $UserAlias + "-disabled.log"
				DisableUser

			} else {

				$logfile = $UserAlias + "-created.log"
				CreateMailbox

			}

		}

		if ($finishButton.visible) {

			$newMailBoxForm.close()
			$Form2.close()
			$launchForm.visible = $True

		}

	} elseif ($_.KeyCode -eq "Escape") {

		$launchForm.close()
		$newMailBoxForm.close()
		$Form2.close()

	}

}

function Handle-OKButtonClick {

	if ($disable) {

			$logfile = $UserAlias + "-disabled.log"
			DisableUser

	} else {

		$logfile = $UserAlias + "-created.log"
		$CheckIT = $emailToITAdminCheckBox.Checked
		$CheckHR = $emailToHRCheckBox.Checked
		$CheckBranch = $branchEmailCheckBox.Checked
		CreateMailbox

		Start-Sleep -s 1

		if ($GroupName_DFS.Length -ne 0 -or $GroupName_Vasto.Length -ne 0 -or $GroupName_Dept.Length -ne 0 -or $GroupName_List.Length -ne 0) {

			Add-Group

		}

	}

}