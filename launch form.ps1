<##############################
	Launch Form
################################>

<#
	Show "Exchange Manager" launch panel form
#>
function Show-launchForm {

	<#
		Config launch panel form and handle default press-key behaviors
	#>
	$launchForm = New-Object System.Windows.Forms.Form
	$launchForm.Text = $title
	$launchForm.Size = New-Object System.Drawing.Size (500, 610)
	$launchForm.StartPosition = "CenterScreen"
	$launchForm.KeyPreview = $True
	$launchForm.Add_KeyDown({

		if ($_.KeyCode -eq "Enter") {

			 Handle-Options

		} elseif ($_.KeyCode -eq "Escape") {

			$launchForm.close()

		}

	})

	<#
		Config "Next" button and its behaviors
	#>
	$nextButton = Config-Button -horizontalPosition 280 -verticalPosition 520 -width 75 -height 23 -text "Next >" -tabIndex 7
	$nextButton.Add_Click({

		 Handle-Options

	})
	$launchForm.Controls.Add($NextButton)

	<#
		Config "Cancel" button and its behaviors
	#>
	$cancelButton = New-Object System.Windows.Forms.Button
	$cancelButton = Config-Button -horizontalPosition 365 -verticalPosition 520 -width 75 -height 23 -text "Cancel" -tabIndex 8
	$cancelButton.Add_Click({

		$launchForm.close()

	})
	$launchForm.Controls.Add($cancelButton)

	<#
		Config "Error message"
	#>	
	$errorMessage = Config-ErrorMessage	-horizontalPosition 20 -verticalPosition 10 -width 500 -height 20
	$launchForm.Controls.Add($errorMessage)


	<#
		Config title label
	#>
	$titleLabel = Config-Label -horizontalPosition 20 -verticalPosition 40 -width 450 -height 20 -text "$Forest" -fontFamily "Calibri" -fontSize 15.75 -fontStyle ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic))
	$launchForm.Controls.Add($titleLabel)

	<# 
		Config "Create a new mailbox" radio button
	#>
	$newMailboxRadioButton = Config-RadioButton -horizontalPosition 40 -verticalPosition 100 -width 200 -height 20 -text "Create a new mailbox" -tabIndex 1
	$newMailboxRadioButton.Checked = $True
	$launchForm.Controls.Add($newMailboxRadioButton)

	<#
		Config "Create mailbox for an existing user" radio button
	#>
	$mailboxForExistingUserRadioButton = Config-RadioButton -horizontalPosition 40 -verticalPosition 120 -width 200 -height 20 -text "Create mailbox for an existing user" -tabIndex 2
	$launchForm.Controls.Add($mailboxForExistingUserRadioButton)

	<#
		"Disable user" radio button
	#>
	$disableUserRadioButton = Config-RadioButton -horizontalPosition 40 -verticalPosition 140 -width 200 -height 20 -text "Disable user" -tabIndex 3
	$launchForm.Controls.Add($disableUserRadioButton)

	<#
		"Create a new distribution group" Radio Button
	#>
	$newDistributionGroupRadioButton = Config-RadioButton -horizontalPosition 40 -verticalPosition 230 -width 200 -height 20 -text "Create a new distribution group" -tabIndex 4
	$launchForm.Controls.Add($newDistributionGroupRadioButton)

	<#
		"Create distribution group for an existing group" Radio Button
	#>
	$distributionGroupForExistingGroupRadioButton = Config-RadioButton -horizontalPosition 40 -verticalPosition 250 -width 300 -height 20 -text "Create distribution group for an existing group" -tabIndex 5
	$launchForm.Controls.Add($distributionGroupForExistingGroupRadioButton)

	<# 
		"Create a new mail contact" Radio Button
	#>
	$newMailContactRadioButton = Config-RadioButton -horizontalPosition 40 -verticalPosition 340 -width 200 -height 20 -text "Create a new mail contact" -tabIndex 6
	$launchForm.Controls.Add($newMailContactRadioButton)

	<#
		"Add a new membership" radio Button
	#>
	$newMembershipRadioButton = Config-RadioButton -horizontalPosition 40 -verticalPosition 440 -width 200 -height 20 -text "Add a new membership" -tabIndex 7
	$launchForm.Controls.Add($newMembershipRadioButton)

	<#
		"Mailbox" group box
	#>
	$mailboxGroupbox = Config-GroupBox -horizontalPosition 20 -verticalPosition 70 -width 420 -height 110 -text "Mailbox"
	$launchForm.Controls.Add($mailboxGroupbox)

	<#
		"Distribution Group" group box
	#>
	$distributionGroupGroupBox = Config-GroupBox -horizontalPosition 20 -verticalPosition 200 -width 420 -height 90 -text "Distribution Group"
	$launchForm.Controls.Add($distributionGroupGroupBox)

	<#
		"Contact" group box
	#>
	$mailContactGroupBox = Config-GroupBox -horizontalPosition 20 -verticalPosition 310 -width 420 -height 70 -text "Mail Contact"
	$launchForm.Controls.Add($mailContactGroupBox)

	<#
		"Add Group Membership" group box (DFS and VASTO)
	#>
	$addGroupMemberShipBox = Config-GroupBox -horizontalPosition 20 -verticalPosition 410 -width 420 -height 70 -text "Add Group Membership"
	$launchForm.Controls.Add($addGroupMemberShipBox)

	#$launchForm.Topmost = $True
	$launchForm.Add_Shown({

		$launchForm.Activate() 

	})
	[void]$launchForm.ShowDialog()

	return $launchForm

}

function Handle-Options {

	if ($newMailboxRadioButton.Checked) {

		$new = $True;
		$title = $title_newuser

	} else {

		$new = $False

	}
	
	if ($mailboxForExistingUserRadioButton.Checked) {

		$title = $title_existinguser

	}
	
	if ($disableUserRadioButton.Checked) {

		$disable = $True
		$title = $title_disableuser

	} else {

		$disable = $False

	}

	if ($newMailContactRadioButton.Checked) {

		$contact = $True;
		$title = $title_newcontact

	} else {

		$contact = $False

	}

	$found = $False

	if ($contact) {

		ShowForm5

	} elseif ($newDistributionGroupRadioButton.Checked -or $distributionGroupForExistingGroupRadioButton.Checked) {

		if ($newDistributionGroupRadioButton.Checked) {

			$newgroup = $True
			$title = $title_newdg

		} else {

			$newgroup = $False
			$title = $title_existingdg

		}

		ShowForm3

	} elseif ($newMembershipRadioButton.Checked) {

		$Form.close()
		ShowFormMembership
		#Invoke-Expression C:\Users\rayj\Desktop\List_Import.ps1   #Call the other script from main method
		#Invoke-Item (start powershell ("C:\Users\rayj\Desktop\List_Import.ps1"))

	} else {

		ShowForm1 -new $new -title $title -disable $disable -found $found

	}

}