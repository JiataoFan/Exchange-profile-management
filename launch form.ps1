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

			if ($NewUser.Checked) {

				$new = $True;
				$title = $title_newuser

			} else {

				$new = $False

			}

			if ($NewContact.Checked) {

				$contact = $True;
				$title = $title_newcontact

			} else {

				$contact = $False

			}

			if ($DisableUser.Checked) {

				$disable = $True;
				$title = $title_disableuser

			} else {

				$disable = $False

			}

			if ($ExistUser.Checked) {

				$title = $title_existinguser

			}

			$found = $False
			if ($contact) {

				ShowForm5

			} elseif ($NewDG.Checked -or $ExistGroup.Checked) {

				if ($NewDG.Checked) {

					$newgroup = $True;
					$title = $title_newdg

				} else {

					$newgroup = $False;
					$title = $title_existingdg

				}

				ShowForm3

			} else {

				ShowForm1

			}

		}

	})

	$launchForm.Add_KeyDown({

		if ($_.KeyCode -eq "Escape") {

			$launchForm.close()

		}

	})

	<#
		Config "Next" button and its behaviors
	#>
	$nextButton = Config-Button -horizontalPosition 280 -verticalPosition 520 -width 75 -height 23 -text "Next >" -tabIndex 7
	$nextButton.Add_Click({

		if ($NewUser.Checked) {

			$new = $True;
			$title = $title_newuser

		} else {

			$new = $False

		}

		if ($NewContact.Checked) {

			$contact = $True;
			$title = $title_newcontact

		} else {

			$contact = $False

		}

		if ($DisableUser.Checked) {

			$disable = $True;
			$title = $title_disableuser

		} else {

			$disable = $False

		}

		if ($ExistUser.Checked) {

			$title = $title_existinguser

		}

		$found = $False

		if ($contact) {

			ShowForm5

		} elseif ($NewDG.Checked -or $ExistGroup.Checked) {

			if ($NewDG.Checked) {

				$newgroup = $True;
				$title = $title_newdg

			} else {

				$newgroup = $False;
				$title = $title_existingdg

			}

			ShowForm3

		} elseif ($NewMembership.Checked) {

			$Form.close()
			ShowFormMembership
			#Invoke-Expression C:\Users\rayj\Desktop\List_Import.ps1   #Call the other script from main method
			#Invoke-Item (start powershell ("C:\Users\rayj\Desktop\List_Import.ps1"))

		} else {

			ShowForm1

		}

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
	$titleLabel = Config-Label -horizontalPosition 20 -verticalPosition 40 -width 450 -height 20 -text "$Forest"
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
	$mailboxGroupbox = Config-Box -horizontalPosition 20 -verticalPosition 70 -width 420 -height 110 -text "Mailbox"
	$launchForm.Controls.Add($mailboxGroupbox)

	<#
		"Distribution Group" group box
	#>
	$distributionGroupGroupBox = Config-Box -horizontalPosition 20 -verticalPosition 200 -width 420 -height 90 -text "Distribution Group"
	$launchForm.Controls.Add($distributionGroupGroupBox)

	<#
		"Contact" group box
	#>
	$mailContactGroupBox = Config-Box -horizontalPosition 20 -verticalPosition 310 -width 420 -height 70 -text "Mail Contact"
	$launchForm.Controls.Add($mailContactGroupBox)

	<#
		"Add Group Membership" group box (DFS and VASTO)
	#>
	$addGroupMemberShipBox = Config-Box -horizontalPosition 20 -verticalPosition 410 -width 420 -height 70 -text "Add Group Membership"
	$launchForm.Controls.Add($addGroupMemberShipBox)

	#$launchForm.Topmost = $True
	$launchForm.Add_Shown({

		$launchForm.Activate() 

	})
	[void]$launchForm.ShowDialog()

	return $launchForm

}

<#
	Config button
#>
function Config-Button {

	param([int]$horizontalPosition, [int]$verticalPosition, [int]$width, [int]$height, [string]$text, [int]$tabIndex)

	$button = New-Object System.Windows.Forms.Button
	$button.Location = New-Object System.Drawing.Size ($horizontalPosition, $verticalPosition)
	$button.Size = New-Object System.Drawing.Size ($width, $height)
	$button.Text = $text
	$button.TabIndex = $tabIndex

	return $button

}

<#
	Config a error message
#>
function Config-ErrorMessage {

	param([int]$horizontalPosition, [int]$verticalPosition, [int]$width, [int]$height)

	$errorMessageBox = New-Object System.Windows.Forms.Label
	$errorMessageBox.Location = New-Object System.Drawing.Size ($horizontalPosition, $verticalPosition)
	$errorMessageBox.Size = New-Object System.Drawing.Size ($width, $height)
	$errorMessageBox.ForeColor = "Red"

	return $errorMessageBox;

}

<#
	Config a label
#>
function Config-Label {

	param([int]$horizontalPosition, [int]$verticalPosition, [int]$width, [int]$height, [string]$text)

	$label = New-Object System.Windows.Forms.Label
	$label.Location = New-Object System.Drawing.Size ($horizontalPosition, $verticalPosition)
	$label.Size = New-Object System.Drawing.Size ($width, $height)
	$label.Text = "$Forest"
	$label.Font = New-Object System.Drawing.Font("Calibri", 15.75, ([System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Italic)), [System.Drawing.GraphicsUnit]::Point)

	return $label

}

<#
	Config group box
#>
function Config-Box {

	param([int]$horizontalPosition, [int]$verticalPosition, [int]$width, [int]$height, [string]$text)

	$box = New-Object System.Windows.Forms.GroupBox
	$box.Location = New-Object System.Drawing.Size ($horizontalPosition, $verticalPosition)
	$box.Size = New-Object System.Drawing.Size ($width, $height)
	$box.Text = [string]$text

	return $box

}

<#
	Config radio button
#>
function Config-RadioButton {

	param([int]$horizontalPosition, [int]$verticalPosition, [int]$width, [int]$height,  [string]$text, [int]$tabIndex)

	$radioButton = New-Object System.Windows.Forms.RadioButton
	$radioButton.Location = New-Object System.Drawing.Size ($horizontalPosition, $verticalPosition)
	$radioButton.Size = New-Object System.Drawing.Size ($width, $height)
	$radioButton.Text = $text
	$radioButton.TabIndex = $tabIndex
	$radioButton.TabStop = $true

	return $radioButton

}