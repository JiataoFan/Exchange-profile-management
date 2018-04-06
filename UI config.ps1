<################################################################################
# UI element configuration functions library
################################################################################>


<#
	Config form window
#>
function Config-Form {

	param([int]$width, [int]$height, [string]$text, [string]$windowState = "Normal", [string]$startPosition = "CenterScreen", [bool]$keyPreview = $True)

	$form = New-Object System.Windows.Forms.Form
	$form.Size = New-Object System.Drawing.Size ($width, $height)
	$form.Text = $text
	$form.WindowState = $windowState
	$form.StartPosition = $startPosition
	$form.KeyPreview = $keyPreview

	return $form

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

	param([int]$horizontalPosition, [int]$verticalPosition, [int]$width, [int]$height, [string]$text,[string]$fontFamily = "Arial", [int]$fontSize = 8, [System.Drawing.FontStyle]$fontStyle = [System.Drawing.FontStyle]::Regular)

	$label = New-Object System.Windows.Forms.Label
	$label.Location = New-Object System.Drawing.Size ($horizontalPosition, $verticalPosition)
	$label.Size = New-Object System.Drawing.Size ($width, $height)
	$label.Text = $text
	$label.Font = New-Object System.Drawing.Font($fontFamily, $fontSize, $fontStyle)

	return $label

}

<#
	Config group box
#>
function Config-GroupBox {

	param([int]$horizontalPosition, [int]$verticalPosition, [int]$width, [int]$height, [string]$text)

	$groupBox = New-Object System.Windows.Forms.GroupBox
	$groupBox.Location = New-Object System.Drawing.Size ($horizontalPosition, $verticalPosition)
	$groupBox.Size = New-Object System.Drawing.Size ($width, $height)
	$groupBox.Text = [string]$text

	return $groupBox

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

<#
	Config text box
#>
function Config-TextBox {

	param([int]$horizontalPosition, [int]$verticalPosition, [int]$width, [int]$height, [int]$tabIndex, [int]$maxLength = 40)

	$textBox = New-Object System.Windows.Forms.TextBox
	$textBox.Location = New-Object System.Drawing.Size ($horizontalPosition, $verticalPosition)
	$textBox.Size = New-Object System.Drawing.Size ($width, $height)
	$textBox.TabIndex = $tabIndex
	$textBox.MaxLength = $maxLength

	return $textBox

}

<#
	Config combo box
#>
function Config-ComboBox {

	param([int]$horizontalPosition, [int]$verticalPosition, [int]$width, [int]$height)

	$comboBox = New-Object System.Windows.Forms.ComboBox
	$comboBox.Location = New-Object System.Drawing.Size ($horizontalPosition, $verticalPosition)
	$comboBox.Size = New-Object System.Drawing.Size ($width, $height)

	return $comboBox

}

<#
	Config checked list box
#>
function Config-CheckedListbox {

	param([int]$horizontalPosition, [int]$verticalPosition, [int]$width, [int]$height, [bool]$checkOnClick = $False, [bool]$horizontalScrollbar = $False, $enabled = $True)

	$checkedListBox = New-Object System.Windows.Forms.CheckedListbox
	$checkedListBox.Location = New-Object System.Drawing.Size ($horizontalPosition, $verticalPosition)
	$checkedListBox.Size = New-Object System.Drawing.Size ($width, $height)
	$checkedListBox.CheckOnClick = $checkOnClick
	$checkedListBox.HorizontalScrollbar = $horizontalScrollbar
	$checkedListBox.Enabled = $enabled

	return $checkedListBox

}

<#
	Config check box
#>
function Config-CheckBox {

	param([int]$horizontalPosition, [int]$verticalPosition, [int]$width, [int]$height, [bool]$checked = $False, [int]$tabIndex, [bool]$enabled = $True)

	$checkBox = New-Object System.Windows.Forms.CheckBox
	$checkBox.Location = New-Object System.Drawing.Size ($horizontalPosition, $verticalPosition)
	$checkBox.Size = New-Object System.Drawing.Size ($width, $height)
	$checkBox.Checked = $True
	$checkBox.TabIndex = $tabIndex
	$checkBox.Enabled = $enabled

	return $checkBox

}