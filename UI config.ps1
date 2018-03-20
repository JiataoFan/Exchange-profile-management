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