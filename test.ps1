
$form = New-Object system.Windows.Forms.Form

$form.KeyPreview = $true

$keyPress = [System.Windows.Forms.KeyEventHandler]{

	#Event Argument: $_ = [System.Windows.Forms.KeyPressEventArgs]
	#if($_.KeyCode -eq "Enter"){

		[void][System.Windows.Forms.MessageBox]::Show('Enter key enterd' + $_.KeyChar)

	#}

}

$form.add_keydown($keyPress)



function Config-ComboBox {

	param([int]$horizontalPosition, [int]$verticalPosition, [int]$width, [int]$height)

	$comboBox = New-Object System.Windows.Forms.ComboBox
	$comboBox.Location = New-Object System.Drawing.Size ($horizontalPosition, $verticalPosition)
	$comboBox.Size = New-Object System.Drawing.Size ($width, $height)
	$comboBox.Items.AddRange(("asdf", "aaaaaaaaaaaaaaaa"))

	return $comboBox

}


$comboBox = Config-ComboBox 10 10 100 100

$form.Controls.Add($comboBox)

$form.ShowDialog()