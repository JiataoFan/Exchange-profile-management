
$form = New-Object system.Windows.Forms.Form

$form.KeyPreview = $true

$keyPress = [System.Windows.Forms.KeyEventHandler]{

	#Event Argument: $_ = [System.Windows.Forms.KeyPressEventArgs]
	#if($_.KeyCode -eq "Enter"){

		[void][System.Windows.Forms.MessageBox]::Show('Enter key enterd' + $_.KeyChar)

	#}

}

$form.add_keydown($keyPress)

$form.ShowDialog()


function Config-Label {

	param([int]$horizontalPosition, [int]$verticalPosition, [int]$width, [int]$height, [string]$text, [string]$fontFamily, [int]$fontSize, [System.Drawing.FontStyle]$fontStyle)

	$label = New-Object System.Windows.Forms.Label
	$label.Location = New-Object System.Drawing.Size ($horizontalPosition, $verticalPosition)
	$label.Size = New-Object System.Drawing.Size ($width, $height)
	$label.Text = $text
	$label.Font = New-Object System.Drawing.Font()

	return $label

}


$label = Config-Label 10 10 10 10