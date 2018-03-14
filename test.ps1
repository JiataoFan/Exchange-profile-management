
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
