<##############################
	Tree view
################################>
function Add-Node {

	param($selectedNode, $name)

	$newNode = New-Object System.Windows.Forms.TreeNode
	$newNode.Name = $name
	$newNode.Text = $name
	$selectedNode.Nodes.Add($newNode) | Out-Null

	return $newNode

}

function Get-NextLevel {

	param($selectedNode, $dn)

	$path = [adsi]("LDAP://" + $dn)
	$OU = New-Object System.DirectoryServices.DirectorySearcher ($path)
	$OU.SearchScope = "onelevel"
	$OU.Filter = "(&(objectClass=organizationalUnit))"
	$OUs = $OU.FindAll()

	if ($OUs -eq $null) {

		$node = Add-Node $selectedNode $path

	} else {

		foreach ($person in $OUs) {

			[string]$ent = $person.properties.Name
			$node = Add-Node $selectedNode $ent
			[string]$dn = $person.properties.distinguishedname
			Get-NextLevel $node $dn

		}

	}

}

function Build-TreeView {

	if ($treeNodes) {

		$treeview1.Nodes.Remove($treeNodes)
		$newMailBoxForm.Refresh()

	}

	$treeNodes = New-Object System.Windows.Forms.TreeNode
	$treeNodes.Text = "Active Directory Hierarchy"
	$treeNodes.Name = "Active Directory Hierarchy"
	$treeNodes.Tag = "root"
	$treeView1.Nodes.Add($treeNodes) | Out-Null

	if ($new) {

		$treeView1.add_AfterSelect({

			[string]$fullpath = $this.SelectedNode.FullPath
			if ($fullpath -eq "Active Directory Hierarchy") {

				$ldappath = ""

			} else {

				[string]$ldappath = $fullpath.substring($fullpath.IndexOf("\"), $fullpath.Length - $fullpath.IndexOf("\"))
				$ldappath = $ldappath.Replace('\', '/')
				$textbox1.Text = $ldappath
				[void]$OU.Items.Add("$ldappath")
				$OU.Text = $ldappath
				$Office.Items.Clear()
				$MainCompany.SelectedIndex = 0

			}

		})

	}

	#Generate Module nodes 
	$OUs = Get-NextLevel $treeNodes $strDomainDN

	$treeNodes.Expand()

}