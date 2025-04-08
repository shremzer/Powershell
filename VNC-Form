###################################################################################################################
###### Creating a form in PowerShell with a drop-down list of users, selecting which, a connection occurs  ########
###################################################################################################################
 
 Add-Type -AssemblyName System.Windows.Forms

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.Text = "                                          CONNECTION BY VNC"
$form.Size = New-Object System.Drawing.Size(500, 400)

# Create a ComboBox
$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(90, 20)
$comboBox.Size = New-Object System.Drawing.Size(260, 20)
$comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

# Import hash-table
$table = Import-Clixml -Path "E:\hash.cli"

# Add names to the ComboBox
$table.GetEnumerator() | Sort-Object -Property key  | ForEach-Object {
    $comboBox.Items.Add($_.Key)
}

# Script-scoped variable to store the selected item
$global:selectedItem = $null

# Add an event handler for when an item is selected
$comboBox.Add_SelectedIndexChanged({
    $global:selectedItem = $comboBox.SelectedItem
})

# Create a button to connect to VNC
$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(150, 160)
$button.Size = New-Object System.Drawing.Size(150, 60)
$button.Text = "CONNECT"

# Add the button to the form
$form.Controls.Add($button)

# Add the ComboBox to the form
$form.Controls.Add($comboBox)

# Add an event handler for the button click
$button.Add_Click({
    if (-not [string]::IsNullOrEmpty($global:selectedItem)) {
        # Find the corresponding IP address or other details from the hash table
        $selectedEntry = $table[$global:selectedItem]
        if ($selectedEntry) {
            $ip = $table.$global:selectedItem   # Assuming the value is the IP address or identifier
            cd "C:\Program Files (x86)\uvnc bvba\UltraVNC\"; .\vncviewer.exe  $ip  -password Fenix1985
        } else {
            [System.Windows.Forms.MessageBox]::Show("No matching entry found for the selected name.")
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select an item from the list.")
    }
})



# Create a button to close the form
$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Location = New-Object System.Drawing.Point(100, 300)
$closeButton.Size = New-Object System.Drawing.Size(260, 30)
$closeButton.Text = "Close"
$closeButton.Add_Click({ $form.Close() })

# Add the close button to the form
$form.Controls.Add($closeButton)

# Show the form
$form.ShowDialog()

$hostname=[System.Net.Dns]::GetHostEntry($table.$global:selectedItem).HostName


$hostname=$hostname.Trim()



# Define the password
$password = ConvertTo-SecureString -String "Fenix1985" -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential("$hostname\admin", $password)


# IF not in Domain
Invoke-command -computername $hostname -scriptblock {msg /v /time:300 */server:$hostname "`tПРОДОЛЖАЙТЕ РАБОТАТЬ   "} -Credential $cred -ErrorAction SilentlyContinue



# IF in AD
Invoke-command -computername $hostname -scriptblock {msg /v /time:300 */server:$hostname "`tПРОДОЛЖАЙТЕ РАБОТАТЬ "} -ErrorAction SilentlyContinue
