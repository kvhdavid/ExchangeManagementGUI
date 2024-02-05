# Created by David Khachatryan 02/01/2024
# Exchange Management Script v1
#


# Prerequisite check
if (!(Get-PSSnapIn Microsoft.Exchange.Management.PowerShell.RecipientManagement -Registered -ErrorAction SilentlyContinue)) {
	throw "Please install the Exchange 2019 CU12 and above Management Tools-Only install. See: https://docs.microsoft.com/en-us/Exchange/manage-hybrid-exchange-recipients-with-management-tools"
	break
}

# Loads Exchange Management module
Add-PSSnapIn Microsoft.Exchange.Management.PowerShell.RecipientManagement

# Loads GUI module
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Enable Visual Styles 
[System.Windows.Forms.Application]::EnableVisualStyles()

function Create-Gui {
    # Create the main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Exchange Management Shell Menu"
    $form.Size = New-Object System.Drawing.Size(400,215)
    $form.StartPosition = "CenterScreen"

    # Create buttons
    $button1 = New-Object System.Windows.Forms.Button
    $button1.Location = New-Object System.Drawing.Point(20,20)
    $button1.Size = New-Object System.Drawing.Size(150,30)
    $button1.Text = "Create New Mailbox"
    $button1.Add_Click({ Create-NewMailbox })
    $form.Controls.Add($button1)

    $button2 = New-Object System.Windows.Forms.Button
    $button2.Location = New-Object System.Drawing.Point(20,70)
    $button2.Size = New-Object System.Drawing.Size(150,30)
    $button2.Text = "Add/Edit/View Aliases"
    $button2.Add_Click({ Add-Edit-ViewAliases })
    $form.Controls.Add($button2)

    $button3 = New-Object System.Windows.Forms.Button
    $button3.Location = New-Object System.Drawing.Point(20,120)
    $button3.Size = New-Object System.Drawing.Size(150,30)
    $button3.Text = "Search for a Mailbox"
    $button3.Add_Click({ Search-Mailbox })
    $form.Controls.Add($button3)

    $button4 = New-Object System.Windows.Forms.Button
    $button4.Location = New-Object System.Drawing.Point(200,20)
    $button4.Size = New-Object System.Drawing.Size(150,30)
    $button4.Text = "Initiate Azure Sync"
    $button4.Add_Click({ Initiate-DirSync })
    $form.Controls.Add($button4)

    $button5 = New-Object System.Windows.Forms.Button
    $button5.Location = New-Object System.Drawing.Point(200,70)
    $button5.Size = New-Object System.Drawing.Size(150,30)
    $button5.Text = "View All Mailboxes"
    $button5.Add_Click({ View-AllMailboxes })
    $form.Controls.Add($button5)

    $button6 = New-Object System.Windows.Forms.Button
    $button6.Location = New-Object System.Drawing.Point(200,120)
    $button6.Size = New-Object System.Drawing.Size(150,30)
    $button6.Text = "Exit"
    $button6.Add_Click({ $form.Close() })
    $form.Controls.Add($button6)

    # Show the form
    $form.ShowDialog()
}

function Create-NewMailbox {
    # Get non-Exchange users from Active Directory
    $nonExchangeUsers = Get-User -Filter "RecipientType -eq 'User' -and RecipientTypeDetails -ne 'DisabledUser'" | Select-Object DisplayName, SamAccountName

    # Create a new form for user selection
    $userSelectionForm = New-Object System.Windows.Forms.Form
    $userSelectionForm.Text = "Select a non-Exchange user"
    $userSelectionForm.Size = New-Object System.Drawing.Size(400,300)
    $userSelectionForm.StartPosition = "CenterScreen"

  # Create a listbox to display non-Exchange users with only display name shown
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(20,20)
$listBox.Size = New-Object System.Drawing.Size(350,200)
$userSelectionForm.Controls.Add($listBox)

# Add items to the listbox
foreach ($user in $nonExchangeUsers) {
    $listBox.Items.Add("$($user.DisplayName) - $($user.SamAccountName)")
}

# Create a button to select a user
$selectButton = New-Object System.Windows.Forms.Button
$selectButton.Location = New-Object System.Drawing.Point(20, 230)
$selectButton.Size = New-Object System.Drawing.Size(150, 30)
$selectButton.Text = "Select User"
$selectButton.Add_Click({
    $selectedText = $listBox.SelectedItem
    if ($selectedText -ne $null) {
        # Extract SamAccountName from the selected text
        $selectedUserSamAccountName = $selectedText -split ' - ' | Select-Object -Last 1

        # Prompt the user for the desired email address (alias) using a larger Windows Form input box
        $aliasInputForm = New-Object System.Windows.Forms.Form
        $aliasInputForm.Text = "Enter Alias"
        $aliasInputForm.Size = New-Object System.Drawing.Size(400, 250)

        $errorLabel = New-Object System.Windows.Forms.Label
        $errorLabel.Location = New-Object System.Drawing.Point(20, 170)
        $errorLabel.Size = New-Object System.Drawing.Size(350, 20)
        $errorLabel.ForeColor = [System.Drawing.Color]::Red
        $aliasInputForm.Controls.Add($errorLabel)

        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(20, 20)
        $label.Size = New-Object System.Drawing.Size(350, 20)
        $label.Text = "Enter the desired email address (Alias) for $selectedUserSamAccountName"
        $aliasInputForm.Controls.Add($label)

        # Display the example on the line below "Enter the desired e-mail address"
        $exampleLabel = New-Object System.Windows.Forms.Label
        $exampleLabel.Location = New-Object System.Drawing.Point(20, 50)
        $exampleLabel.Size = New-Object System.Drawing.Size(350, 20)
        $exampleLabel.Text = "Example: user@example.com"
        $aliasInputForm.Controls.Add($exampleLabel)

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(20, 80)
        $textBox.Size = New-Object System.Drawing.Size(350, 20)
        $aliasInputForm.Controls.Add($textBox)

        $okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(20, 120)
$okButton.Size = New-Object System.Drawing.Size(150, 30)
$okButton.Text = "OK"
$okButton.Add_Click({
    $global:aliasInput = $textBox.Text
    $aliasInputForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $aliasInputForm.Close()
})
$aliasInputForm.Controls.Add($okButton)

        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Location = New-Object System.Drawing.Point(200, 120)
        $cancelButton.Size = New-Object System.Drawing.Size(150, 30)
        $cancelButton.Text = "Cancel"
        $cancelButton.Add_Click({
            $aliasInputForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $aliasInputForm.Close()
        })
        $aliasInputForm.Controls.Add($cancelButton)

        $aliasInputForm.ShowDialog()

       try {
    if ($aliasInputForm.DialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        # Create the remoteRoutingAddress based on the alias input
        $remoteRoutingAddress = $global:aliasInput -replace '@example\.com', '@example0.mail.onmicrosoft.com'

        # Run the command to create a remote mailbox using the specified AD user and primary SMTP address
        # Replace this with your actual command
        Write-Host "Creating remote mailbox for $selectedUserSamAccountName with PrimarySmtpAddress: $global:aliasInput and RemoteRoutingAddress: $remoteRoutingAddress"

        # Your Enable-RemoteMailbox command
        Enable-RemoteMailbox -Identity $selectedUserSamAccountName -PrimarySmtpAddress $global:aliasInput -RemoteRoutingAddress $remoteRoutingAddress

        # Show success message
        [System.Windows.Forms.MessageBox]::Show("Mailbox created successfully for $selectedUserSamAccountName with PrimarySmtpAddress: $global:aliasInput", "Success", "OK", "Information")
        
        # Clear any previous errors
        $errorLabel.Text = ""
    }
}
catch {
    # Display the error in the GUI
    $errorLabel.Text = "Error creating remote mailbox: $_"
}
    }
    $userSelectionForm.Close()
})
$userSelectionForm.Controls.Add($selectButton)


    # Create a button to close the form
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(200,230)
    $closeButton.Size = New-Object System.Drawing.Size(150,30)
    $closeButton.Text = "Close"
    $closeButton.Add_Click({ $userSelectionForm.Close() })
    $userSelectionForm.Controls.Add($closeButton)

    # Show the user selection form
    $userSelectionForm.ShowDialog()
}

function Add-Edit-ViewAliases {
    # Create a form for user input
    $inputForm = New-Object System.Windows.Forms.Form
    $inputForm.Text = "Input user"
    $inputForm.Size = New-Object System.Drawing.Size(300, 150)
    $inputForm.startposition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # Create a label
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Enter username:"
    $label.Location = New-Object System.Drawing.Point(10, 17)
    $inputForm.Controls.Add($label)

    # Create a text box
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10, 40)
$textBox.Size = New-Object System.Drawing.Size(200, 20)
$inputForm.Controls.Add($textBox)

# Create an OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(10, 80)  # Adjusted the Y position
$okButton.Size = New-Object System.Drawing.Size(75, 23)  # Adjusted the button size
$okButton.Add_Click({
    $global:userToManage = $textBox.Text
    $inputForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $inputForm.Close()
})
$inputForm.Controls.Add($okButton)

# Create a Cancel button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Location = New-Object System.Drawing.Point(100, 80)  # Adjusted the X and Y position
$cancelButton.Size = New-Object System.Drawing.Size(75, 23)  # Adjusted the button size
$cancelButton.Add_Click({
    $inputForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $inputForm.Close()
})
$inputForm.Controls.Add($cancelButton)

    # Show the input form as a dialog
    $result = $inputForm.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # Retrieve the user to manage
        $userToManage = $global:userToManage

        # Get primary SMTP address for the specified user
        $primarySMTP = Get-RemoteMailbox -Identity $userToManage | Select-Object -ExpandProperty PrimarySmtpAddress

        # Get aliases for the specified user 
$aliases = Get-RemoteMailbox -Identity $usertomanage | Select-Object -ExpandProperty EmailAddresses | Where-Object { $_.PrefixString -eq 'SMTP' } | ForEach-Object { $_.SmtpAddress }

# Get the primary SMTP address for the specified user
$primaryAlias = Get-RemoteMailbox -Identity $usertomanage | Select-Object -ExpandProperty PrimarySmtpAddress

# Create a form to display aliases
$aliasForm = New-Object System.Windows.Forms.Form
$aliasForm.Text = "Aliases for $userToManage"
$aliasForm.Size = New-Object System.Drawing.Size(400, 300)

# Create a label to explain the asterisk
$asteriskLabel = New-Object System.Windows.Forms.Label
$asteriskLabel.Location = New-Object System.Drawing.Point(10, 0)
$asteriskLabel.Size = New-Object System.Drawing.Size(350, 20)
$asteriskLabel.Text = "Asterisk to the right of the address = Primary Alias"
$aliasForm.Controls.Add($asteriskLabel)

# Create a list box to display aliases
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10, 20)
$listBox.Size = New-Object System.Drawing.Size(350, 200)

# Add items to the list box and mark the primary alias with an asterisk to the right
foreach ($alias in $aliases) {
    if ($alias -eq $primaryAlias) {
        $listBox.Items.Add("$alias *")
    } else {
        $listBox.Items.Add($alias)
    }
}

$aliasForm.Controls.Add($listBox)

# Function to show a simple input box
function Show-InputBox {
    param (
        [string]$prompt,
        [string]$title,
        [string]$default
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(300, 140)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(260, 20)
    $label.Text = $prompt
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10, 50)
    $textBox.Size = New-Object System.Drawing.Size(260, 20)
    $textBox.Text = $default
    $form.Controls.Add($textBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(10, 80)
    $okButton.Size = New-Object System.Drawing.Size(75, 23)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(95, 80)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelButton)

    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    }

    return $null
}

# Create an "Add" button
$addButton = New-Object System.Windows.Forms.Button
$addButton.Text = "Add"
$addButton.Location = New-Object System.Drawing.Point(10, 230)
$addButton.Size = New-Object System.Drawing.Size(75, 23)
$addButton.Add_Click({
    # Prompt for the new alias
    $newAlias = Show-InputBox -prompt "Enter the new alias:" -title "Add Alias" -default ""

    if ($newAlias -ne $null) {
        # Add the new alias
        Set-RemoteMailbox -Identity $userToManage -EmailAddresses @{Add=$newAlias}

        # Refresh the aliases and update the list box
        $aliases = Get-RemoteMailbox -Identity $userToManage | Select-Object -ExpandProperty EmailAddresses | Where-Object { $_.PrefixString -eq 'SMTP' } | ForEach-Object { $_.SmtpAddress }
        Update-ListBox
    }
})
$aliasForm.Controls.Add($addButton)





# Function to update the list box with the latest aliases
function Update-ListBox {
    $listBox.Items.Clear()

    # Add items to the list box and mark the primary alias with an asterisk to the right
    foreach ($alias in $aliases) {
        if ($alias -eq $primaryAlias) {
            $listBox.Items.Add("$alias *")
        } else {
            $listBox.Items.Add($alias)
        }
    }
}

$aliasForm.Controls.Add($addButton)

# Create a "Remove" button
$removeButton = New-Object System.Windows.Forms.Button
$removeButton.Text = "Remove"
$removeButton.Location = New-Object System.Drawing.Point(90, 230)
$removeButton.Size = New-Object System.Drawing.Size(75, 23)
$removeButton.Add_Click({
    # Prompt for confirmation
    $confirmation = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to remove the selected alias?", "Remove Alias", "YesNo", "Warning")

    if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Remove the selected alias
        $selectedAlias = $listBox.SelectedItem -replace ' \*$', ''  # Remove the asterisk if present
        Set-RemoteMailbox -Identity $userToManage -EmailAddresses @{Remove=$selectedAlias}

        # Refresh the aliases and update the list box
        $aliases = Get-RemoteMailbox -Identity $userToManage | Select-Object -ExpandProperty EmailAddresses | Where-Object { $_.PrefixString -eq 'SMTP' } | ForEach-Object { $_.SmtpAddress }
        Update-ListBox
    }
})
$aliasForm.Controls.Add($removeButton)

# Create a "Set Primary Alias" button
$setPrimaryButton = New-Object System.Windows.Forms.Button
$setPrimaryButton.Text = "Set Primary Alias"
$setPrimaryButton.Location = New-Object System.Drawing.Point(170, 230)
$setPrimaryButton.Size = New-Object System.Drawing.Size(120, 23)
$setPrimaryButton.Add_Click({
    # Prompt for confirmation
    $confirmation = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to set the selected alias as the primary?", "Set Primary Alias", "YesNo", "Warning")

    if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Set the selected alias as the primary
        $selectedAlias = $listBox.SelectedItem -replace ' \*$', ''  # Remove the asterisk if present
        Set-RemoteMailbox -Identity $userToManage -PrimarySmtpAddress "$selectedAlias"

        # Refresh the aliases and update the list box
        $aliases = Get-RemoteMailbox -Identity $userToManage | Select-Object -ExpandProperty EmailAddresses | Where-Object { $_.PrefixString -eq 'SMTP' } | ForEach-Object { $_.SmtpAddress }
        Update-ListBox
    }
})
$aliasForm.Controls.Add($setPrimaryButton)

# Create an "Exit" button
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Text = "Exit"
$exitButton.Location = New-Object System.Drawing.Point(300, 230)
$exitButton.Size = New-Object System.Drawing.Size(75, 23)
$exitButton.Add_Click({
    $aliasForm.Close()
    return  # Terminate the entire function when the "Exit" button is clicked
})
$aliasForm.Controls.Add($exitButton)

# Show the alias form
$aliasForm.ShowDialog()

# The function will terminate here if the "Exit" button is clicked
return  # Terminate the entire function if not terminated by the "Exit" button

# Continue with displaying aliases or any other action after alias GUI
# Function to update the list box with the latest aliases
function Update-ListBox {
    $listBox.Items.Clear()

    # Add items to the list box and mark the primary alias with an asterisk to the right
    foreach ($alias in $aliases) {
        if ($alias -eq $primaryAlias) {
            $listBox.Items.Add("$alias *")
        } else {
            $listBox.Items.Add($alias)
        }
    }
}

# Show the alias form
$aliasForm.ShowDialog()

    }
}

function Search-Mailbox {
    # Create a new form for searching
    $searchForm = New-Object System.Windows.Forms.Form
    $searchForm.Text = "Search Mailbox"
    $searchForm.Size = New-Object System.Drawing.Size(300,150)
    $searchForm.StartPosition = "CenterScreen"

    # Create label and textbox for entering mailbox name
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20,20)
    $label.Size = New-Object System.Drawing.Size(100,20)
    $label.Text = "Enter Mailbox:"
    $searchForm.Controls.Add($label)

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Location = New-Object System.Drawing.Point(120,20)
    $textbox.Size = New-Object System.Drawing.Size(150,20)
    $searchForm.Controls.Add($textbox)

    # Create button to initiate search
    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Location = New-Object System.Drawing.Point(20,60)
    $searchButton.Size = New-Object System.Drawing.Size(100,30)
    $searchButton.Text = "Search"
    $searchButton.Add_Click({
        $result = Get-RemoteMailbox -Identity $textbox.Text -ErrorAction SilentlyContinue
        if ($result) {
            Show-MailboxInfo $result
        } else {
            [System.Windows.Forms.MessageBox]::Show("Mailbox not found.", "Search Result")
        }
    })
    $searchForm.Controls.Add($searchButton)

    # Create button to close the search form
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(140,60)
    $closeButton.Size = New-Object System.Drawing.Size(100,30)
    $closeButton.Text = "Close"
    $closeButton.Add_Click({ $searchForm.Close() })
    $searchForm.Controls.Add($closeButton)

    # Show the search form
    $searchForm.ShowDialog()
}

function Show-MailboxInfo($mailbox) {
    # Create a new form to display mailbox info
    $infoForm = New-Object System.Windows.Forms.Form
    $infoForm.Text = "Mailbox Info"
    $infoForm.Size = New-Object System.Drawing.Size(500,300)
    $infoForm.StartPosition = "CenterScreen"

    # Create labels to display basic info
    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(20,20)
    $label1.Size = New-Object System.Drawing.Size(450,20)
    $label1.Text = "DisplayName: $($mailbox.DisplayName)"
    $infoForm.Controls.Add($label1)

    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(20,50)
    $label2.Size = New-Object System.Drawing.Size(450,20)
    $label2.Text = "Email: $($mailbox.PrimarySmtpAddress)"
    $infoForm.Controls.Add($label2)

    $label3 = New-Object System.Windows.Forms.Label
    $label3.Location = New-Object System.Drawing.Point(20,80)
    $label3.Size = New-Object System.Drawing.Size(450,20)
    $label3.Text = "Alias: $($mailbox.Alias)"
    $infoForm.Controls.Add($label3)

    $label4 = New-Object System.Windows.Forms.Label
    $label4.Location = New-Object System.Drawing.Point(20,110)
    $label4.Size = New-Object System.Drawing.Size(450,20)
    $label4.Text = "Aliases:"
    $infoForm.Controls.Add($label4)

    # Display all aliases in a multiline textbox
    $aliasesTextbox = New-Object System.Windows.Forms.TextBox
    $aliasesTextbox.Multiline = $true
    $aliasesTextbox.ScrollBars = "Vertical"
    $aliasesTextbox.Location = New-Object System.Drawing.Point(20,140)
    $aliasesTextbox.Size = New-Object System.Drawing.Size(450,100)
    $aliasesTextbox.Text = $($mailbox.EmailAddresses | Where-Object { $_.PrefixString -eq "smtp" } | ForEach-Object { $_.SmtpAddress }) -join "`r`n"
    $infoForm.Controls.Add($aliasesTextbox)

    # Show the info form
    $infoForm.ShowDialog()
}

# Azure Directory Sync 
function Initiate-DirSync {
    try {
        # Run the Start-ADSyncSyncCycle command
        Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
        [System.Windows.Forms.MessageBox]::Show("Delta sync initiated successfully.", "Initiate DirSync")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error initiating Delta sync: $_", "Initiate DirSync", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

function View-AllMailboxes {
    # Run Get-RemoteMailbox command
    $mailboxes = Get-RemoteMailbox

    # Create a new form to display the results
    $resultsForm = New-Object System.Windows.Forms.Form
    $resultsForm.Text = "View All Mailboxes"
    $resultsForm.Size = New-Object System.Drawing.Size(400,300)
    $resultsForm.StartPosition = "CenterScreen"

    # Create a textbox to display the results
    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Multiline = $true
    $textbox.ScrollBars = "Vertical"
    $textbox.Location = New-Object System.Drawing.Point(20,20)
    $textbox.Size = New-Object System.Drawing.Size(360,200)
    $textbox.Text = $mailboxes | Format-Table -AutoSize | Out-String
    $resultsForm.Controls.Add($textbox)

    # Show the results form
    $resultsForm.ShowDialog()
    }

# Run the GUI
Create-Gui
