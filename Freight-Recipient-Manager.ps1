<#
.SYNOPSIS
    Freight Company Recipient Manager
    
.DESCRIPTION
    Manages contact information for freight companies, including email addresses,
    phone numbers, categories (favorites, regular, specific routes), and notes.
    Integrates with email generation for quick recipient selection.
    
.EXAMPLE
    .\Freight-Recipient-Manager.ps1
    
.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Create Date: 2025-01-07
    Version: 1.0.0
    Change Date: 
    Change Purpose: Initial Release - Recipient Management System
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:RecipientsFile = ".\Data\FreightRecipients.json"
$script:Recipients = @()

#region Data Management Functions
function Initialize-RecipientData {
    <#
    .SYNOPSIS
        Initializes recipient data storage
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    
    $dataDir = Split-Path $script:RecipientsFile -Parent
    if (!(Test-Path $dataDir)) {
        New-Item -ItemType Directory -Path $dataDir -Force | Out-Null
    }
    
    if (Test-Path $script:RecipientsFile) {
        $script:Recipients = Get-Content $script:RecipientsFile -Raw | ConvertFrom-Json
        Write-Host "✅ Loaded $($script:Recipients.Count) freight companies" -ForegroundColor Green
    }
    else {
        # Create with sample data
        $script:Recipients = @(
            @{
                Id = (New-Guid).ToString()
                CompanyName = "Maddy Freight Services"
                ContactName = "Maddy Clark"
                Email = "maddy@maddyfreight.com"
                Phone = "(555) 123-4567"
                Category = "Favorite"
                Specialties = @("RV Parts", "Palletized Freight")
                Routes = @("Michigan to Ohio", "Indiana to Ohio")
                Notes = "Preferred carrier - excellent service"
                LastUsed = (Get-Date).ToString("yyyy-MM-dd")
                TimesUsed = 15
            }
        )
        Save-RecipientData
    }
}

function Save-RecipientData {
    <#
    .SYNOPSIS
        Saves recipient data to JSON file
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    
    $script:Recipients | ConvertTo-Json -Depth 10 | Out-File $script:RecipientsFile -Encoding UTF8
    Write-Host "💾 Recipient data saved" -ForegroundColor Green
}

function Add-Recipient {
    <#
    .SYNOPSIS
        Adds a new freight company recipient
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    param (
        [hashtable]$RecipientData
    )
    
    $RecipientData.Id = (New-Guid).ToString()
    $RecipientData.LastUsed = (Get-Date).ToString("yyyy-MM-dd")
    $RecipientData.TimesUsed = 0
    
    $script:Recipients += $RecipientData
    Save-RecipientData
    
    return $RecipientData.Id
}

function Update-Recipient {
    <#
    .SYNOPSIS
        Updates existing recipient information
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    param (
        [string]$Id,
        [hashtable]$UpdatedData
    )
    
    for ($i = 0; $i -lt $script:Recipients.Count; $i++) {
        if ($script:Recipients[$i].Id -eq $Id) {
            foreach ($key in $UpdatedData.Keys) {
                $script:Recipients[$i].$key = $UpdatedData[$key]
            }
            Save-RecipientData
            return $true
        }
    }
    return $false
}

function Remove-Recipient {
    <#
    .SYNOPSIS
        Removes a recipient from the list
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    param (
        [string]$Id
    )
    
    $script:Recipients = $script:Recipients | Where-Object { $_.Id -ne $Id }
    Save-RecipientData
}

function Get-FilteredRecipients {
    <#
    .SYNOPSIS
        Gets filtered list of recipients
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    param (
        [string]$Category = "All",
        [string]$SearchText = ""
    )
    
    $filtered = $script:Recipients
    
    if ($Category -ne "All") {
        $filtered = $filtered | Where-Object { $_.Category -eq $Category }
    }
    
    if ($SearchText) {
        $filtered = $filtered | Where-Object {
            $_.CompanyName -like "*$SearchText*" -or
            $_.ContactName -like "*$SearchText*" -or
            $_.Email -like "*$SearchText*"
        }
    }
    
    return $filtered
}

function Export-RecipientList {
    <#
    .SYNOPSIS
        Exports recipient list to CSV
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    param (
        [string]$OutputPath
    )
    
    $exportData = $script:Recipients | ForEach-Object {
        [PSCustomObject]@{
            CompanyName = $_.CompanyName
            ContactName = $_.ContactName
            Email = $_.Email
            Phone = $_.Phone
            Category = $_.Category
            Specialties = ($_.Specialties -join "; ")
            Routes = ($_.Routes -join "; ")
            Notes = $_.Notes
            TimesUsed = $_.TimesUsed
            LastUsed = $_.LastUsed
        }
    }
    
    $exportData | Export-Csv -Path $OutputPath -NoTypeInformation
    Write-Host "✅ Exported to: $OutputPath" -ForegroundColor Green
}
#endregion

#region GUI Functions
function New-RecipientManagerGUI {
    <#
    .SYNOPSIS
        Creates the Recipient Manager GUI
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    
    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Freight Company Recipient Manager"
    $form.Size = New-Object System.Drawing.Size(1200, 700)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    #region Top Controls
    # Category Filter
    $lblCategory = New-Object System.Windows.Forms.Label
    $lblCategory.Location = New-Object System.Drawing.Point(20, 20)
    $lblCategory.Size = New-Object System.Drawing.Size(60, 20)
    $lblCategory.Text = "Category:"
    $form.Controls.Add($lblCategory)
    
    $cmbCategory = New-Object System.Windows.Forms.ComboBox
    $cmbCategory.Location = New-Object System.Drawing.Point(85, 18)
    $cmbCategory.Size = New-Object System.Drawing.Size(150, 25)
    $cmbCategory.DropDownStyle = "DropDownList"
    @("All", "Favorite", "Regular", "Occasional", "Route-Specific") | ForEach-Object {
        $cmbCategory.Items.Add($_) | Out-Null
    }
    $cmbCategory.SelectedIndex = 0
    $form.Controls.Add($cmbCategory)
    
    # Search Box
    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Location = New-Object System.Drawing.Point(250, 20)
    $lblSearch.Size = New-Object System.Drawing.Size(50, 20)
    $lblSearch.Text = "Search:"
    $form.Controls.Add($lblSearch)
    
    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(305, 18)
    $txtSearch.Size = New-Object System.Drawing.Size(250, 25)
    $form.Controls.Add($txtSearch)
    
    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Location = New-Object System.Drawing.Point(560, 17)
    $btnSearch.Size = New-Object System.Drawing.Size(80, 27)
    $btnSearch.Text = "🔍 Search"
    $form.Controls.Add($btnSearch)
    
    # Action Buttons
    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Location = New-Object System.Drawing.Point(900, 17)
    $btnAdd.Size = New-Object System.Drawing.Size(120, 30)
    $btnAdd.Text = "➕ Add New"
    $btnAdd.BackColor = [System.Drawing.Color]::LightGreen
    $btnAdd.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btnAdd)
    
    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Location = New-Object System.Drawing.Point(1030, 17)
    $btnExport.Size = New-Object System.Drawing.Size(140, 30)
    $btnExport.Text = "📤 Export to CSV"
    $form.Controls.Add($btnExport)
    #endregion
    
    #region DataGridView
    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Location = New-Object System.Drawing.Point(20, 60)
    $dgv.Size = New-Object System.Drawing.Size(1150, 500)
    $dgv.AllowUserToAddRows = $false
    $dgv.AllowUserToDeleteRows = $false
    $dgv.ReadOnly = $true
    $dgv.SelectionMode = "FullRowSelect"
    $dgv.MultiSelect = $false
    $dgv.AutoSizeColumnsMode = "Fill"
    
    # Define columns
    $dgv.Columns.Add("CompanyName", "Company Name") | Out-Null
    $dgv.Columns.Add("ContactName", "Contact Name") | Out-Null
    $dgv.Columns.Add("Email", "Email") | Out-Null
    $dgv.Columns.Add("Phone", "Phone") | Out-Null
    $dgv.Columns.Add("Category", "Category") | Out-Null
    $dgv.Columns.Add("TimesUsed", "Times Used") | Out-Null
    $dgv.Columns.Add("LastUsed", "Last Used") | Out-Null
    
    $dgv.Columns["TimesUsed"].Width = 90
    $dgv.Columns["LastUsed"].Width = 100
    $dgv.Columns["Category"].Width = 110
    
    $form.Controls.Add($dgv)
    #endregion
    
    #region Bottom Buttons
    $btnEdit = New-Object System.Windows.Forms.Button
    $btnEdit.Location = New-Object System.Drawing.Point(20, 575)
    $btnEdit.Size = New-Object System.Drawing.Size(120, 35)
    $btnEdit.Text = "✏️ Edit Selected"
    $btnEdit.BackColor = [System.Drawing.Color]::LightBlue
    $form.Controls.Add($btnEdit)
    
    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Location = New-Object System.Drawing.Point(150, 575)
    $btnDelete.Size = New-Object System.Drawing.Size(120, 35)
    $btnDelete.Text = "🗑️ Delete"
    $btnDelete.BackColor = [System.Drawing.Color]::LightCoral
    $form.Controls.Add($btnDelete)
    
    $btnCopyEmails = New-Object System.Windows.Forms.Button
    $btnCopyEmails.Location = New-Object System.Drawing.Point(280, 575)
    $btnCopyEmails.Size = New-Object System.Drawing.Size(180, 35)
    $btnCopyEmails.Text = "📋 Copy Selected Emails"
    $form.Controls.Add($btnCopyEmails)
    
    $btnStats = New-Object System.Windows.Forms.Button
    $btnStats.Location = New-Object System.Drawing.Point(900, 575)
    $btnStats.Size = New-Object System.Drawing.Size(140, 35)
    $btnStats.Text = "📊 View Statistics"
    $form.Controls.Add($btnStats)
    
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Location = New-Object System.Drawing.Point(1050, 575)
    $btnClose.Size = New-Object System.Drawing.Size(120, 35)
    $btnClose.Text = "✖️ Close"
    $form.Controls.Add($btnClose)
    #endregion
    
    #region Helper Functions
    function Refresh-RecipientGrid {
        param (
            [string]$Category = "All",
            [string]$SearchText = ""
        )
        
        $dgv.Rows.Clear()
        $filtered = Get-FilteredRecipients -Category $Category -SearchText $SearchText
        
        foreach ($recipient in $filtered) {
            $row = @(
                $recipient.CompanyName,
                $recipient.ContactName,
                $recipient.Email,
                $recipient.Phone,
                $recipient.Category,
                $recipient.TimesUsed,
                $recipient.LastUsed
            )
            $dgv.Rows.Add($row) | Out-Null
            $dgv.Rows[$dgv.Rows.Count - 1].Tag = $recipient.Id
        }
        
        # Update status
        $form.Text = "Freight Company Recipient Manager - $($filtered.Count) companies shown"
    }
    
    function Show-RecipientDialog {
        param (
            [string]$Mode = "Add",
            [object]$ExistingRecipient = $null
        )
        
        $dlg = New-Object System.Windows.Forms.Form
        $dlg.Text = if ($Mode -eq "Add") { "Add New Freight Company" } else { "Edit Freight Company" }
        $dlg.Size = New-Object System.Drawing.Size(550, 600)
        $dlg.StartPosition = "CenterParent"
        $dlg.FormBorderStyle = "FixedDialog"
        $dlg.MaximizeBox = $false
        $dlg.MinimizeBox = $false
        
        $yPos = 20
        
        # Company Name
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(20, $yPos)
        $lbl.Size = New-Object System.Drawing.Size(120, 20)
        $lbl.Text = "Company Name: *"
        $dlg.Controls.Add($lbl)
        
        $txtCompany = New-Object System.Windows.Forms.TextBox
        $txtCompany.Location = New-Object System.Drawing.Point(150, $yPos)
        $txtCompany.Size = New-Object System.Drawing.Size(350, 25)
        if ($ExistingRecipient) { $txtCompany.Text = $ExistingRecipient.CompanyName }
        $dlg.Controls.Add($txtCompany)
        
        $yPos += 40
        
        # Contact Name
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(20, $yPos)
        $lbl.Size = New-Object System.Drawing.Size(120, 20)
        $lbl.Text = "Contact Name:"
        $dlg.Controls.Add($lbl)
        
        $txtContact = New-Object System.Windows.Forms.TextBox
        $txtContact.Location = New-Object System.Drawing.Point(150, $yPos)
        $txtContact.Size = New-Object System.Drawing.Size(350, 25)
        if ($ExistingRecipient) { $txtContact.Text = $ExistingRecipient.ContactName }
        $dlg.Controls.Add($txtContact)
        
        $yPos += 40
        
        # Email
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(20, $yPos)
        $lbl.Size = New-Object System.Drawing.Size(120, 20)
        $lbl.Text = "Email: *"
        $dlg.Controls.Add($lbl)
        
        $txtEmail = New-Object System.Windows.Forms.TextBox
        $txtEmail.Location = New-Object System.Drawing.Point(150, $yPos)
        $txtEmail.Size = New-Object System.Drawing.Size(350, 25)
        if ($ExistingRecipient) { $txtEmail.Text = $ExistingRecipient.Email }
        $dlg.Controls.Add($txtEmail)
        
        $yPos += 40
        
        # Phone
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(20, $yPos)
        $lbl.Size = New-Object System.Drawing.Size(120, 20)
        $lbl.Text = "Phone:"
        $dlg.Controls.Add($lbl)
        
        $txtPhone = New-Object System.Windows.Forms.TextBox
        $txtPhone.Location = New-Object System.Drawing.Point(150, $yPos)
        $txtPhone.Size = New-Object System.Drawing.Size(200, 25)
        if ($ExistingRecipient) { $txtPhone.Text = $ExistingRecipient.Phone }
        $dlg.Controls.Add($txtPhone)
        
        $yPos += 40
        
        # Category
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(20, $yPos)
        $lbl.Size = New-Object System.Drawing.Size(120, 20)
        $lbl.Text = "Category:"
        $dlg.Controls.Add($lbl)
        
        $cmbCat = New-Object System.Windows.Forms.ComboBox
        $cmbCat.Location = New-Object System.Drawing.Point(150, $yPos)
        $cmbCat.Size = New-Object System.Drawing.Size(200, 25)
        $cmbCat.DropDownStyle = "DropDownList"
        @("Favorite", "Regular", "Occasional", "Route-Specific") | ForEach-Object {
            $cmbCat.Items.Add($_) | Out-Null
        }
        $cmbCat.SelectedIndex = 1
        if ($ExistingRecipient) { $cmbCat.Text = $ExistingRecipient.Category }
        $dlg.Controls.Add($cmbCat)
        
        $yPos += 40
        
        # Specialties
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(20, $yPos)
        $lbl.Size = New-Object System.Drawing.Size(120, 20)
        $lbl.Text = "Specialties:"
        $dlg.Controls.Add($lbl)
        
        $txtSpecialties = New-Object System.Windows.Forms.TextBox
        $txtSpecialties.Location = New-Object System.Drawing.Point(150, $yPos)
        $txtSpecialties.Size = New-Object System.Drawing.Size(350, 25)
        if ($ExistingRecipient -and $ExistingRecipient.Specialties) {
            $txtSpecialties.Text = $ExistingRecipient.Specialties -join ", "
        }
        $dlg.Controls.Add($txtSpecialties)
        
        $lblHint = New-Object System.Windows.Forms.Label
        $lblHint.Location = New-Object System.Drawing.Point(150, ($yPos + 27))
        $lblHint.Size = New-Object System.Drawing.Size(350, 15)
        $lblHint.Text = "(Comma-separated: RV Parts, Heavy Equipment, etc.)"
        $lblHint.ForeColor = [System.Drawing.Color]::Gray
        $lblHint.Font = New-Object System.Drawing.Font("Segoe UI", 7)
        $dlg.Controls.Add($lblHint)
        
        $yPos += 60
        
        # Routes
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(20, $yPos)
        $lbl.Size = New-Object System.Drawing.Size(120, 20)
        $lbl.Text = "Routes:"
        $dlg.Controls.Add($lbl)
        
        $txtRoutes = New-Object System.Windows.Forms.TextBox
        $txtRoutes.Location = New-Object System.Drawing.Point(150, $yPos)
        $txtRoutes.Size = New-Object System.Drawing.Size(350, 25)
        if ($ExistingRecipient -and $ExistingRecipient.Routes) {
            $txtRoutes.Text = $ExistingRecipient.Routes -join ", "
        }
        $dlg.Controls.Add($txtRoutes)
        
        $lblHint2 = New-Object System.Windows.Forms.Label
        $lblHint2.Location = New-Object System.Drawing.Point(150, ($yPos + 27))
        $lblHint2.Size = New-Object System.Drawing.Size(350, 15)
        $lblHint2.Text = "(Comma-separated: MI to OH, IN to OH, etc.)"
        $lblHint2.ForeColor = [System.Drawing.Color]::Gray
        $lblHint2.Font = New-Object System.Drawing.Font("Segoe UI", 7)
        $dlg.Controls.Add($lblHint2)
        
        $yPos += 60
        
        # Notes
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(20, $yPos)
        $lbl.Size = New-Object System.Drawing.Size(120, 20)
        $lbl.Text = "Notes:"
        $dlg.Controls.Add($lbl)
        
        $txtNotes = New-Object System.Windows.Forms.TextBox
        $txtNotes.Location = New-Object System.Drawing.Point(150, $yPos)
        $txtNotes.Size = New-Object System.Drawing.Size(350, 80)
        $txtNotes.Multiline = $true
        $txtNotes.ScrollBars = "Vertical"
        if ($ExistingRecipient) { $txtNotes.Text = $ExistingRecipient.Notes }
        $dlg.Controls.Add($txtNotes)
        
        $yPos += 100
        
        # Buttons
        $btnSave = New-Object System.Windows.Forms.Button
        $btnSave.Location = New-Object System.Drawing.Point(300, $yPos)
        $btnSave.Size = New-Object System.Drawing.Size(100, 35)
        $btnSave.Text = "💾 Save"
        $btnSave.BackColor = [System.Drawing.Color]::LightGreen
        $btnSave.DialogResult = "OK"
        $dlg.Controls.Add($btnSave)
        
        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Location = New-Object System.Drawing.Point(410, $yPos)
        $btnCancel.Size = New-Object System.Drawing.Size(100, 35)
        $btnCancel.Text = "❌ Cancel"
        $btnCancel.DialogResult = "Cancel"
        $dlg.Controls.Add($btnCancel)
        
        $dlg.AcceptButton = $btnSave
        $dlg.CancelButton = $btnCancel
        
        $result = $dlg.ShowDialog()
        
        if ($result -eq "OK") {
            if (!$txtCompany.Text -or !$txtEmail.Text) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Company Name and Email are required!",
                    "Validation Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                return $null
            }
            
            $recipientData = @{
                CompanyName = $txtCompany.Text
                ContactName = $txtContact.Text
                Email = $txtEmail.Text
                Phone = $txtPhone.Text
                Category = $cmbCat.Text
                Specialties = @()
                Routes = @()
                Notes = $txtNotes.Text
            }
            
            if ($txtSpecialties.Text) {
                $recipientData.Specialties = $txtSpecialties.Text -split ',' | ForEach-Object { $_.Trim() }
            }
            
            if ($txtRoutes.Text) {
                $recipientData.Routes = $txtRoutes.Text -split ',' | ForEach-Object { $_.Trim() }
            }
            
            return $recipientData
        }
        
        return $null
    }
    #endregion
    
    #region Event Handlers
    
    # Initial load
    Refresh-RecipientGrid
    
    # Category filter change
    $cmbCategory.Add_SelectedIndexChanged({
        Refresh-RecipientGrid -Category $cmbCategory.Text -SearchText $txtSearch.Text
    })
    
    # Search button
    $btnSearch.Add_Click({
        Refresh-RecipientGrid -Category $cmbCategory.Text -SearchText $txtSearch.Text
    })
    
    # Search on Enter key
    $txtSearch.Add_KeyDown({
        if ($_.KeyCode -eq "Enter") {
            Refresh-RecipientGrid -Category $cmbCategory.Text -SearchText $txtSearch.Text
        }
    })
    
    # Add button
    $btnAdd.Add_Click({
        $newData = Show-RecipientDialog -Mode "Add"
        if ($newData) {
            Add-Recipient -RecipientData $newData
            Refresh-RecipientGrid -Category $cmbCategory.Text -SearchText $txtSearch.Text
            
            [System.Windows.Forms.MessageBox]::Show(
                "Freight company added successfully!",
                "Success",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    })
    
    # Edit button
    $btnEdit.Add_Click({
        if ($dgv.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please select a company to edit.",
                "No Selection",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }
        
        $selectedId = $dgv.SelectedRows[0].Tag
        $existing = $script:Recipients | Where-Object { $_.Id -eq $selectedId }
        
        if ($existing) {
            $updatedData = Show-RecipientDialog -Mode "Edit" -ExistingRecipient $existing
            if ($updatedData) {
                Update-Recipient -Id $selectedId -UpdatedData $updatedData
                Refresh-RecipientGrid -Category $cmbCategory.Text -SearchText $txtSearch.Text
                
                [System.Windows.Forms.MessageBox]::Show(
                    "Company information updated successfully!",
                    "Success",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
        }
    })
    
    # Delete button
    $btnDelete.Add_Click({
        if ($dgv.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please select a company to delete.",
                "No Selection",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }
        
        $companyName = $dgv.SelectedRows[0].Cells[0].Value
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to delete:`n`n$companyName`n`nThis action cannot be undone.",
            "Confirm Delete",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($result -eq "Yes") {
            $selectedId = $dgv.SelectedRows[0].Tag
            Remove-Recipient -Id $selectedId
            Refresh-RecipientGrid -Category $cmbCategory.Text -SearchText $txtSearch.Text
            
            [System.Windows.Forms.MessageBox]::Show(
                "Company deleted successfully!",
                "Deleted",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    })
    
    # Copy emails button
    $btnCopyEmails.Add_Click({
        if ($dgv.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please select companies to copy emails from.",
                "No Selection",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }
        
        $emails = @()
        foreach ($row in $dgv.SelectedRows) {
            $email = $row.Cells[2].Value
            if ($email) { $emails += $email }
        }
        
        if ($emails.Count -gt 0) {
            $emailString = $emails -join "; "
            [System.Windows.Forms.Clipboard]::SetText($emailString)
            
            [System.Windows.Forms.MessageBox]::Show(
                "Copied $($emails.Count) email(s) to clipboard:`n`n$emailString",
                "Copied",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    })
    
    # Export button
    $btnExport.Add_Click({
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV files (*.csv)|*.csv"
        $saveDialog.FileName = "FreightCompanies_$(Get-Date -Format 'yyyyMMdd').csv"
        
        if ($saveDialog.ShowDialog() -eq "OK") {
            Export-RecipientList -OutputPath $saveDialog.FileName
            
            [System.Windows.Forms.MessageBox]::Show(
                "Recipient list exported successfully!`n`nFile: $($saveDialog.FileName)",
                "Export Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    })
    
    # Stats button
    $btnStats.Add_Click({
        $totalCompanies = $script:Recipients.Count
        $favorites = ($script:Recipients | Where-Object { $_.Category -eq "Favorite" }).Count
        $totalUsage = ($script:Recipients.TimesUsed | Measure-Object -Sum).Sum
        $mostUsed = $script:Recipients | Sort-Object TimesUsed -Descending | Select-Object -First 1
        
        $statsMsg = @"
📊 RECIPIENT STATISTICS

Total Companies: $totalCompanies
Favorite Companies: $favorites
Total Quotes Sent: $totalUsage

Most Used Company:
  $($mostUsed.CompanyName)
  Times Used: $($mostUsed.TimesUsed)
  Last Used: $($mostUsed.LastUsed)

Categories Breakdown:
$(($script:Recipients | Group-Object Category | ForEach-Object { "  $($_.Name): $($_.Count)" }) -join "`n")
"@
        
        [System.Windows.Forms.MessageBox]::Show(
            $statsMsg,
            "Recipient Statistics",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    })
    
    # Close button
    $btnClose.Add_Click({
        $form.Close()
    })
    
    # Double-click to edit
    $dgv.Add_CellDoubleClick({
        $btnEdit.PerformClick()
    })
    
    #endregion
    
    $form.ShowDialog() | Out-Null
}
#endregion

#region Main Execution
Initialize-RecipientData
New-RecipientManagerGUI
#endregion
