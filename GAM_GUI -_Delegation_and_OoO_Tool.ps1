<# 
.NAME
    GAM - Delegation and OoO GUI Tool
#>
Set-ExecutionPolicy RemoteSigned
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
Invoke-WebRequest "https://www.google.com/favicon.ico" -OutFile "C:\GAM\favicon.ico"

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(400,573)
$Form.text                       = "GAM - Delegation and OoO GUI Tool"
$Form.TopMost                    = $false
$objIcon = New-Object system.drawing.icon ("C:\GAM\favicon.ico")
$Form.StartPosition = "CenterScreen"
$Form.BackColor                  = "#000000"
$Form.TopMost                    = $false
$Form.Icon = $objIcon
$Form.minimumSize = New-Object System.Drawing.Size(420,600) 
$Form.maximumSize = New-Object System.Drawing.Size(420,600)

$Delegation                      = New-Object system.Windows.Forms.Panel
$Delegation.height               = 150
$Delegation.width                = 372
$Delegation.location             = New-Object System.Drawing.Point(12,16)
$Delegation.BorderStyle = "FixedSingle"
$Delegation.ForeColor              = "#ffffff"

$OutOfOffice                      = New-Object system.Windows.Forms.Panel
$OutOfOffice.height               = 350
$OutOfOffice.width                = 372
$OutOfOffice.location             = New-Object System.Drawing.Point(12,190)
$OutOfOffice.BorderStyle = "FixedSingle"
$OutOfOffice.ForeColor              = "#ffffff"

$FromAccountBox                        = New-Object system.Windows.Forms.TextBox
$FromAccountBox.multiline              = $false
$FromAccountBox.width                  = 121
$FromAccountBox.height                 = 20
$FromAccountBox.location               = New-Object System.Drawing.Point(9,83)
$FromAccountBox.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ToAccountBox                        = New-Object system.Windows.Forms.TextBox
$ToAccountBox.multiline              = $false
$ToAccountBox.width                  = 121
$ToAccountBox.height                 = 20
$ToAccountBox.location               = New-Object System.Drawing.Point(140,83)
$ToAccountBox.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$DelegateButton                         = New-Object system.Windows.Forms.Button
$DelegateButton.text                    = "DELEGATE"
$DelegateButton.width                   = 90
$DelegateButton.height                  = 37
$DelegateButton.location                = New-Object System.Drawing.Point(278,70)
$DelegateButton.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$RevokeDelegateButton                         = New-Object system.Windows.Forms.Button
$RevokeDelegateButton.text                    = "REVOKE"
$RevokeDelegateButton.width                   = 90
$RevokeDelegateButton.height                  = 37
$RevokeDelegateButton.location                = New-Object System.Drawing.Point(278,30)
$RevokeDelegateButton.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$LabelDelegation                 = New-Object system.Windows.Forms.Label
$LabelDelegation.text            = "Delegate accounts or revoke delegations."
$LabelDelegation.AutoSize        = $true
$LabelDelegation.width           = 25
$LabelDelegation.height          = 10
$LabelDelegation.location        = New-Object System.Drawing.Point(9,36)
$LabelDelegation.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$LabelOoO               = New-Object system.Windows.Forms.Label
$LabelOoO.text            = "Set Out Of Office alerts for accounts."
$LabelOoO.AutoSize        = $true
$LabelOoO.width           = 25
$LabelOoO.height          = 10
$LabelOoO.location        = New-Object System.Drawing.Point(9,36)
$LabelOoO.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$DelegationDynamicLabel                 = New-Object system.Windows.Forms.Label
$DelegationDynamicLabel.text            = "Status: Ready"
$DelegationDynamicLabel.AutoSize        = $true
$DelegationDynamicLabel.width           = 25
$DelegationDynamicLabel.height          = 10
$DelegationDynamicLabel.location        = New-Object System.Drawing.Point(9,115)
$DelegationDynamicLabel.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$FromAccountLabel                 = New-Object system.Windows.Forms.Label
$FromAccountLabel.text            = "From (child) email:"
$FromAccountLabel.AutoSize        = $true
$FromAccountLabel.width           = 25
$FromAccountLabel.height          = 10
$FromAccountLabel.location        = New-Object System.Drawing.Point(9,60)
$FromAccountLabel.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$ToAccountLabel                 = New-Object system.Windows.Forms.Label
$ToAccountLabel.text            = "To (parent) email:"
$ToAccountLabel.AutoSize        = $true
$ToAccountLabel.width           = 25
$ToAccountLabel.height          = 10
$ToAccountLabel.location        = New-Object System.Drawing.Point(140,60)
$ToAccountLabel.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TitleDelegation                 = New-Object system.Windows.Forms.Label
$TitleDelegation.text            = "ACCOUNT DELEGATION"
$TitleDelegation.AutoSize        = $true
$TitleDelegation.width           = 25
$TitleDelegation.height          = 10
$TitleDelegation.location        = New-Object System.Drawing.Point(9,11)
$TitleDelegation.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$TitleOoO                 = New-Object system.Windows.Forms.Label
$TitleOoO.text            = "OUT OF OFFICE"
$TitleOoO.AutoSize        = $true
$TitleOoO.width           = 25
$TitleOoO.height          = 10
$TitleOoO.location        = New-Object System.Drawing.Point(9,16)
$TitleOoO.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

$UserAccountLabel                 = New-Object system.Windows.Forms.Label
$UserAccountLabel.text            = "User email address:"
$UserAccountLabel.AutoSize        = $true
$UserAccountLabel.width           = 25
$UserAccountLabel.height          = 10
$UserAccountLabel.location        = New-Object System.Drawing.Point(9,60)
$UserAccountLabel.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$UserAccountBox                        = New-Object system.Windows.Forms.TextBox
$UserAccountBox.multiline              = $false
$UserAccountBox.width                  = 121
$UserAccountBox.height                 = 20
$UserAccountBox.location               = New-Object System.Drawing.Point(9,83)
$UserAccountBox.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$SubjectLabel                 = New-Object system.Windows.Forms.Label
$SubjectLabel.text            = "OoO email subject:"
$SubjectLabel.AutoSize        = $true
$SubjectLabel.width           = 25
$SubjectLabel.height          = 10
$SubjectLabel.location        = New-Object System.Drawing.Point(140,60)
$SubjectLabel.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$SubjectBox                        = New-Object system.Windows.Forms.TextBox
$SubjectBox.multiline              = $false
$SubjectBox.width                  = 121
$SubjectBox.height                 = 20
$SubjectBox.location               = New-Object System.Drawing.Point(140,83)
$SubjectBox.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$BodyLabel                 = New-Object system.Windows.Forms.Label
$BodyLabel.text            = "OoO email body (add \n to create a new line):"
$BodyLabel.AutoSize        = $true
$BodyLabel.width           = 25
$BodyLabel.height          = 10
$BodyLabel.location        = New-Object System.Drawing.Point(9,120)
$BodyLabel.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$BodyBox                        = New-Object system.Windows.Forms.TextBox
$BodyBox.multiline              = $true
$BodyBox.width                  = 250
$BodyBox.height                 = 160
$BodyBox.location               = New-Object System.Drawing.Point(9,143)
$BodyBox.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$OoODynamicLabel                 = New-Object system.Windows.Forms.Label
$OoODynamicLabel.text            = "Status: Ready"
$OoODynamicLabel.AutoSize        = $true
$OoODynamicLabel.width           = 25
$OoODynamicLabel.height          = 10
$OoODynamicLabel.location        = New-Object System.Drawing.Point(9,310)
$OoODynamicLabel.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$OoOButton                         = New-Object system.Windows.Forms.Button
$OoOButton.text                    = "OoO ON"
$OoOButton.width                   = 90
$OoOButton.height                  = 37
$OoOButton.location                = New-Object System.Drawing.Point(278,268)
$OoOButton.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$OoOffButton                         = New-Object system.Windows.Forms.Button
$OoOffButton.text                    = "OoO OFF"
$OoOffButton.width                   = 90
$OoOffButton.height                  = 37
$OoOffButton.location                = New-Object System.Drawing.Point(278,228)
$OoOffButton.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Form.controls.AddRange(@($Delegation,$OutOfOffice))
$Delegation.controls.AddRange(@($FromAccountBox, $RevokeDelegateButton, $FromAccountLabel, $ToAccountBox, $DelegationDynamicLabel, $ToAccountLabel, $DelegateButton, $LabelDelegation,$TitleDelegation))
$OutOfOffice.controls.AddRange(@($TitleOoO, $LabelOoO, $OoOffButton, $OoOButton, $UserAccountLabel, $OoODynamicLabel, $UserAccountBox, $SubjectLabel,$SubjectBox, $BodyLabel, $BodyBox))

$gamlocation = "C:\GAM"
if (!(Test-Path $gamlocation)) {
  $DelegationDynamicLabel.Text =  "GAM instance not found. Please install GAM first."
  $DelegationDynamicLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#FF0000")
  $DelegationDynamicLabel.Refresh()
  $OoODynamicLabel.Text =  "GAM instance not found. Please install GAM first."
  $OoODynamicLabel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#FF0000")
  $OoODynamicLabel.Refresh()
}

#region Logic 
$DelegateButton.Add_Click( {delegating})
$RevokeDelegateButton.Add_Click( {revoking})
$OoOButton.Add_Click( {OOO})
$OoOffButton.Add_Click( {OOOff})
function delegating(){
$FromAccount = $FromAccountBox.Text
$ToAccount = $ToAccountBox.Text
$DelegationDynamicLabel.Text =  "Delegating $FromAccount to $ToAccount..."
$DelegationDynamicLabel.Refresh()
sleep(1)
$DelegationOutPut = (cmd.exe /c gam user $FromAccount delegate to $ToAccount) | Out-String
$DelegationDynamicLabel.Text =  "$DelegationOutPut"
$DelegationDynamicLabel.Refresh()
sleep(2)
$DelegationDynamicLabel.Text =  "Finished."
$DelegationDynamicLabel.Refresh()
}
function revoking(){
$FromAccount = $FromAccountBox.Text
$ToAccount = $ToAccountBox.Text
$DelegationDynamicLabel.Text =  "Removing access from $ToAccount to $FromAccount ..."
$DelegationDynamicLabel.Refresh()
sleep(1)
$RevokeOutPut = (cmd.exe /c gam user $FromAccount delete delegate $ToAccount) | Out-String
$DelegationDynamicLabel.Text =  "$RevokeOutPut"
$DelegationDynamicLabel.Refresh()
sleep(2)
$DelegationDynamicLabel.Text =  "Finished."
$DelegationDynamicLabel.Refresh()
}
function OOO(){
$User = $UserAccountBox.Text
$OoOSubject = $SubjectBox.Text
$OoOBody = $BodyBox.Text
$OoODynamicLabel.Text =  "Setting Out of Office alert for $User..."
$OoODynamicLabel.Refresh()
sleep(1)
$OoOOutPut = (cmd.exe /c gam user $User vacation on subject $OoOSubject message $OoOBody startdate 2022-01-01 enddate 2999-04-04) | Out-String
$OoODynamicLabel.Text =  "$OoOOutPut"
$OoODynamicLabel.Refresh()
sleep(1)
$OoODynamicLabel.Text =  "Finished."
$OoODynamicLabel.Refresh()
}
function OOOff(){
$User = $UserAccountBox.Text
$OoODynamicLabel.Text =  "Disabling Out of Office alert for $User..."
$OoODynamicLabel.Refresh()
sleep(1)
$OoOOutPut = (cmd.exe /c gam user $User vacation off) | Out-String
$OoODynamicLabel.Text =  "$OoOOutPut"
$OoODynamicLabel.Refresh()
sleep(1)
$OoODynamicLabel.Text =  "Finished."
$OoODynamicLabel.Refresh()
}
#endregion

[void]$Form.ShowDialog()