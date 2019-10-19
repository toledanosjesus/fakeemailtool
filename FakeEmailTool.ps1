###############################################################
###############################################################
##
##        FAKE EMAIL TOOL
##
##        Version 1.0
##
##        Original tool: 
##        https://github.com/toledanosjesus/fakeemailtool
##
###############################################################
###############################################################

Add-Type -AssemblyName System.Windows.Forms

# FORM CREATION
$Form = New-Object system.Windows.Forms.Form
$Form.Text = "Fake Emails"
$Icon = New-Object system.drawing.icon ("sources\icon.ICO")
$Form.Icon = $Icon
$Form.AutoScroll = $True
$Form.Width = 650
$Form.Height = 650
$form.AutoSize = $True
$Form.MinimizeBox = $True
$Form.MaximizeBox = $False
$Form.WindowState = "Normal"
    # Maximized, Minimized, Normal
$Form.SizeGripStyle = "Hide"
    # Auto, Hide, Show
$Form.ShowInTaskbar = $False
$Form.Opacity = 1
    # 1.0 is fully opaque; 0.0 is invisible
$Form.StartPosition = "CenterScreen"
    # CenterScreen, Manual, WindowsDefaultLocation, WindowsDefaultBounds, CenterParent
$Font = New-Object System.Drawing.Font("Arial",14,[System.Drawing.FontStyle]::Bold)
$Form.Font = $Font
$Form.BackColor = "white"

# Add Image
$img = [System.Drawing.Image]::Fromfile('sources\Logo.png')
$pictureBox = new-object Windows.Forms.PictureBox
$pictureBox.Width = $img.Size.Width
$pictureBox.Height = $img.Size.Height
$pictureBox.Location = New-object System.Drawing.Size(500,500)
$pictureBox.Image = $img
$form.controls.add($pictureBox)


# INTRO
$Intro = New-Object System.Windows.Forms.Label
$Intro.Text = "Fake Emails Tool"
$Intro.Location = New-Object Drawing.Point 50,15
$Intro.AutoSize = $True
$Form.Controls.Add($Intro)

# MAIL TO TEXTBOX
$mailTo = New-Object Windows.Forms.TextBox
$mailTo.Location = New-Object Drawing.Point 160,60
$mailTo.Size = New-Object Drawing.Point 400,30
$FontmailTo = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
$mailTo.Font = $FontmailTo

# Create TextBox and set text, size and location
$mailTotext = New-Object System.Windows.Forms.Label
$mailTotext.Location = New-Object Drawing.Point 50,60
$mailTotext.Size = New-Object Drawing.Point 110,29
$mailTotext.Text = "Mail to:"
$FontmailTotext = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
$mailTotext.Font = $FontmailTotext

$form.Controls.Add($mailTo)
$form.Controls.Add($mailTotext)

# MAIL FROM DISPLAY NAME TEXTBOX
$mailFromName = New-Object Windows.Forms.TextBox
$mailFromName.Location = New-Object Drawing.Point 160,95
$mailFromName.Size = New-Object Drawing.Point 400,30
$FontmailFromName = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
$mailFromName.Font = $FontmailFromName

# Create TextBox and set text, size and location
$mailFromtextName = New-Object System.Windows.Forms.Label
$mailFromtextName.Location = New-Object Drawing.Point 50,95
$mailFromtextName.Size = New-Object Drawing.Point 130,29
$mailFromtextName.Text = "Display Name:"
$FontmailFromtextName = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
$mailFromtextName.Font = $FontmailFromtextName

$form.Controls.Add($mailFromName)
$form.Controls.Add($mailFromtextName)

# MAIL FROM TEXTBOX
$mailFrom = New-Object Windows.Forms.TextBox
$mailFrom.Location = New-Object Drawing.Point 160,130
$mailFrom.Size = New-Object Drawing.Point 400,30
$FontmailFrom = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
$mailFrom.Font = $FontmailFrom

# Create TextBox and set text, size and location
$mailFromtext = New-Object System.Windows.Forms.Label
$mailFromtext.Location = New-Object Drawing.Point 50,130
$mailFromtext.Size = New-Object Drawing.Point 110,29
$mailFromtext.Text = "Mail from:"
$FontmailFromtext = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
$mailFromtext.Font = $FontmailFromtext

$form.Controls.Add($mailFrom)
$form.Controls.Add($mailFromtext)

# MAIL SUBJECT TEXTBOX
$mailSubject = New-Object Windows.Forms.TextBox
$mailSubject.Location = New-Object Drawing.Point 160,165
$mailSubject.Size = New-Object Drawing.Point 400,30
$FontmailSubject = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
$mailSubject.Font = $FontmailSubject

# Create TextBox and set text, size and location
$mailSubjecttext = New-Object System.Windows.Forms.Label
$mailSubjecttext.Location = New-Object Drawing.Point 50,165
$mailSubjecttext.Size = New-Object Drawing.Point 110,29
$mailSubjecttext.Text = "Subject:"
$FontmailSubjecttext = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
$mailSubjecttext.Font = $FontmailSubjecttext

$form.Controls.Add($mailSubject)
$form.Controls.Add($mailSubjecttext)

# MAIL BODY TEXTBOX
$mailBody = New-Object Windows.Forms.TextBox
$mailBody.Multiline = $True
$mailBody.Scrollbars = "Vertical"
$mailBody.Location = New-Object Drawing.Point 160,210
$mailBody.Size = New-Object Drawing.Size 400,180
$FontmailBody = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
$mailBody.Font = $FontmailBody


# Create TextBox and set text, size and location
$mailBodytext = New-Object System.Windows.Forms.Label
$mailBodytext.Location = New-Object Drawing.Point 50,210
$mailBodytext.Size = New-Object Drawing.Point 110,29
$mailBodytext.Text = "Body:"
$FontmailBodytext = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
$mailBodytext.Font = $FontmailBodytext

$form.Controls.Add($mailBody)
$form.Controls.Add($mailBodytext)

# MAIL ATTACHMENT TEXTBOX
$mailAttach = New-Object Windows.Forms.TextBox
$mailAttach.Location = New-Object Drawing.Point 160,400
$mailAttach.Size = New-Object Drawing.Point 360,30
$FontmailAttach = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
$mailAttach.Font = $FontmailAttach

# Create TextBox and set text, size and location
$mailAttachtext = New-Object System.Windows.Forms.Label
$mailAttachtext.Location = New-Object Drawing.Point 50,400
$mailAttachtext.Size = New-Object Drawing.Point 110,29
$mailAttachtext.Text = "Attachment:"
$FontmailAttachtext = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Regular)
$mailAttachtext.Font = $FontmailAttachtext

$form.Controls.Add($mailAttach)
$form.Controls.Add($mailAttachtext)

# ATTACHMENT BUTTON 
# Create Button and set text and location
$battach = New-Object Windows.Forms.Button
$battach.text = "···"
$battach.Size = New-Object Drawing.Point 30,29
$Fontbattach = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Italic)
$battach.Font = $Fontbattach
$battach.Location = New-Object Drawing.Point 530,400

# INPUT HANDLER CANCEL BUTTON – CLOSE APP
# close button – on click close app
$battach.add_click({
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
$FileBrowser.ShowDialog()
$mailAttach.Text = $FileBrowser.FileName
$mailAttach.Refresh()
})

# ADD CONTROLS TO FORM
$form.controls.add($battach)

# SEND BUTTON 
# Create Button and set text and location
$bsend = New-Object Windows.Forms.Button
$bsend.text = "Send"
$Fontbsend = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Italic)
$bsend.Font = $Fontbsend
$bsend.Location = New-Object Drawing.Point 50,435

$bsend.add_click({
$ClickedFont = [System.Drawing.Font]::('Arial', 11, [System.Drawing.FontStyle]::Bold)
$bsend.ForeColor = [System.Drawing.Color]::Green
$bsend.text = "Sent"
$bsend.Font = $ClickedFont

$PSEmailServer = "Your_email_server_address"     # Please, modify this variable
$SMTPPort = "Your_email_server_Port"             # Please, modify this variable
$SMTPUsername = "Your_email_server_username"     # Please, modify this variable
$Password = "Your_email_server_password"         # Please, modify this variable

$SecureStringPassword = $Password | ConvertTo-SecureString -AsPlainText -Force

$EmailCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SMTPUsername,$SecureStringPassword
$mailFrom = $mailFromName.Text + " <" + $mailFrom.Text + ">"

wait-event -Timeout 5
$bsend.Font = $Fontbsend
$bsend.text = "Send"
$bsend.ForeColor = [System.Drawing.Color]::Black

if(!$mailAttach.Text){
Send-MailMessage -From $MailFrom -To $MailTo.Text -Subject $MailSubject.Text -Body $MailBody.Text -Encoding UTF8 -Port $SMTPPort -Credential $EmailCredential
} else {
Send-MailMessage -From $MailFrom -To $MailTo.Text -Subject $MailSubject.Text -Body $MailBody.Text -Attachments $mailAttach.Text -Encoding UTF8 -Port $SMTPPort -Credential $EmailCredential}
})

# ADD CONTROLS TO FORM
$form.controls.add($bsend)

# CANCEL BUTTON 
# Create Button and set text and location
$bcancel = New-Object Windows.Forms.Button
$bcancel.text = "Cancel"
$Fontbcancel = New-Object System.Drawing.Font("Arial",12,[System.Drawing.FontStyle]::Italic)
$bcancel.Font = $Fontbcancel
$bcancel.Location = New-Object Drawing.Point 150,435

# INPUT HANDLER CANCEL BUTTON – CLOSE APP
# close button – on click close app
$bcancel.add_click({
$Form.Close()
})

# ADD CONTROLS TO FORM
$form.controls.add($bcancel)

$form.ShowDialog() | Out-Null