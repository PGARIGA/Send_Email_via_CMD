# Sending Emails with Office 365 SMTP using PowerShell

## Overview
In this guide, we will demonstrate how to send an email with an attachment using Office 365 SMTP through PowerShell. This script is useful for automating email notifications, sending reports, or alerts from a Windows-based system.

## Prerequisites
Before running the script, ensure that:
- PowerShell is installed on your system.
- You have a valid Office 365 account with SMTP authentication enabled.
- Your network allows SMTP traffic on port `587`.
- The sender account has permissions to send emails.

## PowerShell Script
```powershell
$SMTPServer = "smtp.office365.com"
$SMTPPort = "587"
$SMTPUser = "your-email@yourdomain.com"  # Replace with your email
$SMTPPassword = "YourSecurePassword"  # Replace with your secure password

$From = "your-email@yourdomain.com"
$To = "recipient@example.com"  # Replace with the recipient email
$Bcc = "bccrecipient@example.com"  # Optional BCC recipient
$Subject = "Test Email with Attachment"
$Body = "This is a test email sent using Office 365 SMTP with an attachment."
$Attachment = "C:\Path\To\Your\File.xlsx"  # Replace with the actual file path

# Convert password to a secure string
$SecurePassword = ConvertTo-SecureString $SMTPPassword -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential ($SMTPUser, $SecurePassword)

# Send Email with Attachment
Send-MailMessage -From $From -To $To -Bcc $Bcc -Subject $Subject -Body $Body -Attachments $Attachment -SmtpServer $SMTPServer -Port $SMTPPort -UseSsl -Credential $Credential
```

## Explanation
- **SMTP Server & Port**: `smtp.office365.com` on port `587` (TLS enabled) is used for sending emails through Office 365.
- **Authentication**: Uses Office 365 email credentials.
- **Secure Password Handling**: The password is converted into a secure string to avoid exposing it in plaintext.
- **Email Components**:
  - `From`: Sender's email address.
  - `To`: Recipient's email address.
  - `Bcc`: Blind carbon copy (optional).
  - `Subject`: Subject of the email.
  - `Body`: Message content.
  - `Attachment`: Path to the file being attached.
- **Send-MailMessage**: The PowerShell command used to send the email.

## Security Considerations
- **Avoid Hardcoding Credentials**: Instead, store credentials securely using a credential manager or environment variables.
- **Use App Passwords**: If MFA (Multi-Factor Authentication) is enabled, generate and use an app password instead of your main password.
- **Protect Sensitive Information**: Store passwords in a secure vault and never expose them in plaintext in scripts.

## Troubleshooting
- **Authentication Issues**: Ensure that SMTP authentication is enabled for your Office 365 account.
- **Firewall Blocks**: Ensure outbound traffic on port `587` is allowed.
- **Attachment Path Issues**: Verify that the file exists at the specified path.

## Conclusion
This PowerShell script provides a simple yet effective way to send emails with attachments using Office 365 SMTP. By automating email notifications, you can streamline reporting and alerting processes efficiently.

For more PowerShell automation scripts, check out our [GitHub repository](https://github.com/your-repository).

