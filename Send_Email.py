import win32com.client as win32

# Launch Outlook application
outlook = win32.Dispatch('Outlook.Application')

# Create a new mail item (0 = Mail Item)
mail = outlook.CreateItem(0)

# Set email details
mail.To = 'recipient@example.com'
mail.Subject = 'Test Email from Python'
mail.Body = 'Hello,\n\nThis is an automated email sent from Python using Outlook.\n\nRegards,\nYour Name'

# Send the email
mail.Send()

print("Email sent successfully.")





python -c "import win32com.client; print('pywin32 available')"
