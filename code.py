import win32com.client as win32
import pandas as pd

# Start the Outlook application
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

# Email content
subject = 'Subject Title'
body = """Dear Sir/Madam,
The promotional campaign for your company has been offered with a 5% discount.
Best regards,"""

# Load the Excel file containing the email list
df = pd.read_excel('C:......../Desktop/dMailList.xlsx')

# Extract the email addresses as a list
recipients = df['Email'].tolist()

# Print the list of recipients (optional)
recipients

# Loop through each recipient and send the email
for recipient in recipients:
    # Create a new mail item for each email
    mail = outlook.CreateItem(0)
    
    # Configure the email
    mail.Subject = subject
    mail.Body = body
    mail.Recipients.Add(recipient)
    
    # Send the email
    mail.Send()
    print(f'Email sent to {recipient}!')

