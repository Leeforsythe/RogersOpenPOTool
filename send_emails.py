import pandas as pd

# Load your open PO Excel file
po_file_path = 'OPENPO.xlsx'  # Replace with the actual file name of your open POs
open_po_data = pd.read_excel(po_file_path)

# Load your vendor email directory
vendor_email_file_path = 'vendor_emails.xlsx'  # Replace with the actual file name of your vendor email directory
vendor_emails_data = pd.read_excel(vendor_email_file_path)

# Merge open PO data with vendor email directory on VendorName
merged_data = open_po_data.merge(vendor_emails_data, on='VendorName', how='left')

# Print the first few rows to confirm the merge was successful
print(merged_data[['VendorName', 'Email']].head())

# Step 1: Define a basic email template
email_template = """
Dear {vendor_name},

I am writing to follow up on the status of our open purchase orders. Below is the information we have on file:

Purchase Order Number: {po_number}
Quantity on back order: {order_status}

Please let me know if there are any updates or if further action is required on my part.

Thank you,
Lee Forsythe
"""

# Step 2: Prepare email content for each vendor
merged_data['EmailContent'] = merged_data.apply(lambda row: email_template.format(
    vendor_name=row['VendorName'],
    po_number=row.get('SourceIdentifier', 'N/A'),  # Adjust to match your PO number column
    order_status=row.get('OrderStatus', 'N/A')  # Adjust to match your order status column
), axis=1)

# Print out a sample email for testing
print(merged_data[['VendorName', 'Email', 'EmailContent']].head())

import yagmail

# Initialize yagmail for sending emails
sender_email = 'leeforsythe@rogerssupply.com'  # Replace with your Gmail address
app_password = 'cvho cgcn pown nrpy'  # Replace with your app-specific password from Google

# Initialize yagmail SMTP client
yag = yagmail.SMTP(user=sender_email, password=app_password)

# Step 1: Define a basic email template
email_template = """
Dear {vendor_name},

I am writing to follow up on the status of our open purchase orders. Below is the information we have on file:

Purchase Order Number: {po_number}
Quantity on back order: {order_status}

Please let me know if there are any updates or if further action is required on my part.

Thank you,
Lee Forsythe
"""

# Prepare email content for each vendor
merged_data['EmailContent'] = merged_data.apply(lambda row: email_template.format(
    vendor_name=row['VendorName'],
    po_number=row.get('InterfacePONumber'),  # Adjust to match your PO number column
    order_status=row.get('OrderStatus')  # Adjust to match your order status column
), axis=1)

# Send test emails
for index, row in merged_data.iterrows():
    recipient_email = row['Email']
    subject = f"Follow Up on Purchase Order - {row.get('SourceIdentifier', 'N/A')}"
    body = row['EmailContent']

    # Send email
    try:
        yag.send(to=recipient_email, subject=subject, contents=body)
        print(f"Email sent to {recipient_email} regarding PO {row.get('SourceIdentifier', 'N/A')}")
    except Exception as e:
        print(f"Failed to send email to {recipient_email}: {e}")

