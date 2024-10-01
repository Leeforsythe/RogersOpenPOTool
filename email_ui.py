import streamlit as st
import pandas as pd
import yagmail

# Initialize yagmail for sending emails
sender_email = 'leeforsythe@rogerssupply.com'  # Replace with your Gmail address
app_password = 'cvho cgcn pown nrpy'  # Replace with your app-specific password from Google

# Initialize yagmail SMTP client
yag = yagmail.SMTP(user=sender_email, password=app_password)

# Email template
email_template = """
Dear {vendor_name},

I am writing to follow up on the status of our open purchase orders. Below is the information we have on file:

Purchase Order Number: {po_number}
Quantity on back order: {order_status}

Please let me know if there are any updates or if further action is required on my part.

Thank you,
Lee Forsythe
"""

# Vendor email mapping (predefined as before)
vendor_emails_data = pd.read_excel('vendor_emails.xlsx')

# Streamlit UI components
st.title('Open Purchase Order Email Automation')

# File upload for the Open PO Excel file
uploaded_file = st.file_uploader("Upload the Open PO Excel file", type=["xlsx"])

if uploaded_file:
    # Read the uploaded Excel file
    open_po_data = pd.read_excel(uploaded_file)
    
    # Merge with vendor emails
    merged_data = open_po_data.merge(vendor_emails_data, on='VendorName', how='left')

    # Generate email content
    merged_data['EmailContent'] = merged_data.apply(lambda row: email_template.format(
        vendor_name=row['VendorName'],
        po_number=row.get('SourceIdentifier', 'N/A'),  # Adjust if your column name changes
        order_status=row.get('OrderStatus', 'N/A')  # Adjust if your column name changes
    ), axis=1)
    
    # Display the email content for preview
    st.write("Here is a preview of the emails:")
    st.write(merged_data[['VendorName', 'Email', 'EmailContent']].head())

    # Button to send emails
    if st.button('Send Emails'):
        for index, row in merged_data.iterrows():
            recipient_email = row['Email']
            subject = f"Follow Up on Purchase Order - {row.get('SourceIdentifier', 'N/A')}"
            body = row['EmailContent']
            
            # Send email
            try:
                yag.send(to=recipient_email, subject=subject, contents=body)
                st.success(f"Email sent to {recipient_email}")
            except Exception as e:
                st.error(f"Failed to send email to {recipient_email}: {e}")
