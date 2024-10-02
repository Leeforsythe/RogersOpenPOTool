import streamlit as st
import pandas as pd
import yagmail

# Initialize yagmail for sending emails
SENDER_EMAIL = 'leeforsythe@rogerssupply.com'  # Use the correct variable name
APP_PASSWORD = 'cvho cgcn pown nrpy'  # Use the correct variable name

# Initialize yagmail SMTP client
yag = yagmail.SMTP(user=SENDER_EMAIL, password=APP_PASSWORD)

# Updated email template
email_template = """
Dear {vendor_name},

I hope this message finds you well. I'm reaching out regarding an outstanding order we have with you. Please find the details below:

Purchase Order Number: {po_number}
Manufacturer Part Number: {manufacturer_part_number}
Quantity Open: {qty_open}
PO Entry Date: {po_entry_date}

Could you kindly provide an update on the status of this order? If there are any issues or delays, please let us know how we can assist in expediting the process.

Thank you for your attention to this matter. We look forward to your prompt response.

Best regards,
{purchaser_name}
Rogers Supply Company
"""

# Vendor email mapping (predefined as before)
vendor_emails_data = pd.read_excel('vendor_emails.xlsx')

# Streamlit UI components
st.title('Open Purchase Order Email Automation')

# File upload for the Open PO report CSV file
uploaded_file = st.file_uploader("Upload the Open PO Report CSV file", type=["csv"])

if uploaded_file:
    # Read the uploaded CSV file
    open_po_data = pd.read_csv(uploaded_file)
    
    # Debug: Display column names before merge
    st.write("Open PO Data Columns:", open_po_data.columns.tolist())
    st.write("Vendor Emails Data Columns:", vendor_emails_data.columns.tolist())
    
    # Merge with vendor emails (corrected column names)
    merged_data = open_po_data.merge(vendor_emails_data, left_on='Vendor Name', right_on='VendorName', how='left')
    
    # Debug: Display merged column names
    st.write("Merged Data Columns:", merged_data.columns.tolist())
    
    # Clean up duplicate columns
    if 'Purchaser_x' in merged_data.columns:
        merged_data.drop(columns=['Purchaser_x'], inplace=True)
        merged_data.rename(columns={'Purchaser_y': 'Purchaser'}, inplace=True)

    # Generate email content using the new template and additional fields
    merged_data['EmailContent'] = merged_data.apply(lambda row: email_template.format(
        vendor_name=row['Vendor Name'],
        po_number=row['PO Number'],  # Adjust if your column name changes
        manufacturer_part_number=row['Manufacturer Part Number'],  # Adjust as needed
        qty_open=row['Qty Open'] if pd.notna(row['Qty Open']) else 'N/A',  # Adjust as needed
        po_entry_date=row['PO Entry Date'],  # Adjust as needed
        purchaser_name=row['Purchaser']  # Use the "Purchaser" column as the name
    ), axis=1)
    
    # Display the email content for preview
    st.write("Here is a preview of the emails:")
    st.write(merged_data[['Vendor Name', 'Email', 'EmailContent', 'Purchaser']].head())

    # Button to send emails
    if st.button('Send Emails'):
        for index, row in merged_data.iterrows():
            recipient_email = row['Email']
            subject = f"Follow Up on Purchase Order - {row['PO Number']}"
            body = row['EmailContent']
            reply_to_email = row['PurchaserEmail']  # Get the reply-to email address

            # Send email with a "Reply-To" header using headers dictionary
            try:
                yag.send(
                    to=recipient_email, 
                    subject=subject, 
                    contents=body,
                    headers={"Reply-To": reply_to_email}  # Set the "Reply-To" header here
                )
                st.success(f"Email sent to {recipient_email} with replies routed to {reply_to_email}")
            except Exception as e:
                st.error(f"Failed to send email to {recipient_email}: {e}")
