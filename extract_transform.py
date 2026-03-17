#%%
import pandas as pd
import os
import re
from pathlib import Path
import pandas as pd
from typing import Dict, Optional
from pandas.tseries.holiday import USFederalHolidayCalendar
from pandas.tseries.offsets import CustomBusinessDay

#%%
pd.set_option('display.max_colwidth', None)
# Define column names
column_names = ['Title', 'Payment_Type', 'Receiver_email', 'Date', 'Star', 'Reference_num', 'Email_body']

BASE_DIR = Path(__file__).resolve().parent


def resolve_data_file(file_name: str) -> Path:
    """Resolve a file in Raw_Data, including nested export folders."""
    raw_data_dir = BASE_DIR / "Raw_Data"
    direct_file = raw_data_dir / file_name

    if direct_file.exists():
        return direct_file

    matches = list(raw_data_dir.rglob(file_name))
    if matches:
        return matches[0]

    raise FileNotFoundError(
        f"Could not find '{file_name}' in '{raw_data_dir}' or its subfolders."
    )

# Read CSV without headers
data_path = resolve_data_file("March_03_2026.csv")
df = pd.read_csv(data_path,names = column_names , header=None)

# Keep only donation emails (Zelle, CashAPP, Paypal)
df = df[df['Payment_Type'].isin(['Bank Email <bank_email>', 'Cash App <cash@square.com>', 'service@paypal.com <service@paypal.com>' ])]

df['Payment_Type']

paypal_df = df[df['Payment_Type'].isin(['service@paypal.com <service@paypal.com>'])]   
cash_App_df = df[df['Payment_Type'].isin(['Cash App <cash@square.com>'])]
zelle_df = df[df['Payment_Type'].isin(['Bank Email <bank_email>'])] #(#modified for privacy)
zelle_df.head()


#%%
# Filter zelle transactions to add only offerings "XYZ sent you a Zelle Payment"
zelle_receipts = zelle_df[zelle_df['Title'].str.contains('sent you a Zelle® payment', case=False, na=False )]

zelle_receipts.columns.to_list()



#%%
## Extract Necessary Information - Zelle


#%%
import pandas as pd
import re

# Assuming df is your DataFrame with an 'email' column containing the text

def extract_email_info(email_text):
    """Extract name, date, time, amount, and note from email text"""
    
    # Extract full name (appears after "From:" and before "sent you")
    name_match = re.search(r'([A-Z\s]+)\s+sent you', email_text)
    full_name = name_match.group(1).strip() if name_match else None
    
    # Extract date (format: MM/DD/YY)
    date_match = re.search(r'Date:\s*(\d{1,2}/\d{1,2}/\d{2,4})', email_text)
    date = date_match.group(1) if date_match else None
    
    # Extract time (format: H:MM AM/PM)
    time_match = re.search(r'(\d{1,2}:\d{2}\s*[AP]M)', email_text)
    time = time_match.group(1) if time_match else None
    
    # Extract amount (format: $XXX.XX)
    amount_match = re.search(r'Amount:\s*\$?([\d,]+\.\d{2})', email_text)
    amount = amount_match.group(1) if amount_match else None
    
    # Extract note (text inside parentheses after "Note:")
    note_match = re.search(r'Note:\s*(.+?)\s*Date:', email_text, re.DOTALL)
    note = note_match.group(1).strip() if note_match else None
    
    
    return {
        'full_name': full_name,
        'date': date,
        'time': time,
        'amount': amount,
        'note': note
    }

# Apply extraction to the email column
extracted_data = zelle_receipts['Email_body'].apply(extract_email_info)


# Convert to separate columns
zelle_receipts['full_name'] = extracted_data.apply(lambda x: x['full_name'])
zelle_receipts['date'] = extracted_data.apply(lambda x: x['date'])
zelle_receipts['time'] = extracted_data.apply(lambda x: x['time'])
zelle_receipts['amount'] = extracted_data.apply(lambda x: x['amount'])
zelle_receipts['Note'] = extracted_data.apply(lambda x: x['note'])



zelle_receipts['date'].isna().sum()

## Add check date
def time_extract(df):
    from pandas.tseries.holiday import USFederalHolidayCalendar
    from pandas.tseries.offsets import CustomBusinessDay
    
    # Create calendar and business day objects separately
    calendar = USFederalHolidayCalendar()
    us_bd = CustomBusinessDay(calendar=calendar)
    
    # Convert with exact formats
    df['date'] = pd.to_datetime(df['date'], format='%m/%d/%y')
    df['time'] = pd.to_datetime(df['time'], format='%I:%M %p')
    
    # Time condition: <= 9:59 PM
    time_condition = df['time'].dt.hour * 60 + df['time'].dt.minute <= 22 * 60
    
    # Business day check
    holidays = calendar.holidays(start=df['date'].min(), end=df['date'].max())
    is_weekend = df['date'].dt.dayofweek >= 5
    is_holiday = df['date'].isin(holidays)
    date_is_business_day = ~(is_weekend | is_holiday)

    # Keep date only if BOTH: business day AND time <= 9:59 PM
    keep_date = date_is_business_day & time_condition
    
    # Initialize realized_date with original date
    df['realized_date'] = df['date']

    # Apply next business day to rows that need changing
    needs_change = ~keep_date
    if needs_change.any():
        df.loc[needs_change, 'realized_date'] = df.loc[needs_change, 'date'].apply(
            lambda x: x + us_bd
        )
    return df

zelle_final = time_extract(df = zelle_receipts )
## Turn to excel 
path = "Output/extracted_payments_4.xlsx"
data_path = os.path.abspath(path)
zelle_final.to_excel(data_path, index=False)

#%%
### Paypal

import pandas as pd
import re
from datetime import datetime

# Filter only PayPal data
paypal_df = df[df['Payment_Type'].isin(['service@paypal.com <service@paypal.com>'])]

# Filter PayPal transactions to add only "You've got money" offerings
paypal_receipts_df = paypal_df[
    paypal_df['Title'].str.contains("You've got money", case=False, na=False)
].copy()  # BUG FIX 4: .copy() prevents SettingWithCopyWarning


def extract_email_data(email_text):
    """
    Extract structured data from email text with subject "You've got money"
    """
    if not isinstance(email_text, str) or "You've got money" not in email_text:
        return {}

    data = {}

    # Extract Date and Time (e.g. 11/23/25, 10:29 AM)
    date_match = re.search(r'(\d{1,2}/\d{1,2}/\d{2,4}),?\s*(\d{1,2}:\d{2}\s*[AP]M)', email_text)
    if date_match:
        data['Date'] = date_match.group(1)
        data['Time'] = date_match.group(2)

    # Extract amount received
    amount_match = re.search(r'you received\s*\$?([\d,]+\.?\d*)\s*USD', email_text, re.IGNORECASE)
    if amount_match:
        data['amount_received'] = float(amount_match.group(1).replace(',', ''))

    
        
    # Extract Sender name (before "sent you")
    sender_match = re.search(r'^([A-Za-z][^\r\n]+?)\s+sent you', email_text, re.IGNORECASE | re.MULTILINE)
    if sender_match:
        data['full_name'] = sender_match.group(1).strip()

    # Extract Fee
    fee_match = re.search(r'Fee\s*\$?([\d,]+\.?\d*)\s*USD', email_text, re.IGNORECASE)
    if fee_match:
        data['Fee'] = float(fee_match.group(1).replace(',', ''))

    # Extract Total
    total_match = re.search(r'Total\s*\$?([\d,]+\.?\d*)\s*USD', email_text, re.IGNORECASE)
    if total_match:
        data['Total'] = float(total_match.group(1).replace(',', ''))

    # Extract Transaction ID
    trans_id_match = re.search(r'Transaction ID:(.*?)(?:\\r\\n|\n)', email_text)
    if trans_id_match:
        data['Transaction_id'] = trans_id_match.group(1).strip()

    return data


# Apply extraction to the email column
extracted_data = paypal_receipts_df['Email_body'].apply(extract_email_data)

# BUG FIX 2 & 3: Use correct key names matching what extract_email_data stores,
# and use .get() so missing keys return None instead of crashing
paypal_receipts_df['full_name']      = extracted_data.apply(lambda x: x.get('full_name'))
paypal_receipts_df['Date']           = extracted_data.apply(lambda x: x.get('Date'))
paypal_receipts_df['Time']           = extracted_data.apply(lambda x: x.get('Time'))
paypal_receipts_df['Received_amount'] = extracted_data.apply(lambda x: x.get('amount_received'))
paypal_receipts_df['Fee']            = extracted_data.apply(lambda x: x.get('Fee'))
paypal_receipts_df['Total']          = extracted_data.apply(lambda x: x.get('Total'))
paypal_receipts_df['Transaction_id'] = extracted_data.apply(lambda x: x.get('Transaction_id'))

print(paypal_receipts_df.head())
paypal_receipts_df.columns.to_list()
paypal_receipts_df[paypal_receipts_df['full_name']== "None"].count()
#%%
##Cashapp Processing

import pandas as pd
import re
from datetime import datetime

# Filter only Cash App data
cash_app_df = df[df['Payment_Type'].isin(['Cash App <cash@square.com>'])]

import pandas as pd
import re

# Filter only Cash App data
cashapp_df = df[df['Payment_Type'].str.contains('cash@square.com', case=False, na=False)]

# Filter to "Payment received" emails only
cashapp_receipts_df = cashapp_df[
    cashapp_df['Title'].str.contains("Payment received", case=False, na=False)
].copy()


def extract_cashapp_data(email_text):
    """
    Extract structured data from Cash App "Payment received" emails.
    """
    if not isinstance(email_text, str) or "Payment received" not in email_text:
        return {}

    data = {}

    # --- Date & Time ---
    # Targets: "Date:\r\n2/1/26, 11:28 AM"
    date_match = re.search(
        r'Date:\r?\n(\d{1,2}/\d{1,2}/\d{2,4}),\s*(\d{1,2}:\d{2}\s*[AP]M)',
        email_text, re.IGNORECASE
    )
    if date_match:
        data['Date'] = date_match.group(1)   # e.g. 2/1/26
        data['Time'] = date_match.group(2)   # e.g. 11:28 AM

    # --- Full Name ---
    # Targets: "Sender: Yvon N Manahimg" (explicit label, most reliable)
    sender_match = re.search(r'Sender:\s*([^\r\n]+)', email_text, re.IGNORECASE)
    if sender_match:
        data['full_name'] = sender_match.group(1).strip()

    # --- Amount Received ---
    # Targets: "+$35.00"
    amount_match = re.search(r'\+\$?([\d,]+\.?\d*)', email_text)
    if amount_match:
        data['amount_received'] = float(amount_match.group(1).replace(',', ''))

    # --- Note / Purpose ---
    # Targets: "For offering" → captures any text after "For " before newline or "$"
    note_match = re.search(r'\bFor\s+([^\r\n\$]+?)(?:\r?\n|\s*\+?\$)', email_text, re.IGNORECASE)
    if note_match:
        data['Note'] = note_match.group(1).strip().title()  # e.g. "Offering"

    # --- Transaction ID ---
    # Targets: "Transaction number\r\n#D-3Rnpom"
    trans_id_match = re.search(r'Transaction number\r?\n(#[A-Za-z0-9\-]+)', email_text, re.IGNORECASE)
    if trans_id_match:
        data['Transaction_id'] = trans_id_match.group(1).strip()

    return data


# Apply extraction to email body column
extracted_data = cashapp_receipts_df['Email_body'].apply(extract_cashapp_data)

# Map to DataFrame columns
cashapp_receipts_df['full_name']        = extracted_data.apply(lambda x: x.get('full_name'))
cashapp_receipts_df['Date']             = extracted_data.apply(lambda x: x.get('Date'))
cashapp_receipts_df['Time']             = extracted_data.apply(lambda x: x.get('Time'))
cashapp_receipts_df['Received_amount']  = extracted_data.apply(lambda x: x.get('amount_received'))
cashapp_receipts_df['Note']             = extracted_data.apply(lambda x: x.get('Note'))
cashapp_receipts_df['Transaction_id']   = extracted_data.apply(lambda x: x.get('Transaction_id'))

cashapp_receipts_df.columns.to_list()

# %%
