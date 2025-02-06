import os
import re
import sys
import win32com.client
import comtypes.client
import pandas as pd
from PyPDF2 import PdfMerger

USER_EMAIL = "chris.mccormick@safcodental.com"

def get_outlook_folder(namespace, folder_path):
    """Dynamically finds the correct Outlook folder."""
    folders = folder_path.split("/")
    folder = namespace.Folders.Item(folders[0])
    for sub_folder in folders[1:]:
        folder = folder.Folders.Item(sub_folder)
    return folder

def find_matching_email(invoice_number):
    """Searches for an approval email matching the invoice number in specific Outlook folders, prioritizing the email body."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        waiting_folder = get_outlook_folder(namespace, f"{USER_EMAIL}/Inbox/EXPENSES/** WAITING APPROVALS **")
        processed_folder = get_outlook_folder(namespace, f"{USER_EMAIL}/Inbox/EXPENSES/* PROCESSED EXPENSES *")
    except Exception as e:
        print(f"❌ Outlook connection failed: {e}")
        return None
    
    for folder in [processed_folder, waiting_folder]:
        print(f"📂 Checking folder: {folder.Name}")
        folder.Items.Sort("[ReceivedTime]", True)  # Sort by newest first
        
        for message in folder.Items:
            email_text = message.Body.lower().strip()
            subject_text = message.Subject.lower().strip()
            if invoice_number.lower().strip() in email_text or invoice_number.lower().strip() in subject_text:
                print(f"✅ Match found for invoice {invoice_number} in email: {message.Subject}")
                return message
    return None

def print_email_to_pdf(message, invoice_number):
    """Prints the approval email to a PDF."""
    temp_folder = "T:\\Accounts Payable\\OPEX Filing\\Temp Emails"
    os.makedirs(temp_folder, exist_ok=True)
    pdf_path = os.path.join(temp_folder, f"Approval_{invoice_number}.pdf")
    
    try:
        message.SaveAs(pdf_path, 3)  # Save as PDF format
        print(f"📄 Saved approval email for invoice {invoice_number} as PDF: {pdf_path}")
        return pdf_path
    except Exception as e:
        print(f"❌ Error saving email as PDF: {e}")
        return None

def merge_pdfs(invoice_file, approval_pdf):
    """Merges the invoice and approval email into one PDF."""
    if not os.path.exists(invoice_file) or not os.path.exists(approval_pdf):
        print(f"❌ One or both files missing: {invoice_file}, {approval_pdf}")
        return None
    
    try:
        merged_pdf_path = invoice_file.replace("1 - Invoices awaiting approval", "2 - Approved, to be posted")
        merger = PdfMerger()
        merger.append(invoice_file)
        merger.append(approval_pdf)
        merger.write(merged_pdf_path)
        merger.close()
        print(f"✅ Merged PDF saved: {merged_pdf_path}")
        return merged_pdf_path
    except Exception as e:
        print(f"❌ PDF merge failed: {e}")
        return None

def process_existing_invoices():
    """Finds and merges approvals for invoices that were not originally processed by the script."""
    invoice_folder = "T:\\Accounts Payable\\OPEX Filing\\1 - Invoices awaiting approval"
    approved_folder = "T:\\Accounts Payable\\OPEX Filing\\2 - Approved, to be posted"
    
    print(f"🔍 Checking folder: {invoice_folder}")
    if not os.path.exists(invoice_folder):
        print(f"❌ Invoice folder not found: {invoice_folder}")
        return
    
    for file in os.listdir(invoice_folder):
        print(f"🔎 Checking file: {file}")
        match = re.search(r' - \d{6} - ([A-Za-z0-9]+)\.pdf$', file)
        if match:
            invoice_number = match.group(1)
            invoice_file = os.path.join(invoice_folder, file)
            
            approval_email = find_matching_email(invoice_number)
            
            if approval_email:
                approval_pdf = print_email_to_pdf(approval_email, invoice_number)
                merged_pdf = merge_pdfs(invoice_file, approval_pdf)
                os.remove(invoice_file)  # Remove original invoice file
                print(f"✅ Processed approval for invoice: {invoice_number}")
            else:
                print(f"⚠️ No approval found for invoice: {invoice_number}")

if __name__ == "__main__":
    print("🔍 Script started...")
    process_existing_invoices()
    print("✅ Script finished!")