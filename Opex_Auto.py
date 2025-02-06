import win32com.client
import os
from weasyprint import HTML
import pandas as pd
import tkinter as tk
from tkinter import messagebox
import re

USER_EMAIL = "chris.mccormick@safcodental.com"
def load_approvers():
    file_path = "M:\\OPEX Automation\\Cleaned_Vendor_List.xlsx"
    df = pd.read_excel(file_path, sheet_name="Cleaned_Vendor_List")
    approver_dict = {}
    for _, row in df.iterrows():
        vendor = str(row["Vendor"]).strip().lower()
        approver_email = row["Approvers"]
        approver_dict[vendor] = approver_email if isinstance(approver_email, str) else ", ".join(approver_email.dropna())
    print("DEBUG: Approver List Loaded:", approver_dict)
    return approver_dict

def sanitize_filename(filename):
    return re.sub(r'[^a-zA-Z0-9_\-]', '_', filename)

def extract_invoice_details(subject, body):
    invoice_number = re.search(r'Invoice[\s#:]*([A-Za-z0-9-_]+)', subject + body, re.IGNORECASE)
    invoice_number = invoice_number.group(1) if invoice_number else "Unknown"
    invoice_date = re.search(r'(\d{1,2}/\d{1,2}/\d{2,4})', body)
    invoice_date = invoice_date.group(1) if invoice_date else "Unknown"
    invoice_date = re.sub(r'[^0-9]', '', invoice_date)[:6] if invoice_date != "Unknown" else "Unknown"
    return invoice_number, invoice_date

def find_vendor_name(subject, body, approvers):
    subject_lower = subject.lower()
    body_lower = body.lower()
    for vendor in approvers.keys():
        if f" {vendor} " in f" {subject_lower} " or f" {vendor} " in f" {body_lower} ":
            return vendor
    print(f"DEBUG: No exact vendor match found in subject or body.")
    return "Unknown"

def save_email_as_pdf_or_msg(message, save_path):
    # edit pdfkit config and save emails using i forget what the script is called
    try:
        html_body = message.HTMLBody
        if not html_body.strip():
            raise ValueError("Email body is empty or cannot be converted.")
        HTML(html_body).write_pdf(save_path)
        print(f"Saved invoice to {save_path}")
    except Exception as e:
        print(f"Error saving invoice as PDF: {e}")
        try:
            msg_save_path = save_path.replace(".pdf", ".msg")
            message.SaveAs(msg_save_path, 3)
            print(f"Saved email as .msg instead: {msg_save_path}")
        except Exception as msg_error:
            print(f"Error saving email as .msg: {msg_error}")

def process_emails():
    try:
        global approvers
        approvers = load_approvers()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.Folders[USER_EMAIL].Folders["Inbox"]
        messages = list(inbox.Items)
        save_directory = "T:\\Accounts Payable\\OPEX Filing\\1 - Invoices awaiting approval"
        approval_folder = namespace.Folders[USER_EMAIL].Folders["Inbox"].Folders["EXPENSES"].Folders["SAFCO EXPENSES"].Folders["** WAITING APPROVALS **"]
        print("Processing emails...")
        count = 0
        for message in messages[:]:
            try:
                print(f"Checking email: {message.Subject}")
                if message.Categories:
                    print(f"Skipping {message.Subject}, already processed.")
                    continue
                email_category = "Invoice"
                vendor_name = find_vendor_name(message.Subject, message.Body, approvers)
                invoice_number, invoice_date = extract_invoice_details(message.Subject, message.Body)
                approver_email = approvers.get(vendor_name, "AP@safcodental.com")
                if isinstance(approver_email, str):
                    approver_email = [email.strip() for email in approver_email.split(",")]
                valid_emails = [email for email in approver_email if "@" in email]
                if not valid_emails:
                    print(f"ERROR: No valid approver email found for vendor {vendor_name}")
                    continue
                approver_email_str = ", ".join(valid_emails)
                print(f"DEBUG: Vendor: {vendor_name} | Approver Email(s): {approver_email_str}")
                if email_category == "Invoice" and "approval" not in message.Body.lower():
                    filename = f"{vendor_name} - {invoice_date} - {invoice_number}.pdf"
                    pdf_filename = os.path.join(save_directory, sanitize_filename(filename))
                    print(f"Attempting to save invoice to: {pdf_filename}")
                    save_email_as_pdf_or_msg(message, pdf_filename)
                    forward = message.Forward()
                    forward.To = "chris.mccormick@safcodental.com"#approver_email_str
                    #forward.CC = "AP@safcodental.com" if "AP@safcodental.com" not in message.Recipients else ""
                    forward.Subject = f"Approval Required: {vendor_name} - {invoice_date} - {invoice_number}"
                    forward.Body = "Please review and approve the attached invoice."
                    forward.Send()
                    message.Categories = "Processed"
                    message.Move(approval_folder)
                    message.Save()
                    print(f"Forwarded invoice from {vendor_name} to {approver_email_str} and moved original email.")
            except Exception as e:
                print(f"Error processing email '{message.Subject}': {e}")
        print(f"Total emails processed: {count}")
        messagebox.showinfo("Success", f"Email processing completed. {count} emails processed.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        print(f"Error: {e}")

def run_script():
    root = tk.Tk()
    root.title("Email Processor")
    root.geometry("300x150")
    process_button = tk.Button(root, text="Process Emails", command=process_emails, font=("Arial", 12), padx=20, pady=10)
    process_button.pack(pady=40)
    root.mainloop()

if __name__ == "__main__":
    run_script()