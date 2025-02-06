import win32com.client
import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox
import re
from datetime import datetime

USER_EMAIL = "bee.morgan@safcodental.com"
def load_approvers():
    file_path = "T:\\OPEX Automation\\Cleaned_Vendor_List.xlsx"
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
    all_text = ""
    all_text += subject + body
    invoice_number = re.search(r'Invoice[\s#:]*([A-Za-z0-9-_]+)', subject + body, re.IGNORECASE)
    invoice_number = invoice_number.group(1) if invoice_number else "Unknown"
    invoice_dates = extract_dates(all_text)
    return invoice_number, invoice_dates[0]

def extract_dates(text):
    date_formats = [
        r'\d{1,2}[-/]\d{1,2}[-/]\d{2,4}',
        r'\d{1,2} \w+ \d{2,4}',
        r'\d{1,2}-\w+-\d{2,4}',
        r'\d{1,2} \w+ \d{2,4}',
        r'\w{3} \d{1,2}, \d{2,4}',
        r'\w+ \d{1,2},? \d{2,4}',
        r'\d{1,2}-\d{1,2}-\d{2,4}',
        r'\d{2}.\d{2}.\d{4}',
        r'\d{2}-\w{3}-\d{4}',
        r'\d{2}.\d{2}.\d{4}',
        r'\d{2}/[A-Za-z]{3}/\d{4}',
        r'[A-Za-z]+[\xa0\s]*\d{1,2},[\xa0\s]*\d{4}'
    ]
    print(text)
    dates = []
    for date_format in date_formats:
        dates.extend(re.findall(date_format, text))
    print('dates: ', dates)
    valid_dates = []
    for date in dates:
        try:
            valid_dates.append(convert_to_mmddyy(date))
        except Exception as e:
            continue

    print(valid_dates)

    valid_dates = list(set(valid_dates))

    return valid_dates if len(valid_dates) > 0 else None

def convert_to_mmddyy(date):
    try:
        formats = [
            "%m.%d.%Y", "%d.%m.%Y",
            "%m-%d-%Y", "%m/%d/%Y", "%m-%d-%y", "%m/%d/%y",
            "%d-%m-%Y", "%d/%m/%Y", "%d-%m-%y", "%d/%m/%y",
            "%B %d, %Y", "%b %d, %Y", "%B %e, %Y",
            "%d %B %Y", "%d %b %Y",
            "%d-%b-%Y", "%d-%b-%y",
            "%Y-%m-%d"
        ]

        for fmt in formats:
            try:
                return datetime.strptime(date, fmt).strftime("%m%d%y")
            except ValueError:
                pass

        raise ValueError(f"Date format for '{date}' not recognized")

    except Exception as e:
        raise ValueError(f"Error converting date '{date}': {e}") from e
def find_vendor_name(subject, body, approvers):
    subject_lower = subject.lower()
    body_lower = body.lower()
    for vendor in approvers.keys():
        if f" {vendor} " in f" {subject_lower} " or f" {vendor} " in f" {body_lower} ":
            return vendor
    print(f"DEBUG: No exact vendor match found in subject or body.")
    return "Unknown"

def save_attachment_with_unique_name(filename, attachment, folder_path):
    global attachments_extracted
    try:
        base, ext = os.path.splitext(filename)
        save_path = os.path.join(folder_path, filename)
        counter = 1
        while os.path.exists(save_path):
            new_filename = f"{base}_{counter}{ext}"
            save_path = os.path.join(folder_path, new_filename)
            counter += 1
        attachment.SaveAsFile(save_path)
        attachments_extracted += 1
        print(f"Saved attachment as: {os.path.basename(save_path)}")
    except Exception as e:
        print(f"Error saving attachment: {e}")
def extract_attachments(save_path, filename, attachments):
    if not os.path.exists(save_path):
        os.makedirs(save_path)

    for attachment in attachments:
        if attachment.FileName.lower().endswith(('.pdf', '.html')):
            save_attachment_with_unique_name(filename, attachment, save_path)
            print(f"Saved PDF/html file: {attachment.FileName}")
        else:
            print(f"Ignored non-PDF attachment: {attachment.FileName}")


def process_emails():

    try:
        #init approver list
        global approvers
        approvers = load_approvers()

        #init outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.Folders[USER_EMAIL].Folders["Inbox"]
        messages = list(inbox.Items)

        #directories
        save_directory = "T:\\Accounts Payable\\OPEX Filing\\1 - Invoices awaiting approval"
        approval_folder = inbox.Folders["EXPENSES"].Folders["SAFCO EXPENSES"].Folders["** WAITING APPROVALS **"]

        print("Processing emails...")
        count = 0

        for message in messages[:]:
            try:
                print(f"Checking email: {message.Subject}")
                # checks if the message has a category, if it has a category we ignore and iterate past it
                if message.Categories:
                    print(f"Skipping {message.Subject}, already processed.")
                    continue

                # sets eml category to invoice
                email_category = "Invoice"

                # attempts to find vendor name in subject and body (enhance by extracting text from attachments)
                vendor_name = find_vendor_name(message.Subject, message.Body, approvers)

                # attempts to grab invoice num and date from email (again, enhance by extracting text)
                # optional improvement: code every regular expression for all invoices
                # ^^^ use sage exports to grab one invoice from every active vendor, and have
                # ai make the regular expressions based on a csv?
                invoice_number, invoice_date = extract_invoice_details(message.Subject, message.Body)
                print(f"[chris] invoice_number: {invoice_number}, invoice_date: {invoice_date}")

                # if the above function returns none for the number, sets it to unknown.
                if invoice_date == None:
                    invoice_date = "unknown"

                # finds approver using vendor name
                approver_email = approvers.get(vendor_name, "AP@safcodental.com")

                # if approver email exists and is a string:
                # we change approver_email to account for what appears to be multiple approvers?
                # debug: see what we get from approver_email
                if isinstance(approver_email, str):
                    print(f'[chris] approver_email: {approver_email}')
                    approver_email = [email.strip() for email in approver_email.split(",")]

                # for every email in approver_email that was split above, we make sure theres an @ in the email.
                valid_emails = [email for email in approver_email if "@" in email]

                # iterate past this email and go to the next one if no valid emails
                if not valid_emails:
                    print(f"ERROR: No valid approver email found for vendor {vendor_name}")
                    continue

                # debug this
                approver_email_str = ", ".join(valid_emails)
                print(f'[chris] approver_email_str: {approver_email_str}')

                # statement with vendor name and approver emails
                print(f"DEBUG: Vendor: {vendor_name} | Approver Email(s): {approver_email_str}")

                # enter this block if the email category is invoice and the word approval is not in the message body
                # i think it would be better to check if the email came from internal before checking for the word
                # approval, implement this in kevin-the-terminatior
                if email_category == "Invoice": # and "approval" not in message.Body.lower():

                    # set file name and sanitize the name
                    # there should be no need for sanitize_filename but we will keep it as a fail safe.
                    filename = f"{vendor_name} - {invoice_date} - {invoice_number}.pdf"
                    pdf_filename = os.path.join(save_directory, filename)

                    #currently editing this to extract pdf attachments from emails
                    print(f"Attempting to save invoice to: {pdf_filename}")
                    extract_attachments(save_directory, pdf_filename, message.Attachments)

                    # forward the message to the correct approver
                    # in testing atm, sending to myself to see if it works.
                    forward = message.Forward()
                    forward.To = approver_email_str
                    forward.CC = "AP@safcodental.com" if "AP@safcodental.com" not in message.Recipients else ""
                    forward.Subject = f"Approval Required from {approver_email_str}: {vendor_name} - {invoice_date} - {invoice_number}"
                    forward.Body = "Please review and approve the attached invoice."
                    forward.Send()

                    #change the category and move invoice to approval folder, this works properly
                    message.Categories = "Processed"
                    message.Move(approval_folder)
                    #message.Save() this just throws an exception

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