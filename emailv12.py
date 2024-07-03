import re
import dns.resolver
import smtplib
import pandas as pd
import os
import logging
from dns.exception import Timeout
from concurrent.futures import ThreadPoolExecutor, as_completed

# Configure logging
logging.basicConfig(filename='email_validation.log', level=logging.INFO, format='%(asctime)s - %(message)s')

# Regex pattern for email validation
EMAIL_REGEX = r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)"

def check_syntax(email):
    """ Check if the email matches the standard email address pattern. """
    return re.match(EMAIL_REGEX, email) is not None

def check_mail_server(domain):
    """ Check if the domain has MX records. """
    try:
        mx_records = dns.resolver.resolve(domain, 'MX')
        return len(mx_records) > 0
    except (dns.resolver.NoAnswer, dns.resolver.NXDOMAIN, Timeout):
        return False
    except Exception as e:
        logging.error(f"Error resolving MX records for {domain}: {e}")
        return False

def check_connection(email):
    """ Check if we can establish an SMTP connection to the email domain. """
    domain = email.split('@')[1]
    try:
        mx_records = dns.resolver.resolve(domain, 'MX')
        mx_record = str(mx_records[0].exchange)
        print(f"Trying to connect to MX record: {mx_record}")
        with smtplib.SMTP(mx_record, timeout=10) as server:
            server.set_debuglevel(1)  # Enable debug output
            server.ehlo_or_helo_if_needed()
            server.mail('test@example.com')
            code, message = server.rcpt(email)
            if code == 250:
                return True
            else:
                logging.info(f"Failed to validate {email}: {message}")
                return False
    except Exception as e:
        logging.error(f"Error connecting to mail server for {email}: {e}")
        return False

def is_gmail_domain(domain):
    """ Check if the domain is 'gmail.com'. """
    return domain.lower() == 'gmail.com'

def check_catch_all(domain):
    """ Check if the domain has a catch-all email address enabled. """
    try:
        mx_records = dns.resolver.resolve(domain, 'MX')
        mx_record = str(mx_records[0].exchange)
        with smtplib.SMTP(mx_record, timeout=10) as server:
            server.ehlo_or_helo_if_needed()
            server.mail('test@example.com')
            code, _ = server.rcpt(f'nonexistent@{domain}')
            return code == 250
    except Exception as e:
        logging.error(f"Error checking catch-all for {domain}: {e}")
        return False

def validate_email(email):
    """ Validate an email address using various checks. """
    if not check_syntax(email):
        return email, "Invalid syntax"
    
    domain = email.split('@')[1]
    
    if not check_mail_server(domain):
        return email, "No MX records found"
    
    if check_catch_all(domain):
        return email, "Valid"  # Consider catch-all domains as valid
    
    if not check_connection(email):
        return email, "Cannot connect to mail server"
    
    return email, "Valid"

def validate_emails(email_list):
    """ Validate a list of email addresses and return results. """
    results = []
    total_emails = len(email_list)
    
    with ThreadPoolExecutor(max_workers=50) as executor:
        futures = [executor.submit(validate_email, email) for email in email_list]
        
        for future in as_completed(futures):
            results.append(future.result())
            print(f"Processed: {len(results)}/{total_emails}", end='\r')
    
    print("\nValidation complete.")
    return results

def get_user_emails():
    """ Prompt user to enter email addresses. """
    email_list = []
    while True:
        email = input("Enter an email address (or 'done' to finish): ").strip()
        if email.lower() == 'done':
            break
        email_list.append(email)
    return email_list

# Ensure the directory exists before writing files
os.makedirs('outputs', exist_ok=True)

def save_to_excel(emails, valid_filename, invalid_filename):
    """ Save validation results to separate Excel files for valid and invalid emails. """
    valid_emails = [(email, status) for email, status in emails if status == "Valid"]
    invalid_emails = [(email, status) for email, status in emails if status != "Valid"]

    try:
        if valid_emails:
            pd.DataFrame(valid_emails, columns=['Email', 'Status']).to_excel(valid_filename, index=False, sheet_name='Valid')
            print(f"Valid emails saved to '{valid_filename}'")

        if invalid_emails:
            pd.DataFrame(invalid_emails, columns=['Email', 'Status']).to_excel(invalid_filename, index=False, sheet_name='Invalid')
            print(f"Invalid emails saved to '{invalid_filename}'")
    except Exception as e:
        print(f"Error saving to files: {e}")

# Main program flow
if __name__ == "__main__":
    email_list = get_user_emails()
    if email_list:
        results = validate_emails(email_list)
        valid_filename = 'outputs/valid_emails.xlsx'
        invalid_filename = 'outputs/invalid_emails.xlsx'
        save_to_excel(results, valid_filename, invalid_filename)
        
        print("\nDetailed results:")
        for email, status in results:
            print(f"{email}: {status}")
    else:
        print("No emails entered.")
