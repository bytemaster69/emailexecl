import re
import dns.resolver
import smtplib
import pandas as pd
import os
import logging
from dns.exception import Timeout
from concurrent.futures import ThreadPoolExecutor, as_completed
from time import sleep

# Configure logging
logging.basicConfig(filename='email_validation.log', level=logging.INFO, format='%(asctime)s - %(message)s')

# Regex pattern for email validation
EMAIL_REGEX = r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)"

# Set alternative DNS servers
dns.resolver.default_resolver = dns.resolver.Resolver()
dns.resolver.default_resolver.nameservers = ['8.8.8.8', '1.1.1.1']

def retry_connection(func):
    def wrapper(*args, **kwargs):
        retries = 3
        for attempt in range(retries):
            result = func(*args, **kwargs)
            if result or attempt == retries - 1:
                return result
            sleep(2 ** attempt)  # Exponential backoff
    return wrapper

def check_syntax(email):
    """ Check if the email matches the standard email address pattern. """
    return re.match(EMAIL_REGEX, email) is not None

@retry_connection
def check_mail_server(domain):
    """ Check if the domain has MX records. """
    try:
        mx_records = dns.resolver.resolve(domain, 'MX', lifetime=10)
        return len(mx_records) > 0
    except Exception as e:
        logging.error(f"Error resolving MX records for {domain}: {e}")
        return False

@retry_connection
def check_connection(email):
    """ Check if we can establish an SMTP connection to the email domain. """
    domain = email.split('@')[1]
    try:
        mx_records = dns.resolver.resolve(domain, 'MX')
        server_address = str(mx_records[0].exchange)
        with smtplib.SMTP(server_address, timeout=10) as server:
            server.ehlo_or_helo_if_needed()
            server.mail('your-email@example.com')
            code, message = server.rcpt(email)
            if code in [250, 251]:  # Consider 251 also positive
                return True
    except Exception as e:
        logging.error(f"Error connecting to mail server for {email}: {e}")
    return False

def validate_email(email):
    """ Validate an email address using various checks. """
    if not check_syntax(email):
        return email, "Invalid syntax"
    
    domain = email.split('@')[1]
    
    if not check_mail_server(domain):
        return email, "No MX records found"
    
    if not check_connection(email):
        return email, "Cannot connect to mail server"
    
    return email, "Valid"

def validate_emails(email_list):
    """ Validate a list of email addresses and return results. """
    results = []
    with ThreadPoolExecutor(max_workers=50) as executor:
        futures = [executor.submit(validate_email, email) for email in email_list]
        total_emails = len(futures)
        for i, future in enumerate(as_completed(futures)):
            results.append(future.result())
            print(f"Processed: {i + 1}/{total_emails}", end='\r')
    print("\nValidation complete.")
    return results

def get_user_emails():
    """ Prompt user to enter email addresses. """
    email_list = []
    while True:
        email = input("Enter an email address (or 'done' to finish): ").strip()
        if email.lower() == 'done':
            break
        if email:
            email_list.append(email)
    return email_list

# Ensure the directory exists before writing files
os.makedirs('outputs', exist_ok=True)

def save_to_excel(emails, valid_filename, invalid_filename):
    """ Save validation results to separate Excel files for valid and invalid emails. """
    valid_emails = [(email, status) for email, status in emails if status == "Valid"]
    invalid_emails = [(email, status) for email, status in emails if status != "Valid"]

    if valid_emails:
        pd.DataFrame(valid_emails, columns=['Email', 'Status']).to_excel(valid_filename, index=False, sheet_name='Valid')
        print(f"Valid emails saved to '{valid_filename}'")

    if invalid_emails:
        pd.DataFrame(invalid_emails, columns=['Email', 'Status']).to_excel(invalid_filename, index=False, sheet_name='Invalid')
        print(f"Invalid emails saved to '{invalid_filename}'")

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
