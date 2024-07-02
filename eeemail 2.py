import re
import dns.resolver
import smtplib
import pandas as pd
from dns.exception import Timeout

def check_syntax(email):
    regex = r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)"
    return re.match(regex, email) is not None

def check_mail_server(domain):
    max_retries = 3
    retry_count = 0
    while retry_count < max_retries:
        try:
            mx_records = dns.resolver.resolve(domain, 'MX')
            return len(mx_records) > 0
        except (dns.resolver.NoAnswer, dns.resolver.NXDOMAIN):
            return False
        except Timeout:
            retry_count += 1
            print(f"Timeout occurred. Retrying ({retry_count}/{max_retries})...")
    return False

def check_connection(email):
    domain = email.split('@')[1]
    try:
        mx_records = dns.resolver.resolve(domain, 'MX')
        mx_record = mx_records[0].exchange.to_text()
        server = smtplib.SMTP(mx_record)
        server.set_debuglevel(0)
        server.helo()
        server.quit()
        return True
    except Exception as e:
        return False

def is_gmail_domain(domain):
    return domain.lower() == 'gmail.com'

def check_catch_all(domain):
    max_retries = 3
    retry_count = 0
    while retry_count < max_retries:
        try:
            mx_records = dns.resolver.resolve(domain, 'MX')
            mx_record = mx_records[0].exchange.to_text()
            server = smtplib.SMTP(mx_record)
            server.set_debuglevel(0)
            server.helo()

            if is_gmail_domain(domain):
                server.mail('test@gmail.com')
                code, message = server.rcpt('fake-email@gmail.com')
                server.quit()
                return code != 550
            else:
                server.mail('test@example.com')
                code, message = server.rcpt('fake-email@' + domain)
                server.quit()
                return code == 250
        
        except Timeout:
            retry_count += 1
            print(f"Timeout occurred. Retrying ({retry_count}/{max_retries})...")
        except Exception as e:
            return False

    return False

def validate_email(email):
    if not check_syntax(email):
        return False, "Invalid syntax"
    
    domain = email.split('@')[1]
    
    if not check_mail_server(domain):
        return False, "No MX records found"
    
    if not check_connection(email):
        return False, "Cannot connect to mail server"
    
    if check_catch_all(domain):
        return False, "Domain has catch-all enabled"
    
    return True, "Email is valid"

def validate_emails(email_list):
    valid_emails = []
    invalid_emails = []
    
    for email in email_list:
        is_valid, message = validate_email(email)
        if is_valid:
            valid_emails.append(email)
        else:
            invalid_emails.append((email, message))
    
    return valid_emails, invalid_emails

def get_user_emails():
    email_list = []
    while True:
        email = input("Enter an email address (or 'done' to finish): ").strip()
        if email.lower() == 'done':
            break
        email_list.append(email)
    return email_list

def save_to_excel(valid_emails):
    df = pd.DataFrame(valid_emails, columns=['Valid Emails'])
    df.to_excel('valid_emails.xlsx', index=False)

# Main program flow
if __name__ == "__main__":
    email_list = get_user_emails()
    if email_list:
        valid_emails, invalid_emails = validate_emails(email_list)
        if valid_emails:
            save_to_excel(valid_emails)
            print(f"Valid emails saved to 'valid_emails.xlsx'")
        if invalid_emails:
            for email, message in invalid_emails:
                print(f"{email}: INVALID - {message}")
    else:
        print("No emails entered.")
