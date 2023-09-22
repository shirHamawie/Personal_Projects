import win32com.client
import PyPDF2
import re
import os

outlook = None
sub_folder_name = "Wolt"
output_file = True
file_path = "C:\\Users\\t-shamawie\\Downloads\\Wolt Codes.txt"
printing = True
delete_mail = True
save_pdf_path = "C:\\Users\\t-shamawie\\Downloads"
attachment_symbol = "english"
code_pattern = r'CODE:\s+(\w+)'
valid_until_pattern = r'Valid until:\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})'
amount_pattern = r'₪\s+([\d.]+)'
counter = 0
earnings = 0


def connect_to_outlook():
    global outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 represents the inbox folder
    return inbox


def iterate_mailbox(folder):
    if sub_folder_name:
        for sub_folder in folder.Folders:
            if sub_folder.name == sub_folder_name:
                handle_mails(sub_folder)
    else:
        handle_mails(folder)


def handle_attachments(item):
    for attachment in item.Attachments:
        attachment_name = attachment.FileName.lower()
        if attachment_name.endswith(".pdf") and attachment_symbol in attachment_name:
            pdf_path = os.path.join(save_pdf_path, attachment.FileName)
            attachment.SaveAsFile(pdf_path)
            with open(pdf_path, "rb") as pdf_file:
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                text = extract_text(pdf_reader)
                success = handle_code(text)
                global counter
                counter += 1
            if success:
                os.remove(pdf_path)
    return success


def handle_mails(folder_name):
    for item in folder_name.Items:
        success = handle_attachments(item)
        if delete_mail and success:
            item.Delete()


def extract_text(pdf_reader):
    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        text = page.extract_text()
        return text


def write_to_file(code, valid_until, amount):
    with open(file_path, 'a', encoding='utf-8') as file:
        try:
            file.write(f"Amount: {amount}₪\n")
            file.write(f"Code: {code}\n")
            file.write(f"Valid until: {valid_until}\n")
            file.write("-" * 26 + "\n")
        except Exception as e:
            print_error(e)
            if not printing:
                print_values_blue(code, valid_until, amount)


def handle_code(text):
    code_match = re.search(code_pattern, text)
    valid_until_match = re.search(valid_until_pattern, text)
    amount_match = re.search(amount_pattern, text)

    if amount_match:
        amount = amount_match.group(1)
        global earnings
        earnings += float(amount)
    else:
        amount = "Amount not found"

    if code_match:
        code = code_match.group(1)
    else:
        code = "Code not found"

    if valid_until_match:
        valid_until = valid_until_match.group(1).replace(',', '')
    else:
        valid_until = "Valid until date not found"

    if output_file:
        write_to_file(code, valid_until, amount)
    if printing:
        print_values_blue(code, valid_until, amount)

    return True


def print_values_blue(code, valid_until, amount):
    print("\033[94m" + "Amount =", amount + "₪,", "Code =", code + ",", "Valid =", valid_until, end="\n")


def print_error(e):
    print("\033[91m" + "Error Happened:", str(e) + '\033[0m')


def print_green(s, to_add):
    print('\033[92m' + s, to_add)


def run():
    try:
        inbox = connect_to_outlook()
        iterate_mailbox(inbox)
        global outlook
        if outlook:
            del outlook
        if printing:
            print_green(str(counter), "Mails handled")
            print_green(str(earnings) + "₪", "Earnings")
    except Exception as e:
        print_error(e)


run()
