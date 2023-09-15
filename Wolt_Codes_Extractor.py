import win32com.client
import PyPDF2
import re
import os

outlook = None
sub_folder_name = "Wolt"
output_file = True
file_path = "C:\\Users\\t-shamawie\\Downloads\\Wolt Codes.txt"
printing = True
delete_mail = False
save_pdf_path = "C:\\Users\\t-shamawie\\Downloads"
attachment_symbol = "english"
code_pattern = r'CODE:\s+(\w+)'
valid_until_pattern = r'Valid until:\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})'


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
                handle_code(text)
            os.remove(pdf_path)


def handle_mails(folder_name):
    for item in folder_name.Items:
        handle_attachments(item)
        if delete_mail:
            item.Delete()


def extract_text(pdf_reader):
    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        text = page.extract_text()
        return text


def write_to_file(code, valid_until):
    with open(file_path, 'a') as file:
        file.write(f"Code: {code}\n")
        file.write(f"Valid until: {valid_until}\n")
        file.write("-" * 26 + "\n")


def handle_code(text):
    code_match = re.search(code_pattern, text)
    valid_until_match = re.search(valid_until_pattern, text)

    if code_match:
        code = code_match.group(1)
    else:
        code = "Code not found"

    if valid_until_match:
        valid_until = valid_until_match.group(1)
    else:
        valid_until = "Valid until date not found"

    if output_file:
        write_to_file(code, valid_until)
    if printing:
        print("\033[94m" + "Code =", code, "Valid =", valid_until, end="")


def run():
    inbox = connect_to_outlook()
    iterate_mailbox(inbox)
    global outlook
    if outlook:
        del outlook


run()
