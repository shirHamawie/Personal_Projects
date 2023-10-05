import win32com.client
from datetime import datetime, timedelta
import PyPDF2
import re
import os
import tkinter as tk
from tkinter import messagebox

outlook = None
sub_folder_name = "Wolt"
file_path = "C:\\Users\\t-shamawie\\Downloads\\Wolt Codes.txt"
output_file = True
printing = True
delete_mail = True
clean_file = False
save_pdf_path = "C:\\Users\\t-shamawie\\Downloads"
attachment_symbol = "english"
code_pattern = r'CODE:\s+(\w+)'
valid_until_pattern = r'Valid until:\s+([A-Za-z]+\s+\d{1,2},\s+\d{4})'
amount_pattern = r'₪\s+([\d.]+)'
dates_in_past_30_days = []
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
        else:
            success = True
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


def write_to_file(code, date, amount):
    with open(file_path, 'a', encoding='utf-8') as file:
        try:
            file.write(f"Amount: {amount}₪\n")
            file.write(f"Code: {code}\n")
            file.write(f"Date: {date}\n")
            file.write("-" * 26 + "\n")
        except Exception as e:
            print_error(e)
            if not printing:
                print_values_blue(code, date, amount)


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
        date = datetime.strptime(valid_until, '%b %d %Y')
        date = date.replace(year=int(date.year - 5))
        day = date.strftime("%A")
        date = date.strftime('%b %d %Y')
        global dates_in_past_30_days
        if date in dates_in_past_30_days:
            dates_in_past_30_days.remove(date)

    else:
        date = "Valid until date not found"
        day = ""

    if output_file:
        write_to_file(code, date + " " + day, amount)
    if printing:
        print_values_blue(code, date + " " + day, amount)

    return True


def print_values_blue(code, date, amount):
    print("\033[94m" + "Amount =", amount + "₪,", "Code =", code + ",", "Date =", date, end="\n")


def print_error(e):
    print("\033[91m" + "Error Happened:", str(e) + '\033[0m')


def print_green(s, to_add):
    print('\033[92m' + s, to_add)


def print_pink(to_print, end='\n'):
    print('\033[95m', to_print, end=end)


def handled_dates():
    date_pattern = r'Date: (\w{3} \d{2} \d{4})'

    with open(file_path, 'r') as file:
        text = file.read()
        date_matches = re.findall(date_pattern, text)

    date_list = []

    for date_match in date_matches:
        date_obj = datetime.strptime(date_match, '%b %d %Y')
        date = date_obj.strftime('%b %d %Y')
        date_list.append(date)

    return date_list


def manipulate_dates():
    global dates_in_past_30_days
    current_date = datetime.now()
    thirty_days_ago = current_date - timedelta(days=30)
    current_day = current_date
    while current_day >= thirty_days_ago:
        day = current_day.strftime("%A")
        if day != "Friday" and day != "Saturday":
            dates_in_past_30_days.append(current_day.strftime('%b %d %Y'))
        current_day -= timedelta(days=1)

    handled_dates_list = handled_dates()
    for date in handled_dates_list:
        if date in dates_in_past_30_days:
            dates_in_past_30_days.remove(date)


def generate_date_list(month):
    today = datetime.now()

    first_day_of_month = today.replace(month=month, day=1)
    days_to_sunday = (first_day_of_month.weekday() - 6) % 7

    last_day_of_month = (first_day_of_month + timedelta(days=32)).replace(day=1) - timedelta(days=1)
    first_day_of_month -= timedelta(days=days_to_sunday)

    date_list = [first_day_of_month + timedelta(days=i) for i in
                 range((last_day_of_month - first_day_of_month).days + 1)]

    date_list = [date.strftime('%b %d %Y') for date in date_list]

    date_objects = [datetime.strptime(date, '%b %d %Y') for date in date_list]

    return date_list, date_objects


def print_calender(month):
    global dates_in_past_30_days

    today = datetime.now()

    date_list, date_objects = generate_date_list(month)

    print_pink("\t\t" + "--- " + datetime.strptime(str(month), '%m').strftime('%B') + " ---")
    print_pink('Sun  Mon  Tue  Wed  Thu  Fri  Sat')

    for i in range(len(date_list)):
        date = date_list[i]
        date_obj = date_objects[i]
        day = date_obj.strftime('%d')

        if today - date_obj > timedelta(days=30):
            print_pink(' P  ', end='')
        elif date in dates_in_past_30_days:
            print_pink(' X  ', end='')
        else:
            if datetime.strptime(date, '%b %d %Y').month != month:
                print_pink('    ', end='')
            elif date_obj > today:
                print_pink(' A  ', end='')
            elif date_obj.strftime('%A') == 'Saturday' or date_obj.strftime('%A') == 'Friday':
                print_pink(' O  ', end='')
            else:
                print_pink(f'{day:^4}', end='')

        if date_obj.strftime('%A') == 'Saturday':
            print()

    print()


def generate_month_ago_dates():
    today = datetime.now()
    return [today - timedelta(days=i) for i in range(30)]


def print_calenders():
    months_list = set()

    for date in generate_month_ago_dates():
        months_list.add(date.month)

    months_list = sorted(months_list)
    for month in months_list:
        print_calender(month)


def alert_popup(add):
    root = tk.Tk()
    root.withdraw()

    confirmation = messagebox.askyesnocancel("Confirmation", add + "!\nAre you sure you want to continue?",
                                             icon=messagebox.WARNING)
    root.destroy()

    return confirmation


def run():
    try:
        global delete_mail
        global clean_file
        global output_file
        global outlook

        if delete_mail:
            delete = alert_popup("You are about to delete the mails")
            if delete is None:
                exit(1)
            elif not delete:
                delete_mail = False

        if clean_file:
            clean = alert_popup("You are about to clean the output file")
            if clean is None:
                exit(1)
            elif not clean:
                clean_file = False
            else:
                with open(file_path, "w") as _:
                    pass

        if output_file:
            output = alert_popup("You are about to write to the output file")
            if output is None:
                exit(1)
            elif not output:
                output_file = False
            else:
                with open(file_path, "a") as file:
                    current_datetime = datetime.now()
                    date_time = current_datetime.strftime('%Y-%m-%d %H:%M:%S')
                    file.write(date_time + "\n")
                    file.write("=" * 26 + "\n")
        manipulate_dates()
        inbox = connect_to_outlook()
        iterate_mailbox(inbox)
        if outlook:
            del outlook
        if printing:
            print_green(str(counter), "Mails handled")
            print_green(str(earnings) + "₪", "Earnings")
            print()
            print_calenders()
    except Exception as e:
        print_error(e)


run()
