import openpyxl as xl
from datetime import datetime, timedelta, date
from prettytable import PrettyTable
import calendar
import os

path = os.getcwd()+'/'
#path = 'C:/Users/aizat/Documents/project/'

def log_in(user):
    wb = xl.load_workbook(path + user)
    sh = wb['Logins']
    login = input('Enter your login: >>> ')
    password = input('Enter your password: >>> ')
    for lg, pw in sh['B2':'C'+str(sh.max_row)]:
        if lg.value == login and pw.value == password:
            print('You are successfully logged in!'.center(104))
            break
    else:
        while lg.value != login or pw.value != password:
            print('Wrong login or password, please try again'.center(104))
            login = input('Enter your login: >>> ')
            password = input('Enter your password: >>> ')
            for lg, pw in sh['B2':'C'+str(sh.max_row)]:
                if lg.value == login and pw.value == password:
                    print('You are successfully logged in!'.center(104))
                    break
    return login

def reg_in(user):
    wb = xl.load_workbook(path + user)
    sh = wb['Logins']
    name = input('Enter your name and surname: >>> ')
    for i in sh['A']:
        if name == i.value:
            print('You already have an account, please log in'.center(104))
            login = log_in(user)
            break
    else:
        login = input('Create your login: >>> ')
        password = input('Create your password: >>> ')
        c = 1
        while sh.cell(row=c, column=1).value != None:
            c += 1
        for nm, lg, pw in sh['A'+str(c):'C'+str(c)]:
            nm.value, lg.value, pw.value = name, login, password
        print('You have successfully registered'.center(104))
        wb.save(user)
    return login

def define_user(login, user):
    wb = xl.load_workbook(path + user)
    sh = wb['Logins']
    for nm, lg in zip(sh['A'], sh['B']):
        if login == lg.value:
            return lg.row, nm.value

def start(account):
    print('Already have an account?'.center(104))
    while True:
        is_acc = input('(y/n) >>> ').strip().lower()
        if is_acc == 'y':
            login = log_in(account)
            break
        elif is_acc == 'n':
            login = reg_in(account)
            break
        else:
            print('Wrong command'.center(104))
    row, name = define_user(login, account)
    print(f'Greetings {name}!'.center(104))
    print('Dial the menu number to work with the program'.center(104))
    return row, name

def med_hist_header():
    print('-'*104)
    print('| {:<17}{:<21}{:<45}{:<17} | '.format('Date', 'Diagnosis', 'Treatment', 'Term of treatment'))
    print('-'*104)

def med_hist_body(row):
    wb = xl.load_workbook(path + 'patients.xlsx')
    sh = wb['Diagnosis']
    for dg, dt, tr, ds in sh['B'+str(row):'E'+str(row)]:
        dt = datetime.strftime(dt.value, '%d.%m.%Y')
        print('| {:<17}{:<21}{:<45}{:<17} | '.format(dt, dg.value, tr.value, f'{ds.value} days'))
        print('-'*104)

def medical_history(name, string='', header='Medical history'): 
    wb = xl.load_workbook(path + 'patients.xlsx')
    sh = wb['Diagnosis']
    names = []
    for i in sh['A']:
        if i.value == 'Name':
            continue
        if i.value not in names and i.value != None:
            names.append(i.value)
    if name not in names:
        print(string.center(104))
        return string
    else:
        first, second = True, True
        for i in sh['A']:
            if name == i.value:
                if first:
                    if header != '':
                        print(header.center(104))
                    med_hist_header()
                    first = False
                med_hist_body(i.row)

def last_date(name):
    wb = xl.load_workbook(path + 'patients.xlsx')
    sh = wb['Diagnosis']
    dates = [sh['C'+str(i.row)].value for i in sh['A'] if name != 'Name' and name != None and name == i.value]
    last = max(dates)
    for j in sh['C']:
        if last == j.value and name == sh['A'+str(j.row)].value:
            row = j.row
    return last, row

def last_record(name):
    wb = xl.load_workbook(path + 'patients.xlsx')
    sh = wb['Diagnosis']
    try:
        last, row = last_date(name)
        dt = datetime.strftime(last, '%d.%m.%Y')
        dg = sh['B'+str(row)].value
        tr = sh['D'+str(row)].value
        ds = sh['E'+str(row)].value
        print('Last record'.center(104))
        print('-'*104)
        print('| {:<17}{:<21}{:<45}{:<17} | '.format('Date', 'Diagnoses', 'Treatment', 'Term of treatment'))
        print('-'*104)
        print('| {:<17}{:<21}{:<45}{:<17} | '.format(dt, dg, tr, f'{ds} days'))
        print('-'*104)
    except:
        print('Your medical history is empty. It will be filled by your doctor'.center(104))

def count_remainder(row, last):
    wb = xl.load_workbook('patients.xlsx')
    sh = wb['Diagnosis']
    day =  sh['E'+str(row)].value
    delta = datetime.now() - last
    remainder = int(day) - delta.days
    return remainder, day

def treatment(name):
    wb = xl.load_workbook(path + 'patients.xlsx')
    sh = wb['Diagnosis']
    try:
        last, row = last_date(name)
        remainder, day = count_remainder(row, last)
        if remainder > 0:
            print(f"The treatment duration - {day} days, {remainder} days left until the end".center(104))
        else:
            print(f"You have no treatment at the moment".center(104))
    except:
        print('Your medical history is empty'.center(104))

def schedule(): 
    wb = xl.load_workbook(path + 'doctors.xlsx')
    sh = wb['Schedule']
    print('Doctors schedule'.center(104))
    header = [sh.cell(row=1, column=i).value for i in range(1, 7)]
    table = PrettyTable(header)
    for a, b, c, d, e, f in sh['A2':'F10']:
        if a.value == None:
            a.value = ''
        values = [a.value, b.value, c.value, d.value, e.value, f.value]
        if values == [''] + [None for i in range(5)]:
            values = [''] + ['' for i in range(5)]
        table.add_row(values)
    print(table)

def patient_info(row, string, whose, header):
    wb = xl.load_workbook(path + 'patients.xlsx')
    sh = wb['Logins']
    info = [sh.cell(row=row, column=i).value for i in range(4, 8)]
    try:
        info[2] = info[2].strftime('%d.%m.%Y')
        print(header.center(104))
        table = PrettyTable(['Height', 'Weight', 'Birthday date', 'Blood group'])
        table.add_row(['{:^20}'.format(info[0]), '{:^20}'.format(info[1]), '{:^30}'.format(info[2]), '{:^30}'.format(info[3])])
        print(table)
    except:
        print(string.center(104))
        hgt = float(input(f'Enter {whose} height: >>> '))
        wgt = float(input(f'Enter {whose} weight: >>> '))
        bth = input(f'Enter {whose} birthdate: >>> ')
        bth = datetime.strptime(bth, '%d.%m.%Y').date()
        bld = input(f'Enter {whose} bloodgroupe: >>> ')
        for h, w, bd, bg in sh['D'+str(row):'G'+str(row)]:
            h.value, w.value, bd.value, bg.value = hgt, wgt, bth, bld
        wb.save('patients.xlsx')
        patient_info(row, string, whose, header)

def patient():
    row, name = start('patients.xlsx')
    print('''    1 - show the medical history
    2 - show the last record in the medical history
    3 - show the number of days of treatment
    4 - show doctors' schedule
    5 - show my info
    6 - return to the main menu
    7 - exit''')
    while True:
        number = input('>>> ')
        if number == '1':
            s = 'Your medical history is empty. It will be filled by your doctor'
            medical_history(name, s)
        elif number == '2':
            last_record(name)
        elif number == '3':
            treatment(name)
        elif number == '4':
            schedule()
        elif number == '5':
            s = 'Your date is not yet in the system. Please fill in info about yourself'
            patient_info(row, s, 'your', 'My info')
        elif number == '6':
            break
        elif number == '7':
            print(number)
            exit('The program is over, we look forward to your return!'.center(104))
        else:
            print('Such command does not exists'.center(104))

def medassistant():
    row, name = start('medassistants.xlsx')
    print('''    1 - show a list of procedures
    2 - search for a patient
    3 - show a list of medassistants' errands
    4 - execute an errand
    5 - show executed errands
    6 - return to the main menu
    7 - exit''')


def doctor():
    row, name = start('doctors.xlsx')
    print('''    1 - show a list of patients receiving treatment
    2 - show the total number of patients
    3 - show a list of medassistants' errands
    4 - write an errand for a medassistant
    5 - show executed errands
    6 - search for a patient
    7 - diagnose a patient
    8 - return to the main menu
    9 - exit''')


def maindoctor():
    row, name = start('maindoctors.xlsx')
    print('''    1 - show the list of medassistances
    2 - show the list of doctors
    3 - show amount of patients 
    4 - show the employee with the highest salary
    5 - show the employee with the lowest salary
    6 - return to the main menu
    7 - exit''')


commands = {
    'patient': patient,
    'medassistant': medassistant,
    'doctor': doctor,
    'maindoctor': maindoctor
}    

print('Welcome to the information system "AIS Hospital"'.center(104))
while True:
    print('''Enter your account type:
    -patient
    -medassistant
    -doctor
    -maindoctor
    -exit''')
    command = input('>>>> ')  
    if command in commands:
        commands[command]()
    elif command == 'exit':
        exit('The program is over, we look forward to your return!'.center(104))
    else:
        print('Wrong account type, please enter again'.center(104))
