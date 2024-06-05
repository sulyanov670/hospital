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


def patient():
    row, name = start('patients.xlsx')
    print('''    1 - show the medical history
    2 - show the last record in the medical history
    3 - show the number of days of treatment
    4 - show doctors' schedule
    5 - show my info
    6 - return to the main menu
    7 - exit''')


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
