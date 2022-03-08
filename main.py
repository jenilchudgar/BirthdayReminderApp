from sys import intern
from time import sleep
from win10toast_click import ToastNotifier
import openpyxl
import datetime
import webbrowser as web
from pyautogui import press

class MyToast:
    def __init__(self) -> None:
        self.toast = ToastNotifier()

    def show(self,title,msg,duration,on_click):
        self.toast.show_toast(title,msg,duration=duration,callback_on_click=on_click)

def create_person(name,date,month,mob):
    person = {
        "name" : name,
        "date" : date,
        "month" : month,
        "mobile" : mob
    }
    return person

def send_msg(name,mobile): 
    try:
        msg = f"Happy Birthday {name}! Enjoy your day. 🥳🥳"
        web.open(f'https://web.whatsapp.com/send?phone=+91{mobile}&text={msg}')
        sleep(10)
        press("enter")
        print("Successfully Sent!")
    
    except Exception as e:
        print(e)
        print("An Unexpected Error!")

def send_toast(name,no):
    toast = MyToast()
    toast.show(f"Today is {name}'s Birthday!","Should I send I message?",30,lambda: send_msg(name,no))

if __name__ == "__main__":
    book = openpyxl.load_workbook('data.xlsx')
    sheet = book.active

    people = [
            create_person(
                *[
                    sheet.cell(row=i, column=j).value
                    for j in range(1,5)
                ]
            )
        for i in range(2,11)
    ]
    
    date = int(datetime.datetime.now().strftime("%d"))
    month = int(datetime.datetime.now().strftime("%m"))

    for person in people:
        if (person["date"] == date) and (person["month"] == month):
            send_toast(person["name"],person["mobile"])