import xlwt, os
import datetime as dt
from time import sleep

app_name = "Lotto King"
version = "0.01.00"
build_date = "06-27-2017"

email = "David@DREAM-Enterprise.com"
name = app_name + "  V" + version

#define system varibles
width = 60
lines = 28
cent_width = (width-1)
splash_wait = 5
wait = 3

txt_ticket="list_tickets.txt"
lst_ticket=[]
n=0
currdate = dt.date.today().strftime("%Y%m%d")
currtime = dt.datetime.now().strftime("%H%M%S")

#define cosole size and color
os.system("mode con: cols=" + str(width) + " lines=" + str(lines))
os.system("color F")
os.system("cls")
os.system("echo off")

#set styles for worksheet
style0 = xlwt.easyxf('font: name Times New Roman, color-index black, bold on',
                     num_format_str='#,##0.00')
style1 = xlwt.easyxf(num_format_str='YYYY-MM-DD')
style2 = xlwt.easyxf()

#building sheet
book_name = "LottoKing-"+currdate+"-"+currtime+".xls"
book = xlwt.Workbook()
sheet = book.add_sheet(currdate)
sheet.col(0).width = 256*20
sheet.col(1).width = 256*5
sheet.col(2).width = 256*5
sheet.col(3).width = 256*5
sheet.col(4).width = 256*5

def logo():
    
    print()        
    print()
    print()
    print()
   
    print(("------------------------------------------").center(cent_width))    
    print()
    print((" _      ____ _______ _______ ____  ").center(cent_width))
    print(("| |    / __ \__   __|__   __/ __ \ ").center(cent_width))
    print(("| |   | |  | | | |     | | | |  | |").center(cent_width))
    print(("| |   | |  | | | |     | | | |  | |").center(cent_width))
    print(("| |___| |__| | | |     | | | |__| |").center(cent_width))
    print(("|______\____/  |_|     |_|  \____/ ").center(cent_width))
    print()
    print((" _  _______ _   _  _____ ").center(cent_width))
    print(("| |/ /_   _| \ | |/ ____|").center(cent_width))
    print(("| ' /  | | |  \| | |  __ ").center(cent_width))
    print(("|  <   | | | . ` | | |_ |").center(cent_width))
    print(("| . \ _| |_| |\  | |__| |").center(cent_width))
    print(("|_|\_\_____|_| \_|\_____|").center(cent_width))
    print()
    print(("------------------------------------------").center(cent_width))
    print (name.center(cent_width))
    print (email.center(cent_width))
    sleep(wait)
                             
                         
#define commands
def header():
    os.system("cls")
    print ()
    print ((name).center(cent_width))
    print (email.center(cent_width))
    print ()
    sleep (.1)


def start_sheet():
    global n
    sheet.write(n,0, "LOTTO KING",style0)
    sheet.write_merge(n,n,1,2 ,currdate,style1)
    
    n+=1
    sheet.write(n,0, "Ticket Name",style0)
    sheet.write(n,1, "Cost",style0)
    sheet.write(n,2, "Start",style0)
    sheet.write(n,3, "End",style0)
    sheet.write(n,4, "Sales",style0)
    n+=1
   
def read_list():
    with open(txt_ticket) as f:
        temp_ticket=f.readlines()
    
        for i in temp_ticket:
            lst_ticket.append(i.strip('\n'))

def write_tickets():
    global n
    header()
    for i in lst_ticket:
        x=i.split(",")
        sheet.write(n,0,x[0])
        sheet.write(n,1,int(x[1]))
        starting = input("  Starting Number for "+x[0]+": ")
        sheet.write(n,2,int(starting))
        ending = input("  Ending Number for "+x[0]+": ")
        sheet.write(n,3,int(ending))
        sales=((int(ending)-int(starting))*int(x[1]))
        sheet.write(n,4,sales)
        n+=1
    
def add_logo():
    global n
    n+=3
    click="http://www.DREAM-Enterprise.com"
    sheet.write(n,0,xlwt.Formula('HYPERLINK("%s";"DREAM-Enterprise")' % click))
        
def save_sheet():    
    header()
    print("  Saving File: "+book_name)
    print()
    book.save(book_name)
    sleep(1) 
    print("  File Saved")
    print()
    sleep(1)
    print("  Program Ending...")
    sleep(2)
    
if __name__=="__main__":
    logo()
    start_sheet()
    read_list()
    write_tickets()
    
    add_logo()
    save_sheet()