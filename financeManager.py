import csv
from unicodedata import category
import gspread
import time

MONTH = 'nov'
file = f"amex_{MONTH}.csv"
chaseFile = f"chase_{MONTH}.csv"
capitalOneFile = f"capitalOne_{MONTH}.csv"
quickSilverFile = f"quickSilver_{MONTH}.csv"
transactions = []

BILL_NAMES = {"ABODE HOME SECURITY",'AMANDA WANNARKA     COON RAPIDS         MN','TMOBILE*AUTO PAY    800-937-8997        WA','COMCAST ST PAUL CS 1(800)266-2278       MN','MEMBERSHIP FEE',"MIDWEST RADIOLOGY PAROSEVILLE           MN","MINNEAPOLIS RADIOLOGPLYMOUTH            MN",
                "NORTH MEMORIAL HEALTROBBINSDALE         MN","THE URGENCY ROOM 00-BLOOMINGTON         MN","JEWELERS-MUTUAL-PMNT800-558-6411        WI"}
FOOD_NAMES = {"CUB FOODS 0000000001BLAINE              MN",'CUB FOODS #1598 0000BLAINE              MN','ALDI 72108 000000000BLAINE              MN',"ALDI 72016 000000000BLAINE              MN","ALDI                BLAINE              MN","GglPay CUB FOODS #15BLAINE              MN"}
EATING_OUT_NAMES = {"MCDONALD'S 6359     OAKDALE             MN","DONATELLI'S RESTAURAWHITE BEAR LA       MN","WHITE CASTLE  080034BLAINE              MN","GglPay PANERA BREAD 612-656-6147        MN","CHICK-FIL-A #04341 0WOODBURY            MN","PAPA JOHN'S #1722 00BLAINE              MN","ARBYS #7475 BLAINE 0BLAINE              MN","TACO BELL #734948 73BLAINE              MN","DAVANNI'S #17 - COONCOON RAPIDS         MN","DOMINO'S PIZZA      BLAINE              MN","LITTLE CAESARS 3505-651-332-700         MN","DAIRY QUEEN #14832 0BLAINE              MN","LITTLE CAESARS 3505-651-332-700         MN",
                    "CRISP & GREEN - GC  6122089240          MN","PP*CORENUTRITI      BLAINE              MN","GglPay CRISP & GREENBLAINE              MN","PIZZA HUT 039460 000WOODBURY            MN"}
SHOPPING_NAMES = {"COSTCO WHSE #1021",'WAL-MART SUPERCENTERBLAINE              MN','WAL-MART 3498 3498  BLAINE              MN',"WAL-MART SUPERCENTERVADNAIS HEIGHTS     MN","GglPay TARGET       BLAINE              MN","GglPay WALGREENS #72BLAINE              MN","GglPay TARGET       BLAINE              MN","GglPay CVS/PHARMACY BLAINE              MN",
                        }
WANT_NAMES = {"TICKETMASTER","AplPay APPLE.COM/BILINTERNET CHARGE     CA",'COFFEEBARK, LLC     Osceola             WI',"DICK'S FRESH MARKET OSCEOLA             WI","VALLEY SPIRITS      OSCEOLA             WI","AT HOME STORE #134 1BLAINE              MN","MICHAELS STORES 1599BLAINE              MN","TOP TEN LIQUORS 0848BLAINE              MN",
                "LEGO SHOP AT HOME SE(800)835-4386       CT","HI SCORE VIDEO GAMESMINNEAPOLIS         MN","FIVE BELOW 4040 0000BLAINE              MN","GglPay CRISP & GREENBLAINE              MN","UNDERCOVERTOURIST.COHOLLY HILL          FL","MICHAELS STORES 1599BLAINE              MN","FIVE BELOW 4040 0000BLAINE              MN","WAL-MART SUPERCENTERVADNAIS HEIGHTS     MN"}
IGNORE_NAMES = {'Withdrawal from CHASE CREDIT CRD EPAY','Withdrawal from AMEX EPAYMENT ACH PMT','ONLINE PAYMENT - THANK YOU','MOBILE PAYMENT - THANK YOU',"Payment Thank You-Mobile","CAPITAL ONE MOBILE PYMT","CREDIT-CASH BACK REWARD"}
PET_NAMES = {"BLAINE AREA PET HOSPITAL","BLAINE AREA PET HOSPBLAINE              MN","CHEWY.COM           (800)672-4399       FL"}
ENTERTAINMENT_NAMES = {"DISNEY PLUS         BURBANK             CA","SPOTIFY USA         NEW YORK            NY","HBO MAX             NEW YORK            NY","HLU*HULU 10985931514HULU.COM/BILL       CA","GglPay AMC ONLINE 96LEAWOOD             KS","AEN* LIFETIMEMOVIECLNEW YORK CITY       NY"}
GAS_NAMES = {"HOLIDAY STATIONSTOREWOODBURY            MN","KWIK TRIP  925009258BLAINE              MN","HOLIDAY STATIONS 041BLAINE              MN"}
HOME_NAMES = {"THE HOME DEPOT      BLAINE              MN","LOWE'S OF BLAINE, MNBLAINE              MN","MENARDS BLAINE MN 00BLAINE              MN"}
VACATION_NAMES = {"MAC PARKING RESERVATSAINT PAUL          MN","WDW DINING RESV     LAKE BUENA VI       FL"}

sum = 0

# Determine the Sub Category of a Transaction
def subCategory(description):
    if description in PET_NAMES:
        return "Need"
    elif description in BILL_NAMES:
        return "Need"
    elif description in GAS_NAMES:
        return "Need"
    elif description in FOOD_NAMES:
        return "Need"
    elif description in EATING_OUT_NAMES:
        return "Want"
    elif description in ENTERTAINMENT_NAMES:
        return "Want"
    elif description in VACATION_NAMES:
        return "Want"
    elif description == "Transfer to Pet Fund":
        return "Need"
    else:
        return "Unlabled"


#Process the Amex CSV file format
def amex(file,BILL_NAMES,FOOD_NAMES,EATING_OUT_NAMES,SHOPPING_NAMES,WANT_NAMES,IGNORE_NAMES,PET_NAMES,ENTERTAINMENT_NAMES,GAS_NAMES,HOME_NAMES,VACATION_NAMES):
    #Read CSV file
    with open(file,mode='r') as csv_file:
        csv_reader = csv.reader(csv_file)
        headers = next(csv_file) 
        for row in csv_reader:
            date = row[0]
            description = row[1]
            amount = float(row[4]) * -1
            category = 'unlabled'
            if description in HOME_NAMES:
                category = 'Home'
            if description =="DELTA AIR LINES":
                category = 'Vacation'
            if description in WANT_NAMES:
                category = "Wants"
            if description in SHOPPING_NAMES:
                category = "Shopping"
            if description in BILL_NAMES:
                category = 'Bills'
            if description == "TARGET       BLAINE              MN":
                category = 'Shopping'
            if description in FOOD_NAMES:
                category = 'Food'
            if description == 'BT*PIRATE SHIP * POSJACKSON             WY':
                category = "bricklink-store"
            if description in PET_NAMES:
                category = "Pets"
            if description in EATING_OUT_NAMES:
                category = "Eating Out"
            if description in GAS_NAMES and amount>6:
                category = "Gas"
            if description in GAS_NAMES and amount<=6:
                category = "Eating Out"
            if description in ENTERTAINMENT_NAMES:
                category = "Entertainment"
            if description in IGNORE_NAMES:
                continue #ingore transaction
            if description in VACATION_NAMES:
                category = "Vacation"
            if description == "USPS KIOSK 266301955MINNEAPOLIS         MN":
                category == "Other"
            transaction = ((amount,"AMEX Card",description,category,subCategory(description),date))
            print(transaction)
            transactions.append(transaction)
        return transactions

# Process the Chase CSV file format
def chase(chaseFile,BILL_NAMES,FOOD_NAMES,EATING_OUT_NAMES,SHOPPING_NAMES,WANT_NAMES,IGNORE_NAMES,PET_NAMES,ENTERTAINMENT_NAMES,GAS_NAMES,HOME_NAMES,VACATION_NAMES):
    #Read CSV file
    with open(chaseFile,mode='r') as csv_file:
        csv_reader = csv.reader(csv_file)
        headers = next(csv_file) 
        for row in csv_reader:
            date = row[0]
            description = row[2]
            amount = float(row[5])
            category = row[4]
            if description in VACATION_NAMES:
                category = 'Vacation'
            if description in HOME_NAMES:
                category = 'Home'
            if description =="DELTA AIR LINES":
                category = 'Vacation'
            if description in WANT_NAMES:
                category = "Wants"
            if description in SHOPPING_NAMES:
                category = "Shopping"
            if description in BILL_NAMES:
                category = 'Bills'
            if description == "TARGET       BLAINE              MN":
                category = 'Shopping'
            if description in FOOD_NAMES:
                category = 'Food'
            if description == 'BT*PIRATE SHIP * POSJACKSON             WY':
                category = "bricklink-store"
            if description in PET_NAMES:
                category = "Pets"
            if description in EATING_OUT_NAMES:
                category = "Eating Out"
            if description in GAS_NAMES and amount>6:
                category = "Gas"
            if description in GAS_NAMES and amount<=6:
                category = "Eating Out"
            if description in ENTERTAINMENT_NAMES:
                category = "Entertainment"
            if description in IGNORE_NAMES:
                continue #ingore transaction
            if description == "USPS KIOSK 266301955MINNEAPOLIS         MN":
                category == "Other"
            #transaction = ((date,description,amount,category))
            transaction = ((amount,"Amazon Card",description,category,subCategory(description),date))
            print(transaction)
            transactions.append(transaction)
        return transactions

#Process the Capital One CSV file format
# Process the Chase CSV file format
def capitalOne(capitalOneFile,BILL_NAMES,FOOD_NAMES,EATING_OUT_NAMES,SHOPPING_NAMES,WANT_NAMES,IGNORE_NAMES,PET_NAMES,ENTERTAINMENT_NAMES,GAS_NAMES,HOME_NAMES,VACATION_NAMES):
    #Read CSV file
    with open(capitalOneFile,mode='r') as csv_file:
        csv_reader = csv.reader(csv_file)
        headers = next(csv_file) 
        for row in csv_reader:
            date = row[1]
            description = row[4]
            amount = float(row[2])
            category = row[3]
            if description in VACATION_NAMES:
                category = 'Vacation'
            if description in HOME_NAMES:
                category = 'Home'
            if description =="DELTA AIR LINES":
                category = 'Vacation'
            if description in WANT_NAMES:
                category = "Wants"
            if description in SHOPPING_NAMES:
                category = "Shopping"
            if description in BILL_NAMES:
                category = 'Bills'
            if description == "TARGET       BLAINE              MN":
                category = 'Shopping'
            if description in FOOD_NAMES:
                category = 'Food'
            if description == 'BT*PIRATE SHIP * POSJACKSON             WY':
                category = "bricklink-store"
            if description in PET_NAMES:
                category = "Pets"
            if description in EATING_OUT_NAMES:
                category = "Eating Out"
            if description in GAS_NAMES and amount>6:
                category = "Gas"
            if description in GAS_NAMES and amount<=6:
                category = "Eating Out"
            if description in ENTERTAINMENT_NAMES:
                category = "Entertainment"
            if description == "Withdrawal to Emergency Fund XXXXX9141":
                category = "Transfer"
                description = "Transfer to Pet Fund"
            if description == "Withdrawal to Home Emergency Fund XXXXX0254":
                category = "Transfer"
                description = "Transfer to Annual Services Fund"
            if description == "Bills - Withdrawal to Zech Family Account XXXXXXX4866":
                category = "Bills"
                description = "Money for Bills Transfer"
            if description == "Withdrawal from TARGET DEBIT CRD ACH TRAN Blaine MN":
                category = "Shopping"
                description = "Target Debit Card Payment"
            if description =="Withdrawal to 360 Performance Savings XXXXXXX9298":
                category = "Transfer"
                description = "Deficit Savings"
            if description.find("XXXXX0254") != -1:
                description = 'Transfer to Annual Services Fund'
                category = "Transfer"
            if description.find("XXXXX9141") != -1:
                category = "Transfer"
                description = "Transfer to Pet Fund"
            if description == 'Deposit from 360 Checking XXXXX3389':
                category = "Deposit"
            if description == 'Deposit from ROWLISON DIVERSI PAYROLL':
                category = "Deposit"
            if description in IGNORE_NAMES:
                continue #ingore transaction
            if description == "USPS KIOSK 266301955MINNEAPOLIS         MN":
                category == "Other"
            #transaction = ((date,description,amount,category))
            transaction = ((amount,"Captial One Card",description,category,subCategory(description),date))
            print(transaction)
            transactions.append(transaction)
        return transactions
# Process Quick Silver File Format
def quickSilver(quickSilverFile,BILL_NAMES,FOOD_NAMES,EATING_OUT_NAMES,SHOPPING_NAMES,WANT_NAMES,IGNORE_NAMES,PET_NAMES,ENTERTAINMENT_NAMES,GAS_NAMES,HOME_NAMES,VACATION_NAMES):
    #Read CSV file
    with open(quickSilverFile,mode='r') as csv_file:
        csv_reader = csv.reader(csv_file)
        headers = next(csv_file)
        for row in csv_reader:
            try:    
                date = row[0]
            except: break
            description = row[3]
            try:
                amount = float(row[5]) * -1
            except: amount = 0
            category = row[4]
            if description in VACATION_NAMES:
                category = 'Vacation'
            if description in HOME_NAMES:
                category = 'Home'
            if description =="DELTA AIR LINES":
                category = 'Vacation'
            if description in WANT_NAMES:
                category = "Wants"
            if description in SHOPPING_NAMES:
                category = "Shopping"
            if description in BILL_NAMES:
                category = 'Bills'
            if description == "TARGET       BLAINE              MN":
                category = 'Shopping'
            if description in FOOD_NAMES:
                category = 'Food'
            if description == 'BT*PIRATE SHIP * POSJACKSON             WY':
                category = "bricklink-store"
            if description in PET_NAMES:
                category = "Pets"
            if description in EATING_OUT_NAMES:
                category = "Eating Out"
            if description in GAS_NAMES and amount>6:
                category = "Gas"
            if description in GAS_NAMES and amount<=6:
                category = "Eating Out"
            if description in ENTERTAINMENT_NAMES:
                category = "Entertainment"
            if description in IGNORE_NAMES:
                continue #ingore transaction
            if description == "USPS KIOSK 266301955MINNEAPOLIS         MN":
                category == "Other"
            
            transaction = ((amount,"Captial One Card",description,category,subCategory(description),date))
            print(transaction)
            transactions.append(transaction)
        return transactions


def categorySums(worksheet):
    worksheet.update_acell('H2','=SUMIF(D:D,"Bills",A:A)')
    worksheet.update_acell('G2','Bills')
    worksheet.update_acell('H3','=SUMIF(D:D,"Eating Out",A:A)')
    worksheet.update_acell('G3','Eating Out')
    worksheet.update_acell('H4','=SUMIF(D:D,"Entertainment",A:A)')
    worksheet.update_acell('G4','Entertainment')
    worksheet.update_acell('H5','=SUMIF(D:D,"Food",A:A)')
    worksheet.update_acell('G5','Food')
    worksheet.update_acell('H6','=SUMIF(D:D,"bricklink-store",A:A)')
    worksheet.update_acell('G6','bricklink-store')
    worksheet.update_acell('H7','=SUMIF(D:D,"Home",A:A)')
    worksheet.update_acell('G7','Home')
    worksheet.update_acell('H8','=SUMIF(D:D,"Vacation",A:A)')
    worksheet.update_acell('G8','Vacation')
    worksheet.update_acell('H9','=SUMIF(D:D,"Pets",A:A)')
    worksheet.update_acell('G9','Pets')
    worksheet.update_acell('H10','=SUMIF(D:D,"Deposit",A:A)')
    worksheet.update_acell('G10','Deposits')
    worksheet.update_acell('H11','=SUMIF(D:D,"unlabled",A:A)')
    worksheet.update_acell('G11','Unlabled')   


gc = gspread.service_account(filename='C:\\Users\\Aaron\\OneDrive\\FinanceManager\\service_account.json');
#gc = gspread.service_account() # Default location %AppDat% Roaming Gspread

sa = gspread.service_account()
sh = sa.open("Joint Budget")

worksheet = sh.worksheet(f"{MONTH}")


rows = amex(file,BILL_NAMES,FOOD_NAMES,EATING_OUT_NAMES,SHOPPING_NAMES,WANT_NAMES,IGNORE_NAMES,PET_NAMES,ENTERTAINMENT_NAMES,GAS_NAMES,HOME_NAMES,VACATION_NAMES)
rows = chase(chaseFile,BILL_NAMES,FOOD_NAMES,EATING_OUT_NAMES,SHOPPING_NAMES,WANT_NAMES,IGNORE_NAMES,PET_NAMES,ENTERTAINMENT_NAMES,GAS_NAMES,HOME_NAMES,VACATION_NAMES)
rows = capitalOne(capitalOneFile,BILL_NAMES,FOOD_NAMES,EATING_OUT_NAMES,SHOPPING_NAMES,WANT_NAMES,IGNORE_NAMES,PET_NAMES,ENTERTAINMENT_NAMES,GAS_NAMES,HOME_NAMES,VACATION_NAMES)
rows = quickSilver(quickSilverFile,BILL_NAMES,FOOD_NAMES,EATING_OUT_NAMES,SHOPPING_NAMES,WANT_NAMES,IGNORE_NAMES,PET_NAMES,ENTERTAINMENT_NAMES,GAS_NAMES,HOME_NAMES,VACATION_NAMES)


#Insert Header Row
worksheet.insert_row(["Amount","Card Type","Description","Category","Sub-Category","Date"],2)

for row in rows:
    worksheet.append_row([row[0],row[1],row[2],row[3],row[4],row[5]],2)
    time.sleep(1.2) 

print("Program Complete")

#Formatting
categorySums(worksheet)

