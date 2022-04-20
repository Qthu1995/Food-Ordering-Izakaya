# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import datetime as dt
orderdate = dt.datetime.now()

class Item:
    def __init__(self, name, price, qty=1, cate=""):
        self.name = name
        self.price = price
        self.qty = qty
        self.cate = cate
        
    def get_name(self):
        return self.name

    
    def edit_Qty(self, qty):
        self.qty = qty
    
    def Amt(self):
        return float(self.qty) * self.price
    
    def exp_Saledetail(self):
        return[self.cate, self.name, self.qty, self.price, self.Amt()]
    
    def __str__(self):
        qty = str(self.qty)
        name = self.name
        amt = str("{:,.0f}".format(self.Amt()))
        return (
            qty+ " " * (3 - len(qty)) + "x" + name + " " * (80 - len(name))
            + " " * (10 - len(amt)) + amt + "\n" + " " * 6 )
           
    
class Food(Item):
    def __init__(self, name, price, qty=1, cate='Food'):
        super().__init__(name, price, qty)
        self.cate = cate


class Drink(Item):
    def __init__(self, name, price, qty=1, cate='Drink'):
        super().__init__(name, price, qty)
        self.cate = cate
    

class Invoice:
    def __init__(self, cusID):
        self.cusID = cusID
        self.lst_item = []
        self.lst_expDetail = []
    
    def add_Item(self, item):
        self.lst_item.append(item)
    
    def get_totalQty(self):
        totalQty = 0
        for i in self.lst_item:
            totalQty += int(i.qty)
        return totalQty 
    
    def get_totalAmt(self):
        totalAmt = 0
        for i in self.lst_item:
            totalAmt += i.Amt()
        return totalAmt
    
    def exp_Saledetail(self):
        for i in self.lst_item:
            self.lst_expDetail.append([orderdate, self.cusID]+i.exp_Saledetail())
        return self.lst_expDetail
    
    def exp_Saleheader(self):
        return[orderdate, self.cusID, self.get_totalQty(), self.get_totalAmt()]
    
    def print_Invoice(self):
        print(f'\nOrder Details: {self.cusID}')
        print('-'*100)
        for i in self.lst_item:
            print(i.__str__())
        totalAmt = '{:,.0f}'.format(self.get_totalAmt())
        print('-'*100)
        print('Total Amount'+ ' '*(83-len(totalAmt))+totalAmt)
        
import xlwings as xw

data = xw.Book(r'C:\Users\Thu Beo\Desktop\UDEMY_COURSE\Python for data Science\Practise\Izakaya Ordering\Izakaya.xlsx')
lst_Foodname = data.sheets('Food').range('A2').options(expand = 'down').value
lst_Foodprice = data.sheets('Food').range('B2').options(expand = 'down').value
lst_Drinkname = data.sheets('Drink').range('A2').options(expand = 'down').value
lst_Drinkprice = data.sheets('Drink').range('B2').options(expand = 'down').value

sht_Saledetail = data.sheets('Saledetail')

sht_Saleheader = data.sheets('SaleOrderheader')
sht_Customer = data.sheets('Customer')

lst_CusID = data.sheets('Customer').range('A2').options(expand = 'down').value
lst_Cusname = data.sheets('Customer').range('B2').options(expand = 'down').value

def print_MainMenu():
    x1 = "*"*20 + 'ORDER PAGE' + "*"*20
    x2 = "(F)   FOOD"
    x3 = "(D)   DRINK"
    x4 = "-"*50
    main = [x1, x2, x3, x4]
    return print("\n".join(main))

def print_Food():
    main = []
    main.append("*"*14 + "TAKE AWAY MENU" + "*"*15)
    main.append("|NO|" + " "*2 + "|FOOD|" + " "*25 + "|PRICE|")
    for i in range(len(lst_Foodname)):
        x = (
            f"({i+1})" + " "*(5-len(str(i+1))) + lst_Foodname[i]+ " "*(30-len(lst_Foodname[i])) + 
            "{:,.0f}".format(lst_Foodprice[i]))
        main.append(x)
    return print("\n".join(main))

def print_Drink():
    main = []
    main.append("*"*19 + "ORDER DRINK" + "*"*20)
    main.append("|NO|" + " "*2 + "|DRINK NAME|" + " "*25 + "|PRICE|")
    for i in range(len(lst_Drinkname)):
        x = (
            f"({i+1})" + " "*(5-len(str(i+1))) + lst_Drinkname[i]+ " "*(36-len(lst_Drinkname[i])) + 
            "{:,.0f}".format(lst_Drinkprice[i]))
        main.append(x)
    return print("\n".join(main))

def main_Menu():
    phone = input("What is your phone number? ")
    if phone not in lst_CusID:
        print("You are not a member of our store. Please provide information")
        name = input("What's your name? ")
        birth = input("What's your date of birth? ")
        sht_Customer.range(f"A{len(lst_CusID)+2}").value = [phone, name, birth, 0]
    obj_Invoice = Invoice(phone)
    flag_done = True
    while flag_done:
        item = Order()
        obj_Invoice.add_Item(item)
        ques = input("Do you want to order more? (y/n)")
        if ques.lower() != 'y':
            flag_done = False
    sales_detail = obj_Invoice.exp_Saledetail()
    sales_header = obj_Invoice.exp_Saleheader()
    orderid = len(sht_Saleheader.range('A2').options(expand='down').value) + 1
    detailid = len(sht_Saledetail.range('A2').options(expand='down').value) + 1
    # Lay exp_Saledetail cua Class Item de export ra sheet Saledetail
    ex_detail = []
    for i in sales_detail:
        ex_detail.append([orderid] + i)
    sht_Saledetail.range(f"A{detailid + 1}").value = ex_detail
    # Lay exp_Saleheader cua Class Item de export ra sheet SaleOrderHeader
    sht_Saleheader.range(f"A{orderid + 1}").value = [[orderid]+sales_header]

    return obj_Invoice

def Order():
    print_MainMenu()
    select = None
    while select == None:
        ques = input("Please select your operation: ")
        if ques.upper() in ["F", "D"]:
            select = ques.upper()
        else:
            item = None
    if select == "F":
        item = OrderFood()
    elif select == "D":
        item = OrderDrink()
    return item 

def OrderFood():
    print_Food()
    flag_Food = True
    while flag_Food:
        id = input("Please order a number indecating your food? ")
        if id.isnumeric() == True and int(id) in range(1, len(lst_Foodname) + 1):
            foodname = lst_Foodname[int(id)-1]
            foodprice = lst_Foodprice[int(id)-1]
            obj_Food = Food(foodname, foodprice)
            flag_Food = False
        else:
            print(f"ERROR: Invalid Input ({id}). Try again!")
   
    flag_Food = True
    while flag_Food:
        qty = input("How many portions would you like to order?" )
        if qty.isnumeric() == True and int(qty) > 0:
            obj_Food.edit_Qty(qty)
            flag_Food = False
        else:
            print(f"ERROR: Invalid Input ({qty}). Try again!")
    return obj_Food


def OrderDrink():
    print_Drink()
    flag_Drink = True
    while flag_Drink:
        id = input("Please order a number indecating your Drink? ")
        if id.isnumeric() == True and int(id) in range(1, len(lst_Drinkname) + 1):
            drinkname = lst_Drinkname[int(id)-1]
            drinkprice = lst_Drinkprice[int(id)-1]
            obj_Drink = Drink(drinkname, drinkprice)
            flag_Drink = False
        else:
            print(f"ERROR: Invalid Input ({id}). Try again!")
            
    flag_Drink = True
    while flag_Drink:
        qty = input("How many glasses do you want to order?" )
        if qty.isnumeric() == True and int(qty) > 0:
            obj_Drink.edit_Qty(qty)
            flag_Drink = False
        else:
            print(f"ERROR: Invalid Input ({qty}). Try again!")
    return obj_Drink



run = main_Menu()
print(run.print_Invoice())
    