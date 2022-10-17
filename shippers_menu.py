# William Leonard, Saxon Enterprises LLC
# Contact: wfleonard@saxonenterprises.net 732-673-4260
# Shippers Menu program for GH AR Shipping Data 
# 09-16-2022    Error checking menu item entries

import os
import time
import fedexap as fedex
import dhlap as dhl
clear = lambda: os.system('clear')

def print_menu():

    menu_options = {
    1: 'FEDEX Shipping Feed',
    2: 'DHL Shipping Feed',
    3: 'UPS Shipping Feed',
    4: 'Exit',
    }
    print( 'Gabriel Hearst AR for Shippers Program Menu')
    print()
    for key in menu_options.keys():
        print(key, '--', menu_options[key] )

def option1():
     fedex.main()
     
def option2():
     dhl.main()

def option3():
     print('UPS Shipping Data Transfer NOT Available Yet')
     time.sleep(2)

def main():
    while(True):
        clear()
        print_menu()
        print()
        option = ''
        option = input('Enter your option #: ')
        #Check what choice was entered and act accordingly
        if option == "1":
            option1()
        elif option == "2":
            option2()
        elif option == "3":
            option3()
        elif option == "4":
            print('You exited the Shippers Program')
            exit()
        else:
            print('Invalid option. Please enter a number between 1 and 4.')
            time.sleep(1)

if __name__ == "__main__":
    main()