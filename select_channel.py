import sys
from gpiozero import LED
from time import sleep

A = LED(17)
B = LED(27)
INH = LED(22)

INH.on()

while True:
    user_input = input('Select channel (0 1 2 3): ')
    INH.on()

    try:
        user_input = int(user_input)
        print(user_input)

        if user_input == 0:
            INH.off()
            A.off()
            B.off()
        elif user_input == 1:
            INH.off()
            A.on()
            B.off()
        elif user_input == 2:
            INH.off()
            A.off()
            B.on()
        elif user_input == 3:
            INH.off()
            A.on()
            B.on()
        else:
            INH.on()
    
    except ValueError:
        print('wrong, please enter 0 1 2 3.')
        INH.on()









