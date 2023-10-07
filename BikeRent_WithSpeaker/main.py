import win32com.client as wincom
speak = wincom.Dispatch("SAPI.SpVoice")

speak.speak("Enter number of bikes you want to put in stock")
try:
    Bstock= int(input("Enter number of bikes you want to put in stock: "))
except:
    print("enter valid Entry")

class BikeShop:
    def __init__(self,stock):
        self.stock=stock

    def displayBike(self):
        speak.speak(f"total Bikes in Stock are:{ self.stock}")
        print("Bikes in Stock are: ", self.stock)


    def bikeForRent(self,q):

        if q<=0:
            speak.speak("Enter a positive value of greater than zero")
            print("Enter a positive value of greater than zero")
        elif q>self.stock:
            speak.speak("Entered value is greater than current avialable stock ")
            print("Entered value is greater than current avialable stock ")
        else:
            total=q*100
            speak.speak(f"Total Rental Price is: {total}")
            print("Total Rental Price is: ", total)
            speak.speak("Press y for purchasing")
            r = input ("Press y for purchasing")
            if r=="y":
                self.stock=self.stock-q
                speak.speak(f"Thank You, You can pick {q} bikesfrom the store")
                print(f"you can pick {q} from the store")
try:
    obj=BikeShop(Bstock)
except:
    ("invlaid entry")

speak.speak("please select The Desired Option")
while True:
    try:
        user=int(input(
            """
            Select The Desired Option
            
                1 For Display Stock
                2 For Rent a Bike
                3 For Exit
                
            """))


        if user==1:
            obj.displayBike()

        elif user==2:
            speak.speak(("Enter the quantity of required bikes: "))
            try:
                n= int(input("Enter the quantity of required bikes: "))
            except:
                print("enter valid Entry")
            obj.bikeForRent(n)
        else:
            speak.speak("Bye, See You Again")
            break

    except:
        print("Enter Valid Entry")
