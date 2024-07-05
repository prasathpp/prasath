

import random

print("Welcome to Cows and Bulls Game")
print("It is a Two Digit Number Guessing Game")
print("Total Number of Chances for the Player is 7")

secretnumber = str(random.randint(10,99))
chances = 7

while chances<=7:
    playerguess = input("Enter your Guess")

    if secretnumber == playerguess:
        print("Yes! You Guessed it Correct")
        print("Congrats!!!!!")
        break
    else:
        cows = 0
        bulls = 0

        if secretnumber[0] == playerguess[0]:
            bulls+=1
        if secretnumber[1] == playerguess[1]:
            bulls+=1
        if secretnumber[0] == playerguess[1]:
            cows+=1
        if secretnumber[1] == playerguess[0]:
            cows+=1

        print("Bulls : ",bulls)
        print("Cows : ",cows)

        chances-=1

        if chances<1:
            print("Sorry All your Attempts are Incorrect")
            print("Game Over")
            print("The Secret Number is ",secretnumber)
            break

        
        




