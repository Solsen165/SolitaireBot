import pygetwindow as gw
import pyautogui as pg
from PIL import Image
import time
import keyboard
import win32api
import win32con
import os
import cv2
import numpy as np
import pywinauto

# class for the playing cards


class Card:
    def __init__(self, value, suit='', img=None):
        self.value = value
        self.suit = suit                    # suit of the card
        # position of the card on the screen
        self.pos = (0, 0)
        self.pre = None                     # the card which is underneath it
        if img != None:
            self.img = Image.open(img)      # its image
        if (suit == 'Diamond' or suit == 'Heart'):
            self.color = 'Red'
        else:
            self.color = 'Black'

    # function for printing the name of the card
    def detail(self):
        name = ''
        if self.value == 11:
            name = 'Jack'
        elif self.value == 12:
            name = 'Queen'
        elif self.value == 13:
            name = 'King'
        elif self.value == 1:
            name = 'Ace'
        else:
            name = str(self.value)
        print(name, 'of', self.suit + 's')

    # function which returns whether it is possible to place the card on top of another card
    def putOn(self, card2):
        if self.value == 13 and card2.value == 14:
            return self.pre != 14
        elif self.value == card2.value - 1 and self.color != card2.color:
            return True
        else:
            return False


# class for each column (stack) in the table
class Column:
    def __init__(self, bottom, top):
        self.bottom = bottom            # Card on bottom of the stack
        self.top = top                  # Card on top of the stack

    # function for removing the top card
    def remTop(self):
        if self.top == self.bottom:
            self.remBottom()
        else:
            oldpos = self.top.pos
            self.top = self.top.pre
            self.top.pos = (oldpos[0], oldpos[1]-20)

    # function for removing the whole stack
    def remBottom(self):
        oldpos = self.bottom.pos
        self.bottom = self.bottom.pre
        if self.bottom == None:
            click(oldpos)
            self.bottom = identifyCard(takePhoto(oldpos[0], oldpos[1], 0))
        if self.bottom == None:
            self.bottom = Card(14)
        self.bottom.pos = (oldpos[0], oldpos[1]-5)

        self.top = self.bottom

    # function for adding a card to the stack
    def add(self, newBottom, newTop=None):
        if newTop == None:
            newTop = newBottom
        height = newTop.pos[1] - newBottom.pos[1]
        newpos = (self.bottom.pos[0], self.top.pos[1] + height + 20)

        newBottom.pre = self.top
        self.top = newTop
        self.top.pos = newpos


# List containing all 52 cards.
# Used for identifying the cards with their images
deck = []

# Function for populating the deck list with all cards


def initDeck():
    global deck
    # Path for folder containing card images
    path = "GrayCards/"
    for filename in os.listdir(path):
        # Each image is named with the value and suit (For example: 03-Club.png)
        v = int(filename[0:2])
        s = filename[3:-4]
        img = path + '\\' + filename
        newCard = Card(v, s, img)
        deck.append(newCard)


# Function for taking screenshot of a card on the screen
def takePhoto(x, y, top) -> Image:
    # (x,y) are a position on the screen where the card is located

    # limits for card dimensions
    u = y
    d = y
    l = x
    r = x

    # Top means that the card is on the stock at the top of the screen so we take the upper edge and
    # the right edge because they are empty and don't contain other cards to avoid confusion
    if top:
        while pg.pixel(x, u) != (0, 166, 81) and pg.pixel(x, u) != (0, 128, 0):
            u -= 5

        while pg.pixel(r, y) != (0, 128, 0):
            r += 5
        l = r-130

    # otherwise the cards are on the table so we take the lower edge and left edge (right edge is also possible)
    else:
        while pg.pixel(l, y) != (0, 128, 0):
            l -= 5

        while pg.pixel(x, d) != (0, 128, 0):
            d += 5
        u = d-130

    # The size for each card is (95,130)
    im = pg.screenshot(region=(l, u, 95, 130))

    return im


# Function takes the card screenshot and identifies it from the deck list
def identifyCard(im):

    # number of possible matches
    found = 2
    # correct card
    ansCard = 0
    # confidence variable, the higher it is the more accurate but more likely to not recognise it
    conf = 0.8

    # as long as there are more than one matches we will try again but with more confidence
    while found > 1:
        found = 0
        conf += 0.01
        for card in deck:
            if pg.locate(card.img, im, grayscale=True, confidence=conf) != None:
                found += 1
                ansCard = card

    if (ansCard == 0):
        return None
    return ansCard


# Function for setting up the table list from the cards on the screen to be able to perform the algorithm on them
def setTable():
    global table
    # starting position for leftmost stack
    x = 220
    y = 360
    for i in range(7):
        card = identifyCard(takePhoto(x, y, 0))
        card.pos = (x, y)
        table[i].top = card
        if table[i].bottom == None:
            table[i].bottom = card

        # Each stack is 103 pixels away from previous one, and the card is 4 pixels down from previous one
        x += 103
        y += 4


# Function for clicking on the screen with the cursor
def click(pos):
    win32api.SetCursorPos(pos)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.01)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)


# Function for dragging a card from one position to another
def drag(pos1, pos2):
    pg.moveTo(pos1)
    pg.mouseDown()
    pg.moveTo(pos2[0], pos2[1]+10)
    time.sleep(0.01)
    pg.mouseUp()

# Function for displaying the table


def displayTable():
    for c in table:
        c.top.detail()


# Getting the solitaire game window, and activating it
# If the program doesn't open the game window try changing the index to 0, because the function returns list of all windows whose titles contain the word
s = gw.getWindowsWithTitle("Solitaire")[1]

s.restore()
s.activate()

# move window to starting position
s.moveTo(156, 156)

# Initialising the deck
initDeck()

# List for cards on table
table = []

for i in range(7):
    table.append(Column(None, None))
setTable()

# We reverse the table because it is more strategic to begin from the right to remove cards from the biggest stacks first
table.reverse()
print('Table Set up!')

# List for cards on top of piles
piles = []

suits = ['Spade', 'Heart', 'Club', 'Diamond']
piley = 280
pilex = 529
for i in range(4):
    # Each pile starts off empty so we initialise it with card with value 0, so the next one is value 1
    pilecard = Card(0, suits[i])

    # Setting up each pile position
    pilecard.pos = (pilex+20, piley-20)
    piles.append(pilecard)
    pilex += 103


# Variable for next value supposed to be put in the pile
pileTop = 1

# Counter for number of piles finished with current pileTop
# When counter reaches 4, that means that we have to move to the next value
noOfPiles = 0

# If the bot gets stuck, so if it emptys the stock more the 3 times, that means that we have to ignore pileTop and start
# placing available cards on the pile so we can free some other cards
stuck = 0

# Loop for playing the game
while 1:

    # When the piles are finished the game ends
    if pileTop == 14:
        print("The Bot Wins!!")
        break

    # For stopping the game in case of infinite loop, or unwinnable game
    if keyboard.is_pressed('q'):
        break

    # When the bot is stuck we will increase the pileTop to start placing available cards on the piles
    # regardless of whether the order is preserved or not
    if stuck == 3:
        stuck = 0
        pileTop += 1
        noOfPiles = 0

    # Click on stock, to check available card from the stock
    click((220, 280))

    # Taking screenshot of the stock card, and identifying it
    stockImg = pg.screenshot(region=(282, 224, 95, 130))
    stock = identifyCard(stockImg)

    # if the stock is empty or we reached the end of it, we will increase the stuck variable
    if stock == None:
        stuck += 1

    # Boolean for checking if the table is changing or if we need to check next card from the stock
    changed = True

    while (changed):  # as long as the table changes then we will keep cycling through it
        changed = False
        for col1 in table:

            # 14 is value for empty space, so we skip it
            if col1.top.value == 14:
                continue

            # Checking whether we can place the top card in one of the piles
            if col1.top.value <= pileTop:
                for p in piles:
                    if col1.top.suit == p.suit and col1.top.value == p.value+1:

                        # Printing current move
                        col1.top.detail()
                        print('     |\n     v')
                        print(p.suit, 'Pile')
                        print(' ')

                        drag(col1.top.pos, p.pos)
                        p.value = col1.top.value

                        changed = True
                        stuck = 0
                        col1.remTop()
                        noOfPiles += 1
                        break

                # If the top card's value is less then pileTop, then we will negate the increase in noOfPiles
                # so we don't mess up the ordering
                if col1.top.value < pileTop:
                    noOfPiles -= 1

                # if piles are finished we will move to the next number
                if noOfPiles == 4:
                    noOfPiles = 0
                    pileTop += 1

            # If empty space after placing on pile, then we will skip
            if col1.top.value == 14:
                continue

            # Checking if we can put this stack on another stack on table
            for col2 in table:
                if col1.bottom.putOn(col2.top):

                    # Printing current move
                    col1.bottom.detail()
                    print('     |\n     v')
                    if col2.top.value == 14:
                        print("Empty Space")
                    else:
                        col2.top.detail()
                    print(' ')

                    drag(col1.bottom.pos, col2.top.pos)

                    changed = True
                    stuck = 0

                    # Handling the new positions for col1, and col2
                    oldBottom = col1.bottom
                    oldTop = col1.top
                    col1.remBottom()
                    col2.add(oldBottom, oldTop)

    # Same algorithm but for stock card

    # If stock is empty or we reached the end then there is no card to check, so skip
    if stock != None:
        # Position of stock card
        stock.pos = ((323, 280))

        # Checking if stock card can be put in the piles
        if stock.value <= pileTop:
            for p in piles:
                if stock.suit == p.suit and stock.value == p.value + 1:

                    stuck = 0

                    # Printing current move
                    stock.detail()
                    print('     |\n     v')
                    print(p.suit, 'Pile')
                    print(' ')

                    drag(stock.pos, p.pos)
                    p.value = stock.value

                    noOfPiles += 1
                    break

            if stock.value < pileTop:
                noOfPiles -= 1

            if noOfPiles == 4:
                noOfPiles = 0
                pileTop += 1
            stock = None

    if stock != None:
        for col in table:
            if stock.putOn(col.top):
                stuck = 0

                # Printing current move
                stock.detail()
                print('     |\n     v')
                if col.top.value == 14:
                    print("Empty Space")
                else:
                    col.top.detail()
                print(' ')

                drag(stock.pos, col.top.pos)
                col.add(stock)
                break
