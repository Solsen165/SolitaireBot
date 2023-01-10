# SolitaireBot

The bot first loads all card images.

Then it moves through each stack on the screen and identifies each card, then it starts the game algorithm.

## Algorithm:
It first checks if there are any cards that can be put in the piles, (the cards are put in the piles in order of their value)  
after that it checks if there are any available stacks that can be put together,  
when it finds that there are no longer any valid moves, it checks the stock card.

The bot repeats this algorithm until all piles are complete.  
If it gets stuck, it will try to put cards in the piles out of order to try to reveal new cards.

### Note:
The bot doesn't win each time, because the bot doesn't have any deep strategy and not all solitare games are winnable to begin with.
