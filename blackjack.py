from time import sleep
from random import shuffle, choice

delay = 1.8 # animation delay

# cards 

class Deck(object):
    card_names = ['A', '2', '3', '4', '5', '6', '7','8', '9', '10', 'J', 'Q', 'K']

    def __init__(self):
        self.cards = range(52)
        random.shuffle(self.cards)
        self.cards.reverse()

    def deal(self):
        card = self.cards.pop()
        card_name = self.card_names[card % 13]
        if card < 26:
            card_color = 'red'
        else:
            card_color = 'black'
        return (card_name, card_color)

def card_val(card):
    val = Deck.card_names.index(card) + 1
    if val >= 10:
        return 10
    return val

# game flow

def play_round():
    bet()
    # deal initial cards
    hit(5)
    sleep(delay)
    hit(3)
    sleep(delay)
    hit(5)
    sleep(delay)
    hit(3, True)
    # play
    outcome = player()
    if not outcome:
        outcome = dealer()
    if not outcome:
        dealer_hand = abs(hand_total(3))
        player_hand = abs(hand_total(5))
        # abs to correct for negative val in case of Ace
        if player_hand > dealer_hand:
            dealer_says('You win!')
            outcome = 'win'
        elif dealer_hand == player_hand:
            dealer_says('Push!')
            outcome = 'push'
        else:
            dealer_says('You lose!')
            outcome = 'lose'
    return outcome

def flip_card():
    Cell('F3').value = '|'
    sleep(0.3 * delay)
    Cell('F3').value = hole_card[0]
    Cell('F3').font.color = hole_card[1]

def hand_total(row):
    col = 2
    hand = CellRange((row, 5), (row, 25)).value
    hand = filter(None, hand) # remove blank cells
    hand = map((lambda x: str(x)), hand) # change numbers to strings
    if 'X' in hand: # hole card
        hand[1] = hole_card[0]
    hand = map(card_val, hand)
    if (len(hand) == 2) and (10 in hand) and (1 in hand):
        return 'blackjack'
    total = sum(hand)
    if (1 in hand) and (total + 10 <= 21):
        # count one ace as 11
        return -(total + 10) # negative to account for dealer's soft 17
    return total

def dealer_says(txt): Cell('D1').value = txt

## player actions
def bet(talk = True):
    if talk:
        dealer_says("What's your bet?")
    bet_cell = Cell('E6')
    bet_cell.clear()
    active_cell(bet_cell)
    bet_cell.value = 'enter bet'
    bet_cell.font.color = 'gray'
    bet_amt = get_cell_value(bet_cell)
    bet_cell.font.color = 'black'
    while type(bet_amt) not in (int, long):
        try:
            bet_amt = int(bet_amt)
        except ValueError:
            dealer_says("Sorry, I didn't understand that.  How much do you want to bet?")
            bet(False)
            return
    bet_cell.value = '$' + str(bet_amt)
    if bet_amt < 100:
        dealer_says('$100 minimum, please.')
        bet(False)
    elif bet_amt > chips.value:
        dealer_says("It looks like you don't have enough for that.")
        bet(False)
    else:
        Cell('B3').value -= bet_amt
        sleep(0.5 * delay)
        dealer_says('')
    
def player():
    # possible outcomes - blackjack, d_blackjack, blackjack_push, bust, surrender
    actions = ['Hit', 'Stand', 'Double', 'Surrender']
    if Cell('E3').value == 'A':
        actions.insert(3, 'Insurance')
    act = print_and_get_action(actions)
    if act == 'Surrender':
        return 'surrender'
    elif act == 'Insurance':
        insurance()

    # dealer checks for blackjack
    if hand_total(3) == 'blackjack':
        flip_card()
        if hand_total(5) == 'blackjack':
            dealer_says("Two blackjacks - it's a push!")
            return 'blackjack_push'
        else:
            dealer_says("Dealer has blackjack!")
            return 'd_blackjack'

    # player blackjack
    if hand_total(5) == 'blackjack':
        dealer_says('Blackjack - congratulations! Pay out is phi : 1.')
        flip_card()
        return 'blackjack'

    # gameplay
    actions = actions[:3]
    if act not in actions:
        # insurance or surrender
        act = print_and_get_action(actions)

    while act == 'Hit':
        hit(5)
        sleep(0.5 * delay)
        if hand_total(5) > 21:
            dealer_says('Bust!')
            return 'bust'
        act = print_and_get_action(actions)

    if act == 'Double':
        double()
        if hand_total(5) > 21:
            dealer_says('Bust!')
            return 'bust'

    elif act == 'Stand':
        return

def print_and_get_action(actions):
    last_col = 4 + len(actions)
    CellRange((7, 5), (7, last_col)).value = actions
    CellRange((7, 5), (7, last_col)).font.bold = True
    active_cell((7,4))
    row, col = active_cell().position
    while row!= 7 or col not in range(5, last_col + 1):
        sleep(delay)
        row, col = active_cell().position
    CellRange((7, 5), (7, last_col)).clear()
    return actions[col - 5]

def hit(row, hole = False):
    col = 4
    while not Cell(row, col).is_empty():
        col += 1
    card, color = deck.deal()
    if hole:
        global hole_card
        hole_card = (card, color)
        card, color = 'X', 'black'
    Cell(row, col).value = card
    Cell(row, col).font.color = color

def double():
    bet_cell = Cell('E6')
    chips.value -= bet_cell.value
    bet_cell.value *= 2
    hit(5)

def insurance():
    Cell('D4').value = 'Insured:'
    amt = int(Cell('E6').value/2)
    Cell('E4').value = '$' + str(amt)
    chips.value -= amt

## dealer
def dealer():
    # possible outcomes: dealer bust
    flip_card()
    while hand_total(3) < 17:
        # soft 17 will be negative
        sleep(1.5 * delay)
        hit(3)
    if hand_total(3) > 21:
        sleep(delay)
        dealer_says('Dealer busted - you win!')
        return 'd_bust'

def get_cell_value(cell):
    val = cell.value
    while cell.value == val:
        sleep(delay)
    return cell.value

def setup():
    clear_sheet()
    merge_cells('D1', 'AA1')
    Cell('D1').font.bold = True

    CellRange('A3:A4').value = ['Chips:', 'Casino:']
    chips.font.color = 'green'
    chips.value, bank.value = '$1000', '$20,000'

    CellRange('D3, D5, D6').value = ['Dealer:','Hand:', 'Bet:']
    CellRange('A6:B6').value = ['Round', 'Result']
    CellRange('A6:B6').font.underline = True

def result(outcome):
    # outcomes: d_blackjack, blackjack, blackjack_push, d_bust, bust, win, lose, push, surrender
    # deal with insurance
    bet = Cell('E6').value
    phi = 1.6180339887
    amt = 0

    if not Cell('D4').is_empty():
        # insurance
        insurance = Cell('E4').value
        if type(insurance) == str:
            insurance = int(insurance[1:])
        if outcome in ['d_blackjack', 'blackjack_push']:
            # insurance pays 2 to 1
            amt += 2 * insurance
        else:
            amt -= insurance

    # deal with outcome
    if outcome == 'blackjack':
        amt += int(phi * bet)
    elif outcome in ['d_bust', 'win']:
        amt += bet
    elif outcome in ['push', 'blackjack_push']:
        pass
    elif outcome == 'surrender':
        amt -= int(bet / 2)
    elif outcome in ['d_blackjack', 'bust', 'lose']:
        amt -= bet
    return amt


def winnings(amt):
    # amt is the amount won or lost
    amt = int(amt)
    bet = Cell('E6').value
    Cell('E6').clear()
    insurance = Cell('E4').value
    if insurance == None:
        insurance = 0
    elif type(insurance) == str:
        insurance = int(insurance[1:])
    CellRange('D4:E4').clear()
    chips.value += amt + bet + insurance
    bank.value -= amt
    Cell((6 + round), 1).value = round
    Cell((6 + round), 2).value = '$' + str(amt)
    if amt > 0:
        Cell((6 + round), 2).font.color = 'green'
    if amt < 0:
        Cell((6 + round), 2).font.color = 'red'
    if chips.value == 0:
        chips.font.color = 'black'
    elif chips.value <= 0:
        chips.font.color = 'red'
    if bank.value < 0:
        bank.font.color = 'red'

def cleanup():
    CellRange('E3:Z3,E5:Z5').clear()


########### main

# setup


chips = Cell('B3')
bank = Cell('B4')
round = 0
setup()

dealer_says('Welcome to the IronSpread blackjack table!')
sleep(delay)

# game

while 1:
    round += 1
    deck = Deck()
    outcome = play_round()
    amt = result(outcome)
    winnings(amt)
    sleep(2 * delay)

    # check for end conditions
    if chips.value < 0:
        tasks = ['setting the time on my VCR', 
                 'cleaning up this CSV file', 
                 'dealing at the Spreadsheet Blackjack world championship', 
                 'installing my router',
                 'rewriting our VBA code in Python']
        task = choice(tasks)
        dealer_says("You don't have enough money - you can pay your debt by %s!" %task)
        break
    elif chips.value < 100:
        dealer_says("It looks like you're out of money. Come back again soon!")
        break
    elif bank.value <= 0:
        dealer_says('I just opened this place, and you won all my money!')
        sleep(2 * delay)
        dealer_says("Now I have to go back to working as a VBA programmer!")
        break

    cleanup()
