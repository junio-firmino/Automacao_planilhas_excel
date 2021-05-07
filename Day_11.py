import random

def deal_card():
    '''Function choicing the cards in amount.'''
    cards = [11, 2, 3, 4, 5, 6, 7, 8, 9, 10, 10, 10]
    card = random.choice(cards)
    return card

user_cards = []
computer_cards = []
is_game_over = True

for i in range(2):
    user_cards.append(deal_card())
    computer_cards.append(deal_card())

def calculate_score(cards_1):  #
    '''Function calculate the score from the cards chosen.'''
    if sum(cards_1) == 21 and len(cards_1) == 2:
        return 0
    if 11 in cards_1 and sum(cards_1)>21:
        cards_1.remove(11)
        cards_1.append(1)

    return sum(cards_1)

def compare(user_score, computer_score):
    if user_score == computer_score:
        return 'Draw'
    elif computer_score == 0:
        return 'Lose, opponente has Blackjack.'
    elif user_score == 0:
        return 'You win'
    elif user_score>21 and computer_score>21:
        if user_score>computer_score:
            return 'You win because you have more score in this cenarious.'
        else:
            return 'You have less score with the score total over so YOU LOSE!!'
    else:
        return 'You lose'


while is_game_over:
    user_score = calculate_score(user_cards)
    computer_score = calculate_score(computer_cards)
    print(f'Your cards: {user_cards}, Current user cards: {user_score}')
    print(f'Computer first card is {computer_cards} and Current computer cards is {computer_score}')

    if user_score == 0 or computer_score == 0 or user_score>21 or computer_score>17:
        is_game_over = False
        print(compare(user_score,computer_score))
    else:
        user_shoul_deal = input('Choise "y" to get another card, type "n" to pass: ')
        if user_shoul_deal == 'y':
            user_cards.append(deal_card())
            computer_cards.append(deal_card())
        else:
            is_game_over = False
            print(compare(user_score,computer_score))
    # if computer_score != 0 or computer_score < 17:
    #     computer_score = calculate_score(computer_cards)
    #





