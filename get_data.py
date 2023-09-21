import pokemontcgsdk
from pokemontcgsdk import Card, RestClient
import pandas as pd

# Configure the API key
RestClient.configure('#')

# Fetch all cards
cards_data = []
cards = Card.all()

# Collecting card details
for card in cards:
    card_dict = card.__dict__
    cards_data.append(card_dict)

def save_to_excel(data, filename="pokemon_card_data2.xlsx"):
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)


save_to_excel(cards_data)
