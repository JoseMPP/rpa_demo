from price_parser import Price


def get_price_and_currency(price_value):
    price = Price.fromstring(price_value.strip())
    return [price.amount_float,price.currency]
