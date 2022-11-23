from enum import Enum

class Color(Enum):
    WHITE = 'grey'
    DARK = 'dark'
    BLACK = 'black'
    DEFAULT = ''


class Product(Enum):
    JERSEY = 'Ohiko'
    JERSEY_LONG = 'Ohiko_LongSleeves'
    TANK = 'Tank'
    SHORT = 'Short'
    HOODIE_NO_ZIP = 'Argia'
    HOODIE_ZIP = 'Kanpaia'
    SNOOD = 'Snood'
    HEADBAND = 'Headband'
    GLOVES = 'Gloves'
    TIGHT_HALF = 'Half Leg'
    TIGHT_KORSAIR = 'Korsair'
    TIGHT_KORSAIR_PLUS = 'Korsair Plus'



prices_adult = {
    Product.JERSEY: 34,
    Product.JERSEY_LONG: 38,
    Product.TANK: 29,
    Product.SHORT: 22,
    Product.HOODIE_NO_ZIP: 42,
    Product.HOODIE_ZIP: 49,
    Product.SNOOD: 10,
    Product.HEADBAND: 4,
    Product.GLOVES: 26,
    Product.TIGHT_HALF: 22,
    Product.TIGHT_KORSAIR: 35,
    Product.TIGHT_KORSAIR_PLUS: 45,
}

prices_kid = {Product.JERSEY: 26, Product.JERSEY_LONG: 31, Product.TANK: 23, Product.SHORT: 20}
type2excel = {
    'Männer': 'Man',
    'Frauen': 'Woman',
    'Short unisex multisport': 'Multisport',
    'Shorts unisex long': 'Long',
    'Woman tight': 'Tight_Woman',
    'Kinder': 'Kid'
}
size2excel = {'Sechsjährige': '6', 'Achtjährige': '8', 'Zehnjährige': '10'}

