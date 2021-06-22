import pandas as pd
from enum import Enum
from dataclasses import dataclass
from typing import Optional
import os
import wget


def main():
    df = download_orders()
    df = merge_mutual_exclusive_cols(df)
    players_df = download_player_infos()
    order_df = pd.DataFrame(columns='Full name,Product,Gender / Type,Size,Name to be printed,Number to be printed,Special Requests,Price,Is kid,Description'.split(','))
    for _, row in df.iterrows():
        name, items = extract_items(row)
        jersey_name, jersey_number = get_player_info(players_df, name)
        for item in items:
            item.full_name = name
            if item.product not in [Product.HOODIE_NO_ZIP.value, Product.HOODIE_ZIP.value]:
                item.jersey_number = jersey_number
                if item.product != Product.SHORT.value:
                    item.jersey_name = jersey_name
            order_df = order_df.append(item.to_series(), ignore_index=True)
    # print(order_df.to_string())
    prices_df = calculate_prices(order_df, players_df)
    total_price = prices_df['price'].sum()
    print(f'TOTAL PRICE: {total_price}€')
    writer = pd.ExcelWriter('summary.xlsx', engine='openpyxl')
    prices_df.to_excel(writer, index=False)
    writer.save()
    write_order_to_wb(order_df)


def download_orders():
    return download_google_sheet_as_df(
        '1xOnqs-DSKLaIjb3bCxPzd-EgMSz3j9KQ2tJyan5M5eQ',
        'formularantworten.csv'
    )


def download_player_infos():
    players_df = download_google_sheet_as_df(
        '18faU7kEJMGFY7Orl8MvtOYUU21OUFskwfsw3fos85Ww',
        'players.csv'
    )
    players_df = players_df.rename(columns={
        'Vollständiger Name': 'name',
        'Rückennummer': 'number',
        'Name auf Trikot': 'jersey_name'
    })
    # print(players_df.to_string())
    return players_df


def download_google_sheet_as_df(id, filename='temp.csv'):
    if os.path.isfile(filename):
        os.remove(filename)
    wget.download(f'https://docs.google.com/spreadsheets/d/{id}/export?format=csv', out=filename)
    df = pd.read_csv(filename)
    return df


def calculate_prices(df, players_df):
    prices_df = pd.DataFrame(columns='name price num_full_kits summary'.split())
    for name, items in df.groupby('Full name'):
        num_full_kits = calc_num_full_kits(items)
        price = items['Price'].sum() - 10 * num_full_kits
        summary = summarize_order(items)
        prices_df = prices_df.append({'name': name, 'price': price, 'num_full_kits': num_full_kits, 'summary': summary}, ignore_index=True)
        summary_f = ' - ' + summary.replace(', ', '\n - ')
        jersey_name, jersey_number = get_player_info(players_df, name)
        print(f'{name}:\njersey_name: {jersey_name}\njersey_number: {jersey_number}\n{summary_f}\n num full kits: {num_full_kits}\n total price: {price}€\n')
    return prices_df


def summarize_order(items: pd.DataFrame):
    summary = ''
    counts = items['Description'].value_counts()
    for description, count in counts.iteritems():
        prefix = f'{count}X ' if count > 1 else ''
        summary += prefix + description + ', '
    summary = summary[:-2]  # remove trailing comma
    return summary


def calc_num_full_kits(items: pd.DataFrame):
    jerseys = items[items['Product'].isin([Product.JERSEY.value, Product.JERSEY_LONG.value])]
    num_dark_jerseys = len(jerseys[jerseys['Special Requests'] == Color.DARK.value])
    num_light_jerseys = len(jerseys[jerseys['Special Requests'] == Color.LIGHT.value])
    num_shorts = len(items[items['Product'] == Product.SHORT.value])
    num_kits = min(num_shorts, num_light_jerseys, num_dark_jerseys)
    return num_kits


def write_order_to_wb(order_df: pd.DataFrame):
    sheet_name = 'My order'
    filename = 'orderform_template.xlsx'
    writer = pd.ExcelWriter(filename, engine='openpyxl', mode='a')
    order_df.to_excel(writer, sheet_name, index=False, header=False)
    writer.save()


def get_player_info(players_df, name):
    default_number = ''
    default_name = ''
    df = players_df[players_df['name'] == name]
    if not len(df):
        print(f'could not find player {name}')
        return default_name, default_number
    else:
        jersey_name = df['jersey_name'].values[0]
        number = str(int(df['number'].values[0]))
        return jersey_name, number


def merge_mutual_exclusive_cols(df):
    cols = [col for col in df.columns if col.endswith('.1')]
    similar_cols = [col[:-2] for col in cols]
    for c1, c2 in zip(similar_cols, cols):
        df[c1].fillna(df[c2], inplace=True)
        df.drop(c2, axis=1, inplace=True)
    df[Col.SIZE_JERSEY.value].fillna(df[Col.SIZE_JERSEY_KIDS.value], inplace=True)
    df.drop(Col.SIZE_JERSEY_KIDS.value, axis=1, inplace=True)
    return df


class Product(Enum):
    JERSEY = 'Ohiko'
    JERSEY_LONG = 'Ohiko_LongSleeves'
    TANK = 'Tank'
    SHORT = 'Short'
    HOODIE_NO_ZIP = 'Argia'
    HOODIE_ZIP = 'Kanpaia'


prices_adult = {
    Product.JERSEY: 34,
    Product.JERSEY_LONG: 38,
    Product.TANK: 29,
    Product.SHORT: 22,
    Product.HOODIE_NO_ZIP: 42,
    Product.HOODIE_ZIP: 49
}

prices_kid = {
    Product.JERSEY: 26,
    Product.JERSEY_LONG: 31,
    Product.TANK: 23,
    Product.SHORT: 20
}


class Color(Enum):
    LIGHT = 'grey'
    DARK = 'blue'
    BLACK = 'black'
    DEFAULT = ''


@dataclass
class Item:
    product: Product
    type_: str
    size: str
    color: Color
    jersey_name: Optional[str]
    jersey_number: Optional[int]
    price: int = -1
    is_kid: bool = False
    full_name: str = ''

    def __post_init__(self):
        self.size = self.size if self.size not in size2excel.keys() else size2excel[self.size]
        self.type_ = self.type_ if self.type_ not in type2excel.keys() else type2excel[self.type_]
        self.is_kid = self.size in '6 8 10'.split()
        self.price = (prices_kid if self.is_kid else prices_adult)[self.product]

    def to_series(self):
        return pd.Series({
            'Full name': self.full_name,
            'Product': self.product.value,
            'Gender / Type': self.type_,
            'Size': self.size,
            'Name to be printed': self.jersey_name,
            'Number to be printed': self.jersey_number,
            'Special Requests': self.color.value,
            'Price': self.price,
            'Is kid': self.is_kid,
            'Description': self.to_string()
        })

    def to_string(self):
        kid_str = ' (kid)' if self.is_kid else ''
        color_str = ' ' + self.color.name.lower() if not self.color == Color.DEFAULT else ''
        return f'{self.product.name.lower()}{kid_str} {self.type_.lower()}{color_str} {self.size} ({self.price}€)'


type2excel = {
        'Männer': 'Man',
        'Frauen': 'Woman',
        'Short unisex multisport': 'Multisport',
        'Shorts unisex long': 'Long',
        'Woman tight': 'Tight_Woman',
        'Kinder': 'Kid'
    }
size2excel = {
        'Sechsjährige': '6',
        'Achtjährige': '8',
        'Zehnjährige': '10',
    }


class Col(Enum):
    NAME = 'Vollständiger Name (wie in Rückennummer-Tabelle)'
    NUM_DARK = 'Anzahl Trikots [Trikot (kurz) dunkel]'
    NUM_LIGHT = 'Anzahl Trikots [Trikot (kurz) hell]'
    NUM_DARK_LONG = 'Anzahl Trikots [Longsleeve dunkel]'
    NUM_LIGHT_LONG = 'Anzahl Trikots [Longsleeve hell]'
    NUM_BLACK_LONG = 'Anzahl Trikots [Longsleeve schwarz (inoffiziell)]'
    NUM_DARK_TANK = 'Anzahl Trikots [Tank Top dunkel]'
    NUM_LIGHT_TANK = 'Anzahl Trikots [Tank Top hell]'
    TYPE_JERSEY = 'Schnitt (Trikot)'
    SIZE_JERSEY = 'Größe (Trikot)'
    SIZE_JERSEY_KIDS = 'Größe (Kindertrikot)'
    NUM_SHORTS = 'Anzahl (Shorts)'
    TYPE_SHORTS = 'Schnitt (Shorts)'
    SIZE_SHORTS = 'Größe (Shorts)'
    NUM_HOODIES_NO_ZIP = 'Anzahl (Hoodie ohne Reißverschluss)'
    TYPE_HOODIES_NO_ZIP = 'Schnitt (Hoodie ohne Reißverschluss)'
    SIZE_HOODIES_NO_ZIP = 'Größe (Hoodie ohne Reißverschluss)'
    NUM_HOODIES_ZIP = 'Anzahl (Hoodie mit Reißverschluss)'
    TYPE_HOODIES_ZIP = 'Schnitt (Hoodie mit Reißverschluss)'
    SIZE_HOODIES_ZIP = 'Größe (Hoodie mit Reißverschluss)'
    COMMENTS = 'Sonstige Anmerkungen'


def extract_items(row: pd.Series):
    name = row[Col.NAME.value]
    order = []
    for num, prod, color in [
        (Col.NUM_DARK, Product.JERSEY, Color.DARK),
        (Col.NUM_LIGHT, Product.JERSEY, Color.LIGHT),
        (Col.NUM_DARK_LONG, Product.JERSEY_LONG, Color.DARK),
        (Col.NUM_LIGHT_LONG, Product.JERSEY_LONG, Color.LIGHT),
        (Col.NUM_DARK_TANK, Product.TANK, Color.DARK),
        (Col.NUM_LIGHT_TANK, Product.TANK, Color.LIGHT),
    ]:
        order += get_similar_items(row, num, prod, Col.TYPE_JERSEY, Col.SIZE_JERSEY, color)
    order += get_similar_items(row, Col.NUM_SHORTS, Product.SHORT, Col.TYPE_SHORTS, Col.SIZE_SHORTS, Color.DEFAULT)
    order += get_similar_items(row, Col.NUM_HOODIES_NO_ZIP, Product.HOODIE_NO_ZIP, Col.TYPE_HOODIES_NO_ZIP, Col.SIZE_HOODIES_NO_ZIP, Color.DEFAULT)
    order += get_similar_items(row, Col.NUM_HOODIES_ZIP, Product.HOODIE_ZIP, Col.TYPE_HOODIES_ZIP, Col.SIZE_HOODIES_ZIP, Color.DEFAULT)
    return name, order


def get_similar_items(row, num: Col, product: Product, type_: Col, size: Col, color=None):
    if num:
        item = Item(
            product,
            row[type_.value],
            row[size.value],
            color,
            jersey_name='',
            jersey_number='')
        return [item] * row[num.value]
    else:
        return []


if __name__ == '__main__':
    main()