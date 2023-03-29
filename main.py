import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
import pandas as pd
from dataclasses import dataclass
from typing import Optional
import os
import wget
from datetime import datetime
from force_macros import *
import shutil

order_start = order_end = None


rotpot_cfg = dict(
    form_answers_id='1xOnqs-DSKLaIjb3bCxPzd-EgMSz3j9KQ2tJyan5M5eQ',
    player_infos_id='18faU7kEJMGFY7Orl8MvtOYUU21OUFskwfsw3fos85Ww',
    payment_info_id='1xOnqs-DSKLaIjb3bCxPzd-EgMSz3j9KQ2tJyan5M5eQ',
    payment_info_gid=1474649789,
    # start_date="23.09.2022",
    # end_date=None,
    prefix='rotpot',
    kids_design=False
)

kids_cfg = dict(
    form_answers_id='1JUlX7qOaaEINtQlzFd4yXUXie9H2EAdvSw9Q5AFao5Y',
    player_infos_id='1IaqBGPRLDBqFdpOT_YzM2o0eMKb8jv3gM97RKkLuzlk',
    payment_info_id='1JUlX7qOaaEINtQlzFd4yXUXie9H2EAdvSw9Q5AFao5Y',
    payment_info_gid=1320287528,
    # start_date="02.12.2021",
    # end_date="15.08.2022",
    prefix='kids',
    kids_design=True
)


def main():
    process_google_forms(**rotpot_cfg)
    process_google_forms(**kids_cfg)


def process_google_forms(form_answers_id, player_infos_id, payment_info_id, payment_info_gid, prefix, kids_design, start_date=None, end_date=None):
    """
    download, process, print and write the orders to excel
    """
    if prefix:
        print('\n' + '-' * 40 + '\n' + prefix.upper())

    df = download_google_sheet_as_df(form_answers_id, 'formularantworten.csv')
    df['Zeitstempel'] = df['Zeitstempel'].apply(parse_timestamp)
    if start_date is not None:
        day, month, year = [int(x) for x in start_date.split('.')]
        deadline = datetime(year=year, month=month, day=day)
        df = drop_order_older_than(df, deadline)
    if end_date is not None:
        day, month, year = [int(x) for x in end_date.split('.')]
        deadline = datetime(year=year, month=month, day=day)
        df = drop_order_later_than(df, deadline)
    df = merge_mutual_exclusive_cols(df)
    players_df = download_player_infos(player_infos_id)
    order_df = pd.DataFrame(
        columns='Full name,Product,Gender / Type,Size,Name to be printed,Number to be printed,Special Requests,Price,Is kid,Description'.split(','))
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
    payment_dict = download_payment_infos(payment_info_id, payment_info_gid)
    prices_df = calculate_prices(order_df, players_df, payment_dict, sort_by_name=False)
    total_price = prices_df['price'].sum()
    print(f'TOTAL PRICE: {total_price}€')
    writer = pd.ExcelWriter(f'downloaded_tables/summary_{prefix}.xlsx', engine='openpyxl')
    prices_df.to_excel(writer, index=False)
    writer.save()
    max_items_per_orderform = 200
    os.makedirs('generated_orders', exist_ok=True)
    if len(order_df) <= max_items_per_orderform:
        write_order_to_wb(order_df, prefix=prefix)
    else:
        print('order too large. Splitting in a-l and m-z by prename')
        mask = order_df['Full name'].apply(lambda x: x.lower()) < "m"
        write_order_to_wb(order_df[mask], suffix=' a-l', prefix=prefix)
        write_order_to_wb(order_df[~mask], suffix=' m-z', prefix=prefix)


def parse_timestamp(ts: str) -> datetime:
    return datetime.strptime(ts.split(' ')[0], '%d.%m.%Y')


def drop_order_older_than(df: pd.DataFrame, deadline: datetime):
    """
    filters the DataFrame df such that only rows with 'Zeitstempel' later than deadline remain
    """
    mask = df['Zeitstempel'] > deadline
    return df[mask]


def drop_order_later_than(df: pd.DataFrame, deadline: datetime):
    """
    filters the DataFrame df such that only rows with 'Zeitstempel' older than deadline remain
    """
    mask = df['Zeitstempel'] <= deadline
    return df[mask]



def download_payment_infos(id, gid) -> dict:
    """
    downloads information about which players did and did not pay their orders
    """
    summary = download_google_sheet_as_df(id, 'summary.csv', gid=gid)
    d = dict(zip(summary.Name, summary.Bezahlt))
    return d


def download_player_infos(id):
    players_df = download_google_sheet_as_df(id, 'players.csv')
    players_df = players_df.rename(columns={'Vollständiger Name': 'name', 'Rückennummer': 'number', 'Name auf Trikot': 'jersey_name'})
    return players_df


def download_google_sheet_as_df(id, filename='temp.csv', gid=None):
    filename = 'downloaded_tables/' + filename
    if os.path.isfile(filename):
        os.remove(filename)
    gid_str = f'gid={gid}&'if gid else ''
    wget.download(f'https://docs.google.com/spreadsheets/d/{id}/export?{gid_str}format=csv', out=filename)
    return pd.read_csv(filename)


def calculate_prices(df, players_df, payment_dict, sort_by_name=False):
    prices_df = pd.DataFrame(columns='paid name price num_full_kits summary'.split())
    for name, items in df.groupby('Full name', sort=sort_by_name):
        num_full_kits = calc_num_full_kits(items)
        price = items['Price'].sum() - 10 * num_full_kits
        summary = summarize_order(items)
        paid = payment_dict.get(name, 'Nein') == 'Ja'
        paid_en = 'yes' if paid else 'no'
        paid_de = 'Ja' if paid else 'Nein'
        prices_df = prices_df.append({'paid': paid_de, 'name': name, 'price': price, 'num_full_kits': num_full_kits, 'summary': summary}, ignore_index=True)
        summary_f = ' - ' + summary.replace(', ', '\n - ')
        jersey_name, jersey_number = get_player_info(players_df, name)
        num_items = len(items)
        # print(f'{name}:\njersey_name: {jersey_name}\njersey_number: {jersey_number}\n{summary_f}\n num full kits: {num_full_kits}\n total price: {price}€\n paid: {paid_en}\n num items: {num_items}\n')
        full_kits_str = '' if not num_full_kits else f'\n num full kits: {num_full_kits}'
        print(f'{name}:\njersey_name: {jersey_name}\njersey_number: {jersey_number}\n{summary_f}{full_kits_str}\n total price: {price}€\n num items: {num_items}\n')
    return prices_df


def summarize_order(items: pd.DataFrame):
    summary = ''
    counts = items['Description'].value_counts()
    for description, count in counts.iteritems():
        prefix = f'{count}X ' if count > 1 else ''
        summary += prefix + description + ', '
    return summary[:-2]  # remove trailing comma


def calc_num_full_kits(items: pd.DataFrame):
    jerseys = items[items['Product'].isin([Product.JERSEY.value, Product.JERSEY_LONG.value])]
    num_dark_jerseys = len(jerseys[jerseys['Special Requests'] == Color.DARK.value])
    num_light_jerseys = len(jerseys[jerseys['Special Requests'] == Color.WHITE.value])
    num_shorts = len(items[items['Product'] == Product.SHORT.value])
    num_kits = min(num_shorts, num_light_jerseys, num_dark_jerseys)
    return num_kits


def write_order_to_wb(order_df: pd.DataFrame, prefix='', suffix=''):
    sheet_name = 'rotpot_order'
    template_path = 'templates/orderform_template.xlsx'
    if prefix and not prefix.endswith('_'):
        prefix += '_'
    target_path = f'generated_orders/{prefix}orderform{suffix}.xlsx'
    shutil.copy(template_path, target_path)
    writer = pd.ExcelWriter(target_path, engine='openpyxl', mode='a')
    order_df.to_excel(writer, sheet_name, index=False, header=False)
    writer.save()


def get_player_info(players_df, name):
    default_number = default_name = '/'
    df = players_df[players_df['name'] == name]
    if not len(df):
        print(f'could not find player {name}')
        return default_name, default_number
    else:
        jersey_name = df['jersey_name'].values[0]
        number = df['number'].values[0]
        if number == 'keine Nummer':
            number = ''
        else:
            number = str(int(number))
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




@dataclass
class Item:
    product: Product
    type_: str
    size: str
    color: Color
    jersey_name: Optional[str]
    jersey_number: Optional[str]
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
        kid_str = ' kid' if self.is_kid else ''
        color_str = ' ' + self.color.name.lower() if not self.color == Color.DEFAULT else ''
        type_str = ' ' + self.type_.lower() if self.type_ not in [None, 'Kid'] else ''
        size_str = ' ' + self.size if self.size is not None else ''
        return f'{self.product.name.lower()}{color_str}{type_str}{kid_str}{size_str} ({self.price}€)'


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
    NUM_GLOVES = 'Anzahl Friction Gloves'
    TYPE_GLOVES = 'Schnitt Friction Gloves'
    SIZE_GLOVES = 'Größe Friction Gloves'
    NUM_HEADBANDS = 'Anzahl Headbands'
    NUM_SNOODS_WHITE = 'Anzahl Snoods Weiß'
    NUM_SNOODS_BLUE = 'Anzahl Snoods Blau'
    NUM_TIGHTS_HALF = 'Anzahl (Tights) [Half Leg]'
    NUM_TIGHTS_KORSAIR = 'Anzahl (Tights) [Korsair]'
    NUM_TIGHTS_KORSAIR_PLUS = 'Anzahl (Tights) [Korsair Plus]'
    TYPE_TIGHTS = 'Schnitt (Tights)'
    SIZE_TIGHTS = 'Größe (Tights)'
    COMMENTS = 'Sonstige Anmerkungen'


def extract_items(row: pd.Series):
    name = row[Col.NAME.value].strip()
    order = []
    for num, prod, color in [
        (Col.NUM_DARK, Product.JERSEY, Color.DARK),
        (Col.NUM_LIGHT, Product.JERSEY, Color.WHITE),
        (Col.NUM_DARK_LONG, Product.JERSEY_LONG, Color.DARK),
        (Col.NUM_LIGHT_LONG, Product.JERSEY_LONG, Color.WHITE),
        (Col.NUM_BLACK_LONG, Product.JERSEY_LONG, Color.BLACK),
        (Col.NUM_DARK_TANK, Product.TANK, Color.DARK),
        (Col.NUM_LIGHT_TANK, Product.TANK, Color.WHITE),
    ]:
        order += get_similar_items(row, num, prod, Col.TYPE_JERSEY, Col.SIZE_JERSEY, color)

    order += get_similar_items(row, Col.NUM_HEADBANDS, Product.HEADBAND, None, None, Color.DEFAULT)
    order += get_similar_items(row, Col.NUM_SNOODS_WHITE, Product.SNOOD, None, None, Color.WHITE)
    order += get_similar_items(row, Col.NUM_SNOODS_BLUE, Product.SNOOD, None, None, Color.DARK)
    order += get_similar_items(row, Col.NUM_SHORTS, Product.SHORT, Col.TYPE_SHORTS, Col.SIZE_SHORTS, Color.DEFAULT)
    order += get_similar_items(row, Col.NUM_HOODIES_NO_ZIP, Product.HOODIE_NO_ZIP, Col.TYPE_HOODIES_NO_ZIP, Col.SIZE_HOODIES_NO_ZIP, Color.DEFAULT)
    order += get_similar_items(row, Col.NUM_HOODIES_ZIP, Product.HOODIE_ZIP, Col.TYPE_HOODIES_ZIP, Col.SIZE_HOODIES_ZIP, Color.DEFAULT)
    order += get_similar_items(row, Col.NUM_GLOVES, Product.GLOVES, Col.TYPE_GLOVES, Col.SIZE_GLOVES)
    order += get_similar_items(row, Col.NUM_TIGHTS_HALF, Product.TIGHT_HALF, Col.TYPE_TIGHTS, Col.SIZE_TIGHTS)
    order += get_similar_items(row, Col.NUM_TIGHTS_KORSAIR, Product.TIGHT_KORSAIR, Col.TYPE_TIGHTS, Col.SIZE_TIGHTS)
    order += get_similar_items(row, Col.NUM_TIGHTS_KORSAIR_PLUS, Product.TIGHT_KORSAIR_PLUS, Col.TYPE_TIGHTS, Col.SIZE_TIGHTS)
    return name, order


def get_similar_items(row, num: Col, product: Product, type_: Col, size: Col, color=None):
    items = []
    if num:
        try:
            item = Item(
                product,
                row[type_.value] if type_ else None,
                row[size.value] if size else None,
                color or Color.DEFAULT,
                jersey_name='',
                jersey_number='')
            quantity = row[num.value]
            quantity = 0 if pd.isna(quantity) else int(quantity)
            items += [item] * quantity
        except KeyError as err:
            pass
            # print(err)
    return items


if __name__ == '__main__':
    main()
