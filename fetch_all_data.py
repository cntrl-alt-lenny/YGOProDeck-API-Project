# Script by Cntrl_Alt_Lenny

# This script fetches *all* YGOPRODeck API data (including 'misc' fields)
# and downloads the resulting DataFrame as an Excel file in Google Colab.

# 1. It displays a visually pleasing message describing the data points.
# 2. It fetches every possible card detail (main, nested, and misc).
# 3. It saves and downloads the data as an Excel file with "correct" column names.

# ======================
#  INSTALL DEPENDENCIES
# ======================
# This ensures xlsxwriter is available for Pandas ExcelWriter
!pip install xlsxwriter

# ======================
#       IMPORTS
# ======================
import requests
import pandas as pd
from datetime import datetime
import warnings
from IPython.display import display, HTML
from google.colab import files

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ======================
#   VISUAL MESSAGE
# ======================
# Below we itemise each relevant field as documented by YGOPRODeck, including 'misc' fields.
html_message = """
<h2 style="color:DarkBlue;">Fetching All YGOPRODeck API Data</h2>
<p>This script will retrieve <strong>every</strong> piece of card information returned by the
<a href="https://db.ygoprodeck.com/api/v7/cardinfo.php" target="_blank">YGOPRODeck API</a>,
including:</p>
<ul>
  <li><strong>Card Identification:</strong> id (Card ID), konami_id (Konami ID), name (Card Name), type (Card Type), frameType (Frame Type)</li>
  <li><strong>Attributes/Stats:</strong> desc (Description), atk (Attack), def (Defense), level (Level), race (Race), attribute (Attribute), archetype (Archetype), scale (Scale), linkval (Link Value), linkmarkers (Link Markers)</li>
  <li><strong>Banlist Info:</strong> ban_tcg, ban_ocg, ban_goat</li>
  <li><strong>Card Sets (expanded):</strong> set_name, set_code, set_rarity, set_rarity_code, set_price (and more if available)</li>
  <li><strong>Card Images (expanded):</strong> image_url, image_url_small, image_url_cropped</li>
  <li><strong>Card Prices (expanded):</strong> cardmarket_price, tcgplayer_price, ebay_price, amazon_price, coolstuffinc_price</li>
  <li><strong>Misc. Fields (with misc=yes):</strong> beta_name, views, viewsweek, upvotes, downvotes, formats, treated_as, tcg_date, ocg_date, konami_id, md_rarity, has_effect</li>
  <li><strong>ygoprodeck_url:</strong> direct link to the card’s YGOPRODeck page</li>
</ul>
<p>All of the above will be neatly combined into a single Excel file.</p>
"""
display(HTML(html_message))

# ======================
#   FETCH ALL DATA
# ======================
# We include 'misc=yes' to retrieve additional fields.
url = 'https://db.ygoprodeck.com/api/v7/cardinfo.php?misc=yes'
response = requests.get(url)
response.raise_for_status()  # Raises an error if there's a connection/HTTP issue
data = response.json()

# Convert to DataFrame
df = pd.json_normalize(data['data'], sep='_')

# ======================
#   EXPAND NESTED COLUMNS
# ======================
def expand_column(df, column_name, prefix):
    """
    Explodes a list column, normalises it, and merges back into the main DataFrame
    with a prefix on the expanded columns.
    """
    if column_name in df.columns:
        df = df.explode(column_name).reset_index(drop=True)
        expanded_df = pd.json_normalize(df[column_name])
        if not expanded_df.empty:
            expanded_df.columns = [f"{prefix}_{col}" for col in expanded_df.columns]
            df.drop(columns=[column_name], inplace=True)
            expanded_df = expanded_df.reset_index(drop=True)
            df = pd.concat([df, expanded_df], axis=1)
        else:
            df.drop(columns=[column_name], inplace=True)
    return df

# Expand card_sets, card_images, card_prices
df = expand_column(df, 'card_sets', 'set')
df = expand_column(df, 'card_images', 'image')
df = expand_column(df, 'card_prices', 'price')

# ======================
#   RENAME COLUMNS
# ======================
column_mapping = {
    # Core data
    'id': 'Card ID',
    'konami_id': 'Konami ID',
    'name': 'Card Name',
    'type': 'Card Type',
    'frameType': 'Frame Type',
    'desc': 'Description',
    'atk': 'Attack',
    'def': 'Defense',
    'level': 'Level',
    'race': 'Race',
    'attribute': 'Attribute',
    'archetype': 'Archetype',
    'scale': 'Scale',
    'linkval': 'Link Value',
    'linkmarkers': 'Link Markers',
    'ygoprodeck_url': 'YGOPRODeck URL',

    # Banlist info
    'banlist_info.ban_goat': 'Banlist (GOAT)',
    'banlist_info.ban_tcg': 'Banlist (TCG)',
    'banlist_info.ban_ocg': 'Banlist (OCG)',

    # Card sets (expanded)
    'set_set_name': 'Set Name',
    'set_set_code': 'Set Code',
    'set_set_rarity': 'Set Rarity',
    'set_set_rarity_code': 'Set Rarity Code',
    'set_set_price': 'Set Price',
    'set_set_edition': 'Set Edition',        # only appears if tcgplayer_data is used
    'set_set_url': 'Set URL',                # only appears if tcgplayer_data is used

    # Card images (expanded)
    'image_id': 'Image ID',
    'image_image_url': 'Image URL',
    'image_image_url_small': 'Image URL (Small)',
    'image_image_url_cropped': 'Image URL (Cropped)',

    # Card prices (expanded)
    'price_cardmarket_price': 'Price (Cardmarket)',
    'price_tcgplayer_price': 'Price (TCGplayer)',
    'price_ebay_price': 'Price (eBay)',
    'price_amazon_price': 'Price (Amazon)',
    'price_coolstuffinc_price': 'Price (CoolStuffInc)',

    # Misc fields
    'beta_name': 'Beta Name',
    'views': 'Views',
    'viewsweek': 'Views (Week)',
    'upvotes': 'Upvotes',
    'downvotes': 'Downvotes',
    'formats': 'Formats',
    'treated_as': 'Treated As',
    'tcg_date': 'TCG Release Date',
    'ocg_date': 'OCG Release Date',
    'md_rarity': 'MD Rarity',
    'has_effect': 'Has Effect'
}
df.rename(columns=column_mapping, inplace=True)

# ======================
#  REORDER (OPTIONAL)
# ======================
# Define a desired column order. Only columns that actually exist will be retained in that order.
desired_columns = [
    'Card ID', 'Konami ID', 'Card Name', 'Card Type', 'Frame Type', 'Description',
    'Attack', 'Defense', 'Level', 'Race', 'Attribute', 'Archetype', 'Scale',
    'Link Value', 'Link Markers', 'Banlist (GOAT)', 'Banlist (TCG)', 'Banlist (OCG)',
    'Set Name', 'Set Code', 'Set Rarity', 'Set Rarity Code', 'Set Price',
    'Set Edition', 'Set URL',
    'Image ID', 'Image URL', 'Image URL (Small)', 'Image URL (Cropped)',
    'Price (Cardmarket)', 'Price (TCGplayer)', 'Price (eBay)', 'Price (Amazon)', 'Price (CoolStuffInc)',
    'Beta Name', 'Views', 'Views (Week)', 'Upvotes', 'Downvotes', 'Formats',
    'Treated As', 'TCG Release Date', 'OCG Release Date', 'MD Rarity', 'Has Effect',
    'YGOPRODeck URL'
]
existing_columns = [col for col in desired_columns if col in df.columns]
df = df[existing_columns]

# ======================
#  FORMAT CARD ID
# ======================
# Pad 'Card ID' to 8 digits, e.g. 1234567 -> 01234567
if 'Card ID' in df.columns:
    df['Card ID'] = df['Card ID'].astype(str).str.zfill(8)

# ======================
#   SAVE TO EXCEL
# ======================
now = datetime.now()
date_time = now.strftime("%Y-%m-%d %I%M%p")
filename = f"All YGOProDeck API Data - {date_time}.xlsx"

with pd.ExcelWriter(filename, engine='xlsxwriter', engine_kwargs={'options': {'strings_to_urls': False}}) as writer:
    df.to_excel(writer, index=False, sheet_name='RAW Data Sheet')

# ======================
#   DOWNLOAD FILE
# ======================
files.download(filename)

# Execution completes here. The file should automatically be downloaded in Colab.
