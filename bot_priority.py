import pandas as pd
import logging
import re
import time
import html
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

EXCEL_FILE = "Total Site List GP_2.xlsx"  # Apnar Excel file name
BOT_TOKEN = "8355889026:AAHfHVGBSP4QArHG2fvj5IjX45rH5ORNBKE"  # Telegram Bot Token Ekhane Din

SITE_COL_NAMES = [
    'site name', 'gp site name', 'gp id', 'site id', 'site_id', 'siteid', 
    'site code', 'sitecode', 'site', 'airtel code', 'robi new site code', 
    'siteref', 'provider id', 'iss code', 'customer site ref', 'provider site code'
]

LAT_NAMES = ['latitude', 'lat', 'actual lat', 'site_lat', 'latitude (n)']
LON_NAMES = ['longitude', 'lon', 'actual lon', 'site_lon', 'longitude (e)']

# Data Storage by Sheet Name: { 'GP': { 'DHBDA4': [entry] }, 'ISS_1': { ... } }
SHEET_INDEX = {}
SHEET_ORDER = []  # Excel sheets ordering tracking

def clean_value(val):
    if pd.isna(val):
        return None
    val_str = str(val).strip()
    if val_str.lower() in ['nan', 'none', '<na>', 'null', '0', '0.0', '']:
        return None
    return val_str

def load_and_index_excel():
    global SHEET_INDEX, SHEET_ORDER
    print("⏳ Excel File RAM-e indexing kora hocche...")
    start_time = time.time()
    
    try:
        excel_data = pd.read_excel(EXCEL_FILE, sheet_name=None)
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        return

    SHEET_INDEX.clear()
    SHEET_ORDER = list(excel_data.keys())
    total_records = 0

    for sheet_name, df in excel_data.items():
        if df.empty:
            continue

        df.columns = [str(col).strip() for col in df.columns]
        df = df.loc[:, ~df.columns.str.contains('^Unnamed', case=False, na=False)]

        site_col = None
        for col in df.columns:
            if col.strip().lower() in SITE_COL_NAMES:
                site_col = col
                break

        if site_col and site_col in df.columns:
            SHEET_INDEX[sheet_name] = {}
            for _, row in df.iterrows():
                site_key = clean_value(row.get(site_col))
                if not site_key:
                    continue

                site_key_clean = site_key.upper()
                
                entry = {
                    'sheet': sheet_name,
                    'details': [],
                    'lat': None,
                    'lon': None
                }

                for col_name, val in row.items():
                    col_str = str(col_name).strip()
                    val_cleaned = clean_value(val)

                    if col_str.lower().startswith('unnamed') or not val_cleaned:
                        continue

                    entry['details'].append((col_str, val_cleaned))

                    col_lower = col_str.lower()
                    if col_lower in LAT_NAMES:
                        entry['lat'] = val_cleaned
                    elif col_lower in LON_NAMES:
                        entry['lon'] = val_cleaned

                if site_key_clean not in SHEET_INDEX[sheet_name]:
                    SHEET_INDEX[sheet_name][site_key_clean] = []
                SHEET_INDEX[sheet_name][site_key_clean].append(entry)
                total_records += 1

    end_time = time.time()
    print(f"⚡ RAM Indexing Complete! {total_records} records loaded in {round(end_time - start_time, 2)}s.")

def search_site_priority(site_id):
    query = str(site_id).strip().upper()
    
    # Priority Sheet Order: Prothome GP, erpor ISS/ISS_1, erpor baki gulo
    priority_sheets = []
    
    # Check GP sheet
    for s in SHEET_ORDER:
        if s.strip().lower() == 'gp':
            priority_sheets.append(s)

    # Check ISS / ISS_1 sheet
    for s in SHEET_ORDER:
        if s.strip().lower() in ['iss', 'iss_1', 'iss1'] and s not in priority_sheets:
            priority_sheets.append(s)

    # Remaining sheets (ISS_2, ISS_3, ISS_4 etc.)
    for s in SHEET_ORDER:
        if s not in priority_sheets:
            priority_sheets.append(s)

    matched_tabs_count = 0
    results = []

    # Iterate sheet by sheet based on priority
    for sheet_name in priority_sheets:
        if matched_tabs_count >= 2:  # Maximum 2-ti tab pabar por thame jabe
            break

        sheet_data = SHEET_INDEX.get(sheet_name, {})
        matched_keys = [k for k in sheet_data.keys() if query in k]

        if matched_keys:
            tab_has_match = False
            for key in matched_keys:
                entries = sheet_data[key]
                for entry in entries:
                    sheet_name_safe = html.escape(entry['sheet'])
                    sheet_result = [f"📊 <b>Tab / Sheet:</b> <code>{sheet_name_safe}</code>"]
                    
                    for col_str, val_cleaned in entry['details']:
                        col_safe = html.escape(str(col_str))
                        val_safe = html.escape(str(val_cleaned))
                        sheet_result.append(f"• <b>{col_safe}:</b> {val_safe}")

                    lat_val = entry['lat']
                    lon_val = entry['lon']
                    if lat_val and lon_val:
                        try:
                            lat_clean = re.findall(r"[-+]?\d*\.\d+|\d+", str(lat_val))
                            lon_clean = re.findall(r"[-+]?\d*\.\d+|\d+", str(lon_val))
                            if lat_clean and lon_clean:
                                lat_f = float(lat_clean[0])
                                lon_f = float(lon_clean[0])
                                sheet_result.append(f"📍 <b>Coordinates:</b> <code>{lat_f}, {lon_f}</code>")
                                sheet_result.append(f'🗺 <b>Google Map:</b> <a href="https://maps.google.com/?q={lat_f},{lon_f}">Open Location</a>')
                        except ValueError:
                            pass

                    results.append("\n".join(sheet_result))
                    tab_has_match = True

            if tab_has_match:
                matched_tabs_count += 1

    if results:
        header_text = f"🔍 <b>Search Query:</b> <code>{html.escape(query)}</code>\n"
        header_text += f"⚡ <b>Priority Search:</b> Displaying data from maximum {matched_tabs_count} tab(s).\n"
        return header_text + "\n\n" + "\n\n-------------------------\n\n".join(results)
    else:
        return f"❌ Site ID <code>{html.escape(query)}</code> kono Tab-e paowa jayni."

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_text = update.message.text
    result = search_site_priority(user_text)
    
    if len(result) > 4000:
        for i in range(0, len(result), 4000):
            await update.message.reply_text(result[i:i+4000], parse_mode="HTML", disable_web_page_preview=False)
    else:
        await update.message.reply_text(result, parse_mode="HTML", disable_web_page_preview=False)

if __name__ == "__main__":
    load_and_index_excel()

    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    print("🤖 Priority Bot is active...")
    app.run_polling()