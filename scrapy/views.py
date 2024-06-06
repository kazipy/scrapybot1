import json
import re
import pandas as pd
import os
import gspread
from django.http import HttpResponse, JsonResponse
from django.views.decorators.csrf import csrf_exempt
from oauth2client.service_account import ServiceAccountCredentials

# Define your verify token and Google Sheets details
VERIFY_TOKEN = 'your token'
GOOGLE_SHEET_NAME = 'orders'
GOOGLE_CREDENTIALS_FILE = 'json path'

# Define column names for the Google Sheet
COLUMN_NAMES = ['Customer Name', 'Product Name', 'Price', 'Quantity', 'Address']


# Define regular expression patterns for extracting order details
patterns = {
    'Customer Name': re.compile(r'Customer Name:\s*(.*)'),
    'Product Name': re.compile(r'Product Name:\s*(.*)'),
    'Price': re.compile(r'Price:\s*(\d+(\.\d+)?)'),
    'Quantity': re.compile(r'Quantity:\s*(\d+)'),
    'Address': re.compile(r'Address:\s*(.*)')
}

# Authenticate and initialize the Google Sheets API client
def get_google_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)
    sheet = client.open(GOOGLE_SHEET_NAME).sheet1  # Assumes you're working with the first sheet
    return sheet

@csrf_exempt
def webhook(request):
    if request.method == 'GET':
        token_sent = request.GET.get('hub.verify_token')
        if token_sent == VERIFY_TOKEN:
            return HttpResponse(request.GET.get('hub.challenge'))
        return HttpResponse('Invalid verification token')

    elif request.method == 'POST':
        data = json.loads(request.body.decode('utf-8'))
        for entry in data['entry']:
            for event in entry['messaging']:
                if 'message' in event:
                    response = handle_message(event)
                    return response
                elif 'postback' in event:
                    handle_postback(event)
        return HttpResponse('EVENT_RECEIVED')

    return HttpResponse('Invalid request')

def handle_message(event):
    sender_id = event['sender']['id']
    message_text = event['message'].get('text')
    print(f"Received message from {sender_id}: {message_text}")
    if message_text:
        order_details = parse_order(message_text)
        if order_details:
            save_order_to_google_sheet(order_details)
            print(f"Order saved: {order_details}")
            response_data = {
                'status': 'success',
                'order_details': order_details
            }
            return JsonResponse(response_data)
        else:
            print("Order details not found.")
            return JsonResponse({'status': 'error', 'message': 'Order details not found'})

def handle_postback(event):
    sender_id = event['sender']['id']
    payload = event['postback'].get('payload')
    print(f"Received postback from {sender_id}: {payload}")

def parse_order(message):
    order_details = {}
    for key, pattern in patterns.items():
        match = pattern.search(message)
        if match:
            order_details[key] = match.group(1)
        else:
            print(f"Pattern for {key} not matched.")
    if len(order_details) == len(patterns):
        return order_details
    return None

def save_order_to_google_sheet(order_details):
    sheet = get_google_sheet()
    # Check if the first row is empty to add column headers
    if not sheet.row_values(1):
        sheet.append_row(COLUMN_NAMES)

    values = [order_details[key] for key in patterns.keys()]
    sheet.append_row(values)

def save_order_to_excel(order_details):
    df = pd.DataFrame([order_details])
    file_path = 'orders.xlsx'
    
    try:
        if not os.path.exists(file_path):
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, header=True)
            print(f"File created and order saved to {file_path}")
        else:
            with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                df.to_excel(writer, index=False, header=False)
            print(f"Order saved to {file_path}")
    except Exception as e:
        print(f"An error occurred while saving the order: {e}")
