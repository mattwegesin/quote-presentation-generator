import requests
import warnings
import pyotp
from bs4 import BeautifulSoup
import json
warnings.filterwarnings('ignore')

def get_quotes_by_inn_code(inn_code: str):
    session = requests.Session()
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    })
    login_url = "https://quote.hospitalitytechnologies.com/index.php"
    
    login_payload = {
        "username": "api_user",
        "password": "H0spTech123!",
        "submitLogin": "1"
    }
    response = session.post(login_url, data=login_payload, allow_redirects=False)
    
    if response.status_code == 200 and "loginVerify2FA.php" in response.text:
        totp_url = "https://quote.hospitalitytechnologies.com/loginVerify2FA.php"
        secret = "K4OT5KIPWD3IFEBB"
        
        verify_payload = {
            "otp": pyotp.TOTP(secret).now(),
            "otpSubmit": " Verify 2FA "
        }
        session.post(totp_url, data=verify_payload, allow_redirects=False)

    search_url = "https://quote.hospitalitytechnologies.com/access/index.php"
    search_payload = {
        "limitInnCode": inn_code,
        "dateLimit": "ALL",
        "salesRep": "ALL"
    }
    
    search_response = session.post(search_url, data=search_payload)
    soup = BeautifulSoup(search_response.text, 'html.parser')
    quotes = []
    
    # Iterate over all rows in the document to avoid relying on strict table hierarchy
    for row in soup.find_all('tr'):
        cols = [col.text.strip() for col in row.find_all(['td', 'th'])]
        links = [a.get('href') for a in row.find_all('a')]
        
        # A valid quote row typically has 13 columns and starts with a Job # (e.g. #27850)
        if len(cols) >= 11 and cols[0].startswith('#'):
            quote_data = {
                "job": cols[0].split(' ')[0].strip(), # Extract just the #ID
                "create_date": cols[1],
                "approve_date": cols[2],
                "package": cols[6],
                "status": cols[7],
                "assigned_rep": cols[8],
                "inn_code": cols[9],
                "grand_total": cols[10],
                "links": links
            }
            quotes.append(quote_data)
                
    return quotes

import pyotp

def mark_quote_won(quote_id: str):
    try:
        session = requests.Session()
        session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        })
        login_url = "https://quote.hospitalitytechnologies.com/index.php"
        
        login_payload = {
            "username": "api_user",
            "password": "H0spTech123!",
            "submitLogin": "1"
        }
        response = session.post(login_url, data=login_payload, allow_redirects=False)
        
        if response.status_code == 200 and "loginVerify2FA.php" in response.text:
            totp_url = "https://quote.hospitalitytechnologies.com/loginVerify2FA.php"
            secret = "K4OT5KIPWD3IFEBB"
            verify_payload = {
                "otp": pyotp.TOTP(secret).now(),
                "otpSubmit": " Verify 2FA "
            }
            session.post(totp_url, data=verify_payload, allow_redirects=False)

        url = f"https://quote.hospitalitytechnologies.com/access/customerView.php?id={quote_id}"
        payload = {
            "id": quote_id,
            "status": "Won",
            "submitStatusUpdate": "Update Status"
        }
        post_res = session.post(url, data=payload)
        return post_res.status_code == 200
    except Exception as e:
        print(f"Error marking quote as won: {e}")
        return False

def create_and_approve_gsa_quote(inn_code: str, property_name: str, total_monthly: float, room_count: int = 0, meeting_fee: float = 0.0, notes: str = ""):
    try:
        session = requests.Session()
        session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        })
        login_url = "https://quote.hospitalitytechnologies.com/index.php"
        
        login_payload = {
            "username": "api_user",
            "password": "H0spTech123!",
            "submitLogin": "1"
        }
        response = session.post(login_url, data=login_payload, allow_redirects=False)
        
        if response.status_code == 200 and "loginVerify2FA.php" in response.text:
            totp_url = "https://quote.hospitalitytechnologies.com/loginVerify2FA.php"
            secret = "K4OT5KIPWD3IFEBB"
            verify_payload = {
                "otp": pyotp.TOTP(secret).now(),
                "otpSubmit": " Verify 2FA "
            }
            session.post(totp_url, data=verify_payload, allow_redirects=False)
            
        # 1. Create Quote
        create_url = "https://quote.hospitalitytechnologies.com/access/buildQuote.php"
        create_payload = {
            "packageID": "117", # GSA ONLY (US)
            "propertyHotelCode": inn_code,
            "propertyName": property_name or f"{inn_code} GSA",
            "numRooms": str(room_count),
            "submitNewQuote": "Create"
        }
        create_res = session.post(create_url, data=create_payload)
        
        quote_id = None
        if 'id=' in create_res.url:
            quote_id = create_res.url.split('id=')[1].split('&')[0]
        else:
            import re
            match = re.search(r'location=\'([^\']+)\'', create_res.text)
            if match:
                redirect_url = match.group(1)
                if 'id=' in redirect_url:
                    quote_id = redirect_url.split('id=')[1].split('&')[0]

        if not quote_id:
            return None
            
        # 2. Update the prices and meeting room qty
        admin_url = f"https://quote.hospitalitytechnologies.com/access/admin_adminView.php?id={quote_id}"
        admin_res = session.get(admin_url)
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(admin_res.text, 'html.parser')
        form = soup.find('form', id='adminForm')
        if form:
            data = {}
            for input_tag in form.find_all('input'):
                name = input_tag.get('name')
                if not name: continue
                type_attr = input_tag.get('type', 'text').lower()
                value = input_tag.get('value', '')
                if type_attr in ('checkbox', 'radio'):
                    if input_tag.has_attr('checked'):
                        data[name] = value
                elif type_attr not in ('submit', 'button'):
                    data[name] = value
                    
            for select_tag in form.find_all('select'):
                name = select_tag.get('name')
                if name:
                    selected = select_tag.find('option', selected=True)
                    if selected:
                        data[name] = selected.get('value', '')
                    else:
                        opts = select_tag.find_all('option')
                        if opts: data[name] = opts[0].get('value', '')
                        
            for textarea_tag in form.find_all('textarea'):
                name = textarea_tag.get('name')
                if name:
                    data[name] = textarea_tag.text.strip()
            
            if notes:
                # Append to existing notes or replace
                existing_notes = data.get('quoteNotes', '')
                if existing_notes:
                    data['quoteNotes'] = existing_notes + "\n\n" + notes
                else:
                    data['quoteNotes'] = notes
                        
            # Find the items
            first_item_id = None
            meeting_item_id = None
            for inp in form.find_all('input'):
                if inp.get('name', '').startswith('qty'):
                    tr = inp.find_parent('tr')
                    if tr:
                        if 'Meeting Rooms' in tr.text:
                            meeting_item_id = inp.get('name').replace('qty', '')
                        elif not first_item_id:
                            first_item_id = inp.get('name').replace('qty', '')
                            
            if first_item_id:
                # Set first item to total_monthly - meeting_fee
                base_fee = total_monthly - meeting_fee
                data[f'retail{first_item_id}'] = str(base_fee)
                data[f'retailExt{first_item_id}'] = str(base_fee)
                data[f'cost{first_item_id}'] = str(base_fee)
                data[f'costExt{first_item_id}'] = str(base_fee)
                
            if meeting_item_id and meeting_fee > 0:
                data[f'qty{meeting_item_id}'] = "1"
                data[f'cost{meeting_item_id}'] = str(meeting_fee)
                data[f'costExt{meeting_item_id}'] = str(meeting_fee)
                data[f'retail{meeting_item_id}'] = str(meeting_fee)
                data[f'retailExt{meeting_item_id}'] = str(meeting_fee)
                    
            data['submitUpdateQuote'] = "Update Quote"
            session.post(admin_url, data=data)
        
        return quote_id
    except Exception as e:
        print(f"Error creating GSA quote: {e}")
        return None

if __name__ == "__main__":
    inn_code = "TEST123"
    results = get_quotes_by_inn_code(inn_code)
    print(json.dumps(results, indent=2))
