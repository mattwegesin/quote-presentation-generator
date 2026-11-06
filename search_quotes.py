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

if __name__ == "__main__":
    inn_code = "TEST123"
    results = get_quotes_by_inn_code(inn_code)
    print(json.dumps(results, indent=2))
