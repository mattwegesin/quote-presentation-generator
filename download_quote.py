import requests
import warnings
import pyotp
import re
import os
warnings.filterwarnings('ignore')

def download_ihg_quote_xlsx(quote_id: str, output_dir: str = "."):
    """
    Authenticates and downloads the IHG Quote XLSX file for a given quote ID.
    Returns the filepath of the downloaded file.
    """
    session = requests.Session()
    login_url = "https://quote.hospitalitytechnologies.com/index.php"
    payload = {"username": "api_user", "password": "H0spTech123!", "submitLogin": "1"}
    res = session.post(login_url, data=payload, allow_redirects=False)

    if res.status_code == 200 and "loginVerify2FA.php" in res.text:
        totp = pyotp.TOTP("K4OT5KIPWD3IFEBB")
        session.post("https://quote.hospitalitytechnologies.com/loginVerify2FA.php", data={"otp": totp.now(), "otpSubmit": " Verify 2FA "}, allow_redirects=False)

    download_url = f"https://quote.hospitalitytechnologies.com/access/admin_ihgQuote.php?quoteID={quote_id}"
    download_res = session.get(download_url, stream=True)
    
    if download_res.status_code != 200:
        print(f"Failed to download quote {quote_id}. Status Code: {download_res.status_code}")
        return None

    # Try to get the real filename from the server headers
    filename = f"quote_{quote_id}.xlsx"
    if 'Content-Disposition' in download_res.headers:
        d = download_res.headers['Content-Disposition']
        fname_match = re.findall("filename=(.+)", d)
        if len(fname_match) > 0:
            filename = fname_match[0].strip('"; ')

    filepath = os.path.join(output_dir, filename)
    
    with open(filepath, 'wb') as f:
        for chunk in download_res.iter_content(chunk_size=8192):
            f.write(chunk)
            
    return filepath

if __name__ == "__main__":
    # Example usage using the internal ID for quote #27850 (which is 40125)
    quote_internal_id = "40125" 
    print(f"Downloading IHG Quote for ID {quote_internal_id}...")
    saved_path = download_ihg_quote_xlsx(quote_internal_id)
    if saved_path:
        print(f"Successfully downloaded to: {saved_path}")
