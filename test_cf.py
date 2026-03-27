import cloudscraper
import os
from garminconnect import Garmin

email = os.environ.get("EMAIL")
password = os.environ.get("PASSWORD")

if email and password:
    garmin = Garmin(email=email, password=password)
    # Patch the session with cloudscraper to bypass Cloudflare 429 errors
    scraper = cloudscraper.create_scraper()
    garmin.garth.sess = scraper
    
    try:
        if garmin.login():
            print("Successfully logged in with Cloudscraper!")
    except Exception as e:
        print(f"Login failed: {e}")
else:
    print("EMAIL or PASSWORD not set")
