from playwright.sync_api import sync_playwright
import json

def grab_cookies():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        print("Opening Toyota EPCWEB...")
        page.goto("https://epcwebserver.tmhg.jp/home#")

        print(">>> PLEASE LOG IN AND NAVIGATE TO THE SEARCH PAGE <<<")
        page.pause() 

        # Once you hit Resume, we grab the cookies and close the browser
        print("Grabbing cookies...")
        cookies = context.cookies()
        
        # Save cookies to a file
        with open("cookies.json", "w") as f:
            json.dump(cookies, f)
            
        print("Cookies saved to cookies.json! Closing browser.")
        browser.close()

if __name__ == "__main__":
    grab_cookies()
