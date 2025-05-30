import tkinter as tk
from tkinter import messagebox, simpledialog
from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from datetime import datetime
import configparser

# --- Options 1 & 2: Scrape books.toscrape.com ---
def scrape_books():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False,
                                   executable_path=r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
        page = browser.new_page()
        page.goto("https://books.toscrape.com/")
        page.wait_for_selector("article.product_pod")
        books = page.query_selector_all("article.product_pod")

        data = []
        for book in books:
            title = book.query_selector("h3 a").get_attribute("title").strip()
            price = book.query_selector("div.product_price p.price_color").inner_text().strip()
            data.append([title, price])
        browser.close()
        return data

def save_to_excel_books(data, filename="books.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = datetime.today().strftime("%Y-%m-%d")
    ws.append(["Title", "Price"])
    for row in data:
        ws.append(row)
    wb.save(filename)

# --- Options 3 & 4: Scrape quotes.toscrape.com with login ---
def scrape_quotes_with_login(username, password):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False,
                                   executable_path=r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
        page = browser.new_page()
        login_url = "https://quotes.toscrape.com/login"
        page.goto(login_url)

        page.fill('input#username', username)
        page.fill('input#password', password)
        page.click('input[type="submit"]')

        page.wait_for_selector('div.quote')
        quotes = page.query_selector_all('div.quote')
        data = []
        for q in quotes:
            text = q.query_selector("span.text").inner_text().strip()
            author = q.query_selector("small.author").inner_text().strip()
            data.append([text, author])
        browser.close()
        return data

def show_message_box(data, title="Data"):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    message = "\n\n".join(f"{r[0]} — {r[1]}" for r in data)
    messagebox.showinfo(title, message, parent=root)
    root.destroy()

def scrape_with_config_login():
    config = configparser.ConfigParser()
    config.read("configgs.ini")
    username = config['LOGIN']['username']
    password = config['LOGIN']['password']
    data = scrape_quotes_with_login(username, password)
    if data:
        show_message_box(data, title="Quotes (Login via Config)")
    else:
        print("No quotes found after login.")

def scrape_with_user_input():
    root = tk.Tk()
    root.withdraw()
    username = simpledialog.askstring("Login", "Enter username:")
    password = simpledialog.askstring("Login", "Enter password:", show='*')
    if not username or not password:
        print("Username or password missing.")
        return
    data = scrape_quotes_with_login(username, password)
    if data:
        show_message_box(data, title="Quotes (Login via User Input)")
    else:
        print("No quotes found after login.")

def main_menu():
    while True:
        print("\nPlease select an option:")
        print("1. Scrape book data & display in message box (Books to Scrape)")
        print("2. Scrape book data & save in Excel file (Books to Scrape)")
        print("3. Scrape quotes with user id/password from config (Quotes to Scrape)")
        print("4. Scrape quotes with user input (Quotes to Scrape)")
        print("5. Exit")

        choice = input("Enter your choice (1-5): ").strip()

        if choice == '1':
            data = scrape_books()
            if data:
                show_message_box(data, title="Books")
            else:
                print("No books found.")
        elif choice == '2':
            data = scrape_books()
            if data:
                save_to_excel_books(data)
                print("Books saved to 'books.xlsx'.")
            else:
                print("No books found.")
        elif choice == '3':
            scrape_with_config_login()
        elif choice == '4':
            scrape_with_user_input()
        elif choice == '5':
            print("Exiting program.")
            break
        else:
            print("Invalid choice. Please try again.")

if __name__ == "__main__":
    main_menu()
