from flask import Flask, render_template, redirect, url_for, session, request
import pandas as pd
import os
import requests
import json
import openpyxl 
from openpyxl import load_workbook

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Set a secret key for session management

API_KEY = '2EJ4D4oNPhYy4RoOPDZb2Qy9G1UfmI6c'  # Financial Modeling Prep API Key

def is_email_registered(email, filename='users_data.xlsx'):
    """
    Check if the email is already registered in the Excel file.
    """
    try:
        df = pd.read_excel(filename)
        if email in df['Email ID'].values:
            return True
        else:
            return False
    except FileNotFoundError:
        # File not found, so email can't be registered yet
        return False
    
def get_current_stock_price(symbol):
    url = f'https://financialmodelingprep.com/api/v3/quote/{symbol}?apikey={API_KEY}'
    try:
        response = requests.get(url)
        data = response.json()
        if data:
            return data[0]['price']  # Assuming the API returns the current price in 'price' field
    except Exception as e:
        print(f"Error fetching current stock price for {symbol}: {e}")
    return None

def search_symbol_by_company_name(company_name):
    url = f'https://financialmodelingprep.com/api/v3/search?query={company_name}&limit=1&apikey={API_KEY}'
    try:
        response = requests.get(url)
        data = response.json()
        if data:
            # Assuming the first result is the most relevant one
            return data[0]['symbol']
    except Exception as e:
        print(f"Error searching for company name {company_name}: {e}")
    return None

def validate_and_get_symbols(company_names):
    validated_symbols = []
    for name in company_names.split(','):
        name = name.strip()
        symbol = search_symbol_by_company_name(name)
        if symbol:
            validated_symbols.append(symbol)
        else:
            print(f"Company {name} not found or does not have a stock symbol.")
    return validated_symbols

def save_to_excel(email, name, dob, income, country, stocks, username, password, filename='users_data.xlsx'):
    # Serialize the stocks information into JSON
    stocks_json = json.dumps(stocks)

    data_dict = {
        'Email ID': email,
        'Name': name,
        'Date of Birth': dob,
        'Monthly Income': income,
        'Country': country,
        'Username': username,
        'Password': password,
        'Stocks Info': stocks_json  # Storing the serialized JSON
    }
    
    df = pd.DataFrame([data_dict])
    
    if not os.path.exists(filename):
        # If the file does not exist, create it and write the dataframe
        df.to_excel(filename, engine='openpyxl', index=False)
    else:
        # If the file exists, append the new data
        book = load_workbook(filename)
        writer = pd.ExcelWriter(filename, engine='openpyxl')
        writer.book = book
        writer.sheets = {ws.title: ws for ws in book.worksheets}
        
        startrow = writer.sheets['Sheet1'].max_row
        
        # Append without overwriting the existing sheet
        df.to_excel(writer, sheet_name='Sheet1', startrow=startrow, index=False, header=False)
        
        writer.save()

def validate_credentials(username, password, filename='users_data.xlsx'):
    try:
        df = pd.read_excel(filename)
        user_row = df.loc[df['Username'] == username]
        if not user_row.empty and user_row.iloc[0]['Password'] == password:
            return True
        else:
            return False
    except FileNotFoundError:
        return False

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        email = request.form.get('email')
        name = request.form.get('name')
        dob = request.form.get('dob')
        income = request.form.get('income')
        country = request.form.get('country')
        username = request.form.get('username')
        password = request.form.get('password')
        
        # Extract stocks data as before
        stocks = []
        for key in request.form:
            if key.startswith('stocks['):
                index, field = key.strip(']').split('[')[1:]
                index = int(index)
                while index >= len(stocks):
                    stocks.append({'name': '', 'shares': 0, 'invested': 0.0})
                stocks[index][field] = request.form[key]
        
        save_to_excel(email, name, dob, income, country, stocks, username, password)
        
        return redirect(url_for('register'))
    return render_template('register.html')
@app.route('/signin', methods=['GET', 'POST'])
def signin():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        if validate_credentials(username, password):
            session['username'] = username  # Set session
            return redirect(url_for('profile'))  # Redirect to a profile route
        else:
            return 'Invalid Username or Password!'
    return render_template('signin.html')

# Define a new route for the profile page
@app.route('/profile')
@app.route('/profile')
def profile():
    if 'username' in session:
        username = session['username']
        # Fetch user data
        df = pd.read_excel('users_data.xlsx')
        user_data = df.loc[df['Username'] == username].iloc[0]
        symbols = user_data['Stock'].split(', ')
        
        # Calculate total stock value and total profit
        total_stock_value = 0
        total_profit = 0
        for symbol in symbols:
            current_price = get_current_stock_price(symbol)
            # Assuming the purchase price and number of shares are stored in the user_data
            purchase_price = ...  # Fetch from user_data
            num_shares = ...  # Fetch from user_data
            total_stock_value += current_price * num_shares
            total_profit += (current_price - purchase_price) * num_shares
        
        monthly_income = user_data['Monthly Income']
        net_worth = total_stock_value + monthly_income  # Simplified net worth calculation
        
        return render_template('profile.html', monthly_income=monthly_income, total_stock_value=total_stock_value, net_worth=net_worth, symbols=symbols, total_profit=total_profit)
    else:
        return redirect(url_for('signin'))

@app.route('/signout')
def sign_out():
    session.clear()  # Clear the user's session
    return redirect(url_for('signin'))

@app.route('/')
def home():
    return redirect(url_for('signin'))
if __name__ == '__main__':
    app.run(debug=True)