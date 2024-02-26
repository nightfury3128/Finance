from flask import Flask, render_template, redirect, url_for, session, request
import pandas as pd
import os
import requests

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

def save_to_excel(email, name, dob, income, country, companies, username, password):
    filename = 'users_data.xlsx'
    validated_symbols = validate_and_get_symbols(companies)
    
    data_dict = {
        'Email ID': email,
        'Name': name,
        'Date of Birth': dob,
        'Monthly Income': income,
        'Country': country,
        'Username': username,
        'Password': password,
        'Stock Symbols': ', '.join(validated_symbols)
    }
    
    df = pd.DataFrame([data_dict])
    
    if not os.path.exists(filename):
        df.to_excel(filename, engine='openpyxl', index=False)
    else:
        with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            startrow = writer.sheets['Sheet1'].max_row if 'Sheet1' in writer.sheets else 0
            df.to_excel(writer, sheet_name='Sheet1', startrow=startrow, index=False, header=(startrow == 0))

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
        email = request.form['email']
        name = request.form['name']
        dob = request.form['dob']
        income = request.form['income']
        country = request.form['country']
        companies = request.form['companies']  # User inputs company names comma-separated
        username = request.form['username']
        password = request.form['password']
        
        # Check if the email is already registered
        if is_email_registered(email):
            # Handle the case where the email is already registered
            # For example, by returning an error message or redirecting to an error page
            return 'Email already registered. Please use a different email.'
        
        save_to_excel(email, name, dob, income, country, companies, username, password)
        
        return redirect(url_for('signin'))
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