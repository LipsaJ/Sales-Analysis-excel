import pandas as pd
import random
from datetime import date, timedelta

def generate_sales_data(num_rows):
    products = ['Laptop', 'Smartphone', 'Tablet', 'Desktop Computer', 'Printer']
    representatives = ['John', 'Sarah', 'Michael', 'Lisa', 'Robert', 'Susan']
    categories = ['Electronics', 'Appliances', 'Clothing', 'Books', 'Toys']
    regions = ['North', 'South', 'East', 'West']
    shipping_methods = ['Express', 'Standard', 'Next Day']
    payment_methods = ['Credit Card', 'PayPal', 'Cash On Delivery']

    last_names = ['Smith', 'Jonas', 'Pitt', 'Bluth', 'Bing', 'Geller']

    # Dictionary to store product information (name, ID, and price)
    product_info = {
        'Laptop': {'product_id': 101, 'price': 1200},
        'Smartphone': {'product_id': 102, 'price': 500},  # Updated price for Smartphone
        'Tablet': {'product_id': 103, 'price': 800},     # Updated price for Tablet
        'Desktop Computer': {'product_id': 104, 'price': 1500},
        'Printer': {'product_id': 105, 'price': 300},
    }

    data = []

    for _ in range(num_rows):
        product_name = random.choice(products)

        # Retrieve product ID and price from the product_info dictionary
        product_id = product_info[product_name]['product_id']
        sales_amount = product_info[product_name]['price']

        sales_representative = random.choice(representatives)
        date_of_sale = date(2023, 7, 1) + timedelta(days=random.randint(0, 30))
        category = random.choice(categories)
        customer_name = random.choice(['Priya','Nick','Joe','Sophie','Kevin','Danielle','Johnson', 'Smith', 'Brown', 'Lee','Brad','genie','Laila','Arjun','kalki','Abhay','Desco','Kaitlyn','Trevor']) + ' ' + random.choice(last_names)
        region = random.choice(regions)
        shipping_method = random.choice(shipping_methods)
        payment_method = random.choice(payment_methods)

        data.append([product_id, product_name, sales_representative, sales_amount,
                     date_of_sale, category, customer_name, region, shipping_method, payment_method])

    # Create a DataFrame from the data
    df = pd.DataFrame(data, columns=['Product ID', 'Product Name', 'Sales Representative', 'Sales Amount ($)',
                                     'Date of Sale', 'Category', 'Customer Name', 'Region', 'Shipping Method', 'Payment Method'])

    # Save the DataFrame to an Excel file
    df.to_excel('sales_data.xlsx', index=False)

if __name__ == '__main__':
    num_rows = 10000
    generate_sales_data(num_rows)
