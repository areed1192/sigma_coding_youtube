import pyodbc
import textwrap

# Define the Price Data.
price_data = [
    [ 2.00, 3.00, 1.00, 2.40, 100.00, '1/2/2019'],
    [ 3.00, 3.00, 5.00, 9.40, 300.00, '2/1/2020'],
    [ 4.00, 2.00, 1.00, 2.40, 200.00, '3/1/2021']
]

# define the server and the database
server = 'DESKTOP-MOA8O6M\SQLEXPRESS' 
database = 'stock_database'  

# Define the connection string
cnxn = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server}; \
    SERVER='+ server +'; \
    DATABASE='+ database +';\
    Trusted_Connection=yes;'
)

# Create the Cursor.
cursor = cnxn.cursor()

# Loop through to insert each row.
for index, row in enumerate(price_data):
    
    # define an insert query with place holders for the values.
    insert_query = textwrap.dedent('''
        INSERT INTO td_price_data (close_price, high, low, open_price, volume, day_value) 
        VALUES (?, ?, ?, ?, ?, ?);
    ''')
    
    # define the values
    values = (row[0], row[1], row[2], row[3], row[4], row[5])
    
    # insert the data into the database
    cursor.execute(insert_query, values)

# commit the inserts.
cnxn.commit()
    
# grab all the rows from the table
cursor.execute('SELECT * FROM td_price_data')
for row in cursor:
    print(row)
    
# close the cursor and connection  
cursor.close()
cnxn.close()
