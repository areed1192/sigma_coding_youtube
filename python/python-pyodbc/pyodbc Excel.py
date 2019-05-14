import pyodbc

# we may be interested in finding all the drivers we have access to
for driver in pyodbc.drivers():
   
    # print the driver name
    print(driver)
    
    # if the driver name has '.xlsx' in it we found it!
    if '.xlsx' in driver:
        myDriver = driver

# define our connection string
conn_str = (r'DRIVER={'+ myDriver +'};'
            r'DBQ=C:\Users\Alex\Desktop\Video List.xlsx;'
            r'ReadOnly=1') # ReadOnly set to 0 means we can edit the data.

# define our connection, autocommit MUST BE SET TO TRUE, also we can edit data.
cnxn = pyodbc.connect(conn_str, autocommit=True)
crsr = cnxn.cursor()

# loop through all the tables
for worksheet in crsr.tables():
    
    # display the worksheet
    print(worksheet)
    
    # grab the table name.
    tableName = worksheet[2]
    
# define our query to grab the data.
# we want this "SELECT Topic FROM [Sheet1$]"
crsr.execute("SELECT Topic FROM [{}]".format(tableName))

# print each row of data.
for row in crsr:
    print(row.Topic)
