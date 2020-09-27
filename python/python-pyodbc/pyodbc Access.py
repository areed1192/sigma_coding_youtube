import pyodbc

# grab the datasources we have access to
pyodbc.dataSources()

# define components of our connection string
driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
filepath = r'C:\Users\Alex\OneDrive\Career - Work Items\Petco\Financial_Analyst\Petco_Financial_Data.accdb'

# on a side note, here is another way to get the driver setup
myDataSources = pyodbc.dataSources()
access_driver = myDataSources['MS Access Database']
driver = access_driver

# create a connection to the database
cnxn = pyodbc.connect(driver = driver, dbq = filepath, autocommit = True)
crsr = cnxn.cursor()

# grab all the tables
tables_list = list(crsr.tables())

# define the components of a query
table_name = 'ACTUALS_EXPENSE'

# define query
query = "SELECT * FROM {}".format(table_name)

# execute the query
crsr.execute(query)

# fetch a single row, this returns a different type of object
one_row = crsr.fetchone()

# accessing the individual columns
print(one_row[0])
print(one_row.CER)

# specify how many rows you want to fetch at a time, default is one
crsr.fetchmany(5)

# determine the number of rows updated
crsr.rowcount

# This read/write attribute specifies the number of rows to fetch at a time with
# fetchmany(). It defaults to 1 meaning to fetch a single row at a time.
print(crsr.arraysize)

# I could change it, say to 10
crsr.arraysize = 10

# then run it again, notice we get 10 back this time.
crsr.fetchmany()

# get some details about the columns in the table
columns_list = list(crsr.columns())

# 0) table_cat
# 1) table_schem
# 2) table_name
# 3) column_name
# 4) data_type
# 5) type_name
# 6) column_size
# 7) buffer_length
# 8) decimal_digits
# 9) num_prec_radix
# 10) nullable
# 11) remarks
# 12) column_def
# 13) sql_data_type
# 14) sql_datetime_sub
# 15) char_octet_length
# 16) ordinal_position
# 17) is_nullable

# grab some statistics about a particular table
for item in list(crsr.statistics(table_name)):
    print(item)

# foreign keys if there are any
foreign_keys_list = list(crsr.foreignKeys())

# the get type info object returns information about the different data types available in the database.
list(crsr.getTypeInfo())

help(pyodbc)
help(crsr)
help(crsr.connection)

