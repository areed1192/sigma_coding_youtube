# import the CLR
import clr

# Add a reference to the System Namespace, System.IO Namespace, and System.Data Namespace.
clr.AddReference("System")
clr.AddReference("System.IO")
clr.AddReference("System.Data")

# import all the Namespaces
import System
import System.Data


# Let's create a DataTable Class, this Class Object can be found in the System.Data Namespace
data_table_parent = System.Data.DataTable()

# Give the table a name
data_table_parent.TableName = 'PRODUCT_TABLE_1'

'''
    A data table can contain columns, so let's define a few columns to add to our table. Each column,
    will have a data type assigned to it, a name, whether it allows null values, and whether all the
    values have to be unique.
'''

# Define the `Product_ID` column.
column_1 = System.Data.DataColumn()
column_1.ColumnName = "Product_ID"
column_1.DataType = System.Type.GetType("System.String")
column_1.Unique = True
column_1.AllowDBNull = False

# Define the `Product_Name` column.
column_2 = System.Data.DataColumn()
column_2.ColumnName = "Product_Name"
column_2.DataType = System.Type.GetType("System.String")
column_2.Unique = False
column_2.AllowDBNull = True

# Define the `Product_Price` column.
column_3 = System.Data.DataColumn()
column_3.ColumnName = "Product_Price"
column_3.DataType = System.Type.GetType("System.Int32")
column_3.Unique = False
column_3.AllowDBNull = True


# METHOD ONE: ONE AT A TIME
# Add all the columns to the table.
data_table_parent.Columns.Add(column_1)
data_table_parent.Columns.Add(column_2)
data_table_parent.Columns.Add(column_3)

# I CANT HAVE THE SAME COLUMN NAMES, SO LETS CREATE SOME NEW ONES.

# Define the `Product_ID` column.
column_1 = System.Data.DataColumn()
column_1.ColumnName = "Product_ID1"
column_1.DataType = System.Type.GetType("System.String")
column_1.Unique = True
column_1.AllowDBNull = False

# Define the `Product_Name` column.
column_2 = System.Data.DataColumn()
column_2.ColumnName = "Product_Name1"
column_2.DataType = System.Type.GetType("System.String")
column_2.Unique = False
column_2.AllowDBNull = True

# Define the `Product_Price` column.
column_3 = System.Data.DataColumn()
column_3.ColumnName = "Product_Price1"
column_3.DataType = System.Type.GetType("System.Int32")
column_3.Unique = False
column_3.AllowDBNull = True


# METHOD TWO: ALL AT ONCE
# Create a second table .
data_table_parent_2 = System.Data.DataTable()

# Give the table a name
data_table_parent_2.TableName = 'PRODUCT_TABLE_2'

# METHOD ONE: DEFINE AN ARRAY, SET IT TO TYPE SYSTEM.DATA.DATACOLUMN
# Define an Array of Type Data Column
columns_array = System.Array[System.Data.DataColumn]((column_1, column_2, column_3))

# METHOD ONE: DEFINE A REGULAR PYTHON LIST.
# Define a list that contains all the columns
columns_array = [column_1, column_2, column_3]

# Use the `AddRange` method to add all the columns to the table.
data_table_parent_2.Columns.AddRange(columns_array)


# Let's verify method ONE worked by looping through each column in the table's Columns collection.
for column in data_table_parent.Columns:

    # And print the columns name
    print('Table {} has a column named: {}'.format(data_table_parent.TableName, column.ColumnName))

print('\n')

# Let's verify method TWO worked by looping through each column in the table's Columns collection.
for column in data_table_parent_2.Columns:

    # And print the columns name
    print('Table {} has a column named: {}'.format(data_table_parent_2.TableName, column.ColumnName))

print('\n')

# Let's now add some rows to our table.


# METHOD ONE: USE A LOOP
# define a range
for i in range(0, 10):

    # With the Data Table Object call the `NewRow` method to create a DataRow Object.
    row = data_table_parent.NewRow()

    # If the columns have names you can just specify the column you want to add a value to.
    row["Product_ID"] = "0000_" + str(i)
    row["Product_Name"] = "Product " + str(i)
    row["Product_Price"] = i * 100

    # Once you've finshed adding the elements, add the row to the table by calling the Add method from the Rows collection.
    data_table_parent.Rows.Add(row)


# METHOD TWO: USE A LIST
my_elements = ['0000_20', 'Product_20', 20000]

# Define a new row
row = data_table_parent.NewRow()

# Set the ItemArray property
row.ItemArray = my_elements

# Finally add the row
data_table_parent.Rows.Add(row)


# Let's Read the Data from our Table. There are many ways to do this but a simple way is to use the DataReader Object.


# First specify the table you want to read, call the `CreateDataReader` method, and then store it in a variable.
data_reader = data_table_parent.CreateDataReader()

# Now you can create a DataReader object even if the table has now rows. Let's make sure the reader has some rows before we proceed.
if data_reader.HasRows == True:
    
    # If does use a while loop in conjunction with the `Read` method. It will keep going as long as their is data to read.
    while data_reader.Read():

        # Print the data, here I decided to use List Comprehension.
        print([str(data_reader[row]) + " " for row in range(0, data_reader.FieldCount)])


# Here is another way to read the data, loop through the rows collection.
for row in data_table_parent.Rows:

    # Grab the ItemArray property.
    row_values = row.ItemArray

    # Print the values.
    print(list(row_values))


# If we want to be fancy we could write all the content of the table to XML. First define a writer.
system_io_writer = System.IO.StringWriter()

# Then with the table call the `WriteXml` method. Make sure to pass through the writer and in this case I'm going to 
# include the table Schema and the data. Alternative is `System.Data.XmlWriteMode.IgnoreSchema`
data_table_parent.WriteXml(system_io_writer, System.Data.XmlWriteMode.WriteSchema)

# Call the Writer because it now has all the content and then call the `ToString` method to convert it to a string.
xml_content = system_io_writer.ToString()

# Print the content.
print('\n')
print(xml_content)
