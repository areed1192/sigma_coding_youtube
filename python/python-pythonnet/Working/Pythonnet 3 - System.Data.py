# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'

# # System.Data in Pythonnet

import clr

clr.AddReference("System")
clr.AddReference("System.Data")

import System
import System.Data


data_table_parent = System.Data.DataTable()

column_1 = System.Data.DataColumn()
column_1.ColumnName = "Product_ID"
column_1.DataType = System.Type.GetType("System.String")
column_1.Unique = True

column_2 = System.Data.DataColumn()
column_2.ColumnName = "Product_Name"
column_2.DataType = System.Type.GetType("System.String")
column_2.Unique = False

column_3 = System.Data.DataColumn()
column_3.ColumnName = "Product_Price"
column_3.DataType = System.Type.GetType("System.Int32")
column_3.Unique = False

data_table_parent.Columns.Add(column_1)
data_table_parent.Columns.Add(column_2)
data_table_parent.Columns.Add(column_3)

# for column in data_table_parent.Columns:
#     print(column.ColumnName)

for i in range(0, 10):

    row = data_table_parent.NewRow()
    row["Product_ID"] = "0000_" + str(i)
    row["Product_Name"] = "Product " + str(i)
    row["Product_Price"] = i * 100

    data_table_parent.Rows.Add(row)


data_reader = data_table_parent.CreateDataReader()


if data_reader.HasRows == True:

    while data_reader.Read():
        print([str(data_reader[row]) + " " for row in range(0, data_reader.FieldCount)])

