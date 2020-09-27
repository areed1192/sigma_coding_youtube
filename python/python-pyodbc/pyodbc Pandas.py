import pyodbc
import pandas

# Define our Query to create the table.
create_table_query = """
-- Create the Table if it Does not exist.
IF Object_ID('youtube_videos') IS NULL

CREATE TABLE [sigma-coding].[dbo].[youtube_videos]
(
    [video_id] NVARCHAR(MAX) NOT NULL,
    [published_at] DATETIME NULL,
    [channel_id] NVARCHAR(MAX) NULL,
    [channel_title] NVARCHAR(MAX) NULL,
    [video_title] NVARCHAR(MAX) NULL,
    [video_description] NVARCHAR(MAX) NULL,
    [category_id] INT NULL,
    [duration] NVARCHAR(MAX) NULL,
    [definition] NVARCHAR(6) NULL,
    [caption] BIT NULL,
    [licensed_content] BIT NULL,
    [has_custom_thumbnail] BIT NULL, 
    [view_count] INT NULL,
    [like_count] INT NULL,
    [dislike_count] INT NULL,
    [comment_count] INT NULL,
)
"""

# Define the Components of the Connection String.
DRIVER = '{ODBC Driver 17 for SQL Server}'
SERVER_NAME = "ALEX-LAPTOP\ALEX_SQL_SERVER"
DATABASE_NAME = "sigma-coding"

CONNECTION_STRING = """
Driver={driver};
Server={server};
Database={database};
Trusted_Connection=yes;
""".format(
    driver=DRIVER,
    server=SERVER_NAME,
    database=DATABASE_NAME
)

# Create a connection object.
connection_object: pyodbc.Connection = pyodbc.connect(CONNECTION_STRING)

# Create a Cursor Object, using the connection.
cursor_object: pyodbc.Cursor = connection_object.cursor()

# Define the File Path.
data_file = "python/python-pyodbc/youtube_data.csv"

# Load the Data.
youtube_df: pandas.DataFrame = pandas.read_csv(
    data_file,
    infer_datetime_format=True,
    parse_dates=True
)

# Parse the Published At Column.
youtube_df['published_at'] = pandas.to_datetime(youtube_df['published_at'])

# Define the Insert Query.
sql_insert = """
INSERT INTO [sigma-coding].[dbo].[youtube_videos]
 (
    [video_id],
    [published_at],
    [channel_id],
    [channel_title],
    [video_title],
    [video_description],
    [category_id],
    [duration],
    [definition],
    [caption],
    [licensed_content],
    [has_custom_thumbnail], 
    [view_count],
    [like_count],
    [dislike_count],
    [comment_count]
)
VALUES
(
    ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
)
"""

# Create the Table.
cursor_object.execute(create_table_query)

# Commit the Table.
cursor_object.commit()

# Convert the DataFrame to a RecordSet.
df_records = youtube_df.values.tolist()

# Execute it.
cursor_object.executemany(sql_insert, df_records)

# Commit the Transactions.
cursor_object.commit()

# Define the Select Query.
sql_select = "SELECT * FROM [sigma-coding].[dbo].[youtube_videos]"

# Execute the Query.
records = cursor_object.execute(sql_select).fetchall()

# Define the column names.
columns = [column[0] for column in cursor_object.description]

# Dump to a Pandas DataFrame.
youtube_dump_df = pandas.DataFrame.from_records(
    data=records,
    columns=columns
)

# print the head.
print(youtube_dump_df.head())