# *********************************************
# *********************************************
# written using python 3.9.0
# *********************************************
# *********************************************

# *********************************************
# *********************************************
# you need these libraries included
# *********************************************
# *********************************************

import pyodbc
import json
import collections

# *********************************************
# *********************************************
# class to change the returned query array from array[0], array[1], etc to
# array['columnName' ] for each column in the query.
# *********************************************
# *********************************************


class QueryByName():
    def __init__(self, cursor):
        self._cursor = cursor

    def __iter__(self):
        return self

    def __next__(self):
        row = self._cursor.__next__()

        return {description[0]: row[col] for col, description in enumerate(self._cursor.description)}

# *********************************************
# *********************************************
# set up the DB connection.  You'll need to change this to your sql server instance
# or change it up if you're using mysql, postgresql, mongo, etc
# *********************************************
# *********************************************


conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=SAMSUNGLAPTOP\SQLEXPRESS;'
                      'Database=Directory;'
                      'Trusted_Connection=yes;')

# *********************************************
# *********************************************
# open the connection
# *********************************************
# *********************************************

dataRows = conn.cursor()

# *********************************************
# *********************************************
# execute the query
# *********************************************
# *********************************************

dataRows.execute('select * from directory.directory')

# *********************************************
# *********************************************
# create an array to hold the processed data
# *********************************************
# *********************************************

jsonData = []

# *********************************************
# *********************************************
# loop over the query, build the entry for the row, then
# add it to the data array
# *********************************************
# *********************************************

for row in QueryByName(dataRows):
    d = collections.OrderedDict()
    d["directoryID"] = row["DirectoryID"]
    d["name"] = row["PersonName"]
    d["sortName"] = row["SortName"]
    d["email"] = row["EmailAddress"]
    d["department"] = row["Department"]
    d["position"] = row["Position"]
    d["photoFileName"] = row["PhotoFileName"]
    d["bio"] = row["BiographyText"]

    jsonData.append(d)

# *********************************************
# *********************************************
# convert the array to json format
# *********************************************
# *********************************************

outputJSON = json.dumps(jsonData)

# *********************************************
# *********************************************
# write the file
# *********************************************
# *********************************************

with open("directory.json", "w") as f:
    f.write(outputJSON)

# *********************************************
# *********************************************
# close the connection
# *********************************************
# *********************************************

conn.close()
