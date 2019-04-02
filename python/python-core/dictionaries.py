
# Define a New Dictionary
# Come in Key-Value Pairs
NewDict = {'Key1': 1000, 'Key2': 2000, 'Key3': 3000}      
print('Here is my dictionary {}'.format(NewDict), end='\n\n')
           

# Add New Item to Dictionary
# Syntax MyDict[NewKey] = NewValue
NewDict['Key4'] = 4000
print('Here is my dictionary after adding the new value {}'.format(NewDict), end='\n\n')

# Delete Item From Dictionary

# Method One:
# Syntax del MyDict[Key]
del NewDict['Key1']
print('Here is my dictionary after the delete: {}'.format(NewDict), end='\n\n')

# Method Two:
# Use the Pop Method
NewDict.pop('Key1')
print('Here is my dictionary after the delete: {}'.format(NewDict), end='\n\n')


# Delete A Random Key-Value Pair from a Dictionary
RandKeyVal = NewDict.popitem()
print('Here is a random key-value that I deleted: {}'.format(RandKeyVal), end='\n\n')



## Count the Number of Pairs in My Dictionary
DictCount = len(NewDict)
print("The number of Key-Value Pairs in this Dictionary is: {}".format(DictCount),end='\n\n')


# Get All The Keys in a Dictionary
print("Here are the keys in my Dictionary: {}".format(NewDict.keys()), end='\n\n')

# Get All The Values in a Dictionary
print("Here are the values in my Dictionary: {}".format(NewDict.values()), end='\n\n')

# Get All The Key-Value Pairs in a Dictionary
print("Here are the key-value pairs in my Dictionary: {}".format(NewDict.items()), end='\n\n')


# Check if Key Exist in Dictionary
if 'Key1' in NewDict:
    print("The Key Exist",end='\n\n')
else:
    print("The Key Does Not Exist",end='\n\n')


# Check if Value Exist in Dictionary
if 3000 in NewDict.values():
    print("The Value Exist",end='\n\n')
else:
    print("The Value Does Not Exist",end='\n\n')


# Check if Item Exist in Dictionary
if ('Key3', 3000) in NewDict.items():
    print("The Value Exist",end='\n\n')
else:
    print("The Value Does Not Exist",end='\n\n')


# Update a Value in the Dictionary
# The update() method adds element(s) to the dictionary if the key is not in the dictionary. 
# If the key is in the dictionary, it updates the key with the new value.

# Method One:
NewDict['Key3'] = 7000
print(NewDict['Key3'], end='\n\n')


# Method Two:
NewDict.update({'Key3': 8000})
print(NewDict['Key3'], end='\n\n')
    
NewDict.update({'Key5': 8000}) # This will add a new key
print(NewDict['Key5'], end='\n\n')


## Copy the Dictionary
NewDictTwo = NewDict.copy()
print(NewDictTwo, end='\n\n')

# Get A Value From the Dictionary
print(NewDictTwo.get('Key4', 'This Key does not exist'), end='\n\n')
print(NewDictTwo.get('Key5', 'This Key does not exist'), end='\n\n')


# Get A Value From the Dictionary, but if it doesn't exist then add that key-value pair to the dictionary.
print("Here is my dictionary BEFORE SetDefault {}".format(NewDictTwo))
NewDictTwo.setdefault('Key6', "10000")
print("Here is my dictionary AFTER SetDefault {}".format(NewDictTwo), end='\n\n')


# Generate a new dictionary from a sequence of keys and then assign a value to each key in that new dictionary.

MySequence = ('Key9', 'Key10', 'Key11', 'Key12')
NewDictThree = dict.fromkeys(MySequence, "InitialValue")
print(NewDictThree)

NewDictThree= dict.fromkeys(MySequence)
print(NewDictThree)


# Clear All Elements from A Dictionary
NewDictThree.clear()
print(NewDictThree)



# Loop throgh all the keys in a dictionary
for key in NewDictTwo.keys():
    print(key)

# Loop throgh all the values in a dictionary
for value in NewDictTwo.values():
    print(value)

# Loop throgh all the items in a dictionary
for item in NewDictTwo.items():
    print(item)

for key, value in NewDictTwo.items():
    print(key, value)



## Special Dunder Methods

#Does My Dictionary Contain a Key
item = NewDictTwo.__contains__('Key4')

#What is the Size of My Dictionary in Memory
sizeof = NewDictTwo.__sizeof__()

#Convert Dictionary to String.
dictStr = NewDictTwo.__str__()

print(item)
print(sizeof)
print(type(dictStr))
