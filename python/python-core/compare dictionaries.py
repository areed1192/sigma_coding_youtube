NewDict1 = {'Key1': 1000, 'Key2': 2000, 'Key3': 3000}     
NewDict2 = {'Key1': 1000, 'Key2': 2000, 'Key3': 3000, 'Key4': 4000}     

# Find Keys in Common
comkeys = NewDict1.keys() & NewDict2.keys()
# Returns {'Key3', 'Key1', 'Key2'}

# Find The Difference in Keys ORDER MATTERS WITH THIS ONE
difkeys = NewDict2.keys() - NewDict1.keys()
# Returns {'Key4'}

# Return a Set that contains the keys from both dictionaries, but no duplicates.
unikeys = NewDict1.keys() | NewDict2.keys()
# Returns {'Key3', 'Key4', 'Key1', 'Key2'}

# Find Items in Common
comitms = NewDict2.items() & NewDict1.items()
# Returns {('Key3', 3000), ('Key2', 2000), ('Key1', 1000)}

# Find The Difference in Items ORDER MATTERS WITH THIS ONE
difitms = NewDict2.items() - NewDict1.items()
# Returns {('Key4', 4000)}

# Return a Set that contains the key-value pairs from both dictionaries, but no duplicates.
uniitms = NewDict2.items() | NewDict1.items()
# {('Key3', 3000), ('Key2', 2000), ('Key1', 1000), ('Key4', 4000)}

# Is NewDict1 a SubSet of NewDict2? In other words, do all the keys in NewDict1 exist in NewDict2?
set(NewDict1) <= set(NewDict2)
# Returns True

# Is NewDict1 a SuperSet of NewDict2? In other words, does NewDict1 contain all the keys in NewDict2?
set(NewDict1) >= set(NewDict2)
# Returns False

# Does One Dictionary Equal Another Dictionary
NewDict1 == NewDict2
# Returns False
