
# list are ordered
a = ['foo', 'bar', 'baz', 'qux']
b = ['baz', 'qux', 'bar', 'foo']

print(a == b)
print(a is b)
print([1, 2, 3, 4] == [4, 1, 3, 2])

#Lists Can Contain Arbitrary Objects
a = [21.42, 'foobar', 3, 4, 'bark', False, 3.14159]

# declare empty list
my_list = []

# declare list of integers
my_list = [1, 2, 3, 4]

# declare list of mixed datatypes
my_list = [1, 'mystring', 3.4, True]

# Lists Can Be Nested
# declare a nested list
my_list = [9, 'mystring',[1,2,3],['hi','there','my']]

# List Elements Can Be Accessed by Index
# INDEXING LIST


# declare list of integers
my_list = [1, 2, 3, 4]

my_first_element = my_list[0]
#returns 1
my_second_element = my_list[1]
#returns 2
my_third_element = my_list[2]
#returns 3

# returns error, we can only use integers for indexing.
# my_third_element = my_list[2.0]

# returns error, we only have 4 elements. This is accessing a fifth element.
# my_third_element = my_list[4]

# declare nested list of integers
my_list = [[1,2,3],[4,5,6],[7,8,9]]

my_first_element = my_list[0][0]
#returns 1
my_second_element = my_list[1][2]
#returns 6
my_third_element = my_list[2][1]
#returns 8


# SLICING LIST


my_list = ['y','o','d','a']

zero_to_one = my_list[0:1]
two_to_end = my_list[2:]
two_to_first = my_list[:-2]

print(two_to_first)


# CHANGING & ADDING ELEMENTS


my_list = [3, 4, 5, 6]

# change the first element
my_list[0] = 1

# change elements 2 to 3
my_list[2:3] = [7,8]

# add new item to end of the list
my_list.append(9)

# add several new item to the list
my_list.extend([10, 11, 12])

# using the '+' operator
my_list = my_list + [10,10,10]

# using the '*' operator
print(['myName' * 3])

# insert 6 in position 1
my_list.insert(1,6)

# insert 5 & 6 at position 3
my_list[3:3] = [5,6]


# DELETING ELEMENTS


my_list = [1, 2, 3, 4, 5]

# delete the second position
del my_list[2]

# delete multiple items
del my_list [0:1]

# delete the entire list
del my_list


# remove e from the list
my_list = ['m','a','c','e','w','i','n','d','u']
my_list.remove('e')

# remove the item at position 1
my_list.pop(1)

# remove the item at the last position
my_list.pop()

# remove the all the items in a list
my_list.clear()

# find the index of a particular item
fruit_list = ['Apple', 'Oranges','Pears']
print(fruit_list.index('Oranges'))

my_list.sort()
my_list.reverse()
my_new_list = my_list.copy()

# define list with fruits in it
fruit_list = ['Apple', 'Oranges','Pears']

# check if Value exist in List
if 'Apple' in fruit_list:
    print('The fruit exist.')

# loop through list
for fruit in fruit_list:
    print(fruit)

# count the number of items in a list
num_of_items = len(fruit_list)


num_list = [10, 8, 2]

# get the largest number in a list.
max_item = max(num_list)

# get the smallest number in a list.
min_item = min(num_list)
print(max_item)
print(min_item)

# create an enumeration object with an index and value.
print(enumerate(num_list))

# sum all the numbers in a list.
print(sum(num_list))
