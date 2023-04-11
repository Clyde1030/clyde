# Counter

from collections import Counter
a = "aaaaabbbbbccc"
my_counter = Counter(a) # Counter return how many repition each alphebat has
my_counter.keys() # returns the keys
my_counter.values() 
my_counter.most_common(1) # the most common element
my_counter.most_common(2) # the two most common elements
my_counter.most_common(2)[0][0]
list(my_counter.elements())


# namedtuple: 
from collections import namedtuple

color = (55,155,255) # Regular tuple: immutable
print(color[0]) # Does not know what 55 means if it has some meaningful definition before

color = {'red':55, 'green':155, 'blue':255} # Regular dictionary mutable - means you can change the value
color['red']

Color = namedtuple('Color',['red','green','blue']) # First argument is tuple's name and second is keys
namedtupleColor = Color(55, 155, 255) 
namedtupleColor.red

namedtupleColor = Color(red = 55, green = 155, blue = 255) 
namedtupleColor

# OrderedDict
from collections import OrderedDict
ordered_dict = OrderedDict()
ordered_dict['a'] = 1
ordered_dict['b'] = 2
ordered_dict['c'] = 3
ordered_dict['d'] = 4




