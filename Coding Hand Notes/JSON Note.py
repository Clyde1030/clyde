import json


people_string = '''
{
    "people":[
        {
            "name":"John Smith",
            "phone":"615-555-7164",
            "emails":["johnsmith@email.com", "john.smith@work.com"],
            "has_license": false
        },
        {
            "name":"Jane Doe",
            "phone":"560-555-5153",
            "emails":["janedoe@email.com", "jane.doe@work.com"],
            "has_license": true
   
        }
    ]
}
'''

data = json.loads(people_string) # Read as "load-s"
type(data['people']) # All these info will be converted to Pythin dictionary format

for person in data['people']:
    print(person['name'])



# We can also reverse a Python dictionary to a Json file
for person in data['people']:
    del person['phone']
new_strings = json.dumps(data, indent=2, sort_keys= True)




# Load a json file
with open(r'C:\Users\yu-shenglee\Desktop\Python\states.json') as f:
    data = json.load(f)

for state in data['states']:
    print(state['name'], state['abbreviation'])


# Write to a json file
for state in data['states']:
    del state['area_codes']
with open('new_states.json','w') as f:
    json.dump(data, f, indent=2)

from urllib.request import urlopen
with urlopen("http://finance.yahoo.com/webservice/v1/symbols/allcurrencies/quote?format=json") as response:
    source = response.read()

data = json.loads(source)

print(json.dumps(data, indent = 2))

