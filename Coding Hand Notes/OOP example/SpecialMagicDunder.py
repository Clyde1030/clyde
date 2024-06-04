
class Employee:
    
    raise_amount = 1.04

    def __init__(self, first, last, pay):
        self.first = first
        self.last = last
        self.pay = pay
        self.email = first + '.' + last + '@email.com'

    def fullname(self):
        return '{} {}'.format(self.first, self.last)
    
    def apply_raise(self):
        self.pay = int(self.pay * self.raise_amount)


    # __repr__ and __str__ change how our classes are displayed
    def __repr__(self): # for logging and debugging
        return "Employee('{}', '{}', {})".format(self.first, self.last, self.pay)
    
    def __str__(self): # Meant to be displayed to the end users
        return '{} - {}'.format(self.fullname(),self.email)        

    def __add__(self, other):
        return self.pay + other.pay

    def __len__(self):
        return len(self.fullname())


emp_1 = Employee('Corey','Schafer',60000)
emp_2 = Employee('John','Doe',60000)

print(emp_1 + emp_2)
print(len(emp_1))




# print(emp_1)

# When we do this, what it does in the background is actually calling __repr__ and __str__ methods
repr(emp_1) 
emp_1.__repr__()
str(emp_1)
emp_1.__str__()



