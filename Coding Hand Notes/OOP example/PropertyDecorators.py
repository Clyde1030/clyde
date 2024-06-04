
class Employee:

    def __init__(self, first, last):
        self.first = first
        self.last = last

    @property # We can access this method like an attribute even though it's defined as a method
    def email(self):
        return '{}.{}@email.com'.format(self.first, self.last)

    @property
    def fullname(self):
        return '{} {}'.format(self.first, self.last)

    @fullname.setter
    def fullname(self, name):
        first, last = name.split(' ')
        self.first = first
        self.last = last
    
    @fullname.deleter
    def fullname(self):
        print('Delete Name!')
        self.first = None
        self.last = None


emp_1 = Employee('John', 'Smith')

emp_1.fullname = 'Jim Schafer'


print(emp_1.first)
print(emp_1.email)
print(emp_1.fullname)


del emp_1.fullname





