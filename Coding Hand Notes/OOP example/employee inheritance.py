
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


class Developer(Employee): # Developer class inherit all methods and attributes from employee class
    raise_amount = 1.10

    def __init__(self, first, last, pay, prog_lang):
        super().__init__(first, last, pay) # let the class inhirit first last and pay attributes from Employee
        # Employee.__init__(self, first, last, pay) # this does the same as above
        self.prog_lang = prog_lang


class Manager(Employee):

    def __init__(self, first, last, pay, employees = None):
        super().__init__(first, last, pay) # let the class inhirit first last and pay attributes from Employee
        if employees is None:
            self.employees = []
        else:
            self.employees = employees

    def add_emp(self, emp):
        if emp not in self.employees:
            self.employees.append(emp)

    def remove_emp(self, emp):
        if emp in self.employees:
            self.employees.remove(emp)
    
    def print_emps(self):
        for emp in self.employees:
            print('-->', emp.fullname())



dev_1 = Developer('Corey','Schafer',60000,'Python')
dev_2 = Developer('John','Doe',60000,'Java')

mgr_1 = Manager('Sue', 'Smith', 90000, [dev_1])

print(mgr_1.email)
mgr_1.add_emp(dev_2)
mgr_1.remove_emp(dev_1)
mgr_1.print_emps()

# Develop will have email attribute too
print(dev_1.email)
dev_1.apply_raise()
print(dev_2.email)
print(dev_1.pay)

# print(help(Developer)) 
print(dev_2.email)


isinstance(mgr_1, Manager)
isinstance(mgr_1, Employee)
isinstance(mgr_1, Developer)

issubclass(Manager, Developer)
issubclass(Manager, Employee)



