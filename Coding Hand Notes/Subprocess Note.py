import os
import re
import pandas as pd
import calendar


df = pd.DataFrame(columns=['BP#','FA#',"Producer's Name","Producer's NPN", "Trans Date","List Bill","Name of Bank/Company","Policy #","Insured's Name","Plan","Coverage ID","Trans Type","Premium Asset Value","Billing Frequency","Rate","Share","Earnings","Policy Year","Product ID","Product Desc","Target/Excess"])

output = []
with open('Protective.txt','rt') as file:
    for line in file:
        output.append(line)


output[1][0:10]
output[1][22:38]
output[1][93:102]
output[1][105:115]
output[1][116:126]
output[1][127:151]
output[1][22:38]



output[1]









import subprocess


subprocess.run('ls') # list file and folder in this directory
subprocess.run('dir') # list file and folder in this directory
subprocess.run('ls -la',shell=True) # We can pass the entire command as a string with shell=True
subprocess.run('ls -la') # Error
subprocess.run(['ls', '-la']) # same as with shell=True is passing by a string
p1 = subprocess.run(['ls', '-la']) # p1 is a completed process object
print(p1.args) # return the argument of p1
print(p1.returncode) # show if we have error code or not. 0 means no error
print(p1.stdout) 
p1 = subprocess.run(['ls', '-la'],capture_output=True) 
print(p1.stdout) # this is captured as bytes
print(p1.stdout.decode()) # Covert to strings
p1 = subprocess.run(['ls', '-la'],capture_output=True, text=True)  # capture_output = True sends stdout and stderror to subprocess pipe so we can access them
print(p1.stdout)
p1 = subprocess.run(['ls', '-la'],stdout=subprocess.PIPE, text=True) # Send result to PIPE so you can aggregate them with logging result

with open('output.txt','w') as f:
    p1 = subprocess.run(['ls', '-la'],stdout=f, text=True) # Send result to a file 

p1 = subprocess.run(['ls', '-la', 'dne'],capture_output=True, text=True) # We don't get any error by python
print(p1.returncode) # No longer 0
print(p1.stderr)

p1 = subprocess.run(['ls', '-la', 'dne'],capture_output=True, text=True,check=True)
print(p1.stderr)

p1 = subprocess.run(['ls', '-la', 'dne'],stderr=subprocess.DEVNULL)
print(p1.stderr)





