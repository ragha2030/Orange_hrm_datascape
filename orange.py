from selenium import webdriver
from selenium.webdriver.common.by import By
from openpyxl import Workbook
import time
import os



option = webdriver.ChromeOptions()
driver = webdriver.Chrome(options=option)

driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login")
time.sleep(10)
driver.find_element(By.XPATH, '//*[@id="app"]/div[1]/div/div[1]/div/div[2]/div[2]/form/div[1]/div/div[2]/input').send_keys("Admin")
time.sleep(7)
driver.find_element(By.XPATH, '/html/body/div/div[1]/div/div[1]/div/div[2]/div[2]/form/div[2]/div/div[2]/input').send_keys("admin123")
time.sleep(5)
driver.find_element(By.XPATH, '/html/body/div/div[1]/div/div[1]/div/div[2]/div[2]/form/div[3]/button').click()
time.sleep(5)
driver.find_element(By.XPATH, '/html/body/div/div[1]/div[1]/aside/nav/div[2]/ul/li[1]/a').click()
time.sleep(5)


wb = Workbook()
admin_ws = wb.active
admin_ws.title = "Data"

admin_ws['A1'] = 'username'
admin_ws['B1'] = 'user_role'
admin_ws['C1'] = 'Employee_name'
admin_ws['D1'] = 'status'

admin_rows = driver.find_elements(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div')
ws_index = admin_ws.max_row + 2
data_list = []
for i in range(1, len(admin_rows)+1):
        data = {}        
        data['Username'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div[' +str(i)+ ']/div/div[2]').text
        data['User_role'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div[' +str(i)+ ']/div/div[3]').text
        data['Employee_name'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div[' +str(i)+ ']/div/div[4]').text
        data['Status'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div[' +str(i)+ ']/div/div[5]').text
        data_list.append(data)

        admin_ws[f'A{ws_index}'] = data['Username']
        admin_ws[f'B{ws_index}'] = data['User_role']
        admin_ws[f'C{ws_index}'] = data['Employee_name']
        admin_ws[f'D{ws_index}'] = data['Status']
       
        ws_index += 1


driver.find_element(By.XPATH, '/html/body/div/div[1]/div[1]/aside/nav/div[2]/ul/li[2]/a').click()
time.sleep(5)

pim_start_row = len(admin_rows) + 3

admin_ws[f'E{pim_start_row}'] = 'ID'
admin_ws[f'F{pim_start_row}'] = 'First_name'
admin_ws[f'G{pim_start_row}'] = 'Last_name'
admin_ws[f'H{pim_start_row}'] = 'Job_title'
admin_ws[f'I{pim_start_row}'] = 'Employee_status'
admin_ws[f'J{pim_start_row}'] = 'Subunit'

pim_rows = driver.find_elements(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div')

pim_data_list = []

pim_index = pim_start_row + 2

for i in range(1, len(pim_rows)+1):
    data = {}
    data['id'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div[' +str(i)+ ']/div/div[2]').text
    data['First_name'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div[' +str(i)+ ']/div/div[3]').text
    data['Last_name'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div[' +str(i)+ ']/div/div[4]').text
    data['Job_title'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div[' +str(i)+ ']/div/div[5]').text
    data['Employment_status'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div[' +str(i)+ ']/div/div[6]').text
    data['Subunit'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div[' +str(i)+ ']/div/div[7]').text
    pim_data_list.append(data)
   
    admin_ws[f'E{pim_index}'] = data['id']
    admin_ws[f'F{pim_index}'] = data['First_name']
    admin_ws[f'G{pim_index}'] = data['Last_name']
    admin_ws[f'H{pim_index}'] = data['Job_title']
    admin_ws[f'I{pim_index}'] = data['Employment_status']
    admin_ws[f'J{pim_index}'] = data['Subunit']

    pim_index += 1

driver.find_element(By.XPATH, '/html/body/div/div[1]/div[1]/aside/nav/div[2]/ul/li[3]/a ').click()
time.sleep(5)

leave_start_row = pim_start_row + len(pim_data_list) + 2

admin_ws[f'K{leave_start_row}'] = 'Date'
admin_ws[f'L{leave_start_row}'] = 'Employee_name'
admin_ws[f'M{leave_start_row}'] = 'Leave_type'
admin_ws[f'N{leave_start_row}'] = 'Leave_balance'
admin_ws[f'o{leave_start_row}'] = 'Number_of_days'
admin_ws[f'P{leave_start_row}'] = 'Status'
admin_ws[f'Q{leave_start_row}'] = 'Comments'

leave_rows = driver.find_elements(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/div[2]/div')
leave_index = leave_start_row + 1
leave_data_list = []

for i in range(1, len(leave_rows)+1):
    data['Date'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/div[2]/div[' +str(i)+ ']/div/div[2]').text
    data['Employee_name'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/div[2]/div[' +str(i)+ ']/div/div[3]').text
    data['Leave_type'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/div[2]/div[' +str(i)+ ']/div/div[4]').text
    data['Leave_balance'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/div[2]/div[' +str(i)+ ']/div/div[5]').text
    data['Number_of_days'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/div[2]/div[' +str(i)+ ']/div/div[6]').text
    data['status'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/div[2]/div[' +str(i)+ ']/div/div[7]').text
    data['Comments'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[2]/div/div[2]/div[' +str(i)+ ']/div/div[8]').text
    leave_data_list.append(data)

    admin_ws[f'K{leave_index}'] = data['Date']
    admin_ws[f'L{leave_index}'] = data['Employee_name']
    admin_ws[f'M{leave_index}'] = data['Leave_type']
    admin_ws[f'N{leave_index}'] = data['Leave_balance']
    admin_ws[f'O{leave_index}'] = data['Number_of_days']
    admin_ws[f'P{leave_index}'] = data['status']
    admin_ws[f'Q{leave_index}'] = data['Comments']

    leave_index += 1

driver.find_element(By.XPATH, '/html/body/div/div[1]/div[1]/aside/nav/div[2]/ul/li[4]/a').click()
time.sleep(5)


time_start_row = leave_start_row + len(leave_data_list) + 1

admin_ws[f'R{time_start_row}'] = 'Employee name'
admin_ws[f'S{time_start_row}'] = 'Timesheet period'

time_rows = driver.find_elements(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div')
time_index = time_start_row + 1
time_data_list = []
 
for i in range(1, len(time_rows)+1):
    data = {}
    data['Employee_name'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div[' +str(i)+ ']/div/div[1]').text
    data['Timesheet_period'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div[' +str(i)+ ']/div/div[2]').text
    time_data_list.append(data)

    admin_ws[f'R{time_index}'] = data['Employee_name']
    admin_ws[f'S{time_index}'] = data['Timesheet_period']

    time_index += 1



driver.find_element(By.XPATH, '/html/body/div/div[1]/div[1]/aside/nav/div[2]/ul/li[5]/a').click()
time.sleep(5)

recruitment_start_row = time_start_row + len(time_data_list) + 1

admin_ws[f'T{recruitment_start_row}'] = 'Vacancy'
admin_ws[f'U{recruitment_start_row}'] = 'Candidate'
admin_ws[f'V{recruitment_start_row}'] = 'Hiring Manager'
admin_ws[f'W{recruitment_start_row}'] = 'Date of Application'
admin_ws[f'X{recruitment_start_row}'] = 'Status'

recruitment_rows = driver.find_elements(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div')
recruitment_index = recruitment_start_row + 1
recruitment_data_list = []

for i in range(1, len(recruitment_rows)+1):
    data = {}
    data['Vacancy'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div[' +str(i)+ ']/div/div[2]').text
    data['Candidate'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div[' +str(i)+ ']/div/div[3]').text
    data['Hm'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div[' +str(i)+ ']/div/div[4]').text
    data['Doa'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div[' +str(i)+ ']/div/div[5]').text
    data['Status'] = driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div[2]/div[3]/div/div[2]/div[' +str(i)+ ']/div/div[6]').text
    recruitment_data_list.append(data)

    admin_ws[f'T{recruitment_index}'] = data['Vacancy']
    admin_ws[f'U{recruitment_index}'] = data['Candidate']
    admin_ws[f'V{recruitment_index}'] = data['Hm']
    admin_ws[f'W{recruitment_index}'] = data['Doa']
    admin_ws[f'X{recruitment_index}'] = data['Status']

    recruitment_index += 1

driver.find_element(By.XPATH, '/html/body/div/div[1]/div[1]/aside/nav/div[2]/ul/li[6]/a').click()
time.sleep(3)
driver.find_element(By.XPATH, '/html/body/div/div[1]/div[2]/div[2]/div/div/div/div[1]/div[2]/div[1]/a').click()
time.sleep(3)


if os.path.exists("admin_data.xlsx"):
    os.remove("admin_data.xlsx")

wb.save("admin_data.xlsx")
driver.quit()






