from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import UnexpectedAlertPresentException, NoAlertPresentException, NoSuchElementException
 
import pyautogui as p,openpyxl as xl,time
import urllib.request
import pytesseract

from PIL import Image, ImageEnhance, ImageFilter

pytesseract.pytesseract.tesseract_cmd = 'C:\Program Files\Tesseract-OCR/tesseract.exe'

'''
def is_alert_present(driver):
    current_frame = None
    try:
        current_frame = driver.current_window_handle
        a = driver.switch_to_alert()
        a.text
    except NoAlertPresentException:
            # No alert
        return False
    except UnexpectedAlertPresentException:
            # Alert exists
        return True
    finally:
        if current_frame:
            driver.switch_to_window(current_frame)
    return True 
'''


def fill(i):
    roll=w.find_element_by_id("ctl00_ContentPlaceHolder1_txtrollno")
    roll.clear()
    roll.send_keys(i)
    
    Select(w.find_element_by_id("ctl00_ContentPlaceHolder1_drpSemester")).select_by_index(semester-1)
    aw = 1
    
    while aw:
    
        try:
            images = w.find_elements_by_tag_name('img')
            img_url = images[1].get_attribute('src')
            urllib.request.urlretrieve(img_url, 'a.jpeg')
            text = w.find_element_by_id("ctl00_ContentPlaceHolder1_TextBox1")
            text.clear()
            time.sleep(2)
            te = pytesseract.image_to_string(Image.open('a.jpeg')).replace(" ","")  #.upper().replace
            text.send_keys(te)
            time.sleep(2)
            w.find_element_by_id("ctl00_ContentPlaceHolder1_btnviewresult").click()
            aw = 0
    #        current_frame = driver.current_window_handle
    #        a = w.switch_to.alert
    #        a.text
#            print(a)
        except NoAlertPresentException:
                # No alert
            pass
        except UnexpectedAlertPresentException:
            aw = 1
    #        time.sleep(2)
#        fill(i)
#        time.sleep(2)
            # Alert exists
        
#    finally:
#        if current_frame:
#            driver.switch_to_window(current_frame)
        except NoSuchElementException:
            return True 


def getdata():
    details=[]
    details.append(w.find_element_by_id("ctl00_ContentPlaceHolder1_lblRollNoGrading").text)
    details.append(w.find_element_by_id("ctl00_ContentPlaceHolder1_lblNameGrading").text)
    details.append(w.find_element_by_id("ctl00_ContentPlaceHolder1_lblProgramGrading").text)
    details.append(w.find_element_by_id("ctl00_ContentPlaceHolder1_lblBranchGrading").text)
    details.append(w.find_element_by_id("ctl00_ContentPlaceHolder1_lblSemesterGrading").text)
    details.append(w.find_element_by_id("ctl00_ContentPlaceHolder1_lblResultNewGrading").text)
    details.append(w.find_element_by_id("ctl00_ContentPlaceHolder1_lblSGPA").text)
    details.append(w.find_element_by_id("ctl00_ContentPlaceHolder1_lblcgpa").text)
    writedata(details)

def writedata(details):
    for j in range(1,sheet.max_column+1):
        sheet.cell(row=i,column=j).value=details[j-1]
        
#def write_roll(fr,tr):
#    num = int(fr[8:])
#    l = int(tr[8:])
#    c = fr[:8]
#    for j in range(2,l-num+3):
#        sheet.cell(row=j,column=1).value=c + str(num)
#        num = num + 1

# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
w = webdriver.Chrome(executable_path='chromedriver.exe')
#w.maximize_window()
w.get("http://result.rgpv.ac.in/Result/ProgramSelect.aspx")
w.find_element_by_id("radlstProgram_1").click()
path="demo.xlsx"

from_rol=p.prompt("Enter first roll number(in format 0808cs171068)")
to_rol=p.prompt("Enter last roll number ")
semester=int(p.prompt("Enter the semester here"))
#from_rol = "0808cs171061"
#to_rol = "0808cs171090"
#semester = 6

wb=xl.load_workbook(path)
sheet=wb.active

#write_roll(from_rol,to_rol)

print('Your script is running...')

sheet.cell(row=1,column=1).value='Roll No.'
sheet.cell(row=1,column=2).value='Name'
sheet.cell(row=1,column=3).value='Course'
sheet.cell(row=1,column=4).value='Branch'
sheet.cell(row=1,column=5).value='Semester'
sheet.cell(row=1,column=6).value='Result Des.'
sheet.cell(row=1,column=7).value='SGPA'
sheet.cell(row=1,column=8).value='CGPA'

print('Excel sheet has been initialised...')

i=2
num = int(from_rol[8:])
l = int(to_rol[8:])
c = from_rol[:8]
for j in range(2,l-num+3):
    rol=c + str(num)
    num = num + 1
    fill(rol)
    getdata()
    w.find_element_by_id("ctl00_ContentPlaceHolder1_btnReset").click()
    i=i+1
wb.save(path)


print('Task completed...')
w.close()


#selenium.common.exceptions.UnexpectedAlertPresentException