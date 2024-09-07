# Nhớ pip install selenium trước khi chạy code
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Đường dẫn đến WebDriver 
# Tải webdriver của Edge tại https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/
# Dùng trình duyệt khác thì Google nha. Nhớ kiểm tra tính tương thích của WebDriver với trình duyệt đang dùng
# Giải nén và thay đường dẫn vào đây
msdriver = "D:/Downloads/edgedriver_win64/msedgedriver.exe"

# Các phạm vi của ID sinh viên
studentIdList = list(range(24120001, 24122000))
# Tốc độ lưu file, mỗi 10 sinh viên lưu 1 lần
SAVE_RATE = 10
# Phải đăng nhập bằng tài khoản trường từ trước nha
YOUR_STUDENT_ID = "21127469"

# Đường dẫn đến file CSV
filePath = 'NN.csv'

# Khởi tạo WebDriver
service = Service(executable_path=msdriver)
driver = webdriver.Edge(service=service)
wait = WebDriverWait(driver, 10)

def getDataOfStudent(studentId):
    # Chờ và nhập ID vào ô tìm kiếm
    inputElement = wait.until(
        EC.element_to_be_clickable((By.ID, "SearchBox7"))
    )
    try:
        # Nhập văn bản và nhấn Enter
        inputElement.send_keys(f"{studentId}")
        inputElement.send_keys(Keys.ENTER)

        # Chờ và tìm phần tử chứa thông tin
        accountElement = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".DXpY5"))
        )

        # Trích xuất thông tin từ phần tử
        full_text = accountElement.get_attribute("aria-label")
        datum = full_text.split(", ")
        
        return [studentId, datum[1], datum[0]]
    
    except Exception as e:
        print(f"Error for student ID {studentId}: {e}")
        return [studentId, "NOT FOUND", "NOT FOUND"]
    finally:
        # Xóa nội dung ô tìm kiếm
        inputElement.send_keys(Keys.CONTROL + "a")
        inputElement.send_keys(Keys.BACKSPACE)

# Mở trang web
driver.get("https://outlook.office.com/people/")

# Đăng nhập bằng tài khoản trường
element = wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[@aria-label='Sign in with {YOUR_STUDENT_ID}@student.hcmus.edu.vn work or school account.']")))
element.click()

# Mở file CSV và ghi dữ liệu vào file
with open(filePath, mode='w', newline='', encoding='utf-8') as file:
    file.write("StudentID,Name,Email\n")    
    # Tìm kiếm và ghi dữ liệu cho từng ID
    for studentId in studentIdList:
        result = getDataOfStudent(studentId)
        if result[1] == "NOT FOUND":
            continue
        file.write(f"{result[0]},{result[1]},{result[2]}\n")
        if int(studentId) % 10 == 0:
            file.flush()

driver.quit()

# Chạy bằng cách run terminal: python CrawlStudentData.py