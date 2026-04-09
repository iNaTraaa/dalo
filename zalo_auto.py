import time
import random
import openpyxl
import pyautogui
import pyperclip
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from openpyxl.styles import PatternFill

# Cấu hình file
EXCEL_PATH = r'C:\Users\keke\Downloads\telesale\test.xlsx'

MESSAGE_TEMPLATE = """Chào em, Thầy đến từ Khoa Công nghệ Thông tin – Trường Đại học Nguyễn Tất Thành đây 🤝

Thầy cô nhắn để hỗ trợ em tư vấn chọn ngành học phù hợp hơn với bản thân. Em chỉ cần trả lời bằng một số trong các lựa chọn sau nhé:

1. Rất quan tâm
2. Quan tâm
3. Đang cân nhắc
4. Muốn nhận thêm thông tin
5. Đã đăng ký / đang làm hồ sơ
6. Phụ huynh muốn cùng trao đổi
7. Chưa quan tâm / đã chọn trường khác
8. Khác: (nội dung em muốn chia sẻ, ví dụ: học phí, chương trình học, cơ hội việc làm, môi trường học tập hay ngành nào phù hợp với em.)

Thầy cô Khoa Công nghệ Thông tin – ĐH Nguyễn Tất Thành luôn sẵn sàng hỗ trợ và đồng hành cùng em!
Mong nhận được câu trả lời sớm từ em nhé!"""


vang_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") 
do_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")   

print("-" * 30)
try:
    start_stt = int(input("Nhập STT bắt đầu (Cột A): "))
    end_stt = int(input("Nhập STT kết thúc (Cột A): "))
except ValueError:
    print("Vui lòng chỉ nhập số nguyên cho STT!")
    exit()

# Khởi tạo trình duyệt
chrome_options = Options()
chrome_options.add_argument(r"user-data-dir=C:\ZaloAutoProfile")
chrome_options.add_argument("--start-maximized")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
driver.get("https://chat.zalo.me")

print("Đang khởi động... Chờ Zalo ổn định trong 15s")
time.sleep(15)

# Đọc file Excel
try:
    wb = openpyxl.load_workbook(EXCEL_PATH)
    sheet = wb.active
except Exception as e:
    print(f"Lỗi file Excel: {e}")
    driver.quit()
    exit()

# Hàm tô màu cả dòng 
def to_mau_dong(row_index, color_fill):
    for col in range(1, 9): 
        sheet.cell(row=row_index, column=col).fill = color_fill

print("-" * 30)
so_luong_da_xu_ly = 0

for row in range(2, sheet.max_row + 1):
    # Lấy STT hiện tại từ cột A
    stt_val = sheet.cell(row=row, column=1).value
    
    if stt_val is None:
        continue
    try:
        current_stt = int(stt_val)
        if not (start_stt <= current_stt <= end_stt):
            continue
    except ValueError:
        continue 

    phone = sheet.cell(row=row, column=4).value # Cột D (SĐT)
    status = sheet.cell(row=row, column=6).value 

    # Bỏ qua nếu đã gửi thành công trước đó
    if status == "Đã gửi" or not phone:
        continue

    str_phone = str(phone).strip()
    print(f"\n[STT: {current_stt}] --> Đang xử lý số: {str_phone}")
    so_luong_da_xu_ly += 1

    try:
        driver.execute_script("window.focus();")
        # tìm
        search_box = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "contact-search-input"))
        )
        search_box.click()
        search_box.clear()
        search_box.send_keys(str_phone)
        time.sleep(3) 

        search_box.send_keys(Keys.ENTER)
        time.sleep(3)

        # Kiểm tra tồn tại 
        is_not_found = False
        error_texts = ["chưa đăng ký", "Không tìm thấy", "không cho phép"]
        for txt in error_texts:
            if driver.find_elements(By.XPATH, f"//*[contains(text(), '{txt}')]"):
                is_not_found = True
                break

        if is_not_found:
            print(f"   => KHÔNG TÌM THẤY")
            sheet.cell(row=row, column=6).value = "Không tìm thấy"
            to_mau_dong(row, do_fill)
            wb.save(EXCEL_PATH)
            pyautogui.press('esc')
            time.sleep(1)

        else:
            # Gửi tin nhắn
            try:
                btns = driver.find_elements(By.XPATH, "//div[contains(text(), 'Nhắn tin')] | //span[contains(text(), 'Nhắn tin')]")
                if btns:
                    btns[0].click()
                    time.sleep(3)
                
                chat_box = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "richInput"))
                )
                
                # Hàm đếm thông báo chặn để so sánh
                def get_block_count():
                    blocks = driver.find_elements(By.XPATH, "//*[contains(text(), 'chặn không nhận tin nhắn từ người lạ')]")
                    return len([el for el in blocks if el.is_displayed()])

                num_before = get_block_count()
                chat_box.click()
                time.sleep(1)

                pyperclip.copy(MESSAGE_TEMPLATE)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(2) 
                pyautogui.press('enter')
                
                time.sleep(4) 
                num_after = get_block_count()
                
                if num_after > num_before:
                    print("   => BỊ CHẶN")
                    sheet.cell(row=row, column=6).value = "Chặn người lạ"
                    to_mau_dong(row, do_fill)
                    wb.save(EXCEL_PATH)
                    pyautogui.press('esc')
                else:
                    print("   => THÀNH CÔNG")
                    sheet.cell(row=row, column=6).value = "Đã gửi" 
                    to_mau_dong(row, vang_fill)
                    wb.save(EXCEL_PATH)

            except TimeoutException:
                print("   => BỊ CHẶN (Không hiện ô chat)")
                sheet.cell(row=row, column=6).value = "Chặn người lạ"
                to_mau_dong(row, do_fill)
                wb.save(EXCEL_PATH)
                pyautogui.press('esc')

    except Exception as e:
        print(f"   => LỖI: {e}")
        sheet.cell(row=row, column=6).value = "Lỗi"
        to_mau_dong(row, do_fill)
        wb.save(EXCEL_PATH)
        pyautogui.press('esc')

    # Nghỉ để tránh spam
    if so_luong_da_xu_ly % 15 == 0:
        print(f"\n[!] Nghỉ 5 phút sau đợt 15 số...")
        time.sleep(300)
    else:
        wait_time = random.randint(20, 35)
        print(f"Nghỉ {wait_time}s...")
        time.sleep(wait_time)

print("-" * 30)
print(f"HOÀN TẤT TỪ STT {start_stt} ĐẾN {end_stt}!")