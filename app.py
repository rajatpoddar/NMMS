import customtkinter as ctk
import webbrowser
import os
import time
import threading
import requests
import platform
import subprocess
from datetime import datetime
from bs4 import BeautifulSoup
from getmac import get_mac_address

# Excel Modules
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
from io import BytesIO
from PIL import Image as PILImage

# Selenium imports
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

import json

# Server URL configuration - can be overridden via env var, config file, or left as default
DEFAULT_SERVER_URL = "https://nmms.palojori.in"
SERVER_URL = os.environ.get('NMMS_SERVER_URL')

if not SERVER_URL:
    # Check for config file
    config_paths = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json'),
        os.path.expanduser('~/.nmms_config.json'),
    ]
    for cfg_path in config_paths:
        if os.path.exists(cfg_path):
            try:
                with open(cfg_path) as f:
                    config = json.load(f)
                    SERVER_URL = config.get('server_url')
                    if SERVER_URL:
                        break
            except Exception:
                pass

if not SERVER_URL:
    SERVER_URL = DEFAULT_SERVER_URL

def open_file_cross_platform(filepath):
    if platform.system() == 'Darwin':       # macOS
        subprocess.call(('open', filepath))
    elif platform.system() == 'Windows':    # Windows
        os.startfile(filepath)
    else:                                   # Linux
        subprocess.call(('xdg-open', filepath))

class AuthManager:
    @staticmethod
    def get_mac():
        return get_mac_address()

    @staticmethod
    def check_license():
        mac = AuthManager.get_mac()
        try:
            res = requests.get(f"{SERVER_URL}/api/check_status?mac={mac}", timeout=5)
            return res.json()
        except Exception:
            return {"error": "Server unreachable. Check your internet or start the server."}

# ==========================================
# Excel Premium Export Module 
# ==========================================
class ExcelGenerator:
    @staticmethod
    def export(district, block, date, data, status_callback, stop_check):
        filename = os.path.abspath(f"NMMS_Report_{block}_{date.replace('/', '-')}.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "NMMS Attendance"

        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        # ========================================
        # 1. PREMIUM HEADER SECTION
        # ========================================
        ws.merge_cells('A1:L1')
        ws['A1'] = "NMMS TRACKING REPORT"
        ws['A1'].font = Font(size=18, bold=True, color="FFFFFF")
        ws['A1'].fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
        ws['A1'].alignment = center_align
        ws.row_dimensions[1].height = 35

        ws.merge_cells('A2:L2')
        ws['A2'] = f"Date: {date}  |  District: {district}  |  Block: {block}"
        ws['A2'].font = Font(size=12, bold=True, color="000000")
        ws['A2'].fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        ws['A2'].alignment = center_align
        ws.row_dimensions[2].height = 25

        ws.append([]) # Empty row 3 for spacing

        # ========================================
        # 2. COLUMNS SETUP
        # ========================================
        headers = ["Sl. No.", "Work Code", "MSR No.", "Work Name", "Worker Name", "Job Card", "Status", "Taken By", "Designation", "Geo Coordinates", "Photo 1", "Photo 2"]
        ws.append(headers)
        header_row_idx = 4

        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Arial", size=12, bold=True, color="FFFFFF")

        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=header_row_idx, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border

        # Adjust Column Widths
        ws.column_dimensions['A'].width = 8   # Sl No
        ws.column_dimensions['B'].width = 25  # Work Code
        ws.column_dimensions['C'].width = 12  # MSR
        ws.column_dimensions['D'].width = 45  # Work Name
        ws.column_dimensions['E'].width = 25  # Worker Name
        ws.column_dimensions['F'].width = 25  # Job Card
        ws.column_dimensions['G'].width = 12  # Status
        ws.column_dimensions['H'].width = 20  # Taken By
        ws.column_dimensions['I'].width = 15  # Designation
        ws.column_dimensions['J'].width = 25  # Geo Coordinates
        ws.column_dimensions['K'].width = 22  # Photo 1
        ws.column_dimensions['L'].width = 22  # Photo 2

        # ========================================
        # 3. DATA POPULATION
        # ========================================
        row_idx = header_row_idx + 1
        sl_no = 1
        
        for item in data:
            if stop_check(): break
            
            # Smart Extraction for Work Name (Separates it from WorkCode and MSR)
            header_text = item.get('header_text', '')
            work_name = header_text.split('Work Name :')[-1].strip() if 'Work Name :' in header_text else header_text
            
            # Format Coordinates properly
            geo_coords = f"P1: {item.get('geo', 'N/A')}"
            if item.get('photo2_geo') and item.get('photo2_geo') != "N/A":
                geo_coords += f"\nP2: {item.get('photo2_geo')}"

            # Download Images
            img1_obj, img2_obj = None, None
            if item.get('photo_url'):
                try:
                    res = requests.get(item['photo_url'])
                    img = PILImage.open(BytesIO(res.content))
                    img.thumbnail((140, 140)) 
                    img_path = f"temp_img1_{item['msr_no']}.png"
                    img.save(img_path)
                    img1_obj = img_path
                except: pass

            if item.get('photo2_url'):
                try:
                    res = requests.get(item['photo2_url'])
                    img = PILImage.open(BytesIO(res.content))
                    img.thumbnail((140, 140))
                    img_path = f"temp_img2_{item['msr_no']}.png"
                    img.save(img_path)
                    img2_obj = img_path
                except: pass

            # Write Worker Data
            for worker in item.get('workers', []):
                status_callback(f"📊 Exporting row {sl_no} to Excel...")
                
                ws.cell(row=row_idx, column=1, value=sl_no).alignment = center_align
                ws.cell(row=row_idx, column=2, value=item.get('work_code', '')).alignment = center_align
                ws.cell(row=row_idx, column=3, value=item.get('msr_no', '')).alignment = center_align
                ws.cell(row=row_idx, column=4, value=work_name).alignment = center_align
                ws.cell(row=row_idx, column=5, value=worker['name']).alignment = center_align
                ws.cell(row=row_idx, column=6, value=worker['jobcard']).alignment = center_align
                
                status_cell = ws.cell(row=row_idx, column=7, value=worker['status'])
                status_cell.alignment = center_align
                status_cell.font = Font(bold=True, color="00B050" if "present" in worker['status'].lower() else "FF0000")
                
                ws.cell(row=row_idx, column=8, value=item.get('taken_by', 'N/A')).alignment = center_align
                ws.cell(row=row_idx, column=9, value=item.get('designation', 'N/A')).alignment = center_align
                ws.cell(row=row_idx, column=10, value=geo_coords).alignment = center_align

                ws.row_dimensions[row_idx].height = 100 
                
                if img1_obj:
                    ws.add_image(OpenpyxlImage(img1_obj), f"K{row_idx}")
                if img2_obj:
                    ws.add_image(OpenpyxlImage(img2_obj), f"L{row_idx}")

                for col in range(1, 13):
                    ws.cell(row=row_idx, column=col).border = thin_border

                row_idx += 1
                sl_no += 1
                
        # ========================================
        # 4. PREMIUM FOOTER SECTION
        # ========================================
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=12)
        footer_cell = ws.cell(row=row_idx, column=1)
        footer_cell.value = "🚀 Report generated by Nrega Bot NMMS Tracker app"
        footer_cell.font = Font(italic=True, color="595959", bold=True, size=11)
        footer_cell.alignment = Alignment(horizontal="center", vertical="center")
        footer_cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        ws.row_dimensions[row_idx].height = 25
        
        for col in range(1, 13):
            ws.cell(row=row_idx, column=col).border = thin_border

        wb.save(filename)
        
        # Cleanup temp images
        for f in os.listdir():
            if f.startswith("temp_img") and f.endswith(".png"):
                try: os.remove(f)
                except: pass
                
        return filename

# ==========================================
# Main GUI Client
# ==========================================
ctk.set_appearance_mode("System")  
ctk.set_default_color_theme("blue")  

class NMMSApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("NMMS Tracking Report")
        self.geometry("600x650")
        self.mac_id = AuthManager.get_mac()
        self.user_info = {}
        self.stop_scraping = False
        
        self.check_activation()

    def check_activation(self):
        for widget in self.winfo_children(): widget.destroy()
            
        lbl = ctk.CTkLabel(self, text="Verifying License with Server...", font=("Segoe UI", 18))
        lbl.pack(pady=250)
        self.update()
        
        auth_data = AuthManager.check_license()
        
        if "error" in auth_data:
            lbl.configure(text=f"❌ {auth_data['error']}", text_color="red")
            return
            
        if not auth_data.get("registered"):
            self.show_registration_screen()
        elif not auth_data.get("active"):
            self.show_expired_screen()
        else:
            self.user_info = auth_data.get("user_data", {})
            self.days_left = auth_data.get("days_left", 0)
            self.build_main_ui()

    def show_registration_screen(self):
        for widget in self.winfo_children(): widget.destroy()
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(pady=100, padx=40, fill="both", expand=True)
        ctk.CTkLabel(frame, text="Welcome to NMMS Tracker!", font=("Segoe UI", 28, "bold")).pack(pady=10)
        ctk.CTkLabel(frame, text="Your device is not registered on the network.", text_color="gray", font=("Segoe UI", 14)).pack()
        
        def open_browser(): webbrowser.open(f"{SERVER_URL}/register?mac={self.mac_id}")
            
        ctk.CTkButton(frame, text="Open Registration Web Portal", command=open_browser, height=50, font=("Segoe UI", 15, "bold")).pack(pady=30, fill="x")
        ctk.CTkButton(frame, text="I have Registered, Refresh Status", command=self.check_activation, fg_color="#28a745", hover_color="#218838", height=40).pack(fill="x")

    def show_expired_screen(self):
        for widget in self.winfo_children(): widget.destroy()
        ctk.CTkLabel(self, text="⚠️ Subscription Expired", font=("Segoe UI", 26, "bold"), text_color="#dc3545").pack(pady=(200,10))
        ctk.CTkLabel(self, text="Your 30-day trial or subscription has ended.\nPlease contact the administrator to renew.", font=("Segoe UI", 14)).pack()

    def build_main_ui(self):
        for widget in self.winfo_children(): widget.destroy()
        
        # --- Top Header (Main) ---
        header_frame = ctk.CTkFrame(self, fg_color="#1e3c72", corner_radius=0)
        header_frame.pack(fill="x")
        ctk.CTkLabel(header_frame, text="NMMS Tracking Report", font=("Segoe UI", 24, "bold"), text_color="white").pack(pady=(15, 2))
        status_text = f"🟢 Connected | Welcome, {self.user_info.get('name', 'User')} | ⏳ {self.days_left} Days Left"
        ctk.CTkLabel(header_frame, text=status_text, text_color="#a8e6cf", font=("Segoe UI", 12, "bold")).pack(pady=(0, 15))

        # --- Secondary Header (Location Ribbon) ---
        self.location_ribbon = ctk.CTkFrame(self, fg_color="#2a5298", corner_radius=0)
        self.location_ribbon.pack(fill="x")
        
        self.loc_label = ctk.CTkLabel(self.location_ribbon, 
            text=f"📍 {self.user_info.get('state', '')}  |  {self.user_info.get('district', '')}  |  {self.user_info.get('block', '')}", 
            font=("Segoe UI", 14, "bold"), text_color="white")
        self.loc_label.pack(side="left", padx=20, pady=10)
        
        ctk.CTkButton(self.location_ribbon, text="✏️ Edit", width=60, height=28, fg_color="#ff9800", hover_color="#e68a00", 
                      text_color="black", font=("Segoe UI", 12, "bold"), command=self.open_edit_popup).pack(side="right", padx=20, pady=10)

        # --- Dashboard Body ---
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(pady=20, padx=40, fill="both", expand=True)

        ctk.CTkLabel(body, text="Reporting Date (DD/MM/YYYY):", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(10, 5))
        self.date_entry = ctk.CTkEntry(body, height=45, font=("Segoe UI", 16))
        self.date_entry.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.date_entry.pack(fill="x", pady=5)
        
        # --- Start / Stop Buttons Frame ---
        btn_frame = ctk.CTkFrame(body, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(30, 10))

        self.start_btn = ctk.CTkButton(btn_frame, text="▶ Start Extraction", height=50, font=("Segoe UI", 16, "bold"), 
                                       fg_color="#007bff", hover_color="#0056b3", command=self.start_scraping_thread)
        self.start_btn.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.stop_btn = ctk.CTkButton(btn_frame, text="⏹ Stop", height=50, font=("Segoe UI", 16, "bold"), 
                                      fg_color="#dc3545", hover_color="#c82333", state="disabled", command=self.stop_scraping_action)
        self.stop_btn.pack(side="right", fill="x", expand=True, padx=(5, 0))
        
        # --- Status Box ---
        self.status_box = ctk.CTkFrame(body, fg_color="#f8f9fa", border_width=1, border_color="#dee2e6")
        self.status_box.pack(fill="x", pady=20)
        self.status_label = ctk.CTkLabel(self.status_box, text="System Ready.", text_color="#495057", font=("Segoe UI", 14, "bold"))
        self.status_label.pack(pady=15)

        # --- Footer ---
        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.pack(side="bottom", fill="x", pady=10)
        ctk.CTkLabel(footer, text="Developed by Nrega Bot Team.", text_color="gray", font=("Segoe UI", 11)).pack()

    # --- Edit Popup Form ---
    def open_edit_popup(self):
        popup = ctk.CTkToplevel(self)
        popup.title("Edit Configuration")
        popup.geometry("400x350")
        popup.transient(self) # Keep on top
        
        ctk.CTkLabel(popup, text="Temporary Configuration", font=("Segoe UI", 18, "bold")).pack(pady=15)
        
        s_entry = ctk.CTkEntry(popup, placeholder_text="State", height=40)
        s_entry.insert(0, self.user_info.get('state', ''))
        s_entry.pack(fill="x", padx=30, pady=10)

        d_entry = ctk.CTkEntry(popup, placeholder_text="District", height=40)
        d_entry.insert(0, self.user_info.get('district', ''))
        d_entry.pack(fill="x", padx=30, pady=10)

        b_entry = ctk.CTkEntry(popup, placeholder_text="Block", height=40)
        b_entry.insert(0, self.user_info.get('block', ''))
        b_entry.pack(fill="x", padx=30, pady=10)

        def save_edits():
            self.user_info['state'] = s_entry.get().strip().upper()
            self.user_info['district'] = d_entry.get().strip().upper()
            self.user_info['block'] = b_entry.get().strip().upper()
            
            self.loc_label.configure(text=f"📍 {self.user_info['state']}  |  {self.user_info['district']}  |  {self.user_info['block']}")
            self.log_status("✅ Configuration Updated for this session.", "#28a745")
            popup.destroy()

        ctk.CTkButton(popup, text="Save Details", height=40, font=("Segoe UI", 14, "bold"), command=save_edits).pack(pady=20)

    def log_status(self, text, color="#007bff"):
        self.status_label.configure(text=text, text_color=color)
        self.update()

    def stop_scraping_action(self):
        self.stop_scraping = True
        self.log_status("⚠️ Stopping process... Please wait.", "#dc3545")
        self.stop_btn.configure(state="disabled")

    def start_scraping_thread(self):
        s, d, b, date = self.user_info.get('state'), self.user_info.get('district'), self.user_info.get('block'), self.date_entry.get().strip()
        if not s or not d or not b or not date:
            self.log_status("⚠️ Please set location details and date!", "#ff4757")
            return
            
        self.stop_scraping = False
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        threading.Thread(target=self.run_scraper, args=(s, d, b, date), daemon=True).start()

    # --- Core Scraper ---
    def run_scraper(self, state, district, block, date):
        driver = None
        try:
            self.log_status("🚀 Engine Started (Headless Mode)...")
            
            chrome_options = Options()
            chrome_options.add_argument("--headless=new")
            chrome_options.add_argument("--window-size=1920,1080")
            chrome_options.add_argument("--log-level=3")
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
            wait = WebDriverWait(driver, 15)

            if self.stop_scraping: raise Exception("Process Cancelled by User.")

            self.log_status("🌐 Connecting to Server...")
            driver.get("https://mnregaweb4.dord.gov.in/netnrega/View_NMMS_atten_date_new.aspx?fin_year=2025-2026&&digest=HNrisV4bhHnb7Gve3mAKYQ")
            time.sleep(2)

            Select(wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_ddlstate")))).select_by_visible_text(state)
            time.sleep(3)
            Select(wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_ddl_attendance")))).select_by_visible_text(date)
            time.sleep(3)
            wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btn_showreport"))).click()
            time.sleep(3)

            if self.stop_scraping: raise Exception("Process Cancelled by User.")

            self.log_status("📍 Navigating Locations...")
            state_link = wait.until(EC.presence_of_element_located((By.XPATH, f"//a[translate(normalize-space(text()), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = '{state}']")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'}); arguments[0].click();", state_link) 
            time.sleep(3)
            
            dist_link = wait.until(EC.presence_of_element_located((By.XPATH, f"//a[translate(normalize-space(text()), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = '{district}']")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'}); arguments[0].click();", dist_link) 
            time.sleep(3)

            block_row = wait.until(EC.presence_of_element_located((By.XPATH, f"//tr[td[a[translate(normalize-space(text()), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = '{block}'] or translate(normalize-space(text()), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ') = '{block}']]")))
            mr_count_link = block_row.find_element(By.XPATH, "./td[4]/a")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'}); arguments[0].click();", mr_count_link) 
            time.sleep(3)

            self.log_status("📄 Compiling Muster Roll Base...")
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            mr_list = []
            for row in soup.find('div', id='RepPr1').find('table').find_all('tr'):
                cols = row.find_all('td')
                if len(cols) >= 6 and cols[5].find('a'):
                    href = cols[5].find('a')['href']
                    mr_list.append({
                        "work_code": cols[4].text.strip(), "msr_no": cols[5].find('a').text.strip(),
                        "mr_url": href if href.startswith('http') else "https://mnregaweb4.dord.gov.in/netnrega/" + href.lstrip('/')
                    })

            final_data = []
            for idx, mr in enumerate(mr_list):
                if self.stop_scraping: raise Exception("Process Cancelled by User.")
                
                self.log_status(f"🔍 Deep Scraping MR {idx+1}/{len(mr_list)}...")
                driver.get(mr['mr_url'])
                time.sleep(1.5)
                
                s = BeautifulSoup(driver.page_source, 'html.parser')
                
                def get_text(element_id):
                    el = s.find('span', id=element_id)
                    return el.text.strip() if el else "N/A"

                mr.update({
                    "header_text": get_text('ctl00_ContentPlaceHolder1_lbl_dtl'),
                    "taken_by": get_text('ctl00_ContentPlaceHolder1_lbl_Taken_by'),
                    "designation": get_text('ctl00_ContentPlaceHolder1_lbl_Designation'),
                    "geo": get_text('ctl00_ContentPlaceHolder1_lbl_cordinates'),
                    "photo2_geo": get_text('ctl00_ContentPlaceHolder1_lbl_SecondCordinates'),
                    "workers": [{"sno": r.find_all('td')[0].text.strip(), "jobcard": r.find_all('td')[1].text.strip(), "name": r.find_all('td')[2].text.strip(), "status": r.find_all('td')[4].text.strip()} for r in s.find('table', id='ctl00_ContentPlaceHolder1_Gridviewattandance').find_all('tr')[1:]] if s.find('table', id='ctl00_ContentPlaceHolder1_Gridviewattandance') else []
                })
                
                img1 = s.find('img', id='ctl00_ContentPlaceHolder1_img_groupPhoto')
                img2 = s.find('img', id='ctl00_ContentPlaceHolder1_img_SecondGroupPhoto')
                if img1: mr['photo_url'] = img1.get('src') if img1.get('src').startswith('http') else "https://mnregaweb4.dord.gov.in/netnrega/" + img1.get('src').lstrip('/')
                if img2: mr['photo2_url'] = img2.get('src') if img2.get('src').startswith('http') else "https://mnregaweb4.dord.gov.in/netnrega/" + img2.get('src').lstrip('/')
                
                final_data.append(mr)

            driver.quit()
            
            if not self.stop_scraping:
                self.log_status("🖨️ Designing Excel Layout & Rendering Photos...")
                excel_file = ExcelGenerator.export(district, block, date, final_data, self.log_status, lambda: self.stop_scraping)
                
                if not self.stop_scraping:
                    self.log_status("✅ Success! Opening Excel File...", "#28a745")
                    open_file_cross_platform(excel_file)
            
        except Exception as e:
            if "Cancelled" in str(e):
                self.log_status("⛔ Scraper Stopped by User.", "#dc3545")
            else:
                self.log_status(f"❌ Error: {str(e)}", "#dc3545")
            if driver: driver.quit()
        finally:
            self.start_btn.configure(state="normal")
            self.stop_btn.configure(state="disabled")

if __name__ == "__main__":
    app = NMMSApp()
    app.mainloop()