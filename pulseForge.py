import pandas as pd
import streamlit as st
from datetime import datetime, time
import requests
import zipfile
import os
import tempfile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager
import time

# ---------------- Firefox Download Options ----------------
def set_firefox_download_options(options):
    options.set_preference("browser.download.folderList", 2)
    options.set_preference("browser.download.manager.showWhenStarting", False)
    options.set_preference("browser.download.dir", "C:\\Downloads")  # Change if needed
    options.set_preference("browser.helperApps.neverAsk.saveToDisk", 
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/zip")
    options.set_preference("browser.download.manager.useWindow", False)
    options.set_preference("browser.download.manager.showAlertOnComplete", False)
    options.set_preference("browser.download.manager.closeWhenDone", True)
    return options

# ---------------- Date Picker Utility ----------------
def select_date(wait, driver, input_xpath, date_obj, label):
    st.write(f"📅 Selecting {label} ({date_obj.strftime('%d-%m-%Y')})...")

    input_field = wait.until(EC.element_to_be_clickable((By.XPATH, input_xpath)))
    input_field.click()
    time.sleep(0.5)

    target_month = date_obj.strftime("%B")
    target_year = date_obj.strftime("%Y")
    target_day = str(date_obj.day)

    while True:
        header_element = wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'p-datepicker-title')]")))
        current_header = header_element.text
        st.write(f"📖 Calendar currently showing: {current_header}")

        if target_month in current_header and target_year in current_header:
            break

        prev_btn = driver.find_element(By.XPATH, "//button[contains(@class,'p-datepicker-prev')]")
        prev_btn.click()
        time.sleep(0.4)

    day_elements = driver.find_elements(By.XPATH, f"//td[not(contains(@class,'p-datepicker-other-month'))]//span[text()='{target_day}']")
    st.write(f"✅ Found {len(day_elements)} day elements for {target_day}")
    if len(day_elements) == 0:
        raise Exception(f"No matching day found for {target_day} in current calendar view.")
    day_elements[0].click()
    st.write(f"✅ Clicked {label} successfully.")

# ---------------- Export and Download ----------------
def export_and_download(wait, driver, report_type):
    st.write(f"📤 Clicking Export button for {report_type} and selecting XLSX...")
    export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Export']]")))
    driver.execute_script("arguments[0].click();", export_button)

    xlsx_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'p-overlaypanel-content')]//span[text()='XLSX']")))
    driver.execute_script("arguments[0].click();", xlsx_option)
    st.write(f"✅ {report_type} XLSX export triggered. Waiting for download link popup...")

    try:
        download_link = WebDriverWait(driver, 600).until(
            EC.element_to_be_clickable((By.XPATH, f"//div[contains(@class,'p-toast-summary') and contains(text(),'{report_type.lower()}_report')]"))
        )
        st.write(f"✅ Download link found: {download_link.text}")
        driver.execute_script("arguments[0].click();", download_link)
        st.success(f"✅ {report_type} download link clicked successfully!")
        return True
    except:
        st.error(f"❌ Timeout: {report_type} download link did not appear within 10 mins.")
        return False

# ---------------- Main Report Logic ----------------
def run_report(driver, wait, start_date, end_date, report_type):
    st.write(f"🔍 Running {report_type} report from {start_date.strftime('%d-%m-%Y')} to {end_date.strftime('%d-%m-%Y')}")

    select_date(wait, driver, "//input[@placeholder='Enter Start Date' or @id='startDate']", start_date, "Start Date")
    select_date(wait, driver, "//input[@placeholder='Enter End Date' or @id='endDate']", end_date, "End Date")

    st.write("🔍 Clicking Search button...")
    search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[@class='p-button-label' and text()='Search']")))
    driver.execute_script("arguments[0].click();", search_button)
    time.sleep(5)

    return export_and_download(wait, driver, report_type)

# ---------------- Automation Functions ----------------
def automate_report_download(username, password, specific_date, download_path):
    st.write("🌐 Launching Firefox browser...")
    firefox_options = webdriver.FirefoxOptions()
    firefox_options = set_firefox_download_options(firefox_options)
    
    # Set download path
    firefox_options.set_preference("browser.download.dir", download_path)

    try:
        driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()), options=firefox_options)
        wait = WebDriverWait(driver, 20)

        st.write("🌐 Navigating to RMS website...")
        driver.get("https://rms.eyeelectronics.net/")

        # Login
        st.write("🔐 Logging in...")
        wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Username']"))).send_keys(username)
        driver.find_element(By.XPATH, "//input[@placeholder='Password']").send_keys(password)
        driver.find_element(By.XPATH, "//span[text()='Login']").click()
        time.sleep(3)

        # Navigate to Alarm Report
        st.write("📊 Navigating to Alarm Report...")
        rms_station_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Rms Station']")))
        driver.execute_script("arguments[0].click();", rms_station_btn)

        alarm_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='Alarm']")))
        driver.execute_script("arguments[0].click();", alarm_link)

        # Download Motion Report
        st.subheader("📄 Downloading Motion Report")
        report_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[@class='p-button-label' and text()='Report']")))
        driver.execute_script("arguments[0].click();", report_button)
        
        # Select Motion Option
        dropdown_trigger = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='p-multiselect-trigger']")))
        driver.execute_script("arguments[0].click();", dropdown_trigger)
        option_to_select = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Motion']")))
        driver.execute_script("arguments[0].click();", option_to_select)
        
        motion_success = run_report(driver, wait, specific_date, specific_date, "Motion")
        
        # Wait for download to complete
        time.sleep(10)
        
        # Download Vibration Report
        st.subheader("📄 Downloading Vibration Report")
        
        # Select Vibration Option
        dropdown_trigger = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='p-multiselect-trigger']")))
        driver.execute_script("arguments[0].click();", dropdown_trigger)
        option_to_select = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Vibration']")))
        driver.execute_script("arguments[0].click();", option_to_select)
        
        vibration_success = run_report(driver, wait, specific_date, specific_date, "Vibration")
        
        # Wait for download to complete
        time.sleep(10)
        
        driver.quit()
        
        return motion_success and vibration_success

    except Exception as e:
        st.error(f"❌ Error occurred: {e}")
        import traceback
        st.error(f"Full error details: {traceback.format_exc()}")
        return False

def extract_first_file_from_zip(zip_path, extract_to):
    """Extract the first file from a ZIP archive"""
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        # Get the first file in the ZIP
        first_file = zip_ref.namelist()[0]
        # Extract it
        zip_ref.extract(first_file, extract_to)
        return os.path.join(extract_to, first_file)

# ---------------- PulseForge Functions ----------------
# Load username data from repository (ensure this file is not too large)
try:
    username_df = pd.read_excel("USER NAME.xlsx")
except:
    st.error("USER NAME.xlsx file not found. Please make sure it exists in the same directory.")
    username_df = pd.DataFrame(columns=['Zone', 'Name'])

# Define zone priority order for display
zone_priority = ["Sylhet", "Gazipur", "Shariatpur", "Narayanganj", "Faridpur", "Mymensingh"]

# Function to preprocess report files
def preprocess_report(df, alarm_type):
    df["Type"] = alarm_type  # Specify type as either 'Motion' or 'Vibration'
    df['Start Time'] = pd.to_datetime(df['Start Time'], errors='coerce')
    df['End Time'] = pd.to_datetime(df['End Time'], errors='coerce')
    return df

# Function to merge motion and vibration data from report files
def merge_report_files(report_motion_df, report_vibration_df):
    report_motion_df = preprocess_report(report_motion_df, 'Motion')
    report_vibration_df = preprocess_report(report_vibration_df, 'Vibration')
    
    # Merging the two reports
    merged_df = pd.concat([report_motion_df, report_vibration_df], ignore_index=True)
    return merged_df

# Function to count occurrences of Motion and Vibration events per Site Alias and Zone
def count_entries_by_zone(merged_df, start_time_filter=None):
    if start_time_filter is not None:
        merged_df = merged_df[merged_df['Start Time'] >= start_time_filter]

    motion_count = merged_df[merged_df['Type'] == 'Motion'].groupby(['Zone', 'Site Alias ']).size().reset_index(name='Motion Count')
    vibration_count = merged_df[merged_df['Type'] == 'Vibration'].groupby(['Zone', 'Site Alias ']).size().reset_index(name='Vibration Count')
    
    final_df = pd.merge(motion_count, vibration_count, on=['Zone', 'Site Alias '], how='outer').fillna(0)
    final_df['Motion Count'] = final_df['Motion Count'].astype(int)
    final_df['Vibration Count'] = final_df['Vibration Count'].astype(int)
    
    return final_df

# Styling function to color cells based on counts and theme
def highlight_counts(row):
    theme = "dark" if st.get_option("theme.base") == "dark" else "light"
    styles = []
    for val in [row['Motion Count'], row['Vibration Count']]:
        if val >= 10:
            styles.append(f'background-color: {"#8B0000" if theme == "dark" else "lightcoral"}; color: white;')
        elif val > 0:
            styles.append(f'background-color: {"#505050" if theme == "dark" else "lightgray"};')
        else:
            styles.append('')
    return styles

# Function to render DataFrame as an HTML table with color formatting
def render_styled_table(df):
    styled_df = df.style.apply(lambda row: highlight_counts(row), axis=1, subset=['Motion Count', 'Vibration Count'])
    styled_df = styled_df.set_properties(**{'font-size': '12px', 'padding': '4px'}).hide(axis='index')
    return styled_df.to_html()

# Function to send data to Telegram
def send_to_telegram(message, chat_id, bot_token):
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {
        "chat_id": chat_id,
        "text": message,
        "parse_mode": "HTML"
    }
    response = requests.post(url, data=payload)
    return response.ok

# Function to update the 'USER NAME.xlsx' file with the new concern name
def update_username_file(selected_zone, new_concern):
    # Read the existing data
    username_df = pd.read_excel("USER NAME.xlsx")
    
    # Update the concern name for the selected zone
    username_df.loc[username_df['Zone'] == selected_zone, 'Name'] = new_concern
    
    # Save the updated data back to the file
    username_df.to_excel("USER NAME.xlsx", index=False)

# Streamlit app
st.title('PulseForge')

# Create a temporary directory for downloads and extraction
with tempfile.TemporaryDirectory() as temp_dir:
    download_path = temp_dir
    
    # Automation section
    st.header("🔧 Automated Report Download")
    auto_username = st.text_input("RMS Username", value="akib")
    auto_password = st.text_input("RMS Password", type="password", value="akib123")
    auto_date = st.date_input("Report Date", value=datetime.now().date())
    
    if st.button("Download Reports Automatically"):
        with st.spinner("Downloading reports from RMS..."):
            success = automate_report_download(auto_username, auto_password, auto_date, download_path)
            
            if success:
                st.success("Reports downloaded successfully!")
                
                # Find the downloaded ZIP files
                zip_files = [f for f in os.listdir(download_path) if f.endswith('.zip')]
                
                if len(zip_files) >= 2:
                    # Extract the first file from each ZIP
                    motion_zip = [f for f in zip_files if 'motion' in f.lower()][0]
                    vibration_zip = [f for f in zip_files if 'vibration' in f.lower()][0]
                    
                    motion_file_path = extract_first_file_from_zip(os.path.join(download_path, motion_zip), download_path)
                    vibration_file_path = extract_first_file_from_zip(os.path.join(download_path, vibration_zip), download_path)
                    
                    # Store the file paths in session state for processing
                    st.session_state.motion_file_path = motion_file_path
                    st.session_state.vibration_file_path = vibration_file_path
                    st.session_state.reports_downloaded = True
                    
                    st.success(f"Motion report extracted: {os.path.basename(motion_file_path)}")
                    st.success(f"Vibration report extracted: {os.path.basename(vibration_file_path)}")
                else:
                    st.error("Could not find both motion and vibration ZIP files.")
            else:
                st.error("Failed to download reports.")

    # File upload section (manual fallback)
    st.header("📁 Manual File Upload (Fallback)")
    if 'reports_downloaded' not in st.session_state or not st.session_state.reports_downloaded:
        report_motion_file = st.file_uploader("Upload the Motion Report Data", type=["xlsx"])
        report_vibration_file = st.file_uploader("Upload the Vibration Report Data", type=["xlsx"])
    else:
        # Use the automatically downloaded files
        report_motion_file = st.session_state.motion_file_path
        report_vibration_file = st.session_state.vibration_file_path
        st.info("Using automatically downloaded reports")

    if (report_motion_file is not None and report_vibration_file is not None) and \
       (isinstance(report_motion_file, str) or isinstance(report_vibration_file, str) or \
        (hasattr(report_motion_file, 'name') and hasattr(report_vibration_file, 'name'))):
        
        # Read the files whether they are file paths or file objects
        if isinstance(report_motion_file, str):
            report_motion_df = pd.read_excel(report_motion_file, header=2)
        else:
            report_motion_df = pd.read_excel(report_motion_file, header=2)
            
        if isinstance(report_vibration_file, str):
            report_vibration_df = pd.read_excel(report_vibration_file, header=2)
        else:
            report_vibration_df = pd.read_excel(report_vibration_file, header=2)

        merged_df = merge_report_files(report_motion_df, report_vibration_df)

        # Sidebar options
        with st.sidebar:
            st.header("Notifications")
            
            # Date and time filter
            selected_date = st.date_input("Select Start Date", value=datetime.now().date())
            selected_time = st.time_input("Select Start Time", value=time(0, 0))
            start_time_filter = datetime.combine(selected_date, selected_time)

            # Option to send notifications for prioritized zones
            st.write("### Notifications for Prioritized Zones")
            if st.button("Send to Prioritized Zones"):
                for zone in zone_priority:
                    concern = username_df[username_df['Zone'] == zone]['Name'].values
                    zonal_concern = concern[0] if len(concern) > 0 else "Unknown Concern"
                    zone_df = merged_df[(merged_df['Zone'] == zone) & (merged_df['Start Time'] >= start_time_filter)]
                    if not zone_df.empty:
                        message = "<b>Motion & Vibration Alarm Alert</b>\n\n"
                        message += f"<b>{zone}:</b>\nAlarm came after: {start_time_filter.strftime('%Y-%m-%d %I:%M %p')}\n\n"
                        site_summary = count_entries_by_zone(zone_df, start_time_filter)
                        site_summary['Total Alarm Count'] = site_summary['Motion Count'] + site_summary['Vibration Count']
                        site_summary = site_summary.sort_values(by='Total Alarm Count', ascending=False)
                        for _, row in site_summary.iterrows():
                            message += f"#{row['Site Alias ']}: Vibration: {row['Vibration Count']}, Motion: {row['Motion Count']} \n"
                        message += f"\n@{zonal_concern}, please take care."
                        success = send_to_telegram(message, chat_id="-1001509039244", bot_token="7776982987:AAHrCszGk2lWZ9WNYx7FQraQypMWjgKPG5w")
                        if success:
                            st.sidebar.success(f"Data for {zone} sent to Telegram successfully!")
                        else:
                            st.sidebar.error(f"Failed to send data for {zone} to Telegram.")

            # Option to send notifications for other zones
            st.write("### Notifications for Other Zones")
            additional_zones = st.multiselect(
                "Select Zones for Notifications",
                options=merged_df['Zone'].unique(),
                default=[]
            )
            if st.button("Send to Selected Zones"):
                for zone in additional_zones:
                    concern = username_df[username_df['Zone'] == zone]['Name'].values
                    zonal_concern = concern[0] if len(concern) > 0 else "Unknown Concern"
                    zone_df = merged_df[(merged_df['Zone'] == zone) & (merged_df['Start Time'] >= start_time_filter)]
                    if not zone_df.empty:
                        message = "<b>Motion & Vibration Alarm Alert</b>\n\n"
                        message += f"<b>{zone}:</b>\nAlarm came after: {start_time_filter.strftime('%Y-%m-%d %I:%M %p')}\n\n"

                        site_summary = count_entries_by_zone(zone_df, start_time_filter)
                        site_summary['Total Alarm Count'] = site_summary['Motion Count'] + site_summary['Vibration Count']
                        site_summary = site_summary.sort_values(by='Total Alarm Count', ascending=False)
                        for _, row in site_summary.iterrows():
                            message += f"#{row['Site Alias ']}: Vibration: {row['Vibration Count']}, Motion: {row['Motion Count']} \n"
                        message += f"\n@{zonal_concern}, please take care."
                        success = send_to_telegram(message, chat_id="-4625672098", bot_token="7543963915:AAGWMNVfD6BaCLuSyKAPCJgPGrdN5WyGLbo")
                        if success:
                            st.sidebar.success(f"Data for {zone} sent to Telegram successfully!")
                        else:
                            st.sidebar.error(f"Failed to send data for {zone} to Telegram.")

            # Option to update/add zonal concerns
            st.write("### Add/Remove Zonal Concern")
            selected_zone = st.selectbox("Select Zone", options=username_df['Zone'].unique())
            current_concern = username_df.loc[username_df['Zone'] == selected_zone, 'Name'].values[0]
            new_concern = st.text_input("Edit Zonal Concern", value=current_concern)
            if st.button("Update Concern"):
                update_username_file(selected_zone, new_concern)
                st.sidebar.success("Concern updated successfully!")

        # Filtered summary based on selected time filter
        summary_df = count_entries_by_zone(merged_df, start_time_filter)

        # Separate prioritized and non-prioritized zones
        prioritized_df = summary_df[summary_df['Zone'].isin(zone_priority)]
        non_prioritized_df = summary_df[~summary_df['Zone'].isin(zone_priority)]

        # Sort prioritized zones according to the order in zone_priority
        prioritized_df['Zone'] = pd.Categorical(prioritized_df['Zone'], categories=zone_priority, ordered=True)
        prioritized_df = prioritized_df.sort_values('Zone')

        # Display prioritized zones first, sorted by total motion and vibration counts in descending order
        for zone in prioritized_df['Zone'].unique():
            st.write(f"### {zone}")
            zone_df = prioritized_df[prioritized_df['Zone'] == zone]

            # Sort by total motion and vibration counts (sum of both)
            zone_df['Total Alarm Count'] = zone_df['Motion Count'] + zone_df['Vibration Count']
            zone_df = zone_df.sort_values('Total Alarm Count', ascending=False)

            # Display the total alarm count as in the original format
            total_motion = zone_df['Motion Count'].sum()
            total_vibration = zone_df['Vibration Count'].sum()
            st.write(f"Total Motion Alarm count: {total_motion}")
            st.write(f"Total Vibration Alarm count: {total_vibration}")

            # Render and display the HTML table with color formatting
            styled_table_html = render_styled_table(zone_df[['Site Alias ', 'Motion Count', 'Vibration Count']])
            st.markdown(styled_table_html, unsafe_allow_html=True)

        # Display non-prioritized zones in alphabetical order, sorted by total motion and vibration counts
        for zone in sorted(non_prioritized_df['Zone'].unique()):
            st.write(f"### {zone}")
            zone_df = non_prioritized_df[non_prioritized_df['Zone'] == zone]

            # Sort by total motion and vibration counts (sum of both)
            zone_df['Total Alarm Count'] = zone_df['Motion Count'] + zone_df['Vibration Count']
            zone_df = zone_df.sort_values('Total Alarm Count', ascending=False)

            # Display the total alarm count as in the original format
            total_motion = zone_df['Motion Count'].sum()
            total_vibration = zone_df['Vibration Count'].sum()
            st.write(f"Total Motion Alarm count: {total_motion}")
            st.write(f"Total Vibration Alarm count: {total_vibration}")

            # Render and display the HTML table with color formatting
            styled_table_html = render_styled_table(zone_df[['Site Alias ', 'Motion Count', 'Vibration Count']])
            st.markdown(styled_table_html, unsafe_allow_html=True)
    else:
        st.write("Please upload both Motion and Vibration Report Data files or use the automatic download feature.")
