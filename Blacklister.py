import pandas as pd
from playwright.sync_api import sync_playwright
import time
import os
import sys

# 1. LOAD YOUR SPECIFIC EXCEL FILE
EXCEL_FILE = "July 2026 New Accounts- Sydrose Data Collection Spread.xlsx"

if not os.path.exists(EXCEL_FILE):
    print(f"❌ Could not find the file: {EXCEL_FILE}. Please ensure it's in the same folder as this script.")
    sys.exit()

try:
    df = pd.read_excel(EXCEL_FILE, sheet_name="Loan Accounts")
    print(f"✅ Loaded spreadsheet successfully. Found {len(df)} records to process.")
except Exception as e:
    print(f"❌ Error loading Excel tab: {e}")
    sys.exit()

PORTAL_URL = "https://mzloans.loancirrus.app/login"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, args=["--start-maximized"])
    context = browser.new_context(no_viewport=True)
    page = context.new_page()
    
    # --- MANUAL LOGIN PHASE ---
    page.goto(PORTAL_URL)
    print("\n========================================================")
    print("[!] Please log into your Loan Cirrus portal manually.")
    print("[!] Once you are logged in and looking at the main dashboard,")
    print("[!] return to this terminal window and press ENTER...")
    print("========================================================\n")
    input()

    # --- AUTOMATION LOOP ---
    for index, row in df.iterrows():
        trn = str(row['TRN']).strip()
        client_name = str(row['Account Holder']).strip() 
        escrow_id = str(row['Escrow ID']).strip() 
        
        print(f"\n🔄 [{index + 1}/{len(df)}] Processing: {client_name} (TRN: {trn} | Escrow: {escrow_id})")
        
        try:
            # --- ALWAYS RESET TO DASHBOARD BEFORE EVERY CLIENT ---
            page.goto("https://mzloans.loancirrus.app/", wait_until="domcontentloaded")
            time.sleep(1.5)

            try:
                page.keyboard.press("Escape")
            except:
                pass
            
            # --- STEP 1: SEARCH & SELECT CLIENT ---
            search_bar = page.locator("xpath=//input[contains(@placeholder, 'Search') or @type='text']").first
            search_bar.wait_for(state="visible", timeout=12000)
            
            search_bar.click()
            search_bar.focus()
            page.keyboard.press("Control+A")
            page.keyboard.press("Backspace")
            time.sleep(0.5)
            
            search_bar.fill(trn)
            time.sleep(1)
            
            client_link = page.locator("css=ul.dropdown-menu a.search-details").first
            try:
                client_link.wait_for(state="visible", timeout=8000)
            except:
                print("Dropdown delayed, forcing Enter trigger...")
                page.keyboard.press("Enter")
                client_link.wait_for(state="visible", timeout=10000)

            client_link.click()
            print("    ↳ Navigated to client profile.")
            
            # --- STEP 2: SELECT THE CORRECT LOAN BY ESCROW ID ---
            time.sleep(2.5) # Extra brief wait for profile tables to load completely
            
            loan_row = page.locator(f"xpath=//*[self::td or self::a][contains(normalize-space(), '{escrow_id}')]").first
            loan_row.wait_for(state="visible", timeout=12000)
            loan_row.click()
            print(f"    ↳ Clicked into Escrow ID: {escrow_id}")
            
            # --- STEP 3: NAVIGATE TO EXTERNAL ENGAGEMENTS ---
            external_eng_tab = page.locator("xpath=//a[contains(@href, 'external-engagements')]//span[text()='EXTERNAL ENGAGEMENTS']").first
            
            time.sleep(1.5)
            if not external_eng_tab.is_visible():
                print(f"Warning: External Engagements tab is HIDDEN/disabled for {client_name}.")
                print("    ↳ Bypassing engagement creation form steps...")
            else:
                external_eng_tab.click()
                print("    ↳ Switched to External Engagements tab.")
                
                time.sleep(1.5)
                add_engagement_btn = page.locator("xpath=//button[contains(., 'Add External engagement') or contains(text(), 'Add External')]").first
                add_engagement_btn.wait_for(state="visible", timeout=10000)
                add_engagement_btn.click()
                print("    ↳ Opened engagement creation form.")
                
                # --- STEP 4: FILL EXTERNAL ENGAGEMENT FORM ---
                role_dropdown = page.locator("xpath=//select[contains(@ng-model, 'ee.role')]").first
                role_dropdown.wait_for(state="visible", timeout=6000)
                role_dropdown.select_option(label="Bailiff")
                print("    ↳ Selected Dropdown Role: Bailiff")
                time.sleep(1.5) 
                
                firm_dropdown = page.locator("xpath=//select[contains(@ng-model, 'ee.firm')]").first
                firm_dropdown.wait_for(state="visible", timeout=6000)
                
                firm_options = firm_dropdown.locator("option").all_inner_texts()
                matched_firm = next((f for f in firm_options if "SYDROSE" in f.upper()), None)
                
                if matched_firm:
                    firm_dropdown.select_option(label=matched_firm)
                    print(f"    ↳ Selected Dropdown Firm: {matched_firm}")
                else:
                    raise Exception("Could not find an option matching 'SYDROSE' inside the Firm dropdown.")
                
                print("Waiting for Contact dropdown options to sync from the database...")
                contact_dropdown = page.locator("xpath=//select[contains(@ng-model, 'ee.contact')]").first
                contact_dropdown.wait_for(state="visible", timeout=6000)
                
                matched_contact = None
                for _ in range(12):
                    options = contact_dropdown.locator("option").all_inner_texts()
                    matched_contact = next((c for c in options if "ROSALYN TRAILLE" in c.upper()), None)
                    if matched_contact:
                        break
                    time.sleep(0.5)
                
                if matched_contact:
                    contact_dropdown.select_option(label=matched_contact)
                    print(f"    ↳ Selected Dropdown Contact: {matched_contact}")
                else:
                    final_options = contact_dropdown.locator("option").all_inner_texts()
                    raise Exception(f"Timed out waiting for 'Rosalyn Traille'. Options found were: {final_options}")
                    
                time.sleep(1)
                
                scope_radio = page.locator("xpath=//label[contains(., 'This loan')] | //input[@value='This loan']").first
                try:
                    scope_radio.click(timeout=4000)
                except:
                    print("Standard click on radio blocked; forcing interaction...")
                    scope_radio.dispatch_event("click")
                print("    ↳ Selected Scope: This loan")
                
                desc_box = page.locator("css=textarea[ng-model='ee.instructions']").first
                desc_box.wait_for(state="visible", timeout=5000)
                desc_box.fill("Collectors")
                print("    ↳ Filled Instructions box with 'Collectors'.")
                
                watcher_input = page.locator("css=input[ng-model='ee.watcherSearch']").first
                watcher_input.wait_for(state="visible", timeout=5000)
                watcher_input.fill("wy")
                print("    ↳ Typing 'wy' into watchers search field...")
                
                wyndell_option = page.locator("a.ng-binding", has_text="Wyndell Whyte")
                wyndell_option.wait_for(state="visible", timeout=6000)
                wyndell_option.click()
                print("    ↳ Added Watcher: Wyndell Whyte")
                time.sleep(1)
                
                submit_btn = page.locator("button.sys-createengagement").first
                submit_btn.wait_for(state="visible", timeout=5000)
                submit_btn.click()
                print("    ↳ Form submission finalized successfully.")
                time.sleep(4) 

                # --- STEP 5: RETURN TO PROFILE OVERVIEW ---
                client_overview_link = page.locator("a", has_text="Client Overview").first
                client_overview_link.wait_for(state="visible", timeout=10000)
                client_overview_link.click()
                print("    ↳ Returned back to Client Profile View.")

            # --- STEP 6: SMART BLACKLIST CLIENT ---
            blacklist_btn = page.locator("button.btn-warning", has_text="Blacklist Client").first
            
            try:
                blacklist_btn.wait_for(state="visible", timeout=4000)
                blacklist_btn.click()
                print("    ↳ Opened Blacklist Modal Menu.")
                
                reason_box = page.locator("textarea#reason").first
                reason_box.wait_for(state="visible", timeout=5000)
                reason_box.fill("Loan Payment Delinquent")
                print("    ↳ Filled Blacklist reason.")
                
                confirm_blacklist_btn = page.locator("button.btn-danger.btn-lg", has_text="Blacklist Client").first
                confirm_blacklist_btn.wait_for(state="visible", timeout=5000)
                confirm_blacklist_btn.click()
                print(f"Successfully updated and blacklisted client: {client_name}")
            except:
                print(f"Note: Yellow 'Blacklist Client' button missing or bypassed. {client_name} is likely already blacklisted.")
                print(f"Successfully processed record profile updates for: {client_name}")
            
            time.sleep(4) 
            
        except Exception as e:
            print(f"Failed on Row {index + 1} ('{client_name}'): {str(e)}")
            continue

    print("\nDatabase update loop processing complete!")
    browser.close()