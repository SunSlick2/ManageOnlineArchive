# -*- coding: utf-8 -*-
"""
Online_Archive_Sorter_v02_01.py
Forked from InboxSorter_v38.11 and OnlineArchiveSorter_v01.03.

This version supports all 11 rules. 
- Rules targeting "ToDelete" result in permanent deletion.
- Other rules move items to folders within the Online Archive.
- User selects specific folder (Root/Inbox/Sub-Inboxes) at runtime.
"""

import os
import win32com.client
import pandas as pd
import datetime
import openpyxl
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
import threading
import logging
import time
import json
import re
import pythoncom
import sys

class OnlineArchiveSorter:
    CONFIG_FILE_NAME = 'config_archive_v02.json'
    MAIL_ITEM_CLASS = 43 

    def __init__(self):
        self.config = self._load_config()
        self.setup_paths()
        self.setup_logging()
        
        # Load interval from config, default to 500
        self.cache_save_interval = self.config.get("cache_save_interval", 500)
        
        self.email_rules = {}
        self.keyword_rules = {}
        self.smtp_cache = {}
        self.processed_count = 0
        self.items_since_last_save = 0
        
        self.load_data()

    def _load_config(self):
        if not os.path.exists(self.CONFIG_FILE_NAME):
            print(f"Error: Config file {self.CONFIG_FILE_NAME} not found.")
            sys.exit(1)
        with open(self.CONFIG_FILE_NAME, 'r') as f:
            return json.load(f)

    def setup_paths(self):
        self.xls_path = self.config.get('xls_path')
        self.archive_name = self.config.get('archive_folder_name')

    def setup_logging(self):
        self.bulk_logger = self._create_logger('bulk_logger', self.config.get('log_bulk_path'))
        self.invalid_logger = self._create_logger('invalid_logger', self.config.get('log_invalid_path'))

    def _create_logger(self, name, log_file):
        logger = logging.getLogger(name)
        logger.setLevel(logging.INFO)
        if not logger.handlers:
            handler = logging.FileHandler(log_file)
            formatter = logging.Formatter('%(asctime)s|%(levelname)s|%(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
        return logger

    def load_data(self):
        """Loads all 11 rules from Excel as per v38.11 logic."""
        try:
            with pd.ExcelFile(self.xls_path) as xls:
                sheet_map = self.config.get('sheet_map', {})
                
                for rule_name, info in sheet_map.items():
                    df = pd.read_excel(xls, info['sheet'])
                    dest = info['destination_name']
                    
                    # Logic for Email rules
                    if 'Email' in rule_name and 'Keyword' not in rule_name:
                        col = info['column']
                        addresses = df[col].dropna().unique().tolist()
                        for addr in addresses:
                            # Rule 8 specific: ResearchEmail is sender only
                            is_sender_only = (rule_name == "ResearchEmail")
                            self.email_rules[addr.lower()] = {"dest": dest, "sender_only": is_sender_only}
                    
                    # Logic for Keyword rules
                    else:
                        cols = info.get('columns', [info.get('column')])
                        match_field = info.get('match_field', 'subject_only')
                        for col in cols:
                            keywords = df[col].dropna().unique().tolist()
                            for kw in keywords:
                                self.keyword_rules[str(kw).lower()] = {"dest": dest, "field": match_field}

                # Load SMTP Cache
                try:
                    cache_df = pd.read_excel(xls, 'SMTP_Cache')
                    self.smtp_cache = dict(zip(cache_df['ExchangeAddress'].str.lower(), cache_df['SMTPAddress']))
                except:
                    self.smtp_cache = {}
        except Exception as e:
            self.invalid_logger.critical(f"DataLoaderError||{e}")

    def save_smtp_cache(self):
        """Saves cache back to Excel. Critical for Archive due to volume."""
        try:
            df = pd.DataFrame(list(self.smtp_cache.items()), columns=['ExchangeAddress', 'SMTPAddress'])
            with pd.ExcelWriter(self.xls_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name='SMTP_Cache', index=False)
            self.items_since_last_save = 0
        except Exception as e:
            self.invalid_logger.error(f"CacheSaveError||{e}")

    def get_smtp_address(self, item):
        try:
            sender_obj = item.Sender
            if sender_obj.AddressEntryUserType == 0: # olExchangeUserAddressEntry
                ex_addr = sender_obj.Address.lower()
                if ex_addr in self.smtp_cache:
                    return self.smtp_cache[ex_addr]
                eu = sender_obj.GetExchangeUser()
                if eu:
                    smtp = eu.PrimarySmtpAddress
                    self.smtp_cache[ex_addr] = smtp
                    return smtp
            return item.SenderEmailAddress
        except:
            return None

    def get_folder_recursive(self, root_folder, folder_path):
        """Finds or creates folders within the Archive root."""
        current_node = root_folder
        parts = folder_path.split('\\')
        for part in parts:
            try:
                current_node = current_node.Folders.Item(part)
            except:
                current_node = current_node.Folders.Add(part)
        return current_node

    def process_email(self, item, archive_root):
        """Determines if email should be deleted or moved based on 11 rules."""
        try:
            subject = str(item.Subject).lower()
            body = str(item.Body).lower()
            sender_email = (self.get_smtp_address(item) or "").lower()
            
            # 1. Check Email Rules
            if sender_email in self.email_rules:
                rule_info = self.email_rules[sender_email]
                # Rule 8 Logic: Check if sender_only is required
                if not rule_info.get("sender_only") or sender_email: 
                    return self.execute_action(item, rule_info['dest'], archive_root, sender_email, "EmailMatch")

            # 2. Check Keyword Rules
            for kw, info in self.keyword_rules.items():
                match = False
                if info['field'] == 'subject_only' and kw in subject:
                    match = True
                elif info['field'] == 'subject_and_body' and (kw in subject or kw in body):
                    match = True
                
                if match:
                    return self.execute_action(item, info['dest'], archive_root, kw, "KeywordMatch")
            
            return False
        except Exception as e:
            self.invalid_logger.error(f"ItemProcessError||{e}")
            return False

    def execute_action(self, item, dest_name, archive_root, trigger, match_type):
        """Performs Delete or Move."""
        try:
            if dest_name == "ToDelete":
                item.Delete()
                self.bulk_logger.info(f"DELETED|{trigger}|{match_type}|{item.Subject}")
                return True
            else:
                dest_folder = self.get_folder_recursive(archive_root, dest_name)
                item.Move(dest_folder)
                self.bulk_logger.info(f"MOVED|{dest_name}|{trigger}|{match_type}|{item.Subject}")
                return True
        except Exception as e:
            self.invalid_logger.error(f"ActionError|{dest_name}|{e}")
            return False

    def run_archive_processing(self, target_folder_path):
        pythoncom.CoInitialize()
        try:
            print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Connecting to Outlook...")
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            archive_root = None
            
            print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Looking for archive: {self.archive_name}")
            for store in outlook.Stores:
                if store.DisplayName == self.archive_name:
                    archive_root = store.GetRootFolder()
                    break
            
            if not archive_root:
                messagebox.showerror("Error", f"Could not find archive: {self.archive_name}")
                print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] ERROR: Archive not found!")
                return

            print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Archive found. Resolving folder: {target_folder_path}")
            
            # Resolve the specific folder selected by user
            target_folder = archive_root
            if target_folder_path != "ROOT":
                parts = target_folder_path.split('\\')
                for i, part in enumerate(parts):
                    try:
                        print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Looking for folder: '{part}' under '{target_folder.Name}'")
                        
                        # List available folders for debugging
                        folder_names = []
                        for f in target_folder.Folders:
                            folder_names.append(f.Name)
                        print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Available folders: {folder_names}")
                        
                        target_folder = target_folder.Folders.Item(part)
                        print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Successfully found: '{part}'")
                        
                    except Exception as e:
                        error_msg = f"Cannot find folder '{part}' under '{target_folder.Name}'\n\n"
                        error_msg += f"Available folders under '{target_folder.Name}':\n"
                        for f in target_folder.Folders:
                            error_msg += f"  - {f.Name}\n"
                        error_msg += f"\nFull path attempted: {target_folder_path}\n"
                        error_msg += f"Error: {e}"
                        
                        messagebox.showerror("Folder Not Found", error_msg)
                        print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] ERROR: {error_msg}")
                        return
            
            print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Target folder: {target_folder.FolderPath}")
            print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Starting processing...")
            
            self.bulk_logger.info(f"STARTING|Folder: {target_folder.FolderPath}")
            
            self._process_folder_items(target_folder, archive_root)
            
            print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Saving SMTP cache...")
            self.save_smtp_cache() # Final save
            
            print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Processing complete!")
            print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Total items processed: {self.processed_count}")
            
            messagebox.showinfo("Done", f"Processing complete.\nProcessed: {self.processed_count} items.")
            
        except Exception as e:
            self.invalid_logger.critical(f"GlobalRunError||{e}")
            print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] ERROR: {e}")
        finally:
            pythoncom.CoUninitialize()

    def _process_folder_items(self, folder, archive_root):
        """Iterates backwards to ensure stability during moves/deletes."""
        try:
            items = folder.Items
            
            print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Processing items...")
            
            # Initialize progress tracking
            last_progress_time = time.time()
            items_processed_since_last_update = 0
            
            # We'll process without knowing the total count
            i = items.Count
            while i > 0:
                try:
                    item = items.Item(i)
                    if item.Class == self.MAIL_ITEM_CLASS:
                        if self.process_email(item, archive_root):
                            self.processed_count += 1
                            self.items_since_last_save += 1
                            items_processed_since_last_update += 1
                            
                            # Show progress every 100 items or every 30 seconds
                            current_time = time.time()
                            if items_processed_since_last_update >= 100 or (current_time - last_progress_time) >= 30:
                                print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Progress: {self.processed_count} items processed")
                                items_processed_since_last_update = 0
                                last_progress_time = current_time
                            
                            if self.items_since_last_save >= self.cache_save_interval:
                                print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Saving SMTP cache (interval reached)...")
                                self.save_smtp_cache()
                    i -= 1
                except Exception as e:
                    # Likely item moved/deleted or COM error - continue with next item
                    i -= 1
                    continue
                    
            print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] Folder processing complete")
        except Exception as e:
            self.invalid_logger.error(f"FolderProcessingError|{folder.Name}|{e}")
            print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] ERROR processing folder {folder.Name}: {e}")

    def start_gui(self):
        root = tk.Tk()
        root.title(f"Online Archive Sorter v02.01")
        root.geometry("450x400")

        tk.Label(root, text="Select Folder to Process:", font=("Arial", 12, "bold")).pack(pady=10)

        # Folder selection list - HARDCODED
        folders = [
            "ROOT",
            "Inbox",
            "Inbox\\Inbox1",
            "Inbox\\Inbox2",
            "Inbox\\Inbox3",
            "Inbox\\Inbox4"
        ]
        
        selected_folder = tk.StringVar(value="Inbox")
        for f in folders:
            tk.Radiobutton(root, text=f, variable=selected_folder, value=f).pack(anchor="w", padx=50)

        def start_task():
            folder_to_process = selected_folder.get()
            if messagebox.askyesno("Confirm", f"Process matching emails in {folder_to_process}?\n\n'ToDelete' items will be PERMANENTLY DELETED."):
                btn_start.config(state=tk.DISABLED)
                # Close the dialog box
                root.destroy()
                threading.Thread(target=lambda: self.run_archive_processing(folder_to_process), daemon=True).start()

        btn_start = tk.Button(root, text="Start Processing", command=start_task, 
                              bg="#28a745", fg="white", font=("Arial", 11, "bold"), height=2, width=20)
        btn_start.pack(pady=20)

        root.mainloop()

if __name__ == "__main__":
    sorter = OnlineArchiveSorter()
    sorter.start_gui()