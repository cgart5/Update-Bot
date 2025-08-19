import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sqlite3
import threading
import shutil
import tempfile
import time
import queue
from datetime import datetime
from A4GDB import A4GDB
import playwright
import aggressive
#import bot
from config import webpage
import pandas as pd

class ModernSyncGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("A4G Synchronization Manager")
        self.root.geometry("800x600")
        self.root.configure(bg='#1e1e1e')
        
        # Configure style for modern look
        self.setup_styles()
        
        # Initialize variables
        self.att_files = []
        self.kt_files = []
        self.att_folder = ""
        self.kt_folder = ""
        self.db_path = "A4G.db"  # Use A4GDB database name
        self.is_loading = False
        self.is_syncing = False
        self.a4g_db = None

        self.temp_att_dir = None
        self.temp_kt_dir = None

        # Thread communication
        self.message_queue = queue.Queue()
        self.progress_queue = queue.Queue()

        # Progress tracking (route-based)
        self.routes = []
        self.completed_routes = set()
        self.total_routes = 0

        # Create database
        self.init_database()
        
        # Create GUI
        self.create_widgets()

        # Initialize Bot
        self.filepath = ""
        self.url = webpage
        self.bot = None  # Will be initialized when needed

        # Start queue processing
        self.process_queues()

    def process_queues(self):
        """Process messages from worker threads safely in main thread"""
        try:
            while True:
                message = self.message_queue.get_nowait()
                self.log_message(message)
        except queue.Empty:
            pass

        try:
            while True:
                progress_data = self.progress_queue.get_nowait()
                self.update_progress_from_callback(progress_data)
        except queue.Empty:
            pass

        # Schedule next check
        self.root.after(100, self.process_queues)
    
    def gui_progress_callback(self, message, status):
        """Thread-safe callback for bot progress updates"""
        # Handle different types of messages
        if isinstance(message, dict) and status == "service_area_complete":
            # This is a service area completion callback
            self.progress_queue.put(message)
            formatted_message = f"[COMPLETE] Service area '{message.get('service_area', 'Unknown')}' completed - {message.get('facilities_processed', 0)} facilities, {message.get('routes_processed', 0)} routes"
        else:
            # Regular progress message
            formatted_message = f"[{status}] {message}"

        self.message_queue.put(formatted_message)


    def gui_summary_callback(self, summary_text):
        """Thread-safe callback for bot summary"""
        self.message_queue.put(summary_text)

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure modern dark theme
        style.configure('Modern.TFrame', 
                       background='#2d2d2d', 
                       relief='flat')
        style.configure('Header.TLabel', 
                       background='#2d2d2d', 
                       foreground='#ffffff', 
                       font=('Segoe UI', 16, 'bold'))
        style.configure('Modern.TLabel', 
                       background='#2d2d2d', 
                       foreground='#ffffff', 
                       font=('Segoe UI', 10))
        style.configure('Modern.TButton', 
                       background='#404040', 
                       foreground='#ffffff', 
                       font=('Segoe UI', 10),
                       borderwidth=0,
                       focuscolor='none')
        style.map('Modern.TButton',
                 background=[('active', '#505050'), ('pressed', '#606060')])
        style.configure('Accent.TButton', 
                       background='#0078d4', 
                       foreground='#ffffff', 
                       font=('Segoe UI', 10, 'bold'),
                       borderwidth=0,
                       focuscolor='none')
        style.map('Accent.TButton',
                 background=[('active', '#106ebe'), ('pressed', '#005a9e')])
        style.configure('Modern.Treeview',
                       background='#3d3d3d',
                       foreground='#ffffff',
                       fieldbackground='#3d3d3d',
                       borderwidth=0)
        style.configure('Modern.Treeview.Heading',
                       background='#404040',
                       foreground='#ffffff',
                       borderwidth=0)
                # Progress bar styles
        style.configure('Modern.Horizontal.TProgressbar',
                       background='#0078d4',
                       troughcolor='#404040',
                       borderwidth=0,
                       lightcolor='#0078d4',
                       darkcolor='#0078d4')
        
    def init_database(self):
        """Initialize connection to check A4G database"""
        try:
            # Just check if we can connect to the database
            # A4GDB will handle the actual table creation
            conn = sqlite3.connect(self.db_path)
            conn.close()
            print("Database connection verified")  # Use print instead of log_message
        except Exception as e:
            messagebox.showerror("Database Error", f"Failed to connect to database: {str(e)}")
    
    def create_widgets(self):
        # Main container
        main_frame = ttk.Frame(self.root, style='Modern.TFrame')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Header
        header_label = ttk.Label(main_frame, text="A4G Synchronization Manager", 
                                style='Header.TLabel')
        header_label.pack(pady=(0, 30))
        
        # File upload section
        self.create_upload_section(main_frame)
        
        # Database section
        self.create_database_section(main_frame)

        # Progress section
        self.create_progress_section(main_frame)

        # Synchronization section
        self.create_sync_section(main_frame)

        # Status section
        self.create_status_section(main_frame)
        
    def create_upload_section(self, parent):
        upload_frame = ttk.Frame(parent, style='Modern.TFrame')
        upload_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Section title
        title_label = ttk.Label(upload_frame, text="📁 File Upload", 
                               style='Modern.TLabel', font=('Segoe UI', 12, 'bold'))
        title_label.pack(anchor=tk.W, pady=(0, 10))
        
        # File upload options frame
        options_frame = ttk.Frame(upload_frame, style='Modern.TFrame')
        options_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Radio buttons for upload method
        self.upload_method = tk.StringVar(value="folder")
        folder_radio = ttk.Radiobutton(options_frame, text="Select Folders", 
                                      variable=self.upload_method, value="folder",
                                      command=self.toggle_upload_method)
        folder_radio.pack(side=tk.LEFT, padx=(0, 20))
        
        files_radio = ttk.Radiobutton(options_frame, text="Select Individual Files", 
                                     variable=self.upload_method, value="files",
                                     command=self.toggle_upload_method)
        files_radio.pack(side=tk.LEFT)
        
        # File upload buttons frame
        self.buttons_frame = ttk.Frame(upload_frame, style='Modern.TFrame')
        self.buttons_frame.pack(fill=tk.X)
        
        # ATT files section
        self.att_frame = ttk.Frame(self.buttons_frame, style='Modern.TFrame')
        self.att_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        ttk.Label(self.att_frame, text="ATT Files", style='Modern.TLabel').pack(anchor=tk.W)
        self.att_button = ttk.Button(self.att_frame, text="Select ATT Folder", 
                                    command=self.select_att_files, style='Modern.TButton')
        self.att_button.pack(fill=tk.X, pady=(5, 0))
        
        self.att_count_label = ttk.Label(self.att_frame, text="No files selected", 
                                        style='Modern.TLabel')
        self.att_count_label.pack(anchor=tk.W, pady=(5, 0))
        
        # KT files section
        self.kt_frame = ttk.Frame(self.buttons_frame, style='Modern.TFrame')
        self.kt_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))
        
        ttk.Label(self.kt_frame, text="KT Files", style='Modern.TLabel').pack(anchor=tk.W)
        self.kt_button = ttk.Button(self.kt_frame, text="Select KT Folder", 
                                   command=self.select_kt_files, style='Modern.TButton')
        self.kt_button.pack(fill=tk.X, pady=(5, 0))
        
        self.kt_count_label = ttk.Label(self.kt_frame, text="No files selected", 
                                       style='Modern.TLabel')
        self.kt_count_label.pack(anchor=tk.W, pady=(5, 0))
        
    def create_database_section(self, parent):
        db_frame = ttk.Frame(parent, style='Modern.TFrame')
        db_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Section title
        title_label = ttk.Label(db_frame, text="💾 Database Operations", 
                               style='Modern.TLabel', font=('Segoe UI', 12, 'bold'))
        title_label.pack(anchor=tk.W, pady=(0, 10))
        
        # Database buttons
        db_buttons_frame = ttk.Frame(db_frame, style='Modern.TFrame')
        db_buttons_frame.pack(fill=tk.X)
        
        self.load_button = ttk.Button(db_buttons_frame, text="Load Files to Database", 
                                     command=self.load_files_to_db, 
                                     style='Accent.TButton', state=tk.DISABLED)
        self.load_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.view_button = ttk.Button(db_buttons_frame, text="View Database", 
                                     command=self.view_database, style='Modern.TButton')
        self.view_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.clear_button = ttk.Button(db_buttons_frame, text="Clear Database",
                                      command=self.clear_database, style='Modern.TButton')
        self.clear_button.pack(side=tk.LEFT)

        self.export_button = ttk.Button(db_buttons_frame, text="Export DB to Excel",
                                       command=self.export_db_to_excel, style='Modern.TButton')
        self.export_button.pack(side=tk.LEFT, padx=(10, 10))


    def create_progress_section(self, parent):
        """Create service area progress tracking section"""
        progress_frame = ttk.Frame(parent, style='Modern.TFrame')
        progress_frame.pack(fill=tk.X, pady=(0, 15))

        # Section title
        title_label = ttk.Label(progress_frame, text="📊Progress Bar",
                            style='Modern.TLabel', font=('Segoe UI', 12, 'bold'))
        title_label.pack(anchor=tk.W, pady=(0, 10))

        # Progress info frame
        progress_info_frame = ttk.Frame(progress_frame, style='Modern.TFrame')
        progress_info_frame.pack(fill=tk.X, pady=(0, 10))

        # Progress labels
        self.progress_label = ttk.Label(progress_info_frame, text="Ready to start synchronization",
                                    style='Modern.TLabel')
        self.progress_label.pack(side=tk.LEFT)

        # Progress bar
        self.progress_bar = ttk.Progressbar(progress_frame,
                                        style='Modern.Horizontal.TProgressbar',
                                        mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))

    def create_sync_section(self, parent):
        sync_frame = ttk.Frame(parent, style='Modern.TFrame')
        sync_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Section title
        title_label = ttk.Label(sync_frame, text="🔄 Synchronization", 
                               style='Modern.TLabel', font=('Segoe UI', 12, 'bold'))
        title_label.pack(anchor=tk.W, pady=(0, 10))
        
        # Sync button
        self.sync_button = ttk.Button(sync_frame, text="Run Synchronization", 
                                     command=self.run_synchronization, 
                                     style='Accent.TButton', state=tk.DISABLED)
        self.sync_button.pack(anchor=tk.W)
        
    def create_status_section(self, parent):
        status_frame = ttk.Frame(parent, style='Modern.TFrame')
        status_frame.pack(fill=tk.BOTH, expand=True)
        
        # Section title
        title_label = ttk.Label(status_frame, text="📊 Status & Logs", 
                               style='Modern.TLabel', font=('Segoe UI', 12, 'bold'))
        title_label.pack(anchor=tk.W, pady=(0, 10))
        
        # Create frame for text and scrollbar
        text_frame = ttk.Frame(status_frame, style='Modern.TFrame')
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        # Status text area
        self.status_text = tk.Text(text_frame, height=8, bg='#3d3d3d', fg='#ffffff',
                                  font=('Consolas', 9), insertbackground='#ffffff',
                                  selectbackground='#0078d4', relief='flat', bd=0)
        self.status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Scrollbar for status text
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.status_text.yview)
        
        # Initial status message
        self.log_message("Welcome to A4G Synchronization Manager!")
        self.log_message("Please select ATT and KT files to begin...")
    def initialize_route_progress(self):
        """Initialize route progress from database"""
        try:
            if not os.path.exists(self.db_path):
                print(f"🚨 Database file not found: {self.db_path}")
                self.log_message("⚠️ Database file not found - progress tracking unavailable")
                return

            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()

            # Check if Route table exists
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='Route'")
            table_exists = cursor.fetchone()

            if not table_exists:
                print("🚨 Route table not found in database")
                self.log_message("⚠️ Route table not found - progress tracking unavailable")
                conn.close()
                return

            # Get all routes that have zip codes (only routes with zip codes get processed)
            cursor.execute("""
                SELECT DISTINCT r.Rt
                FROM Route r
                WHERE EXISTS (
                    SELECT 1
                    FROM ZipCode z
                    WHERE z.Rt = r.Rt
                )
                ORDER BY r.Rt
            """)
            routes = cursor.fetchall()

            print(f"🔍 Found {len(routes)} routes with zip codes in database")

            if routes:
                self.routes = [route[0] for route in routes]
                self.total_routes = len(self.routes)
                self.completed_routes = set()  # Reset completed routes

                print(f"📊 Routes loaded: {len(self.routes)} total routes")

                self.update_progress_display()
                self.log_message(f"📊 Initialized progress tracking for {len(self.routes)} routes")
            else:
                # If no routes in database, we'll track them as they're discovered
                self.routes = []
                self.total_routes = 0
                self.completed_routes = set()
                print("⚠️ No routes found in database")
                self.log_message("⚠️ No routes found in database - will track as discovered")

            conn.close()

        except Exception as e:
            error_msg = f"Error initializing route progress: {str(e)}"
            print(f"🚨 {error_msg}")
            self.log_message(f"❌ {error_msg}")

    def update_progress_display(self):
        """Update the progress bar and labels"""
        total = len(self.routes) if hasattr(self, 'routes') else 0
        completed = len(self.completed_routes) if hasattr(self, 'completed_routes') else 0

        if hasattr(self, 'progress_bar') and hasattr(self, 'progress_label'):
            if total > 0:
                percentage = (completed / total) * 100
                self.progress_bar['value'] = percentage
                self.progress_label.config(text=f"{completed} / {total} routes completed ({percentage:.1f}%)")
            else:
                self.progress_bar['value'] = 0
                self.progress_label.config(text="Ready to start synchronization")



    def update_progress_from_callback(self, progress_data):
        """Update progress from sync process callbacks"""
        if progress_data.get("type") == "service_area_complete":
            # Get the service area name from the bot's callback data
            service_area = progress_data.get("service_area", "")
            facilities_processed = progress_data.get('facilities_processed', 0)
            routes_processed = progress_data.get('routes_processed', 0)

            if service_area:
                # Log the service area completion message that will be picked up by route monitoring
                completion_message = f"✅ Service area '{service_area}' completed - {facilities_processed} facilities, {routes_processed} routes processed"
                self.log_message(completion_message)

                # Note: The actual progress tracking is now handled by the log message monitoring system
                # which will detect route completions from the detailed log messages

    
    def toggle_upload_method(self):
        """Toggle between folder and file selection methods"""
        method = self.upload_method.get()
        if method == "folder":
            self.att_button.config(text="Select ATT Folder")
            self.kt_button.config(text="Select KT Folder")
        else:
            self.att_button.config(text="Select ATT Files (.xlsx)")
            self.kt_button.config(text="Select KT Files (.xlsx)")
            
        # Reset selections
        self.att_files = []
        self.kt_files = []
        self.att_folder = ""
        self.kt_folder = ""
        self.att_count_label.config(text="No files selected")
        self.kt_count_label.config(text="No files selected")
        self.update_buttons()
        
    def select_att_files(self):
        method = self.upload_method.get()
        
        if method == "folder":
            folder = filedialog.askdirectory(title="Select ATT Folder")
            if folder:
                self.att_folder = folder
                # Count Excel files in the folder
                excel_files = [f for f in os.listdir(folder) 
                              if f.lower().endswith(('.xlsx', '.xls'))]
                self.att_files = [os.path.join(folder, f) for f in excel_files]
                self.att_count_label.config(text=f"Folder selected ({len(excel_files)} Excel files)")
                self.log_message(f"Selected ATT folder: {folder}")
                self.log_message(f"Found {len(excel_files)} Excel files in ATT folder")
        else:
            files = filedialog.askopenfilenames(
                title="Select ATT Excel Files",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )
            if files:
                self.att_files = list(files)
                self.att_count_label.config(text=f"{len(files)} files selected")
                self.log_message(f"Selected {len(files)} ATT Excel files")
                
                # Create temporary ATT folder
                if self.temp_att_dir:
                    shutil.rmtree(self.temp_att_dir)
                self.temp_att_dir = tempfile.mkdtemp(prefix="att_files_")
                
                # Copy files into temp folder
                for f in self.att_files:
                    shutil.copy(f, self.temp_att_dir)
                
                self.att_folder = self.temp_att_dir
        self.update_buttons()
            
    def select_kt_files(self):
        method = self.upload_method.get()
        
        if method == "folder":
            folder = filedialog.askdirectory(title="Select KT Folder")
            if folder:
                self.kt_folder = folder
                # Count Excel files in the folder
                excel_files = [f for f in os.listdir(folder) 
                              if f.lower().endswith(('.xlsx', '.xls'))]
                self.kt_files = [os.path.join(folder, f) for f in excel_files]
                self.kt_count_label.config(text=f"Folder selected ({len(excel_files)} Excel files)")
                self.log_message(f"Selected KT folder: {folder}")
                self.log_message(f"Found {len(excel_files)} Excel files in KT folder")
        else:
            files = filedialog.askopenfilenames(
                title="Select KT Excel Files",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
            )
            if files:
                self.kt_files = list(files)
                self.kt_count_label.config(text=f"{len(files)} files selected")
                self.log_message(f"Selected {len(files)} KT Excel files")

                # Create temporary KT folder
                if self.temp_kt_dir:
                    shutil.rmtree(self.temp_kt_dir)
                self.temp_kt_dir = tempfile.mkdtemp(prefix="kt_files_")
                
                # Copy files into temp folder
                for f in self.kt_files:
                    shutil.copy(f, self.temp_kt_dir)
                
                self.kt_folder = self.temp_kt_dir
                
        self.update_buttons()
            
    def update_buttons(self):
        if self.att_files and self.kt_files and not self.is_loading and not self.is_syncing:
            self.load_button.config(state=tk.NORMAL)
        else:
            self.load_button.config(state=tk.DISABLED)
            
        if self.a4g_db is not None and not self.is_loading and not self.is_syncing:
            self.sync_button.config(state=tk.NORMAL)
        else:
            self.sync_button.config(state=tk.DISABLED)
            
    def load_files_to_db(self):
        if self.is_loading:
            return
            
        self.is_loading = True
        self.load_button.config(text="Loading...", state=tk.DISABLED)
        self.sync_button.config(state=tk.DISABLED)
        
        # Run loading in separate thread to prevent GUI freezing
        thread = threading.Thread(target=self._load_files_thread)
        thread.daemon = True
        thread.start()
        
    def _load_files_thread(self):
        try:
            self.message_queue.put("Initializing A4G Database...")
            
            # Initialize A4GDB with the selected files/folders
            if self.upload_method.get() == "folder":
                self.a4g_db = A4GDB(self.att_folder, self.kt_folder)
            else:
                # For individual files, we'll need to modify A4GDB to accept file lists
                # For now, create temporary folders or modify the A4GDB constructor
                self.a4g_db = A4GDB(self.att_folder, self.kt_folder)
                
            self.message_queue.put("A4G Database initialized successfully")
            
            # Process KT files first
            self.message_queue.put("Processing KT files...")
            self.a4g_db.kt_files()
            self.message_queue.put("KT files processed successfully")
            
            # Process ATT files
            self.message_queue.put("Processing ATT files...")
            self.a4g_db.att_files()
            self.message_queue.put("ATT files processed successfully")
            
            # Update GUI in main thread
            self.root.after(0, self._load_complete)
            
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda: self._load_error(error_msg))
            
    def _load_complete(self):
        self.is_loading = False
        self.load_button.config(text="Load Files to Database", state=tk.NORMAL)
        self.update_buttons()
        self.log_message(f"✅ Successfully processed {len(self.att_files)} ATT files and {len(self.kt_files)} KT files!")
        self.log_message("Database is ready for synchronization")

        # Initialize progress tracking
        self.initialize_route_progress()
        
    def _load_error(self, error_msg):
        self.is_loading = False
        self.load_button.config(text="Load Files to Database", state=tk.NORMAL)
        self.update_buttons()
        self.log_message(f"❌ Error processing files: {error_msg}")
        messagebox.showerror("Processing Error", f"Failed to process files:\n{error_msg}")
        
    def view_database(self):
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Check if tables exist and count records
            tables = ['Service_Area', 'Facility', 'Route', 'ZipCode']
            counts = {}
            
            for table in tables:
                try:
                    cursor.execute(f"SELECT COUNT(*) FROM {table}")
                    counts[table] = cursor.fetchone()[0]
                except sqlite3.OperationalError:
                    counts[table] = 0
            
            conn.close()
            
            info_msg = f"A4G Database Contents:\n\n"
            info_msg += f"Service Areas: {counts.get('Service_Area', 0)}\n"
            info_msg += f"Facilities: {counts.get('Facility', 0)}\n"
            info_msg += f"Routes: {counts.get('Route', 0)}\n"
            info_msg += f"Zip Codes: {counts.get('ZipCode', 0)}"
            
            messagebox.showinfo("Database Info", info_msg)
            self.log_message(f"Database contains {counts.get('Service_Area', 0)} Service Areas, "
                           f"{counts.get('Facility', 0)} Facilities, "
                           f"{counts.get('Route', 0)} Routes, "
                           f"{counts.get('ZipCode', 0)} Zip Codes")
            
        except Exception as e:
            messagebox.showerror("Database Error", f"Failed to read database: {str(e)}")
            
    def clear_database(self):
        if messagebox.askyesno("Clear Database", "Are you sure you want to clear all data from the A4G database?"):
            try:
                conn = sqlite3.connect(self.db_path)
                cursor = conn.cursor()
                
                # Clear A4G database tables in correct order (due to foreign key constraints)
                tables = ['ZipCode', 'Route', 'Facility', 'Service_Area']
                for table in tables:
                    try:
                        cursor.execute(f"DELETE FROM {table}")
                    except sqlite3.OperationalError:
                        pass  # Table might not exist yet
                
                conn.commit()
                conn.close()
                
                self.a4g_db = None
                self.update_buttons()
                self.log_message("🗑️ A4G Database cleared successfully")
                
            except Exception as e:
                messagebox.showerror("Database Error", f"Failed to clear database: {str(e)}")
                
    def run_synchronization(self):
        if messagebox.askyesno("Run Synchronization", "Are you ready to run the synchronization process?"):
            self.is_syncing = True
            self.sync_button.config(text="Synchronizing...", state=tk.DISABLED)
            self.load_button.config(state=tk.DISABLED)

            # Initialize/refresh route progress from database
            self.initialize_route_progress()

            # Reset completed routes for new sync
            if hasattr(self, 'completed_routes'):
                self.completed_routes.clear()
            self.update_progress_display()

            # Log the initialization
            total_routes = len(self.routes) if hasattr(self, 'routes') else 0
            self.log_message(f"🔄 Starting synchronization process...")
            self.log_message(f"📊 Initialized progress tracking for {total_routes} routes")
            print(f"🎯 SYNC STARTED: Tracking {total_routes} routes")

            # Run sync in separate thread
            thread = threading.Thread(target=self._sync_thread)
            thread.daemon = True
            thread.start()
            
    def _sync_thread(self):
        """Run synchronization in separate thread"""
        try:
            # Initialize bot with thread-safe callbacks
            bot = aggressive.Bot(self.url, self.filepath, self.gui_progress_callback, self.gui_summary_callback)
            
            # Run the bot main function
            bot.bot_main()
            
            # Signal completion
            self.root.after(0, self._sync_complete)
            
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda: self._sync_error(error_msg))
            
    def _sync_complete(self):
        self.is_syncing = False
        self.sync_button.config(text="Run Synchronization", state=tk.NORMAL)
        self.update_buttons()
        self.log_message("✅ A4G Synchronization completed successfully!")
        messagebox.showinfo("Sync Complete", "A4G database synchronization completed successfully!")
        
    def _sync_error(self, error_msg):
        self.is_syncing = False
        self.sync_button.config(text="Run Synchronization", state=tk.NORMAL)
        self.update_buttons()
        self.log_message(f"❌ Synchronization failed: {error_msg}")
        messagebox.showerror("Sync Error", f"Synchronization failed:\n{error_msg}")
        
    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"

        # Debug mode: If bot is running, log all messages to console
        if hasattr(self, 'is_syncing') and self.is_syncing:
            print(f"BOT_RUNNING - LOG: {message}")

        # Check for route completion in log messages
        self.check_for_route_completion(message)

        self.status_text.insert(tk.END, log_entry)
        self.status_text.see(tk.END)
        self.root.update_idletasks()

    def check_for_route_completion(self, message):
        """Check log message for route completion and update progress"""
        # Look for route completion patterns in the log message
        import re

        # Pattern 1: "✅ Successfully added postal codes for route: RouteID"
        pattern1 = r"✅ Successfully added postal codes for route:\s*([^\s]+)"
        match1 = re.search(pattern1, message)

        if match1:
            route = match1.group(1).strip()
            print(f"✅ FOUND ROUTE COMPLETION: {route}")
            self.handle_route_completion(route)
            return

        # Pattern 2: Alternative success message format
        pattern2 = r"Successfully added postal codes for route:\s*([^\s]+)"
        match2 = re.search(pattern2, message)

        if match2:
            route = match2.group(1).strip()
            print(f"✅ FOUND ROUTE COMPLETION (ALT): {route}")
            self.handle_route_completion(route)
            return

        # Pattern 3: Any message containing "route" and "successfully" or "completed" (fallback)
        if "route" in message.lower() and ("successfully" in message.lower() or "completed" in message.lower()):
            # Try to extract route name from various formats
            patterns = [
                r"route[:\s]+([A-Za-z0-9_-]+)[^a-zA-Z]*(?:successfully|completed)",
                r"([A-Za-z0-9_-]+)[^a-zA-Z]*route[^a-zA-Z]*(?:successfully|completed)",
                r"(?:successfully|completed)[^a-zA-Z]*route[:\s]+([A-Za-z0-9_-]+)"
            ]

            for pattern in patterns:
                match = re.search(pattern, message, re.IGNORECASE)
                if match:
                    route = match.group(1).strip()
                    if route and len(route) > 1:  # Make sure we got a meaningful route name
                        print(f"✅ FOUND ROUTE COMPLETION (FALLBACK): {route}")
                        self.handle_route_completion(route)
                        return

    def handle_route_completion(self, route):
        """Handle a route completion detected from log messages"""
        # Initialize routes from database if not already done
        if not hasattr(self, 'routes') or not self.routes:
            self.initialize_route_progress()

        # Add the completed route
        if route and hasattr(self, 'completed_routes'):
            self.completed_routes.add(route)
            self.update_progress_display()

            # Always print progress updates so user can see them
            total_routes = len(self.routes) if hasattr(self, 'routes') else 0
            completed_count = len(self.completed_routes)
            progress_msg = f"🎯 ROUTE PROGRESS: {route} completed ({completed_count}/{total_routes})"
            print(progress_msg)

            # Also log it to the GUI
            self.status_text.insert(tk.END, f"[PROGRESS] {progress_msg}\n")
            self.status_text.see(tk.END)

    def __del__(self):
        """Cleanup temporary directories"""
        if self.temp_att_dir and os.path.exists(self.temp_att_dir):
            shutil.rmtree(self.temp_att_dir)
        if self.temp_kt_dir and os.path.exists(self.temp_kt_dir):
            shutil.rmtree(self.temp_kt_dir)
    def export_db_to_excel(self):
        try:
            export_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Export Database to Excel"
            )
            if not export_path:
                return

            conn = sqlite3.connect(self.db_path)
            tables = ['Service_Area', 'Facility', 'Route', 'ZipCode']
            with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                for table in tables:
                    try:
                        df = pd.read_sql_query(f"SELECT * FROM {table}", conn)
                        df.to_excel(writer, sheet_name=table, index=False)
                    except Exception:
                        pass  # Table might not exist

             # Add joined sheet
                try:
                    join_query = """
                    SELECT
                        sa.CTRY AS Country,
                        sa.SA AS ServiceArea,
                        f.FAC AS Facility,
                        r.Rt AS Route,
                        z.Zip AS ZipCode
                    FROM Service_Area sa
                    INNER JOIN Facility f ON sa.SA = f.SA
                    INNER JOIN Route r ON f.FAC = r.FAC
                    INNER JOIN ZipCode z ON r.Rt = z.Rt
                    """
                    joined_df = pd.read_sql_query(join_query, conn)
                    joined_df.to_excel(writer, sheet_name="JoinedData", index=False)
                except Exception as e:
                    self.log_message(f"❌ Failed to export joined data: {str(e)}")
            conn.close()

            self.log_message(f"✅ Database exported to Excel: {export_path}")
            messagebox.showinfo("Export Complete", f"Database exported to:\n{export_path}")
        except Exception as e:
            self.log_message(f"❌ Export failed: {str(e)}")
            messagebox.showerror("Export Error", f"Failed to export database:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ModernSyncGUI(root)
    root.mainloop()
