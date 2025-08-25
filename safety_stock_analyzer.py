#!/usr/bin/env python3
"""
Safety Stock Analyzer - Professional Desktop Application
A powerful tool for analyzing spare parts usage and calculating safety stock levels.
Built with PyQt6 for a modern, professional desktop experience.
"""

import sys
import os
import pandas as pd
import numpy as np
from datetime import datetime
from pathlib import Path

# PyQt6 imports
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QPushButton, QTableWidget, 
                             QTableWidgetItem, QTabWidget, QTextEdit, QFileDialog,
                             QMessageBox, QProgressBar, QStatusBar, QMenuBar,
                             QMenu, QSplitter, QFrame, QGroupBox,
                             QGridLayout, QHeaderView, QAbstractItemView, QComboBox)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QMimeData, QSize
from PyQt6.QtGui import (QFont, QPixmap, QIcon, QDragEnterEvent, QDropEvent, 
                         QPalette, QColor, QAction)

# Data analysis imports
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import seaborn as sns

class SafetyStockAnalyzer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Safety Stock Analyzer - Professional Edition")
        self.setGeometry(100, 100, 1400, 900)
        
        # Initialize data
        self.data = None
        self.analysis_results = None
        self.process_parts = None  # New: Process parts data
        self.process_analysis = None  # New: Process analysis results
        
        # Setup UI
        self.setup_ui()
        self.setup_menu()
        self.setup_status_bar()
        
        # Apply modern styling
        self.apply_modern_style()
        
    def setup_ui(self):
        """Setup the main user interface"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout(central_widget)
        
        # Header
        header = self.create_header()
        main_layout.addWidget(header)
        
        # File upload area
        upload_area = self.create_upload_area()
        main_layout.addWidget(upload_area)
        
        # Main content area
        content_splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Data display area
        data_area = self.create_data_area()
        content_splitter.addWidget(data_area)
        
        # Analysis area
        analysis_area = self.create_analysis_area()
        content_splitter.addWidget(analysis_area)
        
        content_splitter.setSizes([800, 600])
        main_layout.addWidget(content_splitter)
        
        # Control buttons
        control_buttons = self.create_control_buttons()
        main_layout.addWidget(control_buttons)
        
    def create_header(self):
        """Create the application header"""
        header_frame = QFrame()
        header_frame.setFrameStyle(QFrame.Shape.StyledPanel)
        header_frame.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #2E86AB, stop:1 #A23B72);
                border-radius: 10px;
                margin: 5px;
            }
        """)
        
        header_layout = QHBoxLayout(header_frame)
        
        # Title
        title_label = QLabel("🔬 Safety Stock Analyzer")
        title_label.setFont(QFont("Segoe UI", 24, QFont.Weight.Bold))
        title_label.setStyleSheet("color: white; padding: 10px;")
        
        # Subtitle
        subtitle_label = QLabel("Professional Inventory Management & Analysis")
        subtitle_label.setFont(QFont("Segoe UI", 12))
        subtitle_label.setStyleSheet("color: #E0E0E0; padding: 10px;")
        
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        header_layout.addWidget(subtitle_label)
        
        return header_frame
        
    def create_upload_area(self):
        """Create the compact file upload area with drag & drop"""
        upload_frame = QFrame()
        upload_frame.setFrameStyle(QFrame.Shape.StyledPanel)
        upload_frame.setStyleSheet("""
            QFrame {
                background-color: #F8F9FA;
                border: 2px dashed #6C757D;
                border-radius: 8px;
                margin: 3px;
                max-height: 60px;
            }
            QFrame:hover {
                border-color: #007BFF;
                background-color: #E3F2FD;
            }
        """)
        
        upload_layout = QHBoxLayout(upload_frame)
        upload_layout.setContentsMargins(15, 8, 15, 8)
        upload_layout.setSpacing(15)
        
        # Compact upload text
        upload_label = QLabel("📁 Drag & Drop Files Here or Click to Browse")
        upload_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        upload_label.setFont(QFont("Segoe UI", 10))
        upload_label.setStyleSheet("color: #495057;")
        
        # Compact upload button
        upload_btn = QPushButton("Browse Files")
        upload_btn.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        upload_btn.setStyleSheet("""
            QPushButton {
                background-color: #007BFF;
                color: white;
                border: none;
                padding: 6px 16px;
                border-radius: 4px;
                font-weight: bold;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #0056B3;
            }
            QPushButton:pressed {
                background-color: #004085;
            }
        """)
        upload_btn.clicked.connect(self.browse_files)
        
        upload_layout.addWidget(upload_label)
        upload_layout.addWidget(upload_btn)
        
        # Enable drag & drop
        upload_frame.setAcceptDrops(True)
        upload_frame.dragEnterEvent = self.drag_enter_event
        upload_frame.dropEvent = self.drop_event
        
        return upload_frame
        
    def create_data_area(self):
        """Create the data display area"""
        data_group = QGroupBox("📊 Data Overview")
        data_group.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        
        data_layout = QVBoxLayout(data_group)
        
        # Data table
        self.data_table = QTableWidget()
        self.data_table.setAlternatingRowColors(True)
        self.data_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.data_table.horizontalHeader().setStretchLastSection(True)
        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        
        data_layout.addWidget(self.data_table)
        
        return data_group
        
    def create_analysis_area(self):
        """Create the analysis and results area"""
        analysis_group = QGroupBox("🔍 Analysis & Results")
        analysis_group.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        
        analysis_layout = QVBoxLayout(analysis_group)
        
        # Create tabs for different views
        self.analysis_tabs = QTabWidget()
        
        # Summary tab
        summary_tab = QWidget()
        summary_layout = QVBoxLayout(summary_tab)
        self.summary_text = QTextEdit()
        self.summary_text.setReadOnly(True)
        summary_layout.addWidget(self.summary_text)
        self.analysis_tabs.addTab(summary_tab, "📈 Summary")
        
        # Safety Stock tab
        safety_tab = QWidget()
        safety_layout = QVBoxLayout(safety_tab)
        self.safety_table = QTableWidget()
        safety_layout.addWidget(self.safety_table)
        self.analysis_tabs.addTab(safety_tab, "🛡️ Safety Stock")
        
        # Charts tab
        charts_tab = QWidget()
        charts_layout = QVBoxLayout(charts_tab)
        self.charts_canvas = self.create_charts_canvas()
        charts_layout.addWidget(self.charts_canvas)
        self.analysis_tabs.addTab(charts_tab, "📊 Charts")
        
        # Process Analysis tab
        process_tab = QWidget()
        process_layout = QVBoxLayout(process_tab)
        
        # Compact process parts upload area
        process_upload_frame = QFrame()
        process_upload_frame.setFrameStyle(QFrame.Shape.StyledPanel)
        process_upload_frame.setStyleSheet("""
            QFrame {
                background-color: #FFF3CD;
                border: 1px solid #FFC107;
                border-radius: 6px;
                margin: 2px;
                padding: 5px;
            }
        """)
        
        process_upload_layout = QHBoxLayout(process_upload_frame)
        process_upload_layout.setContentsMargins(10, 5, 10, 5)
        
        # Compact label
        process_upload_label = QLabel("🏭 Process Parts:")
        process_upload_label.setFont(QFont("Segoe UI", 9))
        process_upload_layout.addWidget(process_upload_label)
        
        # Compact button
        self.process_upload_btn = QPushButton("📁 Upload")
        self.process_upload_btn.setFont(QFont("Segoe UI", 9))
        self.process_upload_btn.setStyleSheet("""
            QPushButton {
                background-color: #FFC107;
                color: #212529;
                border: none;
                padding: 4px 12px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #E0A800;
            }
        """)
        self.process_upload_btn.clicked.connect(self.upload_process_parts)
        process_upload_layout.addWidget(self.process_upload_btn)
        
        process_upload_layout.addStretch()
        process_layout.addWidget(process_upload_frame)
        
        # Process analysis results table
        self.process_table = QTableWidget()
        process_layout.addWidget(self.process_table)
        
        # Compact process selector for safety stock analysis
        process_selector_frame = QFrame()
        process_selector_frame.setFrameStyle(QFrame.Shape.StyledPanel)
        process_selector_frame.setStyleSheet("""
            QFrame {
                background-color: #E3F2FD;
                border: 1px solid #2196F3;
                border-radius: 6px;
                margin: 2px;
                padding: 5px;
            }
        """)
        
        process_selector_layout = QHBoxLayout(process_selector_frame)
        process_selector_layout.setContentsMargins(10, 5, 10, 5)
        
        # Compact label
        selector_label = QLabel("🔍 Filter by Process:")
        selector_label.setFont(QFont("Segoe UI", 9))
        process_selector_layout.addWidget(selector_label)
        
        # Process dropdown
        self.process_selector = QComboBox()
        self.process_selector.setFont(QFont("Segoe UI", 9))
        self.process_selector.setStyleSheet("""
            QComboBox {
                padding: 4px 8px;
                border: 1px solid #2196F3;
                border-radius: 4px;
                background-color: white;
                min-width: 120px;
            }
        """)
        self.process_selector.addItem("All Processes")
        self.process_selector.currentTextChanged.connect(self.on_process_selection_changed)
        process_selector_layout.addWidget(self.process_selector)
        
        # Compact apply filter button
        self.apply_filter_btn = QPushButton("🔍 Apply")
        self.apply_filter_btn.setFont(QFont("Segoe UI", 9))
        self.apply_filter_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 4px 12px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
        """)
        self.apply_filter_btn.clicked.connect(self.apply_process_filter)
        self.apply_filter_btn.setEnabled(False)
        process_selector_layout.addWidget(self.apply_filter_btn)
        
        process_selector_layout.addStretch()
        process_layout.addWidget(process_selector_frame)
        
        self.analysis_tabs.addTab(process_tab, "🏭 Process Analysis")
        
        analysis_layout.addWidget(self.analysis_tabs)
        
        return analysis_group
        
    def create_charts_canvas(self):
        """Create matplotlib canvas for charts"""
        figure = Figure(figsize=(8, 6))
        canvas = FigureCanvas(figure)
        return canvas
        
    def create_control_buttons(self):
        """Create control buttons"""
        button_frame = QFrame()
        button_layout = QHBoxLayout(button_frame)
        
        # Analyze button
        self.analyze_btn = QPushButton("🔍 Run Analysis")
        self.analyze_btn.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        self.analyze_btn.setStyleSheet("""
            QPushButton {
                background-color: #28A745;
                color: white;
                border: none;
                padding: 12px 24px;
                border-radius: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #218838;
            }
            QPushButton:disabled {
                background-color: #6C757D;
            }
        """)
        self.analyze_btn.clicked.connect(self.run_analysis)
        self.analyze_btn.setEnabled(False)
        
        # Export button
        self.export_btn = QPushButton("📤 Export Results")
        self.export_btn.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        self.export_btn.setStyleSheet("""
            QPushButton {
                background-color: #17A2B8;
                color: white;
                border: none;
                padding: 12px 24px;
                border-radius: 8px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #138496;
            }
            QPushButton:disabled {
                background-color: #6C757D;
            }
        """)
        self.export_btn.clicked.connect(self.export_results)
        self.export_btn.setEnabled(False)
        
        # Clear button
        clear_btn = QPushButton("🗑️ Clear Data")
        clear_btn.setFont(QFont("Segoe UI", 12))
        clear_btn.setStyleSheet("""
            QPushButton {
                background-color: #DC3545;
                color: white;
                border: none;
                padding: 12px 24px;
                border-radius: 8px;
            }
            QPushButton:hover {
                background-color: #C82333;
            }
        """)
        clear_btn.clicked.connect(self.clear_data)
        
        button_layout.addWidget(self.analyze_btn)
        button_layout.addWidget(self.export_btn)
        button_layout.addStretch()
        button_layout.addWidget(clear_btn)
        
        return button_frame
        
    def setup_menu(self):
        """Setup the application menu"""
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu("File")
        
        open_action = QAction("Open File", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.browse_files)
        file_menu.addAction(open_action)
        
        export_action = QAction("Export Results", self)
        export_action.setShortcut("Ctrl+E")
        export_action.triggered.connect(self.export_results)
        file_menu.addAction(export_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("Exit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Analysis menu
        analysis_menu = menubar.addMenu("Analysis")
        
        run_analysis_action = QAction("Run Analysis", self)
        run_analysis_action.setShortcut("F5")
        run_analysis_action.triggered.connect(self.run_analysis)
        analysis_menu.addAction(run_analysis_action)
        
        # Help menu
        help_menu = menubar.addMenu("Help")
        
        about_action = QAction("About", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
        
    def setup_status_bar(self):
        """Setup the status bar"""
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready to load files")
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.status_bar.addPermanentWidget(self.progress_bar)
        
    def apply_modern_style(self):
        """Apply modern styling to the application"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #F8F9FA;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #DEE2E6;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
            QTableWidget {
                gridline-color: #DEE2E6;
                background-color: white;
                alternate-background-color: #F8F9FA;
                selection-background-color: #007BFF;
                selection-color: white;
            }
            QHeaderView::section {
                background-color: #E9ECEF;
                padding: 8px;
                border: 1px solid #DEE2E6;
                font-weight: bold;
            }
            QTabWidget::pane {
                border: 1px solid #DEE2E6;
                background-color: white;
            }
            QTabBar::tab {
                background-color: #E9ECEF;
                padding: 8px 16px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background-color: white;
                border-bottom: 2px solid #007BFF;
            }
        """)
        
    def drag_enter_event(self, event: QDragEnterEvent):
        """Handle drag enter event"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            
    def drop_event(self, event: QDropEvent):
        """Handle drop event"""
        files = []
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.lower().endswith(('.xlsx', '.xls', '.csv', '.txt')):
                files.append(file_path)
                
        if files:
            self.load_files(files)
            
    def browse_files(self):
        """Browse and select files"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Files to Analyze",
            "",
            "All Supported Files (*.xlsx *.xls *.csv *.txt);;Excel Files (*.xlsx *.xls);;CSV Files (*.csv);;Text Files (*.txt)"
        )
        
        if files:
            self.load_files(files)
            
    def upload_process_parts(self):
        """Upload process parts file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Process Parts File",
            "",
            "Excel Files (*.xlsx *.xls);;CSV Files (*.csv)"
        )
        
        if file_path:
            try:
                self.status_bar.showMessage("Loading process parts file...")
                
                # Load process parts file
                if file_path.lower().endswith(('.xlsx', '.xls')):
                    self.process_parts = pd.read_excel(file_path)
                else:
                    self.process_parts = pd.read_csv(file_path)
                
                # Clean process parts data
                self.process_parts = self.clean_process_parts(self.process_parts)
                
                # Run process analysis if spare parts data is available
                if self.data is not None and not self.data.empty:
                    self.run_process_analysis()
                
                self.status_bar.showMessage(f"Process parts loaded: {len(self.process_parts)} records")
                QMessageBox.information(self, "Success", f"Process parts file loaded successfully!\n\nRecords: {len(self.process_parts):,}\nProcesses: {self.process_parts['Process'].nunique():,}")
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load process parts file: {str(e)}")
                self.status_bar.showMessage("Error loading process parts file")
                
    def clean_process_parts(self, df):
        """Clean and validate process parts data"""
        try:
            df_clean = df.copy()
            
            # Ensure required columns exist
            required_cols = ['Process', 'Item Number', 'Part Name']
            missing_cols = [col for col in required_cols if col not in df_clean.columns]
            
            if missing_cols:
                # Try to find similar columns
                for col in missing_cols:
                    if col == 'Process':
                        # Look for process-related columns
                        process_cols = [c for c in df_clean.columns if 'process' in c.lower() or 'operation' in c.lower()]
                        if process_cols:
                            df_clean['Process'] = df_clean[process_cols[0]]
                        else:
                            df_clean['Process'] = 'Unknown'
                    elif col == 'Item Number':
                        # Look for item-related columns
                        item_cols = [c for c in df_clean.columns if 'item' in c.lower() or 'part' in c.lower() or 'number' in c.lower()]
                        if item_cols:
                            df_clean['Item Number'] = df_clean[item_cols[0]]
                        else:
                            df_clean['Item Number'] = 'Unknown'
                    elif col == 'Part Name':
                        # Look for name-related columns
                        name_cols = [c for c in df_clean.columns if 'name' in c.lower() or 'description' in c.lower()]
                        if name_cols:
                            df_clean['Part Name'] = df_clean[name_cols[0]]
                        else:
                            df_clean['Part Name'] = 'Unknown'
            
            # Clean data
            df_clean['Process'] = df_clean['Process'].astype(str).str.strip()
            df_clean['Item Number'] = df_clean['Item Number'].astype(str).str.strip()
            df_clean['Part Name'] = df_clean['Part Name'].astype(str).str.strip()
            
            # Remove rows with missing critical data
            df_clean = df_clean.dropna(subset=['Process', 'Item Number'])
            
            return df_clean
            
        except Exception as e:
            print(f"Error cleaning process parts: {e}")
            return df
            
    def run_process_analysis(self):
        """Run process-based analysis comparing spare parts usage with process parts"""
        if self.data is None or self.data.empty or self.process_parts is None:
            return
            
        try:
            self.status_bar.showMessage("Running process analysis...")
            
            # Find the correct column names from spare parts data
            item_col = None
            qty_col = None
            stock_col = None
            part_col = None
            desc_col = None
            
            for col in self.data.columns:
                if 'item' in col.lower() and 'number' in col.lower():
                    item_col = col
                elif 'req' in col.lower() and 'qty' in col.lower():
                    qty_col = col
                elif 'hand' in col.lower():
                    stock_col = col
                elif 'part' in col.lower() and 'name' in col.lower():
                    part_col = col
                elif 'description' in col.lower() and '2' in col:
                    desc_col = col
            
            if not item_col or not qty_col:
                return
            
            # Create detailed process analysis showing individual parts
            detailed_analysis = []
            
            for _, process_row in self.process_parts.iterrows():
                process_name = process_row['Process']
                item_number = process_row['Item Number']
                part_name = process_row['Part Name']
                
                # Find matching spare parts data
                matching_data = self.data[self.data[item_col] == item_number]
                
                if not matching_data.empty:
                    # Calculate usage metrics for this specific part
                    total_usage = matching_data[qty_col].sum()
                    usage_count = len(matching_data)
                    avg_usage = total_usage / usage_count if usage_count > 0 else 0
                    current_stock = matching_data[stock_col].iloc[0] if stock_col and stock_col in matching_data.columns else 0
                    
                    # Calculate daily metrics
                    daily_usage = total_usage / 30 if total_usage > 0 else 0  # Assuming 30 days
                    daily_std = matching_data[qty_col].std() if len(matching_data) > 1 else 0
                    
                    # Calculate safety stock (using 95% confidence level, Z = 1.65)
                    Z = 1.65
                    lead_time = 30  # Default lead time in days
                    safety_stock = Z * np.sqrt(lead_time) * daily_std
                    reorder_point = (daily_usage * lead_time) + safety_stock
                    
                    # Determine criticality
                    if daily_usage > 10:
                        criticality = "CRITICAL"
                    elif daily_usage > 5:
                        criticality = "HIGH"
                    elif daily_usage > 2:
                        criticality = "MEDIUM"
                    else:
                        criticality = "LOW"
                    
                    detailed_analysis.append({
                        'Process': process_name,
                        'Item Number': item_number,
                        'Part Name': part_name,
                        'Total Usage': total_usage,
                        'Usage Count': usage_count,
                        'Avg Usage per Request': avg_usage,
                        'Current Stock': current_stock,
                        'D_Mean_per_Day': daily_usage,
                        'D_Std_per_Day': daily_std,
                        'Safety Stock': safety_stock,
                        'Reorder Point': reorder_point,
                        'Criticality': criticality
                    })
                else:
                    # Part exists in process but no usage data found
                    detailed_analysis.append({
                        'Process': process_name,
                        'Item Number': item_number,
                        'Part Name': part_name,
                        'Total Usage': 0,
                        'Usage Count': 0,
                        'Avg Usage per Request': 0,
                        'Current Stock': 0,
                        'D_Mean_per_Day': 0,
                        'D_Std_per_Day': 0,
                        'Safety Stock': 0,
                        'Reorder Point': 0,
                        'Criticality': 'NO DATA'
                    })
            
            # Convert to DataFrame and sort
            self.process_analysis = pd.DataFrame(detailed_analysis)
            if not self.process_analysis.empty:
                self.process_analysis = self.process_analysis.sort_values(['Process', 'Total Usage'], ascending=[True, False])
            
            self.display_process_analysis()
            
            # Update process selector dropdown
            self.update_process_selector()
            
            self.status_bar.showMessage(f"Process analysis completed: {len(self.process_analysis)} parts analyzed")
            
        except Exception as e:
            print(f"Error in process analysis: {e}")
            self.status_bar.showMessage("Process analysis failed")
            
    def calculate_process_criticality(self, total_usage):
        """Calculate process criticality based on total usage"""
        if total_usage > 1000:
            return "CRITICAL"
        elif total_usage > 500:
            return "HIGH"
        elif total_usage > 100:
            return "MEDIUM"
        else:
            return "LOW"
            
    def display_process_analysis(self):
        """Display detailed process analysis results showing individual parts"""
        if self.process_analysis is None:
            return
            
        # Set up table with detailed columns
        self.process_table.setRowCount(len(self.process_analysis))
        columns = ['Process', 'Item Number', 'Part Name', 'Total Usage', 'Usage Count', 'Avg Usage per Request', 
                  'Current Stock', 'D_Mean_per_Day', 'D_Std_per_Day', 'Safety Stock', 'Reorder Point', 'Criticality']
        self.process_table.setColumnCount(len(columns))
        self.process_table.setHorizontalHeaderLabels(columns)
        
        # Populate table
        for i, row in self.process_analysis.iterrows():
            for j, col in enumerate(columns):
                value = row[col]
                if isinstance(value, float):
                    if col in ['Avg Usage per Request', 'D_Mean_per_Day', 'D_Std_per_Day', 'Safety Stock', 'Reorder Point']:
                        value = f"{value:.2f}"
                    else:
                        value = f"{value:.0f}"
                item = QTableWidgetItem(str(value))
                
                # Color code criticality levels
                if col == 'Criticality':
                    if value == 'CRITICAL':
                        item.setBackground(QColor(255, 200, 200))  # Light red
                    elif value == 'HIGH':
                        item.setBackground(QColor(255, 255, 200))  # Light yellow
                    elif value == 'MEDIUM':
                        item.setBackground(QColor(200, 255, 200))  # Light green
                    elif value == 'LOW':
                        item.setBackground(QColor(200, 200, 255))  # Light blue
                    elif value == 'NO DATA':
                        item.setBackground(QColor(240, 240, 240))  # Light gray
                
                self.process_table.setItem(i, j, item)
                
        # Auto-resize columns
        self.process_table.resizeColumnsToContents()
        
    def update_process_selector(self):
        """Update the process selector dropdown with available processes"""
        if self.process_parts is None:
            return
            
        # Clear existing items
        self.process_selector.clear()
        self.process_selector.addItem("All Processes")
        
        # Add unique processes
        unique_processes = sorted(self.process_parts['Process'].unique())
        for process in unique_processes:
            self.process_selector.addItem(process)
            
        # Enable the filter button if processes are available
        self.apply_filter_btn.setEnabled(len(unique_processes) > 0)
        
    def on_process_selection_changed(self, process_name):
        """Handle process selection change"""
        print(f"Debug - Process selection changed to: '{process_name}'")
        if process_name == "All Processes":
            self.apply_filter_btn.setEnabled(True)  # Allow showing all processes
        else:
            self.apply_filter_btn.setEnabled(True)  # Allow filtering to specific process
            
    def apply_process_filter(self):
        """Apply process filter to show only parts from selected process"""
        selected_process = self.process_selector.currentText()
        
        if selected_process == "All Processes":
            # Show all process parts
            if self.process_analysis is not None:
                self.display_process_analysis()
            return
            
        if self.process_analysis is None:
            QMessageBox.warning(self, "Warning", "Please run process analysis first")
            return
            
        try:
            self.status_bar.showMessage(f"Filtering process analysis for: {selected_process}")
            
            # Debug: Print unique processes in data
            print(f"Debug - Available processes: {self.process_analysis['Process'].unique()}")
            print(f"Debug - Selected process: '{selected_process}'")
            print(f"Debug - Total records before filter: {len(self.process_analysis)}")
            
            # Filter the process analysis to show only items from the selected process
            filtered_results = self.process_analysis[self.process_analysis['Process'] == selected_process].copy()
            
            # Reset index to ensure proper display
            filtered_results = filtered_results.reset_index(drop=True)
            
            print(f"Debug - Records after filter: {len(filtered_results)}")
            
            if filtered_results.empty:
                QMessageBox.information(self, "Info", f"No parts found for process: {selected_process}")
                return
                
            # Display filtered process results
            self.display_filtered_process_results(filtered_results, selected_process)
            
            self.status_bar.showMessage(f"Filtered process analysis for {selected_process}: {len(filtered_results)} parts")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to apply process filter: {str(e)}")
            self.status_bar.showMessage("Process filter failed")
            
    def display_filtered_process_results(self, filtered_results, process_name):
        """Display filtered process analysis results for a specific process"""
        if filtered_results is None or filtered_results.empty:
            return
            
        print(f"Debug - Displaying {len(filtered_results)} filtered results for {process_name}")
        print(f"Debug - First few items: {filtered_results[['Process', 'Item Number', 'Part Name']].head()}")
        
        # Clear table first
        self.process_table.clear()
        
        # Set up table with same columns as process analysis
        self.process_table.setRowCount(len(filtered_results))
        columns = ['Process', 'Item Number', 'Part Name', 'Total Usage', 'Usage Count', 'Avg Usage per Request', 
                  'Current Stock', 'D_Mean_per_Day', 'D_Std_per_Day', 'Safety Stock', 'Reorder Point', 'Criticality']
        self.process_table.setColumnCount(len(columns))
        self.process_table.setHorizontalHeaderLabels(columns)
        
        # Add process name to header
        self.process_table.setWindowTitle(f"Process Analysis - {process_name}")
        
        # Populate table - use iloc for position-based indexing
        for i in range(len(filtered_results)):
            row = filtered_results.iloc[i]
            for j, col in enumerate(columns):
                value = row[col]
                if isinstance(value, float):
                    if col in ['Avg Usage per Request', 'D_Mean_per_Day', 'D_Std_per_Day', 'Safety Stock', 'Reorder Point']:
                        value = f"{value:.2f}"
                    else:
                        value = f"{value:.0f}"
                item = QTableWidgetItem(str(value))
                
                # Color code criticality levels
                if col == 'Criticality':
                    if value == 'CRITICAL':
                        item.setBackground(QColor(255, 200, 200))  # Light red
                    elif value == 'HIGH':
                        item.setBackground(QColor(255, 255, 200))  # Light yellow
                    elif value == 'MEDIUM':
                        item.setBackground(QColor(200, 255, 200))  # Light green
                    elif value == 'LOW':
                        item.setBackground(QColor(200, 200, 255))  # Light blue
                    elif value == 'NO DATA':
                        item.setBackground(QColor(240, 240, 240))  # Light gray
                
                self.process_table.setItem(i, j, item)
                
        # Auto-resize columns
        self.process_table.resizeColumnsToContents()
        
    def display_filtered_analysis_results(self, filtered_results, process_name):
        """Display filtered analysis results for a specific process"""
        if filtered_results is None or filtered_results.empty:
            return
            
        # Set up table
        self.safety_table.setRowCount(len(filtered_results))
        columns = ['Item Number', 'Part Name', 'Description 2', 'Total Usage', 'Usage Count', 'D_Mean_per_Day', 
                  'D_Std_per_Day', 'Current Stock', 'Safety Stock', 'Reorder Point', 'Criticality']
        self.safety_table.setColumnCount(len(columns))
        self.safety_table.setHorizontalHeaderLabels(columns)
        
        # Add process name to header
        self.safety_table.setWindowTitle(f"Safety Stock Analysis - {process_name}")
        
        # Populate table
        for i, row in filtered_results.iterrows():
            for j, col in enumerate(columns):
                value = row[col]
                if isinstance(value, float):
                    if col in ['D_Mean_per_Day', 'D_Std_per_Day', 'Safety Stock', 'Reorder Point']:
                        value = f"{value:.2f}"
                    else:
                        value = f"{value:.0f}"
                item = QTableWidgetItem(str(value))
                self.safety_table.setItem(i, j, item)
                
        # Auto-resize columns
        self.safety_table.resizeColumnsToContents()
            
    def load_files(self, file_paths):
        """Load and process files"""
        try:
            self.status_bar.showMessage("Loading files...")
            self.progress_bar.setVisible(True)
            self.progress_bar.setMaximum(len(file_paths))
            
            all_data = []
            
            for i, file_path in enumerate(file_paths):
                self.progress_bar.setValue(i + 1)
                self.status_bar.showMessage(f"Loading {os.path.basename(file_path)}...")
                
                # Load file based on extension
                if file_path.lower().endswith(('.xlsx', '.xls')):
                    df = pd.read_excel(file_path)
                elif file_path.lower().endswith('.csv'):
                    df = pd.read_csv(file_path)
                else:
                    df = pd.read_csv(file_path, sep='\t')
                
                # Clean and convert data types
                df = self.clean_dataframe(df)
                    
                # Filter for emergency requests (EM or em in remarks)
                if 'Remark' in df.columns:
                    df = df[df['Remark'].str.contains('EM|em', na=False, regex=True)]
                elif 'Remarks' in df.columns:
                    df = df[df['Remarks'].str.contains('EM|em', na=False, regex=True)]
                    
                all_data.append(df)
                
            if all_data:
                self.data = pd.concat(all_data, ignore_index=True)
                self.display_data()
                self.analyze_btn.setEnabled(True)
                self.status_bar.showMessage(f"Loaded {len(self.data)} emergency request records")
            else:
                self.status_bar.showMessage("No emergency request data found")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load files: {str(e)}")
            self.status_bar.showMessage("Error loading files")
        finally:
            self.progress_bar.setVisible(False)
            
    def clean_dataframe(self, df):
        """Clean and convert data types in the dataframe - SIMPLIFIED APPROACH"""
        try:
            # Make a copy to avoid modifying original
            df_clean = df.copy()
            
            print("=== STARTING DATA CLEANING ===")
            print(f"Original columns: {list(df_clean.columns)}")
            
            # SIMPLE APPROACH: Just convert everything to strings first, then handle specific cases
            for col in df_clean.columns:
                # Convert everything to string first
                df_clean[col] = df_clean[col].astype(str)
                
                # Handle date columns - SIMPLE METHOD
                if any(keyword in col.lower() for keyword in ['date', 'time', 'requested']):
                    print(f"\nProcessing DATE column: {col}")
                    print(f"Sample values: {df_clean[col].head(3).tolist()}")
                    
                    # Try to convert the first few values to see what format we have
                    sample_values = df_clean[col].head(5).tolist()
                    
                    # Check if they look like huge numbers (your case)
                    if any(len(str(val)) > 10 for val in sample_values):
                        print(f"Detected huge numbers in {col} - treating as text for now")
                        # Keep as text for now - we'll handle this in analysis
                        df_clean[col] = df_clean[col].astype(str)
                    else:
                        # Try normal date conversion
                        try:
                            df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce')
                            print(f"Successfully converted {col} to datetime")
                        except:
                            print(f"Failed to convert {col} to datetime - keeping as text")
                            df_clean[col] = df_clean[col].astype(str)
                
                # Handle quantity columns - SIMPLE METHOD
                elif any(keyword in col.lower() for keyword in ['quantity', 'qty', 'amount', 'req']):
                    print(f"\nProcessing QUANTITY column: {col}")
                    try:
                        # Convert to numeric, handle errors
                        df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce')
                        # Fill NaN with 0
                        df_clean[col] = df_clean[col].fillna(0)
                        # Remove negative values
                        df_clean[col] = df_clean[col].abs()
                        print(f"Successfully converted {col} to numeric")
                    except:
                        print(f"Failed to convert {col} to numeric - keeping as text")
                        df_clean[col] = df_clean[col].astype(str)
                
                # Handle text columns
                else:
                    # Just ensure it's string and limit length
                    df_clean[col] = df_clean[col].str[:100]  # Limit to 100 characters
                    df_clean[col] = df_clean[col].fillna('')
            
            print("\n=== DATA CLEANING COMPLETED ===")
            print(f"Final column types:")
            for col in df_clean.columns:
                print(f"  {col}: {df_clean[col].dtype}")
            
            return df_clean
            
        except Exception as e:
            print(f"Error in clean_dataframe: {e}")
            return df
            
    def display_data(self):
        """Display loaded data in the table"""
        if self.data is None or self.data.empty:
            return
            
        # Set up table
        self.data_table.setRowCount(len(self.data))
        self.data_table.setColumnCount(len(self.data.columns))
        self.data_table.setHorizontalHeaderLabels(self.data.columns)
        
        # Populate table
        for i in range(len(self.data)):
            for j in range(len(self.data.columns)):
                value = str(self.data.iloc[i, j])
                item = QTableWidgetItem(value)
                self.data_table.setItem(i, j, item)
                
        # Update summary
        self.update_summary()
        
    def update_summary(self):
        """Update the summary tab"""
        if self.data is None or self.data.empty:
            return
            
        # Find date and quantity columns dynamically
        date_col = None
        quantity_col = None
        item_col = None
        
        for col in self.data.columns:
            if any(keyword in col.lower() for keyword in ['date', 'time', 'requested', 'created']):
                date_col = col
            elif any(keyword in col.lower() for keyword in ['quantity', 'qty', 'amount', 'requested']):
                quantity_col = col
            elif any(keyword in col.lower() for keyword in ['item', 'part', 'number', 'code']):
                item_col = col
        
        summary_text = f"""
📊 DATA SUMMARY
{'='*50}

📁 Records Loaded: {len(self.data):,}
"""
        
        if date_col:
            try:
                # Check if it's actually a datetime column
                if pd.api.types.is_datetime64_any_dtype(self.data[date_col]):
                    date_range = f"{self.data[date_col].min()} to {self.data[date_col].max()}"
                    summary_text += f"📅 Date Range: {date_range}\n"
                else:
                    # It's a text column with huge numbers - show sample values
                    sample_dates = self.data[date_col].head(3).tolist()
                    summary_text += f"📅 Date Column: {date_col} (showing as text)\n"
                    summary_text += f"   Sample values: {sample_dates}\n"
            except Exception as e:
                summary_text += f"📅 Date Column: {date_col} (error: {str(e)})\n"
        
        if item_col:
            try:
                summary_text += f"🔢 Unique Items: {self.data[item_col].nunique():,}\n"
            except:
                summary_text += f"🔢 Item Column: {item_col}\n"
        
        if quantity_col:
            try:
                summary_text += f"📦 Total Quantity Requested: {self.data[quantity_col].sum():,}\n"
            except:
                summary_text += f"📦 Quantity Column: {quantity_col}\n"
        
        summary_text += f"""
📈 COLUMN INFORMATION:
{'='*50}
"""
        
        for col in self.data.columns:
            try:
                if pd.api.types.is_numeric_dtype(self.data[col]):
                    summary_text += f"\n{col}:"
                    summary_text += f"\n  - Min: {self.data[col].min():,}"
                    summary_text += f"\n  - Max: {self.data[col].max():,}"
                    summary_text += f"\n  - Mean: {self.data[col].mean():.2f}"
                    summary_text += f"\n  - Total: {self.data[col].sum():,}"
                elif pd.api.types.is_datetime64_any_dtype(self.data[col]):
                    summary_text += f"\n{col}: Date column with {self.data[col].nunique():,} unique dates"
                else:
                    summary_text += f"\n{col}: {self.data[col].nunique():,} unique values"
            except Exception as e:
                summary_text += f"\n{col}: Error processing column - {str(e)}"
                
        self.summary_text.setPlainText(summary_text)
        
    def run_analysis(self):
        """Run safety stock analysis"""
        if self.data is None or self.data.empty:
            QMessageBox.warning(self, "Warning", "No data loaded for analysis")
            return
            
        try:
            self.status_bar.showMessage("Running analysis...")
            self.progress_bar.setVisible(True)
            self.progress_bar.setMaximum(100)
            
            # Find the correct column names from your data
            item_col = None
            qty_col = None
            stock_col = None
            part_col = None
            desc_col = None
            
            for col in self.data.columns:
                if 'item' in col.lower() and 'number' in col.lower():
                    item_col = col
                elif 'req' in col.lower() and 'qty' in col.lower():
                    qty_col = col
                elif 'hand' in col.lower():
                    stock_col = col
                elif 'part' in col.lower() and 'name' in col.lower():
                    part_col = col
                elif 'description' in col.lower() and '2' in col:
                    desc_col = col
            
            if not item_col or not qty_col:
                QMessageBox.critical(self, "Error", "Required columns not found. Need Item Number and Requested Quantity columns.")
                return
            
            # Group by item number for analysis
            grouped = self.data.groupby(item_col).agg({
                qty_col: ['sum', 'count', 'mean', 'std'],
                stock_col: 'first' if stock_col else 'first',
                part_col: 'first' if part_col else 'first',
                desc_col: 'first' if desc_col else 'first'
            }).reset_index()
            
            # Flatten column names
            if stock_col and part_col and desc_col:
                grouped.columns = ['Item Number', 'Total Usage', 'Usage Count', 'D_Mean_per_Day', 'D_Std_per_Day', 'Current Stock', 'Part Name', 'Description 2']
            elif stock_col and part_col:
                grouped.columns = ['Item Number', 'Total Usage', 'Usage Count', 'D_Mean_per_Day', 'D_Std_per_Day', 'Current Stock', 'Part Name']
                grouped['Description 2'] = 'Unknown'
            elif stock_col and desc_col:
                grouped.columns = ['Item Number', 'Total Usage', 'Usage Count', 'D_Mean_per_Day', 'D_Std_per_Day', 'Current Stock', 'Description 2']
                grouped['Part Name'] = 'Unknown'
            elif part_col and desc_col:
                grouped.columns = ['Item Number', 'Total Usage', 'Usage Count', 'D_Mean_per_Day', 'D_Std_per_Day', 'Part Name', 'Description 2']
                grouped['Current Stock'] = 0
            elif stock_col:
                grouped.columns = ['Item Number', 'Total Usage', 'Usage Count', 'D_Mean_per_Day', 'D_Std_per_Day', 'Current Stock']
                grouped['Part Name'] = 'Unknown'
                grouped['Description 2'] = 'Unknown'
            elif part_col:
                grouped.columns = ['Item Number', 'Total Usage', 'Usage Count', 'D_Mean_per_Day', 'D_Std_per_Day', 'Part Name']
                grouped['Current Stock'] = 0
                grouped['Description 2'] = 'Unknown'
            elif desc_col:
                grouped.columns = ['Item Number', 'Total Usage', 'Usage Count', 'D_Mean_per_Day', 'D_Std_per_Day', 'Description 2']
                grouped['Current Stock'] = 0
                grouped['Part Name'] = 'Unknown'
            else:
                grouped.columns = ['Item Number', 'Total Usage', 'Usage Count', 'D_Mean_per_Day', 'D_Std_per_Day']
                grouped['Current Stock'] = 0
                grouped['Part Name'] = 'Unknown'
                grouped['Description 2'] = 'Unknown'
            
            # Calculate safety stock (using 95% confidence level, Z = 1.65)
            Z = 1.65
            lead_time = 30  # Default lead time in days
            
            grouped['Safety Stock'] = Z * np.sqrt(lead_time) * grouped['D_Std_per_Day']
            grouped['Reorder Point'] = (grouped['D_Mean_per_Day'] * lead_time) + grouped['Safety Stock']
            grouped['Criticality'] = grouped['D_Mean_per_Day'].apply(self.calculate_criticality)
            
            self.analysis_results = grouped
            self.display_analysis_results()
            
            # Temporarily disable charts to avoid matplotlib errors
            try:
                self.create_charts()
            except Exception as e:
                print(f"Charts disabled due to error: {e}")
                # Create a simple text message instead
                self.charts_canvas.figure.clear()
                ax = self.charts_canvas.figure.add_subplot(111)
                ax.text(0.5, 0.5, 'Charts temporarily disabled\nAnalysis completed successfully!', 
                       ha='center', va='center', transform=ax.transAxes, fontsize=12)
                ax.set_title('Analysis Results Available')
                self.charts_canvas.draw()
            
            self.export_btn.setEnabled(True)
            self.status_bar.showMessage("Analysis completed successfully")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Analysis failed: {str(e)}")
            self.status_bar.showMessage("Analysis failed")
        finally:
            self.progress_bar.setVisible(False)
            
    def calculate_criticality(self, daily_usage):
        """Calculate criticality level based on daily usage"""
        if daily_usage > 10:
            return "CRITICAL"
        elif daily_usage > 5:
            return "HIGH"
        elif daily_usage > 2:
            return "MEDIUM"
        else:
            return "LOW"
            
    def display_analysis_results(self):
        """Display analysis results in the safety stock table"""
        if self.analysis_results is None:
            return
            
        # Set up table
        self.safety_table.setRowCount(len(self.analysis_results))
        columns = ['Item Number', 'Part Name', 'Description 2', 'Total Usage', 'Usage Count', 'D_Mean_per_Day', 
                  'D_Std_per_Day', 'Current Stock', 'Safety Stock', 'Reorder Point', 'Criticality']
        self.safety_table.setColumnCount(len(columns))
        self.safety_table.setHorizontalHeaderLabels(columns)
        
        # Populate table
        for i, row in self.analysis_results.iterrows():
            for j, col in enumerate(columns):
                value = row[col]
                if isinstance(value, float):
                    if col in ['D_Mean_per_Day', 'D_Std_per_Day', 'Safety Stock', 'Reorder Point']:
                        value = f"{value:.2f}"
                    else:
                        value = f"{value:.0f}"
                item = QTableWidgetItem(str(value))
                self.safety_table.setItem(i, j, item)
                
        # Auto-resize columns
        self.safety_table.resizeColumnsToContents()
        
    def create_charts(self):
        """Create analysis charts - SIMPLIFIED VERSION"""
        if self.analysis_results is None:
            return
            
        try:
            # Clear previous charts and create new figure
            self.charts_canvas.figure.clear()
            
            # Create a simple success message instead of complex charts
            ax = self.charts_canvas.figure.add_subplot(111)
            ax.text(0.5, 0.5, 'Analysis Completed Successfully!\n\n' + 
                   f'Total Items Analyzed: {len(self.analysis_results)}\n' +
                   f'Critical Items: {len(self.analysis_results[self.analysis_results["Criticality"] == "CRITICAL"])}\n' +
                   f'High Priority: {len(self.analysis_results[self.analysis_results["Criticality"] == "HIGH"])}\n' +
                   f'Medium Priority: {len(self.analysis_results[self.analysis_results["Criticality"] == "MEDIUM"])}\n' +
                   f'Low Priority: {len(self.analysis_results[self.analysis_results["Criticality"] == "LOW"])}', 
                   ha='center', va='center', transform=ax.transAxes, fontsize=11)
            ax.set_title('Safety Stock Analysis Results', fontsize=14, fontweight='bold')
            ax.axis('off')  # Hide axes
            
            # Refresh canvas
            self.charts_canvas.draw()
            
        except Exception as e:
            print(f"Error creating charts: {e}")
            # Create a simple error message chart
            self.charts_canvas.figure.clear()
            ax = self.charts_canvas.figure.add_subplot(111)
            ax.text(0.5, 0.5, 'Charts temporarily disabled\nAnalysis completed successfully!', 
                   ha='center', va='center', transform=ax.transAxes, fontsize=12)
            ax.set_title('Analysis Results Available')
            ax.axis('off')
            self.charts_canvas.draw()
        
    def export_results(self):
        """Export analysis results"""
        if self.analysis_results is None:
            QMessageBox.warning(self, "Warning", "No analysis results to export")
            return
            
        try:
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Export Results",
                f"Safety_Stock_Analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                "Excel Files (*.xlsx);;CSV Files (*.csv)"
            )
            
            if file_path:
                if file_path.endswith('.xlsx'):
                    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                        self.analysis_results.to_excel(writer, sheet_name='Safety Stock Analysis', index=False)
                        if self.data is not None:
                            self.data.to_excel(writer, sheet_name='Raw Data', index=False)
                        if self.process_analysis is not None:
                            self.process_analysis.to_excel(writer, sheet_name='Process Analysis', index=False)
                        if self.process_parts is not None:
                            self.process_parts.to_excel(writer, sheet_name='Process Parts', index=False)
                else:
                    self.analysis_results.to_csv(file_path, index=False)
                    
                QMessageBox.information(self, "Success", f"Results exported to:\n{file_path}")
                self.status_bar.showMessage(f"Results exported to {os.path.basename(file_path)}")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Export failed: {str(e)}")
            
    def clear_data(self):
        """Clear all loaded data and results"""
        if self.data is None and self.process_parts is None:
            # No data to clear
            QMessageBox.information(self, "Info", "No data to clear.")
            return
            
        reply = QMessageBox.question(
            self, 
            "⚠️ Clear Data Confirmation", 
            "Are you sure you want to clear ALL data and results?\n\nThis action cannot be undone.\n\nClick 'Yes' to clear or 'No' to cancel.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No  # Default to No for safety
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self.data = None
            self.analysis_results = None
            self.process_parts = None
            self.process_analysis = None
            
            # Clear tables
            self.data_table.setRowCount(0)
            self.data_table.setColumnCount(0)
            self.safety_table.setRowCount(0)
            self.safety_table.setColumnCount(0)
            self.process_table.setRowCount(0)
            self.process_table.setColumnCount(0)
            
            # Clear process selector
            self.process_selector.clear()
            self.process_selector.addItem("All Processes")
            self.apply_filter_btn.setEnabled(False)
            
            # Clear summary
            self.summary_text.clear()
            
            # Clear charts
            self.charts_canvas.figure.clear()
            self.charts_canvas.draw()
            
            # Disable buttons
            self.analyze_btn.setEnabled(False)
            self.export_btn.setEnabled(False)
            
            # Update status
            self.status_bar.showMessage("Data cleared successfully")
            
            QMessageBox.information(self, "Success", "All data and results have been cleared successfully.")
            
    def closeEvent(self, event):
        """Handle application close event with safety confirmation"""
        if self.data is not None or self.process_parts is not None:
            # Check if there's data that might be lost
            reply = QMessageBox.question(
                self,
                "⚠️ Exit Confirmation",
                "You have unsaved data and analysis results.\n\nAre you sure you want to exit?\n\nClick 'Yes' to exit or 'No' to stay.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No  # Default to No for safety
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                event.accept()
            else:
                event.ignore()
        else:
            # No data loaded, safe to exit
            event.accept()
    
    def show_about(self):
        """Show about dialog"""
        QMessageBox.about(
            self,
            "About Safety Stock Analyzer",
            """
            <h3>Safety Stock Analyzer - Professional Edition</h3>
            <p>A powerful tool for analyzing spare parts usage and calculating safety stock levels.</p>
            <p><b>Features:</b></p>
            <ul>
                <li>📁 Drag & Drop file loading</li>
                <li>📊 Real-time data display</li>
                <li>🔍 Advanced safety stock analysis</li>
                <li>📈 Professional charts and visualizations</li>
                <li>📤 Export results to Excel/CSV</li>
            </ul>
            <p><b>Version:</b> 1.0 Professional</p>
            <p><b>Built with:</b> Python + PyQt6</p>
            """
        )

def main():
    """Main application entry point"""
    app = QApplication(sys.argv)
    app.setApplicationName("Safety Stock Analyzer")
    app.setApplicationVersion("1.0")
    
    # Set application icon (if available)
    try:
        app.setWindowIcon(QIcon("icon.ico"))
    except:
        pass
    
    # Create and show main window
    window = SafetyStockAnalyzer()
    window.show()
    
    # Start event loop
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
