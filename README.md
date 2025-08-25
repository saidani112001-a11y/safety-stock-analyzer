# 🏭 Safety Stock Analyzer - Professional Edition

A powerful desktop application for analyzing spare parts usage and calculating safety stock levels with process-based analysis capabilities.

## ✨ Features

- **📁 Drag & Drop File Loading**: Support for Excel (.xlsx), CSV, and text files
- **📊 Real-time Data Display**: Professional tables with immediate data visualization
- **🔍 Advanced Safety Stock Analysis**: Statistical calculations including D_Mean_per_Day and D_Std_per_Day
- **🏭 Process-Based Analysis**: Upload process parts files and analyze usage by manufacturing processes
- **📈 Professional Charts**: Beautiful visualizations and dashboards
- **📤 Export Capabilities**: Export results to Excel and PDF formats
- **🛡️ Safety Features**: Exit confirmations and data loss prevention
- **💻 Modern UI**: Built with PyQt6 for a professional desktop experience

## 🚀 Quick Start

### Prerequisites
- Python 3.8 or higher
- Windows 10/11 (for .exe build)

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/safety-stock-analyzer.git
   cd safety-stock-analyzer
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**
   ```bash
   python safety_stock_analyzer.py
   ```

### Windows Users - Quick Launch
Double-click `run_app.bat` for automatic dependency installation and launch.

## 🏗️ Building Executable

To create a standalone .exe file:

```bash
# Windows
build_exe.bat

# Or manually
pip install pyinstaller
pyinstaller --onefile --windowed safety_stock_analyzer.py
```

## 📖 Usage Guide

### 1. Load Spare Parts Data
- Drag & drop your Excel/CSV file onto the upload area
- Or click "Browse Files" to select manually
- Supported formats: .xlsx, .csv, .txt

### 2. Run Safety Stock Analysis
- Click "🔍 Run Analysis" button
- View results in the Safety Stock tab
- Check criticality levels and recommendations

### 3. Process Analysis (Optional)
- Go to "🏭 Process Analysis" tab
- Upload your process parts mapping file
- Select specific processes to filter results
- Compare usage patterns across manufacturing processes

### 4. Export Results
- Click "📤 Export Results" to save to Excel
- Includes all analysis data and charts

## 📊 Data Requirements

### Spare Parts File
Your Excel/CSV should contain:
- **Item Number**: Part identifier
- **Date**: Usage date (various formats supported)
- **Quantity**: Amount used
- **Part Name**: Part description
- **Description 2**: Additional part details
- **Current Stock**: Available inventory
- **Remarks**: Filter for "em" or "EM" entries

### Process Parts File (Optional)
Mapping file should contain:
- **Process**: Manufacturing process name
- **Item Number**: Part identifier
- **Part Name**: Part description

## 🔧 Technical Details

- **Framework**: PyQt6 (Professional GUI)
- **Data Processing**: Pandas + NumPy
- **Charts**: Matplotlib + Seaborn
- **File Support**: OpenPyXL, xlrd
- **Architecture**: Object-oriented, modular design

## 🛡️ Safety Features

- **Exit Confirmation**: Prevents accidental data loss
- **Clear Data Confirmation**: Safe data clearing with warnings
- **Data Validation**: Robust error handling and data cleaning
- **Auto-save Prevention**: User-controlled data management

## 📁 Project Structure

```
safety-stock-analyzer/
├── safety_stock_analyzer.py    # Main application
├── requirements.txt             # Python dependencies
├── run_app.bat                 # Windows launcher
├── build_exe.bat               # Executable builder
├── README.md                   # This file
├── LICENSE                     # License information
└── examples/                   # Sample data files
    ├── sample_spare_parts.xlsx
    └── sample_process_parts.xlsx
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support

- **Issues**: Report bugs on GitHub Issues
- **Features**: Request new features via GitHub Issues
- **Documentation**: Check this README and code comments

## 🎯 Roadmap

- [ ] Database integration for persistent storage
- [ ] Multi-language support
- [ ] Cloud synchronization
- [ ] Advanced reporting templates
- [ ] API integration capabilities

## ⭐ Star This Project

If this tool helps you manage your spare parts inventory, please give it a star! ⭐

---

**Built with ❤️ using Python + PyQt6**


