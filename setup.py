from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

with open("requirements.txt", "r", encoding="utf-8") as fh:
    requirements = [line.strip() for line in fh if line.strip() and not line.startswith("#")]

setup(
    name="safety-stock-analyzer",
    version="1.0.0",
    author="Safety Stock Analyzer Team",
    author_email="your.email@example.com",
    description="A powerful desktop application for analyzing spare parts usage and calculating safety stock levels",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/safety-stock-analyzer",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Manufacturing",
        "Intended Audience :: End Users/Desktop",
        "Topic :: Office/Business :: Financial :: Inventory",
        "Topic :: Scientific/Engineering :: Information Analysis",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Operating System :: OS Independent",
        "Operating System :: Microsoft :: Windows",
        "Operating System :: POSIX :: Linux",
        "Operating System :: MacOS",
    ],
    python_requires=">=3.8",
    install_requires=requirements,
    extras_require={
        "dev": [
            "pytest>=6.0",
            "flake8>=3.8",
            "black>=21.0",
            "isort>=5.0",
        ],
        "build": [
            "pyinstaller>=4.0",
            "cx_Freeze>=6.0",
        ],
    },
    entry_points={
        "console_scripts": [
            "safety-stock-analyzer=safety_stock_analyzer:main",
        ],
        "gui_scripts": [
            "safety-stock-analyzer-gui=safety_stock_analyzer:main",
        ],
    },
    include_package_data=True,
    package_data={
        "": ["*.txt", "*.md", "*.bat", "*.yml"],
    },
    keywords="safety stock, inventory management, spare parts, manufacturing, analysis, desktop application, PyQt6",
    project_urls={
        "Bug Reports": "https://github.com/yourusername/safety-stock-analyzer/issues",
        "Source": "https://github.com/yourusername/safety-stock-analyzer",
        "Documentation": "https://github.com/yourusername/safety-stock-analyzer#readme",
    },
)
