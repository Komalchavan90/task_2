import subprocess
import sys

# List of required libraries
required_libs = [
    "P4Python",
    "mysql-connector-python",
    "pandas",
    "openpyxl",
    "xlsxwriter"
]

# Install each library
for lib in required_libs:
    subprocess.check_call([sys.executable, "-m", "pip", "install", lib])

print("\n✅ All required libraries have been installed successfully!")
