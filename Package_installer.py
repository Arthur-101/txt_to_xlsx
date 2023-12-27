### This Program will install all the Necessary Libraries/Packages required 
### to Run the `txt to xlsx`.
### Run this program before running the `Main` file.


import subprocess

all_packages = ['regex', 'pandas', 'numpy', 'openpyxl', 'tk', 'messagebox', 'matplotlib', 'Pillow']

def install_package(package_name):
    try:
        subprocess.check_call(["pip3", "install", package_name])
        print(f"Successfully installed {package_name}")
    except subprocess.CalledProcessError as e:
        print(f"Failed to install {package_name}. Error: {e}")

for package in all_packages:
    install_package(package)
