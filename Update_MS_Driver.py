import os
import platform
import requests
import re
import zipfile
import shutil


def get_edge_version():
    if platform.system() == "Windows":
        try:
            edge_install_path = os.path.join(os.environ["ProgramFiles(x86)"], "Microsoft\\Edge\\Application")
            # Get the list of directories in the Edge installation directory path given above
            edge_versions = [d for d in os.listdir(edge_install_path)
                             if re.match(r"\d+\.\d+\.\d+\.\d+", d)]
            # List the versions in descending order to get the latest version
            edge_versions.sort(key=lambda x: [int(i) for i in x.split('.')], reverse=True)
            if edge_versions:
                return edge_versions[0]  # Return the latest version currently installed
            else:
                print("No Edge installation found.")
                return None
        except Exception as e:
            print("Error accessing Edge version:", e)
            return None
    else:
        print("Unsupported OS. This script is for Windows only.")
        return None


def get_edge_webdriver_url(edge_version):
    url = "https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/?form=MA13LH"
    response = requests.get(url)

    if response.status_code == 200:
        # Search for the download link using regular expressions
        match = re.search(r'href="([^"]*edgedriver[^"]*)"', response.text)
        if match:
            return match.group(1)
        else:
            print("Edge WebDriver download link not found.")
            return None
    else:
        print("Failed to fetch Edge WebDriver download page.")
        return None


def download_webdriver(url, destination):
    try:
        with open(destination, "wb") as f:
            response = requests.get(url)
            f.write(response.content)
        print("WebDriver downloaded successfully:", destination)
    except Exception as e:
        print("Failed to download WebDriver:", e)


def extract_application_file(zip_file_path, destination_directory):
    # Extract application file from the ZIP file and save it to the destination directory
    try:
        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            application_file_name = "msedgedriver.exe"
            zip_ref.extract(application_file_name, destination_directory)
        print("Application file extracted successfully.")
    except Exception as e:
        print("Failed to extract application file:", e)


def main():
    edge_version = get_edge_version()
    if edge_version:
        print("Detected Microsoft Edge version:", edge_version)
        webdriver_url = get_edge_webdriver_url(edge_version)
        if webdriver_url:
            destination_directory = os.getcwd()  # Save to the script's folder
            webdriver_destination = os.path.join(destination_directory, "msedgedriver.zip")
            download_webdriver(webdriver_url, webdriver_destination)
            extract_application_file(webdriver_destination, destination_directory)

            # Specify the destination directories where you want to copy the file
            destination_paths = [
                r"C:\Users\Selenium_Project1",
                r"C:\Users\Selenium_Project2"
                #  ,r"HERE\another\path\here" # add  HERE  More Selenium Projects for MS Edge Driver update
            ]

            # Copy the file to each destination path and print the paths
            for destination_path in destination_paths:
                shutil.copy(os.path.join(destination_directory, "msedgedriver.exe"), destination_path)
                print("Copied msedgedriver.exe to:", destination_path)

            os.remove(webdriver_destination)  # Remove ZIP file after the .exe extraction
        else:
            print("Failed to get Edge WebDriver download URL.")
    else:
        print("Failed to detect Microsoft Edge version.")


if __name__ == "__main__":
    main()
