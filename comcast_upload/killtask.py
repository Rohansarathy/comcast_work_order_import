import subprocess

def kill_chrome_processes():
    try:
        subprocess.run(['taskkill', '/IM', 'chrome.exe', '/F'], check=True)
    except subprocess.CalledProcessError as e:
        print(f"chrome.exe not found or could not be killed: {e}")

    try:
        subprocess.run(['taskkill', '/IM', 'chromedriver.exe', '/F'], check=True)
    except subprocess.CalledProcessError as e:
        print(f"chromedriver.exe not found or could not be killed: {e}")


