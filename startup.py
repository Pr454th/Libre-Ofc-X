import subprocess
import time
import os
import signal

# START LibreOffice headless
def start_libreoffice():
    cmd = [
        "libreoffice",
        "--headless",
        '--accept=socket,host=localhost,port=2002;urp;',
        "--nologo",
        "--nofirststartwizard"
    ]
    subprocess.Popen(cmd)
    print("LibreOffice headless started.")
    time.sleep(2)  # Give it time to boot

# CHECK if LibreOffice is running
def is_libreoffice_running():
    result = subprocess.run(
        ["pgrep", "-f", "soffice.bin"],
        stdout=subprocess.PIPE
    )
    return result.returncode == 0

# STOP LibreOffice
def stop_libreoffice():
    subprocess.run(["pkill", "-f", "soffice.bin"])
    print("LibreOffice stopped.")

# Example usage
if __name__ == "__main__":
    start_libreoffice()
    if is_libreoffice_running():
        print("LibreOffice is running.")
    else:
        print("LibreOffice failed to start.")

    input("Press Enter to stop LibreOffice...")
    stop_libreoffice()
