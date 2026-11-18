
import platform
import win32com.client

print("Operating System:", platform.system())

try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    print("Outlook COM connection successful")
    print("Outlook Version:", outlook.Version)
except Exception as e:
    print("Outlook COM connection failed:", e)
