import win32com.client as win32
import time
import os
from win32com.client import gencache
#import pythoncom

# Initialize CANoe Application
CANoe = win32.DispatchEx("CANoe.Application")
time.sleep(10)
CANoe.Open(r"C:\Users\Dell\Downloads\Jenkins\Jenkins\ACC_RBS\pyrbs.cfg")
time.sleep(2)

# Test configuration details
testConfigName = "AutomatedTests_VT_03"
testUnitsFolder = r"C:\Users\Dell\Downloads\Jenkins\Jenkins\vTEST_Project\ACC_Testing"  # Folder containing .vtuexe files

# Ensure CANoe module
canoeModul = gencache.EnsureModule('{7F31DEB0-5BCC-11D3-8562-00105A3E017B}', 0x0, 1, 54)

# Find all .vtuexe files in the specified folder
testUnitPaths = [
    os.path.join(testUnitsFolder, file)
    for file in os.listdir(testUnitsFolder)
    if file.endswith(".vtuexe")
]



# Add test units to the test configuration
print("Trying to add test cases to CANoe...")
# testConfigurations = CANoe.Configuration.TestConfigurations
# testConfiguration = testConfigurations.Add()
# testConfiguration.Name = testConfigName
# testUnits = canoeModul.ITestUnits2(testConfiguration.TestUnits)
testConfiguration = CANoe.Configuration.TestConfigurations.Add()
testConfiguration.Name = testConfigName
testUnits = canoeModul.ITestUnits2(testConfiguration.TestUnits)

# Add each test unit from the folder
for testUnitPath in testUnitPaths:
    print(f"Adding TestUnit: {testUnitPath}")
    testUnits.Add(testUnitPath)

# # Use a loop to monitor the test state
# max_wait_cycles = 1000
# wait_cycles = 0
#
# while testConfigurations.Running and wait_cycles < max_wait_cycles:
#     if testConfiguration.running == "Finished":  # Hypothetical example
#         print("Test execution completed.")
#         break
#     # Optional: Clear CANoe UI logs if necessary
#     if len(CANoe.UI.Write.Text) > 3:
#         CANoe.UI.Write.Clear()
#
#     # Increment wait cycle and sleep
#     wait_cycles += 1
#     time.sleep(0.1)
#
# if wait_cycles >= max_wait_cycles:
#     raise TimeoutError("Test execution timed out.")

# Start the measurement
time.sleep(1)
CANoe.Measurement.Start()
print("Measurement started.")
time.sleep(2)

# Start test configuration
testConfiguration.Start()
print("Test configuration started running...!")
time.sleep(10)
print("Test execution is completed..!")
time.sleep(5)

# Stop the measurement
if CANoe.Measurement.Running:
    CANoe.Measurement.Stop()
    print("Measurement stopped.")
else:
    print("Measurement was already stopped.")

print("Opening the vtest report..!")
time.sleep(1)

# Open the test report file
report_file_path = os.path.join(r"C:\Users\Dell\Downloads\Jenkins\Jenkins\ACC_RBS\Report_AutomatedTests_VT_02.vtestreport")
if os.path.exists(report_file_path):
    print(f"Opening report file: {report_file_path}")
    os.startfile(report_file_path)
else:
    print(f"Report file not found: {report_file_path}")