import openpyxl

# Create a new Excel workbook
workbook = openpyxl.Workbook()

# Select the worksheet
worksheet = workbook.active
worksheet.title = "Automation Test Report"

# Define column headers
column_headers = ["Test Case ID", "Test Case Description", "Test Steps", "Expected Result", "Actual Result", "Status", "Execution Date", "Tester Name", "Comments", "Test Code"]
worksheet.append(column_headers)

# Define sample test data (including code snippets)
test_data = [
    ["TC001", "Login Test", "1. Open the login page\n2. Enter valid credentials\n3. Click 'Login'", "User is logged in successfully", "User logged in", "Pass", "2023-09-15", "John Doe", "", "```python\n# Python code snippet\nusername = 'user123'\npassword = 'pass123'\nlogin(username, password)\n```"],
    ["TC002", "Logout Test", "1. Click 'Logout' button", "User is logged out successfully", "User logged out", "Pass", "2023-09-15", "Jane Smith", "", "```java\n// Java code snippet\nlogout();\n```"]
]

# Populate the worksheet with test data
for row_data in test_data:
    worksheet.append(row_data)

# Save the workbook to a file
workbook.save("automation_test_report.xlsx")
