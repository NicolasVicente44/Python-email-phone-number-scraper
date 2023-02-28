import pyperclip
import re


# Get clipboard content
clipboard = pyperclip.paste()



# Define regular expression pattern
phoneRegex = re.compile(r'''(
    (\d{3}|\(\d{3}\))? # area code
    (\s|-|\.)? # separator
    (\d{3}) # first 3 digits
    (\s|-|\.) # separator
    (\d{4}) # last 4 digits
    (\s*(ext|x|ext.)\s*(\d{2,5}))? # extension
    )''', re.VERBOSE)

emailRegex = re.compile(r'''(
    [a-zA-Z0-9._%+-]+ # username
    @ # @ symbol
    [a-zA-Z0-9.-]+ # domain name
    (\.[a-zA-Z]{2,4}) # dot-something
    )''', re.VERBOSE)



# Search for phone numbers and email addresses
phoneNumbers = []
emailAddresses = []

for match in phoneRegex.findall(clipboard):
    phoneNumbers.append(match[0])
    
for match in emailRegex.findall(clipboard):
    emailAddresses.append(match[0])




# Copy extracted data to clipboard
output = "Phone Numbers:\n" + "\n".join(phoneNumbers) + "\n\nEmail Addresses:\n" + "\n".join(emailAddresses)
pyperclip.copy(output)




#give feedback based on if phone numbers or emails are found in the selection
if len(phoneNumbers) == 0:
    print("No phone number found in that selection")
else:
    print(f"{len(phoneNumbers)} phone numbers found in that selection")


if len(emailAddresses) == 0:
    print("No email addresses found in that selection")
else:
    print(f"{len(emailAddresses)} email addresses found in that selection")