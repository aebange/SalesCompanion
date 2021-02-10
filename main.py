#######################################################################
# Filename:  main.py
# Author:    Alex Bange
#######################################################################
# Purpose:
#   To be used to parse Salesforce .html files downloaded via browser
#
# Usage
#   DOS> python main.py
#
# Assumptions
#   A) python is in the PATH
#
# Dependencies
#   A) BeautifulSoup
#   B) Pynput
#######################################################################

# Import dependent libraries
from bs4 import BeautifulSoup
from pynput.keyboard import Listener
from os import getcwd

# Open the html file that is to be parsed by the program
# TODO: Automate the detection on this, manual entry obviously won't work
temporary_file = 'John Doe _ Salesforce.html'
target_file_directory = getcwd() + "\\" + temporary_file
with open(target_file_directory, 'r') as f:
    contents = f.read()

# Create the soup object we will use to parse
soup = BeautifulSoup(contents, features="html.parser")


# Gets the contact name from the top of the document so we can verify what the current record is. Returns string
def get_target_name():
    titles = soup.find_all("title")
    for title in titles:
        target = title.contents[0]
        target_name = target.split('|')
        target_name = target_name[0]
        target_name = target_name[:-1]
        return target_name


# Gets the parent div from the section of the html that only pertains to the current record by checking the
# contents of the children against the name listed at the top of the webpage's html (which is always the current record)
def find_target_div(local_target_name):
    local_divs = soup.find_all("div", attrs={"record_flexipage-recordpagedecorator_recordpagedecorator": "",
                                             "class": "record-page-decorator"})
    for local_div in local_divs:
        try:
            local_children = local_div.findChildren("a", attrs={"data-aura-class": "forceOutputLookup"})
            for local_child in local_children:
                if local_child.contents[0] == local_target_name:
                    return local_div
        # TODO: Figure out what the reason was that I added this exception clause and why it is bare
        except:
            pass


# Returns the account name from the Salesforce record as a string
def account_name_parse(local_target_div):
    for local_spans in local_target_div.find_all("span", attrs={"class": "custom-truncate uiOutputText",
                                                                "data-aura-class": "uiOutputText"}):
        for local_span in local_spans:
            return local_span


# Returns the record's phone number. Unclear which number I am grabbing here, check later
def phone_number_parse(local_target_div):
    for local_spans in local_target_div.find_all("span", attrs={'dir': 'ltr'}):
        for local_span in local_spans:
            return local_span


# Returns the contact's email from the Salesforce record as a string
def email_parse(local_target_div):
    for local_as in local_target_div.find_all("a", attrs={"class": "emailuiFormattedEmail",
                                                          "data-aura-class": "emailuiFormattedEmail"}):
        for local_a in local_as:
            return local_a


# Returns the contact's position at the property from the Salesforce record as a string
def position_parse(local_target_div):
    for local_divs in local_target_div.find_all("div", attrs={"class": "slds-item--detail slds-truncate recordCell"}):
        local_children = local_divs.findChildren("span",
                                                 attrs={"class": "uiOutputText", "data-aura-class": "uiOutputText"})
        for local_child in local_children:
            if local_child.contents[0]:
                return local_child.contents[0]


# Provides the full address and independent address fields as separate string objects
# TODO: Update this to work with the new target div method, get rid of prints
def address_information_parse():
    for local_a in soup.find_all("a", attrs={'target': "_blank", 'rel': 'noopener'}):
        # full_address = local_a.get('title')
        local_children = local_a.findChildren("div", attrs={"class": "slds-truncate"})
        local_count = 0
        for local_child in local_children:
            if local_count == 0:
                # street_address = local_child.contents
                local_count += 1
            elif local_count == 1:
                local_pair = local_child.contents[0]
                city_address = local_pair.split(',')[0]
                state_zip_address = local_pair.split(',')[1]
                # state_address = state_zip_address.split(' ')[1]
                state_address, zip_address = state_zip_address.split(' ')[1], state_zip_address.split(' ')[2]
                print("City is {}".format(city_address))
                print("State is {}".format(state_address))
                print("Zip is {}".format(zip_address))


# Fires a function when a ctrl+` is pressed
# TODO: Connect this to Q/A field collection
def on_press(key):
    if str(key) == '<192>':
        # The hotkey has been hit
        pass


# Collect events until a key is pressed
with Listener(
        on_press=on_press) as listener:
    listener.join()


# TODO: Fix this mess and clear prints
targetName = get_target_name()
print(targetName)
targetDiv = find_target_div(targetName)
print(account_name_parse(targetDiv))
print(phone_number_parse(targetDiv))
print(email_parse(targetDiv))
print(position_parse(targetDiv))
# address_information_parse()


contents = f.close()
print("Exit Code 0")
