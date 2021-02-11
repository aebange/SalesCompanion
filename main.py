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
# Used for parsing the Html
from bs4 import BeautifulSoup
# Used for capturing hotkey presses
from pynput.keyboard import Listener
# Used for getting file directory info
from os import getcwd
# Used for copying the .docx template for editing
from shutil import copyfile
from os import rename


# Open the html file that is to be parsed by the program
# TODO: Automate the detection on this, manual entry obviously won't work
temporaryFile = 'John Doe _ Salesforce.html'
targetFile = getcwd() + "\\" + temporaryFile
with open(targetFile, 'r') as f:
    contents = f.read()

templateFile = 'Template.docx'

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


# Returns the contact's MMC account identifier as a string
def mmc_account_parse(local_target_div):
    for local_records in local_target_div.find_all("records-formula-output",
                                                   attrs={"data-output-element-id": "output-field",
                                                          "records-formulaoutput_formulaoutput-host": ""}):
        local_children = local_records.findChildren("lightning-formatted-text")
        for local_child in local_children:
            return local_child.contents[0]


# Provides the full address and independent address fields as separate string objects
def address_information_parse(local_target_div):
    for local_a in local_target_div("a", attrs={'target': "_blank", 'rel': 'noopener'}):
        local_children = local_a.findChildren("div", attrs={"class": "slds-truncate"})
        local_count = 0
        for local_child in local_children:
            if local_count == 0:
                street_address = local_child.contents
                local_count += 1
            elif local_count == 1:
                local_pair = local_child.contents[0]
                city_address_pair = local_pair.split(',')
                # Having issues with python throwing IndexErrors here, no clue why. Doing a workaround
                for item in city_address_pair:
                    city_address = item
                    break
                # Still having IndexErrors with range. Again ran out of patience and am working around with for loops
                index = 0
                for item2 in city_address_pair:
                    if index == 1:
                        state_zip_address = item2
                    else:
                        index += 1
                state_address, zip_address = (state_zip_address.split(' '))[1], (state_zip_address.split(' '))[2]
                return street_address, city_address, state_address, zip_address
            else:
                local_count += 1


# Fires a function when a ctrl+` is pressed
# TODO: Connect this to Q/A field collection
def on_press(key):
    if str(key) == '<192>':
        # The hotkey has been hit
        pass


def lead_sheet_creation(local_template_file, local_account_name, local_city_name, local_state_name):
    cwd = getcwd()
    new_file = copyfile(local_template_file, 'incomplete.docx')
    file_name = 'MMC - MSH Lead - ' + local_account_name + ' - ' + local_city_name + ' - ' + local_state_name + '.docx'
    newer_file = rename(new_file, file_name)
    print(newer_file)


# Waits until a key is pressed, then checks it with the on_press function
def hotkey_listener():
    # Collect events until a key is pressed
    with Listener(
            on_press=on_press) as listener:
        listener.join()


# TODO: Fix this mess and clear prints
contactName = get_target_name()
targetDiv = find_target_div(contactName)
acctName = account_name_parse(targetDiv)
contactPosition = position_parse(targetDiv)
contactPhoneNum = phone_number_parse(targetDiv)
contactEmail = email_parse(targetDiv)
mmcAcctNum = mmc_account_parse(targetDiv)
acctStreet, acctCity, acctState, acctZip = address_information_parse(targetDiv)
lead_sheet_creation(templateFile, acctName, acctCity, acctState)

contents = f.close()
print("Exit Code 0")
