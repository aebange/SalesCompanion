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
from pyautogui import hotkey
# Used for managing file creation, deletion, and position
from shutil import copyfile, rmtree
from os import getcwd, path, remove, chdir, name
from glob import glob
# Used for editing the lead sheet once it's created
from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Pt
from winsound import PlaySound, SND_ALIAS
# Used for handling the windows clipboard
import win32clipboard
from time import sleep

originalDir = getcwd()
templateFile = 'Template.docx'
listenerState = True


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
                street_address = local_child.contents[0]
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


# Waits until a key is pressed, then checks it with the on_press function
def hotkey_listener():
    listener = Listener(
        on_press=on_press)
    listener.start()


# Fires a function when a ctrl+` is pressed
# TODO: Connect this to Q/A field collection
def on_press(key):
    global listenerState
    if str(key) == '<192>':
        # The hotkey has been hit
        hotkey("ctrl", "a")
        sleep(.3)
        hotkey("ctrl", "c")
        listenerState = False


# Clone the template sheet and rename it properly, delete the lead if the lead already exists
def lead_sheet_creation(local_download_dir, local_original_dir, local_template_file, local_account_name,
                        local_city_name, local_state_name):
    chdir(local_original_dir)
    file_name = 'MMC - MSH Lead - ' + local_account_name + ' - ' + local_city_name + ' - ' + local_state_name + '.docx'
    if path.isfile(file_name):
        # The file already exists, delete it so the new one may take its place
        remove(file_name)
    new_file_name = local_download_dir + '\\' + file_name
    copyfile(local_template_file, new_file_name)
    chdir(local_download_dir)
    return file_name


# Modify the contents of the lead sheet document to fill in the lead information
def lead_sheet_completion(local_target_file, local_content_list):
    local_doc = Document(local_target_file)
    style = local_doc.styles['Normal']
    font = style.font
    font.name = 'Trebuchet MS'
    font.size = Pt(11)
    par_format = local_doc.styles['Normal'].paragraph_format
    par_format.line_spacing = Pt(0)
    par_format.space_after = Pt(0)
    par_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    for local_par in local_doc.paragraphs:
        local_par.style = local_doc.styles['Normal']
        if 'CONTNAME' in local_par.text:
            local_par.text = "Name:			" + local_content_list[5]
        if 'PROPNAME' in local_par.text:
            local_par.text = "Property Name:	" + local_content_list[0]
        if 'CONTTITLE' in local_par.text:
            local_par.text = "Title:			" + local_content_list[8]
        if 'STREETCITY' in local_par.text:
            local_par.text = "Address:                  	" + local_content_list[1] + " " + \
                             local_content_list[2] + ", " + local_content_list[3] + " " + local_content_list[4]
        if 'PHONENUM' in local_par.text:
            local_par.text = "Phone: 		" + local_content_list[7]
        if 'CONTEMAIL' in local_par.text:
            local_par.text = "Email:                   	" + local_content_list[6]
        if 'MSHACCT' in local_par.text:
            local_par.text = "MSH Account:            " + local_content_list[10]
        if 'UNITCNT' in local_par.text:
            local_par.text = "Units:                      	" + local_content_list[11]
        if 'PRIMSUPPLY' in local_par.text:
            local_par.text = "Primary Supplier:     		" + local_content_list[12]
        if 'PMC' in local_par.text:
            local_par.text = "Property Mgmt:		" + local_content_list[13]
        if 'MMCACCT' in local_par.text:
            local_par.text = "Property Mgmt:		" + local_content_list[9]
    local_doc.save(local_target_file)


# Checks which of the html files is newest and passes it to be parsed
def select_html_file(local_download_dir):
    html_list = []
    chdir(local_download_dir)
    for file in glob("*.html"):
        html_list.append(file)
    # Parse through the list of .html files until the newest one is found
    while True:
        if len(html_list) > 1:
            if path.getctime(html_list[0]) > path.getctime(html_list[1]):
                html_list.remove(html_list[1])
            else:
                html_list.remove(html_list[0])
        else:
            break
    return html_list[0]


# Used to navigate to the current user's downloads folder in windows
def get_download_path():
    if name == 'nt':
        import winreg
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
            location = winreg.QueryValueEx(key, downloads_guid)[0]
        return location
    else:
        return path.join(path.expanduser('~'), 'downloads')


# Deletes the downloaded html files once parsing is complete to avoid cluttering
def clear_html_files(local_download_dir, local_target_file):
    chdir(local_download_dir)
    if path.isfile(local_target_file):
        # The file already exists, delete it so the new one may take its place
        remove(local_target_file)
    local_target_dir = local_target_file.split('.html')[0] + "_files"
    if path.isdir(local_target_dir):
        # The file already exists, delete it so the new one may take its place
        rmtree(local_target_dir)


# Checks for missing record content and replaces it with N/A
def nonetype_filter(local_content_list):
    for index, item in enumerate(local_content_list):
        if item is None:
            local_content_list[index] = "N/A"
    return local_content_list


# Gets the Q/A field information from the clipboard when ctrl+` is pressed
def get_qa_field():
    win32clipboard.OpenClipboard()
    global listenerState
    while listenerState:
        hotkey_listener()

    data = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()
    line_list = data.splitlines()
    completed_count = 0
    local_account_num, local_units_num, local_vendor, local_pmc = "N/A,", "N/A,", "N/A,", "N/A,"
    local_contents_list = [local_account_num, local_units_num, local_vendor, local_pmc]
    for index, line in enumerate(line_list):
        if line.find("Account Number") != -1:
            local_raw_acct_num = line_list[index + 1]
            local_contents_list[0] = local_raw_acct_num.split(' - o')[0]
            completed_count += 1
        if line.find("Units") != -1:
            local_raw_units_num = line_list[index + 1]
            local_contents_list[1] = local_raw_units_num.split(' - o')[0]
            completed_count += 1
        if line.find("Primary Vendor") != -1:
            local_raw_vendor = line_list[index + 1]
            local_contents_list[2] = local_raw_vendor.split(' - o')[0]
            completed_count += 1
        if line.find("Parent Company") != -1:
            local_raw_pmc = line_list[index + 1]
            local_contents_list[3] = local_raw_pmc.split(' - o')[0]
            completed_count += 1
    if completed_count >= 2:
        return local_contents_list
    else:
        # The clipboard failed to copy the items
        PlaySound('C:/Windows/Media/Hardware Fail.wav', SND_ALIAS)
        exit()


# TODO: Fix this mess and clear prints
downloadDir = get_download_path()

# Get the desired html file
targetHtmlFile = select_html_file(downloadDir)

# Open the targeted html file
targetFile = getcwd() + "\\" + targetHtmlFile
with open(targetFile, 'r') as f:
    contents = f.read()

# Create the soup object we will use to parse the html code
soup = BeautifulSoup(contents, features="html.parser")

# Scrape the values from the html file
contactName = get_target_name()
targetDiv = find_target_div(contactName)
acctName = account_name_parse(targetDiv)
contactPosition = position_parse(targetDiv)
contactPhoneNum = phone_number_parse(targetDiv)
contactEmail = email_parse(targetDiv)
mmcAcctNum = mmc_account_parse(targetDiv)
acctStreet, acctCity, acctState, acctZip = address_information_parse(targetDiv)

# Dump everything into a list
contentList = [acctName, acctStreet, acctCity, acctState, acctZip, contactName, contactEmail,
               contactPhoneNum, contactPosition, mmcAcctNum]

# Filter the html fields that were scraped
filteredContentList = nonetype_filter(contentList)

# Get the Q/A fields from the clipboard
mshAcctNum, acctUnits, acctPrimVendor, acctPMC = get_qa_field()
filteredContentList.append(mshAcctNum)
filteredContentList.append(acctUnits)
filteredContentList.append(acctPrimVendor)
filteredContentList.append(acctPMC)

# Generate a file for the new lead sheet
leadSheetFile = lead_sheet_creation(downloadDir, originalDir, templateFile, filteredContentList[0],
                                    filteredContentList[2], filteredContentList[3])

# Dump the values into the lead sheet
lead_sheet_completion(leadSheetFile, filteredContentList)

clear_html_files(downloadDir, targetFile)

# Close up shop
contents = f.close()
PlaySound('C:/Windows/Media/Speech On.wav', SND_ALIAS)
print("Exit Code 0")
