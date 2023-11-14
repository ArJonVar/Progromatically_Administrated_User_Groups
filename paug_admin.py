#region imports
import smartsheet
from smartsheet.exceptions import ApiError
from smartsheet_grid import grid
import requests
import json
import time
from globals import smartsheet_token, m365_pw, bamb_token, bamb2_token
from logger import ghetto_logger
import os
import subprocess
import re
from PyBambooHR import PyBambooHR
import pandas as pd
import re
#endregion

class PowershellDDLAdmin():
    '''Explain Class'''
    def __init__(self, config):
        self.config = config
        self.path = config.get('dev_path')
        self.smartsheet_token=config.get('smartsheet_token')
        self.bamb_token=config.get('bamb_token')
        self.pw=config.get('m365_pw')
        self.b2token=self.config.get('b2token')
        grid.token=smartsheet_token
        self.sheet_id = 5588740365832068
        self.smart = smartsheet.Smartsheet(access_token=self.smartsheet_token)
        self.smart.errors_as_exceptions(True)
        self.start_time = time.time()
        self.log=ghetto_logger("powershell_ddl_admin.py")
        self.login_command=f'''$secpasswd = ConvertTo-SecureString '{self.pw}' -AsPlainText -Force
            $o365cred = New-Object System.Management.Automation.PSCredential ("mslive@schuchartdow.onmicrosoft.com", $secpasswd)
            Connect-ExchangeOnline -Credential $o365cred'''

#region grab data
    #region get Microsoft data
        #region Distribution Lists
    def handle_distlist_data(self):
        '''gets distribution list data'''
        self.pwrshl_get_distlist()
        self.import_distlist()
        self.pwrshl_get_distlist_members()
        self.extract_data_from_file(self.path+'\_distlist_members.txt')
    def pwrshl_get_distlist(self):
        '''creates a string output of data regarding 'distribution list groups', I added 'get-dynamicdistributiongroup because if I dont, the distribution list command turns out as a list and not a larger dict-ish file'''
        commands = f'''{self.login_command}
            Get-DynamicDistributionGroup
            Write-Output "##1##"
            Get-DistributionGroup'''

        # Open the file to save the output
        with open("_pwrshl_output_dl.txt", "w") as file:

            # Execute the commands
            p = subprocess.run(
                ["powershell", "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-Command", commands],
                capture_output=True,
                text=True
            )

            # Print the output
            file.write(f'--- Output: ---\n{p.stdout}\n\n')
            file.write(f'--- Error: ---\n{p.stderr}\n\n')
    def import_distlist(self):
        '''reads the pwrshell output and produces a list of 'distribution list groups' '''

        self.distlist_dict = []
        dict_buffer = {}
        last_key = None  # This will hold the last key that was inserted into the dictionary        

        start_processing = False  # This flag will tell us when to start processing lines       

        with open('_pwrshl_output_dl.txt', 'r') as f:
            for line in f:
                line = line.strip()     

                if '##1##' in line:
                    start_processing = True
                    continue  # Skip the rest of the loop for this line     

                if not start_processing:
                    continue  # If we haven't reached ##1## yet, skip this line     

                if 'GroupType' in line:
                    # If we're not on the first RunspaceId, append the previous dictionary to the result list
                    if dict_buffer:
                        self.distlist_dict.append(dict_buffer)
                        dict_buffer = {}        

                try:
                    key, value = line.split(':', 1)
                    key = key.strip()
                    dict_buffer[key] = value.strip()
                    last_key = key
                except ValueError:
                    # If we can't split, it's either an empty line or a continuation of the previous value
                    if line.strip() and last_key is not None:
                        # Non-empty line: continuation of previous value
                        dict_buffer[last_key] += line.strip()       

        # Don't forget to add the last dictionary if the file doesn't end with RunspaceId
        if dict_buffer:
            self.distlist_dict.append(dict_buffer)       
    def pwrshl_get_distlist_members(self):
        '''publish a list of email address members of each dist list in the form of a .txt output file to be processed later'''
        # Assuming result is pre-defined or fetched earlier in the script
        self.dl_data = [{'name': item_dict.get('Name'), 'email': item_dict.get('WindowsEmailAddress'), 'cmd_id': 2 + i} for i, item_dict in enumerate(self.distlist_dict)]

        # Create a complete command sequence
        complete_command = self.login_command
        for i, item in enumerate(self.dl_data):
            individual_command = f'''
        Write-Output "##{2 + i}## {item['name']}"
        Get-DistributionGroupMember -Identity "{item['name']}" | Select-Object PrimarySmtpAddress | Out-String -Width 4096
        '''
            complete_command += individual_command

        # Write commands to a temporary PS1 file
        script_file = "temp_script.ps1"
        with open(script_file, 'w', encoding='utf-8') as f:
            f.write(complete_command)

        # Execute the PS1 script and capture its output
        completed_process = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script_file], text=True, capture_output=True)

        # Optionally, remove the temp files
        os.remove(script_file)

        # Write the output and commands to the distlist_members.txt file
        with open("_distlist_members.txt", "w", encoding='utf-8') as file:
            file.write(f'--- Output: ---\n{completed_process.stdout}') 
    def extract_data_from_file(self, file_path):
        # Read the file
        with open(file_path, 'r') as file:
            text = file.read()

        # Regular expression pattern to extract data
        pattern = r'##(\d+)##\s*([\w@.-]+)(?:\n\nPrimarySmtpAddress\s*-+\s*((?:\w+@\w+\.\w+(?:\s|\n)*)+))?'

        # Extract data using the regular expression
        self.matches = re.findall(pattern, text)

        # extracted_data = []
        # for match in matches:
        #     extracted_data.append({
        #         'id': int(match[0]),
        #         'name_or_email': match[1],
        #         'emails': match[2].split() if match[2] else []
        #     })

        # return extracted_data

        # Format the extracted data
        for match, data in zip(self.matches, self.dl_data):
            data['members']= match[2].split() if match[2] else []
        #endregion
        #region Mail Contact
    def handle_mailcontact_data(self):
        '''handles grabbing mail contacts'''
        self.pwrshl_get_contact()
        self.extract_contact_list()
        self.pwrshl_get_contactlist_members()
        self.extract_contactobject()
    def pwrshl_get_contact(self):
        '''creates a string output of data regarding contact groups'''
        commands = f'''{self.login_command}
            Write-Output "##1##"
            Get-MailContact'''

        # Open the file to save the output
        with open("_pwrshl_output_mc.txt", "w") as file:

            # Execute the commands
            p = subprocess.run(
                ["powershell", "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-Command", commands],
                capture_output=True,
                text=True
            )

            # Print the output
            file.write(f'--- Output: ---\n{p.stdout}\n\n')
            file.write(f'--- Error: ---\n{p.stderr}\n\n')
    def extract_contact_list(self):
        '''reads the pwrshell output and produces a list of 'Mail Contact' '''
        with open('_pwrshl_output_mc.txt', 'r') as file:
            content = file.read()

            # Find the relevant part of the text (after ##1##)
            relevant_part = re.search(r'##1##(.+)', content, re.DOTALL)
            if relevant_part:
                relevant_content = relevant_part.group(1)
            else:
                relevant_content = ""

            # Regular expression to extract names
            pattern = r'\n([^\n]+?)\s{2,}'
            names = re.findall(pattern, relevant_content)
            # Process names (optional, based on your needs)
            self.contact_list = [name.strip() for name in names]
    def pwrshl_get_contactlist_members(self):
        '''publish a list of email address members of each dist list in the form of a .txt output file to be processed later'''
        # Create a complete command sequence
        complete_command = self.login_command
        for i, contact in enumerate(self.contact_list):
            individual_command = f'''
        Write-Output "##{2 + i}## {contact}"
        Get-MailContact -Identity "{contact}" | Format-List
        '''
            complete_command += individual_command

        # Write commands to a temporary PS1 file
        script_file = "temp_script.ps1"
        with open(script_file, 'w', encoding='utf-8') as f:
            f.write(complete_command)

        # Execute the PS1 script and capture its output
        completed_process = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script_file], text=True, capture_output=True)

        # Optionally, remove the temp files
        os.remove(script_file)

        # Write the output and commands to the distlist_members.txt file
        with open("pwrshl_output_contact_details.txt", "a", encoding='utf-8') as file:
            file.write(f'--- Output: ---\n{completed_process.stdout}') 
    def extract_contactobject(self):
        '''reads the pwrshell output and produces a list of 'distribution list groups' '''

        contact_objects = []
        contact_dict = {}

        with open('pwrshl_output_contact_details.txt', 'r') as f:
            for line in f:
                line = line.strip()

                # Check if the line starts with '##number##'
                match = re.match(r'##(\d+)##\s*(.*)', line)
                if match:
                    # If we already have data in contact_dict, add it to the list
                    if contact_dict:
                        contact_objects.append(contact_dict)
                    contact_dict = {'id': match.group(1), 'name': match.group(2)}
                elif line and contact_dict:
                    parts = line.split(':', 1)
                    if len(parts) == 2:
                        key, value = parts
                        contact_dict[key.strip()] = value.strip()

        # Add the last contact_dict if it's not empty
        if contact_dict:
            contact_objects.append(contact_dict)
        self.contact_data = []
        contact_data_raw = [{'name': item_dict.get('name'), 'email': item_dict.get('PrimarySmtpAddress'), 'cmd_id': item_dict.get('id')} for i, item_dict in enumerate(contact_objects)]  
        for contact in contact_data_raw:
            if contact not in self.contact_data:
                self.contact_data.append(contact)
        #endregion
    #endregion
    #region get bamboohr data
    def get_bamboohr_data(self):
        '''pulls out the employee directory into a df'''
        bamboo = PyBambooHR.PyBambooHR(subdomain='dowbuilt', api_key=self.bamb_token)

        dir = bamboo.get_employee_directory()
        self.hr_df = pd.DataFrame(dir)
        self.add_position_category()
    def position_category_api_call(self, id):
        '''api call to grab position category, needs seperate api key (automation@dowbuilt.com), uses a dif api token as it needs to be base-64 encoded first to go this way'''

        url = f"https://api.bamboohr.com/api/gateway.php/Dowbuilt/v1/employees/{id}/?fields=customPositionCategory&onlyCurrent=true"

        headers = {
            "accept": "application/json",
            "authorization": f"Basic {self.config.get('b2token')}"
        }       

        response = requests.get(url, headers=headers)

        return json.loads(response.text).get('customPositionCategory')
    def add_position_category(self):
        '''the df comes with default options, we need non default option of position category'''
        pos_cat = [
            self.position_category_api_call(id)
            for id in self.hr_df["id"]
        ]

        self.hr_df["position_category"] = pos_cat
        self.log.log(f"position categories imported from Bamboo API")
    #endregion
    #region get smartsheet data
    def grab_smartsheet_data(self):
        '''grab sheets form Progromatically Administrated User Groups workspace'''
        user_group_def_df=grid(4014864070561668)
        user_group_def_df.fetch_content()
        user_group_exception_df=grid(4146543875542916)
        user_group_exception_df.fetch_content()


    #endregion
    #region post to microsoft
    def pwrshl_post_mailcontact(self, new_contact_list):
        # Define commands as a single string with multiple commands on one line
        commands = f'{self.login_command}'
        for i, new_contact in enumerate(new_contact_list):
            commands += f'''
        Write-Output "##{1+i}##"
        New-MailContact -Name {new_contact.get('name')} -ExternalEmailAddress {new_contact.get('email')}'''         

        # Open the file to save the output
        with open("_new_contacts.txt", "w") as file:
            # Print the commands
            file.write(f'--- Commands: ---\n{commands}\n\n')        

            # Execute the commands
            p = subprocess.run(
                ["powershell", "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-File", commands],
                capture_output=True,
                text=True
            )       

            # Print the output
            file.write(f'--- Output: ---\n{p.stdout}\n\n')
            file.write(f'--- Error: ---\n{p.stderr}\n\n')
    def pwrshl_add_remv_dl_member(self, dl_change_list):
        # Define commands as a single string with multiple commands on one line
        commands = f'{self.login_command}'
        for i, change in enumerate(dl_change_list):
            commands += f'''
        Write-Output "##{1+i}##"
        {change.get('action')}-DistributionGroupMember -Identity {member.get('dl')} -Member {member.get('email')}'''         

        # Open the file to save the output
        with open("_new_contacts.txt", "w") as file:
            # Print the commands
            file.write(f'--- Commands: ---\n{commands}\n\n')        

            # Execute the commands
            p = subprocess.run(
                ["powershell", "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-File", commands],
                capture_output=True,
                text=True
            )       

            # Print the output
            file.write(f'--- Output: ---\n{p.stdout}\n\n')
            file.write(f'--- Error: ---\n{p.stderr}\n\n')           
    #endregion
#endregion


    def run(self):
        '''runs main script as intended'''
        self.handle_distlist_data()
        self.handle_mailcontact_data()
        self.get_bamboohr_data()
        self.grab_smartsheet_data()

if __name__ == "__main__":
    config = {
        'smartsheet_token':smartsheet_token,
        'm365_pw': m365_pw,
        'bamb_token': bamb_token,
        'b2token': bamb2_token,
        'dev_path': r'C:\Egnyte\Shared\IT\Python\Ariel\Dynamic_distribution_lists\V1'
    }
    pda = PowershellDDLAdmin(config)
    pda.run()