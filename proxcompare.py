import subprocess
import os
import json
import datetime
import csv
import tkinter as tk
from tkinter import simpledialog
import sys

# Obtaining Timestamp
now = datetime.datetime.now()
timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")

# Obtaining Temporary Directory
temp_dir = os.environ.get('TEMP')

class obtainProxy:
    def __init__(self, ou):
        self.ou = ou

    def obtain(self):
        print("[INFO][PS] Obtaining Proxy addresses from OU: '"+self.ou+"'")
        pscommand = 'Get-ADUser -Filter * -SearchBase "' + self.ou + '" -Properties UserPrincipalName, ProxyAddresses | Select-Object UserPrincipalName, @{Name="ProxyAddresses";Expression={$_.ProxyAddresses -join ";"}} | ConvertTo-Json -Compress'
        command_output = subprocess.check_output(['powershell', '-Command', pscommand]).decode('utf-8')
        users = json.loads(command_output)
        return users

def genReport(data, type = "txt"):
    if type == "txt":
        txtname = "proxyreport_" + timestamp + ".txt"
        with open(txtname, 'w') as file:
            file.write(data)

print("You'll need to put the O365 JSON file in the same directory as this application")
print("1) Open Powershell")
print("2) Connect-ExchangeOnline")
print("3) Get-EXORecipient -ResultSize Unlimited -Filter {RecipientTypeDetails -eq 'UserMailbox'} | Select-Object 'PrimarySmtpAddress', 'EmailAddresses'| ConvertTo-Json | Out-File -Encoding utf8 -FilePath 'o365.json'")
print("4) Copy the o365.json file to the working directory of this application")
print("5) Once the file is present, press Enter")
input()

print(" ")

print("You are going to be prompted for a user base, this should be formatted like an LDAP search query")
print("For example, if your OU with all required users is:")
print("example.local > OU1 > OU2 > Users")
print("... your Search Base would be:")
print("OU=OU2,OU=OU1,DC=example,DC=local")

print(" ")

root = tk.Tk()
root.withdraw()
proxyinput = simpledialog.askstring(title="Search Base", prompt="Please enter search base\t\t\t")

if proxyinput is None:
    sys.exit(1)
else:
    proxysearchbase = proxyinput

startrun = obtainProxy(proxysearchbase)
userproxys = startrun.obtain()

reportdata = ""
adusermatch = []
adusermismatchcount = 0
o365usermatch = []
o365usermismatchcount = 0



with open('o365.json', 'r') as f:
    o365usersproxys = json.load(f)


reportdata += "Proxy Comparator" + "\n"
reportdata += timestamp + "\n"
reportdata += "Search Base: " + proxysearchbase + "\n"
reportdata += "_____________________________________________________________" + "\n"
reportdata += "\n"

for o365user in o365usersproxys:
    for aduser in userproxys:
        if o365user['PrimarySmtpAddress'] == aduser['UserPrincipalName']:
            o365usermatch.append(o365user['PrimarySmtpAddress'])

for aduser in userproxys:
    adproxys = aduser['ProxyAddresses'].split(";")

    for o365user in o365usersproxys:
        if o365user['PrimarySmtpAddress'] == aduser['UserPrincipalName']:
            adusermatch.append(aduser['UserPrincipalName'])
            print("[UPN MATCH]", o365user['PrimarySmtpAddress'])
            reportdata += "[UPN MATCH] " + aduser['UserPrincipalName'] + "\n"
            for adproxy in adproxys:
                if not adproxy in o365user["EmailAddresses"]:
                    reportdata += "[MISMATCH][AD] UPN: " + aduser['UserPrincipalName'] + " | Address: " + adproxy + "\n"
                    print("AD Proxy", adproxy, "not in O365")
            
            for o365address in o365user["EmailAddresses"]:
                if not o365address in adproxys:
                    reportdata += "[MISMATCH][O365] UPN: " + aduser['UserPrincipalName'] + " | Address: " + o365address + "\n"
                    print("O365 Proxy", o365address, "not in AD")
            
            reportdata += "========================\n"
    
    print(" ")

reportdata += "_____________________________________________________________" + "\n"
reportdata += "\n"

reportdata += "AD USERS NOT FOUND IN O365\n"
reportdata += "__________________________\n"

for aduser in userproxys:
    if not aduser['UserPrincipalName'] in adusermatch:
        reportdata += "[NO MATCH][AD] UPN: " + aduser['UserPrincipalName'] + "\n"
        print("[IN AD][NOT IN O365]", aduser['UserPrincipalName'])
        adusermismatchcount += 1
if adusermismatchcount == 0:
    reportdata += "No missing users" + "\n"

reportdata += "\n"
reportdata += "\n"
reportdata += "\n"
reportdata += "O365 USERS NOT FOUND IN AD\n"
reportdata += "__________________________\n"
print(" ")
for o365user in o365usersproxys:
    if not o365user['PrimarySmtpAddress'] in o365usermatch:
        reportdata += "[NO MATCH][AD] UPN: " + o365user['PrimarySmtpAddress'] + "\n"
        print("[IN O365][NOT IN AD]", o365user['PrimarySmtpAddress'])
        o365usermismatchcount += 1
if o365usermismatchcount == 0:
    reportdata += "No missing users" + "\n"


print(" ")
print("[COMPLETE] A report of this data will be generated in the working directory")
genReport(reportdata)