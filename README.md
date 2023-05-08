# AD Proxy Comparator
This is my second tool that is used to compare proxy addresses in Local AD with all email addresses in Office 365. It's a nice and easy tool that can be used to avoid conflicts or missing addresses during a Azure AD Connect deployment.

**:exclamation: This is still in alpha, please use at own risk**

## How does this work?
It's nice and simple:
- You start by generating a JSON export of all user mailboxes' Email Addresses by using the command in the program. It'll tell you what to do!
- You then run the program on a computer with access to the domain, and it will use Powershell to pull in all user proxy addresses from a "search base" that you specify.
- The program will run, and will then spit out a text file with a full report of all addresses that were mismatched.

## Features
:white_check_mark: All commands are ran locally, and no outputs are sent to anyone  
:white_check_mark: The AD commands are "Get" commands only, so no changes made to AD by the program  
:white_check_mark: Search base filtering locks down the users to check by any OU or domain that you would like