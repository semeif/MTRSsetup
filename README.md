# Microsoft Teams Room account setup
PowerShell script for creation and configuration of Microsoft Teams Room accounts 
Features:
-	Usage of Azure AD PowerShell Module instead of old MSOL Module
-	MFA Support
-	Logon to all 3 Modules (Azure AD, Teams, and Exchange)
-	Module and Admin Check if all PowerShell Modules are installed
-	Module Installation if they are missing
-	Check if MTR License is available and unassigned so it can be used
-	Setting correct Calendar Setting Auto Accept / Auto Answer
-	Waiting for RegistrarPool creation (can take some time) and Room assignment
New:
-	Room List Creation and Room assignment
-	Pick from available Room Lists and assign new meeting room


# First draft of German/English version available:
This version checks the Windows client language. If it is "de-DE" the script will run in German. 
All other language settings force the script to run in English.
If your Client language is not set to "de-DE" but you want to run the script in German 
you can change the $Language variable manually to "de-DE".
We also added the option to set the correct usage location for the MTR account.
You can find the new Version here:
https://github.com/semeif/TRSsetup/blob/main/MTR_Provisioning_En_De.ps1

# Create an MTR Account
Start the script without a Parameter to create a Microsoft teams room system account.

# check whether the needed modules are installed
start the sciript with the parameter modcheck, e.g. "TRSprovisioning.ps modcheck"
If the modules are not installed, the script tries to install the modules.
Keep in mind to start PS with admin privileges.




 MIT License

    Copyright (c) Microsoft Corporation.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
