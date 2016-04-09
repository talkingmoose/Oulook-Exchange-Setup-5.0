(*

--------------------------------------------
Outlook Exchange Setup 5
© Copyright 2008-2016 William Smith
bill@officeformachelp.com

Except where otherwise noted, this work is licensed under
http://creativecommons.org/licenses/by/4.0/

This file is one of four files for assisting a user with configuring
an Exchange account in Microsoft Outlook 2016 for Mac:

1. Outlook Exchange Setup 5.2.1.scpt
2. OutlookExchangeSetupLaunchAgent.sh
3. net.talkingmoose.OutlookExchangeSetupLaunchAgent.plist
4. com.microsoft.Outlook.plist for creating a configuraiton profile

These scripts and files may be freely modified for personal or commercial
purposes but may not be republished for profit without prior consent.

If you find these resources useful or have ideas for improving them,
please let me know. It is only compatible with Outlook 2016 for Mac.

--------------------------------------------

This script assists a user with the setup of his Exchange account
information. Below are basic instructions for using the script.
Consult the Outlook Exchange Setup 5 Administrator's Guide
for complete details.

1.	Customize the "network and  server properties" below with information
	appropriate to your network.
	
2.	Deploy this script to a location on your Macs such as
	"/Library/CompanyName/OutlookExchangeSetup5.2.1.scpt".

3. 	Deploy the recommended "Outlook preferences.mobileconfig"
	configuration profile to eliminate Outlook's startup windows.
	This assumes you're using the volume license edition
	of Office 2016 for Mac.
	
4.	Deploy the OutlookExchangeSetup5.plist file to
	/Library/LaunchAgents. Update the path to point to the
	OutlookExchangeSetup5.2.1.scpt script.
	  
This script assumes the user's full name is in the form of "Last, First",
but is easily modified if the full name is in the form of "First Last".
It works especially well if the Mac is bound to Active Directory where
the user's short name will match his login name. Optionally, you cans set dscl
to pull the user's EMailAddress from a directory service.

*)

-- global logMesage

--------------------------------------------
-- Begin network, server and preferences
--------------------------------------------


--------------- Exchange Server settings ----------------------

property useKerberos : true
-- Set this to true only if Macs in your environment are bound
-- to Active Directory and your network is properly configured.

property ExchangeServer : "exchange.example.com"
-- Address of your organization's Exchange server.

property ExchangeServerRequiresSSL : true
-- True for most servers.

property ExchangeServerSSLPort : 443
-- If ExchangeServerRequiresSSL is true set the port to 443.
-- If ExchangeServerRequiresSSL is false set the port to 80.
-- Use a different port number only if your administrator instructs you.

property DirectoryServer : "gc.example.com"
-- Address of an internal Global Catalog server (a type of Windows domain controller).
-- The LDAP server in a Windows network will be a Global Catalog server,
-- which is separate from the Exchange Server.

property DirectoryServerRequiresAuthentication : true
-- This will almost always be true.

property DirectoryServerRequiresSSL : true
-- This will almost always be true.

property DirectoryServerSSLPort : 3269
-- If DirectoryServerRequiresSSL is true set the port to 3269.
-- If DirectoryServerRequiresSSL is false set the port to 3268.
-- Use a different port number only if your Exchange administrator instructs you.

property DirectoryServerMaximumResults : 1000
-- When searching the Global Catalog server, this number determines
-- the maximum number of entries to display.

property DirectoryServerSearchBase : ""
-- example: "cn=users,dc=domain,dc=com"
-- Usually, this is empty.


--------------- For Active Directory users ---------------------

property getUserInformationFromActiveDirectory : true
-- If Macs are bound to Active Directory they can probably use
-- dscl to return the current user's email address, phone number, title, etc.
-- Use Active Directory when possible, otherwise complete the next section.


--------------- For non Active Directory users ---------------

property domainName : "example.com"
-- Complete this only if not using Active Directory to get user information.
-- The part of your organization's email address following the @ symbol.

property emailFormat : 1
-- Complete this only if not using Active Directory to get user information.
-- When Active Directory is unavailable to determine a user's email address,
-- this script will attempt to parse it from the display name of the user's login.

-- Describe your organization's email format:
-- 1: Email format is first.last@domain.com
-- 2: Email format is first@domain.com
-- 3: Email format is flast@domain.com (first name initial plus last name)
-- 4: Email format is shortName@domain.com

property displayName : 2
-- Complete this only if not using Active Directory to get user information.
-- Describe how the user's display name appears at the bottom of the menu
-- when clicking the Apple menu (Log Out Joe Cool... or Log Out Cool, Joe...).
-- 1: Display name appears as "Last, First"
-- 2: Display name appears as "First Last"

property domainPrefix : ""
-- Optionally append a NetBIOS domain name to the beginning of the user's short name.
-- Be sure to use two backslashes when adding a name.
-- Example: Use "TALKINGMOOSE\\" to set user name "TALKINGMOOSE\username".


--------------- User Experience -------------------------------

property verifyEMailAddress : false
-- If set to "true", a dialog asks the user to confirm his email address.

property verifyServerAddress : false
-- If set to "true", a dialog asks the user to confirm his Exchange server address.

property displayDomainPrefix : false
-- If set to "true", the username appears as "DOMAIN\username".
-- Otherwise, the username appears as "username".

property downloadHeadersOnly : false
-- If set to "true", only email headers are downloaded into Outlook.
-- This takes much less time to sync but a user must be online
-- to download and view messages.

property hideOnMyComputerFolders : false
-- If set to "true", hides local folders.
-- A single Exchange account should do this by default.

property unifiedInbox : false
-- If set to "true", turns on the Group Similar Folders feature
-- in Outlook menu > Preferences > General.

property disableAutodiscover : false
-- If set to "true", disables Autodiscover functionality
-- for the Exchange account. Not recommended for mobile devices
-- that may connect to an internal Exchange server address and
-- connect to a different external Exchange server address.

property errorMessage : "Outlook's setup for your Exchange account failed. Please contact the Help Desk for assistance."
-- Customize this error message for your users in case their account setup fails

--------------------------------------------
-- End network, server and preferences
--------------------------------------------

--------------------------------------------
-- Begin log file setup
--------------------------------------------

-- create the log file in the current user's Logs folder

writeLog("Starting Exchange account setup...")
writeLog("Script: " & name of me)
writeLog(return)

--------------------------------------------
-- End log file setup 
--------------------------------------------

--------------------------------------------
-- Begin logging script properties
--------------------------------------------

writeLog("Setup properties...")
writeLog("Use Kerberos: " & useKerberos)
writeLog("Exchange Server: " & ExchangeServer)
writeLog("Exchange Server Requires SSL: " & ExchangeServerRequiresSSL)
writeLog("Exchange Server Port: " & ExchangeServerSSLPort)
writeLog("Directory Server: " & DirectoryServer)
writeLog("Directory Server Requires Authentication: " & DirectoryServerRequiresAuthentication)
writeLog("Directory Server Requires SSL: " & DirectoryServerRequiresSSL)
writeLog("Directory Server SSL Port: " & DirectoryServerSSLPort)
writeLog("Directory Server Maximum Results: " & DirectoryServerMaximumResults)
writeLog("Directory Server Search Base: " & DirectoryServerSearchBase)
writeLog("Get User Information from Active Directory: " & getUserInformationFromActiveDirectory)
writeLog(return)

if getUserInformationFromActiveDirectory is false then
	writeLog("Domain Name: " & domainName)
	writeLog("Email format: " & emailFormat)
	writeLog("Display Name: " & displayName)
	writeLog("Domain Prefix: " & domainPrefix)
	writeLog(return)
end if

writeLog("Verify Email Address: " & verifyEMailAddress)
writeLog("Verify Server Address: " & verifyServerAddress)
writeLog("Display Domain Prefix: " & displayDomainPrefix)
writeLog("Download Headers Only: " & downloadHeadersOnly)
writeLog("Hide On My Computer Folders: " & hideOnMyComputerFolders)
writeLog("Unified Inbox: " & unifiedInbox)
writeLog("Disable Autodiscover: " & disableAutodiscover)
writeLog("Error Message text: " & errorMessage)
writeLog(return)

--------------------------------------------
-- End logging script properties 
--------------------------------------------

--------------------------------------------
-- Begin collecting user information
--------------------------------------------

-- attempt to read information from Active Directory for the Me Contact record

set userFirstName to ""
set userLastName to ""
set userDepartment to ""
set userOffice to ""
set userCompany to ""
set userWorkPhone to ""
set userMobile to ""
set userFax to ""
set userTitle to ""
set userStreet to ""
set userCity to ""
set userState to ""
set userPostalCode to ""
set userCountry to ""
set userWebPage to ""

if getUserInformationFromActiveDirectory is true then
	
	-- Get information from Active Directoy
	
	-- get the domain's primary NetBIOS domain name
	
	try
		set netbiosDomain to do shell script "/usr/bin/dscl \"/Active Directory/\" -read / SubNodes | awk 'BEGIN {FS=\": \"} {print $2}'"
		if displayDomainPrefix is true then
			set domainPrefix to netbiosDomain & "\\"
		else
			set domainPrefix to ""
		end if
	on error
		
		-- something went wrong
		
		display dialog errorMessage & return & return & "Unable to determine NETBIOS domain name. This computer may not be bound to Active Directory." with icon stop buttons {"OK"} default button {"OK"} with title "Outlook Exchange Setup"
		error number -128
	end try
	
	-- read full user information from Active Directory
	
	try
		set AppleScript's text item delimiters to {": "}
		set userInformation to do shell script "/usr/bin/dscl \"/Active Directory/" & netbiosDomain & "/All Domains/\" -read /Users/$USER AuthenticationAuthority City co company department physicalDeliveryOfficeName sAMAccountName wWWHomePage EMailAddress FAXNumber FirstName JobTitle LastName MobileNumber PhoneNumber PostalCode RealName State Street"
	on error
		
		-- something went wrong
		
		display dialog errorMessage & return & return & "Unable to read user information from network directory." with icon stop buttons {"OK"} default button {"OK"} with title "Outlook Exchange Setup"
		error number -128
	end try
	
	repeat with i from 1 to count of paragraphs in userInformation
		
		set AppleScript's text item delimiters to {": "}
		if paragraph i of userInformation begins with "EMailAddress:" then
			try
				set emailAddress to text item 2 of paragraph i of userInformation
			on error
				set AppleScript's text item delimiters to {""}
				set emailAddress to characters 2 through end of paragraph (i + 1) of userInformation as string
			end try
		end if
		
		set AppleScript's text item delimiters to {": "}
		if paragraph i of userInformation begins with "dsAttrTypeNative:co:" then
			try
				set userCountry to text item 2 of paragraph i of userInformation
			on error
				set AppleScript's text item delimiters to {""}
				set userCountry to characters 2 through end of paragraph (i + 1) of userInformation as string
			end try
		end if
		
		set AppleScript's text item delimiters to {": "}
		if paragraph i of userInformation begins with "dsAttrTypeNative:company:" then
			try
				set userCompany to text item 2 of paragraph i of userInformation
			on error
				set AppleScript's text item delimiters to {""}
				set userCompany to characters 2 through end of paragraph (i + 1) of userInformation as string
			end try
		end if
		
		set AppleScript's text item delimiters to {": "}
		if paragraph i of userInformation begins with "dsAttrTypeNative:department:" then
			try
				set userDepartment to text item 2 of paragraph i of userInformation
			on error
				set AppleScript's text item delimiters to {""}
				set userDepartment to characters 2 through end of paragraph (i + 1) of userInformation as string
			end try
		end if
		
		set AppleScript's text item delimiters to {": "}
		if paragraph i of userInformation begins with "dsAttrTypeNative:physicalDeliveryOfficeName:" then
			try
				set userOffice to text item 2 of paragraph i of userInformation
			on error
				set AppleScript's text item delimiters to {""}
				set userOffice to characters 2 through end of paragraph (i + 1) of userInformation as string
			end try
		end if
		
		set AppleScript's text item delimiters to {": "}
		if paragraph i of userInformation begins with "dsAttrTypeNative:sAMAccountName:" then
			try
				set userShortName to text item 2 of paragraph i of userInformation
			on error
				set AppleScript's text item delimiters to {""}
				set userShortName to characters 2 through end of paragraph (i + 1) of userInformation as string
			end try
		end if
		
		set AppleScript's text item delimiters to {": "}
		if paragraph i of userInformation begins with "dsAttrTypeNative:wWWHomePage:" then
			try
				set userWebPage to text item 2 of paragraph i of userInformation
			on error
				set AppleScript's text item delimiters to {""}
				set userWebPage to characters 2 through end of paragraph (i + 1) of userInformation as string
			end try
		end if
		
		set AppleScript's text item delimiters to {": "}
		if paragraph i of userInformation begins with "City:" then
			try
				set userCity to text item 2 of paragraph i of userInformation
			on error
				set AppleScript's text item delimiters to {""}
				set userCity to characters 2 through end of paragraph (i + 1) of userInformation as string
			end try
		end if
		
		set AppleScript's text item delimiters to {": "}
		if paragraph i of userInformation begins with "FAXNumber:" then
			try
				set userFax to text item 2 of paragraph i of userInformation
			on error
				set AppleScript's text item delimiters to {""}
				set userFax to characters 2 through end of paragraph (i + 1) of userInformation as string
			end try
		end if
		
		set AppleScript's text item delimiters to {": "}
		if paragraph i of userInformation begins with "FirstName:" then
			try
				set userFirstName to text item 2 of paragraph i of userInformation
			on error
				set AppleScript's text item delimiters to {""}
				set userFirstName to characters 2 through end of paragraph (i + 1) of userInformation as string
			end try
		end if
		
		set AppleScript's text item delimiters to {": "}
		if paragraph i of userInformation begins with "JobTitle:" then
			try
				set userTitle to text item 2 of paragraph i of userInformation
			on error
				set AppleScript's text item delimiters to {""}
				set userTitle to characters 2 through end of paragraph (i + 1) of userInformation as string
			end try
		end if
		
		set AppleScript's text item delimiters to {": "}
		if paragraph i of userInformation begins with "LastName:" then
			try
				set userLastName to text item 2 of paragraph i of userInformation
			on error
				set AppleScript's text item delimiters to {""}
				set userLastName to characters 2 through end of paragraph (i + 1) of userInformation as string
			end try
		end if
		
		set AppleScript's text item delimiters to {": "}
		if paragraph i of userInformation begins with "MobileNumber:" then
			try
				set userMobile to text item 2 of paragraph i of userInformation
			on error
				set AppleScript's text item delimiters to {""}
				set userMobile to characters 2 through end of paragraph (i + 1) of userInformation as string
			end try
		end if
		
		set AppleScript's text item delimiters to {": "}
		if paragraph i of userInformation begins with "PhoneNumber:" then
			try
				set userWorkPhone to text item 2 of paragraph i of userInformation
			on error
				set AppleScript's text item delimiters to {""}
				set userWorkPhone to characters 2 through end of paragraph (i + 1) of userInformation as string
			end try
		end if
		
		set AppleScript's text item delimiters to {": "}
		if paragraph i of userInformation begins with "PostalCode:" then
			try
				set userPostalCode to text item 2 of paragraph i of userInformation
			on error
				set AppleScript's text item delimiters to {""}
				set userPostalCode to characters 2 through end of paragraph (i + 1) of userInformation as string
			end try
		end if
		
		set AppleScript's text item delimiters to {": "}
		if paragraph i of userInformation begins with "RealName:" then
			try
				set userFullName to text item 2 of paragraph i of userInformation
			on error
				set AppleScript's text item delimiters to {""}
				set userFullName to characters 2 through end of paragraph (i + 1) of userInformation as string
			end try
		end if
		
		set AppleScript's text item delimiters to {": "}
		if paragraph i of userInformation begins with "State:" then
			try
				set userState to text item 2 of paragraph i of userInformation
			on error
				set AppleScript's text item delimiters to {""}
				set userState to characters 2 through end of paragraph (i + 1) of userInformation as string
			end try
		end if
		
		set AppleScript's text item delimiters to {": "}
		if paragraph i of userInformation begins with "Street:" then
			try
				set userStreet to text item 2 of paragraph i of userInformation
			on error
				set AppleScript's text item delimiters to {""}
				set userStreet to characters 2 through end of paragraph (i + 1) of userInformation as string
			end try
		end if
		
	end repeat
	
	set AppleScript's text item delimiters to {";Kerberosv5;;", ";"}
	
	try
		set userKerberosRealm to text item 2 of userInformation
	end try
	
	set AppleScript's text item delimiters to {""}
	
	if emailAddress is "" then
		
		-- something went wrong
		
		display dialog errorMessage & return & return & "Unable to read email address from network directory." with icon stop buttons {"OK"} default button {"OK"} with title "Outlook Exchange Setup"
		error number -128
	end if
	
else if emailFormat is 1 and displayName is 1 then
	
	-- Pull user information from the account settings of the local user account
	
	set userShortName to short user name of (system info)
	set userFullName to long user name of (system info)
	
	-- first.last@domain.com and full name displays as "Last, First"
	
	set AppleScript's text item delimiters to ", "
	set userFirstName to last text item of userFullName
	set userLastName to word 1 of text item 1 of userFullName
	set AppleScript's text item delimiters to ""
	set emailAddress to userFirstName & "." & userLastName & "@" & domainName
	
else if emailFormat is 1 and displayName is 2 then
	
	-- Pull user information from the account settings of the local user account
	
	set userShortName to short user name of (system info)
	set userFullName to long user name of (system info)
	
	-- first.last@domain.com and full name displays as "First Last"
	
	set AppleScript's text item delimiters to " "
	set userFirstName to word 1 of text item 1 of userFullName
	set userLastName to last text item of userFullName
	set AppleScript's text item delimiters to ""
	set emailAddress to userFirstName & "." & userLastName & "@" & domainName
	
else if emailFormat is 2 and displayName is 1 then
	
	-- Pull user information from the account settings of the local user account
	
	set userShortName to short user name of (system info)
	set userFullName to long user name of (system info)
	
	-- first@domain.com and full name displays as "Last, First"
	
	set AppleScript's text item delimiters to ", "
	set userFirstName to last text item of userFullName
	set userLastName to word 1 of text item 1 of userFullName
	set AppleScript's text item delimiters to ""
	set emailAddress to userFirstName & "@" & domainName
	
else if emailFormat is 2 and displayName is 2 then
	
	-- Pull user information from the account settings of the local user account
	
	set userShortName to short user name of (system info)
	set userFullName to long user name of (system info)
	
	-- first@domain.com if full name displays as "First Last"
	
	set AppleScript's text item delimiters to " "
	set userFirstName to word 1 of text item 1 of userFullName
	set userLastName to last text item of userFullName
	set AppleScript's text item delimiters to ""
	set emailAddress to userFirstName & "@" & domainName
	
else if emailFormat is 3 and displayName is 1 then
	
	-- Pull user information from the account settings of the local user account
	
	set userShortName to short user name of (system info)
	set userFullName to long user name of (system info)
	
	-- flast@domain.com and full name displays as "Last, First"
	
	set AppleScript's text item delimiters to ", "
	set userFirstName to last text item of userFullName
	set userLastName to word 1 of text item 1 of userFullName
	set AppleScript's text item delimiters to ""
	set emailAddress to (character 1 of userFirstName) & userLastName & "@" & domainName
	
else if emailFormat is 3 and displayName is 2 then
	
	-- Pull user information from the account settings of the local user account
	
	set userShortName to short user name of (system info)
	set userFullName to long user name of (system info)
	
	-- flast@domain.com and full name displays as "First Last"
	
	set AppleScript's text item delimiters to " "
	set userFirstName to word 1 of text item 1 of userFullName
	set userLastName to last text item of userFullName
	set AppleScript's text item delimiters to ""
	set emailAddress to (character 1 of userFirstName & userLastName & "@" & domainName)
	
else if emailFormat is 4 and displayName is 1 then
	
	-- Pull user information from the account settings of the local user account
	
	set userShortName to short user name of (system info)
	set userFullName to long user name of (system info)
	
	-- shortName@domain.com and full name displays as "Last, First"
	
	set AppleScript's text item delimiters to ", "
	set userFirstName to last text item of userFullName
	set userLastName to word 1 of text item 1 of userFullName
	set AppleScript's text item delimiters to ""
	set emailAddress to userShortName & "@" & domainName
	
else if emailFormat is 4 and displayName is 2 then
	
	-- Pull user information from the account settings of the local user account
	
	set userShortName to short user name of (system info)
	set userFullName to long user name of (system info)
	
	-- shortName@domain.com and full name displays as "First Last"
	
	set AppleScript's text item delimiters to " "
	set userFirstName to word 1 of text item 1 of userFullName
	set userLastName to last text item of userFullName
	set AppleScript's text item delimiters to ""
	set emailAddress to userShortName & "@" & domainName
	
else
	
	-- something went wrong
	
	display dialog errorMessage & return & return & "Unable to parse account information from local OS X account." with icon stop buttons {"OK"} default button {"OK"} with title "Outlook Exchange Setup"
	error number -128
	
end if

--------------------------------------------
-- End collecting user information
--------------------------------------------

--------------------------------------------
-- Begin logging user information
--------------------------------------------

writeLog("User information...")
writeLog("First Name: " & userFirstName)
writeLog("Last Name: " & userLastName)
writeLog("Email Address: " & emailAddress)
writeLog("Department: " & userDepartment)
writeLog("Office: " & userOffice)
writeLog("Company: " & userCompany)
writeLog("Work Phone: " & userWorkPhone)
writeLog("Mobile Phone: " & userMobile)
writeLog("FAX: " & userFax)
writeLog("Title: " & userTitle)
writeLog("Street: " & userStreet)
writeLog("City: " & userCity)
writeLog("State: " & userState)
writeLog("Postal Code: " & userPostalCode)
writeLog("Country: " & userCountry)
writeLog("Web Page: " & userWebPage)
writeLog(return)

--------------------------------------------
-- End logging user information
--------------------------------------------

--------------------------------------------
-- Begin account setup
--------------------------------------------

tell application "Microsoft Outlook"
	activate
	set working offline to "true"
	try
		set group similar folders to unifiedInbox
		my writeLog("Set Group Similar Folders to " & unifiedInbox & ": Successful.")
	on error
		my writeLog("Set Group Similar Folders to " & unifiedInbox & ": Failed.")
	end try
	
	try
		set hide on my computer folders to hideOnMyComputerFolders
		my writeLog("Set Hide On My Computer Folders to " & hideOnMyComputerFolders & ": Successful.")
	on error
		my writeLog("Set Hide On My Computer Folders to " & hideOnMyComputerFolders & ": Failed.")
	end try
	
	if verifyEMailAddress is true then
		set verifyEmail to display dialog "Please verify your email address is correct." default answer emailAddress with icon 1 with title "Outlook Exchange Setup" buttons {"Cancel", "Verify"} default button {"Verify"}
		set emailAddress to text returned of verifyEmail
		my writeLog("User verified email address as " & emailAddress & ".")
	end if
	
	if verifyServerAddress is true then
		set verifyServer to display dialog "Please verify your Exchange Server name is correct." default answer ExchangeServer with icon 1 with title "Outlook Exchange Setup" buttons {"Cancel", "Verify"} default button {"Verify"}
		set ExchangeServer to text returned of verifyServer
		my writeLog("User verified server address as " & ExchangeServer & ".")
	end if
	
	-- create the Exchange account
	
	try
		set newExchangeAccount to make new exchange account with properties ¬
			{name:"Mailbox - " & userFullName, user name:domainPrefix & userShortName, full name:userFullName, email address:emailAddress, server:ExchangeServer, use ssl:ExchangeServerRequiresSSL, port:ExchangeServerSSLPort, ldap server:DirectoryServer, ldap needs authentication:DirectoryServerRequiresAuthentication, ldap use ssl:DirectoryServerRequiresSSL, ldap max entries:DirectoryServerMaximumResults, ldap search base:DirectoryServerSearchBase, receive partial messages:downloadHeadersOnly, background autodiscover:disableAutodiscover}
		my writeLog("Create Exchange account: Successful.")
	on error
		
		-- something went wrong
		
		my writeLog("Create Exchange account: Failed.")
		
		display dialog errorMessage & return & return & "Unable to create Exchange account." with icon stop buttons {"OK"} default button {"OK"} with title "Outlook Exchange Setup"
		error number -128
		
	end try
	
	-- The following lines enable Kerberos support if the userKerberos property above is set to true.
	
	if useKerberos is true then
		try
			set use kerberos authentication of newExchangeAccount to useKerberos
			set principal of newExchangeAccount to userKerberosRealm
			my writeLog("Set Kerberos authentication: Successful.")
		on error
			
			-- something went wrong
			
			my writeLog("Set Kerberos authentication: Failed.")
			
			display dialog errorMessage & return & return & "Unable to set Exchange account to use Kerberos." with icon stop buttons {"OK"} default button {"OK"} with title "Outlook Exchange Setup"
			error number -128
			
		end try
	end if
	
	try
		-- The Me Contact record is automatically created with the first account.
		-- Set the first name, last name, email address and other information using Active Directory.
		
		set first name of me contact to userFirstName
		set last name of me contact to userLastName
		set email addresses of me contact to {address:emailAddress, type:work}
		set department of me contact to userDepartment
		set office of me contact to userOffice
		set company of me contact to userCompany
		set business phone number of me contact to userWorkPhone
		set mobile number of me contact to userMobile
		set business fax number of me contact to userFax
		set job title of me contact to userTitle
		set business street address of me contact to userStreet
		set business city of me contact to userCity
		set business state of me contact to userState
		set business zip of me contact to userPostalCode
		set business country of me contact to userCountry
		set business web page of me contact to userWebPage
		my writeLog("Populate Me Contact information: Successful.")
	on error
		my writeLog("Populate Me Contact information: Failed.")
	end try
	
	-- Set Outlook to be the default application
	-- for mail, calendars and contacts.
	
	try
		set system default mail application to true
		set system default calendar application to true
		set system default address book application to true
		my writeLog("Set Outlook as default mail, calendar and contacts application: Successful.")
	on error
		my writeLog("Set Outlook as default mail, calendar and contacts application: Failed.")
	end try
	
	-- We're done.
	
end tell

tell application "Microsoft Outlook"
	activate
	set working offline to "false"
end tell

--------------------------------------------
-- End account setup
--------------------------------------------

--------------------------------------------
-- Begin script cleanup
--------------------------------------------


try
	do shell script "/bin/rm $HOME/Library/LaunchAgents/net.talkingmoose.OutlookExchangeSetup5.plist"
	writeLog("Delete OutlookExchangeSetup5.plist file from user LaunchAgents folder: Successful.")
on error
	writeLog("Delete OutlookExchangeSetup5.plist file from user LaunchAgents folder: Failed.")
end try

try
	do shell script "/bin/launchctl remove net.talkingmoose.OutlookExchangeSetup5"
	writeLog("Unload OutlookExchangeSetup5.plist launch agent: Successful.")
on error
	writeLog("Unload OutlookExchangeSetup5.plist launch agent: Failed.")
end try

writeLog(return)
writeLog(return)
writeLog(return)

--------------------------------------------
-- End script cleanup
--------------------------------------------

--------------------------------------------
-- Begin script handlers
--------------------------------------------

on writeLog(logMessage)
	set logFile to (path to home folder as string) & "Library:Logs:OutlookExchangeSetup5.log"
	set rightNow to short date string of (current date) & " " & time string of (current date) & tab
	if logMessage is return then
		set logInfo to return
	else
		set logInfo to rightNow & logMessage & return
	end if
	set openLogFile to open for access file logFile with write permission
	write logInfo to openLogFile starting at eof
	close access file logFile
end writeLog

--------------------------------------------
-- End script handlers
--------------------------------------------
