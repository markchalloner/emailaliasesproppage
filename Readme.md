README
======

Email Aliases Tab for Active Directory Users
--------------------------------------------

This project provides a dll (emailaliasesproppage.dll) that adds a tab to the User objects in Active Directory called Email Aliases. This allows editing of the proxyAddresses attribute through the AD Users GUI even if Exchange is not
installed.

This tab is particularly useful in conjunction with to Google Apps Directory Sync tool to synchronise email aliases to Google.

Warning!
--------

__This code was knocked up in a weekend and has only been tested on a 2003 SBS and 2008 R2 server. I cannot be held responsible for any damage you do to your servers/AD by installing it!__

Installation
------------

###To install on Windows SBS 2003:

* Download the emailaliasesproppage.dll from bin into your system32 folder
* Register the dll with: __regsrv32 emailaliasesproppage__
* Continue to All section

###To install on Windows 2008 R2: 

* Download the emailaliasesproppage.dll from bin into your SysWoW64 folder
* Register the dll with: __%systemroot%\SysWoW64\regsvr32.exe emailaliasesproppage__
* Continue to All section

###All

* Add an admin property page to user-Display CN:
  * Download and run ADSI Edit
  * Connect to your AD
  * Choose Configuration > CN=Configuration,... > CN=DisplaySpecifiers
  * Choose your language id (default is US CN=409)
      * A full list can be found at http://msdn.microsoft.com/nb-no/goglobal/bb964664.aspx
  * In the right hand window double click CN=userDisplay
  * Choose adminPropertyPages and add the value (where _number_ is an unused number): 
    ___number_,{40EC22BE-2BCC-4FB4-85F7-BC775C59A8C5}__
* Close and reopen Active Directory Users and Computers to see change

###To uninstall

* Remove the admin property page from the user-Display CN
* Unregister the dll: __regsrv32 -u emailaliasesproppage__
* Delete the dll

About
-----

The code is based on the userproppage code in the Microsoft Platform SDK and a lot of Googling. It is my first proper attempt with VC++ so apologies if there are obvious mistakes. Feel free to fork and adapt. I would appreciate pull 
requests for any fixes made.
