<#==============================================================================
         File Name : UPS-Inventory.ps1
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr.com)
                   : 
       Description : Uses SNMP to poll and track APC UPS devices using MS Excel.
                   : 
             Notes : Normal operation is with no command line options.  Commandline options noted below.
                   :
      Requirements : Requires the PowerShell SNMP library from https://www.powershellgallery.com/packages/SNMPv3
                   : Currently designed to poll APC UPS devices.  UPS NMC must have SNMP v3 active.
                   : Script checks for active SNMPv1, FTP, and SNMPv3.
                   : Will generate a new spreadsheet if none exists by using a text file located in the same folder
                   : as the script, one IP per line.  Default operation is to check for text file first, then if not
                   : found check for an existing spreadsheet also in the same folder.  If an existing spreadsheet
                   : is located the target list is compliled from column A.  It will also copy a master spreadsheet
                   : to a working copy that gets processed.  Up to 10 backup copies are retained prior to writing
                   : changes to the working copy.
                   : 
          Warnings : Excel is set to be visible (can be changed) so don't mess with it while the script is running or it can crash.
                   :   
             Legal : Public Domain. Modify and redistribute freely. No rights reserved.
                   : SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF 
                   : ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED.
                   :
           Credits : Code snippets and/or ideas came from many sources including but 
                   :   not limited to the following:
                   : 
    Last Update by : Kenneth C. Mazie                                           
   Version History : v1.0 - 08-16-22 - Original 
    Change History : v2.0 - 09-00-22 - Numerous operational & bug fixes prior to v3.
                   : v3.1 - 09-15-22 - Cleaned up final version for posting.
                   : v4.0 - 04-12-23 - Too many changes to list
                   : v4.1 - 07-03-23 - Added age and LDOS dates. 
                   : v5.0 - 01-17-24 - Fixed DNS lookup.  Fixed last test result.  Fixed color coding of hostname for
                   :                   numerous events.  Added hostname cell comments to describe color coding.
                   :                  
==============================================================================#>
