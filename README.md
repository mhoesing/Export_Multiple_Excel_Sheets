Export_Multiple_Excel_Sheets
============================

Export multiple ACL tables to 1 Excel Workbook, each table into individual sheets within the workbook

The video at this link will walk through the scripts and discuss Powershell setup.

http://www.wowt.com/ 

The Powershell script reads all the .CSV comma separated delimited files in a directory and writes them to separate sheets within one Excel Workbook. The Powershell script is called from the ACL script with the EXECUTE command          
     EXECUTE "powershell.exe %v_powershellpath%Combine_CSVs_Into_1_Workbook.ps1 %v_csvpath% %v_xlsfilename%"
    
     all file names and path names have no blanks, underscores are used for sapcing, not blanks

The CSVs are created using ACL EXPORT commands. The ACL tables chosen for EXPORT in this version are specifically listed and include all three ACL tables that are in the project. The script could be modified to solicit user input as to which ACL tables to export using ACCEPT FIELDS "xf" commands, or listing desired .FIL files with a DIRECTORY command assuming those .FIL have a naming convention that facilitates proper selection. The individual .CSVs could be removed after the combined Workbook is created to save disk space, but this script does not remove those CSV files.

The PS1 Powershell script is created by an ACL subscript using LIST UNFORMATTED commands to generate the plain text file with the PS1 extension that Powershell can execute. Thanks Phil Lim for the concept of creating VB scripts within the ACL script, it adapts nicely to Powershell scripts and results in one not having to copy the PS1 and remember where it was placed.

The variables in the front of the script that need to be adjusted for your situation are:
   v_powershellpath    the location of where you wish ACL subscript to write the PS1 script file, this could be the                                ACL project directory
   v_csvpath           the location of the EXPORTed cvs delimited files, this also could be the project path
   v_xlsfilename       the name of the Workbook that conatins all the sheets

The resulting Workbook on my machine with Office 2010 is an .xlsx file.
