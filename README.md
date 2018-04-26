# ExcelTitan
The Excel Add-in that hooks into Excel Titan's repository

I wanted a way to make it easier to find exactly what I'm looking for on the ExcelTitan Github by either file name or topic That's why I have created an Excel Add-in that breaks down the codes into topic but also type and search for a file name.

You aren't restricted to just looking for VBA codes either, this add-in searches my Excel Formulas tutorials and provides links to other useful programs.

## VBA Codes and Formulas
Here are the links to the Github Repositories for:
- [VBA Codes](https://github.com/ExcelTitan/vbaCodes)
- [Formulas](https://github.com/ExcelTitan/Excel_Formulas)

## Compatibility: 
- PC Only
- Excel 2007 - 2016
- Works on 32-bit or 64-bit

## Continual Improvement
Through feedback and my every day use I will continue to improve the functionality of this add-in. 
A change log will also be provided to provide information on updates.

## Installation

- Unzip the zip file to extract the contents
- Move the add-in to your local Microsoft AddIns directory.
- Open Excel
- Click the File tab
- Click Options
- Click the Add-Ins category on the left bar of the Excel Options Dialog Box
- In the Manage box near the bottom, click Excel Add-ins, and then click Go. The Add-Ins dialog box appears.
- Click Browse
- Navigate to the xla or xlam add-in and select it.
- Make sure the checkbox next to the add-in is checked, then click OK.

If the add-in keeps disappearing when you restart Excel, follow the directions below.

Important Note: There are a few additional steps to this installation due to an Office Security Update released in July 2016.


## Trust the Folder Location

The folder that the add-in file is saved in needs to be added as a Trusted Location in Excel. The instructions on how to trust 
the folder location are below.

-  Open the Excel Options menu
   - File > Options
   - Open the Trust Center menu and Add a new location
   - Trust Center > Trust Center Settings… > Trusted Locations > Add new location
   - Browse for the folder that you saved the add-in file in.
   - Press OK, the folder should now appear in the Trusted Locations list.
   - Press OK on the Trust Center and Excel Options menu to close the menus.

-  The add-in is now in a trusted location and the add-in will appear every time you open Excel.


## Unblock the File
Some users still have issues with the add-in’s ribbon disappearing after trusting the folder location. If this happens to you, 
you will need to Unblock the file by changing a file property.

- Locate the Add-in file (.xla, .xlam) in Windows Explorer.
- Right-click the file and select Properties.
- At the bottom of the General tab you should see a Security section. Check the box that says Unblock.
- Press the OK button.

Close Excel completely and re-open it. The add-in should now load and any custom ribbons will appear.

You will only need to do this unblock one time. However, if you download an updated version of the file then you will have to 
repeat the steps above to unblock it.

Hopefully these additional security steps will be fixed in a future update to Office. If you completed all the steps above then 
you should see the add-ins ribbon tab load every time you open Excel.
