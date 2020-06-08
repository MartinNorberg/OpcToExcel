# OpcToExcel
Subscribes to opc tags and writes result to excel
You will net to have access to OpcNetApi

# Getting started
Fill Opc url and Excel report destination.
![Subscriptions](https://github.com/MartinNorberg/OpcToExcel/blob/master/img/SubscribeView.PNG)

Go to settings tab and fill in
Save cyclic - if you want to write to excel report every 'Cycle time', else it will write to file every time value changed.

Excel filename - Name given to excel file. If not filled it will be named 'OPCExport'

New file/day - Will create a new file each day and '_{DateTime.Today} will be added to filename'

If you want these values to get saved you can fill in them in app.config


![Settings](https://github.com/MartinNorberg/OpcToExcel/blob/master/img/SettingsView.PNG)
