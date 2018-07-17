# Excel-Developer-Add-In

## Modules and their Proceedures

AddIn_MAIN
- SpeedUp
- openSettings 	

AddIn_Functions
- OperationCompleted
- OperationCancelled
- WSExists
- IsWorkBookOpen
- openWB
- getFile
- getFolder
- Lastrow
- LastCol
- RndUp
- ClearNameRngs
- FS
- getArrayHeader
- DisabledFunction

AddIn_ErrHandler
- errHandler
- LogDebug
- LogWarning
- LogError
- ProjectDump

AddIn_Settings
- Company
- getSetting
- changeSetting
- LicencingForm
-----------------------------------

## Userforms and Methods

AddIn_SettingForm
- UserForm_Initialize

LogForm
- UserForm_Initialize
- updateText

HelpForm
- UserForm_Initialize

CreateNewHelpForm
- Edit
-----------------------------------

## Settings

| Setting Name | Setting Value | Variable Type | Variable Name |
| --- | --- | --- | --- |
| EnableLogging | True | Boolean | opt_enablelogging |
| EnableContextMenu | True | Boolean | opt_enablecontextmenu |
| EnableExportThisWS | True | Boolean | opt_ExportThisWS |
|CompanyName | Name of the Company | String | CompanyName |
URL1 | http:\\www.companyurl.com |String |URL1_
URL2 ||String |URL2_
URL3 ||String |URL3_
URL4 | |String |URL4_
URL5 | |String |URL5_
Text1 | |String |
Text2 | |String |
Text3 | |String |
Text4 | |String |
Text5 | |String |
Number1 | |Long |
Number2 | |Long |
Number3 | |Long |
Number4 | |Long |
Number5 | |Long |
Boolean1 | |Boolean |
Boolean2 | |Boolean |
Boolean3 | |Boolean |
Boolean4 | |Boolean |
Boolean5 | |Boolean |

-----------------------------------

## VBA Library References

Description   Microsoft Excel 16.0 Object Library
- FullPath   C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE
- Major.Minor   1.9
- Name   Excel
- GUID   {00020813-0000-0000-C000-000000000046}
- Type   0

Description   OLE Automation
- FullPath   C:\Windows\System32\stdole2.tlb
- Major.Minor   2.0
- Name   stdole
- GUID   {00020430-0000-0000-C000-000000000046}
- Type   0

Description   Microsoft Office 16.0 Object Library
- FullPath   C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL
- Major.Minor   2.8
- Name   Office
- GUID   {2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}
- Type   0

Description   Microsoft Forms 2.0 Object Library
- FullPath   C:\WINDOWS\system32\FM20.DLL
- Major.Minor   2.0
- Name   MSForms
- GUID   {0D452EE1-E08F-101A-852E-02608C4D0BB4}
- Type   0

Description   Microsoft Scripting Runtime
- FullPath   C:\Windows\System32\scrrun.dll
- Major.Minor   1.0
- Name   Scripting
- GUID   {420B2830-E718-11CF-893D-00A0C9054228}
- Type   0

Description   Ref Edit Control
- FullPath   C:\Program Files\Microsoft Office\Root\Office16\REFEDIT.DLL
- Major.Minor   1.2
- Name   RefEdit
- GUID   {00024517-0000-0000-C000-000000000046}
- Type   0
-----------------------------------
