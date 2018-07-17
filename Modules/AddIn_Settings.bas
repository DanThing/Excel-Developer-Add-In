Attribute VB_Name = "AddIn_Settings"
'===================================
' **Name**|AddIn_Settings
' **Type**|Module
' **Purpose**|Module to hold functions that get/set this Add-ins setting.
' **Useage**|See each function for details
' **Author**|Daniel Boyce
' **Version**|1.0.20180713
'-----------------------------------
'- AddIn_Settings
'    - Company
'    - getSetting
'    - changeSetting
'    - LicencingForm
'-----------------------------------
'===================================

Option Explicit

'===================================
' Function to get the company name from settings
'
' - @method Company
' - @returns {string}

Public Function Company() As String
    Company = getSetting("CompanyName")
End Function

'===================================
' Function to get the named setting from settings
'
' - @method getSetting
'   - @param {String} settingName
' - @returns {string}

Function getSetting(settingName As String) As Variant
    getSetting = ThisWorkbook.Worksheets("Settings").Range(settingName).value
End Function

'===================================
' Function to change an existing setting
'
' - @method changeSetting
'   - @param {String} settingName
'   - @param {Variant} newValue

Sub changeSetting(settingName As String, newValue As Variant)
    ThisWorkbook.Worksheets("Settings").Range(settingName).value = CStr(newValue)
End Sub

'===================================
' Shows the Addin licencing information to the user.
'
' - @method LicencingForm

Sub LicencingForm()
Dim usrfrm As New LicenceForm
    usrfrm.Show
End Sub

