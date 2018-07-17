VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddIn_SettingForm 
   Caption         =   "Komatsu Australia Excel Add-In Settings"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6060
   OleObjectBlob   =   "AddIn_SettingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddIn_SettingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===================================
' **Name**|AddIn_SettingForm
' **Type**|Userform
' **Purpose**|
' **Author**|Daniel Boyce
' **Version**|1.0.20180702
'-----------------------------------
'- AddIn_SettingForm
'    - btn_Cancel_Click
'    - btn_chkUpdates_Click
'    - btn_update_Click
'    - UserForm_Initialize
'    - UserForm_QueryClose
'-----------------------------------
'===================================

Option Explicit

Private Sub btn_Cancel_Click()
    Unload Me
End Sub

Private Sub btn_chkUpdates_Click()
    Application.Run "DisabledFunction"
End Sub

Private Sub btn_update_Click()

    Application.Run "DeleteFromCellMenu"

    If Not (Me.companyNameUpdate = vbNullString) Or Not (Me.companyNameUpdate = " ") Then
            changeSetting "CompanyName", companyNameUpdate.value
    End If
        
    If Not (Me.URL1Update = vbNullString) Or Not (Me.URL1Update = " ") Then
            changeSetting "URL1_", URL1Update.value
    End If
        
    If Not (Me.URL2Update = vbNullString) Or Not (Me.URL2Update = " ") Then
            changeSetting "URL2_", URL2Update.value
    End If
        
    If Not (Me.URL3Update = vbNullString) Or Not (Me.URL3Update = " ") Then
            changeSetting "URL3_", URL3Update.value
    End If
        
    If Not (Me.URL4Update = vbNullString) Or Not (Me.URL4Update = " ") Then
            changeSetting "URL4_", URL4Update.value
    End If
        
    If Not (Me.URL5Update = vbNullString) Or Not (Me.URL5Update = " ") Then
            changeSetting "URL5_", URL5Update.value
    End If
    
    With ThisWorkbook.Worksheets("Settings")
        .Range("EnableLogging") = Me.opt_enablelogging.value
        .Range("EnableContextMenu") = Me.opt_enablecontextmenu.value
        .Range("EnableSupersession") = Me.opt_supersession.value
        .Range("EnableRemoveRMUR") = Me.opt_RemoveRMUR.value
        .Range("EnableAddItemcodeDashes") = Me.opt_AddItemcodeDashes.value
        .Range("EnableExportThisWS") = Me.opt_ExportThisWS.value
    End With
    
    Application.Run "AddToCellMenu"
    
    Me.Hide
    
End Sub

Private Sub UserForm_Initialize()
    With ThisWorkbook.Worksheets("Settings")
    
        Me.companyNameUpdate.value = .Range("CompanyName").value
        
        Me.URL1Update.value = .Range("URL1_").value
        Me.URL2Update.value = .Range("URL2_").value
        Me.URL3Update.value = .Range("URL3_").value
        Me.URL4Update.value = .Range("URL4_").value
        Me.URL5Update.value = .Range("URL5_").value
        
        Me.opt_enablelogging.value = .Range("EnableLogging").value
        Me.opt_enablecontextmenu.value = .Range("EnableContextMenu").value
        
        Me.opt_supersession.value = .Range("EnableSupersession").value
        Me.opt_RemoveRMUR.value = .Range("EnableRemoveRMUR").value
        Me.opt_AddItemcodeDashes.value = .Range("EnableAddItemcodeDashes").value
        Me.opt_ExportThisWS.value = .Range("EnableExportThisWS").value
    
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Unload Me
End Sub
