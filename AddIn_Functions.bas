Attribute VB_Name = "AddIn_Functions"
'===================================
' **Name**|AddIn_Functions
' **Type**|Module
' **Purpose**|Container Module to hold customised Functions
' **Useage**|See each function for details
' **Arthor**|Daniel Boyce
' **Version**| 1.0.20180702
'-----------------------------------
' - AddIn_Functions
'    - OperationCompleted
'    - OperationCancelled
'    - RemoveRMUR
'    - ExportThisWS
'    - DiffBetween
'    - kAccuracy
'    - CheckForSupersession
'    - WSExists
'    - IsWorkBookOpen
'    - openWB
'    - getFile
'    - getFolder
'    - Lastrow
'    - LastCol
'    - RndUp
'    - ClearNameRngs
'    - FS
'    - getArrayHeader
'    - RemoveSelectRMUR
'    - chkSelectSupersession
'    - DisabledFunction
'-----------------------------------
'===================================

Option Explicit
Option Compare Text

Sub OperationCompleted()
MsgBox "Operation has been Completed.", vbOKOnly, Company
End Sub

Sub OperationCancelled()
MsgBox "Operation has been cancelled by User.", vbOKOnly, Company
SpeedUp False
End
End Sub

'===================================
' Removes the sufix of either 'RM' or 'UR' from an itemcode
'
' - @method RemoveRMUR
'   - @param {Variant} itemcode
' - @returns {String} RemveRMUR
'===================================
Function RemoveRMUR(itemcode As Variant) As String
Dim a As Long
    a = InStr(5, itemcode, "RM", vbTextCompare)
    If a > 6 Then
        RemoveRMUR = Left(itemcode, a - 1)
        Exit Function
    End If
    a = InStrRev(itemcode, "UR", , vbTextCompare)
    If a > 6 Then
         RemoveRMUR = Left(itemcode, a - 1)
        Exit Function
    End If
    RemoveRMUR = itemcode
End Function

'===================================
' Exports a copy of the active worksheet without formulas
'
' - @method ExportThisWS
'===================================
Sub ExportThisWS()
With ActiveSheet
    .Copy
    .Cells().Copy
    .Cells().PasteSpecial xlPasteValues
End With
End Sub

'===================================
' Returns the difference between two numbers
'
' - @method DiffBetween
'   - @param {Variant} firstValue
'   - @param {Variant} secondValue
' - @returns {Long}
'===================================
Function DiffBetween(ByVal firstValue As Variant, ByVal secondValue As Variant) As Long
Dim fst As Long, snd As Long
With Application.WorksheetFunction
    fst = .Max(firstValue, secondValue)
    snd = .Min(firstValue, secondValue)
End With
DiffBetween = fst - snd
End Function

'===================================
' This function is used for the OV accuracy report only.
' The math is as verbally supplied by Yoshikatsu Kinya 17/04/2018
'
' - @method kAccuracy
'   - @param {Variant} planned
'   - @param {Variant} actual
' - @returns {Double}
'===================================
Public Function kAccuracy(planned As Variant, actual As Variant) As Double
If planned > actual Then
    kAccuracy = (actual - planned) / planned
Else
    kAccuracy = actual / planned
End If
If kAccuracy > 1 Then _
    kAccuracy = 0
End Function

'===================================
' This function is used to get the superseeding or preseeding part number
' for a provide partnumber.
'
' This will return a single answer as string if getAll = False.
' This will return an array of all results if getAll = True.
'
' - @method CheckForSupersession
'   - @param {Variant} sku
'   - @param [{XlSearchDirection}] direction (xlNext = 1, xlPrevious = 2)
'   - @param [{boolean}] getAll
' - @returns {variant}

Private Function CheckForSupersession(sku As Variant, Optional supersession_direction As XlSearchDirection = xlNext, Optional getAll As Boolean = False) As Variant
Attribute CheckForSupersession.VB_Description = "Checks to see if the item code(s) selected have been superseeded. This returns the latest Item Code."
On Error GoTo errExit
Dim aCount As Long
Dim tempStr As String
reDo:
    For aCount = LBound(SuperSessionId) To UBound(SuperSessionId)
        If SuperSessionId(aCount, 1) = sku Then
            If SuperSessionChange(aCount, direction) = "" Then
                GoTo notFound
            Else
                sku = SuperSessionChange(aCount, supersession_direction)
                If getAll = True Then
                    tempStr = tempStr & sku & ","
                Else
                    tempStr = sku
                End If
                GoTo reDo
            End If
        End If
    Next aCount
    
safeExit:
    CheckForSupersession = tempStr
Exit Function
    
notFound:
    CheckForSupersession = sku
Exit Function

errExit:
    errHandler Err
End Function

'===================================
' Returns True/False if a  worksheet exists in a given workbook
'
' - @method WSExists
'   - @param {Workbook} TargetWB
'   - @param {String} WSName
' - @returns {Boolean}

Public Function WSExists(ByRef TargetWB As Workbook, ByVal WSName As String) As Boolean
Dim WSO As Worksheet
    On Error GoTo errExit
    Set WSO = TargetWB.Sheets(WSName)
    If Not WSO Is Nothing Then WSExists = True
errExit:
    Set WSO = Nothing
End Function

'===================================
' Returns True/False if a given workbook is currently open
'
' - @method IsWorkBookOpen
'   - @param {String} WorkbookName
' - @returns {Boolean}

Function IsWorkBookOpen(ByVal WorkbookName As String) As Boolean
Dim WBO As Workbook
    On Error GoTo errExit
    Set WBO = Workbooks(WorkbookName)
    If Not WBO Is Nothing Then IsWorkBookOpen = True
    Set WBO = Nothing
    Exit Function
errExit:
    Set WBO = Nothing
End Function

'===================================
' This function checks to see if the workbook is already open.
' If it is then it uses that workbook otherwise it opens the workbook.
'
' - @method openWB
'   - @param {String} fname
' - @return {Workbook}

Function openWB(fname As String) As Workbook
On Error GoTo errExit
    If IsWorkBookOpen(fname) Then
       Set openWB = Workbooks(fname)
       Exit Function
    Else
        Set openWB = Workbooks.Open(fname, False, True)
        Exit Function
    End If
    
    Err.Raise CustomError.err3, "openWB", "Cannot find workbook " & fname & "."
    
errExit:
    errHandler Err
End Function

'===================================
' Shows a dialog box to get the location and file name for the indictaed file.
' If cancel selected then this returns vbNullString
'
' -@method getFile
'   - @param {String} fname
' - @return {String}

Function getFile(fname As String) As String
Dim intChoice As Integer
' User to select file to open
With Application.FileDialog(msoFileDialogOpen)
    .Top = Me.Parent.Application.Top + 100
    .Left = Me.Parent.Application.Left + 100
    .AllowMultiSelect = False
    .Title = fname & " - " & Company
    intChoice = .Show
    If intChoice <> 0 Then
        getFile = .SelectedItems(1)
        Exit Function
    End If
End With
errExit:
getFile = vbNullString
End Function

'===================================
' Shows a dialog box to get the path and name for the indicated folder.
' If cancel selected then this returns vbNullString
'
' - @method getFolder
'   - @param {String} fname
' - @return {String}

Function getFolder(fname As String) As String
Dim intChoice As Integer
With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    .Title = fname & " - " & Company
    intChoice = .Show
    If intChoice <> 0 Then
        getFolder = .SelectedItems(1)
        Exit Function
    End If
End With
getFolder = vbNullString
End Function

'===================================
' This Function takes a worksheet as an input
' and returns the last used row in the sheet
'
' - @method Lastrow
'   - @param {Worksheet} sh
' - @return {Long}

Function Lastrow(Sh As Worksheet)
    Lastrow = Sh.Cells.Find(What:="*", _
                            after:=Sh.Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
End Function

'===================================
' This Function takes a worksheet as an input
' and returns the last used column in the sheet
'
' - @method Lastcol
'   - @param {Worksheet} sh
' - @return {Long}

Function LastCol(Sh As Worksheet)
    LastCol = Sh.Cells.Find(What:="*", _
                            after:=Sh.Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
End Function

'===================================
' This Function takes a Number or range as an input
' and rounds it up to the next integer
'
' - @method RndUp
'   - @param {Variant} numbervalue
' - @return {Long}

Function RndUp(numbervalue As Variant) As Long
With Application.WorksheetFunction
    RndUp = .RoundUp(numbervalue, 0)
End With
End Function

'===================================
' clears the Named Ranged for the indicated worksheet.
'
' - @method ClearNameRngs
'   - @param {Worksheet} ws

Sub ClearNameRngs(WS As Worksheet)
Dim xName As Name
For Each xName In thisWB.Names
    If InStr(1, xName, WS.Name) Then xName.Delete
Next xName
End Sub

'===================================
' Returns the default path seperator.
'   '\' for Windows systems
'   ':' for Classic Mac OS
'   '/' for Unix
'
' - @method FS
' - @returns {string}

Function FS() As String
  FS = Application.PathSeparator
End Function

'===================================
' Function to get the position of a
' value within a 2-dimentional Array.
'
' - @method getArrayHeader
'   - @param {String} lookfor
'   - @param {Variant} inArray
' - @returns {Long}

Function getArrayHeader(lookfor As String, inArray As Variant) As Long
On Error GoTo errExit
Dim itm As Long
    For itm = LBound(inArray(2)) To UBound(inArray(2))
        If inArray(1, itm) = lookfor Then
            getArrayHeader = itm
            Exit Function
        End If
    Next itm
errExit:
    getArrayHeader = 0
    Err.Raise CustomError.Err4, "getArrayHeader"
End Function

'===================================
' Function to remove the suffix 'RM' and 'UR' from a selection range.
'
' - @method RemoveSelectRMUR

Sub RemoveSelectRMUR()
Dim progBar As ProgressBar
Dim sel As Variant
Dim sCount As Long
Dim temp As String
    On Error GoTo errExit
    If Selection.Count < 1 Then
        Err.Raise CustomError.Err5, "RemoveRMUR", "Selection contains no data."
    End If
    Set progBar = New ProgressBar
    With progBar
        .Top = Selection.Parent.Application.Top + 100
        .Left = Selection.Parent.Application.Left + 100
        .Title = Company
        .TotalActions = Selection.Cells().Count
        .showbar
        .StatusMessage = "Removing RM/UR from itemcodes."
    End With
    sel = Selection.value
    For sCount = LBound(sel) To UBound(sel)
        progBar.NextAction
        temp = sel(sCount, 1)
        temp = RemoveRMUR(temp)
        sel(sCount, 1) = temp
    Next sCount
    Selection = sel
safeExit:
    progBar.Terminate
    OperationCompleted
Exit Sub

errExit:
    progBar.Terminate
    errHandler Err
End Sub

'===================================
' Function to return the latest Supersession for a selection range.
'
' - @method chkSelectSupersession

Sub chkSelectSupersession()
Dim wb As Workbook
Dim progBar As ProgressBar
Dim direction As XlSearchDirection
Dim cel As Range
Dim sel As Variant
Dim sCount As Long, lr As Long
Dim temp As String
    On Error GoTo errExit
    If Selection.Cells().Count < 1 Then
        Err.Raise CustomError.Err5, "chkSelectSupersession", "Selection contains no data."
    End If
    Set progBar = New ProgressBar
    With progBar
        .Top = Selection.Parent.Application.Top + 100
        .Left = Selection.Parent.Application.Left + 100
        .Title = Company
        .TotalActions = Selection.Cells().Count
        .showbar
        .StatusMessage = "Getting Supersession data from Komunity."
    End With
    On Error Resume Next
        Set wb = openWB(getSetting(URL3_))
        Set SuperSessionData = wb.Worksheets("ItemList")
    On Error GoTo errExit
    If Err <> 0 Then Err.Raise CustomError.Err2
    lr = Lastrow(SuperSessionData)
    With SuperSessionData
        SuperSessionId = .Range(.Cells(2, 1), .Cells(lr, 1)).value
        SuperSessionChange = .Range(.Cells(2, 7), .Cells(lr, 8)).value
    End With
    If Selection.Cells().Count = 1 Then
        Set cel = Selection
        cel = Application.Run("CheckForSupersession", Selection)
        GoTo safeExit
    End If
    Set cel = Nothing
    sel = Selection.value
        For sCount = LBound(sel) To UBound(sel)
            progBar.NextAction
            temp = sel(sCount, 1)
            temp = Application.Run("CheckForSupersession", temp)
            sel(sCount, 1) = temp
        Next sCount
    Selection = sel
safeExit:
    wb.Close False
    progBar.Terminate
    OperationCompleted
Exit Sub

errExit:
    progBar.Terminate
    errHandler Err
End Sub

'===================================
' Function to alert a user that a selected
' function is not yet implemented
'
' - @method DisabledFunction

Sub DisabledFunction()
    MsgBox "This function is currently disabled.", vbOKOnly, Company
    End
End Sub
