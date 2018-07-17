VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUndoObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'=========================================================================
 ' Module    : clsUndoObject
 ' Company   : JKP Application Development Services (c) 2005
 ' Author    : Jan Karel Pieterse
 ' Created   : 31-8-2005
 ' Purpose   : Class module, Contains each object processed and
 '             handles the exection of the command and the Undo
 ' Copyright : This code is free for you to use for applications
 '             for personal use.
 '             It is not allowed to use this for a commercial program,
 '             unless you have my consent.
 '             If you want to include this code in freeware, make sure you add :
'-------------------------------------------------------------------------

' This code originates from    : Jan Karel Pieterse
 ' Company                     : JKP Application Development Services (c) 2005
 '                                www.jkp-ads.com
'-------------------------------------------------------------------------

 ' The central properties of this class are:
	' ObjectToChange  	The object which is about to be changed 
	' PropertyToChange  The property of the object which is about to be changed 
	' NewValue  		The new value of the object's property 
	' OldValue  		The old value of the object's property 
'-------------------------------------------------------------------------
 
' The main functions are:
	' ExecuteCommand 	actually changes the property (function SetNewValue), after storing the old value of the property (function GetOldValue). 
	' UndoChange 		undoes the change made to the object. 
'=========================================================================
Option Explicit

 Private mUndoObject As Object
 Private msProperty As String
 Private mvNewValue As Variant
 Private mvOldValue As Variant

 Public Property Let PropertyToChange(sProperty As String)
     msProperty = sProperty
 End Property

 Public Property Get PropertyToChange() As String
     PropertyToChange = msProperty
 End Property

 Public Property Set ObjectToChange(oObj As Object)
     Set mUndoObject = oObj
 End Property

 Public Property Get ObjectToChange() As Object
     Set ObjectToChange = mUndoObject
 End Property

 Public Property Let NewValue(vValue As Variant)
     mvNewValue = vValue
 End Property

 Public Property Get NewValue() As Variant
     NewValue = mvNewValue
 End Property

 Public Property Let OldValue(vValue As Variant)
     mvOldValue = vValue
 End Property

 Public Property Get OldValue() As Variant
     OldValue = mvOldValue
 End Property

 Public Function ExecuteCommand() As Boolean
     ExecuteCommand = False
     If mUndoObject Is Nothing Then
     End If
     If mvNewValue = "" Then
     End If
     If msProperty = "" Then
     End If
     If GetOldValue Then
         SetNewValue
         ExecuteCommand = True
     Else
         'Failed to retrieve old value!
     End If
 End Function

 Private Function GetOldValue() As Boolean
     Dim oTemp As Object
     Dim lCount As Long
     Dim lProps As Long
     Dim vProps As Variant
     vProps = Split(PropertyToChange, ".")
     lProps = UBound(vProps)
     Set oTemp = ObjectToChange
     For lCount = 0 To lProps - 1
         Set oTemp = CallByName(oTemp, vProps(lCount), VbGet)
     Next
     If TypeOf oTemp Is Range Then
         If LCase(vProps(lProps)) = "value" Then
             vProps(lProps) = "Formula"
         End If
     End If
     OldValue = CallByName(oTemp, vProps(lProps), VbGet)
     If Err.Number = 0 Then
         GetOldValue = True
     Else
         GetOldValue = False
     End If
 End Function

 Private Function SetNewValue() As Boolean
     Dim oTemp As Object
     Dim lCount As Long
     Dim lProps As Long
     Dim vProps As Variant
     Dim vResult As Variant
     Err.Clear
     Set oTemp = ObjectToChange
     vProps = Split(PropertyToChange, ".")
     lProps = UBound(vProps)
     For lCount = 0 To lProps - 1
         Set oTemp = CallByName(oTemp, vProps(lCount), VbGet)
     Next
     If TypeOf oTemp Is Range Then
         If LCase(vProps(lProps)) = "value" Then
             vProps(lProps) = "Formula"
         End If
     End If
     vResult = CallByName(oTemp, vProps(lProps), VbLet, NewValue)
     If Err.Number = 0 Then
         SetNewValue = True
     Else
         SetNewValue = False
     End If
 End Function

 Public Function UndoChange()
     Dim oTemp As Object
     Dim lCount As Long
     Dim lProps As Long
     Dim vProps As Variant
     Dim vResult As Variant
     Set oTemp = ObjectToChange
     vProps = Split(PropertyToChange, ".")
     lProps = UBound(vProps)
     For lCount = 0 To lProps - 1
         Set oTemp = CallByName(oTemp, vProps(lCount), VbGet)
     Next
     If TypeOf oTemp Is Range Then
         If LCase(vProps(lProps)) = "value" Then
             vProps(lProps) = "Formula"
         End If
     End If
     vResult = CallByName(oTemp, vProps(lProps), VbLet, OldValue)
     If vResult <> "" Then
         UndoChange = True
     Else
         UndoChange = False
     End If
 End Function
