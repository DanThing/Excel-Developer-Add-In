VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExecAndUndo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'===================================
 ' Module    					| clsExecAndUndo
 ' Company   					| JKP Application Development Services (c) 2005
 ' Author    					| Jan Karel Pieterse
 ' Created   					| 31-8-2005
 ' Purpose   					| Class module, stores the objects processed and handles the exection of the commands
 ' Copyright 					| This code is free for you to use for applications for personal use. It is not allowed to use this for a commercial program, unless you have my consent. If you want to include this code in freeware, make sure you add :
'-------------------------------|------------------------------------------
' This code originates from    	|Jan Karel Pieterse
 ' Company                     	|JKP Application Development Services (c) 2005 www.jkp-ads.com
'-------------------------------|------------------------------------------
 
'### Short explanation of the subs and functions shown above

 ' Sub/Function 			|Goal 
 ' AddAndProcessObject 		|This is the entry routine of this class. It receives the object and its properties to be modified and adds a member to the collection of changed objects (members of clsUndoObject), so changes can be undone later on. It calls the routines needed to store previous settings and execute the change. 
 ' ResetUndo 				|Clears the undo stack of this utility (run when instances of the class are destroyed) 
 ' UndoAll 					|Method to call to undo all operations stored in the collection of changed objects. This procedure is called when control-z is pressed or Edit, undo is selected. 
 ' UndoLast 				|Only undoes the last change. This needs to be tied to a button or a shortcut-key programmatically in your application. 
 ' UndoCount 				|Returns the number of objects stored in the collection 

'=====================================

Option Explicit

 Private mcolUndoObjects As Collection
 Private mUndoObject As clsUndoObject

 Public Function AddAndProcessObject(oObj As Object, sProperty As String, vValue As Variant) As Boolean
     Set mUndoObject = New clsUndoObject
     With mUndoObject
         Set .ObjectToChange = oObj
         .NewValue = vValue
         .PropertyToChange = sProperty
         mcolUndoObjects.Add mUndoObject
         If .ExecuteCommand = True Then
             AddAndProcessObject = True
         Else
             AddAndProcessObject = False
         End If
     End With
 End Function

 Private Sub Class_Initialize()
     Set mcolUndoObjects = New Collection
 End Sub

 Private Sub Class_Terminate()
     ResetUndo
 End Sub

 Public Sub ResetUndo()
     While mcolUndoObjects.Count > 0
         mcolUndoObjects.Remove (1)
     Wend
     Set mUndoObject = Nothing
 End Sub

 Public Sub UndoAll()
     Dim lCount As Long
     For lCount = mcolUndoObjects.Count To 1 Step -1
         Set mUndoObject = mcolUndoObjects(lCount)
         mUndoObject.UndoChange
         Set mUndoObject = Nothing
     Next
     ResetUndo
 End Sub

 Public Sub UndoLast()
     Dim lCount As Long
     If mcolUndoObjects.Count >= 1 Then
         Set mUndoObject = mcolUndoObjects(mcolUndoObjects.Count)
         mUndoObject.UndoChange
         mcolUndoObjects.Remove mcolUndoObjects.Count
         Set mUndoObject = Nothing
     Else
         ResetUndo
     End If
 End Sub

 Public Function UndoCount() As Long
     UndoCount = mcolUndoObjects.Count
 End Function
