VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ByVal and ByRef Comparison"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   495
      Left            =   1583
      TabIndex        =   0
      Top             =   1223
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTest_Click()

'Declare x1 and x2 as strings
Dim x1$, x2$

  'Assign values to both variables
  x1$ = "Test1"
  x2$ = "Test2"
  
  'Call the test function
  Call test(x1$, x2$)
  
  'display a message box with the
  'results of the 2 variables from
  'the function
  MsgBox x1$
  MsgBox x2$
  
  'Note that the ByVal variable stayed
  'the same while the ByRef variable
  'was changed.
  'ByVal will send a copy of the variable
  'to the function while ByRef will send
  'the actual variable to the function
  
End Sub

Public Function test(ByVal str1 As String, ByRef str2 As String)

'Simple funciton to add a ! to the beginning of each string

  str1 = "!" & str1
  str2 = "!" & str2
  
End Function
