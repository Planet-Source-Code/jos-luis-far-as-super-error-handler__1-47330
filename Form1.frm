VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Divide By Zero"
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Object Not Found"
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
          Dim obj As CheckBox
1        On Error GoTo Command1_Click_Error

2         obj.Caption = "Hello World"
3         Exit Sub

Command1_Click_Error:

4          If Err <> 0 Then
5              frmError.ErrMsg "Form1", "Command1_Click", Erl, Err
6          End If
End Sub

Private Sub Command2_Click()
          Dim i As Integer
1        On Error GoTo Command2_Click_Error

2         i = i / 0
3         Exit Sub

Command2_Click_Error:

4          If Err <> 0 Then
5        frmError.ErrMsg "Form1", "Command2_Click", Erl, Err
6          End If
End Sub
