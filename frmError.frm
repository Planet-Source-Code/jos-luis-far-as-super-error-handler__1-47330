VERSION 5.00
Begin VB.Form frmError 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "An Error has occured"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Print Error"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4920
      TabIndex        =   26
      ToolTipText     =   "Print the error data on your printer"
      Top             =   6420
      Width           =   1035
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Continue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2760
      TabIndex        =   25
      ToolTipText     =   "Continue program operation."
      Top             =   6420
      Width           =   1035
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&End "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      TabIndex        =   24
      ToolTipText     =   "End the program."
      Top             =   6420
      Width           =   1035
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6000
      TabIndex        =   23
      ToolTipText     =   "Help Information"
      Top             =   6420
      Width           =   1035
   End
   Begin VB.TextBox txtDesc 
      Height          =   855
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   4140
      Width           =   4215
   End
   Begin VB.Label lblCmd 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Continue unless in an error loop."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   660
      TabIndex        =   27
      Top             =   6420
      Width           =   1935
   End
   Begin VB.Image imgExclaim 
      Height          =   1020
      Left            =   60
      Picture         =   "frmError.frx":0000
      Top             =   1320
      Width           =   1125
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Note: The information on this form is copied to the clipboard and saved in error.log file in the application's directory."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   540
      TabIndex        =   22
      Top             =   5100
      Width           =   6495
   End
   Begin VB.Line Line5 
      X1              =   540
      X2              =   6420
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line4 
      X1              =   540
      X2              =   6960
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label lblDesc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Please enter a description of the steps that produced the error."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   540
      TabIndex        =   21
      Top             =   4260
      Width           =   2175
   End
   Begin VB.Label lblProgVersTxt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Version 1.00.0001"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   120
      Width           =   4155
   End
   Begin VB.Label lblProgVers 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Program Version:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   18
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblFileTimeTxt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   480
      Width           =   4155
   End
   Begin VB.Label lblProgPathTxt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   780
      Width           =   4155
   End
   Begin VB.Label lblOSTxt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2760
      TabIndex        =   15
      Top             =   1080
      Width           =   4155
   End
   Begin VB.Label lblFileTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "File Time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1020
      TabIndex        =   14
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblProgPath 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Program Path:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1020
      TabIndex        =   13
      Top             =   780
      Width           =   1695
   End
   Begin VB.Label lblOS 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Operating System:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1020
      TabIndex        =   12
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblLineNumTxt 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   2280
      Width           =   4155
   End
   Begin VB.Label lblProcTxt 
      BackColor       =   &H00FFFFC0&
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   1980
      Width           =   4155
   End
   Begin VB.Label lblModuleTxt 
      BackColor       =   &H00FFFFC0&
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   1680
      Width           =   4155
   End
   Begin VB.Label lblErrDescTxt 
      BackColor       =   &H00FFFFC0&
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2760
      TabIndex        =   8
      Top             =   3240
      Width           =   4155
   End
   Begin VB.Label lblErrCatTxt 
      BackColor       =   &H00FFFFC0&
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   2940
      Width           =   4155
   End
   Begin VB.Label lblErrDesc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Error Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1020
      TabIndex        =   6
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblErrCat 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Error Category:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1020
      TabIndex        =   5
      Top             =   2940
      Width           =   1695
   End
   Begin VB.Line Line3 
      X1              =   480
      X2              =   6900
      Y1              =   2580
      Y2              =   2580
   End
   Begin VB.Label lblLineNum 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Line Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1020
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label lblProc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Procedure:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1020
      TabIndex        =   3
      Top             =   1980
      Width           =   1695
   End
   Begin VB.Label lblModule 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Module:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1020
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   6900
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   6900
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Label lblErrNumTxt 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   2640
      Width           =   4155
   End
   Begin VB.Label lblErrNum 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Error Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1020
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  Copyright:   (c) 1998 GridLinx Software
'  Author:      George Lissauer
'
'  Modified and Fixed in Aug/02/2003 By José Luis Farías - JoseloFarias[at]adinet.com.uy
'  Salto - República Oriental del Uruguay
'  Posted in Planet-Source-Code by himself
'  Don't work in IDE mode. Please Compile the App.
'  Please, if you use this Code in your own proyects, please sendme a program copy (source code if better)
'  Use the freeware VB Add-In MZ-Tools 3 (www.mztools.com)
'  to add Procedure line numbers incremented by 1 (if you like, of course...)
'  And to add the following custom error handler:
'
'   On Error GoTo {PROCEDURE_NAME}_Error
'
'    {PROCEDURE_BODY}
'
'{PROCEDURE_NAME}_Error:
'
'     If Err <> 0 Then
'         frmError.ErrMsg "{MODULE_NAME}", "{PROCEDURE_NAME}", Erl, Err
'     End If
     
Option Explicit

#Const EXE_TYPE = True        'SET TO FALSE IF USED IN A DLL.
Private m_sExeType As String   'Type of exe: dll or exe - used to find timestamp of application

'these are available as long as form is loaded
Private m_lPlatform As Long
Private m_sPlatform As String
Private m_sVersion As String

Const EM_GETLINECOUNT = &HBA
Const EM_GETLINE = &HC4
Const iMAX_CHAR_PER_LINE = 65

Const VbLogToFile = 2            'Should be in VB Constants

Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long

'Win32 API calls
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Public ID, Version, Build As String
Private Sub GetSystemOS()
     Dim OSinfo As OSVERSIONINFO
     Dim RetValue As Integer
     OSinfo.dwOSVersionInfoSize = 148
     OSinfo.szCSDVersion = Space$(128)
     RetValue = GetVersionExA(OSinfo)
     With OSinfo
     Select Case .dwPlatformId
      Case 1
          Select Case .dwMinorVersion
              Case 0
                  ID = "Windows 95"
              Case 10
                  If .dwBuildNumber >= 2183 Then
                      ID = "Windows 98 SE"
                  Else
                      ID = "Windows 98"
                  End If
              Case 90
                  ID = "Windows Millennium Edition"
          End Select
      Case 2
          Select Case .dwMajorVersion
              Case 3
                  ID = "Windows NT 3.51"
              Case 4
                  ID = "Windows NT 4.0"
              Case 5
                  If .dwMinorVersion = 0 Then
                      ID = "Windows 2000"
                  ElseIf .dwMinorVersion = 1 Then
                      ID = "Windows XP"
                  Else
                      ID = "Windows Sever 2003"
                  End If
          End Select
      Case Else
         ID = "Failed"
    End Select
     Version = .dwMajorVersion & "." & .dwMinorVersion
     Build = .dwBuildNumber
     End With
End Sub
Private Sub cmdContinue_Click()
    Dim s As String
    Dim sPath As String
    If Right(App.Path, 1) = "\" Then
        sPath = App.Path
    Else
        sPath = App.Path & "\"
    End If
    App.StartLogging sPath & "Error.log", VbLogToFile
    s = FormatError()
    App.LogEvent s, vbLogEventTypeInformation
    Clipboard.SetText s  'save in clipboard
    Unload Me
    Set frmError = Nothing
End Sub
Private Sub cmdEnd_Click()
    Dim s As String
    Dim sPath As String
    If Right(App.Path, 1) = "\" Then
        sPath = App.Path
    Else
        sPath = App.Path & "\"
    End If
    
    App.StartLogging sPath & "Error.log", VbLogToFile
    s = FormatError()
    App.LogEvent s, vbLogEventTypeInformation
    Clipboard.SetText s  'save in clipboard
    Unload Me
    #If EXE_TYPE = True Then
       End      'COMMENT OUT THE END STATEMENT FOR USE IN DLL
    #End If
    Set frmError = Nothing
End Sub
Private Sub cmdHelp_Click()
    Dim s As String
    s = "Please provide Customer Service with the information displayed."
    s = s & vbCrLf & vbCrLf
    s = s & "This infomation is copied to the clipboard and logged to " & vbCrLf
    s = s & "the file: ERROR.LOG in the:" & vbCrLf & App.Path & " directory."
    s = s & vbCrLf & vbCrLf
    s = s & "Please enter any additional information in the text box."
    s = s & vbCrLf & vbCrLf
    s = s & "If you have a printer connected to your computer, you can " & vbCrLf
    s = s & "print out the information using the Print Error button."
    s = s & vbCrLf & vbCrLf
    s = s & "Continue program operation using the the Continue button."
    s = s & vbCrLf & vbCrLf
    s = s & "If this same error repeats, exit the program with the End button."
    MsgBox s, vbInformation, "An Error Occured in Your Program ..."
End Sub
Private Sub cmdPrint_Click()
    Dim s As String
    s = FormatError()
    Printer.ScaleMode = vbInches
    Printer.ScaleLeft = -0.5
    Printer.Print s
    Printer.EndDoc
End Sub
Private Sub Form_Activate()
    'Get current Windows configuration
    On Error GoTo ActError
    #If EXE_TYPE Then
       m_sExeType = ".exe"
    #Else
       m_sExeType = ".dll"
    #End If
       
    lblFileTimeTxt = FileDateTime(App.Path & "\" & App.EXEName & m_sExeType)
    lblProgPathTxt = ShortPath(App.Path & "\" & App.EXEName & m_sExeType, 45)
    lblProgPathTxt.ToolTipText = App.Path & "\" & App.EXEName & m_sExeType
    lblOSTxt = ID & " Version " & Version & " Build " & Build
ActError:
    
    If Err.Number = 53 Then
         m_sExeType = ".dll"
         DoEvents              'Avoid a locked loop if no exe yet.
         Resume
    Else
         Exit Sub
    End If
End Sub
Private Sub Form_Load()
'    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE)
    ' Center the form
    Me.Move ((Screen.Width - Me.Width) / 2), ((Screen.Height - Me.Height) / 2)
    lblProgVersTxt = FormatVersion(App.Major, App.Minor, App.Revision)
    GetSystemOS
    'GetSystemInfo lpSysInfo
    'm_lProcessor = lpSysInfo.dwProcessorType
    Me.Caption = "An Error Occured in: " & App.Title & " " & Format$(Now, "dddd mmm dd,yyyy @ hh:mm:ss")
    lblErrDescTxt.Caption = "A value has been assigned to a property, that is outside its permissible range."
    lblErrDescTxt.Caption = lblErrDescTxt.Caption & "A value has been assigned to a property, that is outside its permissible range."
    lblNote.Caption = "Note: The infomation on this form is copied to the clipboard and logged to " & _
    "the file: ERROR.LOG in the: " & App.Path & " directory when you click the Continue or End buttons."
End Sub
Private Function FormatVersion(Major&, Minor&, Optional Revision) As String
    If IsMissing(Revision) Then
       FormatVersion = Format(Major, "#0.") & Format(Minor, "00")
    Else
      FormatVersion = Format(Major, "#0.") & Format(Minor, "00.") & Format(Revision, "0000")
    End If
End Function
Private Function FormatError() As String
    Dim s As String
    s = vbCrLf
    s = s & "'" & String(70, "-") & vbCrLf
    s = s & "Error in: " & App.Title & vbCrLf
    s = s & "Time: " & Format$(Now, "dddd mmm dd,yyyy @ hh:mm:ss") & vbCrLf
    s = s & "Version: " & lblProgVersTxt.Caption & vbCrLf
    s = s & "File Time: " & lblFileTimeTxt.Caption & vbCrLf
    s = s & "Program Path: " & lblProgPathTxt.ToolTipText & vbCrLf
    s = s & "Operating System: " & lblOSTxt.Caption & vbCrLf
    s = s & "Module: " & lblModuleTxt.Caption & vbCrLf
    s = s & "Procedure: " & lblProcTxt.Caption & vbCrLf
    s = s & "Line Number: " & lblLineNumTxt.Caption & vbCrLf
    s = s & "Error Number: " & lblErrNumTxt.Caption & vbCrLf
    s = s & "Error Category: " & lblErrCatTxt.Caption & vbCrLf
    s = s & "Error Description: " & vbCrLf & lblErrDescTxt.Caption & vbCrLf
    If Len(Trim$(txtDesc.Text)) > 0 Then
       s = s & "User Description: " & vbCrLf
       s = s & sCommentBlock(txtDesc) & vbCrLf
    End If
    s = s & "'" & String(70, "-") & vbCrLf
    
    FormatError = s
    
End Function
Private Function fGetLine(lLine As Long, ctl As Control) As String
    Dim iByteLo As Integer
    Dim iByteHi As Integer
    Dim sBuffer As String
    Dim iChrs As Integer
    iByteLo = iMAX_CHAR_PER_LINE And (255)
    iByteHi = Int(iMAX_CHAR_PER_LINE / 256)
    sBuffer = Chr$(iByteLo) & Chr$(iByteHi) + Space$(iMAX_CHAR_PER_LINE - 2)
    iChrs = SendMessage(ctl.hwnd, EM_GETLINE, lLine, sBuffer)
    fGetLine = Left$(sBuffer, iChrs)
End Function
Private Function fGetLineCount(ctl As Control) As Long
    Dim lCount As Long
    lCount = SendMessage(ctl.hwnd, EM_GETLINECOUNT, 0&, 0&)
    fGetLineCount = lCount
End Function
Private Function sCommentBlock(ctl As Control) As String
    Dim sText As String
    Dim lLines As Long
    Dim lLine As Long
    lLines = fGetLineCount(ctl)
    sText = ""
    If lLines > 0 Then
       sText = sText & "'" & vbCrLf
       For lLine = 0 To lLines - 1
          sText = sText & "'" & vbTab & fGetLine(lLine, ctl) & vbCrLf
       Next lLine
    End If
    sText = sText & "'"
    sCommentBlock = sText
End Function
Public Sub ErrMsg(sMN As String, sPN As String, iLine As Integer, lErr As Long, Optional vErrCat)
    lblModuleTxt.Caption = sMN
    lblProcTxt.Caption = sPN
    lblLineNumTxt.Caption = Format(iLine, "000")
    lblErrNumTxt.Caption = Format(lErr, "000")
    If IsMissing(vErrCat) Then
       lblErrCatTxt.Caption = "Visual Basic Error"
    Else
       lblErrCatTxt.Caption = vErrCat
    End If
    lblErrDescTxt.Caption = Error$(lErr)
    Me.Show vbModal
End Sub
Private Function ShortPath(sPath As String, iMaxLen As Integer) As String
    Const DRIVE_LENGTH = 3         'Length of Drive, colon & slash in path
    Dim sLeft As String            'Left part of Path
    Dim sRight As String           'Right part of Path
    Dim iNextPos As Integer        'Position of Next "\"
    Dim iStart As Integer          'Position to start from
    If Len(sPath) <= iMaxLen Then
       ShortPath = sPath
       Exit Function
    End If
    iStart = DRIVE_LENGTH + 1                 'Start looking after Drive:\
    sLeft = Left$(sPath, DRIVE_LENGTH)        'Extract the drive from full path
    sRight = Right$(sPath, Len(sPath) - 3)    'Remove drive from right part
    Do While Len(sLeft & sRight) > iMaxLen    'Do until path shorter than Max Length
       iNextPos = InStr(iStart, sPath, "\")   'Find next "\" in path
       If iNextPos = 0 Then Exit Do           'Exit if no more "\" in path
       sLeft = sLeft & "...\"                 'Add another ...\ to short path
       sRight = Right$(sPath, Len(sPath) - iNextPos)   '
       iStart = iNextPos + 1
    Loop
    ShortPath = sLeft & sRight
End Function
