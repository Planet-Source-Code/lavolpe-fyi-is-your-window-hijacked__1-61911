VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   345
      Index           =   1
      Left            =   195
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1650
      Width           =   4275
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1020
      Width           =   4275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   600
      Left            =   240
      TabIndex        =   0
      Top             =   210
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Your window has been hijacked?  No"
      Height          =   315
      Left            =   630
      TabIndex        =   3
      Top             =   2340
      Width           =   3480
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' 2 apis & 2 constants used to compare window procedures
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_WNDPROC As Long = -4
Private Declare Function GetClassLong Lib "user32.dll" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GCL_WNDPROC As Long = -24

Private Sub Command1_Click()
IsHijacked
'^^ added here in case user has too much fun and clicks test button more than once
'   so that the label will change back to green before testing is reapplied

' opt for user to abort test
If MsgBox("Notice your window has not been subclassed by another procedure." & vbNewLine & vbNewLine & _
    "Click OK to temporarily subclass the window & see the result", vbOKCancel + vbQuestion, "Testing") = vbOK Then
    
    SubclassForm hwnd       ' subclass the window
    IsHijacked              ' see if window was hijacked
    SubclassForm hwnd       ' unsubclass the window
    
    MsgBox "Window has now been unsubclassed. No chance to crash application if you hit End", vbInformation + vbOKOnly
End If

End Sub

Private Function IsHijacked() As Boolean
' test routine to check for hijacking.
' For the paranoid ;)

' One possible use? Say you expect your pc to have spyware that may globally subclass
' windows. Well, you can create a simple form with typical objects on it that may
' be bait to such spys (textbox, IE control, etc). Then you can run your compiled
' application and test not only the window handle itself, but all child window handles.
' Should the results not match, I would assume you have something on your pc that is
' monitoring or changing your window messages.

' Note that this does not apply to window hooks: separate topic

' Now for a second scenario. Should you be subclassing your own windows, then you may
' want to keep track of which windowprocedure is the current subclasser and test
' the result against that so you don't fire a false alarm

Dim hWndProc As Long, cWndProc As Long

    ' get current window procedure
    hWndProc = GetWindowLong(hwnd, GWL_WNDPROC)
    ' get the base window/class procedure
    cWndProc = GetClassLong(hwnd, GCL_WNDPROC)
    
    ' update form objects
    Text1(0) = "Expected window procedure handle is " & cWndProc
    Text1(1) = "Actual window procedure handle is " & hWndProc
    
    ' test against base window procedure
    If hWndProc = cWndProc Then
        Label1.Caption = "Your window has been hijacked?  No"
        Label1.BackColor = vbGreen
    Else
        Label1.Caption = "Your window has been hijacked?  Yes"
        Label1.BackColor = vbRed
        IsHijacked = True
    End If
        
End Function

Private Sub Form_Load()
    Call IsHijacked    ' fill in the form display objects
End Sub
