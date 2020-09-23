VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmMain 
   Caption         =   "URL Monitor"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frame 
      Height          =   2325
      Left            =   0
      TabIndex        =   1
      Top             =   4980
      Width           =   8745
      Begin VB.CommandButton cmdGoSelected 
         Cancel          =   -1  'True
         Caption         =   "Go To Selected"
         Height          =   315
         Left            =   1050
         TabIndex        =   8
         Top             =   1950
         Width           =   1335
      End
      Begin VB.CheckBox chkClear 
         Caption         =   "Auto Clear List"
         Height          =   195
         Left            =   2460
         TabIndex        =   6
         Top             =   2010
         Width           =   1395
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go!"
         Height          =   315
         Left            =   8070
         TabIndex        =   5
         Top             =   150
         Width           =   615
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   60
         TabIndex        =   4
         Top             =   150
         Width           =   7965
      End
      Begin VB.ListBox lstMain 
         Height          =   1425
         Left            =   60
         TabIndex        =   3
         Top             =   480
         Width           =   8625
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear List"
         Height          =   315
         Left            =   90
         TabIndex        =   2
         Top             =   1950
         Width           =   915
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click a URL to copy to the Clipboard"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   6000
         TabIndex        =   7
         Top             =   2010
         Width           =   2580
      End
   End
   Begin SHDocVwCtl.WebBrowser brwMain 
      Height          =   4965
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8745
      ExtentX         =   15425
      ExtentY         =   8758
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'URL Monitor v1.0

'By Max Raskin, 15 April 2000

'E-Mail: maxim13@internet-zahav.net

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_SETHORIZONTALEXTENT = &H194

Dim IsBlank As Boolean

Private Sub brwMain_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
'Here is the main idea :
'Add any URL accessed during the entring to the site
'This shows any URL that isn't showed on the standard address bar
'of Internet Explorer, that let's you to get links of files accessed through
'CGI or scripts and add them to your site
    If IsBlank = True Then
        IsBlank = False
        Exit Sub
    End If
    lstMain.AddItem URL
End Sub

Private Sub cmdClear_Click()
'Clear the ListBox
    lstMain.Clear
End Sub

Private Sub cmdGoSelected_Click()
'Navigates if anything in the listbox is selected
    If lstMain.ListIndex <> -1 Then brwMain.Navigate2 lstMain.List(lstMain.ListIndex)
End Sub

Private Sub Form_Load()
'Start the browser with a blank page
    brwMain.Navigate2 "about:blank"
    IsBlank = True
'Add an horizontal scroll bar to the list box
    SendMessage lstMain.hwnd, &H194, 1500, 0&
End Sub

Private Sub Form_Resize()
'Resize the controls as needed and limit form height to 4000
If Me.WindowState <> vbMinimized Then
    If Me.Height < 4000 Then Me.Height = 4000
    brwMain.Width = Me.ScaleWidth
    brwMain.Height = Me.ScaleHeight - frame.Height
    frame.Top = brwMain.Height
    frame.Width = Me.ScaleWidth
    lstMain.Width = Me.ScaleWidth - 150
    txtAddress.Width = Me.ScaleWidth - cmdGo.Width - 150
    cmdGo.Left = txtAddress.Width + 100
End If
End Sub

Private Sub lstMain_Click()
'Copy the selected list entry to the clipboard
On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText lstMain.List(lstMain.ListIndex)
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
'If chkClear value set to true then automaticly clear the list box
    If chkClear.Value = 1 Then lstMain.Clear
'If enter is pressed and txtAddress.Text isn't empty then navigate to the desired URL
    If Trim(txtAddress.Text) <> "" Then If KeyAscii = vbKeyReturn Then brwMain.Navigate2 txtAddress.Text
End Sub

Private Sub cmdGo_Click()
'If chkClear value set to true then automaticly clear the list box
    If chkClear.Value = 1 Then lstMain.Clear
'Navigate to the desired site if txtAddress.Text isn't empty
    If Trim(txtAddress.Text) <> "" Then brwMain.Navigate2 txtAddress.Text
End Sub
