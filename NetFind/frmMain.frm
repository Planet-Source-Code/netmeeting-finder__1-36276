VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Netmeeting"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSearch 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1515
      Top             =   1635
   End
   Begin VB.Frame frmStopIP 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Stop IP"
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2520
      TabIndex        =   10
      Top             =   615
      Width           =   1695
      Begin VB.TextBox txtStopIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000000&
         Height          =   195
         Index           =   0
         Left            =   15
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   14
         Text            =   "255"
         Top             =   195
         Width           =   375
      End
      Begin VB.TextBox txtStopIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000000&
         Height          =   195
         Index           =   1
         Left            =   435
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "255"
         Top             =   195
         Width           =   375
      End
      Begin VB.TextBox txtStopIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   2
         Left            =   870
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "255"
         Top             =   195
         Width           =   375
      End
      Begin VB.TextBox txtStopIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   3
         Left            =   1305
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "255"
         Top             =   195
         Width           =   375
      End
      Begin VB.Label lblSep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   375
         TabIndex        =   17
         Top             =   195
         Width           =   75
      End
      Begin VB.Label lblSep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   795
         TabIndex        =   16
         Top             =   195
         Width           =   75
      End
      Begin VB.Label lblSep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   1230
         TabIndex        =   15
         Top             =   195
         Width           =   75
      End
   End
   Begin VB.Frame frmStart 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Start IP"
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2535
      TabIndex        =   2
      Top             =   150
      Width           =   1695
      Begin VB.TextBox txtStartIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   3
         Left            =   1290
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "255"
         Top             =   195
         Width           =   375
      End
      Begin VB.TextBox txtStartIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   2
         Left            =   870
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "255"
         Top             =   195
         Width           =   375
      End
      Begin VB.TextBox txtStartIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   1
         Left            =   435
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "255"
         Top             =   195
         Width           =   375
      End
      Begin VB.TextBox txtStartIP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   15
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "255"
         Top             =   195
         Width           =   375
      End
      Begin VB.Label lblSep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1230
         TabIndex        =   8
         Top             =   195
         Width           =   75
      End
      Begin VB.Label lblSep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   795
         TabIndex        =   6
         Top             =   195
         Width           =   75
      End
      Begin VB.Label lblSep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   375
         TabIndex        =   4
         Top             =   195
         Width           =   75
      End
   End
   Begin MSWinsockLib.Winsock wskConnect 
      Index           =   0
      Left            =   1980
      Top             =   1635
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1720
   End
   Begin VB.CommandButton cmdSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   495
      Left            =   2535
      TabIndex        =   1
      Top             =   1065
      Width           =   1680
   End
   Begin VB.ListBox lstNetwork 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   135
      Width           =   2370
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Not Scanning"
      Height          =   195
      Left            =   2520
      TabIndex        =   19
      Top             =   1905
      Width           =   975
   End
   Begin VB.Label lblCountFound 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 User(s) Found"
      Height          =   195
      Left            =   2520
      TabIndex        =   18
      Top             =   1635
      Width           =   1125
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intTimesAround As Integer

Private Sub cmdSearch_Click()
    Select Case cmdSearch.Caption
        '// Start the search, but first set some settings
        Case "Start"
            lstNetwork.Clear
            cmdSearch.Caption = "Stop"
            intTimesAround = 0
            Me.Caption = "Find Netmeeting - Scanning..."
            '// Start the scanning
            tmrSearch.Enabled = True
        '// End the search
        Case "Stop"
            cmdSearch.Caption = "Start"
            tmrSearch.Enabled = False
            Me.Caption = "Find Netmeeting"
            '// Stop the search
            lblStatus.Caption = "Not Scanning"
    End Select
End Sub

Private Sub Form_Load()
    Dim strMyIP As String
    Dim strGrpIP1 As String
    Dim strGrpIP2 As String
    
    '// Get the local machines IP address
    strMyIP = wskConnect(0).LocalIP
    
    '// Get the first two groups of the IP address to use for our range
    strGrpIP1 = Left(strMyIP, InStr(1, strMyIP, ".") - 1)
    txtStartIP(0).Text = strGrpIP1
    
    strGrpIP2 = Mid(strMyIP, InStr(1, strMyIP, ".") + 1, InStr(InStr(1, strMyIP, "."), strMyIP, ".") - 1)
    txtStartIP(1).Text = strGrpIP2
    
    '// Set the other start groups to 0
    txtStartIP(2).Text = 0
    txtStartIP(3).Text = 0
    
    '// Set the other stop groups to 255
    txtStopIP(2).Text = 255
    txtStopIP(3).Text = 255

End Sub

Private Sub tmrSearch_Timer()
On Error Resume Next

        '// Display number of users found
        lblCountFound.Caption = lstNetwork.ListCount & " user(s) found."
        
        '// Increase the number of times we have looped
        intTimesAround = intTimesAround + 1
        
        '// Load a new winsock
        Load wskConnect(intTimesAround)
        
        '// If we have more then 50 winsocks creates unload the last one
        If intTimesAround > 50 Then Unload wskConnect(intTimesAround - 50)
        
        '// Loop through the IP range
        If Val(txtStartIP(3)) < Val(txtStopIP(3)) Then
            '// Increase the IP address by one
            txtStartIP(3).Text = txtStartIP(3).Text + 1
        ElseIf Val(txtStartIP(3)) = Val(txtStopIP(3)) Then
            '// Increase the 3rd group in the IP address by one
            txtStartIP(2) = txtStartIP(2) + 1
            '// Reset the 4th group in the IP address to 0
            txtStartIP(3) = 0
        End If
        
        '// Check to see if the scan is complete
        If Val(txtStartIP(2)) > Val(txtStopIP(2)) Then
            '// Let the user know the scan is complete
            MsgBox "Scan Complete"
            '// So stop the scan
            Call cmdSearch_Click
            Exit Sub
        End If
        
        '// Connect to the IP address
        wskConnect(intTimesAround).Connect txtStartIP(0) & "." & txtStartIP(1) & "." & txtStartIP(2) & "." & txtStartIP(3), 1503
        
        '// Update the status with the current IP address
        lblStatus.Caption = "Scanning " & txtStartIP(0) & "." & txtStartIP(1) & "." & txtStartIP(2) & "." & txtStartIP(3)

End Sub

Private Sub txtStartIP_Change(Index As Integer)
On Error Resume Next
    
    '// Exit the sub if the last IP group is changed
    If Index = 3 Then Exit Sub
    
    '// Set the stop IP to the same as the start IP group
    If Index = 0 Or Index = 1 Then
        txtStopIP(Index).Text = txtStartIP(Index).Text
    End If
    
    '// If there are 3 numbers then move to the next group
    If Len(txtStartIP(Index)) = 3 Then
        txtStartIP(Index + 1).SetFocus
        txtStartIP(Index + 1).SelLength = 3
        Exit Sub
    End If
        
End Sub

Private Sub txtStartIP_GotFocus(Index As Integer)
    '// Exit the sub if the first IP group has the focus
    If Index = 0 Then Exit Sub

    '// Check for a period in the IP group
    If InStr(1, txtStartIP(Index - 1), ".") > 0 Then
        txtStartIP(Index - 1).Text = Mid(txtStartIP(Index - 1), 1, Len(txtStartIP(Index - 1)) - 1)
    End If
    
    '// Make sure the number is not greater then 255
    If Val(txtStartIP(Index - 1)) > 255 Then txtStartIP(Index - 1).Text = 255
End Sub

Private Sub txtStartIP_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    '// Exit the sub if the last group has been changed
    If Index = 3 Then Exit Sub

    '// If the decimal key has been pressed then move to the next group
    If KeyCode = vbKeyDecimal Then
        txtStartIP(Index + 1).SetFocus
        txtStartIP(Index + 1).SelLength = 3
    End If

End Sub

Private Sub txtStopIP_Change(Index As Integer)
On Error Resume Next
    '// Exit the sub if the last IP group is changed
    If Index = 3 Then Exit Sub
    
    '// If there are 3 numbers then move to the next group
    If Len(txtStopIP(Index)) = 3 Then
        txtStopIP(Index + 1).SetFocus
        txtStopIP(Index + 1).SelLength = 3
    End If
    
End Sub

Private Sub txtStopIP_GotFocus(Index As Integer)
    '// Exit the sub if the first IP group has the focus
    If Index = 0 Then Exit Sub

    '// Check for a period in the IP group
    If InStr(1, txtStopIP(Index - 1), ".") > 0 Then
        txtStopIP(Index - 1).Text = Mid(txtStopIP(Index - 1), 1, Len(txtStopIP(Index - 1)) - 1)
    End If
    
    '// Make sure the number is not greater then 255
    If Val(txtStopIP(Index - 1)) > 255 Then txtStopIP(Index - 1).Text = 255

End Sub

Private Sub txtStopIP_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    '// Exit the sub if the last group has been changed
    If Index = 3 Then Exit Sub

    '// If the decimal key has been pressed then move to the next group
    If KeyCode = vbKeyDecimal Then
        txtStopIP(Index + 1).SetFocus
        txtStopIP(Index + 1).SelLength = 3
    End If

End Sub

Private Sub wskConnect_Connect(Index As Integer)

    '// We have found a connection so add it to our list
    lstNetwork.AddItem wskConnect(Index).RemoteHostIP
    
    '// Now close the connection
    wskConnect(Index).Close
End Sub


