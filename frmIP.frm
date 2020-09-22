VERSION 5.00
Begin VB.Form frmIP 
   Caption         =   "IP Utility"
   ClientHeight    =   4170
   ClientLeft      =   3060
   ClientTop       =   1785
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   2520
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   4335
      Begin VB.OptionButton optMyInfo 
         Caption         =   "Host Name and IP Address of this Machine"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   3615
      End
      Begin VB.OptionButton optIPAddr 
         Caption         =   "IP Address to Host Name"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   3375
      End
      Begin VB.OptionButton optHostName 
         Caption         =   "Host Name to IP Address"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   960
      TabIndex        =   0
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Host:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   615
   End
End
Attribute VB_Name = "frmIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStart_Click()
    If optHostName.Value = True Then
        Screen.MousePointer = vbHourglass
        Dim sHostName As String
           If SocketsInitialize() Then
              sHostName = Text1.Text
              Text2.Text = GetIPFromHostName(sHostName)
              SocketsCleanup
           Else
                MsgBox "Windows Sockets for 32 bit Windows " & _
                       "environments is not successfully responding."
           End If
        If Text2.Text = "" Then
        Else
            Text2.BackColor = &H80000005
        End If
        Screen.MousePointer = vbNormal
    ElseIf optIPAddr.Value = True Then
        Screen.MousePointer = vbHourglass
        Text1.Text = GetHostNameFromIP(Text2.Text)
        If Text1.Text = "" Then
        Else
            Text1.BackColor = &H80000005
        End If
        Screen.MousePointer = vbNormal
    Else
        Screen.MousePointer = vbHourglass
        Text1.Text = GetIPHostName()
        Text2.Text = GetIPAddress()
        Screen.MousePointer = vbNormal
    End If
End Sub

Private Sub Form_Load()
    optHostName.Value = True
End Sub

Private Sub optHostName_Click()
    Text2.Text = ""
    Text1.Text = ""
    Text2.Enabled = False
    Text2.BackColor = &H8000000F
    Text1.Enabled = True
    Text1.BackColor = &H80000005
End Sub

Private Sub optIPAddr_Click()
    Text2.Text = ""
    Text1.Text = ""
    Text2.Enabled = True
    Text2.BackColor = &H80000005
    Text1.Enabled = False
    Text1.BackColor = &H8000000F
End Sub

Private Sub optMyInfo_Click()
    Text2.Text = ""
    Text1.Text = ""
    Text2.Enabled = False
    Text2.BackColor = &H80000005
    Text1.Enabled = False
    Text1.BackColor = &H80000005
End Sub
