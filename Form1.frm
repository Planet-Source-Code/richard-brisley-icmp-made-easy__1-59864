VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView2 
      Height          =   2535
      Left            =   3240
      TabIndex        =   9
      Top             =   2760
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Hop Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Responding Host"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Time To Live"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   3240
      TabIndex        =   8
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Responding Host"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bytes"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Time To Live"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Round Trip Time"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Text            =   "5"
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "General"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3015
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "255"
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "localhost"
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Target"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Time To Live"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   2640
         Y1              =   960
         Y2              =   960
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "tracert"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ping"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Number of pings"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<:-) :WARNING: Option Explicit makes coding much safer but will cause difficulties until you declare all variables.
'<:-) (Code Fixer identifies most of them as Missing Dims)
'<:-) Run code using [Ctrl]+[F5] to find any others yourself.
'<:-) To auto-insert 'Option Explicit' in all new code: Tools|Options...|Editor tab, check 'Require Variable Declaration'.
Option Explicit
Dim WithEvents objICMP As clsICMP
Attribute objICMP.VB_VarHelpID = -1

Private Sub Command1_Click()
    Call objICMP.getPing(Text1.Text, Text2.Text, Text3.Text)
End Sub

Private Sub Command2_Click()
    Call objICMP.getTraceRT(Text1.Text)
End Sub

Private Sub Form_Load()
    Set objICMP = New clsICMP
End Sub

Private Sub objICMP_Error(msg As String)
    MsgBox msg
End Sub

Private Sub objICMP_PingResponce(strRespondingHost As String, strBytes As String, lngTimeToLiveMs As Long, lngRoundTripTimeMs As Long)
    ListView1.ListItems.Add , , strRespondingHost
    ListView1.ListItems.Item(ListView1.ListItems.Count).SubItems(1) = "OK"
    ListView1.ListItems.Item(ListView1.ListItems.Count).SubItems(2) = strBytes
    ListView1.ListItems.Item(ListView1.ListItems.Count).SubItems(3) = lngTimeToLiveMs
    ListView1.ListItems.Item(ListView1.ListItems.Count).SubItems(4) = lngRoundTripTimeMs
End Sub

Private Sub objICMP_PingStatus(strRespondingHost As String, RCode As String)
    ListView1.ListItems.Add , , strRespondingHost
    ListView1.ListItems.Item(ListView1.ListItems.Count).SubItems(1) = RCode
End Sub

Private Sub objICMP_SockStartup(MaxSockets As Long, MaxUDP As Long, Description As String, Status As String)
    Stop
End Sub

Private Sub objICMP_TracertResponce(strHopNumber As String, strRespondingHost As String, lngTimeToLive As Long)
    ListView2.ListItems.Add , , "OK"
    ListView2.ListItems.Item(ListView2.ListItems.Count).SubItems(1) = strHopNumber
    ListView2.ListItems.Item(ListView2.ListItems.Count).SubItems(2) = strRespondingHost
    ListView2.ListItems.Item(ListView2.ListItems.Count).SubItems(3) = lngTimeToLive
End Sub

Private Sub objICMP_TracertStatus(strStatus As String)
    ListView2.ListItems.Add , , strStatus
End Sub

Private Sub objICMP_WinsockStartup(MaxSockets As Long, MaxUDP As Long, Description As String, Status As String)
    MsgBox Description
    MsgBox Status
End Sub

