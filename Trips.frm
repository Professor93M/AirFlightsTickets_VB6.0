VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Trips 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trips"
   ClientHeight    =   8205
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1455
      Left            =   7680
      Top             =   2760
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2566
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\J1\Desktop\AirFlights\airFlights.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\J1\Desktop\AirFlights\airFlights.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Trips"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<< First"
      Height          =   495
      Left            =   6480
      TabIndex        =   23
      Top             =   6720
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Trips.frx":0000
      Left            =   3120
      List            =   "Trips.frx":0019
      TabIndex        =   22
      Text            =   "Search By"
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Search"
      Height          =   495
      Left            =   4560
      TabIndex        =   21
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox srch 
      Height          =   495
      Left            =   720
      TabIndex        =   19
      Top             =   7200
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New"
      Height          =   495
      Left            =   6480
      TabIndex        =   16
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox staff 
      DataField       =   "staff"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   6000
      Width           =   2295
   End
   Begin VB.TextBox planetype 
      DataField       =   "planeType"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   6000
      Width           =   2295
   End
   Begin VB.TextBox seatno 
      DataField       =   "seatCount"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   9120
      TabIndex        =   5
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox tripno 
      DataField       =   "tripsNo"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox tripcount 
      DataField       =   "tripsCount"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox dest 
      DataField       =   "destination"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Trip Data"
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   11775
      Begin VB.CommandButton Command10 
         Caption         =   "Exit"
         Height          =   495
         Left            =   9240
         TabIndex        =   28
         Top             =   3360
         Width           =   2295
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Main"
         Height          =   495
         Left            =   6360
         TabIndex        =   27
         Top             =   3360
         Width           =   2415
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Last >>"
         Height          =   495
         Left            =   10560
         TabIndex        =   26
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton Command7 
         Caption         =   "< Previous"
         Height          =   495
         Left            =   9240
         TabIndex        =   25
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Next >"
         Height          =   495
         Left            =   7800
         TabIndex        =   24
         Top             =   2640
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Search"
         Height          =   1215
         Left            =   360
         TabIndex        =   20
         Top             =   2640
         Width           =   5535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   495
         Left            =   10560
         TabIndex        =   17
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Update"
         Height          =   495
         Left            =   9240
         TabIndex        =   15
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton ADD 
         Caption         =   "Add"
         Height          =   495
         Left            =   7800
         TabIndex        =   14
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Staff"
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Plane Type"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Seat Count"
         Height          =   255
         Left            =   9000
         TabIndex        =   11
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Trips NO."
         Height          =   255
         Left            =   6120
         TabIndex        =   10
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Trips Count"
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Destination"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   2295
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Trips.frx":0060
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   24
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "destination"
         Caption         =   "destination"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "tripsCount"
         Caption         =   "tripsCount"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "tripsNo"
         Caption         =   "tripsNo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "seatCount"
         Caption         =   "seatCount"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "planeType"
         Caption         =   "planeType"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "staff"
         Caption         =   "staff"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2775.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1665.071
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2250.142
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Trips Form"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   18
      Top             =   360
      Width           =   11775
   End
End
Attribute VB_Name = "Trips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String

Private Sub ADD_Click()
    On Error Resume Next
    Adodc1.Recordset.Update
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Adodc1.Recordset.Update
End Sub

Private Sub Command10_Click()
    End
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Adodc1.Recordset.AddNew
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    If (MsgBox("Delete it?", vbYesNo, "Delete Record") = vbYes) Then
        Adodc1.Recordset.Delete
    End If
End Sub

Private Sub Command4_Click()
    Adodc1.CommandType = adCmdText
    If (Combo1.Text = "ID") Then
        sql = "select * from Trips where id = " & srch & ""
    Else
        sql = "select * from Trips where " & Combo1.Text & " like '%" & srch & "%'"
    End If
    If (srch <> "") Then
        Adodc1.RecordSource = sql
        If Adodc1.Recordset.EOF Then
            MsgBox ("No Data")
        End If
    Else
        Adodc1.RecordSource = "select * from Trips"
    End If
    Adodc1.Refresh
End Sub

Private Sub Command5_Click()
    On Error Resume Next
    Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command6_Click()
    On Error Resume Next
    Adodc1.Recordset.MoveNext
End Sub

Private Sub Command7_Click()
    On Error Resume Next
    Adodc1.Recordset.MovePrevious
End Sub

Private Sub Command8_Click()
    On Error Resume Next
    Adodc1.Recordset.MoveLast
End Sub

Private Sub Command9_Click()
    Main.Show
    Me.Hide
End Sub

Private Sub Form_Load()
    Adodc1.Visible = False
End Sub
