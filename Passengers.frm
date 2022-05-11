VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Passengers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Passengers"
   ClientHeight    =   9015
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Passengers.frx":0000
      Height          =   1215
      Left            =   120
      TabIndex        =   19
      Top             =   3720
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   2143
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
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
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1904.882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1649.764
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1785.26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2039.811
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   975
      Left            =   7440
      Top             =   2640
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1720
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
      RecordSource    =   "Passengers"
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
   Begin VB.ComboBox Combo1 
      DataField       =   "trip_id"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "Passengers.frx":0015
      Left            =   9600
      List            =   "Passengers.frx":0017
      TabIndex        =   17
      Text            =   "Trip ID"
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New"
      Height          =   495
      Left            =   9720
      TabIndex        =   16
      Top             =   6600
      Width           =   975
   End
   Begin VB.ComboBox Degree 
      DataField       =   "degree"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "Passengers.frx":0019
      Left            =   480
      List            =   "Passengers.frx":0023
      TabIndex        =   7
      Text            =   "Degree"
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox fullname 
      DataField       =   "fullname"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   5880
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Passengers.frx":0035
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   10
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
         DataField       =   "fullname"
         Caption         =   "fullname"
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
         DataField       =   "gender"
         Caption         =   "gender"
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
         DataField       =   "age"
         Caption         =   "age"
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
         DataField       =   "price"
         Caption         =   "price"
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
         DataField       =   "degree"
         Caption         =   "degree"
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
         DataField       =   "passID"
         Caption         =   "passID"
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
      BeginProperty Column07 
         DataField       =   "validDate"
         Caption         =   "validDate"
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
      BeginProperty Column08 
         DataField       =   "seatNo"
         Caption         =   "seatNo"
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
      BeginProperty Column09 
         DataField       =   "trip_id"
         Caption         =   "trip_id"
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
            ColumnWidth     =   374.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   12135
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   5280
         TabIndex        =   31
         Top             =   2640
         Width           =   6855
         Begin VB.CommandButton Command11 
            Caption         =   "Exit"
            Height          =   495
            Left            =   5640
            TabIndex        =   37
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Main"
            Height          =   495
            Left            =   4560
            TabIndex        =   36
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Last >>"
            Height          =   495
            Left            =   3000
            TabIndex        =   35
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton Command8 
            Caption         =   "< Previous"
            Height          =   495
            Left            =   2040
            TabIndex        =   34
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Next >"
            Height          =   495
            Left            =   1080
            TabIndex        =   33
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton Command6 
            Caption         =   "<< First"
            Height          =   495
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Search"
         Height          =   1095
         Left            =   0
         TabIndex        =   25
         Top             =   2640
         Width           =   5175
         Begin VB.ComboBox by 
            Height          =   315
            ItemData        =   "Passengers.frx":004A
            Left            =   2280
            List            =   "Passengers.frx":006C
            TabIndex        =   30
            Text            =   "By"
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Search"
            Height          =   495
            Left            =   3840
            TabIndex        =   29
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox srch 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "M/d/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label10 
            Caption         =   "By"
            Height          =   255
            Left            =   2280
            TabIndex        =   27
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Delete"
         Height          =   495
         Left            =   10800
         TabIndex        =   24
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Update"
         Height          =   495
         Left            =   9600
         TabIndex        =   23
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Text2 
         DataField       =   "validDate"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   4920
         TabIndex        =   22
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         DataField       =   "price"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   7200
         TabIndex        =   15
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   495
         Left            =   10800
         TabIndex        =   14
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox seatNo 
         DataField       =   "seatNo"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   7200
         TabIndex        =   6
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox PassID 
         DataField       =   "passID"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2640
         TabIndex        =   5
         Top             =   2040
         Width           =   1935
      End
      Begin VB.ComboBox gender 
         DataField       =   "gender"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Passengers.frx":00BD
         Left            =   2640
         List            =   "Passengers.frx":00C7
         TabIndex        =   4
         Text            =   "Gender"
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox age 
         DataField       =   "age"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   4920
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Gender"
         Height          =   255
         Left            =   2640
         TabIndex        =   28
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Degree"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Trip ID"
         Height          =   255
         Left            =   9480
         TabIndex        =   18
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Price"
         Height          =   255
         Left            =   7200
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Seat No."
         Height          =   255
         Left            =   7200
         TabIndex        =   12
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Valid Date"
         Height          =   255
         Left            =   4920
         TabIndex        =   11
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Passport ID"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Age"
         Height          =   255
         Left            =   4920
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Fullname"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   1095
      Left            =   1440
      Top             =   4080
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1931
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
      Caption         =   "Adodc2"
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Passengers Form"
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
      Left            =   240
      TabIndex        =   20
      Top             =   240
      Width           =   11775
   End
End
Attribute VB_Name = "Passengers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_click()
    If (Combo1.Text <> "") Then
        Adodc2.RecordSource = "select * from trips where id = " & Combo1.Text & ""
        Adodc2.Refresh
        DataGrid2.Refresh
    End If
End Sub

Private Sub Combo1_change()
    If (Combo1.Text <> "") Then
        Adodc2.RecordSource = "select * from trips where id = " & Combo1.Text & ""
        If (Adodc2.Recordset.EOF) Then
            Adodc2.Refresh
            DataGrid2.Refresh
        End If
    End If
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Adodc1.Recordset.Update
End Sub

Private Sub Command10_Click()
    Main.Show
    Me.Hide
End Sub

Private Sub Command11_Click()
    End
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Adodc1.Recordset.AddNew
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    Adodc1.Recordset.Update
End Sub

Private Sub Command4_Click()
    On Error Resume Next
    If (MsgBox("Delete it?", vbYesNo, "Delete Record") = vbYes) Then
        Adodc1.Recordset.Delete
    End If
End Sub

Private Sub Command5_Click()
    On Error Resume Next
    Adodc1.CommandType = adCmdText
    If (by.Text = "ID" Or by.Text = "age" Or by.Text = "tripID") Then
        sql = "select * from passengers where " & by.Text & " = " & srch & ""
    Else
        sql = "select * from passengers where " & by.Text & " like '%" & srch & "%'"
    End If
    If (srch <> "") Then
        Adodc1.RecordSource = sql
        If Adodc1.Recordset.EOF Then
            MsgBox ("No Data")
        End If
    Else
        Adodc1.RecordSource = "select * from passengers"
    End If
    Adodc1.Refresh
End Sub

Private Sub Command6_Click()
    On Error Resume Next
    Adodc1.Recordset.MoveFirst
End Sub

Private Sub Command7_Click()
    On Error Resume Next
    Adodc1.Recordset.MoveNext
End Sub

Private Sub Command8_Click()
    On Error Resume Next
    Adodc1.Recordset.MovePrevious
End Sub

Private Sub Command9_Click()
    On Error Resume Next
    Adodc1.Recordset.MoveLast
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from trips"
    Adodc2.Refresh
    Adodc2.Recordset.MoveFirst
    With Adodc2.Recordset
        Do Until .EOF
            Combo1.AddItem ![id]
            .MoveNext
        Loop
        If (Adodc2.Recordset.EOF) Then
        Else
            Combo1.Text = Adodc1.Recordset.Fields("trip_id")
        End If
    End With
    
End Sub
