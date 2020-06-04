VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Ekspor Excel"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc ado 
      Height          =   330
      Left            =   1313
      Top             =   2640
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "File yang sudah ada (template)"
      Height          =   555
      Left            =   2663
      TabIndex        =   1
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "File Baru"
      Height          =   555
      Left            =   480
      TabIndex        =   0
      Top             =   3120
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2295
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error GoTo Ero
    
    Dim xlApp As New Excel.Application
    
    With xlApp
    
    .Workbooks.Add
    
    'judul
    .Range("A1").Value = "Buku Alamat"
    .Range("A1").Select
    .Selection.Font.Bold = True
    .Selection.Font.Size = 16
    
    'kolom
    .Range("A2").Value = "ID"
    .Range("B2").Value = "Nama"
    .Range("C2").Value = "Alamat"
    .Range("A2:C2").Select
    .Selection.Font.Bold = True
    .Selection.HorizontalAlignment = xlCenter
    
    'data
    ado.Recordset.MoveFirst
    For i = 1 To ado.Recordset.RecordCount
        .Range("A" & CStr(i + 2)).Value = ado.Recordset!ID
        .Range("B" & CStr(i + 2)).Value = ado.Recordset!Nama
        .Range("C" & CStr(i + 2)).Value = ado.Recordset!Alamat
        ado.Recordset.MoveNext
    Next
    
    'membuat list
    xlApp.ActiveSheet.ListObjects.Add , xlApp.Range("A2:C" & CStr(ado.Recordset.RecordCount + 2)), , xlYes
    
    .Range("A1").Select
    .Visible = True
    
    End With
    
    Exit Sub
    
Ero:
    MsgBox Err.Description
    xlApp.ActiveWorkbook.Close False
End Sub

Private Sub Command2_Click()
    On Error GoTo Ero
    
    Dim xlApp As New Excel.Application
    
    With xlApp
    
    .Workbooks.Open App.Path & "\alamat.xlt"
    
    'memasukkan baris/row, untuk mengcopy border
    .Rows("4:4").Select
    For i = 3 To ado.Recordset.RecordCount
        .Selection.Insert Shift:=xlDown
    Next

    'data
    ado.Recordset.MoveFirst
    For i = 1 To ado.Recordset.RecordCount
        .Range("A" & CStr(i + 2)).Value = ado.Recordset!ID
        .Range("B" & CStr(i + 2)).Value = ado.Recordset!Nama
        .Range("C" & CStr(i + 2)).Value = ado.Recordset!Alamat
        ado.Recordset.MoveNext
    Next
    
    .Range("A1").Select
    .Visible = True
    
    End With
    
    Exit Sub
    
Ero:
    MsgBox Err.Description
    xlApp.ActiveWorkbook.Close False
End Sub


Private Sub Form_Load()
    ado.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\alamat.mdb;"
    ado.RecordSource = "select * from tblAlamat"
    ado.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ado.Recordset.Close
End Sub
