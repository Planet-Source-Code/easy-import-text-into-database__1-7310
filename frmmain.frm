VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Easy Database Search"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   495
      Left            =   4800
      TabIndex        =   16
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txtsearchbox 
         Height          =   285
         Left            =   2280
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Enter Part Number to search"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2010
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   5895
      Begin MSMask.MaskEdBox txtretail 
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtlist 
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtdesc 
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox txtss 
         Height          =   285
         Left            =   3840
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtpart 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Superceded to:"
         Height          =   195
         Left            =   2640
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Retail Price"
         Height          =   195
         Left            =   2760
         TabIndex        =   6
         Top             =   960
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "List Price"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Part Number"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   5895
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmmain.frx":0442
         Height          =   1815
         Left            =   120
         OleObjectBlob   =   "frmmain.frx":0456
         TabIndex        =   1
         Top             =   240
         Width           =   5655
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2160
         Width           =   5655
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Dim strMsgBoxRetVal As String
strMsgBoxRetVal = MsgBox("Email: dc_computers@bigpond.com", 65, "About")
End Sub

Private Sub Data1_Reposition()
'Make sure that when Data1 is used to change the records
'that the text boxes reflect the changes
On Error Resume Next

txtpart.Text = Data1.Recordset!part
txtss.Text = Data1.Recordset!ss
txtdesc.Text = Data1.Recordset!Desc
txtlist.Text = Data1.Recordset!List
txtretail.Text = Data1.Recordset!Rec

End Sub

Private Sub Form_Load()
'I like to open mydatabses at run time rather than
'hard code the path in the data control.
'So below it's telling the program that the database
'should be found in the same path as this app.

Set HDesk = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\MyDataBase.MDB", False, False)
Data1.DatabaseName = App.Path & "\MyDataBase.mdb"
Set Data1.Recordset = HDesk.OpenRecordset("MyTable", dbOpenDynaset)

txtpart.Text = Data1.Recordset!part
txtss.Text = Data1.Recordset!ss
txtdesc.Text = Data1.Recordset!Desc
txtlist.Text = Data1.Recordset!List
txtretail.Text = Data1.Recordset!Rec

Data1.Refresh
End Sub

Private Sub Option1_Click()

End Sub

Private Sub txtsearchbox_Change()
'When the user types in a number this code will try
'and find an exact match in the database records
'eg. Type: -1103
'The program will sort all records in the Part Number field
'that start with -1103.
'Now type the rest of the part number so you end up
'with -110383.
'There you go!! There's that part you were looking for.

On Error Resume Next

Dim squote As String
Dim src As String
Dim criteria As String

txtpart.Text = Data1.Recordset!part
txtss.Text = Data1.Recordset!ss
txtdesc = Data1.Recordset!Desc
txtlist = Data1.Recordset!List
txtretail = Data1.Recordset!Rec

squote = Chr$(34)
squotes = Chr$(42)

src = txtsearchbox

If src = "" Then Exit Sub
Data1.RecordSource = "SELECT * FROM mytable WHERE PART Like " & squote & [src] & squotes & squote
Data1.Refresh

End Sub

