VERSION 5.00
Begin VB.Form FrmImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Ascii File"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3690
   Icon            =   "FrmImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   3690
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton Command2 
         Caption         =   "Open DataBase"
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start Import"
         Height          =   495
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.ListBox List1 
         Height          =   1035
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim myfile As String
Dim spart As String
Dim sdesc As String
Dim sssnr As String
Dim srec As String
Dim sslist As String
Dim ssretail As String
    
Set HDesk = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\MyDataBase.MDB", False, False)
Data1.DatabaseName = App.Path & "\MyDataBase.mdb"
Set Data1.Recordset = HDesk.OpenRecordset("MyTable", dbOpenDynaset)
     
'Open the ASCII file and sort it's data
If Right(App.Path, 1) <> "\" Then: myfile = App.Path + "\Prices.dat": Else: myfile = App.Path + "Prices.dat"
Open myfile For Input As #1 Len = 64
  
Do While Not EOF(1)

'Start the Input of files
Input #1, srec
'_________________________________________________________
'   ****Change this area if price file structure changes
'________________________________________________________
spart = Mid$(srec, 1, 14)
sdesc = Mid$(srec, 15, 17)
sssnr = Mid$(srec, 32, 14)
sslist = Mid$(srec, 46, 8)
ssretail = Mid$(srec, 54, 8)
'________________________________________________________

Data1.Recordset.AddNew

Data1.Recordset!part = spart
Data1.Recordset!Desc = sdesc
Data1.Recordset!ss = sssnr
Data1.Recordset!List = sslist
Data1.Recordset!Rec = ssretail
Data1.Recordset.Update
List1.AddItem spart


Loop
   
Close #1
MsgBox List1.ListCount & " files Imported Successfully"


End Sub

Private Sub Command2_Click()
FrmMain.Show
End Sub

