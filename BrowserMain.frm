VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form BrowserMain 
   Caption         =   "Database Browser"
   ClientHeight    =   7845
   ClientLeft      =   6105
   ClientTop       =   2730
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   Picture         =   "BrowserMain.frx":0000
   ScaleHeight     =   7845
   ScaleWidth      =   6120
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   495
      Left            =   4560
      TabIndex        =   25
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Height          =   495
      Left            =   2280
      TabIndex        =   24
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "DataBase"
      Height          =   3135
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   5775
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   5415
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3000
         TabIndex        =   17
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CommandButton cmdDelTable 
         Caption         =   "Delete Table"
         Height          =   495
         Left            =   720
         TabIndex        =   16
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelField 
         Caption         =   "Delete Field"
         Height          =   495
         Left            =   3600
         TabIndex        =   15
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Database Name"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Table"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Fields"
         Height          =   255
         Left            =   3000
         TabIndex        =   21
         Top             =   1560
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Queries"
      Height          =   1815
      Left            =   3120
      TabIndex        =   10
      Top             =   3480
      Width           =   2775
      Begin VB.CommandButton cmdDelQU 
         Caption         =   "Delete Query"
         Height          =   495
         Left            =   720
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "You can only Delete a Query"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   2535
      End
   End
   Begin VB.OptionButton Option 
      Caption         =   "Field is String"
      Height          =   495
      Index           =   2
      Left            =   2760
      TabIndex        =   9
      Top             =   6960
      Width           =   1455
   End
   Begin VB.OptionButton Option 
      Caption         =   "Field is Numeric"
      Height          =   495
      Index           =   1
      Left            =   2760
      TabIndex        =   8
      Top             =   6360
      Width           =   1455
   End
   Begin VB.OptionButton Option 
      Caption         =   "Field is Date"
      Height          =   495
      Index           =   0
      Left            =   2760
      TabIndex        =   7
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "New Table"
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   2655
      Begin VB.CommandButton cmdAddTable 
         Caption         =   "Add Table"
         Height          =   495
         Left            =   720
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtTableName 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "New Field"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   4215
      Begin VB.CommandButton cmdAddField 
         Caption         =   "Add Field"
         Height          =   495
         Left            =   720
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtFieldName 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   7080
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6000
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "BrowserMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DbName As String
Public pathname As String


Private Sub cmdAbout_Click()

Load frmAbout
frmAbout.Show
Me.Hide

End Sub

Private Sub cmdAddField_Click()
'This function will show how to add fields

'Dim our variables
Dim DB As Database
Dim WS As Workspace
Dim TD As TableDef
Dim FD As Field
Dim LC As Integer

'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
'this opens the database
Set DB = WS.OpenDatabase(DbName)
    
'Set the table to open
Set TD = DB.TableDefs(Combo1.Text)

'Once again, you can use dbText, or dbInterger
'or whatever else you wish to set the field type
If BrowserMain.Option(0).Value = True Then
    Set FD = TD.CreateField(txtFieldName.Text, dbDate)
    Else
    If BrowserMain.Option(1).Value = True Then
        Set FD = TD.CreateField(txtFieldName.Text, dbInteger)
        Else
        Set FD = TD.CreateField(txtFieldName.Text, dbText)
    End If
End If

'Bind field to the table
TD.Fields.Append FD

'close the database
DB.Close

Combo2.AddItem txtFieldName.Text
LC = Combo2.ListCount
Combo2.ListIndex = LC - 1
txtFieldName.Text = ""

End Sub

Private Sub cmdAddTable_Click()
'This function will show how to add tables

'Dim our variables
Dim DB As Database
Dim WS As Workspace
Dim TD As TableDef
Dim FD As Field
Dim ANS As String
Dim LC As Integer
ANS = ""
If txtFieldName.Text = "" Then ANS = MsgBox("You must fill in a field name first", vbOKOnly)
If ANS <> "" Then Exit Sub

'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
'this opens the database
Set DB = WS.OpenDatabase(DbName)
    
'Set the table info
Set TD = DB.CreateTableDef(txtTableName.Text)

'Create field and type
If BrowserMain.Option(0).Value = True Then
    Set FD = TD.CreateField(txtFieldName.Text, dbDate)
    Else
    If BrowserMain.Option(1).Value = True Then
        Set FD = TD.CreateField(txtFieldName.Text, dbInteger)
        Else
        Set FD = TD.CreateField(txtFieldName.Text, dbText)
    End If
End If
'Now bind the table to the database
TD.Fields.Append FD
DB.TableDefs.Append TD

'close the database
DB.Close

Combo1.AddItem txtTableName.Text
LC = Combo1.ListCount
Combo1.ListIndex = LC - 1
list_Fields
txtTableName.Text = ""
txtFieldName.Text = ""

End Sub

Private Sub cmdBrowse_Click()

Combo1.Clear
CommonDialog1.ShowOpen
DbName = CommonDialog1.FileName
If DbName = "" Then Exit Sub
list_Tables
Text1.Text = DbName
list_Queries

End Sub

Public Function list_Tables()
'This function will show how to list tables in a database

'Dim our variables
Dim DB As Database
Dim WS As Workspace
Dim TD As TableDef
Dim Temp As String
Dim Max As Long

Combo1.Clear
'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
'this opens the database
Set DB = WS.OpenDatabase(DbName)
    
'find the number of tables in the database
Max = DB.TableDefs.Count

For i = 0 To Max - 1
'Take away six from max because the last 5
'are not for you to mess with

'select the table
Set TD = DB.TableDefs(i)

'make sure table isn't an untouchable
If Left(TD.Name, 4) = "MSys" Then GoTo Skip

'List the tables in the listbox
Combo1.AddItem TD.Name

Skip:
Next i

'close the database
DB.Close

Combo1.ListIndex = 0

End Function

Private Sub cmdCreate_Click()
Dim DB As Database
Dim WS As Workspace
Dim TD  As TableDef
Dim FD As Field
Dim ANS As String
Dim pathname As String

'Check to see if name is blank
If Text1.Text = "" Then
    ANS = MsgBox("You must Fill in a Database Name", vbOKOnly)
    Exit Sub
End If

'if name isn't blank make sure it has DB suffix
If Right(Text1.Text, 4) <> ".mdb" Then Text1.Text = Text1.Text & ".mdb"

'Make sure there is a \ in front
If Left(Text1.Text, 1) <> "\" Then Text1.Text = "\" & Text1.Text

'Make sure that we will not have \\
If Right(App.Path, 1) = "\" Then
'get the path
pathname = Left(App.Path, Len(App.Path) - 1)
Else
pathname = App.Path
End If

'add the filename
pathname = pathname & Text1.Text
DbName = pathname
'make sure table name isn't blank
If txtTableName.Text = "" Then
    ANS = MsgBox("You must fill in a table name", vbOKOnly)
    Exit Sub
End If

'make sure field name isn't blank
If txtFieldName.Text = "" Then
    ANS = MsgBox("You must fill in a field name", vbOKOnly)
    Exit Sub
End If

'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)

'this Creates the database
Set DB = DBEngine.Workspaces(0).CreateDatabase(pathname, dbLangGeneral)

'Create table for readings
Set TD = DB.CreateTableDef(txtTableName.Text)

'Create field and type
If BrowserMain.Option(0).Value = True Then
    Set FD = TD.CreateField(BrowserMain.txtFieldName.Text, dbDate)
    Else
        If BrowserMain.Option(1).Value = True Then
        Set FD = TD.CreateField(BrowserMain.txtFieldName.Text, dbInteger)
        Else
        Set FD = TD.CreateField(BrowserMain.txtFieldName.Text, dbText)
        End If
End If
'Now bind the field to the table
TD.Fields.Append FD

'And the table to the database
DB.TableDefs.Append TD

'close the database
DB.Close

'add and clear
Combo1.AddItem txtTableName.Text
Combo1.ListIndex = 0
Combo2.AddItem txtFieldName.Text
Combo2.ListIndex = 0
Text1.Text = pathname
txtTableName.Text = ""
txtFieldName.Text = ""

End Sub

Private Sub cmdDelField_Click()
'This function will show how to delete records

'Dim our variables
Dim DB As Database
Dim WS As Workspace
Dim TD As TableDef
Dim FD As Field

'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
'this opens the database
Set DB = WS.OpenDatabase(DbName)
    
'Set the table to open
Set TD = DB.TableDefs(Combo1.Text)

'Erase the field
TD.Fields.Delete Combo2.Text

'close the database
DB.Close

Combo2.Clear
list_Fields

End Sub

Private Sub cmdDelQU_Click()
'This function will show how to delete Queries

'Dim our variables
Dim DB As Database
Dim WS As Workspace
Dim TD As TableDef
Dim QU As QueryDef

'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
'this opens the database
Set DB = WS.OpenDatabase(DbName)
    
'Set the table to open
DB.QueryDefs.Delete Combo3.Text

'close the database
DB.Close
Combo3.Clear

list_Queries

End Sub

Private Sub cmdDelTable_Click()
'This function will show how to delete records

'Dim our variables
Dim DB As Database
Dim WS As Workspace
Dim TD As TableDef
Dim FD As Field

'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
'this opens the database
Set DB = WS.OpenDatabase(DbName)
    
'Set the table to open
DB.TableDefs.Delete Combo1.Text

'close the database
DB.Close
Combo1.Clear
Combo2.Clear

list_Tables
list_Fields

End Sub

Private Sub cmdExit_Click()

End

End Sub

Private Sub Combo1_Click()

Combo2.Clear
list_Fields

End Sub

Public Function list_Fields()

'This function will show how to list Fields in a Table

'Dim our variables
Dim DB As Database
Dim WS As Workspace
Dim TD As TableDef
Dim FD As Field
Dim Temp As String
Dim Max As Long
On Error GoTo errhand
'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
'this opens the database
Set DB = WS.OpenDatabase(DbName)
Set TD = DB.TableDefs(Combo1.Text)
'find the number of fields in the table
Max = TD.Fields.Count

For i = 0 To Max - 1

'select the table
Set FD = TD.Fields(i)

'List the fields in the combobox
Combo2.AddItem FD.Name

Next i

Combo2.ListIndex = 0
'close the database
DB.Close
Exit Function

errhand:
DB.Close

End Function


Private Sub Form_Load()

BrowserMain.Option(2).Value = True

End Sub

Public Function list_Queries()
'This function will show how to list tables in a database

'Dim our variables
Dim DB As Database
Dim WS As Workspace
Dim TD As TableDef
Dim QU As QueryDef
Dim Temp As String
Dim Max As Long

'This sets a workspace for the database
Set WS = DBEngine.Workspaces(0)
'this opens the database
Set DB = WS.OpenDatabase(DbName)
    
'find the number of tables in the database
Max = DB.QueryDefs.Count
If Val(Max) = 0 Then Exit Function

For i = 0 To Max - 1

'select the table
Set QU = DB.QueryDefs(i)
If Left(QU.Name, 4) = "~sq_" Then GoTo Skip

'List the tables in the listbox
Combo3.AddItem QU.Name
Combo3.ListIndex = 0

Skip:
Next i

'close the database
DB.Close

End Function

