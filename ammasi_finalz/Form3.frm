VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form form3 
   BackColor       =   &H0080FF80&
   Caption         =   "Form3"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15915
   LinkTopic       =   "Form3"
   ScaleHeight     =   7440
   ScaleWidth      =   15915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton bedit 
      BackColor       =   &H0080C0FF&
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6240
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   13920
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
      BackColor       =   8438015
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\user.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\user.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Inventory"
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
   Begin VB.CommandButton bdelete 
      BackColor       =   &H0080C0FF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton bsearch 
      BackColor       =   &H0080C0FF&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton bsave 
      BackColor       =   &H0080C0FF&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton badd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Add Item"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton blast 
      BackColor       =   &H0080C0FF&
      Caption         =   ">|"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton bnext 
      BackColor       =   &H0080C0FF&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton bprevious 
      BackColor       =   &H0080C0FF&
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton bfirst 
      BackColor       =   &H0080C0FF&
      Caption         =   "|<"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox tfstock 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      DataField       =   "stocks"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7800
      TabIndex        =   4
      Top             =   3120
      Width           =   4215
   End
   Begin VB.TextBox tfclass 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      DataField       =   "class"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7800
      TabIndex        =   3
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   " Stocks Available"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   2160
      TabIndex        =   2
      Top             =   3360
      Width           =   4575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Classification"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   2520
      TabIndex        =   1
      Top             =   2040
      Width           =   3780
   End
   Begin VB.Line Line1 
      X1              =   4080
      X2              =   11520
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Vegetable Inventory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   5160
      TabIndex        =   0
      Top             =   360
      Width           =   5280
   End
End
Attribute VB_Name = "form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub badd_Click()


Adodc1.Recordset.AddNew
tfclass.SetFocus

End Sub

Private Sub bdelete_Click()
Dim ans As String
ans = MsgBox("Are you sure you want to delete this vegetable", vbYesNo + vbQuestion, "Delete Item") '
    If ans = vbYes Then
    Adodc1.Recordset.Delete
    Adodc1.Refresh
    End If
    
    
    
End Sub

Private Sub bedit_Click()
 If Not Adodc1.Recordset.EOF Then
 tfclass.SetFocus
 Else
    MsgBox "No Record to edit", vbExclamation
    End If
    


End Sub

Private Sub bfirst_Click()
Adodc1.Recordset.MoveFirst

End Sub



Private Sub blast_Click()
Adodc1.Recordset.MoveLast

End Sub

Private Sub bnext_Click()
Adodc1.Recordset.MoveNext
End Sub

Private Sub bprevious_Click()
Adodc1.Recordset.MovePrevious
End Sub

Private Sub bsave_Click()
Dim ans As String
ans = MsgBox("Would You Like to Add Item?", vbYesNo + vbQuestion, "save")
    If ans = vbYes Then
    Adodc1.Recordset.Fields("class") = tfclass.Text
    Adodc1.Recordset.Fields("stocks") = tfstock.Text
    Adodc1.Recordset.Update
    MsgBox "Item Added Successfully", vbInformation, "Saved!"
    End If
    
End Sub

Private Sub bsearch_Click()
Dim klass As String
klass = InputBox("Search The Item: ", "Search Item")
Adodc1.Refresh
Do Until Adodc1.Recordset.EOF
    If Adodc1.Recordset.Fields!Class = klass Then
    tfclass.Text = Adodc1.Recordset.Fields("class").Value
    tfstock.Text = Adodc1.Recordset.Fields("stocks").Value
    
    Exit Sub
    

    Else
        Adodc1.Recordset.MoveNext
    End If
    
Loop
    MsgBox " No Vegetable Available", vbCritical, "No Item"


End Sub
