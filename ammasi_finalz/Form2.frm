VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   525
   ClientLeft      =   5685
   ClientTop       =   6150
   ClientWidth     =   8910
   LinkTopic       =   "Form2"
   ScaleHeight     =   525
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   120
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = True
Timer1_Timer

End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 10
    If ProgressBar1.Value = 80 Then
    ProgressBar1.Value = ProgressBar1 + 20
    
    If ProgressBar1.Value >= ProgressBar1.Max Then
        Timer1.Enabled = False
        End If
        MsgBox "Welcome User!", vbInformation, "Short Info"
        Unload Me
        form3.Show
        
    End If
        
        
    
End Sub
