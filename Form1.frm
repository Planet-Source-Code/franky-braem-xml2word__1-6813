VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute"
      Enabled         =   0   'False
      Height          =   420
      Left            =   1410
      TabIndex        =   7
      Top             =   1020
      Width           =   2160
   End
   Begin VB.CommandButton cmdOpenTemplate 
      Caption         =   "..."
      Height          =   285
      Left            =   3585
      TabIndex        =   6
      Top             =   465
      Width           =   390
   End
   Begin VB.CommandButton cmdOpenXMLFile 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3585
      TabIndex        =   5
      Top             =   90
      Width           =   390
   End
   Begin VB.TextBox txtWordTemplate 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   465
      Width           =   2115
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   1500
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Text            =   "Ready."
            TextSave        =   "Ready."
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlgOpen 
      Left            =   180
      Top             =   1125
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtXMLFile 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   2115
   End
   Begin VB.Label Label2 
      Caption         =   "Word Template"
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   525
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "XML File"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExecute_Click()
    
    Dim oXML2Word As XML2Word.clMain
    
    sbStatus.Panels.Item(1).Text = "Busy ..."
    
    Set oXML2Word = New XML2Word.clMain
    
    oXML2Word.WordTemplate = txtWordTemplate.Text
    oXML2Word.XMLFile = txtXMLFile.Text
    oXML2Word.Execute
        
    Set oXML2Word = Nothing
    
    sbStatus.Panels.Item(1).Text = "Ready."

End Sub

Private Sub cmdOpenTemplate_Click()
    
    cdlgOpen.ShowOpen
    txtWordTemplate.Text = cdlgOpen.FileName

End Sub

Private Sub cmdOpenXMLFile_Click()
    
    cdlgOpen.ShowOpen
    txtXMLFile.Text = cdlgOpen.FileName
    
End Sub

Private Sub txtWordTemplate_Change()

    If txtWordTemplate.Text <> "" And txtXMLFile.Text <> "" Then
        cmdExecute.Enabled = True
    Else
        cmdExecute.Enabled = False
    End If
    
End Sub

Private Sub txtXMLFile_Change()
    
    If txtWordTemplate.Text <> "" And txtXMLFile.Text <> "" Then
        cmdExecute.Enabled = True
    Else
        cmdExecute.Enabled = False
    End If

End Sub
