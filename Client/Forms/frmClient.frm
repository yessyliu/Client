VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtpDOB 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   1995
      TabIndex        =   5
      Top             =   1250
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   109379587
      CurrentDate     =   43594
   End
   Begin VB.ComboBox cboGender 
      Height          =   315
      ItemData        =   "frmClient.frx":0000
      Left            =   2000
      List            =   "frmClient.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   800
      Width           =   3015
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   2000
      MaxLength       =   100
      TabIndex        =   3
      Top             =   320
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Date of Birth"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Gender"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Full Name"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Status As String
Dim objClient As clsClient
Dim rsData As ADODB.Recordset
Public ClientID As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Set objClient = New clsClient
    If Status = "NEW" Then
        Call objClient.saveDataClient(txtName.text, cboGender.text, dtpDOB.Value, Status)
    ElseIf Status = "EDIT" Then
        Call objClient.saveDataClient(txtName.text, cboGender.text, dtpDOB.Value, Status, ClientID)
    End If
    Unload Me
    frmClientView.FormStatus = 0
    frmClientView.showClient
    frmClientView.Show
End Sub

Private Sub Form_Load()
    cboGender.AddItem ("Male")
    cboGender.AddItem ("Female")
    
    If Status = "NEW" Then
        txtName.text = ""
        cboGender.ListIndex = 0
    ElseIf Status = "VIEW" Then
        txtName.Locked = True
        dtpDOB.Enabled = False
        cboGender.Locked = True
        cmdSave.Caption = "Back"
        cmdClose.Visible = False
        Call GetData
    ElseIf Status = "EDIT" Then
        Call GetData
    End If
End Sub

Private Sub GetData()
    Set objClient = New clsClient
    Set rsData = objClient.getDataClient(ClientID)
    
        txtName.text = rsData!Name
        dtpDOB.Value = rsData!DOB
        If rsData!Gender = "M" Then
            cboGender.ListIndex = 0
        Else
            cboGender.ListIndex = 1
        End If
End Sub
