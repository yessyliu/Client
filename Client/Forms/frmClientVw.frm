VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmClientView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clients"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid dgClients 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   5106
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "Details"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCreateNew 
      Caption         =   "Create New"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
End
Attribute VB_Name = "frmClientView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FormStatus As Integer
Dim rsData As ADODB.Recordset
Dim objClient As clsClient

Private Sub cmdCreateNew_Click()
    frmClient.Status = "NEW"
    frmClient.Show
End Sub

Private Sub cmdDelete_Click()
    If dgClients.Row >= 0 Then
        If MsgBox("Are you sure want to delete this record? " & vbCrLf & "Name : " & dgClients.Columns(1).Text & vbCrLf _
                  & "Gender : " & IIf(Left(dgClients.Columns(3).Text, 1) = "M", "Male", "Female") & vbCrLf _
                  & "Date of Birth : " & Format(dgClients.Columns(2).Text, "d MMM yyyy"), vbQuestion + vbYesNo + vbDefaultButton1, "Save") = vbYes Then
            objClient.deleteDataClient (dgClients.Columns(0).Text)
            FormStatus = 0
            Call showClient
        End If
    End If
End Sub

Private Sub cmdDetails_Click()
    If dgClients.Row >= 0 Then
        FormStatus = 1
        Call ChangeState
        frmClient.Status = "VIEW"
        frmClient.ClientID = dgClients.Columns(0).Text
        frmClient.Show
    End If
End Sub

Private Sub cmdEdit_Click()
    If dgClients.Row >= 0 Then
        FormStatus = 1
        Call ChangeState
        frmClient.Status = "EDIT"
        frmClient.ClientID = dgClients.Columns(0).Text
        frmClient.Show
    End If
End Sub

Private Sub dgClients_Click()
    If dgClients.Row >= 0 Then
        FormStatus = 1
        Call ChangeState
    End If
End Sub

Private Sub Form_Load()
    FormStatus = 0
    Call ChangeState
    Call showClient
End Sub

Private Sub ChangeState()
    If FormStatus = 1 Then 'Form saat pilih row
        cmdCreateNew.Enabled = True
        cmdEdit.Enabled = True
        cmdDetails.Enabled = True
        cmdDelete.Enabled = True
    ElseIf FormStatus = 0 Then 'Form Awal
        cmdCreateNew.Enabled = True
        cmdEdit.Enabled = False
        cmdDetails.Enabled = False
        cmdDelete.Enabled = False
    End If
End Sub

Public Function showClient()
    Set rsData = New ADODB.Recordset
    Set objClient = New clsClient
    Set rsData = objClient.showClient
    If Not rsData Is Nothing Then
        Set dgClients.DataSource = rsData
        dgClients.Refresh
        dgClients.Columns(0).Visible = False
        dgClients.Columns(2).NumberFormat = "d MMM yyyy"
        
        Dim i As Integer
        i = 1
        While (i < dgClients.Columns.Count)
            dgClients.Columns(i).Locked = True
            i = i + 1
        Wend
    Else
        Set dgClients.DataSource = Nothing
        dgClients.Refresh
    End If
    Call ChangeState
End Function
