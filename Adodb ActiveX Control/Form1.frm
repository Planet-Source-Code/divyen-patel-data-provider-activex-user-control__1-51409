VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Test Form"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add New"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   855
   End
   Begin ADOActiveXControl.AdoActiveXUserControl AdoActiveXUserControl1 
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   3240
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      ConnectionString=   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Adodb ActiveX Control\EmployeeDatabase.mdb;Persist Security Info=False"
      Database        =   "C:\Adodb ActiveX Control\EmployeeDatabase.mdb"
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "Salary"
      DataMember      =   "Employee"
      DataSource      =   "AdoActiveXUserControl1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   2280
      TabIndex        =   8
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "LastName"
      DataMember      =   "Employee"
      DataSource      =   "AdoActiveXUserControl1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   7
      Top             =   1800
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "FirstName"
      DataMember      =   "Employee"
      DataSource      =   "AdoActiveXUserControl1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   1440
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "EmpNumber"
      DataMember      =   "Employee"
      DataSource      =   "AdoActiveXUserControl1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Employee Number"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Specify the ""Employee"" Database Path [ Given in the Folder ]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    AdoActiveXUserControl1.rs.AddNew
    For i = 0 To Text1.Count - 1
        Text1(i).Enabled = True
    Next
    Text1(0).SetFocus
    cmdAdd.Enabled = False
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    AdoActiveXUserControl1.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    With AdoActiveXUserControl1.rs
        .CancelUpdate
        If .RecordCount > 0 Then
            .MoveFirst
        End If
    End With
    cmdAdd.Enabled = True
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    For i = 0 To Text1.Count - 1
        Text1(i).Enabled = False
    Next
    AdoActiveXUserControl1.Enabled = True
End Sub

Private Sub cmdDelete_Click()
    With AdoActiveXUserControl1
    If .rs.RecordCount > 0 Then
            .rs.Delete
            If .rs.RecordCount > 0 Then
                    .rs.MoveNext
                    If .rs.EOF = True Then
                        .rs.MoveLast
                    End If
            Else
                MsgBox "All Record Deleted ...", vbInformation
            End If
    Else
        MsgBox "All Record Deleted ...", vbInformation
    End If
    End With
End Sub

Private Sub cmdSave_Click()
    AdoActiveXUserControl1.rs.Save
    AdoActiveXUserControl1.rs.MoveLast
    cmdAdd.Enabled = True
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    For i = 0 To Text1.Count - 1
        Text1(i).Enabled = False
    Next
    AdoActiveXUserControl1.Enabled = True
End Sub

Private Sub Form_Load()
    For i = 0 To Text1.Count - 1
        Text1(i).Enabled = False
    Next
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
End Sub
