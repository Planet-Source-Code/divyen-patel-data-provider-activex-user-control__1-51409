VERSION 5.00
Begin VB.UserControl AdoActiveXUserControl 
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3285
   DataSourceBehavior=   1  'vbDataSource
   PropertyPages   =   "AdoActiveXUserControl.ctx":0000
   ScaleHeight     =   495
   ScaleWidth      =   3285
   Begin VB.CommandButton cmdAdoNevigation 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   2760
      TabIndex        =   3
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdAdoNevigation 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2280
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdAdoNevigation 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdAdoNevigation 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Record Navigator"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "AdoActiveXUserControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'-------------------------------------
'set DataSourceBhavior=VBDataSource
'set DataBindingBhavior=VBSimpleBound
'-------------------------------------
Dim C_ConnectionString As String

Dim C_Path As String
Dim DBConnection As New ADODB.Connection
Dim rsSchema As New ADODB.Recordset
Public rs As New ADODB.Recordset


Public Property Get ConnectionString() As String
    ConnectionString = C_ConnectionString
End Property

Public Property Let ConnectionString(ByVal vNewValue As String)
    C_ConnectionString = vNewValue
    If DBConnection.State = adStateOpen Then
        DBConnection.Close
    End If
    DBConnection.Open C_ConnectionString
    
    If rsSchema.State = adStateOpen Then rsSchema.Close
    Set rsSchema = DBConnection.OpenSchema(adSchemaTables)
    
    With DataMembers
        .Clear
    End With
    
    While rsSchema.EOF <> True
        If rsSchema!TABLE_TYPE = "TABLE" Then
            With DataMembers
                .Add rsSchema!TABLE_NAME
            End With
        End If
        rsSchema.MoveNext
    Wend
    
    
    rsSchema.Close
    PropertyChanged "ConnectionString"
End Property

Private Sub cmdAdoNevigation_Click(Index As Integer)
If rs.RecordCount > 0 Then
    If Index = 0 Then
        rs.MoveFirst
    ElseIf Index = 1 Then
        rs.MovePrevious
        If rs.BOF = True Then
            rs.MoveFirst
        End If
    ElseIf Index = 2 Then
            rs.MoveNext
        If rs.EOF = True Then
            rs.MoveLast
        End If
    Else
        rs.MoveLast
    End If
End If
End Sub



Private Sub UserControl_GetDataMember(DataMember As String, Data As Object)
    If rs.State = adStateOpen Then rs.Close
    rs.Open DataMember, DBConnection, adOpenKeyset, adLockOptimistic
    Set Data = rs
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    C_Path = PropBag.ReadProperty("Database")
    Let ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & C_Path & ";Persist Security Info=False"
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ConnectionString", C_ConnectionString
    PropBag.WriteProperty "Database", C_Path
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub


Public Property Let Database(ByVal vNewValue As String)
    C_Path = vNewValue
    If Len(C_Path) <> 0 Then
        Let ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & C_Path & ";Persist Security Info=False"
        PropertyChanged "Database"
    End If
End Property

Public Property Get Database() As String
Attribute Database.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Database = C_Path
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

