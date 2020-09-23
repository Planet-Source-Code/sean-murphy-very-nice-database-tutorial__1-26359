VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Address Book"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Data databar 
      Caption         =   "Last | Previous                             Next | Last"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\Desktop\address.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "address"
      Top             =   3720
      Width           =   4455
   End
   Begin VB.TextBox txtComments 
      DataField       =   "comments"
      DataSource      =   "databar"
      Height          =   855
      Left            =   960
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtEmail 
      DataField       =   "email"
      DataSource      =   "databar"
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtPhone 
      DataField       =   "phone"
      DataSource      =   "databar"
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtAddress 
      DataField       =   "address"
      DataSource      =   "databar"
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtName 
      DataField       =   "name"
      DataSource      =   "databar"
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtLastName 
      DataField       =   "LastName"
      DataSource      =   "databar"
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "E-mail:"
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Comments:"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Address:"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Phone:"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "First Name:"
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Last Name:"
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
databar.Recordset.Delete
databar.Recordset.MovePrevious
End Sub

Private Sub Command2_Click()
databar.Recordset.Edit
End Sub

Private Sub Command3_Click()
databar.Recordset.AddNew
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Text2_Change()

End Sub

