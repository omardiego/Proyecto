VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmIngreso 
   Caption         =   "Ingreso"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleMode       =   0  'User
   ScaleWidth      =   14380.69
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   615
      Left            =   5400
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton Cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   3600
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2280
      Top             =   2760
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.TextBox Txtclave 
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   1560
      Width           =   3615
   End
   Begin MSDataListLib.DataCombo Dcusuario 
      Height          =   315
      Left            =   5760
      TabIndex        =   2
      Top             =   840
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   840
      Width           =   3015
   End
End
Attribute VB_Name = "FrmIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmdaceptar_Click()
Dim s1 As String

Me.Adousuario.Recordset.MoveFirst
Me.Adousuario.Recordset.Find "Nom_Us = '" & Me.Dcusuario.Text & "'"
s1 = Me.Adousuario.Recordset("clave").Value
If Me.Txtclave.Text <> "" Then
    If s1 = Me.Txtclave.Text Then
Unload Me
Form2.Show
Else
MsgBox " clave mal ingresada", vbInformation + vbOKOnly, "Software Educativo"
End If
Else
MsgBox "digite la clave por favor", vbInformation + vbOKOnly, "Software Educativo"
End If

End Sub

Private Sub Form_Load()
Me.Adousuario.Refresh
End Sub
