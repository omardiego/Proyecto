VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FrmPaginas 
   Caption         =   "Paginas Web"
   ClientHeight    =   9840
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   ScaleHeight     =   9840
   ScaleWidth      =   14985
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   12535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   18570
      ExtentX         =   32755
      ExtentY         =   22110
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "FrmPaginas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.WebBrowser1.Navigate "D:\Mantenimiento\principal.html"
End Sub
