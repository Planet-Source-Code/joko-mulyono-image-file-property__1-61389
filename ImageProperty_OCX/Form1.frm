VERSION 5.00
Object = "*\AIMGProperty.vbp"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "IMAGE PROPERTY"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin IMGProperty.ImageProperty ImageProperty1 
      Height          =   3600
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   6350
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ImageProperty1.FileName = App.Path & "\WallpaperChanger.jpg"
ImageProperty1.CariInfoGambar
End Sub
