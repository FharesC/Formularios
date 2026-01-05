VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   13500
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   4680
      TabIndex        =   2
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Menu mnu_menu 
      Caption         =   "menu1"
      Index           =   1
      Begin VB.Menu mnu_submenu1 
         Caption         =   "submenu1"
         Index           =   4
      End
      Begin VB.Menu mnu_linea 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_submenu2 
         Caption         =   "submenu2"
         Index           =   5
      End
   End
   Begin VB.Menu mnu_menu2 
      Caption         =   "menu2"
      Index           =   2
   End
   Begin VB.Menu mnu_menu3 
      Caption         =   "menu3"
      Index           =   3
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.valor = Text1.Text
Form2.Show
End Sub

Private Sub Command2_Click()
valor2 = Text1.Text
End Sub

Private Sub mnu_submenu1_Click(Index As Integer)
Form2.valor = "mensaje"
Form2.Show vbModal
End Sub
