VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About this Program"
   ClientHeight    =   4830
   ClientLeft      =   14505
   ClientTop       =   4425
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6855
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name: Vincent Chen
'Date: 2019/06/03
'Purpose: Puzzle Board Version 2

Private Sub cmdReturn_Click()
    Unload frmAbout
End Sub

Private Sub Form_Load()
    lblAbout.Caption = "Puzzle Boards: An ICS4U Culminating Project" _
     & vbCrLf & vbCrLf & "How to play: " & vbCrLf _
     & "The goal of this game is to line up the tiles in numerical order or to create an image. To move a tile, simply click the tile to move it into the empty spot." _
     & vbCrLf & vbCrLf & "Created by: Vincent Chen" & vbCrLf & "Version 2.0 of 2.0"
End Sub
