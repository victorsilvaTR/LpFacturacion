VERSION 5.00
Begin VB.Form FrmStart 
   BorderStyle     =   0  'None
   Caption         =   "LP Facturación"
   ClientHeight    =   8295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8490
   ControlBox      =   0   'False
   Icon            =   "FrmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   8490
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Height          =   8325
      Left            =   0
      Picture         =   "FrmStart.frx":000C
      ScaleHeight     =   8265
      ScaleWidth      =   8490
      TabIndex        =   0
      Top             =   0
      Width           =   8550
      Begin VB.Frame Fr_Invisible 
         Caption         =   "Invisibles"
         Height          =   1335
         Left            =   120
         TabIndex        =   1
         Top             =   5820
         Visible         =   0   'False
         Width           =   2595
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   240
            TabIndex        =   2
            Top             =   480
            Width           =   1275
         End
      End
      Begin VB.Label La_Ver 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "V 0.00.00"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   6240
         TabIndex        =   3
         Top             =   4680
         Width           =   690
      End
   End
End
Attribute VB_Name = "FrmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

   'La_Title = gLexContab
   
   La_Ver = "V " & App.Major & "." & App.Minor & "." & App.Revision

End Sub
