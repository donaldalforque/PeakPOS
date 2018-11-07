VERSION 5.00
Begin VB.Form GLOBAL_MessageFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   ClipControls    =   0   'False
   Icon            =   "GLOBAL_MessageFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSearch 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   1280
      Width           =   1215
   End
   Begin VB.Label lblErrorMessage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save Failed. Price must be numeric."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   4320
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   600
      Picture         =   "GLOBAL_MessageFrm.frx":000C
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "GLOBAL_MessageFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSearch_Click()
    Unload Me
End Sub

