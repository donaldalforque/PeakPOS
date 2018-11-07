VERSION 5.00
Begin VB.Form SYS_OskAlphaFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   Icon            =   "SYS_OskAlphaFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCancel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7080
      Width           =   11895
   End
   Begin VB.CommandButton btnAccept 
      BackColor       =   &H0080FF80&
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   5760
      Width           =   11895
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11895
   End
   Begin VB.CommandButton btn1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1000
   End
   Begin VB.CommandButton btn2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1000
   End
   Begin VB.CommandButton btn3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   1000
   End
   Begin VB.CommandButton btn4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   1000
   End
   Begin VB.CommandButton btn5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   1000
   End
   Begin VB.CommandButton btn6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   1000
   End
   Begin VB.CommandButton btn7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   960
      Width           =   1000
   End
   Begin VB.CommandButton btn8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   960
      Width           =   1000
   End
   Begin VB.CommandButton btn9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   960
      Width           =   1000
   End
   Begin VB.CommandButton btn0 
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   960
      Width           =   1000
   End
   Begin VB.CommandButton btnQ 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1920
      Width           =   1000
   End
   Begin VB.CommandButton btnW 
      BackColor       =   &H00C0FFFF&
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1920
      Width           =   1000
   End
   Begin VB.CommandButton btnE 
      BackColor       =   &H00C0FFFF&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1920
      Width           =   1000
   End
   Begin VB.CommandButton btnR 
      BackColor       =   &H00C0FFFF&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1920
      Width           =   1000
   End
   Begin VB.CommandButton btnT 
      BackColor       =   &H00C0FFFF&
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1920
      Width           =   1000
   End
   Begin VB.CommandButton btnY 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1920
      Width           =   1000
   End
   Begin VB.CommandButton btnU 
      BackColor       =   &H00C0FFFF&
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1920
      Width           =   1000
   End
   Begin VB.CommandButton btnI 
      BackColor       =   &H00C0FFFF&
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1920
      Width           =   1000
   End
   Begin VB.CommandButton btnO 
      BackColor       =   &H00C0FFFF&
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1920
      Width           =   1000
   End
   Begin VB.CommandButton btnP 
      BackColor       =   &H00C0FFFF&
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1920
      Width           =   1000
   End
   Begin VB.CommandButton btnClear 
      BackColor       =   &H00C0FFFF&
      Caption         =   "< BACK SPACE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   960
      Width           =   2300
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   1920
      Width           =   2300
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2880
      Width           =   500
   End
   Begin VB.CommandButton btnA 
      BackColor       =   &H00C0FFFF&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2880
      Width           =   1000
   End
   Begin VB.CommandButton btnS 
      BackColor       =   &H00C0FFFF&
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2880
      Width           =   1000
   End
   Begin VB.CommandButton btnD 
      BackColor       =   &H00C0FFFF&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2880
      Width           =   1000
   End
   Begin VB.CommandButton btnF 
      BackColor       =   &H00C0FFFF&
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2880
      Width           =   1000
   End
   Begin VB.CommandButton btnG 
      BackColor       =   &H00C0FFFF&
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2880
      Width           =   1000
   End
   Begin VB.CommandButton btnH 
      BackColor       =   &H00C0FFFF&
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2880
      Width           =   1000
   End
   Begin VB.CommandButton btnJ 
      BackColor       =   &H00C0FFFF&
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2880
      Width           =   1000
   End
   Begin VB.CommandButton btnK 
      BackColor       =   &H00C0FFFF&
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2880
      Width           =   1000
   End
   Begin VB.CommandButton btnL 
      BackColor       =   &H00C0FFFF&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2880
      Width           =   1000
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2880
      Width           =   2775
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton btnZ 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton btnX 
      BackColor       =   &H00C0FFFF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton btnC 
      BackColor       =   &H00C0FFFF&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton btnV 
      BackColor       =   &H00C0FFFF&
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton btnB 
      BackColor       =   &H00C0FFFF&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton btnN 
      BackColor       =   &H00C0FFFF&
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton btnM 
      BackColor       =   &H00C0FFFF&
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0FFFF&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   3840
      Width           =   2895
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton btnSpace 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SPACE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
      Width           =   6765
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   4800
      Width           =   3735
   End
End
Attribute VB_Name = "SYS_OskAlphaFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public txtControl As TextBox

Private Sub btn0_Click()
    txtValue.text = txtValue.text & "0"
End Sub

Private Sub btn1_Click()
    txtValue.text = txtValue.text & "1"
End Sub

Private Sub btn2_Click()
    txtValue.text = txtValue.text & "2"
End Sub

Private Sub btn3_Click()
    txtValue.text = txtValue.text & "3"
End Sub

Private Sub btn4_Click()
    txtValue.text = txtValue.text & "4"
End Sub

Private Sub btn5_Click()
    txtValue.text = txtValue.text & "5"
End Sub

Private Sub btn6_Click()
    txtValue.text = txtValue.text & "6"
End Sub

Private Sub btn7_Click()
    txtValue.text = txtValue.text & "7"
End Sub

Private Sub btn8_Click()
    txtValue.text = txtValue.text & "8"
End Sub

Private Sub btn9_Click()
    txtValue.text = txtValue.text & "9"
End Sub

Private Sub btnA_Click()
    txtValue.text = txtValue.text & "A"
End Sub

Private Sub btnAccept_Click()
    txtControl.text = txtValue.text
    Unload Me
End Sub

Private Sub btnB_Click()
    txtValue.text = txtValue.text & "B"
End Sub

Private Sub btnC_Click()
    txtValue.text = txtValue.text & "C"
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub Command11_Click()

End Sub

Private Sub btnClear_Click()
    If txtValue.text = "" Then Exit Sub
    txtValue.text = Mid(txtValue.text, 1, Len(txtValue.text) - 1)
End Sub

Private Sub btnD_Click()
    txtValue.text = txtValue.text & "D"
End Sub

Private Sub btnE_Click()
    txtValue.text = txtValue.text & "E"
End Sub

Private Sub btnF_Click()
    txtValue.text = txtValue.text & "F"
End Sub

Private Sub btnG_Click()
    txtValue.text = txtValue.text & "G"
End Sub

Private Sub btnH_Click()
    txtValue.text = txtValue.text & "H"
End Sub

Private Sub btnI_Click()
    txtValue.text = txtValue.text & "I"
End Sub

Private Sub btnJ_Click()
    txtValue.text = txtValue.text & "J"
End Sub

Private Sub btnK_Click()
    txtValue.text = txtValue.text & "K"
End Sub

Private Sub btnL_Click()
    txtValue.text = txtValue.text & "L"
End Sub

Private Sub btnM_Click()
    txtValue.text = txtValue.text & "M"
End Sub

Private Sub btnN_Click()
    txtValue.text = txtValue.text & "N"
End Sub

Private Sub btnO_Click()
    txtValue.text = txtValue.text & "O"
End Sub

Private Sub btnP_Click()
    txtValue.text = txtValue.text & "P"
End Sub

Private Sub btnQ_Click()
    txtValue.text = txtValue.text & "Q"
End Sub

Private Sub btnR_Click()
    txtValue.text = txtValue.text & "R"
End Sub

Private Sub btnS_Click()
    txtValue.text = txtValue.text & "S"
End Sub

Private Sub btnSpace_Click()
    txtValue.text = txtValue.text & " "
End Sub

Private Sub btnT_Click()
    txtValue.text = txtValue.text & "T"
End Sub

Private Sub btnU_Click()
    txtValue.text = txtValue.text & "U"
End Sub

Private Sub btnV_Click()
    txtValue.text = txtValue.text & "V"
End Sub

Private Sub btnW_Click()
    txtValue.text = txtValue.text & "W"
End Sub

Private Sub btnX_Click()
    txtValue.text = txtValue.text & "X"
End Sub

Private Sub btnY_Click()
    txtValue.text = txtValue.text & "Y"
End Sub

Private Sub btnZ_Click()
    txtValue.text = txtValue.text & "Z"
End Sub

Private Sub Command15_Click()
    txtValue.text = txtValue.text & "."
End Sub

