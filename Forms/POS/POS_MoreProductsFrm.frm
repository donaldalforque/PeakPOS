VERSION 5.00
Begin VB.Form POS_MoreProductsFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15105
   Icon            =   "POS_MoreProductsFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   15105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   8400
      Width           =   1700
   End
   Begin VB.CommandButton btnFoodNext 
      BackColor       =   &H00C0FFC0&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   8400
      Width           =   1700
   End
   Begin VB.CommandButton btnFoodPrev 
      BackColor       =   &H00C0FFC0&
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   8400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood60 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   8400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood59 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   8400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood58 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   8400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood57 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   8400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood56 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   8400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood55 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   8400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood54 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   7200
      Width           =   1700
   End
   Begin VB.CommandButton btnFood53 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   7200
      Width           =   1700
   End
   Begin VB.CommandButton btnFood52 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   7200
      Width           =   1700
   End
   Begin VB.CommandButton btnFood51 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   7200
      Width           =   1700
   End
   Begin VB.CommandButton btnFood50 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   7200
      Width           =   1700
   End
   Begin VB.CommandButton btnFood49 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   7200
      Width           =   1700
   End
   Begin VB.CommandButton btnFood48 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   7200
      Width           =   1700
   End
   Begin VB.CommandButton btnFood47 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   7200
      Width           =   1700
   End
   Begin VB.CommandButton btnFood46 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   7200
      Width           =   1700
   End
   Begin VB.CommandButton btnFood45 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   6000
      Width           =   1700
   End
   Begin VB.CommandButton btnFood44 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   6000
      Width           =   1700
   End
   Begin VB.CommandButton btnFood43 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6000
      Width           =   1700
   End
   Begin VB.CommandButton btnFood42 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   6000
      Width           =   1700
   End
   Begin VB.CommandButton btnFood41 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   6000
      Width           =   1700
   End
   Begin VB.CommandButton btnFood40 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6000
      Width           =   1700
   End
   Begin VB.CommandButton btnFood39 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6000
      Width           =   1700
   End
   Begin VB.CommandButton btnFood38 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6000
      Width           =   1700
   End
   Begin VB.CommandButton btnFood37 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6000
      Width           =   1700
   End
   Begin VB.CommandButton btnFood36 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   4800
      Width           =   1700
   End
   Begin VB.CommandButton btnFood35 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   4800
      Width           =   1700
   End
   Begin VB.CommandButton btnFood34 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   4800
      Width           =   1700
   End
   Begin VB.CommandButton btnFood33 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   4800
      Width           =   1700
   End
   Begin VB.CommandButton btnFood32 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   4800
      Width           =   1700
   End
   Begin VB.CommandButton btnFood31 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4800
      Width           =   1700
   End
   Begin VB.CommandButton btnFood30 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4800
      Width           =   1700
   End
   Begin VB.CommandButton btnFood29 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4800
      Width           =   1700
   End
   Begin VB.CommandButton btnFood28 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4800
      Width           =   1700
   End
   Begin VB.CommandButton btnFood27 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   3600
      Width           =   1700
   End
   Begin VB.CommandButton btnFood26 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3600
      Width           =   1700
   End
   Begin VB.CommandButton btnFood25 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3600
      Width           =   1700
   End
   Begin VB.CommandButton btnFood24 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   3600
      Width           =   1700
   End
   Begin VB.CommandButton btnFood23 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   3600
      Width           =   1700
   End
   Begin VB.CommandButton btnFood22 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3600
      Width           =   1700
   End
   Begin VB.CommandButton btnFood21 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3600
      Width           =   1700
   End
   Begin VB.CommandButton btnFood20 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3600
      Width           =   1700
   End
   Begin VB.CommandButton btnFood19 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3600
      Width           =   1700
   End
   Begin VB.CommandButton btnFood18 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood17 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood16 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood15 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood14 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood13 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood12 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood11 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood10 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1200
      Width           =   1700
   End
   Begin VB.CommandButton btnFood8 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1200
      Width           =   1700
   End
   Begin VB.CommandButton btnFood7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1200
      Width           =   1700
   End
   Begin VB.CommandButton btnFood6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1200
      Width           =   1700
   End
   Begin VB.CommandButton btnFood5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1200
      Width           =   1700
   End
   Begin VB.CommandButton btnFood4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1700
   End
   Begin VB.CommandButton btnFood3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
      Width           =   1700
   End
   Begin VB.CommandButton btnFood2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1200
      Width           =   1700
   End
   Begin VB.CommandButton btnFood1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1200
      Width           =   1700
   End
   Begin VB.CommandButton btnNext 
      BackColor       =   &H00C0FFC0&
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   0
      Width           =   1700
   End
   Begin VB.CommandButton btnPrev 
      BackColor       =   &H00C0FFC0&
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   1700
   End
   Begin VB.CommandButton btnCategory7 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   1700
   End
   Begin VB.CommandButton btnCategory6 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   1700
   End
   Begin VB.CommandButton btnCategory5 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1700
   End
   Begin VB.CommandButton btnCategory4 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1700
   End
   Begin VB.CommandButton btnCategory3 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1700
   End
   Begin VB.CommandButton btnCategory2 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1700
   End
   Begin VB.CommandButton btnCategory1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1700
   End
End
Attribute VB_Name = "POS_MoreProductsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurrentCategoryId As Long
Public Sub Populate(ByVal data As String, Optional ByVal RecordId As Long)
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "Category"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Category")
            'lvCategory.ListItems.Clear
            Dim ctr As Integer
            ctr = 1
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        If rec!CategoryId > RecordId Then
                            If ctr = 8 Then Exit Sub 'max out for layout
                            Dim e As Control
                            For Each e In Me.Controls
                                If (TypeOf e Is CommandButton) Then
                                    If e.Name = "btnCategory" & ctr Then
                                        'If IsNull(rec!Name) = False Then
                                            e.Caption = rec!Category
                                        'End If
                                        'If IsNull(rec!barcode) Then
                                            'e.Tag = ""
                                        'Else
                                            e.Tag = rec!CategoryId
                                        'End If
                                        ctr = ctr + 1
                                        Exit For
                                    End If
                                End If
                            Next
                            rec.MoveNext
                        Else
                            rec.MoveNext
                        End If
                    Else
                        rec.MoveNext
                    End If
                Loop
            Else
                MsgBox "No more records to display"
            End If
        Case "CategoryBack"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Category")
            'lvCategory.ListItems.Clear
            'Dim ctr As Integer
            ctr = 1
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        'If rec!CategoryId < RecordId Then
                            If ctr = 8 Then Exit Sub 'max out for layout
                            'Dim e As Control
                            For Each e In Me.Controls
                                If (TypeOf e Is CommandButton) Then
                                    If e.Name = "btnCategory" & ctr Then
                                        'If IsNull(rec!Name) = False Then
                                            e.Caption = rec!Category
                                        'End If
                                        'If IsNull(rec!barcode) Then
                                            'e.Tag = ""
                                        'Else
                                            e.Tag = rec!CategoryId
                                        'End If
                                        ctr = ctr + 1
                                        Exit For
                                    End If
                                End If
                            Next
                            rec.MoveNext
                        'Else
                        '    rec.MoveNext
                        'End If
                    Else
                        rec.MoveNext
                    End If
                Loop
            Else
                MsgBox "No more records to display"
            End If
    End Select
End Sub
Private Sub GetProducts(ByVal CategoryId As Integer, ByVal DisplayType As String, Optional ByVal RecordId As Long)
    'Clear all records for Food
    clearButtons "Food"
    Select Case DisplayType
        Case "Next"
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_Product_GetByCategory"
            cmd.Parameters.Append cmd.CreateParameter("@CategoryId", adInteger, adParamInput, , CategoryId)
            Set rec = cmd.Execute
            Dim ctr As Integer
            ctr = 1
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        If rec!ProductId > RecordId Then
                            Dim e As Control
                            For Each e In Me.Controls
                                If e.Name = "btnFood" & ctr Then
                                    e.Caption = rec!Name & " @ " & FormatNumber(rec!unitprice, 2, vbTrue, vbFalse)
                                    e.Tag = rec!barcode
                                    e.ToolTipText = rec!ProductId
                                    ctr = ctr + 1
                                    Exit For
                                End If
                            Next
                            rec.MoveNext
                        Else
                            rec.MoveNext
                        End If
                    Else
                        rec.MoveNext
                    End If
                Loop
            End If
            con.Close
        Case "Prev"
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_Product_GetByCategory"
            cmd.Parameters.Append cmd.CreateParameter("@CategoryId", adInteger, adParamInput, , CategoryId)
            Set rec = cmd.Execute
            'Dim ctr As Integer
            ctr = 1
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        'If rec!ProductId < RecordId Then
                            'Dim e As Control
                            For Each e In Me.Controls
                                If e.Name = "btnFood" & ctr Then
                                    e.Caption = rec!Name & " @ " & FormatNumber(rec!unitprice, 2, vbTrue, vbFalse)
                                    e.Tag = rec!barcode
                                    e.ToolTipText = rec!ProductId
                                    ctr = ctr + 1
                                    Exit For
                                End If
                            Next
                            rec.MoveNext
                        'Else
                        '    rec.MoveNext
                        'End If
                    Else
                        rec.MoveNext
                    End If
                Loop
            End If
            con.Close
    End Select
    
End Sub
Private Sub btnFontBold(ByVal X As Control)
    Dim ctr As Integer
    Dim e As Control
    ctr = 7
    For Each e In Me.Controls
        If (TypeOf e Is CommandButton) Then
            If e.Name = "btnCategory" & ctr Then
                e.FontBold = False
                ctr = ctr - 1
            End If
        End If
    Next
    X.FontBold = True
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnCategory1_Click()
    btnFontBold btnCategory1
    CurrentCategoryId = Val(btnCategory1.Tag)
    GetProducts Val(btnCategory1.Tag), "Next"
End Sub

Private Sub btnCategory2_Click()
    btnFontBold btnCategory2
    CurrentCategoryId = Val(btnCategory2.Tag)
    GetProducts Val(btnCategory2.Tag), "Next"
End Sub

Private Sub btnCategory3_Click()
    btnFontBold btnCategory3
    CurrentCategoryId = Val(btnCategory3.Tag)
    GetProducts Val(btnCategory3.Tag), "Next"
End Sub

Private Sub btnCategory4_Click()
    btnFontBold btnCategory4
    CurrentCategoryId = Val(btnCategory4.Tag)
    GetProducts Val(btnCategory4.Tag), "Next"
End Sub

Private Sub btnCategory5_Click()
    btnFontBold btnCategory5
    CurrentCategoryId = Val(btnCategory5.Tag)
    GetProducts Val(btnCategory5.Tag), "Next"
End Sub

Private Sub btnCategory6_Click()
    btnFontBold btnCategory6
    CurrentCategoryId = Val(btnCategory6.Tag)
    GetProducts Val(btnCategory6.Tag), "Next"
End Sub

Private Sub btnCategory7_Click()
    btnFontBold btnCategory7
    CurrentCategoryId = Val(btnCategory7.Tag)
    GetProducts Val(btnCategory7.Tag), "Next"
End Sub
Private Sub clearButtons(ByVal ItemType As String)
    Dim e As Control
    Dim ctr As Integer
    
    Select Case ItemType
        Case "All"
            For Each e In Me.Controls
                If (TypeOf e Is CommandButton) Then
                    If e.Name <> "btnPrev" And e.Name <> "btnNext" And e.Name <> "btnFoodPrev" And e.Name <> "btnFoodNext" _
                        And e.Name <> "btnCancel" Then
                        e.Tag = ""
                        e.Caption = ""
                    End If
                End If
            Next
        Case "Food"
            ctr = 60
            For Each e In Me.Controls
                If (TypeOf e Is CommandButton) Then
                    If e.Name = "btnFood" & ctr Then
                        e.Caption = ""
                        e.Tag = ""
                        e.ToolTipText = ""
                        ctr = ctr - 1
                    End If
                End If
            Next
        Case "Category"
            ctr = 7
            For Each e In Me.Controls
                If (TypeOf e Is CommandButton) Then
                    If e.Name = "btnCategory" & ctr Then
                        e.Caption = ""
                        e.Tag = ""
                        ctr = ctr - 1
                    End If
                End If
            Next
    End Select
    
End Sub

Private Sub btnFood1_Click()
    SelectItem btnFood1
End Sub

Private Sub SelectItem(ByVal e As Control)
    With POS_CashierFrm
        .txtBarcode.text = e.Tag
        .txtBarcode_KeyDown 13, 1
        .txtBarcode.text = ""
    End With
End Sub

Private Sub btnFood10_Click()
    SelectItem btnFood10
End Sub

Private Sub btnFood11_Click()
    SelectItem btnFood11
End Sub

Private Sub btnFood12_Click()
    SelectItem btnFood12
End Sub

Private Sub btnFood13_Click()
    SelectItem btnFood13
End Sub

Private Sub btnFood14_Click()
    SelectItem btnFood14
End Sub

Private Sub btnFood15_Click()
    SelectItem btnFood15
End Sub

Private Sub btnFood16_Click()
    SelectItem btnFood16
End Sub

Private Sub btnFood17_Click()
    SelectItem btnFood17
End Sub

Private Sub btnFood18_Click()
    SelectItem btnFood18
End Sub

Private Sub btnFood19_Click()
    SelectItem btnFood19
End Sub

Private Sub btnFood2_Click()
    SelectItem btnFood2
End Sub

Private Sub btnFood20_Click()
    SelectItem btnFood20
End Sub

Private Sub btnFood21_Click()
    SelectItem btnFood21
End Sub

Private Sub btnFood22_Click()
    SelectItem btnFood22
End Sub

Private Sub btnFood23_Click()
    SelectItem btnFood23
End Sub

Private Sub btnFood24_Click()
    SelectItem btnFood24
End Sub

Private Sub btnFood25_Click()
    SelectItem btnFood25
End Sub

Private Sub btnFood26_Click()
    SelectItem btnFood26
End Sub

Private Sub btnFood27_Click()
    SelectItem btnFood27
End Sub

Private Sub btnFood28_Click()
    SelectItem btnFood28
End Sub

Private Sub btnFood29_Click()
    SelectItem btnFood29
End Sub

Private Sub btnFood3_Click()
    SelectItem btnFood3
End Sub

Private Sub btnFood30_Click()
    SelectItem btnFood30
End Sub

Private Sub btnFood31_Click()
    SelectItem btnFood31
End Sub

Private Sub btnFood32_Click()
    SelectItem btnFood32
End Sub

Private Sub btnFood33_Click()
    SelectItem btnFood33
End Sub

Private Sub btnFood34_Click()
    SelectItem btnFood34
End Sub

Private Sub btnFood35_Click()
    SelectItem btnFood35
End Sub

Private Sub btnFood36_Click()
    SelectItem btnFood36
End Sub

Private Sub btnFood37_Click()
    SelectItem btnFood37
End Sub

Private Sub btnFood38_Click()
    SelectItem btnFood38
End Sub

Private Sub btnFood39_Click()
    SelectItem btnFood39
End Sub

Private Sub btnFood4_Click()
    SelectItem btnFood4
End Sub

Private Sub btnFood40_Click()
    SelectItem btnFood40
End Sub

Private Sub btnFood41_Click()
    SelectItem btnFood41
End Sub

Private Sub btnFood42_Click()
    SelectItem btnFood42
End Sub

Private Sub btnFood43_Click()
    SelectItem btnFood43
End Sub

Private Sub btnFood44_Click()
    SelectItem btnFood44
End Sub

Private Sub btnFood45_Click()
    SelectItem btnFood45
End Sub

Private Sub btnFood46_Click()
    SelectItem btnFood46
End Sub

Private Sub btnFood47_Click()
    SelectItem btnFood47
End Sub

Private Sub btnFood48_Click()
    SelectItem btnFood48
End Sub

Private Sub btnFood49_Click()
    SelectItem btnFood49
End Sub

Private Sub btnFood5_Click()
    SelectItem btnFood5
End Sub

Private Sub btnFood50_Click()
    SelectItem btnFood50
End Sub

Private Sub btnFood51_Click()
    SelectItem btnFood51
End Sub

Private Sub btnFood52_Click()
    SelectItem btnFood52
End Sub

Private Sub btnFood53_Click()
    SelectItem btnFood53
End Sub

Private Sub btnFood54_Click()
    SelectItem btnFood54
End Sub

Private Sub btnFood55_Click()
    SelectItem btnFood55
End Sub

Private Sub btnFood56_Click()
    SelectItem btnFood56
End Sub

Private Sub btnFood57_Click()
    SelectItem btnFood57
End Sub

Private Sub btnFood58_Click()
    SelectItem btnFood58
End Sub

Private Sub btnFood59_Click()
    SelectItem btnFood59
End Sub

Private Sub btnFood6_Click()
    SelectItem btnFood6
End Sub

Private Sub btnFood60_Click()
    SelectItem btnFood60
End Sub

Private Sub btnFood7_Click()
    SelectItem btnFood7
End Sub

Private Sub btnFood8_Click()
    SelectItem btnFood8
End Sub

Private Sub btnFood9_Click()
    SelectItem btnFood9
End Sub

Private Sub btnFoodNext_Click()
    'Get last recordId with highest number
    Dim LastRecordId As String
    Dim e As Control
    Dim ctr As Integer
    
    If btnFood60.ToolTipText = "" Then
        LastRecordId = 0
    Else
        LastRecordId = Val(btnFood60.ToolTipText)
    End If
    
    'clear current data records
    clearButtons "Food"
    
    GetProducts CurrentCategoryId, "Next", Val(LastRecordId)
End Sub

Private Sub btnFoodPrev_Click()
    'Get last recordId with highest number
    Dim LastRecordId As String
    Dim e As Control
    Dim ctr As Integer
    
    If btnFood1.ToolTipText = "" Then
        LastRecordId = 0
    Else
        LastRecordId = Val(btnFood1.ToolTipText)
    End If
    
    'clear current data records
    clearButtons "Food"
    
    GetProducts CurrentCategoryId, "Prev", Val(LastRecordId)
End Sub

Private Sub btnNext_Click()
    'Get last recordId with highest number
    Dim LastRecordId As Integer
    Dim e As Control
    Dim ctr As Integer
    
    If btnCategory7.Tag = "" Then
        LastRecordId = 0
    Else
        LastRecordId = btnCategory7.Tag
    End If
    
    'clear current data records
    clearButtons "Category"
    
    Populate "Category", LastRecordId
    
End Sub

Private Sub btnPrev_Click()
    'Get last recordId with highest number
    Dim LastRecordId As Integer
    Dim e As Control
    Dim ctr As Integer
    
    If btnCategory1.Tag = "" Then
        LastRecordId = 0
    Else
        LastRecordId = btnCategory1.Tag
    End If
    
    
    'clear current data records
    clearButtons "Category"
    
    Populate "CategoryBack", LastRecordId
End Sub

Private Sub Form_Load()
    Populate "Category", 0
    CurrentCategoryId = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CurrentCategoryId = 0
End Sub
