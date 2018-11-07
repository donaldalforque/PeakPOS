VERSION 5.00
Begin VB.Form POS_PlayHouseFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15105
   Icon            =   "POS_PlayHouseFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   15105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton btnFood64 
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
      TabIndex        =   71
      Top             =   8400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood65 
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
      TabIndex        =   70
      Top             =   8400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood66 
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
      TabIndex        =   69
      Top             =   8400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood67 
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
      TabIndex        =   68
      Top             =   8400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood68 
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
      TabIndex        =   67
      Top             =   8400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood69 
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
      TabIndex        =   66
      Top             =   8400
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
      TabIndex        =   62
      Top             =   0
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
      TabIndex        =   61
      Top             =   0
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
      TabIndex        =   60
      Top             =   0
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
      TabIndex        =   59
      Top             =   0
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
      TabIndex        =   58
      Top             =   0
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
      TabIndex        =   57
      Top             =   0
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
      TabIndex        =   56
      Top             =   0
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
      TabIndex        =   55
      Top             =   0
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
      TabIndex        =   54
      Top             =   0
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
      TabIndex        =   53
      Top             =   1200
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
      TabIndex        =   52
      Top             =   1200
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
      TabIndex        =   51
      Top             =   1200
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
      TabIndex        =   50
      Top             =   1200
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
      TabIndex        =   49
      Top             =   1200
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
      TabIndex        =   48
      Top             =   1200
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
      TabIndex        =   47
      Top             =   1200
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
      TabIndex        =   46
      Top             =   1200
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
      TabIndex        =   45
      Top             =   1200
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
      TabIndex        =   44
      Top             =   2400
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
      TabIndex        =   43
      Top             =   2400
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
      TabIndex        =   42
      Top             =   2400
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
      TabIndex        =   41
      Top             =   2400
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
      TabIndex        =   40
      Top             =   2400
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
      TabIndex        =   39
      Top             =   2400
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
      TabIndex        =   38
      Top             =   2400
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
      TabIndex        =   37
      Top             =   2400
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
      TabIndex        =   36
      Top             =   2400
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
      TabIndex        =   35
      Top             =   3600
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
      TabIndex        =   34
      Top             =   3600
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
      TabIndex        =   33
      Top             =   3600
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
      TabIndex        =   32
      Top             =   3600
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
      TabIndex        =   31
      Top             =   3600
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
      TabIndex        =   30
      Top             =   3600
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
      TabIndex        =   29
      Top             =   3600
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
      TabIndex        =   28
      Top             =   3600
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
      TabIndex        =   27
      Top             =   3600
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
      TabIndex        =   26
      Top             =   4800
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
      TabIndex        =   25
      Top             =   4800
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
      TabIndex        =   24
      Top             =   4800
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
      TabIndex        =   23
      Top             =   4800
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
      TabIndex        =   22
      Top             =   4800
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
      TabIndex        =   21
      Top             =   4800
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
      TabIndex        =   20
      Top             =   4800
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
      TabIndex        =   19
      Top             =   4800
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
      TabIndex        =   18
      Top             =   4800
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
      TabIndex        =   17
      Top             =   6000
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
      TabIndex        =   16
      Top             =   6000
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
      TabIndex        =   15
      Top             =   6000
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
      TabIndex        =   14
      Top             =   6000
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
      TabIndex        =   13
      Top             =   6000
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
      TabIndex        =   12
      Top             =   6000
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
      TabIndex        =   11
      Top             =   6000
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
      TabIndex        =   10
      Top             =   6000
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
      TabIndex        =   9
      Top             =   6000
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
      TabIndex        =   8
      Top             =   7200
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
      TabIndex        =   7
      Top             =   7200
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
      TabIndex        =   6
      Top             =   7200
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
      TabIndex        =   5
      Top             =   7200
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
      TabIndex        =   4
      Top             =   7200
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
      TabIndex        =   3
      Top             =   7200
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   8400
      Width           =   1700
   End
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
      TabIndex        =   0
      Top             =   8400
      Width           =   1700
   End
   Begin VB.CommandButton btnFood61 
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
      TabIndex        =   65
      Top             =   7200
      Width           =   1700
   End
   Begin VB.CommandButton btnFood62 
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
      TabIndex        =   64
      Top             =   7200
      Width           =   1700
   End
   Begin VB.CommandButton btnFood63 
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
      TabIndex        =   63
      Top             =   7200
      Width           =   1700
   End
End
Attribute VB_Name = "POS_PlayHouseFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Populate(ByVal data As String, Optional ByVal RecordId As Long)
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "Slot"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Slot")
            'lvCategory.ListItems.Clear
            Dim ctr As Integer
            ctr = 1
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        If rec!SlotId > RecordId Then
                            'If ctr = 8 Then Exit Sub 'max out for layout
                            Dim e As Control
                            For Each e In Me.Controls
                                If (TypeOf e Is CommandButton) Then
                                    If e.Name = "btnFood" & ctr Then
                                        'If IsNull(rec!Name) = False Then
                                            e.Caption = rec!Slot
                                        'End If
                                        'If IsNull(rec!barcode) Then
                                            'e.Tag = ""
                                        'Else
                                            e.Tag = rec!SlotId
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
        Case "Session"
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "POS_Session_Get"
            'cmd.Parameters.Append cmd.CreateParameter("@POS_SessionId", adInteger, adParamInput, , POS_SessionId)
            'Dim ctr As Integer
            ctr = 1
            Set rec = cmd.Execute
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!pos_statusid <= 3 Then
                        'Dim e As Control
                        'Me.Caption = rec!SlotId
                        ctr = rec!SlotId
                        For Each e In Me.Controls
                            If (TypeOf e Is CommandButton) Then
                                If e.Name = "btnFood" & ctr Then
                                    'If rec!SlotId = ctr Then
                                        e.Caption = e.Caption & " : " & rec!childname
                                        e.ToolTipText = rec!POS_SessionId
                                        If rec!pos_statusid = 2 Or rec!pos_statusid = 3 Then
                                            e.BackColor = vbRed
                                        Else
                                            e.BackColor = vbGreen
                                        End If
                                    'End If
                                End If
                            End If
                        Next
                        rec.MoveNext
                    Else
                        rec.MoveNext
                    End If
                Loop
            End If
            con.Close
    End Select
End Sub
Private Sub clearButtons(ByVal ItemType As String)
    Dim e As Control
    Dim ctr As Integer
    
    Select Case ItemType
        Case "All"
            For Each e In Me.Controls
                If (TypeOf e Is CommandButton) Then
                    If e.Name <> "btnFoodPrev" And e.Name <> "btnFoodNext" _
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
        Case "Slot"
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
Private Sub LoadSession(ByVal e As Control)
    POS_SessionFrm.LoadVar 0, Val(e.Tag)
    POS_SessionFrm.POS_SessionId = e.ToolTipText
    If Val(e.Tag) <> 0 Then
        POS_SessionFrm.Show (1)
    End If
End Sub
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnFood1_Click()
    LoadSession btnFood1
End Sub

Private Sub btnFood10_Click()
    LoadSession btnFood10
End Sub

Private Sub btnFood11_Click()
    LoadSession btnFood11
End Sub

Private Sub btnFood12_Click()
    LoadSession btnFood12
End Sub

Private Sub btnFood13_Click()
    LoadSession btnFood13
End Sub

Private Sub btnFood14_Click()
    LoadSession btnFood14
End Sub

Private Sub btnFood15_Click()
    LoadSession btnFood15
End Sub

Private Sub btnFood16_Click()
    LoadSession btnFood16
End Sub

Private Sub btnFood17_Click()
    LoadSession btnFood17
End Sub

Private Sub btnFood18_Click()
    LoadSession btnFood18
End Sub

Private Sub btnFood19_Click()
    LoadSession btnFood19
End Sub

Private Sub btnFood2_Click()
    LoadSession btnFood2
End Sub

Private Sub btnFood20_Click()
    LoadSession btnFood20
End Sub

Private Sub btnFood21_Click()
    LoadSession btnFood21
End Sub

Private Sub btnFood22_Click()
    LoadSession btnFood22
End Sub

Private Sub btnFood23_Click()
    LoadSession btnFood23
End Sub

Private Sub btnFood24_Click()
    LoadSession btnFood24
End Sub

Private Sub btnFood25_Click()
    LoadSession btnFood25
End Sub

Private Sub btnFood26_Click()
    LoadSession btnFood26
End Sub

Private Sub btnFood27_Click()
    LoadSession btnFood27
End Sub

Private Sub btnFood28_Click()
    LoadSession btnFood28
End Sub

Private Sub btnFood29_Click()
    LoadSession btnFood29
End Sub

Private Sub btnFood3_Click()
   LoadSession btnFood3
End Sub

Private Sub btnFood30_Click()
    LoadSession btnFood30
End Sub

Private Sub btnFood31_Click()
    LoadSession btnFood31
End Sub

Private Sub btnFood32_Click()
    LoadSession btnFood32
End Sub

Private Sub btnFood33_Click()
    LoadSession btnFood33
End Sub

Private Sub btnFood34_Click()
    LoadSession btnFood34
End Sub

Private Sub btnFood35_Click()
    LoadSession btnFood35
End Sub

Private Sub btnFood36_Click()
    LoadSession btnFood36
End Sub

Private Sub btnFood37_Click()
    LoadSession btnFood37
End Sub

Private Sub btnFood38_Click()
    LoadSession btnFood38
End Sub

Private Sub btnFood39_Click()
    LoadSession btnFood39
End Sub

Private Sub btnFood4_Click()
    LoadSession btnFood4
End Sub

Private Sub btnFood40_Click()
    LoadSession btnFood40
End Sub

Private Sub btnFood41_Click()
    LoadSession btnFood41
End Sub

Private Sub btnFood42_Click()
    LoadSession btnFood42
End Sub

Private Sub btnFood43_Click()
    LoadSession btnFood43
End Sub

Private Sub btnFood44_Click()
    LoadSession btnFood44
End Sub

Private Sub btnFood45_Click()
    LoadSession btnFood45
End Sub

Private Sub btnFood46_Click()
    LoadSession btnFood46
End Sub

Private Sub btnFood47_Click()
    LoadSession btnFood47
End Sub

Private Sub btnFood48_Click()
    LoadSession btnFood48
End Sub

Private Sub btnFood49_Click()
    LoadSession btnFood49
End Sub

Private Sub btnFood5_Click()
    LoadSession btnFood5
End Sub

Private Sub btnFood50_Click()
    LoadSession btnFood50
End Sub

Private Sub btnFood51_Click()
    LoadSession btnFood51
End Sub

Private Sub btnFood52_Click()
    LoadSession btnFood52
End Sub

Private Sub btnFood53_Click()
    LoadSession btnFood53
End Sub

Private Sub btnFood54_Click()
    LoadSession btnFood54
End Sub

Private Sub btnFood55_Click()
    LoadSession btnFood55
End Sub

Private Sub btnFood56_Click()
    LoadSession btnFood56
End Sub

Private Sub btnFood57_Click()
    LoadSession btnFood57
End Sub

Private Sub btnFood58_Click()
    LoadSession btnFood58
End Sub

Private Sub btnFood59_Click()
    LoadSession btnFood59
End Sub

Private Sub btnFood6_Click()
    LoadSession btnFood6
End Sub

Private Sub btnFood60_Click()
    LoadSession btnFood60
End Sub

Private Sub btnFood61_Click()
    LoadSession btnFood61
End Sub

Private Sub btnFood62_Click()
    LoadSession btnFood62
End Sub

Private Sub btnFood63_Click()
    LoadSession btnFood63
End Sub

Private Sub btnFood64_Click()
    LoadSession btnFood64
End Sub

Private Sub btnFood65_Click()
    LoadSession btnFood65
End Sub

Private Sub btnFood66_Click()
    LoadSession btnFood66
End Sub

Private Sub btnFood67_Click()
    LoadSession btnFood67
End Sub

Private Sub btnFood68_Click()
    LoadSession btnFood68
End Sub

Private Sub btnFood69_Click()
    LoadSession btnFood69
End Sub

Private Sub btnFood7_Click()
    LoadSession btnFood7
End Sub

Private Sub btnFood8_Click()
    LoadSession btnFood8
End Sub

Private Sub btnFood9_Click()
    LoadSession btnFood9
End Sub

Private Sub btnFoodNext_Click()
    'Get last recordId with highest number
    Dim LastRecordId As Integer
    Dim e As Control
    Dim ctr As Integer
    
    If btnFood69.Tag = "" Then
        LastRecordId = 0
    Else
        LastRecordId = btnFood69.Tag
    End If
    
    'clear current data records
    clearButtons "All"
    
    Populate "Slot", LastRecordId
    
    Populate "Slot", 0
    Populate "Session"
End Sub

Private Sub btnFoodPrev_Click()
    'Get last recordId with highest number
    Dim LastRecordId As Integer
    Dim e As Control
    Dim ctr As Integer
    
    'If btnFood69.Tag = "" Then
        LastRecordId = 0
    'Else
'        LastRecordId = btnFood69.Tag
    'End If
    
    'clear current data records
    clearButtons "All"
    
    Populate "Slot", LastRecordId
    
    Populate "Slot", 0
    Populate "Session"
End Sub

Public Sub Form_Load()
    Populate "Slot", 0
    Populate "Session"
End Sub

Private Sub Timer1_Timer()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_Sessions_CheckTime"
    cmd.Execute
    con.Close
    
    Populate "Slot", 0
    Populate "Session"
End Sub
