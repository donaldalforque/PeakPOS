VERSION 5.00
Begin VB.Form INV_PriceManagerFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "INV_PriceManagerFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pricing Schemes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1575
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   5295
      Begin VB.CommandButton btnPricingScheme 
         Caption         =   "Set up Pricing Schemes"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   5055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Create multiple prices for your products for special events such as Anniversary, Summer Sale, Promos and etc."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   450
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Standard Retail Price"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   5295
      Begin VB.CommandButton btnSRP 
         Caption         =   "Set up Standard Retail Price (SRP)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "This will be the default price for all your products."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   450
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "INV_PriceManagerFrm.frx":000C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price Manager"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   840
      TabIndex        =   1
      Top             =   405
      Width           =   1665
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Use Price Manager to control and setup your product pricing and discounts."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   450
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   4575
      Left            =   120
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "INV_PriceManagerFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnPricingScheme_Click()
    If EditAccessRights(8) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    INV_ProductPricingSchemeFrm.Show (1)
End Sub

Private Sub btnSRP_Click()
    If EditAccessRights(8) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    INV_SRPSetupFrm.Show
End Sub

Private Sub Command2_Click()
    
End Sub
