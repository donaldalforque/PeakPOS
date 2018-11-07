VERSION 5.00
Begin VB.Form BASE_SystemSettingsFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Settings"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   Icon            =   "BASE_SystemSettingsFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FRE_Inventory 
      BackColor       =   &H00FFFFFF&
      Height          =   6375
      Left            =   1800
      TabIndex        =   31
      Top             =   0
      Width           =   7695
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         Caption         =   "Alerts"
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
         Left            =   240
         TabIndex        =   41
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton Command9 
         Appearance      =   0  'Flat
         Caption         =   "Remove Expired"
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
         Left            =   240
         TabIndex        =   40
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         Caption         =   "Units"
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
         Left            =   240
         TabIndex        =   39
         Top             =   2400
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         Caption         =   "Locations"
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
         Left            =   240
         TabIndex        =   38
         Top             =   2880
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Caption         =   "Charges"
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
         Left            =   240
         TabIndex        =   37
         Top             =   3360
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Caption         =   "Shrinkages"
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
         Left            =   240
         TabIndex        =   36
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Caption         =   "Banks"
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
         Left            =   240
         TabIndex        =   35
         Top             =   4320
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Caption         =   "Funds"
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
         Left            =   240
         TabIndex        =   34
         Top             =   4800
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Caption         =   "Expenses"
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
         Left            =   240
         TabIndex        =   33
         Top             =   5280
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "Workstations"
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
         Left            =   240
         TabIndex        =   32
         Top             =   5760
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Create and setup inventory alert dates and frequency so you get updated with your products available quantity."
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
         TabIndex        =   53
         Top             =   720
         Width           =   7215
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory Settings"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   345
         Left            =   240
         TabIndex        =   52
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Setup inventory alerts for products with low inventory"
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
         Height          =   210
         Left            =   2760
         TabIndex        =   51
         Top             =   1500
         Width           =   4695
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Removes expired product information."
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
         Height          =   210
         Left            =   2760
         TabIndex        =   50
         Top             =   1980
         Width           =   4695
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit of measures for products."
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
         Height          =   210
         Left            =   2760
         TabIndex        =   49
         Top             =   2460
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Locations for product inventories."
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
         Height          =   210
         Left            =   2760
         TabIndex        =   48
         Top             =   2940
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase order charges."
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
         Height          =   210
         Left            =   2760
         TabIndex        =   47
         Top             =   3420
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase order product shrinkages."
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
         Height          =   210
         Left            =   2760
         TabIndex        =   46
         Top             =   3900
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Banks and accounts for cash flow monitoring."
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
         Height          =   210
         Left            =   2760
         TabIndex        =   45
         Top             =   4380
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fund accounts such as petty cash funds."
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
         Height          =   210
         Left            =   2760
         TabIndex        =   44
         Top             =   4860
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Company expenses such as payroll and misc."
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
         Height          =   210
         Left            =   2760
         TabIndex        =   43
         Top             =   5340
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Setup for POS Workstations"
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
         Height          =   210
         Left            =   2760
         TabIndex        =   42
         Top             =   5820
         Visible         =   0   'False
         Width           =   4695
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   0
      TabIndex        =   13
      Top             =   -120
      Width           =   1695
      Begin VB.CommandButton btnTransferStock 
         Caption         =   "Inventory"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   22
         Picture         =   "BASE_SystemSettingsFrm.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2640
         Width           =   1680
      End
      Begin VB.CommandButton btnSalesOrder 
         Caption         =   "Purchasing"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   22
         Picture         =   "BASE_SystemSettingsFrm.frx":6E26
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1800
         Width           =   1680
      End
      Begin VB.CommandButton btnPurchaseOrder 
         Caption         =   "Sales"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   22
         Picture         =   "BASE_SystemSettingsFrm.frx":73A0
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   960
         Width           =   1680
      End
      Begin VB.CommandButton btnPOS 
         Caption         =   "Point of Sale"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   22
         Picture         =   "BASE_SystemSettingsFrm.frx":7A31
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   1680
      End
   End
   Begin VB.Frame FRE_Main 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   7095
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.Frame FRE_POS 
         BackColor       =   &H00FFFFFF&
         Height          =   6375
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Width           =   7695
         Begin VB.CommandButton btnWorkstations 
            Appearance      =   0  'Flat
            Caption         =   "Workstations"
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
            Left            =   240
            TabIndex        =   10
            Top             =   5760
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CommandButton btnExpenses 
            Appearance      =   0  'Flat
            Caption         =   "Expenses"
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
            Left            =   240
            TabIndex        =   9
            Top             =   5280
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CommandButton btnFunds 
            Appearance      =   0  'Flat
            Caption         =   "Funds"
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
            Left            =   240
            TabIndex        =   8
            Top             =   4800
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CommandButton btnBanks 
            Appearance      =   0  'Flat
            Caption         =   "Banks"
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
            Left            =   240
            TabIndex        =   7
            Top             =   4320
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CommandButton btnShrinkages 
            Appearance      =   0  'Flat
            Caption         =   "Shrinkages"
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
            Left            =   240
            TabIndex        =   6
            Top             =   3840
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CommandButton btnCharges 
            Appearance      =   0  'Flat
            Caption         =   "Charges"
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
            Left            =   240
            TabIndex        =   5
            Top             =   3360
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CommandButton btnDiscounts 
            Appearance      =   0  'Flat
            Caption         =   "Discounts"
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
            Left            =   240
            TabIndex        =   4
            Top             =   2880
            Width           =   2175
         End
         Begin VB.CommandButton btnUnits 
            Appearance      =   0  'Flat
            Caption         =   "Loyalty Card Settings"
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
            Left            =   240
            TabIndex        =   3
            Top             =   2400
            Width           =   2175
         End
         Begin VB.CommandButton btnTerms 
            Appearance      =   0  'Flat
            Caption         =   "User Validation"
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
            Left            =   240
            TabIndex        =   2
            Top             =   1920
            Width           =   2175
         End
         Begin VB.CommandButton btnPaymentMethod 
            Appearance      =   0  'Flat
            Caption         =   "Product Display"
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
            Left            =   240
            TabIndex        =   1
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "Setup for POS Workstations"
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
            Height          =   210
            Left            =   2760
            TabIndex        =   29
            Top             =   5820
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Company expenses such as payroll and misc."
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
            Height          =   210
            Left            =   2760
            TabIndex        =   28
            Top             =   5340
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Fund accounts such as petty cash funds."
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
            Height          =   210
            Left            =   2760
            TabIndex        =   27
            Top             =   4860
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Banks and accounts for cash flow monitoring."
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
            Height          =   210
            Left            =   2760
            TabIndex        =   26
            Top             =   4380
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase order product shrinkages."
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
            Height          =   210
            Left            =   2760
            TabIndex        =   25
            Top             =   3900
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase order charges."
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
            Height          =   210
            Left            =   2760
            TabIndex        =   24
            Top             =   3420
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Create custom discounts for different customers."
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
            Height          =   210
            Left            =   2760
            TabIndex        =   23
            Top             =   2940
            Width           =   4695
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Setup loyalty card settings"
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
            Height          =   210
            Left            =   2760
            TabIndex        =   22
            Top             =   2460
            Width           =   4695
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Allows you to enable user validation on POS Buttons"
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
            Height          =   210
            Left            =   2760
            TabIndex        =   21
            Top             =   1980
            Width           =   4695
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Show your frequently ordered products on the POS Screen"
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
            Height          =   210
            Left            =   2760
            TabIndex        =   20
            Top             =   1500
            Width           =   4695
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Point of Sale Settings"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   345
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   2400
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "You can set levels of restrictions for POS Buttons to help ensure integrity in the cashier operations."
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
            TabIndex        =   18
            Top             =   720
            Width           =   7215
         End
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Save && Close"
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
         Left            =   5040
         TabIndex        =   11
         Top             =   6480
         Width           =   1335
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   "Cancel"
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
         Left            =   6480
         TabIndex        =   12
         Top             =   6480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "BASE_SystemSettingsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim UserId As Integer
'Public Sub Populate(ByVal data As String)
'    Dim item As MSComctlLib.ListItem
'    Select Case data
'        Case "Company"
'            Set rec = New ADODB.Recordset
'            Set rec = Global_Data("Company")
'            If Not rec.EOF Then
'                If IsNull(rec!Name) = False Then txtCompanyName.text = rec!Name
'                If IsNull(rec!address1) = False Then txtAddress1.text = rec!address1
'                If IsNull(rec!address2) = False Then txtAddress2.text = rec!address2
'                If IsNull(rec!email) = False Then txtEmail.text = rec!email
'                If IsNull(rec!Phone) = False Then txtPhone.text = rec!Phone
'                If IsNull(rec!fax) = False Then txtFax.text = rec!fax
'                If IsNull(rec!website) = False Then txtWebsite.text = rec!website
'            End If
'        Case "User"
'            Set rec = New ADODB.Recordset
'            Set rec = Global_Data("User")
'            lvUsers.ListItems.Clear
'            If Not rec.EOF Then
'                Do Until rec.EOF
'                    If rec!isActive = "True" Then
'                        Set item = lvUsers.ListItems.add(, , "")
'                            item.SubItems(1) = rec!UserId
'                            item.SubItems(2) = rec!UserNumber
'                            item.SubItems(3) = rec!Name
'                            item.SubItems(4) = rec!Role
'                            item.SubItems(5) = rec!UserRoleId
'                            item.Checked = True
'                    End If
'                    rec.MoveNext
'                Loop
'            End If
'        Case "UserRoles"
'            Set rec = New ADODB.Recordset
'            Set rec = Global_Data("UserRoles")
'            cmbRoles.Clear
'            If Not rec.EOF Then
'                Do Until rec.EOF
'                    If rec!UserRoleId = 1 Then rec.MoveNext
'                    cmbRoles.AddItem rec!Role
'                    cmbRoles.ItemData(cmbRoles.NewIndex) = rec!UserRoleId
'                    rec.MoveNext
'                Loop
'            End If
'            cmbRoles.ListIndex = 0
'        Case "Documents"
'            Set rec = New ADODB.Recordset
'            Set rec = Global_Data("Documents")
'            If Not rec.EOF Then
'                Do Until rec.EOF
'                    Select Case rec!DocNoFormatId
'                        Case 1 '-Purchase Order
'                            txtPrefix_PurchaseOrder.text = rec!prefix
'                            txtNextNumber_PurchaseOrder.text = rec!nextnumber
'                        Case 2 '-Sales Order
'                            txtPrefix_SalesOrder.text = rec!prefix
'                            txtNextNumber_SalesOrder.text = rec!nextnumber
'                        Case 3 '-POS
'                            txtPrefix_POS.text = rec!prefix
'                            txtNextNumber_POS.text = rec!nextnumber
'                        Case 4 '-Warehouse
'                            txtPrefix_CA.text = rec!prefix
'                            txtNextNumber_CA.text = rec!nextnumber
'                        Case 5 '-POS
'                            txtPrefix_CA1.text = rec!prefix
'                            txtNextNumber_CA1.text = rec!nextnumber
'                    End Select
'                    rec.MoveNext
'                Loop
'            End If
'    End Select
'End Sub
'
'Private Sub btnAdd_Click()
'    If IsNumeric(txtUserNumber.text) = False Then
'        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(44)
'        GLOBAL_MessageFrm.Show (1)
'        txtUserNumber.SetFocus
'    ElseIf Trim(txtName.text) = "" Then
'        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(18)
'        GLOBAL_MessageFrm.Show (1)
'        txtName.SetFocus
'    Else
'        On Error GoTo ErrorHandler:
'        Set con = New ADODB.Connection
'        Set cmd = New ADODB.Command
'
'        con.ConnectionString = ConnString
'        con.Open
'        cmd.ActiveConnection = con
'        cmd.CommandType = adCmdStoredProc
'        cmd.CommandText = "BASE_User_Insert"
'
'        cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInputOutput, , 0)
'        cmd.Parameters.Append cmd.CreateParameter("@UserRoleId", adInteger, adParamInput, , cmbRoles.ItemData(cmbRoles.ListIndex))
'        cmd.Parameters.Append cmd.CreateParameter("@UserNumber", adInteger, adParamInput, , Val(txtUserNumber.text))
'        cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 50, txtName.text)
'        cmd.Parameters.Append cmd.CreateParameter("@Username", adVarChar, adParamInput, 50, txtName.text)
'        cmd.Parameters.Append cmd.CreateParameter("@Password", adVarChar, adParamInput, 50, "")
'
'        cmd.Execute
'
'        Dim item As MSComctlLib.ListItem
'        Set item = lvUsers.ListItems.add(, , "")
'            item.SubItems(1) = cmd.Parameters("@UserId")
'            item.SubItems(2) = txtUserNumber.text
'            item.SubItems(3) = txtName.text
'            item.SubItems(4) = cmbRoles.text
'            item.SubItems(5) = cmbRoles.ItemData(cmbRoles.ListIndex)
'            item.Checked = True
'
'        For Each item In lvUsers.ListItems
'            If item.SubItems(1) = cmd.Parameters("@UserId") Then
'                item.Selected = True
'                Exit For
'            End If
'        Next
'
'        txtName.text = ""
'        txtUserNumber.text = ""
'        lvUsers.SetFocus
'        con.Close
'    End If
'    Exit Sub
'ErrorHandler:
'    If IsNumeric(Err.Description) = True Then
'        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(Val(Err.Description))
'        If Err.Description = 47 Then txtUserNumber.SetFocus
'        If Err.Description = 48 Then txtName.SetFocus
'    Else
'        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & Err.Description
'    End If
'    GLOBAL_MessageFrm.Show (1)
'End Sub
'
'Private Sub btnBanks_Click()
'    FIN_BanksFrm.Show (1)
'End Sub
'
'Private Sub btnCancel_Click()
'    Unload Me
'End Sub
'
'
'
'Private Sub btnCompany_Click()
'    FRE_Company.Visible = True
'    FRE_DocNumbers.Visible = False
'    FRE_Users.Visible = False
'    FRE_AutoBackups.Visible = False
'    FRE_References.Visible = False
'    On Error Resume Next
'    txtCompanyName.SetFocus
'End Sub
'
'Private Sub btnDocuments_Click()
'    FRE_Company.Visible = False
'    FRE_DocNumbers.Visible = True
'    FRE_Users.Visible = False
'    FRE_AutoBackups.Visible = False
'    FRE_References.Visible = False
'    txtPrefix_SalesOrder.SetFocus
'End Sub
'
'Private Sub btnExpenses_Click()
'    FIN_ExpenseListFrm.Show (1)
'End Sub
'
'Private Sub btnFunds_Click()
'    FIN_FundsFrm.Show (1)
'End Sub
'
'Private Sub btnLocations_Click()
'    INV_LocationModFrm.Show (1)
'End Sub
'
'Private Sub btnPaymentMethod_Click()
'    BASE_PaymentMethodsFrm.Show (1)
'End Sub
'
'Private Sub btnReferences_Click()
'    FRE_Company.Visible = False
'    FRE_DocNumbers.Visible = False
'    FRE_Users.Visible = False
'    FRE_AutoBackups.Visible = False
'    FRE_References.Visible = True
'    btnPaymentMethod.SetFocus
'End Sub
'
'Private Sub btnRemove_Click()
'    If lvUsers.SelectedItem.SubItems(1) = 1 Then
'        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(46)
'        GLOBAL_MessageFrm.Show (1)
'        Exit Sub
'    End If
'
'    Set con = New ADODB.Connection
'    Set cmd = New ADODB.Command
'    con.ConnectionString = ConnString
'    con.Open
'    cmd.ActiveConnection = con
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandText = "BASE_User_Update"
'    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , lvUsers.SelectedItem.SubItems(1))
'    cmd.Parameters.Append cmd.CreateParameter("@RoleId", adInteger, adParamInput, , cmbRoles.ItemData(cmbRoles.ListIndex))
'    cmd.Parameters.Append cmd.CreateParameter("@Usernumber", adInteger, adParamInput, , lvUsers.SelectedItem.SubItems(2))
'    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 50, Null)
'    cmd.Parameters.Append cmd.CreateParameter("@Username", adVarChar, adParamInput, 50, Null)
'    cmd.Parameters.Append cmd.CreateParameter("@Password", adVarChar, adParamInput, 50, Null)
'    cmd.Parameters.Append cmd.CreateParameter("@Pin", adVarChar, adParamInput, 4, Null)
'    cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , "False")
'    cmd.Execute
'    con.Close
'    lvUsers.ListItems.Remove (lvUsers.SelectedItem.Index)
'End Sub
'
'Private Sub btnSave_Click()
'    'COMPANY
'    Set con = New ADODB.Connection
'    Set cmd = New ADODB.Command
'    con.ConnectionString = ConnString
'    con.Open
'    cmd.ActiveConnection = con
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandText = "BASE_Company_Update"
'    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, txtCompanyName.text)
'    cmd.Parameters.Append cmd.CreateParameter("@Phone", adVarChar, adParamInput, 50, txtPhone.text)
'    cmd.Parameters.Append cmd.CreateParameter("@Fax", adVarChar, adParamInput, 50, txtFax.text)
'    cmd.Parameters.Append cmd.CreateParameter("@Address1", adVarChar, adParamInput, 250, txtAddress1.text)
'    cmd.Parameters.Append cmd.CreateParameter("@Address2", adVarChar, adParamInput, 250, txtAddress2.text)
'    cmd.Parameters.Append cmd.CreateParameter("@Email", adVarChar, adParamInput, 50, txtEmail.text)
'    cmd.Parameters.Append cmd.CreateParameter("@Website", adVarChar, adParamInput, 50, txtWebsite.text)
'    cmd.Execute
'
'    'DOCUMENT FORMAT
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = con
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandText = "GLOBAL_DocNoFormat_Update"
'    cmd.Parameters.Append cmd.CreateParameter("@DocNoFormatId", adInteger, adParamInput, , 1)
'    cmd.Parameters.Append cmd.CreateParameter("@NextNumber", adInteger, adParamInput, , txtNextNumber_PurchaseOrder.text)
'    cmd.Parameters.Append cmd.CreateParameter("@Prefix", adVarChar, adParamInput, 50, txtPrefix_PurchaseOrder.text)
'    cmd.Execute
'
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = con
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandText = "GLOBAL_DocNoFormat_Update"
'    cmd.Parameters.Append cmd.CreateParameter("@DocNoFormatId", adInteger, adParamInput, , 2)
'    cmd.Parameters.Append cmd.CreateParameter("@NextNumber", adInteger, adParamInput, , txtNextNumber_SalesOrder.text)
'    cmd.Parameters.Append cmd.CreateParameter("@Prefix", adVarChar, adParamInput, 50, txtPrefix_SalesOrder.text)
'    cmd.Execute
'
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = con
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandText = "GLOBAL_DocNoFormat_Update"
'    cmd.Parameters.Append cmd.CreateParameter("@DocNoFormatId", adInteger, adParamInput, , 3)
'    cmd.Parameters.Append cmd.CreateParameter("@NextNumber", adInteger, adParamInput, , txtNextNumber_POS.text)
'    cmd.Parameters.Append cmd.CreateParameter("@Prefix", adVarChar, adParamInput, 50, txtPrefix_POS.text)
'    cmd.Execute
'
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = con
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandText = "GLOBAL_DocNoFormat_Update"
'    cmd.Parameters.Append cmd.CreateParameter("@DocNoFormatId", adInteger, adParamInput, , 4)
'    cmd.Parameters.Append cmd.CreateParameter("@NextNumber", adInteger, adParamInput, , txtNextNumber_CA.text)
'    cmd.Parameters.Append cmd.CreateParameter("@Prefix", adVarChar, adParamInput, 50, txtPrefix_CA.text)
'    cmd.Execute
'
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = con
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandText = "GLOBAL_DocNoFormat_Update"
'    cmd.Parameters.Append cmd.CreateParameter("@DocNoFormatId", adInteger, adParamInput, , 5)
'    cmd.Parameters.Append cmd.CreateParameter("@NextNumber", adInteger, adParamInput, , txtNextNumber_CA1.text)
'    cmd.Parameters.Append cmd.CreateParameter("@Prefix", adVarChar, adParamInput, 50, txtPrefix_CA1.text)
'    cmd.Execute
'
'    'Users
'    Dim item As MSComctlLib.ListItem
'    For Each item In lvUsers.ListItems
'        Set cmd = New ADODB.Command
'        cmd.ActiveConnection = con
'        cmd.CommandType = adCmdStoredProc
'        cmd.CommandText = "BASE_User_Update"
'        cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , item.SubItems(1))
'        cmd.Parameters.Append cmd.CreateParameter("@UserRoleId", adInteger, adParamInput, , item.SubItems(5))
'        cmd.Parameters.Append cmd.CreateParameter("@UserNumber", adInteger, adParamInput, , item.SubItems(2))
'        cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 50, Null)
'        cmd.Parameters.Append cmd.CreateParameter("@Username", adVarChar, adParamInput, 50, Null)
'        cmd.Parameters.Append cmd.CreateParameter("@Password", adVarChar, adParamInput, 50, Null)
'        cmd.Parameters.Append cmd.CreateParameter("@Pin", adVarChar, adParamInput, 4, Null)
'        If item.Checked = True Then
'            cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , "True")
'        Else
'            cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , "False")
'        End If
'        cmd.Execute
'    Next
'
'    con.Close
'    Unload Me
'End Sub
'
'
'
'Private Sub btnTerms_Click()
'    BASE_TermsFrm.Show (1)
'End Sub
'
'Private Sub btnUnits_Click()
'    BASE_UnitsFrm.Show (1)
'End Sub
'
'Private Sub btnUsers_Click()
'    FRE_Company.Visible = False
'    FRE_DocNumbers.Visible = False
'    FRE_Users.Visible = True
'    FRE_AutoBackups.Visible = False
'    FRE_References.Visible = False
'    lvUsers.SetFocus
'End Sub
'
'Private Sub btnWorkstations_Click()
'    SYS_WorkstationFrm.Show (1)
'End Sub
'
'Private Sub chkShow_Click()
'    Dim item As MSComctlLib.ListItem
'    Set rec = New ADODB.Recordset
'    Set rec = Global_Data("User")
'    lvUsers.ListItems.Clear
'    If Not rec.EOF Then
'        Do Until rec.EOF
'            If chkShow.value = 1 Then
'                Set item = lvUsers.ListItems.add(, , "")
'                    item.SubItems(1) = rec!UserId
'                    item.SubItems(2) = rec!UserNumber
'                    item.SubItems(3) = rec!Name
'                    item.SubItems(4) = rec!Role
'                    item.SubItems(5) = rec!UserRoleId
'
'                If rec!isActive = "True" Then item.Checked = True
'                lvUsers.ColumnHeaders(1).width = lvUsers.width * 0.04
'                lvUsers.ColumnHeaders(3).width = lvUsers.width * 0.15
'                lvUsers.ColumnHeaders(4).width = lvUsers.width * 0.52
'                lvUsers.ColumnHeaders(5).width = lvUsers.width * 0.25
'                btnRemove.Enabled = False
'            Else
'                If rec!isActive = "True" Then
'                    Set item = lvUsers.ListItems.add(, , "")
'                        item.SubItems(1) = rec!UserId
'                        item.SubItems(2) = rec!UserNumber
'                        item.SubItems(3) = rec!Name
'                        item.SubItems(4) = rec!Role
'                        item.SubItems(5) = rec!UserRoleId
'
'                    If rec!isActive = "True" Then item.Checked = True
'                    lvUsers.ColumnHeaders(1).width = lvUsers.width * 0
'                    lvUsers.ColumnHeaders(3).width = lvUsers.width * 0.15
'                    lvUsers.ColumnHeaders(4).width = lvUsers.width * 0.56
'                    lvUsers.ColumnHeaders(5).width = lvUsers.width * 0.25
'                End If
'                btnRemove.Enabled = True
'            End If
'            rec.MoveNext
'        Loop
'    End If
'End Sub
'
'Private Sub Command3_Click()
'    FRE_Company.Visible = False
'    FRE_DocNumbers.Visible = False
'    FRE_Users.Visible = False
'    FRE_AutoBackups.Visible = True
'    FRE_References.Visible = False
'End Sub
'
'Private Sub Form_Load()
'    Populate "Company"
'    Populate "User"
'    Populate "Documents"
'    Populate "UserRoles"
'
'    chkShow_Click
'    btnCompany_Click
'End Sub
'
'Private Sub lvUsers_DblClick()
'    If lvUsers.ListItems.Count > 0 Then
'        On Error Resume Next
'        BASE_UserRightsFrm.CheckUserId = lvUsers.SelectedItem.SubItems(1)
'        BASE_UserRightsFrm.UserRoleId = lvUsers.SelectedItem.SubItems(4)
'        BASE_UserRightsFrm.cmbRoles.text = lvUsers.SelectedItem.SubItems(3)
'        BASE_UserRightsFrm.txtUserNumber.text = lvUsers.SelectedItem.SubItems(2)
'        BASE_UserRightsFrm.Show (1)
'    End If
'End Sub
'
'Private Sub lvUsers_ItemCheck(ByVal item As MSComctlLib.ListItem)
'    If item.SubItems(1) = "1" Then item.Checked = True
'End Sub
'
'Private Sub txtName_GotFocus()
'    selectText txtName
'End Sub
'
'Private Sub txtNextNumber_CA_Change()
'    If IsNumeric(txtNextNumber_CA.text) = False Then
'        txtNextNumber_CA.text = "1"
'    Else
'        txtNextNumber_CA.text = Int(txtNextNumber_CA.text)
'    End If
'    lblPreview_CA.Caption = txtPrefix_CA.text & Format(txtNextNumber_CA.text, "000000")
'End Sub
'
'Private Sub txtNextNumber_CA1_Change()
'    If IsNumeric(txtNextNumber_CA1.text) = False Then
'        txtNextNumber_CA1.text = "1"
'    Else
'        txtNextNumber_CA1.text = Int(txtNextNumber_CA1.text)
'    End If
'    lblPreview_CA1.Caption = txtPrefix_CA1.text & Format(txtNextNumber_CA1.text, "000000")
'End Sub
'
'Private Sub txtNextNumber_POS_Change()
'    If IsNumeric(txtNextNumber_POS.text) = False Then
'        txtNextNumber_POS.text = "1"
'    Else
'        txtNextNumber_POS.text = Int(txtNextNumber_POS.text)
'    End If
'    lblPreview_POS.Caption = txtPrefix_POS.text & Format(txtNextNumber_POS.text, "000000")
'End Sub
'
'Private Sub txtNextNumber_POS_GotFocus()
'    selectText txtNextNumber_POS
'End Sub
'
'Private Sub txtNextNumber_PurchaseOrder_Change()
'    If IsNumeric(txtNextNumber_PurchaseOrder.text) = False Then
'        txtNextNumber_PurchaseOrder.text = "1"
'    Else
'        txtNextNumber_PurchaseOrder.text = Int(txtNextNumber_PurchaseOrder.text)
'    End If
'    lblPreview_PurchaseOrder.Caption = txtPrefix_PurchaseOrder.text & Format(txtNextNumber_PurchaseOrder.text, "000000")
'End Sub
'
'Private Sub txtNextNumber_PurchaseOrder_GotFocus()
'    selectText txtNextNumber_PurchaseOrder
'End Sub
'
'Private Sub txtNextNumber_SalesOrder_Change()
'    If IsNumeric(txtNextNumber_SalesOrder.text) = False Then
'        txtNextNumber_SalesOrder.text = "1"
'    Else
'        txtNextNumber_SalesOrder.text = Int(txtNextNumber_SalesOrder.text)
'    End If
'    lblPreview_SalesOrder.Caption = txtPrefix_SalesOrder.text & Format(txtNextNumber_SalesOrder.text, "000000")
'End Sub
'
'Private Sub txtNextNumber_SalesOrder_GotFocus()
'    selectText txtNextNumber_SalesOrder
'End Sub
'
'Private Sub txtPrefix_CA_Change()
'     lblPreview_CA.Caption = txtPrefix_CA.text & Format(txtNextNumber_CA.text, "000000")
'End Sub
'
'Private Sub txtPrefix_CA1_Change()
'    lblPreview_CA1.Caption = txtPrefix_CA1.text & Format(txtNextNumber_CA1.text, "000000")
'End Sub
'
'Private Sub txtPrefix_POS_Change()
'    lblPreview_POS.Caption = txtPrefix_POS.text & Format(txtNextNumber_POS.text, "000000")
'End Sub
'
'Private Sub txtPrefix_POS_GotFocus()
'    selectText txtPrefix_POS
'End Sub
'
'Private Sub txtPrefix_PurchaseOrder_Change()
'    lblPreview_PurchaseOrder.Caption = txtPrefix_PurchaseOrder.text & Format(txtNextNumber_PurchaseOrder.text, "000000")
'End Sub
'
'Private Sub txtPrefix_PurchaseOrder_GotFocus()
'    selectText txtPrefix_PurchaseOrder
'End Sub
'
'Private Sub txtPrefix_SalesOrder_Change()
'    lblPreview_SalesOrder.Caption = txtPrefix_SalesOrder.text & Format(txtNextNumber_SalesOrder.text, "000000")
'End Sub
'
'Private Sub txtPrefix_SalesOrder_GotFocus()
'    selectText txtPrefix_SalesOrder
'End Sub
'
'Private Sub txtUserNumber_GotFocus()
'    selectText txtUserNumber
'End Sub
Private Sub TabStrip1_Click()

End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnDiscounts_Click()
    BASE_DiscountSettingsFrm.Show (1)
End Sub

Private Sub btnPaymentMethod_Click()
    POS_DisplayFrm.Show (1)
End Sub

Private Sub btnPOS_Click()
    FRE_POS.Visible = True
    FRE_Inventory.Visible = False
End Sub

Private Sub btnSave_Click()
    Unload Me
End Sub

Private Sub btnTerms_Click()
    POS_UserValidationFrm.Show (1)
End Sub

Private Sub btnTransferStock_Click()
    FRE_POS.Visible = False
    FRE_Inventory.Visible = True
End Sub

Private Sub btnUnits_Click()
    BASE_LoyaltyCardSettingsFrm.Show (1)
End Sub

Private Sub Command10_Click()
    INV_AlertSettingsFrm.Show (1)
End Sub

Private Sub Command9_Click()
    x = MsgBox("Are you sure you want to remove expired product information?", vbQuestion + vbYesNo)
    If x = vbYes Then
        Dim con As New ADODB.Connection
        con.ConnectionString = ConnString
        con.Open
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "INV_ProductExpiry_RemoveExpired"
        cmd.Execute
        con.Close
        
        MsgBox "Expired product record information removed successfully!", vbInformation
    End If
End Sub

Private Sub Form_Load()
    btnPOS_Click
End Sub
