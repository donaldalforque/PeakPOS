VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form BASE_GeneralSettingsFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PeakPOS - General Settings"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   Icon            =   "BASE_GeneralSettingsFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   0
      TabIndex        =   54
      Top             =   -120
      Width           =   1695
      Begin VB.CommandButton btnReset 
         Caption         =   "Data Reset"
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
         Picture         =   "BASE_GeneralSettingsFrm.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   5160
         Width           =   1680
      End
      Begin VB.CommandButton btnDataImport 
         Caption         =   "Data Import"
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
         Picture         =   "BASE_GeneralSettingsFrm.frx":6EED
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   4320
         Width           =   1680
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Backups"
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
         Picture         =   "BASE_GeneralSettingsFrm.frx":74F3
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   3480
         Width           =   1680
      End
      Begin VB.CommandButton btnDocuments 
         Caption         =   "Doc. Numbers"
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
         Picture         =   "BASE_GeneralSettingsFrm.frx":7B61
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   2640
         Width           =   1680
      End
      Begin VB.CommandButton btnReferences 
         Caption         =   "References"
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
         Picture         =   "BASE_GeneralSettingsFrm.frx":818D
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   1800
         Width           =   1680
      End
      Begin VB.CommandButton btnUsers 
         Caption         =   "Users"
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
         Picture         =   "BASE_GeneralSettingsFrm.frx":87C6
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   960
         Width           =   1680
      End
      Begin VB.CommandButton btnCompany 
         Caption         =   "Company"
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
         Picture         =   "BASE_GeneralSettingsFrm.frx":8DDE
         Style           =   1  'Graphical
         TabIndex        =   55
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
      Begin VB.Frame FRE_Import 
         BackColor       =   &H00FFFFFF&
         Height          =   6375
         Left            =   120
         TabIndex        =   132
         Top             =   0
         Width           =   7695
         Begin VB.Frame FRE_Import_Details 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Import Products"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2655
            Left            =   240
            TabIndex        =   147
            Top             =   3480
            Visible         =   0   'False
            Width           =   7335
            Begin VB.CommandButton btnExecute 
               Caption         =   "Execute"
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
               Left            =   5880
               TabIndex        =   156
               Top             =   2040
               Width           =   1335
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Action for duplicate items"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   120
               TabIndex        =   152
               Top             =   840
               Width           =   7095
               Begin VB.OptionButton chkCancel 
                  BackColor       =   &H00FFFFFF&
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
                  Height          =   225
                  Left            =   2880
                  TabIndex        =   155
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   2295
               End
               Begin VB.OptionButton chkIgnore 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Ignore"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   120
                  TabIndex        =   154
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin VB.OptionButton chkOverwrite 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Overwrite"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   1320
                  TabIndex        =   153
                  Top             =   360
                  Width           =   1695
               End
            End
            Begin MSComctlLib.ProgressBar ProgressBar 
               Height          =   375
               Left            =   120
               TabIndex        =   150
               Top             =   2040
               Width           =   5655
               _ExtentX        =   9975
               _ExtentY        =   661
               _Version        =   393216
               BorderStyle     =   1
               Appearance      =   0
               Max             =   1000
               Scrolling       =   1
            End
            Begin VB.CommandButton btnLoadCSV 
               Caption         =   "Load CSV File"
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
               Left            =   120
               TabIndex        =   148
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label lblProgressDetails 
               BackStyle       =   0  'Transparent
               Caption         =   "Import progress.."
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   120
               TabIndex        =   151
               Top             =   1800
               Width           =   5295
            End
            Begin VB.Label Label49 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "View Status Report"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   210
               Left            =   5640
               MouseIcon       =   "BASE_GeneralSettingsFrm.frx":93BF
               MousePointer    =   99  'Custom
               TabIndex        =   157
               Top             =   1680
               Width           =   1545
            End
            Begin VB.Label lblPath 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   1800
               TabIndex        =   149
               Top             =   420
               Width           =   5415
            End
         End
         Begin MSComDlg.CommonDialog DialogBox 
            Left            =   7200
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Filter          =   "*.txt, *.csv"
         End
         Begin VB.CommandButton btnImportProduct 
            Caption         =   "Import Products"
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
            TabIndex        =   136
            Top             =   1440
            Width           =   2175
         End
         Begin VB.CommandButton Command7 
            Appearance      =   0  'Flat
            Caption         =   "Import Sales Orders"
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
            Height          =   375
            Left            =   240
            TabIndex        =   135
            Top             =   2400
            Width           =   2175
         End
         Begin VB.CommandButton Command6 
            Appearance      =   0  'Flat
            Caption         =   "Import Purchase Orders"
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
            Height          =   375
            Left            =   240
            TabIndex        =   134
            Top             =   2880
            Width           =   2175
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            Caption         =   "Import Customers"
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
            Height          =   375
            Left            =   240
            TabIndex        =   133
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label lblcsv_purchaseorders 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Download Sales Order CSV Template"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   4560
            MouseIcon       =   "BASE_GeneralSettingsFrm.frx":9511
            MousePointer    =   99  'Custom
            TabIndex        =   146
            Top             =   2940
            Width           =   2985
         End
         Begin VB.Label lblcsv_salesorders 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Download Sales Order CSV Template"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   4560
            MouseIcon       =   "BASE_GeneralSettingsFrm.frx":9663
            MousePointer    =   99  'Custom
            TabIndex        =   145
            Top             =   2460
            Width           =   2985
         End
         Begin VB.Label lblcsv_customers 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Download Customer CSV Template"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   4560
            MouseIcon       =   "BASE_GeneralSettingsFrm.frx":97B5
            MousePointer    =   99  'Custom
            TabIndex        =   144
            Top             =   1980
            Width           =   2805
         End
         Begin VB.Label lblcsv_products 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Download Product CSV Template"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   4560
            MouseIcon       =   "BASE_GeneralSettingsFrm.frx":9907
            MousePointer    =   99  'Custom
            TabIndex        =   143
            Top             =   1560
            Width           =   2640
         End
         Begin VB.Label Label60 
            BackStyle       =   0  'Transparent
            Caption         =   "Allows you to import product lists, sales orders and customer list using CSV file formats."
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
            TabIndex        =   142
            Top             =   720
            Width           =   7215
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Import"
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
            TabIndex        =   141
            Top             =   240
            Width           =   1410
         End
         Begin VB.Label Label58 
            BackStyle       =   0  'Transparent
            Caption         =   "Import sales orders"
            Enabled         =   0   'False
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
            Left            =   2520
            TabIndex        =   140
            Top             =   2460
            Width           =   1815
         End
         Begin VB.Label Label57 
            BackStyle       =   0  'Transparent
            Caption         =   "Import product list"
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
            Left            =   2520
            TabIndex        =   139
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label Label55 
            BackStyle       =   0  'Transparent
            Caption         =   "Import purchase orders"
            Enabled         =   0   'False
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
            Left            =   2520
            TabIndex        =   138
            Top             =   2940
            Width           =   2055
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "Import customer list"
            Enabled         =   0   'False
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
            Left            =   2520
            TabIndex        =   137
            Top             =   1980
            Width           =   4935
         End
      End
      Begin VB.Frame FRE_AutoBackups 
         BackColor       =   &H00FFFFFF&
         Height          =   6375
         Left            =   120
         TabIndex        =   77
         Top             =   0
         Width           =   7695
         Begin VB.CommandButton btnBackup 
            Caption         =   "Backup"
            Height          =   375
            Left            =   240
            TabIndex        =   100
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Backups"
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
            TabIndex        =   79
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Create backups for your data to ensure data security in case of hardware failure."
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
            TabIndex        =   78
            Top             =   720
            Width           =   6135
         End
      End
      Begin VB.Frame Fre_Reset 
         BackColor       =   &H00FFFFFF&
         Height          =   6375
         Left            =   120
         TabIndex        =   104
         Top             =   0
         Width           =   7695
         Begin VB.CommandButton btnPOSOrderSlip 
            Appearance      =   0  'Flat
            Caption         =   "Reset POS Order Data"
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
            TabIndex        =   119
            Top             =   1920
            Width           =   2175
         End
         Begin VB.CommandButton btnResetAll 
            Appearance      =   0  'Flat
            Caption         =   "Reset All"
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
            TabIndex        =   117
            Top             =   3840
            Width           =   2175
         End
         Begin VB.CommandButton btnMasterReset 
            Appearance      =   0  'Flat
            Caption         =   "Master Data Reset"
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
            TabIndex        =   115
            Top             =   4560
            Width           =   2175
         End
         Begin VB.CommandButton btnResetPurchases 
            Appearance      =   0  'Flat
            Caption         =   "Reset Purchases Data"
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
            TabIndex        =   113
            Top             =   3360
            Width           =   2175
         End
         Begin VB.CommandButton btnResetSalesOrders 
            Appearance      =   0  'Flat
            Caption         =   "Reset Sales Order Data"
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
            TabIndex        =   111
            Top             =   2880
            Width           =   2175
         End
         Begin VB.CommandButton btnResetInventory 
            Appearance      =   0  'Flat
            Caption         =   "Reset Inventory Data"
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
            TabIndex        =   108
            Top             =   2400
            Width           =   2175
         End
         Begin VB.CommandButton btnPOSReset 
            Caption         =   "Reset POS Sales Data"
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
            TabIndex        =   105
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label Label46 
            BackStyle       =   0  'Transparent
            Caption         =   "Reset POS Order data such as hold list or POS Order slip."
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
            Left            =   2520
            TabIndex        =   120
            Top             =   1980
            Width           =   4935
         End
         Begin VB.Label Label45 
            BackStyle       =   0  'Transparent
            Caption         =   "Reset Purchasing, Sales, Inventory and POS data."
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
            Left            =   2520
            TabIndex        =   118
            Top             =   3900
            Width           =   4935
         End
         Begin VB.Line Line1 
            X1              =   240
            X2              =   7320
            Y1              =   4440
            Y2              =   4440
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            Caption         =   "Reset all system data to default settings."
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
            Left            =   2520
            TabIndex        =   116
            Top             =   4620
            Width           =   4935
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Caption         =   "Reset all Purchasing data including invoices and accounts."
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
            Left            =   2520
            TabIndex        =   114
            Top             =   3420
            Width           =   4935
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   "Reset all Sales Order data including invoices and accounts."
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
            Left            =   2520
            TabIndex        =   112
            Top             =   2940
            Width           =   4935
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            Caption         =   "Reset all POS sales records and transactions."
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
            Left            =   2520
            TabIndex        =   110
            Top             =   1560
            Width           =   4935
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "Revert all inventory records including transfers and orders."
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
            Left            =   2520
            TabIndex        =   109
            Top             =   2460
            Width           =   4935
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Reset"
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
            TabIndex        =   107
            Top             =   240
            Width           =   1230
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "This will restore and reset all data to the default settings. WARNING! This process cannot be undone!"
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
            TabIndex        =   106
            Top             =   720
            Width           =   5055
         End
      End
      Begin VB.Frame FRE_DocNumbers 
         BackColor       =   &H00FFFFFF&
         Height          =   6375
         Left            =   120
         TabIndex        =   80
         Top             =   0
         Width           =   7695
         Begin VB.TextBox txtPrefix_AuditStock 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2640
            TabIndex        =   34
            Top             =   5040
            Width           =   855
         End
         Begin VB.TextBox txtNextNumber_AuditStock 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3600
            TabIndex        =   35
            Top             =   5040
            Width           =   1455
         End
         Begin VB.TextBox txtPrefix_NewStock 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2640
            TabIndex        =   32
            Top             =   4680
            Width           =   855
         End
         Begin VB.TextBox txtNextNumber_NewStock 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3600
            TabIndex        =   33
            Top             =   4680
            Width           =   1455
         End
         Begin VB.TextBox txtPrefix_SalesAdjustment 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2640
            TabIndex        =   30
            Top             =   4320
            Width           =   855
         End
         Begin VB.TextBox txtNextNumber_SalesAdjustment 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3600
            TabIndex        =   31
            Top             =   4320
            Width           =   1455
         End
         Begin VB.TextBox txtPrefix_SalesReturn 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2640
            TabIndex        =   28
            Top             =   3960
            Width           =   855
         End
         Begin VB.TextBox txtNextNumber_SalesReturn 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3600
            TabIndex        =   29
            Top             =   3960
            Width           =   1455
         End
         Begin VB.TextBox txtPrefix_PurchaseReturn 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2640
            TabIndex        =   26
            Top             =   3600
            Width           =   855
         End
         Begin VB.TextBox txtNextNumber_PurchaseReturn 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3600
            TabIndex        =   27
            Top             =   3600
            Width           =   1455
         End
         Begin VB.TextBox txtNextNumber_CA1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3600
            TabIndex        =   37
            Top             =   5400
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtPrefix_CA1 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2640
            TabIndex        =   36
            Top             =   5400
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtPrefix_TransferStock 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2640
            TabIndex        =   24
            Top             =   3240
            Width           =   855
         End
         Begin VB.TextBox txtNextNumber_TransferStock 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3600
            TabIndex        =   25
            Top             =   3240
            Width           =   1455
         End
         Begin VB.TextBox txtNextNumber_POS 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3600
            TabIndex        =   23
            Top             =   2880
            Width           =   1455
         End
         Begin VB.TextBox txtPrefix_POS 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2640
            TabIndex        =   22
            Top             =   2880
            Width           =   855
         End
         Begin VB.TextBox txtPrefix_PurchaseOrder 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2640
            TabIndex        =   20
            Top             =   2520
            Width           =   855
         End
         Begin VB.TextBox txtNextNumber_PurchaseOrder 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3600
            TabIndex        =   21
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox txtNextNumber_SalesOrder 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3600
            TabIndex        =   19
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox txtPrefix_SalesOrder 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2640
            TabIndex        =   18
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Audit Stock"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   130
            Top             =   5040
            Width           =   1050
         End
         Begin VB.Label lblPreview_AuditStock 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5160
            TabIndex        =   129
            Top             =   5040
            Width           =   2205
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New Stock"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   128
            Top             =   4680
            Width           =   975
         End
         Begin VB.Label lblPreview_NewStock 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5160
            TabIndex        =   127
            Top             =   4680
            Width           =   2205
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Adjustment"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   126
            Top             =   4320
            Width           =   1620
         End
         Begin VB.Label lblPreview_SalesAdjustment 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5160
            TabIndex        =   125
            Top             =   4320
            Width           =   2205
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Return"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   124
            Top             =   3960
            Width           =   1155
         End
         Begin VB.Label lblPreview_SalesReturn 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5160
            TabIndex        =   123
            Top             =   3960
            Width           =   2205
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Return"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   122
            Top             =   3600
            Width           =   1515
         End
         Begin VB.Label lblPreview_PurchaseReturn 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5160
            TabIndex        =   121
            Top             =   3600
            Width           =   2205
         End
         Begin VB.Label lblPreview_CA1 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5160
            TabIndex        =   95
            Top             =   5400
            Visible         =   0   'False
            Width           =   2205
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cash Advance"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   94
            Top             =   5400
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label lblPreview_TransferStock 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5160
            TabIndex        =   93
            Top             =   3240
            Width           =   2205
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transfer Stock"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   92
            Top             =   3240
            Width           =   1305
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Point of Sale"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   91
            Top             =   2880
            Width           =   1170
         End
         Begin VB.Label lblPreview_POS 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5160
            TabIndex        =   90
            Top             =   2880
            Width           =   2205
         End
         Begin VB.Label lblPreview_PurchaseOrder 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5160
            TabIndex        =   89
            Top             =   2520
            Width           =   2205
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Order"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   88
            Top             =   2520
            Width           =   1425
         End
         Begin VB.Label lblPreview_SalesOrder 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5160
            TabIndex        =   87
            Top             =   2160
            Width           =   2205
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5160
            TabIndex        =   86
            Top             =   1560
            Width           =   765
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Next Number"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3600
            TabIndex        =   85
            Top             =   1560
            Width           =   1260
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prefix"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2640
            TabIndex        =   84
            Top             =   1560
            Width           =   555
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Order"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   83
            Top             =   2160
            Width           =   1065
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Document Numbers"
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
            TabIndex        =   82
            Top             =   240
            Width           =   2325
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Set the pattern for documents numbers here. You can attached prefix to the numbers and can see preview on how it will look."
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
            TabIndex        =   81
            Top             =   720
            Width           =   6135
         End
      End
      Begin VB.Frame FRE_Company 
         BackColor       =   &H00FFFFFF&
         Height          =   6375
         Left            =   120
         TabIndex        =   47
         Top             =   0
         Width           =   7695
         Begin VB.TextBox txtWebsite 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            TabIndex        =   7
            Top             =   3840
            Width           =   3735
         End
         Begin VB.TextBox txtEmail 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            TabIndex        =   6
            Top             =   3480
            Width           =   3735
         End
         Begin VB.TextBox txtFax 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            TabIndex        =   5
            Top             =   3120
            Width           =   2295
         End
         Begin VB.TextBox txtPhone 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            TabIndex        =   4
            Top             =   2760
            Width           =   2295
         End
         Begin VB.TextBox txtAddress2 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            TabIndex        =   3
            Top             =   2400
            Width           =   5055
         End
         Begin VB.TextBox txtAddress1 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            TabIndex        =   2
            Top             =   2040
            Width           =   5055
         End
         Begin VB.TextBox txtCompanyName 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            TabIndex        =   1
            Top             =   1680
            Width           =   5055
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SN #"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   61
            Top             =   3840
            Width           =   405
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VAT-REG TIN #"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   58
            Top             =   3480
            Width           =   1320
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Company Name"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   53
            Top             =   1680
            Width           =   1470
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Website"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   52
            Top             =   3135
            Width           =   780
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   51
            Top             =   2760
            Width           =   600
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Add and save your company profile including contact numbers, websites which will be displayed on your invoices and quotes."
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
            TabIndex        =   50
            Top             =   720
            Width           =   6135
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Company Info"
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
            TabIndex        =   49
            Top             =   240
            Width           =   1635
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   240
            TabIndex        =   48
            Top             =   2040
            Width           =   750
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
         TabIndex        =   45
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
         TabIndex        =   46
         Top             =   6480
         Width           =   1335
      End
      Begin VB.Frame FRE_Users 
         BackColor       =   &H00FFFFFF&
         Height          =   6375
         Left            =   120
         TabIndex        =   62
         Top             =   0
         Width           =   7695
         Begin VB.TextBox txtUserNumber 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1440
            MaxLength       =   4
            TabIndex        =   39
            Top             =   4920
            Width           =   855
         End
         Begin VB.ComboBox cmbRoles 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5400
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   4920
            Width           =   2055
         End
         Begin VB.CheckBox chkShow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Show All"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6480
            TabIndex        =   44
            Top             =   1200
            Width           =   1000
         End
         Begin VB.CommandButton btnRemove 
            Caption         =   "Remove"
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
            Left            =   6120
            TabIndex        =   43
            Top             =   5880
            Width           =   1335
         End
         Begin VB.CommandButton btnAdd 
            Caption         =   "Add"
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
            Left            =   4680
            TabIndex        =   42
            Top             =   5880
            Width           =   1335
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1440
            TabIndex        =   40
            Top             =   5280
            Width           =   3135
         End
         Begin MSComctlLib.ListView lvUsers 
            Height          =   3255
            Left            =   240
            TabIndex        =   38
            Top             =   1560
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   5741
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   529
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "UserId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "User No."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "User"
               Object.Width           =   10583
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Role"
               Object.Width           =   1199
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "UserRoleId"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label lblUserRoles 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Create User Roles"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   195
            Left            =   6120
            MouseIcon       =   "BASE_GeneralSettingsFrm.frx":9A59
            MousePointer    =   99  'Custom
            TabIndex        =   102
            Top             =   5280
            Width           =   1260
         End
         Begin VB.Label lblDivider 
            Height          =   15
            Left            =   240
            TabIndex        =   101
            Top             =   5760
            Width           =   7215
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Number:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   240
            TabIndex        =   98
            Top             =   4935
            Width           =   1110
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   240
            TabIndex        =   97
            Top             =   5295
            Width           =   510
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Role:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   4920
            TabIndex        =   96
            Top             =   4935
            Width           =   405
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Set accounts for multiple individuals and limit their access rights."
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
            TabIndex        =   64
            Top             =   720
            Width           =   6135
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Users"
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
            TabIndex        =   63
            Top             =   240
            Width           =   660
         End
      End
      Begin VB.Frame FRE_References 
         BackColor       =   &H00FFFFFF&
         Height          =   6375
         Left            =   120
         TabIndex        =   65
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
            TabIndex        =   17
            Top             =   5760
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
            TabIndex        =   16
            Top             =   5280
            Width           =   2175
         End
         Begin VB.CommandButton btnFunds 
            Appearance      =   0  'Flat
            Caption         =   "Warehouse Personnel"
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
            TabIndex        =   15
            Top             =   4800
            Width           =   2175
         End
         Begin VB.CommandButton btnBanks 
            Appearance      =   0  'Flat
            Caption         =   "City"
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
            TabIndex        =   14
            Top             =   4320
            Width           =   2175
         End
         Begin VB.CommandButton btnTax 
            Appearance      =   0  'Flat
            Caption         =   "Tax"
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
            TabIndex        =   13
            Top             =   3840
            Width           =   2175
         End
         Begin VB.CommandButton btnPricingScheme 
            Appearance      =   0  'Flat
            Caption         =   "Pricing Scheme"
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
            TabIndex        =   12
            Top             =   3360
            Width           =   2175
         End
         Begin VB.CommandButton btnLocations 
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
            TabIndex        =   11
            Top             =   2880
            Width           =   2175
         End
         Begin VB.CommandButton btnUnits 
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
            TabIndex        =   10
            Top             =   2400
            Width           =   2175
         End
         Begin VB.CommandButton btnTerms 
            Appearance      =   0  'Flat
            Caption         =   "Terms"
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
            Top             =   1920
            Width           =   2175
         End
         Begin VB.CommandButton btnPaymentMethod 
            Appearance      =   0  'Flat
            Caption         =   "Payment Methods"
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
            TabIndex        =   99
            Top             =   5820
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
            TabIndex        =   76
            Top             =   5340
            Width           =   4695
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Warehouse personnel for handling product transfers."
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
            TabIndex        =   75
            Top             =   4860
            Width           =   4695
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Cities for better customer tagging and searching."
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
            TabIndex        =   74
            Top             =   4380
            Width           =   4695
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Product tax codes."
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
            TabIndex        =   73
            Top             =   3900
            Width           =   4695
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Pricing schemes for products."
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
            TabIndex        =   72
            Top             =   3420
            Width           =   4695
         End
         Begin VB.Label Label16 
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
            TabIndex        =   71
            Top             =   2940
            Width           =   4695
         End
         Begin VB.Label Label7 
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
            TabIndex        =   70
            Top             =   2460
            Width           =   4695
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Payment terms for sales orders and payments."
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
            TabIndex        =   69
            Top             =   1980
            Width           =   4695
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Payment methods for orders, invoices and payments."
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
            TabIndex        =   68
            Top             =   1500
            Width           =   4695
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "References"
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
            TabIndex        =   67
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "You can save transaction references such as payment terms, inventory locations, purchases and more."
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
            TabIndex        =   66
            Top             =   720
            Width           =   7215
         End
      End
   End
End
Attribute VB_Name = "BASE_GeneralSettingsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UserId As Integer
Public Sub Populate(ByVal data As String)
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "Company"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Company")
            If Not rec.EOF Then
                If IsNull(rec!Name) = False Then txtCompanyName.Text = rec!Name
                If IsNull(rec!address1) = False Then txtAddress1.Text = rec!address1
                If IsNull(rec!address2) = False Then txtAddress2.Text = rec!address2
                If IsNull(rec!email) = False Then txtEmail.Text = rec!email
                If IsNull(rec!Phone) = False Then txtPhone.Text = rec!Phone
                If IsNull(rec!fax) = False Then txtFax.Text = rec!fax
                If IsNull(rec!website) = False Then txtWebsite.Text = rec!website
            End If
        Case "User"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("User")
            lvUsers.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        Set item = lvUsers.ListItems.add(, , "")
                            item.SubItems(1) = rec!UserId
                            item.SubItems(2) = rec!UserNumber
                            item.SubItems(3) = rec!Name
                            item.SubItems(4) = rec!role
                            item.SubItems(5) = rec!UserRoleId
                            item.Checked = True
                    End If
                    rec.MoveNext
                Loop
            End If
        Case "UserRoles"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("UserRoles")
            cmbRoles.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!UserRoleId = 1 Then rec.MoveNext
                    cmbRoles.AddItem rec!role
                    cmbRoles.ItemData(cmbRoles.NewIndex) = rec!UserRoleId
                    rec.MoveNext
                Loop
            End If
            cmbRoles.ListIndex = 0
        Case "Documents"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Documents")
            If Not rec.EOF Then
                Do Until rec.EOF
                    Select Case rec!DocNoFormatId
                        Case 1 '-Purchase Order
                            txtPrefix_PurchaseOrder.Text = rec!prefix
                            txtNextNumber_PurchaseOrder.Text = rec!nextnumber
                        Case 2 '-Sales Order
                            txtPrefix_SalesOrder.Text = rec!prefix
                            txtNextNumber_SalesOrder.Text = rec!nextnumber
                        Case 3 '-POS
                            txtPrefix_POS.Text = rec!prefix
                            txtNextNumber_POS.Text = rec!nextnumber
                        Case 4 '-Transfer Stock
                            txtPrefix_TransferStock.Text = rec!prefix
                            txtNextNumber_TransferStock.Text = rec!nextnumber
                        Case 5 '-Cash Advance
                            txtPrefix_CA1.Text = rec!prefix
                            txtNextNumber_CA1.Text = rec!nextnumber
                        Case 6 '-Purchase Return
                            txtPrefix_PurchaseReturn.Text = rec!prefix
                            txtNextNumber_PurchaseReturn.Text = rec!nextnumber
                        Case 7 '-Sales Return
                            txtPrefix_SalesReturn.Text = rec!prefix
                            txtNextNumber_SalesReturn.Text = rec!nextnumber
                        Case 8 '-New Stock
                            txtPrefix_NewStock.Text = rec!prefix
                            txtNextNumber_NewStock.Text = rec!nextnumber
                        Case 9 '-Audit Stock
                            txtPrefix_AuditStock.Text = rec!prefix
                            txtNextNumber_AuditStock.Text = rec!nextnumber
                        Case 10 '-Sales Adjustment
                            txtPrefix_SalesAdjustment.Text = rec!prefix
                            txtNextNumber_SalesAdjustment.Text = rec!nextnumber
                    End Select
                    rec.MoveNext
                Loop
            End If
    End Select
End Sub

Private Sub btnAdd_Click()
    If EditAccessRights(26) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    If IsNumeric(txtUserNumber.Text) = False Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(44)
        GLOBAL_MessageFrm.Show (1)
        txtUserNumber.SetFocus
    ElseIf Trim(txtName.Text) = "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(18)
        GLOBAL_MessageFrm.Show (1)
        txtName.SetFocus
    Else
        On Error GoTo ErrorHandler:
        Set con = New ADODB.Connection
        Set cmd = New ADODB.Command
        
        con.ConnectionString = ConnString
        con.Open
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "BASE_User_Insert"
        
        cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInputOutput, , 0)
        cmd.Parameters.Append cmd.CreateParameter("@UserRoleId", adInteger, adParamInput, , cmbRoles.ItemData(cmbRoles.ListIndex))
        cmd.Parameters.Append cmd.CreateParameter("@UserNumber", adInteger, adParamInput, , Val(txtUserNumber.Text))
        cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 50, txtName.Text)
        cmd.Parameters.Append cmd.CreateParameter("@Username", adVarChar, adParamInput, 50, txtName.Text)
        cmd.Parameters.Append cmd.CreateParameter("@Password", adVarChar, adParamInput, 50, "")
        
        cmd.Execute
        
        Dim item As MSComctlLib.ListItem
        Set item = lvUsers.ListItems.add(, , "")
            item.SubItems(1) = cmd.Parameters("@UserId")
            item.SubItems(2) = txtUserNumber.Text
            item.SubItems(3) = txtName.Text
            item.SubItems(4) = cmbRoles.Text
            item.SubItems(5) = cmbRoles.ItemData(cmbRoles.ListIndex)
            item.Checked = True
        
        For Each item In lvUsers.ListItems
            If item.SubItems(1) = cmd.Parameters("@UserId") Then
                item.Selected = True
                Exit For
            End If
        Next
        
        txtName.Text = ""
        txtUserNumber.Text = ""
        lvUsers.SetFocus
        con.Close
    End If
    Exit Sub
ErrorHandler:
    If IsNumeric(Err.Description) = True Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(Val(Err.Description))
        If Err.Description = 47 Then txtUserNumber.SetFocus
        If Err.Description = 48 Then txtName.SetFocus
    Else
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & Err.Description
    End If
    GLOBAL_MessageFrm.Show (1)
End Sub

Private Sub btnBackup_Click()
    Dim x As Variant
    x = MsgBox("The system will now backup the database. It may take a while. Please do not close the system until backup is successful. Proceed?", vbExclamation + vbOKCancel)
    If x = vbCancel Then Exit Sub
    Dim con As New ADODB.Connection
    Dim bName As String
    Set cmd = New ADODB.Command
    Dim sql As String
    
    sql = "BACKUP DATABASE Peak " & _
          "TO DISK = '" & App.path & "'\Backup\Peak.bak'" & _
          "WITH FORMAT," & _
          "MEDIANAME = 'PeakBackup'," & _
          "NAME = 'Full Backup of Peak';"
    
    con.ConnectionString = ConnString
    con.Open
    bName = App.path & "\Backup\Backup.bak"
    con.Execute "BACKUP DATABASE Peak TO DISK='" & bName & "' WITH INIT"
    con.Close
    MsgBox "Backup successful!", vbInformation
End Sub

Private Sub btnBanks_Click()
    BASE_CityFrm.Show (1)
End Sub

Private Sub btnCancel_Click()
    Unload Me
    Set BASE_ContainerFrm = Nothing
End Sub



Private Sub btnCompany_Click()
    FRE_Company.Visible = True
    FRE_DocNumbers.Visible = False
    FRE_Users.Visible = False
    FRE_AutoBackups.Visible = False
    FRE_References.Visible = False
    Fre_Reset.Visible = False
    FRE_Import.Visible = False
    On Error Resume Next
    txtCompanyName.SetFocus
End Sub

Private Sub btnDataImport_Click()
    FRE_Company.Visible = False
    FRE_DocNumbers.Visible = False
    FRE_Users.Visible = False
    FRE_AutoBackups.Visible = False
    FRE_References.Visible = False
    Fre_Reset.Visible = False
    FRE_Import.Visible = True
End Sub

Private Sub btnDocuments_Click()
    FRE_Company.Visible = False
    FRE_DocNumbers.Visible = True
    FRE_Users.Visible = False
    FRE_AutoBackups.Visible = False
    FRE_References.Visible = False
    Fre_Reset.Visible = False
    Fre_Reset.Visible = False
    FRE_Import.Visible = False
    txtPrefix_SalesOrder.SetFocus
End Sub

Private Sub btnExecute_Click()
    If DialogBox.filename = "" Then
        MsgBox "Please load a csv file to execute.", vbCritical
        Exit Sub
    End If
    
    Dim FilePathOnly As String
    FilePathOnly = Left(DialogBox.filename, Len(DialogBox.filename) - Len(DialogBox.FileTitle))

    'Import data
    GetCSVData DialogBox.FileTitle, FilePathOnly

    
    ProgressBar.Min = 0
    ProgressBar.value = 0
    ProgressBar.Max = UniversalCtr
    
    If FRE_Import_Details.Caption = "Import Products" Then
        With CSVRecordset
            If Not .EOF Then
                'Clear Data Import log first
                ClearDataImportLog
                
                Do Until .EOF
                   'CHECK FOR CATEGORY EXISTENCE
                   Dim Category As String
                   Dim CategoryId As Long
                   
                   If IsNull(!Category) = True Then
                        Category = "Default"
                   ElseIf !Category = "" Then
                    Category = "Default"
                   Else
                        Category = !Category
                   End If
                   
                   CategoryId = CategoryImport(Category)
                   
                   'CHECK FOR SUPPLIER EXISTENCE
                   Dim Supplier As String
                   Dim SupplierId As Long
'                   If IsNull(!Supplier) = True Then Supplier = "" Else Supplier = !Supplier
'                   If !Supplier = "" Then Supplier = "Default" Else Supplier = !Supplier
                   If IsNull(!Supplier) = True Then
                        Supplier = "Default"
                   ElseIf !Supplier = "" Then
                    Supplier = "Default"
                   Else
                        Supplier = !Supplier
                   End If

                   SupplierId = SupplierImport(Supplier)
                   
                   'CHECK FOR UNIT EXISTENCE
                   Dim Uom As String
                   Dim UomId As Long
'                   If IsNull(!unit) = True Then Uom = "pcs" Else Uom = !unit
'                   If !unit = "" Then Uom = "pcs" Else Uom = !unit

                   If IsNull(!unit) = True Then
                        Uom = "Default"
                   ElseIf !unit = "" Then
                        Uom = "Default"
                   Else
                        Uom = !unit
                   End If
                   UomId = UomImport(Uom)
                   
                   'TAX
                   Dim TaxId As Long
                   Dim tax As String
                   If !tax = UCase("VAT") Then TaxId = 2 Else TaxId = 1
                   
                   'Barcode
                   Dim Barcode As String
                   If IsNull(!Barcode) = True Then Barcode = "" Else Barcode = !Barcode
                   
                   'On Error GoTo errcode
                   'PRODUCT INPUT AND FILTERS
                   If !itemcode = "" Then
                        MsgBox "Cannot import data with empty itemcode. Please check csv file. Import aborted.", vbCritical
                        Exit Sub
                   ElseIf !Name = "" Then
                        MsgBox "Cannot import data with empty name. Please check csv file. Import aborted.", vbCritical
                        Exit Sub
                   Else
                        
                        'START DATA IMPORT
                        Dim con As New ADODB.Connection
                        
                        con.ConnectionString = ConnString
                        con.Open
                        Set cmd = New ADODB.Command
                        cmd.ActiveConnection = con
                        cmd.CommandType = adCmdStoredProc
                        cmd.CommandText = "SYS_Import_Product"
                        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInputOutput, , 1)
                        cmd.Parameters.Append cmd.CreateParameter("@ItemCode", adVarChar, adParamInput, 500, !itemcode)
                        cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, !Name)
                        cmd.Parameters.Append cmd.CreateParameter("@Barcode", adVarChar, adParamInput, 50, Barcode)
                        cmd.Parameters.Append cmd.CreateParameter("@Alias1", adVarChar, adParamInput, 50, "")
                        cmd.Parameters.Append cmd.CreateParameter("@Alias2", adVarChar, adParamInput, 50, "")
                        cmd.Parameters.Append cmd.CreateParameter("@CategoryId", adInteger, adParamInput, , CategoryId)
                        cmd.Parameters.Append cmd.CreateParameter("@TypeId", adInteger, adParamInput, , 1)
                        cmd.Parameters.Append cmd.CreateParameter("@TaxId", adInteger, adParamInput, , TaxId)
                        cmd.Parameters.Append cmd.CreateParameter("@UnitPrice", adDecimal, adParamInput, , !sellingprice)
                                              cmd.Parameters("@UnitPrice").Precision = 18
                                              cmd.Parameters("@UnitPrice").NumericScale = 2
                        cmd.Parameters.Append cmd.CreateParameter("@Price1", adDecimal, adParamInput, , Null)
                                              cmd.Parameters("@Price1").Precision = 18
                                              cmd.Parameters("@Price1").NumericScale = 2
                        cmd.Parameters.Append cmd.CreateParameter("@Price2", adDecimal, adParamInput, , Null)
                                              cmd.Parameters("@Price2").Precision = 18
                                              cmd.Parameters("@Price2").NumericScale = 2
                        cmd.Parameters.Append cmd.CreateParameter("@Price3", adDecimal, adParamInput, , Null)
                                              cmd.Parameters("@Price3").Precision = 18
                                              cmd.Parameters("@Price3").NumericScale = 2
                        cmd.Parameters.Append cmd.CreateParameter("@UnitPriceMarkUp", adDecimal, adParamInput, , Null)
                                              cmd.Parameters("@UnitPriceMarkUp").Precision = 18
                                              cmd.Parameters("@UnitPriceMarkUp").NumericScale = 2
                        cmd.Parameters.Append cmd.CreateParameter("@Price1MarkUp", adDecimal, adParamInput, , Null)
                                              cmd.Parameters("@Price1MarkUp").Precision = 18
                                              cmd.Parameters("@Price1MarkUp").NumericScale = 2
                        cmd.Parameters.Append cmd.CreateParameter("@Price2MarkUp", adDecimal, adParamInput, , Null)
                                              cmd.Parameters("@Price2MarkUp").Precision = 18
                                              cmd.Parameters("@Price2MarkUp").NumericScale = 2
                        cmd.Parameters.Append cmd.CreateParameter("@Price3MarkUp", adDecimal, adParamInput, , Null)
                                              cmd.Parameters("@Price3MarkUp").Precision = 18
                                              cmd.Parameters("@Price3MarkUp").NumericScale = 2
                        cmd.Parameters.Append cmd.CreateParameter("@UnitCost", adDecimal, adParamInput, , !cost)
                                              cmd.Parameters("@UnitCost").Precision = 18
                                              cmd.Parameters("@UnitCost").NumericScale = 2
                        cmd.Parameters.Append cmd.CreateParameter("@ReorderPoint", adDecimal, adParamInput, , 0)
                                              cmd.Parameters("@ReorderPoint").Precision = 18
                                              cmd.Parameters("@ReorderPoint").NumericScale = 2
                        cmd.Parameters.Append cmd.CreateParameter("@ReorderQuantity", adDecimal, adParamInput, , 0)
                                              cmd.Parameters("@ReorderQuantity").Precision = 18
                                              cmd.Parameters("@ReorderQuantity").NumericScale = 2
                        cmd.Parameters.Append cmd.CreateParameter("@Uom", adVarChar, adParamInput, 50, !unit)
                        cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInput, , SupplierId)
                        If chkIgnore.value = True Then
                            cmd.Parameters.Append cmd.CreateParameter("@Action", adInteger, adParamInput, , 1)
                        ElseIf chkOverwrite.value = True Then
                            cmd.Parameters.Append cmd.CreateParameter("@Action", adInteger, adParamInput, , 2)
                        ElseIf chkCancel.value = True Then
                            cmd.Parameters.Append cmd.CreateParameter("@Action", adInteger, adParamInput, , 3)
                        End If
                        cmd.Execute
                        
                        lblProgressDetails.Caption = "Importing product " & !Name & "..."
                        If ProgressBar.value = ProgressBar.Max Then
                            ProgressBar.Max = ProgressBar.value + 100
                        Else
                            ProgressBar.value = ProgressBar.value + 1
                        End If
                        con.Close
                   End If
                    .MoveNext
                Loop
            End If
        End With
        lblProgressDetails.Caption = "Import complete."
        ProgressBar.value = ProgressBar.Max
    End If
    DialogBox.filename = ""
    Set CSVRecordset = Nothing
    lblPath.Caption = ""
Exit Sub
errcode:
    MsgBox Err.Description, vbCritical
    lblProgressDetails.Caption = "Import cancelled."
End Sub


Private Sub btnExpenses_Click()
    FIN_ExpenseListFrm.Show (1)
End Sub

Private Sub btnFunds_Click()
    BASE_WarehousePersonnelFrm.Show (1)
End Sub

Private Sub btnImportProduct_Click()
    FRE_Import_Details.Visible = True
    FRE_Import_Details.Caption = "Import Products"

    
End Sub

Private Sub btnLoadCSV_Click()
    On Error GoTo errcode
    DialogBox.Flags = cdlOFNHideReadOnly
    DialogBox.Filter = "CSV (*.CSV)|*.CSV|TEXT (*.txt)|*.txt|"
    DialogBox.ShowOpen
    DialogBox.CancelError = True
    lblPath.Caption = DialogBox.filename
    If DialogBox.FileTitle = "" Then Exit Sub
errcode:
    Exit Sub
End Sub

Private Sub btnLocations_Click()
    INV_LocationModFrm.Show (1)
End Sub

Private Sub btnMasterReset_Click()
    Dim x As Variant
    x = MsgBox("WARNING! This will reset all system data to default settings. This process cannot be reverted. Continue?", vbYesNo + vbCritical)
    If x = vbYes Then
        AllowAccess = False
        POS_UserPinFrm.Show (1)
        If AllowAccess = True Then
            Dim con As New ADODB.Connection
            con.ConnectionString = ConnString
            con.Open
            Dim cmd As New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "SYS_Reset_Data"
            cmd.Parameters.Append cmd.CreateParameter("@DataType", adVarChar, adParamInput, 50, "MASTER")
            cmd.Execute
            con.Close
            MsgBox "Records reset successfully! System will now log off.", vbInformation
            Unload Me
            Unload BASE_ContainerFrm
        End If
    End If
End Sub

Private Sub btnPaymentMethod_Click()
    BASE_PaymentMethodsFrm.Show (1)
End Sub

Private Sub btnPOSOrderSlip_Click()
    Dim x As Variant
    x = MsgBox("WARNING! This will reset all POS Order slip and hold lists. Continue?", vbYesNo + vbCritical)
    If x = vbYes Then
        Dim con As New ADODB.Connection
        con.ConnectionString = ConnString
        con.Open
        Dim cmd As New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SYS_Reset_Data"
        cmd.Parameters.Append cmd.CreateParameter("@DataType", adVarChar, adParamInput, 50, "POS ORDER")
        cmd.Execute
        con.Close
        MsgBox "Records reset successfully!", vbInformation
    End If
End Sub

Private Sub btnPOSReset_Click()
    Dim x As Variant
    x = MsgBox("WARNING! This will reset all POS transactions including sales records. Continue?", vbYesNo + vbCritical)
    If x = vbYes Then
        Dim con As New ADODB.Connection
        con.ConnectionString = ConnString
        con.Open
        Dim cmd As New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SYS_Reset_Data"
        cmd.Parameters.Append cmd.CreateParameter("@DataType", adVarChar, adParamInput, 50, "POS")
        cmd.Execute
        con.Close
        MsgBox "Records reset successfully!", vbInformation
    End If
    
End Sub

Private Sub btnPricingScheme_Click()
   BASE_PricingSchemeFrm.Show (1)
End Sub

Private Sub btnReferences_Click()
    FRE_Company.Visible = False
    FRE_DocNumbers.Visible = False
    FRE_Users.Visible = False
    FRE_AutoBackups.Visible = False
    FRE_References.Visible = True
    Fre_Reset.Visible = False
    FRE_Import.Visible = False
    btnPaymentMethod.SetFocus
End Sub

Private Sub btnRemove_Click()
    If lvUsers.SelectedItem.SubItems(1) = 1 Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(46)
        GLOBAL_MessageFrm.Show (1)
        Exit Sub
    End If
    
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_User_Update"
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , lvUsers.SelectedItem.SubItems(1))
    cmd.Parameters.Append cmd.CreateParameter("@RoleId", adInteger, adParamInput, , cmbRoles.ItemData(cmbRoles.ListIndex))
    cmd.Parameters.Append cmd.CreateParameter("@Usernumber", adInteger, adParamInput, , lvUsers.SelectedItem.SubItems(2))
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 50, Null)
    cmd.Parameters.Append cmd.CreateParameter("@Username", adVarChar, adParamInput, 50, Null)
    cmd.Parameters.Append cmd.CreateParameter("@Password", adVarChar, adParamInput, 50, Null)
    cmd.Parameters.Append cmd.CreateParameter("@Pin", adVarChar, adParamInput, 4, Null)
    cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , "False")
    cmd.Execute
    con.Close
    lvUsers.ListItems.Remove (lvUsers.SelectedItem.Index)
End Sub

Private Sub btnReset_Click()
    FRE_Company.Visible = False
    FRE_DocNumbers.Visible = False
    FRE_Users.Visible = False
    FRE_AutoBackups.Visible = False
    FRE_References.Visible = False
    Fre_Reset.Visible = True
    FRE_Import.Visible = False
End Sub

Private Sub btnResetAll_Click()
    Dim x As Variant
    x = MsgBox("WARNING! This will reset all Purchasing, Sales, Inventory and POS records. Continue?", vbYesNo + vbCritical)
    If x = vbYes Then
        Dim con As New ADODB.Connection
        con.ConnectionString = ConnString
        con.Open
        Dim cmd As New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SYS_Reset_Data"
        cmd.Parameters.Append cmd.CreateParameter("@DataType", adVarChar, adParamInput, 50, "POS")
        cmd.Execute
        
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SYS_Reset_Data"
        cmd.Parameters.Append cmd.CreateParameter("@DataType", adVarChar, adParamInput, 50, "INVENTORY")
        cmd.Execute
        
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SYS_Reset_Data"
        cmd.Parameters.Append cmd.CreateParameter("@DataType", adVarChar, adParamInput, 50, "SALESORDER")
        cmd.Execute
        
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SYS_Reset_Data"
        cmd.Parameters.Append cmd.CreateParameter("@DataType", adVarChar, adParamInput, 50, "PURCHASEORDER")
        cmd.Execute
        
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SYS_Reset_Data"
        cmd.Parameters.Append cmd.CreateParameter("@DataType", adVarChar, adParamInput, 50, "SYSTEM")
        cmd.Execute
        
        con.Close
        MsgBox "Records reset successfully!", vbInformation
    End If
End Sub

Private Sub btnResetInventory_Click()
    Dim x As Variant
    x = MsgBox("WARNING! This will reset all inventory records. Product list will not be affected. Continue?", vbYesNo + vbCritical)
    If x = vbYes Then
        Dim con As New ADODB.Connection
        con.ConnectionString = ConnString
        con.Open
        Dim cmd As New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SYS_Reset_Data"
        cmd.Parameters.Append cmd.CreateParameter("@DataType", adVarChar, adParamInput, 50, "INVENTORY")
        cmd.Execute
        con.Close
        MsgBox "Records reset successfully!", vbInformation
    End If
    
End Sub

Private Sub btnResetPurchases_Click()
    Dim x As Variant
    x = MsgBox("WARNING! This will reset all sales order records. Continue?", vbYesNo + vbCritical)
    If x = vbYes Then
        Dim con As New ADODB.Connection
        con.ConnectionString = ConnString
        con.Open
        Dim cmd As New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SYS_Reset_Data"
        cmd.Parameters.Append cmd.CreateParameter("@DataType", adVarChar, adParamInput, 50, "PURCHASEORDER")
        cmd.Execute
        con.Close
        MsgBox "Records reset successfully!", vbInformation
    End If
    
End Sub

Private Sub btnResetSalesOrders_Click()
    Dim x As Variant
    x = MsgBox("WARNING! This will reset all sales order records. Continue?", vbYesNo + vbCritical)
    If x = vbYes Then
        Dim con As New ADODB.Connection
        con.ConnectionString = ConnString
        con.Open
        Dim cmd As New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SYS_Reset_Data"
        cmd.Parameters.Append cmd.CreateParameter("@DataType", adVarChar, adParamInput, 50, "SALESORDER")
        cmd.Execute
        con.Close
        MsgBox "Records reset successfully!", vbInformation
    End If
    
End Sub

Private Sub btnSave_Click()
    'COMPANY
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_Company_Update"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, txtCompanyName.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Phone", adVarChar, adParamInput, 50, txtPhone.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Fax", adVarChar, adParamInput, 50, txtFax.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Address1", adVarChar, adParamInput, 250, txtAddress1.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Address2", adVarChar, adParamInput, 250, txtAddress2.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Email", adVarChar, adParamInput, 50, txtEmail.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Website", adVarChar, adParamInput, 500, txtWebsite.Text)
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
    cmd.Execute
    
    'DOCUMENT FORMAT
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "GLOBAL_DocNoFormat_Update"
    cmd.Parameters.Append cmd.CreateParameter("@DocNoFormatId", adInteger, adParamInput, , 1)
    cmd.Parameters.Append cmd.CreateParameter("@NextNumber", adInteger, adParamInput, , txtNextNumber_PurchaseOrder.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Prefix", adVarChar, adParamInput, 50, txtPrefix_PurchaseOrder.Text)
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
    cmd.Execute
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "GLOBAL_DocNoFormat_Update"
    cmd.Parameters.Append cmd.CreateParameter("@DocNoFormatId", adInteger, adParamInput, , 2)
    cmd.Parameters.Append cmd.CreateParameter("@NextNumber", adInteger, adParamInput, , txtNextNumber_SalesOrder.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Prefix", adVarChar, adParamInput, 50, txtPrefix_SalesOrder.Text)
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
    cmd.Execute
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "GLOBAL_DocNoFormat_Update"
    cmd.Parameters.Append cmd.CreateParameter("@DocNoFormatId", adInteger, adParamInput, , 3)
    cmd.Parameters.Append cmd.CreateParameter("@NextNumber", adInteger, adParamInput, , txtNextNumber_POS.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Prefix", adVarChar, adParamInput, 50, txtPrefix_POS.Text)
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
    cmd.Execute
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "GLOBAL_DocNoFormat_Update"
    cmd.Parameters.Append cmd.CreateParameter("@DocNoFormatId", adInteger, adParamInput, , 4)
    cmd.Parameters.Append cmd.CreateParameter("@NextNumber", adInteger, adParamInput, , txtNextNumber_TransferStock.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Prefix", adVarChar, adParamInput, 50, txtPrefix_TransferStock.Text)
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
    cmd.Execute
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "GLOBAL_DocNoFormat_Update"
    cmd.Parameters.Append cmd.CreateParameter("@DocNoFormatId", adInteger, adParamInput, , 5)
    cmd.Parameters.Append cmd.CreateParameter("@NextNumber", adInteger, adParamInput, , txtNextNumber_CA1.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Prefix", adVarChar, adParamInput, 50, txtPrefix_CA1.Text)
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
    cmd.Execute
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "GLOBAL_DocNoFormat_Update"
    cmd.Parameters.Append cmd.CreateParameter("@DocNoFormatId", adInteger, adParamInput, , 6) 'Purchase Return
    cmd.Parameters.Append cmd.CreateParameter("@NextNumber", adInteger, adParamInput, , txtNextNumber_PurchaseReturn.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Prefix", adVarChar, adParamInput, 50, txtPrefix_PurchaseReturn.Text)
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
    cmd.Execute
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "GLOBAL_DocNoFormat_Update"
    cmd.Parameters.Append cmd.CreateParameter("@DocNoFormatId", adInteger, adParamInput, , 7) 'Sales Return
    cmd.Parameters.Append cmd.CreateParameter("@NextNumber", adInteger, adParamInput, , txtNextNumber_SalesReturn.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Prefix", adVarChar, adParamInput, 50, txtPrefix_SalesReturn.Text)
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
    cmd.Execute
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "GLOBAL_DocNoFormat_Update"
    cmd.Parameters.Append cmd.CreateParameter("@DocNoFormatId", adInteger, adParamInput, , 8) 'New Stock
    cmd.Parameters.Append cmd.CreateParameter("@NextNumber", adInteger, adParamInput, , txtNextNumber_NewStock.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Prefix", adVarChar, adParamInput, 50, txtPrefix_NewStock.Text)
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
    cmd.Execute
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "GLOBAL_DocNoFormat_Update"
    cmd.Parameters.Append cmd.CreateParameter("@DocNoFormatId", adInteger, adParamInput, , 9) 'Audit Stock
    cmd.Parameters.Append cmd.CreateParameter("@NextNumber", adInteger, adParamInput, , txtNextNumber_AuditStock.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Prefix", adVarChar, adParamInput, 50, txtPrefix_AuditStock.Text)
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
    cmd.Execute
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "GLOBAL_DocNoFormat_Update"
    cmd.Parameters.Append cmd.CreateParameter("@DocNoFormatId", adInteger, adParamInput, , 10) 'Sales Adjustment
    cmd.Parameters.Append cmd.CreateParameter("@NextNumber", adInteger, adParamInput, , txtNextNumber_SalesAdjustment.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Prefix", adVarChar, adParamInput, 50, txtPrefix_SalesAdjustment.Text)
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
    cmd.Execute
    
    'Users
    Dim item As MSComctlLib.ListItem
    For Each item In lvUsers.ListItems
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "BASE_User_Update"
        cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , item.SubItems(1))
        cmd.Parameters.Append cmd.CreateParameter("@UserRoleId", adInteger, adParamInput, , item.SubItems(5))
        cmd.Parameters.Append cmd.CreateParameter("@UserNumber", adInteger, adParamInput, , item.SubItems(2))
        cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 50, Null)
        cmd.Parameters.Append cmd.CreateParameter("@Username", adVarChar, adParamInput, 50, Null)
        cmd.Parameters.Append cmd.CreateParameter("@Password", adVarChar, adParamInput, 50, Null)
        cmd.Parameters.Append cmd.CreateParameter("@Pin", adVarChar, adParamInput, 4, Null)
        If item.Checked = True Then
            cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , "True")
        Else
            cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , "False")
        End If
        cmd.Parameters.Append cmd.CreateParameter("@CurrentUserId", adInteger, adParamInput, , UserId)
        cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
        cmd.Execute
    Next
    
    con.Close
    Unload Me
End Sub



Private Sub btnTax_Click()
    BASE_TaxFrm.Show (1)
End Sub

Private Sub btnTerms_Click()
    BASE_TermsFrm.Show (1)
End Sub

Private Sub btnUnits_Click()
    BASE_UnitsFrm.Show (1)
End Sub

Private Sub btnUsers_Click()
    FRE_Company.Visible = False
    FRE_DocNumbers.Visible = False
    FRE_Users.Visible = True
    FRE_AutoBackups.Visible = False
    FRE_References.Visible = False
    Fre_Reset.Visible = False
    FRE_Import.Visible = False
    lvUsers.SetFocus
End Sub

Private Sub btnWorkstations_Click()
    SYS_WorkstationFrm.Show (1)
End Sub

Private Sub chkShow_Click()
    Dim item As MSComctlLib.ListItem
    Set rec = New ADODB.Recordset
    Set rec = Global_Data("User")
    lvUsers.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If chkShow.value = 1 Then
                Set item = lvUsers.ListItems.add(, , "")
                    item.SubItems(1) = rec!UserId
                    item.SubItems(2) = rec!UserNumber
                    item.SubItems(3) = rec!Name
                    item.SubItems(4) = rec!role
                    item.SubItems(5) = rec!UserRoleId
                    
                If rec!isActive = "True" Then item.Checked = True
                lvUsers.ColumnHeaders(1).width = lvUsers.width * 0.04
                lvUsers.ColumnHeaders(3).width = lvUsers.width * 0.15
                lvUsers.ColumnHeaders(4).width = lvUsers.width * 0.52
                lvUsers.ColumnHeaders(5).width = lvUsers.width * 0.25
                btnRemove.Enabled = False
            Else
                If rec!isActive = "True" Then
                    Set item = lvUsers.ListItems.add(, , "")
                        item.SubItems(1) = rec!UserId
                        item.SubItems(2) = rec!UserNumber
                        item.SubItems(3) = rec!Name
                        item.SubItems(4) = rec!role
                        item.SubItems(5) = rec!UserRoleId
                        
                    If rec!isActive = "True" Then item.Checked = True
                    lvUsers.ColumnHeaders(1).width = lvUsers.width * 0
                    lvUsers.ColumnHeaders(3).width = lvUsers.width * 0.15
                    lvUsers.ColumnHeaders(4).width = lvUsers.width * 0.56
                    lvUsers.ColumnHeaders(5).width = lvUsers.width * 0.25
                End If
                btnRemove.Enabled = True
            End If
            rec.MoveNext
        Loop
    End If
End Sub

Private Sub Command2_Click()
    
End Sub

Private Sub Command3_Click()
    FRE_Company.Visible = False
    FRE_DocNumbers.Visible = False
    FRE_Users.Visible = False
    FRE_AutoBackups.Visible = True
    FRE_References.Visible = False
    Fre_Reset.Visible = False
    FRE_Import.Visible = False
End Sub

Private Sub Form_Load()
    Populate "Company"
    Populate "User"
    Populate "Documents"
    Populate "UserRoles"
    
    chkShow_Click
    btnCompany_Click
    
    ViewAccessRights (26)
    ViewAccessRights (27)
End Sub

Private Sub lblInventory_MoreLocations_Click()
    
End Sub

Private Sub Label49_Click()
    Screen.MousePointer = vbHourglass
    BASE_PrintPreviewFrm.isInvoice = False
    BASE_PrintPreviewFrm.Show '(1)
    Dim crxApp As New CRAXDRT.Application
    Dim crxRpt As New CRAXDRT.Report
    Set crxRpt = crxApp.OpenReport(App.path & "\Reports\SYS_DataImportLog.rpt")
    crxRpt.EnableParameterPrompting = False
    BASE_PrintPreviewFrm.tb_Standard.Visible = False
    'crxRpt.RecordSelectionFormula = "{SO_SalesOrder.SalesOrderId}= " & SalesOrderId & ""
    
    crxRpt.DiscardSavedData

    Call ResetRptDB(crxRpt)
    'crxRpt.ParameterFields.GetItemByName("@SalesOrderId").AddCurrentValue Int(SalesOrderId)
    
    BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
    BASE_PrintPreviewFrm.CRViewer.ViewReport
    BASE_PrintPreviewFrm.CRViewer.Zoom 1
    Screen.MousePointer = vbDefault
End Sub

Private Sub lblcsv_products_Click()
    DialogBox.InitDir = App.path & "\Templates\Products.csv"
    DialogBox.filename = "products.csv"
    DialogBox.ShowSave
    Dim FilePathOnly As String
    FilePathOnly = Left(DialogBox.filename, Len(DialogBox.filename) - Len(DialogBox.FileTitle))
    FileCopy App.path & "\Templates\Products.csv", FilePathOnly & DialogBox.FileTitle
End Sub

Private Sub lblUserRoles_Click()
    BASE_UserRolesFrm.Show (1)
End Sub

Private Sub lvUsers_DblClick()
    If lvUsers.ListItems.Count > 0 Then
        On Error Resume Next
        BASE_UserRightsFrm.CheckUserId = lvUsers.SelectedItem.SubItems(1)
        BASE_UserRightsFrm.cUserRoleId = lvUsers.SelectedItem.SubItems(5)
        BASE_UserRightsFrm.cmbRoles.Text = lvUsers.SelectedItem.SubItems(4)
        BASE_UserRightsFrm.txtUserNumber.Text = lvUsers.SelectedItem.SubItems(2)
        BASE_UserRightsFrm.Show (1)
    End If
End Sub

Private Sub lvUsers_ItemCheck(ByVal item As MSComctlLib.ListItem)
    If item.SubItems(1) = "1" Then item.Checked = True
End Sub

Private Sub txtName_GotFocus()
    selectText txtName
End Sub

Private Sub txtNextNumber_TransferStock_Change()
    If IsNumeric(txtNextNumber_TransferStock.Text) = False Then
        txtNextNumber_TransferStock.Text = "1"
    Else
        txtNextNumber_TransferStock.Text = Int(txtNextNumber_TransferStock.Text)
    End If
    lblPreview_TransferStock.Caption = txtPrefix_TransferStock.Text & Format(txtNextNumber_TransferStock.Text, "000000")
End Sub

Private Sub txtNextNumber_AuditStock_Change()
    If IsNumeric(txtNextNumber_AuditStock.Text) = False Then
        txtNextNumber_AuditStock.Text = "1"
    Else
        txtNextNumber_AuditStock.Text = Int(txtNextNumber_AuditStock.Text)
    End If
    lblPreview_AuditStock.Caption = txtPrefix_AuditStock.Text & Format(txtNextNumber_AuditStock.Text, "000000")
End Sub

Private Sub txtNextNumber_CA1_Change()
    If IsNumeric(txtNextNumber_CA1.Text) = False Then
        txtNextNumber_CA1.Text = "1"
    Else
        txtNextNumber_CA1.Text = Int(txtNextNumber_CA1.Text)
    End If
    lblPreview_CA1.Caption = txtPrefix_CA1.Text & Format(txtNextNumber_CA1.Text, "000000")
End Sub

Private Sub txtNextNumber_NewStock_Change()
    If IsNumeric(txtNextNumber_NewStock.Text) = False Then
        txtNextNumber_NewStock.Text = "1"
    Else
        txtNextNumber_NewStock.Text = Int(txtNextNumber_NewStock.Text)
    End If
    lblPreview_NewStock.Caption = txtPrefix_NewStock.Text & Format(txtNextNumber_NewStock.Text, "000000")
End Sub

Private Sub txtNextNumber_POS_Change()
    If IsNumeric(txtNextNumber_POS.Text) = False Then
        txtNextNumber_POS.Text = "1"
    Else
        txtNextNumber_POS.Text = Int(txtNextNumber_POS.Text)
    End If
    lblPreview_POS.Caption = txtPrefix_POS.Text & Format(txtNextNumber_POS.Text, "000000")
End Sub

Private Sub txtNextNumber_POS_GotFocus()
    selectText txtNextNumber_POS
End Sub

Private Sub txtNextNumber_PurchaseOrder_Change()
    If IsNumeric(txtNextNumber_PurchaseOrder.Text) = False Then
        txtNextNumber_PurchaseOrder.Text = "1"
    Else
        txtNextNumber_PurchaseOrder.Text = Int(txtNextNumber_PurchaseOrder.Text)
    End If
    lblPreview_PurchaseOrder.Caption = txtPrefix_PurchaseOrder.Text & Format(txtNextNumber_PurchaseOrder.Text, "000000")
End Sub

Private Sub txtNextNumber_PurchaseOrder_GotFocus()
    selectText txtNextNumber_PurchaseOrder
End Sub

Private Sub txtNextNumber_PurchaseReturn_Change()
    If IsNumeric(txtNextNumber_PurchaseReturn.Text) = False Then
        txtNextNumber_PurchaseReturn.Text = "1"
    Else
        txtNextNumber_PurchaseReturn.Text = Int(txtNextNumber_PurchaseReturn.Text)
    End If
    lblPreview_PurchaseReturn.Caption = txtPrefix_PurchaseReturn.Text & Format(txtNextNumber_PurchaseReturn.Text, "000000")
End Sub

Private Sub txtNextNumber_SalesAdjustment_Change()
    If IsNumeric(txtNextNumber_SalesAdjustment.Text) = False Then
        txtNextNumber_SalesAdjustment.Text = "1"
    Else
        txtNextNumber_SalesAdjustment.Text = Int(txtNextNumber_SalesAdjustment.Text)
    End If
    lblPreview_SalesAdjustment.Caption = txtPrefix_SalesAdjustment.Text & Format(txtNextNumber_SalesAdjustment.Text, "000000")
End Sub

Private Sub txtNextNumber_SalesOrder_Change()
    If IsNumeric(txtNextNumber_SalesOrder.Text) = False Then
        txtNextNumber_SalesOrder.Text = "1"
    Else
        txtNextNumber_SalesOrder.Text = Int(txtNextNumber_SalesOrder.Text)
    End If
    lblPreview_SalesOrder.Caption = txtPrefix_SalesOrder.Text & Format(txtNextNumber_SalesOrder.Text, "000000")
End Sub

Private Sub txtNextNumber_SalesOrder_GotFocus()
    selectText txtNextNumber_SalesOrder
End Sub



Private Sub txtNextNumber_SalesReturn_Change()
    If IsNumeric(txtNextNumber_SalesReturn.Text) = False Then
        txtNextNumber_SalesReturn.Text = "1"
    Else
        txtNextNumber_SalesReturn.Text = Int(txtNextNumber_SalesReturn.Text)
    End If
    lblPreview_SalesReturn.Caption = txtPrefix_SalesReturn.Text & Format(txtNextNumber_SalesReturn.Text, "000000")
End Sub



Private Sub txtPrefix_AuditStock_Change()
    lblPreview_AuditStock.Caption = txtPrefix_AuditStock.Text & Format(txtNextNumber_AuditStock.Text, "000000")
End Sub

Private Sub txtPrefix_AuditStock_GotFocus()
    selectText txtPrefix_AuditStock
End Sub

Private Sub txtPrefix_CA1_Change()
    lblPreview_CA1.Caption = txtPrefix_CA1.Text & Format(txtNextNumber_CA1.Text, "000000")
End Sub

Private Sub txtPrefix_CA1_GotFocus()
    selectText txtPrefix_CA1
End Sub

Private Sub txtPrefix_NewStock_Change()
    lblPreview_NewStock.Caption = txtPrefix_NewStock.Text & Format(txtNextNumber_NewStock.Text, "000000")
End Sub

Private Sub txtPrefix_NewStock_GotFocus()
    selectText txtPrefix_NewStock
End Sub

Private Sub txtPrefix_POS_Change()
    lblPreview_POS.Caption = txtPrefix_POS.Text & Format(txtNextNumber_POS.Text, "000000")
End Sub

Private Sub txtPrefix_POS_GotFocus()
    selectText txtPrefix_POS
End Sub

Private Sub txtPrefix_PurchaseOrder_Change()
    lblPreview_PurchaseOrder.Caption = txtPrefix_PurchaseOrder.Text & Format(txtNextNumber_PurchaseOrder.Text, "000000")
End Sub

Private Sub txtPrefix_PurchaseOrder_GotFocus()
    selectText txtPrefix_PurchaseOrder
End Sub

Private Sub txtPrefix_PurchaseReturn_Change()
    lblPreview_PurchaseReturn.Caption = txtPrefix_PurchaseReturn.Text & Format(txtNextNumber_PurchaseReturn.Text, "000000")
End Sub

Private Sub txtPrefix_PurchaseReturn_GotFocus()
    selectText txtPrefix_PurchaseReturn
End Sub

Private Sub txtPrefix_SalesAdjustment_Change()
    lblPreview_SalesAdjustment.Caption = txtPrefix_SalesAdjustment.Text & Format(txtNextNumber_SalesAdjustment.Text, "000000")
End Sub

Private Sub txtPrefix_SalesAdjustment_GotFocus()
    selectText txtPrefix_SalesAdjustment
End Sub

Private Sub txtPrefix_SalesOrder_Change()
    lblPreview_SalesOrder.Caption = txtPrefix_SalesOrder.Text & Format(txtNextNumber_SalesOrder.Text, "000000")
End Sub

Private Sub txtPrefix_SalesOrder_GotFocus()
    selectText txtPrefix_SalesOrder
End Sub

Private Sub txtPrefix_SalesReturn_Change()
    lblPreview_SalesReturn.Caption = txtPrefix_SalesReturn.Text & Format(txtNextNumber_SalesReturn.Text, "000000")
End Sub

Private Sub txtPrefix_SalesReturn_GotFocus()
    selectText txtPrefix_SalesReturn
End Sub

Private Sub txtPrefix_TransferStock_Change()
    lblPreview_TransferStock.Caption = txtPrefix_TransferStock.Text & Format(txtNextNumber_TransferStock.Text, "000000")
End Sub

Private Sub txtPrefix_TransferStock_GotFocus()
    selectText txtPrefix_TransferStock
End Sub

Private Sub txtUserNumber_GotFocus()
    selectText txtUserNumber
End Sub
