VERSION 5.00
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Begin VB.Form frmExport 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   1005
      Left            =   270
      TabIndex        =   0
      Top             =   90
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   1773
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Begin vbalProgBarLib6.vbalProgressBar progBar 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "5"
         Top             =   540
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   556
         Picture         =   "frmExport.frx":0000
         BackColor       =   14737632
         ForeColor       =   0
         Appearance      =   2
         BorderStyle     =   2
         BarColor        =   16744576
         BarPicture      =   "frmExport.frx":001C
         Max             =   66
         ShowText        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exporting to CSV file...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   1485
         TabIndex        =   2
         Top             =   180
         Width           =   2220
      End
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
