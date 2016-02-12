VERSION 5.00
Begin VB.Form ScratchPad 
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   4365
   ClientTop       =   3285
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   6060
   Begin VB.Frame slurry_tube_frame 
      Caption         =   "Slurry Tube Controls"
      Height          =   2295
      Left            =   720
      TabIndex        =   0
      Top             =   1920
      Width           =   4575
      Begin VB.Line slurrytubetubing 
         Index           =   13
         X1              =   3840
         X2              =   3840
         Y1              =   360
         Y2              =   600
      End
      Begin VB.Line slurrytubetubing 
         Index           =   11
         X1              =   2760
         X2              =   2760
         Y1              =   1080
         Y2              =   1320
      End
      Begin VB.Label ValveClik 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   25
         Left            =   2640
         TabIndex        =   14
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label ValveLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "V44"
         Height          =   255
         Index           =   25
         Left            =   2880
         TabIndex        =   13
         Top             =   1320
         Width           =   375
      End
      Begin VB.Line vcloseline 
         Index           =   25
         X1              =   2880
         X2              =   2640
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line vopenline 
         Index           =   25
         Visible         =   0   'False
         X1              =   2760
         X2              =   2760
         Y1              =   1560
         Y2              =   1320
      End
      Begin VB.Label ValveClik 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   24
         Left            =   3720
         TabIndex        =   12
         Top             =   600
         Width           =   255
      End
      Begin VB.Label ValveLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "V43"
         Height          =   255
         Index           =   24
         Left            =   3960
         TabIndex        =   11
         Top             =   600
         Width           =   375
      End
      Begin VB.Line vcloseline 
         Index           =   24
         X1              =   3960
         X2              =   3720
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line vopenline 
         Index           =   24
         Visible         =   0   'False
         X1              =   3840
         X2              =   3840
         Y1              =   840
         Y2              =   600
      End
      Begin VB.Line slurrytubetubing 
         Index           =   10
         X1              =   1080
         X2              =   1080
         Y1              =   1320
         Y2              =   1440
      End
      Begin VB.Line slurrytubetubing 
         Index           =   9
         X1              =   1080
         X2              =   1680
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line slurrytubetubing 
         Index           =   8
         X1              =   1800
         X2              =   1800
         Y1              =   1080
         Y2              =   1320
      End
      Begin VB.Label slurrylabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "  Wash"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin VB.Label slurrylabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Slurry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Shape Shape2 
         Height          =   495
         Left            =   840
         Top             =   840
         Width           =   495
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00000040&
         Height          =   615
         Left            =   1560
         Top             =   480
         Width           =   495
      End
      Begin VB.Line slurrytubetubing 
         Index           =   5
         X1              =   2760
         X2              =   3120
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line slurrytubetubing 
         Index           =   3
         X1              =   2760
         X2              =   2760
         Y1              =   1560
         Y2              =   1920
      End
      Begin VB.Line slurrytubetubing 
         Index           =   2
         X1              =   2400
         X2              =   2760
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Shape slurrytube 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   975
         Left            =   3120
         Top             =   240
         Width           =   195
      End
      Begin VB.Line slurrytubetubing 
         Index           =   0
         X1              =   1800
         X2              =   2160
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label ValveLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "V28"
         Height          =   255
         Index           =   20
         Left            =   1920
         TabIndex        =   8
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label ValveClik 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   20
         Left            =   1680
         TabIndex        =   7
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label ValveLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "V29"
         Height          =   255
         Index           =   21
         Left            =   3960
         TabIndex        =   6
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label ValveClik 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   21
         Left            =   3720
         TabIndex        =   5
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label ValveLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pump"
         Height          =   255
         Index           =   22
         Left            =   2040
         TabIndex        =   4
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label ValveClik 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   22
         Left            =   2160
         TabIndex        =   3
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label ValveClik 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   23
         Left            =   1680
         TabIndex        =   2
         Top             =   600
         Width           =   255
      End
      Begin VB.Label ValveLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " Paddle"
         Height          =   255
         Index           =   23
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   735
      End
      Begin VB.Line vopenline 
         Index           =   20
         Visible         =   0   'False
         X1              =   1680
         X2              =   1800
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line vcloseline 
         Index           =   20
         X1              =   1800
         X2              =   1800
         Y1              =   1440
         Y2              =   1320
      End
      Begin VB.Line vcloseline 
         Index           =   23
         X1              =   1920
         X2              =   1680
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line vopenline 
         Index           =   23
         Visible         =   0   'False
         X1              =   1800
         X2              =   1800
         Y1              =   840
         Y2              =   600
      End
      Begin VB.Line vcloseline 
         Index           =   22
         X1              =   2280
         X2              =   2280
         Y1              =   1800
         Y2              =   2040
      End
      Begin VB.Line vopenline 
         Index           =   22
         Visible         =   0   'False
         X1              =   2160
         X2              =   2400
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line vopenline 
         Index           =   21
         Visible         =   0   'False
         X1              =   3840
         X2              =   3840
         Y1              =   1440
         Y2              =   1560
      End
      Begin VB.Line vcloseline 
         Index           =   21
         X1              =   3840
         X2              =   3720
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line slurrytubetubing 
         Index           =   1
         X1              =   1800
         X2              =   1800
         Y1              =   1440
         Y2              =   1920
      End
      Begin VB.Line slurrytubetubing 
         Index           =   4
         X1              =   3840
         X2              =   3840
         Y1              =   1560
         Y2              =   1920
      End
      Begin VB.Line slurrytubetubing 
         Index           =   12
         X1              =   3480
         X2              =   3720
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line slurrytubetubing 
         Index           =   6
         X1              =   3840
         X2              =   3840
         Y1              =   840
         Y2              =   1440
      End
      Begin VB.Shape ValveFill 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   22
         Left            =   2160
         Shape           =   3  'Circle
         Top             =   1800
         Width           =   255
      End
      Begin VB.Shape ValveFill 
         FillColor       =   &H00FFFF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   21
         Left            =   3720
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   255
      End
      Begin VB.Shape ValveFill 
         FillColor       =   &H00FFFF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   20
         Left            =   1680
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   255
      End
      Begin VB.Shape ValveFill 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   23
         Left            =   1680
         Shape           =   3  'Circle
         Top             =   600
         Width           =   255
      End
      Begin VB.Shape ValveFill 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   24
         Left            =   3720
         Shape           =   3  'Circle
         Top             =   600
         Width           =   255
      End
      Begin VB.Shape ValveFill 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   25
         Left            =   2640
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   255
      End
      Begin VB.Line slurrytubetubing 
         Index           =   7
         X1              =   3240
         X2              =   3840
         Y1              =   360
         Y2              =   360
      End
   End
End
Attribute VB_Name = "ScratchPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
