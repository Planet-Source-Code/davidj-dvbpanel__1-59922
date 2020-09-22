VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmtest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test"
   ClientHeight    =   4485
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin DVBPanelTesting.DVBPanel DVBPanel4 
      Height          =   915
      Left            =   0
      TabIndex        =   17
      Top             =   3600
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   1614
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   35
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button35 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
   End
   Begin DVBPanelTesting.DVBPanel DVBPanel1 
      Height          =   1635
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2884
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3315
         Left            =   120
         ScaleHeight     =   3315
         ScaleWidth      =   1395
         TabIndex        =   18
         Top             =   0
         Width           =   1395
         Begin VB.OptionButton Option12 
            Caption         =   "Option12"
            Height          =   255
            Left            =   60
            TabIndex        =   30
            Top             =   2940
            Width           =   1035
         End
         Begin VB.OptionButton Option11 
            Caption         =   "Option11"
            Height          =   255
            Left            =   60
            TabIndex        =   29
            Top             =   2670
            Width           =   1035
         End
         Begin VB.OptionButton Option10 
            Caption         =   "Option10"
            Height          =   255
            Left            =   60
            TabIndex        =   28
            Top             =   2415
            Width           =   1035
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Option9"
            Height          =   255
            Left            =   60
            TabIndex        =   27
            Top             =   2160
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Left            =   60
            TabIndex        =   26
            Top             =   120
            Width           =   915
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Option2"
            Height          =   255
            Left            =   60
            TabIndex        =   25
            Top             =   375
            Width           =   915
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Option3"
            Height          =   255
            Left            =   60
            TabIndex        =   24
            Top             =   630
            Width           =   915
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Option4"
            Height          =   255
            Left            =   60
            TabIndex        =   23
            Top             =   885
            Width           =   915
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Option5"
            Height          =   255
            Left            =   60
            TabIndex        =   22
            Top             =   1140
            Width           =   915
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Option6"
            Height          =   255
            Left            =   60
            TabIndex        =   21
            Top             =   1395
            Width           =   915
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Option7"
            Height          =   255
            Left            =   60
            TabIndex        =   20
            Top             =   1650
            Width           =   915
         End
         Begin VB.OptionButton Option8 
            Caption         =   "Option8"
            Height          =   255
            Left            =   60
            TabIndex        =   19
            Top             =   1905
            Width           =   915
         End
      End
   End
   Begin DVBPanelTesting.DVBPanel DVBPanel2 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   3201
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   5295
         Left            =   120
         ScaleHeight     =   5295
         ScaleWidth      =   2880
         TabIndex        =   2
         Top             =   0
         Width           =   2875
         Begin VB.CommandButton Command12 
            Caption         =   "Command12"
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   4800
            Width           =   1575
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Command11"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   4365
            Width           =   1575
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Command10"
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   3945
            Width           =   1575
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Command9"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   3525
            Width           =   1575
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Command8"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   3090
            Width           =   1575
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Command7"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   2670
            Width           =   1575
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Command6"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   2250
            Width           =   1575
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Command5"
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   1815
            Width           =   1575
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Command4"
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   1395
            Width           =   1575
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   975
            Width           =   1575
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   540
            Width           =   2575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   2575
         End
      End
   End
   Begin DVBPanelTesting.DVBPanel DVBPanel3 
      Height          =   1815
      Left            =   2280
      TabIndex        =   1
      Top             =   1680
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3201
      Begin VB.TextBox txtTest 
         Height          =   390
         Index           =   0
         Left            =   240
         MaxLength       =   12
         TabIndex        =   15
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Menu mnFile 
      Caption         =   "File"
   End
   Begin VB.Menu mnView 
      Caption         =   "View"
   End
End
Attribute VB_Name = "frmtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Activate()
    'To show how the focus can change the location of the scroll bars
    txtTest(35).SetFocus
End Sub

Private Sub Form_Load()
    'To show how the focus can change the location of the scroll bars
    'Tab through the array of textboxes to see the scroll bars move
    ' to the control that has focus
    Dim ndx As Integer
    For ndx = 1 To 50
        Load txtTest(ndx)
        txtTest(ndx).Visible = True
        txtTest(ndx).Top = txtTest(ndx - 1).Top + txtTest(ndx - 1).Height + 100
        txtTest(ndx).Left = txtTest(ndx - 1).Left
        txtTest(ndx).Text = "TEST " & ndx
    Next
End Sub
