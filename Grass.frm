VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Grass generator: Created by Johannes.B 2000"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   387
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CM 
      Left            =   1080
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6975
      Begin VB.CommandButton Command2 
         Caption         =   "Save as..."
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   5640
         MaxLength       =   3
         TabIndex        =   14
         Text            =   "0"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   5160
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "100"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   4680
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "0"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3600
         MaxLength       =   7
         TabIndex        =   7
         Text            =   "100000"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2520
         MaxLength       =   6
         TabIndex        =   5
         Text            =   "30000"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   3
         Text            =   "10000"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "DRAW!"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Random up to.. RGB"
         Height          =   255
         Left            =   4560
         TabIndex        =   15
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "3."
         Height          =   255
         Left            =   3420
         TabIndex        =   11
         Top             =   420
         Width           =   135
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "2."
         Height          =   255
         Left            =   2370
         TabIndex        =   10
         Top             =   420
         Width           =   135
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         Height          =   255
         Left            =   1275
         TabIndex        =   9
         Top             =   405
         Width           =   135
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dots"
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tiny circles"
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Big circles"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00004000&
      Height          =   4575
      Left            =   0
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   461
      TabIndex        =   0
      Top             =   840
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Caption = "PLEASE WAIT... STEP 1/3"
For Counter = 1 To Text1.Text Step 1
Picture1.ForeColor = RGB(Rnd * Text4.Text, Rnd * Text5.Text, Rnd * Text6.Text)
Picture1.Circle (Rnd * Picture1.Width, Rnd * Picture1.Height), Rnd * 50
Next
Form1.Caption = "PLEASE WAIT... STEP 2/3"
For Counter = 1 To Text2.Text Step 1
Picture1.ForeColor = RGB(Rnd * Text4.Text, Rnd * Text5.Text, Rnd * Text6.Text)
Picture1.Circle (Rnd * Picture1.Width, Rnd * Picture1.Height), Rnd * 8
Next
Form1.Caption = "PLEASE WAIT... STEP 3/3"
For Counter = 1 To Text3.Text Step 1
Picture1.ForeColor = RGB(Rnd * Text4.Text, Rnd * Text5.Text, Rnd * Text6.Text)
Picture1.PSet (Rnd * Picture1.Width, Rnd * Picture1.Height)
Next
Form1.Caption = "Grass generator: Created by Johannes.B 2000"
End Sub

Private Sub Command2_Click()

CM.CancelError = True
On Error GoTo Sväng
   
CM.Flags = cdlOFNHideReadOnly
CM.Filter = "Windows bitmap (*.BMP)|*.bmp"

CM.ShowSave
SavePicture Picture1.Image, CM.FileName
Exit Sub
Sväng:
Exit Sub
End Sub


Private Sub Form_Resize()
Picture1.Width = Form1.ScaleWidth - 2
Picture1.Height = Form1.ScaleHeight - 57
End Sub

Private Sub Text1_Change()
On Error GoTo nisse
If Text1.Text < 1 Then Text1.Text = 1
Exit Sub
nisse:
Text1.Text = 1
Exit Sub
End Sub


Private Sub Text2_Change()
On Error GoTo nisse
If Text2.Text < 1 Then Text2.Text = 1
Exit Sub
nisse:
Text2.Text = 1
Exit Sub
End Sub


Private Sub Text3_Change()
On Error GoTo nisse
If Text3.Text < 1 Then Text3.Text = 1
Exit Sub
nisse:
Text3.Text = 1
Exit Sub
End Sub


Private Sub Text4_Change()
On Error GoTo balle
If Text4.Text > 255 Or Text4.Text < 0 Then Text4.Text = 255
Exit Sub
balle:
Text4.Text = 1
Exit Sub
End Sub


Private Sub Text5_Change()
On Error GoTo balle
If Text5.Text > 255 Or Text5.Text < 0 Then Text5.Text = 255
Exit Sub
balle:
Text5.Text = 1
Exit Sub
End Sub


Private Sub Text6_Change()
On Error GoTo balle
If Text6.Text > 255 Or Text6.Text < 0 Then Text6.Text = 255
Exit Sub
balle:
Text6.Text = 1
Exit Sub
End Sub


