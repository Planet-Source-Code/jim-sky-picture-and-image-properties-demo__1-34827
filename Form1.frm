VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Picture - Image Properties Demo"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClearPicture 
      Caption         =   "Erase the Picture Property"
      Height          =   315
      Left            =   3660
      TabIndex        =   16
      Top             =   4620
      Width           =   2685
   End
   Begin VB.CommandButton cmdSaveImage 
      Caption         =   "Save Image Property to Bitmap File"
      Height          =   345
      Left            =   450
      TabIndex        =   15
      Top             =   3570
      Width           =   2685
   End
   Begin VB.CommandButton cmdGetPic1Image 
      Caption         =   "Get Pic1 Image and Picture with BitBlt"
      Height          =   435
      Left            =   3660
      TabIndex        =   10
      Top             =   3840
      Width           =   2685
   End
   Begin VB.CommandButton cmdGetPic1 
      Caption         =   "Get Pic1 Picture with Picture Property"
      Height          =   405
      Left            =   3660
      TabIndex        =   7
      Top             =   3420
      Width           =   2685
   End
   Begin VB.CommandButton cmdDrawPic2 
      Caption         =   "Draw a Smaller Circle on Top"
      Height          =   315
      Left            =   3660
      TabIndex        =   6
      Top             =   3090
      Width           =   2685
   End
   Begin VB.CommandButton cmdClearPic2 
      Caption         =   "Clear the Image using Cls "
      Height          =   315
      Left            =   3660
      TabIndex        =   5
      Top             =   4290
      Width           =   2685
   End
   Begin VB.CommandButton cmdSavePic 
      Caption         =   "Save Picture Property to Bitmap File"
      Height          =   345
      Left            =   450
      TabIndex        =   4
      Top             =   3180
      Width           =   2685
   End
   Begin VB.CommandButton CmdLoadPic2 
      Caption         =   "Load Bitmap File Saved from Pic1"
      Height          =   345
      Left            =   3660
      TabIndex        =   3
      Top             =   2730
      Width           =   2685
   End
   Begin VB.CommandButton CmdDrawPic1 
      Caption         =   "Draw Circle on Pic1"
      Height          =   345
      Left            =   450
      TabIndex        =   2
      Top             =   2790
      Width           =   2685
   End
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      Height          =   2025
      Left            =   3630
      ScaleHeight     =   1965
      ScaleWidth      =   2625
      TabIndex        =   1
      Top             =   420
      Width           =   2685
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      Height          =   2040
      Left            =   450
      ScaleHeight     =   1980
      ScaleWidth      =   2640
      TabIndex        =   0
      Top             =   420
      Width           =   2700
   End
   Begin VB.Label Label7 
      Caption         =   "Erase everything using Pic2.Picture = LoadPicture( )"
      Height          =   585
      Left            =   6630
      TabIndex        =   17
      Top             =   4470
      Width           =   2205
   End
   Begin VB.Label Label6 
      Caption         =   $"Form1.frx":0442
      Height          =   1215
      Left            =   6630
      TabIndex        =   14
      Top             =   390
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   $"Form1.frx":04E0
      Height          =   1005
      Left            =   6630
      TabIndex        =   13
      Top             =   3450
      Width           =   2145
   End
   Begin VB.Label Label4 
      Caption         =   "Pic 1"
      Height          =   225
      Left            =   1470
      TabIndex        =   12
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Pic 2"
      Height          =   225
      Left            =   4740
      TabIndex        =   11
      Top             =   2490
      Width           =   705
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":0568
      Height          =   1815
      Left            =   6630
      TabIndex        =   9
      Top             =   1590
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0656
      Height          =   1545
      Left            =   450
      TabIndex        =   8
      Top             =   3960
      Width           =   2715
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdClearPic2_Click()
Pic2.Cls
End Sub

Private Sub cmdClearPicture_Click()
'This will erase the background picture and (unfortuneately) the Image drawn too
Pic2.Picture = LoadPicture()


End Sub

Private Sub CmdDrawPic1_Click()
Pic1.Circle (Pic1.ScaleWidth / 2, Pic1.ScaleHeight / 2), Pic1.ScaleHeight / 4

End Sub



Private Sub cmdDrawPic2_Click()
'draw a smaller circle
Pic2.Circle (Pic2.ScaleWidth / 2, Pic2.ScaleHeight / 2), Pic2.ScaleHeight / 5
End Sub

Private Sub cmdGetPic1_Click()
Pic2.Picture = Pic1.Picture
'if you try Pic2.Image = Pic1.Image you get an error


End Sub

Private Sub cmdGetPic1Image_Click()
Dim rtn As Long
Dim pw As Long ' pixels wide
Dim ph As Long  ' pixels high
'Note the AutoRedraw property of the destination picture must be set to True
Pic2.AutoRedraw = True
Pic2.ScaleMode = 3  ' we need the width in pixels for the  BitBlt routine
pw = Pic2.ScaleWidth
ph = Pic2.ScaleHeight
Pic2.ScaleMode = 0 ' or whatever mode you were using
rtn = BitBlt(Pic2.hDC, 0, 0, pw, ph, Pic1.hDC, 0, 0, vbSrcCopy)
Pic2.Refresh
End Sub

Private Sub CmdLoadPic2_Click()
Pic2.Picture = LoadPicture(App.Path + "\testpic.bmp")
End Sub

Private Sub cmdSaveImage_Click()
'notice that with SavePicture we use the Pic1.Image property Not Pic1.Picture
'if we want to transfer the background picture and the Image we have drawn on it.
SavePicture Pic1.Image, App.Path + "\testpic.bmp"

End Sub

Private Sub cmdSavePic_Click()
'save to file just the background Picture property
SavePicture Pic1.Picture, App.Path + "\testpic.bmp"

End Sub

Private Sub Form_Load()

'you can load a bitmap programmaticallly like this..
Pic1.Picture = LoadPicture(App.Path + "/GBT.jpg")
End Sub
