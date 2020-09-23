VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   344
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   925
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSmallExplodeM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   0
      Left            =   7500
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   66
      Top             =   3960
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSmallExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   7
      Left            =   7500
      Picture         =   "Form1.frx":0342
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   65
      Top             =   3480
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSmallExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   6
      Left            =   7500
      Picture         =   "Form1.frx":0CB4
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   64
      Top             =   3000
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSmallExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   5
      Left            =   7500
      Picture         =   "Form1.frx":1626
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   63
      Top             =   2520
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSmallExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   4
      Left            =   7500
      Picture         =   "Form1.frx":1F98
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   62
      Top             =   2040
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSmallExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   3
      Left            =   7500
      Picture         =   "Form1.frx":290A
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   61
      Top             =   1560
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSmallExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   2
      Left            =   7500
      Picture         =   "Form1.frx":327C
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   60
      Top             =   1080
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSmallExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   1
      Left            =   7500
      Picture         =   "Form1.frx":3BEE
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   59
      Top             =   600
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSmallExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   0
      Left            =   7500
      Picture         =   "Form1.frx":4560
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   58
      Top             =   120
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox PicExplode3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   510
      Left            =   3660
      Picture         =   "Form1.frx":4ED2
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   240
      TabIndex        =   57
      Top             =   60
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.PictureBox picPodM 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2400
      Picture         =   "Form1.frx":A374
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   56
      Top             =   60
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox picPod 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2040
      Picture         =   "Form1.frx":ABB6
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   55
      Top             =   60
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Timer timFrame 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   6060
   End
   Begin VB.PictureBox picBlank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   54
      Top             =   660
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   53
      Top             =   0
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picExplodeM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   0
      Left            =   12000
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   52
      Top             =   4320
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   14
      Left            =   12000
      Picture         =   "Form1.frx":B3F8
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   51
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   13
      Left            =   12000
      Picture         =   "Form1.frx":F3FA
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   50
      Top             =   1920
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   12
      Left            =   12000
      Picture         =   "Form1.frx":133FC
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   49
      Top             =   720
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   11
      Left            =   10980
      Picture         =   "Form1.frx":173FE
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   48
      Top             =   4320
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   10
      Left            =   10980
      Picture         =   "Form1.frx":1B400
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   47
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   9
      Left            =   10980
      Picture         =   "Form1.frx":1F402
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   46
      Top             =   1920
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   8
      Left            =   10980
      Picture         =   "Form1.frx":23404
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   45
      Top             =   720
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   7
      Left            =   9960
      Picture         =   "Form1.frx":27406
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   44
      Top             =   4320
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   6
      Left            =   9960
      Picture         =   "Form1.frx":2B408
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   43
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   5
      Left            =   9960
      Picture         =   "Form1.frx":2F40A
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   42
      Top             =   1920
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   4
      Left            =   9960
      Picture         =   "Form1.frx":3340C
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   41
      Top             =   720
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   3
      Left            =   8940
      Picture         =   "Form1.frx":3740E
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   40
      Top             =   4320
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   2
      Left            =   8940
      Picture         =   "Form1.frx":3B410
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   39
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   1
      Left            =   8940
      Picture         =   "Form1.frx":3F412
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   38
      Top             =   1920
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picExplode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   0
      Left            =   8940
      Picture         =   "Form1.frx":43414
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   37
      Top             =   720
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picBossM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   7020
      Picture         =   "Form1.frx":47416
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   36
      Top             =   5760
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox picBoss 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   6180
      Picture         =   "Form1.frx":49500
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   35
      Top             =   5760
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox picBorgM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   5400
      Picture         =   "Form1.frx":4B5EA
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   34
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picBorg 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   4620
      Picture         =   "Form1.frx":4C1CC
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   33
      Top             =   6000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picFalconM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   0
      Left            =   4380
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   32
      Top             =   4560
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picXwingM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   0
      Left            =   2640
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   31
      Top             =   4920
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   16
      Left            =   5880
      Picture         =   "Form1.frx":4CDAE
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   30
      Top             =   4560
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   15
      Left            =   5880
      Picture         =   "Form1.frx":4F118
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   29
      Top             =   4080
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   14
      Left            =   5880
      Picture         =   "Form1.frx":51482
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   28
      Top             =   3600
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   13
      Left            =   5880
      Picture         =   "Form1.frx":537EC
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   27
      Top             =   3120
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   12
      Left            =   5880
      Picture         =   "Form1.frx":55B56
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   26
      Top             =   2640
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   11
      Left            =   5880
      Picture         =   "Form1.frx":57EC0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   25
      Top             =   2160
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   10
      Left            =   5880
      Picture         =   "Form1.frx":5A22A
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   24
      Top             =   1680
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   9
      Left            =   5880
      Picture         =   "Form1.frx":5C594
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   23
      Top             =   1200
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   8
      Left            =   5880
      Picture         =   "Form1.frx":5E8FE
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   7
      Left            =   4380
      Picture         =   "Form1.frx":60C68
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   21
      Top             =   4080
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   6
      Left            =   4380
      Picture         =   "Form1.frx":62FD2
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   20
      Top             =   3600
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   5
      Left            =   4380
      Picture         =   "Form1.frx":6533C
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   4
      Left            =   4380
      Picture         =   "Form1.frx":676A6
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   18
      Top             =   2640
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   3
      Left            =   4380
      Picture         =   "Form1.frx":69A10
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   2
      Left            =   4380
      Picture         =   "Form1.frx":6BD7A
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   16
      Top             =   1680
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   1
      Left            =   4380
      Picture         =   "Form1.frx":6E0E4
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picFalcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   0
      Left            =   4380
      Picture         =   "Form1.frx":7044E
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   13
      Left            =   2640
      Picture         =   "Form1.frx":727B8
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   13
      Top             =   4320
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   12
      Left            =   2640
      Picture         =   "Form1.frx":75072
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   12
      Top             =   3720
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   11
      Left            =   2640
      Picture         =   "Form1.frx":7792C
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   10
      Left            =   2640
      Picture         =   "Form1.frx":7A1E6
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   9
      Left            =   2640
      Picture         =   "Form1.frx":7CAA0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   8
      Left            =   2640
      Picture         =   "Form1.frx":7F35A
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   7
      Left            =   2640
      Picture         =   "Form1.frx":81C14
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   6
      Left            =   1200
      Picture         =   "Form1.frx":844CE
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   6
      Top             =   4320
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   5
      Left            =   1200
      Picture         =   "Form1.frx":86D88
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   4
      Left            =   1200
      Picture         =   "Form1.frx":89642
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   3
      Left            =   1200
      Picture         =   "Form1.frx":8BEFC
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   2
      Left            =   1200
      Picture         =   "Form1.frx":8E7B6
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   1
      Left            =   1200
      Picture         =   "Form1.frx":91070
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picXwing 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Index           =   0
      Left            =   1200
      Picture         =   "Form1.frx":9392A
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub CreateLaser(X01 As Long, Y01 As Long, x02 As Long, Y02 As Long, LC As LaserColor)
Dim x As Long
Dim FS As Long
    FS = -1
    For x = 0 To UBound(Lasers)
        If Lasers(x).Free Then
            FS = x
            Exit For
        End If
    Next x
    If FS = -1 Then
        ReDim Preserve Lasers(UBound(Lasers) + 1)
        FS = UBound(Lasers)
    End If
    With Lasers(FS)
        .Free = False
        .Segment = BuildLaser
        .SegmentCount = 1
        .x1 = X01
        .x2 = x02
        .y1 = Y01
        .y2 = Y02
        .MainColor = LC
    End With
End Sub
Sub DrawLasers()
Dim z As Long
    For z = 0 To UBound(Lasers)
        If Lasers(z).Free = False Then
        
            With Lasers(z)
                
                If .Segment = BuildLaser Then
                    DrawLineToBuffer .x1, .y1, .x1 + (((.x2 - .x1) / 10) * .SegmentCount), .y2, .MainColor
                    .SegmentCount = .SegmentCount + 1
                    If .SegmentCount >= 11 Then
                        .Segment = HazeEffect
                        .SegmentCount = 1
                        .TimeTillNextAction = 20
                        .LastTickCount = GetTickCount
                        .CurrentColor = .MainColor
                    End If
                ElseIf .Segment = HazeEffect And GetTickCount > (.TimeTillNextAction + .LastTickCount) Then
                    DrawLineToBuffer .x1 + 1, .y1 - .SegmentCount, .x2 - 1, .y2 - .SegmentCount, .CurrentColor
                    DrawLineToBuffer .x1, .y1, .x2, .y2, .MainColor
                    DrawLineToBuffer .x1 + 1, .y1 + .SegmentCount, .x2 - 1, .y2 + .SegmentCount, .CurrentColor
                    .CurrentColor = AdjustBrightness(.CurrentColor, -25)
                    .SegmentCount = .SegmentCount + 1
                    If .SegmentCount >= 11 Then
                        .CurrentColor = .MainColor
                        .LastTickCount = GetTickCount
                        .TimeTillNextAction = 30
                        .Segment = DestroyLaser
                        .SegmentCount = 1
                        CreateExplosion .x2, .y2
                    End If
                ElseIf .Segment = DestroyLaser And GetTickCount > (.TimeTillNextAction + .LastTickCount) Then
    
                    .CurrentColor = AdjustBrightness(.CurrentColor, -20)
                    DrawLineToBuffer .x1 + (((.x2 - .x1) / 10) * .SegmentCount), .y1, .x2, .y2, .CurrentColor
                    
                    .SegmentCount = .SegmentCount + 1
                    If .SegmentCount >= 11 Then
                        .Free = True
                    End If
                End If
                    
            End With
        
        End If
    Next
End Sub



Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        If Defender.ChangingDirection Then Exit Sub
        If Defender.Ship = falcon Then
            If Defender.Direction = DefenderRight Then
                CreateLaser Defender.x + picFalcon(0).Width, Defender.y + (picFalcon(0).Height \ 2) + 3, Defender.x + picFalcon(0).Width + LaserLength, Defender.y + (picFalcon(0).Height \ 2) + 3, greenlaser
            Else
                CreateLaser Defender.x, Defender.y + (picFalcon(0).Height \ 2) + 3, Defender.x - LaserLength, Defender.y + (picFalcon(0).Height \ 2) + 3, greenlaser
            End If
        Else
            If Defender.Direction = DefenderRight Then
                CreateLaser Defender.x + picXwing(0).Width - 30, Defender.y + 2, Defender.x + picXwing(0).Width + LaserLength, Defender.y + 2, redlaser
                CreateLaser Defender.x + picXwing(0).Width - 30, Defender.y + 32, Defender.x + picXwing(0).Width + LaserLength, Defender.y + 32, redlaser
            Else
                CreateLaser Defender.x + 30, Defender.y + 2, Defender.x - LaserLength, Defender.y + 2, redlaser
                CreateLaser Defender.x + 30, Defender.y + 32, Defender.x - LaserLength, Defender.y + 32, redlaser
            End If
        End If
    End If
End Sub
Sub CheckKeys()
    If Defender.ChangingDirection Then Exit Sub
    If GetAsyncKeyState(32) <> 0 Then
        Defender.ChangingDirection = True
    End If
    If GetAsyncKeyState(83) <> 0 Then
        Defender.y = Defender.y - Defender.yMovement
        If Defender.y < 0 Then Defender.y = 0
    End If
    If GetAsyncKeyState(88) <> 0 Then
        Defender.y = Defender.y + Defender.yMovement
        If Defender.Ship = falcon Then
            If Defender.y > (picBuffer.Height - picFalcon(0).Height) Then Defender.y = picBuffer.Height - picFalcon(0).Height
        Else
            If Defender.y > (picBuffer.Height - picXwing(0).Height) Then Defender.y = picBuffer.Height - picXwing(0).Height
        End If
    End If
    If Defender.ChangingDirection Then Exit Sub
    
    If GetAsyncKeyState(65) <> 0 Then
        AdjustStarSpeed
        AdjustAnyLasers
    End If
End Sub
Sub AdjustAnyLasers()
Dim z As Long
    For z = 0 To UBound(Lasers)
        If Lasers(z).Free = False Then
            With Lasers(z)
                If Defender.Direction = DefenderRight Then
                    .x1 = .x1 - 3
                    .x2 = .x2 - 3
                Else
                    .x1 = .x1 + 3
                    .x2 = .x2 + 3
                End If
            End With
        End If
    Next
End Sub
Sub AdjustStarSpeed()
Dim z As Long
    For z = 0 To NumStars - 1
        With StarFieldStars(z)
            If Defender.Direction = DefenderRight Then
                .x = .x - .Speed
                If .x < 0 Then .x = picBuffer.Width
            Else
                .x = .x + .Speed
                If .x > picBuffer.Width Then .x = 0

            End If
        End With
    Next
End Sub
Private Sub Form_Load()
    'Me.Move 0, 0, Screen.Width, Screen.Height
    picBuffer.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    picBlank.Move 0, 0, picBuffer.Width, picBuffer.Height
    
    Me.Show
    DoEvents
    
    ClearBuffer
    pTextOut "Initialising...", "Impact", 22, False, 20, 20, RGB(250, 250, 200)
    BufferToScreen

    
    Initialise
    
    'CreateExplosion 20, 20
    'CreateExplosion 100, 100
    
    timFrame.Enabled = True
    
End Sub
Sub Initialise()

    ReDim Explosion(0)
    Explosion(0).Free = True
    
    ReDim SmallExplosion(0)
    SmallExplosion(0).Free = True
    
    ReDim Lasers(0)
    Lasers(0).Free = True
    
    CreateMasks

    With Defender
        .Ship = falcon
        .x = (picBuffer.Width \ 2) - (picFalcon(0).Width \ 2)
        .y = (picBuffer.Height \ 2) - (picFalcon(0).Height \ 2)
        .Direction = DefenderRight
        .ChangingDirection = False
        .CurrentFrame = 0
        .yMovement = 2
    End With
    
    SetupStarField
    
End Sub
Sub SetupStarField()
Dim x As Long
Dim RNDCOL As Long
    ReDim StarFieldStars(NumStars)
    For x = 0 To NumStars - 1
        With StarFieldStars(x)
            .x = Int(Rnd * picBuffer.Width)
            .y = Int(Rnd * picBuffer.Height)
            .Speed = 2 + Int(Rnd * 8)
            RNDCOL = Int(Rnd * 250)
            .StartColor = RGB(RNDCOL, RNDCOL, RNDCOL)
            .CurrentColor = .StartColor
            .ColorDirection = 1 + Int(Rnd * 2)
        End With
    Next
End Sub
Sub CreateMasks()
Dim lF As Long
    For lF = 0 To picXwing.Count - 1
        If lF > 0 Then
            Load picXwingM(lF)
        End If
        CreateMask picXwing(lF).hdc, picXwingM(lF).hdc, picXwing(lF).Width, picXwing(lF).Height
    Next
    For lF = 0 To picFalcon.Count - 1
        If lF > 0 Then
            Load picFalconM(lF)
        End If
        CreateMask picFalcon(lF).hdc, picFalconM(lF).hdc, picFalcon(lF).Width, picFalcon(lF).Height
    Next
    For lF = 0 To picExplode.Count - 1
        If lF > 0 Then
            Load picExplodeM(lF)
        End If
        CreateMask picExplode(lF).hdc, picExplodeM(lF).hdc, picExplode(lF).Width, picExplode(lF).Height
    Next
    For lF = 0 To picSmallExplode.Count - 1
        If lF > 0 Then
            Load picSmallExplodeM(lF)
        End If
        CreateMask picSmallExplode(lF).hdc, picSmallExplodeM(lF).hdc, picSmallExplode(lF).Width, picSmallExplode(lF).Height
    Next
End Sub
Sub CreateMask(picInHDC As Long, picOutHDC As Long, lW As Long, lH As Long)
Dim x As Long
Dim y As Long
    For x = 0 To lW
        For y = 0 To lH
            If GetPixel(picInHDC, x, y) <> 0 Then
                SetPixel picOutHDC, x, y, 0
            End If
        Next
    Next
End Sub
Sub ClearBuffer()
    BitBlt picBuffer.hdc, 0, 0, picBuffer.Width, picBuffer.Height, picBlank.hdc, 0, 0, vbSrcCopy
End Sub
Sub BufferToScreen()
    BitBlt Me.hdc, 0, 0, picBuffer.Width, picBuffer.Height, picBuffer.hdc, 0, 0, vbSrcCopy
    'Me.Refresh
End Sub

Private Sub timFrame_Timer()
    ClearBuffer
    CreateFrameToBuffer
    BufferToScreen
    CheckKeys
End Sub
Sub CreateFrameToBuffer()
    DrawStarField
    
    'draw environment
    'draw enemy
    'draw friendlys
    
    DrawLasers
    DrawDefenderToBuffer
    ExplosionsToBuffer
    TextToBuffer
End Sub
Sub TextToBuffer()
    pTextOut "[A] Thrust  [S] Up  [X] Down  [ENTER] Fire  [SPACE] Change Direction", "Tahoma", 8, True, 5, picBuffer.Height - 20, RGB(51, 100, 51)
End Sub
Sub DrawStarField()
Dim z As Long
Dim Haze As Long
    For z = 0 To NumStars - 1
        With StarFieldStars(z)
            Haze = AdjustBrightness(.CurrentColor, -30, True)
            SetPixel picBuffer.hdc, .x, .y, .CurrentColor
            SetPixel picBuffer.hdc, .x - 1, .y, Haze
            SetPixel picBuffer.hdc, .x, .y - 1, Haze
            SetPixel picBuffer.hdc, .x + 1, .y, Haze
            SetPixel picBuffer.hdc, .x, .y + 1, Haze
            
            If .CurrentColor = 16777215 Then
                .ColorDirection = 2 'dimmer
            ElseIf .CurrentColor = 0 Then
                .ColorDirection = 1 'brighter
                .CurrentColor = RGB(30, 30, 30)
            End If
            If .ColorDirection = 1 Then
                .CurrentColor = AdjustBrightness(.CurrentColor, 3, True)
            Else
                .CurrentColor = AdjustBrightness(.CurrentColor, -3, True)
            End If
        End With
    Next
End Sub
Sub DrawDefenderToBuffer()
    With Defender
        If .Ship = falcon Then
            BitBlt picBuffer.hdc, .x, .y, picFalcon(0).Width, picFalcon(0).Height, picFalconM(.CurrentFrame).hdc, 0, 0, vbSrcAnd
            BitBlt picBuffer.hdc, .x, .y, picFalcon(0).Width, picFalcon(0).Height, picFalcon(.CurrentFrame).hdc, 0, 0, vbSrcPaint
        Else
            BitBlt picBuffer.hdc, .x, .y, picXwing(0).Width, picXwing(0).Height, picXwingM(.CurrentFrame).hdc, 0, 0, vbSrcAnd
            BitBlt picBuffer.hdc, .x, .y, picXwing(0).Width, picXwing(0).Height, picXwing(.CurrentFrame).hdc, 0, 0, vbSrcPaint
        End If
        If .ChangingDirection Then
            If .Direction = DefenderRight Then
                .CurrentFrame = .CurrentFrame + 1
                If .Ship = falcon Then
                    If .CurrentFrame >= picFalcon.Count - 1 Then
                        .ChangingDirection = False
                        .Direction = DefenderLeft
                    End If
                Else
                    If .CurrentFrame >= picXwing.Count - 1 Then
                        .ChangingDirection = False
                        .Direction = DefenderLeft
                    End If
                End If
            Else
                .CurrentFrame = .CurrentFrame - 1
                If .CurrentFrame = 0 Then
                    .ChangingDirection = False
                    .Direction = DefenderRight
                End If
            End If
        End If
    End With
End Sub
Sub ExplosionsToBuffer()
Dim lExp As Long
    For lExp = 0 To UBound(Explosion)
        With Explosion(lExp)
            If Not .Free Then
                BitBlt picBuffer.hdc, .x, .y, picExplode(.CurrentFrame).Width, picExplode(.CurrentFrame).Height, picExplodeM(.CurrentFrame).hdc, 0, 0, vbSrcAnd
                BitBlt picBuffer.hdc, .x, .y, picExplode(.CurrentFrame).Width, picExplode(.CurrentFrame).Height, picExplode(.CurrentFrame).hdc, 0, 0, vbSrcPaint
                If (.LastFrameTime + .TimeTillNextFrame) <= GetTickCount Then
                    .LastFrameTime = GetTickCount
                    .CurrentFrame = .CurrentFrame + 1
                    If .CurrentFrame > picExplode.Count - 1 Then
                        .Free = True
                    End If
                End If
            End If
        End With
    Next
    For lExp = 0 To UBound(SmallExplosion)
        With SmallExplosion(lExp)
            If Not .Free Then
                BitBlt picBuffer.hdc, .x, .y, picSmallExplode(.CurrentFrame).Width, picSmallExplode(.CurrentFrame).Height, picSmallExplodeM(.CurrentFrame).hdc, 0, 0, vbSrcAnd
                BitBlt picBuffer.hdc, .x, .y, picSmallExplode(.CurrentFrame).Width, picSmallExplode(.CurrentFrame).Height, picSmallExplode(.CurrentFrame).hdc, 0, 0, vbSrcPaint
                If (.LastFrameTime + .TimeTillNextFrame) <= GetTickCount Then
                    .LastFrameTime = GetTickCount
                    .CurrentFrame = .CurrentFrame + 1
                    If .CurrentFrame > picSmallExplode.Count - 1 Then
                        .Free = True
                    End If
                End If
            End If
        End With
    Next

End Sub


Sub pTextOut(sIn As String, sFont As String, iFontSize As Integer, bFontBold As Boolean, xPos As Integer, yPos As Integer, lColor As Long)
    
    SetTextColor picBuffer.hdc, 0
    picBuffer.Font = sFont
    picBuffer.FontSize = iFontSize
    picBuffer.FontBold = bFontBold
    
    TextOut picBuffer.hdc, xPos + 1, yPos + 1, sIn, Len(sIn)
    TextOut picBuffer.hdc, xPos - 1, yPos - 1, sIn, Len(sIn)
    TextOut picBuffer.hdc, xPos - 1, yPos + 1, sIn, Len(sIn)
    TextOut picBuffer.hdc, xPos + 1, yPos - 1, sIn, Len(sIn)
    TextOut picBuffer.hdc, xPos - 1, yPos, sIn, Len(sIn)
    TextOut picBuffer.hdc, xPos + 1, yPos, sIn, Len(sIn)
    
    SetTextColor picBuffer.hdc, lColor
    TextOut picBuffer.hdc, xPos, yPos, sIn, Len(sIn)

End Sub
Sub DrawLineToBuffer(x1 As Long, y1 As Long, x2 As Long, y2 As Long, lColor As Long)
Dim hRPen As Long
Dim Point As POINTAPI
    picBuffer.ForeColor = lColor
    Point.x = x1: Point.y = y1
    MoveToEx picBuffer.hdc, x1, y1, Point
    LineTo picBuffer.hdc, x2, y2
End Sub
Sub ApplyBlend()
Dim Blend As BLENDFUNCTION
Dim BlendPtr As Long
    Blend.SourceConstantAlpha = 130
    
    CopyMemory BlendPtr, Blend, 4
    
    AlphaBlend picBuffer.hdc, 0, 0, picBuffer.Width, picBuffer.Height, picBlank.hdc, 0, 0, picBuffer.Width, picBuffer.Height, BlendPtr
End Sub
Sub CreateExplosion(xP As Long, yP As Long)
Dim bFreeSlot As Boolean
Dim a As Long
    bFreeSlot = False
    For a = 0 To UBound(Explosion)
        If Explosion(a).Free = True Then
            bFreeSlot = True
            Exit For
        End If
    Next
    If Not bFreeSlot Then
        ReDim Preserve Explosion(UBound(Explosion) + 1)
        a = UBound(Explosion)
    End If
    
    With Explosion(a)
        .Free = False
        .x = xP - 30 'adjust coords of explosion to appear in middle
        .y = yP - 54
        .TimeTillNextFrame = 50
        .LastFrameTime = GetTickCount
        .CurrentFrame = 0
    End With
    
End Sub

Sub CreateSmallExplosion(xP As Long, yP As Long)
Dim bFreeSlot As Boolean
Dim a As Long
    bFreeSlot = False
    For a = 0 To UBound(SmallExplosion)
        If SmallExplosion(a).Free = True Then
            bFreeSlot = True
            Exit For
        End If
    Next
    If Not bFreeSlot Then
        ReDim Preserve SmallExplosion(UBound(SmallExplosion) + 1)
        a = UBound(SmallExplosion)
    End If
    
    With SmallExplosion(a)
        .Free = False
        .x = xP - 14 'adjust coords of explosion to appear in middle
        .y = yP - 14
        .TimeTillNextFrame = 50
        .LastFrameTime = GetTickCount
        .CurrentFrame = 0
    End With
    
End Sub
