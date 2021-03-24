VERSION 5.00
Begin VB.Form ATE 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ATE"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   14685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmd_instsel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "OSC +  50 Ohm "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2520
      Width           =   1335
   End
   Begin VB.ComboBox ComboLT2BOFF 
      Height          =   315
      Left            =   7200
      TabIndex        =   33
      Text            =   "LT2B TERM OFF"
      Top             =   7800
      Width           =   1575
   End
   Begin VB.ComboBox ComboLT2BON 
      Height          =   315
      Left            =   5400
      TabIndex        =   32
      Text            =   "LT2B TERM ON"
      Top             =   7800
      Width           =   1695
   End
   Begin VB.ComboBox ComboLT2AOFF 
      Height          =   315
      ItemData        =   "ATEnew.frx":0000
      Left            =   3480
      List            =   "ATEnew.frx":0002
      TabIndex        =   31
      Text            =   "LT2A TERM OFF"
      Top             =   7800
      Width           =   1695
   End
   Begin VB.ComboBox ComboLT1 
      Height          =   315
      Left            =   240
      TabIndex        =   30
      Text            =   "5V Check"
      Top             =   7800
      Width           =   1215
   End
   Begin VB.ComboBox ComboLT2AON 
      Height          =   315
      Left            =   1680
      TabIndex        =   29
      Text            =   "LT2A TERM ON"
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton Cmd_cpin2 
      Caption         =   "LT2A-CPIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   28
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_ine2 
      Caption         =   "LT2A-INE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   27
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_cpin 
      Caption         =   "LT1-CPIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   26
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_encb_LT2QE 
      Caption         =   "LT2-EN2 Quad ended"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   24
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Cmd_enca_LT2QE 
      Caption         =   "LT2-EN1 Quad ended"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   23
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Cmd_encb_LT2SE 
      Caption         =   "LT2-EN2 single ended"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   22
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Cmd_enca_LT2SE 
      Caption         =   "LT2-EN1 single ended"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   21
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Cmd_encb_LT1 
      Caption         =   "LT1-EN2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   20
      Top             =   4320
      Width           =   975
   End
   Begin VB.Frame Frame5 
      Caption         =   "Voltage Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   7200
      Width           =   8775
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   1440
         X2              =   1440
         Y1              =   360
         Y2              =   1560
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "INE/CPIN Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   17
      Top             =   5400
      Width           =   8775
      Begin VB.CommandButton Cmd_ine 
         Caption         =   "LT1-INE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   4200
         X2              =   4200
         Y1              =   120
         Y2              =   1320
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Encoder Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      Width           =   8775
      Begin VB.CommandButton Cmd_enca_LT1 
         Caption         =   "LT1-EN1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   975
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   2520
         X2              =   2520
         Y1              =   120
         Y2              =   1320
      End
   End
   Begin VB.CommandButton Cmd_instsel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Oscilloscope"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_instsel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Multimeter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Instrument Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   11175
      Begin VB.CommandButton Cmd_instsel 
         BackColor       =   &H00C0C0FF&
         Caption         =   "SigGen + 50 Ohm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Cmd_instsel 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Sig Gen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton Cmd_channel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Channel 8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Cmd_channel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Channel 7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Cmd_channel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Channel 6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Cmd_channel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Channel 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Cmd_channel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Channel 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Cmd_channel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Channel 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Cmd_channel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Channel 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Channel Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   11175
      Begin VB.CommandButton Cmd_channel 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Channel 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton Cmd_reset 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Reset ATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6960
      Width           =   2175
   End
   Begin VB.CommandButton Cmd_disconnect 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Disconnect ATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton Cmd_connect 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Connect ATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   2175
   End
End
Attribute VB_Name = "ATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Maps the channel selected to the relays we need to turn on:

'Channel 1   :   Port B1,B3,B7+ Toggle A6
'Channel 2   :   Port B1,B3 + Toggle A6
'Channel 3   :   Port B1,B5+ Toggle A6
'Channel 4   :   Port B1+ Toggle A6
'Channel 5   :   Port B4,B6+ Toggle A6
'Channel 6   :   Port B6+ Toggle A6
'Channel 7   :   Port B2+ Toggle A6
'Channel 8   :   None (All zeros)+ Toggle A6
'Sig Gen     :   Port B6 + Toggle A7
'Meter       :   Port B1 + Toggle A7
'Scope       :   None (All zeros) + Toggle A7
'50 OHM Load :   Port B3 + Toggle A7
'=== ABSOLUTELY MUST NOT CHANGE CHANNEL WHILST PULSING ===



Private Sub Cmd_channel_Click(Index As Integer)
Dim intLoop As Integer
'only one channel allowed to be enabled at a time
For intLoop = 1 To 8
    If Cmd_channel(intLoop).BackColor = &HC0FFC0 Then    'backcolour is green
        Cmd_channel(intLoop).BackColor = &HC0C0FF        'set the backcolor to red
    End If
Next

Cmd_channel(Index).BackColor = &HC0FFC0      'set channel selected to green
WriteMicropulse ("STX")
Select Case Index

    Case 1
        ATEWritePort 1, 138
        Delay 100
        ATEWritePort 0, 0
        Delay 100
        ATEWritePort 0, 64
        
    Case 2
        ATEWritePort 1, 10
        Delay 100
        ATEWritePort 0, 0
        Delay 100
        ATEWritePort 0, 64
        
    Case 3
        ATEWritePort 1, 34
        Delay 100
        ATEWritePort 0, 0
        Delay 100
        ATEWritePort 0, 64
        
    Case 4
        ATEWritePort 1, 2
        Delay 100
        ATEWritePort 0, 0
        Delay 100
        ATEWritePort 0, 64
        
    Case 5
        ATEWritePort 1, 80
        Delay 100
        ATEWritePort 0, 0
        Delay 100
        ATEWritePort 0, 64
        
    Case 6
        ATEWritePort 1, 64
        Delay 100
        ATEWritePort 0, 0
        Delay 100
        ATEWritePort 0, 64
        
    Case 7
        ATEWritePort 1, 4
        Delay 100
        ATEWritePort 0, 0
        Delay 100
        ATEWritePort 0, 64
        
    Case 8
        ATEWritePort 1, 0
        Delay 100
        ATEWritePort 0, 0
        Delay 100
        ATEWritePort 0, 64
End Select

Delay 250   'short delay to allow for relay settling

End Sub




Private Sub Cmd_connect_Click()
    
If Not g_blnATEConnected Then

    If ATEConnect = True Then
        g_blnATEConnected = True
       MsgBox ("ATE Connected")
       
    Else
        g_blnATEConnected = False
        MsgBox ("Failed to connect to ATE")
    End If

End If

End Sub

Private Sub Cmd_cpin_Click()

Dim bytValue As Byte

'' Port B and Port C needs to be zero first for this port B and port C needs to be zero and toggle A4,A5,A6,A7

ATEWritePort 1, 0
Delay 100
ATEWritePort 2, 0
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 16    '' Toggles A4
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 32     ''Toggles A5
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 64      ''Toggles A6
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 128     ''Toggles A7
Delay 100


'' Checking cpin on LT1

ATEWritePort 1, 8       '' connect the Pin
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 16      ''Toggles A4
Delay 100
WriteMicropulse ("CENA 1")
Delay 200
ATEWritePort 2, 0        '' C2=0
Delay 100
WriteMicropulse ("CPIN 0 0")

'' Read the D4 (PORT D )

ATEReadPort 3, bytValue
If bytValue = 16 Then      ''D4 should be 1 OR 0 HERE* NOT SURE CHECK WHEN TESTING
MsgBox (" LT1 CPIN PIn checked for CPIN 0  0, Result=OK ")
Else
MsgBox ("LT1 CPIN Pin is not Working ")
End If

WriteMicropulse ("CPIN 0 1")

'' Read the D4 (PORT D )

ATEReadPort 3, bytValue
If bytValue = 0 Then      ''D4 should be 1 OR 0 HERE* NOT SURE CHECK WHEN TESTING
MsgBox (" LT1 CPIN PIn checked for CPIN 0 1 , Result=OK ")
Else
MsgBox ("LT1 CPIN Pin is not Working ")
End If
End Sub

Private Sub Cmd_cpin2_Click()
Dim bytValue As Byte

'' Port B and Port C needs to be zero first for this port B and port C needs to be zero and toggle A4,A5,A6,A7

ATEWritePort 1, 0
Delay 100
ATEWritePort 2, 0
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 16    '' Toggles A4
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 32     ''Toggles A5
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 64      ''Toggles A6
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 128     ''Toggles A7
Delay 100


'' Checking cpin on LT2

ATEWritePort 1, 0      '' connect the Pin
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 16      ''Toggles A4
Delay 100
WriteMicropulse ("CENA 1")
Delay 200
ATEWritePort 2, 0        '' C2=0
Delay 100
WriteMicropulse ("CPIN 0 0")

'' Read the D4 (PORT D )

ATEReadPort 3, bytValue
If bytValue = 16 Then      ''D4 should be 1 OR 0 HERE* NOT SURE CHECK WHEN TESTING
MsgBox (" LT2 CPIN PIn checked for CPIN 0  0, Result=OK ")
Else
MsgBox ("LT2 CPIN Pin is not Working ")
End If

WriteMicropulse ("CPIN 0 1")

'' Read the D4 (PORT D )

ATEReadPort 3, bytValue
If bytValue = 0 Then      ''D4 should be 1 OR 0 HERE* NOT SURE CHECK WHEN TESTING
MsgBox (" LT2 CPIN PIn checked for CPIN 0 1 , Result=OK ")
Else
MsgBox ("LT2 CPIN Pin is not Working ")
End If

End Sub

Private Sub Cmd_disconnect_Click()
If g_blnATEConnected Then
    
    If ATEDisconnect = True Then
        g_blnATEConnected = False
        MsgBox ("ATE Disconnected")
    Else
        MsgBox ("Error Disconnecting ATE"), vbExclamation
    End If
    
End If
End Sub

Private Sub Cmd_enca_LT1_Click()

MsgBox ("In MPComms Type STA Then do shift+enter")

WriteMicropulse ("LCP 1 0")
Delay 100
WriteMicropulse ("LCP 2 0")
Delay 100

'' Connecting pin 3 and 4 to CP1 and CP2

ATEWritePort 2, 8    '' C3=1,C4=0
Delay 100
ATEWritePort 1, 229
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 36
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100

'' Trigger 10 times
Dim counter As Integer
counter = 1
For counter = 1 To 10
ATEWritePort 2, 128    ''C7=1
Delay 250
ATEWritePort 2, 0      ''C7=0
Delay 1000
counter = counter + 1
Next
Delay 100
'' Connecting pin 3 and 4 to CP2 and CP1

ATEWritePort 2, 8    '' C3=1,C4=0
Delay 100
ATEWritePort 1, 197
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 36
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100

'' Trigger 10 times
Dim counter2 As Integer
counter2 = 1
For counter2 = 1 To 10
ATEWritePort 2, 128    ''C7=1
Delay 250
ATEWritePort 2, 0      ''C7=0
Delay 1000
counter2 = counter2 + 1
Next
Delay 100
MsgBox ("LT1-EN1 Checked")
End Sub

Private Sub Cmd_enca_LT2QE_Click()

MsgBox ("In MPComms Type STA Then do shift+enter")

WriteMicropulse ("LCP 1 0")
Delay 100
WriteMicropulse ("LCP 2 0")
Delay 100

''pin 2,3,4 and 5 needs to be connnected to CP1+, CP1-, CP2+ and CP2-.

ATEWritePort 2, 8    '' C3=1,C4=0
Delay 100
ATEWritePort 1, 81
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 89
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100

'' Trigger 10 times
Dim counter As Integer
counter = 1
For counter = 1 To 10
ATEWritePort 2, 128    ''C7=1
Delay 250
ATEWritePort 2, 0      ''C7=0
Delay 1000
counter = counter + 1
Next
Delay 100

''pin 4,5,2 and 3 needs to be connnected to CP1+, CP1-, CP2+ and CP2-.

ATEWritePort 2, 8    '' C3=1,C4=0
Delay 100
ATEWritePort 1, 65
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 25
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100

'' Trigger 10 times
Dim counter2 As Integer
counter2 = 1
For counter2 = 1 To 10
ATEWritePort 2, 128    ''C7=1
Delay 250
ATEWritePort 2, 0      ''C7=0
Delay 1000
counter2 = counter2 + 1
Next
Delay 100
MsgBox ("Quad Ended LT2-EN1 Checked")
End Sub

Private Sub Cmd_enca_LT2SE_Click()

MsgBox ("In MPComms Type STA Then do shift+enter")

WriteMicropulse ("LCP 1 0")
Delay 100
WriteMicropulse ("LCP 2 0")
Delay 100

'' Connecting pin 2 and 4 to CP1 and CP2

ATEWritePort 2, 8    '' C3=1,C4=0
Delay 100
ATEWritePort 1, 229
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 4
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100

'' Trigger 10 times
Dim counter As Integer
counter = 1
For counter = 1 To 10
ATEWritePort 2, 128    ''C7=1
Delay 250
ATEWritePort 2, 0      ''C7=0
Delay 1000
counter = counter + 1
Next
Delay 100
'' Connecting pin 2 and 4 to CP2 and CP1

ATEWritePort 2, 8    '' C3=1,C4=0
Delay 100
ATEWritePort 1, 197
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 4
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100

'' Trigger 10 times
Dim counter2 As Integer
counter2 = 1
For counter2 = 1 To 10
ATEWritePort 2, 128    ''C7=1
Delay 250
ATEWritePort 2, 0      ''C7=0
Delay 1000
counter2 = counter2 + 1
Next
Delay 100
MsgBox ("Single ended LT2-EN1 Checked")
End Sub

Private Sub Cmd_encb_LT1_Click()

MsgBox ("In MPComms Type STA Then do shift+enter")

WriteMicropulse ("LCP 1 0")
Delay 100
WriteMicropulse ("LCP 2 0")
Delay 100

'' Connecting pin 2 and 5 to CP1 and CP2

ATEWritePort 2, 8    '' C3=1,C4=0
Delay 100
ATEWritePort 1, 229
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 0
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100

'' Trigger 10 times
Dim counter As Integer
counter = 1
For counter = 1 To 10
ATEWritePort 2, 128    ''C7=1
Delay 250
ATEWritePort 2, 0      ''C7=0
Delay 1000
counter = counter + 1
Next
Delay 100
'' Connecting pin 2 and 5 to CP2 and CP1

ATEWritePort 2, 8    '' C3=1,C4=0
Delay 100
ATEWritePort 1, 197
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 0
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100

'' Trigger 10 times
Dim counter2 As Integer
counter2 = 1
For counter2 = 1 To 10
ATEWritePort 2, 128    ''C7=1
Delay 250
ATEWritePort 2, 0      ''C7=0
Delay 1000
counter2 = counter2 + 1
Next
Delay 100
MsgBox ("LT1-EN2 Checked")
End Sub

Private Sub Cmd_encb_LT2QE_Click()

MsgBox ("In MPComms Type STA Then do shift+enter")

WriteMicropulse ("LCP 1 0")
Delay 100
WriteMicropulse ("LCP 2 0")
Delay 100

''pin 2,3,4 and 5 needs to be connnected to CP1+, CP1-, CP2+ and CP2-.

ATEWritePort 2, 8    '' C3=1,C4=0
Delay 100
ATEWritePort 1, 17
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 89
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100

'' Trigger 10 times
Dim counter As Integer
counter = 1
For counter = 1 To 10
ATEWritePort 2, 128    ''C7=1
Delay 250
ATEWritePort 2, 0      ''C7=0
Delay 1000
counter = counter + 1
Next
Delay 100

''pin 4,5,2 and 3 needs to be connnected to CP1+, CP1-, CP2+ and CP2-.

ATEWritePort 2, 8    '' C3=1,C4=0
Delay 100
ATEWritePort 1, 1
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 25
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100

'' Trigger 10 times
Dim counter2 As Integer
counter2 = 1
For counter2 = 1 To 10
ATEWritePort 2, 128    ''C7=1
Delay 250
ATEWritePort 2, 0      ''C7=0
Delay 1000
counter2 = counter2 + 1
Next
Delay 100
MsgBox ("Quad Ended LT2-EN2 Checked")
End Sub

Private Sub Cmd_encb_LT2SE_Click()

MsgBox ("In MPComms Type STA Then do shift+enter")

WriteMicropulse ("LCP 1 0")
Delay 100
WriteMicropulse ("LCP 2 0")
Delay 100

'' Connecting pin 2 and 4 to CP1 and CP2

ATEWritePort 2, 8    '' C3=1,C4=0
Delay 100
ATEWritePort 1, 165
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 4
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100

'' Trigger 10 times
Dim counter As Integer
counter = 1
For counter = 1 To 10
ATEWritePort 2, 128    ''C7=1
Delay 250
ATEWritePort 2, 0      ''C7=0
Delay 1000
counter = counter + 1
Next
Delay 100
'' Connecting pin 2 and 4 to CP2 and CP1

ATEWritePort 2, 8    '' C3=1,C4=0
Delay 100
ATEWritePort 1, 133
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 4
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100

'' Trigger 10 times
Dim counter2 As Integer
counter2 = 1
For counter2 = 1 To 10
ATEWritePort 2, 128    ''C7=1
Delay 250
ATEWritePort 2, 0      ''C7=0
Delay 1000
counter2 = counter2 + 1
Next
Delay 100
MsgBox ("Single ended LT2-EN2 Checked")
End Sub

Private Sub Cmd_ine2_Click()
Dim bytValue As Byte

'' Port B and Port C needs to be zero first for this port B and port C needs to be zero and toggle A4,A5,A6,A7

ATEWritePort 1, 0
Delay 100
ATEWritePort 2, 0
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 16    '' Toggles A4
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 32     ''Toggles A5
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 64      ''Toggles A6
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 128     ''Toggles A7
Delay 100

ATEWritePort 1, 0   '' Checking ine on LT2
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 16      ''Toggles A4
Delay 100
WriteMicropulse ("INE 0 0")
Delay 200
ATEWritePort 2, 4        '' C4=1
Delay 100



ATEReadPort 3, bytValue     '' Read the D4 (PORT D )
If bytValue = 16 Then      ''D4 should be 1
MsgBox (" LT2 ine PIn checked, Result=OK ")
Else
MsgBox ("LT2 ine Pin is not Working ")
End If

End Sub

Private Sub Cmd_instsel_Click(Index As Integer)

Dim intLoop As Integer

For intLoop = 0 To 4
    If Cmd_instsel(intLoop).BackColor = &HC0FFC0 Then     'backcolour is green
        Cmd_instsel(intLoop).BackColor = &HC0C0FF        'set the backcolor to red
    End If
Next

Cmd_instsel(Index).BackColor = &HC0FFC0                  'set channel selected to green

Select Case Index                                     'work out the relays we need to turn on
    Case 0                       'sig gen
        ATEWritePort 1, 64
        Delay 100
        ATEWritePort 0, 0
        Delay 100
        ATEWritePort 0, 128
        
    Case 1                        'meter
        ATEWritePort 1, 2
        Delay 100
        ATEWritePort 0, 0
        Delay 100
        ATEWritePort 0, 128
                           
    Case 2                        'scope
        ATEWritePort 1, 0
        Delay 100
        ATEWritePort 0, 0
        Delay 100
        ATEWritePort 0, 128
        
    Case 3                       'oscilloscpe + 50 ohm load
        ATEWritePort 1, 8
        Delay 100
        ATEWritePort 0, 0
        Delay 100
        ATEWritePort 0, 128
    Case 4
        ATEWritePort 1, 72      'Siggen + 50 ohm load
        Delay 100
        ATEWritePort 0, 0
        Delay 100
        ATEWritePort 0, 128
        
        
End Select

End Sub

Private Sub Cmd_ine_Click()

Dim bytValue As Byte

'' Port B and Port C needs to be zero first for this port B and port C needs to be zero and toggle A4,A5,A6,A7

ATEWritePort 1, 0
Delay 100
ATEWritePort 2, 0
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 16    '' Toggles A4
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 32     ''Toggles A5
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 64      ''Toggles A6
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 128     ''Toggles A7
Delay 100

ATEWritePort 1, 8   '' Checking ine on LT1
Delay 100
ATEWritePort 0, 0
Delay 100
ATEWritePort 0, 16      ''Toggles A4
Delay 100
WriteMicropulse ("INE 0 0")
Delay 200
ATEWritePort 2, 4        '' C4=1
Delay 100

ATEReadPort 3, bytValue     '' Read the D4 (PORT D )
If bytValue = 16 Then      ''D4 should be 1
MsgBox (" LT1 ine PIn checked, Result=OK ")
Else
MsgBox ("LT1 ine Pin is not Working ")
End If

End Sub

Private Sub Cmd_reset_Click()
'reset the device
If g_blnATEConnected Then

    If ATEReset = False Then
        MsgBox ("Error Resetting ATE"), vbExclamation
    Else
        MsgBox ("ATE Has Been Reset")
    End If

End If
End Sub

Private Sub ComboLT1_Click()
ATEWritePort 2, 0     ''C7=0
Delay 250
ATEWritePort 2, 128   ''C7=1,C3=0,C4=0
Delay 250
If ComboLT1.Text = "LT1" Then
ATEWritePort 1, 2
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 0
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100
Dim bytValue As Byte
ATEReadPort 3, bytValue     '' Read the D0 (PORT D )
If bytValue = 1 Then      ''D0 should be 1 (Make sure all the other pins of port D are zero, otherwise change the code)
MsgBox (" LT1 5 VOLTS checked, Result=OK ")
Else
MsgBox ("LT1 Pin 1 doesn't have 5 Volts")
End If
End If
If ComboLT1.Text = "LT2" Then
ATEWritePort 1, 0
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 0
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100
ATEReadPort 3, bytValue     '' Read the D0 (PORT D )
If bytValue = 1 Then      ''D0 should be 1 (Make sure all the other pins of port D are zero, otherwise change the code By making port B and C 0 First)
MsgBox (" LT2 5 VOLTS checked, Result=OK ")
Else
MsgBox ("LT2 Pin 1 doesn't have 5 Volts")
End If
End If
End Sub
Private Sub ComboLT2AOFF_Click()
WriteMicropulse ("TERM 0 1")
Delay 100

If ComboLT2AOFF.Text = "Pin 2 (5.0 V)" Then
ATEWritePort 2, 0     ''C7=0
Delay 250
ATEWritePort 2, 154   ''C7=1,C3=1,C4=1,C1=1,C0=0
Delay 250
ATEWritePort 1, 192
Delay 100
ATEWritePort 0, 0          '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 128
Delay 100
ATEWritePort 0, 0          '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100
Dim bytValue As Byte
ATEReadPort 3, bytValue     '' Read the D3 (PORT D)
If bytValue = 8 Then        ''D3 should be 1 (Make sure all the other pins of port D are zero, otherwise change the code)
MsgBox (" LT2 Pin2 checked, Result = 5.0 V ")
Else
MsgBox ("LT2 Pin 2 doesn't have 5.0 Volts")
End If

ElseIf ComboLT2AOFF.Text = "Pin 4 (5.0 V)" Then
ATEWritePort 2, 0     ''C7=0
Delay 250
ATEWritePort 2, 154   ''C7=1,C3=1,C4=1,C1=1,C0=0
Delay 250
ATEWritePort 1, 192
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 4
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100
ATEReadPort 3, bytValue     '' Read the D3 (PORT D )
If bytValue = 8 Then      ''D3 should be 1 (Make sure all the other pins of port D are zero, otherwise change the code By making port B and C 0 First)
MsgBox (" LT2 Pin 4 checked, Result=5.0 V ")
Else
MsgBox ("LT2 Pin 4 doesn't have 5 Volts")
End If

ElseIf ComboLT2AOFF.Text = "Pin 3 (2.5 V)" Then
ATEWritePort 2, 0     ''C7=0
Delay 250
ATEWritePort 2, 153   ''C7=1,C3=1,C4=1,C1=0,C0=1
Delay 250
ATEWritePort 1, 192
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 160
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100
ATEReadPort 3, bytValue     '' Read the D2 (PORT D )
If bytValue = 4 Then      ''D2 should be 1 (Make sure all the other pins of port D are zero, otherwise change the code By making port B and C 0 First)
MsgBox (" LT2 Pin 3 checked, Result=2.5 V")
Else
MsgBox ("LT2 Pin 3 doesn't have 2.5 Volts")
End If

ElseIf ComboLT2AOFF.Text = "Pin 5 (2.5 V)" Then
ATEWritePort 2, 0     ''C7=0
Delay 250
ATEWritePort 2, 153   ''C7=1,C3=1,C4=1,C1=0,C0=1
Delay 250
ATEWritePort 1, 192
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 0
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100
ATEReadPort 3, bytValue     '' Read the D2 (PORT D )
If bytValue = 4 Then      ''D2 should be 1 (Make sure all the other pins of port D are zero, otherwise change the code By making port B and C 0 First)
MsgBox (" LT2 Pin 5 checked, Result=2.5 V ")
Else
MsgBox ("LT2 Pin 5 doesn't have 2.5 Volts")
End If
End If
End Sub

Private Sub ComboLT2AON_Click()
WriteMicropulse ("TERM 0 0")
Delay 100

ATEWritePort 2, 0     ''C7=0
Delay 250
ATEWritePort 2, 152   ''C7=1,C3=1,C4=1,C1=0,C0=0
Delay 250

If ComboLT2AON.Text = "Pin 2 (3.3 V)" Then
ATEWritePort 1, 192
Delay 100
ATEWritePort 0, 0          '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 128
Delay 100
ATEWritePort 0, 0          '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100
Dim bytValue As Byte
ATEReadPort 3, bytValue     '' Read the D1 (PORT D)
If bytValue = 2 Then        ''D1 should be 1 (Make sure all the other pins of port D are zero, otherwise change the code)
MsgBox (" LT2 Pin2 checked, Result=3.3 V ")
Else
MsgBox ("LT2 Pin 2 doesn't have 3.3 Volts")
End If

ElseIf ComboLT2AON.Text = "Pin 3 (3.3 V)" Then
ATEWritePort 1, 192
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 160
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100
ATEReadPort 3, bytValue     '' Read the D1 (PORT D )
If bytValue = 2 Then      ''D1 should be 1 (Make sure all the other pins of port D are zero, otherwise change the code By making port B and C 0 First)
MsgBox (" LT2 Pin 3 checked, Result=3.3 V ")
Else
MsgBox ("LT2 Pin 3 doesn't have 3.3 Volts")
End If

ElseIf ComboLT2AON.Text = "Pin 4 (3.3 V)" Then
ATEWritePort 1, 192
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 4
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100
ATEReadPort 3, bytValue     '' Read the D1 (PORT D )
If bytValue = 2 Then      ''D1 should be 1 (Make sure all the other pins of port D are zero, otherwise change the code By making port B and C 0 First)
MsgBox (" LT2 Pin 4 checked, Result=3.3 V ")
Else
MsgBox ("LT2 Pin 4 doesn't have 3.3 Volts")
End If

ElseIf ComboLT2AON.Text = "Pin 5 (3.3 V)" Then
ATEWritePort 1, 192
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 0
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100
ATEReadPort 3, bytValue     '' Read the D1 (PORT D )
If bytValue = 2 Then      ''D1 should be 1 (Make sure all the other pins of port D are zero, otherwise change the code By making port B and C 0 First)
MsgBox (" LT2 Pin 5 checked, Result=3.3 V ")
Else
MsgBox ("LT2 Pin 5 doesn't have 3.3V Volts")
End If
End If
End Sub



Private Sub ComboLT2BOFF_Click()

WriteMicropulse ("TERM 0 1")
Delay 100

If ComboLT2BOFF.Text = "Pin 2 (5.0 V)" Then
ATEWritePort 2, 0     ''C7=0
Delay 250
ATEWritePort 2, 154   ''C7=1,C3=1,C4=1,C1=1,C0=0
Delay 250
ATEWritePort 1, 128
Delay 100
ATEWritePort 0, 0          '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 128
Delay 100
ATEWritePort 0, 0          '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100
Dim bytValue As Byte
ATEReadPort 3, bytValue     '' Read the D3 (PORT D)
If bytValue = 8 Then        ''D3 should be 1 (Make sure all the other pins of port D are zero, otherwise change the code)
MsgBox (" LT2 Pin2 checked, Result = 5.0 V ")
Else
MsgBox ("LT2 Pin 2 doesn't have 5.0 Volts")
End If

ElseIf ComboLT2BOFF.Text = "Pin 4 (5.0 V)" Then
ATEWritePort 2, 0     ''C7=0
Delay 250
ATEWritePort 2, 154   ''C7=1,C3=1,C4=1,C1=1,C0=0
Delay 250
ATEWritePort 1, 128
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 4
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100
ATEReadPort 3, bytValue     '' Read the D3 (PORT D )
If bytValue = 8 Then      ''D3 should be 1 (Make sure all the other pins of port D are zero, otherwise change the code By making port B and C 0 First)
MsgBox (" LT2 Pin 4 checked, Result=5.0 V ")
Else
MsgBox ("LT2 Pin 4 doesn't have 5 Volts")
End If

ElseIf ComboLT2BOFF.Text = "Pin 3 (2.5 V)" Then
ATEWritePort 2, 0     ''C7=0
Delay 250
ATEWritePort 2, 153   ''C7=1,C3=1,C4=1,C1=0,C0=1
Delay 250
ATEWritePort 1, 128
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 160
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100
ATEReadPort 3, bytValue     '' Read the D2 (PORT D )
If bytValue = 4 Then      ''D2 should be 1 (Make sure all the other pins of port D are zero, otherwise change the code By making port B and C 0 First)
MsgBox (" LT2 Pin 3 checked, Result=2.5 V")
Else
MsgBox ("LT2 Pin 3 doesn't have 2.5 Volts")
End If

ElseIf ComboLT2BOFF.Text = "Pin 5 (2.5 V)" Then
ATEWritePort 2, 0     ''C7=0
Delay 250
ATEWritePort 2, 153   ''C7=1,C3=1,C4=1,C1=0,C0=1
Delay 250
ATEWritePort 1, 128
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 0
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100
ATEReadPort 3, bytValue     '' Read the D2 (PORT D )
If bytValue = 4 Then      ''D2 should be 1 (Make sure all the other pins of port D are zero, otherwise change the code By making port B and C 0 First)
MsgBox (" LT2 Pin 5 checked, Result=2.5 V ")
Else
MsgBox ("LT2 Pin 5 doesn't have 2.5 Volts")
End If
End If

End Sub

Private Sub ComboLT2BON_Click()
WriteMicropulse ("TERM 0 0")
Delay 100

ATEWritePort 2, 0     ''C7=0
Delay 250
ATEWritePort 2, 152   ''C7=1,C3=1,C4=1,C1=0,C0=0
Delay 250

If ComboLT2BON.Text = "Pin 2 (3.3 V)" Then
ATEWritePort 1, 128
Delay 100
ATEWritePort 0, 0          '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 128
Delay 100
ATEWritePort 0, 0          '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100
Dim bytValue As Byte
ATEReadPort 3, bytValue     '' Read the D1 (PORT D)
If bytValue = 2 Then        ''D1 should be 1 (Make sure all the other pins of port D are zero, otherwise change the code)
MsgBox (" LT2 Pin2 checked, Result=3.3 V ")
Else
MsgBox ("LT2 Pin 2 doesn't have 3.3 Volts")
End If

ElseIf ComboLT2BON.Text = "Pin 3 (3.3 V)" Then
ATEWritePort 1, 128
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 160
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100
ATEReadPort 3, bytValue     '' Read the D1 (PORT D )
If bytValue = 2 Then      ''D1 should be 1 (Make sure all the other pins of port D are zero, otherwise change the code By making port B and C 0 First)
MsgBox (" LT2 Pin 3 checked, Result=3.3 V ")
Else
MsgBox ("LT2 Pin 3 doesn't have 3.3 Volts")
End If

ElseIf ComboLT2BON.Text = "Pin 4 (3.3 V)" Then
ATEWritePort 1, 128
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 4
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100
ATEReadPort 3, bytValue     '' Read the D1 (PORT D )
If bytValue = 2 Then      ''D1 should be 1 (Make sure all the other pins of port D are zero, otherwise change the code By making port B and C 0 First)
MsgBox (" LT2 Pin 4 checked, Result=3.3 V ")
Else
MsgBox ("LT2 Pin 4 doesn't have 3.3 Volts")
End If

ElseIf ComboLT2BON.Text = "Pin 5 (3.3 V)" Then
ATEWritePort 1, 128
Delay 100
ATEWritePort 0, 0    '' Toggle A4 FROM 0 TO 1
Delay 100
ATEWritePort 0, 16
Delay 100
ATEWritePort 1, 0
Delay 100
ATEWritePort 0, 0    '' Toggle A5 FROM 0 TO 1
Delay 100
ATEWritePort 0, 32
Delay 100
ATEReadPort 3, bytValue     '' Read the D1 (PORT D )
If bytValue = 2 Then      ''D1 should be 1 (Make sure all the other pins of port D are zero, otherwise change the code By making port B and C 0 First)
MsgBox (" LT2 Pin 5 checked, Result=3.3 V ")
Else
MsgBox ("LT2 Pin 5 doesn't have 3.3V Volts")
End If
End If
End Sub

Private Sub Form_Load()
ComboLT1.AddItem "LT1"
ComboLT1.AddItem "LT2"
ComboLT2AON.AddItem "Pin 2 (3.3 V)"
ComboLT2AON.AddItem "Pin 3 (3.3 V)"
ComboLT2AON.AddItem "Pin 4 (3.3 V)"
ComboLT2AON.AddItem "Pin 5 (3.3 V)"
ComboLT2AOFF.AddItem "Pin 2 (5.0 V)"
ComboLT2AOFF.AddItem "Pin 4 (5.0 V)"
ComboLT2AOFF.AddItem "Pin 3 (2.5 V)"
ComboLT2AOFF.AddItem "Pin 5 (2.5 V)"
ComboLT2BON.AddItem "Pin 2 (3.3 V)"
ComboLT2BON.AddItem "Pin 3 (3.3 V)"
ComboLT2BON.AddItem "Pin 4 (3.3 V)"
ComboLT2BON.AddItem "Pin 5 (3.3 V)"
ComboLT2BOFF.AddItem "Pin 2 (5.0 V)"
ComboLT2BOFF.AddItem "Pin 4 (5.0 V)"
ComboLT2BOFF.AddItem "Pin 3 (2.5 V)"
ComboLT2BOFF.AddItem "Pin 5 (2.5 V)"
End Sub
