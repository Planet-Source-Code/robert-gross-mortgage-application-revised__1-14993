VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7020
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   2055
      ScaleWidth      =   2775
      TabIndex        =   2
      ToolTipText     =   "Your Dream Home!"
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5400
      Top             =   480
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Copperplate Gothic Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   464
      TabIndex        =   0
      Top             =   0
      Width           =   7020
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mortgage Calculator"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   3120
      TabIndex        =   3
      Top             =   960
      Width           =   3615
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Code: Fading Text
' Author: Jonathan Roach
' Email: stormdev@golden.net
'
' Comments:
' Enjoy the code, votes and emails are always welcomed.
' Happy Coding !
'
' Update Notes,
'
' I have added the ability to having multiple messages
' displayed and rotated through the fade in and out
' routines. The new variable curMsg and MsgArray have been
' added to accomplish this, you can increase the size of the
' array or decrease it depending on how many messages you
' want to cycle through. Enjoy the code !
'
'Code revised by: R.Gross on 12-4-2000
'
Dim colVal As Long          ' Variable to contain our rgb color value
Dim counter As Integer      ' Counter to be used to increment/decrement rgb value
Dim fIn As Boolean          ' Boolean, true/false variable to determine if we are fading in or fading out.
' *** Update ***
Dim curMsg As Integer           ' Determines which message from the MsgArray to display
Dim MsgArray(1 To 9) As String  ' New variable Array to allow us to display multiple messages

Private Sub cmdClose_Click()
' Shutdown all components and end the program
Timer1.Enabled = False
Unload frmAbout

End Sub

Private Sub Form_Load()
' Setup our messages
MsgArray(1) = "Mortgage Calculator"
MsgArray(2) = "Find out what your new house will cost!"
MsgArray(3) = "Put down as large a down payment as you can afford"
MsgArray(4) = "Check out the neighborhood for best re-sale value"
MsgArray(5) = "Shop around for a loan"
' Use a blank entry for a greater pause before the next message !
MsgArray(6) = ""
MsgArray(7) = "Thanks for using this program"
MsgArray(8) = "Comments are always welcome"
MsgArray(9) = "RGross_41@msn.com"
' Set the current message
curMsg = 1
' Set our counter to zero
counter = 0
' This is our boolean (true/false) variable which will be used
' to determin if we are fading the text in or out. To start we
' set it to true so our text will fade in first.
fIn = True
' Show the form
Me.Show
cmdClose.SetFocus
End Sub

Private Sub Timer1_Timer()
' This is the timer that controls our fading in and out, if you
' increase the timers interval property, the fades will take longer,
' you can also alter the fade speed by increasing or decreasing the
' the value that is added to counter in the lines below.
'
' Check and see if we are fading the text in, fIn=true.
If fIn = True Then
' If we are fading in then begin the fade routine.
' Check if our counter variable is less than or equal to our goal color,
' I used 170 because that is the rgb value for the red component that
' I wanted the color fade to stop at, but you can use any rgb value and
' the red green or blue components.
'
    If counter <= 200 Then
        ' Store our rgb color value into the colVal variable
        ' Notice the counter variable as the red parameter in
        ' the rgb call, the zeros for the green and blue
        ' components will yield black, you can try putting
        ' the counter variable in different components and
        ' changing component values to get the color you
        ' desire.
        colVal = RGB(counter, 0, 0)
        ' Set the forecolor of our picturebox to the new color
        Picture1.ForeColor = colVal
        ' Increment our counter variable, a higher + value
        ' will increase fade speed while a lower value will
        ' slow the fade down and make it more smooth.
        counter = counter + 10
    Else
        ' If we have reached our final color then set the fIn variable to false
        ' because we just faded the text in we need it false to fade out next.
        fIn = False
        ' Reset the counter
        counter = 0
    End If
Else
' If fIn was false that means we are fading out now
' Basically the same code as above, so I won't comment heavily.
    If counter <= 200 Then
        colVal = RGB(200 - counter, 0, 0)
        Picture1.ForeColor = colVal
        counter = counter + 4
    Else
        fIn = True
        counter = 0
        ' Increment our Message Pointer
        curMsg = curMsg + 1
    End If
End If
' Check if the curMsg is at the last ("Upper Boundary"UBound) position of our Array.
If curMsg <= UBound(MsgArray) Then
    ' Print the current message
    SetTextPosition
    Picture1.Print MsgArray(curMsg)
Else
    ' Reset our message pointer to the first message to loop through the MsgArray
    curMsg = 1
End If
End Sub

Private Sub SetTextPosition()
' Set our current x and y coordinates for writing to the picturebox.
' This line sets the CurrentX property based on the size of the msg
' and centers that line in the textbox horizontally.
Picture1.CurrentX = (Picture1.ScaleWidth - Picture1.TextWidth(MsgArray(curMsg))) \ 2
' You could also use the above routine to center the text vertically, by replacing
' the .ScaleWidth with .ScaleHeight and .TextWidth with .TextHeight, for the CurrentY
' Co-ordinate.
' ie: Picture1.CurrentY = (Picture1.ScaleHeight - Picture1.TextHeight(MsgArray(curMsg))) \ 2
'
' For Now I will just hard code it as 1
Picture1.CurrentY = 1
End Sub
