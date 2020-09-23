VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "ColorThang"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   3135
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Give Color"
      Height          =   195
      Left            =   1815
      TabIndex        =   13
      Top             =   3795
      Width           =   1170
   End
   Begin VB.TextBox txtHex2 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4305
      Width           =   2925
   End
   Begin VB.TextBox txtHex 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4035
      Width           =   2925
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Give Color"
      Height          =   195
      Left            =   1815
      TabIndex        =   14
      Top             =   2970
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Give Color"
      Height          =   195
      Left            =   1815
      TabIndex        =   15
      Top             =   2160
      Width           =   1185
   End
   Begin VB.TextBox txtB1 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   990
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4755
      Width           =   400
   End
   Begin VB.TextBox txtG2 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2205
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4755
      Width           =   400
   End
   Begin VB.TextBox txtG1 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   525
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4755
      Width           =   400
   End
   Begin VB.TextBox txtB2 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2670
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   4755
      Width           =   400
   End
   Begin VB.TextBox txtR2 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1725
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4755
      Width           =   400
   End
   Begin VB.TextBox txtR1 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4755
      Width           =   400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   210
      Left            =   135
      TabIndex        =   16
      Top             =   30
      Width           =   720
   End
   Begin VB.TextBox txtLong2 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3495
      Width           =   2925
   End
   Begin VB.TextBox txtLong 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3225
      Width           =   2925
   End
   Begin VB.PictureBox picColor2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   390
      ScaleHeight     =   675
      ScaleWidth      =   810
      TabIndex        =   24
      Top             =   1410
      Width           =   840
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   390
      ScaleHeight     =   675
      ScaleWidth      =   810
      TabIndex        =   23
      Top             =   465
      Width           =   840
   End
   Begin VB.TextBox txtColor2 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2670
      Width           =   2925
   End
   Begin VB.TextBox txtColor 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2400
      Width           =   2925
   End
   Begin VB.PictureBox picChooseColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   2115
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   1665
      ScaleWidth      =   840
      TabIndex        =   0
      Top             =   375
      Width           =   870
   End
   Begin VB.Label Label15 
      Caption         =   "R"
      Height          =   195
      Left            =   0
      TabIndex        =   33
      Top             =   4275
      Width           =   120
   End
   Begin VB.Label Label14 
      Caption         =   "L"
      Height          =   195
      Left            =   0
      TabIndex        =   32
      Top             =   4020
      Width           =   120
   End
   Begin VB.Label Label13 
      Caption         =   "Hex Color :"
      Height          =   225
      Left            =   120
      TabIndex        =   31
      Top             =   3825
      Width           =   1170
   End
   Begin VB.Label Label12 
      Caption         =   "/"
      Height          =   225
      Left            =   1515
      TabIndex        =   30
      Top             =   4755
      Width           =   105
   End
   Begin VB.Label Label11 
      Caption         =   "Right R G B :"
      Height          =   240
      Left            =   1680
      TabIndex        =   29
      Top             =   4560
      Width           =   1260
   End
   Begin VB.Label Label10 
      Caption         =   "Left R G B :"
      Height          =   195
      Left            =   15
      TabIndex        =   28
      Top             =   4560
      Width           =   1350
   End
   Begin VB.Label Label9 
      Caption         =   "R"
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   3495
      Width           =   105
   End
   Begin VB.Label Label8 
      Caption         =   "L"
      Height          =   210
      Left            =   0
      TabIndex        =   26
      Top             =   3255
      Width           =   90
   End
   Begin VB.Label Label7 
      Caption         =   "Long Color :"
      Height          =   240
      Left            =   120
      TabIndex        =   25
      Top             =   3000
      Width           =   1125
   End
   Begin VB.Label Label6 
      Caption         =   "R"
      Height          =   240
      Left            =   0
      TabIndex        =   22
      Top             =   2670
      Width           =   120
   End
   Begin VB.Label Label5 
      Caption         =   "L"
      Height          =   225
      Left            =   0
      TabIndex        =   21
      Top             =   2415
      Width           =   90
   End
   Begin VB.Label Label4 
      Caption         =   "HTML Color :"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   2205
      Width           =   1125
   End
   Begin VB.Label Label3 
      Caption         =   "Right or Left Click here :"
      Height          =   270
      Left            =   990
      TabIndex        =   19
      Top             =   30
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Right Color :"
      Height          =   240
      Left            =   30
      TabIndex        =   18
      Top             =   1185
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Left Color :"
      Height          =   255
      Left            =   45
      TabIndex        =   17
      Top             =   255
      Width           =   990
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Project by : Fatty
'
' I made this out of Curiosity. I was bored, and wanted to figure out how to do this
' then I added everything else (HTML, Longs, and RGB). I also added that stuff because
' when I was making this PlanetSourceCode was having small problems, and I couldn't
' upload so I had nothing to do but Improve this code. This took me SO LONG to figure
' out it's not even funny. Since the Hex colors go BGR, instead of RGB. So to convert
' HTML Color codes back to the long value, you have to move the string all around.
' Which I would have never figured out if not for my Visual Basic 6 : From the
' Ground up book.

Private Sub Command1_Click()
 frmAbout.Show ' Obviously... Show the about form.
End Sub

Private Sub Command2_Click()
On Error Resume Next
 If txtColor.Text <> "" Then
  picColor.BackColor = LongfromHTML(txtColor.Text)
 End If
 If txtColor2.Text <> "" Then
  picColor2.BackColor = LongfromHTML(txtColor2.Text)
 End If
End Sub

Private Sub Command3_Click()
 On Error Resume Next
 If txtLong.Text <> "" Then
  picColor.BackColor = txtLong.Text
 End If
 If txtLong2.Text <> "" Then
  picColor2.BackColor = txtLong2.Text
 End If
End Sub

Private Sub Command4_Click()
 If Len(txtHex.Text) <> 0 Then
  If Left(txtHex.Text, 2) = "&H" Then
   crap$ = Right(txtHex.Text, Len(txtHex.Text) - 2)
   If Right(crap$, 1) = "&" Then
   crap$ = Left(crap$, Len(crap$) - 1)
   End If
  End If
  poop$ = HexToRGB(crap$)
  picColor.BackColor = poop$
  MsgBox poop$
 End If
 If Len(txtHex2.Text) <> 0 Then
  If Left(txtHex2.Text, 2) = "&H" Then
   crap2$ = Right(txtHex2.Text, Len(txtHex2.Text) - 2)
   If Right(crap2$, 1) = "&" Then
   crap2$ = Left(crap2$, Len(crap2$) - 1)
   End If
  End If
  poop2$ = HexToRGB(crap2$)
  picColor2.BackColor = poop2$
  MsgBox poop2$
 End If
End Sub

Private Sub picChooseColor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next ' Just to be safe
 If Button = vbLeftButton Then ' Check which button was pressed
  picColor.BackColor = picChooseColor.Point(X, Y) ' If it was the Left button then make the
' Left button picture that color. What this does is get the color from
' the EXACT Pixel that your mouse is on. This would be a good thing
' for a Paint program.
 ElseIf Button = vbRightButton Then ' Check the button
  picColor2.BackColor = picChooseColor.Point(X, Y) ' Same as above get color, blah blah
 End If ' Can't forget to end the if...
End Sub

Private Sub picChooseColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' I put this one to update the color in the Picture Box when you move the mouse.
On Error Resume Next ' Just to be safe
 
 If Button = vbLeftButton Then ' Check the button pressed
  picColor.BackColor = picChooseColor.Point(X, Y) ' Get the color of Pixel.
 ElseIf Button = vbRightButton Then 'Check button.
  picColor2.BackColor = picChooseColor.Point(X, Y) ' Get the color of Pixel.
 End If ' End the if
End Sub

Private Sub picChooseColor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' This just sets the TextBoxs to the color according to what button was pressed.
On Error Resume Next

 If Button = vbLeftButton Then
  txtColor.Text = HTMLfromLong(picColor.BackColor)
  txtLong.Text = LongfromHTML(txtColor.Text)
  Call RGBOut(txtLong.Text, txtR1, txtG1, txtB1)
  Tmp = HTMLfromLong(picColor.BackColor)
  Tmp = Replace(Tmp, "#", "")
  s1 = Left(Tmp, 2)
  s2 = Mid(Tmp, 3, 2)
  s3 = Right(Tmp, 2)
  Tmp = s3 & s2 & s1
  txtHex.Text = "&H" & Tmp & "&"
 ElseIf Button = vbRightButton Then
  txtColor2.Text = HTMLfromLong(picColor2.BackColor)
  txtLong2.Text = LongfromHTML(txtColor2.Text)
  Call RGBOut(txtLong2.Text, txtR2, txtG2, txtB2)
  Tmp = HTMLfromLong(picColor2.BackColor)
  Tmp = Replace(Tmp, "#", "")
  s1 = Left(Tmp, 2)
  s2 = Mid(Tmp, 3, 2)
  s3 = Right(Tmp, 2)
  Tmp = s3 & s2 & s1
  txtHex2.Text = "&H" & Tmp & "&"
 End If
End Sub

Public Function HTMLfromLong(lngColor As Long) As String
Dim Red As String, Green As String, Blue As String

On Error GoTo ExitThis 'Cautions...
 Red$ = Hex(lngColor And 255) ' Makes Red$ = the hex of lngColor
 Green$ = Hex(lngColor \ 256 And 255) ' This one was a little more complicated.
 Blue$ = Hex(lngColor \ 65536 And 255) ' Took me a while for this one.
 If Len(Red$) < 2 Then Red$ = "0" & Red
 If Len(Green$) < 2 Then Green$ = "0" & Green
 If Len(Blue$) < 2 Then Blue$ = "0" & Blue
 HTMLfromLong = "#" & Red$ & Green$ & Blue$ ' Set this function to equal the HTML Color.

ExitThis:
 Exit Function
End Function

Public Function LongfromHTML(HTMLColor As String)
Dim s1$, s2$, s3$
'Okay here, you have to reverse stuff around. Reason being is,
'Hex color is BBGGRR (Blue, Green, Red). So I have to take the first two
'(Red) and move it to the end, then move the last two (blue) to the front.

 HTMLColor$ = Replace(HTMLColor$, "#", "")
 s1$ = Left$(HTMLColor$, 2)
 s2$ = Mid$(HTMLColor$, 3, 2)
 s3$ = Right$(HTMLColor$, 2)
 LongfromHTML = HexToRGB(s3$ & s2$ & s1$)
End Function

Private Function HexToRGB(HexCode As String) As Currency
'This baby Converts the Hex code, (which was reversed... see above...) to
'it's Long Value.
On Error GoTo error
 Dim Tmp$
 Dim Nums1 As Integer, Nums2 As Integer
 Dim Nums3 As Long, Nums4 As Long
 Const Hx = "&H"
 Const Big = 65536
 Const Lil = 256, Two = 2
  Tmp = HexCode
  If UCase(Left$(HexCode, 2)) = "&H" Then Tmp = Mid$(HexCode, 3)
  Tmp = Right$("0000000" & Tmp, 8)
  If IsNumeric(Hx & Tmp) Then
   Nums1 = CInt(Hx & Right$(Tmp, Two))
   Nums3 = CLng(Hx & Mid$(Tmp, 5, Two))
   Nums2 = CInt(Hx & Mid$(Tmp, 3, Two))
   Nums4 = CLng(Hx & Left$(Tmp, Two))
   HexToRGB = CCur(Nums4 * Lil + Nums2) * Big + (Nums3 * Lil) + Nums1
  End If
  Exit Function
error:
 MsgBox Err.Description, vbExclamation, "Error"
End Function

Public Sub RGBOut(lngColorVal As String, ROut As TextBox, GOut As TextBox, BOut As TextBox)
' This is the same as the HTMLtoLong code, cept it isn't getting the Hex value.
 ROut.Text = lngColorVal And 255
 GOut.Text = (lngColorVal \ 256 And 255)
 BOut.Text = (lngColorVal \ 65536 And 255)
End Sub

' Below is SUPER usless, but I felt like doing it.

Private Sub txtColor_KeyPress(KeyAscii As Integer)

 Select Case KeyAscii% ' Make a Select Case of KeyAscii
 
  Case 8 ' Backspace
   If Len(txtColor.Text) = 0 Then Exit Sub
   poop = Left(txtColor.Text, Len(txtColor.Text) - 1)
   txtColor.Text = ""
   txtColor.SelText = poop
  
  Case 35 ' The # Sign
  If Len(txtColor.Text) = 7 Then Exit Sub
   If Left(txtColor.Text, 1) = "#" Then
    Exit Sub
   End If
   txtColor.SelText = Chr$(KeyAscii)
  
  Case 48 To 57 ' 0-9
  If Len(txtColor.Text) = 7 Then Exit Sub
   txtColor.SelText = Chr$(KeyAscii) ' cursor would stay in front of everything.

  Case 66 To 70 ' a-f
  If Len(txtColor.Text) = 7 Then Exit Sub
   txtColor.SelText = UCase(Chr$(KeyAscii))
   
  Case 97 To 102 ' A-F
  If Len(txtColor.Text) = 7 Then Exit Sub
   txtColor.SelText = UCase(Chr$(KeyAscii))
   
 End Select ' End the Select Case
End Sub

Private Sub txtColor2_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii% ' Make a Select Case of KeyAscii
 
  Case 8
   If Len(txtColor2.Text) = 0 Then Exit Sub ' Backspace
   poop = Left(txtColor2.Text, Len(txtColor2.Text) - 1)
   txtColor2.Text = ""
   txtColor2.SelText = poop
  
  Case 35 ' The # Sign
  If Len(txtColor2.Text) = 7 Then Exit Sub
  If Left(txtColor2.Text, 1) = "#" Then
    Exit Sub
   End If
   txtColor2.SelText = Chr$(KeyAscii)
  
  Case 48 To 57 ' 0-9
  If Len(txtColor2.Text) = 7 Then Exit Sub
   txtColor2.SelText = Chr$(KeyAscii) ' cursor would stay in front of everything.

  Case 66 To 70 ' a-f
  If Len(txtColor2.Text) = 7 Then Exit Sub
   txtColor2.SelText = UCase(Chr$(KeyAscii))
   
  Case 97 To 102 ' A-F
  If Len(txtColor2.Text) = 7 Then Exit Sub
   txtColor2.SelText = UCase(Chr$(KeyAscii))
   
 End Select ' End the Select Case
End Sub

Private Sub txtHex2_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii% ' Make a Select Case of KeyAscii
  Case 8
   poop = Left(txtHex2.Text, Len(txtHex2.Text) - 1)
   txtHex2.Text = ""
   txtHex2.SelText = poop
  Case 38
   If Left(txtHex2.Text, 2) = "&H" Then
    If Right(txtHex2.Text, 1) = "&" Then
     Exit Sub
    End If
   End If
   If txtHex2.Text = "" Then
    txtHex2.SelText = "&H"
    Exit Sub
   End If
   txtHex2.SelText = Chr$(KeyAscii)
  Case 48 To 57 ' if you don't know, 48-57 are 0-9 Keys
   txtHex2.SelText = Chr$(KeyAscii) ' Add the number, I used SelText cause the blinking
                                  ' cursor would stay in front of everything.
  Case 66 To 70 ' a-f
   txtHex2.SelText = UCase(Chr$(KeyAscii))
   
  Case 97 To 102 ' A-F
   txtHex2.SelText = UCase(Chr$(KeyAscii))
  Case 104
   txtHex2.SelText = UCase(Chr$(KeyAscii))
 End Select ' End the Select Case
End Sub

Private Sub txtLong_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii% ' Make a Select Case of KeyAscii
  Case 8
   If txtLong.SelLength > 0 Then
    crack = Mid$(txtLong.Text, txtLong.SelStart + 1, txtLong.SelLength)
    sex = txtLong.Text
    txtLong.Text = ""
    txtLong.SelText = Replace(sex, crack, "")
    Exit Sub
   End If
   poop = Left(txtLong.Text, Len(txtLong.Text) - 1)
   txtLong.Text = ""
   txtLong.SelText = poop
  Case 48 To 57 ' if you don't know, 48-57 are 0-9 Keys
   txtLong.SelText = Chr$(KeyAscii) ' Add the number, I used SelText cause the blinking
                                  ' cursor would stay in front of everything.
 End Select ' End the Select Case
End Sub

Private Sub txtLong2_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii% ' Make a Select Case of KeyAscii
  Case 8
   If txtLong.SelLength > 0 Then
    crack = Mid$(txtLong.Text, txtLong.SelStart + 1, txtLong.SelLength)
    sex = txtLong.Text
    txtLong.Text = ""
    txtLong.SelText = Replace(sex, crack, "")
    Exit Sub
   End If
   poop = Left(txtLong.Text, Len(txtLong.Text) - 1)
   txtLong.Text = ""
   txtLong.SelText = poop
  Case 48 To 57 ' if you don't know, 48-57 are 0-9 Keys
   txtLong2.SelText = Chr$(KeyAscii) ' Add the number, I used SelText cause the blinking
                                  ' cursor would stay in front of everything.
 End Select ' End the Select Case
End Sub

Private Sub txtHex_KeyPress(KeyAscii As Integer)
On Error Resume Next
 Select Case KeyAscii% ' Make a Select Case of KeyAscii
  Case 8
   poop = Left(txtHex.Text, Len(txtHex.Text) - 1)
   txtHex.Text = ""
   txtHex.SelText = poop
  Case 38
   If Left(txtHex.Text, 2) = "&H" Then
    If Right(txtHex.Text, 1) = "&" Then
     Exit Sub
    End If
   End If
   If txtHex.Text = "" Then
    txtHex.SelText = "&H"
    Exit Sub
   End If
   txtHex.SelText = Chr$(KeyAscii)
  Case 48 To 57 ' if you don't know, 48-57 are 0-9 Keys
   txtHex.SelText = Chr$(KeyAscii) ' Add the number, I used SelText cause the blinking
                                  ' cursor would stay in front of everything.
  Case 66 To 70 ' a-f
   txtHex.SelText = UCase(Chr$(KeyAscii))
   
  Case 97 To 102 ' A-F
   txtHex.SelText = UCase(Chr$(KeyAscii))
  Case 104
   txtHex.SelText = UCase(Chr$(KeyAscii))
 End Select ' End the Select Case
End Sub
