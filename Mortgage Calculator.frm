VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Mortgage Calculator"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8835
   ForeColor       =   &H00C0C0C0&
   Icon            =   "Mortgage Calculator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Mortgage Calculator.frx":0442
   ScaleHeight     =   4005
   ScaleWidth      =   8835
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picHome 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2835
      Left            =   0
      Picture         =   "Mortgage Calculator.frx":4B588
      ScaleHeight     =   2805
      ScaleWidth      =   3750
      TabIndex        =   0
      Top             =   0
      Width           =   3780
   End
   Begin VB.CommandButton cmdAmortization 
      Caption         =   "Amor&tization"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   10
      Top             =   3240
      Width           =   1215
   End
   Begin MSMask.MaskEdBox mskRate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   5
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      ToolTipText     =   "Enter interest rate like this: for 5.25% enter 5.25"
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#0.0##%"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskLoanAmount 
      CausesValidation=   0   'False
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      ToolTipText     =   "Enter amount to borrow"
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      MaxLength       =   12
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$#,##0.00;($#,##0.00)"
      Mask            =   "#,###,###.##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtPayment 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtTerm 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      ToolTipText     =   "Length of Loan in Years"
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   9
      ToolTipText     =   "Calculate Payment Amount"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Monthly Payment:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4080
      TabIndex        =   7
      Top             =   2115
      Width           =   2370
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Term:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5685
      TabIndex        =   5
      Top             =   1515
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Interest rate:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4710
      TabIndex        =   3
      Top             =   915
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "&Amount:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5355
      TabIndex        =   1
      Top             =   315
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Sub cmdAmortization_Click()
'Dim LoanAmount As Double    'Amount to borrow

Dim objExcel As Excel.Application
Dim objWorkBook As Excel.Workbook
Dim objWorkSheet As Excel.Worksheet

    On Error GoTo OLE_ERROR
'Create Excel object
    Set objExcel = GetObject("", "Excel.Application")
    Set objWorkBook = objExcel.Workbooks.Open(App.Path & "\" & "Amortiz.xls")
    objExcel.Visible = True
    frmMain.Visible = False

'Place the data entered into Excel Amortiz spreadsheet
    Set objWorkSheet = ActiveSheet
    With objWorkSheet
    .Cells(14, 3) = Format(mskLoanAmount.Text, "Currency")
    .Cells(15, 3) = mskRate.Text
    .Cells(16, 3) = txtTerm.Text
    End With

Set objExcel = Nothing
Set objWorkBook = Nothing
Set objWorkSheet = Nothing

EndApplication:
    frmMain.Visible = True
    Exit Sub
'
'Error handler
'
OLE_ERROR:
   MsgBox Error$(Err)
   If Not (objExcel Is Nothing) Then
      '
      'Close Excel and release the Object Variable
      Set objExcel = Nothing
      Set objWorkBook = Nothing
      Set objWorkSheet = Nothing
   End If
   Resume EndApplication

End Sub

Private Sub cmdAmortization_LostFocus()
'Clear data out for new entry
mskRate.Text = ""
txtTerm.Text = ""
txtPayment.Text = ""

End Sub

Private Sub cmdCalculate_Click()
'Test for data present in all text boxes

Dim LoanAmount As Double    'Amount to borrow
Dim Rate As Double          'Interest Rate of Loan
Dim Nper As Integer         'Total payment periods
Dim Payment As Double       'Monthly Payment

On Error GoTo ErrHandler

If mskLoanAmount.Text <= 0 Or _
    mskLoanAmount.Text = "_,___,___.__" Or _
    mskRate.Text <= 0 Or txtTerm.Text <= 0 _
    Then GoTo ErrHandler

LoanAmount = mskLoanAmount.Text
Rate = mskRate.Text / 12
Nper = txtTerm.Text * 12

Payment = Pmt(Rate, Nper, -LoanAmount)
txtPayment.Text = Format(Payment, "Currency", 2)

Exit Sub

ErrHandler:
MsgBox "Please enter valid data into fields.", vbExclamation

End Sub



Private Sub mnuFileExit_Click()
  Unload Me

End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show
End Sub

' Select all the text in this TextBox.
Private Sub SelectField(ByVal text_box As TextBox)
    text_box.SelStart = 0
    text_box.SelLength = Len(text_box.Text)
End Sub

' Select all the text in this TextBox.
Private Sub SelectmField(ByVal text_box As MaskEdBox)
    text_box.SelStart = 0
    text_box.SelLength = Len(text_box.Text)
End Sub

Private Sub mskLoanAmount_Change()

'Clear all text box values when this data changes
mskRate.Text = ""
txtTerm.Text = ""
txtPayment.Text = ""

End Sub

Private Sub mskLoanAmount_GotFocus()
    SelectmField mskLoanAmount
    
End Sub

Private Sub mskLoanAmount_LostFocus()

If mskLoanAmount.Text = "_,___,___.__" Then
    mskLoanAmount.SetFocus
    MsgBox "Please enter valid data into fields.", vbExclamation
End If

End Sub

Private Sub mskRate_LostFocus()
'Format the Rate to be as a percentage but allow user to input
'a whole number; for example: 7.5% entered as 7.5 and not as
'0.075.

On Error GoTo ErrHandler

If mskRate.Text <= 0 Then GoTo ErrHandler

mskRate.Text = mskRate.Text / 100
Exit Sub

ErrHandler:
mskRate.SetFocus
MsgBox "Please enter valid data into field.", vbExclamation

End Sub

Private Sub txtPayment_Change()
txtPayment.Text = Format(txtPayment.Text, "Currency", 2)

End Sub

Private Sub mskRate_GotFocus()
    SelectmField mskRate

End Sub

Private Sub txtTerm_GotFocus()
    SelectField txtTerm

End Sub

Private Sub txtTerm_LostFocus()
'Test to insure that data has been entered in text box
If txtTerm.Text = "" Then
    txtTerm.SetFocus
    MsgBox "Please enter valid data into fields.", vbExclamation
End If

End Sub
