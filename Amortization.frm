VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Amortization 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Amortization Table"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Amortization.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4895
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColor       =   16777215
      GridLines       =   3
      GridLinesFixed  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Amortization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
'The FlexGrid's column headers...
    With MSFlexGrid1
        .TextMatrix(0, 0) = "No."
        .TextMatrix(0, 1) = "Begin Bal."
        .TextMatrix(0, 2) = "Interest"
        .TextMatrix(0, 3) = "Principal"
        .TextMatrix(0, 4) = "End Bal."
        .TextMatrix(0, 5) = "Cum. Interest"
        .TextMatrix(0, 6) = "Total Paid"
    End With
'Set Column Headers to bold font, change font color,
'size, and back color. Also change row height
    With MSFlexGrid1
        .Row = 0
        .Col = 0
        .CellFontBold = True
        .CellFontName = "Garamond"
        .CellFontSize = 14
        .CellForeColor = &HC00000
        .CellBackColor = &HFFFFFF
        .Row = 0
        .Col = 1
        .CellFontBold = True
        .CellFontName = "Garamond"
        .CellFontSize = 14
        .CellForeColor = &HC00000
        .CellBackColor = &HFFFFFF
        .Row = 0
        .Col = 2
        .CellFontBold = True
        .CellFontName = "Garamond"
        .CellFontSize = 14
        .CellForeColor = &HC00000
        .CellBackColor = &HFFFFFF
        .Row = 0
        .Col = 3
        .CellFontBold = True
        .CellFontName = "Garamond"
        .CellFontSize = 14
        .CellForeColor = &HC00000
        .CellBackColor = &HFFFFFF
        .Row = 0
        .Col = 4
        .CellFontBold = True
        .CellFontName = "Garamond"
        .CellFontSize = 14
        .CellForeColor = &HC00000
        .CellBackColor = &HFFFFFF
        .Row = 0
        .Col = 5
        .CellFontBold = True
        .CellFontName = "Garamond"
        .CellFontSize = 14
        .CellForeColor = &HC00000
        .CellBackColor = &HFFFFFF
        .Row = 0
        .Col = 6
        .CellFontBold = True
        .CellFontName = "Garamond"
        .CellFontSize = 14
        .CellForeColor = &HC00000
        .CellBackColor = &HFFFFFF
        
        .RowHeight(0) = 350
    End With
    
'Set the Rate back so calculations will be correct.
'Rate was manipulated so as to keep the user from having to
'enter the Rate as a decimal value, e.g., 0.05 for 5%
Dim Rate As Double          'Interest Rate of Loan
frmMain.mskRate.Text = frmMain.mskRate.Text * 100
Rate = frmMain.mskRate.Text / 1200

Dim iPayments As Integer    'Equals number of monthly payments
iPayments = Val(frmMain.txtTerm.Text) * 12

'**********************************************************
'Beginning of the Amortization Calculations copied from   *
'GraphAMort by Eoin Armstrong on the Planet Source Code site. *
'**********************************************************
'This sub populates the MSFlexgrid

    Dim iCountPay As Integer    'Use to Loop
    Dim dCurrBal As Currency      'Starting loan balance
    Dim dCurrInt As Currency      'Interest this period
    Dim dCurrPrin As Currency     'Principal paid this period
    Dim dEndBal As Currency       'Loan balance at end of period
    Dim dCumInt As Currency       'Cumulative interest at end of period
    Dim dTotPaid As Currency      'Total paid to date at end of period

        dTotPaid = frmMain.txtPayment.Text                 'assign monthly payment
        dCurrBal = frmMain.mskLoanAmount.Text              'assign loan balance for 1st period
        dCurrInt = dCurrBal / 12 * (frmMain.mskRate / 100) 'assign interest for the first period
        dCurrPrin = dTotPaid - dCurrInt                    'assign principal paid for the first period
        dEndBal = dCurrBal - dCurrPrin                     'assign balance at end of the first period
        dCumInt = dCurrInt                                 'assign cumulative interest for the first period
        dTotPaid = frmMain.txtPayment.Text                 'assign running total paid for first period
    
    With MSFlexGrid1
        .Rows = iPayments + 1   'number of rows in the
                                'FlexGrid + 1 to allow
                                'for the header row
        .TextMatrix(1, 1) = Format(dCurrBal, "Currency", 2)  ' print current balance,
        .TextMatrix(1, 2) = Format(dCurrInt, "Currency", 2)  ' current interest,
        .TextMatrix(1, 3) = Format(dCurrPrin, "Currency", 2) ' current principal,
        .TextMatrix(1, 4) = Format(dEndBal, "Currency", 2)   ' end balance,
        .TextMatrix(1, 5) = Format(dCumInt, "Currency", 2)   ' cumulative interest, and
        .TextMatrix(1, 6) = Format(dTotPaid, "Currency", 2)  ' running total paid
    End With                                                 ' for the first period.
                                                             ' Remember that the FlexGrid row and column indices both begin at 0.
                                                             ' So the current balance will print in the 3rd column from the left,
                                                             ' the current interest will appear in the 4th, and so on.
    ' the all-important loop
    For iCountPay = 1 To iPayments ' 1 to the number of rows
        With MSFlexGrid1
            .TextMatrix(iCountPay, 0) = iCountPay ' populates the No. of payments column i.e. the first column on the left
            
            If iCountPay > 1 Then ' the loop for the financial figures really begins
                                  ' from the 3rd row (1st row=headers, 2nd was the first period,
                                  ' and was already populated above.
                
                ' starting balance for this period is read from ending balance from previous period
                dCurrBal = .TextMatrix(iCountPay - 1, 4)
                ' print the starting balance for this period
                .TextMatrix(iCountPay, 1) = Format(dCurrBal, "Currency", 2)
                ' multiply the starting balance by the interst rate (which is divided by 100 to provide a percentage)
                ' and divide by 12 to get the interest paid for this month.
                dCurrInt = .TextMatrix(iCountPay, 1) / 12 * (frmMain.mskRate / 100)
                ' print the interest paid this period
                .TextMatrix(iCountPay, 2) = Format(dCurrInt, "Currency", 2)
                ' this periods principal paid is the monthly total minus the interest paid
                dCurrPrin = frmMain.txtPayment - dCurrInt
                ' print the principal paid for this period
                .TextMatrix(iCountPay, 3) = Format(dCurrPrin, "Currency", 2)
                ' this periods end balance is the start balance minus principal paid this period
                dEndBal = dCurrBal - dCurrPrin
                ' print the end balance for this period
                .TextMatrix(iCountPay, 4) = Format(dEndBal, "Currency", 2)
                ' cumulative interest for this period is the cumulative interst paid in the last period
                ' plus interest paid this period
                dCumInt = .TextMatrix(iCountPay - 1, 5) + dCurrInt
                ' print the cumulative interest for this period
                .TextMatrix(iCountPay, 5) = Format(dCumInt, "Currency", 2)
                ' the total paid to date this period is simply the monthly payment
                ' multiplied by the payment number
                dTotPaid = .TextMatrix(1, 6) * iCountPay
                ' print the total paid at the end of this period
                .TextMatrix(iCountPay, 6) = Format(dTotPaid, "Currency", 2)
            End If
        End With
    Next iCountPay ' increment the all-important loop
    
    With MSFlexGrid1
        .ColWidth(0) = 640   ' Adjust the column widths
        .ColWidth(1) = 1975  ' in 'twips'.
        .ColWidth(2) = 1700
        .ColWidth(3) = 1800
        .ColWidth(4) = 1700
        .ColWidth(5) = 2000
        .ColWidth(6) = 1975
        Dim j
        For j = 0 To 6           ' All columns
            .ColAlignment(j) = 3 ' to be centrally aligned
        Next j
    End With

'Color alternating rows of the Grid to make it easier to
'read.
Dim iCols As Integer
    Do Until MSFlexGrid1.Row = iPayments - 1
        MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            For iCols = 0 To 6
                MSFlexGrid1.Col = iCols
                MSFlexGrid1.CellBackColor = &HC0C0C0
            Next iCols
                MSFlexGrid1.Row = MSFlexGrid1.Row + 1
    Loop
'This code is to correct the error that I would get from
'the above code if I left the Do Until statement = iPayments
'So, I just subtracted 1 from iPayments and put the code
'below to color the very last row.

            MSFlexGrid1.Row = iPayments
            For iCols = 0 To 6
                MSFlexGrid1.Col = iCols
                MSFlexGrid1.CellBackColor = &HC0C0C0
            Next iCols

MSFlexGrid1.Visible = True      ' show the grid

End Sub

