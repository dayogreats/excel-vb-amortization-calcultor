Sub CalculateAmortizationSchedule()
    Dim LoanAmount As Double
    Dim InterestRate As Double
    Dim LoanTerm As Integer
    Dim MonthlyInterestRate As Double
    Dim MonthlyPayment As Double
    Dim i As Integer
    Dim PaymentDate As Date
    Dim PrincipalPayment As Double
    Dim InterestPayment As Double
    Dim TotalPayment As Double
    Dim OutstandingBalance As Double
    
    ' Read loan details from the worksheet or form controls
    LoanAmount = Range("A1").Value ' Loan amount cell reference
    InterestRate = Range("A2").Value ' Interest rate cell reference (in decimal form)
    LoanTerm = Range("A3").Value ' Loan term in months cell reference
    
    ' Calculate monthly interest rate and payment
    MonthlyInterestRate = InterestRate / 12
    MonthlyPayment = LoanAmount * MonthlyInterestRate / (1 - (1 + MonthlyInterestRate) ^ -LoanTerm)
    
    ' Clear the amortization schedule area
    Range("E2:H100").ClearContents ' Adjust the range as needed
    
    ' Calculate and display the amortization schedule
    For i = 1 To LoanTerm
        PaymentDate = DateAdd("m", i - 1, Date) ' Calculate payment date
        PrincipalPayment = MonthlyPayment - (LoanAmount * MonthlyInterestRate)
        InterestPayment = MonthlyPayment - PrincipalPayment
        TotalPayment = MonthlyPayment
        OutstandingBalance = LoanAmount - PrincipalPayment
        
        ' Display payment details in the worksheet
        Cells(i + 1, 5).Value = i
        Cells(i + 1, 6).Value = PaymentDate
        Cells(i + 1, 7).Value = PrincipalPayment
        Cells(i + 1, 8).Value = InterestPayment
        Cells(i + 1, 9).Value = TotalPayment
        Cells(i + 1, 10).Value = OutstandingBalance
        
        ' Update the outstanding balance for the next iteration
        LoanAmount = OutstandingBalance
    Next i
End Sub
