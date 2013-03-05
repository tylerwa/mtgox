Private pID As Variant
Private pRate As Variant
Private pFee As Variant
Private pDate As Variant
Private pValue As Variant
Private pIndex As Variant
Private pType As Variant
Private pText As Variant
Private pBalance As Variant
Private pDetails As Variant
'***********************************************************************
Public Property Get orderIndex() As Variant
    orderIndex = pIndex
End Property
Public Property Get OrderDate() As Variant
    OrderDate = pDate
End Property
Public Property Get OrderType() As Variant
    OrderType = pType
End Property
Public Property Get OrderText() As Variant
    OrderText = pText
End Property
Public Property Get OrderValue() As Variant
    OrderValue = pValue
End Property
Public Property Get OrderID() As Variant
    OrderID = pID
End Property
Public Property Get OrderRate() As Variant
    OrderRate = pRate
End Property
Public Property Get OrderFee() As Variant
    OrderFee = pFee
End Property
Public Property Get OrderBalance() As Variant
    OrderBalance = pBalance
End Property
'***********************************************************************
Public Property Get OrderDetail(ByVal index As Integer) As Variant
    OrderDetail = pDetails(index)
End Property
Public Property Get OrderDetails() As Variant
    OrderDetails = pDetails
End Property
Public Property Let SetOrderDetails(ByRef passedDetails As Variant)
    pDetails = passedDetails
    ParseOrderDetails
End Property
Public Property Let SetOrderIndex(ByVal newIndex As Integer)
    pDetails(0) = newIndex
    pIndex = newIndex
End Property
'***********************************************************************
Public Property Get MaxOrderDetails() As Variant
    MaxOrderDetails = UBound(pDetails)
End Property
Public Property Get MinOrderDetails() As Variant
    MinOrderDetails = LBound(pDetails)
End Property
Public Property Let SetMaxOrderDetails(ByVal newSize As Variant)
    ReDim Preserve pDetails(newSize)
End Property
'***********************************************************************
Public Property Let RawLine(ByVal text As String)
    Dim index As Integer
    Dim temp As Variant
    text = Replace(text, Chr(13), "")
    temp = Split(text, ",")
    ReDim pDetails(UBound(temp))
    pDetails = temp
    For index = LBound(pDetails) To UBound(pDetails)
        If Mid(pDetails(index), 1, 1) = """" Then
            pDetails(index) = Mid(pDetails(index), 2, Len(pDetails(index)) - 2)
        End If
    Next index
    ParseOrderDetails
End Property
Public Sub ParseOrderDetails()
    Dim intBeg As Integer
    Dim dolla As Integer
    Dim intEnd As Integer
    pIndex = pDetails(0)
    pDate = pDetails(1)
    pType = pDetails(2)
    pText = pDetails(3)
    pValue = pDetails(4)
    pBalance = pDetails(5)
    Select Case pType
        Case "withdraw"
            pID = "Withdraw @ " & pDate
            pRate = 0
            pFee = 0
        Case "deposit"
            pID = "Deposit @ " & pDate
            pRate = 0
            pFee = 0
        Case "spent", "in", "earned", "out"
            pID = Mid(pText, InStr(1, pText, "tid:") + 4, 16)
            pRate = Right(pText, 7)
            pFee = 0
        Case "fee"
            intEnd = InStr(1, pText, "%")
            intBeg = InStr(1, pText, "(")
            dolla = InStr(1, pText, "$") + 1
            pID = Mid(pText, InStr(1, pText, "tid:") + 4, 16)
            pRate = Mid(pText, dolla, InStr(dolla, pText, "(") - dolla)
            pFee = Mid(pText, intBeg + 1, intEnd - intBeg)
        Case Else
            pID = 0
            pRate = 0
            pFee = 0
    End Select
End Sub
Public Function PrintOrderDetail(ByVal index As Integer) As Variant
    On Error Resume Next
    Select Case index
        Case 0: PrintOrderDetail = pIndex
        Case 1: PrintOrderDetail = pDate
        Case 2: PrintOrderDetail = pType
        Case 3: PrintOrderDetail = pText
        Case 4: PrintOrderDetail = pValue
        Case 5: PrintOrderDetail = pBalance
        Case 6: PrintOrderDetail = pID
        Case 7: PrintOrderDetail = pRate
        Case 8: PrintOrderDetail = pFee
        Case Else: PrintOrderDetail = ""
    End Select
    On Error GoTo 0
End Function
