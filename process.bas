Option Explicit
Sub Initiate()
    Application.ScreenUpdating = False
    Dim orderBook() As New Transaction
    Dim historyBTC() As New Transaction
    Dim historyUSD() As New Transaction
    Dim orderIDs() As Variant
    ReDim orderIDs(0)
    historyUSD = Import(ThisWorkbook.path & "/history_USD.csv")
    historyBTC = Import(ThisWorkbook.path & "/history_BTC.csv")
    CompileOrderIDs historyBTC, orderIDs
    CompileOrderIDs historyUSD, orderIDs
    orderBook = Aggregate(historyBTC, historyUSD, orderIDs)
    Sort orderBook
    SwapOrderWithFee orderBook
    PrintToWorksheet orderBook, "all"
    Application.ScreenUpdating = True
End Sub
Private Function Import(ByVal fullPath As String) As Transaction()
    Dim index As Integer
    Dim importData() As New Transaction
    Dim detailCount As Integer
    Dim currentLine As String
    Dim nextLine As String
    Dim textStream As Object
    Dim textStreamFSO As Object
    Set textStreamFSO = CreateObject("Scripting.FileSystemObject")
    Set textStream = textStreamFSO.OpenTextFile(fullPath, 1)
    index = 0
    Do While Not textStream.atendofstream
        ReDim Preserve importData(index)
        currentLine = textStream.readline
        If index = 0 Then
            detailCount = UBound(Split(currentLine, ","))
            currentLine = textStream.readline
            importData(index).RawLine = currentLine
        ElseIf UBound(Split(currentLine, ",")) < detailCount Then
            nextLine = textStream.readline
            importData(index).RawLine = currentLine & nextLine
        Else
            importData(index).RawLine = currentLine
        End If
        index = index + 1
    Loop
    Set textStream = Nothing
    Set textStreamFSO = Nothing
    Import = importData()
End Function
Private Sub CompileOrderIDs(ByRef sourceGroup() As Transaction, ByRef orderIDs As Variant)
    Dim sourceIndex As Integer
    Dim listIndex As Integer
    Dim alreadyListed As Boolean
    For sourceIndex = LBound(sourceGroup) To UBound(sourceGroup)
        alreadyListed = False
        For listIndex = LBound(orderIDs) To UBound(orderIDs)
            If sourceGroup(sourceIndex).OrderID = orderIDs(listIndex) Then
                alreadyListed = True
            End If
        Next
        If Not alreadyListed Then
            orderIDs(UBound(orderIDs)) = sourceGroup(sourceIndex).OrderID
            If sourceIndex < UBound(sourceGroup) Then
                ReDim Preserve orderIDs(UBound(orderIDs) + 1)
            End If
        End If
    Next
End Sub
Private Function Aggregate(ByRef historyBTC() As Transaction, ByRef historyUSD() As Transaction, _
    ByRef orderIDs As Variant) As Transaction()
    Dim orderBook() As New Transaction
    Dim listIndex As Integer
    Dim btcIndex As Integer
    Dim usdIndex As Integer
    Dim index As Integer
    index = -1
    For listIndex = LBound(orderIDs) To UBound(orderIDs)
        For btcIndex = LBound(historyBTC) To UBound(historyBTC)
            If historyBTC(btcIndex).OrderID = orderIDs(listIndex) Then
                If historyBTC(btcIndex).OrderType <> "out" Then
                    index = index + 1
                    ReDim Preserve orderBook(index)
                    orderBook(index).SetOrderDetails = historyBTC(btcIndex).OrderDetails
                End If
            End If
        Next
        For usdIndex = LBound(historyUSD) To UBound(historyUSD)
            If historyUSD(usdIndex).OrderID = orderIDs(listIndex) Then
                If historyUSD(usdIndex).OrderType <> "spent" Then
                    index = index + 1
                    ReDim Preserve orderBook(index)
                    orderBook(index).SetOrderDetails = historyUSD(usdIndex).OrderDetails
                End If
            End If
        Next
    Next
    Aggregate = orderBook()
End Function
Public Sub Sort(ByRef orderBook() As Transaction)
    Dim finished As Boolean
    Dim index As Long
    Dim temp As New Transaction
    Do
        finished = True
        For index = LBound(orderBook) To UBound(orderBook) - 1
            If orderBook(index).OrderDate < orderBook(index + 1).OrderDate Then
                finished = False
                temp.SetOrderDetails = orderBook(index).OrderDetails
                orderBook(index).SetOrderDetails = orderBook(index + 1).OrderDetails
                orderBook(index + 1).SetOrderDetails = temp.OrderDetails
            End If
        Next
    Loop While Not finished
End Sub
Public Sub SwapOrderWithFee(ByRef orderBook() As Transaction)
    Dim justSwapped As Boolean
    Dim index As Long
    Dim temp As New Transaction
    justSwapped = False
    For index = LBound(orderBook) To UBound(orderBook) - 1
        If orderBook(index).OrderType = "in" Or orderBook(index).OrderType = "earned" Then
            If justSwapped = True Then
                justSwapped = False
            Else
                temp.SetOrderDetails = orderBook(index).OrderDetails
                orderBook(index).SetOrderDetails = orderBook(index + 1).OrderDetails
                orderBook(index + 1).SetOrderDetails = temp.OrderDetails
                justSwapped = True
            End If
        End If
        orderBook(index).SetOrderIndex = UBound(orderBook) - index
    Next
    orderBook(UBound(orderBook)).SetOrderIndex = UBound(orderBook) - index
End Sub
Private Sub PrintToWorksheet(ByRef orderGroup() As Transaction, ByVal nameOfSheet As String)
    Dim orderIndexA As Integer
    Dim orderIndexB As Integer
    Dim tempSheet As Worksheet
    On Error Resume Next
    Set tempSheet = ActiveWorkbook.Worksheets(nameOfSheet)
    On Error GoTo 0
    If tempSheet Is Nothing Then
        Worksheets.Add.name = nameOfSheet
    End If
    InsertHeaders nameOfSheet
    For orderIndexA = LBound(orderGroup) To UBound(orderGroup)
        For orderIndexB = orderGroup(orderIndexA).MinOrderDetails To orderGroup(orderIndexA).MaxOrderDetails + 3
            Sheets(nameOfSheet).Cells(orderIndexA + 2, orderIndexB + 1) = "'" & _
                orderGroup(orderIndexA).PrintOrderDetail(orderIndexB)
        Next
    Next
    Set tempSheet = Nothing
End Sub
Sub InsertHeaders(ByVal nameOfSheet As String)
    Dim tempSheet As Worksheet
    Set tempSheet = Sheets(nameOfSheet)
    With tempSheet
        .Cells(1, 1) = "Index"
        .Cells(1, 2) = "Date"
        .Cells(1, 3) = "Type"
        .Cells(1, 4) = "Info"
        .Cells(1, 5) = "Value"
        .Cells(1, 6) = "Balance"
        .Cells(1, 7) = "TID"
        .Cells(1, 8) = "Rate"
        .Cells(1, 9) = "Fee %"
    End With
End Sub
