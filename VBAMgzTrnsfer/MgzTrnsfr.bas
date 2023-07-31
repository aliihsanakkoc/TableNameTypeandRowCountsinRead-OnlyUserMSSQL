Attribute VB_Name = "MgzTrnsfr"
Option Explicit

Public WsData, WsMesafe As Worksheet
Public ProductArticleBarcode, RequestStore, RequestStoreLocation, WarehouseName As String
Public RequestStoreDemandAmount, ProductBeginIndex, ProductEndIndex, LastRowData, LastRowMesafe, i As Integer
Public arrIndex As Variant: Public dictStoreLocation As New Dictionary

Public Sub transfer()

With Application
    .DisplayAlerts = False
    .ScreenUpdating = False
    .EnableEvents = False
    .Calculation = xlCalculationManual
End With
Stop
WarehouseName = "DEPO"

Set WsData = ThisWorkbook.Worksheets("DATA")
LastRowData = WsData.Cells(Rows.Count, 1).End(xlUp).row

Set WsMesafe = ThisWorkbook.Worksheets("MESAFE")
LastRowMesafe = WsMesafe.Cells(Rows.Count, 1).End(xlUp).row

CreateIndexArray
Stop
StoreLocation

For i = 2 To LastRowData
    If WsData.Cells(i, 10).Value < 0 Then
    ProductArticleBarcode = WsData.Cells(i, 1).Value
    RequestStore = WsData.Cells(i, 2).Value
    RequestStoreLocation = WsData.Cells(i, 3).Value
    RequestStoreDemandAmount = WsData.Cells(i, 10).Value
    GiverSelection
    End If
Next i

With Application
    .DisplayAlerts = True
    .ScreenUpdating = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
End With

End Sub

Private Sub CreateIndexArray()

    Dim j As Integer
    
    ReDim arrIndex(1 To LastRowData - 1, 1 To 2)

    For j = 2 To LastRowData
        arrIndex(j - 1, 1) = WsData.Cells(j, 1).Value
        arrIndex(j - 1, 2) = j
    Next j
    
End Sub

Private Sub FindIndex(ProductArticleBarcode)
    
    Dim j As Integer
    Dim dictIndex As New Dictionary

    For j = LBound(arrIndex) To UBound(arrIndex)
        If arrIndex(j, 1) = ProductArticleBarcode Then dictIndex.Add arrIndex(j, 2), arrIndex(j, 2)
    Next j
    
    ProductBeginIndex = Application.WorksheetFunction.Min(dictIndex.Items)
    ProductEndIndex = Application.WorksheetFunction.Max(dictIndex.Items)
    
    Set dictIndex = Nothing
 Stop
End Sub

Private Sub GiverSelection()
        
    Dim j As Integer
    Dim dictGiver As New Dictionary
    
    FindIndex (ProductArticleBarcode)
    
    For j = ProductBeginIndex To ProductEndIndex
        If WsData.Cells(j, 10).Value > 0 Then dictGiver.Add WsData.Cells(j, 2).Value, WsData.Cells(j, 10).Value
    Next j

    If dictGiver.Count = 0 Then GoTo NextProduct
    
    If dictGiver.Exists(WarehouseName) = True And dictGiver(WarehouseName) > 0 Then
    DeterminingTheAmountToBeGiven (WarehouseName)
    Else: GiveFromStore dictGiver.Items, dictGiver.Keys
    End If
    
    Set dictGiver = Nothing

NextProduct:

End Sub

Private Sub GiveFromStore(storesAmountGive As Variant, stores As Variant)
     
    Dim dictGiverStores As New Dictionary
    Dim giverStoreName As String
    Dim j, x As Integer
   
    For j = LBound(storesAmountGive) To UBound(storesAmountGive)
        If storesAmountGive(j) >= (RequestStoreDemandAmount * -1) Then dictGiverStores.Add stores(j), RequestStoreLocation & dictStoreLocation(stores(j))
    Next j
    
    If dictGiverStores.Count = 0 Then
        For x = LBound(storesAmountGive) To UBound(storesAmountGive)
            If storesAmountGive(x) > 0 Then dictGiverStores.Add stores(x), RequestStoreLocation & dictStoreLocation(stores(x))
        Next x
    End If
    
    giverStoreName = NearestStoreSelect(dictGiverStores.Keys, dictGiverStores.Items)
    
    Set dictGiverStores = Nothing
    
    DeterminingTheAmountToBeGiven (giverStoreName)
    
End Sub

Private Function NearestStoreSelect(giverStore As Variant, combinedLocation As Variant) As String
    
    Dim dictNearestStoreSelect As New Dictionary
    Dim x, y As Integer
    Dim key As Variant
    
    For x = LBound(giverStore) To UBound(giverStore)
        For y = 1 To LastRowMesafe
            If combinedLocation(x) = WsMesafe.Cells(y, 1).Value Then dictNearestStoreSelect.Add giverStore(x), WsMesafe.Cells(y, 4).Value
        Next y
    Next x
        
    For Each key In dictNearestStoreSelect.Keys
        If Application.WorksheetFunction.Min(dictNearestStoreSelect.Items) = dictNearestStoreSelect(key) Then NearestStoreSelect = key
    Next key

    Set dictNearestStoreSelect = Nothing

End Function

Private Sub DeterminingTheAmountToBeGiven(giverName As String)
    
    Dim givenAmount, giverIndex, j As Integer
    
    WsData.Cells(i, 11).Value = giverName

    For j = ProductBeginIndex To ProductEndIndex
        If WsData.Cells(j, 2).Value = giverName Then
        givenAmount = WsData.Cells(j, 10).Value
        giverIndex = j
        Exit For
        End If
    Next j
    
    If givenAmount + RequestStoreDemandAmount >= 0 Then
    WsData.Cells(giverIndex, 10).Value = givenAmount + RequestStoreDemandAmount
    WsData.Cells(i, 12) = RequestStoreDemandAmount * -1
    Else: WsData.Cells(giverIndex, 10).Value = 0: WsData.Cells(i, 12) = givenAmount
    End If

End Sub

Private Sub StoreLocation()

    Dim j As Integer
    
    On Error Resume Next
    For j = 2 To LastRowData
        dictStoreLocation.Add WsData.Cells(j, 2).Value, WsData.Cells(j, 3).Value
    Next j
    On Error GoTo 0
Stop
End Sub

Public Function MagazaDurum(ByVal MagazaAdi As Range, _
                            ByVal SatisMiktar As Range, _
                            ByVal StokMiktar As Range, _
                            ByVal RaftaBeklemeSuresi As Range, _
                            ByVal SatisMiktariKacGunluk As Integer, _
                            ByVal KacGundeBirMagazalarArasiTransferYapiliyor As Integer) As Integer

Dim StokYeterlilikMiktar, periyod As Integer

periyod = KacGundeBirMagazalarArasiTransferYapiliyor

If RaftaBeklemeSuresi > SatisMiktariKacGunluk Then RaftaBeklemeSuresi = SatisMiktariKacGunluk

StokYeterlilikMiktar = Int(Application.WorksheetFunction.Round((SatisMiktar / RaftaBeklemeSuresi) * periyod, 0))

If MagazaAdi = "DEPO" Then
MagazaDurum = StokMiktar
ElseIf SatisMiktar = 0 And RaftaBeklemeSuresi < 7 Then MagazaDurum = 0
ElseIf SatisMiktar = 0 And RaftaBeklemeSuresi >= 7 Then MagazaDurum = StokMiktar - 1
Else: MagazaDurum = StokMiktar - StokYeterlilikMiktar
End If

End Function
