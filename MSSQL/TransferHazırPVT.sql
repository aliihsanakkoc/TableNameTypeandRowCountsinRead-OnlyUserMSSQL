CREATE TABLE #Transfer (StateDescription NVARCHAR(10),WarehouseCode NVARCHAR(10), ItemCode NVARCHAR(30), 
ItemDescription NVARCHAR(200), ColorCode NVARCHAR(10), SizeCode NVARCHAR(10), Sayi FLOAT)

INSERT INTO #Transfer
SELECT 'SATI�',[Depo Kodu], [Madde Kodu], [Madde A��klamas�], [Renk Kodu], [ItemDim1Code], SUM([Miktar]) FROM [TR].[Faturalar]
WHERE [Madde Kodu] IN ('') AND [Fatura Tarihi]>''
GROUP BY [Depo Kodu], [Madde Kodu], [Madde A��klamas�], [Renk Kodu], [ItemDim1Code]

INSERT INTO #Transfer
SELECT 'ENVANTER', [Depo Kodu], [Madde Kodu], [Madde A��klamas�], [Renk Kodu], [ItemDim1Code], [Envanter] FROM [TR].[Envanter] 
WHERE [Madde Kodu] IN ('')

CREATE TABLE #EnSonTransferTarihi (StateDescription NVARCHAR(10),WarehouseCode NVARCHAR(10), ItemCode NVARCHAR(30), 
ItemDescription NVARCHAR(200), ColorCode NVARCHAR(10), SizeCode NVARCHAR(10), TransferTarihi DATE)

INSERT INTO #EnSonTransferTarihi
SELECT 'RAFTAKALMA', [Depo Kodu], [Madde Kodu], [Madde A��klamas�], [Renk Kodu], [ItemDim1Code], MAX([Belge Tarihi]) FROM [TR].[StokGirisCikis]
WHERE [S�re� Kodu]='S' AND [Stok Hareket �ekli]=1 AND [�ade]=0 AND
[Madde Kodu] IN ('')
GROUP BY [Depo Kodu], [Madde Kodu], [Madde A��klamas�], [Renk Kodu], [ItemDim1Code]


INSERT INTO #Transfer
SELECT T.StateDescription,T.WarehouseCode, T.ItemCode, T.ItemDescription, T.ColorCode, T.SizeCode, DATEDIFF(DAY, T.TransferTarihi, CAST(GETDATE() AS DATE))  FROM #EnSonTransferTarihi T

SELECT TransferPivot.* FROM
(SELECT * FROM #Transfer) Trnsfr
PIVOT (SUM(Sayi) FOR StateDescription IN ([SATI�],[ENVANTER],[RAFTAKALMA]) ) TransferPivot
