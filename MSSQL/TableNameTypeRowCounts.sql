--don't forget to choose the database you want in available databases
CREATE TABLE #TableNameTypeRowCounts 
(Id INT IDENTITY(1,1), TableName NVARCHAR(200), TableType NVARCHAR(10), RowCounts INT )
INSERT INTO #TableNameTypeRowCounts
SELECT T.TABLE_NAME, T.TABLE_TYPE, NULL FROM INFORMATION_SCHEMA.TABLES T

DECLARE @I AS INT=1
DECLARE @LastRow AS INT= (SELECT COUNT(*) FROM #TableNameTypeRowCounts)
DECLARE @TableName AS NVARCHAR(200)
DECLARE @Query as NVARCHAR(300)
DECLARE @RowCounts AS INT

WHILE @I<=@LastRow
BEGIN	
	DROP TABLE #TableRow
	CREATE TABLE #TableRow ( rowsCounter INT)
	SET @TableName=(SELECT T.TableName FROM #TableNameTypeRowCounts T WHERE T.Id=@I)
	BEGIN TRY
		SET @Query = 'SELECT COUNT(*) AS rowsCounter FROM ' +@TableName
		INSERT #TableRow EXEC sp_executesql @Query
		SELECT @RowCounts=rowsCounter From #TableRow
		UPDATE #TableNameTypeRowCounts SET RowCounts=@RowCounts WHERE Id=@I
	END TRY
	BEGIN CATCH
		UPDATE #TableNameTypeRowCounts SET RowCounts=-1 WHERE Id=@I
		SET @I=@I+1
	END CATCH
SET @I=@I+1	
END

SELECT * FROM #TableNameTypeRowCounts
