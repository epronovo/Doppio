SELECT 'CRS335MI'[minm],'AddCtrlObj'[trnm],* FROM (
	SELECT CAST((ROW_NUMBER() OVER (ORDER BY ptysalesofficeid) + 200000) AS TEXT)[ACRF],
	SUBSTR(ptysalesofficepublicname,1,40)[TX40],
	SUBSTR(ptysalesofficepublicname,1,15)[TX15]
	FROM SyndigoSalesOffices
) as temp
WHERE 1=1
AND TX40 <> ''
