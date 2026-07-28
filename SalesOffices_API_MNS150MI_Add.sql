SELECT 'MNS150MI'[minm],'Add'[trnm],* FROM (
	SELECT CAST((ROW_NUMBER() OVER (ORDER BY ptysalesofficeid) + 200000) AS TEXT)[USID],
	SUBSTR(ptysalesofficepublicname,1,36)[NAME],
	'300'[DFCO],
	'4'[ULTP]
	FROM SyndigoSalesOffices
) as temp
WHERE 1=1
AND NAME <> ''
