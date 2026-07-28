SELECT 'CRS620MI'[minm],'AddSupplier'[trnm],* FROM (
	SELECT CAST((ROW_NUMBER() OVER (ORDER BY ptysalesofficeid) + 200000) AS TEXT)[SUNO],
	ptysalesofficeid[SCNO],
	SUBSTR(ptysalesofficepublicname,1,36)[SUNM],
	'0'[SUTY],
	countryCode[CSCD],
	'MDY'[DTFM],
	'D01'[ORTY],
	'1'[DT4T],
	'1'[DTCD],
	countryCode||'D'[CUCD],
	'1'[CRTP],
	'1'[ATPR],
	'ISP'[SUCL],
	'FOB'[TEDL],
	'001'[MODL],
	'FOB'[TEAF],
	'GB'[LNCD],
	'001'[TEPA],
	'N30'[TEPY],
	'EFT'[PYME]
	FROM SyndigoSalesOffices
) as temp
WHERE 1=1
AND SUNM <> ''
