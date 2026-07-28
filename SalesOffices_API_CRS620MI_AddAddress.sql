SELECT 'CRS620MI'[minm],'AddAddress'[trnm],* FROM (
	SELECT CAST((ROW_NUMBER() OVER (ORDER BY ptysalesofficeid) + 200000) AS TEXT)[SUNO],
	'1'[ADTE],
	'001'[ADID],
	SUBSTR(ptysalesofficepublicname,1,36)[SUNM],
	SUBSTR(street,1,36)[ADR1],
	CASE WHEN LENGTH(street) > 36 THEN SUBSTR(street,37) ELSE '' END[ADR2],
	SUBSTR(city,1,20)[TOWN],
	stateCode[ECAR],
	postalCode[PONO],
	countryCode[CSCD]
	FROM SyndigoSalesOffices
) as temp
WHERE 1=1
AND SUNM <> '' AND ADR1 <> ''
