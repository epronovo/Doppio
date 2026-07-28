SELECT 'CRS610MI'[minm],'Add'[trnm],* FROM (
	SELECT 'Z00003'[CUTM],
	SUBSTR(ptysalesofficepublicname,1,36)[CUNM],
	SUBSTR(street,1,36)[CUA1],
	'GB'[LNCD],
	ptysalesofficeid[CUNO],
	CASE WHEN LENGTH(street) > 36 THEN SUBSTR(street,37) ELSE '' END[CUA2],
	postalCode[PONO],
	countryCode[CSCD],
	stateCode[ECAR],
	'10'[STAT],
	SUBSTR(city,1,20)[TOWN]
	FROM SyndigoSalesOffices
) as temp
WHERE 1=1
AND CUNM <> '' AND CUA1 <> ''
AND NOT (LENGTH(CUNM) > 36 OR LENGTH(CUA1) > 36 OR LENGTH(TOWN) > 20)
