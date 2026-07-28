SELECT 'MMS200MI'[minm],'AddItmBasic'[trnm],* FROM (
	SELECT thgskuid[ITNO],
	thgskuname[ITDS],
	thgskudescription[FUDS],
	thgskustatus[STAT],
	'HJDATALOAD'[RESP],
	thguom[UNMS],
	itemtype[ITTY],
	thgprimarymakebuy[MABU],
	'Y1'[PRVG],
	itemgroup[ITGR],
	productgroup[ITCL],
	thghjlegacyerpproductnumber[ITNE]
	FROM SyndigoSKU
) as temp
WHERE 1=1
AND ITNO <> '' AND ITDS <> ''
