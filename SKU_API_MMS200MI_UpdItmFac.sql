SELECT 'MMS200MI'[minm],'UpdItmFac'[trnm],* FROM (
	SELECT 'RI1'[FACI],
	thgskuid[ITNO],
	'0'[LEA4],
	thghtcs[CSNO],
	'HJPCM1'[WSCA],
	'0'[CPL0],
	'0'[CPDC],
	'1'[VAMT],
	'0'[ALTS],
	'100'[REWH],
	'1'[DLET],
	'1'[MARC],
	'1'[FATM]
	FROM SyndigoSKU
) as temp
WHERE 1=1
AND ITNO <> ''
