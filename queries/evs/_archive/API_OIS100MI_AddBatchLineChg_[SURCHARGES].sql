SELECT 	
	''[MESSAGE],
	oaorno[ORNO],
	obponr[PONR],
	'0'[POSX],
	CASE WHEN surcharge= 'SUR' then 'SUR08' when surcharge = 'SUR1' then 'SUR15' when surcharge = 'SUR2' then 'SUR25' else '' end[CRID],
	ORDN55||ordl55[CRD0],
	''[END]
	FROM (
	SELECT DISTINCT cono55,o.cusn55,o.dseq55,o.ordn55,o.ordl55,dtso40,s.uprc55,s.OutstandingQty,
	CASE
	  WHEN txb120 LIKE 'SUR%' THEN trim(txb120)
	  WHEN txb220 LIKE 'SUR%' THEN trim(txb220)
	  WHEN txb320 LIKE 'SUR%' THEN trim(txb320)
	  WHEN txb420 LIKE 'SUR%' THEN trim(txb420)
	  WHEN txb520 LIKE 'SUR%' THEN trim(txb520)
	  ELSE NULL
	END AS surcharge
	FROM oep55 o
	JOIN OCOU_Summary s on s.ORDN55 = o.ordn55 and s.ordl55 = o.ordl55 
	LEFT JOIN oep40 on ordn40 = o.ordn55 
	LEFT JOIN usp20 on cusn20 = o.cusn55 and dseq20 = o.dseq55
	WHERE DTSO40 BETWEEN 1250305 AND 1250406
	AND (trim(txb120) IN ('SUR','SUR1','SUR2') 
	  or trim(txb220) IN ('SUR','SUR1','SUR2') 
	  or trim(txb320) IN ('SUR','SUR1','SUR2') 
	  or trim(txb420) IN ('SUR','SUR1','SUR2') 
	  or trim(txb520) IN ('SUR','SUR1','SUR2'))
) as temp
left join oohead on OAYREF = ORDN55
left join ooline on OBORNO = OAORNO and OBCUPO = ordl55
WHERE 1=1
AND cast(UPRC55 as double) <> 0	
