SELECT DISTINCT
	''[MESSAGE],
	oaorno[ORNO],
	l.CATN55[ITNO],
	CAST(ROUND(l.OutstandingQty) AS TEXT)[ORQT],
	'US1'[WHLO],
	'20' || substr(o.CRQD55, 2, 2) || substr(o.CRQD55, 4, 2) || substr(o.CRQD55, 6, 2)[DWDT], -- JBA Order Line Promised Date
	'20' || substr(o.DTDR55, 2, 2) || substr(o.DTDR55, 4, 2) || substr(o.DTDR55, 6, 2)[CODT],
	CASE WHEN o.DTDR55 = 0 THEN '' ELSE '20' || substr(o.DTDR55, 2, 2) || substr(o.DTDR55, 4, 2) || substr(o.DTDR55, 6, 2) END [UID1], -- EDI Ship Date JBA Order Line Ship Date
	l.UPRC55[SAPR],
	CASE WHEN cam.ADID is null then '' else printf('%05d', cam.ADID) end [ADID],
	l.ORDL55[CUPO],
	CASE WHEN k.ordn55 is not null then '2' else '' end as [LTYP],
	CASE WHEN cast(l.CUNO as Integer) IN ('10005','10008','10009','10012','14385') AND instr(CUSO40, '/') > 0 THEN substr(CUSO40, 1, instr(CUSO40, '/') - 1)
		 WHEN cast(l.CUNO as Integer) IN ('10020','10151','10427','10851','12210','14928','14933','14936','15190') THEN trim(substr(CUSO40, 1, 10))
		 ELSE CUSO40 END AS CUOR,
	CASE WHEN cast(l.CUNO as Integer) IN ('10005','10008','10009','10012','14385') AND instr(CUSO40, '/') > 0 THEN substr(CUSO40,instr(CUSO40, '/') + 1,instr(substr(CUSO40, instr(CUSO40, '/') + 1), '/') - 1	) 
		 ELSE '' END AS UCA1, -- Cust line no.
	CASE WHEN cast(l.CUNO as Integer) IN ('10005','10008','10009','10012','14385') AND instr(CUSO40, '/') > 0 THEN substr(CUSO40,instr(CUSO40, '/') + instr(substr(CUSO40, instr(CUSO40, '/') + 1), '/') + 1) 
		 ELSE '' END AS UCA2, -- Release Number
	CASE WHEN ORPOPN is null then '' else ORPOPN end as [UCA4], -- Customer's Part Number
	''[END],
	l.ORDN55,
	l.ORDL55
FROM OCOU_Summary l
LEFT JOIN oep40 ON ORDN40 = l.ORDN55
LEFT JOIN OCOU_Kaller k ON k.ORDN55 = l.ORDN55 and k.ordl55 = l.ordl55
LEFT JOIN OCOU_SOD_NEEDASRPART s ON s.ORDN55 = l.ORDN55 and s.ordl55 = l.ordl55
LEFT JOIN CAML_CustomerAddressMaster cam on cam.CUSN05 = CUSN40 and cam.DSEQ05 = DSEQ40 
LEFT JOIN oep55 o ON o.ORDN55 = l.ORDN55 and o.ordl55 = l.ordl55
LEFT JOIN OXHEAD on OAYREF = l.ORDN55
LEFT JOIN MITMAS on MMITNO = l.CATN55
LEFT JOIN OCUSIT ON ORCUNO = CAST(l.CUNO as Integer) and ORITNO = l.CATN55
WHERE 1=1
	AND l.CATN55 not IN ('CA','COC','EXP001','FIRST ARTICLE CA','PLATING CERT','PPAP LEVEL1','PPAP LEVEL2','PPAP LEVEL3','PPAP LEVEL4','QE INSP REQ''D','BAGGING')  -- NO CHARGES
	AND k.ordn55 is null   	-- NO OCOU_Kaller PARTS
	AND s.ordn55 is null   	-- NO SOD PARTS
	AND MMITNO is not null 	-- Item must exist in M3
