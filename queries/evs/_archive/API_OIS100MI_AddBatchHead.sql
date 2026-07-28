SELECT DISTINCT 
	''[MESSAGE],
	CASE WHEN l.DSEQ55 IN ('WEB', '999') THEN '10000'
		 WHEN cam.CUNO is not null THEN cam.CUNO
		 WHEN cml.CUNO is not null THEN cml.CUNO
		 END [CUNO],
 	CASE WHEN l.DSEQ55 IN ('WEB', '999') OR cam.ADID is null THEN ''
 		 ELSE printf('%05d', cam.ADID) 
 		 END [ADID],
	CASE WHEN cml.CUNO IN ('10002','10003','10004','10005','10007','10008','10009','10012','10020','10025','10026','10030','10048','10151','12210','14384','14385','14928','14929','14930','14931','14933','14935','14936','15190','15192') THEN 'EDI' 
	     WHEN l.DSEQ55 IN ('WEB', '999') THEN 'WEB'
	     ELSE 'D01'
	     END [ORTP],
	'20' || substr(hdr.DTDR40, 2, 2) || substr(hdr.DTDR40, 4, 2) || substr(hdr.DTDR40, 6, 2)[RLDT],
	'088'[FACI],
	CASE WHEN cast(l.CUNO as Integer) IN ('10005','10008','10009','10012','14385') AND instr(CUSO40, '/') > 0 THEN substr(CUSO40, 1, instr(CUSO40, '/') - 1)
		 WHEN cast(l.CUNO as Integer) IN ('10020','10151','10427','10851','12210','14928','14933','14936','15190') THEN trim(substr(CUSO40, 1, 10))
		 ELSE CUSO40 END AS CUOR,
	l.ORDN55[YREF],
	CASE WHEN trim(hdr.CURN40) = 'CAN' THEN 'CAD' ELSE trim(hdr.CURN40) END [CUCD],
	'20' || substr(hdr.DTCO40, 2, 2) || substr(hdr.DTCO40, 4, 2) || substr(hdr.DTCO40, 6, 2)[ORDT],
	'MECSVC'[RESP],
	CASE WHEN chrg.MODL IS NULL THEN ''
		 WHEN chrg.MODL IS 'FED' AND adr.COCD05 <>'USA' THEN 'FIG'	 
		 WHEN chrg.MODL IS 'FSD' AND adr.COCD05 <>'USA' THEN 'FIE'
		 ELSE chrg.MODL END AS MODL,
	COALESCE(A030,'')[UCA1], 						-- Freight Account
	CASE WHEN cml.CUNO = '10003' THEN 'MSC' 		
		 ELSE '' END AS [UCA2], 					-- Department
	CASE WHEN cml.CUNO = '10002' THEN '200009801' 	-- Grainger
		 WHEN cml.CUNO = '10003' THEN '000027127' 	-- MSC Industrial
		 ELSE '' END AS [UCA3], 					-- Customer's Vendor ID
	CASE WHEN cast(l.CUNO as Integer) IN ('10005','10008','10009','10012','14385') AND instr(CUSO40, '/') > 0 THEN substr(CUSO40,instr(CUSO40, '/') + instr(substr(CUSO40, instr(CUSO40, '/') + 1), '/') + 1) 
		 ELSE '' END AS [UCA5], 					-- 830 Release No.
	''[UCA6], 										-- Block Reason code
	COALESCE(TDMBMD,'')[UCA7], 						-- Ship-to DUNS ID
	''[END],
	hdr.CUSN40,
	hdr.DSEQ40
FROM OCOU_Summary l
LEFT JOIN oep40 hdr ON hdr.ORDN40 = l.ORDN55													-- order header
LEFT JOIN OCOU_Kaller k ON k.ORDN55 = l.ORDN55 and k.ordl55 = l.ordl55
LEFT JOIN OCOU_SOD_NEEDASRPART s ON s.ORDN55 = l.ORDN55 and s.ordl55 = l.ordl55
LEFT JOIN slp05 adr on adr.CUSN05 = l.CUSN55 and adr.DSEQ05 = l.DSEQ55 							-- ship-to address	
LEFT JOIN CAML_CustomerAddressMaster cam on cam.CUSN05 = trim(hdr.CUSN40) and cam.DSEQ05 = trim(hdr.DSEQ40)
LEFT JOIN CML_Current cml on cml.CUSN65 = trim(hdr.CUSN40)
LEFT JOIN oep50 chrg on chrg.ORDN50 = hdr.ORDN40													-- charges
LEFT JOIN CUGEX1 ON pk01 = cml.CUNO and pk03 = printf('%05d', cam.ADID)						 	-- freight account
LEFT JOIN MBMTRD ON TDIDTR = 84 and TDEXTP = cml.CUNO and TDMVXD = printf('%05d', cam.ADID)  	-- ship-to DUNS ID
WHERE 1=1
	AND l.CATN55 not IN ('CA','COC','EXP001','FIRST ARTICLE CA','PLATING CERT','PPAP LEVEL1','PPAP LEVEL2','PPAP LEVEL3','PPAP LEVEL4','QE INSP REQ''D','BAGGING')  -- NO CHARGES
	AND k.ordn55 is null   -- NO OCOU_Kaller PARTS
	AND s.ordn55 is null   -- NO SOD PARTS
ORDER BY 3