SELECT DISTINCT 
	''[MESSAGE],
	oaorno[ORNO],
	'1'[ADRT],
 	trim(dsa.ONAM45)[CUNM],
 	trim(dsa.OAD145)[CUA1],
 	trim(dsa.OAD245)[CUA2],
 	trim(dsa.OAD345)[CUA3],
 	trim(dsa.OAD445)[TOWN],
 	SUBSTR(trim(dsa.OAD545), 1, 2)[ECAR],
 	CASE WHEN TRIM(SUBSTR(dsa.OAD545, -3)) = 'USA' THEN 'US' 
 		 WHEN TRIM(SUBSTR(dsa.OAD545, -3)) = 'CDN' THEN 'CA' 
 		 ELSE TRIM(SUBSTR(dsa.OAD545, -3))
 		 END[CSCD],
 	TRIM(dsa.OPST45)[PONO],
	CUSN40,
	DSEQ40,
	''[END]
FROM OCOU_Summary l
LEFT JOIN oep40 ON ORDN40 = l.ORDN55
left JOIN OCOU_Kaller k ON k.ORDN55 = l.ORDN55 and k.ordl55 = l.ordl55
left JOIN OCOU_SOD_NEEDASRPART s ON s.ORDN55 = l.ORDN55 and s.ordl55 = l.ordl55
LEFT join CAML_CustomerAddressMaster cam on cam.CUSN05 = CUSN40 and cam.DSEQ05 = DSEQ40 
left join oxhead on OAYREF = l.ORDN55
JOIN oep45 dsa on ordn45 = l.ORDN55 and seqn45 = '1'
WHERE 1=1
	AND l.DSEQ55 IN ('WEB', '999')
	AND l.CATN55 not IN ('CA','COC','EXP001','FIRST ARTICLE CA','PLATING CERT','PPAP LEVEL1','PPAP LEVEL2','PPAP LEVEL3','PPAP LEVEL4','QE INSP REQ''D','BAGGING')  -- NO CHARGES
	AND k.ordn55 is null   -- NO OCOU_Kaller PARTS
	AND s.ordn55 is null   -- NO SOD PARTS
ORDER BY 3