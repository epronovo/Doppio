SELECT l.ORDN55,count(Distinct l.OSRC55)cnt
FROM OCOU_Summary l
LEFT JOIN oep40 ON ORDN40 = l.ORDN55
left JOIN OCOU_Kaller k ON k.ORDN55 = l.ORDN55 and k.ordl55 = l.ordl55
left JOIN OCOU_SOD_NEEDASRPART s ON s.ORDN55 = l.ORDN55 and s.ordl55 = l.ordl55
LEFT join CAML_CustomerAddressMaster cam on cam.CUSN05 = CUSN40 and cam.DSEQ05 = DSEQ40 
WHERE 1=1
	AND l.CATN55 not IN ('CA','COC','EXP001','FIRST ARTICLE CA','PLATING CERT','PPAP LEVEL1','PPAP LEVEL2','PPAP LEVEL3','PPAP LEVEL4','QE INSP REQ''D','BAGGING')  -- NO CHARGES
	AND k.ordn55 is null   -- NO OCOU_Kaller PARTS
	AND s.ordn55 is null   -- NO SOD PARTS
GROUP BY l.ORDN55
having cnt > 1