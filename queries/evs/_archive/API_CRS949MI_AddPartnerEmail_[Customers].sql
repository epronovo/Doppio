SELECT DISTINCT
''[MESSAGE],
'088'[DIVI],'231'[DONR],printf('%05d', caml.CUNO)[PRF1],printf('%05d', caml.ADID)[PRF2],'MAIL'[MEDC],'M3_IDM_XML_CSV'[SIID],COALESCE(EMAIAU,'springsales@asraymond.com')[TOMA],'noreply_m3@onebarnes.com'[FRMA],'ASRaymond Ord Confirm-<UF01>-<UF02>-<UF03>'[SUBJ],'0'[CPPL],'1'[FIET],'CUSTOM_MAIL_TEM'[EMBT],'01'[FILM],'ASRaymond Ord Confirm-<UF01>-<UF02>-<UF03>'[FNAM],'0'[PRTP]
,caml.CUSN05,caml.DSEQ05
FROM CAML_CustomerAddressMaster caml 
JOIN SLP05 s ON s.CUSN05 = caml.CUSN05 AND s.DSEQ05 = caml.DSEQ05
LEFT JOIN AUXADR_CONT1 ON CUSNAU = caml.CUSN05 AND DSEQAU = caml.DSEQ05
UNION
SELECT DISTINCT
''[MESSAGE],
'088'[DIVI],'380'[DONR],printf('%05d', caml.CUNO)[PRF1],printf('%05d', caml.ADID)[PRF2],'MAIL'[MEDC],'M3_IDM_XML_CSV'[SIID],COALESCE(EMAIAU,'springsales@asraymond.com')[TOMA],'noreply_m3@onebarnes.com'[FRMA],'ASRaymond Inv-<UF01>-<UF02>-<UF03>'[SUBJ],'0'[CPPL],'1'[FIET],'CUSTOM_MAIL_TEM'[EMBT],'01'[FILM],'ASRaymond Inv-<UF01>-<UF02>-<UF03>'[FNAM],'0'[PRTP]
,caml.CUSN05,caml.DSEQ05
FROM CAML_CustomerAddressMaster caml 
JOIN SLP05 s ON s.CUSN05 = caml.CUSN05 AND s.DSEQ05 = caml.DSEQ05
LEFT JOIN AUXADR_INVB1 ON CUSNAU = caml.CUSN05 AND DSEQAU = '000' AND TYPEAU='C'
ORDER BY 1,2,3,4
