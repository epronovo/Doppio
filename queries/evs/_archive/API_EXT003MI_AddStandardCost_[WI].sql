SELECT 
    ''[MESSAGE],
    Facility[FACI],
    Item[ITNO],
    ''[STRT],
    '3'[PCTP],
    '082625'[PCDT],
    Cost[CSU1] 
FROM CTOS_TerexItems, CTOS_Facilities 
LEFT JOIN MCHEAD ON KOITNO = Item
WHERE Facility = 'WI' AND KOITNO IS NULL
ORDER BY 1,2
