select
	''[MESSAGE],
	IBPUNO[PUNO],
	IBPNLI[PNLI],
	IBPNLS[PNLS],
	'20' || substr(RECD09, 2, 2) || substr(RECD09, 4, 2) || substr(RECD09, 6, 2)[CODT],
	REPLACE(CAST(CAST(IBORQA AS INTEGER) AS TEXT), ',', '') [CFQA]
from MPLINE
JOIN pmp09 on ordn09 = SUBSTR(ibpuno,2,7) and line09 = ibpnli