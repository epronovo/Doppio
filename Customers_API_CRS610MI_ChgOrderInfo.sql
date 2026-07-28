SELECT 'CRS610MI'[minm],'ChgOrderInfo'[trnm],
* FROM (
	SELECT 
	ptyorgid[CUNO],
	CASE WHEN ptyorgtype = 'Scholastic' THEN 'SCH' 
		WHEN ptyorgtype = 'Commercial' THEN 'COM' 
		WHEN ptyorgtype = 'College' THEN 'COL' 
		WHEN ptyorgtype = 'Greek' THEN 'GRK' 
	ELSE '' END AS [CUCL],
	case when ptyorglegacyaccountrecievablenumber = '' THEN 'MDM:'||mdr_pid else ptyorglegacyaccountrecievablenumber end as [OREF],
	'?'[SMCD],
	''[PYNO],
	''[INRC],
	''[DOGR],
	'0'[ADBO],
	'0'[AICD],
	'0'[BOP1],
	case when ptyorgmailtoaddresscity = '' then UPPER(ptyorgmdrmcity) else UPPER(ptyorgmailtoaddresscity) end as [TOWN],
	case when mdr_dept_name_full = '' and ptyorgmailtostreet1 = '' then UPPER(ptyorgmdrmstreet)
	     when mdr_dept_name_full <> '' then UPPER('dept:' || mdr_dept_name_full) 
	     else ptyorgmailtostreet1 end as [CUA1],
	SUBSTR(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(UPPER(ptyorgpublicname), "'", ''), '*', ''), '"', ''), ' & ', '&'), '& ', '&'), ' &', '&'), '&', ' & '),1,36) as [CUNM]
	FROM SyndigoOrg
	WHERE mdr_pid = ''
) as temp
WHERE 1=1 
AND CUNM <> '' AND CUA1 <> ''
AND NOT (LENGTH(CUNM) > 36 OR LENGTH(CUA1) > 36 OR LENGTH(TOWN) > 20)
