SELECT 'CRS610MI'[minm],'ChgFinancial'[trnm],* FROM (
	SELECT 
	ptyorgid[CUNO],
	CSCD||'D' AS [CUCD],
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
