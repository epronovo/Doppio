SELECT 'CRS610MI'[minm],'Add'[trnm],* FROM (
	SELECT 'Z00001'[CUTM],
	'300'[CONO],
	''[DIVI],
	'GB'[LNCD],
	ptyorgid[CUNO],
	SUBSTR(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(UPPER(ptyorgpublicname), "'", ''), '*', ''), '"', ''), ' & ', '&'), '& ', '&'), ' &', '&'), '&', ' & '),1,36) as [CUNM],
	case when mdr_dept_name_full = '' and ptyorgmailtostreet1 = '' then UPPER(ptyorgmdrmstreet)
	     when mdr_dept_name_full <> '' then UPPER('dept:' || mdr_dept_name_full) 
	     else ptyorgmailtostreet1 end as [CUA1],
	case when mdr_dept_name_full = '' then ptyorgmailtostreet2 else ptyorgmdrmstreet end as [CUA2],
	case when ptyorgmailtoaddresspostalcode = '' then ptyorgmdrmzipcode else ptyorgmailtoaddresspostalcode end as [PONO],
	PHONE[PHNO],
	''[PHN2],
	FAX[TFNO],
	'0'[CUTP],
	''[YREF],
	''[YRE2],
	CSCD[CSCD],
	case when ptyorgmailtoaddressstate = '' then ptyorgmdrmstate else ptyorgmailtoaddressstate end as [ECAR],
	'20'[STAT],
	case when ptyorgmailtoaddresscity = '' then UPPER(ptyorgmdrmcity) else UPPER(ptyorgmailtoaddresscity) end as [TOWN],
	CSCD||ptyorgmailtoaddressstate[EDES]
	FROM SyndigoOrg
	WHERE mdr_pid = ''
) as temp
WHERE 1=1 
AND CUNM <> '' AND CUA1 <> ''
AND NOT (LENGTH(CUNM) > 36 OR LENGTH(CUA1) > 36 OR LENGTH(TOWN) > 20)