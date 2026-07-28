CREATE TABLE CUGEX1 AS
select 'OCUSAD'[FILE],opcuno[PK01],'01'[PK02],opadid[PK03]
,trim(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(REPLACE(TLTX60,'GROUND',''),'*',''),'ACCT',''),'#',''),'UPS',''),'FEDEX',''),'COLLECT',''),'COLLECT',''),'2ND DAY',''),'BLUE',''),'RED',''),'NDA',''),'P1',''),'FED',''),'PER KEVIN','')) [A030]
from OSYTXL 
left join OCUSAD on optxid = tltxid
where TLTXID IN (select optxid from OCUSAD where OPCUNO = '10003')
and TLTX60 <> ''
and TLLINO = 2
and (
tltx60 like '%UPS A%'
or tltx60 like '%UPS GRO%'
or tltx60 like '%UPS NDA%'
or tltx60 like '%UPS RED%'
or tltx60 like '%UPS COL%'
or tltx60 like '%FEDEX%'
or tltx60 like '%FED GRO%'
)
