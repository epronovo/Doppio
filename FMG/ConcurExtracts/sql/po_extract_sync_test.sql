/*=====================================================================================================
  Concur PO extract - three-proc sync test
  -----------------------------------------------------------------------------------------------
  Procs under test:
     p_concurinvoicepo_get_poheader   (no param)  -> 200 header records
     p_concurinvoicepo_get_poline     (@po_num)   -> 300 line records (UNION: real lines + charges)
     p_concurinvoicepo_get_poreceipt  (no param)  -> receipt records

  Rule being tested:  if a PO is filtered OUT of the header, no line and no receipt row may
                      survive for that PO either.

  Section 1 - PO list (EDIT HERE; leave the table empty to run against everything)
  Section 2 - Header decision matrix: which filter keeps or kills each PO
  Section 3 - What each proc will emit for those POs (mirrors proc logic, key columns only)
  Section 4 - SYNC CHECKS.  Every row returned in 4A-4I is a defect (4J is informational).
              4A is now a regression test for the vendor filter added to _get_poreceipt.
  Section 5 - One-line verdict per PO
  Section 6 - Run the real procs (optional, to eyeball actual output)

  Read-only.  Run the whole file at once (it is a single batch after the USE).
=====================================================================================================*/
use App_Concur;
go
set nocount on;


/*-----------------------------------------------------------------------------------------------
  1.  PO's UNDER TEST                                                                  << EDIT >>
      Comment the INSERT out entirely to test EVERY PO in the extract.
-----------------------------------------------------------------------------------------------*/
if object_id('tempdb..#po') is not null drop table #po;
create table #po (po_num varchar(20) not null primary key);

insert into #po (po_num) values
     ('7806040')
    ,('1100000006')
;

declare @all bit = case when exists (select 1 from #po) then 0 else 1 end;


/*-----------------------------------------------------------------------------------------------
  Build the three working sets, each mirroring its proc's own filters exactly.
-----------------------------------------------------------------------------------------------*/

---------------------------------------------------------------------- headers + every filter flag
if object_id('tempdb..#hdr') is not null drop table #hdr;

select
     poh.system_id
    ,poh.system_entity_id
    ,poh.entity_id
    ,poh.po_num
    ,purchase_order_number = poh.entity_id + poh.po_num
    ,poh.vend_num
    ,poh.vend_seq
    ,poh.status_code
    ,ven.vendor_code
    ,ven.vendor_address_code
    -- filter 1: coalesce(poh.status_code,1)=1                    [header + line + receipt: all three]
    ,f1_status_ok            = f.f1
    -- filter 2: ven.vendor_code is not null                      [header + receipt]
    ,f2_vendor_hdr_ok        = f.f2
    -- filter 2b: nullif(ltrim(rtrim(vendor_code)),'') not null   [line proc CHARGE leg only]
    ,f2b_vendor_charge_ok    = f.f2b
    -- filter 3: PO needs >=1 line that has a receipt             [HEADER ONLY, joined on entity_id]
    --            (the receipt proc's inner join to vw_concur_extract_po_line satisfies this implicitly)
    ,f3_line_with_receipt_ok = f.f3
    ,ships_header            = case when f.f1 = 1 and f.f2 = 1 and f.f3 = 1 then 1 else 0 end
    ,line_count    = (select count(*) from dbo.vw_concur_extract_po_line pol
                       where pol.system_id=poh.system_id and pol.system_entity_id=poh.system_entity_id
                         and pol.po_num=poh.po_num)
    ,charge_count  = (select count(*) from dbo.vw_concur_extract_po_line_charges polc
                       where polc.system_id=poh.system_id and polc.system_entity_id=poh.system_entity_id
                         and polc.po_num=poh.po_num and polc.delete_flag='N')
    ,receipt_count = (select count(*) from dbo.vw_concur_extract_po_line_receipt polr
                       where polr.system_id=poh.system_id and polr.system_entity_id=poh.system_entity_id
                         and polr.po_num=poh.po_num)
into #hdr
from dbo.vw_concur_extract_po_header poh
    left outer join dbo.vw_concur_extract_po_vendor ven
        on  ven.system_id        = poh.system_id
        and ven.system_entity_id = poh.system_entity_id
        and ven.vend_num         = poh.vend_num
        and ven.vend_seq         = poh.vend_seq
    cross apply (
        select
             f1  = case when coalesce(poh.status_code,1) = 1 then 1 else 0 end
            ,f2  = case when ven.vendor_code is not null then 1 else 0 end
            ,f2b = case when nullif(ltrim(rtrim(ven.vendor_code)),'') is not null then 1 else 0 end
            ,f3  = case when exists (
                        select 1
                        from dbo.vw_concur_extract_po_line pol
                            inner join dbo.vw_concur_extract_po_line_receipt plr
                                    on  pol.po_num    = plr.po_num
                                    and pol.po_line   = plr.po_line
                                    and pol.po_seq    = plr.po_seq
                                    and pol.entity_id = plr.entity_id
                        where pol.entity_id = poh.entity_id
                          and pol.po_num    = poh.po_num
                   ) then 1 else 0 end
    ) f
where (@all = 1 or poh.po_num in (select po_num from #po));


---------------------------------------------------------------------- 300 records, both UNION legs
if object_id('tempdb..#line') is not null drop table #line;

select
     leg              = cast('1-line' as varchar(8))
    ,pol.system_id, pol.system_entity_id, pol.entity_id, pol.po_num, pol.po_line, pol.po_seq
    ,charge_sequence  = cast(null as varchar(10))
    ,external_id      = cast(pol.entity_id + pol.po_num + pol.po_line + pol.po_seq as varchar(100))
    ,external_id_len  = len(cast(pol.entity_id + pol.po_num + pol.po_line + pol.po_seq as varchar(100)))
    -- the proc uses cast(), not try_cast: a NULL here means the proc would THROW
    ,line_number      = try_cast(pol.po_line + pol.po_seq as int)
    ,expense_type     = cast('' as varchar(64))
    ,account_code     = cast(isnull(edac.account_code,'0005050') as varchar(64))
    ,[description]    = cast(replace(pol.po_line_desc,'"','""') as varchar(255))
    ,quantity         = cast(pol.qty_ordered as decimal(12,2))
    ,unit_price       = cast(pol.unit_price as decimal(23,8))
into #line
from dbo.vw_concur_extract_po_line pol
    inner join dbo.vw_concur_extract_po_header poh (nolock)
        on  poh.system_id        = pol.system_id
        and poh.system_entity_id = pol.system_entity_id
        and poh.po_num           = pol.po_num
        and coalesce(poh.status_code,1) = 1
    left outer join dbo.entity_department_account_code edac
        on edac.entity_id = pol.entity_id and edac.gl_account_code = poh.ledger_code
where (@all = 1 or pol.po_num in (select po_num from #po))

union all

select
     leg              = cast('2-charge' as varchar(8))
    ,polc.system_id, polc.system_entity_id, polc.entity_id, polc.po_num, polc.po_line, polc.po_seq
    ,charge_sequence  = cast(polc.charge_sequence as varchar(10))
    ,external_id      = cast(polc.entity_id + polc.po_num + polc.po_line + polc.po_seq as varchar(100))
    ,external_id_len  = len(cast(polc.entity_id + polc.po_num + polc.po_line + polc.po_seq as varchar(100)))
    ,line_number      = try_cast('99' + polc.po_line + '99' + polc.charge_sequence as int)
    ,expense_type     = cast('' as varchar(64))
    ,account_code     = cast(isnull(case when polc.short_description like '%TAX%'     then '5555'
                                         when polc.short_description like '%OTHER%'   then '4444'
                                         when polc.short_description like '%FREIGHT%' then '9999'
                                    end,'') as varchar(64))
    ,[description]    = cast(replace(pol.po_line_desc,'"','""') as varchar(255))  -- proc reads the LINE desc
    ,quantity         = cast(1 as decimal(12,2))
    ,unit_price       = cast(polc.charge_amount as decimal(23,8))
from dbo.vw_concur_extract_po_line_charges polc
    inner join dbo.vw_concur_extract_po_header poh (nolock)
        on  poh.system_id        = polc.system_id
        and poh.system_entity_id = polc.system_entity_id
        and poh.po_num           = polc.po_num
        and coalesce(poh.status_code,1) = 1
    left outer join dbo.vw_concur_extract_po_line pol
        on  pol.system_id        = polc.system_id
        and pol.system_entity_id = polc.system_entity_id
        and pol.po_num           = polc.po_num
        and pol.po_line          = polc.po_line
        and pol.po_seq           = polc.po_seq
where polc.delete_flag = 'N'
  and (@all = 1 or polc.po_num in (select po_num from #po))
  and exists (
        select 1
        from dbo.vw_concur_extract_po_vendor ven
        where ven.system_id        = poh.system_id
          and ven.system_entity_id = poh.system_entity_id
          and ven.vend_num         = poh.vend_num
          and ven.vend_seq         = poh.vend_seq
          and nullif(ltrim(rtrim(ven.vendor_code)),'') is not null
      );


---------------------------------------------------------------------------------- receipt records
if object_id('tempdb..#rcpt') is not null drop table #rcpt;

select
     polr.system_id, polr.system_entity_id, polr.entity_id, polr.po_num, polr.po_line, polr.po_seq
    ,purchase_order_number  = polr.entity_id + polr.po_num
    ,line_item_external_id  = cast(polr.entity_id + polr.po_num + polr.po_line + polr.po_seq as varchar(100))
    ,line_item_len          = len(cast(polr.entity_id + polr.po_num + polr.po_line + polr.po_seq as varchar(100)))
    ,goods_receipt_number   = coalesce(polr.external_id,
                                polr.system_entity_id + polr.po_num + isnull(polr.po_seq,'')
                                + cast(row_number() over (partition by polr.system_id, polr.system_entity_id,
                                                                       polr.po_num, polr.po_seq
                                                          order by polr.po_num, polr.po_seq desc) as varchar))
    ,gr_number_is_fallback  = case when polr.external_id is null then 1 else 0 end  -- fallback is not stable run to run
    ,delivery_slip_number   = cast(polr.packing_slip as varchar(255))
    ,uom_code               = cast(polr.unit_of_measure_order as varchar(10))
    ,received_quantity      = polr.qty_received
    ,received_date          = convert(char(10), polr.date_received, 23)
into #rcpt
from dbo.vw_concur_extract_po_line_receipt polr
-- mirrors p_concurinvoicepo_get_poreceipt after the 2026-08-26 rewrite: the poh and pol INNER JOINs
-- became EXISTS (semi-joins), so neither can duplicate receipt rows, and the vendor test folded in.
where (@all = 1 or polr.po_num in (select po_num from #po))
  and exists (
        select 1
        from dbo.vw_concur_extract_po_header poh
            inner join dbo.vw_concur_extract_po_vendor ven
                on  ven.system_id        = poh.system_id
                and ven.system_entity_id = poh.system_entity_id
                and ven.vend_num         = poh.vend_num
                and ven.vend_seq         = poh.vend_seq
                and ven.vendor_code is not null
        where poh.system_id        = polr.system_id
          and poh.system_entity_id = polr.system_entity_id
          and poh.po_num           = polr.po_num
          and (poh.status_code = 1 or poh.status_code is null)
      )
  and exists (
        select 1
        from dbo.vw_concur_extract_po_line pol
        where pol.system_id        = polr.system_id
          and pol.system_entity_id = polr.system_entity_id
          and pol.po_num           = polr.po_num
          and pol.po_line          = polr.po_line
          and pol.po_seq           = polr.po_seq
      );


/*===============================================================================================
  2.  HEADER DECISION MATRIX - which filter keeps or kills each PO
===============================================================================================*/
select
     [#] = '2. header decision matrix'
    ,h.entity_id, h.po_num, h.purchase_order_number, h.system_id, h.system_entity_id
    ,h.status_code, h.vendor_code, h.vendor_address_code
    ,[status=1]              = case when h.f1_status_ok=1            then 'pass' else 'FAIL' end
    ,[vendor_code not null]  = case when h.f2_vendor_hdr_ok=1        then 'pass' else 'FAIL' end
    ,[vendor_code non-blank] = case when h.f2b_vendor_charge_ok=1    then 'pass' else 'FAIL' end
    ,[has line w/ receipt]   = case when h.f3_line_with_receipt_ok=1 then 'pass' else 'FAIL' end
    ,h.line_count, h.charge_count, h.receipt_count
    ,[HEADER SHIPS?]         = case when h.ships_header=1 then 'YES' else 'no' end
from #hdr h
order by h.entity_id, h.vend_num, h.po_num;


/*===============================================================================================
  3.  WHAT EACH PROC WILL EMIT
===============================================================================================*/
select [#] = '3a. 200 header records'
    ,h.purchase_order_number, h.entity_id, h.po_num, h.vendor_code, h.vendor_address_code, h.status_code
from #hdr h
where h.ships_header = 1
order by h.entity_id, h.vend_num, h.po_num;

select [#] = '3b. 300 line records'
    ,l.leg, l.po_num, l.po_line, l.po_seq, l.charge_sequence, l.line_number
    ,l.external_id, l.external_id_len, l.expense_type, l.account_code
    ,[field 7/8 rule] = case
        when nullif(l.expense_type,'') is null     and nullif(l.account_code,'') is null     then 'VIOLATION - both empty (Error 5001)'
        when nullif(l.expense_type,'') is not null and nullif(l.account_code,'') is not null then 'VIOLATION - both populated'
        else 'ok' end
    ,l.quantity, l.unit_price, l.[description]
from #line l
order by l.po_num, l.line_number, l.leg;

select [#] = '3c. receipt records'
    ,r.purchase_order_number, r.line_item_external_id, r.line_item_len
    ,r.goods_receipt_number, r.gr_number_is_fallback
    ,r.received_quantity, r.received_date, r.uom_code, r.delivery_slip_number
from #rcpt r
order by r.po_num desc, try_cast(r.po_seq as int), try_cast(r.po_line as int);


/*===============================================================================================
  4.  SYNC CHECKS - every row returned below is a defect.  No rows = the three procs agree.
===============================================================================================*/

------------------------------------------------------------------ 4A. receipts for a suppressed PO
select
     [#] = '4A. RECEIPT SHIPS BUT HEADER DOES NOT  (regression test - expect 0 rows)'
    ,r.entity_id, r.po_num
    ,receipt_rows = count(*)
    ,why = isnull(stuff(
              case when max(h.f1_status_ok)            = 0 then ' + status_code <> 1'       else '' end
            + case when max(h.f2_vendor_hdr_ok)        = 0 then ' + vendor_code is null'    else '' end
            + case when max(h.f3_line_with_receipt_ok) = 0 then ' + no line with a receipt' else '' end
            ,1,3,''),'no matching header row at all')
from #rcpt r
    left outer join #hdr h
        on  h.system_id        = r.system_id
        and h.system_entity_id = r.system_entity_id
        and h.po_num           = r.po_num
where isnull(h.ships_header,0) = 0
group by r.entity_id, r.po_num;

--------------------------------------------------------------------- 4B. lines for a suppressed PO
select
     [#] = '4B. LINE SHIPS BUT HEADER DOES NOT'
    ,l.leg, l.entity_id, l.po_num
    ,line_rows = count(*)
    ,why = isnull(stuff(
              case when max(h.f1_status_ok)            = 0 then ' + status_code <> 1'       else '' end
            + case when max(h.f2_vendor_hdr_ok)        = 0 then ' + vendor_code is null'    else '' end
            + case when max(h.f3_line_with_receipt_ok) = 0 then ' + no line with a receipt' else '' end
            ,1,3,''),'no matching header row at all')
from #line l
    left outer join #hdr h
        on  h.system_id        = l.system_id
        and h.system_entity_id = l.system_entity_id
        and h.po_num           = l.po_num
where isnull(h.ships_header,0) = 0
group by l.leg, l.entity_id, l.po_num;

------------------------------------------------- 4C. header ships but the line proc returns nothing
select
     [#] = '4C. HEADER SHIPS WITH NO 300 RECORDS'
    ,h.entity_id, h.po_num, h.line_count, h.charge_count
    ,note = 'header qualified on entity_id+po_num, line proc keys on system_id/system_entity_id'
from #hdr h
where h.ships_header = 1
  and not exists (select 1 from #line l
                   where l.system_id = h.system_id
                     and l.system_entity_id = h.system_entity_id
                     and l.po_num = h.po_num);

---------------------------------------------- 4D. blank vendor_code: real lines ship, charges do not
select
     [#] = '4D. CHARGE LINES DROPPED WHILE HEADER + LINES SHIP (blank but not null vendor_code)'
    ,h.entity_id, h.po_num, h.vendor_code, h.charge_count
from #hdr h
where h.ships_header = 1
  and h.f2b_vendor_charge_ok = 0
  and h.charge_count > 0;

------------------------------------------------------ 4E. external_id collision between the two legs
select
     [#] = '4E. EXTERNAL ID SHARED BY MORE THAN ONE 300 RECORD'
    ,l.po_num, l.external_id
    ,legs           = count(distinct l.leg)
    ,records_sent   = count(*)
    ,min_line_number = min(l.line_number)
    ,max_line_number = max(l.line_number)
from #line l
group by l.po_num, l.external_id
having count(*) > 1;

------------------------------------------------ 4F. receipt points at a line key that will not ship
select
     [#] = '4F. RECEIPT line_item_external_id HAS NO MATCHING 300 RECORD (Error 1001)'
    ,r.entity_id, r.po_num, r.po_line, r.po_seq
    ,r.line_item_external_id, r.line_item_len
    ,sample_line_key = (select top 1 l.external_id from #line l
                         where l.po_num = r.po_num and l.leg = '1-line'
                         order by l.external_id)
    ,note = case when r.line_item_external_id is null
                 then 'NULL key - po_seq is null and the receipt proc has no isnull() on it'
                 else 'no 300 record carries this external_id' end
from #rcpt r
where not exists (select 1 from #line l
                   where l.leg = '1-line' and l.external_id = r.line_item_external_id);

--------------------------------------------------------- 4G. key length / padding across the sources
select [#] = '4G. KEY LENGTHS BY SOURCE (differing lengths = CHAR padding)'
      ,src = 'line', po_num, key_len = external_id_len, [rows] = count(*)
from #line group by po_num, external_id_len
union all
select [#] = '4G. KEY LENGTHS BY SOURCE (differing lengths = CHAR padding)'
      ,src = 'receipt', po_num, key_len = line_item_len, [rows] = count(*)
from #rcpt group by po_num, line_item_len
order by po_num, src;

------------------------------------------------------------------------ 4H. duplicate header records
select [#] = '4H. DUPLICATE 200 HEADER RECORDS', h.purchase_order_number, header_rows = count(*)
from #hdr h
where h.ships_header = 1
group by h.purchase_order_number
having count(*) > 1;

------------------------------------------------------------- 4I. line_number would throw in the proc
select [#] = '4I. line_number WOULD THROW ON CAST TO INT'
      ,l.leg, l.po_num, l.po_line, l.po_seq, l.charge_sequence
from #line l
where l.line_number is null;


------------------------------------------ 4J. would the OLD inner-join form have duplicated receipts?
select
     [#] = '4J. JOIN FAN-OUT the EXISTS rewrite removed (informational)'
    ,r.entity_id, r.po_num, r.po_line, r.po_seq
    ,header_matches = (select count(*) from dbo.vw_concur_extract_po_header poh
                        where poh.system_id=r.system_id and poh.system_entity_id=r.system_entity_id
                          and poh.po_num=r.po_num)
    ,line_matches   = (select count(*) from dbo.vw_concur_extract_po_line pol
                        where pol.system_id=r.system_id and pol.system_entity_id=r.system_entity_id
                          and pol.po_num=r.po_num and pol.po_line=r.po_line and pol.po_seq=r.po_seq)
from #rcpt r
where (select count(*) from dbo.vw_concur_extract_po_header poh
        where poh.system_id=r.system_id and poh.system_entity_id=r.system_entity_id
          and poh.po_num=r.po_num) > 1
   or (select count(*) from dbo.vw_concur_extract_po_line pol
        where pol.system_id=r.system_id and pol.system_entity_id=r.system_entity_id
          and pol.po_num=r.po_num and pol.po_line=r.po_line and pol.po_seq=r.po_seq) > 1;


/*===============================================================================================
  5.  VERDICT - one row per PO
===============================================================================================*/
select
     [#] = '5. verdict'
    ,po.po_num
    ,header_rows  = (select count(*) from #hdr  h where h.po_num = po.po_num and h.ships_header = 1)
    ,line_rows    = (select count(*) from #line l where l.po_num = po.po_num and l.leg = '1-line')
    ,charge_rows  = (select count(*) from #line l where l.po_num = po.po_num and l.leg = '2-charge')
    ,receipt_rows = (select count(*) from #rcpt r where r.po_num = po.po_num)
    ,verdict = case
        when (select count(*) from #hdr h where h.po_num = po.po_num and h.ships_header = 1) = 0
             and ((select count(*) from #line l where l.po_num = po.po_num) > 0
               or (select count(*) from #rcpt r where r.po_num = po.po_num) > 0)
            then 'OUT OF SYNC - header suppressed, lines/receipts still ship'
        when (select count(*) from #hdr h where h.po_num = po.po_num and h.ships_header = 1) = 0
            then 'excluded everywhere - consistent'
        when (select count(*) from #line l where l.po_num = po.po_num and l.leg = '1-line') = 0
            then 'OUT OF SYNC - header ships with no 300 records'
        when exists (select 1 from #rcpt r
                      where r.po_num = po.po_num
                        and not exists (select 1 from #line l
                                         where l.leg = '1-line'
                                           and l.external_id = r.line_item_external_id))
            then 'OUT OF SYNC - receipt references a line that will not ship'
        when exists (select 1 from #line l where l.po_num = po.po_num
                      group by l.external_id having count(*) > 1)
            then 'OUT OF SYNC - duplicate external_id across the 300 records'
        else 'in sync' end
from (select po_num from #hdr
      union select po_num from #line
      union select po_num from #rcpt) po
order by po.po_num;


/*===============================================================================================
  6.  OPTIONAL - run the real procs and eyeball the actual output.
      The header and receipt procs take no parameters, so they return EVERY PO.
===============================================================================================*/
-- exec dbo.p_concurinvoicepo_get_poheader;
-- exec dbo.p_concurinvoicepo_get_poline  @po_num = '7806040';
-- exec dbo.p_concurinvoicepo_get_poline  @po_num = '1100000006';
-- exec dbo.p_concurinvoicepo_get_poreceipt;
