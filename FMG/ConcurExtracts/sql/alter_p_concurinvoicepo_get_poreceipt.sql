/*=====================================================================================================
  p_concurinvoicepo_get_poreceipt - vendor filter + join tuning
  -----------------------------------------------------------------------------------------------
  WHAT CHANGES
    1. Vendor filter added, so the receipt file can no longer ship a PO the header dropped.
       (p_concurinvoicepo_get_poheader: "where ... and ven.vendor_code is not null")
    2. The poh and pol INNER JOINs become EXISTS.  Neither contributed a column to the SELECT list -
       both were existence tests written as joins, and either could multiply receipt rows if its view
       returns more than one row per key.  EXISTS is a semi-join: it stops at the first match and
       cannot fan out.  poh's vend_num / vend_seq are still needed, so the header and vendor tests
       fold into one EXISTS - "a qualifying header with a vendor code exists".
    3. coalesce(poh.status_code,1)=1  ->  (poh.status_code = 1 or poh.status_code is null).
       Same meaning, but SARGable - the original wraps the column in a function so it can never seek.

  WHAT DOES NOT CHANGE
    The SELECT list.  It IS the 29-field file layout (streamed by reader.GetValue(i) in proc column
    order), so nothing there may be added, removed or reordered.
    The ORDER BY.  Its casts prevent any index from satisfying the sort, but changing the file's
    record order is a separate decision - left alone deliberately.
    The ROW_NUMBER fallback for goods_receipt_number.  See variant B at the bottom.

  OUTPUT IMPACT
    Identical to today's output EXCEPT: rows for POs with no vendor_code disappear (that is the fix),
    and if either join was fanning out, the duplicate receipt records disappear too.  Run the three
    pre-deploy checks below first - they quantify all of it.
=====================================================================================================*/


/*-----------------------------------------------------------------------------------------------
  PRE-DEPLOY CHECK 1 - which receipt rows does the vendor filter remove?
  These are receipts that have been shipping for POs Concur never received a 200 header for.
  Nothing returned = the change is insurance rather than a fix.
-----------------------------------------------------------------------------------------------*/
/*
select
     polr.entity_id
    ,polr.po_num
    ,purchase_order_number = polr.entity_id + polr.po_num
    ,poh.vend_num
    ,poh.vend_seq
    ,receipt_rows_dropped  = count(*)
from App_Concur.dbo.vw_concur_extract_po_line_receipt polr
    inner join App_Concur.dbo.vw_concur_extract_po_header poh
        on  poh.system_id        = polr.system_id
        and poh.system_entity_id = polr.system_entity_id
        and poh.po_num           = polr.po_num
        and coalesce(poh.status_code,1) = 1
    inner join App_Concur.dbo.vw_concur_extract_po_line pol
        on  pol.system_id        = polr.system_id
        and pol.system_entity_id = polr.system_entity_id
        and pol.po_num           = polr.po_num
        and pol.po_line          = polr.po_line
        and pol.po_seq           = polr.po_seq
where not exists (
        select 1
        from App_Concur.dbo.vw_concur_extract_po_vendor ven
        where ven.system_id        = poh.system_id
          and ven.system_entity_id = poh.system_entity_id
          and ven.vend_num         = poh.vend_num
          and ven.vend_seq         = poh.vend_seq
          and ven.vendor_code is not null
    )
group by polr.entity_id, polr.po_num, poh.vend_num, poh.vend_seq
order by receipt_rows_dropped desc;
*/


/*-----------------------------------------------------------------------------------------------
  PRE-DEPLOY CHECK 2 - are the two INNER JOINs fanning out TODAY?
  rows_after > rows_before means the current proc is emitting duplicate receipt records, and the
  EXISTS rewrite will remove them.  It also means today's ROW_NUMBER fallback values are computed
  over inflated partitions and will shift - see check 3.
-----------------------------------------------------------------------------------------------*/
/*
select
     rows_before = (select count(*) from App_Concur.dbo.vw_concur_extract_po_line_receipt)
    ,rows_after  = (
        select count(*)
        from App_Concur.dbo.vw_concur_extract_po_line_receipt polr
            inner join App_Concur.dbo.vw_concur_extract_po_header poh
                on  poh.system_id        = polr.system_id
                and poh.system_entity_id = polr.system_entity_id
                and poh.po_num           = polr.po_num
            inner join App_Concur.dbo.vw_concur_extract_po_line pol
                on  pol.system_id        = polr.system_id
                and pol.system_entity_id = polr.system_entity_id
                and pol.po_num           = polr.po_num
                and pol.po_line          = polr.po_line
                and pol.po_seq           = polr.po_seq);

-- and if they differ, which keys duplicate:
select top 50
     polr.po_num, polr.po_line, polr.po_seq
    ,header_matches = (select count(*) from App_Concur.dbo.vw_concur_extract_po_header poh
                        where poh.system_id=polr.system_id and poh.system_entity_id=polr.system_entity_id
                          and poh.po_num=polr.po_num)
    ,line_matches   = (select count(*) from App_Concur.dbo.vw_concur_extract_po_line pol
                        where pol.system_id=polr.system_id and pol.system_entity_id=polr.system_entity_id
                          and pol.po_num=polr.po_num and pol.po_line=polr.po_line and pol.po_seq=polr.po_seq)
from App_Concur.dbo.vw_concur_extract_po_line_receipt polr
where (select count(*) from App_Concur.dbo.vw_concur_extract_po_header poh
        where poh.system_id=polr.system_id and poh.system_entity_id=polr.system_entity_id
          and poh.po_num=polr.po_num) > 1
   or (select count(*) from App_Concur.dbo.vw_concur_extract_po_line pol
        where pol.system_id=polr.system_id and pol.system_entity_id=polr.system_entity_id
          and pol.po_num=polr.po_num and pol.po_line=polr.po_line and pol.po_seq=polr.po_seq) > 1;
*/


/*-----------------------------------------------------------------------------------------------
  PRE-DEPLOY CHECK 3 - how many receipts actually use the ROW_NUMBER fallback?
  fallback_rows = 0 means the window function is pure overhead: it is computed for every row and
  discarded for every row, and it forces a sort.  In that case variant B is free.
-----------------------------------------------------------------------------------------------*/
/*
select
     total_rows    = count(*)
    ,has_ext_id    = sum(case when polr.external_id is not null then 1 else 0 end)
    ,fallback_rows = sum(case when polr.external_id is null     then 1 else 0 end)
from App_Concur.dbo.vw_concur_extract_po_line_receipt polr;
*/


/*=====================================================================================================
  VARIANT A - deploy this.  Output-neutral apart from the fix itself.
=====================================================================================================*/
USE [App_Concur]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER procedure [dbo].[p_concurinvoicepo_get_poreceipt]  as
/*
	declare @extract_id int = 21
--*/
begin
select
	200 record_type
	,polr.entity_id + polr.po_num [purchase_order_number]
	,cast(polr.entity_id + polr.po_num + polr.po_line + polr.po_seq as varchar(100))[line_item_external_id] --polr.item_num_internal [line_item_external_id]
	,coalesce(polr.external_id, polr.system_entity_id+polr.po_num+isnull(polr.po_seq,'') + cast(ROW_NUMBER () OVER(PARTITION by polr.system_id,polr.system_entity_id,polr.po_num,polr.po_seq order by polr.po_num,polr.po_seq desc) as varchar)) [goods_receipt_number]
	,cast(polr.packing_slip as varchar(255)) [delivery_slip_number]
	,cast(polr.unit_of_measure_order as varchar(10)) [uom_code]
	,polr.qty_received [received_quantity]
	,convert(char(10), polr.date_received, 23) [received_date]
	,'' [is_deleted]
	,cast(null as varchar(50))future_use_1
	,cast(null as varchar(50))future_use_2
	,cast(null as varchar(50))future_use_3
	,cast(null as varchar(50))future_use_4
	,cast(null as varchar(50))future_use_5
	,cast(null as varchar(50))future_use_6
	,cast(null as varchar(50))future_use_7
	,cast(null as varchar(50))future_use_8
	,cast(null as varchar(50))future_use_9
	,cast(null as varchar(50))future_use_10
	,cast(null as varchar(50))custom_1
	,cast(null as varchar(50))custom_2
	,cast(null as varchar(50))custom_3
	,cast(null as varchar(50))custom_4
	,cast(null as varchar(50))custom_5
	,cast(null as varchar(50))custom_6
	,cast(null as varchar(50))custom_7
	,cast(null as varchar(50))custom_8
	,cast(null as varchar(50))custom_9
	,cast(null as varchar(50))custom_10
from App_Concur.dbo.vw_concur_extract_po_line_receipt polr
-- A qualifying header WITH a vendor code must exist.  Was two INNER JOINs (poh, then the vendor test
-- in the WHERE); poh contributed nothing to the SELECT list, only vend_num / vend_seq to the vendor
-- lookup, so both collapse into one semi-join that cannot duplicate receipt rows.
where exists (
		select 1
		from App_Concur.dbo.vw_concur_extract_po_header poh
			inner join App_Concur.dbo.vw_concur_extract_po_vendor ven
				on  ven.system_id        = poh.system_id
				and ven.system_entity_id = poh.system_entity_id
				and ven.vend_num         = poh.vend_num
				and ven.vend_seq         = poh.vend_seq
				and ven.vendor_code is not null
				--  strict variant - only if p_concurinvoicepo_get_poheader is changed to match:
				--  and nullif(ltrim(rtrim(ven.vendor_code)),'') is not null
		where poh.system_id        = polr.system_id
			and poh.system_entity_id = polr.system_entity_id
			and poh.po_num           = polr.po_num
			and (poh.status_code = 1 or poh.status_code is null)   -- was coalesce(poh.status_code,1)=1
	)
-- The receipt's PO line must exist.  Was an INNER JOIN to vw_concur_extract_po_line that selected
-- no columns at all - a pure existence test, and a duplication risk if that view is not unique on
-- system_id + system_entity_id + po_num + po_line + po_seq.
	and exists (
		select 1
		from App_Concur.dbo.vw_concur_extract_po_line pol
		where pol.system_id        = polr.system_id
			and pol.system_entity_id = polr.system_entity_id
			and pol.po_num           = polr.po_num
			and pol.po_line          = polr.po_line
			and pol.po_seq           = polr.po_seq
	)
order by polr.po_num desc, cast(polr.po_seq as int), cast(polr.po_line as int)
end
GO


/*=====================================================================================================
  VARIANT B - the goods_receipt_number fallback.  DO NOT deploy without reading this.

  Today:  coalesce(polr.external_id,
                   system_entity_id + po_num + isnull(po_seq,'') + ROW_NUMBER() OVER (
                       PARTITION BY system_id, system_entity_id, po_num, po_seq
                       ORDER BY     po_num, po_seq desc))

  Two problems.  The window function is evaluated for EVERY row and discarded wherever external_id is
  populated - coalesce does not short-circuit an operator - and it forces a sort of the whole set.
  Worse, the ORDER BY inside the OVER uses the partitioning columns themselves, so every row in a
  partition ties and SQL Server may number them in any order it likes.  The value is not stable
  between runs, and goods_receipt_number is how Concur identifies a goods receipt: a number that
  shifts makes Concur create a SECOND receipt against the PO line rather than update the first.

  So the current fallback is already unsafe.  Any change here re-keys receipts Concur has already
  accepted - they load as new goods receipts, not updates.  Sequence:
    1. Run PRE-DEPLOY CHECK 3.  If fallback_rows = 0 the fallback has never fired, nothing in Concur
       is keyed on it, and B is free - deploy it and delete the window function.
    2. If fallback_rows > 0, work out why external_id is null in vw_concur_extract_po_line_receipt
       and fix it at the source.  A real external_id beats any synthesised key.
    3. Only if neither is possible, use the deterministic form below - and expect Concur to create
       duplicate goods receipts for the rows whose number changes.

  Replacement expression (adds po_line to the key, orders on columns that actually vary):
-----------------------------------------------------------------------------------------------*/
/*
	,coalesce(polr.external_id
		,polr.system_entity_id + polr.po_num + polr.po_line + isnull(polr.po_seq,'')
		 + cast(ROW_NUMBER() OVER (
				PARTITION BY polr.system_id, polr.system_entity_id, polr.po_num, polr.po_line, polr.po_seq
				ORDER BY     polr.date_received, polr.packing_slip) as varchar)
		) [goods_receipt_number]

	-- residual risk: two receipts against the same line on the same date with the same packing slip
	-- still tie, and a receipt inserted with an earlier date renumbers the ones after it.

	-- if PRE-DEPLOY CHECK 3 returns fallback_rows = 0, use this instead and drop the window function
	-- entirely - one less sort over the whole result set:
	,polr.external_id [goods_receipt_number]
*/
