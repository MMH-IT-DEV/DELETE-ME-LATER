select 
[sites].site_name as 'Site'
, item_number as 'Item Number'
, min_stock_level as 'Min Stock Level'
, max_stock_level as 'Max Stock Level'
, reorder_qty as 'Reorder Qty'
from [dbo].[item]
left outer join [dbo].[sites] on 1=1
where sites.record_status = 1 and item.record_status = 1 and sites.site_id > 0 and item.item_id > 0 
and (min_stock_level > 0 or max_stock_level > 0 or reorder_qty > 0)