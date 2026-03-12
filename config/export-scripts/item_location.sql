select 
item.item_number as 'Item Number',
sites.site_name as 'Site',
location.code as 'Location',
CASE WHEN inventory.min_stock_level < 0 THEN 0 ELSE inventory.min_stock_level END 'Min Stock Level',
CASE WHEN inventory.max_stock_level < 0 THEN 0 ELSE inventory.max_stock_level END 'Max Stock Level',
CASE WHEN inventory.reorder_qty < 0 THEN 0 ELSE inventory.reorder_qty END 'Reorder Qty',
inventory.primary_location as 'Primary Location'
from dbo.inventory
inner join dbo.location
on inventory.location_id = location.location_id
inner join dbo.sites
on location.site_id = sites.site_id
inner join dbo.item
on inventory.item_id = item.item_id
where inventory.record_status = 1 and item.record_status = 1 and location.record_status = 1 and sites.record_status = 1