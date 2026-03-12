select
item.item_number as 'Item Number',
sites.site_name as 'Site',
location.code as 'Location',
ISNULL(inventory_transactions.serial_number,'') as 'Serial Number',
ISNULL(inventory_transactions.lot,'') as 'Lot',
ISNULL(inventory_transactions.date_code,'') as 'Date Code',
ISNULL(inventory_transactions.pallet, '') as 'Pallet',
ISNULL(inventory_transactions.po,'') as 'Ref Number',
ISNULL(supplier.supplier_number, '') as 'Vendor Number',
ISNULL(customer.customer_number,'') as 'Customer',
ISNULL(convert(varchar, inventory_transactions.date_acquired, 23),'') as 'Date Acquired',
inventory_transactions.remaining_qty as 'Quantity',
inventory_transactions.cost as 'Unit Cost'
from dbo.inventory_transactions
inner join dbo.item
on inventory_transactions.item_id = item.item_id
inner join dbo.location
on inventory_transactions.location_id = location.location_id
inner join dbo.sites
on location.site_id = sites.site_id
left outer join dbo.supplier
on inventory_transactions.supplier_id = supplier.supplier_id
left outer join dbo.customer
on inventory_transactions.customer_id = customer.customer_id
where inventory_transactions.record_status = 1 and trans_type not in (350)
order by item.item_number, inventory_transactions.trans_date