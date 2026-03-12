select
item.item_number as 'Item Number',
sites.site_name as 'Site',
location.code as 'Location',
ISNULL(inventory_transactions.serial_number,'') as 'Serial Number',
ISNULL(inventory_transactions.lot,'') as 'Lot',
ISNULL(inventory_transactions.date_code,'') as 'Date Code',
CASE
	WHEN inventory_transactions.date_acquired = '1800-01-01' THEN ''
	ELSE ISNULL(convert(varchar, inventory_transactions.date_acquired, 23),'')
END 'Due Date',
inventory_transactions.remaining_qty as 'Quantity',
ISNULL(supplier.supplier_number, '') as 'Vendor Number',
ISNULL(customer.customer_number,'') as 'Customer Number',
ISNULL(REPLACE(REPLACE(SUBSTRING(
        (
            SELECT ' '+notes.note_text  AS [text()]
            FROM dbo.notes
            WHERE notes.note_id = inventory_transactions.transaction_id and notes.table_name = 'inventory_transactions' and notes.note_text is not null and notes.note_text <> ''
            ORDER BY notes.note_date
            FOR XML PATH ('')
        ), 2, 1000), '&#x0D;', '    '), CHAR(10), '    '), '') as 'Notes'
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
where inventory_transactions.record_status = 1 and trans_type = 350
order by item.item_number, inventory_transactions.trans_date