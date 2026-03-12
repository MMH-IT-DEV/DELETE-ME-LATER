select
item.item_number as 'Item Number',
supplier.supplier_number 'Vendor Number',
item_supplier.supplier_item_sku as 'Vendor SKU',
ISNULL(item_supplier.supplier_item_description, '') as 'Vendor Item Description',
item_supplier.lead_time as 'Lead Time',
item_supplier.default_supplier as 'Preferred Vendor',
unit_of_measure.uom_code 'Unit',
unit_of_measure.preferred_uom 'Default Unit',
unit_of_measure.cost 'Price',
ISNULL(item_supplier.note, '') as 'Notes'
from dbo.item_supplier
inner join dbo.item
on item.item_id = item_supplier.item_id
inner join dbo.supplier
on supplier.supplier_id = item_supplier.supplier_id
inner join dbo.unit_of_measure
on item_supplier.supplier_id = unit_of_measure.supplier_id and item_supplier.item_id = unit_of_measure.item_id
where item_supplier.record_status = 1 and item.item_type not in (4,5)
Order by item_number asc, quantity_eaches asc