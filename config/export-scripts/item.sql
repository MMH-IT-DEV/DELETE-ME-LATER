select 
item_type.item_type_name as 'Type',
item.item_number as 'Item Number',
ISNULL(item.alt_item_number, '') as 'Alt Item Number',
ISNULL(item.description, '') as 'Description',
ISNULL(sales_tax_code.tax_code, '') as 'Tax Code',
item.cost as 'Cost',
item_costing_method.method_name as 'Cost Method',
ISNULL(manufacturer.name, '') as 'Manufacturer',
ISNULL(category.description, '') as 'Category',
CASE 
	WHEN ISNULL(item.unit_of_measure, '') = '' THEN 'ea'
	WHEN item.unit_of_measure = '' THEN 'ea'
	ELSE item.unit_of_measure
END	as 'Stocking Unit',
CASE 
	WHEN ISNULL(item.unit_of_measure, '') = '' THEN 'ea'
	WHEN item.unit_of_measure = '' THEN 'ea'
	ELSE item.unit_of_measure
END	as 'Purchase Unit',
CASE 
	WHEN ISNULL(item.unit_of_measure, '') = '' THEN 'ea'
	WHEN item.unit_of_measure = '' THEN 'ea'
	ELSE item.unit_of_measure
END	as 'Sales Unit',
item.list_price as 'List Price',
item.sale_price as 'Sales Price',
item.track_serial_numbers as 'Serial Number',
item.auto_sn_value as 'Serial Number Format',
item.track_lots as 'Lot',
item.track_date_codes as 'Date Code',
item.track_pos as 'Ref Number',
item.track_suppliers as 'Vendor',
item.track_customers as 'Customer',
item.checkout_duration as 'Checkout Length',
item.width as 'Width',
item.height as 'Height',
item.length as 'Depth',
item.dimension_unit as 'Length Unit',
item.weight as 'Weight',
item.weight_unit as 'Weight Unit',
item.auto_sn as 'Auto Generate Serial Number',
item.use_subitem_cost as 'Use Sub Item Cost',
ISNULL(REPLACE(REPLACE(SUBSTRING(
        (
            SELECT ' '+notes.note_text  AS [text()]
            FROM dbo.notes
            WHERE notes.note_id = item.item_id and notes.table_name = 'item' and notes.note_text is not null and notes.note_text <> ''
            ORDER BY notes.note_date
            FOR XML PATH ('')
        ), 2, 1000), '&#x0D;', '    '), CHAR(10), '    '), '') as 'Notes',
ISNULL(convert(varchar, item.usrdefdate1, 23),'') as 'Custom Date 1',
ISNULL(convert(varchar, item.usrdefdate2, 23),'') as 'Custom Date 2',
ISNULL(convert(varchar, item.usrdefdate3, 23),'') as 'Custom Date 3',
ISNULL(convert(varchar, item.usrdefdate4, 23),'') as 'Custom Date 4',
ISNULL(convert(varchar, item.usrdefdate5, 23),'') as 'Custom Date 5',
ISNULL(item.usrdefnumber1,0) as 'Custom Number 1',
ISNULL(item.usrdefnumber2,0) as 'Custom Number 2',
ISNULL(item.usrdefnumber3,0) as 'Custom Number 3',
ISNULL(item.usrdefnumber4,0) as 'Custom Number 4',
ISNULL(item.usrdefnumber5,0) as 'Custom Number 5',
ISNULL(item.usrdeftext1,'') as 'Custom Text 1',
ISNULL(item.usrdeftext2,'') as 'Custom Text 2',
ISNULL(item.usrdeftext3,'') as 'Custom Text 3',
ISNULL(item.usrdeftext4,'') as 'Custom Text 4',
ISNULL(item.usrdeftext5,'') as 'Custom Text 5',
ISNULL(item.usrdeftext6,'') as 'Custom Text 6',
ISNULL(item.usrdeftext7,'') as 'Custom Text 7',
ISNULL(item.usrdeftext8,'') as 'Custom Text 8',
ISNULL(item.usrdeftext9,'') as 'Custom Text 9',
ISNULL(item.usrdeftext10,'') as 'Custom Text 10'
from dbo.item
inner join dbo.item_type
on item.item_type = item_type.item_type_id
left outer join dbo.item_ext
on item.item_id = item_ext.item_id
left outer join dbo.sales_tax_code
on item_ext.sales_tax_code_id = sales_tax_code.tax_code_id
left outer join dbo.item_costing_method
on item.cost_calc_method = item_costing_method.costing_method_id
left outer join dbo.manufacturer
on item.manufacturer_id = manufacturer.manufacturer_id
left outer join dbo.category
on item.category_id = category.category_id
where item.record_status = 1
and item.item_id > 0