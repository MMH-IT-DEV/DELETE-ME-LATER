select
item.item_number as 'Item Number',
sub_item.item_number as 'Sub Item Number',
item_group_spec.quantity as 'Quantity'
from dbo.item_group_spec
inner join dbo.item
on item_group_spec.item_id = item.item_id
inner join dbo.item sub_item
on item_group_spec.sub_item_id = sub_item.item_id
where item_group_spec.record_status = 1
order by item.item_number