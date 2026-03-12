select distinct
category.description as 'Category Name'
from dbo.category
where category.record_status = 1