select distinct
unit_of_measure.uom_code as 'Unit Name',
unit_of_measure.uom_code as 'Abbreviation',
'Simple Count' as 'UoM Type',
'F' as 'Discrete',
cast(unit_of_measure.quantity_eaches as nvarchar(128)) as 'Conversion Factor',
'each' as 'Related UoM'
from unit_of_measure
where unit_of_measure.record_status = 1 and unit_of_measure.uom_code is not null and unit_of_measure.uom_code <> ''
and unit_of_measure.uom_code <> 'Each'
union
select distinct
item.unit_of_measure as 'Unit Name',
item.unit_of_measure as 'Abbreviation',
'Simple Count' as 'UoM Type',
'F' as 'Discrete',
'1.0000' as 'Conversion Factor',
'each' as 'Related UoM'
from item
where item.record_status = 1 and item.unit_of_measure is not null and item.unit_of_measure <> ''
union
select distinct
item.dimension_unit as 'Unit Name',
item.dimension_unit as 'Abbreviation',
'Length' as 'UoM Type',
'F' as 'Discrete',
'' as 'Conversion Factor',
'' as 'Related UoM'
from item
where item.record_status = 1 and item.dimension_unit is not null and item.dimension_unit <> ''
union
select distinct
item.weight_unit as 'Unit Name',
item.weight_unit as 'Abbreviation',
'Mass' as 'UoM Type',
'F' as 'Discrete',
'' as 'Conversion Factor',
'' as 'Related UoM'
from item
where item.record_status = 1 and item.dimension_unit is not null and item.dimension_unit <> ''