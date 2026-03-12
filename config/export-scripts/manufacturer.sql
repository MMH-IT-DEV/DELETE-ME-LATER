select 
ISNULL(manufacturer.name,'') as 'Manufacturer Name',
ISNULL(manufacturer.email,'') as 'Business Email',
ISNULL(manufacturer.website, '') as 'Website',
ISNULL(manufacturer.contact_name, '') as 'Contact Name',
ISNULL(phone.phone_number, '') as 'Contact Phone',
ISNULL(phone.extension, '') as 'Contact Extension',
ISNULL(cell.phone_number, '') as 'Contact Cell',
ISNULL(fax.phone_number, '') as 'Contact Fax',
ISNULL(address.address1, '') as 'Address 1',
ISNULL(address.address2, '') as 'Address 2',
ISNULL(address.mail_stop, '') as 'Mail Stop',
ISNULL(address.city, '') as 'City',
ISNULL(address.state, '') as 'State',
ISNULL(address.postal_code, '') as 'Postal Code',
ISNULL(address.country, '') as 'Country',
ISNULL(REPLACE(REPLACE(SUBSTRING(
        (
            SELECT ' '+notes.note_text  AS [text()]
            FROM dbo.notes
            WHERE notes.note_id = manufacturer.manufacturer_id and notes.table_name = 'manufacturer' and notes.note_text is not null and notes.note_text <> ''
            ORDER BY notes.note_date
            FOR XML PATH ('')
        ), 2, 1000), '&#x0D;', '    '), CHAR(10), '    '), '') as 'Notes',
ISNULL(convert(varchar, manufacturer.usrdefdate1, 23),'') as 'Custom Date 1',
ISNULL(convert(varchar, manufacturer.usrdefdate2, 23),'') as 'Custom Date 2',
ISNULL(convert(varchar, manufacturer.usrdefdate3, 23),'') as 'Custom Date 3',
ISNULL(convert(varchar, manufacturer.usrdefdate4, 23),'') as 'Custom Date 4',
ISNULL(convert(varchar, manufacturer.usrdefdate5, 23),'') as 'Custom Date 5',
ISNULL(manufacturer.usrdefnumber1,0) as 'Custom Number 1',
ISNULL(manufacturer.usrdefnumber2,0) as 'Custom Number 2',
ISNULL(manufacturer.usrdefnumber3,0) as 'Custom Number 3',
ISNULL(manufacturer.usrdefnumber4,0) as 'Custom Number 4',
ISNULL(manufacturer.usrdefnumber5,0) as 'Custom Number 5',
ISNULL(manufacturer.usrdeftext1,'') as 'Custom Text 1',
ISNULL(manufacturer.usrdeftext2,'') as 'Custom Text 2',
ISNULL(manufacturer.usrdeftext3,'') as 'Custom Text 3',
ISNULL(manufacturer.usrdeftext4,'') as 'Custom Text 4',
ISNULL(manufacturer.usrdeftext5,'') as 'Custom Text 5',
ISNULL(manufacturer.usrdeftext6,'') as 'Custom Text 6',
ISNULL(manufacturer.usrdeftext7,'') as 'Custom Text 7',
ISNULL(manufacturer.usrdeftext8,'') as 'Custom Text 8',
ISNULL(manufacturer.usrdeftext9,'') as 'Custom Text 9',
ISNULL(manufacturer.usrdeftext10,'') as 'Custom Text 10'
from dbo.manufacturer
left outer join dbo.phone 
on manufacturer.manufacturer_id = phone.phone_id and table_name = 'manufacturer' and phone_type_id = 1 and phone_number is not null and phone_number <> ''
left outer join dbo.phone as cell
on manufacturer.manufacturer_id = cell.phone_id and cell.table_name = 'manufacturer' and cell.phone_type_id = 2 and cell.phone_number is not null and cell.phone_number <> ''
left outer join dbo.phone as fax
on manufacturer.manufacturer_id = fax.phone_id and fax.table_name = 'manufacturer' and fax.phone_type_id = 3 and fax.phone_number is not null and fax.phone_number <> ''
left outer join dbo.address
on manufacturer.manufacturer_id = address.address_id and address.table_name = 'manufacturer' and address.address_type_id = 1 and address.record_type = 'Manufacturer Billing'
where manufacturer.record_status = 1