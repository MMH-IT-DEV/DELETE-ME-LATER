select 
supplier.supplier_number as 'Vendor Number',
ISNULL(supplier.name,supplier.supplier_number) as 'Vendor Name',
ISNULL(supplier.email,'') as 'Business Email',
ISNULL(supplier.website, '') as 'Website',
ISNULL(supplier.contact_name, '') as 'Contact Name',
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
            WHERE notes.note_id = supplier.supplier_id and notes.table_name = 'supplier' and notes.note_text is not null and notes.note_text <> ''
            ORDER BY notes.note_date
            FOR XML PATH ('')
        ), 2, 1000), '&#x0D;', '    '), CHAR(10), '    '), '') as 'Notes',
ISNULL(convert(varchar, supplier.usrdefdate1, 23),'') as 'Custom Date 1',
ISNULL(convert(varchar, supplier.usrdefdate2, 23),'') as 'Custom Date 2',
ISNULL(convert(varchar, supplier.usrdefdate3, 23),'') as 'Custom Date 3',
ISNULL(convert(varchar, supplier.usrdefdate4, 23),'') as 'Custom Date 4',
ISNULL(convert(varchar, supplier.usrdefdate5, 23),'') as 'Custom Date 5',
ISNULL(supplier.usrdefnumber1,0) as 'Custom Number 1',
ISNULL(supplier.usrdefnumber2,0) as 'Custom Number 2',
ISNULL(supplier.usrdefnumber3,0) as 'Custom Number 3',
ISNULL(supplier.usrdefnumber4,0) as 'Custom Number 4',
ISNULL(supplier.usrdefnumber5,0) as 'Custom Number 5',
ISNULL(supplier.usrdeftext1,'') as 'Custom Text 1',
ISNULL(supplier.usrdeftext2,'') as 'Custom Text 2',
ISNULL(supplier.usrdeftext3,'') as 'Custom Text 3',
ISNULL(supplier.usrdeftext4,'') as 'Custom Text 4',
ISNULL(supplier.usrdeftext5,'') as 'Custom Text 5',
ISNULL(supplier.usrdeftext6,'') as 'Custom Text 6',
ISNULL(supplier.usrdeftext7,'') as 'Custom Text 7',
ISNULL(supplier.usrdeftext8,'') as 'Custom Text 8',
ISNULL(supplier.usrdeftext9,'') as 'Custom Text 9',
ISNULL(supplier.usrdeftext10,'') as 'Custom Text 10'
from dbo.supplier
left outer join dbo.phone 
on supplier.supplier_id = phone.phone_id and table_name = 'supplier' and phone_type_id = 1 and phone_number is not null and phone_number <> ''
left outer join dbo.phone as cell
on supplier.supplier_id = cell.phone_id and cell.table_name = 'supplier' and cell.phone_type_id = 2 and cell.phone_number is not null and cell.phone_number <> ''
left outer join dbo.phone as fax
on supplier.supplier_id = fax.phone_id and fax.table_name = 'supplier' and fax.phone_type_id = 3 and fax.phone_number is not null and fax.phone_number <> ''
left outer join dbo.address
on supplier.supplier_id = address.address_id and address.table_name = 'supplier' and address.address_type_id = 1 and address.record_type = 'Supplier Billing'
where supplier.record_status = 1
and supplier.supplier_id > 0