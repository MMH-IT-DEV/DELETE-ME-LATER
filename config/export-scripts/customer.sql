select 
customer.customer_number as 'Customer Number',
ISNULL(customer.name,'') as 'First Name',
ISNULL(customer.company_name,'') as 'Company Name',
ISNULL(customer.department,'') as 'Department',
ISNULL(customer.email,'') as 'Contact Email',
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
            SELECT ' '+notes.note_text AS [text()] --REPLACE(REPLACE(notes.note_text, char(13), ''), char(10), '')
            FROM dbo.notes
            WHERE notes.note_id = customer.customer_id and notes.table_name = 'customer' and notes.note_text is not null and notes.note_text <> ''
            ORDER BY notes.note_date
            FOR XML PATH ('')
        ), 2, 1000), '&#x0D;', '    '), CHAR(10), '    '), '') as 'Notes',
ISNULL(convert(varchar, customer.usrdefdate1, 20),'') as 'Custom Date 1',
ISNULL(convert(varchar, customer.usrdefdate2, 20),'') as 'Custom Date 2',
ISNULL(convert(varchar, customer.usrdefdate3, 20),'') as 'Custom Date 3',
ISNULL(convert(varchar, customer.usrdefdate4, 20),'') as 'Custom Date 4',
ISNULL(convert(varchar, customer.usrdefdate5, 20),'') as 'Custom Date 5',
ISNULL(customer.usrdefnumber1,0) as 'Custom Number 1',
ISNULL(customer.usrdefnumber2,0) as 'Custom Number 2',
ISNULL(customer.usrdefnumber3,0) as 'Custom Number 3',
ISNULL(customer.usrdefnumber4,0) as 'Custom Number 4',
ISNULL(customer.usrdefnumber5,0) as 'Custom Number 5',
ISNULL(customer.usrdeftext1,'') as 'Custom Text 1',
ISNULL(customer.usrdeftext2,'') as 'Custom Text 2',
ISNULL(customer.usrdeftext3,'') as 'Custom Text 3',
ISNULL(customer.usrdeftext4,'') as 'Custom Text 4',
ISNULL(customer.usrdeftext5,'') as 'Custom Text 5',
ISNULL(customer.usrdeftext6,'') as 'Custom Text 6',
ISNULL(customer.usrdeftext7,'') as 'Custom Text 7',
ISNULL(customer.usrdeftext8,'') as 'Custom Text 8',
ISNULL(customer.usrdeftext9,'') as 'Custom Text 9',
ISNULL(customer.usrdeftext10,'') as 'Custom Text 10'
from dbo.customer
left outer join dbo.phone 
on customer.customer_id = phone.phone_id and table_name = 'customer' and phone_type_id = 1 and phone_number is not null and phone_number <> ''
left outer join dbo.phone as cell
on customer.customer_id = cell.phone_id and cell.table_name = 'customer' and cell.phone_type_id = 2 and cell.phone_number is not null and cell.phone_number <> ''
left outer join dbo.phone as fax
on customer.customer_id = fax.phone_id and fax.table_name = 'customer' and fax.phone_type_id = 3 and fax.phone_number is not null and fax.phone_number <> ''
left outer join dbo.address
on customer.customer_id = address.address_id and address.table_name = 'customer' and address.address_type_id = 1 and address.record_type = 'Customer Billing'
where customer.record_status = 1