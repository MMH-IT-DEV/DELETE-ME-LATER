select 
location.code as 'Location',
ISNULL(location.description, '') as 'Location Description',
sites.site_name as 'Site',
ISNULL(REPLACE(REPLACE(SUBSTRING(
        (
            SELECT ' '+ notes.note_text  AS [text()]
            FROM dbo.notes
            WHERE notes.note_id = location.location_id and notes.table_name = 'location' and notes.note_text is not null and notes.note_text <> ''
            ORDER BY notes.note_date
            FOR XML PATH ('')
        ), 2, 1000), '&#x0D;', '    '), CHAR(10), '    '), '') as 'Notes',
ISNULL(convert(varchar, location.usrdefdate1, 20),'') as 'Custom Date 1',
ISNULL(convert(varchar, location.usrdefdate2, 20),'') as 'Custom Date 2',
ISNULL(convert(varchar, location.usrdefdate3, 20),'') as 'Custom Date 3',
ISNULL(convert(varchar, location.usrdefdate4, 20),'') as 'Custom Date 4',
ISNULL(convert(varchar, location.usrdefdate5, 20),'') as 'Custom Date 5',
ISNULL(location.usrdefnumber1,0) as 'Custom Number 1',
ISNULL(location.usrdefnumber2,0) as 'Custom Number 2',
ISNULL(location.usrdefnumber3,0) as 'Custom Number 3',
ISNULL(location.usrdefnumber4,0) as 'Custom Number 4',
ISNULL(location.usrdefnumber5,0) as 'Custom Number 5',
ISNULL(location.usrdeftext1,'') as 'Custom Text 1',
ISNULL(location.usrdeftext2,'') as 'Custom Text 2',
ISNULL(location.usrdeftext3,'') as 'Custom Text 3',
ISNULL(location.usrdeftext4,'') as 'Custom Text 4',
ISNULL(location.usrdeftext5,'') as 'Custom Text 5',
ISNULL(location.usrdeftext6,'') as 'Custom Text 6',
ISNULL(location.usrdeftext7,'') as 'Custom Text 7',
ISNULL(location.usrdeftext8,'') as 'Custom Text 8',
ISNULL(location.usrdeftext9,'') as 'Custom Text 9',
ISNULL(location.usrdeftext10,'') as 'Custom Text 10'
from dbo.location
left outer join dbo.sites
on location.site_id = sites.site_id
where location.record_status = 1