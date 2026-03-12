select 
sites.site_name as 'Site',
ISNULL(sites.description, '') as 'Site Description',
ISNULL(REPLACE(REPLACE(SUBSTRING(
        (
            SELECT ' '+notes.note_text  AS [text()]
            FROM dbo.notes
            WHERE notes.note_id = sites.site_id and notes.table_name = 'sites' and notes.note_text is not null and notes.note_text <> ''
            ORDER BY notes.note_date
            FOR XML PATH ('')
        ), 2, 1000), '&#x0D;', '    '), CHAR(10), '    '), '') as 'Notes',
ISNULL(convert(varchar, sites.usrdefdate1, 23),'') as 'Custom Date 1',
ISNULL(convert(varchar, sites.usrdefdate2, 23),'') as 'Custom Date 2',
ISNULL(convert(varchar, sites.usrdefdate3, 23),'') as 'Custom Date 3',
ISNULL(convert(varchar, sites.usrdefdate4, 23),'') as 'Custom Date 4',
ISNULL(convert(varchar, sites.usrdefdate5, 23),'') as 'Custom Date 5',
ISNULL(sites.usrdefnumber1,0) as 'Custom Number 1',
ISNULL(sites.usrdefnumber2,0) as 'Custom Number 2',
ISNULL(sites.usrdefnumber3,0) as 'Custom Number 3',
ISNULL(sites.usrdefnumber4,0) as 'Custom Number 4',
ISNULL(sites.usrdefnumber5,0) as 'Custom Number 5',
ISNULL(sites.usrdeftext1,'') as 'Custom Text 1',
ISNULL(sites.usrdeftext2,'') as 'Custom Text 2',
ISNULL(sites.usrdeftext3,'') as 'Custom Text 3',
ISNULL(sites.usrdeftext4,'') as 'Custom Text 4',
ISNULL(sites.usrdeftext5,'') as 'Custom Text 5',
ISNULL(sites.usrdeftext6,'') as 'Custom Text 6',
ISNULL(sites.usrdeftext7,'') as 'Custom Text 7',
ISNULL(sites.usrdeftext8,'') as 'Custom Text 8',
ISNULL(sites.usrdeftext9,'') as 'Custom Text 9',
ISNULL(sites.usrdeftext10,'') as 'Custom Text 10'
from dbo.sites
where sites.record_status = 1 and
sites.site_id > 0