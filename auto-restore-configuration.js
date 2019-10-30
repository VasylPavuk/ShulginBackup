({
    'sql':
    {
        'connectionString': 'Provider=sqloledb;Data Source=.;Initial Catalog=master;User Id=sa;Password=myPassword;',
        'backupFolder': 'backupFolder',
        'databases': ['Database1','Database2']
    },
    'compress':
    {
        'extension': '.7z',
        'cmd': '\"C:\\Program Files\\7-Zip\\7z.exe\" x -y \"$ARCHIVE\"'
    },
    'telegram':
    {
        'token': '0000000000:AAA-BBBBBBBBBBBBBBBBBBB_CCCCCCCCCCC',
        'chat_id': '-1'
    }
})
