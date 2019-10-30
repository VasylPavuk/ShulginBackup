({
    'sql':
    {
        'connectionString': 'Provider=sqloledb;Data Source=.;Initial Catalog=master;Integrated Security=SSPI;',
        //'connectionString': 'Provider=sqloledb;Data Source=.;Initial Catalog=master;User Id=sa;Password=myPassword;'
        'backupFolder': 'C:\\Git\\ShulginBackup\\backup-folder\\',
        'databases': ['master', 'msdb', 'Deezze.Log']
    },
    'compress':
    {
        'extension': '.7z',
        'cmd': '\"C:\\Program Files\\7-Zip\\7z.exe\" a -mx9 \"$ARCHIVE\" \"$SOURCE\"'
    },
    'copyfile1':
    {
        'targetFolder': 'C:\\Git\\ShulginBackup\\backup-folder\\TargetFolder'
    },
    'ftp':
    {
        'hostname': 'ftp://public:public@localhost'
        
    },
    'telegram':
    {
        'token': '0000000000:AAA-BBBBBBBBBBBBBBBBBBB_CCCCCCCCCCC',
        'chat_id': '-1'
    }
})
