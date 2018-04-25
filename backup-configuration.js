({
	'sql':
	{
		'connectionString': 'Provider=sqloledb;Data Source=.;Initial Catalog=master;Integrated Security=SSPI;',
		//'connectionString': 'Provider=sqloledb;Data Source=.;Initial Catalog=master;User Id=sa;Password=myPassword;'
		'backupFolder': 'C:\\Users\\Vasya\\Documents\\Projects\\Scripts\\ShulginAvto\\Backup-Shulgin\\backup-folder\\',
		'databases': ['master', 'msdb', 'userInfo']
	},
	'compress':
	{
		'extension': '.7z',
		'cmd': '\"C:\\Program Files\\7-Zip\\7z.exe\" a -mx9 \"$ARCHIVE\" \"$SOURCE\"'
	},
	'copyfile':
	{
		'targetFolder': 'C:\\Users\\Vasya\\Documents\\Projects\\Scripts\\ShulginAvto\\Backup-Shulgin\\backup-folder\\TargetFolder'
	},
    'ftp':
    {
        'hostname': 'ftp://public:public@localhost'
        
    }
})
