/*
	SQL Server database backup script
	Made for Shulgin Adriy by Vasyl Pavuk
	March 2018
*/
function BackupClass(configuration)
{
	this.configuration = configuration; // configuration
	this.messages = new Array(); 		// the list of plain messages
	this.fails = 0;						// amount of failed operations during backup process
	this.FSO = new ActiveXObject('Scripting.FilesystemObject');
	this.connADO = new ActiveXObject('ADODB.Connection');
	this.connADO.ConnectionString = this.configuration.sql.connectionString;
	this.connADO.CommandTimeout = 0;
	this.connADO.Open();
/*
	function "Expand" add additional zeros to the "Value" till it's length become as "Length"
	parameters:
		Value: the value to be expanded
		Length: the length of result string expected
	returns:
		String like "00Value" with length of "Length" parameter (or more symbols if Value is longer)
*/
	function Expand(Value, Length)
	{
		for(Value = new String(Value);Value.length < Length;)
			Value = "0"+Value;
		return 	Value;
	}
/*
	function formatDate is designed to format date "d" as the following mask: YYYY-MM-DD
	parameters:
		d: incoming date to format
	returns: formated string like "YYYY-MM-DD"
*/
	function formatDate(d)
	{
		var result = d.getFullYear()+'-'+Expand(d.getMonth()+1,2)+'-'+Expand(d.getDate(),2);
		return result;
	}
/*
	function formatDateTime is designed to format date "d" as the following mask: YYYY-MM-DD
	parameters:
		d: incoming date to format
	returns: formated string like "YYYY-MM-DD--HH-MM"
*/
	function formatDateTime(d)
	{
		var result = d.getFullYear()+'-'+Expand(d.getMonth()+1,2)+'-'+Expand(d.getDate(),2)+'--'+
			Expand(d.getHours(),2)+'-'+Expand(d.getMinutes(), 2)+'-'+Expand(d.getSeconds(),2);
		return result;
	}
/*
	function "logMessage" to add new message to log
*/
	this.logMessage = function(message, type)
	{
		message = formatDateTime(new Date())+': '+message;
		this.messages.push(message);
		WScript.Echo(message);
	}
/*
	function "createFolder" creates folder and parent folders(if necessary)
	parameters:
		path: folder path to create
	returns:
		"true" or "false" means succeeded or failed
*/
	this.createFolder = function(path)
	{
			var folderParts = path.replace(/[\/]+/g, '\\').split('\\');
			var currentFolder = '';
			for(var index = 0; index < folderParts.length; index++)
			{
				currentFolder += folderParts[index]+'\\';
				if(!this.FSO.FolderExists(currentFolder))
					try
					{
						this.FSO.CreateFolder(currentFolder);
					}
					catch(err)
					{
						this.logMessage(err.description+' during folder creation '+currentFolder, 'error');
						return false;
					}
			}
			return true;
	}
/*
	function "proceedBackups" creates database backup files, compress it if necessary and sends file by FTP
	parameters:
		backupType: available values "full", "differential", "log"
	returns: this.fails - amount of fails during backup,compress and send routines
*/
	this.proceedBackups = function(backupType)
	{
		var currentDate = formatDate(new Date());
		
		for(var dbIndex = 0; dbIndex < this.configuration.sql.databases.length; dbIndex++)
		{
			// 1. backup file
			var backupFile = this.backupDatabase(this.configuration.sql.databases[dbIndex], backupType);
			// 2. compress file
			if(!backupFile)
				continue; // nothing to compress or send
			var compressedBackup = this.compressFile(backupFile, true);
			// 3. Send file to Disaster Recovery (DR) server
			if(compressedBackup)
				this.copyArchive(compressedBackup);
			// 4. Send Report
			
		}
	}
/*
	function "checkDatabaseExists" checks that database exists on server
	parameters:
		databaseName: the name of database to be checked
	returns: 0 				- database does not exists
			 value that > 0	- database id on SQL Server
*/
	this.checkDatabaseExists = function (databaseName)
	{
		var checkQuery = "select db_id('"+databaseName+"') databaseId";
		var rs = this.connADO.Execute(checkQuery);
		var result = String(rs('databaseId'));
		if(result=='null')
			return 0;
		return parseInt(result);
	}
/*
	function "backupDatabase" performs single database backup
	parameters:
		databaseName: 	the name of database that backup should performed
		backupType: 	available values "full", "differential", "log"
	returns:			path of backup file
*/
	this.backupDatabase=function(databaseName, backupType)
	{
		if(this.checkDatabaseExists(databaseName) > 0)
		{
			var backupDate = new Date(), backupFolder, backupFile, backupCommand;
			backupFolder = this.configuration.sql.backupFolder;
			if(!this.createFolder(backupFolder))
			{
				this.logMessage('Failed to create folder: '+backupFolder);
				return;
			}
			this.logMessage('Backup database ['+databaseName+"] type="+backupType, 'information');
			switch(backupType.toLowerCase())
			{
				case 	'full':
					backupFile = this.FSO.BuildPath(backupFolder, formatDateTime(backupDate)+'--full-['+databaseName+'].bak');
					backupCommand = "backup database ["+databaseName+"] to disk = N'"+backupFile+"';"
				break;
				case 	'differential':
				case	'diff':
					if(databaseName == 'master')
						return;
					backupFile = this.FSO.BuildPath(backupFolder, formatDateTime(backupDate)+'--differential-['+databaseName+'].bak');
					backupCommand = "backup database ["+databaseName+"] to disk = N'"+backupFile+"' with differential;";
				break;
				case	'log':
					if(databaseName == 'master')
						return;
					backupFile = this.FSO.BuildPath(backupFolder, formatDateTime(backupDate)+'--log-['+databaseName+'].trn');
					backupCommand = "backup log ["+databaseName+"] to disk = N'"+backupFile+"';"
				break;
				default:
					this.logMessage('Incorrect type of backup provided: '+backupType, 'information');
			}
			try
			{
				this.logMessage(backupCommand, 'information');
				var rs = this.connADO.Execute(backupCommand);
				while(rs)
				{
					var e = new Enumerator(this.connADO.Errors);
					for( ; !e.atEnd(); e.moveNext())
						this.logMessage('Msg '+e.item().NativeError+':\t '+ e.item().description, 'information');
					rs = rs.NextRecordset();
				}
			}
			catch(err)
			{
                // oops, something goes wrong
				this.fails++;
				this.logMessage('Backup database "'+databaseName+'" failed:');
				this.errorsCount++;
				if(rs)
					while(rs)
					{
						var e = new Enumerator(this.connADO.Errors);
						for( ;!e.atEnd(); e.moveNext())
							this.logMessage('Msg '+e.item().NativeError+':\t '+ e.item().description);
						rs = rs.NextRecordset();
					}
				else
				{
					var e = new Enumerator(this.connADO.Errors);
					for( ;!e.atEnd(); e.moveNext())
						this.logMessage('Msg '+e.item().NativeError+':\t '+ e.item().description);                     
				}
				return false; // failed to backup database
			}
		}
		else
		{
			this.logMessage('database ['+databaseName+'] does not exists!', 'error');
			this.fails++;
			return false; // failed to backup database
		}
		if(this.FSO.FileExists(backupFile))
			return backupFile;
		return false; // failed to backup database
	}
/*
	function "compressFile" performs a single file compression
	parameters:
		sourceFile: file path to compress
		deleteSource: shows delete file or no after successful compress;
	returns: path of file compressed
*/
	this.compressFile = function(sourceFile, deleteSource)
	{
		var wSh=new ActiveXObject("WScript.Shell");
		this.logMessage('Compress file:\t'+sourceFile);
		var currentDirectory = wSh.CurrentDirectory;
		wSh.CurrentDirectory=this.FSO.GetParentFolderName(sourceFile);
		var archiveFile = this.FSO.GetBaseName(sourceFile)+this.configuration.compress.extension;
		var CMD = this.configuration.compress.cmd.replace('$ARCHIVE', archiveFile).replace('$SOURCE', this.FSO.GetFileName(sourceFile));
		var retCode=wSh.Run(CMD, 1, true);
		if(deleteSource&&(retCode==0)&&FSO.FileExists(archiveFile))
			FSO.DeleteFile(sourceFile)
		else
			return false;
		wSh.CurrentDirectory = currentDirectory;
		return this.FSO.BuildPath(this.FSO.GetParentFolderName(sourceFile), archiveFile);
	}
/*
	function "copyArchive" makes a copy of file to configured folder
	parameters:
		sourceFile: the source file to be copied
*/
	this.copyArchive = function(sourceFile)
	{
		var targetFolder = this.configuration.copyfile.targetFolder;
		var targetFile = this.FSO.BuildPath(targetFolder, this.FSO.GetFileName(sourceFile));
		this.logMessage('this.copyArchive: '+sourceFile+'\t'+targetFile);
		try
		{
			this.FSO.CopyFile(sourceFile, targetFile);
		}
		catch(err)
		{
			this.logMessage(err.description);
		}
	}
/*
	function "saveLog" for save log messages collected to file on disk
	parameters:
		folderPath: - folder to store log
*/
	this.saveLog = function(folderPath)
	{
		var logFileName = this.FSO.BuildPath(folderPath, formatDateTime(new Date())+".log");
		var logFile = this.FSO.CreateTextFile(logFileName, true);
		logFile.Write(this.messages.join('\r\n'));
		logFile.Close();
	}
}
/********************************************************************************
	BEGIN EXECUTION
********************************************************************************/
// read configuration from file
var FSO = new ActiveXObject('Scripting.FilesystemObject');
var configFile = FSO.OpenTextFile('backup-configuration.js', 1);
var configurationText = configFile.ReadAll();
configFile.Close();

var configuration = eval(configurationText);

var backupObject = new BackupClass(configuration);
backupObject.proceedBackups('full');
backupObject.proceedBackups('diff');
backupObject.proceedBackups('log');
backupObject.logMessage('Total fails found:'+backupObject.fails);
backupObject.saveLog(FSO.GetParentFolderName(WScript.ScriptFullName));
