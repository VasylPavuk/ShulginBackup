autoRestore();

function autoRestore()
{

    var FSO = new ActiveXObject('Scripting.FilesystemObject');
    var wshNetwork = new ActiveXObject('WScript.Network');
    var wSh = new ActiveXObject('WScript.Shell');
    wSh.CurrentDirectory = FSO.GetParentFolderName(WScript.ScriptFullName);
    var logFolder = FSO.BuildPath(wSh.CurrentDirectory, 'Logs');
    createFolder(logFolder);
    var logFileName = FSO.BuildPath(logFolder, formatDate(new Date())+'.log');
    var logFile = FSO.OpenTextFile(logFileName, 8 /* append file */, true);
    logMessage('Start autoRestore()');

    var configuration = loadConfiguration();

    var connADO = new ActiveXObject('ADODB.Connection');
    connADO.ConnectionString = configuration.sql.connectionString;
    connADO.CommandTimeout = 0
    connADO.Open();

    for(var databaseIndex = 0; databaseIndex < configuration.sql.databases.length; databaseIndex++)
        restoreDatabase(configuration.sql.databases[databaseIndex]);

    connADO.Close();
    logMessage('Finish autoRestore()\n\n');

    function loadConfiguration()
    {
        var configurationFile = FSO.OpenTextFile('auto-restore-configuration.js', 1);
        var configurationText = configurationFile.ReadAll();
        configurationFile.Close();
        return eval (configurationText);
    }
    
    function logMessage(message)
    {
        message = formatDateTime(new Date()) + ' '+message;
        logFile.WriteLine(message);
    }
    /*
        function "createFolder" creates folder and parent folders(if necessary)
        parameters:
            path: folder path to create
        returns:
            "true" or "false" means succeeded or failed
    */
    function createFolder(path)
    {
        if(FSO.FolderExists(path))
            return 1;
        var parentFolder = FSO.GetParentFolderName(path);
        if(!FSO.FolderExists(parentFolder))
            createFolder(parentFolder);

        if(!FSO.FolderExists(path))
            try
            {
                FSO.CreateFolder(path);
            }
            catch(err)
            {
                logMessage(err.description+' during folder creation '+currentFolder, messageType['error']);
                sendTelegramMessage(wshNetwork.ComputerName + ':: Failed to create folder: '+path);
                return false;
            }
        return true;
    }

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
        return     Value;
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

    function sendTelegramMessage(messageText)
    {
        var token=configuration.telegram.token,
            chat_id = configuration.telegram.chat_id,
            method = 'sendMessage';
        var url = 'https://api.telegram.org/bot'+token+'/'+method+'?chat_id='+chat_id+'&parse_mode=html'+'&text='+(messageText);
        var xmlhttp = new ActiveXObject("Msxml2.XMLHTTP");
        WScript.Echo(url);
        xmlhttp.open('GET', url, false);
        xmlhttp.send(null);
        WScript.Echo(xmlhttp.status);
        if(xmlhttp.status == 200)
            return true;
        return false
    }

    function restoreDatabase(databaseName)
    {
        /* TODO:
            database not exists - load from last full + logs
            database exists and online - do nothing + worning
            database exists and restoring - check the last backup restored and proceed next one
        */

        var dbId = parseInt(connADO.Execute("select isnull(db_id('"+databaseName+"'), 0);")(0));
        if(dbId == 0)
        {    // database does not exists, restore from full + logs
            var lastFull = -1;
            var backupsList = listBackups(configuration.sql.backupFolder);
            for(var fileIndex = 0; fileIndex < backupsList.length; fileIndex++)
                if(backupsList[fileIndex].indexOf('--full-') > -1)
                    lastFull = fileIndex;

                for(var fileIndex = lastFull; fileIndex < backupsList.length; fileIndex++)
                    restoreFile(decompressBackup(backupsList[fileIndex]));
        }
        if(dbId > 0)
        {
            // database exists, check state
            var dbState = connADO.Execute("select [state_desc] from sys.databases where [name] = '"+databaseName+"';");
            if(dbState('state_desc') != 'RESTORING')
            {
                // nothing to do
                return;
            }
            var restoreRS = connADO.Execute("SET NOCOUNT ON; SELECT TOP 1 rsh.restore_type, rsh.restore_date, bmf.physical_device_name\n\
            FROM    msdb.dbo.restorehistory rsh\n\
                    INNER JOIN msdb.dbo.backupset bs ON rsh.backup_set_id = bs.backup_set_id\n\
                    INNER JOIN msdb.dbo.backupmediafamily bmf ON bmf.media_set_id = bs.media_set_id\n\
            WHERE   rsh.destination_database_name = N'"+databaseName+"'\n\
            ORDER BY rsh.restore_date DESC;");
            var lastRestoreFile = FSO.GetBaseName(String(restoreRS('physical_device_name')));
            var backupsList = listBackups(configuration.sql.backupFolder);
            var lastBacckup = -1;
            for(var fileIndex = 0; fileIndex < backupsList.length; fileIndex++)
                if(FSO.GetBaseName(backupsList[fileIndex]) == lastRestoreFile)
                    lastBacckup = fileIndex;
            if(lastBacckup > -1)
                for(var fileIndex = lastBacckup+1; fileIndex < backupsList.length; fileIndex++)
                    restoreFile(decompressBackup(backupsList[fileIndex]));
        }

        function listBackups (baseFolder)
        {
            var backupsList = new Array();
            var strToSearch = '['+databaseName+']';
            
            backupsSearch(baseFolder);
            backupsList.sort();
            return backupsList;
            
            function backupsSearch(path)
            {
                var folder = FSO.GetFolder(path);
                for(var files = new Enumerator(folder.Files); !files.atEnd(); files.moveNext())
                    if((files.item().Name.indexOf(strToSearch) > -1))
                        backupsList.push(files.item().Path);
                for(var folders = new Enumerator(folder.Subfolders); !folders.atEnd(); folders.moveNext())
                    backupsSearch(folders.item().Path);
            }
        }
        
        function decompressBackup(archiveName)
        {
            var folderName = FSO.GetParentFolderName(archiveName);
            var baseName = FSO.GetBaseName(archiveName);
            var archiveExtension = FSO.GetExtensionName(archiveName);
            wSh.CurrentDirectory = FSO.GetParentFolderName(archiveName);
            var backupFile = FSO.BuildPath(folderName, baseName+".bak");

            var CMD = configuration.compress.cmd.replace('$ARCHIVE', archiveName);
            wSh.Run(CMD, 1, true);
            wSh.CurrentDirectory = FSO.GetParentFolderName(WScript.ScriptFullName);
            if(FSO.FileExists(FSO.BuildPath(folderName, baseName+'.bak')))
            {
                return FSO.BuildPath(folderName, baseName+'.bak');
            }
            if(FSO.FileExists(FSO.BuildPath(folderName, baseName+'.trn')))
            {
                return FSO.BuildPath(folderName, baseName+'.trn');
            }
            return false;
        }

        function restoreFile(backupFile)
        {
            var query='';
            if(backupFile.indexOf('--full-') > -1)
            {
                query = "restore database ["+databaseName+"] from disk = N'"+
                    backupFile+"' with norecovery, replace";
            }
            if(backupFile.indexOf('--log-') > -1)
            {
                query = "restore log ["+databaseName+"] from disk = N'"+
                    backupFile+"' with norecovery";
            }
            if(query != '')
            {
                try
                {
                    logMessage(query);
                    connADO.Execute(query);
                }
                catch(err)
                {
                    logMessage(err.description);
                    //sendTelegramMessage(err.description);
                    sendTelegramMessage('<b>'+wshNetwork.ComputerName+'::autoRestore()</b>error: \n'+query);
                }
            }
            if(FSO.FileExists(backupFile))
                FSO.DeleteFile(backupFile);
        }
    }
}
