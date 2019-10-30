/*
    SQL Server database backup script
    Made for Shulgin Adriy by Vasyl Pavuk
    October 2019
*/
function BackupServer(backupType)
{
    var errorsCount = 0;
    var FSO = new ActiveXObject('Scripting.FilesystemObject');
    var wSh = new ActiveXObject('WScript.Shell');
    wSh.CurrentDirectory = FSO.GetParentFolderName(WScript.ScriptFullName);
    var wshNetwork = new ActiveXObject('WScript.Network');
    
    var configFile = FSO.OpenTextFile('backup-configuration.js', 1);
    var configurationText = configFile.ReadAll();
    configFile.Close();
    var configuration = eval(configurationText); // configuration
    
    var messages = new Array();             // the list of plain messages
    var messageType={'information': 1, 'warinng': 2, 'error': 3};
    var fails = 0;                          // amount of failed operations during backup process

    var connADO = new ActiveXObject('ADODB.Connection');
    connADO.ConnectionString = configuration.sql.connectionString;
    connADO.CommandTimeout = 0;
    connADO.Open();
    
    var backupDateFolder;
    var logFolder = FSO.BuildPath(FSO.GetParentFolderName(WScript.ScriptFullName),'/Log');
    {
        var xDate = formatDate(new Date()).split('-');
        backupDateFolder = xDate[0]+'\\'+xDate[0]+'-'+xDate[1]+'\\'+xDate.join('-');
    }

    for(var dbIndex = 0; dbIndex < configuration.sql.databases.length; dbIndex++)
    {
        // 1. backup file
        var backupFile = backupDatabase(configuration.sql.databases[dbIndex], backupType);

        // 2. compress file
        if(!backupFile)
            continue; // nothing to compress or send
        var compressedBackup = compressFile(backupFile, true);
        
        // 3. Send file to Disaster Recovery (DR) server
        if(compressedBackup)
            if(configuration.copyfile)
                copyArchive(compressedBackup);

        if(compressedBackup)
            if(configuration.ftp)
                uploadByFTP(compressedBackup);

        // 4. Send Report
    }
    saveLog();
    //sendTelegramMessage('backup script complete (<i>testing message</i>)');
    /********************************************************************************/

    /*
        function "backupDatabase" performs single database backup
        parameters:
            databaseName:   the name of database that backup should performed
            backupType:     available values "full", "differential", "log"
        returns:            path of backup file
    */
    function backupDatabase(databaseName, backupType)
    {
        /*
            function "checkDatabaseExists" checks that database exists on server
            parameters:
                databaseName: the name of database to be checked
            returns: 0;         //database does not exists
                     value that > 0    - database id on SQL Server
        */
        checkDatabaseExists = function (databaseName)
        {
            var checkQuery = "select db_id('"+databaseName+"') as databaseId";
            var rs = connADO.Execute(checkQuery);
            var result = String(rs('databaseId'));
            if(result=='null')
                return 0;
            return parseInt(result);
        }

        /*
            function "lastBackupDate" returns days past since last backup
            parameters:
                databaseName: the name of database to be checked
                backupType: type of backup that checked
            returns: number of days sicne last backup
        */
        function lastBackupDateFunction(databaseName, backupType)
        {
            var query = "exec sp_executesql\n\
                @stmt = N'select  datediff(day, max([backup_finish_date]), getdate()) [days_past]\n\
                from    [msdb].[dbo].[backupset]\n\
                where   [database_name] = @DatabaseName and [type] = @BackupType;',\n\
                @params = N'@DatabaseName sysname, @BackupType nchar(1)', @databaseName = N'"+databaseName+"', @backupType = N'"+backupType+"'";
            var rs = connADO.Execute(query);
            return parseInt(rs('days_past'));
        }

        /*
            function "getRecoveryModel" returns recovery model of database
            parameters:
                databaseName: the name of database to be checked
            returns:
                1: FULL
                2: BULK_LOGGED
                3: SIMPLE
        */
        getRecoveryModel = function(databaseName)
        {
            var query = "select [recovery_model] from sys.databases where [name] = '"+databaseName+"';";
            var rs = connADO.Execute(query);
            return parseInt(rs('recovery_model'));
        }

        if(checkDatabaseExists(databaseName) > 0)
        {
            var backupDate = new Date(), backupFolder, backupFile, backupCommand;
            backupFolder = FSO.BuildPath(configuration.sql.backupFolder, backupDateFolder);
            if(!createFolder(backupFolder))
            {
                logMessage('Failed to create folder: '+backupFolder, messageType['error']);
                return;
            }
            logMessage('Backup database ['+databaseName+"] type="+backupType, messageType['information']);
            var lastBackupDate = lastBackupDateFunction(databaseName, 'D');
            switch(backupType.toLowerCase())
            {
            case    'full':
                    backupFile = FSO.BuildPath(backupFolder, formatDateTime(backupDate)+'--full-['+databaseName+'].bak');
                    backupCommand = "backup database ["+databaseName+"] to disk = N'"+backupFile+"';"
                break;
                case    'differential':
                case    'diff':
                    if(lastBackupDate > 2)
                        backupDatabase(databaseName, 'full');
                    if(databaseName == 'master')
                        return;
                    backupFile = FSO.BuildPath(backupFolder, formatDateTime(backupDate)+'--differential-['+databaseName+'].bak');
                    backupCommand = "backup database ["+databaseName+"] to disk = N'"+backupFile+"' with differential;";
                break;
                case    'log':
                    if(lastBackupDate > 2)
                        backupDatabase(databaseName, 'full');
                    if(databaseName == 'master')
                        return;
                    if(getRecoveryModel(databaseName) == 3) // SIMPLE RECOVERY
                        return;
                    backupFile = FSO.BuildPath(backupFolder, formatDateTime(backupDate)+'--log-['+databaseName+'].trn');
                    backupCommand = "backup log ["+databaseName+"] to disk = N'"+backupFile+"';"
                break;
                default:
                    logMessage('Incorrect type of backup provided: '+backupType, messageType['information']);
            }
            try
            {
                logMessage(backupCommand, messageType['information']);
                var rs = connADO.Execute(backupCommand);
                while(rs)
                {
                    var e = new Enumerator(connADO.Errors);
                    for( ; !e.atEnd(); e.moveNext())
                        logMessage('Msg '+e.item().NativeError+':\t '+ e.item().description, messageType['information']);
                    rs = rs.NextRecordset();
                }
            }
            catch(err)
            {
                // oops, something goes wrong
                fails++;
                logMessage('Backup database "'+databaseName+'" failed:', messageType['error']);
                errorsCount++;
                if(rs)
                    while(rs)
                    {
                        var e = new Enumerator(connADO.Errors);
                        for( ;!e.atEnd(); e.moveNext())
                            logMessage('Msg '+e.item().NativeError+':\t '+ e.item().description, messageType['error']);
                        rs = rs.NextRecordset();
                    }
                else
                {
                    var e = new Enumerator(connADO.Errors);
                    for( ;!e.atEnd(); e.moveNext())
                        logMessage('Msg '+e.item().NativeError+':\t '+ e.item().description, messageType['error']);                     
                }
                sendTelegramMessage(wshNetwork.ComputerName + ':: Failed to backup: '+backupCommand);
                return false; // failed to backup database
            }
        }
        else
        {
            logMessage('database ['+databaseName+'] does not exists!', messageType['error']);
            fails++;
            return false; // failed to backup database
        }
        if(FSO.FileExists(backupFile))
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
    function compressFile(sourceFile, deleteSource)
    {
        var wSh = new ActiveXObject("WScript.Shell");
        logMessage('Compress file:\t'+sourceFile, messageType['information']);
        var currentDirectory = wSh.CurrentDirectory;
        wSh.CurrentDirectory = FSO.GetParentFolderName(sourceFile);
        var archiveFile = FSO.GetBaseName(sourceFile)+configuration.compress.extension;
        var CMD = configuration.compress.cmd.replace('$ARCHIVE', archiveFile).replace('$SOURCE', FSO.GetFileName(sourceFile));
        var wshExec  = wSh.Exec(CMD);
        while(wshExec.Status==0)
            WScript.Sleep(1000);
        var outText = wshExec.StdOut.ReadAll();
        if(deleteSource&&(wshExec.ExitCode ==0)&&FSO.FileExists(archiveFile)&&(outText.indexOf('Everything is Ok') >-1 ))
        {
            logMessage(outText, messageType.information);
            FSO.DeleteFile(sourceFile)
        }
        else
        {
            logMessage(outText, messageType.error);
            sendTelegramMessage(wshNetwork.ComputerName + ':: Failed to compress file: '+messageType.error);
            return false;
        }
        wSh.CurrentDirectory = currentDirectory;
        return FSO.BuildPath(FSO.GetParentFolderName(sourceFile), archiveFile);
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
/*
    function "logMessage" to add new message to log
*/
    function logMessage(message, type)
    {
        message = [formatDateTime(new Date()), message, type];
        messages.push(message);
        WScript.Echo(message);
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
    function "copyArchive" makes a copy of file to configured folder
    parameters:
        sourceFile: the source file to be copied
*/
    function copyArchive(sourceFile)
    {
        var targetFolder = FSO.BuildPath(configuration.copyfile.targetFolder, backupDateFolder);
        createFolder(targetFolder);
        var targetFile = FSO.BuildPath(targetFolder, FSO.GetFileName(sourceFile));
        logMessage('copyArchive: '+sourceFile+'\t'+targetFile, messageType['information']);
        try
        {
            FSO.CopyFile(sourceFile, targetFile);
        }
        catch(err)
        {
            logMessage(err.description, messageType.error);
            sendTelegramMessage(wshNetwork.ComputerName + ':: Failed to copy archive to ' + targetFile);
        }
    }

    function uploadByFTP(sourceFile)
    {
        var wSh = new ActiveXObject('WScript.Shell');
        var workTime = formatDateTime(new Date());
        var winSCP = FSO.BuildPath(FSO.GetParentFolderName(WScript.ScriptFullName),'/WinSCP/WinSCP.exe');
        var scriptFileName = FSO.BuildPath(logFolder, '/upload--'+workTime+'.wscp');
        var logFileName = FSO.BuildPath(logFolder, '/upload--'+workTime);
        createFolder(logFolder);
        var scriptFile = new ActiveXObject('ADODB.Stream');
        scriptFile.Type = 2;
        scriptFile.Charset = 'UTF-8';
        scriptFile.Open();
        var hostName = 'open '+configuration.ftp.hostname;
        /*
        if(configuration.certificate)
            hostName += ' -certificate="25:b8:ae:19:b4:34:4d:8f:1c:71:65:69:5e:80:20:7b:f0:c0:0e:ca"';
        */
        scriptFile.WriteText(hostName+'\n');
        scriptFile.WriteText('put '+sourceFile+'\n');
        scriptFile.WriteText('exit');
        scriptFile.SaveToFile(scriptFileName, 2);
        scriptFile.Close();
        var CMD = winSCP + ' /console /script='+scriptFileName+' /log='+logFileName+'.log /xmllog='+logFileName+'.xml';
        logMessage('Send file '+sourceFile+' to by FTP', messageType['information']);
        var wshExec = wSh.Exec(CMD);
        while(wshExec.Status == 0)
            WScript.Sleep(1000);
        if(wshExec.ExitCode != 0)
        {
            logMessage('Send file failed. Check details in the file '+logFileName+', Exit Code='+wshExec.ExitCode, messageType['error']);
        }
        else
        {
            logMessage('File '+sourceFile+' successful sent', messageType['information']);
        }
        FSO.DeleteFile(scriptFileName);
    }
/*
    function "saveLog" for save log messages collected to file on disk
    parameters:
        folderPath: - folder to store log
*/
    function saveLog()
    {
        var logFileName = FSO.BuildPath(logFolder, formatDateTime(new Date())+".html");
        createFolder(logFolder);
        var logHTML = new ActiveXObject("Msxml2.DOMDocument");
        var html = logHTML.createElement('html');
        logHTML.appendChild(html);
        var body = logHTML.createElement('body');
        html.appendChild(body);
        var table = logHTML.createElement('table');
        body.appendChild(table);
        table.setAttribute('border', '1');
        table.setAttribute('style', 'border-collapse:collapse');
        var headerRow = logHTML.createElement('tr');
        table.appendChild(headerRow);
        var headerDate = logHTML.createElement('th');
        headerRow.appendChild(headerDate);
        headerDate.appendChild(logHTML.createTextNode('Date/Time'));
        var headerMessage = logHTML.createElement('th');
        headerRow.appendChild(headerMessage);
        headerMessage.appendChild(logHTML.createTextNode('Message'));
        for(var messageId = 0; messageId < messages.length; messageId++)
        {
            var dataRow = logHTML.createElement('tr');
            table.appendChild(dataRow);
            var dateTimeCell = logHTML.createElement('td');
            dataRow.appendChild(dateTimeCell);
            var datePre = logHTML.createElement('pre');
            dateTimeCell.appendChild(datePre);
            dateTimeCell.setAttribute('valign', 'top');
            datePre.appendChild(logHTML.createTextNode(messages[messageId][0]));

            var messageCell = logHTML.createElement('td');
            dataRow.appendChild(messageCell);
            var messagePre = logHTML.createElement('pre');
            messageCell.appendChild(messagePre);

            var messageSpan = logHTML.createElement('span');
            messagePre.appendChild(messageSpan);
            messageSpan.appendChild(logHTML.createTextNode(messages[messageId][1]));

            if((messages[messageId][2])==messageType.error)
            {
                messageSpan.setAttribute('style', 'color:red');
            }
        }
        createFolder(logFolder);
        logHTML.save(logFileName);
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
}
/********************************************************************************
    BEGIN EXECUTION
********************************************************************************/
// read configuration from file
var currentTime = new Date();
if ((currentTime.getHours() >=3) && (currentTime.getHours() < 5 ) )
{
    BackupServer('full');
}
else
{
    BackupServer('log');
}
