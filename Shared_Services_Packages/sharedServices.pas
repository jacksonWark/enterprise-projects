unit sharedServices;

interface

  uses
    System.SysUtils, System.Classes, System.IOUtils, System.DateUtils, System.IniFiles,
    Winapi.Windows, WinAPi.ShlObj,
    vcl.Forms, Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.Controls,
    ExeInfo;

  function ComputerName : string;

  //INI
  function RetValidPath(path : string) : string;
  function ProcessLocalIni : boolean;
  function GetRootLoc : string;
  function RootIniLoc : string;
  function DbIniLoc : string;
  function SynIniLoc : string;
  function GenIniLoc : string;
  function UserIniLoc : string;
  function CreateCopyUserDefs(defVals : TStringList) : boolean;
  procedure CheckDefaults(defVals : TStringList);
  function GetLoginInfo(connectionName : string; syn : boolean = false) : TArray<string>;
  function GetLoginInfoODBC(dbalias : string) : TArray<string>;
  function GetODBC(dbalias,schema : string) : string;
  function GetFooter(iniPath : string) : string;
  function GetSpecialFolderPath(folder : integer) : string;

  //Other files
  function CommonFiles : string;
  function LocalTempLoc : string;
  function LocalLogLoc : string;
  procedure SetFileAge(tempFileDays,logFileDays : integer);
  function ClearLocalTemp : boolean;
  function ClearLocalLog : boolean;
  Function ClearDir(age : integer; path : string) : boolean;

  //InitPanel
  procedure InitPanelOpen(parentForm : TComponent; mainText : string);
  procedure InitPanelAddSubHeading(text : string);
  procedure InitPanelClose;

  //Mail
  procedure EmergencyMail(subject,msg : string);

  var
    userRootLoc,
    rootLoc,
    tempLocalLoc,
    logLocalLoc,
    userIni : string;
    systemInfo : TExeInfo;
    tempAge,logAge : integer;
    initPanel : TPanel;
    labelArray : TArray<TLabel>;


implementation



function ComputerName : String;
var
  buffer: array[0..255] of char;
  size: dword;
begin
  size := 256;
  if GetComputerName(buffer, size) then
    Result := buffer
  else
    Result := ''
end;

{$REGION 'INI'}

function RetValidPath(path : string) : string;
begin
  result := path;
  if (path <> '') and (length(path) > 1) then
  begin
    if path[length(path)] <> '\' then
      result := result + '\';
  end
  else
    result := '';
end;


function ProcessLocalIni : boolean;
var
  appLocalIni : TIniFile;
  iniPath, iniServerPath, iniUserServerPath : string;
  doRead : boolean;
  splitPath : TArray<string>;
begin
  result := true;
  doRead := false;

  if (rootLoc = '') or (userRootLoc = '') then
  begin
    try
      iniPath := ExtractFileDir(application.ExeName) + '\local.ini';
      if FileExists(iniPath) then
        doRead := true;

      appLocalIni := TIniFile.Create( iniPath );
      if doRead then
      begin
        iniServerPath := appLocalIni.ReadString('APPDATA','serverPath','');
        iniUserServerPath := appLocalIni.ReadString('USERDATA','serverPath','');
        if (iniServerPath <> '') then
        begin
          rootLoc := RetValidPath(iniServerPath);

          //if LOCAL then find users directory on machine
          if ((uppercase(iniUserServerPath) = 'NONE') or (iniUserServerPath = '')) then
            userRootLoc := 'NONE'
          else
            userRootLoc := RetValidPath(iniUserServerPath);
        end
        else
          result := false;
      end
      else
      begin
        //write empty config so it's easier to configure and reduce typing errors
        appLocalIni.WriteString('APPDATA','serverPath','');
        appLocalIni.WriteString('USERDATA','serverPath','');
        result := false;
      end;
      appLocalIni.Free;

    except
      on E: Exception do
      begin
        result := false;
      end;
    end;
  end;
end;


function GetRootLoc : string;
begin
    
  if rootLoc <> '' then
    result := rootLoc
  else
  begin

    if ProcessLocalIni then
    begin
      if DirectoryExists(rootLoc) = false then
      begin
        try
          TDirectory.CreateDirectory(rootLoc);
        except
          on E: Exception do begin result := ''; exit; end;
        end;  
      end;
      result := rootLoc;
    end
    else
      result := '';
  end;
end;


function RootIniLoc : string;
begin
  result := GetRootLoc;
  if result <> '' then
  begin
    result := result + 'ini\';
    if directoryExists(result) = false then
    begin
      try
        TDirectory.CreateDirectory(result);
      except
        on E: Exception do result := '';
      end;
    end;
  end;
end;


function DbIniLoc : string;
begin
  result := RootIniLoc;
  if result <> '' then
  begin
    result := result + 'database\';
    if directoryExists(result) = false then
    begin
      try
        TDirectory.CreateDirectory(result);
      except
        on E: Exception do result := '';
      end;
    end;
  end;
end;


function SynIniLoc : string;
begin
  result := RootIniLoc;
  if result <> '' then
  begin
    result := result + 'synergize\';
    if directoryExists(result) = false then
    begin
      try
        TDirectory.CreateDirectory(result);
      except
        on E: Exception do result := '';
      end;
    end;
  end;
end;


function GenIniLoc : string;
begin
  result := RootIniLoc;
  if result <> '' then
  begin
    result := result + 'general\';
    if directoryExists(result) = false then
    begin
      try
        TDirectory.CreateDirectory(result);
      except
        on E: Exception do result := '';
      end;
    end;
    result := result + string.Join('',StringReplace(extractFileName(application.exename),'.exe','.ini',[]).Split(['NEW','New','new']));
  end;
end;


function UserIniLoc : string;
var
  path : string;
begin
  if userIni <> '' then
    result := userIni
  else
  begin
    result := '';
    if systemInfo = nil then
      systemInfo := TExeInfo.Create(nil);

    if (userRootLoc <> '') or ProcessLocalIni then
    begin
      path := userRootLoc + StringReplace(extractFileName(application.ExeName),'.exe','',[]) + '\';
      if directoryExists(path) = false then
      begin
        try
          TDirectory.CreateDirectory(path);
        except
          on E: Exception do
            exit
        end;

      end;
      path := path + systemInfo.UserName + '.ini';
      userIni := path;
      result := userIni;
    end;
  end;

end;


function CreateCopyUserDefs(defVals : TStringList) : boolean;
var
  iniPath : string;
  ini : TIniFile;
  I : integer;
  keyVal,sectionField : TArray<string>;
begin
  result := true;
  if fileExists(UserIniLoc) = false then
  begin
    iniPath := RootIniLoc + 'default\';
    if directoryExists(iniPath) = false then
    begin
      try TDirectory.CreateDirectory(iniPath);
      except on E: Exception do exit
      end;
    end;
    iniPath := iniPath + stringreplace(extractFilename(application.ExeName),'.exe','.ini',[]);
    if (fileExists(iniPath) = false) then //create it with default values
    begin
      try
        ini := TIniFile.Create(iniPath);
      except
        on E: Exception do
        begin
          result := false;
          exit;
        end;
      end;
      for I := 0 to defVals.Count - 1 do
      begin
        setlength(keyVal,0); setlength(sectionField,0);
        keyVal := defVals[I].Split(['=']);
        sectionField := keyVal[0].Split(['-']);
        try
          ini.WriteString(sectionField[0],sectionField[1],keyVal[1]);
        except
          on E: Exception do
          begin
            result := false;
            ini.Free;
            exit;
          end;
        end;
      end;
      ini.Free;
      if defVals.Count = 0 then
      begin
        result := false;
        exit;
      end;
    end;
    //now copy from default directory to user directory
    try
      TFile.Copy(iniPath,UserIniLoc);
    except
      on E: Exception do
        result := false;
    end;
  end;

end;


procedure CheckDefaults(defVals : TStringList);
var
  iniPath : string;
  ini : TIniFile;
  I : integer;
  keyVal,sectionField : TArray<string>;
begin
  if fileExists(UserIniLoc) then
  begin
    iniPath := RootIniLoc + 'default\' + stringreplace(extractFilename(application.ExeName),'.exe','.ini',[]);
    if fileExists(iniPath) then
    begin
      try
        ini := TIniFile.Create(iniPath);
      except
        on E: Exception do
        begin
          exit;
        end;
      end;
      for I := 0 to defVals.Count - 1 do
      begin
        setlength(keyVal,0); setlength(sectionField,0);
        keyVal := defVals[I].Split(['=']);
        sectionField := keyVal[0].Split(['-']);
        try
          ini.WriteString(sectionField[0],sectionField[1],keyVal[1]);
        except
          on E: Exception do
          begin
            ini.Free;
            exit;
          end;
        end;
      end;
      ini.Free;
    end;
  end;
end;


function GetLoginInfo(connectionName : string; syn : boolean = false) : TArray<string>;
var
  iniPath, typ : string;
  proceed : boolean;
  config : TIniFile;
begin
  if syn then
  begin
    iniPath := synIniLoc + connectionName + '.ini';
    typ := 'Synergize';
  end
  else
  begin
    iniPath := dbIniLoc + connectionName + '.ini';
    typ := 'Database';
  end;

  if (iniPath <> '') then
  begin
    if fileExists(iniPath) then
    begin
      try
        proceed := true;
        config := TIniFile.Create(iniPath);
      except on E: Exception do
        begin //couldnt open file are permissions okay?
          SetLength(result,1);
          result[0] := 'Failed to open database configuration file. Please contact IT to permissions are configured properly';
          proceed := false;
        end;
      end;
      if proceed then
      begin
        setLength(result,3);
        result[0] := config.ReadString('Login','username','');
        result[1] := config.ReadString('Login','password','');
        if syn then
          result[2] := config.ReadString('Login','CurrentSchema','')
        else
          result[2] := config.ReadString('Login','CurrentSchema','');
        config.Free;
      end;
    end
    else
    begin
      SetLength(result,1);
      result[0] := 'The '+typ+' configuration ('+connectionName+') file does not exist. Please contact IT to ensure local.ini is configured properly';
    end;
  end
  else
  begin
    SetLength(result,1);
    result[0] := 'Unable to find '+typ+' configuration was supplied.';
  end;
end;


function GetLoginInfoODBC(dbalias : string) : TArray<string>;
var
  iniPath, typ : string;
  proceed : boolean;
  config : TIniFile;
begin
  iniPath := dbIniLoc + dbalias + '.ini';

  if (iniPath <> '') then
  begin
    if fileExists(iniPath) then
    begin
      try
        proceed := true;
        config := TIniFile.Create(iniPath);
      except on E: Exception do
        begin //couldnt open file are permissions okay?
          SetLength(result,1);
          result[0] := 'Failed to open database configuration file. Please contact IT to permissions are configured properly';
          proceed := false;
        end;
      end;
      if proceed then
      begin
        setLength(result,4);
        result[1] := config.ReadString('Login','username','');
        result[2] := config.ReadString('Login','password','');
        result[3] := config.ReadString('Login','CurrentSchema','');
        config.Free;

        var odbc := GetODBC(dbalias,uppercase(result[3]));
        if odbc <> '' then
          result[0] := odbc
        else
        begin
          SetLength(result,1);
          result[0] := 'Unable to locate a suitable ODBC connection. Please contact IT.';
        end;
      end;
    end
    else
    begin
      SetLength(result,1);
      result[0] := 'The Database configuration ('+dbalias+') file does not exist. Please contact IT to ensure local.ini is configured properly';
    end;
  end
  else
  begin
    SetLength(result,1);
    result[0] := 'Unable to find '+typ+' configuration was supplied.';
  end;
end;


function GetODBC(dbalias,schema : String) : string;
var
  path: String;
  i: Integer;
  inifile: TiniFile;
  iniContents : TStringList ;
begin
  result := '';
  path := GetSpecialFolderPath(CSIDL_PROFILE) + '\db2cli.ini';
  if FileExists(path)=True then
  begin
    iniContents := TStringList .Create;
    inifile:=tinifile.Create(path);
    try
      inifile.ReadSections(iniContents);
      for i:=0 to  iniContents.Count - 1 do
      begin
        if  (uppercase(inifile.ReadString(inicontents[I],'DBALIAS','')) = uppercase(dbalias)) then
        begin
          var iniSchema := uppercase(inifile.ReadString(inicontents[I],'CURRENTSCHEMA',''));
          if (iniSchema = schema) or (iniSchema = 'TMAPP') then
          begin
            var funcGood := false;
            var funcPath := uppercase(inifile.ReadString(inicontents[I],'CURRENTFUNCTIONPATH',''));
            var splitFunc := funcPath.Split([',']);
            for var J := 0 to length(splitFunc) - 1 do
              if splitFunc[J] = schema then   //checking only for schema in function path, and not SYSFUN,SYSPROC,SYSIBM
                funcGood := true;
            if funcGood then
            begin
              result := inicontents[I];
              break
            end;
          end;
        end;
      end;
    finally
      inifile.Free;
      inicontents.Free;
    end;
  end;
end;


function GetFooter(iniPath : string) : string;
begin
  result := '';
  try
    var config := TIniFile.Create(iniPath);
    var footerCt := config.ReadInteger('MailBody','Fcount',-1);
    if footerCt > -1 then
    begin
      for var I := 0 to footerCt do
        result := result + config.ReadString('MailBody','F' + i.ToString,'') + slinebreak;
    end;
  except
    result := '';
  end;
end;


function GetSpecialFolderPath(folder : integer) : string;
const
 SHGFP_TYPE_CURRENT = 0;
var
 path: array [0..MAX_PATH] of char;
begin
 if SUCCEEDED(SHGetFolderPath(0,folder,0,SHGFP_TYPE_CURRENT,@path[0])) then
   Result := path
 else
   Result := 'C:\Temp';
end;


{$ENDREGION}

{$REGION 'Other Files'}

function CommonFiles : string;
begin
  result := GetRootLoc;
  if result <> '' then
  begin
    result := result + 'Common Files\';
    if directoryExists(result) = false then
    begin
      try
        TDirectory.CreateDirectory(result);
      except
        on E: Exception do result := '';
      end;
    end;
  end;
end;


function LocalTempLoc : string;
var
  path : string;
begin
  //Look at tempLocalLoc. If it exists then simply return
  if tempLocalLoc <> '' then
  begin
    result := tempLocalLoc;
  end
  else
  begin
    try
      if systemInfo = nil then
        systemInfo := TExeInfo.Create(nil);

      path := 'C:\Users\' + systemInfo.UserName + '\AppData\Local\example\'; //removed
      if DirectoryExists(path) = false then
      begin
        //create it
        TDirectory.CreateDirectory(path);
      end;

      path := path + string.Join('',extractFilename(application.ExeName).Split(['.exe','NEW','New','new'])) + '\';
      if DirectoryExists(path) = false then
      begin
        //create it
        TDirectory.CreateDirectory(path);
      end;
      tempLocalLoc := path;
      result :=  tempLocalLoc;
    except
      on E: Exception do
      begin
        result := '';
      end;
    end;
  end;
end;


function LocalLogLoc : string;
var
  path : string;
begin
  result := '';
  //Look at logLocalLoc. If it exists then simply return
  if logLocalLoc <> '' then
  begin
    result := logLocalLoc;
  end
  else
  begin
    //look for where the app log directory should be. If it is not there then create it
    try
      if systemInfo = nil then
        systemInfo := TExeInfo.Create(nil);

      path := 'C:\Users\' + systemInfo.UserName + '\AppData\Local\example\'; //removed
      if DirectoryExists(path) = false then
      begin
        //create it
        TDirectory.CreateDirectory(path);
      end;
      path := path + 'logs\';
      if DirectoryExists(path) = false then
      begin
        //create it
        TDirectory.CreateDirectory(path);
      end;
      path := path + string.Join('',extractFilename(application.ExeName).Split(['.exe','NEW','New','new'])) + '\logs\';
      if DirectoryExists(path) = false then
      begin
        //create it
        TDirectory.CreateDirectory(path);
      end;
      logLocalLoc := path;
      result := logLocalLoc;
    except
      on E: Exception do
      begin
        result := '';
      end;
    end;
  end;
end;


procedure SetFileAge(tempFileDays,logFileDays : integer);
begin
  if tempFileDays = 0 then
    raise Exception.Create('sharedServicesUtils: setting value tempFileDays = 0 will cause ClearLocalTemp to fail.')
  else
    tempAge := tempFileDays;

  if logFileDays = 0 then
    raise Exception.Create('sharedServicesUtils: setting value logFileDays = 0 will cause ClearLocalLog to fail.')
  else
    logAge := logFileDays;
end;


function ClearLocalTemp : boolean;
begin
  result := false;
  if tempAge = 0 then
  begin
    raise Exception.Create('sharedServicesUtils: function ClearLocalTemp was called without initializing temorary file age by calling SetFileAge.');
  end
  else
  begin
    result := ClearDir(tempAge,LocalTempLoc);
  end;


end;


function ClearLocalLog : boolean;
begin
  result := false;
  if logAge = 0 then
  begin
    raise Exception.Create('sharedServicesUtils: function ClearLocalLog was called without initializing log file age by calling SetFileAge.');
  end
  else
  begin
    result := ClearDir(logAge,LocalLogloc);
  end;
end;


//Returns the number of failures
function ClearDir(age : integer; path : string) : boolean;
var
  I,fail : integer;
  currDate, fileDate : TDateTime;
  files : TArray<string>;
begin
  result := false;
  try
    files := TDirectory.GetFiles(path,'*',TSearchOption.soAllDirectories);
  except
    on E: Exception do
      exit;
  end;

  if length(files) = 0 then
    result := true
  else
  begin
    currDate := Now;
    for I := 0 to (length(files) - 1) do
    begin
      try
        FileAge(files[I],fileDate);
        if WithinPastDays(currDate,fileDate,age) = false then
          TFile.Delete(files[I]);
      except
        on E: Exception do
          inc(fail);
      end;
    end;
    if fail <> I then
      result := true;
  end;

end;

{$ENDREGION}

{$REGION 'InitPanel'}

procedure InitPanelOpen(parentForm : TComponent; mainText : string);
var
  dims : TRect;
begin
  initPanel := TPanel.Create(parentForm);
  with initPanel do
  begin
    //Align := alClient;
    parent := (parentForm as TWinControl);
    top := 0;
    left := 0;
    dims := (parentForm as TForm).ClientRect;
    width := dims.Width;
    height := dims.Height;
    ShowCaption := false;
    bevelOuter := bvNone;
    BringToFront;
    visible := true;
  end;

  SetLength(labelArray,1);
  labelArray[0] := TLabel.Create(initPanel);
  with labelArray[0] do
  begin
    parent := initPanel;
    align := alTop;
    margins.Top := 40;
    margins.Left := 0;
    margins.Right := 0;
    margins.Bottom := 10;
    AlignWithmargins := true;
    font.Size := 16;
    caption := mainText;
    Alignment := taCenter;
    visible := true;
  end;

  Application.ProcessMessages;
end;


procedure InitPanelAddSubHeading(text : string);
var
  len,I : integer;
begin
  len := length(labelArray) + 1;
  setLength(labelArray,len);
  labelArray[len - 1] := TLabel.Create(initPanel);
  with labelArray[len - 1] do
  begin
    parent := initPanel;
    align := alTop;
    margins.Top := 10;
    margins.Left := 0;
    margins.Right := 0;
    margins.Bottom := 0;
    AlignWithmargins := true;
    font.Size := 12;
    caption := text;
    Alignment := taCenter;
    visible := true;
  end;

  for I := (len - 1) downto 0 do
  begin
    labelArray[I].Top := 0;
  end;

  initpanel.BringToFront;
  Application.ProcessMessages;
end;


procedure InitPanelClose;
var
  I : integer;
begin
  for I := 0 to (length(labelArray) - 1) do
    labelArray[I].Free;

  initPanel.Free;
  Application.ProcessMessages;
end;

{$ENDREGION}


end.
