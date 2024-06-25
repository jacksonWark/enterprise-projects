unit service;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Classes, System.JSON, System.AnsiStrings, System.DateUtils,
  System.IniFiles, System.StrUtils, System.Math, System.RegularExpressions,
  Vcl.Graphics, Vcl.Controls, Vcl.SvcMgr, Vcl.Dialogs, ExeInfo, Vcl.ExtCtrls,
  IdEMailAddress, IdComponent, IdTCPConnection, IdTCPClient, IdExplicitTLSClientServerBase,
  IdMessageClient, IdSMTPBase, IdSMTP, IdBaseComponent, IdMessage, IdHTTP, IdIOHandlerSocket,
  IdIOHandler, IdIOHandlerStack, IdSSL, IdSSLOpenSSL,
  FireDAC.Stan.Param,
  //System.Variants, System.NetEncoding, System.Net.URLClient, Data.DB, IdAuthentication,
  System.Generics.Collections, Registry,
  outlook, sharedservices, crypto;

type
  TTGIConAPIsvc = class(TService)
    ExeInfo: TExeInfo;
    procedure ServiceExecute(Sender: TService);
    procedure ServiceAfterInstall(Sender: TService);

  private
    { Private declarations }
    //TGI Stuff
    token : string;
    tokenExpire : TDateTime;
    //TGI Stuff
    //JSON
    currID : string;
    //Query building
    insDataFlds, insDataVals : string;
    insDataFirstFld : boolean;
    insArrFlds, insArrVals, currArrTable : string;
    insArrFirstFld : boolean;
    //email stuff
    emailNewFlds, emailDataAlert : string;
    OutCtrl : TOutlookController;
    mailLog, mailLogDtl : Boolean;
    //xml,url,
    logErrorName: String;
    LogID,LogDtlID: Integer;
    IsDefaultEmail,testProfile,aTestMode: Boolean;

    procedure autoRun;
    procedure runProfile;
    procedure ActiveProcessesReady;
    procedure LogErr(eID: Integer; aError: String);
    procedure LogDtl(sendData: String);
    procedure InsLog;
    procedure qUpdLogDtl(aReply: String);
    function checkSetup : Boolean;
    function GetSysEmail: String;
    function ValidEmail(email: string): boolean;
    procedure SetupGraphMail(isTest : boolean);
    procedure graphMail(aTo,aSubj,aBody : string; isTest : boolean);
    //TGI Stuff
    function TryGetAuthToken(user, pwd : string) : string;
    function TryGetData(url : string) : string;
    //JSON
    procedure ParseAllData(json : string);
    procedure ParseEntity(jObj : TJSONObject);
    procedure ParseEntityDyn(jObj : TJSONObject; objName : string = 'noname');
    procedure ParseArray(jArray : TJSONArray; parents : string);
    procedure ParseJSONField(key : string; value : TJSONValue; parents : string);
    procedure AddToQuery(table,column,SQLval : string);
    procedure DoLocComment(JSONObj : TJSONObject);
    procedure AddJSONParam(const jObj : TJSONObject; query, objName, name, dataT : string; maxLen : integer = 0);
    //Utils
    function CreateDateTime(JSONdate : string; timeLocal : boolean = false) : string;
    function TruncStr(str : string; len : integer) : string;
    function JSONGetDataType(jVal : TJSONValue) : string;
    function JSONIsDate(jVal : string) : boolean;
    //LOC Comment
    function BuildLocComment(lat,long : double; latStr,longStr,zone : string) : string;
    function ClosestZone(lat,long : double) : string;
    function CompassPoint(LatFrom, LongFrom, LatTo, LongTo: Double): String;

    function DBLogon : string;

    //Query Building
    const insCmd = 'insert into TABLE';
    const insVals = 'values (';
    const senderMail = 'example@email.com'
  public
    function GetServiceController: TServiceController; override;
    { Public declarations }
  end;

var
  TGIConAPIsvc: TTGIConAPIsvc;
  iniFile: TiniFile;

implementation

{$R *.dfm}

uses dmSvc;

{$REGION 'TGIConAPI'}

procedure TTGIConAPIsvc.ActiveProcessesReady;
var
  aFile : TextFile;
  aFileExists: boolean;
begin
  IsDefaultEmail := False;
  LogID := 0;
  try
    with dmR.qList do
    begin
      try
        Open();
        while not(eof) do
        begin
          try
            autoRun;
          except
            on E: Exception do
              begin
                aFileExists := FileExists(logErrorName);
                AssignFile(aFile, logErrorName);
                if aFileExists then
                  System.Append(aFile)
                else
                  ReWrite(aFile);
                WriteLn(aFile, DateTimeToStr(Now)+' TGIConAPI log - Failed to autoRun');
                WriteLn(aFile, 'could not complete on server: '+ExeInfo.ComputerName);
                WriteLn(aFile, 'Error: '+ E.Message);
                CloseFile(aFile);
                GraphMail(GetSysEmail,'TGIConAPI log - Failed to autoRun',
                          ' could not complete on server: '+ExeInfo.ComputerName
                          +' - Error Code : activeProcessReady'
                          +'Error: '+ E.Message
                          ,testProfile);
              end;
          end;
          Next;
        end;
        Close;
      finally
      end;
    end;
  except
    on E: Exception do
    begin
      aFileExists := FileExists(logErrorName);
      AssignFile(aFile, logErrorName);
      if aFileExists then
        Append(aFile)
      else
        ReWrite(aFile);
      WriteLn(aFile, DateTimeToStr(Now)+' TGIConAPI log - Failed to connect to DB2');
      WriteLn(aFile, 'could not complete on server: '+ExeInfo.ComputerName);
      WriteLn(aFile, 'Error: '+ E.Message);
      CloseFile(aFile);
    end;
  end;
end;

procedure TTGIConAPIsvc.autoRun;
var
  aday,amo,ayr,ahr,amin,asec,aa: Word;
  dstr: String;
  aFile: TextFile;
  aFileExists: boolean;
  RoundTime: TDateTime;
begin
  try
    //check for a valid parameter
    logErrorName := ExtractFileDir(ParamStr(0))+'\TGIConAPIErrors.log';
    //test profile check
    testProfile := False;
    if ((AnsiLeftStr(upperCase(dmR.qList.FieldByName('NAME').Value),6) = 'TESTAL')
        or (AnsiLeftStr(upperCase(dmR.qList.FieldByName('NAME').Value),8) = 'ITTESTAW')) then
      testProfile := True;

    if ExeInfo.ComputerName = 'OIN' then
    begin
      testProfile := True;
    end;

    with dmR.qProfDtl do
	  begin
      Open();
      if RecordCount=0 then
      begin
        LogErr(0,'Invalid or expired profile: '+dmR.qList.FieldByName('NAME').Value);
        GraphMail(GetSysEmail,'TGIConAPI log - Failed to run',
          'Invalid or expired profile: '+dmR.qList.FieldByName('NAME').Value
          +' on server:'+ExeInfo.ComputerName,testProfile);
      end else
      begin
        DecodeDate(Now,ayr,amo,aday);
        DecodeTime(Now,ahr,amin,asec,aa);
        dstr:=IntToStr(ayr)+Format('%.*d',[2,amo])+Format('%.*d',[2,aday])+'-'
          +Format('%.*d',[2,ahr])+Format('%.*d',[2,amin])+Format('%.*d',[2,asec]);
        //set testmode
        if dmR.qProfDtl.FieldByName('TEST_MODE').Value = 'Y' then
          aTestMode := True
        else
          aTestMode := False;
        try
          with DmR.qUpdTime do
          begin
            RoundTime := EncodeDateTime(ayr,amo,aday,ahr,amin,0,0);
            ParamByName('NEXT_RUN').Value := IncMinute(RoundTime, dmR.qProfDtl.FieldByName('MINS').Value);
            ExecSQL;
          end;
          runProfile;
        except
          on E: Exception do
          begin
            LogErr(0,'Failed to complete: '+E.Message);
            GraphMail(GetSysEmail, 'TGIConAPI log - Failed to complete',
              'Profile: '+dmR.qList.FieldByName('NAME').Value
              +' could not complete on server:'+ExeInfo.ComputerName
              +' - Error Code : autoRun3'
              +' - Error: '+E.Message,testProfile);
          end;
        end;
      end;
      Close;
    end;
  except
    on E: Exception do
    begin
      aFileExists := FileExists(logErrorName);
      AssignFile(aFile, logErrorName);
      if aFileExists then
        Append(aFile)
      else
        ReWrite(aFile);
      WriteLn(aFile, DateTimeToStr(Now)+' TGIConAPI - autoRun failed');
      WriteLn(aFile, 'Profile: '+dmR.qList.FieldByName('NAME').Value);
      WriteLn(aFile, 'could not complete on server: '+ExeInfo.ComputerName);
      WriteLn(aFile, 'Error: '+ E.Message);
      CloseFile(aFile);
      GraphMail(GetSysEmail, 'TGIConAPI - autoRun failed',
        'Profile: '+dmR.qList.FieldByName('NAME').Value
        +' could not complete on server: '+ExeInfo.ComputerName
        +' - Error Code : autoRun1'
        +' - Error: '+E.Message,testProfile);
    end;
  end;
end;

procedure TTGIConAPIsvc.runProfile;
var
  aResult: string;
begin
  if checkSetup then
    begin
      LogDtl(dmR.qProfDtl.FieldByName('WEB_SERVICE_URL').Value);
      aResult := TryGetData(dmR.qProfDtl.FieldByName('WEB_SERVICE_URL').Value)
    end
  else
    aResult := '';

  //use reply
  if ContainsText(aResult, '{"statusCode":200') then
  begin
    if dmR.qProfDtl.FieldByName('TM_MODE').Value = 'WRITE' then
    begin
      qUpdLogDtl(aResult);
      ParseAllData(aResult);
    end;
  end
  else
  begin
    InsLog;
    LogErr(LogID, aResult);
  end;
end;

function TTGIConAPIsvc.checkSetup: Boolean;
begin
  result := True;
  if dmR.qProfDtl.Active = False then
    dmR.qProfDtl.Open;
  if dmR.qProfDtl.FieldByName('WEB_SERVICE_URL').Value = '' then
  begin
    result := False;
    LogErr(1,'Web Service URL is blank');
  end;
  if dmR.qProfDtl.FieldByName('COMMUNITY_CODE').Value = '' then
  begin
    result := False;
    LogErr(1,'Community Code is blank');
  end;
  if dmR.qProfDtl.FieldByName('COMPANY_CODE').Value = '' then
  begin
    result := False;
    LogErr(1,'Company Code is blank');
  end;
  if dmR.qProfDtl.FieldByName('AUTH').Value = '' then
  begin
    result := False;
    LogErr(1,'Authorization URL is blank');
  end;
  if dmR.qProfDtl.FieldByName('USER_ID').Value = '' then
  begin
    result := False;
    LogErr(1,'USER_ID is blank');
  end;
  if dmR.qProfDtl.FieldByName('PSWD').Value = '' then
  begin
    result := False;
    LogErr(1,'Password is blank');
  end;
  if result = False then
    LogErr(1,' aborting...');
end;


procedure TTGIConAPIsvc.InsLog;
var
  dtFrom: TDateTime;
begin
  dtFrom := Now;
  with dmR.qInsLog do
    begin
      ParamByName('TGI_ID').Value := dmR.qProfDtl.FieldByName('TGI_ID').Value;
      ParamByName('TX_DATE').Value := dtFrom;
      ExecSQL;
      LogDtlID := 0;
    end;
    //find log id
    with dmR.qLogID do
    begin
      ParamByName('TGI_ID').Value := dmR.qProfDtl.FieldByName('TGI_ID').Value;
      ParamByName('TX_DATE').Value := dtFrom;
      Open();
      LogID := FieldByName('LOG_ID').Value;
      Close;
    end;
end;

procedure TTGIConAPIsvc.LogDtl(sendData: String);
begin
  if dmR.qList.FieldByName('USE_LOG').Value = 'True' then
  begin
    if LogID = 0 then
      InsLog;
    with dmR.qInsLogDtl do
    begin
      ParamByName('LOG_ID').Value := LogID;
      ParamByName('SEND_DATA').Value := sendData;
      ExecSQL;
      dmR.qLogDtlID.ParamByName('LOG_ID').Value := LogID;
      dmR.qLogDtlID.Open;
      LogDtlID := dmR.qLogDtlID.FieldByName('LOGDTL_ID').Value;
      dmR.qLogDtlID.Close;
    end;
  end;
end;

procedure TTGIConAPIsvc.LogErr(eID: Integer; aError: String);
begin
  with dmR.qInsErr do
  begin
    if (eID<>0) and (LogID=0) and (dmR.qList.FieldByName('USE_LOG').Value = 'True') then
      InsLog;

    ParamByName('LOG_ID').Value := LogID;
    ParamByName('ERROR_CODE').Value := AnsiLeftStr(aError,100);
    if (eID=0) or (dmR.qList.FieldByName('USE_LOG').Value = 'True') then
      ExecSQL;
  end;
end;

procedure TTGIConAPIsvc.qUpdLogDtl(aReply: String);
begin
  if dmR.qList.FieldByName('USE_LOG').Value = 'True' then
  begin
    with dmR.qUpdLogDtl do
    begin
      ParamByName('LOGDTL_ID').Value := LogDTLID;
      ParamByName('RECD_DATA').Value := aReply;
      //ExecSQL;            /////////////////////////////////////////////////////////
    end;
  end;
end;

{$ENDREGION}

{$REGION 'Admin'}

function TTGIConAPIsvc.GetSysEmail: String;
begin
  IsDefaultEmail := False;
  try
    dmR.qOpen(dmR.qAdmin);
    if ((dmR.qAdmin.FieldByName('SMTP_USER').IsNull)
        or (ValidEmail(dmR.qAdmin.FieldByName('SMTP_USER').Value)=False)) then
    begin
      result := senderMail;
      IsDefaultEmail := True;
    end
    else
      result := dmR.qAdmin.FieldByName('SMTP_USER').Value;
  except
    //find registry setting
    try
      try
        iniFile:=TiniFile.Create(ExtractFileDir(ParamStr(0))+'TGIConAPISvc.ini');
        result := IniFile.ReadString('Admin','SMTP_USER','');
        IsDefaultEmail := True;
      finally
        inifile.Free;
      end;
    except
      result := senderMail;
      IsDefaultEmail := True;
    end;
  end;
end;

procedure TTGIConAPIsvc.SetupGraphMail(isTest : boolean);
begin
  if isTest then
    OutCtrl := TOutlookCOntroller.Create(true,true,exeInfo.UserName,false)
  else
    OutCtrl := TOutlookCOntroller.Create(mailLog,mailLogDtl,exeInfo.UserName,false);

  OutCtrl.Auth(false);
end;


procedure TTGIConAPIsvc.graphMail(aTo,aSubj,aBody : string; isTest : boolean);
var
  recipients : TArray<string>;
begin
  if OutCtrl = nil then
    SetupGraphMail(isTest);

  SetLength(recipients,2);
  Recipients[0] := aTo;
  Recipients[1] := senderMail;

  if isTest = false then
  begin
    SetLength(Recipients,length(Recipients) + 1);
    Recipients[length(Recipients) - 1] := senderMail;
  end;

  OutCtrl.SendMail(
    'html',
    senderMail,
    senderMail,
    senderMail,
    aSubj,
    aBody,
    Recipients,
    {CC}nil,
    {BCC}nil,
    {ATTACH}nil
  );
end;

{$ENDREGION}

{$REGION 'Service'}

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  TGIConAPIsvc.Controller(CtrlCode);
end;

function TTGIConAPIsvc.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;

procedure TTGIConAPIsvc.ServiceAfterInstall(Sender: TService);
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('\SYSTEM\CurrentControlSet\Services\' + Name, false) then
    begin
      Reg.WriteString('Description', 'TGIConAPI Service');
      Reg.CloseKey;
    end;
  finally
    Reg.Free;
  end;
end;

procedure TTGIConAPIsvc.ServiceExecute(Sender: TService);
const
  SecBetweenRuns = 10;     
var
  Count: Integer;
  Reg: TRegistry;
begin
  Count := 0;

  //new - Variables for building SQL queries
  insDataFlds := insCmd + 'DATA(';
  insDataVals := insVals;
  insDataFirstFld := true;

  insArrFlds := insCmd;
  insArrVals := insVals;
  insArrFirstFld := true;
  currArrTable := 'null';

  var ret := DBLogon;
  if ret = '' then
  begin
    while not Terminated do
    begin
      Inc(Count);
      if Count >= SecBetweenRuns then
      begin
        Count := 0;

        ActiveProcessesReady;

      end;
      Sleep(1000);
      ServiceThread.ProcessRequests(False);
    end;
  end
  else
  begin
    GraphMail(senderMail,'TGIConAPI - Configuration Error',
              'A database connection was not able to be established on startup. Message: ' + ret,testProfile);
  end;
end;

{$ENDREGION}

{$REGION 'API functions'}

function TTGIConAPIsvc.TryGetAuthToken(user, pwd : string) : string;
var
  HTTP : TidHTTP;
  SSL : TidSSLIOHandlerSocketOpenSSL;
  jObj : TJSONObject;
  ret : string;
  statusCode : integer;
begin

  //Create HTTP client
  HTTP := TidHTTP.Create;
  HTTP.HTTPOptions := HTTP.HTTPOptions + [hoNoProtocolErrorException] + [hoWantProtocolErrorContent];

  SSL := TidSSLIOHandlerSocketOpenSSL.Create;
  SSL.SSLOptions.Method := TidSSLVersion(1);
  HTTP.IOHandler := SSL;

  //Create headers
  HTTP.Request.Username := user;
  HTTP.Request.Password := pwd;
  HTTP.Request.BasicAuthentication := true;

  HTTP.Request.CustomHeaders.Values['X-Community-Code']
    := dmR.qProfDtl.FieldByName('COMMUNITY_CODE').Value;
  HTTP.Request.CustomHeaders.Values['X-Company-Code']
    := dmR.qProfDtl.FieldByName('COMPANY_CODE').Value;

  try
    ret := HTTP.Get(dmR.qProfDtl.FieldByName('AUTH').Value);
  finally
    HTTP.Free;
    SSL.Free;
  end;

  if ret <> '' then
  begin
    if ret.Contains('{') and ret.Contains('}') then
    begin
      //parse json
      jObj := TJSONObject.ParseJSONValue(ret) as TJSONObject;

      if jObj.FindValue('statusCode') <> nil then
      begin
        statusCode := (jObj.GetValue('statusCode') as TJSONNumber).AsInt;
      //check for error codes
        result := 'Error ' + inttostr(statusCode)
                + ':' + jObj.Get('message').JsonValue.Value;
      end
      else if jObj.FindValue('token') <> nil then
      begin
        token := jObj.Get('token').JsonValue.Value;
        tokenExpire := IncMinute(Now,59);

        result := '200';
      end
      else
        result := 'Error - no code: ' + ret;

      jObj.Free;
    end
    else
      result := 'Error 0: Response not in JSON (Response: ' + ret + ')';
  end
  else
    //if no response at all. This probably wont ever happen
    result := 'Error 0: No response';
end;


function TTGIConAPIsvc.TryGetData(url : string) : string;
var
  HTTP : TidHTTP;
  SSL : TidSSLIOHandlerSocketOpenSSL;
  authRes : string;
  authOk : boolean;
  jObj : TJSONObject;
  statusCode : integer;
begin
  //Create HTTP client
  HTTP := TidHTTP.Create;
  HTTP.HTTPOptions := HTTP.HTTPOptions + [hoNoProtocolErrorException] + [hoWantProtocolErrorContent];

  SSL := TidSSLIOHandlerSocketOpenSSL.Create;
  SSL.SSLOptions.Method := TidSSLVersion(1);
  HTTP.IOHandler := SSL;

  authOk := true;
  if (token = '') or (Now > tokenExpire) then
  begin
    authRes := TryGetAuthToken(dmR.qProfDtl.FieldByName('USER_ID').Value,
               dmR.qProfDtl.FieldByName('PSWD').Value);
    authOk := authRes.Contains('200');
    result := 'Unable to get Auth Token. Error - ' + authRes;
  end;

  if authOk then
  begin
    HTTP.Request.CustomHeaders.Values['Authorization'] := 'Bearer ' + token;
    //HTTP.Request.CustomHeaders.Add('Authorization:Bearer '+ token);

    HTTP.Request.CustomHeaders.Values['X-Community-Code']
      := dmR.qProfDtl.FieldByName('COMMUNITY_CODE').Value;
    HTTP.Request.CustomHeaders.Values['X-Company-Code']
      := dmR.qProfDtl.FieldByName('COMPANY_CODE').Value;

    try
      result := HTTP.Get(url);
      if result.Contains('{') and result.Contains('}') then
      begin
        jObj := jObj.ParseJSONValue(result) as TJSONObject;

        if jObj.FindValue('statusCode') <> nil then
        begin
          statusCode :=(jObj.GetValue('statusCode') as TJSONNumber).AsInt;
          if statusCode = 401 then
          begin
            authRes := TryGetAuthToken(dmR.qProfDtl.FieldByName('USER_ID').Value,
                       dmR.qProfDtl.FieldByName('PSWD').Value);

            if authRes.Contains('200') then
            begin
              HTTP.Request.CustomHeaders.Values['Authorization'] := 'Bearer ' + token;
              result := HTTP.Get(url);
            end
            else
              result := 'Unable to get Auth Token. Error - ' + authRes;
          end
          else if statusCode <> 200 then
            result := inttostr(statusCode) + ':' + jObj.Get('message').Value;
        end
        else
        begin
          if (jObj.Get('message').Value = 'Public key not found in jwks.json')
          or (jObj.Get('message').Value.Contains('Unexpected token')) then
          begin
            authRes := TryGetAuthToken(dmR.qProfDtl.FieldByName('USER_ID').Value,
                       dmR.qProfDtl.FieldByName('PSWD').Value);

            if authRes.Contains('200') then
            begin
              HTTP.Request.CustomHeaders.Values['Authorization'] := 'Bearer ' + token;
              result := HTTP.Get(url);
            end
            else
              result := 'Unable to get Auth Token. Error - ' + authRes;
          end
        end;

        jObj.Free;
      end
      else
        //for when response is not JSON. This likely will not ever happen
        result := 'Error 0: No response';

    finally
      HTTP.Free;
      SSL.Free;
    end;
  end
  else
  begin
    HTTP.Free;
    SSL.Free;
  end;

end;

{$ENDREGION}

{$REGION 'JSON'}

procedure TTGIConAPIsvc.ParseAllData(json : string);
var
  I : integer;
  jObj : TJSONObject;
  jArr : TJSONArray;
  outTxt : TStreamWriter;
  dirStr, strEmlBody, strEmlSub : String;
begin
  //Write JSON to a text file
  dmR.qOpen(dmR.qAdmin);
  dirStr := dmR.qAdmin.FieldByName('LOG_DIR').Value + '\'
    + AnsiReplaceStr(dmR.qProfDtl.FieldByName('NAME').Value,' ','');
  if DirectoryExists(dirStr) = False then
    ForceDirectories(dirStr);

  outTxt := TStreamWriter.Create(
    dirStr + '\' +
    'JSON_' +
    inttostr(yearof(now)) + '-' +
    inttostr(Monthof(now)) + '-' +
    inttostr(dayof(now)) + '_' +
    inttostr(hourof(now)) + '-' +
    inttostr(minuteof(now)) + '-' +
    inttostr(secondof(now)) + '.txt'
  );
  try
    outTxt.Write(json);

    //parse json
    jObj := TJSONObject.Create;
    jObj := TJSONObject.ParseJSONValue(json) as TJSONObject;
    jArr := jObj.GetValue('data') as TJSONArray;

    for I := 0 to (jArr.Count - 1) do
    begin
      if (jArr[I] as TJSONObject).FindValue('id') <> nil then
      begin
        currID := TruncStr((jArr[I] as TJSONObject).GetValue('id').ToString.Trim(['"']),36);
        dmR.qDATAGetID.ParamByName('ENTITYID').Value := currID;
        dmR.qDATAGetID.Open;

        if dmR.qDATAGetID.RecordCount < 1 then
        begin
          ParseEntityDyn(jArr[I] as TJSONObject);

          //Do actual insert here
          dmR.qIns.SQL.Clear;
          dmR.qIns.SQL.Add(insDataFlds + ')');
          dmR.qIns.SQL.Add(insDataVals + ')');
          dmR.qIns.ExecSQL;

          insDataFlds := insCmd + 'DATA(';
          insDataVals := insVals;
          insDataFirstFld := true;
          currID := '';
        end;

        dmR.qDATAGetID.Close;
      end;
    end;

    //do emails here. then set any flags to 'False' that are true
    dmR.qGetSchema.Filter := 'NEW_FIELD_ALERT = ''True''';
    dmR.qGetSchema.Filtered := true;
    for I := 0 to (dmR.qGetSchema.RecordCount - 1) do
    begin
      if I <> 0 then
        emailNewFlds := emailNewFlds + ', '
      else
        strEmlSub := 'NEW FIELD ALERT';
      emailNewFlds := emailNewFlds + dmR.qGetSchemaJSON_FIELD.Value;
      dmR.qGetSchema.Edit;
      dmR.qGetSchemaNEW_FIELD_ALERT.Value := 'False';
      dmR.qGetSchema.Post;
      dmR.qGetSchema.Next;
    end;

    dmR.qGetSchema.Filter := 'DATA_TYPE_ALERT = ''True''';
    for I := 0 to (dmR.qGetSchema.RecordCount - 1) do
    begin
      if I <> 0 then
        emailDataAlert := emailDataAlert + ', '
      else
      begin
        if strEmlSub = '' then
          strEmlSub := 'DATA TYPE ALERT'
        else
          strEmlSub := strEmlSub + ' & DATA TYPE ALERT';
      end;
      emailDataAlert := emailDataAlert + dmR.qGetSchemaJSON_FIELD.Value;
      dmR.qGetSchema.Edit;
      dmR.qGetSchemaDATA_TYPE_ALERT.Value := 'False';
      dmR.qGetSchema.Post;
      dmR.qGetSchema.Next;
    end;
    dmR.qGetSchema.Filtered := false;

    if emailNewFlds <> '' then
      strEmlBody := 'New fields: ' + slinebreak + emailNewFlds + slinebreak + slinebreak;

    if emailDataAlert <> '' then
      strEmlBody := strEmlBody + 'Data Type Alerts: ' + slinebreak + emailDataAlert;

    if strEmlBody <> '' then
    begin
      with dmR do
      begin
        GraphMail(GetSysEmail,'TGIConAPI - ' + strEmlSub, strEmlBody,testProfile);
      end;
    end;

    emailNewFlds := '';
    emailDataAlert := '';

  finally
    outTxt.Free;
    jObj.Free;
  end;
end;


procedure TTGIConAPIsvc.ParseEntityDyn(jObj : TJSONObject; objName : string = 'noname');
var
  I : integer;
  parentString, thisKey : string;
begin
  if objName = '' then
    parentString := 'noname'
  else
    parentString := objName;

  for I := 0 to jObj.Count - 1 do
  begin
    if jObj.Pairs[I].JsonString.Value = '' then
      thisKey := 'noname'
    else
      thisKey := jObj.Pairs[I].JsonString.Value;

    if jObj.Pairs[I].JsonValue is TJSONObject then
    begin
      if jObj.Pairs[I].JsonString.Value = 'deviceEvent' then
        DoLocComment(jObj.Pairs[I].JsonValue as TJSONObject);

      ParseEntityDyn(jObj.Pairs[I].JsonValue as TJSONObject, parentString + '_' + thisKey);
    end
    else if jObj.Pairs[I].JsonValue is TJSONArray then
    begin
      ParseArray(jObj.Pairs[I].JsonValue as TJSONArray, parentString + '_' + thisKey);
    end
    else
    begin
      if (thisKey = 'timestampUtc') and (parentString <> 'noname_weather_properties') then
        if ContainsStr(insDataFlds, 'TIMESTAMPSERVER') = false then
          AddToQuery('Table','TIMESTAMPSERVER',CreateDateTime(jObj.Pairs[I].JsonValue.Value, true));

      ParseJSONField(thisKey,jObj.Pairs[I].JsonValue,parentString);
    end;

  end;

end;


procedure TTGIConAPIsvc.ParseArray(jArray : TJSONArray; parents : string);
var
  I : integer;
  thisName : string;
begin
  for I := 0 to jArray.Count - 1 do
  begin
    if jArray.Items[I].Value = '' then
      thisName := 'noname'
    else
      thisName := jArray.Items[I].Value;

    if jArray.Items[I] is TJSONObject then
    begin
      ParseEntityDyn(jArray[I] as TJSONObject, parents + '_' + thisName);
    end
    else if jArray.Items[I] is TJSONArray then
    begin
      ParseArray(jArray[I] as TJSONArray, parents + '_' + thisName);
    end;

    //NEED TO EXECUTE THE INSERT HERE SO EACH OBJECT IN AN ARRAY IS A SEPARATE ENTRY IN THE TABLE
    if currArrTable <> 'null' then
    begin
      insArrFlds := insArrFlds + ',id';
      insArrVals := insArrVals + ',' + QuotedStr(currID);

      dmR.qIns.SQL.Clear;
      dmR.qIns.SQL.Add(insArrFlds + ')');
      dmR.qIns.SQL.Add(insArrVals + ')');

      dmR.qIns.ExecSQL;

      currArrTable := 'null';
      insArrFirstFld := true;
      insArrFlds := insCmd;
      insArrVals := insVals;
    end;
  end;
end;


procedure TTGIConAPIsvc.ParseJSONField(key : string; value : TJSONValue; parents : string);
var
  fieldName, tableName, tableCol, dataType, JSONDataType, SQLVal, temp : string;
  dataSize : integer;
begin
  fieldName := StringReplace(parents,'noname_','',[rfReplaceAll]);
  fieldName := StringReplace(fieldName,'_noname','',[rfReplaceAll]);
  fieldName := StringReplace(fieldName,'noname','',[rfReplaceAll]);
  if fieldName <> '' then
    fieldName := fieldName + '_';

  fieldName := fieldName + key;
  //showMessage('Field name is '  + fieldName);

  if dmR.qGetSchema.Active = false then
    dmR.qGetSchema.Open;

  dmR.qGetSchema.Filter := 'JSON_FIELD = ''' + fieldName + '''';
  dmR.qGetSchema.Filtered := true;

  if dmR.qGetSchema.RecordCount = 0 then
  begin
    //showMessage('Field ' + fieldName + ' does not exist in table');

    //insert a new entry into the table with fieldName as JSON_FIELD and set NEW_FIELD_ALERT = 'True'
    dmR.qInsSchemaField.ParamByName('JSON_FIELD').Value := fieldName;
    dmR.qInsSchemaField.ExecSQL;
  end
  else
  begin
    //check if there is a table and column specified. if so continue
    tableName := 'schema.'+dmR.qGetSchemaTM_TABLE_NAME.Value; //schema name removed
    tableCol := dmR.qGetSchemaTM_TABLE_COL.Value;
    if  ((tableName <> '') and (tableCol <> '')) then
    begin
      dmR.qGetColInfo.Filter := 'TABNAME = ' + QuotedStr(tableName) + ' and COLNAME = ' + QuotedStr(uppercase(tableCol));
      dmR.qGetColInfo.Filtered := True;
      dmR.qGetColInfo.Open;

      if dmR.qGetColInfo.RecordCount > 0 then
      begin
        dataType := dmR.qGetColInfo.FieldByName('TYPENAME').Value;
        dataSize := dmR.qGetColInfo.FieldByName('LENGTH').Value;

        //check the data type and size of the column. if we cant make it work set data type alert to 'true'
        JSONDataType := JSONGetDataType(value);

        if (dataType = 'VARCHAR') and (JSONDataType = 'string') then
        begin
          if length(value.Value) > dataSize then
          begin
            //if just size is wrong still set alert but truncate value and use it
            //set alert
            dmR.qSchemaSetDTAlert.Params[0].Value := fieldName;
            temp := value.Value;
            dmR.qSchemaSetDTAlert.ExecSQL;
            SQLVal := QuotedStr(TruncStr(value.Value,dataSize));
          end
          else
            SQLVal := QuotedStr(value.Value);
        end
        else if ((dataType = 'INTEGER') or  (dataType = 'DOUBLE')) and (JSONDataType = 'number') then
        begin
          SQLVal := value.Value;
        end
        else if (dataType = 'TIMESTAMP') and (JSONDataType = 'date') then
        begin
          SQLVal := CreateDateTime(value.Value)
        end
        else if JSONDataType = 'null' then
        begin
          SQLVal := 'nilnilnil'
        end
        else
        begin
          SQLVal := 'nilnilnil';
          //set alert
          dmR.qSchemaSetDTAlert.Params[0].Value := fieldName;
          temp := value.Value;
          dmR.qSchemaSetDTAlert.ExecSQL;
        end;

        if SQLVal <> 'nilnilnil' then
        begin
          AddToQuery(tableName,tableCol,SQLVal);
        end;

      end;

      dmR.qGetColInfo.Close;
      dmR.qGetColInfo.Filter := '';
      dmR.qgetColInfo.Filtered := False;

    end;
  end;

  dmR.qGetSchema.Filtered := false;

end;


procedure TTGIConAPIsvc.AddToQuery(table,column,SQLval : string);
begin
  if table = 'table' then //table name removed 
  begin
    if insDataFirstFld = false then
    begin
      insDataFlds := insDataFlds + ',';
      insDataVals := insDataVals + ',';
    end
    else
      insDataFirstFld := False;

    insDataFlds := insDataFlds + column;
    insDataVals := insDataVals + SQLval;
  end
  else if (table = 'landmark table') or (table = 'weather metadata table') then //table names removed
  begin

    if currArrTable = 'null' then
    begin
      currArrTable := table; 
      insArrFlds := insArrFlds + replaceStr(table,'table','') + '('; //table name removed
    end;

    if insArrFirstFld = false then
    begin
      insArrFlds := insArrFlds + ',';
      insArrVals := insArrVals + ',';
    end
    else
      insArrFirstFld := False;

    insArrFlds := insArrFlds + column;
    insArrVals := insArrVals + SQLval;
  end;
end;




//deprecated
procedure TTGIConAPIsvc.ParseEntity(jObj : TJSONObject);
var
  I : integer;
  jSubObj : TJSONObject;
  jSubObj2 : TJSONObject;
  jArr : TJSONArray;
  lat, long : double;
  zoneStr : string;
  splitStr : TArray<string>;
begin
  with dmR do
  begin
    qDATAGetID.ParamByName('ENTITYID').Value := TruncStr(jObj.GetValue('id').ToString.Trim(['"']),36);
    qDATAGetID.Open;

    if qDATAGetID.RecordCount < 1 then
    begin
      AddJSONParam(jObj,'DATA','','id','STR',36);
      AddJSONParam(jObj,'DATA','','timestampUtc','DATE');
      qInsDATA.ParamByName('TIMESTAMPSERVER').Value := CreateDateTime(jObj.GetValue('timestampUtc').ToString.Trim(['"']),true);

      //asset JSON object
      jSubObj := jObj.GetValue('asset') as TJSONObject;

      AddJSONParam(jSubObj,'DATA','ASSET','companyCode','STR',10);
      AddJSONParam(jSubObj,'DATA','ASSET','assetId','STR',10);
      AddJSONParam(jSubObj,'DATA','ASSET','esn','STR',20);

      //asset->properties JSON object
      jSubObj := jSubObj.GetValue('properties') as TJSONObject;

      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','status','STR',10);
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','type','STR',20);
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','group','STR',20);
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','subGroup','STR',20);

      //NEW NEW NEW
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','companyName','STR',40);
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','description','STR',80);
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','purchaseDate','DATE');
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','fleetDate','DATE');
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','retireDate','DATE');
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','activationDate','DATE');
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','plateNumber','STR',20);
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','serialNumber','STR',40);
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','vin','STR',17);
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','manufacture','STR',20);
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','model','STR',20);
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','modelYear','INT');
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','userDate1','DATE');
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','userDate2','DATE');
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','userDate3','DATE');
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','userAmount1','FLOAT');
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','userAmount2','FLOAT');
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','userAmount3','FLOAT');
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','userCode1','STR',40);
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','userCode2','STR',40);
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','userCode3','STR',40);
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','userText1','STR',80);
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','userText2','STR',80);
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','userText3','STR',80);
      AddJSONParam(jSubObj,'DATA','ASSET_PROPERTIES','comment','STR',100);
      //NEW NEW NEW

      //deviceEvent JSON object
      jSubObj := jObj.GetValue('deviceEvent') as TJSONObject;

      AddJSONParam(jSubObj,'DATA','DEVICEEVENT','messageType','STR',10);
      AddJSONParam(jSubObj,'DATA','DEVICEEVENT','batteryHealthy','STR',5);
      AddJSONParam(jSubObj,'DATA','DEVICEEVENT','gpsDataValid','STR',5);
      AddJSONParam(jSubObj,'DATA','DEVICEEVENT','input1Missed','STR',5);
      AddJSONParam(jSubObj,'DATA','DEVICEEVENT','input2Missed','STR',5);
      AddJSONParam(jSubObj,'DATA','DEVICEEVENT','gpsFailCounter','INT');
      AddJSONParam(jSubObj,'DATA','DEVICEEVENT','latitude','FLOAT');
      AddJSONParam(jSubObj,'DATA','DEVICEEVENT','longitude','FLOAT');
      AddJSONParam(jSubObj,'DATA','DEVICEEVENT','input1TriggeredMessage','STR',5);
      AddJSONParam(jSubObj,'DATA','DEVICEEVENT','input1StateOpen','STR',5);
      AddJSONParam(jSubObj,'DATA','DEVICEEVENT','input2TriggeredMessage','STR',5);
      AddJSONParam(jSubObj,'DATA','DEVICEEVENT','input2StateOpen','STR',5);
      AddJSONParam(jSubObj,'DATA','DEVICEEVENT','subType','STR',10);
      AddJSONParam(jSubObj,'DATA','DEVICEEVENT','vibrationTriggeredMessage','STR',5);
      AddJSONParam(jSubObj,'DATA','DEVICEEVENT','vibrationStateIsActive','STR',5);
      AddJSONParam(jSubObj,'DATA','DEVICEEVENT','gpsDataIsFrom3dFix','STR',5);
      AddJSONParam(jSubObj,'DATA','DEVICEEVENT','deviceInMotion','STR',5);
      AddJSONParam(jSubObj,'DATA','DEVICEEVENT','gpsHighConfidence','STR',5);
      AddJSONParam(jSubObj,'DATA','DEVICEEVENT','esn','STR',5);

      //Do stuff for location comment and zone code
      if (jSubObj.FindValue('latitude') <> nil) and (jSubObj.FindValue('longitude') <> nil) then
      begin
        lat := strtofloat(jSubObj.GetValue('latitude').ToString.Trim(['"']));
        long := strtofloat(jSubObj.GetValue('longitude').ToString.Trim(['"']));
        zoneStr := ClosestZone(lat, long);
        splitStr := zoneStr.Split(['_']);

        qInsDATA.ParamByName('TGI_CLOSEST_ZONE').Value := splitStr[0];
        qInsDATA.ParamByName('TGI_TMWIN_LOC').Value := BuildLocComment(lat,long,splitStr[1],splitStr[2],splitStr[0]);
      end;
      //location->properties JSON object
      jSubObj := (jObj.GetValue('location') as TJSONObject).GetValue('properties') as TJSONObject;

      AddJSONParam(jSubObj,'DATA','LOCATION_PROPERTIES','road','STR',40);
      AddJSONParam(jSubObj,'DATA','LOCATION_PROPERTIES','city','STR',30);
      AddJSONParam(jSubObj,'DATA','LOCATION_PROPERTIES','state','STR',20);
      AddJSONParam(jSubObj,'DATA','LOCATION_PROPERTIES','countryCode','STR',2);
      AddJSONParam(jSubObj,'DATA','LOCATION_PROPERTIES','country','STR',20);
      AddJSONParam(jSubObj,'DATA','LOCATION_PROPERTIES','postalCode','STR',10);

      //stats JSON object
      jSubObj := jObj.GetValue('stats') as TJSONObject;

      //stats->distance JSON object
      jSubObj2 := jSubObj.GetValue('distance') as TJSONObject;

      AddJSONParam(jSubObj2,'DATA','STATS_DISTANCE','m','FLOAT');
      AddJSONParam(jSubObj2,'DATA','STATS_DISTANCE','km','FLOAT');
      AddJSONParam(jSubObj2,'DATA','STATS_DISTANCE','mi','FLOAT');


      //stats->speed JSON object
      jSubObj2 := jSubObj.GetValue('speed') as TJSONObject;

      AddJSONParam(jSubObj2,'DATA','STATS_SPEED','mph','FLOAT');
      AddJSONParam(jSubObj2,'DATA','STATS_SPEED','kph','FLOAT');

      //stats->heading JSON object
      jSubObj2 := jSubObj.GetValue('heading') as TJSONObject;

      AddJSONParam(jSubObj2,'DATA','STATS_HEADING','compass','STR',3);
      AddJSONParam(jSubObj2,'DATA','STATS_HEADING','bearing','FLOAT');

      //stats->odometer JSON object
      jSubObj2 := jSubObj.GetValue('odometer') as TJSONObject;

      AddJSONParam(jSubObj2,'DATA','STATS_ODOMETER','m','FLOAT');
      AddJSONParam(jSubObj2,'DATA','STATS_ODOMETER','km','FLOAT');
      AddJSONParam(jSubObj2,'DATA','STATS_ODOMETER','mi','FLOAT');

      //NEW NEW NEW
      //stats->lifeTimeOdometer
      jSubObj2 := jSubObj.GetValue('lifeTimeOdometer') as TJSONObject;

      AddJSONParam(jSubObj2,'DATA','STATS_LIFETIMEODOMETER','km','FLOAT');
      AddJSONParam(jSubObj2,'DATA','STATS_LIFETIMEODOMETER','mi','FLOAT');
      //NEW NEW NEW

      //weather->properties JSON object
      jSubObj := (jObj.GetValue('weather') as TJSONObject).GetValue('properties') as TJSONObject;

      AddJSONParam(jSubObj,'DATA','WEATHER_PROPERTIES','summary','STR',20);
      AddJSONParam(jSubObj,'DATA','WEATHER_PROPERTIES','pressure','INT');
      AddJSONParam(jSubObj,'DATA','WEATHER_PROPERTIES','humidity','INT');
      AddJSONParam(jSubObj,'DATA','WEATHER_PROPERTIES','sunriseUtc','DATE');
      AddJSONParam(jSubObj,'DATA','WEATHER_PROPERTIES','sunsetUtc','DATE');
      AddJSONParam(jSubObj,'DATA','WEATHER_PROPERTIES','timestampUtc','DATE');

      //TODO
      //if cloudCover exists then we can add that too

      //weather->properties->temperature JSON object
      if jSubObj.FindValue('temperature') <> nil then
      begin
        jSubObj2 := jSubObj.GetValue('temperature') as TJSONObject;

        AddJSONParam(jSubObj2,'DATA','WEATHER_PROPERTIES_TEMPERATURE','c','INT');
        AddJSONParam(jSubObj2,'DATA','WEATHER_PROPERTIES_TEMPERATURE','f','INT');
      end;

      //weather->properties->feelsLike JSON object
      jSubObj2 := jSubObj.GetValue('feelsLike') as TJSONObject;

      AddJSONParam(jSubObj2,'DATA','WEATHER_PROPERTIES_FEELSLIKE','c','INT');
      AddJSONParam(jSubObj2,'DATA','WEATHER_PROPERTIES_FEELSLIKE','f','INT');

      //weather->properties->wind->speed JSON object
      jSubObj2 := (jSubObj.GetValue('wind') as TJSONObject).GetValue('speed') as TJSONObject;

      AddJSONParam(jSubObj2,'DATA','WEATHER_PROPERTIES_WIND_SPEED','mph','INT');
      AddJSONParam(jSubObj2,'DATA','WEATHER_PROPERTIES_WIND_SPEED','kph','INT');

      //TODO
      //if wind->gust exists the we should grab that too


      qInsDATA.ExecSQL;
      qInsDATA.Close;

      //weather->properties->metadata JSON array
      if jSubObj.FindValue('metadata') <> nil then
      begin
        jArr := jSubObj.GetValue('metadata') as TJSONArray;

        for I := 0 to (jArr.Count - 1) do
        begin
          //weather->properties->metadata[] object
          jSubObj2 := jArr[I] as TJSONObject;

          AddJSONParam(jObj,'WEATHER','','id','STR',36);

          AddJSONParam(jSubObj2,'WEATHER','METADATA','id','INT');
          AddJSONParam(jSubObj2,'WEATHER','METADATA','main','STR',10);
          AddJSONParam(jSubObj2,'WEATHER','METADATA','description','STR',20);
          AddJSONParam(jSubObj2,'WEATHER','METADATA','icon','STR',5);

          qInsWEATHER.ExecSQL;
          qInsWEATHER.Close;

        end;
      end;

      //landmarks JSON array
      if jSubObj.FindValue('landmarks') <> nil then
      begin
        jArr := jObj.GetValue('landmarks') as TJSONArray;

        for I := 0 to (jArr.Count - 1) do
        begin
          //landmarks[] JSON object
          jSubObj := jArr[I] as TJSONObject;
          //landmarks[]->properties JSON object
          jSubObj2 := jSubObj.GetValue('properties') as TJSONObject;

          AddJSONParam(jObj,'LANDMARKS','','id','STR',36);

          //qInsLANDMARKS.ParamByName('LANDMARKID').Value := strtoint(jSubObj.GetValue('landmarkId').ToString.Trim(['"']));
          AddJSONParam(jSubObj,'LANDMARKS','','landmarkId','STR',36);
          AddJSONParam(jSubObj,'LANDMARKS','','name','STR',40);

          AddJSONParam(jSubObj2,'LANDMARKS','PROPERTIES','status','STR',10);
          AddJSONParam(jSubObj2,'LANDMARKS','PROPERTIES','landmarkType','STR',10);
          AddJSONParam(jSubObj2,'LANDMARKS','PROPERTIES','detentionType','STR',10);
          AddJSONParam(jSubObj2,'LANDMARKS','PROPERTIES','geoFenceType','STR',10);

          if jSubObj2.GetValue('geoFenceType').ToString = '"Circle"' then
            AddJSONParam(jSubObj2,'LANDMARKS','PROPERTIES','radius','INT');

          AddJSONParam(jSubObj2,'LANDMARKS','PROPERTIES','lastUpdatedDate','DATE');
          AddJSONParam(jSubObj2,'LANDMARKS','PROPERTIES','billable','STR',5);
          AddJSONParam(jSubObj2,'LANDMARKS','PROPERTIES','address','STR',40);
          AddJSONParam(jSubObj2,'LANDMARKS','PROPERTIES','subType','STR',10);

          qInsLANDMARKS.ExecSQL;
          qInsLANDMARKS.Close;

        end;
      end;
    end;

    qDATAGetID.Close;
  end;
end;

//deprecated
procedure TTGIConAPIsvc.AddJSONParam(const jObj : TJSONObject; query, objName, name, dataT : string; maxLen : integer = 0);
var
  str, paramName : string;
  val : variant;
begin
  if jObj.FindValue(name) <> nil then
  begin

    str := jObj.GetValue(name).ToString.Trim(['"']);

    if dataT = 'INT' then
      if (str = 'null') or (str = '') then
        val := 0
      else
        val := strtoint(str)
    else if dataT = 'FLOAT' then
      if (str = 'null') or (str = '') then
        val := 0.0
      else
        val := strtofloat(str)
    else if dataT = 'DATE' then
      if (str = 'null') or (str = '') then
        val := CreateDateTime('1980-01-01T00:00:00.000Z')
      else
        val := CreateDateTime(str)
    else if dataT = 'DATELOC' then
      if (str = 'null') or (str = '') then
        val := CreateDateTime('1980-01-01T00:00:00.000Z', true)
      else
        val := CreateDateTime(str, true)
    else
    begin
      if maxLen > 0 then
        if str = 'null' then
          str := 'null'
        else
          str := TruncStr(str,maxLen);

      val := str;
    end;

    paramName := uppercase(objName);
    if objName.Length = 0 then
      paramName := uppercase(name)
    else
      paramName := paramName + '_' + uppercase(name);

    if query = 'DATA' then
    begin
      dmR.qInsDATA.ParamByName(paramName).Value := val;
    end
    else if query = 'WEATHER' then
    begin
      dmR.qInsWEATHER.ParamByName(paramName).Value := val;
    end
    else if query = 'LANDMARKS' then
    begin
      dmR.qInsLANDMARKS.ParamByName(paramName).Value := val;
    end

  end;
end;

{$ENDREGION}

{$REGION 'Utils'}

function TTGIConAPIsvc.CreateDateTime(JSONdate : string; timeLocal : boolean = false) : string;
var
  splitStr,splitTime : TArray<string>;
  diffUTC,monthNow,dayNow : integer;
  tZone : TTimeZone;
  compDate,currDate :TDateTime;
begin
  //probably want this in the profile, or at least adjusting for daylight savings time
  tZone := TTimeZone.Local;
  diffUTC := tZone.GetUtcOffset(now).Hours;
  //do dst time
  //need to find out what the dates are
  currDate := now;
  monthNow := monthof(currDate);
  dayNow := dayof(currDate);
  if monthNow in [1,2,12] then
    diffUTC := diffUTC - 1
  else if (monthNow = 3) or (monthNow = 11) then
  begin
    compDate := EncodeDateTime(yearOf(currDate),monthNow,
                ( 7 - (dayOfweek(StartOfTheMonth(currDate)) - 1) + 1 ),
                2,0,0,0 );
    if monthNow = 3 then
    begin
      compDate := incWeek(compDate,1);
      if compareDate(currDate,compDate) < 0 then
        diffUTC := diffUTC - 1;
    end
    else
    begin
      if compareDate(currDate,compDate) >= 0 then
        diffUTC := diffUTC - 1;
    end;
  end;

  splitStr := JSONdate.Split(['T']);

  splitTime := ReplaceStr(splitStr[1],'Z','').Split([':','.']);

  result := splitStr[0] + ' ';

  if timeLocal then
  begin
    result := result + inttostr(strtoint(splitTime[0]) + diffUTC);
  end
  else
  begin
    result := result + splitTime[0];
  end;

  result := result + ':' + splitTime[1] + ':' + splitTime[2];

  result := QuotedStr(result);
end;


function TTGIConAPIsvc.TruncStr(str : string; len : integer) : string;
begin
  result := str;
  if length(str) > len then
    SetLength(str,len);
  result := str;
end;


function TTGIConAPIsvc.JSONGetDataType(jVal : TJSONValue) : string;
begin
  if jVal.Value = 'null' then
    result := 'null'
  else
  begin
    if jVal is TJSONNumber then
      result := 'number'
    else
    begin
      if JSONIsDate(jVal.Value) then
        result := 'date'
      else
        result := 'string'
    end;
  end;
end;


function TTGIConAPIsvc.JSONIsDate(jVal : string) : boolean;
var
  regEx : TRegEx;
  res : TMatch;
begin
  result := false;

  regEx := TRegEx.Create('\d\d\d\d-\d?\d-\d?\dT\d\d:\d\d:\d\d\.?\d?\d?\d?Z');

  res := regEx.Match(jVal);

  if res.Success then
    result := true

end;


function TTGIConAPIsvc.ValidEmail(email: string): boolean;
const
  atom_chars = [#33..#255] - ['(', ')', '<', '>', '@', ',', ';', ':',
                                '\', '/', '"', '.', '[', ']', #127];
  quoted_string_chars = [#0..#255] - ['"', #13, '\'];
  letters = ['A'..'Z', 'a'..'z'];
  letters_digits = ['0'..'9', 'A'..'Z', 'a'..'z'];
  subdomain_chars = ['-', '0'..'9', 'A'..'Z', 'a'..'z'];
type
  States = (STATE_BEGIN, STATE_ATOM, STATE_QTEXT, STATE_QCHAR,
      STATE_QUOTE, STATE_LOCAL_PERIOD, STATE_EXPECTING_SUBDOMAIN,
      STATE_SUBDOMAIN, STATE_HYPHEN);
var
  State: States;
  i, n, subdomains: integer;
  c: char;
begin
  State := STATE_BEGIN;
  n := Length(email);
  i := 1;
  subdomains := 1;
  while (i <= n) do begin
    c := email[i];
    case State of      STATE_BEGIN:
      if CharInSet(c, atom_chars) then
        State := STATE_ATOM
      else if c = '"' then
        State := STATE_QTEXT
      else          break;
    STATE_ATOM:
      if c = '@' then
        State := STATE_EXPECTING_SUBDOMAIN
      else if c = '.' then
        State := STATE_LOCAL_PERIOD
      else if not (CharInSet(c, atom_chars)) then
        break;
    STATE_QTEXT:
      if c = '\' then
        State := STATE_QCHAR
      else if c = '"' then
        State := STATE_QUOTE
      else if not (CharInSet(c, quoted_string_chars)) then
        break;
    STATE_QCHAR:
      State := STATE_QTEXT;
    STATE_QUOTE:
      if c = '@' then
        State := STATE_EXPECTING_SUBDOMAIN
      else if c = '.' then
        State := STATE_LOCAL_PERIOD
      else
        break;
    STATE_LOCAL_PERIOD:
      if CharInSet(c, atom_chars) then
        State := STATE_ATOM
      else if c = '"' then
        State := STATE_QTEXT
      else
        break;
    STATE_EXPECTING_SUBDOMAIN:
      if CharInSet(c, letters_digits) then
        State := STATE_SUBDOMAIN
      else
        break;
    STATE_SUBDOMAIN:
      if c = '.' then begin
        inc(subdomains);
        State := STATE_EXPECTING_SUBDOMAIN
      end else if c = '-' then
        State := STATE_HYPHEN
      else if not (CharInSet(c, letters_digits)) then
        break;      STATE_HYPHEN:
      if CharInSet(c, letters_digits) then
        State := STATE_SUBDOMAIN
      else if c <> '-' then
        break;
    end;
    inc(i);
  end;
  if i <= n then
    Result := False
  else
    Result := (State = STATE_SUBDOMAIN) and (subdomains >= 2);
end;

{$ENDREGION}

{$REGION 'LOC comment'}

procedure TTGIConAPIsvc.DoLocComment(JSONObj : TJSONObject);
var
  lat, long : double;
  zoneStr : string;
  splitStr : TArray<string>;
begin
  //Do stuff for location comment and zone code
  if (JSONObj.FindValue('latitude') <> nil) and (JSONObj.FindValue('longitude') <> nil) then
  begin
    lat := strtofloat(JSONObj.GetValue('latitude').ToString.Trim(['"']));
    long := strtofloat(JSONObj.GetValue('longitude').ToString.Trim(['"']));
    zoneStr := ClosestZone(lat, long);
    splitStr := zoneStr.Split(['_']);

    //dm.qInsDATA.ParamByName('TGI_CLOSEST_ZONE').Value := splitStr[0];
    //dm.qInsDATA.ParamByName('TGI_TMWIN_LOC').Value := BuildLocComment(lat,long,splitStr[1],splitStr[2],splitStr[0]);

    if insDataFirstFld = false then
    begin
      insDataFlds := insDataFlds + ',';
      insDataVals := insDataVals + ',';
    end
    else
      insDataFirstFld := False;

    insDataFlds := insDataFlds + 'TGI_CLOSEST_ZONE';
    insDataVals := insDataVals + QuotedStr(splitStr[0]);

    insDataFlds := insDataFlds + ',TGI_TMWIN_LOC';
    insDataVals := insDataVals + ',' + QuotedStr(BuildLocComment(lat,long,splitStr[1],splitStr[2],splitStr[0]));
  end;
end;


function TTGIConAPIsvc.CompassPoint(LatFrom, LongFrom, LatTo, LongTo: Double): String;
const
  compassp: array[0..15] of string = ('N', 'NNE', 'NE', 'ENE', 'E',
    'ESE', 'SE', 'SSE', 'S', 'SSW', 'SW', 'WSW', 'W', 'WNW', 'NW', 'NNW');
var
  BearingDeg: Integer;
  x, y, LongDiff: Double;
begin
  LongDiff := LongTo - LongFrom;
  x := Sin(DegToRad(LongDiff)) * Cos(DegToRad(LatTo));
  y := Cos(DegToRad(LatFrom)) * Sin(DegToRad(LatTo))
    - Sin(DegToRad(LatFrom)) * Cos(DegToRad(LatTo)) * Cos(DegToRad(LongDiff));
  BearingDeg := Trunc(RadToDeg(ArcTan2(x, y)));
  if BearingDeg < 0 then
    BearingDeg := BearingDeg +360;
  Result := compassp[(((BearingDeg * 100) + 1125) mod 36000) div 2250];
  //Sample: Kansas City To St. Louis results in E
  //mTest.Lines.Add(CompassPoint(39.099912, -94.581213, 38.627089, -90.200203));
end;


function TTGIConAPIsvc.ClosestZone(lat,long : double) : string;
var
  newLat,newLong : string;
begin
  with dmR do
  begin
    qConvDegDms.ParamByName('val').Value := lat;
    qConvDegDms.ParamByName('kind').Value := 'LAT';

    qConvDegDms.Open;
    newLat := qConvDegDms.FieldByName('1').AsString;
    qConvDegDms.Close;

    qConvDegDms.ParamByName('val').Value := long;
    qConvDegDms.ParamByName('kind').Value := 'LONG';

    qConvDegDms.Open;
    newLong := qConvDegDms.FieldByName('1').AsString;
    qConvDegDms.Close;

    spFindZone.ParamByName('IPOSLAT').Value := newLat;
    spFindZone.ParamByName('IPOSLONG').Value := newLong;

    spFindZone.ExecProc;

    result := spFindZone.ParamByName('OZONE_ID').AsString;

    if result.Contains('XGST') then
      result := replaceStr(result,'XGST','');

    result := result + '_' + newLat + '_' + newLong;
  end;
end;


function TTGIConAPIsvc.BuildLocComment(lat,long : double; latStr,longStr,zone : string) : string;
var
  zoneLatStr,zoneLongStr,COMPPT,zoneDesc : string;
  zoneLat,zoneLong : double;
  zoneLatInt,zoneLongInt,latInt,longInt : integer;
begin
  with dmR do
  begin
    //Get coords for closest zone
    qGetZoneCoord.ParamByName('ZONE').Value := zone;
    qGetZoneCoord.Open;
    zoneLatStr := qGetZoneCoord.FieldByName('POSLAT').AsString;
    zoneLongStr := qGetZoneCoord.FieldByName('POSLONG').AsString;

    zoneDesc := qGetZoneCoord.FieldByName('SHORT_DESCRIPTION').AsString;
    qGetZoneCoord.Close;

    latInt := Round( (lat * 3600) );
    longInt := Round( (long * 3600) );

    //Convert to degrees to get compass direction
    qConvDmsDeg.ParamByName('val').Value := zoneLatStr;

    qConvDmsDeg.Open;
    zoneLat := qConvDmsDeg.FieldByName('1').AsFloat;
    zoneLatInt := Round( (zoneLat * 3600) );
    qConvDmsDeg.Close;

    qConvDmsDeg.ParamByName('val').Value := zoneLongStr;

    qConvDmsDeg.Open;
    zoneLong := qConvDmsDeg.FieldByName('1').AsFloat;
    zoneLongInt := Round( (zoneLong * 3600) );
    qConvDmsDeg.Close;

    //get dist between unit and closest zone
    qGetDist.ParamByName('STARTLAT').Value := latInt;
    qGetDist.ParamByName('STARTLONG').Value := longInt;
    qGetDist.ParamByName('ENDLAT').Value := zoneLatInt;
    qGetDist.ParamByName('ENDLONG').Value := zoneLongInt;

    qGetDist.Open;
    result := qGetDist.Fields[0].AsString + 'M ';
    qGetDist.Close;

    //Get compass direction
    COMPPT := CompassPoint(zoneLat,zoneLong,lat,long);
    result := result + COMPPT + ' of ' + zoneDesc;
  end;

end;

{$ENDREGION}


function TTGIConAPIsvc.DBLogon : string;
var
  config : TIniFile;
  iniPath,dbname : string;
  dbConnArr : TArray<string>;
begin
  result := '';
  if RootIniLoc = '' then
    result := 'The local.ini file in the app directory ('+extractfiledir(exeinfo.OriginalFileName)+') is not configured properly'
  else
  begin
    if fileexists(GenIniLoc) then
    begin
      try
        config := TIniFile.Create(GenIniLoc);
        dbName := config.ReadString('DB','database','');
        mailLog := config.ReadBool('Mail','logging',true);
        if mailLog then mailLogDtl := config.ReadBool('Mail','detailedLogging',false);
        config.Free;
      except on E:Exception do
        begin
          result := 'Unable to load the database from the service''s configuration file.'
        end;
      end;
    end
    else
    begin
      try
        config := TIniFile.Create(GenIniLoc);
        config.WriteString('DB','database','');
        config.WriteBool('Mail','logging',true);
        config.WriteBool('Mail','detailedLogging',false);
        config.Free;
      finally
        result := 'No configuration file exists. Attemped to create one at "'+GenIniLoc+'"';
      end;
    end;
    if result = '' then
    begin
      if dbname <> '' then
      begin
        dbConnArr := GetLoginInfo(dbName);
        if (length(dbConnArr) > 1) then
        begin
          if (dbConnArr[0] <> '') and (dbConnArr[1] <> '') then
          begin
            var err : string;
            err := dmr.Logon(dbname,dbConnArr[0],decpass(dbConnArr[1]),dbConnArr[2]);   //do database logon
            if err <> '' then result := err;
          end
          else result := 'The password and/or username is invalid in the database configuration specified in the service''s configuration file.';
        end
        else if (length(dbConnArr) = 1) then result := dbConnArr[0];
      end
      else
        result :=  'There is no database configuration specified in the the service''s configuration file.';
    end;
  end;
end;


end.
