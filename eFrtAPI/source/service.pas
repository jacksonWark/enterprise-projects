unit service;

interface

uses
  Winapi.Windows, Winapi.Messages, 
  System.SysUtils, System.Classes, System.IniFiles, System.StrUtils, System.AnsiStrings, System.DateUtils,
  Vcl.Graphics, Vcl.Controls, Vcl.SvcMgr, Vcl.Dialogs, Vcl.ExtCtrls,
  ExeInfo, Registry, ComObj,
  IdComponent, IdTCPConnection, IdTCPClient, IdExplicitTLSClientServerBase, IdMessageClient, IdSMTPBase, IdSMTP, IdBaseComponent, IdMessage, IdEMailAddress, IdHTTP, IdIOHandlerSocket, IdIOHandler, IdIOHandlerStack, IdSSL, IdSSLOpenSSL,
  FireDAC.Stan.Param,
  Xml.xmldom, Xml.XMLIntf, Xml.XMLDoc, xml.omnixmldom,
  Outlook, SharedServices, crypto;

type
  TeFrtAPIsvc = class(TService)
    Timer1: TTimer;
    ExeInfo: TExeInfo;
    XMLDoc: TXMLDocument;
    procedure Timer1Timer(Sender: TObject);
    procedure ServiceExecute(Sender: TService);
    procedure ServiceAfterInstall(Sender: TService);

  private
    { Private declarations }
    xml, url, logErrorName, uName: String;
    LogID, LogDtlID: Integer;
    IsDefaultEmail, testProfile, aTestMode, isOffline: Boolean;

    OutCtrl : TOutlookController;
    mailLog, mailLogDtl : Boolean;

    procedure autoRun;
    procedure runProfile;
    procedure ActiveProcessesReady;
    procedure LogErr(eID: Integer; aError: String);
    procedure LogDtl(sendData: String);

    procedure SetupGraphMail(isTest : boolean);
    procedure graphMail(aTo,aSubj,aBody : string; isTest : boolean);

    procedure InsLog;
    procedure qUpdLogDtl(aReply: String);
    procedure qUpdLogDtlACH(ActID: Integer; User1, User2, User3: String);
    procedure qUpdLogDtlODR(OdrID: Integer; OdrComment: String);
    procedure qUpdLogDtlTLO(TloID: Integer; User9, User10: String);
    function setupXMLurl(requestType: String) : Boolean;
    function checkSetup : Boolean;
    function GetSysEmail: String;
    function GetSysSMTP: String;
    function ValidEmail(email: string): boolean;
    function tryReqAS(shipid,shiptype,acccodes : string; testmode : boolean) : string;
    function tryReqRW(shipid,entrytype : string; inclwgt : boolean; qty : TArray<system.integer>;
      lens,wids,hgts : TArray<double>; rw : TArray<integer>; linearFeet : integer; testmode : boolean) : string;

    //Shared Services
    function DBLogon : string;
  public
    function GetServiceController: TServiceController; override;
    { Public declarations }
  end;

var
  eFrtAPIsvc: TeFrtAPIsvc;
  iniFile: TiniFile;

implementation

{$R *.dfm}

uses dmSvc;

procedure ServiceController(CtrlCode: DWord); stdcall;
begin
  eFrtAPIsvc.Controller(CtrlCode);
end;

procedure TeFrtAPIsvc.ActiveProcessesReady;                                    
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
            if isOffline = False then
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
                WriteLn(aFile, DateTimeToStr(Now)+' eFrtAPI log - Failed to autoRun');
                WriteLn(aFile, 'could not complete on server: '+ExeInfo.ComputerName);
                WriteLn(aFile, 'Error: '+ E.Message);
                CloseFile(aFile);
                GraphMail(GetSysEmail, 'eFrtAPI log - Failed to autoRun',
                  ' could not complete on server: '+ExeInfo.ComputerName
                  +' - Error Code : activeProcessReady'
                  +'Error: '+ E.Message,testProfile);
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
      WriteLn(aFile, DateTimeToStr(Now)+' eFrtAPI log - Failed to connect to DB2');
      WriteLn(aFile, 'could not complete on server: '+ExeInfo.ComputerName);
      WriteLn(aFile, 'Error: '+ E.Message);
      CloseFile(aFile);
    end;
  end;
  isOffline := False;
end;


procedure TeFrtAPIsvc.autoRun;                                                  
var
  aday,amo,ayr,ahr,amin,asec,aa: Word;
  dstr: String;
  aFile: TextFile;
  aFileExists: boolean;
begin
  try
    //check for a valid parameter
    logErrorName := ExtractFileDir(ParamStr(0))+'\eFrtAPIErrors.log';
    //test profile check
    testProfile := False;
    if dmR.qList.FieldByName('FRT_FUNCTION').Value = 'tryReqAS' then
      dmR.qToProcess := dmR.qToProcessAS
    else if dmR.qList.FieldByName('FRT_FUNCTION').Value = 'tryReqRW' then
      dmR.qToProcess := dmR.qToProcessRW
    else if dmR.qList.FieldByName('FRT_FUNCTION').Value = 'tryReqLONG' then
      dmR.qToProcess := dmR.qToProcessLONG
    else
    begin
      //error - no function properly defined
      LogErr(0, 'No function properly defined for profile: '
        + dmR.qList.FieldByName('NAME').Value);
      GraphMail(GetSysEmail,'eFrtAPI log - Failed to run',
          'No function properly defined for profile: '+dmR.qList.FieldByName('NAME').Value
          +' on server: '+ExeInfo.ComputerName,testProfile);
      abort;
    end;

    if ExeInfo.ComputerName = 'OIN' then
    begin
      testProfile := True;
    end
    else
    begin
      dmR.qToProcess.SQL[23] := '      between current date - 30 days and current timestamp -'
        +dmR.qList.FieldByName('PU_DELAY').AsString+' hours';
      dmR.qToProcess.SQL[26] := '    else T.CREATED_TIME between current date - 30 days and current timestamp - '
        +dmR.qList.FieldByName('DEL_DELAY').AsString+' hours end';
    end;
    with dmR.qProfDtl do
	  begin
      Open();
      if RecordCount=0 then
      begin
        LogErr(0,'Invalid or expired profile: '+dmR.qList.FieldByName('NAME').Value);
        GraphMail(GetSysEmail, 'eFrtAPI log - Failed to run',
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
          runProfile;
          with DmR.qUpdTime do
          begin
            ParamByName('NEXT_RUN').Value := IncMinute(Now, dmR.qProfDtl.FieldByName('MINS').Value);
            ExecSQL;
          end;
        except
          on E: Exception do
          begin
            if isOffline = False then
            begin
              //too many requests error email send
              LogErr(0,'Failed to complete: '+E.Message);
              GraphMail(
                GetSysEmail, 'eFrtAPI log - Failed to complete',
                'Profile: '+dmR.qList.FieldByName('NAME').Value
                +' could not complete on server:'+ExeInfo.ComputerName
                +' - Error Code : autoRun3'
                +' - Error: '+E.Message,testProfile
              );

              if ContainsText(E.Message,'429 Too Many Requests') then
              begin
                with DmR.qUpdTime do
                begin
                  ParamByName('NEXT_RUN').Value := IncMinute(Now, dmR.qProfDtl.FieldByName('MINS').Value);
                  ExecSQL;
                end;
              end;

            end;
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
      WriteLn(aFile, DateTimeToStr(Now)+' eFrtAPI - autoRun failed');
      WriteLn(aFile, 'Profile: '+dmR.qList.FieldByName('NAME').Value);
      WriteLn(aFile, 'could not complete on server: '+ExeInfo.ComputerName);
      WriteLn(aFile, 'Error: '+ E.Message);
      CloseFile(aFile);
      GraphMail(
        GetSysEmail, 'eFrtAPI - autoRun failed',
        'Profile: '+dmR.qList.FieldByName('NAME').Value
        +' could not complete on server: '+ExeInfo.ComputerName
        +' - Error Code : autoRun1'
        +' - Error: '+E.Message,testProfile
      );
    end;
  end;
end;

function TeFrtAPIsvc.checkSetup: Boolean;
begin
  result := True;
  if dmR.qProfDtl.Active = False then
    dmR.qProfDtl.Open;
  if dmR.qProfDtl.FieldByName('WEB_SERVICE_URL').Value = '' then
  begin
    result := False;
    LogErr(1,'Web Service URL is blank');
  end;
  if dmR.qProfDtl.FieldByName('REQUESTOR').Value = '' then
  begin
    result := False;
    LogErr(1,'Requestor is blank');
  end;
  if dmR.qProfDtl.FieldByName('AUTH').Value = '' then
  begin
    result := False;
    LogErr(1,'Authorization is blank');
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

function TeFrtAPIsvc.GetServiceController: TServiceController;
begin
  Result := ServiceController;
end;

function TeFrtAPIsvc.GetSysEmail: String;
begin
  IsDefaultEmail := False;
  try
    dmR.qOpen(dmR.qAdmin);
    if ((dmR.qAdmin.FieldByName('SMTP_USER').IsNull)
        or (ValidEmail(dmR.qAdmin.FieldByName('SMTP_USER').Value)=False)) then
    begin
      result := 'example@email.com';
      IsDefaultEmail := True;
    end
    else
      result := dmR.qAdmin.FieldByName('SMTP_USER').Value;
  except
    //find registry setting
    try
      try
        iniFile:=TiniFile.Create(ExtractFileDir(ParamStr(0))+'eFrtAPISvc.ini');
        result := IniFile.ReadString('Admin','SMTP_USER','example@email.com');
        IsDefaultEmail := True;
      finally
        inifile.Free;
      end;
    except
      result := 'example@email.com';
      IsDefaultEmail := True;
    end;
  end;
end;

function TeFrtAPIsvc.GetSysSMTP: String;
begin
  try
    dmR.qOpen(dmR.qAdmin);
    if dmR.qAdmin.FieldByName('SMTP_SERVER').IsNull then
      result := 'default' //removed 
    else
      result := dmR.qAdmin.FieldByName('SMTP_SERVER').Value;
  except
    //find registry setting
    try
      iniFile:=TiniFile.Create(ExtractFileDir(ParamStr(0))+'eFrtAPISvc.ini');
      result := IniFile.ReadString('Admin','SMTP_SERVER',''); //removed
    finally
      inifile.Free;
    end;
  end;
end;

procedure TeFrtAPIsvc.InsLog;
var
  dtFrom: TDateTime;
begin
  dtFrom := Now;
  with dmR.qInsLog do
    begin
      ParamByName('FRT_ID').Value := dmR.qProfDtl.FieldByName('FRT_ID').Value;
      ParamByName('TX_DATE').Value := dtFrom;
      ExecSQL;
      LogDtlID := 0;
    end;
    //find log id
    with dmR.qLogID do
    begin
      ParamByName('FRT_ID').Value := dmR.qProfDtl.FieldByName('FRT_ID').Value;
      ParamByName('TX_DATE').Value := dtFrom;
      Open();
      LogID := FieldByName('LOG_ID').Value;
      Close;
    end;
end;

procedure TeFrtAPIsvc.LogDtl(sendData: String);
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

procedure TeFrtAPIsvc.LogErr(eID: Integer; aError: String);
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

procedure TeFrtAPIsvc.qUpdLogDtl(aReply: String);
begin
  if dmR.qList.FieldByName('USE_LOG').Value = 'True' then
  begin
    with dmR.qUpdLogDtl do
    begin
      ParamByName('LOGDTL_ID').Value := LogDTLID;
      ParamByName('RECD_DATA').Value := aReply;
      ExecSQL;
    end;
  end;
end;

procedure TeFrtAPIsvc.qUpdLogDtlACH(ActID: Integer; User1, User2, User3: String);
begin
  if dmR.qList.FieldByName('USE_LOG').Value = 'True' then
  begin
    with dmR.qUpdLogDtlACH do
    begin
      ParamByName('LOGDTL_ID').Value := LogDTLID;
      ParamByName('ACH_ACT_ID').Value := ActID;
      ParamByName('ACH_USER1').Value := User1;
      ParamByName('ACH_USER2').Value := User2;
      ParamByName('ACH_USER3').Value := User3;
      ExecSQL;
    end;
  end;
end;

procedure TeFrtAPIsvc.qUpdLogDtlODR(OdrID: Integer; OdrComment: String);
begin
  if dmR.qList.FieldByName('USE_LOG').Value = 'True' then
  begin
    with dmR.qUpdLogDtlODR do
    begin
      ParamByName('LOGDTL_ID').Value := LogDTLID;
      ParamByName('ODR_DLID').Value := OdrID;
      ParamByName('ODR_COMMENT').Value := OdrComment;
      ExecSQL;
    end;
  end;
end;

procedure TeFrtAPIsvc.qUpdLogDtlTLO(TloID: Integer; User9, User10: String);
begin
  if dmR.qList.FieldByName('USE_LOG').Value = 'True' then
  begin
    with dmR.qUpdLogDtlTLO do
    begin
      ParamByName('LOGDTL_ID').Value := LogDTLID;
      ParamByName('TLO_DLID').Value := TloID;
      ParamByName('TLO_USER9').Value := User9;
      ParamByName('TLO_USER10').Value := User10;
      ExecSQL;
    end;
  end;
end;

procedure TeFrtAPIsvc.runProfile;
var
  aResult, errorCode, errorMsg, LAttrValue, u1, u2, u3, u9, u10,
    User10Note, FrtFunc, tmpStr: string;
  iResult, i, iODR: Integer;
  LDocument: IXMLDocument;
  LNodeElement, LNode: IXMLNode;
  l, w, h : TArray<double>;
  q, rw : TArray<integer>;
begin
  iODR := 0;
  with dmR.qToProcess do
  begin
    Open;
    while not eof do
    begin
      FrtFunc := dmR.qList.FieldByName('FRT_FUNCTION').Value;
      if FrtFunc = 'tryReqAS' then
      begin
        User10Note := 'API ACC ERROR';
        aResult := tryReqAS(
                    FieldByName('TRACE_B10').AsString,
                    FieldByName('P_OR_D').AsString, FieldByName('EDI_CODE').AsString, aTestMode
                    );
      end
      else if (FrtFunc = 'tryReqRW') or (FrtFunc = 'tryReqLONG') then
      begin
        if dmR.qRWodr.Active then
          dmR.qRWodr.Close;
        if FrtFunc = 'tryReqRW' then
          User10Note := 'API DIM ERROR';
        else
          User10Note := 'API LONG ERROR';

        //find the most recent entry
        dmR.qRWodr.ParamByName('ORDER_ID').Value := FieldByName('DETAIL_LINE_ID').Value;
        dmR.qRWodr.Open;
        iODR := dmR.qRWodr.FieldByName('ID').Value;
        //update USER1 not reqd for all but the most recent
        while not(dmR.qRWodr.eof) do
        begin
          if dmR.qRWodr.FieldByName('ID').Value <> iODR then
          begin
            dmR.qUpdODRrw.ParamByName('ID').Value := dmR.qRWodr.FieldByName('ID').Value;
            dmR.qUpdODRrw.ParamByName('USER1').Value := 'API-Not Reqd-' + DateTimeToStr(Now);
            dmR.qUpdODRrw.ParamByName('LOC_COMMENT').Value := '';
            if dmR.qProfDtl.FieldByName('TM_MODE').Value = 'WRITE' then
              dmR.qUpdODRrw.ExecSQL;
          end;
          dmR.qRWodr.Next;
        end;
        dmR.qRWodr.First;
        if (FieldByName('EDICUBE_WGT').AsFloat * 1.03) >= dmR.qRWodr.FieldByName('USER2').AsFloat then
        begin
          //update USER1 for most recent if cube lower than original
          dmR.qUpdODRrw.ParamByName('ID').Value := iODR;
          dmR.qUpdODRrw.ParamByName('USER1').Value := 'API-Not Reqd-' + DateTimeToStr(Now);
          dmR.qUpdODRrw.ParamByName('LOC_COMMENT').Value := '';
          if dmR.qProfDtl.FieldByName('TM_MODE').Value = 'WRITE' then
            dmR.qUpdODRrw.ExecSQL;
        end
        else
        begin
          //prepare to send
          //get cube and set to arrays
          dmR.qRWdtl.ParamByName('DLID').Value := FieldByName('DETAIL_LINE_ID').Value;
          dmR.qRWdtl.Open;
          SetLength(q, dmR.qRWdtl.RecordCount);
          SetLength(l, dmR.qRWdtl.RecordCount);
          SetLength(w, dmR.qRWdtl.RecordCount);
          SetLength(h, dmR.qRWdtl.RecordCount);
          SetLength(rw, dmR.qRWdtl.RecordCount);
          i := 0;
          while not (dmR.qRWdtl.eof) do
          begin
            q[i] := dmR.qRWdtl.FieldByName('DIM_PIECES').Value;
            l[i] := dmR.qRWdtl.FieldByName('DIM_LENGTH').Value;
            w[i] := dmR.qRWdtl.FieldByName('DIM_WIDTH').Value;
            h[i] := dmR.qRWdtl.FieldByName('DIM_HEIGHT').Value;
            rw[i] := Round(
                dmR.qRWdtl.FieldByName('DIM_PIECES').Value
                * dmR.qRWdtl.FieldByName('DIM_LENGTH').Value
                * dmR.qRWdtl.FieldByName('DIM_WIDTH').Value
                * dmR.qRWdtl.FieldByName('DIM_HEIGHT').Value * 0.000578704
              );
            inc(i);
            dmR.qRWdtl.Next;
          end;
          dmR.qRWdtl.Close;
        end;
        //send
        if FrtFunc = 'tryReqLONG' then
          aResult := tryReqRW(FieldByName('TRACE_B10').AsString, 'L',
            False, q,l,w,h,rw,dmR.qRWodr.FieldByName('USER2').AsInteger, aTestMode)
        else
          aResult := tryReqRW(FieldByName('TRACE_B10').AsString, 'D',
            False, q,l,w,h,rw,13, aTestMode);
        dmR.qRWodr.Close;
      end
      else
      begin
        User10Note := 'API OTH ERROR';
        //lengths etc
        aResult := '';
      end;

      //use reply
      if ContainsText(aResult, 'ACCEPTED') then
      begin
        if dmR.qProfDtl.FieldByName('TM_MODE').Value = 'WRITE' then
        begin
          //achtlo
          u1 := 'API-Sent-' + DateTimeToStr(Now);
          u2 := '';
          u3 := '';
          if FrtFunc = 'tryReqAS' then
            dmR.spODRstat.ParamByName('ISTAT_COMMENT').Value := u1
          else if (FrtFunc = 'tryReqRW') or (FrtFunc = 'tryReqLONG') then
            dmR.qUpdODRrw.ParamByName('LOC_COMMENT').Value := '';
        end;
      end
      else if ContainsText(aResult, 'Our web site is currently off-line') then
      begin
        //send email to admin
        GraphMail(
          GetSysEmail, 
          'eFrtAPI notice - client web service is offline',
          'Web service: '+dmR.qList.FieldByName('WEB_SERVICE_URL').Value
          +' is offline according to the response file. All processing has been paused. '
          + 'eFrtAPI will try again in 2 hours.', testProfile
        );

        //set next run time to 2 hours for ALL profiles set to run in the next 2 hours
        DmR.qUpdTimeAll.ExecSQL;

        //bail on all remaining orders
        isOffline := True;
        abort;
      end
      else
      begin
        if AnsiLeftStr(aResult, 7)= 'ERROR: ' then
        begin
          tmpStr := AnsiRightStr(aResult, Length(aResult)-7);
          i := AnsiPos(' ', tmpStr);
          errorCode := AnsiLeftStr(tmpStr, i - 1);
          errorMsg := AnsiRightStr(tmpStr, Length(tmpStr) - i);
        end
        else
        begin
          LDocument := TXMLDocument.Create(nil);
          LDocument.LoadFromXML(aResult);
          LNodeElement := LDocument.ChildNodes.FindNode('asresult');
          if (LNodeElement <> nil) then
          begin
            for I := 0 to LNodeElement.ChildNodes.Count - 1 do
            begin
              LNode := LNodeElement.ChildNodes.Get(I);
              if LNode.NodeName = 'errorcode' then
                errorcode := LNode.Text
              else if LNode.NodeName = 'errormsg' then
                errormsg := AnsiLeftStr(LNode.Text,60);
            end;
          end;
          LNodeElement := LDocument.ChildNodes.FindNode('dimsrwrslt');
          if (LNodeElement <> nil) then
          begin
            for I := 0 to LNodeElement.ChildNodes.Count - 1 do
            begin
              LNode := LNodeElement.ChildNodes.Get(I);
              if LNode.NodeName = 'errorcode' then
                errorcode := LNode.Text
              else if LNode.NodeName = 'errormsg' then
                errormsg := AnsiLeftStr(LNode.Text,60);
            end;
          end;
        end;

        if dmR.qProfDtl.FieldByName('TM_MODE').Value = 'WRITE' then
        begin
          dmR.spODRstat.ParamByName('ISTAT_COMMENT').Value
            :=  'API-' + errorCode + '-' + errorMsg;
          //achtlo
          u1 := 'API-errorcode-' + errorcode;
          u2 := AnsiLeftStr(errormsg,40);
          if length(errormsg) > 40 then
            u3 := AnsiRightStr(errormsg, length(errormsg)-40)
          else
            u3 := '';

          //if rejected update user10 (unless LONGFRTCHG or TDGFEE) modified 2021-09-22 v
          if (((FrtFunc = 'tryReqAS') and (FieldByName('ACODE_ID').AsString<>'TDGFEE')
            and (FieldByName('ACODE_ID').AsString<>'LONGFRTCHG'))
            or (FrtFunc = 'tryReqRW') or (FrtFunc = 'tryReqLONG')) then
          begin
            if FieldByName('USER9').Value = '' then                             
            begin                                                                 
              if dmR.qAdmin.Active = False then
                dmR.qAdmin.Open;
              u9 := dmR.qAdmin.FieldByName('USER9_CODE').Value;
            end
            else
              u9 := FieldByName('USER9').Value;
            if FieldByName('USER10').Value = '' then
              u10 := User10Note
            else if ContainsText(FieldByName('USER10').Value,User10Note) = True  then
              u10 := FieldByName('USER10').Value
            else
              u10 := AnsiLeftStr(FieldByName('USER10').Value + User10Note,40);

            dmR.qUpdTLO.ParamByName('DETAIL_LINE_ID').Value := FieldByName('DETAIL_LINE_ID').Value;
            dmR.qUpdTLO.ParamByName('USER9').Value := u9;
            dmR.qUpdTLO.ParamByName('USER10').Value := u10;
            dmR.qUpdTLO.ExecSQL;
          end;

          //Log
          qUpdLogDtlTLO(FieldByName('DETAIL_LINE_ID').Value, U9, U10);
          if (FrtFunc = 'tryReqRW') or (FrtFunc = 'tryReqLONG') then
            dmR.qUpdODRrw.ParamByName('LOC_COMMENT').Value := AnsiLeftStr(errorMsg,80);
        end;
      end;

      if dmR.qProfDtl.FieldByName('TM_MODE').Value = 'WRITE' then
      begin
        if FrtFunc = 'tryReqAS' then
        begin
          //write to odrstat
          dmR.spODRstat.ParamByName('ISTAT_COMMENT').Value := u1;
          dmR.spODRstat.ParamByName('IORDER_ID').Value :=  dmR.qToProcess.FieldByName('DETAIL_LINE_ID').Value;
          dmR.spODRstat.ParamByName('ICHANGED').Value :=  Now;
          dmR.spODRstat.ParamByName('ISTATUS_CODE').Value :=  'ACCNOTIFY';
          dmR.spODRstat.ParamByName('IUPDATED_BY').Value :=  uName;    
          dmR.spODRstat.ParamByName('IINS_DATE').Value :=  Now;
          dmR.spODRstat.ParamByName('IUSER1').Value :=  dmR.qToProcess.FieldByName('ACODE_ID').Value;
          dmR.spODRstat.Execute;
          //Log
          qUpdLogDtlODR(dmR.spODRstat.ParamByName('IORDER_ID').Value,
            dmR.spODRstat.ParamByName('ISTAT_COMMENT').Value);

          //write to acharge_tlorder
          dmR.qUpdACHTLO.ParamByName('ACT_ID').Value := dmR.qToProcess.FieldByName('ACT_ID').Value;
          dmR.qUpdACHTLO.ParamByName('USER1').Value := u1;
          dmR.qUpdACHTLO.ParamByName('USER2').Value := u2;
          dmR.qUpdACHTLO.ParamByName('USER3').Value := u3;
          dmR.qUpdACHTLO.ExecSQL;
          //Log
          qUpdLogDtlACH(dmR.qToProcess.FieldByName('ACT_ID').Value, u1, u2, u3);
        end
        else if (FrtFunc = 'tryReqRW') or (FrtFunc = 'tryReqLONG') then
        begin
          dmR.qUpdODRrw.ParamByName('ID').Value := iODR;
          dmR.qUpdODRrw.ParamByName('USER1').Value := u1;
          dmR.qUpdODRrw.ExecSQL;
        end;
      end;
      Next;
    end;

    //if RW then clean up empty user1's
    if dmR.qList.FieldByName('FRT_FUNCTION').Value = 'tryReqRW' then
    begin
      dmR.qUpdNotReqd.Open;
      while not(dmR.qUpdNotReqd.eof) do
      begin
        dmR.qUpdODRrw.ParamByName('ID').Value := dmR.qUpdNotReqd.FieldByName('ID').Value;
        dmR.qUpdODRrw.ParamByName('USER1').Value := 'API-Not Reqd-' + DateTimeToStr(Now);
        dmR.qUpdODRrw.ParamByName('LOC_COMMENT').Value := '';
        if dmR.qProfDtl.FieldByName('TM_MODE').Value = 'WRITE' then
          dmR.qUpdODRrw.ExecSQL;
        dmR.qUpdNotReqd.Next;
      end;
      dmR.qUpdNotReqd.Close;
    end;

    Close;
  end;


end;

procedure TeFrtAPIsvc.ServiceAfterInstall(Sender: TService);
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create(KEY_READ or KEY_WRITE);
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('\SYSTEM\CurrentControlSet\Services\' + Name, false) then
    begin
      Reg.WriteString('Description', 'eFrtAPI Service');
      Reg.CloseKey;
    end;
  finally
    Reg.Free;
  end;
end;

procedure TeFrtAPIsvc.ServiceExecute(Sender: TService);
const
  SecBetweenRuns = 60; 
var
  Count: Integer;
  Reg: TRegistry;
begin
  Count := 0;
  DefaultDOMVendor := sOmniXmlVendor;
  isOffline := False;
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
  try
    GraphMail(
      'example@email.com','eFrtAPI log - Failed to start service',
      'Failed to create database connection on server:'+ExeInfo.ComputerName+'. Error: '+ret,
      true
    );
    var config := TIniFile.Create(GenIniLoc);
    config.WriteString('DB','database','');
    config.Free;
  finally
  end;
end;

function TeFrtAPIsvc.setupXMLurl(requestType: String): Boolean;
begin
  if checkSetup = True then
  begin
    url := dmR.qProfDtl.FieldByName('WEB_SERVICE_URL').Value;

    xml := 
      '<?xml version="1.0" encoding="ISO-8859-1"?>'
       + requestType
       + '<requestor>' + dmR.qProfDtl.FieldByName('REQUESTOR').Value + '</requestor>'
       + '<authorization>' + dmR.qProfDtl.FieldByName('AUTH').Value + '</authorization>'
       + '<login>' + dmR.qProfDtl.FieldByName('USER_ID').Value + '</login>'
       + '<passwd>' + dmR.qProfDtl.FieldByName('PSWD').Value + '</passwd>';

  end;
  result := checkSetup;
end;

procedure TeFrtAPIsvc.Timer1Timer(Sender: TObject);
var
  aFile : TextFile;
  aFileExists: boolean;
begin
  Timer1.Enabled := False;

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
                WriteLn(aFile, DateTimeToStr(Now)+' eFrtAPI log - Failed to autoRun');
                WriteLn(aFile, 'could not complete on server: '+ExeInfo.ComputerName);
                WriteLn(aFile, 'Error: '+ E.Message);
                CloseFile(aFile);
                GraphMail(GetSysEmail, 'eFrtAPI log - Failed to autoRun',
                  ' could not complete on server: '+ExeInfo.ComputerName
                  +' - Error Code : timer2'
                  +'Error: '+ E.Message,testProfile);
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
      WriteLn(aFile, DateTimeToStr(Now)+' eFrtAPI log - Failed to connect to DB2');
      WriteLn(aFile, 'could not complete on server: '+ExeInfo.ComputerName);
      WriteLn(aFile, 'Error: '+ E.Message);
      CloseFile(aFile);
      GraphMail(GetSysEmail, 'eFrtAPI log - Failed to connect to DB2',
        ' could not complete on server: '+ExeInfo.ComputerName
        +' - Error Code : timer1'
        +'Error: '+ E.Message,testProfile);
    end;
  end;
end;

function TeFrtAPIsvc.tryReqAS(shipid, shiptype, acccodes: string;
  testmode: boolean): string;
var
  strStream : TStringStream;

  I : integer;
  splitStr : Tarray<string>;
  HTTP : TidHTTP;
  SSL : TidSSLIOHandlerSocketOpenSSL;
begin
  if setupXMLurl('<asrequest>') then
  begin
    if testmode then xml := xml + '<testmode>Y</testmode>';

    if length(shipid) > 10 then //exit and show error code that ship id is too long
      exit('ERROR: shipid exceeds maximum length of 10');
    else if length(shipid) < 10 then //exit and show error code that ship id is not there
      exit('ERROR: shipid must 10 digits');
    else
      xml := xml + '<shipid>' + shipid + '</shipid>';

    if (shiptype = 'P') or (shiptype = 'D') then
      xml := xml + '<shiptype>' + shiptype + '</shiptype>'
    else if (shiptype = 'p') or (shiptype = 'd') then
      xml := xml + '<shiptype>' + UpperCase(shiptype) + '</shiptype>'
    else //exit and show error code saying that ship type is invalid
      exit('ERROR: shiptype invalid - must be P or D');

    if length(acccodes) > 0 then
    begin
      if containstext(acccodes,',') then
      begin
        splitStr := acccodes.Split([',']);

        if length(splitStr) > 10 then
          exit('ERROR: acccodes exceeds maximum of 10 codes');

        for I := 0 to (length(splitStr) - 1) do
          xml := xml + '<detail><acccode>' + splitStr[I] + '</acccode></detail>';
      end
      else
        xml := xml + '<detail><acccode>' + acccodes + '</acccode></detail>';
    end
    else //exit and show error code saying there are no acc codes
      result := exit('ERROR: no acccodes');

    xml := xml + '</asrequest>';
    LogDtl(xml);

    strStream := TStringStream.Create(xml);

    //Create HTTP client
    HTTP := TidHTTP.Create;
    SSL := TidSSLIOHandlerSocketOpenSSL.Create;
    SSL.SSLOptions.Method := TidSSLVersion(1);
    HTTP.IOHandler := SSL;

    try
      result := HTTP.Post(url, strStream);
      qUpdLogDtl(result);
    finally
      HTTP.Free;
      SSL.Free;
    end;

    strStream.Free;

    //if result is nothing return error code
    if result = '' then result := 'ERROR: HTTP-POST returned nothing';
  end;
end;

function TeFrtAPIsvc.tryReqRW(shipid, entrytype: string; inclwgt: boolean;
  qty: TArray<system.integer>; lens, wids, hgts: TArray<double>;
  rw: TArray<integer>; linearFeet: integer; testmode: boolean): string;
var
  strStream : TStringStream;
  I : integer;
  HTTP : TidHTTP;
  SSL : TidSSLIOHandlerSocketOpenSSL;
begin
  if setupXMLurl('<dimsrwrequest>') then
  begin
    url := 'https://www.tst-cfexpress.com/xml/dims-rw';

    if testmode then xml := xml + '<testmode>Y</testmode>';

    if length(shipid) > 10 then //exit and show error code that ship id is too long
      exit('ERROR: shipid exceeds maximum length of 10');
    else if length(shipid) < 10 then //exit and show error code that ship id is not there
      exit('ERROR: shipid must 10 digits');
    else
      xml := xml + '<shipid>' + shipid + '</shipid>';

    if (entrytype = 'D') or (entrytype = 'R') or (entrytype = 'L') then
      xml := xml + '<entrytype>' + entrytype + '</entrytype>'
    else if (entrytype = 'd') or (entrytype = 'r') or (entrytype = 'l') then
      xml := xml + '<entrytype>' + UpperCase(entrytype) + '</entrytype>'
    else //exit and show error code saying that ship type is invalid
      exit('ERROR: entrytype invalid - must be D, R, or L');

    if inclwgt then xml := xml + '<inclwgt>Y</inclwgt>';

    if (entryType = 'D') or (entrytype = 'd') then
    begin
      for I := 0 to (length(lens) - 1) do
      begin
        xml := xml + '<detail>';

        if qty[I] = 0 then
          exit('ERROR: qty must be non-zero');

        xml := xml + '<qty>' + inttostr(qty[I]) + '</qty>';

        if (lens[I] = 0) or (wids[I] = 0) or (hgts[I] = 0) then
          exit('ERROR: dimensions must be non-zero');

        if lens[I] > 636 then
          exit('ERROR: len cannot exceed 636 inches');

        if wids[I] > 96 then
          exit('ERROR: wid cannot exceed 96 inches');

        if hgts[I] > 96 then
          exit('ERROR: hgt cannot exceed 96 inches');

        xml := xml + '<len>' + floattostr(lens[I]) + '</len>'
                   + '<wid>' + floattostr(wids[I]) + '</wid>'
                   + '<hgt>' + floattostr(hgts[I]) + '</hgt>';

        if inclwgt then
        begin
          if rw[I] = 0 then
            exit('ERROR: rw must be non-zero');

          xml := xml + '<rw>' + inttostr(rw[I]) + '</rw>';
        end;

        xml := xml + '</detail>';

      end;

    end
    else if (entryType = 'R') or (entrytype = 'r') then
    begin
      for I := 0 to (length(rw) - 1) do
        begin
          if qty[I] = 0 then
            exit('ERROR: qty must be non-zero');

          if rw[I] = 0 then
            exit('ERROR: rw must be non-zero');

          xml := xml + '<detail><qty>' + inttostr(qty[I]) + '</qty><rw>' + inttostr(rw[I]) + '</rw></detail>';
        end;
      xml := xml + '</detail>';
    end
    else if (entryType = 'L') or (entrytype = 'l') then
    begin
      if (linearFeet > 9) and (linearFeet < 54) then
        xml := xml + '<detail><linearFeet>' + inttostr(linearFeet) + '</linearFeet></detail>'
      else //return error
        exit('ERROR: linearFeet must be between 10 and 53 feet');
    end;

    xml := xml + '</dimsrwrequest>';

    LogDtl(xml);

    strStream := TStringStream.Create(xml);

    //Create HTTP client
    HTTP := TidHTTP.Create;
    SSL := TidSSLIOHandlerSocketOpenSSL.Create;
    SSL.SSLOptions.Method := TidSSLVersion(1);
    HTTP.IOHandler := SSL;

    try
      result := HTTP.Post(url, strStream);
      qUpdLogDtl(result);
    finally
      HTTP.Free;
      SSL.Free;
    end;

    strStream.Free;

    //if result is nothing return error code
    if result = '' then result := 'ERROR: HTTP-POST returned nothing';
  end;
end;

function TeFrtAPIsvc.ValidEmail(email: string): boolean;
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


procedure TeFrtAPIsvc.SetupGraphMail(isTest : boolean);
begin
  if isTest then
    OutCtrl := TOutlookCOntroller.Create(
      true,
      true,
      exeInfo.UserName,
      false
    )
  else
    OutCtrl := TOutlookCOntroller.Create(
      mailLog,
      mailLogDtl,
      exeInfo.UserName,
      false
    );

  OutCtrl.Auth(false);
end;


procedure TeFrtAPIsvc.graphMail(aTo,aSubj,aBody : string; isTest : boolean);
var
  recipients : TArray<string>;
begin
  if OutCtrl = nil then
    SetupGraphMail(isTest);

  SetLength(recipients,2);
  Recipients[0] := aTo;
  Recipients[1] := 'example@email.com'; //removed

  if isTest = false then
  begin
    SetLength(Recipients,length(Recipients) + 1);
    Recipients[length(Recipients) - 1] := 'example@email.com'; //removed
  end;

  OutCtrl.SendMail(
    'html',
    'example@email.com',
    'example@email.com', // removed
    'example@email.comm',
    aSubj, 
    aBody, 
    Recipients,
    {CC}nil,
    {BCC}nil,
    {ATTACH}nil
  );
end;


function TeFrtAPIsvc.DBLogon : string;
var
  config : TIniFile;
  iniPath,dbname : string;
  dbConnArr : TArray<string>;
begin
  result := '';
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
		      uName := uppercase(dbConnArr[0]);
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

end.
