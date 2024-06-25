unit dmod;

interface

uses
  System.SysUtils, System.Classes, System.StrUtils, System.Character, System.Math,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param, FireDAC.Stan.Error, FireDAC.Stan.Async, FireDAC.Stan.Def, FireDAC.Stan.Pool,
  FireDAC.Phys, FireDAC.Phys.Intf, FireDAC.Phys.ODBC, FireDAC.Phys.DB2Def, FireDAC.Phys.ODBCDef, FireDAC.Phys.DB2,
  FireDAC.DatS,  FireDAC.DApt.Intf,  FireDAC.DApt, FireDAC.UI.Intf, FireDAC.VCLUI.Wait,
  FireDAC.Comp.Client, FireDAC.Comp.DataSet,
  Data.DB,
  FireDAC.Comp.BatchMove, FireDAC.Comp.BatchMove.Text, FireDAC.Comp.BatchMove.Dataset,
  utils, FireDAC.VCLUI.Async, FireDAC.Comp.UI;

type
  TDataMod = class(TDataModule)
    qClientID: TFDQuery;
    DB: TFDConnection;
    qUpdClientDelta: TFDQuery;
    qUpdOpARDelta: TFDQuery;
    qUpdCloARDelta: TFDQuery;
    bMove: TFDBatchMove;
    custData: TFDMemTable;
    qGetCustID: TFDQuery;
    qInsCust: TFDQuery;
    qUpdCust: TFDQuery;
    qFindZone: TFDQuery;
    qFindProvZone: TFDQuery;
    qSelCli: TFDQuery;
    spUpdCont: TFDStoredProc;
    spCreateClient: TFDStoredProc;
    spUpdBilling: TFDStoredProc;
    spUpdCredit: TFDStoredProc;
    qUserID: TFDQuery;
    qUserName: TFDQuery;
    qClientUpd: TFDQuery;
    qClientDeltaActive: TFDQuery;
    qOpARDeltaActive: TFDQuery;
    qCloARDeltaActive: TFDQuery;
    qHRCID: TFDQuery;
    qValidProvShort: TFDQuery;
    qValidProv: TFDQuery;
    spCalcClient: TFDStoredProc;
    qDeltaGetARIDs: TFDQuery;
    qDeltaDelDupAR: TFDQuery;
    qClearOldClient: TFDQuery;
    qClearOldAR: TFDQuery;
    qCredLim: TFDQuery;
    qCredHist: TFDQuery;
    qCustData: TFDQuery;
    qCredLimHist: TFDQuery;
    qClientIDCLIENT_ID: TStringField;
    qClientIDNAME: TStringField;
    qClientIDCITY: TStringField;
    qClientIDINS_TIMESTAMP: TSQLTimeStampField;
    qChgID_Cli: TFDQuery;
    qChgID_CliAR: TFDQuery;
    qChgID_CliBal: TFDQuery;
    qChgID_CliStatus: TFDQuery;
    qChgID_CredHist: TFDQuery;
    qChgID_CustData: TFDQuery;
    qChgID_TLcust: TFDQuery;
    qChgID_TLdest: TFDQuery;
    qChgID_TLcont: TFDQuery;
    qChgID_TLbill: TFDQuery;
    qChgID_TLcare: TFDQuery;
    qChgID_TLpickup: TFDQuery;
    qChgID_TLorig: TFDQuery;
    qChgImpact: TFDQuery;
    qGetPhCID: TFDQuery;
    qGetNLID: TFDQuery;
    qInsNL: TFDQuery;
    qInsNLDTL: TFDQuery;
    qChgID_Phone: TFDQuery;
    qChgID_Delta: TFDQuery;
    qTaxExRpt: TFDQuery;
    qNameSearch: TFDQuery;
    StringField1: TStringField;
    qClientUpdTest: TFDQuery;
  private
    { Private declarations }

    clientName, clientCity : string;

    //change these to INI values?
    
  public
    { Public declarations }

    function initDB(dbase,user,pwd,schema : string) : string;

    function UpdCustDelta(I : integer = 0) : string;
    function UpdOARDelta(I : integer = 0) : string;
    function UpdCARDelta(I : integer = 0) : string;
    function CustDeltaActive(I : integer = 0) : string;
    function OARDeltaActive(I : integer = 0) : string;
    function CARDeltaActive(I : integer = 0) : string;
    function CalcAging(I : integer = 0) : string;
    function CleanDeltaTable : string;
    function ClearOldDelta : string;

    function DoClientIns(ID : string) : boolean;
    function DoClientUpd : boolean;
    function DoSameCredLim(CID : string; credLim : integer) : boolean;

    function GetFieldIDs : TArray<integer>;

    procedure CSVtoDS(csvPath,sep,delim : string; insert : boolean);
    procedure CreateMapping(map : TFDBatchMoveMappings);
    procedure UpdateMapping(map : TFDBatchMoveMappings);
    function GetRecordCSV(delim,sep,kind,ID : string) : string;

    function GetZoneCode(ClientID : string; ship : boolean = false) : string;
    function IncNumber(ClientID : string) : string;

    function ExistsClientID(ClientID : string) : boolean;
    function NameSearch(name,client_id : string) : string;
    function GetClientID(hrcID : string) : string;
    function GetClientName : string;
    function GetClientCity : string;
    function ChangeClientID(oldClientID,newClientID,name : string) : boolean;
    function ImpactAssess(ClientID : string) : string;

    function ValidateSalesRep(salesRep : string) : string;
    function ValidateProvince(province : string) : string;
    function PrettyPhone(phone : string) : string;
    function ValidateStatus_1(value,default : string) : TArray<string>;
    function TestFunc : TArray<string>;
    function RevBoolStr(TF : boolean) : string;
    function HandleBoolStr(boolStr : string) : boolean;

  end;

var
  DataMod: TDataMod;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

function TDataMod.initDB(dbase,user,pwd,schema : string) : string;
begin
  try
    DB.Params.Values['Alias'] := dbase;
    DB.Params.Values['ODBCAdvanced'] := 'CurrentSchema='+schema;
    DB.Params.username := user;
    DB.Params.password := pwd;
    DB.Connected := true;
    Db.ExecSQL('SET CURRENT SCHEMA '+schema);
    Db.ExecSQL('SET CURRENT PATH "SYSFUN","SYSPROC","SYSIBMADM",' + uppercase(schema));
    result := 'success';
  except
    on E : Exception do
    begin
      result := 'Failed to connect to database: ' + E.Message;
    end;
  end;

end;

function TDataMod.TestFunc : TArray<string>;
begin

  qValidProv.Prepare;
  qValidProv.Params[0].Value := 'BCWIL';

  setLength(result,4);
  result[0] := qValidProv.Command.CommandIntf.CommandText;
  result[1] := qValidProv.Command.SchemaName;
  result[2] := qValidProv.Command.SQLText;
  result[3] := qValidProv.Text;
  datamod.qValidProv.open;
end;

{$REGION 'Delta Updates'}

function TDataMod.UpdCustDelta(I : integer = 0) : string;
begin
  result := '';
  try
    qUpdClientDelta.ExecSQL;
  except
    on E: Exception do
    begin
      if (ContainsText(E.Message,'SQL0964C  The transaction log for the database is full.  SQLSTATE=57011')) and (I < numRetries) then
      begin
        Sleep(waitLength);
        inc(I);
        logMsg('[WARNING] - Failed to run qUpdCustDelta, transaction log full. Retrying... ');
        result := UpdCustDelta(I);
      end
      else
        result := 'Failed to run qUpdCustDelta: ' + E.Message;
    end;
  end;
  if (I > 0) and (result = '') then
  begin
    LogMsg('[WARNING] - ...Success');
    result := 'success';
  end
  else if (I = 0) and (result = 'success') then
    result := '';
end;


function TDataMod.UpdOARDelta(I : integer = 0) : string;
begin
  result := '';
  try
    qUpdOpARDelta.ExecSQL;
  except
    on E: Exception do
    begin
      if (ContainsText(E.Message,'SQL0964C  The transaction log for the database is full.  SQLSTATE=57011')) and (I < numRetries) then
      begin
        Sleep(waitLength);
        inc(I);
        logMsg('[WARNING] - Failed to run qUpdOpARDelta, transaction log full. Retrying... ');
        result := UpdOARDelta(I);
      end
      else
        result := 'Failed to run qUpdOpARDelta: ' + E.Message;
    end;
  end;
  if (I > 0) and (result = '') then
  begin
    LogMsg('[WARNING] - ...Success');
    result := 'success';
  end
  else if (I = 0) and (result = 'success') then
    result := '';
end;


function TDataMod.UpdCARDelta(I : integer = 0) : string;
begin
  result := '';
  try
    qUpdCloARDelta.ExecSQL;
  except
    on E: Exception do
    begin
      if (ContainsText(E.Message,'SQL0964C  The transaction log for the database is full.  SQLSTATE=57011')) and (I < numRetries) then
      begin
        Sleep(waitLength);
        inc(I);
        logMsg('[WARNING] - Failed to run qUpdCloARDelta, transaction log full. Retrying... ');
        result := UpdCARDelta(I);
      end
      else
        result := 'Failed to run qUpdCloARDelta: ' + E.Message;
    end;
  end;
  if (I > 0) and (result = '') then
  begin
    LogMsg('[WARNING] - ...Success');
    result := 'success';
  end
  else if (I = 0) and (result = 'success') then
    result := '';
end;


function TDataMod.CustDeltaActive(I : integer = 0) : string;
begin
  result := '';
  try
    qClientDeltaActive.ExecSQL;
  except
    on E: Exception do
    begin
      if (ContainsText(E.Message,'SQL0964C  The transaction log for the database is full.  SQLSTATE=57011')) and (I < numRetries) then
      begin
        Sleep(waitLength);
        inc(I);
        logMsg('[WARNING] - Failed to run qClientDeltaActive, transaction log full. Retrying... ');
        result := CustDeltaActive(I);
      end
      else
        result := 'Failed to run qClientDeltaActive: ' + E.Message;
    end;
  end;
  if (I > 0) and (result = '') then
  begin
    LogMsg('[WARNING] - ...Success');
    result := 'success';
  end
  else if (I = 0) and (result = 'success') then
    result := '';
end;


function TDataMod.OARDeltaActive(I : integer = 0) : string;
begin
  result := '';
  try
    qOpARDeltaActive.ExecSQL;
  except
    on E: Exception do
    begin
      if (ContainsText(E.Message,'SQL0964C  The transaction log for the database is full.  SQLSTATE=57011')) and (I < numRetries) then
      begin
        Sleep(waitLength);
        inc(I);
        logMsg('[WARNING] - Failed to run qOpARDeltaActive, transaction log full. Retrying... ');
        result := OARDeltaActive(I);
      end
      else
        result := 'Failed to run qOpARDeltaActive: ' + E.Message;
    end;
  end;
  if (I > 0) and (result = '') then
  begin
    LogMsg('[WARNING] - ...Success');
    result := 'success';
  end
  else if (I = 0) and (result = 'success') then
    result := '';
end;


function TDataMod.CARDeltaActive(I : integer = 0) : string;
begin
  result := '';
  try
    qCloARDeltaActive.ExecSQL;
  except
    on E: Exception do
    begin
      if (ContainsText(E.Message,'SQL0964C  The transaction log for the database is full.  SQLSTATE=57011')) and (I < numRetries) then
      begin
        Sleep(waitLength);
        inc(I);
        logMsg('[WARNING] - Failed to run qCloARDeltaActive, transaction log full. Retrying... ');
        result := CARDeltaActive(I);
      end
      else
        result := 'Failed to run qCloARDeltaActive: ' + E.Message;
    end;
  end;
  if (I > 0) and (result = '') then
  begin
    LogMsg('[WARNING] - ...Success');
    result := 'success';
  end
  else if (I = 0) and (result = 'success') then
    result := '';
end;


function TDataMod.CalcAging(I : integer = 0) : string;
begin
  result := '';
  LogMsg('[INFO] - Starting CalcAging...');
  try
    spCalcClient.ExecProc;
  except
    on E: Exception do
    begin
      if (ContainsText(E.Message,'SQL0964C  The transaction log for the database is full.  SQLSTATE=57011')) and (I < numRetries) then
      begin
        Sleep(waitLength);
        inc(I);
        logMsg('[WARNING] - Failed to run spCalcClient, transaction log full. Retrying... ');
        result := CalcAging(I);
      end
      else
        result := 'Failed to run spCalcClient: ' + E.Message;
    end;
  end;
  if (I > 0) and (result = '') then
  begin
    LogMsg('[WARNING] - ...Success');
    result := 'success';
  end
  else if (I = 0) and (result = 'success') then
    result := '';
end;


function TDataMod.CleanDeltaTable : string;
var
  I: Integer;
  arID,balzero : string;
begin
  result := '';
//try first query qDeltaGetARIDs and if it fails log and dont do the rest
  try
    qDeltaGetARIDs.Open;
    if qDeltaGetARIDs.RecordCount > 0 then
    begin    //loop through records from first query and run qDeltaDelDupAR for each with different parameters
      for I := 0 to qDeltaGetARIDs.RecordCount - 1 do
      begin
        try
          arID := qDeltaGetARIDs.FieldByName('AR_ID').AsString;
          balzero := qDeltaGetARIDs.FieldByName('BAL_ZERO').AsString;
          qDeltaDelDupAR.ParamByName('AR_ID').Value := arID;
          qDeltaDelDupAR.ParamByName('BAL_ZERO').Value := balzero;
          qDeltaDelDupAR.ExecSQL;
          qDeltaGetARIDs.Next;
        except on E: Exception do
          begin
            LogMsg('[WARNING] - Failed to run qDeltaDelDupAR for AR_ID = '+ arID +' BAL_ZERO = '+ balzero +'. Message: '+ E.Message);
          end;
        end;
      end;
    end
    else
      result := '[WARNING] - qDeltaGetARIDs returned no records and did not make any changes to the delta table';
  except on E: Exception do
    result := '[WARNING] - qDeltaGetARIDs failed to run and did not make any changes to the delta table. Message: ' + E.Message;
  end;
end;


function TDataMod.ClearOldDelta : string;
var
  dateAR, dateClient : TDateTime;
begin
  result := '';
  dateAR := IncMonth(now,-1);
  dateClient := IncMonth(dateAR,-1);

  try //run query for client
    qClearOldClient.ParamByName('MYDATE').Value := dateClient;
    qClearOldClient.ExecSQL;
  except on E: Exception do result := '[WARNING] - qClearOldClient failed. Message: ' + E.Message;
  end;

  try //run query for AR
    qClearOldAR.ParamByName('MYDATE').Value := dateAR;
    qClearOldAR.ExecSQL;
  except on E: Exception do
    begin
      if result <> '' then
        result := result + '[SEP]';

      result := result + '[WARNING] - qClearOldAR failed. Message: ' + E.Message;
    end;
  end;
end;

{$ENDREGION}

function TDataMod.DoClientIns(ID : STRING) : boolean;
const
  labelArr : array[0..4] of string = ('CDF.HRC_ID','CDF.TYPE_OF_BUSINESS','CDF.SERVICE_DESC','CDF.BUS_IND','CDF.OP_SINCE');
var
  IDs : TArray<integer>;
  I,J : integer;
  temp,salesRep,busName,city,prov,shipID,currency,openClose, openCloseMsg, POD_REQ : string;
  tempArr,status1 : TArray<string>;
  useBillingAddr,useOpenClose,shipUse,mailUse : boolean;
begin
  //Do Main insert
  result := true;

  busName := uppercase(custData.FieldByName('NAME').asString);
  if busName = '' then busName := uppercase(custData.FieldByName('LEGAL_NAME').asString);
  if Length(busName) > 40 then Setlength(busName,40);
  //determine if we want USD
  currency := uppercase(custData.FieldByName('COUNTRY').AsString);
  if (currency = 'US') or (currency = 'USA') or (currency = 'UNITED STATES') then
    currency := 'USD'
  else
    currency := 'CAD';

  salesRep := custData.FieldByName('SALES_REP').asString; if Length(temp) > 128 then Setlength(salesRep,128); //validate sales rep here
  salesRep := ValidateSalesRep(salesRep);

  openClose := custData.FieldByName('OPEN_TIME_CLOSE_TIME').AsString;
  useOpenClose := openClose <> '';  openCloseMsg := '';

  if HandleBoolStr(custData.FieldByName('POD_REQUIRED').AsString) then POD_REQ := 'True' else POD_REQ := 'False';

  var contact := custData.FieldByName('CONTACT').asString;
  if contact = '' then begin contact := 'UNKNOWN'; LogMsg('[WARNING] - Field "CONTACT" is empty for CLIENT_ID = ' + ID + '. Setting to "UNKNOWN"') end
  else if Length(contact) > 40 then Setlength(contact,40);

  //Temporarily reverse these values as HRC resolves this issue in the extracts
  //Changed back as High Radius claims to have resolved issue
  shipUse := HandleBoolStr(custData.FieldByName('S_USE').AsString);
  mailUse := HandleBoolStr(custData.FieldByName('M_USE').AsString);
  //if shipping address exists then create another entry in CLient table
  if {(custData.FieldByName('S_USE').AsBoolean)} shipUse then
  begin
    temp := custData.FieldByName('SHIP.ADDRESS_1').asString + custData.FieldByName('SHIP.ADDRESS_2').asString + custData.FieldByName('SHIP.CITY').asString +
            custData.FieldByName('SHIP.POSTAL_CODE').asString + custData.FieldByName('SHIP.PROVINCE').asString + custData.FieldByName('SHIP.COUNTRY').asString +
            custData.FieldByName('SHIP.EMAIL_ADDRESS').asString + custData.FieldByName('SHIP.BUSINESS_PHONE').asString + custData.FieldByName('SHIP.FAX_PHONE').asString +
            custData.FieldByName('SHIP.CITY').asString;

    if temp <> '' then
    begin
      temp := custData.FieldByName('SHIP.CITY').asString;
      if temp <> '' then
      begin
        temp := uppercase(string.Join('',temp.Split(['.',',','-','_'])));
        tempArr := temp.Split([' ']);
        for I := 0 to length(tempArr) - 1 do
          for J := 1 to min(3 - length(city),length(tempArr[I])) do
            city := city + tempArr[I][J];
        //with client id - check if city name letters are in name or there is no city
        if (ContainsText(ID,city)) then
        //if yes then add a number until it works
          shipID := IncNumber(ID)
        else
        begin
          //if no then add city letters and check if used
          shipID := ID + city;
          if length(shipID) > 10 then setlength(shipID,10);
          if ExistsClientID(shipID) then
          //if used then add numbers until it works
            shipID := IncNumber(shipID);
        end;
      end
      else
        shipID := IncNumber(ID);

      if shipID <> '' then
      begin
        try
          with spCreateClient do
          begin
            Params.ClearValues;
            ParamByName('IOCLIENT_ID').Value := shipID;
            ParamByName('ILANGUAGE').Value := 'E';
            ParamByName('IDEFAULT_DELIVERY_Z').asString := GetZoneCode(shipID, true);
            ParamByName('iCUSTOMER_SINCE').Value := Date;
            ParamByName('ICLIENT_IS').Value := '111100';
            ParamByName('ICURRENCY_CODE').Value := currency;
            //NAME
            ParamByName('INAME').Value := busName;
            //ADDRESS_1
            temp := uppercase(custData.FieldByName('SHIP.ADDRESS_1').asString); if Length(temp) > 40 then Setlength(temp,40);
            ParamByName('IADDRESS_1').Value := temp;
            //ADDRESS_2
            temp := uppercase(custData.FieldByName('SHIP.ADDRESS_2').asString); if Length(temp) > 40 then Setlength(temp,40);
            ParamByName('IADDRESS_2').Value := temp;
            //CITY
            temp := uppercase(custData.FieldByName('SHIP.CITY').asString); if Length(temp) > 30 then Setlength(temp,30);
            ParamByName('ICITY').Value := temp;
            //POSTAL_CODE
            temp := uppercase(StringReplace(custData.FieldByName('SHIP.POSTAL_CODE').asString,'-','',[])); if Length(temp) > 10 then Setlength(temp,10);
            ParamByName('IPOSTAL_CODE').Value := temp;
            //PROVINCE
            //temp := uppercase(custData.FieldByName('SHIP.PROVINCE').asString); if Length(temp) > 4 then Setlength(temp,4);
            ParamByName('IPROVINCE').Value := ValidateProvince(custData.FieldByName('SHIP.PROVINCE').asString);
            //COUNTRY
            ParamByName('ICOUNTRY').Value := uppercase(custData.FieldByName('SHIP.COUNTRY').asString);
            //EMAIL_ADDRESS
            temp := custData.FieldByName('SHIP.EMAIL_ADDRESS').asString; if Length(temp) > 128 then Setlength(temp,128);
            ParamByName('IEMAIL_ADDRESS').Value := temp;
            //BUSINESS_PHONE
            temp := PrettyPhone(custData.FieldByName('SHIP.BUSINESS_PHONE').asString); if Length(temp) > 20 then Setlength(temp,20);
            ParamByName('IBUSINESS_PHONE').Value := temp;
            //FAX_PHONE
            temp := PrettyPhone(custData.FieldByName('SHIP.FAX_PHONE').asString); if Length(temp) > 20 then Setlength(temp,20);
            ParamByName('IFAX_PHONE').Value := temp;
            ExecProc;
            {$IFDEF Release}
            if useOpenClose then
            begin
              openCloseMsg := 'HRCxTM: Action Required[body]A record was created in the CLIENT table for ID "'+ShipID+
                         '" with OPEN_CLOSE_TIME value "'+openClose+'". Please insert this information manually in Truckmate.';
            end;
            {$ENDIF}
          end;
        except
          on E: Exception do
          begin
            LogMsg('[ERROR] - DoClientIns on spCreateClient for shipping account(' + shipID + '): ' + E.Message);
            exit(false);
          end;
        end;

        try
          with qClientUpd do
          begin
            Params.ClearValues;
            ParamByName('ID').value := shipID;
            ParamByName('SALES_REP').Value := salesRep;
            //ParamByName('COLLECTOR').Value := collector;                      //version 1.2.3.0 - set upon ini read and is an SQL[2] line
            ParamByName('POD_REQUIRED').Value := POD_REQ;
            ExecSQL;
          end;
        except
          on E: Exception do
          begin
            LogMsg('[ERROR] - DoClientIns on qClientUpd for shipping account(' + shipID + '): ' + E.Message);
            exit(false);
          end;
        end;
      end
      else
      begin
        LogMsg('[ERROR] - Unable to generate CLIENT_ID for shipping account for client ' + ID);
        EmailPatch(GetRecordCSV(inboundDelim,inboundSep,'ship',ID),'Unable to generate CLIENT_ID for shipping account for client ' + ID + slineBreak +
                   'The data for the client is attached to this email and called "clientData.txt". ' +
                   'Please review the client and rename the file to the desired CLIENT_ID + ".txt",' +
                   'then copy it to the folder ' + appPath + 'Patch\ on BTSSMATOOLS1.');
      end;
    end;
  end;
  //Now create main entry in CLIENT table
  try
    with spCreateClient do
    begin
      Params.ClearValues;
      ParamByName('IOCLIENT_ID').Value := ID;
      ParamByName('ILANGUAGE').Value := 'E';
      ParamByName('ICLIENT_IS').Value := '111100';
      ParamByName('ICURRENCY_CODE').Value := currency;
      //NAME
      ParamByName('INAME').Value := busName;
      //EMAIL_ADDRESS
      temp := custData.FieldByName('EMAIL_ADDRESS').asString; if Length(temp) > 128 then Setlength(temp,128);
      ParamByName('IEMAIL_ADDRESS').Value := temp;
      //BUSINESS_PHONE
      temp := PrettyPhone(custData.FieldByName('BUSINESS_PHONE').asString); if Length(temp) > 20 then Setlength(temp,20);
      ParamByName('IBUSINESS_PHONE').Value := temp;
      //FAX_PHONE
      temp := PrettyPhone(custData.FieldByName('FAX_PHONE').asString); if Length(temp) > 20 then Setlength(temp,20);
      ParamByName('IFAX_PHONE').Value := temp;
      //ADDRESS_1
      temp := uppercase(custData.FieldByName('ADDRESS_1').asString); if Length(temp) > 40 then Setlength(temp,40);
      ParamByName('IADDRESS_1').Value := temp;
      //ADDRESS_2
      temp := uppercase(custData.FieldByName('ADDRESS_2').asString); if Length(temp) > 40 then Setlength(temp,40);
      ParamByName('IADDRESS_2').Value := temp;
      //CITY
      temp := uppercase(custData.FieldByName('CITY').asString); if Length(temp) > 30 then Setlength(temp,30);
      ParamByName('ICITY').Value := temp;
      //POSTAL_CODE
      temp := uppercase(StringReplace(custData.FieldByName('POSTAL_CODE').asString,'-','',[])); if Length(temp) > 10 then Setlength(temp,10);
      ParamByName('IPOSTAL_CODE').Value := temp;
      //PROVINCE
      //temp := uppercase(custData.FieldByName('PROVINCE').asString); if Length(temp) > 4 then Setlength(temp,4);
      ParamByName('IPROVINCE').Value := ValidateProvince(custData.FieldByName('PROVINCE').asString);
      //COUNTRY
      ParamByName('ICOUNTRY').Value := uppercase(custData.FieldByName('COUNTRY').asString);
      //CONTACT
      //temp := custData.FieldByName('CONTACT').asString; if temp = '' then temp := 'UNKNOWN' else if Length(temp) > 40 then Setlength(temp,40);
      ParamByName('ICONTACT').Value := contact;
      ParamByName('IDEFAULT_DELIVERY_Z').asString := GetZoneCode(ID);
      ParamByName('iCUSTOMER_SINCE').Value := Date;
      {$IFDEF Release}
      if useOpenClose and (openCloseMsg = '') then
      begin
        openCloseMsg := 'HRCxTM: Action Required[body]A record was created in the CLIENT table for ID "'+ID+
                         '" with OPEN_CLOSE_TIME value "'+openClose+'". Please insert this information manually in Truckmate.';
       { var openClose := custData.FieldByName('OPEN_TIME_CLOSE_TIME').AsString.Split([' - ']); if length(openClose) > 0 then begin if length(openClose) = 1 then begin var hourMin := openClose[0].Split([':']); if strtoint(hourMin[0]) <= 12 then ParamByName('IOPEN_TIME').Value := openClose[0] else ParamByName('ICLOSE_TIME').Value := openClose[0]; end else begin ParamByName('IOPEN_TIME').Value := openClose[0]; ParamByName('ICLOSE_TIME').Value := openClose[1]; end; end; }
      end;
      {$ENDIF}
      ExecProc;
    end;
  except
    on E: Exception do
    begin
      LogMsg('[ERROR] - DoClientIns on spCreateClient for main entry(' + ID + '): ' + E.Message);
      exit(false);
    end;
  end;

  try
    with spUpdBilling do
    begin
      Params.ClearValues;
      ParamByName('ICLIENT_ID').Value := ID;
      //BILL_CUSTOMER                              Need to try to match a code here or else it wont go in
      temp := custData.FieldByName('BILL_CUSTOMER').asString;
      if ExistsClientID(temp) then
      begin
        var suggestion := NameSearch(temp,ID);
        if suggestion <> '' then // we cant find a client id using the name field
          EmailError('HRCxTM Notification','While creating new Client "'+ID+'", the user supplied the parent company name "'+temp+
          '". After searching the database of existing clients, the suggestion "'+suggestion+
          '" was created. If you think this is appropriate please edit the value "3rd Party Billing"in the "BIlling" tab of "Customer & Vendor Profiles" manually.',ITrecips)
        else
          EmailError('HRCxTM Notification','While creating new Client "'+ID+'", the user supplied the parent company name "'+temp+
          '". After searching the database of existing clients, there were no suggestions created. '+
          'If you do not think this is appropriate please edit the value "3rd Party Billing"in the "BIlling" tab of "Customer & Vendor Profiles" manually.',ITrecips);
      end;
      ParamByName('ICBTAX_1').Value := RevBoolStr(HandleBoolStr(custData.FieldByName('CBTAX_1').AsString));
      ExecProc;
    end;
  except
    on E: Exception do
    begin
      LogMsg('[ERROR] - DoClientIns on spUpdBilling for main entry(' + ID + '): ' + E.Message);
      exit(false);
    end;
  end;

  try
    with spUpdCredit do
    begin
      Params.ClearValues;
      ParamByName('ICLIENT_ID').Value := ID;
      ParamByName('ICREDIT_LIMIT').Value := custData.FieldByName('CREDIT_LIMIT').AsString;
      ExecProc;
    end;
  except
    on E: Exception do
    begin
      LogMsg('[ERROR] - DoClientIns on spUpdCredit (' + ID + '): ' + E.Message);
      exit(false);
    end;
  end;

  try
    with qClientUpd do
    begin
      Params.ClearValues;
      ParamByName('ID').value := ID;
      //DUNS_ID
      temp := custData.FieldByName('DUNS_ID').asString; if Length(temp) > 11 then Setlength(temp,11); ParamByName('DUNS_ID').Value := temp;
      //SALES_REP
      ParamByName('SALES_REP').Value := salesRep;

      ParamByName('CBTAX_1').Value := RevBoolStr(HandleBoolStr(custData.FieldByName('CBTAX_1').AsString));
      //POD_REQUIRED
      ParamByName('POD_REQUIRED').Value := POD_REQ;

      //Billing and mailing addressess
      //if mailing address is same as address use it as alt_address
      if {custData.FieldByName('M_USE').AsBoolean} mailUse then
      begin
        temp := custData.FieldByName('M_ALT_ADDRESS_1').asString + custData.FieldByName('M_ALT_ADDRESS_2').asString + custData.FieldByName('M_ALT_CITY').asString +
            custData.FieldByName('M_ALT_POSTAL_CODE').asString + custData.FieldByName('M_ALT_PROVINCE').asString + custData.FieldByName('M_ALT_COUNTRY').asString;
        if temp <> '' then
        begin
          //ALT_ADDRESS_1
          temp := uppercase(custData.FieldByName('M_ALT_ADDRESS_1').asString); if Length(temp) > 40 then Setlength(temp,40);
          ParamByName('ALT_ADDRESS_1').Value := temp;
          //ALT_ADDRESS_2
          temp := uppercase(custData.FieldByName('M_ALT_ADDRESS_2').asString); if Length(temp) > 40 then Setlength(temp,40);
          ParamByName('ALT_ADDRESS_2').Value := temp;
          //ALT_CITY
          temp := uppercase(custData.FieldByName('M_ALT_CITY').asString); if Length(temp) > 30 then Setlength(temp,30);
          ParamByName('ALT_CITY').Value := temp;
          //ALT_POSTAL_CODE
          temp := uppercase(StringReplace(custData.FieldByName('M_ALT_POSTAL_CODE').asString,'-','',[])); if Length(temp) > 10 then Setlength(temp,10);
          ParamByName('ALT_POSTAL_CODE').Value := temp;
          //ALT_PROVINCE
          ParamByName('ALT_PROVINCE').Value := ValidateProvince(custData.FieldByName('M_ALT_PROVINCE').asString);
          //ALT_COUNTRY
          ParamByName('ALT_COUNTRY').Value := uppercase(custData.FieldByName('M_ALT_COUNTRY').asString);
        end;
      end;
      ExecSQL;
    end;
  except
    on E: Exception do
    begin
      LogMsg('[ERROR] - DoClientIns on qInsClientUpd for main entry(' + ID + '): ' + E.Message);
      exit(false);
    end;
  end;

  //Now we can update BILL_CUSTOMER if there is a shipping account
  if shipID <> '' then
  begin
    try
      with spUpdBilling do
      begin
        Params.ClearValues;
        ParamByName('ICLIENT_ID').Value := shipID;
        //BILL_CUSTOMER  --  Need to try to match a code here or else it wont go in
        ParamByName('IBILL_CUSTOMER').Value := ID;
        ParamByName('ICBTAX_1').Value := RevBoolStr(HandleBoolStr(custData.FieldByName('CBTAX_1').AsString));
        ExecProc;
      end;
    except
      on E: Exception do
      begin
        LogMsg('[ERROR] - DoClientIns on spUpdBilling for shipping account(' + shipID + '): ' + E.Message);
        exit(false);
      end;
    end;
  end;
  //update custom defined fields
  //need to get ids for custom defined fields
  //order: HRC_ID,TYPE_OF_BUSINESS,SERVICE_DESC
  IDs := GetFieldIDs;
  for I := 0 to length(IDs) - 1 do
  begin
    try
      with qInsCust do
      begin
        Params.ClearValues;
        ParamByName('CUSTDEF_ID').Value := IDs[I];
        ParamByName('CLIENT_ID').Value := ID;
        ParamByName('DATA_STRING').Value := custData.FieldByName(labelArr[I]).Value;
        ExecSQL;
      end;
    except
      on E: Exception do
      begin
        LogMsg('[ERROR] - DoClientIns on qInsCust for client ' + ID + ' and field ' + labelArr[I] + ': ' + E.Message);
        exit(false);
      end;
    end;
  end;
  //need to put stuff in phone and phonecont tables  //update phonecont with title=eInvPDF //based on "send invoices & statements by" update with email or fax
  var conts := 0;//track records updated/created to see if we have a problem with the notifylist
  status1 := ValidateStatus_1(custData.FieldByName('STATUS_1').AsString,custData.FieldByName('EMAIL_ADDRESS').asString);
  if ( (custData.FieldByName('STATEMENT_DELIVERY').AsString = 'Email') and ( MatchStr('E:'+custData.FieldByName('PCP.EMAIL').AsString,status1) ) )
  or ( (custData.FieldByName('STATEMENT_DELIVERY').AsString = 'Fax')   and ( MatchStr('[FAX:'+custData.FieldByName('PCP.FAX_NUM').AsString,status1) ) ) then
  begin //update auto created entry in phone cont table - TITLE = eInvPDF, CONT_NAME = 'A/R', EMAIL = STATUS_1
    for var k := 0 to length(status1) - 1 do
    begin
      if status1[k] <> '' then
      begin
        var splitstatus := status1[k].Split([':']);
        var statusval : string;
        if splitstatus[0] = 'F' then
          statusval := '[FAX:' + splitstatus[1] + ']'
        else
          statusval := splitStatus[1];
        try
          with spUpdCont do  //this procedure updates the contact if it exists, creates a new one if it does not, or removes the existing if iREMOVE is set to TRUE
          begin
            Params.ClearValues;
            ParamByName('iPROFILE_TYPE').Value := 'CLIENT';
            ParamByName('iKEY_VALUE').Value := ID;
            ParamByName('iTITLE').Value := 'eInvPDF';
            ParamByName('iNAME').Value := 'A/R';
            ParamByName('iEMAIL').Value := statusval;                       //[FAX:]
            ExecProc;
            inc(conts);
          end;
        except
          on E: Exception do
          begin
            LogMsg('[ERROR] - DoClientIns on spUpdCont for (contact-client-email) (A/R-' + ID + '-' + statusval +'): ' + E.Message);
            exit(false);
          end;
        end;
      end;
    end;
    try //make another entry into phone cont table - TITLE = A/P, CONT_NAME = CONTACT, EMAIL = PCP.EMAIL, FAX_NUM = PCP.FAX_NUM, PHONENUMBER = PCP.PHONENUMBER
      with spUpdCont do
      begin
        Params.ClearValues;
        ParamByName('iPROFILE_TYPE').Value := 'CLIENT';
        ParamByName('iKEY_VALUE').Value := ID;
        ParamByName('iTITLE').Value := 'A/P';
        ParamByName('iNAME').Value := {custData.FieldByName('CONTACT').AsString} contact;
        if custData.FieldByName('PCP.EMAIL').AsString <> '' then ParamByName('iEMAIL').Value := custData.FieldByName('PCP.EMAIL').AsString;
        if custData.FieldByName('PCP.FAX_NUM').AsString <> '' then ParamByName('iFAX_NUM').Value := PrettyPhone(custData.FieldByName('PCP.FAX_NUM').AsString);
        if custData.FieldByName('PCP.PHONENUMBER').AsString <> '' then ParamByName('iPHONENUMBER').Value := PrettyPhone(custData.FieldByName('PCP.PHONENUMBER').AsString);
        ExecProc;
        inc(conts);
      end;
    except
      on E: Exception do
      begin
        LogMsg('[ERROR] - DoClientIns on spUpdCont for (contact-client) (' + {custData.FieldByName('CONTACT').AsString} contact + '-' + ID + '): ' + E.Message);
        exit(false);
      end;
    end;
  end
  else
  begin
    for var k := 0 to length(status1) - 1 do
    begin
      if status1[k] <> '' then
      begin
        var splitstatus := status1[k].Split([':']);
        var statusval : string;
        if splitstatus[0] = 'F' then
          statusval := '[FAX:' + splitstatus[1] + ']'
        else
          statusval := splitStatus[1];
        //update auto created entry in phone cont table - TITLE = eInvPDF, CONT_NAME = contact, EMAIL = STATUS_1
        try
          with spUpdCont do
          begin
            Params.ClearValues;
            ParamByName('iPROFILE_TYPE').Value := 'CLIENT';
            ParamByName('iKEY_VALUE').Value := ID;
            ParamByName('iTITLE').Value := 'eInvPDF';
            ParamByName('iNAME').Value := {custData.FieldByName('CONTACT').AsString} contact;
            ParamByName('iEMAIL').Value := statusval;
            ExecProc;
            inc(conts);
          end;
        except
          on E: Exception do
          begin
            LogMsg('[ERROR] - DoClientIns on spUpdCont for (contact-client) (' + {custData.FieldByName('CONTACT').AsString} contact + '-' + ID + '): ' + E.Message);
            exit(false);
          end;
        end;
      end;
    end;
  end;
  //if alt_contact is not blank insert into phonecont
  //cont_name=alt_contact, title=controller, email=controlleremail, phone=controllerphone
  //maybe use the procedure PROFILE_API_UPDATE_CONTACTS
  if (custData.FieldByName('ALT_CONTACT').AsString <> '') and ({custData.FieldByName('CONTACT').AsString} contact <> custData.FieldByName('ALT_CONTACT').AsString) then
  begin
    //make another entry into phone cont table - TITLE = CONTROLLER, CONT_NAME = ALT_CONTACT, EMAIL = PCC.EMAIL, PHONENUMBER = PCC.PHONENUMBER
    try
      with spUpdCont do
      begin
        Params.ClearValues;
        ParamByName('iPROFILE_TYPE').Value := 'CLIENT';
        ParamByName('iKEY_VALUE').Value := ID;
        ParamByName('iTITLE').Value := 'CONTROLLER';
        ParamByName('iNAME').Value := custData.FieldByName('ALT_CONTACT').AsString;
        if custData.FieldByName('PCC.EMAIL').AsString <> '' then ParamByName('iEMAIL').Value := custData.FieldByName('PCC.EMAIL').AsString;
        if custData.FieldByName('PCC.PHONENUMBER').AsString <> '' then ParamByName('iPHONENUMBER').Value := PrettyPhone(custData.FieldByName('PCC.PHONENUMBER').AsString);
        ExecProc;
        inc(conts);
      end;
    except
      on E: Exception do
      begin
        LogMsg('[ERROR] - DoClientIns on spUpdCont for (contact-client) (' + custData.FieldByName('ALT_CONTACT').AsString + '-' + ID + '): ' + E.Message);
        exit(false);
      end;
    end;
  end;
  //Create NotifyList records
  if conts > 0 then
  begin
    try
      try
        with qInsNL do
        begin
          ParamByName('CLIENTID').Value := ID;
          ExecSQL;
        end;
      except
        on E: Exception do
        begin
          LogMsg('[ERROR] - DoClientIns on qInsNL for "' + ID + '": ' + E.Message);
          exit(false);
        end;
      end;
      try
        qGetPhCID.ParamByName('CLIENTID').Value := ID;
        qGetPhCID.Open;
        qGetNLID.ParamByName('CLIENTID').Value := ID;
        qGetNLID.Open;
        with qInsNLDTL do
        begin
          ParamByName('NL_ID').Value := qGetNLID.FieldByName('NL_ID').Value;
          ParamByName('PHONECONTID').Value := qGetPhCID.FieldByName('PHONECONTID').Value;
          ExecSQL;
        end;
      except
        on E: Exception do
        begin
          LogMsg('[ERROR] - DoClientIns on qInsNLDTL for "' + ID + '": ' + E.Message);
          exit(false);
        end;
      end;
    finally
      if qGetPhCID.Active then
          qGetPhCID.Close;
      if qGetNLID.Active then
          qGetNLID.Close;
    end;
  end
  else
    EmailError('HRCxTM Alert','Client "'+ID+'" did not have any valid email or fax values to create entries in the PHONECONT table. No entry in CLIENT_NOTIFYLIST was created either.',ITRecips);
  if useOpenClose and (opencloseMsg <> '') then
  begin
    var openCloseArr := openCloseMsg.Split(['[body]']);
    //EmailError(openCloseArr[0],openCloseArr[1],ITRecips+acctRecips);
  end;
end;


function TDataMod.DoClientUpd : boolean;
var
  IDs : TArray<integer>;
  //I : integer;
  clientID, hrcID : string;
begin
  result := true;

  clientID := custData.FieldByName('CLIENT_ID').asString; if Length(clientID) > 10 then Setlength(clientID,10);
  hrcID := custData.FieldByName('CDF.HRC_ID').AsString;

  if clientID = '' then
  begin
    if hrcID <> '' then
    begin
      LogMsg('[WARNING] - DoClientUpd: Record has no CLIENT_ID but is associated with the HRC_ID = "' + hrcID + '"');
      clientID := GetClientID(hrcID);
      if clientID <> '' then
      begin
        LogMsg('[INFO] - DoClientUpd: Updating CLIENT_ID = "' + clientID + '"');
      end
      else
      begin
        LogMsg('[ERROR] - DoClientUpd: Record has no matching CLIENT_ID. Update procedure will not be run for this record.');
        EmailError('HRCxTM Alert',
                   'When the High Radius integtration ran it encountered a record that has an HRC_ID = "' + hrcID +
                   '" but no associated CLIENT_ID. It is likely that a CLIENT_ID needs to be entered manually and has not been entered yet.' +
                   'To do so please refer to the email alert that was sent regarding this HRC_ID',
                   acctRecips);
        result := false;
        exit
      end;
    end
    else
    begin
      LogMsg('[ERROR] - Record has no CLIENT_ID or CDF.HRC_ID. Update procedure will not be run for this record.');
      result := false;
      exit
    end;

  end;

  if not ExistsClientID(clientID) then
  begin
    LogMsg('[ERROR] - DoClientUpd: No Client exists in table for CLIENT_ID = "' + clientID + '". Update procedure will not be run for this record.');
    result := false;
    exit;
  end;
  //if credit limit in DB is same as csv limit then create credit history entry
    //else try update credit limit
  if DoSameCredLim(clientID,custData.FieldByName('CREDIT_LIMIT').AsInteger) = false then
  begin
    try
      with spUpdCredit do
      begin
        Params.ClearValues;
        ParamByName('ICLIENT_ID').Value := clientID;
        ParamByName('ICREDIT_LIMIT').Value := custData.FieldByName('CREDIT_LIMIT').Value;

        ExecProc;
      end;
    except
      on E: Exception do
      begin
        LogMsg('[ERROR] - DoClientUpd on spUpdCredit (' + clientID + '): ' + E.Message);
        result := false;
      end;
    end;
  end;
  //update custom defined fields
  //need to get ids for custom defined fields
  //order: HRC_ID,EXEMPT_TYPE,TYPE_OF_BUSINESS,INTERLINE_CODE,SERVICE_DESC,DNB_NUM
  if GetClientID(hrcID) = '' then
  begin
    IDs := GetFieldIDs;

    try
      with qUpdCust do
      begin
        Params.ClearValues;
        ParamByName('CUSTDEF_ID').Value := IDs[0];
        ParamByName('CLIENT_ID').Value := clientID;
        ParamByName('DATA_STRING').Value := custData.FieldByName('CDF.HRC_ID').Value;
        ExecSQL;
      end;
    except
      on E: Exception do
      begin
        LogMsg('[ERROR] - DoClientUpd on qInsCust for client ' + clientID + ' and field HRC_ID: ' + E.Message);
        result := false;
      end;
    end;
  end;

  if result then
    LogMsg('[INFO] - Updated Client "' + clientID + '"');
end;


function TDataMod.DoSameCredLim(CID : string; credLim : integer) : boolean;
var
  DBLim : integer;
begin
  result := false;
  try  //try run query
    qCredLim.ParamByName('CLIENTID').Value := CID;
    qCredLim.Open;
    DBLim := qCredLim.FieldByName('CREDIT_LIMIT').AsInteger;
    qCredLim.Close;
  //if credit limit from db and function parameter are same then create credit history entry
    if credLim = DBLim then
    begin
      try
        qCredLimHist.ParamByName('CLIENTID').Value := CID;
        qCredLimHist.Open;
        if qCredLimHist.IsEmpty = false then
        begin
          try
            qCredHist.Params.ClearValues;
            qCredHist.ParamByName('CLIENT_ID').Value := CID;
            qCredHist.ParamByName('CREDIT_LIMIT').Value := credLim;
            qCredHist.ExecSQL;
            result := true;
          except on E: Exception do
            begin
              LogMsg('[WARNING] - Failed to insert entry into CREDIT_HISTORY table for Client ID "' + CID + '". Message: ' + E.Message);
              EmailError('HRCxTM Alert','Failed to insert entry into CREDIT_HISTORY for Client ID "' + CID + '". Check log for error message.',ITRecips);
            end;
          end;
        end;
      finally
        qCredLimHist.Close;
      end;
    end;
  except on E: Exception do
    begin
      if qCredLim.Active then
        qCredLim.Close;
      LogMsg('[WARNING] - Failed to get credit limit for Client ID "' + CID + '". Message: ' + E.Message);
      EmailError('HRCxTM Alert','Failed to get credit limit for Client ID "' + CID + '". Unable to determine whether an entry should be created' +
                 ' in table CREDIT_HISTORY. Please check manually if CLIENT table value for CREDIT_LIMIT was the same as value in extract.',ITRecips);
    end;
  end;
end;


//order: HRC_ID,TYPE_OF_BUSINESS,SERVICE_DESC,BUS_IND,OP_SINCE
function TDataMod.GetFieldIDs : TArray<integer>;
const
  labelArr : array[0..4] of string = ('HRC_ID','TYPE_OF_BUSINESS','SERVICE_DESC','BUS_IND','OP_SINCE');
var                                                                                          //If this is added back in the date format needs to be considered
  I : integer;                                                                               //basically another validate function
begin
  SetLength(result, 5);

  for I := 0 to 4 do
  begin
    try
      with qGetCustID do
      begin
        Params.ClearValues;
        ParamByName('LABEL_NAME').Value := labelArr[I];
        Open;
        result[I] := FieldByName('CUSTDEF_ID').AsInteger;
        Close;
      end;
    except
      on E: Exception do
      begin
        LogMsg('[ERROR] - qGetCustID for field ' + labelArr[I] + ': ' + E.Message);
      end;
    end;
  end;
end;


procedure TDataMod.CSVtoDS(csvPath,sep,delim : string; insert : boolean);
var
  reader : TFDBatchMoveTextReader;
  writer : TFDBatchMoveDataSetWriter;
  mapping : TFDBatchMoveMappings;
begin
  //if custData.RecordCount <> 0 then
  if custData.Active then
  begin
    try
      custData.EmptyDataSet;
      custData.Close;
    except
      on E: Exception do
      begin
        LogMsg('[ERROR] - CSVtoDS(' + csvPath + ') - trying to reset custData: ' + E.Message);
        exit
      end;
    end;
  end;

  try
    reader := TFDBatchMoveTextReader.Create(bMove);
    reader.FileName := csvPath;
    //reader.DataDef.Delimiter := delim[1];
    if delim <> '' then
      reader.DataDef.Delimiter := delim[1]
    else
      reader.DataDef.Delimiter := #0; 
    reader.DataDef.Separator := sep[1];
    reader.DataDef.WithFieldNames := true;
  except
    on E: Exception do
    begin
      LogMsg('[ERROR] - CSVtoDS(' + csvPath + ') - setting up reader: ' + E.Message);
      if reader <> nil then reader.Free;
      exit
    end;
  end;

  try
    writer := TFDBatchMoveDataSetWriter.Create(bMove);
    writer.DataSet := custData;
    //writer.Optimise := true;
  except
    on E: Exception do
    begin
      LogMsg('[ERROR] - CSVtoDS(' + csvPath + ') - setting up writer: ' + E.Message);
      if reader <> nil then reader.free;
      if writer <> nil then writer.Free;
      exit
    end;
  end;


  try
    mapping := TFDBatchMoveMappings.Create(bMove);
    if insert then
      CreateMapping(mapping)
    else
      UpdateMapping(mapping);
    bMove.Mappings := mapping;
  except
    on E: Exception do
    begin
      LogMsg('[ERROR] - CSVtoDS(' + csvPath + ') - creating mapping: ' + E.Message);
      if reader <> nil then reader.free;
      if writer <> nil then writer.Free;
      if mapping <> nil then mapping.Free;      
      exit
    end;
  end;

  try
    bMove.GuessFormat;
    bMove.Execute;

    custData.Active := true;
  except
    on E: Exception do
    begin
      LogMsg('[ERROR] - CSVtoDS(' + csvPath + ') - executing BatchMove: ' + E.Message);
    end;
  end;

  try
    if reader <> nil then
      reader.Free;
    if writer <> nil then
      writer.Free;
    if mapping <> nil then
      mapping.Free;
  except
    on E:Exception do
    begin
      LogMsg('[ERROR] - Trying to close BatchMove objects: ' + E.Message);
    end;
  end;
end;


procedure TDataMod.CreateMapping(map : TFDBatchMoveMappings);
var
  mapitem : TFDBatchMoveMappingItem;
begin
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'CDF.HRC_ID'; DestinationFieldName := 'CDF.HRC_ID' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'DUNS_ID'; DestinationFieldName := 'DUNS_ID' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'CREDIT_LIMIT'; DestinationFieldName := 'CREDIT_LIMIT' end;

  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'SALES_REP'; DestinationFieldName := 'SALES_REP' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'BILL_CUSTOMER'; DestinationFieldName := 'BILL_CUSTOMER' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'CBTAX_1'; DestinationFieldName := 'CBTAX_1' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'CDF.TYPE_OF_BUSINESS'; DestinationFieldName := 'CDF.TYPE_OF_BUSINESS' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'CDF.SERVICE_DESC'; DestinationFieldName := 'CDF.SERVICE_DESC' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'CDF.BUS_IND'; DestinationFieldName := 'CDF.BUS_IND' end;

  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'LEGAL_NAME'; DestinationFieldName := 'LEGAL_NAME' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'NAME'; DestinationFieldName := 'NAME' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'ADDRESS_1'; DestinationFieldName := 'ADDRESS_1' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'ADDRESS_2'; DestinationFieldName := 'ADDRESS_2' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'CITY'; DestinationFieldName := 'CITY' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'PROVINCE'; DestinationFieldName := 'PROVINCE' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'POSTAL_CODE'; DestinationFieldName := 'POSTAL_CODE' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'COUNTRY'; DestinationFieldName := 'COUNTRY' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'EMAIL_ADDRESS'; DestinationFieldName := 'EMAIL_ADDRESS' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'BUSINESS_PHONE'; DestinationFieldName := 'BUSINESS_PHONE' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'FAX_PHONE'; DestinationFieldName := 'FAX_PHONE' end;

  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'M_USE'; DestinationFieldName := 'M_USE' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'M_ALT_ADDRESS_1'; DestinationFieldName := 'M_ALT_ADDRESS_1' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'M_ALT_ADDRESS_2'; DestinationFieldName := 'M_ALT_ADDRESS_2' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'M_ALT_CITY'; DestinationFieldName := 'M_ALT_CITY' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'M_ALT_PROVINCE'; DestinationFieldName := 'M_ALT_PROVINCE' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'M_ALT_POSTAL_CODE'; DestinationFieldName := 'M_ALT_POSTAL_CODE' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'M_ALT_COUNTRY'; DestinationFieldName := 'M_ALT_COUNTRY' end;

  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'S_USE'; DestinationFieldName := 'S_USE' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'SHIP.ADDRESS_1'; DestinationFieldName := 'SHIP.ADDRESS_1' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'SHIP.ADDRESS_2'; DestinationFieldName := 'SHIP.ADDRESS_2' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'SHIP.CITY'; DestinationFieldName := 'SHIP.CITY' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'SHIP.PROVINCE'; DestinationFieldName := 'SHIP.PROVINCE' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'SHIP.POSTAL_CODE'; DestinationFieldName := 'SHIP.POSTAL_CODE' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'SHIP.COUNTRY'; DestinationFieldName := 'SHIP.COUNTRY' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'SHIP.EMAIL_ADDRESS'; DestinationFieldName := 'SHIP.EMAIL_ADDRESS' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'SHIP.BUSINESS_PHONE'; DestinationFieldName := 'SHIP.BUSINESS_PHONE' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'SHIP.FAX_PHONE'; DestinationFieldName := 'SHIP.FAX_PHONE' end;

  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'OPEN_TIME_CLOSE_TIME'; DestinationFieldName := 'OPEN_TIME_CLOSE_TIME' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'OP_SINCE'; DestinationFieldName := 'CDF.OP_SINCE' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'STATEMENT_DELIVERY'; DestinationFieldName := 'STATEMENT_DELIVERY' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'STATUS_1'; DestinationFieldName := 'STATUS_1' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'POD_REQUIRED'; DestinationFieldName := 'POD_REQUIRED' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'CONTACT'; DestinationFieldName := 'CONTACT' end;

  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'PCP.EMAIL'; DestinationFieldName := 'PCP.EMAIL' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'PCP.PHONENUMBER'; DestinationFieldName := 'PCP.PHONENUMBER' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'PCP.FAX_NUM'; DestinationFieldName := 'PCP.FAX_NUM' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'ALT_CONTACT'; DestinationFieldName := 'ALT_CONTACT' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'PCC.EMAIL'; DestinationFieldName := 'PCC.EMAIL' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'PCC.PHONENUMBER'; DestinationFieldName := 'PCC.PHONENUMBER' end;

  custData.FieldDefs.Clear;
  custData.Fields.Clear;
  //Now set up fields in Dataset
  with CustData.FieldDefs do
  begin
    with AddFieldDef do begin DataType := ftString;Name := 'DUNS_ID';Size := 11; end;
    with AddFieldDef do begin DataType := ftString;Name := 'CDF.HRC_ID';Size := 15; end;
    with AddFieldDef do begin DataType := ftFloat;Name := 'CREDIT_LIMIT'; end;

    with AddFieldDef do begin DataType := ftString;Name := 'SALES_REP';Size := 128; end;
    with AddFieldDef do begin DataType := ftString;Name := 'BILL_CUSTOMER';Size := 10; end;
    with AddFieldDef do begin DataType := ftString;Name := 'CBTAX_1';Size := 15; end;
    with AddFieldDef do begin DataType := ftString;Name := 'CDF.TYPE_OF_BUSINESS';Size := 255; end;
    with AddFieldDef do begin DataType := ftString;Name := 'CDF.SERVICE_DESC';Size := 255; end;
    with AddFieldDef do begin DataType := ftString;Name := 'CDF.BUS_IND';Size := 255; end;

    with AddFieldDef do begin DataType := ftString;Name := 'LEGAL_NAME';Size := 40; end;
    with AddFieldDef do begin DataType := ftString;Name := 'NAME';Size := 40; end;
    with AddFieldDef do begin DataType := ftString;Name := 'ADDRESS_1';Size := 40; end;
    with AddFieldDef do begin DataType := ftString;Name := 'ADDRESS_2';Size := 40; end;
    with AddFieldDef do begin DataType := ftString;Name := 'CITY';Size := 30; end;
    with AddFieldDef do begin DataType := ftString;Name := 'PROVINCE';Size := 20; end;
    with AddFieldDef do begin DataType := ftString;Name := 'POSTAL_CODE';Size := 10; end;
    with AddFieldDef do begin DataType := ftString;Name := 'COUNTRY';Size := 2; end;
    with AddFieldDef do begin DataType := ftString;Name := 'EMAIL_ADDRESS';Size := 128; end;
    with AddFieldDef do begin DataType := ftString;Name := 'BUSINESS_PHONE';Size := 20; end;
    with AddFieldDef do begin DataType := ftString;Name := 'FAX_PHONE';Size := 20; end;

    with AddFieldDef do begin DataType := ftBoolean;Name := 'M_USE'; end;
    with AddFieldDef do begin DataType := ftString;Name := 'M_ALT_ADDRESS_1';Size := 40; end;
    with AddFieldDef do begin DataType := ftString;Name := 'M_ALT_ADDRESS_2';Size := 40; end;
    with AddFieldDef do begin DataType := ftString;Name := 'M_ALT_CITY';Size := 30; end;
    with AddFieldDef do begin DataType := ftString;Name := 'M_ALT_PROVINCE';Size := 20; end;
    with AddFieldDef do begin DataType := ftString;Name := 'M_ALT_POSTAL_CODE';Size := 10; end;
    with AddFieldDef do begin DataType := ftString;Name := 'M_ALT_COUNTRY';Size := 2; end;

    with AddFieldDef do begin DataType := ftBoolean;Name := 'S_USE'; end;
    with AddFieldDef do begin DataType := ftString;Name := 'SHIP.ADDRESS_1';Size := 40; end;
    with AddFieldDef do begin DataType := ftString;Name := 'SHIP.ADDRESS_2';Size := 40; end;
    with AddFieldDef do begin DataType := ftString;Name := 'SHIP.CITY';Size := 30; end;
    with AddFieldDef do begin DataType := ftString;Name := 'SHIP.PROVINCE';Size := 20; end;
    with AddFieldDef do begin DataType := ftString;Name := 'SHIP.POSTAL_CODE';Size := 10; end;
    with AddFieldDef do begin DataType := ftString;Name := 'SHIP.COUNTRY';Size := 2; end;
    with AddFieldDef do begin DataType := ftString;Name := 'SHIP.EMAIL_ADDRESS';Size := 128; end;
    with AddFieldDef do begin DataType := ftString;Name := 'SHIP.BUSINESS_PHONE';Size := 20; end;
    with AddFieldDef do begin DataType := ftString;Name := 'SHIP.FAX_PHONE';Size := 20; end;

    with AddFieldDef do begin DataType := ftString;Name := 'OPEN_TIME_CLOSE_TIME';Size := 40; end;
    with AddFieldDef do begin DataType := ftString;Name := 'CDF.OP_SINCE';Size := 255 end;
    with AddFieldDef do begin DataType := ftString;Name := 'STATEMENT_DELIVERY';Size := 5; end;
    with AddFieldDef do begin DataType := ftString;Name := 'STATUS_1';Size := 40; end;
    with AddFieldDef do begin DataType := ftString;Name := 'POD_REQUIRED';Size := 15 end;
    with AddFieldDef do begin DataType := ftString;Name := 'CONTACT';Size := 40; end;

    with AddFieldDef do begin DataType := ftString;Name := 'PCP.EMAIL';Size := 128; end;
    with AddFieldDef do begin DataType := ftString;Name := 'PCP.PHONENUMBER';Size := 20; end;
    with AddFieldDef do begin DataType := ftString;Name := 'PCP.FAX_NUM';Size := 20; end;
    with AddFieldDef do begin DataType := ftString;Name := 'ALT_CONTACT';Size := 40; end;
    with AddFieldDef do begin DataType := ftString;Name := 'PCC.EMAIL';Size := 128; end;
    with AddFieldDef do begin DataType := ftString;Name := 'PCC.PHONENUMBER';Size := 20; end;

  end;
  CustData.CreateDataSet;
end;


procedure TDataMod.UpdateMapping(map : TFDBatchMoveMappings);
begin
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'CLIENT_ID'; DestinationFieldName := 'CLIENT_ID' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'Approved Credit Limit'; DestinationFieldName := 'CREDIT_LIMIT' end;
  with TFDBatchMoveMappingItem.Create(map) do begin SourceFieldName := 'CDF.HRC_ID'; DestinationFieldName := 'CDF.HRC_ID' end;

  custData.FieldDefs.Clear;
  custData.Fields.Clear;
  //Now set up fields in Dataset
  with CustData.FieldDefs do
  begin
    with AddFieldDef do begin DataType := ftString;Name := 'CLIENT_ID';Size := 10; end;
    with AddFieldDef do begin DataType := ftFloat;Name := 'CREDIT_LIMIT'; end;
    with AddFieldDef do begin DataType := ftString;Name := 'CDF.HRC_ID';Size := 15; end;
  end;
  CustData.CreateDataSet;
end;


function TDataMod.GetRecordCSV(delim,sep,kind,ID : string) : string;
var
  d,s : string;
begin
  d := delim; s := sep;
  try
    with custData do
    begin
      if kind = 'ins' then
      begin
        result := d + FieldByName('CDF.HRC_ID').asString+d+s+FieldByName('CREDIT_LIMIT').asString+s+d+
                  FieldByName('LEGAL_NAME').asString+d+s+d+FieldByName('NAME').asString+d+s+d+
                  FieldByName('ADDRESS_1').asString+d+s+d+FieldByName('ADDRESS_2').asString+d+s+d+FieldByName('CITY').asString+d+s+d+FieldByName('PROVINCE').asString+d+s+d+FieldByName('POSTAL_CODE').asString+d+s+d+FieldByName('COUNTRY').asString+d+s+d+FieldByName('EMAIL_ADDRESS').asString+d+s+d+FieldByName('BUSINESS_PHONE').asString+d+s+d+FieldByName('FAX_PHONE').asString+d+s+d+
                  FieldByName('CDF.TYPE_OF_BUSINESS').asString+d+s+d+FieldByName('DUNS_ID').asString+d+s+d+FieldByName('BILL_CUSTOMER').asString+d+s+
                  FieldByName('M_USE').asString+s+d+FieldByName('M_ALT_ADDRESS_1').asString+d+s+d+FieldByName('M_ALT_ADDRESS_2').asString+d+s+d+FieldByName('M_ALT_CITY').asString+d+s+d+FieldByName('M_ALT_PROVINCE').asString+d+s+d+FieldByName('M_ALT_POSTAL_CODE').asString+d+s+d+FieldByName('M_ALT_COUNTRY').asString+d+s+d+
                  FieldByName('CDF.SERVICE_DESC').asString+d+s+
                  FieldByName('S_USE').asString+s+d+FieldByName('SHIP.ADDRESS_1').asString+d+s+d+FieldByName('SHIP.ADDRESS_2').asString+d+s+d+FieldByName('SHIP.CITY').asString+d+s+d+FieldByName('SHIP.PROVINCE').asString+d+s+d+FieldByName('SHIP.POSTAL_CODE').asString+d+s+d+FieldByName('SHIP.COUNTRY').asString+d+s+d+FieldByName('SHIP.EMAIL_ADDRESS').asString+d+s+d+FieldByName('SHIP.BUSINESS_PHONE').asString+d+s+d+FieldByName('SHIP.FAX_PHONE').asString+d+s+d+
                  FieldByName('OPEN_TIME_CLOSE_TIME').asString+d+s+FieldByName('CBTAX_1').asString+s+d+FieldByName('CDF.BUS_IND').asString+d+s+FieldByName('CDF.OP_SINCE').asString+s+d+FieldByName('SALES_REP').asString+d+s+d+FieldByName('STATEMENT_DELIVERY').asString+d+s+d+FieldByName('STATUS_1').asString+d+s+FieldByName('POD_REQUIRED').asString+s+d+FieldByName('CONTACT').asString+d+s+d+
                  FieldByName('PCP.EMAIL').asString+d+s+d+FieldByName('PCP.PHONENUMBER').asString+d+s+d+FieldByName('PCP.FAX_NUM').asString+d+s+d+FieldByName('ALT_CONTACT').asString+d+s+d+FieldByName('PCC.EMAIL').asString+d+s+d+FieldByName('PCC.PHONENUMBER').asString+d+s+
                  slinebreak;
      end
      else if kind = 'ship' then
      begin
        result := d+d+s+s+d+FieldByName('LEGAL_NAME').asString+d+s+d+FieldByName('NAME').asString+d+s+d+FieldByName('ADDRESS_1').asString+d+s+d+
                  FieldByName('ADDRESS_2').asString+d+s+d+FieldByName('CITY').asString+d+s+d+FieldByName('PROVINCE').asString+d+s+d+FieldByName('POSTAL_CODE').asString+d+s+d+
                  FieldByName('COUNTRY').asString+d+s+d+FieldByName('EMAIL_ADDRESS').asString+d+s+d+FieldByName('BUSINESS_PHONE').asString+d+s+d+FieldByName('FAX_PHONE').asString+d+s+d+
                  d+s+d+d+s+d+ID+d+s+s+d+d+s+d+d+s+d+d+s+d+d+s+d+d+s+d+d+s+d+d+s+s+d+d+s+d+d+s+d+d+s+d+d+s+d+d+s+d+d+s+d+d+s+d+d+s+d+d+s+d+
                  FieldByName('OPEN_TIME_CLOSE_TIME').asString+d+s+s+d+d+s+s+d+FieldByName('SALES_REP').asString+d+s+d+d+s+d+d+s+s+d+d+s+d+d+s+d+d+s+d+d+s+d+d+s+d+d+s+d+d+s+slinebreak;
      end
      else if kind = 'upd' then
      begin
        result := d + FieldByName('CLIENT_ID').asString+d+s+FieldByName('CREDIT_LIMIT').asString+s+d+FieldByName('CDF.HRC_ID').asString+d+s+slinebreak;
      end
      else
      begin
        result := '';
        LogMsg('[ERROR] - Invalid type for GetRecordCSV: ' + kind);
      end;
    end;
  except
    on E: Exception do
    begin
      result := '';
      LogMsg('[ERROR] - GetRecordCSV: ' + E.Message);
    end;
  end;
end;


function TDataMod.GetZoneCode(ClientID : string; ship : boolean = false) : string;
var
  city,prov,country  : string;
begin
  result := '';
  //Check if there is a record we can operate on
  if (custData.RecordCount <> 0) and (custData.Eof = false) then
  begin
    if ship then
    begin
      city := custData.FieldByName('SHIP.CITY').AsString;
      prov := ValidateProvince(custData.FieldByName('SHIP.PROVINCE').AsString);
    end
    else
    begin
      city := custData.FieldByName('CITY').AsString;
      prov := ValidateProvince(custData.FieldByName('PROVINCE').AsString);
    end;
    //check that city, province exist
    if (city <> '') and (prov <> '') then
    begin
      //strip punctuation from city province and country and uppercase
       city := uppercase(stringReplace(city,'.','',[rfReplaceAll])) + '%';
       prov := '%' + uppercase(stringReplace(prov,'.','',[rfReplaceAll])) + '%';

      //Run first query
      try
        qFindZone.ParamByName('CITY').Value := city;
        qFindZone.ParamByName('PROV').Value := prov;
        qFindZone.Open;
      except
        on E: Exception do
        begin
          result := 'PENDING';
          LogMsg('[ERROR] - Opening qFindZone for CLIENT_ID "' + clientID + '": ' + E.Message);
          EmailError('HRCxTM alert',
                   'ZONE_ID set to "' + result + '" for CLIENT_ID = ' + CLIENTID +
                   '. If this is not appropriate it should be corrected as this may affect dispatch and rating',acctrecips);
          exit;
        end;
      end;

      //If we try to get ZONE_ID and it exists we use it
      if qFindZone.FieldByName('ZONE_ID').AsString <> '' then
        result := qFindZone.FieldByName('ZONE_ID').AsString;

      qFindZone.Close;
    end;

    if result = '' then
    //If first method didtn't work or didn't happen run query to check if there is a zone code for the province
    begin
      try
        prov := stringReplace(prov,'%','',[rfReplaceAll]);
        qFindProvZone.ParamByName('PROV').Value := prov;
        qFindProvZone.Open;
      except
        on E: Exception do
        begin
          result := 'PENDING';
          LogMsg('[ERROR] - Opening qFindProvZone for CLIENT_ID "' + clientID + '": ' + E.Message);
          EmailError('HRCxTM alert',
                   'ZONE_ID set to "' + result + '" for CLIENT_ID = ' + CLIENTID +
                   '. If this is not appropriate it should be corrected as this may affect dispatch and rating',acctrecips);
          exit;
        end;
      end;

      //if we look at ZONE_ID and it exists  result = ZONE_ID
      if qFindProvZone.FieldByName('ZONE_ID').AsString <> '' then
        result := qFindProvZone.FieldByName('ZONE_ID').AsString
      else
      begin
        if ship then
          country := uppercase(custData.FieldByName('SHIP.COUNTRY').AsString)
        else
          country := uppercase(custData.FieldByName('COUNTRY').AsString);

        if (country = 'CA') or (country = 'CAN') or (country = 'CANADA') then
          result := 'CANADA'
        else if (country = 'US') or (country = 'USA') or (country = 'UNITED STATES') then
          result := 'USA'
        else
          result := 'PENDING';
      end;
      qFindProvZone.Close;

      LogMsg('[INFO] - ZONE_ID set to "' + result + '" for CLIENT_ID = ' + CLIENTID);
      //send an email to accounting. This currently does not actually send
      EmailError('HRCxTM alert',
                   'ZONE_ID set to "' + result + '" for CLIENT_ID = ' + CLIENTID +
                   '. If this is not appropriate it should be corrected as this may affect dispatch and rating',acctrecips);
    end;
  end;
end;


function TDataMod.IncNumber(ClientID : string) : string;
var
  I,ind, num : integer;
begin
  result := '';
  ind := length(ClientID) - 1;
  num := 0;
  if ( IsNumber(ClientID[ind]) ) then
  begin
    num := strtoint(String(ClientID[ind])) + 1;
    ClientID[ind] := inttostr(num)[1];
  end
  else
  begin
    ClientID := CLientID + '0';
    inc(ind);
  end;

  if ExistsClientID(ClientID) then
  begin
    for I := (num + 1) to 9 do
    begin
      ClientID[ind] := inttostr(I)[1];
      if ExistsClientID(ClientID) = false then
      begin
        result := ClientID;
        exit;
      end
    end;
  end
  else
    result := ClientID;
end;


function TDataMod.ExistsClientID(clientID : string) : boolean;
begin
  result := false;

  try
    qClientID.ParamByName('clientid').Value := clientID;
    qClientID.Open;

    if qClientID.FieldByName('client_id').Value = clientID then
      result := true;
    
    clientName := qClientID.FieldByName('name').asstring;
    clientCity := qClientID.FieldByName('city').asstring;

    qClientID.Close;
  except
    on E: Exception do
    begin
      LogMsg('[ERROR] - ExistsClientID for CLIENT_ID "' + clientID + '": ' + E.Message);
    end;

  end;
end;


function TDataMod.GetClientID(hrcID : string) : string;
begin
  try
    qHRCID.Params[0].Value := hrcID;
    qHRCID.Open;

    if qHRCID.RecordCount > 0 then
      result := qHRCID.FieldByName('SRC_TABLE_KEY').AsString
    else
      result := '';

    qHRCID.Close;

  except
    on E:Exception do
    begin
      LogMsg('[ERROR] - GetClientID for HRC_ID "' + hrcID + '": ' + E.Message);
      result := '';
    end;
  end;

end;


function TDataMod.NameSearch(name, client_id : string) : string;
begin
  try
    var modName := uppercase(trim(name));
    modName := string.Join('',(modName.Split([' LTD',' INCORPORATED',' INC',' CORPORATION',' COMPANY',' CORP','.','!','@','#','$','%','^','*','(',')','+','=','<','>','?','/','\','|',';',':','''','"','`','~','[',']','{','}'])));
    modName := string.Join(' ',modName.Split(['  ']));
    var words := modName.Split([' ']);
    var nonempty : TArray<string>;
    for var I := 0 to length(words) - 1 do
      if words[I] <> '' then
      begin
        setlength(nonempty,length(nonempty) + 1);
        nonempty[length(nonempty)-1] := words[I];
      end;

    if length(nonempty) = 0 then
      exit('');
         
    modName := '%' + nonempty[0] + '%';
    qNameSearch.Params[0].Value := modName;
    qNameSearch.Params[1].Value := client_id;
    qNameSearch.Open;

    if (qNameSearch.RecordCount > 0) then
    begin
      if (qNameSearch.RecordCount < 3) then
      begin
        result := qNameSearch.FieldByName('CLIENT_ID').AsString;
        qNameSearch.Close;
      end
      else
      begin
        qNameSearch.Close;
        if length(nonempty) > 1 then
          modName := '%' + nonempty[0] + ' ' + nonempty[1] + '%'
        else
          exit('');
        qNameSearch.Params[0].Value := modName;
        try
          qNameSearch.Open;
          result := qNameSearch.FieldByName('CLIENT_ID').AsString;
        finally
           qNameSearch.Close;
        end;
      end;
    end
    else
    begin
      result := '';
      qNameSearch.Close;
    end;
  except
    on E:Exception do
    begin
      LogMsg('[ERROR] - NameSearch for name "' + name + '": ' + E.Message);
      result := '';
    end;
  end;
end;



function TDataMod.GetClientName : string;
begin
  result := clientName;
end;


function TDataMod.GetClientCity : string;
begin
  result := clientCity;
end;


function TDataMod.ChangeClientID(oldClientID,newClientID,name : string) : boolean;
begin
  result := true;
  try
    DB.StartTransaction;

    qChgID_Cli.ParamByName('OLD_CLIENT_ID').Value := oldClientID;
    qChgID_Cli.ParamByName('NEW_CLIENT_ID').Value := newClientID;
    qChgID_Cli.ParamByName('CLIENT_NAME').Value := name;
    qChgID_Cli.ExecSQL;

    qChgID_CliAR.ParamByName('OLD_CLIENT_ID').Value := oldClientID;
    qChgID_CliAR.ParamByName('NEW_CLIENT_ID').Value := newClientID;
    qChgID_CliAR.ExecSQL;

    qChgID_CliBal.ParamByName('OLD_CLIENT_ID').Value := oldClientID;
    qChgID_CliBal.ParamByName('NEW_CLIENT_ID').Value := newClientID;
    qChgID_CliBal.ExecSQL;

    qChgID_CliStatus.ParamByName('OLD_CLIENT_ID').Value := oldClientID;
    qChgID_CliStatus.ParamByName('NEW_CLIENT_ID').Value := newClientID;
    qChgID_CliStatus.ExecSQL;

    qChgID_CredHist.ParamByName('OLD_CLIENT_ID').Value := oldClientID;
    qChgID_CredHist.ParamByName('NEW_CLIENT_ID').Value := newClientID;
    qChgID_CredHist.ExecSQL;

    qChgID_CustData.ParamByName('OLD_CLIENT_ID').Value := oldClientID;
    qChgID_CustData.ParamByName('NEW_CLIENT_ID').Value := newClientID;
    qChgID_CustData.ExecSQL;

    qChgID_Phone.ParamByName('OLD_CLIENT_ID').Value := oldClientID;
    qChgID_Phone.ParamByName('NEW_CLIENT_ID').Value := newClientID;
    qChgID_Phone.ParamByName('CLIENT_NAME').Value := name;
    qChgID_Phone.ExecSQL;

    qChgID_Delta.ParamByName('OLD_CLIENT_ID').Value := oldClientID;
    qChgID_Delta.ParamByName('NEW_CLIENT_ID').Value := newClientID;
    qChgID_Delta.ExecSQL;


    qChgID_TLcust.ParamByName('OLD_CLIENT_ID').Value := oldClientID;
    qChgID_TLcust.ParamByName('NEW_CLIENT_ID').Value := newClientID;
    qChgID_TLcust.ParamByName('CLIENT_NAME').Value := name;
    qChgID_TLcust.ExecSQL;

    qChgID_TLorig.ParamByName('OLD_CLIENT_ID').Value := oldClientID;
    qChgID_TLorig.ParamByName('NEW_CLIENT_ID').Value := newClientID;
    qChgID_TLorig.ParamByName('CLIENT_NAME').Value := name;
    qChgID_TLorig.ExecSQL;

    qChgID_TLdest.ParamByName('OLD_CLIENT_ID').Value := oldClientID;
    qChgID_TLdest.ParamByName('NEW_CLIENT_ID').Value := newClientID;
    qChgID_TLdest.ParamByName('CLIENT_NAME').Value := name;
    qChgID_TLdest.ExecSQL;

    qChgID_TLcont.ParamByName('OLD_CLIENT_ID').Value := oldClientID;
    qChgID_TLcont.ParamByName('NEW_CLIENT_ID').Value := newClientID;
    qChgID_TLcont.ParamByName('CLIENT_NAME').Value := name;
    qChgID_TLcont.ExecSQL;

    qChgID_TLbill.ParamByName('OLD_CLIENT_ID').Value := oldClientID;
    qChgID_TLbill.ParamByName('NEW_CLIENT_ID').Value := newClientID;
    qChgID_TLbill.ParamByName('CLIENT_NAME').Value := name;
    qChgID_TLbill.ExecSQL;

    qChgID_TLcare.ParamByName('OLD_CLIENT_ID').Value := oldClientID;
    qChgID_TLcare.ParamByName('NEW_CLIENT_ID').Value := newClientID;
    qChgID_TLcare.ParamByName('CLIENT_NAME').Value := name;
    qChgID_TLcare.ExecSQL;

    qChgID_TLpickup.ParamByName('OLD_CLIENT_ID').Value := oldClientID;
    qChgID_TLpickup.ParamByName('NEW_CLIENT_ID').Value := newClientID;
    qChgID_TLpickup.ParamByName('CLIENT_NAME').Value := name;
    qChgID_TLpickup.ExecSQL;

    DB.Commit;
  except on E:EDB2NativeException do
    begin
      var qString := 'on "' + (E as EDB2NativeException).FDObjName + '"';
      DB.Rollback;
      result := false;
      LogMsg('[ERROR] - ChangeClientID('+oldClientID+','+newClientID+','+name+') '+qString+' E: ' + E.Message);
    end;
  end;
end;


function TDataMod.ImpactAssess(ClientID : string) : string;
begin
  result := '';

  try
    qChgImpact.ParamByName('CLIENT_ID').Value := clientID;
    qChgImpact.Open;
    var total := qChgImpact.FieldByName('TOT_COUNT').AsInteger;
    var billed := qChgImpact.FieldByName('BILLED').AsInteger;
    qChgImpact.Close;
    if total > 5 then
      result := inttostr(total) + ' records in TLORDER '  ;
    if billed > 0 then
    begin
      if result <> '' then
        result := result + 'and ';
      result := result + inttostr(billed) + ' records that have already been billed ';
    end;
    if result <> '' then
      result := result + 'for CLIENT_ID = "' + clientID + '"';

  except on E:EDB2NativeException do
    begin
      LogMsg('[ERROR] - ImpactAssess('+ClientID+') '+E.FDObjName+' E: '+E.Message);
      if qChgImpact.Active then qChgImpact.Close;      
    end;
  end;
end;


function TDataMod.ValidateSalesRep(salesRep : string) : string;
begin
  //Search in USERS table for USER_ID with salesRep
  try
    with qUserID do
    begin

      ParamByName('USERID').Value := salesRep;
      Open;
    end;
  except
    on E: Exception do
    begin
      LogMsg('[ERROR] - ValidateSalesRep(' + salesRep + ') using qUserID: ' + E.Message);
      result := '';
      exit;
    end;
  end;

  if (qUserID.RecordCount > 0) and (qUserID.FieldByName('USER_ID').AsString <> '') then
  //If it exists then return it
    result := salesRep
  else
  //If not search USERS for USERNAME with salesRep
  begin
    try
      with qUserName do
      begin
        ParamByName('USERNAME').Value := salesRep;
        Open;
      end;
    except
      on E: Exception do
      begin
        LogMsg('[ERROR] - ValidateSalesRep(' + salesRep + ') using qUserID: ' + E.Message);
        result := '';
        exit;
      end;
    end;

    if (qUserName.RecordCount > 0) and (qUserName.FieldByName('USER_ID').AsString <> '') then
      result := qUserName.FieldByName('USER_ID').AsString
    else
    begin
      result := '';
      //And send an email?
    end;
    qUserName.Close;
  end;

  qUserID.Close;
end;


function TDataMod.ValidateProvince(province : string) : string;
var
  prov : string;
begin
  result := '';
  if length(province) < 3 then
  begin
    prov := uppercase(province);
    try
      qValidProvShort.Params[0].Value := prov;
      qValidProvShort.Open;
      if (qValidProvShort.RecordCount > 0) and (qValidProvShort.FieldByName('ZONE_ID').AsString <> '') then
        result := qValidProvShort.FieldByName('ZONE_ID').AsString
      else
        result := prov;
     qValidProvShort.Close;
    except
      on E:Exception do
      begin
        LogMsg('[ERROR] - ValidateProvince('+province+') running query qValidProvShort: ' + E.Message);
        result := prov;
      end;
    end;
  end
  else
  begin
    prov := '%' + uppercase(province) + '%';
    try
      qValidProv.Params[0].Value := prov;
      qValidProv.Open;
      if (qValidProv.RecordCount > 0) and (qValidProv.FieldByName('ZONE_ID').AsString <> '') then
        result := qValidProv.FieldByName('ZONE_ID').AsString
      else
      begin
        prov := uppercase(province);
        SetLength(prov,4);
        result := prov;
      end;
      qValidProv.Close;
    except
      on E:Exception do
      begin
        LogMsg('[ERROR] - ValidateProvince('+province+') running query qValidProv: ' + E.Message);
        prov := uppercase(province);
        SetLength(prov,4);
        result := prov;
      end;
    end;
  end;
end;


function TDataMod.PrettyPhone(phone : string) : string;
var
  len : integer;

begin
  len := length(phone);
  if (len < 10) or (len > 15) then
    result := phone
  else
  begin
    var ext : boolean; if len > 10 then ext := true else ext := false;
    var prettyPhone : string;
    for var I := 1 to len do
    begin
      if isNumber(phone[I]) then
      begin
        prettyPhone := prettyPhone + phone[I];
        if (I = 3) or (I = 6) then
          prettyPhone := prettyPhone + '-'
        else if (I = 10) and ext then
          prettyPhone := prettyPhone + 'x';
      end
      else
      begin
        result := phone;
        exit;
      end;
    end;
    result := prettyPhone;
  end;
end;


function TDataMod.ValidateStatus_1(value,default : string) : TArray<string>;
// Trying to catch multiple values and anticipate what people might use to separate them
// Is there a comma, semicolon, slash
var
  I,nulls : integer;
  faxval : string;
  function IsNumeric(str : string) : boolean;
  begin
    result := true;
    var I : integer;
    for I := 1 to length(str) do
      if IsNumber(str[I]) = false then
        exit(false);
  end;
begin
  faxval := ''; nulls := 0;
  if containsStr(value,',') or containsStr(value,';') or containsStr(value,'/') or containsStr(value,'\') then
    result := value.Split([',',';','/','\'])
  else
    result := value.Split([' ']);

  if length(result) > 0 then
  begin
    for I := length(result) - 1 downto 0 do
    begin
      var val := trim(result[I]);
      if ContainsStr(val,'@') and ContainsStr(val,'.') then
      begin
        if (val <> '') and (length(val) > 6) then
        begin
          result[I] := 'E:' + val
        end
        else
        begin
          result[I] := '';
          inc(nulls);
        end;
      end
      else
      begin
        faxval := string.join('',val.Split(['-','(',')','.','+',' '])) + faxval;
        if length(faxval) > 9 then
        begin
          if IsNumeric(faxval) then
          begin
            result[I] := 'F:' + faxval;
          end
          else
          begin
            result[I] := '';
            inc(nulls);
          end;
          faxval := '';
        end
        else
        begin
          result[I] := '';
          inc(nulls);
        end;
      end;

    end;
    //for each check to see if they are legit
    if nulls = length(result) then
    begin
      SetLength(result,1);
      Result[0] := 'E:' + default;
    end;
  end
  else
    result := ['E:' + default];
end;


function TDataMod.RevBoolStr(TF : boolean) : string;
begin
  result := '';
  if TF then result := 'False' else result := 'True';
end;

function TDataMod.HandleBoolStr(boolStr : string) : boolean;
begin
  result := false;
  var temp := trim(uppercase(boolstr));
  if (temp = 'TRUE') or (temp = 'YES') or (temp = '1') then result := true
  else result := false;
end;

end.
