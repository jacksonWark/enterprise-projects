unit DMod;

interface

uses
  System.SysUtils, System.Classes, System.DateUtils, System.Generics.Collections, System.StrUtils, System.UITypes,
  Vcl.Dialogs,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.UI.Intf,
  FireDAC.Phys.Intf, FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async,
  FireDAC.Phys, FireDAC.Phys.DB2, FireDAC.Phys.DB2Def, FireDAC.VCLUI.Wait,
  FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf, FireDAC.DApt,
  FireDAC.Comp.Client, Data.DB, FireDAC.Comp.DataSet, Data.DBXDb2, Data.FMTBcd,
  Data.SqlExpr, ExeInfo, FireDAC.Phys.ODBCBase;

type
  TDataModule1 = class(TDataModule)
    DB2: TFDConnection;
    QFindRows: TFDQuery;
    procGEN_ID: TFDStoredProc;
    QUpdOld: TFDQuery;
    QInsertNew: TFDQuery;
    QFindInt: TFDQuery;
    QGetAveData: TFDQuery;
    QUpdAve: TFDQuery;
    QNewAveEntry: TFDQuery;
    QCalcAve: TFDQuery;
    QGetFCAAVE: TFDQuery;
    QGetAve: TFDQuery;
    QAveByDate: TFDQuery;
    QReplace: TFDQuery;
    QInsertClient: TFDQuery;
    SQLConnection1: TSQLConnection;
    spGEN_ID: TSQLStoredProc;
    ExeInfo: TExeInfo;
    QFindIntACD_ID: TIntegerField;
    QFindIntACD_RANGE_FROM: TFloatField;
    QFindIntACD_RANGE_TO: TFloatField;
    QFindIntACD_START_DATE: TSQLTimeStampField;
    QFindIntVENDOR_ID: TStringField;
    QFindRowsACD_ID: TIntegerField;
    QFindRowsACD_RANGE_FROM: TFloatField;
    QFindRowsACD_RANGE_TO: TFloatField;
    QFindRowsACD_START_DATE: TSQLTimeStampField;
    QFindRowsUSER_COND: TStringField;
    QFindRowsCLIENT_ID: TStringField;
    QUpdAve2: TFDQuery;
    qReplaceAveEntry: TFDQuery;
    qUpdAveReplace: TFDQuery;
    QGetFCAAVETX_DATE: TSQLTimeStampField;
    QGetFCAAVETL_RATE: TFloatField;
    QGetFCAAVEFCAAVE_LTL: TFloatField;
    QGetFCAAVEFCAAVE_TL: TFloatField;
    FDPhysDB2DriverLink1: TFDPhysDB2DriverLink;
    QGetFCAAVEltl_rate: TFloatField;
    QGetAveDataNew: TFDQuery;
    QGetAveDataNewTX_DATE: TSQLTimeStampField;
    QAveByDateTX_DATE: TSQLTimeStampField;
    QAveByDateLTL_RATE: TFloatField;
    QAveByDateTL_RATE: TFloatField;
    QGetAveDataNewLTL_RATE: TFloatField;
    QGetAveDataNewTL_RATE: TFloatField;
    QAveByDateIS_USED: TBooleanField;
    QFindClients: TFDQuery;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }

  public
    { Public declarations }
    function InitDB(alias,user,pwd,schema : string) : boolean;

    function GenID(mode : string) : integer;

    function TryGetAverage(currDate : TDateTime) : TArray<double>;

    procedure UpdateFCAAVEbatch(ltl,tl : string);

    function TryGetAverage2 : TArray<double>;
    function ExtractMonth(date : string) : integer;
    function MonthName(num : integer) : string;
    function IsZero(data : TArray<TArray<double>>) : boolean;
    function DoDialog(msg : string; dtype : integer) : boolean;
    function UpdateAveTable(vals : TArray<TArray<double>>; startDate,endDate : string; month : integer) : TArray<double>;

    procedure AddAveEntry(effDate : TDateTime; LTL,TL : double; used : boolean);
    function GetAverages : TArray<double>;
    function GetAveDate : TDateTime;
    function IsAveEntry(aDate : TDateTime) : boolean;

    //For replacing FCAAVE values
    procedure TryReplaceFCAAVE(day : TDateTime; ltl,tl : double);

    function GetFCAAVETable : TDataset;
    function GetDate(code : string; endDate : TDateTime) : TDateTime;
    function GetINTFCDate(code : string; endDate : TDateTime) : TDateTime;
    function VerifyClient(clientString, DBclient : string) : Boolean;
    function GenList(dataString : string) : TArray<TArray<String>>;

    procedure UpdateForCode(code : string; range : integer; LTL,TL,other : string; effDate, endDate : TDateTime; replace : boolean);
    procedure UpdateInterliner(vendor,rates : string; effDate, endDate : TDateTime; replace : boolean);

    procedure UpdateAndInsert(acdID : integer; rate : double; endDate, startDate : TDateTime;  vendorID, clientID : string);
    procedure ReplaceEntry(acdID : integer; rate : double; endDate : TDateTime);
  end;

var
  DataModule1: TDataModule1;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

function TDataModule1.InitDB(alias,user,pwd,schema :string) : boolean;
begin
  result := true;
  DB2.Params.Values['Alias'] := alias;
  DB2.Params.Values['Password'] := pwd;
  DB2.Params.Values['User_Name'] := user;

  try
    DB2.Connected := True;
    DB2.ExecSQL('SET CURRENT SCHEMA '+schema);
    DB2.ExecSQL('SET CURRENT PATH "SYSFUN","SYSPROC","SYSIBMADM",' + uppercase(schema));
  except
    exit(false)
  end;
end;


function TDataModule1.GenID(mode : string) : integer;
begin
  if mode = 'ACD' then procGEN_ID.ParamByName('IGEN_NAME').AsString := 'GEN_ACD_ID'
  else if mode = 'CLIENT' then procGEN_ID.ParamByName('IGEN_NAME').AsString := 'GEN_ACD_CLIENTS_ID';

  procGEN_ID.execute;
  result := procGEN_ID.ParamByName('ONEXT_ID').Value;

  procGEN_ID.ExecProc;
  result := procGEN_ID.ParamByName('ONEXT_ID').AsInteger;
end;


//Out of use as of Jul 28 2023
function TDataModule1.TryGetAverage(currDate : TDateTime) : TArray<double>;
begin
  result := nil;

  //get all entries in FCAAVE table that have not been used to calc an average yet
  QGetAveDataNew.Open;

  //as long as the result contains a column name it should have calculated an average
  //NOTE: this doesn't neccesarily mean it used the right dates
  if not QGetAveDataNew.FieldByName('TX_Date').IsNull then
  begin
    //calc ave pulls the average but does not update the table yet
    QCalcAve.Open;

    setlength(result,2);
    result[0] := QCalcAve.FieldByName('1').asFloat;

    QCalcAve.Next;

    result[1] := QCalcAve.FieldByName('1').asFloat;

    QCalcAve.Close;

    //This will update the table so the entries used to get this average will all have the averages and is_used set to true
    QUpdAve.ParamByName('LTLAVE').Value := result[0];
    QUpdAve.ParamByName('TLAVE').Value := result[1];
    QUpdAve.ExecSQL;

    //end;
  end;

  QGetAveDataNew.Close;
end;


procedure TDataModule1.UpdateFCAAVEbatch(ltl,tl : string);
var
  numLTL,numTL : double;
begin
  try
    QGetAveDataNew.Open;
  except on E:Exception do
    raise Exception.Create('Exception in "UpdateFCAAVEbatch" on "QGetAveDataNew.Open": ' + E.Message);
  end;
  if not QGetAveDataNew.FieldByName('TX_Date').IsNull then
  begin
    numLTL := strtofloat(replaceStr(ltl,'/',''));
    numTL := strtofloat(replaceStr(tl,'/',''));
    try
    QUpdAve.ParamByName('LTLAVE').Value := numLTL;
    QUpdAve.ParamByName('TLAVE').Value := numTL;
    QUpdAve.ExecSQL;
    except on E:Exception do
      raise Exception.Create('Exception in "UpdateFCAAVEbatch" on "QUpdAve.ExecSQL": ' + E.Message);
    end;
  end;
  QGetAveDataNew.Close;
end;


function TDataModule1.TryGetAverage2 : TArray<double>;
var
  numVals, month, I : integer;
  startDate,endDate : string;
  vals : TArray<TArray<double>>;
begin
  result := nil;
  SetLength(vals,1);
  SetLength(vals[0],2);

  //get all entries in FCAAVE table that have not been used to calc an average yet
  try
    QGetAveDataNew.Open;
  except on E:Exception do
    raise Exception.Create('Exception in "TryGetAverage2" on "QGetAveDataNew.Open": ' + E.Message);
  end;

  if not QGetAveDataNew.FieldByName('TX_Date').IsNull then
  begin
    endDate := QGetAveDataNew.FieldByName('TX_Date').AsString;
    month := ExtractMonth(endDate);
    vals[0][0] := QGetAveDataNew.FieldByName('LTL_RATE').AsFloat;
    vals[0][1] := QGetAveDataNew.FieldByName('TL_RATE').AsFloat;
    numVals := 1;


    if QGetAveDataNew.RecordCount > 1 then
    begin
      for I := 2 to QGetAveDataNew.RecordCount do
      begin
        QGetAveDataNew.Next;

        if ExtractMonth(QGetAveDataNew.FieldByName('TX_Date').AsString) = month then
        begin
          setLength(vals, length(vals) + 1);
          setLength(vals[numVals], 2);
          vals[numVals][0] := QGetAveDataNew.FieldByName('LTL_RATE').AsFloat;
          vals[numVals][1] := QGetAveDataNew.FieldByName('TL_RATE').AsFloat;
          inc(numVals);

          if I = QGetAveDataNew.RecordCount then
          begin
            startDate := QGetAveDataNew.FieldByName('TX_Date').AsString;
            if result = nil then
              result := UpdateAveTable(vals,startDate,endDate,month)
            else
              UpdateAveTable(vals,startDate,endDate,month);
          end;
        end
        else
        begin
          QGetAveDataNew.Prior;
          startDate := QGetAveDataNew.FieldByName('TX_Date').AsString;
          QGetAveDataNew.Next;

          if result = nil then
            result := UpdateAveTable(vals,startDate,endDate,month)
          else
            UpdateAveTable(vals,startDate,endDate,month);

          month := ExtractMonth(QGetAveDataNew.FieldByName('TX_Date').AsString);
          endDate := QGetAveDataNew.FieldByName('TX_Date').AsString;
          setLength(vals, 1);
          setLength(vals[0], 2);
          vals[0][0] := QGetAveDataNew.FieldByName('LTL_RATE').AsFloat;
          vals[0][1] := QGetAveDataNew.FieldByName('TL_RATE').AsFloat;
          numVals := 1;
        end;

      end;
    end
    else
      raise Exception.Create('FCAAVE was not calculated for '+MonthName(month)+' as there are fewer than 4 values for the month.')
  end;
  QGetAveDataNew.Close;
end;


function TDataModule1.UpdateAveTable(vals : TArray<TArray<double>>; startDate,endDate : string; month : integer) : TArray<double>;
var
  ltlTotal,tlTotal,ltlAve,tlAve : double;
  I,len : integer;
begin
  result := nil;
  len := length(vals);
  if len < 4 then
    raise Exception.Create('FCAAVE was not calculated for '+MonthName(month)+' as there are fewer than 4 values for the month.')
  else if len > 5 then
    raise Exception.Create('FCAAVE was not calculated for '+MonthName(month)+' as there are more than 5 values for the month.')
  else
  begin
    if IsZero(vals) then
      raise Exception.Create('FCAAVE was not calculated for '+MonthName(month)+' as one of the values in the table is a zero.')
    else
    begin
      //do stuff here
      for I := 0 to len - 1 do
      begin
        ltlTotal := ltlTotal + vals[I][0];
        tlTotal := tlTotal + vals[I][1];
      end;

      ltlAve := ltlTotal/len;
      tlAve := tlTotal/len;

      SetLength(result, 2);
      result[0] := ltlAve;
      result[1] := tlAve;

      try
        QupdAve2.ParamByName('LTLAVE').Value := ltlAve;
        QupdAve2.ParamByName('TLAVE').Value := tlAve;
        QupdAve2.ParamByName('START_DATE').Value := startDate;
        QupdAve2.ParamByName('END_DATE').Value := endDate;
        QUpdAve2.ExecSQL;
      except on E:Exception do
        raise Exception.Create('Exception in "UpdateAveTable" on "QUpdAve2.ExecSQL": ' + E.Message)
      end;
    end;
  end;
end;


function TDataModule1.ExtractMonth(date : string) : integer;
begin
  result := StrtoInt(date.Split(['/'])[0]);
end;


function TDataModule1.MonthName(num : integer) : string;
var
  months : TArray<string>;
begin
  months := ['January','February','March','April','May','June','July','August','September','October','November','December'];
  result := months[num];
end;


function TDataModule1.IsZero(data : TArray<TArray<double>>) : boolean;
var
  I,J : integer;
begin
  result := false;
  for I := 0 to length(data) - 1 do
  begin
    for J := 0 to length(data[I]) - 1 do
      if data[I][J] = 0 then
      begin
        result := true;
        exit
      end;
  end;
end;


//dtype: 0 - information, 1 - confirmation, 2 - error
function TDataModule1.DoDialog(msg : string; dtype : integer) : boolean;
var
  modRes : integer;
  ddType : TMsgDlgType;
  buttons : TMsgDlgButtons;
begin
  if dtype = 0 then
  begin
    ddType := mtInformation;
    buttons := [mbOk];
  end
  else if dtype = 1 then
  begin
    ddType := mtConfirmation;
    buttons := [mbYes,mbNo];
  end
  else if dtype = 2 then
  begin
    ddType := mtError;
    buttons := [mbOk];
  end;

  modRes := MessageDlg(msg, ddType,buttons,0);

  if dtype = 1 then
    if modRes = mrYes then result := true
    else result := false
  else
    if modRes = mrOK then result := true
    else result := false;

end;


procedure TDataModule1.AddAveEntry(effDate : TDateTime; LTL,TL : double; used : boolean);
begin
  try
    QAveByDate.ParamByName('inDate').AsDateTime := DateOf(effDate);
    QAveByDate.Open;
  except on E:Exception do
    raise Exception.Create('Exception in "AddAveEntry" on "QAveByDate.Open": ' + E.Message)
  end;

  if QAveByDate.FieldByName('TX_DATE').IsNull then
  begin
    try
      QNewAveEntry.ParamByName('DATE').Value := effDate;
      QNewAveEntry.ParamByName('LTL').Value := LTL;
      QNewAveEntry.ParamByName('TL').Value := TL;
      QNewAveEntry.ParamByName('USED').Value := used;
      QNewAveEntry.ExecSQL;
    except on E:Exception do
      raise Exception.Create('Exception in "AddAveEntry" on "QNewAveEntry.ExecSQL": ' + E.Message)
    end;
  end;

  QAveByDate.Close;
end;


function TDataModule1.GetAverages : TArray<double>;
begin
  result := nil;
  try
    QGetAve.Open;         //qCalcAve
    setlength(result,2);
    if QGetAve.FieldByName('FCAAVE_LTL').IsNull then
    begin
      result[0] := 0;
      result[1] := 0;
    end
    else
    begin
      result[0] := QGetAve.FieldByName('FCAAVE_LTL').Value;
      result[1] := QGetAve.FieldByName('FCAAVE_TL').Value;
    end;
    QGetAve.Close;
  except on E:Exception do
    raise Exception.Create('Exception in "GetAverages": ' + E.Message)
  end;
end;


Function TDataModule1.GetAveDate : TDateTime;
begin
  try
    QGetAve.open;
  except on E:Exception do
    raise Exception.Create('Exception in "GetAveDate" on "QGetAve.open": ' + E.Message)
  end;
  result := QGetAve.FieldByName('TX_Date').Value;

  QGetAve.Close;
end;


//Currently not in use
function TDataModule1.IsAveEntry(aDate : TDateTime) : boolean;
begin
  QAveByDate.ParamByName('inDate').Value := aDate;
  QAveByDate.Open;

  if QAveByDate.FieldByName('TX_DATE').IsNull then
    result := false
  else
    result := true;

end;


procedure TDataModule1.TryReplaceFCAAVE(day : TDateTime; ltl, tl : double);
var
  month, year : integer;
  ltlave, tlave : double;
  startDate, endDate : TDateTime;
  isused : boolean;
begin
  try
    qAveByDate.Params[0].value := DateOf(day);
    qAveByDate.Open;
  except on E:Exception do
    raise Exception.Create('Exception in "TryReplaceFCAAVE" on "qAveByDate.Open": ' + E.Message)
  end;

  if qAveByDate.RecordCount > 0 then
  begin
    if (qAveByDate.FieldByName('LTL_RATE').AsFloat <> ltl) and (qAveByDate.FieldByName('TL_RATE').AsFloat <> tl) then
    begin
      try
        qReplaceAveEntry.ParamByName('LTL').value := ltl;
        qReplaceAveEntry.ParamByName('TL').value := tl;
        qReplaceAveEntry.ParamByName('INDATE').Value := qAveByDate.FieldByName('TX_DATE').Value;
        qReplaceAveEntry.ExecSQL;
      except on E:Exception do
        raise Exception.Create('Exception in "TryReplaceFCAAVE" on "qReplaceAveEntry.ExecSQL": ' + E.Message)
      end;

      isused := qAveByDate.FieldByName('IS_USED').AsBoolean;
      if isused then
      begin
        //update the average for this month
        month := monthOf(day);
        year := yearOF(day);

        startDate := StartOfAMonth(year,month);
        endDate := EndOfAMonth(year,month);

        //Calc averages
        try
          qCalcAve.ParamByName('START').Value := startDate;
          qCalcAve.ParamByName('END').Value := endDate;
          qCalcAve.Open;
        except on E:Exception do
          raise Exception.Create('Exception in "TryReplaceFCAAVE" on "qCalcAve.Open": ' + E.Message)
        end;
        if qCalcAve.RecordCount >= 2 then
        begin
          ltlave := qCalcAve.Fields[0].AsFloat;
          qCalcAve.Next;
          tlave := qCalcAve.Fields[0].AsFloat;

          //update averages
          try
            qUpdAveReplace.ParamByName('START_DATE').Value := startDate;
            qUpdAveReplace.ParamByName('END_DATE').Value := endDate;
            qUpdAveReplace.ParamByName('LTLAVE').Value := ltlave;
            qUpdAveReplace.ParamByName('TLAVE').Value := tlave;
            qUpdAveReplace.ExecSQL;
          except on E:Exception do
            raise Exception.Create('Exception in "TryReplaceFCAAVE" on "qUpdAveReplace.ExecSQL": ' + E.Message)
          end;
        end;
        qCalcAve.Close;
      end;
    end;
  end;

  qAveByDate.Close;

end;


function TDataModule1.GetFCAAVETable : TDataset;
begin
  result := nil;
  try
    QGetFCAAVE.Open;
    result := QGetFCAAVE.GetClonedDataSet(false);
    QGetFCAAVE.Close;
  except on E: Exception do
    raise Exception.Create('Exception in "GetFCAAVETable" on "QGetFCAAVE.Open": '+E.Message);
  end;
end;


function TDataModule1.GetDate(code : string; endDate : TDateTime) : TDateTime;
begin
  result := encodeDate(1899,12,31);
  try
    QFindRows.ParamByName('END_DATE').AsDateTime := endDate - 1;
    QFindRows.ParamByName('ACODE_ID').Value := code;
    QFindRows.Open;
  except on E:Exception do
    raise Exception.Create('Exception in "GetDate" on "QFindRows.Open": ' + E.Message)
  end;

  if not QFindRows.FieldByName('ACD_START_DATE').IsNull then
    result := QFindRows.FieldByName('ACD_START_DATE').value;

  QFindRows.Close;
end;


function TDataModule1.GetINTFCDate(code : string; endDate : TDateTime) : TDateTime;
begin
  result := encodeDate(1899,12,31);
  try
    QFindInt.ParamByName('END_DATE').AsDateTime := endDate - 1;
    QFindInt.ParamByName('VENDOR_ID').Value := code;
    QFindInt.Open;
  except on E:Exception do
    raise Exception.Create('Exception in "GetINTFCDate" on "QFindInt.Open": ' + E.Message)
  end;

  if not QFindInt.FieldByName('ACD_START_DATE').IsNull then
    result := QFindInt.FieldByName('ACD_START_DATE').value;

  QFindInt.Close;
end;


//Check the current rate versus the current record from the DB we are looking at to make sure the client is right
function TDataModule1.VerifyClient(clientString, DBclient : string) : Boolean;
begin

  if ContainsText(clientString, '...') then result := '...' = DBclient
  else result := clientString = DBclient;

end;

//Take the input strings specifying rates and clients and turn them into 2D arrays
function TDataModule1.GenList(dataString : string) : TArray<TArray<String>>;
var
  commaList : TArray<String>;
  I : integer;
begin
  if ContainsText(dataString, ',') then
  begin
    commaList := dataString.Split([',']);
    SetLength(result,length(commaList));
    for I := 0 to (length(commaList) - 1) do
    begin
      result[I] := commaList[I].Split(['/']);
    end;
  end
  else
  begin
    SetLength(result,1);
    result[0] := dataString.Split(['/']);
  end;
end;

//Params:
//  code ----> A string representing the ACODE_ID value from the ACHARGE_DETAIL table
//  range ---> Integer representing the break point in the weight range where LTL ends and TL starts. Doesn't apply to 'other' type rates
//  Rates:  String containing one or more rates that can be associated with a CLIENT_ID, separated by commas
//      ex: /66.3,COMLOG/33.15    1st rate no CLIENT_ID and rate=66.3, second rate CLIENT_ID=COMLOG and rate=33.15
//    LTL -----> applies to the weight range 0-range
//    TL ------> applies to the weight range range-1000000
//    other ---> applies to the weight range 0-1000000
//  effDate -> DateTime value to be applied to the ACD_START_DATE field
//  endDate -> DateTime value representing the furthest point in the future. A code that has this value as its ACD_END_DATE is considered an active rate.
//             NOTE: This value should be a constant and does not need to be passed as a parameter
//  replace -> A boolean value that determines whether to create a new entry or update an old one. True updates an existing record.
//             False makes the existing record inactive, copies it with the new rates and makes it active.
procedure TDataModule1.UpdateForCode(code : string; range : integer; LTL,TL,other : string; effDate, endDate : TDateTime; replace : boolean);
var
  I,usrCondWgt : integer;
  clientID,usrCond,usrCondSign : string;
  doUpdate : boolean;
  ltlList, tlList, otherList : TArray<TArray<string>>;
  rate : double;
  dType : TFieldType;
begin
  try
    QFindRows.ParamByName('END_DATE').AsDateTime := endDate - 1;
    QFindRows.ParamByName('ACODE_ID').AsString := code;
    QFindRows.Open;
  except on E:Exception do
    raise Exception.Create('Exception in "UpdateForCode" on "QFindRows.Open": ' + E.Message)
  end;

  doUpdate := false;

  //Handle data from the strings containing one or more 'LTL', 'TL', or 'other' type rates
  ltlList := GenList(LTL);
  tlList := GenList(TL);
  otherList := GenList(other);

  while not QFindRows.Eof do
  begin
    //New on 2022 Apr 20: We need to check if range from and range to are both zero so we know we are looking at an entry that is 'valuation' type
    //Then look for a 'USER_COND' and read the contents to determine which rate to apply to which. For now gonna be specific to the 'INTSURFSC' ACODE_ID
    //BUT WE MAY NEED TO ADDRESS THIS ISSUE EVENTUALLY
    if (QFindRows.FieldByName('ACD_RANGE_FROM').Value = 0) and (QFIndRows.FieldByName('ACD_RANGE_TO').Value = 0) then
    begin
      //dType := QFindRows.FieldByName('USER_COND').DataType;
      usrCond := ''; usrCondSign := ''; usrCondWgt := 0;
      usrCond := QFindRows.FieldByName('USER_COND').Value;
      //usrCondWgt := strtoint(ReplaceStr(ReplaceStr(ReplaceStr(usrCond.Trim,'WEIGHT',''),'<',''),'>=','').Trim);
      usrCondSign := ReplaceStr(ReplaceStr(usrCond.Trim,'WEIGHT',''),'10000','').Trim;

      if usrCondSign = '<' then
      begin
        rate := strtofloat(ltlList[0][1]);
        doUpdate := true;
      end
      else if usrCondSign = '>=' then
      begin
        rate := strtofloat(tlList[0][1]);
        doUpdate := true;
      end;
    end
    else
    begin //determine whether the value we need to set is LTL, TL or other and set it
      if (QFindRows.FieldByName('ACD_RANGE_FROM').Value < 3 ) then
      begin
        if (QFindRows.FieldByName('ACD_RANGE_TO').Value > range ) then
        begin
          for I := 0 to (length(otherList) - 1) do
          begin
            if VerifyClient(otherList[I][0],QFindRows.FieldByName('CLIENT_ID').Value) then
            begin
              rate := strtofloat(otherList[I][1]);
              clientID := otherList[I][0];
              doUpdate := true;
              break;
            end;
          end;
        end
        else
        begin
          for I := 0 to (length(ltlList) - 1) do
          begin
            if VerifyClient(ltlList[I][0],QFindRows.FieldByName('CLIENT_ID').Value) then
            begin
              rate := strtofloat(ltlList[I][1]);
              clientID := ltlList[I][0];
              doUpdate := true;
              break;
            end;
          end;
        end;
      end
      else
      begin
        for I := 0 to (length(tlList) - 1) do
        begin
          if VerifyClient(tlList[I][0],QFindRows.FieldByName('CLIENT_ID').Value) then
          begin
            rate := strtofloat(tlList[I][1]);
            clientID := tlList[I][0];
            doUpdate := true;
            break;
          end;
        end;
      end;
    end;

    if doUpdate and (rate > 0) then
    begin
      if replace then
      begin
        ReplaceEntry(QFindRows.FieldByName('ACD_ID').Value, rate, endDate);
      end
      else
      begin
        UpdateAndInsert(QFindRows.FieldByName('ACD_ID').Value,rate,endDate,effDate,'',clientID);
      end;
      if code = 'FCAAVE' then
      begin
        //If we wanna add back the ability to update FCAAVE in the batch update with historical data, we need to update the FCAAVE table from here.
        //See paper notes
        //But can we use the old method to bluntly update all unused entries in the FCAAVE table with just LTL and TL values?
        //Ok we are going with the last one for now
        //UpdateFCAAVEbatch(strtofloat(ltlList[I][1]),TL);
      end;
      doUpdate := false; rate := 0;
    end;

    QFindRows.Next;
  end;

  QFindRows.Close;
end;


procedure TDataModule1.UpdateInterliner(vendor, rates : string; effDate, endDate : TDateTime; replace : boolean);
var
  rateList : TArray<string>;
  rate : double;
begin
  try
    QFindInt.ParamByName('VENDOR_ID').Value := vendor;
    QFindInt.ParamByName('END_DATE').Value := endDate - 1;
    QFindInt.Open;
  except on E:Exception do
    raise Exception.Create('Exception in "UpdateInterliner" on "QFindInt.Open": ' + E.Message)
  end;

  while not QFindInt.Eof do
  begin
    rateList := rates.Split([':'], 2);

    //determine whether the value we need to set is LTL, TL or other and set it
    if (QFindInt.FieldByName('ACD_RANGE_FROM').Value < 3 ) then
    begin
        rate := strToFloat(rateList[0]);
    end
    else
    begin
      rate := strToFloat(rateList[1]);
    end;

    if (rate > 0) then
    begin
      if replace then
      begin
        ReplaceEntry(QFindInt.FieldByName('ACD_ID').Value, rate, endDate);
      end
      else
      begin
        UpdateAndInsert(QFindInt.FieldByName('ACD_ID').Value,rate,endDate,effDate,vendor,'');
      end;
    end;

    QFindInt.Next;
  end;

  QFindInt.Close;
end;



procedure TDataModule1.ReplaceEntry(acdID : integer; rate : double; endDate : TDateTime);
begin
  try
    QReplace.ParamByName('ACD_ID').Value := acdID;
    QReplace.ParamByName('RATE').Value := rate;
    QReplace.ParamByName('END_DATE').Value := endDate;
    QReplace.ExecSQL;
  except on E:Exception do
    raise Exception.Create('Exception in "ReplaceEntry" on "QReplace.ExecSQL": ' + E.Message)
  end;
end;



procedure TDataModule1.UpdateAndInsert(acdID : integer; rate : double; endDate, startDate : TDateTime; vendorID, clientID : string);
var
  acdIDNew, I : integer;
  clientString : String;
  clientList : TArray<String>;
begin
  //update old entry with end date just before our new start date
  try
    QUpdOld.ParamByName('ACD_ID').AsInteger := acdID;
    QUpdOld.ParamByName('ACD_END_DATE').AsDateTime := EndOfTheDay(startDate - 1);
    QUpdOld.ExecSQL;
  except on E:Exception do
    raise Exception.Create('Exception in "UpdateAndInsert" on "QUpdOld.ExecSQL": ' + E.Message)
  end;

  acdIDNew := GenID('ACD');

  //create new entry
  try
    QInsertNew.ParamByName('ACD_ID').AsInteger := acdIDNew;
    QInsertNew.ParamByName('ACD_ID2').AsInteger := acdID;
    QInsertNew.ParamByName('ACD_START_DATE').AsDateTime := startDate;
    QInsertNew.ParamByName('ACD_END_DATE').AsDateTime := endDate;
    QInsertNew.ParamByName('ACD_PERCENT').Value := rate;
    QInsertNew.ParamByName('VENDOR_ID').Value := vendorID;
    QInsertNew.ExecSQL;
  except on E:Exception do
    raise Exception.Create('Exception in "UpdateAndInsert" on "QInsertNew.ExecSQL": ' + E.Message)
  end;

  //If client = '' do nothing, if = '...' then do
  if clientID <> '' then
  begin
    if ContainsText(clientID,'...') then
    begin
      clientString := StringReplace(clientID,'...\','',[rfReplaceAll]);
      clientList := clientString.Split(['-']);
      for I := 0 to (length(clientList) - 1) do
        begin
          try
            QInsertClient.ParamByName('ACD_ID').AsInteger := acdIDNew;
            QInsertClient.ParamByName('ID').AsInteger := GenID('CLIENT');
            QInsertClient.ParamByName('CLIENT_ID').AsString := clientList[I];
            QInsertClient.ExecSQL;
          except on E:Exception do
            raise Exception.Create('Exception in "UpdateAndInsert" on "QInsertClient.ExecSQL": ' + E.Message)
          end;
        end;
    end
    else
    begin
      try
        QInsertClient.ParamByName('ACD_ID').AsInteger := acdIDNew;
        QInsertClient.ParamByName('ID').AsInteger := GenID('CLIENT');
        QInsertClient.ParamByName('CLIENT_ID').AsString := clientID;
        QInsertClient.ExecSQL;
      except on E:Exception do
        raise Exception.Create('Exception in "UpdateAndInsert" on "QInsertClient.ExecSQL": ' + E.Message);
      end;
    end;
  end;

end;

end.
