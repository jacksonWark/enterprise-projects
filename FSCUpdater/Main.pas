unit Main;

{ NOTES

    Pending improvements:
    1. Current system for doing FCAAVE doesn't cover some issues.
      - We can calculate past averages indefinitely but there is no way to use any but the most recent.
      - Maybe there should be a selection feature to choose which month+year you want to use.
      - Also should there be the ability to edit the FCAAVE table? (maybe not!)
        Maybe just delete an entry.
      - Currently the only way to add an FCAAVE entry is by doing an update for FCA rates
          This is a good thing because users cannot accidentally pollute our table
          This is a bad thing because users cannot recover from errors on their own
}

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes, System.IOUtils, System.RegularExpressions, System.IniFiles, System.DateUtils,
  System.Types, System.Math, System.UITypes, System.Generics.Collections, System.StrUtils,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.Imaging.pngimage, Vcl.StdCtrls,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdHTTP, IdIOHandler, IdIOHandlerSocket, IdIOHandlerStack,
  IdSSL, IdSSLOpenSSL, IdStack,
  Vcl.ComCtrls, Vcl.Mask, Data.DB, Vcl.Grids, Vcl.DBGrids,
  IdExplicitTLSClientServerBase, IdMessageClient, IdSMTPBase, IdSMTP,
  sharedServices, crypto, DBConfig, TaskDialog, TaskDialogEx;

type
  TForm1 = class(TForm)
  {$REGION 'Auto Generated'}
    bUpdate: TButton;
    iRefresh: TImage;
    pRefresh: TPanel;
    pStatusPanel: TPanel;
    lEffDate: TLabel;
    iRefreshPress: TImage;
    iRefreshHover: TImage;
    HTTPclient: TIdHTTP;
    IdSSLIOHandlerSocketOpenSSL1: TIdSSLIOHandlerSocketOpenSSL;
    pFCA: TPanel;
    pFuel: TPanel;
    pWeeklyEmail: TPanel;
    pMonthlyEmail: TPanel;
    EffDatePicker: TDateTimePicker;
    PageControl1: TPageControl;
    Main: TTabSheet;
    FCAAVETableTab: TTabSheet;
    lTitle: TLabel;
    FCAAVETable: TDBGrid;
    DataSource1: TDataSource;
    pAve: TPanel;
    pFSURCHG: TPanel;
    lFUELSURCHGTL: TLabel;
    lFUELSURCHG: TLabel;
    lFUELSURCHGLTL: TLabel;
    pAgentFSC: TPanel;
    meAGENTFSC: TMaskEdit;
    Memo1: TMemo;
    meFSC: TMaskEdit;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    pComOve: TPanel;
    meCO: TMaskEdit;
    pCOtitle: TPanel;
    lComoxOverland: TLabel;
    cbCO: TCheckBox;
    pCOlbl: TPanel;
    lCOLTL: TLabel;
    lCOTL: TLabel;
    pFounTire: TPanel;
    Panel3: TPanel;
    lFountainTire: TLabel;
    cbFT: TCheckBox;
    pFounTireLbl: TPanel;
    lFTLTL: TLabel;
    lFTTL: TLabel;
    meFT: TMaskEdit;
    pQuikX: TPanel;
    pQXtitle: TPanel;
    lQuikXTransport: TLabel;
    cbQX: TCheckBox;
    pQXlbl: TPanel;
    lQXLTL: TLabel;
    lQXTL: TLabel;
    meQX: TMaskEdit;
    pCdnDom: TPanel;
    Label6: TLabel;
    cbCdnDom: TCheckBox;
    Panel1: TPanel;
    Label9: TLabel;
    lCdnDomTL: TLabel;
    lCdnDomLTL: TLabel;
    Label8: TLabel;
    lCdnDomEffDate: TLabel;
    Label14: TLabel;
    pMapei: TPanel;
    pMapeiTitle: TPanel;
    pMapeiLbl: TPanel;
    meMap: TMaskEdit;
    lMapei: TLabel;
    cbMapei: TCheckBox;
    lMapLTL: TLabel;
    lMapTL: TLabel;
    pNFastFreight: TPanel;
    pNFastFreightTitle: TPanel;
    pNFastFreightLbl: TPanel;
    Label10: TLabel;
    Label11: TLabel;
    lNFastFreight: TLabel;
    meNFF: TMaskEdit;
    cbNFF: TCheckBox;
    pAlcanRSGILE: TPanel;
    pAlcanRSGILEtitle: TPanel;
    lRTrsgile: TLabel;
    cbRTRS: TCheckBox;
    pAlcanRSGILELbl: TPanel;
    lRTLTL: TLabel;
    lRTTL: TLabel;
    lRTother: TLabel;
    meRT: TMaskEdit;
    pBrewers: TPanel;
    pBrewersTitle: TPanel;
    lBrewers: TLabel;
    cbBrew: TCheckBox;
    pBrewersLbl: TPanel;
    Label13: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    meBrew: TMaskEdit;
    pManiMantra: TPanel;
    pManiMantraTitle: TPanel;
    pManiMantraLbl: TPanel;
    Label7: TLabel;
    Label19: TLabel;
    meMM: TMaskEdit;
    lManitoulinMantra: TLabel;
    pNatTire: TPanel;
    pNatTireLbl: TPanel;
    meNT: TMaskEdit;
    lNationalTire: TLabel;
    cbNatTire: TCheckBox;
    pEncorp: TPanel;
    pEncorpLbl: TPanel;
    meEn: TMaskEdit;
    lEncorp: TLabel;
    cbEncorp: TCheckBox;
    pWFraser: TPanel;
    pWFraserTitle: TPanel;
    pWFraserLbl: TPanel;
    Label21: TLabel;
    lWestFraser: TLabel;
    cbWFra: TCheckBox;
    meWFra: TMaskEdit;
    pAPPCAR: TPanel;
    Label20: TLabel;
    cbAPPCAR: TCheckBox;
    Label25: TLabel;
    Label28: TLabel;
    lFCAave: TLabel;
    pFcaaveTop: TPanel;
    bReCalcFCAAVE: TButton;
    bRefreshFCAAVE: TButton;
    Panel12: TPanel;
    pFCAAVEmain: TPanel;
    Panel13: TPanel;
    Panel14: TPanel;
    Label22: TLabel;
    Panel15: TPanel;
    Panel18: TPanel;
    p18bot: TPanel;
    p18top: TPanel;
    Label23: TLabel;
    Label24: TLabel;
    Label26: TLabel;
    cbFCAAVE: TCheckBox;
    Panel16: TPanel;
    laveMonth: TLabel;
    Label27: TLabel;
    Label29: TLabel;
    laveLTLdata: TLabel;
    laveTLdata: TLabel;
    Label30: TLabel;
    bImpHEChart: TButton;
    pSelDB: TPanel;
    bSelDB: TButton;
    bChgDB: TButton;
    bImpNEWREDChart: TButton;
    ScrollBox2: TScrollBox;
    pUpd: TPanel;
    pFCAmain: TPanel;
    pFCAtitle: TPanel;
    pFCAltl: TPanel;
    pFCAtl: TPanel;
    lFCA: TLabel;
    lFCA2: TLabel;
    lFCA4: TLabel;
    lFCA4P: TLabel;
    lFCA90: TLabel;
    lFUAP: TLabel;
    l2LTL: TLabel;
    l4LTL: TLabel;
    l4PtLTL: TLabel;
    l90LTL: TLabel;
    lFUAPLTL: TLabel;
    lLTL: TLabel;
    meLTL: TMaskEdit;
    l2TL: TLabel;
    l4PtTL: TLabel;
    l4TL: TLabel;
    l90TL: TLabel;
    lFUAPTL: TLabel;
    lTL: TLabel;
    meTL: TMaskEdit;
    Panel17: TPanel;
    Panel19: TPanel;
    lFUELENDA: TLabel;
    lFuelLTL: TLabel;
    lFuelTL: TLabel;
    Panel10: TPanel;
    Panel28: TPanel;
    Panel29: TPanel;
    Panel30: TPanel;
    Panel35: TPanel;
    Panel36: TPanel;
    Label37: TLabel;
    Panel37: TPanel;
    Label38: TLabel;
    Label35: TLabel;
    Label36: TLabel;
    cbFCA: TCheckBox;
    Panel25: TPanel;
    Panel27: TPanel;
    Label39: TLabel;
    Panel33: TPanel;
    Label40: TLabel;
    cbFSC: TCheckBox;
    lFUELSURCHGInfo: TLabel;
    Panel2: TPanel;
    Label33: TLabel;
    cbFuelenda: TCheckBox;
    Panel11: TPanel;
    Panel31: TPanel;
    Label34: TLabel;
    Panel32: TPanel;
    Label42: TLabel;
    cbAgentFSC: TCheckBox;
    AgentPicker: TDateTimePicker;
    Panel38: TPanel;
    Panel39: TPanel;
    Panel40: TPanel;
    Panel41: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    lAgentFSC: TLabel;
    lAGENT30data: TLabel;
    lAGENT40data: TLabel;
    lCLETRUdata: TLabel;
    cbMM: TCheckBox;
    pMonthlyTitle: TPanel;
    Label31: TLabel;
    Panel22: TPanel;
    Label43: TLabel;
    pWeeklyTitle: TPanel;
    Label17: TLabel;
    Panel8: TPanel;
    lWeeklyEmailTitle: TLabel;
    Label12: TLabel;
    pFUELENDA: TPanel;
    pPRGFuel: TPanel;
    Label18: TLabel;
    Label41: TLabel;
    Label32: TLabel;
    lFuelLTLData: TLabel;
    lFuelTLData: TLabel;
    pNEWRED: TPanel;
    Panel5: TPanel;
    lNEWRED: TLabel;
    Panel6: TPanel;
    Label45: TLabel;
    lnewredLTL: TLabel;
    Label47: TLabel;
    lnewredTL: TLabel;
    Panel7: TPanel;
    Label49: TLabel;
    cbNEWRED: TCheckBox;
    Label44: TLabel;
    lWebRateData: TLabel;
    Label46: TLabel;
    Label48: TLabel;
    Label50: TLabel;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    lWebNEWRED: TLabel;
    AdvTaskDialog1: TAdvTaskDialog;
    AdvTaskDialogEx1: TAdvTaskDialogEx;
    AdvInputTaskDialogEx1: TAdvInputTaskDialogEx;
    procedure iRefreshMouseEnter(Sender: TObject);
    procedure iRefreshHoverMouseLeave(Sender: TObject);
    procedure iRefreshHoverMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure iRefreshHoverMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FormActivate(Sender: TObject);
    procedure eLTLKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    {procedure bEmailRatesClick(Sender: TObject);}
    procedure DateTimePicker1Change(Sender: TObject);
    procedure iRefreshHoverClick(Sender: TObject);
    procedure bUpdateClick(Sender: TObject);
    procedure meCOKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure meQXKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure meFTKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure meRTKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure meBrewKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure meGrimKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure meMapKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure meMMKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure meWFraKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure meNFFKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure meEnKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure meNTKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure cbFCAAVEClick(Sender: TObject);
    procedure meAGENTFSCKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure meFSCKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure meLTLChange(Sender: TObject);
    procedure bRefreshFCAAVEClick(Sender: TObject);
    procedure bReCalcFCAAVEClick(Sender: TObject);
    procedure bImpHEChartClick(Sender: TObject);
    procedure bChgDBClick(Sender: TObject);
    procedure bImpNEWREDChartClick(Sender: TObject);
    procedure ScrollBox2MouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
    procedure DateTimePicker2Change(Sender: TObject);
    procedure lFUAPDblClick(Sender: TObject);

  {$ENDREGION}
  {$REGION 'private'}
  private
    { Private declarations }
    function Scrape(myURL : String) : String;
    function RegEx(BuffString, Mode : String) : String;
    function GetClosestMonday : TDateTime;
    function GetClosestWeekday(weekday : integer) : TDateTime;
    function NearestNorCanDate : TDateTime;
    function ConvertDate(date : string) : string;
    function LastMonday(date : TDateTime) : Boolean;

    procedure CheckFCARates;
    function RefineRatesHTML(myURL: String) : TArray<String>;
    procedure CalcRates(ltl, tl : string);

    procedure MakeHEChart;
    procedure LoadHEChart;
    procedure LoadNEWREDChart;
    procedure CheckNorCanRate(onStartup : boolean);
    procedure CheckNEWREDRate(onStartup : boolean);

    procedure CheckAllRates;
    procedure CheckAve;
    procedure cbCalcAve;

    function ApplySavedConfig : integer;
    procedure DBSetup;
    procedure DBPopup;
    procedure StartupDBReq;

    procedure UpdFCAAVETable;
    procedure NewWeek;

    procedure PrepareUpdateList;
    procedure ClearUpdateList;
    procedure AddToList(header, data : string);

    function ChkdRatesStr : string;

    procedure PrepareINTDict;
    procedure ClearINTDict;
    function InputValid(val : string) : Boolean;

    procedure ErrMsg(msg : string);

    function RegExCdnDom : string;
    function ExtractValCdnDom(regExMatch : string; rate : boolean = false) : string;
    procedure UpdCdnDom;

    procedure UpdateFCAAVETable;
  {$ENDREGION}
  {$REGION 'public'}
  public
    { Public declarations }
    Debug : Boolean;
    resized : Boolean;
    terminate : Boolean;

    Database : string;
    dbConnArr : TArray<string>;

    dateShift : Boolean;
    error : Boolean;
    lastDate : TDateTime;
    NorCanDate, NEWREDDate : TDateTime;
    dateSettings : TFormatSettings;
    //HEChart : TDictionary<double,string>;
    HEChart, NEWREDChart : TDictionary<integer,string>;
    HEMax, HEMin, NEWREDMax, NEWREDMin : integer;
    HEloaded, NEWREDloaded : boolean;
    INTFC : TArray<string>;
    FCArates : TArray<string>;
    unshiftedDate : TDateTime;
    NorCanHtml,NEWREDHtml : string;
    FCAAVES : TArray<double>;

    urlNorCanCfg, NorCanPatternCfg,
    urlCdnDomCfg, CdnDomPatternCfg : String;   

    JudgementDay : TDateTime;

    const regFormSize = 580;
    const expFormSize = 789;
    const updButtonOffset = 91;
    const urlFSC = 'https://yrc.com/fuel-surcharge-canada/';
    const urlNorCan = 'https://www2.nrcan.gc.ca/eneene/sources/pripri/prices_bycity_e.cfm?PriceYear=0&ProductID=5&LocationID=4&dummy=#PriceGraph';
    const urlNEWRED = 'https://www2.nrcan.gc.ca/eneene/sources/pripri/prices_bycity_e.cfm?productID=5&locationID=4&frequency=W&priceYear=';
    const NorCanPattern = '';
    const urlCdnDom = 'http://specs.maberendezvous.ca/servlet/transportation_view';
    const CdnDomPattern = '';
    
  {$ENDREGION}
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

uses DMod;


{$REGION 'Start Up'}

procedure TForm1.FormActivate(Sender: TObject);
var
  TF : boolean;
  path : string;
begin
  //Show loading panel
  terminate := false;
  InitPanelOpen(self,'Starting FSCUpdater...');

  JudgementDay := EncodeDateTime(2099,12,31,23,59,59,999);

  Debug := false;
  resized := false;
  dateShift := false;
  error := false;

  dateSettings := TFormatSettings.Create;
  dateSettings.DateSeparator := '-';
  dateSettings.ShortDateFormat := 'mm-dd-yyyy';


  //For database configuration
  InitPanelAddSubheading('Reading configuration files');
  DBSetup;
  if not terminate then
  begin
    if pSelDB.Visible then
      InitPanelAddSubheading('Failed connecting to database')
    else
    begin
      InitPanelAddSubheading('Connected to database');
    end;

    PageControl1.Enabled := True;

    //Make the Huckleberry-Endako chart to be used to determine rates based on fuel prices
    //MakeHEChart;
    LoadHEChart;
    LoadNEWREDChart;

    //Currently checks only the fuel prices, no longer calculating the average for the previous months FCA rates
    InitPanelAddSubheading('Getting rates from the web');
    CheckAllRates;

    //Setting the default effective date for FCA rates to the upcoming monday
    effDatePicker.DateTime := GetClosestMonday;

    //Setting the default effective date for FUELENDA to the current date
    DateTimePicker1.DateTime :=  Now;

    //Set date to the closest wednesday that is not in the past
    AgentPicker.DateTime := GetClosestWeekday(4);

    InitPanelClose;
  end
  else
    Application.Terminate;
end;


procedure TForm1.StartupDBReq;
begin
  NewWeek;
  UpdateFCAAVETable;
end;


function TForm1.ApplySavedConfig : integer;
var
  GenIniFile, DBIniFile, userIni : TIniFile;
  hexPass,genIniPath,userIniPath,dbIniPath : string;
  ret : boolean;
  defvals : TStringList;
begin
  //This is gonna be the same in most apps so make a function that takes a default location
  //IF user ini file doesnt exists open general and create default user ini
  userIniPath := UserIniLoc; 

  genIniPath := GenIniLoc;
  if genIniPath = '' then
  begin
    ShowMessage('Failed loading configuration file.' +
                ' Make sure the local ini file in the application directory(' + extractFileDir(application.ExeName) + ') is properly configured.');
    exit
  end
  else
  begin
    try
      GenIniFile := TIniFile.Create(genIniPath);
    except on E: Exception do
      begin
        ShowMessage('Failed loading configuration file. Contact IT to ensure you have the correct file permissions.');
        exit
      end;
    end;
  end;

  defvals := TStringList.Create;
  urlNorCanCfg := genIniFile.ReadString('Scrape','FUELENDA','');
  NorCanPatternCfg := genIniFile.ReadString('Scrape','patternFUELENDA','');
  urlCdnDomCfg := genIniFile.ReadString('Scrape','FSCCDNDOM','');
  CdnDomPatternCfg := genIniFile.ReadString('Scrape','patternFSCCDNDOM','');
  genIniFile.ReadSectionValues('UserDefaults',defvals);
  genIniFile.Free;

  if CreateCopyUserDefs(defvals) = false then
  begin
    ShowMessage('Unable to create user configuration file. Contact IT to ensure you have the correct file permissions.');
    exit
  end;

  if userIniPath = '' then
  begin
    ShowMessage('Failed loading user configuration file.' +
                  ' Make sure the local ini file in the application directory(' + extractFileDir(application.ExeName) + ') is properly configured.');
    exit
  end
  else
  begin
    try
      userIni := TIniFile.Create(userIniPath);
    except on E: Exception do
      begin
        ShowMessage('Failed loading user configuration file. Contact IT to ensure you have the correct file permissions.');
        exit
      end;
    end;
  end;

  result := 0;
  // Read database name from user ini file apply to edit box
  if userIni.ReadString('DB', 'database', '') <> '' then
  begin
    Database := userIni.ReadString('DB', 'database', '');
    inc(result);

    userIni.Free;

    dbIniPath := DBIniLoc + Database + '.ini';

    dbConnArr := GetLoginInfo(Database);
    if length(dbConnArr) > 1 then
    begin
      if dbConnArr[0] <> '' then inc(result);
      if dbConnArr[1] <> '' then inc(result);
    end
    else
      ShowMessage(dbConnArr[0]);
  end
  else
    userIni.Free;
end;


procedure TForm1.DBSetup;
var
  res : boolean;
begin

  DataModule1.DB2.Connected := False;

  if ApplySavedConfig = 3 then
  begin
    res := Dmod.DataModule1.InitDB(Database,dbConnArr[0],decPass(dbConnArr[1]),dbConnArr[2]);
    if res = false then
    begin
      showmessage('Database configuration invalid. Unable to log in.');
      pSelDB.Visible := true;
      pSelDB.Caption := 'If you didn''t want to select "' + Database + '" then select a different database below. Otherwise please contact IT to make sure the database is configured properly';
      pSelDB.BringToFront;
    end
    else
    begin
      if pSelDB.Visible then
        pSelDB.Visible := false;
      StartupDBReq;
    end;
  end
  else
  begin
    showmessage('Loading database configuration from file failed. Please contact IT. App closing...');
    terminate := true;
  end;

end;


procedure Tform1.DBPopup;
var
  dbconf : TDBConfigForm;
  modRes,dlgRes : integer;
begin
  dbconf := DBConfigForm;
  dbconf := TDBConfigForm.Create(nil);
  dbconf.init;

  modRes := dbconf.ShowModal;

  if modRes = mrYes then
  begin
    DBSetup;
  end
  else
    ShowMessage('Database not changed');
end;


procedure TForm1.bChgDBClick(Sender: TObject);
begin
  DBPopup;
end;


procedure TForm1.UpdFCAAVETable;
begin
  try
    DataSource1.DataSet := DataModule1.GetFCAAVETable;
  except on E:Exception do
    ShowMessage(e.Message);
  end;
  if DataSource1.DataSet = nil then
    ShowMessage('Unable to get FCAAVE table');
end;


procedure TForm1.CheckAllRates;
begin
  NorCanDate := NearestNorCanDate;
  NEWREDDate := GetClosestMonday + 1;
  CheckNorCanRate(true);
  CheckNEWREDRate(true);

  //Removed actual scrape since website no longer exists 08\18\2023
  //UpdCdnDom;
end;


procedure TForm1.CheckAve;
var
  monDate : TDateTime;
begin
    FCAAVES := DataModule1.TryGetAverage2;
end;


procedure TForm1.NewWeek;
var
  FCADate,FuelDate{, endDate} : TDateTime;
begin
  try
    FCADate := DataModule1.GetDate('FCA', JudgementDay);
    FuelDate := DataModule1.GetDate('FUELENDA', JudgementDay);

  //Read whether FCA rates or Fuel rates have been updated this week, and check or uncheck boxes accordingly
    if FCADate > (Today - (DayOfWeek(Today) - 1)) then cbFCA.Checked := false;
    if FuelDate > (Today - (DayOfWeek(Today) - 1)) then cbFUELENDA.Checked := false;
  except on E:Exception do
    ShowMessage(e.Message);
  end;
end;


procedure TForm1.UpdCdnDom;
var
  scrapeData : string;
  splitData : TArray<string>;
begin
  scrapeData := RegExCdnDom;
  splitData := scrapeData.Split([':']);

  if length(splitData) = 3 then
  begin
    lCdnDomEffDate.Caption := splitData[0];
    lCdnDomLTL.Caption := splitData[1];
    lCdnDomTL.Caption := splitData[2];
  end
  else
    ShowMessage('Failed to get web rate for FSCCDNDOM. Values are: [' + string.Join(',',splitData)+']');
end;


{$ENDREGION}

{$REGION 'Refresh Button'}

//Currently not in use. Refresh button for scraping FCA rates from the web
procedure TForm1.iRefreshHoverClick(Sender: TObject);
begin
  CheckAllRates;
end;

procedure TForm1.iRefreshHoverMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  iRefreshHover.Visible := False;
  iRefreshPress.Visible := True;
end;

procedure TForm1.iRefreshHoverMouseLeave(Sender: TObject);
begin
  iRefreshHover.Visible := False;
  iRefresh.Visible := True;
end;

procedure TForm1.iRefreshHoverMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  iRefreshHover.Visible := True;
  iRefreshPress.Visible := False;
end;

procedure TForm1.iRefreshMouseEnter(Sender: TObject);
begin
  iRefreshHover.Visible := True;
  iRefresh.Visible := False;
end;

{$ENDREGION}

{$REGION 'Input box responses'}

//Event handlers to make the UX more convenient
procedure TForm1.meAGENTFSCKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  splitStr : TArray<string>;
begin
  if Key = vkReturn then
  begin
    cbAgentFSC.Checked := True;
    splitStr := StringReplace(meAGENTFSC.EditText, '_', '', [rfReplaceAll]).Split([':']);

    SetRoundMode(TRoundingMode.rmUp);

    lAGENT30data.Caption := floattoStr(RoundTo((strTofloat(splitStr[0]) * 0.3), -2));
    lAGENT40data.Caption := floattoStr(RoundTo((strTofloat(splitStr[0]) * 0.4), -2));
    lCLETRUdata.Caption := floattoStr(RoundTo((strTofloat(splitStr[1]) * 0.4), -2));

    SetRoundMode(TRoundingMode.rmNearest);
  end;
end;

procedure TForm1.meBrewKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = vkReturn then
  begin
    cbBrew.Checked := True;
  end;
end;

procedure TForm1.meCOKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = vkReturn then
  begin
    cbCO.Checked := True;
  end;
end;


procedure TForm1.meEnKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = vkReturn then
  begin
    cbEncorp.Checked := True;
  end;
end;


procedure TForm1.meFSCKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = vkReturn then
  begin
    cbFSC.Checked := True;
  end;
end;

procedure TForm1.meFTKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = vkReturn then
  begin
    cbFT.Checked := True;
  end;
end;


procedure TForm1.meGrimKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  {
  Updated on Oct 4 2021 by Jackson
  if Key = vkReturn then
  begin
    cbGrim.Checked := True;
  end;
  }
end;


procedure TForm1.meLTLChange(Sender: TObject);
begin
  if TMaskEdit(Sender).Text > '' then
    cbFCA.Checked := True;
end;

procedure TForm1.meMapKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = vkReturn then
  begin
    cbMapei.Checked := True;
  end;
end;


procedure TForm1.meMMKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = vkReturn then
  begin
    cbMM.Checked := True;
  end;
end;

procedure TForm1.meNFFKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = vkReturn then
  begin
    cbNFF.Checked := True;
  end;
end;

procedure TForm1.meNTKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = vkReturn then
  begin
    cbNatTire.Checked := True;
  end;
end;

procedure TForm1.meQXKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = vkReturn then
  begin
    cbQX.Checked := True;
  end;
end;

procedure TForm1.meRTKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = vkReturn then
  begin
    cbRTRS.Checked := True;
  end;
end;

procedure TForm1.meWFraKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = vkReturn then
  begin
    cbWFra.Checked := True;
  end;
end;



procedure TForm1.eLTLKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = vkReturn then
  begin
    CalcRates(meLTL.EditText,meTL.EditText);
  end;
end;


procedure TForm1.DateTimePicker1Change(Sender: TObject);
begin
  NorCanDate := DateTimePicker1.DateTime;
  CheckNorCanRate(false);
end;

procedure TForm1.DateTimePicker2Change(Sender: TObject);
begin
  NEWREDDate := DateTimePicker2.DateTime;
  CheckNEWREDRate(false);
end;

procedure TForm1.ScrollBox2MouseWheel(Sender: TObject; Shift: TShiftState;
  WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
var
  LTopLeft, LTopRight, LBottomLeft, LBottomRight: SmallInt;
  LPoint: TPoint;
  ScrollBox: TScrollBox;
begin
  ScrollBox := TScrollBox(Sender);
  LPoint := ScrollBox.ClientToScreen(Point(0,0));
  LTopLeft := LPoint.X;
  LTopRight := LTopLeft + ScrollBox.ClientWidth;
  LBottomLeft := LPoint.Y;
  LBottomRight := LBottomLeft + ScrollBox.ClientWidth;
  if (MousePos.X >= LTopLeft) and
    (MousePos.X <= LTopRight) and
    (MousePos.Y >= LBottomLeft) and
    (MousePos.Y <= LBottomRight) then
  begin
    ScrollBox.VertScrollBar.Position := ScrollBox.VertScrollBar.Position - (WheelDelta div 2);
    Handled := True;
  end;
end;

{$ENDREGION}

{$REGION 'Error Functions'}

procedure TForm1.ErrMsg(msg : string);
begin
  if error = false then
  begin
    MessageDlg(msg,mtError,[mbOk],0);
    error := true;
  end;
end;


{$ENDREGION}

{$REGION 'Utility functions'}

function TForm1.GetClosestMonday : TDateTime;
var
  todayDate : TDateTime;
  diff : integer;
begin
  todayDate := Today;

  if DayOfWeek(todayDate) = 1 then
  begin
    result := todayDate + 1;
  end
  else if DayOfWeek(todayDate) <> 2 then
  begin
    diff := 9 - DayOfWeek(todayDate);
    result := todayDate + diff;
  end
  else
  begin
    result := todayDate;
  end;
end;


function TForm1.GetClosestWeekday(weekday : integer) : TDateTime;
var
  todayDate : TDateTime;
  diff : integer;
begin
  todayDate := Today;

  if DayOfWeek(todayDate) < weekday then
  begin
    diff := weekday - DayOfWeek(todayDate);
    result := todayDate + diff;
  end
  else if DayOfWeek(todayDate) > weekday then
  begin
    diff := (7 + weekday) - DayOfWeek(todayDate);
    result := todayDate + diff;
  end
  else
  begin
    result := todayDate;
  end;
end;


function TForm1.LastMonday(date : TDateTime) : Boolean;
var
  newDate : TDateTime;
  diff, month : integer;
begin
  month := MonthOf(date);

  if DayOfWeek(date) = 1 then newDate := date + 1
  else if DayOfWeek(date) <> 2 then
  begin
    diff := 9 - DayOfWeek(date);
    newDate := date + diff;
  end
  else newDate := date + 7;

  if month = 12 then
  begin
    if MonthOf(newDate) = 1 then result := True
    else result := false;
  end
  else
  begin
    if MonthOf(newDate) = (month + 1) then result := True
    else result := false;
  end;

end;


procedure TForm1.lFUAPDblClick(Sender: TObject);
var
  x: Integer;
begin

  //x := AdvInputTaskDialogEx1.Execute;
  //ShowMessage(x.ToString);
  //ShowMessage(AdvInputTaskDialogEx1.InputText);
end;

//Pull the HTML document for a given URL
function TForm1.Scrape(myURL : String) : String;
var
  buff : string;
  htmlOut : TStreamWriter;
begin

  with HTTPclient do
  begin
    try
      buff := Get(myURL);
    except
      on EIdSocketError do messagedlg('There was a problem trying to get the web rate for the URL: ' + myURL,mtInformation,[mbOK],0);
    end;
  end;

  if Debug then
  begin

    //TODO - change this to the local temp or whatever
    htmlOut := TStreamWriter.Create(TFile.Create(ExtractFilePath(Application.ExeName) + 'htmlOut.txt'), TEncoding.ANSI );
    htmlOut.OwnStream;

    htmlOut.Write(buff);
    htmlOut.free;
  end;

  result := buff;
end;

//Regular expressions for all the current needs of this project
function TForm1.RegEx(BuffString, Mode : String) : String;
var
  RegEx : TRegEx;
  Pattern : String;
  Match : TMatch;
  DebugLog : TStreamWriter;
begin
  //TODO - change this to the local temp or whatever
  if Debug then DebugLog := TStreamWriter.Create('RegEx.txt', True);

  if Mode = 'wholeText' then
  begin
    Pattern :=  '<table role="presentation" border="1" cellspacing="0" cellpadding="0">';
    Pattern := Pattern + '\s*<tbody>(\s*<tr>(\s*<td( colspan="2")? valign="top">(<strong>[ A-Za-z]*</strong>)?[ A-Za-z0-9&;:,%\+\.\-]*</td>)+\s*</tr>)+';
  end;

  if Mode = 'rates' then
  begin
    Pattern := '\d\d?\.\d';
  end;

  if Mode = 'rate' then
  begin
    Pattern := '\d+\.\d';
  end;

  if Mode = 'date' then
  begin
      Pattern := '\d\d-\d\d-\d\d\d\d';
  end;

  if Mode = 'norcan' then
  begin
      Pattern := '<td headers="headerDate empty header1" class="nowrap">' + ConvertDate(DateToStr(NorCanDate, dateSettings)) + '</td>';
      Pattern := Pattern + '\s*<td headers="header4_1_1 header3_1 header1">\d+\.\d+</td>'
  end;
  if Mode = 'newred' then
  begin
      Pattern := '<td headers="headerDate empty header1" class="nowrap">' + ConvertDate(DateToStr(NEWREDDate, dateSettings)) + '</td>';
      Pattern := Pattern + '\s*<td headers="header4_1_1 header3_1 header1">\d+\.\d+</td>'
  end;


  if Debug then DebugLog.WriteLine(Pattern);

  // Use regular expression class to try to match the pattern we created
  RegEx.Create(Pattern);
  Match := RegEx.Match(BuffString);

  // If there is a match return it as a string, otherwise return an empty string
  if Match.Success = True then
    begin
      Result := Match.Value;

      if Debug then DebugLog.WriteLine('Success' + sLineBreak);

      if Mode = 'rates' then
      begin
        Match := Match.NextMatch;
        if Match.Success = True then
        begin
          Result := Result + ':' + Match.Value;
          if Debug then DebugLog.WriteLine('Success on second rate' + sLineBreak);
        end
        else
        begin
          Result := '';
          if Debug then DebugLog.WriteLine('Failure on second rate' + sLineBreak);
        end;
      end;
    end
  else
    begin
      Result := '';
      if Debug then DebugLog.WriteLine('Failure' + sLineBreak);
    end;

    if Debug then DebugLog.Free;

end;


function TForm1.RegExCdnDom : string;
var
  //outFile : TStreamWriter;
  RegEx : TRegEx;
  Match : TMatch;
  html,Pattern,effDate,LTL,TL : string;
begin
  var URL : string;
  if urlNorCanCfg = '' then
    URL := urlCdnDom
  else
    URL := urlCdnDomCfg;
  html := Scrape(URL);

  //outFile := TStreamWriter.Create(ExtractFileDir(application.ExeName) + '\testCdnDomRegEx.txt');

  // Try to capture effective date
  Pattern := '<!-- !!!!!!!!CREATE INNER TABLES -->\s*(<table width=100% border=0 cellspacing=0><tr>\s*<td( bgcolor=#F7F7F7)?><span class=small>[ a-zA-Z0-9,]+</span></td></tr></table>\s*)+';
  // Use regular expression class to try to match the pattern we created
  RegEx.Create(Pattern);
  Match := RegEx.Match(html);
  if Match.Success then
  begin
    Match := Match.NextMatch;

    if Match.Success then
    begin
      effDate := ExtractValCdnDom(Match.Value);
    end;
  end;

  //Now try to match the LTL and TL
  Match := RegEx.Match(html,'<!-- !!!!!!!!CREATE INNER TABLES -->\s*(<table width=100% border=0 cellspacing=0><tr>\s*<td( bgcolor=#F7F7F7)?><span class=small>(&nbsp;)+[ a-zA-Z0-9\.]+</span></td></tr></table>\s*)+');
  if Match.Success then
  begin
    //This is the fuel price thing
    Match := Match.NextMatch;
    if Match.Success then
    begin
      LTL := ExtractValCdnDom(Match.Value, true);

      Match := Match.NextMatch;
      if Match.Success then
      begin
        TL := ExtractValCdnDom(Match.Value, true);
      end;
    end;
  end;

  //outfile.WriteLine(effDate + ':' + LTL + ':' + TL);
  //outfile.Free;

  result := effDate + ':' + LTL + ':' + TL;
end;


function TForm1.ExtractValCdnDom(regExMatch : string; rate : boolean = false) : string;
var
  cleaned : string;
  cleanedArr : TArray<string>;
begin
  cleaned := ReplaceStr(regExMatch,'<!-- !!!!!!!!CREATE INNER TABLES -->','');
  cleaned := ReplaceStr(cleaned,'<table width=100% border=0 cellspacing=0><tr>','');
  cleaned := ReplaceStr(cleaned,'<td bgcolor=#F7F7F7><span class=small>','');
  cleaned := ReplaceStr(cleaned,'<td>','');
  cleaned := ReplaceStr(cleaned,'<span class=small>','');
  cleaned := ReplaceStr(cleaned,',','');
  cleaned := ReplaceStr(cleaned,slinebreak,'');
  cleaned := ReplaceStr(cleaned,'</span></td></tr></table>',',');
  if rate then
    cleaned := ReplaceStr(cleaned,'&nbsp;',',');

  cleanedArr := cleaned.Split([',']);

  result := trim(cleanedArr[length(cleanedArr) -2]);
end;


{$ENDREGION}

{$REGION 'FCA Functions/Procedures'}

function TForm1.RefineRatesHTML(myURL : String) : TArray<String>;
var
  rawHTML, targetStr, rates, Date : String;
  ratesList : TArray<String>;
begin
  rawHTML := Scrape(myURL);

  targetStr := RegEx(rawHTML, 'wholeText');

  rates := RegEx(targetStr, 'rates');
  ratesList := rates.Split([':']);

  Date := RegEx(targetStr, 'date');

  SetLength(Result, 3);
  result[0] := ratesList[0];
  result[1] := ratesList[1];
  result[2] := date;
end;


procedure TForm1.CheckFCARates;
var
  webVals : TArray<String>;
  webDate : TDateTime;
begin

  webVals := RefineRatesHTML(urlFSC);

  webDate := StrToDate(webVals[2], dateSettings);

  // -1 here means LessThanValue
  if CompareDate(lastDate, webDate) = -1 then
  begin
    pStatusPanel.Caption := 'New Rates Downloaded';
    lastDate := webDate;

    meLTL.EditText := webVals[0];
    meTL.EditText := webVals[1];

    CalcRates(webVals[0], webVals[1]);

  end;

end;


procedure TForm1.CalcRates(ltl,tl : string);
var
  fLTL,fTL : double;
  sLTL,sTL : string;
begin
  SetRoundMode(TRoundingMode.rmUp);


  if (ltl = '__.__') then fLTL := 0
  else
  begin
    sLTL := StringReplace(ltl, '_', '', [rfReplaceAll]);
    fLTL := RoundTo(strtofloat(sLTL), -2);
  end;

  if (tl = '__.__') then fTL := 0
  else
  begin
    sTL := StringReplace(tl, '_', '', [rfReplaceAll]);
    fTL := RoundTo(strtofloat(sTL), -2);
  end;


  //FCA2 rates
  l2LTL.Caption := FloatToStr(RoundTo(fLTL * 1.02, -2));
  l2TL.Caption := FloatToStr(RoundTo(fTL * 1.02, -2));

  //FCA4 rates
  l4LTL.Caption := FloatToStr(RoundTo(fLTL * 1.04, -2));
  l4TL.Caption := FloatToStr(RoundTo(fTL * 1.04, -2));

  //FCA 4 Point rates
  l4ptLTL.Caption := FloatToStr(RoundTo(fLTL + 4, -2));
  l4ptTL.Caption := FloatToStr(RoundTo(fTL + 4, -2));

  //FCA90  rates
  l90LTL.Caption := FloatToStr(RoundTo(fLTL * 0.9, -2));
  l90TL.Caption := FloatToStr(RoundTo(fTL * 0.9, -2));

  //FUELUAP
  lFUAPLTL.Caption := FloatToStr(RoundTo(fLTL * 0.6, -2));
  lFUAPTL.Caption := FloatToStr(RoundTo(fTL * 0.6, -2));

  SetRoundMode(TRoundingMode.rmNearest);
end;


//Fetch the averages in case someone forgot to add them to the table when we calculated them
procedure TForm1.cbFCAAVEClick(Sender: TObject);
var
  monDate : TDateTime;
begin
  if cbFCAAVE.Checked then
  begin
    try
      datamodule1.DB2.StartTransaction;
      cbCalcAve;
      datamodule1.DB2.Commit;
    except on E:Exception do
      begin
        ShowMessage(e.Message + slinebreak + 'Rolling back transaction');
        datamodule1.DB2.Rollback;
      end;
    end;
  end
  else
  begin
    laveLTLdata.Caption := '';
    laveTLdata.Caption := '';
    laveMonth.Caption := '';
    FCAAVES := nil;
  end;
end;

procedure TForm1.cbCalcAve;
begin
  //This was allowing users to calculate the average prematurely and update with the wrong value
  //CheckAve;
  FCAAVES := DataModule1.GetAverages;

  if not ((FCAAVES[0] = 0) or (FCAAVES[1] = 0)) then
  begin
    laveMonth.Caption := 'From: ' + inttostr(monthof(Datamodule1.GetAveDate)) + '/' + inttostr(yearof(Datamodule1.GetAveDate));
    laveLTLdata.Caption := floattostr(Roundto(FCAAVES[0], -2));
    laveTLdata.Caption := floattostr(RoundTo(FCAAVES[1], -2));
    UpdateFCAAVETable;
  end;
end;


{$ENDREGION}

{$REGION 'Norcan functions/procedures'}

//Representation of the Huckleberry-Endako chart to be used to determine rates based on fuel prices
//Dictionary with a key of the fuel price, and a value of the corresponding LTL and TL rates seperated by a colon character
procedure TForm1.MakeHEChart;
var
  //perLitre : double;
  pLt : integer;
  tlval : double;
  I : integer;
  valStr : string;
const
  ltlvals : array[1..236] of string = ('0.00','0.17','0.33','0.50','0.66','0.83','1.00','1.16','1.33','1.50','1.66','1.83','2',
  {original size [1..105]}             '2.16','2.33','2.49','2.66','2.83','2.99','3.16','3.33','3.49','3.66','3.82','3.99','4.16',
                                       '4.32','4.49','4.66','4.82','4.99','5.15','5.32','5.49','5.65','5.82','5.99','6.15','6.32',
                                       '6.48','6.65','6.82','6.98','7.15','7.32','7.48','7.65','7.81','7.98','8.15','8.31','8.48',
                                       '8.65','8.81','8.98','9.14','9.31','9.48','9.64','9.81','9.98','10.14','10.31','10.47',
                                       '10.64','10.81','10.97','11.14','11.31','11.47','11.64','11.8','11.97','12.14','12.3',
                                       '12.47','12.64','12.8','12.97','13.13','13.3','13.47','13.63','13.8','13.97','14.13',
                                       '14.3','14.46','14.63','14.8','14.96','15.13','15.3','15.46','15.63','15.79','15.96',
                                       '16.13','16.29','16.46','16.63','16.79','16.96','17.12','17.29',
                                       {New temporary extension of the chart for rising fuel prices}
                                       '17.47','17.64','17.8','17.97','18.14','18.3','18.47','18.64','18.8','18.97','19.13',
                                       '19.3','19.47','19.63','19.8','19.97','20.13','20.3','20.47','20.63','20.8','20.97',
                                       '21.13','21.3','21.46','21.63','21.8','21.96','22.13','22.3','22.46','22.63','22.8',
                                       '22.96','23.13','23.29','23.46','23.63','23.79','23.96','24.13','24.29','24.46','24.63',
                                       '24.79','24.96','25.12','25.29','25.46','25.62','25.79','25.96','26.12','26.29','26.46',
                                       '26.62','26.79','26.96','27.12','27.29','27.45','27.62','27.79','27.95','28.12','28.29',
                                       '28.45','28.62','28.79','28.95','29.12','29.28','29.45','29.62','29.78','29.95','30.12',
                                       '30.28','30.45','30.62','30.78','30.95','31.11','31.28','31.45','31.61','31.78','31.95',
                                       '32.11','32.28','32.45','32.61','32.78','32.95','33.11','33.28','33.44','33.61','33.78',
                                       '33.94','34.11','34.28','34.44','34.61','34.78','34.94','35.11','35.27','35.44','35.61',
                                       '35.77','35.94','36.11','36.27','36.44','36.61','36.77','36.94','37.1','37.27','37.44',
                                       '37.6','37.77','37.94','38.1','38.27','38.44','38.6','38.77','38.94','39.1');
begin

  //perLitre := $0.80;
  pLt := 80;
  tlval := 0.00;

  HEChart := TDictionary<integer,string>.create;

  for I := 1 to 236 do
  begin
    valStr := ltlvals[I] + ':' + FloatToStr(RoundTo(tlval, -2));
    HEChart.Add(pLt, valStr);
    pLt := pLt + 1;
    tlval := tlval + 0.25;
  end;
end;


procedure TForm1.LoadHEChart;
var
  dFile : TStreamReader;
  I : integer;
  buffer,filePath : string;
  lines,vals : TArray<string>;
begin
  HEloaded := false;
  filePath := CommonFiles + 'HEChart.csv';
  if FileExists(filePath) then
  begin
    //Load and read csv file containing chart
    dFile := TStreamReader.Create(filePath);
    buffer := dFile.ReadToEnd;
    if buffer <> '' then
    begin
      //Split into rows
      lines := buffer.Split([slinebreak]);
      if length(lines) > 5 then
      begin
        HEChart := TDictionary<integer,string>.create;
        for I := 4 to (length(lines) - 1) do
        begin
          vals := lines[I].Split([',']);
          if (length(vals) = 3) and ((vals[0] <> '') and (vals[1] <> '') and (vals[2] <> '')) then
          begin
            HEChart.Add(Round((strtofloat(vals[0]) * 100)),vals[1] + ':' + vals[2]);
            if I = 4 then HEmin := Round((strtofloat(vals[0]) * 100));
            if I = (length(lines) - 2) then HEmax := Round((strtofloat(vals[0]) * 100));
          end;
        end;
        HEloaded := true;
      end
      else
      begin
        errMsg('There is an issue with the file containing the Huckleberry-Endako Chart. Try loading one using the button in the top right labeled "Import Huckleberry-Endako Chart", or contact support.');
      end;
    end
    else
    begin
      errMsg('File for Huckleberry-Endako Chart is empty. Try loading a new one using the button in the top right labeled "Import Huckleberry-Endako Chart", or contact support.');
    end;
    dFile.Free;
  end
  else
  begin
    errMsg('Unable to find file for Huckleberry-Endako Chart. Try loading one using the button in the top right labeled "Import Huckleberry-Endako Chart", or contact support.');
  end;

end;


procedure TForm1.LoadNEWREDChart;
var
  dFile : TStreamReader;
  I : integer;
  buffer,filePath : string;
  lines,vals : TArray<string>;
begin
  NEWREDloaded := false;
  filePath := CommonFiles + 'NEWREDChart.csv';
  if FileExists(filePath) then
  begin
    //Load and read csv file containing chart
    dFile := TStreamReader.Create(filePath);
    buffer := dFile.ReadToEnd;
    if buffer <> '' then
    begin
      //Split into rows
      lines := buffer.Split([slinebreak]);
      if length(lines) > 5 then
      begin
        NEWREDChart := TDictionary<integer,string>.create;
        for I := 4 to (length(lines) - 1) do
        begin
          vals := lines[I].Split([',']);
          if (length(vals) = 3) and ((vals[0] <> '') and (vals[1] <> '') and (vals[2] <> '')) then
          begin
            NEWREDChart.Add(Round((strtofloat(vals[0]) * 100)),vals[1] + ':' + vals[2]);
            if I = 4 then NEWREDmin := Round((strtofloat(vals[0]) * 100));
            if I = (length(lines) - 2) then NEWREDmax := Round((strtofloat(vals[0]) * 100));
          end;
        end;
        NEWREDloaded := true;
      end
      else
      begin
        errMsg('There is an issue with the file containing the NEWRED Chart. Try loading one using the button in the top right labeled "Import NEWRED Chart", or contact support.');
      end;
    end
    else
    begin
      errMsg('File for NEWRED Chart is empty. Try loading a new one using the button in the top right labeled "Import NEWRED Chart", or contact support.');
    end;
    dFile.Free;
  end
  else
  begin
    errMsg('Unable to find file for NEWRED Chart. Try loading one using the button in the top right labeled "Import NEWRED Chart", or contact support.');
  end;

end;


//Return the nearest valid date for fuel price. Excludes weekends
function TForm1.NearestNorCanDate : TDateTime;
var
   currDate: TDateTime;
   targetDate : TDateTime;
begin

  currDate:= Date;

  if DayOfWeek(currDate) = 1 then targetDate := IncDay(currDate, -2)
  else if DayOfWeek(currDate) = 7 then targetDate := IncDay(currDate, -1)
  else targetDate := currDate;

  result := targetDate
end;


//Converts between date formats. Works both ways.
//Norcan uses YYYY-MM-DD, where FCA uses DD-MM-YYYY.
function TForm1.ConvertDate(date : string) : string;
var
  splitStr : TArray<string>;
begin
  SetLength(splitStr,3);

  if RegEx(date,'date') = '' then
  begin
    splitStr := date.Split(['-'], 3);
    result := splitStr[1] + '-' + splitStr[2] + '-' + splitStr[0];
  end
  else
  begin
    splitStr := date.Split(['-'], 3);
    result := splitStr[2] + '-' + splitStr[0] + '-' + splitStr[1];
  end;

end;


//Scrape fuel price from the Norcan website. Use two different Regular Expression operations to extract the Price.
//Then use the Huckleberry-Endako chart to determine the rates. Add the rates to execution list for the update
procedure TForm1.CheckNorCanRate(onStartup : boolean);
var
  htmlrefined, webRate, rates : string;
  D : integer;
  //B : boolean;
  //arr : TArray<double>;
  //I: Integer;
begin
  var URL,year : string;
  if urlNorCanCfg = '' then
    URL := urlNorCan{ + year}
  else
    URL := urlNorCanCfg;
  
  if NorCanHtml = '' then NorCanHtml := Scrape(URL);

  htmlrefined := RegEx(NorCanHtml,'norcan');

  webRate := RegEx(htmlrefined, 'rate');


  if webRate <> '' then
  begin
    lWebRateData.Caption := webRate;

    DateTimePicker1.DateTime := NorCanDate;
    D := trunc(SimpleRoundTo(strtofloat(webRate), 0));                        

    if HEloaded then
    begin
      if D < HEmin then D := HEmin;
      if D > HEmax then D := HEmax;
      if HEChart.TryGetValue(D, rates) then
      begin
        var ratesList := rates.Split([':'], 2);
        lFuelLTLData.Caption := ratesList[0];
        lFuelTLData.Caption := ratesList[1];
        cbFuelenda.Checked := true;
      end
      else
        MessageDlg('Unable to use HE Chart. Index value is: ' + inttostr(D)
                    + slinebreak + booltostr(HEChart.ContainsKey(D),true),mtError, [mbOK],0,mbOK);
    end
    else
      MessageDlg('HE Chart not loaded. Unable to convert value: ' + inttostr(D),mtError, [mbOK],0,mbOK);

    if (dateShift = True) and (onStartup = false) then
    begin
      MessageDlg('The date you selected ('
                 + datetostr(unshiftedDate, dateSettings)
                 + ') does not have a valid rate. Displaying ('
                 + datetostr(NorCanDate, dateSettings)
                 + ') instead.', mtInformation, [mbOK], 0, mbOK);

      DateTimePicker1.DateTime := NorCanDate;
      dateShift := false;
    end;

  end
  else
  begin
    if dateShift = false then
    begin
      dateShift := true;
      unshiftedDate := NorCanDate;
    end;

    NorCanDate := NorCanDate - 1;

    CheckNorCanRate(onStartup);

  end;

end;


procedure TForm1.CheckNEWREDRate(onStartup : boolean);
var
  htmlrefined, webRate, rates : string;
  D : integer;
begin
  var URL,year : string;
  DateTimeToString(year, 'YYYY', NorCanDate);
  URL := urlNEWRED + year;

  if NEWREDHtml = '' then NEWREDHtml := Scrape(URL);

  htmlrefined := RegEx(NEWREDHtml,'newred');

  webRate := RegEx(htmlrefined, 'rate');


  if webRate <> '' then
  begin
    lWebNEWRED.Caption := webRate;

    DateTimePicker2.DateTime := NEWREDDate;
    D := trunc(SimpleRoundTo(strtofloat(webRate), 0));

    if NEWREDloaded then begin
      if D < NEWREDmin then D := NEWREDmin;
      if D > NEWREDmax then D := NEWREDmax;
      if NEWREDChart.TryGetValue(D, rates) then
      begin
        var ratesList := rates.Split([':'], 2);
        lNEWREDLTL.Caption := ratesList[0];
        lNEWREDTL.Caption := ratesList[1];
        cbNEWRED.Checked := true
      end
      else
        MessageDlg('Unable to use NEWRED Chart. Index value is: ' + inttostr(D)
                    + slinebreak + booltostr(NEWREDChart.ContainsKey(D),true),mtError, [mbOK],0,mbOK);
    end
    else
      MessageDlg('NEWRED Chart not loaded. Unable to convert value: ' + inttostr(D),mtError, [mbOK],0,mbOK);

    if (dateShift = True) and (onStartup = false) then
    begin
      MessageDlg('The date you selected ('
                 + datetostr(unshiftedDate, dateSettings)
                 + ') does not have a valid rate. Displaying ('
                 + datetostr(NEWREDDate, dateSettings)
                 + ') instead.', mtInformation, [mbOK], 0, mbOK);

      DateTimePicker2.DateTime := NEWREDDate;
      dateShift := false;
    end;

  end
  else
  begin
    if dateShift = false then
    begin
      dateShift := true;
      unshiftedDate := NEWREDDate;
    end;

    NEWREDDate := NEWREDDate - 1;

    CheckNEWREDRate(onStartup);

  end;

end;



{$ENDREGION}

{$Region 'Before Update Operations'}

//Generate the execution list for our database operations.
//Format is: ACODE_ID:ACD_RANGE_TO:CLIENT_ID/LTL:CLIENT_ID/TL:CLIENT_ID/OTHER,CLIENT_ID/OTHER... etc
//IF there are multiple CLIENT_IDs for an ACODE_ID it looks like the following:
//:...\CLIENT_ID1-CLIENT_ID2-etc/rate:
procedure TForm1.PrepareUpdateList;


procedure TForm1.ClearUpdateList;
begin
  FCArates := nil;
end;


//Helper procedure for adding entries to the execution list
procedure TForm1.AddToList(header, data : string);
var
  splitStr : TArray<string>;
begin
  SetLength(FCArates, length(FCArates) + 1);

  if ContainsText(data, ':') then
  begin

    splitStr := StringReplace(data, '_', '', [rfReplaceAll]).Split([':']);

    if length(splitStr) = 3 then
    begin
      FCArates[length(FCArates) - 1] := header + ':/' + splitStr[0] + ':/' + splitStr[1] + ':/' + splitStr[2];
    end
    else
    begin
      FCArates[length(FCArates) - 1] := header + ':/' + splitStr[0] + ':/' + splitStr[1] + ':/0';
    end;
  end
  else FCArates[length(FCArates) - 1] := header + ':/0:/0:/' + StringReplace(data, '_', '', [rfReplaceAll]);

end;


//No longer generates a dictionary, instead uses an array for updating the interliner rates.
//Now we just do it seperately
//Format is Vendor_ID/(LTL or other):TL
procedure TForm1.PrepareINTDict;


procedure TForm1.ClearINTDict;
begin
  INTFC := nil;
end;


//A simple check for the input masked edit boxes
function TForm1.InputValid(val : string) : Boolean;
var
  temp : string;
begin
  result := false;

  temp := StringReplace(val, '_', '', [rfReplaceAll]);
  temp := StringReplace(temp, '.', '', [rfReplaceAll]);
  temp := StringReplace(temp, ':', '', [rfReplaceAll]);

  if Length(val) = 17 then
  begin
    if Length(temp) >= 6 then result := True;
  end;

  if Length(val) = 11 then
  begin
    if Length(temp) >= 4 then result := True;
  end;

  if Length(val) = 5 then
  begin
    if Length(temp) >= 2 then result := True;
  end;

end;


function TForm1.ChkdRatesStr : string;
begin
  //Example
  if rate.Checked then result := result + 'Rate Name' + slinebreak;
end;

{$ENDREGION}


procedure TForm1.bUpdateClick(Sender: TObject);
var
  I, conf, replace, doAve : integer;
  splitStr : TArray<string>;
  monthlydate,zero,currEffDate : TDateTime;
  notUpd, activeDates : string;
  FCAAVEreplace : boolean;
begin
  try
    datamodule1.DB2.StartTransaction;
    FCAAVEreplace := false;
    //Create string to display dates being updated
    activeDates := 'Effective Date: ' + datetostr(effDatePicker.DateTime, datesettings);
    //Removed: checks to populate the UI with effective dates for checked rates
    activeDates := activeDates + slinebreak + slinebreak;

    conf := messagedlg(activeDates + 'Are you sure you want to update the following rate(s):' + slinebreak + slinebreak + ChkdRatesStr,
                       mtConfirmation, [mbYes,mbNo], 0);

    if conf = mrNo then Exit;

    currEffDate := effDatePicker.DateTime;

    if cbFCA.Checked then
    begin
      CalcRates(meLTL.EditText,meTL.EditText);
      if inputvalid(meLTL.EditText) and inputValid(meTL.EditText) then
        DataModule1.AddAveEntry(currEffDate,
                                strtofloat(StringReplace(meLTL.EditText, '_', '', [rfReplaceAll])),
                                strtofloat(StringReplace(meTL.EditText, '_', '', [rfReplaceAll])), false);
    end;

    //create a date that represents zero
    zero := encodeDate(1899,12,31);

    PrepareUpdateList;
    PrepareIntDict;

    if error then begin error := false; exit; end;

    replace := 0;
  
    for I := 0 to (length(FCArates) - 1) do
    begin
      splitStr := FCArates[I].Split([':'], 5);

      if splitStr[0] = 'AGENTFSC' then
      begin
        currEffDate := agentPicker.DateTime;
      end;

      if DataModule1.GetDate(splitStr[0], JudgementDay) = currEffDate then
      begin
        if replace = 0 then
        begin
          replace := messageDlg(
            'An entry already exists for code ' + splitStr[0] + ' with the effective date ('
            + datetostr(currEffDate, dateSettings) + '). Would you like to replace this entry?',
            mtConfirmation, 
            [mbYes,mbNo,mbYesToAll,mbNoToAll], 
            0
          );
        end;

        if replace = mrNo then replace := 0
        else if (replace = mrYes) or (replace = mrYesToAll) then
        begin //do replace procedure
          DataModule1.UpdateForCode(
            splitStr[0], 
            StrToint(splitStr[1]),
            splitStr[2], 
            splitStr[3], 
            splitstr[4],
            currEffDate, 
            JudgementDay, 
            true
          );

          if splitStr[0] = 'FCA' then
          begin
            DataModule1.TryReplaceFCAAVE( 
              currEffDate,
              strtofloat( StringReplace(splitStr[2],'/','',[]) ),
              strtofloat( StringReplace( splitStr[3].Split([','])[0] ,'/','',[]) ) 
            );
            FCAAVEreplace := true;
          end;
          
          if replace = mrYes then replace := 0;
        end;
      end
      else if DataModule1.GetDate(splitStr[0], JudgementDay) < zero then
      begin //add to list to inform user it was not updated and why
        notUpd := notUpd + splitStr[0] + slineBreak;
      end
      else
      begin
        DataModule1.UpdateForCode(splitStr[0], StrToint(splitStr[1]),
                                  splitStr[2], splitStr[3], splitstr[4],
                                  currEffDate, JudgementDay, false);
      end;

      if splitStr[0] = 'AGENTFSC' then
      begin
        currEffDate := effDatePicker.DateTime;
      end;

    end;

    //Do update for the rates with the code INTFC
    for I := 0 to (length(INTFC) - 1) do
    begin
      splitStr := INTFC[I].Split(['/'], 2);
      
      if DataModule1.GetINTFCDate(splitStr[0], JudgementDay) = currEffDate then
      begin
        if replace = 0 then
        begin
          replace := messageDlg('An entry already exists for code INTFC-' + splitStr[0] + ' with the effective date ('
                              + datetostr(currEffDate, dateSettings)
                              + '). Would you like to replace this entry?',
                              mtConfirmation, [mbYes,mbNo,mbYesToAll,mbNoToAll], 0);
        end;

        if replace = mrNo then replace := 0
        else if (replace = mrYes) or (replace = mrYesToAll) then
        begin //do replace procedure
          DataModule1.UpdateInterliner(splitStr[0], splitStr[1], currEffDate, JudgementDay, true);

          if replace = mrYes then replace := 0;
        end;
      end
      else if DataModule1.GetINTFCDate(splitStr[0], JudgementDay) < zero then //add to list to inform user it was not updated and why
         notUpd := notUpd + 'INTFC: ' + splitStr[0] + slineBreak
      else
        DataModule1.UpdateInterliner(splitStr[0], splitStr[1], currEffDate, JudgementDay, false);
    end;

    //Do update for these rates if it is the last monday of the month
    if LastMonday(currEffDate) and (cbFCA.Checked) then
    begin
      if monthOf(currEffDate) = 12 then
        monthlydate := StartOfAMonth(YearOf(currEffDate) + 1, 1)
      else
        monthlydate := StartOfAMonth(YearOf(currEffDate), MonthOf(currEffDate) + 1);
      if DataModule1.GetDate('MONTHRATE', JudgementDay) < zero then
        notUpd := notUpd + 'MONTHRATE' + slineBreak
      else if DataModule1.GetDate('MONTHRATE', JudgementDay) <> monthlydate then
        DataModule1.UpdateForCode('MONTHRATE',10003, '/' + StringReplace(meLTL.EditText, '_', '', [rfReplaceAll]), '/' + StringReplace(meTL.EditText, '_', '', [rfReplaceAll]), '/0', WESFREdate, JudgementDay, false);

      //Now update FCAAVE
      if (FCAAVES = nil) and (cbFCAAVE.Checked = false) and (FCAAVEreplace = false) then
      begin
        CheckAve;

        if (FCAAVES <> nil) and (DataModule1.GetDate('FCAAVE', JudgementDay) <> monthlydate) then
        begin
          doAve := messageDlg('FCAAVE is ready to be updated with the values:' + slineBreak + slinebreak
                              + 'LTL: ' + floattostr(RoundTo(FCAAVES[0], -2)) + slinebreak
                              + 'TL: ' + floattostr(RoundTo(FCAAVES[1], -2))  + slinebreak
                              + 'Would you like to update?',
                              mtConfirmation, [mbYes,mbNo], 0);
          if doAve = mrYes then
          begin
            DataModule1.UpdateForCode('FCAAVE',10003,'/' + floattostr(RoundTo(FCAAVES[0], -2)),'/' + floattostr(RoundTo(FCAAVES[1], -2)),
                                      '/' + floattostr(RoundTo(FCAAVES[0], -2)),monthlydate, JudgementDay, false);
            FCAAVES := nil;
          end;
          UpdateFCAAVETable;
        end;
      end;
    end;
    datamodule1.DB2.Commit;
  
    //Display message if rates were not updated
    if (length(notUpd) > 0) then
    begin
      var msgBody := 'The following codes did not have an active entry in the database: '+slinebreak+slinebreak+notUpd+slinebreak+slinebreak 
                  +  'As a result, they were not updated. You may have to manage the ACHARGE_DETAIL table manually.'; 
      messageDlg(msgBody,mtInformation, [mbOk], 0);
    end;

    ClearUpdateList;
    ClearINTDict;
    UpdateFCAAVETable;
  
  except on E:Exception do
    begin
      ShowMessage(e.Message + slinebreak + 'Rolling back transaction');
      datamodule1.DB2.Rollback;
    end;
  end;
end;


procedure TForm1.bImpHEChartClick(Sender: TObject);
var
  Origpath, Destpath : string;
  dlg : TOpenDialog;
begin
  //Do something
  Origpath := ''; Destpath := '';
  try
    dlg := TOpenDialog.Create(nil);
    if dlg.Execute then
    begin
      if FileExists(dlg.FileName) then
      begin
        Origpath := dlg.FileName;
        try
          Destpath := CommonFiles + 'HEChart.csv';
          if FileExists(Destpath) then
            TFile.Delete(Destpath);
          TFile.Copy(Origpath,Destpath);
          try
            LoadHEChart;
            if not error then
              ShowMessage('Success!');
          except on E:Exception do 
            errMsg('Failed to load HE Chart. Please report this messaage to IT. Message: ' + E.Message);  
          end;
        except
          on E: Exception do
          begin
            ErrMsg('Please report the following message to IT: ' + E.Message);
          end;
        end;
      end
      else
        ErrMsg('The file you specified does not exist.');
    end;
  finally
    dlg.Free;
  end;
end;


procedure TForm1.bImpNEWREDChartClick(Sender: TObject);
var
  Origpath, Destpath : string;
  dlg : TOpenDialog;
begin
  //Do something
  Origpath := ''; Destpath := '';
  try
    dlg := TOpenDialog.Create(nil);
    if dlg.Execute then
    begin
      if FileExists(dlg.FileName) then
      begin
        Origpath := dlg.FileName;
        try
          Destpath := CommonFiles + 'NEWREDChart.csv';
          if FileExists(Destpath) then
            TFile.Delete(Destpath);
          TFile.Copy(Origpath,Destpath);
          try
            LoadNEWREDChart;
            if not error then
              ShowMessage('Success!');
          except on E:Exception do 
            errMsg('Failed to load NEWRED Chart. Please report this messaage to IT. Message: ' + E.Message);  
          end;
        except
          on E: Exception do
          begin
            ErrMsg('Please report the following message to IT: ' + E.Message);
          end;
        end;
      end
      else
        ErrMsg('The file you specified does not exist.');
    end;
  finally
    dlg.Free;
  end;
end;


procedure TForm1.bReCalcFCAAVEClick(Sender: TObject);
begin
  try
    datamodule1.DB2.StartTransaction;
    CheckAve;
    UpdateFCAAVETable;
    datamodule1.DB2.Commit;
  except on E:Exception do
    begin
      ShowMessage(e.Message + slinebreak + 'Rolling back transaction');
      datamodule1.DB2.Rollback;
    end;
  end;
end;


procedure TForm1.bRefreshFCAAVEClick(Sender: TObject);
begin
  UpdateFCAAVETable;
end;


procedure TForm1.UpdateFCAAVETable;
begin
  UpdFCAAVETable;

  FCAAVETable.Columns[0].Width := 250;
  FCAAVETable.Columns[1].Width := 100;
  FCAAVETable.Columns[2].Width := 100;
  FCAAVETable.Columns[3].Width := 100;
  FCAAVETable.Columns[4].Width := 100;
end;

end.
