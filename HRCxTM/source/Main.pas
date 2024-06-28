unit Main;

{

for now replace all and delim with special character: "


}

interface

uses
  Winapi.Windows, Winapi.Messages, ShellApi, ShlObj,
  System.SysUtils, System.Variants, System.Classes, System.IniFiles, System.Character, System.IOUtils,
  System.StrUtils, System.Math, System.DateUtils, System.UITypes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,
  Data.DB, Vcl.Grids, Vcl.DBGrids,
  ScBridge, ScUtils, ScSFTPUtils, ScSSHUtils, ScSFTPClient, ScSSHClient,
  Hydra.Core.BaseModuleManager, Hydra.VCL.ModuleManager, Hydra.VCL.Interfaces,
  CryRptsDotNet_Import, CryUtilsNonVCL,
  {OutlookUtils,} utils, sharedServices, DBConfigAdmin, crypto;

type
  TForm1 = class(TForm)
    SSHClient: TScSSHClient;
    SFTPClient: TScSFTPClient;
    FileStor: TScFileStorage;
    ModMgr: THYModuleManager;
    bRun: TButton;
    lExchSuccess: TLabel;
    Panel1: TPanel;
    Panel2: TPanel;
    p5: TPanel;
    bPatch: TButton;
    lPatchSuccess: TLabel;
    p3: TPanel;
    lDownloadSuccess: TLabel;
    bDownload: TButton;
    p2: TPanel;
    lUploadSuccess: TLabel;
    bUpload: TButton;
    p1: TPanel;
    lExtractSuccess: TLabel;
    bExtract: TButton;
    bTest: TButton;
    Panel8: TPanel;
    pRetry: TPanel;
    bRetry: TButton;
    pDB: TPanel;
    lDB: TLabel;
    Label2: TLabel;
    bDBpass: TButton;
    p4: TPanel;
    lProcessSuccess: TLabel;
    bProcess: TButton;
    Panel3: TPanel;
    lChangeCID: TLabel;
    bChangeCID: TButton;
    Panel4: TPanel;
    Panel5: TPanel;
    Label3: TLabel;
    Label4: TLabel;
    eOldCID: TEdit;
    eNewCID: TEdit;
    Label1: TLabel;
    eCIDname: TEdit;
    procedure SSHClientServerKeyValidate(Sender: TObject; NewServerKey: TScKey;
      var Accept: Boolean);
    procedure SFTPClientCreateLocalFile(Sender: TObject; const LocalFileName,
      RemoteFileName: string; Attrs: TScSFTPFileAttributes;
      var Handle: NativeUInt);
    procedure FormActivate(Sender: TObject);
    procedure SFTPClientDirectoryList(Sender: TObject; const Path: string;
      const Handle: TArray<System.Byte>; FileInfo: TScSFTPFileInfo;
      EOF: Boolean);
    procedure bRunClick(Sender: TObject);
    procedure bPatchClick(Sender: TObject);
    procedure bExtractClick(Sender: TObject);
    procedure bUploadClick(Sender: TObject);
    procedure bDownloadClick(Sender: TObject);
    procedure bTestClick(Sender: TObject);
    procedure bDBpassClick(Sender: TObject);
    procedure bRetryClick(Sender: TObject);
    procedure bProcessClick(Sender: TObject);
    procedure bChangeCIDClick(Sender: TObject);
  private
    { Private declarations }

    cryPlugin : IHYVCLNonVisualPlugin;
    sftpFileList, filelist, NewIDSummary : TArray<TArray<string>>;
    profFilter : TArray<string>;
    profFilterPath, testCollector : string;

    CIDwordsUsed : integer;

    function Startup : boolean;

    function RunDataExchange : boolean;

    function Extract : boolean;
    function Upload : boolean;
    function Download : boolean;
    function Process : integer;
    function PatchClientIDs : boolean;
    function TaxExemptReport : boolean;

    function ProcessCustExtract(fileName : string; insert : boolean) : string;
    procedure MoveProcessedFile(name,newPath : string; n : integer = 0);

    //Config and Logging
    function DBlogon : string;
    function DoDBConfig : boolean;
    function LoadConfig : boolean;
    procedure SetupFileDirs;

    //SFTP
    function InitSFTP : boolean;
    function UploadSFTP(source,dest : string) : boolean;
    function DnloadSFTP(source,dest : string) : boolean;
    function MoveSFTP(source,dest : string) : boolean;
    function GetFileList(dirs : TArray<string>) : boolean;
    function DestroySFTP : boolean;

    //Extracts
    function ExpCSV(reportPath,csvPath : string)  : boolean;
    function CleanupCSV(filename : string) : string;
    function IsEmptyCSV(filename : string) : boolean;

    //Client_ID
    function SameCompany(clientID, cname : string; numWords : integer) : boolean;
    function GenClientID(name,city : string) : string;
    function GenMethodA(words : TArray<string>) : string;
    function GenMethodB(words : TArray<string>) : string;
    function StripVowels(words : TArray<string>) : TArray<string>;
    function DoMethodB(words : TArray<string>; city : string) : string;
    function AddNumbers(clientID : string) : string;
    function isProfane(str : string) : boolean;
    function CreateProfanityFilter : TArray<string>;

    //testing
    procedure TestExp;
    procedure TestSFTP;
    procedure TestSFTPFileList;
    procedure TestCSVtoDS;
    procedure TestDeltaUpdate;
    procedure TestInsertUpdate;
    procedure TestClientIDGen;
    procedure TestPatch;
    procedure testCSVtoDSVis;
    procedure TestBlankCSV;
    procedure TestValidProv;
    procedure TestCleanupCSV;
    procedure TestDeltaCleanup;

    const headerstring = 'removed';

  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

uses dmod;


procedure TForm1.FormActivate(Sender: TObject);
var
  TF : string;
begin

  if Startup = false then
  begin
    pRetry.Visible := true;
  end;

end;


procedure TForm1.bRetryClick(Sender: TObject);
begin
  if Startup then
    pRetry.Visible := false;
end;


function TForm1.Startup : boolean;
var
  err, params : string;
  I : integer;
  sftp : boolean;
begin
  result := true;

  if LoadConfig then
  begin
    err := DBlogon;

    if err = 'success' then
    begin
      lDB.Caption := dbDatabase;

      if ParamCount = 1 then
      begin
        Application.Minimize;  //or maybe MainFrm.Visible := False;
        if paramstr(1) = 'auto' then
        begin
          try
            if RunDataExchange then
              LogMsg('[INFO] - Process "RunDataExchange" ran successfully')
            else
              LogMsg('[ERROR] - Process "RunDataExchange" FAILED');
          finally
            Application.Terminate;
          end;
        end
        else if paramstr(1) = 'extract' then
        begin
          try
            if Extract then
              LogMsg('[INFO] - Process "Extract" ran successfully')
            else
              LogMsg('[ERROR] - Process "Extract" FAILED');
          finally
            Application.Terminate;
          end;
        end
        else if paramstr(1) = 'upload' then
        begin
          try
            if InitSFTP then
            begin
              if Upload then
                LogMsg('[INFO] - Process "Upload" ran successfully')
              else
                LogMsg('[ERROR] - Process "Upload" FAILED');
              DestroySFTP;
            end
            else
              LogMsg('[ERROR] - Unable to connect to SFTP server. Process "Upload" aborted.');
          finally
            Application.Terminate;
          end;
        end
        else if paramstr(1) = 'download' then
        begin
          try
            if InitSFTP then
            begin
              if Download then
                LogMsg('[INFO] - Process "Download" ran successfully')
              else
                LogMsg('[ERROR] - Process "Download" FAILED');
              DestroySFTP;
            end
            else
              LogMsg('[ERROR] - Unable to connect to SFTP server. Process "Download" aborted.');
          finally
            Application.Terminate;
          end;
        end
        else if paramstr(1) = 'patch' then
        begin
          try
            if PatchClientIds then
              LogMsg('[INFO] - Process "PatchClientIds" ran successfully')
            else
              LogMsg('[ERROR] - Process "PatchClientIds" FAILED');
          finally
            Application.Terminate;
          end;
        end
        else if paramstr(1) = 'taxexempt' then
        begin
          try
            if TaxExemptReport then
              LogMsg('[INFO] - Process "TaxExemptReport" ran successfully')
            else
              LogMsg('[ERROR] - Process "TaxExemptReport" FAILED');
          finally
            Application.Terminate;
          end;
        end
        else
        begin
          LogMsg('[ERROR] - Invalid parameter string "' + paramstr(1) + '". Aborting...');
          Application.Terminate;
        end;
      end
      else if ParamCount = 2 then
      begin
        Application.Minimize;
        if ((paramstr(1) = 'extract') and (paramstr(2) = 'upload'))
        or ((paramstr(1) = 'upload') and (paramstr(2) = 'extract')) then
        begin
          try
            if Extract then
            begin
              LogMsg('[INFO] - Process "Extract" ran successfully');

              if InitSFTP then
              begin
                if Upload then
                  LogMsg('[INFO] - Process "Upload" ran successfully')
                else
                  LogMsg('[ERROR] - Process "Upload" FAILED');
                DestroySFTP;
              end
              else
                LogMsg('[ERROR] - Unable to connect to SFTP server. Process "Upload" aborted.');
            end
            else
              LogMsg('[ERROR] - Process "Extract" FAILED. Aborting process "Upload"...');
          finally
            Application.Terminate;
          end;
        end
        else if ((paramstr(1) = 'extract') and (paramstr(2) = 'download'))
             or ((paramstr(1) = 'download') and (paramstr(2) = 'extract')) then
        begin
          try
            if Extract then
              LogMsg('[INFO] - Process "Extract" ran successfully')
            else
              LogMsg('[ERROR] - Process "Extract" FAILED');

            if InitSFTP then
            begin
              if Download then
                LogMsg('[INFO] - Process "Download" ran successfully')
              else
                LogMsg('[ERROR] - Process "Download" FAILED');
              DestroySFTP;
            end
            else
              LogMsg('[ERROR] - Unable to connect to SFTP server. Process "Download" aborted.');

          finally
            Application.Terminate;
          end;
        end
        else if ((paramstr(1) = 'extract') and (paramstr(2) = 'patch'))
             or ((paramstr(1) = 'patch') and (paramstr(2) = 'extract')) then
        begin
          try
            if Extract then
              LogMsg('[INFO] - Process "Extract" ran successfully')
            else
              LogMsg('[ERROR] - Process "Extract" FAILED');

            if PatchClientIds then
              LogMsg('[INFO] - Process "PatchClientIds" ran successfully')
            else
              LogMsg('[ERROR] - Process "PatchClientIds" FAILED');
          finally
            Application.Terminate;
          end;
        end
        else if ((paramstr(1) = 'download') and (paramstr(2) = 'upload'))
             or ((paramstr(1) = 'upload') and (paramstr(2) = 'download')) then
        begin
          try
            if InitSFTP then
            begin
              if Upload then
                LogMsg('[INFO] - Process "Upload" ran successfully')
              else
                LogMsg('[ERROR] - Process "Upload" FAILED');
              if Download then
                LogMsg('[INFO] - Process "Download" ran successfully')
              else
                LogMsg('[ERROR] - Process "Download" FAILED');
              DestroySFTP;
            end
            else
              LogMsg('[ERROR] - Unable to connect to SFTP server. Process "Upload" and "Download" aborted.');
          finally
            Application.Terminate;
          end;
        end
        else if ((paramstr(1) = 'patch') and (paramstr(2) = 'upload'))
             or ((paramstr(1) = 'upload') and (paramstr(2) = 'patch')) then
        begin
          try
            if InitSFTP then
            begin
              if Upload then
                LogMsg('[INFO] - Process "Upload" ran successfully')
              else
                LogMsg('[ERROR] - Process "Upload" FAILED');
              DestroySFTP;
            end
            else
              LogMsg('[ERROR] - Unable to connect to SFTP server. Process "Upload" aborted.');

            if PatchClientIds then
              LogMsg('[INFO] - Process "PatchClientIds" ran successfully')
            else
              LogMsg('[ERROR] - Process "PatchClientIds" FAILED');
          finally
            Application.Terminate;
          end;
        end
        else if ((paramstr(1) = 'patch') and (paramstr(2) = 'download'))
             or ((paramstr(1) = 'download') and (paramstr(2) = 'patch')) then
        begin
          try
            if InitSFTP then
            begin
              if Download then
                LogMsg('[INFO] - Process "Download" ran successfully')
              else
                LogMsg('[ERROR] - Process "Download" FAILED');
              DestroySFTP;
            end
            else
              LogMsg('[ERROR] - Unable to connect to SFTP server. Process "Download" aborted.');

            if PatchClientIds then
              LogMsg('[INFO] - Process "PatchClientIds" ran successfully')
            else
              LogMsg('[ERROR] - Process "PatchClientIds" FAILED');
          finally
            Application.Terminate;
          end;
        end
        else
        begin
          LogMsg('[ERROR] - Invalid parameters "' + paramstr(1) + '" "' + paramstr(2) + '". Aborting...');
          Application.Terminate;
        end;

      end
      else if ParamCount = 3 then
      begin
        Application.Minimize;
        if  ((paramstr(1) = 'extract') or (paramstr(1) = 'upload') or (paramstr(1) = 'download'))
        and ((paramstr(2) = 'extract') or (paramstr(2) = 'upload') or (paramstr(2) = 'download'))
        and ((paramstr(3) = 'extract') or (paramstr(3) = 'upload') or (paramstr(3) = 'download')) then
        begin
          sftp := false;
          try
            if Extract then
            begin
              LogMsg('[INFO] - Process "Extract" ran successfully');
              if InitSFTP then
              begin
                sftp := true;
                if Upload then
                  LogMsg('[INFO] - Process "Upload" ran successfully')
                else
                  LogMsg('[ERROR] - Process "Upload" FAILED');
              end
              else
                LogMsg('[ERROR] - Unable to connect to SFTP server. Process "Upload" aborted.');
            end
            else
              LogMsg('[ERROR] - Process "Extract" FAILED. Aborting process "Upload"...');

            if not sftp then
              if InitSFTP then
                sftp := true;
            if sftp then
            begin
              if Download then
                LogMsg('[INFO] - Process "Download" ran successfully')
              else
                LogMsg('[ERROR] - Process "Download" FAILED');
            end
            else
              LogMsg('[ERROR] - Unable to connect to SFTP server. Process "Download" aborted.');
            if sftp then
              DestroySFTP;
          finally
            Application.Terminate;
          end;
        end
        else if  ((paramstr(1) = 'extract') or (paramstr(1) = 'upload') or (paramstr(1) = 'patch'))
             and ((paramstr(2) = 'extract') or (paramstr(2) = 'upload') or (paramstr(2) = 'patch'))
             and ((paramstr(3) = 'extract') or (paramstr(3) = 'upload') or (paramstr(3) = 'patch')) then
        begin
          try
            if Extract then
            begin
              LogMsg('[INFO] - Process "Extract" ran successfully');
              if InitSFTP then
              begin
                sftp := true;
                if Upload then
                  LogMsg('[INFO] - Process "Upload" ran successfully')
                else
                  LogMsg('[ERROR] - Process "Upload" FAILED');
                DestroySFTP;
              end
              else
                LogMsg('[ERROR] - Unable to connect to SFTP server. Process "Upload" aborted.');
            end
            else
              LogMsg('[ERROR] - Process "Extract" FAILED. Aborting process "Upload"...');

            if PatchClientIds then
              LogMsg('[INFO] - Process "PatchClientIds" ran successfully')
            else
              LogMsg('[ERROR] - Process "PatchClientIds" FAILED');
          finally
            Application.Terminate;
          end;
        end
         else if  ((paramstr(1) = 'extract') or (paramstr(1) = 'patch') or (paramstr(1) = 'download'))
             and ((paramstr(2) = 'extract') or (paramstr(2) = 'patch') or (paramstr(2) = 'download'))
             and ((paramstr(3) = 'extract') or (paramstr(3) = 'patch') or (paramstr(3) = 'download')) then
        begin
          try
            if Extract then
              LogMsg('[INFO] - Process "Extract" ran successfully')
            else
              LogMsg('[ERROR] - Process "Extract" FAILED');

            if InitSFTP then
            begin
              if Download then
                LogMsg('[INFO] - Process "Download" ran successfully')
              else
                LogMsg('[ERROR] - Process "Download" FAILED');
              DestroySFTP;
            end
            else
              LogMsg('[ERROR] - Unable to connect to SFTP server. Process "Download" aborted.');

            if PatchClientIds then
              LogMsg('[INFO] - Process "PatchClientIds" ran successfully')
            else
              LogMsg('[ERROR] - Process "PatchClientIds" FAILED');
          finally
            Application.Terminate;
          end;
        end
         else if  ((paramstr(1) = 'patch') or (paramstr(1) = 'upload') or (paramstr(1) = 'download'))
             and ((paramstr(2) = 'patch') or (paramstr(2) = 'upload') or (paramstr(2) = 'download'))
             and ((paramstr(3) = 'patch') or (paramstr(3) = 'upload') or (paramstr(3) = 'download')) then
        begin
          try
            if InitSFTP then
            begin
              if Upload then
                LogMsg('[INFO] - Process "Upload" ran successfully')
              else
                LogMsg('[ERROR] - Process "Upload" FAILED');
              if Download then
                LogMsg('[INFO] - Process "Download" ran successfully')
              else
                LogMsg('[ERROR] - Process "Download" FAILED');
              DestroySFTP;
            end
            else
              LogMsg('[ERROR] - Unable to connect to SFTP server. Process "Upload" and "Download" aborted.');

            if PatchClientIds then
              LogMsg('[INFO] - Process "PatchClientIds" ran successfully')
            else
              LogMsg('[ERROR] - Process "PatchClientIds" FAILED');
          finally
            Application.Terminate;
          end;
        end
        else
        begin
          LogMsg('[ERROR] - Invalid parameters "' + paramstr(1) + '" "' + paramstr(2) + '" "' + paramstr(3) + '". Aborting...');
          Application.Terminate;
        end;

      end
      else if ParamCount = 4 then
      begin
        Application.Minimize;
        try
          if RunDataExchange then
            LogMsg('[INFO] - Process "RunDataExchange" ran successfully')
          else
            LogMsg('[ERROR] - Process "RunDataExchange" FAILED');
        finally
          Application.Terminate;
        end;
      end;

    end
    else
    begin
      if ParamCount > 0 then
      begin
        LogMsg('[ERROR] - Aborting - ' + err);
        EmailError('HRCxTM Fatal Error','Aborted - ' + err,ITRecips);
        Application.Terminate;
      end
      else
      begin
        showMessage('A connection to the database was unable to be established: ' + err + '.' + slinebreak + slinebreak +
                    'If you are sure the database is not the source of this issue and the config value [DB]-database is correct, ' +
                    'then try clicking "Configure Database" and entering login information to retry.');
      end;
      result := false;
    end;

  end
  else
  begin
    if ParamCount > 0 then
    begin
      EmailError('HRCxTM Fatal Error','Unable to load configuration from file. Process aborted',ITRecips);
      Application.Terminate;
    end
    else
    begin
      showMessage('Unable to load configuration from file. Please look for "HRCxTM.ini" in ' + GenIniLoc + ' to determine the cause');
    end;
    result := false;
  end;
end;


function TForm1.RunDataExchange : boolean;
var
  upld,dnld : boolean;
begin
  //init SFTP
  LogMsg('[INFO] - Begin "RunDataExchange"');
  result := true; upld := true; dnld := true;
  if initSFTP then
  begin

    if Extract then
    begin
      if not Upload then
        upld := false;
    end
    else
      upld := false;

    if not Download then
      dnld := false;

    if (upld = false) and (dnld = false) then
    begin
      LogMsg('[ERROR] - All outbound and inbound extracts failed.');
      EmailError('HRCxTM Fatal Error','All outbound and inbound extracts failed.',ITRecips);
      result := false;
    end;

    destroySftp;
  end
  else
  begin
    LogMsg('[ERROR] - SFTP connect failed. Process aborted');
    EmailError('HRCxTM Fatal Error','SFTP connection failed. Process aborted. Check the log at ' + logPath,ITRecips);
    result := false;
  end;
  //Run Client ID Manual Patch procedure
  PatchClientIDs;
end;


function TForm1.Extract : boolean;
var
  err,dateStr, expFName, uplFName, expFolder, activeFolder : string;
  I,failCt : integer;
  extractData : TArray<TArray<string>>;
  files : TArray<string>;
begin
  LogMsg('[INFO] - Begin "Extract"');
  result := true;

  try
  // Get files from outbound directory
    files := TDirectory.GetFiles(appPath + 'Outbound');

    if length(files) > 0 then
    begin
      LogMsg('[INFO] - Clearing directory ' + appPath + 'Outbound\');
      for I := 0 to length(files) - 1 do
        DelFile(files[I]);
    end;
  except
    on E:Exception do
    begin
      LogMsg('[ERROR] - Unable to fully clear directory ' + appPath + 'Outbound\.');
    end;
  end;

  //load cryRptsDotNet.dll
  err := CryUtilsNonVCL.LoadDLL(ModMgr,dllPath);
  if err <> '' then
    begin
    LogMsg('[ERROR] - Not running exports - Crystal Reports error:' + err);
    EmailError('HRCxTM Error','Did not run exports. Error: ' + err,ITRecips);
    result := false;
  end
  else
  begin
    failCt := 0;

    //Do stored procedure for updating AR aging and yearly sales values
    err := DataMod.CalcAging;
    if err <> '' then
    begin
      LogMsg('[ERROR] - ' + err);
      EmailError('HRCxTM Error',err,ITRecips);
    end
    else
      LogMsg('[INFO] - CalcAging ran successfully');

    //Delete files from Raw folder

    //for closedAR, openAR and Customer extracts
    //load report, export to csv, and sftp upload
    //if there are any errors dont do consecutive steps, log and potentially notify someone
    System.SysUtils.FormatSettings.DateSeparator := '_';System.SysUtils.FormatSettings.TimeSeparator := '_';

    expFolder := 'Outbound\Raw\'; activeFolder := 'Outbound\';
    //0 is customer master, 1 is open AR, 2 is closed AR
    extractData := [['Customer Master Extract','customer_master_','CustomerExtract.rpt',sftpCustMstPath],
                    ['Open AR Extract','open_ar_','OpenARExtract.rpt',sftpArPath],
                    ['Closed AR Extract','closed_ar_','ClosedARExtract.rpt',sftpArPath]];

    for I := 0 to 2 do
    begin
      //First try to clean out unneccessary entries in the delta table and log any errors but continue
      err := DataMod.CleanDeltaTable;
      if err <> '' then
      begin
        LogMsg(err);
        err := '';
      end;

      //Clear out 1 month old processed entries in the AR delta table, and 2 month old processed entries in the client delta table
      //should this be a config value?
      err := DataMod.ClearOldDelta;
      if err <> '' then
      begin
        if ContainsText(err,'[SEP]') then
        begin
          LogMsg( err.Split(['[SEP]'])[0] );
          LogMsg( err.Split(['[SEP]'])[1] );
        end
        else
          LogMsg(err);
        err := '';
      end;

      //Mark delta entries as active or about to processed
      if I = 0 then
        err := DataMod.CustDeltaActive
      else if I = 1 then
        err := DataMod.OARDeltaActive
      else if I = 2 then
        err := DataMod.CARDeltaActive;

      if err = '' then
      begin
        DateTimeToString(dateStr,'dd/mm/yy/hh/mm/ss',Now);
        //expFName := appPath + extractData[I][1] + dateStr + '.csv';
        expFName := appPath + expFolder + extractData[I][1] + dateStr + '_raw.csv';
        if ExpCsv(rptsPath + extractData[I][2],expFName) then
        begin
          uplFName := CleanupCSV(expFName);
          if uplFName <> '' then
          begin
            MoveProcessedFile(uplFName,activeFolder);
            LogMsg('[INFO] - Created extract: ' + appPath + activeFolder + extractFileName(uplFName));
          end
          else
          begin
            MoveProcessedFile(expFName,'Outbound\Error\');
            inc(failCt);
          end;
        end
        else
        begin
          LogMsg('[ERROR] - Failed to export ' + extractData[I][2] + '. Not sent to HRC');
          inc(failCt);
        end;

      end
      else
      begin
        LogMsg('[ERROR] - Failed to set delta records to active for extract "' + extractData[I][0] + '". ' + err);
        inc(failCt);
      end;
    end;

    if failCt = 3 then
    begin
      LogMsg('[ERROR] - None of the extracts were generated.');
      EmailError('HRCxTM Error','None of the extracts were generated. Check the log file at: ' + logPath,ITRecips);
      result := false;
    end
    else if failCt > 0 then
    begin
      LogMsg('[WARNING] - Process "Extract" ran and generated ' + inttostr(3 - failCt) + ' out of 3 extracts.');
    end;


    //Cleanup Crystal
    err := CryUtilsNonVCL.ReleasePlugin(modMgr,cryPlugin);
    if err <> '' then
      LogMsg('[ERROR] - ReleasePlugin(modMgr,cryPlugin): ' + err);

    err := CryUtilsNonVCL.UnloadDLL(modMgr);
    if err <> '' then
      LogMsg('[ERROR] - UnloadDLL(modMgr): ' + err);

    System.SysUtils.FormatSettings.DateSeparator := '-';System.SysUtils.FormatSettings.TimeSeparator := ':';
  end;

end;


//YOU MUST USE InitSFTP BEFORE USING THIS FUNCTION!!!
function TForm1.Upload : boolean;
var
  err,fileName : string;
  I,extractType,failCt : integer;
  extractData : TArray<TArray<string>>;
  files : TArray<string>;
  currTime, sixAM : TDateTime;
begin
  result := true;
  failCt := 0;

  LogMsg('[INFO] - Begin "Upload"');

  try
  // Get files from outbound directory
    files := TDirectory.GetFiles(appPath + 'Outbound');

    if length(files) > 0 then
    begin

      //0 is customer master, 1 is open AR, 2 is closed AR
      extractData := [['Customer Master Extract','customer_master_',sftpCustMstPath],
                      ['Open AR Extract','open_ar_',sftpArPath],
                      ['Closed AR Extract','closed_ar_',sftpArPath]];

      for I := 0 to length(files) - 1 do
      begin

        //identify extract type. 0 is Customer Master, 1 is Open AR, 2 is Closed AR
        fileName := extractFileName(files[I]);
        if ContainsText(fileName,extractData[0][1]) then
          extractType := 0
        else if ContainsText(fileName,extractData[1][1]) then
          extractType := 1
        else if ContainsText(fileName,extractData[2][1]) then
          extractType := 2
        else
          extractType := -1;

        if extractType <> -1 then
        begin

          if UploadSFTP(files[I],extractData[extractType][2] + fileName) then
          begin
            //if successful update table so that we dont try to do it again for these values next time
            if extractType = 0 then
              err := DataMod.UpdCustDelta
            else if extractType = 1 then
              err := DataMod.UpdOARDelta
            else if extractType = 2 then
              err := DataMod.UpdCARDelta;

            if err = '' then
            begin
              MoveProcessedFile(files[I], 'Outbound\Processed\');
              LogMsg('[INFO] - Uploaded extract: ' + files[I]);
            end
            else
            begin
              LogMsg(err);
              MoveProcessedFile(files[I],'Outbound\Error\');
              EmailError('HRCxTM Error','Unable to update delta table values for ' + extractData[extractType][0] + ': ' + err,ITRecips);
              inc(failCt);
            end;
          end
          else
          begin
            LogMsg('[ERROR] - Failed to SFTP upload file ' + files[I]);
            MoveProcessedFile(files[I],'Outbound\Error\');
            inc(failCt);
          end;
        end
        else
        begin
          LogMsg('[ERROR] - File ' + files[I] + ' is not a recognized extract file. Moving to Error directory.');
          MoveProcessedFile(files[I],'Outbound\Error\');
          inc(failCt);
        end;

      end;

      if failCt = length(files) then
      begin
        LogMsg('[ERROR] - None of the extracts were uploaded to HRC');
        EmailError('HRCxTM Error','None of the extracts were uploaded to HRC. Check the log file at: ' + logPath,ITRecips);
        result := false;
      end
      else if failCt > 0 then
      begin
        LogMsg('[WARNING]Process "Upload" ran and uploaded ' + inttostr(length(files) - failCt) +
               ' files to SFTP out of ' + inttostr(length(files)) + ' files in directory ' + appPath + 'Outbound');
      end
      else
      begin
        currTime := Time;
        sixAM := EncodeTime(7,15,0,0);     //Apparently the cron is at 9:15 CT, so 7:15 PT
        if TimeOf(currTime) > sixAM then
          EmailError('HRCxTM alert','Process "upload" finished successfully but completed after 7:15 AM PT. ' +
                     'Please check SFTP server to ensure extracts were processed and they may have been missed',ITRecips);
      end;

    end
    else
    begin
      LogMsg('[WARNING] - There were no extracts to upload');
      result := false;
    end;

  except
    on E:Exception do
    begin
      logMsg('[ERROR] - Failed loading data from directory ' + appPath + 'Outbound . Message: ' + E.Message);
      result := false;
    end;
  end;

end;


//YOU MUST USE InitSFTP BEFORE USING THIS FUNCTION!!!
function TForm1.Download : boolean;
var
  fName, downloadDir, proc : string;
  procType : boolean;
  I,J,failCt,numFiles : integer;
  //dirs : TArray<string>;
begin
  result := true;

  LogMsg('[INFO] - Begin "Download"');

  //try to download SFTP the customer create and update files
  GetFileList([sftpCustCrtPath,sftpCustUpdPath]);

  for I := 0 to 1 do
  begin
    if I = 0 then
      downloadDir := sftpCustCrtPath
    else if I = 1 then
      downloadDir := sftpCustUpdPath;

    for J := 0 to length(sftpFileList[I]) - 1 do
    begin
      fName := appPath + 'Inbound\' + StringReplace(sftpFileList[I][J],downloadDir,'',[]);
      if (DnloadSFTP(sftpFileList[I][J],fName)) then
      begin
        SetLength(fileList[I],length(fileList[I]) + 1);
        fileList[I][J] := fName;
      end
      else
      begin
        LogMsg('[ERROR] - Failed to download ' + fileList[I][J] + ' from HRC');
        sftpFileList[I][J] := '';
      end;
    end;
  end;

  failCt := Process;

  numFiles := length(fileList[0]) + length(fileList[1]);
  if numFiles = 0 then
  begin
    LogMsg('[WARNING] - There were no files to download');
    EmailError('HRCxTM Alert','There were no inbound extracts downloaded. Please check manually to verify this is not a mistake.',ITRecips{????});
    result := false;
  end
  else if failCt = numFiles then
  begin
    LogMsg('[ERROR] - All inbound extracts failed to be processed');
    //send out an email to alert someone
    EmailError('HRCxTM Error','All inbound extracts failed to be processed. Check the log at ' + logPath,ITRecips);
    result := false;
  end
  else if failCt > 0  then
    LogMsg('[WARNING] - Process "Download" downloaded ' + inttostr(numFiles) + ' out of ' + inttostr(length(sftpFileList[0]) + length(sftpFileList[1])) + ' files on SFTP server' +
           ' and processed ' + inttostr(numFiles - failCt) + ' out of ' + inttostr(numFiles) + ' files.')
  else if failCt = -1 then
    LogMsg('[INFO] - Did not detect valid files to process in inbound directory.');

  sftpFileList := nil;
  fileList := nil;
end;


function TForm1.Process : integer;
var
  I,J : integer;
  procType, areFiles : boolean;
  downloadDir, proc : string;
begin
  areFiles := false;
  result := 0;
  SetLength(NewIDSummary,0);
  if fileList = nil then  //generate list from local dir
  begin
    setlength(fileList,2);
    fileList[0] := TDirectory.GetFiles(appPath + 'Inbound\','New_Customer_Creation_*',TSearchOption.soTopDirectoryOnly);
    fileList[1] := TDirectory.GetFiles(appPath + 'Inbound\','Output_Existing_Customer_Update_*',TSearchOption.soTopDirectoryOnly);
  end;
  for var N := 0 to 1 do if areFiles = false then
    for var M := 0 to length(fileList[0]) - 1 do
      if filelist[N][M] <> '' then
        if fileExists(filelist[N][M]) then
        begin
          areFiles := true;
          break;
        end
        else
          LogMsg('[WARNING] - File "' + filelist[N][M] +'" does not exist.');

  if areFiles then
  begin
    for I := 0 to 1 do   //Loop through file names and process them all
    begin
      for j := 0 to length(fileList[I]) - 1 do
      begin
        if I = 0 then
        begin
          procType := true;
          downloadDir := sftpCustCrtPath;
        end
        else if I = 1 then
        begin
          procType := false;
          downloadDir := sftpCustUpdPath;
        end;

        proc := ProcessCustExtract(fileList[I][J],procType);
        if containsStr(proc,'error') then
          inc(result)
        else
        begin
          if proc <> '' then
            LogMsg('[INFO] - Processed extract: ' + fileList[I][J]);
        end;
        MoveProcessedFile(fileList[I][J], 'Inbound\' + proc + '\');
        if sftpFileList <> nil then
          MoveSFTP(sftpFileList[I][J],downloadDir + stringreplace(proc,'\empty','',[]) + '/' + StringReplace(sftpFileList[I][J],downloadDir,'',[]));
      end;
    end;
    {$IFDEF Release}
    if length(NewIDSummary) > 0 then
    begin
      var body := 'HRCxTM ran at ' + timetostr(timeOf(now)) + ' and created ' + inttostr(length(NewIDSummary)) + ' new clients in Truckmate.'
                + slinebreak + slinebreak + slinebreak + 'Client Name/Legal Name | Client ID generated' + slinebreak;
      for var K := 0 to length(NewIDSummary)-1 do
        body := body + slinebreak + NewIDSummary[K][0] + ' | ' + NewIDSummary[K][1];
      EmailError('HRCxTM Client Creation Summary ' + datetimetostr(now),body,ITRecips + AcctRecips);
    end;
    {$ENDIF}
  end
  else
    result := -1;
end;


function TForm1.PatchClientIDs : boolean;
var
  files : TArray<string>;
  I : integer;
  fileName,ID : string;
begin
  result := true;

  LogMsg('[INFO] - Begin "PatchClientIDs"');
  //Look in csvPath + 'Patch\' directory and create a list of all files
  try
    files := TDirectory.GetFiles(appPath + 'Patch');

    //loop through all files
    for I := 0 to length(files) - 1 do
    begin
      //store filename and strip path
      fileName := ExtractFileName(files[I]);
      //if filename = 'clientData.txt' then ignore
      if fileName <> 'clientData.txt' then
      begin
        ID := ReplaceStr(fileName,'.txt','');

        if length(ID) > 10 then
        begin
          //send email to acct to use different ID
          LogMsg('[ERROR] - Failed creating a client that required a manually entered CLIENT_ID. CLIENT_ID: ' + ID);
          EmailError('HRCxTM Error: Action Required','Failed creating a client that required a manually entered CLIENT_ID. The supplied CLIENT_ID is: ' + ID + slinebreak +
                     'This CLIENT_ID is too long as it is greater than 10 characters. Please rename the file to a different CLIENT_ID. The file is located at ' + files[I],acctRecips);
          result := false;
          continue
        end;

        if DataMod.ExistsClientID(ID) then
        begin
          //send email to acct to use different ID
          LogMsg('[ERROR] - Failed creating a client that required a manually entered CLIENT_ID - already exists. CLIENT_ID: ' + ID);
          EmailError('HRCxTM Error: Action Required','Failed creating a client that required a manually entered CLIENT_ID. The supplied CLIENT_ID is: ' + ID + slinebreak +
                     'This CLIENT_ID is already in use. Please rename the file to a different CLIENT_ID. The file is located at ' + files[I],acctRecips);
          result := false;
          continue
        end;

        DataMod.CSVtoDS(files[I],inboundSep,inboundDelim,true);

        if DataMod.custData.RecordCount <> 0 then
        begin
          //do insert
          DataMod.custData.First;
          if not DataMod.DoClientIns(ID) then
          begin
            //Send email to IT to determine cause
            LogMsg('[ERROR] - Process "Patch" failed to insert a client record with filename: ' + fileName);
            EmailError('HRCxTM Error','Process "Patch" failed to insert a client record with filename: ' + fileName + slinebreak +
                       '. Please check logs to determine cause.',ITRecips);
            result := false;
            continue
          end
          else
          begin
            MoveProcessedFile(files[I], 'Patch\Processed\');
            LogMsg('[INFO] - Client created with CLIENT_ID = ' + ID);
            //DelFile(files[I]);
          end;
        end
        else
        begin
          //send email to acct to fix formatting
          LogMsg('[ERROR] - Failed creating a client that required a manually entered CLIENT_ID. CLIENT_IS: ' + ID);
          EmailError('HRCxTM Error: Action Required','Failed creating a client that required a manually entered CLIENT_ID. The supplied CLIENT_IS is: ' + ID + slinebreak +
                     'Please check the files contents as there is likely a formatting error. The file is located at ' + files[I], acctRecips);
          result := false;
          continue
        end;
      end
      else
      begin
        //send email to acct to fix formatting
        LogMsg('[ERROR] - Filename not changed for file in Patch folder.');
        EmailError('HRCxTM Error: Action Required','Failed creating a client that required a manually entered CLIENT_ID. The file name was not changed to a CLIENT_ID.' +
        ' The file is located at ' + files[I], acctRecips);
        result := false;
        continue
      end;
    end;
  except
    on E:Exception do
    begin
      logMsg('[ERROR] - Failed loading data from directory ' + appPath + 'Patch . Message: ' + E.Message);
      result := false;
    end;
  end;

end;


function TForm1.TaxExemptReport : boolean;
begin
  result := true;
  LogMsg('[INFO] - Begin "TaxExemptReport"');
  try
    if DataMod.qTaxExRpt.Active then
      DataMod.qTaxExRpt.Close;
    DataMod.qTaxExRpt.Open;
    if DataMod.qTaxExRpt.RecordCount > 0 then
    begin
      try
        var tableHTML := '<table style="border:1px solid;border-collapse:collapse">';
        var style := ' style="border:1px solid;padding:5px">';
        tableHTML := tableHTML + '<thead><th'+style+'CLIENT_ID</th><th'+style+'NAME</th><th'+style+'CBTAX_1</th><th'+style+'HRC_ID</th><th'+style+'CREATED</th></thead>';
        for var I := 1 to DataMod.qTaxExRpt.RecordCount do
        begin
          tableHtml := tableHtml+'<tr><td'+style+DataMod.qTaxExRpt.fieldbyName('CLIENT_ID').asString+'</td>'+
                                 '<td'+style+DataMod.qTaxExRpt.fieldbyName('NAME').asString+'</td>'+
                                 '<td'+style+DataMod.qTaxExRpt.fieldbyName('CBTAX_1').asString+'</td>'+
                                 '<td'+style+DataMod.qTaxExRpt.fieldbyName('HRC_ID').asString+'</td>'+
                                 '<td'+style+DataMod.qTaxExRpt.fieldbyName('CREATED').asString+'</td></tr>';
          DataMod.qTaxExRpt.Next;
        end;
        tableHtml := tableHtml + '</table>';
        var mailBody := 'Below are all Clients created by HighRadius, marked as Tax Exempt in Truckmate, in the last 2 weeks (from today''s date).' + slinebreak + tableHtml;
        emailError('HighRadius Tax Exempt Client Report',mailBody,taxExRecips,nil,TaxExCCS);
      except on E:Exception do
        begin
          LogMsg('[ERROR] - Failed creating report table for email. Aborting send. Message: ' + E.Message);
          result := false;
        end;
      end;
    end
    else
      LogMsg('[Warning] - No records returned from "qTaxExRpt"');
  except on E:Exception do
    begin
      LogMsg('[ERROR] - Failed running "qTaxExRpt". Message: ' + E.Message);
      result := false;
    end;
  end;
end;


procedure Tform1.MoveProcessedFile(name,newPath : string; n : integer = 0);
begin
  try
    TFile.Move(name, appPath + newPath + extractFileName(name));
  except
    on E: Exception do
    begin
      LogMsg('[ERROR] - Attempting to move file "' + name + '" to ' + appPath + newPath + ':' + E.Message);
      if n < numRetries then
      begin
        Sleep(waitLength);
        LogMsg('[INFO] - Retrying move...');
        MoveProcessedFile(name,newPath,n + 1);
      end;
    end;
  end;
end;


function TForm1.ProcessCustExtract(fileName : string; insert : boolean) : string;
var
  I,errCt : integer;
  failedRecs,kind, ID, NAME, hrcID, typestr : string;
  attach : TArray<string>;
  procedure AddToSummary(id,name : string);
  begin
    var len := length(NewIDSummary);
    SetLength(NewIDSummary,len + 1);
    SetLength(NewIDSummary[len],2);
    NewIDSummary[len][0] := name;
    NewIDSummary[len][1] := id;
  end;
begin
  result := 'processed';
  errCt := 0;
  if insert then
    failedRecs := headerstring + slinebreak
  else
    failedRecs := updheaderstr + slinebreak;
  if containsstr(filename,'.csv') then
  begin
    if not isEmptyCSV(fileName) then
    begin
      try
        if insert then
          DataMod.CSVtoDS(fileName,inboundSep,inboundDelim,insert)  //is the move issue caused here?
        else
          DataMod.CSVtoDS(fileName,inboundSep,inboundDelim,insert);

        if DataMod.custData.RecordCount <> 0 then
        begin
          DataMod.custData.First;
          for I := 0 to DataMod.custData.RecordCount - 1 do
          begin
            if insert then
            begin
              try
                ID := '';
                NAME := DataMod.custData.FieldByName('NAME').AsString;
                if NAME = '' then
                  NAME := DataMod.custData.FieldByName('LEGAL_NAME').AsString;
                ID := GenClientID(NAME,DataMod.custData.FieldByName('CITY').AsString);
                if ID <> '' then
                begin
                  DataMod.DB.StartTransaction;
                  if DataMod.DoClientIns(ID) then
                  begin
                    DataMod.DB.Commit;
                    LogMsg('[INFO] - Created client "' +NAME + ' (' + ID + ')"');
                    AddToSummary(ID,NAME);
                  end
                  else
                  begin
                    DataMod.DB.Rollback;
                    failedRecs := failedRecs + DataMod.GetRecordCSV(inboundDelim,inboundSep,'ins',ID);
                    inc(errCt);
                  end
                end
                else
                begin
                  hrcID := DataMod.custData.FieldByName('CDF.HRC_ID').AsString;
                  if hrcID <> '' then
                  begin
                    //create file to attach at csvPath + 'Patch\clientData.txt'
                    LogMsg('[ERROR] - Unable to generate CLIENT_ID for client with HRC_ID: ' + hrcID);
                    EmailPatch(DataMod.GetRecordCSV(inboundDelim,inboundSep,'ins',ID),'Unable to generate CLIENT_ID for client with HRC_ID: ' + hrcID + slineBreak +
                               'The data for the client is attached to this email and called "clientData.txt". ' +
                               'Please review the client and rename the file to the desired CLIENT_ID + ".txt",' +
                               'then copy it to the folder ' + appPath + 'Patch\ on BTSSMATOOLS1.');
                  end
                  else
                  begin
                    LogMsg('[ERROR] - Unable to generate CLIENT_ID for client named: ' + NAME +'. No HRC_ID.');
                    EmailError('HRCxTM Error: Action Required','Unable to generate CLIENT_ID for client named: ' + NAME +
                    '. There is also no HRC_ID present, please contact HRC.',ITRecips);
                  end;
                  failedRecs := failedRecs + DataMod.GetRecordCSV(inboundDelim,inboundSep,'ins',ID);
                  inc(errCt);
                end;
              except
                on E: Exception do
                begin
                  //What does it mean for this to be a failure - all records must fail
                  //but we want to count all failures and save them into a csv
                  failedRecs := failedRecs + DataMod.GetRecordCSV(inboundDelim,inboundSep,'ins',ID);
                  inc(errCt);
                end;
              end;
            end
            else
            begin
              try
                DataMod.DB.StartTransaction;
                if DataMod.DoClientUpd then
                  DataMod.DB.Commit
                else
                begin
                  DataMod.DB.Rollback;
                  failedRecs := failedRecs + DataMod.GetRecordCSV(inboundDelim,inboundSep,'upd',ID);
                  inc(errCt);
                end
              except
                on E: Exception do
                begin
                  logMsg('[ERROR] - DoClientUpd: ' + E.Message);
                  inc(errCt);
                end;
              end;
            end;

            if not DataMod.custData.Eof then
              DataMod.custData.Next;
          end;

          if errCt = DataMod.custData.RecordCount then
          begin
            result := 'error';
            LogMsg('[ERROR] - Failed to insert all records for file: "' + filename + '"');
          end
          else if errCt > 0 then
          begin
            //Send email with failed records
            if insert then kind := 'insert' else kind := 'update';
            LogMsg('[ERROR] - ' + inttostr(errCt)+' records failed to '+kind+'. Dumping below:'+slineBreak+failedRecs);
            EmailError('HRCxTM Error', inttostr(errCt)+' records failed to '+kind+'.If appropriate please advise accounting. They are listed below.'
            +slinebreak+slinebreak+failedRecs,ITRecips);
          end;
        end
        else
        begin
          result := 'error';
          LogMsg('[ERROR] - No records were able to be created from csv file: "' + filename + '"');
          EmailError('HRCxTM Alert','No records were able to be created from csv file: "' + filename + '"' + slinebreak +
                     'Check the error folder and logs. This may be due to a formatting error.',ITRecips);
        end;

      except
        on E: Exception do
        begin
          result := 'error';
          LogMsg('[ERROR] - Failed processing "' + fileName + '" into a dataset: ' + E.Message);
          EmailError('HRCxTM Alert','Failed processing "' + fileName + '" into a dataset: ' + E.Message + slinebreak +
                     'Check the error folder and logs. This may be due to a formatting error.',ITRecips);
        end;
      end;
    end
    else
    begin
      result := 'processed\empty';
      LogMsg('[WARNING] - Extract "' + fileName + '" is empty. No changes to the database were performed.');
      //Do we email?
      //if send on empty then send email
      if sendOnEmpty then
      begin
        if insert then
          typestr := 'New Customer Creation'
        else
          typestr := 'Existing Customer Update';

        EmailError('HRCxTM Alert','While running the HighRadius integration, the ' + typestr + ' file was empty.' + slinebreak +
                   'If you were expecting a(n) Update/New Client, contact IT or HRC. If there were no updates expected, please ingore this message. ' +
                   'If you would no longer like to recieve these messages, contact IT to change the config value "sendOnEmptyExtract" to "0".',acctRecips);
      end;
    end;
  end
  else
  begin
    result := 'error';
    var errormsg := '[WARNING] - Did not process inbound extract "' + filename + '": Invalid file type - ';
    var filestr : string;
    if containsStr(filename,'\') then
    begin
      var tokens := filename.Split(['\']);
      filestr := tokens[length(tokens)-1];
    end
    else
      filestr := fileName;
    if containsstr(filestr,'.') then
    begin
      var fileExt := '.' + filestr.Split(['.'])[1];
      errormsg := errormsg + '"' + fileExt + '"';
    end
    else
      errormsg := errormsg + 'No file extension';
    LogMsg(errormsg);
    EmailError('HRCxTM Alert',errormsg + slinebreak + 'Check the error folder and logs.',ITRecips);
  end;
end;


{$REGION 'Config and Logging'}


function TForm1.DBlogon : string;
var
  dbConfig : TInifile;
  nulls{, dbIni} : string;
begin
  result := '';
  nulls := '';
  dbConnArr := GetLoginInfoODBC(dbDatabase);
  if length(dbConnArr) > 1 then
  begin
    if dbConnArr[0] = '' then
      nulls := 'username';
    if dbConnArr[1] = '' then
      nulls := ', password';
    if nulls <> '' then
    begin
      result := 'Null value read for database ' + nulls + '.';
      exit
    end;
  end
  else
  begin
    if length(dbConnArr) = 1 then result := dbConnArr[0];
    exit;
  end;
  odbcName := dbConnArr[0];
  result := DataMod.initDB(dbDatabase,dbConnArr[1],decPass(dbConnArr[2]),dbConnArr[3]);
end;


function TForm1.DoDBConfig : boolean;
var
  DBForm : TDBConfigAdminForm;
  modRes : integer;
  config : TIniFile;
begin
  result := false;
  try
    DBForm := TDBConfigAdminForm.Create(nil);
    DBForm.init(false);
    modRes := DBForm.ShowModal;

    if modRes = mrYes then
    begin
      result := True;
      try
        config := TInifile.Create(GenIniLoc);
      except
        on E : Exception do
        begin
          //email IT instead
          ShowMessage('Unable to open configuration file: ' + E.Message);
          result := false;
          exit
        end;
      end;
      dbDatabase := config.ReadString('DB','database','');
      if dbDatabase = '' then
      begin
        ShowMessage('Config value DB-database is blank. Try again');
        result := false;
      end
      else
      begin
        var ret := DBlogon;
        if ret <> 'success' then
        begin
          result := false;
          ShowMessage(ret);
        end;
      end;
      config.Free;
    end;
  finally
    DBForm.Release;
  end;
end;


function TForm1.LoadConfig : boolean;
var
  config : TInifile;
  iniFileLoc, temp, temp2 : string;
  tempInt : integer;
  tempArr : TArray<string>;
  pchar : PWideChar;
  binBytes : TBytes;
begin
  result := true;
  iniFileLoc := GenIniLoc;
  if iniFileLoc <> '' then
  begin
    if fileExists(iniFileLoc) then
    begin
      try
        config := TInifile.Create(iniFileLoc);
      except
        on E : Exception do
        begin
          //email IT instead
          EmailError('HRCxTM error: Action required','Unable to open configuration file: ' + E.Message,ITRecips);
          result := false;
          exit
        end;
      end;
      //read saved values
      //Database
      temp := config.ReadString('DB','database','');
      if temp <> '' then
      begin
        dbDatabase := temp;
      end
      else begin LogMsg('[ERROR] - Could not read config value DB-database'); result := false; exit end;

      //SFTP
      temp := config.ReadString('SFTP','host',''); if temp <> '' then sftpHost := temp
      else begin LogMsg('[ERROR] - Could not read config value SFTP-host'); result := false; exit end;
      tempInt := config.ReadInteger('SFTP','port',0); if tempInt <> -1 then sftpPort := tempInt
      else begin LogMsg('[ERROR] - Could not read config value SFTP-port'); result := false; exit end;
      temp := config.ReadString('SFTP','user',''); if temp <> '' then sftpUser := temp
      else begin LogMsg('[ERROR] - Could not read config value SFTP-user'); result := false; exit end;
      temp := config.ReadString('SFTP','privKeyName',''); if temp <> '' then sftpPrivKeyName := temp
      else begin LogMsg('[ERROR] - Could not read config value SFTP-privKeyName'); result := false; exit end;
      temp := config.ReadString('SFTP','arPath',''); if temp <> '' then sftpArPath := temp
      else begin LogMsg('[ERROR] - Could not read config value SFTP-arPath'); result := false; exit end;

      temp := config.ReadString('SFTP','customerMasterPath',''); if temp <> '' then sftpCustMstPath := temp
      else begin LogMsg('[ERROR] - Could not read config value SFTP-customerMasterPath'); result := false; exit end;
      temp := config.ReadString('SFTP','custCreatePath',''); if temp <> '' then sftpCustCrtPath := temp
      else begin LogMsg('[ERROR] - Could not read config value SFTP-custCreatePath'); result := false; exit end;
      temp := config.ReadString('SFTP','custUpdatePath',''); if temp <> '' then sftpCustUpdPath := temp
      else begin LogMsg('[ERROR] - Could not read config value SFTP-custUpdatePath'); result := false; exit end;

      temp := config.ReadString('SFTP','keyPath',''); if temp <> '' then FileStor.Path := temp
      else begin LogMsg('[ERROR] - Could not read config value SFTP-keyPath'); result := false; exit end;

      //Crystal Reports
      temp := config.ReadString('CrystalReports','rptsPath',''); if temp <> '' then rptsPath := temp
      else begin LogMsg('[ERROR] - Could not read config value CrystalReports-rptsPath'); result := false; exit end;
      temp := config.ReadString('CrystalReports','pathToDLL',''); if temp <> '' then dllPath := temp
      else begin LogMsg('[ERROR] - Could not read config value CrystalReports-pathToDLL'); result := false; exit end;

      //APPLICATION
      temp := config.ReadString('Application','path','');
      if temp <> '' then
      begin
        if temp[length(temp)] <> '\' then
        begin
          temp := temp + '\';
        end;

        appPath := temp + dbDatabase + '\';
        logPath := appPath + 'Logs\';
      end
      else begin LogMsg('[ERROR] - Could not read config value Application-path'); result := false; exit end;

      temp := config.ReadString('Application','profanityFilterPath',''); if temp <> '' then profFilterPath := temp
      else begin LogMsg('[ERROR] - Could not read config value Application-profanityFilterPath'); result := false; exit end;
      //version 1.2.3.0 - set upon ini read and is an SQL[2] line
      collector := config.ReadString('Application','collectorNewClient','');

      numRetries := config.ReadInteger('Application','numberOfRetries',3);
      waitLength := (config.ReadInteger('Application','retryWaitLength',3)*1000);

      //CSV
      temp := config.ReadString('CSV','inboundSeperator',''); if temp <> '' then inboundSep := temp
      else begin LogMsg('[ERROR] - Could not read config value CSV-inboundSeperator'); result := false; exit end;
      inboundDelim := config.ReadString('CSV','inboundDelimiter','');

      //EMAIL
      temp := config.ReadString('Mail','from',''); if temp <> '' then emailFrom := temp
      else begin LogMsg('[ERROR] - Could not read config value Mail-from'); result := false; exit end;
      mailLog := config.ReadBool('Mail','logging',true);
      if mailLog then mailLogDtl := config.ReadBool('Mail','detailedLogging',false);

      temp := config.ReadString('Mail','listSeperator','');
      if temp <> '' then
      begin
        temp2 := temp;
        temp := config.ReadString('Mail','accountingRecipients','');
        if temp <> '' then
        begin
          if ContainsText(temp,temp2) then
          begin
            tempArr := temp.Split([temp2[1]]);
            acctRecips := tempArr;
          end
          else
            acctRecips := [temp];
        end
        else begin LogMsg('[ERROR] - Could not read config value Mail-accountingRecipients'); result := false; exit end;

        temp := config.ReadString('Mail','ITRecipients','');
        if temp <> '' then
        begin
          if ContainsText(temp,temp2) then
          begin
            tempArr := temp.Split([temp2[1]]);
            ITRecips := tempArr;
          end
          else
            ITRecips := [temp];
        end
        else begin LogMsg('[ERROR] - Could not read config value Mail-ITRecipients'); result := false; exit end;

        //taxExRecips, taxExCCS
        temp := config.ReadString('Mail','taxExemptRecipients','');
        if temp <> '' then
        begin
          if ContainsText(temp,temp2) then
          begin
            tempArr := temp.Split([temp2[1]]);
            taxExRecips := tempArr;
          end
          else
            taxExRecips := [temp];
        end;
        temp := config.ReadString('Mail','taxExemptCCs','');
        if temp <> '' then
        begin
          if ContainsText(temp,temp2) then
          begin
            tempArr := temp.Split([temp2[1]]);
            taxExCCS := tempArr;
          end
          else
            taxExCCS := [temp];
        end;
      end
      else begin LogMsg('[ERROR] - Could not read config value Mail-listSeperator'); result := false; exit end;

      temp := config.ReadString('Mail','sendOnEmptyExtract','');
      if temp <> '' then
      begin
        if temp = '0' then
          sendOnEmpty := false
        else if temp = '1' then
          sendOnEmpty := true
        else
          sendOnEmpty := true;
      end
      else
        sendOnEmpty := true;

      config.Free;
    end
    else
    begin
      try
        config := TInifile.Create(iniFileLoc);
      except
        on E : Exception do
        begin
          //LogMsg('There is no configuration file and failed to create one: ' + E.Message);
          EmailError('HRCxTM error: Action required','');
          result := false;
          exit
        end;
      end;

      try
        //write default values
        config.WriteString('DB','database',''); dbDatabase := '';

        //SFTP
        config.WriteString('SFTP','host',''); sftpHost := '';
        config.WriteInteger('SFTP','port',0); sftpPort := 0;
        config.WriteString('SFTP','user',''); sftpUser := '';
        config.WriteString('SFTP','privKeyName','eyHR1'); sftpPrivkeyName := '';
        config.WriteString('SFTP','arPath',''); sftpArPath := '';
        config.WriteString('SFTP','customerMasterPath',''); sftpCustMstPath := '';
        config.WriteString('SFTP','custCreatePath',''); sftpCustCrtPath := '';
        config.WriteString('SFTP','custUpdatePath',''); sftpCustUpdPath := '';
        config.WriteString('SFTP','keyPath',extractFileDir(application.ExeName)); FileStor.Path := extractFileDir(application.ExeName);

        //Crystal Reports
        config.WriteString('CrystalReports','rptsPath',''); rptsPath := '';
        config.WriteString('CrystalReports','pathToDLL',''); rptsPath := '';

        //APPLICATION
        config.WriteString('Application','path','');
        appPath := '';
        logPath := appPath + '';
        profFilterPath := RootIniLoc + '';
        config.WriteString('Application','profanityFilterPath','');
        config.WriteString('Application','collectorNewClient','');
          collector := '';      //version 1.2.3.0 - set upon ini read and is an SQL[2] line
        config.WriteInteger('Application','numberOfRetries',3); numRetries := 3;
        config.WriteInteger('Application','retryWaitLength', 3000); waitLength := 3000;

        //CSV
        config.WriteString('CSV','inboundDelimiter','"'); inboundDelim := '';
        config.WriteString('CSV','inboundSeperator',','); inboundSep := '|';

        //EMAIL    
        config.WriteString('Mail','from',''); emailFrom := '';
        config.WriteString('Mail','listSeperator',';'); //emailListSep := ';';
        config.WriteString('Mail','accountingRecipients',''); acctRecips := [''];
        config.WriteString('Mail','ITRecipients',''); ITRecips := [''];
        config.WriteString('Mail','taxExemptRecipients',''); taxExRecips := [''];
        config.WriteString('Mail','taxExemptCCs',''); taxExCCS := [''];
        config.WriteString('Mail','sendOnEmptyExtract','1'); sendOnEmpty := true;
        config.WriteBool('Mail','logging',true);
        config.WriteBool('Mail','detailedLogging',false);

        config.Free;
      except
        on E: Exception do
        begin
          LogMsg('[ERROR] - Failed to write values to config file: ' + E.Message);
          EmailError('HRCxTM error','Failed to write values to config file: ' + E.Message,ITRecips);
          result := false;
          exit
        end;
      end;
    end;

    SetupFileDirs;
  end
  else
    result := false;
end;


procedure TForm1.SetupFileDirs;
var
  path : string;
begin
  try
    //Create main directory
    path := appPath;
    if appPath[length(appPath)] <> '\' then
      path := path + '\';

    System.IOUtils.TDirectory.CreateDirectory(path);
    if DirectoryExists(path,true) then
    begin
      System.IOUtils.TDirectory.CreateDirectory(path + 'Inbound\');
      System.IOUtils.TDirectory.CreateDirectory(path + 'Inbound\Processed\');
      System.IOUtils.TDirectory.CreateDirectory(path + 'Inbound\Error\');

      System.IOUtils.TDirectory.CreateDirectory(path + 'Outbound\');
      System.IOUtils.TDirectory.CreateDirectory(path + 'Outbound\Processed\');
      System.IOUtils.TDirectory.CreateDirectory(path + 'Outbound\Error\');
      System.IOUtils.TDirectory.CreateDirectory(path + 'Outbound\Raw\');

      System.IOUtils.TDirectory.CreateDirectory(path + 'Patch\');
      System.IOUtils.TDirectory.CreateDirectory(path + 'Patch\Processed\');
      System.IOUtils.TDirectory.CreateDirectory(path + 'Logs\');
    end;

  except
    on E:Exception do
    begin
      EmailError('HRCxTM error: Action Required','Unable to create application directory at ' + appPath + '.' +
                 'Therefore no logging, SFTP sending/receiving, or CLIENT_ID patching will occur. Message: ' +
                 E.Message,ITRecips);
    end;
  end;
end;


{$ENDREGION}


{$REGION 'SFTP'}

function TForm1.InitSFTP : boolean;
var
  I,J : integer;
begin
  result := false;
  J := 0;
  LogMsg('[INFO] - Connecting to SFTP...');
  for I := 0 to 3 do
  begin
    if I > 0 then
    begin
      J := (J*2) + 1;
      LogMsg('[INFO] - Retrying in ' + inttostr(J) + ' minute(s)...');
      Sleep(J*60000);
    end;

    try
      //Replace this with values from config file/registry
      SSHClient.Authentication := atPublicKey;
      SSHClient.HostName := sftpHost;
      SSHClient.Port := sftpPort;
      SSHClient.User := sftpUser;
      SSHClient.PrivateKeyName := sftpPrivKeyName;
     except on E: Exception do
      begin
        LogMsg('[ERROR] - Failed setting config parameters for SFTP: ' + E.Message);
        //result := false;
        continue
      end;
    end;

    try
      SSHClient.Connect;
    except on E: Exception do
      begin
        LogMsg('[ERROR] - Failed connecting to SSH: ' + E.Message);
        //result := false;
        continue
      end;
    end;

    try
      SFTPClient.Initialize;
      result := true;
    except on E: Exception do
      begin
        LogMsg('[ERROR] - Failed connecting to SFTP: ' + E.Message);
        //result := false;
      end;
    end;

    if result = true then
      break
  end;

  if result = false then
    EmailError('HRCxTM Error','Failed connecting to SFTP server. Please check the logs located at ' + logPath,ITRecips)
  else
    LogMsg('[INFO] - ...Connected!');
end;


function TForm1.UploadSFTP(source,dest : string) : boolean;
begin
  result := true;
  try
    SFTPClient.UploadFile(source, dest, false);
  except on E: Exception do
    begin
      LogMsg('[ERROR] - Failed uploading' + source + ' to SFTP: ' +E.Message);
      result := false;
    end;
  end;
end;


function TForm1.DnloadSFTP(source,dest : string) : boolean;
begin
  result := true;
  if fileExists(dest) then
  begin
    LogMsg('[ERROR] - Destination file already exists. File: ' + dest + ' not downloaded.');
    result := false;
  end
  else
  begin
    try
      SFTPClient.DownloadFile(source,dest, false);
    except on E: Exception do
      begin
        LogMsg('[ERROR] - Failed downloading ' + source + ' from SFTP: ' +E.Message);
        result := false;
      end;
    end;
  end;
end;


function TForm1.MoveSFTP(source,dest : string) : boolean;
begin
  result := true;
  try
    SFTPClient.RenameFile(source,dest);
  except on E: Exception do
    begin
      LogMsg('[ERROR] - Failed moving ' + source +  ' to ' + dest + ' on SFTP: ' +E.Message);
      result := false;
    end;
  end;
end;


function TForm1.GetFileList(dirs : TArray<string>) : boolean;
var
  Handle: TScSFTPFileHandle;
  I : integer;
begin
  fileList := nil;
  sftpFilelist := nil;

  SetLength(fileList,2);
  SetLength(sftpFilelist,2);
  try
    for I := 0 to length(dirs) - 1 do
    begin

      Handle := SFTPClient.OpenDirectory(dirs[I]);
      try
        while not SFTPClient.EOF(Handle) do
          SFTPClient.ReadDirectory(Handle);
      finally
        SFTPClient.CloseHandle(Handle);
      end;
    end;
  except
    on E:Exception do
    begin
      LogMsg('[ERROR] - Trying to get list of files from SFTP on GetFileList: ' + E.Message);
    end;
  end;
end;


function TForm1.DestroySFTP : boolean;
begin
  result := true;
  try
    SFTPClient.Disconnect;
  except on E: Exception do
    begin
      LogMsg('[ERROR] - Failed disconnecting SFTP: ' + E.Message);
      result := false;
      exit
    end;
  end;

  try
    SSHClient.Disconnect;
  except on E: Exception do
    begin
      LogMsg('[ERROR] - Failed disconnecting SSH: ' + E.Message);
      result := false;
      exit
    end;
  end;

  if SFTPClient.Active then
  begin
    LogMsg('[ERROR] - SFTP did not disconnect');
    result := false;
  end;

  if SSHClient.Connected then
  begin
    LogMsg('[ERROR] - SSH did not disconnect');
    result := false
  end;

end;

{$ENDREGION}


{$REGION 'Extracts'}


function TForm1.ExpCSV(reportPath,csvPath : string) : boolean;
var
  err : string;
begin
  result := true;

  err := CryUtilsNonVCL.CreateNonVisPlugin(ModMgr, cryPlugin);
  if err <> '' then
  begin
    result := false;
    LogMsg(err);
    with cryPlugin as IReportManager do CleanupReport;
    exit
  end;

  err := CryUtilsNonVCL.LoadAndLogon(cryPlugin,reportPath,odbcName,dbDatabase,dbConnArr[1],decPass(dbConnArr[2]));
  if err <> '' then
  begin
    result := false;
    LogMsg(err);
    with cryPlugin as IReportManager do CleanupReport;
    exit
  end;

  with cryPlugin as IReportManager do
  begin
    try
      ExportReport(extractFileDir(csvPath) + '\',extractFileName(csvPath),'csv','','|',false);
    except on E: Exception do
      begin
        LogMsg('[ERROR] - Failed exporting report ' + extractFileName(reportPath) + ': ' + E.Message);
        result := false;
      end;
    end;

    CleanupReport;
  end;

end;


function TForm1.CleanupCSV(filename : string) : string;
var
  writeFile : TStreamWriter;
  readFile : TStreamReader;
  buffer, newFileName : string;
  buff : tbytes;
  csvFile : TFileStream;
begin
  result := '';

  try
    readFile := TStreamReader.Create(filename);
    buffer := readFile.ReadToEnd;
    readFile.Close;
    readFile.BaseStream.Free;
    readFile.Free;
  except
    on E: Exception do
    begin
      LogMsg('[ERROR] - CleanupCSV - Failed to read delimited file ' + filename + ': ' + E.Message);
      result := '';
    end;
  end;

  newFileName := ReplaceStr(filename,'_raw.csv','.csv');

  buffer := StringReplace(buffer,'"','',[rfReplaceAll]);

  try
    writeFile := TStreamWriter.Create(newFileName, false);
    writeFile.WriteLine(buffer);
    writeFile.Close;
    writeFile.BaseStream.Free;
    writeFile.Free;
  except
    on E: Exception do
    begin
      LogMsg('[ERROR] - CleanupCSV - Failed to write cleaned file ' + newFileName + ': ' + E.Message);
      result := '';
    end;
  end;

  result := newFileName;
  DelFile(filename);
end;


function TForm1.IsEmptyCSV(filename : string) : boolean;
const
  updHeaders : string = 'CLIENT_ID,Approved Credit Limit,CDF.HRC_ID';
var
  readFile : TStreamReader;
  buffer : string;
  lines,headers,comp : TArray<string>;
  I,J, len : integer;
  match : boolean;
begin
  result := false;
  try
    readFile := TStreamReader.Create(filename);
    buffer := readFile.ReadToEnd;
    readFile.Free;
    if buffer <> '' then
    begin
      if ContainsText(buffer,slinebreak) then
      begin
        lines := buffer.Split([slineBreak]);
        //look at buffArr[0] to see if it has the headers
        if (length(lines) > 1) and (lines[1] = '') and (lines[0] <> '') then
        begin
          headers := lines[0].Split([inboundSep]);
        end
        else
          exit
      end
      else if ContainsText(buffer,#$A) then
      begin
        lines := buffer.Split([#$A]);
        //look at buffArr[0] to see if it has the headers
        if (length(lines) > 1) and (lines[1] = '') and (lines[0] <> '') then
        begin
          headers := lines[0].Split([inboundSep]);
        end
        else
          exit
      end
      else
      begin
        headers := buffer.Split([inboundSep]);
      end;

      len := length(headers);
      if (len = 49) or (len = 3) then
      begin
        if len = 49 then
          comp := headerstring.split([','])
        else if len = 3 then
          comp := updHeaders.split([','])
        else
        begin
          result := false;
          exit
        end;

        result := true;
        for I := 0 to len do
        begin
          match := false;
          for J := 0 to len do
          begin
            if headers[I] = comp[J] then
            begin
              match := true;
              break;
            end;
          end;
          if match = false then
          begin
            result := false;
            exit
          end;
        end;
      end;
    end;
  except
    on E: Exception do
    begin
      LogMsg('[ERROR] - Unable to open file ' + filename + ' while running "isEmptyCSV": ' + E.Message);
      if readFile <> nil then
        readFile.Free;
    end;
  end;
end;


{$ENDREGION}


{$REGION 'Client_ID'}

              //SameCompany
function TForm1.SameCompany(clientID, cname : string; numWords : integer) : boolean;
var
  csvComp,dbComp : string;
begin
  result := false;
  with DataMod do
  begin
    qClientID.ParamByName('CLIENTID').value := clientID;
    qClientID.Open;
    try
      if qClientID.FieldByName('CLIENT_ID').Value = clientID then
      begin
        dbComp := string.Join('',uppercase(qClientID.FieldByName('NAME').value).Split(['LTD',' INC.',' CO.','.',',','!','@','#','$','%','^','&','*','(',')',
                                                                                       '+','=','<','>','?','/','\','|',';',':','''','"','`','~','[',']','{','}']));
        dbComp := string.Join(' ',dbComp.Split(['_','-','  ',' OF ',' AND ',' CO ']));

        var dbArr := dbComp.Split([' ']);
        dbComp := '';
        var csvArr := cname.Split([' ']);
        for var I := 0 to numWords - 1 do
        begin
          dbComp := dbComp + dbArr[I];
          csvComp := csvComp + csvArr[I];
        end;

        if dbComp = csvComp  then
          result := true;
      end;
    finally
      qClientID.Close;
    end;
  end;
end;

//JosephStyons https://stackoverflow.com/questions/54797/how-do-you-implement-levenshtein-distance-in-delphi

function TForm1.GenClientID(name,city : string) : string;
const
  rmArrSpace : array[0..2] of string = (' OF ',' AND ',' CO ');
  rmArrNoSpace : array[0..37] of string
  = ('LTD',' INC.',' CO.','L.L.C.','LLC.',' LLC.','_','-','.',',','!','@','#','$','%','^',
     '&','*','(',')','+','=','<','>','?','/','\','|',';',':','''','"','`','~','[',']','{','}');
var
  wordArr : TArray<string>;
  modifiedName, modifiedCity : string;
  I : integer;
  nameLen : integer;
begin //cases we need to be prepared for: 1 - word with only one letter. 2 - multiple words with only 1 letter 3 - total letters is < 6. 4 - Same first two words
      //we try normal, try add city, try use 3rd word instead of second. constraint is 6 >= CLIENT_ID <= 10
  result := '';
  modifiedName := trim(uppercase(name));  //store modified name, uppercase it
  modifiedCity := trim(uppercase(city));
  //remove strings ' OF' ' AND' ' LTD' ' CO'
  //remove characters _ - . , & ( ) [ ] { } / \ $ # @ ! % ^ * ? " ' < > | : ; ~ `
  modifiedName := String.Join('', modifiedName.Split(rmArrNoSpace));
  modifiedName := String.Join(' ', modifiedName.Split(rmArrSpace));
  //replace '  ' with ' '
  modifiedName := modifiedName.Replace('  ',' ');

  //split into separate words
  wordArr := modifiedName.Split([' ']);
  nameLen := 0;
  for I := 0 to length(wordArr) - 1 do
    nameLen := nameLen + length(wordArr[I]);

  if nameLen <= 6 then  //if 6 chars or less use all of them
  begin
    var doNum := false;
    var usedCity := false;
    result := string.Join('',wordArr);
    if nameLen < 6 then  //if less than 6 chars add additional information we have
    begin
      if modifiedCity <> '' then //add first 3 chars from city name if we have it
      begin
        for I := 1 to Min(Length(modifiedCity),3) do
          result := result + modifiedCity[I];
        usedCity := true;
      end
      else
        for I := 1 to 6 - nameLen do  //pad with zeros until we have 6 chars
          result := result + '0';
    end;

    if IsProfane(result) then  //if it is profane, return nothing as adding to it will only create a bad word with extra stuff on the end
      result := ''
    else
    begin
      if DataMod.ExistsClientID(result) then //if the client id already exists try to add the city if we havent already
      begin
        if (modifiedCity <> '') and (usedCity = false) then
        begin
          for I := 1 to Min(Length(modifiedCity),3) do
            result := result + modifiedCity[I];

          if DataMod.ExistsClientID(result) then
            doNum := true;
        end
        else
          doNum := true;

        if doNum then  //if we still have an ID that is already used, add a single digit to the end starting with 1 and trying until 9
          result := AddNumbers(result);
      end;
    end;
  end
  else
  begin
    var lastWord := wordArr[length(wordArr)-1];
    if (lastWord = 'INC') or (lastWord = 'CO') or (lastWord = 'LLC') then
      SetLength(wordArr,length(wordArr)-1);

    result := GenMethodA(wordArr);
    var done := false;
    if not IsProfane(result) then
    begin
      if DataMod.ExistsClientID(result) then
      begin
        if SameCompany(result,modifiedName,CIDwordsUsed) then
        begin
          var doNum := false;
          if (modifiedCity <> '') then
          begin
            for I := 1 to Min(Length(modifiedCity),3) do
              result := result + modifiedCity[I];

            if DataMod.ExistsClientID(result) then
              doNum := true;
          end
          else
            doNum := true;
          if doNum then
            result := AddNumbers(result);
          if result <> '' then
            done := true;
        end
        else
        begin
          if length(wordArr) > 2 then //if not same company, and more than 2 words in name - add up to 3 chars from 3rd word
          begin
            for var k := 1 to Min(Length(wordArr[2]),3) do
              result := result + wordArr[2][k];
            if not (DataMod.ExistsClientID(result) or IsProfane(result)) then
              done := true;
          end;
        end;
      end
      else
        done := true;
    end;

    if not done then
    begin //strip vowels and do method a again
      var ret : string;
      ret := GenMethodA(StripVowels(wordArr));

      if (ret = result) or (IsProfane(ret)) or (DataMod.ExistsClientID(ret)) then
      begin //Use method B
        result := DoMethodB(wordArr, modifiedCity); //If this fails it returns ''
      end
      else
        result := ret;
    end
    else
  end;
end;


function Tform1.DoMethodB(words : TArray<string>; city : string) : string;
var
  doNum : boolean;
  I : integer;
  modCity : string;
begin
  result := ''; doNum := false;
  //generate client id
  result := GenMethodB(words);
  //check for swears. at this point we dont ha ve another method so it has failed
  if IsProfane(result) then
    result := ''
  else
  begin
    //does this client ID already exist? can we make a new one with city name or numbers that is unique?
    if DataMod.ExistsClientID(result) then
    begin
      if city <> '' then
      begin
        modCity := string.Join('',city.Split([' ','.',',']));
        for I := 1 to Min(Length(modCity),3) do
          result := result + modCity[I];
        if DataMod.ExistsClientID(result) then
          doNum := true;
      end
      else
        doNum := true;

      if doNum then
      begin
        //add numbers until it doesnt exist in table
        result := AddNumbers(result);
      end;
    end;
  end;
end;


function TForm1.AddNumbers(clientID : string) : string;
var
  I : integer;
  modClientID : string;
begin
  result := '';
  for I := 1 to 9 do
  begin
    modClientID :=  clientID + inttostr(I);
    if DataMod.ExistsClientID(modClientID) = false then
    begin
      result := modClientID;
      break
    end;
  end;
end;


function TForm1.GenMethodA(words : TArray<string>) : string;
var
  I,J, totalLen : integer;
  lenArr : TArray<integer>;
begin
  result := '';
  if length(words) = 1 then
  begin
    SetLength(lenArr,1);
    lenArr[0] := min(length(words[0]),6);
  end
  else if length(words) = 2 then
  begin
    SetLength(lenArr,2);
    if length(words[0]) > length(words[1]) then
    begin
      lenArr[1] := Min(length(words[1]),3);
      lenArr[0] := 6 - lenArr[1];
    end
    else
    begin
      lenArr[0] := Min(length(words[0]),3);
      lenArr[1] := 6 - lenArr[0];
    end;
  end
  else if length(words) = 3 then
  begin
    SetLength(lenArr,3);
    if ((length(words[0]) + length(words[1])) < 3) then
    begin
      lenArr[0] := length(words[0]);
      lenArr[1] := length(words[1]);
      lenArr[2] := 6 - (lenArr[0] + lenArr[1]);
    end
    else if ((length(words[0]) + length(words[2])) < 3) then
    begin
      lenArr[0] := length(words[0]);
      lenArr[2] := length(words[2]);
      lenArr[1] := 6 - (lenArr[0] + lenArr[2]);
    end
    else if ((length(words[1]) + length(words[2])) < 3) then
    begin
      lenArr[1] := length(words[1]);
      lenArr[2] := length(words[2]);
      lenArr[0] := 6 - (lenArr[1] + lenArr[2]);
    end
    else
    begin
      lenArr[0] := Min(length(words[0]),3);
      lenArr[1] := Min(length(words[1]),3);
      if (lenArr[0] + lenArr[1]) < 6 then
        lenArr[2] := Min(6 - (lenArr[0] + lenArr[1]),length(words[2]))
      else
        SetLength(lenArr, 2);
    end;
  end
  else
  begin
    totalLen := 0;
    for I := 0 to length(words) - 1 do
    begin
      if totalLen < 6 then
      begin
        SetLength(lenArr, length(lenArr) + 1);
        lenArr[I] := Min(Min(length(words[I]),3),6 - totalLen);
        totalLen := totalLen + lenArr[I];
      end;
    end;
  end;

  CIDwordsUsed := 0;
  for I := 0 to length(lenArr) - 1 do
  begin
    if lenArr[I] > 0 then
      inc(CIDwordsUsed);
    for J := 1 to lenArr[I] do
      result := result + words[I][J];
  end;

end;


function TForm1.GenMethodB(words : TArray<string>) : string;
var
  I,J, totalLen,lenA,lenB,lenC : integer;
  lenArr : TArray<integer>;
begin
  result := '';
  if length(words) = 1 then
    //This has to be the same as result as method A so return nil
  begin
    result := '';
    exit
  end
  else if length(words) = 2 then
  begin
    SetLength(lenArr,2);
    lenA := length(words[0]); lenB := length(words[1]);
    if (lenA > lenB) and (lenB > 3) then
    begin
      //Normally take either 3 or fewer letters from word[1]
      //Instead take up to 5(better if we mix 2 words if possible) from word[1] and rest from word[0]
      lenArr[1] := Min(lenB,5);
      lenArr[0] := 6 - lenArr[1];
    end
    else if (lenA < lenB) and (lenA > 3) then
    begin
      //Normally take 3 or fewer letters from word[0]
      //Instead take up to 5(better if we mix 2 words if possible) from word[0] and rest from word[1]
      lenArr[0] := Min(lenA,5);
      lenArr[1] := 6 - lenArr[0];
    end
    else if (lenA = lenB) then
    begin
      if lenA > 3 then
      begin
        lenArr[0] := 4;
        lenArr[1] := 2;
      end
    end
    else
    begin
    // if smaller word is < 4 then result is the same as method A
      result := '';
      exit
    end;
  end
  else if length(words) = 3 then
  begin
    SetLength(lenArr,3);
    lenA := length(words[0]); lenB := length(words[1]); lenC := length(words[2]);
    if ((lenA + lenB) < 3) then
    begin
      lenArr[0] := lenA;
      lenArr[1] := lenB;
      lenArr[2] := 7 - (lenArr[0] + lenArr[1]);
    end
    else if ((lenA + lenC) < 3) then
    begin
      lenArr[0] := lenA;
      lenArr[2] := lenC;
      lenArr[1] := 7 - (lenArr[0] + lenArr[2]);
    end
    else if ((lenB + lenC) < 3) then
    begin
      lenArr[1] := lenB;
      lenArr[2] := lenC;
      lenArr[0] := 7 - (lenArr[1] + lenArr[2]);
    end
    else
    begin
      lenArr[0] := Min(lenA,4);
      lenArr[1] := Min(lenB,4);
      if ((lenArr[0] - 1) + (lenArr[1] - 1)) < 6 then
        lenArr[2] := Min(7 - ((lenArr[0] - 1) + (lenArr[1] - 1)),lenC)
      else
        SetLength(lenArr, 2);
    end;
  end
  else
  begin
    totalLen := 0;
    for I := 0 to length(words) - 1 do
    begin
      if totalLen < 6 then
      begin
        SetLength(lenArr, length(lenArr) + 1);
        if length(words[I]) > 2 then
        begin
          lenArr[I] := Min(Min(length(words[I]),4),7 - totalLen);
          totalLen := totalLen + (lenArr[I] - 1);
        end
        else
        begin
          lenArr[I] := length(words[I]);
          totalLen := totalLen + lenArr[I];
        end;
      end;
    end;
  end;

  for I := 0 to length(lenArr) - 1 do
  begin
    for J := 1 to lenArr[I] do
      if (length(words) > 2) and (length(words[I]) > 1) and (J = 2) then
        result := result + ''
      else
        result := result + words[I][J];
  end;
end;


function TForm1.StripVowels(words : TArray<string>) : TArray<string>;  //assumes uppercase
var
  tempArr : array of array[0..1] of string ;
  first,rest : string;
  I, len : integer;
begin
  //make a paralell array with no leading characters
  len := length(words);
  setlength(result,len);
  for I := 0 to len - 1 do
  begin
    if length(words[I]) > 1 then
    begin
      first := words[I][1];
      rest := words[I];
      delete(rest,1,1);
      rest := string.Join('', rest.Split(['A','E','I','O','U','Y']) );
      result[I] := first + rest;
    end
    else
      result[I] := words[I];
  end;
end;


function TForm1.IsProfane(str : string) : boolean;
var
  I : integer;
begin
  result := false;
  if profFilter = nil then
  begin
    profFilter := CreateProfanityFilter;
    if profFilter = nil then
      exit;
  end;

  for I := 0 to (length(profFilter) - 1) do
  begin
    if ContainsText(str,profFilter[I]) then
    begin
      result := true;
      break;
    end;
  end;
end;


function TForm1.CreateProfanityFilter : TArray<string>;
var
  filterFile : TStreamReader;
  buffer : string;
begin
  //Try to open file and catch exceptions
  //File location is from config file
  try
    filterFIle := TStreamReader.Create(profFilterPath);
    //read contents of file to a string
    buffer := filterFile.ReadToEnd;
    filterFile.Close;
    filterFile.BaseStream.Free;
    filterFile.Free;
  except
    on E: Exception do
    begin
      LogMsg('[ERROR] - Failed reading profanity filter file: ' + E.Message);
      exit
    end;
  end;

  //split on newline and return array
  buffer := uppercase(buffer);
  result := buffer.Split([slinebreak]);
end;

{$ENDREGION}


{$REGION 'Unit tests'}

procedure TForm1.TestExp;
var
  err : string;
begin
  err := CryUtilsNonVCL.LoadDLL(ModMgr,dllPath);
  if err <> '' then
    LogMsg(err)
  else
    LogMsg('Loaded DLL');

  if ExpCSV(rptsPath + 'CustomerExtract.rpt',appPath + '\Outbound\exportCSVtest.csv')
  then
  begin
    CleanUpCsv(appPath + '\Outbound\exportCSVtest.csv');
    LogMsg('Exported successfully');
  end;

  err := CryUtilsNonVCL.ReleasePlugin(modMgr,cryPlugin);
  if err <> '' then
    LogMsg(err)
  else
    LogMsg('Released plugin');


  err := CryUtilsNonVCL.UnloadDLL(modMgr);
  if err <> '' then
    LogMsg(err)
  else
    LogMsg('Unloaded DLL');
end;


procedure TForm1.TestSFTP;
begin
  InitSFTP;

  UploadSFTP(ExtractFileDir(application.ExeName) + '\localTest.csv', sftpCustPath + 'remoteTest.csv');
  DnloadSFTP(sftpCustPath + 'remoteTest.csv', ExtractFileDir(application.ExeName) + '\remoteTest.csv');

  DestroySFTP;
end;


procedure TForm1.TestSFTPFileList;
var
  I,J : Integer;
  expFName, downloadDir : string;
  dirs : TArray<string>;
begin
  InitSFTP;

  SetLength(dirs,2);
  dirs[0] := sftpCustCrtPath;
  dirs[1] := sftpCustUpdPath;
  GetFileList(dirs);

  for I := 0 to 1 do
  begin
    if I = 0 then
      downloadDir := sftpCustCrtPath
    else if I = 1 then
      downloadDir := sftpCustUpdPath;

    for J := 0 to length(sftpFileList[I]) - 1 do
    begin
      expFName := appPath + 'Inbound\' + StringReplace(sftpFileList[I][J],downloadDir,'',[]);
      if (DnloadSFTP(sftpFileList[I][J],expFName)) then
      begin
        SetLength(fileList[I],length(fileList[I]) + 1);
        fileList[I][J] := expFName;
      end
      else
      begin
        LogMsg('Failed to download ' + fileList[I][J] + ' from HRC');
        sftpFileList[I][J] := '';
      end;
    end;
  end;
  sftpFileList := nil;
  fileList := nil;
  DestroySFTP;
end;



procedure TForm1.TestDeltaUpdate;
var
  err,dateStr,expFName,uplFName : string;
begin
  if initSFTP then
  begin
    //load cryRptsDotNet.dll
    err := CryUtilsNonVCL.LoadDLL(ModMgr,dllPath);
    if err <> '' then
      LogMsg('Not running exports - ' + err)
    else
    begin //for closedAR, openAR and Customer extracts. load report, export to csv, and sftp upload
          //if there are any errors dont do consecutive steps, log and potentially notify someone
      System.SysUtils.FormatSettings.DateSeparator := '_';System.SysUtils.FormatSettings.TimeSeparator := '_';

      DateTimeToString(dateStr,'dd/mm/yy/hh/mm/ss',Now);
      expFName := appPath + 'Outbound\customer_master_' + dateStr + '.csv';
      if ExpCsv(rptsPath + 'CustomerExtract.rpt',expFName) then
      begin
        uplFName := CleanupCSV(expFName);
        if uplFName <> '' then
        begin
          //if successful update table so that we dont try to do it again for these values next timev
          err := DataMod.UpdCustDelta;
          if err = '' then
          begin
            MoveProcessedFile(uplFName, 'Outbound\processed\');
          end
          else
            LogMsg(err);
        end;
      end
      else
        LogMsg('Failed to export OpenARExtract.rpt. Not sent to HRC');
    end;

    err := CryUtilsNonVCL.ReleasePlugin(modMgr,cryPlugin);
    if err <> '' then
      LogMsg(err);

    err := CryUtilsNonVCL.UnloadDLL(modMgr);
    if err <> '' then
      LogMsg(err);

    DestroySftp;
  end;
end;


procedure TForm1.TestInsertUpdate;
begin

  ProcessCustExtract(extractFileDir(application.ExeName) + '\TESTDATA_notrailingcomma.csv',true);
end;


procedure TForm1.TestClientIDGen;
var
  temp : string;
  I : integer;
  logFile : TStreamWriter;
begin
 { temp := GenClientID('BANK OF MONTREAL','DAWSON CREEK');
  temp := GenClientID('BELL POLE CO. LTD.','LUMBY');
  temp := GenClientID('RICK & VERNA BOONSTRA','TELKWA');
  temp := GenClientID('NISGA''A VALLEY HEALTH AUTHORITY','NEW AIYANSH');
  temp := GenClientID('TOP-VALU FOOD PRODUCTS','VANCOUVER');
  temp := GenClientID('Dr. Jon LTD.','Vancouver');
  temp := GenClientID('Dr. Jon LTD.','');
  temp := GenClientID('L. O. ENTERPRISES','');
  temp := GenClientID('Shinzo Takashima','St. Louis');
  temp := GenClientID('Grimshaw Trucking','');
  temp := GenClientID('N GT Racing','');
  temp := '';       }

  logFile := TStreamWriter.Create(logPath + 'TestClientIDGen.txt', false);

  with DataMod.qSelCli do
  begin
    Open;

    if RecordCount > 0 then
    begin
      logFile.WriteLine('Original CLIENT_ID,Generated CLIENT_ID,SAME,NAME,CITY');
      for I := 1 to RecordCount do
      begin
        temp := '';
        if I < 10 then
          temp := '00'
        else if I < 100 then
          temp := '0';
        temp := GenClientID(FieldByName('NAME').asString,FieldByName('CITY').asString);
        logFile.WriteLine(
          FieldByName('CLIENT_ID').asString +
          ',' + temp +
          ',' + booltostr((temp = FieldByName('CLIENT_ID').asString),true) +
          ',' + FieldByName('NAME').asString +
          ',' + FieldByName('CITY').asString 
        );
        Next;
      end;
    end;

  end;

  logFile.Free;
end;


procedure TForm1.TestPatch;
begin
  PatchClientIDs;
end;


procedure TForm1.TestValidProv;
const
  arr : array[0..5] of string = ('Alberta','Alabama','Western Australia','Ontario','ON','BC');
var
  I : integer;
begin
  for I := 0 to 5 do
  begin
    lExchSuccess.Caption := lExchSuccess.Caption + '; ' + DataMod.ValidateProvince(arr[I]);
  end;

end;


procedure TForm1.TestCleanupCSV;
begin
  CleanupCSV(extractFileDir(application.exename) + '\open_ar_31_08_22_00_00_19_raw.csv');
end;


procedure TForm1.TestDeltaCleanup;
var
  res : string;
begin
  //res := DataMod.CleanDeltaTable;
  res := DataMod.ClearOldDelta;
  if res <> '' then
  begin
    if ContainsText(res,'[SEP]') then
    begin
      LogMsg( res.Split(['[SEP]'])[0] );
      LogMsg( res.Split(['[SEP]'])[1] );
    end
    else
      LogMsg(res);
    res := '';
  end;
end;

{$ENDREGION}


procedure TForm1.bExtractClick(Sender: TObject);
begin
  if Extract then
  begin
    lExtractSuccess.Caption := 'Success';
    LogMsg('[INFO] - Ran Extract successfully');
  end
  else
    lExtractSuccess.Caption := 'Failure'
end;


procedure TForm1.bUploadClick(Sender: TObject);
begin
  if InitSFTP then
  begin
    if Upload then
    begin
      lUploadSuccess.Caption := 'Success';
      LogMsg('[INFO] - Ran Upload successfully');
    end
    else
      lUploadSuccess.Caption := 'Failure';

    DestroySFTP;
  end
  else
    lUploadSuccess.Caption := 'Failure. Unable to establish SFTP connection.'
end;


procedure TForm1.bChangeCIDClick(Sender: TObject);
var
  oldCID,newCID,name : string;
begin
  oldCID := trim(uppercase(eOldCID.Text));
  newCID := trim(uppercase(eNewCID.Text));
  name := trim(uppercase(eCIDname.Text));
  if DataMod.ExistsClientID(oldCID) then //Does Client ID exist
  begin
    if name = '' then //is name blank
    begin
      name := DataMod.GetClientName;
      if MessageDlg('You didn''t enter a name. Would you like to proceed with the current name of "'+name+'"?',
                    mtConfirmation, [mbYes,mbNo],0) = mrNo then exit;
    end;

    if newCID = '' then //is new CID blank
    //generate cid  and ask for conf
    begin
      newCID := GenClientID(name,DataMod.GetClientCity);
      if MessageDlg('You didn''t enter a new CLIENT_ID. Would you like to proceed with the generated CLIENT_ID of "'+newCID+'"?',
                    mtConfirmation, [mbYes,mbNo],0) = mrNo then exit;
    end
    else if length(newCID) > 10 then   //is new one too long
    begin
      newCID := copy(newCID,1,10);
      if MessageDlg('The new CLIENT_ID that was entered is too long. Would you like to proceed with the truncated value of "'+newCID+'"?',
                    mtConfirmation, [mbYes,mbNo],0) = mrNo then exit;
    end;

    var impact := DataMod.ImpactAssess(oldCID);
    if impact <> '' then
    begin
      if MessageDlg('There are ' + impact +'. Are you sure you would you like to proceed?',
                    mtConfirmation, [mbYes,mbNo],0) = mrNo then exit;
    end;

    if DataMod.ChangeClientID(oldCID,newCID,name) then
      lChangeCID.Caption := 'Success!'
    else
      lChangeCID.Caption := 'Failure';

  end
  else
    ShowMessage('The CLIENT_ID "'+oldCID+'" does not exist in the database');
end;

procedure TForm1.bDBpassClick(Sender: TObject);
var
  err : string;
begin
  if DoDBConfig then
  begin
    //display message saying connected and written to config. you may use app
    showmessage('Connected to DB.');
    lDB.Caption := dbDatabase;
  end
  else
  begin
    //display message saying unable to connect to DB, same as initial
    showMessage('A connection to the database was unable to be established: ' + err +
                '. To retry click the "Configure Database" button and select a different database.');
  end;
end;


procedure TForm1.bDownloadClick(Sender: TObject);
begin
  if InitSFTP then
  begin
    if Download then
    begin
      lDownloadSuccess.Caption := 'Success';
      LogMsg('[INFO] - Ran Download successfully');
    end
    else
      lDownloadSuccess.Caption := 'Failure';

    DestroySFTP;
  end
  else
    lDownloadSuccess.Caption := 'Failure. Unable to establish SFTP connection.'
end;


procedure TForm1.bPatchClick(Sender: TObject);
begin
  if PatchClientIDs then
  begin
    lPatchSuccess.Caption := 'Success';
    logMsg('[INFO] - Ran Patch successfully');
  end
  else
    lPatchSuccess.Caption := 'At least one Client was not created successfully. Check your email or have IT look at the logs.';
end;


procedure TForm1.bProcessClick(Sender: TObject);
var
  num : integer;
begin
  num := Process;
  if num = 0 then
  begin
    lProcessSuccess.Caption := 'Success';
    LogMsg('[INFO] - Processed all local extracts successfully');
  end
  else if num = -1 then
    lProcessSuccess.Caption := 'No Extracts Present'
  else
    lProcessSuccess.Caption := 'Failed to process ' + inttostr(num) + ' extract(s)';
end;

procedure TForm1.bRunClick(Sender: TObject);
begin
 if RunDataExchange then
 begin
    lExchSuccess.Caption := 'Success';
    LogMsg('[INFO] - Ran Exchange successfully');
 end
  else
    lExchSuccess.Caption := 'Failed';
end;


procedure TForm1.bTestClick(Sender: TObject);
begin
  //Run a test function or procedure here
end;


procedure TForm1.SSHClientServerKeyValidate(Sender: TObject;
  NewServerKey: TScKey; var Accept: Boolean);
begin
  accept := true;
end;


procedure TForm1.SFTPClientCreateLocalFile(Sender: TObject; const LocalFileName,
  RemoteFileName: string; Attrs: TScSFTPFileAttributes; var Handle: NativeUInt);
var
  dwFlags: DWORD;
begin
  if aAttrs in Attrs.ValidAttributes then begin
    dwFlags := 0;
    if faReadonly in Attrs.Attrs then
      dwFlags := dwFlags or FILE_ATTRIBUTE_READONLY;
    if faSystem in Attrs.Attrs then
      dwFlags := dwFlags or FILE_ATTRIBUTE_SYSTEM;
    if faHidden in Attrs.Attrs then
      dwFlags := dwFlags or FILE_ATTRIBUTE_HIDDEN;
    if faArchive in Attrs.Attrs then
      dwFlags := dwFlags or FILE_ATTRIBUTE_ARCHIVE;
    if faCompressed in Attrs.Attrs then
      dwFlags := dwFlags or FILE_ATTRIBUTE_COMPRESSED;
  end
  else
    dwFlags := FILE_ATTRIBUTE_NORMAL;

  Handle := CreateFile(PChar(LocalFileName),
    GENERIC_READ or GENERIC_WRITE, 0, nil, CREATE_NEW, dwFlags, 0);

end;


procedure TForm1.SFTPClientDirectoryList(Sender: TObject; const Path: string;
  const Handle: TArray<System.Byte>; FileInfo: TScSFTPFileInfo; EOF: Boolean);
var
  len : integer;
begin
  if (FileInfo = nil) or (FileInfo.Filename = '.') or (FileInfo.Filename = '..')
  or (uppercase(FileInfo.Filename) = 'ERROR') or (uppercase(FileInfo.Filename) = 'PROCESSED') then
    Exit;
  //Do something for each file
  //If file is new then use it
  if (path = '/outbound/prod/crd/custcreate/') then
  begin
    len := length(sftpFileList[0]) + 1;
    setlength(sftpFileList[0],len);
    sftpFileList[0][len - 1] := Path + FileInfo.Filename;
  end
  else if path = '/outbound/prod/crd/custupdate/' then
  begin
    len := length(sftpFileList[1]) + 1;
    setlength(sftpFileList[1],len);
    sftpFileList[1][len - 1] := Path + FileInfo.Filename;
  end;
end;


end.
