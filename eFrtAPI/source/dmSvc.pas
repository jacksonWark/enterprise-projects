unit dmSvc;

interface

uses
  System.SysUtils, System.Classes, FireDAC.UI.Intf, FireDAC.VCLUI.Wait,
  FireDAC.Phys.DB2Def, FireDAC.Phys.MSSQLDef, FireDAC.Phys.ODBCDef,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.Phys.Intf,
  FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys,
  FireDAC.Phys.DB2, FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf,
  FireDAC.DApt, FireDAC.Phys.ODBC, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client, FireDAC.Phys.MSSQL, FireDAC.Phys.ODBCBase,
  FireDAC.Comp.UI, service;

type
  TdmR = class(TDataModule)
    FDGUIxWaitCursor1: TFDGUIxWaitCursor;
    FDPhysDB2DriverLink1: TFDPhysDB2DriverLink;
    db2: TFDConnection;
    qAdmin: TFDQuery;
    qList: TFDQuery;
    qUpdTime: TFDQuery;
    qLogID: TFDQuery;
    qInsLogDtl: TFDQuery;
    qInsErr: TFDQuery;
    qToProcess: TFDQuery;
    spODRstat: TFDStoredProc;
    qUpdACHTLO: TFDQuery;
    qUpdTLO: TFDQuery;
    qListDS: TDataSource;
    qProfDtl: TFDQuery;
    qInsLog: TFDQuery;
    qUpdLogDtl: TFDQuery;
    qLogDtlID: TFDQuery;
    qUpdLogDtlACH: TFDQuery;
    qUpdLogDtlODR: TFDQuery;
    qUpdLogDtlTLO: TFDQuery;
    qToProcessAS: TFDQuery;
    qToProcessRW: TFDQuery;
    qRWdtl: TFDQuery;
    qRWodr: TFDQuery;
    qUpdODRrw: TFDQuery;
    qToProcessLONG: TFDQuery;
    qUpdTimeAll: TFDQuery;
    qUpdNotReqd: TFDQuery;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    function Logon(database,user,password,schema : string) : string;

    procedure qOpen(q: TFDQuery);
  end;

var
  dmR: TdmR;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

{ TdmR }


function TdmR.Logon(database,user,password,schema : string) : string;
begin
  result := '';
  try
    if db2.Connected then
      db2.Connected := false;
    db2.Params.Values['Alias'] := database;
    db2.Params.UserName := user;
    db2.Params.Password := password;
  except on E: Exception do
    begin
      result := 'Failed to set parameters for database: ' + E.Message;
      exit;
    end;
  end;
  try
    db2.Connected := true;
    db2.ExecSQL('SET CURRENT SCHEMA ' + uppercase(schema));
    db2.ExecSQL('SET CURRENT PATH "SYSFUN","SYSPROC","SYSIBMADM","' + uppercase(schema)+'"');
  except on E: Exception do result := 'Failed to connect to database: ' + database;
  end;
end;

procedure TdmR.qOpen(q: TFDQuery);
begin
  if q.Active then
  begin
    q.Refresh;
    q.First;
  end
  else
    q.Open();
end;

end.
