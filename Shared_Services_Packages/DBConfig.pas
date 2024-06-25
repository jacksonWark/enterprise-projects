unit dbConfig;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes, System.ImageList, System.IniFiles, System.UITypes, System.IOUtils,
  Vcl.Graphics,Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ImgList,Vcl.StdCtrls, Vcl.ExtCtrls, sharedServices;

type
  TDBConfigForm = class(TForm)
    ImageList1: TImageList;
    pSelect: TPanel;
    Panel4: TPanel;
    Panel3: TPanel;
    dbList: TListBox;
    Panel2: TPanel;
    Label2: TLabel;
    bSelect: TButton;
    procedure bSelectClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
    { Private declarations }
    synergize : boolean;
    iniDir : string;
  public
    { Public declarations }
    procedure init(syn : boolean = false);
  end;

var
  DBConfigForm: TDBConfigForm;

implementation

{$REGION 'FORM'}

procedure TDBConfigForm.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var
  dlgRes : integer;
begin
  if self.ModalResult = mrCancel then
  begin
    dlgRes := messageDlg('Are you sure you want to close the configuration window without finishing?',mtConfirmation,[mbYes,mbNo],0);
    if dlgRes = mrNo then
      CanClose := false;
  end;
end;


procedure TDBConfigForm.init(syn: boolean = false);
var
  files : TArray<string>;
  I : integer;
begin
  if syn then
  begin
    synergize := true;
    iniDir := synIniLoc;
    DBConfigForm.Caption := 'Synergize Configuration';
    label2.Caption := 'Synergize Name';
  end
  else
  begin
    synergize := false;
    iniDir := dbIniLoc;
  end;
  try
    files := TDirectory.GetFiles(iniDir);
    if length(files) > 0 then
    begin
      for I := 0 to length(files) - 1 do
        dbList.Items.Add(StringReplace(extractFilename(files[I]),'.ini','',[]));
    end;
  except
    on E: Exception do
    begin
      ShowMessage('File Access issues. Closing');
      Self.ModalResult := mrAbort;
    end;
  end;
end;


procedure TDBConfigForm.bSelectClick(Sender: TObject);
var
  config : TIniFile;
begin
  if (dbList.Items.Count <> 0) and (dbList.ItemIndex <> -1) and (dbList.Items[dblist.ItemIndex] <> '') then
  begin //set genini database to this value and close form
    try
      config := TiniFile.Create(UserIniLoc);   //change to user INI
    except
      on E: Exception do
      begin
        Showmessage('Unable to create ini file.');
        Self.ModalResult := mrAbort;
      end;
    end;

    try
      if synergize then
        config.WriteString('DB','synergizeDatabase',dbList.Items[dbList.ItemIndex])
      else
        config.WriteString('DB','database',dbList.Items[dbList.ItemIndex]);
      config.Free;

      Self.ModalResult := mrYes;
    except
      on E: Exception do
      begin
        Showmessage('Unable to write to ini file.');
        Self.ModalResult := mrAbort;
      end;
    end;
  end
  else
  begin
    ShowMessage('Please select a database configuration to use.');
  end;
end;


{$ENDREGION}

end.
