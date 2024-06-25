unit Outlook;

interface

uses
  System.Classes, System.SysUtils, System.StrUtils, System.Math, System.NetEncoding, System.JSON, System.IOUtils, System.Win.Registry,
  Vcl.Dialogs, Vcl.Forms,
  CloudCustomOutlookJW, CloudOutlookWinJW, CloudCustomOutlookMailJW, CloudOutlookMailJW, CloudBaseJW,
  OutlookConfig;

type
  TOutlookController = class
    private
      OutlookMail : TAdvOutlookMail;
      authRetry : integer;

      clientid, clientsecret, redirectURI, tenantString : string;

      function GetAttachments(destination : string; numMails,pageOffset : integer; appUserID : string = '') : TArray<string>;
      function genHTML(input : string;  footer : string = '') : string;
    public
      constructor Create(logging : boolean = false; detailedLogging : boolean = false;user : string = ''; visual : boolean = true; Owner : TComponent = nil);
      function Auth(onBehalfOf : boolean = false) : boolean;

      function SendMail(format,fromEmail,sender,replyTo,subject,body : string; toEmails,ccEmails,bccEmails,attachmentPaths : TArray<string>; footer : string = '';
                        asUser : boolean = false) : boolean;
      function Send(fromEmail,subject,body : string; toEmails : TArray<string>;replyTo : string = ''; ccEmails : TArray<string> = nil;
                    bccEmails: TArray<string> = nil;attachmentPaths : TArray<string> = nil;format : string = 'HTML'; footer : string = ''; asUser : boolean = false) : boolean;
      function testDraft(format,fromEmail,sender,replyTo,subject,body : string; toEmails,ccEmails,bccEmails,attachmentPaths : TArray<string>;
                        asUser : boolean = false) : boolean;

      function GetMailsFromFolder(mailFolderName : string; userEmail : string = ''; numMails : integer = 100; pageOffset : integer = 0) : TOutlookMailItems;
      function GetAllMailsFromFolder(mailFolderName : string; userEmail : string = '') : TOutlookMailItems;
      function DownloadAttachments(mailFolderName,localDestination : string; userEmail : string = ''; numMails : integer = 100; pageOffset : integer = 0) : TArray<string>;
  end;


{
CURRENT VERSION (Jun 10, 2024): 1.0.2.0
Aug 8 2022:                     1.0.1.0
                                1.0.0.1
April 29 2022:                  1.0.0.0

Eventually would like to add the following:

  Functions to move emails from one folder to another
    One that behaves like GetMailsFromFolder and Download attachments, and
    One that moves just one specific email

}

implementation


constructor TOutlookController.Create(logging : boolean = false; detailedLogging : boolean = false; user : string = ''; visual : boolean = true; Owner : TComponent = nil);
var
  logPath, compName : string;
  reg : TRegistry;
  ret : TArray<string>;
begin
  //Check Registry settings
  try
    reg := TRegistry.Create;
    if reg.OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\',true) then
    begin
      reg.WriteInteger('SecureProtocols',2688);
    end;
  finally
    reg.Free;
  end;

  ret := GetConfigVals;
  if length(ret) >= 4 then
  begin
    if ( (ret[0] <> '') and (ret[1] <> '') and (ret[2] <> '') and (ret[3] <> '')) then
    begin
      clientid := ret[0];
      clientsecret := ret[1];
      redirectURI := ret[2];
      tenantString := ret[3];

      OutlookMail := TAdvOutlookMail.Create(Owner);

      OutlookMail.App.Key := clientid;
      OutlookMail.App.Secret := clientsecret;
      OutlookMail.App.CallBackURL := redirectURI;
      OutlookMail.App.Tenant := tenantString;

      if logging then
      begin
        OutlookMail.Logging := true;
        if detailedLogging then OutlookMail.LogLevel := llDetail
        else OutlookMail.LogLevel := llBasic;

        if ret[4] = '' then logPath :=  TPath.GetCachePath + '\'
        else logPath := ret[4];

        var dateStr : string; dateTimetoString(dateStr,'YYYY-MM-DD',now);

        if user <> '' then user := '\' + user;

        logPath := logPath + '\appMailLogs\' + replacestr(ExtractFileName(application.ExeName),'.exe','')+ user + '\' + dateStr + '\TMSOutlookMail.LOG';

        try
          System.IOUtils.TDirectory.CreateDirectory(ExtractFileDir(logPath));
        except
          on E: Exception do
          begin
            compName := GetEnvironmentVariable('COMPUTERNAME');
            SendMail(
              'TEXT',
              'example@email.com', //removed
              '',
              '',
              'OutlookUtils setup error',
              'The log directory (' + ExtractFileDir(logPath) + ') does not exist and OutlookUtils could not create it on ' + compName + '. Error: ' + E.Message,
              ['example@email.com'],
              nil,
              nil,
              nil
            );
          end;
        end;
        OutlookMail.LogFileName := logPath;
      end;

      authRetry := 0;

      //do token test and return err msg
      var authRet : boolean;
      try
        authRet := Auth;
      except on E:exception do
        Exception.Create('Exception on Auth Test. Message: ' + e.Message);
      end;
      if authRet = false then raise Exception.Create('Failed to authenticate on Auth Test. Try resetting your credentials.');
    end
    else
    begin
      var errMsg := 'Failed to create TOutlookController: Credential Value(s) missing - ';
      var crName := ['Client ID: ','Client Secret: ','Redirect URI: ','Tenant: '];
      for var I := 0 to 3 do
        if ret[I] = '' then
          errMsg := errMsg + crName[I] + ret[I] + ',';
      raise Exception.Create(errMsg);
    end;
  end
  else
  begin
    var errMsg := 'Failed to create TOutlookController: Credential(s) missing - ';
    var crName := ['Client ID, ','Client Secret, ','Redirect URI, ','Tenant '];
    for var I := 0 to length(ret) - 1 do
      if ret[I] = '' then
        errMsg := errMsg + crName[I];

    raise Exception.Create(errMsg);
  end;
end;

//Get folders, select desired folder, get list of emails in that folder and pass back to public
//using numMails and pageOffset assumes that any previous calls used the same page size(numMails)
function TOutlookController.GetMailsFromFolder(mailFolderName : string; userEmail : string; numMails,pageOffset : integer) : TOutlookMailItems;
var
  I : Integer;
  ret : boolean;
begin
  result := nil; ret := false;

  if userEmail <> '' then
    ret := OutlookMail.GetFolders(userEmail)
  else
    ret := OutlookMail.GetFolders;

  if ret = false then
  begin
    if userEmail <> '' then
      ret := Auth
    else
      ret := Auth(true);
  end;

  if ret then
  begin
    for I := 0 to OutlookMail.Folders.Count - 1 do
    begin
      if OutlookMail.Folders[I].DisplayName = mailFolderName then
      begin
        if userEmail <> '' then
          OutlookMail.GetMails(OutlookMail.Folders[I].ID,numMails,pageOffset,userEmail)
        else
          OutlookMail.GetMails(OutlookMail.Folders[I].ID,numMails,pageOffset);

        result := OutlookMail.Items;
      end;
    end;
  end;
end;

//Get folders, select desired folder, get list of emails in that folder and pass back to public
function TOutlookController.GetAllMailsFromFolder(mailFolderName : string; userEmail : string) : TOutlookMailItems;
var
  I,J, items: Integer;
  ret : boolean;
begin
  result := nil; ret := false;

  if userEmail <> '' then
    ret := OutlookMail.GetFolders(userEmail)
  else
    ret := OutlookMail.GetFolders;

  if ret = false then
  begin
    if userEmail <> '' then
      ret := Auth
    else
      ret := Auth(true);
  end;

  if ret then
  begin
    for I := 0 to OutlookMail.Folders.Count - 1 do
    begin
      if OutlookMail.Folders[I].DisplayName = mailFolderName then
      begin
        items := OutlookMail.Folders[I].ItemCount;
        if items > 100 then
        begin
          SetRoundMode(rmUp);
          //SetLength(result,items);
          for J := 0 to (Round(items/100)) do
          begin
            if userEmail <> '' then
              OutlookMail.GetMails(OutlookMail.Folders[I].ID,100,J,userEmail)
            else
              OutlookMail.GetMails(OutlookMail.Folders[I].ID,100,J);

            result := OutlookMail.Items;
          end;
        end
        else
        begin
          if userEmail <> '' then
            OutlookMail.GetMails(OutlookMail.Folders[I].ID,100,0,userEmail)
          else
            OutlookMail.GetMails(OutlookMail.Folders[I].ID);

          result := OutlookMail.Items;
        end;
      end;
    end;
  end;
end;

//since this is only called by DownloadAttachments we dont have to worry about number of mails or pages, as that is just based on the result of GetMailsFromFolder
function TOutlookCOntroller.GetAttachments(destination : string; numMails,pageOffset : integer; appUserID : string) : TArray<string>;
var
  I,J,K : integer;
  item : TOutlookMailItem;
  //jVal : TJSONValue;
  jObj : TJSONObject;
  jArr : TJSONArray;
  respJSON,fileName,fileType,fileByteString : string;
  decodedBytes,fileBytes : TBytes;
begin
  result := nil;
  if OutlookMail.Items.Count <> 0 then
  begin
    K := 0;
    for I := 0 to OutlookMail.Items.Count - 1 do
    begin
      item := OutlookMail.Items[I];

      if item.HasAttachments then
      begin
        SetLength(result,length(result)+1);
        result[K] := 'ID: ' + item.ID + ';From: ' + item.FromEmail + ';To: ';
        for J := 0 to item.RecipientNames.Count - 1 do
        begin
          result[K] := result[K] + item.RecipientNames[J];
          if J <> item.RecipientNames.Count - 1 then
            result[K] := result[K] + ',';
        end;
        result[K] := result[K] + ';Subject: ' + item.Subject + ';Attachments: ';

        //Send HTTP request to get attachment data for this email
        if appUserID <> '' then
          respJSON := OutlookMail.GetAttachmentList(item.ID,appUserID)
        else
          respJSON := OutlookMail.GetAttachmentList(item.ID);

        jObj := TJSONObject.ParseJSONValue(respJSON) as TJSONObject;

        if jObj.GetValue('error') <> nil then
        begin
          jObj := jObj.GetValue('error') as TJSONObject;
          result[K] := 'error: ' + jObj.GetValue('code').Value
                     + ' - ' + jObj.GetValue('message').Value;
        end
        else
        begin
          jArr := jObj.GetValue('value') as TJSONArray;

          for J := 0 to jArr.Count - 1 do
          begin
            //jObj := jArr.Get(J) as TJSONObject;
            jObj := jArr.Items[J] as TJSONObject;

            fileName := (jObj.GetValue('name') as TJSONValue).Value;
            fileType := (jObj.GetValue('contentType') as TJSONValue).Value;
            fileByteString := (jObj.GetValue('contentBytes') as TJSONValue).Value;

            fileBytes := Bytesof(fileByteString);

            //Decode and create file
            decodedBytes := TNetEncoding.Base64.Decode(fileBytes);

            TFile.WriteAllBytes(destination + '\' + fileName,decodedBytes);

            result[K] := result[K] + fileName;
            if J <> jArr.Count - 1 then
              result[k] := result[k] + ',';
          end;
          result[K] := result[K] + ';';
          inc(K);
          jArr.Free;
        end;
        //not sure why this is causing problems
        {jObj.Free;}
      end;
    end;
  end
  else
end;


function TOutlookController.Auth(onBehalfOf : boolean) : boolean;
var
  retVal : boolean;
begin
  result := false;
  if onBehalfOf then
    OutlookMail.DoAuth
  else
  begin
    retVal := false;
    retVal := OutlookMail.DoSvcAuth;

    if retVal then
      result := retVal
    else if authRetry < 3 then
    begin
      inc(authRetry);
      result := Auth;
    end;
  end;
end;



function TOutlookController.Send(
  fromEmail,subject,body : string;
  toEmails : TArray<string>;
  replyTo : string = '';
  ccEmails : TArray<string> = nil;
  bccEmails: TArray<string> = nil;
  attachmentPaths : TArray<string> = nil;
  format : string = 'HTML';
  footer : string = '';
  asUser : boolean = false
) : boolean;
var reply : string;
begin
  result := true;
  try
    if replyTo = '' then
      reply := fromEmail
    else
      reply := replyTo;

    if ( (fromEmail <> '') and (subject <> '') and (body <> '') and ((toEmails <> nil) and (toEmails[0] <> '')) ) then
      result := SendMail(format,fromEmail,fromEmail,reply,subject,body,toEmails,ccEmails,bccEmails,attachmentPaths,footer,asUser)
    else
    begin
      result := false;
      var compName : string;
      compName := GetEnvironmentVariable('COMPUTERNAME');
      var tosStr : string;
      if toEmails = nil then
        tosStr := 'nil'
      else if length(toEmails) < 1 then
        tosStr := ''''''
      else if length(toEmails) = 1 then
        tosStr := '[' + toEmails[0] + ']'
      else
      begin
        var I : integer;
        tosStr := '[';
        for I := 0 to length(toEmails) - 1 do
        begin
          tosStr := tosStr + toEmails[I];
          if I <> length(toEmails) - 1 then
            tosStr := tosStr + ', ';
        end;
        tosStr := tosStr + ']';
      end;
      SendMail(
        'TEXT',
        'example@email.com', //removed
        '',
        '',
        'OutlookUtils Send error',
        'Required values (fromEmail=' + fromEmail + '; subject=' + subject + '; body=' + body + '; toEmails=' + tosStr + ')' +
        ' are invalid and OutlookUtils did not send an email from app ' + application.ExeName + ' on ' + compName + slinebreak,
        ['example@email.com'], //removed
        nil,
        nil,
        nil
      );
    end;
  except on E: Exception do
    begin
      OutlookMail.LogException('SendMail: Exception - '+E.Message);    //New jun 2024
      result := false;
    end;
  end;
end;


function TOutlookController.SendMail(format,fromEmail,sender,replyTo,subject,body : string; toEmails,ccEmails,bccEmails,attachmentPaths : TArray<string>; footer : string = ''; asUser : boolean = false ) : boolean;
const
  MB = 1048576;
var
  mailitem : TOutlookMailItem;
  I, fSize, totalSize: Integer;
  incl : TStringList;
  attach : TArray<string>;
  attachLen : TArray<integer>;
  f : file;
begin
  result := true;
  try
    try
      mailItem := TOutlookMailItem.Create(nil);

      mailItem.SenderEmail := sender;
      mailItem.FromEmail := fromEmail;
      mailItem.Subject := subject;
      mailItem.Body := body;

      for I := 0 to (length(toEmails) - 1) do
      begin
        mailItem.RecipientEmails.Add(toEmails[I]);
      end;

      for I := 0 to (length(ccEmails) - 1) do
        mailItem.CcRecipientEmails.Add(ccEmails[I]);

      for I := 0 to (length(bccEmails) - 1) do
        mailItem.BccRecipientEmails.Add(bccEmails[I]);

      if uppercase(format) = 'HTML' then
      begin
        mailitem.MailType := TOutlookMailType.mtHTML;
        mailItem.Body := genHTML(body,footer);
      end
      else if uppercase(format) = 'HTMLRAW' then
      begin
        mailitem.MailType := TOutlookMailType.mtHTML;
        mailItem.Body := body + slineBreak + footer;
      end
      else if uppercase(format) = 'TEXT' then
      begin
        mailitem.MailType := TOutlookMailType.mtPlainText;
        mailItem.Body := body + slineBreak + footer;
      end;

      totalSize := 0;
      incl := TStringList.Create;
      for I := 0 to (length(attachmentPaths) - 1) do
      begin
        AssignFile(f,attachmentPaths[I]);
        Reset(f,1);
        fSize := FileSize(f);
        CloseFile(f);
        if (fSize < (3*MB)) and (((totalSize + fSize)) < (3*MB)) then
        begin
          totalSize := totalSize + fSize;
          incl.Add(attachmentPaths[I]);
        end
        else
        begin
          setlength(attach,length(attach) + 1); setlength(attachlen,length(attachlen) + 1);
          attach[length(attach) - 1] := attachmentPaths[I];
          attachlen[length(attach) - 1] := fSize;
        end;
      end;

      if incl.Count > 0 then
        mailItem.Attachments := incl;

      if length(attach) > 0 then
      begin
        //Do draft method
        result := OutlookMail.DraftSendMail(mailitem,attach,attachlen,replyTo,asUser);
        if result = false then
        begin
          Auth(asUser);
          result := OutlookMail.DraftSendMail(mailitem,attach,attachlen,replyTo,asUser);
        end;
      end
      else
      begin
        //Do send method
        result := OutlookMail.SendMessage(mailitem,replyTo,asUser);
        if result = false then
        begin
          Auth(asUser);
          result := OutlookMail.SendMessage(mailitem,replyTo,asUser);
        end;
      end;
    except on E: Exception do
      begin
        OutlookMail.LogException('SendMail: Exception - '+E.Message);    //New jun 2024
        result := false;
      end;
    end;
  finally
    mailitem.Free;
  end;
end;


function TOutlookController.TestDraft(format,fromEmail,sender,replyTo,subject,body : string; toEmails,ccEmails,bccEmails,attachmentPaths : TArray<string>; asUser : boolean) : boolean;
var
  mailitem : TOutlookMailItem;
  I: Integer;
begin
  mailItem := TOutlookMailItem.Create(nil);

  mailItem.SenderEmail := sender;
  mailItem.FromEmail := fromEmail;
  mailItem.Subject := subject;
  mailItem.Body := body;

  for I := 0 to (length(toEmails) - 1) do
  begin
    mailItem.RecipientEmails.Add(toEmails[I]);
  end;

  for I := 0 to (length(ccEmails) - 1) do
    mailItem.CcRecipientEmails.Add(ccEmails[I]);

  for I := 0 to (length(bccEmails) - 1) do
    mailItem.BccRecipientEmails.Add(bccEmails[I]);

  for I := 0 to (length(attachmentPaths) - 1) do
    mailItem.Attachments.Add(attachmentPaths[I]);

  if uppercase(format) = 'HTML' then
  begin
    mailitem.MailType := TOutlookMailType.mtHTML;
    mailItem.Body := genHTML(body);
  end
  else if uppercase(format) = 'HTMLRAW' then
  begin
    mailitem.MailType := TOutlookMailType.mtHTML;
  end
  else if uppercase(format) = 'TEXT' then
    mailitem.MailType := TOutlookMailType.mtPlainText;

  if asUser then
    result := OutlookMail.DraftSendMail(mailitem,'',sender,replyTo)
  else
    result := OutlookMail.DraftSendMail(mailitem,fromEmail,sender,replyTo);

  if result = false then
  begin
    if asUser then
    begin
      Auth(true);
      result := OutlookMail.DraftSendMail(mailitem,'',sender,replyTo);
    end
    else
    begin
      Auth;
      result := OutlookMail.DraftSendMail(mailitem,fromEmail,sender,replyTo);
    end;
  end;

  mailitem.Free;
end;


function TOutlookCOntroller.DownloadAttachments(mailFolderName,localDestination : string; userEmail : string; numMails,pageOffset : integer) : TArray<string>;
begin
  result := nil;
  if OutlookMail.Items.Count > 0 then
    OutlookMail.Items.Clear;

  if GetMailsFromFolder(mailFolderName,userEmail,numMails,pageOffset) <> nil then
    result := GetAttachments(localDestination,numMails,pageOffset,userEmail);
end;


function TOutlookController.genHTML(input : string; footer : string = '' ) : string;
var
  parList : TArray<string>;
  I : integer;

  function clearEmptyElements(inArr : TArray<string>) : TArray<string>;
  begin
    setlength(result,0);
    for var n := 0 to length(inArr) - 1 do
      if inArr[n] <> '' then
        result := result + [inArr[n]];
  end;
begin
  parList := clearEmptyElements(input.Split([slinebreak]));
  result := '<!DOCTYPE html><html><body>';

  for I := 0 to length(parList) - 1 do
  begin
    if (footer = '') and (I = length(parList) - 1) then
      result := result + '<p style="margin-bottom:2em">' + parList[I] + '</p>'
    else
      result := result + '<p>' + parList[I] + '</p>';
  end;

  if footer <> '' then
  begin
  var footerParList := clearEmptyElements(footer.Split([slinebreak]));
    for var J := 0 to length(footerParList)-1 do
    begin
      result := result + '<p style="';
      if j = 0 then
        result := result + 'margin:2em 2px 2px 2px">'
      else if j = length(footerParList) - 1 then
        result := result + 'margin:2px 2px 2em 2px">'
      else
        result := result + 'margin:2px">';
      result := result + footerParList[j] + '</p>';
    end;
  end;

   result := result + '</body></html>';
end;


end.
