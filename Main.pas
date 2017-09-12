unit Main;

interface

uses
 Winapi.Windows,Winapi.Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, Dialogs, Winapi.ShellAPI,JSON, WinApi.WinInet, Vcl.StdCtrls,
  IniFiles, Generics.Collections, Generics.Defaults;
type
  TForm4 = class(TForm)
    function  request(host,path:String; AData: AnsiString; blnSSL: Boolean = True): AnsiString;
    procedure btnClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ReCaption(btn: TButton);
    procedure Log(Text:string);
  private
    { Private declarations }
  public
    { Public declarations }
  end;



var
  Form4: TForm4;
  vlogin, vtoken, vdomen:string;
  DomenDirectory: TDictionary<string, string>;
implementation

{$R *.dfm}

procedure TForm4.btnClick(Sender: TObject);
var response,temptoken:string;
Site:Pchar;
Params: TStringList;
ParamsStr: string;
JObject: TJSONObject;
JPair: TJSONPair;
i:integer;
MemberName,vsuccess: string;
FJSONObject:TJSONObject;
begin
     try
      vdomen:=StringReplace(TButton(Sender).Caption,vlogin+'@','',[rfReplaceAll, rfIgnoreCase]);
      DomenDirectory.TryGetValue(vDomen,vtoken);
      Params:=TStringList.Create;
      try
      if Length(vdomen)>0 then
      Params.Add('domain='+vdomen);
      Params.Add('login='+vlogin);
        if Params <> nil then
       begin
        Params.Delimiter := '&';
        ParamsStr := Params.DelimitedText;
       end;

       response:=request('pddimp.yandex.ru','api2/admin/email/get_oauth_token',AnsiString(ParamsStr),true);
       response:= StringReplace(response, ' xmlns:xi="http://www.w3.org/2001/XInclude"', '', [rfReplaceAll, rfIgnoreCase]);


       FJSONObject:=TJSONObject.ParseJSONValue(response) as TJSONObject;
       JObject := (FJSONObject as TJSONObject);//привели значение пары к классу TJSONObject

       try
       {проходим по каждой паре}
       for I := 0 to JObject.Count-1 do
       begin
        JPair:=JObject.Pairs[i];//получили пару по её индексу
        MemberName:=JPair.JsonString.Value;//определили имя
         {ищем в какое свойство записывать значение}
        if MemberName='oauth-token' then
         temptoken:=JPair.JsonValue.Value ;
        if MemberName='success' then
         vsuccess:=JPair.JsonValue.Value
       end;
          except
          raise Exception.Create('Ошибка разбора JSON');
          end;

      if vsuccess='ok' then
       begin
       Site := PChar('https://passport.yandex.ru/passport?mode=oauth&type=trusted-pdd-partner&error_retpath=mail.yandex.ru&access_token='+temptoken);
       ShellExecute( Handle, 'open',Site, nil, nil, SW_RESTORE );
       end
       else
       begin
       ShowMessage('Не удалось получить доступ к электронной почте!ERROR:'+response);
       Log('Не удалось получить доступ к электронной почте!ERROR:'+response);
       end;
     finally
      Params.Free;
     end;
     Except
     on E: Exception do
     Log(E.Message);
     end;
end;

function TForm4.request(host,path:String; AData: AnsiString; blnSSL: Boolean = True): AnsiString;
var
  Header      : TStringStream;
  pSession    : HINTERNET;
  pConnection : HINTERNET;
  pRequest    : HINTERNET;
  port        : Integer;
  bytes, b, pos: Cardinal;
  ResponseString: AnsiString;
begin
  Result := '';

  pSession := InternetOpen(nil, INTERNET_OPEN_TYPE_PRECONFIG, nil, nil, 0);

  if Assigned(pSession) then
  try

    if blnSSL then
      Port := INTERNET_DEFAULT_HTTPS_PORT
    else
      Port := INTERNET_DEFAULT_HTTP_PORT;
    pConnection := InternetConnect(pSession, PChar(host), port, nil, nil, INTERNET_SERVICE_HTTP, 0, 0);

    if Assigned(pConnection) then
    try

      pRequest := HTTPOpenRequest(pConnection, 'POST', PWideChar(path),nil, nil, nil, INTERNET_FLAG_SECURE, 0);
       if not Assigned(pRequest) then
        raise Exception.Create('Ошибка при выполнении функции HttpOpenRequest' + sLineBreak + SysErrorMessage(GetLastError));


      if Assigned(pRequest) then
      try
        Header := TStringStream.Create('');
        try
          with Header do
          begin
            WriteString('Host: ' + host + sLineBreak);
            WriteString('Content-type: application/x-www-form-urlencoded'+ SLineBreak);
            WriteString('PddToken: '+vtoken+ SlineBreak+SLineBreak);
          end;

          HttpAddRequestHeaders(pRequest, PChar(Header.DataString), Length(Header.DataString), HTTP_ADDREQ_FLAG_ADD);

          if HTTPSendRequest(pRequest, nil, 0, Pointer(AData), Length(AData)) then
          begin
              pos := 1;
              b := 1;
              ResponseString := '';
          while b > 0 do begin
          if not InternetQueryDataAvailable( pRequest, bytes, 0, 0 ) then
            raise Exception.Create('Ошибка при выполнении функции InternetQueryDataAvailable' + sLineBreak + SysErrorMessage(GetLastError));
          SetLength( ResponseString, Length(ResponseString) + bytes );
          InternetReadFile( pRequest, @ResponseString[Pos], bytes, b );
          Inc(Pos, b);
          end;
          Result :=ResponseString;
          end;


        finally
          Header.Free;
        end;
      finally
        InternetCloseHandle(pRequest);
      end;

    finally
      InternetCloseHandle(pConnection);
    end;
  finally
    InternetCloseHandle(pSession);
  end;
end;

procedure TForm4.FormCreate(Sender: TObject);
var Buffer: array[0..Pred(MAX_COMPUTERNAME_LENGTH+1)] of Char;
     Size: cardinal;
      Init1: TIniFile;
      i:Integer;
      vd,vt: String;
      btn:TButton;
      dir:String;
begin
  try
  Size := SizeOf(Buffer);
  GetComputerName(Buffer, Size);
  if buffer<>'RDP' then
  Application.Terminate;
  Size := SizeOf(Buffer);
  GetUserName(Buffer, Size);
  vlogin:=buffer;
  //Чтение
  FileSetAttr(ExtractFilePath(Application.ExeName)+'Enter2MailBox.ini', faHidden);
  Init1:= TIniFile.Create(ExtractFilePath(Application.ExeName)+'Enter2MailBox.ini');
  // Создаем справочник
  DomenDirectory := TDictionary<string, string>.Create;
  for i :=1 to 4 do
  begin
  vd:= init1.ReadString('Domen info','Domen'+InttoStr(i), '');
  vt:= init1.ReadString('Domen info','Token'+InttoStr(i), '');
  if vd<>'' then
  begin
  btn:=TButton.Create(nil);//Задать владельца
  btn.Parent:=TForm(Sender);//Задать родителя.
  btn.Align:=alTop;
  btn.Caption:=vlogin+'@'+vd;
  ReCaption(btn);
  btn.onClick:=btnClick;
  btn.Name:='btn'+InttoStr(i);
  if not DomenDirectory.ContainsKey(vd) then
  DomenDirectory.add(vd,vt) ;
  end;
  end;
  TForm(Sender).AutoSize:=True;
  Except
  on E: Exception do
     Log(E.Message);
  end;
 end;

procedure TForm4.Log(Text:string);
var
  F : TextFile;
  FileName : String;
  dt:string;
begin
  FileName := ExtractFilePath(Application.ExeName) + 'Log.txt';
  AssignFile(F, FileName);
  if FileExists(FileName) then
    Append(F)
  else
    Rewrite(F);
    dt:=DateToStr(Date);
    dt:=dt+' '+TimeToStr(Time);
  WriteLn(F, text+': ' +dt);
  CloseFile(F);
end;



procedure TForm4.ReCaption(btn: TButton);
var S: string;
begin
  S := btn.Caption;
  Canvas.Font.Size := btn.Font.Size;
  if (Form4.clientwidth < Canvas.TextWidth(s) + 10) then
  Form4.clientwidth:= Canvas.TextWidth(s) + 10;
end;

end.
