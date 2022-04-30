program MailTest;

{$mode delphi}{$H+}
{$define use_tftp}

(*
  Test program for GMAIL accounts

  pjde 2022

*)

{$hints off}
{$notes off}
uses
  RaspberryPi3,
  GlobalConfig,
  GlobalConst,
  GlobalTypes,
  Platform,
  Threads,
  SysUtils,
  Classes, Console, uLog, uTCPS,
{$ifdef use_tftp}
  uTFTP, Winsock2,
{$endif}
  Ultibo, umbedTLS, uIMAP
  { Add additional units here };

type

  { TMailSystem }

  TMailSystem = class
    imap : TIMAP;
    procedure DoCmdDone (Sender : TObject; Id, Type_ : integer; Res : string);
    procedure DoMsgHeader (Sender : TObject; Msg : TMailMsg);
    procedure DoSearchResult (Sender : TObject; Res : TSearchRes);
    procedure DoMarker (Sender : TObject; id : string);
    procedure DoConnect (Sender : TObject);
    procedure DoDisconnect (Sender : TObject);
    procedure DoStreamConvert (anAttachment : TMailAttachment; aStream : TStream);
    constructor Create;
    destructor Destroy; override;
  end;

var
  Console1, Console2, Console3 : TWindowHandle;
  IPAddress : string;
  i, j, k, x : integer;
  s : array [0..255] of char;
  t : string;
  ms : TMailSystem;
  ml : TMailMsg;
  ch : char;
  ma : TMailAttachment;
  sl : TStringList;
  aStream : TMemoryStream;
  f : TFileStream;
  fn : string;

function display (s : string) : string; overload;
var
  i : integer;
begin
  Result := '';
  for i := 1 to length (s) do
    if s[i] in [' ' .. '~'] then
      Result := Result + s[i]
    else
      Result := Result + '<' + ord (s[i]).ToString + '>';
end;

function display (s : TStream) : string; overload;
var
  ch : char;
begin
  s.Seek (0, soFromBeginning);
  Result := '';
  ch := #0;
  while s.Position < s.Size do
    begin
      s.Read (ch, 1);
      if ch in [' ' .. '~'] then
        Result := Result + ch
      else
        Result := Result + '<' + ord (ch).ToString + '>';
    end;
end;

procedure Log1 (s : string);
begin
  ConsoleWindowWriteLn (Console1, s);
end;

procedure Log2 (s : string);
begin
  ConsoleWindowWriteLn (Console2, s);
end;

procedure Log3 (s : string);
begin
  ConsoleWindowWriteLn (Console3, s);
end;

procedure Msg2 (Sender : TObject; s : string);
begin
  Log2 ('TFTP - ' + s);
end;

{$ifdef use_tftp}
function WaitForIPComplete : string;
var
  TCP : TWinsock2TCPClient;
begin
  TCP := TWinsock2TCPClient.Create;
  Result := TCP.LocalAddress;
  if (Result = '') or (Result = '0.0.0.0') or (Result = '255.255.255.255') then
    begin
      while (Result = '') or (Result = '0.0.0.0') or (Result = '255.255.255.255') do
        begin
          sleep (1000);
          Result := TCP.LocalAddress;
        end;
    end;
  TCP.Free;
end;
{$endif}

procedure WaitForSDDrive;
begin
  while not DirectoryExists ('C:\') do sleep (500);
end;

{ TMailSystem }

procedure TMailSystem.DoCmdDone (Sender : TObject; Id, Type_ : integer; Res : string);
begin
  Log1 ('Command ' + IntToStr (Id) + ' Complete. Result ' + Res);
end;

procedure TMailSystem.DoMsgHeader (Sender : TObject; Msg : TMailMsg);
begin
  with Msg do
    begin
      Log1 ('Msg Header ' + IntToStr (Msg.UID) + ' -------------------');
      Log1 (' From "' + From + '" To "' + Recipient + '"');
      Log1 (' Subject "' + Subject + '" Date "' + Date + '"');
    end;
end;

procedure TMailSystem.DoSearchResult (Sender : TObject; Res : TSearchRes);
begin
//
end;

procedure TMailSystem.DoMarker (Sender : TObject; id : string);
var
  i : integer;
begin
  if id = 'INITIATED' then
    Log1 ('We are good to go.')
  else if id = 'DONE' then
    imap.Logout;
end;

procedure TMailSystem.DoConnect (Sender : TObject);
begin
  Log1 ('Connected....');
  imap.Login;
  imap.SelectFolder ('INBOX');
  imap.SearchSince (Date - 4);      // last 4 days
  imap.Marker ('INITIATED');
end;

procedure TMailSystem.DoDisconnect (Sender : TObject);
begin
  Log1 ('Disconnected....');
end;

procedure TMailSystem.DoStreamConvert (anAttachment: TMailAttachment;
  aStream: TStream);
begin
  Log1 ('Stream save Complete..' + anAttachment.Name + ' Size ' + IntToStr (aStream.Size));
end;

constructor TMailSystem.Create;
begin
  imap := TIMAP.Create;
  imap.OnCmdDone := ms.DoCmdDone;
  imap.OnSearchResult := DoSearchResult;
  imap.OnMsgHeader := DoMsgHeader;
  imap.OnMarker := DoMarker;
  imap.OnStreamConvert := DoStreamConvert;
  imap.OnConnect := DoConnect;
  imap.OnDisconnect := DoDisconnect
end;

destructor TMailSystem.Destroy;
begin
  inherited;
end;

begin
  Console1 := ConsoleWindowCreate (ConsoleDeviceGetDefault, CONSOLE_POSITION_LEFT, true);
  Console2 := ConsoleWindowCreate (ConsoleDeviceGetDefault, CONSOLE_POSITION_TOPRIGHT, false);
  Console3 := ConsoleWindowCreate (ConsoleDeviceGetDefault, CONSOLE_POSITION_BOTTOMRIGHT, false);
  SetLogProc (@Log1);
  Log3 ('IMAP Demo for accessing GMAIL accounts and the like.');
  WaitForSDDrive;
  Log2 ('SD Drive ready.');
  Log2 ('');
  mbedtls_version_get_string_full (s);
  Log3 ('TLS Version : ' + s);
  Log3 ('');
  Log3 ('1 - Clear.');
  Log3 ('2 - Connect and Search last 4 days.');
  Log3 ('3 - Logout.');
  Log3 ('4 - Fetch Header of last message in Search Results.');
  Log3 ('5 - Fetch Text / Attachments of last message in Search Results.');
  Log3 ('6 - Save Attachments of last message in Search Results.');
  Log3 ('8 - List Search Results.');
  Log3 ('9 - List Fetched Mail.');
  Log3 ('0 - Close (without Logging out).');
  Log3 ('');

{$ifdef use_tftp}
  IPAddress := WaitForIPComplete;
  Log3 ('Network ready. Local Address : ' + IPAddress + '.');
  Log3 ('');
  Log2 ('TFTP - Syntax "tftp -i ' + IPAddress + ' put kernel7.img"');
  SetOnMsg (@Msg2);
{$endif}

  ms := TMailSystem.Create;
  with ms do
    begin
      imap.Host := 'imap.gmail.com';
      imap.Username := 'your email address@gmail.com';
      imap.Password := 'your app password';
    end;
  ch := #0;
  while true do
    begin
      if ConsoleGetKey (ch, nil) then
        case ch of
          '1' : ConsoleWindowClear (Console1);
          '2' : ms.imap.Connect;
          '3' : ms.imap.AddCommand ('LOGOUT');
          '4' :
            if length (ms.imap.SearchResults) > 0 then
              ms.imap.FetchHeader (ms.imap.SearchResults[high (ms.imap.SearchResults)]);
          '5' :
            if length (ms.imap.SearchResults) > 0 then
              ms.imap.FetchText (ms.imap.SearchResults[high (ms.imap.SearchResults)]);
          '6' :
            begin
              if length (ms.imap.searchResults) > 0 then
                begin
                  ml := ms.imap.GetMail (ms.imap.SearchResults[high (ms.imap.SearchResults)]);
                  if ml <> nil then
                    with ml do
                      begin
                        if Attachments.Count > 0 then
                          begin
                            for j := 0 to Attachments.Count - 1 do
                              begin
                                ma := Attachments[j];
                                Log1 ('Saving Attachment ' + ma.Name);
                                aStream := ma.ConvertToStream;
                                try
                                  f := TFileStream.Create (ma.Name, fmCreate);
                                  aStream.Seek (0, soFromBeginning);
                                  aStream.SaveToStream (f);
                                  f.Free;
                                except
                                  Log1 ('Error saving attachment ' + ma.Name);
                                  end;
                                Log1 ('Attachment Saved');
                                aStream.Free;
                              end;
                          end;
                      end;
                end;
            end;
          '8' :
            begin
              s := '';
              for i := 0 to high (ms.imap.SearchResults) do
                if i = 0 then s := IntToStr (ms.imap.SearchResults[i])
               else s := s + ' ' + IntToStr (ms.imap.SearchResults[i]);
              Log1 (s);
            end;
          '9' :
            begin
              for i := 0 to ms.imap.Mail.Count - 1 do
                begin
                  with TMailMsg (ms.imap.Mail[i]) do
                    begin
                      Log1 (IntToStr (UID) + ' From "' + From + '" To "' + Recipient + '"');
                      Log1 ('   Subject "' + Subject + '" Date "' + Date + '"');
                      if Text.Count > 0 then
                        begin
                          Log1 ('Text -----');
                          for j := 0 to Text.Count - 1 do
                            Log1 ('Text : "' + Text[j] + '"');
                        end;
                      if HTML.Count > 0 then
                        begin
                          Log1 ('HTML -----');
                          for j := 0 to HTML.Count - 1 do
                            Log1 ('HTML : "' + HTML[j] + '"');
                        end;
                      Log1 ('-------------------------------------');
                      if Attachments.Count > 0 then
                        begin
                          for j := 0 to Attachments.Count - 1 do
                            begin
                              ma := Attachments[j];
                              Log1 ('---- Attachment ' + IntToStr (j) + ' -------');
                              Log1 ('FileName ' + ma.Name);
                              Log1 ('Content Type ' + ma.ContentType);
                            end;
                        end;
                    end;
                 end;
             end;
          '0' : ms.imap.Close;
          'Q', 'q' : break;
        end;
    end;
  ThreadHalt (0);



end.

