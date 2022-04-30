unit uIMAP;

{$mode Delphi}

interface

uses
  Classes, SysUtils, uTCPs, uTimers;

const
  IMAPPort     = 993;
  Tag          = 'TAG';
  UM_NEXT      = UM_BASE + 2;

  ctUnknown    = 0;
  ctLogin      = 1;
  ctFolder     = 2;
  ctUnseen     = 3;
  ctHeader     = 4;
  ctText       = 5;
  ctSearch     = 6;
  ctMarker     = 8;
  ctAll        = 9;
  ctDate       = 10;
  ctSubject    = 11;
  ctLogout     = 12;

type

  TMailMsg = class;
  TMailAttachment = class;
  TIMAP = class;

  TSearchRes = array of integer;
  TOnSearchResultEvent = procedure (Sender : TObject; Res : TSearchRes) of object;
  TOnMailMsgEvent = procedure (Sender : TObject; Msg : TMailMsg) of object;
  TOnCmdDoneEvent = procedure (Sender : TObject; Id : integer; Type_ : integer; Res : string) of object;
  TOnMarkerEvent = procedure (Sender : TObject; Id : string) of object;
  TOnStreamConvertEvent = procedure (anAttachment : TMailAttachment; aStream : TStream) of object;

  { TMailAttachment }

  TMailAttachment = class
    Owner : TIMAP;
    Name : string;
    ContentType : string;
    Contents : TStringList;
    procedure SaveToStream;
    function ConvertToStream : TMemoryStream;
    constructor Create (anOwner : TIMAP);
    destructor Destroy; override;
  end;

  { TSaveThread }

  TSaveThread = class (TThread)
    FName : integer;
    FAttachment : TMailAttachment;
    FStream : TStream;
    procedure Execute; override;
    constructor Create (anAttachment : TMailAttachment);
    destructor Destroy; override;
  end;


  { TMailMsg }

  TMailMsg = class
    UID : integer;
    Recipient : string;
    From : string;
    Subject : string;
    Date : string;
    ContentType : string;
    Text : TStringList;
    HTML : TStringList;
    Attachments : TList;
    function GetAttachment (byName : string) : TMailAttachment;
    procedure ClearAttachments;
    constructor Create;
    destructor Destroy; override;
  end;

  { TIMAP }

  TIMAP = class
  private
    TagNumber : integer;
    Sock : TTCPSClient;
    Buffer : string;
    HeaderDecoding : TMailMsg;
    TextDecoding : TMailMsg;
    Processing : boolean;
    Queue : TStringList;
    QueueHandle : THandle;
    CurrType : integer;
    CurrText : TStringList;
    IsText : boolean;
    IsAttachment : Boolean;
    CurrAttachment : TMailAttachment;
    FileName : string;
    LastWasBlank : Boolean;
    ContentType : string;
    Boundaries : TStringList;
    FOnSearchResult : TOnSearchResultEvent;
    FOnMsgHeader : TOnMailMsgEvent;
    FOnCmdDone : TOnCmdDoneEvent;
    FOnMarker : TOnMarkerEvent;
    FOnStreamConvert : TOnStreamConvertEvent;
    FOnConnect, FOnDisconnect : TNotifyEvent;
    procedure DoDebug (Sender : TObject; level : integer; s : string);
    procedure DoVerify (Sender : TObject; Flags : LongWord; var Allow : boolean);
    procedure DoAppRead (Sender : TObject; Buf : pointer; len : cardinal);
    procedure DoConnect (Sender : TObject; Vers, Suite : string);
    procedure DoDisconnect (Sender : TObject);
    procedure DoQueueProc (Sender : THandle; aMsg, aNo : integer);
  public
    Host : string;
    Username : string;
    Password : string;
    CurrentFolder : string;
    SearchResults : TSearchRes;
    Capabilities : TStringList;
    CmdDecoding : string;
    Mail : TList; // of TMailMsg
    Existing : integer;
    Recent : integer;
    UseUIDs : boolean; // to be implemented
    procedure Connect;
    procedure Close;
    procedure Test;
    function Login : integer;
    function Logout : integer;
    function Search (filter : string) : integer;
    function SearchUnseen : integer;
    function SearchAll : integer;
    function SearchSince (byDate : TDateTime) : integer;
    function SearchSubject (bySubject : string) : integer;
    function GetMail (byID : integer) : TMailMsg;
    procedure ClearMail;
    function FetchHeader (byID : Integer) : integer;
    function FetchText (byID : integer) : integer;
    function SelectFolder (fn : string) : integer;
    function Marker (id : string) : integer;
    function AddCommand (s : string) : integer; overload;
    function AddCommand (s : string; Type_ : integer) : integer; overload;
    function Connected : boolean;
    constructor Create;
    destructor Destroy; override;
    property OnSearchResult : TOnSearchResultEvent read FOnSearchResult write FOnSearchResult;
    property OnMsgHeader : TOnMailMsgEvent read FOnMsgHeader write FOnMsgHeader;
    property OnCmdDone : TOnCmdDoneEvent read FOnCmdDone write FOnCmdDone;
    property OnMarker : TOnMarkerEvent read FOnMarker write FOnMarker;
    property OnConnect : TNotifyEvent read FOnConnect write FOnConnect;
    property OnDisconnect : TNotifyEvent read FOnDisconnect write FOnDisconnect;
    property OnStreamConvert : TOnStreamConvertEvent read FOnStreamConvert write FOnStreamConvert;
  end;

function Split (s : string; delim : string) : TStringList;
function SplitQuoted (s : string) : TStringList;

implementation

uses uLog;

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

function Split (s : string; delim : string) : TStringList;
var
  i : integer;
begin
  Result := TStringList.Create;
  i := Pos (delim, s);
  while i > 0 do
    begin
      Result.Add (Copy (s, 1, i - 1));
      s := copy (s, i + length (delim));
      i := Pos (delim, s);
    end;
  if Length (s) > 0 then
    Result.Add (s);
end;

function SplitQuoted (s : string) : TStringList;
var
  i : integer;
  quoted : boolean;
begin
  Quoted := false;
  Result := TStringList.Create;
  Result.Add ('');
  for i := 1 to length (s) do
    begin
      if s[i] = '"' then
        Quoted := not Quoted
      else if Quoted then
        Result[Result.Count - 1] := Result[Result.Count - 1] + s[i]
      else if (s[i] = ' ') or (s[i] = #9) then
        begin
           if Result[Result.Count - 1] <> '' then
             Result.Add ('');
        end
      else
        Result[Result.Count - 1] := Result[Result.Count - 1] + s[i];
    end;
  if Result[Result.Count - 1] = '' then
    Result.Delete (Result.Count - 1);
end;

{ TSaveThread }

procedure TSaveThread.Execute;
var
  i, j, k : integer;
  s : string;
  len : integer;
  b, b0, b1, b2, b3 : integer;
const
  Base64In : array [0..127] of byte =
    (
        255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,
        255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,
        255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,
        255, 255, 255, 255,  62, 255, 255, 255,  63,  52,  53,  54,  55,
         56,  57,  58,  59,  60,  61, 255, 255, 255,  64, 255, 255, 255,
          0,   1,   2,   3,   4,   5,   6,   7,   8,   9,  10,  11,  12,
         13,  14,  15,  16,  17,  18,  19,  20,  21,  22,  23,  24,  25,
        255, 255, 255, 255, 255, 255,  26,  27,  28,  29,  30,  31,  32,
         33,  34,  35,  36,  37,  38,  39,  40,  41,  42,  43,  44,  45,
         46,  47,  48,  49,  50,  51, 255, 255, 255, 255, 255
    );

begin

  with FAttachment do
    begin
      if Copy (ContentType, 1, 4) = 'text' then
        begin
          for i := 0 to Contents.Count - 1 do
            begin
              s := Contents[i] + #13#10;
              FStream.Write (s[1], length (s));
            end;
        end
      else
        begin
          Log ('Starting Stream Transfer ' + ContentType + ' Count ' + IntTostr (Contents.Count));
          i := 0;
          s := '';
          while i < Contents.Count do
            begin
              s := s + Contents[i];
              i := i + 1;
              j := 1;
              len := Length (s);
              while j <= len - 3 do
                begin
                  b0 := Base64In[byte (s[j])];
                  b1 := Base64In[byte (s[j + 1])];
                  b2 := Base64In[byte (s[j + 2])];
                  b3 := Base64In[byte (s[j + 3])];
                  b := ((b0 and $3F) shl 2) + ((b1 and $30) shr 4);
                  FStream.Write (b, 1);
                  if b2 <> $40 then
                    begin
                      b := ((b1 and $0f) shl 4) + ((b2 and $3c) shr 2);
                      FStream.Write (b, 1);
                      if b3 <> $40 then
                        begin
                          b := ((b2 and $03) shl 6) + (b3 and $3f);
                          FStream.Write (b, 1);
                        end;
                    end;
                  j := j + 4;
                end;  // j < len - 3
              if j = len + 1 then
                s := ''
              else
                Delete (s, 1, j - 4);
            end;  // Contents
        end;
     end;
  Log ('Save complete');
  if Assigned (FAttachment.Owner) then
    if Assigned (FAttachment.Owner.FOnStreamConvert) then
      FAttachment.Owner.FOnStreamConvert (FAttachment, FStream);
end;

constructor TSaveThread.Create (anAttachment : TMailAttachment);
begin
  inherited Create (false);
  FAttachment := anAttachment;
  FStream := TMemoryStream.Create;
  FreeOnTerminate := true;
end;

destructor TSaveThread.Destroy;
begin
  FStream.Free;
  inherited Destroy;
end;

procedure TIMAP.DoDebug (Sender : TObject; level : integer; s : string);
begin
  //
end;

procedure TIMAP.DoVerify (Sender : TObject; Flags : LongWord; var Allow : boolean);
begin
  Allow := true;
end;

procedure TIMAP.DoAppRead (Sender : TObject; Buf : pointer; len : cardinal);
var
  i, j, x : integer;
  s : string;
  m : TMailMsg;
  Parts : TStringList;
  n, v : string;

begin
  if len <> 0 then
    begin
      SetLength (s, len);
      Move (Buf^, s[1], len);
      Buffer := Buffer + s;
      x := Pos (#13#10, Buffer);
      while x > 0 do
        begin
          s := Copy (Buffer, 1, x - 1);
          Buffer := Copy (Buffer, x + 2);
   //       Log (display (Copy (s, 1, 100)));
          Parts := SplitQuoted (s);
          if TextDecoding <> nil then    // if decoding text body
            begin
              if s = ')' then
                begin
                  TextDecoding := nil;
                end
              else if (Copy (s, 1, 2) = '--') then // boundary
                begin
                  if LastWasBlank then
                    begin
                //      Log ('Boundary - last was blank');
                  //    Log ('End of boundary - Content Type = ' + ContentType);
                      if isAttachment then
                        begin
                          if CurrAttachment <> nil then
                            if CurrAttachment.Contents.Count > 0 then
                               if CurrAttachment.Contents[CurrAttachment.Contents.Count - 1] = '' then
                                  CurrAttachment.Contents.Delete (CurrAttachment.Contents.Count - 1);
                        end
                      else if (ContentType = 'text/plain') and (TextDecoding.Text.Count > 0) then
                        begin
                          if TextDecoding.Text[TextDecoding.Text.Count - 1] = '' then
                            TextDecoding.Text.Delete (TextDecoding.Text.Count - 1);
                        end
                      else if (ContentType = 'text/html') and (TextDecoding.HTML.Count > 0) then
                        begin
                          if TextDecoding.HTML[TextDecoding.HTML.Count - 1] = '' then
                            TextDecoding.HTML.Delete (Textdecoding.HTML.Count - 1);
                        end;
                    end;
                  if Copy (s, length (s) - 1) = '--' then // boundary end
                    begin
                      if Boundaries.Count > 0 then
                        if Copy (s, 1, length (s) - 2) = Boundaries[Boundaries.Count - 1] then
                           begin
                           //  Log ('Found end of boundary ' + Boundaries[Boundaries.Count - 1]);
                             if LastWasBlank then
                               begin
                                 if IsAttachment then
                                   begin
                                   end
                                 else
                                   begin

                                   end; // not attachment
                               end;
                           end;
                    end
                  else
                    begin
                   //   Log ('Adding Boundary ' + s);
                      Boundaries.Add (s);
                    end;
                  isText := false;
                  isAttachment := false;
                  FileName := '';
                  ContentType := '';
                  CurrAttachment := nil;
                  CurrText := nil;
                end
              else if IsText then
                begin
               //   Log ('Is Text - Content Type ' + ContentType + ' Is Attachment ' + ny[IsAttachment]);
                  if IsAttachment then
                    begin
                      if CurrAttachment <> nil then
                        CurrAttachment.Contents.Add (s);
                    end
                  else if ContentType = 'text/plain' then
                    TextDecoding.Text.Add (s)
                  else if ContentType = 'text/html' then
                    TextDecoding.HTML.Add (s);
                end
              else // not is Text
                begin
                  if (s = '') and (not IsText) then
                    begin
                      IsText := true;
                   //   Log ('Setting Is Text true');
                    end
                  else if Parts.Count > 1 then
                    begin
                      for i := 1 to Parts.Count - 1 do
                        if Copy (Parts[i], length (Parts[i])) = ';' then
                          Parts[i] := Copy (Parts[i], 1, length (Parts[i]) - 1); // remove any ;
                      if Parts[0] = 'Content-Type:' then
                        begin
                          ContentType := Parts[1];
                        //  Log ('Content Type = ' + ContentType);
                          IsAttachment := false;
                          for j := 2 to Parts.Count - 1 do
                            begin
                              x := Pos ('=', Parts[i]);
                              if x > 0 then
                                begin
                                  n := Copy (Parts[i], 1, x - 1);
                                  v := Copy (Parts[i], x + 1);
                                  if n = 'name' then
                                    begin
                                      //Log ('This is an attachment');
                                      IsAttachment := true;
                                      FileName := v;
                                      CurrAttachment := TMailAttachment.Create (Self);
                                      CurrAttachment.Name := v;
                                      CurrAttachment.ContentType := ContentType;
                                      TextDecoding.Attachments.Add (CurrAttachment);
                                    end
                                  else if s = 'boundary' then
                                    begin
                                //      Log ('Start boundary ' + s);
                                      Boundaries.Add (s);
                                    end;
                                end;
                            end;
                        end
                      else if Parts[0] = 'Content-Disposition:' then
                        begin
                          for j := 2 to Parts.Count - 1 do
                            begin
                              x := Pos ('=', Parts[i]);
                              if x > 0 then
                                begin
                                  n := Copy (Parts[i], 1, x - 1);
                                  v := Copy (Parts[i], x + 1);
                                  if n = 'filename' then
                                    begin
                                      //Log ('This is an attachment');
                                      IsAttachment := true;
                                      FileName := v;
                                      CurrAttachment := TextDecoding.GetAttachment (v);
                                      if CurrAttachment = nil then
                                        begin
                                          CurrAttachment := TMailAttachment.Create (Self);
                                          CurrAttachment.Name := v;
                                          CurrAttachment.ContentType := ContentType;
                                          TextDecoding.Attachments.Add (CurrAttachment);
                                        end;
                                    end;
                                end
                              else // no '='
                                begin
                                  if Parts[i] = 'attachment' then
                                    IsAttachment := true;
                                end;
                            end;
                        end
                    end;
                end;

              LastWasBlank := (s = '');
            end // Text Decoding
          else if Parts.Count >= 2 then
            begin
              if Parts[0] = '*' then
                begin
                  if Parts[1] = 'CAPABILITY' then
                    begin
                      Capabilities.Clear;
                      for j := 2 to Parts.Count - 1 do
                        Capabilities.Add (Parts[j]);
                    end
                  else if Parts[1] = 'SEARCH' then
                    begin
                      SetLength (SearchResults, Parts.Count - 2);
                      for j := 2 to Parts.Count - 1 do
                        SearchResults[j - 2] := StrToIntDef (Parts[j], 0);
                      if Assigned (FOnSearchResult) then
                        FOnSearchResult (Self, SearchResults);
                    end
                  else if Parts.Count = 3 then
                    begin
                      if Parts[2] = 'EXISTS' then
                        Existing := StrToIntDef (Parts[1], 0)
                      else if Parts[2] = 'RECENT' then
                        Recent := StrToIntDef (Parts[1], 0);
                    end
                  else if Parts.Count > 3 then
                    begin
                      if Parts[2] = 'FETCH' then
                        begin
                          x := StrToIntDef (Parts[1], 0);
                          if x > 0 then
                            begin
                              if Pos ('TEXT', Parts[3]) > 0 then // fetching text
                                begin
                                  TextDecoding := GetMail (x);
                                  if TextDecoding = nil then
                                    begin
                                      TextDecoding := TMailMsg.Create;
                                      TextDecoding.UID := x;
                                      Mail.Add (TextDecoding);
                                    end;
                                  TextDecoding.Text.Clear;
                                  TextDecoding.HTML.Clear;
                                  CurrText := nil; // TextDecoding.Text;
                                  CurrAttachment := nil;
                                  IsText := false;
                                  IsAttachment := false;
                                  LastWasBlank := true; // expect boundary header
                                  Boundaries.Clear;
                                end
                              else   // assume header
                                begin
                                  HeaderDecoding := GetMail (x);
                                  if HeaderDecoding = nil then
                                    begin
                                      HeaderDecoding := TMailMsg.Create;
                                      HeaderDecoding.UID := x;
                                      Mail.Add (HeaderDecoding);
                                    end;
                                end;
                            end;
                        end
                    end;
                end // parts[0] = '*'
              else if Copy (Parts[0], 1, length (Tag)) = Tag then
                begin
                  x := StrToIntDef (Copy (Parts[0], length (Tag) + 1), 0);
                  if HeaderDecoding <> nil then
                    begin
                      if Assigned (FOnMsgHeader) then FOnMsgHeader (Self, HeaderDecoding);
                      HeaderDecoding := nil;
                    end;
                  TextDecoding := nil;
                  if Assigned (FOnCmdDone) then FOnCmdDone (Self, x, CurrType, Parts[1]);
                  CurrType := ctUnknown;
                  CurrText := nil;
                  SetTimer (QueueHandle, 1, 2);
                end // tag
              else if HeaderDecoding <> nil then
                begin
                  if Parts.Count > 1 then
                    begin
                      s := '';
                      for j := 1 to Parts.Count - 1 do
                        if j = 1 then
                          s := Parts[j]
                        else
                          s := s + ' ' + Parts[j];
                      if parts[0] = 'Subject:' then
                        HeaderDecoding.Subject := s
                      else if parts[0] = 'From:' then
                        HeaderDecoding.From := s
                      else if parts[0] = 'To:' then
                        HeaderDecoding.Recipient := s
                      else if parts[0] = 'Date:' then
                        HeaderDecoding.Date := s;
                    end;
                end;  // Header Decoding
            end;  // parts count >= 2
          Parts.Free;
          x := Pos (#13#10, Buffer);
        end;  // x > 0
    end;
end;

procedure TIMAP.DoConnect (Sender : TObject; Vers, Suite : string);
begin
//  Log ('Connected. Vers ' + Vers + ' Suite ' + Suite);
end;

procedure TIMAP.DoDisconnect (Sender : TObject);
begin
  if Assigned (FOnDisconnect) then FOnDisconnect (Self);
end;

procedure TIMAP.Connect;
begin
  Log ('Connecting using ' + Username + '  ' + Password);
  Sock.HostName := Host;
  Sock.RemoteAddress := '';
  Sock.RemotePort := IMAPPort;
  Buffer := '';
  isText := false;
  Processing := false;
  TextDecoding := nil;
  HeaderDecoding := nil;
  if Sock.Connect then
    if Assigned (FOnConnect) then FOnConnect (Self);
end;

procedure TIMAP.Close;
begin
  Sock.AppClose;
end;

function TIMAP.Login : integer;
begin
  Result := AddCommand (format ('LOGIN "%s" "%s"', [Username, Password]), ctLogin);
end;

function TIMAP.Logout : integer;
begin
  Result := AddCommand ('LOGOUT', ctLogout);
end;

function TIMAP.Search (filter : string) : integer;
begin
  if UseUIDs then
    Result := AddCommand (format ('UID SEARCH %s', [filter]), ctSearch)
  else
    Result := AddCommand (format ('SEARCH %s', [filter]), ctSearch);
end;

function TIMAP.SearchUnseen : integer;
begin
  if UseUIDs then
    Result := AddCommand ('UID SEARCH NOT SEEN', ctUnseen)
  else
    Result := AddCommand ('SEARCH NOT SEEN', ctUnseen);
end;

function TIMAP.SearchAll : integer;
begin
  if UseUIDs then
    Result := AddCommand ('UID SEARCH ALL', ctAll)
  else
    Result := AddCommand ('SEARCH ALL', ctAll);
end;

function TIMAP.SearchSince (byDate : TDateTime) : integer;
var
  d : string;
begin
  d := FormatDateTime ('dd-mmm-yyyy', byDate);
  if UseUIDs then
    Result := AddCommand (format ('UID SEARCH SINCE %s', [d]), ctAll)
  else
    Result := AddCommand (format ('SEARCH SINCE %s', [d]), ctAll);
end;

function TIMAP.SearchSubject (bySubject : string) : integer;
begin
  if UseUIDs then
    Result := AddCommand (format ('UID SEARCH SUBJECT %s', [bySubject]), ctSubject)
  else
    Result := AddCommand (format ('SEARCH SUBJECT %s', [bySubject]), ctSubject);
end;

function TIMAP.GetMail (byID: integer) : TMailMsg;
var
  i : integer;
  m : TMailMsg;
begin
  Result := nil;
  for i := 0 to Mail.Count - 1 do
    begin
      m := Mail[i];
      if m.UID = byID then
        begin
          Result := m;
          break;
        end;
    end;
end;

function TIMAP.FetchHeader (byID: Integer) : integer;
begin
  if useUIDs then
    Result := AddCommand (format ('UID FETCH %d body[header]', [byID]), ctHeader)
  else
    Result := AddCommand (format ('FETCH %d body[header]', [byID]), ctHeader);
end;

function TIMAP.FetchText (byID: integer) : integer;
begin
  if useUIDs then
     Result := AddCommand (format ('UID FETCH %d body[text]', [byID]), ctText)
  else
      Result := AddCommand (format ('FETCH %d body[text]', [byID]), ctText);
end;

function TIMAP.SelectFolder (fn : string) : integer;
begin
  Result := AddCommand (format ('SELECT "%s"', [fn]), ctFolder);
end;

function TIMAP.AddCommand (s : string) : integer;
begin
  Result := AddCommand (s, ctUnknown);
end;

function TIMAP.AddCommand (s : string; Type_ : integer) : integer;
begin
  TagNumber := TagNumber + 1;
  if TagNumber > $00fffff then TagNumber := 1;
  Queue.AddObject (s, TObject ((TagNumber and $00ffffff) + (Type_ * $1000000)));
  Result := TagNumber;
  if not Processing then
    begin
      Processing := true;
      SetTimer (QueueHandle, 1, 2);
    end;
end;

function TIMAP.Connected : boolean;
begin
  Result := Sock.Connected;
end;

constructor TIMAP.Create;
begin
  TagNumber := 0;
  Host := '';
  Username := '';
  Password := '';
  Buffer := '';
  CmdDecoding := '';
  UseUIDs := false;
  CurrType := ctUnknown;
  CurrText := nil;
  HeaderDecoding := nil;
  TextDecoding := nil;
  Processing := false;
  IsText := false;
  LastWasBlank := false;
  CurrAttachment := nil;
  IsAttachment := false;
  FileName := '';
  Boundaries := TStringList.Create;
  Capabilities := TStringList.Create;
  SetLength (SearchResults, 0);
  Mail := TList.Create;
  Queue := TStringList.Create;
  QueueHandle := AllocateHandle (DoQueueProc);
  Sock := TTCPSClient.Create;
  Sock.OnDebug := DoDebug;
  Sock.OnVerify := DoVerify;
  Sock.OnAppRead := DoAppRead;
  Sock.OnConnect := DoConnect;
  Sock.OnDisconnect := DoDisconnect;
end;

destructor TIMAP.Destroy;
begin
  try
    Sock.AppClose;
  finally
    Sock.Free;
    end;
  DeAllocateHandle (QueueHandle);
  Queue.Free;
  Capabilities.Free;
  Boundaries.Free;
  SetLength (SearchResults, 0);
  ClearMail;
  Mail.Free;
  inherited;
end;

procedure TIMAP.DoQueueProc (Sender : THandle; aMsg, aNo : integer);
var
  s : string;
  x : integer;
begin
  if aMsg <> UM_TIMER then exit;
  KillTimer (Sender, aNo);
  case aNo of
    1 :
      begin
        if Queue.Count > 0 then
          begin
            s := Queue[0];
            x := integer (Queue.Objects[0]) and $ffffff;
            CurrType := integer (Queue.Objects[0]) div $1000000;
            Queue.Delete (0);
            if CurrType = ctMarker then
              begin
                if Assigned (FOnMarker) then FOnMarker (Self, s);
                SetTimer (QueueHandle, 1, 2);
              end
            else
              try
                s := format ('%s%d %s'#13#10, [Tag, x, s]);
        //        log (display (s));
                Sock.AppWrite (s);
                Processing := true;
              except
                Processing := false;
                end;
          end
        else
          Processing := false;
      end;
  end;
end;

procedure TIMAP.Test;
begin
  SetTimer (QueueHandle, 1, 2);
end;

function TIMAP.Marker (id : string) : integer;
begin
  Result := AddCommand (id, ctMarker);
end;

procedure TIMAP.ClearMail;
var
  i : integer;
begin
  HeaderDecoding := nil;
  TextDecoding := nil;
  for i := 0 to Mail.Count - 1 do
    TMailMsg (Mail[i]).Free;
  Mail.CLear;
end;

{ TMailMsg }

constructor TMailMsg.Create;
begin
  UID := 0;
  Recipient := '';
  From := '';
  Subject := '';
  Date := '';
  ContentType := '';
  Text := TStringList.Create;
  HTML := TStringList.Create;
  Attachments := TList.Create;
end;

destructor TMailMsg.Destroy;
begin
  Text.Free;
  HTML.Free;
  ClearAttachments;
  Attachments.Free;
  inherited;
end;

function TMailMsg.GetAttachment (byName : string) : TMailAttachment;
var
  i : integer;
begin
  for i := 0 to Attachments.Count - 1 do
    begin
      Result := Attachments[i];
      if Result.Name = byName then exit;
    end;
  Result := nil;
end;

procedure TMailMsg.ClearAttachments;
var
  i : integer;
begin
  for i := 0 to Attachments.Count - 1 do
    TMailAttachment (Attachments[i]).Free;
  Attachments.Clear;
end;

{ TMailAttachment }

procedure TMailAttachment.SaveToStream ();
var
  fn : string;
  st : TSaveThread;
begin
  try
     st := TSaveThread.Create (Self);
     st.Start;
  except
  end;
end;

function TMailAttachment.ConvertToStream: TMemoryStream;
var
  i, j, k : integer;
  s : string;
  len : integer;
  b, b0, b1, b2, b3 : integer;
const
  Base64In : array [0..127] of byte =
    (
        255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,
        255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,
        255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, 255,
        255, 255, 255, 255,  62, 255, 255, 255,  63,  52,  53,  54,  55,
         56,  57,  58,  59,  60,  61, 255, 255, 255,  64, 255, 255, 255,
          0,   1,   2,   3,   4,   5,   6,   7,   8,   9,  10,  11,  12,
         13,  14,  15,  16,  17,  18,  19,  20,  21,  22,  23,  24,  25,
        255, 255, 255, 255, 255, 255,  26,  27,  28,  29,  30,  31,  32,
         33,  34,  35,  36,  37,  38,  39,  40,  41,  42,  43,  44,  45,
         46,  47,  48,  49,  50,  51, 255, 255, 255, 255, 255
    );

begin
  Result := TMemoryStream.Create;
  if Copy (ContentType, 1, 4) = 'text' then
    begin
      for i := 0 to Contents.Count - 1 do
        begin
          s := Contents[i] + #13#10;
          Result.Write (s[1], length (s));
        end;
    end
  else
    begin
      i := 0;
      s := '';
      while i < Contents.Count do
        begin
          s := s + Contents[i];
          i := i + 1;
          j := 1;
          len := Length (s);
          while j <= len - 3 do
            begin
              b0 := Base64In[byte (s[j])];
              b1 := Base64In[byte (s[j + 1])];
              b2 := Base64In[byte (s[j + 2])];
              b3 := Base64In[byte (s[j + 3])];
              b := ((b0 and $3F) shl 2) + ((b1 and $30) shr 4);
              Result.Write (b, 1);
              if b2 <> $40 then
                begin
                  b := ((b1 and $0f) shl 4) + ((b2 and $3c) shr 2);
                  Result.Write (b, 1);
                  if b3 <> $40 then
                    begin
                      b := ((b2 and $03) shl 6) + (b3 and $3f);
                      Result.Write (b, 1);
                    end;
                end;
              j := j + 4;
            end;  // j < len - 3
          if j = len + 1 then
            s := ''
          else
            Delete (s, 1, j - 4);
        end;  // Contents
    end;
end;

constructor TMailAttachment.Create (anOwner : TIMAP);
begin
  Owner := anOwner;
  Name := '';
  ContentType := '';
  Contents := TStringList.Create;
end;

destructor TMailAttachment.Destroy;
begin
  Contents.Free;
  inherited;
end;

end.

