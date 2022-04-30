unit uTimers;

{$mode delphi}
{$H+}

interface

uses
  Classes, SysUtils;

const
  UM_BASE           = 100;
  UM_TIMER          = UM_BASE + 1;

type
  TMessageProc = procedure (Sender : THandle; aMsg, aNo : integer) of object;

function AllocateHandle (aProc : TMessageProc) : THandle;
procedure DeallocateHandle (aHandle : THandle);
procedure SetTimer (aHandle : THandle; aNo : integer; anInterval : cardinal);
procedure KillTimer (aHandle : THandle; aNo : integer);
procedure PostMessage (aHandle : THandle; aMsg, aNo : integer);

implementation

uses SyncObjs, uLog, GlobalConst;

type

  TMessageData = record
    Msg : integer;
    No : integer;
  end;
  PMessageData = ^TMessageData;

  TMessageBank = class;

  { TTimerThread }

  TTimerThread = class (TThread)
    FNo : integer;
    FInterval : integer;
    FEvent : TEvent;
    FOwner : TMessageBank;
    procedure Execute; override;
    constructor Create (anOwner : TMessageBank; aNo : integer; anInterval : Cardinal);
    destructor Destroy; override;
  end;

  { TQueueThread }

  TQueueThread = class (TThread)
    FEvent : TEvent;
    procedure Execute; override;
    constructor Create;
    destructor Destroy; override;
  end;

  { TMessageBank }

  TMessageBank = class
    FProc : TMessageProc;
    FThreads : TList;
    FQueue : TList;
    FQueueEvent : TEvent;
    function GetThread (aNo : integer) : TTimerThread;
    function AddThread (aNo : integer; anInterval : Cardinal) : TTimerThread;
    function AddMsg (aMsg, aNo : integer) : PMessageData;
    procedure DoNext;
    constructor Create;
    destructor Destroy; override;
  end;

var
  MessageBanks : TList;
  i : integer;

function GetBank (byHandle : THandle) : TMessageBank;
var
  x : integer;
begin
  x := MessageBanks.IndexOf (pointer (byHandle));
  if x < 0 then
    Result := nil
  else
    Result := TMessageBank (MessageBanks[x]);
end;

function AllocateHandle (aProc : TMessageProc) : THandle;
var
  aBank : TMessageBank;
begin
 // Log ('allocating timers');
  aBank := TMessageBank.Create;
  aBank.FProc := aProc;
  MessageBanks.Add (aBank);
  Result := THandle (aBank);
//  Log ('timers allocated Timer Banks ' + IntToStr (MessageBanks.Count));
end;

procedure DeallocateHandle (aHandle : THandle);
var
  aBank : TMessageBank;
begin
  aBank := GetBank (aHandle);
  if aBank = nil then exit;
  MessageBanks.Remove (aBank);
  aBank.Free;;
end;

procedure SetTimer (aHandle : THandle; aNo : integer; anInterval : cardinal);
var
  aBank : TMessageBank;
  aThread : TTimerThread;
begin
 // Log ('Setting timer ' + aNo.ToString);
  aBank := GetBank (aHandle);
  if aBank = nil then exit;
  aThread := aBank.GetThread (aNo);
  if aThread <> nil then aThread.Free;
  aBank.AddThread (aNo, anInterval);
end;

procedure KillTimer (aHandle : THandle; aNo : integer);
var
  aBank : TMessageBank;
  aThread : TTimerThread;
begin
  aBank := GetBank (aHandle);
  if aBank = nil then exit;
  aThread := aBank.GetThread (aNo);
  if aThread = nil then exit;
  aBank.FThreads.Remove (aThread);
  aThread.FOwner := nil;
  aThread.Terminate;
//  aThread.FEvent.SetEvent;
end;

procedure PostMessage (aHandle : THandle; aMsg, aNo : integer);
var
  aBank :  TMessageBank;
begin
//  Log ('posting message ' + aMsg.ToString + ' No ' + aNo.ToString);
  aBank := GetBank (aHandle);
  if aBank = nil then exit;
  aBank.AddMsg (aMsg, aNo);
end;

{ TQueueThread }

procedure TQueueThread.Execute;
begin

end;

constructor TQueueThread.Create;
begin
  inherited Create (true);

end;

destructor TQueueThread.Destroy;
begin
  inherited Destroy;
end;

{ TTimerThread }

procedure TTimerThread.Execute;
var
  res : TWaitResult;
begin
  while not Terminated do
    begin
//      Log ('executing thread loop for ' + FNo.ToString);
      FEvent.ResetEvent;
      res := FEvent.WaitFor (FInterval);
//      Log (' res is ' + IntToStr (ord (res)) + ' vs ' + IntToStr (ord (wrTimeout)));
 //    if Assigned (FOwner) then FOwner.AddMsg (UM_TIMER, FNo);
  //   if (res = wrTimeout) and Assigned (FOwner) then
      if ((ord (res) = ERROR_WAIT_TIMEOUT) or (res = wrTimeout)) and Assigned (FOwner) then
        if Assigned (FOwner.FProc) then FOwner.FProc (THandle (FOwner), UM_TIMER, FNo);
    end;
 // Log ('Thread for ' + FNo.ToString + ' Terminated...');
end;

constructor TTimerThread.Create (anOwner : TMessageBank; aNo : integer; anInterval : cardinal);
begin
  inherited Create (true);
  FNo := aNo;
  FInterval := anInterval;
  FOwner := anOwner;
  FreeOnTerminate := true;
  FEvent := TEvent.Create (nil, true, false, '');
end;

destructor TTimerThread.Destroy;
begin
  FEvent.Free;
  inherited Destroy;
end;

{ TMessageBank }

function TMessageBank.GetThread (aNo: integer): TTimerThread;
var
  x :  integer;
begin
  Result := nil;
  for x := 0 to FThreads.Count - 1 do
    if TTimerThread (FThreads[x]).FNo = aNo then
      begin
        Result := TTimerThread (FThreads[x]);
        exit;
      end;
end;

function TMessageBank.AddThread (aNo: integer; anInterval: Cardinal): TTimerThread;
begin
  Result := TTimerThread.Create (Self, aNo, anInterval);
  FThreads.Add (Result);
  Result.Start;
end;

function TMessageBank.AddMsg (aMsg, aNo: integer): PMessageData;
begin
  New (Result);
  Result^.Msg:= aMsg;
  Result^.No:= aNo;
  FQueue.Add (Result);
end;

procedure TMessageBank.DoNext;
var
  aMsg : PMessageData;
begin
  if FQueue.Count = 0 then
    begin

    end
  else
    begin
      aMsg := FQueue[0];
      FQueue.Delete (0);
      if Assigned (FProc) then FProc (THandle (Self), aMsg.Msg, aMsg.No);
      Dispose (aMsg);
    end;
end;

constructor TMessageBank.Create;
begin
  FThreads := TList.Create;
  FQueue := TList.Create;
  FQueueEvent := TEvent.Create (nil, true, false, '');
end;

destructor TMessageBank.Destroy;
var
  i : integer;
begin
  FQueueEvent.Free;
  for i := 0 to FQueue.Count - 1 do
    Dispose (PMessageData (FQueue[i]));
  FQueue.Free;
  for i := 0 to FThreads.Count - 1 do
    TTimerThread (FThreads[i]).Terminate;
  FThreads.Free;
  inherited Destroy;
end;

initialization
  MessageBanks := TList.Create;
finalization
  for i := 0 to MessageBanks.Count - 1 do
    TMessageBank (MessageBanks[i]).Free;
  MessageBanks.Free;;
end.

