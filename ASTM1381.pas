unit ASTM1381;

interface

uses
 HComPort, SysUtils, Classes, Forms, StrUtils, SyncObjs, TypInfo, AnsiStrings;

const
 ACK = 6;
 ENQ = 5;
 ETB = 23;
 ETX = 3;
 LF = 10;
 NAK = 21; //15H
 STX = 2;
 EOT = 4;
 CR = 13;
 
 C_ACK = AnsiChar(6);
 C_ENQ = AnsiChar(5);
 C_ETB = AnsiChar(23);
 C_ETX = AnsiChar(3);
 C_LF = AnsiChar(10);
 C_NAK = AnsiChar(21);
 C_STX = AnsiChar(2);
 C_EOT = AnsiChar(4);
 C_CR = AnsiChar(13);
 
type
 TTransferState = (tsNeutral, tsEstabilished, tsTermination, tsTransfer, tsWaitForPacket);
 TReturnState = (rsOther, rsACK, rsNAK);
 TNotifyStr = procedure(s: AnsiString) of object;
 TNotifyProc = procedure;
 //TASTM1381Transfer = class;
 
 {TSendThread = class(TThread)
 private
   FOwner: TASTM1381Transfer;
 protected
   procedure Execute; override;
 public
   constructor Create(AOwner: TASTM1381Transfer);
 end;}
 
 TASTM1381Transfer = class(TObject)
 private
  //FPushPop: TPushPop;
  //FSendThread: TSendThread;
  //FOnByteReceived: TNotifyStr;
  FHSerialPort: THSerialPort;
  FPrevState: TTransferState;
  FTransferState: TTransferState;
  FLastMessage: TStringList;
  FLastPacket: AnsiString;
  FOnMessageRecieved: TNotifyProc;
  FOnPortError: TNotifyProc;
  FLastError: AnsiString;
  procedure OnRxChar(Sender: TObject; ResvSize: Integer);
  procedure OnError(Sender: TObject; Msg: AnsiString; var Error: integer);
  procedure OpenConnect;
  function WaitForACK: TReturnState;
  procedure ReadPacket;
  function Checksum(const s: AnsiString): Boolean;
  function GetCrc(const s: AnsiString): Word;
  procedure ACKResponse;
  procedure NAKResponse;
  procedure PacketReceived;
  procedure SetTransferState(const Value: TTransferState);
 public
  constructor Create;
  destructor Destroy; override;
  procedure CloseConnect;
  function CalcCRC(const s: AnsiString): Word;
  procedure SendFrame(const Frame: AnsiString);
  property LastMessage: TStringList read FLastMessage;
  property LastError: AnsiString read FLastError;
  property OnMessageRecieved: TNotifyProc read FOnMessageRecieved write FOnMessageRecieved;
  property OnPortError: TNotifyProc read FOnPortError write FOnPortError;
  //property OnByteReceived: TNotifyStr read FOnByteReceived write FOnByteReceived;
  property HSerialPort: THSerialPort read FHSerialPort;
  property TransferState: TTransferState read FTransferState write SetTransferState;
 end;
 
implementation

uses Math;

var
 WriteLogStrCriticalSection: TCriticalSection;
 
procedure CreateLogFile;
var
 f: textfile;
begin
 if FileExists('log.txt') then
  Exit;
 AssignFile(f, 'log.txt');
 rewrite(f);
 closefile(f);
end;

procedure WriteLog(const s: AnsiString); overload;
var
 f: textfile;
begin
 WriteLogStrCriticalSection.Enter;
 try
  AssignFile(f, 'log.txt');
{$I-}
  Append(f);
  WriteLn(f, AnsiString(DateTimeToStr(Now) + ' ') + s);
  CloseFile(f);
{$I+}
 finally
  WriteLogStrCriticalSection.Leave;
 end;
end;

function CopyLim(const s: AnsiString; StrBeg, StrEnd: integer): AnsiString;
begin
 Result := Copy(s, StrBeg, StrEnd - StrBeg + 1);
end;

function ReplReturns(const s: AnsiString): AnsiString;
begin
 Result := AnsiReplaceStr(AnsiReplaceStr(s, AnsiChar($0D), AnsiString('#$0D')), AnsiChar($0A), AnsiString('#$0A'));
end;

{ TASTM1381Transfer }

procedure TASTM1381Transfer.ACKResponse;
begin
 WriteLog('ACKResponse');
 FHSerialPort.SendByte(ACK);
end;

function TASTM1381Transfer.CalcCRC(const s: AnsiString): Word;
var
 i: Integer;
 s1: AnsiString;
 b: Byte;
begin
 WriteLog('CalcCRC');
 b := 0;
 for i := 1 to Length(s) do
  b := b + Ord(s[i]) mod 256;
 s1 := AnsiString(IntToHex(b, 2));
 Result := Ord(s1[2]) * 256 + Ord(s1[1]);
end;

//<STX> [F1] [DATA] <ETX> [C1] [C2] <CR> <LF>

function TASTM1381Transfer.Checksum(const s: AnsiString): Boolean;
var
 Temp: Integer;
begin
 WriteLog('Checksum');
 Result := False;
 Temp := Max(Pos(C_ETX, s), Pos(C_ETB, s));
 if Temp = 0 then
  Exit;
 Result := CalcCRC(Copy(s, 1, Temp)) = GetCrc(s);
end;

procedure TASTM1381Transfer.CloseConnect;
begin
 WriteLog('CloseConnect');
 FHSerialPort.SendByte(EOT);
 TransferState := tsNeutral;
end;

constructor TASTM1381Transfer.Create;
begin
 inherited;
 WriteLog('Create');
 //FPushPop := TPushPop.Create;
 //FSendThread := TSendThread.Create(Self);
 TransferState := tsNeutral;
 FLastMessage := TStringList.Create;
 FHSerialPort := THSerialPort.Create;
 FHSerialPort.OnRxChar := OnRxChar;
 FHSerialPort.OnError := OnError;
 //FHSerialPort.Open;
end;

destructor TASTM1381Transfer.Destroy;
begin
 WriteLog('Destroy');
 //FSendThread.Free;
 //FPushPop.Free;
 FHSerialPort.Free;
 FLastMessage.Free;
 inherited;
end;

function TASTM1381Transfer.GetCrc(const s: AnsiString): Word;
var
 s1: AnsiString;
 BegStr, EndStr: Integer;
begin
 WriteLog(AnsiString('GetCrc ') + ReplReturns(s));
 BegStr := Max(Pos(C_ETX, s), Pos(C_ETB, s));
 EndStr := PosEx(C_CR, string(s), BegStr);
 s1 := CopyLim(s, BegStr + 1, EndStr - 1);
 Result := Ord(s1[2]) * 256 + Ord(s1[1]);
end;

procedure TASTM1381Transfer.NAKResponse;
begin
 WriteLog('NAKResponse');
 FHSerialPort.SendByte(NAK);
end;

procedure TASTM1381Transfer.OpenConnect;
var
 Counter: Integer;
 ReturnState: TReturnState;
begin
 Counter := 0;
 repeat
  FHSerialPort.SendByte(ENQ);
  ReturnState := WaitForACK;
  Inc(Counter);
  if ReturnState = rsNAK then
   Sleep(10000);
 until (TransferState = tsEstabilished) or (Counter = 7) or (ReturnState = rsOther);
end;

procedure TASTM1381Transfer.ReadPacket;
begin
 WriteLog('ReadPacket');
 FPrevState := TransferState;
 TransferState := tsWaitForPacket;
end;

procedure TASTM1381Transfer.OnRxChar(Sender: TObject; ResvSize: Integer);
var
 Temp: Byte;
 s: AnsiString;
begin
 FHSerialPort.ReadString(s, ResvSize);
 //if Assigned(OnByteReceived) then
 //  OnByteReceived(s);
 WriteLog(AnsiString('State on Rx: ' + GetEnumName(TypeInfo(TTransferState), Integer(TransferState))));
 if Length(s) = 0 then
 begin
  WriteLog('Rx null');
  Exit;
 end;
 repeat
  Temp := Ord(s[1]);
  case Temp of
   ENQ:
    begin
     WriteLog('Rx ENQ');
     case TransferState of
      tsNeutral, tsTransfer:
       begin
        FLastMessage.Clear;
        FLastPacket := '';
        TransferState := tsEstabilished;
        ACKResponse;
        Delete(s, 1, 1);
       end;
     else
      raise Exception.Create('Wrong state in ENQ!');
      {tsEstabilished:
        begin
          FLastMessage.Clear;
          FLastPacket := '';
          TransferState := tsNeutral;
          NAKResponse;
        end;}
     end;
    end;
   STX:
    begin
     WriteLog('Rx STX');
     case TransferState of
      tsNeutral, tsEstabilished, tsTransfer:
       begin
        FLastPacket := Copy(s, 2, MaxInt);
        s := '';
        ReadPacket;
       end;
     else
      raise Exception.Create('Wrong state in STX!');
     end;
    end;
   EOT:
    begin
     WriteLog('Rx EOT');
     case TransferState of
      tsTransfer, tsWaitForPacket:
       begin
        FLastPacket := '';
        Delete(s, 1, 1);
        TransferState := tsNeutral;
        if Assigned(FOnMessageRecieved) and (LastMessage.Count > 0) then
         FOnMessageRecieved;
       end;
     else
      raise Exception.Create('Wrong state in EOT!');
     end;
    end;
  else
   begin
    WriteLog('Rx ' + ReplReturns(s));
    FLastPacket := FLastPacket + s;
    s := '';
   end;
  end;
  if (TransferState = tsWaitForPacket) and (Pos(C_CR + C_LF, FLastPacket) <> 0) then
   PacketReceived;
 until s = '';
end;

procedure TASTM1381Transfer.PacketReceived;
var
 Temp: Integer;
begin
 WriteLog('PacketReceived');
 case FPrevState of
  tsNeutral: NAKResponse;
  tsEstabilished, tsTransfer:
   begin
    if Checksum(FLastPacket) then
    begin
     Temp := Max(Pos(C_ETX, FLastPacket), Pos(C_ETB, FLastPacket));
     FLastMessage.Add(string(Copy(FLastPacket, 1, Temp)));
     ACKResponse;
    end
    else
     NAKResponse;
    TransferState := tsTransfer;
   end;
 else
  raise Exception.Create('Wrong state in PacketReceived!');
 end;
end;

procedure TASTM1381Transfer.OnError(Sender: TObject; Msg: AnsiString; var Error: integer);
begin
 WriteLog('OnError');
 FLastError := Msg + AnsiString('. Error: ' + SysErrorMessage(Error));
 if Assigned(FOnPortError) then
  FOnPortError;
 Error := 0;
end;

procedure TASTM1381Transfer.SendFrame(const Frame: AnsiString);
var
 Counter: Integer;
 ReturnState: TReturnState;
begin
 {FPushPop.Push(Frame);
 if FSendThread.Suspended then
   FSendThread.Resume}
 WriteLog('SendFrame');
 case TransferState of
  tsEstabilished: TransferState := tsTransfer;
  tsNeutral: OpenConnect;
 else
  raise Exception.Create('Wrong state in SendFrame!');
 end;
 if TransferState = tsNeutral then
  Exit;
 Counter := 0;
 with FHSerialPort do
  repeat
   SendByte(STX);
   SendString(Frame);
   SendByte(CR);
   SendByte(ETX);
   SendWord(CalcCRC(Frame + C_CR + C_ETX));
   SendByte(CR);
   SendByte(LF);
   ReturnState := WaitForACK;
   Inc(Counter);
   if ReturnState <> rsACK {= rsNAK} then
    Sleep(10000);
  until (Counter = 7) or (ReturnState = rsACK);
 if (Counter = 7) and (ReturnState <> rsACK) then
  CloseConnect;
 //FOwner.CloseConnect;
end;

function TASTM1381Transfer.WaitForACK: TReturnState;
var
 s: AnsiString;
begin
 WriteLog('WaitForACK');
 Result := rsOther;
 FHSerialPort.ReadString(s, 1, 10000);
 //if Assigned(OnByteReceived) then
 //  OnByteReceived(s);
 if s = C_ACK then
 begin
  Result := rsACK;
  case TransferState of
   tsNeutral: TransferState := tsEstabilished;
   tsEstabilished: ;
  else
   raise Exception.Create('Wrong state in WaitForACK!');
  end;
 end
 else if s = C_NAK then
  Result := rsNAK;
end;

procedure TASTM1381Transfer.SetTransferState(const Value: TTransferState);
begin
 WriteLog('Set state from: ' + GetEnumName(TypeInfo(TTransferState), Integer(FTransferState)) + ' to: ' +
  GetEnumName(TypeInfo(TTransferState), Integer(Value)));
 FTransferState := Value;
end;

{ TSendThread }

(*constructor TSendThread.Create(AOwner: TASTM1381Transfer);
begin
  inherited Create(True);
  FOwner := AOwner;
end;

procedure TSendThread.Execute;
var
  Counter: Integer;
  ReturnState: TReturnState;
  s: AnsiString;
begin
  repeat
    if FOwner.FPushPop.Count = 0 then
    begin
      Sleep(1000);
      Continue;
    end;
    s := FOwner.FPushPop.Pop;
    WriteLog('SendFrame');
    case FOwner.TransferState of
      tsEstabilished: FOwner.TransferState := tsTransfer;
      tsNeutral: FOwner.OpenConnect;
    else
      raise Exception.Create('Wrong state in SendFrame!');
    end;
    if FOwner.TransferState = tsNeutral then
      Exit;
    Counter := 0;
    with FOwner.FHSerialPort do
      repeat
        SendByte(STX);
        SendString(s);
        SendByte(CR);
        SendByte(ETX);
        SendWord(FOwner.CalcCRC(s + C_CR + C_ETX));
        SendByte(CR);
        SendByte(LF);
        ReturnState := FOwner.WaitForACK;
        Inc(Counter);
        if ReturnState <> rsACK{= rsNAK} then
          Sleep(10000);
      until (Counter = 7) or (ReturnState = rsACK);
    //if (Counter = 7) and (ReturnState <> rsACK) then
    //  FOwner.CloseConnect;
    FOwner.CloseConnect;
  until Terminated;
end;*)

initialization
 WriteLogStrCriticalSection := TCriticalSection.Create;
 CreateLogFile;
 
finalization
 WriteLogStrCriticalSection.Free;
 
end.

