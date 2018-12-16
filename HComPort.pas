unit HComPort;

interface

uses
  Windows, Classes, SysUtils;

type
  TBaudRate = (br110, br300, br600, br1200, br2400, br4800, br9600,
    br14400, br19200, br38400, br56000, br57600, br115200);
  TStopBits = (sbOneStopBit, sbOne5StopBits, sbTwoStopBits);
  TParity = (prNone, prOdd, prEven, prMark, prSpace);
  TFlowControl = (fcNone, fcRtsCts, fcXonXoff, fcBoth);
  TEvent = (evRxChar, evTxEmpty, evRxFlag, evRing, evBreak, evCTS,
    evDSR, evError, evRLSD);
  TEvents = set of TEvent;

  TComRxCharEvent = procedure(Sender: TObject; ResvSize: Integer) of object;
  TComErrorEvent = procedure(Sender: TObject; Msg: AnsiString; var Error: integer) of object;

  THSerialPort = class;

  TComThread = class(TThread)
  private
    Owner: THSerialPort;
    Mask: DWORD;
    StopEvent: THandle;
  protected
    procedure Execute; override;
    procedure DoEvents;
    procedure Stop;
  public
    constructor Create(AOwner: THSerialPort);
    destructor Destroy; override;
  end;

  THSerialPort = class
  private
    ComHandle: THandle;
    EventThread: TComThread;
    FConnected: Boolean;
    FBaudRate: TBaudRate;
    FParity: TParity;
    FStopBits: TStopBits;
    FFlowControl: TFlowControl;
    FDataBits: Byte;
    FEvents: TEvents;
    FEnableDTR: Boolean;
    FWriteBufSize: Integer;
    FReadBufSize: Integer;
    FActiveDCD: Boolean;
    FOnRxChar: TComRxCharEvent;
    FOnTxEmpty: TNotifyEvent;
    FOnBreak: TNotifyEvent;
    FOnRing: TNotifyEvent;
    FOnCTS: TNotifyEvent;
    FOnDSR: TNotifyEvent;
    FOnDCD: TNotifyEvent;
    FOnError: TComErrorEvent;
    FOnRxFlag: TNotifyEvent;
    FOnOpen: TNotifyEvent;
    FOnClose: TNotifyEvent;
    FPort: string;
    procedure SetDataBits(Value: Byte);
    procedure DoOnRxChar;
    procedure DoOnTxEmpty;
    procedure DoOnBreak;
    procedure DoOnRing;
    procedure DoOnRxFlag;
    procedure DoOnCTS;
    procedure DoOnDSR;
    procedure DoOnError(Msg: AnsiString; Error: integer);
    procedure DoOnDCD;
    function CheckActiveDCD: Boolean;
    procedure CreateHandle;
    procedure DestroyHandle;
    procedure SetupState;
    function ValidHandle: Boolean;
    procedure InitSerialPort;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Open;
    procedure Close;
    function InQue: Integer;
    function OutQue: Integer;
    function ActiveCTS: Boolean;
    function ActiveDSR: Boolean;
    function ActiveDCD: Boolean;
    function ActiveRing: Boolean;
    function Write(var Buffer; Count: Integer): Integer;
    function SendByte(b: byte): Integer;
    function SendWord(w: word): Integer;
    function SendString(Str: AnsiString): Integer;
    function WriteString(Str: AnsiString): Integer;
    function Read(var Buffer; Count: Integer; TimeOut: DWORD): Integer;
    function ReadString(var Str: AnsiString; Count: Integer; TimeOut: DWORD = INFINITE): Integer;
    procedure PurgeIn;
    procedure PurgeOut;
    function GetComHandle: THandle;
    property Connected: Boolean read FConnected;
    property BaudRate: TBaudRate read FBaudRate write FBaudRate;
    property Port: string read FPort write FPort;
    property Parity: TParity read FParity write FParity;
    property StopBits: TStopBits read FStopBits write FStopBits;
    property FlowControl: TFlowControl read FFlowControl write FFlowControl;
    property DataBits: Byte read FDataBits write SetDataBits;
    property Events: TEvents read FEvents write FEvents;
    property EnableDTR: Boolean read FEnableDTR write FEnableDTR;
    property WriteBufSize: Integer read FWriteBufSize write FWriteBufSize;
    property ReadBufSize: Integer read FReadBufSize write FReadBufSize;
    property OnRxChar: TComRxCharEvent read FOnRxChar write FOnRxChar;
    property OnTxEmpty: TNotifyEvent read FOnTxEmpty write FOnTxEmpty;
    property OnBreak: TNotifyEvent read FOnBreak write FOnBreak;
    property OnRing: TNotifyEvent read FOnRing write FOnRing;
    property OnCTS: TNotifyEvent read FOnCTS write FOnCTS;
    property OnDSR: TNotifyEvent read FOnDSR write FOnDSR;
    property OnDCD: TNotifyEvent read FOnDCD write FOnDCD;
    property OnRxFlag: TNotifyEvent read FOnRxFlag write FOnRxFlag;
    property OnError: TComErrorEvent read FOnError write FOnError;
    property OnOpen: TNotifyEvent read FOnOpen write FOnOpen;
    property OnClose: TNotifyEvent read FOnClose write FOnClose;
  end;

  EComError = class(Exception);

implementation

const
  dcb_Binary = $00000001;
  dcb_Parity = $00000002;
  dcb_OutxCtsFlow = $00000004;
  dcb_OutxDsrFlow = $00000008;
  dcb_DtrControl = $00000030;
  dcb_DsrSensivity = $00000040;
  dcb_TXContinueOnXOff = $00000080;
  dcb_OutX = $00000100;
  dcb_InX = $00000200;
  dcb_ErrorChar = $00000400;
  dcb_Null = $00000800;
  dcb_RtsControl = $00003000;
  dcb_AbortOnError = $00004000;

  // Component code

constructor TComThread.Create(AOwner: THSerialPort);
var
  AMask: Integer;
begin
  inherited Create(True);
  StopEvent := CreateEvent(nil, True, False, nil);
  Owner := AOwner;
  AMask := 0;
  if evRxChar in Owner.FEvents then
    AMask := AMask or EV_RXCHAR;
  if evRxFlag in Owner.FEvents then
    AMask := AMask or EV_RXFLAG;
  if evTxEmpty in Owner.FEvents then
    AMask := AMask or EV_TXEMPTY;
  if evRing in Owner.FEvents then
    AMask := AMask or EV_RING;
  if evCTS in Owner.FEvents then
    AMask := AMask or EV_CTS;
  if evDSR in Owner.FEvents then
    AMask := AMask or EV_DSR;
  if evRLSD in Owner.FEvents then
    AMask := AMask or EV_RLSD;
  if evError in Owner.FEvents then
    AMask := AMask or EV_ERR;
  if evBreak in Owner.FEvents then
    AMask := AMask or EV_BREAK;
  SetCommMask(Owner.ComHandle, AMask);
  Resume;
end;

procedure TComThread.Execute;
var
  EventHandles: array[0..1] of THandle;
  Overlapped: TOverlapped;
  dwSignaled, BytesTrans: DWORD;
begin
  FillChar(Overlapped, SizeOf(Overlapped), 0);
  Overlapped.hEvent := CreateEvent(nil, True, True, nil);
  EventHandles[0] := StopEvent;
  EventHandles[1] := Overlapped.hEvent;
  repeat
    WaitCommEvent(Owner.ComHandle, Mask, @Overlapped);
    dwSignaled := WaitForMultipleObjects(2, @EventHandles, False, INFINITE);
    case dwSignaled of
      WAIT_OBJECT_0: Break;
      WAIT_OBJECT_0 + 1: if GetOverlappedResult(Owner.ComHandle, Overlapped, BytesTrans, False) then
         {Synchronize(}DoEvents{)};
    else
      Break;
    end;
  until False;
  Owner.PurgeIn;
  Owner.PurgeOut;
  CloseHandle(Overlapped.hEvent);
  CloseHandle(StopEvent);
end;

procedure TComThread.Stop;
begin
  SetEvent(StopEvent);
end;

destructor TComThread.Destroy;
begin
  Stop;
  inherited Destroy;
end;

procedure TComThread.DoEvents;
begin
  if (EV_RXCHAR and Mask) > 0 then
    Owner.DoOnRxChar;
  if (EV_TXEMPTY and Mask) > 0 then
    Owner.DoOnTxEmpty;
  if (EV_BREAK and Mask) > 0 then
    Owner.DoOnBreak;
  if (EV_RING and Mask) > 0 then
    Owner.DoOnRing;
  if (EV_CTS and Mask) > 0 then
    Owner.DoOnCTS;
  if (EV_DSR and Mask) > 0 then
    Owner.DoOnDSR;
  if (EV_RXFLAG and Mask) > 0 then
    Owner.DoOnRxFlag;
  if (EV_RLSD and Mask) > 0 then
    Owner.DoOnDCD;
  if (EV_ERR and Mask) > 0 then
    Owner.DoOnError('Communication Error', GetLastError);
end;

constructor THSerialPort.Create;
begin
  inherited;
  FConnected := False;
  FBaudRate := br9600;
  FParity := prNone;
  FPort := 'COM1';
  FStopBits := sbOneStopBit;
  FDataBits := 8;
  FEvents := [evRxChar, evTxEmpty, evRxFlag, evRing, evBreak, evCTS, evDSR, evError, evRLSD];
  FEnableDTR := True;
  FWriteBufSize := 20480;
  FReadBufSize := 20480;
  ComHandle := INVALID_HANDLE_VALUE;
  FActiveDCD := False;
end;

destructor THSerialPort.Destroy;
begin
  Close;
  inherited;
end;

procedure THSerialPort.CreateHandle;
begin
  ComHandle := CreateFile(PChar(FPort), GENERIC_READ or GENERIC_WRITE,
    0, nil, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL or FILE_FLAG_OVERLAPPED, 0);
  if ValidHandle then
    InitSerialPort;
  if not ValidHandle then
    DoOnError('Unable to open com port ', GetLastError);
end;

procedure THSerialPort.DestroyHandle;
begin
  if ValidHandle then
  begin
    PurgeComm(ComHandle, PURGE_TXABORT + PURGE_RXABORT);
    CloseHandle(ComHandle);
  end;
  if FActiveDCD then
  begin
    FActiveDCD := False;
    if Assigned(FOnDCD) then
      FOnDCD(Self);
  end;
  ComHandle := INVALID_HANDLE_VALUE;
end;

function THSerialPort.ValidHandle: Boolean;
begin
  Result := ComHandle <> INVALID_HANDLE_VALUE;
end;

procedure THSerialPort.Open;
begin
  Close;
  CreateHandle;
end;

procedure THSerialPort.Close;
begin
  if FConnected then
  begin
    //EventThread.FreeOnTerminate := True;
    //EventThread.Terminate;
    //EventThread.WaitFor;
    EventThread.Free;
    DestroyHandle;
    FConnected := False;
    if Assigned(FOnClose) then
      FOnClose(Self);
  end;
end;

procedure THSerialPort.SetupState;
var
  DCB: TDCB;
  Timeouts: TCommTimeouts;
begin
  FillChar(DCB, SizeOf(DCB), 0);

  DCB.DCBlength := SizeOf(DCB);

  GetCommState(ComHandle, DCB);

  case FBaudRate of
    br110: DCB.BaudRate := CBR_110;
    br300: DCB.BaudRate := CBR_300;
    br600: DCB.BaudRate := CBR_600;
    br1200: DCB.BaudRate := CBR_1200;
    br2400: DCB.BaudRate := CBR_2400;
    br4800: DCB.BaudRate := CBR_4800;
    br9600: DCB.BaudRate := CBR_9600;
    br14400: DCB.BaudRate := CBR_14400;
    br19200: DCB.BaudRate := CBR_19200;
    br38400: DCB.BaudRate := CBR_38400;
    br56000: DCB.BaudRate := CBR_56000;
    br57600: DCB.BaudRate := CBR_57600;
    br115200: DCB.BaudRate := CBR_115200;
  end;
  DCB.ByteSize := FDataBits;
  case FParity of
    prNone: DCB.Parity := NOPARITY;
    prOdd: DCB.Parity := ODDPARITY;
    prEven: DCB.Parity := EVENPARITY;
    prMark: DCB.Parity := MARKPARITY;
    prSpace: DCB.Parity := SPACEPARITY;
  end;
  case FStopBits of
    sbOneStopBit: DCB.StopBits := ONESTOPBIT;
    sbOne5StopBits: DCB.StopBits := ONE5STOPBITS;
    sbTwoStopBits: DCB.StopBits := TWOSTOPBITS;
  end;
  DCB.EvtChar := #0;

  DCB.XonChar := #17;
  DCB.XoffChar := #19;
  DCB.XonLim := 0;
  DCB.XoffLim := 0;
  DCB.Flags := 4113; //dcb_Binary or dcb_DtrControl or dcb_RtsControl or dcb_Parity;

  { DCB.XonLim:=FWriteBufSize div 4;
   DCB.XoffLim:=1;

   DCB.Flags:=DCB.Flags or dcb_Binary;
   if FEnableDTR then DCB.Flags:=DCB.Flags or (dcb_DtrControl and (DTR_CONTROL_ENABLE shl 4));

   case FFlowControl of
     fcRtsCts: DCB.Flags:=DCB.Flags or dcb_OutxCtsFlow or (dcb_RtsControl and (RTS_CONTROL_HANDSHAKE shl 12));
     fcXonXoff: DCB.Flags:=DCB.Flags or dcb_OutX or dcb_InX;
     fcBoth: DCB.Flags:=DCB.Flags or dcb_OutX or dcb_InX or dcb_OutxCtsFlow or (dcb_RtsControl and (RTS_CONTROL_HANDSHAKE shl 12));
   end; }
  if not SetCommState(ComHandle, DCB) then
    DoOnError('Unable to set com state', GetLastError);

  if not GetCommTimeouts(ComHandle, Timeouts) then
    DoOnError('Unable to set com Timeout', GetLastError);

  timeouts.ReadIntervalTimeout := $FFFFFFFF;
  timeouts.ReadTotalTimeoutMultiplier := CBR_56000;
  case FBaudRate of
    br2400: timeouts.WriteTotalTimeoutMultiplier := CBR_2400;
    br4800: timeouts.WriteTotalTimeoutMultiplier := CBR_4800;
    br9600: timeouts.WriteTotalTimeoutMultiplier := CBR_9600;
    br19200: timeouts.WriteTotalTimeoutMultiplier := CBR_19200;
    br38400: timeouts.WriteTotalTimeoutMultiplier := CBR_38400;
    br57600: timeouts.WriteTotalTimeoutMultiplier := CBR_57600;
    br115200: timeouts.WriteTotalTimeoutMultiplier := CBR_128000;
  end;

  if not SetCommTimeouts(ComHandle, Timeouts) then
    DoOnError('Unable to set com Timeout', GetLastError);

  EscapeCommFunction(ComHandle, SETDTR);

  if not SetupComm(ComHandle, FReadBufSize, FWriteBufSize) then
    DoOnError('Unable to set com', GetLastError);
end;

function THSerialPort.InQue: Integer;
var
  Errors: DWORD;
  ComStat: TComStat;
begin
  if not ClearCommError(ComHandle, Errors, @ComStat) then
    DoOnError('Unable to read com InQue', GetLastError);
  Result := ComStat.cbInQue;
end;

function THSerialPort.OutQue: Integer;
var
  Errors: DWORD;
  ComStat: TComStat;
begin
  if not ClearCommError(ComHandle, Errors, @ComStat) then
    DoOnError('Unable to read com OutQue', GetLastError);
  Result := ComStat.cbOutQue;
end;

function THSerialPort.CheckActiveDCD: Boolean;
var
  dwS: DWORD;
begin
  if not GetCommModemStatus(ComHandle, dwS) then
    DoOnError('Unable to read com status', GetLastError);
  Result := (dws and MS_RLSD_ON <> 0);
end;

function THSerialPort.ActiveCTS: Boolean;
var
  dwS: DWORD;
begin
  if not GetCommModemStatus(ComHandle, dwS) then
    DoOnError('Unable to read com status', GetLastError);
  Result := (dws and MS_CTS_ON <> 0);
end;

function THSerialPort.ActiveDSR: Boolean;
var
  dwS: DWORD;
begin
  if not GetCommModemStatus(ComHandle, dwS) then
    DoOnError('Unable to read com status', GetLastError);
  Result := (dws and MS_DSR_ON <> 0);
end;

function THSerialPort.ActiveRing: Boolean;
var
  dwS: DWORD;
begin
  if not GetCommModemStatus(ComHandle, dwS) then
    DoOnError('Unable to read com status', GetLastError);
  Result := (dws and MS_RING_ON <> 0);
end;

function THSerialPort.Write(var Buffer; Count: Integer): Integer;
var
  Overlapped: TOverlapped;
  BytesWritten: DWORD;
begin
  FillChar(Overlapped, SizeOf(Overlapped), 0);
  Overlapped.hEvent := CreateEvent(nil, True, True, nil);
  WriteFile(ComHandle, Buffer, Count, BytesWritten, @Overlapped {nil});
  WaitForSingleObject(Overlapped.hEvent, INFINITE {10000});
  if not GetOverlappedResult(ComHandle, Overlapped, BytesWritten, False) then
    DoOnError('Unable to write to port', GetLastError);
  CloseHandle(Overlapped.hEvent);
  Result := BytesWritten;
end;

function THSerialPort.WriteString(Str: AnsiString): Integer;
begin
  Result := Write(Str[1], Length(Str));
end;

function THSerialPort.Read(var Buffer; Count: Integer; TimeOut: DWORD): Integer;
var
  Overlapped: TOverlapped;
  BytesRead: DWORD;
begin
  FillChar(Overlapped, SizeOf(Overlapped), 0);
  Overlapped.hEvent := CreateEvent(nil, False, False, nil);
  ReadFile(ComHandle, Buffer, Count, BytesRead, @Overlapped);
  WaitForSingleObject(Overlapped.hEvent, TimeOut);
  if not GetOverlappedResult(ComHandle, Overlapped, BytesRead, False) then
    DoOnError('Unable to read from port', GetLastError);
  CloseHandle(Overlapped.hEvent);
  Result := BytesRead;
end;

function THSerialPort.ReadString(var Str: AnsiString; Count: Integer; TimeOut: DWORD = INFINITE): Integer;
begin
  SetLength(Str, Count);
  Result := Read(Str[1], Count, TimeOut);
  SetLength(Str, Result);
end;

procedure THSerialPort.PurgeIn;
begin
  if not PurgeComm(ComHandle, PURGE_RXABORT or PURGE_RXCLEAR) then
    DoOnError('Unable to purge com', GetLastError);
end;

procedure THSerialPort.PurgeOut;
begin
  if not PurgeComm(ComHandle, PURGE_TXABORT or PURGE_TXCLEAR) then
    DoOnError('Unable to purge com', GetLastError);
end;

function THSerialPort.GetComHandle: THandle;
begin
  Result := ComHandle;
end;

function THSerialPort.SendByte(b: byte): Integer;
var
  F: TextFile;
  function TactCounter: Int64;
  asm
 db $0F,$31
  end;
begin
  Result := Write(b, 1);
  Exit;
  AssignFile(f, 'log.txt');
  if FileExists('log.txt') then
    Append(f)
  else
    Rewrite(f);
  WriteLn(f, 'Send on ' + IntToStr(TactCounter));
  WriteLn(f, Char(b));
  CloseFile(f);
end;

function THSerialPort.SendString(Str: AnsiString): Integer;
var
  F: TextFile;
  function TactCounter: Int64;
  asm
 db $0F,$31
  end;
begin
  Result := Write(Str[1], Length(Str));
  Exit;
  AssignFile(f, 'log.txt');
  if FileExists('log.txt') then
    Append(f)
  else
    Rewrite(f);
  WriteLn(f, 'Send on ' + IntToStr(TactCounter));
  WriteLn(f, Str);
  CloseFile(f);
end;

function THSerialPort.SendWord(w: word): Integer;
var
  F: TextFile;
  function TactCounter: Int64;
  asm
 db $0F,$31
  end;
begin
  Result := Write(w, 2);
  Exit;
  AssignFile(f, 'log.txt');
  if FileExists('log.txt') then
    Append(f)
  else
    Rewrite(f);
  WriteLn(f, 'Send on ' + IntToStr(TactCounter));
  WriteLn(f, w);
  CloseFile(f);
end;

procedure THSerialPort.SetDataBits(Value: Byte);
begin
  if Value <> FDataBits then
    if Value > 8 then
      FDataBits := 8
    else if Value < 5 then
      FDataBits := 5
    else
      FDataBits := Value;
end;

procedure THSerialPort.DoOnRxChar;
begin
  if Assigned(FOnRxChar) then
    FOnRxChar(Self, InQue);
end;

procedure THSerialPort.DoOnBreak;
begin
  if Assigned(FOnBreak) then
    FOnBreak(Self);
end;

procedure THSerialPort.DoOnRing;
begin
  if Assigned(FOnRing) then
    FOnRing(Self);
end;

procedure THSerialPort.DoOnTxEmpty;
begin
  if Assigned(FOnTxEmpty) then
    FOnTxEmpty(Self);
end;

procedure THSerialPort.DoOnCTS;
begin
  if Assigned(FOnCTS) then
    FOnCTS(Self);
end;

procedure THSerialPort.DoOnDSR;
begin
  if Assigned(FOnDSR) then
    FOnDSR(Self);
end;

procedure THSerialPort.DoOnDCD;
begin
  FActiveDCD := CheckActiveDCD;
  if Assigned(FOnDCD) then
    FOnDCD(Self);
end;

procedure THSerialPort.DoOnError(Msg: AnsiString; Error: integer);
begin
  if Assigned(FOnError) then
    FOnError(Self, Msg, Error);
  if Error <> 0 then
    EComError.Create(string(Msg) + Format('Error Number : %d', [Error]));
end;

procedure THSerialPort.DoOnRxFlag;
begin
  if Assigned(FOnRxFlag) then
    FOnRxFlag(Self);
end;

function THSerialPort.ActiveDCD: Boolean;
begin
  Result := (FConnected) and (FActiveDCD);
end;

procedure THSerialPort.InitSerialPort;
begin
  SetupState;
  EventThread := TComThread.Create(Self);
  FConnected := True;
  FActiveDCD := CheckActiveDCD;
  if Assigned(FOnOpen) then
    FOnOpen(Self);
end;

end.

