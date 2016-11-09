unit uSerialPort;

interface

uses
  Windows, SysUtils, Classes, IniFiles, uLogFile;

type
  TtSerailPort = class(TComponent)
  private
    FRequeryDataInterval: Cardinal;
    FComPort: Integer;
    FBaudRate: Integer;
    FTimeOuts: COMMTIMEOUTS;
    FParity: Integer;
    FStopBits: Integer;
    FOutQueue: Integer;
    FInQueue: Integer;
    FByteSize: Integer;
    FBytes: Cardinal;
    FUseSerialPort: Boolean;
    FEndChar: Word;
    FLogFile: TtLogFile;
    FWriteToLog: Boolean;
    procedure SetBaudRate(const Value: Integer);
    procedure SetBytes(const Value: Cardinal);
    procedure SetByteSize(const Value: Integer);
    procedure SetComPort(const Value: Integer);
    procedure SetInQueue(const Value: Integer);
    procedure SetOutQueue(const Value: Integer);
    procedure SetParity(const Value: Integer);
    procedure SetRequeryDataInterval(const Value: Cardinal);
    procedure SetStopBits(const Value: Integer);
    procedure SetTimeOuts(const Value: COMMTIMEOUTS);
    procedure SetUseSerialPort(const Value: Boolean);
    procedure SetEndChar(const Value: Word);
    procedure SetLogFile(const Value: TtLogFile);
    procedure SetWriteToLog(const Value: Boolean);
    { Private declarations }
  protected
    { Protected declarations }
  public
    { Public declarations }
    FTheStruct:TCOMSTAT;
    FErrors: Cardinal;
    FDCB:TDCB;
    FCom:Cardinal;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy(); override;
      {Функция для открытия ComPort}
    function OpenSerialPort: Boolean;
      {Функция для закрытия ComPort}
    function CloseSerialPort: Boolean;
      {Функция для чтения из ComPort}
    function Read: String;
      {Функция для записи в ComPort}
    function Write(AString: String): Boolean;
      {Функция для очистки ComPort}
    function Clear: Boolean;
  published
    { Published declarations }
    //'Номер Com Port
    property ComPort: Integer read FComPort write SetComPort;
    //'Скорость считывания с Com Port считывателя магнитных карт.',
    property BaudRate: Integer read FBaudRate write SetBaudRate;
    //'Количество байт считывания с Com Port считывателя магнитных карт.',
    property ByteSize: Integer read FByteSize write SetByteSize;
    //'Схема считывания с Com Port считывателя магнитных карт.',
    property Parity: Integer read FParity write SetParity;
    //'Количество стоповых битов считывания с Com Port считывателя магнитных карт.',
    property StopBits: Integer read FStopBits write SetStopBits;
    //'Количество считываемых байтов считываемых с Com Port считывателя магнитных карт.',
    property InQueue: Integer read FInQueue write SetInQueue;
    //'Количество записываемых байтов считываемых с Com Port считывателя магнитных карт.',
    property OutQueue: Integer read FOutQueue write SetOutQueue;
    // 'Интервал с которым данные будет считываться программой с Com Port 1000 = 1 сек.',
    property RequeryDataInterval: Cardinal read FRequeryDataInterval write SetRequeryDataInterval;

    property Bytes:Cardinal read FBytes write SetBytes;
    property TimeOuts: COMMTIMEOUTS read FTimeOuts write SetTimeOuts;
      {Отключает чтенеи из SerialPort }
    property UseSerialPort: Boolean read FUseSerialPort write SetUseSerialPort;
      {Завершающий символ}
    property EndChar: Word  read FEndChar write SetEndChar;
      {Лог файл для логирования работы порта }
    property LogFile: TtLogFile read FLogFile write SetLogFile;
      {Логоировать раоту компонента}
    property WriteToLog: Boolean read FWriteToLog write SetWriteToLog;
  end;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('CarsComponets', [TtSerailPort]);
end;

{ TtSerailPort }

function TtSerailPort.Clear: Boolean;
const
  cErrorWhilePurgeSerialPort = 'Ошибка при очистке Com Port.';
begin
  if not UseSerialPort then begin
    Result:= True;
    Exit;
  end;
  if not PurgeComm(FCom, PURGE_TXCLEAR or PURGE_RXCLEAR) then begin
    raise Exception.Create(cErrorWhilePurgeSerialPort+#10#13+SysErrorMessage(GetLastError));
    Result:= False;
  end
  else
    Result:= True;
end;

function TtSerailPort.CloseSerialPort: Boolean;
begin
  if not UseSerialPort then begin
    Result:= True;
    Exit;
  end;
  CloseHandle(FCom);
  Result:= True;
end;

constructor TtSerailPort.Create(AOwner: TComponent);
begin
  inherited;
  with Self do
  begin
    LogFile:= TtLogFile.Create(Self);
  end;
end;

destructor TtSerailPort.Destroy;
begin
  LogFile.Free;
  inherited;
end;

function TtSerailPort.OpenSerialPort: Boolean;
const
  cErrorCanNoOpenComPort = 'Ошибка при инициальзации считывателя магнитных карт!';
  cErrorCanNotSetState = 'Невозможно установить параметры работы Com Port!';
begin
  if not UseSerialPort then begin
    Result:= True;
    Exit;
  end;

  Result:= True;
  FCom:=CreateFile(PChar('COM'+IntToStr(ComPort)),GENERIC_READ or GENERIC_WRITE,
    0,nil,OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,0);
  if FCom=INVALID_HANDLE_VALUE then
  begin
    CloseHandle(FCom);
    raise Exception.Create(cErrorCanNoOpenComPort+#10#13+SysErrorMessage(GetLastError));
    Result:= False;
    exit;
  end;
  FDCB.DCBlength:= sizeof(tdcb);
  SetupComm(FCom,FInQueue,FOutQueue);
  GetCommState(FCom,FDCB);
  with FDCB do
  begin
    BaudRate:=FBaudRate;
    ByteSize:=FByteSize;
    Parity:=FParity;
    StopBits:=FStopBits;
  end;
  if not SetCommState(FCom,FDCB) then
  begin
    CloseHandle(FCom);
    raise Exception.Create(cErrorCanNotSetState);
    Result:= False;
    exit;
  end;
end;

function TtSerailPort.Read: String;
const
  cAlertSkipRead = 'Считываение данных из SerialPort пропущено.';
  cPrintRead = 'Считываем данные из SerialPort.';
var
  lBuffer: string;
begin
  LogFile.Add(llINFO, cPrintRead);
  if not UseSerialPort then begin
    LogFile.Add(llINFO, cAlertSkipRead);
    Result:= '';
    Exit;
  end;
  Result:= '';
  ClearCommError(FCom,FErrors,@FTheStruct);
  repeat
    if FTheStruct.cbInQue>0 then
    begin //что-то пришло
        SetLength(lBuffer,FTheStruct.cbInQue);
        ReadFile(FCom,lBuffer[1],FTheStruct.cbInQue,FBytes,nil);
        SetLength(lBuffer,Bytes);
        //прочитано: buffer
        LogFile.Add(llINFO, 'lBuffer = '+lBuffer);
        Result:= Result+lBuffer;
    end
    else begin
      Result:= 'Nothing';
      Break;
    end;
  until (Result[Length(Result)] =  Char(EndChar));
  LogFile.Add(llINFO, 'Result = '+Result);
end;

procedure TtSerailPort.SetBaudRate(const Value: Integer);
begin
  FBaudRate := Value;
end;

procedure TtSerailPort.SetBytes(const Value: Cardinal);
begin
  FBytes := Value;
end;

procedure TtSerailPort.SetByteSize(const Value: Integer);
begin
  FByteSize := Value;
end;

procedure TtSerailPort.SetComPort(const Value: Integer);
begin
  FComPort := Value;
end;

procedure TtSerailPort.SetEndChar(const Value: Word);
begin
  FEndChar := Value;
end;

procedure TtSerailPort.SetInQueue(const Value: Integer);
begin
  FInQueue := Value;
end;

procedure TtSerailPort.SetLogFile(const Value: TtLogFile);
begin
  FLogFile := Value;
end;

procedure TtSerailPort.SetOutQueue(const Value: Integer);
begin
  FOutQueue := Value;
end;

procedure TtSerailPort.SetParity(const Value: Integer);
begin
  FParity := Value;
end;

procedure TtSerailPort.SetRequeryDataInterval(const Value: Cardinal);
begin
  FRequeryDataInterval := Value;
end;

procedure TtSerailPort.SetStopBits(const Value: Integer);
begin
  FStopBits := Value;
end;

procedure TtSerailPort.SetTimeOuts(const Value: COMMTIMEOUTS);
begin
  FTimeOuts := Value;
end;

procedure TtSerailPort.SetUseSerialPort(const Value: Boolean);
begin
  FUseSerialPort := Value;
end;

procedure TtSerailPort.SetWriteToLog(const Value: Boolean);
begin
  if FWriteToLog <> Value then begin
    FWriteToLog := Value;
    LogFile.Active:= FWriteToLog;
  end;
end;

function TtSerailPort.Write(AString: String): Boolean;
const
  cErrorWhileWriteToComport = 'Ошибка при записи в Com Port';
begin
  if not UseSerialPort then begin
    Result:= True;
    Exit;
  end;
  if not WriteFile(FCom,AString,FTheStruct.cbOutQue,FBytes,nil) then begin
    raise Exception.Create(cErrorWhileWriteToComport+#10#13+SysErrorMessage(GetLastError));
    Result:= False;
  end
  else Result:= True;

end;

end.
