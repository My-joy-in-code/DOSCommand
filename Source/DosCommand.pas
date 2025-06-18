{
  this component let you execute a dos program (exe, com or batch file) and catch
  the ouput in order to put it in a memo or in a listbox, ...
  you can also send inputs.
  the cool thing of this component is that you do not need to wait the end of
  the program to get back the output. it comes line by line.


  *********************************************************************
  ** maxime_collomb@yahoo.fr                                         **
  **                                                                 **
  **   for this component, iCount just translated C code             **
  ** from Community.borland.com                                      **
  ** (http://www.vmlinux.org/jakov/community.borland.com/10387.html) **
  **                                                                 **
  **   if you have a good idea of improvement, please                **
  ** let me know (maxime_collomb@yahoo.fr).                          **
  **   if you improve this component, please send me a Copy          **
  ** so i can put it on www.torry.net.                               **
  *********************************************************************

  History :
  ---------

  06-19-2025 : changed by samso->www.delphipraxis.net
  - revert EOL-Detection in DoReadLine
  - Environment handling refactored
  05-31-2025 : changed by samso->www.delphipraxis.net
  - recover WaitFor in TDosCommand.Stop to avoid memory leaks during application termination
  05-30-2025 : changed by samso->www.delphipraxis.net
  - added GetOEMCodepage
  - Char decoding and encoding uses now by default the Windows OEM Codepage
  05-29-2025 : changed by samso->www.delphipraxis.net
  - DoReadLine refactored
  05-27-2025 : changed by samso->www.delphipraxis.net
  - asserts removed, OSExceptions are used instead
  - Code added to read the last bytes of stdout after the process has ended
  05-24-2025 : changed by samso->www.delphipraxis.net
  - added property "Codepage" to set default text encoding and decoding in console
  - Remove TCharDecoding and TCharEncoding - use TEncoding instead
  - changed property TDosCommand.Lines: Now filled after excuting of the process
  - TDosThread can used standalone
  03-19-2024 : changed by samso->www.delphipraxis.net
  - TTimeoutCalculator instead TProcessTimer to avoid the use of timers in TThread
  - Timeout as floatingpoint for timeouts<1s
  06-21-2011 : version 2.03 (by sirius in www.delphi-treff.de || www.delphipraxis.net)
  (marked sirius2)
  - added property "current directory" to set for child process
  - added possibility to add environment variables (see properties)
  - deleted possible memory leaks
  - removed class TTimer (not threadsafe) from TProcesstimer (added global var TimerInstances and stuff ->threadsafe)
  - try to prepare for unicode  !!! not tested
  - added event OnCharDecoding (for unicode console output); no need for ANSI Text in console
  - added event OnCharEncoding (for unicode console input); no need for ANSI Text in console
  - removed bug causing EInvalidpointer with unhandled FatalException in TDosThread (dont free FatalException object)
  - added MB_TASKMODAL to Error Messagebox of Thread
  - create pipehandles like in Article ID: 190351 of Microsoft knowledge base

  05-11-2009 : version 2.02 (by sirius in www.delphi-treff.de || www.delphipraxis.net)
  - added synchronisation (see later)
  - deleted FOwner in TDosCommand (not needed)
  - added TInputlines (sync of InputLines)
  - added critical section and locked functions to Processtimer
  - added TDosCommand.ThreadTerminate
  - moved FTimer.Ending to Thread.onTerminate
  - reraised Excpetion from Thread in onTerminate
  - added try-finally to whole Thread.Execute method
  - added Synchronize to FLines and FOutPutLines
  - added Synchronize to NewLine event
  - added event for every New char
  - added TReadPipe(Thread) for waiting to Pipe (blocking ReadFile)
  - added TSyncstring to Synchronize string transfer (from TReadPipe)
  - restructured FExecute -> Execute (New methods)
  - added WaitForMultipleObjects and stuff for less Processing-Time

  13-06-2003 : version 2.01
  - Added exception when executing with empty CommandLine
  - Added IsRunning property to check if a command is currently
  running
  18-05-2001 : version 2.0
  - Now, catching the beginning of a line is allowed (usefull if the
  prog ask for an entry) => the method OnNewLine is modified
  - Now can send inputs
  - Add a couple of FreeMem for sa & sd [thanks Gary H. Blaikie]
  07-05-2001 : version 1.2
  - Sleep(1) is added to give others processes a chance
  [thanks Hans-Georg Rickers]
  - the loop that catch the outputs has been re-writen by
  Hans-Georg Rickers => no more broken lines
  30-04-2001 : version 1.1
  - function IsWinNT() is changed to
  (Win32Platform = VER_PLATFORM_WIN32_NT) [thanks Marc Scheuner]
  - empty lines appear in the redirected output
  - property OutputLines is added to redirect output directly to a
  memo, richedit, listbox, ... [thanks Jean-Fabien Connault]
  - a timer is added to offer the possibility of ending the process
  after XXX sec. after the beginning or YYY sec after the last
  output [thanks Jean-Fabien Connault]
  - managing process priorities flags in the CreateProcess
  thing [thanks Jean-Fabien Connault]
  20-04-2001 : version 1.0 on www.torry.net
  *******************************************************************
  How to use it :
  ---------------
  - just put the line of command in the property 'CommandLine'
  - execute the process with the method 'Execute'
  - if you want to stop the process before it has ended, use the method 'Stop'
  - if you want the process to stop by itself after XXX sec of activity,
  use the property 'MaxTimeAfterBeginning'
  - if you want the process to stop after XXX sec without an output,
  use the property 'MaxTimeAfterLastOutput'
  - to directly redirect outputs to a memo or a richedit, ...
  use the property 'OutputLines'
  (DosCommand1.OutputLines := Memo1.Lines;)
  - you can access all the outputs of the last command with the property 'Lines'
  - you can change the priority of the process with the property 'Priority'
  value of Priority must be in [HIGH_PRIORITY_CLASS, IDLE_PRIORITY_CLASS,
  NORMAL_PRIORITY_CLASS, REALTIME_PRIORITY_CLASS]
  - you can have an event for each New line and for the end of the process
  with the events 'procedure OnNewLine(Sender: TObject; NewLine: string;
  OutputType: TOutputType);' and 'procedure OnTerminated(Sender: TObject);'
  - you can send inputs to the dos process with 'SendLine(Value: string;
  Eol: Boolean);'. Eol is here to determine if the program have to add a
  CR/LF at the end of the string.
  *******************************************************************
  How to call a dos function (win 9x/Me) :
  ----------------------------------------

  Example : Make a dir :
  ----------------------
  - if you want to get the Result of a 'c:\dir /o:gen /l c:\windows\*.txt'
  for example, you need to make a batch file
  --the batch file : c:\mydir.bat
  @echo off
  dir /o:gen /l %1
  rem eof
  --in your code
  DosCommand.CommandLine := 'c:\mydir.bat c:\windows\*.txt';
  DosCommand.Execute;

  Example : Format a disk (win 9x/Me) :
  -------------------------
  --a batch file : c:\myformat.bat
  @echo off
  format %1
  rem eof
  --in your code
  var diskname: string;
  --
  DosCommand1.CommandLine := 'c:\myformat.bat a:';
  DosCommand1.Execute; //launch format process
  DosCommand1.SendLine('', True); //equivalent to press enter key
  DiskName := 'test';
  DosCommand1.SendLine(DiskName, True); //enter the name of the volume
  ******************************************************************* }
unit DosCommand;

interface

uses
  System.SysUtils, System.Classes, System.SyncObjs, Winapi.Windows, Winapi.Messages, registry;

const
  DefaultOEMCodepage = 850;  // if not found in Registry

type
  EDosCommand = class(Exception); // MK: 20030613

  ECreatePipeError = class(Exception); // exception raised when a pipe cannot be created
  ECreateProcessError = class(Exception); // exception raised when the process cannot be created
  EProcessTimer = class(Exception); // exception raised by TProcessTimer

  TOutputType = (otEntireLine, otBeginningOfLine); // to know if the newline is finished.
  TEndStatus = (esStop, // stopped via TDoscommand.Stop
    esProcess, // ended via Child-Process
    esStill_Active, // still active
    esNone, // not executed yet
    esError, // ended via Exception
    esTime); // ended because of time

type
  // Thread safety is not required as the TTimeoutCalculator is only used in TDosThread.
  TTimeoutCalculator = record
  strict private
    FSinceBeginning: Cardinal;
    FSinceLastOutput: Cardinal;
    fBeginTimeStamp: Cardinal;
    fLastOutputTimeStamp: Cardinal;
    const
      MaxTimeToWait = 1000;
    function GetTimeToWait: Cardinal;
    function GetTimeout: Boolean;
  private
  public
    procedure Beginning(LastOutput, SinceBeginning: Single); // call this at the beginning of a process
    procedure NewOutput; // call this when a New output is received
    property TimeToWait: Cardinal read GetTimeToWait;
    property Timeout: Boolean read GetTimeout;
  end;

  TNewLineEvent = procedure(ASender: TObject; const ANewLine: string; AOutputType: TOutputType) of object;
  // if New line is read via pipe
  TNewCharEvent = procedure(ASender: TObject; ANewChar: Char) of object;
  // every New char from pipe
  TErrorEvent = procedure(ASender: TObject; AE: Exception; var AHandled: Boolean) of object;
  // if Exception occurs in TDosThread -> if not handled, Messagebox will be shown
  TTerminateProcessEvent = procedure(ASender: TObject; var ACanTerminate: Boolean) of object;
  // called when Dos-Process has to be terminated (via TerminateProcess); just asking if thread can terminate process

  // added by sirius (synchronizes inputlines between Mainthread and TDosThread)
  TInputLines = class(TSimpleRWSync)
  strict private
    FEvent: TEvent;
    FList: TStrings;
    function get_Strings(AIndex: Integer): string;
    procedure set_Strings(AIndex: Integer; const AValue: string);
  public
    constructor Create; reintroduce;
    destructor Destroy; override;
    function Add(const AValue: string): Integer;
    function Count: Integer;
    procedure Delete(AIndex: Integer);
    function LockList: TStrings;
    procedure UnlockList;
    property Event: TEvent read FEvent;
    property Strings[AIndex: Integer]: string read get_Strings write set_Strings; default;
  end;

  //by sirius; syncronized string (TReadPipe<->TDosThread)
  TSyncString = class(TSimpleRWSync)
  strict private
    FValue: string;
    function get_Value: string;
    procedure set_Value(const AValue: string);
  public
    procedure Add(const AValue: string);
    procedure Delete(ACount: Integer);
    function Length: Integer;
    property Value: string read get_Value write set_Value;
  end;

  TReadPipe = class(TThread) // by sirius (wait for pipe input without sleep(1))
    // writes pipe input into TSyncString --> set event  --> TDosThread can read input
  strict private
    FEvent: TEvent;
    fEncoding: TEncoding;
    Fread_stdout, Fwrite_stdout: THandle;
    FSyncString: TSyncString;
  strict protected
    procedure Execute; override;
  public
    constructor Create(AReadStdout, AWriteStdout: THandle; AEncoding: TEncoding); reintroduce;
    destructor Destroy; override;
    procedure Terminate; reintroduce;
    property Event: TEvent read FEvent;
    property ReadString: TSyncString read FSyncString;
  end;

  TDosCommand = class;

  // sirius2: added EnvironmentStrings and (En/De)coding-events
  // the thread that is waiting for outputs through the pipe
  TDosThread = class(TThread)
  strict private
    FCommandLine: string;
    FCurrentDir: string;
    FEnvironment: String;
    FInputLines: TInputLines; // FiFo for Input
    FInputToOutput: Boolean;
    FLines: TStringList;
    fEncoding: TEncoding;
    FOnNewChar: TNewCharEvent;
    FOnNewLine: TNewLineEvent;
    FOnTerminateProcess: TTerminateProcessEvent;
    FOutputLines: TStrings;
    FOwner: TDosCommand;
    FPriority: Integer;
    FProcessInformation: TProcessInformation;
    FTerminateEvent: TEvent;
    FTimer: TTimeoutCalculator;
    procedure DoEndStatus(AValue: TEndStatus);
    procedure DoLinesAdd(const AStr: string);
    procedure DoNewChar(AChar: Char);
    procedure DoNewLine(const AStr: string; AOutputType: TOutputType);
    procedure DoReadLine(AReadString: TSyncString; var ALast: string; var ALineBeginned: Boolean);
    procedure DoSendLine(AWritePipe: THandle; var ALast: string; var ALineBeginned: Boolean);
    procedure DoTerminateProcess;
  private
    FExitCode: Cardinal;
  strict protected // DoSync-Methods are in Main-Thread-Context (called via Synchronize)
    FCanTerminate: Boolean;
    procedure Execute; override;
  public
    constructor Create(AOwner: TDosCommand; const ACl, ACurrDir: string; AOl: TStrings;
      AMaxO, AMaxB: Single; AOnl: TNewLineEvent; AOnc: TNewCharEvent; Ot: TNotifyEvent;
      AOtp: TTerminateProcessEvent; Ap: Integer; Aito: Boolean; AEnv: TStrings; Encoding: TEncoding); reintroduce;
    destructor Destroy; override;
    procedure Terminate; reintroduce;
    property InputLines: TInputLines read FInputLines;
    property Lines: TStringList read FLines;
  end;

  TDosCommand = class(TComponent) // the component to put on a form
  strict private
    FCommandLine: string;
    FCurrentDir: string;
    FEnvironment: TStrings; // sirius2
    FExitCode: Cardinal;
    FInputToOutput: Boolean;
    FLines: TStringList;
    FMaxTimeAfterBeginning: Single;
    FMaxTimeAfterLastOutput: Single;
    FonExecuteError: TErrorEvent;
    FOnNewChar: TNewCharEvent;
    FOnNewLine: TNewLineEvent;
    FOnTerminated: TNotifyEvent;
    FOnTerminateProcess: TTerminateProcessEvent;
    FOutputLines: TStrings;
    FPriority: Integer;
    FThread: TDosThread;
    fEncoding: TEncoding;
    fCodepage: Word;
    function get_EndStatus: TEndStatus;
    function get_IsRunning: Boolean;
  private
    FEndStatus: Integer;
    FProcessInformation: TProcessInformation;
    procedure SetCodepage(const Value: Word);
  strict protected
    procedure ThreadTerminated(ASender: TObject);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    ///<summary>read OEM Codepage from registry</summary>
    class function GetOEMCodepage: Word; static;
    ///<summary>the user call this to execute the command</summary>
    procedure Execute;
    ///<summary>add a line in the input pipe</summary>
    procedure SendLine(const AValue: string; AEol: Boolean);
    ///<summary>the user can stop the process with this method. Stops process and waits until end</summary>
    procedure Stop(Wait: Boolean=False);
    property EndStatus: TEndStatus read get_EndStatus;
    property ExitCode: Cardinal read FExitCode;
    ///<summary>When true, a command is still running // MK: 20030613</summary>
    property IsRunning: Boolean read get_IsRunning;
    ///<summary>if the user want to access all the outputs of a process, he can use this property.
    /// Lines is deleted before execution and filled after execution</summary>
    property Lines: TStringList read FLines;
    ///<summary>can be lines of a memo, a richedit, a listbox, ...</summary>
    property OutputLines: TStrings read FOutputLines write FOutputLines;
    ///<summary>stops process and waits, only for createprocess</summary>
    property Priority: Integer read FPriority write FPriority;
    ///<summary>Processinformation from createprocess</summary>
    property ProcessInformation: TProcessInformation read FProcessInformation;
    ///<summary>Codepage for character encoding and decoding</summary>
    property Codepage: Word read fCodepage write SetCodepage;
  published
    ///<summary>command to execute</summary>
    property CommandLine: string read FCommandLine write FCommandLine;
    ///<summary>currentdir for childprocess (if empty -> currentdir is
    /// same like currentdir in parent process), by sirius</summary>
    property CurrentDir: string read FCurrentDir write FCurrentDir;
    ///<summary>add Environment variables for process (if empty -> environment of parent process is used)</summary>
    property Environment: TStrings read FEnvironment;
    ///<summary>check it if you want that the inputs appear also in the outputs</summary>
    property InputToOutput: Boolean read FInputToOutput write FInputToOutput;
    ///<summary>Time in seconds until the process is stopped</summary>
    property MaxTimeAfterBeginning: Single read FMaxTimeAfterBeginning write FMaxTimeAfterBeginning;
    ///<summary>Time in seconds after the last output until the process is stopped</summary>
    property MaxTimeAfterLastOutput: Single read FMaxTimeAfterLastOutput write FMaxTimeAfterLastOutput;
    ///<summary>event if DosCommand.execute is aborted via Exception</summary>
    property OnExecuteError: TErrorEvent read FonExecuteError write FonExecuteError;
    ///<summary>event for each New char that is received through the pipe</summary>
    property OnNewChar: TNewCharEvent read FOnNewChar write FOnNewChar;
    ///<summary>event for each New line that is received through the pipe</summary>
    property OnNewLine: TNewLineEvent read FOnNewLine write FOnNewLine;
    ///<summary>event for the end of the process (normally, time out or by user (DosCommand.Stop;))</summary>
    property OnTerminated: TNotifyEvent read FOnTerminated write FOnTerminated;
    ///<summary>event to ask for processtermination</summary>
    property OnTerminateProcess: TTerminateProcessEvent read FOnTerminateProcess write FOnTerminateProcess;
  end;

implementation

uses
  System.Types;

resourcestring
  SStillRunning = 'DosCommand still running';
  SNotRunning = 'DosCommand not running';
  SNoCommandLine = 'No Commandline to execute';
  SProcessError = 'Error creating Process: %s - (%s)';
  SPipeError = 'Error creating Pipe: %s';

const
  MaxBufSize = 1024;
  fin: String = #0;

procedure RaisePipeError;
begin
  raise ECreatePipeError.CreateResFmt(@SPipeError, [SysErrorMessage(GetLastError)]);
end;

function GetTickDiff(t1, t2: Cardinal): Cardinal; inline;
begin
  {$ifopt Q+}{$define recoveroverflowcheck}{$Q-}{$else}{$undef recoveroverflowcheck}{$endif}
  Result := t2 - t1;
  {$ifdef recoveroverflowcheck}{$Q+}{$endif}
end;

{ TTimeoutCalculator }

procedure TTimeoutCalculator.Beginning(LastOutput, SinceBeginning: Single);
begin
  fBeginTimeStamp := GetTickCount;
  fLastOutputTimeStamp := fBeginTimeStamp;
  if (LastOutput>0) and (LastOutput<MAXDWORD div 1000)
  then
    fSinceLastOutput := trunc(LastOutput * 1000)
  else
    fSinceLastOutput := 0;
  if (SinceBeginning>0) and (SinceBeginning<MAXDWORD div 1000)
  then
    fSinceBeginning := trunc(SinceBeginning * 1000)
  else
    fSinceBeginning := 0;
end;

function TTimeoutCalculator.GetTimeout: Boolean;
var
  Current: Cardinal;
begin
  Current := GetTickCount;
  if (FSinceLastOutput>0) and (GetTickDiff(fLastOutputTimeStamp, Current)>=FSinceLastOutput) then
    Result := True
  else
  if (FSinceBeginning>0) and (GetTickDiff(fBeginTimeStamp, Current)>=FSinceBeginning) then
    Result := True
  else
    Result := False;
  {$ifdef debug}
  if Result then
    outputdebugstring('Timeout = true')
  else
    outputdebugstring('Timeout = false');
  {$endif}
end;

function TTimeoutCalculator.GetTimeToWait: Cardinal;
var
  Current: Cardinal;
  LTime: Cardinal;
begin
  Result := infinite;
  Current := GetTickCount;
  if FSinceLastOutput>0 then
  begin
    LTime := GetTickDiff(fLastOutputTimeStamp, Current);
    if LTime>=FSinceLastOutput then // Timeout
      Result := 0
    else
      Result := FSinceLastOutput - LTime;
  end;
  if (Result>0) and (FSinceBeginning>0) then
  begin
    LTime := GetTickDiff(fBeginTimeStamp, Current);
    if LTime>=FSinceBeginning then // Timeout
      Result := 0
    else
    begin
      LTime := FSinceBeginning - LTime;
      if LTime<Result then
        Result := LTime
    end;
  end;
  if Result>MaxTimeToWait
  then
    Result := MaxTimeToWait;
  {$ifdef debug}
  if Result=infinite
  then
    outputdebugstring('TimeToWait = infinite')
  else
    outputdebugstring(PChar(Format('TimeToWait = %f', [Result/1000])));
  {$endif}
end;

procedure TTimeoutCalculator.NewOutput;
begin
  fLastOutputTimeStamp := GetTickCount;
end;

function GetEnvironmentLength(p: PChar): Integer;
// Length without list termination char (#0)
// GetEnvironmentLength('ABC'#0'EFG'#0#0) = 8
var
  L: Integer;
  prev: Char;
begin
  if p=nil then
    exit(0);
  L := 0;
  repeat
    inc(L);
    prev := p^;
    inc(p);
  until (prev=#0) and (p^=#0);
  Result := L;
end;

function GetEnvironment(List: TStrings): string;
var
  lpEnvironment: PChar;
  EnvLength: Integer;
  TxtLength: Integer;
  i: Integer;
  pEnvText: PChar;
begin
  // set environment variables and add parent env
  if List.Count = 0 then
    Result := ''
  else
  begin
    // Calc length of text including #0 characters
    TxtLength := List.Count + 1;
    for i := 0 to List.Count - 1 do
      inc(TxtLength, Length(List[i]));
    lpEnvironment := GetEnvironmentStrings;
    try
      EnvLength := GetEnvironmentLength(lpEnvironment);
      if EnvLength>0 then
      begin
        inc(TxtLength, EnvLength);
        SetLength(Result, TxtLength);
        pEnvText := Pointer(Result);
        StrMove(pEnvText, lpEnvironment, EnvLength);
        inc(pEnvText, EnvLength);
      end
      else
      begin
        SetLength(Result, TxtLength);
        pEnvText := Pointer(Result);
      end
    finally
      freeEnvironmentStrings(lpEnvironment);
    end;

    for i := 0 to List.Count - 1 do
    begin
      StrPCopy(pEnvText, List[i]);
      inc(pEnvText, Length(List[i])+1); // skip #0
    end;
    pEnvText^ := #0; // Append End of List
  end;
end;

{ TDosThread }

constructor TDosThread.Create(AOwner: TDosCommand; const ACl, ACurrDir: string; AOl: TStrings;
      AMaxO, AMaxB: Single; AOnl: TNewLineEvent; AOnc: TNewCharEvent; Ot: TNotifyEvent;
      AOtp: TTerminateProcessEvent; Ap: Integer; Aito: Boolean; AEnv: TStrings; Encoding: TEncoding);
begin
  inherited Create(False);
  fEncoding := Encoding;
  FEnvironment := GetEnvironment(AEnv);
  if AOwner<>nil then
  begin
    FOwner := AOwner;
    FOwner.FEndStatus := Ord(esStill_Active);
  end;
  FLines := TStringList.Create;
  FCommandLine := ACl;
  FCurrentDir := ACurrDir;
  FOutputLines := AOl;
  FInputLines := TInputLines.Create;
  FInputToOutput := Aito;
  FOnNewLine := AOnl;
  FOnNewChar := AOnc;
  FOnTerminateProcess := AOtp;
  Self.OnTerminate := Ot;
  FTimer.Beginning(AMaxO, AMaxB);
  FPriority := Ap;
  FTerminateEvent := TEvent.Create(nil, True, False, '');
end;

destructor TDosThread.Destroy;
begin
  inherited Destroy;
  FInputLines.Free;
  FLines.Free;
  FTerminateEvent.Free;
end;

procedure TDosThread.DoEndStatus(AValue: TEndStatus);
begin
  if FOwner<>nil then
    TInterlocked.Exchange(FOwner.FEndStatus, Ord(AValue));
end;

procedure TDosThread.DoLinesAdd(const AStr: string);
begin
  FLines.Add(AStr);
  if Assigned(FOutputLines)
  then begin
    Queue(procedure
    begin
    FOutputLines.Add(AStr);
    end);
  end;
end;

procedure TDosThread.DoNewChar(AChar: Char);
begin
  if Assigned(FOnNewChar) then
  begin
    Queue(procedure
    begin
      FOnNewChar(FOwner, AChar);
    end);
  end;
end;

procedure TDosThread.DoNewLine(const AStr: string; AOutputType: TOutputType);
begin
  if Assigned(FOnNewLine) then
  begin
    Queue(procedure
    begin
      FOnNewLine(FOwner, AStr, AOutputType);
    end);
  end;
end;

procedure TDosThread.DoReadLine(AReadString: TSyncString; var ALast: string; var ALineBeginned: Boolean);
var
  Reads, Line: string;
  idx: Integer;
  LengthReads: Integer;
  LengthLine: Integer;
  ch: Char;
  EOL: Boolean;
begin
  // check to see if there is any data to read from stdout
  Reads := AReadString.Value;
  LengthReads := Length(Reads);
  if (LengthReads > 0) then
  begin
    AReadString.Delete(LengthReads);
    if (Reads=fin)
    then
      exit;
    Line := ALast; // take the begin of the line (if exists)
    LengthLine := Length(ALast);
    SetLength(Line, LengthLine + LengthReads); // set max possible length
    for idx := 1 to LengthReads do
    begin
      if Terminated then
        Exit;

      ch := Reads[idx];
      DoNewChar(ch);
      EOL := False;
      case ch of
        Char(0): ;// ignore
        Char(13): EOL := True;
        Char(10): EOL := (idx=1) or (Reads[idx-1]<>Char(13));
      else
        begin
          inc(LengthLine);
          Line[LengthLine] := ch; // add a character
        end;
      end;
      if EOL then
      begin
        FTimer.NewOutput; // a New ouput has been caught
        SetLength(Line, LengthLine);
        DoLinesAdd(Line); // add the line
        DoNewLine(Line, otEntireLine);
        SetLength(Line, LengthReads - idx);
        LengthLine := 0;
      end;
    end;
    SetLength(Line, LengthLine);
    ALast := Line; // no EOL found in the rest, maybe in the next output
    if (LengthLine>0) then
    begin
      DoNewLine(Line, otBeginningOfLine);
      ALineBeginned := True;
    end;
  end
end;

procedure TDosThread.DoSendLine(AWritePipe: THandle; var ALast: string; var ALineBeginned: Boolean);
var
  sSends: string;
  bWrite: Cardinal;
  pBuf: TBytes;
  sBuffer: string;
begin
  sSends := FInputLines[0];

  if Length(sSends) > 1 then
  begin
    // First Char of sSends is "_" or white space
    if sSends[1] = '_' then
      sSends := sSends + Char(13) + Char(10);
    Delete(sSends, 1, 1);
    pBuf := fEncoding.GetBytes(sSends);
    // send it to stdin
    if not WriteFile(AWritePipe, Pointer(pBuf)^, Length(pBuf), bWrite, nil)
    then
      RaiseLastOSError;
    if FInputToOutput then // if we have to output the inputs
    begin
      if Assigned(FOutputLines) then
      begin
        if ALineBeginned then
        begin // if we are continuing a line
          ALast := ALast + sSends;
          sBuffer := ALast;
          Queue(procedure
          begin
            FOutputLines[FOutputLines.Count - 1] := sBuffer;
          end);
          ALineBeginned := False;
        end
        else // if it's a New line
        begin
          Queue(procedure
          begin
            FOutputLines.Add(sSends);
          end);
        end;
      end;
      DoNewLine(ALast, otEntireLine);
      ALast := '';
    end;

    FInputLines.Delete(0); // delete the line that has been send

  end;
end;

procedure TDosThread.DoTerminateProcess;
begin
  FCanTerminate := True;
  if Assigned(FOnTerminateProcess) then
    Queue(procedure
    begin
      FOnTerminateProcess(FOwner, FCanTerminate);
    end);
  if FCanTerminate then
  begin
    TerminateProcess(FProcessInformation.hProcess, 0);
    CloseHandle(FProcessInformation.hProcess);
    CloseHandle(FProcessInformation.hThread);
  end;
end;

procedure TDosThread.Execute;
var
  si: TSTARTUPINFO;
  sa: PSECURITYATTRIBUTES; // security information for pipes
  sd: PSECURITY_DESCRIPTOR;
  outputread, outputreadtmp, outputwrite, myoutputwrite, errorwrite, inputRead,
    inputWrite, inputWritetmp, MyProcessHandle: THandle; // pipe handles
  last: string;
  LineBeginned: Boolean;
  currDir: PChar;
  ReadPipeThread: TReadPipe;
  lpEnvironment: Pointer;
  WaitHandles: array [0 .. 3] of THandle;
begin // Execute
  NameThreadForDebugging('TDosThread');
  try
    FreeOnTerminate := Assigned(OnTerminate) or (fOwner<>nil);
    sa := nil;
    sd := nil;
    inputWritetmp := 0;
    try
      GetMem(sa, sizeof(SECURITY_ATTRIBUTES));
      if (Win32Platform = VER_PLATFORM_WIN32_NT) then
      begin // initialize security descriptor (Windows NT)
        GetMem(sd, sizeof(SECURITY_DESCRIPTOR));
        InitializeSecurityDescriptor(sd, SECURITY_DESCRIPTOR_REVISION);
        SetSecurityDescriptorDacl(sd, True, nil, False);
        sa.lpSecurityDescriptor := sd;
      end
      else
      begin
        sa.lpSecurityDescriptor := nil;
        sd := nil;
      end;
      sa.nLength := sizeof(SECURITY_ATTRIBUTES);
      sa.bInheritHandle := True; // allow inheritable handles
      MyProcessHandle := GetCurrentProcess;
      // sirius2: Pipe creation changed to microsoft knowledge base article ID: 190351
      if not(CreatePipe(outputreadtmp, outputwrite, sa, 0)) then
        RaisePipeError;

      if not(DuplicateHandle(MyProcessHandle, outputwrite, MyProcessHandle,
        @errorwrite, 0, True, DUPLICATE_SAME_ACCESS)) then
        RaisePipeError;
      if not(DuplicateHandle(MyProcessHandle, outputwrite, MyProcessHandle,
        @myoutputwrite, 0, False, DUPLICATE_SAME_ACCESS)) then
        RaisePipeError;

      if not(CreatePipe(inputRead, inputWritetmp, sa, 0)) then
        RaisePipeError;

      if not(DuplicateHandle(MyProcessHandle, outputreadtmp,
        MyProcessHandle, @outputread, 0, False, DUPLICATE_SAME_ACCESS)) then
        RaisePipeError;

      if not(DuplicateHandle(MyProcessHandle, inputWritetmp,
        MyProcessHandle, @inputWrite, 0, False, DUPLICATE_SAME_ACCESS)) then
        RaisePipeError;

      if not CloseHandle(outputreadtmp) then
        RaisePipeError;
      if not CloseHandle(inputWritetmp) then
        RaisePipeError;

      GetStartupInfo(si); // set startupinfo for the spawned process
      { The dwFlags member tells CreateProcess how to make the process.
        STARTF_USESTDHANDLES validates the hStd* members. STARTF_USESHOWWINDOW
        validates the wShowWindow member. }
      si.dwFlags := STARTF_USESTDHANDLES or STARTF_USESHOWWINDOW;
      si.wShowWindow := SW_HIDE;
      // set the New handles for the child process
      si.hStdOutput := outputwrite;
      si.hStdError := errorwrite;
      si.hStdInput := inputRead;

      if FCurrentDir = '' then
        currDir := nil
      else
        currDir := PChar(FCurrentDir);

      lpEnvironment := Pointer(FEnvironment);

      // spawn the child process

      if not(CreateProcess(nil, PChar(FCommandLine), nil, nil, True,
        CREATE_NEW_CONSOLE or FPriority
{$IFDEF UNICODE}
        or CREATE_UNICODE_ENVIRONMENT
{$ENDIF}
        , lpEnvironment, currDir, si, FProcessInformation)) then
      begin
        raise ECreateProcessError.CreateResFmt(@SProcessError,
          [FCommandLine, SysErrorMessage(getlasterror)]);
      end;

      if FOwner<>nil then
      begin
        Queue(procedure
        begin
          FOwner.FProcessInformation := FProcessInformation;
        end);
        // take Processinformation to mainthread
      end;

      // sirius2: close unneeded handles
      if not CloseHandle(outputwrite) then
        RaisePipeError;
      if not CloseHandle(inputRead) then
        RaisePipeError;
      if not CloseHandle(errorwrite) then
        RaisePipeError;

      ReadPipeThread := TReadPipe.Create(outputread, myoutputwrite, fEncoding);

      last := ''; // Buffer to save last output without finished with EOL
      LineBeginned := False;

      GetExitCodeProcess(FProcessInformation.hProcess, FExitCode);
      // init ExitCode

      try
        repeat // main program loop

          // thread is waiting to one of
          WaitHandles[0] := ReadPipeThread.Event.Handle;
          // New output from childprocess
          WaitHandles[1] := FInputLines.Event.Handle;
          // user has New line (TDosCommand.Sendline) to deliver
          WaitHandles[2] := FProcessInformation.hProcess;
          // Process-Ending (child)
          WaitHandles[3] := FTerminateEvent.Handle;
          // Termination of this thread (from mainthread)

          case WaitforMultipleObjects(Length(WaitHandles), @WaitHandles, False, FTimer.TimeToWait) of
            Wait_Object_0 + 2:
              begin // Process terminated
                {$ifdef Debug}
                outputdebugstring('Process terminated');
                {$endif}
                GetExitCodeProcess(FProcessInformation.hProcess, FExitCode);
              end;

            Wait_Object_0 + 1:
              begin // InputEvent
                if ((FInputLines.Count > 0) and not Terminated) then
                  DoSendLine(inputWrite, last, LineBeginned);
                if FInputLines.Count > 0 then
                  FInputLines.Event.SetEvent;
              end;

            Wait_Object_0 + 0:
              begin // ReadEvent
                while ReadPipeThread.ReadString.Length > 0 do
                  DoReadLine(ReadPipeThread.ReadString, last, LineBeginned);
              end;
          end;

        until Terminated or (FExitCode <> STILL_ACTIVE) or FTimer.Timeout;

        if (FExitCode <> STILL_ACTIVE) then
        begin
          DoEndStatus(esProcess);
          FCanTerminate := False;
        end
        else
        begin
          FCanTerminate := True;
          if Terminated then
            DoEndStatus(esStop)
          else
            DoEndStatus(esTime);
          DoTerminateProcess; // stop Process
        end;

      finally
        if FCanTerminate then
          Waitforsingleobject(FProcessInformation.hProcess, 1000);
        GetExitCodeProcess(FProcessInformation.hProcess, FExitCode);
        ReadPipeThread.Terminate;
        ReadPipeThread.WaitFor;
        // If not empty get last output from ReadPipeThread
        while ReadPipeThread.ReadString.Length > 0 do
          DoReadLine(ReadPipeThread.ReadString, last, LineBeginned);
        ReadPipeThread.Free;
      end;
      if (last <> '') then
      begin // If not empty flush last output
        DoLinesAdd(last);
        DoNewLine(last, otEntireLine);
      end;
    finally
      FreeMem(sd);
      FreeMem(sa);

      if not CloseHandle(outputread) then
        RaisePipeError;
      if not CloseHandle(inputWrite) then
        RaisePipeError;
      if not CloseHandle(myoutputwrite) then
        RaisePipeError;
    end;
  except
    on E: Exception do
    begin
      {$ifdef debug}
      OutputDebugString(PChar('EXCEPTION: TDosThread ' + E.Message));
      {$endif}
      DoEndStatus(esError);
      raise;
    end;
  end;
end;

procedure TDosThread.Terminate;
begin
  inherited Terminate;
  FTerminateEvent.SetEvent;
end;

{ TDosCommand }

constructor TDosCommand.Create(AOwner: TComponent);
var
  LCodepage: Word;
begin
  inherited Create(AOwner);
  FCommandLine := '';
  FCurrentDir := '';
  FLines := TStringList.Create;
  FMaxTimeAfterBeginning := 0;
  FMaxTimeAfterLastOutput := 0;
  FPriority := NORMAL_PRIORITY_CLASS;
  FEndStatus := Ord(esNone);
  FEnvironment := TStringList.Create;
  LCodepage := GetOEMCodepage;
  if LCodepage<>0 then
    Codepage := LCodepage
  else
    Codepage := DefaultOEMCodepage;
end;

destructor TDosCommand.Destroy;
begin
  Stop(True); // Wait until the end to avoid memory leaks during application termination
  FLines.Free;
  FEnvironment.Free;
  if not TEncoding.IsStandardEncoding(fEncoding) then
    fEncoding.Free;
  inherited Destroy;
end;

procedure TDosCommand.Execute;
begin
  if IsRunning then
    raise EDosCommand.CreateRes(@SStillRunning);

  if FCommandLine = '' then // MK: 20030613
    raise EDosCommand.CreateRes(@SNoCommandLine);
  FLines.Clear; // clear old outputs
  FThread := TDosThread.Create(Self, FCommandLine, FCurrentDir, FOutputLines,
    FMaxTimeAfterLastOutput, FMaxTimeAfterBeginning,
    FOnNewLine, FOnNewChar, ThreadTerminated, FOnTerminateProcess, FPriority,
    FInputToOutput, FEnvironment, fEncoding);
end;

class function TDosCommand.GetOEMCodepage: Word;
const
  CodepageKey = '\SYSTEM\CurrentControlSet\Control\Nls\CodePage';
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create(KEY_READ);
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey(CodepageKey, False) then
      Result := StrToIntDef(Reg.ReadString('OEMCP'), 0)
    else
      Result := 0;
  finally
    Reg.Free
  end;
end;

function TDosCommand.get_EndStatus: TEndStatus;
begin
  Result := TEndStatus(FEndStatus);
end;

function TDosCommand.get_IsRunning: Boolean;
begin
  Result := Assigned(FThread); // sirius
end;

procedure TDosCommand.SendLine(const AValue: string; AEol: Boolean);
const
  EolCh: array [Boolean] of Char = (' ', '_');
begin
  if (FThread <> nil) then
    FThread.InputLines.Add(EolCh[AEol] + AValue)
  else
    raise EDosCommand.CreateRes(@SNotRunning);
end;

procedure TDosCommand.SetCodepage(const Value: Word);
begin
  if (Value<>fCodepage)
  then begin
    if not TEncoding.IsStandardEncoding(fEncoding) then
      fEncoding.Free;
    if Value=CP_ACP then
      fEncoding := TEncoding.Default
    else
      fEncoding := TEncoding.GetEncoding(Value);
    fCodepage := fEncoding.CodePage;
  end;
end;

procedure TDosCommand.Stop(Wait: Boolean);
begin
  if (FThread <> nil) then
  begin
    if Wait then
    begin
      FThread.FreeOnTerminate := False;
      FThread.Terminate;
      FThread.WaitFor;
      FreeAndNil(FThread);
    end
    else
      FThread.Terminate;
  end;
end;

procedure TDosCommand.ThreadTerminated(ASender: TObject);
var
  E: Exception;
  handled: Boolean;
  procedure show(E: Exception);
  begin
    MessageBox(0, PChar(E.ClassName + sLineBreak + E.Message),
      PChar(ExtractFileName(ParamStr(0))), MB_OK or MB_ICONSTOP or
      MB_TASKMODAL);
  end;

begin
  if (ASender is TDosThread) then
  begin
    FExitCode := TDosThread(ASender).FExitCode;
    FLines.AddStrings(TDosThread(ASender).Lines);
    if (ASender=FThread) and FThread.FreeOnTerminate then
      FThread := nil;
    if Assigned(TDosThread(ASender).FatalException) then
    begin
      E := TThread(ASender).FatalException as Exception;
      try
        if Assigned(FonExecuteError) then
        begin
          handled := False;
          FonExecuteError(Self, E, handled);
          if not handled then
            show(E);
        end
        else
          show(E);
      except
        on E: Exception do
          show(E);
      end;
    end;
  end;
  if Assigned(FOnTerminated) then
    FOnTerminated(Self);
end;

{ TInputLines }

constructor TInputLines.Create;
begin
  inherited Create;
  FList := TStringList.Create;
  FEvent := TEvent.Create(nil, False, False, '');
end;

destructor TInputLines.Destroy;
begin
  LockList;
  try
    FList.Free;
    FList := nil;
  finally
    UnlockList;
    FEvent.Free;
  end;
  inherited Destroy;
end;

function TInputLines.Add(const AValue: string): Integer;
var
  pList: TStrings;
begin
  pList := LockList;
  try
    Result := pList.Add(AValue);
  finally
    UnlockList;
  end;
  FEvent.SetEvent;
end;

function TInputLines.Count: Integer;
var
  pList: TStrings;
begin
  pList := LockList;
  try
    Result := pList.Count;
  finally
    UnlockList;
  end;
end;

procedure TInputLines.Delete(AIndex: Integer);
var
  pList: TStrings;
begin
  pList := LockList;
  try
    pList.Delete(AIndex);
  finally
    UnlockList;
  end;
end;

function TInputLines.get_Strings(AIndex: Integer): string;
var
  pList: TStrings;
begin
  pList := LockList;
  try
    Result := pList[AIndex];
  finally
    UnlockList;
  end;
end;

function TInputLines.LockList: TStrings;
begin
  BeginWrite;
  Result := FList;
end;

procedure TInputLines.set_Strings(AIndex: Integer; const AValue: string);
var
  pList: TStrings;
begin
  pList := LockList;
  try
    pList[AIndex] := AValue;
  finally
    UnlockList;
  end;
end;

procedure TInputLines.UnlockList;
begin
  EndWrite;
end;

{ TSyncString }

procedure TSyncString.Add(const AValue: string);
begin
  BeginWrite;
  try
    FValue := FValue + AValue;
  finally
    EndWrite;
  end;
end;

procedure TSyncString.Delete(ACount: Integer);
begin
  BeginWrite;
  try
    System.Delete(FValue, 1, ACount);
  finally
    EndWrite;
  end;
end;

function TSyncString.get_Value: string;
begin
  BeginRead;
  try
    Result := FValue;
  finally
    EndRead;
  end;
end;

function TSyncString.Length: Integer;
begin
  BeginRead;
  try
    Result := System.Length(FValue);
  finally
    EndRead;
  end;
end;

procedure TSyncString.set_Value(const AValue: string);
begin
  BeginWrite;
  try
    FValue := AValue;
  finally
    EndWrite;
  end;
end;

{ TReadPipe }

constructor TReadPipe.Create(AReadStdout, AWriteStdout: THandle; AEncoding: TEncoding);
begin
  inherited Create(False);
  fEncoding := AEncoding;
  FEvent := TEvent.Create(nil, False, False, '');
  FSyncString := TSyncString.Create;
  Fread_stdout := AReadStdout;
  Fwrite_stdout := AWriteStdout;
  FreeOnTerminate := False;
end;

destructor TReadPipe.Destroy;
begin
  inherited Destroy;
  FEvent.Free;
  FSyncString.Free;
end;

procedure TReadPipe.Execute;
  function AvailableBytes(var avail: Cardinal): Boolean;
  begin
    Result := PeekNamedPipe(Fread_stdout, nil, 0, nil, @avail, nil);
  end;
var
  LBytes: TBytes;
  LByteCount: Cardinal;
  LBuffSize: Cardinal;
  Error: Cardinal;
  Avail: Cardinal;
begin
  try
    LBuffSize := 128;
    NameThreadForDebugging('TReadPipe');
    repeat
      if AvailableBytes(Avail) and (Avail>LBuffSize) then
        LBuffSize := Avail;
      SetLength(LBytes, LBuffSize);
      LByteCount := 0;
      // wait for input
      if not ReadFile(Fread_stdout, Pointer(LBytes)^, LBuffSize, LByteCount, nil) then
      begin
        Error := GetLastError;
        if Error <> Error_broken_pipe then
          RaiseLastOSError(Error);
      end;
      FSyncString.Add(fEncoding.GetString(LBytes, 0, LByteCount));
      if not Terminated then
        FEvent.SetEvent;
    until Terminated and not(AvailableBytes(Avail) and (Avail>0));
  except
    {$ifdef debug}
    on E: Exception do
      OutputDebugString(PChar('EXCEPTION: TReadPipe Execute ' + E.Message));
    {$endif}
  end;
end;

procedure TReadPipe.Terminate;
var
  bwrite: Cardinal;
begin
  inherited Terminate;
  // write dummy string to stdout to trigger ReadFile.
  if not WriteFile(Fwrite_stdout, Pointer(fin)^, Length(fin), bwrite, nil) then
    RaiseLastOSError;
end;

end.
