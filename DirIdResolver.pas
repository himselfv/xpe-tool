unit DirIdResolver;

interface

(*
  Usage: set public variables (BootDrive, WindowsDir etc) according to your
  configuration, then call Resolve() to get specific dirs.

  Works for English-native windows. Now, you can have MUI installed, but
  the folders underneath will still be in English.

  Does not resolve CurrentUser-relative variables.
*)

type
  TDirIdResolver = class
  public
    BootDrive: WideString;
    SystemDrive: WideString;
    Windows: WideString;
    ProgramFiles: WideString;
    DocumentsAndSettings: WideString;
    function System32: WideString;
    function AllUsers: WideString;
    function Resolve(Id: integer): WideString;
    function ResolveStr(Str: WideString): WideString;
  end;

implementation
uses SysUtils;

function TDirIdResolver.System32: WideString;
begin
  Result := Windows + '\System32';
end;

function TDirIdResolver.AllUsers: WideString;
begin
  Result := DocumentsAndSettings + 'All Users';
end;

function TDirIdResolver.Resolve(Id: integer): WideString;
begin
  case Id of
    10: Result := Windows;
    11: Result := System32;
    12: Result := System32 + '\Drivers';
    17: Result := Windows + '\Inf';
    18: Result := Windows + '\Help';
    20: Result := Windows + '\Fonts';
    21: Result := System32 + '\Viewers';
    23: Result := System32 + '\Spool\Drivers\Color';
    24: Result := SystemDrive;
    25: Result := Windows;
    30: Result := BootDrive;
    50: Result := Windows + '\System';

    51: Result := System32 + '\spool';
    52: Result := System32 + '\spool\Drivers\w32x86';
    54: Result := BootDrive;
    55: Result := System32 + '\spool\PRTPROCS\w32x86';

    16404: Result := Windows + '\Fonts';
    16406: Result := AllUsers + '\Start Menu';
    16407: Result := AllUsers + '\Start Menu\Programs';
    16408: Result := AllUsers + '\Start Menu\Programs\Startup';
    16409: Result := AllUsers + '\Desktop';
    16415: Result := AllUsers + '\Favorites';
    16419: Result := AllUsers + '\Application Data';

    16420: Result := Windows;
    16421: Result := System32;
    16422: Result := ProgramFiles;
    16425: Result := System32;
    16427: Result := ProgramFiles + '\Common Files';

    16429: Result := AllUsers + '\Templates';
    16430: Result := AllUsers + '\Documents';
    16431: Result := AllUsers + '\Start Menu\Programs\Administrative Tools';
    16437: Result := AllUsers + '\Documents\My Music';
    16438: Result := AllUsers + '\Documents\My Pictures';
    16439: Result := AllUsers + '\Documents\My Videos';

    16440: Result := Windows + '\resources';
    16441: Result := Windows + '\resources409';
  else
    raise Exception.Create('Directory id %'+IntToStr(Id)+' cannot be resolved.');
  end;
end;

function StrCut(ps, pe: PWideChar): WideString;
var i: integer;
begin
  if integer(pe) < integer(ps) then
    Result := '';
  SetLength(Result, (integer(pe)-integer(ps)) div SizeOf(WideChar));
  i := 1;
  while integer(ps) < integer(pe) do begin
    Result[i] := ps^;
    Inc(ps);
    Inc(i);
  end;
end;

function TDirIdResolver.ResolveStr(Str: WideString): WideString;
var st, pos: PWideChar;
  InVar: boolean;
  val: integer;
begin
  if str='' then begin
    Result := '';
    exit;
  end;

  Result := '';  
  InVar := false;
  st := @str[1];
  pos := @str[1];
  while pos^ <> #00 do begin
    if pos^ <> '%' then begin
      Inc(pos);
      continue;
    end;

    if not InVar then begin
      Result := Result + StrCut(st, pos);
      InVar := true;
      Inc(pos);
      st := pos;
      continue;
    end;

    val := StrToInt(StrCut(st, pos));
    Result := Result + Resolve(val);
    InVar := false;
    Inc(pos);
    st := pos;
  end;

  if not InVar then
    Result := Result + StrCut(st, pos);
end;

end.
