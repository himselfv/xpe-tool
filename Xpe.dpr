program Xpe;

{$APPTYPE CONSOLE}

(*
  xpe find "Internet Explorer"
  xpe info [component id]
  xpe dep [component id]

  xpe files [component id] >file.list
  xpe registry-export [component id] >export.reg
*)

uses
  SysUtils, ActiveX, AdoInt, TntClasses, Variants, Windows,
  CodepageUtils, WideStrUtils,
  ResourceList in 'ResourceList.pas',
  XpeConsts in 'XpeConsts.pas';

type
  EBadUsage = class(Exception);

procedure BadUsage(msg: string = '');
begin
  raise EBadUsage.Create(msg);
end;

type
  TCommandInfo = record
    cmd: string;
    p: string;
    desc: string;
  end;

const
  COMMANDS: array[0..5] of TCommandInfo = (
    ( cmd: 'help'; p: '[command]';
      desc: 'Prints help.' ),
    ( cmd: 'find'; p: '<part of component name>';
      desc: 'Lists all components with the specified name.' ),
    ( cmd: 'info'; p: '<component id>';
      desc: 'Prints basic component info.' ),
    ( cmd: 'deps'; p: '<component id>';
      desc: 'Prints component dependencies.' ),
    ( cmd: 'files'; p: '<component id>';
      desc: 'Prints the list of files included in component.' ),
    ( cmd: 'registry-export'; p: '<component id>';
      desc: 'Prints registry export file (.reg) of component registry entries.' )
  );

procedure PrintUsage;
var i: integer;
begin
  writeln('Usage: '+ExtractFilename(Paramstr(0))+' <command>');
  writeln('Commands: ');
  for i := 0 to Length(COMMANDS) - 1 do
    writeln('  '+COMMANDS[i].cmd+' '+COMMANDS[i].p);
end;

procedure PrintHelp(cmd: string);
var i, ind: integer;
begin
  if cmd='' then begin
    PrintUsage;
    exit;
  end;

  ind := -1;
  for i := 0 to Length(COMMANDS) - 1 do
    if SameText(COMMANDS[i].cmd, cmd) then begin
      ind := i;
      break;
    end;
  if ind<0 then begin
    writeln('Unknown command: '+cmd);
    PrintUsage;
    exit;
  end;

  writeln('Usage: '+ExtractFilename(Paramstr(0))+' '+COMMANDS[ind].cmd+' '+COMMANDS[ind].p);
  writeln(COMMANDS[i].desc);
end;


var
  Config: TTntStringList;
  conn: _Connection;

procedure Init;
begin
  Config := TTntStringList.Create;
  Config.LoadFromFile(ChangeFileExt(paramstr(0), '.cfg'));
  CoInitializeEx(nil, COINIT_MULTITHREADED);
  conn := CoConnection.Create;
  conn.Open(Config.Values['dsn'], Config.Values['user'], Config.Values['pass'], 0);
end;

procedure Free;
begin
  conn := nil;
  CoUninitialize();
  FreeAndNil(Config);
end;



function getComponentId(ParamId: integer): integer;
begin
  if ParamCount<ParamId then BadUsage('ComponentId missing');
  if not TryStrToInt(ParamStr(ParamId), Result) then
    BadUsage('ComponentId must be integer: '+ParamStr(ParamId));
end;

//Converts variant to integer, or to zero if it's null.
function int(value: OleVariant): integer;
begin
  if VarIsNull(value) or VarIsClear(value) then
    Result := 0
  else
    Result := integer(Value);
end;

//Converts variant to string, or to '' if it's null.
function str(value: OleVariant): WideString;
begin
  if VarIsNull(value) or VarIsClear(value) then
    Result := ''
  else
    Result := WideString(Value);
end;

//Prints contents of ComponentObjects list
procedure writelnList(rs: _Recordset);
begin
  while not rs.EOF do begin
    writeln(Format('%-8s %-48s %s', [
      str(rs.Fields['ComponentId'].Value),
      str(rs.Fields['DisplayName'].Value),
      str(rs.Fields['Revision'].Value)
    ]));
    rs.MoveNext();
  end;
end;

//Prints Property=Value pair in a standard format
procedure writelnPair(name, value: string);
begin
  writeln(Format('%-16s %-48s', [
    Name+':',
    Value
  ]));
end;

//Prints a specified field from recordset in Field=Value format
procedure writelnField(rs: _Recordset; FieldName: string);
begin
  writelnPair(FieldName, str(rs.Fields[FieldName].Value));
end;

//MAIN: find <part of name>
procedure ListComponents(PartOfName: WideString);
var rs: _Recordset;
  RecordsAffected: OleVariant;
begin
  rs := conn.Execute('SELECT ComponentId, DisplayName, Revision FROM ComponentObjects '
    +'WHERE DisplayName LIKE '''+PartOfName+''' ORDER BY ComponentId ASC', RecordsAffected, 0);
  writelnList(rs);
end;

//MAIN: print <id>
procedure PrintComponentInfo(ComponentId: integer);
var rs: _Recordset;
  RecordsAffected: OleVariant;
  flags: string;
begin
  rs := conn.Execute('SELECT DateImported, Revision, Visibility, Released, Editable, '
      +'DateCreated, DateRevised, DisplayName, Version, Description, Copyright, '
      +'Vendor, Owners, Authors, Testers, IsMacro '
    +'FROM ComponentObjects WHERE ComponentId='+IntToStr(ComponentId),
    RecordsAffected, 0);
  if rs.EOF then begin
    writeln('Component not found.');
    exit;
  end;

  writelnPair('ComponentId', IntToStr(ComponentId));
  writelnField(rs,'DisplayName');
  writelnField(rs,'Description');
  writelnField(rs,'Version');
  writelnField(rs,'Revision');
  writelnField(rs,'Copyright');
  writelnField(rs,'Vendor');
  writelnField(rs,'Owners');
  writelnField(rs,'Authors');
  writelnField(rs,'Testers');
  writelnField(rs,'Visibility');

  flags := '';
  if boolean(rs.Fields['Released'].Value) then
    flags := flags + 'RELEASED ';
  if boolean(rs.Fields['Editable'].Value) then
    flags := flags + 'EDITABLE ';
  if boolean(rs.Fields['IsMacro'].Value) then
    flags := flags + 'MACRO ';
  writelnPair('Flags', flags);

  writelnField(rs,'DateCreated');
  writelnField(rs,'DateRevised');
  writelnField(rs,'DateImported');
end;

//MAIN: dep <id>
procedure ListDependencies(ComponentId: integer);
var rs: _Recordset;
  RecordsAffected: OleVariant;
begin
  rs := conn.Execute('SELECT ComponentObjects.ComponentId, '
    +'ComponentObjects.DisplayName, ComponentObjects.Revision '
    +'FROM ComponentDependencyLists, ComponentObjects '
    +'WHERE ComponentDependencyLists.DependOnGuid=ComponentObjects.ComponentVIGUID '
    +'AND ComponentDependencyLists.ComponentID='+IntToStr(ComponentId)+' '
    +'ORDER BY ComponentObjects.ComponentId ASC', RecordsAffected, 0);
  writelnList(rs);
end;


//Reads ExtendedResource value from recordset according to ExtendedResourceFormat.
function readFormat(rs: _Recordset; format: integer): OleVariant;
begin
  case format of
    PROPFORMAT_BINARY:
      Result := rs.Fields['BinaryValue'].Value;
    PROPFORMAT_SZ:
      Result := rs.Fields['StringValue'].Value;
    PROPFORMAT_INTEGER:
      Result := rs.Fields['IntegerValue'].Value;
    PROPFORMAT_BOOL:
      Result := rs.Fields['BoolValue'].Value;
    PROPFORMAT_MULTISZ:
      Result := rs.Fields['StringValue'].Value;
    PROPFORMAT_EXPRESSION:
      Result := rs.Fields['StringValue'].Value;
    PROPFORMAT_GUID:
      Result := rs.Fields['GuidValue'].Value;
  else
    Result := Unassigned;
  end;
end;

//Parses ExtendedResource list containing File Resources turning it into a File Resource List
function parseFileResources(rs: _Recordset): TResourceList;
var ResourceId: OleVariant;
  Resource: TFileResource;
  ParamName: string;
begin
  Result := TResourceList.Create(TFileResource);
  try
    while not rs.EOF do begin
      ResourceId := rs.Fields['ResourceId'].Value;
      if VarIsClear(ResourceId) or VarIsNull(ResourceId) then begin
        rs.MoveNext;
        continue;
      end;

      Resource := TFileResource(Result.GetResource(ResourceId));
      ParamName := rs.Fields['Name'].Value;

      if SameText(ParamName, 'BuildOrder') then
        Resource.BuildOrder := rs.Fields['IntegerValue'].Value
      else
      if SameText(ParamName, 'DstPath') then
        Resource.DstPath := str(rs.Fields['StringValue'].Value)
      else
      if SameText(ParamName, 'DstName') then
        Resource.DstName := str(rs.Fields['StringValue'].Value)
      else
      if SameText(ParamName, 'SrcPath') then
        Resource.SrcPath := str(rs.Fields['StringValue'].Value)
      else
      if SameText(ParamName, 'SrcName') then
        Resource.SrcName := str(rs.Fields['StringValue'].Value)
      else
      if SameText(ParamName, 'SrcFileCRC') then
        Resource.SrcFileCRC := rs.Fields['IntegerValue'].Value
      else
      if SameText(ParamName, 'SrcFileSize') then
        Resource.SrcFileSize := rs.Fields['IntegerValue'].Value
      else
      if SameText(ParamName, 'NoExpand') then
        Resource.NoExpand := rs.Fields['BoolValue'].Value
      else
      if SameText(ParamName, 'Overwrite') then
        Resource.Overwrite := rs.Fields['BoolValue'].Value;

      rs.MoveNext;
    end;
  except
    FreeAndNil(Result);
    raise;
  end;
end;

//MAIN: files <id>
procedure FileList(ComponentId: integer);
var rs: _Recordset;
  RecordsAffected: OleVariant;
  Resources: TResourceList;
  Res: TFileResource;
  i: integer;
  Src, Dst: WideString;
begin
  rs := conn.Execute('SELECT * FROM ExtendedProperties WHERE '
    +'ResourceTypeId='+IntToStr(RESOURCETYPE_FILE)+' AND ' //file
    +'OwnerId='+IntToStr(ComponentId)+' '
    +'ORDER BY ResourceId ', RecordsAffected, 0);
  Resources := parseFileResources(rs);
  try
    for i := 0 to Resources.Count - 1 do begin
      Res := TFileResource(Resources.Items[i]);
      Dst := Res.DstPath + '\' + Res.DstName;

      if Res.SrcPath <> '' then
        Src := Res.SrcPath + '\'
      else Src := '';

      if Res.SrcName='' then
        Src := Src + Res.DstName
      else
        Src := Src + Res.SrcName;

      writeln('"'+Src+'" "'+Dst+'"');
    end;
  finally
    FreeAndNil(Resources);
  end;
end;


//Parses ExtendedResource list containing Registry Resources turning it into a Registry Resource List
function parseRegistryResources(rs: _Recordset): TResourceList;
var ResourceId: OleVariant;
  Resource: TRegistryResource;
  ParamName: string;
begin
  Result := TResourceList.Create(TRegistryResource);
  try
    while not rs.EOF do begin
      ResourceId := rs.Fields['ResourceId'].Value;
      if VarIsClear(ResourceId) or VarIsNull(ResourceId) then begin
        rs.MoveNext;
        continue;
      end;

      Resource := TRegistryResource(Result.GetResource(ResourceId));
      ParamName := rs.Fields['Name'].Value;

      if SameText(ParamName, 'BuildOrder') then
        Resource.BuildOrder := rs.Fields['IntegerValue'].Value
      else
      if SameText(ParamName, 'Description') then
        Resource.Description := str(rs.Fields['StringValue'].Value)
      else
      if SameText(ParamName, 'DisplayName') then
        Resource.DisplayName := str(rs.Fields['StringValue'].Value)
      else
      if SameText(ParamName, 'KeyPath') then
        Resource.KeyPath := str(rs.Fields['StringValue'].Value)
      else
      if SameText(ParamName, 'ValueName') then
        Resource.ValueName := str(rs.Fields['StringValue'].Value)
      else
      if SameText(ParamName, 'RegType') then
        Resource.RegType := rs.Fields['IntegerValue'].Value
      else
      if SameText(ParamName, 'RegOp') then
        Resource.RegOp := rs.Fields['IntegerValue'].Value
      else
      if SameText(ParamName, 'RegCond') then
        Resource.RegCond := rs.Fields['IntegerValue'].Value
      else
      if SameText(ParamName, 'RegValue') then begin
        Resource.RegValueFormat := rs.Fields['Format'].Value;
        Resource.RegValue := readFormat(rs, Resource.RegValueFormat);
      end;

      rs.MoveNext;
    end;
  except
    FreeAndNil(Result);
    raise;
  end;
end;


//Sorts registry resources according to their key path, to group same key items.
//Also moves "key only" items to the top.
procedure SortRegistryResources(Resources: TResourceList);
var i, j, cmp: integer;
  ResI: TRegistryResource;
begin
  for i := 1 to Resources.Count - 1 do begin
    ResI := TRegistryResource(Resources.Items[i]);

   //Find new place
    j := i-1;
    while j>=0 do begin
      cmp := WideCompareStr(ResI.KeyPath, TRegistryResource(Resources.Items[j]).KeyPath);
      if (cmp < 0) or ((cmp=0) and (ResI.ValueName='')) then
        Dec(j)
      else
        break;
    end;
    Inc(j);

   //Move
    if j<i then begin
      cmp := i;
      while cmp>j do begin
        Resources.Items[cmp] := Resources.Items[cmp-1];
        Dec(cmp);
      end;
      Resources.Items[j] := ResI;
    end;
  end;
end;

//Escapes registry file type string
function regEscapeStr(str: WideString): WideString;
begin
  str := WideReplaceStr(str, '\', '\\');
  Result := WideReplaceStr(str, '"', '\"');
end;

//Writes registry file type comment
procedure regComment(msg: WideString);
begin
  writeln(';XPE Exporter: '+msg);
end;


const
  sHexSymbols: AnsiString = '0123456789abcdef';

//Prints the contents of binary array in val as registry-compatible hexadecimal string
function regOleToHex(val: OleVariant): string;
var a: array of byte;
  i: integer;
begin
  if VarIsClear(val) or VarIsNull(val) then begin
    Result := ''; //fine, supported hex value in .reg files
    exit;
  end;

  a := val;
  if Length(a)<=0 then begin
    Result := '';
    exit;
  end;

  SetLength(Result, (Length(a)-1)*3 + 2);
  Result[1] := sHexSymbols[(a[0] shr 4) + 1];
  Result[2] := sHexSymbols[(a[0] mod 16) + 1];

  for i := 1 to Length(a) - 1 do begin
    Result[i*3+0]:=',';
    Result[i*3+1]:=sHexSymbols[(a[i] shr 4) + 1];
    Result[i*3+2]:=sHexSymbols[(a[i] mod 16) + 1];
  end;
end;

procedure RegistryExport(ComponentId: integer);
var rs: _Recordset;
  RecordsAffected: OleVariant;
  Resources: TResourceList;
  Res: TRegistryResource;
  i: integer;
  LastKey: WideString;
  ValueName: WideString;
  ValueStr: WideString;
begin
  rs := conn.Execute('SELECT * FROM ExtendedProperties WHERE '
    +'ResourceTypeId='+IntToStr(RESOURCETYPE_REGISTRY)+' AND ' //registry
    +'OwnerId='+IntToStr(ComponentId)+' '
    +'ORDER BY ResourceId ', RecordsAffected, 0);
  Resources := parseRegistryResources(rs);
  try
    SortRegistryResources(Resources);
    writeln('Windows Registry Editor Version 5.00');
    LastKey := '';

    for i := 0 to Resources.Count - 1 do begin
      Res := TRegistryResource(Resources.Items[i]);
      if Res.KeyPath='' then begin
        regComment('Invalid item "'+ValueName+'": empty KeyPath.');
        continue;
      end;

     { We only support a limited set of REGCOND options. It's either DELETE... }
      if Res.RegCond=REGCOND_DELETE then begin
       //Deleting the whole key
        if Res.ValueName='' then begin
          writeln('');
          writeln('-['+Res.KeyPath+']');
          continue;
        end;

       //Else deleting only a single named value.
       { If we're in a different key than before, we need to start it }
        if not WideSameText(Res.KeyPath, LastKey) then begin
          writeln('');
          writeln('['+Res.KeyPath+']');
          LastKey := Res.KeyPath;
        end;

        ValueName := '"'+regEscapeStr(Res.ValueName)+'"';
        writeln(ValueName+'=-');
        continue;
      end;
     { Anything else is considered WRITE. }

     { If we're in a different key than before, we need to start it }
      if not WideSameText(Res.KeyPath, LastKey) then begin
        writeln('');
        writeln('['+Res.KeyPath+']');
        LastKey := Res.KeyPath;
      end;

      if (Res.ValueName='') and (VarIsEmpty(Res.RegValue) or VarIsNull(Res.RegValue)) then
        continue; //It's just a key, no default value

     { Default values OR named values }
      if Res.ValueName<>'' then
        ValueName := '"'+regEscapeStr(Res.ValueName)+'"'
      else
        ValueName := '@';

     {Res.RegValueFormat can be either EXPECTED_FORMAT (according to Res.RegType), or EXPRESSION, or BINARY.
      Expression is not supported: we don't know how to parse it.}

     {If it's binary, we just output it in a supported way: by specifying registry type as "hex(REG_TYPE)"
      The only exception is REG_BINARY itself which is just "hex"}
      if Res.RegValueFormat=PROPFORMAT_BINARY then begin
        if Res.RegType<>REG_BINARY then
          ValueStr := 'hex('+IntToStr(Res.RegType)+'):'
        else
          ValueStr := 'hex:';
        ValueStr := ValueStr + regOleToHex(Res.RegValue);
      end else

     {These are REGISTRY types, not the internal XPE types}
      case Res.RegType of {Only these are supported:}
        REG_SZ, REG_EXPAND_SZ, REG_MULTI_SZ:
          if (Res.RegValueFormat=PROPFORMAT_SZ)
          or (Res.RegValueFormat=PROPFORMAT_MULTISZ) then
            ValueStr := '"'+regEscapeStr(str(Res.RegValue))+'"'
          else begin
            regComment('Value "'+ValueName+'": unsupported ValueFormat of '+IntToStr(Res.RegValueFormat)
              +' for RegType of '+IntToStr(Res.RegType)+'.');
            continue; //unsupported
          end;
        REG_BINARY: begin
          regComment('Value "'+ValueName+'": unsupported ValueFormat of '+IntToStr(Res.RegValueFormat)
            +' for RegType of '+IntToStr(Res.RegType)+'.');
          continue; //PROPFORMAT_BINARY solved earlier, rest is unsupported
        end;
        REG_DWORD:
          if Res.RegValueFormat=PROPFORMAT_INTEGER then
            ValueStr := 'dword:'+IntToHex(int(Res.RegValue), 8)
          else begin
            regComment('Value "'+ValueName+'": unsupported ValueFormat of '+IntToStr(Res.RegValueFormat)
              +' for RegType of '+IntToStr(Res.RegType)+'.');
            continue; //unsupported
          end;
      else
        regComment('Value "'+ValueName+'": unsupported RegType of '+IntToStr(Res.RegType)+'.');
        continue;
      end;

     {Write out}
      writeln(ValueName+'='+ValueStr)
    end;
  finally
    FreeAndNil(Resources);
  end;
end;



var
  cmd: string;
begin
  try
    Init;

    if ParamCount<1 then BadUsage();
    cmd := ParamStr(1);

    if SameText(cmd, 'help') then begin
      if ParamCount>=2 then
        PrintHelp(ParamStr(2))
      else
        PrintHelp('');
    end else
    if SameText(cmd, 'find') then begin
      if ParamCount<2 then BadUsage('Name missing');
      ListComponents(ParamStr(2));
    end else
    if SameText(cmd, 'info') then begin
      PrintComponentInfo(getComponentId(2));
    end else
    if SameText(cmd, 'deps') then begin
      ListDependencies(getComponentId(2));
    end else
    if SameText(cmd, 'files') then begin
      FileList(getComponentId(2));
    end else
    if SameText(cmd, 'registry-export') then begin
      RegistryExport(getComponentId(2));
    end else
      BadUsage('Invalid command: '+cmd);

    Free;
  except
    on E: EBadUsage do begin
      if E.Message<>'' then
        Writeln('Bad usage: '+E.Message);
      PrintUsage;
    end;
    on E:Exception do
      Writeln(E.Classname, ': ', WinToOEM(E.Message));
  end;
end.
