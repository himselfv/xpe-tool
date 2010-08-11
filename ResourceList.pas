unit ResourceList;

interface
uses SysUtils, Windows, XpeConsts;

type
  TResource = class
  public
    ResourceId: integer;
    BuildOrder: integer;
  end;
  CResource = class of TResource;

  TFileResource = class(TResource)
  public
    DstPath: WideString;
    DstName: WideString;
    NoExpand: boolean;
    Overwrite: boolean;
    SrcFileCRC: integer;
    SrcFileSize: integer;
    SrcName: WideString;
    SrcPath: WideString;
  end;

  TRegistryResource = class(TResource)
  public
    Description: WideString;
    DisplayName: WideString;
    RegType: integer; {REG_SZ, REG_DWORD, etc}
    RegOp: integer;
    RegCond: integer; {See XpeConsts}
    KeyPath: WideString;
    ValueName: WideString;
    RegValue: OleVariant; {Can be of different format. Depends on RegType and user choice. Stored in RegValueFormat}
    RegValueFormat: integer;
  end;

  TResourceList = class
  public
    ResourceClass: CResource;
    Count: integer;
    Items: array of TResource;
    constructor Create(AResourceClass: CResource);
    destructor Destroy; override;
    procedure Clear;
    function FindResource(ResourceId: integer): TResource;
    function GetResource(ResourceId: integer): TResource;
  end;

implementation

constructor TResourceList.Create(AResourceClass: CResource);
begin
  inherited Create;
  ResourceClass := AResourceClass;
  Count := 0;
end;

destructor TResourceList.Destroy;
begin
  Clear;
  inherited;
end;

procedure TResourceList.Clear;
var i: integer;
begin
  for i := 0 to Count - 1 do
    FreeAndNil(Items[i]);
  Count := 0;
end;

function TResourceList.FindResource(ResourceId: integer): TResource;
var i: integer;
begin
  Result := nil;
  for i := 0 to Count - 1 do
    if Items[i].ResourceId=ResourceId then begin
      Result := Items[i];
      break;
    end;
end;

function TResourceList.GetResource(ResourceId: integer): TResource;
begin
  Result := FindResource(ResourceId);
  if Result=nil then begin
    if Count >= Length(Items) then
      SetLength(Items, Count + 20);
    Result := ResourceClass.Create;
    Result.ResourceId := ResourceId;
    Items[Count] := Result;
    Inc(Count);
  end;
end;

end.
