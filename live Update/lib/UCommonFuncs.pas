unit UCommonFuncs;

interface
uses  Windows, SysUtils, Variants, Classes,
  FileCtrl, UControleUAC, UEncryptage, Uregistry,
  UConstantes,USystemInfo,UfronctionBase,Registry;
type
  TElement = class (Tobject)
  private
    fElement : string;
    fSignature : string;
  public
    property Element : string read fElement write fElement;
    property Signature : string read fSignature write fSignature;
  end;

  TProductList = class (Tlist)
    private
      function GetItems(Indice : integer): TElement;
      procedure SetItems(Indice : integer; const Value: TElement);
      function Add(AObject: TElement): Integer;
    public
      property Items [Indice : integer] : TElement read GetItems write SetItems;
      function Find(Element : string) : TElement;
      procedure clear; override;
      destructor destroy ; override;
  end;

function ISVersionsDIff (TheListExe : TStringList): boolean;
procedure ChargeListeProducts(FileTemp : string;TTL : TproductList);

implementation

procedure ChargeListeProducts(FileTemp : string;TTL : TproductList);
var TT : TStringList;
    TL : TElement;
    II : Integer;
    Chps,CLef,LEct : string;
begin
  TT := TStringList.Create;
  TRY
    TT.LoadFromFile(FileTemp);
    for II := 0 to TT.Count -1 do
    begin
      TL := TElement.Create;
      lect := TT.Strings[II];
      Chps := readtokenST(LEct,SEPARATOR);
      if lect <> '' then
      begin
        CLef := lect;
      end;
      TL.Element := Chps;
      TL.Signature := CLef;
      TTL.Add(TL);
    end;
  FINALLY
    TT.free;
  END;
end;


function ExisteDiff (Tloc,TDis :TProductList; TheListExe : TStringList) : boolean;
var II : Integer;
    LT,DT : TElement;
    theName : string;
begin
  Result := false;
  for II := 0 to Tloc.Count -1 do
  begin
    LT := Tloc.items[II];
    TheName := LT.Element;
    DT := TDis.Find(LT.Element);
    if DT <> nil then
    begin
      if DT.Signature <> LT.Signature then
      begin
        Result := True;
        if ExtractFileExt(TheName)='.exe' then TheListExe.Add(TheName);
      end;
    end;
  end;
end;


function ISVersionsDIff (TheListExe : TStringList): boolean;
var TTL : TProductList; // provenance du serveur
    TLL : TProductList; // locaux
    LTemp : string;
    LLOc : string;
    TheRepo : string;
    TReg : TRegistry;
begin
  result := false;
  PrepareRepository;
  TheListExe.Clear;
  TTL := TProductList.Create;
  TLL := TProductList.Create;
  TRY
    Treg := TRegistry.Create;
    TRY
      with Treg do
      begin
        RootKey := HKEY_LOCAL_MACHINE;
        if not OpenKey ('SOFTWARE\Wow6432Node\Cegid\Cegid Business',False) then Exit;
        TheRepo := IncludeTrailingBackslash(IncludeTrailingBackslash(ReadString('DIRCOPY'))+'REPOSITORY');
        CloseKey;
      end;
    finally
      Treg.Free;
    end;
    // -- Fichiers locaux
    LLOc := IncludeTrailingBackslash (TheRepo)+CONFLSEFILE;
    ChargeListeProducts (LLOc,TLL);
    LTemp := IncludeTrailingBackslash (GetTempDir)+CONFLSEFILE;
    ChargeListeProducts (LTemp,TTL);
    Result := ExisteDiff (TLL,TTL,TheListExe); if Result then Exit;
    //
    TLL.clear;
    TTL.clear;
    LLOc := IncludeTrailingBackslash (TheRepo)+CONFCEGIDFILE;
    ChargeListeProducts (LLOc,TLL);
    LTemp := IncludeTrailingBackslash (GetTempDir)+CONFCEGIDFILE;
    ChargeListeProducts (LTemp,TTL);
    Result := ExisteDiff (TLL,TTL,TheListExe); if Result then Exit;
    //
    TLL.clear;
    TTL.clear;
    LLOc := IncludeTrailingBackslash (TheRepo)+PARAMUPDATE;
    ChargeListeProducts (LLOc,TLL);
    LTemp := IncludeTrailingBackslash (GetTempDir)+PARAMUPDATE;
    ChargeListeProducts (LTemp,TTL);
    Result := ExisteDiff (TLL,TTL,TheListExe); if Result then Exit;
    //
    TLL.clear;
    TTL.clear;
    LLOc := IncludeTrailingBackslash (TheRepo)+CONFSTDFILE;
    ChargeListeProducts (LLOc,TLL);
    LTemp := IncludeTrailingBackslash (GetTempDir)+CONFSTDFILE;
    ChargeListeProducts (LTemp,TTL);
    Result := ExisteDiff (TLL,TTL,TheListExe); if Result then Exit;

    TLL.clear;
    TTL.clear;
    LLOc := IncludeTrailingBackslash (TheRepo)+CONFUPDFILE;
    ChargeListeProducts (LLOc,TLL);
    LTemp := IncludeTrailingBackslash (GetTempDir)+CONFUPDFILE;
    ChargeListeProducts (LTemp,TTL);
    Result := ExisteDiff (TLL,TTL,TheListExe); if Result then Exit;

  FINALLY
    TTL.free;
    TLL.Free;
  END;
end;


{ TProductList }

function TProductList.Add(AObject: TElement): Integer;
begin
  result := inherited Add(Aobject);
end;

procedure TProductList.clear;
var Indice : integer;
begin
  for Indice := count -1 downto 0 do
  begin
    if Items[Indice] <> nil then
    begin
      TElement(Items [Indice]).free;
      Items[Indice] := nil;
    end;
  end;
  inherited;
end;

destructor TProductList.destroy;
begin
  Clear;
  inherited;
end;

function TProductList.Find(Element: string): TElement;
var II : integer;
    TT : TElement;
begin
  Result := nil;
  for II := 0 to Count -1 do
  begin
    TT := Items[II];
    if TT.Element = Element then
    begin
      Result := TT;
      break;
    end;
  end;
end;

function TProductList.GetItems(Indice: integer): TElement;
begin
  result := TElement (Inherited Items[Indice]);
end;

procedure TProductList.SetItems(Indice: integer; const Value: TElement);
begin
  Inherited Items[Indice]:= Value;
end;



end.
