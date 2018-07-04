unit UverifServeur;


interface
uses Windows,SysUtils,Registry,Classes,XmlDoc,XMLIntf,forms,UcontroleUAC,UFolders,UXmlUpdInfo,Uregistry;

type

	TSoft = class
  	private
    	fNomSoft : string;
      fCode : string;
      fInstallversion : string;
      fLocalversion : string;
      fServerversion : string;
      fAjour : boolean;
    public
    	property Code : string read fcode write fcode;
      property Nom : string read fNomSoft write fNomSoft;
      property Installversion : string read fInstallversion write fInstallversion;
      property Localversion : string read fLocalversion write fLocalversion;
      property Serverversion : string read fServerversion write fServerversion;
      property Ajour : boolean read fAjour write fajour;
    	constructor create;
      destructor destroy; override;
  end;

  TlistSoft = class (Tlist)
  private
    function GetItems(Indice: integer): TSoft;
    procedure SetItems(Indice: integer; const Value: TSoft);
  public
    procedure clear; override;
    function Add(AObject: TSoft): Integer;
    property Items [Indice : integer] : TSoft read GetItems write SetItems;
    function findSoft(nomSoft: string): TSoft;
    destructor destroy; override;
  end;

	TControleversion = class
  private
    fServerLocation : string;
    fLocalLocation : string;
    fListSofts : TlistSoft;
    XmlUpdInfo : TXmlUpdInfo;
    function GetXmlversion (nomXml : string) : string;
    procedure GetLocalLocation;
  public
  	property ServerLocation : string read FServerLocation write fServerlocation;
  	property LocalLocation : string read FLocalLocation write fLocalLocation;
  	procedure GetServerLocation;
    procedure SetinstalledListSoft;
    procedure  getInfoServeur;
    procedure SetMaj;
  	constructor create;
    Destructor destroy; override;
  end;


procedure SetInfoMaj;

implementation

procedure SetInfoMaj;
var ControlVersions : TControleVersion;
begin
	ControlVersions := TControleVersion.create;
  TRY
  	with ControlVersions do
    begin
    	GetServerLocation;
    	GetLocalLocation;
    	if fServerLocation = '' then exit;
    	SetinstalledListSoft;
    	getInfoServeur;
      SetMaj;
    end;
  FINALLY
  	ControlVersions.free;
  END;
end;

{ TControleversion }

constructor TControleversion.create;
begin
	fListSofts := TlistSoft.create;
  XmlUpdInfo := TXmlUpdInfo.create;
  XmlUpdInfo.NewXml;
end;


destructor TControleversion.destroy;
begin
	fListSofts.free;
  XmlUpdInfo.free;
  inherited;
end;

procedure TControleversion.getInfoServeur;
var repert : string;
		indice : integer;
    soft : string;
begin
	repert := IncludeTrailingBackslash(fServerLocation) + 'KIT\PRODUCTS';
  for Indice := 0 to fListSofts.Count - 1 do
  begin
     soft := IncludeTrailingBackslash ( repert)+flistSofts.Items [indice].Code+'.xml';
     if FileExists(Soft) then
     begin
       flistSofts.Items[indice].Serverversion  := GetXmlversion (Soft);
     end;
  end;
end;

procedure TControleversion.GetLocalLocation;
var reg : TRegistry;
begin
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('Software\LSE\AutoUpdate\', False) then
    begin
      fLocalLocation := Reg.ReadString('LocalFolder');
    end;
  finally
  	reg.free;
  end;
end;

procedure TControleversion.GetServerLocation;
var reg : TRegistry;
begin
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('Software\LSE\AutoUpdate\', False) then
    begin
      fServerLocation := Reg.ReadString('ServerFolder');
    end;
  finally
  	reg.free;
  end;
end;

function TControleversion.GetXmlversion(nomXml: string): string;
var product : IXMLNode ;
		DOMDOC : TXmlDocument;
begin
	DomDOc := TXmlDocument.create(Application);
  TRY
    With DOMDoc do
    begin
      domdoc.LoadFromFile(NomXml);
      if domdoc.IsEmptyDoc then exit;
      Active := true;
      product := DomDoc.DocumentElement.ChildNodes.FindNode('PRODUCT');
      result := product.GetAttributeNS('version','');
    end;
  FINALLY
  	domdoc.free;
  end;
end;

procedure TControleversion.SetinstalledListSoft;
var reg,reg1 : Tregistry;
		list : TStringList;
    indice : integer;
    key : string;
    OneSoft : TSoft;
begin
  list := TStringList.Create;
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('Software\Cegid\Cegid business\', False) then
    begin
      reg.GetKeyNames(list);
      for Indice := 0 to list.Count -1 do
      begin
        key := list.Strings[indice];
        Reg1 := TRegistry.Create;
        Reg1.RootKey := HKEY_LOCAL_MACHINE;
        if Reg1.OpenKey('Software\Cegid\Cegid business\'+Key+'\', False) then
        begin
          if Reg1.ValueExists('Version') then
          begin
            OneSoft := fListSofts.findSoft(Key);
            if OneSoft = nil then
            begin
              OneSoft := Tsoft.create;
              OneSoft.Nom := Key;
            end;
            OneSoft.Installversion := Reg1.ReadString('version');
            OneSoft.Code  := Reg1.ReadString('Code');
            fListSofts.Add(OneSoft);
          end;
        end;
        reg1.Free;
      end;
    end;
  finally
    Reg.CloseKey;
    Reg.Free;
    List.free;
  end;
end;


procedure TControleversion.SetMaj;
var Indice : integer;
		TheSoft : Tsoft;
begin
	for Indice := 0 to fListSofts.Count -1 do
  begin
    TheSoft := fListSofts.Items [indice];
    if TheSoft.Serverversion > theSoft.Installversion then
    begin
    	SetInfoStocke ('ServerUpdated','1');
      XmlUpdInfo.Addproduct(TheSoft.Nom ,TheSoft.Code, TheSoft.Installversion ,TheSoft.Serverversion,'S');
    end else if GetInfoStocke('Updmethod')='0' then
    begin
    	if TheSoft.localversion > theSoft.Installversion then
      begin
        SetInfoStocke ('ClientUpdated','1');
      	XmlUpdInfo.Addproduct(TheSoft.Nom ,TheSoft.Code,TheSoft.Installversion ,TheSoft.localversion,'C');
      end;
  	end;
  end;
  XmlUpdInfo.SaveXml;
end;

{ TlistSoft }

function TlistSoft.Add(AObject: TSoft): Integer;
begin
  result := Inherited Add(AObject);
end;

procedure TlistSoft.clear;
var indice : integer;
begin
  for Indice := 0 to Count -1 do
  begin
    TSoft(Items [Indice]).free;
  end;
  inherited;
end;


destructor TlistSoft.destroy;
begin
  clear;
  inherited;
end;

function TlistSoft.findSoft(nomSoft: string): TSoft;
var Indice : integer;
begin
  result := nil;
  for Indice := 0 to Count -1 do
  begin
    if Items[Indice].fNomSoft = nomSoft then
    begin
      result:=Items[Indice];
      break;
    end;
  end;
end;

function TlistSoft.GetItems(Indice: integer): TSoft;
begin
  result := TSoft (Inherited Items[Indice]);
end;


procedure TlistSoft.SetItems(Indice: integer; const Value: TSoft);
begin
   Inherited Items[Indice]:= Value;
end;

{ TSoft }

constructor TSoft.create;
begin
	fCode  := '';
  fNomSoft  := '';
  fInstallversion := '';
  fServerversion := '';
  fAjour := true;
end;

destructor TSoft.destroy;
begin
  //
  inherited;
end;

end.
