unit UgestionINI;

interface
uses Classes, IniFiles,CBPPath,UTOB,Udefinitions ;

type

	TEnreg = class(TObject)
  	private
      fId : string;
      fRessource : string;
      fServer : string;
      fDataBase : string;
      fuser : string;
      fLibelle : string;
      fPassword : string;
    public
      property ID : string read fId write fId;
      property Ressource : string read fRessource write fRessource;
      property Server : string read fServer write fServer;
      property Database : string read fDataBase write fDataBase;
      property User : string read fuser write fuser;
      property Password : string read fPassword write fPassword;
      property Libelle : string Read fLibelle write fLibelle;
      constructor Create;
  end;

  TListeEnreg = class (TList)
  	private
      function Add(AObject: TEnreg): Integer;
      function GetItems(Indice: integer): TEnreg;
      procedure SetItems(Indice: integer; const Value: TEnreg);
    public
      constructor create;
    	destructor destroy; override;
      property Items [Indice : integer] : TEnreg read GetItems write SetItems;
      function findID (ID : string ): TEnreg;
      procedure clear; override;
  end;

  TGestionINI = class (TObject)
  	private
  		FIniFile: TIniFile;
      FOpened : Boolean;
      fListeEnreg : TListeEnreg;
      fTOBList : TOB;
      procedure ConstitueListe;
      procedure AjouteUnEnreg (NomSection :string);
      procedure AjouteUneFille(UnEnreg : Tenreg);
      procedure AjoutelesChamps(TOBL : TOB);
    public
      constructor create;
      destructor destroy; override;
      function GetList : TListeEnreg;
      function GetTOBList : TOb;
      function FindId (ID : string) : TEnreg;
    	function EcritEnreg(TheEnreg: Tenreg; Mode: TactionFiche): boolean;
      function SupprimeEnreg (TheEnreg : TEnreg) : Boolean;
  end;

implementation

uses SysUtils;

function TGestionINI.EcritEnreg (TheEnreg : Tenreg ; Mode : TactionFiche) : boolean;
begin
  Result := true;
  TRY
    FiniFile.WriteString(TheEnreg.ID,'ID',TheEnreg.ID);
    FiniFile.WriteString(TheEnreg.ID,'SERVER',TheEnreg.Server);
    FiniFile.WriteString(TheEnreg.ID,'DATABASE',TheEnreg.Database);
    FiniFile.WriteString(TheEnreg.ID,'USER',TheEnreg.user);
    FiniFile.WriteString(TheEnreg.ID,'RESSOURCE',TheEnreg.Ressource);
    FiniFile.WriteString(TheEnreg.ID,'LIBELLE',TheEnreg.Libelle);
    FiniFile.WriteString(TheEnreg.ID,'PASSWORD',TheEnreg.Password);
    if Mode= TaCreat then
    begin
      AjouteUnEnreg(TheEnreg.ID);
    end;
  except
    Result := false;
  end;
end;

{ TGestionINI }

procedure TGestionINI.AjoutelesChamps(TOBL: TOB);
begin
	TOBL.AddChampSupValeur('SEL','');
	TOBL.AddChampSupValeur('ID','');
	TOBL.AddChampSupValeur('LIBELLE','');
	TOBL.AddChampSupValeur('RESSOURCE','');
	TOBL.AddChampSupValeur('SERVER','');
	TOBL.AddChampSupValeur('DATABASE','');
	TOBL.AddChampSupValeur('PASSWORD','');
	TOBL.AddChampSupValeur('USER','');
end;

procedure TGestionINI.AjouteUneFille(UnEnreg: Tenreg);
var TOBL : TOB;
begin
	TOBL := TOB.Create('UNE FILLE',fTOBList,-1);
  AjoutelesChamps(TOBL);
  TOBL.SetString('ID',UnEnreg.ID);
  TOBL.SetString('LIBELLE',UnEnreg.Libelle);
  TOBL.SetString('RESSOURCE',UnEnreg.Ressource);
  TOBL.SetString('SERVER',UnEnreg.Server);
  TOBL.SetString('DATABASE',UnEnreg.Database);
  TOBL.SetString('USER',UnEnreg.user);
  TOBL.SetString('PASSWORD',UnEnreg.Password);
end;

procedure TGestionINI.AjouteUnEnreg(NomSection: string);
var LesChamps : TStringList;
    Indice : integer;
    Ligne,Champ,Valeur : string;
    UnEnreg : TEnreg;
    LeNom : string;
begin
  LeNom := UpperCase (NomSection);
  lesChamps := TStringList.Create;
  //
  FiniFile.ReadSectionValues (LeNom,LesChamps);
  if lesChamps.Count > 0 then
  begin
    UnEnreg := TEnreg.create;
    UnEnreg.ID := LeNom;
    // Ajout d'un descripteur de fichier
    for Indice :=0 to lesChamps.Count -1 do
    begin
      // Recup des infos
      Ligne := lesChamps.Strings [Indice];
      Champ := Copy(Ligne,1,Pos('=',Ligne)-1);
      Valeur := Copy(Ligne,Pos('=',Ligne)+1,Length(Ligne));
      // desciption du fichier
      if Champ = 'ID' then UnEnreg.ID := Valeur else
      if Champ = 'SERVER' then UnEnreg.Server := Valeur else
      if Champ = 'DATABASE' then UnEnreg.Database  := Valeur else
      if Champ = 'USER' then UnEnreg.User := Valeur else
      if Champ = 'RESSOURCE' then UnEnreg.Ressource := Valeur else
      if Champ = 'LIBELLE' then UnEnreg.Libelle := Valeur;
      if Champ = 'PASSWORD' then UnEnreg.Password := Valeur;
    end;
    fListeEnreg.Add (UnEnreg);
    AjouteUneFille(UnEnreg);
  end;
  //
 LesChamps.Clear;
 LesChamps.free;
end;

procedure TGestionINI.ConstitueListe;
var Indice : integer;
    Sections : TStringList;
begin
  fTOBList.ClearDetail;
  fListeEnreg.clear;
  //
  Sections := TStringList.Create;
  // Lecture des Sections
  FiniFile.ReadSections (Sections);
  for Indice := 0 to Sections.Count -1 do
  begin
    AjouteUnEnreg (Uppercase(Sections.strings[Indice]));
  end;
  Sections.Clear;
  Sections.free;
end;

constructor TGestionINI.create;
var ffiledescriptor : string;
begin
	fListeEnreg := TListeEnreg.create;
  fTOBList := TOB.Create('LA LISTE DES ID',nil,-1);
  FOpened := false;
  (*
  ffiledescriptor :=IncludeTrailingBackslash(IncludeTrailingBackslash (TcbpPath.GetCegidDataDistri)+'WebServices')+'Definitions.ini';
  *)
  ffiledescriptor := IncludeTrailingBackslash(getCheminIni)+'Definitions.ini';
	FIniFile := TIniFile.Create (ffiledescriptor);
  if FIniFile <> nil then FOpened := True;
  if FOpened then
  begin
  	ConstitueListe;
  end;
end;

destructor TGestionINI.destroy;
begin
  if Finifile <> nil then FreeAndNil(FIniFile);
	if fListeEnreg <> nil then FreeAndnil(fListeEnreg);
  if fTOBList <> nil then freeAndNil(fTOBList);
  inherited;
end;

function TGestionINI.FindId(ID: string): TEnreg;
begin
	Result := fListeEnreg.findID(ID);
end;

function TGestionINI.GetList: TListeEnreg;
begin
	Result := fListeEnreg;
end;

function TGestionINI.GetTOBList: TOb;
begin
	Result := fTOBList;
end;

function TGestionINI.SupprimeEnreg(TheEnreg: TEnreg): Boolean;
begin
	FIniFile.EraseSection (TheEnreg.ID);
	ConstitueListe;
end;

{ TListeEnreg }

function TListeEnreg.Add(AObject: TEnreg): Integer;
begin
	Result := inherited ADD(Aobject);
end;

procedure TListeEnreg.clear;
var indice : integer;
begin
  if count > 0 then
  begin
    for Indice := count -1 downto 0 do
    begin
      if TEnreg(Items [Indice])<> nil then
      begin
         TEnreg (Items [Indice]).free;
         Items[Indice] := nil;
      end;
    end;
  end;
  Pack;
  inherited;
end;

constructor TListeEnreg.create;
begin
//
end;

destructor TListeEnreg.destroy;
begin
	clear;
  inherited;
end;

function TListeEnreg.findID(ID: string): TEnreg;
var indice : Integer;
begin
	Result := nil;
  for Indice := 0 to count -1 do
  begin
    if TEnreg(Items [Indice])<> nil then
    begin
       if TEnreg (Items [Indice]).ID = ID then
       begin
         Result := Items [Indice];
         break;
       end;
    end;
  end;
end;

function TListeEnreg.GetItems(Indice: integer): TEnreg;
begin
  result := TEnreg (Inherited Items[Indice]);
end;

procedure TListeEnreg.SetItems(Indice: integer; const Value: TEnreg);
begin
  Inherited Items[Indice]:= Value;
end;

{ TEnreg }

constructor TEnreg.Create;
begin
  fId :='';
  fRessource :='';
  fServer :='';
  fDataBase :='';
  fuser :='';
  fLibelle :='';
  fPassword :='';
end;

end.
