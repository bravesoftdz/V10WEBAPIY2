unit UgestionListe;

interface
Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, fe_main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,MainEagl,
{$ENDIF}
     uTob,
     forms,
     HCtrls,
     HEnt1,
     sysutils,
     ComCtrls;

Type

TChamps = Class(TObject)
	private
		fnom : string;
    flibelle : string; // nom dans la base
    ftype : string;
    flargeur : integer;
    fletitre : string;
    ftablette : string;
    faligment : string;
		FF : string;
    fDec : integer;
    fSep : boolean;
    fObli : boolean;
    fOkLib : boolean;
    fOkVisu : boolean;
    fOkNulle : boolean;
    fOkCumul : boolean;
	public
  	constructor create;
    destructor destroy; override;
    property LeNom : string read fnom write fnom;
    property LeLibelle : string read flibelle write flibelle;
    property TypeData : string read ftype write ftype;
    property largeur : integer read flargeur write flargeur;
    property Titre : string read fletitre write fletitre;
    property aligment : string read faligment write faligment;
    property Format : string read FF write FF;
    property Decim : integer read fDec write fdec;
    property sep : boolean read fsep write fsep;
    property oblig : boolean read fObli write fObli;
    property Oklib : boolean read fOkLib write fOkLib;
    property Okvisu : boolean read fOkVisu write fOkVisu;
    property OkNulle : boolean read fOkNulle write fOkNulle;
    property okcumul : boolean read fOkCumul write fOkCumul;
    property tablette : string read ftablette  write ftablette ;
end;

TlistChamps = class (Tlist)
  private
    function Add(AObject: Tchamps): Integer;
    function GetItems(Indice: integer): Tchamps;
    procedure SetItems(Indice: integer; const Value: TChamps);
  public
    property Items [Indice : integer] : Tchamps read GetItems write SetItems;
    function findchamps (nomchamps : string ): TChamps;
    function remove (UnChamps : Tchamps) : integer;
    procedure clear; override;
end;

TgestionListe = class (TObject)
	private
  	fListChamp : TlistChamps;
  	fListChampInit : TlistChamps;
    fListChampListe : TlistChamps;
  	fOrderBy : string;
    fListe : string;
    fCaption : string;
    FNomTable, FLien, FSortBy, fstCols : string;
    Ftitre : Hstring;
    FLargeur, FAlignement, FParams, FPerso: string;
    ftitle, fNC : Hstring;
    fOkTri, fOkNumCol : boolean;
    fnbcols : integer;
    procedure ChargeListChamps;
    procedure ChargeListChampsliste;
  public
  	constructor create;
    destructor destroy;  override;
    property ListChamps : TlistChamps read fListChamp write fListChamp;
    property ListChampsListe : TlistChamps read fListChampListe write fListChampListe;
    property Orderby : string read fOrderBy write fOrderBy;
    property Liste: string read fListe write fliste;
    property caption : string read fCaption write fcaption;
    property Nomtable : string read FNomTable write FNomTable;
    property Lien : string read FLien write FLien;
    property sort : string read FSortBy write FSortBy;
    property titre : Hstring read ftitre write ftitre;
    function ChargeListe (NomListe : string) : boolean;
end;

function TriLibelle (Champs1,Champs2 : pointer) : integer;
function TriName (Champs1,Champs2 : pointer) : integer;

implementation


function TriLibelle(Champs1, Champs2: pointer) : integer;
begin
	Result := CompareText(TChamps(champs1).LeLibelle, TChamps(Champs2).LeLibelle);
end;

function TriName(Champs1, Champs2: pointer) : integer;
begin
	Result := CompareText(TChamps(champs1).LeNom, TChamps(Champs2).LeNom);
end;

{ TgestionListe }

function TgestionListe.ChargeListe(NomListe: string): boolean;
begin
	result := false;
  fListe := Nomliste;
  ChargeHListe (fListe, FNomTable, FLien, FSortBy, fstCols, FTitre,
                FLargeur, FAlignement, FParams, ftitle, fNC, FPerso, fOkTri, fOkNumCol);
  ChargeListChamps;
  fListChampInit.Assign(fListChamp ,lacopy); // recopy de la liste dans init
  ChargeListChampsliste;
end;

procedure TgestionListe.ChargeListChamps;
var req : string;
		QQ : Tquery;
    Indice : integer;
    fChamps : TChamps;
begin
  Req := 'SELECT * FROM '+FNomTable+' WHERE 1=2'; // pour recup la structure
  QQ := OpenSql (Req,true,-1,'',true);
  for indice := 0 to QQ.FieldCount - 1 do
  begin
  	fChamps := TChamps.create;
    fChamps.fnom   := QQ.Fields[indice].FieldName;
    fChamps.ftype := ChampToType(fChamps.fnom);
    if fchamps.ftype = 'COMBO' then fchamps.ftablette := ChampToTT(fchamps.Fnom);
    fChamps.flibelle  := ChampToLibelle (fChamps.fnom);
    fChamps.fletitre  := fChamps.flibelle;
    fListChamp.Add(fChamps);
  end;
  ferme(QQ);
end;

procedure TgestionListe.ChargeListChampsliste;
var laliste,lelement : string;
		UnChamps : Tchamps;
    alignement,lalargeur,letitre,FF,lesaligments,leslargeurs,lestitres : string;
    Dec : integer;
    Sep,Obli,OkLib,OkVisu,Oknulle,OkCumul : boolean;
begin
  laliste := fstcols;
  lesaligments := FAlignement;
  leslargeurs := FLargeur;
  lestitres := Ftitre;
  repeat
    lelement := READTOKENST (laliste);
    if lelement <> '' then
    begin
      alignement := ReadTokenSt(lesaligments);
      lalargeur := readtokenst(leslargeurs);
      letitre := readtokenst(lestitres);
      TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
      unChamps := fListChamp.findChamps(lelement);
      if UnCHamps<>nil then
      begin
      	flistChamp.remove(unChamps);
        fListChampListe.add(unChamps);
      end else
      begin
      	unChamps := Tchamps.create;
        UnChamps.fnom := lelement;
        UnChamps.ftype := '<<FORMULE>>';
        UnChamps.flibelle := 'Formule';
      end;
      Unchamps.flargeur := strtoint(lalargeur);
      Unchamps.Titre := letitre;
      Unchamps.fDec := Dec;
      UnChamps.FF := FF;
      UnChamps.fSep := Sep;
      UnChamps.fObli := Obli;
      UnChamps.fOkLib := OkLib;
      UnCHamps.fOkVisu := OkVisu;
      UnChamps.fOkNulle := OkNulle;
      UnChamps.fOkCumul := OkCumul;
    end;
  until laliste = '';
end;

constructor TgestionListe.create;
begin
  fOrderBy := '';
  fListe := '';
  fCaption :='';
  FNomTable := '';
  FLien:= '';
  FSortBy := '';
  fstCols := '';
  Ftitre :='';
  FLargeur := '';
  FAlignement := '';
  FParams := '';
  FPerso:= '';
  ftitle := '';
  fNC :='';
  fOkTri := false;
  fOkNumCol :=false;
  fnbcols:=0;
	fListChampInit := TlistChamps.create;
	fListChamp  := TlistChamps.create;
  fListChampListe := TlistChamps.create;
end;


destructor TgestionListe.destroy;
begin
	fListChampInit.clear;
  flistChampInit.free;
  fListChamp.clear;
  fListChamp.free;
  fListChampListe.clear;
  fListChampListe.free;
  inherited;
end;

{ TChamps }

constructor TChamps.create;
begin
  fnom :='';
  ftype :='';
  flargeur :=0;
  fletitre :='';
  ftablette :='';
  faligment :='';
  FF :='';
  fDec :=0;
  fSep :=false;
  fObli :=false;
  fOkLib :=false;
  fOkVisu :=true;
  fOkNulle :=false;
  fOkCumul :=false;

end;

destructor TChamps.destroy;
begin

  inherited;
end;

{ TlistChamps }

function TlistChamps.Add(AObject: Tchamps): Integer;
begin
  result := Inherited Add(AObject);
end;

procedure TlistChamps.clear;
var indice : integer;
begin
  if count > 0 then
  begin
    for Indice := 0 to count -1 do
    begin
      TChamps(Items [Indice]).free;
    end;
  end;
  inherited;
end;

function TlistChamps.findchamps(nomchamps: string): TChamps;
var Indice : integer;
begin
  result := nil;
  for Indice := 0 to Count -1 do
  begin
    if Items[Indice].fnom  = nomchamps then
    begin
      result:=Items[Indice];
      break;
    end;
  end;
end;

function TlistChamps.GetItems(Indice: integer): Tchamps;
begin
  result := TChamps (Inherited Items[Indice]);
end;

function TlistChamps.remove(UnChamps: Tchamps): integer;
begin
	result := inherited remove(Unchamps);
end;

procedure TlistChamps.SetItems(Indice: integer; const Value: TChamps);
begin
  Inherited Items[Indice]:= Value;
end;

end.
