unit UGestionErreurExportLSE;

interface
uses Classes,UTOB,Hctrls,HEnt1,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul, fe_main,
{$else}
     eMul, MainEagl,
{$ENDIF}
AglInit
;

type

  TTypeError = (TteTiers,TteArticle,TTeNone);

	TErrorExport = class (Tobject)
  private
  	fTypeError : TTypeError;
    fTiers : string;
    fAuxiliaire : string;
    fNatureAuxi : string;
    fArticle : string;
    fTypeArticle : string;
    fCodeArticle : string;
    fCorrected : boolean;
  public
    constructor create;
    destructor destroy; override;
    property TypeErreur : TTypeError read fTypeError write fTypeError;
    property Tiers : string read fTiers write fTiers;
    property Auxiliaire : string read fAuxiliaire write fAuxiliaire;
    property NatureAuxi : string read fNatureAuxi write fNatureAuxi;
    property Article : string read fArticle write fArticle;
    property TypeArticle : string read fTypeArticle write fTypeArticle;
    property CodeArticle : string read fCodeArticle write fCodeArticle;
    property corrected : boolean read fCorrected write fCorrected;
  end;

  TlistErrorExport = class (Tlist)
  private
    function GetItems(Indice: integer): TerrorExport;
    procedure SetItems(Indice: integer; const Value: TErrorExport);
    function findArticle (Article : string) : boolean;
    function findTiers (Tiers : string) : boolean;
  public
    property Items [Indice : integer] : TErrorExport read GetItems write SetItems;
    function Add(TypeErr: TTypeError; Tiers, Auxiliaire,NatureAuxi, Article,CodeArticle,TypeArticle: string): Integer;
    function InErrorCount : integer;
    destructor destroy; override;
    procedure clear; override;
  end;

function VerifExportAuxiliaire (Code : string ): boolean;
function VerifExportArticle (Code : string ): boolean;

implementation

function VerifExportAuxiliaire (Code : string ): boolean;
var QQ : Tquery;
begin
  result := false;
  QQ := OpenSql ('SELECT YTC_TEXTELIBRE2 FROM TIERSCOMPL WHERE YTC_AUXILIAIRE="'+Code+'"',true,1,'',true);
  if not QQ.eof then
  begin
    if QQ.findField('YTC_TEXTELIBRE2').Asstring <> '' then result := true;
  end;
  ferme (QQ);
end;


function VerifExportArticle (Code : string ): boolean;
var QQ : Tquery;
begin
  result := false;
  QQ := OpenSql ('SELECT GA_CHARLIBRE1 FROM ARTICLE WHERE GA_ARTICLE="'+Code+'"',true,1,'',true);
  if not QQ.eof then
  begin
    if QQ.findField('GA_CHARLIBRE1').Asstring <> '' then result := true;
  end;
  ferme (QQ);
end;


{ TErrorExport }

constructor TErrorExport.create;
begin
	fTypeError := TTeNone;
  fTiers := '';
  fArticle  := '';
  fCorrected := false;
end;

destructor TErrorExport.destroy;
begin
  inherited;
end;

{ TlistErrorExport }

function TlistErrorExport.Add(TypeErr: TTypeError; Tiers,Auxiliaire, NatureAuxi, Article,CodeArticle,TypeArticle: string): Integer;
var Item : TErrorExport;
begin
  result := -1;
  if TypeErr = TteArticle then
  begin
  	if findArticle(Article) then exit;
  end else if TypeErr = TteTiers then
  begin
  	if findTiers(Tiers) then exit;
  end;
	Item := TErrorExport.create;
  Item.TypeErreur := TypeErr;
  Item.Tiers := Tiers ;
  Item.Auxiliaire := Auxiliaire ;
  Item.NatureAuxi  := NatureAuxi ;
  Item.Article := Article;
  Item.CodeArticle := CodeArticle;
  Item.TypeArticle := TypeArticle;
  result := Inherited Add(Item);
end;

procedure TlistErrorExport.clear;
var Indice : integer;
begin
  inherited;
  for Indice := 0 to Count -1 do
  begin
    TerrorExport(Items [Indice]).free;
  end;
end;

destructor TlistErrorExport.destroy;
begin
	clear;
  inherited;
end;

function TlistErrorExport.findArticle(Article: string): boolean;
var indice : integer;
begin
  result := false;
	for Indice := 0 to Count -1 do
  begin
    if (Items [Indice].TypeErreur = TteArticle) and (Items [Indice].Article = Article) then
    begin
    	result := true;
      break;
    end;
  end;
end;

function TlistErrorExport.findTiers(Tiers: string): boolean;
var indice : integer;
begin
  result := false;
	for Indice := 0 to Count -1 do
  begin
    if (Items [Indice].TypeErreur = TteTiers) and (Items [Indice].Tiers  = Tiers) then
    begin
    	result := true;
      break;
    end;
  end;
end;

function TlistErrorExport.GetItems(Indice: integer): TerrorExport;
begin
  result := TerrorExport (Inherited Items[Indice]);
end;

function TlistErrorExport.InErrorCount: integer;
var indice : integer;
begin
	result := 0;
  for Indice := 0 to Count -1 do
  begin
    if not Items [Indice].corrected then inc(result);
  end;
end;

procedure TlistErrorExport.SetItems(Indice: integer; const Value: TErrorExport);
begin
  Inherited Items[Indice]:= Value;
end;

end.
