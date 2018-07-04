unit FactLivrFromRecep;

interface
uses sysutils,classes,windows,messages,controls,forms,hmsgbox,stdCtrls,clipbrd,nomenUtil,
     HCtrls,SaisUtil,HEnt1,Ent1,EntGC,UtilPGI,UTOB,HTB97,FactUtil,FactComm,Menus,
     FactTOB, FactPiece, FactArticle,FactAdresseBTP,ParamSoc,CalcOlegenericAff,
{$IFDEF EAGLCLIENT}
  maineagl,
{$ELSE}
  fe_main,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  FactVariante,
{$IFDEF BTP}
     BTPUtil,UtilPhases,
{$ENDIF}
     HRichOLE;

type

TLivraisonFromRecepStock = class
  private
    FF : TForm;
  	fTOBpiece : TOB;
    fTobAffaire : TOB;
    fUsable : boolean;
    TheAffaire : string;
    fautorisee : boolean;
    procedure SetAffaire(const Value: TOB);
    function GetNbrRecep: boolean;
    function IsAutorise : boolean;
  public
  	constructor create (TT : TForm);
    destructor destroy; override;
    property Document : TOB read fTOBPiece write fTOBPiece;
    property Affaire : TOB read fTobAffaire write SetAffaire;
    Property IsExists : boolean read GetNbrRecep;
    procedure ChangeAffaire;
end;

implementation

uses facture;

{ TLivraisonFromRecepStock }

procedure TLivraisonFromRecepStock.ChangeAffaire;
begin
	if (fTOBAffaire <> nil) and (IsAutorise) then
  begin
		if FF is TFFacture then TFFacture(FF).TheConso.recupereReceptionsHorsLien (fTOBpiece,fTOBaffaire);
  end;
end;

constructor TLivraisonFromRecepStock.create (TT : TForm);
begin
	FF := TT;
	fusable := false;
  TheAffaire := '';
end;

destructor TLivraisonFromRecepStock.destroy;
begin
  inherited;
end;

function TLivraisonFromRecepStock.GetNbrRecep: boolean;
begin
	result:=False;
	if fautorisee then result := (TFFacture(FF).TheConso.NbrReceptionsHorsLien > 0);
end;

function TLivraisonFromRecepStock.IsAutorise: boolean;
begin
	result := (fTOBaffaire <> nil) and
  	 			  (fTOBaffaire.getValue('AFF_AFFAIRE') <> '') AND
     				(fTOBaffaire.getValue('AFF_AFFAIRE') <> TheAffaire);
  fAutorisee := result;
  if fautorisee then TheAffaire := fTOBaffaire.getValue('AFF_AFFAIRE');
end;

procedure TLivraisonFromRecepStock.SetAffaire(const Value: TOB);
begin
  fTobAffaire := Value;
  if fTOBaffaire = nil then exit;
  if fTOBaffaire.getValue('AFF_AFFAIRE') = '' then exit;
  if TheAffaire <> '' then exit;
  TheAffaire := fTOBaffaire.getValue('AFF_AFFAIRE');
  TFFacture(FF).TheConso.recupereReceptionsHorsLien (fTOBpiece,fTOBaffaire);
end;

end.
 
