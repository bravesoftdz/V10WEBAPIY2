unit UCotraitanceOuv;

interface
Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     MaineAGL,
{$ENDIF}
     Hent1,
     EntGc,
		 Menus,
		 HCtrls,
     HMsgBox,
     SysUtils,
     AglInit,
     forms,
     UentCommun,
     uTob;
type

	TpieceCotraitOuv = class (Tobject)
  	private
    	fusable : boolean;
      fCreatedPop : boolean;
      fmaxItems : integer;
    	FF : TForm;
      fTOBpiece : TOB;
      fTOBligne : TOB;
      fTOBaffaire : TOB;
      fTOBOuvrage : TOB;
      GS : THGrid;
      SGN_NATURETRAVAIL : integer;
      SGN_LIBFOU : integer;
      fPOPGS : TPopupMenu;
      fMenuItem: array[0..2] of TMenuItem;
    	procedure DefiniMenuPop(Parent: Tform);
      procedure GSAffectCotrait (Sender : tobject);
      procedure GSAffectInterne (Sender : tobject);
    	procedure Setaffaire(const Value: TOB);
    	procedure MenuEnabled(State: boolean);
      procedure AffecteCoTraitance(Fournisseur: string; Indice: integer);
    	procedure AffecteInterne(Indice: integer);
    	procedure AffecteCoTraitanceSousDetail(TOBOUV: TOB; Fournisseur,LibelleFou,NatureTravail: string);
    	procedure AffecteInterneSousDetail(TOBOUV: TOB);
    	procedure SetGrilleSaisie(State: boolean);
    	procedure Setusable(const Value: boolean);
    	function NoMoreCotraitance(fTOBLigne: TOB): boolean;
    public
    	constructor create (laForm : Tform);
      destructor destroy; override;
      property Affaire : TOB read fTOBAffaire write Setaffaire;
      property TOBpiece : TOB read fTOBpiece write fTOBpiece;
      property TOBOuvrage:TOB read fTOBouvrage write fTOBouvrage;
      property InUse : boolean read fusable write Setusable;
      property Coltravail : integer read SGN_NATURETRAVAIL write SGN_NATURETRAVAIL;
      property CollibFou : integer read SGN_LIBFOU write SGN_LIBFOU;
  end;

var PieceTraitUsable : boolean;

implementation

uses FactTOB,LigNomen,factOuvrage,facture,UCotraitance;


{ TpieceCotraitOuv }

constructor TpieceCotraitOuv.create ( LaForm : Tform);
var ThePop : Tcomponent;
begin
	fusable := true; // par défaut
	if not VH_GC.SeriaCoTraitance then fusable := false; // meme pas la peine de continuer
  FF := laForm;
	if (FF is TFLigNomen) then
  begin
    fusable := false;
    GS := TFligNomen(FF).G_NLIG;
  end else fusable := false;
  //
  ThePop := FF.Findcomponent  ('PopupG_NLIG');
  if ThePop <> nil then
  begin
    fCreatedPop := false;
    fPOPGS := TPopupMenu(thePop);
  end;
  DefiniMenuPop(FF);
  PieceTraitUsable := fusable;
  SGN_NATURETRAVAIL := -1;
end;

destructor TpieceCotraitOuv.destroy;
var indice : integer;
begin
  for Indice := 0 to fMaxItems -1 do
  begin
  	if fMenuItem[Indice]<> nil then fMenuItem[Indice].Free;
  end;
	if fcreatedPop then fPOPGS.free;
  PieceTraitUsable := false;
  inherited;
end;

procedure TpieceCotraitOuv.DefiniMenuPop (Parent : Tform);
var Indice : integer;
begin
  fMaxItems := 0;
  if not fcreatedPop then
  begin
    fMenuItem[fMaxItems] := TmenuItem.Create (parent);
    with fMenuItem[fMaxItems] do
      begin
    	Name := 'BCOTRAITSEP';
      Caption := '-';
      Visible := false;
      end;
    inc (fMaxItems);
  end;
  // Affectation à un cotraitant
  fMenuItem[fMaxItems] := TmenuItem.Create (parent);
  with fMenuItem[fMaxItems] do
    begin
    Name := 'BCOTRAITAFFECT';
    Caption := 'Affectation à un cotraitant';
    Visible := false;
    OnClick := GSAffectCotrait;
    end;
  inc (fMaxItems);
  // Récupération des travaux
  fMenuItem[fMaxItems] := TmenuItem.Create (parent);
  with fMenuItem[fMaxItems] do
    begin
    Name := 'BCOINTERNE';
//    Caption := 'Travaux effectués par nous même';
    Caption := 'Affectation en interne';
    Visible := false;
    OnClick := GSAffectInterne;
    end;
  inc (fMaxItems);

  for Indice := 0 to fMaxItems -1 do
  begin
  	if fMenuItem [Indice] <> nil then fPOPGS.Items.Add (fMenuItem[Indice]);
  end;
end;

procedure  TpieceCotraitOuv.AffecteCoTraitanceSousDetail(TOBOUV: TOB;Fournisseur,LibelleFou,NatureTravail : string);
var TOBOL : TOB;
		Indice : integer;
begin
	for Indice := 0 to TOBOUV.detail.count -1 do
  begin
		TOBOL := TOBOUV.detail[Indice];
    TOBOL.putValue('BLO_FOURNISSEUR',Fournisseur);
  	TOBOL.putValue('GA_FOURNPRINC',Fournisseur);
    TOBOL.putValue('LIBELLEFOU',LibelleFou);
    TOBOL.putValue('BLO_NATURETRAVAIL',NatureTravail);
    TOBOL.putValue('BLO_DPA',TOBOL.GetValue('BLO_PUHT'));
    TOBOL.putValue('BLO_DPR',TOBOL.GetValue('BLO_PUHT'));
    TOBOL.putValue('BLO_COEFFR',0);
    TOBOL.putValue('BLO_COEFFG',0);
    TOBOL.putValue('BLO_COEFFC',0);
    TOBOL.putValue('BLO_COEFMARG',1);
    TOBOL.putValue('BLO_NONAPPLICFRAIS','X');
    TOBOL.putValue('BLO_NONAPPLICFC','X');
    TOBOL.putValue('BLO_NONAPPLICFG','X');
    if TOBOL.detail.count > 0 then
    begin
    	AffecteCoTraitanceSousDetail (TOBOL,Fournisseur,LibelleFou,NatureTravail);
    end;
  end;
end;

procedure TpieceCotraitOuv.AffecteCoTraitance (Fournisseur: string;Indice: integer);
var TOBL : TOB;
		LibelleFou : string;
begin
  fTOBligne := TFLigNomen (FF).TOBLpiece;
  TOBL := fTOBOuvrage.detail[Indice-1]; if TOBL=nil then exit;
  LibelleFou := GetLibIntervenant (Fournisseur);
  TOBL.putvalue('BLO_NATURETRAVAIL','001');
  TOBL.putvalue('BLO_FOURNISSEUR',Fournisseur);
  TOBL.putValue('GA_FOURNPRINC',Fournisseur);
  TOBL.putvalue('LIBELLEFOU',LibelleFou);
  TOBL.putvalue('BLO_DPA',TOBL.getValue('BLO_PUHT'));
  TOBL.putvalue('BLO_DPR',TOBL.getValue('BLO_PUHT'));
  TOBL.putvalue('BLO_NONAPPLICFC','X');
  TOBL.putvalue('BLO_NONAPPLICFG','X');
  TOBL.putvalue('BLO_NONAPPLICFRAIS','X');
  TOBL.putvalue('BLO_COEFFC',0.0);
  TOBL.putvalue('BLO_COEFFR',0.0);
  TOBL.putvalue('BLO_COEFFG',0.0);
  TOBL.putvalue('BLO_COEFMARG',1.0);
  fTOBLigne.putvalue('GL_RECALCULER','X');
  fTOBLigne.putvalue('GLC_NATURETRAVAIL','010');
  fTOBLigne.putvalue('GL_FOURNISSEUR','');
  fTOBPiece.putvalue('GP_RECALCULER','X');
  if IsOuvrage(TOBL) then
  begin
    AffecteCoTraitanceSousDetail (TOBL,Fournisseur,LibelleFou,'001');
  end;
  TFligNomen(FF).AfficheLaligne(indice);
end;

procedure TpieceCotraitOuv.GSAffectCotrait(Sender: tobject);
var retour : string;
		indice : integer;
begin
	if GS.nbSelected = 0 then BEGIN PGiInfo('Aucune ligne sélectionnée');exit; END;
  // Sélection du cotraitant
  retour := AGLLanceFiche('BTP','BTMULAFFAIREINTER','BAI_AFFAIRE='+fTOBAffaire.getValue('AFF_AFFAIRE'),'','');
  // --
  if retour <> '' then
  begin
    for Indice := 1 to GS.RowCount do
    begin
      if (GS.IsSelected (Indice) ) then
      begin
        AffecteCoTraitance (Retour,Indice);
        TFLigNomen(FF).AfficheLaLigne(Indice)
      end;
    end;
  end;
  GS.AllSelected := false;
end;

procedure TpieceCotraitOuv.MenuEnabled(State: boolean);
var Indice : integer;
begin
  for Indice := 0 to fMaxItems -1 do
  begin
    fMenuItem[Indice].visible := State;
  end;
end;

procedure TpieceCotraitOuv.Setaffaire(const Value: TOB);
begin
  fTOBAffaire := Value;
end;

procedure TpieceCotraitOuv.AffecteInterneSousDetail(TOBOUV : TOB);
var TOBOL : TOB;
		Indice : integer;
begin
	for Indice := 0 to TOBOUV.detail.count -1 do
  begin
		TOBOL := TOBOUV.detail[Indice];
    TOBOL.putValue('BLO_FOURNISSEUR','');
    TOBOL.putValue('GA_FOURNPRINC','');
    TOBOL.putValue('LIBELLEFOU','');
    TOBOL.putValue('BLO_NATURETRAVAIL','');
    TOBOL.putValue('BLO_DPA',TOBOL.GetValue('BLO_PUHT'));
    TOBOL.putValue('BLO_DPR',TOBOL.GetValue('BLO_PUHT'));
    TOBOL.putValue('BLO_COEFFR',0);
    TOBOL.putValue('BLO_COEFFG',0);
    TOBOL.putValue('BLO_COEFFC',0);
    TOBOL.putValue('BLO_COEFMARG',1);
    TOBOL.putValue('BLO_NONAPPLICFRAIS','-');
    TOBOL.putValue('BLO_NONAPPLICFC','-');
    TOBOL.putValue('BLO_NONAPPLICFG','-');
    if TOBOL.detail.count > 0 then
    begin
    	AffecteInterneSousDetail (TOBOL);
    end;
  end;
end;

function TpieceCotraitOuv.NoMoreCotraitance(fTOBLigne : TOB) : boolean;
var fTOBouvrage,TOBP : TOB;
		indiceNomen,Indice : integer;
begin
	result := true;
  if Isouvrage(fTOBLigne) then
  begin
    if TFligNomen(FF).Thefacture <> nil then
    begin
      fTOBOuvrage := TFfacture(TFligNomen(FF).Thefacture).TheTOBOuvrage;
      IndiceNomen := fTOBLigne.getvalue('GL_INDICENOMEN');
      if IndiceNomen > 0 then
      begin
        TOBP := fTOBOuvrage.detail[IndiceNomen-1];
        for Indice := 0 to TOBP.detail.count -1 do
        begin
        	if TOBP.detail[Indice].getValue('BLO_NATURETRAVAIL')<>'' then
          begin
          	result:= false;
            break;
          end;
        end;
      end;
    end;
  end;
end;

procedure TpieceCotraitOuv.AffecteInterne (Indice : integer);
var TOBL : TOB;
		Fournisseur : string;
begin
	Fournisseur := '';
  fTOBligne := TFLigNomen (FF).TOBLpiece;
  TOBL := fTOBOuvrage.detail[Indice-1]; if TOBL=nil then exit;
  TOBL.putvalue('BLO_NATURETRAVAIL','');
  TOBL.putvalue('BLO_FOURNISSEUR','');
  TOBL.putValue('GA_FOURNPRINC','');
  TOBL.putvalue('LIBELLEFOU','');
  TOBL.putvalue('BLO_DPA',TOBL.getValue('BLO_PUHT'));
  TOBL.putvalue('BLO_DPR',TOBL.getValue('BLO_PUHT'));
  TOBL.putvalue('BLO_NONAPPLICFC','X');
  TOBL.putvalue('BLO_NONAPPLICFG','-');
  TOBL.putvalue('BLO_NONAPPLICFRAIS','X');
  TOBL.putvalue('BLO_COEFFC',0.0);
  TOBL.putvalue('BLO_COEFFR',0.0);
  TOBL.putvalue('BLO_COEFFG',0.0);
  TOBL.putvalue('BLO_COEFMARG',1.0);
  fTOBLigne.putvalue('GL_RECALCULER','X');
  if NoMoreCotraitance(fTOBLigne) then
  begin
    fTOBLigne.putvalue('GLC_NATURETRAVAIL','');
    fTOBLigne.putvalue('GL_FOURNISSEUR','');
  end;
  fTOBPiece.putvalue('GP_RECALCULER','X');
  if IsOuvrage(TOBL) then
  begin
    AffecteInterneSousDetail (TOBL);
  end;
  TFligNomen(FF).AfficheLaligne(indice);
end;

procedure TpieceCotraitOuv.GSAffectInterne(Sender: tobject);
var Indice : integer;
begin
	if GS.nbSelected = 0 then BEGIN PGiInfo('Aucune ligne sélectionnée');exit; END;
  if PgiAsk ('Désirez-vous réellement affecter ces taches en interne ?','cotraitance') = mryes then
  begin
    for Indice := 1 to GS.RowCount do
    begin
      if (GS.IsSelected (Indice) ) then
      begin
        AffecteInterne (Indice);
      end;
    end;
  end;
  GS.OnBeforeFlip := nil;
  GS.OnFlipSelection  := nil;
  GS.AllSelected := false;
end;

procedure TpieceCotraitOuv.SetGrilleSaisie (State : boolean);
begin
	if SGN_NATURETRAVAIL <> -1 then
  begin
    GS.ColEditables[SGN_NATURETRAVAIL] := False;
    if not state then
    begin
      GS.Colwidths[SGN_NATURETRAVAIL] := -1;
    end else
    begin
      GS.Colwidths[SGN_NATURETRAVAIL] := 20;
      GS.ColFormats[SGN_NATURETRAVAIL] := 'CB=BTNATURETRAVAIL';
      GS.ColDrawingModes[SGN_NATURETRAVAIL]:= 'IMAGE'
    end;
	end;
  if SGN_LIBFOU <> -1 then
  begin
  	GS.ColEditables[SGN_LIBFOU] := False;
    if not state then
    begin
      GS.Colwidths[SGN_LIBFOU] := -1;
    end else
    begin
      GS.Colwidths[SGN_LIBFOU] := 90;
    end;
  end;

end;

procedure TpieceCotraitOuv.Setusable(const Value: boolean);
begin
  fusable := Value;
  SetGrilleSaisie(fusable);
  MenuEnabled(fusable);
end;

end.
