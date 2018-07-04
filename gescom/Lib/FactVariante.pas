unit FactVariante;

interface
uses sysutils,classes,windows,messages,controls,forms,hmsgbox,stdCtrls,clipbrd,nomenUtil,
     HCtrls,SaisUtil,HEnt1,Ent1,EntGC,UtilPGI,UTOB,HTB97,FactUtil,FactComm,Menus,ParamSoc,
{$IFDEF EAGLCLIENT}
  maineagl,
{$ELSE}
  fe_main,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HRichOLE;

const MAXITEMS = 2;

type

	TTypeParagraphe = (Ttpdebut,TtpFin);

	TVariante = class
  private
    fCreatedPop : boolean;
    FF : TForm;
    fusable : boolean;
    fgrid : THGrid;
    POPGS : TPopupMenu;
    TOBPiece,TOBLigPiece : TOB;
    TobPieceTrait : TOB;
    TOBOuvrage : TOB;
    TOBOuvragesP : TOB;
    TOBAffaire : TOB;
    TOBtiers : TOB;
    TOBBases,TOBBasesL : TOB;
    fTOBSSTrait : TOB;
    TOBL : TOB;
    DEV : Rdevise;
    MesMenuItem: array[0..MAXITEMS] of TMenuItem;
    fMaxItems : integer;
    procedure SetGrid(const Value: THGrid); virtual;
    procedure GSSetVarianteorNot (Sender : Tobject);
    procedure DefiniMenuPop(Parent: Tform);
    procedure Setdocument(const Value: TOB);
    procedure SetLigne(const Value: TOB);
    procedure SetAffaire(const Value: TOB);
    procedure PositionneParagVariant (TOBParcours,TOBL : TOB; Ligne : integer; Variante : Boolean);
{$IFDEF BTP}
    procedure PositionneOuvrageVariant(TOBParcours,TOBL : TOB;Ligne : integer; Variante : Boolean);
{$ENDIF}
    procedure SetOuvrage(const Value: TOB);
    procedure SetBases(const Value: TOB);
    procedure SetBasesL(const Value: TOB);
  public
    constructor create (TT : TForm);
    destructor  destroy ; override;
    property    Grid : THGrid write setGrid ;
    property    document : TOB read TobPiece write Setdocument;
    property    PieceTrait : TOB read TOBPieceTrait write TOBPieceTrait;
    property    Ouvrage : TOB read TobOuvrage write SetOuvrage;
    property    Tiers : TOB read TobTiers write TOBTiers;
    property    OuvragesP : TOB read TobOuvragesP write TOBOuvragesP;
    property    affaire :TOb read TOBaffaire write Setaffaire;
    property    Bases : TOB read TOBBases write SetBases;
    property    BasesL : TOB read TOBBasesL write SetBasesL;
    property    Ligne : TOB read TOBL write SetLigne;
    property    Devise : Rdevise read DEV write DEV;
    property    TOBSSTRAIT : TOB read fTOBSSTrait write fTOBSSTrait;
end;

procedure SetTypeLigne (TOBL : TOB ; Variante : boolean);
function IsGerableEnVariante ( TOBL : TOB) : boolean;
function IsVariante ( TOBL : TOB) : boolean;
function IsArticleVariante ( TOBL : TOB) : boolean;
function IsParagrapheVariante ( TOBL : TOB) : boolean; overload;
function IsParagrapheVariante ( TOBL : TOB; Niveau : integer) : boolean; overload;
function IsDebutParagrapheVariante ( TOBL : TOB) : boolean; overload;
function IsDebutParagrapheVariante ( TOBL : TOB; Niveau : integer) : boolean; overload;
function IsFinParagrapheVariante ( TOBL : TOB) : boolean; overload;
function IsFinParagrapheVariante ( TOBL : TOB; Niveau :integer) : boolean; overload;
function IsLigneInVariante (TOBPiece,TOBL : TOB ; Ligne : integer) : boolean;
procedure SetLigneCommentaire (TobPiece,TOBL : TOB; Ligne : integer);
procedure SetParagraphe (TobPiece,TOBL : TOB; Ligne, niveau : integer; var variante : boolean; TypeParag :TTypeParagraphe=Ttpdebut );
procedure SupprimeLesVariantes (TOBOrigine : TOB);
function VariantePieceAutorisee (TOBPiece : TOB) : boolean;

implementation

uses facture,lignomen
{$IFDEF BTP}
     ,BTPUtil,FactOuvrage,FactureBtp,FactCalc
{$ENDIF}
;
function GetPrefixeTable (TOBL : TOB) : string;
begin
  result := '';
  if TOBL.nomtable = 'LIGNE' then result := 'GL' else
  if TOBL.nomtable = 'LIGNEOUVPLAT' then result := 'BOP' else
  if TOBL.nomtable = 'LIGNEOUV' then result := 'BLO';
end;

{***********A.G.L.***********************************************
Auteur  ...... : SANTUCCI
Créé le ...... : 18/07/2003
Modifié le ... :   /  /
Description .. : ISVariante -- permet de savoir si la ligne
Suite ........ : (LIGNE,LIGNEOUV) est une ligne de variante ou non
Suite ........ :
Suite ........ : renvoie true si variante  false sinon
Mots clefs ... : VARIANTE;LIGNE;LIGNEOUV
*****************************************************************}
function IsVariante ( TOBL : TOB) : boolean;
var prefixe : string;
    TypeLigne : string;
begin
  result:= false;
  if TOBL = nil then exit;
  Prefixe := GetPrefixeTable (TOBL);
  if Prefixe = '' then exit;
  TypeLigne := TOBL.GetValue(Prefixe+'_TYPELIGNE');
  if (TypeLigne = 'ARV') or
     (TypeLigne = 'COV') or
     (TypeLigne = 'SDV') or
     (copy (TypeLigne,1,2) = 'DV') or
     (copy (TypeLigne,1,2) = 'TV')  then result := true;
end;

function IsArticleVariante ( TOBL : TOB) : boolean;
begin
	result := (IsArticle(TOBL) and IsVariante(TOBL));
end;

function IsParagrapheVariante ( TOBL : TOB) : boolean;
begin
	result := (IsParagraphe(TOBL) and IsVariante(TOBL));
end;

function IsParagrapheVariante ( TOBL : TOB; Niveau : integer) : boolean;
begin
	result := (IsParagraphe(TOBL,Niveau) and IsVariante(TOBL));
end;

function IsDebutParagrapheVariante ( TOBL : TOB) : boolean;
begin
	result := (IsDebutParagraphe(TOBL) and IsVariante(TOBL));
end;

function IsDebutParagrapheVariante ( TOBL : TOB; Niveau : integer) : boolean; 
begin
	result := (IsDebutParagraphe(TOBL,Niveau) and IsVariante(TOBL));
end;

function IsFinParagrapheVariante ( TOBL : TOB) : boolean;
begin
	result := (IsFinParagraphe(TOBL) and IsVariante(TOBL));
end;

function IsFinParagrapheVariante ( TOBL : TOB; Niveau :integer) : boolean; overload;
begin
	result := (IsFinParagraphe(TOBL,Niveau) and IsVariante(TOBL));
end;

{***********A.G.L.***********************************************
Auteur  ...... : SANTUCCI
Créé le ...... : 18/07/2003
Modifié le ... :   /  /
Description .. : ISGerableEnVariante -- permet de savoir si la ligne
Suite ........ : (LIGNE,LIGNEOUV) peut etre géré en variante ou non
Suite ........ :
Suite ........ : renvoie true si gérable en variante  false sinon
Mots clefs ... : VARIANTE;LIGNE;LIGNEOUV
*****************************************************************}
function IsGerableEnVariante ( TOBL : TOB) : boolean;
var prefixe : string;
    TypeLigne : string;
begin
{$IFDEF BTP}
  result:= false;
  Prefixe := GetPrefixeTable (TOBL);
  if prefixe = '' then exit;
  TypeLigne := TOBL.GetValue(prefixe+'_TYPELIGNE');
  if (TypeLigne = 'ARV') or (copy (TypeLigne,1,2) = 'DV') or
     (TypeLigne = 'ART') or (copy (TypeLigne,1,2) = 'DP') then result := true;
{$ELSE}
  result := false;
{$ENDIF}
end;


{***********A.G.L.***********************************************
Auteur  ...... : SANTUCCI
Créé le ...... : 18/07/2003
Modifié le ... :   /  /
Description .. : SetTypeLigne -- permet de définir le type de ligne sur
Suite ........ : (LIGNE,LIGNEOUV) en fonction du type déjà présent sur la ligne
Suite ........ :
Suite ........ :
Mots clefs ... : VARIANTE;LIGNE;LIGNEOUV
*****************************************************************}
procedure SetTypeLigne (TOBL : TOB ; Variante : boolean);
var Prefixe : string;
    ChampTypeLigne,TypeLigne : string;
begin
  Prefixe := GetPrefixeTable (TOBL);
  if Prefixe = '' then exit;
  ChampTypeLigne := Prefixe+'_TYPELIGNE';
  TypeLigne := TOBL.GetValue(ChampTypeLigne);
  if (TypeLigne='ART') or (TypeLigne='ARV') then
  BEGIN
    if Variante then
    begin
      TOBL.PutValue(ChampTypeLigne,'ARV');
    end else
    begin
      TOBL.PutValue(ChampTypeLigne,'ART');
    end;
  END ELSE
  if (TypeLigne='COM') or (TypeLigne='COV') then
  BEGIN
    if Variante then TOBL.PutValue(ChampTypeLigne,'COV')
                else TOBL.PutValue(ChampTypeLigne,'COM');
  END ELSE

  if (copy(TypeLigne,1,2)='DP') or (copy(TypeLigne,1,2)='DV') then
  BEGIN
    if Variante then TOBL.PutValue(ChampTypeLigne,'DV'+copy(TypeLigne,3,1))
                else TOBL.PutValue(ChampTypeLigne,'DP'+copy(TypeLigne,3,1));
  END ELSE
  if (copy(TypeLigne,1,2)='TP') or (copy(TypeLigne,1,2)='TV')  then
  BEGIN
  if Variante then TOBL.PutValue(ChampTypeLigne,'TV'+copy(TypeLigne,3,1))
              else TOBL.PutValue(ChampTypeLigne,'TP'+copy(TypeLigne,3,1));
  END;
end;

function IsLigneInVariante (TOBPiece,TOBL : TOB ; Ligne : integer) : boolean;
var IndicePar : integer;
    Niveau : integer;
begin
  result := false;
  if Tobpiece = nil then exit;
  Niveau := TOBL.GetValue('GL_NIVEAUIMBRIC');
  if isDebutParagraphe (TOBL) then Niveau := Niveau -1;
  if Niveau = 0 then exit;
  if Ligne < 2 then exit;
  IndicePar := RecupDebutParagraph (TOBpiece,Ligne-1,Niveau,true);
  if IsVariante (TOBPiece.detail[IndicePar]) then result := true;
end;


procedure SetLigneCommentaire (TobPiece,TOBL : TOB; Ligne : integer);
var IndicePar : integer;
begin
  TOBL.PutValue('GL_TYPELIGNE','COM'); // par défaut
  if TOBL.GetValue('GL_NIVEAUIMBRIC') = 0 then exit;
  if TOBPiece = nil then exit;
  if Ligne = 0 then exit;
  IndicePar := RecupDebutParagraph (TOBpiece,Ligne,TOBL.GetValue('GL_NIVEAUIMBRIC'),true);
  if IsVariante (TOBPiece.detail[IndicePar]) then TOBL.PutValue('GL_TYPELIGNE','COV');
end;


procedure SetParagraphe (TobPiece,TOBL : TOB; Ligne, niveau : integer; var variante : boolean; TypeParag : TTypeParagraphe );
var IndicePar : integer;
begin
	if TypeParag = Ttpdebut then
  begin
  	TOBL.PutValue('GL_TYPELIGNE','DP'+IntToStr(niveau)); // par défaut
  end else
  begin
  	TOBL.PutValue('GL_TYPELIGNE','TP'+IntToStr(niveau)); // par défaut
  end;
//  variante := false;
  if niveau = 1 then
  begin
  	TOBL.PutValue('GL_DATELIVRAISON',TOBPiece.GEtValue('GP_DATELIVRAISON'));
  	if TOBL.GetValue('GL_DATELIVRAISON') = iDate1900 then TOBL.PutValue('GL_DATELIVRAISON',TOBPiece.getValue('GP_DATEPIECE'));
  	TOBL.PutValue('GL_NIVEAUIMBRIC',niveau);
  	exit;
  end;
  if TOBPiece = nil then exit;
  if Ligne = 0 then exit;
  IndicePar := RecupDebutParagraph (TOBpiece,Ligne,niveau -1,true);
  if IndicePar >= 0 then TOBL.PutValue('GL_DATELIVRAISON',TOBPiece.detail[IndicePar].GEtValue('GL_DATELIVRAISON'))
  									else TOBL.PutValue('GL_DATELIVRAISON',TOBPiece.GEtValue('GP_DATELIVRAISON'));
  if TOBL.GetValue('GL_DATELIVRAISON') = iDate1900 then TOBL.PutValue('GL_DATELIVRAISON',TOBPiece.getValue('GP_DATEPIECE'));
  TOBL.PutValue('GL_NIVEAUIMBRIC',niveau);
  if IsVariante (TOBPiece.detail[IndicePar]) then
  begin
		if TypeParag = Ttpdebut then
    begin
    	TOBL.PutValue('GL_TYPELIGNE','DV'+IntToStr(niveau));
    end else
    begin
    	TOBL.PutValue('GL_TYPELIGNE','TV'+IntToStr(niveau));
    end;
    variante := true;
  end;
end;

procedure SupprimeLesVariantes (TOBOrigine : TOB);
var Indice : integer;
    TOBL : TOB;
begin
  for Indice := TOBOrigine.detail.count -1 downto 0 do
  begin
     TOBL := TOBOrigine.detail[Indice];
     if IsVariante (TOBL) then TOBL.free;
  end;
end;

{ TVariante }

constructor TVariante.create(TT: TForm);
var ThePop : Tcomponent;
begin
  FF := TT;
  fusable := true;
{$IFDEF BTP}
  if FF is TFFacture then ThePop := TFFacture(TT).Findcomponent  ('POPBTP')
                     else ThePop := TFLigNomen(TT).Findcomponent  ('POPUPG_NLIG');
  if ThePop = nil then
  BEGIN
    // pas de menu BTP trouve ..on le cree
    POPGS := TPopupMenu.Create(TT);
    POPGS.Name := 'POPBTP';
    fCreatedPop := true;
  END else
  BEGIN
    fCreatedPop := false;
    POPGS := TPopupMenu(thePop);
  END;
  DefiniMenuPop(TT);
{$ENDIF}
  if TT is TFFActure then  grid := TFFacture(TT).GS;
  if TT is TFLigNomen then grid := TFLigNomen(TT).G_NLIG;
end;

procedure TVariante.DefiniMenuPop(Parent: Tform);
var Indice : integer;
begin
  fmaxitems := 0;
  if not fcreatedPop then
  begin
    MesMenuItem[fmaxitems] := TmenuItem.Create (parent);
    with MesMenuItem[fmaxitems] do
      begin
      Caption := '-';
      end;
    inc (fmaxitems);
  end;
  // CTRL+ALT+A = Définir comme variante
  MesMenuItem[fmaxitems] := TmenuItem.Create (parent);
  with MesMenuItem[fmaxitems] do
    begin
    Caption := TraduireMemoire ('Définir comme variante');  // par défaut
    Name := 'BVARIANTE';
    OnClick := GSSetVarianteorNot;
    enabled := false;
    end;
  MesMenuItem[fmaxitems].ShortCut := ShortCut( Word('A'), [ssCtrl,ssAlt]);
  inc (fmaxitems);

  for Indice := 0 to fmaxitems -1 do
    begin
      if MesMenuItem [Indice] <> nil then POPGS.Items.Add (MesMenuItem[Indice]);
    end;
end;

destructor TVariante.destroy;
var indice : integer;
begin
  for Indice := 0 to fmaxitems -1 do
  begin
    MesMenuItem[Indice].Free;
  end;
{$IFDEF BTP}
  if fcreatedPop then POPGS.free;
{$ENDIF}
  inherited;
end;

procedure TVariante.GSSetVarianteorNot(Sender: Tobject);
var
    prefixe : string;
    variante : boolean;
begin
  Prefixe := GetPrefixeTable (TOBL);
  if prefixe = '' then exit;
  Variante := not IsVariante (TOBL);
  if Prefixe = 'GL' then
	begin
    if IsFinParagraphe (TOBL) then exit;
    if IsDebutParagraphe(TOBL) then
    begin
      PositionneParagVariant  (TOBPiece,TOBL,fgrid.row,Variante);
    end else
    if IsOuvrage(TOBL) then
    begin
      PositionneOuvrageVariant (TOBPiece,TOBL,fgrid.row,Variante);
    end else
    begin
      if Variante <> Isvariante(TOBL) then
      begin
        if IsVariante(TOBL) then
        begin
          ZeroLigneMontant (TOBL);
        end else
        begin
          AssigneDocumentTva2014(TOBPiece,TOBSSTRAIT);
          DeduitLigneModifie (TOBL,TOBpiece,TOBPieceTrait,TOBOuvrage, TOBouvragesP,TOBBases,TOBBasesL,TOBTiers,DEV, tamodif);
          InitDocumentTva2014;
        end;
      end;
      TOBL.putValue('GL_RECALCULER','X');
    end;
  end;
  SetTypeLigne (TOBL,Variante);
  fgrid.InvalidateRow (fgrid.row);
  if FF is TFFacture then
  begin
    TOBPiece.putValue('GP_RECALCULER','X');
    if (TFFacture(FF).BCalculDocAuto.down) or (TOBL.GetValue ('GL_TYPEPRESENT') <> 0) then
    begin
      TFFacture(FF).CalculeLaSaisie (-1,-1,True);
    end else
    begin
      TFFacture(FF).CalculeLaSaisie (-1,fgrid.row,True);
//      TFFacture(FF).CacheTotalisations (true);
      TFFacture(FF).GS.Refresh;
      //FV1 : 16/11/20015 - pas d'indice hors limite mais une violation d'accès
      TFFacture(FF).GS.ClearSelected;

    end;
{$IFDEF BTP}
  end
  else if FF is TFLigNomen then
  begin
   TFLigNomen(FF).CalculeOuvrageLoc;
   TFLigNomen(FF).AfficheTotalisation;
   TFLigNomen(FF).AfficheMontantLigne (fGrid.row);
   //FV1 : 16/11/20015 - pas d'indice hors limite mais une violation d'accès
   TFLigNomen(FF).G_NLIG.ClearSelected;
   TFLigNomen(FF).G_NLIG.Refresh;
{$ENDIF}
  end;
  
  SetLigne (TOBL);
end;

function VariantePieceAutorisee (TOBPiece : TOB) : boolean;
var Naturepiece : string;
begin
  result := true;
  NaturePiece := TOBPiece.GetValue('GP_NATUREPIECEG');
{$IFDEF BTP}
  if (NaturePiece <> GetParamSoc('SO_AFNATPROPOSITION')) and
     (NaturePiece <> GetParamSoc('SO_AFNATAFFAIRE')) then result := false;
  if result then
  begin
    if GetEtatAffaire (TOBpiece.getValue('GP_AFFAIREDEVIS'))<>'ENC' then
      result := false
    else
      result := true;
  end;
{$ELSE}
{$ENDIF}
end;

{$IFDEF BTP}
procedure TVariante.PositionneOuvrageVariant(TOBParcours,TOBL : TOB;Ligne : integer; Variante : boolean);
var indicenomen , ITypeLig,Indrech,indice : integer;
    TobDet : TOB;
    prefixe : string;
begin
  Prefixe := GetPrefixeTable (TOBL);
  if prefixe = '' then exit;
  if Prefixe = 'GL' then
  begin
    IndiceNomen := TOBL.GetNumChamp('GL_INDICENOMEN');
    ITypeLig := TOBL.GetNumchamp('GL_TYPELIGNE');
    IndRech := TOBL.GetValeur(IndiceNomen);
    if Variante <> Isvariante(TOBL) then
    begin
      if IsVariante(TOBL) then
      begin
        ZeroLigneMontant (TOBL);
      end else
      begin
        DeduitLigneModifie (TOBL,TOBpiece,TOBPieceTrait,TOBOuvrage,TOBOuvragesP,TOBBases,TOBBasesL,TOBTiers,DEV, taModif);
      end;
    end;
    SetTypeLigne (TOBL,Variante);
    TOBL.PutValue ('GL_RECALCULER','X');
    fgrid.InvalidateRow (TOBL.GetValue('GL_NUMLIGNE'));
    if TOBL.GetValue ('GL_TYPEPRESENT') = 0 then exit;
    GestionDetailOuvrage (FF,TOBParcours,fgrid.row);
  end else
  begin
    SetTypeLigne (TOBL,Variante);
  end;
end;
{$ENDIF}

procedure TVariante.PositionneParagVariant (TobParcours,TOBL : TOB; Ligne : integer; Variante : boolean);
var Arow : integer;
    Indice : integer;
    ParagStop : string;
    TheLigne : TOB;
    Niveau : integer;
begin
  niveau := strtoint(copy (TOBL.GetValue('GL_TYPELIGNE'),3,1));
  for Indice := Ligne to TobParcours.detail.count do
  begin
    TheLigne := TOBParcours.detail[Indice-1];
    if IsFinParagraphe(TheLigne,niveau) then
    begin
      SetTypeLigne (TheLigne , Variante);
      fgrid.InvalidateRow (Indice);
      break;
    end else if IsOuvrage (TheLigne) then
    begin
      PositionneOuvrageVariant (TobParcours,TheLigne,Indice,variante);
    end else
    begin
      if not IsDebutparagraphe (TheLigne) then
      begin
        if Variante <> Isvariante(TheLigne) then
        begin
          if IsVariante(TheLigne) then
          begin
            ZeroLigneMontant (TheLigne);
          end else
          begin
            AssigneDocumentTva2014(TOBPiece,TOBSSTRAIT);
            DeduitLigneModifie (TheLigne,TOBpiece,TOBPieceTrait,TOBOuvrage,TOBouvragesP,TOBBases,TOBBasesL,TOBTiers,DEV, TaModif);
            InitDocumentTva2014;
          end;
        end;
        TheLigne.putValue('GL_RECALCULER','X');
      end;
      SetTypeLigne (TheLigne , Variante);
      fgrid.InvalidateRow (Indice);
    end;
  end;
  SetTypeLigne (TOBL , Variante);
end;

procedure TVariante.SetAffaire(const Value: TOB);
begin
  TobAffaire := value;
end;

procedure TVariante.SetBases(const Value: TOB);
begin
  TOBBases := value;
end;

procedure TVariante.SetBasesL(const Value: TOB);
begin
  TOBBasesL := value;
end;

procedure TVariante.Setdocument(const Value: TOB);
var ThisTOb : TOB;
begin
  ThisTOB := Value;

  if ThisTOb = nil then BEGIN Fusable := False; Exit; END;

  if ThisTOB.NomTable = 'LIGNE' then
  begin
    TOBPiece := ThisTOB.Parent;
    TOBligPiece := ThisTob;
  end else TOBPiece := ThisTOB;
  if not VariantePieceAutorisee (TOBPiece)  then BEGIN fusable := false; Exit; END;
end;

procedure TVariante.SetGrid(const Value: THGrid);
begin
   fgrid := Value;
{$IFDEF BTP}
  if fCreatedPop then if not assigned(fgrid.popupmenu) then fgrid.PopupMenu := POPGS;
{$ENDIF}

end;

procedure TVariante.SetLigne(const Value: TOB);
var CurItem,TheItem : TmenuItem;
    indice : integer;
    TypeLigne : string;
    perfixe : string;
begin
  if not fusable then exit;
  TOBL := value;
  if TOBL = nil then exit;
  TheItem := nil;
  {$IFDEF BTP}
  if POPGS <> nil then
  begin
    for indice := 0 to POPGS.Items.Count do
    begin
      CurItem := POPGS.Items [indice];
      if CurItem.Name = 'BVARIANTE' then BEGIN TheItem := CurItem; Break; END;
    end;
  end;
  {$ENDIF}
  if TheItem = nil then exit;
  if TOBL.NomTable = 'LIGNE' then
  begin
    TypeLigne := TOBL.GetValue('GL_TYPELIGNE');
    if IsLigneInVariante (TOBPiece,TOBL,Fgrid.row) and isVariante(TOBL) then BEGIN TheItem.enabled := false; exit; END;
    if not isGerableEnVariante (TOBL) or not fusable then TheItem.Enabled := false
                                                     else TheItem.Enabled := true;
    if not fusable then exit;
    if isVariante (TOBL) then TheItem.caption := TraduireMemoire ('Prendre en compte dans le calcul')
                         else TheItem.caption := TraduireMemoire ('Définir comme variante');
  end else
  begin
    TypeLigne := TOBL.GetValue('BLO_TYPELIGNE');
    // provenance d'un article en prix pose
    if TFLigNomen(FF).TypeArp Then BEGIN TheItem.Enabled := false;Exit; END;
    if (not isGerableEnVariante (TOBL)) or (not fusable) or (isVariante(TobLigPiece)) then TheItem.Enabled := false
                                                                                      else TheItem.Enabled := true;
    if not fusable then exit;
    if isVariante (TOBL) then TheItem.caption := TraduireMemoire ('Prendre en compte dans le calcul')
                         else TheItem.caption := TraduireMemoire ('Définir comme variante');
  end;
end;

procedure TVariante.SetOuvrage(const Value: TOB);
begin
  TobOuvrage := value;
end;

end.
