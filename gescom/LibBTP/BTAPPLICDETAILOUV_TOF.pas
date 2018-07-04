{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 29/11/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTAPPLICDETAILOUV ()
Mots clefs ... : TOF;BTAPPLICDETAILOUV
*****************************************************************}
Unit BTAPPLICDETAILOUV_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,MainEagl,
{$ENDIF}
		 AglInit,
     UTOB,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Menus,
     UTOF,
     SaisUtil,
     UtilPgi,
     UEntCommun ;

const MAXITEMS = 2;

Type

	TModeApplicationDOUV = (TMAONone,TMAOParag,TMAODocument);


	TApplicDetOuv = class
  private
  	IndiceDepart,IndiceFin : integer;
    fCreatedPop : boolean;
    FF : TForm;
    fusable : boolean;
    fgrid : THGrid;
    POPGS : TPopupMenu;
    TOBPiece , TOBOuvrages: TOB;
    TOBBases,TOBBasesL : TOB;
    TOBDetOuv : TOB;
    TOBL : TOB;
    MesMenuItem: array[0..MAXITEMS] of TMenuItem;
    fMaxItems : integer;
		procedure GSApplicSousDetail (Sender : TObject);
    procedure SetGrid(const Value: THGrid); virtual;
    procedure DefiniMenuPop(Parent: Tform);
    procedure Setdocument(const Value: TOB);
    procedure SetLigne(const Value: TOB);
    procedure MemoriseSousDetail (TOBL : TOB);
    procedure RechercheDepartFin (TOB : TOB; ModeApplication : TModeApplicationDOUV);
		procedure RemplaceSousDetail;
    procedure ReplDansSousDetailOuvrage (TOBLL : TOB);

    function RecupDebutPar (Ligne, Niveau : integer) : Integer;
    function RecupFinPar (ligne, Niveau : integer) : Integer;
    procedure ReplInsideSousDetailOuvrage(TOBLL, TOBOD: TOB);
  public
    constructor create (TT : TForm);
    destructor  destroy ; override;
    property    Grid : THGrid write setGrid ;
    property    document : TOB read TobPiece write Setdocument;
    property    Ouvrages : TOB read TobOuvrages write TOBOuvrages;
    property    PiedBAses : TOB read TOBBAses write TOBBases;
    property    LigneBases : TOB read TOBBasesL write TOBBasesL;
    property    Ligne : TOB read TOBL write SetLigne;
end;



  TOF_BTAPPLICDETAILOUV = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;


Implementation
uses facture,FactureBtp,FactUtil,FactOuvrage,FactCalc,NomenUtil,FactPiece;

function DefiniApplicDetailOuv (TOBL : TOB) : TModeApplicationDOUV;
var TOBP : TOB;
begin
	result := TMAONone;
	if TOBL.GetValue('GL_TYPEARTICLE') <> 'OUV' then exit;
  TOBP := TOB.Create('LES PARAMS',nil,-1);
  TOBP.AddChampSupValeur ('RESULT',-1);
  TOBP.AddChampSupValeur ('PARAM',TOBL.GEtVAlue('GL_NIVEAUIMBRIC'));
  TheTOB := TOBP;
  AGLLanceFiche ('BTP','BTAPPLICDETOUV','','','');
  if (TheTOB <> nil) and (TheTOB.GetValue('RESULT') <> -1) then
  begin
  	if TheTOB.GetValue('RESULT') = 0 then result := TMAOParag
    																 else result := TMAODocument;
  end;
  TheTOB := nil;
  TOBP.Free;
end;

procedure TOF_BTAPPLICDETAILOUV.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTAPPLICDETAILOUV.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTAPPLICDETAILOUV.OnUpdate ;
begin
  Inherited ;
  if TRadioButton (GEtControl('RPARAGRAPHE')).Checked then laTOB.putValue('RESULT',0)
  																										else laTOB.putValue('RESULT',1);
  TheTOB := LaTOB;
end ;

procedure TOF_BTAPPLICDETAILOUV.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTAPPLICDETAILOUV.OnArgument (S : String ) ;
begin
  Inherited ;
  TRadioButton (GEtControl('RPARAGRAPHE')).enabled := (LATOB.getValue('PARAM') <> 0);
  TRadioButton (GEtControl('RPARAGRAPHE')).Checked := (LATOB.getValue('PARAM') <> 0);
  TRadioButton (GEtControl('RDOCUMENT')).Checked := (LATOB.getValue('PARAM') = 0);
end ;

procedure TOF_BTAPPLICDETAILOUV.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTAPPLICDETAILOUV.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTAPPLICDETAILOUV.OnCancel () ;
begin
  Inherited ;
end ;

{ TApplicDetOuv }

constructor TApplicDetOuv.create(TT: TForm);
var ThePop : Tcomponent;
begin
	TOBDetOuv := TOB.Create ('DETAIL OUVRAGE SAUVE',nil,-1);

  FF := TT;
  fusable := true;
{$IFDEF BTP}
  if FF is TFFacture then ThePop := TFFacture(TT).Findcomponent  ('POPBTP');
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
end;

procedure TApplicDetOuv.DefiniMenuPop(Parent: Tform);
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
  //
  MesMenuItem[fmaxitems] := TmenuItem.Create (parent);
  with MesMenuItem[fmaxitems] do
  begin
    Caption := TraduireMemoire ('Appliquer ce &sous détail');  // par défaut
    Name := 'BAPPLICDETOUV';
    OnClick := GSApplicSousDetail;
    enabled := false;
  end;
  inc (fmaxitems);

  for Indice := 0 to fmaxitems -1 do
  begin
  	if MesMenuItem [Indice] <> nil then POPGS.Items.Add (MesMenuItem[Indice]);
  end;
end;

destructor TApplicDetOuv.destroy;
var indice : integer;
begin
	TOBDetOuv.free;

  for Indice := 0 to fmaxitems -1 do
  begin
    MesMenuItem[Indice].Free;
  end;
{$IFDEF BTP}
  if fcreatedPop then POPGS.free;
{$ENDIF}
  inherited;
end;

procedure TApplicDetOuv.GSApplicSousDetail(Sender: TObject);
var ModeApplication : TModeApplicationDOUV;
begin
	IndiceDepart := 0;
  IndiceFin := 0;
	ModeApplication := DefiniApplicDetailOuv (TOBL) ;
  if ModeApplication <> TMAONone then
  begin
  	MemoriseSousDetail (TOBL);
    RechercheDepartFin (TOBL,ModeApplication);
		RemplaceSousDetail;
  end;
end;


procedure TApplicDetOuv.MemoriseSousDetail(TOBL: TOB);
var TOBOUV,TOBO : TOB;
		indice : integer;
begin
	TOBDetOuv.ClearDetail; // nettoyage
  Indice := TOBL.getValue('GL_INDICENOMEN');
  TOBOUV := TOBOuvrages.detail[Indice-1];
  TOBO := TOB.Create ('UN OUVRAGE',TOBDETOUV,-1);
  TOBO.Dupliquer (TOBOUV,true,true);
end;

procedure TApplicDetOuv.RechercheDepartFin(TOB: TOB; ModeApplication: TModeApplicationDOUV);
begin
	if ModeApplication = TMAODocument then
  begin
  	IndiceDepart := 0;
    IndiceFin := TOBPiece.detail.count -1;
  end else
  begin
    IndiceDepart := RecupDebutPar (TOBL.GetValue('GL_NUMLIGNE')-1,TOBL.GetValue('GL_NIVEAUIMBRIC'));
    IndiceFin := RecupFinPar (TOBL.GetValue('GL_NUMLIGNE')-1,TOBL.GetValue('GL_NIVEAUIMBRIC'));
  end;
end;

function TApplicDetOuv.RecupDebutPar(Ligne, Niveau: integer): Integer;
var Indice : integer;
begin
	result := -1;
  for Indice := Ligne downto 0 do
  begin
  	if IsDebutParagraphe (TOBPiece.Detail[Indice],Niveau) then
    begin
    	result := indice;
      break;
    end;
  end;
end;

function TApplicDetOuv.RecupFinPar(Ligne, Niveau: integer): Integer;
var Indice : integer;
begin
	result := -1;
  for Indice := Ligne to TOBPiece.detail.count -1  do
  begin
  	if IsFinParagraphe (TOBPiece.Detail[Indice],Niveau) then
    begin
    	result := indice;
      break;
    end;
  end;
end;

procedure TApplicDetOuv.RemplaceSousDetail;
var Indice,IndiceNomen : integer;
		TOBLL,TOBOD : TOB;
    Valeurs : T_Valeurs;
    DEV : rdevise;
    bc,Cancel,first : boolean;
    SavCol,SavRow : integer;
begin
  first := true;
  DEV := TFFacture(ff).DEV;
  savcol := TFFacture(FF).GS.col;
  for Indice := IndiceDepart to IndiceFin do
  begin
    TOBLL := TOBPiece.Detail[Indice];
    if TOBLL = TOBL then continue; // pas la peine c'est la ligne de reference
    if (TOBLL.getValue('GL_ARTICLE')=TOBL.GetValue('GL_ARTICLE')) and (not IsSousDetail(TOBLL)) then
    begin
      indiceNomen := TOBLL.GetValue('GL_INDICENOMEN');
      if IndiceNomen = 0 then continue;
      TOBOD := TOBOuvrages.detail[IndiceNomen-1];
      TOBOD.ClearDetail;
      TOBOD.Dupliquer (TOBDetOuv.detail[0],true,true);
      TOBLL.SetDouble('GL_DPA',TOBL.GEtDouble('GL_DPA'));
      TOBLL.SetDouble('GL_DPR',TOBL.GEtDouble('GL_DPR'));
      TOBLL.SetDouble('GL_PUHT',TOBL.GEtDouble('GL_PUHT'));
      TOBLL.SetDouble('GL_PUTTC',TOBL.GEtDouble('GL_PUTTC'));
      TOBLL.SetDouble('GL_PUHTDEV',TOBL.GEtDouble('GL_PUHTDEV'));
      TOBLL.SetDouble('GL_PUTTCDEV',TOBL.GEtDouble('GL_PUTTCDEV'));
      TOBLL.SetDouble('GL_COEFFG',TOBL.GEtDouble('GL_COEFFG'));
      TOBLL.SetDouble('GL_COEFFC',TOBL.GEtDouble('GL_COEFFC'));
      TOBLL.SetDouble('GL_COEFFR',TOBL.GEtDouble('GL_COEFFR'));
      TOBLL.SetDouble('GL_COEFMARG',0);
      TOBLL.PutValue('GL_RECALCULER','X');
    end else
    begin
    	if IsOuvrage (TOBLL) then
      begin
      	ReplDansSousDetailOuvrage (TOBLL);
      end;
    end;
    if (first) and (TOBLL.getValue('GL_RECALCULER')='X') then
    begin
  		TOBPiece.PutValue('GP_RECALCULER','X');
      first := false;
    end;
  end;
//
	if TOBPiece.GetValue('GP_RECALCULER')='X' then
  begin
    {if TFFacture(FF).CalculPieceAutorise then }TFFacture(FF).CalculeLaSaisie (-1,-1,True);
    TFFacture(FF).SupLesLibDetail (TOBpiece);
    TFFacture(FF).AffichageDesDetailOuvrages;
    NumeroteLignesGC (nil,TOBPiece,true,false);
    TFFacture(FF).NettoieGrid;
    DefiniRowCount (FF,TOBPiece);
    for Indice := 0 to TOBPiece.Detail.Count - 1 do TFFacture(FF).AfficheLaLigne(Indice + 1);
    SavRow := TOBL.GetValue('GL_NUMLIGNE');

    bc:=False ; Cancel:=False ;
    TFFacture(FF).ActiveEventGrille (false);
    TFfacture(FF).GS.col := SavCol;
    TFfacture(FF).GS.row := Savrow;

    TFFacture(FF).PosValueCell (TFfacture(FF).GS.Cells[TFfacture(FF).GS.Col,TFfacture(FF).GS.Row]) ;
    TFFacture(FF).ActiveEventGrille (true);
  end;
end;

procedure TApplicDetOuv.ReplInsideSousDetailOuvrage (TOBLL,TOBOD : TOB);
var indice,Ind : integer;
		TOBOL,TOBOO,TOBOOO : TOB;
begin
  for Indice := 0 To TOBOD.Detail.count -1 do
  begin
    TOBOL := TOBOD.Detail [Indice];
    if TOBOL.getValue('BLO_ARTICLE')=TOBL.GEtValue('GL_ARTICLE') then
    begin
    	TOBOL.ClearDetail;
      for Ind := 0 to TOBDetOuv.detail[0].detail.count -1 do
      begin
        TOBOOO := TOBDetOuv.detail[0].detail[Ind];
        TOBOO := TOB.Create ('LIGNEOUV',TOBOL,-1);
        TOBOO.Dupliquer (TOBOOO,True,true);
        //
        TOBPiece.putValue('GP_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO')+1);
        TOBOO.putValue('BLO_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO'));
      end;
      TOBOL.SetDouble('BLO_DPA',TOBL.GEtDouble('GL_DPA'));
      TOBOL.SetDouble('BLO_DPR',TOBL.GEtDouble('GL_DPR'));
      TOBOL.SetDouble('BLO_PUHT',TOBL.GEtDouble('GL_PUHT'));
      TOBOL.SetDouble('BLO_PUTTC',TOBL.GEtDouble('GL_PUTTC'));
      TOBOL.SetDouble('BLO_PUHTDEV',TOBL.GEtDouble('GL_PUHTDEV'));
      TOBOL.SetDouble('BLO_PUTTCDEV',TOBL.GEtDouble('GL_PUTTCDEV'));
      TOBOL.SetDouble('BLO_COEFFG',TOBL.GEtDouble('GL_COEFFG'));
      TOBOL.SetDouble('BLO_COEFFC',TOBL.GEtDouble('GL_COEFFC'));
      TOBOL.SetDouble('BLO_COEFFR',TOBL.GEtDouble('GL_COEFFR'));
      TOBOL.SetDouble('BLO_COEFMARG',0);
      //
      TOBOL.PutValue('ANCPA',TOBOO.GetValue('BLO_DPA'));
      TOBOL.PutValue('ANCPR',TOBOO.GetValue('BLO_DPR'));
      GetValoDetail (TOBOL); // pour le cas des Article en prix posés
      TOBOL.PutValue('GA_PVHT',TOBOL.GetValue('BLO_PUHTDEV'));
      //
      TOBLL.PutValue('GL_RECALCULER','X');
    end else
    begin
    	if TOBOL.detail.Count > 0 then
      begin
      	ReplInsideSousDetailOuvrage (TOBLL,TOBOL);
      end;
    end;
  end;
end;

procedure TApplicDetOuv.ReplDansSousDetailOuvrage(TOBLL: TOB);
var IndiceOuv : integer;
		TOBOD : TOB;
begin
  IndiceOuv := TOBLL.GetValue('GL_INDICENOMEN');
  if IndiceOuv =0 then exit;
  TOBOD := TOBOuvrages.detail[IndiceOuv-1];
  ReplInsideSousDetailOuvrage (TOBLL,TOBOD);
  //
  if TOBLL.Getvalue('GL_RECALCULER')='X' then NumeroteLigneOuv (TOBOD,TOBLL,1,1,0,0,0);
  //
end;

procedure TApplicDetOuv.Setdocument(const Value: TOB);
var ThisTOb : TOB;
begin
  ThisTOB := Value;
  if ThisTOb = nil then BEGIN Fusable := False; Exit; END;
  TOBPiece := ThisTOB;
end;

procedure TApplicDetOuv.SetGrid(const Value: THGrid);
begin
   fgrid := Value;
{$IFDEF BTP}
  if fCreatedPop then if not assigned(fgrid.popupmenu) then fgrid.PopupMenu := POPGS;
{$ENDIF}
end;

procedure TApplicDetOuv.SetLigne(const Value: TOB);
var CurItem,TheItem : TmenuItem;
    indice : integer;
    TypeArticle : string;
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
      if CurItem.Name = 'BAPPLICDETOUV' then BEGIN TheItem := CurItem; Break; END;
    end;
  end;
  {$ENDIF}
  if TheItem = nil then exit;
  TypeArticle := TOBL.GetValue('GL_TYPEARTICLE');
  if (TypeArticle <> 'OUV') and (TOBL.GetValue('GL_INDICENOMEN')>0) then TheItem.enabled := false ELSE TheItem.enabled := true;
end;


Initialization
  registerclasses ( [ TOF_BTAPPLICDETAILOUV ] ) ;
end.

