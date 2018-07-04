{***********UNITE*************************************************
Auteur  ...... : SANTUCCI Lionel
Créé le ...... : 12/01/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSAISDOCEXT ()
Mots clefs ... : TOF;BTSAISDOCEXT
*****************************************************************}
Unit UTOF_BTSAISDOCEXT ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     HTB97,
     AglInit,
     EntGc,
     SaisUtil,
     fiche,
{$IFNDEF EAGLCLIENT}
     Fe_Main,
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     maineagl
     eMul,
{$ENDIF}
     Graphics,
     Dialogs,
     grids,
     Hpanel,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOB,
     FactTOB,
     FactUtil,
     FactVariante,
     utofAfBaseCodeAffaire,
     FactTiers,FactComm,StockUtil,
     UTOF ;

Type

  T_GestionDocExt = (GdeVisuFRF, GdeSelectRF, GdeSelect);
  // ------------------------------
  // Pour la visu document en franc
  // ------------------------------
  // Entht,TOBPiece,TOBOuvrage : a renseigner
  //
  // ------------------------------
  // Pour la sélection des receptions fournisseur
  // ------------------------------
  // TOBTiers,TOBAffaire,TOBRecep : a renseigner
  //
  // ------------------------------


  T_VisuDoc = class
  private
  	fTOBTiers : TOB;
  	fTOBAffaire : TOB;
    fTOBPiece : TOB;
    fTobOuvrage : TOB;
    fTOBArticles : TOB;
    fModeGestion : T_GestionDocExt;
    fEnHt : boolean;
    FTOBDetRecep : TOB;
    fTOBresult : TOB;
    procedure SetTOBPiece ( ThePiece : TOB);
    procedure ConvertitPieceToFrancs(TheDocument: TOB);
    procedure PrepareAffichageRecep ( TheDocument : TOB);
    procedure TrouveRefTiers(CodeArt, NatTiers, CodeTiers: string; TOBARt : TOB);
    procedure InsereFille (TOBL,TOBDR : TOB);
    procedure AffecteDispo(TOBA, TOBL: TOB);
  public
    property TOBAffaire : TOB read fTOBAffaire write fTOBAffaire;
    property TOBPiece : TOB read fTOBPiece write SetTOBPiece;
    property TOBOuvrage : TOB read fTOBOuvrage write fTOBOuvrage;
    property TOBRecep : TOB read FTOBDetRecep write FTOBDetRecep;
    property TOBTiers : TOB read fTOBTiers write fTOBTiers;
    Property TOBresult : TOB read fTOBresult;
    property TOBArticles : TOB read fTOBArticles;
    property Modegestion : T_GestionDocExt read fModegestion write fModegestion;
    property EnHt : boolean read fEnHt;
    constructor create;
    destructor destroy; override;
    procedure VoirDocument;
  end;

  TOF_BTSAISDOCEXT = Class (TOF_AFBASECODEAFFAIRE)
  private
  	fNbCols : integer;
    lesColonnes : string;
    TOBPiece : TOB;
    TypeGestion : T_GestionDocExt;
    GS : THGrid;
    EnHt : boolean;
    PENTETE,PPIED : THPanel;
    TitrePiece : THLabel;
    REFERENCE : THedit;
    LCLIENT,LDEVISE,LDEVISE1 : THLabel;
    TDEVISE : THEDIT;
    BOUVRIR,Bvalider,BRechercher,BselectAll : TToolBarButton97;
    SG_NUML,SG_TYPEA,SG_REFA,SG_LIB,SG_QTE,SG_UV,SG_PUV,SG_MONTANT,SG_SELECT,SG_REFLIVRFOU : integer;
    FirstFind: boolean;
    FindDialog: TFindDialog;
    procedure GetComposants;
    procedure DefiniGrille;
    procedure BeforeShow;
    procedure AfficheLaLigne(ARow: integer);
    procedure DefiniLesColonnes;
    procedure GetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas;
      AState: TGridDrawState);
    procedure ReajusteAffichage(Arow: integer);
    procedure BouvrirClick (Sender : Tobject);
    procedure BRechercherClick(Sender: TObject);
    procedure FindDialogFind(Sender: TObject);
    procedure BselectAllClick(Sender: Tobject);
  public
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
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

procedure T_VisuDoc.ConvertitPieceToFrancs (TheDocument : TOB);
var Indice : integer;
    TOBL : TOB;
begin
  TheDocument.AddChampSupValeur ('TOTALPIECE',0,false);
  if EnHt then
    TheDocument.putvalue('TOTALPIECE',arrondi(TheDocument.Getvalue('GP_TOTALHTDEV')*6.55957,V_PGI.OkDecV))
  else
    TheDocument.putvalue('TOTALPIECE',arrondi(TheDocument.Getvalue('GP_TOTALTTCDEV')*6.55957,V_PGI.OkDecV));

  for Indice := 0 to TheDocument.detail.count -1 do
  begin
    TOBL := TheDocument.detail[Indice];
    if EnHt then
    begin
      TOBL.PutValue('GL_PUHTDEV',arrondi(TOBL.GetValue('GL_PUHTDEV')*6.55957,V_PGI.OkDecP));
      TOBL.PutValue('GL_MONTANTHTDEV',arrondi(TOBL.GetValue('GL_MONTANTHTDEV')*6.55957,V_PGI.OkDecV));
    end else
    begin
      TOBL.PutValue('GL_PUTTCDEV',arrondi(TOBL.GetValue('GL_PUTTCDEV')*6.55957,V_PGI.OkDecP));
      TOBL.PutValue('GL_MONTANTTTCDEV',arrondi(TOBL.GetValue('GL_MONTANTTTCDEV')*6.55957,V_PGI.OkDecV));
    end;
  end;
end;

constructor T_VisuDoc.create;
begin
  fEnHt := true; // par défaut
  fTOBPiece := nil;
  fTOBArticles := TOB.Create ('LES ARTICLES',nil,-1);
end;

destructor T_VisuDoc.destroy;
begin
  inherited;
  fTOBArticles.free;
  if fTOBresult <> nil then fTOBResult.free;
end;

procedure T_VisuDoc.AffecteDispo (TOBA,TOBL : TOB);
var TOBD : TOB;
		req : string;
    QQ : Tquery;
begin
  Req := 'SELECT * FROM DISPO WHERE GQ_ARTICLE="'+TOBL.GetValue('GL_ARTICLE')+'" AND GQ_DEPOT="'+TOBL.GEtValue('GL_DEPOT')+'"';
  QQ := OpenSql (Req,True);
  TRY
    if not QQ.eof then
    begin
      TOBD := TOB.Create ('DISPO',TOBA,-1);
      TOBD.SelectDB ('',QQ,true);
    end;
  FINALLY
  	ferme (QQ);
  END;
end;

procedure T_VisuDoc.PrepareAffichageRecep(TheDocument: TOB);
var TOBR,TOBL,TOBA : TOB;
		Indice,Numl : integer;
    Article,RefSais,Refpiece : String;
    Qte,Value : double;
begin
	fEnHt :=  (TOBTiers.GetValue('T_FACTUREHT')='X');
  TheDocument.putValue('GP_NATUREPIECEG','BLF');
  InitTOBPiece (TheDocument);
  TheDocument.putValue('GP_TIERS',TOBTiers.getValue('T_TIERS'));
  TheDocument.putValue('GP_DEVISE',TOBTiers.getValue('T_DEVISE'));
  TheDocument.putValue('GP_AFFAIRE',TOBAffaire.GetValue('AFF_AFFAIRE'));
  TheDocument.putValue('GP_AFFAIRE0',TOBAffaire.GetValue('AFF_AFFAIRE0'));
  TheDocument.putValue('GP_AFFAIRE1',TOBAffaire.GetValue('AFF_AFFAIRE1'));
  TheDocument.putValue('GP_AFFAIRE2',TOBAffaire.GetValue('AFF_AFFAIRE2'));
  TheDocument.putValue('GP_AFFAIRE3',TOBAffaire.GetValue('AFF_AFFAIRE3'));
  TheDocument.putValue('GP_AVENANT',TOBAffaire.GetValue('AFF_AVENANT'));
  TiersVersPiece (TOBTiers,TheDocument);
  AffaireVersPiece (TheDocument,TOBAffaire);
  //
  for Indice := 0 to FTOBDetRecep.detail.count -1 do
  begin
    TOBR := FTOBDetRecep.detail[Indice];
    if TOBR.GetValue('_REFLIVRFOU_') = '' then
    begin
    	Qte := TOBR.GetValue('BCO_QTERETOUR') * (-1);
    end else
    begin
    	Qte := TOBR.GetValue('BCO_QUANTITE') - TOBR.GetValue('BCO_QTEVENTE');
    end;
    If Qte = 0 Then Continue;
    Value := TOBR.GetValue('MTACHAT');
    Article := TOBR.GetValue('GL_ARTICLE');
    RefSais := Trim(Copy(TOBR.GetValue('GL_ARTICLE'), 1, 18));
    RefPiece := EncodeRefPiece (TOBR);
    TOBA := fTOBArticles.findFirst(['GA_ARTICLE'],[Article],true);
    if TOBA = nil then
    begin
    	TOBA := TOB.Create ('ARTICLE',fTOBARticles,-1);
      TOBA.PutValue('GA_ARTICLE',Article);
      TOBA.LoadDB;
      TOBA.AddChampSupValeur('REFARTBARRE', '', False);
      TOBA.AddChampSupValeur('REFARTTIERS', '', False);
      TOBA.AddChampSupValeur('REFARTSAISIE', '', False);
      TOBA.AddChampSupValeur('UTILISE', '-', False);
			TrouveRefTiers (refSais,TOBTiers.GetValue('T_NATUREAUXI'),TOBTIers.GetValue('T_TIERS'),TOBA);
    end;
    TOBL := TheDocument.FindFirst (['GL_ARTICLE','GL_PIECEORIGINE'],[Article,RefPiece],true);
    if TOBL = nil then
    begin
      TOBL := NewTOBLigne(TheDocument, 0);
      TOBL.ClearDetail;
      NumL := TheDocument.Detail.Count;
      InitialiseTOBLigne(TheDocument, TOBTiers, TOBAffaire, NumL, false);
    	ArticleVersLigne (TheDocument,TOBA,nil,TOBL,TOBTiers,true);
//
			TOBA.PutValue('REFARTSAISIE', RefSais);
//
      TOBL.PutValue('GL_ARTICLE', TOBA.GetValue('GA_ARTICLE'));
      TOBL.PutValue('GL_PIECEORIGINE', RefPiece);
      TOBL.PutValue('_LIBREFLIVRFOU_', TOBR.GetValue('_LIBREFLIVRFOU_'));
      if TOBA.GetValue('REFARTTIERS') <> '' then TOBL.PutValue('GL_REFARTSAISIE', TOBA.GetValue('REFARTTIERS'))
      																			 else TOBL.PutValue('GL_REFARTSAISIE', TOBA.GetValue('REFARTSAISIE'));
      TOBL.PutValue('GL_CODEARTICLE', TOBA.GetValue('GA_CODEARTICLE'));
      TOBL.PutValue('GL_REFARTBARRE', TOBA.GetValue('REFARTBARRE'));
      if TOBA.GetValue('REFARTTIERS') <> '' then TOBL.PutValue('GL_REFARTTIERS', TOBA.GetValue('REFARTTIERS'))
      																			 else TOBL.PutValue('GL_REFARTTIERS', TOBA.GetValue('REFARTTIERS'));
      TOBL.PutValue('GL_TYPEREF', 'ART');
    end;
    TOBL.PutValue('GL_QUALIFQTEACH',TOBR.GetValue('GL_QUALIFQTEACH'));
//    QTe := Qte * GetRatio (TOBL,nil,trsVente);
    TOBL.putValue('GL_LIBELLE',TOBR.GetValue('GL_LIBELLE'));
    TOBL.putValue('GL_QTEFACT',TOBL.GetValue('GL_QTEFACT') + Qte);
    TOBL.PutValue('GL_QTESTOCK',TOBL.GEtValue('GL_QTEFACT'));
    TOBL.PutValue('GL_QTERESTE',TOBL.GEtValue('GL_QTEFACT'));
    TOBL.PutValue('GL_QTERELIQUAT',0);
    TOBL.PutValue('GL_MTRELIQUAT', 0);
    TOBL.PutValue('GL_DEPOT',TOBR.GetValue('GL_DEPOT'));
    if Qte > 0 then
    begin
    	TOBL.PutValue('GL_MONTANTHTDEV',TOBL.GetValue('GL_MONTANTHTDEV')+ Value);
      TOBL.PutValue('GL_MTRESTE',     TOBL.GEtValue('GL_MONTANTHTDEV'));
      if TOBL.GetValue('GL_QTEFACT') <> 0 then
      begin
        TOBL.PutValue('GL_DPA',Arrondi(TOBL.GetValue('GL_MONTANTHTDEV')/TOBL.GetValue('GL_QTEFACT'),V_PGI.okdecP));
      end else
      begin
        TOBL.PutValue('GL_DPA',Arrondi(TOBL.GetValue('GL_MONTANTHTDEV'),V_PGI.okdecP));
      end;
    end;
    AffecteDispo (TOBA,TOBL);
    InsereFille (TOBL,TOBR);
  end;
  //
end;

procedure T_VisuDoc.SetTOBPiece(ThePiece: TOB);
begin
  fTOBpiece := ThePiece;
  fEnHt := (ThePiece.GetValue('GP_FACTUREHT') = 'X');
end;

procedure T_VisuDoc.VoirDocument;
var TheModegestion,TheAction : string;
    THeDocument : TOB;
begin
	TheDocument := nil;
  if ModeGestion = GdeVisuFRF then
  begin
    THeDocument := TOB.Create ('PIECE',nil,-1);
    THeDocument.Dupliquer (TOBPiece,true,true,true);
    TheModeGestion := 'TOFRANCS';
    ConvertitPieceToFrancs(THeDocument);
    TheAction := 'CONSULTATION';
  end else if ModeGestion = GdeSELECT then
    begin
      TheModeGestion := 'SELECT';
  end else if ModeGestion = GdeSelectRF then
  begin
  	if (fTOBTiers = nil) or (fTOBAffaire = nil) then exit;
    THeDocument := TOB.Create ('PIECE',nil,-1);
    TheModeGestion := 'LISTERECEPFOU';
		PrepareAffichageRecep (TheDocument);
    TheAction := 'SELECTION';
  end;
  if TheDocument.detail.count > 0 then
  begin
    TheTOB := THeDocument;
    AglLanceFiche ('BTP','BTSAISDOCEXT','','','ACTION='+TheAction+';MODEGESTION='+TheModegestion);
    if TheTOB <> nil then
    begin
      fTOBresult := TOB.Create ('LES LIGNES RESULTATS',nil,-1);
      fTOBresult.dupliquer (TheTOB,true,true);
      TheTOB.free;
    end;
    TheTOB := nil;
  end;
  if TheDocument <> nil then THeDocument.free;
end;

procedure T_VisuDoc.TrouveRefTiers(CodeArt, NatTiers, CodeTiers: string; TOBART : TOB);
var TobRefTiers: TOB;
  RefTiers: string;
  Qry: TQuery;
  SQl : string;
begin
  if (CodeArt = '') or (NatTiers = '') or (CodeTiers = '') then exit;
  if not TOBArt.FieldExists('REFARTTIERS') then TOBArt.AddChampSup('REFARTTIERS', False);
  if ((NatTiers = 'CLI') or (NatTiers = 'AUD')) then
  begin
    Qry := OpenSQL('SELECT * FROM ARTICLETIERS WHERE GAT_REFTIERS="' + CodeTiers + '"' +
      'AND GAT_ARTICLE="' + CodeArt + '" ', True);
    if not Qry.EOF then
    begin
      TobRefTiers := TOB.Create('ARTICLETIERS', nil, -1);
      TobRefTiers.SelectDB('', Qry);
      RefTiers := TobRefTiers.GetValue('GAT_REFARTICLE');
      TobRefTiers.Free;
    end else
    begin
      RefTiers := '';
    end;
  end else
    if ((NatTiers = 'FOU') or (NatTiers = 'AUC')) then
  begin
    SQL := 'SELECT * FROM CATALOGU WHERE GCA_TIERS="' + CodeTiers + '" AND GCA_ARTICLE="' + CodeArt + '" ' +
      'AND GCA_DATESUP>="' + USDateTime(Date) + '" AND GCA_DATEREFERENCE<="' + USDateTime(Date) + '"';
    Qry := OpenSQL(SQL, True);
    if not Qry.EOF then
    begin
      TobRefTiers := TOB.Create('CATALOGU', nil, -1);
      TobRefTiers.SelectDB('', Qry);
      RefTiers := TobRefTiers.GetValue('GCA_REFERENCE');
      TobRefTiers.Free;
    end else
    begin
      RefTiers := '';
    end;
  end;
  Ferme(Qry);
  TOBArt.PutValue('REFARTTIERS', RefTiers)
end;


procedure TOF_BTSAISDOCEXT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISDOCEXT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISDOCEXT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISDOCEXT.OnLoad ;
begin
  Inherited ;
  definigrille;
  BeforeShow;
end ;

procedure TOF_BTSAISDOCEXT.OnArgument (S : String ) ;
var Arguments,ChampMul,ValMul,Critere : string;
    X : integer;
begin
  Inherited ;
  TOBPiece := LaTOB;
  EnHt := (TOBPiece.GetVAlue('GP_FACTUREHT')='X');
  // Recup des arguments
  Arguments := S;
  repeat
    Critere := uppercase(Trim(ReadTokenSt(Arguments)));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));

        if ChampMul = 'MODEGESTION' then
        begin
          if ValMul = 'TOFRANCS' then
          begin
            ecran.Caption := TraduireMemoire('Visualisation d''un document en francs');
            TypeGestion := GdeVisuFRF
          end else if valmul = 'LISTERECEPFOU' then
          begin
            ecran.Caption := TraduireMemoire('Sélection des réceptions fournisseurs');
            TypeGestion := GdeSelectRF;
          end else
          begin
          end
        end else if ChampMul = 'ACTION' then
        begin
        	if ValMul <> 'SELECTION' then
          begin

          end;
        end;
      end;
    end;
  until Critere = '';
  GetComposants;
end ;

procedure TOF_BTSAISDOCEXT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISDOCEXT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISDOCEXT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISDOCEXT.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('GP_AFFAIRE'));
Aff0:=THEdit(GetControl('GP_AFFAIRE0'));
Aff1:=THEdit(GetControl('GP_AFFAIRE1'));
Aff2:=THEdit(GetControl('GP_AFFAIRE2'));
Aff3:=THEdit(GetControl('GP_AFFAIRE3'));
Aff4:=THEdit(GetControl('GP_AVENANT'));
Tiers:=THEdit(GetControl('GP_TIERS'));
end;

procedure TOF_BTSAISDOCEXT.GetComposants;
begin
  GS := THGrid (GetControl('GS'));
  GS.GetCellCanvas := GetCellCanvas ;
  PENTETE := THPanel (GetControl('PENTETE'));
  PPIED := THPanel (GetControl('PPIED'));
  TitrePiece := THLabel (GetControl('TITREPIECE'));
  REFERENCE := THedit(GetControl('REFERENCE'));
  LCLIENT := THLabel (GetControl('LCLIENT'));
  TDEVISE := THEdit(GetCOntrol('TDEVISE'));
  LDEVISE := THLabel (GetCOntrol('LDEVISE'));
  LDEVISE1 := THLabel (GetCOntrol('LDEVISE1'));
  BOUVRIR := TToolBarButton97(GetControl('Bouvrir'));
  BVALIDER := TToolBarButton97(GetControl('BValider'));
  BRechercher := TToolBarButton97(GetCONTROL('Brechercher'));
  BselectAll := TToolBarButton97(GetCONTROL('BselectAll'));
  if (TypeGestion = GdeSelectRF) then
  begin
  	BOUVRIR.visible := true;
  	BOUVRIR.enabled := true;
    BOuvrir.onclick := BouvrirClick;
    FindDialog := TFindDialog.Create (Ecran);
  	Brechercher.OnClick := BRechercherClick;
  	FindDialog.OnFind := FindDialogFind;
    BselectAll.onClick := BselectAllClick;
  end else
  begin
  	BOUVRIR.visible := false;
  	BVALIDER.enabled := false;
    Brechercher.Visible := false;
    BselectAll.visible := false;
  end;
end;

procedure TOF_BTSAISDOCEXT.DefiniLesColonnes;
var Nam, St, FF, FFP: string;
    i : integer;
begin
  GS.colCount := fNbCols;
  GS.FixedCols := 1;
  St := LesColonnes;
  //
  FF := '#';
  if V_PGI.OkDecQ > 0 then
  begin
    FF := '0.';
    for i := 1 to V_PGI.OkDecQ - 1 do
    begin
      FF := FF + '#';
    end;
    FF := FF + '0';
  end;
  FFP := '#';
  if V_PGI.OkDecP > 0 then
  begin
    FFP := '0.';
    for i := 1 to V_PGI.OkDecP - 1 do
    begin
      FFP := FFP + '#';
    end;
    FFP := FFP + '0';
  end;
  for i := 0 to GS.ColCount - 1 do
  begin
    Nam := ReadTokenSt(St);
    if (Nam = 'GL_NUMLIGNE') then
    begin
      SG_NUML := i;
      GS.ColAligns [i] := taCenter;
      GS.ColLengths [i] := 38;
      GS.ColWidths [i] := 38;
      GS.Cells [i,0] := 'N°';
    end else if (Nam = 'GL_TYPEARTICLE') then
    begin
      SG_TYPEA := i;
      GS.ColAligns [i] := taLeftJustify;
      GS.ColLengths [i] := 0;
      GS.ColWidths [i] := 0;
      GS.Cells [i,0] := '';
    end else if (Nam = 'GL_REFARTSAISIE') then
    begin
      SG_REFA := i;
      GS.ColAligns [i] := taLeftJustify;
      GS.ColLengths [i] := 88;
      GS.ColWidths [i] := 88;
      GS.Cells [i,0] := 'Référence';
    end else if (Nam = 'GL_CODEARTICLE') then
    begin
      SG_REFA := i;
      GS.ColAligns [i] := taLeftJustify;
      GS.ColLengths [i] := 88;
      GS.ColWidths [i] := 88;
      GS.Cells [i,0] := 'Référence';
    end else if (Nam = 'GL_LIBELLE') then
    begin
      SG_LIB := i;
      GS.ColAligns [i] := taLeftJustify;
      GS.ColLengths [i] := 353;
      GS.ColWidths [i] := 353;
      GS.Cells [i,0] := 'Désignation';
    end else if (Nam = 'GL_QTEFACT') then
    begin
      SG_QTE := i;
      GS.ColFormats[i] := FF;
      GS.ColAligns [i] := taRightJustify;
      GS.ColLengths [i] := 63;
      GS.ColWidths [i] := 63;
      GS.Cells [i,0] := 'Quantité';
    end else if (Nam = 'GL_QUALIFQTEVTE') then
    begin
      SG_UV := i;
      GS.ColAligns [i] := taCenter;
      GS.ColLengths [i] := 32;
      GS.ColWidths [i] := 32;
      GS.Cells [i,0] := 'Unité';
    end else if (Nam = 'GL_PUHTDEV') or (Nam = 'GL_PUTTCDEV') then
    begin
      SG_PUV := i;
      GS.ColFormats[i] := FFP;
      GS.ColAligns [i] := taRightJustify;
      GS.ColLengths [i] := 69;
      GS.ColWidths [i] := 69;
      GS.Cells [i,0] := 'Prix';
    end else if (Nam = 'GL_DPA') then
    begin
      SG_PUV := i;
      GS.ColFormats[i] := FFP;
      GS.ColAligns [i] := taRightJustify;
      GS.ColLengths [i] := 69;
      GS.ColWidths [i] := 69;
      GS.Cells [i,0] := 'Prix';
    end else if (Nam = 'GL_MONTANTHTDEV') or (Nam = 'GL_MONTANTTTCDEV') then
    begin
      SG_MONTANT := i;
      GS.ColFormats[i] := FFP; { NEWPIECE }
      GS.ColAligns [i] := taRightJustify ;
      GS.ColLengths [i] := 80;
      GS.ColWidths [i] := 80;
      GS.Cells [i,0] := 'Montant';
    end else if (Nam = 'MTACHAT') then
    begin
      SG_MONTANT := i;
      GS.ColFormats[i] := FFP;
      GS.ColAligns [i] := taRightJustify;
      GS.ColLengths [i] := 80;
      GS.ColWidths [i] := 80;
      GS.Cells [i,0] := 'Montant';
    end else if (Nam = 'SELECT') then
    begin
      SG_SELECT := i;
      GS.ColAligns [i] := taLeftJustify;
      GS.ColLengths [i] := 20;
      GS.ColWidths [i] := 20;
      GS.Cells [i,0] := 'Sel';
    end else if (Nam = '_LIBREFLIVRFOU_') then
    begin
      SG_REFLIVRFOU := i;
      GS.ColAligns [i] := taCenter;
      GS.ColLengths [i] := 255;
      GS.ColWidths [i] := 255;
      GS.Cells [i,0] := 'Provenance';
    end;
  end;
end;

procedure TOF_BTSAISDOCEXT.DefiniGrille;
begin
	if TypeGestion = GdeVisuFRF then
  begin
  	// Visu document en franc
  if EnHt then
  begin
    	fNbCols := 8;
    lesColonnes := 'GL_NUMLIGNE;GL_TYPEARTICLE;GL_REFARTSAISIE;GL_LIBELLE;GL_QTEFACT;GL_QUALIFQTEVTE;GL_PUHTDEV;GL_MONTANTHTDEV;';
  end else
  begin
    lesColonnes := 'GL_NUMLIGNE;GL_TYPEARTICLE;GL_REFARTSAISIE;GL_LIBELLE;GL_QTEFACT;GL_QUALIFQTEVTE;GL_PUTTCDEV;GL_MONTANTTTCDEV;';
    end;
  end else if TypeGestion = GdeSelectRF then
  begin
  	// Sélection des lignes à livrer
    fNbCols := 6;
    lesColonnes := 'SELECT;GL_CODEARTICLE;GL_LIBELLE;GL_QTEFACT;GL_QUALIFQTEVTE;_LIBREFLIVRFOU_';
  end else
  begin
  	// Selection
  end;
  //
  definilesColonnes;
end;

procedure TOF_BTSAISDOCEXT.ReajusteAffichage (Arow : integer);
var TOBL : TOB;
begin
  TOBL := GEtTOBligne (TOBPiece,Arow);
  if TOBL.GetValue('GL_QTEFACT')=0 then GS.Cells[SG_QTE,Arow] := '';
  if EnHt then
  begin
    if TOBL.GetValue('GL_PUHTDEV')=0 then GS.Cells[SG_PUV,Arow] := '';
    if TOBL.GetValue('GL_MONTANTHTDEV')=0 then GS.Cells[SG_MONTANT,Arow] := '';
  end else
  begin
    if TOBL.GetValue('GL_PUTTCDEV')=0 then GS.Cells[SG_PUV,Arow] := '';
    if TOBL.GetValue('GL_MONTANTTTCDEV')=0 then GS.Cells[SG_MONTANT,Arow] := '';
  end;
end;

procedure TOF_BTSAISDOCEXT.AfficheLaLigne (ARow : integer);
var TOBL: TOB;
begin
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  TOBL.PutLigneGrid(GS, ARow, False, False, LesColonnes);
  ReajusteAffichage (Arow);
end;

procedure TOF_BTSAISDOCEXT.BeforeShow;
var Indice : integer;
    DEV : Rdevise;
    TypeRef : string;
begin
  UpdateCaption (ecran);
  TitrePiece.Caption := GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_LIBELLE');
  TypeRef := GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_REFINTEXT');
  if Typeref = 'INT' then
  begin
    REFERENCE.text := TOBPiece.GetValue('GP_REFINTERNE');
  end else if Typeref = 'EXT' then
  begin
    REFERENCE.text := TOBPiece.GetValue('GP_REFINTERNE');
  end else REFERENCE.text := TOBPiece.GetValue('GP_REFINTERNE');

  LCLIENT.Caption := '';
  if (TypeGestion = GdeSelect) or (TypeGestion = GdeSelectRF) then
  begin
//    DEV.Code := TOBPiece.getValue('GP_DEVISE');
//    GetInfosDevise(DEV);
		PENTETE.visible := false;
    PPIED.visible := false;
    GS.MultiSelect := True;
  end else if Typegestion = GdeVisuFRF then
  begin
    DEV.Code := 'FRF';
    GetInfosDevise(DEV);
    PPIED.visible := true;
    GS.MultiSelect := false;
  end;
  LDevise.Caption := DEV.Symbole;
  LDevise1.Caption := DEV.Symbole;
  TDEVISE.Text := DEV.Libelle;
  GS.RowCount := TOBPiece.Detail.Count + 1;
  TOBPiece.PutEcran(Ecran);
  For Indice := 0 to TOBPiece.detail.count -1 do
  begin
    AfficheLaLigne (Indice+1);
  end;
  if (TypeGestion = GdeSelectRF) then
  begin
  	TFFiche(Ecran).HMTrad.ResizeGridColumns (GS);
  end;
end;


procedure TOF_BTSAISDOCEXT.GetCellCanvas(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
var TOBL: TOB;
  //ARect : TRect ;
begin
  if ACol < GS.FixedCols then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  {Ligne non imprimable}
  if TOBL.GetValue('GL_NONIMPRIMABLE') = 'X' then if (ARow <> GS.Row) then Canvas.Font.Color := clBlue;
  {Ligne supprimée}
  if TOBL.FieldExists('SUPPRIME') then
    if TOBL.GetValue('SUPPRIME') = 'X' then if (ARow <> GS.Row) then Canvas.Font.Style := Canvas.Font.Style + [fsStrikeOut];
  {Lignes de sous-total}
  if TOBL.GetValue('GL_TYPELIGNE') = 'TOT' then Canvas.Font.Style := Canvas.Font.Style + [fsBold, fsItalic];
  {Lignes début de paragraphe}// Modif BTP
  (* VARIANTE
  if Copy(TOBL.GetValue('GL_TYPELIGNE'), 1, 2) = 'DP' then
  *)
  if IsDebutParagraphe (TOBL) then
  begin
    if TOBL.GetValue('GL_NIVEAUIMBRIC') > 1 then Canvas.Font.Style := Canvas.Font.Style + [fsItalic]
    else Canvas.Font.Style := Canvas.Font.Style + [fsBold, fsItalic];
  end;
  {Lignes fin de paragraphe}// Modif BTP
  (* VARIANTE  -- if Copy(TOBL.GetValue('GL_TYPELIGNE'), 1, 2) = 'TP' then *)
  if IsFinParagraphe (TOBL) then
  begin
    if TOBL.GetValue('GL_NIVEAUIMBRIC') > 1 then Canvas.Font.Style := Canvas.Font.Style + [fsItalic]
    else Canvas.Font.Style := Canvas.Font.Style + [fsBold, fsItalic];
  end;
  {Ligne de commentaire rattachée}// Modif BTP
  if IsLigneDetail(nil, TOBL) then
  begin
    Canvas.Font.Style := Canvas.Font.Style + [fsbold, fsItalic];
    Canvas.Font.Color := clActiveCaption;
  end;
  {Ligne de retenue de garantie}// Modif BTP
  if (TOBL.GetValue('GL_TYPELIGNE') = 'RG') then
  begin
    Canvas.Font.Style := Canvas.Font.Style + [fsbold, fsItalic];
    Canvas.Font.Color := clRed;
  end;
  {$IFDEF BTP}
  {Ligne Traitée}
  if (TOBL.GetValue('GL_VIVANTE') = '-') and not (IsLigneDetail(nil, TOBL)) then if (ARow <> GS.Row) then
    Canvas.Font.Style := Canvas.Font.Style + [fsStrikeOut];
  {$ENDIF}
  // VARIANTE
  if (isVariante(TOBL)) then
  begin
    Canvas.Font.Style := Canvas.Font.Style + [fsbold,fsItalic];
    Canvas.Font.Color := clMaroon;
  end;
  // Ligne de référence de livraison ... Commande fournisseur ou réception
  if TOBL.GetValue('GL_TYPELIGNE') = 'RL' then
  begin
    Canvas.Font.Style := Canvas.Font.Style + [fsBold];
    Canvas.Font.Color := clRed;
  end;
end;

procedure T_VisuDoc.InsereFille(TOBL, TOBDR: TOB);
var TOBD : TOB;
begin
  TOBD := TOB.Create ('UNE FILLE',TOBL,-1);
  TOBD.dupliquer (TOBDR,false,true);
  TOBD.AddChampSupValeur('_REFLIVRFOU_',TOBDR.GetValue('_REFLIVRFOU_'));     // reference de la recption en cours
  TOBD.AddChampSupValeur('_NUMRECEPFOUR_',TOBDR.GetValue('_NUMRECEPFOUR_')); // reference de la reception pour le retour
  TOBD.AddChampSupValeur('BCO_NUMMOUV',TOBDR.GetValue('BCO_NUMMOUV'));
  TOBD.AddChampSupValeur('BCO_QUANTITE',TOBDR.GetValue('BCO_QUANTITE'));
  TOBD.AddChampSupValeur('BCO_QTEVENTE',TOBDR.GetValue('BCO_QTEVENTE'));
  TOBD.AddChampSupValeur('BCO_QTELIVRE',TOBDR.GetValue('BCO_QTELIVRE'));
  TOBD.AddChampSupValeur('BCO_QTERETOUR',TOBDR.GetValue('BCO_QTERETOUR'));
  TOBD.AddChampSupValeur('BCO_LIENVENTE',TOBDR.GetValue('BCO_LIENVENTE'));
  TOBD.AddChampSupValeur('MTACHAT',TOBDR.GetValue('MTACHAT'));
  TOBD.AddChampSupValeur('_NATURE_',TOBDR.GetValue('_NATURE_'));
  TOBD.AddChampSupValeur('A TRAITER','-');
end;

procedure TOF_BTSAISDOCEXT.BouvrirClick(Sender: Tobject);
var UneTOB : TOB;
		Indice : integer;
begin
  Inherited ;
  TheTOB := TOB.Create ('LE RESULTAT',nil,-1);
  for Indice := 0 to TOBpiece.detail.count -1 do
	begin
  	if GS.IsSelected (Indice+1) then
    begin
    	UneTOB := TOB.Create ('LIGNESEL',TheTOB,-1);
      UneTOB.Dupliquer (TOBPiece.detail[Indice],true,true);
    end;
  end;
  ecran.Close;
end;


procedure TOF_BTSAISDOCEXT.BRechercherClick(Sender: TObject);
begin
  FirstFind := true;
  FindDialog.Execute;
end;


procedure TOF_BTSAISDOCEXT.FindDialogFind(Sender: TObject);
begin
  Rechercher(GS, FindDialog, FirstFind);
end;

procedure TOF_BTSAISDOCEXT.BselectAllClick (Sender : Tobject);
begin
  GS.AllSelected := not GS.AllSelected;
end;

Initialization
  registerclasses ( [ TOF_BTSAISDOCEXT ] ) ;
end.

