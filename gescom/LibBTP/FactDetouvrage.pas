unit FactDetouvrage;

(* UNITE PERMETTANT DE MODIFIER LE SOUS DETAIL D'UN OUVRAGE DE PREMIER NIVEAU DIRECTEMENT DEPUIS LA SAISIE DE PIECE *)

interface
Uses Classes,
		 forms,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
		 AglInitGc,
     Hent1,
     HCtrls,
     EntGc,
     HMsgBox,
     SysUtils,
     AglInit,
     UentCommun,
     SaisUtil,
     NomenUtil,
     Choix,
     uTob;

type

	TmodeInsert = (TmiCurrent,TmiUp);
  TChpModif = (TchpQte,TChpPerte,TchpRendement);

	TModifSousDetail = class (Tobject)
  	private
    	XX : Tform;
    	fTOBpiece : TOB;
      fTOBaffaire : TOB;
      fTOBOuvrages : TOB;
      fTOBtiers : TOB;
      fTOBArticles : TOB;
      fDev : Rdevise;
      ART : THCritMaskEdit;
      fEnHt : boolean;
      //
      function FindDetailOuvrage (TOBL : TOB) : TOB;
    	procedure AfficheLaLigneSousDetail(TOBOD, TOBL, TOBPere: TOB);
      procedure CalculeLaLigneDoc (TOBL : TOB; WithRecalc : boolean=true);
      function EnHt (TOBL : TOB) : boolean;
    	function VerifNomen(TobArt: TOB; Arow: integer; var retour,Libelle: string): boolean;
    	function chargeArticle(TheArticle: string; TOBA: TOB): boolean;
    	function GetCodeArticle(var TypeArticle, Article, CodeArticle,libelle: string; Arow, Acol: integer): Boolean;
    	procedure SetLigneOuvrage(TOBLO, TOBL, TOBOD: TOB; Article,CodeArticle: string; Libelle: string='');
    	procedure ReInitligneOuv(TOBOD: TOB);
    	function ChargeOuvrage(TOBL: TOB; CodeNomen: string; TobOuvrage: TOB; Niv,OrdreCompo: integer; var Tresult: T_valeurs): boolean;
      procedure RenumeroteLesFilles(TOBLN: TOB; Niveau: integer);
      procedure SetNumerotation(TOBL, TOBMere: TOb; Niveau, Arow: integer);
      procedure AjouteMtAvancOuvrage(TOBL: TOB; montant: double; Sens,Typefacturation: string);
      procedure AppliqueCoefSousDetail(TOBO: TOB; Coef,PP: Double);
    public
    	constructor create (Theform : Tform);
      destructor destroy; override;
      property TOBpiece : TOB read fTOBpiece write fTOBpiece;
      property TOBaffaire : TOB read fTOBAffaire write fTOBAffaire;
      property TOBOuvrages : TOB read fTOBOuvrages write fTOBOuvrages;
      property TOBArticles : TOB read fTOBArticles write fTOBArticles;
      property TOBTiers : TOB read fTOBtiers write fTOBtiers;
      //
      property DEV : Rdevise read fDev  write fDev;
      // -- générique --
      function GetTypeFacturation : string;
      // --  Modif via la saisie --
      procedure ModifLibelleSousDet (TOBL : TOB; libelle : string);
      function ModifQteSousDet (TOBL : TOB; Qte : double; DEV : RDevise) : boolean;
      function ModifPASousDet (TOBL : TOB; Pua : double) : boolean;
      function ModifPVSousDet (TOBL : TOB; Pu : double) : boolean;
      procedure ModifUniteSousDet (TOBL : TOB; Unite : string);
    	procedure ModifPourcentAvancOuv(TOBL: TOB; Pourcent: double;Typefacturation: string; DEV: Rdevise);
    	function ModifPourcentAvancSousDet(TOBL: TOB; Pourcent: double;Typefacturation: string; DEV : Rdevise) : boolean;
      procedure ModifQteAvancOuv (TOBL: TOB; Qte: double;Typefacturation: string; DEV : Rdevise);
    	procedure ModifQteAvancSousDet(TOBL: TOB; Qte: double;Typefacturation: string; DEV : Rdevise);
    	procedure ModifMontantAvancOuv(TOBL: TOB; Montant: double;Typefacturation: string; DEV : Rdevise);
    	procedure ModifMontantAvancSousDet(TOBL: TOB; Montant: double;Typefacturation: string; DEV : Rdevise);
      procedure ModifCoefMargOuv (TOBL: TOB; Coef,PP: double; DEV : Rdevise);
      procedure ModifCoefMargSousDet(TOBL: TOB; Coef,PP: double;DEV: Rdevise);
      procedure ModifCoefMarqSousDet(TOBL: TOB; CM, PM: double;DEV: Rdevise);
      procedure ModifCoefMarqOuv(TOBL: TOB; CM, PM: double; DEV: Rdevise);
      function ModifQteSaisSousDet(TOBL: TOB; Ichps: TChpModif; Valeur: double; DEV: RDevise): boolean;

      procedure TraiteCodeArticle (TOBL : TOB; CodeArticle : string; Libelle : string='');
      procedure RechercheCodeArticle (Acol,Arow : integer; TOBL : TOB; CodeArticle : string);
      // -- Actions sur une ligne --
      procedure AddSousDetail (TOBL : TOB; ModeInsert : TmodeInsert=TmiCurrent);
      procedure DelSousDetail(TOBL : TOB);
      // -----
      procedure GetBlocNote (TOBL : TOB);
    	procedure SetBlocNote(TOBL: TOB);
      // --
      function FindLigneOuvrageinDoc (TOBD : TOB) : TOB;
      function GetLigneOuvrage (TOBL : TOB) : integer;
      procedure SetChange;
  end;

procedure RenseigneTOBOuv(TOBPiece,TOBOuv, TOBL, TobArticle: TOB;libelle: string; DEV : RDevise; AvecPv: boolean=true);

implementation
uses factureBTP,factOuvrage,UtilPGI,FactArticle,facture,FactTOB,
		 FactUtil,Factvariante,FactComm,factPiece,UtilArticle,
     TarifUtil,FactDomaines,BTPUtil,UtilTOBPiece,UCotraitance,
     ParamSoc,UspecifPOC;

{ TModifSousDetail }

procedure TModifSousDetail.AddSousDetail(TOBL: TOB; ModeInsert : TmodeInsert);
var TOBRef,TOBOD,TOBOUV,TOBLDet,TOBNL,TOBLRef : TOB;
		Numlig,IndiceNomen,Arow : integer;
    NatureTravail : double;
    LibelleFou : string;
    QteSit,MtSit : double;
begin
  if TFFacture(XX).Action = taConsult then Exit;

	Arow := TOBL.getindex+1; // emplacement de la ligne dans le document
  //
  if ModeInsert = TmiCurrent then
  begin
  	TOBLRef := TOBL;
  end else
  begin
  	TOBLRef := GetTOBLigne(fTOBpiece,Arow-1);
  end;

  TOBRef := FindLigneOuvrageinDoc (TOBLRef); // la ligne dans le document
  if TOBref = nil then exit;

  if (Pos(GetTypeFacturation,'DAC;AVA')>0) and
     (Pos(TOBPiece.getValue('GP_NATUREPIECEG'),'FBT;DAC;FBP;BAC')>0) and
     (TOBL.GetString('GL_PIECEPRECEDENTE')<>'') then Exit;

  if IsLignefacture (TOBREF) then
  begin
    PgiError(Traduirememoire('Impossible : une facturation a déjà été effectuée sur cet ouvrage.'));
    exit;
  end;
  NatureTravail := valeur(TOBref.getValue('GLC_NATURETRAVAIL'));
  LibelleFou := TOBRef.getValue('LIBELLEFOU');
  if ModeInsert = TmiCurrent then
  begin
    TOBOD := FindDetailOuvrage (TOBLREF);
    if TOBOD = nil then exit;
    Numlig := TOBOD.GetIndex;
  end else
  begin
    Numlig := -1;
  end;

  IndiceNomen := TOBRef.getValue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
  TOBOUV := fTOBOuvrages.Detail[IndiceNomen-1];
  // création du nouveau sous détail
  TOBLDet := TOB.create('LIGNEOUV',TOBOUV,numlig);
  InsertionChampSupOuv (TOBLDET,false);
  NumeroteLigneOuv (TOBOUV,TOBRef,1,1,0,0,0);
  //
  fTOBPiece.putValue('GP_UNIQUEBLO',fTOBPiece.getValue('GP_UNIQUEBLO')+1);
  TOBLDet.putValue('BLO_UNIQUEBLO',fTOBPiece.getValue('GP_UNIQUEBLO'));
  if (natureTravail > 0) and (NatureTravail <= 10) then
  begin
  	TOBLDet.putValue('BLO_NATURETRAVAIL',TOBRef.GetValue('GLC_NATURETRAVAIL'));
  	TOBLDet.putValue('BLO_FOURNISSEUR',TOBRef.GetValue('GL_FOURNISSEUR'));
    if TOBLDet.FieldExists('BLF_NATURETRAVAIL') then
    begin
      TOBLDet.putvalue('BLF_NATURETRAVAIL',TOBRef.GetValue('GLC_NATURETRAVAIL'));
      TOBLDet.putvalue('BLF_FOURNISSEUR',TOBRef.GetValue('GL_FOURNISSEUR'));
    end;
  end;
  TOBLDet.putvalue('LIBELLEFOU',LibelleFou);

  // création de la nouvelle ligne dans la grille et dans la piece
  TFFacture (XX).GS.CacheEdit;
  TFFacture (XX).GS.SynEnabled := False;
  InsertTOBLigne(fTOBPiece, ARow);
  TFFacture(XX).GS.InsertRow(ARow);
  InitialiseTOBLigne(fTOBPiece, fTOBTiers, fTOBAffaire, ARow);
  TFfacture(XX).GS.Row := ARow;
  TFFacture(XX).GS.MontreEdit;
  TFFacture(XX).GS.SynEnabled := True;
  //
  TOBNL := GetTOBligne(fTOBpiece,Arow);
  if (IsVariante(TOBRef)) then TOBNL.PutValue('GL_TYPELIGNE', 'SDV')
                          else TOBNL.PutValue('GL_TYPELIGNE', 'SD');
  TOBNL.putValue('GL_INDICENOMEN',TOBref.getValue('GL_INDICENOMEN'));
  TOBNL.putValue('UNIQUEBLO',TOBLDet.getValue('BLO_UNIQUEBLO'));
  NumeroteLignesGC(TFFACTURE(XX).GS, fTOBPiece);
  TOBLDET.putValue('BLO_NUMORDRE',TOBLDet.GetIndex+1);
  TFFACTURE(XX).ShowDetail(ARow);
  TFFACTURE(XX).StCellCur := '';
end;

procedure TModifSousDetail.AfficheLaLigneSousDetail (TOBOD,TOBL,TOBpere : TOB);
var MontantPA,MontantAchat,MontantDev,QteDuDetail,QteDuPv,QteOuv : double;
begin
  //
  QteDuDetail := TOBOD.GetValue('BLO_QTEDUDETAIL'); if QteDuDetail = 0 then QteDuDetail := 1;
  QteDuPv := TOBOD.GetValue('BLO_PRIXPOURQTE'); if QteDuPv = 0 then QteDuPv := 1;
  QteOuv := TOBPere.GetValue('GL_QTEFACT');
  //
  MontantPA :=Arrondi((TOBOD.GetValue('BLO_QTEFACT')*TOBOD.GetValue('BLO_DPA'))/QteDuPv,V_PGI.okdecP);
  //
  if TOBOD.GetValue('BLO_TYPEARTICLE') <> 'POU' then
  begin
    if TFFacture(XX).AffSousDetailUnitaire then
    begin
      MontantAchat := Arrondi((MontantPA)/(QteDuPv*QteDuDetail),V_PGI.okdecP);
      if EnHt(TOBL) then
      begin
        MontantDev :=Arrondi(( TOBOD.GetValue('BLO_QTEFACT')*TOBOD.getValue('BLO_PUHTDEV'))/(QteDuPv*QteDuDetail),V_PGI.okdecP);
      end else
      begin
        MontantDev :=Arrondi((TOBOD.GetValue('BLO_QTEFACT')*TOBOD.getValue('BLO_PUTTCDEV'))/(QteDuPv*QteDuDetail),V_PGI.okdecP);
      end;
    end else
    begin
    MontantAchat := Arrondi((MontantPA * QteOuv)/(QteDuPv*QteDuDetail),V_PGI.okdecP);
    if EnHt(TOBL) then
    begin
      MontantDev :=Arrondi(( QteOuv * TOBOD.GetValue('BLO_QTEFACT')*TOBOD.getValue('BLO_PUHTDEV'))/(QteDuPv*QteDuDetail),V_PGI.okdecP);
    end else
    begin
      MontantDev :=Arrondi((QteOuv * TOBOD.GetValue('BLO_QTEFACT')*TOBOD.getValue('BLO_PUTTCDEV'))/(QteDuPv*QteDuDetail),V_PGI.okdecP);
    end;
  end;
  end;
  //
  TOBL.PutValue('GL_CODEARTICLE', TOBOD.GetValue('BLO_CODEARTICLE'));
  TOBL.PutValue('GL_ARTICLE', TOBOD.GetValue('BLO_ARTICLE'));
  TOBL.PutValue('GL_REFARTSAISIE', TOBOD.GetValue('BLO_CODEARTICLE'));
  TOBL.PutValue('GL_LIBELLE', TOBOD.GetValue('BLO_LIBELLE'));
  TOBL.PutValue('GL_TYPEARTICLE', TOBOD.GetValue('BLO_TYPEARTICLE'));
  TOBL.PutValue('GL_INDICENOMEN', TOBpere.getValue('GL_INDICENOMEN'));
  TOBL.PutValue('GL_FAMILLETAXE1', TOBOD.GetValue('BLO_FAMILLETAXE1'));
  TOBL.PutValue('UNIQUEBLO', TOBOD.GetValue('BLO_UNIQUEBLO'));
  TOBL.PutValue('GLC_NATURETRAVAIL', TOBOD.GetValue('BLO_NATURETRAVAIL'));
  TOBL.SetDOuble('GL_QTESAIS', TOBOD.GetValue('BLO_QTESAIS'));
  TOBL.SetDOuble('GL_PERTE', TOBOD.GetValue('BLO_PERTE'));
  TOBL.SetDOuble('GL_RENDEMENT', TOBOD.GetValue('BLO_RENDEMENT'));

  if TOBOD.GetValue('BLO_NATURETRAVAIL') <> '' then TOBL.PutValue('LIBELLEFOU', TOBOD.GetValue('LIBELLEFOU'));
  //
  if TOBOD.GetValue('BLO_TYPEARTICLE') = 'POU' then
  begin
    TOBL.PutValue('GL_QTEFACT', TOBOD.GetValue('BLO_QTEFACT')/(QteDuPv*QteDuDetail) )
  end else
  begin
    if TFFacture(XX).AffSousDetailUnitaire then
    begin
      TOBL.PutValue('GL_QTEFACT', TOBOD.GetValue('BLO_QTEFACT'));
    end else
    begin
      TOBL.PutValue('GL_QTEFACT', (TOBOD.GetValue('BLO_QTEFACT') * TOBpere.GetValue('GL_QTEFACT')/(QteDuPv*QteDuDetail)));
    end;
  end;


  TOBL.PutValue('GL_QTERESTE',TOBL.getValue('GL_QTEFACT'));
  TOBL.PutValue('GL_QTESTOCK',TOBL.getValue('GL_QTEFACT'));

  TOBL.PutValue('GL_QUALIFQTEVTE', TOBOD.GetValue('BLO_QUALIFQTEVTE'));
  TOBL.PutValue('GL_DPA', TOBOD.GetValue('BLO_DPA') / QteDupv);
  if EnHt (TOBL) then TOBL.PutValue('GL_PUHTDEV', Arrondi(TOBOD.GetValue('BLO_PUHTDEV'),V_PGI.okdecP))
                 else TOBL.PutValue('GL_PUTTCDEV', Arrondi(TOBOD.GetValue('BLO_PUTTCDEV'),V_PGI.okdecP));
  if TOBOD.GetValue('BLO_TYPEARTICLE') <> 'POU' then
  begin
    TOBL.PutValue('GL_MONTANTPA', MontantAchat);
    if EnHt(TOBL) then TOBL.PUtValue('GL_MONTANTHTDEV', MontantDEV)
                  else TOBL.PUtValue('GL_MONTANTTTCDEV', MontantDEV);
  end;
  // --- GUINIER ---
  if CtrlOkReliquat(TOBL, 'GL') then TOBL.PutValue('GL_QTERESTE',TOBL.getValue('GL_MONTANTHTDEV'));
  //
  TOBL.PutValue('BLF_POURCENTAVANC',TOBOD.getValue('BLF_POURCENTAVANC'));
  TOBL.PutValue('BLF_QTECUMULEFACT',TOBOD.getValue('BLF_QTECUMULEFACT'));
  TOBL.PutValue('BLF_QTESITUATION',TOBOD.getValue('BLF_QTESITUATION'));
  TOBL.PutValue('BLF_MTSITUATION',TOBOD.getValue('BLF_MTSITUATION'));
  //
  TFFActure(XX).AfficheLaLigne (TOBL.GetIndex+1);
	TFFActure(XX).GS.InvalidateRow(TOBL.GetIndex+1);
end;

procedure TModifSousDetail.CalculeLaLigneDoc(TOBL: TOB; WithRecalc : boolean);
var indiceNomen : integer;
		TOBOP : TOB;
    valeurs : T_Valeurs;
    Qte,QTeDudetail : double;
begin
  IndiceNomen := TOBL.getValue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
  TOBOP := fTOBOuvrages.Detail[IndiceNomen-1];
  if Withrecalc then
  begin
    QteDuDetail := TOBOP.getDouble('BLO_QTEDUDETAIL'); if QTeDudetail = 0 then QteDuDetail := 1;
    InitTableau (Valeurs);
    CalculeOuvrageDoc (TOBOP,1,1,true,DEV,valeurs,(fTOBPIECE.GetValue('GP_FACTUREHT')='X'));
    GetValoDetail (TOBOP); // pour le cas des Article en prix posés
    Qte := TOBL.Getvalue('GL_QTEFACT'); Qte := Qte / QteDUDetail;
    TOBL.Putvalue('GL_MONTANTPAFG',valeurs[10]*Qte);
    TOBL.Putvalue('GL_MONTANTPAFR',valeurs[11]*Qte);
    TOBL.Putvalue('GL_MONTANTPAFC',valeurs[12]*Qte);
    TOBL.Putvalue('GL_MONTANTFG',valeurs[13]*Qte);
    TOBL.Putvalue('GL_MONTANTFR',valeurs[14]*Qte);
    TOBL.Putvalue('GL_MONTANTFC',valeurs[15]*Qte);
    TOBL.Putvalue('GL_MONTANTPA',Arrondi((Qte * TOBL.GetValue('GL_DPA')),V_PGI.okdecV));
    TOBL.Putvalue('GL_MONTANTPR',Arrondi((Qte * TOBL.GetValue('GL_DPR')),V_PGI.okdecV));

    TOBL.Putvalue('GL_DPA',valeurs[16]);
    TOBL.Putvalue('GL_DPR',valeurs[17]);
    //
    TOBL.Putvalue('GL_PUHTDEV',valeurs[2]);
    TOBL.Putvalue('GL_PUTTCDEV',valeurs[3]);
    TOBL.Putvalue('GL_PUHT',DeviseToPivotEx(TOBL.GetValue('GL_PUHTDEV'),DEV.Taux,DEV.quotite,V_PGI.OkdecP));
    TOBL.Putvalue('GL_PUTTC',DevisetoPivotEx(TOBL.GetValue('GL_PUTTCDEV'),DEV.taux,DEV.quotite,V_PGI.okdecP));
    TOBL.Putvalue('GL_PUHTBASE',TOBL.GetValue('GL_PUHT'));
    TOBL.Putvalue('GL_PUTTCBASE',TOBL.GetValue('GL_PUTTC'));
    TOBL.Putvalue('GL_DPA',valeurs[0]);
    TOBL.Putvalue('GL_DPR',valeurs[1]);
    TOBL.Putvalue('GL_PMAP',valeurs[6]);
    TOBL.Putvalue('GL_PMRP',valeurs[7]);
    TOBL.putvalue('GL_TPSUNITAIRE',valeurs[9]);
    StockeInfoTypeLigne (TOBL,Valeurs);
  end;
  if (TOBL.GetString('GL_PIECEPRECEDENTE')='') and (TOBL.FieldExists('BLF_QTESITUATION')) then
  begin
    TOBL.PutValue('BLF_QTESITUATION', TOBL.GetValue('GL_QTEFACT')); { NEWPIECE }
    TOBL.PutValue('BLF_MTSITUATION',Arrondi(TOBL.GetValue('BLF_QTESITUATION')*TOBL.GetValue('GL_PUHTDEV')/TOBL.GEtValue('GL_PRIXPOURQTE'),DEV.decimale));
    TOBL.PutValue('BLF_POURCENTAVANC',100);
  end;
  TOBL.PutValue('GL_RECALCULER', 'X');
  ZeroligneMontant (TOBL);
  TFFacture(XX).CalculheureTot (TOBL);
  TOBPiece.PutValue('GP_RECALCULER', 'X');
  TFFacture(XX).CalculeLaSaisie(-1, TOBL.getIndex+1, True);
end;

constructor TModifSousDetail.create (TheForm : Tform);
begin
 XX := TheForm;
 ART := THCritMaskEdit.Create (XX);
end;

procedure RenseigneTOBOuv(TOBPIECE,TOBOuv,TOBL, TobArticle : TOB; libelle : string;DEV : Rdevise; AvecPv:boolean);
var SavTypeLigne : string;
    MTPAF         : Double;
    UA            : string;
begin

  MTPAF := 0;

	SavTypeLigne := TOBOUV.GetValue('BLO_TYPELIGNE'); // on sauvegarde avant de se faire avoir par les fonctions en dessous

  CopieOuvFromLigne (TOBOuv,TOBL);

  TOBOUv.putValue('BLO_TYPELIGNE',SavTypeLigne);

  if TobArticle <> nil then
  begin
    // Modif BTP
    TOBOuv.PutValue('BLO_TENUESTOCK',  TobArticle.GetValue('GA_TENUESTOCK'));
    TOBOuv.PutValue('BLO_QUALIFQTEVTE',TobArticle.GetValue('GA_QUALIFUNITEVTE'));
    TOBOuv.PutValue('BLO_QUALIFQTESTO',TobArticle.GetValue('GA_QUALIFUNITESTO'));
    TOBOuv.PutValue('BLO_PRIXPOURQTE', TobArticle.GetValue('GA_PRIXPOURQTE'));
    MTPAF := TobArticle.GetValue('GA_PAHT');
    TOBOuv.PutValue('BLO_PMAP',TobArticle.GetValue('GA_PMAP'));

    //FV1 : 05/03/2015 - Gestion du Prix d'Achat dans la fiche Article...
    if MTPAF = 0 then
    Begin
     MTPAF    := TOBArticle.GetVAlue('GA_PAUA');
     UA       := TOBArticle.GetValue('GA_QUALIFUNITEACH');
     if MTPAF = 0 Then
     begin
       MTPAF  := TOBArticle.GetValue('GA_PAHT');
       UA     := TOBArticle.getValue('GCA_QUALIFUNITEVTE');
     end;
    End;
    //
    TOBOuv.PutValue ('BLO_DPA', MTPAF);
    TOBOuv.PutValue('BLO_QUALIFQTEACH', UA);
    //
    TOBOuv.PutValue('ANCPA',TobOuv.GetValue('BLO_DPA'));
    TOBOuv.PutValue('ANCPR',TobOuv.GetValue('BLO_DPR'));   
    CopieOuvFromArt (TOBPiece,TOBOuv,TobArticle,DEV);

    if AvecPv then
    begin
      TOBOuv.PutValue('BLO_PUHT',     TobArticle.GetValue('GA_PVHT'));
      TOBOuv.PutValue('BLO_PUHTBASE', TobArticle.GetValue('GA_PVHT'));
      TOBOuv.PutValue('BLO_PUHTDEV',  pivottodevise(TobOuv.GetValue('BLO_PUHTBASE'),DEV.Taux,DEV.quotite,V_PGI.OkdecP ));
      TOBOuv.PutValue('ANCPV',        TobOuv.GetValue('BLO_PUHTDEV'));
      TOBOuv.PutValue('BLO_PUTTC',    TobArticle.GetValue('GA_PVTTC'));
      TOBOuv.PutValue('BLO_PUTTCBASE',TobArticle.GetValue('GA_PVTTC'));
      TOBOuv.PutValue('BLO_PUTTCDEV', pivottodevise(TOBOuv.GetValue('BLO_PUTTCBASE'),DEV.Taux,DEV.quotite,V_PGI.OkdecP ));
    end;

    // Modif BTP
    if Libelle <> '' then TOBOuv.PutValue('BLO_LIBELLE',Libelle);
    TOBOUv.putValue('BNP_TYPERESSOURCE',TOBArticle.getValue('BNP_TYPERESSOURCE'));
    //
    if (Pos (TOBOUV.getValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne) > 0) and (not IsLigneExternalise(TOBOUV)) then
      TOBOUV.PutValue('BLO_TPSUNITAIRE',1)
    else
      TOBOUV.PutValue('BLO_TPSUNITAIRE',0);
    //
    if (IsPrestationSt(TOBOUV)) then
    begin
      if TOBPiece.getValue('GP_APPLICFGST')='X' then
      begin
        TOBOuv.putValue('BLO_NONAPPLICFRAIS','-');
      end else
      begin
        TOBOuv.putValue('BLO_NONAPPLICFRAIS','X');
      end;
      if TOBPiece.getValue('GP_APPLICFCST')='X' then
      begin
        TOBOuv.putValue('BLO_NONAPPLICFC','-');
      end else
      begin
        TOBOuv.putValue('BLO_NONAPPLICFC','X');
      end;
    end else
    begin
      TOBOuv.putValue('BLO_NONAPPLICFRAIS','-');
      TOBOuv.putValue('BLO_NONAPPLICFC','-');
    end;

    if TobArticle.getValue('GA_PRIXPASMODIF')<>'X' Then
    begin
      AppliqueCoefDomaineActOuv (TOBOuv); //,TobArticle);
      if VH_GC.BTCODESPECIF = '001' then ApliqueCoeflignePOC (TOBOuv,DEV);
    end;

    CalculeLigneAcOuv (TOBOUV,DEV,true,TOBARTICLE);

    if (Valeur(TOBL.getValue('GLC_NATURETRAVAIL')) <> 0) and (valeur(TOBL.getValue('GLC_NATURETRAVAIL'))<10) then
    begin
      TOBOuv.PutValue('BLO_NATURETRAVAIL',TOBL.getValue('GLC_NATURETRAVAIL'));
      TOBOuv.PutValue('BLO_FOURNISSEUR',TOBL.getValue('GL_FOURNISSEUR'));
      TOBOuv.PutValue('GA_FOURNPRINC',TOBL.getValue('GL_FOURNISSEUR'));
    end;
    TOBOuv.SetBoolean('GESTRELIQUAT',TobArticle.GetBoolean('GA_RELIQUATMT'));
  end else
  begin
    if AvecPv then
    begin
      TOBOuv.PutValue('BLO_PUHT',0);
      TOBOuv.PutValue('BLO_PUHTBASE',0);
      TOBOuv.PutValue('BLO_PUHTDEV',0);
      TOBOuv.PutValue('ANCPV',0);
      TOBOuv.PutValue('BLO_PUTTC',0);
      TOBOuv.PutValue('BLO_PUTTCBASE',0);
      TOBOuv.PutValue('BLO_PUTTCDEV',0);
      TOBOuv.PutValue('BLO_DPA',0);
      TOBOuv.PutValue('BLO_DPR',0);
      TOBOuv.PutValue('BLO_PMAP',0);
      TOBOuv.PutValue('BLO_PMRP',0);
      TOBOuv.PutValue('BLO_QUALIFQTEVTE','');
      TOBOuv.PutValue('BLO_QUALIFQTESTO','');
      TOBOuv.PutValue('BLO_QUALIFQTEACH','');
    end;
    TOBOuv.PutValue('ANCPA',0);
    TOBOuv.PutValue('ANCPR',0);
    CalculeLigneAcOuv (TOBOUV,DEV,True,TOBARTICLE);
  end;
end;

procedure TModifSousDetail.DelSousDetail(TOBL: TOB);
var TOBref,TOBOD,TOBOUV : TOB;
		IndiceNomen: integer;
    Acol,Arow : integer;
    bc,cancel : boolean;
begin
//  if (Pos(GetTypeFacturation,'DAC;AVA')>0) and (Pos(TOBPiece.getValue('GP_NATUREPIECEG'),'FBT;DAC;FBP;BAC')>0) then Exit;

	Arow := TFFActure(XX).GS.row;
  TOBRef := FindLigneOuvrageinDoc (TOBL); // la ligne d'ouvrage dans le document
  if TOBref = nil then exit;
  
  if (Pos(GetTypeFacturation,'DAC;AVA')>0) and
     (Pos(TOBPiece.getValue('GP_NATUREPIECEG'),'FBT;DAC;FBP;BAC')>0) and
     (TOBRef.GetString('GL_PIECEPRECEDENTE')<>'') then Exit;

  if IsLignefacture (TOBREF) then
  begin
    PgiError(Traduirememoire('Impossible : une facturation a déjà été effectuée sur cet ouvrage.'));
    exit;
  end;
  //
  TFFActure(XX).deduitLaLigne (TOBRef); // pour calculer correctement la pièce
  //
  TOBOD := FindDetailOuvrage (TOBL); // le détail de l'ouvrage
  if TOBOD = nil then exit;
  //
  IndiceNomen := TOBL.getValue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
  TOBOUV := fTOBOuvrages.Detail[IndiceNomen-1]; // l'ouvrage
  TOBOD.free; // Suppression du sous détail dans l'ouvrage
  NumeroteLigneOuv (TOBOUV,TOBRef,1,1,0,0,0);
  //
  //
  TFFacture(XX).GS.BeginUpdate;
  TFFacture(XX).GS.CacheEdit;
  TFFacture(XX).GS.SynEnabled := False;
  TFFacture(XX).GS.DeleteRow(Arow);
  if ((fTOBPiece.Detail.Count > 1) and (ARow <> fTOBPiece.Detail.Count)) then
  begin
    fTOBPiece.Detail[ARow - 1].Free;
  end
  else
  begin
    fTOBPiece.Detail[ARow - 1].InitValeurs;
    InitLesSupLigne(fTOBPiece.Detail[ARow - 1]);
    InitialiseTOBLigne(fTOBPiece, fTOBTiers, fTOBAffaire, Arow);
  end;
  JustNumerote (fTOBPiece,Arow-1);
  CalculeLaLigneDoc (TOBRef);
  TFFacture(XX).GS.MontreEdit;
  TFFacture(XX).GS.SynEnabled := true;
  TFFacture(XX).GS.EndUpdate;
  ACol := TFfacture(XX).GS.fixedcols;
  TFFActure(XX).GSRowEnter(TFFActure(XX).GS, ARow, bc, False);
  TFfacture(XX).GSCellEnter(TFFActure(XX).GS, ACol, ARow, Cancel);
  TFfacture(XX).GS.col := ACol;
  TFfacture(XX).GS.row := Arow;
  TFfacture(XX).StCellCur := TFfacture(XX).GS.cells[Acol,Arow];
  TFfacture(XX).ShowDetail(ARow);
end;

destructor TModifSousDetail.destroy;
begin
  ART.free;
  inherited;
end;

function TModifSousDetail.EnHt(TOBL: TOB): boolean;
begin
	result :=  (TOBL.getValue('GL_FACTUREHT')='X');
end;

function TModifSousDetail.FindDetailOuvrage(TOBL: TOB): TOB;
var TOBOP : TOB;
		indiceNomen : integer;
    NumUnique,Indice : integer;
begin
	result := nil;
  IndiceNomen := TOBL.getValue('GL_INDICENOMEN');
  NumUnique := TOBL.getValue('UNIQUEBLO');
  TOBOP := fTOBOuvrages.detail[IndiceNomen-1];
  if TOBOP <> nil then
  begin
  	for Indice := 0 to TOBOP.detail.count -1 do
    begin
      if TOBOP.detail[Indice].getValue('BLO_UNIQUEBLO')=NumUnique then
      begin
      	result := TOBOP.detail[Indice];
        break;
      end;
    end;
  end;
end;

function TModifSousDetail.FindLigneOuvrageinDoc(TOBD: TOB): TOB;
var Inddep,Indice : integer;
		indOuvrage : integer;
begin
	result := nil;
  Inddep := TOBD.GetIndex ;
  IndOuvrage := TOBD.GetValue('GL_INDICENOMEN');
  for Indice := Inddep-1 downto 0 do
  begin
    if (ISOuvrage(TOBpiece.detail[Indice])) and
    	 (TOBpiece.detail[Indice].getValue('GL_INDICENOMEN')=IndOuvrage) and (IsArticle(TOBpiece.detail[Indice])) then
    begin
    	result := TOBpiece.detail[Indice];
      break;
    end;
  end;
end;


function TModifSousDetail.GetLigneOuvrage(TOBL: TOB): integer;
var TheTOB : TOB;
begin
	result := -1;
  TheTOB := FindLigneOuvrageinDoc(TOBL);
  if TheTOB<> nil then result := TheTOB.GetIndex+1;
end;


procedure TModifSousDetail.TraiteCodeArticle(TOBL: TOB;CodeArticle: string; Libelle : string);
begin

end;

procedure TModifSousDetail.ModifLibelleSousDet(TOBL: TOB; libelle: string);
var TOBOD : TOB;
begin
  TOBOD := FindDetailOuvrage (TOBL);
  if TOBOD = nil then exit;
  TOBOD.putValue('BLO_LIBELLE',Libelle);
  TOBL.putValue('GL_LIBELLE',Libelle);
  TFFActure(XX).AfficheLaLigne (TOBL.GetIndex+1);
end;

Function TModifSousDetail.ModifPASousDet(TOBL: TOB; Pua : double) : boolean;
var TOBO,TOBOD,TOBA : TOB;
		AncPUa,valeurPR,Qte : double;
    zvaleur : T_valeurs;
    gestionHt : boolean;
begin
  result := true;
  TOBO := FindLigneOuvrageinDoc (TOBL);
  if TOBO <> nil then
  begin
  	TOBOD := FindDetailOuvrage (TOBL);
    if TOBOD = nil then exit;
    //
    if IsLignefacture (TOBOD) then
    begin
      PgiError(Traduirememoire('Impossible : une facturation a déjà été effectuée sur cette ligne.'));
      result := false;
      exit;
    end;
    gestionHt := EnHt(TOBO);
  	TFFActure(XX).deduitLaLigne (TOBO); // pour calculer correctement la pièce
    //
    AncPua := TOBOD.GetValue('ANCPA');
    Qte := TOBOD.GetValue('BLO_QTEFACT');
    if (Pua = AncPua) then exit;

    TOBA := FindTOBArtSais (fTOBArticles,TOBOD.GetValue('BLO_ARTICLE'));

    if (TOBOD.Detail.Count > 0) then
    begin
      if TOBOD.getValue('BLO_TYPEARTICLE')='OUV' then
      begin
        ReajusteMontantOuvrage (fTOBARTicles,fTOBpiece,TOBO,TOBOD,AncPua,ValeurPr,Pua,DEV,gestionHt,true,false,true);
        GetValoDetail (TOBOD);
        CalculOuvFromDetail (TOBOD,DEV,Zvaleur);
        TOBOD.Putvalue ('BLO_DPA',Zvaleur[0]);
        TOBOD.Putvalue ('BLO_DPR',Zvaleur[1]);
        TOBOD.putvalue ('GA_PVHT',Zvaleur[2]);
        TOBOD.Putvalue ('BLO_PUHT',Zvaleur[2]);
        TOBOD.Putvalue ('BLO_PUHTBASE',Zvaleur[2]);
        TOBOD.Putvalue ('BLO_PUHTDEV',pivottodevise(Zvaleur[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
        TOBOD.Putvalue ('BLO_PUTTC',Zvaleur[3]);
        TOBOD.Putvalue ('BLO_PUTTCDEV',pivottodevise(Zvaleur[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
        TOBOD.PutValue ('ANCPV', TOBOD.GetValue('BLO_PUHTDEV'));
        TOBOD.PutValue ('ANCPA', TOBOD.GetValue('BLO_DPA'));
        TOBL.PutValue  ('ANCPR', TOBOD.GetValue('BLO_DPR'));
        TOBOD.Putvalue ('BLO_PMAP',Zvaleur[6]);
        TOBOD.Putvalue ('BLO_PMRP',Zvaleur[7]);
        //
        TOBOD.PutValue('BLO_TPSUNITAIRE',Zvaleur[9]);
        TOBOD.PutValue('GA_HEURE',Zvaleur[9]);
        //
        TOBOD.PutValue('BLO_MONTANTFG',Arrondi(Zvaleur[13]*Qte,V_PGI.OkdecP));
        TOBOD.PutValue('BLO_MONTANTFC',Arrondi(Zvaleur[15]*Qte,V_PGI.OkdecP));
        TOBOD.PutValue('BLO_MONTANTFR',Arrondi(Zvaleur[14]*Qte,V_PGI.OkdecP));
        CalculeLigneAcOuvCumul (TOBOD);

      end else if TOBOD.getValue('BLO_TYPEARTICLE')='ARP' then
      begin
        TOBOD.detail[0].putvalue ('GA_PAHT',Pua);
        TOBL.detail[0].putvalue ('BLO_DPA',Pua);
        CalculeLigneAcOuv (TOBOD.detail[0],DEV,true,TOBA);
        TOBOD.detail[0].putvalue ('GA_PVHT',TOBOD.detail[0].GetValue ('BLO_PUHT'));
        CalculMontantHtDevLigOuv (TOBOD.detail[0],DEV);

        CalculOuvFromDetail (TOBOD,DEV,Zvaleur);
        TOBOD.Putvalue ('BLO_DPA',Zvaleur[0]);
        TOBOD.Putvalue ('BLO_DPR',Zvaleur[1]);
        TOBOD.putvalue ('GA_PVHT',Zvaleur[2]);
        TOBOD.Putvalue ('BLO_PUHT',Zvaleur[2]);
        TOBOD.Putvalue ('BLO_PUHTBASE',Zvaleur[2]);
        TOBOD.Putvalue ('BLO_PUHTDEV',pivottodevise(Zvaleur[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
        TOBOD.Putvalue ('BLO_PUTTC',Zvaleur[3]);
        TOBOD.Putvalue ('BLO_PUTTCDEV',pivottodevise(Zvaleur[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
        TOBOD.Putvalue ('BLO_PMAP',Zvaleur[6]);
        TOBOD.Putvalue ('BLO_PMRP',Zvaleur[7]);
        TOBOD.PutValue('BLO_TPSUNITAIRE',Zvaleur[9]);
        TOBOD.PutValue('GA_HEURE',Zvaleur[9]);
        TOBOD.PutValue('BLO_MONTANTFG',Arrondi(Zvaleur[13]*Qte,V_PGI.OkdecP));
        TOBOD.PutValue('BLO_MONTANTFC',Arrondi(Zvaleur[15]*Qte,V_PGI.OkdecP));
        TOBOD.PutValue('BLO_MONTANTFR',Arrondi(Zvaleur[14]*Qte,V_PGI.OkdecP));
        CalculeLigneAcOuvCumul (TOBOD);
        GetValoDetail (TOBOD);
      end;
    end else
    begin
      TOBOD.PutValue('BLO_DPA', Pua);
      CalculeLigneAcOuv (TOBOD,DEV,true,TOBA);
    end;
    TOBOD.PutValue('ANCPA', pua);
    //
    //
    CalculMontantHtDevLigOuv (TOBOD,DEV);
    TOBOD.SetDouble ('GA_PAHT',TOBOD.GetDouble ('BLO_DPA'));
    TOBOD.SetDouble ('GA_PVHT',TOBOD.GetDouble ('BLO_PUHTDEV'));
    TOBOD.PutValue('GA_HEURE',TOBOD.GetDouble ('BLO_TPSUNITAIRE'));
    AfficheLaLigneSousDetail (TOBOD,TOBL,TOBO);
    //
    CalculeLaLigneDoc (TOBO);
    //
  end;
end;

function TModifSousDetail.ModifPVSousDet(TOBL: TOB; Pu : double) : boolean;
var TOBO,TOBOD,TOBA : TOB;
		AncPU,valeurPR,Qte : double;
    zvaleur : T_valeurs;
    gestionHt : boolean;
begin
  result := true;
  TOBO := FindLigneOuvrageinDoc (TOBL);
  if TOBO <> nil then
  begin
  	TOBOD := FindDetailOuvrage (TOBL);
    if TOBOD = nil then exit;
    if IsLignefacture (TOBOD) then
    begin
      PgiError(Traduirememoire('Impossible : une facturation a déjà été effectuée sur cette ligne.'));
      result := false;
      exit;
    end;
    //
    gestionHt := EnHt(TOBO);
  	TFFActure(XX).deduitLaLigne (TOBO); // pour calculer correctement la pièce
    //
    AncPu := TOBOD.GetValue('ANCPV');
    if Pu = AncPu then exit;
    //
    if (TOBOD.Detail.Count > 0) then
    begin
			Qte := TOBOD.GetValue('BLO_QTEFACT');
    	ReajusteMontantOuvrage (fTOBARTicles,fTOBpiece,TOBO,TOBOD,AncPu,ValeurPr,Pu,DEV,gestionHt);
      if GestionHt then
      BEGIN
        TOBOD.PutValue('BLO_PUHTDEV', Pu);
        TOBOD.PutValue('BLO_PUHT', devisetopivotEx(Pu,DEV.taux,DEV.quotite,V_PGI.OkdecP));
        CalculeLigneHTOuv  (TOBOD,TOBPiece,DEV);
      END ELSE
      BEGIN
        TOBOD.PutValue('BLO_PUTTCDEV', Pu);
        CalculeLigneTTCOuv  (TOBOD,TOBPiece,DEV);
      END;
      TOBOD.PutValue('ANCPV', pu);
      GetValoDetail (TOBOD);
      CalculOuvFromDetail (TOBOD,DEV,Zvaleur);
      TOBOD.Putvalue ('BLO_DPA',Zvaleur[0]);
      TOBOD.Putvalue ('BLO_DPR',Zvaleur[1]);
      TOBOD.putvalue ('GA_PVHT',Zvaleur[2]);
      TOBOD.Putvalue ('BLO_PUHT',Zvaleur[2]);
      TOBOD.Putvalue ('BLO_PUHTBASE',Zvaleur[2]);
      TOBOD.Putvalue ('BLO_PUHTDEV',pivottodevise(Zvaleur[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
      TOBOD.Putvalue ('BLO_PUTTC',Zvaleur[3]);
      TOBOD.Putvalue ('BLO_PUTTCDEV',pivottodevise(Zvaleur[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
      TOBOD.PutValue ('ANCPV', TOBOD.GetValue('BLO_PUHTDEV'));
      TOBOD.PutValue ('ANCPA', TOBOD.GetValue('BLO_DPA'));
      TOBL.PutValue  ('ANCPR', TOBOD.GetValue('BLO_DPR'));
      TOBOD.Putvalue ('BLO_PMAP',Zvaleur[6]);
      TOBOD.Putvalue ('BLO_PMRP',Zvaleur[7]);
      //
      TOBOD.PutValue('BLO_TPSUNITAIRE',Zvaleur[9]);
      TOBOD.PutValue('GA_HEURE',Zvaleur[9]);
      //
      TOBOD.PutValue('BLO_MONTANTFG',Arrondi(Zvaleur[13]*Qte,V_PGI.OkdecP));
      TOBOD.PutValue('BLO_MONTANTFC',Arrondi(Zvaleur[15]*Qte,V_PGI.OkdecP));
      TOBOD.PutValue('BLO_MONTANTFR',Arrondi(Zvaleur[14]*Qte,V_PGI.OkdecP));
      CalculeLigneAcOuvCumul (TOBOD);
    end else
    begin
      if GestionHt then
      BEGIN
        TOBOD.PutValue('BLO_PUHTDEV', Pu);
        TOBOD.PutValue('BLO_PUHT', devisetopivotEx(Pu,DEV.taux,DEV.quotite,V_PGI.OkdecP));
        CalculeLigneHTOuv  (TOBOD,TOBPiece,DEV);
      END ELSE
      BEGIN
        TOBOD.PutValue('BLO_PUTTCDEV', Pu);
        CalculeLigneTTCOuv  (TOBOD,TOBPiece,DEV);
      END;
      TOBOD.PutValue('ANCPV', pu);
      TOBA := FindTOBArtSais (fTOBArticles,TOBOD.GetValue('BLO_ARTICLE'));
      CalculeLigneAcOuv (TOBOD,DEV,TOBL.GetValue('GLC_NONAPPLICFRAIS')='X',TOBA);
    end;
    //
    //
    //
    CalculMontantHtDevLigOuv (TOBOD,DEV);
    AfficheLaLigneSousDetail (TOBOD,TOBL,TOBO);
    //
    CalculeLaLigneDoc (TOBO);
    //
  end;
end;

function TModifSousDetail.ModifQteSaisSousDet (TOBL: TOB; Ichps : TChpModif ; Valeur : double; DEV : RDevise) : boolean;
var TOBO,TOBOD,TOBA : TOB;
		QtePere,QteMaj,QteDuDetail : double;
    zvaleur : T_valeurs;
begin
	result := true;
  TOBO := FindLigneOuvrageinDoc (TOBL);
  if TOBO <> nil then
  begin
  	TOBOD := FindDetailOuvrage (TOBL);
    if TOBOD = nil then exit;
    //
  	QtePere := TOBO.getValue('GL_QTEFACT');
    QteDuDetail := TOBOD.getValue('BLO_QTEDUDETAIL'); if QteDuDetail = 0 then QteDuDetail := 1;
    if IChps = TchpQte then
    begin
      if TFFacture(XX).AffSousDetailUnitaire then
      begin
        TOBOD.putvalue('BLO_QTESAIS',Valeur);
      end else
      begin
        Qtemaj := Arrondi(Valeur/QtePere * QteDuDetail,V_PGI.okdecQ);
        TOBOD.putvalue('BLO_QTESAIS',QteMaj);
      end;
    end else if IChps = TChpPerte then
    begin
      TOBOD.putvalue('BLO_PERTE',Valeur);
    end else if IChps = TChpRendement then
    begin
      TOBOD.putvalue('BLO_RENDEMENT',Valeur);
    end;
    if POS(TOBOD.GetString('BNP_TYPERESSOURCE'),'SAL;INT;ST') > 0 then
    begin
      // prise en compte du rendement
      if TOBOD.GetDouble('BLO_RENDEMENT') <> 0 then TOBOD.SetDouble('BLO_QTEFACT',TOBOD.GetDouble ('BLO_QTESAIS')/TOBOD.GetDouble('BLO_RENDEMENT'))
                                               else TOBOD.SetDouble('BLO_QTEFACT',TOBOD.GetDouble ('BLO_QTESAIS'));
    end else
    begin
      // prise en compte du coef de perte
      if TOBOD.GetDouble('BLO_PERTE') <> 0 then TOBOD.SetDouble('BLO_QTEFACT',TOBOD.GetDouble ('BLO_QTESAIS')*TOBOD.GetDouble('BLO_PERTE'))
                                           else TOBOD.SetDouble('BLO_QTEFACT',TOBOD.GetDouble ('BLO_QTESAIS'));
    end;
    if TOBOD.GetValue('GA_TYPEARTICLE')='PRE' then
    begin
      TOBOD.PutValue('GA_HEURE',TOBOD.GetValue('BLO_QTEFACT'));
    end;
    //
    if TOBOD.detail.count > 0 then
    begin
      GetValoDetail (TOBOD);
      CalculOuvFromDetail (TOBOD,DEV,Zvaleur);
      TOBOD.Putvalue ('BLO_DPA',Zvaleur[0]);
      TOBOD.Putvalue ('BLO_DPR',Zvaleur[1]);
      TOBOD.putvalue ('GA_PVHT',Zvaleur[2]);
      TOBOD.Putvalue ('BLO_PUHT',Zvaleur[2]);
      TOBOD.Putvalue ('BLO_PUHTBASE',Zvaleur[2]);
      TOBOD.Putvalue ('BLO_PUHTDEV',pivottodevise(Zvaleur[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
      TOBOD.Putvalue ('BLO_PUTTC',Zvaleur[3]);
      TOBOD.Putvalue ('BLO_PUTTCDEV',pivottodevise(Zvaleur[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
      TOBOD.PutValue('ANCPV', TOBOD.GetValue('BLO_PUHTDEV'));
      TOBOD.PutValue('ANCPA', TOBOD.GetValue('BLO_DPA'));
      TOBOD.PutValue('ANCPR', TOBOD.GetValue('BLO_DPR'));
      TOBOD.Putvalue ('BLO_PMAP',Zvaleur[6]);
      TOBOD.Putvalue ('BLO_PMRP',Zvaleur[7]);
      //
      TOBOD.PutValue('BLO_TPSUNITAIRE',Zvaleur[9]);
      TOBOD.PutValue('GA_HEURE',Zvaleur[9]);
      //
      TOBOD.PutValue('BLO_MONTANTFG',Arrondi(Zvaleur[13]*QteMaj,V_PGI.OkdecP));
      TOBOD.PutValue('BLO_MONTANTFC',Arrondi(Zvaleur[15]*QteMaj,V_PGI.OkdecP));
      TOBOD.PutValue('BLO_MONTANTFR',Arrondi(Zvaleur[14]*QteMaj,V_PGI.OkdecP));
      CalculeLigneAcOuvCumul (TOBOD);
    end else
    begin
      TOBA := FindTOBArtSais (fTOBArticles,TOBOD.GetValue('BLO_ARTICLE'));
      CalculeLigneAcOuv (TOBOD,DEV,TOBL.GetValue('GLC_NONAPPLICFRAIS')='X',TOBA);
    end;
    //
    CalculMontantHtDevLigOuv (TOBOD,DEV);
    AfficheLaLigneSousDetail (TOBOD,TOBL,TOBO);
    //
    CalculeLaLigneDoc (TOBO);
    //
  end;
end;

function TModifSousDetail.ModifQteSousDet(TOBL: TOB; Qte : double; DEV : RDevise) : boolean;
var TOBO,TOBOD,TOBA : TOB;
		QtePere,QteMaj,QteDuDetail : double;
    zvaleur : T_valeurs;
    Isfacturation : Boolean;
begin
	result := true;
  Isfacturation := (Pos(TOBL.GetValue('GL_NATUREPIECEG'),'FBT;FBP')>0);
  TOBO := FindLigneOuvrageinDoc (TOBL);
  if TOBO <> nil then
  begin
  	TOBOD := FindDetailOuvrage (TOBL);
    if TOBOD = nil then exit;

    if IsLignefacture (TOBOD) then
    begin
      PgiError(Traduirememoire('Impossible : une facturation a déjà été effectuée sur cette ligne.'));
      result := false;
      exit;
    end;

  	TFFActure(XX).deduitLaLigne (TOBO); // pour calculer correctement la pièce
    //
  	QtePere := TOBO.getValue('GL_QTEFACT');
    QteDuDetail := TOBOD.getValue('BLO_QTEDUDETAIL'); if QteDuDetail = 0 then QteDuDetail := 1;
    if TFFacture(XX).AffSousDetailUnitaire then
    begin
      TOBOD.putvalue('BLO_QTEFACT',Qte);
    end else
    begin
    Qtemaj := Arrondi(Qte/QtePere * QteDuDetail,V_PGI.okdecQ);
    TOBOD.putvalue('BLO_QTEFACT',QteMaj);
    end;
    if TOBOD.GetValue('GA_TYPEARTICLE')='PRE' then
    begin
      TOBOD.PutValue('GA_HEURE',Qtemaj);
    end;
		//
    if TOBOD.detail.count > 0 then
    begin
      GetValoDetail (TOBOD);
      CalculOuvFromDetail (TOBOD,DEV,Zvaleur);
      TOBOD.Putvalue ('BLO_DPA',Zvaleur[0]);
      TOBOD.Putvalue ('BLO_DPR',Zvaleur[1]);
      TOBOD.putvalue ('GA_PVHT',Zvaleur[2]);
      TOBOD.Putvalue ('BLO_PUHT',Zvaleur[2]);
      TOBOD.Putvalue ('BLO_PUHTBASE',Zvaleur[2]);
      TOBOD.Putvalue ('BLO_PUHTDEV',pivottodevise(Zvaleur[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
      TOBOD.Putvalue ('BLO_PUTTC',Zvaleur[3]);
      TOBOD.Putvalue ('BLO_PUTTCDEV',pivottodevise(Zvaleur[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
      TOBOD.PutValue('ANCPV', TOBOD.GetValue('BLO_PUHTDEV'));
      TOBOD.PutValue('ANCPA', TOBOD.GetValue('BLO_DPA'));
      TOBOD.PutValue('ANCPR', TOBOD.GetValue('BLO_DPR'));
      TOBOD.Putvalue ('BLO_PMAP',Zvaleur[6]);
      TOBOD.Putvalue ('BLO_PMRP',Zvaleur[7]);
      //
      TOBOD.PutValue('BLO_TPSUNITAIRE',Zvaleur[9]);
      TOBOD.PutValue('GA_HEURE',Zvaleur[9]);
      //
      TOBOD.PutValue('BLO_MONTANTFG',Arrondi(Zvaleur[13]*QteMaj,V_PGI.OkdecP));
      TOBOD.PutValue('BLO_MONTANTFC',Arrondi(Zvaleur[15]*QteMaj,V_PGI.OkdecP));
      TOBOD.PutValue('BLO_MONTANTFR',Arrondi(Zvaleur[14]*QteMaj,V_PGI.OkdecP));
      CalculeLigneAcOuvCumul (TOBOD);
    end else
    begin
      TOBA := FindTOBArtSais (fTOBArticles,TOBOD.GetValue('BLO_ARTICLE'));
      CalculeLigneAcOuv (TOBOD,DEV,TOBL.GetValue('GLC_NONAPPLICFRAIS')='X',TOBA);
    end;
    //
    CalculMontantHtDevLigOuv (TOBOD,DEV);
    AfficheLaLigneSousDetail (TOBOD,TOBL,TOBO);
    //
    CalculeLaLigneDoc (TOBO);
    //
  end;
end;

procedure TModifSousDetail.ModifUniteSousDet(TOBL: TOB; Unite : string);
var TOBOD : TOB;
begin
  TOBOD := FindDetailOuvrage (TOBL);
  if TOBOD = nil then exit;
  TOBOD.putValue('BLO_QUALIFQTEVTE',Unite);
  TOBL.putValue('GL_QUALIFQTEVTE',Unite);
  TFFActure(XX).AfficheLaLigne (TOBL.GetIndex+1);
end;

procedure TModifSousDetail.RechercheCodeArticle(Acol,Arow : integer; TOBL: TOB;CodeArticle: string);
var TypeArticle,Libelle,Article : string;
		TOBOD,TOBLO: TOB;
begin
  TOBLO := FindLigneOuvrageinDoc (TOBL);
  if TOBLO <> nil then
  begin
    TOBOD := FindDetailOuvrage (TOBL);
    if TOBOD = nil then exit;
    if GetCodeArticle(TypeArticle,Article,CodeArticle,Libelle,Acol,Arow) then
    begin
  		TFFActure(XX).deduitLaLigne (TOBLO); // pour calculer correctement la pièce
    	SetLigneOuvrage (TOBLO,TOBL,TOBOD,Article,CodeArticle,Libelle);
      AfficheLaLigneSousDetail (TOBOD,TOBL,TOBLO);
      CalculeLaLigneDoc (TOBLO);
      TFFACTURE(XX).StCellCur := TFFActure(XX).GS.Cells[Acol,Arow];
    end;
  end;
end;

procedure TModifSousDetail.ReInitligneOuv(TOBOD : TOB);
begin
  TOBOD.PutValue('BLO_NATURETRAVAIL','');
  TOBOD.PutValue('BLO_FOURNISSEUR','');
  TOBOD.PutValue('BLO_CODEARTICLE', '');
  TOBOD.PutValue('BLO_ARTICLE', '');
  TOBOD.PutValue('BLO_LIBELLE', '');
  TOBOD.putValue('BLO_COMPOSE','');
  TOBOD.putValue('BLO_TYPEARTICLE','');
end;

procedure TModifSousDetail.RenumeroteLesFilles(TOBLN : TOB; Niveau : integer);
var LigPere1,LigPere2,LigPere3,LigPere4,LigPere5 : integer;
    Indice,CeluiLa : integer;
    NiveauANumeroter : integer;
    TOBLLN : TOB;
begin
  NiveauANumeroter := Niveau +1;
  if Niveau = 1 then
  begin
    LigPere1 := TOBLN.GetValue('BLO_N1');
    LigPere2 := 0;
    LigPere3 := 0;
    LigPere4 := 0;
    LigPere5 := 0;
  end else if Niveau = 2 then
  begin
    LigPere1 := TOBLN.GetValue('BLO_N1');
    LigPere2 := TOBLN.GetValue('BLO_N2');
    LigPere3 := 0;
    LigPere4 := 0;
    LigPere5 := 0;
  end else if Niveau = 3 then
  begin
    LigPere1 := TOBLN.GetValue('BLO_N1');
    LigPere2 := TOBLN.GetValue('BLO_N2');
    LigPere3 := TOBLN.GetValue('BLO_N3');
    LigPere4 := 0;
    LigPere5 := 0;
  end else if Niveau = 4 then
  begin
    LigPere1 := TOBLN.GetValue('BLO_N1');
    LigPere2 := TOBLN.GetValue('BLO_N2');
    LigPere3 := TOBLN.GetValue('BLO_N3');
    LigPere4 := TOBLN.GetValue('BLO_N4');
    LigPere5 := 0;
  end;
  for Indice := 0 to TOBLN.detail.count -1 do
  begin
    TOBLLN := TOBLN.detail[Indice];
    if NiveauANumeroter = 2 then BEGIN Inc(LigPere2); CeluiLa := LigPere2; END else
    if NiveauANumeroter = 3 then BEGIN Inc(LigPere3) ; CeluiLa := LigPere3; END else
    if NiveauANumeroter = 4 then BEGIN Inc(LigPere4) ; CeluiLa := LigPere4; END else
    if NiveauANumeroter = 5 then BEGIN Inc(LigPere5); ; CeluiLa := LigPere5; END;
    SetNumerotation (TOBLLN,TOBLN,NiveauANumeroter, CeluiLa);
    if TOBLLN.detail.count > 0 then
    begin
      RenumeroteLesFilles (TOBLLN,NiveauANumeroter);
    end;
  end;
end;

procedure TModifSousDetail.SetNumerotation(TOBL,TOBMere: TOb; Niveau : integer ;Arow: integer);
begin
  if Niveau=1 then
  begin
    // provenance de la ligne de document
    TOBL.putValue('BLO_N1',Arow);
  end else
  begin
    // provenance d'un detail d'ouvrage
    if Niveau = 2 then
    begin
      TOBL.PutValue('BLO_N1',TOBMere.getValue('BLO_N1'));
      TOBL.PutValue('BLO_N2',Arow);
    end else if Niveau = 3 then
    begin
      TOBL.PutValue('BLO_N1',TOBMere.getValue('BLO_N1'));
      TOBL.PutValue('BLO_N2',TOBMere.getValue('BLO_N2'));
      TOBL.PutValue('BLO_N3',Arow);
    end else if Niveau = 4 then
    begin
      TOBL.PutValue('BLO_N1',TOBMere.getValue('BLO_N1'));
      TOBL.PutValue('BLO_N2',TOBMere.getValue('BLO_N2'));
      TOBL.PutValue('BLO_N3',TOBMere.getValue('BLO_N3'));
      TOBL.PutValue('BLO_N4',Arow);
    end else if Niveau = 5 then
    begin
      TOBL.PutValue('BLO_N1',TOBMere.getValue('BLO_N1'));
      TOBL.PutValue('BLO_N2',TOBMere.getValue('BLO_N2'));
      TOBL.PutValue('BLO_N3',TOBMere.getValue('BLO_N3'));
      TOBL.PutValue('BLO_N4',TOBMere.getValue('BLO_N4'));
      TOBL.PutValue('BLO_N5',Arow);
    end;
  end;
end;


procedure TModifSousDetail.SetLigneOuvrage (TOBLO,TOBL,TOBOD: TOB;Article,CodeArticle: string; Libelle : string);
var TOBART,TOBOO : TOB;
		Tresult : T_valeurs;
    QteDuDetail,QteSit,MtSit : double;
    IndiceNomen : integer;
begin
  IndiceNomen := TOBL.getValue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
  TOBOO := fTOBOuvrages.Detail[IndiceNomen-1];
  QteDuDetail := TOBOO.detail[0].getValue('BLO_QTEDUDETAIL');
  TOBART := FindTOBArtSais(fTOBArticles ,Article);
  if TOBArt = nil then exit;
  if (Valeur(TOBLO.getValue('GLC_NATURETRAVAIL')) <> 0) and (Valeur(TOBLO.getValue('GLC_NATURETRAVAIL')) < 10) then
  begin
    TOBOD.PutValue('BLO_NATURETRAVAIL',TOBLO.getValue('GLC_NATURETRAVAIL'));
    TOBOD.PutValue('BLO_FOURNISSEUR',TOBLO.getValue('GL_FOURNISSEUR'));
  end;
  TOBOD.PutValue('LIBELLEFOU',TOBLO.getValue('LIBELLEFOU'));
  TOBOD.Putvalue ('BLO_QTEDUDETAIL',qtedudetail);
  TOBOD.PutValue('BLO_NATUREPIECEG',TOBLO.getValue('GL_NATUREPIECEG'));
  TOBOD.PutValue('BLO_SOUCHE',TOBLO.getValue('GL_SOUCHE'));
  TOBOD.PutValue('BLO_NUMERO',TOBLO.getValue('GL_NUMERO'));
  TOBOD.PutValue('BLO_INDICEG',TOBLO.getValue('GL_INDICEG'));
  TOBOD.PutValue('BLO_DOMAINE',TOBLO.getValue('GL_DOMAINE'));
  TOBOD.PutValue('BLO_CODEARTICLE', TOBArt.getValue('GA_CODEARTICLE'));
  TOBOD.PutValue('BLO_ARTICLE', Article);
  TOBOD.PutValue('BLO_QTEFACT', 1.0);
  TOBOD.PutValue('BNP_TYPERESSOURCE', TOBArt.getValue('BNP_TYPERESSOURCE'));
  TOBOD.PutValue('BLO_QTESAIS', TOBOD.GetDouble('BLO_QTEFACT'));
  //
  if GetParamSocSecur('SO_BTGESTQTEAVANC',false)  and (ISDocumentChiffrage(TOBLO)) then
  begin
    if (POS(TOBOD.GetString('BNP_TYPERESSOURCE'),'SAL;INT;ST') > 0)  then
    begin
      TOBOD.PutValue('BLO_QUALIFHEURE',     TOBArt.GetValue('GA_QUALIFHEURE'));
      TOBOD.PutValue('BLO_RENDEMENT',       TOBArt.GetValue('GA_COEFPROD'));
      TOBOD.PutValue('BLO_PERTE',0);
    end else
    begin
      TOBOD.SetString('BLO_QUALIFHEURE','');
      TOBOD.PutValue('BLO_PERTE',           TOBArt.GetValue('GA_PERTEPROP'));
      TOBOD.PutValue('BLO_RENDEMENT',0);
    end;
    TOBOD.SetDouble('BLO_QTEFACT',CalculeQteFact (TOBOD));
  end;

  TOBOD.PutValue('BLO_COMPOSE',TOBLO.getValue('GL_ARTICLE')) ;
  TOBOD.Putvalue('BLO_NOMENCLATURE',TOBArt.getValue('GA_CODEARTICLE')) ;
  if Libelle = '' then
  begin
    TOBOD.PutValue('BLO_LIBELLE', TOBArt.GetValue('GA_LIBELLE'));
  end else
  BEGIN
    TOBOD.PutValue('BLO_LIBELLE', Libelle);
  END;
  TOBOD.PutValue('BLO_ORDRECOMPO',1);
  RenseigneTOBOuv(fTOBpiece,TOBOD,TOBLO,TOBART,Libelle,DEV);
  if isOuvrage(TOBOD) then
  begin
    TOBOD.putValue('BLO_COMPOSE',Article);
    if not ChargeOuvrage (TOBLO,CodeArticle, TOBOD , 2,TOBOD.GetValue('BLO_NUMORDRE'),Tresult) then
    begin
    	ReInitligneOuv(TOBOD);
      exit;
    end;
    RenumeroteLesFilles (TOBOD,1);
    TOBOD.Putvalue ('BLO_DPA',Tresult[0]);
    TOBOD.Putvalue ('BLO_DPR',Tresult[1]);
    TOBOD.Putvalue ('BLO_PUHT',Tresult[2]);
    TOBOD.Putvalue ('BLO_PUHTBASE',Tresult[2]);
    TOBOD.Putvalue ('BLO_PUHTDEV',pivottodevise(Tresult[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
    TOBOD.Putvalue ('BLO_PUTTC',Tresult[3]);
    TOBOD.Putvalue ('BLO_PUTTCDEV',pivottodevise(Tresult[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
    TOBOD.Putvalue ('BLO_PMAP',Tresult[6]);
    TOBOD.PutValue('ANCPV', TOBOD.getValue('BLO_PUHTDEV'));
    TOBOD.PutValue('ANCPA', TOBOD.getValue('BLO_DPA'));
    TOBOD.PutValue('ANCPR', TOBOD.getValue('BLO_DPR'));
    TOBOD.Putvalue ('BLO_PMRP',Tresult[7]);
  //  TOBOD.Putvalue ('BLO_QTEDUDETAIL',qtedudetail);
    TOBOD.putValue('BLO_TPSUNITAIRE',TResult[9]);
    TOBOD.putValue('GA_HEURE',Tresult[9]);
    TOBOD.putValue('GA_PAHT',Tresult[0]);
    TOBOD.PutValue('BLO_TPSUNITAIRE',Tresult[9]);
    TOBOD.PutValue('GA_HEURE',Tresult[9]);
    TOBOD.PutValue('BLO_MONTANTFG',Arrondi(Tresult[13]*TOBOD.getValue('BLO_QTEFACT'),V_PGI.OkdecP));
    TOBOD.PutValue('BLO_MONTANTFC',Arrondi(Tresult[15]*TOBOD.getValue('BLO_QTEFACT'),V_PGI.OkdecP));
    TOBOD.PutValue('BLO_MONTANTFR',Arrondi(Tresult[14]*TOBOD.getValue('BLO_QTEFACT'),V_PGI.OkdecP));
    CalculeLigneAcOuvCumul (TOBOD);
    CalculMontantHtDevLigOuv (TOBOD,DEV);
    GetValoDetail (TOBOD);
  end;
  if (TOBOD.nomtable = 'LIGNEOUV') and
     (TOBLO.GetString('GL_PIECEPRECEDENTE')='') and
     (TOBOD.FieldExists('BLF_QTESITUATION')) then
  begin
    QteSit := arrondi(TOBOD.getValue('BLO_QTEFACT')*TOBLO.getvalue('GL_QTEFACT') /TOBOD.getValue('BLO_QTEDUDETAIL'),V_PGI.okdecQ);
    MtSit := arrondi(TOBOD.getValue('BLO_QTEFACT')*TOBLO.getvalue('GL_QTEFACT') /TOBOD.getValue('BLO_QTEDUDETAIL')*TOBOD.getValue('BLO_PUHTDEV'),DEV.Decimale);
    TOBOD.PutValue('BLF_QTESITUATION', QteSit); { NEWPIECE }
    TOBOD.PutValue('BLF_MTSITUATION',MtSit);
    TOBOD.PutValue('BLF_POURCENTAVANC',100);
  end;
end;

function  TModifSousDetail.VerifNomen(TobArt : TOB ; Arow : integer; Var retour : string; var Libelle : string ) : boolean;
var Q : TQuery;
    st1 : string;
    nb : integer;
    TypeArticle : string;
begin
  result := true;
  st1 := '';
  nb := 0;
  TYpeArticle := TOBArt.GetValue('GA_TYPEARTICLE');
  if (TypeArticle <> 'NOM') and
  	 ((TypeArticle <> 'OUV') and (TypeArticle <> 'ARP')) then
  BEGIN
  	exit;
  END;

  Q := OpenSQL('Select Count(GNE_NOMENCLATURE) from NOMENENT Where GNE_ARTICLE="' +
  		 TOBArt.GetValue('GA_ARTICLE') + '"', True,-1, '', True);
  if not(Q.Eof) then nb := Q.Fields[0].AsInteger;
  Ferme(Q);

  if nb > 1 then
  BEGIN
    st1:=Choisir('Choix d''une nomenclature','NOMENENT','GNE_LIBELLE',
                 'GNE_NOMENCLATURE','GNE_ARTICLE="'+ TOBArt.GetValue('GA_ARTICLE')+'"','');
    if st1 = '' then
    begin
    	result := false;
      exit;
    end;
  END else if nb = 1 then
  BEGIN
  	st1 := TOBArt.GetValue('GA_CODEARTICLE');
  END else result := false;

  if st1 <> '' then
  begin
    Q := OpenSQL('Select GNE_LIBELLE from NOMENENT Where GNE_NOMENCLATURE="' + st1 + '"', True,-1, '', True);
    Libelle := Q.FindField('GNE_Libelle').AsString;
    TOBART.putValue('REFARTSAISIE',TOBArt.GetValue('GA_CODEARTICLE'));
    Ferme(Q);
  end;
  Retour := st1;
end;

function TModifSousDetail.ChargeOuvrage(TOBL: TOB; CodeNomen : string ; TobOuvrage : TOB ; Niv,OrdreCompo : integer ;var  Tresult: T_valeurs):boolean;
var
  i_ind1, i_ind2,IndPou : integer;
  Select : string;
  Q,QQ : TQuery;
  TOBDetOuv, TobLigne, TobFille : TOB;
  TobArt : TOB;
  // Modif BTP
  TOBLigNom : TOB;
  QteDudetail,QTe,QTedupv,MontantInterm,MontantPr : Double;
  Valloc : T_Valeurs;
  OKGestQteAvance : boolean;
begin
  OKGestQteAvance := (GetParamSocSecur('SO_BTGESTQTEAVANC',false))  and (ISDocumentChiffrage(TOBL)) ;

  if niv > 5 then
  BEGIN
    PGIBox ('Nombre de niveaux trop important : 5 Maximum','Ouvrages');
    Result := false;
    Exit;
  END;
  result := true;
  QteDuDetail := 1;
  indpou := -1;
  //  CodeNomen obligatoire
  if CodeNomen <> '' then
  begin
    InitTableau (Tresult);
    TobOuvrage.PutValue('SOUSNOMEN','X');   //En cas de récursivité
    TobLigne := TOB.Create('NOMENLIG', nil, -1);
    //  Chargement du niveau courant de nomenclature
    Select := 'Select NOMENLIG.*,GNE_QTEDUDETAIL,GNE_LIBELLE from NOMENLIG LEFT OUTER JOIN NOMENENT ON GNE_NOMENCLATURE = GNL_NOMENCLATURE where GNL_NOMENCLATURE = "' + CodeNomen + '"';
    Q := OpenSQL(Select, true,-1, '', True);
    if not Q.EOF then
    begin
      TobLigne.LoadDetailDB('DETAILNOMEN', '', '', Q, False);
      ferme (Q);
      for i_ind1:=0 to TobLigne.Detail.count -1 do
      begin
        TOBLIGNom := TOBLIgne.detail[i_ind1];
        QteDuDetail := TOBLigNom.getValue('GNE_QTEDUDETAIL');
        if QteDuDetail = 0 then QteDuDetail := 1;
        TOBDetOuv:=TOB.Create('LIGNEOUV',TobOuvrage,-1);
        InsertionChampSupOuv (TOBDetOuv,false);
        //
        fTOBPiece.putValue('GP_UNIQUEBLO',fTOBPiece.getValue('GP_UNIQUEBLO')+1);
        TOBDetOuv.putValue('BLO_UNIQUEBLO',fTOBPiece.getValue('GP_UNIQUEBLO'));
        //
        // prise en compte document d'origine
        TOBDetOuv.PutValue('BLO_NATUREPIECEG',TOBL.getValue('GL_NATUREPIECEG'));
        TOBDetOuv.PutValue('BLO_SOUCHE',TOBL.getValue('GL_SOUCHE'));
        TOBDetOuv.PutValue('BLO_NUMERO',TOBL.getValue('GL_NUMERO'));
        TOBDetOuv.PutValue('BLO_INDICEG',TOBL.getValue('GL_INDICEG'));
        // ------
        TOBDetOuv.PutValue('BLO_COMPOSE',TobOuvrage.GetValue('BLO_ARTICLE'));
        TOBDetOuv.PutValue('BLO_NIVEAU',Niv);
        TOBDetOuv.PutValue('BLO_ORDRECOMPO',OrdreCompo);

        TOBDetOuv.PutValue('BLO_NUMORDRE',i_ind1+1);
        TOBDetOuv.PutValue('CODENOMEN',TobLigNom.GetValue('GNL_SOUSNOMEN'));
        TOBDETOuv.putvalue('BLO_QTEDUDETAIL',TOBLigNom.getvalue('GNE_QTEDUDETAIL'));
        TOBDetOuv.PutValue('BLO_ARTICLE', TOBLigNom.GetValue('GNL_ARTICLE'));

        //
    		TOBART := FindTOBArtSais (fTOBArticles,TobLigNom.GetValue('GNL_ARTICLE'));
        if TOBARt = nil then
        begin
          QQ := OPenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                        'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                        'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE GA_ARTICLE ="'+
                          TobLigNom.GetValue('GNL_ARTICLE')+'"',True,-1, '', True);
          TOBArt:= TOB.create ('ARTICLE',fTOBarticles,-1);
          TOBArt.selectDB ('',QQ);
          ferme (QQ);
        end;
        InitChampsSupArticle (TOBART);
        //
        TOBDetOuv.PutValue('BLO_QTEFACT', TOBLigNom.GetValue('GNL_QTE'));
        if (OKGestQteAvance) then
        begin
          TOBDetOuv.SetSTring('BLO_QUALIFHEURE',TOBLigNom.GetString('GNL_QUALIFUNITEACH'));
          TOBDetOuv.SetDouble('BLO_RENDEMENT',  TOBLigNom.GetDouble('GNL_RENDEMENT'));
          TOBDetOuv.SetDouble('BLO_PERTE', TOBLigNom.GetDouble('GNL_PERTE'));
          TOBDetOuv.SetDouble('BLO_QTESAIS',TOBLigNom.GetDouble('GNL_QTESAIS'));
        end else
        begin
          TOBDetOuv.SetDouble('BLO_QTESAIS',TOBDetOuv.GetDouble('BLO_QTEFACT'));
        end;

        RenseigneTOBOuv(fTOBpiece,TOBDetOuv,TOBL,TobArt,TOBArt.getValue('GA_LIBELLE'),DEV);
        TOBDetOuv.PutValue('BLO_CODEARTICLE',TobLigNom.GetValue('GNL_CODEARTICLE'));
        // Modif BTP
        TOBDetOuv.PutValue('BLO_REFARTSAISIE',TobLigNom.GetValue('GNL_CODEARTICLE'));
        // ---
        TOBDetOuv.PutValue('BLO_LIBELLE',TOBLigNom.GetValue('GNL_LIBELLE'));
        if TobDetOuv.GetValue('CODENOMEN') <> '' then
        begin
          if not ChargeOuvrage(TOBL,TobDetOuv.GetValue('CODENOMEN'), TobDetOuv , Niv+1,TobDetOuv.GetValue('BLO_NUMORDRE'),valloc) then
          begin
            TOBDetOuv.free;
            TOBligne.free;
            Result := false;
            Exit;
          end;
          TOBDetOuv.Putvalue ('BLO_DPA',valloc[0]);
          TOBDetOuv.Putvalue ('BLO_DPR',valloc[1]);
          TOBDetOuv.Putvalue ('BLO_PUHT',Valloc[2]);
          TOBDetOuv.Putvalue ('BLO_PUHTBASE',Valloc[2]);
          TOBDetOuv.Putvalue ('BLO_PUHTDEV',pivottodevise(Valloc[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
          TOBDetOuv.Putvalue ('BLO_PUTTC',Valloc[3]);
          TOBDetOuv.Putvalue ('BLO_PUTTCDEV',pivottodevise(Valloc[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
          TOBDetOuv.Putvalue ('BLO_PMAP',Valloc[6]);
          TOBDetOuv.Putvalue ('BLO_PMRP',VAlloc[7]);
          TOBDetOuv.Putvalue ('BLO_TPSUNITAIRE',VAlloc[9]);
          TOBDetOuv.PutValue('BLO_MONTANTFG',Arrondi(ValLoc[13]*TOBDetOuv.getValue('BLO_QTEFACT'),V_PGI.OkdecP));
          TOBDetOuv.PutValue('BLO_MONTANTFC',Arrondi(ValLoc[15]*TOBDetOuv.getValue('BLO_QTEFACT'),V_PGI.OkdecP));
          TOBDetOuv.PutValue('BLO_MONTANTFR',Arrondi(ValLoc[14]*TOBDetOuv.getValue('BLO_QTEFACT'),V_PGI.OkdecP));
          CalculeLigneAcOuvCumul (TOBDETOUV);
          CalculMontantHtDevLigOuv (TOBDetOuv,DEV);
        end;
        TOBDetOuv.PutValue('ANCPV',TOBDetOuv.Getvalue('BLO_PUHTDEV') );
        TOBDetOuv.PutValue('ANCPA',TOBDetOuv.Getvalue('BLO_DPA') );
        TOBDetOuv.PutValue('ANCPR',TOBDetOuv.Getvalue('BLO_DPR') );
        if TOBLigNom.GetValue('GNL_PRIXFIXE') <> 0 then
        begin
          TOBDetOuv.PutValue('BLO_PUHT', TOBLigNom.GetValue('GNL_PRIXFIXE'));
          TOBDetOuv.PutValue('BLO_PUHTDEV', pivottodevise(TOBLigNom.GetValue('GNL_PRIXFIXE'),DEV.Taux,DEV.quotite,V_PGI.OkdecP));
        end;
        Qte := TOBDetOuv.getValue('BLO_QTEFACT');
        QteDuPv := TOBDetOuv.getValue('BLO_PRIXPOURQTE'); if QteDuPv = 0 then QteDuPv := 1;
        CalculeLigneAcOuv (TOBDETOUV,DEV,true,TOBART);
        CalculMontantHtDevLigOuv (TOBDetOuv,DEV);
        if TobDetOuv.Getvalue('BLO_TYPEARTICLE') <> 'POU' then
        begin
          Tresult[0] := Tresult [0] + arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_DPA')),V_PGI.okdecP);
          Tresult[1] := Tresult [1] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_DPR')),V_PGI.okdecP);
          Tresult[2] := Tresult [2] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PUHT')),V_PGI.okdecP);
          Tresult[3] := Tresult [3] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PUTTC')),V_PGI.okdecP);
          Tresult[6] := Tresult [6] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PMAP')),V_PGI.okdecP);
          Tresult[7] := Tresult [7] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PMRP')),V_PGI.okdecP);
          Tresult[9] := Tresult [9] +(Qte * TOBDetOuv.GetValue ('BLO_TPSUNITAIRE'));
          Tresult[13] := Tresult [13] +TOBDetOuv.GetValue('BLO_MONTANTFG');
          Tresult[14] := Tresult [14] +TOBDetOuv.GetValue('BLO_MONTANTFR');
          Tresult[15] := Tresult [15] +TOBDetOuv.GetValue('BLO_MONTANTFC');
        end else
        begin
          indpou := 1;
        end;
      end;
    end else Ferme(Q);
    TOBLigne.free;
    FormatageTableau (Tresult,V_PGI.OkdecP);
    // calcule des eventuels pourcentages
    if indpou >= 0 then
    begin
      for I_ind1 := 0 to TOBOuvrage.detail.count -1 do
      begin
        TOBLIGNom := TOBOuvrage.detail[i_ind1];
        if TobLigNom.Getvalue('BLO_TYPEARTICLE') <> 'POU' then continue;
        TOBLIGNom.Putvalue ('BLO_DPA',Tresult[0]);
        TOBLIGNom.Putvalue ('BLO_DPR',Tresult[1]);
        TOBLIGNom.Putvalue ('BLO_PUHT',Tresult[2]);
        TOBLIGNom.Putvalue ('BLO_PUHTBASE',Tresult[2]);
        TOBLIGNom.Putvalue ('BLO_PUHTDEV',pivottodevise(Tresult[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
        TOBLIGNom.Putvalue ('BLO_PUTTC',Tresult[3]);
        TOBLIGNom.Putvalue ('BLO_PUTTCDEV',pivottodevise(Tresult[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
        TOBLIGNom.Putvalue ('BLO_PMAP',Tresult[6]);
        TOBLIGNom.Putvalue ('BLO_PMRP',Tresult[7]);
        Qte := TOBDetOuv.getValue('BLO_QTEFACT');
        CalculMontantHtDevLigOuv (TOBLigNom,DEV);
        QteDuPv := TOBDetOuv.getValue('BLO_PRIXPOURQTE'); if QteDuPv = 0 then QteDuPv := 1;
        Tresult[0] := Tresult [0] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_DPA')),V_PGI.OkdecP);
        Tresult[1] := Tresult [1] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_DPR')),V_PGI.OkdecP);
        Tresult[2] := Tresult [2] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PUHT')),V_PGI.OkdecP);
        Tresult[3] := Tresult [3] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PUTTC')),V_PGI.OkdecP);
        Tresult[6] := Tresult [6] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PMAP')),V_PGI.OkdecP);
        Tresult[7] := Tresult [7] +arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PMRP')),V_PGI.OkdecP);
        FormatageTableau (Tresult,V_PGI.OkdecP );
      end;
    end;
    for i_ind2 := 0 to TobOuvrage.Detail.Count - 1 do
    begin
      TobFille := TobOuvrage.Detail[i_ind2];
      if (TobFille.getvalue('ANCPV') <> TOBFille.getValue('BLO_PUHTDEV')) and
      (TobFille.detail.count > 0) then
      begin
        MontantInterm := TOBfille.getValue('BLO_PUHTDEV');
        ReajusteMontantOuvrage (nil,TOBPiece,TOBL,TOBFille,TOBFille.getValue('ANCPV'),MontantPr,MontantInterm,DEV,fEnHt);
        TOBFille.putvalue('BLO_PUHTDEV',MontantInterm);
        TOBDetOuv.PutValue('ANCPV',TOBDetOuv.Getvalue('BLO_PUHTDEV') );
        CalculMontantHtDevLigOuv (TOBDetOuv,DEV);
      end;
    end;
    GestionDetailPrixPose (TobOuvrage);
  end;
  Tresult := CalculSurTableau ('/',Tresult,QteDudetail);
  FormatageTableau (Tresult,V_PGI.OkdecP);
end;


function TModifSousDetail.chargeArticle (TheArticle : string ; TOBA : TOB) : boolean;
var QQ : Tquery;
begin
	result := false;
  QQ := OPenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
              		'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
              		'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE GA_ARTICLE ="'+TheArticle+'"',True,-1, '', True);
//	QQ := OPenSql ('SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+TheArticle+'"',true);
  if not QQ.eof then
  begin
  	TOBA.selectdb ('',QQ);
    InitChampsSupArticle (TOBA);
    result := true;
  end;
  ferme (QQ);
end;


function TModifSousDetail.GetCodeArticle(var TypeArticle : string; var Article : string; var CodeArticle: string; var libelle : string; Arow,Acol : integer) : Boolean;
var RechArt : T_RechArt ;
    OkArt,Trouvok   : Boolean ;
    st1 : string;
    TOBArt,TobTemp : TOB;
    TOBPiece,TOBAA : TOB;
    sWhere, stChamps : string;
begin
  Libelle := '';
  TypeArticle := '';
  OkArt:=False ;
  TOBArt := TOB.Create('ARTICLE',nil,-1);
  //-----------------------------------------------------------------
  //  Recherche du code article saisi
  //-----------------------------------------------------------------
  RechArt := TrouverArticle (Codearticle, TOBArt);
  TypeArticle := TOBArt.GetValue('GA_TYPEARTICLE');
  case RechArt of
  	traOk : begin
      //-----------------------------------------------------------------
      //  Code article trouve, on regarde s'il a des nomenclatures associées
      //  si oui et plusieurs, ouverture d'une fenetre de selection
      //  si oui et une seule, on recupere directement son code
      //-----------------------------------------------------------------
      if (Pos(TypeArticle, 'NOM;OUV,ARP') >0) then
      begin
        Trouvok := VerifNomen(TobArt,ARow,st1,Libelle);
        if st1 <> '' then
        begin
          CodeArticle := st1;
          OkArt:=True ;
        end else OkArt:=false;
      end else OkArt := true;
    end;
    traAucun : begin
      //-----------------------------------------------------------------
      // Recherche sur code via LookUp ou Recherche avancée
      //-----------------------------------------------------------------
      ART.Text:= CodeArticle;
      sWhere:=FabricWhereNatArt(fTOBPiece.getValue('GP_NATUREPIECEG'),'','') ;
      StChamps:='GA_CODEARTICLE='+ART.Text+';XX_WHERE=AND '+sWhere ;
      DispatchRecherche(ART, 1, '', stChamps , '');
      if ART.Text <>'' then
      begin
        if not chargeArticle(Art.text,TOBARt) then
        begin
        	result := false;
          exit;
        end;
        CodeArticle := '';
        trouvOk := VerifNomen(TobArt,ARow,st1,Libelle);
        if st1 <> '' then
        begin
          CodeArticle := st1;
          OkArt:=True ;
        end else if TOBArt.GetValue('GA_STATUTART')='GEN' then
        begin
          if ChoisirDimension (TOBArt.GetValue ('GA_ARTICLE'), TOBArt) then
          begin
            OkArt := True;
            Article := TOBART.GetValue('GA_ARTICLE');
          end;
        end else if TrouvOk then OkArt:=True ;
      end;
  	end ;
  traGrille : begin
      //-----------------------------------------------------------------
      // Forcement objet dimension avec saisie obligatoire
      //if GetArticleLookUp (GF_ARTICLE, 'GA_STATUTART = "DIM"') then
      //-----------------------------------------------------------------
      if ChoisirDimension (TOBArt.GetValue ('GA_ARTICLE'), TOBArt) then OkArt := True;
      CodeArticle := '';
      if (TypeArticle= 'ARP') then
      begin
        Trouvok := VerifNomen(TobArt,ARow,st1,Libelle);
        if st1 <> '' then
        begin
          //-----------------------------------------------------------------
          //  on enregistre le code nomenclature dans la colonne GLN_COMPOSE
          //-----------------------------------------------------------------
          CodeArticle := st1;
          OkArt:=True ;
        end else OkArt:=false;
      end else OkArt := true;
  	end;
  end ; // Case

  if (Tobart.GetValue('GA_FERME') = 'X') then
  begin
    PgiInfo('Impossible : Cet article est fermé', TFFacture(XX).Caption);
    OkArt := False;
  end;

  if (OkArt) then
  begin
    //--------------------------------------------------------------------------
    //  on va charger la fiche article dans TobArticles
    //--------------------------------------------------------------------------
    TOBAA := TOB.Create('ARTICLE',fTOBArticles,-1);
    TOBAA.Dupliquer(TOBArt, True, True);
    Article := TOBAA.GetValue('GA_ARTICLE');
    if CodeArticle = '' then CodeArticle := TOBAA.GetValue('GA_CODEARTICLE');
    if Libelle = '' then Libelle := TOBAA.GetValue('GA_LIBELLE');
  end else
  begin
  	Article := '';
  	CodeArticle := '';
    Libelle := '';
  end;
  TOBArt.free;
  Result := OkArt;
end;

procedure TModifSousDetail.SetChange;
begin
	if (fTOBpiece <> nil) and (fTOBPiece.getValue('GP_FACTUREHT')<>'') then fEnHt := (fTOBPiece.getValue('GP_FACTUREHT')='X');
end;

procedure TModifSousDetail.AjouteMtAvancOuvrage(TOBL: TOB;montant: double; Sens:string; Typefacturation : string);
var Pourcent,PuHtdev : double;
  TOBpiece : TOB;
begin
  if Sens = '+' then
  begin
  	TOBL.PutValue('BLF_MTSITUATION',TOBL.GetValue('BLF_MTSITUATION')+Montant);
  	if TOBL.GetString('GL_PIECEPRECEDENTE')<>'' then TOBL.PutValue('BLF_MTCUMULEFACT',TOBL.GetValue('BLF_MTCUMULEFACT')+Montant);
  end else
  begin
  	TOBL.PutValue('BLF_MTSITUATION',TOBL.GetValue('BLF_MTSITUATION')-Montant);
  	if TOBL.GetString('GL_PIECEPRECEDENTE')<>'' then TOBL.PutValue('BLF_MTCUMULEFACT',TOBL.GetValue('BLF_MTCUMULEFACT')-Montant);
  end;

  if TOBL.GetString('GL_PIECEPRECEDENTE')='' then
  begin
    TOBpiece := TOBL.Parent;
    // dans ce cas on doit recalculer le PU unitaire de la ligne
    if TOBL.GetDouble('GL_QTEFACT') <> 0 then
    begin
      PuHtdev := ARRONDI(TOBL.GetDouble('BLF_MTSITUATION')/TOBL.GetDouble('GL_QTEFACT'),V_PGI.OkDecP);
    end else
    begin
      PuHtdev := TOBL.GetDouble('BLF_MTSITUATION');
    end;
    TOBL.SetDouble('GL_PUHTDEV',PuHtdev);
    TOBL.SetDouble('GL_COEFMARG',0);
    TOBL.PutValue('QTECHANGE','X');
    TOBL.putValue('GL_RECALCULER','X');
    TOBPiece.putValue('GP_RECALCULER','X');
  end else
  begin
  if (Pos(TypeFacturation,'DAC;AVA')>0) then // avancement
  begin
    Pourcent := Arrondi(TOBL.GetValue('BLF_MTCUMULEFACT')/TOBL.GetValue('BLF_MTMARCHE')*100,2);
  	TOBL.PutValue('BLF_QTESITUATION',TOBL.GetValue('BLF_QTECUMULEFACT')-TOBL.GetValue('BLF_QTEDEJAFACT'));
  end else
  begin
    if TOBL.GetValue('BLF_MTMARCHE') <> 0 then
    begin
    	Pourcent := Arrondi(TOBL.GetValue('BLF_MTSITUATION')/TOBL.GetValue('BLF_MTMARCHE')*100,2);
    end else
    begin
      Pourcent := 100;
    end;
  end;
  TOBL.PutValue('BLF_POURCENTAVANC',pourcent);
  end;

  if (Pos(TypeFacturation,'DAC;AVA')>0) then // avancement
  begin
    TOBL.PutValue('GL_QTEFACT',TOBL.GetValue('BLF_QTESITUATION'));
    TOBL.PutValue ('GL_QTERESTE',TOBL.GetValue('BLF_QTESITUATION'));
    TOBL.PutValue ('GL_QTESTOCK',TOBL.GetValue('BLF_QTESITUATION'));
    //
    // --- GUINIER ---
    if CtrlOkReliquat(TOBL, 'GL') then TOBL.PutValue('GL_QTERESTE',TOBL.getValue('BLF_MTSITUATION'));

  end;
end;

procedure TModifSousDetail.ModifQteAvancSousDet (TOBL: TOB;Qte: double; Typefacturation : string; DEV : Rdevise);
var TOBO,TOBOD,TOBA : TOB;
		QtePere,QteMaj,QteDuDetail : double;
    zvaleur : T_valeurs;
    Adeduire,Pourcent,QteOuv : double;
begin
  TOBO := FindLigneOuvrageinDoc (TOBL);
  if TOBO <> nil then
  begin
    QteOuv := TOBO.GetValue('GL_QTEFACT');
  	TOBOD := FindDetailOuvrage (TOBL);
    if TOBOD = nil then exit;
    //
    QteDuDetail := TOBOD.getValue('BLO_QTEDUDETAIL'); if QteDuDetail = 0 then QteDuDetail := 1;
//  	TFFActure(XX).deduitLaLigne (TOBO); // pour calculer correctement la pièce
    ADeduire := TOBOD.GetValue('BLF_MTSITUATION');
    //
    TOBOD.PutValue('BLF_QTESITUATION',Qte);
    if QteOuv >0 then
    begin
    	TOBOD.PutValue('BLO_QTEFACT',Arrondi(Qte/QteOuv*QteDuDetail,V_PGI.OkDecQ));
    end else
    begin
    	TOBOD.PutValue('BLO_QTEFACT',Arrondi(Qte*QteDuDetail,V_PGI.OkDecQ));
    end;
    TOBOD.PutValue('BLF_MTSITUATION',Arrondi(TOBOD.GetValue('BLF_QTESITUATION')*TOBOD.GetValue('BLO_PUHTDEV')/TOBOD.GEtValue('BLO_PRIXPOURQTE'),DEV.decimale));
    //
    if TOBL.GetString('GL_PIECEPRECEDENTE')<>'' then
    begin
    if (Pos(TypeFacturation,'DAC;AVA')>0) then // avancement
    begin
    	TOBOD.PutValue('BLF_QTECUMULEFACT',TOBOD.GetValue('BLF_QTEDEJAFACT')+Qte);
    	TOBOD.PutValue('BLF_MTCUMULEFACT',TOBOD.GetValue('BLF_MTDEJAFACT')+TOBOD.GetValue('BLF_MTSITUATION'));
      Pourcent := Arrondi(TOBOD.GetValue('BLF_MTCUMULEFACT')/TOBOD.GetValue('BLF_MTMARCHE')*100,2);
    end else
    begin
      if TOBOD.GetValue('BLF_MTMARCHE') > 0 then
      begin
      	Pourcent := Arrondi(TOBOD.GetValue('BLF_MTSITUATION')/TOBOD.GetValue('BLF_MTMARCHE')*100,2);
      end else
      begin
      	Pourcent := Arrondi(100,2);
      end;
    	CalculMontantHtDevLigOuv (TOBOD,DEV);
    end;
    TOBOD.PutValue('BLF_POURCENTAVANC',Pourcent);
    TOBOD.PutValue('BLO_COEFMARG',0);
    end;
    //
    AfficheLaLigneSousDetail (TOBOD,TOBL,TOBO);
    //
    AjouteMtAvancOuvrage(TOBO,Adeduire,'-',Typefacturation);
    AjouteMtAvancOuvrage(TOBO,TOBOD.GetValue('BLF_MTSITUATION'),'+',Typefacturation);
    //
    TOBO.PutValue('QTECHANGE','X');
    TOBO.putValue('GL_RECALCULER','X');
    TOBPiece.putValue('GP_RECALCULER','X');
    //
    CalculeLaLigneDoc (TOBO,false);
    //
  end;
end;

procedure TModifSousDetail.ModifQteAvancOuv(TOBL: TOB; Qte: double;Typefacturation: string; DEV: Rdevise);
var pourcent,PrixPourQte,QteSit,MtSit,MtLoc : double;
		indiceNomen,II : integer;
    TOBO,TOBOD : TOB;
begin
  TOBL.PutValue('BLF_QTESITUATION',Qte);
  TOBL.PutValue('GL_QTEFACT',TOBL.GetValue('BLF_QTESITUATION'));
  TOBL.PutValue ('GL_QTERESTE',TOBL.GetValue('BLF_QTESITUATION'));
  TOBL.PutValue ('GL_QTESTOCK',TOBL.GetValue('BLF_QTESITUATION'));
  //
  TOBL.PutValue('BLF_MTSITUATION',Arrondi(TOBL.GetValue('BLF_QTESITUATION')*TOBL.GetValue('GL_PUHTDEV')/TOBL.GEtValue('GL_PRIXPOURQTE'),DEV.decimale));
  TOBL.PutValue('BLF_MTPRODUCTION',TOBL.GetValue('BLF_MTSITUATION'));
  // --- GUINIER ---
  if CtrlOkReliquat(TOBL, 'GL') then TOBL.PutValue('GL_QTERESTE',TOBL.getValue('BLF_MTSITUATION'));
  //
  if (Pos(TypeFacturation,'DAC;AVA')>0) then // avancement
  begin
    if TOBL.GetString('GL_PIECEPRECEDENTE')='' then
    begin
       Pourcent := 100;
    end else
    begin
  	TOBL.PutValue('BLF_QTECUMULEFACT',TOBL.GetValue('BLF_QTEDEJAFACT')+Qte);
  	TOBL.PutValue('BLF_MTCUMULEFACT',TOBL.GetValue('BLF_MTDEJAFACT')+TOBL.GetValue('BLF_MTSITUATION'));
    Pourcent := Arrondi(TOBL.GetValue('BLF_MTCUMULEFACT')/TOBL.GetValue('BLF_MTMARCHE')*100,2);
    end;
  end else
  begin
    if TOBL.GetString('GL_PIECEPRECEDENTE')='' then
    begin
       Pourcent := 100;
    end else
    begin
    Pourcent := Arrondi(TOBL.GetValue('BLF_MTSITUATION')/TOBL.GetValue('BLF_MTMARCHE')*100,2);
  end;
  end;
  TOBL.PutValue('BLF_POURCENTAVANC',Pourcent);
  TOBL.putValue('GL_RECALCULER','X');
  TOBPiece.putValue('GP_RECALCULER','X');
  //
	indiceNomen := TOBL.Getvalue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
  TOBO := TOBOuvrages.detail[IndiceNomen-1];
  for II := 0 to TOBO.detail.count -1 do
  begin
  	TOBOD := TOBO.detail[II];
    TOBOD.PutValue('BLF_POURCENTAVANC',Pourcent);
    PrixPourQte :=  TOBOD.GetValue('BLO_QTEDUDETAIL'); if PrixPourQte = 0 then PrixPourQte := 1;
    //
    if TOBL.GetString('GL_PIECEPRECEDENTE')='' then
    begin
      QteSit := arrondi(TOBOD.getValue('BLO_QTEFACT')*TOBL.getvalue('GL_QTEFACT') /TOBOD.getValue('BLO_QTEDUDETAIL'),V_PGI.okdecQ);
      MtLoc := Arrondi(TOBOD.getValue('BLO_QTEFACT') / TOBOD.getValue('BLO_QTEDUDETAIL') * TOBOD.getValue('BLO_PUHTDEV'),V_PGI.okdecP);
      MtSit := arrondi(TOBL.getvalue('GL_QTEFACT') * MtLoc,DEV.Decimale);
      TOBOD.PutValue('BLF_QTESITUATION', QteSit); { NEWPIECE }
      TOBOD.PutValue('BLF_MTSITUATION',MtSit);
      TOBOD.PutValue('BLF_POURCENTAVANC',100);
    end else
    begin
    if (Pos(TypeFacturation,'DAC;AVA')>0) then // avancement
    begin
      TOBOD.PutValue('BLF_QTECUMULEFACT',Arrondi(TOBOD.GetValue('BLF_QTEMARCHE')*TOBOD.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
      TOBOD.PutValue('BLF_MTCUMULEFACT',arrondi(TOBOD.GetValue('BLF_MTMARCHE')*TOBOD.GetValue('BLF_POURCENTAVANC')/100,DEV.Decimale));
      TOBOD.PutValue('BLF_QTESITUATION',TOBOD.GetValue('BLF_QTECUMULEFACT')-TOBOD.GetValue('BLF_QTEDEJAFACT'));
      TOBOD.PutValue('BLF_MTSITUATION',TOBOD.GetValue('BLF_MTCUMULEFACT')-TOBOD.GetValue('BLF_MTDEJAFACT'));
    end else
    begin
      QteSit :=  Arrondi(TOBOD.GetValue('BLF_QTEMARCHE') / TOBL.GetValue('BLF_QTEMARCHE') * Qte ,V_PGI.OkDecQ);
        MtSit := Arrondi ( QteSit*TOBOD.GetValue('BLO_PUHTDEV'),DEV.decimale);
      TOBOD.PutValue('BLF_QTESITUATION',QteSit);
      TOBOD.PutValue('BLF_MTSITUATION',MtSIT);
    end;
  end;
  end;

end;

procedure TModifSousDetail.ModifMontantAvancOuv(TOBL: TOB; Montant: double;Typefacturation: string; DEV: Rdevise);
var pourcent : double;
		indiceNomen,II, imAX : integer;
    MaxMt,PuDev : Double;
    DiffMt,SumMt : Double;
    TOBO,TOBOD : TOB;
begin
	TOBL.SetString('QTECHANGE','X');
  TOBL.PutValue('BLF_MTSITUATION',montant);
  TOBL.PutValue('BLF_MTPRODUCTION',TOBL.GetValue('BLF_MTSITUATION'));

  if TOBL.GetString('GL_PIECEPRECEDENTE') = '' then
  begin
    if TOBL.GetDouble('GL_QTEFACT') <> 0 then
    begin
      PuDev := ARRONDI(montant/TOBL.GetDouble('GL_QTEFACT'),V_PGI.okdecP);
      TFFacture(XX).AppliquePrixOuvrage (TOBL,PuDev);
    end;
    Exit;
  end;
  //
  TOBL.PutValue('BLF_MTCUMULEFACT',TOBL.GetValue('BLF_MTDEJAFACT')+TOBL.GetValue('BLF_MTSITUATION'));
  if (Pos(TypeFacturation,'DAC;AVA')>0) then // avancement
  begin
    Pourcent := Arrondi(TOBL.GetValue('BLF_MTCUMULEFACT')/TOBL.GetValue('BLF_MTMARCHE')*100,2);
  end else
  begin
    Pourcent := Arrondi(TOBL.GetValue('BLF_MTSITUATION')/TOBL.GetValue('BLF_MTMARCHE')*100,2);
  end;
  //
  TOBL.PutValue('BLF_POURCENTAVANC',pourcent);
  TOBL.PutValue('BLF_QTECUMULEFACT',Arrondi(TOBL.GetValue('BLF_QTEMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
  TOBL.PutValue('BLF_QTESITUATION',TOBL.GetValue('BLF_QTECUMULEFACT')-TOBL.GetValue('BLF_QTEDEJAFACT'));
  TOBL.putValue('GL_RECALCULER','X');
  TOBPiece.putValue('GP_RECALCULER','X');

	indiceNomen := TOBL.Getvalue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
  TOBO := TOBOuvrages.detail[IndiceNomen-1];
  for II := 0 to TOBO.detail.count -1 do
  begin
  	TOBOD := TOBO.detail[II];
    TOBOD.PutValue('BLF_POURCENTAVANC',Pourcent);
    //
    TOBOD.PutValue('BLF_QTECUMULEFACT',Arrondi(TOBOD.GetValue('BLF_QTEMARCHE')*TOBOD.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
    TOBOD.PutValue('BLF_QTESITUATION',TOBOD.GetValue('BLF_QTECUMULEFACT')-TOBOD.GetValue('BLF_QTEDEJAFACT'));
    //
    TOBOD.PutValue('BLF_MTCUMULEFACT',arrondi(TOBOD.GetValue('BLF_MTMARCHE')*TOBOD.GetValue('BLF_POURCENTAVANC')/100,DEV.Decimale));
    TOBOD.PutValue('BLF_MTSITUATION',TOBOD.GetValue('BLF_MTCUMULEFACT')-TOBOD.GetValue('BLF_MTDEJAFACT'));
    if TOBOD.GetDouble('BLF_MTSITUATION') > MaxMt then
    begin
      MaxMt := TOBOD.GetDouble('BLF_MTSITUATION') ;
      imAX := II;
    end;
    SumMt := SumMt + TOBOD.GetDouble('BLF_MTSITUATION');
  end;

  SumMt := ARRONDI(SumMt,DEV.Decimale);
  DiffMt := ARRONDI(TOBL.GetDouble('BLF_MTSITUATION')-SumMt,DEV.Decimale);
  if DiffMt <> 0 then
  begin
  	TOBOD := TOBO.detail[imAX];
    TOBOD.SetDouble('BLF_MTCUMULEFACT',TOBOD.GetDouble('BLF_MTCUMULEFACT')+DiffMt);
    TOBOD.PutValue('BLF_MTSITUATION',TOBOD.GetDouble('BLF_MTSITUATION')+DiffMt);
  end;

end;

procedure TModifSousDetail.ModifMontantAvancSousDet (TOBL: TOB;Montant: double; Typefacturation : string; DEV : RDevise);
var TOBO,TOBOD,TOBA : TOB;
		QtePere,QteMaj,QteDuDetail : double;
    zvaleur : T_valeurs;
    Adeduire,Pourcent : double;
begin
  TOBO := FindLigneOuvrageinDoc (TOBL);
  if TOBO <> nil then
  begin
  	TOBOD := FindDetailOuvrage (TOBL);
    if TOBOD = nil then exit;
    //
//  	TFFActure(XX).deduitLaLigne (TOBO); // pour calculer correctement la pièce
    ADeduire := TOBOD.GetValue('BLF_MTSITUATION');
    //
//    TOBOD.PutValue('BLF_MTCUMULEFACT',TOBOD.GetValue('BLF_MTDEJAFACT')-Adeduire);
    TOBOD.PutValue('BLF_MTSITUATION',montant);
    if TOBL.GetString('GL_PIECEPRECEDENTE')='' then
    begin
      if (TOBO.GeTDouble('GL_QTEFACT')<> 0) and (TOBOD.getValue('BLO_QTEFACT')<>0) then
      begin
        if TOBL.GetString('GL_FACTUREHT')='X' then
        begin
          TOBOD.SetDouble('BLO_PUHTDEV',Arrondi(TOBOD.GetValue('BLF_MTSITUATION')/(TOBO.GetValue('GL_QTEFACT')*TOBOD.getValue('BLO_QTEFACT')/TOBOD.getValue('BLO_QTEDUDETAIL')),V_PGI.okDecP));
        end else
        begin
          TOBOD.SetDouble('BLO_PUTTCDEV',Arrondi(TOBOD.GetValue('BLF_MTSITUATION')/(TOBO.GetValue('GL_QTEFACT')*TOBOD.getValue('BLO_QTEFACT')/TOBOD.getValue('BLO_QTEDUDETAIL')),V_PGI.OkDecP));
        end;
      end else
      begin
        if (TOBO.GetString('GL_FACTUREHT')='X') then
        begin
          TOBOD.SetDouble('BLO_PUHTDEV',TOBOD.GetValue('BLF_MTSITUATION'));
        end else
        begin
          TOBOD.SetDouble('BLO_PUTTCDEV',TOBOD.GetValue('BLF_MTSITUATION'));
        end;
      end;
      CalculeLigneAcOuvCumul (TOBOD);
    end else
    begin
    TOBOD.PutValue('BLF_MTCUMULEFACT',TOBOD.GetValue('BLF_MTDEJAFACT')+TOBOD.GetValue('BLF_MTSITUATION'));
    if (Pos(TypeFacturation,'DAC;AVA')>0) then // avancement
    begin
      Pourcent := Arrondi(TOBOD.GetValue('BLF_MTCUMULEFACT')/TOBOD.GetValue('BLF_MTMARCHE')*100,2);
    end else
    begin
      Pourcent := Arrondi(TOBOD.GetValue('BLF_MTSITUATION')/TOBOD.GetValue('BLF_MTMARCHE')*100,2);
    end;
    //
    TOBOD.PutValue('BLF_POURCENTAVANC',pourcent);
    //
    TOBOD.PutValue('BLF_QTECUMULEFACT',Arrondi(TOBOD.GetValue('BLF_QTEMARCHE')*TOBOD.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
    TOBOD.PutValue('BLF_QTESITUATION',TOBOD.GetValue('BLF_QTECUMULEFACT')-TOBL.GetValue('BLF_QTEDEJAFACT'));
    end;
    //
    AfficheLaLigneSousDetail (TOBOD,TOBL,TOBO);
    //
    AjouteMtAvancOuvrage(TOBO,Adeduire,'-',Typefacturation);
    AjouteMtAvancOuvrage(TOBO,montant,'+',Typefacturation);
    //
    TOBO.PutValue('QTECHANGE','X');
    TOBO.putValue('GL_RECALCULER','X');
    TOBPiece.putValue('GP_RECALCULER','X');
    //
    CalculeLaLigneDoc (TOBO,false);
    //
  end;
end;

procedure TModifSousDetail.AppliqueCoefSousDetail (TOBO : TOB;Coef,PP : Double);
var Ii : integer;
    TOBOD : TOB;
begin
  //
  for II := 0 to TOBO.detail.count -1 do
  begin
    TOBOD := TOBO.detail[II];
    TOBOD.PutValue('BLO_COEFMARG',Coef);
    TOBOD.PutValue('POURCENTMARG',PP);
    if TOBOD.detail.count > 0 then
    begin
      AppliqueCoefSousDetail (TOBOD,Coef,PP)
    end else
    begin
      CalculeLigneAcOuv (TOBOD,DEV,true);
    end;
  end;
end;

procedure TModifSousDetail.ModifPourcentAvancOuv (TOBL: TOB;Pourcent: double; Typefacturation : string; DEV : Rdevise);
var TOBO,TOBOD : TOB;
		indiceNomen,II,imAX : integer;
    MaxMt : Double;
    DiffMt,SumMt : Double;
begin
//
	MaxMt := 0;
  imAX := -1;
  TOBL.PutValue('BLF_POURCENTAVANC',Pourcent);
  //
  TOBL.PutValue('BLF_QTECUMULEFACT',Arrondi(TOBL.GetValue('BLF_QTEMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
  TOBL.PutValue('BLF_QTESITUATION',TOBL.GetValue('BLF_QTECUMULEFACT')-TOBL.GetValue('BLF_QTEDEJAFACT'));
  TOBL.PutValue('GL_QTEFACT',TOBL.GetValue('BLF_QTESITUATION'));
  TOBL.PutValue ('GL_QTERESTE',TOBL.GetValue('BLF_QTESITUATION'));
  TOBL.PutValue ('GL_QTESTOCK',TOBL.GetValue('BLF_QTESITUATION'));
  //
  TOBL.PutValue('BLF_MTCUMULEFACT',arrondi(TOBL.GetValue('BLF_MTMARCHE')*TOBL.GetValue('BLF_POURCENTAVANC')/100,DEV.Decimale));
  TOBL.PutValue('BLF_MTSITUATION',TOBL.GetValue('BLF_MTCUMULEFACT')-TOBL.GetValue('BLF_MTDEJAFACT'));
  //
  // --- GUINIER ---
  if CtrlOkReliquat(TOBL, 'GL') then TOBL.PutValue('GL_QTERESTE',TOBL.getValue('BLF_MTSITUATION'));
  //
  TOBL.PutValue('QTECHANGE','X');
  TOBL.putValue('GL_RECALCULER','X');
  TOBPiece.putValue('GP_RECALCULER','X');
  TFFacture(XX).AfficheLaLigne (TOBL.GetIndex +1);
	//
	indiceNomen := TOBL.Getvalue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
  TOBO := TOBOuvrages.detail[IndiceNomen-1];
  for II := 0 to TOBO.detail.count -1 do
  begin
  	TOBOD := TOBO.detail[II];
    TOBOD.PutValue('BLF_POURCENTAVANC',Pourcent);
    //
    TOBOD.PutValue('BLF_QTECUMULEFACT',Arrondi(TOBOD.GetValue('BLF_QTEMARCHE')*TOBOD.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
    TOBOD.PutValue('BLF_QTESITUATION',TOBOD.GetValue('BLF_QTECUMULEFACT')-TOBOD.GetValue('BLF_QTEDEJAFACT'));
    //
    TOBOD.PutValue('BLF_MTCUMULEFACT',arrondi(TOBOD.GetValue('BLF_MTMARCHE')*TOBOD.GetValue('BLF_POURCENTAVANC')/100,DEV.Decimale));
    TOBOD.PutValue('BLF_MTSITUATION',TOBOD.GetValue('BLF_MTCUMULEFACT')-TOBOD.GetValue('BLF_MTDEJAFACT'));
    if TOBOD.GetDouble('BLF_MTSITUATION') > MaxMt then
    begin
      MaxMt := TOBOD.GetDouble('BLF_MTSITUATION') ;
      imAX := II;
    end;
    SumMt := SumMt + TOBOD.GetDouble('BLF_MTSITUATION');
  end;
  SumMt := ARRONDI(SumMt,DEV.Decimale);
  DiffMt := ARRONDI(TOBL.GetDouble('BLF_MTSITUATION')-SumMt,DEV.Decimale);
  if (DiffMt <> 0) and (imax >= 0) then
  begin
  	TOBOD := TOBO.detail[imAX];
    TOBOD.SetDouble('BLF_MTCUMULEFACT',TOBOD.GetDouble('BLF_MTCUMULEFACT')+DiffMt);
    TOBOD.PutValue('BLF_MTSITUATION',TOBOD.GetDouble('BLF_MTSITUATION')+DiffMt);
  end;
  TOBL.PutValue('BLF_MTPRODUCTION',TOBL.GetValue('BLF_MTSITUATION'));
end;

function TModifSousDetail.ModifPourcentAvancSousDet(TOBL: TOB;Pourcent: double; Typefacturation : string; DEV : RDevise) : boolean;
var TOBO,TOBOD,TOBA : TOB;
		QtePere,QteMaj,QteDuDetail : double;
    zvaleur : T_valeurs;
    Adeduire : double;
begin
  Result := true;
  TOBO := FindLigneOuvrageinDoc (TOBL);
  if TOBO <> nil then
  begin
    if (TOBO.GetString ('GL_PIECEPRECEDENTE')='') then
    begin
      Result := false;
      exit;
    end;
    //
  	TOBOD := FindDetailOuvrage (TOBL);
    if TOBOD = nil then exit;
    //
  	TFFActure(XX).deduitLaLigne (TOBO); // pour calculer correctement la pièce
    ADeduire := TOBOD.GetValue('BLF_MTSITUATION');
    //
    TOBOD.PutValue('BLF_POURCENTAVANC',Pourcent);
    //
    TOBOD.PutValue('BLF_QTECUMULEFACT',Arrondi(TOBOD.GetValue('BLF_QTEMARCHE')*TOBOD.GetValue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
    TOBOD.PutValue('BLF_QTESITUATION',TOBOD.GetValue('BLF_QTECUMULEFACT')-TOBOD.GetValue('BLF_QTEDEJAFACT'));
    //
    TOBOD.PutValue('BLF_MTCUMULEFACT',arrondi(TOBOD.GetValue('BLF_MTMARCHE')*TOBOD.GetValue('BLF_POURCENTAVANC')/100,DEV.Decimale));
    TOBOD.PutValue('BLF_MTSITUATION',TOBOD.GetValue('BLF_MTCUMULEFACT')-TOBOD.GetValue('BLF_MTDEJAFACT'));
    AfficheLaLigneSousDetail (TOBOD,TOBL,TOBO);
    //
    AjouteMtAvancOuvrage(TOBO,Adeduire,'-',Typefacturation);
    AjouteMtAvancOuvrage(TOBO,TOBOD.GetValue('BLF_MTSITUATION'),'+',Typefacturation);
    //
    TOBO.PutValue('QTECHANGE','X');
    TOBO.putValue('GL_RECALCULER','X');
    TOBPiece.putValue('GP_RECALCULER','X');
    //
    CalculeLaLigneDoc (TOBO,false);
    //
  end;
end;



function TModifSousDetail.GetTypeFacturation: string;
begin
  result := RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS'));
end;

procedure TModifSousDetail.GetBlocNote(TOBL: TOB);
var TOBOD : TOB;
begin
  TOBOD := FindDetailOuvrage (TOBL);
	if TOBOD <> nil then TOBL.SetString('GL_BLOCNOTE',TOBOD.GetString('BLO_BLOCNOTE'));
end;

procedure TModifSousDetail.SetBlocNote(TOBL: TOB);
var TOBOD : TOB;
begin
  TOBOD := FindDetailOuvrage (TOBL);
	if TOBOD <> nil then TOBOD.SetString('BLO_BLOCNOTE',TOBL.GetString('GL_BLOCNOTE'));
end;

procedure TModifSousDetail.ModifCoefMargSousDet (TOBL: TOB;Coef,PP: double; DEV : Rdevise);
var TOBO,TOBOD,TOBA : TOB;
begin
  TOBO := FindLigneOuvrageinDoc (TOBL);
  if TOBO <> nil then
  begin
  	TOBOD := FindDetailOuvrage (TOBL);
    if TOBOD = nil then exit;
    if IsLignefacture (TOBOD) then
    begin
      PgiError(Traduirememoire('Impossible : une facturation a déjà été effectuée sur cette ligne.'));
      exit;
    end;
    //
    TOBOD.PutValue('BLO_COEFMARG',Coef);
    TOBOD.PutValue('POURCENTMARG',PP);
    AppliqueCoefSousDetail (TOBOD,Coef,PP);
    CalculeLigneAcOuv (TOBOD,DEV,true);
    //
    TOBL.PutValue('GL_COEFMARG',Coef);
    TOBL.PutValue('POURCENTMARG',PP);
    TOBL.SetDouble('GL_PUHTDEV',TOBOD.getDouble('BLO_PUHTDEV'));
    //
    TOBO.putValue('GL_RECALCULER','X');
    TOBPiece.putValue('GP_RECALCULER','X');
    CalculeLaLigneDoc (TOBO);
    //
  end;
end;

procedure TModifSousDetail.ModifCoefMargOuv(TOBL: TOB; Coef,PP: double; DEV: Rdevise);
var pourcent,PrixPourQte,QteSit,MtSit,MtLoc,PUHT,PM : double;
		indiceNomen,II : integer;
    TOBO,TOBOD : TOB;
begin
  TOBL.PutValue('GL_COEFMARG',Coef);
  TOBL.PutValue('POURCENTMARG',PP);
  PUHT := Arrondi(TOBL.GetValue('GL_DPR')* Coef,V_PGI.OkdecP);
  if PUHT <> 0 then
  begin
    PM := Arrondi((PUHT - TOBL.GetValue('GL_DPR')) / PUHT,4) * 100;
  end else
  begin
    PM := 0;
  end;
  TOBL.PutValue('POURCENTMARQ',PM);

  TOBL.putValue('GL_RECALCULER','X');
  TOBPiece.putValue('GP_RECALCULER','X');
  //
	indiceNomen := TOBL.Getvalue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
  TOBO := TOBOuvrages.detail[IndiceNomen-1];
  for II := 0 to TOBO.detail.count -1 do
  begin
  	TOBOD := TOBO.detail[II];
    TOBOD.PutValue('BLO_COEFMARG',Coef);
    TOBOD.PutValue('POURCENTMARG',PP);
    AppliqueCoefSousDetail (TOBOD,Coef,PP);
    CalculeLigneAcOuv (TOBOD,DEV,true);
  end;
end;


procedure TModifSousDetail.ModifCoefMarqSousDet (TOBL: TOB;CM,PM: double; DEV : Rdevise);
var TOBO,TOBOD,TOBA : TOB;
    Coef,Pourc : double;
    PUHT : double;
begin
  TOBO := FindLigneOuvrageinDoc (TOBL);
  if TOBO <> nil then
  begin
  	TOBOD := FindDetailOuvrage (TOBL);
    if TOBOD = nil then exit;
    if IsLignefacture (TOBOD) then
    begin
      PgiError(Traduirememoire('Impossible : une facturation a déjà été effectuée sur cette ligne.'));
      exit;
    end;
    //
    if CM <> 1 then
    begin
      PUHT :=  Arrondi(TOBOD.GetValue('BLO_DPR')/ (1 - CM),V_PGI.OkdecP);
    end else
    begin
      PuHt :=  TOBOD.GetValue('BLO_DPR');
    end;
    if PUHT <> 0 then
    begin
      Coef := 1 + Arrondi((PUHT - TOBOD.GetValue('BLO_DPR')) / TOBOD.GetValue('BLO_DPR'),4);
      Pourc := (Coef - 1) * 100;
    end else
    begin
      Coef := 0;
      Pourc := 0;
    end;
    // ---
    TOBOD.PutValue('BLO_COEFMARG',Coef);
    TOBOD.PutValue('POURCENTMARG',Pourc);
    TOBOD.PutValue('POURCENTMARQ',PM);
    AppliqueCoefSousDetail (TOBOD,Coef,Pourc);
    CalculeLigneAcOuv (TOBOD,DEV,true);
    //
    TOBL.PutValue('GL_COEFMARG',Coef);
    TOBL.PutValue('POURCENTMARG',Pourc);
    TOBL.PutValue('POURCENTMARQ',PM);
    TOBL.SetDouble('GL_PUHTDEV',TOBOD.getDouble('BLO_PUHTDEV'));
    //
    TOBO.putValue('GL_RECALCULER','X');
    TOBPiece.putValue('GP_RECALCULER','X');
    CalculeLaLigneDoc (TOBO);
    //
  end;
end;

procedure TModifSousDetail.ModifCoefMarqOuv(TOBL: TOB; CM,PM: double; DEV: Rdevise);
var pourcent,PrixPourQte,QteSit,MtSit,MtLoc : double;
		indiceNomen,II : integer;
    TOBO,TOBOD : TOB;
    Coef,PP,PUHT : double;
begin

  if CM <> 1 then
  begin
    PUHT :=  Arrondi(TOBL.GetValue('GL_DPR')/ (1 - CM),V_PGI.OkdecP);
  end else
  begin
    PuHt :=  TOBL.GetValue('GL_DPR');
  end;
  
  if PUHT <> 0 then
  begin
    Coef := 1+ Arrondi((PUHT - TOBL.GetValue('GL_DPR')) / TOBL.GetValue('GL_DPR'),4);
    PP := (Coef -1) * 100;
  end else
  begin
    Coef := 0;
    PP := 0;
  end;

  TOBL.PutValue('GL_COEFMARG',Coef);
  TOBL.PutValue('POURCENTMARG',PP);
  TOBL.PutValue('POURCENTMARQ',PM);

  TOBL.putValue('GL_RECALCULER','X');
  TOBPiece.putValue('GP_RECALCULER','X');
  //
	indiceNomen := TOBL.Getvalue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
  TOBO := TOBOuvrages.detail[IndiceNomen-1];
  for II := 0 to TOBO.detail.count -1 do
  begin
  	TOBOD := TOBO.detail[II];
    TOBOD.PutValue('BLO_COEFMARG',Coef);
    TOBOD.PutValue('POURCENTMARG',PP);
    TOBOD.PutValue('POURCENTMARQ',PM);
    AppliqueCoefSousDetail (TOBOD,Coef,PP);
    CalculeLigneAcOuv (TOBOD,DEV,true);
  end;
end;



end.

