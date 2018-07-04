unit UrectifSituations;

interface

Uses HEnt1, HCtrls, Controls, ComCtrls, StdCtrls, ExtCtrls,
     SysUtils, Classes, Graphics, Grids, Forms, Saisutil, EntGC, AGLInit, FactUtil,FactRg,ParamSoc,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fe_Main,
{$ENDIF}
{$IFDEF AFFAIRE}
     AffaireUtil,
{$endif}
     UtilFO, Messages, Windows,
{$IFDEF BTP}
     BTPUtil,FactOuvrage,
{$ENDIF}
     UtilPgi,
     FactTOB,
     UTOB,
     uEntCommun,
     UtilChainage;

type

  TRectifSituation = class
    private
      TOBPiece,TOBpieceTrait,TOBSSTRAIT,TOBTimbres,TOBBasesL,TOBBases,TOBEches,TOBPorcs,TOBAdresses : TOB;
      TOBAnaP,TOBAnaS,TOBLienOle,TOBPieceRG,TOBBasesRG,TOBOuvrage,TOBOuvragesP : TOB;
      TOBTiers,TOBArticles,TOBNomenclature,TOBAffaire,TOBCPTA,TOBMetres : TOB;
      TOBAffaireInterv,TOBACOMPTES : TOB;
      TOBLIGNEFAC : TOB;
      TOBRESULTATS : TOB;
      TOBVTECOLL : TOB;
      // --
      XTOBPiece,XTOBpieceTrait,XTOBSSTRAIT,XTOBTimbres,XTOBBasesL,XTOBBases,XTOBEches,XTOBPorcs,XTOBAdresses : TOB;
      XTOBAnaP,XTOBAnaS,XTOBLienOle,XTOBPieceRG,XTOBBasesRG,XTOBOuvrage,XTOBOuvragesP : TOB;
      XTOBTiers,XTOBArticles,XTOBNomenclature,XTOBAffaire,XTOBCPTA,XTOBMetres : TOB;
      XTOBVTECOLL : TOB;
      XTOBAffaireInterv,XTOBACOMPTES : TOB;
      ModeDebug : boolean;

      XX : TForm;
      DEV : Rdevise;
      // ---
      IndiceOuv : integer;
      LastSituation : boolean;
      nextCledoc : r_cledoc;
      NewDate : TdateTime;
      NumSituation : integer;
      ModeGeneration : string;
      InAvancement : boolean;
      // ---
      procedure ChargeLaPiece (Cledoc : r_cledoc);
      procedure ConstitueAvoir (NumSituation,Numfac : integer; DatePiece : TdateTime);
      procedure ConstitueLigneFac;
      procedure ConstitueLigneFacOuvr(TOBL, TOBOuvrage: TOB);
      procedure ConstitueNewFacture;
      procedure ConstitueSituation;
      procedure Enregistre ;
      procedure EnregistreCreationDoc (Ioerror : TIOErr; DebutTitre,LastErrorMessage : string);
      procedure EnregistreDocument;
      procedure EnregistreDocumentModifie;
      function GetNumSituation(TOBPiece: TOB): integer;
      function FindTobLigne(TOBLF: TOB; FromCourant :boolean=false): TOB;
      procedure Init;
      procedure LoadLesArticles (Cledoc : r_cledoc);
      procedure PositionneAvancementDevis;
      procedure PositionneToutModif;
      procedure RectifieAvancement;
      procedure RecalculePiece;
      procedure RecalculePieceCourante;
      procedure SetInfoNewDocument;
      procedure SetPieceMorte (cledoc : r_cledoc);
      procedure ToutAllouer;
      procedure ToutLiberer;
    public
      LastErrorMessage : string;
      Validate : boolean;
      constructor create (FF : TForm);
      destructor destroy; override;
      procedure RectifieSituation;
  end;



implementation
uses FactCpta,Facture,UtilTOBpiece,UCotraitance,
     BTSAISIEPAGEMETRE_TOF,FactTimbres,factligneBase,
     FactAdresse,FactNomen,FactCalc,factPiece,FactGrp,
     FactTiers,factSituations,factComm,Splash,UCumulCollectifs
     ;


{ TRectifSituation }

procedure TRectifSituation.ChargeLaPiece(Cledoc: r_cledoc);
var Q : TQuery;
begin
  IndiceOuv := 1;
  LoadLesArticles (cledoc);
  LoadPieceLignes(CleDoc, TobPiece,true,false,false,false);
  PieceAjouteSousDetail(TOBPiece,true,false,true);
  RemplirTOBTiers (TOBTiers,TOBPiece.GetString('GP_TIERS'),TOBPiece.GetString('GP_NATUREPIECEG'),false);
  LoadLaTOBPieceTrait(TOBpieceTrait,cledoc,'');
  LoadLesSousTraitants(cledoc,TOBSSTRAIT);
  loadlesMetres (Cledoc,TOBMetres);
  LoadlesTimbres(cledoc,TOBTimbres);
  // Lecture bases Lignes
  Q := OpenSQL('SELECT * FROM LIGNEBASE WHERE ' + WherePiece(CleDoc, ttdLigneBase, False), True,-1, '', True);
  TOBBasesL.LoadDetailDB('LIGNEBASE', '', '', Q, False);
  Ferme(Q);
{$IFDEF BTP}
  OrdonnelignesBases (TOBBasesL);
{$ENDIF}
  // Lecture bases
  Q := OpenSQL('SELECT * FROM PIEDBASE WHERE ' + WherePiece(CleDoc, ttdPiedBase, False), True,-1, '', True);
  TOBBases.LoadDetailDB('PIEDBASE', '', '', Q, False);
  Ferme(Q);
  // Lecture Echéances
  Q := OpenSQL('SELECT *,(SELECT T_LIBELLE FROM TIERS WHERE T_TIERS=GPE_FOURNISSEUR AND T_NATUREAUXI="FOU") AS LIBELLE  FROM PIEDECHE WHERE ' + WherePiece(CleDoc, ttdEche, False), True,-1, '', True);
  TOBEches.LoadDetailDB('PIEDECHE', '', '', Q, False);
  Ferme(Q);
  // Lecture Ports
  Q := OpenSQL('SELECT *," " AS SEL FROM PIEDPORT WHERE ' + WherePiece(CleDoc, ttdPorc, False), True,-1, '', True);
  TOBPorcs.LoadDetailDB('PIEDPORT', '', '', Q, False);
  Ferme(Q);
  // Lecture Adresses
  LoadLesAdresses(TOBPiece, TOBAdresses);
  // Lecture Nomenclatures
  LoadLesNomen(TOBPiece, TOBNomenclature, TOBArticles, CleDoc);
  // Lecture Analytiques
  LoadLesAna(TOBPiece, TOBAnaP, TOBAnaS);
  // Modif BTP
  // chargement textes debut et fin
  LoadLesBlocNotes([tModeGridStd], TOBLienOle, Cledoc);
  // Chargement des retenues de garantie et Tva associe
  if GetParamSoc('SO_RETGARANTIE') then LoadLesRetenues(TOBPiece, TOBPieceRG, TOBBasesRG, TaModif);
  // Lecture Ouvrages
  LoadLesOuvrages(TOBPiece, TOBOuvrage, TOBArticles, Cledoc, IndiceOuv,nil);
  if (TFFACTURE(XX).IsEcritLesOuvPlat)  then LoadLesOuvragesPlat(TOBPiece, TOBOuvragesP, Cledoc);
  LoadLaTOBAffaireInterv(TOBAffaireInterv,TOBPiece.getValue('GP_AFFAIRE'));
  LoadLesVTECOLLECTIF (CleDoc,TOBVTECOLL);
end;

procedure TRectifSituation.LoadLesArticles (Cledoc : r_cledoc);
var Sql : String;
    QQ : TQuery;
begin
  Sql := 'SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
         'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
         'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE '+
         'WHERE GA_ARTICLE IN (SELECT DISTINCT GL_ARTICLE FROM LIGNE WHERE '+
         WherePiece(cledoc,ttdLigne,false)+') UNION ALL '+
         'SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
         'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
         'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE '+
         'WHERE GA_ARTICLE IN (SELECT DISTINCT BLO_ARTICLE FROM LIGNEOUV WHERE '+
         WherePiece(cledoc,ttdOuvrage,false)+') AND NOT GA_ARTICLE IN (SELECT DISTINCT GL_ARTICLE FROM LIGNE WHERE '+
         WherePiece(cledoc,ttdLigne,false)+') ORDER BY GA_ARTICLE';
  QQ := OpenSql(Sql,true,-1,'',true);
  if not QQ.eof then TOBArticles.LoadDetailDB ('ARTICLE','','',QQ,false);
  ferme (QQ);
end;

procedure TRectifSituation.ConstitueAvoir (NumSituation,Numfac : integer; DatePiece : TdateTime);

  procedure AjusteDocument;
  var indice : integer;
      TOBL : TOB;
      IndiceOuv : integer;
  begin
    {$IFDEF BTP}
    for Indice := 0 to TOBPiece.detail.count -1 do
    begin
      TOBL := TOBpiece.detail[indice];
      IndiceOuv := TOBL.GetValue('GL_INDICENOMEN');
      if (IndiceOuv > 0) and (IsOuvrage(TOBL)) then ReAffecteLigOuv(IndiceOuv, TobL, TobOuvrage);
      // protection des coefs de marge, fg, fc , FR pour la validation : cf pb plumard rzzzz rrzzz Avril 2010
      ProtectionCoef  (TOBL);
    end;
    {$ENDIF}
  end;

  procedure AjouteCommentaireProvenance (NumSituation,Numfac : integer; DatePiece : TdateTime);
  var Comment : string;
      TOBB : TOB;
  begin
    if NumSituation <> 0 then
    begin
      Comment := 'Avoir sur Situation N° '+InttoStr(NumSituation)+' / Facture N° '+InttoStr(Numfac)+' du '+ DateToStr(DatePiece);
    end else
    begin
      Comment := 'Avoir sur Facture N° '+InttoStr(Numfac)+' du '+ DateToStr(DatePiece);
    end;
    TOBB:=NewTOBLigne(TOBPiece,1);
    TOBB.PutValue('GL_NUMORDRE', 0);
    TOBB.PutValue('GL_LIBELLE',Comment)    ; TOBB.PutValue('GL_TYPELIGNE','COM') ;
    TOBB.PutValue('GL_TYPEDIM','NOR')   ; TOBB.PutValue('GL_CODEARTICLE','') ;
    TOBB.PutValue('GL_ARTICLE','')      ; TOBB.PutValue('GL_QTEFACT',0) ;
    TOBB.PutValue('GL_QTESTOCK',0)      ; TOBB.PutValue('GL_PUHTDEV',0) ;
    TOBB.PutValue('GL_QTERESTE',0)      ; { NEWPIECE }
    TOBB.PutValue('GL_PUTTCDEV',0)      ; TOBB.PutValue('GL_TYPEARTICLE','') ;
    TOBB.PutValue('GL_PUHT',0)          ; TOBB.PutValue('GL_PUHTNET',0) ;
    TOBB.PutValue('GL_PUTTC',0)         ; TOBB.PutValue('GL_PUTTCNET',0) ;
    TOBB.PutValue('GL_PUHTBASE',0)      ; TOBB.PutValue('GL_FAMILLETAXE1','') ;
    TOBB.PutValue('GL_TYPENOMENC','')   ; TOBB.PutValue('GL_QUALIFMVT','') ;
    TOBB.PutValue('GL_REFARTSAISIE','') ; TOBB.PutValue('GL_REFARTBARRE','') ;
    TOBB.PutValue('GL_REFCATALOGUE','') ; TOBB.PutValue('GL_TYPEREF','') ;
    TOBB.PutValue('GL_REFARTTIERS','')  ;
    {Modif AC 4/07/03 Pas de GL_CODESDIM sur les lignes commentaire}
    TOBB.PutValue('GL_CODESDIM','')  ;
    {Fin Modif AC}
    {Modif JLD 20/06/2002}
    TOBB.PutValue('GL_ESCOMPTE',TOBPiece.GetValue('GP_ESCOMPTE')) ;
    TOBB.PutValue('GL_REMISEPIED',TOBPiece.GetValue('GP_REMISEPIED')) ;
    {Fin modif}

    //JS 17/06/03
    TOBB.PutValue('GL_INDICESERIE',0) ; TOBB.PutValue('GL_INDICELOT',0) ;
    TOBB.PutValue('GL_REMISELIGNE',0) ;
    // Modif BTP
    TOBB.PutValue('GL_TYPEARTICLE','COM');
    TOBB.PutValue('GL_PUHTNETDEV',0)    ; TOBB.PutValue('GL_PUTTCNETDEV',0) ;
    TOBB.PutValue('GL_BLOCNOTE','')    ; TOBB.PutValue('GL_QUALIFQTEVTE','') ;
    TOBB.PutValue('GL_INDICENOMEN',0) ;
    // ---
    ZeroLigne(TOBB) ;
  end;

var NewDoc : R_Cledoc;
    NewNum : integer;
begin
  FillChar(NewDoc,sizeof(NewDoc),#0);
  NewDoc.Naturepiece  := 'ABT';
  NewDoc.DatePiece := NewDate;
  NewDoc.Souche       := GetSoucheG(NewDoc.NaturePiece, TOBPiece.getString('GP_ETABLISSEMENT'), TOBPiece.getString('GP_DOMAINE'));
  NewNum := SetNumberAttribution(NewDoc.Naturepiece,NewDoc.Souche, NewDoc.DatePiece,1);
  if NewNum = 0 then
  begin
    LastErrorMessage := 'Erreur : N° de compteur des avoirs';
    V_PGI.ioerror := oeUnknown;
    exit;
  end else
  begin
    NewDoc.NumeroPiece := NewNum ;
  end;
  MajFromCleDoc (TOBpiece,NewDoc);

  AjouteCommentaireProvenance(NumSituation,Numfac,DatePiece);
  NumeroteLignesGC(nil,TOBpiece, true, false);
  InverseLesPieces(TOBPiece, 'PIECE');
  InverseLesPieces(TOBBases, 'PIEDBASE');
  InverseLesPieces(TOBEches, 'PIEDECHE');
  InverseLesPieces(TOBPorcs, 'PIEDPORT'); //mcd 07/06/02 port non pris en compte
  InverseLesPieces(TOBPieceRG, 'PIECERG');
  InverseLesPieces(TOBBasesRG, 'PIEDBASERG');
  InverseLesPieces(TOBOuvragesP, 'LIGNEOUVPLAT');
  AjusteDocument; // pour les ouvrages
  PositionneToutModif;
  // --
end;

constructor TRectifSituation.create (FF : TForm);
begin
  ModeDebug := false;
  Validate := false;
  XX := FF;
  FillChar(nextCledoc,sizeof(nextCledoc),#0);
  ToutAllouer;
  lastSituation := true;
end;

destructor TRectifSituation.destroy;
begin
  ToutLiberer;
  inherited;
end;


procedure TRectifSituation.Enregistre;
begin
  PositionneToutModif;
  XTOBpiece := TOBPiece;
  XTOBOUvrage := TOBOuvrage;
  XTOBOuvragesP := TOBOuvragesP;
  XTOBBases := TOBBases;
  XTOBBasesL := TOBBasesL;
  XTOBEches := TOBEches;
  XTOBPieceTrait := TOBPieceTrait;
  XTOBAffaireInterv := TOBAffaireInterv;
  XTOBTiers := TOBTiers;
  XTOBArticles := TOBArticles;
  XTOBCpta := TOBCpta;
  XTOBAcomptes := TOBAcomptes;
  XTOBPorcs := TOBPorcs ;
  XTOBPIECERG := TOBPIECERG;
  XTOBBASESRG := TOBBASESRG;
  XTOBVTECOLL := TOBVTECOLL;
  XTOBanaP := TOBanaP;
  XTOBanaS := TOBanaS;
  XTOBSSTRAIT := TOBSSTRAIT;
  XTOBMetres := TOBMetres;
  XTOBLIENOLE := TOBLIENOLE;
  XTOBNomenclature := TOBNomenclature;
  XTOBAdresses := TOBAdresses;
  XTOBAffaire := TOBAffaire;
  XTOBTimbres := TOBTimbres;
  EnregistreDocument;
end;

procedure TRectifSituation.EnregistreDocumentModifie;
var newDoc : r_cledoc;
    NewNum : integer;
begin
  TFFacture (XX).BeforeEnreg;
  XTOBpiece := TFFacture(XX).LaPieceCourante;
  XTOBOUvrage := TFFacture(XX).TheTOBOuvrage;
  XTOBOuvragesP := TFFacture(XX).TheTOBOuvrageP;
  XTOBBases := TFFacture(XX).TheTOBBases;
  XTOBBasesL := TFFacture(XX).TheTOBBasesL;
  XTOBEches := TFFacture(XX).TheEches;
  XTOBPieceTrait := TFFacture(XX).ThePieceTrait;
  XTOBAffaireInterv := TFFacture(XX).ThePieceAffaire;
  XTOBTiers := TFFacture(XX).TheTOBTiers;
  XTOBArticles := TFFacture(XX).TheTOBArticles;
  XTOBCpta := TOBCPTA;  TOBCPTA.ClearDetail;
  XTOBAcomptes := TFFacture(XX).TheAcomptes;
  XTOBPorcs := TFFacture(XX).TheTOBPorcs ;
  XTOBPIECERG := TFFacture(XX).TheTOBPieceRG;
  XTOBBASESRG := TFFacture(XX).TheTOBbasesRG;
  XTOBVTECOLL := TFFacture(XX).XTOBVTECOLL ;
  XTOBanaP := TFFacture(XX).TheTOBANAP;
  XTOBanaS := TFFacture(XX).TheTOBANAS;
  XTOBSSTRAIT := TFFacture (XX).TheTOBSSTRAIT;
  XTOBMetres := TFFacture (XX).TheMetres;
  XTOBLIENOLE := TFFacture (XX).TheTOBliensOle;
  XTOBNomenclature := TFFacture (XX).TheNomenclature;
  XTOBAdresses := TFFacture(XX).TheTOBAdresses;
  XTOBAffaire := TFFacture(XX).TheTOBAffaire;
  XTOBTimbres := TFFacture (XX).TheTOBTimbres;
  //
  XTOBpiece.SetString('GP_REFCOMPTABLE','');
  //
  FillChar(NewDoc,sizeof(NewDoc),#0);
  NewDoc.Naturepiece  := TFFacture(XX).CleDoc.NaturePiece;
  NewDoc.DatePiece := NewDate;
  NewDoc.Souche       := GetSoucheG(NewDoc.NaturePiece, TOBPiece.getString('GP_ETABLISSEMENT'), TOBPiece.getString('GP_DOMAINE'));
  NewNum := SetNumberAttribution(NewDoc.Naturepiece,NewDoc.Souche, NewDoc.DatePiece,1);
  if NewNum = 0 then
  begin
    LastErrorMessage := 'Erreur : N° de compteur des Factures';
    V_PGI.ioerror := oeUnknown;
    exit;
  end else
  begin
    NewDoc.NumeroPiece := NewNum ;
  end;
  MajFromCleDoc (XTOBpiece,NewDoc);
  RecalculePieceCourante;
  //
  EnregistreDocument;
end;



procedure TRectifSituation.EnregistreDocument;
var OldEcr, OldStk: RMVT;
    DebutTitre : string;
begin
  LastErrorMessage := ' ---> OK';
  // --
  SetInfoNewDocument;
  if (Pos(XTOBPiece.getString('GP_NATUREPIECEG'),'FBT;FBP')>0) then
  begin
    if NumSituation <> 0 then
    begin
      DebutTitre := 'Génération Situation N° '+InttoStr(NumSituation)+' / Facture N° '+XTOBPiece.getString('GP_NUMERO')+' du '+ DateToStr(XTOBpiece.GetDateTime('GP_DATEPIECE'));
    end else
    begin
      DebutTitre := 'Génération Facture N° '+XTOBPiece.getString('GP_NUMERO')+' du '+ DateToStr(XTOBpiece.GetDateTime('GP_DATEPIECE'));
    end;
  end else
  begin
    DebutTitre := 'Génération Avoir N° '+XTOBPiece.getString('GP_NUMERO')+' au '+ DateToStr(XTOBpiece.GetDateTime('GP_DATEPIECE'));
  end;
  TRY
    //
    ValideLesLignesCompl(XTOBpiece,nil);
    if V_PGI.ioerror <> oeOk then
    begin
      LastErrorMessage := 'Erreur Ecriture LIGNECOMPL';
      exit;
    end;
    ConstitueLigneFac;
    //
    if not PassationComptable ( XTOBPiece,XTOBOUvrage,XTOBOuvragesP, XTOBBases,XTOBBasesL,
                               XTOBEches,XTOBPieceTrait,XTOBAffaireInterv,XTOBTiers, XTOBArticles, XTOBCpta, XTOBAcomptes,
                               XTOBPorcs, XTOBPIECERG, XTOBBASESRG, XTOBanaP,XTOBanaS,XTOBSSTRAIT,XTOBVTECOLL, DEV, OldEcr,
                               OldStk, True, true) then
    begin
      LastErrorMessage := 'Erreur génération comptable';
      exit;
    end;
    if XTOBVTECOLL.detail.count > 0 then
    begin
      PrepareInsertCollectif (XTOBPiece,XTOBVTECOLL);
      if not XTOBVTECOLL.InsertDB(nil) then
      begin
        LastErrorMessage := 'Erreur mise à jour des TTC / COLLECTIFS';
        V_PGI.IoError := oeUnknown;
        exit;
      end;
    end;
    if not XTOBPorcs.InsertDB(nil) then
    begin
      LastErrorMessage := 'Erreur mise à jour des Ports/frais';
      V_PGI.IoError := oeUnknown;
      exit;
    end;
    if not XTOBPieceTrait.InsertDB(nil) then
    begin
      LastErrorMessage := 'Erreur mise à jour des PIECETRAIT';
      V_PGI.IoError := oeUnknown;
      exit;
    end;
    if not XTOBSSTrait.InsertDB(nil) then
    begin
      LastErrorMessage := 'Erreur mise à jour des PIECEINTERV';
      V_PGI.IoError := oeUnknown;
      exit;
    end;
    if not XTOBPieceRG.InsertDB(nil) then
    begin
      LastErrorMessage := 'Erreur mise à jour des RG';
      V_PGI.IoError := oeUnknown;
      exit;
    end;
    if not XTOBBasesRG.InsertDB(nil) then
    begin
      LastErrorMessage := 'Erreur mise à jour des Bases RG';
      V_PGI.IoError := oeUnknown;
      exit;
    end;
    if not XTOBTimbres.InsertDB(nil) then
    begin
      LastErrorMessage := 'Erreur mise à jour des Timbres';
      V_PGI.IoError := oeUnknown;
      exit;
    end;
    if not XTOBPiece.InsertDBByNivel(False) then
    begin
      LastErrorMessage :='Erreur mise à jour ligne';
      V_PGI.IoError := oeUnknown;
      exit;
    end;
    ValideLesOuv(XTOBOuvrage, XTOBPiece);
    if (not XTOBOuvragesP.InsertDBByNivel(false)) then
    begin
      LastErrorMessage := 'Erreur mise à jour OUVRAGES PLAT';
      V_PGI.IoError := oeUnknown;
      exit;
    end;
    if not XTOBBases.InsertDB(nil) then
    begin
      LastErrorMessage := 'Erreur mise à jour BASES';
      V_PGI.IoError := oeUnknown;
      exit;
    end;
    if not XTOBBasesL.InsertDB(nil) then
    begin
      LastErrorMessage := 'Erreur mise à jour BASES LIGNE';
      V_PGI.IoError := oeUnknown;
      exit;
    end;
    if not XTOBEches.InsertDB(nil) then
    begin
      LastErrorMessage := 'Erreur mise à jour ECHEANCES';
      V_PGI.IoError := oeUnknown;
      exit;
    end;
    if not XTOBAnaP.InsertDBByNivel(false) then
    begin
      LastErrorMessage := 'Erreur mise à jour ANALYTIQUES Pieces';
      V_PGI.IoError := oeUnknown;
      exit;
    end;
    if not XTOBAnaS.InsertDBByNivel(false) then
    begin
      LastErrorMessage := 'Erreur mise à jour ANALYTIQUES Stocks';
      V_PGI.IoError := oeUnknown;
      exit;
    end;
    if not XTOBmetres.InsertDBByNivel(false) then
    begin
      LastErrorMessage := 'Erreur mise à jour METRES';
      V_PGI.IoError := oeUnknown;
      exit;
    end;
    if V_PGI.IoError = oeOk then ValideLesNomen(XTOBNomenclature);
    if (V_PGI.ioerror = oeOk) and (TOBLIGNEFAC.Detail.count > 0) then
    begin
      if not TOBLIGNEFAC.InsertDBByNivel(false) then
      begin
        LastErrorMessage := 'Erreur mise à jour LIGNEFAC';
        V_PGI.IoError := oeUnknown;
        exit;
      end;
    end;
    if (V_PGI.ioerror = oeOk) then ConstitueSituation;
  FINALLY
    EnregistreCreationDoc (V_PGI.ioerror,DebutTitre,LastErrorMessage);
  END;

end;

function TRectifSituation.FindTobLigne (TOBLF : TOB; FromCourant :boolean=false) : TOB;
var II : integer;
    TOBL,ThePiece : TOB;
    prefixe : string;
begin
  result := nil;
  if TOBLF.NomTable = 'LIGNEFAC' then prefixe := 'BLF'
  else if TOBLF.nomtable = 'LIGNE' then prefixe := 'GL';

  if FromCourant then ThePiece := XTOBpiece
                 else ThePiece := TOBpiece;
  for II := 0 to ThePiece.detail.count -1 do
  begin
    TOBL := ThePiece.detail[II]; if TOBL.GetString('GL_TYPELIGNE')<> 'ART' then continue;
    if TOBL.GetInteger('GL_NUMORDRE')=TOBLF.GetInteger(prefixe+'_NUMORDRE') then
    begin
      result := TOBL;
      break;
    end;
  end;
end;


function TRectifSituation.GetNumSituation (TOBPiece : TOB) : integer;
var QQ: TQuery;
    Req : string;
begin
  result := 0;
  req := 'SELECT BST_NUMEROSIT FROM BSITUATIONS WHERE '+
         'BST_NATUREPIECE="'+TOBPIece.GetValue('GP_NATUREPIECEG')+'" AND '+
         'BST_SOUCHE="'+TOBPIece.GetVAlue('GP_SOUCHE')+'" AND '+
         'BST_NUMEROFAC="'+inttoStr(TOBPiece.GetValue('GP_NUMERO'))+'"';
  QQ := OpenSql(Req,true,-1,'',true);
  if not QQ.eof then result := QQ.fields[0].asInteger;
  ferme (QQ);
end;

procedure TRectifSituation.Init;
begin
  // Pièce
  TOBPiece.ClearDetail; TOBPiece.InitValeurs; 
  TOBPieceTrait.ClearDetail;
  TOBBasesL.ClearDetail;
  TOBBases.ClearDetail;
  TOBEches.ClearDetail;
  TOBPorcs.ClearDetail;
  TOBTiers.ClearDetail; TOBTiers.InitValeurs; 
  TOBArticles.ClearDetail;
  TOBAdresses.ClearDetail;
  TOBNomenclature.ClearDetail;
  TOBAffaire.ClearDetail; TOBAffaire.InitValeurs; 
  TOBCPTA.ClearDetail;
  TOBANAP.ClearDetail;
  TOBANAS.ClearDetail;
  TOBOuvrage.ClearDetail;
  TOBOuvragesP.ClearDetail;
  TOBLIENOLE.ClearDetail;
  TOBPieceRG.ClearDetail;
  TOBBasesRG.ClearDetail;
  TOBTimbres.ClearDetail;
  TOBSSTRAIT.ClearDetail;
  TOBMetres.ClearDetail;
  TOBAffaireInterv.ClearDetail;
  TOBACOMPTES.ClearDetail;
  TOBLIGNEFAC.clearDetail;
  //
  XTOBPiece := nil;
  XTOBPieceTrait:= nil;
  XTOBBasesL:= nil;
  XTOBBases:= nil;
  XTOBEches:= nil;
  XTOBPorcs:= nil;
  XTOBTiers:= nil;
  XTOBArticles:= nil;
  XTOBAdresses:= nil;
  XTOBNomenclature:= nil;
  XTOBAffaire:= nil;
  XTOBCPTA:= nil;
  XTOBANAP:= nil;
  XTOBANAS:= nil;
  XTOBOuvrage:= nil;
  XTOBOuvragesP:= nil;
  XTOBLIENOLE:= nil;
  XTOBPieceRG:= nil;
  XTOBBasesRG:= nil;
  XTOBVTECOLL := nil;
  XTOBTimbres:= nil;
  XTOBSSTRAIT:= nil;
  XTOBMetres:= nil;
  XTOBAffaireInterv:= nil;
  XTOBACOMPTES:= nil;
end;

procedure TRectifSituation.PositionneToutModif;
begin
  // Pièce
  TOBPiece.SetAllModifie(true);
  TOBPieceTrait.SetAllModifie(true);
  TOBBasesL.SetAllModifie(true);
  TOBBases.SetAllModifie(true);
  TOBEches.SetAllModifie(true);
  TOBPorcs.SetAllModifie(true);
  TOBTiers.SetAllModifie(true);
  TOBArticles.SetAllModifie(true);
  TOBAdresses.SetAllModifie(true);
  TOBNomenclature.SetAllModifie(true);
  TOBAffaire.SetAllModifie(true);
  TOBCPTA.SetAllModifie(true);
  TOBANAP.SetAllModifie(true);
  TOBANAS.SetAllModifie(true);
  TOBOuvrage.SetAllModifie(true);
  TOBOuvragesP.SetAllModifie(true);
  TOBLIENOLE.SetAllModifie(true);
  TOBPieceRG.SetAllModifie(true);
  TOBBasesRG.SetAllModifie(true);
  TOBTimbres.SetAllModifie(true);
  TOBSSTRAIT.SetAllModifie(true);
  TOBMetres.SetAllModifie(true);
  TOBAffaireInterv.SetAllModifie(true);
  TOBACOMPTES.SetAllModifie(true);
end;

procedure TRectifSituation.ConstitueLigneFacOuvr(TOBL, TOBOuvrage : TOB);
var IndiceNomen,II : integer;
    TOBOO,TOBODF,TOBO : TOB;
begin
  IndiceNomen := TOBL.getValue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;

  TOBOO := TOBOuvrage.detail[IndiceNomen-1];

  for II := 0 TO TOBOO.detail.count -1 do
  begin
    TOBO := TOBOO.detail[II];
    TOBODF := TOB.Create('LIGNEFAC',TOBLIGNEFAC,-1);
    TOBODF.putValue('BLF_NATUREPIECEG',TOBL.GetValue('GL_NATUREPIECEG'));
    TOBODF.putValue('BLF_SOUCHE',TOBL.GetValue('GL_SOUCHE'));
    TOBODF.putValue('BLF_DATEPIECE',TOBL.GetValue('GL_DATEPIECE'));
    TOBODF.putValue('BLF_NUMERO',TOBL.GetValue('GL_NUMERO'));
    TOBODF.putValue('BLF_INDICEG',TOBL.GetValue('GL_INDICEG'));
    TOBODF.putValue('BLF_NUMORDRE',0);
    TOBODF.putValue('BLF_UNIQUEBLO',TOBO.GetValue('BLO_UNIQUEBLO'));
    TOBODF.putValue('BLF_MTMARCHE',TOBO.GetValue('BLF_MTMARCHE'));
    TOBODF.putValue('BLF_MTDEJAFACT',TOBO.GetValue('BLF_MTDEJAFACT'));
    TOBODF.putValue('BLF_QTEMARCHE',TOBO.GetValue('BLF_QTEMARCHE'));
    TOBODF.putValue('BLF_QTEDEJAFACT',TOBO.GetValue('BLF_QTEDEJAFACT'));
    TOBODF.putValue('BLF_MTCUMULEFACT',TOBO.GetValue('BLF_MTCUMULEFACT'));
    TOBODF.putValue('BLF_MTSITUATION',TOBO.GetValue('BLF_MTSITUATION'));
    TOBODF.putValue('BLF_QTECUMULEFACT',TOBO.GetValue('BLF_QTECUMULEFACT'));
    TOBODF.putValue('BLF_QTESITUATION',TOBO.GetValue('BLF_QTESITUATION'));
    TOBODF.putValue('BLF_POURCENTAVANC',TOBO.GetValue('BLF_POURCENTAVANC'));
    TOBODF.putValue('BLF_POURCENTAVANC',TOBO.GetValue('BLF_POURCENTAVANC'));
    TOBODF.putValue('BLF_MTPRODUCTION',0);
    TOBODF.SetAllModifie (true);
  end;

end;

procedure TRectifSituation.ConstitueLigneFac;
var II : integer;
    TOBL,TOBLF : TOB;
begin
  if pos(XTOBpiece.getString('GP_NATUREPIECEG'),'FBT;DAC;FBP;BAC') = 0 then exit;
  //
  for II := 0 to XTOBPiece.detail.count -1 do
  begin
    TOBL:=XTOBPiece.Detail[II] ;
    if IsSousDetail(TOBL) then continue;
    if (TOBL.getValue('BLF_MTMARCHE') <> 0) or (TOBL.getValue('BLF_MTSITUATION')<>0) then
    begin
      TOBLF := TOB.Create ('LIGNEFAC',TOBLIGNEFAC,-1);
      SetLigneFacture (TOBL,TOBLF);
      if IsOuvrage(TOBL) then ConstitueLigneFacOuvr(TOBL, XTOBOuvrage);
    end;
  end;
end;

procedure TRectifSituation.ConstitueNewFacture;
var newDoc : r_cledoc;
    NewNum : integer;
begin
  FillChar(NewDoc,sizeof(NewDoc),#0);
  NewDoc.Naturepiece  := TFFacture(XX).CleDoc.NaturePiece;
  NewDoc.DatePiece := NewDate;
  NewDoc.Souche       := GetSoucheG(NewDoc.NaturePiece, TOBPiece.getString('GP_ETABLISSEMENT'), TOBPiece.getString('GP_DOMAINE'));
  NewNum := SetNumberAttribution(NewDoc.Naturepiece,NewDoc.Souche, NewDoc.DatePiece,1);
  if NewNum = 0 then
  begin
    LastErrorMessage := 'Erreur : N° de compteur des Factures';
    V_PGI.ioerror := oeUnknown;
    exit;
  end else
  begin
    NewDoc.NumeroPiece := NewNum ;
  end;
  MajFromCleDoc (TOBpiece,NewDoc);
end;

procedure TRectifSituation.ConstitueSituation;

  function GetLastIndice (ssAffaire : String; NumeroSit : integer) : integer;
  var req : string;
      QQ : Tquery;
  begin
    result := 0;
    Req:='SELECT BST_INDICESIT FROM BSITUATIONS WHERE '+
        'BST_SSAFFAIRE="'+ ssAffaire + '" AND '+
        'BST_NUMEROSIT='+ IntToStr(NumeroSit)+' '+
        'ORDER BY BST_INDICESIT DESC';
    QQ:=OpenSQL(Req,TRUE,1,'',true);
    if not QQ.eof then
    begin
      result := QQ.Fields[0].AsInteger;
    end;
    ferme (QQ);
  end;
  
var TOBSIT : TOB;
    XD,XP,TXD,TXP,TTC : double;
begin
  if not InAvancement then exit;
  if pos(XTOBpiece.getString('GP_NATUREPIECEG'),'FBT;DAC;FBP;BAC') = 0 then exit;
  TOBSIT := TOB.Create ('BSITUATIONS',nil,-1);
  TRY
    GetMontantRG (XTOBPieceRG,XTOBBasesRG,XD,XP,DEV,True,True);
    GetcumultaxesRG (XTOBBasesRG,XTOBPieceRG,TXD,TXP,DEV);
    TOBSIT.SetString('BST_NATUREPIECE',XTOBPiece.getString('GP_NATUREPIECEG'));
    TOBSIT.SetString('BST_SOUCHE',XTOBPiece.getString('GP_SOUCHE'));
    TOBSIT.SetInteger('BST_NUMEROFAC',XTOBPiece.getInteger('GP_NUMERO'));
    TOBSIT.SetInteger('BST_NUMEROSIT',NumSituation);
    TOBSIT.putValue('BST_DATESIT',XTOBPiece.GetValue('GP_DATEPIECE'));
    TOBSIT.putValue('BST_AFFAIRE',XTOBPiece.GetValue('GP_AFFAIRE'));
    TOBSIT.putValue('BST_SSAFFAIRE',XTOBPiece.GetValue('GP_AFFAIREDEVIS'));
    TTC := XTOBPiece.GetValue('GP_TOTALTTCDEV') - XD - TXD;
    TOBSIt.putvalue('BST_MONTANTHT',XTOBPiece.GetValue('GP_TOTALHTDEV'));
    TOBSIt.putvalue('BST_MONTANTTVA',TTC - XTOBPiece.GetValue('GP_TOTALHTDEV'));
    TOBSIt.putvalue('BST_MONTANTTTC',TTC);
    TOBSIT.PutValue('BST_MONTANTREGL',0);
    TOBSIT.PutValue('BST_MONTANTACOMPTE',0);
    TOBSIT.PutValue('BST_VIVANTE','X');
    TOBSIT.PutValue('BST_INDICESIT',GetLastIndice(XTOBPiece.GetValue('GP_AFFAIREDEVIS'),NumSituation)+1);
    TOBSit.SetAllModifie (true);
    if not TOBSIT.InsertDB(nil) then
    begin
      LastErrorMessage := 'Erreur Ecriture Situation';
      V_PGI.ioerror := OeUnknown;
    end;
  FINALLY
    TOBSIT.free;
  END;
end;

procedure TRectifSituation.PositionneAvancementDevis;


  function GetLigneDevisOrigine (RefOrigine : string) : TOB;
  var cledoc : r_cledoc;
      Req : string;
      QQ : TQuery;
  begin
    result := nil;
    DecodeRefPiece (RefOrigine,cledoc);
    Req := 'SELECT GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE,GL_NUMORDRE,GL_QTEPREVAVANC,Gl_QTESIT,GL_POURCENTAVANC '+
           'FROM LIGNE '+
           'WHERE '+WherePiece(cledoc,ttdLigne,true,true);
    QQ := OpenSql (REq,true,1,'',true);
    if not QQ.eof then
    begin
      result := TOB.Create ('LIGNE',nil,-1);
      TOB(result).SelectDB ('',QQ);
    end;
    ferme (QQ);
  end;

  function GetLignefacOrigine (RefOrigine : string; UniqueBlo : integer = 0) : TOB;
  var cledoc : r_cledoc;
      Req : string;
      QQ : TQuery;
  begin
    result := nil;
    DecodeRefPiece (RefOrigine,cledoc);
    if UniqueBlo = 0 then
    begin
      Req := 'SELECT * '+
             'FROM LIGNEFAC '+
             'WHERE '+WherePiece(cledoc,ttdLignefac,true,true);
    end else
    begin

      Req := 'SELECT * '+
             'FROM LIGNEFAC '+
             'WHERE '+
             'BLF_NATUREPIECEG="'+cledoc.NaturePiece+'" AND '+
             'BLF_SOUCHE="'+cledoc.Souche+'" AND '+
             'BLF_NUMERO='+InttoStr(cledoc.NumeroPiece)+' AND '+
             'BLF_INDICEG='+InttoStr(cledoc.Indice)+' AND '+
             'BLF_UNIQUEBLO='+InttoStr(UniqueBlo);
    end;
    QQ := OpenSql (REq,true,1,'',true);
    if not QQ.eof then
    begin
      result := TOB.Create ('LIGNEFAC',nil,-1);
      TOB(result).SelectDB ('',QQ);
    end;
    ferme (QQ);
  end;

  procedure PositionneInfoLigneFacOrigine (TOBLF,TOBLFO: TOB);
  begin
    if TOBLFO = nil then exit; // pas trouvé ????
    if InAvancement then
    begin
      TOBLFO.SetDouble('BLF_MTCUMULEFACT',TOBLF.GetDouble('BLF_MTCUMULEFACT'));
      TOBLFO.SetDouble('BLF_MTDEJAFACT',TOBLF.GetDouble('BLF_MTDEJAFACT'));
      TOBLFO.SetDouble('BLF_QTECUMULEFACT',TOBLF.GetDouble('BLF_QTECUMULEFACT'));
      TOBLFO.SetDouble('BLF_QTEDEJAFACT',TOBLF.GetDouble('BLF_QTEDEJAFACT'));
    end else
    begin
      TOBLFO.SetDouble('BLF_MTCUMULEFACT',TOBLF.GetDouble('BLF_MTCUMULEFACT'));
      TOBLFO.SetDouble('BLF_QTECUMULEFACT',TOBLF.GetDouble('BLF_QTECUMULEFACT'));
    end;
    TOBLFO.SetDouble('BLF_POURCENTAVANC',TOBLF.GetDouble('BLF_POURCENTAVANC'));
    TOBLFO.UpdateDB;
  end;

  procedure PositionneInfoDevisOrigine (TOBLF,TOBLD: TOB);
  begin
    if InAvancement then
    begin
      TOBLD.SetDouble('GL_QTESIT',TOBLF.GetDouble('BLF_QTECUMULEFACT'));
      TOBLD.SetDouble('GL_QTEPREVAVANC',TOBLF.GetDouble('BLF_QTECUMULEFACT'));
      TOBLD.SetDouble('GL_POURCENTAVANC',TOBLF.GetDouble('BLF_POURCENTAVANC'));
    end else
    begin
      TOBLD.SetDouble('GL_QTESIT',TOBLF.GetDouble('BLF_QTECUMULEFACT'));
      TOBLD.SetDouble('GL_QTEPREVAVANC',TOBLF.GetDouble('BLF_QTECUMULEFACT'));
      TOBLD.SetDouble('GL_POURCENTAVANC',TOBLF.GetDouble('BLF_POURCENTAVANC'));
    end;
    TOBLD.UpdateDB;
  end;


var II : integer;
    TOBL,TOBLFO,TOBLD,TOBLF : TOB;
    StOrig : string;
begin
  // mise en place de l'avancement sur le devis d'origine
  if not InAvancement then exit;
  if pos(XTOBpiece.getString('GP_NATUREPIECEG'),'FBT;DAC;FBP;BAC') = 0 then exit;
  //
  TOBLD := nil;
  TOBLFO := nil;
  for II := 0 to TOBLignefac.detail.count -1 do
  begin
    TOBLF:=TOBLignefac.Detail[II] ;
    TOBL := FindTobLigne (TOBLF,true); // ligne dans XTOBpiece
    if TOBl = nil then continue;
    StOrig:=TOBL.GetValue('GL_PIECEORIGINE') ;
    if TOBLF.GetInteger('BLF_UNIQUEBLO') = 0 then
    begin
      TOBLD := GetLigneDevisOrigine (stOrig);
      TOBLFO := GetLignefacOrigine (stOrig);
      PositionneInfoDevisOrigine (TOBLF,TOBLD);
      PositionneInfoLigneFacOrigine (TOBLF,TOBLFO);
    end else
    begin
      TOBLFO := GetLignefacOrigine (stOrig,TOBLF.GetInteger('BLF_UNIQUEBLO'));
      PositionneInfoLigneFacOrigine (TOBLF,TOBLFO);
    end;
    if TOBLD <> nil then FreeAndNil(TOBLD);
    if TOBLFO <> nil then FreeAndNil(TOBLFO);
  end;
end;


procedure TRectifSituation.RecalculePieceCourante;
var II : integer;
begin
  ReinitMontantPieceTrait (XTOBPiece,XTOBaffaire,XTOBPieceTrait);
  ZeroFacture (XTOBpiece);
  ZeroMontantPorts (XTOBPorcs);
  XTOBBases.ClearDetail;
  XTOBBasesL.ClearDetail;
  PutValueDetail(XTOBPiece, 'GP_RECALCULER', 'X'); // positionne le recalcul du document
  for II := 0 to XTOBPiece.detail.Count -1 do ZeroLigneMontant(XTOBPiece.detail[II]);
  XTOBVTECOLL.ClearDetail;
  CalculFacture(XTOBAffaire,XTOBPiece,XTOBPieceTrait,XTOBSSTRAIT,XTOBOUvrage, XTOBOuvragesP,XTOBBases, XTOBBasesL,XTOBTiers, XTOBArticles, XTOBPorcs, XTOBPieceRG, XTOBBasesRG,XTOBVTECOLL, DEV, false, taModif,false);
  if XTOBTimbres.Detail.Count > 0 then
  begin
    GereEcheancesGC(XTOBPiece, XTOBTiers, XTOBEches, XTOBAcomptes, XTOBPieceRG,XTOBPieceTrait,XTOBPorcs,taCreat, DEV, False);
    CalculeTimbres (XTOBTimbres,XTOBpiece,XTOBEches);
  end;
  CalculeReglementsIntervenants(XTOBSSTRAIT, XTOBPiece,XTOBPIECERG,XTOBAcomptes,XTOBPorcs,XTOBPiece.getValue('GP_AFFAIRE'),XTOBPieceTrait,DEV);
  GereEcheancesGC(XTOBPiece, XTOBTiers, XTOBEches, XTOBAcomptes, XTOBPieceRG,XTOBPieceTrait,XTOBPorcs,tacreat, DEV, False);
end;

procedure TRectifSituation.RecalculePiece;
var II : integer;
begin
  ReinitMontantPieceTrait (TOBPiece,TOBaffaire,TOBPieceTrait);
  ZeroFacture (TOBpiece);
  ZeroMontantPorts (TOBPorcs);
  TOBBases.ClearDetail;
  TOBBasesL.ClearDetail;
  PutValueDetail(TOBPiece, 'GP_RECALCULER', 'X'); // positionne le recalcul du document
  for II := 0 to TOBPiece.detail.Count -1 do ZeroLigneMontant(TOBPiece.detail[II]);
  TOBVTECOLL.ClearDetail;
  CalculFacture(TOBAffaire,TOBPiece,TOBPieceTrait,TOBSSTRAIT,TOBOUvrage, TOBOuvragesP,TOBBases, TOBBasesL,TOBTiers, TOBArticles, TOBPorcs, TOBPieceRG, TOBBasesRG, TOBVTECOLL, DEV, false, TaModif,false);
  if TOBTimbres.Detail.Count > 0 then
  begin
    GereEcheancesGC(TOBPiece, TOBTiers, TOBEches, TOBAcomptes, TOBPieceRG,TOBPieceTrait, TOBPorcs,taCreat, DEV, False);
    CalculeTimbres (TOBTimbres,TOBpiece,TOBEches);
  end;
  CalculeReglementsIntervenants(TOBSSTRAIT, TOBPiece,TOBPIECERG,TOBAcomptes,TOBPorcs,TOBPiece.getValue('GP_AFFAIRE'),TOBPieceTrait,DEV);
  GereEcheancesGC(TOBPiece, TOBTiers, TOBEches, TOBAcomptes, TOBPieceRG,TOBPieceTrait,TOBPorcs,taCreat, DEV, False);
end;

procedure TRectifSituation.RectifieAvancement;

  function FindLigneOuv (TOBO,TOBOUS : TOB) : TOB;
  var II : integer;
      TOBOOS : TOB;
  begin
    result := nil;
    For II := 0 to TOBOUS.detail.count -1 do
    begin
      TOBOOS := TOBOUS.detail[II];
      if TOBOOS.GetInteger('BLO_UNIQUEBLO')=TOBO.GetInteger('BLO_UNIQUEBLO') then
      begin
        result := TOBOOS;
        exit;
      end;
    end;
  end;

  procedure SetAvancementLigne (TOBLS,TOBL : TOB);
  begin
    TOBLS.setDouble('BLF_QTEDEJAFACT', TOBL.getDouble('BLF_QTEDEJAFACT')+TOBL.getDouble('BLF_QTESITUATION'));
    TOBLS.setDouble('BLF_MTDEJAFACT', TOBL.getDouble('BLF_MTDEJAFACT')+TOBL.getDouble('BLF_MTSITUATION'));
    // recalcule de la situation
    TOBLS.setDouble('BLF_MTSITUATION', TOBLS.getDouble('BLF_MTCUMULEFACT')-TOBLS.getDouble('BLF_MTDEJAFACT'));
    TOBLS.setDouble('BLF_QTESITUATION', TOBLS.getDouble('BLF_QTECUMULEFACT')-TOBLS.getDouble('BLF_QTEDEJAFACT'));
    IF TOBLS.getDouble('BLF_MTMARCHE') = 0 then TOBLS.setDouble('BLF_POURCENTAVANC',100)
                                           else  TOBLS.setDouble('BLF_POURCENTAVANC', Arrondi(TOBLS.getDouble('BLF_MTCUMULEFACT')/TOBLS.getDouble('BLF_MTMARCHE')*100,2));
  end;

  procedure RectifieAvancementOuvrage (TOBL,TOBLS,TOBOuvrageRef,TOBOuvrageSuiv : TOB);
  var IndiceNomen,IndiceNomS,II : integer;
      TOBOUV,TOBOUS,TOBO,TOBOS : TOB;
  begin
    IndiceNomen := TOBL.GetInteger('GL_INDICENOMEN');
    IndiceNoms := TOBLS.GetInteger('GL_INDICENOMEN');
    TOBOUV := TOBOuvrageRef.detail[IndiceNomen-1];
    TOBOUS := TOBOuvrageSuiv.detail[IndiceNomen-1];
    for II := 0 to TOBOUV.detail.count -1 do
    begin
      TOBO := TOBOUV.detail[II];
      TOBOS := FindLigneOuv(TOBO,TOBOUS);
      if TOBOS = nil then continue;
      SetAvancementLigne (TOBOS,TOBO);
    end;
  end;

var II : integer;
    TOBL,TOBLS : TOB;
    QteDejafacture,MtDejafacture : double;
begin
  // Application de l'avancement de la situation modifié sur la nouvelle situation (tobpiece)
  For II := 0 to TFFacture(XX).LaPieceCourante.detail.count -1 do
  begin
    TOBL := TFFacture(XX).LaPieceCourante.detail[II];
    if TOBL.GetString('GL_TYPELIGNE') <> 'ART' then continue;
    if not TOBL.IsFieldModified('GL_MONTANTHTDEV') then continue;
    TOBLS := FindTobLigne(TOBL); // lecture sur TOBpiece
    if TOBLS = nil then continue;
    SetAvancementLigne (TOBLS,TOBL);
    if IsOuvrage(TOBL) then
    begin
      RectifieAvancementOuvrage (TOBL,TOBLS,TFFacture(XX).TheTOBOuvrage,TOBOuvrage);
    end;
  end;
end;

procedure TRectifSituation.RectifieSituation;
var Splash : TFsplashScreen;
begin
  ModeDebug := V_PGI.SAV;
  Splash := TFsplashScreen.Create (application);
  Splash.Label1.Caption  := TraduireMemoire('Génération des pièces rectificatives...');
  Splash.Show;
  Splash.BringToFront;
  Splash.refresh;
  TRY
    ModeGeneration := RenvoieTypeFact(TFFACTURE(XX).LaPieceCourante.GetValue('GP_AFFAIREDEVIS'));
    InAvancement := (Pos(ModeGeneration,'AVA;DAC;')>0);
    DEV := TFFacture(XX).DEV;
    //
    lastSituation := true;
    NumSituation := 0;
    // ---------------
    nextCledoc.NaturePiece := '';
    nextCledoc.NumeroPiece := 0;
    // --------------------
    if not DemandeDatesFacturation (NewDate) then exit;
    //
    Init;
    if (TFfacture(XX).Cledoc.NaturePiece = '') or (TFfacture(XX).Cledoc.NumeroPiece =0) then exit;

    ChargeLaPiece (TFfacture(XX).Cledoc);
    //
    if (InAvancement)  then NumSituation := GetNumSituation (TOBPiece);
    if (InAvancement)  then lastSituation := DerniereSituation (TOBpiece,true);
    if (InAvancement) and  (not lastSituation) then GetSituationSuivante (TOBpiece,NextCledoc);
    //
    if V_PGI.IOError = OeOk then ConstitueAvoir (NumSituation,TOBpiece.getInteger('GP_NUMERO'),TOBpiece.GetDateTime ('GP_DATEPIECE'));
    if V_PGI.IOError = OeOk then Enregistre;
    if V_PGI.ioError = OeOk then SetPieceMorte (TFfacture(XX).Cledoc); // positionne la piece origine comme morte
    if V_PGI.IOError = OeOk then EnregistreDocumentModifie; // en provenance des elements du TFFACTURE   (nouvelle situation courante)
    //
    if (not lastSituation) then
    begin
      NumSituation := 0;
      // Generation de l'avoir par rapport a la piece suivante
      if V_PGI.IOError = OeOk then Init;
      if V_PGI.IOError = OeOk then ChargeLaPiece (nextCledoc);
      if (InAvancement)  then NumSituation := GetNumSituation (TOBPiece); // recup du numero de situation courant
      if V_PGI.IOError = OeOk then ConstitueAvoir (NumSituation,TOBpiece.getInteger('GP_NUMERO'),TOBpiece.GetDateTime ('GP_DATEPIECE'));
      if V_PGI.IOError = OeOk then Enregistre;
      //
      // Generation de na nouvelle suituation suivante
      NumSituation := 0;
      if V_PGI.IOError = OeOk then Init;
      if V_PGI.IOError = OeOk then ChargeLaPiece (nextCledoc);
      if (InAvancement)  then NumSituation := GetNumSituation (TOBPiece); // recup du numero de situation courant
      if V_PGI.IOError = OeOk then SetPieceMorte (nextCledoc) ; // positionne la piece origine comme morte
      //
      if V_PGI.IOError = OeOk then ConstitueNewFacture;
      if V_PGI.IOError = OeOk then RectifieAvancement;
      if V_PGI.IOError = OeOk then RecalculePiece;
      if V_PGI.IOError = OeOk then Enregistre;
    end else
    begin
      if V_PGI.IOError = OeOk then PositionneAvancementDevis;
    end;
  FINALLY
    Splash.free;
    Splash := nil;
    Validate := (V_PGI.ioerror = oeOk );
    TheTOB := TOBRESULTATS;
    AglLanceFiche ('BTP','BTRECALCPIECE_RAP','','','ERREURSOUPAS');
    TheTOB := Nil;
  END;
end;

procedure TRectifSituation.SetInfoNewDocument;

  procedure ReajusteSousDetail (TOBOUV : TOB; NewDoc : r_cledoc);
  var i : integer;
      TOBB : TOB;
  begin
    for i := 0 to TOBOUV.Detail.Count - 1 do
    begin
      TOBB := TOBOUV.Detail[i];
      MajFromCleDoc(TOBB, NewDoc);
      ReajusteSousDetail (TOBB,NewDoc);
    end;
  end;

var i,J,K : integer;
    TOBB,TOBDA,TOBAA : TOB;
    RefA : string;
    NewDoc : r_cledoc;
begin
  NewDoc := TOB2Cledoc(XTOBpiece);
  RefA:=EncodeRefCPGescom(XTOBPiece) ; if RefA='' then Exit ;
  //
  for i := 0 to XTOBPiece.Detail.Count - 1 do
  begin
    TOBB := XTOBPiece.Detail[i];
    MajFromCleDoc(TOBB, NewDoc);
    for J := 0 to TOBB.detail.count -1 do
    begin
      TOBAA := TOBB.detail[J];
      for K := 0 to TOBAA.detail.count -1 do
      begin
        TOBDA := TOBAA.detail[K];
        TOBDA.PutValue('YVA_IDENTIFIANT',RefA) ;
      end;
    end;
  end;
  for i := 0 to XTOBOuvrage.Detail.Count - 1 do
  begin
    TOBB := XTOBOuvrage.Detail[i];
    TOBB.PutValue('UTILISE','X');
    MajFromCleDoc(TOBB, NewDoc);
    if TOBB.detail.count > 0 then
    begin
      ReajusteSousDetail (TOBB,NewDoc);
    end;
  end;
  for i := 0 to XTOBOuvragesP.Detail.Count - 1 do
  begin
    TOBB := XTOBOuvragesP.Detail[i];
    for J := 0 to  TOBB.detail.count -1 do
    begin
      TOBAA := TOBB.Detail[j];
      MajFromCleDoc(TOBAA, NewDoc);
    end;
  end;
  for i := 0 to XTOBNomenclature.Detail.Count - 1 do
  begin
    TOBB := XTOBNomenclature.Detail[i];
    MajFromCleDoc(TOBB, NewDoc);
    if TOBB.detail.count > 0 then
    begin
      ReajusteSousDetail (TOBB,NewDoc);
    end;
  end;
  for i := 0 to XTOBmetres.Detail.Count - 1 do
  begin
    TOBB := XTOBmetres.Detail[i];
    MajFromCleDoc(TOBB, NewDoc);
  end;
  for i := 0 to XTOBBases.Detail.Count - 1 do
  begin
    TOBB := XTOBBases.Detail[i];
    MajFromCleDoc(TOBB, NewDoc);
  end;
  for i := 0 to XTOBBasesL.Detail.Count - 1 do
  begin
    TOBB := XTOBBasesL.Detail[i];
    for J := 0 to TOBB.detail.count -1 do
    begin
      TOBAA := TOBB.detail[J];
      MajFromCleDoc(TOBAA, NewDoc);
    end;
  end;
  for i := 0 to XTOBEches.Detail.Count - 1 do
  begin
    TOBB := XTOBEches.Detail[i];
    MajFromCleDoc(TOBB, NewDoc);
  end;
  for i := 0 to XTOBPieceRG.Detail.Count - 1 do
  begin
    TOBB := XTOBPieceRG.Detail[i];
    MajFromCleDoc(TOBB, NewDoc);
  end;
  for i := 0 to XTOBBasesRG.Detail.Count - 1 do
  begin
    TOBB := XTOBBasesRG.Detail[i];
    MajFromCleDoc(TOBB, NewDoc);
  end;
  for i := 0 to XTOBPorcs.Detail.Count - 1 do
  begin
    TOBB := XTOBPorcs.Detail[i];
    MajFromCleDoc(TOBB, NewDoc);
  end;
  for i := 0 to XTOBSSTRAIT.Detail.Count - 1 do
  begin
    TOBB := XTOBSSTRAIT.Detail[i];
    MajFromCleDoc(TOBB, NewDoc);
  end;
  for i := 0 to XTOBpieceTrait.Detail.Count - 1 do
  begin
    TOBB := XTOBpieceTrait.Detail[i];
    MajFromCleDoc(TOBB, NewDoc);
  end;
  for i := 0 to XTOBAdresses.Detail.Count - 1 do
  begin
    TOBB := XTOBAdresses.Detail[i];
    MajFromCleDoc(TOBB, NewDoc);
  end;
  for i := 0 to XTOBLienOle.Detail.Count - 1 do
  begin
    TOBB := XTOBLienOle.Detail[i];
    MajFromCleDoc(TOBB, NewDoc);
  end;
end;

procedure TRectifSituation.SetPieceMorte (cledoc : r_cledoc);
var SQl : string;
begin
  Sql := 'UPDATE PIECE SET GP_VIVANTE="-" WHERE '+WherePiece(cledoc,ttdPiece,true);
  if ExecuteSQL(SQL)=0 then
  begin
    LastErrorMessage := 'Erreur Positionnement pièce morte';
    V_PGI.IOError := oeUnknown;
    exit;
  end;
  Sql := 'UPDATE LIGNE SET GL_VIVANTE="-" WHERE '+WherePiece(cledoc,ttdLigne,false);
  if ExecuteSQL(SQL)=0 then
  begin
    LastErrorMessage := 'Erreur Positionnement pièce morte';
    V_PGI.IOError := oeUnknown;
    exit;
  end;
  if NumSituation >0 then
  begin
    Sql := 'UPDATE BSITUATIONS SET BST_VIVANTE="-" WHERE '+WherePiece(cledoc,ttdSit,true);
    if ExecuteSQL(SQL)=0 then
    begin
      LastErrorMessage := 'Erreur Positionnement pièce morte';
      V_PGI.IOError := oeUnknown;
      exit;
    end;
  end;
end;

procedure TRectifSituation.ToutAllouer;
begin
  TOBRESULTATS := TOB.Create ('LE COMPTE RENDU',nil,-1);
  // Pièce
  TOBPiece := TOB.Create('PIECE', nil, -1);
  // Modif BTP
  AddLesSupEntete(TOBPiece);
  // ---
  TOBPieceTrait := TOB.Create ('LES LIGNES EXTRENALISE',nil,-1);
  TOBBasesL := TOB.Create('LES BASES LIGNES', nil, -1);
  TOBBases := TOB.Create('BASES', nil, -1);
  TOBEches := TOB.Create('Les ECHEANCES', nil, -1);
  TOBPorcs := TOB.Create('PORCS', nil, -1);
  // Fiches
  TOBTiers := TOB.Create('TIERS', nil, -1);
  TOBTiers.AddChampSup('RIB', False);
  TOBArticles := TOB.Create('ARTICLES', nil, -1);
  // Adresses
  TOBAdresses := TOB.Create('LESADRESSES', nil, -1);
  if GetParamSoc('SO_GCPIECEADRESSE') then
  begin
    TOB.Create('PIECEADRESSE', TOBAdresses, -1); {Livraison}
    TOB.Create('PIECEADRESSE', TOBAdresses, -1); {Facturation}
  end else
  begin
    TOB.Create('ADRESSES', TOBAdresses, -1); {Livraison}
    TOB.Create('ADRESSES', TOBAdresses, -1); {Facturation}
  end;
  // Divers
  TOBNomenclature := TOB.Create('NOMENCLATURES', nil, -1);
  // Affaires
  TOBAffaire   := TOB.Create('AFFAIRE', nil, -1);
  // Comptabilité
  TOBCPTA      := CreerTOBCpta;
  TOBANAP      := TOB.Create('', nil, -1);
  TOBANAS      := TOB.Create('', nil, -1);
  TOBOuvrage   := TOB.Create('OUVRAGES', nil, -1);
  TOBOuvragesP := TOB.Create('LES OUVRAGES PLAT', nil, -1);
  TOBLIENOLE   := TOB.Create('LIENS', nil, -1);
  TOBPieceRG   := TOB.create('PIECERRET', nil, -1);
  TOBBasesRG   := TOB.create('BASESRG', nil, -1);
  // --
  TOBTimbres   := TOB.Create ('LES TIMBRES',nil,-1);
  TOBSSTRAIT   := TOB.Create ('LES SOUS TRAITS',nil,-1);
  TOBMetres    := TOB.Create('LES METRES',nil,-1);
  TOBAffaireInterv := TOB.Create ('LES CO-SOUSTRAITANTS',nil,-1);
  TOBACOMPTES := TOB.create('LES ACOMPTES',nil,-1);
  TOBLIGNEFAC := TOB.Create ('LES LIGNES FAC',nil,-1); 
  TOBVTECOLL := TOB.Create ('LES VENTIL COLL',nil,-1);
end;

procedure TRectifSituation.ToutLiberer;
begin
  // Pièce
  TOBPiece.free;
  TOBPieceTrait.free;
  TOBBasesL.free;
  TOBBases.free;
  TOBEches.free;
  TOBPorcs.free;
  TOBTiers.free;
  TOBArticles.free;
  TOBAdresses.free;
  TOBNomenclature.free;
  TOBAffaire.free;
  TOBCPTA.free;
  TOBANAP.free;
  TOBANAS.free;
  TOBOuvrage.free;
  TOBOuvragesP.free;
  TOBLIENOLE.free;
  TOBPieceRG.free;
  TOBBasesRG.free;
  TOBTimbres.free;
  TOBSSTRAIT.free;
  TOBMetres.free;
  TOBAffaireInterv.free;
  TOBACOMPTES.free;
  TOBLIGNEFAC.free;
  TOBRESULTATS.free;
  TOBVTECOLL.Free;
end;

procedure TRectifSituation.EnregistreCreationDoc(Ioerror : TIOErr; DebutTitre,LastErrorMessage : string);
var TT : TOB;
begin
  TT := TOB.Create ('UNE LIGNE',TOBRESULTATS,-1);
  if Ioerror=OeOk then  TT.AddChampSupValeur ('RAPPORT',DebutTitre + ' --> OK')
                  else  TT.AddChampSupValeur ('RAPPORT',DebutTitre + ' --> ERREUR ('+lastErrorMessage+')');
end;

end.
