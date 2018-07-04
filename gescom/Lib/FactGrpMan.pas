unit FactGrpMan ;

interface

Uses HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}  DB, Fe_Main,
      {$IFDEF V530} EdtDOC,{$ELSE} EdtRdoc, {$ENDIF}
{$ENDIF}
{$IFDEF BTP}
     FactOuvrage,factspec,BTPUtil,FactGrpBtp,UtilPhases,
{$ENDIF}
     SysUtils, Dialogs, SaisUtil, UtilPGI, AGLInit, FactUtil, UtilSais,
     EntGC, Classes, HMsgBox, SaisComm, FactComm, ed_tools, FactNomen, UtilGrp,
     FactCpta, FactCalc, ParamSoc, UtilArticle,
{$IFDEF AFFAIRE}
     FactActivite,  AffaireUtil,
{$ENDIF}
     FactTOB, FactPiece, factrg, FactAdresse, FactTiers, FactLotSerie,BTStructChampSup,uEntCommun;

Function  RegroupeLesPiecesMan ( TOBSource : TOB ; NewNat : String ; Eclate,DeGroupeRemise,AvecComment : boolean ; NewDate : TDateTime = 0 ) : R_CleDoc;

implementation
uses UtilTOBPiece,FactTimbres;

Var NewNature,VenteAchat : String ;
    NewNum,IndiceNomen,PremNum,LastNum,NbP,EvDep,IndiceSerie : integer ;
    ClePiece :R_CleDoc;
    DEV    : RDEVISE ;
    TOBGenere,TOBArticles,TOBTiers,TOBRef,TOBEches,TOBNomenclature,TOBAdresses,TOBSerie,
    TOBCpta,TOBAnaP,TOBAnaS,TOBBases,TOBBasesL,TOBAcomptes,TOBPorcs,TOBComms, TOBCatalogu : TOB ;
    NeedVisa,Commentaires : boolean ;
{$IFDEF AFFAIRE}
    GereActivite,DelActivite : boolean;
{$ENDIF} // AFFAIRE
    MontantVisa  : double ;
    ChampsRupt,PiecesG             : TStrings ;
    LaNewDate                             : TDateTime ;
    // modif BTP
    TOBOuvrage,TOBPieceRG,TOBBasesRG,TOBLienOle,TOBOuvragesP : TOB;
    IndiceRg : integer;
    TobLigneTarif : TOB ;   

Type T_GenVal = Class
                TOBSource : TOB ;
                CodeTiers : String ;
                Eclate,DeGroupeRemise : Boolean ;
                Procedure G_GenereParTiers ;
                End ;

{============================= MANIPULATION DES TOBS ==========================}
Procedure G_CreerLesTOB ;
BEGIN
TOBGenere:=TOB.Create('PIECE',Nil,-1) ;
// Modif BTP
AddLesSupEntete (TOBGenere);
// ---
TOBRef:=TOB.Create('LIGNE',Nil,-1) ;
TOBArticles:=TOB.Create('',Nil,-1) ;
TOBCataLogu:=TOB.Create('',Nil,-1) ;
TOBBases:=TOB.Create('BASES',Nil,-1) ;
TOBBasesL:=TOB.Create('LES BASES LIGNE',Nil,-1) ;
TOBTiers:=TOB.Create('TIERS',Nil,-1) ; TOBTiers.AddChampSup('RIB',False) ;
TOBEches:=TOB.Create('Les ECHEANCES',Nil,-1) ;
TOBSerie:=TOB.Create('',Nil,-1) ;
TOBAcomptes:=TOB.Create('',Nil,-1) ;
TOBPorcs:=TOB.Create('',Nil,-1) ;
TOBNomenclature:=TOB.Create('NOMENCLATURES',Nil,-1) ;
TOBComms:=TOB.Create('COMMERCIAUX',Nil,-1) ;
TOBCpta:=CreerTOBCpta ;
TOBAnaP:=TOB.Create('',Nil,-1) ; TOBAnaS:=TOB.Create('',Nil,-1) ;
TOBAdresses:=TOB.Create('',Nil,-1);
PiecesG:=TStringList.Create ;
ChampsRupt:=TStringList.Create ;
// Modif BTP
TOBOuvrage := TOB.create ('OUVRAGES',nil,-1);
TOBOuvragesP := TOB.create ('OUVRAGES',nil,-1);
TOBLienOLE := TOB.create ('LIENSOLE',nil,-1);
TOBPIECERG := TOB.create ('RETENUES',nil,-1);
TOBBasesRG := TOB.create ('BASESRG',nil,-1);
TobLigneTarif := Tob.Create('_LIGNETARIF',nil,-1);
PutLesSupEnteteLigneTarif(TobGenere, TobLigneTarif, nil);
PutTobPieceInTobLigneTarif(TobGenere, TobLigneTarif);
END ;

Procedure G_FreeLesTOB ;
BEGIN
TOBGenere.Free ; TOBRef.Free ; TOBArticles.Free ;
TOBBases.Free ; TOBbasesL.free; TOBTiers.Free ; TOBNomenclature.Free ;
TOBAdresses.Free ; TOBCataLogu.Free; TOBSerie.Free;
TOBCpta.Free ; TOBAnaP.Free ; TOBAnaS.Free ; TOBEches.Free ;
TOBAcomptes.Free ; TOBPorcs.Free ; TOBComms.Free ;
PiecesG.Free ;
ChampsRupt.Clear ; ChampsRupt.Free ;
// Modif BTP
TOBOuvrage.free; TOBOuvragesP.free; TOBLIENOLE.free; TOBPIECERG.free; TOBBasesRG.free;
// ----
TobLigneTarif.Free;
END ;

Procedure G_ChargeChampsRupt ;
BEGIN
{$IFDEF BTP}
  DefiniChampRuptBtp (ChampsRupt);
{$ELSE}
Q:=OpenSQL('SELECT GPR_CONDGRP FROM PARPIECEGRP WHERE GPR_NATUREPIECEG="'+NewNature+'"',True,-1, '', True) ;
if Not Q.EOF then
   BEGIN
{$IFDEF EAGLCLIENT}
   ChampsRupt.Text:=Q.FindField('GPR_CONDGRP').AsString ;
{$ELSE}
   if Not TMemoField(Q.FindField('GPR_CONDGRP')).IsNull then
      ChampsRupt.Assign(TMemoField(Q.FindField('GPR_CONDGRP'))) ;
{$ENDIF}
   END else
   BEGIN
   with ChampsRupt do begin
    Add('PIECE;GP_TIERS;=;');
    Add('PIECE;GP_DEVISE;=;');
    Add('PIECE;GP_TAUXDEV;=;');
    Add('PIECE;GP_COTATION;=;');
    Add('PIECE;GP_SAISIECONTRE;=;');
    Add('PIECE;GP_REGIMETAXE;=;');
    Add('PIECE;GP_ESCOMPTE;=;');
    Add('PIECE;GP_FACTUREHT;=;');
    Add('PIECE;GP_TIERSPAYEUR;=;');
    END;
   END ;
Ferme(Q) ;
{$ENDIF}
END ;

Procedure G_InitGenere ;
Var Q : TQuery ;
BEGIN
NeedVisa:=(GetInfoParPiece(NewNature,'GPP_VISA')='X') ;
MontantVisa:=GetInfoParPiece(NewNature,'GPP_MONTANTVISA') ;
VenteAchat:=GetInfoParPiece(NewNature,'GPP_VENTEACHAT') ;
Premnum:=-1 ; LastNum:=-1 ; NbP:=0 ;
{Journal des évènements}
Q:=OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT',True,-1, '', True) ;
if Not Q.EOF then EvDep:=Q.Fields[0].AsInteger ;
Ferme(Q) ;
Inc(EvDep) ;
Q:=OpenSQL('SELECT * FROM JNALEVENT WHERE GEV_TYPEEVENT="GEN" AND GEV_NUMEVENT=-1',False) ;
Q.Insert ; InitNew(Q) ;
Q.FindField('GEV_NUMEVENT').AsInteger:=EvDep ;
Q.FindField('GEV_TYPEEVENT').AsString:='GEN' ;
Q.FindField('GEV_LIBELLE').AsString:=Copy('Génération de '+RechDom('GCNATUREPIECEG',NewNature,False),1,35) ;
Q.FindField('GEV_DATEEVENT').AsDateTime:=Date ;
Q.FindField('GEV_UTILISATEUR').AsString:=V_PGI.User ;
Q.FindField('GEV_ETATEVENT').AsString:='ERR' ;
Q.Post ;
Ferme(Q) ;
{Gestion de la maj de l'activité}
{$IFDEF AFFAIRE}
EtudieActivite(NewNature,taCreat,false,GereActivite,DelActivite);
{$endif}
END ;

{================================ NOMENCLATURES ===============================}
Procedure G_LoadNomenLigne ( TOBL : TOB ) ;
Var i,OldL,Lig,MaxNiv,k,Niv : integer ;
    TOBLN,TOBNomen,TOBGroupeNomen,TOBLoc : TOB ;
    Q    : TQuery ;
    CleDoc : R_CleDoc ;
BEGIN
///// NE PAS TOUCHER ///////////////////////////
OldL:=-1 ; //IndiceNomen:=1 ; MR : modif déjà effectué en version 535 annulée depuis !!!!!!!!!
/////////////////////////////////////////////////
if ((TOBL.GetValue('GL_TYPEARTICLE')<>'NOM') or (TOBL.GetValue('GL_TYPENOMENC')<>'ASS')) then Exit ;
TOBNomen:=TOB.Create('',Nil,-1) ; CleDoc:=TOB2CleDoc(TOBL) ;
Q:=OpenSQL('SELECT * FROM LIGNENOMEN WHERE '+WherePiece(CleDoc,ttdNomen,False)+' AND GLN_NUMLIGNE='+IntToStr(TOBL.GetValue('GL_NUMLIGNE')),True,-1, '', True) ;
TOBNomen.LoadDetailDB('LIGNENOMEN','','',Q,True,False) ;
Ferme(Q) ;
if TOBNomen.Detail.Count=0 then
   BEGIN
   MessageAlerte('ATTENTION : Il existe des nomenclatures sans décomposition') ;
   V_PGI.IoError:=oeUnknown;
   END;
for i:=0 to TOBNomen.Detail.Count-1 do
    BEGIN
    TOBLN:=TOBNomen.Detail[i] ;
    Lig:=TOBLN.GetValue('GLN_NUMLIGNE') ;
    if OldL<>Lig then
       BEGIN
       TOBGroupeNomen:=TOB.Create('',TOBNomenclature,-1) ; MaxNiv:=-1 ;
       TOBGroupeNomen.AddChampSup('UTILISE',False) ; TOBGroupeNomen.PutValue('UTILISE','-') ;
       for k:=i to TOBNomen.Detail.Count-1 do
           BEGIN
           TOBLoc:=TOBNomen.Detail[k] ;
           if TOBLoc.GetValue('GLN_NUMLIGNE')<>Lig then Break ;
           Niv:=TOBLoc.GetValue('GLN_NIVEAU') ; if Niv>MaxNiv then MaxNiv:=Niv ;
           END ;
       CreerLesTOBNomen(TOBGroupeNomen,TOBNomen,TOBArticles,Lig,MaxNiv,i, TobL.GetValue ('GL_DEPOT')) ; // DBR : Depot unique chargé
       END ;
    OldL:=Lig ;
    END ;
TOBL.PutValue('GL_INDICENOMEN',IndiceNomen) ; Inc(IndiceNomen) ;
TOBNomen.Free ;
END ;

{================================ ADRESSES =================================}
Procedure G_ChargeLesAdresses ( TOBPiece : TOB ) ;
var i_ind, NumL, NumF : integer ;
BEGIN
for i_ind:=TOBAdresses.Detail.Count-1 downto 0 do TOBAdresses.Detail[i_ind].Free ;
if GetParamSoc('SO_GCPIECEADRESSE') then
   BEGIN
   TOB.Create('PIECEADRESSE',TOBAdresses,-1) ; {Livraison}
   TOB.Create('PIECEADRESSE',TOBAdresses,-1) ; {Facturation}
   END else
   BEGIN
   TOB.Create('ADRESSES',TOBAdresses,-1) ; {Livraison}
   TOB.Create('ADRESSES',TOBAdresses,-1) ; {Facturation}
   END ;
LoadLesAdresses(TOBPiece,TOBAdresses) ;
TOBAdresses.SetAllModifie(True);
NumL:=TOBPiece.GetValue('GP_NUMADRESSELIVR');
NumF:=TOBPiece.GetValue('GP_NUMADRESSEFACT');
if GetParamSoc('SO_GCPIECEADRESSE') then
   BEGIN
   TOBGenere.PutValue('GP_NUMADRESSELIVR',+1) ;
   if NumF=NumL then TOBGenere.PutValue('GP_NUMADRESSEFACT',+1)
                else TOBGenere.PutValue('GP_NUMADRESSEFACT',+2);
   END else
   BEGIN
   TOBGenere.PutValue('GP_NUMADRESSELIVR',-1) ;
   if NumF=NumL then TOBGenere.PutValue('GP_NUMADRESSEFACT',-1)
                else TOBGenere.PutValue('GP_NUMADRESSEFACT',-2);
   END ;
END ;

{================================ Numero Série ===============================}
Procedure G_LoadSerie ( TOBL : TOB ) ;
Var TOBSerLig: TOB ;
    Q    : TQuery ;
    CleDoc : R_CleDoc ;
BEGIN
{$IFDEF CCS3}
Exit ;
{$ELSE}
if ((TOBL.GetValue('GL_TYPELIGNE')<>'ART') or (TOBL.GetValue('GL_INDICESERIE')<=0))  then Exit ;
TOBSerLig:=TOB.Create('',Nil,-1) ; CleDoc:=TOB2CleDoc(TOBL) ;
Q:=OpenSQL('SELECT * FROM LIGNESERIE WHERE '+WherePiece(CleDoc,ttdSerie,False)+' AND GLS_NUMLIGNE='+IntToStr(TOBL.GetValue('GL_NUMLIGNE')) + ' AND GLS_RANG=0',True,-1, '', True) ;
TOBSerLig.LoadDetailDB('LIGNESERIE','','',Q,True,False) ;
Ferme(Q) ;
if TOBSerLig.Detail.Count>0 then
   begin
   TOBSerLig.ChangeParent(TOBSerie,-1);
   TOBL.PutValue('GL_INDICESERIE',IndiceSerie) ; Inc(IndiceSerie) ;
   end else
   begin
   TOBSerLig.Free;
   TOBL.PutValue('GL_INDICESERIE',0) ;
   end;
{$ENDIF}
END ;

{==================================== LIGNES ==================================}
procedure G_LigneCommentaire(TobOrigine, TobMere: TOB);
Var NewL,TOBL : TOB ;
    RefP : String ;
BEGIN
if TobOrigine.Detail.Count > 0 then
  TOBL := TobOrigine.Detail[0]
else
  TOBL := Nil;
NewL := NewTOBLigne(TobMere, 0);
if TOBL<>nil then
begin
  NewL.Dupliquer(TOBL,False,True);
  NewTOBLigneFille(NewL);
end;
RefP:=RechDom('GCNATUREPIECEG',TobOrigine.GetValue('GP_NATUREPIECEG'),False)
   +' N° '+IntToStr(TobOrigine.GetValue('GP_NUMERO'))
   +' du '+DateToStr(TobOrigine.GetValue('GP_DATEPIECE'))
   +'  '+TobOrigine.GetValue('GP_REFINTERNE') ;
RefP:=Copy(RefP,1,70) ;
NewL.PutValue('GL_NUMLIGNE', 0);
NewL.PutValue('GL_NUMORDRE', 0);
NewL.PutValue('GL_LIBELLE',RefP)    ; NewL.PutValue('GL_TYPELIGNE','COM') ;
NewL.PutValue('GL_TYPEDIM','NOR')   ; NewL.PutValue('GL_CODEARTICLE','') ;
NewL.PutValue('GL_ARTICLE','')      ; NewL.PutValue('GL_QTEFACT',0) ;
NewL.PutValue('GL_QTESTOCK',0)      ; NewL.PutValue('GL_PUHTDEV',0) ;
NewL.PutValue('GL_QTERESTE',0)      ; { NEWPIECE }
//--- GUINIER ---
NewL.PutValue('GL_MTRESTE',0)       ; {NEWPIECE}

NewL.PutValue('GL_PUTTCDEV',0)      ; NewL.PutValue('GL_TYPEARTICLE','') ;
NewL.PutValue('GL_PUHT',0)          ; NewL.PutValue('GL_PUHTNET',0) ;
NewL.PutValue('GL_PUTTC',0)         ; NewL.PutValue('GL_PUTTCNET',0) ;
NewL.PutValue('GL_PUHTBASE',0)      ; NewL.PutValue('GL_FAMILLETAXE1','') ;
NewL.PutValue('GL_TYPENOMENC','')   ; NewL.PutValue('GL_QUALIFMVT','') ;
NewL.PutValue('GL_REFARTSAISIE','') ; NewL.PutValue('GL_REFARTBARRE','') ;
NewL.PutValue('GL_REFCATALOGUE','') ; NewL.PutValue('GL_TYPEREF','') ;
NewL.PutValue('GL_REFARTTIERS','')  ;
{Modif JLD 20/06/2002}
NewL.PutValue('GL_ESCOMPTE',TobOrigine.GetValue('GP_ESCOMPTE')) ;
NewL.PutValue('GL_REMISEPIED',TobOrigine.GetValue('GP_REMISEPIED')) ;
{Fin modif}

//JS 17/06/03
NewL.PutValue('GL_INDICESERIE',0) ; NewL.PutValue('GL_INDICELOT',0) ;
NewL.PutValue('GL_REMISELIGNE',0) ;
// Modif BTP
NewL.PutValue('GL_TYPEARTICLE','EPO');
NewL.PutValue('GL_PUHTNETDEV',0)    ; NewL.PutValue('GL_PUTTCNETDEV',0) ;
NewL.PutValue('GL_BLOCNOTE','')    ; NewL.PutValue('GL_QUALIFQTEVTE','') ;
// ---
ZeroLigne(NewL) ;
END ;

Procedure G_ChargeLesLignes ( TOBPiece : TOB ) ;
Var //Q : TQuery ;
    CD : R_CleDoc ;
BEGIN
CD:=TOB2CleDoc(TOBPiece) ;
(* Avant :
Q:=OpenSQL('SELECT * FROM LIGNE WHERE '+WherePiece(CD,ttdLigne,False)+' ORDER BY GL_NUMLIGNE',True) ;
if Not Q.EOF then TOBPiece.LoadDetailDB('LIGNE','','',Q,True,True) ;
Ferme(Q) ;
 Apres : *)
LoadLignes(CD, TobPiece);
if TOBPiece.Detail.Count>0 then
   BEGIN
   {Ajout ligne de commentaire}
//   if Commentaires then G_LigneCommentaire(TOBPiece) ;
   {Déroulement des lignes}
   PieceAjouteSousDetail(TOBPiece);  // GM 08/02/02
{
   AddLesSupLigne(TOBPiece.Detail[0],True) ;
   for i:=0 to TOBPiece.Detail.Count-1 do
       BEGIN
       TOBL:=TOBPiece.Detail[i] ; InitLesSupLigne(TOBL) ;
       TOB.Create('',TOBL,-1) ;  // manque la 2eme fille !!! GM
       END ;
}
   END ;
END ;

// Modif BTP
procedure G_ChargeLesRG(TOBL:TOB) ;
Var Q : TQuery ;
    CD : R_CleDoc ;
    i  : integer ;
    TOBDocP,TOBInsere : TOB ;
BEGIN
if TOBL.getValue('GL_TYPELIGNE')<>'RG' then exit;
TOBDocP := TOB.create ('LATOB',nil,-1);
CD:=TOB2CleDoc(TOBL) ;
Q:=OpenSQL('SELECT * FROM PIECERG WHERE '+WherePiece(CD,ttdRetenuG,False)+' AND PRG_NUMLIGNE='+inttostr(TOBL.GetValue('GL_NUMLIGNE')),True,-1, '', True) ;
if Not Q.EOF then TOBDocP.LoadDetailDB('PIECERG','','',Q,True,False) ;
Ferme(Q) ;
for i:= 0 to TOBDocP.detail.count -1 do
    begin
    TOBInsere:=TOB.create ('PIECERG',TOBPieceRG,-1);
    TOBInsere.Dupliquer (TOBDocP.detail[i],true,true);
    TOBInsere.AddChampSupValeur ('INDICERG',IndiceRG);
    TOBL.PutValue ('INDICERG',IndiceRG);
    inc(IndiceRg);
    end;
TOBDocP.free; //TobDocP := nil;
end;
// --

{==================================== PIECE ===================================}
Procedure G_FinirPieceGenere ;
Var i : integer ;
    TOBL : TOB ;
    RefPiece,RefA : String ;
    RemDiff : boolean ;
    RemCur,TotRemDev,TotMontantDev,RemMoy : double ;
BEGIN
{Traitement remise pied}
RemDiff:=False ; RemCur:=0 ;
TotRemDev:=0 ; TotMontantDev:=0 ;
for i:=0 to TOBGenere.Detail.Count-1 do
    BEGIN
    TOBL:=TOBGenere.Detail[i] ;
    if i=0 then RemCur:=TOBL.GetValue('GL_REMISEPIED') else
     if TOBL.GetValue('GL_REMISEPIED')<>RemCur then RemDiff:=True ;
    if TOBL.GetValue('GL_REMISABLEPIED')='X' then
       BEGIN
       TotRemDev:=TotRemDev+TOBL.GetValue('GL_TOTREMPIEDDEV') ;
       TotMontantDev:=TotMontantDev+TOBL.GetValue('GL_MONTANTHTDEV') ;
       END ;
    END ;
if ((RemDiff) and (TotRemDev<>0) and (TotMontantDev<>0)) then
   BEGIN
   RemMoy:=Arrondi(100.0*TotRemDev/TotMontantDev,8) ;
   TOBGenere.PutValue('GP_REMISEPIED',RemMoy) ;
   for i:=0 to TOBGenere.Detail.Count-1 do
       BEGIN
       TOBL:=TOBGenere.Detail[i] ;
       TOBL.PutValue('GL_REMISEPIED',RemMoy) ;
       END ;
   END ;
RefA:=EncodeRefCPGescom(TOBGenere) ; UG_MajAnalPiece(TOBAnaP,TOBAnaS,RefA) ;
for i:=0 to TOBGenere.Detail.Count-1 do
    BEGIN
    TOBL:=TOBGenere.Detail[i] ;
    {Traçabilité}
    RefPiece:=EncodeRefPiece(TOBL) ; TOBL.PutValue('GL_PIECEPRECEDENTE',RefPiece) ;
    if TOBL.GetValue('GL_PIECEORIGINE')='' then TOBL.PutValue('GL_PIECEORIGINE',RefPiece) ;
    {Pièce vers lignes}
    UG_PieceVersLigne(TOBGenere,TOBL) ;
    {Qtes}
    TOBL.PutValue('GL_QTERELIQUAT',0)   ;
    TOBL.PutValue('GL_MTRELIQUAT' ,0)   ;
    TOBL.PutValue('GL_SOLDERELIQUAT','-') ;
    {Analytique}
    UG_MajAnalLigne(TOBL,RefA) ;
    TobL.PutValue ('GL_NUMORDRE', 0);
{$IFDEF BTP}
    // c'est un regroupement donc on fait suivre l'info
    TOBL.putValue('NUMCONSOPREC',TOBL.getValue('BLP_NUMMOUV'));
    TOBL.PutValue('BLP_NUMMOUV',0);
{$ENDIF}
    ZeroLigneMontant (TOBL);
    END ;
{Numéros lignes}
NumeroteLignesGC(Nil, TOBGenere, True) ;
END ;

Procedure G_AlimNewPiece ( TOBPiece : TOB ; DupPiece : Boolean ) ;
Var NewSouchePiece : String;
//  BBI : correction fiche 10342
    i_ind : integer;
BEGIN
TOBGenere.ClearDetail ; TOBGenere.InitValeurs ;
if DupPiece then
   begin
   TOBGenere.Dupliquer(TOBPiece,False,True) ;
   AddLesSupEntete (TOBGenere);
   end
   else
   begin
   for i_ind:=1 to 9 do TOBGenere.PutValue('GP_LIBRETIERS'+inttostr(i_ind),TOBPiece.GetValue('GP_LIBRETIERS'+inttostr(i_ind)));
   TOBGenere.PutValue('GP_LIBRETIERSA',TOBPiece.GetValue('GP_LIBRETIERSA'));
   for i_ind:=1 to 3 do TOBGenere.PutValue('GP_LIBREPIECE'+inttostr(i_ind),TOBPiece.GetValue('GP_LIBREPIECE'+inttostr(i_ind)));
   TOBGenere.PutValue('GP_REPRESENTANT',TOBPiece.GetValue('GP_REPRESENTANT'));
   // Correction FQ 12110
   TOBGenere.putValue('GP_AFFAIRE',TOBPiece.getValue('GP_AFFAIRE'));
   TOBGenere.putValue('GP_AFFAIRE1',TOBPiece.getValue('GP_AFFAIRE1'));
   TOBGenere.putValue('GP_AFFAIRE2',TOBPiece.getValue('GP_AFFAIRE2'));
   TOBGenere.putValue('GP_AFFAIRE3',TOBPiece.getValue('GP_AFFAIRE3'));
   TOBGenere.putValue('GP_AVENANT',TOBPiece.getValue('GP_AVENANT'));
   // -----
   end;
//  BBI : fin correction fiche 10342
TOBGenere.PutValue('GP_NATUREPIECEG',NewNature) ;
TOBGenere.PutValue('GP_DATEPIECE',LaNewDate) ;
TOBGenere.PutValue('GP_INDICEG',0) ;
NewSouchePiece:=GetSoucheG( NewNature, TOBGenere.GetValue('GP_ETABLISSEMENT'),TOBGenere.GetValue('GP_DOMAINE'));
TOBGenere.PutValue('GP_SOUCHE',NewSouchePiece);
NewNum:=GetNumSoucheG(NewSouchePiece,LaNewDate) ; TOBGenere.PutValue('GP_NUMERO',NewNum) ;
if GetInfoParPiece(NewNature,'GPP_ACTIONFINI')='ENR' then TOBGenere.PutValue('GP_VIVANTE','-') else TOBGenere.PutValue('GP_VIVANTE','X') ;
TOBGenere.PutValue('GP_DATECREATION',Date) ;
TOBGenere.PutValue('GP_DATEMODIF',NowH) ;
TOBGenere.PutValue('GP_ETATVISA','NON') ;
TOBGenere.PutValue('GP_EDITEE','-') ;
TOBGenere.PutValue('GP_CREEPAR','GEN') ;
TOBGenere.PutValue('GP_VENTEACHAT',VenteAchat) ;
TOBGenere.PutValue('GP_DATETAUXDEV',TOBPiece.GetValue('GP_DATETAUXDEV')) ;
TOBGenere.PutValue('GP_REFINTERNE',TOBPiece.GetValue('GP_REFINTERNE')) ;
TOBGenere.PutValue('GP_REFEXTERNE',TOBPiece.GetValue('GP_REFEXTERNE')) ;
TOBGenere.PutValue('GP_RIB',TOBPiece.GetValue('GP_RIB')) ;
{Autres informations}
TOBNomenclature.ClearDetail ; IndiceNomen:=1 ;
TOBSerie.ClearDetail ; IndiceSerie:=1 ;
{Analytique}
TOBAnaP.ClearDetail ; TOBAnaS.ClearDetail ; {UG_LoadAnaPiece(TOBPiece,TOBAnaP,TOBAnaS)} ;
// Modif BTP
{Ouvrages}
TOBOUvrage.clearDetail;
TOBOUvragesP.clearDetail;
{retenues de garantie}
TOBPieceRG.clearDetail;
{tva sur retenues de garantie}
TOBBasesRG.clearDetail;
// ------
END ;

procedure G_ValideLesSeries ;
Var i,IndSerie : integer ;
    TOBSerLig,TOBL : TOB ;
BEGIN
{$IFDEF CCS3}
Exit ;
{$ELSE}
if TOBSerie.Detail.Count<=0 then Exit ;
TOBSerie.Detail[0].AddChampSup('UTILISE',True) ;
for i:=0 to TOBGenere.Detail.Count-1 do
    BEGIN
    IndSerie:=GetChampLigne(TOBGenere,'GL_INDICESERIE',i+1) ;
    if IndSerie>0 then
       BEGIN
       TOBL:=TOBGenere.Detail[i] ; TOBSerLig:=TOBSerie.Detail[IndSerie-1] ;
       TOBSerLig.PutValue('UTILISE','X') ;
       UpdateLesSeries(TOBL,TOBSerLig) ;
       MAJDispoSerie(TOBL,TOBSerLig,False);
       END ;
    END ;
for i:=TOBSerie.Detail.Count-1 downto 0 do
    BEGIN
    TOBSerLig:=TOBSerie.Detail[i] ;
    if TOBSerLig.GetValue('UTILISE')<>'X' then TOBSerLig.Free ;
    END ;
if Not TOBSerie.InsertDB(Nil) then V_PGI.IoError:=oeUnknown ;
{$ENDIF}
END ;

{============================ GENERATION, COMPTABILITE ========================}
Procedure G_GenereCompta ;
Var OldEcr,OldStk : RMVT ;
BEGIN
FillChar(OldEcr,Sizeof(OldEcr),#0) ; FillChar(OldStk,Sizeof(OldStk),#0) ;
if Not PassationComptable(TOBGenere,TOBOuvrage ,TOBOuvragesP,TOBBases,TOBBasesL,TOBEches,nil,nil,TOBTiers,TOBArticles,TOBCpta,TOBAcomptes,TOBPorcs,TOBPIECERG,TOBBasesRG,TOBanaP,TOBanaS,nil,nil,DEV,OldEcr,OldStk,True)
   then V_PGI.IoError:=oeLettrage ;
LibereParamTimbres;   
END ;

procedure G_InitToutModif ;
Var NowFutur : TDateTime ;
BEGIN
NowFutur:=NowH ;
TOBGenere.SetAllModifie(True) ; TOBGenere.SetDateModif(NowFutur) ;
TOBBases.SetAllModifie(True)  ;
TOBBasesL.SetAllModifie(True)  ;
TOBEches.SetAllModifie(True)  ;
TOBAcomptes.SetAllModifie(True)  ;
TOBPorcs.SetAllModifie(True)  ;
TOBSerie.SetAllModifie(True)  ;
TOBTiers.SetAllModifie(True)  ;
TOBAnaP.SetAllModifie(True)   ;
// Modif BTP
TOBOuvrage.setallmodifie (true);
TOBOuvragesP.setallmodifie (true);
TOBLIENOLE.setallmodifie (true);
TOBPIECERG.setallmodifie (true);
TOBBasesRG.setallmodifie (true);
END ;

Procedure G_AjouteEvent ;
Var St : String ;
BEGIN
St:='Pièce N° '+IntToStr(TOBGenere.GetValue('GP_NUMERO'))
   +', Tiers '+TOBGenere.GetValue('GP_TIERS')
   +', Total HT de '+StrfPoint(TOBGenere.GetValue('GP_TOTALHTDEV'))+' '+RechDom('TTDEVISETOUTES',TOBGenere.GetValue('GP_DEVISE'),False) ;
PiecesG.Add(St) ;
END ;

Procedure G_TraiteEuro ;
Var leModeOppose : boolean ;
BEGIN
if Not ForceEuroGC then Exit ;
LeModeOppose:=(TOBGenere.GetValue('GP_SAISIECONTRE')='X') ; if Not LeModeOppose then Exit ;
if TOBGenere.GetValue('GP_DEVISE')<>V_PGI.DevisePivot then Exit ;
PutValueDetail(TOBGenere,'GP_SAISIECONTRE','-') ;
ForcePieceEuro(TOBGenere,TOBBases,TOBEches,TOBPorcs,TOBNomenclature,TOBAcomptes,TOBOuvrage,TOBPIeceRG,TOBBasesRG)  ;
END ;

Procedure G_GenereLaPiece ( TOBPiece : TOB ) ;
Var St,Domaine,Nature,Lib : string ;
    NeedVisaEnd : boolean ;
    MontantVisaEnd : Double ;
    GestionConso : TGestionPhase;
BEGIN
if TOBGenere.Detail.Count<=0 then Exit ;
{$IFDEF BTP}
gestionConso := TGestionPhase.create;
{$ENDIF}
{Calculs et finitions}
G_FinirPieceGenere ;
UG_MajLesComms(TOBGenere,TOBArticles,TOBComms) ;
ValideLaPeriode(TOBGenere) ;
{$IFDEF BTP}
//AjouteLignesVirtuellesOuv (TOBGenere,TOBOuvrage,TOBArticles,TOBCpta,TOBTiers,TOBCatalogu,nil,nil,DEV);
{$ENDIF}
TOBBases.clearDetail; TOBBasesL.clearDetail; ZeroFacture (TOBGenere);
ZeroMontantPorts (TOBPorcs);
PutValueDetail(TOBGenere,'GP_RECALCULER','X');
CalculFacture(nil,TOBGenere,nil,nil,TOBOuvrage,TOBOuvragesP,TOBBases,TOBBasesL,TOBTiers,TOBArticles,TOBPorcs,TOBPieceRG,TOBBasesRG,nil,DEV) ;
{$IFDEF BTP}
//DetruitLignesVirtuellesOuv (TOBGenere,DEV);
{$ENDIF}
CalculeSousTotauxPiece(TOBGenere) ;
{Visa}
Domaine:=TOBGenere.GetValue('GP_DOMAINE') ; Nature:=TOBGenere.GetValue('GP_NATUREPIECEG') ; Lib:='' ;
if Domaine<>'' then Lib:=GetInfoParPieceDomaine(Nature,Domaine,'GDP_LIBELLE') ;
if Lib<>'' then
   BEGIN
   NeedVisaEnd:=(GetInfoParPieceDomaine(Nature,Domaine,'GDP_VISA')='X') ;
   MontantVisaEnd:=GetInfoParPieceDomaine(Nature,Domaine,'GDP_MONTANTVISA') ;
   END else
   BEGIN
   NeedVisaEnd:=NeedVisa ; MontantVisaEnd:=MontantVisa ;
   END ;
if NeedVisaEnd then
   BEGIN
   if Abs(TOBGenere.GetValue('GP_TOTALHT'))>=MontantVisaEnd then TOBGenere.PutValue('GP_ETATVISA','ATT') ;
   END ;
{Echéances}
TOBEches.ClearDetail ;GereEcheancesGC(TOBGenere,TOBTiers,TOBEches,TOBAcomptes,TOBPieceRG,nil,TOBPorcs,taCreat,DEV,False) ;
G_InitToutModif ;
G_TraiteEuro ;
{Enregistrement physique}
if Not SetNumeroDefinitif(TOBGenere,TOBBases,TOBBasesL,TOBEches,TOBNomenclature,TOBAcomptes,TOBPieceRG,TOBBasesRG, nil) then V_PGI.IoError:=oeSaisie ;
// Correction OPTIMISATION
if V_PGI.IoError = oeOk then ValideLesArticlesFromOuv(TOBarticles,TOBOuvrage);
// --
if V_PGI.IoError=oeOk then ValideLesLignes(TOBGenere,TOBArticles,TobCatalogu,TOBNomenclature,TOBOuvrage,TOBPieceRG,TOBBasesRG,False,True) ;
if V_PGI.IoError=oeOk then ValideLesLignesCompl(TOBGenere, nil);
if V_PGI.IoError=oeOk then ValideLesAdresses(TOBGenere,TOBPiece,TOBAdresses) ;
if V_PGI.IoError=oeOk then ValideLesArticles(TOBGenere,TOBArticles) ;
if V_PGI.IoError=oeOk then ValideleTiers(TOBGenere,TOBTiers) ;
{$IFDEF BTP}
//PutValueDetail (TOBgenere,'GP_RECALCULER','X');
//AjouteLignesVirtuellesOuv (TOBGenere,TOBOuvrage,TOBArticles,TOBCpta,TOBTiers,TOBCatalogu,nil,nil,DEV);
//CalculFacture(TOBGenere,TOBOuvrage,TOBOuvragesP,TOBBases,TOBbasesL,TOBTiers,TOBArticles,TOBPorcs,TOBPieceRG,TOBBasesRG,DEV);
{$ENDIF}
if V_PGI.IoError = oeOk then ValideLesBases(TOBgenere,TobBases,TOBBasesL);
if V_PGI.IoError=oeOk then UG_ValideAnals(TOBGenere,TOBAnaP,TOBAnaS) ;
if V_PGI.IoError=oeOk then G_GenereCompta ;
if V_PGI.IoError=oeOk then G_ValideLesSeries ;
if V_PGI.IoError=oeOk then ValideLesAcomptes(TOBGenere,TOBAcomptes) ;
if V_PGI.IoError=oeOk then ValideLesPorcs(TOBGenere,TOBPorcs) ;
{$IFDEF BTP}
//DetruitLignesVirtuellesOuv (TOBGenere,DEV);
{$ENDIF}
if V_PGI.IoError=oeOk then TOBGenere.InsertDBByNivel(False) ;
if V_PGI.IoError=oeOk then TOBBases.InsertDB(Nil) ;
if V_PGI.IoError=oeOk then TOBBasesL.InsertDB(Nil) ;
if V_PGI.IoError=oeOk then TOBEches.InsertDB(Nil) ;
if V_PGI.IoError=oeOk then TOBAnaP.InsertDB(Nil) ;
if V_PGI.IoError=oeOk then ValideLesNomen(TOBNomenclature) ;
{$IFDEF BTP}
if V_PGI.IoError=oeOk then ValideLesOuv(TOBOuvrage,TOBgenere) ;
(* CONSO *)
if (V_PGI.IOError = OeOk) Then GestionConso.GenerelesPhases(TOBGenere,nil,false,false,True,taCreat);
(* ---- *)
{$ENDIF}
// Modif BTP
if V_PGI.IoError=oeOk then ValideLesRetenues (TOBGenere,TOBPIECERG);
if V_PGI.IoError=oeOk then ReajusteCautionDocOrigine (TOBGenere,TOBPIECERG,true);
if V_PGI.IoError=oeOk then ValideLesBasesRG(TOBgenere,TOBBasesRG);
// ----
{$IFDEF AFFAIRE}
if V_PGI.IoError=oeOk then if ctxAffaire in V_PGI.PGIContexte then MajGestionAffaire(TOBGenere,TOBPiece,nil,nil,TOBAComptes,'VAL',DEV) ;
if V_PGI.IoError=oeOk then ValideActivite(TOBGenere,Nil,TOBArticles, GereActivite,False,DelActivite,False);
{$ENDIF}
{$IFDEF BTP}
gestionConso.free;
{$ENDIF}
{Compte rendu}
if V_PGI.IoError=oeOk then
   BEGIN
   if PremNum=-1 then PremNum:=TOBGenere.GetValue('GP_NUMERO') ;
   LastNum:=TOBGenere.GetValue('GP_NUMERO') ;
   St:=TOBGenere.GetValue('GP_NATUREPIECEG') + ';' +
       DateToStr (TOBGenere.GetValue('GP_DATEPIECE')) + ';' +
       TOBGenere.GetValue('GP_SOUCHE') + ';' +
       IntToStr (TOBGenere.GetValue('GP_NUMERO')) + ';' +
       IntToStr (TOBGenere.GetValue('GP_INDICEG')) + ';';
   StringToCleDoc(St,ClePiece) ;
   Inc(Nbp) ; G_AjouteEvent ;
   END ;
END ;

Procedure G_InversePiece ( TOBPiece,TobTiers : TOB ) ; { NEWPIECE }
Var CleDoc : R_CleDoc ;
    TOBBases_O,TOBEches_O,TOBN_O,TOBPiece_O,TOBAcomptes_O,TOBPorcs_O, TobDetPiece_O,TOBSerie_O : TOB ;
    TobLigneTarif_O: Tob;
    NowFutur : TDateTime ;
//    PasHisto : boolean ;
    Q        : TQuery ;
    OldEcr,OldStk : RMVT ;
    i        : integer;
    // Modif BTP
    TOBOuvrage_O,TOBLIENOLE_O,TOBPIECERG_O,TOBBasesRG_O : TOB;
{$IFDEF BTP}
    IndiceOuv : integer;
{$ENDIF}
    //
    OldCleDoc: R_CleDoc;
    OldTobL, NewTobL: Tob;
BEGIN
CleDoc:=TOB2CleDoc(TOBPiece) ;
TOBPiece_O:=TOB.Create('PIECE',Nil,-1) ;
// Modif BTP
{$IFDEF BTP}
IndiceOuv := 1;
{$ENDIF}
AddLesSupEntete (TOBPiece_O);
// ---
TOBPiece_O.Dupliquer(TOBPiece,False,True) ;
// Au niveau filles on duplique uniquement les lignes ( sans analytique rajoutée en dessous)
for i :=0 to TOBPiece.Detail.count-1 do
    BEGIN
    TobDetPiece_O:=TOB.Create('LIGNE',TOBPiece_O,-1) ;
    TOBDetPiece_O.Dupliquer(TOBPiece.Detail[i],false,True) ;
    END;
// Numeros de série
TOBSerie_O:=TOB.Create('',Nil,-1) ;
TOBSerie_O.Dupliquer(TOBSerie,True,True) ;
TOBSerie_O.SetAllModifie(False) ;
// Lecture bases
TOBBases_O:=TOB.Create('BASES',Nil,-1) ;
Q:=OpenSQL('SELECT * FROM PIEDBASE WHERE '+WherePiece(CleDoc,ttdPiedBase,False),True,-1, '', True) ;
TOBBases_O.LoadDetailDB('PIEDBASE','','',Q,False) ;
Ferme(Q) ;
// Lecture Echéances
TOBEches_O:=TOB.Create('Les ECHEANCES',Nil,-1) ;
Q:=OpenSQL('SELECT * FROM PIEDECHE WHERE '+WherePiece(CleDoc,ttdEche,False),True,-1, '', True) ;
TOBEches_O.LoadDetailDB('PIEDECHE','','',Q,False) ;
Ferme(Q) ;
// Lecture Acomptes
TOBAcomptes_O:=TOB.Create('',Nil,-1) ;
Q:=OpenSQL('SELECT * FROM ACOMPTES WHERE '+WherePiece(CleDoc,ttdAcompte,False),True,-1, '', True) ;
TOBAcomptes_O.LoadDetailDB('ACOMPTES','','',Q,False) ;
Ferme(Q) ;
// Lecture Ports
TOBPorcs_O:=TOB.Create('',Nil,-1) ;
Q:=OpenSQL('SELECT * FROM PIEDPORT WHERE '+WherePiece(CleDoc,ttdPorc,False),True,-1, '', True) ;
TOBPorcs_O.LoadDetailDB('PIEDPORT','','',Q,False) ;
Ferme(Q) ;
// Lecture nomenclatures
TOBN_O:=TOB.Create('NOMENCLATURES',Nil,-1) ;
LoadLesNomen(TOBPiece,TOBN_O,TOBArticles,CleDoc) ;
// Modif BTP
// Lecture Ouvrage
TOBOuvrage_O:=TOB.Create('OUVRAGES',Nil,-1) ;
{$IFDEF BTP}
LoadLesOuvrages(TOBPiece,TOBOuvrage_O,TOBArticles,CleDoc,IndiceOuv) ;
{$ENDIF}
// Lecture tEXTES DEBUT ET fin
TOBLIENOLE_O:=TOB.Create('OLE',Nil,-1) ;
Q:=OpenSQL('SELECT * FROM LIENSOLE WHERE '+WherePiece(CleDoc,ttdLienOle,False),True,-1, '', True) ;
TOBLIENOLE_O.LoadDetailDB('LIENSOLE','','',Q,False) ;
Ferme(Q) ;
TOBPIECERG_O:=TOB.create('RETENUES',nil,-1);
Q:=OpenSQL('SELECT * FROM PIECERG WHERE '+WherePiece(CleDoc,ttdretenuG,False),True,-1, '', True) ;
TOBPieceRg.selectDB ('',Q) ;
Ferme(Q) ;
TOBBasesRg_O:=TOB.create('BASESRG',nil,-1);
Q:=OpenSQL('SELECT * FROM PIEDBASERG WHERE '+WherePiece(CleDoc,ttdBaseRG,False),True,-1, '', True) ;
TOBBasesRg.selectDB ('',Q) ;
Ferme(Q) ;
TOBLigneTarif_O:=TOB.Create('',Nil,-1) ;
(* SUPPRIME -- SUITE A DROPTABLE DANS PGIMAJVER
Q:=OpenSQL('SELECT * FROM LIGNETARIF WHERE '+WherePiece(CleDoc,ttdLigneTarif,False),True) ;
TOBLigneTarif_O.LoadDetailDB('LIGNETARIF','','',Q,False) ;
*)
Ferme(Q) ;                                                                                   
// ----------
{Mise à jour inverse}
NowFutur:=NowH ; TOBPiece_O.SetDateModif(NowFutur) ; TOBPiece_O.CleWithDateModif:=True ;
if V_PGI.IoError=oeOk then DetruitCompta(TOBPiece_O,NowFutur,OldEcr,OldStk) ;
if V_PGI.IoError=oeOK then
begin
  {Flag pour InverseStockTransfo}
  TobPiece_O.AddChampSupValeur('GENAUTO', 'X');
  InverseStockTransfo(TOBPiece_O, TobGenere, TOBArticles, TOBCatalogu, TOBN_O);
end;
{ Remise à 0 du Reste à livrer dans la pièce d'origine }
{ Remise à jour du Reste à livrer dans la pièce d'origine }
{ TobPiece_O contient l'ancienne pièce, TobGenere contient la nouvelle piece (pas encore totalement mise à jour) }
for i := 0 to TobPiece_O.Detail.Count - 1 do
begin
  OldTobL := TobPiece_O.Detail[i];
  if EstLigneArticle(OldTobL) then
  begin
    OldCleDoc := Tob2CleDoc(OldTobL);
    OldCleDoc.NumLigne := OldTobL.GetValue('GL_NUMLIGNE');
    OldCleDoc.NumOrdre := OldTobL.GetValue('GL_NUMORDRE');
    NewTobL := FindCleDocInTob(OldCleDoc, TobGenere, True);
    if Assigned(NewTobL) then
    begin
    	// Modif LS -- Du fait que cela soit une regroupement manuel la quantité entiere est passé dans la piece regroupe..d'ou la
      // raz du qteresre
			// Ancien      OldTobL.PutValue('GL_QTERESTE', OldTobL.GetValue('GL_QTERESTE') - NewTobl.GetValue('GL_QTESTOCK'));
      OldTobL.PutValue('GL_QTERESTE', 0);
      //--- GUINIER ---
      OldTobL.PutValue('GL_MTRESTE',  0);

    end;
  end;
end;
if not HistoPiece(TobPiece_O) then { NEWPIECE }
begin
  { NEWPIECE }
  if not TOBPiece_O.UpdateDB(False) then
    V_PGI.IoError := oeSaisie;
  if V_PGI.IoError = oeOk then
  begin
    if not DetruitAncien(TOBPiece_O,TOBBases_O,TOBEches_O,TOBN_O,nil,TOBAcomptes_O,TOBPorcs_O,TOBSerie_O,TOBOuvrage_O,TOBLienOle_O,TOBPieceRG_O,TOBBasesRg_O, nil) then
      V_PGI.IoError:=oeSaisie ;
  end;
end
else
begin
  if UpdateAncien(TOBPiece_O,TOBGenere,TOBAcomptes_O,TobTiers,True) then
  begin
    { NEWPIECE }
    if not TOBPiece_O.UpdateDB(False) then V_PGI.IoError := oeSaisie;
    RazLesSeries(nil,TobSerie_O);
    if not TOBSerie_O.UpdateDB(False) then V_PGI.IoError := oeSaisie;
  end
  else
    V_PGI.IoError := oeSaisie;
end;
//if V_PGI.IoError=oeOk then InverseStockTransfo(TOBPiece_O, TobPiece, TOBArticles, TOBCatalogu, TOBN_O); { NEWPIECE }

// if V_PGI.IoError=oeOk then ValideLesArticles(TOBPiece_O,TOBArticles) ; // DBR Suite au Différentiel
{$IFDEF AFFAIRE}
if V_PGI.IoError=oeOk then ValideActivite(Nil,TOBPiece_O, TOBArticles, GereActivite, False,DelActivite,False);
{$ENDIF}
{$IFDEF BTP}
if (V_PGI.IoError=oeOk) then MajQteFacture(TobGenere, TobPiece_O);
{$ENDIF}
{$IFNDEF CCS3}
if V_PGI.IoError=oeOk then InverseStockSerie(TOBPiece_O,TOBSerie_O) ;
{$ENDIF}
{Libérations}
TOBPiece_O.Free ; TOBBases_O.Free ; TOBEches_O.Free ; TOBN_O.Free ; TOBAcomptes_O.Free ; TOBPorcs_O.Free ;
TOBSerie_O.Free ;
// Modif BTP
TOBOuvrage_O.free; TOBPIECERG_O.free; TOBBasesRG_O.free;
// --------
TobLigneTarif_O.Free ; 
END ;

Procedure G_LigneVersPieceGenere ( TOBL : TOB ) ;
BEGIN
TOBGenere.PutValue('GP_ETABLISSEMENT',TOBL.GetValue('GL_ETABLISSEMENT')) ;
TOBGenere.PutValue('GP_TIERS',TOBL.GetValue('GL_TIERS')) ;
TOBGenere.PutValue('GP_TIERSFACTURE',TOBL.GetValue('GL_TIERSFACTURE')) ;
TOBGenere.PutValue('GP_TIERSPAYEUR',TOBL.GetValue('GL_TIERSPAYEUR')) ;
TOBGenere.PutValue('GP_TIERSLIVRE',TOBL.GetValue('GL_TIERSLIVRE')) ;
//TOBGenere.PutValue('GP_REPRESENTANT',TOBL.GetValue('GL_REPRESENTANT')) ; Fiche eQualité 10425
TOBGenere.PutValue('GP_APPORTEUR',TOBL.GetValue('GL_APPORTEUR')) ;
TOBGenere.PutValue('GP_DEPOT',TOBL.GetValue('GL_DEPOT')) ;
TOBGenere.PutValue('GP_SAISIECONTRE',TOBL.GetValue('GL_SAISIECONTRE')) ;
TOBGenere.PutValue('GP_REGIMETAXE',TOBL.GetValue('GL_REGIMETAXE')) ;
TOBGenere.PutValue('GP_FACTUREHT',TOBL.GetValue('GL_FACTUREHT')) ;
{Modif JLD 20/06/2002}
TOBGenere.PutValue('GP_ESCOMPTE',TOBL.Parent.GetValue('GP_ESCOMPTE')) ;
TOBGenere.PutValue('GP_REMISEPIED',TOBL.Parent.GetValue('GP_REMISEPIED')) ;
{Fin Modif JLD}
//  BBI : correction fiche 10342
   TOBGenere.PutValue('GP_TVAENCAISSEMENT',TOBL.Parent.GetValue('GP_TVAENCAISSEMENT')) ;
//  BBI : fin correction fiche 10342
TOBGenere.PutValue('GP_DATELIVRAISON',TOBL.GetValue('GL_DATELIVRAISON')) ;
{Traitement des devises}
DEV.Code:=TOBL.GetValue('GL_DEVISE') ; GetInfosDevise(DEV) ; TOBGenere.PutValue('GP_DEVISE',DEV.Code) ;
DEV.Taux:=TOBL.GetValue('GL_TAUXDEV') ; TOBGenere.PutValue('GP_TAUXDEV',DEV.Taux) ;
TOBGenere.PutValue('GP_COTATION',TOBL.GetValue('GL_COTATION')) ;
{Traitement Tiers}
TOBGenere.PutValue('GP_MODEREGLE',TOBTiers.GetValue('T_MODEREGLE')) ;
END ;

Function G_DiffChamps ( TOBL : TOB ; Champ : String ) : boolean ;
Var V1,V2 : Variant ;
BEGIN
Result:=False ;
if TOBL.FieldExists(Champ) then
   BEGIN
   V1:=TOBL.GetValue(Champ) ; V2:=TOBRef.GetValue(Champ) ;
   if V1<>V2 then Result:=True ;
   END ;
END ;

Function G_yaRupture ( TOBL : TOB ) : Boolean ;
Var i : integer ;
    Champ,Junk : String ;
BEGIN
Result:=True ;
for i:=0 to ChampsRupt.Count-1 do
    BEGIN
    Champ:=ChampsRupt[i] ; Junk:=ReadTokenSt(Champ) ; Champ:=ReadTokenSt(Champ) ;
    if G_DiffChamps(TOBL,Champ) then Exit ;
    END ;
{Ne pas regrouper sur codes affaires différents}
(* GM/PCS supprimer sinon le regroupement sur des affaires différentes ne marche pas
if ctxGCAFF in V_PGI.PGIContexte then  // En contexte affaire multiaffaire / pièce autorisé
    BEGIN
    Champ:='GL_AFFAIRE'  ; if G_DiffChamps(TOBL,Champ) then Exit ;
    Champ:='GL_AFFAIRE1' ; if G_DiffChamps(TOBL,Champ) then Exit ;
    Champ:='GL_AFFAIRE2' ; if G_DiffChamps(TOBL,Champ) then Exit ;
    Champ:='GL_AFFAIRE3' ; if G_DiffChamps(TOBL,Champ) then Exit ;
    Champ:='GL_AVENANT'  ; if G_DiffChamps(TOBL,Champ) then Exit ;
    END;
{Tout est OK}
*)
Result:=False ;
END ;

Procedure G_ChargeCompletePiece ( TOBPiece : TOB ) ;
BEGIN
G_ChargeLesLignes(TOBPiece) ;
UG_AjouteLesArticles(TOBPiece,TOBArticles,TOBCpta,TOBTiers,TOBAnaP,TOBAnaS,TOBCatalogu,True) ;
UG_AjouteLesRepres(TOBPiece,TOBComms) ;
END ;

Procedure T_GenVal.G_GenereParTiers ;
Var i,k,LastI,iacc : integer ;
    TOBL,TOBPiece,NewL,TOBTP : TOB ;
    Premier,NewPiece,DupPiece : boolean ;
    RemCur        : double ;
    AnaLigne      : boolean;
    Ok_ReliquatMt : Boolean;
    stDomaine     : string; // DBR Fiche 10665
BEGIN
stDomaine := ''; // DBR Fiche 10665
TOBArticles.ClearDetail ; RemCur:=0 ;
Premier:=True ; NewPiece:=False ; DupPiece:=False ; LastI:=0 ;
// TobTiers  alimentée une seule fois avant G_ChargeCompletePiece PA le 04/09/2000
if RemplirTOBTiers(TOBTiers,CodeTiers,NewNature,False)<>trtOk then V_PGI.IoError:=oeUnknown ;
AnaLigne := UG_IsAnalMultiple (TOBSource);
if not AnaLigne then UG_LoadAnaPiece(TOBSource.detail[0],TOBAnaP,TOBAnaS) ;
for i:=0 to TOBSource.Detail.Count-1 do
    BEGIN
    TOBPiece:=TOBSource.Detail[i] ;
    G_ChargeCompletePiece(TOBPiece) ;
    if Eclate then BEGIN NewPiece:=True ; DupPiece:=True ; END ;
    if DeGroupeRemise then
       BEGIN
       if i=0 then RemCur:=TOBPiece.GetValue('GP_REMISEPIED') else
        if TOBPiece.GetValue('GP_REMISEPIED')<>RemCur then NewPiece:=True ;
       END ;
    stDomaine := TobPiece.GetValue ('GP_DOMAINE'); // DBR Fiche 10665
    for k:=0 to TOBPiece.Detail.Count-1 do
        BEGIN
        TOBL:=TOBPiece.Detail[k] ;
        Ok_ReliquatMt := CtrlOkReliquat(TOBL, 'GL');
        if ((Premier) or (NewPiece) or (G_yaRupture(TOBL))) then
           BEGIN
           if Not Premier then
              BEGIN
              TOBAcomptes.ClearDetail ;
              TOBPorcs.ClearDetail ;
              for iacc:=LastI to i-1 do
                  BEGIN
                  UG_ChargeLesAcomptes(TOBSource.Detail[iacc],TOBAcomptes) ;
                  UG_ChargeLesPorcs(TOBSource.Detail[iacc],TOBPorcs) ;
                  {Inversion et mise à jour pièces d'origine}
                  G_InversePiece(TOBSource.Detail[iacc],TobTiers) ;
                  END ;
              if TobGenere.GetValue ('GP_DOMAINE') = '&&&' then TobGenere.PutValue ('GP_DOMAINE', ''); // DBR Fiche 10665
              if i>0 then G_GenereLaPiece(TOBSource.Detail[i-1])
                     else G_GenereLaPiece(TOBSource.Detail[0]) ;
              LastI:=i ;
              END ;
           G_AlimNewPiece(TOBPiece,DupPiece) ;
           G_ChargeLesAdresses(TOBPiece);
           G_LigneVersPieceGenere(TOBL) ;
           END ;
        if (k = 0) and Commentaires then
          G_LigneCommentaire(TobPiece, TobGenere);
        TOBRef.Dupliquer(TOBL,False,True) ; Premier:=False ;
        NewL:=NewTOBLigne(TOBGenere,0) ;
        NewL.Dupliquer(TOBL,True,True) ;
        // Correction FQ 12110 -------------------
        if NewL.getValue('GL_AFFAIRE')<>TOBGenere.GetValue('GP_AFFAIRE') then
        begin
          TOBGenere.putValue('GP_AFFAIRE','');
          TOBGenere.putValue('GP_AFFAIRE1','');
          TOBGenere.putValue('GP_AFFAIRE2','');
          TOBGenere.putValue('GP_AFFAIRE3','');
          TOBGenere.putValue('GP_AVENANT','');
        end;
        // ---------------------------------------
        if Tobl.GetValue('GL_QTERESTE') <> 0 then
        begin
          NewL.PutValue('GL_QTESTOCK', Tobl.GetValue('GL_QTERESTE'));
          NewL.PutValue('GL_QTEFACT',  Tobl.GetValue('GL_QTERESTE'));
          NewL.PutValue('GL_QTERESTE', Tobl.GetValue('GL_QTERESTE'));
        end
        else
        begin
          NewL.PutValue('GL_QTESTOCK', 0);
          NewL.PutValue('GL_QTEFACT',  0);
          NewL.PutValue('GL_QTERESTE', 0);
        end;

        // --- GUINIER ---
        if Ok_ReliquatMt then
        begin
          if Tobl.GetValue('GL_MTRESTE') <> 0 then
          begin
            NewL.PutValue('GL_MTRESTE', Tobl.GetValue('GL_MTRESTE'));
          end
          else
          begin
            NewL.PutValue('GL_MTRESTE', 0);
          end;
        end;

// DBR Fiche 10665 - Debut
        NewL.PutValue ('GL_DOMAINE', stDomaine);
        if (TobGenere.GetValue ('GP_DOMAINE') = '') and (stDomaine <> '') then
        begin
          TobGenere.PutValue ('GP_DOMAINE', stDomaine);
        end
        else
          if TobGenere.GetValue ('GP_DOMAINE') <> stDomaine then
          begin
            TobGenere.PutValue ('GP_DOMAINE', '&&&');
          end;
// DBR Fiche 10665 - Fin
          G_LoadNomenLigne(NewL) ;
          G_LoadSerie (NewL) ;
          // Modif BTP
{$IFDEF BTP}
          UG_LoadOuvrageLigne(NewL,TOBOuvrage,TOBArticles,IndiceNomen,True) ;
{$ENDIF}
          G_ChargeLesRG(newL) ;
          // -------
          if AnaLigne then UG_LoadAnaLigneOuPas(TOBPiece,NewL) ;
          NewPiece:=False ;
        END ;
{$IFNDEF CCS3}

    {Exception gestion de documents TIERSPIECE}
    TOBTP:=TOBTiers.FindFirst(['GTP_NATUREPIECEG'],[NewNature],False) ;
    if TOBTP<>Nil then if TOBTP.GetValue('GTP_REGROUPE')<>'X' then BEGIN NewPiece:=True ; DupPiece:=True ; END ;
{$ENDIF}
    END ;
{Dernière rupture}
TOBAcomptes.ClearDetail ;
TOBPorcs.ClearDetail ;
for iacc:=LastI to TOBSource.Detail.Count-1 do
    BEGIN
    UG_ChargeLesAcomptes(TOBSource.Detail[iacc],TOBAcomptes) ;
    UG_ChargeLesPorcs(TOBSource.Detail[iacc],TOBPorcs) ;
    {Inversion et mise à jour pièces d'origine}
    G_InversePiece(TOBSource.Detail[iacc],TobTiers) ;
    END ;
if TobGenere.GetValue ('GP_DOMAINE') = '&&&' then TobGenere.PutValue ('GP_DOMAINE', ''); // DBR Fiche 10665
UG_RegroupeLesAcomptes(TOBAcomptes); // JT eQualité 11013
G_GenereLaPiece(TOBSource.Detail[TOBSource.Detail.Count-1]);
END ;

Procedure G_MajEvent ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT * FROM JNALEVENT WHERE GEV_NUMEVENT='+IntToStr(EvDep)+' AND GEV_TYPEEVENT="GEN"',False) ;
if Not Q.EOF then
   BEGIN
   Q.Edit ;
   Q.FindField('GEV_ETATEVENT').AsString:='OK' ;
   TMemoField(Q.FindField('GEV_BLOCNOTE')).Assign(PiecesG) ;
   Q.Post ;
   END ;
Ferme(Q) ;
END ;

Function RegroupeLesPiecesMan ( TOBSource : TOB ; NewNat : String ; Eclate,DeGroupeRemise,AvecComment : boolean ; NewDate : TDateTime = 0 ) : R_CleDoc;

	procedure  SetMorte(Tobpiece : TOB);
  var Cledoc : R_cledoc;
  		Sql : string;
  begin
    Cledoc := TOB2CleDoc(TOBPiece);
		Sql := 'UPDATE LIGNE SET GL_VIVANTE="-" WHERE '+WherePiece (Cledoc,ttdLigne,false);
    ExecuteSQL(Sql);
		Sql := 'UPDATE PIECE SET GP_VIVANTE="-" WHERE '+WherePiece (Cledoc,ttdPiece,false);
    ExecuteSQL(Sql);
  end;

	procedure  SetPieceprecedenteMorte(TOBPieces : TOB);
  var Indice : integer;
  begin
		For Indice := 0 to TOBPieces.detail.count -1 do
    begin
      SetMorte(TOBPieces.detail[Indice]);
    end;

  end;

Var CodeTiers,StRendu : String ;
    io   : TIoErr ;
    GenVal : T_GenVal ;
BEGIN
MemoriseChampsSupLigneETL (NewNat,true);
MemoriseChampsSupLigneOUV (NewNature);
MemoriseChampsSupPIECETRAIT;

{Initialisations}
FillChar(ClePiece,Sizeof(ClePiece),#0) ;
Result:=ClePiece;
if TOBSource.Detail.Count=0 then exit ;
NewNature:=NewNat ; Commentaires:=AvecComment ;
if NewDate>0 then LaNewDate:=NewDate else LaNewDate:=V_PGI.DateEntree ;
G_InitGenere ;
{$IFDEF BTP}
UG_AjouteLesChampsBtp (TobSource);
{$ENDIF}
{Création des TOB}
G_CreerLesTOB ;
G_ChargeChampsRupt ;
{Lecture des pièces}
CodeTiers:=TOBSource.Detail[0].GetValue('GP_TIERS') ;
GenVal:=T_GenVal.Create ;
GenVal.CodeTiers:=CodeTiers ; GenVal.TOBSource:=TOBSource ;
GenVal.Eclate:=Eclate ; GenVal.DeGroupeRemise:=DeGroupeRemise ;
io:=Transactions(GenVal.G_GenereParTiers,1) ;
GenVal.Free ;
Case io of
    oeOk : BEGIN
							SetPieceprecedenteMorte(TOBSource);
    					G_MajEvent ;
    			 END;
    oeUnknown : BEGIN
                MessageAlerte('ATTENTION : La génération ne s''est pas complètement effectuée') ;
                END ;
    oeSaisie : BEGIN
               MessageAlerte('ATTENTION : Certaines pièce déjà en cours de traitement n''ont pas été enregistrées !') ;
               END ;
    oeLettrage : BEGIN
                 MessageAlerte('ATTENTION : Certaines pièces ne peuvent pas passer en comptabilité et n''ont pas été enregistrées !') ;
                 END ;
    END ;
{Libérations}
G_FreeLesTOB ;
{Compte rendu}
if ((PremNum>0) and (io=oeOk)) then
   BEGIN
   StRendu:='Le traitement s''est correctement effectué. #13#10'+IntToStr(NbP)+' pièce générée de nature '
           +RechDom('GCNATUREPIECEG',NewNat,False)+' de N° '+IntToStr(PremNum) ;
   PGIInfo(StRendu,'Compte-rendu de génération') ;
   Result:=ClePiece;
   END ;
END ;

end.
