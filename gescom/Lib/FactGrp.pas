unit FactGrp ;

interface

uses {$IFDEF VER150} variants,{$ENDIF} HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DB, Fe_Main,
      {$IFDEF V530} EdtDOC,{$ELSE} EdtRdoc, {$ENDIF}
{$ENDIF}
{$IFDEF BTP}
     // Modif BTP
     FactOuvrage, BTPUtil,FactAcompte, FactGrpBtp,
     (* CONSO *)
     UtilPhases,UtilGrpBtp,FactTobBtp,
     (* ------ *)
     FactAdresseBTP,Factrgpbesoin,FactTvaMilliem,
     // ----
{$ENDIF}
     SysUtils, Dialogs, SaisUtil, UtilPGI, AGLInit, FactUtil, UtilSais,
     EntGC, Classes, HMsgBox, SaisComm, FactComm, ed_tools, FactNomen, UtilGrp,
     FactCpta, FactCalc, ParamSoc, UtilArticle,  Stockutil,
{$IFDEF AFFAIRE}
      FactActivite, AffaireUtil, FactAffaire,
{$ENDIF}
     Factspec,FactRg, FactTOB, FactPiece, FactArticle, FactTiers, FactAdresse,FactureBtp,
     FactLotSerie, FactContreM,BTStructChampSup,
     DepotUtil,UtilTOBPiece,uEntCommun,
     UCotraitance,
     UspecifPOC ; 

Type R_NumPTiers = RECORD
								Nature : string;
                Souche : string;
                DatePiece : TdateTime;
                NbPiece,NumDebP,NumFinP : integer;
                END ;



Function RegroupeLesPieces ( TOBSource : TOB ; NewNat : String ; Eclate,DeGroupeRemise,AvecComment : boolean ; ChoixEclateDomaine : integer ;
                              NewDate : TDateTime = 0; AvecCompteRendu : Boolean = True; AlimListePiece : Boolean = False;ContinuOnError : Boolean = False;RepriseAffaireEntete:boolean=false;
                              TraitFinTravaux:boolean=false;CodeAffCde:string='';WithVariantes: boolean=false;NewAffaire:string='';ModeTrait : string='';
                              MontantHtRef : Double = 0; MontantTTCRef : double= 0; NumfacRef : string='') : integer ;
Function  TransfoBatchPiece ( OldNat,NewNat,Souche : String ; Numero,Indice : integer ; NewDate : TDateTime; AvecComment : boolean = False; AvecCompteRendu : boolean = False ) : integer ;
// Modif LSA le 18/06/03
//Function  CreerPiecesFromLignes(TOBLigne:TOB; GenereMode:String;CompteRendu:boolean=true; TOBResult : TOB=nil):Boolean;
Function  CreerPiecesFromLignes(TOBLigne:TOB; GenereMode:String; DatePriseEncompte : TDateTime;CompteRendu:boolean=true;WithLiaisonDoc : boolean=false; TOBResult : TOB=nil; TOBDESOUVRAGES : TOB=nil):Boolean;
function  AjouteLignesToPiece (TOBLIgnes,ThePiece : TOB; GenereMode:String;WithLiaisonDoc : boolean=false;CompteRendu : boolean=true;TOBResult : TOB=nil) : boolean;
Procedure DupliqueLaPiece ( TOBPiece, TOBAffNew : TOB ; NewNat, NewTiers : String ;AvecCompteRendu,bEscompteTiers,bModeRegleTiers :Boolean ; NewDate:TDateTime=0) ;
Procedure G_ChargeLesLignes ( TOBPiece : TOB; ModeGenere : string='' ) ; //MODIFBTP
Procedure G_MajAnalLigne ( TOBL : TOB ; RefA : String ) ; // Modif BTP
Procedure AjouteUnArticle ( TOBL,TOBArticles,TOBCaTalogu : TOB ) ;
function G_DupliquePieceBTP ( var NewCleDoc: R_CLEDOC) : boolean;
Procedure G_LigneCommentaire ( TOBPiece : TOB ) ;

implementation

uses
FactVariante
{$IFDEF GPAOLIGHT}
  ,wOrdreCMP
{$ENDIF GPAOLIGHT}
{$IFDEF EDI}
  ,EDITiers
  ,EDITransfert
{$ENDIF EDI}
  ,FactTarifs
  ,PiecesRecalculs
  ,FactDomaines
  ,FactTimbres
  ,CalcOLEGenericBTP
  ,UTofListeInv
  ,BTGENODANAL_TOF
  ,AGLInitBTP
  ,UmessageIS
  ,UCumulCollectifs
  ;

type

	T_CledCFA = record
  							MontantHt : Double ;
                MontantTTC : Double;
                NumfacFou : string;
              end;

Const FromShellOLE : boolean = False ;
			SansStock : boolean = false;

Var NewNature,VenteAchat : String ;
    NewNum,IndiceNomen,PremNum,LastNum,NbP,EvDep,IndiceSerie : integer ;
    DEV    : RDEVISE ;
    TOBGenere,TOBArticles,TOBTiers,TOBRef,TOBRefPiece,TOBEches,TOBNomenclature,TOBAdresses,TOBTarif,
    TOBCpta,TOBAnaP,TOBAnaS,TOBBases,TOBBasesL,TOBAcomptes,TOBPorcs,TOBComms, TOBCatalogu, TOBListePiece,TOBSerie : TOB ;
    TobLigneTarif: TOB;
    NeedVisa,Commentaires,EclateDomaine,GereActivite,DelActivite,bDuplicPiece,EclateParPiece,
    bContinuOnError,bContremarque,ForceNumero : boolean ;
    MontantVisa  : double ;
    TTiers,ChampsRupt,PiecesG             : TStrings ;
    LaNewDate                             : TDateTime ;
    LesNumPieceTiers : Array of R_NumPTiers;
    // Modif BTP
    TOBOuvrage,TOBOuvragesP,TOBLIENOLE,TOBPIECERG,TOBBasesRg,TOBLiaison,TOBPieceTrait,TOBSSTRAIT,TOBTimbres : TOB;
    indiceRg : Integer;
    RepriseAffaireRef : boolean;
    TOBAcomptesGen,TOBAcomptesGen_O : TOB;
    FinTravaux : boolean;
    WithLiaisonInterDoc : boolean;
    // Modif MODE
    CodeAff: string;
    LiaisonOrli: boolean;
    // Modif affaire
    TobRevision, TobVariable: TOB;
    TOBMessages : TOB;
    TOBVTECOLL : TOB;
    // -- génération des factures fournisseurs
{$IFDEF BTP}
    GestionPhase : TGestionphase;
		TheRepartTva : TREPARTTVAMILL;
{$ENDIF}
		CleCFA : T_CledCFA;
    // -----

Type T_GenVal = Class
                TOBSource : TOB ;
                CodeTiers : String ;
                DeGroupeRemise : Boolean ;
                ChoixEclateDomaine : integer ;
                WithVariantes : boolean;
                NewAffaire : string;
                ModeTraitement : string;
                Procedure G_GenereParTiers ;
                procedure AddLesSupEnteteBtp(TOBPIECE: TOB);
                End ;

Type T_GenereLesPieces = Class(TObject)
     TOBSource : TOB ;
     MajPrix:boolean;
     GenereMode : string;
     // Modif BTP
     OldTypePiece : string;
     UpdateDev : boolean;
     RendActive : boolean;
     GestionReliquat : boolean;
     TOBResult : TOB;
     WithLiaison : boolean;
     TOBDESOUVRAGES : TOB;
     TOBDESSSTRAIT : TOB;
     // --
     Procedure GenereLesPieces ;
     destructor destroy ; override;
     constructor create;
     End ;

Type
  TPieceModifie = class
     MajPrix:boolean;
     // Modif BTP
     UpdateDev : boolean;
     GestionReliquat : boolean;
     TOBResult : TOB;
     WithLiaison : boolean;
     TOBAffaire : TOB;
     // --
     Procedure CompleteLaPiece ;
     destructor destroy ; override;
  end;

Type T_DuplicVal = Class
                TOBPiece, TobAffNew : TOB ;
                CodeTiers, NewCodeTiers, NewCodeAff : String;
                EscompteTiers , ModeRegleTiers : Boolean;
                Procedure G_ValideDuplicPiece ;
                End ;

Type T_ValideNumP = Class
     TOBSource : TOB ;
     CodeTiers : String;
     DatePiece : TDateTime;
     Procedure G_ValideNumPieceTiers ;
     Procedure PrepareNumPieceTiers;
     End ;

procedure G_GenereFraisChantier;
var cledoc,CledocPiece : R_CLEDOC;
    newClef : string;
    Requete : string;
begin
	if Pos(TOBgenere.getValue('GP_NATUREPIECEG'),'DBT;BCE;') = 0 then exit;
	RetrouvePieceFraisGRPBtp (cledoc,TOBgenere,false);
  if Cledoc.NumeroPiece = 0 then exit;
  if G_DupliquePieceBTP (cledoc) then
  begin
    CledocPiece := TOB2CleDoc (TOBgenere);
    NewClef := CleDocToString (cledoc);
    Requete := 'UPDATE PIECE SET GP_PIECEFRAIS="'+NewClef+'" WHERE '+WherePiece (CleDocPiece,TtdPiece,true);
    if ExecuteSql (requete) < 1 then
    begin
      V_PGI.IOError := OeUnknown;
    end;
  end;
end;

{============================= MANIPULATION DES TOBS ==========================}
Function EcritTOBPiece(NaturePiece, Tiers, Affaire, Etab,Domaine:String; TOBL:TOB;Numpiece : integer;Remise,Escompte : double; IndiceDePiece : integer;DatePriseEnCompte : TdateTime):TOB;
Var CleDoc:R_CLEDOC;
    EnHT, SaisieContre:Boolean;
BEGIN
  FillChar(CleDoc,Sizeof(CleDoc),#0);

  CleDoc.NaturePiece:=NaturePiece ;

  if DatePriseEnCOmpte <> iDate1900 then
     CleDoc.DatePiece:=DatePriseEnCompte
  else
     CleDoc.DatePiece:=V_PGI.DateEntree;

  CleDoc.Souche:=GetSoucheG(CleDoc.NaturePiece,Etab,Domaine) ;
  CleDoc.Indice:=0 ;
  EnHT:=False ;
  SaisieContre:=False;

  Result:=CreerTOBPieceVide(CleDoc, Tiers, Affaire, Etab,Domaine, EnHT, SaisieContre,Numpiece);

  // Modif BTP
  if Result <> nil then
     begin
{$IFDEF BTP}
     if NaturePiece = 'FPR' then result.PutValue('GP_AFFAIREDEVIS', Affaire);
{$ENDIF}
     Result.PutValue('GP_DEPOT', TOBL.GetValue('GL_DEPOT'));
     Result.PutValue('GP_REMISEPIED', Remise);
     Result.PutValue('GP_ESCOMPTE', Escompte);
     if (IndiceDePiece <> 0) and (Numpiece = -1) then Result.PutValue('GP_NUMERO', Result.GetValue('GP_NUMERO')+IndiceDePiece);
  end;

  // --
END;

Procedure EcritTOBLigne(NewTOBL, OldTOBL, NewTOBP:TOB; iLigne:Integer);

  procedure DupliqueLaLigne (NewTOBL,OldTOBL : TOB) ;
  var Indice : integer;
      NomChamp : string;
  begin
    for indice := 1 to NewTOBL.NbChamps do
    begin
      NomChamp := NewTOBL.GetNomChamp(indice);
      if OldTOBL.fieldExists (NomChamp) then
      begin
        NewTOBL.putValue(NomChamp,OldTOBL.GetValue(NomChamp));
      end;
    end;
    // Champs Suplémentaire
    for indice := 1000 to 1000 + NewTOBL.ChampsSup.Count - 1 do
    begin
      NomChamp := NewTOBL.GetNomChamp(indice);
      if OldTOBL.fieldExists (NomChamp) then
      begin
        NewTOBL.putValue(NomChamp,OldTOBL.GetValue(NomChamp));
      end;
    end;
    //
  end;
  
BEGIN
DupliqueLaLigne (NewTOBL,OldTOBL);
//NewTOBL.Dupliquer(OldTOBL, False, True, False);
PieceVersLigne(NewTOBP, NewTOBL);
NewTOBL.PutValue('GL_DOMAINE', oldTOBL.GetValue('GL_DOMAINE')) ;
NewTOBL.PutValue('GL_NUMLIGNE', iLigne) ;
//NewTOBL.PutValue('GL_NUMORDRE', iLigne) ;
NewTOBL.PutValue('GL_NUMORDRE', 0) ;
NewTOBL.PutValue('GL_DEPOT', OldTOBL.GetValue('GL_DEPOT')) ;
END;

Procedure G_CreerLesTOB ;
BEGIN
TOBMessages := TOB.Create ('LES MESSAGES',nil,-1);
TOBGenere:=TOB.Create('PIECE',Nil,-1) ;
TOBRefPiece:=TOB.Create('PIECE',Nil,-1) ;
// Modif BTP
AddLesSupEntete (TOBGenere);AddLesSupEntete (TOBRefPiece);
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
// Modif BTP
TOBOuvrage := TOB.create ('OUVRAGES',nil,-1);
TOBOuvragesP := TOB.create ('OUVRAGES PLAT',nil,-1);
TOBLienOLE := TOB.create ('LIENSOLE',nil,-1);
TOBPIECERG := TOB.create ('RETENUES',nil,-1);
TOBBasesRG := TOB.create ('BASESRG',nil,-1);
TOBLiaison := TOB.Create ('LES LIENS DEV-CHA',nil,-1);
TOBVTECOLL := TOB.Create ('LES CUMUL COLLECTIF',nil,-1);
// -----
//Affaire-ONYX
TOBRevision := TOB.Create('Les révisions', nil, -1);
TOBVariable := TOB.Create('Les variables', nil, -1);
//
TOBComms:=TOB.Create('COMMERCIAUX',Nil,-1) ;
TOBCpta:=CreerTOBCpta ;
TOBAnaP:=TOB.Create('',Nil,-1) ; TOBAnaS:=TOB.Create('',Nil,-1) ;
TOBAdresses:=TOB.Create('',Nil,-1);
TOBTarif:=TOB.Create('TARIF',Nil,-1) ;
TOBLigneTarif := TOB.Create('_LIGNETARIF_', Nil, -1);
PutLesSupEnteteLigneTarif(TobGenere, TobLigneTarif, nil);
PutTobPieceInTobLigneTarif(TobGenere, TobLigneTarif);
TOBListePiece := TOB.Create('',Nil,-1);
TTiers:=TStringList.Create ;
PiecesG:=TStringList.Create ;
ChampsRupt:=TStringList.Create ;
// modif BTP
TOBAcomptesGen := TOB.Create ('LESACOMPTES',nil,-1);
TOBAcomptesGen_O := TOB.create ('ACOMPTESORIG',nil,-1);
TOBPieceTrait := TOB.Create('LES PIECES TRAIT',nil,-1);
TOBSSTrait := TOB.Create('LES CO TRAIT',nil,-1);
{$IFDEF BTP}
TheRepartTva := TREPARTTVAMILL.create (nil);
{$ENDIF}
// --
END ;

Procedure G_FreeLesTOB (AlimListePiece : Boolean);
BEGIN
TOBMessages.free;
TOBGenere.Free ; TOBRefPiece.Free ; TOBRef.Free ; TOBArticles.Free ;
TOBBases.Free ; TOBTiers.Free ; TOBNomenclature.Free ; TOBBasesL.free;
// Modif BTP
TOBOuvrage.free; TOBOUvragesP.free; TOBLIENOLE.free; TOBPIECERG.free; TOBBasesRG.free;
TOBLiaison.free;
TOBVTECOLL.free;
// ----
TOBAdresses.Free ; TOBCataLogu.Free; TOBSerie.Free;
TOBCpta.Free ; TOBAnaP.Free ; TOBAnaS.Free ; TOBEches.Free ;
TOBAcomptes.Free ; TOBPorcs.Free ; TOBComms.Free ;
if Not(AlimListePiece) then TobListePiece.Free;
TTiers.Clear ; TTiers.Free ; PiecesG.Free ;
TOBTarif.Free ;
TOBLigneTarif.Free; 
ChampsRupt.Clear ; ChampsRupt.Free ;
LesNumPieceTiers:=Nil;
// Modif BTP
TOBAcomptesGen.free;
TOBAcomptesGen_O.Free ;
//Affaire-ONYX
TOBRevision.free;
TOBVariable.free;
// --
{$IFDEF BTP}
TheRepartTva.free;
{$ENDIF}
TOBPieceTrait.Free;
TOBSSTrait.Free;
END ;


Procedure G_ChargeChampsRupt(OldNat : string) ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT GPR_CONDGRP FROM PARPIECEGRP WHERE GPR_NATUREPIECEG="'+NewNature+
           '" AND GPR_NATPIECEGRP="' +OldNat+ '"',True,-1, '', True) ;
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
{$IFDEF BTP}
     DefiniChampRuptBtp (ChampsRupt);
{$ELSE}
   with ChampsRupt do
        BEGIN
        Add('PIECE;GP_TIERS;=;')        ; Add('PIECE;GP_DEVISE;=;');
        Add('PIECE;GP_TAUXDEV;=;')      ; Add('PIECE;GP_COTATION;=;');
        Add('PIECE;GP_SAISIECONTRE;=;') ; Add('PIECE;GP_REGIMETAXE;=;');
        Add('PIECE;GP_ESCOMPTE;=;')     ; Add('PIECE;GP_FACTUREHT;=;');
        Add('PIECE;GP_TIERSPAYEUR;=;')  ; Add('PIECE;GP_TVAENCAISSEMENT;=;');
        END;
{$ENDIF}
   END ;
Ferme(Q) ;
END ;


//FV1 : 17/06/2015 - FS#1393 - SOTRELEC - problème génération d'avoir global à partir facture sans code affaire en en-tête (Debut)
Procedure G_ChargeChampsRuptAvoirGlobal(OldNat : string) ;
BEGIN
  with ChampsRupt do
  BEGIN
    Add('PIECE;GP_TIERS;=;');
    Add('PIECE;GP_DEVISE;=;');
    Add('PIECE;GP_TAUXDEV;=;');
    Add('PIECE;GP_COTATION;=;');
    Add('PIECE;GP_SAISIECONTRE;=;');
    Add('PIECE;GP_REGIMETAXE;=;');
    Add('PIECE;GP_ESCOMPTE;=;');
    Add('PIECE;GP_FACTUREHT;=;');
    Add('PIECE;GP_TIERSPAYEUR;=;');
    Add('PIECE;GP_TVAENCAISSEMENT;=;');
  END;
END ;
//FV1 : 17/06/2015 - FS#1393 - SOTRELEC - problème génération d'avoir global à partir facture sans code affaire en en-tête (fin)

procedure EcritlesOuvragesP (natureG : string; TOBOuvragesP : TOB);
var PassP : string;
begin
if GetParamSocSecur('SO_GCDESACTIVECOMPTA',false) then exit; // SI pas de compta on sort
PassP:=GetInfoParPiece(NatureG,'GPP_TYPEECRCPTA') ; if ((PassP='') or (PassP='RIE')) then Exit ;
if not TOBOuvragesP.InsertDBByNivel(false) then V_PGI.ioerror := OeUnknown;
end;

Procedure G_InitGenere ( ChoixEclateDomaine : integer ) ;
Var Q : TQuery ;
BEGIN
NeedVisa:=(GetInfoParPiece(NewNature,'GPP_VISA')='X') ;
MontantVisa:=GetInfoParPiece(NewNature,'GPP_MONTANTVISA') ;
VenteAchat:=GetInfoParPiece(NewNature,'GPP_VENTEACHAT') ;
if ChoixEclateDomaine=1 then EclateDomaine:=True else
 if ChoixEclateDomaine=0 then EclateDomaine:=False else
    EclateDomaine:=(GetInfoParPiece(NewNature,'GPP_ECLATEDOMAINE')='X') ;
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
OldL:=-1 ; //IndiceNomen:=1 ; MR : modif déjà effectué en version 535 annulée depuis !!!!!!!!!
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
{$IFDEF BTP}
procedure RecupereAdressesOrigine (TOBPiece,TOBAdresses : TOB);
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
LoadLesAdressesOrigine(TOBPiece,TOBAdresses) ;
TOBAdresses.SetAllModifie(True);
NumL:=TOBPiece.GetValue('GP_NUMADRESSELIVR'); NumF:=TOBPiece.GetValue('GP_NUMADRESSEFACT');
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
end;
{$ENDIF}

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
if not bContremarque then LoadLesAdresses(TOBPiece,TOBAdresses) ;
TOBAdresses.SetAllModifie(True);
NumL:=TOBPiece.GetValue('GP_NUMADRESSELIVR'); NumF:=TOBPiece.GetValue('GP_NUMADRESSEFACT');
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
{NEWPIECE}
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

{================================ ANALYTIQUES =================================}
Procedure G_LoadAnaPiece ( TOBPiece : TOB ) ;
Var Q : TQuery ;
    RefA : String ;
BEGIN
RefA:=EncodeRefCPGescom(TOBPiece) ; if RefA='' then Exit ;
Q:=OpenSQL('SELECT * FROM VENTANA WHERE YVA_TABLEANA="GL" AND YVA_IDENTIFIANT="'+RefA+'" AND YVA_IDENTLIGNE="000"',True,-1, '', True) ;
if Not Q.EOF then TOBAnaP.LoadDetailDB('VENTANA','','',Q,False) ;
Ferme(Q) ;
END ;

Procedure G_LoadAnaLigne ( TOBPiece,TOBL : TOB ) ;
Var Q : TQuery ;
    RefA : String ;
    NumL : integer ;
    TOBAL : TOB ;
BEGIN
if TOBL=Nil then Exit ;
if TOBL.Detail.Count<=0 then TOBAL:=TOB.Create('',TOBL,-1) else TOBAL:=TOBL.Detail[0] ;
RefA:=EncodeRefCPGescom(TOBPiece) ; if RefA='' then Exit ;
NumL:=TOBL.GetValue('GL_NUMLIGNE') ;
Q:=OpenSQL('SELECT * FROM VENTANA WHERE YVA_TABLEANA="GL" AND YVA_IDENTIFIANT="'+RefA+'" AND YVA_IDENTLIGNE="'+FormatFloat('000',NumL)+'" ORDER BY YVA_AXE, YVA_NUMVENTIL',True,-1, '', True) ;
if Not Q.EOF then TOBAL.LoadDetailDB('VENTANA','','',Q,False) ;
Ferme(Q) ;
END ;

{==================================== LIGNES ==================================}
Procedure G_LigneCommentaire ( TOBPiece : TOB ) ;
Var NewL,TOBL : TOB ;
    RefP : String ;
BEGIN
if TOBPiece.Detail.Count>0 then TOBL:=TOBPiece.Detail[0] else TOBL:=Nil ;
NewL:=NewTOBLigne(TOBPiece,1) ; if TOBL<>Nil then NewL.Dupliquer(TOBL,False,True) ;
NewL.PutValue('GL_NUMORDRE', 0) ;
RefP:=RechDom('GCNATUREPIECEG',TOBPiece.GetValue('GP_NATUREPIECEG'),False)
   +' N° '+IntToStr(TOBPiece.GetValue('GP_NUMERO'))
   +' du '+DateToStr(TOBPiece.GetValue('GP_DATEPIECE'))
   +'  '+TOBPiece.GetValue('GP_REFINTERNE') ;
RefP:=Copy(RefP,1,70) ;
NewL.PutValue('GL_LIBELLE',RefP)    ; NewL.PutValue('GL_TYPELIGNE','COM') ;
NewL.PutValue('GL_TYPEDIM','NOR')   ; NewL.PutValue('GL_CODEARTICLE','') ;
NewL.PutValue('GL_ARTICLE','')      ; NewL.PutValue('GL_QTEFACT',0) ;
NewL.PutValue('GL_QTESTOCK',0)      ; NewL.PutValue('GL_PUHTDEV',0) ;
NewL.PutValue('GL_QTERESTE',0)      ; { NEWPIECE }
// --- GUINIER ---
NewL.PutValue('GL_MTRESTE',0)       ; { NEWPIECE }

NewL.PutValue('GL_PUTTCDEV',0)      ; NewL.PutValue('GL_TYPEARTICLE','') ;
NewL.PutValue('GL_PUHT',0)          ; NewL.PutValue('GL_PUHTNET',0) ;
NewL.PutValue('GL_PUTTC',0)         ; NewL.PutValue('GL_PUTTCNET',0) ;
NewL.PutValue('GL_PUHTBASE',0)      ; NewL.PutValue('GL_FAMILLETAXE1','') ;
NewL.PutValue('GL_TYPENOMENC','')   ; NewL.PutValue('GL_QUALIFMVT','') ;
NewL.PutValue('GL_REFARTSAISIE','') ; NewL.PutValue('GL_REFARTBARRE','') ;
NewL.PutValue('GL_REFCATALOGUE','') ; NewL.PutValue('GL_TYPEREF','') ;
NewL.PutValue('GL_REFARTTIERS','')  ;
{Modif AC 4/07/03 Pas de GL_CODESDIM sur les lignes commentaire}
NewL.PutValue('GL_CODESDIM','')  ;
{Fin Modif AC}
{Modif JLD 20/06/2002}
NewL.PutValue('GL_ESCOMPTE',TOBPiece.GetValue('GP_ESCOMPTE')) ;
NewL.PutValue('GL_REMISEPIED',TOBPiece.GetValue('GP_REMISEPIED')) ;
{Fin modif}

//JS 17/06/03
NewL.PutValue('GL_INDICESERIE',0) ; NewL.PutValue('GL_INDICELOT',0) ;
NewL.PutValue('GL_REMISELIGNE',0) ;
// Modif BTP
NewL.PutValue('GL_TYPEARTICLE','EPO');
NewL.PutValue('GL_PUHTNETDEV',0)    ; NewL.PutValue('GL_PUTTCNETDEV',0) ;
NewL.PutValue('GL_BLOCNOTE','')    ; NewL.PutValue('GL_QUALIFQTEVTE','') ;
NewL.PutValue('GL_INDICENOMEN',0) ;
// ---
ZeroLigne(NewL) ;
// Modif MODE : marquage de la ligne commentaire ajoutée à la pièce d'origine
if CodeAff <> '' then
   BEGIN
   NewL.AddChampSupValeur('LIGNECOMMENT', 'X');
   if NewL.FieldExists('GEL_QTEAFFECTEE') then NewL.PutValue('GEL_QTEAFFECTEE', 0);
   END;
// ---
NewL.PutValue('GL_FACTUREHT',TOBPiece.GetValue('GP_FACTUREHT')) ;
NewL.PutValue('GL_DEVISE',TOBPiece.GetValue('GP_DEVISE')) ;
NewL.PutValue('GP_TAUXDEV',TOBPiece.GetValue('GP_TAUXDEV')) ;
NewL.PutValue('GL_COTATION',TOBPiece.GetValue('GP_COTATION')) ;
NewL.PutValue('GL_TIERS',TOBPiece.GetValue('GP_TIERS')) ;
NewL.PutValue('GL_TIERSFACTURE',TOBPiece.GetValue('GP_TIERSFACTURE')) ;
NewL.PutValue('GL_TIERSPAYEUR',TOBPiece.GetValue('GP_TIERSPAYEUR')) ;
NewL.PutValue('GL_TIERSLIVRE',TOBPiece.GetValue('GP_TIERSLIVRE')) ;
NewL.PutValue('GL_REGIMETAXE',TOBPiece.GetValue('GP_REGIMETAXE')) ;

END ;

Procedure G_ChargeLesAcomptes ( TOBPiece : TOB ) ;
Var Q : TQuery ;
    CD : R_CleDoc ;
BEGIN
CD:=TOB2CleDoc(TOBPiece) ;
Q:=OpenSQL('SELECT * FROM ACOMPTES WHERE '+WherePiece(CD,ttdAcompte,False),True,-1, '', True) ;
if Not Q.EOF then TOBAcomptes.LoadDetailDB('ACOMPTES','','',Q,True,False) ;
Ferme(Q) ;
END ;

Procedure G_SommeLesPorcs ( TOBPL,TOBP : TOB ) ;
Var X : Double ;
BEGIN
X:=TOBP.GetValue('GPT_BASEHT')      ; X:=Arrondi(X+TOBPL.GetValue('GPT_BASEHT'),9)      ; TOBP.PutValue('GPT_BASEHT',X) ;
X:=TOBP.GetValue('GPT_BASEHTDEV')   ; X:=Arrondi(X+TOBPL.GetValue('GPT_BASEHTDEV'),9)   ; TOBP.PutValue('GPT_BASEHTDEV',X) ;
X:=TOBP.GetValue('GPT_BASETTC')     ; X:=Arrondi(X+TOBPL.GetValue('GPT_BASETTC'),9)     ; TOBP.PutValue('GPT_BASETTC',X) ;
X:=TOBP.GetValue('GPT_BASETTCDEV')  ; X:=Arrondi(X+TOBPL.GetValue('GPT_BASETTCDEV'),9)  ; TOBP.PutValue('GPT_BASETTCDEV',X) ;
X:=TOBP.GetValue('GPT_TOTALHT')     ; X:=Arrondi(X+TOBPL.GetValue('GPT_TOTALHT'),9)     ; TOBP.PutValue('GPT_TOTALHT',X) ;
X:=TOBP.GetValue('GPT_TOTALHTDEV')  ; X:=Arrondi(X+TOBPL.GetValue('GPT_TOTALHTDEV'),9)  ; TOBP.PutValue('GPT_TOTALHTDEV',X) ;
X:=TOBP.GetValue('GPT_TOTALTTC')    ; X:=Arrondi(X+TOBPL.GetValue('GPT_TOTALTTC'),9)    ; TOBP.PutValue('GPT_TOTALTTC',X) ;
X:=TOBP.GetValue('GPT_TOTALTTCDEV') ; X:=Arrondi(X+TOBPL.GetValue('GPT_TOTALTTCDEV'),9) ; TOBP.PutValue('GPT_TOTALTTCDEV',X) ;
if TOBPL.GetValue('GPT_TYPEPORT')='MT' then
    TOBP.PutValue('GPT_MONTANTMINI',TOBP.GetValue('GPT_TOTALHTDEV')) ;
END ;

Procedure G_ChargeLesPorcs ( TOBPiece : TOB ) ;
Var Q : TQuery ;
    CD : R_CleDoc ;
    i,k : integer ;
    PorcsLoc,TOBP,TOBPL : TOB ;
    FTi : Array[1..5] of String ;
    CodePort,TypePort : String ;
BEGIN
PorcsLoc:=TOB.Create('LESPORCS',Nil,-1) ; CD:=TOB2CleDoc(TOBPiece) ;
Q:=OpenSQL('SELECT * FROM PIEDPORT WHERE '+WherePiece(CD,ttdPorc,False),True,-1, '', True) ;
if Not Q.EOF then PorcsLoc.LoadDetailDB('PIEDPORT','','',Q,True,False) ;
Ferme(Q) ;
for i:=0 to PorcsLoc.Detail.Count-1 do
    BEGIN
    TOBPL:=PorcsLoc.Detail[i] ;
    CodePort:=TOBPL.GetValue('GPT_CODEPORT') ;
    TypePort:=TOBPL.GetValue('GPT_TYPEPORT') ;
    for k:=1 to 5 do FTi[k]:=TOBPL.GetValue('GPT_FAMILLETAXE'+IntToStr(k)) ;
    TOBP:=TOBPorcs.FindFirst(['GPT_CODEPORT','GPT_TYPEPORT','GPT_FAMILLETAXE1','GPT_FAMILLETAXE2','GPT_FAMILLETAXE3','GPT_FAMILLETAXE4','GPT_FAMILLETAXE5'],
                             [CodePort,TypePort,FTi[1],FTi[2],FTi[3],FTi[4],FTi[5]],True) ;
    if TOBP<>Nil then
       G_SommeLesPorcs(TOBPL,TOBP)
       else
       BEGIN
       TOBP:=TOB.Create('PIEDPORT',TOBPorcs,-1);
       TOBP.Dupliquer(PorcsLoc.Detail[i],True,True);
       END;
    END ;
PorcsLoc.Free ;
END ;

// Modif BTP
procedure G_ChargeLesRG(TOBL:TOB) ;
Var Q : TQuery ;
    CD : R_CleDoc ;
    i  : integer ;
    TOBDocP,TOBInsere,TOBPiece : TOB ;
    CautionU,CautionUdev : double;
BEGIN
if TOBL.getValue('GL_TYPELIGNE')<>'RG' then exit;
TOBPiece := TOBL.Parent;
TOBDocP := TOB.create ('LATOB',nil,-1);
CD:=TOB2CleDoc(TOBL) ;
Q:=OpenSQL('SELECT * FROM PIECERG WHERE '+WherePiece(CD,ttdRetenuG,False)+' AND PRG_NUMLIGNE='+inttostr(TOBL.GetValue('GL_NUMORDRE')),True,-1, '', True) ;
if Not Q.EOF then TOBDocP.LoadDetailDB('PIECERG','','',Q,True,False) ;
Ferme(Q) ;
for i:= 0 to TOBDocP.detail.count -1 do
    begin
    TOBInsere:=TOB.create ('PIECERG',TOBPieceRG,-1);
    TOBInsere.Dupliquer (TOBDocP.detail[i],true,true);
    TOBInsere.AddChampSupValeur ('INDICERG',IndiceRG);
    TOBInsere.AddChampSupValeur ('PIECEPRECEDENTE',encoderefPiece(TOBDOCP.detail[i],0));
    TOBInsere.AddChampSupValeur ('NUMORDRE',0,false);
    TOBInsere.putValue('PRG_CAUTIONMTU',0);
    TOBInsere.putValue('PRG_CAUTIONMTUDEV',0);
//
    CautionU :=0; CautionUdev := 0;
    GetCautionAlreadyUsed (TOBPiece,TOBL,TOBL.GEtValue('GL_PIECEPRECEDENTE'),CautionU,CautionUdev);
    (*
    CautionU := CautionU - TOBDOCP.detail[i].GEtValue('PRG_MTTTCRG');
    if CautionU < 0 then CautionU := 0;
    CautionUDev := CautionUDev - TOBDOCP.detail[i].GEtValue('PRG_MTTTCRGDEV');
    if CautionUDev < 0 then CautionUDev := 0;
    *)
    TOBInsere.AddChampSupValeur ('CAUTIONUTIL',CautionU);
    TOBInsere.AddChampSupValeur ('CAUTIONUTILDEV',CautionUDev);
    TOBInsere.AddChampSupValeur ('CAUTIONAPRES',0);
    TOBInsere.AddChampSupValeur ('CAUTIONAPRESDEV',0);
//
    TOBL.PutValue ('INDICERG',IndiceRG);
    inc(IndiceRg);
    end;
TOBDocP.free;
end;
// --

Procedure G_TrierDomaine ( TOBPiece : TOB ) ;
Var i,iDom,iType,PremArt : integer ;
    sDom,PremDom,sType   : String ;
    TOBL   : TOB ;
BEGIN
iType := 0;
iDom := 0;
if TOBPiece=Nil then Exit ;
if Not EclateDomaine then Exit ;
if TOBPiece.Detail.Count<=1 then Exit ;
PremDom:='@@@' ; PremArt:=-1 ;
for i:=0 to TOBPiece.Detail.Count-2 do
    BEGIN
    TOBL:=TOBPiece.Detail[i] ;
    if i=0 then BEGIN iDom:=TOBL.GetNumChamp('GL_DOMAINE') ; iType:=TOBL.GetNumChamp('GL_TYPELIGNE') ; END ;
    sType:=TOBL.GetValeur(iType) ;
    if sType='ART' then
       BEGIN
       sDom:=TOBL.GetValue('GL_DOMAINE') ;
       if PremDom='@@@' then PremDom:=sDom ;
       if PremArt=-1 then PremArt:=i ;
       END else
       BEGIN
       TOBL.PutValeur(iDom,sDom) ;
       END ;
    END ;
for i:=0 to PremArt-1 do TOBPiece.Detail[i].PutValue('GL_DOMAINE',PremDom) ;
TOBPiece.Detail.Sort('GL_DOMAINE;GL_NUMLIGNE;') ;
END ;

// Modif MODE
procedure G_ChargeLignesAffCde(TOBPiece: TOB);
var Q: TQuery;
  CD: R_CleDoc;
  sSql, RefArt, RefDim, Stg: string;
  ii, NumL: integer;
  qte: double;
  TOBAff, TOBA, TOBL: TOB;
begin
  if (TOBPiece = nil) or (TOBPiece.Detail.Count <= 0) then Exit;
  CD := TOB2CleDoc(TOBPiece);
  {Chargement des quantités affectées des lignes de la pièce}
  sSql := 'SELECT GEL_NUMORDRE,GEL_QTEAFFECTEE,GEL_ARTICLE,GEL_CODESDIM FROM AFFCDELIGNE WHERE '
    + ' GEL_NATUREPIECEG="' + CD.NaturePiece + '" AND GEL_SOUCHE="' + CD.Souche + '" '
    + ' AND GEL_NUMERO=' + IntToStr(CD.NumeroPiece)
    + ' AND GEL_INDICEG=' + IntToStr(CD.Indice) + ' AND GEL_CODEAFF="' + CodeAff + '"'
    + ' ORDER BY GEL_NUMORDRE';
  TOBAff := TOB.Create('Les AffCde', nil, -1);
  Q := OpenSQL(sSql, True,-1, '', True);
  if not Q.EOF then TOBAff.LoadDetailDB('AFFCDELIGNE', '', '', Q, True, True);
  Ferme(Q);
  for ii := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[ii];
    TOBL.AddChampSupValeur('GEL_QTEAFFECTEE', 0);
    NumL := TOBL.GetValue('GL_NUMORDRE');
    RefArt := TOBL.GetValue('GL_ARTICLE');
    RefDim := TOBL.GetValue('GL_CODESDIM');
    TOBA := TOBAff.FindFirst(['GEL_NUMORDRE'], [NumL], False);
    if TOBA <> nil then
    begin
      if (TOBA.GetValue('GEL_ARTICLE') = RefArt) and
        (TOBA.GetValue('GEL_CODESDIM') = RefDim) then
      begin
        qte := TOBA.GetValue('GEL_QTEAFFECTEE');
        TOBL.PutValue('GEL_QTEAFFECTEE', qte);
      end else
      begin
        Stg := DateTimeToStr(Now) + ' Commande ' + IntToStr(CD.NumeroPiece)
          + ' Indice ' + IntToStr(CD.Indice) + ' Ligne ' + IntToStr(NumL)
          + ' Quantité commandée=' + StrF00(TOBL.GetValue('GL_QTESTOCK'), V_PGI.OkDecQ);
        if RefArt <> '' then Stg := Stg + ' Article ' + RefArt else Stg := Stg + ' Générique ' + RefDim;
        Stg := Stg + ' : erreur chargement quantité affectée ' + StrF00(TOBA.GetValue('GEL_QTEAFFECTEE'), V_PGI.OkDecQ);
        if TOBA.GetValue('GEL_ARTICLE') <> '' then Stg := Stg + ' sur article ' + TOBA.GetValue('GEL_ARTICLE')
        else Stg := Stg + ' sur générique ' + TOBA.GetValue('GEL_CODESDIM');
        LogAGL(Stg);
      end;
    end;
  end;
  TOBAff.Free;
end;

// NAC - Optimisation génération des prépas
Function G_GetQteAffectee(TOBLigne: TOB): double;
var
  Val: variant;
begin
  Result := 0;
  if CodeAff <> '' then
  begin
    {quantité affectée}
    if TOBLigne.FieldExists('GEL_QTEAFFECTEE') then
    begin
      Val := TOBLigne.GetValue('GEL_QTEAFFECTEE');
      if not VarIsNull(Val) then Result := VarAsType(Val, varDouble);
    end;
    if Result < 0 then Result := 0;
  end;
end;

// NAC - Optimisation génération des prépas
Function G_GetMtAffectee(TOBLigne: TOB): double;
var
  Val: variant;
begin
  Result := 0;
  if CodeAff <> '' then
  begin
    {quantité affectée}
    if TOBLigne.FieldExists('GEL_MTHTAFFECTEE') then
    begin
      Val := TOBLigne.GetValue('GEL_MTHTAFFECTEE');
      if not VarIsNull(Val) then Result := VarAsType(Val, varDouble);
    end;
    if Result < 0 then Result := 0;
  end;
end;

Procedure G_InitLiaisonOrli ;
var
  QQ: TQuery;
  sSql: string;
begin
  LiaisonOrli := False ;
  if CodeAff = '' then Exit ;
  sSql := 'SELECT GEA_LIAISONORLI FROM AFFCDEENTETE WHERE GEA_CODEAFF="'+ CodeAff +'"';
  QQ := OpenSQl(sSql, True,-1, '', True);
  if not QQ.Eof then LiaisonOrli := (QQ.FindField('GEA_LIAISONORLI').AsString = 'X');
  Ferme(QQ);
end;
// -----

Procedure G_ChargeLesLignes ( TOBPiece : TOB ; ModeGenere : string='' ) ;
Var {Q : TQuery ;}
    CD : R_CleDoc ;
BEGIN
if not bContremarque then
   begin
   CD:=TOB2CleDoc(TOBPiece) ;
(* Avant :
   Q:=OpenSQL('SELECT * FROM LIGNE WHERE '+WherePiece(CD,ttdLigne,False)+' ORDER BY GL_NUMLIGNE',True) ;
   if Not Q.EOF then TOBPiece.LoadDetailDB('LIGNE','','',Q,True,True) ;
   Ferme(Q) ;
   Apres : *)
   LoadLignes(CD, TobPiece);
   // Modif MODE
   if CodeAff <> '' then G_ChargeLignesAffCde(TOBPiece);
   // -----
   end;
// VARIANTE
	 if (ModeGenere <> 'DEVTOCHAN') and (ModeGenere <> 'ETUTODEV') and (ModeGenere <> 'VALIDETU') then SupprimeLesVariantes (TOBpiece);
// --
if TOBPiece.Detail.Count>0 then
   BEGIN
   {Trier par domaine si éclatement}
   G_TrierDomaine(TOBPiece) ;
   {Ajout ligne de commentaire}
   if Commentaires then G_LigneCommentaire(TOBPiece) ;
   {Déroulement des lignes}
   PieceAjouteSousDetail(TOBPiece,false,false,true);
   {
   AddLesSupLigne(TOBPiece.Detail[0],True) ;
   for i:=0 to TOBPiece.Detail.Count-1 do
       BEGIN
       TOBL:=TOBPiece.Detail[i] ; InitLesSupLigne(TOBL) ;
       TOB.Create('',TOBL,-1) ;
       TOB.Create('',TOBL,-1) ; // GM 070202  pour 2eme fille
       END ;
   }
   END ;
END ;

{==================================== PIECE ===================================}
Procedure G_PieceVersLigne ( TOBPiece,TOBL : TOB ) ;
BEGIN
TOBL.PutValue('GL_DATEPIECE',TOBPiece.GetValue('GP_DATEPIECE')) ;
TOBL.PutValue('GL_SOUCHE',TOBPiece.GetValue('GP_SOUCHE')) ;
TOBL.PutValue('GL_NATUREPIECEG',TOBPiece.GetValue('GP_NATUREPIECEG')) ;
TOBL.PutValue('GL_INDICEG',TOBPiece.GetValue('GP_INDICEG')) ;
TOBL.PutValue('GL_NUMERO',TOBPiece.GetValue('GP_NUMERO')) ;
TOBL.PutValue('GL_VIVANTE',TOBPiece.GetValue('GP_VIVANTE')) ;
TOBL.PutValue('GL_DATEMODIF',TOBPiece.GetValue('GP_DATEMODIF')) ;
TOBL.PutValue('GL_DATECREATION',TOBPiece.GetValue('GP_DATECREATION')) ;
TOBL.PutValue('GL_ETABLISSEMENT',TOBPiece.GetValue('GP_ETABLISSEMENT')) ;
TOBL.PutValue('GL_REMISEPIED',TOBPiece.GetValue('GP_REMISEPIED')) ;
TOBL.PutValue('GL_CREERPAR',TOBPiece.GetValue('GP_CREEPAR')) ;
// Modif Btp
if ((TOBL.GetValue('GL_TYPENOMENC') = 'OUV') or (TOBL.GetValue('GL_TYPENOMENC') = 'NOM')) and
   (GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'),'GPP_RECUPPRE') = 'SUI') then
begin
	TOBL.PutValue('GL_TYPEPRESENT',GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'),'GPP_TYPEPRESENT')) ;
end;
// --------------
END ;

Procedure G_MajAnalPiece ( RefA : String ) ;
Var i : integer ;
    TOBV : TOB ;
BEGIN
if TOBAnaP=Nil then Exit ;
for i:=0 to TOBAnaP.Detail.Count-1 do
    BEGIN
    TOBV:=TOBAnaP.Detail[i] ;
    if TOBV.NomTable='VENTANA' then TOBV.PutValue('YVA_IDENTIFIANT',RefA) ;
    END ;
END ;

Procedure G_MajAnalLigne ( TOBL : TOB ; RefA : String ) ;
Var i,NumL,k : integer ;
    TOBV,TOBAL : TOB ;
BEGIN
NumL:=TOBL.GetValue('GL_NUMLIGNE') ;
for i:=0 to TOBL.Detail.Count-1 do
    BEGIN
    TOBAL:=TOBL.Detail[i] ;
    for k:=0 to TOBAL.Detail.Count-1 do
        BEGIN
        TOBV:=TOBAL.Detail[k] ;
        if TOBV.NomTable='VENTANA' then
           BEGIN
           TOBV.PutValue('YVA_IDENTIFIANT',RefA) ;
           TOBV.PutValue('YVA_IDENTLIGNE',FormatFloat('000',NumL)) ;
           END ;
       END ;
    END ;
END ;

Procedure AjouteUnArticle ( TOBL,TOBArticles,TOBCaTalogu : TOB ) ;
Var TOBArt,TOBCata : TOB ;
    RefUnique,RefCata,RefFour : String ;
    QQ : TQuery;
BEGIN
if (TOBL.GetValue('GL_ENCONTREMARQUE')='X')  then
   Begin
   RefCata:=TOBL.GetValue('GL_REFCATALOGUE') ;
   RefFour:=GetCodeFourDCM(TOBL) ;
   TOBCata:=InitTOBCata(TOBCatalogu,RefCata,RefFour) ;
   LoadTOBDispoContreM(TOBL, TOBCata, true) ;
   end;
RefUnique:=TOBL.GetValue('GL_ARTICLE') ; if RefUnique='' then Exit ;
TOBArt:=TOBArticles.FindFirst(['GA_ARTICLE'],[RefUnique],False) ;
if TOBArt=Nil then
   BEGIN
   TOBArt:=CreerTOBArt(TOBArticles) ;
   QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
   							  'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
    							'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+RefUnique+'"',true,-1, '', True);
//   TOBArt.SelectDB('"'+RefUnique+'"',Nil) ;
	 TOBART.selectDb ('',QQ);
   InitChampsSupArticle (TOBART);
   ferme (QQ);
   END ;
If (TOBL.GetValue ('GL_TYPEARTICLE') <>'CTR') And (TOBL.GetValue ('GL_TYPEARTICLE') <>'PRE') And (TOBL.GetValue ('GL_TYPEARTICLE') <>'FRA') then // Ajout PA le 18/08/01
   LoadTOBDispo(TOBArt,True, '"' + TobL.GetValue ('GL_DEPOT') + '"') ; // DBR : Depot unique chargé
// NAC - Optimisation génération des prépas
TOBArticles.SetAllModifie(False);
END ;

Procedure G_MajLesComms ( TOBPiece : TOB ) ;
Var i : integer ;
    EtatVC : String ;
    TOBL : TOB ;
BEGIN
for i:=0 to TOBPiece.Detail.Count-1 do
    BEGIN
    TOBL:=TOBPiece.Detail[i] ;
    EtatVC:=TOBL.GetValue('GL_VALIDECOM') ;
    if EtatVC='VAL' then TOBL.PutValue('GL_VALIDECOM','AFF') else
     if EtatVC='AFF' then CommVersLigne(TOBPiece,TOBArticles,TOBComms,i+1,False) ;
    END ;
END ;

Procedure G_FinirPieceGenere (ModeGeneration : string = '');
Var i : integer ;
    TOBL : TOB ;
    RefPiece,RefA : String ;
    RemDiff,CalculPv : boolean ;
    RemCur,TotRemDev,TotMontantDev,RemMoy : double ;
BEGIN
  // MODIFBTP : Traitement des situations
  {Traitement remise pied}
  RemDiff:=False ; RemCur:=0 ;// RemMoy:=0 ;
  TotRemDev:=0 ; TotMontantDev:=0 ;
  for i:=0 to TOBGenere.Detail.Count-1 do
  BEGIN
    TOBL:=TOBGenere.Detail[i] ;
    if i=0 then RemCur:=TOBL.GetValue('GL_REMISEPIED') else
    if (TOBL.GetValue('GL_REMISEPIED')<>RemCur) and  (TOBL.GetValue('GL_REMISABLEPIED')='X') then RemDiff:=True ;
    if TOBL.GetValue('GL_REMISABLEPIED')='X' then
    BEGIN
      TotRemDev:=TotRemDev+TOBL.GetValue('GL_TOTREMPIEDDEV') ;
      TotMontantDev:=TotMontantDev+TOBL.GetValue('GL_MONTANTHTDEV') ;
    END ;
    // NEW ONE (OPTIMISATION CALCUL)
    ZeroLigneMontant (TOBL);  // reinit des montants de la ligne
    // --
  END ;
  ZeroFacture(TOBGenere) ;
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
  {$ifdef AFFAIRE}
  MemoriseNumLigne(TobGenere);

  {$endif}
  for i:=0 to TOBGenere.Detail.Count-1 do
  BEGIN
    TOBL:=TOBGenere.Detail[i] ;
    {Traçabilité}
    if Not bDuplicPiece then  // Ajout PA zones non reprise en duplication de pièce
    begin
      RefPiece:=EncodeRefPiece(TOBL) ; TOBL.PutValue('GL_PIECEPRECEDENTE',RefPiece) ;
      if TOBL.GetValue('GL_PIECEORIGINE')='' then TOBL.PutValue('GL_PIECEORIGINE',RefPiece) ;
    end;
    {Pièce vers lignes}
    UG_PieceVersLigneTarif(TobGenere, TobLigneTarif, TOBL, i + 1);
    UG_PieceVersLigne(TOBGenere,TOBL) ;

    {Divers}
    // Modif MODE
    if CodeAff = '' then
    begin
      if (ModeGeneration <> 'CBTTOLIVC') and (ModeGeneration <> 'CCTOLIVCLI') and (ModeGeneration <> 'CCTOLIVCLI') then
      begin
        TOBL.PutValue('GL_QTERELIQUAT', 0);
        TOBL.PutValue('GL_MTRELIQUAT', 0);
      end;
      TOBL.PutValue('GL_SOLDERELIQUAT', '-');
    end;
    // -----
    {Analytique}
    UG_MajAnalLigne(TOBL,RefA) ;
    TobL.PutValue ('GL_NUMORDRE', 0);
  END ;
  if isPieceGerableFraisDetail (TOBGenere.GetValue('GP_NATUREPIECEG')) or (VenteAchat='VEN') then
  begin
  (*
  	BeforeApplicFrais (TOBGenere,TOBOUvrage,false);
  	RecalculeCoefFrais (TOBGenere,TOBPorcs);   // calcul des Taux Fc et TAux FG
    TOBGenere.putValue('GP_MONTANTPR',0);
    TOBGenere.putValue('GP_MONTANTFG',0); // car deja calcule precedemment
    AppliqueFraisPiece (TOBGenere,TOBOuvrage,true,false,(TOBGenere.getVAlue('GP_FACTUREHT')='X'),DEV);
  *)
    calculPv :=  (pos(NewNature,'FBT;DAC;FBP;BAC')=0) and (ModeGeneration<>'FACDETAILCONSO');
    CalculeMontantsDoc (TOBgenere,TOBOuvrage,CalculPv);
  end;
  {Numéros lignes}
  TobGenere.putvalue('GP_CODEORDRE',0); // gm le 18/07/03  sinon on repart du numordre de la piece d'origine
  NumeroteLignesGC(Nil,TOBGenere, true);
  {$IFDEF BTP}
  EnregistreNumOrdreRG (TOBgenere,TOBPieceRg); // En vu du recalcul de la piece qui risque de repositionner le NUMORDRE
  ConstitueLienInterDoc (TOBgenere,TOBLiaison);
  {$ENDIF}
  {$ifdef AFFAIRE}
  RenumeroteVariableRev(TobGenere,TobVariable,TobRevision);
  {$endif}
END ;

Function G_GetNumPieceTiers ( TOBPiece : TOB) : Integer;
var Souche : string;
Var i: integer;
begin
Result := 0;
if Not(EclateParPiece) then Exit;
Souche := TobPiece.GetValue('GP_SOUCHE');
for i:=Low(LesNumPieceTiers)to High(LesNumPieceTiers) do
   begin
   if LesNumPieceTiers[i].Souche = Souche then
      begin
      Result := LesNumPieceTiers[i].NumDebP;
      Inc(LesNumPieceTiers[i].NumDebP);
      if Result > LesNumPieceTiers[i].NumFinP then Result := 0;
      Break;
      end;
   end;
end;

Procedure G_AlimNewPiece ( TOBPiece : TOB ; DupPiece : Boolean; ModeGenere : string='' ) ;
Var NewSouchePiece : String;
    i_ind : integer;
BEGIN
TOBGenere.ClearDetail ; TOBGenere.InitValeurs ;
if DupPiece then
	Begin
   TOBGenere.Dupliquer(TOBPiece,False,True);
	 AddLesSupEntete (TOBGenere);    // gm 12/03/2002 suite aux modifs BTP
  End
   else
   begin
   for i_ind:=1 to 9 do TOBGenere.PutValue('GP_LIBRETIERS'+inttostr(i_ind),TOBRefPiece.GetValue('GP_LIBRETIERS'+inttostr(i_ind)));
   TOBGenere.PutValue('GP_LIBRETIERSA',TOBRefPiece.GetValue('GP_LIBRETIERSA'));
   for i_ind:=1 to 3 do TOBGenere.PutValue('GP_LIBREPIECE'+inttostr(i_ind),TOBRefPiece.GetValue('GP_LIBREPIECE'+inttostr(i_ind)));
   TOBGenere.PutValue('GP_REPRESENTANT',TOBRefPiece.GetValue('GP_REPRESENTANT'));
   end;
TOBGenere.PutValue('GP_NATUREPIECEG',NewNature) ;
TOBGenere.PutValue('GP_DATEPIECE',LaNewDate) ;
TOBGenere.PutValue('GP_INDICEG',0) ;
NewSouchePiece:=GetSoucheG( NewNature, TOBGenere.GetValue('GP_ETABLISSEMENT'),TOBGenere.GetValue('GP_DOMAINE'));
TOBGenere.PutValue('GP_SOUCHE',NewSouchePiece);
if EclateParPiece then    // PCS 12/03/2003  CDISCOUNT
   NewNum := G_GetNumPieceTiers (TobGenere)
  else
   NewNum:=GetNumSoucheG(NewSouchePiece,LaNewDate) ;
TOBGenere.PutValue('GP_NUMERO',NewNum) ;
if GetInfoParPiece(NewNature,'GPP_ACTIONFINI')='ENR' then TOBGenere.PutValue('GP_VIVANTE','-') else TOBGenere.PutValue('GP_VIVANTE','X') ;
TOBGenere.PutValue('GP_DATECREATION',Date) ;
TOBGenere.PutValue('GP_CREATEUR',V_PGI.User);
TOBGenere.PutValue('GP_DATEMODIF',NowH) ;
TOBGenere.PutValue('GP_ETATVISA','NON') ;
TOBGenere.PutValue('GP_EDITEE','-') ;
//
// Modified by f.vautrain 24/10/2017 15:30:16 - FS#2753 - NCN : développement selon devis ML-171509-3A
If (ModeGenere = 'ETUTODEV') then
  TOBGenere.PutValue('GP_CREEPAR','AO')
else
  TOBGenere.PutValue('GP_CREEPAR','GEN') ;
  
TOBGenere.PutValue('GP_VENTEACHAT',VenteAchat) ;
TOBGenere.PutValue('GP_DATETAUXDEV',TOBPiece.GetValue('GP_DATETAUXDEV')) ;
TOBGenere.PutValue('GP_REFINTERNE',TOBPiece.GetValue('GP_REFINTERNE')) ;
TOBGenere.PutValue('GP_REFEXTERNE',TOBPiece.GetValue('GP_REFEXTERNE')) ;
TOBGenere.PutValue('GP_RIB',TOBPiece.GetValue('GP_RIB')) ;
// MODIF BTP
TOBGenere.PutValue('GP_ARRONDILIGNE',TOBPiece.GetValue('GP_ARRONDILIGNE')) ;
if VenteAchat = 'VEN' Then
begin
  TOBGenere.PutValue('GP_PIECEENPA',GetInfoParPiece(NewNature, 'GPP_FORCEENPA'));
end;
//----
AlimNewPieceBTP ( TOBGenere,TOBPiece);

TOBGenere.PutValue('GP_MODEREGLE',TOBPiece.GetValue('GP_MODEREGLE')) ; // Ajout TG 11/09/2001

{Autres informations}
TOBNomenclature.ClearDetail ; IndiceNomen:=1 ; IndiceRg := 1;
TOBSerie.ClearDetail ; IndiceSerie:=1 ;
{Analytique}
TOBAnaP.ClearDetail ; TOBAnaS.ClearDetail ;
// Modif BTP
{Ouvrages}
TOBOUvrage.clearDetail;
TOBBasesL.clearDetail;
TOBOuvragesP.clearDetail;
TOBPieceRG.clearDetail;
TOBBasesRG.clearDetail;
TOBVTECOLL.ClearDetail;
// ------

UG_LoadAnaPiece(TOBPiece,TOBAnaP,TOBAnaS) ;
TobLigneTarif.ClearDetail;
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

Procedure G_GenereCompta ;
Var OldEcr,OldStk : RMVT ;
BEGIN
FillChar(OldEcr,Sizeof(OldEcr),#0) ; FillChar(OldStk,Sizeof(OldStk),#0) ;
if VH_GC.GCIfDefCEGID then
  RechLibTiersCegid (VenteAchat,TobTiers,TOBGenere,TobAdresses);
if Not PassationComptable(TOBGenere,TOBOuvrage, TOBOuvragesP,TOBBases,TOBBasesL,TOBEches,TOBPieceTrait,nil,TOBTiers,TOBArticles,TOBCpta,TOBAcomptes,TOBPorcs,TOBPIECERG,TOBBasesRG,TOBAnaP,TOBAnaS,nil,TOBVTECOLL,DEV,OldEcr,OldStk,Not(bContinuOnError))
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
  TOBOuvragesP.SetAllModifie(True)  ;
  TOBEches.SetAllModifie(True)  ;
  TOBAcomptes.SetAllModifie(True)  ;
  TOBPorcs.SetAllModifie(True)  ;
  TOBSerie.SetAllModifie(True)  ;
  TOBTiers.SetAllModifie(True)  ;
  TOBAnaP.SetAllModifie(True)   ; TOBAnaS.SetAllModifie(True)   ;
  // Modif BTP
  TOBPieceRG.SetAllModifie (true);
  TOBBasesRG.SetAllModifie (true);
  TOBVTECOLL.SetAllModifie(true);
  TobLigneTarif.SetAllModifie (True);
  //Affaire-ONYX
  TOBRevision.SetAllModifie(True);
  TOBVariable.SetAllModifie(True);
END ;

Procedure G_AjouteEvent ;
Var St : String ;
BEGIN
St:='Pièce N° '+IntToStr(TOBGenere.GetValue('GP_NUMERO'))
   +', Tiers '+TOBGenere.GetValue('GP_TIERS')
   +', Total HT de '+StrfPoint(TOBGenere.GetValue('GP_TOTALHTDEV'))+' '+RechDom('TTDEVISETOUTES',TOBGenere.GetValue('GP_DEVISE'),False) ;
PiecesG.Add(St) ;
END ;

Procedure G_AjouteListePiece;
Var TobDet : TOB;
BEGIN
TobDet:=TOB.Create('',TOBListePiece,-1);
TobDet.AddChampSupValeur ('NATUREPIECE',TOBGenere.GetValue('GP_NATUREPIECEG'));
TobDet.AddChampSupValeur ('NUMERO',TOBGenere.GetValue('GP_NUMERO'));
TobDet.AddChampSupValeur ('SOUCHE',TOBGenere.GetValue('GP_SOUCHE'));
END;

Procedure G_TraiteEuro ;
Var leModeOppose : boolean ;
BEGIN
if Not ForceEuroGC then Exit ;
LeModeOppose:=(TOBGenere.GetValue('GP_SAISIECONTRE')='X') ; if Not LeModeOppose then Exit ;
if TOBGenere.GetValue('GP_DEVISE')<>V_PGI.DevisePivot then Exit ;
PutValueDetail(TOBGenere,'GP_SAISIECONTRE','-') ;
ForcePieceEuro(TOBGenere,TOBBases,TOBEches,TOBPorcs,TOBNomenclature,TOBAcomptes,TOBOuvrage,TOBPIeceRG,TOBBasesRG)  ;
END ;

// Modif MODE
procedure G_ValideLesLignesGeneriques ( TOBPiece: TOB; ElimineLignesZero: boolean = False );
var
  Ind           : Integer;
  iNoTypeDim    : Integer;
  iNoCodeDim    : Integer;
  iNoCodeArt    : Integer;
  iNoQteReste   : Integer;
  iNoQteReliq   : integer;
  iNoQteStock   : Integer;
  iNoQteFact    : integer;
  INoMtReste    : Integer;
  INoMtReliquat : Integer;
  INoMontantHT  : Integer;
  TOBL          : TOB;
  TOBGen        : TOB;
  TypeDim       : string;
  CodeDim       : string;
  Qte           : Double;
  QteReste      : Double;
  QteReliq      : Double;
  QteStock      : Double;
  QteFact       : double;
  MtReste       : Double;
  MtReliq       : Double;
  MontantHT     : Double;
  Montant       : Double;
  Renum         : boolean;  // MODIF MODE 30/10/2003
  Ok_ReliquatMt : Boolean;
  {****************************************************************************}
  procedure RAZGenerique (Maj: boolean);
  begin
    if (TOBGen <> nil) and (Maj) then
    begin
      TOBGen.PutValeur(iNoQteReste,   QteReste);
      TOBGen.PutValeur(iNoQteReliq,   QteReliq);
      TOBGen.PutValeur(iNoQteStock,   QteStock);
      TOBGen.PutValeur(iNoQteFact ,   QteFact);
      TOBGen.PutValeur(iNoMtReste,    MtReste);
      TOBGen.PutValeur(INoMtReliquat, MtReliq);
      TOBGen.PutValeur(iNoMontantHT,  MontantHT);

    end;
    CodeDim   := '';
    TOBGen    := nil;
    QteReste  := 0;
    QteReliq  := 0;
    QteStock  := 0;
    QteFact   := 0;
    MtReste   := 0;
    MtReliq   := 0;
    MontantHT := 0;
  end;
  {****************************************************************************}
begin
  iNoTypeDim    := 0;
  iNoCodeDim    := 0;
  iNoCodeArt    := 0;
  iNoQteReste   := 0;
  iNoQteReliq   := 0;
  iNoQteStock   := 0;
  iNoQteFact    := 0;
  iNoMtReste    := 0;
  INoMtReliquat := 0;
  INoMontantHT  := 0;

  RAZGenerique(False);

  for Ind := 0 to TOBPiece.Detail.Count -1 do
  begin
    TOBL := TOBPiece.Detail[Ind];
    Ok_ReliquatMt := CtrlOkReliquat(TOBL, 'GL');
    if Ind = 0 then
    begin
      iNoTypeDim    := TOBL.GetNumChamp('GL_TYPEDIM');
      iNoCodeDim    := TOBL.GetNumChamp('GL_CODESDIM');
      iNoCodeArt    := TOBL.GetNumChamp('GL_CODEARTICLE');
      iNoQteReste   := TOBL.GetNumChamp('GL_QTERESTE');
      iNoQteReliq   := TOBL.GetNumChamp('GL_QTERELIQUAT');
      iNoQteStock   := TOBL.GetNumChamp('GL_QTESTOCK');
      iNoQteFact    := TOBL.GetNumChamp('GL_QTEFACT');
      INoMtReste    := TOBL.GetNumChamp('GL_MTRESTE');
      INoMtReliquat := TOBL.GetNumChamp('GL_MTRELIQUAT');
      INoMontantHT  := TOBL.GetNumChamp('GL_MONTANTHTDEV');
    end;

    TypeDim := TOBL.GetValeur(iNoTypeDim);
    if TypeDim = 'GEN' then
    begin
      // cas d'une ligne générique
      RAZGenerique(True);
      TOBGen := TOBL;
      CodeDim := TOBGen.GetValeur(iNoCodeDim);
    end else
    if TypeDim = 'DIM' then
    begin
      // cas d'une ligne dimensionnée
      if (TOBGen <> nil) and (CodeDim = TOBL.GetValeur(iNoCodeArt)) then
      begin
        Qte := TOBL.GetValeur(iNoQteReste);
        QteReste := QteReste + Qte;
        Qte := TOBL.GetValeur(iNoQteReliq);
        QteReliq := QteReliq + Qte;
        Qte := TOBL.GetValeur(iNoQteStock);
        QteStock := QteStock + Qte;
        Qte := TOBL.GetValeur(iNoQteFact);
        QteFact := QteFact + Qte;
        // --- GUINIER ---
        if Ok_ReliquatMt then
        begin
          Montant := TOBL.GetValeur(iNoMtReste);
          MtReste := MtReste + Montant;
          Montant := TOBL.GetValeur(iNoQteReliq);
          MtReliq := MtReliq + Montant;
        end;
      end;
    end else
    begin
      // cas des autres lignes
      RAZGenerique(True);
    end;
  end;
  RAZGenerique(True);
  // MODIF MODE 30/10/2003
  if ElimineLignesZero then
  begin
    // suppression des lignes sans quantité
    Renum := False;
    for Ind := TOBPiece.Detail.Count -1 downto 0 do
    begin
      TOBL := TOBPiece.Detail[Ind];
      if ((not TOBL.FieldExists('LIGNECOMMENT')) or (TOBL.GetValue('LIGNECOMMENT') <> 'X')) and
         ((TOBL.GetValeur(iNoQteStock) = 0) and (TOBL.GetValeur(iNoQteReste) = 0)) then
      begin
        TOBL.Free;
        Renum := True;
      end;
    end;
    if Renum then NumeroteLignesGC(nil, TOBPiece);
  end;
  // FIN MODIF MODE 30/10/2003
end;

{$IFDEF BTP}
Procedure G_EnregistreLaPieceCumule(TOBPiece,TOBAffaire: TOB; GestionReliquat:boolean=false;FromLigneDetail : boolean=false);
Var Domaine,Nature,Lib : String ;
//    NumPiece : Integer;
    NeedVisaEnd : boolean ;
    MontantVisaEnd : Double ;
    ind : integer;
begin

  GestionPhase := TGestionPhase.create;
  UG_DetruitLaPiece (TOBGenere);

  if ctxAffaire in V_PGI.PGIContexte then
     BEGIN
     // en GI, on peut avoir des ports sans lignes sur les pièces associées aux affaires ...
     if (TOBGenere.Detail.Count<=0) and (TOBPorcs.Detail.Count<=0) then Exit ;
     END else
     BEGIN
     if TOBGenere.Detail.Count<=0 then Exit;
     END ;
  {$IFDEF AFFAIRE}
  InverseQteLigneArticle(NewNature, TobPiece, TOBGenere,TOBBases,TOBEches,TobPorcs); // pour passer d'une FAC en AVC
  {$ENDIF}

{Calculs et finitions}
	bDuplicPiece := true;  // Modif LS Pour Correction QTERESTE
  G_FinirPieceGenere ;
	bDuplicPiece := false; // Modif LS Pour Correction QTERESTE

// Reactive la pièce si morte
TOBGenere.PutValue('GP_VIVANTE','X');
// --
UG_MajLesComms(TOBGenere,TOBArticles,TOBComms) ;
ValideLaPeriode(TOBGenere) ;
{$IFDEF BTP}
//AjouteLignesVirtuellesOuv (TOBGenere,TOBOuvrage,TOBArticles,TOBCpta,TOBTiers,TOBCatalogu,nil,nil,DEV);
{$ENDIF}
PutValueDetail(TOBGenere,'GP_RECALCULER','X');
BeforeApplicFrais (TOBgenere,TOBOuvrage);
CalculeMontantsDoc (TOBgenere,TOBOuvrage,false);
ZeroFacture (TOBgenere);
for Ind := 0 to TOBgenere.detail.count -1 do ZeroLigneMontant (TOBgenere.detail[Ind]);
PutValueDetail (TOBgenere,'GP_RECALCULER','X');
TOBBases.ClearDetail;
TOBBasesL.ClearDetail;
ZeroMontantPorts (TOBPorcs);
ReinitMontantPieceTrait (TOBGenere,TOBaffaire,TOBPieceTrait);
CalculFacture(TOBAffaire,TOBGenere,TOBPieceTrait,TOBSSTrait,TOBOuvrage,TOBOuvragesP,TOBBases,TOBBasesL,TOBTiers,TOBArticles,TOBPorcs,TOBPieceRG,TOBBasesRG,nil,DEV) ;
{$IFDEF BTP}
//DetruitLignesVirtuellesOuv (TOBGenere,DEV);
{$ENDIF}
CalculeSousTotauxPiece(TOBGenere) ;
{Visa}
Domaine:=TOBGenere.GetValue('GP_DOMAINE') ;
Nature:=TOBGenere.GetValue('GP_NATUREPIECEG') ; Lib:='' ;

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
TOBEches.ClearDetail ; GereEcheancesGC(TOBGenere,TOBTiers,TOBEches,TOBAcomptes,TOBPieceRG,nil,TOBPorcs,taCreat,DEV,False) ;
G_InitToutModif ;
G_TraiteEuro ;
// Modif MODE
{Pour interface avec Orli} // plus besoin bicose ajout de ligne a piece deja créé.
//if LiaisonOrli then TOBGenere.PutValue('GP_REFINTERNE', IntToStr(TOBGenere.GetValue('GP_NUMERO')));
// ----------
// Correction OPTIMISATION
if V_PGI.IoError = oeOk then ValideLesArticlesFromOuv(TOBarticles,TOBOuvrage);
// --
if V_PGI.IoError=oeOk then ValideLesLignes(TOBGenere,TOBArticles,TobCatalogu,TOBNomenclature,TOBOuvrage,TOBPieceRG,TOBBasesRG,false,True,true,true,TOBPiece) ;
if V_PGI.IoError=oeOk then ValideLesLignesCompl(TOBGenere, nil);
//if V_PGI.IoError=oeOk then ValideLesAdresses(TOBGenere,TOBPiece,TOBAdresses) ;
if V_PGI.IoError=oeOk then ValideLesArticles(TOBGenere,TOBArticles) ;
if V_PGI.IoError=oeOk then ValideLesCatalogues(TOBGenere,TOBCatalogu) ;
if V_PGI.IoError=oeOk then ValideleTiers(TOBGenere,TOBTiers) ;
{$IFDEF BTP}
if V_PGI.IoError=oeOk then OuvrageDifferencie (TOBGenere,TOBOuvrage,True,DEV);
if V_PGI.IoError=oeOk then ValideLesOuv(TOBOuvrage,TOBgenere) ;
// NEW ONE (OPTIMISATION CALCUL)
//if V_PGI.IoError = oeOk then ValideLesOuvPlat (TOBOuvragesP, TOBgenere);
if V_PGI.IoError = oeOk then ValideLesBases(TOBgenere,TobBases,TOBBasesL);
//--
//PutValueDetail (TOBgenere,'GP_RECALCULER','X');
//AjouteLignesVirtuellesOuv (TOBGenere,TOBOuvrage,TOBArticles,TOBCpta,TOBTiers,TOBCatalogu,nil,nil,DEV);
//CalculFacture(TOBGenere,TOBOuvrage,TOBOuvragesP,TOBBases,TOBBasesL,TOBTiers,TOBArticles,TOBPorcs,TOBPieceRG,TOBBasesRG,DEV);
{$ENDIF}
if V_PGI.IoError=oeOk then UG_ValideAnals(TOBGenere,TOBAnaP,TOBAnaS) ;
if V_PGI.IoError=oeOk then G_GenereCompta ;
if (V_PGI.IOError = OeOk) and (TOBVTECOLL.detail.count >0) then
begin
  PrepareInsertCollectif(TOBGenere,TOBVTECOLL);
  TOBVTECOLL.InsertDB(nil);
end;
if V_PGI.IoError=oeOk then G_ValideLesSeries ;
if V_PGI.IoError=oeOk then ValideLesPorcs(TOBGenere,TOBPorcs) ;
{$IFDEF BTP}
//DetruitLignesVirtuellesOuv (TOBGenere,DEV);
{$ENDIF}
// NEW ONE (OPTIMISATION CALCUL)
//if V_PGI.IoError = oeOk then TOBOuvragesP.InsertDBByNivel(false);
if V_PGI.IoError = oeOk then TOBBasesL.InsertDB(nil);
//--
if V_PGI.IoError=oeOk then TOBGenere.InsertDBByNivel(False) ;
if V_PGI.IoError=oeOk then TOBBases.InsertDB(Nil) ;
if V_PGI.IoError=oeOk then TOBEches.InsertDB(Nil) ;
if V_PGI.IoError=oeOk then TOBAnaP.InsertDB(Nil) ;
if V_PGI.IoError=oeOk then TOBAnaS.InsertDB(Nil) ;
if V_PGI.IoError=oeOk then ValideLesNomen(TOBNomenclature) ;
{$IFDEF BTP}
if (V_PGI.IoError=oeOk) and (WithLiaisonInterDoc) then TOBLiaison.InsertDB (nil);
{$ENDIF}
// Modif BTP
if V_PGI.IoError=oeOk then ValideLesRetenues (TOBGenere,TOBPIECERG);
if V_PGI.IoError=oeOk then ValideLesBasesRG(TOBgenere,TOBBasesRG);
// ----

if V_PGI.IoError=oeOk then ValideLesSousTrait(TOBGenere,TOBSSTRait,DEV);
if V_PGI.IoError=oeOk then ValideLesPieceTrait(TOBGenere,TOBaffaire,TOBPieceTrait,TOBSSTRAIT,DEV);

{$IFDEF AFFAIRE}
if V_PGI.IoError=oeOk then if (ctxAffaire in V_PGI.PGIContexte) or (VH_GC.GASeria) then MajGestionAffaire(TOBGenere,TOBPiece,TOBPieceRG,TOBBasesRG,TOBAcomptes,'VAL',DEV,FinTravaux) ;
if V_PGI.IoError=oeOk then ValideActivite(TOBGenere,Nil,TOBArticles, GereActivite,False,DelActivite,False);
  //Affaire-ONYX
if (V_PGI.IoError = oeOk) and (TobVariable <> nil) then ValideLesVariables(TOBGenere, TOBPiece, TOBVariable);
if (V_PGI.IoError = oeOk) and (TOBRevision <> nil) then ValideLesRevisions(TOBGenere, TOBPiece, TOBRevision);
{$ENDIF}
// MODIF BTP -- LS
{$IFDEF BTP}
if (V_PGI.IoError=oeOk) and (FromLigneDetail) then UG_ReajustePiecesPrecedente (TOBGenere,TOBPiece,TOBArticles);
{$ENDIF}
// Modif BRL 180609 : Ajout maj phases
if (V_PGI.IOError = OeOk) Then GestionPhase.GenerelesPhases(TOBGenere,nil,False,False,False,TaCreat);
if (V_PGI.IOError = OeOk) Then GestionPhase.free;
// --
{Compte rendu}
if V_PGI.IoError=oeOk then
   BEGIN
   if PremNum=-1 then PremNum:=TOBGenere.GetValue('GP_NUMERO') ;
   LastNum:=TOBGenere.GetValue('GP_NUMERO') ;
   Inc(Nbp) ; G_AjouteEvent ; G_AjouteListePiece;
   END ;
end;
{$ENDIF}

procedure GetValorisation (TOBPiece : TOB; var TotalPa,TotalPr,TotalPV : double);
VAR CLEDOC : r_cledoc;
		QQ : TQuery;
begin
	TotalPa := 0;
	TotalPR := 0;
  TotalPV := 0;
  DecodeRefPiece(TOBpiece.getValue('PIECEORIGINE'),cledoc);
  QQ := OpenSql ('SELECT GP_MONTANTPA,GP_MONTANTPR,GP_TOTALHTDEV,GP_TOTALREMISEDEV,GP_TOTALESCDEV FROM PIECE WHERE '+
  								WherePiece(Cledoc,ttdpiece,false),True,-1,'',True);
  if not QQ.eof then
  begin
  	TotalPA := QQ.findFIeld('GP_MONTANTPA').AsFloat;
  	TotalPR := QQ.findFIeld('GP_MONTANTPR').AsFloat;
    TotalPv := QQ.findFIeld('GP_TOTALHTDEV').AsFloat + QQ.findFIeld('GP_TOTALREMISEDEV').AsFloat + QQ.findFIeld('GP_TOTALESCDEV').AsFloat;
  end;
  ferme (QQ);
end;

function GenereEcartArticle (TOBPiece : TOB; var TotalPa,TotalPr,TotalPV : double) : boolean;
var TOBL,TOBArt,TOBAffaire : TOB;
		Numl : integer;
    Artecart : string;
    Qart : TQuery;
    EnHt,Created : boolean;
    TOBtaxesL : TOB;
    EcartPv,EcartPa,EcartPR : double;
begin
  Result := false;
  ecartPA := Arrondi(TotalPa - TOBgenere.getValue('GP_MONTANTPA'),V_PGI.okdecP);
  ecartPR := Arrondi(TotalPr - TOBgenere.getValue('GP_MONTANTPR'),V_PGI.okdecP);
  ecartPV := Arrondi( TotalPV - (TOBgenere.getValue('GP_TOTALHTDEV') +
             TOBgenere.getValue('GP_TOTALREMISEDEV') +
             TOBgenere.getValue('GP_TOTALESCDEV')),V_Pgi.okdecP);
  if (ecartPa=0) and (ecartPr=0) and (ecartPv=0) then exit;
	created := false;
	TOBart := TOB.create ('ARTICLE',nil,-1);
  TOBTaxesL := TOB.Create ('LES TAXES LIGNES',nil,-1);
  TRY
    ArtEcart := GetParamsoc('SO_BTECARTPMA');
    if ArtEcart = '' then
    begin
      PgiError (Traduirememoire('Impossible d''affecter la différence : l''article d''écart n''a pas été paramétré'));
    end;
    Qart := opensql('Select GA_ARTICLE from ARTICLE Where GA_CODEARTICLE="' + ArtEcart + '"', true,-1, '', True);
    if not Qart.eof then
    begin
      ArtEcart := Qart.fields[0].AsString;
      ferme (Qart);
    end else
    begin
      ferme (Qart);
      PgiError (Traduirememoire('Impossible d''affecter la différence : l''article d''écart n''a pas été paramétré'));
    end;
    QArt := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                   'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                   'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+ArtEcart+'"',true,-1, '', True);
    if not QArt.EOF then
    begin
      TOBArt.SelectDB('', QArt);
      InitChampsSupArticle (TOBART);
      ferme (Qart);
    end else
    begin
      ferme (Qart);
      PgiError (Traduirememoire('Impossible d''affecter la différence : l''article d''écart n''a pas été paramétré'));
    end;
    //
    TOBaffaire := FindCetteAffaire (TOBPiece.getValue('GP_AFFAIRE'));
    if TOBAffaire = nil then
    begin
    	StockeCetteAffaire (TOBPiece.getValue('GP_AFFAIRE'));
    	TOBaffaire := FindCetteAffaire (TOBPiece.getValue('GP_AFFAIRE'));
      created := true;
    end;
    TOBL := TOB.Create ('LIGNE',TOBPiece, -1); // ajout d'une ligne
    numl := TOBPiece.detail.count;
    NewTOBLigneFille(TOBL);
    AddLesSupLigne(TOBL,false);
    InitLesSupLigne(TOBL);
    InitialiseTOBLigne(TOBPiece, TOBTiers, TOBAffaire, NumL);
    TOBL.putValue('GL_CREERPAR','GEN'); // Définition pour écart
    EnHt := (TOBL.getValue('GL_FACTUREHT')='X');
    // -- Pour ne pas recalculer a partir des PA * coef ..etc..
    TOBL.putValue('GLC_NONAPPLICFRAIS','X');
    TOBL.putValue('GLC_NONAPPLICFC','X');
    TOBL.putValue('GLC_NONAPPLICFG','-');
    //
    TOBArt.PutValue('REFARTSAISIE', copy(TOBArt.GetValue('GA_ARTICLE'), 1, 18));
    TOBL.PutValue('GL_ARTICLE', TOBArt.GetValue('GA_ARTICLE'));
    TOBL.PutValue('GL_REFARTSAISIE', TOBArt.GetValue('REFARTSAISIE'));
    TOBL.PutValue('GL_CODEARTICLE', TOBArt.GetValue('GA_CODEARTICLE'));
    TOBL.PutValue('GL_REFARTBARRE', TOBArt.GetValue('REFARTBARRE'));
    TOBL.PutValue('GL_REFARTTIERS', TOBArt.GetValue('REFARTTIERS'));
    TOBL.PutValue('GL_TYPEREF', 'ART');
    ArticleVersLigne(TOBPiece, TOBArt, nil, TOBL, TOBTiers);
    TOBL.PutValue('GL_QTEFACT', 1);
    TOBL.PutValue('GL_QTESTOCK', 1);
    TOBL.PutValue('GL_QTERESTE', 1);
    // --- GUINIER ---
    TOBL.PutValue('GL_MTRESTE',  0);

    TOBL.PutValue('GL_PRIXPOURQTE', 1);
    TOBL.putValue('GL_COEFFR',0);
    TOBL.putValue('GL_COEFFC',0);
    TOBL.PutValue('GL_DPA', EcartPA);
    TOBL.PutValue('GL_DPR', EcartPR);
    if EcartPa <> 0 then TOBL.PutValue('GL_COEFFG',Arrondi(EcartPR/EcartPa,4));
    TOBL.PutValue('GL_PUHTDEV', EcartPv);
    TOBL.putValue('GL_COEFMARG',0);
    TOBL.PutValue('GL_RECALCULER', 'X');
    Result := True;
  FINALLY
  	TOBARt.free;
    TOBtaxesL.free;
    if created then
    begin
    	TOBaffaire.free;
    end;
  END;
end;

procedure Retablibases (TOBgenere,TOBBases : TOB; DEV : RDevise);
var II : integer;
		TOBB : TOB;
    Indmax : integer;
    Max,DiffVal,MTTVAC,MTHT,DeltaHt : Double;
begin
  Max     := 0;
  Indmax  := -1;
  MtTvaC  := 0;
  MTHT    := 0;
  for II := 0 to TOBBases.Detail.Count - 1 do
  begin
    TOBB := TOBBases.Detail[II];
    if Abs(TOBB.GetDouble('GPB_VALEURDEV')) > Max then
    begin
    	MAx := Abs(TOBB.GetDouble('GPB_VALEURDEV'));
      //
      MTHT := MTHT + TOBB.GetDouble('GPB_BASEDEV');
      MTTVAC := MTTVAC + TOBB.GetDouble('GPB_VALEURDEV');
      //
      Indmax := II;
    end;
  end;
  DiffVal := (TOBGenere.GetDouble('GP_TOTALTTCDEV') - TOBGenere.GetDouble('GP_TOTALHTDEV')) - MTTVAC ;
  DeltaHt := TOBGenere.GetDouble('GP_TOTALHTDEV') -  MTHT ;
  if Indmax >= 0 then
  begin
    TOBB := TOBBases.detail[Indmax];
    TOBB.SetDouble('GPB_BASEDEV',TOBB.GetDouble('GPB_BASEDEV')+DeltaHt);
    TOBB.SetDouble('GPB_VALEURDEV',TOBB.GetDouble('GPB_VALEURDEV')+DiffVal);
    //
    TOBB.SetDouble('GPB_BASETAXE',DEVISETOPIVOT(TOBB.GetDouble('GPB_BASEDEV'),DEV.Taux,DEV.Quotite));
    TOBB.SetDouble('GPB_VALEURTAXE',DEVISETOPIVOT(TOBB.GetDouble('GPB_VALEURDEV'),DEV.Taux,DEV.Quotite));
  end;
end;

Procedure G_GenereLaPiece ( TOBPiece,TOBAffaire,TOBPieceTrait : TOB ;RendActive : boolean=false; GestionReliquat : boolean = false; FromLigneDetail : boolean=false; ModeGeneration : string='') ;
Var Domaine,Nature,Lib : String ;
    NumPiece : Integer;
    NeedVisaEnd : boolean ;
    MontantVisaEnd : Double ;
    AutoriseEclatOuvrage : boolean;
    Ind,I : integer;
    TotalPa,TotalPr,TotalPV : double;
    GestionConso : TGestionPhase;
    CodeMessage : string;
BEGIN
	AutoriseEclatOuvrage := false;
  if ModeGeneration = 'DEVTOCTE' then AutoriseEclatOuvrage := true;
  if (ModeGeneration = 'FACDETAILCONSO') or (ModeGeneration = 'FACCONSO') then autoriseEclatOuvrage := true;
  if (Modegeneration = '') and (pos (NewNature ,'FAC;FPR')>0) then autoriseEclatOuvrage := true;
  if ctxAffaire in V_PGI.PGIContexte then
     BEGIN
     // en GI, on peut avoir des ports sans lignes sur les pièces associées aux affaires ...
     if (TOBGenere.Detail.Count<=0) and (TOBPorcs.Detail.Count<=0) then Exit ;
     END else
     BEGIN
     if TOBGenere.Detail.Count<=0 then Exit;
     END ;
  {$IFDEF AFFAIRE}
  RepriseQuantiteVar(NewNature, TobPiece, TOBGenere,TobVariable);  // recup des qté non arrondi
  {$ENDIF}

  {$IFDEF BTP}
	TheRepartTva.TOBPiece := TOBgenere;
  TheRepartTva.TOBOuvrages := TOBOuvrage;
  TheRepartTva.TOBBases := TOBBases;
  TheRepartTva.AppliqueFromPiece;
  {$ENDIF}
  {Calculs et finitions}
  G_FinirPieceGenere (ModeGeneration);
  UG_MajLesComms(TOBGenere,TOBArticles,TOBComms) ;
  ValideLaPeriode(TOBGenere) ;
  {$IFDEF BTP}
  //AjouteLignesVirtuellesOuv (TOBGenere,TOBOuvrage,TOBArticles,TOBCpta,TOBTiers,TOBCatalogu,nil,nil,DEV);
  {$ENDIF}
  PutValueDetail(TOBGenere,'GP_RECALCULER','X');
  BeforeApplicFrais (TOBgenere,TOBOuvrage);
  CalculeMontantsDoc (TOBgenere,TOBOuvrage,false);
  (*
  if TOBGenere.GetValue('GP_NATUREPIECEG')='DBT' then // Cas de la génération de devis depuis une etude
  begin
	PieceTraitUsable := true;
  end;
  *)
  {$IFDEF BTP}
  if ModeGeneration = 'DEVTOCTE' then
  begin
		if (TOBGenere.FieldExists('PIECEORIGINE') and (TOBGenere.GetString('PIECEORIGINE')<>'')) then
    begin
      GetValorisation (TOBgenere,TotalPa,TotalPR,TotalPV);
      if (TOBgenere.getValue('GP_MONTANTPA')<>TotalPa) or
         (TOBgenere.getValue('GP_MONTANTPR')<>TotalPr) or
         (TOBgenere.getValue('GP_TOTALHTDEV')+TOBgenere.getValue('GP_TOTALREMISEDEV')+TOBgenere.getValue('GP_TOTALESCDEV')<>TotalPV) then
      begin
        ZeroFacture (TOBgenere);
        for Ind := 0 to TOBgenere.detail.count -1 do ZeroLigneMontant (TOBgenere.detail[Ind]);
        PutValueDetail (TOBgenere,'GP_RECALCULER','X');
        TOBBases.ClearDetail;
        TOBBasesL.ClearDetail;
        ZeroMontantPorts (TOBPorcs);
				ReinitMontantPieceTrait (TOBGenere,TOBaffaire,TOBPieceTrait);
        TOBVTECOLL.ClearDetail;
        CalculFacture(TOBAffaire,TOBGenere,TOBpieceTrait,TOBSSTRAIT,TOBOuvrage,TOBOuvragesP,TOBBases,TOBBasesL,TOBTiers,TOBArticles,TOBPorcs,TOBPieceRG,TOBBasesRG,TOBVTECOLL,DEV) ;
        if GenereEcartArticle (TOBgenere,TotalPa,TotalPr,TotalPV) then
        begin
          NumeroteLignesGC(Nil,TOBGenere, true);
          ZeroLigneMontant (TOBgenere.detail[TOBGenere.detail.count -1]);
          TOBgenere.detail[TOBGenere.detail.count -1].putvalue('GL_RECALCULER','X');
          PutValueDetail (TOBgenere,'GP_RECALCULER','X');
          TOBVTECOLL.ClearDetail;
          CalculFacture(TOBAffaire,TOBGenere,TOBpieceTrait,TOBSSTRAIT,TOBOuvrage,TOBOuvragesP,TOBBases,TOBBasesL,TOBTiers,TOBArticles,TOBPorcs,TOBPieceRG,TOBBasesRG, TOBVTECOLL, DEV) ;
        end;
      end;
    end;
  end else
  begin
    if ModeGeneration = 'ETUTODEV' then
    begin
      TOBgenere.setBoolean ('AFF_OKSIZERO',GetParamSocSecur('SO_OKSIZERO',False));
    end;
    ZeroFacture (TOBgenere);
    for Ind := 0 to TOBgenere.detail.count -1 do ZeroLigneMontant (TOBgenere.detail[Ind]);
    PutValueDetail (TOBgenere,'GP_RECALCULER','X');
    TOBBases.ClearDetail;
    TOBBasesL.ClearDetail;
    ZeroMontantPorts (TOBPorcs);
		ReinitMontantPieceTrait (TOBGenere,TOBaffaire,TOBPieceTrait);
    TOBVTECOLL.ClearDetail;
		CalculFacture(TOBAffaire,TOBGenere,TOBPieceTrait,TOBSSTRAIT,TOBOuvrage,TOBOuvragesP,TOBBases,TOBBasesL,TOBTiers,TOBArticles,TOBPorcs,TOBPieceRG,TOBBasesRG,TOBVTECOLL,DEV) ;
  end;
  //
  {$ENDIF}
  CalculeSousTotauxPiece(TOBGenere) ;
  Domaine:=TOBGenere.GetValue('GP_DOMAINE') ; Nature:=TOBGenere.GetValue('GP_NATUREPIECEG') ; Lib:='' ;

  {verification que le montant d'une piece  passant en compta soit differente de 0}
  if (Pos(Nature,PieceFacturation) > 0) and (TOBGenere.GetValue('GP_TOTALHT') = 0 ) then
  begin
	MessageAlerte('ATTENTION : Le montant de cette pièce est nul et donc ne sera pas générée') ;
	V_PGI.IoError:=oeUnknown;
  exit;
  end;
  {Visa}
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
  TOBEches.ClearDetail;
  GereEcheancesGC(TOBGenere,TOBTiers,TOBEches,TOBAcomptes,TOBPieceRG,TOBPieceTrait,TOBPorcs,taCreat,DEV,False) ;
  G_InitToutModif ;
  G_TraiteEuro ;

  {Enregistrement physique}
  // Modif BTP
  if ForceNumero then
   BEGIN
   NumPiece := TobGenere.getValue('GP_NUMERO');
   if NumPiece = 0 then
      V_PGI.IoError := oeSaisie
   else
      begin
      if Not SetDefinitiveNumber(TobGenere, TobBases, TOBBasesL, TobEches, TobNomenclature, TobAcomptes, TobPieceRG, TobBasesRG, TobLigneTarif, NumPiece) then V_PGI.IoError:=oeSaisie ;
      {$IFDEF AFFAIRE}
      if not SetDefinitiveNumberAffaire(TOBGenere, TobRevision, TobVariable, NumPiece) then V_PGI.IoError := oeSaisie;
      {$ENDIF}
      end
  end else
  //
   if EclateParPiece then
      BEGIN
      NumPiece := TobGenere.GetValue('GP_NUMERO'); // G_GetNumPieceTiers (TobGenere);   PCS 12/03/2003
      TobGenere.PutValue('GP_NUMERO',0);
      if NumPiece = 0 then
         V_PGI.IoError:=oeSaisie
      else
         begin
         if Not SetDefinitiveNumber(TobGenere, TobBases, TOBBasesL,TobEches, TobNomenclature, TobAcomptes, TobPieceRG, TobBasesRG, TobLigneTarif,NumPiece) then V_PGI.IoError:=oeSaisie ;
         {$IFDEF AFFAIRE}
         if not SetDefinitiveNumberAffaire(TOBGenere, TobRevision, TobVariable, NumPiece) then V_PGI.IoError := oeSaisie;
         {$ENDIF}
         end
  end else
      BEGIN
      if Not SetNumeroDefinitif(TobGenere, TobBases, TOBBasesL,TobEches, TobNomenclature, TobAcomptes, TobPieceRG, TobBasesRG, TobLigneTarif) then V_PGI.IoError:=oeSaisie ;
      END;


  // Modif MODE
  {Pour interface avec Orli}
  if LiaisonOrli then TOBGenere.PutValue('GP_REFINTERNE', IntToStr(TOBGenere.GetValue('GP_NUMERO')));
  // ----------
  // NA - FQ MODE 10586
  // MODIF MODE 30/10/2003
  if CodeAff <> '' then G_ValideLesLignesGeneriques(TOBGenere, True);

  {$IFDEF AFFAIRE} // gm 14/10/03
  InverseQteLigneArticle(NewNature, TobPiece, TOBGenere,TOBBases,TOBEches,TobPorcs); // pour passer d'une FAC en AVC
  {$ENDIF}
  // Correction OPTIMISATION
  if V_PGI.IoError = oeOk then ValideLesArticlesFromOuv(TOBarticles,TOBOuvrage);
  // --
  if V_PGI.IoError=oeOk then ValideLesLignes(TOBGenere,TOBArticles,TobCatalogu,TOBNomenclature,TOBOuvrage,TOBPieceRG,TOBBasesRG,RendActive,True,false,true,TOBPiece,AutoriseEclatOuvrage) ;

  if TOBGenere.getValue('GP_RECALCULER')='X' then
  begin
  ZeroFacture (TOBgenere);
  for Ind := 0 to TOBgenere.detail.count -1 do ZeroLigneMontant (TOBgenere.detail[Ind]);
  PutValueDetail (TOBgenere,'GP_RECALCULER','X');
  TOBBases.ClearDetail;
  TOBBasesL.ClearDetail;
  ZeroMontantPorts (TOBPorcs);
	ReinitMontantPieceTrait (TOBGenere,TOBaffaire,TOBPieceTrait);
  TOBVTECOLL.ClearDetail;
  CalculFacture(TOBAffaire,TOBGenere,TOBPieceTrait,TOBSSTRAIT,TOBOuvrage,TOBOuvragesP,TOBBases,TOBBasesL,TOBTiers,TOBArticles,TOBPorcs,TOBPieceRG,TOBBasesRG,TOBVTECOLL,DEV) ;
	CalculeSousTotauxPiece(TOBGenere) ;
  end;
  {$IFDEF BTP}
  // Gestion Controle facture achat
  if (CleCFA.NumfacFou <> '') then
  begin
    TOBGenere.SetString('GP_REFEXTERNE',CleCFA.NumfacFou);
    if (CleCFA.MontantTTC <> 0) and (
    (Abs(TOBGenere.GetDouble ('GP_TOTALHTDEV'))<>Abs(CleCFA.MontantHT)) OR
    (Abs(TOBGenere.GetDouble ('GP_TOTALTTCDEV'))<>Abs(CleCFA.MontantTTC))) then
    begin
      //
      //
      if GetInfoParPiece (TOBGenere.GetString('GP_NATUREPIECEG'), 'GPP_ESTAVOIR') = 'X' then
      begin
        TOBGenere.PutValue ('GP_TOTALTTC',  DEVISETOPIVOT(CleCFA.MontantTTC,DEV.Quotite,DEV.Quotite)*  -1);
        TOBGenere.PutValue ('GP_TOTALTTCDEV', CleCFA.MontantTTC* -1);
        TOBGenere.PutValue ('GP_TOTALHT',  DEVISETOPIVOT(CleCFA.MontantHT,DEV.Quotite,DEV.Quotite)*  -1);
        TOBGenere.PutValue ('GP_TOTALHTDEV', CleCFA.MontantHT* -1);
      end else
      begin
        TOBGenere.PutValue ('GP_TOTALTTC', DEVISETOPIVOT(CleCFA.MontantTTC,DEV.Quotite,DEV.Quotite));
        TOBGenere.PutValue ('GP_TOTALTTCDEV', CleCFA.MontantTTC);
        TOBGenere.PutValue ('GP_TOTALHT', DEVISETOPIVOT(CleCFA.MontantHT,DEV.Quotite,DEV.Quotite));
        TOBGenere.PutValue ('GP_TOTALHTDEV', CleCFA.MontantHT);
      end;
      // HOP HOP HOP et maintenant on egalise les bases de TVA (pour eviter les messages a la c..)
      if TOBBases.detail.count > 0 then
      begin
        Retablibases (TOBgenere,TOBBases,DEV);
      end;
      // ----
      VH_GC.ModeGestionEcartComptable := 'CPA'; {DBR CPA}
      TOBEches.ClearDetail;
      GereEcheancesGC(TOBGenere,TOBTiers,TOBEches,TOBAcomptes,TOBPieceRG,TOBPieceTrait,TOBPorcs,taCreat,DEV,False) ;
    end;
  end;
  {$ENDIF}
  if V_PGI.IoError=oeOk then ValideLesLignesCompl(TOBGenere, nil);
  if V_PGI.IoError=oeOk then ValideLesAdresses(TOBGenere,TOBPiece,TOBAdresses) ;
  if V_PGI.IoError=oeOk then ValideLesArticles(TOBGenere,TOBArticles) ;
  if V_PGI.IoError=oeOk then ValideLesCatalogues(TOBGenere,TOBCatalogu) ;
  if V_PGI.IoError=oeOk then ValideleTiers(TOBGenere,TOBTiers) ;
  {$IFDEF BTP}
  //PutValueDetail (TOBgenere,'GP_RECALCULER','X');
  //AjouteLignesVirtuellesOuv (TOBGenere,TOBOuvrage,TOBArticles,TOBCpta,TOBTiers,TOBCatalogu,nil,nil,DEV,true,TOBAnap);
  //CalculFacture(TOBGenere,TOBOuvrage,TOBOUvragesP,TOBBases,TOBBasesL,TOBTiers,TOBArticles,TOBPorcs,TOBPieceRG,TOBBasesRG,DEV);
  {$ENDIF}
  if V_PGI.IoError=oeOk then UG_ValideAnals(TOBGenere,TOBAnaP,TOBAnaS) ;
  if V_PGI.IoError=oeOk then G_GenereCompta ;
  if V_PGI.IoError=oeOk then G_ValideLesSeries ;
  if V_PGI.IoError=oeOk then ValideLesAcomptes(TOBGenere,TOBAcomptes,TOBAcomptesGen_O) ;
  if V_PGI.IoError=oeOk then ValideLesPorcs(TOBGenere,TOBPorcs) ;
  {$IFDEF BTP}
  if V_PGI.IoError=oeOk then OuvrageDifferencie (TOBGenere,TOBOuvrage,True,DEV);
  if V_PGI.IoError=oeOk then ValideLesOuv(TOBOuvrage,TOBgenere) ;
  // NEW ONE (OPTIMISATION CALCUL)
  if (V_PGI.IoError = oeOk) and (Pos(TOBgenere.GetValue('GP_NATUREPIECEG'),'FBT;ABT;ABP;FBP;BAC')>0) then ValideLesOuvPlat (TOBOuvragesP, TOBgenere);
  if V_PGI.IoError = oeOk then ValideLesBases(TOBgenere,TobBases,TOBBasesL);
  //DetruitLignesVirtuellesOuv (TOBGenere,DEV);
  {$ENDIF}
  for I := 0 to TOBGenere.detail.count -1 do
  begin
    TRY
      ProtectionCoef (TOBGenere.detail[I]);
//      if V_PGI.IoError=oeOk then TOBGenere.detail[I].InsertOrUpdateDB(false) ;
    EXCEPT
      on E: Exception do
      begin
        PGIError(E.message + ' - Ligne '+InttoStr(I));
        V_PGI.IOError := OeUnknown;
        Exit;
      end;
    END;
  end;
  //
  TRY
    if V_PGI.IoError=oeOk then TOBGenere.InsertOrUpdateDB(false) ;
  EXCEPT
    on E: Exception do
    begin
      PGIError(E.message);
      V_PGI.IOError := OeUnknown;
      Exit;
    end;
  END;
  // NEW ONE (OPTIMISATION CALCUL)

  if V_PGI.IoError = oeOk then EcritlesOuvragesP (TOBgenere.getValue('GP_NATUREPIECEG'),TOBOuvragesP);
  if V_PGI.IoError = oeOk then TOBBasesL.InsertDB(nil);
  //--
  if V_PGI.IoError=oeOk then TOBBases.InsertDB(Nil) ;
  //if V_PGI.IoError=oeOk then TOBEches.InsertDB(Nil) ;
  if V_PGI.IoError=oeOk then TOBEches.InsertOrUpdateDB(True) ;

  if V_PGI.IoError=oeOk then TOBAnaP.InsertDB(Nil) ;
  if V_PGI.IoError=oeOk then TOBAnaS.InsertDB(Nil) ;
  if V_PGI.IoError=oeOk then ValideLesNomen(TOBNomenclature) ;
  {$IFDEF BTP}
  // --
  if (V_PGI.IoError=oeOk) and (WithLiaisonInterDoc) then TOBLiaison.InsertDB (nil);
  {$ENDIF}
  // Modif BTP
  //FV1 - gestion des cotraitants
  if V_PGI.IoError=oeOk then ValideLesSousTrait(TOBGenere,TOBSSTRait,DEV);
  if V_PGI.IoError=oeOk then ValideLesPieceTrait(TOBGenere,TOBaffaire,TOBPieceTrait,TOBSSTRAIT,DEV);
  //
  if V_PGI.IoError=oeOk then ValideLesRetenues (TOBGenere,TOBPIECERG);
  if V_PGI.IoError=oeOk then ReajusteCautionDocOrigine (TOBGenere,TOBPIECERG,true);
  if V_PGI.IoError=oeOk then ValideLesBasesRG(TOBgenere,TOBBasesRG);
  // ----
  {$IFDEF BTP}
  if (V_PGI.IoError=oeOk) and (TOBGenere.GetValue('GP_NATUREPIECEG')=GetParamSoc ('SO_BTNATBESOINCHA')) then ReajusteQteReste(TOBGenere) ;
  {$ENDIF}
  {$IFDEF AFFAIRE}
  if V_PGI.IoError=oeOk then if (ctxAffaire in V_PGI.PGIContexte) or (VH_GC.GASeria) then MajGestionAffaire(TOBGenere,TOBPiece,TOBPieceRG,TOBBasesRG,TOBAcomptes,'VAL',DEV,FinTravaux) ;
  if V_PGI.IoError=oeOk then ValideActivite(TOBGenere,Nil,TOBArticles, GereActivite,False,DelActivite,False);
  //Affaire-ONYX
  if (V_PGI.IoError = oeOk) and (TobVariable <> nil)  and  (GetParamSoc('SO_AFVARIABLES')) then
  ValideLesVariables(TOBGenere, TOBPiece, TOBVariable);
  if (V_PGI.IoError = oeOk) and (TOBRevision <> nil)  and  (GetParamSoc('SO_AFREVISIONPRIX'))then
  ValideLesRevisions(TOBGenere, TOBPiece, TOBRevision);
  {$ENDIF}
  {$IFDEF BTP}
  if (V_PGI.IoError=oeOk) then MajAffaireApresGeneration (TobGenere,TOBPieceRG,TOBBasesRG,TOBAcomptes,DEV,FinTravaux);
  if (V_PGI.IoError=oeOk) and
	((TOBGenere.GetValue('GP_NATUREPIECEG')=VH_GC.AFNatAffaire) or (TOBGenere.GetValue('GP_NATUREPIECEG')='BCE')) then
	begin
  	MajSousAffaire (TOBGenere,nil,'00',TaCreat,false,false,ModeGeneration);
	end;
  // MODIF BTP -- LS
  if (V_PGI.IoError=oeOk) and (FromLigneDetail) and ((ModeGeneration = 'CBTTOLIVC') or (ModeGeneration = 'RFOTOLIVC') or (ModeGeneration = 'CCTOLIVCLI')) then
  begin
	UG_reajusteQtePiecePrecedent (TOBGenere,TOBPiece);
  end;
  if (V_PGI.IoError=oeOk) and (FromLigneDetail) then UG_ReajustePiecesPrecedente (TOBGenere,TOBPiece,TOBArticles);
  if (V_PGI.IoError=oeOk) then TheRepartTva.Ecrit;

  if (ModeGeneration ='GENERECF') or (ModeGeneration ='DEVTOCTE') then
  begin
  if (V_PGI.IOError = OeOk) Then GestionConso := TGestionPhase.create;
  if (V_PGI.IOError = OeOk) Then GestionConso.GenerelesPhases(TOBGenere,nil,true,false,false,taCreat);
  if (V_PGI.IOError = OeOk) Then GestionConso.clear;
  if (V_PGI.IOError = OeOk) Then GestionConso.free;
  end;
  // --
  {$ENDIF}
  if V_PGI.IoError=oeOk then TOBLigneTarif.InsertDB(Nil);
  {Compte rendu}
  if V_PGI.IoError=oeOk then
   BEGIN
   if PremNum=-1 then PremNum:=TOBGenere.GetValue('GP_NUMERO') ;
   LastNum:=TOBGenere.GetValue('GP_NUMERO') ;
   Inc(Nbp) ; G_AjouteEvent ; G_AjouteListePiece;
   END ;

  {$IFDEF GPAOLIGHT}
  if V_PGI.IoError=oeOk then
  begin
    if Pos(Nature + ';', GetParamSoc('SO_WMISEENPROD')) <> 0 then wReceptionneWOLFromGL(TobGenere);
  end;
  {$ENDIF GPAOLIGHT}
  {$IFDEF EDI}
  if V_PGI.IoError = oeOk then
    if StrToBool_(GetInfoParPiece(TOBGenere.GetValue('GP_NATUREPIECEG'), 'GPP_PIECEEDI')) then
  if IsEDITiers(TOBGenere) then EDICreateETR(TOBGenere, EDIGetFieldFromETS('ETS_CODEMESSAGE', EDIGetCleETS(TOBGenere.GetValue('GP_TIERS'), TOBGenere.GetValue('GP_NATUREPIECEG'))));
  {$ENDIF EDI}

  if V_PGI.IoError = oeOk then ReaffecteDispoDiff(TOBArticles) ;
  // frais de chantier associés
  if (V_PGI.IoError=oeOk) then G_genereFraisChantier;
  if (V_PGI.IoError=oeOk) then
  begin
  if (GetparamSocSecur ('SO_OPTANALSTOCK',false)) and (IsLivraisonClient(TOBGenere)) then
  begin
    UpdateStatusMoisOD (TOBGenere);
  end;
  end;

  //FV1 : 01/12/2015 - FS#1811 - TEAM RESEAUX : interventions - ajout option clôturé si facturé selon paramètre société..
  if (ModeGeneration = 'FACDETAILCONSO') or (ModeGeneration = 'FACCONSO') then
  begin
  if V_PGI.IoError=oeOk then
  Begin
  if GetParamSocSecur('SO_FACTCLOTURE', false) then
    PositionneEtatAffaire(TobPiece.GetValue('GP_AFFAIRE'), 'CLO')
  else
    PositionneEtatAffaire(TobPiece.GetValue('GP_AFFAIRE'), 'FAC');
  end;
  end;

  if V_PGI.IoError=oeOk then
  Begin
    if GetParamSocSecur('SO_BTECHGACTIF',false) then
    begin
      CodeMessage := GetInfoParPiece(TOBGenere.GetValue('GP_NATUREPIECEG'),'GPP_ECHGMESSOUT');
      if CodeMessage <> '' then
      begin
        if not GenereMessageIS (CodeMessage,TOBMessages,TOBgenere,TOBBases,TOBTiers) then V_PGI.ioerror := oeUnknown;
      end;
    end;
  end;
END ;

// Modif MODE
function CreerReliquat(TOBPiece, TOBPiece_O : TOB): boolean;
Var
//  Reste: Double;
  NewEtat,RefSuiv: String;
  i: Integer;
  TOBL,TOBLOld: TOB;
  CleDoc, CleDoc_O: R_CleDoc;
  Ok_ReliquatMT : Boolean;
begin
//Result:=False ;
CleDoc_O := TOB2CleDoc(TOBPiece_O);
for i:=0 to TobPiece.Detail.Count-1 do  { NEWPIECE }
begin
  Tobl := TobPiece.Detail[i];
  Ok_ReliquatMT := CtrlOkReliquat(TOBL, 'GL');
  CleDoc := TOB2CleDoc(TOBL);
  if CompareCleDoc(CleDoc, CleDoc_O) then
  begin
    { Récupère la ligne dans la pièce d'origine }
    TobLOld := GetTOBPrec(TobL, TobPiece_O, True);
    if TobLOld = Nil then Continue;
    if TobL.GetValue('GL_SOLDERELIQUAT') = '-' then
    begin
      { Mise à jour du reste à livrer de la ligne de la pièce d'origine }
      TobLOld.PutValue('GL_QTERESTE', Arrondi(TobLOld.GetValue('GL_QTERESTE') - TobL.GetValue('GL_QTESTOCK'), 6));
      if TobLOld.GetValue('GL_QTERESTE') < 0 then TobLOld.PutValue('GL_QTERESTE', 0);
      // --- GUINIER ---
      if Ok_ReliquatMT then
      begin
        TobLOld.PutValue('GL_MTRESTE', Arrondi(TobLOld.GetValue('GL_MTRESTE') - TobL.GetValue('GL_MONTANTHTDEV'), 3));
        if TobLOld.GetValue('GL_MTRESTE') < 0 then TobLOld.PutValue('GL_MTRESTE', 0);
      end;
    end
    else
    begin
      TobLOld.PutValue('GL_QTERESTE', 0);
      // --- GUINIER ---
      if Ok_ReliquatMT then TobLOld.PutValue('GL_MTRESTE', 0);
    end;
  end;
end;
NewEtat := '';
if ToutesLignesSoldees(TobPiece_O) then
  NewEtat := '-'
else
  NewEtat := 'X';
TobPiece_O.PutValue('GP_VIVANTE', NewEtat);
{ Etat des lignes = Etat de l'entête }
for i := 0 to TobPiece_O.Detail.Count - 1 do
  TobPiece_O.detail[i].PutValue('GL_VIVANTE', NewEtat);
RefSuiv:=EncodeRefPiece(TOBPiece) ;
TOBPiece_O.PutValue('GP_DEVENIRPIECE',RefSuiv) ;
TOBPiece_O.PutValue('GP_REFCOMPTABLE','') ;
//TOBPiece_O.UpdateDB(False);
Result:=True ;
END ;
// -----

Procedure G_InversePiece ( TOBPiece,TobTiers : TOB ) ; { NEWPIECE }
Var CleDoc : R_CleDoc;
    TOBBases_O, TOBEches_O, TOBN_O, TOBPiece_O, TOBAcomptes_O, TOBPorcs_O, TobDetPiece_O, TOBSerie_O: TOB;
    TobLigneTarif_O: TOB;

    TobVariable_O, TobRevision_O: TOB;
    NowFutur : TDateTime ;
//    PasHisto : boolean ;
    Q        : TQuery ;
    OldEcr,OldStk   : RMVT ;
    i        : integer;
    // Modif BTP
    TOBOuvrage_O,TOBLIENOLE_O,TOBPIECERG_O,TOBBasesRG_O : TOB;
    IndiceOuv : integer;
    // Modif MODE
    TOBL : TOB;
    Qte, Ecart : double;
//    Val : variant;
    {$IFNDEF BTP}
    RazOk : boolean;
    {$ENDIF}
    // -----
    OldCleDoc: R_CleDoc;
    OldTobL, NewTobL: Tob;
    MajQteReste   : boolean;
    OK_ReliquatMt : Boolean;
    MtReste       : Double;
    EcartMt       : Double;
BEGIN
  //MajQteReste := true;
  // Modif MODE
  {$IFNDEF BTP}
  RazOk := True;
  {$ENDIF}
  // -----
  IndiceOuv := 1;
  CleDoc:=TOB2CleDoc(TOBPiece) ;
  TOBPiece_O:=TOB.Create('PIECE',Nil,-1) ;
  // Modif BTP
  AddLesSupEntete (TOBPiece_O);
  // ---
  TOBPiece_O.Dupliquer(TOBPiece,False,True) ;
  {$IFDEF BTP}
  MajQteReste :=  (GetInfoParPiece (TOBPiece_O.GetValue('GP_NATUREPIECEG'),'GPP_RELIQUAT')='X');
  {$ENDIF}

  // Au niveau filles on duplique uniquement les lignes ( sans analytique rajoutée en dessous)
  for i :=0 to TOBPiece.Detail.count-1 do
  BEGIN
    // Modif MODE : ne pas reprendre la ligne commentaire ajoutée
    if not (TOBPiece.Detail[i].FieldExists('LIGNECOMMENT')) or (TOBPiece.Detail[i].GetValue('LIGNECOMMENT') <> 'X') then
    BEGIN
      TobDetPiece_O := TOB.Create('LIGNE', TOBPiece_O, -1);
      TOBDetPiece_O.Dupliquer(TOBPiece.Detail[i], False, True);
    END;
    // -----
  END;

  TOBPiece_O.SetAllModifie(False) ;

  // NAC - Optimisation génération des prépas
  // ----------
  // Modif MODE
  if CodeAff <> '' then
  begin
    {Calcul de la quantité restante pour chaque ligne de la pièce d'origine}
    for i := 0 to TOBPiece_O.Detail.Count - 1 do
    begin
      TOBL := TOBPiece_O.Detail[i];
      if (TOBL.GetValue('GL_ARTICLE') = '') and (TOBL.GetValue('GL_CODESDIM') = '') then Continue;
      //
      OK_ReliquatMt := CtrlOkReliquat(TOBL, 'GL');
      //
      Qte := G_GetQteAffectee(TOBL);
      Ecart := Arrondi(TOBL.GetValue('GL_QTERESTE') - Qte, 6);
      if Ecart > 0 then TOBPiece_O.Modifie := True;
      if OK_ReliquatMt then
      begin
        MtReste := G_GetMtAffectee(TOBL);
        EcartMt := Arrondi(TOBL.GetValue('GL_MTRESTE') - MTReste, 3);
        if EcartMt > 0 then TOBPiece_O.Modifie := True;
      end;
      if not TOBPiece_O.Modifie then Break;
    end;

    {Génération des pièces reliquats}
    if TOBPiece_O.Modifie then
    begin
      // Suppression MODE 30/10/2003
      //RazOk := False;
    end;
  end;

  // Numeros de série
  TOBSerie_O:=TOB.Create('',Nil,-1) ;
  TOBSerie_O.Dupliquer(TOBSerie,True,True) ;
  TOBSerie_O.SetAllModifie(False) ;

  // Lecture nomenclatures
  TOBN_O:=TOB.Create('NOMENCLATURES',Nil,-1) ;
  LoadLesNomen(TOBPiece,TOBN_O,TOBArticles,CleDoc) ;

  //if RazOk then
  //begin
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
  TOBPieceRg_O.selectDB ('',Q) ;
  Ferme(Q) ;
  TOBBasesRg_O:=TOB.create('BASESRG',nil,-1);
  Q:=OpenSQL('SELECT * FROM PIEDBASERG WHERE '+WherePiece(CleDoc,ttdBaseRG,False),True,-1, '', True) ;
  TOBBasesRg_O.selectDB ('',Q) ;
  Ferme(Q) ;

  TOBLigneTarif_O:=TOB.create('_LIGNETARIF_',nil,-1);
(* SUPPRIME -- SUITE A DROPTABLE DANS PGIMAJVER
  Q:=OpenSQL('SELECT * FROM LIGNETARIF WHERE '+WherePiece(CleDoc,ttdLigneTarif,False),True) ;
  TOBLigneTarif_O.LoadDetailDB('LIGNETARIF', '', '', Q, False);
  Ferme(Q) ;
*)
  //Affaire
  TobVariable_O := TOB.Create('', nil, -1);
  if GetParamSoc('SO_AFVARIABLES')  then
  begin
    Q := OpenSQL('SELECT * FROM AFORMULEVARQTE WHERE ' + WherePiece(CleDoc, ttdVariable, False), True,-1, '', True);
    TobVariable_O.LoadDetailDB('AFORMULEVARQTE', '', '', Q, False);
    Ferme(Q);
  end;
  TobRevision_O := TOB.Create('', nil, -1);
  if GetParamSoc('SO_AFREVISIONPRIX') then
  begin
    Q := OpenSQL('SELECT * FROM AFREVISION WHERE ' + WherePiece(CleDoc, ttdRevision, False), True,-1, '', True);
    TobRevision_O.LoadDetailDB('AFREVISION', '', '', Q, False);
    Ferme(Q);
  end;
  //end;

{Mise à jour inverse}
NowFutur:=NowH ; TOBPiece_O.SetDateModif(NowFutur) ; TOBPiece_O.CleWithDateModif:=True ;

{$IFDEF AFFAIRE}
// gm 16/10/03 pour ne pas supprimer la compta des factures,qd on fait
// une génération automatique de factures en avoirs
// à généraliser plus tard si cette fonction est reprise ailleurs
//if  (ctxTempo in V_PGI.PGIContexte) then
//begin
  if not(GetInfoParPiece(Tobgenere.GetValue('GP_NATUREPIECEG'), 'GPP_ESTAVOIR') = 'X') then
  begin
    if V_PGI.IoError=oeOk then DetruitCompta(TOBPiece_O,NowFutur,OldEcr,OldStk) ;
  end;
//end
//else
{$ENDIF}
//if V_PGI.IoError=oeOk then DetruitCompta(TOBPiece_O,OldEcr,OldStk) ;
if V_PGI.IoError=oeOK then
begin
  {Flag pour InverseStockTransfo}
  TobPiece_O.AddChampSupValeur('GENAUTO', 'X');
  InverseStockTransfo(TOBPiece_O, TobGenere, TOBArticles, TOBCatalogu, TOBOuvrage_O);
end;

// Modif MODE
{$IFNDEF BTP}
if RazOk then
begin
{$ENDIF}
  { Remise à jour du Reste à livrer dans la pièce d'origine }
  { TobPiece_O contient l'ancienne pièce, TobGenere contient la nouvelle piece (pas encore totalement mise à jour) }
  if MajQteReste then
  begin
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
          // --- GUINIER ---
          OldTobL.PutValue('GL_QTERESTE', OldTobL.GetValue('GL_QTERESTE') - NewTobl.GetValue('GL_QTESTOCK'));
          if OldTOBL.getValue('GL_QTERESTE') < 0 then OldTobL.PutValue('GL_QTERESTE', 0);
          //
          if OK_ReliquatMt then
          begin
            OldTobL.PutValue('GL_MTRESTE', OldTobL.GetValue('GL_MTRESTE') - NewTobl.GetValue('GL_MONTANTHTDEV'));
            if OldTOBL.getValue('GL_MTRESTE') < 0 then OldTobL.PutValue('GL_MTRESTE', 0);
          end;
        end
        else
          if OldTobl.FieldExists ('QTERESTE') then
          begin
            OldTobl.PutValue ('GL_QTERESTE', OldTobl.GetValue ('QTERESTE'));
            TobPiece_O.Modifie := true;
          end;
          // --- GUINIER ---
          if OK_ReliquatMt then
          begin
            if OldTobl.FieldExists ('MTRESTE') then
            begin
              OldTobl.PutValue ('GL_MTRESTE', OldTobl.GetValue ('MTRESTE'));
              TobPiece_O.Modifie := true;
            end;
          end;
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
      {$IFDEF AFFAIRE}
      if not DetruitAffaireComplement(TobVariable_O, TobRevision_O) then
        V_PGI.IoError := oeSaisie;
      {$ENDIF}
      if not DetruitAncien(TOBPiece_O,TOBBases_O,TOBEches_O,TOBN_O,nil,TOBAcomptes_O,TOBPorcs_O,TOBSerie_O,TOBOuvrage_O,TOBLienOle_O,TOBPieceRG_O,TOBBasesRg_O,TobLigneTarif_O) then
        V_PGI.IoError:=oeSaisie ;
    end;
  end
  else
  begin
    if UpdateAncien(TOBPiece_O,TOBGenere,TOBAcomptes_O,TobTiers,True) then
    begin
      { NEWPIECE }
      if TOBPiece_O.UpdateDB(False) then
      begin
        { MODE DCA 28/08/2003 - Génération de cdes fourn.(CF) depuis des cdes de réassort(FCF) }
        { Ajout requête pour forcer la mise à jour de GP_DATEMODIF : non modifiable avant pour gérer le multi-utilisateur : TOBPiece_O.CleWithDateModif  }
        if (ExecuteSQL('UPDATE PIECE SET GP_DATEMODIF="' + USTime(NowFutur) + '" WHERE ' + WherePiece(CleDoc, ttdPiece,
          False)) <> 1) then V_PGI.IoError := oeSaisie;
      end
      else V_PGI.IoError := oeSaisie;
      RazLesSeries(nil,TobSerie_O);
      if not TOBSerie_O.UpdateDB(False) then V_PGI.IoError := oeSaisie;
    end
    else
      V_PGI.IoError := oeSaisie;
  end;
{$IFNDEF BTP}
end else
begin
  // Ajout MODE 30/10/2003
  CreerReliquat(TOBGenere, TOBPiece_O);
  if CodeAff <> '' then G_ValideLesLignesGeneriques(TOBPiece_O);
  if not TOBPiece_O.UpdateDB(False) then V_PGI.IoError := oeSaisie;
end ;
{$ENDIF}
//if V_PGI.IoError=oeOk then ValideLesArticles(TOBPiece_O,TOBArticles) ;
// La mise a jour différentiel du stock implique qu'on ne touche pas tobdispo ici mais a la validation de la pièce regroupante DBR
{$IFDEF AFFAIRE}
if V_PGI.IoError=oeOk then ValideActivite(Nil,TOBPiece_O, TOBArticles, GereActivite, False,DelActivite,False);
{$ENDIF}
{$IFDEF BTP}
if (V_PGI.IoError=oeOk) then MajQteFacture(TobGenere, TobPiece_O);
{$ENDIF}
{$IFNDEF CCS3}
if V_PGI.IoError=oeOk then InverseStockSerieTransfo(TOBPiece_O, TobPiece, TOBSerie_O) ; { NEWPIECE }
{$ENDIF}
{Libérations}
TOBPiece_O.Free ; TOBBases_O.Free ; TOBEches_O.Free ; TOBN_O.Free ; TOBAcomptes_O.Free ; TOBPorcs_O.Free ;
TOBSerie_O.Free ;
// Modif BTP
TOBOuvrage_O.free; TOBPIECERG_O.free; TOBBasesRG_O.free; TOBLIenOLE_O.free;
TobLigneTarif_O.Free; 
//Affaire-ONYX
TobVariable_O.free;
TobRevision_O.free;
// --------
END ;

Procedure G_LigneVersPieceGenere ( TOBL : TOB; DupPiece:Boolean=False ) ;
BEGIN
TOBGenere.PutValue('GP_ETABLISSEMENT',TOBL.GetValue('GL_ETABLISSEMENT')) ;
// Transformation de Commande de réassort
if TOBL.GetValue('GL_NATUREPIECEG')='FCF' then
   begin
   TOBGenere.PutValue('GP_TIERS',TOBL.GetValue('GL_FOURNISSEUR')) ;
   TOBGenere.PutValue('GP_TIERSFACTURE',TOBL.GetValue('GL_FOURNISSEUR')) ;
   TOBGenere.PutValue('GP_TIERSPAYEUR',TOBL.GetValue('GL_FOURNISSEUR')) ;
   TOBGenere.PutValue('GP_TIERSLIVRE',TOBL.GetValue('GL_FOURNISSEUR')) ;
   end else
   begin
   TOBGenere.PutValue('GP_TIERS',TOBL.GetValue('GL_TIERS')) ;
   TOBGenere.PutValue('GP_TIERSFACTURE',TOBL.GetValue('GL_TIERSFACTURE')) ;
   TOBGenere.PutValue('GP_TIERSPAYEUR',TOBL.GetValue('GL_TIERSPAYEUR')) ;
   TOBGenere.PutValue('GP_TIERSLIVRE',TOBL.GetValue('GL_TIERSLIVRE')) ;
   end ;
if Not DupPiece then
   begin
//   TOBGenere.PutValue('GP_REPRESENTANT',TOBL.GetValue('GL_REPRESENTANT')) ; fiche eQualite 10425
   TOBGenere.PutValue('GP_APPORTEUR',TOBL.GetValue('GL_APPORTEUR')) ;
   TOBGenere.PutValue('GP_DEPOT',TOBL.GetValue('GL_DEPOT')) ;
   TOBGenere.PutValue('GP_SAISIECONTRE',TOBL.GetValue('GL_SAISIECONTRE')) ;
   TOBGenere.PutValue('GP_REGIMETAXE',TOBL.GetValue('GL_REGIMETAXE')) ;
   TOBGenere.PutValue('GP_FACTUREHT',TOBL.GetValue('GL_FACTUREHT')) ;
   {Modif JLD 20/06/2002}
   TOBGenere.PutValue('GP_ESCOMPTE',TOBL.Parent.GetValue('GP_ESCOMPTE')) ;
   TOBGenere.PutValue('GP_REMISEPIED',TOBL.Parent.GetValue('GP_REMISEPIED')) ;
   TOBGenere.PutValue('GP_TVAENCAISSEMENT',TOBL.Parent.GetValue('GP_TVAENCAISSEMENT')) ;
   {Fin Modif JLD}
   TOBGenere.PutValue('GP_DATELIVRAISON',TOBL.GetValue('GL_DATELIVRAISON')) ;
   end;
{Traitement des devises}
DEV.Code:=TOBL.GetValue('GL_DEVISE') ; GetInfosDevise(DEV) ; TOBGenere.PutValue('GP_DEVISE',DEV.Code) ;
DEV.Taux:=TOBL.GetValue('GL_TAUXDEV') ; TOBGenere.PutValue('GP_TAUXDEV',DEV.Taux) ;
TOBGenere.PutValue('GP_COTATION',TOBL.GetValue('GL_COTATION')) ;
{Traitement Tiers}
if TOBGenere.GetValue('GP_MODEREGLE') = '' then  // Ajout TG 11/09/2001 (condition)
 TOBGenere.PutValue('GP_MODEREGLE',TOBTiers.GetValue('T_MODEREGLE')) ;
{if not VH_GC.GCIfDefCEGID then
{Traitement Affaire}
{if ctxGCAFF in V_PGI.PGIContexte then
   begin
   TOBGenere.PutValue('GP_AFFAIRE', TOBL.GetValue('GL_AFFAIRE')) ;
   TOBGenere.PutValue('GP_AFFAIRE1',TOBL.GetValue('GL_AFFAIRE1')) ;
   TOBGenere.PutValue('GP_AFFAIRE2',TOBL.GetValue('GL_AFFAIRE2')) ;
   TOBGenere.PutValue('GP_AFFAIRE3',TOBL.GetValue('GL_AFFAIRE3')) ;
   TOBGenere.PutValue('GP_AVENANT', TOBL.GetValue('GL_AVENANT')) ;
   end;}

{Cas particuliers}
if EclateDomaine then TOBGenere.PutValue('GP_DOMAINE',TOBL.GetValue('GL_DOMAINE')) ;
END ;

Function G_DiffChampsEntete ( TOBPiece : TOB ; Champ : String ) : boolean ;
Var V1,V2 : Variant ;
BEGIN
Result:=False ;
if TOBPiece.FieldExists(Champ) then
   BEGIN
   V1:=TOBPiece.GetValue(Champ) ; V2:=TOBRefPiece.GetValue(Champ) ;
     if V1<>V2 then
     begin
     	Result:=True ;
     end;
   END ;
END ;

Function G_yaRuptureEntete ( TOBPiece : TOB; TestReroupementAff : boolean=true ) : Boolean ;
Var i : integer ;
    Champ,Junk : String ;
BEGIN
Result:=True ;
{Tester différences génériques}
for i:=0 to ChampsRupt.Count-1 do
    BEGIN
    Champ:=ChampsRupt[i] ; Junk:=ReadTokenSt(Champ) ; Champ:=ReadTokenSt(Champ) ;
    if G_DiffChampsEntete(TOBPiece,Champ) then Exit ;
    END ;
{$IFDEF BTP}
//MODIFBTP : test si affaire avec ou sans regroupement
if TestReroupementAff then
begin
	if RenvoieCodeReg(TOBPiece.Getvalue('GP_AFFAIRE')) = 'AUC' then Exit;
end;
// modif BRL 27/01/05 if RenvoieCodeReg(TOBPiece.Getvalue('GP_AFFAIREDEVIS')) = 'AUC' then Exit;
{$ENDIF}
{Tout est OK}
Result:=False ;
END ;

Function G_DiffChamps ( TOBL : TOB ; Champ : String ) : boolean ;
Var V1,V2 : Variant ;
BEGIN
Result:=False ;
if (TOBL.FieldExists(Champ)) and (TOBL.GetValue('GL_REFARTSAISIE')<>'') then
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
if EclateParPiece Then begin Result := False; Exit; end; // Une piece source = Une piece dest. pas d'éclatement par ligne PA 02/10/2001
{Tester différences génériques}
for i:=0 to ChampsRupt.Count-1 do
    BEGIN
    Champ:=ChampsRupt[i] ; Junk:=ReadTokenSt(Champ) ; Champ:=ReadTokenSt(Champ) ;
    if G_DiffChamps(TOBL,Champ) then Exit ;
    END ;
{Eclater par domaine}
if EclateDomaine then
   BEGIN
   Champ:='GL_DOMAINE'  ; if G_DiffChamps(TOBL,Champ) then Exit ;
   END ;
{if not VH_GC.GCIfDefCEGID then
{Ne pas regrouper sur codes affaires différents}
{if ctxGCAFF in V_PGI.PGIContexte then  // En contexte affaire multiaffaire / pièce autorisé
   BEGIN
   Champ:='GL_AFFAIRE'  ; if G_DiffChamps(TOBL,Champ) then Exit ;
   Champ:='GL_AFFAIRE1' ; if G_DiffChamps(TOBL,Champ) then Exit ;
   Champ:='GL_AFFAIRE2' ; if G_DiffChamps(TOBL,Champ) then Exit ;
   Champ:='GL_AFFAIRE3' ; if G_DiffChamps(TOBL,Champ) then Exit ;
   Champ:='GL_AVENANT'  ; if G_DiffChamps(TOBL,Champ) then Exit ;
   END; }
{Tout est OK}
Result:=False ;
END ;

Procedure G_ChargeCompletePiece ( TOBPiece : TOB ; ModeGenere : string=''; WithRg : boolean= true) ;
BEGIN
G_ChargeLesLignes(TOBPiece, ModeGenere) ;
if WithRg then UG_ChargeLesRG (TOBPiece,TOBPieceRG, TOBBasesRG);
UG_AjouteLesArticles(TOBPiece,TOBArticles,TOBCpta,TOBTiers,TOBAnaP,TOBAnaS,TOBCatalogu,False) ;
UG_AjouteLesRepres(TOBPiece,TOBComms) ;
// NAC - Optimisation génération des prépas
TOBArticles.SetAllModifie(False);
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

{============================ T_GenVal ========================}
procedure T_GenVal.AddLesSupEnteteBtp (TOBPIECE : TOB);
begin
if not (TOBPiece.FieldExists ('AFF_GENERAUTO')) then TOBPIece.addchampSupValeur ('AFF_GENERAUTO','',true);
if not (TOBPiece.FieldExists ('AFF_OKSIZERO')) then TOBPIece.addchampSupValeur ('AFF_OKSIZERO','',true);
if not (TOBPiece.FieldExists ('RUPTMILLIEME')) then TOBPIece.addchampSupValeur ('RUPTMILLIEME','',true);
end;

Procedure T_GenVal.G_GenereParTiers ;
Var i,k,LastI,iacc : integer ;
    TOBL,TOBPiece,TOBPieceTiers,NewL,TOBTP,TOBSOUSTRAIT : TOB ;
    Premier,NewPiece,DupPiece,InitRem : boolean ;
    RemCur : double ;
    // Modif MODE
    RefPiece    : string;
    Qte, Ecart  : double;
    Mt, MtEcart : Double;
    // -----
    // MODIF_DBR_DEBUT
    OldTobl, NewTobl : Tob;
    OldCleDoc,LocCledoc: R_CleDoc;
    iOld : integer;
    // MODIF_DBR_FIN
    LocalModegestion : string;
    Naff0,Naff1,Naff2,Naff3,NAvenant : string;
    Ok_ReliquatMt : Boolean;
    TOBAFFAIREL : TOB;
    QQ : TQuery;
BEGIN
  RemCur := 0;
TOBSOUSTRAIT := TOB.Create ('LES SOUS TRAITANTS',nil,-1);

InitRem := False;
LocalModegestion := ModeTraitement;
TOBArticles.ClearDetail ;
Premier:=True ; NewPiece:=False ; DupPiece:=False ; LastI:=0 ;

if RemplirTOBTiers(TOBTiers,CodeTiers,NewNature,False)<>trtOk then V_PGI.IoError:=oeUnknown ;
TOBPieceTiers:=TOB.Create('PieceMemeTiers',Nil,-1) ;

  if NewAffaire <> '' then
  begin
    BTPCodeAffaireDecoupe (NewAffaire,Naff0,Naff1,Naff2,Naff3,NAvenant,TaModif,True);
    TOBANAP.clearDetail;
    TOBANAS.clearDetail;
  END;
  if ModeTraitement = 'ETUTODEV' then
  BEGIN
    TOBANAP.clearDetail;
    TOBANAS.clearDetail;
  end;

  for i:=0 to TOBSource.Detail.Count-1 do
  BEGIN
    if TOBSource.Detail[i].GetValue('GP_TIERS')<>CodeTiers then Continue ;
    TOBPiece:=TOB.Create('PIECE',TOBPieceTiers,-1);
    LocCledoc := TOB2CleDoc(TOBSource.Detail[i]);
    // Modif BTP
    //AddLesSupEntete (TOBPiece);
    // ---
    //TOBPiece.Dupliquer(TOBSource.Detail[i],True,True);
    LoadPiece (LocCledoc,TOBpiece);
    if ModeTraitement = 'ETUTODEV' then
    begin
      TOBPiece.SetString('GP_AFFAIRE',TOBSource.Detail[i].GetString('GP_AFFAIRE'));
      TOBPiece.SetString('GP_AFFAIRE1',TOBSource.Detail[i].GetString('GP_AFFAIRE1'));
      TOBPiece.SetString('GP_AFFAIRE2',TOBSource.Detail[i].GetString('GP_AFFAIRE2'));
      TOBPiece.SetString('GP_AFFAIRE3',TOBSource.Detail[i].GetString('GP_AFFAIRE3'));
      TOBPiece.SetString('GP_AVENANT',TOBSource.Detail[i].GetString('GP_AVENANT'));
    end;
  END;

  for i:=0 to TOBPieceTiers.Detail.Count-1 do
  BEGIN
    TOBPiece:=TOBPieceTiers.Detail[i] ;
    G_ChargeCompletePiece(TOBPiece,LocalModeGestion,false) ;
    if i=0 then
    begin
    	TOBRefPiece.Dupliquer(TOBPiece,False,True) ;
      LocCledoc := TOB2CleDoc(TOBpiece);
      LoadLesSousTraitants(LocCledoc ,TOBSOUSTRAIT);
      Setallused (TOBSOUSTRAIT);
    end;
    if EclateParPiece then BEGIN NewPiece:=True ; DupPiece:=True ; END ;
{$IFDEF BTP}
    if i=0 then
    begin
      RemCur:=TOBPiece.GetValue('GP_REMISEPIED') ;
    end else
    begin
      if TOBPiece.GetValue('GP_REMISEPIED')<>RemCur then InitRem:=True ;
    end;
{$ENDIF}
    if DeGroupeRemise then
       BEGIN
       if i=0 then RemCur:=TOBPiece.GetValue('GP_REMISEPIED') else
        if TOBPiece.GetValue('GP_REMISEPIED')<>RemCur then NewPiece:=True ;
       END ;
    if G_yaRuptureEntete (TOBPiece,((NewNature<>'FBC')and (NewNature<>'FF')and (NewNature<>'AF')) ) then
       BEGIN
       NewPiece:=True ;
       TOBRefPiece.Dupliquer(TOBPiece,False,True) ;
       END ;

       if not NewPiece then
       begin

        LocCledoc := TOB2CleDoc(TOBpiece);
        IF TOBSOUSTRAIT.detail.Count = 0 then
        begin
          LoadLesSousTraitants(LocCledoc ,TOBSOUSTRAIT);
          Setallused (TOBSOUSTRAIT);
        end else
        begin
          CompletelesSSTrait(LocCledoc,TOBSOUSTRAIT);
        end;
       end;
    for k:=0 to TOBPiece.Detail.Count-1 do
        BEGIN
        TOBL:=TOBPiece.Detail[k] ;
        if ModeTraitement = 'ETUTODEV' then
        BEGIN
          if TOBL.detail.count > 0 then TOBL.detail[0].clearDetail;
          if TOBL.detail.count > 1 then TOBL.detail[1].clearDetail;
          TOBL.SetString('GL_AFFAIRE', TOBPiece.GetString('GP_AFFAIRE')) ;
          TOBL.SetString('GL_AFFAIRE1',TOBPiece.GetString('GP_AFFAIRE1')) ;
          TOBL.SetString('GL_AFFAIRE2',TOBPiece.GetString('GP_AFFAIRE2')) ;
          TOBL.SetString('GL_AFFAIRE3',TOBPiece.GetString('GP_AFFAIRE3')) ;
          TOBL.SetString('GL_AVENANT', TOBPiece.GetString('GP_AVENANT')) ;
        END;
        // --- GUINIER ---
        Ok_ReliquatMt := CtrlOkReliquat(TOBL, 'GL');
        if ((Premier) or (NewPiece) or (G_yaRupture(TOBL))) then
           BEGIN
           if Not Premier then
              BEGIN
              TOBAcomptes.ClearDetail ; TOBPorcs.ClearDetail ; TOBAcomptesGen_O.clearDetail;
              Tobrevision.ClearDetail ; TobVariable.ClearDetail; TOBPieceTrait.ClearDetail;
              TOBSSTRAIT.ClearDetail;
{$IFDEF BTP}
							TheRepartTva.Initialise;
{$ENDIF}
              for iacc:=LastI to i-1 do
                  BEGIN
                  if newNature <> 'AVC' then
                  begin
                  if TOBAcomptesGen.detail.count = 0 then
                     begin
                     UG_ChargeLesAcomptes(TOBPieceTiers.Detail[iacc],TOBAcomptes);
                     end else
                     begin
                     UG_RecupLesAcomptes (TOBPieceTiers.Detail[iacc],TOBAcomptes,TOBAcomptesGen,TOBAcomptesGen_O);
                     end;
                  end else
                  begin
                  	TOBgenere.putValue('GP_ACOMPTEDEV',0);
                  	TOBgenere.putValue('GP_ACOMPTE',0);
                  end;
                  UG_ChargeLesPorcs(TOBPieceTiers.Detail[iacc],TOBPorcs) ;
                  {$IFDEF AFFAIRE}
                  if GetParamSoc('SO_AFREVISIONPRIX') then G_ChargeLesRevisions(TOBPieceTiers.Detail[iacc], TOBRevision);
                  if GetParamSoc('SO_AFVARIABLES')    then G_ChargeLesVariables(TOBPieceTiers.Detail[iacc], TOBVariable);
                  {$ENDIF}
                  {Inversion et mise à jour pièces d'origine}
                  G_InversePiece(TOBPieceTiers.Detail[iacc], TobTiers);
                  END ;
             // MODIF_DBR_DEBUT
              { Remise à jour du Reste à livrer dans la pièce d'origine }
              { TobPiece_O contient l'ancienne pièce, TobGenere contient la nouvelle piece (pas encore totalement mise à jour) }
              for iOld := 0 to TobPiece.Detail.Count - 1 do
              begin
                OldTobL := TobPiece.Detail[iOld];
                if EstLigneArticle(OldTobL) then
                begin
                  OldCleDoc := Tob2CleDoc(OldTobL);
                  OldCleDoc.NumLigne := OldTobL.GetValue('GL_NUMLIGNE');
                  OldCleDoc.NumOrdre := OldTobL.GetValue('GL_NUMORDRE');
                  NewTobL := FindCleDocInTob(OldCleDoc, TobGenere, True);
                  if Assigned(NewTobL) then
                  begin
                    OldTobL.AddChampSupValeur('QTERESTE', OldTobL.GetValue('GL_QTERESTE') - NewTobl.GetValue('GL_QTESTOCK'));
                    if Ok_ReliquatMt then OldTobL.AddChampSupValeur('MTRESTE',  OldTobL.GetValue('GL_MTRESTE') - NewTobl.GetValue('GL_MONTANTHTDEV'));
                  end;
                end;
              end;
              TOBSSTRAIT.Dupliquer(TOBSOUSTRAIT,True,true);
              TOBSOUSTRAIT.ClearDetail;
             // MODIF_DBR_FIN
              if i>0 then G_GenereLaPiece(TOBPieceTiers.Detail[i-1],nil,TOBPieceTrait,false,false,false,LocalModegestion )
                     else G_GenereLaPiece(TOBPieceTiers.Detail[0],nil,TOBPieceTrait,false,false,false,LocalModegestion ) ;
              LastI:=i ;
              END;
           G_AlimNewPiece(TOBPiece,DupPiece,ModeTraitement) ;
           if ModeTraitement = 'ETUTODEV' then
           begin
             TOBGenere.PutValue('AFF_GENERAUTO', TOBpiece.getString('AFF_GENERAUTO')) ;
           end;
           if NewAffaire <> '' then
           begin
             TOBGenere.PutValue('GP_AFFAIRE', NewAffaire) ;
             TOBGenere.PutValue('GP_AFFAIRE1',Naff1) ;
             TOBGenere.PutValue('GP_AFFAIRE2',Naff2) ;
             TOBGenere.PutValue('GP_AFFAIRE3',Naff3) ;
             TOBGenere.PutValue('GP_AVENANT', Navenant) ;
           end;

           G_ChargeLesAdresses(TOBPiece);
           G_LigneVersPieceGenere(TOBL,DupPiece) ;
           END ;
        //FV1 : 22/06/2015 -FS#1393 - SOTRELEC - problème génération d'avoir global à partir facture sans code affaire en en-tête
        {if (ModeTraitement = 'GENAVOIRGLOBAL') then
        begin
          TOBL.SetString('GL_AFFAIRE', TOBPiece.GetString('GP_AFFAIRE'));
          TOBL.SetString('GL_AFFAIRE1',TOBPiece.GetString('GP_AFFAIRE1'));
          TOBL.SetString('GL_AFFAIRE2',TOBPiece.GetString('GP_AFFAIRE2'));
          TOBL.SetString('GL_AFFAIRE3',TOBPiece.GetString('GP_AFFAIRE3'));
          TOBL.SetString('GL_AVENANT', TOBPiece.GetString('GP_AVENANT'));
        end
        else}
        if (ModeTraitement = 'GENEREFF') AND (TobL.GetString('GL_AFFAIRE')<> TOBGenere.GetSTring('GP_AFFAIRE')) then
        begin
          TOBGenere.SetString('GP_AFFAIRE','');
          TOBGenere.SetString('GP_AFFAIRE1','') ;
          TOBGenere.SetString('GP_AFFAIRE2','') ;
          TOBGenere.SetString('GP_AFFAIRE3','') ;
          TOBGenere.SetString('GP_AVENANT', '') ;
        end;

        TOBRef.Dupliquer(TOBL,False,True) ;
      // Ajout de ligne commentaire si rupture sur ligne,non rupture sur Piece et pour la nature FCF
        if (TOBL.GetValue('GL_NATUREPIECEG') = 'FCF') and (not Premier) and (not NewPiece) and (TOBGenere.Detail.count = 0) then
        begin
          // Ligne Commentaire
          NewL:=NewTOBLigne(TOBGenere,0) ;
          NewL.Dupliquer(TOBPiece.Detail[0],True,True) ;
          NewL:=NewTOBLigne(TOBGenere,2) ;
          NewL.Dupliquer(TOBL,True,True) ;
        end else
        begin
          // Suppression MODE 30/10/2003
          NewL:=NewTOBLigne(TOBGenere,0) ;
          NewL.Dupliquer(TOBL,True,True) ;
          NewL.PutValue('GL_RECALCULER','X');

        end ;
{$IFDEF BTP}
        if InitRem then
        begin
          NewL.PutValue('GL_REMISABLELIGNE','-');
          NewL.PutValue('GL_REMISABLEPIED','-');
          NewL.PutValue('GL_TOTREMPIED',0);
          NewL.PutValue('GL_TOTREMPIEDDEV',0);
          NewL.PutValue('GL_REMISEPIED',0);
          NewL.PutValue('GL_TOTREMLIGNE',0);
          NewL.PutValue('GL_TOTREMLIGNEDEV',0);
          NewL.PutValue('GL_REMISELIGNE',0);
        end;
{$ENDIF}
        Premier:=False ;
        if (NewAffaire <> '') then
        begin
          TOBL.PutValue('GL_AFFAIRE', NewAffaire) ;
          TOBL.PutValue('GL_AFFAIRE1',Naff1) ;
          TOBL.PutValue('GL_AFFAIRE2',Naff2) ;
          TOBL.PutValue('GL_AFFAIRE3',Naff3) ;
          TOBL.PutValue('GL_AVENANT', Navenant) ;
        end else if ModeTraitement ='ETUTODEV' then
        begin
          TOBL.SetString('GL_AFFAIRE', TOBPiece.GetString('GP_AFFAIRE')) ;
          TOBL.SetString('GL_AFFAIRE1',TOBPiece.GetString('GP_AFFAIRE1')) ;
          TOBL.SetString('GL_AFFAIRE2',TOBPiece.GetString('GP_AFFAIRE2')) ;
          TOBL.SetString('GL_AFFAIRE3',TOBPiece.GetString('GP_AFFAIRE3')) ;
          TOBL.SetString('GL_AVENANT', TOBPiece.GetString('GP_AVENANT')) ;
        end;

        if RepriseAffaireRef then
           begin
           TOBL.PutValue('GL_AFFAIRE', TOBPiece.GetValue('GP_AFFAIRE')) ;
           TOBL.PutValue('GL_AFFAIRE1',TOBPiece.GetValue('GP_AFFAIRE1')) ;
           TOBL.PutValue('GL_AFFAIRE2',TOBPiece.GetValue('GP_AFFAIRE2')) ;
           TOBL.PutValue('GL_AFFAIRE3',TOBPiece.GetValue('GP_AFFAIRE3')) ;
           TOBL.PutValue('GL_AVENANT', TOBPiece.GetValue('GP_AVENANT')) ;
           end;
        // Transformation de Commande de réassort
        if TOBL.GetValue('GL_NATUREPIECEG')='FCF' then
           begin
           NewL.PutValue('GL_TIERS',TOBL.GetValue('GL_FOURNISSEUR')) ;
           NewL.PutValue('GL_TIERSFACTURE',TOBL.GetValue('GL_FOURNISSEUR')) ;
           NewL.PutValue('GL_TIERSPAYEUR',TOBL.GetValue('GL_FOURNISSEUR')) ;
           NewL.PutValue('GL_TIERSLIVRE',TOBL.GetValue('GL_FOURNISSEUR')) ;
           TOBGenere.PutValue('GP_BLOCNOTE',TOBPiece.GetValue('GP_BLOCNOTE')) ;
           end ;
        // Modif MODE
        if CodeAff <> '' then
        begin
          {Calcul de la quantité et du reliquat pour la pièce résultat à partir de la quantité affectée}
          // NAC - Optimisation génération des prépas
          Qte := G_GetQteAffectee(NewL);
          Ecart := Arrondi(NewL.GetValue('GL_QTERESTE') - Qte, 6);   //Comment que ça marche là pour un montant '??????
          if Ecart < 0 then Ecart := 0;
          NewL.PutValue('GL_QTESTOCK', Qte);
          NewL.PutValue('GL_QTEFACT', Qte);
          NewL.PutValue('GL_QTERESTE', Qte);
          NewL.PutValue('GL_QTERELIQUAT', Ecart);
          if Ok_ReliquatMt then
          begin
            Mt := G_GetMtAffectee(NewL);
            MtEcart := Arrondi(NewL.GetValue('GL_MTRESTE') - Mt, 3);   //Comment que ça marche là pour un montant '??????
            if Ecart < 0 then Ecart := 0;
            NewL.PutValue('GL_MTRESTE', Mt);
            NewL.PutValue('GL_MTRELIQUAT', MtEcart);
          end;
          NewL.PutValue('GL_RECALCULER', 'X');
          RefPiece := EncodeRefPiece(NewL);
          NewL.PutValue('GL_PIECEPRECEDENTE', RefPiece);
          if NewL.GetValue('GL_PIECEORIGINE') = '' then NewL.PutValue('GL_PIECEORIGINE', RefPiece);
        end
        else
        begin
          {V500_003 Début}
          if Tobl.GetValue('GL_QTERESTE') > 0 then
          begin
            NewL.PutValue('GL_QTEFACT',  Tobl.GetValue('GL_QTERESTE'));
            NewL.PutValue('GL_QTESTOCK', Tobl.GetValue('GL_QTERESTE'));
            NewL.PutValue('GL_QTERESTE', Tobl.GetValue('GL_QTERESTE'));
          end
          else
          begin
            if TobL.GetValue('GL_QTEFACT') >= 0 then
            begin
              NewL.PutValue('GL_QTEFACT',  0);
              NewL.PutValue('GL_QTESTOCK', 0);
              NewL.PutValue('GL_QTERESTE', 0);
            end
            else
            begin
              NewL.PutValue('GL_QTEFACT',  TobL.GetValue('GL_QTEFACT'));
              NewL.PutValue('GL_QTESTOCK', TobL.GetValue('GL_QTESTOCK'));
              NewL.PutValue('GL_QTERESTE', TobL.GetValue('GL_QTESTOCK'));
            end;
          end;

          if Ok_ReliquatMt then
          begin
            if Tobl.GetValue('GL_MTRESTE') > 0 then
            begin
              NewL.PutValue('GL_MTRESTE',  Tobl.GetValue('GL_MTRESTE'));
            end
            else
            begin
              if TobL.GetValue('GL_MONTANTHTDEV') >= 0 then
              begin
                NewL.PutValue('GL_MTRESTE',0);
              end
              else
              begin
                NewL.PutValue('GL_MTRESTE',Tobl.GetValue('GL_MONTANTHTDEV'));
              end;
            end;
          end;
          NewL.PutValue('GL_MONTANTHTDEV',  Tobl.GetValue('GL_MONTANTHTDEV'));
          // --
          NewL.PutValue('GL_DATELIVRAISON',TOBGenere.GetValue('GP_DATELIVRAISON'));
          {V500_003 Fin}
        end;
        // -----
        G_LoadNomenLigne(NewL) ;
        G_LoadSerie (NewL) ;
        // Modif BTP
{$IFDEF BTP}
        UG_LoadOuvrageLigne(NewL,TOBOuvrage,TOBArticles,indicenomen,RepriseAffaireRef) ;
{$ENDIF}
        // MODIF BTP
        G_ChargeLesRG(newL) ;
        // -------
        UG_LoadAnaLigne(TOBPiece,NewL) ;
        LoadLigneTarif(TobLigneTarif, NewL);
        NewPiece:=False ;
        END ;
{$IFNDEF CCS3}
    {Exception gestion de documents TIERSPIECE}
    TOBTP:=TOBTiers.FindFirst(['GTP_NATUREPIECEG'],[NewNature],False) ;
    if TOBTP<>Nil then if TOBTP.GetValue('GTP_REGROUPE')<>'X' then BEGIN NewPiece:=True ; DupPiece:=True ; END ;
{$ENDIF}
    END ;
{Dernière rupture}
TOBAcomptes.ClearDetail ; TOBPorcs.ClearDetail ; TOBAcomptesGen_O.clearDetail;
Tobrevision.ClearDetail; TobVariable.ClearDetail;
{$IFDEF BTP}
TheRepartTva.Initialise;
{$ENDIF}
for iacc:=LastI to TOBPieceTiers.Detail.Count-1 do
    BEGIN
    if newNature <> 'AVC' then
    begin
      if TOBAcomptesGen.detail.count = 0 then
      begin
        UG_ChargeLesAcomptes(TOBPieceTiers.Detail[iacc],TOBAcomptes);
      end else
      begin
        UG_RecupLesAcomptes (TOBPieceTiers.Detail[iacc],TOBAcomptes,TOBAcomptesGen,TOBAcomptesGen_O);
      end;
    end else
    begin
      TOBgenere.putValue('GP_ACOMPTEDEV',0);
      TOBgenere.putValue('GP_ACOMPTE',0);
    end;
    UG_ChargeLesPorcs(TOBPieceTiers.Detail[iacc],TOBPorcs) ;
    {Inversion et mise à jour pièces d'origine}
    {$IFDEF AFFAIRE}
    if GetParamSoc('SO_AFREVISIONPRIX') then G_ChargeLesRevisions(TOBPieceTiers.Detail[iacc], TOBRevision);
    if GetParamSoc('SO_AFVARIABLES')    then G_ChargeLesVariables(TOBPieceTiers.Detail[iacc], TOBVariable);
    {$ENDIF}
    G_InversePiece(TOBPieceTiers.Detail[iacc], TobTiers);
    END ;

UG_RegroupeLesAcomptes(TOBAcomptes); // JT eQualité 11013
if TOBPieceTiers.Detail.Count>0 then
begin
	TOBSSTRAIT.Dupliquer(TOBSOUSTRAIT,True,true);
	G_GenereLaPiece(TOBPieceTiers.Detail[TOBPieceTiers.Detail.Count-1],nil,TOBPieceTrait,false,false,false,LocalModegestion);
end;
TOBPieceTiers.Free;
TOBSOUSTRAIT.Free;
END ;

{============================ T_GenereLesPieces ========================}
// Modif BTP
constructor T_GenereLesPieces.create;
begin
  MajPrix:= false;
  // Modif BTP
  OldTypePiece := '';
  UpdateDev := false;
  RendActive := false;
  GestionReliquat := false;
  TOBResult := nil;
  WithLiaison := false;
end;

destructor T_GenereLesPieces.destroy;
begin
  if TOBResult <> nil then TOBResult.free;
  inherited;
end;
// --

procedure T_GenereLesPieces.GenereLesPieces;
Var i,j: integer;
    TOBL, TobLigneTarif, TOBA, TOBCata: TOB; 
    CodeTiers: string;
    prix: double;
    AvecTarif:T_ActionTarifArt;
    LaPiecegeneree : TOB;
    TOBPiece : TOB;
{$IFDEF BTP}
    venteachat : string;
{$ENDIF}
		cledoc : R_CLEDOC;
BEGIN
{$IFDEF BTP}
TOBPiece := TOB.Create ('PIECE', nil,-1);
{$ENDIF}
TOBLigneTarif := nil;
TRY
{$IFDEF BTP}
  if OldTypePiece = '' then exit;
  WithLiaisonInterDoc := WithLiaison;
  TOBPiece.putValue('GP_NATUREPIECEG',OldTypePiece);
{$ENDIF}
  for i:=0 to TOBSource.Detail.Count-1 do
     BEGIN
     TOBArticles.ClearDetail ;
     TOBGenere.ClearDetail;
{$IFDEF BTP}
     // MODIF BTP -- LS
     TOBPiece.ClearDetail;
     TOBOuvrage.cleardetail;
     TOBOuvragesP.cleardetail;
     TOBPieceTrait.Cleardetail; //FV1
     TOBSSTRAIT.ClearDetail; // LS
     // --
{$ENDIF}
     TOBGenere:=TOBSource.Detail[i];
     AddLesSupEntete (TOBGenere);

      if GenereMode = 'DEVTOCTE' then
      begin
        DecodeRefPiece(TOBGenere.GetString('PIECEORIGINE'),cledoc);
        LoadLesSousTraitants (cledoc,TOBSSTRAIT);
    		Setallused (TOBSSTRAIT);
      end;


     CodeTiers:=TOBGenere.getValue('GP_TIERS');
     if RemplirTOBTiers(TOBTiers,CodeTiers,NewNature,False)<>trtOk then V_PGI.IoError:=oeUnknown ;
// MODIF LS
// --
     G_ChargeCompletePiece(TOBGenere,self.GenereMode) ;
  {$IFDEF BTP}
		 VenteAchat := GetInfoParPiece(TobGenere.getValue('GP_NATUREPIECEG'), 'GPP_VENTEACHAT');
     if TOBgenere.FieldExists ('PIECEORIGINE') then
     begin
      RecupereAdressesOrigine (TOBGenere,TOBAdresses);
     end else
  {$ENDIF}
     G_ChargeLesAdresses(TOBGenere);
     TiersVersPiece(TOBTiers,TOBGenere);
  {$IFDEF BTP}
      if (GenereMode = 'CHANTOCBT') or (GenereMode = 'CBTTOLIVC') or (Generemode = 'DEVTOCHAN') or (GenereMode ='CCTOLIVCLI') then
      begin
      	TOBGenere.putValue('GP_FACTUREHT','X'); // correction pour devis TTC
      end;
     if not TOBgenere.FieldExists ('PIECEORIGINE') then
     begin
  {$ENDIF}
     	TiersVersAdresses(TOBTiers,TOBAdresses,TOBGenere) ;
{$IFDEF BTP}
			if venteAchat = 'VEN' Then LivAffaireVersAdresses (nil,TOBAdresses,TOBgenere);
		 end;
     // MODIF BTP -- LS
     UG_ChargePiecesPrecedente (TOBPiece,TOBGenere,GestionReliquat);
     // --
{$ENDIF}
     if (MajPrix) or (UpdateDev) then
        BEGIN
        FillChar(DEV,Sizeof(DEV),#0) ;
        DEV.Code:=TOBGenere.Getvalue('GP_DEVISE') ;
        GetInfosDevise(DEV) ;
        DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,TOBGenere.Getvalue('GP_DATECREATION')) ;
        END ELSE
        BEGIN
        // Correction du DEV non positionné si provenance de creefromligne
        DEV.Code:=TOBGenere.GetValue('GP_DEVISE') ; GetInfosDevise(DEV) ;
        DEV.Taux:=TOBGenere.GetValue('GP_TAUXDEV') ;
        END;

     TOBgenere.putValue('GP_TAUXDEV',DEV.Taux);
     AttribCotation (TOBgenere);
     IndiceNomen := 1;
     for j:=0 to TOBGenere.Detail.Count-1 do
     BEGIN
        TOBL:=TOBGenere.Detail[j] ;
        TOBA:=TOBArticles.FindFirst(['GA_ARTICLE'],[TOBL.getvalue('GL_ARTICLE')],False) ;
        TOBCata:=TOBCatalogu.FindFirst(['GCA_REFERENCE', 'GCA_TIERS'],[TOBL.GetValue('GL_REFCATALOGUE'), GetCodeFourDCM(TOBL)], False) ;

        G_LoadNomenLigne(TOBL) ;
  {$IFDEF BTP}
  			if GenereMode = 'DEVTOCTE' then
        begin
        	G_RecupereDetailOuvrage(TOBL,TOBDESOUVRAGES,TOBOuvrage,TOBArticles,indicenomen,RepriseAffaireRef,NewNature) ;
        end
        else
        begin
          if (GenereMode<>'FACCONSO') and (GenereMode<>'FACDETAILCONSO') then
          begin
        	  UG_LoadOuvrageLigne(TOBL,TOBOuvrage,TOBArticles,indicenomen,RepriseAffaireRef) ;
          end;
        end;
        if (Pos(genereMode,'DEVTOCHAN;CHANTOCBT')>0) then
        begin
          // Dans le cas de génération de contre etude ou de génération de prévision de chantier
          // on affecte le sous traitants au document actuel
          if TOBL.GetString('GLC_NATURETRAVAIL')='002' then
          begin
            UG_SetSSTrait (TOBL,TOBDESSSTRAIT,TOBSSTRAIT);
          end;
        end;

  {$ENDIF}
        UG_LoadAnaLigne(TOBGenere,TOBL) ;
        // MODIF BTP
        G_ChargeLesRG(TOBL) ;
        // -------
        //Prix
        if MajPrix then
        BEGIN
//           Prix:=0;
           AvecTarif:=ataCancel;
           if (GenereMode <> 'REAP') and (GenereMode <> 'REAPECLAT') and
           		(GenereMode <> 'VALIDDECIS') and (GenereMode <> 'VALIDDECISECLAT') then // en cas de REAPPRO, on garde ce qui est dans la piece d'origine
           begin
               IF TOBL.GetValue('GL_TYPEREF')<>'CAT' then
                  AvecTarif:=TarifVersLigne(TobA, TobTiers, TobL, TobLigneTarif, TobGenere, TobTarif, True, True, DEV) ;

               IF (TOBcata<>nil) And ((TOBL.GetValue('GL_TYPEREF')='CAT') or (AvecTarif<>ataOk)) then
               BEGIN
                  Prix:=TOBCata.GetValue('GCA_PRIXBASE');
                  if TOBL.GetValue('GL_FACTUREHT')='X' then
                  BEGIN
                     TOBL.PutValue('GL_PUHT',Prix) ; TOBL.PutValue('GL_PUHTDEV',Prix) ;
                  END else
                  BEGIN
                     TOBL.PutValue('GL_PUTTC',Prix) ; TOBL.PutValue('GL_PUTTCDEV',Prix) ;
                  END ;
                  TOBL.PutValue('GL_DPA',TOBCata.GetValue('GCA_DPA')) ;
               END;
           end;
        END;
     END;

     TOBAcomptes.ClearDetail ;
     UG_ChargeLesAcomptes(TOBGenere,TOBAcomptes) ;
     UG_ChargeLesPorcs(TOBGenere,TOBPorcs) ;
     G_GenereLaPiece(TOBPiece,Nil,TOBPieceTrait,RendActive,GestionReliquat,True,GenereMode);
     // Modif BTP
     if TOBResult <> nil then
     begin
        {$IFDEF BTP}
        if self.GenereMode = 'DEVTOCHAN' then
        begin
          if gestionPhase <> nil then GestionPhase.GenerelesPhases (TOBGenere,nil,false,false,false,TaCreat);
        end else
        if (self.GenereMode = 'CHANTOCBT') or
           (self.GenereMode = 'REAP') or
           (self.GenereMode = 'REAPECLAT') or
           (self.GenereMode = 'VALIDDECIS') or
           (self.GenereMode = 'VALIDDECISECLAT') or
           (self.GenereMode = 'RFOTOLIVC') or
           (self.GenereMode = 'CBTTOLIVC') or
           (Self.GenereMode = 'CCTOLIVCLI') 
           then
        begin
{$IFDEF BTP}
          if (self.GenereMode = 'RFOTOLIVC') then
          begin
            GestionPhase.RepertorieReceptionsHLiensFromLIV (Tobgenere,TOBgenere.GetValue('GP_AFFAIRE'));
          end;
{$ENDIF}
          if gestionPhase <> nil then GestionPhase.AssocieAuxPhases (TobGenere,nil,false,false,false,TaCreat);
        end;
        {$ENDIF}
        LaPieceGeneree := TOB.Create ('PIECE',TOBResult,-1);
        LaPiecegeneree.Dupliquer (TOBgenere,true,true);
     end;
{$IFDEF BTP}
     if gestionPhase <> nil then GestionPhase.Clear;
{$ENDIF}
// --
     END ;
FINALLY
{$IFDEF BTP}
  TOBPiece.free;
{$ENDIF}
END;

end;

{============================ T_ValideNumP ========================}
Procedure T_ValideNumP.PrepareNumPieceTiers;
Var i,j,NbSouche : integer;
    Domaine, Etab, Souche : String;
    bNewSouche : Boolean;
begin
NbSouche :=0;
LesNumPieceTiers := Nil;
for i:=0 to TOBSource.Detail.Count-1 do
   BEGIN
   if TOBSource.Detail[i].GetValue('GP_TIERS') = CodeTiers then
      begin
      bNewSouche := True;
      Domaine := TOBSource.Detail[i].GetValue('GP_DOMAINE');
      Etab := TOBSource.Detail[i].GetValue('GP_ETABLISSEMENT');
      Souche := GetSoucheG(NewNature,Etab,Domaine);
      if Souche = '' then begin V_PGI.IoError:=oeUnknown ; exit; end;
      for j:=Low(LesNumPieceTiers)to High(LesNumPieceTiers) do
      if LesNumPieceTiers[j].Souche = Souche then
         begin
         bNewSouche := False;
         Inc(LesNumPiecetiers[j].NbPiece);
         Break;
         end;

      if bNewSouche then
         begin
         Inc(NbSouche);
         SetLength(LesNumPieceTiers,NbSouche);
         LesNumPieceTiers[NbSouche-1].Nature    := NewNature;
         LesNumPieceTiers[NbSouche-1].Souche    := Souche;
         LesNumPieceTiers[NbSouche-1].DatePiece := TOBSource.Detail[i].GetDateTime('GP_DATEPIECE');
         LesNumPieceTiers[NbSouche-1].NbPiece   := 1;
         end;
      end;
   END;
end;

Procedure T_ValideNumP.G_ValideNumPieceTiers;
Var j : Integer;
		LaDate : TdateTime;
BEGIN
for j:=Low(LesNumPieceTiers)to High(LesNumPieceTiers) do
   BEGIN
   if LesNumPieceTiers[j].Souche <> '' then
      BEGIN
  		if DatePiece <> 0 then LaDate := DatePiece else LaDate := LesNumPieceTiers[j].DatePiece;
      // FV1 :       LesNumPieceTiers[j].NumDebP := SetNumberAttribution (LesNumPieceTiers[j].Souche,LesNumPieceTiers[j].NbPiece);
      LesNumPieceTiers[j].NumDebP := SetNumberAttribution (LesNumPieceTiers[j].Nature, LesNumPieceTiers[j].Souche,LaDate,LesNumPieceTiers[j].NbPiece);
      if LesNumPieceTiers[j].NumDebP = 0 then V_PGI.IoError:=oeUnknown ;
      LesNumPieceTiers[j].NumFinP := LesNumPieceTiers[j].NumDebP + LesNumPieceTiers[j].NbPiece -1;
      END;
   END;
END;


{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 31/07/2001
Modifié le ... : 31/07/2001
Description .. : Génération automatique d'un ensemble de pièces en un
Suite ........ : autre ensemble de nature d'arrivée définie. Tous les
Suite ........ : contrôles et paramétrages sont actifs. Les critères évidents
Suite ........ : d'éclatement, ainsi que certains paramètres passés
Suite ........ : à la fonction déterminent le niveau de regroupement des
Suite ........ : pièces entre elles.
Suite ........ : Si traitement OK, la fonction retourne 0
Mots clefs ... : TRANSFORMATION;GENERATION;REGROUPEMENT;AUTOMATIQUE;
*****************************************************************}
Function RegroupeLesPieces ( TOBSource : TOB ; NewNat : String ; Eclate,DeGroupeRemise,AvecComment : boolean ; ChoixEclateDomaine : integer ;
                              NewDate : TDateTime = 0; AvecCompteRendu : Boolean = True; AlimListePiece : Boolean = False;
                              ContinuOnError : Boolean = False;RepriseAffaireEntete:boolean=false;
                              TraitFinTravaux:boolean=false;CodeAffCde:string='';WithVariantes: boolean=false;NewAffaire:string='';ModeTrait : string=''; MontantHtRef : Double = 0; MontantTTCRef : double= 0; NumfacRef : string='') : integer ;
Var StSort,CodeTiers,StRendu : String ;
    TOBPiece : TOB ;
    i,k,NbErr  : integer ;
    io   : TIoErr ;
    GenVal : T_GenVal ;
    zdeb,zfin,OldNat : string;
    ValideNumP : T_ValideNumP;
BEGIN
Result:=0 ; io:=oeok; NbErr := 0;
{Initialisations}
zdeb := FormatDateTime('ttttt',NowH);
NewNature:=NewNat ; Commentaires:=AvecComment ; bDuplicPiece := False; bContinuOnError := ContinuOnError;
EclateParPiece := Eclate;
bContremarque:=False;
RepriseAffaireRef := RepriseAffaireEntete;
// Modif BTP
FinTravaux := TraitFinTravaux;
ForceNumero := false;
WithLiaisonInterDoc := false;
// Gestion controle facture Achat en génération
FillChar(CleCfa,Sizeof(CleCfa),#0);
CleCfa.MontantHt := MontantHtRef;
CleCFA.MontantTTC := MontantTTCRef;
CleCFA.NumfacFou := NumfacRef;
// --
// Modif MODE
CodeAff:=CodeAffCde;
G_InitLiaisonOrli;
// -----
if NewDate>0 then LaNewDate:=NewDate else LaNewDate:=V_PGI.DateEntree ;
G_InitGenere(ChoixEclateDomaine) ;
{Tri préalable des infos d'origine}
{$IFDEF BTP}
  DefiniTriEtRuptureBtp ( TOBSource,stSort);
{$ELSE}
  StSort:='GP_TIERS;GP_FACTUREHT;GP_REGIMETAXE;GP_TVAENCAISSEMENT;GP_DEVISE;GP_SAISIECONTRE;GP_ESCOMPTE;GP_DATEPIECE;GP_NUMERO;' ;
{$ENDIF}
  TOBSource.Detail.Sort(StSort) ;

  {Création des TOB}
  G_CreerLesTOB ;

  if TOBSource.Detail.Count>0 then OldNat:=TOBSource.Detail[0].GetValue('GP_NATUREPIECEG') else OldNat:='';

  //FV1 : 17/06/2015 - FS#1393 - SOTRELEC - problème génération d'avoir global à partir facture sans code affaire en en-tête (Debut)
  {if ModeTrait = 'GENAVOIRGLOBAL' then
  begin
    G_ChargeChampsRuptavoirGlobal(OldNat);
  end
  //FV1 : 17/06/2015 - FS#1393 - SOTRELEC - problème génération d'avoir global à partir facture sans code affaire en en-tête (fin)
  else }
  if ModeTrait <> 'GENEREFF' then // Modification BRL 29/07/2014 : pas de rupture en génération automatique depuis BSV
    G_ChargeChampsRupt(OldNat) ;
  {Lecture des pièces}
  for i:=0 to TOBSource.Detail.Count-1 do
  BEGIN
    TOBPiece:=TOBSource.Detail[i] ;
    RegroupeLesAcomptesBTP (TOBPiece,TOBAcomptesGen);
    CodeTiers:=TOBPiece.GetValue('GP_TIERS') ;
    if TTiers.IndexOf(CodeTiers)<0 then TTiers.Add(CodeTiers) ;
  END ;

  for k:=0 to TTiers.Count-1 do
  BEGIN
    CodeTiers:=TTiers[k] ;
    GenVal:=T_GenVal.Create ;
    GenVal.CodeTiers:=CodeTiers ; GenVal.TOBSource:=TOBSource ;
    GenVal.DeGroupeRemise:=DeGroupeRemise ;
    GenVal.ChoixEclateDomaine:=ChoixEclateDomaine ;
    GenVal.WithVariantes := WithVariantes;
    GenVal.NewAffaire := NewAffaire;
    GenVal.ModeTraitement := ModeTrait;
// PA 02/10/2001
    if EclateParPiece then // si éclatement on peut déterminer le nombre de pièces et réserver les numéro pour génération multi-utilisateur
        begin
        ValideNumP := T_ValideNumP.Create;
        ValideNumP.TOBSource := TOBSource; ValideNumP.CodeTiers := CodeTiers;
        ValideNumP.PrepareNumPieceTiers;
        ValideNumP.DatePiece := 0;
        if  NewDate > 0 then ValideNumP.DatePiece := NewDate;
        io:=Transactions(ValideNumP.G_ValideNumPieceTiers,5) ;
        ValideNumP.Free;
        end;
    if io=oeOk then io:=Transactions(GenVal.G_GenereParTiers,0);
    GenVal.Free ;
    Case io of
      oeOk : G_MajEvent ;
      oeUnknown  : BEGIN
                   if Not (FromShellOLE) And Not(ContinuOnError) then MessageAlerte('ATTENTION : La génération ne s''est pas complètement effectuée') ;
                   Result:=1 ; if Not(ContinuOnError) then Break else begin Inc(NbErr); io:=oeOk; end;
                   END ;
      oeSaisie   : BEGIN
                   if Not (FromShellOLE) And Not(ContinuOnError) then MessageAlerte('ATTENTION : Certaines pièces déjà en cours de traitement n''ont pas été enregistrées !') ;
                   Result:=2 ; if Not(ContinuOnError) then Break else begin Inc(NbErr); io:=oeOk; end;
                   END ;
      oeLettrage : BEGIN
                   if Not (FromShellOLE) And Not(ContinuOnError) then MessageAlerte('ATTENTION : Certaines pièces ne peuvent pas passer en comptabilité et n''ont pas été enregistrées !') ;
                   Result:=3 ; if Not(ContinuOnError) then Break else begin Inc(NbErr); io:=oeOk; end;
                   END ;
      oeStock    : BEGIN
                   if Not (FromShellOLE) And Not(ContinuOnError) then MessageAlerte('ATTENTION : Stock insuffisant pour certaines pièces qui n''ont pas été enregistrées !') ;
                   Result:=4 ; if Not(ContinuOnError) then Break else begin Inc(NbErr); io:=oeOk; end;
                   END ;
       END ;
    // Modif MODE
    if (CodeAff <> '') and (io <> oeOk) then
       BEGIN
       StRendu := DateTimeToStr(Now) + ' ERREUR n°' + IntToStr(Ord(io)) + ' Affectation n°' + CodeAff
                + ' Tiers n°' + CodeTiers + ' Pièces n°';
       for i := 0 to TOBSource.Detail.Count - 1 do
         BEGIN
         if TOBSource.Detail[i].GetValue('GP_TIERS') <> CodeTiers then Continue;
         StRendu := StRendu + IntToStr(TOBSource.Detail[i].GetValue('GP_NUMERO'))
           + '/' + IntToStr(TOBSource.Detail[i].GetValue('GP_INDICEG')) + ' ';
         END;
       if V_PGI.TransacErrorMessage <> '' then StRendu := StRendu + ' => ' + V_PGI.TransacErrorMessage;
       LogAGL(StRendu);
       END;
    // -----
  END ;
  
{Libérations}
G_FreeLesTOB(AlimListePiece) ;
zfin := FormatDateTime('ttttt',NowH);
{Compte rendu}
if ((AvecCompteRendu) and (PremNum>0) and (io=oeOk)) then
   BEGIN
   StRendu:='Le traitement s''est correctement effectué. #13#10'+IntToStr(NbP)+' pièce(s) générée(s) de nature '
           +RechDom('GCNATUREPIECEG',NewNat,False)+' du N° '+IntToStr(PremNum)+' au N° '+IntToStr(LastNum) ;
   if VH_GC.GCIfDefCEGID then
	StRendu := StRendu + '#13#10 Traitement de '+zdeb+' à '+zfin ;
   if (NbErr<>0) And (ContinuOnError) then  StRendu := StRendu + '#13#10 Nombre de Tiers non traités :' + IntToStr(NbErr);
   PGIInfo(StRendu,'Compte-rendu de génération') ;
   END ;
if ((AlimListePiece) and (PremNum>0) and (io=oeOk)) then
   BEGIN
   if TheTob <> Nil then TheTob.Free; TheTob := TOBListePiece;
   END;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 31/07/2001
Modifié le ... : 31/07/2001
Description .. : Transformation unitaire d'une pièce identifiée par sa nature
Suite ........ : et son numéro. Ne fait aucun contrôle.
Suite ........ : A été développée pour implémentation OLE
Mots clefs ... : TRANSFORMATION;GENERATION;OLE;
*****************************************************************}
Function  TransfoBatchPiece ( OldNat,NewNat,Souche : String ; Numero,Indice : integer ; NewDate : TDateTime; AvecComment : boolean = False; AvecCompteRendu : boolean = False ) : integer ;
Var TOBSource,TOBPiece : TOB ;
    Q                  : TQuery ;
    Find,Vivante       : Boolean ;
    SQL                : String ;
BEGIN
TOBSource:=TOB.Create('Père',Nil,-1) ; Find:=False ; Vivante:=False ;
SQL:='SELECT * FROM PIECE WHERE GP_NATUREPIECEG="'+OldNat+'" AND GP_SOUCHE="'+Souche+'" '
    +'AND GP_NUMERO='+IntToStr(Numero)+' AND GP_INDICEG='+IntToStr(Indice) ;
Q:=OpenSQL(SQL,True,-1, '', True) ;
if Not Q.EOF then
   BEGIN
   TOBPiece:=TOB.Create('PIECE',TOBSource,-1) ;
   // Modif BTP
   AddLesSupEntete (TOBPiece);
   // ---
   TOBPiece.SelectDB('',Q) ; Find:=True ;
   Vivante:=(TOBPiece.GetValue('GP_VIVANTE')='X') ;
   END ;
Ferme(Q) ;
if Not Find then BEGIN Result:=-1 ; TOBSource.Free ; Exit ; END ;
if Not Vivante then BEGIN Result:=-2 ; TOBSource.Free ; Exit ; END ;
FromShellOLE:=True ;
Result:=RegroupeLesPieces(TOBSource,NewNat,True,True,AvecComment,0,NewDate,AvecCompteRendu,False,false,false,false,'',false,'','GENERECF') ;
FromShellOLE:=False ;
{0=OK, -1=pas trouvé, -2=pièce non vivante, 1=???, 2=en cours de traitement, 3=pb comptable, 4=pb stock}
END ;


procedure DefiniLesLiensRegroup (TOBgenere : TOB);
var Indice,INDL : integer;
		TOBP,TOBL,TOBREF : TOB;
    TheAffaire : string;
begin

	for Indice := 0 to TOBGenere.detail.count-1 do
  begin
  	TOBP := TOBGenere.detail[Indice];
    for IndL := 0 TO TOBP.detail.count - 1 do
    begin
  		TOBL := TOBP.detail[IndL];
      if IndL = 0 then
      begin
        TheAffaire := TOBL.GetValue('GL_AFFAIRE');
      end else
      begin
        if TOBL.getValue('GL_AFFAIRE') <> TheAffaire then
        begin
          // REinit du code affaire puisque document multi-affaire
          TOBP.putValue('GP_AFFAIRE','');
          TOBP.putValue('GP_AFFAIRE1','');
          TOBP.putValue('GP_AFFAIRE2','');
          TOBP.putValue('GP_AFFAIRE3','');
          TOBP.putValue('GP_AVENANT','');
        end;
      end;
      if TOBL.FieldExists('RATACHE') then
      begin
        TOBRef := TOBP.FindFirst (['RATACHEMENT'],[TOBL.GetValue('RATACHE')],true);
        if TOBRef <> nil then TOBL.PutValue('GL_LIGNELIEE',TOBRef.GetValue('GL_NUMORDRE'));
      end;
    end;
  end;
end;

procedure PositionneValeurPV (TOBL : TOB; RecalcPA : boolean=true);
var Select : string;
		Q : TQuery;
    CoefPaPr, CoefPrPv ,coefUsUv ,CoefUAUS, FUA,FUS,FUV,PP: double;
    DEV : RDevise;
begin
  // NB provenance d'un document d'achat --> docuent de vente
  if TOBL.GetValue('GL_ARTICLE') = '' then exit;
    //
  DEV.Code := TOBL.GetString('GL_DEVISE');
  GetInfosDevise(DEV) ;
  DEV.Taux:=TOBL.GetValue('GL_TAUXDEV') ;
  
  FUV := RatioMesure('PIE', TOBL.GetValue('GL_QUALIFQTEVTE'));
  if FUV = 0 then FUV := 1;

  FUS := RatioMesure('PIE', TOBL.GetValue('GL_QUALIFQTESTO'));
  if FUS = 0 then FUS := 1;

  FUA := RatioMesure('PIE', TOBL.GetValue('GL_QUALIFQTEACH'));
  if FUA = 0 then FUA := 1;

  COEFUSUV := TOBL.GetDouble('GL_COEFCONVQTEVTE');
  COEFUAUS := TOBL.GetDouble('GL_COEFCONVQTE');
  //
  CoefPaPr := 1;
  CoefPrPv := 1;
  Select := 'SELECT GA_COEFFG,GA_COEFCALCHT FROM ARTICLE WHERE GA_ARTICLE="'+TOBL.GetValue('GL_ARTICLE')+'"';
  Q := OpenSql (Select,true,-1, '', True);
  if not Q.eof then
  begin
    if Q.findfield('GA_COEFFG').AsFloat > 0 then CoefPaPR := Q.findfield('GA_COEFFG').AsFloat;
    if Q.findfield('GA_COEFCALCHT').AsFloat > 0 then CoefPRPV := Q.findfield('GA_COEFCALCHT').AsFloat;
  end;
  ferme (Q);

  if TOBL.getValue('GL_DOMAINE')<>'' Then
  begin
    GetCoefDomaine (TOBL.getValue('GL_DOMAINE'),CoefPaPr,CoefPRPV);
  end;

  //
  if CoefPaPr = 0 then CoefPaPr := 1;
  if CoefPrPV = 0 then CoefPrPV := 1;
  //
  TOBL.putValue('GL_COEFFG',CoefPaPr-1);
  TOBL.putValue('GL_COEFMARG',CoefPrPv);
  TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('GL_COEFMARG')-1)*100,2));

  if RecalcPA then
  begin
    if CoefUAUS <>0 then
    begin
      if coefUsUv <> 0 then
      begin
        TOBL.PutValue('GL_DPA',TOBL.GetValue('GL_PUHTNET')/ CoefUAUS / CoefUSUV );
      end else
      begin
        TOBL.PutValue('GL_DPA',TOBL.GetValue('GL_PUHTNET')/ CoefUAUS * FUA / FUS );
      end;
    end else
    begin
      TOBL.PutValue('GL_DPA',TOBL.GetValue('GL_PUHTNET')* FUS / FUA );
    end;
  end;

  if VH_GC.BTCODESPECIF = '001' then
  begin
    CalculeDonneelignePOC (TOBL,CoefPaPr,CoefPrPv);
    if CoefPaPr = 0 then CoefPaPr := 1;
    if CoefPrPv = 0 then CoefPrPv := 1;
    TOBL.putValue('GL_COEFFG',CoefPaPr-1);
    TOBL.putValue('GL_COEFMARG',CoefPrPv);
    TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('GL_COEFMARG')-1)*100,2));
  end;

  TOBL.PutValue('GL_DPR', Arrondi(TOBL.GEtValue('GL_DPA')*CoefPaPr,V_PGI.okdecP));
  if TOBL.GetValue('GL_DPR') = 0 then TOBL.PutValue('GL_DPR',TOBL.GEtVAlue('GL_DPA'));
  TOBL.PutValue('GL_PUHTDEV',Arrondi(TOBL.GEtValue('GL_DPR')*CoefPRPV,V_PGI.okdecP));
  if TOBL.GetValue('GL_PUHTDEV') = 0 then TOBL.PutValue('GL_PUHTDEV',TOBL.GEtVAlue('GL_DPR'));
  TOBL.putValue('GL_PUHT',DeviseToPivotEx(TOBL.GetValue('GL_PUHTDEV'),DEV.Taux,DEV.quotite,V_PGI.okdecP));
  if TOBL.GetValue('GL_PUHT') <> 0 then
  begin
    PP := Arrondi((TOBL.GetValue('GL_PUHT') - TOBL.GetValue('GL_DPR')) / TOBL.GetValue('GL_PUHT'),4) * 100;
  end else
  begin
    PP := 0;
  end;
  TOBL.PutValue('POURCENTMARQ',PP);
end;

{============================ CreerPiecesFromLignes ========================}
Function  CreerPiecesFromLignes(TOBLigne:TOB; GenereMode:String; DatePriseEncompte : TDateTime;CompteRendu:boolean=true;WithLiaisonDoc : boolean=false; TOBResult : TOB=nil; TOBDESOUVRAGES : TOB=nil):Boolean;
var TOBSource             : TOB;
    NewTOBP               : TOB;
    NewTOBL               : TOB;
    TOBL                  : TOB;
    QteReliquat           : double;
    MtReliquat            : double;

    StRendu               : string;
    TrierPar              : string;
    RuptureCol            : string;
    ValPrec               : String;
    NaturePiece,CodeTiers,CodeAffaire,Etablissement,Domaine:String;
    DateLigne,DatePiece   :TDateTime;
    i,iLigne,iPiece       :Integer;
    {Fin, }AvecComment,MajPrix:Boolean;
    io                    :TIoErr ;
    GenPiece              : T_GenereLesPieces;
    RatioStk              : Double;
    Qte                   : Double;
    Mt                    : Double;
    ratioVA               : Double;
    RatioVente            :double;
    NumPiece              : integer;
    // Modif BTP
    UpdateDev             : boolean;
    Remise,Escompte       : double;
    GestionReliquat       : boolean;
    OldTypePiece          : string;
{$IFDEF BTP}
    MsgErreur		          : String;
    ListeAppel	          : String;
		Intervention          : boolean;
    AjoutAdresseDocument  : boolean;
    DocumentInit          : string;
    DocumentInitPrec      : string;
    ModeLivraison         : integer;
    MemoTemp              : TStrings;
    QteLivrable           : double;
    NBLigneArt            : integer;
    ValPu                 : double;
    CoefPaPr              : double;
    LastAffaire           : string;
    LesTOBOuvrages        : TOB;
    LesTOBSSTRAIT         : TOB;
    Ok_Reliquat           : Boolean;
    MtLivrable            : Double;
{$ENDIF}
    // --

   Function InitMode:Boolean;
   BEGIN
      GestionReliquat := false;
{$IFDEF BTP}
      AjoutAdresseDocument := false;
      DocumentInitPrec := '';
      ModeLivraison := -2;
      NBLigneArt := 0;
{$ENDIF}
      NumPiece := -1;
      ForceNumero := false;
      Remise := 0; Escompte := 0;
      GenereMode := UpperCase(GenereMode);
      bDuplicPiece := True;
      if (GenereMode = 'FACDETAILCONSO') or (GenereMode = 'FACCONSO') then
     	   Begin
         OldTypePiece := 'CON';
         // ContreMarque achat
         if Getparamsoc('SO_FACTPROV') then
            NaturePiece:='FPR'
         else
            NaturePiece:='FAC';
         TrierPar:='GL_AFFAIRE;GL_TIERS;BLP_NUMMOUV;';
         RuptureCol:= 'GL_AFFAIRE;';
         if (GenereMode = 'FACCONSO') then
          	begin
            TrierPar:='GL_TIERS;GL_CHANTIER;BLP_NUMMOUV';
	          RuptureCol:= 'GL_TIERS;GL_CHANTIER';
        	  end;
       	AjoutAdresseDocument := false;
        AvecComment:=False;
        DatePiece:=0;
        MajPrix:=True;
        UpdateDev := false;
        Result:=True;
        end
      else if (GenereMode = 'VTOA') or (GenereMode = 'REAP') or (GenereMode = 'REAPECLAT') or
      		(GenereMode = 'VALIDDECIS') or (GenereMode = 'VALIDDECISECLAT')  then
      BEGIN
      OldTypePiece := 'CC';
      // ContreMarque achat
      NaturePiece:='CF';
      TrierPar:='GL_TIERS;';
      RuptureCol:='GL_TIERS;';
      if (GenereMode = 'REAP') or (GenereMode = 'REAPECLAT') OR
      	 (GenereMode = 'VALIDDECIS') or (GenereMode = 'VALIDDECISECLAT') then
      begin
        OldTypePiece := 'REA';
{$IFDEF BTP}
        AjoutAdresseDocument := true;
        RuptureCol := RuptureCol + 'DOCUMENT;';
        // TrierPar:=TrierPar+'GL_DEPOT;ORIGINE;GL_IDENTIFIANTWOL;';
//        if GenereMode = 'REAPECLAT' then RuptureCol:=RuptureCol+'GL_DEPOT;GL_AFFAIRE;GL_IDENTIFIANTWOL;';
//        if GenereMode = 'REAPECLAT' then RuptureCol:=RuptureCol+'GL_DEPOT;GL_AFFAIRE;GL_IDENTIFIANTWOL;';
{$ENDIF}
      end;
      if GetInfoParPiece(NaturePiece,'GPP_ECLATEDOMAINE')='X' then
         BEGIN
         TrierPar:=TrierPar+'GL_DOMAINE;';
         RuptureCol:=RuptureCol+'GL_DOMAINE;'; // DBR - point virgule ajouté si rup sur affaire aussi - fiche 10164
         END;
      TrierPar:=TrierPar+'GL_DATELIVRAISON;';
      if ctxgcaff in V_PGI.PGIContexte then // DBR - fiche 10164
      begin
        if GetInfoParPiece(NaturePiece,'GPP_ECLATEAFFAIRE')='X' then
          RuptureCol:=RuptureCol+'GL_AFFAIRE;';
      end;
{$IFDEF BTP}
      if (GenereMode = 'REAP') or (GenereMode = 'REAPECLAT') or
      	 (GenereMode = 'VALIDDECIS') or (GenereMode = 'VALIDDECISECLAT') then
      begin
        TrierPar := ''; // on ne trie plus car arrive deja préparé
      end;
{$ENDIF}
      AvecComment:=False;
      DatePiece:=0;
      MajPrix:=True;
      UpdateDev := false;
      Result:=True;
      END else if GenereMode = 'VTOBL' then
      BEGIN
      // ContreMarque livraison
      NaturePiece:='BLC';
      TrierPar:='GL_TIERS;GL_DEPOT;';
      RuptureCol:='GL_TIERS;GL_DEPOT;';
      if GetInfoParPiece(NaturePiece,'GPP_ECLATEDOMAINE')='X' then
         BEGIN
         TrierPar:=TrierPar+'GL_DOMAINE;';
         RuptureCol:=RuptureCol+'GL_DOMAINE';
         END;
      TrierPar:=TrierPar+'GL_DATELIVRAISON;';
      AvecComment:=False;
      DatePiece:=0;
      MajPrix:=False;
      UpdateDev := false;
      Result:=True;
{$IFDEF BTP}
      END else if GenereMode = 'RFOTOLIVC' then   // generation des livraisons depuis les receptions
      BEGIN
      	//
        OldTypePiece := 'BLF';
        NaturePiece:='LBT';
        TrierPar:='GL_TIERS;GL_AFFAIRE;GL_NUMLIGNE;';
        RuptureCol:='GL_TIERS;GL_AFFAIRE;';
        AvecComment:=false;
        DatePiece:=0;
        MajPrix:=true;
        UpdateDev := true;
        Result:=True;
        //
      END else if GenereMode = 'DEVTOCHAN' then
      BEGIN
      // Prévision de chantier
      OldTypePiece := 'DBT';
      NaturePiece:=GetParamSoc('SO_BTNATCHANTIER');
      TrierPar:='';
      RuptureCol:='GL_TIERS;GL_DEPOT;';
      ForceNumero := false;
      AvecComment:=False;
      DatePiece:=0;
      MajPrix:=False;
      UpdateDev := true;
      Result:=True;
      END else if GenereMode = 'CHANTOCBT' then
      BEGIN
      // Besoin de chantier
      OldTypePiece := GetParamSoc('SO_BTNATCHANTIER');
      NaturePiece:='CBT';
      TrierPar:='GL_TIERS;ORIGINE;GL_NUMLIGNE;';
      RuptureCol:='GL_TIERS;ORIGINE';
      if GetInfoParPiece(NaturePiece,'GPP_ECLATEDOMAINE')='X' then
         BEGIN
         TrierPar:=TrierPar+'GL_DOMAINE;';
         RuptureCol:=RuptureCol+'GL_DOMAINE';
         END;
      TrierPar:=TrierPar+'GL_DATELIVRAISON;';
      AvecComment:=False;
      DatePiece:=0;
      MajPrix:=False;
      UpdateDev := false;
      Result:=True;
      END else if GenereMode ='CCTOLIVCLI' then
      begin
        // Besoin de chantier
        OldTypePiece := 'CC';
        GestionReliquat := true;
  //      NaturePiece:='BLC';
        NaturePiece:='BLC';
        TrierPar:='GL_TIERS;ORIGINE;GL_NUMLIGNE;';
        RuptureCol:='GL_TIERS;ORIGINE';
        if GetInfoParPiece(NaturePiece,'GPP_ECLATEDOMAINE')='X' then
           BEGIN
           TrierPar:=TrierPar+'GL_DOMAINE;';
           RuptureCol:=RuptureCol+'GL_DOMAINE';
           END;
        AvecComment:=False;
        DatePiece:=0;
        MajPrix:=False;
        UpdateDev := false;
        Result:=True;
      END else if GenereMode = 'CBTTOLIVC' then
      BEGIN
      // Besoin de chantier
      OldTypePiece := 'CBT';
      GestionReliquat := true;
//      NaturePiece:='BLC';
      NaturePiece:='LBT';
      TrierPar:='GL_TIERS;ORIGINE;GL_NUMLIGNE;';
      RuptureCol:='GL_TIERS;ORIGINE';
      if GetInfoParPiece(NaturePiece,'GPP_ECLATEDOMAINE')='X' then
         BEGIN
         TrierPar:=TrierPar+'GL_DOMAINE;';
         RuptureCol:=RuptureCol+'GL_DOMAINE';
         END;
      AvecComment:=False;
      DatePiece:=0;
      MajPrix:=False;
      UpdateDev := false;
      Result:=True;
{$ENDIF}
      END else if GenereMode = 'DEVTOCTE' then {Devis vers Contre etude}
      BEGIN
      OldTypePiece := 'DBT';
      NaturePiece:='BCE';
      TrierPar:='';
      RuptureCol:='GL_TIERS;';
      AvecComment:=false;
      DatePiece:=0;
      MajPrix:=False;
      UpdateDev := false;
      Result:=True;
      END else
         Result:=False;
   END;

   Procedure TrierTOB(ParQuoi:String);
   BEGIN
      if ParQuoi <> '' then TOBLigne.Detail.Sort(ParQuoi);
   END;

   Function GetColRupture(TOBL:TOB):String;
   var sChamp,sRupt:string;
       Tobresult : variant;
   BEGIN
   result:='';
   sRupt:=RuptureCol;
   sChamp:=ReadTokenSt(sRupt);
   While sChamp<>'' do
      BEGIN
      if Result<>'' then Result := Result+';';
      TOBResult := TOBL.GetValue(sChamp);
      if (VarType(TOBresult)) and (varTypeMask) = VarInteger then
      begin
        TobResult := VarToStr (ToBresult);
      end;
      Result := Result+TOBResult;
      sChamp:=ReadTokenSt(sRupt);
      END;
   END;

   Function RecupNumPiece(TOBL : TOB) : integer;
   begin
      if (GenereMode = 'DEVTOCHAN') then
         BEGIN
         Result:=TOBL.GetValue('GL_NUMERO');
         END else Result := -1;
   end;

   Function RecupCodeTiers(TOBL:TOB):String;
   BEGIN
      if (GenereMode = 'VTOA') or
         (GenereMode = 'REAP') or
         (GenereMode = 'REAPECLAT') or
				 (GenereMode = 'VALIDDECIS') or
         (GenereMode = 'VALIDDECISECLAT') OR
         (GenereMode = 'DEVTOCHAN') or
         (GenereMode = 'CHANTOCBT') or
         (GenereMode = 'CBTTOLIVC') or
         (GenereMode = 'CCTOLIVCLI') or
         (GenereMode = 'RFOTOLIVC') or
         (UpperCase (GenereMode) = 'DEVTOCTE') or
         (UpperCase (GenereMode) = 'FACCONSO') or
         (UpperCase (GenereMode) = 'FACDETAILCONSO') then
         Result:=TOBL.GetValue('GL_TIERS')
      else
         Result := '';
   END;

   Function RecupCodeAffaire(TOBL:TOB):String;
   BEGIN
      if (GenereMode = 'VTOA') or
         (GenereMode = 'REAP') or
         (GenereMode = 'REAPECLAT') or
				 (GenereMode = 'VALIDDECIS') or
         (GenereMode = 'VALIDDECISECLAT') or
         (GenereMode = 'DEVTOCHAN') or
         (GenereMode = 'CHANTOCBT') or
         (GenereMode = 'CBTTOLIVC') or
         (GenereMode = 'CCTOLIVCLI') or
         (GenereMode = 'RFOTOLIVC') or
         (UpperCase (GenereMode) = 'DEVTOCTE') or
         (UpperCase (GenereMode) = 'FACCONSO') or
         (UpperCase (GenereMode) = 'FACDETAILCONSO') then
       begin
           if UpperCase (GenereMode) = 'FACCONSO' then
           begin
              Result:=TOBL.GetValue('GL_CHANTIER');
              if Result='' then PgiInfo('Ligne de l''appel '+TOBL.GetValue('GL_AFFAIRE')+' non traitée.'+chr(13)+' Chantier associé non renseigné dans l''appel.','FACTURATION IMPOSSIBLE');
           end
           else
           begin
              Result:=TOBL.GetValue('GL_AFFAIRE');
           end;
       end
       else
       begin
         Result := '';
       end;
   END;

   Function RecupRemise(TOBL:TOB):double;
   BEGIN
      if (GenereMode = 'DEVTOCHAN') or
      	 (UpperCase (GenereMode) = 'DEVTOCTE') then
         BEGIN
         Result:=TOBL.GetValue('GL_REMISEPIED');
         END else Result := 0;
   END;

   Function RecupEscompte(TOBL:TOB):double;
   BEGIN
      if (GenereMode = 'DEVTOCHAN') or
      	 (UpperCase (GenereMode) = 'DEVTOCTE') then
         BEGIN
         Result:=TOBL.GetValue('GL_ESCOMPTE');
         END else Result := 0;
   END;

BEGIN
	//
//  MemoriseChampsSupLigneETL (NaturePiece ,True);
  MemoriseChampsSupLigneETL (NewNature ,True);
  MemoriseChampsSupLigneOUV (NewNature);
  MemoriseChampsSupPIECETRAIT;
  //
	LastAffaire := '';
	EclateParPiece := false;
  result:=true;
  if (TOBLigne = nil)  then EXIT;
  if (TOBLigne.Detail.Count = 0) then EXIT;
  if not InitMode then EXIT;
  if TrierPar <> '' then TrierTOB(TrierPar);
  TOBSource := TOB.Create('LES PIECES', nil, -1);
  LesTOBOuvrages := TOB.Create ('LES OUVRAGES DES PIECES',nil,-1);
  LesTOBSSTRAIT := TOB.Create ('LES SS-TRAITS',nil,-1);
  // Modif BTP
  AddLesSupEntete (TOBSource);
{$IFDEF BTP}
  GestionPhase := TGestionphase.create;
{$ENDIF}
  // ---
  GenereMode := UpperCase(GenereMode);
  bContremarque:=True;
  // Modif MODE
  CodeAff:=''; LiaisonOrli:=False;
  // -----
  {$IFDEF BTP}
  if (GenereMode = 'DEVTOCHAN') or (GenereMode = 'CHANTOCBT') then bContremarque := false;
  {$ENDIF}
  {Boucle de génération des NewTOBPiece et NewTOBLigne}
//  Fin:=False;
  {i:=0; }iLigne:=0; iPiece:=0;
  ValPrec := '';
  CoefPapr := 1;
{$IFDEF BTP}
  if (GenereMode = 'CBTTOLIVC') or (GenereMode = 'CCTOLIVCLI') then
  begin
  	InitCalcDispo;
  end;
{$ENDIF}
  NewTOBP := nil;
  MsgErreur := '';
  ListeAppel:= '';

  if (UpperCase (GenereMode) = 'FACCONSO') or
   (UpperCase (GenereMode) = 'FACDETAILCONSO') then
  begin
     //Modification pour trier les appels par numéro d'affaire.
     TOBLigne.Detail.Sort('GL_AFFAIRE;BLP_NUMMOUV;');
  end;

  for i:= 0 to TOBLigne.detail.count -1 do
  begin
    TOBL:=TOBLigne.Detail[i];
    Ok_Reliquat := CtrlOkReliquat(TOBL, 'GL');
{$IFDEF BTP}
    if (GenereMode = 'CHANTOCBT') then
    begin
      if (TOBL.getValue('GL_TYPEARTICLE')='PRE') and (IsPrestationInterne(TOBL.GetValue('GL_ARTICLE'))) Then continue;
      // correction Fiche qualitée 11092
      if (TOBL.getValue('GL_TYPEARTICLE')='FRA')  Then continue;
    end;
		if (GenereMode = 'CBTTOLIVC') or (GenereMode = 'CCTOLIVCLI') then
    begin
      QteLivrable := GetQtelivrable (TOBL,GestionPhase,nil,SansStock); // juste pour avoir les PA affectés
      QteLivrable := TOBL.GetValue('GL_QTEFACT');
//      if Qtelivrable = 0 then continue; // si rien a livrer sur cette ligne --> ligne suivante
    end
    else
    begin
      QteLivrable := TOBL.GetValue('GL_QTERESTE');
      if Ok_Reliquat then MtLivrable := TOBL.GetValue('GL_MTRESTE');
    end;


{$ENDIF}
    if ((GenereMode = 'REAP')              or
        (GenereMode = 'REAPECLAT')         or
    	  (GenereMode = 'VALIDDECIS')        or
        (GenereMode = 'VALIDDECISECLAT')) and (TOBL.GetValue('GL_TIERS')='')     then continue;
    //
    if (GetColRupture(TOBL) <> ValPrec) then
    begin
      // Creation de la nouvelle TOBPiece
      CodeTiers := RecupCodeTiers(TOBL);
      if GetInfoParPiece(NaturePiece,'GPP_ECLATEAFFAIRE')='X' then // DBR - Fiche 10164
         CodeAffaire := RecupCodeAffaire(TOBL)
      else
         CodeAffaire := '';
      Remise := 0;
      Escompte := 0;
      {$IFDEF BTP}
    	if genereMode = 'FACCONSO' then
        begin
       	CodeAffaire := RecupCodeAffaire(TOBL);
        LastAffaire := TOBL.GetValue('GL_AFFAIRE');
        if CodeAffaire = '' then
           Begin
           if ListeAppel = '' then
              ListeAppel := TOBL.GetValue('GL_AFFAIRE')
           else
              if pos(TOBL.GetValue('GL_AFFAIRE'), ListeAppel) = 0 then
                 ListeAppel := ListeAppel + ';' + TOBL.GetValue('GL_AFFAIRE');
           Continue;
           end;
        end
      else
        begin
      	CodeAffaire := RecupCodeAffaire(TOBL);
      	LastAffaire := CodeAffaire;
        end;
      if GenereMode = 'DEVTOCHAN' then
        begin
      	// c'est finit maintenant on cree une piece distincte avec son propre compteur
(*      NumPiece := RecupNumPiece(TOBL); *)
//      Remise := RecupRemise(TOBL);
//      Escompte := RecupEscompte (TOBL);
        end;
      Etablissement := TOBL.GetValue('GL_ETABLISSEMENT');
      {$ENDIF}
      NewTOBP := EcritTOBPiece(NaturePiece, CodeTiers, CodeAffaire, Etablissement, Domaine, TOBL,Numpiece,Remise,Escompte,iPiece,DatePriseEnCompte);
      if NewTobP = nil then
      BEGIN
        TOBSource.free ;
        Exit;
      END;
      NewTOBP.ChangeParent(TOBSource,-1);
      // NEW ONE (OPTIMISATION CALCUL)
      ZeroFacture (NewTOBP);
      // --
      if TOBDESOUVRAGES <> nil then
      begin
      	// repositionne la tob des ouvrages pour les documents traites
        G_AjouteOuvragesPiece (LesTOBOuvrages,TOBDESOUVRAGES,TOBL,NewTOBP);
      end;

      DatePiece:=NewTOBP.GetValue('GP_DATELIVRAISON');
{$IFDEF BTP}
      if (GenereMode = 'CHANTOCBT') or (GenereMode = 'CBTTOLIVC') or (Generemode = 'DEVTOCHAN') or (GenereMode ='CCTOLIVCLI') then
      begin
        if (GenereMode = 'CHANTOCBT') or
    	     (GenereMode = 'VALIDDECIS') or
           (GenereMode = 'VALIDDECISECLAT')  then
        begin
          NewTOBP.PutValue('GP_DATELIVRAISON',TOBL.GetValue('GL_DATELIVRAISON'));
          DatePiece:=NewTOBP.GetValue('GP_DATELIVRAISON');
        end;
        // récupération du mode de livraison --> Suivi dans les pièces suivantes
        NewTOBP.PutValue('GP_IDENTIFIANTWOT',TOBL.GetValue('GL_IDENTIFIANTWOL'));
        if TOBL.FieldExists ('GP_REFINTERNE') then NewTOBP.PutValue('GP_REFINTERNE',TOBL.GetValue('GP_REFINTERNE'));
        if TOBL.FieldExists ('GP_REPRESENTANT') then NewTOBP.PutValue('GP_REPRESENTANT',TOBL.GetValue('GP_REPRESENTANT'));
       	// Dans le cas ou le client est TTC on enregistre quand meme en HT
        NewTOBP.PutValue('GP_FACTUREHT','X');
        // --
        if (GenereMode = 'DEVTOCHAN') then
        begin
          NewTOBP.addChampSupValeur('PIECEORIGINE',TOBL.GetValue('PIECEORIGINE'),false);
          NewTOBP.PutValue('GP_NUMADRESSEFACT',TOBL.GetValue('GP_NUMADRESSEFACT'));
          NewTOBP.PutValue('GP_NUMADRESSELIVR',TOBL.GetValue('GP_NUMADRESSELIVR'));
        end;
      end;
      if (GenereMode = 'CBTTOLIVC') or (GenereMode = 'RFOTOLIVC') or (GenereMode ='CCTOLIVCLI') then
      begin
      	NewTOBP.PutValue('GP_DATEPIECE',DatePriseEncompte);
      	NewTOBP.PutValue('GP_DATELIVRAISON',DatePriseEncompte);
      end;
      if (GenereMode = 'CHANTOCBT') or (GenereMode = 'CBTTOLIVC') or (GenereMode='CCTOLIVCLI') then
      begin
        // ceci afin de récupérer les adresses -- dans ce cas relation 1-1 entre document
        NewTOBP.addChampSupValeur('PIECEORIGINE',TOBL.GetValue('GL_PIECEORIGINE'),false);
        NewTOBP.PutValue('GP_NUMADRESSEFACT',TOBL.GetValue('GP_NUMADRESSEFACT'));
        NewTOBP.PutValue('GP_NUMADRESSELIVR',TOBL.GetValue('GP_NUMADRESSELIVR'));
        if TOBL.FieldExists ('GP_REFINTERNE') then NewTOBP.PutValue('GP_REFINTERNE',TOBL.GetValue('GP_REFINTERNE'));
      end;
      if (GenereMode = 'DEVTOCTE') then
      begin
        // ceci afin de récupérer les adresses -- dans ce cas relation 1-1 entre document
        NewTOBP.addChampSupValeur('PIECEORIGINE',TOBL.GetValue('PIECEORIGINE'),false);
        NewTOBP.PutValue('GP_NUMADRESSEFACT',TOBL.GetValue('GP_NUMADRESSEFACT'));
        NewTOBP.PutValue('GP_NUMADRESSELIVR',TOBL.GetValue('GP_NUMADRESSELIVR'));
        NewTOBP.PutValue('GP_PIECEENPA',GetInfoParPiece(NewNature, 'GPP_FORCEENPA'));
        //
        if TOBL.FieldExists ('GP_DOMAINE')    then NewTOBP.SetString('GP_DOMAINE',TOBL.GetValue('GP_DOMAINE'));
        // --
        if TOBL.FieldExists ('GP_PIECEFRAIS') then NewTOBP.addChampSupValeur('GP_PIECEFRAIS',TOBL.GetValue('GP_PIECEFRAIS'),false);
        if TOBL.FieldExists ('GP_REFINTERNE') then NewTOBP.PutValue('GP_REFINTERNE',TOBL.GetValue('GP_REFINTERNE'));
        if TOBL.FieldExists ('GP_APPLICFGST') then NewTOBP.PutValue('GP_APPLICFGST',TOBL.GetValue('GP_APPLICFGST'));
        if TOBL.FieldExists ('GP_APPLICFCST') then NewTOBP.PutValue('GP_APPLICFCST',TOBL.GetValue('GP_APPLICFCST'));
        NewTOBP.PutValue('GP_COEFFC',TOBLigne.GetValue('GP_COEFFC'));
        NewTOBP.PutValue('GP_COEFFR',TOBLigne.GetValue('GP_COEFFR'));
      end;
      if NewTOBP.GetValue('GP_VENTEACHAT')='VEN' Then
      begin
      	NewTOBP.PutValue('GP_PIECEENPA',GetInfoParPiece(NaturePiece, 'GPP_FORCEENPA'));
      end;
{$ENDIF}
      ValPrec := GetColRupture(TOBL);
      Inc(iPiece);
      iLigne:=0;
{$IFDEF BTP}
      DocumentInitPrec := '';
      ModeLivraison := TOBL.GetValue('GL_IDENTIFIANTWOL');
{$ENDIF}
      if (GenereMode = 'DEVTOCHAN') or (GenereMode = 'RFOTOLIVC') or
      	 (GenereMode = 'CBTTOLIVC') or (GenereMode = 'FACCONSO') or (genereMode ='CCTOLIVCLI') or
         (GenereMode = 'FACDETAILCONSO') or (GenereMode = 'DEVTOCTE')  then
         begin
           if (GenereMode = 'FACDETAILCONSO') or
              (genereMode = 'FACCONSO') then
              Intervention := true
           else
              Intervention := false;
           if TOBL.GetString('NUMBON')<> '' then
           begin
            G_LigneCommentaireNumBon (newTOBP,TOBL,true,True);
           end else
           begin
            G_LigneCommentaireBTP (newTOBP,TOBL,true,Intervention);
           end;
         end;
    end;

{$IFDEF BTP}
		if (GenereMode = 'FACCONSO') and (TOBL.GetValue('GL_AFFAIRE')<>LastAffaire) then
    begin
      if TOBL.GetString('NUMBON')<> '' then
      begin
    	  G_LigneCommentaireNumBon (newTOBP,TOBL,true,True);
      end else
      begin
    	  G_LigneCommentaireBTP (newTOBP,TOBL,true,True);
      end;
      LastAffaire := TOBL.GetValue('GL_AFFAIRE');
    end;

    DocumentInit := RecupPieceInit(TOBL);
    if (GenereMode = 'VALIDDECISECLAT') or (GenereMode = 'REAPECLAT') then
    begin
      if (not IsSameDocument(DocumentInit,DocumentInitPrec) or
         (ModeLivraison <> TOBL.GetValue('GL_IDENTIFIANTWOL')) ) and (AjoutAdresseDocument) and (Not IsDetailBesoin(TOBL)) then
      begin
        MemoTemp := TStringList.Create;
        TRY
          Inc(iLigne);
          NewTOBL := NewTOBLigne(newTOBP,0); 
          (*
          NewTOBL := TOB.Create('LIGNE', NewTOBP, -1);
          AddLesSupLigne(newTOBl,False);
          *)
          InitLigneComm (TOBL,NewTobP,NewTOBL,iLigne);
          NewTOBL.PutValue('GL_IDENTIFIANTWOL',TOBL.GetValue('GL_IDENTIFIANTWOL'));
          AffecteAdresseLiv(TOBL,MemoTemp); // recherche adresse de livraison du document initial
          NewTOBL.putvalue('GL_BLOCNOTE',MemoTemp.text);
        FINALLY
          Memotemp.free;
          DocumentInitPrec := DocumentInit;
          ModeLivraison := TOBL.GetValue('GL_IDENTIFIANTWOL');
        END;
      end;
    end;
    if (GenereMode = 'VALIDDECIS') or (GenereMode = 'VALIDDECISECLAT')  or
    	 (GenereMode = 'REAP') or (GenereMode = 'REAPECLAT')  then
    begin
      NewTOBP.PutValue('GP_DATELIVRAISON',TOBL.GetValue('GL_DATELIVRAISON'));
      DatePiece:=NewTOBP.GetValue('GP_DATELIVRAISON');
    end;
    if (Pos(genereMode,'DEVTOCHAN;CHANTOCBT')>0) then
    begin
      // Dans le cas de génération de contre etude ou de génération de prévision de chantier
      if TOBL.GetString('GLC_NATURETRAVAIL')='002' then
      begin
        UG_recupSSTrait (TOBL,LesTOBSSTRAIT);
      end;
    end;

{$ENDIF}
    {Génére TOBligne fille}
    Inc(iLigne);
    NewTOBL := TOB.Create('LIGNE', NewTOBP, -1);
{$IFDEF BTP}
		AddLesSupLigne (NewTOBL,False);
    //		AddLesSupLignesBtp (NewTOBL,false);
{$ENDIF}
    EcritTOBLigne(NewTOBL, TOBL, NewTOBP, iLigne);
{$IFDEF BTP}
    NewTOBL.putValue('GL_IDENTIFIANTWOL',TOBL.GetValue('GL_IDENTIFIANTWOL')); // ??
    if (Generemode = 'DEVTOCHAN') or (Generemode = 'DEVTOCTE') then
    begin
    	if NewTOBL.fieldExists ('GLC_DOCUMENTLIE') then
      begin
      	if TOBL.GetValue('GLC_DOCUMENTLIE')='' then
        begin
    			NewTOBL.putValue('GLC_DOCUMENTLIE',EncodeRefPiece (TOBL,0,true)); // ??
        end else
        begin
    			NewTOBL.putValue('GLC_DOCUMENTLIE',TOBL.GetValue('GLC_DOCUMENTLIE'));
        end;
      end;
      if (Generemode = 'DEVTOCTE') then
      begin
      	NewTOBL.PutValue('GL_PIECEORIGINE', NewTOBP.GetValue('PIECEORIGINE'));
      end;
      NewTOBL.Data := TOBL.Data;
    end;
    if (UpperCase (GenereMode) = 'FACCONSO') or
       (UpperCase (GenereMode) = 'FACDETAILCONSO') then
       begin
       NewTOBL.putValue('GL_INDICENOMEN', 0);
       if UpperCase (GenereMode) = 'FACCONSO' then
          NEWTOBL.PutValue('GLC_AFFAIRELIEE', TOBL.GETVALUE('GL_AFFAIRE'));
       end;
{$ENDIF}
    if GetInfoParPiece (NaturePiece, 'GPP_ECLATEAFFAIRE') = '-' then // DBR - fiche 10164
    begin
      NewTobl.PutValue ('GL_AFFAIRE', TobL.GetValue ('GL_AFFAIRE'));
    end;
    DateLigne:=NewTOBL.GetValue('GL_DATELIVRAISON');
    if Dateligne<V_PGI.DateEntree then
    begin
      DateLigne:=V_PGI.DateEntree;
      NewTOBL.PutValue('GL_DATELIVRAISON',DateLigne);
    end;

    if (GenereMode <> 'CBTTOLIVC') and (GenereMode <> 'RFOTOLIVC') and (GenereMode <> 'CCTOLIVCLI') then
    begin
      if (iLigne=1) or (DateLigne>DatePiece) then
      Begin
        NewTOBP.PutValue('GP_DATELIVRAISON',DateLigne) ;
        DatePiece:=DateLigne;
      End;
    end;

{$IFDEF BTP}
    if (UpperCase(GenereMode) = 'RFOTOLIVC')  then
    begin
    // là on a forcement les lignes de réceptions a livrer sur chantiers sans Besoin préalable
//      RatioStk:=GetRatio(TOBL,Nil,trsDocument);
      //
      RatioVente := GetRatio(TOBL, nil, trsVente);
      //
      RatioStk := Ratiovente ;
      Qte:=ARRONDI(NewTOBL.GetValue('GL_QTEFACT') * RatioStk,V_PGI.OkDecQ);

      if (NewTobl.getValue('GL_TYPELIGNE')='ART') and (Qte <> 0 ) Then inc(NBLigneArt);
      //
      NewTOBL.PutValue('GL_QTESTOCK', Qte);
      NewTOBL.PutValue('GL_QTEFACT',  Qte);
      NewTOBL.PutValue('GL_QTERELIQUAT',0);
      NewTOBL.PutValue('GL_REMISEPIED', 0);
      NewTOBL.PutValue('GL_REMISELIGNE',0);
      NewTOBL.PutValue('GL_QTERESTE', Qte);  { NEWPIECE }
      // --- GUINIER ---
      if  Ok_Reliquat then
      begin
        Mt := ARRONDI(NewTOBL.GetValue('GL_MontantHTDEV'),V_PGI.OkDecP);
        NewTOBL.PutValue('GL_MTRELIQUAT',0);
        NewTOBL.PutValue('GL_MTRESTE', Mt); { NEWPIECE }
      end;
      //
      NewTOBL.PutValue('GL_TYPEDIM','NOR');
      // Recalcule des PVHTDEV et PVHT en reappliquant les coef fg et marge de l'article
      PositionneValeurPV (newTOBL);
      //
    end
    else
    if (UpperCase(GenereMode) = 'CBTTOLIVC')  or (UpperCase(GenereMode) = 'CCTOLIVCLI') then
    begin
      QteReliquat := Abs(TOBL.getValue('GL_QTERESTE'))-Abs(QTeLivrable);
      if QTereliquat <= 0 then NewTOBL.PutValue('GL_QTERELIQUAT',0)
                          else NewTOBL.PutValue('GL_QTERELIQUAT',TOBL.getValue('GL_QTERESTE')-QTeLivrable);
      NewTOBL.PutValue('GL_QTESTOCK', QteLivrable);
      NewTOBL.PutValue('GL_QTEFACT',  QteLivrable);
      NewTOBL.PutValue('GL_QTERESTE', QteLivrable);  { NEWPIECE }
      // --- GUINIER ----
      if Ok_Reliquat then
      begin
        MtReliquat := Abs(TOBL.getValue('GL_MONTANTHTDEV'))-Abs(MtLivrable);
        if MtReliquat <= 0 then
          NewTOBL.PutValue('GL_MTRELIQUAT', 0)
        else
          NewTOBL.PutValue('GL_MTRELIQUAT', TOBL.getValue('GL_MONTANTHTDEV') - MtLivrable);
        NewTOBL.PutValue('GL_MTRESTE',     MtLivrable);  { NEWPIECE }
      end;
      //
      NewTOBL.PutValue('GL_TYPEDIM','NOR');
      //
      if (NewTobl.getValue('GL_TYPELIGNE')='ART') and (QteLivrable <> 0 ) Then inc(NBLigneArt);
      //
      if (TOBL.FieldExists ('DPARECUPFROMRECEP')) and (TOBL.GetValue ('DPARECUPFROMRECEP') > 0) then
      begin
        if TOBL.GetValue('GL_DPA') <> 0 then CoefPaPR := TOBL.GetValue('GL_DPR')/TOBL.GetValue('GL_DPA');
        if CoefPaPr = 0 then CoefPaPr := 1;
      	NewTOBL.PutValue('GL_DPA', TOBL.GetValue ('DPARECUPFROMRECEP'));
        NewTOBL.PutValue('GL_DPR', Arrondi(NewTOBL.GEtValue('GL_DPA')*CoefPaPr,V_PGI.okdecP));
      end;
      PositionneValeurPV (newTOBL,false);
    end else
    if (UpperCase (GenereMode) = 'DEVTOCHAN') or
       (UpperCase (GenereMode) = 'CHANTOCBT') or
       (UpperCase (GenereMode) = 'DEVTOCTE') or
       (UpperCase (GenereMode) = 'CBTTOLIV') or
       (UpperCase (GenereMode) = 'VALIDDECIS') or
       (UpperCase (GenereMode) = 'VALIDDECISECLAT') or
       (UpperCase (GenereMode) = 'FACCONSO') or
       (UpperCase (GenereMode) = 'FACDETAILCONSO') then
    begin
      //Maj des qté
      //Converti de l'unité de stock à l'unité du type de ligne
      if (UpperCase (GenereMode) = 'CHANTOCBT') then
        Qte:=NewTOBL.GetValue('GL_QTERESTE')
      else
      	Qte:=NewTOBL.GetValue('GL_QTEFACT');
      NewTOBL.PutValue('GL_QTESTOCK',Qte);
      NewTOBL.PutValue('GL_QTEFACT', Qte);
      NewTOBL.PutValue('GL_QTERELIQUAT',0);
      NewTOBL.PutValue('GL_QTERESTE', Qte);  { NEWPIECE }
      // --- GUINIER ---
      if Ok_Reliquat then
      begin
        if (UpperCase (GenereMode) = 'CHANTOCBT') then
          Mt := NewTOBL.GetValue('GL_MTRESTE')
        else
  	      Mt := NewTOBL.GetValue('GL_MONTANTHTDEV');
        NewTOBL.PutValue('GL_MTRELIQUAT',0);
        NewTOBL.PutValue('GL_MTRESTE', Mt);  { NEWPIECE }
      end;

      NewTOBL.PutValue('GL_TYPEDIM','NOR');
      //
      if (NewTobl.getValue('GL_TYPELIGNE')='ART') and (Qte <> 0 ) Then inc(NBLigneArt);
      //
    end else
    begin
{$ENDIF}
      RatioStk:=GetRatio(NewTOBL,Nil,trsStock);
      Qte:=NewTOBL.GetValue('GL_QTESTOCK')*RatioStk;
      NewTOBL.PutValue('GL_QTEFACT',Qte);
      NewTOBL.PutValue('GL_QTESTOCK',Qte);
      NewTOBL.PutValue('GL_QTERELIQUAT',0);
      NewTOBL.PutValue('GL_QTERESTE', Qte);  { NEWPIECE }
      // --- GUINIER ---
      if  Ok_Reliquat then
      begin
        Mt :=NewTOBL.GetValue('GL_MONTANTHTDEV');
        NewTOBL.PutValue('GL_MTRELIQUAT',0);
        NewTOBL.PutValue('GL_MTRESTE', Mt);  { NEWPIECE }
      end;

      NewTOBL.PutValue('GL_TYPEDIM','NOR');
      //
      if (NewTobl.getValue('GL_TYPELIGNE')='ART') and (Qte <> 0 ) Then inc(NBLigneArt);
      //
{$IFDEF BTP}
      if (UpperCase (GenereMode) = 'REAP') or
         (UpperCase (GenereMode) = 'REAPECLAT') then
      begin
        ValPU := NewTOBL.GetValue('GL_PUHT')/ RatioStk;
        NewTOBL.PutValue('GL_PUHT',ValPu);
        NewTobl.putvalue('GL_PUHTDEV',ValPu);
      end;
    end;
    if (GenereMode = 'CBTTOLIVC') or (GenereMode = 'RFOTOLIVC') or (GenereMode ='CCTOLIVCLI') then
    begin
      NewTOBL.PutValue('GL_DATEPIECE',NewTOBP.GetValue('GP_DATEPIECE'));
      NewTOBL.PutValue('GL_DATELIVRAISON',NewTOBP.GetValue('GP_DATEPIECE'));
    end;
    //
    ZeroLigne (NewTOBL);
{$ENDIF}
  end;

{$IFDEF BTP}
 if (GenereMode = 'CBTTOLIVC') or (GenereMode= 'CCTOLIVCLI') then
  begin
  	InitCalcDispo;
  end;

  if NBLigneArt = 0 then // si pas de ligne d'article présent dans la tob
  begin
    TOBSource.free;
    Exit;
  end;
{$ENDIF}
  if TOBSource.detail.count = 0 then exit;
  //
  NumeroteLignesGC(Nil,NewTOBP) ;
  // MODIF LS --- maintenant que c'est bien trié
  if (GenereMode = 'REAP') or (GenereMode = 'REAPECLAT') or
  	 (GenereMode = 'VALIDDECIS') or (GenereMode = 'VALIDDECISECLAT') then DefiniLesLiensRegroup (TOBSource);
  // ----
  {Initialisations}
  NewNature:=NaturePiece ; Commentaires:=AvecComment ; bContinuOnError := False;
  bDuplicPiece := True;
  G_InitGenere(-1) ;
  {Création des TOB}
  G_CreerLesTOB ;
  {Mise à jour des tables avec TOBGenere}
  GenPiece:=T_GenereLesPieces.Create;
  GenPiece.TOBSource:=TOBSource ;
  GenPiece.MajPrix:=MajPrix;
  GenPiece.GenereMode:=GenereMode;
  GenPiece.GestionReliquat := GestionReliquat;
  GenPiece.TOBDESOUVRAGES := LesTOBOuvrages;
  GenPiece.TOBDESSSTRAIT := LesTOBSSTRAIT;
  // Modif BTP
  GenPiece.UpdateDev := UPdateDev;
  GenPiece.TOBResult := TOB.Create ('LES PIECES GENEREEES',nil,-1);
  GenPiece.OldTypePiece := OldTypePiece;
  GenPiece.WithLiaison := WithLiaisonDoc;
{$IFDEF BTP}
  if (GenereMode = 'CHANTOCBT') then GenPiece.RendActive := true;
{$ENDIF}
  // --
  io:=Transactions(GenPiece.GenereLesPieces,0) ;
  Case io of
     oeOk       : BEGIN
                  G_MajEvent;
                  StRendu:='Le traitement s''est correctement effectué. #13#10'+IntToStr(NbP)+' pièce(s) générée(s) de nature '
                             +RechDom('GCNATUREPIECEG',NewNature,False)+' du N° '+IntToStr(PremNum)+' au N° '+IntToStr(LastNum) ;
                  if CompteRendu then PGIInfo(StRendu,'Compte-rendu de génération') ;
                  // Modif BTP
                  if TOBResult <> nil then
                  begin
                    repeat
                      GenPiece.TOBResult.detail[0].ChangeParent(TOBresult,-1);
                    until GenPiece.TOBResult.detail.count = 0;
                  end;
                  // --
                  result:=true;
                  END ;
      oeUnknown : BEGIN
                  MessageAlerte('ATTENTION : La génération ne s''est pas complètement effectuée');
                  result:=false;
                  END;
      oeSaisie  : BEGIN
                  MessageAlerte('ATTENTION : Certaines pièces déjà en cours de traitement n''ont pas été enregistrées !') ;
                  result:=false;
                  END ;
     oeLettrage : BEGIN
                  MessageAlerte('ATTENTION : Certaines pièces ne peuvent pas passer en comptabilité et n''ont pas été enregistrées !') ;
                  result:=false;
                  END ;
     END ;

  {Libérations}
  G_FreeLesTOB(false) ;
  GenPiece.free;
  LesTOBOuvrages.free;
  LesTOBSSTRAIT.Free;
{$IFDEF BTP}
  GestionPhase.free;

  if listeAppel <> '' then
     Begin
     While ListeAppel <> '' do
        Begin
        if MsgErreur = '' then
           MsgErreur  := 'Consommation de l''appel : ' + ReadTokenST(ListeAppel) + ' non Facturée' + Chr(10)
        else
           MsgErreur  := MsgErreur + 'Consommation de l''appel : ' + ReadTokenST(ListeAppel) + ' non Facturée' + Chr(10);
        end;
     MsgErreur := MsgErreur + chr(10) + 'Si vous voulez les facturer renseignez le code chantier de référence dans ce(s) appel';
     pgibox(msgerreur, 'Erreur de Facturation');
     end;

{$ENDIF}
END ;


{============================ dupliqueLaPiece ========================}
function G_DupliquePieceBTP ( var NewCleDoc: R_CLEDOC) : boolean;
var QQ : TQuery;
		TOBFrs,TOBOuvrage,TOBBasesL : TOB;
    NewNum : integer;
    indice : integer;
    IndNum : integer;
begin
	result := false;
  TOBFrs := TOB.Create ('PIECE',nil,-1);
  TOBOuvrage := TOB.Create ('LES LIGNEOUV',nil,-1);
  TOBBasesL := TOB.Create ('LES LIGNES BASES',nil,-1);
  QQ := OpenSql ('SELECT * FROM PIECE WHERE '+WherePiece(NewCleDoc,ttdPiece,false),true,-1, '', True);
  if not QQ.eof then
  begin
    TOBFrs.selectDb ('',QQ);
    ferme(QQ);
    QQ := OpenSql ('SELECT * FROM LIGNE WHERE '+WherePiece(NewCleDoc,ttdligne,false),true,-1, '', True);
    TOBFrs.LoadDetailDB ('LIGNE','','',QQ,false);
    ferme(QQ);
    QQ := OpenSql('SELECT * FROM LIGNEOUV WHERE '+WherePiece(NewCleDoc,ttdOuvrage,false),true,-1, '', True);
    TOBOuvrage.LoadDetailDB ('LIGNEOUV','','',QQ,false);
    ferme (QQ);
    QQ := OpenSql('SELECT * FROM LIGNEBASE WHERE '+WherePiece(NewCleDoc,ttdLigneBase,false),true,-1, '', True);
    TOBBasesL.LoadDetailDB ('LIGNEBASE','','',QQ,false);
    ferme (QQ);
    //
    //FV1 :     newNum := SetNumberAttribution (NewCleDoc.Souche,1);
    newNum := SetNumberAttribution (NewCleDoc.NaturePiece, NewCleDoc.Souche,NewCleDoc.DatePiece,1);
    TOBfrs.putValue('GP_NUMERO',NewNum);
    For Indice := 0 to TOBfrs.detail.count -1 do
    begin
      if Indice = 0 then IndNum := TOBFrs.detail[Indice].GetNumChamp ('GL_NUMERO');
      TOBFrs.detail[Indice].PutValeur (IndNum,NewNum);
    end;
    For Indice := 0 to TOBOuvrage.detail.count -1 do
    begin
      if Indice = 0 then IndNum := TOBOuvrage.detail[Indice].GetNumChamp ('BLO_NUMERO');
      TOBOuvrage.detail[Indice].PutValeur (IndNum,NewNum);
    end;
    For Indice := 0 to TOBBasesL.detail.count -1 do
    begin
      if Indice = 0 then IndNum := TOBBasesL.detail[Indice].GetNumChamp ('BLB_NUMERO');
      TOBBAsesL.detail[Indice].PutValeur (IndNum,NewNum);
    end;
    //
    if Not TOBFrs.InsertDBByNivel (false) then V_PGI.IoError:=oeUnknown ;
    if V_PGI.ioError = OeOk then
    begin
    	if TOBOuvrage.detail.count > 0 then
      begin
    		if Not TOBOuvrage.InsertDB(nil,false) then V_PGI.IoError:=oeUnknown ;
      end;
    end;
    if V_PGI.ioError = OeOk then
    begin
    	if TOBBasesL.detail.count > 0 then
      begin
    		if Not TOBBasesL.InsertDB(nil,false) then V_PGI.IoError:=oeUnknown ;
      end;
    end;
    if V_PGI.ioError = OeOk then
    begin
    	NewCleDoc.NumeroPiece := NewNum;
      result := true;
    end;
  end;
  ferme(QQ);
  TOBFrs.free;
  TOBOuvrage.free;
  TOBBasesL.free;
end;

Procedure DupliqueLaPiece ( TOBPiece, TOBAffNew : TOB ; NewNat, NewTiers : String ;AvecCompteRendu,bEscompteTiers,bModeRegleTiers :Boolean ; NewDate:TDateTime=0) ;
Var StRendu,NewAff : String ;
    io   : TIoErr ;
    DuplicVal : T_DuplicVal ;
BEGIN
{Initialisations}
EclateParPiece := false; // GISE  10/07/03 car sinon le boolean n'étant pas initialisé
                         // on  avait soit true , soit false et ca plantait en dupplic affaire
NewNature:=NewNat ;  bContinuOnError:=false;
bDuplicPiece := True; bContremarque:=False;
//
MemoriseChampsSupLigneETL (NewNature ,True);
MemoriseChampsSupLigneOUV (NewNature);
MemoriseChampsSupPIECETRAIT;
//

// Modif BTP
ForceNumero := false;
// --
if TOBAffNew <> Nil then NewAff := TOBAffNew.GetValue('AFF_AFFAIRE')
                    else NewAff := '';
// Modif MODE
CodeAff:=''; LiaisonOrli:=False;
// -----

if NewDate>0 then LaNewDate:=NewDate else LaNewDate:=V_PGI.DateEntree ;
G_InitGenere(-1) ;
{Création des TOB}
G_CreerLesTOB ;
{Lecture des pièces}
DuplicVal:=T_DuplicVal.Create;
try
  if NewTiers = TOBPiece.GetValue('GP_TIERS') then NewTiers := '';
  if NewAff = TOBPiece.GetValue('GP_AFFAIRE') then NewAff := '';
DuplicVal:=T_DuplicVal.Create;
  DuplicVal.CodeTiers:=TOBPiece.GetValue('GP_TIERS') ;
  DuplicVal.TOBPiece:=TOBPiece ;
  DuplicVal.NewCodeTiers := NewTiers;
  DuplicVal.NewCodeAff := NewAff;
  DuplicVal.TOBAffNew:=TOBAffNew;
  DuplicVal.EscompteTiers :=bEscompteTiers;
  DuplicVal.ModeRegleTiers:=bModeRegleTiers;
  io:=Transactions(DuplicVal.G_ValideDuplicPiece,1);
finally
	DuplicVal.Free ;
end;

Case io of
   oeOk :  ;
   oeUnknown : BEGIN
             MessageAlerte('ATTENTION : La génération ne s''est pas complètement effectuée') ;
             END ;
   oeSaisie : BEGIN
             MessageAlerte('ATTENTION : Certaines pièces déjà en cours de traitement n''ont pas été enregistrées !') ;
             END ;
   oeLettrage : BEGIN
             MessageAlerte('ATTENTION : Certaines pièces ne peuvent pas passer en comptabilité et n''ont pas été enregistrées !') ;
             END ;
   oeStock   : BEGIN
             MessageAlerte('ATTENTION : Stock insuffisant pour certaines pièces qui n''ont pas été enregistrées !') ;
             END ;
   END ;

{Libérations}
G_FreeLesTOB(False) ;
{Compte rendu}
if ((AvecCompteRendu) and (PremNum>0) and (io=oeOk)) then
   BEGIN
   StRendu:='Pièce N° '+ IntToStr(PremNum) + ' générée correctement ';
   PGIInfo(StRendu,'Compte-rendu de génération') ;
   END ;
END ;


procedure G_NewAffversPiece( TobAff : TOB);
begin
TOBGenere.PutValue('GP_AFFAIRE', TobAff.GetValue('AFF_AFFAIRE'));
TOBGenere.PutValue('GP_AFFAIRE1',TobAff.GetValue('AFF_AFFAIRE1'));
TOBGenere.PutValue('GP_AFFAIRE2',TobAff.GetValue('AFF_AFFAIRE2'));
TOBGenere.PutValue('GP_AFFAIRE3',TobAff.GetValue('AFF_AFFAIRE3'));
TOBGenere.PutValue('GP_AVENANT', TobAff.GetValue('AFF_AVENANT'));
TOBGenere.PutValue('GP_APPORTEUR', TobAff.GetValue('AFF_APPORTEUR'));
end;


{============================ T_DuplicVal ========================}
Procedure T_DuplicVal.G_ValideDuplicPiece ;
Var k           : integer ;
    TOBL        : TOB;
    NewL        : TOB ;
    TiersPiece  : string;
BEGIN

  if (NewCodeTiers <> '') then TiersPiece := NewCodeTiers else TiersPiece := CodeTiers;

  G_AlimNewPiece(TOBPiece,true);

  if RemplirTOBTiers(TOBTiers,TiersPiece,NewNature,False)<>trtOk then V_PGI.IoError:=oeUnknown ;

  DEV.Code:=TOBPiece.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  G_ChargeCompletePiece(TOBPiece) ;
  UG_ChargeLesAcomptes(TOBPiece,TOBAcomptes);
  UG_ChargeLesPorcs(TOBPiece,TOBPorcs) ;
  {$IFDEF AFFAIRE}
  G_ChargeLesVariables(TOBPiece, TOBVariable);
  {$ENDIF}

  if (TOBPiece.GetValue('GP_NUMADRESSELIVR')<>0 ) or (TOBPiece.GetValue('GP_NUMADRESSELIVR')<>0) then
    G_ChargeLesAdresses(TOBPiece);

  if (NewCodeTiers <> '') then
  begin
    TobGenere.PutValue('GP_TIERS',Tobtiers.GetValue('T_TIERS'));
    TiersVersPiece(TobTiers,TobGenere);
    if Not(EscompteTiers) then
    begin
      TobGenere.PutValue('GP_ESCOMPTE',  TobPiece.GetValue('GP_ESCOMPTE')) ;
      TobGenere.PutValue('GP_QUALIFESCOMPTE',TobPiece.GetValue('GP_QUALIFESCOMPTE')) ;
    end;
    if Not(ModeRegleTiers) then TobGenere.PutValue('GP_MODEREGLE', TobPiece.GetValue('GP_MODEREGLE'));
  end;

  if (NewCodeAff <> '') then G_NewAffversPiece(TobAffNew);

  for k:=0 to TOBPiece.Detail.Count-1 do
  BEGIN
    TOBL:=TOBPiece.Detail[k] ;
    NewL:=NewTOBLigne(TOBGenere,0) ;
    NewL.Dupliquer(TOBL,True,True) ;
    G_LoadNomenLigne(NewL) ;
    {$IFDEF BTP}
    UG_LoadOuvrageLigne(NewL,TOBOuvrage,TOBArticles,indicenomen,RepriseAffaireRef) ;
    {$ENDIF}
    UG_LoadAnaLigne(TOBPiece,NewL) ;
    if (NewCodeTiers <> '') or (NewCodeAff <> '') then PieceVersLigne (TOBGenere,NewL);
    if (NewCodeAff <> '')  then AffaireVersLigne (TobGenere,NewL,TobAffNew);
  END ;

//G_GenereLaPiece(TOBPiece,nil,nil);
G_GenereLaPiece(TOBPiece,nil,TOBPieceTrait);

END ;

function  AjouteLignesToPiece (TOBLIgnes,ThePiece : TOB; GenereMode:String;WithLiaisonDoc : boolean=false;CompteRendu : boolean=true;TOBResult : TOB=nil) : boolean;
var io:TIoErr ;
    stRendu : string;
    NaturePiece : string;
    NumPiece : integer;
    Gestionreliquat : boolean;
    Remise , Escompte : double;
    TrierPar,RuptureCol : string;
    AvecComment,MajPrix,UpdateDev : boolean;
    Indice,IndiceNomen : integer;
    TOBL,OLDTOBPiece,TOBAffaire : TOB;
    QQ : Tquery;
    cledoc : R_Cledoc;
    MaxLig,MaxNumOrdre : integer;
//    IndLig,IndOrdre : integer;
    Qte       : double;
    Mt        : Double;
    DateLigne : TDateTime;
    XX : TPieceModifie;
    CodeTiers : string;
    created : Boolean;
    ok_ReliquatMt : Boolean;

   Function InitMode:Boolean;
   BEGIN
      MaxLig :=0;
      MaxNumOrdre := 0;
      result := false;
      GestionReliquat := false;
      NumPiece := ThePiece.GetValue('GP_NUMERO');
      Remise := 0; Escompte := 0;
      TrierPar := '';
      RuptureCol := '';
      AvecComment := false;
      MajPrix := false;
      UpdateDev := false;
      GenereMode := UpperCase(GenereMode);
{$IFDEF BTP}
      if GenereMode = 'DEVTOCHAN' then
      BEGIN
      // Prévision de chantier
      NaturePiece:=GetParamSoc('SO_BTNATCHANTIER');
      TrierPar:='';
      RuptureCol:='';
      AvecComment:=True;
      MajPrix:=False;
      UpdateDev := false;
      Result:=True;
      END;
{$ENDIF}
   END;

   Procedure TrierTOB(ParQuoi:String);
   BEGIN
      if ParQuoi <> '' then TOBLignes.Detail.Sort(ParQuoi);
   END;

begin
  created := false;

  // initialisation
  result := false;
  NaturePiece := '';
  NewNature := ThePiece.GetValue('GP_NATUREPIECEG');

  // Verif si traitement à lancer
  if TOBLignes.detail.count = 0 then
  begin
    if CompteRendu then
    begin
      PGIInfo('Rien à traiter','Compte-rendu de génération') ;
    end;
    exit;
  end;

  if not InitMode then Exit;

  if GenereMode = 'DEVTOCHAN' then
  begin
    // Dans ce cas la provenance de la ligne d'origine se fait par 'GL_NATUREPIECEG,GL_NUMERO,GL_NUMLIGNE+N1+N2+N3+N4+N5'
    NaturePiece:=GetParamSoc('SO_BTNATCHANTIER');
    bContremarque:=False;
  end;
  //
  MemoriseChampsSupLigneETL (NewNature ,True);
  MemoriseChampsSupLigneOUV (NewNature);
  MemoriseChampsSupPIECETRAIT;
  //
  {Création des TOB}
  G_CreerLesTOB ;

  TrierTOB(TrierPar);
  OLDTOBPiece := TOB.Create ('PIECE',nil,-1);

  TRY
    CodeTiers:=ThePiece.getValue('GP_TIERS');

    if RemplirTOBTiers(TOBTiers,CodeTiers,ThePiece.getValue('GP_NATUREPIECEG'),False)<>trtOk then V_PGI.IoError:=oeUnknown ;
    // Chargement de la piece de base
    TOBGenere.Dupliquer (ThePiece,true,true); // initiale
    //
    G_ChargeCompletePiece(TOBGenere,GenereMode) ;
    TOBaffaire := FindCetteAffaire (ThePiece.getValue('GP_AFFAIRE'));
    if TOBAffaire = nil then
    begin
    	StockeCetteAffaire (ThePiece.getValue('GP_AFFAIRE'));
    	TOBaffaire := FindCetteAffaire (ThePiece.getValue('GP_AFFAIRE'));
      created := true;
    end;

    cledoc := TOB2CleDoc (TOBGenere);
		LoadLesSousTraitants (cledoc,TOBSSTRAIT);  // récupération des sous traitants préalablement définis
    Setallused (TOBSSTRAIT);
    AddLesSupEntete (TOBGenere);
    UG_VideLesTobs (TOBBases,TOBEches,TOBNomenclature,TOBAcomptes,TOBPorcs,TOBSerie,TOBOuvrage,TOBLienOle,TOBPieceRG,TOBBasesRg,TobVariable,TobRevision,TobLigneTarif  );

    G_ChargeLesAdresses(TOBGenere);
    UG_ChargeLesAcomptes(TOBGenere,TOBAcomptes);
    UG_ChargeLesPorcs(TOBGenere,TOBPorcs) ;

    {$IFDEF AFFAIRE}
    G_ChargeLesRevisions(TOBGenere, TOBRevision);
    G_ChargeLesVariables(TOBGenere, TOBVariable);
    {$ENDIF}

  // --
    // positionne les lignes déjà présente comme non nouvelle
    Indicenomen := 1;

    for Indice := 0 to TOBGenere.detail.count -1 do
    begin
      TOBL := TOBGenere.detail[Indice];
      // Joli bug ma foi qui reinitialise les champs sup apres que l'on se soit fait mal a les recuperer
      //      InitTobLigne (TOBL);
			if TOBL.detail.count > 0 then TOBL.ClearDetail;
			if Assigned(TOBL) then NewTOBLigneFille(TOBL);

        // -----
      G_LoadNomenLigne(TOBL) ;
      G_LoadSerie (TOBL) ;
      // Modif BTP
{$IFDEF BTP}
      UG_LoadOuvrageLigne(TOBL,TOBOuvrage,TOBArticles,indicenomen,false) ;
{$ENDIF}
      // MODIF BTP
      G_ChargeLesRG(TOBL) ;
      // -------
      UG_LoadAnaLigne(TOBGenere,TOBL) ;
      // Pour garder la meme structure que les lignes en provenance des détails a ajouter
      TOBL.AddChampSupValeur('GP_REFINTERNE','');
      TOBL.AddChampSupValeur('GP_REPRESENTANT','');
      //
      TOBL.AddChampSupValeur ('PIECEORIGINE','');
      TOBL.AddChampSupValeur ('GP_NUMADRESSEFACT',-1);
      TOBL.AddChampSupValeur ('GP_NUMADRESSELIVR',-1);
      //
      TOBL.addChampsupValeur ('NEWONE','-');
    end;

    if TOBGenere.detail.count > 0 then
      MaxLig := TOBGenere.detail[TOBGenere.detail.count-1].GetValue('GL_NUMLIGNE')
    else
      MaxLig := 0;

    // Traitement des lignes à ajouter
    for Indice := 0 to TOBLignes.detail.count -1 do
    begin
      Inc(MaxLig);
      // positionne les clefs de l'ancien document
      cledoc := TOB2CleDoc (TOBLignes.detail[Indice]);
      MajFromCleDoc (OldTOBPiece,cledoc);
      // --
      if Indice = 0 then
      begin
      	TOBL := TOBLignes.detail[0];
      	G_LigneCommentaireBTP (ToBgenere,TOBL,True);
      	Inc(MaxLig);
      end;

      if (Pos(genereMode,'DEVTOCHAN;CHANTOCBT')>0) then
      begin
        // Dans le cas de génération de contre etude ou de génération de prévision de chantier
        if TOBL.GetString('GLC_NATURETRAVAIL')='002' then
        begin
          UG_recupSSTrait (TOBLignes.detail[Indice],TOBSSTRAIT);
        end;
      end;

//    TOBL := TOB.Create('LIGNE', TOBGenere, -1);
			TOBL := NewTOBLigne(TOBgenere,-1);
      if TOBGenere.getValue('GP_VIVANTE')= '-' Then TOBGenere.PutValue('GP_VIVANTE','X'); // reactive la piece
      EcritTOBLigne(TOBL, TOBLignes.detail[Indice], TOBGenere, Maxlig);
      if (GenereMode = 'DEVTOCHAN') or (Generemode = 'DEVTOCTE')  then
      begin
      	TOBL.AddChampSupValeur ('GLC_DOCUMENTLIE',EncodeRefPiece (TOBLignes.detail[Indice]));
      end;
      // Pour garder la meme structure que les lignes en provenance des détails a ajouter
      TOBL.AddChampSupValeur ('PIECEORIGINE',TOBLignes.detail[Indice].getValue('PIECEORIGINE'));
      TOBL.AddChampSupValeur ('GP_NUMADRESSEFACT',TOBLignes.detail[Indice].getValue('GP_NUMADRESSEFACT'));
      TOBL.AddChampSupValeur ('GP_NUMADRESSELIVR',TOBLignes.detail[Indice].getValue('GP_NUMADRESSELIVR'));
      //
      TOBL.addChampsupValeur ('NEWONE','X');
      //
      UG_AjouteLesArticles(OldTOBPiece,TOBArticles,TOBCpta,TOBTiers,TOBAnaP,TOBAnaS,TOBCatalogu,False) ;
      UG_AjouteLesRepres(OldTOBPiece,TOBComms) ;
      // -----
      G_LoadNomenLigne(TOBL) ;
      G_LoadSerie (TOBL) ;
      // Modif BTP
{$IFDEF BTP}
      UG_LoadOuvrageLigne(TOBL,TOBOuvrage,TOBArticles,indicenomen,true) ;
{$ENDIF}
      // MODIF BTP
      G_ChargeLesRG(TOBL) ;
      // -------
      UG_LoadAnaLigne(OldTOBPiece,TOBL) ;

  {$IFDEF BTP}
      TOBL.putValue('GL_IDENTIFIANTWOL',TOBLignes.detail[Indice].GetValue('GL_IDENTIFIANTWOL')); // ??
      if (genereMode = 'DEVTOCHAN') then TOBL.Data := TOBLignes.detail[Indice].Data;
      ok_ReliquatMt := CtrlOkReliquat(TOBL, 'GL');

  {$ENDIF}

      DateLigne:=TOBLignes.detail[Indice].GetValue('GL_DATELIVRAISON');
      if Dateligne<V_PGI.DateEntree then
      begin
        DateLigne:=V_PGI.DateEntree;
        TOBL.PutValue('GL_DATELIVRAISON',DateLigne);
      end;

      if GenereMode = 'DEVTOCHAN' then
      begin
        Qte:=TOBL.GetValue('GL_QTEFACT');
        TOBL.PutValue('GL_QTEFACT', Qte);
        TOBL.PutValue('GL_QTESTOCK',Qte);
        TOBL.PutValue('GL_QTERELIQUAT',0);
        TOBL.PutValue('GL_QTERESTE', Qte);  { NEWPIECE }
        // --- GUINIER ---
        if ok_ReliquatMt then
        begin
          Mt:=TOBL.GetValue('GL_MONTANTHTDEV');
          TOBL.PutValue('GL_MTRELIQUAT',0);
          TOBL.PutValue('GL_MTRESTE', Mt);  { NEWPIECE }
        end;
        TOBL.PutValue('GL_TYPEDIM','NOR');

        TOBL.PutValue('GL_QTESIT',0);
        TOBL.PutValue('GL_POURCENTAVANC',0);
        TOBL.PutValue('GL_QTEPREVAVANC',0);
        TOBL.PutValue('GL_TOTPREVAVANC',0);
        TOBL.PutValue('GL_TOTPREVDEVAVAN',0);
      end;
    end;
    NumeroteLignesGC(Nil,TOBGenere) ;
    // MODIF LS --- maintenant que c'est bien trié
    DefiniLesLiensRegroup (TOBgenere);
    // ----
    Io := OeOk;
    XX := TPieceModifie.Create;
    XX.TOBResult := TOBResult;
    XX.GestionReliquat := Gestionreliquat;
    XX.MajPrix := Majprix;
    XX.UpdateDev := UpdateDev;
    XX.WithLiaison := WithLiaisonDoc  ;
    XX.TOBAffaire := TOBAffaire;
    //
    XX.CompleteLaPiece; // lance le traitement de completion
    // --
      Case io of
        oeOk      : BEGIN
                    G_MajEvent;
                    StRendu:='Le traitement s''est correctement effectué. '
                               +RechDom('GCNATUREPIECEG',NewNature,False)+' N° '+IntToStr(TOBGenere.getValue('GP_NUMERO'))+' complété(e)' ;
                    if CompteRendu then PGIInfo(StRendu,'Compte-rendu de génération') ;
                    // --
                    result:=true;
                    END ;
        oeUnknown : BEGIN
                    MessageAlerte('ATTENTION : La génération ne s''est pas complètement effectuée');
                    result:=false;
                    END;
        oeSaisie  : BEGIN
                    MessageAlerte('ATTENTION : Certaines pièces déjà en cours de traitement n''ont pas été enregistrées !') ;
                    result:=false;
                    END ;
       oeLettrage : BEGIN
                    MessageAlerte('ATTENTION : Certaines pièces ne peuvent pas passer en comptabilité et n''ont pas été enregistrées !') ;
                    result:=false;
                    END ;
       END ;


  FINALLY
    if QQ <> nil then Ferme (QQ);
   {Libérations}
   OldTOBPiece.free;
   G_FreeLesTOB(false) ;
   if created then TOBAffaire.free;
  END;

end;

{ TPieceModifie }

destructor TPieceModifie.destroy;
begin
//  if TOBResult <> nil then TOBResult.free;
  inherited;
end;

procedure TPieceModifie.CompleteLaPiece;
var {CodeTiers : string;}
    Indice : integer;
    {TOBPiece,}TOBL,TOBA,TOBCata : TOB;
    Prix : double;
    AvecTarif:T_ActionTarifArt;
begin
  WithLiaisonInterDoc := WithLiaison;
  if (MajPrix) or (UpdateDev) then
  BEGIN
    FillChar(DEV,Sizeof(DEV),#0) ;
    DEV.Code:=TOBGenere.Getvalue('GP_DEVISE') ;
    GetInfosDevise(DEV) ;
    DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,TOBgenere.Getvalue('GP_DATECREATION')) ;
  END;

  for Indice := 0 to TOBgenere.detail.count -1 do
  begin
    TOBL:=TOBgenere.Detail[Indice] ;
    if TOBL.GetValue('GL_ARTICLE') <> '' then
    begin
      TOBA:=TOBArticles.FindFirst(['GA_ARTICLE'],[TOBL.getvalue('GL_ARTICLE')],False) ;
      TOBCata:=TOBCatalogu.FindFirst(['GCA_REFERENCE', 'GCA_TIERS'],[TOBL.GetValue('GL_REFCATALOGUE'), GetCodeFourDCM(TOBL)], False) ;
      if MajPrix then
      BEGIN
//         Prix:=0;
         AvecTarif:=ataCancel;
  {$IFDEF BTP}
         IF (TOBL.GetValue('GL_TYPEREF')<>'CAT') and (TOBL.GetValue('GL_TYPEREF')<>'') then
  {$ELSE}
         IF TOBL.GetValue('GL_TYPEREF')<>'CAT' then
  {$ENDIF}
  				AvecTarif:=TarifVersLigne(TOBA,TOBTiers,TOBL,TobLigneTarif,TOBGenere,TOBTarif,True,True,DEV) ;

         IF (TOBcata<>nil) And ((TOBL.GetValue('GL_TYPEREF')='CAT') or (AvecTarif<>ataOk)) then
         BEGIN
            Prix:=TOBCata.GetValue('GCA_PRIXBASE');
            if TOBL.GetValue('GL_FACTUREHT')='X' then
            BEGIN
               TOBL.PutValue('GL_PUHT',Prix) ; TOBL.PutValue('GL_PUHTDEV',Prix) ;
            END else
            BEGIN
               TOBL.PutValue('GL_PUTTC',Prix) ; TOBL.PutValue('GL_PUTTCDEV',Prix) ;
            END ;
            TOBL.PutValue('GL_DPA',TOBCata.GetValue('GCA_DPA')) ;
         END;

      END;
      ZeroLigneMontant (TOBL);  // reinit des montants de la ligne
    end;
  END;
  //
  G_EnregistreLaPieceCumule(TOBgenere,TOBAffaire,GestionReliquat,true);
  //
  if TOBResult <> nil then
  begin
    TOBResult.Dupliquer (TOBgenere,true,true);
  end;
end;

end.
