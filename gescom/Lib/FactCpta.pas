unit FactCpta ;

interface

Uses HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DB,  Fe_Main,
      {$IFDEF V530} EdtDOC,{$ELSE} EdtRdoc, {$ENDIF}
{$ENDIF}
     SysUtils, Dialogs, SaisUtil, UtilPGI, AGLInit, FactUtil, UtilSais,LettUtil,
     EntGC, Classes, HMsgBox, SaisComm, FactComm, ed_tools,LetBatch, FactArticle,
     FactRg,factcalc,
{$IFDEF AFFAIRE}
     AffaireUtil,
     AffaireCpta,
{$ENDIF}
{$IFDEF BTP}
			FactTvaMilliem,
      Factrgpbesoin,
{$ENDIF}
     TiersPayeur, Paramsoc,VentilCpta, EcrPiece_TOF,ULibEcriture,UtilIntercompta,
     TiersUtil,uEntCommun
     , UConnectWSConst;

Type T_RetCpta = (rcOk,rcRef,rcPar,rcMaj) ;

     T_GCAcompte = RECORD
                   JalRegle,CpteRegle : String ;
                   Libelle  : String ;
                   ModePaie,LienCpta : String ;
                   Montant  : Double ;
                   TypeCarte : string ;
                   CBLibelle : string;
                   CBInternet : string ;
                   DateExpire : string;
                   NumCheque : string ;
                   IsReglement : boolean ;
                   IsContrepartie : boolean ; // IsContrepartie=True, si contrepartie autre que le compte de tiers
                   CpteContre  : string ;     // Compte général de contrepartie, si différent du compte de tiers
                   LibelleContre  : string ;  // Libellé de l'écriture de contrepartie
                   DateEche : string ;        // Date échéance : à renseigner si différente de la date de la pièce
                   // MOdif BTP
                   IsModif : boolean;
                   NumEcr : integer;
                   LaTobAcc : TOB;
                   IsComptabilisable : boolean;
                   // Modif JT le 09/10/03 -- Gestion du tiers Facturé
                   AuxiFacture : String;                                                                      
                   // Modif LS du 20/03/2012 -- Gestion de la date de paiement
                   DateEcr : string;
                   END ;


     T_TypeEcr = (tecTiers,tecHT,tecTaxe,tecRemise,tecStock,TecRG,TecTaxeRg);

// Modif BTP
Procedure ReinitAcompteGCCpta ( TOBAcc : TOB ) ;
// --
Procedure LanceZoomEcriture ( TOBPiece : TOB ; DeStock : boolean ) ;
Procedure DetruitCompta ( TOBPiece : TOB ; DateSupp : TdateTime; Var OldEcr,OldStk : RMVT;ForModif : boolean=false ) ;
Function  PassationComptable ( TOBPiece,TOBOuvrages,TOBOuvragesP,TOBBases,TOBBasesL,TOBEches,TOBPieceTrait,TOBAffInterv,TOBTiers,TOBArticles,TOBCpta,TOBAcomptes,TOBPorcs,TOBPieceRG,TOBBasesRG,TOBAnaP,TOBAnaS,TOBPieceInterv,TOBVTECOLL  : TOB ; DEV : Rdevise ; OldEcr,OldStk : RMVT ; bParle : Boolean ; NewPiece : Boolean = True ) : boolean ;
Function  ChargeAjouteCompta ( TOBCpta,TOBPiece,TOBL,TOBA,TOBTiers,TOBCata,TobAffaire : TOB ; AvecAnal : boolean;TOBGCS:TOB=nil ) : TOB ;
{$IFDEF BTP}
Function  ChargeAjouteComptaBTP ( TOBCpta,TOBPiece,TOBL,TOBA,TOBTiers,TOBCata,TobAffaire : TOB ; FamTaxe : string; AvecAnal : boolean;TOBGCS:TOB=nil ) : TOB ;
{$ENDIF}
Procedure LoadLesAna ( TOBPiece,TOBAnaP,TOBAnaS : TOB ) ;
Procedure PreVentileLigne ( TOBC,TOBAP,TOBAS,TOBL : TOB ) ;
procedure ValideAnalytiques ( TOBPiece,TOBAnaP,TOBAnaS : TOB ) ;
Function  EnregistreAcompte ( TOBPiece,TOBTiers : TOB ; LeAcc : T_GCAcompte ; Quiet : boolean = false ) : TOB ;
Procedure DetruitAcompteGCCpta ( TOBPiece,TOBAcc : TOB ; ModifRefGC : boolean = False; VideRefGC : boolean = False) ;
Procedure MajAccRegleDiff ( TOBPiece,TOBAcomptes : TOB ; OldNum : integer) ;
Function  InsertionDifferee ( TOBEcr : TOB ) : boolean ;
Function  CreerTOBCpta : TOB ;
// cd 050701
Function  DecodeRefCPGescom ( RefG : String ) : R_CleDoc ;
Function  DecodeRefGCComptable ( RefC : String ) : RMVT ;
// cd 050701
Procedure RechLibTiersCegid (VenteAchat : string; TobTiers,TOBPiece,TobAdresses: TOB) ;
Function  FindTOBCode ( TOBGC : TOB ; FamArt,FamTiers,FamAff,Etabl,Regime,FamTaxe : String ) : TOB ;
Function  FabricSQLCompta ( FamArt,FamTiers,FamAff,Etabl,Regime,FamTaxe,VenteAchat : String ) : String ;
Function  CreerTOBCodeCpta ( TOBParent : TOB ) : TOB ;
Function  LibEcrATraiter(TypeEcr : T_TypeEcr;Nature : string; LibEcr : Boolean):Boolean;
Procedure RenseigneTiersFact ( TOBPiece,TOBTiers,TOBTiersCpta : TOB ) ; // JT 09/10/2003 (utilsé dans UtofGcAcomptes
{$IFDEF BTP}
function VerifEcriturePieceModifiable (TOBPiece : TOB) : boolean;
{$ENDIF}
Procedure SuiteVentilLigne ( RefA : String ; LaTV,TOBAna,TOBL : TOB ; Ecr : boolean ) ;
function TraitementVOuvrage (TOBEcr,TOBPieceinterv,TOBAFFInterv,TOBPiece,TOBL,TOBOuvrages,TOBOuvragesP,TOBTiers,TOBArticles,TOBCpta,TOBPorcs: TOB;MM : RMVT ; NbDec : integer; EnHt : boolean; DEV : Rdevise) : T_RetCpta;
function TraitementVLigne (TOBEcr,TOBAFFInterv,TOBPieceinterv,TOBA,TOBPiece,TOBL,TOBTiers,TOBArticles,TOBCpta,TOBPorcs : TOB ; MM : RMVT ; NbDec : integer; EnHt : boolean ;DEV : Rdevise; TOBMereLigne : TOb=nil ) : T_RetCpta;
Procedure GCCreerPiecePayeur(TOBECR, TOBTiers : TOB; DocInfo : T_WSDocumentInf);//DocType : string; DocNumber : integer) ;
Function  EnregistreReglSSTrait ( TOBPiece,TOBClient,TOBFournisseur : TOB ; LeAcc : T_GCAcompte ; Quiet : boolean = false ) : TOB ;
Procedure DetruitLesCompta ( TOBlistPiece : TOB ; DateSupp : TdateTime; Var OldEcr,OldStk : RMVT;ForModif : boolean=false ) ;
Function TraiteFamilleAffaire ( CodeAffaire:string;  TobAffaire : TOB) : string;
function GetWSDocumentInfFromTob(TobGP : TOB) : T_WSDocumentInf;
function GetWSDocumentInfFromRmvt(lRmvt : RMVT) : T_WSDocumentInf;

implementation
Uses
  FactVariante
  , FactOuvrage
  , UCoTraitance
  , FactTimbres
  , FactTOB
  , factLigneBase
  , UtilTobPiece
  , Factretenues
  , ULiquidTva2014
  , UCumulCollectifs
  , UConnectWSCEGID
  {$IFDEF MODE}
  , FactCptaMode
  {$ENDIF}
  ;

type MODIF_Diff = class(TOB)
     private
       function YouzTob(T : TOB) : integer;
       procedure LoadFromBuffer ( Buffer : String );
     end;

Var PremHT,PremTVA,LastMsg,NbEches,NbAcc : integer ;
    DistinctAffaire : Boolean ;
    LaDEV           : RDEVISE ;
    TTA,TTStock,TTVarStk : TList ;
    RGComptabilise : Boolean;
    TobTVASurEncaiss : TOB;
		GereTvaMixte : boolean;
    TheTOBRG : TOB;

Const TTCSANSTVA : boolean = False ;

Type T_CodeCpta = Class
                  CptHT,LibArt,Affaire,Depot : String ;
                  FamTaxe : Array[1..5] of String ;
                  SommeTaxeD,SommeTaxeP : Array[1..5] of double ;
                  XD,XP,Qte : Double ;
                  LibU : boolean ;
                  Anal : TList ;
                  // BTP
                  FamilleTvaActive : integer;
                  // --
                  Constructor Create ;
                  Destructor Destroy ;  override;
                  End ;

Type T_DetAnal = Class
                 Section,Ax : String ;
                 MD,MP,Qte1,TotQte1 : Double ;
                 End ;

Type T_MontantArtFi = RECORD   // Montants des articles financiers pour une pièce.
                   Montant     : Double ;
                   MontantDev  : Double ;
                   END ;

Type T_CreatPont = (ccpEscompte,ccpRemise,ccpHT,ccpTaxe,ccpStock,ccpVarStk,ccpRG) ;

const TexteMessage: array[1..22] of string 	= (
          {1}  'Compte général d''escompte absent ou incorrect'
          {2} ,'Compte général de remise absent ou incorrect'
          {3} ,'Compte général de taxe absent ou incorrect'
          {4} ,'Compte collectif tiers absent ou incorrect'
          {5} ,'Compte général de HT absent ou incorrect'
          {6} ,'Compte général d''écart de conversion absent ou incorrect'
          {7} ,'Journal comptable non renseigné pour cette nature de pièce'
          {8} ,'Nature comptable non renseignée'
          {9} ,'Erreur sur la numérotation du journal comptable'
         {10} ,'Certains comptes généraux, auxiliaires ou analytiques sont incorrects'
         {11} ,'Compte général de stock ou variation absent ou incorrect'
         {12} ,'Compte de retenue de garantie absent ou incorrect'
         {13} ,'Compte général de taxe RG absent ou incorrect'
         {14} ,'Compte général de banque (associé au journal ou au mode de paiement) absent ou incorrect'
         {15} ,'Le compte comptable général n''existe pas dans votre plan comptable'
         {16} ,'Journal non renseigné'
         {17} ,'Erreur sur le client'
         {18} ,''
         {19} ,'Ecart sur contrôle des pièces d''achat' {DBR CPA}
         {20} ,''
         {21} ,''
         {22} ,''
              );

procedure CallWSCegid(WseT : T_WSEntryType; TobEcr : TOB; DocInfo : T_WSDocumentInf);
var
  SendEntry : TSendEntryY2;
begin
  if TGetParamWSCEGID.ConnectToY2 then
  begin                                          
    SendEntry := TSendEntryY2.Create;         
    try
      SendEntry.ServerName := '';
      SendEntry.DBName     := '';
      {$IF not defined(APPSRV)}
      SendEntry.SendEntryCEGID(WseT, TobEcr, DocInfo);
      {$IFEND !APPSRV}
    finally
      SendEntry.Free;
    end;
  end;
end;

procedure RePackBase (TObbases,TOBBasesST : TOB; DEV : RDevise);

  function GetDeviseFourn(Fournisseur : string) : string;
  var QQ: TQuery;
      SQL : String;
  begin
    SQl := 'SELECT T_DEVISE FROM TIERS WHERE T_TIERS="'+Fournisseur+'" AND T_NATUREAUXI="FOU"';
    QQ := OPenSql(SQL,true,1,'',true);
    if not QQ.eof then result := QQ.Fields[0].AsString;
    ferme (QQ);
  end;

var
  indice : integer;
  TOBB,TOBBE,TOBBST : TOB;
  FamilleTaxe : string;
  Tauxtaxe : double;
begin
  if TOBBases.detail.Count = 0 then Exit;
  Indice := 0;
  repeat
    TOBB := TOBBases.detail[Indice];
    if TOBB.getValue('GPB_FOURN')='' then
    begin
      inc(Indice);
    end else
    begin
      if (TOBBasesST <> nil) then
      begin
        if IsAutoLiquidationTvaST (TOBB.GetString('GPB_FOURN')) then
        begin
          FamilleTaxe := GetTvaST (TOBB.GetValue('GPB_FOURN'));
          Tauxtaxe := GetTauxTaxeST(TOBB.GetString('GPB_FOURN'),FamilleTaxe);
          TOBBST := TOBBasesST.findfirst(['GPB_CATEGORIETAXE','GPB_FAMILLETAXE','GPB_FOURN'],
                                         ['TX1',FamilleTaxe,''],true);
          if TOBBST = nil then
          begin
            TOBBST := TOB.Create ('PIEDBASE',TOBBasesST,-1);
            TOBBST.SetString('GPB_CATEGORIETAXE',TOBB.GetString('GPB_CATEGORIETAXE'));
            TOBBST.SetString('GPB_FAMILLETAXE',FamilleTaxe);
            TOBBST.SetDouble('GPB_TAUXTAXE',Tauxtaxe);
            TOBBST.SetDouble('GPB_TAUXDEV',TOBB.GetValue('GPB_TAUXDEV'));
            TOBBST.SetDouble('GPB_COTATION',TOBB.GetValue('GPB_COTATION'));
            TOBBST.SetString('GPB_FOURN','');
          end;
          //
          TOBBST.PutValue('GPB_BASEDEV',TOBBST.GetValue('GPB_BASEDEV')+TOBB.GetValue('GPB_BASEACHAT'));
          TOBBST.PutValue('GPB_BASETAXE',TOBBST.getValue('GPB_BASEDEV'));
          //
          TOBBST.PutValue('GPB_VALEURDEV',TOBBST.GetValue('GPB_VALEURDEV')+TOBB.GetValue('GPB_VALEURACHAT'));
          TOBBST.PutValue('GPB_VALEURTAXE',TOBBST.GetValue('GPB_VALEURDEV'));
        end;
      end;
      TOBBE := TOBBASES.findfirst(['GPB_CATEGORIETAXE','GPB_FAMILLETAXE','GPB_FOURN'],
                                  [TOBB.GetValue('GPB_CATEGORIETAXE'),TOBB.GetValue('GPB_FAMILLETAXE'),''],true);
      if TOBBE <> nil then
      begin
         TOBBE.PutValue('GPB_BASETAXE',TOBBE.GetValue('GPB_BASETAXE')+TOBB.GetValue('GPB_BASETAXE'));
         TOBBE.PutValue('GPB_BASETAXETTC',TOBBE.GetValue('GPB_BASETAXETTC')+TOBB.GetValue('GPB_BASETAXETTC'));
         TOBBE.PutValue('GPB_BASEDEV',TOBBE.GetValue('GPB_BASEDEV')+TOBB.GetValue('GPB_BASEDEV'));
         TOBBE.PutValue('GPB_BASETTCTTC',TOBBE.GetValue('GPB_BASETTCTTC')+TOBB.GetValue('GPB_BASETTCTTC'));
         TOBBE.PutValue('GPB_VALEURTAXE',TOBBE.GetValue('GPB_VALEURTAXE')+TOBB.GetValue('GPB_VALEURTAXE'));
         TOBBE.PutValue('GPB_VALEURDEV',TOBBE.GetValue('GPB_VALEURDEV')+TOBB.GetValue('GPB_VALEURDEV'));
         TOBB.Free;
      end else
      begin
        TOBB.PutValue('GPB_FOURN','');
        inc(indice);
      end;
    end;
  until Indice >= TOBBases.detail.count;
end;

function ComptabiliseTrait (Fournisseur: string; TOBAFFInterv,TOBPieceInterv : TOB) : boolean;
var TOBT : TOB;
begin
	result := false;
  if not (GetParamSocSecur('SO_BTCPTAPAIEDIRECT',false)) then exit;
  if TOBAFFInterv = nil then
  begin
    Result := True;
    exit;
  end;
  TOBT := TOBAFFInterv.FindFirst(['BAI_TIERSFOU'],[Fournisseur],true);
  if TOBT <> nil then exit; // pas de comptabilisation pour les elements cotraitants
//	for II := 0 to TOBAFFInterv.detail.count -1 do
  //
  TOBT := TOBPieceInterv.FindFirst(['BPI_TIERSFOU'],[Fournisseur],True);
  if TOBT <> nil then
  begin
    if (TOBT.getvalue('BPI_TYPEPAIE')='002') or
       ((TOBT.getvalue('BPI_TYPEPAIE')='001') and (GetParamSocSecur('SO_BTCPTAPAIEDIRECT',false))) then
    begin
      // Comptabilisation des elements osu traites si reglement par la societe
      // (Donc pas en direct)
      result := true;
    end;
  end;
end;


function IsCompteTaxeSurEncaissement(LeCompte : string) : boolean;
begin
  Result := false;
  if LeCompte = '' then exit;
  Result := ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = "' + LeCompte + '" AND G_TVASURENCAISS = "X"');
end;

procedure CumulTVASurEncaissement(LeCompte : string; TobL : TOB);
var TobTmp : TOB;
    Cpt : integer;
    Cat, Fam : string;
    Montant, MontantDev, TotalHt, TotalHtDev, TauxTva : double;
    OkEncaiss,prefixe : string;
begin
  prefixe := GetPrefixeTable (TOBL);

  if (LeCompte = '') or (not assigned(TobL)) then exit;
  { Test si le compte est TVA sur encaissement }
  if IsCompteTaxeSurEncaissement(LeCompte) then
    OkEncaiss := 'X'
    else
    OkEncaiss := '-';
  if OkEncaiss = '-' then exit;
  { Créer une fille ou un cumul pour les 5 types de taxes }
  for Cpt := 1 to 5 do
  begin
    Cat := 'TX' + IntToStr(Cpt);
    Fam := TobL.GetString(prefixe+'_FAMILLETAXE' + IntToStr(Cpt));
    if Fam <> '' then
    begin
      TauxTva := 1;
      if OkEncaiss = 'X' then
      begin
        Montant := TobL.GetDouble(prefixe+'_TOTALTAXE' + IntToStr(Cpt));
        MontantDev := Arrondi(TobL.GetDouble(prefixe+'_TOTALTAXEDEV' + IntToStr(Cpt)),V_PGI.okdecv);
        TotalHt := TobL.GetDouble(prefixe+'_TOTALHT');
        TotalHtDev := Arrondi(TobL.GetDouble(prefixe+'_TOTALHTDEV'),V_PGI.okdecV);
        if MontantDev = 0 then continue; // Correction FQ 12653
        if TotalHtDev <> 0 then TauxTva := arrondi(MontantDev/TotalHtDev,4)
                           else TauxTva := 1;  
      end else
      begin
        Montant := 0;
        MontantDev := 0;
        TotalHt := 0;
        TotalHtDev := 0;  
      end;
      TobTmp := TobTVASurEncaiss.FindFirst(['INDEX'],[Cat + ';' +Fam], true);
      if not assigned(TobTmp) then
      begin
        TobTmp := TOB.Create('', TobTVASurEncaiss, -1);
        TobTmp.AddChampSupValeur('INDEX', Cat + ';' + Fam);
        TobTmp.AddChampSupValeur('BASE', TotalHt);
        TobTmp.AddChampSupValeur('BASEDEV', TotalHtDev);
        TobTmp.AddChampSupValeur('MONTANT', Montant);
        TobTmp.AddChampSupValeur('MONTANTDEV', MontantDev);
        TobTmp.AddChampSupValeur('TAUX', TauxTva);
        TobTmp.AddChampSupValeur('OKENCAISS', OkEncaiss);
      end else
      if Montant <> 0 then
      begin
        TobTmp.Setdouble('MONTANT', (TobTmp.GetDouble('MONTANT') + Montant));
        TobTmp.Setdouble('MONTANTDEV', (TobTmp.GetDouble('MONTANTDEV') + MontantDev));
        TobTmp.Setdouble('BASE', TobTmp.GetDouble('BASE') + TotalHt);
        TobTmp.Setdouble('BASEDEV', TobTmp.GetDouble('BASEDEV') + TotalHtDev);
      end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 20/02/2006
Modifié le ... :
Description .. : Test si on gère la TVA mixte
Suite ........ : si ok, en fin de test c'est G_TVASURENCAISS (sur lignes HT)
Suite ........ : qui indique si on peut le faire
Mots clefs ... : LIAISON;COMPTABILITE;PONT;PASSATION;
*****************************************************************}
function OkTVAMixte(TobTiers: TOB; MM : RMVT; NatureJal : string; AccepteOD : boolean=false) : boolean;
begin
  Result := false;
  { J_NATUREJAL = vide }
  if NatureJal = '' then exit;
  { Pas de journal défini }
  if MM.Jal = '' then exit;
  { On ne gère pas la TVA sur encaissement }
  if Not VH^.OuiTvaEnc then exit;
  { Régime ou la tva sur encaissement non renseigné sur le tiers }
  if ((TobTiers.GetString('T_REGIMETVA') = '') or (TobTiers.GetString('T_TVAENCAISSEMENT') = '')) then exit;
  { La racine du collectif ne correspond pas à celui défini }
  if not EstCollFact(TobTiers.GetString('T_COLLECTIF')) then exit;
  { Mauvaise nature comptable de la pièce }
  if ((MM.Nature <> 'FC') and (MM.Nature <> 'AC') and (MM.Nature <> 'FF') and (MM.Nature <> 'AF') and
      ((MM.Nature <> 'OD') and (AccepteOD))) then Exit;
  { On est en vente : Ok }
  if NatureJal = 'VTE' then
    Result := true
  { En achat, test si le FO est défini comme mixte }
  else if NatureJal = 'ACH' then
    Result := (Pos(TobTiers.GetString('T_TVAENCAISSEMENT'),'TE;TM')>0);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 27/03/2000
Modifié le ... : 27/03/2000
Description .. : Encodage d'un identifiant d'écriture comptable dans la pièce Gescom
Mots clefs ... : LIAISON;COMPTABILITE;PONT;PASSATION;
*****************************************************************}
Function EncodeRefGCComptable ( TOBEcr : TOB ) : String ;
BEGIN
Result:='' ; if TOBEcr=Nil then Exit ;
Result:=TOBEcr.GetValue('E_JOURNAL')+';'+FormatDateTime('ddmmyyyy',TOBEcr.GetValue('E_DATECOMPTABLE'))+';'
       +IntToStr(TOBEcr.GetValue('E_NUMEROPIECE'))+';'+TOBEcr.GetValue('E_QUALIFPIECE')+';'+TOBEcr.GetValue('E_NATUREPIECE')+';' ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 27/03/2000
Modifié le ... : 27/03/2000
Description .. : Décodage d'un identifiant d'écriture en pièce Gescom vers un identifiant comptable
Mots clefs ... : LIAISON;COMPTABILITE;PONT;PASSATION;
*****************************************************************}
Function DecodeRefGCComptable ( RefC : String ) : RMVT ;
Var XX : RMVT ;
    StC : String ;
BEGIN
FillChar(XX,Sizeof(XX),#0) ;
if RefC = '' then exit;
StC:=ReadTokenSt(RefC) ; XX.Jal:=StC ;
StC:=ReadTokenSt(RefC) ; XX.DateC:=EvalueDate(StC) ; XX.Exo:=QuelExoDT(XX.DateC) ;
StC:=ReadTokenSt(RefC) ; XX.Num:=StrToInt(StC) ;
StC:=ReadTokenSt(RefC) ; XX.Simul:=StC ;
StC:=ReadTokenSt(RefC) ; XX.Nature:=StC ;
Result:=XX ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 27/03/2000
Modifié le ... : 27/03/2000
Description .. : Décodage d'un identifiant de pièce Gescom en Compta vers un identifiant Gescom
Mots clefs ... : LIAISON;COMPTABILITE;PONT;PASSATION;
*****************************************************************}
Function DecodeRefCPGescom ( RefG : String ) : R_CleDoc ;
Var CD : R_CleDoc ;
    StC  : String ;
BEGIN
FillChar(CD,Sizeof(CD),#0) ;
StC:=ReadTokenSt(RefG) ; CD.NaturePiece:=StC ;
StC:=ReadTokenSt(RefG) ; CD.Souche:=StC ;
StC:=ReadTokenSt(RefG) ; CD.DatePiece:=EvalueDate(StC) ;
StC:=ReadTokenSt(RefG) ; CD.NumeroPiece:=StrToInt(StC) ;
StC:=ReadTokenSt(RefG) ; CD.Indice:=StrToInt(StC) ;
Result:=CD ;
END ;

//******************************************************************************
//**************** Traitement des libellés des écritures ***********************
//******************************************************************************
Function LibEcrATraiter(TypeEcr : T_TypeEcr;Nature : string; LibEcr : Boolean):Boolean;
Var Valeurs, Ligne : string;
begin
Result := true; if not(LibEcr) then Exit;
Case TypeEcr of   // correspondance avec les valeurs de la tablette
   tecTiers : Ligne := 'TIE';
   tecHT    : Ligne := 'HT';
   tecTaxe  : Ligne := 'TAX';
   tecRemise: Ligne := 'REM';
   // Modif BTP
   tecRG: Ligne := 'RG';
   TecTaxeRG: Ligne := 'TRG';
   else BEGIN Result:=False ; Exit ; END ;
   End;
Valeurs := Trim(GetInfoParPiece(Nature,'GPP_TYPEECRALIM'));
if Pos(Ligne,Valeurs)> 0 then Result:=true else Result:=False;
end;

Function RecupValeursLibEcr(Tobpiece,TobTiers: TOB; ValeurEcr :string):string;
Var tmp,champ,Prefixe,Valeur,CodeTiers,TiersCpta,SQL : string;
		QQ : TQuery;
BEGIN
  CodeTiers:=TOBPiece.GetValue('GP_TIERS') ; TiersCpta:=TOBPiece.GetValue('GP_TIERSFACTURE') ;
  Result := ''; if ValeurEcr ='' then Exit;
  tmp:=ReadTokenSt(ValeurEcr);
  While (tmp <> '') do
  begin
     Champ:=RechDom('GCLIBECRITURECPTA',tmp,True) ;
     if Champ <> '' then
     begin
        Valeur := '';
        Prefixe := Copy (Champ,1,Pos('_',Champ));
        if Prefixe = 'GP_' then
        begin
          if (champ = 'GP_TIERS') then
          begin
            if ((CodeTiers=TiersCpta) or (TiersCpta='')) then
            begin
              if TobPiece.Fieldexists(Champ) then Valeur:=Trim(Tobpiece.GetValue(Champ));
            end else
            begin
              if TobPiece.Fieldexists('GP_TIERSFACTURE') then Valeur:=Trim(Tobpiece.GetValue('GP_TIERSFACTURE'));
            end;
          end else
          begin
            if TobPiece.Fieldexists(Champ) then Valeur:=Trim(Tobpiece.GetValue(Champ));
          end;
        end else if Prefixe = 'T_' then
        begin
          if (champ = 'T_LIBELLE') then
          begin
            if ((CodeTiers=TiersCpta) or (TiersCpta='')) then
            begin
            	if TobTiers.FieldExists(Champ) then Valeur:=Trim(TobTiers.GetValue(Champ));
            end else
            begin
              Valeur := '';
              SQl := 'SELECT T_LIBELLE FROM TIERS WHERE T_NATUREAUXI="'+TOBTiers.GetString('T_NATUREAUXI')+'" AND '+
              			 'T_TIERS="'+TiersCpta+'"';
              QQ := OpenSql(SQL,True,1,'',True);
              if not QQ.eof then
              begin
								Valeur := QQ.Fields[0].AsString;
              end;
              ferme (QQ);
            end;
          end else
          begin
            if TobTiers.FieldExists(Champ) then Valeur:=Trim(TobTiers.GetValue(Champ));
          end;
        end else if Prefixe = 'AFF_' then
        begin
          Valeur := GetChampsAffaire(Tobpiece.GetValue('GP_AFFAIRE'),Champ);
        end;
        if Result ='' then Result:=Valeur else Result:=Result+' '+Valeur;
     end;
     tmp:=ReadTokenSt(ValeurEcr);
  end;
END;

Procedure AlimLibEcr (TobE,TobPiece,TobTiers: TOB;LibDefaut:string; TypeEcr : T_TypeEcr; LibEcr, ComptaDiff : Boolean);
Var Nature,Fixe1,Fixe2,ValeurEcr1,ValeurEcr2,st1,st2,stTot : string;
BEGIN
  if TobE = Nil then Exit;
  // JT - eQualité 10168 - Compta diff, affecte le libelle du tiers sauf pour SIC.
  // Le lib est retraité dans CptaDiff selon si on garde le détail par tiers ou non
  if (ComptaDiff) and (not VH_GC.GCIfDefCEGID) then
  begin
{$IFDEF BTP}
  	if LibEcr then TOBE.PutValue('E_LIBELLE',LibDefaut) else TOBE.PutValue('E_REFINTERNE',LibDefaut);
{$ELSE}
    TOBE.PutValue('E_LIBELLE',LibDefaut);
    exit;
{$ENDIF}
  end;
  if LibEcr then TOBE.PutValue('E_LIBELLE',LibDefaut) else TOBE.PutValue('E_REFINTERNE',LibDefaut);
  if (TobPiece=Nil) or (TobTiers=Nil) then Exit;
  Nature :=TOBPiece.GetValue('GP_NATUREPIECEG');
  if Not(LibEcrATraiter(TypeEcr,Nature,LibEcr)) then Exit;

  if LibEcr then
     begin
     Fixe1:=Trim(GetInfoParPiece(Nature,'GPP_RACINELIBECR1')); Fixe2:=Trim(GetInfoParPiece(Nature,'GPP_RACINELIBECR2'));
     ValeurEcr1:=GetInfoParPiece(Nature,'GPP_VALEURLIBECR1'); ValeurEcr2:=GetInfoParPiece(Nature,'GPP_VALEURLIBECR2');
     end else
     begin
     Fixe1:=Trim(GetInfoParPiece(Nature,'GPP_RACINEREFINT1')); Fixe2:=Trim(GetInfoParPiece(Nature,'GPP_RACINEREFINT2'));
     ValeurEcr1:=GetInfoParPiece(Nature,'GPP_VALEURREFINT1'); ValeurEcr2:=GetInfoParPiece(Nature,'GPP_VALEURREFINT2');
     end;

  if (Fixe1='') And (Fixe2='') And (ValeurEcr1='') And (ValeurEcr2='') then Exit;
  st1 := RecupValeursLibEcr(Tobpiece,TobTiers,ValeurEcr1);
  st2 := RecupValeursLibEcr(Tobpiece,TobTiers,ValeurEcr2);
  if Fixe1<>'' then stTot := Fixe1 +' '+st1 else stTot := st1;
  if Fixe2<>'' then stTot := stTot+' '+Fixe2+' '+st2 else stTot := stTot+' '+st2;
  Trim(stTot);
  if stTot <> '' then
     begin
     if LibEcr then TOBE.PutValue('E_LIBELLE',Copy(stTot,1,35)) else TOBE.PutValue('E_REFINTERNE',Copy(stTot,1,35));
     end;
END;

Function RecupTypeEcrCpta ( Nature,Etablissement : String) : String ;
BEGIN
Result:=GetInfoParPieceCompl(Nature,Etablissement,'GPC_TYPEECRCPTA') ;
if (Result='') then Result:=GetInfoParPiece(Nature,'GPP_TYPEECRCPTA') ;
END;

Procedure LanceZoomEcriture ( TOBPiece : TOB ; DeStock : boolean ) ;
Var XX : RMVT ;
    RefCP : String ;
BEGIN
if DeStock then RefCP:=TOBPiece.GetValue('GP_REFCPTASTOCK') else RefCP:=TOBPiece.GetValue('GP_REFCOMPTABLE') ;
if ((RefCP='') or (RefCP='DIFF')) then Exit ;
XX:=DecodeRefGCComptable(RefCP) ;
XX.CodeD:=TOBPiece.GetValue('GP_DEVISE') ;
XX.Etabl:=TOBPiece.GetValue('GP_ETABLISSEMENT') ;
LanceZoomBordereau(XX) ;
END ;

procedure MajAnaEntete ( TOBP,TOBAP,TOBAS : TOB ) ;
Var TOBV : TOB ;
    RefA : String ;
    k    : integer ;
BEGIN
RefA:=EncodeRefCPGescom(TOBP) ;
for k:=0 to TOBAP.Detail.Count-1 do
    BEGIN
    TOBV:=TOBAP.Detail[k] ;
    TOBV.PutValue('YVA_IDENTIFIANT',RefA) ;
    TOBV.PutValue('YVA_NATUREID','GC') ;
    TOBV.PutValue('YVA_IDENTLIGNE',FormatFloat('000',0)) ;
    END ;
for k:=0 to TOBAS.Detail.Count-1 do
    BEGIN
    TOBV:=TOBAS.Detail[k] ;
    TOBV.PutValue('YVA_IDENTIFIANT',RefA) ;
    TOBV.PutValue('YVA_NATUREID','GC') ;
    TOBV.PutValue('YVA_IDENTLIGNE',FormatFloat('000',0)) ;
    END ;
END ;

procedure ValideAnalytiques ( TOBPiece,TOBAnaP,TOBAnaS : TOB ) ;
Var RefA : String ;
    TOBL,TOBAL,TOBV : TOB ;
    i,k        : integer ;
BEGIN
MajAnaEntete(TOBPiece,TOBAnaP,TOBAnaS) ;
RefA:=EncodeRefCPGescom(TOBPiece) ;
for i:=0 to TOBPiece.Detail.Count-1 do
    BEGIN
    TOBL:=TOBPiece.Detail[i] ;
    if TOBL.Detail.count = 0 then
    begin
      NewTOBLigneFille (TOBL);
    end;
    TOBAL:=TOBL.Detail[0] ;
    for k:=0 to TOBAL.Detail.Count-1 do
        BEGIN
        TOBV:=TOBAL.Detail[k] ;
        TOBV.PutValue('YVA_IDENTIFIANT',RefA) ;
        TOBV.PutValue('YVA_NATUREID','GC') ;
        TOBV.PutValue('YVA_IDENTLIGNE',FormatFloat('000',i+1)) ;
        END ;
    TOBAL:=TOBL.Detail[1] ;
    for k:=0 to TOBAL.Detail.Count-1 do
        BEGIN
        TOBV:=TOBAL.Detail[k] ;
        TOBV.PutValue('YVA_IDENTIFIANT',RefA) ;
        TOBV.PutValue('YVA_NATUREID','GC') ;
        TOBV.PutValue('YVA_IDENTLIGNE',FormatFloat('000',i+1)) ;
        END ;
    END ;
END ;

Procedure LoadLesAna ( TOBPiece,TOBAnaP,TOBAnaS : TOB ) ;
Var Q : TQuery ;
    TOBL,TOBA,TOBAL : TOB ;
    i,NumL : integer ;
    RefA   : String ;
BEGIN
RefA:=EncodeRefCPGescom(TOBPiece) ; if RefA='' then Exit ;
Q:=OpenSQL('SELECT * FROM VENTANA WHERE (YVA_TABLEANA="GL" OR YVA_TABLEANA="GS") AND YVA_IDENTIFIANT="'+RefA+'"',True,-1, '', True) ;
if Not Q.EOF then
   BEGIN
   TOBAnaP.LoadDetailDB('VENTANA','','',Q,False) ;
   for i:=TOBAnaP.Detail.Count-1 downto 0 do
       BEGIN
       TOBA:=TOBAnaP.Detail[i] ;
       if TOBA.GetValue('YVA_TABLEANA')='GS' then TOBA.ChangeParent(TOBAnaS,0) ;
       END ;
   END ;
Ferme(Q) ;
TOBAnaP.Detail.Sort('-YVA_AXE;-YVA_IDENTLIGNE') ;
for i:=TOBAnaP.Detail.Count-1 downto 0 do
    BEGIN
    TOBA:=TOBAnaP.Detail[i] ; NumL:=ValeurI(TOBA.GetValue('YVA_IDENTLIGNE')) ;
    if ((NumL>0) and (NumL<=TOBPiece.Detail.Count)) then
       BEGIN
       TOBL:=TOBPiece.Detail[NumL-1] ;
       if TOBL.Detail.Count<=0 then TOBAL:=TOB.Create('',TOBL,-1) else TOBAL:=TOBL.Detail[0] ;
       TOBA.ChangeParent(TOBAL,-1) ;
       END ;
    END ;
TOBAnaS.Detail.Sort('-YVA_AXE;-YVA_IDENTLIGNE') ;
for i:=TOBAnaS.Detail.Count-1 downto 0 do
    BEGIN
    TOBA:=TOBAnaS.Detail[i] ; NumL:=ValeurI(TOBA.GetValue('YVA_IDENTLIGNE')) ;
    if ((NumL>0) and (NumL<=TOBPiece.Detail.Count)) then
       BEGIN
       TOBL:=TOBPiece.Detail[NumL-1] ; TOBAL:=TOBL.Detail[1] ;
       TOBA.ChangeParent(TOBAL,-1) ;
       END ;
    END ;
END ;

Procedure PreVentileCEGID ( TOBL : TOB ) ;
Var TOBPiece,TOBA,TOBAna : TOB ;
    RefA,prefixe     : String ;
BEGIN
  prefixe := GetPrefixeTable (TOBL);

if TOBL=Nil then Exit ;
if TOBL.NomTable<>'LIGNE' then Exit ;
TOBPiece:=TOBL.Parent ; if TOBPiece=Nil then Exit ;
RefA:=EncodeRefCPGescom(TOBPiece) ;
TOBAna:=TOBL.Detail[0] ; TOBAna.ClearDetail ;
{Axe 1 = Activité = 710 en dur ...}
TOBA:=TOB.Create('VENTANA',TOBAna,-1) ;
TOBA.PutValue('YVA_TABLEANA','GL') ; TOBA.PutValue('YVA_NATUREID','GC') ;
TOBA.PutValue('YVA_AXE','A1')      ; TOBA.PutValue('YVA_IDENTIFIANT',RefA) ;
TOBA.PutValue('YVA_IDENTLIGNE',FormatFloat('000',TOBL.GetValue(prefixe+'_NUMLIGNE'))) ;
TOBA.PutValue('YVA_SECTION','710') ;
TOBA.PutValue('YVA_POURCENTAGE',100.0) ; TOBA.PutValue('YVA_NUMVENTIL',1) ;
{Axe 1 = Agences = Table libre pièce 2}
TOBA:=TOB.Create('VENTANA',TOBAna,-1) ;
TOBA.PutValue('YVA_TABLEANA','GL') ; TOBA.PutValue('YVA_NATUREID','GC') ;
TOBA.PutValue('YVA_AXE','A2')      ; TOBA.PutValue('YVA_IDENTIFIANT',RefA) ;
TOBA.PutValue('YVA_IDENTLIGNE',FormatFloat('000',TOBL.GetValue(prefixe+'_NUMLIGNE'))) ;
if  TOBPiece.GetValue('GP_LIBRETIERS2')='' then TOBA.PutValue('YVA_SECTION',VH^.Cpta[fbAxe2].Attente)
                                           else TOBA.PutValue('YVA_SECTION',TOBPiece.GetValue('GP_LIBRETIERS2')) ;
TOBA.PutValue('YVA_POURCENTAGE',100.0) ; TOBA.PutValue('YVA_NUMVENTIL',1) ;
{Axe 3 = Marché + Canal = Tables libres pièce 1 et 3}
TOBA:=TOB.Create('VENTANA',TOBAna,-1) ;
TOBA.PutValue('YVA_TABLEANA','GL') ; TOBA.PutValue('YVA_NATUREID','GC') ;
TOBA.PutValue('YVA_AXE','A3')      ; TOBA.PutValue('YVA_IDENTIFIANT',RefA) ;
TOBA.PutValue('YVA_IDENTLIGNE',FormatFloat('000',TOBL.GetValue(prefixe+'_NUMLIGNE'))) ;
if (TOBPiece.GetValue('GP_LIBRETIERS1')='') or (TOBPiece.GetValue('GP_LIBRETIERS3')='')
   then TOBA.PutValue('YVA_SECTION',VH^.Cpta[fbAxe3].Attente)
   else TOBA.PutValue('YVA_SECTION',TOBPiece.GetValue('GP_LIBRETIERS1')+TOBPiece.GetValue('GP_LIBRETIERS3')) ;
TOBA.PutValue('YVA_POURCENTAGE',100.0) ; TOBA.PutValue('YVA_NUMVENTIL',1) ;
END ;

Procedure SuiteVentilLigne ( RefA : String ; LaTV,TOBAna,TOBL : TOB ; Ecr : boolean ) ;
Var TOBA,TOBV : TOB ;
    i    : integer ;
    prefixe  : string;
BEGIN

  prefixe := GetPrefixeTable (TOBL);

if LaTV.Detail.Count<=0 then Exit ;
for i:=0 to LaTV.Detail.Count-1 do
    BEGIN
    TOBV:=LaTV.Detail[i] ;
    TOBA:=TOB.Create('VENTANA',TOBAna,-1) ;
    if Ecr then TOBA.PutValue('YVA_TABLEANA','GL') else TOBA.PutValue('YVA_TABLEANA','GS') ;
    TOBA.PutValue('YVA_IDENTIFIANT',RefA) ;
    TOBA.PutValue('YVA_NATUREID','GC') ;
    TOBA.PutValue('YVA_IDENTLIGNE',FormatFloat('000',TOBL.GetValue(prefixe+'_NUMLIGNE'))) ;
    if TOBV.NomTable='VENTANA' then
       BEGIN
       TOBA.PutValue('YVA_SECTION',TOBV.GetValue('YVA_SECTION')) ;
       TOBA.PutValue('YVA_POURCENTAGE',TOBV.GetValue('YVA_POURCENTAGE')) ;
       TOBA.PutValue('YVA_NUMVENTIL',TOBV.GetValue('YVA_NUMVENTIL')) ;
       TOBA.PutValue('YVA_AXE',TOBV.GetValue('YVA_AXE')) ;
       END else
       BEGIN
       TOBA.PutValue('YVA_SECTION',TOBV.GetValue('V_SECTION')) ;
       TOBA.PutValue('YVA_POURCENTAGE',TOBV.GetValue('V_TAUXMONTANT')) ;
       TOBA.PutValue('YVA_NUMVENTIL',TOBV.GetValue('V_NUMEROVENTIL')) ;
       TOBA.PutValue('YVA_AXE','A'+Copy(TOBV.GetValue('V_NATURE'),3,1)) ;
       END ;
    END ;
END ;

Procedure PreVentileLigne ( TOBC,TOBAP,TOBAS,TOBL : TOB ) ;
Var RefA : String ;
II : Integer;
    TOBPiece,TOBAna,TOBStk,LaTV,LaTS : TOB ;
BEGIN
  if TOBL=Nil then Exit ;
  TOBPiece:=TOBL.Parent ;
  TOBAna:=Nil ;
  TOBStk:=Nil ;
  if VH_GC.GCIfDefCEGID then  // gm le 27/12/02 pour ne pas générer d'ana sur les pièce affaire
  if (TOBPiece.GetValue('GP_NATUREPIECEG') <> GetParamSoc('SO_AFNATAFFAIRE')) and
  	 (TOBPiece.GetValue('GP_NATUREPIECEG') <> GetParamSoc('SO_AFNATPROPOSITION'))  then
  Begin
  	if TOBPiece.GetValue('GP_VENTEACHAT')='VEN' then BEGIN PreVentileCEGID(TOBL) ; Exit ; END ;
  End;
  RefA:=EncodeRefCPGescom(TOBPiece) ;
  if (TOBAP<>Nil) and (TOBL.Detail[0].Detail.count = 0) then
  BEGIN
    TOBAna:=TOBL.Detail[0] ;
//    TOBAna.ClearDetail ;
  END ;
  if (TOBAS<>Nil) and (TOBL.Detail[1].detail.count = 0) then
  BEGIN
  	TOBStk:=TOBL.Detail[1] ;
//    TOBStk.ClearDetail ;
  END ;
  if TOBC<>Nil then
  BEGIN
    LaTV:=TOBC.Detail[0] ; LaTS:=TOBC.Detail[1] ;
    if TOBAP<>Nil then if ((LaTV.Detail.Count<=0) and (TOBAP.Detail.Count>0)) then LaTV:=TOBAP ;
    if TOBAS<>Nil then if ((LaTS.Detail.Count<=0) and (TOBAS.Detail.Count>0)) then LaTS:=TOBAS ;
  END else
  BEGIN
  	LaTV:=TOBAP ; LaTS:=TOBAS ;
  END ;
  if ((LaTV<>Nil) and (TOBAna<>Nil)) then SuiteVentilLigne(RefA,LaTV,TOBAna,TOBL,True) ;
  if ((LaTS<>Nil) and (TOBStk<>Nil)) then SuiteVentilLigne(RefA,LaTS,TOBStk,TOBL,False) ;
  // FS#506 - DELABOUDINIERE : Pb affectation analytique en saisie de facture d'achat
  // FS#583 - TEAM RESEAUX - problème de ventilation analytique
  if TOBC <> nil then
  begin
    for II := 0 to TOBC.detail.count -1 do
    begin
      TOBC.detail[II].ClearDetail;
    end;
  end;
END ;

Procedure PreVentileLigneDetailOuv ( TOBC,TOBAP,TOBAS,TOBOD,TOBL : TOB ) ;
Var RefA : String ;
    TOBPiece,TOBAna,TOBStk,LaTV,LaTS : TOB ;
BEGIN
  if TOBL=Nil then Exit ;
  TOBPiece:=TOBL.Parent ;
  TOBAna:=Nil ;
  TOBStk:=Nil ;
  RefA:=EncodeRefCPGescom(TOBPiece) ;
  if (TOBAP<>Nil) and (TOBOD.Detail[0].Detail.count = 0) then
  BEGIN
    TOBAna:=TOBOD.Detail[0] ;
  END ;
  if (TOBAS<>Nil) and (TOBOD.Detail[1].detail.count = 0) then
  BEGIN
  	TOBStk:=TOBOD.Detail[1] ;
  END ;
  if TOBC<>Nil then
  BEGIN
    LaTV:=TOBC.Detail[0] ; LaTS:=TOBC.Detail[1] ;
    if TOBAP<>Nil then if ((LaTV.Detail.Count<=0) and (TOBAP.Detail.Count>0)) then LaTV:=TOBAP ;
    if TOBAS<>Nil then if ((LaTS.Detail.Count<=0) and (TOBAS.Detail.Count>0)) then LaTS:=TOBAS ;
  END else
  BEGIN
  	LaTV:=TOBAP ; LaTS:=TOBAS ;
  END ;
  if ((LaTV<>Nil) and (TOBAna<>Nil)) then
  begin
  	SuiteVentilLigne(RefA,LaTV,TOBAna,TOBOD,True) ;
  end;
  if ((LaTS<>Nil) and (TOBStk<>Nil)) then
  begin
  	SuiteVentilLigne(RefA,LaTS,TOBStk,TOBOD,False) ;
  end;
END ;



Procedure SuiteVentilPorcs ( RefA : String ; TOBpiece, LaTV,TOBAna : TOB ; Ecr : boolean ) ;
Var TOBA,TOBV : TOB ;
    i    : integer ;
BEGIN
if LaTV.Detail.Count<=0 then Exit ;
for i:=0 to LaTV.Detail.Count-1 do
    BEGIN
    TOBV:=LaTV.Detail[i] ;
    TOBA:=TOB.Create('VENTANA',TOBAna,-1) ;
    if Ecr then TOBA.PutValue('YVA_TABLEANA','GL') else TOBA.PutValue('YVA_TABLEANA','GS') ;
    TOBA.PutValue('YVA_IDENTIFIANT',RefA) ;
    TOBA.PutValue('YVA_NATUREID','GC') ;
    TOBA.PutValue('YVA_IDENTLIGNE',FormatFloat('00000000000',TOBpiece.getValue('GP_NUMERO')+I)) ;
    if TOBV.NomTable='VENTANA' then
       BEGIN
       TOBA.PutValue('YVA_SECTION',TOBV.GetValue('YVA_SECTION')) ;
       TOBA.PutValue('YVA_POURCENTAGE',TOBV.GetValue('YVA_POURCENTAGE')) ;
       TOBA.PutValue('YVA_NUMVENTIL',TOBV.GetValue('YVA_NUMVENTIL')) ;
       TOBA.PutValue('YVA_AXE',TOBV.GetValue('YVA_AXE')) ;
       END else
       BEGIN
       TOBA.PutValue('YVA_SECTION',TOBV.GetValue('V_SECTION')) ;
       TOBA.PutValue('YVA_POURCENTAGE',TOBV.GetValue('V_TAUXMONTANT')) ;
       TOBA.PutValue('YVA_NUMVENTIL',TOBV.GetValue('V_NUMEROVENTIL')) ;
       TOBA.PutValue('YVA_AXE','A'+Copy(TOBV.GetValue('V_NATURE'),3,1)) ;
       END ;
    END ;
END ;

Procedure PreVentilePorcs ( TOBL,TOBAP,TOBAS,TOBPiece : TOB ) ;
Var RefA : String ;
    TOBAnaP,TOBanaS,LaTV,LaTS : TOB ;
BEGIN
  if TOBL=Nil then Exit ;
  RefA:=EncodeRefCPGescom(TOBPiece) ;
  LaTV:=TOBAP ; LaTS:=TOBAS ;
  TOBAnaP := TOBL.detail[0];
  TOBAnaS := TOBL.detail[1];
  if (LaTV<>Nil) then SuiteVentilporcs(RefA,TOBpiece,LaTV,TOBAnaP,True) ;
  if (LaTS<>Nil) then SuiteVentilporcs(RefA,TOBPiece,LaTS,TOBAnaS,False) ;
END ;


Procedure RenseigneClefCompta ( TOBPiece : TOB ; Var MM : RMVT ) ;
BEGIN
MM.Etabl:=TOBPiece.GetValue('GP_ETABLISSEMENT') ;
MM.DateC:=TOBPiece.GetValue('GP_DATEPIECE') ; MM.Exo:=QuelExoDT(MM.DateC) ;
MM.CodeD:=TOBPiece.GetValue('GP_DEVISE') ;
MM.DateTaux:=TOBPiece.GetValue('GP_DATETAUXDEV') ;
MM.TauxD:=TOBPiece.GetValue('GP_TAUXDEV') ;
MM.ModeSaisieJal:='-' ;
END ;

Constructor T_CodeCpta.Create ;
BEGIN
inherited Create ;
Anal:=TList.Create ;
END ;

Destructor T_CodeCpta.Destroy ;
BEGIN
Anal.Free ;
inherited Destroy ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 27/03/2000
Modifié le ... : 27/03/2000
Description .. : Destruction de l'écriture comptable en transformation ou modification de pièce Gescom
Mots clefs ... : LIAISON;COMPTABILITE;PONT;PASSATION;
*****************************************************************}
Procedure DetruitEcritureFromMM ( var MM : RMVT ; InTrans : Boolean ; DateAnnul : TDateTime; ForModif : boolean=false ) ;
Var
  Q         : TQuery;
  Ax        : String;
  Nb        : integer;
  i         : integer;
  k         : integer;
  NumL      : integer;
  NumA      : integer;
  Indice    : integer;
  TOBEcr    : TOB;
  TOBAna    : TOB;
  TOBE      : TOB;
  TOBA      : TOB;
  TOBAxe    : TOB;
  TOBECREX  : TOB;
  TOBTiers  : TOB;
  okok      : boolean;
  DocInfo   : T_WSDocumentInf;
begin
  {Mise à jour inverse des soldes des comptes}
  if MM.Simul = 'N' then
  begin
    TOBEcr   := TOB.Create('',Nil,-1) ;
    TOBAna   := TOB.Create('',Nil,-1) ;
    TOBEcrEx := TOB.Create('',Nil,-1) ;
    TOBTiers := TOB.Create ('TIERS', nil, -1);
    try
      Q := OpenSQL('SELECT * FROM ECRITURE WHERE '+WhereEcriture(tsGene,MM,False),True,-1, '', True) ;
      try
        if Not Q.EOF then
          TOBEcr.LoadDetailDB('ECRITURE', '', '', Q, False, True) ;
      finally
        Ferme(Q);
      end;
      if TOBEcr.detail.count > 0 then 
      begin
        {Analytiques}
        if Assigned(TOBEcr.FindFirst(['E_ANA'], ['X'], True)) then
        begin
          Q := OpenSQL('SELECT * FROM ANALYTIQ WHERE ' + WhereEcriture(tsAnal,MM,False),True,-1, '', True) ;
          try
            if Not Q.EOF then TOBAna.LoadDetailDB('ANALYTIQ', '', '', Q, False, True) ;
          finally
            Ferme(Q) ;
          end;
          if TOBAna.Detail.count > 0 then
          begin
            {Changement de parent}
            TOBAna.Detail.Sort('Y_NUMLIGNE;Y_NUMVENTIL') ;
            for i := TOBAna.Detail.Count-1 downto 0 do
            begin
              TOBA := TOBAna.Detail[i] ;
              NumL := TOBA.GetValue('Y_NUMLIGNE') ;
              Ax   := TOBA.GetValue('Y_AXE') ;
              NumA := Ord(Ax[2])-48 ;
              TOBE := TOBEcr.FindFirst(['E_NUMLIGNE'],[NumL],False) ;
              if TOBE.Detail.Count <= 0 then
              begin
                for k:=1 to 5 do
                  TOB.Create('A'+IntToStr(k),TOBE,-1) ;
              end;
              TOBAxe := TOBE.Detail[NumA-1];
              TOBA.ChangeParent(TOBAxe,-1);
            end;
          end;
        end;
        DelettreEcriture(MM, TobEcr);
        Q := OpenSql('SELECT * FROM TIERS WHERE T_AUXILIAIRE = "' + TOBEcr.detail[0].GetString('E_AUXILIAIRE') + '"', True, 1, '', true);
        try
          if not Q.eof then TOBTiers.SelectDB('', Q);
        finally
          ferme(Q);
        end;
        if ForModif then
        begin
          {MAJ Soldes}
          MajSoldesEcritureTOB(TOBEcr,False) ;
          Nb := 0;
          try
            Nb := ExecuteSQL('DELETE FROM ECRITURE WHERE ' + WhereEcriture(tsGene,MM,False)) ;
          except
            on E: Exception do
              PgiError('Erreur : ' + E.Message, 'Suppression écriture');
          end;
          if ((Nb <= 0) and (InTrans)) then
          begin
            MessageValid := 'Erreur Lors de la suppression des écritures comptables';
            V_PGI.IoError:=oeUnknown ;
            Exit ;
          end;
          try
            ExecuteSQL('DELETE FROM ANALYTIQ WHERE ' + WhereEcriture(tsAnal,MM,False)) ; {! il peut ne pas y avoir d'analytique}
          except
            on E: Exception do
              PgiError('Erreur SQL : ' + E.Message, 'Suppression analytique');
          end;
        end else
        begin
          DelettreEcritureinMem (TOBECR);
          {Mise à jour de la pièce d'origine -- Pour éviter de faire repointer sur la pièce disparuuuu..au coin de..}
          for Indice := 0 to TOBEcr.detail.count -1 do
            TOBEcr.detail[Indice].SetString('E_REFGESCOM','');
          {Extourne ecriture}
          TOBEcrEx := ExtourneEcriture(TOBECR, false, StrToDate(DateToStr(DateAnnul)), TobEcr.detail[0].GetString('E_QUALIFPIECE'),
                                       GetParamSocSecur('SO_MONTANTNEGATIF', false), '', '', '', False);
          Okok := False;
          try
            Okok := TOBEcrEx.InsertDBByNivel(False);
          except
            on E: Exception do
              PgiError('Erreur SQL : ' + E.Message, 'Extourne pièce');
          end;
          DocInfo := GetWSDocumentInfFromRmvt(MM);
          if not Okok then
          begin
            V_PGI.IoError := oeLettrage;
          end else
          begin
            try
              MajSoldesEcritureTOB(TobEcrEx, true);
            except
              on E: Exception do
                PgiError('Erreur SQL : ' + E.Message, 'Extourne pièce - Solde');
            end;
            if V_PGI.IOError <> oeOk then
            begin
              V_PGI.IoError := oeLettrage;
            end else
              CallWSCegid(wsetExtourne, TOBEcrEx, DocInfo);
          end;
          {Ecriture sur tiers payeur de l'écriture extournée}
          GCCreerPiecePayeur(TOBEcrEx, TOBTiers, DocInfo); //MM.DocType, MM.DocNumber);
          {Update ecritures principales}
          try
            TOBEcr.UpdateDB(false);
          except
            on E: Exception do
              PgiError('Erreur SQL : ' + E.Message, 'Enreg. Pièce comptable');
          end;
        end;
      end;
    finally
      TOBEcr.Free ;
      TOBEcrEx.free;
      TOBAna.Free ;
      TOBTiers.free;
    end;
  end else
  begin
    Nb := 0;
    try
      Nb:=ExecuteSQL('DELETE FROM ECRITURE WHERE '+WhereEcriture(tsGene,MM,False)) ;
    except
      on E: Exception do
        PgiError('Erreur SQL : ' + E.Message, 'Suppression simul.');
    end;
    if ((Nb<=0) and (InTrans)) then
    begin
      V_PGI.IoError:=oeUnknown ;
      Exit ;
    end;
    try
      ExecuteSQL('DELETE FROM ANALYTIQ WHERE '+WhereEcriture(tsAnal,MM,False)) ; {! il peut ne pas y avoir d'analytique}
    except
      on E: Exception do
        PgiError('Erreur SQL : ' + E.Message, 'Suppression simul. Analytique');
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : JT (eQualité 10246)
Créé le ...... : 10/10/2003
Modifié le ... :
Description .. : Modification de la table ECRITURE lors annulation de la saisie
Suite ........ : d'une pièce en gardant l'acompte/règlement créé
Mots clefs ... : LIAISON;COMPTABILITE;ACOMPTE;REGLEMENT
*****************************************************************}
Procedure ModifEcritureFromMM ( MM : RMVT ; InTrans : Boolean ; NewRef : string) ;
Var Nb : integer ;
begin
  Nb := ExecuteSQL('UPDATE ECRITURE SET E_REFGESCOM="'+NewRef+'" WHERE '+WhereEcriture(tsGene,MM,False)) ;
  if ((Nb<=0) and (InTrans)) then
    V_PGI.IoError := oeUnknown;
end ;

function MODIF_Diff.YouzTob(T : TOB) : integer;
begin
  T.ChangeParent(Self,-1) ;
  Result:=0 ;
end;

procedure MODIF_Diff.LoadFromBuffer(Buffer : String );
begin
  TOBLoadFromBuffer(Buffer,YouzTob) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : JT (eQualité 10246)
Créé le ...... : 10/10/2003
Modifié le ... :
Description .. : Modification de la table COMPTADIFFEREE lors annulation de la saisie
Suite ........ : d'une pièce en gardant l'acompte/règlement créé
Mots clefs ... : LIAISON;COMPTABILITE;DIFFEREE;ACOMPTE;REGLEMENT
*****************************************************************}
Procedure ModifEcritureDifferee(ReqCptaDiff, NewRef : string; InTrans : Boolean);
Var Qry : TQuery;
    Sql, Buffer, Buffer1 : string;
    TobCptaDiff, TobTmp, TobTmp1 : TOB;
    TobLue : MODIF_DIFF;
    Cpt : integer;
begin
  Sql := 'SELECT * '+ ReqCptaDiff +'AND (GCD_NATURECOMPTA="OC" OR GCD_NATURECOMPTA="RC")';
  Qry := OpenSQL(Sql,True,-1, '', True);
  TobCptaDiff := TOB.Create('LES DIFFEREES',nil,-1);
  TobCptaDiff.LoadDetailDB('COMPTADIFFEREE','','',Qry,False);
  Ferme(Qry);
  for Cpt := 0 to TobCptaDiff.Detail.count -1 do
  begin
    Buffer := TobCptaDiff.Detail[Cpt].GetValue('GCD_BLOBECRITURE');
    TobLue := MODIF_DIFF.Create('',Nil,-1);
    TobLue.LoadFromBuffer(Buffer);
    TobTmp := TobLue.FindFirst(['!E_REFGESCOM'],[''],True);
    while TobTmp <> nil  do
    begin
      TobTmp.PutValue('E_REFGESCOM',NewRef);
      TobTmp.PutValue('E_BLOCNOTE','');
      TobTmp := TobLue.FindNext(['!E_REFGESCOM'],[''],True);
    end;
    TobTmp1 := TOB.Create('',nil,-1);
    TobTmp1.Dupliquer(TobLue.Detail[0],True,True);
    Buffer1 := TobTmp1.SaveToBuffer(True,False,'');
    TobCptaDiff.Detail[Cpt].PutValue('GCD_BLOBECRITURE',Buffer1);
    FreeAndNil(TobLue);
    FreeAndNil(TobTmp1);
  end;
  TobCptaDiff.UpdateDB(False);
  FreeAndNil(TobCptaDiff);
end ;

Procedure DetruitLesCompta ( TOBlistPiece : TOB ; DateSupp : TdateTime; Var OldEcr,OldStk : RMVT;ForModif : boolean=false ) ;
var II : Integer;
begin
	for II := 0 to TOBlistPiece.detail.Count -1 do
  begin
    DetruitCompta (TOBlistPiece.detail[II],DateSupp,OldEcr,OldStk,ForModif);
  end;
end;

Procedure DetruitCompta ( TOBPiece : TOB ; DateSupp : TdateTime; Var OldEcr,OldStk : RMVT;ForModif : boolean=false ) ;
Var
  RefComptable : string;
  RefCptaStock : String ;
  DD           : TDateTime ;
begin
  FillChar(OldEcr,Sizeof(OldEcr),#0) ;
  FillChar(OldStk,Sizeof(OldStk),#0) ;
  if Assigned(TOBPiece) then
  begin
    RefComptable := TOBPiece.GetString('GP_REFCOMPTABLE') ;
    RefCptaStock := TOBPiece.GetString('GP_REFCPTASTOCK') ;
    if ((RefComptable='') and (RefCptaStock='')) then Exit ;
    if RefComptable='DIFF' then
    begin
      DD := TOBPiece.GetValue('GP_DATEPIECE');
      RefComptable:=EncodeRefCPGescom(TOBPiece);
      if ExecuteSQL('DELETE FROM COMPTADIFFEREE WHERE GCD_DATEPIECE="'+UsDateTime(DD)+'" AND GCD_REFPIECE="'+RefComptable+'"') <> 1 then
        V_PGI.IoError := oeunknown;
    end else
    begin
      DD:=TOBPiece.GetValue('GP_DATEPIECE') ;
      if RefComptable<>'' then
      begin
        OldEcr := DecodeRefGCComptable(RefComptable) ;
        OldEcr.DocType := TOBPiece.GetString('GP_NATUREPIECEG');
        OldEcr.DocNumber := TOBPiece.GetInteger('GP_NUMERO');
        DetruitEcritureFromMM(OldEcr, True, DD, ForModif) ;
      end;
      if RefCptaStock<>'' then
      begin
        OldStk:=DecodeRefGCComptable(RefCptaStock) ;
        OldStk.DocType := TOBPiece.GetString('GP_NATUREPIECEG');
        OldStk.DocNumber := TOBPiece.GetInteger('GP_NUMERO');
        DetruitEcritureFromMM(OldStk, True, DD, ForModif) ;
      end ;
    end ;
  end;       
end ;

Procedure RecupLesDC ( TOBL : TOB ; Var XD,XP : Double; OnlyTimbre : Boolean=false ) ;
Var CodeD : String ;
BEGIN
XD:=0 ; XP:=0 ;
if TOBL.NomTable='PIEDCOLLECTIF' then
   BEGIN
     XP:=TOBL.GetValue('BPB_REAJUSTE');
     XD:=TOBL.GetValue('BPB_REAJUSTEDEV') ;
   END else
if TOBL.NomTable='PIEDECHE' then
   BEGIN
   if not OnlyTimbre then
   begin
     XP:=TOBL.GetValue('GPE_MONTANTECHE') + TOBL.GetValue('GPE_TIMBRE');
     CodeD:=TOBL.GetValue('GPE_DEVISE') ;
     if CodeD<>V_PGI.DevisePivot then XD:=TOBL.GetValue('GPE_MONTANTDEV') + TOBL.GetValue('GPE_TIMBRE') else XD:=XP ;
   end else
   begin
     XP:=TOBL.GetValue('GPE_TIMBRE');
     CodeD:=TOBL.GetValue('GPE_DEVISE') ;
     if CodeD<>V_PGI.DevisePivot then XD:=TOBL.GetValue('GPE_TIMBRE') else XD:=XP ;
   end;
   END else
if TOBL.NomTable='LIGNE' then
   BEGIN
   CodeD:=TOBL.GetValue('GL_DEVISE');
   if TTCSANSTVA then
      BEGIN
      {Cas du non assujetti}
      XP:=TOBL.GetValue('GL_MONTANTTTC') ;
      if CodeD<>V_PGI.DevisePivot then XD:=TOBL.GetValue('GL_MONTANTTTCDEV') else XD:=XP ;
      END else
      BEGIN
      XP:=TOBL.GetValue('GL_MONTANTHT') ;
      if CodeD<>V_PGI.DevisePivot then XD:=TOBL.GetValue('GL_MONTANTHTDEV') else XD:=XP ;
      END ;
   END else
if TOBL.NomTable='LIGNEOUVPLAT' then
   BEGIN
   CodeD:=TOBL.GetValue('BOP_DEVISE');
   if TTCSANSTVA then
      BEGIN
      {Cas du non assujetti}
      XP:=TOBL.GetValue('BOP_MONTANTTTC') ;
      if CodeD<>V_PGI.DevisePivot then XD:=TOBL.GetValue('BOP_MONTANTTTCDEV') else XD:=XP ;
      END else
      BEGIN
      XP:=TOBL.GetValue('BOP_MONTANTHT') ;
      if CodeD<>V_PGI.DevisePivot then XD:=TOBL.GetValue('BOP_MONTANTHTDEV') else XD:=XP ;
      END ;
   END else
if TOBL.NomTable='PIEDBASE' then
   BEGIN
   XP:=TOBL.GetValue('GPB_BASETAXE') ;
   CodeD:=TOBL.GetValue('GPB_DEVISE');
   if CodeD<>V_PGI.DevisePivot then XD:=TOBL.GetValue('GPB_BASEDEV') else XD:=XP ;
   END else
if TOBL.NomTable='PIEDPORT' then
   BEGIN
   CodeD:=TOBL.GetValue('GPT_DEVISE') ;
   if TTCSANSTVA then
      BEGIN
      {Cas du non assujetti ou facturation TTC}
      XP:=TOBL.GetValue('GPT_TOTALTTC') ;
      if CodeD<>V_PGI.DevisePivot then XD:=TOBL.GetValue('GPT_TOTALTTCDEV') else XD:=XP ;
      END else
      BEGIN
      XP:=TOBL.GetValue('GPT_TOTALHT') ;
      if CodeD<>V_PGI.DevisePivot then XD:=TOBL.GetValue('GPT_TOTALHTDEV') else XD:=XP ;
      END ;
   END else
    ;
END ;

Procedure RecupLesDCTaxes ( TOBL : TOB ; Var XD,XP : Double ) ;
Var CodeD : String ;
BEGIN
XD:=0 ; XP:=0 ;
XP:=TOBL.GetValue('GPB_VALEURTAXE') ;
CodeD:=TOBL.GetValue('GPB_DEVISE');
if CodeD<>V_PGI.DevisePivot then XD:=TOBL.GetValue('GPB_VALEURDEV') else XD:=XP ;
END ;

Procedure RecupLesDCEsc ( TOBL : TOB ; Var XD,XP : Double ) ;
Var CodeD : String ;
BEGIN
XD:=0 ; XP:=0 ;
XP:=TOBL.GetValue('GP_TOTALESC') ;
CodeD:=TOBL.GetValue('GP_DEVISE');
if CodeD<>V_PGI.DevisePivot then XD:=TOBL.GetValue('GP_TOTALESCDEV') else XD:=XP ;
END ;

Procedure RecupLesDCRem ( TOBL : TOB ; Var XD,XP : Double ) ;
Var CodeD : String ;
BEGIN
XD:=0 ; XP:=0 ;
XP:=TOBL.GetValue('GP_TOTALREMISE') ;
CodeD:=TOBL.GetValue('GP_DEVISE');
if CodeD<>V_PGI.DevisePivot then XD:=TOBL.GetValue('GP_TOTALREMISEDEV') else XD:=XP ;
END ;

procedure CompteGenVersEcr (TOBPiece,TOBE : TOB; Compte : string);
var QQ : Tquery;
begin
	if (TOBPiece.getValue('GP_TVAENCAISSEMENT')<>'TE') and (TOBPiece.getValue('GP_TVAENCAISSEMENT')<>'TM') and (TOBE.GetValue('E_TVAENCAISSEMENT')='X') then
  begin
    QQ := OpenSql ('SELECT G_TVASURENCAISS FROM GENERAUX WHERE G_GENERAL="'+Compte+'"',true,-1, '', True);
    if not QQ.eof then
    begin
      TOBE.PutValue('E_TVAENCAISSEMENT',QQ.FindField ('G_TVASURENCAISS').AsString);
    end;
    ferme (QQ);
  end;
end;

Procedure PieceVersECRSsTrait ( MM : RMVT ; TOBPiece,TOBClient,TOBE : TOB  ; Acc : boolean ) ;
Var
  RefI     : String ;
  DD       : TDateTime ;
  ExigeTVA : string;
BEGIN
  {RMVT}
  TOBE.PutValue('E_JOURNAL',MM.Jal)          ; TOBE.PutValue('E_EXERCICE',MM.Exo) ;
  TOBE.PutValue('E_DATECOMPTABLE',MM.DateC)  ; TOBE.PutValue('E_ETABLISSEMENT',MM.Etabl) ;
  TOBE.PutValue('E_DEVISE',MM.CodeD)         ; TOBE.PutValue('E_TAUXDEV',MM.TauxD) ;
  TOBE.PutValue('E_DATETAUXDEV',MM.DateTaux) ; TOBE.PutValue('E_QUALIFPIECE',MM.Simul) ;
  TOBE.PutValue('E_NATUREPIECE',MM.Nature)   ; TOBE.PutValue('E_NUMEROPIECE',MM.Num) ;
  TOBE.PutValue('E_VALIDE',CheckToString(MM.Valide)) ;
  TOBE.PutValue('E_DATEECHEANCE',MM.DateE) ;
  TOBE.PutValue('E_PERIODE',GetPeriode(MM.DateC)) ; TOBE.PutValue('E_SEMAINE',NumSemaine(MM.DateC)) ;
  {Pièce}
  TOBE.PutValue('E_REGIMETVA',TOBPiece.GetValue('GP_REGIMETAXE')) ;
  TOBE.PutValue('E_REFLIBRE',TOBPiece.GetValue('GP_REFINTERNE')) ;
  TOBE.PutValue('E_REFEXTERNE',TOBPiece.GetValue('GP_REFEXTERNE')) ;
  DD:=TOBPiece.GetValue('GP_DATEREFEXTERNE') ; if DD<=iDate1900 then DD:=TOBPiece.GetValue('GP_DATELIVRAISON') ;
  TOBE.PutValue('E_DATEREFEXTERNE',DD) ;
  TOBE.PutValue('E_UTILISATEUR',TOBPiece.GetValue('GP_UTILISATEUR')) ;
  TOBE.PutValue('E_DATECREATION',TOBPiece.GetValue('GP_DATECREATION')) ;
  TOBE.PutValue('E_DATEMODIF',TOBPiece.GetValue('GP_DATEMODIF')) ;
  TOBE.PutValue('E_COTATION',TOBPiece.GetValue('GP_COTATION')) ;
  // TOBE.PutValue('E_REFGESCOM',EncodeRefCPGescom(TOBPiece)) ;
  if TobPiece.GetValue('GP_GENERAUTO')<>'' then TOBE.PutValue('E_CREERPAR',TOBPiece.GetValue('GP_GENERAUTO'))
                                           else TOBE.PutValue('E_CREERPAR',TOBPiece.GetValue('GP_CREEPAR'));
  {Tiers}
  TOBE.PutValue('E_LIBELLE',TOBClient.GetValue('T_LIBELLE')) ;
  TOBE.PutValue('E_CONFIDENTIEL',TOBClient.GetValue('T_CONFIDENTIEL')) ;
  {Affaire}
  TOBE.PutValue('E_AFFAIRE',TOBPiece.GetValue('GP_AFFAIRE')) ;
  {Divers}
  TOBE.PutValue('E_QUALIFORIGINE','GC') ; TOBE.PutValue('E_VISION','DEM') ;
  TOBE.PutValue('E_ECRANOUVEAU','N')    ; TOBE.PutValue('E_MODESAISIE',MM.ModeSaisieJal) ;
  TOBE.PutValue('E_ETATLETTRAGE','RI')  ; TOBE.PutValue('E_CONTROLETVA','RIE')  ;
  //
  RefI:=RefI+TOBPiece.GetValue('GP_NATUREPIECEG')+' '+IntToStr(TOBPiece.GetValue('GP_NUMERO')) ;
  //
  AlimLibEcr(TobE,TobPiece,TOBClient,Copy(RefI,1,35),tecTiers,False,(MM.Simul='S'));
  TOBE.PutValue('E_ETAT','0000000000') ;
  {Tables libres}
  TOBE.PutValue('E_LIBREDATE',TOBPiece.GetValue('GP_DATELIBREPIECE1')) ;
  {#TVAENC}
  ExigeTVA := PositionneExige(TOBClient);
  TOBE.SetBoolean('E_TVAENCAISSEMENT',false);
  TOBE.SetString('E_TRESOSYNCHRO','RIE');
  TOBE.SetString('E_IO', 'X'); // Sur demande de la compta le 18/01/2006
  (* koaaa ca ??? ca va pas non ??
  if TOBPiece.GetValue('GP_TVAENCAISSEMENT')='TE' then StE:='X' else StE:='-' ;
  TOBE.PutValue('E_TVAENCAISSEMENT',StE) ;
  *)
  // CEGID V9
  TOBE.SetDateTime ('E_DATEORIGINE', iDate1900);
  TOBE.SetDateTime('E_DATPER', iDate1900);
  TOBE.SetInteger ('E_ENTITY', 0);
  TOBE.SetString ('E_REFGUID', '');
  // ----
end;

Procedure PieceVersECR ( MM : RMVT ; TOBPiece,TOBTiers,TOBE : TOB  ; Acc : boolean ) ;
Var RefI,NatEcr{,StE} : String ;
    DD              : TDateTime ;
    ExigeTVA : string;
BEGIN
  {RMVT}
  TOBE.PutValue('E_JOURNAL',MM.Jal)          ; TOBE.PutValue('E_EXERCICE',MM.Exo) ;
  TOBE.PutValue('E_DATECOMPTABLE',MM.DateC)  ; TOBE.PutValue('E_ETABLISSEMENT',MM.Etabl) ;
  TOBE.PutValue('E_DEVISE',MM.CodeD)         ; TOBE.PutValue('E_TAUXDEV',MM.TauxD) ;
  TOBE.PutValue('E_DATETAUXDEV',MM.DateTaux) ; TOBE.PutValue('E_QUALIFPIECE',MM.Simul) ;
  TOBE.PutValue('E_NATUREPIECE',MM.Nature)   ; TOBE.PutValue('E_NUMEROPIECE',MM.Num) ;
  TOBE.PutValue('E_VALIDE',CheckToString(MM.Valide)) ;
  TOBE.PutValue('E_DATEECHEANCE',MM.DateE) ;
  TOBE.PutValue('E_PERIODE',GetPeriode(MM.DateC)) ; TOBE.PutValue('E_SEMAINE',NumSemaine(MM.DateC)) ;
  {Pièce}
  TOBE.PutValue('E_REGIMETVA',TOBPiece.GetValue('GP_REGIMETAXE')) ;
  TOBE.PutValue('E_REFLIBRE',TOBPiece.GetValue('GP_REFINTERNE')) ;
  TOBE.PutValue('E_REFEXTERNE',TOBPiece.GetValue('GP_REFEXTERNE')) ;
  if VH_GC.GCIfDefCEGID then
  begin
    TOBE.PutValue('E_DATEREFEXTERNE',TOBPiece.GetValue('GP_DATEPIECE')) ;
    TOBE.PutValue('E_TABLE3',TOBPiece.GetValue('GP_REPRESENTANT')) ;
  end
  else
  begin
  DD:=TOBPiece.GetValue('GP_DATEREFEXTERNE') ; if DD<=iDate1900 then DD:=TOBPiece.GetValue('GP_DATELIVRAISON') ;
  TOBE.PutValue('E_DATEREFEXTERNE',DD) ;
  end;
  TOBE.PutValue('E_UTILISATEUR',TOBPiece.GetValue('GP_UTILISATEUR')) ;
  TOBE.PutValue('E_DATECREATION',TOBPiece.GetValue('GP_DATECREATION')) ;
  TOBE.PutValue('E_DATEMODIF',TOBPiece.GetValue('GP_DATEMODIF')) ;
  TOBE.PutValue('E_COTATION',TOBPiece.GetValue('GP_COTATION')) ;
  TOBE.PutValue('E_REFGESCOM',EncodeRefCPGescom(TOBPiece)) ;
  if TobPiece.GetValue('GP_GENERAUTO')<>'' then TOBE.PutValue('E_CREERPAR',TOBPiece.GetValue('GP_GENERAUTO'))
                                           else TOBE.PutValue('E_CREERPAR',TOBPiece.GetValue('GP_CREEPAR'));
  {Tiers}
  TOBE.PutValue('E_LIBELLE',TOBTiers.GetValue('T_LIBELLE')) ;
  TOBE.PutValue('E_CONFIDENTIEL',TOBTiers.GetValue('T_CONFIDENTIEL')) ;
  {Affaire}
  {$IFDEF CHR}
  TOBE.PutValue('E_AFFAIRE',TOBPiece.GetValue('GP_HRDOSSIER')) ;
  {$ELSE}
  TOBE.PutValue('E_AFFAIRE',TOBPiece.GetValue('GP_AFFAIRE')) ;
  {$ENDIF}
  {Divers}
  TOBE.PutValue('E_QUALIFORIGINE','GC') ; TOBE.PutValue('E_VISION','DEM') ;
  TOBE.PutValue('E_ECRANOUVEAU','N')    ; TOBE.PutValue('E_MODESAISIE',MM.ModeSaisieJal) ;
  TOBE.PutValue('E_ETATLETTRAGE','RI')  ; TOBE.PutValue('E_CONTROLETVA','RIE')  ;
  if Acc then
     BEGIN
     NatEcr:=TOBE.GetValue('E_NATUREPIECE') ;
     if ((NatEcr='OC') or (NatEcr='OF')) then RefI:='Acompte 'else RefI:='Règlement ' ;
     END else RefI:='' ;
  if VH_GC.GCIfDefCEGID then
     RefI:=IntToStr(TOBPiece.GetValue('GP_NUMERO'))
  else
    RefI:=RefI+TOBPiece.GetValue('GP_NATUREPIECEG')+' '+IntToStr(TOBPiece.GetValue('GP_NUMERO')) ;

  AlimLibEcr(TobE,TobPiece,TobTiers,Copy(RefI,1,35),tecTiers,False,(MM.Simul='S'));
  TOBE.PutValue('E_ETAT','0000000000') ;
  {Tables libres}
  TOBE.PutValue('E_LIBREDATE',TOBPiece.GetValue('GP_DATELIBREPIECE1')) ;
  {#TVAENC}
  ExigeTVA := PositionneExige(TOBTiers);
  TOBE.SetBoolean('E_TVAENCAISSEMENT',((ExigeTVA = 'TE') or (ExigeTVA = 'TM')));
  TOBE.SetString('E_IO', 'X'); // Sur demande de la compta le 18/01/2006
  (* koaaa ca ??? ca va pas non ??
  if TOBPiece.GetValue('GP_TVAENCAISSEMENT')='TE' then StE:='X' else StE:='-' ;
  TOBE.PutValue('E_TVAENCAISSEMENT',StE) ;
  *)
  // CEGID V9
  TOBE.SetDateTime ('E_DATEORIGINE', iDate1900);
  TOBE.SetDateTime('E_DATPER', iDate1900);
  TOBE.SetInteger ('E_ENTITY', 0);
  TOBE.SetString ('E_REFGUID', '');
  // ----
END ;

Function CreerCompteGC ( Var TOBG : TOB ; Cpt,CodeTVA,CodeTPF : String ; Quoi : T_CreatPont ; NatP : String = '' ) : boolean ;
Var LibG,NatG,Abr,SensG : String ;
BEGIN
TOBG:=TOB.Create('GENERAUX',Nil,-1) ;
TOBG.PutValue('G_GENERAL',Cpt) ;
TOBG.PutValue('G_CREERPAR','GC') ; TOBG.PutValue('G_SUIVITRESO','RIE') ;
Case Quoi of
   ccpEscompte : BEGIN
                 LibG:='Escompte (créé liaison comptable)' ;
                 Abr:='Escompte' ; NatG:='CHA' ; SensG:='M' ;
                 END ;
     ccpRemise : BEGIN
                 LibG:='R.R.R. (créé liaison comptable)' ;
                 Abr:='R.R.R.' ; NatG:='CHA' ; SensG:='M' ;
                 END ;
       ccpTaxe : BEGIN
                 if CodeTVA<>'' then BEGIN LibG:='TVA (créé liaison comptable)' ; Abr:='TVA' ; END
                                else BEGIN LibG:='TAXE (créé liaison comptable)' ; Abr:='TAXE' ; END ;
                 NatG:='DIV' ; SensG:='M' ;
                 END ;
         ccpHT : BEGIN
                 if ((NatP='FF') or (NatP='AF'))
                    then BEGIN LibG:='ACHAT (créé liaison comptable)' ; Abr:='ACHAT' ; NatG:='CHA' ; SensG:='D' ; END
                    else BEGIN LibG:='VENTE (créé liaison comptable)' ; Abr:='VENTE' ; NatG:='PRO' ; SensG:='C' ; END ;
                 END ;
      ccpStock : BEGIN
                 LibG:='Stock (créé liaison comptable)' ;
                 Abr:='Stock' ; NatG:='DIV' ; SensG:='M' ;
                 END ;
     ccpVarStk : BEGIN
                 LibG:='Var. Stock (créé liaison comptable)' ;
                 Abr:='Var. Stock' ; NatG:='CHA' ; SensG:='M' ;
                 END ;
       // Modif BTP
       ccpRG   : BEGIN
                 if ((NatP='FF') or (NatP='AF'))
                    then BEGIN LibG:='R.G. Ach.(créé liaison comptable)' ; Abr:='R.G Ach' ; NatG:='DIV' ; SensG:='C' ; END
                    else BEGIN LibG:='R.G. Vte.(créé liaison comptable)' ; Abr:='R.G Vte' ; NatG:='DIV' ; SensG:='D' ; END ;
                 END ;
       // ----
   END ;
TOBG.PutValue('G_NATUREGENE',NatG) ; TOBG.PutValue('G_LIBELLE',LibG) ;
TOBG.PutValue('G_TVA',CodeTVA) ; TOBG.PutValue('G_TPF',CodeTPF) ;
TOBG.PutValue('G_ABREGE',Abr) ;
// CEGID V9
TOBG.PutValue('G_GUIDASSOCIER','') ;
TOBG.PutValue('G_INVISIBLE','-') ;
TOBG.PutValue('G_NONTAXABLE','X') ;
TOBG.PutValue('G_RESTRICTIONA1','') ;
TOBG.PutValue('G_RESTRICTIONA2','') ;
TOBG.PutValue('G_RESTRICTIONA3','') ;
TOBG.PutValue('G_RESTRICTIONA4','') ;
TOBG.PutValue('G_RESTRICTIONA5','') ;
TOBG.PutValue('G_TYPECPTTVA',TypeCptTVADefaut(Cpt,NatG)) ;
// -----
Result:=TOBG.InsertDB(Nil) ;
if Not Result then BEGIN TOBG.Free ; TOBG:=Nil ; END ;
END ;

Function ListeChampsTOBG  : String ;
BEGIN
Result:='G_GENERAL, G_LIBELLE, G_ABREGE, G_NATUREGENE, G_CREERPAR, G_SUIVITRESO, G_TVA, G_TPF, G_CONSO, G_LETTRABLE, '
       +'G_POINTABLE, G_VENTILABLE, G_VENTILABLE1, G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5' ;
END ;

Function CreerTOBGeneral ( Cpt : String ) : TOB ;
Var TOBG : TOB ;
    Q    : TQuery ;
BEGIN
Result:=Nil ; TOBG:=Nil ;
if Cpt='' then Exit ;
Q:=OpenSQL('SELECT '+ListeChampsTOBG+' FROM GENERAUX WHERE G_GENERAL="'+Cpt+'"',True,-1, '', True) ;
if Not Q.EOF then
   BEGIN
   TOBG:=TOB.Create('LEGENERAL',Nil,-1) ;
   TOBG.SelectDB('',Q) ;
   END ;
Ferme(Q) ;
Result:=TOBG ;
END ;

Function CreerLigneEscompte ( TOBEcr,TOBPiece,TOBTiers : TOB ; MM : RMVT ) : T_RetCpta ;
Var TauxEsc   : Double ;
    CpteEsc   : String ;
    TOBG,TOBE : TOB ;
    OkVent    : boolean ;
    NumL      : integer ;
    DP,CP,DD,CD,XD,XP : Double ;
BEGIN
Result:=rcOk ;
TauxEsc:=TOBPiece.GetValue('GP_ESCOMPTE') ; if TauxEsc=0 then Exit ;
if TOBPiece.GetValue('GP_TOTALESCDEV')=0 then Exit ;
if ((MM.Nature='FC') or (MM.Nature='AC')) then CpteEsc:=VH_GC.GCCpteEscVTE else CpteEsc:=VH_GC.GCCpteEscACH ;
// Erreur sur Compte Escompte
if CpteEsc='' then BEGIN Result:=rcPar ; LastMsg:=1 ; Exit ; END ;
{Etude du compte général d'escompte}
TOBG:=CreerTOBGeneral(CpteEsc) ;
// Erreur sur Compte Escompte
if TOBG=Nil then
   BEGIN
   if VH_GC.GCPontComptable='ATT' then Result:=rcPar else
    if VH_GC.GCPontComptable='REF' then Result:=rcRef else
       BEGIN
       if Not CreerCompteGC(TOBG,CpteEsc,'','',ccpEscompte) then Result:=rcPar ;
       END ;
   if Result<>rcOk then BEGIN LastMsg:=1 ; Exit ; END ;
   END ;
OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
{Ligne d'écriture}
TOBE:=TOB.Create('ECRITURE',TOBEcr,-1) ;
PieceVersECR(MM,TOBPiece,TOBTiers,TOBE,False) ;
{Général}
TOBE.PutValue('E_GENERAL',CpteEsc) ;
TOBE.PutValue('E_CONSO',TOBG.GetValue('G_CONSO')) ;
{Divers}
TOBE.PutValue('E_TYPEMVT','HT') ;
AlimLibEcr(TobE,TobPiece,TobTiers,'Escompte '+StrfPoint(TauxEsc)+' %',tecRemise,True,(MM.Simul='S'));
{Contreparties}
TOBE.PutValue('E_CONTREPARTIEGEN',TOBTiers.GetValue('T_COLLECTIF')) ;
TOBE.PutValue('E_CONTREPARTIEAUX',TOBTiers.GetValue('T_AUXILIAIRE')) ; 
{Eche+Vent}
NumL:=TOBEcr.Detail.Count-NbEches+1 ; TOBE.PutValue('E_NUMLIGNE',NumL) ;
if OkVent then TOBE.PutValue('E_ANA','X') ;
{Montants}
DP:=0 ; CP:=0 ; DD:=0 ; CD:=0 ;
RecupLesDCEsc(TOBPiece,XD,XP) ;
if MM.Nature='FC' then BEGIN DD:=XD ; DP:=XP ; CD:=0 ; CP:=0 ; END else
if MM.Nature='FF' then BEGIN CD:=XD ; CP:=XP ; DD:=0 ; DP:=0 ; END else
if MM.Nature='AC' then BEGIN CD:=-XD ; CP:=-XP ; DD:=0 ; DP:=0 ; END else
if MM.Nature='AF' then BEGIN DD:=-XD ; DP:=-XP ; CD:=0 ; CP:=0 ; END ;
TOBE.PutValue('E_DEBIT',DP)     ; TOBE.PutValue('E_CREDIT',CP) ;
TOBE.PutValue('E_DEBITDEV',DD)  ; TOBE.PutValue('E_CREDITDEV',CD) ;
if ((DP=0) and (CP=0)) then TOBE.Free ;
TOBG.Free ;
END ;


Procedure CumulAnalPorcs ( CumHT : T_CodeCpta ; TOBCode : TOB ; XD,XP : Double ) ;
Var XDT : T_DetAnal ;
    TOBA,TOBAL : TOB ;
    Pourc,QteL : Double ;
    Section,Ax : String ;
    i,k,iTrouv : integer ;
BEGIN
XDT := nil;
if TOBCode = Nil then Exit;
if TOBCode.Detail.Count<=0 then Exit ;
TOBAL:=TOBCode.Detail[0] ;
QteL:=1 ;
for i:=0 to TOBAL.Detail.Count-1 do
    BEGIN
    TOBA:=TOBAL.Detail[i] ;
    if pos(TOBA.NomTable,'VENTIL;VENTEANA') = 0 then Continue;
    if TOBA.NomTable = 'VENTIL' then
    begin
			Ax := 'A'+Copy(TOBA.GetValue('V_NATURE'),3,1);
      Pourc:= TOBA.GetValue('V_TAUXMONTANT');
      Section:=TOBA.GetString('V_SECTION');
    end else
    begin
      Ax:=TOBA.GetValue('YVA_AXE') ; Pourc:=TOBA.GetValue('YVA_POURCENTAGE') ;
      Section:=TOBA.GetValue('YVA_SECTION') ;
    end;
    iTrouv:=-1 ;
    for k:=0 to CumHT.Anal.Count-1 do
        BEGIN
        XDT:=T_DetAnal(CumHT.Anal[k]) ;
        if ((XDT.Section=Section) and (XDT.Ax=Ax)) then BEGIN iTrouv:=k ; Break ; END ;
        END ;
    if iTrouv<0 then XDT:=T_DetAnal.Create ;
    XDT.Section:=Section ; XDT.aX:=aX ;
    XDT.MD:=XDT.MD+Pourc*XD/100 ; XDT.MP:=XDT.MP+Pourc*XP/100 ;
    XDT.Qte1:=Arrondi(XDT.Qte1+QteL*Pourc/100.0,V_PGI.OkDecQ) ;
    if iTrouv<0 then CumHT.Anal.Add(XDT) ;
    END ;
END ;



Function ChargeAjouteComptaRem ( TOBCpta,TOBPiece,TOBTiers,TOBAnaP,TOBAnaS : TOB; CpteVte,CpteAch : string ) : TOB ;
Var
  Nature,Etab,FamArt,FamTiers,FamAff,FamTaxe : String ;
  NatV,sRang : String ;
  TOBC : TOB ;
  VenteAchat : string;
BEGIN
//
	VenteAchat := GetInfoParPiece(TOBPiece.getValue('GP_NATUREPIECEG'),'GPP_VENTEACHAT');
	if VenteAchat='VEN' then NatV:='HV' else NatV:='HA' ;
  Etab:=TOBPiece.GetValue('GP_ETABLISSEMENT') ;
  Nature:=TOBPiece.GetValue('GP_NATUREPIECEG') ;
  FamTiers:=TOBTiers.GetValue('T_COMPTATIERS') ; FamArt:='' ;
  FamAff := '';
  FamTaxe := '';
  TOBC:=TOB.Create ('CODECPTA',TOBCPTA,-1);
  TOBC.SetInteger('GCP_RANG',9999);
  TOBC.SetString('GCP_COMPTAARTICLE','---');
  TOBC.SetString('GCP_COMPTATIERS','---');
  TOBC.SetString('GCP_COMPTAAFFAIRE','---');
  TOBC.SetString('GCP_ETABLISSEMENT',Etab);
  TOBC.SetString('GCP_REGIMETAXE','---');
  TOBC.SetString('GCP_CPTEGENEACH',CpteAch);
  TOBC.SetString('GCP_CPTEGENEVTE',CpteVte);
  TOBC.SetString('GCP_FAMILLETAXE','');
  TOBC.SetString('GCP_VENTEACHAT','');
  TOB.Create('Ecr',TOBC,-1) ; TOB.Create('Stk',TOBC,-1) ;
  TOBCpta.Detail.Sort('-GCP_COMPTAARTICLE;-GCP_COMPTATIERS;-GCP_COMPTAAFFAIRE;-GCP_REGIMETAXE;-GCP_FAMILLETAXE;-GCP_ETABLISSEMENT');
  //
  if TOBC.detail[0].detail.count = 0 then
  begin
    if GetParamSoc ('SO_GCAXEANALYTIQUE') then
    begin // mcd 20/11/02 mis ene place anamlytique générique
      GetVentilPort(TOBPiece,TOBTiers,TOBC,NatV,sRang)
    end else
    begin
      GetVentilAffaire(TOBPiece,nil,nil,TOBTiers,TOBC,NatV,sRang) ;
    end ;
  end;
  Result:=TOBC ;
END ;


Function CreerLigneRemise ( TOBEcr,TOBPiece,TOBTiers,TOBAnaP,TOBAnaS,TOBCpta : TOB ; MM : RMVT ) : T_RetCpta ;
Var TauxRem   : Double ;
    CPteRem,STa   : String ;
    TOBG,TOBE,TOBCode : TOB ;
    OkVent    : boolean ;
    NumL,kk,i      : integer ;
    DP,CP,DD,CD,XD,XP : Double ;
    LibEcr : string;
    CumHT                : T_CodeCpta ;
BEGIN
  Result:=rcOk ;
  TauxRem:=TOBPiece.GetValue('GP_REMISEPIED') ; if TauxRem=0 then Exit ;
  if TOBPiece.GetValue('GP_TOTALREMISEDEV')=0 then Exit ;
  if ((MM.Nature='FC') or (MM.Nature='AC')) then CpteRem:=VH_GC.GCCpteRemVTE else CpteRem:=VH_GC.GCCpteRemACH ;

  // Erreur sur Compte remise
  if CpteRem='' then BEGIN Result:=rcPar ; LastMsg:=2 ; Exit ; END ;
  LibEcr := 'Remise pied '+StrfPoint(TauxRem)+' %';
// pour pouvoir gérer l'analyique sur les remises --->
  TOBCode:=ChargeAjouteComptaRem(TOBCpta,TOBPiece,TOBTiers,TOBAnaP,TOBAnaS,VH_GC.GCCpteRemVTE,VH_GC.GCCpteRemACH) ;
// ----
  {Etude du compte général de remise}
  TOBG:=CreerTOBGeneral(CpteRem) ;
  // Erreur sur Compte remise
  if TOBG=Nil then
   BEGIN
   if VH_GC.GCPontComptable='ATT' then Result:=rcPar else
    if VH_GC.GCPontComptable='REF' then Result:=rcRef else
       BEGIN
       if Not CreerCompteGC(TOBG,CpteRem,'','',ccpRemise) then Result:=rcPar ;
       END ;
   if Result<>rcOk then BEGIN LastMsg:=2 ; Exit ; END ;
   END ;
  OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
  {Ligne d'écriture}
  TOBE:=TOB.Create('ECRITURE',TOBEcr,-1) ;
  PieceVersECR(MM,TOBPiece,TOBTiers,TOBE,False) ;
  {Général}
  TOBE.PutValue('E_GENERAL',CpteRem) ;
  TOBE.PutValue('E_CONSO',TOBG.GetValue('G_CONSO')) ;
  {Divers}
  TOBE.PutValue('E_TYPEMVT','HT') ;
  AlimLibEcr(TobE,TobPiece,TobTiers,'Remise pied '+StrfPoint(TauxRem)+' %',tecRemise,True,(MM.Simul='S'));
  {Contreparties}
  TOBE.PutValue('E_CONTREPARTIEGEN',TOBTiers.GetValue('T_COLLECTIF')) ;
  TOBE.PutValue('E_CONTREPARTIEAUX',TOBTiers.GetValue('T_AUXILIAIRE')) ;
  {Eche+Vent}
  NumL:=TOBEcr.Detail.Count-NbEches+1 ; TOBE.PutValue('E_NUMLIGNE',NumL) ;
  if OkVent then TOBE.PutValue('E_ANA','X') ;
  {Montants}
  DD:=0 ; CD:=0 ; DP:=0 ; CP:=0 ;
  RecupLesDCRem(TOBPiece,XD,XP) ;
  if MM.Nature='FC' then BEGIN DD:=XD ; DP:=XP ; CD:=0 ; CP:=0 ; END else
  if MM.Nature='FF' then BEGIN CD:=XD ; CP:=XP ; DD:=0 ; DP:=0 ; END else
  if MM.Nature='AC' then BEGIN CD:=-XD ; CP:=-XP ; DD:=0 ; DP:=0 ; END else
  if MM.Nature='AF' then BEGIN DD:=-XD ; DP:=-XP ; CD:=0 ; CP:=0 ; END ;
  TOBE.PutValue('E_DEBIT',DP)     ; TOBE.PutValue('E_CREDIT',CP) ;
  TOBE.PutValue('E_DEBITDEV',DD)  ; TOBE.PutValue('E_CREDITDEV',CD) ;
  if ((DP=0) and (CP=0)) then TOBE.Free ;
  CumHT := T_CodeCpta.Create ;
  if ((DP<>0) or (CP<>0)) then
  begin
  // -- Ajout analytique --
    CumHT.CptHT:=CpteRem ;
    CumHT.LibU:=True ;
    CumHT.LibArt:=LibEcr ;
    CumHT.Affaire:=TOBPiece.getString('GP_AFFAIRE') ;
    for kk:=1 to 5 do
      CumHT.FamTaxe[kk]:='' ;
    CumHT.XD:=CumHT.XD+XD ;
    CumHT.XP:=CumHT.XP+XP ;
    CumHT.Qte:=CumHT.Qte+1 ;
    if CumHT.LibArt<>LibEcr then
    begin
      CumHT.LibU:=False ;
      CumHT.LibArt:='' ;
    end;
    for kk:=1 to 5 do
    BEGIN
      CumHT.SommeTaxeD[kk]:=0 ;
      CumHT.SommeTaxeP[kk]:=0 ;
    END ;
    CumulAnalPorcs(CumHT,TOBCode,XD,XP) ;
    TTA.Add(CumHT) ;
  end;
  if CumHT.Anal.Count>0 then
  begin
     TOBE.AddChampSup('ANAL',False) ; TOBE.AddChampSup('AXES',False) ;
     TOBE.PutValue('ANAL',TTA.count-1) ;
     Sta:='' ; for i:=1 to 5 do if TOBG.GetValue('G_VENTILABLE'+IntToStr(i))='X' then Sta:=Sta+IntToStr(i) ;
     TOBE.PutValue('AXES',Sta) ;
  end;
  TOBG.Free ;
END ;


function GetCompteTaxe ( CatTaxe,regime,Famtaxe : string; Achat : boolean; TvaEnc : boolean; var CodeTva : string; var CodeTPF : string) : string;
begin
	result:= '';
    // ---
{$IFDEF BTP}
	if BTTypeTaxe (CatTaxe) = bttTVA then
{$ELSE}
	if CatTaxe='TX1' then
{$ENDIF}
	BEGIN
  	if TvaEnc then result :=TVA2ENCAIS(Regime,FamTaxe,Achat) ;
    if result='' then result:=TVA2CPTE(Regime,FamTaxe,Achat) ;
    CodeTVA:=FamTaxe ;
  END else
  BEGIN
  	if TvaEnc then result:=TPF2ENCAIS(Regime,FamTaxe,Achat) ;
    if result='' then result:=TPF2CPTE(Regime,FamTaxe,Achat) ;
    CodeTPF:=FamTaxe ;
  END ;
end;

Function CreerLignesTaxes ( TOBEcr,TOBPiece,TOBBases,TOBBasesST,TOBBasesCharge,TOBTiers,TOBPieceRG,TOBBasesRg,TOBPOrcs  : TOB ; MM : RMVT ) : T_RetCpta ;

Var i : integer ;
    TOBB,TOBE,TOBG,TOBTTC,TOBTVAAUTOLIQUIDE,TOBTL,TOBTVAAUTOLIQENC : TOB ;
    XD,XP : Double ;
    CD,CP,DD,DP,EP,ED,TRD,TRP : Double ;
    CTX,Regime,CatTaxe,FamTaxe,CodeTVA,CodeTPF : String ;
    Achat,OkVent,TvaEnc : boolean ;
    NumL  : integer ;
    // Modif BTP
    RGP,RGD,TheMontantRGTTC : double;
    OkPourMixte,IsEncaiss,EnHt : boolean;
    IndexMixte : string;
    TOBTMP : TOB;
    //

  function RecalcTva (TOBpiece,TOBB : TOB) : boolean;
  var regimeTaxe,FamilleTaxe : string;
    	BaseDev ,Base,Taux : double;
      EnHt : Boolean;
      TOBT : TOB;
      FormuleTaxe : string;
  begin
		Result := false;
  	RegimeTaxe := TOBPiece.GetString('GP_REGIMETAXE');
    Familletaxe := TOBB.GetString('GPB_FAMILLETAXE');
    BaseDev := TOBB.Getdouble('GPB_BASEDEV');
    Base := TOBB.Getdouble('GPB_BASETAXE');
    EnHt := (TOBPIece.GetString('GP_FACTUREHT')='X');
    TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],['TX1',RegimeTaxe,FamilleTaxe],False) ;
    if TOBT<>Nil then
    BEGIN
      Result := true;
      if Achat then
      begin
      	Taux:=TOBT.GetValue('TV_TAUXACH')
      end else
      begin
      	Taux:=TOBT.GetValue('TV_TAUXVTE')
      end;
      FormuleTaxe:=TOBT.GetValue('TV_FORMULETAXE') ;
      if EnHT then
      begin
        TOBB.SetDouble('GPB_VALEURDEV',	Arrondi(CalculeMontantTaxe(BaseDev,Taux,FormuleTaxe,nil),2));
        TOBB.SetDouble('GPB_VALEURTAXE',	Arrondi(CalculeMontantTaxe(Base,Taux,FormuleTaxe,nil),2));
      end else
      begin
        TOBB.SetDouble('GPB_VALEURDEV',	Arrondi(CalculeMontantTaxeTTC(BaseDev,Taux,FormuleTaxe,nil),2)) ;
        TOBB.SetDouble('GPB_VALEURTAXE',	Arrondi(CalculeMontantTaxeTTC(Base,Taux,FormuleTaxe,nil),2));
      end;

    END ;
	end;

  procedure GenereEcriture (LibelleForce : string='');
  begin
    {Ligne d'écriture}
    TOBTTC:=Nil ;
    if TOBEcr.Detail.Count>0 then TOBTTC:=TOBEcr.Detail[0] ;
    TOBE:=TOB.Create('ECRITURE',TOBEcr,-1) ;
    PieceVersECR(MM,TOBPiece,TOBTiers,TOBE,False) ;
    {Général}
    TOBE.PutValue('E_GENERAL',CTX) ;
    TOBE.PutValue('E_CONSO',TOBG.GetValue('G_CONSO')) ;
    {Divers}
    AlimLibEcr(TobE,TobPiece,TobTiers,TOBG.GetValue('G_LIBELLE'),tecTaxe,True,(MM.Simul='S'));
    if LibelleForce <> '' then
    begin
      TOBE.SetString('E_LIBELLE',LibelleForce);
    end;
    {$IFDEF BTP}
    if BTTypeTaxe (CatTaxe) = bttTVA then
    BEGIN
      TOBE.PutValue('E_TYPEMVT','TVA') ; if PremTVA<0 then PremTVA:=TOBEcr.Detail.Count ;
    END else if BTTypeTaxe (CatTaxe) = bttTPF then
    begin
      TOBE.PutValue('E_TYPEMVT','TPF');
    end else
    begin
      TOBE.PutValue('E_TYPEMVT',CatTaxe) ;
    end;
    {$ELSE}
    if CatTaxe='TX1' then
    BEGIN
      TOBE.PutValue('E_TYPEMVT','TVA') ; if PremTVA<0 then PremTVA:=TOBEcr.Detail.Count ;
    END else
    if CatTaxe='TX2' then TOBE.PutValue('E_TYPEMVT','TPF')
                     else TOBE.PutValue('E_TYPEMVT',CatTaxe) ;
    {$ENDIF}
    {Contreparties}
    TOBE.PutValue('E_CONTREPARTIEGEN',TOBTiers.GetValue('T_COLLECTIF')) ;
    TOBE.PutValue('E_CONTREPARTIEAUX',TOBTiers.GetValue('T_AUXILIAIRE')) ;
    TOBE.SetBoolean('E_TVAENCAISSEMENT', IsEncaiss);
    {Eche+Vent}
    NumL:=TOBEcr.Detail.Count-NbEches+1 ; TOBE.PutValue('E_NUMLIGNE',NumL) ;
    if OkVent then TOBE.PutValue('E_ANA',OKVent) ;
    if TOBG.GetValue('G_LETTRABLE')='X' then
    BEGIN
      TOBE.PutValue('E_ECHE','X') ; TOBE.PutValue('E_NUMECHE',1) ; TOBE.PutValue('E_ETATLETTRAGE','AL') ;
      TOBE.PutValue('E_ETAT','0000000000') ;
      if TOBTTC<>Nil then
      BEGIN
        TOBE.PutValue('E_MODEPAIE',TOBTTC.GetValue('E_MODEPAIE')) ;
        TOBE.PutValue('E_DATEECHEANCE',TOBTTC.GetValue('E_DATEECHEANCE')) ;
      END ;
    END ;
    {Montants}
    DP:=0 ; CP:=0 ; DD:=0 ; CD:=0 ;
    if MM.Nature='FC' then BEGIN CD:=XD-ED-RGD+TRD ; CP:=XP-EP-RGP+TRP ; DD:=0 ; DP:=0 ; END else
    if MM.Nature='FF' then BEGIN DD:=XD-ED-RGD ; DP:=XP-EP-RGP ; CD:=0 ; CP:=0 ; END else
    if MM.Nature='AC' then BEGIN DD:=-XD+ED+RGD-TRD ; DP:=-XP+EP+RGP-TRD ; CD:=0 ; CP:=0 ; END else
    if MM.Nature='AF' then BEGIN CD:=-XD+ED+RGD ; CP:=-XP+EP+RGP ; DD:=0 ; DP:=0 ; END ;
    CP := arrondi(CP,V_PGI.OkDecV); CD := Arrondi (CD,LaDev.decimale);
    DP := arrondi(DP,V_PGI.OkDecV); DD := Arrondi (DD,LaDev.decimale);
    if (DP < 0) or (CP <0) then
    begin
      DP := DP * (-1);
      CP := CP * (-1);
      DD := DD * (-1);
      CD := CD * (-1);
      //
      TOBE.PutValue('E_DEBIT',CP)     ; TOBE.PutValue('E_CREDIT',DP) ;
      TOBE.PutValue('E_DEBITDEV',CD)  ; TOBE.PutValue('E_CREDITDEV',DD) ;
    end else
    begin
      TOBE.PutValue('E_DEBIT',DP)     ; TOBE.PutValue('E_CREDIT',CP) ;
      TOBE.PutValue('E_DEBITDEV',DD)  ; TOBE.PutValue('E_CREDITDEV',CD) ;
    end;
    if ((DP=0) and (CP=0)) then TOBE.Free ;
  end;

  procedure RAZVariables;
  begin
    RGP := 0;
    RGD := 0;
    ED:= 0;
    EP:= 0;
    CodeTVA := '';
    CodeTPF := '';
  end;


  procedure EnregLiquidTva (TOBB : TOB);
  var Invert : boolean;
  begin
    RecupLesDCTaxes(TOBB,XD,XP) ;
    if XP=0 then Exit;
    CatTaxe:=TOBB.GetValue('GPB_CATEGORIETAXE') ; FamTaxe:=TOBB.GetValue('GPB_FAMILLETAXE') ;
    CodeTva := ''; CodeTPF := '';
    CTX := GetCompteTaxe ( CatTaxe,regime,Famtaxe,Achat,TvaEnc,CodeTva,CodeTPF);
    //
    // Erreur sur Compte Taxe
    if CTX='' then BEGIN Result:=rcPar ; LastMsg:=3 ; exit ; END ;

    {Etude du compte général de Taxe}
    TOBG:=CreerTOBgeneral(CTX) ;
    IsEncaiss := TvaEnc;

    // Erreur sur Compte Taxe
    if TOBG=Nil then
    BEGIN
      if VH_GC.GCPontComptable='ATT' then Result:=rcPar else
      if VH_GC.GCPontComptable='REF' then Result:=rcRef else
      BEGIN
        if Not CreerCompteGC(TOBG,CTX,CodeTVA,CodeTPF,ccpTaxe) then Result:=rcPar ;
      END ;
      if Result<>rcOk then BEGIN LastMsg:=3 ; exit ; END ;
    END ;

    OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;

    GenereEcriture;
    TOBG.Free ;
    //
    Invert := not Achat;
    XP := XP * (-1);
    XD := XD * (-1);
    //
    CodeTva := ''; CodeTPF := '';
    CTX := GetCompteTaxe ( CatTaxe,regime,Famtaxe,Invert,TvaEnc,CodeTva,CodeTPF);
    //
    // Erreur sur Compte Taxe
    if CTX='' then BEGIN Result:=rcPar ; LastMsg:=3 ; exit ; END ;

    {Etude du compte général de Taxe}
    TOBG:=CreerTOBgeneral(CTX) ;
    IsEncaiss := TvaEnc;

    // Erreur sur Compte Taxe
    if TOBG=Nil then
    BEGIN
      if VH_GC.GCPontComptable='ATT' then Result:=rcPar else
      if VH_GC.GCPontComptable='REF' then Result:=rcRef else
      BEGIN
        if Not CreerCompteGC(TOBG,CTX,CodeTVA,CodeTPF,ccpTaxe) then Result:=rcPar ;
      END ;
      if Result<>rcOk then BEGIN LastMsg:=3 ; exit ; END ;
    END ;
    OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
    GenereEcriture;
    TOBG.Free ;

  end;


  procedure EnregLiquidTvaST (TOBB : TOB);
  var Invert : boolean;
      Libelle : string;
  begin
    if TOBB.GetValue('GPB_VALEURTAXE')=0 then
    begin
      XP:=Arrondi(CalculeMontantTaxe(TOBB.GetValue('GPB_BASETAXE'),TOBB.GetValue('GPB_TAUXTAXE'),'',nil),V_PGI.OkDecV);
      XD := Arrondi(CalculeMontantTaxe(TOBB.GetValue('GPB_BASEDEV'),TOBB.GetValue('GPB_TAUXTAXE'),'',nil),LaDEV.decimale);
    end else
    begin
    RecupLesDCTaxes(TOBB,XD,XP) ;
    end;
    if XP=0 then Exit;
    CatTaxe:=TOBB.GetValue('GPB_CATEGORIETAXE') ; FamTaxe:=TOBB.GetValue('GPB_FAMILLETAXE') ;
    CodeTva := ''; CodeTPF := '';
    CTX := GetCompteTaxe ( CatTaxe,regime,Famtaxe,Achat,TvaEnc,CodeTva,CodeTPF);
    //
    // Erreur sur Compte Taxe
    if CTX='' then BEGIN Result:=rcPar ; LastMsg:=3 ; exit ; END ;

    {Etude du compte général de Taxe}
    TOBG:=CreerTOBgeneral(CTX) ;
    IsEncaiss := TvaEnc;

    // Erreur sur Compte Taxe
    if TOBG=Nil then
    BEGIN
      if VH_GC.GCPontComptable='ATT' then Result:=rcPar else
      if VH_GC.GCPontComptable='REF' then Result:=rcRef else
      BEGIN
        if Not CreerCompteGC(TOBG,CTX,CodeTVA,CodeTPF,ccpTaxe) then Result:=rcPar ;
      END ;
      if Result<>rcOk then BEGIN LastMsg:=3 ; exit ; END ;
    END ;

    OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
    Libelle := 'TVA Collectée (Autoliquidée)';
    GenereEcriture (Libelle);
    TOBG.Free ;
    //
    Invert := not Achat;
    XP := XP * (-1);
    XD := XD * (-1);
    //
    CodeTva := ''; CodeTPF := '';
    CTX := GetCompteTaxe ( CatTaxe,regime,Famtaxe,Invert,TvaEnc,CodeTva,CodeTPF);
    //
    // Erreur sur Compte Taxe
    if CTX='' then BEGIN Result:=rcPar ; LastMsg:=3 ; exit ; END ;

    {Etude du compte général de Taxe}
    TOBG:=CreerTOBgeneral(CTX) ;
    IsEncaiss := TvaEnc;

    // Erreur sur Compte Taxe
    if TOBG=Nil then
    BEGIN
      if VH_GC.GCPontComptable='ATT' then Result:=rcPar else
      if VH_GC.GCPontComptable='REF' then Result:=rcRef else
      BEGIN
        if Not CreerCompteGC(TOBG,CTX,CodeTVA,CodeTPF,ccpTaxe) then Result:=rcPar ;
      END ;
      if Result<>rcOk then BEGIN LastMsg:=3 ; exit ; END ;
    END ;

    OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
    Libelle := 'TVA Déductible (Autoliquidée)';
    GenereEcriture(LIbelle);
    TOBG.Free ;

  end;

BEGIN
  Result:=rcOk ;
  if TTCSANSTVA then Exit ;
  Regime:=TOBPiece.GetValue('GP_REGIMETAXE') ;
  Achat:=((MM.Nature='AF') or (MM.Nature='FF')) ;

  GetMontantRGReliquat(TOBPIeceRG, XD, XP,True);
  TheMontantRGTTC := XD;
  //
  {#TVAENC}
  TvaEnc:=(TOBPiece.GetValue('GP_TVAENCAISSEMENT')='TE') or (TOBPiece.GetValue('GP_TVAENCAISSEMENT')='TM') ;
  OkPourMixte := ((GereTvaMixte) and (TvaEnc));
  (*
  if (OkPourMixte) then
  begin
    for i:=0 to TobTVASurEncaiss.Detail.Count-1 do
    BEGIN
      TOBTmp:=TobTVASurEncaiss.Detail[i] ;
      TobTmp.Setdouble('MONTANT',Arrondi(TOBTmp.getDouble('BASE')*TOBTMP.GetDouble('TAUX'),2));
      TobTmp.Setdouble('MONTANTDEV',Arrondi(TOBTmp.getDouble('BASEDEV')*TOBTMP.GetDouble('TAUX'),2));
    END;
  end;
  *)
  for i:=0 to TOBBases.Detail.Count-1 do
  BEGIN
    // Modif BTP
    RGP:=0;RGD:=0;  TRD := 0; TRP := 0;
    // --
    CodeTVA:='' ; CodeTPF:='' ;
    TOBB:=TOBBases.Detail[i] ; RecupLesDCTaxes(TOBB,XD,XP) ;
    CatTaxe:=TOBB.GetValue('GPB_CATEGORIETAXE') ;
    FamTaxe:=TOBB.GetValue('GPB_FAMILLETAXE') ;
//    if (FamTaxe = VH_GC.AutoLiquiTVAST) and (Achat) and (TOBPiece.getValue('GP_DATEPIECE') >= StrToDate('01/01/2014')) then
    if (TOBPiece.getBoolean('GP_AUTOLIQUID')) and (Achat) and (TOBPiece.getValue('GP_DATEPIECE') >= StrToDate('01/01/2014')) then
    begin
      TOBTL := TOB.Create('PIEDBASE',nil,-1);
      TOBTL.Dupliquer(TOBB,False,true);
      if RecalcTva(TOBPiece,TOBTL) then
      begin
        EnregLiquidTva(TOBTL);
      end;
      TOBTL.Free;
      continue;
    end else if (Regime = 'INT') and (Achat) then
    begin
      TOBTL := TOB.Create('PIEDBASE',nil,-1);
      TOBTL.Dupliquer(TOBB,False,true);
      if RecalcTva(TOBPiece,TOBTL) then
      begin
        EnregLiquidTva(TOBTL);
    end;
      TOBTL.Free;
      continue;
    end;
    if XP=0 then Continue ;
    { Si Tva mixte, il faut déduire le montant présent dans TobTVASurEncaiss }
    if (OkPourMixte) then
    begin
      TobTmp := TobTVASurEncaiss.FindFirst(['INDEX'], [CatTaxe + ';' + FamTaxe], true);
      if assigned(TobTmp) then
      begin
        EP := TobTmp.Getdouble('MONTANT');
        ED := TobTmp.Getdouble('MONTANTDEV');
      end;
      // Si l'on déduit la Tva encaissement --> le reste ne l'est forcement pas
      TvaEnc := false;
    end;
    //
    CodeTva := ''; CodeTPF := '';
    CTX := GetCompteTaxe ( CatTaxe,regime,Famtaxe,Achat,TvaEnc,CodeTva,CodeTPF);
    //
    // Erreur sur Compte Taxe
    if CTX='' then BEGIN Result:=rcPar ; LastMsg:=3 ; Break ; END ;

    {Etude du compte général de Taxe}
    TOBG:=CreerTOBgeneral(CTX) ;
    IsEncaiss := TvaEnc;

    // Erreur sur Compte Taxe
    if TOBG=Nil then
    BEGIN
      if VH_GC.GCPontComptable='ATT' then Result:=rcPar else
      if VH_GC.GCPontComptable='REF' then Result:=rcRef else
      BEGIN
        if Not CreerCompteGC(TOBG,CTX,CodeTVA,CodeTPF,ccpTaxe) then Result:=rcPar ;
      END ;
      if Result<>rcOk then BEGIN LastMsg:=3 ; Break ; END ;
    END ;

    OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
    GetTvaRetenuesDivHT (TOBPorcs,TRP,TRD,CatTaxe);

    GenereEcriture;
    TOBG.Free ;
  END ;
  //
  for i:=0 to TOBBasesCharge.Detail.Count-1 do
  BEGIN
    // Modif BTP
    RGP:=0;RGD:=0;  TRD := 0; TRP := 0;
    // --
    CodeTVA:='' ; CodeTPF:='' ;
    TOBB:=TOBBasesCharge.Detail[i] ; RecupLesDCTaxes(TOBB,XD,XP) ;
    CatTaxe:=TOBB.GetValue('GPB_CATEGORIETAXE') ;
    FamTaxe:=TOBB.GetValue('GPB_FAMILLETAXE') ;
    if XP=0 then Continue ;
    { Si Tva mixte, il faut déduire le montant présent dans TobTVASurEncaiss }
    if (OkPourMixte) then
    begin
      TobTmp := TobTVASurEncaiss.FindFirst(['INDEX'], [CatTaxe + ';' + FamTaxe], true);
      if assigned(TobTmp) then
      begin
        EP := TobTmp.Getdouble('MONTANT');
        ED := TobTmp.Getdouble('MONTANTDEV');
      end;
      // Si l'on déduit la Tva encaissement --> le reste ne l'est forcement pas
      TvaEnc := false;
    end;
    //
    CodeTva := ''; CodeTPF := '';
    CTX := GetCompteTaxe ( CatTaxe,regime,Famtaxe,not Achat,TvaEnc,CodeTva,CodeTPF);
    //
    // Erreur sur Compte Taxe
    if CTX='' then BEGIN Result:=rcPar ; LastMsg:=3 ; Break ; END ;

    {Etude du compte général de Taxe}
    TOBG:=CreerTOBgeneral(CTX) ;
    IsEncaiss := TvaEnc;

    // Erreur sur Compte Taxe
    if TOBG=Nil then
    BEGIN
      if VH_GC.GCPontComptable='ATT' then Result:=rcPar else
      if VH_GC.GCPontComptable='REF' then Result:=rcRef else
      BEGIN
        if Not CreerCompteGC(TOBG,CTX,CodeTVA,CodeTPF,ccpTaxe) then Result:=rcPar ;
      END ;
      if Result<>rcOk then BEGIN LastMsg:=3 ; Break ; END ;
    END ;

    OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
    GenereEcriture;
    TOBG.Free ;
  END ;
  // Autoliquidation de tva a produire si ST avec TVA autoliquidée
  if (Result = rcOk) and (TOBBasesST <> nil) then
  begin
    for i:=0 to TOBBasesST.Detail.Count-1 do
    BEGIN
      // Modif BTP
      RGP:=0;RGD:=0;
      // --
      CodeTVA:='' ; CodeTPF:='' ;
      TOBB:=TOBBasesST.Detail[i] ;
      CatTaxe:=TOBB.GetValue('GPB_CATEGORIETAXE') ;
      FamTaxe:=TOBB.GetValue('GPB_FAMILLETAXE') ;
      EnregLiquidTvaST(TOBB);
    END ;
  end;
  { Si nécessaire, traite TobTVASurEncaiss }
  if (Result = rcOk) and (OkPourMixte) then
  begin
    { Force la récup de la TVA sur encaissement}
    TvaEnc := true;
    IsEncaiss := true;
    for i := 0 to TobTVASurEncaiss.detail.count -1 do
    begin
      RAZVariables;
      TOBB := TobTVASurEncaiss.detail[i];
      { Récup les montants }
      XD := TOBB.GetDouble('MONTANTDEV');
      XP := TOBB.GetDouble('MONTANT');
      if XP = 0 then continue;
      { Récup Catégorie et Famille }
      IndexMixte := TOBB.GetString('INDEX');
      CatTaxe := ReadTokenSt(IndexMixte);
      FamTaxe := IndexMixte;
      CTX := GetCompteTaxe ( CatTaxe,regime,Famtaxe,Achat,TvaEnc,CodeTva,CodeTPF);
      // Erreur sur Compte Taxe
      if CTX='' then BEGIN Result:=rcPar ; LastMsg:=3 ; Break ; END ;
      { Etude du compte général de Taxe }
      TOBG := CreerTOBgeneral(CTX);
      // Erreur sur Compte Taxe
      if TOBG=Nil then
      BEGIN
        if VH_GC.GCPontComptable='ATT' then Result:=rcPar else
        if VH_GC.GCPontComptable='REF' then Result:=rcRef else
        BEGIN
          if Not CreerCompteGC(TOBG,CTX,CodeTVA,CodeTPF,ccpTaxe) then Result:=rcPar ;
        END ;
        if Result<>rcOk then BEGIN LastMsg:=3 ; Break ; END ;
      END ;
      OkVent := TOBG.GetBoolean('G_VENTILABLE');
      {Ligne d'écriture}
      GenereEcriture;
      TOBG.Free ;
    end;
  end;
  //

END ;

Procedure DetruitEcheancesAcompte ( TOBPiece,TOBEches,TOBAcomptes : TOB ) ;
Var i,ie : integer ;
    TOBE : TOB ;
BEGIN
ie := 0;
if TOBAcomptes=Nil then Exit ;
if TOBAcomptes.Detail.Count<=0 then Exit ;
for i:=TOBEches.Detail.Count-1 downto 0 do
    BEGIN
    TOBE:=TOBEches.Detail[i] ;
    if TOBE.GetValue('GPE_ACOMPTE')='X' then TOBE.Free ;
    END ;
for i:=0 to TOBEches.Detail.Count-1 do
    BEGIN
    if i=0 then ie:=TOBEches.Detail[i].GetNumChamp('GPE_NUMECHE') ;
    TOBEches.Detail[i].PutValeur(ie,i+1) ;
    END ;
END ;

Function InsereEcheancesAcompte ( TOBPiece,TOBEches,TOBAcomptes : TOB ) : integer ;
Var X : Double ;
    TOBE,TOBACC : TOB ;
    i,NbCre,iTrouv,k,ie : integer ;
    MP  : String ;
BEGIN
TOBE := nil;
ie := 0;
Result:=0 ; NbCre:=0 ;
if TOBAcomptes=Nil then Exit ;
for i:=0 to TOBAcomptes.Detail.Count-1 do
    BEGIN
    TOBACC:=TOBAcomptes.Detail[i] ;
    MP:=TOBACC.GetValue('GAC_MODEPAIE') ; iTrouv:=-1 ;
    for k:=0 to NbCre-1 do
        BEGIN
        TOBE:=TOBEches.Detail[k] ;
        if TOBE.GetValue('GPE_MODEPAIE')=MP then BEGIN iTrouv:=i ; Break ; END ;
        END ;
    // forcer le regroupement des échéances d'acompte sur la première échéance
    // PCS 27102003 if VH_GC.GCIfDefCEGID then
    if ((VH_GC.GCIfDefCEGID) OR (not GetParamSoc('SO_BTCOMPTAREGL'))) and (TOBEches.Detail.Count>0) then
    begin
      iTrouv:=0 ; TOBE:=TOBEches.Detail[0] ;
    end;
    if iTrouv<0 then
       BEGIN
       TOBE:=TOB.Create('PIEDECHE',TOBEches,0) ; Inc(NbCre) ;
       TOBE.PutValue('GPE_NATUREPIECEG',TOBPiece.GetValue('GP_NATUREPIECEG')) ;
       TOBE.PutValue('GPE_SOUCHE',TOBPiece.GetValue('GP_SOUCHE')) ;
       TOBE.PutValue('GPE_NUMERO',TOBPiece.GetValue('GP_NUMERO')) ;
       TOBE.PutValue('GPE_INDICEG',TOBPiece.GetValue('GP_INDICEG')) ;
       TOBE.PutValue('GPE_TIERS',TOBPiece.GetValue('GP_TIERS')) ;
       TOBE.PutValue('GPE_DATEPIECE',TOBPiece.GetValue('GP_DATEPIECE')) ;
//       TOBE.PutValue('GPE_DATEECHE',TOBPiece.GetValue('GP_DATEPIECE')) ;
       TOBE.PutValue('GPE_DATEECHE',TOBACC.GetValue('GAC_DATEECHEANCE')) ;
       TOBE.PutValue('GPE_TAUXDEV',TOBPiece.GetValue('GP_TAUXDEV')) ;
       TOBE.PutValue('GPE_DATETAUXDEV',TOBPiece.GetValue('GP_DATETAUXDEV')) ;
       TOBE.PutValue('GPE_DEVISE',TOBPiece.GetValue('GP_DEVISE')) ;
       TOBE.PutValue('GPE_SAISIECONTRE',TOBPiece.GetValue('GP_SAISIECONTRE')) ;
       TOBE.PutValue('GPE_COTATION',TOBPiece.GetValue('GP_COTATION')) ;
       TOBE.PutValue('GPE_MODEPAIE',MP) ; TOBE.PutValue('GPE_ACOMPTE','X') ;
       END ;
    X:=TOBE.GetValue('GPE_MONTANTECHE') ; X:=Arrondi(X+TOBACC.GetValue('GAC_MONTANT'),6) ; TOBE.PutValue('GPE_MONTANTECHE',X) ;
    X:=TOBE.GetValue('GPE_MONTANTDEV')  ; X:=Arrondi(X+TOBACC.GetValue('GAC_MONTANTDEV'),6) ; TOBE.PutValue('GPE_MONTANTDEV',X) ;
    END ;
for i:=0 to TOBEches.Detail.Count-1 do
    BEGIN
    if i=0 then ie:=TOBEches.Detail[i].GetNumChamp('GPE_NUMECHE') ;
    TOBEches.Detail[i].PutValeur(ie,i+1) ;
    END ;
Result:=NbCre ;
END ;

procedure CalculRgEcheances (TOBpiece,TOBeches,TOBechesRG,TOBPieceRG,TOBBasesRG,TOBAcomptes:TOB);
var Indice : integer;
    Ratio,XD,XP,XXD,XXP,MontantTTC : double;
    TOBERG,TOBE: TOB;
    MontantCalc,Ecart : double;
    Numcaution : string;
begin
MontantCalc := 0;
//GetCumulRG(TOBPieceRG,XD,XP); // Valeur TTC de RG
GetRg(TOBPieceRG,False,True,XP,XD,numcaution);

GetSommeAcomptes(TOBAcomptes,XXP,XXD) ;
MontantTTC:=TOBPiece.GetValue('GP_TOTALTTCDEV')-XD-XXD ;
for Indice := 0 to TOBEches.detail.count -1 do
    begin
    TOBE := TOBEches.detail[Indice];
    TOBERG:=TOB.create ('PIEDECHE',TOBEchesRG,-1);
    TOBERG.Dupliquer (TOBE,true,true);
    {Calcul du ratio de l'echeance par rapport au montant TTC du document}
    if MontantTTC<>0 then Ratio := TOBERG.GetValue('GPE_MONTANTDEV')/MontantTTC else ratio:=0;
    TOBERG.PutValue('GPE_MONTANTDEV',arrondi(XD * Ratio,LaDev.Decimale ));
    MontantCalc:=MontantCalc + TOBERG.GetValue('GPE_MONTANTDEV');
    end;
{Reajustement des montants echeances RG par rapport au total RG}
Ecart := arrondi(MontantCalc-XD,LAdev.Decimale );
if Ecart <>0  then
   begin
   TOBERG := TOBEchesRG.detail[0];
   TOBERG.PutValue('GPE_MONTANTDEV',TOBERG.GetValue('GPE_MONTANTDEV') + Ecart);
   end;
{Reajustement des contrevaleur et pivot par rapport au montant devise}
for Indice := 0 to TOBEchesRG.detail.count - 1 do
    begin
    TOBERG := TOBECHESRG.detail[Indice];
    CoordinateMont (TOBERG,'GPE_MONTANTDEV','GPE_MONTANTECHE',LaDev);
    end;
end;

procedure RecupRGEcheances (TOBH,TOBEchesRG : TOB; var XXD,XXP : double);
var TOBERG : TOB;
begin
XXD:=0; XXP:=0;
if TOBEChesRG = nil then exit;
TOBERG := TOBEchesRG.FindFirst (['GPE_NUMECHE'],[TOBH.GetValue('GPE_NUMECHE')],true);
if TOBERG <> nil then
   begin
   XXD := TOBERG.GetValue('GPE_MONTANTDEV');
   XXP := TOBERG.GetValue('GPE_MONTANTECHE');
   end;
end;

Function CreerLignesRgtCaisse(TobEcr, TobPiece, TobEches, TobTiers : TOB; MM : RMVT) : T_RetCpta;
var TobRgt, TobGene : TOB;
    Journal, JournalOld(*, Nature *): string;
    Numero : Longint;
    //NewNum, AxeParam : boolean;
    Qry : TQuery;
    Cpt : integer;

    procedure CreerLigneEcr(IsTiers : boolean);
    var Montant : double;
    begin
      If LaDev.Taux = 0 then
        LaDev.Taux := GetTaux(LaDev.Code, LaDev.DateTaux, TobPiece.GetValue('GP_DATEPIECE'));
      Montant := TobEches.detail[Cpt].GetValue('GPE_MONTANTECHE');
      TobRgt := TOB.Create('ECRITURE',TobEcr,-1);
      PieceVersECR(MM,TOBPiece,TOBTiers,TobRgt,False);
      Qry := OpenSql('SELECT MP_JALCAISSE, J_CONTREPARTIE '+
                     'FROM MODEPAIE LEFT JOIN JOURNAL ON J_JOURNAL=MP_JALCAISSE '+
                     'WHERE MP_MODEPAIE = "'+TobEches.detail[Cpt].GetValue('GPE_MODEPAIE')+'"',True,-1, '', True);
      Journal := Qry.FindField('MP_JALCAISSE').AsString;
      if Journal <> JournalOld then
      begin
        JournalOld := Journal;
        Numero := GetNewNumJal(Journal,False,MM.DateC);
      end;
      TobRgt.PutValue('E_JOURNAL',Journal);
      TobRgt.PutValue('E_NUMEROPIECE',Numero);
      if IsTiers then
      begin
        TobRgt.PutValue('E_GENERAL',TobTiers.GetValue('T_COLLECTIF'));
        TobRgt.PutValue('E_AUXILIAIRE',TobTiers.GetValue('T_AUXILIAIRE'));
        TobRgt.PutValue('E_CONTREPARTIEAUX','');
        TobRgt.PutValue('E_ETATLETTRAGE','AL');
        TobRgt.PutValue('E_EMETTEURTVA','X');
        CSetMontants(TobRgt, 0, Montant, LaDev, False);
      end else
      begin
        TobRgt.PutValue('E_GENERAL',Qry.FindField('J_CONTREPARTIE').AsString);
        TobRgt.PutValue('E_AUXILIAIRE','');
        TobRgt.PutValue('E_CONTREPARTIEAUX',TobTiers.GetValue('T_AUXILIAIRE'));
        TobRgt.PutValue('E_ETATLETTRAGE','RI');
        TobRgt.PutValue('E_EMETTEURTVA','-');
        CSetMontants(TobRgt, Montant, 0, LaDev, False);
      end;
      Ferme(Qry);
      TobGene := CreerTOBGeneral(TobRgt.GetValue('E_GENERAL'));
      if TobGene.GetValue('G_VENTILABLE')='X' then
        TobRgt.PutValue('E_ANA','X')
        else
        TobRgt.PutValue('E_ANA','-');
      TobRgt.PutValue('E_CONSO',TobGene.GetValue('G_CONSO'));
      TobRgt.PutValue('E_BLOCNOTE',TobPiece.GetValue('GP_BLOCNOTE'));
      TobRgt.PutValue('E_RIB',TobTiers.GetValue('RIB')); // JT eQualité 11032
      TobRgt.PutValue('E_TYPEMVT','DIV');
      TobRgt.PutValue('E_ECHE','X');
      TobRgt.PutValue('E_MODEPAIE',TobEches.detail[Cpt].GetValue('GPE_MODEPAIE'));
      TobRgt.PutValue('E_DATEECHEANCE',TobPiece.GetValue('GP_DATEPIECE'));
      TobRgt.PutValue('E_CODEACCEPT',MPTOACC(TobRgt.GetValue('E_MODEPAIE')));
      AlimLibEcr(TobRgt,TobPiece,TobTiers,'',tecTiers,False,(MM.Simul='S'));
      if TobGene <> nil then
        FreeAndNil(TobGene);
    end;

begin
  Result := rcOk;
  Journal := '';
  JournalOld := '';
  Numero := 0;
//  AxeParam := (GetParamSoc ('SO_GCAXEANALYTIQUE'));
  if TobPiece.GetValue('GP_VENTEACHAT') = 'VEN' then
    MM.Nature := 'RC'
    else
    MM.Nature := 'RF';
  for Cpt := 0 to TobEches.detail.count -1 do
  begin
    CreerLigneEcr(True);  //Ligne Tiers
    CreerLigneEcr(False); //Ligne Banque
    if Not InsertionDifferee(TobEcr) then
    begin
      TobEcr.Free;
      Result := rcMaj;
      V_PGI.IoError:=oeUnknown ;
      Exit ;
    end;
    TobEcr.ClearDetail;
  end;
end;

function CreerLigneTimbres ( TOBEcr,TOBECHE,TOBTTC,TOBpiece,TOBTiers : TOB; var MM : RMVT) :  T_RetCpta;
var
  TOBPar,TOBE,TOBG : TOB;
  CTX : string;
  OkVent,TvAENc : Boolean;
  XP,XD, DP, CP, DD,CD: double ;
  Numl : Integer;

  procedure GenereEcriture;
  begin
    {Ligne d'écriture}
    TOBE:=TOB.Create('ECRITURE',TOBEcr,-1) ;
    PieceVersECR(MM,TOBPiece,TOBTiers,TOBE,False) ;
    {Général}
    TOBE.PutValue('E_GENERAL',CTX) ;
    TOBE.PutValue('E_CONSO',TOBG.GetValue('G_CONSO')) ;
    {Divers}
    AlimLibEcr(TobE,TobPiece,TobTiers,TOBG.GetValue('G_LIBELLE'),tecTaxe,True,(MM.Simul='S'));
    {$IFDEF BTP}
    TOBE.PutValue('E_TYPEMVT','DIV') ;
    {$ENDIF}
    {Contreparties}
    TOBE.PutValue('E_CONTREPARTIEGEN',TOBTiers.GetValue('T_COLLECTIF')) ;
    TOBE.PutValue('E_CONTREPARTIEAUX',TOBTiers.GetValue('T_AUXILIAIRE')) ;
    TOBE.SetBoolean('E_TVAENCAISSEMENT', tvaenc);
    {Eche+Vent}
    NumL:=TOBEcr.Detail.Count ; TOBE.PutValue('E_NUMLIGNE',NumL) ;
    if OkVent then TOBE.PutValue('E_ANA',OKVent) ;
    if TOBG.GetValue('G_LETTRABLE')='X' then
    BEGIN
      TOBE.PutValue('E_ECHE','X') ; TOBE.PutValue('E_NUMECHE',TOBTTC.GetValue('E_NUMECHE')) ; TOBE.PutValue('E_ETATLETTRAGE','AL') ;
      TOBE.PutValue('E_ETAT','0000000000') ;
      if TOBTTC<>Nil then
      BEGIN
        TOBE.PutValue('E_MODEPAIE',TOBTTC.GetValue('E_MODEPAIE')) ;
        TOBE.PutValue('E_DATEECHEANCE',TOBTTC.GetValue('E_DATEECHEANCE')) ;
      end;
    END ;
    {Montants}
    DP:=0 ; CP:=0 ; DD:=0 ; CD:=0 ;
    if MM.Nature='FC' then BEGIN CD:=XD ; CP:=XP; DD:=0 ; DP:=0 ; END else
    if MM.Nature='FF' then BEGIN DD:=XD ; DP:=XP ; CD:=0 ; CP:=0 ; END else
    if MM.Nature='AC' then BEGIN DD:=-XD ; DP:=-XP ; CD:=0 ; CP:=0 ; END else
    if MM.Nature='AF' then BEGIN CD:=-XD; CP:=-XP ; DD:=0 ; DP:=0 ; END ;
    CP := arrondi(CP,V_PGI.OkDecV); CD := Arrondi (CD,LaDev.decimale);
    TOBE.PutValue('E_DEBIT',DP)     ; TOBE.PutValue('E_CREDIT',CP) ;
    TOBE.PutValue('E_DEBITDEV',DD)  ; TOBE.PutValue('E_CREDITDEV',CD) ;
    if ((DP=0) and (CP=0)) then TOBE.Free ;
  end;

begin
	Result:=rcOk ;
  TvaEnc:=(TOBPiece.GetValue('GP_TVAENCAISSEMENT')='TE') or (TOBPiece.GetValue('GP_TVAENCAISSEMENT')='TM') ;

	if TOBECHE.GetValue('GPE_CODE')='' then exit;
  TOBPar := FindTimbre( TOBECHE.GetValue('GPE_CODE'));
  if TOBPar = nil then
  begin

  end else
  begin
    RecupLesDC (TOBECHE,XD,Xp,True);  // pour les timbres
    CTX  := TOBPar.GetValue('BTP_GENERAL');
  	TOBG:=CreerTOBgeneral(CTX) ;
    OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
    //
    GenereEcriture;
    TOBG.Free;
  end;
end;

function CreerEcrZero(TOBEcr,TOBPiece,TOBTiers: TOB ; var MM : RMVT) : T_RetCpta;
Var TOBTTC,TOBG : TOB ;
    i,cpt,NumE,NumL : integer ;
    GColl,ModePaie : String ;
    OkVent : boolean ;
    DP,CP,DD,CD,XD,XP : Double ;
    // modif BTP
    TOBEchesRG : TOB;
    XXD,XXP : double;
    LastDate : TdateTime;
begin
  Result:=rcOk ; NbAcc:=0 ;
  {Etude du compte général collectif}
  GColl:=TOBTiers.GetValue('T_COLLECTIF') ;
  TOBG:=CreerTOBGeneral(GColl) ;
  {Erreur sur le collectif}
  if TOBG=Nil then BEGIN Result:=rcPar ; LastMsg:=4 ; Exit ; END ;
  {-- }
  cpt:=1;
  OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
  TOBTTC:=TOB.Create('ECRITURE',TOBEcr,-1) ;
  PieceVersECR(MM,TOBPiece,TOBTiers,TOBTTC,False) ;
  {Tiers}
  TOBTTC.PutValue('E_AUXILIAIRE',TOBTiers.GetValue('T_AUXILIAIRE')) ;
  TOBTTC.PutValue('E_GENERAL',GColl) ;
  TOBTTC.PutValue('E_CONTREPARTIEAUX',TOBTiers.GetValue('T_AUXILIAIRE')) ;
  TOBTTC.PutValue('E_CONTREPARTIEGEN',GColl) ;
  TOBTTC.PutValue('E_CONSO',TOBTiers.GetValue('T_CONSO')) ;
  {Pièce}
  TOBTTC.PutValue('E_BLOCNOTE',TOBPiece.GetValue('GP_BLOCNOTE')) ;
  TOBTTC.PutValue('E_RIB',TobTiers.GetValue('RIB')); // JT eQualité 11032
  {Divers}
  TOBTTC.PutValue('E_TYPEMVT','') ;
  TOBTTC.PutValue('E_ETATLETTRAGE','AL') ;
  TOBTTC.PutValue('E_EMETTEURTVA','X') ;
  AlimLibEcr(TobTTC,TobPiece,TobTiers,TOBTiers.GetValue('T_LIBELLE'),tecTiers,True,(MM.Simul='S'));
  {Eche+Vent}
  NumL:=1 ; NumE:=1;
  TOBTTC.PutValue('E_ECHE','X') ;
  if OkVent then TOBTTC.PutValue('E_ANA','X') ;
  TOBTTC.PutValue('E_NUMLIGNE',1) ; TOBTTC.PutValue('E_NUMECHE',1) ;
  {Echéances}
  ModePaie:= 'CHQ';
  TOBTTC.PutValue('E_MODEPAIE',ModePaie) ;
  TOBTTC.PutValue('E_DATEECHEANCE',TOBPiece.GetDateTime('GP_DATEPIECE'));
  LastDate := TOBPiece.GetDateTime('GP_DATEPIECE');
  TOBTTC.PutValue('E_CODEACCEPT',MPTOACC(ModePaie)) ;
  TOBTTC.PutValue('E_DATEPAQUETMIN',MM.DateC) ;
  TOBTTC.PutValue('E_DATEPAQUETMAX',MM.DateC) ;
  TOBTTC.PutValue('E_ENCAISSEMENT','ENC') ;
  {Montants}
  DP:=0 ; CP:=0 ; DD:=0 ; CD:=0 ;
  XD := 1.0; XP := 1.0;
  // --
  if MM.Nature='FC' then BEGIN DD:=XD ; DP:=XP ; CD:=0 ; CP:=0 ; END else
  if MM.Nature='FF' then BEGIN CD:=XD ; CP:=XP ; DD:=0 ; DP:=0 ; END else
  if MM.Nature='AC' then BEGIN CD:=-(XD) ; CP:=-(XP) ; DD:=0 ; DP:=0 ; END else
  if MM.Nature='AF' then BEGIN DD:=-(XD) ; DP:=-(XP) ; CD:=0 ; CP:=0 ; END ;
  TOBTTC.PutValue('E_DEBIT',DP)     ; TOBTTC.PutValue('E_CREDIT',CP) ;
  TOBTTC.PutValue('E_DEBITDEV',DD)  ; TOBTTC.PutValue('E_CREDITDEV',CD) ;
  TOBTTC.PutValue('E_ETATREVISION','-') ;
  {-- }
  cpt:=2;
  TOBTTC:=TOB.Create('ECRITURE',TOBEcr,-1) ;
  PieceVersECR(MM,TOBPiece,TOBTiers,TOBTTC,False) ;
  {Tiers}
  TOBTTC.PutValue('E_AUXILIAIRE',TOBTiers.GetValue('T_AUXILIAIRE')) ;
  TOBTTC.PutValue('E_GENERAL',GColl) ;
  TOBTTC.PutValue('E_CONTREPARTIEAUX',TOBTiers.GetValue('T_AUXILIAIRE')) ;
  TOBTTC.PutValue('E_CONTREPARTIEGEN',GColl) ;
  TOBTTC.PutValue('E_CONSO',TOBTiers.GetValue('T_CONSO')) ;
  {Pièce}
  TOBTTC.PutValue('E_BLOCNOTE',TOBPiece.GetValue('GP_BLOCNOTE')) ;
  TOBTTC.PutValue('E_RIB',TobTiers.GetValue('RIB')); // JT eQualité 11032
  {Divers}
  TOBTTC.PutValue('E_TYPEMVT','') ;
  TOBTTC.PutValue('E_ETATLETTRAGE','AL') ;
  TOBTTC.PutValue('E_EMETTEURTVA','X') ;
  AlimLibEcr(TobTTC,TobPiece,TobTiers,TOBTiers.GetValue('T_LIBELLE'),tecTiers,True,(MM.Simul='S'));
  {Eche+Vent}
  NumL:=2 ; NumE:=1 ;
  TOBTTC.PutValue('E_ECHE','X') ;
  if OkVent then TOBTTC.PutValue('E_ANA','X') ;
  TOBTTC.PutValue('E_NUMLIGNE',2) ; TOBTTC.PutValue('E_NUMECHE',1) ;
  {Echéances}
//  ModePaie:=TOBH.GetValue('GPE_MODEPAIE') ;
  ModePaie:= 'CHQ';
  TOBTTC.PutValue('E_MODEPAIE',ModePaie) ;
  TOBTTC.PutValue('E_DATEECHEANCE',TOBPiece.GetDateTime('GP_DATEPIECE'));
  LastDate := TOBPiece.GetDateTime('GP_DATEPIECE');
  TOBTTC.PutValue('E_CODEACCEPT',MPTOACC(ModePaie)) ;
  TOBTTC.PutValue('E_DATEPAQUETMIN',MM.DateC) ;
  TOBTTC.PutValue('E_DATEPAQUETMAX',MM.DateC) ;
  TOBTTC.PutValue('E_ENCAISSEMENT','DEC') ;
  TOBTTC.PutValue('E_ETATREVISION','-') ;
  {Montants}
  DP:=0 ; CP:=0 ; DD:=0 ; CD:=0 ;
  XD := 1.0; XP := 1.0;
  // --
  if MM.Nature='FC' then BEGIN DD:=0 ; DP:=0 ; CD:=XD ; CP:=XP ; END else
  if MM.Nature='FF' then BEGIN CD:=0 ; CP:=0 ; DD:=XD ; DP:=XP ; END else
  if MM.Nature='AC' then BEGIN CD:=0 ; CP:=0 ; DD:=(-XD) ; DP:=(-XP) ; END else
  if MM.Nature='AF' then BEGIN DD:=0 ; DP:=0 ; CD:=-(XD) ; CP:=-(XP) ; END ;
  TOBTTC.PutValue('E_DEBIT',DP)     ; TOBTTC.PutValue('E_CREDIT',CP) ;
  TOBTTC.PutValue('E_DEBITDEV',DD)  ; TOBTTC.PutValue('E_CREDITDEV',CD) ;
  // Stockage date d'échéance
  MM.LastDateEche := LastDate;

  TOBG.Free ;

end;

Function CreerPaiementDirect ( TOBEcr,TOBPiece,TOBEches,TOBTiers,TOBAcomptes,TOBpieceRG,TOBBasesRG,TOBPieceTrait,TOBPorcs : TOB; DEV : RDevise; var MM : RMVT ) : T_RetCpta ;
Var TOBTTC,TOBG,TOBH : TOB ;
    i,cpt,NumE,NumL : integer ;
    GColl,ModePaie : String ;
    OkVent : boolean ;
    DP,CP,DD,CD,XD,XP : Double ;
    // modif BTP
    TOBEchesRG : TOB;
    XXD,XXP : double;
    LastDate : TdateTime;
begin
  Result:=rcOk ;
  NbAcc:=0 ;
  XXD := 0;
  XXP := 0;
  {Etude du compte général collectif}
  GColl:= GetParamSocSecur('SO_BTREGLGENERAL','');
  //FV1 - 26/07/2016
  TOBG:=CreerTOBGeneral(GColl) ;
  {Erreur sur le collectif}
  if TOBG=Nil then BEGIN Result:=rcPar ; LastMsg:=4 ; Exit ; END ;
  OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
  {Boucle sur les echéances}
  cpt:=0;
  for i:=0 to TOBEches.Detail.Count-1 do
  BEGIN
    TOBH:=TOBEches.Detail[i] ;
    // Le règlement d'un article financier ne doit pas être traité.
    inc(cpt);
  	if (TOBH.GetString('GPE_FOURNISSEUR') = '') then continue;
    TOBTTC:=TOB.Create('ECRITURE',TOBEcr,-1) ;
    PieceVersECR(MM,TOBPiece,TOBTiers,TOBTTC,False) ;
    {Tiers}
    TOBTTC.PutValue('E_AUXILIAIRE','') ;
    TOBTTC.PutValue('E_GENERAL',GColl) ;
    TOBTTC.PutValue('E_CONSO',TOBTiers.GetValue('T_CONSO')) ;
    {Divers}
    TOBTTC.PutValue('E_TYPEMVT','TTC') ;
    TOBTTC.PutValue('E_ETATLETTRAGE','AL') ;
    if TOBG.GetValue('G_LETTRABLE')='X' then
    BEGIN
      TOBTTC.PutValue('E_ECHE','X') ; TOBTTC.PutValue('E_NUMECHE',1) ; TOBTTC.PutValue('E_ETATLETTRAGE','AL') ;
      TOBTTC.PutValue('E_ETAT','0000000000') ;
    END ;
    TOBTTC.PutValue('E_EMETTEURTVA','X') ;
    AlimLibEcr(TobTTC,TobPiece,TobTiers,TOBTiers.GetValue('T_LIBELLE'),tecTiers,True,(MM.Simul='S'));
    // --
  	if TOBH.GetString('GPE_FOURNISSEUR') <> '' then
    begin
      TOBTTC.SetString('E_LIBELLE',Copy(TOBTTC.getString('E_LIBELLE')+'-ST- '+TOBH.GetString('GPE_FOURNISSEUR'),1,35));
    end;
    // --
    {Eche+Vent}
    NumL:=TOBEcr.Detail.Count-NbEches+1 ; 
    TOBTTC.PutValue('E_NUMLIGNE',NumL) ;
    {Echéances}
    ModePaie:=TOBH.GetValue('GPE_MODEPAIE') ;
    TOBTTC.PutValue('E_MODEPAIE',ModePaie) ;
    TOBTTC.PutValue('E_DATEECHEANCE',TOBH.GetValue('GPE_DATEECHE')) ;
    LastDate := TOBH.GetValue('GPE_DATEECHE');
    TOBTTC.PutValue('E_CODEACCEPT',MPTOACC(ModePaie)) ;
    TOBTTC.PutValue('E_DATEPAQUETMIN',MM.DateC) ;
    TOBTTC.PutValue('E_DATEPAQUETMAX',MM.DateC) ;
    //FV1 - 23/02/2018 : FS#2964 - Viviane - Il faut charger le champ E_DATEVALEUR avec le champ E_DATECOMPTABLE pour la ligne « TTC »
    TOBTTC.PutValue('E_DATEVALEUR', MM.DateC) ;
    {Montants}
    DP:=0 ; CP:=0 ; DD:=0 ; CD:=0 ;
    RecupLesDC(TOBH,XD,XP) ;
    // Modif BTP
//    if (TOBpieceRG.detail.count>0) and (TOBH.GetValue('GPE_ACOMPTE')<>'X')then RecupRGEcheances (TOBH,TOBEchesRG,XXD,XXP) else begin XXD:=0;XXP:=0; end;
    // --
    if MM.Nature='FC' then BEGIN DD:=XD+XXD ; DP:=XP+XXP ; CD:=0 ; CP:=0 ; END else
    if MM.Nature='FF' then BEGIN CD:=XD+XXD ; CP:=XP+XXP ; DD:=0 ; DP:=0 ; END else
    if MM.Nature='AC' then BEGIN CD:=-(XD+XXD) ; CP:=-(XP+XXP) ; DD:=0 ; DP:=0 ; END else
    if MM.Nature='AF' then BEGIN DD:=-(XD+XXD) ; DP:=-(XP+XXP) ; CD:=0 ; CP:=0 ; END ;
    TOBTTC.PutValue('E_DEBIT',DP)     ; TOBTTC.PutValue('E_CREDIT',CP) ;
    TOBTTC.PutValue('E_DEBITDEV',DD)  ; TOBTTC.PutValue('E_CREDITDEV',CD) ;
    //FV1 - 23/04/2018 : FS#3043 - Viviane : renseigner le champ "E_ENCAISSEMENT" lors de la validation d'une facture
    TOBTTC.PutValue('E_ENCAISSEMENT', SensEnc(DP,CP)) ;
    //
  END ;

end;


Function CreerLignesTiers ( TOBEcr,TOBPiece,TOBEches,TOBTiers,TOBAcomptes,TOBpieceRG,TOBBasesRG,TOBPieceTrait,TOBPorcs,TOBVTECOLL : TOB; DEV : RDevise; var MM : RMVT ) : T_RetCpta ;
Var TOBTTC,TOBG,TOBH,TOBL : TOB ;
    i,II,cpt,NumE,NumL : integer ;
    GColl,ModePaie,Collectif : String ;
    OkVent : boolean ;
    DP,CP,DD,CD,XD,XP : Double ;
    // modif BTP
    TOBEchesRG : TOB;
    XXD,XXP : double;
    LastDate : TdateTime;
BEGIN
  //
  Result:=rcOk ;
  NbAcc:=0 ;
  {Etude du compte général collectif}
  GColl:=TOBTiers.GetValue('T_COLLECTIF') ;
  TOBG:=CreerTOBGeneral(GColl) ;
  // modif BTP
  TOBechesRG:=TOB.create('Les ECHEANCES',nil,-1);
  // --
  {Erreur sur le collectif}
  if TOBG=Nil then BEGIN Result:=rcPar ; LastMsg:=4 ; Exit ; END ;
  OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
  {Gestion de l'acompte}
  NbAcc:=InsereEcheancesAcompte(TOBPiece,TOBEches,TOBAcomptes) ;
  {Gestion de la retenue de garantie si non comptabilise}
  if (TOBpieceRG.detail.count>0) and (not RGComptabilise) then CalculRgEcheances(TOBPiece,TOBEches,TOBEchesRG,TOBPieceRG,TOBBasesRG,TOBAcomptes);
  {Boucle sur les echéances}
  LastDate := IDate1900;
  XXD      := 0;
  XXP      := 0;
  cpt:=0;
  for i:=0 to TOBEches.Detail.Count-1 do
  BEGIN
    TOBH:=TOBEches.Detail[i] ;
    // Le règlement d'un article financier ne doit pas être traité.
    {$IFDEF MODE}
    if (TestReglementArticleFi(TOBH)=true) then Continue;
    {$ENDIF}
  	if (TOBH.GetString('GPE_FOURNISSEUR') <> '') and (not GetparamSocSecur('SO_BTREGLSTTIERS',true)) then continue;
    if (((pos(MM.Nature,'FC;AC')>0) and (TOBH.GetString('GPE_FOURNISSEUR') = '')) or (pos(MM.Nature,'FF;AF')>0)) and (getparamSocSecur('SO_BTVENTCOLLECTIF',false)) and (TOBVTECOLL <> nil)  then
    begin
      ReajusteCollectifs (TOBVTECOLL,TOBH.GetDouble('GPE_MONTANTECHE'),TOBH.GetDouble('GPE_MONTANTDEV'));
      for II := 0 to TOBVTECOLL.detail.count -1 do
      begin
        inc(cpt);
        TOBl:= TOBVTECOLL.detail[II];
        TOBTTC:=TOB.Create('ECRITURE',TOBEcr,-1) ;
        PieceVersECR(MM,TOBPiece,TOBTiers,TOBTTC,False) ;
        {Tiers}
        Collectif := TOBl.getString('BPB_COLLECTIF'); if Collectif = '' then Collectif := GColl;
        TOBTTC.PutValue('E_AUXILIAIRE',TOBTiers.GetValue('T_AUXILIAIRE')) ;
        TOBTTC.PutValue('E_GENERAL',Collectif) ;
        TOBTTC.PutValue('E_CONSO',TOBTiers.GetValue('T_CONSO')) ;
        {Pièce}
        TOBTTC.PutValue('E_BLOCNOTE',TOBPiece.GetValue('GP_BLOCNOTE')) ;
        TOBTTC.PutValue('E_RIB',TobTiers.GetValue('RIB')); // JT eQualité 11032
        {Divers}
        TOBTTC.PutValue('E_TYPEMVT','TTC') ;
        TOBTTC.PutValue('E_ETATLETTRAGE','AL') ;
        TOBTTC.PutValue('E_EMETTEURTVA','X') ;
        AlimLibEcr(TobTTC,TobPiece,TobTiers,TOBTiers.GetValue('T_LIBELLE'),tecTiers,True,(MM.Simul='S'));
        if VH_GC.GCIfDefCEGID then
        if TOBTiers.FieldExists('LIBTIERSCEGID') then TOBTTC.PutValue('E_LIBELLE',TOBTiers.GetValue('LIBTIERSCEGID')) ;
        {Eche+Vent}
        if OkVent then BEGIN NumL:=cpt ; NumE:=1 ; END else BEGIN NumL:=1 ; NumE:=cpt ; END ;
        TOBTTC.PutValue('E_ECHE','X') ;
        if OkVent then TOBTTC.PutValue('E_ANA','X') ;
        TOBTTC.PutValue('E_NUMLIGNE',NumL) ; TOBTTC.PutValue('E_NUMECHE',NumE) ;
        {Echéances}
        ModePaie:=TOBH.GetValue('GPE_MODEPAIE') ;
        TOBTTC.PutValue('E_MODEPAIE',ModePaie) ;
        TOBTTC.PutValue('E_DATEECHEANCE',TOBH.GetValue('GPE_DATEECHE')) ;
        LastDate := TOBH.GetValue('GPE_DATEECHE');
        TOBTTC.PutValue('E_CODEACCEPT',MPTOACC(ModePaie)) ;
        TOBTTC.PutValue('E_DATEPAQUETMIN',MM.DateC) ;
        TOBTTC.PutValue('E_DATEPAQUETMAX',MM.DateC) ;
        //FV1 - 23/02/2018 : FS#2964 - Viviane - Il faut charger le champ E_DATEVALEUR avec le champ E_DATECOMPTABLE pour la ligne « TTC »
        TOBTTC.PutValue('E_DATEVALEUR', MM.DateC) ;
        {Montants}
        DP:=0 ; CP:=0 ; DD:=0 ; CD:=0 ;
        RecupLesDC(TOBL,XD,XP) ;
        if MM.Nature='FC' then BEGIN DD:=XD+XXD ; DP:=XP+XXP ; CD:=0 ; CP:=0 ; END else
        if MM.Nature='FF' then BEGIN CD:=XD+XXD ; CP:=XP+XXP ; DD:=0 ; DP:=0 ; END else
        if MM.Nature='AC' then BEGIN CD:=-(XD+XXD) ; CP:=-(XP+XXP) ; DD:=0 ; DP:=0 ; END else
        if MM.Nature='AF' then BEGIN DD:=-(XD+XXD) ; DP:=-(XP+XXP) ; CD:=0 ; CP:=0 ; END ;
        TOBTTC.PutValue('E_DEBIT',DP)     ; TOBTTC.PutValue('E_CREDIT',CP) ;
        TOBTTC.PutValue('E_DEBITDEV',DD)  ; TOBTTC.PutValue('E_CREDITDEV',CD) ;
        //FV1 - 23/04/2018 : FS#3043 - Viviane : renseigner le champ "E_ENCAISSEMENT" lors de la validation d'une facture
        TOBTTC.PutValue('E_ENCAISSEMENT', SensEnc(DP,CP)) ;
        //
        if ((DP=0) and (CP=0)) then TOBTTC.Free else
        BEGIN
          if OkVent then NbEches:=1 else Inc(NbEches) ;
        END ;

        // TRAITEMENT SPECIFIQUE POUR LA SOCIETE POUCHAIN : BRL Mars 2011
        // Test sur code SIREN unique pour l'entreprise (9 premiers caractères du code SIRET)
        if (TOBTTC.Detail.count <> 0) and (Copy(GetParamsocSecur('SO_SIRET',''),1,9) = '403001001') then
        BEGIN
          // Mise à jour de la table libre 1 ecriture à partir de la table libre pièce 3
          // pour gestion des factures en litiges
          TOBTTC.PutValue('E_TABLE0', TOBPiece.GetValue('GP_LIBREPIECE3'));
        END;
      end;
    end else
    begin
      inc(cpt); 
      TOBTTC:=TOB.Create('ECRITURE',TOBEcr,-1) ;
      PieceVersECR(MM,TOBPiece,TOBTiers,TOBTTC,False) ;
      {Tiers}
      TOBTTC.PutValue('E_AUXILIAIRE',TOBTiers.GetValue('T_AUXILIAIRE')) ;
      TOBTTC.PutValue('E_GENERAL',GColl) ;
      TOBTTC.PutValue('E_CONSO',TOBTiers.GetValue('T_CONSO')) ;
      {Pièce}
      TOBTTC.PutValue('E_BLOCNOTE',TOBPiece.GetValue('GP_BLOCNOTE')) ;
      TOBTTC.PutValue('E_RIB',TobTiers.GetValue('RIB')); // JT eQualité 11032
      {Divers}
      TOBTTC.PutValue('E_TYPEMVT','TTC') ;
      TOBTTC.PutValue('E_ETATLETTRAGE','AL') ;
      TOBTTC.PutValue('E_EMETTEURTVA','X') ;
      AlimLibEcr(TobTTC,TobPiece,TobTiers,TOBTiers.GetValue('T_LIBELLE'),tecTiers,True,(MM.Simul='S'));
      // --
      if TOBH.GetString('GPE_FOURNISSEUR') <> '' then
      begin
        TOBTTC.SetString('E_LIBELLE',Copy(TOBTTC.getString('E_LIBELLE')+'-ST- '+TOBH.GetString('GPE_FOURNISSEUR'),1,35));
      end;
      // --
      if VH_GC.GCIfDefCEGID then
      if TOBTiers.FieldExists('LIBTIERSCEGID') then TOBTTC.PutValue('E_LIBELLE',TOBTiers.GetValue('LIBTIERSCEGID')) ;
      {Eche+Vent}
      if OkVent then BEGIN NumL:=cpt ; NumE:=1 ; END else BEGIN NumL:=1 ; NumE:=cpt ; END ;
      TOBTTC.PutValue('E_ECHE','X') ;
      if OkVent then TOBTTC.PutValue('E_ANA','X') ;
      TOBTTC.PutValue('E_NUMLIGNE',NumL) ; TOBTTC.PutValue('E_NUMECHE',NumE) ;
      {Echéances}
      ModePaie:=TOBH.GetValue('GPE_MODEPAIE') ;
      TOBTTC.PutValue('E_MODEPAIE',ModePaie) ;
      TOBTTC.PutValue('E_DATEECHEANCE',TOBH.GetValue('GPE_DATEECHE')) ;
      LastDate := TOBH.GetValue('GPE_DATEECHE');
      TOBTTC.PutValue('E_CODEACCEPT',MPTOACC(ModePaie)) ;
      TOBTTC.PutValue('E_DATEPAQUETMIN',MM.DateC) ;
      TOBTTC.PutValue('E_DATEPAQUETMAX',MM.DateC) ;
      //FV1 - 23/02/2018 : FS#2964 - Viviane - Il faut charger le champ E_DATEVALEUR avec le champ E_DATECOMPTABLE pour la ligne « TTC »
      TOBTTC.PutValue('E_DATEVALEUR', MM.DateC) ;
      {Montants}
      DP:=0 ; CP:=0 ; DD:=0 ; CD:=0 ;
      RecupLesDC(TOBH,XD,XP) ;
      // Modif BTP
  //    if (TOBpieceRG.detail.count>0) and (TOBH.GetValue('GPE_ACOMPTE')<>'X')then RecupRGEcheances (TOBH,TOBEchesRG,XXD,XXP) else begin XXD:=0;XXP:=0; end;
      // --
      if MM.Nature='FC' then BEGIN DD:=XD+XXD ; DP:=XP+XXP ; CD:=0 ; CP:=0 ; END else
      if MM.Nature='FF' then BEGIN CD:=XD+XXD ; CP:=XP+XXP ; DD:=0 ; DP:=0 ; END else
      if MM.Nature='AC' then BEGIN CD:=-(XD+XXD) ; CP:=-(XP+XXP) ; DD:=0 ; DP:=0 ; END else
      if MM.Nature='AF' then BEGIN DD:=-(XD+XXD) ; DP:=-(XP+XXP) ; CD:=0 ; CP:=0 ; END ;
      TOBTTC.PutValue('E_DEBIT',DP)     ; TOBTTC.PutValue('E_CREDIT',CP) ;
      TOBTTC.PutValue('E_DEBITDEV',DD)  ; TOBTTC.PutValue('E_CREDITDEV',CD) ;
      //FV1 - 23/04/2018 : FS#3043 - Viviane : renseigner le champ "E_ENCAISSEMENT" lors de la validation d'une facture
      TOBTTC.PutValue('E_ENCAISSEMENT', SensEnc(DP,CP)) ;
      //
      if ((DP=0) and (CP=0)) then TOBTTC.Free else
      BEGIN
        if OkVent then NbEches:=1 else Inc(NbEches) ;
      END ;

      // TRAITEMENT SPECIFIQUE POUR LA SOCIETE POUCHAIN : BRL Mars 2011
      // Test sur code SIREN unique pour l'entreprise (9 premiers caractères du code SIRET)
      if (TOBTTC.Detail.count <> 0) and (Copy(GetParamsocSecur('SO_SIRET',''),1,9) = '403001001') then
      BEGIN
        // Mise à jour de la table libre 1 ecriture à partir de la table libre pièce 3
        // pour gestion des factures en litiges
        TOBTTC.PutValue('E_TABLE0', TOBPiece.GetValue('GP_LIBREPIECE3'));
      END;
      if TOBH.GetValue('GPE_CODE')<>'' then
      begin
        CreerLigneTimbres (TOBEcr,TOBH,TOBTTC,TOBPiece,TOBTiers,MM);
      END;
    end;
  END ;

  if (GetParamSocSecur ('SO_BTOPTRGCOLLTIE',false)) and (TOBPieceRg <> nil) and (TOBPieceRg.detail.count <>  0) then
  begin
    inc(cpt); // on rajoute une echeance pour la RG
    // nouvelle gestion des retenues de garantie pour eviter le probleme de controle de TVA
    TOBTTC:=TOB.Create('ECRITURE',TOBEcr,-1) ;
    PieceVersECR(MM,TOBPiece,TOBTiers,TOBTTC,False) ;
    {Tiers}
    TOBTTC.PutValue('E_AUXILIAIRE',TOBTiers.GetValue('T_AUXILIAIRE')) ;
    TOBTTC.PutValue('E_GENERAL',GColl) ;
    TOBTTC.PutValue('E_CONSO',TOBTiers.GetValue('T_CONSO')) ;
    {Pièce}
    TOBTTC.PutValue('E_BLOCNOTE',TOBPiece.GetValue('GP_BLOCNOTE')) ;
    TOBTTC.PutValue('E_RIB',TobTiers.GetValue('RIB')); // JT eQualité 11032
    {Divers}
    TOBTTC.PutValue('E_TYPEMVT','TTC') ;
    TOBTTC.PutValue('E_ETATLETTRAGE','AL') ;
    TOBTTC.PutValue('E_EMETTEURTVA','X') ;
    AlimLibEcr(TobTTC,TobPiece,TobTiers,'Retenue Garantie',tecRG,True,(MM.Simul='S'));
    if VH_GC.GCIfDefCEGID then
    if TOBTiers.FieldExists('LIBTIERSCEGID') then TOBTTC.PutValue('E_LIBELLE',TOBTiers.GetValue('LIBTIERSCEGID')) ;
    {Eche+Vent}
    if OkVent then BEGIN NumL:=cpt ; NumE:=1 ; END else BEGIN NumL:=1 ; NumE:=cpt ; END ;
    TOBTTC.PutValue('E_ECHE','X') ;
    if OkVent then TOBTTC.PutValue('E_ANA','X') ;
    TOBTTC.PutValue('E_NUMLIGNE',NumL) ; TOBTTC.PutValue('E_NUMECHE',NumE) ;
    {Echéances}
    ModePaie := GetParamSocSecur ('SO_BTMODEPAIEASS','RG');
    TOBTTC.PutValue('E_MODEPAIE',ModePaie) ;
    TOBTTC.PutValue('E_DATEECHEANCE',LastDate) ;
    //FV1 - 23/02/2018 : FS#2964 - Viviane - Il faut charger le champ E_DATEVALEUR avec le champ E_DATECOMPTABLE pour la ligne « TTC »
    TOBTTC.PutValue('E_DATEVALEUR', MM.DateC) ;
    {Montants}
    DP:=0 ; CP:=0 ; DD:=0 ; CD:=0 ;
    GetMontantRGReliquat(TOBPIeceRG, XD, XP,True);
    // --
    if MM.Nature='FC' then BEGIN DD:=XD ; DP:=XP ; CD:=0 ; CP:=0 ; END else
    if MM.Nature='FF' then BEGIN CD:=XD ; CP:=XP ; DD:=0 ; DP:=0 ; END else
    if MM.Nature='AC' then BEGIN CD:=-XD ; CP:=-XP ; DD:=0 ; DP:=0 ; END else
    if MM.Nature='AF' then BEGIN DD:=-XD ; DP:=-XP ; CD:=0 ; CP:=0 ; END ;
    TOBTTC.PutValue('E_DEBIT',DP)     ; TOBTTC.PutValue('E_CREDIT',CP) ;
    TOBTTC.PutValue('E_DEBITDEV',DD)  ; TOBTTC.PutValue('E_CREDITDEV',CD) ;
    //FV1 - 23/04/2018 : FS#3043 - Viviane : renseigner le champ "E_ENCAISSEMENT" lors de la validation d'une facture
    TOBTTC.PutValue('E_ENCAISSEMENT', SensEnc(DP,CP)) ;
    //
    if ((DP=0) and (CP=0)) then TOBTTC.Free else
    BEGIN
      TheTOBRg := TOBTTC;
      if OkVent then NbEches:=1 else Inc(NbEches) ;
    END ;
  end;

  // Stockage date d'échéance
  MM.LastDateEche := LastDate;

  TOBG.Free ;
  // Modif BTP
  TOBEchesRg.free;
  // --
  {Gestion de l'acompte}
  DetruitEcheancesAcompte(TOBPiece,TOBEches,TOBAcomptes) ;
END ;


Procedure AjusteTaxesBases ( TOBPiece,TOBBases,TOBPorcs : TOB ; NbDec : integer ) ;
Type DPE = (D,P) ;
Var TaxesB,TaxesL : Array[DPE,1..5] of Double ;
    MaxT          : Array[1..5] of Double ;
    LeMax         : Array[1..5] of integer ;
    EcartD,EcartP,XD,XP : Double ;
    i,k,NumT : integer ;
    CumHT : T_CodeCpta ;
    CatT : String ;
    TOBB,TOBP,TT : TOB ;
BEGIN
  if TTCSANSTVA then Exit ;
  FillChar(TaxesB,Sizeof(TaxesB),#0) ; FillChar(TaxesL,Sizeof(TaxesL),#0) ; FillChar(MaxT,Sizeof(MaxT),#0) ;
  for i:=1 to 5 do LeMax[i]:=-1 ;

  for i:=0 to TTA.Count-1 do
    BEGIN
    CumHT:=T_CodeCpta(TTA[i]) ;
    for k:=1 to 5 do
        BEGIN
        TaxesL[D,k]:=TaxesL[D,k]+CumHT.SommeTaxeD[k] ;
        TaxesL[P,k]:=TaxesL[P,k]+CumHT.SommeTaxeP[k] ;
      if ((CumHT.SommeTaxeD[k]<>0) and (Abs(CumHT.SommeTaxeD[k])>MaxT[k])) then
      BEGIN
        MaxT[k]:=Abs(CumHT.SommeTaxeD[k]) ;
        LeMax[k]:=i ;
       END ;
    END ;
  END ;
  NumT := 0;
  for i:=0 to TOBBases.Detail.Count-1 do
    BEGIN
    TOBB:=TOBBases.Detail[i] ;
    CatT:=TOBB.GetValue('GPB_CATEGORIETAXE') ;
    NumT:=Ord(CatT[3])-48 ;
    RecupLesDCTaxes(TOBB,XD,XP) ;
    TaxesB[D,NumT]:=TaxesB[D,NumT]+XD ;
    TaxesB[P,NumT]:=TaxesB[P,NumT]+XP ;
    END ;
  // prise en compte des TVAs pour les Retenues divers HT
  for I:= 0 to TOBPorcs.detail.count -1 do
  begin
    TOBP:=TOBPorcs.Detail[i] ;
    if TOBP.GetBoolean('GPT_FRAISREPARTIS') then continue;
    if TOBP.GetBoolean('GPT_RETENUEDIVERSE') and (TOBP.GetString('GPT_TYPEPORT')<>'HT') and (TOBP.GetString('GPT_TYPEPORT')<>'MI') and (TOBP.GetString('GPT_TYPEPORT')<>'MT') then continue;
    for k:=1 to 5 do
    BEGIN
      TaxesB[D,NumT]:=TaxesB[D,NumT]+TOBP.GetDouble('GPT_TOTALTAXEDEV'+InttoStr(k)) ;
      TaxesB[P,NumT]:=TaxesB[P,NumT]+TOBP.GetDouble('GPT_TOTALTAXE'+InttoStr(k)) ;
    END;
  end;
  {Ajustement final}
  for k:=1 to 5 do
  BEGIN
    EcartD:=Arrondi(TaxesB[D,k]-TaxesL[D,k],NbDec) ;
    EcartP:=Arrondi(TaxesB[P,k]-TaxesL[P,k],V_PGI.OkDecV) ;
    if ((EcartD<>0) or (EcartP<>0)) then if LeMax[k]>=0 then
       BEGIN
       CumHT:=T_CodeCpta(TTA[LeMax[k]]) ;
      // Correctif // LS
      if IsCompteTaxeSurEncaissement(CumHT.CptHT) then
      begin
        TT := TobTVASurEncaiss.FindFirst(['INDEX'],['TX1' + ';' +CumHT.FamTaxe[k]], true);
        if TT <> nil then
        begin
          TT.SetDouble('MONTANTDEV',TT.GetDouble('MONTANTDEV')+EcartD);
          TT.SetDouble('MONTANT',TT.GetDouble('MONTANT')+EcartP);
        end;
      end;
       CumHT.SommeTaxeD[k]:=Arrondi(CumHT.SommeTaxeD[k]+EcartD,NbDec) ;
       CumHT.SommeTaxeP[k]:=Arrondi(CumHT.SommeTaxeP[k]+EcartP,V_PGI.OkDecV) ;
       END ;
    END ;
END ;

Procedure AjusteHTBases ( TOBPiece,TOBBases,TOBBasesCharge,TOBPorcs : TOB ; R_ArtFi : T_MontantArtFi; NbDec : integer ) ;
Var TauxEsc,TauxRem,EscD,EscP,RemD,RemP : Double ;
    TotMHT_D,TotMHT_P,TotTPF_D,TotTPF_P,TotBase_D,TotBase_P : Double ;
    i,k,kk : integer ;
    CumHT : T_CodeCpta ;
    XDT : T_DetAnal ;
    TOBB,TOBB1 : TOB ;
    XD,XP,MaxM,EcartD,EcartP : Double ;
    LeMax : integer ;
    FamBase : String ;
    stLaFamTaxeEcart : string; // DBR Ecart sur ligne d'écriture
BEGIN
  {JLD 23/09/2003 : bricole infâme mais sans risque a priori}
  if uppercase(Copy(GetParamSoc('SO_NIF'),1,2)) = 'ES' then Exit ;
  {Fin modif JLD}
  if TTCSANSTVA then Exit ;
  if TOBBases.Detail.Count<=0 then Exit ;
  TauxEsc:=TOBPiece.GetValue('GP_ESCOMPTE') ; TauxRem:=TOBPiece.GetValue('GP_REMISEPIED') ;
  RecupLesDCEsc(TOBPiece,EscD,EscP) ; RecupLesDCRem(TOBPiece,RemD,RemP) ;
  if ((TauxRem<>0) or (TauxEsc<>0)) then
   BEGIN
   {Si Escompte ou remise pied alors ajustement en masse}
   TotMHT_D:=0  ; TotMHT_P:=0  ; TotTPF_D:=0 ; TotTPF_P:=0 ; TotBase_D:=0 ; TotBase_P:=0 ;
   MaxM:=0 ; LeMax:=-1 ;
   for i:=0 to TTA.Count-1 do
       BEGIN
       CumHT:=T_CodeCpta(TTA[i]) ;
       TotMHT_D:=TotMHT_D+CumHT.XD ; TotMHT_P:=TotMHT_P+CumHT.XP ;
       if Abs(CumHT.XD)>MaxM then BEGIN MaxM:=Abs(CumHT.XD) ; LeMax:=i ; END ;
       for k:=2 to 5 do
           BEGIN
        {$IFDEF BTP}
					 if BTTypeTaxe (CumHT.FamTaxe[k]) <> bttTPF then continue;
        {$ENDIF}
           TotTPF_D:=TotTPF_D+CumHT.SommeTaxeD[k] ;
           TotTPF_P:=TotTPF_P+CumHT.SommeTaxeP[k] ;
           END ;
       END ;
   for i:=0 to TOBBases.Detail.Count-1 do
       BEGIN
       TOBB:=TOBBases.Detail[i] ;
      {$IFDEF BTP}
	 		 if BTTypeTaxe (TOBB.GetValue('GPB_CATEGORIETAXE')) = bttTVA then
      {$ELSE}
       if TOBB.GetValue('GPB_CATEGORIETAXE')='TX1' then
      {$ENDIF}
          BEGIN
          RecupLesDC(TOBB,XD,XP) ;
          TotBase_D:=TotBase_D+XD ; TotBase_P:=TotBase_P+XP ;
          END ;
       END ;
    for i:=0 to TOBBasesCharge.Detail.Count-1 do
    BEGIN
      TOBB:=TOBBasesCharge.Detail[i] ;
      {$IFDEF BTP}
      if BTTypeTaxe (TOBB.GetValue('GPB_CATEGORIETAXE')) = bttTVA then
      {$ELSE}
      if TOBB.GetValue('GPB_CATEGORIETAXE')='TX1' then
      {$ENDIF}
      BEGIN
        RecupLesDC(TOBB,XD,XP) ;
        TotBase_D:=TotBase_D+XD ; TotBase_P:=TotBase_P+XP ;
      END ;
    END ;
   {Ajustement final}
   EcartD:=(TotMHT_D+TotTPF_D-EscD-RemD)-TotBase_D + R_ArtFi.MontantDev ;
   EcartP:=(TotMHT_P+TotTPF_P-EscP-RemP)-TotBase_P + R_ArtFi.Montant ;
    {DBR CPA DEBUT}
   if ((EcartD<>0) or (EcartP<>0)) then
    //   if ((EcartD<>0) or (EcartP<>0)) then if LeMax>=0 then
    {DBR CPA FIN}
     BEGIN
      {DBR CPA DEBUT}
      if VH_GC.ModeGestionEcartComptable = 'CPA' then
      begin
       if (EcartD > 0.0000001) or (EcartD < -0.0000001) or (EcartP > 0.0000001) or (EcartP < -0.0000001) then
         BEGIN
           // DBR Ecart sur ligne d'écriture
           CumHT:=T_CodeCpta.Create;
           if TOBPiece.GetValue('GP_VENTEACHAT') = 'VEN' then
           begin
             if EcartD > 0 then
               CumHT.CptHT:=GetParamSoc ('SO_GCECARTCREDIT')
             else if EcartD < 0 then
               CumHT.CptHT:=GetParamSoc ('SO_GCECARTDEBIT');
             if CumHT.CptHT = '' then CumHT.CptHT := GetParamSoc ('SO_CLIATTEND');
           end else
           begin
             CumHT.CptHT:=GetParamSoc ('SO_GCCPTEHTACH');
             if CumHT.CptHT = '' then CumHT.CptHT := GetParamSoc ('SO_FOUATTEND');
           end;
           CumHT.LibU:=True ;
           CumHT.LibArt:=TexteMessage [19];
           //CumHT.LibArt := 'Ecart';
           CumHT.Affaire:='';
           for kk:=1 to 5 do CumHT.FamTaxe[kk]:='';
           //CumHT.FamTaxe [1] := stLaFamTaxeEcart;

           CumHT.XD:=-EcartD ;
           CumHT.XP:=-EcartP;
           CumHT.Qte:=1 ;
           TTA.Add(CumHT) ;
         end;
      end else if LeMax >= 0 then
      begin
        {DBR CPA FIN}
        CumHT:=T_CodeCpta(TTA[LeMax]) ;
        CumHT.XD:=Arrondi(CumHT.XD-EcartD,NbDec) ;
        CumHT.XP:=Arrondi(CumHT.XP-EcartP,V_PGI.OkDecV) ;
        {Ré-ajuster l'analytique cf Pileje courrier 8 point 2.3}
        if CumHT.Anal.Count>0 then
           BEGIN
           XDT:=T_DetAnal(CumHT.Anal[0]) ;
           XDT.MD:=Arrondi(XDT.MD-EcartD,NbDec) ;
           XDT.MP:=Arrondi(XDT.MP-EcartP,V_PGI.OkDecV) ;
           END ;
      END;
     end; {DBR CPA}
   END else
   BEGIN
   {Si Ni Escompte Ni remise pied alors ajustement par base Tva}
   for k:=0 to TOBBases.Detail.Count-1 do
      BEGIN
       TOBB:=TOBBases.Detail[k] ;
      {$IFDEF BTP}
	 		 if (BTTypeTaxe (TOBB.GetValue('GPB_CATEGORIETAXE')) <> bttTVA) then continue;
      {$ELSE}
       if TOBB.GetValue('GPB_CATEGORIETAXE')<>'TX1' then Continue ;
      {$ENDIF}
       FamBase:=TOBB.GetValue('GPB_FAMILLETAXE') ;
       TotMHT_D:=0 ; TotMHT_P:=0 ; TotTPF_D:=0 ; TotTPF_P:=0 ;
       MaxM:=0 ;
       LeMax:=-1 ;
       for i:=0 to TTA.Count-1 do
           BEGIN
           CumHT:=T_CodeCpta(TTA[i]) ; if CumHT.FamTaxe[1]<>FamBase then Continue ;
           TotMHT_D:=TotMHT_D+CumHT.XD ; TotMHT_P:=TotMHT_P+CumHT.XP ;
           if Abs(CumHT.XD)>MaxM then
           BEGIN
            MaxM:=Abs(CumHT.XD) ;
            stLaFamTaxeEcart := CumHT.FamTaxe [k+1]; // DBR Ecart sur ligne d'écriture
            LeMax:=i ; // DBR Ecart sur ligne d'écriture
           END ;
           for kk:=2 to 5 do
               BEGIN
          {$IFDEF BTP}
					 		 if BTTypeTaxe (CumHT.FamTaxe[k]) <> bttTPF then continue;
          {$ENDIF}
               TotTPF_D:=TotTPF_D+CumHT.SommeTaxeD[kk] ;
               TotTPF_P:=TotTPF_P+CumHT.SommeTaxeP[kk] ;
               END ;
           END ;
       {Ajustement final}
       RecupLesDC(TOBB,XD,XP) ; TotBase_D:=XD ; TotBase_P:=XP ;
      // recup des montant HT des retenue Diverses HT
      GetRetenuesCollectif (TOBPorcs,XP,XD,'HT');
      TotBase_D:=TotBase_D + XD ; TotBase_P:=TotBase_P + XP;
      //
      TOBB1:=TOBBasesCharge.findFirst(['GPB_CATEGORIETAXE','GPB_FAMILLETAXE'],[TOBB.GetValue('GPB_CATEGORIETAXE'),FamBase],true);
      if TOBB1 <> nil then
      begin
        RecupLesDC(TOBB1,XD,XP);
        TotBase_D:=TotBase_D + XD;
        TotBase_P:=TotBase_P + XP;
      end;
      //
       EcartD:=(TotMHT_D+TotTPF_D)-TotBase_D + R_ArtFi.MontantDev;
       EcartP:=(TotMHT_P+TotTPF_P)-TotBase_P + R_ArtFi.Montant;
      {DBR CPA DEBUT}
       if ((EcartD<>0) or (EcartP<>0)) then
        // DBR - Test ci dessous car EcartD ou EcartP proche de 0 mais pas égal à 0 (10e-14)
      {DBR CPA FIN}
       begin
        {DBR CPA DEBUT}
        if VH_GC.ModeGestionEcartComptable = 'CPA' then
        begin
          {DBR CPA FIN}
         if (EcartD > 0.0000001) or (EcartD < -0.0000001) or (EcartP > 0.0000001) or (EcartP < -0.0000001) then
          //        if ((EcartD<>0) or (EcartP<>0)) then if LeMax>=0 then // DBR Ecart sur ligne d'écriture
           BEGIN
             // DBR Ecart sur ligne d'écriture
             CumHT:=T_CodeCpta.Create;
             if TOBPiece.GetValue('GP_VENTEACHAT') = 'VEN' then
             begin
              if EcartD > 0 then CumHT.CptHT:=GetParamSoc ('SO_GCECARTCREDIT')
               else if EcartD < 0 then
                 CumHT.CptHT:=GetParamSoc ('SO_GCECARTDEBIT');
               if CumHT.CptHT = '' then CumHT.CptHT := GetParamSoc ('SO_CLIATTEND');
             end else
             begin
               CumHT.CptHT:=GetParamSoc ('SO_GCCPTEHTACH');
               if CumHT.CptHT = '' then CumHT.CptHT := GetParamSoc ('SO_FOUATTEND');
             end;
             CumHT.LibU:=True ;
             CumHT.LibArt:=TexteMessage [19];
             //CumHT.LibArt := 'Ecart';
             CumHT.Affaire:='';
             for kk:=1 to 5 do CumHT.FamTaxe[kk]:='';
             CumHT.FamTaxe [1] := stLaFamTaxeEcart;

             CumHT.XD:=-EcartD ;
             CumHT.XP:=-EcartP;
             CumHT.Qte:=1 ;
             TTA.Add(CumHT) ;
          END;
          {DBR CPA DEBUT}
        end else if LeMax >= 0 then
        begin
          {DBR CPA FIN}
          CumHT:=T_CodeCpta(TTA[LeMax]) ;
          CumHT.XD:=Arrondi(CumHT.XD-EcartD,NbDec) ;
          CumHT.XP:=Arrondi(CumHT.XP-EcartP,V_PGI.OkDecV) ;
          {Ré-ajuster l'analytique cf Pileje courrier 8 point 2.3}
          if CumHT.Anal.Count>0 then
             BEGIN
             XDT:=T_DetAnal(CumHT.Anal[0]) ;
             XDT.MD:=Arrondi(XDT.MD-EcartD,NbDec) ;
             XDT.MP:=Arrondi(XDT.MP-EcartP,V_PGI.OkDecV) ;
             END ;
        END ;
       end; {DBR CPA}
      END ;
   END ;
END ;

Function CreerLignesEcrHT ( TOBEcr,TOBPiece,TOBTiers : TOB ; MM : RMVT ; IndiceCpta : integer ) : T_RetCpta ;
Var CumHT : T_CodeCpta ;
    CHT : String ;
    TOBG,TOBE : TOB ;
    XD,XP  : Double ;
    CD,CP  : Double ;
    DD,DP  : Double ;
    OkVent : boolean ;
    NumL,i : integer ;
    Sta,LibDefaut  : String ;
BEGIN
Result:=rcOk ;
CP := 0; DP := 0; CD := 0 ; DD := 0;
CumHT:=T_CodeCpta(TTA[IndiceCpta]) ;
// Erreur sur Compte HT
CHT:=CumHT.CptHT ; if CHT='' then BEGIN Result:=rcPar ; LastMsg:=5 ; Exit ; END ;
{Etude du compte général HT}
TOBG:=CreerTOBGeneral(CHT) ;
// Erreur sur Compte HT
if TOBG=Nil then
   BEGIN
   if VH_GC.GCPontComptable='ATT' then Result:=rcPar else
    if VH_GC.GCPontComptable='REF' then Result:=rcRef else
       BEGIN
       if Not CreerCompteGC(TOBG,CHT,CumHT.FamTaxe[1],CumHT.FamTaxe[2],ccpHT,MM.Nature) then Result:=rcPar ;
       END ;
   if Result<>rcOk then BEGIN LastMsg:=5 ; Exit ; END ;
   END ;
OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
{Ligne d'écriture}
TOBE:=TOB.Create('ECRITURE',TOBEcr,-1) ;
if CumHT.Anal.Count>0 then
   BEGIN
   TOBE.AddChampSup('ANAL',False) ; TOBE.AddChampSup('AXES',False) ;
   TOBE.PutValue('ANAL',IndiceCpta) ;
   Sta:='' ; for i:=1 to 5 do if TOBG.GetValue('G_VENTILABLE'+IntToStr(i))='X' then Sta:=Sta+IntToStr(i) ;
   TOBE.PutValue('AXES',Sta) ;
   END ;
PieceVersECR(MM,TOBPiece,TOBTiers,TOBE,False) ;
CompteGenVersEcr (TOBPiece,TOBE,CumHt.CptHT );
{Général}
TOBE.PutValue('E_GENERAL',CHT) ;
TOBE.PutValue('E_CONSO',TOBG.GetValue('G_CONSO')) ;
{Divers}
TOBE.PutValue('E_TYPEMVT','HT') ;
TOBE.PutValue('E_QTE1',CumHT.Qte) ;
TOBE.PutValue('E_TVA',CumHT.FamTaxe[1]) ; TOBE.PutValue('E_TPF',CumHT.FamTaxe[2]) ;
if CumHT.Affaire<>'' then TOBE.PutValue('E_AFFAIRE',CumHT.Affaire) ;
if CumHT.LibU then LibDefaut:=CumHT.LibArt else LibDefaut:=TOBG.GetValue('G_LIBELLE');
AlimLibEcr(TobE,TobPiece,TobTiers,LibDefaut,tecHT,True,(MM.Simul='S'));
{Contreparties}
TOBE.PutValue('E_CONTREPARTIEGEN',TOBTiers.GetValue('T_COLLECTIF')) ;
TOBE.PutValue('E_CONTREPARTIEAUX',TOBTiers.GetValue('T_AUXILIAIRE')) ;
{Eche+Vent}
NumL:=TOBEcr.Detail.Count-NbEches+1 ; TOBE.PutValue('E_NUMLIGNE',NumL) ;
if PremHT<0 then PremHT:=TOBEcr.Detail.Count ;
if OkVent then TOBE.PutValue('E_ANA','X') ;
{Montants}
XD:=CumHT.XD ; XP:=CumHT.XP ;
if MM.Nature='FC' then BEGIN CD:=XD ; CP:=XP ; DD:=0 ; DP:=0 ; END else
if MM.Nature='FF' then BEGIN DD:=XD ; DP:=XP ; CD:=0 ; CP:=0 ; END else
if MM.Nature='AC' then BEGIN DD:=-XD ; DP:=-XP ; CD:=0 ; CP:=0 ; END else
if MM.Nature='AF' then BEGIN CD:=-XD ; CP:=-XP ; DD:=0 ; DP:=0 ; END ;
TOBE.PutValue('E_DEBIT',DP)     ; TOBE.PutValue('E_CREDIT',CP) ;
TOBE.PutValue('E_DEBITDEV',DD)  ; TOBE.PutValue('E_CREDITDEV',CD) ;
if ((DP=0) and (CP=0)) then TOBE.Free ;
TOBG.Free ;
END ;

Function FabricSQLCompta ( FamArt,FamTiers,FamAff,Etabl,Regime,FamTaxe,VenteAchat : String ) : String ;
Var wArt,wTiers,wAff,WVenteAchat : String ;
BEGIN
  if (GetParamSocSecur('SO_GCVENTCPTAART', False)) then
  begin
  	if Famart = '' then
    begin
    	wArt:='AND (GCP_COMPTAARTICLE="")';
    end else
    begin
    	wArt:='AND (GCP_COMPTAARTICLE="'+FamArt+'" OR GCP_COMPTAARTICLE="")';
    end;
  end else
  begin
    wArt:='' ;
  end;
  if (GetParamSocSecur('SO_GCVENTCPTATIERS', False)) then
  begin
  	if Famtiers = '' then
    begin
	    wTiers:='AND (GCP_COMPTATIERS="")';
    end else
    begin
  		wTiers:='AND (GCP_COMPTATIERS="'+FamTiers+'" OR GCP_COMPTATIERS="")';
  	end;
  end else
  begin
  	wTiers:='' ;
  end;
  if (GetParamSocSecur('SO_GCVENTCPTAAFF', False)) then
  begin
  	if FamAFF = '' then
    begin
  		wAff:='AND (GCP_COMPTAAFFAIRE="")';
    end else
    begin
  		wAff:='AND (GCP_COMPTAAFFAIRE="'+FamAff+'" OR GCP_COMPTAAFFAIRE="")';
    end;
  end else
  begin
  	wAff:='' ;
  end;
  if VenteAchat <> '' then
  begin
    WVenteAchat := 'AND (GCP_VENTEACHAT="'+VenteAchat+'" OR GCP_VENTEACHAT="")';
  end else
  begin
    WVenteAchat := 'AND GCP_VENTEACHAT=""';
  end;
  Result:='SELECT * FROM CODECPTA WHERE (GCP_ETABLISSEMENT="'+Etabl+'" OR GCP_ETABLISSEMENT="") '
         +'AND (GCP_REGIMETAXE="'+Regime+'" OR GCP_REGIMETAXE="") '
         +'AND (GCP_FAMILLETAXE="'+FamTaxe+'" OR GCP_FAMILLETAXE="") '
         +wArt+' '+wTiers+' '+wAff+' '+WVenteAchat+' '
         +'ORDER BY GCP_COMPTAARTICLE DESC, GCP_COMPTATIERS DESC, GCP_COMPTAAFFAIRE DESC, GCP_REGIMETAXE DESC, GCP_FAMILLETAXE DESC, GCP_ETABLISSEMENT DESC' ;
END ;

Function  CreerTOBCpta : TOB ;
Var TOBCpta : TOB ;
BEGIN
TOBCpta:=TOB.Create('',Nil,-1) ;
TOBCpta.AddChampSup('LASTSQL',False) ; TOBCpta.PutValue('LASTSQL','') ;
Result:=TOBCpta ;
END ;

Function  CreerTOBCodeCpta ( TOBParent : TOB ) : TOB ;
Var TOBC : TOB ;
BEGIN
TOBC:=TOB.Create('CODECPTA',TOBParent,-1) ;
TOB.Create('Ecr',TOBC,-1) ; TOB.Create('Stk',TOBC,-1) ;
Result:=TOBC ;
END ;

Function FindTOBCode ( TOBGC : TOB ; FamArt,FamTiers,FamAff,Etabl,Regime,FamTaxe : String ) : TOB ;
Var TOBC : TOB ;
    i : integer ;
    fArt,fTiers,fAff : string;
BEGIN
Result:=Nil ;
for i:=TOBGC.Detail.Count-1 downto 0 do
    BEGIN
    TOBC:=TOBGC.Detail[i] ;
    if ((GetParamSocSecur('SO_GCVENTCPTAART', False)) and (TOBC.GetValue('GCP_COMPTAARTICLE')<>'') and (TOBC.GetValue('GCP_COMPTAARTICLE')<>FamArt)) then Continue ;
    if ((GetParamSocSecur('SO_GCVENTCPTATIERS', False)) and (TOBC.GetValue('GCP_COMPTATIERS')<>'') and (TOBC.GetValue('GCP_COMPTATIERS')<>FamTiers)) then Continue ;
    if ((GetParamSocSecur('SO_GCVENTCPTAAFF', False)) and (TOBC.GetValue('GCP_COMPTAAFFAIRE')<>'') and (TOBC.GetValue('GCP_COMPTAAFFAIRE')<>FamAff)) then Continue ;
    if ((TOBC.GetValue('GCP_REGIMETAXE')<>'') and (TOBC.GetValue('GCP_REGIMETAXE')<>Regime)) then Continue ;
    if ((TOBC.GetValue('GCP_FAMILLETAXE')<>'') and (TOBC.GetValue('GCP_FAMILLETAXE')<>FamTaxe)) then Continue ;
    if ((TOBC.GetValue('GCP_ETABLISSEMENT')<>'') and (TOBC.GetValue('GCP_ETABLISSEMENT')<>Etabl)) then Continue ;
    Result:=TOBC ; Break ;
    END ;
END ;

Function TraiteFamilleAffaire ( CodeAffaire:string;  TobAffaire : TOB) : string;
begin
Result := '';
if Not(GetParamSocSecur('SO_GCVENTCPTAAFF', False)) then Exit;
if CodeAffaire = '' then Exit;

if (TobAffaire<>Nil) then
   if (TobAffaire.GetValue('AFF_AFFAIRE') = CodeAffaire) then
      begin Result := TobAffaire.GetValue('AFF_COMPTAAFFAIRE'); Exit; end;
{$IFDEF AFFAIRE}
Result:=GetChampsAffaire(CodeAffaire,'AFF_COMPTAAFFAIRE');
{$endif}
end;

Function  ChargeAjouteCompta ( TOBCpta,TOBPiece,TOBL,TOBA,TOBTiers,TOBCata,TobAffaire : TOB ; AvecAnal : boolean;TOBGCS:TOB=nil ) : TOB ;
Var VenteAchat,Nature,Regime,Etab,FamArt,FamTiers,FamAff,SQL,SQLAnal,FamTaxe : String ;
    NatV,sRang,LastSQL : String ;
    TOBC,TOBEcr,TOBStk,TOBAna,TOBXX : TOB ;
    Q    : TQuery ;
    ii   : integer ;
    prefixe,CodeAffaire : string;
BEGIN
  prefixe := GetPrefixeTable (TOBL);
  Result:=Nil ;
  // VARIANTE
  if IsVariante(TOBL) then exit;
  if IsLigneFromCentralis(TOBL) then exit;
  // --
  Regime:=TOBPiece.GetValue('GP_REGIMETAXE')   ; Etab:=TOBPiece.GetValue('GP_ETABLISSEMENT') ;
  FamTiers:=TOBTiers.GetValue('T_COMPTATIERS') ;
  if TOBL.GetValue(prefixe+'_TYPEREF')='CAT' then FamArt:=TOBCata.GetValue('GCA_COMPTAARTICLE')
                       											 else FamArt:=TOBA.GetValue('GA_COMPTAARTICLE') ;
  // --
  //FamAff := TraiteFamilleAffaire (TOBL.GetValue(prefixe+'_AFFAIRE'), TobAffaire );
  //--
  // --
  CodeAffaire := TOBL.GetValue(prefixe+'_AFFAIRE');
  if Copy(CodeAffaire,1,1)='W' then
  begin
    CodeAffaire := GetChampsAffaire (CodeAffaire,'AFF_CHANTIER');
    if Trim(Copy(CodeAffaire,2,14)) <> '' then   // correction BRL : certaines fois le code n'est pas à blanc mais à "A               00"
    begin
      FamAff := TraiteFamilleAffaire (CodeAffaire, TobAffaire );
    end else
    begin
      FamAff := TraiteFamilleAffaire (TOBL.GetValue(prefixe+'_AFFAIRE'), TobAffaire );
    end;
  end else
  begin
  	FamAff := TraiteFamilleAffaire (TOBL.GetValue(prefixe+'_AFFAIRE'), TobAffaire );
  end;

  Nature:=TOBPiece.GetValue('GP_NATUREPIECEG') ;
  VenteAchat:=GetInfoParPiece(Nature,'GPP_VENTEACHAT') ;
  if VenteAchat='VEN' then NatV:='HV' else NatV:='HA' ;
  //
  FamTaxe := TOBL.GetValue(prefixe+'_FAMILLETAXE1');
  //
  TOBC:=FindTOBCode(TOBCpta,FamArt,FamTiers,FamAff,Etab,Regime,FamTaxe) ;
  LastSQL:='' ; SQL:='' ;
  if TOBC=Nil then
  BEGIN
    SQL:=FabricSQLCompta(FamArt,FamTiers,FamAff,Etab,Regime,FamTaxe,VenteAchat) ;
    //   if TOBCpta.FieldExists('LASTSQL') then LastSQL:=TOBCpta.GetValue('LASTSQL') ;
  END ;
  if ((TOBC=Nil)) then
  BEGIN
    Q:=OpenSQL(SQL,True,-1, '', True) ;
    if Not Q.EOF then
    BEGIN
      TOBXX := TOB.Create ('LES VENTIL',nil,-1);
      TOBXX.LoadDetailDB('CODECPTA','','',Q,false);
      TOBC:=FindTOBCode(TOBXX,FamArt,FamTiers,FamAff,Etab,Regime,FamTaxe) ;
      if TOBC <> nil then
      begin
        TOB.Create('Ecr',TOBC,-1) ; TOB.Create('Stk',TOBC,-1) ;
        TOBC.ChangeParent(TOBCpta,-1);
      end;
      TOBXX.Free;
    END ;
    Ferme(Q) ;
    if TOBC = nil then
    begin
      TOBC := CreerTOBCodeCpta (TOBCpta);
      TOBC.SetString('GCP_COMPTAARTICLE',FamArt);
      TOBC.SetString('GCP_COMPTATIERS',FamTiers);
      TOBC.SetString('GCP_COMPTAAFFAIRE',FamAff);
      TOBC.SetString('GCP_REGIMETAXE',Regime);
      TOBC.SetString('GCP_FAMILLETAXE',FamTaxe);
      TOBC.SetString('GCP_ETABLISSEMENT',Etab);
      TOBC.SetString('GCP_VENTEACHAT',VenteAchat);
      TOBC.SetString('GCP_CPTEGENEACH',VH_GC.GCCpteHTACH);
      TOBC.SetString('GCP_CPTEGENEVTE',VH_GC.GCCpteHTVTE);
    end;
    TOBCpta.Detail.Sort('-GCP_COMPTAARTICLE;-GCP_COMPTATIERS;-GCP_COMPTAAFFAIRE;-GCP_REGIMETAXE;-GCP_FAMILLETAXE;-GCP_ETABLISSEMENT');
    if ((TOBC<>Nil) and (AvecAnal)) then
    BEGIN
      sRang:=IntToStr(TOBC.GetValue('GCP_RANG')) ;
      SQLAnal:='SELECT * FROM VENTIL WHERE (V_NATURE like "'+NatV+'%" OR V_NATURE LIKE "ST%") AND V_COMPTE="'+sRang+'"' ;
      Q:=OpenSQL(SQLAnal,True,-1, '', True) ;
      if Not Q.EOF then
      BEGIN
        TOBEcr:=TOBC.Detail[0] ; TOBStk:=TOBC.Detail[1] ;
        TOBEcr.LoadDetailDB('VENTIL','','',Q,False,True) ;
        for ii:=TOBEcr.Detail.Count-1 downto 0 do
        BEGIN
            TOBAna:=TOBEcr.Detail[ii] ;
            if Copy(TOBAna.GetValue('V_NATURE'),1,2)='ST' then TOBAna.ChangeParent(TOBStk,0) ;
        END ;
        TOBEcr.Detail.Sort('V_NUMEROVENTIL') ; TOBStk.Detail.Sort('V_NUMEROVENTIL') ;
      END ;
      Ferme(Q) ;
      //CONTREM ???? : TOB peut être Nil: et ce n'est pas protégé dans la fonction suivante
    END ;
  END;
  if (TOBC<>Nil) and (AvecAnal) and (TOBA<>Nil) then
  begin
    if GetParamSoc ('SO_GCAXEANALYTIQUE') then
    begin // mcd 20/11/02 mis ene place anamlytique générique
      GetVentilGescom(TOBPiece,TOBL,TOBA,TOBTiers,TOBC,NatV,sRang,TOBGCS)
    end
    else
    begin
      {$IFDEF AFFAIRE}
      if (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) then
      GetVentilAffaire(TOBPiece,TOBL,TOBA,TOBTiers,TOBC,NatV,sRang) ;
      {$endif}
      if ctxMode in V_PGI.PGIContexte then GetVentilGescom(TOBPiece,TOBL,TOBA,TOBTiers,TOBC,NatV,sRang,TOBGCS)
    end ;
  end;
  Result:=TOBC ;
END ;

{$IFDEF BTP}
Function  ChargeAjouteComptaBTP ( TOBCpta,TOBPiece,TOBL,TOBA,TOBTiers,TOBCata,TobAffaire : TOB ; FamTaxe : string; AvecAnal : boolean;TOBGCS:TOB=nil ) : TOB ;
Var VenteAchat,Nature,Regime,Etab,FamArt,FamTiers,FamAff,SQL,SQLAnal : String ;
    NatV,sRang,LastSQL : String ;
    TOBC,TOBEcr,TOBStk,TOBAna,TOBXX : TOB ;
    Q    : TQuery ;
    ii   : integer ;
    prefixe : string;
    CodeAffaire : string;
    NatureTravail : double;
BEGIN
  prefixe := GetPrefixeTable (TOBL);
  NatureTravail := 0;
  if TOBL.FieldExists('BLF_NATURETRAVAIL') then
  begin
  	NatureTravail := valeur(TOBL.GetValue('BLF_NATURETRAVAIL'));
  end;
  if Naturetravail = 0 then
  begin
    if TOBL.NomTable = 'LIGNEOUVPLAT' then
    begin
      NatureTravail := valeur(TOBL.GetValue('BOP_NATURETRAVAIL'));
    end else
    begin
      NatureTravail := valeur(TOBL.GetValue('GLC_NATURETRAVAIL'));
    end;
  end;
  Result:=Nil ;
  // VARIANTE
  if IsVariante(TOBL) then exit;
  // --
  Regime:=TOBPiece.GetValue('GP_REGIMETAXE')   ; Etab:=TOBPiece.GetValue('GP_ETABLISSEMENT') ;
  FamTiers:=TOBTiers.GetValue('T_COMPTATIERS') ;
  FamArt := '';
  if TOBL.GetValue(prefixe+'_TYPEREF')='CAT' then FamArt:=TOBCata.GetValue('GCA_COMPTAARTICLE')
  																					 else FamArt:=TOBA.GetValue('GA_COMPTAARTICLE') ;
  if NatureTravail > 0 then
  begin
    // -- sous traitance --
    if getParamSocSecur('SO_BTVENTILSOUSTRAIT','') <> '' then FamArt := getParamSocSecur('SO_BTVENTILSOUSTRAIT','');  {Famille comptable article}
  end;
  // --
  CodeAffaire := TOBL.GetValue(prefixe+'_AFFAIRE');
  if Copy(CodeAffaire,1,1)='W' then
  begin
    CodeAffaire := GetChampsAffaire (CodeAffaire,'AFF_CHANTIER');
    if Trim(Copy(CodeAffaire,2,14)) <> '' then  // correction BRL : certaines fois le code n'est pas à blanc mais à "A               00"
    begin
      FamAff := TraiteFamilleAffaire (CodeAffaire, TobAffaire );
    end else
    begin
      FamAff := TraiteFamilleAffaire (TOBL.GetValue(prefixe+'_AFFAIRE'), TobAffaire );
    end;
  end else
  begin
  	FamAff := TraiteFamilleAffaire (TOBL.GetValue(prefixe+'_AFFAIRE'), TobAffaire );
  end;

  Nature:=TOBPiece.GetValue('GP_NATUREPIECEG') ;
  VenteAchat:=GetInfoParPiece(Nature,'GPP_VENTEACHAT') ;
  if VenteAchat='VEN' then NatV:='HV' else NatV:='HA' ;
  TOBC:=FindTOBCode(TOBCpta,FamArt,FamTiers,FamAff,Etab,Regime,FamTaxe) ;
  LastSQL:='' ; SQL:='' ;
  if TOBC=Nil then
  BEGIN
    SQL:=FabricSQLCompta(FamArt,FamTiers,FamAff,Etab,Regime,FamTaxe,VenteAchat) ;
    //   if TOBCpta.FieldExists('LASTSQL') then LastSQL:=TOBCpta.GetValue('LASTSQL') ;
  END ;
  if ((TOBC=Nil) {and (SQL<>LastSQL)}) then
  BEGIN
    Q:=OpenSQL(SQL,True,-1, '', True) ;
    if Not Q.EOF then
    BEGIN
      TOBXX := TOB.Create ('LES VENTIL',nil,-1);
      TOBXX.LoadDetailDB('CODECPTA','','',Q,false);
      TOBC:=FindTOBCode(TOBXX,FamArt,FamTiers,FamAff,Etab,Regime,FamTaxe) ;
      if TOBC <> nil then
      begin
        TOB.Create('Ecr',TOBC,-1) ; TOB.Create('Stk',TOBC,-1) ;
        TOBC.ChangeParent(TOBCpta,-1);
      end;
      TOBXX.Free;
    END ;
    Ferme(Q) ;
    if TOBC= nil then
    begin
      TOBC := CreerTOBCodeCpta (TOBCpta);
      TOBC.SetString('GCP_COMPTAARTICLE',FamArt);
      TOBC.SetString('GCP_COMPTATIERS',FamTiers);
      TOBC.SetString('GCP_COMPTAAFFAIRE',FamAff);
      TOBC.SetString('GCP_REGIMETAXE',Regime);
      TOBC.SetString('GCP_FAMILLETAXE',FamTaxe);
      TOBC.SetString('GCP_ETABLISSEMENT',Etab);
      TOBC.SetString('GCP_VENTEACHAT',VenteAchat);
      TOBC.SetString('GCP_CPTEGENEACH',VH_GC.GCCpteHTACH);
      TOBC.SetString('GCP_CPTEGENEVTE',VH_GC.GCCpteHTVTE);
    end;
    TOBCpta.Detail.Sort('-GCP_COMPTAARTICLE;-GCP_COMPTATIERS;-GCP_COMPTAAFFAIRE;-GCP_REGIMETAXE;-GCP_FAMILLETAXE;-GCP_ETABLISSEMENT');
    if ((TOBC<>Nil) and (AvecAnal)) then
    BEGIN
      sRang:=IntToStr(TOBC.GetValue('GCP_RANG')) ;
      SQLAnal:='SELECT * FROM VENTIL WHERE (V_NATURE like "'+NatV+'%" OR V_NATURE LIKE "ST%") AND V_COMPTE="'+sRang+'"' ;
      Q:=OpenSQL(SQLAnal,True,-1, '', True);
      if Not Q.EOF then
      BEGIN
        TOBEcr:=TOBC.Detail[0] ; TOBStk:=TOBC.Detail[1] ;
        TOBEcr.LoadDetailDB('VENTIL','','',Q,False,True) ;
        for ii:=TOBEcr.Detail.Count-1 downto 0 do
        BEGIN
          TOBAna:=TOBEcr.Detail[ii] ;
          if Copy(TOBAna.GetValue('V_NATURE'),1,2)='ST' then TOBAna.ChangeParent(TOBStk,0) ;
        END ;
        TOBEcr.Detail.Sort('V_NUMEROVENTIL') ; TOBStk.Detail.Sort('V_NUMEROVENTIL') ;
      END ;
      Ferme(Q) ;
      //CONTREM ???? : TOB peut être Nil: et ce n'est pas protégé dans la fonction suivante
    END ;
  END;
  if (TOBC<>Nil) and (AvecAnal) and (TOBA<>Nil) then
  begin
    if GetParamSoc ('SO_GCAXEANALYTIQUE') then
    begin // mcd 20/11/02 mis ene place anamlytique générique
      GetVentilGescom(TOBPiece,TOBL,TOBA,TOBTiers,TOBC,NatV,sRang,TOBGCS)
    end else
    begin
      {$IFDEF AFFAIRE}
      if (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) then
      GetVentilAffaire(TOBPiece,TOBL,TOBA,TOBTiers,TOBC,NatV,sRang) ;
      {$endif}
      if ctxMode in V_PGI.PGIContexte then GetVentilGescom(TOBPiece,TOBL,TOBA,TOBTiers,TOBC,NatV,sRang,TOBGCS)
    end ;
  end;

  Result:=TOBC ;
END ;
{$ENDIF}

Function ChargeAjouteComptaPorc ( TOBCpta,TOBPiece,TOBP,TOBTiers,TOBAnaP,TOBAnaS : TOB ) : TOB ;
Var Nature,Regime,Etab,FamArt,FamTiers,FamAff,SQL,FamTaxe : String ;
    NatV,sRang : String ;
    TOBC,TOBEcr,TOBXX : TOB ;
    Q    : TQuery ;
    VenteAchat : string;
BEGIN
//
	VenteAchat := GetInfoParPiece(TOBPiece.getValue('GP_NATUREPIECEG'),'GPP_VENTEACHAT');
  if (Pos(TOBP.GetString('GPT_TYPEPORT'),'PT;MIC;MTC;')>0) and (TOBP.GetDouble('GPT_TOTALTTCDEV')<0) then
  begin
    if VenteAchat='VEN' then NatV:='HA' else NatV:='HV' ;
  end else
  begin
	if VenteAchat='VEN' then NatV:='HV' else NatV:='HA' ;
  end;
  Regime:=TOBPiece.GetValue('GP_REGIMETAXE')   ; Etab:=TOBPiece.GetValue('GP_ETABLISSEMENT') ;
  Nature:=TOBPiece.GetValue('GP_NATUREPIECEG') ;
  FamTiers:=TOBTiers.GetValue('T_COMPTATIERS') ; FamArt:=TOBP.GetValue('GPT_COMPTAARTICLE') ;
  // mcd 25/04/02 FamAff:='' ;
  FamAff := TraiteFamilleAffaire (TOBPiece.GetValue('GP_AFFAIRE'), Nil );
  FamTaxe := TOBP.GetString('GPT_FAMILLETAXE1');
  TOBC:=FindTOBCode(TOBCpta,FamArt,FamTiers,FamAff,Etab,Regime,FamTaxe) ;
  if TOBC<>Nil then BEGIN Result:=TOBC ; Exit ; END ;
  SQL:=FabricSQLCompta(FamArt,FamTiers,FamAff,Etab,Regime,FamTaxe,VenteAchat) ;
  Q:=OpenSQL(SQL,True,-1, '', True) ;
  if Not Q.EOF then
  BEGIN
    TOBXX := TOB.Create ('LES VENTIL',nil,-1);
    TOBXX.LoadDetailDB('CODECPTA','','',Q,false);
    TOBC:=FindTOBCode(TOBXX,FamArt,FamTiers,FamAff,Etab,Regime,FamTaxe) ;
    if TOBC <> nil then
    begin
      TOB.Create('Ecr',TOBC,-1) ; TOB.Create('Stk',TOBC,-1) ;
      TOBC.ChangeParent(TOBCpta,-1);
      TOBCpta.Detail.Sort('-GCP_COMPTAARTICLE;-GCP_COMPTATIERS;-GCP_COMPTAAFFAIRE;-GCP_REGIMETAXE;-GCP_FAMILLETAXE;-GCP_ETABLISSEMENT');
    end else
    begin
      if TOBCpta.FieldExists('LASTSQL') then TOBCpta.PutValue('LASTSQL',SQL) ;
    end;
    TOBXX.Free;
(*
    TOBC:=CreerTOBCodeCpta(TOBCpta) ; TOBC.SelectDB('',Q) ;
    TOBCpta.Detail.Sort('GCP_REGIMETAXE;GCP_ETABLISSEMENT') ;
*)
  END ;
  Ferme(Q) ;

  if (TOBC<>Nil)  then
  BEGIN
    sRang:=IntToStr(TOBC.GetValue('GCP_RANG')) ;
    Q:=OpenSQL('SELECT * FROM VENTIL WHERE V_NATURE like "'+NatV+'%" AND V_COMPTE="'+sRang+'"',True,-1, '', True) ;
    if Not Q.EOF then
    BEGIN
      TOBEcr:=TOBC.Detail[0] ;
      TOBEcr.LoadDetailDB('VENTIL','','',Q,False,True) ;
      TOBEcr.Detail.Sort('V_NUMEROVENTIL') ;
    END ;
    Ferme(Q) ;
    //
    if (TOBC<>Nil)  then
    begin
      if TOBC.detail[0].detail.count = 0 then
      begin
        if GetParamSoc ('SO_GCAXEANALYTIQUE') then
        begin // mcd 20/11/02 mis ene place anamlytique générique
          GetVentilPort(TOBPiece,TOBTiers,TOBC,NatV,sRang)
        end else
        begin
        	GetVentilAffaire(TOBPiece,nil,nil,TOBTiers,TOBC,NatV,sRang) ;
        end ;
      end;
    end;
    //
  END ;
  Result:=TOBC ;
END ;

Procedure CumulAnal ( CumHT : T_CodeCpta ; TOBL : TOB ; XD,XP : Double ) ;
Var XDT : T_DetAnal ;
    TOBA,TOBAL : TOB ;
    Pourc,QteL : Double ;
    Section,Ax      : String ;
    i,k,iTrouv : integer ;
    prefixe : string;
BEGIN
  prefixe := GetPrefixeTable (TOBL);
XDT := nil;
if TOBL.Detail.Count<=0 then Exit ;
TOBAL:=TOBL.Detail[0] ;
QteL:=TOBL.GetValue(prefixe+'_QTEFACT') ;
for i:=0 to TOBAL.Detail.Count-1 do
    BEGIN
    TOBA:=TOBAL.Detail[i] ;
    if prefixe = 'BOP' then
    begin
      TOBA.setDouble ('YVA_MONTANT',XP);
      TOBA.setDouble ('YVA_MONTANTDEV',XD);
    end;
    Ax:=TOBA.GetValue('YVA_AXE') ; Pourc:=TOBA.GetValue('YVA_POURCENTAGE') ;
    Section:=TOBA.GetValue('YVA_SECTION') ;
    iTrouv:=-1 ;
    for k:=0 to CumHT.Anal.Count-1 do
        BEGIN
        XDT:=T_DetAnal(CumHT.Anal[k]) ;
        if ((XDT.Section=Section) and (XDT.Ax=Ax)) then BEGIN iTrouv:=k ; Break ; END ;
        END ;
    if iTrouv<0 then XDT:=T_DetAnal.Create ;
    XDT.Section:=Section ; XDT.aX:=aX ;
    XDT.MD:=XDT.MD+Pourc*XD/100 ; XDT.MP:=XDT.MP+Pourc*XP/100 ;
    XDT.Qte1:=Arrondi(XDT.Qte1+QteL*Pourc/100.0,V_PGI.OkDecQ) ;
    if iTrouv<0 then CumHT.Anal.Add(XDT) ;
    END ;
END ;

Procedure CumulDetOuvAnalSurLigne (TOBL,TOBOD : TOB ) ;
Var XDT : T_DetAnal ;
    TOBA,TOBAL,TOBALL,TOBAAA,TOBPiece : TOB ;
    Pourc,QteL : Double ;
    Section,Ax      : String ;
    i,k,iTrouv : integer ;
    prefixe,RefA : string;
BEGIN
  //
  // traitement sur la ligne de document
  if TOBL.detail.count < 2 then TOB.Create('', TOBL, -1); {Analytique vente/achat}
  if TOBL.detail.count < 2 then TOB.Create('', TOBL, -1); {Analytique stock}
  prefixe := GetPrefixeTable (TOBOD);
  XDT := nil;
  if TOBOD.Detail.Count<=0 then Exit ;
  TOBAL:=TOBOD.Detail[0] ; // analytique detail ouvrage
  TOBALL := TOBL.detail[0]; // analytique ligne
  QteL:=TOBOD.GetValue(prefixe+'_QTEFACT') ;
  // calcul sur la ligne d'ouvrage
  for i:=0 to TOBAL.Detail.Count-1 do
  BEGIN
    TOBA:=TOBAL.Detail[i] ;
    Ax:=TOBA.GetValue('YVA_AXE') ; Pourc:=TOBA.GetValue('YVA_POURCENTAGE') ;
    Section:=TOBA.GetValue('YVA_SECTION') ;

    TOBAAA := TOBALL.findfirst(['YVA_AXE','YVA_SECTION'],[Ax,Section],true);
    if TOBAAA = nil then
    begin
      TOBPiece := TOBL.parent;
  		RefA:=EncodeRefCPGescom(TOBPiece) ;
    	TOBAAA := TOB.Create ('VENTANA',TOBALL,-1);
			TOBAAA.SetString('YVA_TABLEANA','GL');
      TOBAAA.SetString('YVA_IDENTIFIANT',RefA) ;
      TOBAAA.SetString('YVA_NATUREID','GC') ;
      TOBAAA.SetString('YVA_IDENTLIGNE',FormatFloat('000',TOBL.GetValue('GL_NUMLIGNE'))) ;
      TOBAAA.SetString('YVA_AXE',Ax);
      TOBAAA.SetString('YVA_SECTION',Section);
      TOBAAA.SetInteger ('YVA_NUMVENTIL',TOBAAA.GetIndex);
    end;
    TOBAAA.SetDouble('YVA_MONTANT',TOBAAA.GetDouble('YVA_MONTANT')+TOBA.GetDouble('YVA_MONTANT'));
    TOBAAA.SetDouble('YVA_MONTANTDEV',TOBAAA.GetDouble('YVA_MONTANTDEV')+TOBA.GetDouble('YVA_MONTANTDEV'));
    if TOBL.GetDouble('GL_MONTANTHTDEV') <> 0 then
    begin
    	TOBAAA.SetDouble('YVA_POURCENTAGE',Arrondi(TOBAAA.GetDouble('YVA_MONTANTDEV')/TOBL.GetDouble('GL_MONTANTHTDEV')*100,2));
    end;
  END ;
END ;

Function CreerLignesHTTVA ( TOBEcr,TOBAFFInterv,TOBPieceinterv,TOBPiece,TOBOuvrages,TOBOuvragesP,TOBBases,TOBBasesCharge,TOBTiers,TOBArticles,TOBCpta,TOBPorcs,TOBAnaP,TOBAnaS : TOB ; MM : RMVT ; NbDec : integer; DEV : Rdevise ) : T_RetCpta ;
Var i,k,itrouv,kk : integer ;
    TOBA,TOBL,TOBCode,TOBP : TOB ;
    RefUnique            : String ;
    CptHT,LibArt,Affaire : String ;
    CumHT                : T_CodeCpta ;
    OkTX                 : Boolean ;
    XD,XP : Double ;
    XDT : T_DetAnaL ;
    R_ArtFi : T_MontantArtFi ;
    // Modif BTP
    FamTaxeHt : string;
    {Ratio : double;} (* provenance de tva au millieme *)
    Indice : integer;
    TOBBasesL : TOB;
    NumOrdre : integer;
    EnHt : boolean;
    TOBB : TOB;
    // --
BEGIN
	EnHt := (TOBPiece.getvalue('GP_FACTUREHT')='X');
  Result:=rcOk ;
  R_ArtFi.Montant:=0; R_ArtFi.MontantDev:=0;
  {Articles}
  for i:=0 to TOBPiece.Detail.Count-1 do
  BEGIN
    {Ratio := 1; }
    TOBL:=TOBPiece.Detail[i] ;
    {Tests préalables}
    //CONTREM ?????
    RefUnique:=TOBL.GetValue('GL_ARTICLE') ; if RefUnique='' then Continue ;
    TOBA:=FindTOBArtRow(TOBPiece,TOBArticles,i+1) ; if TOBA=Nil then Continue ;
    Numordre := TOBL.GetInteger('GL_NUMORDRE');
    if IsLigneFromCentralis(TOBL) then continue;
    if (TOBA.GetValue('GA_TYPEARTICLE')='FI') then
    begin
      // Les articles financiers ne génèrent pas de lignes d'écritures HT.
      // Cumul des montants, pour ajustement dans AjusteHTBases
      RecupLesDC(TOBL,XD,XP) ;
      R_ArtFi.Montant:=R_ArtFi.Montant+XP; R_ArtFi.MontantDev:=R_ArtFi.MontantDev+XD;
      Continue ;
    end;
    if (TOBL.GetValue('GL_TYPEARTICLE')='OUV') and (TOBL.GetValue('GL_TYPENOMENC')='OUV') and (TOBL.GetValue('GL_INDICENOMEN')>0) then
    begin
      result := TraitementVOuvrage (TOBEcr,TOBPieceinterv,TOBAFFInterv,TOBPiece,TOBL,TOBOuvrages,TOBOuvragesP,TOBTiers,TOBArticles,TOBCpta,TOBPorcs,MM, NbDec,EnHt,DEV);
    end else
    begin
      result := TraitementVLigne (TOBEcr,TOBAFFInterv,TOBPieceinterv,TOBA,TOBPiece,TOBL,TOBTiers,TOBArticles,TOBCpta,TOBPorcs,MM, NbDec, EnHt,DEV);
    end;
    if result<>rcOk then break;
  end;

{Ports et frais}
  if result = rcOk then
  begin
  for i:=0 to TOBPorcs.Detail.Count-1 do
  BEGIN
    TOBP:=TOBPorcs.Detail[i] ;
    {Tests préalables}
    RefUnique:=TOBP.GetValue('GPT_CODEPORT') ; if RefUnique='' then Continue ;
    // Correction fiche qualité 11967
    if TOBP.GetValue('GPT_FRAISREPARTIS') = 'X' then continue;
    if TOBP.GetBoolean('GPT_RETENUEDIVERSE') then continue;
    // -----
    LibArt:=Copy(TOBP.GetValue('GPT_LIBELLE'),1,35) ; Affaire:='' ; CptHT:='' ;
    {Recherche du paramétrage compta associé, sinon l'ajouter}
    TOBCode:=ChargeAjouteComptaPorc(TOBCpta,TOBPiece,TOBP,TOBTiers,TOBAnaP,TOBAnaS) ;
    if TOBCode<>Nil then
    BEGIN
      if (TOBP.GetDouble('GPT_TOTALTTCDEV')<0) and (GetParamSocSecur('SO_VENTILMONTANTSURCHARGE', False)) then
      begin
        if ((MM.Nature='FC') or (MM.Nature='AC')) then CptHT:=TOBCode.GetValue('GCP_CPTEGENEACH')
                                                  else CptHT:=TOBCode.GetValue('GCP_CPTEGENEVTE') ;
      end else
      begin
        if ((MM.Nature='FC') or (MM.Nature='AC')) then CptHT:=TOBCode.GetValue('GCP_CPTEGENEVTE')
                                                  else CptHT:=TOBCode.GetValue('GCP_CPTEGENEACH') ;
      end;
    END ;
    if (TOBCode=Nil) or ((CptHT='') and (VH_GC.GCPontComptable='ATT')) then
    BEGIN
      if (TOBP.GetDouble('GPT_TOTALTTCDEV')<0) and (GetParamSocSecur('SO_VENTILMONTANTSURCHARGE', False)) then
      begin
        if ((MM.Nature='FC') or (MM.Nature='AC')) then CptHT:=VH_GC.GCCptePortACH else CptHT:=VH_GC.GCCptePortVTE ;
      end else
      begin
        if ((MM.Nature='FC') or (MM.Nature='AC')) then CptHT:=VH_GC.GCCptePortVTE else CptHT:=VH_GC.GCCptePortACH ;
      end;
    END ;
    // Erreur sur le compte HT port
    if CptHT='' then BEGIN Result:=rcPar ; LastMsg:=5 ; Break ; END ;
    {Sommations par compte et familles de taxes identiques}
    RecupLesDC(TOBP,XD,XP) ;
    iTrouv:=-1 ;
    for k:=0 to TTA.Count-1 do
    BEGIN
      CumHT:=T_CodeCpta(TTA[k]) ;
      if CumHT.CptHT<>CptHT then Continue ;
      if DistinctAffaire then BEGIN if CumHT.Affaire<>Affaire then Continue ; END ;
      OkTx:=True ;
      for kk:=1 to 5 do if CumHT.FamTaxe[kk]<>TOBP.GetValue('GPT_FAMILLETAXE'+IntToStr(kk)) then OkTx:=False ;
      if OkTX then BEGIN iTrouv:=k ; Break ; END ;
    END ;
    if iTrouv<0 then
    BEGIN
      CumHT:=T_CodeCpta.Create ; CumHT.CptHT:=CptHT ; CumHT.LibU:=True ;
      CumHT.LibArt:=LibArt ; CumHT.Affaire:=Affaire ;
      for kk:=1 to 5 do CumHT.FamTaxe[kk]:=TOBP.GetValue('GPT_FAMILLETAXE'+IntToStr(kk)) ;
    END else
    BEGIN
      CumHT:=T_CodeCpta(TTA[iTrouv]) ;
    END ;
    CumHT.XD:=CumHT.XD+XD ; CumHT.XP:=CumHT.XP+XP ;
    CumHT.Qte:=CumHT.Qte+1 ;
    if CumHT.LibArt<>LibArt then BEGIN CumHT.LibU:=False ; CumHT.LibArt:='' ; END ;
    for kk:=1 to 5 do
    BEGIN
      CumHT.SommeTaxeD[kk]:=CumHT.SommeTaxeD[kk]+TOBP.GetValue('GPT_TOTALTAXEDEV'+IntToStr(kk)) ;
      CumHT.SommeTaxeP[kk]:=CumHT.SommeTaxeP[kk]+TOBP.GetValue('GPT_TOTALTAXE'+IntToStr(kk)) ;
    END ;
    CumulAnalPorcs(CumHT,TOBCode,XD,XP) ;
    if iTrouv<0 then TTA.Add(CumHT) ;
  END ;
  {Arrondissage des cumuls}
  for i:=0 to TTA.Count-1 do
  BEGIN
    CumHT:=T_CodeCpta(TTA[i]) ;
    CumHT.XD:=Arrondi(CumHT.XD,NbDec) ; CumHT.XP:=Arrondi(CumHT.XP,V_PGI.OkDecV) ;
    for kk:=1 to 5 do
    BEGIN
      CumHT.SommeTaxeD[kk]:=Arrondi(CumHT.SommeTaxeD[kk],NbDec) ;
      CumHT.SommeTaxeP[kk]:=Arrondi(CumHT.SommeTaxeP[kk],V_PGI.OkDecV) ;
    END ;
    for kk:=0 to CumHT.Anal.Count-1 do
    BEGIN
      XDT:=CumHT.Anal[kk] ;
      XDT.MD:=Arrondi(XDT.MD,NbDec) ;
      XDT.MP:=Arrondi(XDT.MP,V_PGI.OkDecV) ;
      XDT.TotQte1:=CumHT.Qte ;
    END ;
  END ;
  I := 0;
  {Ajustements par rapport au pied de facture}
    AjusteTaxesBases(TOBPiece,TOBBases,TOBPorcs,NbDec) ;
    AjusteHTBases(TOBPiece,TOBBases,TOBBasesCharge,TOBPorcs,R_ArtFi,NbDec) ;
  {Création des lignes d'écriture}
  for i:=0 to TTA.Count-1 do
  BEGIN
    if Result=rcOk then Result:=CreerLignesEcrHT(TOBEcr,TOBPiece,TOBTiers,MM,i) ;
    if Result<>rcOk then Break ;
  END ;
  end;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 21/03/2003
Modifié le ... : 21/03/2003
Description .. : Empêche la traduction comptable de montants négatifs si le
Suite ........ : paramétrage en comptabilité l'interdit. Dans ce cas,
Suite ........ : changement de signe et de sens de la ligne d'écriture
Mots clefs ... : COMPTABILISATION;NEGATIF;SIGNE;SENS;
*****************************************************************}
Procedure AnalyseLesNegatifs ( TOBEcr : TOB ) ;
Var TOBL  : TOB ;
    i     : integer ;
    XD,XC : double ;
BEGIN
if VH^.MontantNegatif then Exit ;
if TOBEcr=Nil then Exit ;
if TOBEcr.Detail.Count<=0 then Exit ;
for i:=0 to TOBEcr.Detail.Count-1 do
    BEGIN
    TOBL:=TOBEcr.Detail[i] ;
    XD:=TOBL.GetValue('E_DEBIT') ; XC:=TOBL.GetValue('E_CREDIT') ;
    if ((XD<0) or (XC<0)) then
       BEGIN
       TOBL.PutValue('E_CREDIT',-XD) ; TOBL.PutValue('E_DEBIT',-XC) ;
       {Devise}
       XD:=TOBL.GetValue('E_DEBITDEV') ; XC:=TOBL.GetValue('E_CREDITDEV') ;
       TOBL.PutValue('E_CREDITDEV',-XD) ; TOBL.PutValue('E_DEBITDEV',-XC) ;
       END ;
    END ;
END ;

Procedure MakeContrePTiers ( TOBEcr,TOBTiers : TOB ) ;
Var TOBHT,TOBL  : TOB ;
    i,iAux,iCpt : integer ;
    CptHT,Auxi : String ;
BEGIN
iAux := 0; iCpt := 0;
if PremHT<=0 then Exit ;
TOBHT:=TOBEcr.Detail[PremHT-1] ;
Auxi:=TOBTiers.GetValue('T_AUXILIAIRE') ; CptHT:=TOBHT.GetValue('E_GENERAL') ;
for i:=0 to TOBEcr.Detail.Count-1 do
    BEGIN
    TOBL:=TOBEcr.Detail[i] ;
    if i=0 then BEGIN iAux:=TOBL.GetNumChamp('E_AUXILIAIRE') ; iCpt:=TOBL.GetNumChamp('E_CONTREPARTIEGEN') ; END ;
    if TOBL.GetValeur(iAux)=Auxi then TOBL.PutValeur(iCpt,CptHT) ;
    END ;
END ;

Procedure GCVentilLigneTOB ( TOBAna : TOB ; Section : String ; NumVentil,NbDec : integer ; Pourc : double ; Debit : boolean ; Qte1,TotQte1 : double ) ;
BEGIN
VentilLigneTOB(TOBAna,Section,NumVentil,NbDec,Pourc,Debit) ;
TOBAna.PutValue('Y_QTE1',Qte1) ; TOBAna.PutValue('Y_TOTALQTE1',TotQte1) ;
Pourc:=0 ; if TotQte1<>0 then Pourc:=Arrondi(100.0*Qte1/TotQte1,4) ;
TOBAna.PutValue('Y_POURCENTQTE1',Pourc) ;
END ;

Procedure VentilerFromGC ( TOBEcr : TOB ; NbDec : integer ; TTAnal : TList ; NomC : String ) ;
Var CumHT     : T_CodeCpta ;
    IndiceA,i,NumAxe,NumVentil,Nb1,Nb2,Nb3 : integer ;
    XDT     : T_DetAnal ;
    Vent    : Array[1..5] of boolean ;
    TOBAxe,TOBAna  : TOB ;
    Axes,CptGene,Section : String ;
    Pourcentage,Qte1,TotQte1 : Double ;
    fb       : TFichierBase ;
BEGIN
IndiceA:=TOBEcr.GetValue(NomC) ; Axes:=TOBEcr.GetValue('AXES') ;
CumHT:=T_CodeCpta(TTAnal[IndiceA]) ; Nb1:=0 ; Nb2:=0 ; Nb3:=0 ;
CptGene:=TOBEcr.GetValue('E_GENERAL') ; FillChar(Vent,Sizeof(Vent),#0) ;
for NumAxe:=1 to 5 do Vent[NumAxe]:=(Pos(IntToStr(NumAxe),Axes)>0) ;
if ((Vent[1]) and (CumHT.XP<>0)) then
   BEGIN
   TOBAxe:=TOBEcr.Detail[0] ;
   for i:=0 to CumHT.Anal.Count-1 do
       BEGIN
       XDT:=T_DetAnal(CumHT.Anal[i]) ; if XDT.Ax<>'A1' then Continue ;
       inc(Nb1) ;
       TOBAna:=TOB.Create('ANALYTIQ',TOBAxe,-1) ; EcrVersAna(TOBEcr,TOBAna) ;
       Section:=XDT.Section ; NumVentil:=Nb1 ; Pourcentage:=Arrondi(100.0*XDT.MP/CumHT.XP,6) ;
       Qte1:=XDT.Qte1 ; TotQte1:=XDT.TotQte1 ;
       GCVentilLigneTOB(TOBAna,Section,NumVentil,NbDec,Pourcentage,TOBEcr.GetValue('E_DEBIT')<>0,Qte1,TotQte1) ;
       END ;
   if Nb1<=0 then
      BEGIN
      TOBAna:=TOB.Create('ANALYTIQ',TOBAxe,-1) ; EcrVersAna(TOBEcr,TOBAna) ;
      Section:=VH^.Cpta[fbAxe1].Attente ; NumVentil:=1 ; Pourcentage:=100.0 ; Qte1:=1 ; TotQte1:=0 ;
      GCVentilLigneTOB(TOBAna,Section,NumVentil,NbDec,Pourcentage,TOBEcr.GetValue('E_DEBIT')<>0,Qte1,TotQte1) ;
      END ;
   END ;
if ((Vent[2]) and (CumHT.XP<>0)) then
   BEGIN
   TOBAxe:=TOBEcr.Detail[1] ;
   for i:=0 to CumHT.Anal.Count-1 do
       BEGIN
       XDT:=T_DetAnal(CumHT.Anal[i]) ; if XDT.Ax<>'A2' then Continue ;
       inc(Nb2) ;
       TOBAna:=TOB.Create('ANALYTIQ',TOBAxe,-1) ; EcrVersAna(TOBEcr,TOBAna) ;
       Section:=XDT.Section ; NumVentil:=Nb2 ; Pourcentage:=Arrondi(100.0*XDT.MP/CumHT.XP,6) ;
       Qte1:=XDT.Qte1 ; TotQte1:=XDT.TotQte1 ;
       GCVentilLigneTOB(TOBAna,Section,NumVentil,NbDec,Pourcentage,TOBEcr.GetValue('E_DEBIT')<>0,Qte1,TotQte1) ;
       END ;
   if Nb2<=0 then
      BEGIN
      TOBAna:=TOB.Create('ANALYTIQ',TOBAxe,-1) ; EcrVersAna(TOBEcr,TOBAna) ;
      Section:=VH^.Cpta[fbAxe2].Attente ; NumVentil:=1 ; Pourcentage:=100.0 ; Qte1:=1 ; TotQte1:=1 ;
      GCVentilLigneTOB(TOBAna,Section,NumVentil,NbDec,Pourcentage,TOBEcr.GetValue('E_DEBIT')<>0,Qte1,TotQte1) ;
      END ;
   END ;
if ((Vent[3]) and (CumHT.XP<>0)) then
   BEGIN
   TOBAxe:=TOBEcr.Detail[2] ;
   for i:=0 to CumHT.Anal.Count-1 do
       BEGIN
       XDT:=T_DetAnal(CumHT.Anal[i]) ; if XDT.Ax<>'A3' then Continue ;
       inc(Nb3) ;
       TOBAna:=TOB.Create('ANALYTIQ',TOBAxe,-1) ; EcrVersAna(TOBEcr,TOBAna) ;
         // mcd 09/01/08 la ligne suivante a été reprise de VDEV où l'analytique sur 3eme axe fonctionne
       Section:=XDT.Section ; NumVentil:=Nb3 ; Pourcentage:=Arrondi(100.0*XDT.MP/CumHT.XP,6) ;
       Qte1:=XDT.Qte1 ; TotQte1:=XDT.TotQte1 ;
       GCVentilLigneTOB(TOBAna,Section,NumVentil,NbDec,Pourcentage,TOBEcr.GetValue('E_DEBIT')<>0,Qte1,TotQte1) ;
       END ;
   if Nb3<=0 then
      BEGIN
      TOBAna:=TOB.Create('ANALYTIQ',TOBAxe,-1) ; EcrVersAna(TOBEcr,TOBAna) ;
      Section:=VH^.Cpta[fbAxe3].Attente ; NumVentil:=1 ; Pourcentage:=100.0 ; Qte1:=1 ; TotQte1:=1 ;
      GCVentilLigneTOB(TOBAna,Section,NumVentil,NbDec,Pourcentage,TOBEcr.GetValue('E_DEBIT')<>0,Qte1,TotQte1) ;
      END ;
   END ;
for NumAxe:=4 to 5 do if Vent[NumAxe] then
    BEGIN
    fb:=AxeToFb('A'+IntToStr(NumAxe)) ; TOBAxe:=TOBEcr.Detail[NumAxe-1] ;
    if TOBAxe.Detail.Count<=0 then
       BEGIN
       TOBAna:=TOB.Create('ANALYTIQ',TOBAxe,-1) ; EcrVersAna(TOBEcr,TOBAna) ;
       Section:=VH^.Cpta[fb].Attente ;
       NumVentil:=1 ; Pourcentage:=100.0 ; Qte1:=1 ; TotQte1:=1 ;
       GCVentilLigneTOB(TOBAna,Section,NumVentil,NbDec,Pourcentage,TOBEcr.GetValue('E_DEBIT')<>0,Qte1,TotQte1) ;
       END ;
    END ;
ArrondirAnaTOB(TOBEcr,NbDec) ;
END ;

Procedure MakeAnalGC ( TOBEcr,TOBPiece : TOB ; MM : RMVT ; NbDec : integer ) ;
Var TOBL : TOB ;
    i,k,iAna  : integer ;
BEGIN
iAna := 0;
for i:=0 to TOBEcr.Detail.Count-1 do
    BEGIN
    TOBL:=TOBEcr.Detail[i] ;
    if i=0 then iAna:=TOBL.GetNumChamp('E_ANA') ;
    if TOBL.GetValeur(iAna)='X' then
       BEGIN
       for k:=1 to 5 do TOB.Create('A'+IntToStr(k),TOBL,-1) ;
       if TOBL.FieldExists('ANAL') then VentilerFromGC(TOBL,NbDec,TTA,'ANAL') else
        if TOBL.FieldExists('STOCK') then VentilerFromGC(TOBL,NbDec,TTStock,'STOCK') else
         if TOBL.FieldExists('VARSTK') then VentilerFromGC(TOBL,NbDec,TTVarStk,'VARSTK') else
            VentilerTOB(TOBL,'',0,NbDec,False) ;
       END ;
    END ;
END ;

Function CreerLigneRetenueDiv ( TOBEcr,TOBPiece,TOBPorcs,TOBBases,TOBTiers : TOB ; MM : RMVT ) : T_RetCpta ;
Var CPteRD   : String ;
    TOBG,TOBE,TOBTTC,TOBTVARG,TOBP : TOB ;
    OkVent    : boolean ;
    NumL      : integer ;
    DP,CP,DD,CD,XD,XP,XXp,XXD,TheMontantTTC : Double ;
    DEV : Rdevise;
    II : integer;
BEGIN
  Result:=rcOk ;
//  if GetParamSocSecur ('SO_BTOPTRGCOLLTIE',false) then exit;
  TOBTVARG := TOB.Create ('LES TVA RET',nil,-1);
  if TOBPorcs = nil then exit;
  if TOBPorcs.detail.count = 0 then exit;
  DEV.Code:=TOBPiece.GetValue('GP_DEVISE') ; GetInfosDevise(DEV) ;
  //
  //FV1 - 27/02/2018 : FS#2963 - TREUIL - Absence de mode de paiement dans l'écriture sur la ligne du Port et Frais
  TOBTTC:=Nil ;
  if TOBEcr.Detail.Count>0 then TOBTTC:=TOBEcr.Detail[0] ;
  //
  RecalculeRetenues (TOBpiece,TOBPorcs,TOBBases,DEV);
  //
  for II := 0 To TOBPorcs.detail.count -1 do
  begin
    TOBP := TOBPorcs.detail[II];
    if not TOBP.GetBoolean('GPT_RETENUEDIVERSE') then continue;
    {Montants}

// On garde les montants signés. Voir + bas : BRL 4/08
//    XP := Abs(TOBP.getDouble('GPT_TOTALTTC'));
//    XD := Abs(TOBP.getDouble('GPT_TOTALTTCDEV'));

    XP := TOBP.getDouble('GPT_TOTALTTC') ;
    XD := TOBP.getDouble('GPT_TOTALTTCDEV');
    if (XP =0) and (XD = 0) then continue;
    CPteRD := '';
    if getparamSocSecur('SO_BTVENTCOLLECTIF',false) then
    begin
      CPteRD := FindCollectifRDPlus (TOBP,TObpiece);
    end;
    //
    if CPteRD = '' then
    begin
      CPteRD := TOBP.GEtString('GPT_COLLECTIF');
    end;
    // Erreur sur Compte retenues de garanties
    if CPteRD='' then BEGIN Result:=rcPar ; LastMsg:=12 ; Exit ; END ;
    {Etude du compte général }
    TOBG:=CreerTOBGeneral(CPteRD) ;
    // Erreur sur Compte RG
    if TOBG=Nil then
       BEGIN
       if VH_GC.GCPontComptable='ATT' then Result:=rcPar else
        if VH_GC.GCPontComptable='REF' then Result:=rcRef else
           BEGIN
           if Not CreerCompteGC(TOBG,CPteRD,'','',ccpRG) then Result:=rcPar ;
           END ;
       if Result<>rcOk then BEGIN LastMsg:=2 ; Exit ; END ;
       END ;
    OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
    {Ligne d'écriture}
    TOBE:=TOB.Create('ECRITURE',TOBEcr,-1) ;
    PieceVersECR(MM,TOBPiece,TOBTiers,TOBE,False) ;
    {Général}
    TOBE.PutValue('E_GENERAL',CPteRD) ;
    TOBE.PutValue('E_CONSO',TOBG.GetValue('G_CONSO')) ;
    {Divers}
    TOBE.PutValue('E_TYPEMVT','TTC') ;
    {Client}
    TOBE.PutValue('E_AUXILIAIRE',TOBTiers.GetValue('T_AUXILIAIRE')) ;
    //FV1 : 27/02/2018 - FS#2957 - DSA - En écriture comptable le libellé des Ports & Frais est erroné
    if GetParamSocSecur('SO_LIBRETENUE', True) then AlimLibEcr(TobE,TobPiece,TobTiers,TOBP.GetString('GPT_LIBELLE'),tecRG,True,(MM.Simul='S'));
    //
    {Eche+Vent}
    NumL:=TOBEcr.Detail.Count-NbEches+1 ; TOBE.PutValue('E_NUMLIGNE',NumL) ;
    if OkVent then TOBE.PutValue('E_ANA','X') ;

//    if TOBG.GetValue('G_LETTRABLE')='X' then
    BEGIN
      TOBE.PutValue('E_ECHE','X') ; TOBE.PutValue('E_NUMECHE',1) ; TOBE.PutValue('E_ETATLETTRAGE','AL') ;
      TOBE.PutValue('E_ETAT','0000000000') ;
      //??????? c'est quoi ce truc ?????
      //if TOBTTC <> nil then
      if TOBTTC<>Nil then
      BEGIN
        TOBE.PutValue('E_MODEPAIE',TOBTTC.GetValue('E_MODEPAIE')) ;
        TOBE.PutValue('E_DATEECHEANCE',TOBTTC.GetValue('E_DATEECHEANCE')) ;
      END ;
    END ;
    TOBE.PutValue('E_EMETTEURTVA','X') ;
    {Montants}
    //GetSommeRG(TOBPieceRG,XP,XD,XE) ;

    DP:=0 ; CP:=0 ; DD:=0 ; CD:=0 ;
    if MM.Nature='FC' then BEGIN DD:=XD ; DP:=XP ; CD:=0 ; CP:=0 ; END else
    if MM.Nature='FF' then BEGIN CD:=XD ; CP:=XP ; DD:=0 ; DP:=0 ; END else
    if MM.Nature='AC' then BEGIN CD:=-XD ; CP:=-XP ; DD:=0 ; DP:=0 ; END else
    if MM.Nature='AF' then BEGIN DD:=-XD ; DP:=-XP ; CD:=0 ; CP:=0 ; END ;

    // si valeur négative, on inverse le sens : BRL 4/08
    if DP < 0 then
    begin
     CP:=-DP;
     DP:=0;
     end;
    if DD < 0 then
    begin
     CD:=-DD;
     DD:=0;
     end;

    TOBE.PutValue('E_DEBIT',DP)     ; TOBE.PutValue('E_CREDIT',CP) ;
    TOBE.PutValue('E_DEBITDEV',DD)  ; TOBE.PutValue('E_CREDITDEV',CD) ;
    if ((DP=0) and (CP=0)) then TOBE.Free else TheTOBRG := TOBE;
  end;
  TOBG.Free ;
  TOBTVARG.free;
end;

// Modif BTP
Function CreerLigneRetenueGar ( TOBEcr,TOBPiece,TOBPieceRG,TOBTiers,TOBBasesRG : TOB ; MM : RMVT ) : T_RetCpta ;
Var CPteRG   : String ;
    TOBG,TOBE,TOBTTC : TOB ;
    OkVent    : boolean ;
    NumL      : integer ;
    DP,CP,DD,CD,XD,XP,XXp,XXD,TheMontantRGTTC : Double ;
    DEV : Rdevise;
BEGIN
  Result:=rcOk ;
  if GetParamSocSecur ('SO_BTOPTRGCOLLTIE',false) then exit;
  TOBTTC:=Nil ; if TOBEcr.Detail.Count>0 then TOBTTC:=TOBEcr.Detail[0] ;
  if TOBPieceRg = nil then exit;
  if TOBPieceRg.detail.count = 0 then exit;
  DEV.Code:=TOBPiece.GetValue('GP_DEVISE') ; GetInfosDevise(DEV) ;
  {Montants}
  GetMontantRGReliquat(TOBPIeceRG, XD, XP,True);
  TheMontantRGTTC := XD;
  CpteRg := '';
  
  if getparamSocSecur('SO_BTVENTCOLLECTIF',false) then
  begin
    CpteRg := FindCollectifRGPlus (TOBBasesRG,TObpiece);
  end;
  
  if CpteRg = '' then
  begin
    if ((MM.Nature='FC') or (MM.Nature='AC')) then CpteRg:=GetParamSoc('SO_GCCPTERGVTE')
    else if ((MM.Nature='FF') or (MM.Nature='AF')) then CpteRg:=GetParamSoc('SO_GCCPTERGACH');
  end;
  
  // Erreur sur Compte retenues de garanties
  if CpteRG='' then BEGIN Result:=rcPar ; LastMsg:=12 ; Exit ; END ;
  {Etude du compte général }
  TOBG:=CreerTOBGeneral(CpteRG) ;
  // Erreur sur Compte RG
  if TOBG=Nil then
     BEGIN
     if VH_GC.GCPontComptable='ATT' then Result:=rcPar else
      if VH_GC.GCPontComptable='REF' then Result:=rcRef else
         BEGIN
         if Not CreerCompteGC(TOBG,CpteRg,'','',ccpRG) then Result:=rcPar ;
         END ;
     if Result<>rcOk then BEGIN LastMsg:=2 ; Exit ; END ;
     END ;
  OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
  {Ligne d'écriture}
  TOBE:=TOB.Create('ECRITURE',TOBEcr,-1) ;
  PieceVersECR(MM,TOBPiece,TOBTiers,TOBE,False) ;
  {Général}
  TOBE.PutValue('E_GENERAL',CpteRG) ;
  TOBE.PutValue('E_CONSO',TOBG.GetValue('G_CONSO')) ;
  {Divers}
  TOBE.PutValue('E_TYPEMVT','TTC') ;
  {Client}
  TOBE.PutValue('E_AUXILIAIRE',TOBTiers.GetValue('T_AUXILIAIRE')) ;
  AlimLibEcr(TobE,TobPiece,TobTiers,'Retenue Garantie',tecRG,True,(MM.Simul='S'));
  {Eche+Vent}
  NumL:=TOBEcr.Detail.Count-NbEches+1 ; TOBE.PutValue('E_NUMLIGNE',NumL) ;
  if OkVent then TOBE.PutValue('E_ANA','X') ;

  if TOBG.GetValue('G_LETTRABLE')='X' then
  BEGIN
    TOBE.PutValue('E_ECHE','X') ; TOBE.PutValue('E_NUMECHE',1) ; TOBE.PutValue('E_ETATLETTRAGE','AL') ;
    TOBE.PutValue('E_ETAT','0000000000') ;
    if TOBTTC <> nil then
    if TOBTTC<>Nil then
    BEGIN
      TOBE.PutValue('E_MODEPAIE',TOBTTC.GetValue('E_MODEPAIE')) ;
      TOBE.PutValue('E_DATEECHEANCE',TOBTTC.GetValue('E_DATEECHEANCE')) ;
    END ;
  END ;
  TOBE.PutValue('E_EMETTEURTVA','X') ;
  {Montants}
  //GetSommeRG(TOBPieceRG,XP,XD,XE) ;

  DP:=0 ; CP:=0 ; DD:=0 ; CD:=0 ;
  if MM.Nature='FC' then BEGIN DD:=XD ; DP:=XP ; CD:=0 ; CP:=0 ; END else
  if MM.Nature='FF' then BEGIN CD:=XD ; CP:=XP ; DD:=0 ; DP:=0 ; END else
  if MM.Nature='AC' then BEGIN CD:=-XD ; CP:=-XP ; DD:=0 ; DP:=0 ; END else
  if MM.Nature='AF' then BEGIN DD:=-XD ; DP:=-XP ; CD:=0 ; CP:=0 ; END ;
  TOBE.PutValue('E_DEBIT',DP)     ; TOBE.PutValue('E_CREDIT',CP) ;
  TOBE.PutValue('E_DEBITDEV',DD)  ; TOBE.PutValue('E_CREDITDEV',CD) ;
  if ((DP=0) and (CP=0)) then TOBE.Free else TheTOBRG := TOBE;
  TOBG.Free ;
END ;

Function CreerLignesTaxesRG ( TOBEcr,TOBPiece,TOBPieceRG,TOBTiers,TOBBases,TOBBasesRG : TOB ; MM : RMVT ) : T_RetCpta ;
Var i : integer ;
    TOBB,TOBE,TOBG,TOBTTC : TOB ;
    RGD,RGP : Double ;
    CD,CP,DD,DP : Double ;
    CTX,Regime,CatTaxe,FamTaxe,CodeTVA,CodeTPF : String ;
    Achat,OkVent,TvaEnc : boolean ;
    NumL  : integer ;
    XD,XP,TheMontantRGTTC : Double;
BEGIN
Result:=rcOk ;
(*
GetCumulRG(TOBPieceRG,XP,XD);
TheMontantRGTTC := XD;
GetReliquatCaution (TOBPieceRg,XP,XD);
TheMontantRGTTC := TheMontantRGTTC - XD;
if TheMontantRGTTC < 0 then TheMontantRGTTC := 0;
if TheMontantRGTTC <= 0 then exit;
*)
GetMontantRGReliquat(TOBPIeceRG, XD, XP,True);
TheMontantRGTTC := XD;

Regime:=TOBPiece.GetValue('GP_REGIMETAXE') ;
Achat:=((MM.Nature='AF') or (MM.Nature='FF')) ;
{#TVAENC} TvaEnc:=(Pos(TOBPiece.GetValue('GP_TVAENCAISSEMENT'),'TE;TM')>0);
for i:=0 to TOBBases.Detail.Count-1 do
    BEGIN
    CodeTVA:='' ; CodeTPF:='' ;
    TOBB:=TOBBases.Detail[i] ; GetLATaxeRG(TOBB,RGD,RGP) ;
    if RGD=0 then Continue ;
    CatTaxe:=TOBB.GetValue('PBR_CATEGORIETAXE') ; FamTaxe:=TOBB.GetValue('PBR_FAMILLETAXE') ;
{$IFDEF BTP}
		if BTTypeTaxe (CatTaxe) = bttTVA then
{$ELSE}
    if CatTaxe='TX1' then
{$ENDIF}
       BEGIN
       CTX:='' ;
       if TvaEnc then
        begin
        CTX:=TVA2ENCAIS(Regime,FamTaxe,Achat,true) ;
        // si le compte de Tva sur RG encaiss n'existe pas on prend celui qui n'est pas defini sur RG
        if Ctx = '' then CTX:=TVA2ENCAIS(Regime,FamTaxe,Achat) ;
        end;
       if CTX='' then CTX:=TVA2CPTE(Regime,FamTaxe,Achat) ;
       CodeTVA:=FamTaxe ;
       END else
       BEGIN
       CTX:='' ; if TvaEnc then CTX:=TPF2ENCAIS(Regime,FamTaxe,Achat) ;
       if CTX='' then CTX:=TPF2CPTE(Regime,FamTaxe,Achat) ;
       CodeTPF:=FamTaxe ;
       END ;
    // Erreur sur Compte Taxe
    if CTX='' then BEGIN Result:=rcPar ; LastMsg:=13 ; Break ; END ;
    {Etude du compte général de Taxe}
    TOBG:=CreerTOBgeneral(CTX) ;
    // Erreur sur Compte Taxe
    if TOBG=Nil then
       BEGIN
       if VH_GC.GCPontComptable='ATT' then Result:=rcPar else
        if VH_GC.GCPontComptable='REF' then Result:=rcRef else
           BEGIN
           if Not CreerCompteGC(TOBG,CTX,CodeTVA,CodeTPF,ccpTaxe) then Result:=rcPar ;
           END ;
       if Result<>rcOk then BEGIN LastMsg:=13 ; Break ; END ;
       END ;
    OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
    {Ligne d'écriture}
    TOBTTC:=Nil ; TOBTTC := TheTOBRG;
//    if TOBEcr.Detail.Count>0 then TOBTTC:=TOBEcr.Detail[0] ;
    TOBE:=TOB.Create('ECRITURE',TOBEcr,-1) ;
    PieceVersECR(MM,TOBPiece,TOBTiers,TOBE,False) ;
    {Général}
    TOBE.PutValue('E_GENERAL',CTX) ;
    TOBE.PutValue('E_CONSO',TOBG.GetValue('G_CONSO')) ;
    {Divers}
    AlimLibEcr(TobE,TobPiece,TobTiers,TOBG.GetValue('G_LIBELLE'),tecTaxeRG,True,(MM.Simul='S'));
{$IFDEF BTP}
		if BTTypeTaxe (CatTaxe) = bttTVA then
    BEGIN
       TOBE.PutValue('E_TYPEMVT','TVA') ; if PremTVA<0 then PremTVA:=TOBEcr.Detail.Count ;
    END else if BTTypeTaxe (CatTaxe) = bttTPF then
    BEGIN
    	TOBE.PutValue('E_TYPEMVT','TPF')
    END else
    BEGIN
    	TOBE.PutValue('E_TYPEMVT',CatTaxe) ;
    END;
{$ELSE}
    if CatTaxe='TX1' then
       BEGIN
       TOBE.PutValue('E_TYPEMVT','TVA') ; if PremTVA<0 then PremTVA:=TOBEcr.Detail.Count ;
       END else if CatTaxe='TX2' then TOBE.PutValue('E_TYPEMVT','TPF')
           else TOBE.PutValue('E_TYPEMVT',CatTaxe) ;
{$ENDIF}
    {Contreparties}
    TOBE.PutValue('E_CONTREPARTIEGEN',TOBTiers.GetValue('T_COLLECTIF')) ;
    TOBE.PutValue('E_CONTREPARTIEAUX',TOBTiers.GetValue('T_AUXILIAIRE')) ;
    TOBE.SetBoolean('E_TVAENCAISSEMENT', (Pos(TOBPiece.GetValue('GP_TVAENCAISSEMENT'),'TE;TM')>0));
    {Eche+Vent}
    NumL:=TOBEcr.Detail.Count-NbEches+1 ; TOBE.PutValue('E_NUMLIGNE',NumL) ;
    if OkVent then TOBE.PutValue('E_ANA','X') ;
    if TOBG.GetValue('G_LETTRABLE')='X' then
       BEGIN
       TOBE.PutValue('E_ECHE','X') ; TOBE.PutValue('E_NUMECHE',1) ; TOBE.PutValue('E_ETATLETTRAGE','AL') ;
       TOBE.PutValue('E_ETAT','0000000000') ;
       if TOBTTC<>Nil then
          BEGIN
          TOBE.PutValue('E_MODEPAIE',TOBTTC.GetValue('E_MODEPAIE')) ;
          TOBE.PutValue('E_DATEECHEANCE',TOBTTC.GetValue('E_DATEECHEANCE')) ;
          END ;
       END ;
    {Montants}
    DP:=0 ; CP:=0 ; DD:=0 ; CD:=0 ;
    if MM.Nature='FC' then BEGIN CD:=RGD ; CP:=RGP ; DD:=0 ; DP:=0 ; END else
    if MM.Nature='FF' then BEGIN DD:=RGD ; DP:=RGP ; CD:=0 ; CP:=0 ; END else
    if MM.Nature='AC' then BEGIN DD:=-RGD ; DP:=-RGP ; CD:=0 ; CP:=0 ; END else
    if MM.Nature='AF' then BEGIN CD:=-RGD ; CP:=-RGP ; DD:=0 ; DP:=0 ; END ;
    TOBE.PutValue('E_DEBIT',DP)     ; TOBE.PutValue('E_CREDIT',CP) ;
    TOBE.PutValue('E_DEBITDEV',DD)  ; TOBE.PutValue('E_CREDITDEV',CD) ;
    if ((DP=0) and (CP=0)) then TOBE.Free ;
    TOBG.Free ;
    END ;
END ;

// --

Procedure AjusteSurTVA ( TOBEcr : TOB ; EcartD,EcartP : Double ) ;
Var LaLig : integer ;
    TOBL : TOB ;
    DD,DP,CD,CP : Double ;
BEGIN
if PremTVA>0 then LaLig:=PremTVA else if PremHT>0 then LaLig:=PremHT else LaLig:=1 ;
TOBL:=TOBEcr.Detail[LaLig-1] ;
DD:=TOBL.GetValue('E_DEBITDEV')  ; CD:=TOBL.GetValue('E_CREDITDEV') ;
DP:=TOBL.GetValue('E_DEBIT')     ; CP:=TOBL.GetValue('E_CREDIT') ;
if DD<>0 then
   BEGIN
   TOBL.PutValue('E_DEBITDEV',DD-EcartD) ; TOBL.PutValue('E_DEBIT',DP-EcartP) ;
   END else
   BEGIN
   TOBL.PutValue('E_CREDITDEV',CD+EcartD) ; TOBL.PutValue('E_CREDIT',CP+EcartP) ;
   END ;
END ;

Procedure EquilibreEcrGC ( TOBEcr,TOBPiece,TOBTiers : TOB ; MM : RMVT ; NbDec : integer ) ;
Var DD,DP,CD,CP : Double ;
    i           : integer ;
    ModeOppose  : boolean ;
    TOBL        : TOB ;
    EcartP,EcartD : Double ;
BEGIN
DD:=0 ; CD:=0 ; DP:=0 ; CP:=0 ;
ModeOppose:=(TOBPiece.GetValue('GP_SAISIECONTRE')='X') ;
for i:=0 to TOBEcr.Detail.Count-1 do
    BEGIN
    TOBL:=TOBEcr.Detail[i] ;
    DD:=DD+TOBL.GetValue('E_DEBITDEV')  ; CD:=CD+TOBL.GetValue('E_CREDITDEV') ;
    DP:=DP+TOBL.GetValue('E_DEBIT')     ; CP:=CP+TOBL.GetValue('E_CREDIT') ;
    TOBL.PutValue('E_ENCAISSEMENT',SensEnc(DP,CP)) ;
    END ;
EcartD:=Arrondi(DD-CD,NbDec) ; EcartP:=Arrondi(DP-CP,V_PGI.OkDecV) ;
if ((EcartD=0) and (EcartP=0)) then Exit ;
if ((EcartD<>0) and (EcartP<>0)) then
   BEGIN
   if DelphiRunning then ShowMessage('Tout faux, Devise : '+StrfPoint(EcartD)+'   Pivot : '+StrfPoint(EcartP)) ;
   AjusteSurTVA(TOBEcr,EcartD,EcartP) ;
   END else if ((Not ModeOppose) and (EcartP<>0)) then
   BEGIN
   if DelphiRunning then ShowMessage('Saisie fausse, Devise : '+StrfPoint(EcartD)+'   Pivot : '+StrfPoint(EcartP)) ;
   AjusteSurTVA(TOBEcr,EcartD,EcartP) ;
   END else if EcartP<>0 then
   BEGIN
   if DelphiRunning then ShowMessage('Saisie devise juste et compta fausse, Devise : '+StrfPoint(EcartD)+'   Pivot : '+StrfPoint(EcartP)) ;
   AjusteSurTVA(TOBEcr,EcartD,EcartP) ;
   END else
   BEGIN
   {Ecart divers}
   AjusteSurTVA(TOBEcr,EcartD,EcartP) ;
   END ;
END ;

Procedure GCLettrerAcomptes (MM : RMVT ; TOBEcr,TOBAcomptes,TOBPiece : TOB ) ;
Var Q : TQuery ;
    SQL,Auxi,Gene,CodeL,CodeD,{SaisieEuro,}NatEcr,QualifP : String ;
    i,ie          : integer ;
    L             : TL_Rappro ;
    TLETT         : TList ;
    Achat         : boolean ;
    TOBCompta,TOBE,TOBEcrLiees,TOBACC,NewE,TOBLett,TOBTrouve : TOB ;
BEGIN
  if TOBAcomptes=Nil then Exit ;
  if TOBAcomptes.Detail.Count<=0 then Exit ;
  if TOBEcr.Detail.Count<=0 then Exit ;
  Auxi:=TOBEcr.Detail[0].GetValue('E_AUXILIAIRE') ; if Auxi='' then Exit ;
  Gene:=TOBEcr.Detail[0].GetValue('E_GENERAL')    ; if Gene='' then Exit ;
  CodeD:=TOBEcr.Detail[0].GetValue('E_DEVISE')    ;
  //SaisieEuro:=TOBEcr.Detail[0].GetValue('E_SAISIEEURO') ;
  if TOBPiece.GetValue('GP_VENTEACHAT')='ACH' then Achat:=True else Achat:=False ;
  {Lecture echéances d'acompte en comptabilité}
  QualifP:='N' ;
  TOBCompta:=TOB.Create('',Nil,-1) ;
  for i:=0 to TOBAcomptes.Detail.Count-1 do
  BEGIN
    TOBACC:=TOBAcomptes.Detail[i] ;
    if TOBACC.GetValue('GAC_ISREGLEMENT')='X' then BEGIN if Achat then NatEcr:='RF' else NatEcr:='RC' ; END
                else BEGIN if Achat then NatEcr:='OF' else NatEcr:='OC' ; END ;
    SQL:='SELECT * FROM ECRITURE WHERE E_AUXILIAIRE="'+Auxi+'" AND E_GENERAL="'+Gene+'" '
         +'AND E_QUALIFPIECE="'+QualifP+'" AND E_ECHE="X" AND (E_NATUREPIECE="'+NatEcr+'" OR E_NATUREPIECE="OD") AND (E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL") '
         +'AND E_DEVISE="'+CodeD+'" '
         +'AND E_JOURNAL="'+TOBACC.GetValue('GAC_JALECR')+'" AND E_NUMEROPIECE='+IntToStr(TOBACC.GetValue('GAC_NUMECR')) ;
    Q:=OpenSQL(SQL,True,-1, '', True) ;
    if Not Q.EOF then TOBCompta.LoadDetailDB('ECRITURE','','',Q,True) ;
    Ferme(Q) ;
  END ;
  if TOBCompta.Detail.Count<=0 then BEGIN TOBCompta.Free ; Exit ; END ;
  {Lecture des éventuelles écritures déjà lettrées avec ces acomptes}
  TOBEcrLiees:=TOB.Create('',Nil,-1) ;
  for i:=0 to TOBCompta.Detail.Count-1 do
  BEGIN
    TOBE:=TOBCompta.Detail[i] ;
    CodeL:=TOBE.GetValue('E_LETTRAGE') ; if CodeL='' then Continue ;
    SQL:='SELECT * FROM ECRITURE WHERE E_AUXILIAIRE="'+Auxi+'" AND E_GENERAL="'+Gene+'" '
         +'AND E_QUALIFPIECE="N" AND E_ECHE="X" AND E_ETATLETTRAGE="PL" '
         +'AND E_LETTRAGE="'+CodeL+'" AND E_DEVISE="'+CodeD+'" ' ;
    Q:=OpenSQL(SQL,True,-1, '', True) ;
    if Not Q.EOF then TOBEcrLiees.LoadDetailDB('ECRITURE','','',Q,True) ;
    Ferme(Q) ;
  END ;
  for i:=TOBEcrLiees.Detail.Count-1 downto 0 do
  BEGIN
    {Ne pas se retrouver soi-même}
    TOBLett:=TOBEcrLiees.Detail[i] ;
    TOBTrouve:=TOBCompta.FindFirst(['E_JOURNAL','E_NUMEROPIECE','E_NUMLIGNE','E_NUMECHE'],[TOBLett.GetValue('E_JOURNAL'),TOBLett.GetValue('E_NUMEROPIECE'),TOBLett.GetValue('E_NUMLIGNE'),TOBLett.GetValue('E_NUMECHE')],True) ;
    if TOBTrouve=Nil then TOBLett.ChangeParent(TOBCompta,-1) ;
  END ;
  TOBEcrLiees.Free ;
  {Rajout des échéances d'acompte de l'écriture de gescom}
  for i:=0 to TOBEcr.Detail.Count-1 do
  BEGIN
    TOBE:=TOBEcr.Detail[i] ; ie:=TOBE.GetNumChamp('E_ECHE') ;
    if TOBE.GetValeur(ie)<>'X' then Break ;
    if i>=NbAcc then Break ;
    NewE:=TOB.Create('ECRITURE',TOBCompta,-1) ;
    NewE.Dupliquer(TOBE,False,True) ;
  END ;
  {Constitution du paquet à lettrer}
  TLett:=TList.Create ;
  for i:=0 to TOBCompta.Detail.Count-1 do
  BEGIN
    TOBE:=TOBCompta.Detail[i] ; L:=TL_Rappro.Create ;
    {Comptes et caractéristiques}
    L.General:=Gene ; L.Auxiliaire:=Auxi ;
    L.DateC:=TOBE.GetValue('E_DATECOMPTABLE') ; L.DateE:=TOBE.GetValue('E_DATEECHEANCE') ; L.DateR:=TOBE.GetValue('E_DATEREFEXTERNE') ;
    L.RefI:=TOBE.GetValue('E_REFINTERNE') ; L.RefL:=TOBE.GetValue('E_REFLIBRE') ;
    L.RefE:=TOBE.GetValue('E_REFEXTERNE') ; L.Lib:=TOBE.GetValue('E_LIBELLE') ;
    L.Jal:=TOBE.GetValue('E_JOURNAL') ; L.Numero:=TOBE.GetValue('E_NUMEROPIECE') ;
    L.NumLigne:=TOBE.GetValue('E_NUMLIGNE') ; L.NumEche:=TOBE.GetValue('E_NUMECHE') ;
    L.CodeL:=TOBE.GetValue('E_LETTRAGE') ; L.CodeD:=CodeD ; L.TauxDEV:=TOBE.GetValue('E_TAUXDEV') ;
    if MM.CodeLettrage <> '' then L.CodeL := MM.CodeLettrage;
    L.Nature:=TOBE.GetValue('E_NATUREPIECE') ;
    L.Facture:=((L.Nature='FC') or (L.Nature='FF') or (L.Nature='AC') or (L.Nature='AF') or (L.Nature='OD')) ;
    L.Client:=((L.Nature='FC') or (L.Nature='AC') or (L.Nature='RC') or (L.Nature='OC') or (L.Nature='OD')) ;
    L.EditeEtatTva:=(TOBE.GetValue('E_EDITEETATTVA')='X') ;
    L.Solution:=0 ; L.Exo:=TOBE.GetValue('E_EXERCICE') ;
    {Montants, Lettrage}
    L.DebDev:=TOBE.GetValue('E_DEBITDEV') ; L.CredDev:=TOBE.GetValue('E_CREDITDEV') ;
    L.Debit:=TOBE.GetValue('E_DEBIT') ; L.Credit:=TOBE.GetValue('E_CREDIT') ;
    //    L.SaisieEuro:=(SaisieEuro='X') ;
    if L.CodeD<>V_PGI.DevisePivot then
    BEGIN
      L.DebitCur:=TOBE.GetValue('E_DEBITDEV') ; L.CreditCur:=TOBE.GetValue('E_CREDITDEV') ;
    END else
    BEGIN
      L.DebitCur:=TOBE.GetValue('E_DEBIT') ; L.CreditCur:=TOBE.GetValue('E_CREDIT') ;
    END ;
    {Objet}
    TLETT.Add(L) ;
  END ;
  {Lettrage}
  LettrerUnPaquet(TLett,False,False) ;
  {Libérations}
  VideListe(TLett) ; TLett.Free ;
  TOBCompta.Free ;
END ;

Procedure GCCreerPiecePayeur(TOBECR, TOBTiers: TOB; DocInfo : T_WSDocumentInf); //DocType : string; DocNumber : integer) ;
Var
  OkTp      : Boolean ;
  Jal       : string;
  Nat       : string ;
  XX        : RMVT ;
  TOBE      : TOB ;
  TobPayeur : TOB;
  SendEntry : TSendEntryY2;
begin
  {Test préalables}
  OkTP     := False;
  if    (not VH^.OuiTP)
     or (not assigned(TobEcr))
     or (not assigned(TobTiers))
     or (TobEcr.Detail.Count = 0)
  then
    exit;
  TobE := TobEcr.Detail[0];
  Jal  := TobE.GetString('E_JOURNAL');
  Nat  := TobE.GetString('E_NATUREPIECE') ;
  if   (TOBE.GetString('E_QUALIFPIECE') <> 'N')                                 // Pas écriture de type Normale
    or (not EstJalFact(Jal))                                                    // Pas un journal facturable
    or (pos(';' + Nat + ';', ';FC;FF;AC;AF;') = 0)                              // Pas une facture ou un avoir
    or (((Nat = 'FC') or (Nat = 'AC')) and (VH^.JalVTP = ''))                   // Flux vente et journal non renseigné
    or (((Nat = 'FF') or (Nat = 'AF')) and (VH^.JalATP = ''))                   // Flux achat et journal non renseigné
    or (((Nat = 'AC') or (Nat = 'AF')) and (TobTiers.GetBoolean('T_AVOIRRBT'))) // Avoir
  then
    exit;
  OkTp := (    (TOBTiers.GetString('T_PAYEUR') <> '')    // Possède un tiers payeur
           and (not TOBTiers.GetBoolean('T_ISPAYEUR'))); // N'est pas lui même un payeur
  if OkTP then
  begin
    {Création de la pièce payeur}
    FillChar(XX,Sizeof(XX),#0) ;
    XX.Jal       := Jal ;
    XX.Nature    := Nat ;
    XX.Simul     := 'N' ;
    XX.Valide    := False ;
    XX.ANouveau  := False ;
    XX.Etabl     := TOBE.GetString('E_ETABLISSEMENT') ;
    XX.DateC     := TOBE.GetDateTime('E_DATECOMPTABLE') ;
    XX.Exo       := QuelExoDT(XX.DateC) ;
    XX.CodeD     := TOBE.GetString('E_DEVISE') ;
    XX.DateTaux  := TOBE.GetDateTime('E_DATETAUXDEV') ;
    XX.Num       := TOBE.GetInteger('E_NUMEROPIECE') ;
    XX.TauxD     := TOBE.GetDouble('E_TAUXDEV') ;
    XX.DateE     := TOBE.GetDateTime('E_DATEECHEANCE') ;
    XX.DocType   := DocInfo.dType;
    XX.DocNumber := DocInfo.dNumber;
    TobPayeur := TOB.Create('_PAYEUR', nil, -1);
    try
      GenerePiecesPayeur(XX, '', nil, TobPayeur);
      if TobPayeur.Detail.count > 0 then
        CallWSCegid(wsetPayer, TobPayeur, DocInfo); //XX.DocType, XX.DocNumber);
    finally
      FreeAndNil(TobPayeur);
    end;
  end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 26/12/2000
Modifié le ... : 26/12/2000
Description .. : Enregistrement de la TOB Ecriture détaillée sous forme de
Suite ........ : texte dans un blob dans le cas d'un choix de lien comptable
Suite ........ : différé.
Mots clefs ... : COMPTABILITE;PASSATION;DIFFERE;
*****************************************************************}
Function InsertionDifferee ( TOBEcr : TOB ) : boolean ;
Var EcrDiff,TOBE : TOB ;
    StEcr,LaRef,LaNat,LeUser : String ;
    DD           : TDateTime ;
    LeRang       : integer ;
    Q            : TQuery ;
BEGIN
Result:=True ;
if TOBEcr=Nil then Exit ;
if TOBEcr.Detail.Count<=0 then Exit ;
TOBE:=TOBEcr.Detail[0] ;
EcrDiff:=TOB.Create('COMPTADIFFEREE',Nil,-1) ;
DD:=TOBE.GetValue('E_DATECOMPTABLE')      ; EcrDiff.PutValue('GCD_DATEPIECE',DD) ;
LaRef:=TOBE.GetValue('E_REFGESCOM')       ; EcrDiff.PutValue('GCD_REFPIECE',LaRef) ;
LaNat:=TOBE.GetValue('E_NATUREPIECE')     ; EcrDiff.PutValue('GCD_NATURECOMPTA',LaNat) ;
LeUser:=V_PGI.User                        ; EcrDiff.PutValue('GCD_USER',LeUser) ;
StEcr:=TOBEcr.SaveToBuffer(True,False,'') ; EcrDiff.PutValue('GCD_BLOBECRITURE',StEcr) ;
Q:=OpenSQL('SELECT MAX(GCD_RANG) FROM COMPTADIFFEREE WHERE GCD_DATEPIECE="'+UsDateTime(DD)+'" AND GCD_REFPIECE="'+LaRef+'" AND GCD_USER="'+LeUser+'"',True,-1, '', True) ;
if Not Q.EOF then LeRang:=Q.Fields[0].AsInteger+1 else LeRang:=0 ;
Ferme(Q) ;
EcrDiff.PutValue('GCD_RANG',LeRang) ;
Result:=EcrDiff.InsertDB(Nil) ;
EcrDiff.Free ;
END ;

procedure ConstituRestitutionAcpt (TOBEcr,TOBPiece,TOBTiers: TOB; DEV : RDevise;  var MM : RMVT) ;

  function OkVent(TobG : TOB) : boolean;
  begin
    Result := TobG.GetBoolean('G_VENTILABLE');
  end;

  procedure AjouteEcritureTiers(TOBEcr,TOBPiece,TOBTiers,TTX : TOB; var MM : RMVT);
  var
    TOBTTC,TOBG : TOB;
    GColl : String ;
    NumL : integer;
    ModePaie : string;
  begin
    GColl:=TOBTiers.GetValue('T_COLLECTIF') ;
    TOBG:=CreerTOBGeneral(GColl) ;
    //
    TOBTTC:=TOB.Create('ECRITURE',TOBEcr,-1) ;
    PieceVersECR(MM,TOBPiece,TOBTiers,TOBTTC,False) ;
    {Tiers}
    TOBTTC.PutValue('E_AUXILIAIRE',TOBTiers.GetValue('T_AUXILIAIRE')) ;
    TOBTTC.PutValue('E_GENERAL',GColl) ;
    TOBTTC.PutValue('E_CONSO',TOBTiers.GetValue('T_CONSO')) ;
    {Pièce}
    TOBTTC.PutValue('E_BLOCNOTE',TOBPiece.GetValue('GP_BLOCNOTE')) ;
    TOBTTC.PutValue('E_RIB',TTX.geTString('E_RIB')); // JT eQualité 11032
    {Divers}
    TOBTTC.PutValue('E_TYPEMVT','TTC') ;
    TOBTTC.PutValue('E_ETATLETTRAGE','AL') ;
    TOBTTC.PutValue('E_EMETTEURTVA','X') ;
    TOBTTC.PutValue('E_AFFAIRE',TTX.geTString('E_AFFAIRE')) ;
    TOBTTC.PutValue('E_LIBELLE',TTX.geTString('E_LIBELLE')) ;
    TOBTTC.PutValue('E_REFINTERNE',TTX.geTString('E_REFINTERNE')) ;
    {Eche+Vent}
    TOBTTC.PutValue('E_ECHE','X') ;
    TOBTTC.SetBoolean('E_ANA', OkVent(TobG)) ;
    NumL:=TOBEcr.Detail.Count-NbEches+1 ;
    TOBTTC.PutValue('E_NUMLIGNE',NumL) ; TOBTTC.PutValue('E_NUMECHE',1) ;
    {Echéances}
    ModePaie:=TTX.GetValue('E_MODEPAIE') ;
    TOBTTC.PutValue('E_MODEPAIE',ModePaie) ;
    TOBTTC.PutValue('E_DATEECHEANCE',TTX.GetValue('E_DATEECHEANCE')) ;
    TOBTTC.PutValue('E_CODEACCEPT',MPTOACC(ModePaie)) ;
    TOBTTC.PutValue('E_DATEPAQUETMIN',MM.DateC) ;
    TOBTTC.PutValue('E_DATEPAQUETMAX',MM.DateC) ;
    {Montants}
    TOBTTC.PutValue('E_DEBIT',0)     ; TOBTTC.PutValue('E_CREDIT',TOBPiece.getDouble ('GP_ACOMPTE')) ;
    TOBTTC.PutValue('E_DEBITDEV',0)  ; TOBTTC.PutValue('E_CREDITDEV',TOBPiece.getDouble ('GP_ACOMPTEDEV')) ;
    TOBG.free;
  end;

  procedure  AjouteEcritureHT (TOBEcr,TOBPiece,TOBTiers,TTX : TOB; var MM : RMVT);
  var
    TOBE,TOBG : TOB;
    CHT : string;
    result : T_RetCpta;
    NumL : integer;
  begin
    CHT := TTX.GetString('E_GENERAL');
    TOBG:=CreerTOBGeneral(CHT) ;
    // Erreur sur Compte HT
    if TOBG=Nil then
    BEGIN
      if VH_GC.GCPontComptable='ATT' then Result:=rcPar else
      if VH_GC.GCPontComptable='REF' then Result:=rcRef else Result:=rcPar ;
      if Result<>rcOk then BEGIN LastMsg:=5 ; Exit ; END ;
    END ;
    {Ligne d'écriture}
    TOBE:=TOB.Create('ECRITURE',TOBEcr,-1) ;
    PieceVersECR(MM,TOBPiece,TOBTiers,TOBE,False) ;
    CompteGenVersEcr (TOBPiece,TOBE,CHT );
    {Général}
    TOBE.PutValue('E_GENERAL',CHT) ;
    TOBE.PutValue('E_CONSO',TOBG.GetValue('G_CONSO')) ;
    {Divers}
    TOBE.PutValue('E_TYPEMVT','HT') ;
    TOBE.PutValue('E_QTE1',1) ;
    TOBE.PutValue('E_TVA',TTX.geTString('E_TVA')) ;
    TOBE.PutValue('E_TPF',TTX.geTString('E_TPF')) ;
    TOBE.PutValue('E_AFFAIRE',TTX.geTString('E_AFFAIRE')) ;
    TOBE.PutValue('E_LIBELLE',TTX.geTString('E_LIBELLE')) ;
    TOBE.PutValue('E_REFINTERNE',TTX.geTString('E_REFINTERNE')) ;
    {Contreparties}
    TOBE.PutValue('E_CONTREPARTIEGEN',TOBTiers.GetValue('T_COLLECTIF')) ;
    TOBE.PutValue('E_CONTREPARTIEAUX',TOBTiers.GetValue('T_AUXILIAIRE')) ;
    {Eche+Vent}
    NumL:=TOBEcr.Detail.Count-NbEches+1 ; TOBE.PutValue('E_NUMLIGNE',NumL) ;
    TOBE.SetBoolean('E_ANA', OkVent(TobG)) ;
    {Montants}
    TOBE.PutValue('E_DEBIT',TTX.GetDouble('MTRESTITUE')); TOBE.PutValue('E_CREDIT',0) ;
    TOBE.PutValue('E_DEBITDEV',TTX.GetDouble('MTRESTITUEDEV'))  ; TOBE.PutValue('E_CREDITDEV',0) ;
    TOBG.Free ;
  end;

  procedure  AjouteEcritureTVA (TOBEcr,TOBPiece,TOBTiers,TTX : TOB; var MM : RMVT);
  var
    TOBE,TOBG : TOB;
    CHT : string;
    result : T_RetCpta;
    NumL : integer;
  begin
    CHT := TTX.GetString('E_GENERAL');
    TOBG:=CreerTOBgeneral(CHT) ;
    TOBE:=TOB.Create('ECRITURE',TOBEcr,-1) ;
    PieceVersECR(MM,TOBPiece,TOBTiers,TOBE,False) ;
    {Général}
    TOBE.PutValue('E_GENERAL',TTX.GetString('E_GENERAL') ) ;
    TOBE.PutValue('E_CONSO',TTX.GetString('E_CONSO')) ;
    TOBE.PutValue('E_TYPEMVT',TTX.GetString('E_TYPEMVT')) ;
    {Contreparties}
    TOBE.PutValue('E_CONTREPARTIEGEN',TTX.GetValue('E_CONTREPARTIEGEN')) ;
    TOBE.PutValue('E_CONTREPARTIEAUX',TTX.GetValue('E_CONTREPARTIEAUX')) ;
    TOBE.SetBoolean('E_TVAENCAISSEMENT', TTX.GetBoolean('E_TVAENCAISSEMENT'));
    TOBE.PutValue('E_LIBELLE',TTX.GetString('E_LIBELLE')) ;
    TOBE.PutValue('E_REFINTERNE',TTX.GetString('E_REFINTERNE')) ;
    {Eche+Vent}
    NumL:=TOBEcr.Detail.Count-NbEches+1 ; TOBE.PutValue('E_NUMLIGNE',NumL) ;
    TOBE.SetBoolean('E_ANA', OKVent(TobG)) ;
    if TTX.GetValue('E_ECHE')='X' then
    BEGIN
      TOBE.PutValue('E_ECHE','X') ; TOBE.PutValue('E_NUMECHE',1) ; TOBE.PutValue('E_ETATLETTRAGE','AL') ;
      TOBE.PutValue('E_ETAT','0000000000') ;
      TOBE.PutValue('E_MODEPAIE',TTX.GetValue('E_MODEPAIE')) ;
      TOBE.PutValue('E_DATEECHEANCE',TTX.GetValue('E_DATEECHEANCE')) ;
    END ;
    {Montants}
    TOBE.PutValue('E_DEBIT',TTX.GetDouble('MTRESTITUE')); TOBE.PutValue('E_CREDIT',0) ;
    TOBE.PutValue('E_DEBITDEV',TTX.GetDouble('MTRESTITUEDEV'))  ; TOBE.PutValue('E_CREDITDEV',0) ;
  end;


var TT,TT1,TTX,TTM : TOB;
    MM1 : RMVT;
    Zcledoc : r_cledoc;
    QQ : TQuery;
    II : integer;
    ratio,Valmax,TC,TCDEV : double;
    SQL : string;
begin
  TT := TOB.Create ('ECRITURES ACPT',nil,-1);
  TRY
    FillChar(MM1,Sizeof(MM1),#0) ;
    DecoderefPiece (TOBPiece.getString('GP_ATTACHEMENT'),ZCledoc);
    QQ := OpenSql ('SELECT GP_REFCOMPTABLE FROM PIECE WHERE '+WherePiece (Zcledoc,TTdPiece,false),true,1,'',true);
    if not QQ.eof then
    begin
      MM1 := DecodeRefGCComptable(QQ.Fields[0].AsString);
    end;
    ferme (QQ);
    if MM1.Jal <> '' then
    begin
      SQl := 'SELECT ECRITURE.*, 0 AS MTRESTITUE, 0 AS MTRESTITUEDEV '+
             'FROM ECRITURE '+
             'WHERE  '+
             'E_EXERCICE="'+MM1.Exo+'" AND '+
             'E_JOURNAL="'+MM1.Jal +'" AND '+
             'E_QUALIFPIECE="'+MM1.Simul+'" AND '+
             'E_NUMEROPIECE='+InttoStr(MM1.Num)+' '+
             'ORDER BY E_NUMLIGNE';
      TT.LoadDetailDBFromSQL  ('XX1',SQl,false);
      if TT.detail.count > 0 then
      begin
        TT1 := nil;
        TTM := nil;
        TC := 0;  TCDEV := 0;
        for II := 0 to TT.detail.count -1 do
        begin
          TT1 := TT.detail[II];
          if TT1.getString('E_TYPEMVT')= 'TTC' then break;
        end;
        
        if TT1 = nil then exit;
        ratio := TOBpiece.getDouble('GP_ACOMPTE')/TT1.GetDouble('E_DEBIT');
        ValMAx := 0;
        for II := 0 to TT.detail.count -1 do
        begin
          TTX := TT.detail[II];
          if TTX.getString('E_TYPEMVT')='TTC' then
          begin
            TTX.SetDouble('MTRESTITUE',TOBpiece.getDouble('GP_ACOMPTE')); // ecriture sur tiers donc...
            TTX.SetDouble('MTRESTITUEDEV',TOBpiece.getDouble('GP_ACOMPTEDEV')); // ecriture sur tiers donc...
          end else
          begin
            if TTX.GetDouble('E_CREDIT') > valMax then
            begin
              TTM := TTX;
              ValMax := TTX.GetDouble('E_CREDIT');
            end;
            TTX.SetDouble('MTRESTITUE',Arrondi(TTX.GetDouble('E_CREDIT')*ratio,V_PGI.OkDecV)); // ecriture sur tiers donc...
            TTX.SetDouble('MTRESTITUEDEV',Arrondi(TTX.GetDouble('E_CREDITDEV')*ratio,V_PGI.OkDecV)); // ecriture sur tiers donc...
            TC := TC + TTX.GetDouble('MTRESTITUE');
            TCDEV := TCDEV + TTX.GetDouble('MTRESTITUEDEV');
          end;
        end;
        if arrondi(Tc,V_PGI.okdecV)<> TOBpiece.getDouble('GP_ACOMPTE') then
        begin
          TTM.SetDouble ('MTRESTITUE',arrondi(TTM.getDouble('MTRESTITUE')+ TOBpiece.getDouble('GP_ACOMPTE') - TC,V_PGI.okdecV));
          TTM.SetDouble ('MTRESTITUEDEV',Arrondi(TTM.getDouble('MTRESTITUEDEV')+ TOBpiece.getDouble('GP_ACOMPTEDEV') - TCDEV,V_PGI.okdecV));
        end;
        for II := 0 to TT.detail.count -1 do
        begin
          TTX := TT.detail[II];
          if TTX.getString('E_TYPEMVT')='TTC' then
          begin
            AjouteEcritureTiers(TOBEcr,TOBPiece,TOBTiers,TTX,MM);
          end else if TTX.getString('E_TYPEMVT')='HT' then
          begin
            AjouteEcritureHT(TOBEcr,TOBPiece,TOBTiers,TTX,MM);
          end else
          begin
            AjouteEcritureTVA(TOBEcr,TOBPiece,TOBTiers,TTX,MM);
          end;
        end;
      end;
    end;
  FINALLY
    TT.free;
  END;
end;

Function PasseEnCompta ( TOBPiece,TOBOuvrages,TOBOuvragesP,TOBBases,TOBBasesST,TOBBasesCharge,TOBEches,TOBpieceTrait,TOBAFFInterv,TOBTiers,TOBArticles,TOBCpta,TOBAcomptes,TOBPorcs,TOBPieceRG,TOBBasesRG,TOBAnaP,TOBAnaS,TOBPieceinterv,TOBVTECOLL : TOB ; NbDec : integer ; OldEcr : RMVT ; DEV : RDevise; NewPiece : Boolean) : T_RetCpta ;
Var
  QualifP,PassP,LienP,NatureG,NatCpta,JalCpta,RefCP,Domaine,NatureJal : String ;
  TOBEcr,TOBT,TOBBASRET : TOB ;
  MM          : RMVT ;
  NumCpta     : Longint ;
  ii          : integer ;
  QQ          : TQuery ;
  TOBBasesGLob : TOB;
  TOBrepart : TOB;
  I : integer;
  LesMilliemes,TvaEncaissement : string;
  ExePiece : string;
  SendEntry : TSendEntryY2;
  DocInfo : T_WSDocumentInf;
begin
  Result    := rcOk ;
  if GetParamSocSecur('SO_GCDESACTIVECOMPTA',false) then exit; // SI pas de compta on sort
  LastMsg   := -1;
  NbEches   := 0;
  NbAcc     := 0;
  NumCpta   := 0;
  NatureG   := TOBPiece.GetString('GP_NATUREPIECEG') ;
  Domaine   := TOBPiece.GetString('GP_DOMAINE') ;
  ExePiece  := CodeExo (TOBPiece.GetDateTime('GP_DATEPIECE'));
  LienP     := GetInfoParPiece(NatureG,'GPP_TYPEPASSCPTA') ;
  PassP     := GetInfoParPiece(NatureG,'GPP_TYPEECRCPTA') ;
  if ((PassP='') or (PassP='RIE')) then Exit ;
  FillChar(MM, Sizeof(MM),#0) ;
  MM.Etabl := TOBPiece.GetString('GP_ETABLISSEMENT');
  if (Not NewPiece) and (OldEcr.Jal <> '') and (OldEcr.Num <> 0) and (OldEcr.Exo = ExePiece) then
  begin
    NumCpta:=OldEcr.Num;
    JalCpta := oldEcr.Jal;
    QualifP := oldEcr.simul;
    NatCpta := OldEcr.nature;
    MM.Jal := OldEcr.Jal;
    MM.nature := NatCpta;
    MM.Simul:=QualifP ;
  end else
  begin
    NewPiece := True;
    JalCpta  := '';
    if LienP='DIF' then PassP:='SIM' ;
    {Passation différée --> passer en simu pour ne pas affecter les soldes et la numérotation}
    QualifP :='N';
    if      PassP = 'PRE' then QualifP := 'P'
    else if PassP = 'SIM' then QualifP := 'S';
    {Journal}
    TOBT := TOBTiers.FindFirst(['GTP_NATUREPIECEG'],[NatureG],False) ;
    if TOBT <> Nil then JalCpta := TOBT.GetValue('GTP_JOURNALCPTA');
    if ((JalCpta='') and (Domaine<>'')) then JalCpta:=GetInfoParPieceDomaine(NatureG,Domaine,'GDP_JOURNALCPTA') ;
    if JalCpta='' then JalCpta := GetInfoParPieceCompl(NatureG,MM.Etabl,'GPC_JOURNALCPTA') ;
    if JalCpta='' then JalCpta := GetInfoParPiece(NatureG,'GPP_JOURNALCPTA') ;
    // Erreur sur Journal
    if JalCpta='' then
    begin
      Result:=rcPar ;
      LastMsg:=7 ;
      Exit ;
    end;
    TOBPiece.SetString('GP_JALCOMPTABLE', JalCpta) ;
    MM.Jal := JalCpta ;
    {Nature}
    NatCpta := GetInfoParPiece(NatureG, 'GPP_NATURECPTA') ;
    // Erreur sur la nature
    if NatCpta = '' then
    begin
      Result:=rcPar ;
      LastMsg:=8 ;
      Exit ;
    end;
    MM.Nature := NatCpta ;
  end;
  TTCSansTVA       := ((TOBPiece.GetString('GP_VENTEACHAT')='ACH') and (GetParamSocSecur('SO_GCACHATTTC', False)));
  TOBEcr           := TOB.Create('COMPTABILITE', Nil, -1);  // TOB Ecriture Générique
  TOBBasesGlob     := TOB.Create ('BASESTVAGLOB',nil,-1);
  TOBRepart        := TOB.Create ('REPARTITION TVA',nil,-1);
  TobTVASurEncaiss := TOB.Create('TVA ENCAISSEMENT', nil, -1); // Tob des montant tva sur encaissement
  try
    // Repartition TVA au 1/1000e
    if TOBPiece.FieldExists('RUPTMILLIEME') then
    begin
      LesMilliemes := TOBPiece.GetValue('RUPTMILLIEME');
    end else
    begin
      LesMilliemes := '';
    end;
    DecodeRepartTva (LesMilliemes,TOBRepart);
    if RGComptabilise then FusionneTva ( TobBasesRg,TOBBasesGlob,TOBPIECERG,TOBRepart,Pos(NatureG,'ABT;ABP')>0);
    // Création du RECORD clef comptable
    MM.Simul:=QualifP ;
    RenseigneClefCompta(TOBPiece,MM) ; PremHT:=-1 ; PremTVA:=-1 ;
    //Si pièce de caisse en différée, gestion des règlements avec la table PIEDECHE)
    if (MM.Simul = 'S') and (IsNatCaisse(NatureG)) then
    begin
      Result := CreerLignesRgtCaisse(TobEcr, TobPiece, TobEches, TobTiers, MM);
      if Result = rcOk then
        TobEcr.ClearDetail;
    end;

    {Valide}
    MM.Valide := (QualifP = 'N');
    {Numéro}
    if (OldEcr.Jal='') or (NewPiece) then
      NumCpta:=GetNewNumJal(JalCpta,(MM.Simul='N'),MM.DateC) ;

    {Type de journal}
    MM.ModeSaisieJal:='-' ;
    QQ:=OpenSQL('SELECT J_MODESAISIE, J_NATUREJAL FROM JOURNAL WHERE J_JOURNAL="'+MM.Jal+'"',True,-1, '', True) ;
    if Not QQ.EOF then
    begin
      MM.ModeSaisieJal:=QQ.Fields[0].AsString ;
      NatureJal := QQ.Fields[1].AsString ;
    end;
    Ferme(QQ) ;
    // Erreur sur le numéro
    if NumCpta <= 0 then
    begin
      Result := rcPar ;
      LastMsg := 9 ;
      Exit;
    end;
    //
    GereTvaMixte := OkTvaMixte(TOBTiers, MM, NatureJal);
    //
    MM.Num:=NumCpta ;
    {Préparer analytiques}
    TTA:=TList.Create ;
    if TOBPiece.geTDouble('GP_TOTALTTCDEV')<> 0 then
    begin
      {Lignes de Tiers}
      Result:=CreerLignesTiers(TOBEcr,TOBPiece,TOBEches,TOBTiers,TOBAcomptes,TOBPieceRg,TOBbasesRG,TOBpieceTrait,TOBPorcs,TOBVTECOLL,DEV, MM) ;
      {Lignes HT}
      if Result=rcOk then Result:=CreerLignesHTTva(TOBEcr,TOBAFFInterv,TOBPieceinterv,TOBPiece,TOBOuvrages,TOBOuvragesP,TOBBases,TOBBasesCharge,TOBTiers,TOBArticles,TOBCpta,TOBPorcs,TOBanaP,TOBAnaS,MM,NbDec,DEV) ;
      {Taxes}
      if Result=rcOk then Result:=CreerLignesTaxes(TOBEcr,TOBPiece,TOBBases,TOBBasesST,TOBBasesCharge,TOBTiers,TOBPIeceRG,TOBBASESGlob,TOBPorcs,MM) ;
      {Escompte et remises}
      if Result=rcOk then Result:=CreerLigneRemise(TOBEcr,TOBPiece,TOBTiers,TOBanaP,TOBAnaS,TOBCpta,MM) ;
      if Result=rcOk then Result:=CreerLigneEscompte(TOBEcr,TOBPiece,TOBTiers,MM) ;
      // Modif BTP
      {Lignes Retenu de garantie}
      if RGComptabilise then
         begin
         if Result=rcOk then Result:=CreerLigneRetenueGar(TOBEcr,TOBPIece,TOBPieceRG,TOBTiers,TOBBasesRG,MM) ;
         if Result=rcOk then Result:=CreerLigneRetenueDiv(TOBEcr,TOBPIece,TOBPorcs,TOBBases,TOBTiers,MM) ;
      //   if Result=rcOk then Result:=CreerLignesTaxesRG(TOBEcr,TOBPiece,TOBPieceRG,TOBTiers,TOBBASESGlob,TOBBasesRg,MM) ;
         end;
      if (GetparamSocSecur('SO_BTCPTAPAIEDIRECT',false)) and (not GetparamSocSecur('SO_BTREGLSTTIERS',true)) then
      begin
        if Result=rcOk then Result:=CreerPaiementDirect(TOBEcr,TOBPiece,TOBEches,TOBTiers,TOBAcomptes,TOBPieceRg,TOBbasesRG,TOBpieceTrait,TOBPorcs,DEV, MM) ;
      end;
      if (not GetParamsocSecur ('SO_BTCOMPTAREGL',false)) and (GetParamsocSecur ('SO_ACOMPTESFAC', false)) and (MM.Nature='FC') then
      begin
        // contitution de la partie d'écriture comptable correspondante à la deduction de l'acompte
        if Result=rcOk then
        begin
          if (TOBpiece.getString('GP_ATTACHEMENT') <> '') and (tobpiece.getDouble('GP_ACOMPTE')<>0)  then ConstituRestitutionAcpt (TOBEcr,TOBPiece,TOBTiers,DEV, MM) ;
        end;
      end;
      // ------
    end else
    begin
      result := CreerEcrZero(TOBEcr,TOBPiece,TOBTiers,MM) ;
    end;
    if ((Result=rcOk) and (TOBEcr.Detail.Count>=2)) then
    begin
      {Ultimes ajustements}
      AnalyseLesNegatifs(TOBEcr) ;
      MakeContrePTiers(TOBEcr,TOBTiers) ;
      {Analytiques}
      MakeAnalGC(TOBEcr,TOBPiece,MM,NbDec) ;
      {Retour inverse}
      if LienP <> 'DIF' then
        RefCP := EncodeRefGCComptable(TOBEcr.Detail[0])
      else
        RefCP := 'DIFF' ;
      TOBPiece.SetString('GP_REFCOMPTABLE',RefCP) ;
      {Equilibrage}
      if TOBPiece.GeTDouble('GP_TOTALTTCDEV')<> 0 then
        EquilibreEcrGC(TOBEcr, TOBPiece, TOBTiers, MM, NbDec) ;
      {MAJ des soldes comptables}
      if QualifP = 'N' then
      begin
        MajSoldesEcritureTOB(TOBEcr,True) ;
        if V_PGI.IoError<>oeOk then
        begin
          Result  := rcMaj ;
          LastMsg := 10 ;
        end;
      end;
      if Result=rcOk then
      begin
        {Insertion}
        if LienP<>'DIF' then
        begin
          DocInfo := GetWSDocumentInfFromTob(TobPiece);
          TobEcr.AddChampSup('IBAN', False);
          if Not TOBEcr.InsertDBByNivel(False) then
            V_PGI.IoError := oeUnknown
          else
          if TGetParamWSCEGID.ConnectToY2 then
          begin
            { Si RIB renseigné, recherche IBAN correspondant }
            if TOBEcr.Detail[0].GetString('E_RIB') <> '' then
            begin
              QQ := OpenSQL('SELECT R_CODEIBAN FROM RIB WHERE R_AUXILIAIRE = "' + TobTiers.GetString('T_AUXILIAIRE') + '" AND R_PRINCIPAL = "X"', true);
              try
                TOBEcr.SetString('IBAN', QQ.Fields[0].AsString);
              finally
                Ferme(QQ);
              end;
            end else
              TOBEcr.SetString('IBAN', '');
            CallWSCegid(wsetDocument, TOBEcr,DocInfo);
          end;
          {Lettrage}
          if QualifP='N' then GCLettrerAcomptes(OldEcr,TOBEcr,TOBAcomptes,TOBPiece) ;
          {Tiers payeur}
          if QualifP='N' then GCCreerPiecePayeur(TOBECR, TOBTiers, DocInfo); 
          {TVA encaissement}
          {$IFDEF BTP}
          if (QualifP='N') or (QualifP='S') then
            ElementsTvaEnc(MM,False,TOBPiece.GetString('GP_TVAENCAISSEMENT'));
          {$ELSE BTP}
           if QualifP='N' then ElementsTvaEnc(MM,False,TOBPiece.GetValue('GP_TVAENCAISSEMENT')) ;
          {$ENDIF BTP}
          TOBEcr.DelChampSup('IBAN', False);
          {Fermeture de la piece apres comptabilisation} //AC
          if GetInfoParPiece(NatureG,'GPP_ACTIONFINI')='COM' then
          begin
            TOBPiece.SetBoolean('GP_VIVANTE', False) ;
            for ii := 0 to TOBPiece.Detail.count-1 do
              TOBPiece.Detail[ii].SetBoolean('GL_VIVANTE', False) ;
          end;
        end else
        begin
           if Not InsertionDifferee(TOBEcr) then
            V_PGI.IoError := oeUnknown ;
        end;
      end;
    end;
  finally
    FreeAndNil(TobEcr);
    FreeAndNil(TOBBasesGlob);
    FreeAndNil(TOBRepart);
    FreeAndNil(TobTVASurEncaiss);
  end;
  for ii:=0 to TTA.Count-1 do
  begin
    VideListe(T_CodeCpta(TTA[ii]).Anal) ;
    T_CodeCpta(TTA[ii]).Anal.Free ;
    T_CodeCpta(TTA[ii]).Anal:=Nil ;
  end;
  VideListe(TTA) ;
  TTA.Free ;
  TTA := Nil ;
end;

Function CreerLignesEcrStock ( TOBEcr,TOBPiece,TOBTiers : TOB ; MM : RMVT ; IndiceCpta : integer ) : T_RetCpta ;
Var CumStock : T_CodeCpta ;
    CHT : String ;
    TOBG,TOBE : TOB ;
    XD,XP : Double ;
    CD,CP : Double ;
    DD,DP : Double ;
    OkVent,EstAvoir : boolean ;
    NumL,i : integer ;
    Sta,LibDefaut,Nature,VA : String ;
BEGIN
  Result:=rcOk ;
  CumStock:=T_CodeCpta(TTStock[IndiceCpta]) ;
  // Erreur sur Compte HT
  CHT:=CumStock.CptHT ; if CHT='' then BEGIN Result:=rcPar ; LastMsg:=11 ; Exit ; END ;
  {Etude du compte général HT}
  TOBG:=CreerTOBGeneral(CHT) ;
  // Erreur sur Compte HT
  if TOBG=Nil then
     BEGIN
     if VH_GC.GCPontComptable='ATT' then Result:=rcPar else
      if VH_GC.GCPontComptable='REF' then Result:=rcRef else
         BEGIN
         if Not CreerCompteGC(TOBG,CHT,CumStock.FamTaxe[1],CumStock.FamTaxe[2],ccpStock,MM.Nature) then Result:=rcPar ;
         END ;
     if Result<>rcOk then BEGIN LastMsg:=11 ; Exit ; END ;
     END ;
  OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
  {Ligne d'écriture}
  TOBE:=TOB.Create('ECRITURE',TOBEcr,-1) ;
  if CumStock.Anal.Count>0 then
     BEGIN
     TOBE.AddChampSup('STOCK',False) ; TOBE.AddChampSup('AXES',False) ;
     TOBE.PutValue('STOCK',IndiceCpta) ;
     Sta:='' ; for i:=1 to 5 do if TOBG.GetValue('G_VENTILABLE'+IntToStr(i))='X' then Sta:=Sta+IntToStr(i) ;
     TOBE.PutValue('AXES',Sta) ;
     END ;
  PieceVersECR(MM,TOBPiece,TOBTiers,TOBE,False) ;
  {Général}
  TOBE.PutValue('E_GENERAL',CHT) ;
  TOBE.PutValue('E_CONSO',TOBG.GetValue('G_CONSO')) ;
  {Divers}
  TOBE.PutValue('E_TYPEMVT','DIV') ;
  TOBE.PutValue('E_QTE1',CumStock.Qte) ;
  TOBE.PutValue('E_TVA',CumStock.FamTaxe[1]) ; TOBE.PutValue('E_TPF',CumStock.FamTaxe[2]) ;
  TOBE.PutValue('E_REFEXTERNE',CumStock.Depot) ;
  if CumStock.Affaire<>'' then TOBE.PutValue('E_AFFAIRE',CumStock.Affaire) ;
  LibDefaut:=TOBTiers.GetValue('T_LIBELLE') ;
  AlimLibEcr(TobE,TobPiece,TobTiers,LibDefaut,tecStock,True,(MM.Simul='S'));
  {Eche+Vent}
  NumL:=TOBEcr.Detail.Count ; TOBE.PutValue('E_NUMLIGNE',NumL) ;
  if OkVent then TOBE.PutValue('E_ANA','X') ;
  {Montants}
  XD:=CumStock.XD ; XP:=CumStock.XP ;
  VA:=TOBPiece.GetValue('GP_VENTEACHAT') ; Nature:=TOBPiece.GetValue('GP_NATUREPIECEG') ;
  EstAvoir:=(GetInfoParPiece(Nature,'GPP_ESTAVOIR')='X');
  if VA='VEN' then
     BEGIN
     if EstAvoir then BEGIN DD:=-XD ; DP:=-XP ; CD:=0 ; CP:=0 ; END
                 else BEGIN CD:=XD ; CP:=XP ; DD:=0 ; DP:=0 ; END ;
     END else
     BEGIN
     if EstAvoir then BEGIN CD:=-XD ; CP:=-XP ; DD:=0 ; DP:=0 ; END
                 else BEGIN DD:=XD ; DP:=XP ; CD:=0 ; CP:=0 ; END ;
     END ;
  TOBE.PutValue('E_DEBIT',DP)     ; TOBE.PutValue('E_CREDIT',CP) ;
  TOBE.PutValue('E_DEBITDEV',DD)  ; TOBE.PutValue('E_CREDITDEV',CD) ;
  if ((DP=0) and (CP=0)) then TOBE.Free ;
  TOBG.Free ;
END ;

Function CreerLignesEcrVarStk ( TOBEcr,TOBPiece,TOBTiers : TOB ; MM : RMVT ; IndiceCpta : integer ) : T_RetCpta ;
Var CumVarStk : T_CodeCpta ;
    CHT : String ;
    TOBG,TOBE : TOB ;
    XD,XP : Double ;
    CD,CP : Double ;
    DD,DP : Double ;
    OkVent,EstAvoir : boolean ;
    NumL,i : integer ;
    Sta,LibDefaut,VA,Nature  : String ;
BEGIN
Result:=rcOk ;
CumVarStk:=T_CodeCpta(TTVarStk[IndiceCpta]) ;
// Erreur sur Compte HT
CHT:=CumVarStk.CptHT ; if CHT='' then BEGIN Result:=rcPar ; LastMsg:=11 ; Exit ; END ;
{Etude du compte général HT}
TOBG:=CreerTOBGeneral(CHT) ;
// Erreur sur Compte HT
if TOBG=Nil then
   BEGIN
   if VH_GC.GCPontComptable='ATT' then Result:=rcPar else
    if VH_GC.GCPontComptable='REF' then Result:=rcRef else
       BEGIN
       if Not CreerCompteGC(TOBG,CHT,CumVarStk.FamTaxe[1],CumVarStk.FamTaxe[2],ccpVarStk,MM.Nature) then Result:=rcPar ;
       END ;
   if Result<>rcOk then BEGIN LastMsg:=11 ; Exit ; END ;
   END ;
OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
{Ligne d'écriture}
TOBE:=TOB.Create('ECRITURE',TOBEcr,-1) ;
if CumVarStk.Anal.Count>0 then
   BEGIN
   TOBE.AddChampSup('VARSTK',False) ; TOBE.AddChampSup('AXES',False) ;
   TOBE.PutValue('VARSTK',IndiceCpta) ;
   Sta:='' ; for i:=1 to 5 do if TOBG.GetValue('G_VENTILABLE'+IntToStr(i))='X' then Sta:=Sta+IntToStr(i) ;
   TOBE.PutValue('AXES',Sta) ;
   END ;
PieceVersECR(MM,TOBPiece,TOBTiers,TOBE,False) ;
{Général}
TOBE.PutValue('E_GENERAL',CHT) ;
TOBE.PutValue('E_CONSO',TOBG.GetValue('G_CONSO')) ;
{Divers}
TOBE.PutValue('E_TYPEMVT','DIV') ;
TOBE.PutValue('E_QTE1',CumVarStk.Qte) ;
TOBE.PutValue('E_TVA',CumVarStk.FamTaxe[1]) ; TOBE.PutValue('E_TPF',CumVarStk.FamTaxe[2]) ;
TOBE.PutValue('E_REFEXTERNE',CumVarStk.Depot) ;
if CumVarStk.Affaire<>'' then TOBE.PutValue('E_AFFAIRE',CumVarStk.Affaire) ;
LibDefaut:=TOBTiers.GetValue('T_LIBELLE') ;
AlimLibEcr(TobE,TobPiece,TobTiers,LibDefaut,tecStock,True,(MM.Simul='S'));
{Eche+Vent}
NumL:=TOBEcr.Detail.Count ; TOBE.PutValue('E_NUMLIGNE',NumL) ;
if OkVent then TOBE.PutValue('E_ANA','X') ;
{Montants}
XD:=CumVarStk.XD ; XP:=CumVarStk.XP ;
VA:=TOBPiece.GetValue('GP_VENTEACHAT') ; Nature:=TOBPiece.GetValue('GP_NATUREPIECEG') ;
EstAvoir:=(GetInfoParPiece(Nature,'GPP_ESTAVOIR')='X');
if VA<>'VEN' then
   BEGIN
   if EstAvoir then BEGIN DD:=-XD ; DP:=-XP ; CD:=0 ; CP:=0 ; END
               else BEGIN CD:=XD ; CP:=XP ; DD:=0 ; DP:=0 ; END ;
   END else
   BEGIN
   if EstAvoir then BEGIN CD:=-XD ; CP:=-XP ; DD:=0 ; DP:=0 ; END
               else BEGIN DD:=XD ; DP:=XP ; CD:=0 ; CP:=0 ; END ;
   END ;
TOBE.PutValue('E_DEBIT',DP)     ; TOBE.PutValue('E_CREDIT',CP) ;
TOBE.PutValue('E_DEBITDEV',DD)  ; TOBE.PutValue('E_CREDITDEV',CD) ;
if ((DP=0) and (CP=0)) then TOBE.Free ;
TOBG.Free ;
END ;

Procedure RecupLesDCStock ( TOBL,TOBArticles : TOB ; Achat : boolean ; Var XD,XP : Double ) ;
Var TOBArt,TOBDispo : TOB ;
    Depot,RefUnique,ModeValo : String ;
    Qte                      : Double ;
    prefixe : string;
BEGIN
prefixe := GetPrefixeTable (TOBL);

XD:=0 ; XP:=0 ;
Depot:=TOBL.GetValue(prefixe+'_DEPOT') ; if Depot='' then Depot:=TOBL.Parent.GetValue('GP_DEPOT') ;
RefUnique:=TOBL.GetValue(prefixe+'_ARTICLE') ; if RefUnique='' then Exit ;
Qte:=TOBL.GetValue(prefixe+'_QTESTOCK') ;
TOBArt:=TOBArticles.FindFirst(['GA_ARTICLE'],[RefUnique],False) ; if TOBArt=Nil then Exit ;
TOBDispo:=TOBArt.FindFirst(['GQ_ARTICLE','GQ_DEPOT'],[RefUnique,Depot],False) ;
if Achat then
   BEGIN
   if TTCSANSTVA then XD:=TOBL.GetValue(prefixe+'_MONTANTTTC') else XD:=TOBL.GetValue(prefixe+'_MONTANTHT') ;
   XP:=XD ;
   END else
   BEGIN
   ModeValo:=GetParamSoc('SO_GCMODEVALOSTOCK') ;
   if TOBDispo<>Nil then
      BEGIN
      if ModeValo='PMA' then XD:=TOBDispo.GetValue('GQ_PMAP') else
       if ModeValo='PMR' then XD:=TOBDispo.GetValue('GQ_PMRP') else
        if ModeValo='DPA' then XD:=TOBDispo.GetValue('GQ_DPA') else
         if ModeValo='DPR' then XD:=TOBDispo.GetValue('GQ_DPR') else
            XD:=TOBDispo.GetValue('GQ_PMAP') ;
      END else
      BEGIN
      if ModeValo='PMA' then XD:=TOBArt.GetValue('GA_PMAP') else
       if ModeValo='PMR' then XD:=TOBArt.GetValue('GA_PMRP') else
        if ModeValo='DPA' then XD:=TOBArt.GetValue('GA_DPA') else
         if ModeValo='DPR' then XD:=TOBArt.GetValue('GA_DPR') else
            XD:=TOBDispo.GetValue('GA_PMAP') ;
      END ;
   XP:=XD ; XD:=Arrondi(XD*Qte,V_PGI.OkDecV) ; XP:=XD ;
   END ;
END ;

Function CreerLignesStock ( TOBEcr,TOBPiece,TOBTiers,TOBArticles,TOBCpta : TOB ; MM : RMVT ; NbDec : integer ) : T_RetCpta ;
Var i,k,itrouv,kk : integer ;
    TOBA,TOBL,TOBCode : TOB ;
    RefUnique,Depot,sVA : String ;
    CptStock,CptVarStk,LibArt,Affaire : String ;
    CumStock,CumVarStk   : T_CodeCpta ;
    OkTenu : Boolean ;
    XD,XP : Double ;
    XDT : T_DetAnaL ;
    prefixe : string;
BEGIN

Result:=rcOk ;
sVA:=TOBPiece.GetValue('GP_VENTEACHAT') ;
for i:=0 to TOBPiece.Detail.Count-1 do
    BEGIN
    TOBL:=TOBPiece.Detail[i] ;
    if i = 0 then prefixe := GetPrefixeTable (TOBL);
    {Tests préalables}
    RefUnique:=TOBL.GetValue(prefixe+'_ARTICLE') ; if RefUnique='' then Continue ;
    OkTenu:=(TOBL.GetValue(prefixe+'_TENUESTOCK')='X') ; if Not OkTenu then Continue ;
    TOBA:=FindTOBArtRow(TOBPiece,TOBArticles,i+1) ; if TOBA=Nil then Continue ;
    LibArt:=Copy(TOBL.GetValue(prefixe+'_LIBELLE'),1,35) ;
    Affaire:=TOBL.GetValue(prefixe+'_AFFAIRE') ; Depot:=TOBL.GetValue(prefixe+'_DEPOT') ;
    {Recherche du paramétrage compta associé, sinon l'ajouter}
    TOBCode:=ChargeAjouteCompta(TOBCpta,TOBPiece,TOBL,TOBA,TOBTiers,Nil,Nil,False) ;
    if TOBCode<>Nil then
       BEGIN
       CptStock:=TOBCode.GetValue('GCP_CPTEGENESTOCK') ;
       CptVarStk:=TOBCode.GetValue('GCP_CPTEGENEVARSTK') ;
       END else
       BEGIN
       CptStock:='' ; CptVarStk:='' ;
       END ;
    if CptStock='' then CptStock:=GetParamSoc('SO_GCCPTESTOCK') ;
    if CptVarStk='' then CptVarStk:=GetParamSoc('SO_GCCPTEVARSTK') ;
    // Erreur sur le compte stock
    if ((CptStock='') or (CptVarStk='')) then BEGIN Result:=rcPar ; LastMsg:=11 ; Break ; END ;
    {Sommations par compte et familles de taxes identiques}
    RecupLesDCStock(TOBL,TOBArticles,(sVA='ACH'),XD,XP) ;
    {Stock}
    iTrouv:=-1 ;
    for k:=0 to TTStock.Count-1 do
        BEGIN
        CumStock:=T_CodeCpta(TTStock[k]) ;
        if ((CumStock.CptHT=CptStock) and (CumStock.Depot=Depot)) then BEGIN iTrouv:=k ; Break ; END ;
        END ;
    if iTrouv<0 then
        BEGIN
        CumStock:=T_CodeCpta.Create ; CumStock.CptHT:=CptStock ;
        CumStock.Depot:=Depot ; CumStock.LibU:=True ;
        CumStock.LibArt:=LibArt ; CumStock.Affaire:=Affaire ;
        END else
        BEGIN
        CumStock:=T_CodeCpta(TTStock[iTrouv]) ;
        END ;
    CumStock.XD:=CumStock.XD+XD ; CumStock.XP:=CumStock.XP+XP ;
    CumStock.Qte:=CumStock.Qte+TOBL.GetValue(prefixe+'_QTEFACT') ;
    if CumStock.LibArt<>LibArt then BEGIN CumStock.LibU:=False ; CumStock.LibArt:='' ; END ;
    CumulAnal(CumStock,TOBL,XD,XP) ;
    if iTrouv<0 then TTStock.Add(CumStock) ;
    {Variation Stock}
    iTrouv:=-1 ;
    for k:=0 to TTVarStk.Count-1 do
        BEGIN
        CumVarStk:=T_CodeCpta(TTVarStk[k]) ;
        if ((CumVarStk.CptHT=CptVarStk) and (CumVarStk.Depot=Depot)) then BEGIN iTrouv:=k ; Break ; END ;
        END ;
    if iTrouv<0 then
        BEGIN
        CumVarStk:=T_CodeCpta.Create ; CumVarStk.CptHT:=CptVarStk ; CumVarStk.LibU:=True ;
        CumVarStk.Depot:=Depot ; CumVarStk.LibArt:=LibArt ; CumVarStk.Affaire:=Affaire ;
        END else
        BEGIN
        CumVarStk:=T_CodeCpta(TTVarStk[iTrouv]) ;
        END ;
    CumVarStk.XD:=CumVarStk.XD+XD ; CumVarStk.XP:=CumVarStk.XP+XP ;
    CumVarStk.Qte:=CumVarStk.Qte+TOBL.GetValue(prefixe+'_QTEFACT') ;
    if CumVarStk.LibArt<>LibArt then BEGIN CumVarStk.LibU:=False ; CumVarStk.LibArt:='' ; END ;
    CumulAnal(CumVarStk,TOBL,XD,XP) ;
    if iTrouv<0 then TTVarStk.Add(CumVarStk) ;
    END ;
{Arrondissage des cumuls}
for i:=0 to TTStock.Count-1 do
    BEGIN
    CumStock:=T_CodeCpta(TTStock[i]) ;
    CumStock.XD:=Arrondi(CumStock.XD,NbDec) ; CumStock.XP:=Arrondi(CumStock.XP,V_PGI.OkDecV) ;
    for kk:=0 to CumStock.Anal.Count-1 do
        BEGIN
        XDT:=CumStock.Anal[kk] ;
        XDT.MD:=Arrondi(XDT.MD,NbDec) ;
        XDT.MP:=Arrondi(XDT.MP,V_PGI.OkDecV) ;
        if CumStock.XD<>0 then XDT.Qte1:=Arrondi(CumStock.Qte*XDT.MD/CumStock.XD,V_PGI.OkDecQ) ;
        XDT.TotQte1:=CumStock.Qte ;
        END ;
    END ;
for i:=0 to TTVarStk.Count-1 do
    BEGIN
    CumVarStk:=T_CodeCpta(TTVarStk[i]) ;
    CumVarStk.XD:=Arrondi(CumVarStk.XD,NbDec) ; CumVarStk.XP:=Arrondi(CumVarStk.XP,V_PGI.OkDecV) ;
    for kk:=0 to CumVarStk.Anal.Count-1 do
        BEGIN
        XDT:=CumVarStk.Anal[kk] ;
        XDT.MD:=Arrondi(XDT.MD,NbDec) ;
        XDT.MP:=Arrondi(XDT.MP,V_PGI.OkDecV) ;
        if CumVarStk.XD<>0 then XDT.Qte1:=Arrondi(CumVarStk.Qte*XDT.MD/CumVarStk.XD,V_PGI.OkDecQ) ;
        XDT.TotQte1:=CumVarStk.Qte ;
        END ;
    END ;
{Création des lignes stock}
for i:=0 to TTStock.Count-1 do
    BEGIN
    if Result=rcOk then Result:=CreerLignesEcrStock(TOBEcr,TOBPiece,TOBTiers,MM,i) ;
    if Result<>rcOk then Break ;
    END ;
{Création des lignes Variation Stock}
if Result=rcOk then for i:=0 to TTVarStk.Count-1 do
    BEGIN
    if Result=rcOk then Result:=CreerLignesEcrVarStk(TOBEcr,TOBPiece,TOBTiers,MM,i) ;
    if Result<>rcOk then Break ;
    END ;
END ;

Procedure MakeContrePStock ( TOBEcr : TOB ) ;
Var CptStock,CptVarStk,Gene : String ;
    TOBL                    : TOB ;
    i : integer ;
BEGIN
CptStock:='' ; CptVarStk:='' ;
for i:=0 to TOBEcr.Detail.Count-1 do
    BEGIN
    TOBL:=TOBEcr.Detail[i] ;
    Gene:=TOBL.GetValue('E_GENERAL') ;
    if ((CptStock='') and (TOBL.FieldExists('STOCK'))) then CptStock:=Gene ;
    if ((CptVarStk='') and (TOBL.FieldExists('VARSTK'))) then CptVarStk:=Gene ;
    if ((CptStock<>'') and (CptVarStk<>'')) then Break ;
    END ;
for i:=0 to TOBEcr.Detail.Count-1 do
    BEGIN
    TOBL:=TOBEcr.Detail[i] ;
    if TOBL.FieldExists('STOCK') then TOBL.PutValue('E_CONTREPARTIEGEN',CptVarStk) else
     if TOBL.FieldExists('VARSTK') then TOBL.PutValue('E_CONTREPARTIEGEN',CptStock) ;
    END ;
END ;

Procedure RenseigneClefStock ( TOBPiece : TOB ; Var MM : RMVT ) ;
BEGIN
MM.Etabl:=TOBPiece.GetValue('GP_ETABLISSEMENT') ;
MM.DateC:=TOBPiece.GetValue('GP_DATEPIECE') ; MM.Exo:=QuelExoDT(MM.DateC) ;
MM.CodeD:=V_PGI.DevisePivot ; MM.TauxD:=1 ;
MM.DateTaux:=TOBPiece.GetValue('GP_DATETAUXDEV') ;
MM.ModeSaisieJal:='-' ;
//MM.ModeOppose:=False ;
END ;

Procedure RenseigneTiersFact ( TOBPiece,TOBTiers,TOBTiersCpta : TOB ) ;
Var CodeTiers,NatureAuxi : String ;
    TiersCpta, Rib       : String ;
    Q                    : TQuery ;
BEGIN
CodeTiers:=TOBTiers.GetValue('T_TIERS') ; TiersCpta:=TOBPiece.GetValue('GP_TIERSFACTURE') ;
// JT eQualité 11032 (Gestion RIB)
if not TOBTiers.FieldExists('RIB') then TOBTiers.AddChampSup('RIB',False) ;
if TOBTiers.GetValue('RIB')='' then TOBTiers.PutValue('RIB',TOBPiece.GetValue('GP_RIB')) ;
// Fin JT
if ((CodeTiers=TiersCpta) or (TiersCpta='')) then Exit ;
NatureAuxi:=TOBTiers.GetValue('T_NATUREAUXI') ;
Q:=OpenSQL('SELECT * FROM TIERS WHERE T_TIERS="'+TiersCpta+'" AND T_NATUREAUXI="'+NatureAuxi+'"',True,-1, '', True) ;
if Not Q.EOF then TOBTiersCpta.SelectDB('',Q) ;
Ferme(Q) ;
// JT eQualité 11032 (Gestion RIB)
Q := OpenSQL('SELECT * FROM RIB WHERE R_AUXILIAIRE="' + TobTiersCpta.GetValue('T_AUXILIAIRE') +'" '+
             'AND R_PRINCIPAL="X"', True,-1, '', True);
if not Q.EOF then
begin
  Rib := EncodeRIB(Q.FindField('R_ETABBQ').AsString, Q.FindField('R_GUICHET').AsString,
         Q.FindField('R_NUMEROCOMPTE').AsString, Q.FindField('R_CLERIB').AsString,
         Q.FindField('R_DOMICILIATION').AsString);
  TobTiersCpta.PutValue('RIB', Rib);
end;
Ferme(Q);
// Fin JT
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 15/01/2002
Modifié le ... : 15/01/2002
Description .. : Passation stock comptable de la pièce Gescom
Mots clefs ... : LIAISON;COMPTABILITE;PONT;PASSATION;STOCK;
*****************************************************************}
Function PasseEnStock ( TOBPiece,TOBTiers,TOBArticles,TOBCpta : TOB ; NbDec : integer ; OldStk : RMVT ) : T_RetCpta ;
Var NatureG,NatCpta,JalCpta,RefCP : String ;
    TOBEcr      : TOB ;
    MM          : RMVT ;
    NumCpta     : Longint ;
    ii          : integer ;
    QQ          : TQuery ;
BEGIN
Result:=rcOk ; LastMsg:=-1 ;
NatureG:=TOBPiece.GetValue('GP_NATUREPIECEG') ; if Not isComptaStock(NatureG) then Exit ;
{TOB Ecriture Générique}
TOBEcr:=TOB.Create('STOCK COMPTABLE',Nil,-1) ;
// Création du RECORD clef comptable
FillChar(MM,Sizeof(MM),#0) ; MM.Simul:='N' ;
RenseigneClefStock(TOBPiece,MM) ;
{Journal}
JalCpta:=GetParamSoc('SO_GCJALSTOCK') ;
// Erreur sur Journal
if JalCpta='' then BEGIN TOBEcr.Free ; Result:=rcPar ; LastMsg:=7 ; Exit ; END ;
MM.Jal:=JalCpta ;
{Nature}
NatCpta:='OD' ; MM.Nature:=NatCpta ;
{Valide}
MM.Valide:=True ;
{Numéro}
if ((MM.Jal=OldStk.Jal) and (MM.Simul=OldStk.Simul) and (MM.Nature=OldStk.Nature)) then
   BEGIN
   NumCpta:=OldStk.Num ; MM.DateC:=OldStk.DateC ;
   END else
   BEGIN
   NumCpta:=GetNewNumJal(JalCpta,(MM.Simul='N'),MM.DateC) ;
   END ;
{Type de journal}
MM.ModeSaisieJal:='-' ;
QQ:=OpenSQL('SELECT J_MODESAISIE FROM JOURNAL WHERE J_JOURNAL="'+MM.Jal+'"',True,-1, '', True) ;
if Not QQ.EOF then MM.ModeSaisieJal:=QQ.Fields[0].AsString ;
Ferme(QQ) ;
// Erreur sur le numéro
if NumCpta<=0 then BEGIN TOBEcr.Free ; Result:=rcPar ; LastMsg:=9 ; Exit ; END ;
MM.Num:=NumCpta ;
{Préparer analytiques}
TTStock:=TList.Create ; TTVarStk:=TList.Create ;
{Lignes Stock}
if Result=rcOk then Result:=CreerLignesStock(TOBEcr,TOBPiece,TOBTiers,TOBArticles,TOBCpta,MM,NbDec) ;
if ((Result=rcOk) and (TOBEcr.Detail.Count>=2)) then
   BEGIN
   {Ultimes ajustements}
   AnalyseLesNegatifs(TOBEcr) ;
   MakeContrePStock(TOBEcr) ;
   {Analytiques}
   MakeAnalGC(TOBEcr,TOBPiece,MM,NbDec) ;
   {Retour inverse}
   RefCP:=EncodeRefGCComptable(TOBEcr.Detail[0]) ; TOBPiece.PutValue('GP_REFCPTASTOCK',RefCP) ;
   {MAJ des soldes comptables}
   MajSoldesEcritureTOB(TOBEcr,True) ;
   if V_PGI.IoError<>oeOk then BEGIN Result:=rcMaj ; LastMsg:=10 ; END ;
   if Result=rcOk then
      BEGIN
      {Insertion}
      if Not TOBEcr.InsertDBByNivel(False) then V_PGI.IoError:=oeUnknown ;
      END ;
   END ;
TOBEcr.Free ;
for ii:=0 to TTStock.Count-1 do
    BEGIN
    VideListe(T_CodeCpta(TTStock[ii]).Anal) ;
    T_CodeCpta(TTStock[ii]).Anal.Free ; T_CodeCpta(TTStock[ii]).Anal:=Nil ;
    END ;
VideListe(TTStock) ; TTStock.Free ; TTStock:=Nil ;
for ii:=0 to TTVarStk.Count-1 do
    BEGIN
    VideListe(T_CodeCpta(TTVarStk[ii]).Anal) ;
    T_CodeCpta(TTVarStk[ii]).Anal.Free ; T_CodeCpta(TTVarStk[ii]).Anal:=Nil ;
    END ;
VideListe(TTVarStk) ; TTVarStk.Free ; TTVarStk:=Nil ;
END ;

procedure CalcBasesTaxe (TOBPiece,TOBBasesL,TOBL,TOBTiers,TOBAFFInterv : TOB; EnHt: boolean; DEV : rdevise);
var prefixe : string;
    iMontant : integer;
    TOBTaxesL : TOB;
//    NatureTravail : double;
    Fournisseur : string;
begin
  TOBTaxesL := TOB.Create ('TAXES LIGNE',nil,-1);
  TOBTaxesL.AddChampSupValeur ('UN CHAMP',''); // pour le tob debug
  TRY
    prefixe := GetPrefixeTable (TOBL);
    (*
    begin
    	ZeroLigneMontant (TOBL);
      exit;
    end;
    *)
    if EnHT then Imontant:=TOBL.GetNumChamp (prefixe+'_TOTALHTDEV')
            else Imontant:=TOBL.GetNumChamp (prefixe+'_TOTALTTCDEV') ;

    CalculTaxesLigne(TOBL,TOBPiece,TOBTiers,TOBTaxesL,DEV,DEV.Decimale,TOBL.GetValeur(Imontant),EnHt) ;
    ChangeParentLignesBases (TOBBasesL,TOBtaxesL,Dev.decimale);
  FINALLY
    TOBtaxesL.Free;
  END;
end;


procedure RecupLigneOuvPlats (TOBpiece,TOBL,TOBOuvragePMere,TOBouvragesP : TOB);
var TOBPlatM,TOBPlat : TOB;
begin
  TOBPlatM := FindOuvragesPlat (TOBL,TOBOuvragePMere);
  if TOBPlatM <> nil then
  begin
    TOBPlat := AddMereLignePlat (TOBOuvragesP,TOBL.GetValue('GL_NUMORDRE'));
    TOBPLat.Dupliquer(TOBPlatM,True,true);
  end;
end;

procedure RecalculeLigneFromDet (TOBPiece,TOBArticles,TOBBasesL,TOBL,TOBOuvrages,TOBOuvragesP,TOBTiers,TOBAFFInterv,TOBPieceInterv: TOb;EnHt : boolean;DEV : Rdevise);
var Indice : integer;
    TOBPlat,TOBPP,TOBA,TOBTaxesL : TOB;
    refUnique,Fournisseur : string;
    NatureTravail : double;
begin

	TOBTaxesL := TOB.Create ('LES TAXES LIGNES',nil,-1);
  TRY
    TOBPlat := FindOuvragesPlat (TOBL,TOBOuvragesP);
    if TOBPlat = nil then
    begin
      TOBPlat := AddMereLignePlat (TOBOuvragesP,TOBL.GetValue('GL_NUMORDRE'));
    end else
    begin
      TOBPlat.ClearDetail;
    end;
    NatureTravail := 0;

    if TOBL.FieldExists('BLF_NATURETRAVAIL') then NatureTravail := Valeur(TOBL.GetValue('BLF_NATURETRAVAIL'));
    if (NatureTravail = 0) then
    begin
      NatureTravail := Valeur(TOBL.GetValue('GLC_NATURETRAVAIL'));
    end;
    //
    GetOuvragePlat (TOBpiece,TOBL,TOBOuvrages,TOBPlat,TOBTiers,DEV,true,(NatureTravail>10));
    //
    ZeroLigneMontant (TOBL);
    if (TOBL.FieldExists('BLF_NATUREPIECEG')) and (TOBL.GetString('BLF_NATUREPIECEG')<>'') then
    begin
      TOBL.PutValue('BLF_MTSITUATION',0);
      TOBL.PutValue('BLF_QTESITUATION',0);
      TOBL.PutValue('BLF_POURCENTAVANC',0);
    end;
    Indice := 0;
    repeat
      TOBTaxesL.ClearDetail;
      TOBPP := TOBplat.detail[Indice];
      refUnique := TOBPP.GetValue('BOP_ARTICLE');
      TOBA := TOBarticles.findFirst(['GA_ARTICLE'],[refUnique],true);
      if TOBA=Nil then Continue ;
      NatureTravail := Valeur(TOBPP.GetValue('BOP_NATURETRAVAIL'));
      Fournisseur := TOBPP.GetValue('BOP_FOURNISSEUR');
      if (NatureTravail>0) and (NatureTravail < 10) and
         (Fournisseur <> '') and (not ComptabiliseTrait (fournisseur,TOBAffInterv,TOBPieceInterv)) then
      begin
        TOBPP.Free;
        continue;
      end;
      ZeroLigneMontant (TOBPP);
      if EnHT then
      BEGIN
        CalculeLigneHT(TOBPP,TOBTaxesL,TOBPiece,DEV, DEV.Decimale ,False,TOBTiers) ;
        ChangeParentLignesBases (TOBBasesL,TOBtaxesL,DEV.Decimale);
        //
        SommeLignePlat(TOBPP,TOBL,EnHT) ;
      END else
      BEGIN
        CalculeLigneTTC(TOBPP,TOBTaxesL,TOBPiece,DEV, DEV.Decimale,False,TOBTiers) ;
        ChangeParentLignesBases (TOBBasesL,TOBtaxesL,DEV.Decimale);
        //
        SommeLignePlat(TOBPP,TOBL,EnHT) ;
      END ;
      if TOBL.FieldExists('BLF_NATUREPIECEG') and (TOBL.GetString('BLF_NATUREPIECEG')<>'') then
      begin
      	TOBL.putValue('BLF_MTSITUATION',TOBL.GetValue('BLF_MTSITUATION')+TOBPP.getValue('BLF_MTSITUATION'));
      end;
      Inc(Indice);
    until Indice  >= TOBplat.detail.Count;
  FINALLY
    TOBTaxesL.Free;
  end;
end;

procedure CalculBaseRGCalc (TOBPieceinterv,TOBAFFInterv,TOBTiers,TOBPiece,TOBPieceTrait,TOBBasesL,TOBPIeceRG,TOBBasesRg,TOBS: TOB ; TOBPorcs: TOB; R_Valeurs : R_CPercent;DEV:Rdevise;GenerationFac:boolean=false);
var
   TOBL,TOBBases,TOBPIECELOC,TOBRG,TOBLLOC,TOBPorcsLoc,TOBBasesLoc,TOBBL: TOB;
   Indice,INdL,IndiceRg : Integer;
   IPuHt,IpuHtdev,IpuTTC,iputtcdev,iMontHt,imontHtDev,iMontTTC : integer;
   iMontTTCdev: integer;
   bid1,bid2,bid3,bid4 : double;
   Applicretenue : boolean;
   PortHt,PortTTC : double;
   NumOrdre : integer;
   TOBB : TOB;
   TOBTaxes : TOB;
   NatureTravail : double;
begin
  if TOBPieceRG = nil then exit;
  if IndiceRG = 0 then exit;
  TOBRG := TOBPieceRG.findfirst (['INDICERG'],[IndiceRG],true);
  if TOBRG.GetBoolean('PRG_MTMANUEL') then exit;
  //
  ApplicRetenue:=(GetInfoParPiece(TOBPIECE.GetValue('GP_NATUREPIECEG'),'GPP_APPLICRG')='X');
  TOBPIECELOC:=TOB.create ('PIECE',nil,-1);
  TOBPorcsLoc := TOB.Create ('LES PORTS LOC',nil,-1);
  TOBBasesLoc := TOB.Create ('LES BASES LIGNES',nil,-1);
  TOBLLOC:=TOB.Create('LIGNE',nil,-1);
  TOBBases := TOB.create('BASES PIED(LIGNES)',nil,-1);
  //
  AddLesSupEntete (TOBpieceLoc);
  TOBPieceLOC.dupliquer (TOBPIECE,false,true);
  ZeroFacture (TOBpieceloc);
  //
  AddLesSupLigne (TOBLLOc,false);
  //
  IndiceRg := TOBS.GetValue('INDICERG');

  TOBPorcsLoc.Dupliquer (TOBPorcs,true,true);
  //
  TOBBasesLoc.Dupliquer (TOBBasesL,true,true);
  // nettoyage avant création
  VidePieceRGLigne (TOBPieceRG,IndiceRG);
  // --
  ZeroLignePourcent(TOBS) ;
  ZeroLignePourcent(TOBLLOC) ;
  IndL := 0;
  IPuHt := 0;
  IPuHtdev := 0;
  IPuTTC := 0;
  IPuttcdev := 0;
  IMOntHt := 0;
  IMOntHtDev := 0;
  IMOntttc := 0;
  IMOntttcDev := 0;
  for Indice := R_Valeurs.Depart to R_Valeurs.Fin do
  begin
    TOBL := GetTOBLigne(TOBPiece,Indice+1 ) ; if TOBL = nil then break;
    // --
    NatureTravail := valeur(TOBL.GetValue('BLF_NATURETRAVAIL'));
    if (NatureTravail >0) and (NatureTravail <10) and (TOBL.GetValue('BLF_FOURNISSEUR')<>'') and
       (not ComptabiliseTrait (TOBL.GetValue('BLF_FOURNISSEUR'),TOBAFFInterv,TOBPieceinterv)) then continue;
    // --
    NumOrdre := TOBL.getValue('GL_NUMORDRE');
    if Indice = R_Valeurs.Depart then
    begin
      IndL := TOBL.GetNumChamp ('GL_TYPELIGNE');
      // PU HT
      IPuHtDev := TOBL.GetNumChamp ('GL_PUHTDEV');
      IPuHt := TOBL.GetNumChamp ('GL_PUHT');
      // PU TTC
      IPuTTCDev := TOBL.GetNumChamp ('GL_PUTTCDEV');
      IPuTTC := TOBL.GetNumChamp ('GL_PUTTC');
      // Montant Ht
      IMontHt := TOBL.GetNumChamp ('GL_TOTALHT');
      IMontHtDev := TOBL.GetNumChamp ('GL_TOTALHTDEV');
      // Montant TTC
      IMontTTC := TOBL.GetNumChamp ('GL_TOTALTTC');
      IMontTTCDev := TOBL.GetNumChamp ('GL_TOTALTTCDEV');
      // -------
    end;
    if TOBL.GetValeur(INDL) = 'EPO' then break;
    // VARIANTE
    if (isVariante(TOBL)) or (IsLigneFromCentralis(TOBL)) then continue;
    // --
    if (TOBL.getvaleur(INDL) = 'RG' ) then continue;
    if (TOBL.getvalue('GL_ARTICLE') = '' ) then continue;
    SommeLignePiedHT (TOBL,TOBpieceLoc);
    TOBBL := FindLignesbases (numordre,TOBBasesL);
    if TOBBL <> nil then CumuleLesBases (TOBBases, TOBBL,DEV.Decimale);
  end;
  // --
  RecalculPiedPort (tamodif,TOBpieceLoc,TOBporcsloc,TOBBases);
  AddlesPorts (TOBPIECELOC,TOBPorcsLoc,TOBBases,TOBBasesLoc,TOBtiers,DEv,TaCreat);

  TOBRG := TOBPieceRG.findfirst (['INDICERG'],[IndiceRG],true);
  if TOBRG <> nil then
  begin
    CalculeRGSimple (TOBPorcsLoc,TOBPIECELOC,TOBRG,TOBBases,TOBBASESRG,DEV);
    RGVersLigne (TOBPIECELOC,TOBRG,TOBBasesRG,TOBS,nil,Applicretenue,GenerationFac);
  end;
  // nettoyage avant départ
  TOBPIECELOC.free;
  TOBPorcsLoc.free;
  TOBBasesLoc.free;
  TOBLLOC.free;
  TOBBases.free;
end;

procedure RecalculeRGCpta (TOBPieceinterv,TOBAFFInterv,TOBTiers,TOBPieceTrait,TOBPiece,TOBBasesL,TOBPieceRG,TOBBasesRG,TOBporcs: TOB; DEV : Rdevise);
var TOBL : TOB;
I : Integer;
T_RG :R_CPercent;
begin
  TOBL := TOBPIECE.findfirst(['GL_TYPELIGNE'],['RG'],true);
  WHILE TOBL <> nil do
  BEGIN
    // Calcul Retenue de garantie dans le cas de document regroup
    I:= TOBL.GetIndex;
    T_RG.Traite := I;
    T_RG.Depart := RecupDebutRG (TOBPiece,I - 1);
    if I > 0 then T_RG.Fin := I - 1
             else T_RG.Fin := 0;
    CalculBaseRGCalc (TOBPieceinterv,TOBAFFInterv,TOBTiers,TOBPiece,TOBPieceTrait,TOBBasesL,TOBPieceRg,TOBBasesRg,TOBL,tobporcs,T_RG,DEV,True);
    TOBL := TOBPIECE.findnext(['GL_TYPELIGNE'],['RG'],true);
  END;
end;

procedure CalculeLaPieceCpta (TOBPiece,TOBAcomptes,TOBEches,TOBArticles, TOBBases, TOBBasesST, TOBOuvrages,TOBOuvragesP,TOBtiers,TOBAFFInterv,TOBPorcs,TOBPieceRG,TOBBasesRG,TOBPieceTrait,TOBPieceinterv,TOBOuvPlatMere : TOB; DEV : Rdevise);
var EnHt : boolean;
    TOBBasesL,TOBL,TOBA,TOBB : TOB;
    NumOrdre,NBDEc : integer;
    RefUnique,Fournisseur : string;
    i : integer;
    NatureTravail : double;
begin
  AssigneDocumentTva2014 (TOBPiece,TOBPieceinterv);
  EnHt := (TOBPiece.getvalue('GP_FACTUREHT')='X');
  NBDec := DEV.decimale;
  TOBBases.ClearDetail;
  TOBbasesL := TOB.Create('LES LIGNES BASES',nil,-1);
  //
  i := 0;
  repeat
    TOBBasesL.clearDetail;
    {Ratio := 1; }
    TOBL:=TOBPiece.Detail[i] ;
    NatureTravail := 0;
    Fournisseur := '';

    if (TOBL.FieldExists('BLF_NATURETRAVAIL')) then
    begin
    	NatureTravail := Valeur(TOBL.GetValue('BLF_NATURETRAVAIL'));
    end;

    if (NatureTravail = 0) and (TOBL.GetString('GLC_NATURETRAVAIL')<> '') then
    begin
    	NatureTravail := Valeur(TOBL.GetValue('GLC_NATURETRAVAIL'));
    end;

    if TOBL.fieldExists('BLF_FOURNISSEUR') then
    begin
    	Fournisseur := TOBL.GetValue('BLF_FOURNISSEUR')
    end;
    if (Fournisseur = '') and (NatureTravail <> 0) then
    begin
      Fournisseur := TOBL.GetValue('GL_FOURNISSEUR');
    end;
    if (VH_GC.BTCODESPECIF = '001') and (IsLigneSoustraite (TOBL)) then
    begin
       TOBL.Free;
       continue;
    end;

    //
    if (NatureTravail > 0) and (NatureTravail<10) and (Fournisseur<>'') and
       (not ComptabiliseTrait (Fournisseur,TOBAFFInterv,TOBPieceInterv)) then
    begin
       TOBL.Free;
       continue;
    end;
    NumOrdre := TOBL.getValue('GL_NUMORDRE');
    {Tests préalables}
    //CONTREM ?????
    RefUnique:=TOBL.GetValue('GL_ARTICLE') ; if RefUnique='' then begin inc(i); Continue ; end;
    TOBA:=FindTOBArtRow(TOBPiece,TOBArticles,i+1) ; if TOBA=Nil then begin inc(i); Continue ; end;
    if IsLigneFromCentralis(TOBL) then begin inc(i); Continue ; end;
    if IsVariante(TOBL) then begin inc(i); Continue ; end;
    if ( IsOuvrageComptabiliseDetail(TOBL) and ( TOBL.GetValue('GL_INDICENOMEN')>0 ) ) then
    begin
      RecalculeLigneFromDet (TOBPiece,TOBArticles,TOBBasesL,TOBL,TOBOuvrages, TOBOuvragesP,TOBTiers,TOBAFFInterv,TOBPieceinterv,EnHt,DEV);
    end else
    begin
			RecupLigneOuvPlats (TOBpiece,TOBL,TOBOuvPlatMere,TOBouvragesP);
      CalcBasesTaxe (TOBPiece,TOBBasesL,TOBL,TOBTiers,TOBAFFInterv,EnHt,DEV);
    end;
    // ---
    TOBB := FindLignesbases (NumOrdre,TOBBasesL);
    if TOBB <> nil then
    begin
      ReajusteLigneViaBase (TOBL,TOBB,EnHt,DEV);
    end;
    if EnHt then
    begin
      SommeLignePiedHT(TOBL,TOBPiece) ;
    end else
    begin
      SommeLignePiedTTC(TOBL,TOBPiece) ;
    end;
    if TOBB <> nil then CumuleLesBases(TOBBases,TOBB,NbDec);
    Inc(I);
  until i >= TOBPiece.detail.count ;
  TOBBasesL.clearDetail;
  //
  RecalculPiedPort (TaModif,TOBpiece,TOBporcs,TOBBases);
  AddLesports(TOBpiece,TOBPorcs,TOBBases,TOBBasesL,TOBTiers,DEV,TaModif);
  EquilibrageBases(TOBpiece,TOBBases,EnHT,DEV);
  RecalculeRGCpta (TOBPieceinterv,TOBAFFInterv,TOBTiers,TOBPieceTrait,TOBPiece,TOBBasesL,TOBPieceRG,TOBBasesRG,TOBporcs,DEV);
  if not RGMultiple(TOBpiece) then
  begin
  	RecalculeRG(TOBPORCS,TOBPIECE, TOBPIECERG, TOBBASES, TOBBASESRG,TOBPieceTrait, DEV);
  end;
  //
  GereEcheancesGC (TOBpiece,TOBTiers,TOBEches,TOBAcomptes,TOBPieceRG,TOBPieceTrait,TOBPorcs,taModif,DEV,false);
  RePackBase (TObbases,TOBBasesST,DEV); // du fait que les bases soient scindés via co ou sous traitants
  TOBbasesL.free;
  InitDocumentTva2014;
end;

procedure EnleveLesNonConcernesRG (TOBPieceRG : TOB);
var Indice : integer;
begin
  if TOBPieceRG.detail.count = 0 then exit;
  Indice := 0;
	repeat
		if TOBPieceRG.detail[Indice].GEtString('PRG_FOURN') <> '' then TOBPieceRG.detail[Indice].free
    																													else Inc(Indice);
  Until Indice >= TOBPieceRG.detail.count;
end;


Function PassationComptable ( TOBPiece,TOBOuvrages,TOBOuvragesP,TOBBases,TOBBasesL,TOBEches,TOBPieceTrait,TOBAffInterv,TOBTiers,TOBArticles,TOBCpta,TOBAcomptes,TOBPorcs,TOBPieceRG,TOBBasesRG,TOBAnaP,TOBAnaS,TOBPieceInterv,TOBVTECOLL  : TOB ; DEV : Rdevise ; OldEcr,OldStk : RMVT ; bParle : Boolean; NewPiece : Boolean = True ) : boolean ;

  procedure AddPorcBase (TOBBase,TOBPorcs : TOB; Sens : string='+');
  var TOBB : TOB;
  begin
    TOBB := TOBBase.FindFirst(['GPB_CATEGORIETAXE','GPB_FAMILLETAXE'],['TX1',TOBPorcs.GetString('GPT_FAMILLETAXE1')],true);
    if TOBB = nil then
    begin
      TOBB := TOB.Create ('PIEDBASE',TOBBase,-1);
      TOBB.AddChampSupValeur('XXXXX','X');
      TOBB.SetString('GPB_CATEGORIETAXE','TX1');
      TOBB.SetString('GPB_FAMILLETAXE',TOBPorcs.GetString('GPT_FAMILLETAXE1'));
      TOBB.SetDouble('GPB_TAUXTAXE',Arrondi(((TOBPorcs.GetDouble('GPT_BASETTCDEV')/TOBPorcs.GetDouble('GPT_BASEHTDEV'))-1)*100,0));
      TOBB.SetString('GPB_DEVISE',TOBPorcs.GetString('GPT_DEVISE'));
    end;
    if Sens = '+' then
    begin
      TOBB.SetDouble('GPB_BASETAXE',Arrondi(TOBB.GetDouble('GPB_BASETAXE')+TOBPorcs.GetDouble('GPT_TOTALHT'),DEV.Decimale));
      TOBB.SetDouble('GPB_BASETAXETTC',Arrondi(TOBB.GetDouble('GPB_BASETAXETTC')+TOBPorcs.GetDouble('GPT_TOTALTTC'),DEV.Decimale));
      TOBB.SetDouble('GPB_VALEURTAXE',Arrondi(TOBB.GetDouble('GPB_VALEURTAXE')+TOBPorcs.GetDouble('GPT_TOTALTAXE1'),DEV.Decimale));
      TOBB.SetDouble('GPB_BASEDEV',Arrondi(TOBB.GetDouble('GPB_BASEDEV')+TOBPorcs.GetDouble('GPT_TOTALHTDEV'),DEV.Decimale));
      TOBB.SetDouble('GPB_BASETTCDEV',Arrondi(TOBB.GetDouble('GPB_BASETTCDEV')+TOBPorcs.GetDouble('GPT_TOTALTTCDEV'),DEV.Decimale));
      TOBB.SetDouble('GPB_VALEURDEV',Arrondi(TOBB.GetDouble('GPB_VALEURDEV')+TOBPorcs.GetDouble('GPT_TOTALTAXEDEV1'),DEV.Decimale));
    end else
    begin
      TOBB.SetDouble('GPB_BASETAXE',Arrondi(TOBB.GetDouble('GPB_BASETAXE')-TOBPorcs.GetDouble('GPT_TOTALHT'),DEV.Decimale));
      TOBB.SetDouble('GPB_BASETAXETTC',Arrondi(TOBB.GetDouble('GPB_BASETAXETTC')-TOBPorcs.GetDouble('GPT_TOTALTTC'),DEV.Decimale));
      TOBB.SetDouble('GPB_VALEURTAXE',Arrondi(TOBB.GetDouble('GPB_VALEURTAXE')-TOBPorcs.GetDouble('GPT_TOTALTAXE1'),DEV.Decimale));
      TOBB.SetDouble('GPB_BASEDEV',Arrondi(TOBB.GetDouble('GPB_BASEDEV')-TOBPorcs.GetDouble('GPT_TOTALHTDEV'),DEV.Decimale));
      TOBB.SetDouble('GPB_BASETTCDEV',Arrondi(TOBB.GetDouble('GPB_BASETTCDEV')-TOBPorcs.GetDouble('GPT_TOTALTTCDEV'),DEV.Decimale));
      TOBB.SetDouble('GPB_VALEURDEV',Arrondi(TOBB.GetDouble('GPB_VALEURDEV')-TOBPorcs.GetDouble('GPT_TOTALTAXEDEV1'),DEV.Decimale));
    end;
  end;

Var Res : T_RetCpta ;
    Tex : String ;
    TOBTiersCpta : TOB ;
    TOBEchesCpta,TOBBasesCpta,TOBBasesCptaST,TOBPieceCpta,TOBPieceRGCpta,TOBBaseRGCpta,TOBPorcsCpta,TOBAcomptesCpta,TOBOuvragesPCpta,TOBPieceTraitCpta,TOBBasesCharges : TOB;
    I : Integer;
BEGIN
  TOBBasesCharges := TOB.Create ('LES BASES DES CHARGES',nil,-1);
//AC 18/08/03 NV GESTION COMPTA DIFF
//Si compta différée pour pièce et acompte, ne gère rien içi
Result := True;
if TOBPiece.GetBoolean('GP_HORSCOMPTA') then Exit;
if not IsComptaPce(TOBPiece.GetValue('GP_NATUREPIECEG')) then exit;
// Fin AC
// Modif BTP
LaDEV.Code:=TOBPiece.GetValue('GP_DEVISE') ; GetInfosDevise(LaDEV) ;
RGComptabilise := GetParamsoc('SO_GCVENTILRG') ;
// --
if VH_GC.GCIfDefCEGID then
begin
  Result:=True ;
  if ((TOBPiece.GetValue('GP_TOTALHT')=0) and (TOBPiece.GetValue('GP_TOTALTTC')=0)) then Exit ;
end;

//
if (GetParamSocSecur('SO_BTCPTAPAIEDIRECT',false)) and IsComptabilisationInterv (TOBPieceTrait) then // comptabilisation co traitance et/ou sous traitance
begin
  TOBPorcsCpta := TOB.Create('BASES RG ENTREP',nil,-1);
  TOBEchesCpta := TOB.create ('LES ECHEANCES ENTREPRISE',nil,-1);
  TOBBasesCpta := TOB.Create ('LES BASES ENTREPRISE',nil,-1);
  TOBBasesCptaST := TOB.Create ('LES BASES ST AUTOLIQUIDE',nil,-1);
  TOBPieceRGCpta := TOB.Create('PIECE RG ENTREP',nil,-1);
  TOBPieceRGCpta.Dupliquer(TOBPieceRG,true,true);
  EnleveLesNonConcernesRG (TOBPieceRGCpta);
  TOBPieceCpta := TOB.Create ('PIECE',nil,-1); // la piece de l'entreprise
  TOBBaseRGCpta := TOB.Create('BASES RG ENTREP',nil,-1);
  TOBAcomptesCpta := TOB.Create('ACOMPTES ENTREP',nil,-1);
  TOBOuvragesPCpta := TOB.Create ('LES OUVRAGES PLAT CPTA',nil,-1);
	TOBPieceTraitCpta := TOB.Create ('LES INTERV',nil,-1);
  //
  AddLesSupEntete (TOBpieceCpta);
  TOBPieceCpta.Dupliquer (TOBpiece,true,true);
  //
  TOBPorcsCpta.Dupliquer(TOBPorcs,true,true);
  TOBAcomptesCpta.Dupliquer(TOBAcomptes,true,true);
  //
  ZeroFacture (TOBpieceCpta);
  ZeroMontantPorts (TOBPorcsCpta);
  //
//  RecupReglementsEnt (TOBAcomptesCpta,TOBAcomptes);
  // Redefini une piece et des montants pour l'entreprise mandataire
  CalculeLaPieceCpta (TOBpieceCpta,TOBAcomptesCpta,TOBEchesCpta,TOBArticles,TOBBasesCpta,TOBBasesCptaST,TOBOuvrages,TOBOuvragesPCpta,TOBTiers,TOBAFFinterv,TOBPorcsCpta,TOBPieceRGCpta,TOBBaseRGCpta,TOBPieceTrait,TOBPieceInterv,TOBOuvragesP, DEV);
  //
end else
begin
	TOBPorcsCpta := TOBPorcs;
  TOBEchesCpta := TOBEches;
  TOBBasesCpta := TOB.Create ('LES BASES ENTREPRISE',nil,-1);
  TOBBasesCpta.Dupliquer(TOBBases,True,true);
  TOBPieceRGCpta := TOBpieceRg;
  TOBpieceCpta := TOBpiece;
  TOBBaseRGCpta := TOBBasesRG;
  TOBAcomptesCpta := TOBAcomptes;
  TOBOuvragesPCpta := TOBOuvragesP;
  TOBPieceTraitCpta := TOBPieceTrait;
  TOBBasesCptaST := nil;
	RePackBase (TOBBasesCpta,TOBBasesCptaST,DEV);
end;
//
for I := 0 TO TOBPorcsCpta.detail.count -1 do
begin
  if TOBPorcsCpta.detail[I].GetString('GPT_RETENUEDIVERSE')='X' then continue;
  if (Pos(TOBPorcsCpta.detail[I].GetString('GPT_TYPEPORT'),'PT;MIC;MTC;')>0) and (TOBPorcsCpta.detail[I].GetDouble('GPT_TOTALTTCDEV')<0) then
  begin
    AddPorcBase(TOBBasesCpta,TOBPorcsCpta.detail[I],'-');
    AddPorcBase(TOBBasesCharges,TOBPorcsCpta.detail[I]);
  end;
end;
//
Result:=False ; DistinctAffaire:=GetParamsoc('SO_GCDISTINCTAFFAIRE') ;
TOBTiersCpta:=TOB.Create('TIERS',Nil,-1) ; TOBTiersCpta.Dupliquer(TOBTiers,True,True) ;
RenseigneTiersFact(TOBPiece,TOBTiers,TOBTiersCpta) ;
Res:=PasseEnCompta(TOBPieceCpta,TOBOuvrages,TOBOuvragesPCpta,TOBBasesCpta,TOBBasesCptaST,TOBBasesCharges,TOBEchesCpta,TOBPieceTrait,TOBAFFInterv,TOBTiersCpta,TOBArticles,TOBCpta,TOBAcomptesCpta ,TOBPorcsCpta,TOBPieceRGCpta,TOBBaseRGCpta ,TOBanaP,TOBAnaS,TOBPieceinterv,TOBVTECOLL,DEV.Decimale,OldEcr,DEV,NewPiece) ;
if Res=rcOk then Res:=PasseEnStock(TOBPiece,TOBTiers,TOBArticles,TOBCpta,DEV.Decimale,OldStk) ;
TOBBasesCpta.free;
if (GetParamSocSecur('SO_BTCPTAPAIEDIRECT',false)) and (IsComptabilisationInterv (TOBPieceTrait))  then  // gestion co traitance et sous traitance
begin
	TOBpiece.putValue('GP_REFCOMPTABLE',TOBPieceCpta.GetValue('GP_REFCOMPTABLE'));
	TOBPorcsCpta.free;
  TOBEchesCpta.free;
  TOBPieceRGCpta.free;
  TOBpieceCpta.free;
  TOBBaseRGCpta.free;
  TOBAcomptesCpta.free;
  TOBPieceTraitCpta.free;
  if TOBBasesCptaST <> nil then TOBBasesCptaST.free;
end;
if LastMsg>0 then Tex:=TexteMessage[LastMsg] else Tex:='' ;
Case Res of
   rcOk  : Result:=True ;
   else result := false;
   END ;
TOBTiersCpta.Free ;
TOBBasesCharges.free;
END ;

{==============================================================================}
{============================ GESTION DES ACOMPTES ============================}
{==============================================================================}
Type T_CPAcompte = Class
                   LeAcc : T_GCAcompte ;
                   TOBPiece,TOBTiers,NewACC,TOBFournisseur : TOB ;
                   // Modif BTP
                   IsModif : boolean;
                   // --
                   Procedure ValideAcompte ;
                   Procedure ValideReglSousTrait ;
                   Constructor Create ;
                   Destructor Destroy ;   override;
                   private
                   procedure MajMontantAcc;
                   End ;

Constructor T_CPAcompte.Create ;
BEGIN
Inherited Create ;
NewACC:=Nil ;
END ;

Destructor T_CPAcompte.Destroy ;
BEGIN
if NewACC<>Nil then NewAcc.Free ;
Inherited Destroy ;
END ;

Function ACC_FabriqueMM ( TOBPiece : TOB ; Var LeAcc : T_GCAcompte; EcritureSSTrait : Boolean=false ) : RMVT ;
Var MM : RMVT ;
    NatureG,QualifP : String ;
BEGIN
  FillChar(MM,Sizeof(MM),#0) ;
  {Journal}
  MM.Jal:=LeAcc.JalRegle ;
  if MM.Jal='' then BEGIN LastMsg:=16 ; Exit ; END ;
  {Type Ecr}
  NatureG:=TOBPiece.GetValue('GP_NATUREPIECEG') ;
  LeAcc.LienCpta:=GetInfoParPiece(NatureG,'GPP_TYPEPASSACC') ;
  QualifP:='N' ;
  if VH_GC.GCIfDefCEGID then
  QualifP:='S' ;
  if LeAcc.LienCpta='DIF' then MM.Simul:='S' else MM.Simul:=QualifP ;
  if MM.Simul='N' then MM.Valide:=True else MM.Valide:=False ;
  {Numéro}
  MM.Num:=GetNewNumJal(MM.Jal,(MM.Simul='N'),MM.DateC) ;
  if MM.Num<=0 then BEGIN LastMsg:=9 ; Exit ; END ;
  {TOBPiece}
  //
  if LeAcc.DateEcr <> '' then
  	MM.DateC:= StrToDate(LeAcc.DateEcr)
  else
  	MM.DateC:=TOBPiece.GetValue('GP_DATEPIECE') ;
  // -- Date Echeance ---------
  if LeAcc.DateEche   <> '' then
  	MM.DateE:= StrToDate(LeAcc.DateEche)
  else
  	MM.DateE:=TOBPiece.GetValue('GP_DATEPIECE') ;
  // --------------------------
  MM.CodeD:=TOBPiece.GetValue('GP_DEVISE') ; MM.DateTaux:=TOBPiece.GetValue('GP_DATETAUXDEV') ;
  MM.TauxD:=TOBPiece.GetValue('GP_TAUXDEV') ;
  MM.Etabl:=TOBPiece.GetValue('GP_ETABLISSEMENT') ;
  //MM.ModeOppose:=(TOBPiece.GetValue('GP_SAISIECONTRE')='X') ;
  LaDEV.Code:=MM.CodeD ; GetInfosDevise(LaDEV) ;
  {Fin du MM}
  MM.Exo:=QuelExoDT(MM.DateC) ;
  if not EcritureSSTrait then
  begin
    if TOBPiece.GetValue('GP_VENTEACHAT')='ACH' then
    BEGIN
      if LeAcc.isReglement then MM.Nature:='RF' else MM.Nature:='OF' ;
    END else
    BEGIN
      if LeAcc.isReglement then MM.Nature:='RC' else MM.Nature:='OC' ;
    END ;
  end else
  begin
  	MM.Nature:='OD';
  end;
  Result:=MM ;
END ;

Procedure ACC_GetlesMontants ( TOBPiece : TOB ; Montant : double ; Var XP,XD : Double ) ;
Var Taux : Double ;
BEGIN
XP:=0 ; XD:=0 ;
Taux:=TOBPiece.GetValue('GP_TAUXDEV') ;
if TOBPiece.GetValue('GP_DEVISE')<>V_PGI.DevisePivot then
   BEGIN
   XD:=Montant ;
   XP:=DeviseToEuro(Montant,Taux,LaDEV.Quotite) ;
   END else
   BEGIN
   XP:=Montant ; XD:=XP ;
   END ;
END ;

Function ACC_CompleteTiersSSTrait ( TOBPiece,TOBTiers,TOBFournisseur,TOBTTC : TOB ; LeAcc : T_GCAcompte; SensClient : boolean ) : boolean ;
Var TOBG,TOBGF : TOB ;
    i,NumE,NumL  : integer ;
    GColl,GFour, NatEcr  : String ;
    OkVent : boolean ;
    XP,XD : Double ;
BEGIN
  Result:=False ;
  TOBG := nil;
  TOBGF := nil;
  //
  TRY
    {Etude du compte général collectif client}
    GColl:=TOBTiers.GetValue('T_COLLECTIF') ;
    TOBG:=CreerTOBGeneral(GColl) ;
    // Erreur sur le collectif
    if TOBG=Nil then BEGIN LastMsg:=4 ; Exit ; END ;
    {Etude du compte général collectif sous traitant}
    GFour:=TOBFournisseur.GetValue('T_COLLECTIF') ;
    TOBGF:=CreerTOBGeneral(GColl) ;
    // Erreur sur le collectif
    if TOBGF=Nil then BEGIN LastMsg:=4 ; Exit ; END ;
    OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
    {Renseignements complémentaires}
    // --
    if SensClient then
    begin
    	TOBTTC.PutValue('E_AUXILIAIRE',TOBTiers.GetValue('T_AUXILIAIRE')) ;
      TOBTTC.PutValue('E_GENERAL',GColl) ;
      TOBTTC.PutValue('E_CONSO',TOBTiers.GetValue('T_CONSO')) ;
      TOBTTC.PutValue('E_CONTREPARTIEGEN',GFour) ;
      TOBTTC.PutValue('E_CONTREPARTIEAUX',TOBFournisseur.GetString('T_AUXILIAIRE')) ;
    	TOBTTC.PutValue('E_RIB',TobTiers.GetValue('RIB')); // JT eQualité 11032
    	NumL:=1 ; NumE:=1 ;
    end else
    begin
    	NumL:=2 ; NumE:=1 ;
    	TOBTTC.PutValue('E_AUXILIAIRE',TOBFournisseur.GetValue('T_AUXILIAIRE')) ;
      TOBTTC.PutValue('E_GENERAL',GFour) ;
      TOBTTC.PutValue('E_CONSO',TOBFournisseur.GetValue('T_CONSO')) ;
      TOBTTC.PutValue('E_CONTREPARTIEGEN',GColl) ;
      TOBTTC.PutValue('E_CONTREPARTIEAUX',TOBTiers.GetValue('T_AUXILIAIRE')) ;
    end;
    {Pièce}
    TOBTTC.PutValue('E_BLOCNOTE',TOBPiece.GetValue('GP_BLOCNOTE')) ;
    {Divers}
    TOBTTC.PutValue('E_TYPEMVT','DIV') ;
    TOBTTC.PutValue('E_ETATLETTRAGE','AL') ;
    TOBTTC.PutValue('E_EMETTEURTVA','-') ;
    // Modif BTP
    TOBTTC.PutValue('E_NUMTRAITECHQ',LeAcc.NumCheque) ;
    // --
    {Eche+Vent}
    TOBTTC.PutValue('E_ECHE','X') ;
    if OkVent then TOBTTC.PutValue('E_ANA','X') ;
    TOBTTC.PutValue('E_NUMLIGNE',NumL) ; TOBTTC.PutValue('E_NUMECHE',NumE) ;
    {Echéances}
    TOBTTC.PutValue('E_MODEPAIE',LeAcc.ModePaie) ;
    TOBTTC.PutValue('E_CODEACCEPT',MPTOACC(LeAcc.ModePaie)) ;
    TOBTTC.PutValue('E_LIBELLE',LeAcc.Libelle) ;
    //TOBTTC.PutValue('E_DATEECHEANCE',TOBTTC.GetValue('E_DATECOMPTABLE')) ;
    TOBTTC.PutValue('E_DATEECHEANCE',TOBTTC.GetValue('E_DATECOMPTABLE')) ;
    {Montants}
    ACC_GetLesMontants(TOBPiece,LeAcc.Montant,XP,XD) ;
    NatEcr:=TOBTTC.GetValue('E_NATUREPIECE') ;
    if SensClient then
    BEGIN
       TOBTTC.PutValue('E_DEBIT',0)     ; TOBTTC.PutValue('E_CREDIT',XP) ;
       TOBTTC.PutValue('E_DEBITDEV',0)  ; TOBTTC.PutValue('E_CREDITDEV',XD) ;
       TOBTTC.PutValue('E_ENCAISSEMENT',SensEnc(0,XP)) ;
    END else
    BEGIN
       TOBTTC.PutValue('E_DEBIT',XP)     ; TOBTTC.PutValue('E_CREDIT',0) ;
       TOBTTC.PutValue('E_DEBITDEV',XD)  ; TOBTTC.PutValue('E_CREDITDEV',0) ;
       TOBTTC.PutValue('E_ENCAISSEMENT',SensEnc(XP,0)) ;
    END ;
    {Analytiques}
    if TOBTTC.GetValue('E_ANA')='X' then
    BEGIN
       for i:=1 to 5 do TOB.Create('A'+IntToStr(i),TOBTTC,-1) ;
       VentilerTOB(TOBTTC,'',LaDEV.Decimale,2,False) ;
    END ;
    Result:=True ;
  FINALLY
  	{Libérations}
  	if TOBG <> nil then TOBG.Free ;
  	if TOBGF <> nil then TOBGF.Free ;
  END;
end;

Function ACC_CompleteTiers ( TOBPiece,TOBTiers,TOBTTC : TOB ; LeAcc : T_GCAcompte ) : boolean ;
Var TOBG : TOB ;
    i,NumE,NumL  : integer ;
    GColl,NatEcr  : String ;
    OkVent : boolean ;
    XP,XD : Double ;
BEGIN
  Result:=False ;
  {Etude du compte général collectif}
  GColl:=TOBTiers.GetValue('T_COLLECTIF') ;
  TOBG:=CreerTOBGeneral(GColl) ;
  // Erreur sur le collectif
  if TOBG=Nil then BEGIN LastMsg:=4 ; Exit ; END ;
  OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
  {Renseignements complémentaires}
  TOBTTC.PutValue('E_AUXILIAIRE',TOBTiers.GetValue('T_AUXILIAIRE')) ;
  // --
  TOBTTC.PutValue('E_GENERAL',GColl) ;
  TOBTTC.PutValue('E_CONSO',TOBTiers.GetValue('T_CONSO')) ;
  {Pièce}
  TOBTTC.PutValue('E_BLOCNOTE',TOBPiece.GetValue('GP_BLOCNOTE')) ;
  TOBTTC.PutValue('E_RIB',TobTiers.GetValue('RIB')); // JT eQualité 11032
  {Divers}
  TOBTTC.PutValue('E_TYPEMVT','DIV') ;
  TOBTTC.PutValue('E_ETATLETTRAGE','AL') ;
  TOBTTC.PutValue('E_EMETTEURTVA','X') ;
  TOBTTC.PutValue('E_CONTREPARTIEGEN',LeAcc.CpteRegle) ;
  TOBTTC.PutValue('E_CONTREPARTIEAUX','') ;
  // Modif BTP
  TOBTTC.PutValue('E_NUMTRAITECHQ',LeAcc.NumCheque) ;
  // --
  {Eche+Vent}
  NumL:=1 ; NumE:=1 ;
  TOBTTC.PutValue('E_ECHE','X') ;
  if OkVent then TOBTTC.PutValue('E_ANA','X') ;
  TOBTTC.PutValue('E_NUMLIGNE',NumL) ; TOBTTC.PutValue('E_NUMECHE',NumE) ;
  {Echéances}
  TOBTTC.PutValue('E_MODEPAIE',LeAcc.ModePaie) ;
  TOBTTC.PutValue('E_CODEACCEPT',MPTOACC(LeAcc.ModePaie)) ;
  TOBTTC.PutValue('E_LIBELLE',LeAcc.Libelle) ;
  //TOBTTC.PutValue('E_DATEECHEANCE',TOBTTC.GetValue('E_DATECOMPTABLE')) ;
  TOBTTC.PutValue('E_DATEECHEANCE',LeAcc.DateEche) ;
  {Montants}
  ACC_GetLesMontants(TOBPiece,LeAcc.Montant,XP,XD) ;
  NatEcr:=TOBTTC.GetValue('E_NATUREPIECE') ;
  if ((NatEcr='OC') or (NatEcr='RC')) then
     BEGIN
     TOBTTC.PutValue('E_DEBIT',0)     ; TOBTTC.PutValue('E_CREDIT',XP) ;
     TOBTTC.PutValue('E_DEBITDEV',0)  ; TOBTTC.PutValue('E_CREDITDEV',XD) ;
     TOBTTC.PutValue('E_ENCAISSEMENT',SensEnc(0,XP)) ;
     END else
     BEGIN
     TOBTTC.PutValue('E_DEBIT',XP)     ; TOBTTC.PutValue('E_CREDIT',0) ;
     TOBTTC.PutValue('E_DEBITDEV',XD)  ; TOBTTC.PutValue('E_CREDITDEV',0) ;
     TOBTTC.PutValue('E_ENCAISSEMENT',SensEnc(XP,0)) ;
     END ;
  {Analytiques}
  if TOBTTC.GetValue('E_ANA')='X' then
     BEGIN
     for i:=1 to 5 do TOB.Create('A'+IntToStr(i),TOBTTC,-1) ;
     VentilerTOB(TOBTTC,'',LaDEV.Decimale,2,False) ;
     END ;
  {Libérations}
  TOBG.Free ;
  Result:=True ;
END ;

Function ACC_CompleteBanque ( TOBPiece,TOBTiers,TOBTTC : TOB ; LeAcc : T_GCAcompte ) : boolean ;
Var TOBG : TOB ;
    i,NumE,NumL  : integer ;
    Gene,sNatGene,NatEcr : String ;
    OkVent,OkLett : boolean ;
    XP,XD : Double ;
BEGIN
Result:=False ;
{Etude du compte général collectif}
Gene:=LeAcc.CpteRegle ;
TOBG:=CreerTOBGeneral(Gene) ;
// Erreur sur le collectif
if TOBG=Nil then BEGIN LastMsg:=14 ; Exit ; END ;
OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
{Renseignements complémentaires}
TOBTTC.PutValue('E_AUXILIAIRE','') ;
TOBTTC.PutValue('E_GENERAL',Gene) ;
TOBTTC.PutValue('E_CONSO',TOBG.GetValue('G_CONSO')) ;
{Pièce}
TOBTTC.PutValue('E_BLOCNOTE',TOBPiece.GetValue('GP_BLOCNOTE')) ;
TOBTTC.PutValue('E_RIB',TobTiers.GetValue('RIB')); // JT eQualité 11032
{Divers}
TOBTTC.PutValue('E_TYPEMVT','DIV') ;
TOBTTC.PutValue('E_ETATLETTRAGE','RI') ;
TOBTTC.PutValue('E_EMETTEURTVA','-') ;
TOBTTC.PutValue('E_CONTREPARTIEGEN',TOBTiers.GetValue('T_COLLECTIF')) ;
TOBTTC.PutValue('E_CONTREPARTIEAUX',TOBTiers.GetValue('T_AUXILIAIRE')) ;
{Eche+Vent}
NumL:=2 ; NumE:=0 ;
TOBTTC.PutValue('E_ECHE','-') ;
if OkVent then TOBTTC.PutValue('E_ANA','X') ;
TOBTTC.PutValue('E_NUMLIGNE',NumL) ; TOBTTC.PutValue('E_NUMECHE',NumE) ;
sNatGene:=TOBG.GetValue('G_NATUREGENE') ;
if (TOBG.GetValue('G_POINTABLE')='X') and ((sNatGene='BQE') or (sNatGene='CAI')) then
   BEGIN
   TOBTTC.PutValue('E_ECHE','X') ; TOBTTC.PutValue('E_NUMECHE',1) ;
   END ;
{Echéances}
OkLett:=False ;
if TOBG.GetValue('G_LETTRABLE')='X' then OkLett:=True else if ((sNatGene='COC') or (sNatGene='COF')) then OkLett:=True ;
if OkLett then
   BEGIN
   TOBTTC.PutValue('E_ECHE','X') ; TOBTTC.PutValue('E_NUMECHE',1) ; TOBTTC.PutValue('E_ETATLETTRAGE','AL') ;
   TOBTTC.PutValue('E_ETAT','0000000000') ;
   if ((sNatGene='COC') or (sNatGene='COF')) then
      BEGIN
      TOBTTC.PutValue('E_AUXILIAIRE',TOBTiers.GetValue('T_AUXILIAIRE')) ;
      TOBTTC.PutValue('E_CONSO',TOBTiers.GetValue('T_CONSO')) ;
      END ;
   END ;
TOBTTC.PutValue('E_MODEPAIE',LeAcc.ModePaie) ;
TOBTTC.PutValue('E_LIBELLE',LeAcc.Libelle) ;
//TOBTTC.PutValue('E_DATEECHEANCE',TOBTTC.GetValue('E_DATECOMPTABLE')) ;
TOBTTC.PutValue('E_DATEECHEANCE',LeAcc.DateEche) ;
{Montants}
ACC_GetLesMontants(TOBPiece,LeAcc.Montant,XP,XD) ;
NatEcr:=TOBTTC.GetValue('E_NATUREPIECE') ;
if ((NatEcr='OC') or (NatEcr='RC')) then
   BEGIN
   TOBTTC.PutValue('E_DEBIT',XP)     ; TOBTTC.PutValue('E_CREDIT',0) ;
   TOBTTC.PutValue('E_DEBITDEV',XD)  ; TOBTTC.PutValue('E_CREDITDEV',0) ;
   TOBTTC.PutValue('E_ENCAISSEMENT',SensEnc(XP,0)) ;
   END else
   BEGIN
   TOBTTC.PutValue('E_DEBIT',0)     ; TOBTTC.PutValue('E_CREDIT',XP) ;
   TOBTTC.PutValue('E_DEBITDEV',0)  ; TOBTTC.PutValue('E_CREDITDEV',XD) ;
   TOBTTC.PutValue('E_ENCAISSEMENT',SensEnc(0,XP)) ;
   END ;
{Analytiques}
if TOBTTC.GetValue('E_ANA')='X' then
   BEGIN
   for i:=1 to 5 do TOB.Create('A'+IntToStr(i),TOBTTC,-1) ;
   VentilerTOB(TOBTTC,'',LaDEV.Decimale,2,False) ;
   END ;
{Libérations}
TOBG.Free ;
Result:=True ;
END ;

Procedure T_CPAcompte.MajMontantAcc;
var CodeD : string;
Taux : double;
BEGIN
CodeD:=TOBPiece.GetValue('GP_DEVISE') ; Taux:=TOBPiece.GetValue('GP_TAUXDEV') ;
if CodeD<>V_PGI.DevisePivot then
   BEGIN
   NewACC.PutValue('GAC_MONTANT',DeviseToEuro(LeAcc.Montant,Taux,LaDEV.Quotite)) ;
   END else
   BEGIN
   NewACC.PutValue('GAC_MONTANT',LeAcc.Montant) ;
   END ;
END;

Procedure T_CPAcompte.ValideReglSousTrait ;
Var
  TOBEcr,TOBE   : TOB ;
  MM            : RMVT ;
  FlagAcc : String ;
  // Modif BTP
  QQ : Tquery;
  XP,XD : double;
  Indice,i : integer;
  NatEcr : string;
  LATobAcc : TOB;
  SensCommun : boolean;
  TOBEcart,TOBE1   : TOB ;
  DocInfo : T_WSDocumentInf;
BEGIN
  SensCommun := true;
  TOBE := nil;
  TOBEcart := nil;
  if (LeAcc.LienCpta <> 'DIF') and (LeACC.IsModif) then
  BEGIN
    // MISE A JOUR DU REGLEMENT
    LaTobAcc := LeAcc.LaTobAcc;
    DocInfo := GetWSDocumentInfFromTob(TobPiece);
    if LeAcc.IsComptabilisable then
    begin
    	MM.Nature:='OD';
      QQ := OpenSql ('SELECT * FROM ECRITURE WHERE E_JOURNAL="'+LeACC.JalRegle +'" AND E_NATUREPIECE="'+MM.Nature+'" AND E_NUMEROPIECE="'+inttostr(LATobAcc.GetValue('GAC_NUMECR'))+'"',true,-1, '', True);
      if QQ.eof Then BEGIN ferme(QQ);V_PGI.IoError:=oeUnknown ; Exit ; END ;
      TOBEcr:=TOB.Create('',Nil,-1) ;
      TOBEcart := TOB.Create ('',Nil,-1);
      TOBECR.LoadDetailDB ('ECRITURE','','',QQ,false);
      TOBEcart.dupliquer (TOBEcr,true,true);
      ferme(QQ);
      for Indice := 0 to TOBEcr.detail.count -1 do
      begin
        TOBE := TOBEcr.detail[indice]; TOBE1 := TOBEcart.detail[Indice];
        if TOBE.GetValue('E_AUXILIAIRE')<> '' then
        begin
          TOBE.SetString('E_REFGESCOM','');
          // ecriture sur le compte de tiers
          {Montants}
          ACC_GetLesMontants(TOBPiece,LeAcc.Montant,XP,XD) ;
          NatEcr:=TOBE.GetValue('E_NATUREPIECE') ;
          //
          TOBE.PutValue('E_DEBIT',0)     ; TOBE.PutValue('E_CREDIT',XP) ;
          TOBE.PutValue('E_DEBITDEV',0)  ; TOBE.PutValue('E_CREDITDEV',XD) ;
          TOBE.PutValue('E_ENCAISSEMENT',SensEnc(0,XP)) ;
          //
          if XP-TOBE1.GetValue('E_CREDIT') < 0 Then SensCommun := false;
          TOBE1.putvalue ('E_CREDIT',abs(XP-TOBE1.GetValue('E_CREDIT')));
          TOBE1.putvalue ('E_CREDITDEV',abs(XD-TOBE1.GetValue('E_CREDITDEV')));
          {Analytiques}
          if TOBE.GetValue('E_ANA')='X' then
          BEGIN
            for i:=1 to 5 do TOB.Create('A'+IntToStr(i),TOBE,-1) ;
            VentilerTOB(TOBE,'',LaDEV.Decimale,2,False) ;
          END ;
        end else
        begin
          // ecriture sur le compte fournisseur
          {Montants}
          TOBE.SetString('E_REFGESCOM','');
          ACC_GetLesMontants(TOBPiece,LeAcc.Montant,XP,XD) ;
          NatEcr:=TOBE.GetValue('E_NATUREPIECE') ;
          TOBE.PutValue('E_DEBIT',XP)     ; TOBE.PutValue('E_CREDIT',0) ;
          TOBE.PutValue('E_DEBITDEV',XD)  ; TOBE.PutValue('E_CREDITDEV',0) ;
          TOBE.PutValue('E_ENCAISSEMENT',SensEnc(XP,0)) ;
          //
          if XP-TOBE1.GetValue('E_DEBIT') < 0 Then SensCommun := false;
          TOBE1.putvalue ('E_DEBIT',abs(XP-TOBE1.GetValue('E_DEBIT')));
          TOBE1.putvalue ('E_DEBITDEV',abs(XD-TOBE1.GetValue('E_DEBITDEV')));
          {Analytiques}
          if TOBE.GetValue('E_ANA')='X' then
          BEGIN
            for i:=1 to 5 do TOB.Create('A'+IntToStr(i),TOBE,-1) ;
            VentilerTOB(TOBE,'',LaDEV.Decimale,2,False) ;
          END ;
        end;
      end;
      // phase de mise a jour
      if Not TOBEcr.InsertOrUpdateDB(true) then
      begin
        TOBEcr.Free;
        TOBEcart.free;
        V_PGI.IoError:=oeUnknown ;
        Exit ;
      end else
      begin
        CallWSCegid(wsetSubContractPayment, TOBEcr, DocInfo);
      end;
    end;
    NewAcc.putvalue('GAC_MONTANTDEV',LeAcc.Montant);
    MajMontantAcc ;
    if LeAcc.IsComptabilisable then
    begin
      if (TOBEcart.detail[0].getValue('E_DEBIT') - TOBEcart.detail[0].getValue('E_CREDIT')) <> 0 then
      MajSoldesEcritureTOB(TOBEcart,SensCommun) ;
      TobEcart.free;
    end;
  END ELSE
  BEGIN
    if LeAcc.IsComptabilisable then
    begin
      {Record comptable}
      MM:=ACC_FabriqueMM(TOBPiece,LeAcc,true) ;
      if LastMsg>0 then BEGIN V_PGI.IoError:=oeUnknown ; Exit ; END ;
      {Ecritures}
      TOBEcr:=TOB.Create('',Nil,-1) ;
      {Ligne de tiers -- Client payeur}
      TOBE:=TOB.Create('ECRITURE',TOBEcr,-1) ;
      PieceVersECRSsTrait(MM,TOBPiece,TOBTiers,TOBE,True) ;
      if Not ACC_CompleteTiersSSTrait(TOBPiece,TOBTiers,TOBFournisseur,TOBE,LeAcc,True) then BEGIN TOBEcr.Free ; V_PGI.IoError:=oeUnknown ; Exit ; END ;
      TOBE.SetString('E_REFGESCOM',''); // pour eviter le bug de la ref gescom possitionné sur l'écriture de reglement
      {Ligne Sous traitant}
      TOBE:=TOB.Create('ECRITURE',TOBEcr,-1) ;
      PieceVersECRSsTrait(MM,TOBPiece,TOBTiers,TOBE,True) ;
      if Not ACC_CompleteTiersSSTrait(TOBPiece,TOBTiers,TOBFournisseur,TOBE,LeAcc,false) then BEGIN TOBEcr.Free ; V_PGI.IoError:=oeUnknown ; Exit ; END ;
      TOBE.SetString('E_REFGESCOM',''); // pour eviter le bug de la ref gescom possitionné sur l'écriture de reglement
      {Enregistrement}
      if IsComptaPce(TOBPiece.GetValue('GP_NATUREPIECEG')) then  //AC 18/08/03 NV GESTION COMPTA DIFF
      begin
        if LeAcc.LienCpta<>'DIF' then
        BEGIN
          {MAJ des soldes comptables}
          if MM.Simul='N' then MajSoldesEcritureTOB(TOBEcr,True) ;
          FlagAcc:='R' ;
          {Insertion}
          if Not TOBEcr.InsertDB(Nil) then BEGIN TOBEcr.Free ; V_PGI.IoError:=oeUnknown ; Exit ; END ;
        END else
        BEGIN
          FlagAcc:='D' ;
          if Not InsertionDifferee(TOBEcr) then BEGIN TOBEcr.Free ; V_PGI.IoError:=oeUnknown ; Exit ; END ;
        END ;
      end ;
      {Création TOB Acompte}
      if TOBEcr.Detail.Count<=0 then
      begin
        TOBEcr.Free ;
        V_PGI.IoError:=oeUnknown ;
        Exit ;
      end else
        CallWSCegid(wsetSubContractPayment, TOBEcr, DocInfo);
      TOBE:=TOBEcr.Detail[0] ;
    end;
    NewAcc:=TOB.Create('ACOMPTES',Nil,-1) ;
    NewAcc.AddChampSup('NEWACC',False) ; NewAcc.PutValue('NEWACC',FlagAcc) ;
    // Modif BTP
    NewAcc.AddChampSup('MONTANTINIT',False) ; NewAcc.PutValue('MONTANTINIT',LeACC.Montant) ;
    // --
    NewACC.PutValue('GAC_NATUREPIECEG',TOBPiece.GetValue('GP_NATUREPIECEG')) ;
    NewACC.PutValue('GAC_SOUCHE',TOBPiece.GetValue('GP_SOUCHE')) ;
    NewACC.PutValue('GAC_NUMERO',TOBPiece.GetValue('GP_NUMERO')) ;
    NewACC.PutValue('GAC_INDICEG',TOBPiece.GetValue('GP_INDICEG')) ;
    //
    if LeAcc.DateEcr <> '' then
    begin
      NewACC.PutValue('GAC_DATEECR',StrToDate(LeAcc.DateEcr));
    end else
    begin
      NewACC.PutValue('GAC_DATEECR',TOBPiece.GetValue('GP_DATEPIECE'));
    end;
    // --------- Date Echeance ----------------------------------------------
    if LeAcc.DateEche <> '' then
    begin
      NewACC.PutValue('GAC_DATEECHEANCE',StrToDate(LeAcc.DateEche));
    end else
    begin
      NewACC.PutValue('GAC_DATEECHEANCE',TOBPiece.GetValue('GP_DATEPIECE'));
    end;
    // ----------------------------------------------------------------------

    //
    if LeAcc.IsComptabilisable then
    begin
      NewACC.PutValue('GAC_JALECR',TOBE.GetValue('E_JOURNAL')) ;
      NewACC.PutValue('GAC_NUMECR',TOBE.GetValue('E_NUMEROPIECE')) ;
    end else
    begin
      NewACC.PutValue('GAC_JALECR','') ;
      NewACC.PutValue('GAC_NUMECR',0) ;
    end;
    NewACC.PutValue('GAC_MODEPAIE',LeAcc.ModePaie) ;
    NewACC.PutValue('GAC_MONTANTDEV',LeAcc.Montant) ;
    NewACC.PutValue('GAC_CBLIBELLE',LeAcc.CBLibelle) ;
    NewACC.PutValue('GAC_CBINTERNET',LeAcc.CBInternet) ;
    NewACC.PutValue('GAC_DATEEXPIRE',LeAcc.DateExpire) ;
    NewACC.PutValue('GAC_TYPECARTE',LeAcc.TypeCarte) ;
    NewACC.PutValue('GAC_NUMCHEQUE',LeAcc.NumCheque) ;
    NewACC.PutValue('GAC_AUXILIAIRE',LeAcc.AuxiFacture ) ;
    NewACC.PutValue('GAC_QUALIFPIECE', MM.Simul) ;
    // --
    if LeAcc.IsReglement then NewACC.PutValue('GAC_ISREGLEMENT','X') else NewACC.PutValue('GAC_ISREGLEMENT','-');
    // Modif BTP
    NewACC.PutValue('GAC_LIBELLE',LeAcc.Libelle );
    MajMontantAcc ;
  END;
  {Libérations}
  TOBEcr.Free ;
end;

Procedure T_CPAcompte.ValideAcompte ;
Var
  TOBEcr,TOBE   : TOB ;
  MM            : RMVT ;
  FlagAcc : String ;
  QQ : Tquery;
  XP,XD : double;
  Indice,i : integer;
  NatEcr : string;
  LATobAcc : TOB;
  SensCommun : boolean;
  TOBEcart,TOBE1   : TOB ;
  DocInfo : T_WSDocumentInf;
BEGIN
  {$IFDEF BTP}
  SensCommun := true;
  TOBE := nil;
  TOBEcart := nil;
  DocInfo := GetWSDocumentInfFromTob(TobPiece);
  if (LeAcc.LienCpta <> 'DIF') and (LeACC.IsModif) then
  begin
    LaTobAcc := LeAcc.LaTobAcc;
    if LeAcc.IsComptabilisable then
    begin
      if TOBPiece.GetValue('GP_VENTEACHAT')='ACH' then
      begin
        if LeAcc.isReglement then
          MM.Nature:='RF'
        else
          MM.Nature:='OF';
      end else
      begin
        if LeAcc.isReglement then
          MM.Nature:='RC'
        else
          MM.Nature:='OC' ;
      end;
      QQ := OpenSql ('SELECT * FROM ECRITURE WHERE E_JOURNAL="'+LeACC.JalRegle +'" AND E_NATUREPIECE="'+MM.Nature+'" AND E_NUMEROPIECE="'+inttostr(LATobAcc.GetValue('GAC_NUMECR'))+'"',true,-1, '', True);
      try
        if not QQ.Eof then
        begin
          TOBEcr   := TOB.Create('',Nil,-1) ;
          TOBEcart := TOB.Create ('',Nil,-1);
          TOBECR.LoadDetailDB ('ECRITURE','','',QQ,false);
          TOBEcart.dupliquer (TOBEcr,true,true);
        end else
          V_PGI.IoError := oeUnknown ;
      finally
        Ferme(QQ);
      end;
      if V_PGI.IoError <> oeOk then
        Exit;
      for Indice := 0 to TOBEcr.detail.count -1 do
      begin
        TOBE := TOBEcr.detail[indice];
        TOBE1 := TOBEcart.detail[Indice];
        if TOBE.GetValue('E_AUXILIAIRE')<> '' then
        begin
          TOBE.SetString('E_REFGESCOM','');
          // ecriture sur le compte de tiers
          {Montants}
          ACC_GetLesMontants(TOBPiece,LeAcc.Montant,XP,XD) ;
          NatEcr := TOBE.GetValue('E_NATUREPIECE') ;
          if ((NatEcr='OC') or (NatEcr='RC')) then
          begin
            TOBE.PutValue('E_DEBIT',0)     ; TOBE.PutValue('E_CREDIT',XP) ;
            TOBE.PutValue('E_DEBITDEV',0)  ; TOBE.PutValue('E_CREDITDEV',XD) ;
            TOBE.PutValue('E_ENCAISSEMENT',SensEnc(0,XP)) ;
            //
            if XP-TOBE1.GetValue('E_CREDIT') < 0 Then SensCommun := false;
            TOBE1.putvalue ('E_CREDIT',abs(XP-TOBE1.GetValue('E_CREDIT')));
            TOBE1.putvalue ('E_CREDITDEV',abs(XD-TOBE1.GetValue('E_CREDITDEV')));
          end else
          begin
            TOBE.PutValue('E_DEBIT',XP)     ; TOBE.PutValue('E_CREDIT',0) ;
            TOBE.PutValue('E_DEBITDEV',XD)  ; TOBE.PutValue('E_CREDITDEV',0) ;
            TOBE.PutValue('E_ENCAISSEMENT',SensEnc(XP,0)) ;
            //
            if XP-TOBE1.GetValue('E_DEBIT') < 0 Then SensCommun := false;
            TOBE1.putvalue ('E_DEBIT',abs(XP-TOBE1.GetValue('E_DEBIT')));
            TOBE1.putvalue ('E_DEBITDEV',abs(XD-TOBE1.GetValue('E_DEBITDEV')));
          end;
          {Analytiques}
          if TOBE.GetBoolean('E_ANA') then
          begin
            for i := 1 to 5 do
              TOB.Create('A'+IntToStr(i),TOBE,-1) ;
            VentilerTOB(TOBE,'',LaDEV.Decimale,2,False) ;
          end;
        end else
        begin
          // ecriture sur le compte de banque
          {Montants}
          TOBE.SetString('E_REFGESCOM','');
          ACC_GetLesMontants(TOBPiece,LeAcc.Montant,XP,XD) ;
          NatEcr:=TOBE.GetValue('E_NATUREPIECE') ;
          if ((NatEcr='OC') or (NatEcr='RC')) then
          begin
            TOBE.PutValue('E_DEBIT',XP)     ; TOBE.PutValue('E_CREDIT',0) ;
            TOBE.PutValue('E_DEBITDEV',XD)  ; TOBE.PutValue('E_CREDITDEV',0) ;
            TOBE.PutValue('E_ENCAISSEMENT',SensEnc(XP,0)) ;
            //
            if XP-TOBE1.GetValue('E_DEBIT') < 0 Then SensCommun := false;
            TOBE1.putvalue ('E_DEBIT',abs(XP-TOBE1.GetValue('E_DEBIT')));
            TOBE1.putvalue ('E_DEBITDEV',abs(XD-TOBE1.GetValue('E_DEBITDEV')));
          end else
          begin
            TOBE.PutValue('E_DEBIT',0)     ; TOBE.PutValue('E_CREDIT',XP) ;
            TOBE.PutValue('E_DEBITDEV',0)  ; TOBE.PutValue('E_CREDITDEV',XD) ;
            TOBE.PutValue('E_ENCAISSEMENT',SensEnc(0,XP)) ;
            //
            if XP-TOBE1.GetValue('E_CREDIT') < 0 Then SensCommun := false;
            TOBE1.putvalue ('E_CREDIT',abs(XP-TOBE1.GetValue('E_CREDIT')));
            TOBE1.putvalue ('E_CREDITDEV',abs(XD-TOBE1.GetValue('E_CREDITDEV')));
          end;
          {Analytiques}
          if TOBE.GetValue('E_ANA')='X' then
          begin
             for i:=1 to 5 do TOB.Create('A'+IntToStr(i),TOBE,-1) ;
             VentilerTOB(TOBE,'',LaDEV.Decimale,2,False) ;
          end;
        end;
      end;
      // phase de mise a jour
      if Not TOBEcr.InsertOrUpdateDB(true) then
      begin
        TOBEcr.Free;
        TOBEcart.free;
        V_PGI.IoError:=oeUnknown ;
        Exit ;
      end else
        CallWSCegid(wsetPayment, TOBEcr, DocInfo);
    end;
    NewAcc.putvalue('GAC_MONTANTDEV',LeAcc.Montant);
    MajMontantAcc ;
    if LeAcc.IsComptabilisable then
    begin
     if (TOBEcart.detail[0].getValue('E_DEBIT') - TOBEcart.detail[0].getValue('E_CREDIT')) <> 0 then
        MajSoldesEcritureTOB(TOBEcart,SensCommun) ;
     TobEcart.free;
    end;
  end else
  begin
    {$ENDIF}
    if LeAcc.IsComptabilisable then
    begin
      {Record comptable}
      MM:=ACC_FabriqueMM(TOBPiece,LeAcc) ;
      if LastMsg>0 then BEGIN V_PGI.IoError:=oeUnknown ; Exit ; END ;
      {Ecritures}
      TOBEcr:=TOB.Create('',Nil,-1) ;
      {Ligne de tiers}
      TOBE:=TOB.Create('ECRITURE',TOBEcr,-1) ;
      PieceVersECR(MM,TOBPiece,TOBTiers,TOBE,True) ;
      if Not ACC_CompleteTiers(TOBPiece,TOBTiers,TOBE,LeAcc) then BEGIN TOBEcr.Free ; V_PGI.IoError:=oeUnknown ; Exit ; END ;
      TOBE.SetString('E_REFGESCOM',''); // pour eviter le bug de la ref gescom possitionné sur l'écriture de reglement
      {Ligne de banque}
      TOBE:=TOB.Create('ECRITURE',TOBEcr,-1) ;
      PieceVersECR(MM,TOBPiece,TOBTiers,TOBE,True) ;
      if Not ACC_CompleteBanque(TOBPiece,TOBTiers,TOBE,LeAcc) then BEGIN TOBEcr.Free ; V_PGI.IoError:=oeUnknown ; Exit ; END ;
      TOBE.SetString('E_REFGESCOM',''); // pour eviter le bug de la ref gescom possitionné sur l'écriture de reglement
      {Enregistrement}
      if IsComptaAcc(TOBPiece.GetValue('GP_NATUREPIECEG')) then  //AC 18/08/03 NV GESTION COMPTA DIFF
      begin
        if LeAcc.LienCpta <> 'DIF' then
        begin
          {MAJ des soldes comptables}
          if MM.Simul='N' then MajSoldesEcritureTOB(TOBEcr,True) ;
          FlagAcc:='R' ;
          {Insertion}
          if Not TOBEcr.InsertDB(Nil) then
          begin
            TOBEcr.Free ;
            V_PGI.IoError:=oeUnknown ;
            Exit ;
          end else
            CallWSCegid(wsetPayment, TOBEcr, DocInfo);
        end else
        begin
          FlagAcc:='D' ;
          if Not InsertionDifferee(TOBEcr) then BEGIN TOBEcr.Free ; V_PGI.IoError:=oeUnknown ; Exit ; END ;
        end;
      end ;
      {Création TOB Acompte}
      if TOBEcr.Detail.Count<=0 then BEGIN TOBEcr.Free ; V_PGI.IoError:=oeUnknown ; Exit ; END ;
      TOBE:=TOBEcr.Detail[0] ;
    end;
    NewAcc:=TOB.Create('ACOMPTES',Nil,-1) ;
    NewAcc.AddChampSup('NEWACC',False) ; NewAcc.PutValue('NEWACC',FlagAcc) ;
    // Modif BTP
    NewAcc.AddChampSup('MONTANTINIT',False) ; NewAcc.PutValue('MONTANTINIT',LeACC.Montant) ;
    // --
    NewACC.PutValue('GAC_NATUREPIECEG',TOBPiece.GetValue('GP_NATUREPIECEG')) ;
    NewACC.PutValue('GAC_SOUCHE',TOBPiece.GetValue('GP_SOUCHE')) ;
    NewACC.PutValue('GAC_NUMERO',TOBPiece.GetValue('GP_NUMERO')) ;
    NewACC.PutValue('GAC_INDICEG',TOBPiece.GetValue('GP_INDICEG')) ;
    //
    if LeAcc.DateEcr <> '' then
    begin
    NewACC.PutValue('GAC_DATEECR',StrToDate(LeAcc.DateEcr));
    end else
    begin
    NewACC.PutValue('GAC_DATEECR',TOBPiece.GetValue('GP_DATEPIECE'));
    end;
    // --------- Date Echeance ----------------------------------------------
    if LeAcc.DateEche <> '' then
    begin
    NewACC.PutValue('GAC_DATEECHEANCE',StrToDate(LeAcc.DateEche));
    end else
    begin
    NewACC.PutValue('GAC_DATEECHEANCE',TOBPiece.GetValue('GP_DATEPIECE'));
    end;
    // ----------------------------------------------------------------------

    //
    if LeAcc.IsComptabilisable then
    begin
     NewACC.PutValue('GAC_JALECR',TOBE.GetValue('E_JOURNAL')) ;
     NewACC.PutValue('GAC_NUMECR',TOBE.GetValue('E_NUMEROPIECE')) ;
    end else
    begin
     NewACC.PutValue('GAC_JALECR','') ;
     NewACC.PutValue('GAC_NUMECR',0) ;
    end;
    NewACC.PutValue('GAC_MODEPAIE',LeAcc.ModePaie) ;
    NewACC.PutValue('GAC_MONTANTDEV',LeAcc.Montant) ;
    NewACC.PutValue('GAC_CBLIBELLE',LeAcc.CBLibelle) ;
    NewACC.PutValue('GAC_CBINTERNET',LeAcc.CBInternet) ;
    NewACC.PutValue('GAC_DATEEXPIRE',LeAcc.DateExpire) ;
    NewACC.PutValue('GAC_TYPECARTE',LeAcc.TypeCarte) ;
    NewACC.PutValue('GAC_NUMCHEQUE',LeAcc.NumCheque) ;
    NewACC.PutValue('GAC_AUXILIAIRE',LeAcc.AuxiFacture ) ;
    NewACC.PutValue('GAC_QUALIFPIECE', MM.Simul) ;
    // --
    if LeAcc.IsReglement then NewACC.PutValue('GAC_ISREGLEMENT','X') else NewACC.PutValue('GAC_ISREGLEMENT','-');
    {$IFDEF CHR}
    NewACC.PutValue('GAC_LIBELLE', DateToStr(TOBPiece.GetValue('GP_DATEPIECE')) + ';' + copy(TobPiece.Getvalue('GP_HRDOSSIER'), 13, 3) + ';' + '1');
    {$ELSE}
    // Modif BTP
    NewACC.PutValue('GAC_LIBELLE',LeAcc.Libelle );
    // --
    {$ENDIF}
    MajMontantAcc ;
    {$IFDEF BTP}
  end;
  {$ENDIF}
  {Libérations}
  TOBEcr.Free ;
END ;


Function  EnregistreReglSSTrait ( TOBPiece,TOBClient,TOBFournisseur : TOB ; LeAcc : T_GCAcompte ; Quiet : boolean = false ) : TOB ;
Var io  : TIoErr ;
    Acc : T_CPAcompte ;
    Titre : String ;
    TOBAcc,TOBTiersCpta : TOB ;
BEGIN
Result:=Nil ;
Acc:=T_CPAcompte.Create ; LastMsg:=-1 ;
TOBTiersCpta:=TOB.Create('TIERS',Nil,-1) ; TOBTiersCpta.Dupliquer(TOBClient,True,True) ;
RenseigneTiersFact(TOBPiece,TOBClient,TOBTiersCpta) ;
LeAcc.AuxiFacture := TOBTiersCpta.GetValue('T_AUXILIAIRE');
Acc.LeAcc:=LeAcc ; Acc.TOBPiece:=TOBPiece ; Acc.TOBTiers:=TOBTiersCpta ; Acc.TOBFournisseur := TOBFournisseur;
// Modif BTP
if LeAcc.latobAcc <> nil then Acc.NewAcc := LeAcc.LaTobAcc;
// --
io:=Transactions(Acc.ValideReglSousTrait,0) ;
if io<>oeOk then
   BEGIN
   Titre:='ATTENTION : Règlement sous traitant non enregistré' ;
   if LastMsg>0 then Titre:=Titre+' : '+TexteMessage[LastMsg] ;
   if not Quiet then MessageAlerte(Titre) ;
   END else
   BEGIN
   if (Acc.NewAcc<>Nil) and (LeAcc.LaTOBacc = nil) then
      BEGIN
      TOBAcc:=TOB.Create('ACOMPTES',Nil,-1) ;
      TOBAcc.Dupliquer(Acc.NewAcc,True,True) ;
      Result:=TOBAcc ;
      END ELSE
      BEGIN
      // Modification
      TOBAcc:=TOB.Create('ACOMPTES',Nil,-1) ;
      TOBAcc.Dupliquer(Acc.NewAcc,True,True) ;
      Result := TOBAcc;
      END;
   END ;
ACC.TOBPiece:=Nil ; ACC.TOBTiers:=Nil ; Acc.Free ;
TobTiersCpta.Free ;
end;

Function  EnregistreAcompte ( TOBPiece,TOBTiers : TOB ; LeAcc : T_GCAcompte ; Quiet : boolean = false ) : TOB ;
Var io  : TIoErr ;
    Acc : T_CPAcompte ;
    Titre : String ;
    TOBAcc,TOBTiersCpta : TOB ;
BEGIN
  Result:=Nil ;
  Acc:=T_CPAcompte.Create ; LastMsg:=-1 ;
  TOBTiersCpta:=TOB.Create('TIERS',Nil,-1) ; TOBTiersCpta.Dupliquer(TOBTiers,True,True) ;
  RenseigneTiersFact(TOBPiece,TOBTiers,TOBTiersCpta) ;
  LeAcc.AuxiFacture := TOBTiersCpta.GetValue('T_AUXILIAIRE');
  Acc.LeAcc:=LeAcc ; Acc.TOBPiece:=TOBPiece ; Acc.TOBTiers:=TOBTiersCpta ;
  // Modif BTP
  if LeAcc.latobAcc <> nil then Acc.NewAcc := LeAcc.LaTobAcc;
  // --
  io:=Transactions(Acc.ValideAcompte,2) ;
  if io<>oeOk then
     BEGIN
     Titre:='ATTENTION : Acompte/Règlement non enregistré' ;
     if LastMsg>0 then Titre:=Titre+' : '+TexteMessage[LastMsg] ;
     if not Quiet then MessageAlerte(Titre) ;
     END else
     BEGIN
     if (Acc.NewAcc<>Nil) and (LeAcc.LaTOBacc = nil) then
        BEGIN
        TOBAcc:=TOB.Create('ACOMPTES',Nil,-1) ;
        TOBAcc.Dupliquer(Acc.NewAcc,True,True) ;
        Result:=TOBAcc ;
        END ELSE
        BEGIN
        // Modification
        TOBAcc:=TOB.Create('ACOMPTES',Nil,-1) ;
        TOBAcc.Dupliquer(Acc.NewAcc,True,True) ;
        Result := TOBAcc;
        END;
     END ;
  ACC.TOBPiece:=Nil ; ACC.TOBTiers:=Nil ; Acc.Free ;
  TobTiersCpta.Free ;
END ;

Procedure DetruitAcompteGCCpta(TOBPiece,TOBAcc : TOB ; ModifRefGC : boolean = False; VideRefGC : boolean = False) ;
Var
  FlagAcc,JalEcr,NewRef,QualifP, ReqCptaDiff : String ;
  NumEcr         : integer ;
  Q              : TQuery ;
  DD             : TDateTime ;
  TOBEcr         : TOB ;
  XX             : RMVT ;
BEGIN
  if (Assigned(TobAcc)) and (TobAcc.FieldExists('NEWACC')) then
  begin
    NewRef  := EncodeRefCPGescom(TOBPiece) ;
    FlagAcc := TOBAcc.GetValue('NEWACC') ;
    if VH_GC.GCIfDefCEGID then
       QualifP := 'S'
    else
      QualifP := 'N';
    if FlagAcc = 'R' then
    begin
      FillChar(XX,SizeOf(XX),#0);
      JalEcr := TOBAcc.GetValue('GAC_JALECR');
      NumEcr := TOBAcc.GetValue('GAC_NUMECR');
      Q := OpenSQL('SELECT * FROM ECRITURE WHERE E_JOURNAL="'+JalEcr+'" AND E_NUMEROPIECE='+IntToStr(NumEcr)+' AND E_QUALIFPIECE="'+QualifP+'"',True,-1, '', True) ;
      if Not Q.EOF then
      begin
        TOBEcr := TOB.Create('ECRITURE',Nil,-1) ;
        try
          TOBEcr.SelectDB('',Q,True);
          XX := TOBToIdent(TOBEcr,False);
        finally
          TOBEcr.Free;
        end;
      end ;
      Ferme(Q) ;
      if XX.Jal <> '' then
      begin
        XX.DocType   := TobAcc.GetString('GAC_NATUREPIECEG');
        XX.DocNumber := TobAcc.GetInteger('GAC_NUMERO');
        if not ModifRefGC then  // JT eQualité 10246
        begin
          DetruitEcritureFromMM(XX, False, NowH)
        end else
        begin
          if VideRefGC then // JT eQualité 10246
            ModifEcritureFromMM(XX, False, '')
          else
            ModifEcritureFromMM(XX, False, NewRef);
        end;
      end;
    end else
    begin
      DD := TOBPiece.GetValue('GP_DATEPIECE') ;
      ReqCptaDiff := 'FROM COMPTADIFFEREE'
                   + ' WHERE GCD_REFPIECE = "' + NewRef + '"'
                   + '   AND GCD_DATEPIECE = "' + UsDateTime(DD) + '"'
                   + '   AND GCD_USER = "' + V_PGI.User + '"';
      if not ModifRefGC then // JT eQualité 10246
      begin
        ExecuteSQL('DELETE '+ ReqCptaDiff)
      end else
      begin
        if VideRefGC then // JT eQualité 10246
          ModifEcritureDifferee(ReqCptaDiff,'',False)
        else
          ModifEcritureDifferee(ReqCptaDiff,NewRef,False);
      end;
    end;
  end;
end;

Procedure ReinitAcompteGCCpta ( TOBAcc : TOB ) ;
Var FlagAcc,JalEcr,QualifP : String ;
    NumEcr         : integer ;
    Q              : TQuery ;
    TOBEcr         : TOB ;
    XX             : RMVT ;
    indice         : integer;
begin
if TOBAcc=Nil then Exit ;
if Not TOBAcc.FieldExists('NEWACC') then Exit ;
FlagAcc:=TOBAcc.GetValue('NEWACC') ;
QualifP:='N' ;
if VH_GC.GCIfDefCEGID then
   QualifP:='S' ;
if FlagAcc='R' then
   BEGIN
   FillChar(XX,SizeOf(XX),#0) ;
   JalEcr:=TOBAcc.GetValue('GAC_JALECR') ; NumEcr:=TOBAcc.GetValue('GAC_NUMECR') ;
   Q:=OpenSQL('SELECT * FROM ECRITURE WHERE E_JOURNAL="'+JalEcr+'" AND E_NUMEROPIECE='+IntToStr(NumEcr)+' AND E_QUALIFPIECE="'+QualifP+'"',True,-1, '', True) ;
   if Not Q.EOF then
      BEGIN
      TOBEcr:=TOB.Create('ECRITURE',Nil,-1) ;
      TOBEcr.loaddetailDB('ECRITURE','','',Q,false);
      for Indice := 0 to TOBEcr.detail.count -1 do
          begin
         {Reinit Le N° piece}
         TOBEcr.detail[Indice].putvalue('E_REFGESCOM','');
         {Reinit Le N° Affaire}
         TOBEcr.detail[Indice].putvalue('E_AFFAIRE','');
         end;
      TOBECR.SetAllModifie (true);
      TOBEcr.UpdateDB (true);
      TOBEcr.Free ;
      END ;
   Ferme(Q) ;
   END;
end;

Procedure MajAccRegleDiff ( TOBPiece,TOBAcomptes : TOB ; OldNum : integer) ;
Var LienCpta,NaturePieceG,OldRef,NewRef : String ;
    DD                    : TDateTime ;
    NewNum,iNum           : integer ;
BEGIN
if ((TOBPiece=Nil) or (TOBAcomptes=Nil)) then Exit ;
if TOBAcomptes.Detail.Count<=0 then Exit ;
NaturePieceG:=TOBPiece.GetValue('GP_NATUREPIECEG') ;
iNum:=TOBPiece.GetNumChamp('GP_NUMERO') ; NewNum:=TOBPiece.GetValeur(iNum) ;
LienCpta:=GetInfoParPiece(NaturePieceG,'GPP_TYPEPASSACC') ; if LienCpta<>'DIF' then Exit ;
TOBPiece.PutValeur(iNum,OldNum) ; OldRef:=EncodeRefCPGescom(TOBPiece) ;
TOBPiece.PutValeur(iNum,NewNum) ; NewRef:=EncodeRefCPGescom(TOBPiece) ;
DD:=TOBPiece.GetValue('GP_DATEPIECE') ;
ExecuteSQL('UPDATE COMPTADIFFEREE SET GCD_REFPIECE="'+NewRef+'", GCD_RANG=GCD_RANG+10 WHERE GCD_DATEPIECE="'+UsDateTime(DD)+'" AND GCD_REFPIECE="'+OldRef+'" AND GCD_USER="'+V_PGI.User+'"') ;
END ;

Procedure RechLibTiersCegid (VenteAchat : string; TobTiers,TOBPiece,TobAdresses: TOB) ;
BEGIN
if VenteAchat='VEN' then
   BEGIN
   TOBTiers.AddChampSup('LIBTIERSCEGID',False) ;
   TOBTiers.PutValue('LIBTIERSCEGID',TOBTiers.GetValue('T_LIBELLE')) ;
   if TOBPiece.GetValue('GP_TIERSFACTURE')=TOBPiece.GetValue('GP_TIERSLIVRE') then
      BEGIN
      if TOBAdresses.Detail.Count=2 then TOBTiers.PutValue('LIBTIERSCEGID',TOBAdresses.Detail[1].GetValue('GPA_LIBELLE')) ;
      END ;
   END ;
END;

function VerifEcriturePieceModifiable (TOBPiece : TOB) : boolean;
var MM : RMVT;
		Q : Tquery;
    TOBECR : TOB;
    Indice : integer;
    NbAcomptes : Integer;
    Sql : string;
    Okok : boolean;
    DateClo : TDateTime;
begin
  result := true;
  NbAcomptes := 0;
  Sql := 'SELECT COUNT(*) AS NBR FROM ACOMPTES WHERE '+
  				'GAC_NATUREPIECEG="'+TOBPiece.getString('GP_NATUREPIECEG')+'" AND '+
          'GAC_SOUCHE="'+TOBPiece.getString('GP_SOUCHE')+'" AND '+
          'GAC_NUMERO='+InttOStr(TOBPiece.getInteger('GP_NUMERO'))+' AND '+
          'GAC_INDICEG='+InttOStr(TOBPiece.getInteger('GP_INDICEG'));
  Q := OpenSQL(Sql,True,1,'',true);
  if not Q.Eof then
  begin
    NbAcomptes := Q.Fields[0].asInteger;
  end;
  ferme (Q);
  TOBEcr := TOB.Create ('LES ECRITURES',nil,-1);
  TRY
    //
    if GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'),'GPP_TYPEPASSCPTA') = 'AUC' then Exit;
    //
    if isExerciceClo (TOBPiece.GetValue('GP_DATEPIECE')) then
    begin
      PGIInfo(TraduireMemoire('Cette pièce est sur un exercice clôturé en comptabilité.'),'ATTENTION');
      Result := False;
      Exit;
    end;

    // Contrôle Journal période cloturée
    if GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'),'GPP_JOURNALCPTA') <> '' then
    begin
      Q := OpenSql ('SELECT CCJ_PERIODCLOTURE FROM CLOTPERJOU WHERE CCJ_JOURNAL="'+
          GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'),'GPP_JOURNALCPTA')+'" AND CCJ_ENTITY=0',true,1,'',true);
      if not Q.Eof then
      begin
        DateClo := FINDEMOIS(Q.fields[0].AsDateTime);
        if TOBPiece.GetValue('GP_DATEPIECE') <= DateClo then
        begin
          PGIInfo(TraduireMemoire('Cette pièce est sur une période clôturée en comptabilité.'),'ATTENTION');
          Result := False;
        end;
      end;
      ferme (Q);
    end;
    if not Result then Exit;

    // Contrôle période cloturée
    if (Result) and (TOBPiece.GetValue('GP_DATEPIECE')<= GetParamSocSecur('SO_DATECLOTUREPER',iDate1900)) then
    begin
      PGIInfo(TraduireMemoire('Cette pièce est sur une période clôturée en comptabilité.'),'ATTENTION');
      Result := False;
      Exit;
    end;
    //
    if TOBPiece.GetValue('GP_REFCOMPTABLE') = '' then Exit;
    // contrôle lettrage
    MM := DecodeRefGCComptable (TOBPiece.GetValue('GP_REFCOMPTABLE'));
    Q:=OpenSQL('SELECT E_LETTRAGE,E_VALIDE,E_EXPORTE FROM ECRITURE WHERE '+WhereEcriture(tsGene,MM,False),True,-1, '', True) ;
    okok := not Q.Eof;
    if (result) and (okok) then
    begin
      TOBECR.loadDetailDb ('ECRITURE','','',Q,false);
      for Indice := 0 to TOBECR.detail.count -1 do
      begin
        if (TOBECR.Detail[Indice].getValue('E_LETTRAGE')<>'') then
        begin
          result := false;
          PGIInfo(TraduireMemoire('Cette pièce est lettrée en comptabilité via le suivi des règlements.'),'ATTENTION');
          break;
        end;
        if (TOBECR.Detail[Indice].getValue('E_EXPORTE')='X') and (not GetParamSocSecur('SO_SO_BTAUTOREXPORT',false)) then
        begin
          result := false;
          PGIInfo(TraduireMemoire('Cette pièce a été exportée (COMSX).'),'ATTENTION');
          break;
        end;
  // mis en commentaire par BRL car pose pb pour les factures passées en compta comme validée par défaut
  //      if TOBECR.Detail[Indice].getValue('E_VALIDE')='X' then BEGIN result := false; Break; END;
      end;
    end;
    Ferme (Q);
  FINALLY
  	TOBECR.free;
  END;
end;

function TraitementVOuvrage (TOBEcr,TOBPieceinterv,TOBAFFInterv,TOBPiece,TOBL,TOBOuvrages,TOBOuvragesP,TOBTiers,TOBArticles,TOBCpta,TOBPorcs: TOB;MM : RMVT ; NbDec : integer; EnHt : boolean; DEV : Rdevise):  T_RetCpta;
var Indice : integer;
    TOBPlat,TOBPP,TOBA : TOB;
    refUnique : string;
begin
  result := rcOk;
  TOBPlat := FindOuvragesPlat (TOBL,TOBOuvragesP);
  if TOBPlat = nil then
  begin
    TOBPlat := AddMereLignePlat (TOBOuvragesP,TOBL.GetValue('GL_NUMORDRE'));
    GetOuvragePlat (TOBpiece,TOBL,TOBOuvrages,TOBPlat,TOBTiers,DEV,(TOBPiece.getvalue('GP_NATUREPIECEG')<>'DBT'));
  end;
  //
  if TOBPlat = nil then BEGIN result := rcRef ; exit; END;
  TOBL.clearDetail; // on vire pour enfin recuperer les ventil du detail d'ouvrage
  for Indice := 0 to TOBPlat.detail.count -1 do
  begin
    TOBPP := TOBplat.detail[Indice];
    refUnique := TOBPP.GetValue('BOP_ARTICLE');
    TOBA := TOBarticles.findFirst(['GA_ARTICLE'],[refUnique],true);
    if TOBA=Nil then Continue ;
    if TOBPP.detail.count < 2 then TOB.Create ('_1',TOBPP,-1);
    if TOBPP.detail.count < 2 then TOB.Create ('_2',TOBPP,-1);
    result := TraitementVLigne (TOBEcr,TOBAFFInterv,TOBPieceinterv,TOBA,TOBpiece,TOBPP,TOBTiers,TOBARticles,TOBCpta,TOBPorcs,MM,Nbdec,EnHt,DEV,TOBL);
    if result <> rcOk then break;
      // la ventilation analytique est stocké sur la ligne de document et pas dans les details d'ouvrages
    CumulDetOuvAnalSurLigne(TOBL,TOBPP);
    TOBPP.clearDetail;
  end;
end;

function TraitementVLigne (TOBEcr,TOBAFFInterv,TOBPieceinterv,TOBA,TOBPiece,TOBL,TOBTiers,TOBArticles,TOBCpta,TOBPorcs : TOB ; MM : RMVT ; NbDec : integer; EnHt : boolean; DEV : Rdevise; TOBMereLigne : TOb=nil ) : T_RetCpta;
Var i,k,itrouv,kk : integer ;
    TOBCode,TOBP : TOB ;
    RefUnique            : String ;
    CptHT,LibArt,Affaire : String ;
    CumHT                : T_CodeCpta ;
    OkTX                 : Boolean ;
    XD,XP : Double ;
    XDT : T_DetAnaL ;
    R_ArtFi : T_MontantArtFi ;
    // Modif BTP
    FamTaxeHt : string;
    {Ratio : double;} (* provenance de tva au millieme *)
    Indice : integer;
    Prefixe : string;
    NatureTravail : double;
    LocAnaP,LocAnaS : TOb;
    Fournisseur : string;
    // --
begin
  Result:=rcOk ;
  prefixe := GetPrefixeTable (TOBL);
  LibArt:=Copy(TOBL.GetValue(prefixe+'_LIBELLE'),1,35) ;
  Affaire:=TOBL.GetValue(prefixe+'_AFFAIRE') ; CptHT:='' ;
  NatureTravail := 0;
  Fournisseur := '';
  //
  if (VH_GC.BTCODESPECIF = '001') and (IsLigneSoustraite (TOBL)) then Exit;
  //
  if TOBL.GetValue('BLF_NATURETRAVAIL')<>'' then
  begin
  	NatureTravail := valeur(TOBL.GetValue('BLF_NATURETRAVAIL'));
  end;
  if NatureTravail = 0 then
  begin
    if prefixe = 'GL' then
    begin
  		NatureTravail := valeur(TOBL.GetValue('GLC_NATURETRAVAIL'));
    end else
    begin
  		NatureTravail := valeur(TOBL.GetValue(prefixe+'_NATURETRAVAIL'));
    end;
  end;

  if (NatureTravail <> 0) and (TOBL.GetString('BLF_FOURNISSEUR')<>'') then
  begin
  	Fournisseur := TOBL.GetValue('BLF_FOURNISSEUR');
  end;
  if (NatureTravail <> 0) and (Fournisseur = '') then
  begin
  	Fournisseur := TOBL.GetValue(prefixe+'_FOURNISSEUR');
  end;

  if (NatureTravail > 0) and ( NatureTravail < 10 ) and
     ((Fournisseur<>'') and (not ComptabiliseTrait (Fournisseur,TOBAFFInterv,TOBPieceinterv)) ) then
  begin
    exit;
  end;
  {Sommations par compte et familles de taxes identiques}
  RecupLesDC(TOBL,XD,XP) ; // totalisation de la ligne
  for Indice := 0 to VH_GC.TOBParamTaxe.detail.count -1 do
  begin
    if VH_GC.TOBParamTaxe.detail[Indice].getValue('BPT_TYPETAXE') <> 'TVA' Then continue;
    if TOBL.getValue(prefixe+'_FAMILLETAXE'+intToStr(Indice+1))='' then continue;
    // Positionnement en vue de la gestion du compte en fonction de famille de taxe
    TOBCode:=ChargeAjouteComptaBTP(TOBCpta,TOBPiece,TOBL,TOBA,TOBTiers,Nil,Nil,TOBL.GetValue(prefixe+'_FAMILLETAXE'+IntToStr(Indice+1)),True) ;
    if TOBCode<>Nil then
    BEGIN
      if TOBMereLigne <> nil then
      begin
    		PreVentileLigneDetailOuv(TOBCode,TOBL.detail[0],TOBL.detail[1],TOBL,TOBMereLigne);
      end else
      begin
        if (TOBL.Detail[0].Detail.Count <= 0) then LocAnaP := TOBL.detail[0] else LocAnaP := nil;
        if (TOBL.Detail[1].Detail.Count <= 0) then LocAnaS := TOBL.detail[1] else LocAnaS := nil;
        PreVentileLigne(TOBCode, LocAnaP, LocAnaS, TOBL);
      end;
      if ((MM.Nature='FC') or (MM.Nature='AC')) then CptHT:=TOBCode.GetValue('GCP_CPTEGENEVTE')
                                                else CptHT:=TOBCode.GetValue('GCP_CPTEGENEACH') ;
    END ;
    if (TOBCode=Nil) or ((CptHT='') and (VH_GC.GCPontComptable='ATT')) then
    BEGIN
      if ((MM.Nature='FC') or (MM.Nature='AC')) then CptHT:=VH_GC.GCCpteHTVTE else CptHT:=VH_GC.GCCpteHTACH ;
    END ;
    // Erreur sur le compte HT article
    if CptHT='' then BEGIN Result:=rcPar ; LastMsg:=5 ; Exit ; END ;
    iTrouv:=-1 ;
    //
    if GereTvaMixte then CumulTVASurEncaissement(CptHT, TobL);
    //
    for k:=0 to TTA.Count-1 do
    BEGIN
      CumHT:=T_CodeCpta(TTA[k]) ;
      //
      if CumHT.CptHT<>CptHT then Continue ;
      if DistinctAffaire then BEGIN if CumHT.Affaire<>Affaire then Continue ; END ;
      OkTx:=True ;
      if CumHT.famTaxe[1]<> TOBL.GetValue(prefixe+'_FAMILLETAXE'+IntToStr(Indice+1)) then OkTx := false;
      //          for kk:=1 to 5 do if CumHT.FamTaxe[kk]<>TOBL.GetValue('GL_FAMILLETAXE'+IntToStr(kk)) then OkTx:=False ;
      if OkTX then BEGIN iTrouv:=k ; Break ; END ;
    END ;
    if iTrouv<0 then
    BEGIN
      CumHT:=T_CodeCpta.Create ; CumHT.CptHT:=CptHT ; CumHT.LibU:=True ;
      CumHT.LibArt:=LibArt ; CumHT.Affaire:=Affaire ;
      //         for kk:=1 to 5 do CumHT.FamTaxe[kk]:=TOBL.GetValue('GL_FAMILLETAXE'+IntToStr(kk)) ;
      CumHt.FamTaxe[1] := TOBL.GetValue(prefixe+'_FAMILLETAXE'+IntToStr(Indice+1));
    END else
    BEGIN
      CumHT:=T_CodeCpta(TTA[iTrouv]) ;
    END ;
    if (not TOBL.FieldExists('MILLIEME'+intToStr(Indice+1))) or (TOBL.GetValue('MILLIEME'+InttoStr(Indice+1)) = 0) then
    begin
      CumHT.XD:=CumHT.XD+XD ; CumHT.XP:=CumHT.XP+XP ;
    end else
    begin
      CumHT.XD:=CumHT.XD+ (XD * TOBL.GetValue('MILLIEME'+InttoStr(Indice+1)) / 1000) ;
      CumHT.XP:=CumHT.XP+(XP* TOBL.GetValue('MILLIEME'+IntToStr(Indice+1)) / 1000)  ;
    end;
    CumHT.Qte:=CumHT.Qte+TOBL.GetValue(prefixe+'_QTEFACT') ;
    if CumHT.LibArt<>LibArt then BEGIN CumHT.LibU:=False ; CumHT.LibArt:='' ; END ;
    //
    CumHT.SommeTaxeD[Indice+1]:=CumHT.SommeTaxeD[Indice+1]+TOBL.GetValue(prefixe+'_TOTALTAXEDEV'+IntToStr(Indice+1)) ;
    CumHT.SommeTaxeP[Indice+1]:=CumHT.SommeTaxeP[Indice+1]+TOBL.GetValue(prefixe+'_TOTALTAXE'+IntToStr(Indice+1)) ;
    //
    CumulAnal(CumHT,TOBL,XD,XP) ;
    if iTrouv<0 then TTA.Add(CumHT) ;
  END ;
end;

function GetWSDocumentInfFromRmvt(lRmvt: RMVT): T_WSDocumentInf;
begin
  Result.dType   := lRmvt.DocType;
  Result.dStub   := lRmvt.Souche;
  Result.dNumber := lRmvt.DocNumber;
  Result.dIndex  := 0;
end;

function GetWSDocumentInfFromTob(TobGP: TOB): T_WSDocumentInf;
begin
  Result.dType   := TobGP.GetString('GP_NATUREPIECEG');
  Result.dStub   := TobGP.GetString('GP_SOUCHE');
  Result.dNumber := TobGP.GetInteger('GP_NUMERO');
  Result.dIndex  := TobGP.GetInteger('GP_INDICEG');
end;

end.
