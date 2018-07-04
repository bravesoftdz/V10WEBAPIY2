unit MulSMPUtil;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics,
  Controls, Forms, Dialogs, UiUtil, ExtCtrls,
  ComCtrls, Buttons, StdCtrls, Mask, Menus,
  Grids, ColMemo, HRichEdt, HSysMenu,
  Hctrls, hmsgbox, Hqry,  HTB97, HPanel, HRichOLE,
  Hcompte, Hent1, Ent1,
{$IFDEF EAGLCLIENT}
  eMul,
  MaineAGL,
{$ELSE}
  Mul,
  DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DBGrids,
  HDB,
  FE_MAIN,
{$ENDIF}
  Saisie,
  SaisUtil,
  SaisComm,
  EncUtil,
  UTOB,
  UtilPGI,
  ParamSoc,
  uLibAnalytique,
  uLibEcriture,
  UtilSais,
  GenerMP       // pour TGenereMP
  ;

Const
  MAXSELECTECHE = 4999 ;

Function  AttribSL ( smp : TSuiviMP ) : TSorteLettre ;
Function  AttribListe ( smp : TSuiviMP ; Suivi : String = '' ) : String ;
Function  QuelProfil(smp : TSuiviMP ; Var OkOk : Boolean) : tProfilTraitement ;
Function  ProfilOk( smp : TSuiviMP ) : Boolean ;
Function  IsEnc(smp : TSuiviMP) : Boolean ;
Function  IsCompensation(smp : TSuiviMP) : Boolean ;  {FP 21/02/2006}
Function  ModifRibSurMul(Q : TQuery ; TID : Boolean ; Update : Boolean) : Boolean ;
Procedure CreerCodeLot ( CodeLot : String ; PourLot : Boolean = FALSE ; Abrege : String = '' ; Libre : String = '') ;
Function  AttribPlus ( smp : TSuiviMP ) : String ;
Function  AttribSoucheLotPreparation ( smp : TSuiviMP ) : String ;
Function  EstMultiCatMP(smp : TSuiviMP) : Boolean ;
procedure LaCATEGORIEChange(HCat : THValCombobox ; XX_WHEREMP : THEdit ; SorteLettre : TSorteLettre) ;
Function  TrouveJalCptEff(smp : tSuiviMP) : Integer ;
Function  AttribHelp ( smp : TSuiviMP ; Suivi : Boolean = FALSE) : Integer ;
{$IFDEF EAGLCLIENT}
Procedure SwapModeGrid(GS : THGrid ; Q : THQuery) ;
{$ELSE}
Procedure SwapModeGrid(GS : thDBGrid ; Q : thQuery) ;
{$ENDIF}
Function  SmpToStr ( smp : TSuiviMP ) : String ;
Function  StrToSmp ( vStsmp : String ) : TSuiviMP ;
function  IsVirInternational(Smp : TSuiviMP) : Boolean;{JP 15/05/07 : FQ 17309}

// Nouvelles fonctions importées de MulGenereMP
function  GenereEscompteMP ( TOBEcr,TOBEscomptes,TOBGHT,TOBGTVA : TOB ; vAvecTVA : Boolean ; JalEsc : String ; ParmsMP : TGenereMP ) : TOB ;
procedure InitLigneEsc ( TOBL : TOB ; MMEsc : RMVT ; DeltaJour : Integer ; ParmsMP : TGenereMP ) ;
procedure FinirLignesEscompte ( T1,T2,T3,TOBGHT,TOBGTVA : TOB ; CompteHT,CompteTVA : String ; TauxEsc,TauxTva,NewTaux: double ; vAvecTVA : Boolean ; ParmsMP : TGenereMP ) ;
procedure TraiteTOBEscompte ( TOBL, TOBEscomptes, TOBGHT, TOBGTVA : TOB ; vBoAuProrata : Boolean ;
                              vStMeth : String ; ParmsMP : TGenereMP ; var PasPris : Boolean ) ;
procedure AjouteComptePourEsc ( Cpt : String ; TOBRech : TOB ) ;
function  CalcNewTauxEscompte( TauxEsc, X : Double ; TOBL : TOB ;
                               vBoAuProrata : Boolean ; vStMethEsc : String ; ParmsMP : TGenereMP ;
                               Var Delta : Integer ; var TauxEscF : Double ; Var PasPris : Boolean ) : Double ;
procedure InitCommuns ( TOBD : TOB ) ;
procedure TripoteDetail ( TOBD : TOB ; ParmsMP : TGenereMP ) ;
procedure AffecteG ( TOBL : TOB ; ParmsMP : TGenereMP ) ;
Function  GetLeMontant ( Montant,Couv,Sais : Double ; tsm : TSorteMontant ) : Double ;
procedure FinirDestination ( TOBDest, TOBR : TOB ; ParmsMP : TGenereMP ) ;
procedure FinirAnas ( TOBL : TOB ; ParmsMP : TGenereMP ) ;
Function  VerifCrits ( TOBL : TOB ; ValsRupt : String ; ParmsMP : TGenereMP ) : boolean ;
procedure DupliquerTOBsmp (  sBqe, sCat, ValsRupt : String ; TOBOrig, TOBDest : TOB ; ParmsMP : TGenereMP ) ;
Procedure RecalcLigne ( TOBL : TOB ; ParmsMP : TGenereMP ) ;
procedure PassageDevise( TOBDest : TOB ; ParmsMP : TGenereMP ) ;
procedure ConstitueDest ( sBQE, sCat, ValsRupt : String ; TOBOrig, TOBDest : TOB ; ParmsMP : TGenereMP ) ;
function  CoherAna ( ideb, ifin : integer ; TOBDest : TOB ; ParmsMP : TGenereMP ) : boolean ;
procedure AjouteLesAnas( TOBOrig : TOB ) ;
procedure VireInutiles ( TOBOrig : TOB; swapSelect : Boolean = False);
procedure AnasRupt ( TOBL : TOB ; ParmsMP : TGenereMP ) ;
procedure RegroupeAnas ( TOBL : TOB ) ;
procedure InitEscompteSup ( TOBE : TOB ) ;
Function  FindCreerTOBEscomptePourLigne ( TOBEscomptes, TOBLigne : TOB ; QueFind : boolean ) : TOB ;
Function ExtractDiscerne ( TOBOrig : TOB ; sBqe, sCat, ValsRupt : String ; ParmsMP : TGenereMP ) : TOB ;
Procedure VerifRep(LaSousDir : String) ;
Procedure CompleteMMRefCCMP( ParmsMP : TGenereMP ; Var FR : tFormuleRefCCMP ) ;
Procedure CompleteMM ( ParmsMP : TGenereMP ; SorteLettre : TSorteLettre ;
                       vBoSpooler, vBoFichier : Boolean ; vStRepSpooler : String ;
                       Var MM : RMVT ) ;
Function  CompteToJal ( sBqe : String ; var OkJal : Boolean ) : String ;
Procedure FlagOrigEsc ( TOBOrig : TOB ) ;
// New
procedure GlobaliserDest( TOBDest : TOB ; ParmsMP : TGenereMP ) ;

// Repot + correction SaisUtil // Ajout SBO 03/08/2004
Function  NormaliserPieceSimuTOB ( TOBL : TOB ; MajTOBL : Boolean ) : Boolean ; // Ajout SBO : Update piece simu --> pièce "normal"

Function  GetYEnATropMsg( smp : TSuiviMP ) : string ;

implementation

Uses
  {$IFDEF MODENT1}
  CPObjetGen,
  CPTypeCons,
  CPProcGen,
  CPVersion,
  {$ENDIF MODENT1}
  FichComm;

Function AttribSL ( smp : TSuiviMP ) : TSorteLettre ;
BEGIN
Result:=tslAucun ;
Case smp of
   smpEncTraEdt,  smpEncTraEdtNC   : Result:=tslTraite ;
   smpDecChqEdt,  smpDecChqEdtNC   : Result:=tslCheque ;
   smpDecVirEdt,  smpDecVirEdtNC   : Result:=tslVir ;
   smpDecVirInEdt,smpDecVirInEdtNC : Result:=tslVir ;
   smpEncPreEdt,smpEncPreEdtNC     : Result:=tslPre ;
   smpDecBorEdt,  smpDecBorEdtNC   : Result:=tslBOR ;
   END ;
END ;

Function AttribListe ( smp : TSuiviMP ; Suivi : String = '' ) : String ;
BEGIN
Case smp of
   smpEncPreBqe     : Result:='CPENCPRELEV' ;
   smpEncTraEdt     : Result:='CPENCTRAITEEDT' ;
   smpEncTraPor     : Result:='CPENCTRAITEPOR' ;
   smpEncChqPor     : Result:='CPENCCHQPOR' ;
   smpEncCBPor      : Result:='CPENCCBPOR' ;
   smpEncTraEdtNC   : Result:='CPENCTRAIEDTNC' ;
   smpEncTraEnc     : Result:='CPENCTRAITEENC' ;
   smpEncTraEsc     : Result:='CPENCTRAITEESC' ;
   smpEncChqBqe     : Result:='CPENCCHQBQE' ;
   smpEncCBBqe      : Result:='CPENCCBBQE' ;
   smpEncTraBqe     : Result:='CPENCTRAITEBQE' ;
   smpEncPreEdt     : Result:='CPENCPREEDT' ;
   smpEncPreEdtNC   : Result:='CPENCPREEDTNC' ;
   smpEncDiv        : Result:='CPENCDIV' ;
   smpEncTous       : Result:='CPENCTOUS' ;
   smpDecChqEdt     : Result:='CPDECCHEQUEEDT' ;
   smpDecChqEdtNC   : Result:='CPDECCHQEDTNC' ;
   smpDecVirBqe     : Result:='CPDECVIREMENT' ;
   smpDecVirEdt     : Result:='CPDECVIREDT' ;
   smpDecVirEdtNC   : Result:='CPDECVIREDT' ; {A priori jamais utilisé}
   smpDecVirInBqe   : Result:='CPDECVIRIN' ;
   smpDecVirInEdt   : Result:='CPDECVIRINEDT' ;
   smpDecVirInEdtNC : Result:='CPDECVIRINEDNC' ;
   smpDecBorEdt     : Result:='CPDECBOREDT' ;
   smpDecBorEdtNC   : Result:='CPDECBOREDTNC' ;
   smpDecTraPor     : Result:='CPDECTRAITEPOR' ;
   smpDecBorDec     : Result:='CPDECTRAITEENC' ;
   smpDecBorEsc     : Result:='CPDECTRAITEESC' ;
   smpDecTraBqe     : Result:='CPDECTRAITEBQE' ;
   smpDecDiv        : Result:='CPDECDIV' ;
   smpDecTous       : Result:='CPDECTOUS' ;
   {b FP 21/02/2006}
   smpCompenCli,
//   smpCompenFou     : Result:='CPCOMPENSA0'+IntToStr(TCompensation.GetNumPlanCorrespondance);
   smpCompenFou     : if V_PGI.Sav then Result:='CPCOMPENSA0'+IntToStr(TCompensation.GetNumPlanCorrespondance)
                    else Result:='CPCOMPEN0'+IntToStr(TCompensation.GetNumPlanCorrespondance);
   {e FP 21/02/2006}
   else Result:='CPENCPRELEV' ;
   END ;
Result:=Result+Suivi ;
END ;

{JP 15/05/07 : FQ 17309 : Pour savoir si on est sur un mul pour virements internationaux
{---------------------------------------------------------------------------------------}
function  IsVirInternational(Smp : TSuiviMP) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := Smp in [smpDecVirInBqe, smpDecVirInEdt, smpDecVirInEdtNC];
end;

Function AttribPlus ( smp : TSuiviMP ) : String ;
BEGIN
Case smp of
   smpEncPreBqe     : Result:=' AND YX_ABREGE="CPENCPRELEV"' ; // FQ 18742 : SBO 05/09/2006
   smpEncTraEdt     : Result:=' AND YX_ABREGE="CPENCTRAITEEDT"' ;
   smpEncTraPor     : Result:=' AND YX_ABREGE="CPENCTRAITEPOR"' ;
   smpEncChqPor     : Result:=' AND YX_ABREGE="CPENCCHQPOR"' ;
   smpEncCBPor      : Result:=' AND YX_ABREGE="CPENCCBPOR"' ;
   smpEncTraEdtNC   : Result:=' AND YX_ABREGE="CPENCTRAITEEDT"' ;
   smpEncTraEnc     : Result:=' AND YX_ABREGE="CPENCTRAITEENC"' ;
   smpEncTraEsc     : Result:=' AND YX_ABREGE="CPENCTRAITEESC"' ;
   smpEncChqBqe     : Result:=' AND YX_ABREGE="CPENCCHQBQE"' ;
   smpEncCBBqe      : Result:=' AND YX_ABREGE="CPENCCBBQE"' ;
   smpEncTraBqe     : Result:=' AND YX_ABREGE="CPENCTRAITEBQE"' ;
   smpEncPreEdt     : Result:=' AND YX_ABREGE="CPENCPREEDT"' ;
   smpEncPreEdtNC   : Result:=' AND YX_ABREGE="CPENCPREEDT"' ;
   smpEncDiv        : Result:=' AND YX_ABREGE="CPENCDIV"' ;
   smpEncTous       : Result:=' AND YX_ABREGE="CPENCTOUS"' ;
   smpDecChqEdt     : Result:=' AND YX_ABREGE="CPDECCHEQUEEDT"' ;
   smpDecChqEdtNC   : Result:=' AND YX_ABREGE="CPDECCHEQUEEDT"' ;
   smpDecVirBqe     : Result:=' AND YX_ABREGE="CPDECVIREMENT"' ;
   smpDecVirEdt     : Result:=' AND YX_ABREGE="CPDECVIREDT"' ;
   smpDecVirEdtNC   : Result:=' AND YX_ABREGE="CPDECVIREDT"' ;
   smpDecVirInBqe   : Result:=' AND YX_ABREGE="CPDECVIRIN"' ;
   smpDecVirInEdt   : Result:=' AND YX_ABREGE="CPDECVIRINEDT"' ;
   smpDecVirInEdtNC : Result:=' AND YX_ABREGE="CPDECVIRINEDT"' ;
   smpDecBorEdt     : Result:=' AND YX_ABREGE="CPDECBOREDT"' ;
   smpDecBorEdtNC   : Result:=' AND YX_ABREGE="CPDECBOREDT"' ;
   smpDecTraPor     : Result:=' AND YX_ABREGE="CPDECTRAITEPOR"' ;
   smpDecBorDec     : Result:=' AND YX_ABREGE="CPDECTRAITEENC"' ;
   smpDecBorEsc     : Result:=' AND YX_ABREGE="CPDECTRAITEESC"' ;
   smpDecTraBqe     : Result:=' AND YX_ABREGE="CPDECTRAITEBQE"' ;
   smpDecDiv        : Result:=' AND YX_ABREGE="CPDECDIV"' ;
   smpDecTous       : Result:=' AND YX_ABREGE="CPDECTOUS"' ;
   {b FP 21/02/2006}
   smpCompenCli     : Result:=' AND YX_ABREGE="CPCOMPENSCLI"';
   smpCompenFou     : Result:=' AND YX_ABREGE="CPCOMPENSFOU"';
   {e FP 21/02/2006}
   else Result:='' ;
   END ;
END ;

Function AttribSoucheLotPreparation ( smp : TSuiviMP ) : String ;
BEGIN
Case smp of
   smpEncPreBqe     : Result:='CGP' ;
   smpEncTraEdt     : Result:='CLT' ;
   smpEncTraPor     : Result:='CPT' ;
   smpEncChqPor     : Result:='CPC' ;
   smpEncCBPor      : Result:='CPB' ;
   smpEncTraEdtNC   : Result:='CLT' ;
   smpEncTraEnc     : Result:='CET' ;
   smpEncTraEsc     : Result:='CST' ;
   smpEncChqBqe     : Result:='CEC' ;
   smpEncCBBqe      : Result:='CEB' ;
   smpEncTraBqe     : Result:='CRT' ;
   smpEncPreEdt     : Result:='CLP' ;
   smpEncDiv        : Result:='CDV' ;
   smpEncTous       : Result:='ENCTOUS' ;
   smpDecChqEdt     : Result:='FLC' ;
   smpDecChqEdtNC   : Result:='FLC' ;
   smpDecVirBqe     : Result:='FGV' ;
   smpDecVirEdt     : Result:='FLV' ;
   smpDecVirEdtNC   : Result:='FLV' ;
   smpDecVirInBqe   : Result:='FHV' ;
   smpDecVirInEdt   : Result:='FIV' ;
   smpDecVirInEdtNC : Result:='FIV' ;
   smpDecBorEdt     : Result:='FLB' ;
   smpDecBorEdtNC   : Result:='FLB' ;
   smpDecTraPor     : Result:='FGB' ;
   smpDecBorDec     : Result:='FRB' ;
   smpDecBorEsc     : Result:='BORESC' ;
   smpDecTraBqe     : Result:='FRB' ;
   smpDecDiv        : Result:='FDV' ;
   smpDecTous       : Result:='DECTOUS' ;
   {b FP 21/02/2006}
   smpCompenCli     : Result:='CCP' ;
   smpCompenFou     : Result:='FCP' ;
   {e FP 21/02/2006}
   else Result:='' ;
   END ;
END ;

Function QuelProfil(smp : TSuiviMP ; Var OkOk : Boolean) : tProfilTraitement ;

BEGIN
OkOk:=TRUE ; Result:=prAucun ;
Case smp Of
  smpEncPreBqe,smpEncTraEdt,smpEncTraEnc,smpEncTraEsc,smpEncTraBqe,
  smpEncTous,smpEncTraEdtNC,smpEncTraPor,smpEncPreEdt,smpEncPreEdtNC,smpEncDiv,smpEncChqPor,
  smpEncCBPor,smpEncChqBqe,smpEncCBBqe : Result:=prClient ;

  smpDecVirBqe,smpDecVirEdt,smpDecVirEdtNC,
  smpDecVirInBqe,smpDecVirInEdt,smpDecVirInEdtNC,
  smpDecChqEdt,smpDecChqEdtNC,
  smpDecBorEdt,smpDecBorEdtNC,smpDecBorDec,smpDecBorEsc,
  smpDecTraBqe,smpDecTraPor,smpDecDiv,smpDecTous : Result:=prFournisseur ;
  Else OkOk:=FALSE ;
  END ;
END ;

Function ProfilOk( smp : TSuiviMP ) : Boolean ;
Var ind : tProfilTraitement ;
    OkOk : Boolean ;
BEGIN
Result:=FALSE ;
Ind:=QuelProfil(smp,OkOk) ; If Not OkOk Then Exit ;
If Not VH^.ProfilCCMPExiste[Ind] Then BEGIN Result:=TRUE ; Exit ; END ;
If Not VH^.ProfilUserC[Ind].ProfilOk Then Exit ;
Result:=TRUE ;
END ;

{b FP 21/02/2006}
Function  IsCompensation(smp : TSuiviMP) : Boolean ;
begin
  Result := (Smp in [smpCompenCli, smpCompenFou]);
end;
{e FP 21/02/2006}

Function IsEnc(smp : TSuiviMP) : Boolean ;
BEGIN
Result:=True;
  Case Smp Of
   smpEncDiv    : Result:=TRUE  ; // Décaissements divers
   smpEncPreEdt : Result:=TRUE  ; // Lettres-Prélèvements fournisseurs
   smpEncPreEdtNC : Result:=TRUE  ; // Lettres-Prélèvements fournisseurs
   smpEncPreBqe : Result:=TRUE  ; // Prélèvements clients
   smpEncChqPor : Result:=TRUE  ; // Mise en portefeuille chèques clients
   smpEncCBPor  : Result:=TRUE  ; // Mise en portefeuille Cartes bleues clients
   smpEncTraPor : Result:=TRUE  ; // Mise en portefeuille traite clients
   smpEncTraEdt : Result:=TRUE  ; // Lettres-traite clients
 smpEncTraEdtNC : Result:=TRUE  ; // Lettres-traite clients
   smpEncTraEnc : Result:=TRUE  ; // Traite clients à l'encaissement
   smpEncTraEsc : Result:=TRUE  ; // Traite clients à l'escompte
   smpEncChqBqe : Result:=TRUE  ; // Chèques clients en banque
   smpEncCBBqe  : Result:=TRUE  ; // Cartes bleues clients en banque
   smpEncTraBqe : Result:=TRUE  ; // Traite clients en banque
     smpEncTous : Result:=TRUE  ; //
   smpDecChqEdt : Result:=FALSE ; // Lettres-chèques fournisseurs
 smpDecChqEdtNC : Result:=FALSE ; // Lettres-chèques fournisseurs
   smpDecVirEdt : Result:=FALSE ; // Virements fournisseurs
 smpDecVirEdtNC : Result:=FALSE ; // Virements fournisseurs
   smpDecVirBqe : Result:=FALSE ; // Virements fournisseurs
 smpDecVirInEdt : Result:=FALSE ; // Virements internationaux fournisseurs
smpDecVirInEdtNC: Result:=FALSE ; // Virements internationaux fournisseurs
 smpDecVirInBqe : Result:=FALSE ; // Virements internationaux fournisseurs
   smpDecBorEdt : Result:=FALSE ; // Lettres-Bor fournisseurs
 smpDecBorEdtNC : Result:=FALSE ; // Lettres-Bor fournisseurs
   smpDecBorDec : Result:=FALSE ; // Traites fournisseurs en encaissement
   smpDecBorEsc : Result:=FALSE ; // Traites fournisseurs à l'escompte
   smpDecTraBqe : Result:=FALSE ; // Traites fournisseurs en banque
   smpDecTraPor : Result:=FALSE ; // Mise en portefeuille traite clients
   smpDecDiv    : Result:=FALSE ; // Décaissements divers
     smpDecTous : Result:=FALSE ; //
{b FP 21/02/2006}
   smpCompenCli : Result:=FALSE ;
   smpCompenFou : Result:=FALSE ;
{e FP 21/02/2006}
    END ;
END ;

Function ModifRIBMP ( Var RIB : String ; Aux : String ; IsAux : Boolean) : boolean ;
Var Auxi   : String ;
    Num{,Nb} : integer ;
    QR      : TQuery ;
    {SQL,}Etab,Guichet,NumCompte,Cle,Dom : String ;
//    M                                  : RMVT ;
BEGIN
Result:=False ; Auxi:=Aux ;
Num:=FicheRIB_AGL(Auxi,taModif,True,RIB,IsAux) ; if Num<=0 then Exit ;
Result:=TRUE ;
QR:=OpenSQL('Select * from RIB Where R_AUXILIAIRE="'+Auxi+'" AND R_NUMERORIB='+IntToStr(Num),True) ;
if Not QR.EOF then
   BEGIN
   Etab:=QR.FindField('R_ETABBQ').AsString ;
   Guichet:=QR.FindField('R_GUICHET').AsString ;
   NumCompte:=QR.FindField('R_NUMEROCOMPTE').AsString ;
   Cle:=QR.FindField('R_CLERIB').AsString ;
   Dom:=QR.FindField('R_DOMICILIATION').AsString ;
   END ;
Ferme(QR) ;
RIB:=EncodeRIB(Etab,Guichet,NumCompte,Cle,Dom) ;
(*
if Maj then
   BEGIN
   M:=OBMToIdent(O,True) ;
   SQL:='Update ECRITURE Set E_RIB="'+RIB+'" Where '+WhereEcriture(tsGene,M,True);
   Nb:=ExecuteSQL(SQL) ; if Nb=1 then Result:=True ;
   END ;
*)
END ;

Function ModifRibSurMul(Q : TQuery ; TID : Boolean ; Update : Boolean) : Boolean ;
Var //OkModif : boolean ;
    RIB,Aux,OldRib{,Gen} : String ;
    RJal,RExo : String ;
    RDate : TDateTime ;
    RNumP,RNumL,RNumEche : Integer ;
    OkPourUpdate : Boolean ;
begin
Result:=FALSE ;
If Q=NIL Then Exit ;
{$IFDEF EAGLCLIENT}
{$ELSE}
  If Not Q.Active Then Exit ;
{$ENDIF}
if Q.EOF Or Q.BOF Then Exit ;
If Q.FindField('E_RIB')=NIL Then
  BEGIN
  HShowMessage('0;?Caption?;Fonction non disponible : Il faut paramétrer le RIB dans la liste de restitution;W;O;O;O;','','') ;
  Exit ;
  END ;
If Not TID Then If Q.FindField('E_AUXILIAIRE')=NIL Then
  BEGIN
  HShowMessage('0;?Caption?;Fonction non disponible : Il faut paramétrer le code tiers dans la liste de restitution;W;O;O;O;','','') ;
  Exit ;
  END ;
If TID Then If Q.FindField('E_GENERAL')=NIL Then
  BEGIN
  HShowMessage('0;?Caption?;Fonction non disponible : Il faut paramétrer le compte général dans la liste de restitution;W;O;O;O;','','') ;
  Exit ;
  END ;
RIB:=Q.FindField('E_RIB').AsString ; OldRib:=Rib ;
If TID Then Aux:=Q.FindField('E_GENERAL').AsString Else Aux:=Q.FindField('E_AUXILIAIRE').AsString ;
ModifRIBMP(RIB,AUX,Not TID) ;
//OkModif:=FALSE ;
OkPourUpdate:=(Q.FindField('E_JOURNAL')<>NIL) And (Q.FindField('E_DATECOMPTABLE')<>NIL) And
              (Q.FindField('E_NUMEROPIECE')<>NIL) And (Q.FindField('E_NUMLIGNE')<>NIL) And
              (Q.FindField('E_NUMECHE')<>NIL) ;
If (OldRib<>Rib) And Update And OkPourUpdate Then
  BEGIN
   Try
   RJal:=Q.FindField('E_JOURNAL').AsString ;
   RExo:=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) ;
   RDate:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
   RNumP:=Q.FindField('E_NUMEROPIECE').AsInteger ;
   RNumL:=Q.FindField('E_NUMLIGNE').AsInteger ;
   RNumEche:=Q.FindField('E_NUMECHE').AsInteger ;
   BeginTrans ;
   {JP 16/11/07 : FQ 21847 : Gestion de E_UTILISATEUR}
   ExecuteSQL('UPDATE Ecriture SET E_RIB = "' + RIB + '", E_UTILISATEUR = "' + V_PGI.User +
               '", E_DATEMODIF = "' + UsTime(NowH) + '" WHERE E_JOURNAL="'+RJal+'"'
              +' AND E_EXERCICE="'+RExo+'"'
              +' AND E_DATECOMPTABLE="'+USDATETIME(RDate)+'"'
              +' AND E_QUALIFPIECE="N"'
              +' AND E_NUMEROPIECE='+IntToStr(RNumP)
              +' AND E_NUMLIGNE='+IntToStr(RNumL)
              +' AND E_NUMECHE='+IntToStr(RNumEche)) ;
   CommitTrans ;
   Result:=TRUE ;
   Except
   Rollback ;
   End ;
  END ;
end;

Procedure CreerCodeLot ( CodeLot : String ; PourLot : Boolean = FALSE ; Abrege : String = '' ; Libre : String = '') ;
var STyp : String ;
  {$IFDEF EAGLCLIENT}
    newLot : TOB ;
  {$ELSE}
    Q : TQuery ;
  {$ENDIF}
begin
  if CodeLot='' then Exit ;
  if PourLot
    then STyp:='CNL'
    else STyp:='CBE' ;
{$IFDEF EAGLCLIENT}
  if not ExisteSQL('SELECT * FROM CHOIXEXT WHERE YX_TYPE="'+STyp+'" AND YX_CODE="'+CodeLot+'"') then
    begin
    newLot := TOB.Create('CHOIXEXT', nil, -1) ;
    newLot.InitValeurs ;
    newLot.putValue('YX_TYPE',    STyp);
    newLot.putValue('YX_CODE',    CodeLot);
    newLot.putValue('YX_LIBELLE', RechDom('ttUtilisateur',V_PGI.User,False)+' le '+DateToStr(Date) );
    if PourLot then
      begin
      newLot.putValue('YX_ABREGE', Abrege);
      newLot.putValue('YX_LIBRE',  Libre);
      end ;
    newLot.InsertDB(nil) ;
    newLot.Free;
    end ;
{$ELSE}
  Q:=OpenSQL('SELECT * FROM CHOIXEXT WHERE YX_TYPE="'+STyp+'" AND YX_CODE="'+CodeLot+'"',False) ;
  if Q.EOF then
    BEGIN
    Q.Insert ; InitNew(Q) ;
    Q.FindField('YX_TYPE').AsString:=STyp ;
    Q.FindField('YX_CODE').AsString:=CodeLot ;
    Q.FindField('YX_LIBELLE').AsString:=RechDom('ttUtilisateur',V_PGI.User,False)+' le '+DateToStr(Date) ;
    If PourLot Then
      BEGIN
      Q.FindField('YX_ABREGE').AsString:=Abrege ;
      Q.FindField('YX_LIBRE').AsString:=Libre ;
      END ;
    Q.Post ;
    END ;
  Ferme(Q) ;
{$ENDIF}
END ;

Function EstMultiCatMP(smp : TSuiviMP) : Boolean ;
BEGIN
Result:=FALSE ;
If smp in [smpEncTraEdt,smpEncTraEdtNC,smpEncDiv,smpDecDiv] Then Result:=TRUE ;
END ;

procedure LaCATEGORIEChange(HCat : THValCombobox ; XX_WHEREMP : THEdit ; SorteLettre : TSorteLettre) ;
Var St,sMode,sCat,Junk : String ;
    i : integer ;
    Okok,OKLT,OKLC : boolean ;
    WhereMpVide : Boolean ;
begin
WhereMPVide:=FALSE ;
XX_WHEREMP.Text:='' ;
if HCAT.Value='' then
  BEGIN
  if (SorteLettre in [tslTraite,tslCheque,tslBor])=FALSE Then WhereMpVide:=TRUE Else // On prend tous les modes de paiement
    BEGIN // On ne prend que les modes de paiement à éditer
    for i:=0 to VH^.MPACC.Count-1 do
        BEGIN
        St:=VH^.MPACC[i] ;
        sMode:=ReadtokenSt(St) ; Junk:=ReadtokenSt(St) ; sCat:=ReadtokenSt(St) ;
        OKLC:=(ReadTokenSt(St)='X') ; OKLT:=(ReadTokenSt(St)='X') ;
        Case SorteLettre of
           tslTraite,tslBOR : Okok:=OKLT ;
                  tslCheque : Okok:=OKLC ;
                  else Okok:=FALSE ;
           END ;
        if Okok then XX_WHEREMP.Text:=XX_WHEREMP.Text+' OR E_MODEPAIE="'+sMode+'"' ;
        END ;
    END ;
  END Else
  BEGIN
  for i:=0 to VH^.MPACC.Count-1 do
      BEGIN
      St:=VH^.MPACC[i] ;
      sMode:=ReadtokenSt(St) ; Junk:=ReadtokenSt(St) ; sCat:=ReadtokenSt(St) ;
      OKLC:=(ReadTokenSt(St)='X') ; OKLT:=(ReadTokenSt(St)='X') ;
      if sCat=HCAT.Value then
         BEGIN
         (*
         If smp in [smpEncTraEdt,smpEncTraEdtNC] Then OkOk:=TRUE Else
           BEGIN
         *)
           Case SorteLettre of
              tslTraite,tslBOR : Okok:=OKLT ;
                     tslCheque : Okok:=OKLC ;
                     else Okok:=True ;
              END ;
           (*
           END ;
           *)
         if Okok then XX_WHEREMP.Text:=XX_WHEREMP.Text+' OR E_MODEPAIE="'+sMode+'"' ;
         END ;
      END ;
  END ;
If WhereMPVide Then XX_WHEREMP.Text:='' Else
  BEGIN
  if XX_WHEREMP.Text='' then
     BEGIN
     XX_WHEREMP.Text:='E_MODEPAIE="aaa"' ;
     END else // ne rien trouver
     BEGIN
     St:=XX_WHEREMP.Text ; Delete(St,1,4) ; XX_WHEREMP.Text:=St ;
     END ;
  END ;
//If smp in [smpEncDiv,smpDecDiv,smpEncTraEdt,smpEncTraEdtNC] Then PreciseMP ;
end;

Function TrouveJalCptEff(smp : tSuiviMP) : Integer ;
Var i : integer ;
    TOBL : TOB ;
    OkOk : Boolean ;
    NatCpt : String ;
    Cli : Boolean ;
BEGIN
Result:=-1 ; OkOk:=FALSE ;
If VH^.TOBJalEffet=Nil Then Exit ;
for i:=0 to VH^.TOBJalEffet.Detail.Count-1 do
    BEGIN
    TOBL:=VH^.TOBJalEffet.Detail[i] ;
    NatCpt:=TOBL.GetValue('G_NATUREGENE') ;
    Cli:=IsEnc(smp) ;
    Case Cli Of
      TRUE : OkOk:=(NatCpt='COC') Or (NatCpt='TID') Or (NatCpt='COD') Or (NatCpt='DIV') ;
      FALSE : OkOk:=(NatCpt='COF') Or (NatCpt='TIC') Or (NatCpt='COD') Or (NatCpt='DIV') ;
      End ;
    If OkOk Then BEGIN Result:=i ; Break ; END ;
    END ;
END ;

function  AttribHelp ( smp : TSuiviMP ; Suivi : Boolean = FALSE) : Integer ;
begin
  case smp of
    smpEncPreBqe     : If Suivi Then Result:=999999231 Else Result:=999999230 ;
    smpEncTraEdt     : If Suivi Then Result:=999999200 Else Result:=999999202 ;
    smpEncChqPor     : If Suivi Then Result:=0 Else Result:=0 ;
    smpEncCBPor      : If Suivi Then Result:=0 Else Result:=0 ;
    smpEncTraPor     : If Suivi Then Result:=999999203 Else Result:=999999204 ;
    smpEncTraEdtNC   : If Suivi Then Result:=0 Else Result:=999999201 ;
    smpEncTraEnc     : If Suivi Then Result:=999999205 Else Result:=999999206 ;
    smpEncTraEsc     : If Suivi Then Result:=999999207 Else Result:=999999208 ;
    smpEncChqBqe     : If Suivi Then Result:=0 Else Result:=0 ;
    smpEncCBBqe      : If Suivi Then Result:=0 Else Result:=0 ;
    smpEncTraBqe     : If Suivi Then Result:=999999209 Else Result:=999999210 ;
    smpEncPreEdt     : If Suivi Then Result:=999999200 Else Result:=999999202 ;
    smpEncPreEdtNC   : If Suivi Then Result:=0 Else Result:=999999201 ;
    smpEncDiv        : If Suivi Then Result:=0 Else Result:=0 ;
    smpEncTous       : If Suivi Then Result:=0 Else Result:=0 ;
    smpDecChqEdt     : If Suivi Then Result:=999999200 Else Result:=999999202 ;
    smpDecChqEdtNC   : If Suivi Then Result:=0 Else Result:=999999201 ;
    smpDecVirBqe     : If Suivi Then Result:=999999436 Else Result:=999999437 ;
    smpDecVirEdt     : If Suivi Then Result:=999999433 Else Result:=999999433 ;
    smpDecVirEdtNC   : If Suivi Then Result:=0 Else Result:=999999435 ;
    smpDecVirInBqe   : If Suivi Then Result:=999999441 Else Result:=999999442 ;
    smpDecVirInEdt   : If Suivi Then Result:=999999438 Else Result:=999999439 ;
    smpDecVirInEdtNC : If Suivi Then Result:=0 Else Result:=999999440 ;
    smpDecBorEdt     : If Suivi Then Result:=999999200 Else Result:=999999202 ;
    smpDecBorEdtNC   : If Suivi Then Result:=0 Else Result:=999999201 ;
    smpDecTraPor     : If Suivi Then Result:=999999203 Else Result:=999999204 ;
    smpDecBorDec     : If Suivi Then Result:=999999209 Else Result:=999999210 ;
    smpDecBorEsc     : If Suivi Then Result:=0 Else Result:=0 ;
    smpDecTraBqe     : If Suivi Then Result:=0 Else Result:=0 ;
    smpDecDiv        : If Suivi Then Result:=0 Else Result:=0 ;
    smpDecTous       : If Suivi Then Result:=0 Else Result:=0 ;
    else If Suivi Then Result:=0 Else Result:=0 ;
  end;
end;


{$IFDEF EAGLCLIENT}
procedure SwapModeGrid(GS : THGrid ; Q : THQuery) ;
Var OkOk : Boolean ;
    Selecting : Boolean ;
begin
    Okok := (goRowSelect in GS.Options) ;
    Selecting := (goEditing in GS.Options) ;
    Q.Manuel := true;
    if (Selecting) then
    begin
        GS.Options := GS.Options-[goEditing,goTabs,goAlwaysShowEditor];
        GS.Options := GS.Options+[goRowSelect];
    end
    else
    begin
        GS.Options := GS.Options-[goRowSelect];
        GS.Options := GS.Options+[goEditing,goTabs,goAlwaysShowEditor];
    end;
    Q.Manuel := false;
    if (Okok <> (goRowSelect in GS.Options)) then GS.Refresh;
end ;
{$ELSE}
Procedure SwapModeGrid(GS : thDBGrid ; Q : thQuery) ;
Var OkOk : Boolean ;
    Selecting : Boolean ;
begin
    Okok:=(dgRowSelect in GS.Options) ;
    Selecting:=(dgEditing in GS.Options) ;
    Q.Manuel:=TRUE ;
    if Selecting then
    begin
        Q.Close; {JP 09/05/07 : pour éviter "ne peut faire cette opération sur une ensemble de données ouvert}
        GS.Options:=GS.Options-[dgEditing,dgTabs,dgAlwaysShowEditor] ;
        GS.Options:=GS.Options+[dgRowSelect] ;
        Q.RequestLive:=FALSE ;
        Q.Open; {JP 09/05/07 : pour éviter "ne peut faire cette opération sur une ensemble de données ouvert}
    end
    else
    begin
        Q.Close; {JP 09/05/07 : pour éviter "ne peut faire cette opération sur une ensemble de données ouvert}
        GS.Options:=GS.Options-[dgRowSelect] ;
        GS.Options:=GS.Options+[dgEditing,dgTabs,dgAlwaysShowEditor] ;
        Q.RequestLive:=TRUE ;
        Q.Open; {JP 09/05/07 : pour éviter "ne peut faire cette opération sur une ensemble de données ouvert}
    end;
    Q.Manuel:=FALSE ;
    if Okok<>(dgRowSelect in GS.Options) then GS.Refresh ;
end;
{$ENDIF}

Function  SmpToStr ( smp : TSuiviMP ) : String ;
begin
  Case smp of
    smpEncPreBqe     : Result:='smpEncPreBqe' ;
    smpEncTraEdt     : Result:='smpEncTraEdt' ;
    smpEncTraPor     : Result:='smpEncTraPor' ;
    smpEncChqPor     : Result:='smpEncChqPor' ;
    smpEncCBPor      : Result:='smpEncCBPor' ;
    smpEncTraEdtNC   : Result:='smpEncTraEdtNC' ;
    smpEncTraEnc     : Result:='smpEncTraEnc' ;
    smpEncTraEsc     : Result:='smpEncTraEsc' ;
    smpEncChqBqe     : Result:='smpEncChqBqe' ;
    smpEncCBBqe      : Result:='smpEncCBBqe' ;
    smpEncTraBqe     : Result:='smpEncTraBqe' ;
    smpEncPreEdt     : Result:='smpEncPreEdt' ;
    smpEncPreEdtNC   : Result:='smpEncPreEdtNC' ;
    smpEncDiv        : Result:='smpEncDiv' ;
    smpEncTous       : Result:='smpEncTous' ;
    smpDecChqEdt     : Result:='smpDecChqEdt' ;
    smpDecChqEdtNC   : Result:='smpDecChqEdtNC' ;
    smpDecVirBqe     : Result:='smpDecVirBqe' ;
    smpDecVirEdt     : Result:='smpDecVirEdt' ;
    smpDecVirEdtNC   : Result:='smpDecVirEdtNC' ;
    smpDecVirInBqe   : Result:='smpDecVirInBqe' ;
    smpDecVirInEdt   : Result:='smpDecVirInEdt' ;
    smpDecVirInEdtNC : Result:='smpDecVirInEdtNC' ;
    smpDecBorEdt     : Result:='smpDecBorEdt' ;
    smpDecBorEdtNC   : Result:='smpDecBorEdtNC' ;
    smpDecTraPor     : Result:='smpDecTraPor' ;
    smpDecBorDec     : Result:='smpDecBorDec' ;
    smpDecBorEsc     : Result:='smpDecBorEsc' ;
    smpDecTraBqe     : Result:='smpDecTraBqe' ;
    smpDecDiv        : Result:='smpDecDiv' ;
    smpDecTous       : Result:='smpDecTous' ;
    {b FP 21/02/2006}
    smpCompenCli     : Result:='smpCompenCli' ;
    smpCompenFou     : Result:='smpCompenFou' ;
    {e FP 21/02/2006}
    else Result:='smpAucun' ;
  end ;
end ;

Function  StrToSmp ( vStsmp : String ) : TSuiviMP ;
begin
  if vStSmp = 'smpEncPreBqe'        then Result := smpEncPreBqe
  else if vStSmp = 'smpEncTraEdt'   then Result := smpEncTraEdt
  else if vStSmp = 'smpEncTraPor'   then Result := smpEncTraPor
  else if vStSmp = 'smpEncChqPor'   then Result := smpEncChqPor
  else if vStSmp = 'smpEncCBPor'    then Result := smpEncCBPor         
  else if vStSmp = 'smpEncTraEdtNC' then Result := smpEncTraEdtNC
  else if vStSmp = 'smpEncTraEnc'   then Result := smpEncTraEnc
  else if vStSmp = 'smpEncTraEsc'   then Result := smpEncTraEsc
  else if vStSmp = 'smpEncChqBqe'   then Result := smpEncChqBqe
  else if vStSmp = 'smpEncCBBqe'    then Result := smpEncCBBqe
  else if vStSmp = 'smpEncTraBqe'   then Result := smpEncTraBqe
  else if vStSmp = 'smpEncPreEdt'   then Result := smpEncPreEdt
  else if vStSmp = 'smpEncPreEdtNC' then Result := smpEncPreEdtNC
  else if vStSmp = 'smpEncDiv'      then Result := smpEncDiv
  else if vStSmp = 'smpEncTous'     then Result := smpEncTous
  else if vStSmp = 'smpDecChqEdt'   then Result := smpDecChqEdt
  else if vStSmp = 'smpDecChqEdtNC' then Result := smpDecChqEdtNC
  else if vStSmp = 'smpDecVirBqe'   then Result := smpDecVirBqe
  else if vStSmp = 'smpDecVirEdt'   then Result := smpDecVirEdt
  else if vStSmp = 'smpDecVirEdtNC' then Result := smpDecVirEdtNC
  else if vStSmp = 'smpDecVirInBqe' then Result := smpDecVirInBqe
  else if vStSmp = 'smpDecVirInEdt' then Result := smpDecVirInEdt
  else if vStSmp = 'smpDecVirInEdtNC' then Result := smpDecVirInEdtNC
  else if vStSmp = 'smpDecBorEdt'   then Result := smpDecBorEdt
  else if vStSmp = 'smpDecBorEdtNC' then Result := smpDecBorEdtNC
  else if vStSmp = 'smpDecTraPor'   then Result := smpDecTraPor
  else if vStSmp = 'smpDecBorDec'   then Result := smpDecBorDec
  else if vStSmp = 'smpDecBorEsc'   then Result := smpDecBorEsc
  else if vStSmp = 'smpDecTraBqe'   then Result := smpDecTraBqe
  else if vStSmp = 'smpDecDiv'      then Result := smpDecDiv
  else if vStSmp = 'smpDecTous'     then Result := smpDecTous
 {b FP 21/02/2006}
  else if vStSmp = 'smpCompenCli'   then Result := smpCompenCli
  else if vStSmp = 'smpCompenFou'   then Result := smpCompenFou
  {e FP 21/02/2006}
  else Result := smpAucun
end ;

// ============================================
// Nouvelles fonctions importées de MulGenereMP
// ============================================
function  GenereEscompteMP ( TOBEcr, TOBEscomptes, TOBGHT, TOBGTVA : TOB ;
                             vAvecTVA : Boolean ; JalEsc : String ; ParmsMP : TGenereMP ) : TOB ;
var i                         : Integer ;
    TOBResult                 : TOB ;
    TOBL, TOBE, TOBEsc        : TOB ;
    DetEsc1, DetEsc2, DetEsc3 : TOB ;
    Jal, Exo                  : String ;
    CompteHT, CompteTVA       : String ;
    TauxEsc, TauxTVA          : Double ;
    NewTaux                   : Double ;
    NumP, NumL, NumE, NumEsc  : integer ;
    MMEsc                     : RMVT ;
    DeltaJour                 : Integer ;
begin
  TOBResult := TOB.Create('',Nil,-1) ;
//  if Not GereEscompte then Exit ; // Test a faire dorénavant avant appel !!!
  for i:=0 to TOBEcr.Detail.Count-1 do
    begin
    TOBL := TOBEcr.Detail[i] ;
    if Not TOBL.FieldExists('ESCOMPTABLE') then Continue ;
    // Recherche si escompte attaché à la ligne
    Jal       := TOBL.GetValue('E_JOURNAL') ;
    Exo       := TOBL.GetValue('E_EXERCICE') ;
    NumP      := TOBL.GetValue('E_NUMEROPIECE') ;
    NumL      := TOBL.GetValue('E_NUMLIGNE') ;
    NumE      := TOBL.GetValue('E_NUMECHE') ;
    NewTaux   := TOBL.GetValue('ESNEWTAUX') ;
    DeltaJour := TOBL.GetValue('ESDELTAJOUR') ;

		// Récup paramétrage escompte
    TOBE := TOBEscomptes.FindFirst(['E_JOURNAL','E_EXERCICE','E_NUMEROPIECE','E_NUMLIGNE','E_NUMECHE'],
                                   [ Jal,        Exo,         NumP,           NumL,        NumE ],
                                   False ) ;
    if TOBE<>Nil then
      begin
      if (TOBE.GetValue('SANSESCOMPTE')='X') or (TOBE.GetValue('ESCOMPTABLE')<>'X')	then Continue ;
      end
    else
      if TOBEscomptes.GetValue('ESCOMPTABLE')<>'X' then Continue ;

    TauxEsc		:= TOBEscomptes.GetValue('TAUXESC') ;
    if TauxEsc = 0 then Continue ;
    TauxTVA		:= TOBEscomptes.GetValue('TAUXTVA') ;
    // Informations pièce escompte
    CompteTVA	:= TOBEscomptes.GetValue('COMPTETVA') ;
    CompteHT	:= TOBEscomptes.GetValue('COMPTEHT') ;
    TOBEsc    := TOB.Create('',Nil,-1) ;

    NumEsc  := GetNewNumJal(JalEsc,False,ParmsMP.DateC) ;
    if NumEsc<=0 then Continue ;
    MMEsc.Etabl     := TOBL.GetValue('E_ETABLISSEMENT') ;
    MMEsc.Jal       := JalEsc ;
    MMEsc.DateC     := ParmsMP.DateC ;
    MMEsc.Exo       := QuelExoDT(MMEsc.DateC) ;
    MMEsc.CodeD     := TOBL.GetValue('E_DEVISE') ;
    MMEsc.Simul     := 'S' ;
    MMEsc.Nature    := 'OD' ;
    MMEsc.DateTaux  := TOBL.GetValue('E_DATETAUXDEV') ;
    MMEsc.Num       := NumEsc ;
    MMEsc.TauxD     := TOBL.GetValue('E_TAUXDEV') ;
    MMEsc.Valide    := False ;
    MMEsc.ModeSaisieJal := '-' ;

    // Constitution ligne tiers
    DetEsc1:=TOB.Create('ECRITURE',TOBEsc,-1) ;
    DetEsc1.Dupliquer(TOBL,False,True) ;
    InitLigneEsc(DetEsc1,MMEsc,DeltaJour, ParmsMP) ;
    DetEsc1.PutValue('E_LIBELLE', IntToStr(numP) ); // insertion du numéro de pièce de la facture pour reprise dans libelle (Fiche10441)
    // Constitution ligne HT
    DetEsc2:=TOB.Create('ECRITURE',TOBEsc,-1) ;
    DetEsc2.Dupliquer(TOBL,False,True) ;
    InitLigneEsc(DetEsc2,MMEsc,DeltaJour, ParmsMP) ;
    DetEsc2.PutValue('E_LIBELLE', IntToStr(numP) ); // insertion du numéro de pièce de la facture pour reprise dans libelle (Fiche10441)
    // Constitution ligne TVA
    DetEsc3:=TOB.Create('ECRITURE',TOBEsc,-1) ;
    DetEsc3.Dupliquer(TOBL,False,True) ;
    InitLigneEsc(DetEsc3,MMEsc,DeltaJour, ParmsMP) ;
    DetEsc3.PutValue('E_LIBELLE', IntToStr(numP) ); // insertion du numéro de pièce de la facture pour reprise dans libelle (Fiche10441)
    // Finition
    FinirLignesEscompte( DetEsc1, DetEsc2, DetEsc3, TOBGHT, TOBGTVA,
                         CompteHT, CompteTVA,
                         TauxEsc, TauxTva, NewTaux,
                         vAvecTVA, ParmsMP) ;
	  // MAJ TOB Escompte avec le montant de l'escomte calculé
    if (TOBE<>Nil) then
    	begin
      // BPY le 23/08/2004 => fiche n° 13363 : il faut signé les escompte du a un probleme d'escompte sur avoir dans l'edition d'un reglement
      // attention : les escompte sont pour l'instant signé coté fournisseur et donc pas dans le bon sens pour les client !
      TOBE.PutValue('MONTANTESC', DetEsc1.GetValue('E_DEBIT') - DetEsc1.GetValue('E_CREDIT')) ;
      // fin BPY
      TOBE.PutValue('TAUXESC', 100 - NewTaux ) ;
      end ;

    if (not vAvecTVA) or (TauxTVA=0) or (CompteTVA='') then
      DetEsc3.Free ;
    // Insertion pièce escompte
    if TOBEsc.InsertDB(Nil) then
      if TOBEsc.Detail.Count>0 then
        begin
        TOB.Create( 'ECRITURE' , TOBResult , -1 ) ;
        TOBResult.Detail[ TOBResult.Detail.Count - 1 ].Dupliquer( TOBEsc.Detail[0], False, True ) ;
        end ;
    TOBEsc.Free ;
    end ;
  Result := TOBResult ;
End ;

procedure InitLigneEsc ( TOBL : TOB ; MMEsc : RMVT ; DeltaJour : Integer ; ParmsMP : TGenereMP ) ;

    Procedure _ModMontant(var debit : double ; var credit : double ; couverture : double);
    begin
        if (not (couverture = 0.0)) then
        begin
            if (not (debit = 0.0)) then debit := debit - couverture
            else if (not (credit = 0.0)) then credit := credit - couverture;
        end;
    end;

var
    Deb,Cred,Couv : double;
begin
  TOBL.PutValue('E_JOURNAL',        MMEsc.Jal) ;
  TOBL.PutValue('E_ETABLISSEMENT',  MMEsc.Etabl) ;
  TOBL.PutValue('E_NUMEROPIECE',    MMEsc.Num) ;
  TOBL.PutValue('E_NATUREPIECE',    MMEsc.Nature) ;
  TOBL.PutValue('E_QUALIFPIECE',    MMEsc.Simul) ;
  TOBL.PutValue('E_EXERCICE',       MMEsc.Exo) ;
  TOBL.PutValue('E_DATECOMPTABLE',  MMEsc.DateC) ;
  TOBL.PutValue('E_DATECREATION',   Date) ;
  TOBL.PutValue('E_DEVISE',         MMEsc.CodeD) ;
  TOBL.PutValue('E_DATETAUXDEV',    MMEsc.DateTaux) ;
  TOBL.PutValue('E_TAUXDEV',        MMEsc.TauxD) ;
  TOBL.PutValue('E_MODESAISIE',     MMEsc.ModeSaisieJal) ;

// BPY le 22/10/2004 => Fiche n° 14827 : les montant d'escompte sont calculé a partire de
//                                       E_CREDIT, E_DEBIT sans tenir compte de la couverture

    // recup des montant euro
    Deb := TOBL.GetValue('E_DEBIT');
    Cred := TOBL.GetValue('E_CREDIT');
    Couv := TOBL.GetValue('E_COUVERTURE');
    // modification des montant selon la couverture
    _ModMontant(Deb,Cred,Couv);
    // reset des monatnt
    TOBL.PutValue('E_DEBIT',Arrondi(Deb,V_PGI.OkDecV));
    TOBL.PutValue('E_CREDIT',Arrondi(Cred,V_PGI.OkDecV));
    TOBL.PutValue('E_COUVERTURE',0);

    // recup des montant en devise
    Deb := TOBL.GetValue('E_DEBITDEV');
    Cred := TOBL.GetValue('E_CREDITDEV');
    Couv := TOBL.GetValue('E_COUVERTUREDEV');
    // modification des montant selon la couverture
    _ModMontant(Deb,Cred,Couv);
    // reset des monatnt
    TOBL.PutValue('E_DEBITDEV',Arrondi(Deb,ParmsMP.DEV.Decimale));
    TOBL.PutValue('E_CREDITDEV',Arrondi(Cred,ParmsMP.DEV.Decimale));
    TOBL.PutValue('E_COUVERTUREDEV',0);

// Fin BPY

  TOBL.PutValue('E_LETTRAGE',       '') ;
  TOBL.PutValue('E_LETTRAGEDEV',    '-') ;
  TOBL.PutValue('E_ETATLETTRAGE',   'RI') ;
  TOBL.PutValue('E_REFPOINTAGE',    '') ;
  TOBL.PutValue('E_SUIVDEC',        '') ;
  TOBL.PutValue('E_NOMLOT',         '') ;
  TOBL.PutValue('E_TIERSPAYEUR',    '') ;
  TOBL.PutValue('E_ECRANOUVEAU',    'N') ;
  TOBL.PutValue('E_SAISIMP',        0) ;
  TOBL.PutValue('E_REFGESCOM',      '') ;
  TOBL.PutValue('E_PERIODE',        GetPeriode(MMEsc.DateC)) ;
  TOBL.PutValue('E_SEMAINE',        NumSemaine(MMEsc.DateC));
  TOBL.PutValue('E_QUALIFORIGINE',  'ESC') ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 27/12/2002
Modifié le ... : 30/12/2002
Description .. : 21/12/2002 : Modification du libellé et de la ref interne
Suite ........ : générée dans l'écriture d'escompte
Suite ........ :
Suite ........ : 30/12/2002 : Nouvelle méthode de calcul du montant de
Suite ........ : l'escompte pour éviter les éventuelles problèmes du à
Suite ........ : l'arrondi du montant escompté
Mots clefs ... :
*****************************************************************}
procedure FinirLignesEscompte ( T1, T2, T3, TOBGHT, TOBGTVA     : TOB ;
                                CompteHT, CompteTVA             : String ;
                                TauxEsc,TauxTva,NewTaux         : Double ;
                                vAvecTVA                        : Boolean ;
                                ParmsMP                         : TGenereMP ) ;
Var TOBG              : TOB ;
    i, NbDec, k, ia   : integer ;
    tsm               : TSorteMontant ;
    Gene, Aux, AuxLib : String ;
    TLoc, TAna, TAxe  : TOB ;
    TotLig, XTTC, XHT,
     XTVA, XX         : Array[1..2,TSorteMontant] of Double ;
    QTiers            : TQuery ;
    MontantTTC1       : Double ;
    MontantTTC2       : Double ;
    MontantEsc1       : Double ;
    MontantEsc2       : Double ;
    AvecTVA           : Boolean ;
    lStLibelle        : String ;    // pour construire le libellé de la pièce
BEGIN
  // Initialisations
  AvecTVA := vAvecTVA and (TauxTVA<>0) and (CompteTVA<>'') ;
  FillChar( TotLig, Sizeof(TotLig), #0) ;
  FillChar( XTTC,   Sizeof(XTTC),   #0) ;
  FillChar( XHT,    Sizeof(XHT),    #0) ;
  FillChar( XTVA,   Sizeof(XTVA),   #0) ;

  // FQ 18594 Gestion du montant MP saisie dans la grille
  {
  if Arrondi( T1.GetValue('E_SAISIMP'), ParmsMP.DEV.Decimale ) <> 0 then
    begin
    if T1.GetValue('E_DEBITDEV') <> 0 then
      begin
      TotLig[1,tsmPivot]  := T1.GetValue('E_SAISIMP') ;
      TotLig[2,tsmPivot]  := 0 ;
      TotLig[1,tsmDevise] := T1.GetValue('E_SAISIMP') ;
      TotLig[2,tsmDevise] := 0 ;
      end
    else
      begin
      TotLig[1,tsmPivot]  := 0  ;
      TotLig[2,tsmPivot]  := T1.GetValue('E_SAISIMP') ;
      TotLig[1,tsmDevise] := 0  ;
      TotLig[2,tsmDevise] := T1.GetValue('E_SAISIMP') ;
      end ;
    end
  else
  }  begin
    TotLig[1,tsmPivot]  := T1.GetValue('E_DEBIT')  ;
    TotLig[2,tsmPivot]  := T1.GetValue('E_CREDIT') ;

    TotLig[1,tsmDevise] := T1.GetValue('E_DEBITDEV') ;
    TotLig[2,tsmDevise] := T1.GetValue('E_CREDITDEV') ;
    end ;

  TotLig[1,tsmEuro]   := 0  ;
  TotLig[2,tsmEuro]   := 0 ;
  Gene := T1.GetValue('E_GENERAL') ;
  Aux  := T1.GetValue('E_AUXILIAIRE') ; TLoc:=nil;
  // Construction du libelle de l'escompte (maxi 35 caractères) Fiche 10441 :
  // - Escpt S/pièce N° <Numéro pièce d'origine> <Libelle tiers tronqué>
  AuxLib := Aux ;
  if Aux <> '' then
    begin
    QTiers := OpenSQL('SELECT T_LIBELLE FROM TIERS WHERE T_AUXILIAIRE="'+Aux+'"', True) ;
    if not QTiers.Eof then
      AuxLib := QTiers.FindField('T_LIBELLE').AsString ;
    Ferme(QTiers) ;
    end ;
  lStLibelle := 'Escpte S/pièce N° ' + T1.GetValue('E_LIBELLE') + ' ' ;
  lStLibelle := Copy( lStLibelle + AuxLib, 1, 35) ;

  // Ligne tiers
  T1.PutValue('E_NUMLIGNE',           1) ;
  T1.PutValue('E_NUMECHE',            1) ;
  T1.PutValue('E_CONTREPARTIEGEN',    CompteHT ) ;
  T1.PutValue('E_CONTREPARTIEAUX',    '') ;
  T1.PutValue('E_ETATLETTRAGE',       'AL') ;
  // Modification libelle de l'écriture d'escompte Fiche 10441
  T1.PutValue('E_LIBELLE',            lStLibelle ) ;
  // FQ 22141 : initialiser la date d'échéance à la date comptable du règlement
  T1.PutValue('E_DATEECHEANCE',       ParmsMP.DateC ) ;
  // Fin FQ 22141

  // Ligne HT
  TOBG := TOBGHT.FindFirst(['G_GENERAL'],[CompteHT],True) ;
  T2.PutValue('E_NUMLIGNE',         2) ;
  T2.PutValue('E_NUMECHE',          0) ;
  T2.PutValue('E_TYPEMVT',          QuelTypeMvt(CompteHT, TOBG.GetValue('G_NATUREGENE'), T2.GetValue('E_NATUREPIECE') ) ) ;
  T2.PutValue('E_GENERAL',          CompteHT) ;
  T2.PutValue('E_AUXILIAIRE',       '') ;
  T2.PutValue('E_ECHE',             '-') ;
  T2.PutValue('E_ETAT',             '0000000000') ;
  T2.PutValue('E_CONTREPARTIEGEN',  Gene) ;
  T2.PutValue('E_CONTREPARTIEAUX',  Aux) ;
  T2.PutValue('E_TVA',              TOBG.GetValue('G_TVA')) ;
  T2.PutValue('E_TPF',              TOBG.GetValue('G_TPF')) ;
  if TOBG.GetValue('G_VENTILABLE')='X' then
    T2.PutValue('E_ANA','X') ;
  // Modification libelle de l'écriture d'escompte Fiche 10441
  T2.PutValue('E_LIBELLE',          lStLibelle ) ;

  // Ligne TVA
  if AvecTVA then
    begin
    TOBG := TOBGTVA.FindFirst(['G_GENERAL'],[CompteTVA],True) ;
    T3.PutValue('E_NUMLIGNE',         3) ;
    T3.PutValue('E_NUMECHE',          0) ;
    T3.PutValue('E_TYPEMVT',          QuelTypeMvt(CompteTVA,TOBG.GetValue('G_NATUREGENE'),T3.GetValue('E_NATUREPIECE'))) ;
    T3.PutValue('E_GENERAL',          CompteTVA) ;
    T3.PutValue('E_AUXILIAIRE',       '') ;
    T3.PutValue('E_ECHE',             '-') ;
    T3.PutValue('E_ETAT',             '0000000000') ;
    T3.PutValue('E_CONTREPARTIEGEN',  Gene) ;
    T3.PutValue('E_CONTREPARTIEAUX',  Aux) ;
    if TOBG.GetValue('G_VENTILABLE')='X' then
      T3.PutValue('E_ANA','X') ;
    // Modification libelle de l'écriture d'escompte Fiche 10441
    T3.PutValue('E_LIBELLE',          lStLibelle ) ;
  end ;

  // calcul des montants
  for i:=0 to 2 do
    begin
    tsm := TSorteMontant(i) ;
    if i=0 then
      NbDec:=V_PGI.OkDecV
    else if i=1
      then NbDec := ParmsMP.DEV.Decimale
      else NbDec := V_PGI.OkDecE ;
    // == TTC ==
    // Récupération des montant de la facture
    MontantTTC1 := TotLig[1,tsm] ;                		      // Montant TTC Débit
    MontantTTC2 := TotLig[2,tsm] ;                                    // Montant TTC Crédit
    // Montant Escompté calculé comme précédemment
    MontantEsc1 := Arrondi( (MontantTTC1 * (NewTaux/100)) , NbDec) ;  // Montant Escompté Débit
    MontantEsc2 := Arrondi( (MontantTTC2 * (NewTaux/100)) , NbDec) ;  // Montant Escompté Crédit
    // Montant de l'escompte calculé par différence
    XTTC[1,tsm]	:= Arrondi(MontantTTC2 - MontantEsc2, NbDec) ;        // Montant Escompte Débit
    XTTC[2,tsm]	:= Arrondi(MontantTTC1 - MontantEsc1, NbDec) ;	      // Montant Escompte Crédit
    // == HT ==
    XHT[1,tsm] := Arrondi(XTTC[2,tsm] / (1.0+TauxTva/100.0), NbDec) ;
    XHT[2,tsm] := Arrondi(XTTC[1,tsm] / (1.0+TauxTva/100.0), NbDec) ;
    // == TVA ==
    XTVA[1,tsm] := Arrondi(XTTC[2,tsm] - XHT[1,tsm], NbDec) ;
    XTVA[2,tsm] := Arrondi(XTTC[1,tsm] - XHT[2,tsm], NbDec) ;
    END ;

  // Affectation des montants
  for k:=1 to 3 do
    begin
    if k=1 then
      BEGIN
      TLoc  := T1 ;
      XX    := XTTC ;
      END
    else if k=2 then
           BEGIN
           TLoc := T2 ;
           XX   := XHT ;
           END
         else
           BEGIN
           if not AvecTVA then Continue ;
           TLoc := T3 ;
           XX   := XTVA ;
           END ;
    TLoc.PutValue('E_DEBIT',      XX[1,tsmPivot]) ;
    TLoc.PutValue('E_CREDIT',     XX[2,tsmPivot]) ;
    TLoc.PutValue('E_DEBITDEV',   XX[1,tsmDevise]) ;
    TLoc.PutValue('E_CREDITDEV',  XX[2,tsmDevise]) ;
    // Analytiques
    if TLoc.GetValue('E_ANA')='X' then
      BEGIN
      AlloueAxe(TLoc);

      VentilerTOB(TLoc,'',0, ParmsMP.DEV.Decimale, False) ;
      for i:=0 to TLoc.Detail.Count-1 do
        begin
        TAxe:=TLoc.Detail[i] ;
        for ia:=0 to TAxe.Detail.Count-1 do
          begin
          TAna := TAxe.Detail[ia] ;
          TAna.PutValue('Y_QUALIFPIECE',    'S' ) ;
          TAna.PutValue('Y_NUMEROPIECE',    TLoc.GetValue('E_NUMEROPIECE') ) ;
          TAna.PutValue('Y_JOURNAL',        TLoc.GetValue('E_JOURNAL') ) ;
          TAna.PutValue('Y_EXERCICE',       TLoc.GetValue('E_EXERCICE') ) ;
          TAna.PutValue('Y_DATECOMPTABLE',  TLoc.GetValue('E_DATECOMPTABLE') ) ;
          TAna.PutValue('Y_PERIODE',        TLoc.GetValue('E_PERIODE') ) ;
          TAna.PutValue('Y_SEMAINE',        TLoc.GetValue('E_SEMAINE') ) ;
          TAna.PutValue('Y_DEVISE',         TLoc.GetValue('E_DEVISE') ) ;
          end ;
        end ;
      end ;
    end ;
end ;

procedure TraiteTOBEscompte ( TOBL, TOBEscomptes, TOBGHT, TOBGTVA : TOB ;
                              vBoAuProrata : Boolean ;    // ajout 07/04/2002
                              vStMeth : String ;          // ajout 07/04/2002
                              ParmsMP : TGenereMP ;
                              var PasPris : Boolean) ;
Var TOBE            : TOB ;
    Jal,Exo         : String ;
    NumP,NumL,NumE  : integer ;
    TauxEsc, X, NewTaux : Double ;
    CompteHT,CompteTva : String ;
    OkEsc           : boolean ;
    Delta           : Integer ;
    TauxEscF 		    : Double ;
BEGIN
//   if Not GereEscompte then Exit ;  // Test a faire avant appel de la fonction
  if TOBL = Nil then Exit ;
  Jal   := TOBL.GetValue('E_JOURNAL')      ;
  Exo   := TOBL.GetValue('E_EXERCICE') ;
  NumP  := TOBL.GetValue('E_NUMEROPIECE') ;
  NumL  := TOBL.GetValue('E_NUMLIGNE') ;
  NumE  := TOBL.GetValue('E_NUMECHE') ;
  TOBE := TOBEscomptes.FindFirst(['E_JOURNAL','E_EXERCICE','E_NUMEROPIECE','E_NUMLIGNE','E_NUMECHE'],
                                 [ Jal,        Exo,         NumP,           NumL,        NumE ],
                                 False) ;
  if TOBE = Nil then
    begin
    OkEsc     := TOBEscomptes.GetValue('ESCOMPTABLE')='X' ;
    TauxEsc   := TOBEscomptes.GetValue('TAUXESC') ;
    CompteTVA := TOBEscomptes.GetValue('COMPTETVA') ;
    AjouteComptePourEsc(CompteTVA,TOBGTVA) ;
    CompteHT  := TOBEscomptes.GetValue('COMPTEHT')   ;
    AjouteComptePourEsc(CompteHT,TOBGHT) ;
    end
  else
    begin
    // Si ligne SANS ESCOMPTE --> Exit
    if TOBE.FieldExists('SANSESCOMPTE') then
       if TOBE.GetValue('SANSESCOMPTE')='X' then Exit ;
    OkEsc     := TOBE.GetValue('ESCOMPTABLE')='X' ;
    TauxEsc   := TOBEscomptes.GetValue('TAUXESC') ;
    CompteTVA := TOBEscomptes.GetValue('COMPTETVA') ;
    AjouteComptePourEsc(CompteTVA,TOBGTVA) ;
    CompteHT  := TOBEscomptes.GetValue('COMPTEHT')   ;
    AjouteComptePourEsc(CompteHT,TOBGHT) ;
    end ;

  if Not OkEsc then Exit ;
  if TauxEsc=0 then Exit ;

  if TOBL.GetValue('E_SAISIMP')<>0
    then X := TOBL.GetValue('E_SAISIMP')
    else X := TOBL.GetValue('E_DEBIT') + TOBL.GetValue('E_CREDIT') - TOBL.GetValue('E_COUVERTURE') ;

  NewTaux := CalcNewTauxEscompte( TauxEsc, X, TOBL,
                                  vBoAuProrata, vStMeth, ParmsMP,
                                  Delta, TauxEscF, PasPris ) ;
  If NewTaux<=0 Then Exit ;

  // Ajout arrondi - montant local
  X := Arrondi( (NewTaux*X) / 100.0, V_PGI.OkDecV ) ;
  TOBL.PutValue('E_SAISIMP',X) ;

  if X<>0 then
    begin
    TOBL.AddChampSup('ESCOMPTABLE', False) ;
    TOBL.AddChampSup('ESNEWTAUX',   False) ;
    TOBL.AddChampSup('ESDELTAJOUR', False) ;
    TOBL.AddChampSup('ESTAUXFOURN', False) ;
    TOBL.PutValue('ESCOMPTABLE',  'X') ;
    TOBL.PutValue('ESNEWTAUX',    NewTaux) ;
    TOBL.PutValue('ESDELTAJOUR',  Delta) ;
    TOBL.PutValue('ESTAUXFOURN',  TauxEscF) ;
    end ;
end;

procedure AjouteComptePourEsc ( Cpt : String ; TOBRech : TOB ) ;
var TOBG  : TOB ;
    QQ    : TQuery ;
begin
  if Cpt='' then Exit ;
  TOBG := TOBRech.FindFirst(['G_GENERAL'],[Cpt],True) ;
  if TOBG<>Nil then Exit ;
  QQ:=OpenSQL('SELECT * FROM GENERAUX WHERE G_GENERAL="'+Cpt+'"',True) ;
  if Not QQ.EOF then
    TOBRech.LoadDetailDB('GENERAUX','','',QQ,True) ;
  Ferme(QQ) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 27/12/2002
Modifié le ... :   /  /
Description .. : 27/12/2001 : Changement de la méthode de calcul du
Suite ........ : prorata : (Écart en jours entre date d'échéance et la date
Suite ........ : comptable)/(Écart en jours entre date comptable et la date
Suite ........ : d'échéance)*taux d'escompte initial = taux d'escompte réel
Mots clefs ... :
*****************************************************************}
function  CalcNewTauxEscompte(  TauxEsc, X   : Double ;
                                TOBL         : TOB ;
                                vBoAuProrata : Boolean ;    // Calcul proratisé
                                vStMethEsc   : String ;     // Méthode de calcul
                                ParmsMP      : TGenereMP ;
                                Var Delta    : Integer ;
                                var TauxEscF : Double ;
                                Var PasPris  : Boolean ) : Double ;
var NewTaux, XPaye, XEscompte : Double ;
    DateEch, DateCpt          : TDateTime ;
    XDelta, MaxDelta          : Double ;
begin
  // Init variables locales
  Delta		:= 0 ;
	DateEch	:=TOBL.GetValue('E_DATEECHEANCE') ;			// Date échéance de la facture
	DateCpt	:=TOBL.GetValue('E_DATECOMPTABLE') ;		// Date comptable de la facture
	TauxEscF:=TOBL.GetValue('T_ESCOMPTE') ;					// Taux d'escompte du tiers
  // Suivant méthode d'escompte, maj taux d'escompte
	If ( (TauxEscF<>0) And (vStMethEsc='TFX') )
    Then tauxEsc  := TauxEscF
    Else TauxEscF := 0 ;
  // Calcul du taux et montant escompte
	NewTaux:=(1.0-TauxEsc/100.0) ;
  Result:=Newtaux ;
	XPaye:=NewTaux*X ;
  XEscompte:=Arrondi(X-XPaye,V_PGI.OkDecV) ;
	// Calcul simple ou au prorata...
  If vBoAuProrata Then
  	begin
		// Calcul des écarts de date pour l'éventuelle proratisation
		XDelta:=Arrondi(DateEch-ParmsMP.DateC,0) ;
		MaxDelta := Arrondi(DateEch-DateCpt,0) ;
		If (XDelta<=0) or (MaxDelta=0) Then
      BEGIN
      Result:=0 ;
      PasPris:=TRUE ;
      END
    Else
    	Begin
      // Nouveau montant d'escompte proratisé
      XEscompte:=XEscompte*(XDelta/MaxDelta) ;
      // Nouveau taux d'escompte proratisé
      Result	:= (100*(X - XEscompte)) / X ;
     	// MAJ Delta de paiement
			Delta := Round(XDelta) ;
      end ;
    end
  else
  	Result:=100*Result ;
end ;

procedure InitCommuns ( TOBD : TOB ) ;
var i : integer ;
begin
  // Lettrage
  TOBD.PutValue('E_LETTRAGE',       '') ;
  TOBD.PutValue('E_ETATLETTRAGE',   'RI') ;
  TOBD.PutValue('E_COUVERTURE',     0) ;
  TOBD.PutValue('E_COUVERTUREDEV',  0) ;
  TOBD.PutValue('E_LETTRAGEDEV',    '-') ;
  TOBD.PutValue('E_ETAT',           '0000000000') ;
  // Eches + Anas
  TOBD.PutValue('E_ECHE', '-') ;
  TOBD.PutValue('E_ANA',  '-') ;
  // Saisie
  TOBD.PutValue('E_MODESAISIE',   '-') ;
//  if GeneOppose then
//    TOBD.PutValue('E_SAISIEEURO', 'X') ;
  TOBD.PutValue('E_QUALIFORIGINE','') ;
  // Enc et Dec
  TOBD.PutValue('E_SUIVDEC',      '') ;
  TOBD.PutValue('E_NOMLOT',       '') ;
  TOBD.PutValue('E_REFRELEVE',    '') ;
  TOBD.PutValue('E_ENCAISSEMENT', SensEnc( TOBD.GetValue('E_DEBITDEV') , TOBD.GetValue('E_CREDITDEV') ) ) ;
  // Tiers
  TOBD.PutValue('E_DATERELANCE',  iDate1900) ;
  TOBD.PutValue('E_NIVEAURELANCE',0) ;
  // TVA
  for i:=1 to 4 do
    TOBD.PutValue('E_ECHEENC'+IntToStr(i),  0) ;
  TOBD.PutValue('E_ECHEDEBIT',    0) ;
  // Divers
  TOBD.PutValue('E_VALIDE',       '-') ;
  TOBD.PutValue('E_AVOIRRBT',     '-') ;
  TOBD.PutValue('E_PIECETP',      '') ;
  TOBD.PutValue('E_EQUILIBRE',    '-') ;
  TOBD.PutValue('E_CFONBOK',      '-') ;
  TOBD.PutValue('E_EDITEETATTVA', '-') ;
  TOBD.PutValue('E_REFGESCOM',    '') ;
  {JP 09/11/05 : FQ 15771 : Si la contrepartie n'est pas lettrable, il faut que E_NUMECHE soit réinitialisé}
  TOBD.PutValue('E_NUMECHE',      0) ;
end ;

procedure TripoteDetail ( TOBD : TOB ; ParmsMP : TGenereMP ) ;
BEGIN
  InitCommuns(TOBD) ;
  if ParmsMP.Ref2<>'' then
    TOBD.PutValue('E_REFINTERNE', ParmsMP.Ref2) ;
  if ParmsMP.Lib2<>'' then
    TOBD.PutValue('E_LIBELLE',    ParmsMP.Lib2) ;
  if ParmsMP.RefExt2<>'' then
    TOBD.PutValue('E_REFEXTERNE', ParmsMP.RefExt2) ;
  if ParmsMP.RefLib2<>'' then
    TOBD.PutValue('E_REFLIBRE',   ParmsMP.RefLib2) ;
  TOBD.PutValue('E_GENERAL',      ParmsMP.CptG) ;
  if ParmsMP.MPG<>'' then
    TOBD.PutValue('E_MODEPAIE',   ParmsMP.MPG) ;
  // Cas particuliers
  if ParmsMP.GenColl then
    begin
    TOBD.PutValue('E_ETATLETTRAGE', 'AL') ;
    TOBD.PutValue('E_ECHE',         'X') ;
    TOBD.PutValue('E_NUMECHE',      1) ;
    end
  else
    begin
    TOBD.PutValue('E_AUXILIAIRE',   '') ;
    if ParmsMP.GenLett then
      begin
      TOBD.PutValue('E_ETATLETTRAGE', 'AL') ;
      TOBD.PutValue('E_ECHE',         'X') ;
      TOBD.PutValue('E_NUMECHE',      1) ;
      end ;
   if ParmsMP.GenPoint then
      begin
      TOBD.PutValue('E_ETATLETTRAGE', 'RI') ;
      TOBD.PutValue('E_ECHE',         'X') ;
      TOBD.PutValue('E_NUMECHE',      1) ;
      end ;
   end ;
  if ParmsMP.GenVent
    then TOBD.PutValue('E_ANA',       'X')
    else TOBD.ClearDetail ;
end ;

procedure AffecteG ( TOBL : TOB ; ParmsMP : TGenereMP ) ;
var TOBA, TOBAxe : TOB ;
    i, ia        : integer ;
begin
  TOBL.PutValue('E_EXERCICE',         QuelExoDT(ParmsMP.DateC)) ;
  TOBL.PutValue('E_DATECOMPTABLE',    ParmsMP.DateC) ;
  TOBL.PutValue('E_PERIODE',          GetPeriode(ParmsMP.DateC)) ;
  TOBL.PutValue('E_SEMAINE',          NumSemaine(ParmsMP.DateC)) ;
  TOBL.PutValue('E_JOURNAL',          ParmsMP.JalG) ;

  TOBL.PutValue('E_QUALIFPIECE',      'S') ;
  if ((TOBL.GetValue('E_ECHE')='X') and (ParmsMP.DateE>0) and (ParmsMP.ForceEche)) then
    TOBL.PutValue('E_DATEECHEANCE',   ParmsMP.DateE) ;
  if (TOBL.GetValue('E_ECHE')='X') And  (ParmsMP.MPG<>'') then
    TOBL.PutValue('E_MODEPAIE',       ParmsMP.MPG) ;
  for ia:=0 to TOBL.Detail.Count-1 do
    begin
    TOBAxe := TOBL.Detail[ia] ;
    for i:=0 to TOBAxe.Detail.Count-1 do
      begin
      TOBA:=TOBAxe.Detail[i] ;
      TOBA.PutValue('Y_EXERCICE',       TOBL.GetValue('E_EXERCICE')) ;
      TOBA.PutValue('Y_DATECOMPTABLE',  TOBL.GetValue('E_DATECOMPTABLE')) ;
      TOBA.PutValue('Y_JOURNAL',        TOBL.GetValue('E_JOURNAL')) ;
      TOBA.PutValue('Y_NUMEROPIECE',    TOBL.GetValue('E_NUMEROPIECE')) ;
      TOBA.PutValue('Y_NUMLIGNE',       TOBL.GetValue('E_NUMLIGNE')) ;
      TOBA.PutValue('Y_NUMVENTIL',      IntToStr(i+1)) ;
      TOBA.PutValue('Y_GENERAL',        TOBL.GetValue('E_GENERAL')) ;
      TOBA.PutValue('Y_REFINTERNE',     TOBL.GetValue('E_REFINTERNE')) ;
      TOBA.PutValue('Y_LIBELLE',        TOBL.GetValue('E_LIBELLE')) ;
      TOBA.PutValue('Y_NATUREPIECE',    TOBL.GetValue('E_NATUREPIECE')) ;
      TOBA.PutValue('Y_AUXILIAIRE',     TOBL.GetValue('E_AUXILIAIRE')) ;
      TOBA.PutValue('Y_PERIODE',        TOBL.GetValue('E_PERIODE')) ;
      TOBA.PutValue('Y_SEMAINE',        TOBL.GetValue('E_SEMAINE')) ;
      TOBA.PutValue('Y_QUALIFPIECE',    TOBL.GetValue('E_QUALIFPIECE')) ;
      END ;
    END ;
END ;

function  GetLeMontant ( Montant,Couv,Sais : Double ; tsm : TSorteMontant ) : Double ;
begin
	Result := 0 ;
	if Arrondi(Montant,4) = 0 then Exit ;
  if Arrondi(Sais,4) = 0
  	then Result := Arrondi(Montant-Couv, 6)
		else Result := Sais ;
end ;

procedure FinirDestination ( TOBDest, TOBR : TOB ; ParmsMP : TGenereMP ) ;
var i,j,Ind,NewNum,k : integer ;
    TOBL        : TOB ;
    XD,XC,EcheD,ValMax    : Double ;
    Eche: array[1..4] of Double;
    lStNatPiece : String ;
    szNat       : String ;
    szNatGene   : String ;
    TxTva, RegTva : String ;
    lDtDateModif : TDateTime ;
begin

//  ValMax:=0;
  // Traitement des lignes d'échéances
  for i:=0 to TOBDest.Detail.Count-1 do
    BEGIN
    TOBL := TOBDest.Detail[i] ;
     // Montants
    XD   := GetLeMontant(TOBL.GetValue('E_DEBIT'),TOBL.GetValue('E_COUVERTURE'),TOBL.GetValue('E_SAISIMP'),tsmPivot) ;
    XC   := GetLeMontant(TOBL.GetValue('E_CREDIT'),TOBL.GetValue('E_COUVERTURE'),TOBL.GetValue('E_SAISIMP'),tsmPivot) ;
    TOBL.PutValue('E_DEBIT',      XC ) ;
    TOBL.PutValue('E_CREDIT',     XD ) ;
    // Montants Dev
    XD   := GetLeMontant(TOBL.GetValue('E_DEBITDEV'),TOBL.GetValue('E_COUVERTUREDEV'),TOBL.GetValue('E_SAISIMP'),tsmDevise) ;
    XC   := GetLeMontant(TOBL.GetValue('E_CREDITDEV'),TOBL.GetValue('E_COUVERTUREDEV'),TOBL.GetValue('E_SAISIMP'),tsmDevise) ;
    TOBL.PutValue('E_DEBITDEV',   XC ) ;
    TOBL.PutValue('E_CREDITDEV',  XD ) ;
    if TOBL.GetValue('E_ETATLETTRAGE')<>'RI' then
      begin
      TOBL.PutValue('E_LETTRAGE',     '' ) ;
      TOBL.PutValue('E_ETATLETTRAGE', 'AL' ) ;
      TOBL.PutValue('E_ECHE',         'X' ) ;
      TOBL.PutValue('E_NUMECHE',      1 ) ;
      end ;

    // BPY le 22/10/2004 => Fiche n°14827 : pb de lettrage du au lettrage sur reference
    //                                      set de la ref a n'importe nawak en cas de reglement globalisé s'il y a plusieurs lignes!
{   // SBO 29/10/2005 : Demande RR pour oter code BPY pour SIC
    if (not (ParmsMP.Ref1 = '')) then
      TOBL.PutValue('E_REFINTERNE',ParmsMP.Ref1)
    else if ((not (ParmsMP.GroupeEncaDeca = 'DET')) and (TOBDest.Detail.Count > 1)) then
      TOBL.PutValue('E_REFINTERNE','Rglt. ' + FormatDateTime('DD/MM/YYYY',V_PGI.DateEntree));
}    // Fin BPY

    if ParmsMP.Ref1<>'' then
      TOBL.PutValue('E_REFINTERNE',      ParmsMP.Ref1) ;

    if ParmsMP.Lib1<>'' then
      TOBL.PutValue('E_LIBELLE',      ParmsMP.Lib1) ;
    if ParmsMP.RefExt1<>'' then
      TOBL.PutValue('E_REFEXTERNE',   ParmsMP.RefExt1) ;
    if ParmsMP.RefLib1<>'' then
      TOBL.PutValue('E_REFLIBRE',     ParmsMP.RefLib1) ;
    if ParmsMP.Lib1<>'' then
      TOBL.PutValue('E_LIBELLE',      ParmsMP.Lib1) ;
    TOBL.PutValue('E_SAISIMP',        0) ;
    TOBL.PutValue('E_COUVERTURE',     0) ;
    TOBL.PutValue('E_COUVERTUREDEV',  0) ;
    FinirAnas( TOBL, ParmsMP ) ;
    end ;


  // Traitement des lignes de contrepartie
  for i:=TOBR.Detail.Count-1 downto 0 do
  begin
    TOBL := TOBR.Detail[i] ;
    TxTva:='';
    RegTva:='';
    For k:=1 to 4 do Eche[k]:=0.0;
    EcheD:=0;
    ValMax:=0;
    for j := (TobDest.Detail.Count - 1) downto 0 do
    begin

      if TobDest.Detail[j].GetValue('GROUPEIDX') = TOBL.GetValue('GROUPEIDX') then
      begin
         {Reprise des infos de TVA de la ligne ayant la valeur la + élevée}
          For k:=1 to 4 do    {ValMax est valable pour toutes les lignes}
          begin
            Eche[k] := Eche[k]+TobDest.Detail[j].GetDouble('E_ECHEENC'+inttostr(k)) ;
            If TobDest.Detail[j].GetDouble('E_ECHEENC'+inttostr(k))>ValMax then
            begin
              TxTva:=TobDest.Detail[j].GetString('E_TVA');
              RegTva:=TobDest.Detail[j].GetString('E_REGIMETVA');
              ValMax:=TobDest.Detail[j].GetDouble('E_ECHEENC'+inttostr(k));
            end;
          end;
          EcheD := EcheD+TobDest.Detail[j].GetDouble('E_ECHEDEBIT') ;

          {Transfert des infos de TVA}
          for k:=1 to 4 do
            TOBL.PutValue('E_ECHEENC'+inttostr(k), Eche[k] ) ;

          TOBL.PutValue('E_ECHEDEBIT',  EcheD ) ;
          TOBL.PutValue('E_TVA', txTva);
          TOBL.PutValue('E_REGIMETVA', RegTva);
      end ;

    end;

   end;

   for i:=TOBR.Detail.Count-1 downto 0 do
   begin
    TOBL := TOBR.Detail[i] ;

    // Correction pour les groupes de lignes avec un solde à zéro --> on ôte tout le bloc !
    if ( TOBL.GetValue('E_DEBIT') + TOBL.GetValue('E_CREDIT') ) = 0 then
    begin
      // On ôte le groupe de ligne d'échéance correspondant
      for j := (TobDest.Detail.Count - 1) downto 0 do
      begin
      if TobDest.Detail[j].GetValue('GROUPEIDX') = TOBL.GetValue('GROUPEIDX') then
        TobDest.Detail[j].Free ;
      end ;

      // On libère la ligne de contrepartie
      TOBL.Free ;
      // On passe au bloc suivant
      continue ;
      end ;
      // Fin Correction pour les lignes soldé --> pas de ligne de contrepartie !
      AnasRupt( TOBL, ParmsMP ) ;

      // Insertion dans la pièce finale
      Ind  := TOBL.GetValue('INDICE') ;
      TOBL.ChangeParent(TOBDest,Ind) ;
    end ;


  //////////////////////////
  // Uniformisations des infos pièces tel que numéro, date modif, nature pièce...
  NewNum := GetNewNumJal(ParmsMP.JalG,False,ParmsMP.DateC) ;
  lDtDateModif := NowH ;

  // ===> Ajout SBO 28/07/2004 :
  // On détermine une fois pour toute la nature de la pièce (avec Méthode pour FFF )
  if TOBDest.Detail.Count > 0 then
    begin
    TOBL := TOBDest.Detail[0] ;
    szNatGene := TOBL.GetValue('G_NATUREGENE');

    // FQ 14821 ( Les pièces générées sur un journal d'OD doivent être de type OD ) // SBO 05/11/2004
    if (ParmsMP.NatJalG = 'OD') then lStNatPiece := 'OD'
    else if ((szNatGene) <> 'COD') then
      begin
      szNat := TOBL.GetValue('E_NATUREPIECE');
      if (szNat = 'OD') then
        begin
        if (szNatGene = 'COC') then lStNatPiece := 'RC'
        else if (szNatGene = 'COF') then lStNatPiece := 'RF'
        else if (IsEnc(ParmsMP.smp)) then lStNatPiece := 'RC'
        else if (not(IsEnc(ParmsMP.smp))) then lStNatPiece := 'RF' ;
        end
      else if (szNat = 'OC') or (szNat = 'AC') or (szNat = 'FC') or (szNat = 'RC') then lStNatPiece := 'RC'
      else if (szNat = 'OF') or (szNat = 'AF') or (szNat = 'FF') or (szNat = 'RF') then lStNatPiece := 'RF'
      else lStNatPiece := ParmsMP.NatureG;
      end
    else lStNatPiece := ParmsMP.NatureG ; // Collectif divers // Attention danger potentiel, devrait toujours être OD aussi !
    end
  else lStNatPiece := 'OD' ;
  // ===> Fin Ajout SBO 28/07/2004

  // En attendant, génération uniquement sur des journaux de type pièce
  for i:=0 to TOBDest.Detail.Count-1 do
    begin
    TOBL := TOBDest.Detail[i] ;

    TOBL.PutValue('E_NUMLIGNE'	 , i+1) ;
    TOBL.PutValue('E_NUMEROPIECE', NewNum) ;

    // VL 301003 FQ 12958
    if ParmsMP.ExportCFONB then
      TOBL.PutValue('E_CFONBOK',   '#');
    // FIN VL

    TOBL.PutValue('E_NUMENCADECA', ParmsMP.NumEncaDeca) ;
    TOBL.PutValue('E_MODESAISIE' , '-') ;
    {JP 09/11/05 : FQ 15771 : Il ne faut pas mettre E_NUMECHE systématiquement à 1, notamment
                   lorsque le compte n'est pas lettrable : D'ailleurs, la ligne est-elle bien
                   utile, car me semble-t-il les infos de lettrage ont déjà été traitées ? }
    if TOBL.GetValue('E_ETATLETTRAGE')<>'RI' then
      TOBL.PutValue('E_NUMECHE' 	 , 1) ;

    // Uniformisations
    TOBL.PutValue('E_NATUREPIECE',  lStNatPiece ) ;   // Ajout SBO 28/07/2004
    TOBL.PutValue('E_DATECREATION', Date ) ;          // Ajout SBO 28/07/2004
    TOBL.PutValue('E_DATEMODIF',    lDtDateModif ) ;  // Ajout SBO 28/07/2004

    // Correction recopie Code gescom
    TOBL.PutValue('E_REFGESCOM',    '') ;
    TOBL.PutValue('E_VALIDE',       '-') ;
    AffecteG(TOBL, ParmsMP) ;

  end;
end ;

procedure FinirAnas ( TOBL : TOB ; ParmsMP : TGenereMP ) ;
Var i,NumA                  : integer ;
    TOBA,TOBAxe             : TOB ;
    EDP,ECP,EDD,ECD : Double ;  // Montants écritures
    TDP,TCP,TDD,TCD : double ;  // Totaux montants
    XD,XC                   : double ;  // Soldes
begin
  if TOBL.Detail.Count<=0 then Exit ;
  EDP := TOBL.GetValue('E_DEBIT')     ;
  ECP := TOBL.GetValue('E_CREDIT') ;
  EDD := TOBL.GetValue('E_DEBITDEV')  ;
  ECD := TOBL.GetValue('E_CREDITDEV') ;
  for NumA:=0 to TOBL.Detail.Count-1 do
    BEGIN
    TOBAxe := TOBL.Detail[NumA] ;
    TDP:=0 ; TCP:=0 ;
    TDD:=0 ; TCD:=0 ;
    for i:=0 to TOBAxe.Detail.Count-1 do
      BEGIN
      TOBA:=TOBAxe.Detail[i] ;
      if ParmsMP.Ref1<>'' then
        TOBA.PutValue('Y_REFINTERNE',   ParmsMP.Ref1) ;
      if ParmsMP.Lib1<>'' then
        TOBA.PutValue('Y_LIBELLE',      ParmsMP.Lib1) ;
      TOBA.PutValue('Y_TOTALECRITURE',  Arrondi(EDP+ECP,6) ) ;
      TOBA.PutValue('Y_TOTALDEVISE',    Arrondi(EDD+ECD,6) ) ;
      if i < TOBAxe.Detail.Count-1 then
        begin
        // Pivot
        XD:=Arrondi(EDP*TOBA.GetValue('Y_POURCENTAGE')/100.0,V_PGI.OkDecV) ;
        XC:=Arrondi(ECP*TOBA.GetValue('Y_POURCENTAGE')/100.0,V_PGI.OkDecV) ;
        TDP:=Arrondi(TDP+XD,6) ;
        TCP:=Arrondi(TCP+XC,6) ;
        TOBA.PutValue('Y_DEBIT',XD) ;
        TOBA.PutValue('Y_CREDIT',XC) ;
        // Devise
        XD:=Arrondi(EDD*TOBA.GetValue('Y_POURCENTAGE')/100.0,ParmsMP.DEV.Decimale) ;
        XC:=Arrondi(ECD*TOBA.GetValue('Y_POURCENTAGE')/100.0,ParmsMP.DEV.Decimale) ;
        TDD:=Arrondi(TDD+XD,6) ;
        TCD:=Arrondi(TCD+XC,6) ;
        TOBA.PutValue('Y_DEBITDEV',  XD) ;
        TOBA.PutValue('Y_CREDITDEV', XC) ;
        end
      else
        begin
        // Pivot
        XD:=Arrondi(EDP-TDP,V_PGI.OkDecV) ;
        XC:=Arrondi(ECP-TCP,V_PGI.OkDecV) ;
        TOBA.PutValue('Y_DEBIT',  XD) ;
        TOBA.PutValue('Y_CREDIT', XC) ;
        // Devise
        XD:=Arrondi(EDD-TDD,ParmsMP.DEV.Decimale) ;
        XC:=Arrondi(ECD-TCD,ParmsMP.DEV.Decimale) ;
        TOBA.PutValue('Y_DEBITDEV',  XD) ;
        TOBA.PutValue('Y_CREDITDEV', XC) ;
        end ;
      end ;
    end ;
end ;

Function  VerifCrits ( TOBL : TOB ; ValsRupt : String ; ParmsMP : TGenereMP ) : boolean ;
Var Sts,NomChamp,ValChamp,StV : String ;
begin
  Result:=True ;
  if TOBL=Nil then Exit ;
  if ParmsMP.ChampsRupt='' then Exit ;
  if ValsRupt='' then Exit ;
  Result:=False ;
  StS := ParmsMP.ChampsRupt ;
  StV := ValsRupt ;
  Repeat
    NomChamp := ReadTokenSt(Sts) ;
    ValChamp := ReadTokenSt(StV) ;
    if NomChamp<>'' then
      if TOBL.GetValue(NomChamp)<>ValChamp then Exit ;
  Until ((Sts='') or (NomChamp='')) ;
  Result:=True ;
end ;

procedure DupliquerTOBsmp (  sBqe, sCat, ValsRupt : String ;
                             TOBOrig, TOBDest     : TOB ;
                             ParmsMP              : TGenereMP  ) ;
Var i : integer ;
    TOBL,TOBN : TOB ;
    lBQE,lCat : String ;
begin
  if ParmsMP.smp in [smpEncTous,smpDecTous] then
    // duplication fine (??)
    for i:=0 to TOBOrig.Detail.Count-1 do
      begin
      TOBL := TOBOrig.Detail[i] ;
      lBQE := TOBL.GetValue('E_BANQUEPREVI') ;
      lCat := MPToCategorie(TOBL.GetValue('E_MODEPAIE')) ;
      if ((lBqe=sBqe) and (lCat=sCat)) then
        if VerifCrits(TOBL,ValsRupt,ParmsMP) then
          begin
          TOBN := TOB.Create( 'ECRITURE', TOBDest, -1 ) ;
          TOBN.Dupliquer(TOBL,True,True,False) ;
          end ;
      end
  else
    if ValsRupt<>'' then
      // duplication moyenne (??)
      for i:=0 to TOBOrig.Detail.Count-1 do
        begin
        TOBL := TOBOrig.Detail[i] ;
        if VerifCrits(TOBL,ValsRupt,ParmsMP) then
          begin
          TOBN := TOB.Create('ECRITURE',TOBDest,-1) ;
          TOBN.Dupliquer(TOBL,True,True,False) ;
          end ;
        end
    else
      // grosse duplication (...)
      TOBDest.Dupliquer(TOBOrig,True,True,False) ;
end ;

Procedure RecalcLigne ( TOBL : TOB ; ParmsMP : TGenereMP ) ;
Var DD,CD,DP,CP      : Double ;
    XP               : Double ;
    TOBA,TOBAx       : TOB ;
    i,k              : integer ;
begin
  DD := TOBL.GetValue('E_DEBITDEV') ;
  CD := TOBL.GetValue('E_CREDITDEV') ;
  DP := DeviseToEuro(DD,ParmsMP.DEV.Taux,ParmsMP.DEV.Quotite) ;
  CP := DeviseToEuro(CD,ParmsMP.DEV.Taux,ParmsMP.DEV.Quotite) ;
  TOBL.PutValue('E_DEBIT',      DP) ;
  TOBL.PutValue('E_CREDIT',     CP) ;
  XP := DP + CP ;
  for k:=0 to TOBL.Detail.Count-1 do
    BEGIN
    TOBAx := TOBL.Detail[k] ;
    for i:=0 to TOBAx.Detail.Count-1 do
      BEGIN
      TOBA := TOBAx.Detail[i] ;
      DD:=TOBA.GetValue('Y_DEBITDEV') ; CD:=TOBA.GetValue('Y_CREDITDEV') ;
      DP := DeviseToEuro(DD,ParmsMP.DEV.Taux,ParmsMP.DEV.Quotite) ;
      CP := DeviseToEuro(CD,ParmsMP.DEV.Taux,ParmsMP.DEV.Quotite) ;
      TOBA.PutValue('Y_DEBIT',        DP) ;
      TOBA.PutValue('Y_CREDIT',       CP) ;
      TOBA.PutValue('Y_TOTALECRITURE',XP) ;
      END ;
    END ;
end ;

procedure PassageDevise( TOBDest : TOB ; ParmsMP : TGenereMP ) ;
Var TOBL : TOB ;
    sDev : String ;
    i    : integer ;
    Cote : Double ;
begin
  if TOBDest.Detail.Count <= 0 then Exit ;
  TOBL := TOBDest.Detail[0] ;
  sDev := TOBL.GetValue('E_DEVISE') ;
  if ((sDev=V_PGI.DevisePivot) or (sDev=V_PGI.DeviseFongible)) then Exit ;
  if V_PGI.TauxEuro<>0
    then Cote := ParmsMP.DEV.Taux / V_PGI.TauxEuro
    else Cote := 1 ;
  for i:=0 to TOBDest.Detail.Count-1 do
    BEGIN
    TOBL := TOBDest.Detail[i] ;
    TOBL.PutValue('E_TAUXDEV',  ParmsMP.DEV.Taux ) ;
    TOBL.PutValue('E_COTATION', Cote ) ;
    RecalcLigne( TOBL, ParmsMP ) ;
    END ;
end ;

procedure ConstitueDest ( sBQE, sCat, ValsRupt : String ; TOBOrig, TOBDest : TOB ; ParmsMP : TGenereMP ) ;
Var TOBR,TOBL,TOBDETR,TOBA,TOBAxe,TOBLocAxe,TOBLOCEcr : TOB ;
    i,Lasti,LastRupt,k,ia,iloc, iMemeTiers : integer ;
    DP,CP,DD,CD,XD,XC : double ;
    TypG,OldAux,OldGen,St, szOldCodeAccept : String ;
    OldEche : TDateTime ;
    Rupt,Premier, bCodeAccept : boolean ;
    YY           : RMVT ;
    StSort : String ;
begin
  // Initialisation des variables locales
	Lasti := 0 ;
  OldEche := iDate1900 ;
  // Duplication
  TOBDest.ClearDetail ;
  DupliquerTOBsmp( sBqe, sCat, ValsRupt, TOBOrig, TOBDest, ParmsMP ) ;
  // Méthode de tri
  TypG := ParmsMP.GroupeEncadeca ;
  {$IFDEF CEGID}
  if TypG='DET' then StSort:='E_AUXILIAIRE;E_GENERAL;E_DATEECHEANCE;E_NUMEROPIECE' else
   if TypG='GLO' then StSort:='E_AUXILIAIRE;E_GENERAL;E_DATEECHEANCE;E_NUMEROPIECE' else
    if TypG='AUX' then StSort:='E_AUXILIAIRE;E_GENERAL;E_DATEECHEANCE;E_NUMEROPIECE' else
     if TypG='ECT' then StSort:='E_DATEECHEANCE;E_AUXILIAIRE;E_GENERAL;E_NUMEROPIECE' else
      if TypG='ECH' then StSort:='E_AUXILIAIRE;E_GENERAL;E_DATEECHEANCE;E_NUMEROPIECE' ;
  {$ELSE}
  If EstSerie(S3) Then
    BEGIN
    if TypG='DET' then StSort:='E_AUXILIAIRE;E_GENERAL;E_DATEECHEANCE;E_NUMEROPIECE' else
     if TypG='GLO' then StSort:='E_AUXILIAIRE;E_GENERAL;E_DATEECHEANCE;E_NUMEROPIECE' else
      if TypG='AUX' then StSort:='E_AUXILIAIRE;E_GENERAL;E_DATEECHEANCE;E_NUMEROPIECE' else
       if TypG='ECT' then StSort:='E_DATEECHEANCE;E_AUXILIAIRE;E_GENERAL;E_NUMEROPIECE' else
        if TypG='ECH' then StSort:='E_AUXILIAIRE;E_GENERAL;E_DATEECHEANCE;E_NUMEROPIECE' ;
    END Else
    BEGIN
    if TypG='DET' then StSort:='E_BANQUEPREVI;E_NOMLOT;E_AUXILIAIRE;E_GENERAL;E_DATEECHEANCE;E_NUMEROPIECE' else
     if TypG='GLO' then StSort:='E_BANQUEPREVI;E_NOMLOT;E_AUXILIAIRE;E_GENERAL;E_DATEECHEANCE;E_NUMEROPIECE' else
      if TypG='AUX' then StSort:='E_BANQUEPREVI;E_NOMLOT;E_AUXILIAIRE;E_GENERAL;E_DATEECHEANCE;E_NUMEROPIECE' else
       if TypG='ECT' then StSort:='E_BANQUEPREVI;E_NOMLOT;E_DATEECHEANCE;E_AUXILIAIRE;E_GENERAL;E_NUMEROPIECE' else
        if TypG='ECH' then StSort:='E_BANQUEPREVI;E_NOMLOT;E_AUXILIAIRE;E_GENERAL;E_DATEECHEANCE;E_NUMEROPIECE' ;
    END ;
  {$ENDIF}
  TOBDest.Detail.Sort(StSort) ;
  PassageDevise( TOBDest , ParmsMP ) ;


// Mise en place de la globalisation des tiers ici...
// ...En attente...

  TOBR := TOB.Create('',Nil,-1) ;
  DP:=0 ; CP:=0 ;
  DD:=0 ; CD:=0 ;
  Rupt    := False ;
  Premier := True ;
  bCodeAccept := True;
  LastRupt:= 0 ;
  iMemeTiers := 1;
  for i:=0 to TOBDest.Detail.Count-1 do
    begin
    TOBL := TOBDest.Detail[i] ;
    Lasti:=i ;
    // Sauvegarde info origine
    YY := TOBToIdent(TOBL,True) ;
    St := EncodeLC(YY) ;
    TOBL.PutValue('E_TRACE',St) ;

    // Traitement
    if Premier then
      BEGIN
      OldAux := TOBL.GetValue('E_AUXILIAIRE') ;
      OldGen := TOBL.GetValue('E_GENERAL') ;
      OldEche:= TOBL.GetValue('E_DATEECHEANCE') ;
      szOldCodeAccept:= TOBL.GetValue('E_CODEACCEPT');
      iMemeTiers := 1;
      END ;
    if TypG='DET' then Rupt:=(i>0) else
     if TypG='GLO' then Rupt:=False else
      if TypG='AUX' then Rupt:=((OldAux<>TOBL.GetValue('E_AUXILIAIRE')) or (OldGen<>TOBL.GetValue('E_GENERAL'))) else
       if TypG='ECT' then Rupt:=(OldEche<>TOBL.GetValue('E_DATEECHEANCE')) else
        if TypG='ECH' then Rupt:=((OldAux<>TOBL.GetValue('E_AUXILIAIRE')) or (OldGen<>TOBL.GetValue('E_GENERAL')) or (OldEche<>TOBL.GetValue('E_DATEECHEANCE'))) ;
    Premier:=False ;
    if Rupt then
      BEGIN
      TOBDETR := TOB.Create('ECRITURE',TOBR,-1) ;
      TOBDETR.Dupliquer( TOBDest.Detail[i-1], False, True, False) ;
      TOBDETR.AddChampSup('INDICE',False) ;
      TOBDETR.PutValue('INDICE',i) ;
      TOBDETR.AddChampSupValeur('GROUPEIDX',LastRupt) ;
      if ((DP<>0) and (CP<>0)) then
        BEGIN
        if Abs(DP)>Abs(CP) then
          BEGIN
          DP := DP-CP ; CP:=0 ;
          DD := DD-CD ; CD:=0 ;
          END
        else
          BEGIN
          CP := CP-DP ; DP:=0 ;
          CD := CD-DD ; DD:=0 ;
          END ;
        END ;
      // FQ 12438
      if (iMemeTiers > 1) and (not bCodeAccept) then
        TOBDETR.PutValue('E_CODEACCEPT', MPTOACC(TOBDETR.GetValue('E_MODEPAIE')) );
      TOBDETR.PutValue('E_DEBIT',     DP) ;
      TOBDETR.PutValue('E_CREDIT',    CP) ;
      TOBDETR.PutValue('E_DEBITDEV',  DD) ;
      TOBDETR.PutValue('E_CREDITDEV', CD) ;

      // Reproduire anals
      if CoherAna(LastRupt,i-1, TOBDest, ParmsMP ) then
        BEGIN
        for k:=LastRupt to i-1 do
          BEGIN
          TOBLOCEcr := TOBDest.Detail[k] ;
          if ((TOBLOCEcr.Detail.Count>0) and (TOBDETR.Detail.Count<=0)) then
            AlloueAxe(TobDeTr);

          for ia:=0 to TOBLOCEcr.Detail.Count-1 do
            BEGIN
            TOBLocAxe := TOBLOCEcr.Detail[ia] ;
            TOBAxe    := TOBDETR.Detail[ia] ;
            for iloc:=0 to TOBLOCAxe.Detail.Count-1 do
              BEGIN
              TOBA:=TOB.Create('ANALYTIQ',TOBAxe,-1) ;
              TOBA.Dupliquer(TOBLOCAxe.Detail[iloc],False,True,False) ;
              END ;
            END ;
          END ;
        END ;
      // Tripotage
      TripoteDetail( TOBDETR, ParmsMP ) ;
      DP:=0 ; CP:=0 ;
      DD:=0 ; CD:=0 ;
      {JP 26/01/05 : FQ 17284 : Suppression du -1. En effet, si LastRupt = i - 1, la ventilation de la dernière
                     ligne d'une rupture est reprise dans la ventilation de la rupture suivante cf CoherAna et
                     for k:=LastRupt to i-1 do.
                     Par exemple : TOBL : lig 0 : Tiers1 - Eche1 - Ventil1
                                          lig 1 : Tiers1 - Eche1 - Ventil2
                                          lig 2 : Tiers2 - Eche1 - Ventil1
                     le Traitement de l'analytique doit se faire de 0 à 1 et de 2 à 2 alors qu'en mettant LastRupt
                     à i - 1, on faisait le traitement de 0 à 1 et de 1 à 2 !!!}
      LastRupt := i; //- 1;
      iMemeTiers := 1;
      END
    // FQ 12438
    else begin
      inc(iMemeTiers);
      bCodeAccept := (szOldCodeAccept = TOBL.GetValue('E_CODEACCEPT'));
    end;

    // Indicateur d'annulation du traitement pour ce tiers
    TOBL.AddChampSupValeur('GROUPEIDX', LastRupt) ;

    XD := GetLeMontant( TOBL.GetValue('E_DEBIT'), TOBL.GetValue('E_COUVERTURE'), TOBL.GetValue('E_SAISIMP'), tsmPivot) ;
    XC := GetLeMontant( TOBL.GetValue('E_CREDIT'), TOBL.GetValue('E_COUVERTURE'), TOBL.GetValue('E_SAISIMP'), tsmPivot) ;
    DP := Arrondi( DP+XD, V_PGI.OkDecV) ;
    CP := Arrondi( CP+XC, V_PGI.OkDecV) ;

    XD := GetLeMontant( TOBL.GetValue('E_DEBITDEV'), TOBL.GetValue('E_COUVERTUREDEV'), TOBL.GetValue('E_SAISIMP'), tsmDevise) ;
    XC := GetLeMontant( TOBL.GetValue('E_CREDITDEV'), TOBL.GetValue('E_COUVERTUREDEV'), TOBL.GetValue('E_SAISIMP'), tsmDevise) ;
    DD := Arrondi( DD+XD, ParmsMP.DEV.Decimale) ;
    CD := Arrondi( CD+XC, ParmsMP.DEV.Decimale) ;

    OldAux := TOBL.GetValue('E_AUXILIAIRE') ;
    OldGen := TOBL.GetValue('E_GENERAL') ;
    OldEche:= TOBL.GetValue('E_DATEECHEANCE') ;
    szOldCodeAccept:= TOBL.GetValue('E_CODEACCEPT');
    END ;

  if Not Premier then
    BEGIN
    TOBDETR:=TOB.Create('ECRITURE',TOBR,-1) ;
    TOBDETR.Dupliquer(TOBDest.Detail[Lasti],False,True,False) ;
    TOBDETR.AddChampSup('INDICE',False) ;
    TOBDETR.PutValue('INDICE',-1) ;
    TOBDETR.AddChampSupValeur('GROUPEIDX',LastRupt) ;
    if ((DP<>0) and (CP<>0)) then
      BEGIN
      if Abs(DP)>Abs(CP) then
        BEGIN
        DP := DP-CP ; CP:=0 ;
        DD := DD-CD ; CD:=0 ;
        END
      else
        BEGIN
        CP := CP-DP ; DP:=0 ;
        CD := CD-DD ; DD:=0 ;
        END ;
      END ;
    // FQ 12438
    if (iMemeTiers > 1) and (not bCodeAccept) then
      TOBDETR.PutValue('E_CODEACCEPT', MPTOACC(TOBDETR.GetValue('E_MODEPAIE')) );
    TOBDETR.PutValue('E_DEBIT',     DP) ;
    TOBDETR.PutValue('E_CREDIT',    CP) ;
    TOBDETR.PutValue('E_DEBITDEV',  DD) ;
    TOBDETR.PutValue('E_CREDITDEV', CD) ;

    // Reproduire anals
    for k:=LastRupt to TOBDest.Detail.Count-1 do
      BEGIN
      // Ligne ECR de destination issue de la source
      TOBLOCEcr := TOBDest.Detail[k] ;
      // Si pas de TOB intermédiaire des axes, les créer
      if ((TOBLOCEcr.Detail.Count>0) and (TOBDETR.Detail.Count<=0)) then
        AlloueAxe(TobDeTr);

      for ia:=0 to TOBLOCEcr.Detail.Count-1 do
        BEGIN
        TOBLocAxe := TOBLOCEcr.Detail[ia] ;
        TOBAxe    := TOBDETR.Detail[ia] ;
        for iloc:=0 to TOBLOCAxe.Detail.Count-1 do
          BEGIN
          // Duplication des analytiques des origines sur la nouvelle ligne
          TOBA := TOB.Create('ANALYTIQ',TOBAxe,-1) ;
          TOBA.Dupliquer(TOBLOCAxe.Detail[iloc],False,True,False) ;
          END ;
        END ;
      END ;
    TripoteDetail( TOBDETR, ParmsMP ) ;
    END ;
  FinirDestination( TOBDest, TOBR, ParmsMP ) ;   ;
  TOBR.Free ;
end ;

function  CoherAna ( ideb, ifin : integer ; TOBDest : TOB ; ParmsMP : TGenereMP ) : boolean ;
var TOBL,TOBAxe : TOB ;
    i,ia,NumAxe : integer ;
    BB : Array[1..MaxAxe] of Boolean ;
begin
  Result := False ;
  for i := ideb to ifin do
    begin
    TOBL := TOBDest.Detail[i] ;
    FillChar(BB,Sizeof(BB),#0) ;
    for NumAxe := 0 to TOBL.Detail.Count-1 do
        begin
        TOBAxe := TOBL.Detail[NumAxe] ;
        if TOBAxe.Detail.Count>0 then
          BB[NumAxe+1] := True ;
        end ;
    for ia:=1 to MaxAxe do
      if BB[ia]<>ParmsMP.Ventils[ia] then Exit ;
    end ;
  Result := True ;
end ;

procedure AjouteLesAnas( TOBOrig : TOB ) ;
Var i,k,j       : integer ;
    TOBL, NewA  : TOB ;
    TOBAxe,DetA : TOB ;
    SQL         : String ;
    QQ          : TQuery ;
begin
  NewA := TOB.Create('',Nil,-1) ;
  for i := 0 to TOBOrig.Detail.Count-1 do
    BEGIN
    TOBL := TOBOrig.Detail[i] ;
    if TOBL.GetValue('E_ANA')<>'X' then Continue ;
    SQL := 'SELECT * FROM ANALYTIQ WHERE Y_JOURNAL="'+TOBL.GetValue('E_JOURNAL')+'" '
        +'AND Y_EXERCICE="'+TOBL.GetValue('E_EXERCICE')+'" AND Y_DATECOMPTABLE="'+UsDateTime(TOBL.GetValue('E_DATECOMPTABLE'))+'" '
        +'AND Y_NUMEROPIECE='+IntToStr(TOBL.GetValue('E_NUMEROPIECE'))+' AND Y_NUMLIGNE='+IntToStr(TOBL.GetValue('E_NUMLIGNE'))+' '
        +'AND Y_QUALIFPIECE="N" ORDER BY Y_AXE, Y_NUMVENTIL' ;
    QQ := OpenSQL(SQL,True) ;
    if Not QQ.EOF then
      NewA.LoadDetailDB('ANALYTIQ','','',QQ,False,True) ;
    Ferme(QQ) ;

    AlloueAxe(TobL);
    for k:=1 to MaxAxe do
      BEGIN
      TOBAxe := TobL.Detail[k - 1];
      for j := NewA.Detail.Count - 1 downto 0 do
        BEGIN
        DetA := NewA.Detail[j] ;
        if DetA.GetValue('Y_AXE')='A'+IntToStr(k) then
          DetA.ChangeParent(TOBAxe,0) ;
        END ;
      END ;
    END ;
  NewA.Free ;
end ;

procedure VireInutiles ( TOBOrig : TOB; swapSelect : Boolean = False);
Var i : integer ;
    TOBL : TOB ;
begin
  for i:=TOBOrig.Detail.Count-1 downto 0 do
    begin
    TOBL := TOBOrig.Detail[i] ;
    if swapSelect then begin
      if TOBL.GetValue('MARQUE')='X' then TOBL.Free;
      end
    else
      if TOBL.GetValue('MARQUE')<>'X' then TOBL.Free ;

    end ;
end ;

procedure AnasRupt ( TOBL : TOB ; ParmsMP : TGenereMP ) ;
Var ModeA  : String ;
begin
  if TOBL.GetValue('E_ANA')<>'X' then Exit ;
  if ((ParmsMP.ModeAnal='VEN') or (ParmsMP.ModeAnal='ATT') or (TOBL.Detail.Count<=0)) then
    begin
    //SG6 28.02.05
    TOBL.ClearDetail;

    if ((TOBL.Detail.Count<=0) and (ParmsMP.ModeAnal='DET'))
      then ModeA := 'ATT'
      else ModeA := ParmsMP.ModeAnal ;
    if TOBL.Detail.Count<=0 then
      AlloueAxe(TobL);

    VentilerTOB( TOBL, '', 0, ParmsMP.DEV.Decimale, (ModeA='ATT') ) ;
    end
  else
    begin
    {Report en détail groupé par section}
    RegroupeAnas(TOBL) ;
    end ;
end ;

procedure RegroupeAnas ( TOBL : TOB ) ;
Var i, k, NumA        : integer ;
    TOBA, NewA, DetA  : TOB ;
    TOBAxe, NewAxe    : TOB ;
    Debit, CasPart    : boolean ;
    Section           : String ;
    X,Pourc           : Double ;
    DP,CP,DD,CD : Double ;
    TotEcrD,TotEcrP   : Double ;
begin
  // Tob mère contenant les nouvelles ventilations analytituqes par Axes
  NewA    := TOB.Create('',Nil,-1) ;
  // Cré les axes analytiques dans la tob NewA
  AlloueAxe(NewA);

  {TobDebug(TobL);}

  TotEcrP := TOBL.GetValue('E_DEBIT')     + TOBL.GetValue('E_CREDIT')     ;
  TotEcrD := TOBL.GetValue('E_DEBITDEV')  + TOBL.GetValue('E_CREDITDEV')  ;
  // Parcours tous les axes analytiques
  for NumA:=0 to TOBL.Detail.Count-1 do
    begin
    TOBAxe := TOBL.Detail[NumA] ;

    NewAxe := NewA.Detail[NumA] ;
    // Parcours toutes les ventilations de l'axe en cours
    for i:=0 to TOBAxe.Detail.Count-1 do
      begin
      TOBA    := TOBAxe.Detail[i] ;
      Section := TOBA.GetValue('Y_SECTION') ;
      // Si on ne trouve pas la section en cours dans une précédente TOB
      DetA    := NewAxe.FindFirst(['Y_SECTION'],[Section],False) ;
      // La crée
      if DetA=Nil then
        begin
        DetA := TOB.Create('ANALYTIQ', NewAxe, -1) ;
        DetA.Dupliquer(TOBA,False,True) ;
        DetA.PutValue('Y_DEBIT',      0) ;
        DetA.PutValue('Y_CREDIT',     0) ;
        DetA.PutValue('Y_DEBITDEV',   0) ;
        DetA.PutValue('Y_CREDITDEV',  0) ;
        end ;
      // Additionne les montant des mêmes sections
      X := DetA.GetValue('Y_DEBIT') ;
      X := X + TOBA.GetValue('Y_DEBIT') ;
      DetA.PutValue('Y_DEBIT',      X) ;
      X := DetA.GetValue('Y_CREDIT') ;
      X := X + TOBA.GetValue('Y_CREDIT') ;
      DetA.PutValue('Y_CREDIT',     X) ;
      X := DetA.GetValue('Y_DEBITDEV') ;
      X := X + TOBA.GetValue('Y_DEBITDEV') ;
      DetA.PutValue('Y_DEBITDEV',   X) ;
      X := DetA.GetValue('Y_CREDITDEV') ;
      X := X + TOBA.GetValue('Y_CREDITDEV') ;
      DetA.PutValue('Y_CREDITDEV',  X) ;
      end ;
    end ;
  TOBL.ClearDetail ;
  Debit := TOBL.GetValue('E_DEBIT') <> 0 ;

  // Parcours tous les axes analytiques
  AlloueAxe(TobL);
  for i:=1 to MaxAxe do
    begin
    TOBAxe := TobL.Detail[i - 1];
    for NumA := 0 to NewA.Detail.Count-1 do
      begin
      NewAxe := NewA.Detail[NumA] ;

      // Si ce n'est pas la même table : Passe au suivant
      if NewAxe.GetString('Y_AXE') <> TOBAxe.GetString('Y_AXE') then Continue ;

      // Parcours les écritures
      for k:=NewAxe.Detail.Count-1 downto 0 do
        begin
        DetA := NewAxe.Detail[k] ;
        DetA.ChangeParent(TOBAxe, 0) ;

        CasPart:=False ;
        DP := DetA.GetValue('Y_DEBIT') ;
        CP := DetA.GetValue('Y_CREDIT') ;
        DD := DetA.GetValue('Y_DEBITDEV') ;
        CD := DetA.GetValue('Y_CREDITDEV') ;
        if ((Debit) and (CP<>0)) then
          begin
          DP := DP-CP ; CP:=0 ;
          DD := DD-CD ; CD:=0 ;
          CasPart := True ;
          end
        else
          if ((Not Debit) and (DP<>0)) then
            begin
            CP := CP-DP ; DP:=0 ;
            CD := CD-DD ; DD:=0 ;
            CasPart := True ;
            end ;
            if CasPart then
               begin
               DetA.PutValue('Y_DEBIT',     DP) ;
               DetA.PutValue('Y_CREDIT',    CP) ;
               DetA.PutValue('Y_DEBITDEV',  DD) ;
               DetA.PutValue('Y_CREDITDEV', CD) ;
               end ;

        DetA.PutValue('Y_TOTALECRITURE',  TotEcrP) ;
        DetA.PutValue('Y_TOTALDEVISE',    TotEcrD) ;
        Pourc := 0 ;
        if TotEcrP<>0 then
          Pourc := Arrondi( 100.0*( DetA.GetValue('Y_DEBIT') + DetA.GetValue('Y_CREDIT') ) / TotEcrP, 6 ) ;
        if Pourc <> 0 then DetA.PutValue('Y_POURCENTAGE', Pourc);
        end ;
      end ;
    end ;
  NewA.Free ;
end ;

procedure InitEscompteSup ( TOBE : TOB ) ;
begin
//  if not GereEscompte then Exit ; // Test a faire avant appel
  TOBE.AddChampSup('ESCOMPTABLE',False) ; TOBE.PutValue('ESCOMPTABLE','X') ;
  TOBE.AddChampSup('COMPTEHT',False)    ; TOBE.PutValue('COMPTEHT','') ;
  TOBE.AddChampSup('COMPTETVA',False)   ; TOBE.PutValue('COMPTETVA','') ;
  TOBE.AddChampSup('TAUXESC',False)     ; TOBE.PutValue('TAUXESC',0.00) ;
  TOBE.AddChampSup('TAUXTVA',False)     ; TOBE.PutValue('TAUXTVA',0.00) ;
  TOBE.AddChampSup('MONTANTESC',False)  ; TOBE.PutValue('MONTANTESC',0.00) ;
// Ajout Champ SANSESCOMPTE
	TOBE.AddChampSup('SANSESCOMPTE',False); TOBE.PutValue('SANSESCOMPTE','-') ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 27/11/2002
Modifié le ... :   /  /
Description .. : Recherche la tob escompte correspondant à la ligne dont la
Suite ........ : tob est passé en paramètre.
Suite ........ : Création en option
Mots clefs ... :
*****************************************************************}
Function  FindCreerTOBEscomptePourLigne ( TOBEscomptes, TOBLigne : TOB ; QueFind : boolean ) : TOB ;
Var TOBE            : TOB ;
    Jal,Exo         : String ;
    NumP,NumL,NumE  : integer ;
begin
  Result:=Nil ;
	if (TOBLigne = nil) then Exit ;

  Jal		:= TOBLigne.GetValue('E_JOURNAL') ;
	Exo		:= TOBLigne.GetValue('E_EXERCICE') ;
	NumP	:= TOBLigne.GetValue('E_NUMEROPIECE') ;
	NumL	:= TOBLigne.GetValue('E_NUMLIGNE') ;
	NumE	:= TOBLigne.GetValue('E_NUMECHE') ;
	TOBE	:= TOBEscomptes.FindFirst(['E_JOURNAL','E_EXERCICE','E_NUMEROPIECE','E_NUMLIGNE','E_NUMECHE'],
  																[Jal,Exo,NumP,NumL,NumE],False) ;
	if ((TOBE=Nil) and (Not QueFind)) then
    BEGIN
    TOBE:=TOB.Create('ESCOMPTE',TOBEscomptes,-1) ;
    {Infos ecriture}
    TOBE.AddChampSup('E_JOURNAL',False)     ; TOBE.PutValue('E_JOURNAL',Jal) ;
    TOBE.AddChampSup('E_EXERCICE',False)    ; TOBE.PutValue('E_EXERCICE',Exo) ;
    TOBE.AddChampSup('E_NUMEROPIECE',False) ; TOBE.PutValue('E_NUMEROPIECE',NumP) ;
    TOBE.AddChampSup('E_NUMLIGNE',False)    ; TOBE.PutValue('E_NUMLIGNE',NumL) ;
    TOBE.AddChampSup('E_NUMECHE',False)     ; TOBE.PutValue('E_NUMECHE',NumE) ;
    {infos Escompte}
    InitEscompteSup(TOBE) ;
    END ;
	Result:=TOBE ;
end ;

Function ExtractDiscerne ( TOBOrig : TOB ; sBqe, sCat, ValsRupt : String ; ParmsMP : TGenereMP ) : TOB ;
Var TOB2, TOBE, TOBE2 : TOB ;
    i                 : integer ;
begin
  TOB2:=Nil ;
  for i:=0 to TOBOrig.Detail.Count-1 do
    begin
    TOBE:=TOBOrig.Detail[i] ;
    if ((sBqe<>'') and (TOBE.GetValue('E_BANQUEPREVI')<>sBqe)) then Continue ;
    if ((scat<>'') and (MPToCategorie(TOBE.GetValue('E_MODEPAIE'))<>sCat)) then Continue ;
    if Not VerifCrits( TOBE, ValsRupt, ParmsMP ) then Continue ;
    if TOB2 = Nil then
      TOB2 := TOB.Create('',Nil,-1) ;
    TOBE2 := TOB.Create('ECRITURE',TOB2,-1) ;
    TOBE2.Dupliquer(TOBE,True,True) ;
    end ;
  Result := TOB2 ;
end ;

Procedure CompleteMMRefCCMP( ParmsMP : TGenereMP ; Var FR : tFormuleRefCCMP ) ;
begin
  If not (ParmsMP.smp in [smpDecChqEdt,smpDecChqEdtNC]) Then
    begin
    FR.Ref1 := ParmsMP.Ref1 ;
    FR.Ref2 := ParmsMP.Ref2 ;
    end ;
  FR.Lib1     := ParmsMP.Lib1 ;
  FR.Lib2     := ParmsMP.Lib2 ;
  FR.RefExt1  := ParmsMP.RefExt1 ;
  FR.RefLib1  := ParmsMP.RefLib1 ;
  FR.RefExt2  := ParmsMP.RefExt2 ;
  FR.RefLib2  := ParmsMP.RefLib2 ;
end ;

Procedure CompleteMM ( ParmsMP : TGenereMP ; SorteLettre : TSorteLettre ;
                       vBoSpooler, vBoFichier : Boolean ; vStRepSpooler : String ;
                       Var MM : RMVT ) ;
 begin
  MM.smp         := ParmsMP.smp ;
  MM.NumEncaDeca := ParmsMP.NumEncaDeca ;   //RRO le 05/02/2003 - Dans tous les cas -
  Case ParmsMP.smp of
    smpEncTraEdt,smpEncTraEdtNC :
                 BEGIN
                 MM.SorteLettre := SorteLettre ;
                 MM.Globalise   := ( ParmsMP.GroupeEncaDeca = 'GLO' ) ; // SBO : Correction Fiche 12351 (remplace MM.Globalise := False ;)
                 MM.Section     := ParmsMP.CptG ;
                 CompleteMMRefCCMP( ParmsMP, MM.FormuleRefCCMP ) ;
                 END ;
    smpEncPreEdt,smpEncPreEdtNC : BEGIN
                 MM.SorteLettre := SorteLettre ;
                 MM.Globalise   := ( ParmsMP.GroupeEncaDeca = 'GLO' ) ; // SBO : Correction Fiche 12351 (remplace MM.Globalise := False ;)
                 MM.General     := ParmsMP.CptG ;
                 CompleteMMRefCCMP( ParmsMP, MM.FormuleRefCCMP ) ;
                 END ;
    smpDecChqEdt,smpDecChqEdtNC : BEGIN
                 MM.SorteLettre := SorteLettre ;
                 MM.Globalise   := ( ParmsMP.GroupeEncaDeca = 'GLO' ) ; // SBO : Correction Fiche 12351 (remplace MM.Globalise := False ;)
                 MM.General     := ParmsMP.CptG ;
                 CompleteMMRefCCMP( ParmsMP, MM.FormuleRefCCMP ) ;
                 END ;
    smpDecVirEdt,smpDecVirEdtNC,smpDecVirInEdt,smpDecVirInEdtNC : BEGIN
                 MM.SorteLettre := SorteLettre ;
                 MM.Globalise   := ( ParmsMP.GroupeEncaDeca = 'GLO' ) ; // SBO : Correction Fiche 12351 (remplace MM.Globalise := False ;)
                 MM.General     := ParmsMP.CptG ;
                 CompleteMMRefCCMP( ParmsMP, MM.FormuleRefCCMP ) ;
                 END ;
    smpDecBorEdt,smpDecBorEdtNC :
                 BEGIN
                 MM.SorteLettre := SorteLettre ;
                 MM.Globalise   := ( ParmsMP.GroupeEncaDeca = 'GLO' ) ; // SBO : Correction Fiche 12351 (remplace MM.Globalise := False ;)
                 MM.Section     := ParmsMP.CptG ;
                 CompleteMMRefCCMP( ParmsMP, MM.FormuleRefCCMP ) ;
                 ParmsMP.TIDTIC := False; // FQ 13422
                 END ;
      else
        begin
        if ParmsMP.NatJalG = 'BQE'
          then MM.General := ParmsMP.CptG
          else MM.Section := ParmsMP.CptG ;
        end ;
    end ;
  MM.GroupeEncadeca := ParmsMP.GroupeEncaDeca ;
  if ParmsMP.MPG<>'' then
    MM.MPGUnique := ParmsMP.MPG ;
  MM.ExportCFONB := ParmsMP.ExportCFONB ;
  MM.FormatCFONB := ParmsMP.FormatCFONB  ;
  MM.EnvoiTrans  := ParmsMP.ModeTeletrans ;
  If Not VH^.OldTeleTrans Then
    MM.EnvoiTrans := '' ;
  MM.Bordereau  := ParmsMP.EnvoiBordereau ;
  MM.Document   := ParmsMP.ModeleBordereau ;
  MM.TIDTIC     := ParmsMP.TIDTIC ;
  MM.Valide     := False ;
  If ParmsMP.smp In [smpDecChqEdt,smpDecChqEdtNC,smpEncTraEdt,smpEncTraEdtNC,smpDecBorEdt,smpDecBorEdtNC,smpDecVirEdt,smpDecVirEdtNC,smpDecVirInEdt,smpDecVirInEdtNC,smpEncPreEdt,smpEncPreEdtNC] Then
    begin
    MM.MSED.Spooler     := vBoSpooler ;
    MM.MSED.RepSpooler  := vStRepSpooler ;
    VerifRep(MM.MSED.RepSpooler) ;
    MM.MSED.SoucheSpooler   := 0 ;
    MM.MSED.XFichierSpooler := vBoFichier ;
    Case ParmsMP.smp Of
      smpDecChqEdt,smpDecChqEdtNC     : MM.MSED.RacineSpooler := GetParamSocSecur('SO_CPRACINECHEQUE','') ;
      smpEncTraEdt,smpEncTraEdtNC     : MM.MSED.RacineSpooler := GetParamSocSecur('SO_CPRACINETRAITE','') ;
      smpDecBorEdt,smpDecBorEdtNC     : MM.MSED.RacineSpooler := GetParamSocSecur('SO_CPRACINEBOR','') ;
      smpDecVirEdt,smpDecVirEdtNC     : MM.MSED.RacineSpooler := GetParamSocSecur('SO_CPRACINEVIREMENT','') ;
      smpDecVirInEdt,smpDecVirInEdtNC : MM.MSED.RacineSpooler := GetParamSocSecur('SO_RACINEVIRIN','') ; // VL 301003 FQ 12958
      smpEncPreEdt,smpEncPreEdtNC     : MM.MSED.RacineSpooler := GetParamSocSecur('SO_CPRACINEPRELEVEMENT','') ;
      end ;
    end ;
end ;

Procedure VerifRep(LaSousDir : String) ;
begin
  if LaSousDir<>'' then
    begin
    if LaSousDir[Length(LaSousDir)]='\' then
      LaSousDir := Copy( LaSousDir, 1, Length(LaSousDir)-1 ) ;
    if not DirectoryExists(LaSousDir) then
      CreateDir(LaSousDir) ;
    end ;
end ;

Function  CompteToJal ( sBqe : String ; var OkJal : Boolean ) : String ;
Var QQ : TQuery ;
    Jal : String ;
begin
  Jal   := '' ;
  OkJal := TRUE ;
  QQ    := OpenSQl('SELECT J_JOURNAL FROM JOURNAL WHERE J_CONTREPARTIE="'+sBQE+'"',True) ;
  if Not QQ.EOF then
    Jal := QQ.FindField('J_JOURNAL').AsString ;
  Ferme(QQ) ;
  if Jal='' then
    begin
    OkJal := FALSE ;
    QQ    := OpenSQl('SELECT J_JOURNAL FROM JOURNAL WHERE J_NATUREJAL="OD"',True) ;
    if Not QQ.EOF then
      Jal := QQ.FindField('J_JOURNAL').AsString ;
    Ferme(QQ) ;
    end ;
  Result:=Jal ;
end ;

Procedure FlagOrigEsc ( TOBOrig : TOB ) ;
Var XX : RMVT ;
    TOBE : TOB ;
    i    : integer ;
begin
//if Not GereEscompte then Exit ; // Test a faire avant
  if TOBOrig = Nil then Exit ;
  for i:=0 to TOBOrig.Detail.Count-1 do
    begin
    TOBE:=TOBOrig.Detail[i] ;
    if Not TOBE.FieldExists('ESCOMPTABLE') then Continue ;
    if TOBE.GetValue('ESCOMPTABLE')<>'X' then Continue ;
    XX := TOBToIdent(TOBE,True) ;
    ExecuteSQL('UPDATE ECRITURE SET E_QUALIFORIGINE="GAE" WHERE '+WhereEcriture(tsGene,XX,True)) ;
    end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 13/08/2003
Modifié le ... : 13/08/2003
Description .. : Regroupe les lignes de tiers avec cumul des montants et
Suite ........ : libère les lignes en doublons.
Mots clefs ... :
*****************************************************************}
procedure GlobaliserDest( TOBDest : TOB ; ParmsMP : TGenereMP ) ;
Var
  i       : Integer ;
  TobL    : TOB ;     // Tob de la ligne courante a traité
  TobRupt : TOB ;     // Tob de la ligne de référence pour les cumuls
  Rupture : Boolean ; // Indicateur de rupture
  oldAux  : String ;
  oldGen  : String ;
begin
  TOBRupt := TOBDest.Detail[ ( TOBDest.Detail.Count - 1 ) ] ;
  oldAux  := '' ;
  oldGen  := '' ;
  // Parcours inversé pour éviter les sauts de lignes
  for i:=TOBDest.Detail.Count-1 downto 0 do
    begin
    TOBL := TOBDest.Detail[i] ;
    Rupture := ((OldAux<>TOBL.GetValue('E_AUXILIAIRE')) or (OldGen<>TOBL.GetValue('E_GENERAL'))) ;
    if Rupture then
      begin
      // Finition de TobRupt
      if oldAux<>'' then // pour éviter le premier passage
        if TobRupt.GetValue('E_COUVERTURE') < 0 then
          begin
          TobRupt.PutValue('E_COUVERTURE',     TobRupt.GetValue('E_COUVERTURE') * (-1) ) ;
          TobRupt.PutValue('E_COUVERTUREDEV',  TobRupt.GetValue('E_COUVERTUREDEV') * (-1) ) ;
          end ;
      // MAJ des conditions de rupture et de la ligne de référence
      TobRupt := TobL ;
      oldGen  := TOBL.GetValue('E_GENERAL') ;
      oldAux  := TOBL.GetValue('E_AUXILIAIRE') ;
      // montant de couverture passe en signé pendant l'opération de cumul...
      if TobRupt.GetValue('E_DEBIT') <= TobRupt.GetValue('E_CREDIT') then
        TobRupt.PutValue('E_COUVERTURE',     TobRupt.GetValue('E_COUVERTURE') * (-1) ) ;
      end
    else
      begin
      // MAJ des cumuls
      TobRupt.PutValue('E_DEBIT',      (TobRupt.GetValue('E_DEBIT') + TobL.GetValue('E_DEBIT')) ) ;
      TobRupt.PutValue('E_CREDIT',     (TobRupt.GetValue('E_CREDIT') + TobL.GetValue('E_CREDIT')) ) ;
      TobRupt.PutValue('E_COUVERTURE', (TobRupt.GetValue('E_COUVERTURE') + TobL.GetValue('E_COUVERTURE')) ) ;
      // MAJ des cumuls en devise
      TobRupt.PutValue('E_DEBITDEV',      (TobRupt.GetValue('E_DEBITDEV') + TobL.GetValue('E_DEBITDEV')) ) ;
      TobRupt.PutValue('E_CREDITDEV',     (TobRupt.GetValue('E_CREDITDEV') + TobL.GetValue('E_CREDITDEV')) ) ;
      TobRupt.PutValue('E_COUVERTUREDEV', (TobRupt.GetValue('E_COUVERTUREDEV') + TobL.GetValue('E_COUVERTUREDEV')) ) ;
      // Libération de la ligne
      TOBL.Free ;
      end ;
    end ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 03/08/2004
Modifié le ... :   /  /    
Description .. : Transforme une pièce de simulation en pièce "normale"
Suite ........ : (Affectation d'un nouveau numéro de pièce + 
Suite ........ : E_QUALIFPIECE + MAJ des soldes des comptes , sections, 
Suite ........ : tiers, ...)
Suite ........ : 
Suite ........ : Reprise et adaptation/correction de la fonction de SaisUtil 
Suite ........ : normaliserPieceSimu
Mots clefs ... : 
*****************************************************************}
function NormaliserPieceSimuTOB(TOBL: TOB; MajTOBL: Boolean): Boolean;
Var TOBEcr  : TOB ;          // Tob de la pièce
    TOBE    : TOB ;          // pointeur sur ligne d'écriture
    TOBA    : TOB ;          // pointeur sur ligne d'analytique
    QQ      : TQuery ;
    Jal     : String ;
    NewNum  : integer ;
    i,k     : integer ;
    DD      : TDateTime ;
    nbEnreg : LongInt ;
BEGIN
  Result := False ;
  if TOBL=Nil then Exit ;

  // Chargement des lignes d'écritures
  TOBEcr := TOB.Create('',Nil,-1) ;
  QQ:=OpenSQL('SELECT * FROM ECRITURE WHERE '+WhereEcritureTob(tsGene,TOBL,False),True) ;
  TOBEcr.LoadDetailDB('ECRITURE','','',QQ,False) ;
  Ferme(QQ) ;
  if TOBEcr.Detail.Count<=0 then
     BEGIN
     TOBEcr.Free ;
     Exit ;
     END ;

  // Chargement de l'analytique (uLibAnalytique)
  ChargeAnalytique( TobEcr ) ;

  // détermination du numéro de pièce "Normal"
  Jal := TOBEcr.Detail[0].GetValue('E_JOURNAL') ;
  DD  := TOBEcr.Detail[0].GetValue('E_DATECOMPTABLE') ;
  NewNum := GetNewNumJal(Jal,True,DD) ;
  if NewNum<=0 then
    BEGIN
    TOBEcr.Free ;
    Exit ;
    END ;

  // Parcours des lignes d'écritures
  for i:=0 to TOBEcr.Detail.Count-1 do
    BEGIN
    TOBE:=TOBEcr.Detail[i] ;
    TOBE.PutValue('E_QUALIFPIECE','N') ;
    TOBE.PutValue('E_NUMEROPIECE',NewNum) ;
    // Parcours des lignes d'analytiques
    for k:=0 to TOBE.Detail.Count-1 do
        BEGIN
        TOBA:=TOBE.Detail[k] ;
        TOBA.PutValue('Y_QUALIFPIECE','N') ;
        TOBA.PutValue('Y_NUMEROPIECE',NewNum) ;
        END ;
    END ;

  // MAJ des soldes des comptes / tiers / sections...
  MajSoldesEcritureTOB(TOBEcr,True) ;

  // MAJ base des lignes d'écriture
  nbEnreg := ExecuteSQL ( 'UPDATE ECRITURE SET E_QUALIFPIECE="N", E_NUMEROPIECE=' + IntToStr(NewNum)
                         + ' WHERE ' + WhereEcritureTob(tsGene,TobL,False) ) ;
  Result := nbEnreg > 0 ;

  // MAJ base des lignes d'analytique
  if Result then
    ExecuteSQL ( 'UPDATE ANALYTIQ SET Y_QUALIFPIECE="N", Y_NUMEROPIECE=' + IntToStr(NewNum)
                +' WHERE ' + WhereEcritureTob(tsAnal,TobL,False) ) ;

  // MAJ TOB parmaètre si demandé
  if MajTOBL then
     BEGIN
     TOBL.PutValue('E_NUMEROPIECE',NewNum) ;
     TOBL.PutValue('E_QUALIFPIECE','N') ;
     END ;

  TOBEcr.Free ;

  CPStatutDossier ;

END ;


Function  GetYEnATropMsg( smp : TSuiviMP ) : string ;
begin
  result := 'Vous ne pouvez pas sélectionner plus de ' + IntToStr(MAXSELECTECHE) + ' échéances dans ce module.' ;
  if IsEnc( smp ) then
    begin
    result := result + '#13#10 (Utilisez les modules "Génération des encaissements" ' ;
    if EstSpecif('51206') then
      result := result + 'ou "Encaissement de masses" ' ;
    result := result + '#13#10  du menu "Autres traitements" pour gérer les gros volumes.)' ;
    end
  else
    begin
    result := result + '#13#10 (Utilisez les modules "Génération des décaissements" ' ;
    if EstSpecif('51206') then
      result := result + 'ou "Décaissements de masses" ' ;
    result := result + '#13#10  du menu "Autres traitements" pour gérer les gros volumes.)' ;
    end ;
end ;

end.
