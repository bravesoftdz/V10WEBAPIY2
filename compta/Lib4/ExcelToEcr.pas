{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... :   /  /
Modifié le ... : 08/12/2003
Description .. : Passage en eAGL
Mots clefs ... :
*****************************************************************}

{******************************************************************************}
{
  Modifications :
    - CA - 12/11/2000 - Arrondi sur 6 décimales pour le calcul du solde global
    - CA - 12/11/2001 - Gestion du cas ou E_AUXILIAIRE est absent de la TOB
    - CA - 10/03/2003 - Gestion des arrondi sur V_PGI.OkDecV
    - CA - 27/11/2003 - Mise à jour des champs de synchro
    - CA - 27/11/2003 - Rechargement des tablettes liées au journaux
    - CA - 27/11/2003 - E_NUMGROUPEECR forcé à 1 ( mode libre obligatoire )
    - CA - 22/01/2004 - FQ 13096 - Contrôle de la nature pour détection du lettrage inutile.
    - JP - 01/12/2005 - FQ 16865 - Le contrôle sur l'existance de la pièce était fait sur
                        VH.Encours.Code alors que l'exercice pouvait être VH.Suivant.Code
    - JP - 05/12/2005 - FQ 16865 - On va finalement toujours chercher le numéro de pièce
                        dans la table souche lorsque l'on est sur un journal "de saisie"
******************************************************************************}
unit ExcelToEcr ;

interface

uses SysUtils,
     Classes, // TStrings, TStringList
     Hctrls,  // OpenSQL, Ferme, ExisteSQL, ReadTokenSt, RempliValCombo, ReadTokenPipe, ChampToNum
{$IFDEF EAGLCLIENT}

{$ELSE}
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}

{$IFDEF VER150}
     Variants,
{$ENDIF}
     HEnt1,     // NumSemaine, V_PGI,
     Ent1,      // GetPeriode, VH, BourreEtLess, fbGene, BourreOuTronque
     UTob,      // Tob, TQuery
     Paramsoc,
     HmsgBox,
     Echeance,
     UlibEcriture,// CRemplirInfoPointage, CBlocageBor
     ULibExercice;


type
  TEcrRevisionPGI = class(TObject)
    TheTOBEcr : TOB;
    msgError : string;
    msgInfo : string;
    theJournal : string;
    theFolio : string;
  private
    procedure Init(LeJournal,LeFolio : string);
    function  Valide : boolean;
    function  CheckDetails: boolean;
    function  EcrEquilibree: boolean;
    function  StructureOK : boolean;
    procedure PurgeStructure;
    function  IgnoreEcr(T: TOB): boolean;
  public
    destructor Destroy; override;
  end;

function  ChargeJournauxRevisionPGI : variant;
function  ChargeSectionsRevisionPGI ( Ax : String ) : variant;
function  IntegreEcritureRevisionPGI( LeJournal,LeFolio : string ; CreerComptes : boolean = False; NatureJ : string = ''; Contrepartie : string = ''; RecupLibelle : Boolean=FALSE ) : variant ;
function  GetFoliosRevisionPGI( stJournal : string ) : Variant;
Function  ExcelToCompta ( TOBEPZ : TOB ; CreerComptes : boolean; NatureJ : string = ''; Contrepartie : string = ''; RecupLibelle : Boolean=FALSE ) : integer ;
// ajout me 21/06/2002
function  GetInfoAuxilaire (Auxi : string; Compteexite : Boolean ; var TOBE, TobG, TOBT: TOB; CreerComptes : Boolean) : Boolean;
Function CreationCompteAuxi ( Nature : String ; TOBT,TOBE : TOB ) : boolean ;
procedure CGetEch ( vTOB : TOB ; vInfo : TInfoEcriture = nil ) ;

implementation

Uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcGen,
  CPProcMetier,
  {$ENDIF MODENT1}
  SaisUtil,    // RMVT
  UtilPGI,     // CheckToString, EuroToFranc, FrancToEuro
  UtilSais,    // MajSoldesEcritureTOB
  ULibAnalytique, // AlloueAxe
  UtilSoc,     // MarquerPublifi
  uLibRevision  // TrouveCycleRevisionDuGene
     ;

Const EPZ_OK          = 0 ;
      EPZ_NOSELECTION = 1 ; EPZ_ERRJAL      = 2 ;
      EPZ_ERRDATE     = 3 ; EPZ_ERRNAT      = 4 ;
      EPZ_ERRGENE     = 5 ; EPZ_ERRINSERT   = 6 ;
      EPZ_ERREQUIL    = 7 ; EPZ_ERRZERO     = 8 ;
      EPZ_ERRINEX     = 9 ; EPZ_BLOCAGE     = 10;
      EPZ_ERRJALLIB   = 11; EPZ_ERRJALCTR   = 12;
      EPZ_ERRJALVAL   = 13; EPZ_ERRDATECTR  = 14;
      EPZ_ERREXISTE   = 15; EPZ_JOUNALFERME = 16;
      EPZ_ERREURETAB  = 17; EPZ_ERRMODEJRL = 18;
      EPZ_ERRCREATCPTE = 19; EPZ_COMPTEVISA = 20; // fiche 17699
      EPZ_ERRJALODA   = 21; // fiche 20233

Procedure EPZ_AlimCommun ( TOBE : TOB ; MM : RMVT ; NatureJ : string) ;
Var RegTVA : String ;
    Q      : TQuery ;
BEGIN
{RMVT}
TOBE.PutValue('E_JOURNAL',MM.Jal)          ; TOBE.PutValue('E_EXERCICE',MM.Exo) ;

if (MM.ModeSaisieJal <> 'LIB') then  // cas type journal libre, on garde la date
TOBE.PutValue('E_DATECOMPTABLE',MM.DateC)  ;

TOBE.PutValue('E_ETABLISSEMENT',MM.Etabl) ;
TOBE.PutValue('E_DEVISE',MM.CodeD)         ; TOBE.PutValue('E_TAUXDEV',MM.TauxD) ;
TOBE.PutValue('E_DATETAUXDEV',MM.DateTaux) ; TOBE.PutValue('E_QUALIFPIECE',MM.Simul) ;
TOBE.PutValue('E_NATUREPIECE',MM.Nature); TOBE.PutValue('E_NUMEROPIECE',MM.Num) ;
{ CA - 27/11/2003 - Mise à jour du NUMGROUPEECR : forcément = 1 car mode libre obligatoire }
TOBE.PutValue('E_NUMGROUPEECR',1);
TOBE.PutValue('E_VALIDE',CheckToString(MM.Valide)) ;
TOBE.PutValue('E_DATEECHEANCE',MM.DateC) ;
TOBE.PutValue('E_PERIODE',GetPeriode(MM.DateC)) ; TOBE.PutValue('E_SEMAINE',NumSemaine(MM.DateC)) ;
TOBE.PutValue('E_MODESAISIE',MM.ModeSaisieJal) ;
{Communs}
TOBE.PutValue('E_UTILISATEUR',V_PGI.User) ; TOBE.PutValue('E_TVAENCAISSEMENT','-') ;
TOBE.PutValue('E_DATECREATION',Date)      ; TOBE.PutValue('E_DATEMODIF',NowH) ;
TOBE.PutValue('E_COTATION',1)             ;
TOBE.PutValue('E_QUALIFORIGINE','EXC')    ; TOBE.PutValue('E_VISION','DEM') ;
TOBE.PutValue('E_ECRANOUVEAU','N')        ; TOBE.PutValue('E_ETATLETTRAGE','RI')  ;
TOBE.PutValue('E_CONTROLETVA','RIE')      ; TOBE.PutValue('E_ETAT','0000000000') ;
{ CA - 27/11/2003 - Mise à jour des champs de synchro }
TOBE.putValue('E_ETATREVISION','-')       ; TOBE.PutValue('E_IO','X');
{Régime TVA}
RegTVA:=VH^.RegimeDefaut ;
if RegTVA='' then
   BEGIN
   Q:=OpenSQL('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="RTV"',True) ;
   if Not Q.EOF then RegTVA:=Q.Fields[0].AsString ;
   Ferme(Q) ;
   END ;
TOBE.PutValue('E_REGIMETVA',RegTVA) ;
// ajout me 21/06/2002
TOBE.PutValue('E_CONFIDENTIEL', '0');
END ;

Procedure EPZ_AlimExcel ( TOBE,TOBExc : TOB) ;
Var DP,CP      : Double ;
    Section    : String ;
    Etb        : String ;
    TSec,TEta         : TOB;
BEGIN
DP:=Valeur(TOBExc.GetValue('E_DEBIT')) ; CP:=Valeur(TOBExc.GetValue('E_CREDIT')) ;
DP := arrondi(DP,V_PGI.OkDecV);  CP := Arrondi(CP, V_PGI.OkDecV) ; // fiche 17424
TOBE.PutValue('E_DEBIT',DP)    ; TOBE.PutValue('E_CREDIT',CP) ;
TOBE.PutValue('E_DEBITDEV',DP) ; TOBE.PutValue('E_CREDITDEV',CP) ;
TOBE.PutValue('E_REFINTERNE',TOBExc.GetValue('E_REFINTERNE')) ;
TOBE.PutValue('E_LIBELLE',TOBExc.GetValue('E_LIBELLE')) ;
if TOBExc.FieldExists('E_ETABLISSEMENT') then
   BEGIN
   Etb:=Trim(TOBExc.GetValue('E_ETABLISSEMENT')) ;
   If Etb<>'' Then
   begin
      TEta := TOB.Create('ETABLISS', nil, -1);
      TEta.PutValue ('ET_ETABLISSEMENT', Etb);
      if not (TEta.SelectDB('"' + Etb + '"', nil)) then
      begin
           TEta.PutValue ('ET_LIBELLE', 'Etablissement :' + Etb);
           TEta.InsertDB(nil, TRUE);
      end;
      TEta.Free;
   end
   else
   Etb:=GetParamSocSecur('SO_ETABLISDEFAUT', '');
   TOBE.PutValue('E_ETABLISSEMENT',Etb) ;

   END ;
if TOBExc.FieldExists('SECTION') then
   BEGIN
   TOBE.AddChampSup('SECTION',False) ;
   Section:=TOBExc.GetValue('SECTION') ;
   // FQ18071 : interdire les sections fermées
   if Not ExisteSQL('SELECT S_SECTION FROM SECTION WHERE S_AXE="A1" AND S_SECTION="'+Section+'" AND S_FERME="-"') then
      Section:=VH^.Cpta[fbAxe1].Attente ;
   TOBE.PutValue('SECTION',Section) ;
   END ;

if TOBExc.FieldExists('E_TABLE0') then
   BEGIN
       if  TOBExc.Getvalue('E_TABLE0') <> '' then
       begin
           TOBE.AddChampSup('SECTION',False) ;
            Section:= UpperCase(BourreOuTronque(TOBExc.GetValue('E_TABLE0'),fbaxe1)); //fiche 13747
            TSec := TOB.Create('SECTION', nil, -1);
            TSec.Putvalue ('S_AXE', 'A1');
            // FQ18071 : interdire les sections fermées
            if Not ExisteSQL('SELECT S_SECTION FROM SECTION WHERE S_AXE="A1" AND S_SECTION="'+Section+'" AND S_FERME="-"') then
              Section:=VH^.Cpta[fbAxe1].Attente ;

            TSec.Putvalue ('S_SECTION', Section);
            if not (TSec.SelectDB('"A1";"' + Section + '"', nil)) then
            begin
                 TSec.Putvalue ('S_LIBELLE', Section);
                 TSec.Putvalue ('S_SENS', 'M');
                 TSec.PutValue ('S_CONFIDENTIEL', '0');
                 TSec.InsertDB(nil, TRUE);
            end;
            TSec.Free;
           TOBE.PutValue('SECTION',Section) ;
       end;
   END ;

if TOBExc.FieldExists('E_LIBRETEXTE0') then
begin
     if  (TOBExc.Getvalue('E_TABLE0') = '') and  (TOBExc.Getvalue('E_LIBRETEXTE0') <> '') then
           TOBE.PutValue('E_LIBRETEXTE0', TOBExc.GetValue('E_LIBRETEXTE0')) ;
end;

END ;

Function EPZ_PremMP (Modereg : string) : String ;
Var St : String ;
Q      : Tquery;
BEGIN
Result:='CHQ' ;
if VH^.MPAcc.Count>0 then BEGIN St:=VH^.MPAcc[0] ; Result:=ReadTokenSt(St) ; END ;
if Modereg <> '' then
begin
     Q := OpenSql ('SELECT MR_MP1 from MODEREGL where MR_MODEREGLE="'+Modereg+'"', FALSE);
     if not Q.EOF then
     Result := Q.FindField ('MR_MP1').asstring;
     Ferme (Q);
end;
END ;

function EPZ_AlimFromGene ( TOBE,TOBExc,TOBGene,TOBTiers : TOB; NatureJ : string; Treso : Boolean=FALSE;Contrepartie : string = '' ) : Boolean;
Var Gene,NatG,Auxi                  : String ;
    TOBG,TOBT                       : TOB ;
    isColl,isTIDC,isLett,isLettauxi : Boolean ;
    DP,CP                           : double;
    isbanque                        : Boolean;
    Modereg                         : string;
BEGIN
Result := TRUE;
Gene:=TOBExc.GetValue('E_GENERAL') ; TOBE.PutValue('E_GENERAL',Gene) ;

Auxi:='' ;  Modereg:='';
if (TOBExc.FieldExists ('E_AUXILIAIRE')) THEN  // ajout me 21/06/2002
begin
     Auxi:=TOBExc.GetValue('E_AUXILIAIRE') ;
     TOBE.PutValue('E_AUXILIAIRE',Auxi) ;
end;

TOBG:=TOBGene.FindFirst(['G_GENERAL'],[Gene],True) ;
TOBE.PutValue('E_CONSO',TOBG.GetValue('G_CONSO')) ;
{Ventilations}
if TOBG.GetValue('G_VENTILABLE')='X' then TOBE.PutValue('E_ANA','X') ;
{Echéances/Lettrage}
NatG:=TOBG.GetValue('G_NATUREGENE') ;
isColl:=((NatG='COC') or (NatG='COD') or (NatG='COF') or (NatG='COS')) ;
isTIDC:=((NatG='TID') or (NatG='TIC')) ; isLett:=False ; isLettauxi := FALSE;
isbanque := (gene = BourreEtLess(contrepartie,fbGene) );

// ajout me
if NatureJ <> '' then
begin
   DP:=Valeur(TOBExc.GetValue('E_DEBIT')) ;
   CP:=Valeur(TOBExc.GetValue('E_CREDIT')) ;
   if (Treso) then
   begin
         if (DP > 0) then
         begin                                 // ajout me 30-10-2002
              if (isbanque) then TOBE.PutValue('E_NATUREPIECE', 'RC')
              else
              if (not isColl) or (NatG= 'COF') then TOBE.PutValue('E_NATUREPIECE', 'RF')
              else
              if (NatG = 'COC') then TOBE.PutValue('E_NATUREPIECE', 'OC');
         end
         else
         if (CP > 0) then
         begin                                    // ajout me 30-10-2002
              if (isbanque) then TOBE.PutValue('E_NATUREPIECE', 'RF')
              else
              if (not isColl) or (NatG= 'COC') then TOBE.PutValue('E_NATUREPIECE', 'RC')
              else
              if (NatG= 'COF') then TOBE.PutValue('E_NATUREPIECE', 'OF');
         end;
   end
   else
   if (NatureJ = 'ACH') and (NatG= 'COC')then
            TOBE.PutValue('E_NATUREPIECE', 'OD');
// ajout me 15-10-2002
   if (NatureJ = 'VTE') and (NatG= 'COF')then
            TOBE.PutValue('E_NATUREPIECE', 'OD');

   if (not isColl) and (Auxi <> '') then
     TOBE.PutValue('E_AUXILIAIRE','') ;

end;
if (Auxi <> '') THEN  // CA - 22/01/2004 - FQ 13096 - Contrôle de la nature inutile.
begin
  TOBT:=TOBTiers.FindFirst(['T_AUXILIAIRE'],[Auxi],True) ;
  if TOBT <> nil then
  begin
       isLettauxi:=(TOBT.GetValue('T_LETTRABLE')='X');
       Modereg := TOBT.GetValue('T_MODEREGLE');
  end;
end;
if isbanque then
begin
        if TOBG.FieldExists ('G_POINTABLE') and
        (TOBG.GetValue('G_POINTABLE')='X')   then CRemplirInfoPointage(TOBE);
end;
if isColl and isLettauxi then isLett:=True else if isTIDC then isLett:=(TOBG.GetValue('G_LETTRABLE')='X') ;
if isLett then
   BEGIN
    TOBE.PutValue('E_MODEPAIE',EPZ_PremMP(Modereg)) ;
    TOBE.PutValue('E_CODEACCEPT',MPTOACC(TOBE.GetValue('E_MODEPAIE'))) ;
    // ajout me 31-10-2002
    CRemplirInfoLettrage(TOBE);
    if (Auxi = '') then
    begin
        if NatG='COC' then Auxi:=VH^.TiersDefCli else if NatG='COF' then Auxi:=VH^.TiersDefFou else
        if NatG='COS' then Auxi:=VH^.TiersDefSal else if NatG='COD' then Auxi:=VH^.TiersDefDiv ;
        if ((Auxi='') and (isColl)) then Auxi:=VH^.TiersDefDiv ;
        if (Auxi <> '') and TOBE.FieldExists ('E_AUXILIAIRE') then  // CA - 12/11/2001
        begin
                if Not ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="'+Auxi+'"') then
                 begin Result:=FALSE ; exit; end;
                 TOBE.PutValue('E_AUXILIAIRE',Auxi) ;
        end
     end;
   END
   else
   BEGIN
       if isColl and (Auxi = '') then
       begin
            if NatG='COC' then Auxi:=VH^.TiersDefCli else if NatG='COF' then Auxi:=VH^.TiersDefFou else
            if NatG='COS' then Auxi:=VH^.TiersDefSal else if NatG='COD' then Auxi:=VH^.TiersDefDiv ;
            if ((Auxi='') and (isColl)) then Auxi:=VH^.TiersDefDiv ;
            if TOBE.FieldExists ('E_AUXILIAIRE') then  // CA - 12/11/2001
            begin
                if Not ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="'+Auxi+'"') then
                 begin Result:=FALSE ; exit; end;
                 TOBE.PutValue('E_AUXILIAIRE',Auxi) ;
                 TOBE.PutValue('E_MODEPAIE',EPZ_PremMP(Modereg)) ;
                 TOBE.PutValue('E_CODEACCEPT',MPTOACC(TOBE.GetValue('E_MODEPAIE'))) ;
                 // ajout me 31-10-2002
                 CRemplirInfoLettrage(TOBE);
            end;
       end
       else if ((NatG='CHA') or (NatG='PRO') or (NatG='IMO')) then
       begin
         TOBE.PutValue('E_TVA',TOBG.GetValue('G_TVA')) ;
         TOBE.PutValue('E_TPF',TOBG.GetValue('G_TPF')) ;
         TOBE.PutValue('E_TYPEMVT','HT') ;
       end ;
   END;

END ;

Procedure EPZ_VentilTOB ( TOBE : TOB ) ;
BEGIN
if TOBE.GetValue('E_ANA')='X' then
   BEGIN
   AlloueAxe( TOBE ) ; // SBO 25/01/2006
   // pour ventilation type
   if (TOBE.Getvalue('E_LIBRETEXTE0') <> '')  then
   begin
         VentilerTOB(TOBE,'',V_PGI.OkDecV,2,False,'',true,TOBE.Getvalue('E_LIBRETEXTE0')) ;
         TOBE.Putvalue('E_LIBRETEXTE0', '');
   end
   else
          VentilerTOB(TOBE,'',V_PGI.OkDecV,2,False);

   END ;
END ;

Procedure EPZ_EquilTOB ( TOBEcr : TOB ) ;
Var XD,XC,DE,CE,Delta  : Double ;
    i      : integer ;
    TOBE   : TOB ;
    G1,A1,G2,A2,CptG,CptA : String ;
BEGIN
if TOBEcr.Detail.Count<=2 then Exit ;
G1:=TOBEcr.Detail[0].GetValue('E_GENERAL') ;
if TOBEcr.Detail[0].FieldExists('E_AUXILIAIRE') then   // CA - 12/11/2001
  A1:=TOBEcr.Detail[0].GetValue('E_AUXILIAIRE')
else A1 := '';
G2:=TOBEcr.Detail[1].GetValue('E_GENERAL') ;
if TOBEcr.Detail[1].FieldExists('E_AUXILIAIRE') then   // CA - 12/11/2001
  A2:=TOBEcr.Detail[1].GetValue('E_AUXILIAIRE')
else A2 := '';
DE:=0 ; CE:=0 ;
for i:=0 to TOBEcr.Detail.Count-1 do
    BEGIN
    TOBE:=TOBEcr.Detail[i] ;
    XD:=TOBE.GetValue('E_DEBIT') ; XC:=TOBE.GetValue('E_CREDIT') ;
    DE:=DE+XD ; CE:=CE+XC ;
    if i=0 then BEGIN CptG:=G2 ; CptA:=A2 ; END else BEGIN CptG:=G1 ; CptA:=A1 ; END ;
    TOBE.PutValue('E_CONTREPARTIEGEN',CptG) ; TOBE.PutValue('E_CONTREPARTIEAUX',CptA) ;
    TOBE.PutValue('E_ENCAISSEMENT',SensEnc(XD,XC)) ;
    END ;
//Delta:=Arrondi(DE-CE,6) ;
Delta:=Arrondi(DE-CE,V_PGI.OkDecV) ;  // CA - 10/03/2003
if Delta=0 then Exit ;
TOBE:=TOBEcr.Detail[TOBEcr.Detail.Count-1] ;
DE:=TOBE.GetValue('E_DEBIT') ; CE:=TOBE.GetValue('E_CREDIT') ;
if DE<>0 then BEGIN DE:=DE-Delta ; TOBE.PutValue('E_DEBIT',DE) ; END
         else BEGIN CE:=CE+Delta ; TOBE.PutValue('E_CREDIT',CE) ; END ;
END ;

Function EPZ_CVDATE ( StD : String ) : TDateTime ;
BEGIN
if isNumeric(StD) then Result:=ValeurI(StD) else Result:=StrToDate(StD) ;
END ;

Procedure RemplirNatureSens ( Gene,Auxi : String ; Var Nature,Sens : String ) ;
Var CC,C2,C3 : Char ;
BEGIN
CC:=Gene[1] ; C2:=Gene[2] ; C3:=Gene[3] ;
Nature:='' ; Sens:='' ;
Case CC of
   '1' : BEGIN Nature:='DIV' ; Sens:='C' ; END ;
   '2' : BEGIN
         Nature:='IMO' ; Sens:='D' ;
         if C2 in ['8','9'] then BEGIN Nature:='DIV' ; Sens:='C' ; END ;
         END ;
   '3' : BEGIN
         Nature:='DIV' ; Sens:='D' ;
         if C2='9' then BEGIN Nature:='DIV' ; Sens:='C' ; END ;
         END ;
   '4' : BEGIN
         Nature:='DIV' ; if Auxi<>'' then Nature:='COD' ;
         Case C2 of
            '0' : BEGIN
                  Sens:='C' ; if C3='9' then Sens:='D' ;
                  if Auxi<>'' then
                       Nature:='COF'
                  else   // ajout me 30-10-2002
                       if gene = GetParamSocSecur('SO_DEFCOLFOU', '') then Nature:='COF';
                  END ;
            '1' : BEGIN
                  Sens:='D' ; if C3='9' then Sens:='C' ;
                  if Auxi<>'' then Nature:='COC'
                  else   // ajout me 30-10-2002
                       if gene = GetParamSocSecur('SO_DEFCOLCLI', '') then Nature:='COC';
                  END ;
            '8' : BEGIN Sens:='D' ; END ;
             else BEGIN Sens:='C' ; END ;
            END ;
         END ;
   '5' : BEGIN
         Nature:='DIV' ; Sens:='D' ;
         if C2='1' then BEGIN Nature:='BQE' ; Sens:='D' ; END else
         if C2='3' then BEGIN Nature:='CAI' ; Sens:='D' ; END else
         if C2='9' then BEGIN Nature:='DIV' ; Sens:='C' ; END ;
         END ;
   '6' : BEGIN Nature:='CHA' ; Sens:='D' ; END ;
   '7' : BEGIN Nature:='PRO' ; Sens:='C' ; END ;
   END ;
END ;

Function CreationCompte ( Gene : String ; TOBG,TOBE : TOB ) : boolean ;
Var Sens,Nature,Auxi,LibRef : String ;
Q1  :TQuery;
lStEtatCycle : string;
BEGIN
Result:=False ;
if Length(Gene)<=3 then Exit ;
if TOBG=Nil then Exit ;
if TOBE.FieldExists('E_AUXILIAIRE') then   // CA - 12/11/2001
  Auxi:=TOBE.GetValue('E_AUXILIAIRE')
else Auxi := '';
LibRef:=TOBE.GetValue('E_LIBELLE') ; if LibRef='' then LibRef:=TOBE.GetValue('E_REFINTERNE') ;
// ajout me 30-10-2002
if (LibRef='') or (LibRef='...') then LibRef:='Créé par Excel' ;
RemplirNatureSens(Gene,Auxi,Nature,Sens) ;
if ((Nature='') or (Sens='')) then Exit ;
TOBG.PutValue('G_GENERAL',Gene) ;
TOBG.PutValue('G_NATUREGENE',Nature) ;
if ((Nature='COC') or (Nature='COF') or (Nature='COS') or (Nature='COD')) then
     TOBG.PutValue('G_COLLECTIF','X') ;
TOBG.PutValue('G_SENS',Sens) ;
TOBG.PutValue('G_LIBELLE',LibRef) ;
TOBG.PutValue('G_ABREGE',Copy(LibRef,1,17)) ;

// ajout me 30-10-2002
TOBG.PutValue('G_VENTILABLE1', '-');
TOBG.PutValue('G_VENTILABLE2', '-');
TOBG.PutValue('G_VENTILABLE3', '-');
TOBG.PutValue('G_VENTILABLE4', '-');
TOBG.PutValue('G_VENTILABLE5', '-');
TOBG.PutValue('G_VENTILABLE', '-');
TOBG.PutValue('G_CONFIDENTIEL', '0');

  // GCO - 29/10/2007 - FQ 21760
  if VH^.Revision.Plan <> '' then
  begin
    TOBG.PutValue('G_CYCLEREVISION', TrouveCycleRevisionDuGene( Gene ));
    TobG.PutValue('G_VISAREVISION', '-');

    if TobG.GetValue('G_CYCLEREVISION') = '' then
    begin // Aucun cycle de révision ne contient le général
      if GetParamSocSecur('SO_CPREVISBLOQUEGENE', False, True) then Exit;
    end
    else
    begin
      lStEtatCycle := GetColonneSQL('CREVCYCLE',  'CCY_ETATCYCLE',
                      'CCY_CODECYCLE = "' + TobG.GetValue('G_CYCLEREVISION') + '" AND ' +
                      'CCY_EXERCICE = "' + VH^.EnCours.Code + '"');

      if (lStEtatCycle = cValide) or (lStEtatCycle = cSupervise) then
      begin
        TobG.PutValue('G_VISAREVISION', 'X');
      end;
    end;
  end
  else
  begin
    TOBG.PutValue('G_CYCLEREVISION', '');
    TOBG.PutValue('G_VISAREVISION', '-');
  end;
  MiseAJourCREVGeneraux( Gene, TobG.GetValue('G_VISAREVISION')= 'X');
  // FIN GCO - 30/10/2007

// fiche 17505
if TOBE.FieldExists('E_TABLE0') and (TOBE.Getvalue('E_TABLE0') <> '') then
begin
           TOBG.PutValue('G_VENTILABLE', 'X');
           TOBG.PutValue('G_VENTILABLE1', 'X');
end
else
begin
     if (TOBE.FieldExists('E_LIBRETEXTE0')) and (TOBE.Getvalue('E_LIBRETEXTE0') <> '') then
     begin
           TOBG.PutValue('G_VENTILABLE', 'X');
         Q1 := OpenSql ('SELECT * FROM VENTIL WHERE V_COMPTE IN (SELECT DISTINCT(CC_CODE) FROM CHOIXCOD '+
            'WHERE CC_TYPE="VTY" AND CC_CODE="'+TOBE.GetValue('E_LIBRETEXTE0')+'") AND V_NATURE LIKE "TY%"', TRUE);
         While Not Q1.EOF do
         BEGIN
          Nature := Q1.FindField ('V_NATURE').Asstring;
          Case Nature[3] of
              '1' : TOBG.PutValue('G_VENTILABLE1', 'X');
              '2' : TOBG.PutValue('G_VENTILABLE2', 'X');
              '3' : TOBG.PutValue('G_VENTILABLE3', 'X');
              '4' : TOBG.PutValue('G_VENTILABLE4', 'X');
              '5' : TOBG.PutValue('G_VENTILABLE5', 'X');
          end;
          Q1.Next;
         END;
         Ferme (Q1);
     end;
end;

Result:=TOBG.InsertDB(Nil) ;
END ;

// ajout me 16-10-2002 pour récupérer les libelles dans la fiche excel
// si les comptes existent pour une balance
Function ModifLibellecompte ( Gene : String ; TOBG,TOBE : TOB ) : boolean ;
Var LibRef : String ;
BEGIN
Result:=False ;
LibRef:=TOBE.GetValue('E_LIBELLE') ;
if LibRef='' then exit;
TOBG.PutValue('G_GENERAL',Gene) ;
TOBG.PutValue('G_LIBELLE',LibRef) ;
TOBG.PutValue('G_ABREGE',Copy(LibRef,1,17)) ;
Result:=TOBG.InsertOrUpdateDB(FALSE) ;
END ;

// ajout me
Function CreationCompteAuxi ( Nature : String ; TOBT,TOBE : TOB ) : boolean ;
Var Gene,Auxi,LibRef : String ;
BEGIN
Result:=False ;
if TOBT=Nil then Exit ;
if Nature='' then Exit;
if TOBE.FieldExists('E_AUXILIAIRE') then   // CA - 12/11/2001
  Auxi:=TOBE.GetValue('E_AUXILIAIRE')
else Auxi := '';
// ajout me 15-10-2002
if Auxi = '' then exit;

Gene := TOBE.GetValue('E_GENERAL');
LibRef:=TOBE.GetValue('E_LIBELLE') ; if LibRef='' then LibRef:=TOBE.GetValue('E_REFINTERNE') ;
// ajout me 15-10-2002
if (LibRef='') or (LibRef='...') then LibRef:='Créé par Excel' ;
TOBT.PutValue('T_AUXILIAIRE',Auxi) ;
TOBT.PutValue('T_TIERS',Auxi) ;
TOBT.PutValue('T_COLLECTIF',Gene) ;
TOBT.PutValue('T_NATUREAUXI',Nature) ;
TOBT.PutValue('T_LIBELLE',LibRef) ;
TOBT.PutValue('T_ABREGE',Copy(LibRef,1,17)) ;
// ajout me 15-10-2002
TOBT.PutValue('T_REGIMETVA', VH^.RegimeDefaut);
//TOBT.PutValue('T_DEVISE', V_PGI.DevisePivot);
TOBT.PutValue('T_LETTRABLE', 'X');

// ajout me 30-10-2002
TOBT.PutValue('T_MULTIDEVISE', 'X');
TOBT.PutValue('T_DEVISE', '');
TOBT.PutValue('T_CONFIDENTIEL', '0');
if (Nature = 'CLI') or (Nature = 'FOU') then
begin
  TOBT.PutValue('T_MODEREGLE', GetParamSocSecur('SO_GCMODEREGLEDEFAUT', ''));
  TOBT.PutValue('T_TVAENCAISSEMENT', GetParamSocSecur('SO_CODETVADEFAUT', ''));
end;
Result:=TOBT.InsertDB(Nil) ;
END ;


Function ExcelToCompta ( TOBEPZ : TOB ; CreerComptes : boolean; NatureJ : string = ''; Contrepartie : string = ''; RecupLibelle : Boolean=FALSE ) : integer ;
Var Jal,NatEcr,Gene                : String ;
    TOBE,TOBG,TOBGene,TOBECR,TOBEP : TOB ;
    Piece, Num,i, PremLig          : integer ;
    OkG,OkZero,OkCompte,OkDate,OkV : Boolean ;
    MM                             : RMVT ;
    Q                              : TQuery ;
    DateC,LimDate,SavDate,Datefin  : TDateTime ;
    TotD,TotC,XD,XC                : Double ;
    Auxi                           : string; // ajout me 21/06/2002
    NatA, SIM                      : string;
    TOBT,TOBTIERS                  : TOB;
    DP,CP                          : double;
    TJ                             : TOB;
    Contrep,Where                  : string;
    Exo                            : TExoDate;
    OkCreat                        : Boolean;
BEGIN
Result:=EPZ_OK ;
OkDate := TRUE;     OkCreat := TRUE;
if TOBEPZ=Nil then BEGIN Result:=EPZ_NOSELECTION ; Exit ; END ;
if TOBEPZ.Detail.Count<=1 then BEGIN Result:=EPZ_NOSELECTION ; Exit ; END ;
TOBE:=TOBEPZ.Detail[0] ;
{Date comptable}
if Not TOBE.FieldExists('E_DATECOMPTABLE') then BEGIN Result:=EPZ_ERRDATE ; Exit ; END ;
if TOBE.FieldExists('E_ETABLISSEMENT') then  // fiche 13670
begin
    If not ExisteSQL ('SELECT ET_ETABLISSEMENT from ETABLISS Where ET_ETABLISSEMENT="'+ TOBE.GetValue('E_ETABLISSEMENT')+'"') then
    BEGIN
         Result:=EPZ_ERREURETAB ; Exit ;
    END ;
end;

Datefin := FinDeMois(StrToDate(TOBE.getvalue('E_DATECOMPTABLE')));
if (VH^.DateCloturePer>0) and (Datefin <=VH^.DateCloturePer) then
BEGIN Ferme (Q); Result := EPZ_ERRDATECTR ; Exit ; END ;

DateC:=EPZ_CVDATE(TOBE.GetValue('E_DATECOMPTABLE')) ;
if DateC<VH^.Encours.Deb then BEGIN Result:=EPZ_ERRDATE ; Exit ; END ;
if VH^.Suivant.Code<>'' then LimDate:=VH^.Suivant.Fin else LimDate:=VH^.Encours.Fin ;
if DateC>LimDate then BEGIN Result:=EPZ_ERRDATE ; Exit ; END ;
{Journal+Numéro}
if Not TOBE.FieldExists('E_JOURNAL') then BEGIN Result:=EPZ_ERRJAL ; Exit ; END ;
Jal:=TOBE.GetValue('E_JOURNAL') ; if Jal='' then BEGIN Result:=EPZ_ERRJAL ; Exit ; END ;
Piece := TOBE.GetValue('E_NUMEROPIECE');

if not ( CEstBloqueJournal(Jal, false) ) then
    CBloqueurJournal(true, Jal) ;

SIM :='N';
if TOBE.FieldExists('E_QUALIFPIECE') then  // fiche 15390  et fiche 21559
   SIM := TOBE.GetValue('E_QUALIFPIECE');

// CA - 29/11/2001  - DEBUT
{ Numéro de folio }
  if Piece = 0 then
    // Calcul du numéro de pièce
    Num:=GetNewNumJal(Jal, (SIM = 'N'), DateC) // fiche 21559
  else Num := Piece;
  if Num<=0 then
         Num := 1;
{ Calcul du premier numéro de ligne }
  if Piece = 0 then PremLig := 0
  else
  begin
    Where := 'WHERE E_EXERCICE="'+VH.Encours.Code+
      '" AND E_JOURNAL="'+Jal+'" AND E_NUMEROPIECE='+IntToStr(Num)+' AND E_QUALIFPIECE="N" AND '+
      'E_PERIODE='+IntToStr(GetPeriode(DateC));
    if ExisteSQL('SELECT E_VALIDE FROM ECRITURE ' + Where + 'AND E_VALIDE="X"') then
    BEGIN Ferme (Q); Result := EPZ_ERRJALVAL ; CBloqueurJournal(false, Jal); Exit ; END ;
    Q := OpenSQL ('SELECT MAX(E_NUMLIGNE) AS MAXLIGNE FROM ECRITURE ' + Where,True);
    if not Q.Eof then PremLig := Q.FindField('MAXLIGNE').AsInteger
    else BEGIN Ferme (Q); Result:=EPZ_ERRJAL ; CBloqueurJournal(false, Jal);  Exit ; END ;
    Ferme (Q);
  end;
// CA - 29/11/2001 - FIN

{Nature}

// ajout me 21/06/2002
if (NatureJ <> '') then
begin
      if NatureJ = 'ACH' then NatEcr:='FF' else
      if NatureJ = 'VTE' then NatEcr:='FC'
      else NatEcr := 'OD';
end else NatEcr:='OD' ;

{Comptes généraux}
if Not TOBE.FieldExists('E_GENERAL') then BEGIN Result:=EPZ_ERRGENE ; CBloqueurJournal(false, Jal); Exit ; END ;
TOBGene:=TOB.Create('',Nil,-1) ; OkG:=True ; OkV := TRUE; OkZero:=True ; TotD:=0 ; TotC:=0 ;


   TOBTIERS:=TOB.Create('',nil,-1) ;
   SavDate := DateC;

for i:=0 to TOBEPZ.Detail.Count-1 do
    BEGIN
    TOBE:=TOBEPZ.Detail[i] ; Gene:=TOBE.GetValue('E_GENERAL') ;
// CA - 19/03/2002 - Corrige pb des comptes avec longueur supérieur à SO_LGCPTEGEN
    Gene:=BourreEtLess(Gene,fbGene) ; TOBE.PutValue('E_GENERAL',Gene) ;
    if OKDate and (SavDate <> EPZ_CVDATE(TOBE.GetValue('E_DATECOMPTABLE'))) then
       OkDate := FALSE;

    SavDate := EPZ_CVDATE(TOBE.GetValue('E_DATECOMPTABLE'));

    TOBG:=TOB.Create('GENERAUX',TOBGene,-1) ;
// ajout me
    Auxi := '';
    if (TOBE.FieldExists ('E_AUXILIAIRE')) THEN  // ajout me 21/06/2002
    begin
         if (TOBE.GetValue('E_AUXILIAIRE') <> '') then
         begin
              Auxi:=BourreEtLess(TOBE.GetValue('E_AUXILIAIRE'),fbAux) ;
              TOBE.PutValue('E_AUXILIAIRE',Auxi) ;
         end;
         TOBT:=TOB.Create('TIERS',TOBTIERS,-1) ;
    end;
    if (Gene='') and (Auxi = '') then BEGIN OkG:=False ; Break ; END ;

    OkCreat := TRUE;
    if Not TOBG.SelectDB('"'+Gene+'"',Nil) then
       BEGIN
       OkCompte:=False ;
       if not CreerComptes then  begin OkCreat := FALSE;  break; end;
       if CreerComptes then if CreationCompte(Gene,TOBG,TOBE) then OkCompte:=True ;
       if Auxi <> '' then
         BEGIN
              // ajout me 21/06/2002
              if not GetInfoAuxilaire (Auxi, Okcompte, TOBE, TobG, TOBT, CreerComptes) then
              begin
                        if not CreerComptes then OkCreat := FALSE; 
                        OkG:=False ; Break ;
              end;
              // Compte fermé ? FQ18071
              if TOBE.FieldExists ('E_AUXILIAIRE') and (TOBE.GetString('E_AUXILIAIRE') <> '') and
                 Assigned(TOBT) and ( TOBT.GetString('T_FERME') = 'X' ) then
                begin
                OkG:=False ;
                Break ;
                end;
         END ;
       END
       ELSE
       BEGIN
              // Compte fermé ? FQ18071
              if  TOBG.GetString('G_FERME') = 'X' then
                begin
                OkG:=False ;
                Break ;
                end ;
              // Fiche 17699
              if  TOBG.GetString('G_VISAREVISION') = 'X' then
                begin
                OkV := False ;
                Break ;
                end ;
              // ajout me 16-10-2002
              if RecupLibelle then ModifLibellecompte (Gene,TOBG,TOBE);
              // ajout me 21/06/2002
              OkCompte:=TRUE ;
              // fiche 18108
              if not GetInfoAuxilaire (Auxi, Okcompte, TOBE, TobG, TOBT, CreerComptes) then
              begin
                        if not CreerComptes then OkCreat := FALSE;
                        OkG:=False ; Break ;
              end;
              // Compte fermé ? FQ1 8071
              if TOBE.FieldExists ('E_AUXILIAIRE') and (TOBE.GetString('E_AUXILIAIRE') <> '') and
                 Assigned(TOBT) and ( TOBT.GetString('T_FERME') = 'X' ) then
                begin
                OkG:=False ;
                Break ;
                end;
       END;
    XD:=Valeur(TOBE.GetValue('E_DEBIT')) ; XC:=Valeur(TOBE.GetValue('E_CREDIT')) ;
    TotD:=TotD+XD ; TotC:=TotC+XC ;
    // ajout me pour révision afin d'effectuer une création de compte
    // if ((XD=0) and (XC=0)) then BEGIN OkZero:=False ; Break ; END ;
    END ;
if not OkCreat then begin Result := EPZ_ERRCREATCPTE; CBloqueurJournal(false, Jal);  Exit; end; //fiche 18108
if Not OkG then BEGIN TOBGene.Free ; Result:=EPZ_ERRGENE ; CBloqueurJournal(false, Jal); Exit ; END ;
if Not OkZero then BEGIN TOBGene.Free ; Result:=EPZ_ERRZERO ; CBloqueurJournal(false, Jal); Exit ; END ;
if Not OkV then Begin TOBGene.Free ; Result:= EPZ_COMPTEVISA; CBloqueurJournal(false, Jal); Exit; end; // fiche 17699


//if Arrondi(TotD-TotC,6)<>0 then
if Arrondi(TotD-TotC,V_PGI.OkDecV)<>0 then  // CA - 10/03/2003
  BEGIN
   TOBGene.Free ; Result:=EPZ_ERREQUIL ; CBloqueurJournal(false, Jal); Exit ;
  END ;
{Consitution MM}
FillChar(MM,Sizeof(MM),#0) ;
MM.Etabl:=VH^.EtablisDefaut ; MM.Jal:=Jal ; MM.Num:=Num ;   MM.Simul := SIM; // fiche 21559
if not ExisteSQL('select ET_ETABLISSEMENT from ETABLISS Where ET_ETABLISSEMENT="' + MM.Etabl+'"' ) then  // fiche 13950 10/08/2005
BEGIN
     Result:=EPZ_ERREURETAB; CBloqueurJournal(false, Jal);  Exit ;
END;

MM.Exo:=QuelExoDT(DateC) ; MM.DateC:=DateC ; MM.DateTaux:=DateC ;
MM.Nature:=NatEcr ; MM.CodeD:=V_PGI.DevisePivot ; MM.TauxD:=1 ; MM.Treso := FALSE;

Q:=OpenSQL('SELECT J_MODESAISIE,J_NATUREJAL,J_CONTREPARTIE, J_FERME FROM JOURNAL WHERE J_JOURNAL="'+Jal+'"',True) ;
if Not Q.EOF then
begin
     MM.ModeSaisieJal:=Q.FindField('J_MODESAISIE').AsString;
     // ajout me
     MM.Treso := (Q.FindField('J_NATUREJAL').AsString = 'BQE');
     if Q.FindField('J_FERME').AsString ='X' then  // fiche 15820
     BEGIN Ferme(Q); TOBGene.Free ; Result:=EPZ_JOUNALFERME ; CBloqueurJournal(false, Jal);  Exit ; END ;

     if not OkDate and (MM.ModeSaisieJal <> 'LIB') then // Fiche 16548
            Result := EPZ_ERRJALLIB;   // fiche  16548

     if (Q.FindField('J_NATUREJAL').AsString = 'ODA') then
     begin
      Ferme(Q); TOBGene.Free ; Result := EPZ_ERRJALODA; CBloqueurJournal(false, Jal);  Exit ;
     end;

     Contrep:=BourreEtLess(Contrepartie,fbGene);
     if (NatureJ = 'BQE') and (BourreEtLess(Q.FindField('J_CONTREPARTIE').AsString,fbGene) <> Contrep) then
     BEGIN Ferme(Q); TOBGene.Free ; Result:=EPZ_ERRJALCTR ; CBloqueurJournal(false, Jal);  Exit ; END ;

     if (MM.Simul <> 'N') and (MM.ModeSaisieJal <> '-') then  // fiche 18106
     BEGIN Ferme(Q); TOBGene.Free ; Result:=EPZ_ERRMODEJRL ; CBloqueurJournal(false, Jal);  Exit ; END ;

end
else
begin
        TJ := TOB.Create ('JOURNAL',nil,-1);
        TJ.PutValue('J_JOURNAL',JAL);
        TJ.PutValue('J_LIBELLE',JAL+' CREATION PAR EXCEL');
        TJ.PutValue('J_ABREGE',JAL);
        if NatureJ = '' then NatureJ := 'OD';
        TJ.PutValue('J_NATUREJAL',NatureJ);
        TJ.PutValue('J_COMPTEURNORMAL',JAL);
        TJ.PutValue('J_CONTREPARTIE',Contrepartie);
        TJ.PutValue('J_MODESAISIE','LIB');
        TJ.putvalue ('J_CENTRALISABLE',  '-');
        TJ.putvalue ('J_MULTIDEVISE', 'X');
        TJ.putvalue ('J_CREERPAR', 'EXC');
        // ajout me 30-10-2002
        TJ.putvalue ('J_ABREGE', JAL);
        TJ.putvalue ('J_COMPTEURSIMUL',  'SIM');
        if NatureJ = 'BQE' then TJ.PutValue ('J_TYPECONTREPARTIE', 'MAN');
        TJ.InsertDB(nil);
        TJ.Free;
        TJ := TOB.Create ('SOUCHE',nil,-1);
        TJ.PutValue('SH_TYPE','CPT');
        TJ.PutValue('SH_SOUCHE',Jal);
        TJ.PutValue('SH_LIBELLE','SOUCHE :'+Jal);
        TJ.PutValue('SH_NUMDEPART',1);
        TJ.InsertDB(nil);
        TJ.Free;
        if (MM.Simul <> 'N') then // fiche 18106
           MM.ModeSaisieJal := '-'
        else
           MM.ModeSaisieJal := 'LIB';
        MM.Treso := (NatureJ = 'BQE');
        { CA - 27/11/2003 - Mise à jour des tablettes }
        AvertirMultiTable ( 'TTJOURNAL');
end;
Ferme(Q) ;

if MM.ModeSaisieJal='ERR' then BEGIN TOBGene.Free ; Result:=EPZ_ERRJAL ; CBloqueurJournal(false, Jal);  Exit ; END ;
if MM.ModeSaisieJal = '-' then   // Mode saisie
begin
    {JP 05/12/05 : FQ 16865 : En fait si l'on veux éviter les trous et les doublons, il faut quoiqu'il arrive
                   récupérer le numéro de souche sans faire de test
     JP 01/12/05 : FQ 16865 : MM.Exo peut valoir VH.Suivant.Code, ce qui rend la requête ci-dessous inopérante
                   et autorise à se retrouver avec le même numéro de pièce sur une même période !!!
    //Where := 'WHERE E_EXERCICE="'+VH.Encours.Code+
    Where := 'WHERE E_EXERCICE = "' + MM.Exo +

      '" AND E_JOURNAL="'+Jal+'" AND E_NUMEROPIECE='+IntToStr(Num)+' AND E_QUALIFPIECE="N" AND '+
      'E_PERIODE='+IntToStr(GetPeriode(DateC));
    if ExisteSQL('SELECT E_NUMEROPIECE FROM ECRITURE ' + Where ) then  // fiche 13950 10/08/2005
    BEGIN}
    if Piece = 0 then  // fiche 19595  GetNewNumJal est déjà fait
      MM.num := num
    else
      MM.num :=  GetNewNumJal(Jal,(MM.Simul = 'N'),DateC);   // fiche 21559
   PremLig := 0;
end;
{Ecriture}
TOBEcr:=TOB.Create('',Nil,-1) ;
for i:=0 to TOBEPZ.Detail.Count-1 do
    BEGIN
    TOBEP:=TOBEPZ.Detail[i] ;
    // ajout me pour révision

    DP := Valeur(TOBEP.GetValue('E_DEBIT')) ;
    CP := Valeur(TOBEP.GetValue('E_CREDIT')) ;
    if (DP = 0) and (CP = 0) then continue;

    TOBE:=TOB.Create('ECRITURE',TOBEcr,-1) ;
    DateC:=StrTodate(TOBEP.GetValue('E_DATECOMPTABLE'));

    // ajout me pour vérifier la cohérence des dates de lâ pièce  06-08-2004
    if (not CQuelExercice(DateC, Exo)) or (MM.Exo <> Exo.code) then
    begin Result:=EPZ_ERRDATE ; CBloqueurJournal(false, Jal);  exit; end;

    TOBE.putValue('E_DATECOMPTABLE', DateC);
    EPZ_AlimCommun(TOBE,MM, NatureJ) ; // ajout me 21/06/2002
    EPZ_AlimExcel(TOBE,TOBEP) ;
    if not EPZ_AlimFromGene(TOBE,TOBEP,TOBGene,TOBTiers, NatureJ, MM.Treso, Contrepartie) then
    begin Result:=EPZ_ERRINEX ; CBloqueurJournal(false, Jal);  exit; end;

    TOBE.PutValue('E_NUMLIGNE',PremLig+i+1) ;
    // ajout me 11/08/2005 fiche 15868
    TOBE.PutValue( 'E_QUALIFQTE1'        , '...' ) ;
    TOBE.PutValue( 'E_QUALIFQTE2'        , '...' ) ;

    END ;
{Equilibre, ajustements}
EPZ_EquilTOB(TOBEcr) ;

for i:=0 to TOBEcr.Detail.Count-1 do
begin
    CGetEch(TOBEcr.Detail[i]) ;         // calcul de la date d'échéance
    EPZ_VentilTOB(TOBEcr.Detail[i]) ;   // Ventilation analytique
end;

// ajout me 30-10-2002
if not CBlocageBor(
                TOBEcr.detail[0].GetValue('E_JOURNAL'),
                TOBEcr.detail[0].GetValue('E_DATECOMPTABLE'),
                TOBEcr.detail[0].GetValue('E_NUMEROPIECE'),FALSE) then
                begin Result:=EPZ_BLOCAGE ; CBloqueurJournal(false, Jal);  exit; end;

{Insertion}
if TOBEcr.InsertDBByNivel(False) then
   BEGIN
   MajSoldesEcritureTOB(TOBEcr,True) ;
   ADevalider(MM.Jal,MM.DateC) ;
   CPSTATUTDOSSIER ;
   MarquerPublifi(True) ;
   if (Result <> EPZ_ERRJALLIB)  then
      Result := EPZ_OK
   END else
   BEGIN
      Result:=EPZ_ERRINSERT ;
   END ;
{Libérations}
 CBloqueurJournal(false, Jal);
// ajout me 30-10-2002
CBloqueurBor(
                   TOBEcr.Detail[0].GetValue('E_JOURNAL'),
                   TOBEcr.Detail[0].GetValue('E_DATECOMPTABLE'),
                   TOBEcr.Detail[0].GetValue('E_NUMEROPIECE'), false, false);

TOBGene.Free ;
TOBEcr.Free ;
TOBTIERS.free;


END ;

(*  Fonction ne marchant pas car tablette TZSECTION NON rechargeable en mémoire
function  ChargeSectionsRevisionPGI ( Ax : String ) : variant;
var V       : Variant;
    It, Va  : TStrings;
    i       : integer;
    StTable : String ;
    Q1      : TQuery;
BEGIN
Result:=varEmpty ;
It:=TStringList.Create ; Va:=TStringList.Create ;
StTable:='TZSECTION' ; if Ax<>'A1' then StTable:=StTable+Ax[2] ;
RemplirValCombo(StTable,'','',It,Va,False,False);
V:=VarArrayCreate([0,It.Count],varVariant);
for i:=0 to It.Count-1 do V[i]:=Va[i]+' : '+It[i] ;
Result:=V ;
It.Free ; Va.Free ;
END ;
*)

function  ChargeSectionsRevisionPGI ( Ax : String ) : variant;
var V       : Variant;
    Va      : TStrings;
    i       : integer;
    Q1      : TQuery;
BEGIN
Result:=varEmpty ;
Va:=TStringList.Create ;
Q1 := OpenSql ('Select S_SECTION, S_LIBELLE FROM SECTION Where S_AXE="'+Ax+'"', TRUE);
While not Q1.EOF do
begin
     Va.Add(Q1.FindField('S_SECTION').asstring) ;
     Q1.next;
end;
ferme (Q1);
V:=VarArrayCreate([0, Va.Count],varVariant);
for i:=0 to Va.Count-1 do  V[i]:= Va[i] ;
Result:=V ; Va.Free ;
END ;

function ChargeJournauxRevisionPGI : variant;
var V : Variant;
    It, Va : TStrings;
    i : integer;
begin
Result:=varEmpty ;
It:=TStringList.Create ; Va:=TStringList.Create ;
RemplirValCombo('TTJALGUIDE','','',It,Va,False,False);
V:=VarArrayCreate([0,It.Count],varVariant);
for i:=0 to It.Count-1 do V[i]:=Va[i]+' : '+It[i] ;
Result:=V ;
It.Free ; Va.Free ;
end;

Function MessExc ( ii : integer ) : String ;
BEGIN
Result:='Erreur d''enregistrement' ;
Case ii of
      EPZ_NOSELECTION : Result:='La sélection doit contenir au moins 2 lignes' ;
      EPZ_ERRJAL      : Result:='Journal ou numérotation incorrecte' ;
     //fiche 13092 EPZ_ERRJALVAL   : Result:='Journal ou numérotation validé' ;
      EPZ_ERRJALVAL   : Result:='Le bordereau est validé, l''écriture ne peut pas être comptabilisée';
      EPZ_ERRDATE     : Result:='La date est incorrecte' ;
      EPZ_ERRGENE     : Result:='Un des comptes est absent ou incorrect' ;
      EPZ_ERRINSERT   : Result:='Insertion non effectuée' ;
      EPZ_ERREQUIL    : Result:='Ecriture non équilibrée' ;
      EPZ_ERRZERO     : Result:='Il existe des lignes à zéro' ;
      EPZ_ERRINEX     : Result:='Compte d''attente inexistant';
      EPZ_BLOCAGE     : Result:='Bordereau bloqué par un autre utilisateur';
      EPZ_ERRJALLIB   : Result := 'EPZ_ERRJALLIB' ;
      EPZ_ERRJALCTR   : Result:='Le compte de banque doit être le compte de contrepartie du journal';
      EPZ_ERRDATECTR  : Result:= 'La période est clôturée à cette date';
      EPZ_ERREXISTE   : Result:= 'Cette pièce existe déjà';
      EPZ_JOUNALFERME : Result:= 'Le journal est fermé';
      EPZ_ERREURETAB  : Result:= 'Etablissement inexistant';
      EPZ_ERRMODEJRL  : Result:= 'Le mode de saisie du journal n''est pas pièce';
      EPZ_ERRCREATCPTE: Result:= 'Des comptes saisis n''existent pas, veuillez les modifier';
      EPZ_COMPTEVISA  : Result := 'Mouvement lié à un compte visé non modifiable'; // fiche 17699
      EPZ_ERRJALODA   : Result := 'Le journal est de nature OD analytique' // fiche 20233
      END ;
END ;

function GetFoliosRevisionPGI( stJournal : string ) : Variant;
var V : Variant;
    i : integer;
    Q : TQuery;
    stJal,stSelect,stWhere,stOrderBy : string;
begin
  V := VarArrayCreate([0,0], varVariant);
  result := V;
  if stJournal='' then exit;
  stJal := ReadTokenPipe(stJournal,' : ');
  stSelect := 'select distinct(e_numeropiece),e_journal,e_qualifpiece,e_valide,e_numligne,e_numeche,e_exercice from ecriture';
  stWhere := ' where e_journal="'+stJal+'" and e_qualifpiece="N" and e_valide="-" and e_numligne<=1 and e_numeche<=1 '+
		     'and e_exercice="' + VH^.Encours.Code + '"';
  stOrderBy := ' order by e_numeropiece';
  Q := OPenSQL(stSelect+stWhere+stOrderBy,true);
  if Q.EOF then begin Ferme(Q); exit; end;
  i:=0;
  while not Q.EOF do
    begin
    VarArrayRedim(V,i+1);
    V[i] := Q.Fields[0].AsString;
    Q.Next; Inc(i);
    end;
  Ferme(Q);
  result := V;
end;

// NatureJ ajout me 21/06/2002
function  IntegreEcritureRevisionPGI( LeJournal,LeFolio : string ; CreerComptes : boolean = False; NatureJ : string = ''; Contrepartie : string = ''; RecupLibelle : Boolean=FALSE ) : variant ;
Var TheEcr : TEcrRevisionPGI ;
    ii     : integer ;
begin
Result:='' ;
TheEcr:=TEcrRevisionPGI.Create;
TheEcr.Init(LeJournal,LeFolio) ;
TheEcr.TheTOBEcr.Detail[0].Free ;
TheEcr.PurgeStructure;
if TheEcr.Valide then
   BEGIN
   ii:=ExcelToCompta(TheEcr.TheTOBEcr,CreerComptes, NatureJ, Contrepartie,RecupLibelle) ;
   if ii>0 then Result:=MessExc(ii) ;
   END else
   BEGIN
   Result:=TheEcr.msgError ;
   END ;
TheEcr.Free ;
end;

function TEcrRevisionPGI.Valide : boolean;
begin
  Result:=False ; msgError:='' ;
  if TheTOBEcr=nil then msgError := 'Aucune écriture a intégrer' else
   if TheTOBEcr.Detail.Count<2 then msgError := 'Une écriture nécessite au moins 2 lignes' else
    if not StructureOK then else
     if not CheckDetails then else
     if not EcrEquilibree then else
  Result := msgError='';
end;


destructor TEcrRevisionPGI.Destroy;
begin
if TheTOBEcr<>Nil then TheTOBEcr.Free;
inherited Destroy;
end;

function TEcrRevisionPGI.StructureOK : boolean;
var T : TOB;
    bExiste : boolean;
begin
  T := TheTOBEcr.Detail[0] ;
  bExiste := T.FieldExists ('E_DATECOMPTABLE') and
           T.FieldExists ('E_DEBIT') and
           T.FieldExists ('E_CREDIT') and
           T.FieldExists ('E_REFINTERNE') and
           T.FieldExists ('E_LIBELLE') and
           T.FieldExists ('E_GENERAL');
  if not bExiste then msgError := 'Structure de traitement invalide (vérifier le nom des colonnes).';
  Result:=(msgError='') ;
end;

procedure TEcrRevisionPGI.PurgeStructure;
var
    T : TOB;
    i : integer;
    NomCh : string;
begin
  T := TheTOBEcr.Detail[0] ;  // 1ere tob des données
  i := T.ChampsSup.Count-1;
  while i>=0 do
    begin
    NomCh:=TCS(T.ChampsSup[i]).Nom ;
    if ChamptoNum(NomCh)=-1 then
      T.DelChampSup(NomCh,true);
    Dec(i);
    end;
end;

function TEcrRevisionPGI.EcrEquilibree : boolean;
var Solde, Debit, Credit : double;
begin
  msgError := '';
  Debit := TheTOBEcr.Somme('E_DEBIT',[''],[''],false);
  Credit := TheTOBEcr.Somme('E_CREDIT',[''],[''],false);
//  Solde:=Arrondi(Credit-Debit,6) ; // CA - 12/11/2001
  Solde:=Arrondi(Credit-Debit,V_PGI.OkDecV) ; // CA - 10/03/2003
  if Solde<>0 then msgError := Format('L''écriture doit être équilibrée (écart %f)',[Solde]);
  result := msgError = '';
end;

function TEcrRevisionPGI.CheckDetails : boolean;
var nCount,i : integer;
    T : TOB;
    debit, credit : double;
    stdate, compte, libelle : string;
    bExisteauxi       : Boolean; // ajout me
    Auxi              : string;
begin
  msgError := ''; msgInfo := ''; i:=0;
  Auxi := '';
  nCount := TheTOBEcr.Detail.Count;
  while i < nCount do
    begin
    T := TheTOBEcr.Detail[i];
    if IgnoreEcr(T) then begin Dec(nCount); continue; end;

    // ajout me 4/07/2002
    bExisteauxi := T.GetNumChamp('E_AUXILIAIRE')<>-1;
    if bExisteauxi then Auxi := T.GetValue('E_AUXILIAIRE')
    else Auxi := '';

    msgError := ''; msgInfo := '';
    stdate := T.GetValue('E_DATECOMPTABLE');
    debit := Valeur(T.GetValue('E_DEBIT')); credit := Valeur(T.GetValue('E_CREDIT'));
    libelle := T.GetValue('E_LIBELLE')    ; compte := T.GetValue('E_GENERAL');
//EPZ    if IgnoreEcr(T) then begin Dec(nCount); continue; end else Inc(i);
    if (debit<>0) and (credit<>0) then msgError := 'Toute ligne doit avoir un et un seul montant' else
    // ajout me pour révision
//    if (not (ctxPCL in V_PGI.PGIContexte)) and (debit=0) and (credit=0) then msgError := 'Toute ligne doit avoir au moins un montant' else

    if libelle='' then msgError := 'Toute ligne doit contenir un libellé' else
    // ajout me 4/07/2002  pour Auxi
    if (compte='') and (Auxi = '') then msgError := 'Toute ligne doit avoir un compte renseigné' else
           ;
    if msgError<>'' then break;
    if not IsValidDate(stDate) then
       begin
       msgInfo := 'La Date est invalide '+stDate;
       T.PutValue('E_DATECOMPTABLE',VH^.EnCours.Fin);
       end;
    T.PutValue('E_DEBIT',Debit); T.PutValue('E_CREDIT',Credit);
    T.AddChampSupValeur('E_JOURNAL',theJournal);
    // CA - 29/11/2001
    if IsNumeric(theFolio) or (theFolio='') then
    begin
      if theFolio = '' then T.AddChampSupValeur('E_NUMEROPIECE ',0)
      else T.AddChampSupValeur('E_NUMEROPIECE',StrToInt(theFolio));
    end else msgError := 'Le numéro de folio doit être numérique ou vide';
    Inc(i);
    end;
  Result:=(msgError='') ;
end;

procedure TEcrRevisionPGI.Init(LeJournal,LeFolio : string);
var St : string;

{
function LoadFromClpBrdSlk : TOB;
var
  MyHandle: THandle;
  TextPtr: PChar;
  theSylk : TSlkGeneral;
  mySt : string;
  T : TOB;
begin
  ClipBoard.Open;
try
  MyHandle := Clipboard.GetAsHandle(CF_SYLK);
  TextPtr := GlobalLock(MyHandle);
  mySt := StrPas(TextPtr);
  GlobalUnlock(MyHandle);
finally
  Clipboard.Close;
  theSylk := TSlkGeneral.Create;
  theSylk.ParseString(St);
  T := TOB.Create('TOBSYLK',nil,-1);
  T.GetTobFromTheSylk(theSylk);
  theSylk.Destroy;
  result := T;
end;
end;
}

begin
// EN ATTENDANT LA 545...
TheTOBEcr:=TOBLoadFromClipBoard;
//TheTOBEcr:=LoadFromClpBrdSlk;
//
//EPZ 30/09/2002 déplcé dans le IntegreEcritureRevisionPGI
//PurgeStructure;
St:=LeJournal; theJournal:=ReadTokenPipe(St,' : ');
// ajout me 21/06/2002
theJournal := UpperCase (theJournal);
theFolio:=LeFolio ;
end;

function TEcrRevisionPGI.IgnoreEcr(T : TOB) : boolean;
var stdate, debit, credit : string;
    bSupp : boolean;
begin
  bSupp := not (T.FieldExists ('E_DATECOMPTABLE') and
           T.FieldExists ('E_DEBIT') and
           T.FieldExists ('E_CREDIT') and
           T.FieldExists ('E_REFINTERNE') and
           T.FieldExists ('E_LIBELLE') and
           T.FieldExists ('E_GENERAL'));
  if bSupp then else
    begin
    stdate := T.GetValue('E_DATECOMPTABLE');
    debit := T.GetValue('E_DEBIT'); if debit='' then debit := '0';
    credit := T.GetValue('E_CREDIT'); if credit='' then credit := '0';
    bSupp := (not IsValidDate(stDate)) or
             (not IsNumeric(debit,true)) or
             (not IsNumeric(credit,true));
    end;
  if bSupp then T.Free;
  Result := bSupp;
end;

function GetInfoAuxilaire (Auxi : string; Compteexite : Boolean ; var TOBE, TobG, TOBT : TOB; CreerComptes : Boolean) : Boolean;
var
Gene          : string;
NatA,NAtG     : string;
begin
     Result := TRUE;
     if Auxi = '' then exit;
     if (not (TOBE.FieldExists ('E_AUXILIAIRE'))) then begin Result := FALSE; exit; end;
     NatG:=TOBG.GetValue('G_NATUREGENE') ;

     if not Compteexite then // cas compte n'existe pas
     begin
           Gene := VH^.Cpta[fbGene].Attente;
           if TOBT.SelectDB('"'+Auxi+'"',Nil) then
           begin
                TOBE.PutValue('E_AUXILIAIRE', Auxi);
                NatA := TOBT.GetValue ('T_NATUREAUXI');
                // ajout me 15-10-2002
                if TOBT.GetValue ('T_COLLECTIF') <> '' then
                   Gene := TOBT.GetValue ('T_COLLECTIF')
                else
                begin
                if NatA='CLI' then Gene :=VH^.DefautCli else if NatA='FOU' then Gene :=VH^.DefautFou else
                if NatA='SAL' then Gene :=VH^.DefautSal else if NatA='DIV' then Gene := VH^.DefautDDivers;
                end;
           end
           else
           begin
                     if (NatG = 'COC') then NatA := 'CLI' else
                     if (NatG = 'COF') then NatA := 'FOU' else
                     if (NatG = 'COS') then NatA := 'SAL' else
                     if (NatG = 'COD') then NatA := 'DIV';
                     if not CreerComptes then  begin Result := FALSE; exit; end;
                     if not CreationCompteAuxi(NatA,TOBT,TOBE) then
                            TOBE.PutValue('E_AUXILIAIRE', '');
           end;
           if TOBG.SelectDB('"'+Gene+'"',Nil) then
                TOBE.PutValue('E_GENERAL', Gene)
           else Result := FALSE;
     end
     else
     begin
                 if TOBT.SelectDB('"'+Auxi+'"',Nil) then
                 begin
                 NatA := TOBT.GetValue ('T_NATUREAUXI');
                 if ((NatG = 'COC') and (NatA = 'CLI')) or
                    ((NatG = 'COF') and (NatA = 'FOU')) or
                    ((NatG = 'COS') and (NatA = 'SAL')) or
                    ((NatG = 'COD') and (NatA = 'DIV')) then
                           TOBE.PutValue('E_AUXILIAIRE', Auxi)
                    else
                            TOBE.PutValue('E_AUXILIAIRE', '');
                 end
                 else
                 begin
                     if (NatG = 'COC') then NatA := 'CLI' else
                     if (NatG = 'COF') then NatA := 'FOU' else
                     if (NatG = 'COS') then NatA := 'SAL' else
                     if (NatG = 'COD') then NatA := 'DIV';
                     if not CreerComptes then  begin Result := FALSE; exit; end;
                     if not CreationCompteAuxi(NatA,TOBT,TOBE) then
                            TOBE.PutValue('E_AUXILIAIRE', '');
                 end;
     end;
end;

procedure CGetEch ( vTOB : TOB ; vInfo : TInfoEcriture = nil ) ;
var
 lTOB                : TOB ;
 lDtDateEch          : TDateTime ;
 lModeR              : T_MODEREGL ;
 lStModePaie         : string ;
 lBoBqe              : boolean ;
 lBoResult           : boolean ;
 lStModeAux          : string ;
 lRdSolde            : double ;
 lStDeca             : string ;
 lBoDetruireInfo     : boolean ;
begin

  lBoDetruireInfo  := False;
 if vInfo = nil then
  begin
   vInfo           := TInfoEcriture.Create ;
   lBoDetruireInfo := true ;
  end ;


 lBoBqe := false ;
 lTOB   := vTOB ;
 if not vInfo.LoadCompte(lTOB.GetValue('E_GENERAL')) then exit ;

 if ( vInfo.Compte.GetValue('G_COLLECTIF') = 'X' ) then
  begin
   if not vInfo.LoadAux(lTOB.GetValue('E_AUXILIAIRE')) or
    ( ( vInfo.Aux.InIndex > - 1 ) and  (vInfo.Aux.GetValue('T_LETTRABLE') = '-' ) ) then Exit ;
  end
   else
    if ( vInfo.Compte.GetValue('G_POINTABLE')='X') and ( (vInfo.Compte.GetValue('G_NATUREGENE')='BQE') or (vInfo.Compte.GetValue('G_NATUREGENE')='CAI') ) then
     begin
      lBoBqe   := true ;
     end
      else
       if ( vInfo.Compte.GetValue('G_LETTRABLE') = '-' ) then
        begin
         CSupprimerInfoLettrage(lTOB) ;
         lTOB.PutValue('E_MODEPAIE', '') ;
         exit ;
        end ;

 lStModePaie   := lTOB.GetValue('E_MODEPAIE') ;

 vInfo.Devise.Load([lTOB.GetValue('E_DEVISE')]) ;
 vInfo.Devise.AffecteTaux(lTOB.GetValue('E_DATECOMPTABLE')) ;
 lDtDateEch             := lTOB.GetValue('E_DATECOMPTABLE') ;
 lModeR.Action          := taCreat ;
 lModeR.TotalAPayerP    := lTOB.GetValue('E_DEBIT')    + lTOB.GetValue('E_CREDIT') ;
 lModeR.TotalAPayerD    := lTOB.GetValue('E_DEBITDEV') + lTOB.GetValue('E_CREDITDEV') ;
 lModeR.CodeDevise      := vInfo.Devise.Dev.Code ;
 lModeR.Symbole         := vInfo.Devise.Dev.Symbole ;
 lModeR.Quotite         := vInfo.Devise.Dev.Quotite ;
 lModeR.TauxDevise      := vInfo.Devise.Dev.Taux ;
 lModeR.Decimale        := vInfo.Devise.Dev.Decimale ;
 lModeR.DateFact        := lDtDateEch ;
 lModeR.DateBL          := lDtDateEch ;
 lModeR.DateFactExt     := lDtDateEch ;

 if ( vInfo.Aux.InIndex > - 1 ) then
  lModeR.ModeInitial := vInfo.Aux.GetValue('T_MODEREGLE') ;

 if lStModeAux <> '' then
  lModeR.Aux := lStModeAux
   else
    lModeR.Aux := lTOB.GetValue('E_AUXILIAIRE') ;

 if lModeR.Aux ='' then
  lModeR.Aux := lTOB.GetValue('E_GENERAL') ;

 CalculModeregle(lModeR, false) ;

 lDtDateEch              := lModeR.TabEche[1].DateEche ;
 lStModePaie             := lModeR.TabEche[1].ModePaie ;

 lBoResult := true ;

 lRdSolde  := Arrondi(lTOB.GetValue('E_DEBIT') - lTOB.GetValue('E_CREDIT') , V_PGI.OkDecV ) ;
 if lRdSolde = 0 then lStDeca := ''
  else
   if lRdSolde > 0 then lStDeca := 'ENC'
    else lStDeca := 'DEC' ;

 if lBoResult or ( lTOB.GetValue('E_DATEECHEANCE') = iDate1900 ) then
   begin
    if lBoBqe then
     begin
      lTOB.PutValue('E_MODEPAIE'         , lStModePaie) ;
      lTOB.PutValue('E_DATEECHEANCE'     , lDtDateEch) ;
      lTOB.PutValue('E_ORIGINEPAIEMENT'  , lDtDateEch) ;
      lTOB.PutValue('E_ECHE'             , 'X') ;
      lTOB.PutValue('E_NUMECHE'          , 1) ;
      lTOB.PutValue('E_ETATLETTRAGE'     , 'RI') ;
      lTOB.PutValue('E_ENCAISSEMENT'     , lStDeca) ;
      lTOB.PutValue('E_DATEVALEUR'       , lTOB.GetValue('E_DATECOMPTABLE')) ;
      lTOB.PutValue('E_DATEPAQUETMIN'    , lTOB.GetValue('E_DATECOMPTABLE')) ;
      lTOB.PutValue('E_DATEPAQUETMAX'    , lTOB.GetValue('E_DATECOMPTABLE')) ;
     end
      else
       begin
        lTOB.PutValue('E_MODEPAIE'       , lStModePaie ) ;
        lTOB.PutValue('E_DATEECHEANCE'   , lDtDateEch ) ;
        lTOB.PutValue('E_ORIGINEPAIEMENT', lDtDateEch ) ;
        lTOB.PutValue('E_ECHE'           , 'X') ;
        lTOB.PutValue('E_NUMECHE'        , 1) ;
        lTOB.PutValue('E_ENCAISSEMENT'   , lStDeca) ;
        lTOB.PutValue('E_DATEVALEUR'     , lTOB.GetValue('E_DATECOMPTABLE')) ;
        lTOB.PutValue('E_DATEPAQUETMIN'  , lTOB.GetValue('E_DATECOMPTABLE')) ;
        lTOB.PutValue('E_DATEPAQUETMAX'  , lTOB.GetValue('E_DATECOMPTABLE')) ;
        if lTOB.GetValue('E_ETATLETTRAGE')='RI' then lTOB.PutValue('E_ETATLETTRAGE', 'AL') ;
       end ; // if
   end; // if

 if lBoDetruireInfo then vInfo.Free ;

end ;

end.
