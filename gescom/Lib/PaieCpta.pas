unit PaieCpta ;

interface

Uses HEnt1, HCtrls, UTOB, Ent1, DB, DBTables, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls,
     FE_Main, SysUtils, Dialogs, SaisUtil, UtilPGI, AGLInit, FactUtil, UtilSais,
     EntGC, Classes, EdtDoc, HMsgBox, SaisComm, FactComm, ed_tools, Saisie ;

Type T_RetCpta = (rcOk,rcRef,rcPar) ;

Function  DetruitCompta ( TOBPiece : TOB ) : RMVT ;
Function  PassationComptable ( TOBPiece,TOBBases,TOBEches,TOBTiers,TOBArticles,TOBCpta : TOB ; NbDec : integer ; OldEcr : RMVT ) : boolean ;
Procedure LanceZoomEcriture ( TOBPiece : TOB ) ;
Function  ChargeAjouteCompta ( TOBCpta,TOBPiece,TOBA,TOBTiers : TOB ; AvecAnal : boolean ) : TOB ;
Procedure LoadLesAna ( TOBPiece,TOBAna : TOB ) ;
Procedure PreVentileLigne ( TOBC,TOBAP,TOBL : TOB ) ;
Function  EncodeRefCPGescom ( TOBPiece : TOB ) : String ;

implementation

Var PremHT,PremTVA,LastMsg,NbEches : integer ;
    EstAvoir                       : Boolean ;
    TTA : TList ;

Type T_CodeCpta = Class
                  CptHT,LibArt : String ;
                  FamTaxe : Array[1..5] of String ;
                  SommeTaxeD,SommeTaxeP,SommeTaxeE : Array[1..5] of double ;
                  XD,XP,XE,Qte : Double ;
                  LibU : boolean ;
                  Anal : TList ;
                  Constructor Create ;
                  Destructor Destroy ;
                  End ;

Type T_DetAnal = Class
                 Section,Ax : String ;
                 MD,MP,ME : Double ;
                 End ;

Type T_CreatPont = (ccpEscompte,ccpRemise,ccpHT,ccpTaxe) ;

const TexteMessage: array[1..15] of string 	= (
          {1}  'Compte général d''escompte absent ou incorrect'
          {2} ,'Compte général de remise absent ou incorrect'
          {3} ,'Compte général de taxe absent ou incorrect'
          {4} ,'Compte collectif tiers absent ou incorrect'
          {5} ,'Compte général de HT absent ou incorrect'
          {6} ,'Compte général d''écart de conversion absent ou incorrect'
          {7} ,'Journal comptable non renseigné pour cette nature de pièce'
          {8} ,'Nature comptable non rernseignée'
          {9} ,'Erreur sur la numérotation du journal comptable'
         {10} ,''
         {11} ,''
         {12} ,''
         {13} ,''
         {14} ,''
         {15} ,''
              );
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
Description .. : Encodage d'un identifiant de pièce Gescom dans l'écriture Comptable
Mots clefs ... : LIAISON;COMPTABILITE;PONT;PASSATION;
*****************************************************************}
Function EncodeRefCPGescom ( TOBPiece : TOB ) : String ;
BEGIN
Result:='' ; if TOBPiece=Nil then Exit ;
Result:=TOBPiece.GetValue('GP_NATUREPIECEG')+';'+TOBPiece.GetValue('GP_SOUCHE')+';'
       +FormatDateTime('ddmmyyyy',TOBPiece.GetValue('GP_DATEPIECE'))+';'
       +IntToStr(TOBPiece.GetValue('GP_NUMERO'))+';'+IntToStr(TOBPiece.GetValue('GP_INDICEG'))+';' ;
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

Procedure LanceZoomEcriture ( TOBPiece : TOB ) ;
Var XX : RMVT ;
    RefCP : String ;
BEGIN
RefCP:=TOBPiece.GetValue('GP_REFCOMPTABLE') ; if RefCP='' then Exit ;
XX:=DecodeRefGCComptable(RefCP) ;
XX.CodeD:=TOBPiece.GetValue('GP_DEVISE') ;
XX.ModeOppose:=(TOBPiece.GetValue('GP_SAISIECONTRE')='X') ;
XX.Etabl:=TOBPiece.GetValue('GP_ETABLISSEMENT') ;
LanceSaisie(Nil,taConsult,XX) ;
END ;

Procedure LoadLesAna ( TOBPiece,TOBAna : TOB ) ;
Var Q : TQuery ;
    TOBL,TOBA,TOBAL : TOB ;
    RefA : String ;
    i,NumL : integer ;
BEGIN
RefA:=EncodeRefCPGescom(TOBPiece) ; if RefA='' then Exit ;
Q:=OpenSQL('SELECT * FROM VENTANA WHERE YVA_TABLEANA="GL" AND YVA_IDENTIFIANT="'+RefA+'"',True) ;
if Not Q.EOF then TOBAna.LoadDetailDB('VENTANA','','',Q,False) ;
Ferme(Q) ;
TOBAna.Detail.Sort('-YVA_AXE;-YVA_IDENTLIGNE') ;
for i:=TOBAna.Detail.Count-1 downto 0 do
    BEGIN
    TOBA:=TOBAna.Detail[i] ; NumL:=ValeurI(TOBA.GetValue('YVA_IDENTLIGNE')) ;
    if NumL>0 then
       BEGIN
       TOBL:=TOBPiece.Detail[NumL-1] ; TOBAL:=TOBL.Detail[0] ;
       TOBA.ChangeParent(TOBAL,-1) ;
       END ;
    END ;
END ;

Procedure PreVentileLigne ( TOBC,TOBAP,TOBL : TOB ) ;
Var i,NumA : integer ;
    RefA : String ;
    TOBV,TOBA,TOBPiece,TOBAna,LaTV : TOB ;
BEGIN
if TOBL=Nil then Exit ;
TOBPiece:=TOBL.Parent ; RefA:=EncodeRefCPGescom(TOBPiece) ;
TOBAna:=TOBL.Detail[0] ; TOBAna.ClearDetail ;
if TOBC<>Nil then
   BEGIN
   LaTV:=TOBC ;
   if ((LaTV.Detail.Count<=0) and (TOBAP.Detail.Count>0)) then LaTV:=TOBAP ; 
   END else
   BEGIN
   LaTV:=TOBAP ;
   END ;
if LaTV.Detail.Count<=0 then Exit ;
for i:=0 to LaTV.Detail.Count-1 do
    BEGIN
    TOBV:=LaTV.Detail[i] ;
    TOBA:=TOB.Create('VENTANA',TOBAna,-1) ;
    TOBA.PutValue('YVA_TABLEANA','GL') ;
    TOBA.PutValue('YVA_IDENTIFIANT',RefA) ;
    TOBA.PutValue('YVA_NATUREID','GC') ;
    TOBA.PutValue('YVA_IDENTLIGNE',FormatFloat('000',TOBL.GetValue('GL_NUMLIGNE'))) ;
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

Procedure RenseigneClefCompta ( TOBPiece : TOB ; Var MM : RMVT ) ;
BEGIN
{A remplir}
MM.Etabl:=TOBPiece.GetValue('GP_ETABLISSEMENT') ;
MM.DateC:=TOBPiece.GetValue('GP_DATEPIECE') ;
{ok}
MM.Exo:=QuelExoDT(MM.DateC) ;
MM.CodeD:=V_PGI.DevisePivot ; MM.DateTaux:=MM.DateC ;
MM.TauxD:=TOBPiece.GetValue('GP_TAUXDEV') ;
MM.ModeSaisieJal:='-' ; MM.ModeOppose:='-' ;
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
Function DetruitCompta ( TOBPiece : TOB ) : RMVT ;
Var RefComptable,Ax : String ;
    MM           : RMVT ;
    Nb,i,k,NumL,NumA : integer ;
    Q            : TQuery ;
    TOBEcr,TOBAna,TOBE,TOBA,TOBAxe : TOB ;
BEGIN
FillChar(MM,Sizeof(MM),#0) ; Result:=MM ;
if TOBPiece=Nil then Exit ;
RefComptable:=TOBPiece.GetValue('GP_REFCOMPTABLE') ; if RefComptable='' then Exit ;
MM:=DecodeRefGCComptable(RefComptable) ; Result:=MM ;
{Mise à jour inverse des soldes des comptes}
if MM.Simul='N' then
   BEGIN
   {Ecritures}
   TOBEcr:=TOB.Create('',Nil,-1) ;
   Q:=OpenSQL('SELECT * FROM ECRITURE WHERE '+WhereEcriture(tsGene,MM,False),True) ;
   if Not Q.EOF then TOBEcr.LoadDetailDB('ECRITURE','','',Q,False,True) ;
   Ferme(Q) ;
   {Analytiques}
   TOBAna:=TOB.Create('',Nil,-1) ;
   Q:=OpenSQL('SELECT * FROM ANALYTIQ WHERE '+WhereEcriture(tsAnal,MM,False),True) ;
   if Not Q.EOF then TOBAna.LoadDetailDB('ANALYTIQ','','',Q,False,True) ;
   Ferme(Q) ;
   {Changement de parent}
   TOBAna.Detail.Sort('Y_NUMLIGNE;Y_NUMVENTIL') ;
   for i:=TOBAna.Detail.Count-1 downto 0 do
       BEGIN
       TOBA:=TOBAna.Detail[i] ;
       NumL:=TOBA.GetValue('Y_NUMLIGNE') ;
       Ax:=TOBA.GetValue('Y_AXE') ; NumA:=Ord(Ax[2])-48 ;
       TOBE:=TOBEcr.FindFirst(['E_NUMLIGNE'],[NumL],False) ;
       if TOBE.Detail.Count<=0 then BEGIN for k:=1 to 5 do TOB.Create('A'+IntToStr(k),TOBE,-1) ; END ;
       TOBAxe:=TOBE.Detail[NumA-1] ; TOBA.ChangeParent(TOBAxe,-1) ;
       END ;
   {MAJ Soldes}
   MajSoldesEcritureTOB(TOBEcr,False) ;
   {Frees}
   TOBEcr.Free ; TOBAna.Free ;
   END ;
Nb:=ExecuteSQL('DELETE FROM ECRITURE WHERE '+WhereEcriture(tsGene,MM,False)) ;
if Nb<=0 then BEGIN V_PGI.IoError:=oeunknown ; Exit ; END ;
Nb:=ExecuteSQL('DELETE FROM ANALYTIQ WHERE '+WhereEcriture(tsAnal,MM,False)) ; {! il peut ne pas y avoir d'analytique}
END ;

Procedure RecupLesDC ( TOBL : TOB ; Var XD,XP,XE : Double ) ;
Var CodeD : String ;
BEGIN
XD:=0 ; XP:=0 ; XE:=0 ;
// Simplifier à l'extrême
if TOBL.NomTable='PIEDECHE' then
   BEGIN
   XP:=TOBL.GetValue('GPE_MONTANTECHE') ; XE:=TOBL.GetValue('GPE_MONTANTCON') ;
   CodeD:=TOBL.GetValue('GPE_DEVISE');
   if CodeD<>V_PGI.DevisePivot then XD:=TOBL.GetValue('GPE_MONTANTDEV') else XD:=XP ;
   END else
if TOBL.NomTable='LIGNE' then
   BEGIN
   XP:=TOBL.GetValue('GL_MONTANTHT') ; XE:=TOBL.GetValue('GL_MONTANTHTCON') ;
   CodeD:=TOBL.GetValue('GL_DEVISE');
   if CodeD<>V_PGI.DevisePivot then XD:=TOBL.GetValue('GL_MONTANTHTDEV') else XD:=XP ;
   END else
if TOBL.NomTable='PIEDBASE' then
   BEGIN
   XP:=TOBL.GetValue('GPB_BASETAXE') ; XE:=TOBL.GetValue('GPB_BASECON') ;
   CodeD:=TOBL.GetValue('GPB_DEVISE');
   if CodeD<>V_PGI.DevisePivot then XD:=TOBL.GetValue('GPB_BASEDEV') else XD:=XP ;
   END else
    ;
END ;

Procedure RecupLesDCTaxes ( TOBL : TOB ; Var XD,XP,XE : Double ) ;
Var CodeD : String ;
BEGIN
XD:=0 ; XP:=0 ; XE:=0 ;
XP:=TOBL.GetValue('GPB_VALEURTAXE') ; XE:=TOBL.GetValue('GPB_VALEURCON') ;
CodeD:=TOBL.GetValue('GPB_DEVISE');
if CodeD<>V_PGI.DevisePivot then XD:=TOBL.GetValue('GPB_VALEURDEV') else XD:=XP ;
END ;

Procedure RecupLesDCEsc ( TOBL : TOB ; Var XD,XP,XE : Double ) ;
Var CodeD : String ;
BEGIN
XD:=0 ; XP:=0 ; XE:=0 ;
XP:=TOBL.GetValue('GP_TOTALESC') ; XE:=TOBL.GetValue('GP_TOTALESCCON') ;
CodeD:=TOBL.GetValue('GP_DEVISE');
if CodeD<>V_PGI.DevisePivot then XD:=TOBL.GetValue('GP_TOTALESCDEV') else XD:=XP ;
END ;

Procedure RecupLesDCRem ( TOBL : TOB ; Var XD,XP,XE : Double ) ;
Var CodeD : String ;
BEGIN
XD:=0 ; XP:=0 ; XE:=0 ;
XP:=TOBL.GetValue('GP_TOTALREMISE') ; XE:=TOBL.GetValue('GP_TOTALREMISECON') ;
CodeD:=TOBL.GetValue('GP_DEVISE');
if CodeD<>V_PGI.DevisePivot then XD:=TOBL.GetValue('GP_TOTALREMISEDEV') else XD:=XP ;
END ;

Procedure PieceVersECR ( MM : RMVT ; TOBPiece,TOBTiers,TOBE : TOB ) ;
Var i : integer ;
    RefI : String ;
BEGIN
{RMVT}
TOBE.PutValue('E_JOURNAL',MM.Jal)          ; TOBE.PutValue('E_EXERCICE',MM.Exo) ;
TOBE.PutValue('E_DATECOMPTABLE',MM.DateC)  ; TOBE.PutValue('E_ETABLISSEMENT',MM.Etabl) ;
TOBE.PutValue('E_DEVISE',MM.CodeD)         ; TOBE.PutValue('E_TAUXDEV',MM.TauxD) ;
TOBE.PutValue('E_DATETAUXDEV',MM.DateTaux) ; TOBE.PutValue('E_QUALIFPIECE',MM.Simul) ;
TOBE.PutValue('E_NATUREPIECE',MM.Nature)   ; TOBE.PutValue('E_NUMEROPIECE',MM.Num) ;
TOBE.PutValue('E_VALIDE',CheckToString(MM.Valide)) ;
TOBE.PutValue('E_DATEECHEANCE',MM.DateC) ;
TOBE.PutValue('E_PERIODE',GetPeriode(MM.DateC)) ; TOBE.PutValue('E_SEMAINE',NumSemaine(MM.DateC)) ;
{Pièce}
TOBE.PutValue('E_REGIMETVA',TOBPiece.GetValue('GP_REGIMETAXE')) ;
TOBE.PutValue('E_REFLIBRE',TOBPiece.GetValue('GP_REFINTERNE')) ;
TOBE.PutValue('E_REFEXTERNE',TOBPiece.GetValue('GP_REFEXTERNE')) ;
TOBE.PutValue('E_DATEREFEXTERNE',TOBPiece.GetValue('GP_DATEREFEXTERNE')) ;
TOBE.PutValue('E_UTILISATEUR',TOBPiece.GetValue('GP_UTILISATEUR')) ;
TOBE.PutValue('E_DATECREATION',TOBPiece.GetValue('GP_DATECREATION')) ;
TOBE.PutValue('E_DATEMODIF',TOBPiece.GetValue('GP_DATEMODIF')) ;
TOBE.PutValue('E_AFFAIRE',TOBPiece.GetValue('GP_AFFAIRE')) ;
TOBE.PutValue('E_COTATION',TOBPiece.GetValue('GP_COTATION')) ;
TOBE.PutValue('E_SAISIEEURO',TOBPiece.GetValue('GP_SAISIECONTRE')) ;
TOBE.PutValue('E_REFPAIE',EncodeRefCPPaie(TOBPiece)) ;
{Tiers}
TOBE.PutValue('E_LIBELLE',TOBTiers.GetValue('T_LIBELLE')) ;
TOBE.PutValue('E_CONFIDENTIEL',TOBTiers.GetValue('T_CONFIDENTIEL')) ;
{Divers}
TOBE.PutValue('E_QUALIFORIGINE','GC') ; TOBE.PutValue('E_VISION','DEM') ;
TOBE.PutValue('E_ECRANOUVEAU','N')    ; TOBE.PutValue('E_MODESAISIE','-') ;
TOBE.PutValue('E_CONTROLETVA','RIE')  ;
RefI:=RechDom('GCNATUREPIECEG',TOBPiece.GetValue('GP_NATUREPIECEG'),False)+' N° '+IntToStr(TOBPiece.GetValue('GP_NUMERO')) ;
TOBE.PutValue('E_REFINTERNE',RefI) ;
TOBE.PutValue('E_ETAT','0000000000') ;
{Tables libres}
TOBE.PutValue('E_LIBREDATE',TOBPiece.GetValue('GP_DATELIBREPIECE1')) ;
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
   END ;
TOBG.PutValue('G_NATUREGENE',NatG) ; TOBG.PutValue('G_LIBELLE',LibG) ;
TOBG.PutValue('G_TVA',CodeTVA) ; TOBG.PutValue('G_TPF',CodeTPF) ;
TOBG.PutValue('G_ABREGE',Abr) ;
Result:=TOBG.InsertDB(Nil) ;
if Not Result then BEGIN TOBG.Free ; TOBG:=Nil ; END ;
END ;

Function CreerLignesTiers ( TOBEcr,TOBPiece,TOBEches,TOBTiers : TOB ; MM : RMVT ) : T_RetCpta ;
Var TOBTTC,TOBG,TOBH : TOB ;
    i,NumE,NumL  : integer ;
    GColl  : String ;
    Q      : TQuery ;
    OkVent : boolean ;
    DP,CP,DD,CD,DE,CE,XD,XE,XP : Double ;
BEGIN
Result:=rcOk ;
{Etude du compte général collectif}
GColl:=TOBTiers.GetValue('T_COLLECTIF') ; TOBG:=Nil ;
Q:=OpenSQL('SELECT * FROM GENERAUX WHERE G_GENERAL="'+GColl+'"',True) ;
if Not Q.EOF then TOBG:=TOB.CreateDB('GENERAUX',Nil,-1,Q) ;
Ferme(Q) ;
// Erreur sur le collectif
if TOBG=Nil then BEGIN Result:=rcPar ; LastMsg:=4 ; Exit ; END ;
OkVent:=(TOBG.GetValue('G_VENTILABLE')='X') ;
{une seule ligne ecriture/eche pour le tiers}
TOBTTC:=TOB.Create('ECRITURE',TOBEcr,-1) ;
PieceVersECR(MM,TOBPiece,TOBTiers,TOBTTC) ;
{Tiers}
TOBTTC.PutValue('E_AUXILIAIRE',TOBTiers.GetValue('T_AUXILIAIRE')) ;
TOBTTC.PutValue('E_GENERAL',GColl) ;
TOBTTC.PutValue('E_CONSO',TOBTiers.GetValue('T_CONSO')) ;
{Pièce}
TOBTTC.PutValue('E_RIB',TOBPiece.GetValue('GP_RIB')) ;
{Divers}
TOBTTC.PutValue('E_TYPEMVT','TTC') ;
TOBTTC.PutValue('E_ETATLETTRAGE','AL') ;
TOBTTC.PutValue('E_EMETTEURTVA','X') ;
{Eche+Vent}
NumL:=1 ; NumE:=1 ;
TOBTTC.PutValue('E_ECHE','X') ;
if OkVent then TOBTTC.PutValue('E_ANA','X') ;
TOBTTC.PutValue('E_NUMLIGNE',NumL) ; TOBTTC.PutValue('E_NUMECHE',NumE) ;
{Echéances}
TOBTTC.PutValue('E_MODEPAIE',TOBH.GetValue('GPE_MODEPAIE')) ;
TOBTTC.PutValue('E_DATEECHEANCE',TOBH.GetValue('GPE_DATEECHE')) ;
{Montants}
RecupLesDC(TOB du bulletin ,XD,XP,XE) ;
if MM.Nature='FC' then BEGIN DD:=XD ; DP:=XP ; DE:=XE ; CD:=0 ; CP:=0 ; CE:=0 ; END else
if MM.Nature='FF' then BEGIN CD:=XD ; CP:=XP ; CE:=XE ; DD:=0 ; DP:=0 ; DE:=0 ; END else
if MM.Nature='AC' then BEGIN CD:=-XD ; CP:=-XP ; CE:=-XE ; DD:=0 ; DP:=0 ; DE:=0 ; END else
if MM.Nature='AF' then BEGIN DD:=-XD ; DP:=-XP ; DE:=-XE ; CD:=0 ; CP:=0 ; CE:=0 ; END ;
TOBTTC.PutValue('E_DEBIT',DP)     ; TOBTTC.PutValue('E_CREDIT',CP) ;
TOBTTC.PutValue('E_DEBITDEV',DD)  ; TOBTTC.PutValue('E_CREDITDEV',CD) ;
TOBTTC.PutValue('E_DEBITEURO',DE) ; TOBTTC.PutValue('E_CREDITEURO',CE) ;
if ((DP=0) and (CP=0)) then TOBTTC.Free else Inc(NbEches) ;
TOBG.Free ;
END ;


Procedure AjusteTaxesBases ( TOBPiece,TOBBases : TOB ; NbDec : integer ) ;
Type DPE = (D,P,E) ;
Var TaxesB,TaxesL : Array[DPE,1..5] of Double ;
    MaxT          : Array[1..5] of Double ;
    LeMax         : Array[1..5] of integer ;
    EcartD,EcartP,EcartE,XD,XP,XE : Double ;
    i,k,NumT : integer ;
    CumHT : T_CodeCpta ;
    CatT : String ;
    TOBB : TOB ;
BEGIN
FillChar(TaxesB,Sizeof(TaxesB),#0) ; FillChar(TaxesL,Sizeof(TaxesL),#0) ; FillChar(MaxT,Sizeof(MaxT),#0) ;
for i:=1 to 5 do LeMax[i]:=-1 ;
for i:=0 to TTA.Count-1 do
    BEGIN
    CumHT:=T_CodeCpta(TTA[i]) ;
    for k:=1 to 5 do
        BEGIN
        TaxesL[D,k]:=TaxesL[D,k]+CumHT.SommeTaxeD[k] ;
        TaxesL[P,k]:=TaxesL[P,k]+CumHT.SommeTaxeP[k] ;
        TaxesL[E,k]:=TaxesL[E,k]+CumHT.SommeTaxeE[k] ;
        if ((CumHT.SommeTaxeD[k]<>0) and (Abs(CumHT.SommeTaxeD[k])>MaxT[k]))
           then BEGIN MaxT[k]:=Abs(CumHT.SommeTaxeD[k]) ; LeMax[k]:=i ; END ;
       END ;
    END ;
for i:=0 to TOBBases.Detail.Count-1 do
    BEGIN
    TOBB:=TOBBases.Detail[i] ;
    CatT:=TOBB.GetValue('GPB_CATEGORIETAXE') ; NumT:=Ord(CatT[3])-48 ;
    RecupLesDCTaxes(TOBB,XD,XP,XE) ;
    TaxesB[D,NumT]:=TaxesB[D,NumT]+XD ;
    TaxesB[P,NumT]:=TaxesB[P,NumT]+XP ;
    TaxesB[E,NumT]:=TaxesB[E,NumT]+XE ;
    END ;
{Ajustement final}
for k:=1 to 5 do
    BEGIN
    EcartD:=Arrondi(TaxesB[D,k]-TaxesL[D,k],NbDec) ;
    EcartP:=Arrondi(TaxesB[P,k]-TaxesL[P,k],V_PGI.OkDecV) ;
    EcartE:=Arrondi(TaxesB[E,k]-TaxesL[E,k],V_PGI.OkDecE) ;
    if ((EcartD<>0) or (EcartP<>0) or (EcartE<>0)) then if LeMax[k]>=0 then
       BEGIN
       CumHT:=T_CodeCpta(TTA[LeMax[k]]) ;
       CumHT.SommeTaxeD[k]:=Arrondi(CumHT.SommeTaxeD[k]+EcartD,NbDec) ;
       CumHT.SommeTaxeP[k]:=Arrondi(CumHT.SommeTaxeP[k]+EcartP,V_PGI.OkDecV) ;
       CumHT.SommeTaxeE[k]:=Arrondi(CumHT.SommeTaxeE[k]+EcartE,V_PGI.OkDecE) ;
       END ;
    END ;
END ;

Procedure AjusteHTBases ( TOBPiece,TOBBases : TOB ; NbDec : integer ) ;
Var TauxEsc,TauxRem,EscD,EscP,EscE,RemD,RemP,RemE : Double ;
    TotMHT_D,TotMHT_P,TotMHT_E,TotTPF_D,TotTPF_P,TotTPF_E,
      TotBase_D,TotBase_P,TotBase_E : Double ;
    i,k,kk : integer ;
    CumHT : T_CodeCpta ;
    TOBB : TOB ;
    XD,XP,XE,MaxM,EcartD,EcartP,EcartE : Double ;
    LeMax : integer ;
    FamBase : String ;
BEGIN
TauxEsc:=TOBPiece.GetValue('GP_ESCOMPTE') ; TauxRem:=TOBPiece.GetValue('GP_REMISEPIED') ;
RecupLesDCEsc(TOBPiece,EscD,EscP,EscE) ; RecupLesDCRem(TOBPiece,RemD,RemP,RemE) ;
if ((TauxRem<>0) or (TauxEsc<>0)) then
   BEGIN
   {Si Escompte ou remise pied alors ajustement en masse}
   TotMHT_D:=0 ; TotMHT_P:=0 ; TotMHT_E:=0 ;
   TotTPF_D:=0 ; TotTPF_P:=0 ; TotTPF_E:=0 ;
   TotBase_D:=0 ; TotBase_P:=0 ; TotBase_E:=0 ;
   MaxM:=0 ; LeMax:=-1 ;
   for i:=0 to TTA.Count-1 do
       BEGIN
       CumHT:=T_CodeCpta(TTA[i]) ;
       TotMHT_D:=TotMHT_D+CumHT.XD ; TotMHT_P:=TotMHT_P+CumHT.XP ; TotMHT_E:=TotMHT_E+CumHT.XE ;
       if Abs(CumHT.XD)>MaxM then BEGIN MaxM:=Abs(CumHT.XD) ; LeMax:=i ; END ;
       for k:=2 to 5 do
           BEGIN
           TotTPF_D:=TotTPF_D+CumHT.SommeTaxeD[k] ;
           TotTPF_P:=TotTPF_P+CumHT.SommeTaxeP[k] ;
           TotTPF_E:=TotTPF_E+CumHT.SommeTaxeE[k] ;
           END ;
       END ;
   for i:=0 to TOBBases.Detail.Count-1 do
       BEGIN
       TOBB:=TOBBases.Detail[i] ;
       if TOBB.GetValue('GPB_CATEGORIETAXE')='TX1' then
          BEGIN
          RecupLesDC(TOBB,XD,XP,XE) ;
          TotBase_D:=TotBase_D+XD ; TotBase_P:=TotBase_P+XP ; TotBase_E:=TotBase_E+XE ;
          END ;
       END ;
   {Ajustement final}
   EcartD:=(TotMHT_D+TotTPF_D-EscD-RemD)-TotBase_D ;
   EcartP:=(TotMHT_P+TotTPF_P-EscP-RemP)-TotBase_P ;
   EcartE:=(TotMHT_E+TotTPF_E-EscE-RemE)-TotBase_E ;
   if ((EcartD<>0) or (EcartP<>0) or (EcartE<>0)) then if LeMax>=0 then
      BEGIN
      CumHT:=T_CodeCpta(TTA[LeMax]) ;
      CumHT.XD:=Arrondi(CumHT.XD-EcartD,NbDec) ;
      CumHT.XP:=Arrondi(CumHT.XP-EcartP,V_PGI.OkDecV) ;
      CumHT.XE:=Arrondi(CumHT.XE-EcartE,V_PGI.OkDecE) ;
      END ;
   END else
   BEGIN
   {Si Ni Escompte Ni remise pied alors ajustement par base Tva}
   for k:=0 to TOBBases.Detail.Count-1 do
       BEGIN
       TOBB:=TOBBases.Detail[k] ; if TOBB.GetValue('GPB_CATEGORIETAXE')<>'TX1' then Continue ;
       FamBase:=TOBB.GetValue('GPB_FAMILLETAXE') ;
       TotMHT_D:=0 ; TotMHT_P:=0 ; TotMHT_E:=0 ;
       TotTPF_D:=0 ; TotTPF_P:=0 ; TotTPF_E:=0 ;
       MaxM:=0 ; LeMax:=-1 ;
       for i:=0 to TTA.Count-1 do
           BEGIN
           CumHT:=T_CodeCpta(TTA[i]) ; if CumHT.FamTaxe[1]<>FamBase then Continue ;
           TotMHT_D:=TotMHT_D+CumHT.XD ; TotMHT_P:=TotMHT_P+CumHT.XP ; TotMHT_E:=TotMHT_E+CumHT.XE ;
           if Abs(CumHT.XD)>MaxM then BEGIN MaxM:=Abs(CumHT.XD) ; LeMax:=i ; END ;
           for kk:=2 to 5 do
               BEGIN
               TotTPF_D:=TotTPF_D+CumHT.SommeTaxeD[kk] ;
               TotTPF_P:=TotTPF_P+CumHT.SommeTaxeP[kk] ;
               TotTPF_E:=TotTPF_E+CumHT.SommeTaxeE[kk] ;
               END ;
           END ;
       {Ajustement final}
       RecupLesDC(TOBB,XD,XP,XE) ; TotBase_D:=XD ; TotBase_P:=XP ; TotBase_E:=XE ;
       EcartD:=(TotMHT_D+TotTPF_D)-TotBase_D ;
       EcartP:=(TotMHT_P+TotTPF_P)-TotBase_P ;
       EcartE:=(TotMHT_E+TotTPF_E)-TotBase_E ;
       if ((EcartD<>0) or (EcartP<>0) or (EcartE<>0)) then if LeMax>=0 then
          BEGIN
          CumHT:=T_CodeCpta(TTA[LeMax]) ;
          CumHT.XD:=Arrondi(CumHT.XD-EcartD,NbDec) ;
          CumHT.XP:=Arrondi(CumHT.XP-EcartP,V_PGI.OkDecV) ;
          CumHT.XE:=Arrondi(CumHT.XE-EcartE,V_PGI.OkDecE) ;
          END ;
       END ;
   END ;
END ;

Function CreerLignesEcrHT ( TOBEcr,TOBPiece,TOBTiers : TOB ; MM : RMVT ; IndiceCpta : integer ) : T_RetCpta ;
Var CumHT : T_CodeCpta ;
    CHT : String ;
    TOBG,TOBE : TOB ;
    Q    : TQuery ;
    XD,XP,XE : Double ;
    CD,CP,CE : Double ;
    DD,DP,DE : Double ;
    OkVent : boolean ;
    NumL,i : integer ;
    Sta  : String ;
BEGIN
Result:=rcOk ;
CumHT:=T_CodeCpta(TTA[IndiceCpta]) ;
// Erreur sur Compte HT
CHT:=CumHT.CptHT ; if CHT='' then BEGIN Result:=rcPar ; LastMsg:=5 ; Exit ; END ;
{Etude du compte général HT}
TOBG:=Nil ;
Q:=OpenSQL('SELECT * FROM GENERAUX WHERE G_GENERAL="'+CHT+'"',True) ;
if Not Q.EOF then TOBG:=TOB.CreateDB('GENERAUX',Nil,-1,Q) ;
Ferme(Q) ;
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
PieceVersECR(MM,TOBPiece,TOBTiers,TOBE) ;
{Général}
TOBE.PutValue('E_GENERAL',CHT) ;
TOBE.PutValue('E_CONSO',TOBG.GetValue('G_CONSO')) ;
{Divers}
TOBE.PutValue('E_TYPEMVT','HT') ;
TOBE.PutValue('E_QTE1',CumHT.Qte) ;
TOBE.PutValue('E_TVA',CumHT.FamTaxe[1]) ; TOBE.PutValue('E_TPF',CumHT.FamTaxe[2]) ;
if CumHT.LibU then TOBE.PutValue('E_LIBELLE',CumHT.LibArt)
              else TOBE.PutValue('E_LIBELLE',TOBG.GetValue('G_LIBELLE')) ;
{Eche+Vent}
NumL:=TOBEcr.Detail.Count-NbEches+1 ; TOBE.PutValue('E_NUMLIGNE',NumL) ;
if PremHT<0 then PremHT:=NumL ;
if OkVent then TOBE.PutValue('E_ANA','X') ;
{Montants}
XD:=CumHT.XD ; XP:=CumHT.XP ; XE:=CumHT.XE ;
if MM.Nature='FC' then BEGIN CD:=XD ; CP:=XP ; CE:=XE ; DD:=0 ; DP:=0 ; DE:=0 ; END else
if MM.Nature='FF' then BEGIN DD:=XD ; DP:=XP ; DE:=XE ; CD:=0 ; CP:=0 ; CE:=0 ; END else
if MM.Nature='AC' then BEGIN DD:=-XD ; DP:=-XP ; DE:=-XE ; CD:=0 ; CP:=0 ; CE:=0 ; END else
if MM.Nature='AF' then BEGIN CD:=-XD ; CP:=-XP ; CE:=-XE ; DD:=0 ; DP:=0 ; DE:=0 ; END ;
TOBE.PutValue('E_DEBIT',DP)     ; TOBE.PutValue('E_CREDIT',CP) ;
TOBE.PutValue('E_DEBITDEV',DD)  ; TOBE.PutValue('E_CREDITDEV',CD) ;
TOBE.PutValue('E_DEBITEURO',DE) ; TOBE.PutValue('E_CREDITEURO',CE) ;
if ((DP=0) and (CP=0)) then TOBE.Free ;
TOBG.Free ;
END ;

Function FabricSQLCompta ( FamArt,FamTiers,FamAff,Etabl,Regime : String ) : String ;
Var wArt,wTiers,wAff : String ;
BEGIN
if VH_GC.GCVentCptaArt then wArt:='AND GCP_COMPTAARTICLE="'+FamArt+'"' else wArt:='' ;
if VH_GC.GCVentCptaTiers then wTiers:='AND GCP_COMPTATIERS="'+FamTiers+'"' else wTiers:='' ;
if VH_GC.GCVentCptaAff then wAff:='AND GCP_COMPTAAFFAIRE="'+FamAff+'"' else wAff:='' ;
Result:='SELECT * FROM CODECPTA WHERE (GCP_ETABLISSEMENT="'+Etabl+'" OR GCP_ETABLISSEMENT="") '
       +'AND (GCP_REGIMETAXE="'+Regime+'" OR GCP_REGIMETAXE="") '
       +wArt+' '+wTiers+' '+wAff+' '
       +'ORDER BY GCP_REGIMETAXE DESC, GCP_ETABLISSEMENT DESC' ;
END ;

Function FindTOBCode ( TOBGC : TOB ; FamArt,FamTiers,FamAff,Etabl,Regime : String ) : TOB ;
Var TOBC : TOB ;
    i : integer ;
BEGIN
Result:=Nil ;
for i:=TOBGC.Detail.Count-1 downto 0 do
    BEGIN
    TOBC:=TOBGC.Detail[i] ;
    if ((VH_GC.GCVentCptaArt) and (TOBC.GetValue('GCP_COMPTAARTICLE')<>FamArt)) then Continue ;
    if ((VH_GC.GCVentCptaTiers) and (TOBC.GetValue('GCP_COMPTATIERS')<>FamTiers)) then Continue ;
    if ((VH_GC.GCVentCptaAff) and (TOBC.GetValue('GCP_COMPTAAFFAIRE')<>FamAff)) then Continue ;
    if ((TOBC.GetValue('GCP_REGIMETAXE')<>'') and (TOBC.GetValue('GCP_REGIMETAXE')<>Regime)) then Continue ;
    if ((TOBC.GetValue('GCP_ETABLISSEMENT')<>'') and (TOBC.GetValue('GCP_ETABLISSEMENT')<>Etabl)) then Continue ;
    Result:=TOBC ; Break ;
    END ;
END ;

Function  ChargeAjouteCompta ( TOBCpta,TOBPiece,TOBA,TOBTiers : TOB ; AvecAnal : boolean ) : TOB ;
Var VenteAchat,Nature,Regime,Etab,FamArt,FamTiers,FamAff,SQL : String ;
    NatV,sRang : String ;
    TOBC : TOB ;
    Q    : TQuery ;
BEGIN
Result:=Nil ;
Regime:=TOBPiece.GetValue('GP_REGIMETAXE')   ; Etab:=TOBPiece.GetValue('GP_ETABLISSEMENT') ;
FamTiers:=TOBTiers.GetValue('T_COMPTATIERS') ; FamArt:=TOBA.GetValue('GA_COMPTAARTICLE') ; FamAff:='' ;
Nature:=TOBPiece.GetValue('GP_NATUREPIECEG') ;
TOBC:=FindTOBCode(TOBCpta,FamArt,FamTiers,FamAff,Etab,Regime) ;
if TOBC<>Nil then BEGIN Result:=TOBC ; Exit ; END ;
SQL:=FabricSQLCompta(FamArt,FamTiers,FamAff,Etab,Regime) ;
Q:=OpenSQL(SQL,True) ;
if Not Q.EOF then
   BEGIN
   TOBC:=TOB.Create('CODECPTA',TOBCpta,-1) ; TOBC.SelectDB('',Q) ;
   TOBCpta.Detail.Sort('GCP_REGIMETAXE;GCP_ETABLISSEMENT') ;
   END ;
Ferme(Q) ;
if ((TOBC<>Nil) and (AvecAnal)) then
   BEGIN
   VenteAchat:=GetInfoParPiece(Nature,'GPP_VENTEACHAT') ;
   if VenteAchat='VEN' then NatV:='HV' else NatV:='HA' ;
   sRang:=IntToStr(TOBC.GetValue('GCP_RANG')) ;
   Q:=OpenSQL('SELECT * FROM VENTIL WHERE V_NATURE like "'+NatV+'%" AND V_COMPTE="'+sRang+'"',True) ;
   if Not Q.EOF then
      BEGIN
      TOBC.LoadDetailDB('VENTIL','','',Q,False,True) ;
      TOBC.Detail.Sort('V_NUMEROVENTIL') ;
      END ;
   Ferme(Q) ;
   END ;
Result:=TOBC ;
END ;

Procedure CumulAnal ( CumHT : T_CodeCpta ; TOBL : TOB ; XD,XP,XE : Double ) ;
Var XDT : T_DetAnal ;
    TOBA,TOBAL : TOB ;
    Pourc : Double ;
    Section,Ax      : String ;
    i,k,iTrouv : integer ;
BEGIN
if TOBL.Detail.Count<=0 then Exit ;
TOBAL:=TOBL.Detail[0] ;
for i:=0 to TOBAL.Detail.Count-1 do
    BEGIN
    TOBA:=TOBAL.Detail[i] ;
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
    XDT.MD:=XDT.MD+Pourc*XD/100 ; XDT.MP:=XDT.MP+Pourc*XP/100 ; XDT.ME:=XDT.ME+Pourc*XE/100 ;
    if iTrouv<0 then CumHT.Anal.Add(XDT) ;
    END ;
END ;

Function CreerLignesHT ( TOBEcr,TOBPiece,TOBBases,TOBTiers,TOBArticles,TOBCpta : TOB ; MM : RMVT ; NbDec : integer ) : T_RetCpta ;
Var i,k,itrouv,kk : integer ;
    TOBA,TOBL,TOBCode : TOB ;
    RefUnique         : String ;
    CptHT,LibArt      : String ;
    CumHT             : T_CodeCpta ;
    OkTX              : Boolean ;
    XD,XP,XE : Double ;
    XDT : T_DetAnaL ;
BEGIN
Result:=rcOk ;
for i:=0 to TOBPiece.Detail.Count-1 do
    BEGIN
    TOBL:=TOBPiece.Detail[i] ;
    {Tests préalables}
    RefUnique:=TOBL.GetValue('GL_ARTICLE') ; if RefUnique='' then Continue ;
    TOBA:=FindTOBArtRow(TOBPiece,TOBArticles,i+1) ; if TOBA=Nil then Continue ;
    LibArt:=Copy(TOBL.GetValue('GL_LIBELLE'),1,35) ;
    {Recherche du paramétrage compta associé, sinon l'ajouter}
    TOBCode:=ChargeAjouteCompta(TOBCpta,TOBPiece,TOBA,TOBTiers,False) ;
    if TOBCode<>Nil then
       BEGIN
       if ((MM.Nature='FC') or (MM.Nature='AC')) then CptHT:=TOBCode.GetValue('GCP_CPTEGENEVTE')
                                                 else CptHT:=TOBCode.GetValue('GCP_CPTEGENEACH') ;
       END else
       BEGIN
       if ((MM.Nature='FC') or (MM.Nature='AC')) then CptHT:=VH_GC.GCCpteHTVTE else CptHT:=VH_GC.GCCpteHTACH ;
       END ;
    // Erreur sur le compte HT article
    if CptHT='' then BEGIN Result:=rcPar ; LastMsg:=5 ; Break ; END ;
    {Sommations par compte et familles de taxes identiques}
    RecupLesDC(TOBL,XD,XP,XE) ;
    iTrouv:=-1 ;
    for k:=0 to TTA.Count-1 do
        BEGIN
        CumHT:=T_CodeCpta(TTA[k]) ;
        if CumHT.CptHT<>CptHT then Continue ;
        OkTx:=True ;
        for kk:=1 to 5 do if CumHT.FamTaxe[kk]<>TOBL.GetValue('GL_FAMILLETAXE'+IntToStr(kk)) then OkTx:=False ;
        if OkTX then BEGIN iTrouv:=k ; Break ; END ;
        END ;
    if iTrouv<0 then
       BEGIN
       CumHT:=T_CodeCpta.Create ; CumHT.CptHT:=CptHT ; CumHT.LibU:=True ; CumHT.LibArt:=LibArt ;
       for kk:=1 to 5 do CumHT.FamTaxe[kk]:=TOBL.GetValue('GL_FAMILLETAXE'+IntToStr(kk)) ;
       END else
       BEGIN
       CumHT:=T_CodeCpta(TTA[iTrouv]) ;
       END ;
    CumHT.XD:=CumHT.XD+XD ; CumHT.XP:=CumHT.XP+XP ; CumHT.XE:=CumHT.XE+XE ;
    CumHT.Qte:=CumHT.Qte+TOBL.GetValue('GL_QTEFACT') ;
    if CumHT.LibArt<>LibArt then BEGIN CumHT.LibU:=False ; CumHT.LibArt:='' ; END ;
    for kk:=1 to 5 do
        BEGIN
        CumHT.SommeTaxeD[kk]:=CumHT.SommeTaxeD[kk]+TOBL.GetValue('GL_TOTALTAXEDEV'+IntToStr(kk)) ;
        CumHT.SommeTaxeP[kk]:=CumHT.SommeTaxeP[kk]+TOBL.GetValue('GL_TOTALTAXE'+IntToStr(kk)) ;
        CumHT.SommeTaxeE[kk]:=CumHT.SommeTaxeE[kk]+TOBL.GetValue('GL_TOTALTAXECON'+IntToStr(kk)) ;
        END ;
    CumulAnal(CumHT,TOBL,XD,XP,XE) ;
    if iTrouv<0 then TTA.Add(CumHT) ;
    END ;
{Arrondissage des cumuls}
for i:=0 to TTA.Count-1 do
    BEGIN
    CumHT:=T_CodeCpta(TTA[i]) ;
    CumHT.XD:=Arrondi(CumHT.XD,NbDec) ; CumHT.XP:=Arrondi(CumHT.XP,V_PGI.OkDecV) ; CumHT.XE:=Arrondi(CumHT.XE,V_PGI.OkDecE) ;
    for kk:=1 to 5 do
        BEGIN
        CumHT.SommeTaxeD[kk]:=Arrondi(CumHT.SommeTaxeD[kk],NbDec) ;
        CumHT.SommeTaxeP[kk]:=Arrondi(CumHT.SommeTaxeP[kk],V_PGI.OkDecV) ;
        CumHT.SommeTaxeE[kk]:=Arrondi(CumHT.SommeTaxeE[kk],V_PGI.OkDecE) ;
        END ;
    for kk:=0 to CumHT.Anal.Count-1 do
        BEGIN
        XDT:=CumHT.Anal[kk] ;
        XDT.MD:=Arrondi(XDT.MD,NbDec) ;
        XDT.MP:=Arrondi(XDT.MP,V_PGI.OkDecV) ;
        XDT.ME:=Arrondi(XDT.ME,V_PGI.OkDecE) ;
        END ;
    END ;
{Ajustements par rapport au pied de facture}
AjusteTaxesBases(TOBPiece,TOBBases,NbDec) ;
AjusteHTBases(TOBPiece,TOBBases,NbDec) ;
{Création des lignes d'écriture}
for i:=0 to TTA.Count-1 do
    BEGIN
    if Result=rcOk then Result:=CreerLignesEcrHT(TOBEcr,TOBPiece,TOBTiers,MM,i) ;
    if Result<>rcOk then Break ;
    END ;
END ;

Procedure MakeContrePTiers ( TOBEcr,TOBTiers : TOB ) ;
Var TOBHT,TOBL : TOB ;
    i     : integer ;
    CptHT,Auxi : String ;
BEGIN
if PremHT<=0 then Exit ;
TOBHT:=TOBEcr.Detail[PremHT-1] ;
Auxi:=TOBTiers.GetValue('T_AUXILIAIRE') ; CptHT:=TOBHT.GetValue('E_GENERAL') ;
for i:=0 to TOBEcr.Detail.Count-1 do
    BEGIN
    TOBL:=TOBEcr.Detail[i] ;
    if TOBL.Getvalue('E_AUXILIAIRE')=Auxi then TOBL.PutValue('E_CONTREPARTIEGEN',CptHT) ;
    END ;
END ;

Procedure VentilerFromGC ( TOBEcr : TOB ; NbDec : integer ) ;
Var CumHT     : T_CodeCpta ;
    IndiceA,i,NumAxe,NumVentil,Nb1,Nb2 : integer ;
    XDT     : T_DetAnal ;
    Vent    : Array[1..5] of boolean ;
    TOBAxe,TOBAna  : TOB ;
    Axes,CptGene,Section : String ;
    Pourcentage   : Double ;
    fb       : TFichierBase ;
BEGIN
IndiceA:=TOBEcr.GetValue('ANAL') ; Axes:=TOBEcr.GetValue('AXES') ;
CumHT:=T_CodeCpta(TTA[IndiceA]) ; Nb1:=0 ; Nb2:=0 ;
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
       Section:=XDT.Section ; NumVentil:=i+1 ; Pourcentage:=Arrondi(100.0*XDT.MP/CumHT.XP,6) ;
       VentilLigneTOB(TOBAna,Section,NumVentil,NbDec,Pourcentage,TOBEcr.GetValue('E_DEBIT')<>0) ;
       END ;
   if Nb1<=0 then
      BEGIN
      TOBAna:=TOB.Create('ANALYTIQ',TOBAxe,-1) ; EcrVersAna(TOBEcr,TOBAna) ;
      Section:=VH^.Cpta[fbAxe1].Attente ; NumVentil:=1 ; Pourcentage:=100.0 ;
      VentilLigneTOB(TOBAna,Section,NumVentil,NbDec,Pourcentage,TOBEcr.GetValue('E_DEBIT')<>0) ;
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
       Section:=XDT.Section ; NumVentil:=i+1 ; Pourcentage:=Arrondi(100.0*XDT.MP/CumHT.XP,6) ;
       VentilLigneTOB(TOBAna,Section,NumVentil,NbDec,Pourcentage,TOBEcr.GetValue('E_DEBIT')<>0) ;
       END ;
   if Nb2<=0 then
      BEGIN
      TOBAna:=TOB.Create('ANALYTIQ',TOBAxe,-1) ; EcrVersAna(TOBEcr,TOBAna) ;
      Section:=VH^.Cpta[fbAxe2].Attente ; NumVentil:=1 ; Pourcentage:=100.0 ;
      VentilLigneTOB(TOBAna,Section,NumVentil,NbDec,Pourcentage,TOBEcr.GetValue('E_DEBIT')<>0) ;
      END ;
   END ;
for NumAxe:=3 to 5 do if Vent[NumAxe] then
    BEGIN
    fb:=AxeToFb('A'+IntToStr(NumAxe)) ;
    TOBAxe:=TOBEcr.Detail[NumAxe-1] ;
    if TOBAxe.Detail.Count<=0 then
       BEGIN
       TOBAna:=TOB.Create('ANALYTIQ',TOBAxe,-1) ; EcrVersAna(TOBEcr,TOBAna) ;
       Section:=VH^.Cpta[fb].Attente ; NumVentil:=1 ; Pourcentage:=100.0 ;
       VentilLigneTOB(TOBAna,Section,NumVentil,NbDec,Pourcentage,TOBEcr.GetValue('E_DEBIT')<>0) ;
       END ;
    END ;
ArrondirAnaTOB(TOBEcr,NbDec) ;
END ;

Procedure MakeAnalGC ( TOBEcr,TOBPiece,TOBTiers : TOB ; MM : RMVT ; NbDec : integer ) ;
Var TOBL,TOBAL,TOBV : TOB ;
    i,k  : integer ;
    RefA : String ;
BEGIN
for i:=0 to TOBEcr.Detail.Count-1 do
    BEGIN
    TOBL:=TOBEcr.Detail[i] ;
    if TOBL.GetValue('E_ANA')='X' then
       BEGIN
       for k:=1 to 5 do TOB.Create('A'+IntToStr(k),TOBL,-1) ;
       if TOBL.FieldExists('ANAL') then
          BEGIN
          VentilerFromGC(TOBL,NbDec) ;
          END else
          BEGIN
          VentilerTOB(TOBL,'',0,NbDec,False) ;
          END ;
       END ;
    END ;
RefA:=EncodeRefCPGescom(TOBPiece) ;
for i:=0 to TOBPiece.Detail.Count-1 do
    BEGIN
    TOBL:=TOBPiece.Detail[i] ; TOBAL:=TOBL.Detail[0] ;
    for k:=0 to TOBAL.Detail.Count-1 do
        BEGIN
        TOBV:=TOBAL.Detail[k] ;
        TOBV.PutValue('YVA_IDENTIFIANT',RefA) ;
        TOBV.PutValue('YVA_NATUREID','GC') ;
        TOBV.PutValue('YVA_IDENTLIGNE',FormatFloat('000',i+1)) ;
        END ;
    END ;
END ;

Function CreerLigneEcart ( TOBEcr,TOBPiece,TOBTiers : TOB ; MM : RMVT ; EcartD,EcartP,EcartE : Double ) : T_RetCpta ;
Var CptECC : String ;
    TOBE   : TOB ;
    NumL   : integer ;
    Debit  : boolean ;
BEGIN
Result:=rcOk ;
if EcartP<>0 then Debit:=(EcartP<0) else Debit:=(EcartE<0) ;
if Debit then
   BEGIN
   CptECC:=VH^.EccEuroDebit ;
   if CptECC='' then BEGIN CptECC:=VH^.EccEuroCredit ; EcartP:=-EcartP ; EcartE:=-EcartE ; EcartD:=-EcartD ; END ;
   END else
   BEGIN
   CptECC:=VH^.EccEuroCredit ;
   if CptECC='' then BEGIN CptECC:=VH^.EccEuroDebit ; EcartP:=-EcartP ; EcartE:=-EcartE ; EcartD:=-EcartD ; END ;
   END ;
// Erreur sur Compte de conversion
if CptECC='' then BEGIN Result:=rcPar ; LastMsg:=6 ; Exit ; END ;
{Ligne d'écriture}
TOBE:=TOB.Create('ECRITURE',TOBEcr,-1) ;
PieceVersECR(MM,TOBPiece,TOBTiers,TOBE) ;
{Général}
TOBE.PutValue('E_GENERAL',CptECC) ;
{Divers}
TOBE.PutValue('E_TYPEMVT','ECC') ;
TOBE.PutValue('E_LIBELLE','Ecart de conversion') ;
{Contreparties}
TOBE.PutValue('E_CONTREPARTIEGEN',TOBTiers.GetValue('T_COLLECTIF')) ;
TOBE.PutValue('E_CONTREPARTIEAUX',TOBTiers.GetValue('T_AUXILIAIRE')) ;
{Eche+Vent}
NumL:=TOBEcr.Detail.Count-NbEches+1 ; TOBE.PutValue('E_NUMLIGNE',NumL) ;
{Montants}
if Debit then
   BEGIN
   TOBE.PutValue('E_DEBIT',-EcartP) ; TOBE.PutValue('E_DEBITDEV',-EcartD)  ; TOBE.PutValue('E_DEBITEURO',-EcartE) ;
   END else
   BEGIN
   TOBE.PutValue('E_CREDIT',EcartP) ; TOBE.PutValue('E_CREDITDEV',EcartD) ; TOBE.PutValue('E_CREDITEURO',EcartE) ;
   END ;
END ;

Procedure AjusteSurTVA ( TOBEcr : TOB ; EcartD,EcartP,EcartE : Double ) ;
Var LaLig : integer ;
    TOBL : TOB ;
    DD,DP,DE,CD,CP,CE : Double ;
BEGIN
if PremTVA>0 then LaLig:=PremTVA else if PremHT>0 then LaLig:=PremHT else LaLig:=1 ;
TOBL:=TOBEcr.Detail[LaLig-1] ;
DD:=TOBL.GetValue('E_DEBITDEV')  ; CD:=TOBL.GetValue('E_CREDITDEV') ;
DP:=TOBL.GetValue('E_DEBIT')     ; CP:=TOBL.GetValue('E_CREDIT') ;
DE:=TOBL.GetValue('E_DEBITEURO') ; CE:=TOBL.GetValue('E_CREDITEURO') ;
if DD<>0 then
   BEGIN
   TOBL.PutValue('E_DEBITDEV',DD-EcartD) ; TOBL.PutValue('E_DEBIT',DP-EcartP) ; TOBL.PutValue('E_DEBITEURO',DE-EcartE) ;
   END else
   BEGIN
   TOBL.PutValue('E_CREDITDEV',CD+EcartD) ; TOBL.PutValue('E_CREDIT',CP+EcartP) ; TOBL.PutValue('E_CREDITEURO',CE+EcartE) ;
   END ;
END ;

Procedure EquilibreEcrGC ( TOBEcr,TOBPiece,TOBTiers : TOB ; MM : RMVT ; NbDec : integer ) ;
Var DD,DP,DE,CD,CP,CE : Double ;
    i          : integer ;
    ModeOppose : boolean ;
    TOBL       : TOB ;
    EcartE,EcartP,EcartD : Double ;
BEGIN
DD:=0 ; CD:=0 ; DP:=0 ; CP:=0 ; DE:=0 ; CE:=0 ;
ModeOppose:=(TOBPiece.GetValue('GP_SAISIECONTRE')='X') ;
for i:=0 to TOBEcr.Detail.Count-1 do
    BEGIN
    TOBL:=TOBEcr.Detail[i] ;
    DD:=DD+TOBL.GetValue('E_DEBITDEV')  ; CD:=CD+TOBL.GetValue('E_CREDITDEV') ;
    DP:=DP+TOBL.GetValue('E_DEBIT')     ; CP:=CP+TOBL.GetValue('E_CREDIT') ;
    DE:=DE+TOBL.GetValue('E_DEBITEURO') ; CE:=CE+TOBL.GetValue('E_CREDITEURO') ;
    TOBL.PutValue('E_ENCAISSEMENT',SensEnc(DP,CP)) ;
    END ;
EcartD:=Arrondi(DD-CD,NbDec) ; EcartP:=Arrondi(DP-CP,V_PGI.OkDecV) ; EcartE:=Arrondi(DE-CE,V_PGI.OkDecE) ;
if ((EcartD=0) and (EcartP=0) and (EcartE=0)) then Exit ;
if ((EcartD<>0) and (EcartE<>0) and (EcartP<>0)) then
   BEGIN
   if DelphiRunning then ShowMessage('Tout faux, Devise : '+StrfPoint(EcartD)+'   Pivot : '+StrfPoint(EcartP)+'   ContreV : '+StrfPoint(EcartE)) ;
   AjusteSurTVA(TOBEcr,EcartD,EcartP,EcartE) ;
   END else if ((ModeOppose) and (EcartE<>0)) or ((Not ModeOppose) and (EcartP<>0)) then
   BEGIN
   if DelphiRunning then ShowMessage('Saisie fausse, Devise : '+StrfPoint(EcartD)+'   Pivot : '+StrfPoint(EcartP)+'   ContreV : '+StrfPoint(EcartE)) ;
   AjusteSurTVA(TOBEcr,EcartD,EcartP,EcartE) ;
   END else if ((EcartE<>0) and (EcartP<>0)) then
   BEGIN
   if DelphiRunning then ShowMessage('Saisie devise juste et compta fausse, Devise : '+StrfPoint(EcartD)+'   Pivot : '+StrfPoint(EcartP)+'   ContreV : '+StrfPoint(EcartE)) ;
   AjusteSurTVA(TOBEcr,EcartD,EcartP,EcartE) ;
   END else
   BEGIN
   {Ecart de conversion}
   if VH_GC.GCEcartConvert<>'CPT' then AjusteSurTVA(TOBEcr,EcartD,EcartP,EcartE)
                                  else CreerLigneEcart(TOBEcr,TOBPiece,TOBTiers,MM,EcartD,EcartP,EcartE) ;
   END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 27/03/2000
Modifié le ... : 27/03/2000
Description .. : Passation comptable de la pièce Gescom
Mots clefs ... : LIAISON;COMPTABILITE;PONT;PASSATION;
*****************************************************************}
Function PasseEnCompta ( TOBPiece,TOBBases,TOBEches,TOBTiers,TOBArticles,TOBCpta : TOB ; NbDec : integer ; OldEcr : RMVT ) : T_RetCpta ;
Var QualifP,PassP,NatureG,NatCpta,Viv,JalCpta,RefCP : String ;
    TOBEcr,TOBT : TOB ;
    MM          : RMVT ;
    NumCpta     : Longint ;
BEGIN
Result:=rcOk ; LastMsg:=-1 ; NbEches:=0 ;
{TOB Ecriture Générique}
TOBEcr:=TOB.Create('COMPTABILITE',Nil,-1) ;
// Création du RECORD clef comptable
FillChar(MM,Sizeof(MM),#0) ; MM.Simul:=QualifP ;
RenseigneClefCompta(TOBPiece,MM) ; PremHT:=-1 ; PremTVA:=-1 ;
{Journal}

//***PG*** JalCpta:=Paramètre de la paie pour renseigner code journal
// Erreur sur Journal
if JalCpta='' then BEGIN TOBEcr.Free ; Result:=rcPar ; LastMsg:=7 ; Exit ; END ;
TOBPiece.PutValue('GP_JALCOMPTABLE',JalCpta) ; MM.Jal:=JalCpta ;
{Nature}
NatCpta:='OD' ; MM.Nature:=NatCpta ;
{Valide}
MM.Valide:=True ;
{Numéro}
if ((MM.Jal=OldEcr.Jal) and (MM.Simul=OldEcr.Simul) and (MM.Nature=OldEcr.Nature))
   then NumCpta:=OldEcr.Num else NumCpta:=GetNewNumJal(JalCpta,(MM.Simul='N'),MM.DateC) ;
// Erreur sur le numéro
if NumCpta<=0 then BEGIN TOBEcr.Free ; Result:=rcPar ; LastMsg:=9 ; Exit ; END ;
MM.Num:=NumCpta ;
{Préparer analytiques}
TTA:=TList.Create ;
{Lignes de Tiers}
Result:=CreerLignesTiers(TOBEcr,TOBPiece,TOBEches,TOBTiers,MM) ;
{Lignes HT}
if Result=rcOk then Result:=CreerLignesHT(TOBEcr,TOBPiece,TOBBases,TOBTiers,TOBArticles,TOBCpta,MM,NbDec) ;
if ((Result=rcOk) and (TOBEcr.Detail.Count>=2)) then
   BEGIN
   {Ultimes ajustements}
   MakeContrePTiers(TOBEcr,TOBTiers) ;
   {Analytiques}
   MakeAnalGC(TOBEcr,TOBPiece,TOBTiers,MM,NbDec) ;
   {Retour inverse}
   RefCP:=EncodeRefGCComptable(TOBEcr.Detail[0]) ;
   TOBPiece.PutValue('GP_REFCOMPTABLE',RefCP) ;
   {Equilibrage}
   EquilibreEcrGC(TOBEcr,TOBPiece,TOBTiers,MM,NbDec) ;
   {MAJ des soldes comptables}
   MajSoldesEcritureTOB(TOBEcr,True) ;
   {Insertion}
   if Not TOBEcr.InsertDBByNivel(False) then V_PGI.IoError:=oeUnknown ;
   END ;
TOBEcr.Free ;
VideListe(TTA) ; TTA.Free ; TTA:=Nil ;
END ;

Function PassationComptable ( TOBPiece,TOBBases,TOBEches,TOBTiers,TOBArticles,TOBCpta : TOB ; NbDec : integer ; OldEcr : RMVT ) : Boolean ;
Var Res : T_RetCpta ;
    Tex : String ;
BEGIN
Result:=False ;
Res:=PasseEnCompta(TOBPiece,TOBBases,TOBEches,TOBTiers,TOBArticles,TOBCpta,NbDec,OldEcr) ;
if LastMsg>0 then Tex:=TexteMessage[LastMsg] else Tex:='' ;
Case Res of
   rcOk  : Result:=True ;
   rcPar : HShowMessage('0;Paramétrage comptable incorrect;'+Tex+';W;O;O;O;','','') ;
   rcRef : HShowMessage('0;Passage en comptabilité impossible;'+Tex+';W;O;O;O;','','') ;
   END ;
END ;

end.
