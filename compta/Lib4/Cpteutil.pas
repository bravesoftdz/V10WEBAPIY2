{*** MODIF CA - 12/10/2001 }
{*
  - Modification paramètre Histo de type booléen par BalSit de type chaîne.
  si BalSit <> '', on bosse sur la table CBALSIT
  - Modification de la requête CBALSIT pour travailler en fonction de BSI_TYPEBAL
{**************************}

unit CPTEUTIL ;

interface


uses SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,
     Forms, Dialogs, StdCtrls, Spin,
{$IFDEF EAGLCLIENT}

{$ELSE}
     DB,
  {$IFNDEF DBXPRESS} {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF}, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     Hctrls,
     FileCtrl,
{$IFDEF VER150}
  variants,
{$ENDIF}
     ExtCtrls,
     HStatus,
     ComCtrls,
     HEnt1,
     Ent1,
     {$IFDEF MODENT1}
     CPTypeCons,
     CPProcMetier,
     CPVersion,
     {$ELSE}
     tCalcCum,
     {$ENDIF MODENT1}
     UTob,
     UlibExercice,
     UtilPGI,  // CA - 28/12/2001
     ParamSoc   ,UentCommun  //FP FQ16854  GetParamSocSecur
     ;

{ Pb sur V_PGI.DeviseLibre : A revalider }
{
select g_general , t_auxiliaire from generaux G, tiers T where t_collectif=g.g_general and
g_general="41100000" and exists ( Select e_general, e_auxiliaire, e_exercice from ecriture
  where e_general=G.G_general and e_auxiliaire=t.t_auxiliaire and e_exercice="003")
order by g_general, t_auxiliaire}

Type TTabNatLibre = Array[1..10] Of String ;
Type TOKNatLibre = Array[1..10] Of Boolean ;


function WhatTypeEcr(Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState) : SetttTypePiece ;
function WhatTypeEcr2(St : String ) : SetttTypePiece; // CA - 22/11/2001
Function WhatExiste(PreE,PreJ,PreB,Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState ;
                    LDate1,LDate2 : TDateTime ; CodeExo : String ; WhatQualif : Byte ;  DEV : String ; Axe : String = '' ;
                    FiltreSup : String = '') : String ;
Function WhatExisteNat(PreE,PreJ,PreB,PreJ1,Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState ;
                       LDate1,LDate2 : TDateTime ; CodeExo : String ; WhatQualif : Byte ; DEV : String) : String ;
Function WhatExisteMul(PreE,PreJ1,PreJ2,PreB,Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState ;
                       LDate1,LDate2 : TDateTime ; CodeExo,Where1,Where2 : String ; WhatQualif : Byte ; Axe : String = '') : String ;
Function WhatExisteMulNat(PreE,PreJ1,PreJ2,PreB,Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState ;
                       LDate1,LDate2 : TDateTime ; CodeExo,Where1,Where2 : String ; WhatQualif : Byte) : String ;
Function WhereSupp(P : String ; SetTyp : SetttTypePiece) : String ;
Function WhereSuppEdtTiers(Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState) : String ;
Function JalODA ( CodCpt : String3) : String3 ;
Procedure DecodeNatLibre(fb : TFichierBase ; St : String ; Var Nat : TTabNatLibre ; DefautPourDeb : Boolean) ;
Procedure DecodeOkNatLibre(St : String ; Var OkNatLibre : TOkNatLibre) ;
function  EnDevise(FiltreDev,DeviseEnPivot : Boolean) : Boolean ;
Procedure FactoriseComboTypeRub ( C : THValComboBox ) ;

Type TabDC4 = Array[1..4] Of TabDC ;
{ 0 : AN
  1 : Courantes
  2 : Simu
  3 : Prevision
  4 : Situation
  5 : Révision
  6 : Ce que on ne trouve pas
  7 : Total de 1..5 (0 = AN Ne concerne que les écritures normales !!! + Flip Flop 1 <-> 7
}

{
X un : Calcul d'un cumul pour une fiche de base. La requete est réinitilisé à chaque appel.
X Deux : Calcul d'un cumul pour une fiche composite. La requete est réinitilisé à chaque appel.
 UnSurTableLibre : Calcul d'un cumul pour une table libre. La requete est réinitilisé à chaque appel.
 DeuxSurTableLibre : Calcul d'un cumul pour une table composite. La requete est réinitilisé à chaque appel.
 CPCpteJal : Calcul d'un cumul pour una journal et une racine de généraux. La requete est réinitilisé à chaque appel.
 DeuxSansCumul : Calcul d'un cumul pour tous les journaux et comptes. La requète n'est lancée q'uune fois et reste "vivante"
 CPUn : Calcul d'un cumul pour une racine de compte. La requète n'est lancée qu'une fois et reste "vivante"
 CPUnUDF : Calcul d'un cumul pour tous les élements d'une fiche de base. La requète n'est lancée qu'une fois et reste "vivante"
X unBud : Calcul d'un cumul pour une fiche de base budget. La requete est réinitilisé à chaque appel.
X DeuxBud : Calcul d'un cumul pour une fiche composite budget. La requete est réinitilisé à chaque appel.
 UnNatEcr : Calcul d'un cumul une nature d'écriture. La requete est réinitilisé à chaque appel.
 DeuxNatEcr : Calcul d'un cumul sur deux natures d'écriture. La requete est réinitilisé à chaque appel.
}

tCalcOkNatlibre = Record
                  OkNat1,OkNat2 : TOkNatLibre ;
                  End ;

Type tCalcNatLibre = Record
                     CptDeb1,CptFin1,CptDeb2,CptFin2 : String ;
                     NatDeb1,NatFin1,NatDeb2,NatFin2 : TTabNatLibre ;
                     End ;

Const MaxFiltreSup = 2 ;

Type tUnFiltreSup = Record
                    NomChamp : String ;
                    ValChamp : String ;
                    End ;

Type tFiltreSup = Array[0..MaxFiltreSup] Of tUnFiltreSup ;

Type tPrepCalc = Record
                 fb1,fb2 : TFichierBase ;
                 SetTyp : SetttTypePiece ;
                 FiltreDev,FiltreEtab,FiltreExo : Boolean ;
                 DeviseEnPivot : Boolean ;
                 Decimale,DecimaleEuro : Integer ;
                 EnEuro : Boolean ;
                 FiltreCpt1,FiltreCpt2 : Boolean ;
                 Tous1,Tous2 : Boolean ;
                 Methode : TTypeCalc ;
                 tz1,tz2 : TZoomTable ;
                 OkNatLibre : TCalcOKNatLibre ;
                 SpeedActif : Boolean ;
//                 Histo : Boolean ;
                  BalSit : string;
                 FiltreSup : String ;
                 ChampSup : String ;
                 End ;

Type tExecCalc = Record
                 Cpt1,Cpt2,Axe1,Axe2 : String ;
                 CalcNatLibre : tCalcNatLibre ;
                 Date1,Date2 : TDateTime ;
                 Devise,Etab,Exo : String3 ;
                 TotCpt : TabTot ;
                 TotCptEuro : TabTot ;
                 SommeTout : Boolean ;
                 QIsActive : Boolean ;
                 Params : ZSParams ;
                 BalSit : boolean; // CA - 21/11/2001 : idicateur de balance de situation
                 ValChampSup : Variant ;
                 End ;

Type TGCalculCum = Class
                   Multiple : TTypeCalc ;
                   Q, QBalSit : TQuery ; // CA - 21/11/2001
                   PrepCalc : tPrepCalc ;
                   ExecCalc : tExecCalc ;
//                   TOBCum : TOB ;
                   UseTC : Boolean ;
                   fstBase : string; // CA - 17/04/2007 : pour préciser la base concernée dans le cas du multisoc
//                   Constructor Create(FMultiple : TTypeCalc ; Ffb1,Ffb2 : TFichierBase ; FSetTyp : SetttTypePiece ; FFiltreDev,FFiltreEtab,FFiltreExo,FDeviseEnPivot,FEnEuro : Boolean ; FDec,FDecE : Integer ;
//                                      FSpeedActif : Boolean = FALSE ; FHisto : Boolean = FALSE ; FFiltreSup : String = '') ;
                   Constructor Create(FMultiple : TTypeCalc ; Ffb1,Ffb2 : TFichierBase ; FSetTyp : SetttTypePiece ; FFiltreDev,FFiltreEtab,FFiltreExo,FDeviseEnPivot,FEnEuro : Boolean ; FDec,FDecE : Integer ;
                                      FSpeedActif : Boolean = FALSE ; FBalSit : string = '' ; FFiltreSup : String = '' ;
                                      FChampSup : String = ''; FBase : string='') ;
                   Constructor CreateNat(FMultiple : TTypeCalc ; Ffb1,Ffb2 : TFichierBase ;
                                      FSetTyp : SetttTypePiece ; FFiltreDev,FFiltreEtab,FFiltreExo,FDeviseEnPivot,FEnEuro : Boolean ;
                                      FDec,FDecE : Integer ; Ftz1,Ftz2 : TZoomTable) ;
                   Procedure InitCalcul(C1,C2,A1,A2,DE,ET,EX : String ; D1,D2 : TDateTime ; ST : Boolean;
                                        SurBalSit : boolean = False) ; // CA - 21/11/2001
                   Procedure InitCalculVCS(C1,C2,A1,A2,DE,ET,EX : String ; D1,D2 : TDateTime ; ST : Boolean;
                                           VCS : Variant ; SurBalSit : boolean = False) ; // CA - 21/11/2001
                   Procedure ReInitCalcul(C1,C2 : String ; D1,D2 : TDateTime; Exercice : String3 = '';
                                          SurBalSit : boolean = False) ;  // CA - 21/11/2001
                   Procedure ReInitCalculVCS(C1,C2 : String ; D1,D2 : TDateTime; VCS : Variant ;
                                            Exercice : String3 = ''; SurBalSit : boolean = False) ;  // CA - 21/11/2001
                   Procedure Calcul ;
                   Destructor Destroy ; override ;
                   end ;


Function PrepareTotCpt(fb : TFichierBase ; SetTyp : SetttTypePiece ; FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot : Boolean ; Multiple : TTypeCalc; stBase : string = '') : TQuery ;
procedure ExecuteTotCpt(Var Q : TQuery ; Cpt : String17 ; Date1,Date2 : TDateTime ; Devise,Etab,Exo : String3 ;
                        Var TotCpt,TotCptEuro : TabTot ; SommeTout : Boolean ; Dec,DecE : Integer ;
                        fb : tfichierbase ; SetType : SetttTypePiece ; Multiple : TTypeCalc ; DeviseEnPivot : Boolean; stBase : string = '') ;

Function PrepareTotCptSolde(SetTyp : SetttTypePiece ; FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot : Boolean; stBase : string = '') : TQuery ;
procedure ExecuteTotCptSolde(Var Q : TQuery ; Cpt1 : String17 ; Date1,Date2 : TDateTime ; Devise,Etab,Exo : String3 ;
                             Var TotCptMvt,TotCptSld,TotCptMvtEuro,TotCptSldEuro : TabTot ; EnSoldeTotal : Boolean ; SommeTout : Boolean ;
                             Dec,DecE : Integer ; MonnaieOpposee : Boolean ; DeviseEnPivot : Boolean ; SetTyp : SetttTypePiece; stBase : string = '') ;

Function PrepareTotJal(CodCpt:String3 ; SetTyp : SetttTypePiece ; FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot : Boolean; stBase : string = '') : TQuery ;
procedure ExecuteTotJal(Var Q : TQuery ; Cpt : String17 ; Date1,Date2 : TDateTime ; Devise,Etab,Exo : String3 ;
                        Var TotCpt,TotCptEuro : TabTot ; SommeTout : Boolean ; Dec,DecE : Integer ;
                        SetType : SetttTypePiece ; DeviseEnPivot : Boolean) ;

Procedure CumulVersSolde(Var Cum : TabDC) ;

Function PrepareTotCptJointure(fb : TFichierBase ; SetTyp : SetttTypePiece ; FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot : Boolean ;
                               PrefixeJointure,ChampJointure,NomChampJointure,ValeurChampJointure : String; stBase : string = '') : TQuery ;
procedure ExecuteTotCptJointure(Var Q : TQuery ; Date1,Date2 : TDateTime ; Devise,Etab,Exo : String3 ;
                                Var TotCpt : TabTot ; SommeTout : Boolean ; Dec,DecE : Integer ;
                                fb : tfichierBase ; SetTyp : SetttTypePiece ; DeviseEnPivot : Boolean ;
                                PrefixeJointure,ChampJointure,NomChampJointure,ValeurChampJointure : String; stBase : string = '') ;

Function  GetDate(s : string) : TDateTime ;


////////////////////////////////////

//CETTE PARTIE A ETE DEPLACE DANS ULIBEXERCICE à L'ENLEVER
{$IFNDEF COMSX}
Function DansExo(Exo : TExoDate ; Date1,Date2 : TDateTime) : Boolean ;
Function  QUELEXODATE(Date1,Date2 : TDateTime ; Var MonoExo : Boolean ; Var Exo : TExoDate) : Boolean ;
Function QuelExoDate2(Date1,Date2 : TDateTime ; Var MonoExo : Boolean ; Var Exo1,Exo2 : TExoDate) : Boolean ;
function  WhatDate(St : String ; Var DD1,DD2 : tDateTime ; Var Err : Integer ; Var Exo : TExoDate) : boolean ;
{$ENDIF}
////////////////////////////////////


{b FP FQ16854: Identique à CPBALANCE_TOF, il faudrait placer cette classe dans une unité commune}
type
  TSQLAnaCroise = class
  private
    FAxes: array[1..MaxAxe] of Boolean; // Axes ventilables ?
    FPremierAxe: Integer;

    procedure LoadInfo;
    function  GetPremierAxe: String;
    function  AxeToSousPlan(NatureCpt: String): Integer;  {A partir de l'axe, retourne le numéro du sous plan}
  public
    constructor Create;
    destructor  Destroy; override;

    function GetConditionAxe(NatureCpt: String): String;
    function GetChampSection(NatureCpt: String): String;

    class function ConditionAxe(NatureCpt: String): String;
    class function ChampSection(NatureCpt: String): String;
  end;
{e FP FQ16854}

implementation

uses LicUtil,
     HMsgBox ;


Const TMaxDuree=1/24 ;
      MaxTaille=100000 ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 18/06/2002
Modifié le ... : 20/06/2002
Description .. : Cree la Tob des données correspondantes aux paramètres
Suite ........ : utilisateur
Mots clefs ... :
*****************************************************************}
Procedure AlimenteTob(LaFille :Tob ; LeMultiple : TTypeCalc ; LeFicBase1,LeFicBase2 : TFichierBase ; Dev,Etab,Exo : String ; D1,D2 : TDateTime ;
                      FFiltresupp, ChampSup : String ; DeviseEnPivot,FEnEuro : Boolean ; FSetTyp : SetttTypePiece; stBase : string ) ;
Var St,Sum_Montant : String ;
    Q : TQuery ;
    FiltreDev,FiltreEtab,FiltreExo : Boolean ;
    S,Prefixe : String ;
    Ana : Boolean ;
    LeOrder : String ;
    SecondChamps : Boolean ;
begin
Ana:=( ((LeFicBase1 in [FbAxe1..FbAxe5]) and (LeMultiple=un)) or
      (((LeFicBase1 in [FbAxe1..Fbaxe5]) or (LeFicBase2 in [FbAxe1..Fbaxe5])) and (LeMultiple=Deux)) )  ;
FiltreDev:=(Dev<>'') ;
FiltreEtab:=(Etab<>'') ;
FiltreExo:=(Exo<>'') ;
S:='' ;
Prefixe:='E_' ;

If Ana then Prefixe:='Y_' Else Prefixe:='E_' ;

SecondChamps:=(LeMultiple=Deux) ;

St:='' ; LeOrder:='' ;

If LeFicBase1=FbGene then St:=Prefixe+'GENERAL' ;
If LeFicBase1=FbAux then St:=Prefixe+'AUXILIAIRE';
{b FP FQ16854}
//If LeFicBase1 in [FbAxe1..FbAxe5] then St:=Prefixe+'SECTION' ;
If (LeFicBase1 in [FbAxe1..FbAxe5]) then St:=TSQLAnaCroise.ChampSection(FbToAxe(LeFicBase1));
{e FP FQ16854}

If ChampSup <> '' then S:='SELECT '+ChampSup+', '+St+', ' else S:='SELECT '+St+', ' ;
LeOrder:=St; St:='' ;

If SecondChamps then
  begin
    If LeFicBase2=FbGene then St:=Prefixe+'GENERAL' ;
    If LeFicBase2=FbAux then St:=Prefixe+'AUXILIAIRE';
    {b FP FQ16854}
    //If LeFicBase2 in [FbAxe1..FbAxe5] then St:=Prefixe+'SECTION' ;
    If LeFicBase2 in [FbAxe1..FbAxe5] then St:=TSQLAnaCroise.ChampSection(FbToAxe(LeFicBase2));
    {e FP FQ16854}
    LeOrder:=LeOrder+', '+St ;
    S:=S+St+', ' ;
    st:='' ;
  end ;

S:=S+Prefixe+'EXERCICE, ' ;

If (not Ana) then
  If EnDevise(FiltreDev,DeviseEnPivot) Then SUM_Montant:='sum(E_DEBITDEV) Debit, sum(E_CREDITDEV) Credit, 0 DebitEuro, 0 CreditEuro, '
                                       Else SUM_Montant:='sum(E_DEBIT) Debit, sum(E_CREDIT) Credit, 0 DebitEuro, 0 CreditEuro, '
             else
  If EnDevise(FiltreDev,DeviseEnPivot) Then SUM_Montant:='sum(Y_DEBITDEV) Debit, sum(Y_CREDITDEV) Credit, 0 DebitEuro, 0 CreditEuro,  '
                                       Else SUM_Montant:='sum(Y_DEBIT) Debit, sum(Y_CREDIT) Credit, 0 DebitEuro, 0 CreditEuro, ' ;

If Ana then St:=SUM_MONTANT+'Y_ECRANOUVEAU, Y_QUALIFPIECE FROM ANALYTIQ '
       else St:=SUM_MONTANT+'E_ECRANOUVEAU, E_QUALIFPIECE FROM ECRITURE ' ;
S:=S+St ; st:='' ;

//Le Where
S:=S+' Where '+Prefixe+'DATECOMPTABLE >="'+USDateTime(D1)+'" And '+Prefixe+'DATECOMPTABLE <= "'+UsDateTime(D2)+'" ' ;

{b FP FQ16854}
//If Ana then if (LeFicBase1 in [FbAxe1..fbAxe5]) then st:='AND Y_AXE="'+FbToAxe(LeFicbase1)+'" '
//                                                else st:='AND Y_AXE="'+FbToAxe(LeFicBase2)+'" ' ;
If Ana then if (LeFicBase1 in [FbAxe1..fbAxe5]) then st:='AND '+TSQLAnaCroise.ConditionAxe(FbToAxe(LeFicbase1))+' '
                                                else st:='AND '+TSQLAnaCroise.ConditionAxe(FbToAxe(LeFicBase2))+' ';
{e FP FQ16854}
S:=S+St ; St:='';

If FiltreDev Then St:='AND '+Prefixe+'DEVISE="'+DEV+'" ' ;
S:=S+St ; st:='' ;

If FiltreEtab Then St:='AND '+Prefixe+'ETABLISSEMENT="'+Etab+'" ' ;
S:=S+St ; St:='' ;

If FiltreExo Then St:='AND '+Prefixe+'EXERCICE="'+EXO+'"' ;
S:=S+St ; St:='' ;

If FFiltreSupp <> '' then S:=S+' AND ('+FFiltreSupp+')' ;

St:=WhereSupp(Prefixe,FSetTyp) ;
If St<>'' Then S:=S+St ;
St:='' ;

if ana then
  begin
  end ;
// Le Group by + Le Order by
If Champsup <> '' then LeOrder:=ChampSup+', '+LeOrder ;
S:=S+' GROUP BY '+LeOrder+','+Prefixe+'EXERCICE, '+Prefixe+'ECRANOUVEAU,'+Prefixe+'QUALIFPIECE ';

S:=S+' ORDER BY '+LeOrder+','+Prefixe+'EXERCICE, '+Prefixe+'ECRANOUVEAU,'+Prefixe+'QUALIFPIECE ';

Q:=OpenSql(S,TRUE,-1,'',False,stBase) ;

LaFille.loadDetailDb('LaTitFillote','','',Q,True) ;
Ferme(Q) ;
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 18/06/2002
Modifié le ... :   /  /
Description .. : Verifie cohérence des données transmises
Mots clefs ... :
*****************************************************************}
Function VerifieValeur(Zon1,Zon2 : String ; FMultiple : TTypeCalc ; Ffb1,Ffb2 : TFichierBase ; D1,D2 :TDatetime ) : Boolean ;
Begin
Result:=True ;
If not (FMultiple in [un..deux]) then begin Result:=False ; Exit ; End ;
If Zon1='' then begin result:=False ; exit ; End ;
If ((Zon2='') and (FMultiple=deux)) then begin Result:=False ; Exit ; end ;
If not (Ffb1 in [FbAxe1..FbAux]) then begin result:=False; exit ; end ;
if D1<Idate1900 then begin Result:=False ; exit ; end ;
If D2>Idate2099 then begin Result:=False ; exit ; end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 20/06/2002
Modifié le ... :   /  /
Description .. : Calcul (en octets) de la taille de la Tob mère en mémoire
Mots clefs ... :
*****************************************************************}
Function CalculTailleTobMere(LaTobMere : Tob ; LeMultiple : TTypeCalc ; LeFicBase1, LeFicBase2 : TFichierBase) : Integer ;
var LaTaille, Longueur : Integer ;
    i, NbrLignes : Integer ;
    LaFille : Tob ;
begin
//Result :=0 ;
Longueur :=0 ; LaTaille :=0 ;

{$IFDEF NOVH}
If LeFicBase1=FbGene then Longueur := GetInfoCpta(FbGene).Lg ;
If LeFicbase1=FbAux  then Longueur := GetInfoCpta(FbAux ).Lg ;
If LeFicBase1=FbAxe1 then Longueur := GetInfoCpta(FbAxe1).Lg ;
If LeFicBase1=FbAxe2 then Longueur := GetInfoCpta(FbAxe2).Lg ;
If LeFicBase1=FbAxe3 then Longueur := GetInfoCpta(FbAxe3).Lg ;
If LeFicBase1=FbAxe4 then Longueur := GetInfoCpta(FbAxe4).Lg ;
If LeFicBase1=FbAxe5 then Longueur := GetInfoCpta(FbAxe5).Lg ;
If LeMultiple=Deux then
  begin
  If LeFicBase2=FbGene then Longueur:=Longueur + GetInfoCpta(FbGene).Lg ;
  If LeFicbase2=FbAux  then Longueur:=Longueur + GetInfoCpta(FbAux ).Lg ;
  If LeFicBase2=FbAxe1 then Longueur:=Longueur + GetInfoCpta(FbAxe1).Lg ;
  If LeFicBase2=FbAxe2 then Longueur:=Longueur + GetInfoCpta(FbAxe2).Lg ;
  If LeFicBase2=FbAxe3 then Longueur:=Longueur + GetInfoCpta(FbAxe3).Lg ;
  If LeFicBase2=FbAxe4 then Longueur:=Longueur + GetInfoCpta(FbAxe4).Lg ;
  If LeFicBase2=FbAxe5 then Longueur:=Longueur + GetInfoCpta(FbAxe5).Lg ;
  end ;
{$ELSE}
If LeFicBase1=FbGene then Longueur:=VH^.Cpta[FbGene].Lg ;
If LeFicbase1=FbAux then Longueur:=VH^.Cpta[FbAux].Lg ;
If LeFicBase1=FbAxe1 then Longueur:=VH^.Cpta[FbAxe1].Lg ;
If LeFicBase1=FbAxe2 then Longueur:=VH^.Cpta[FbAxe2].Lg ;
If LeFicBase1=FbAxe3 then Longueur:=VH^.Cpta[FbAxe3].Lg ;
If LeFicBase1=FbAxe4 then Longueur:=VH^.Cpta[FbAxe4].Lg ;
If LeFicBase1=FbAxe5 then Longueur:=VH^.Cpta[FbAxe5].Lg ;
If LeMultiple=Deux then
  begin
  If LeFicBase2=FbGene then Longueur:=Longueur+VH^.Cpta[FbGene].Lg ;
  If LeFicbase2=FbAux then Longueur:=Longueur+VH^.Cpta[FbAux].Lg ;
  If LeFicBase2=FbAxe1 then Longueur:=Longueur+VH^.Cpta[FbAxe1].Lg ;
  If LeFicBase2=FbAxe2 then Longueur:=Longueur+VH^.Cpta[FbAxe2].Lg ;
  If LeFicBase2=FbAxe3 then Longueur:=Longueur+VH^.Cpta[FbAxe3].Lg ;
  If LeFicBase2=FbAxe4 then Longueur:=Longueur+VH^.Cpta[FbAxe4].Lg ;
  If LeFicBase2=FbAxe5 then Longueur:=Longueur+VH^.Cpta[FbAxe5].Lg ;
  end ;
{$ENDIF NOVH}

for i:=0 To LaTobMere.detail.Count-1 do //Calcul de la taille de la Tob Mere
  begin
  LaFille:=LaTobMere.Detail[i] ; NbrLignes:=LaFille.Detail.Count ;
  LaTaille:=LaTaille+((29+Longueur)*NbrLignes) ;
  end ;
Result:=LaTaille ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 20/06/2002
Modifié le ... :   /  /
Description .. : libere en fonction de criteres des  filles de la Tob Mere
Mots clefs ... :
*****************************************************************}
Procedure RearrangeLaTobMere(LaTobMere : Tob ; LeMultiple : TTypeCalc ; Ffb1, Ffb2 : TFichierBase) ;
var LaTaille : Integer ;
    TableAna, TableAnaTob : Boolean ;
    i, j, NbrAcces, NBA, Idx : Integer ;
    TF :Tob ;
    T1,T2 : TFichierBase ;
begin
Idx:=0;
LaTaille:=CalculTailleTobMere(LaTobMere,LeMultiple, Ffb1, Ffb2)  ;
If LaTaille<=MaxTaille then exit ;

TableAna:=( ((Ffb1 in [FbAxe1..FbAxe5]) and (LeMultiple=un)) or
           (((Ffb1 in [FbAxe1..Fbaxe5]) or (Ffb2 in [FbAxe1..Fbaxe5])) and (LeMultiple=Deux)) )  ;
//TrieLaTobMere (LaTobMere) ;

For i:=LaTobMere.Detail.Count-1 DownTo 0 do
  begin
  While LaTaille>=MaxTaille do
    begin
    TF:=LaTobMere.Detail[i] ;
    T1:=TF.GetValue('FicBase1') ; T2:=TF.GetValue('FicBase2') ;
    TableAnaTob:=( (( T1 in [FbAxe1..FbAxe5]) and (TF.GetValue('Multiple')=un)) or
                  (((T1 in [FbAxe1..Fbaxe5]) or (T2 in [FbAxe1..Fbaxe5])) and (TF.GetValue('Multiple')=Deux)) )  ;

    If TableAna=TableAnaTob then
      begin
      NbrAcces:=TF.GetValue('NombreAcces') ;
      For j:=0 to LaTobMere.Detail.Count-1 do
        begin
        NBA:=LaTobMere.Detail[j].GetValue('NombreAcces') ;
        if NBA<NbrAcces then begin Idx:=j ; NbrAcces:=NBA ;end ;
        end ;
      TF:=LaTobMere.Detail[idx] ;
      Tf.Free ; {Tf:=Nil ;} LaTaille:=CalculTailleTobMere(LaTobMere,LeMultiple, Ffb1, Ffb2) ;
      end else
      begin
      Tf.Free ; {Tf:=Nil ;} LaTaille:=CalculTailleTobMere(LaTobMere,LeMultiple, Ffb1, Ffb2) ;
      end ;

    end ;
  end ;
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 18/06/2002
Modifié le ... :   /  /
Description .. : Cree la ligne de tob avec les paramètres de base transmis
Mots clefs ... :
*****************************************************************}
Function CreeLaFille(LaTobMere : Tob ; LeMultiple : TTypeCalc ; LeFicBase1,LeFicBase2 : TFichierBase ; Dev,Etab,Exo : String ; D1,D2 : TDateTime ; FFiltresupp,ChampSup : String ; DeviseEnPivot,FEnEuro : Boolean ; FSetTyp : SetttTypePiece; stBase : string) : Tob ;
var LaFille{, LaPetiteFille}  : Tob ;
begin

RearrangeLaTobMere(LaTobMere, LeMultiple, LeFicBase1,LeFicBase2) ;

LaFille:=Tob.create('LaFille',LaTobMere,-1) ;

LaFille.AddChampSup('Multiple',False) ; LaFille.Putvalue('Multiple',LeMultiple) ;
LaFille.AddChampSup('FicBase1',False); LaFille.putValue('FicBase1',LeFicBase1) ;
LaFille.AddChampSup('FicBase2',False); LaFille.putValue('FicBase2',LeFicBase2) ;
LaFille.AddChampSup('Dev',False) ; LaFille.PutValue('Dev',Dev) ;
LaFille.AddChampSup('Etab',False) ; LaFille.PutValue('Etab',Etab) ;
LaFille.AddChampSup('Exo',False) ; LaFille.PutValue('Exo',Exo) ;
LaFille.AddChampSup('Date1',False) ; LaFille.PutValue('Date1',D1) ;
LaFille.AddChampSup('Date2',False) ; LaFille.PutValue('Date2',D2) ;
LaFille.AddChampSup('Filtresupp',False) ; LaFille.PutValue('FiltreSupp',FFiltreSupp) ;
LaFille.AddChampSup('ChampSup',False) ; LaFille.PutValue('ChampSup',ChampSup) ;
LaFille.AddChampSup('NombreAcces',False) ; LaFille.PutValue('NombreAcces',1) ;
LaFille.AddChampSup('HeureAcces',False) ; LaFille.PutValue('HeureAcces',Now) ;

//If LaPetiteFille <>nil then begin LapetiteFille.Free ; LapetiteFille:=Nil end Else LapetiteFille:=Tob.create('LaTitFillote',Nil,-1) ;
//LaPetiteFille:=AlimenteTob(LaFille,LeMultiple,LeFicBase1,LeFicBase2,Dev,Etab,Exo,D1,D2,FFiltresupp, DeviseEnPivot,FEnEuro, FSetTyp ) ;
AlimenteTob(LaFille,LeMultiple,LeFicBase1,LeFicBase2,Dev,Etab,Exo,D1,D2,FFiltresupp,ChampSup,DeviseEnPivot,FEnEuro, FSetTyp, stBase ) ;
Result:=Lafille ;
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 18/06/2002
Modifié le ... :   /  /
Description .. : Recherche la tob des parametres (+données) ou la cree
Mots clefs ... :
*****************************************************************}
Function RechercheOuCreeLaFille(LaTobMere : Tob ;FMultiple : TTypeCalc ; Ffb1,Ffb2 : TFichierBase ; Dev,Etab,Exo : String ; D1,D2 :TDatetime ; DeviseEnPivot,FEnEuro : Boolean ; FSetTyp : SetttTypePiece ; FFiltreSupp, ChampSup : String; stBase : string ) :Tob ;
var LaFille : Tob ;
    Nbr : Integer ;
    St : String ;
begin
if FMultiple=Un then Ffb2:=Ffb1  ;

LaFille:=LaTobMere.FindFirst(['Multiple','FicBase1','FicBase2','Dev','Etab','Exo','Date1','Date2','Filtresupp','ChampSup'],
                             [FMultiple,Ffb1,Ffb2,Dev,Etab,Exo,D1,D2,FFiltresupp,ChampSup],False) ;

if LaFille<>Nil then
  begin
  Nbr:=LaFille.GetValue('NombreAcces') ;
  LaFille.PutValue('NombreAcces',Nbr+1);
  LaFille.PutValue('HeureAcces',Now);
  end else
  BEGIN
  {$IFDEF EAGLCLIENT}
  if V_PGI.Superviseur then
  {$ELSE}
  If V_PGI.PassWord=CryptageSt(DayPass(Date)) Then
  {$ENDIF}
    BEGIN
    St:='Non Trouvé : dev : '+dev+','+'Etab : '+Etab+', Exo : '+Exo+', Date1 :'+DateToStr(D1)+', Date2 :'+DateToStr(D2) ;
    HShowMessage('0;Non trouvé :;'+St+';E;O;O;O;','','') ;
    END ;
  LaFille:=CreeLaFille(LaTobMere,FMultiple,Ffb1,Ffb2,Dev,Etab,Exo,D1,D2,FFiltresupp,ChampSup,DeviseEnPivot,FEnEuro,FSetTyp ,stBase) ;
  END ;

Result:=LaFille ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 18/06/2002
Modifié le ... : 18/06/2002
Description .. : Recherche le cumul pour une ou un couple de valeurs
Mots clefs ... : CUMUL;GETCUMUL
*****************************************************************}
Function DonneCumul(LaTobMere : Tob ; Zon1,Zon2 : String ; FMultiple : TTypeCalc ; Ffb1,Ffb2 : TFichierBase ; FSetTyp : SetttTypePiece ;
                    Dev,Etab,Exo : String ; DeviseEnPivot,FEnEuro : Boolean ; D1,D2 :TDatetime ; ChampSup : Variant ; ValChampSup : Variant ; FFiltreSup : String =''; stBase : string = '' ) : TabTot ;
var tot: tabtot ;
    ch1, ch2, Prefixe, LeQualifiant, LEcrANouveau : String ;
    UnSeulChamp, CestOK : Boolean ;
    TobDet : Tob ;
    i: Integer ;
    LaTob: Tob ;
    Ana : Boolean ;
begin

//FillChar
For i:=0 to 7 do
  begin
  Tot[i].TotDebit:=0 ;
  Tot[i].TotCredit:=0 ;
  end ;
Result:=Tot ;

CestOk:=VerifieValeur(Zon1,Zon2,FMultiple,Ffb1,Ffb2,D1,D2) ;
if (not CestOk) then exit ;

Ana:=( ((Ffb1 in [FbAxe1..FbAxe5]) and (Fmultiple=un)) or
      (((Ffb1 in [FbAxe1..Fbaxe5]) or (Ffb2 in [FbAxe1..Fbaxe5])) and (FMultiple=Deux)) )  ;

LaTob:=RechercheOuCreeLaFille(LaTobMere, FMultiple, Ffb1,Ffb2, Dev,Etab,Exo, D1,D2, DeviseEnPivot,FEnEuro, FSetTyp, FFiltreSup, ChampSup, stBase ) ;

{$IFDEF TESTGV}
LaTobMere.saveToFile('c:\Halley\mere.txt',false,true,true) ;
{$ELSE}
{$ENDIF}

If Ana then Prefixe:='Y_' Else Prefixe:='E_' ;

Case Ffb1 Of
  FbGene : Ch1:=Prefixe+'GENERAL' ;
  FbAux  : Ch1:=Prefixe+'AUXILIAIRE';
  {b FP FQ16854}
  //fbAxe1..fbAxe5 : Ch1:=Prefixe+'SECTION' ;
  fbAxe1..fbAxe5 : Ch1:=TSQLAnaCroise.ChampSection(FbToAxe(Ffb1)) ;
  {e FP FQ16854}
  END ;
UnSeulChamp:=(FMultiple=Un) ;

If (not UnSeulChamp) then
  begin
    Case Ffb2 Of
    FbGene : Ch2:=Prefixe+'GENERAL' ;
    FbAux  : Ch2:=Prefixe+'AUXILIAIRE';
    {b FP FQ16854}
    //fbAxe1..fbAxe5 : Ch2:=Prefixe+'SECTION' ;
    fbAxe1..fbAxe5 : Ch2:=TSQLAnaCroise.ChampSection(FbToAxe(Ffb2));
    {e FP FQ16854}
    end ;
  end ;


If VarIsNull(ValChampSup) then
  begin
  if UnSeulChamp then TobDet:=LaTob.FindFirst([Ch1],[Zon1],true)
                 else TobDet:=LaTob.FindFirst([Ch1,Ch2],[Zon1,Zon2],true) ;
  end else
  begin
  if UnSeulChamp then TobDet:=LaTob.FindFirst([Ch1,ChampSup],[Zon1,ValChampSup],true)
                 else TobDet:=LaTob.FindFirst([Ch1,Ch2,ChampSup],[Zon1,Zon2,ValChampSup],true) ;
  end ;

while TobDet <> nil do
  begin
  LeQualifiant:=TobDet.GetValue(Prefixe+'QUALIFPIECE') ;
  LEcrANouveau:=TobDet.GetValue(Prefixe+'ECRANOUVEAU') ;

  If LEcrANouveau='N' then
    begin
    If LeQualifiant='N' then begin Tot[1].TotDebit:=TobDet.GetValue('DEBIT') ; Tot[1].TotCredit:=TobDet.GetValue('CREDIT') ; Tot[7].TotDebit:=Tot[7].TotDebit+Tot[1].TotDebit ; Tot[7].TotCredit:=Tot[7].TotCredit+Tot[1].TotCredit ; end else
      If LeQualifiant='S' then begin Tot[2].TotDebit:=TobDet.GetValue('DEBIT') ; Tot[2].TotCredit:=TobDet.GetValue('CREDIT') ; Tot[7].TotDebit:=Tot[7].TotDebit+Tot[2].TotDebit ; Tot[7].TotCredit:=Tot[7].TotCredit+Tot[2].TotCredit ; end else
        If LeQualifiant='P' then begin Tot[3].TotDebit:=TobDet.GetValue('DEBIT') ; Tot[3].TotCredit:=TobDet.GetValue('CREDIT') ; Tot[7].TotDebit:=Tot[7].TotDebit+Tot[3].TotDebit ; Tot[7].TotCredit:=Tot[7].TotCredit+Tot[3].TotCredit ; end else
          If LeQualifiant='U' then begin Tot[4].TotDebit:=TobDet.GetValue('DEBIT') ; Tot[4].TotCredit:=TobDet.GetValue('CREDIT') ; Tot[7].TotDebit:=Tot[7].TotDebit+Tot[4].TotDebit ; Tot[7].TotCredit:=Tot[7].TotCredit+Tot[4].TotCredit ; end else
            If LeQualifiant='R' then begin Tot[5].TotDebit:=TobDet.GetValue('DEBIT') ; Tot[5].TotCredit:=TobDet.GetValue('CREDIT') ; Tot[7].TotDebit:=Tot[7].TotDebit+Tot[5].TotDebit ; Tot[7].TotCredit:=Tot[7].TotCredit+Tot[5].TotCredit ; end else
                                     begin Tot[6].TotDebit:=TobDet.GetValue('DEBIT') ; Tot[6].TotCredit:=TobDet.GetValue('CREDIT') ; end ;
    end ;
  If ((LEcrANouveau='OAN') or (LEcrANouveau='H')) then begin Tot[0].TotDebit:=TobDet.GetValue('DEBIT') ; Tot[0].TotCredit:=TobDet.GetValue('CREDIT') ; end ;

    If VarIsNull(ValChampSup) then
      begin
      if UnSeulChamp then TobDet:=LaTob.FindNext([Ch1],[Zon1],true)
                     else TobDet:=LaTob.FindNext([Ch1,Ch2],[Zon1,Zon2],true) ;
      end else
      begin
      if UnSeulChamp then TobDet:=LaTob.FindNext([Ch1,ChampSup],[Zon1,ValChampSup],true)
                     else TobDet:=LaTob.FindNext([Ch1,Ch2,ChampSup],[Zon1,Zon2,ValChampSup],true) ;
      end ;

  end ;
Result:=Tot ;
end ;

Procedure TrieLaTobMere (LaTobMere : Tob ) ;
var TFi,TFj,TFi2,TFj2 : Tob ;
    i,j,NAi,NAj : Integer ;
begin
for i:=0 to LaTobMere.Detail.Count-1 do
  begin
  TFi:=LaTobMere.Detail[i] ; NAi:=TFi.GetValue('NombreAcces') ;
  for j:=1 to LaTobMere.Detail.count-1 do
    begin
    TFj:=LaTobMere.Detail[j] ; NAj:=TFj.GetValue('NombreAcces') ;
    If NAj>NAi then
      begin
      TFi2:=TFi ; TFj2:=TFj ; LaTobMere.Detail[i]:=TFj2 ; LaTobMere.Detail[j]:=TFi2 ;
      end ;
    end ;
  end ;
end ;



Function VCSOK(VCS : Variant) : Boolean ;
BEGIN
Case VarType(VCS) of
   varEmpty,varNull : Result:=FALSE
   Else Result:=TRUE ;
   End ;
END ;

Function GetDate(s : string) : TDateTime ;
var j,err : integer ;
    sd,sb : string ;
    OkWhatDate : boolean ;
    ddeb,dfin : TDateTime ;
    Exo : TExoDate ;
begin
result:=V_PGI.DateEntree ;
j := Pos('/',s) ;
if j>0 then
   begin
   sd := Uppercase(Copy(s,1,j-1)) ; // pour WhatDate
   sb := Uppercase(Copy(s,j+1,1)) ; // D (Début) ou F (Fin)
   OkWhatDate:=WhatDate(sd,ddeb,dfin,err,exo) ;
   if OKWhatDate then
      begin
      if sb='D' then result:=ddeb else
      if sb='F' then result:=dfin ;
      end ;
   end ;
end ;


{=============================================================================}
Procedure InitMinMax(var C1,C2 : String) ;
BEGIN
If C1='' Then C1:='!' ; If C2='' Then C2:='·················' ;
END ;

{=============================================================================}
Procedure DecodeNatLibre(fb : TFichierBase ; St : String ; Var Nat : TTabNatLibre ; DefautPourDeb : Boolean) ;
Var Lg : Array[1..10] Of Integer ;
    i,j,Deb : Integer ;
BEGIN
Fillchar(Lg,SizeOf(Lg),#0) ; Fillchar(Nat,SizeOf(Nat),#0) ;
Case fb Of fbGene : i:=1 ; fbAux : i:=2 ; fbAxe1..fbAxe5 : i:=3 ; fbBudgen : i:=4 ; fbBudSec1..fbBudSec5 : i:=5 ; Else Exit ; END ;
{$IFDEF NOVH}
For j:=1 To 10 Do Lg[i] := GetLgTableLibre(i, j); Deb:=1 ;
{$ELSE}
For j:=1 To 10 Do Lg[i]:=VH^.LgTableLibre[i,j] ; Deb:=1 ;
{$ENDIF NOVH}
For i:=1 To 10 Do
  BEGIN
  If Lg[i]>0 Then
     BEGIN
     Nat[i]:=Copy(St,Deb,Lg[i]) ;
     If Nat[i][1]='*' Then
        BEGIN
        If DefautPourDeb Then Nat[i]:='!' Else Nat[i]:='·' ;
        END ;
     END ;
  Deb:=Lg[i]+1 ;
  END ;
END ;

{=============================================================================}
Function DansExo(Exo : TExoDate ; Date1,Date2 : TDateTime) : Boolean ;
begin
Result:=(Date1>=Exo.Deb) And (Date2<=Exo.Fin) ;
end ;

(*======================================================================*)
Function DansExoCloture(DD1,DD2 : TDateTime ; Var LExo : TExoDate) : Boolean ;
Var i : Integer ;
BEGIN
Result:=FALSE ;
For i:=1 To 5 Do
  BEGIN
  {$IFDEF NOVH}
  If DansExo(GetExoClo[i],DD1,DD2) then BEGIN LExo:= GetExoClo[i] ; Result:=TRUE ; Exit ; END ;
  {$ELSE}
  If DansExo(VH^.ExoClo[i],DD1,DD2) then BEGIN LExo:=VH^.ExoClo[i] ; Result:=TRUE ; Exit ; END ;
  {$ENDIF NOVH}
  END ;
END ;

Function DansExoViaQuery(Date1,Date2 : TDateTime ; Var MonoExo : Boolean ; Var Exo1,Exo2 : TExoDate) : Boolean ;
Var Q : TQuery ;
    Deb,Fin : TDateTime ;
    PremFois : Boolean ;
    Exo : TExoDate ;
    i,j,k : Word ;
BEGIN
Result:=FALSE ;
Q:=OpenSQL('SELECT * FROM EXERCICE',TRUE) ; PremFois:=TRUE ;
While Not Q.Eof Do
   BEGIN
   Deb:=Q.FindField('EX_DATEDEBUT').AsDateTime ;
   Fin:=Q.FindField('EX_DATEFIN').AsDateTime ;
   If ((date1>=Deb) And (date1<=Fin)) Or ((date2>=Deb) And (date2<=Fin))then
      BEGIN
      If PremFois Then MonoExo:=TRUE Else MonoExo:=FALSE ;
      Exo.Deb:=Deb ; Exo.Fin:=Fin ; Exo.Code:=Q.FindField('EX_EXERCICE').AsString ;
      Exo.DateButoir:=Q.FindField('EX_DATECUM').AsDateTime ;
      Exo.DateButoirRub:=Q.FindField('EX_DATECUMRUB').AsDateTime ;
      Exo.DateButoirBud:=Q.FindField('EX_DATECUMBUD').AsDateTime ;
      Exo.DateButoirBudgete:=Q.FindField('EX_DATECUMBUDGET').AsDateTime ;
      NombrePerExo(Exo,i,j,k) ; Exo.NombrePeriode:=k ;
      If PremFois Then Exo1:=Exo Else Exo2:=Exo ;
      PremFois:=FALSE ;
      Result:=TRUE ;
      END ;
   Q.Next ;
   END ;
Ferme(Q) ;

END ;

(*======================================================================*)
Function QuelExoDate(Date1,Date2 : TDateTime ; Var MonoExo : Boolean ; Var Exo : TExoDate) : Boolean ;
Var LExo,LExo2 : TExoDate ;
BEGIN
MonoExo:=FALSE ; Result:=FALSE ;
{$IFDEF NOVH}
if DansExo(GetPrecedent,Date1,Date2) then BEGIN MonoExo:=TRUE ; Exo:=GetPrecedent ; END else
   if DansExo(GetEnCours,Date1,Date2) then BEGIN MonoExo:=TRUE ; Exo:=GetEnCours ; END else
      if DansExo(GetSuivant,Date1,Date2) then BEGIN MonoExo:=TRUE ; Exo:=GetSuivant ; END Else
         If DansExoCloture(Date1,Date2,LExo) Then BEGIN MonoExo:=TRUE ; Exo:=LExo ; END Else
            If DansExoViaQuery(Date1,Date2,MonoExo,LExo,LExo2) Then Exo:=LExo Else Exit ;
{$ELSE}
if DansExo(VH^.Precedent,Date1,Date2) then BEGIN MonoExo:=TRUE ; Exo:=VH^.Precedent ; END else
   if DansExo(VH^.EnCours,Date1,Date2) then BEGIN MonoExo:=TRUE ; Exo:=VH^.EnCours ; END else
      if DansExo(VH^.Suivant,Date1,Date2) then BEGIN MonoExo:=TRUE ; Exo:=VH^.Suivant ; END Else
         If DansExoCloture(Date1,Date2,LExo) Then BEGIN MonoExo:=TRUE ; Exo:=LExo ; END Else
            If DansExoViaQuery(Date1,Date2,MonoExo,LExo,LExo2) Then Exo:=LExo Else Exit ;
{$ENDIF NOVH}
Result:=TRUE ;
END ;

(*======================================================================*)
Function QuelExoDate2(Date1,Date2 : TDateTime ; Var MonoExo : Boolean ; Var Exo1,Exo2 : TExoDate) : Boolean ;
Var LExo,Lexo2 : TExoDate ;
BEGIN
MonoExo:=FALSE ; Result:=FALSE ;
{$IFDEF NOVH}
if DansExo(GetPrecedent,Date1,Date2) then BEGIN MonoExo:=TRUE ; Exo1:=GetPrecedent ; Exo2:=Exo1 ; END else
   if DansExo(GetEnCours,Date1,Date2) then BEGIN MonoExo:=TRUE ; Exo1:=GetEnCours ; Exo2:=Exo1 ; END else
      if DansExo(GetSuivant,Date1,Date2) then BEGIN MonoExo:=TRUE ; Exo1:=GetSuivant ; Exo2:=Exo1 ; END Else
         If DansExoCloture(Date1,Date2,LExo) Then BEGIN MonoExo:=TRUE ; Exo1:=LExo ; Exo2:=Exo1 ; END Else
            If DansExoViaQuery(Date1,Date2,MonoExo,LExo,LExo2) Then BEGIN Exo1:=LExo ; Exo2:=LExo2 ; END Else Exit ;
{$ELSE}
if DansExo(VH^.Precedent,Date1,Date2) then BEGIN MonoExo:=TRUE ; Exo1:=VH^.Precedent ; Exo2:=Exo1 ; END else
   if DansExo(VH^.EnCours,Date1,Date2) then BEGIN MonoExo:=TRUE ; Exo1:=VH^.EnCours ; Exo2:=Exo1 ; END else
      if DansExo(VH^.Suivant,Date1,Date2) then BEGIN MonoExo:=TRUE ; Exo1:=VH^.Suivant ; Exo2:=Exo1 ; END Else
         If DansExoCloture(Date1,Date2,LExo) Then BEGIN MonoExo:=TRUE ; Exo1:=LExo ; Exo2:=Exo1 ; END Else
            If DansExoViaQuery(Date1,Date2,MonoExo,LExo,LExo2) Then BEGIN Exo1:=LExo ; Exo2:=LExo2 ; END Else Exit ;
{$ENDIF NOVH}
Result:=TRUE ;
END ;

{=============================================================================}
Procedure CumulVersSolde(Var Cum : TabDC) ;
Var Solde : Double ;
BEGIN
Solde:=Cum.TotDebit-Cum.TotCredit ;
if Solde<0 then BEGIN Cum.TotDebit:=0 ; Cum.TotCredit:=Abs(Solde) ; END else
                BEGIN Cum.TotDebit:=Solde ; Cum.TotCredit:=0 ; END ;
END ;

{=============================================================================}
Procedure Sommation (Var TotCpt : TabTot ; Avec6 : Boolean) ;
Var i,IMax : Integer ;
    Tot : TabDC ;
BEGIN
Fillchar(Tot,SizeOf(Tot),#0) ; If Avec6 Then IMax:=6 Else IMax:=5 ;
For i:=1 To IMax Do
  BEGIN
  Tot.TotDebit:=Tot.TotDebit+TotCpt[i].TotDebit ;
  Tot.TotCredit:=Tot.TotCredit+TotCpt[i].TotCredit ;
  END ;
TotCpt[7].TotDebit:=TotCpt[1].TotDebit ; TotCpt[7].TotCredit:=TotCpt[1].TotCredit ;
TotCpt[1].TotDebit:=Tot.TotDebit ; TotCpt[1].TotCredit:=Tot.TotCredit ;
END ;

{=============================================================================}
Function WhereSupp(P : String ; SetTyp : SetttTypePiece) : String ;
BEGIN
Result:='' ;
If tpReel In SetTyp Then Result:=Result+'OR '+P+'QUALIFPIECE="N" ' ;
If tpSim  In SetTyp Then Result:=Result+'OR '+P+'QUALIFPIECE="S" ' ;
If tpPrev In SetTyp Then Result:=Result+'OR '+P+'QUALIFPIECE="P" ' ;
If tpSitu In SetTyp Then Result:=Result+'OR '+P+'QUALIFPIECE="U" ' ;
If tpRevi In SetTyp Then Result:=Result+'OR '+P+'QUALIFPIECE="R" ' ;
If tpCloture In SetTyp Then Result:=Result+'OR '+P+'QUALIFPIECE="C" ' ;
If tpIfrs In SetTyp Then Result:=Result+'OR '+P+'QUALIFPIECE="I" ' ;
If Result<>'' Then
   BEGIN
   Result:=Copy(Result,4,Length(Result)-3) ;
   Result:=' AND ('+Result+') ' ;
   END ;
END ;

function EnDevise(FiltreDev,DeviseEnPivot : Boolean) : Boolean ;
begin
Result:=FALSE ;
If FiltreDev Then
   BEGIN
   If DeviseEnPivot Then Else Result:=TRUE ;
   END Else Result:=FALSE ;
end ;

Function RajouteOrderBy(St : String) : String ;

BEGIN
//i:=Pos(
END ;

{=============================================================================}
Function FabricReqCpt1 ( fb : TFichierBase ; SetTyp : SetttTypePiece ; FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot : Boolean ; Multiple : TTypeCalc ;
                         FiltreSup,ChampSup : String ;
                         Cpte : String ; Date1,Date2 : TDateTime ; Etab,Devise,Exo : String ; VCS : Variant;stBase : string) : TQuery ;
Var Q   : TQuery ;
    StCalc,StD : String ;
    SUM_Montant : String ;
    P : String[2] ;
    SQL : String ;
    JointureGene,JointureAux,JointureAna, JointureAnly : Boolean ;
BEGIN
(*
Q:=TQuery.Create(Application) ;
Q.DatabaseName:=DBSOC.DataBaseName ;
Q.SessionName:=DBSOC.SessionName ;
Q.SQL.Clear ;
*)
//Q:=NIL ;

Jointureana := False;           {FP FQ16854}

Case Fb Of
  FbGene : StCalc:='SELECT E_GENERAL, E_EXERCICE, ' ;
  FbAux  : StCalc:='SELECT E_AUXILIAIRE, E_EXERCICE, ';
  {b FP FQ16854}
  //fbAxe1..fbAxe5 : StCalc:='SELECT Y_SECTION, Y_EXERCICE, ' ;
  fbAxe1..fbAxe5 : StCalc:='SELECT '+TSQLAnaCroise.ChampSection(FbToAxe(fb))+', Y_EXERCICE, ' ;
  {e FP FQ16854}
  fbJal : StCalc:='SELECT E_JOURNAL, E_EXERCICE, ' ;
  END ;
//Q.SQL.Add(StCalc) ;
SQL:='' ;
SQL:=SQL+StCalc ;
Case fb Of
  fbGene,fbAux,fbJal : If EnDevise(FiltreDev,DeviseEnPivot) Then SUM_Montant:='sum(E_DEBITDEV), sum(E_CREDITDEV), 0, 0, '
                                                            Else SUM_Montant:='sum(E_DEBIT), sum(E_CREDIT), 0, 0, ' ;
  fbAxe1..fbAxe5 : If EnDevise(FiltreDev,DeviseEnPivot) Then SUM_Montant:='sum(Y_DEBITDEV), sum(Y_CREDITDEV), 0, 0,  '
                                                                     Else SUM_Montant:='sum(Y_DEBIT), sum(Y_CREDIT), 0, 0, ' ;
  END ;
Case fb Of
  fbGene,fbAux,fbJal : StCalc:=SUM_MONTANT+'E_ECRANOUVEAU, E_QUALIFPIECE FROM ECRITURE ' ;
  fbAxe1..fbAxe5 : StCalc:=SUM_MONTANT+'Y_ECRANOUVEAU, Y_QUALIFPIECE FROM ANALYTIQ ' ;
  END ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
If FiltreSup<>'' Then
  BEGIN
  StCalc:='' ;
  JointureGene:=FALSE ; JointureAux:=FALSE ; JointureAna:=FALSE ; JointureAnly := FALSE; // ajout me  fiche 17455
  If Pos('G_',FiltreSup)>0 Then JointureGene:=TRUE ;
  If Pos('T_',FiltreSup)>0 Then JointureAux:=TRUE ;
  If Pos('S_',FiltreSup)>0 Then JointureAna:=TRUE ;
  If Pos('Y_',FiltreSup)>0 Then JointureAnly:=TRUE ;  // ajout me fiche 17455
  If JointureGene Then StCalc:=' LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL ' ;
  If JointureAux Then StCalc:=StCalc+' LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE ' ;
  {b FP FQ16854}
  //If Jointureana Then StCalc:=StCalc+' LEFT JOIN SECTION ON Y_SECTION=S_SECTION ' ;
  If Jointureana Then StCalc:=StCalc+' LEFT JOIN SECTION ON '+TSQLAnaCroise.ChampSection(FbToAxe(fb))+'=S_SECTION ' ;
  {e FP FQ16854}
  
  // ajout me fiche Com 17455  getcumul2 rubrique + analytiq // Fiche 19387
  If JointureAnly and (not (fb in [FbAxe1..FbAxe5])) Then
  StCalc:=StCalc+' LEFT JOIN ANALYTIQ ON '+TSQLAnaCroise.ChampSection(FbToAxe(fb))+'=Y_SECTION ';

  If StCalc<>'' Then SQL:=SQL+StCalc ;
  END ;
{$IFDEF NOVH}
{$ELSE}
If not EstSpecif('51185') Then
BEGIN
{$ENDIF NOVH}
  Case Fb Of
    FbGene : StCalc:='WHERE E_GENERAL="'+Cpte+'" AND E_DATECOMPTABLE>="'+UsDateTime(Date1)+'" AND E_DATECOMPTABLE<="'+UsDateTime(Date2)+'" ' ;
    FbAux  : StCalc:='WHERE E_AUXILIAIRE="'+Cpte+'" AND E_DATECOMPTABLE>="'+UsDateTime(Date1)+'" AND E_DATECOMPTABLE<="'+UsDateTime(Date2)+'" ';
{b FP FQ16854}
    //fbAxe1..fbAxe5 : StCalc:='WHERE Y_SECTION="'+Cpte+'" AND Y_AXE="'+fbToAxe(fb)+'" AND Y_DATECOMPTABLE>="'+UsDateTime(Date1)+'" AND Y_DATECOMPTABLE<="'+UsDateTime(Date2)+'" ' ;
    fbAxe1..fbAxe5 : StCalc:='WHERE '+TSQLAnaCroise.ChampSection(fbToAxe(fb))+'="'+Cpte+'"'+
                       ' AND '+TSQLAnaCroise.ConditionAxe(fbToAxe(fb))+
                       ' AND Y_DATECOMPTABLE>="'+UsDateTime(Date1)+'" AND Y_DATECOMPTABLE<="'+UsDateTime(Date2)+'" ' ;
{e FP FQ16854}
    FbJal  : StCalc:='WHERE E_JOURNAL="'+Cpte+'" AND E_DATECOMPTABLE>="'+UsDateTime(Date1)+'" AND E_DATECOMPTABLE<="'+UsDateTime(Date2)+'" ';
    END ;
{$IFDEF NOVH}
{$ELSE}
  END
  Else
  BEGIN
  Case Fb Of
    FbGene : StCalc:='WHERE E_GENERAL=:CPTE AND E_DATECOMPTABLE>=:DD1 AND E_DATECOMPTABLE<=:DD2 ' ;
    FbAux  : StCalc:='WHERE E_AUXILIAIRE=:CPTE AND E_DATECOMPTABLE>=:DD1 AND E_DATECOMPTABLE<=:DD2 ';
{b FP FQ16854}
    //fbAxe1..fbAxe5 : StCalc:='WHERE Y_SECTION=:CPTE AND Y_AXE="'+fbToAxe(fb)+'" AND Y_DATECOMPTABLE>=:DD1 AND Y_DATECOMPTABLE<=:DD2 ' ;
    fbAxe1..fbAxe5 : StCalc:='WHERE '+TSQLAnaCroise.ChampSection(fbToAxe(fb))+'=:CPTE'+
                       ' AND '+TSQLAnaCroise.ConditionAxe(fbToAxe(fb))+
                       ' AND Y_DATECOMPTABLE>=:DD1 AND Y_DATECOMPTABLE<=:DD2 ' ;
{e FP FQ16854}
    FbJal  : StCalc:='WHERE E_JOURNAL=:CPTE AND E_DATECOMPTABLE>=:DD1 AND E_DATECOMPTABLE<=:DD2 ';
    END ;
  END ;
{$ENDIF NOVH}

{b FP FQ16854}
if (FiltreSup<>'') and Jointureana then
  StCalc:=StCalc+' AND S_AXE="'+fbToAxe(fb)+'" ';
{e FP FQ16854}

//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
If FiltreDev Then
   BEGIN
   {$IFDEF NOVH}
   {$ELSE}
   If not EstSpecif('51185') Then
     BEGIN
   {$ENDIF NOVH}
     Case Fb Of
       FbGene,FbAux,fbJal : StCalc:='AND E_DEVISE="'+Devise+'" ' ;
       fbAxe1..fbAxe5 : StCalc:='AND Y_DEVISE="'+Devise+'" ' ;
       END ;
    {$IFDEF NOVH}
    {$ELSE}
     END  Else
     BEGIN
     Case Fb Of
       FbGene,FbAux,fbJal : StCalc:='AND E_DEVISE=:DEV ' ;
       fbAxe1..fbAxe5 : StCalc:='AND Y_DEVISE=:DEV ' ;
       END ;
     END ;
    {$ENDIF NOVH}
//   Q.SQL.Add(StCalc) ;
   SQL:=SQL+StCalc ;
   END ;
If FiltreEtab Then
   BEGIN
   {$IFDEF NOVH}
   {$ELSE}
   If not EstSpecif('51185') Then
     BEGIN
   {$ENDIF NOVH}
     Case Fb Of
       FbGene,FbAux,fbJal : StCalc:='AND E_ETABLISSEMENT="'+Etab+'" ' ;
       fbAxe1..fbAxe5 : StCalc:='AND Y_ETABLISSEMENT="'+Etab+'" ' ;
       END ;
     {$IFDEF NOVH}
     {$ELSE}
     END Else
     BEGIN
     Case Fb Of
       FbGene,FbAux,fbJal : StCalc:='AND E_ETABLISSEMENT=:ETAB ' ;
       fbAxe1..fbAxe5 : StCalc:='AND Y_ETABLISSEMENT=:ETAB ' ;
       END ;
     END ;
     {$ENDIF NOVH}
//   Q.SQL.Add(StCalc) ;
   SQL:=SQL+StCalc ;
   END ;
If FiltreExo Then
   BEGIN
   {$IFDEF NOVH}
   {$ELSE}
   If not EstSpecif('51185') Then
     BEGIN
   {$ENDIF NOVH}
     Case Fb Of
       FbGene,FbAux,fbJal : StCalc:='AND E_EXERCICE="'+Exo+'" ' ;
       fbAxe1..fbAxe5 : StCalc:='AND Y_EXERCICE="'+Exo+'" ' ;
       END ;
     {$IFDEF NOVH}
     {$ELSE}
     END Else
     BEGIN
     Case Fb Of
       FbGene,FbAux,fbJal : StCalc:='AND E_EXERCICE=:EXO ' ;
       fbAxe1..fbAxe5 : StCalc:='AND Y_EXERCICE=:EXO ' ;
       END ;
     END ;
     {$ENDIF NOVH}
//   Q.SQL.Add(StCalc) ;
   SQL:=SQL+StCalc ;
   END ;
P:='E_' ; If fb In [fbAxe1..fbAxe5] Then P:='Y_' ;
StCalc:=WhereSupp(P,SetTyp) ; //If StCalc<>'' Then Q.SQL.Add(StCalc) ;
If StCalc<>'' Then SQL:=SQL+StCalc ;
If FiltreSup<>'' Then
  BEGIN
  StCalc:=' AND ('+FiltreSup+') ' ; //Q.SQL.Add(StCalc) ;
  SQL:=SQL+StCalc ;
  END ;
If ChampSup<>'' Then
  BEGIN
  StCalc:='' ;
  {$IFDEF NOVH}
  {$ELSE}
  If not EstSpecif('51185') Then
    BEGIN
  {$ENDIF NOVH}
    If Not VarIsNull(VCS) Then
      BEGIN
       Case VarType(VCS) of
        varEmpty,varNull  : (*QX.ParamByName('CHAMPSUP').AsString:=''*) ;
        varByte,varSmallint,varInteger  : StCalc:=' AND ('+ChampSup+'='+IntToStr(VCS)+') ' ;
        varSingle,varDouble,varCurrency : BEGIN
                                          StD:=Formatfloat('###0.00',VCS) ; StD:=FindEtReplace(StD,',','.',TRUE) ;
                                          StCalc:=' AND ('+ChampSup+'='+StD+') '
                                          END ;
        varDate  : StCalc:=' AND ('+ChampSup+'="'+UsDateTime(VCS)+'") ' ;
        else  StCalc:=' AND ('+ChampSup+'="'+VCS+'") ' ;
        END ;
      END ;
    {$IFDEF NOVH}
    {$ELSE}
    END Else
    BEGIN
    StCalc:=' AND ('+ChampSup+'=:CHAMPSUP) ' ; //Q.SQL.Add(StCalc) ;
    END ;
    {$ENDIF NOVH}
  SQL:=SQL+StCalc ;
  END ;
Case Fb Of
  FbGene : StCalc:='GROUP BY E_GENERAL,E_EXERCICE,E_ECRANOUVEAU,E_QUALIFPIECE ' ;
  FbAux  : StCalc:='GROUP BY E_AUXILIAIRE,E_EXERCICE,E_ECRANOUVEAU,E_QUALIFPIECE ' ;
  {b FP FQ16854}
  //fbAxe1..fbAxe5 : StCalc:='GROUP BY Y_SECTION,Y_EXERCICE,Y_ECRANOUVEAU,Y_QUALIFPIECE ' ;
  fbAxe1..fbAxe5 : StCalc:='GROUP BY '+TSQLAnaCroise.ChampSection(FbToAxe(Fb))+',Y_EXERCICE,Y_ECRANOUVEAU,Y_QUALIFPIECE ' ;
  {e FP FQ16854}
  FbJal : StCalc:='GROUP BY E_JOURNAL,E_EXERCICE,E_ECRANOUVEAU,E_QUALIFPIECE ' ;
  END ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
(*
ChangeSQL(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
*)
{$IFDEF EAGLCLIENT}
Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase) ;
{$ELSE}
{$IFDEF NOVH}
Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase);
{$ELSE}
If not EstSpecif('51185') Then Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase)
                          Else Q:=PrepareSQL(SQL,TRUE) ;
{$ENDIF NOVH}
{$ENDIF}
Result:=Q ;
END ;

{=============================================================================}
Function FabricReqHisto1 ( fb : TFichierBase ; SetTyp : SetttTypePiece ; FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot : Boolean ; Multiple : TTypeCalc;
                           BalSit : string ; Cpte : String ; FiltreSup : String; stBase : string) : TQuery ;
Var Q   : TQuery ;
    StCalc,SQL : String ;
    SUM_Montant : String ;
//    P : String[2] ;
    JointureGene,JointureAux,JointureAna : Boolean ;
BEGIN
{Q:=Nil ;} SQL:='' ;
StCalc:='SELECT BSE_COMPTE1, BSE_COMPTE2, ' ;
SQL:=SQL+StCalc ;
If EnDevise(FiltreDev,DeviseEnPivot) Then SUM_Montant:='BSE_DEBITDEV, BSE_CREDITDEV '
                                   Else SUM_Montant:='BSE_DEBIT,BSE_CREDIT ' ;
StCalc:=SUM_MONTANT+' FROM CBALSITECR ' ;
SQL:=SQL+StCalc ;
If FiltreSup<>'' Then
  BEGIN
  StCalc:='' ;
  JointureGene:=FALSE ; JointureAux:=FALSE ; JointureAna:=FALSE ;
  If Pos('G_',FiltreSup)>0 Then JointureGene:=TRUE ;
  If Pos('T_',FiltreSup)>0 Then JointureAux:=TRUE ;
  If Pos('S_',FiltreSup)>0 Then JointureAna:=TRUE ;
  If JointureGene Then StCalc:=' LEFT JOIN GENERAUX ON BSE_COMPTE1=G_GENERAL ' ;
  If JointureAux Then StCalc:=StCalc+' LEFT JOIN TIERS ON BSE_COMPTE1=T_AUXILIAIRE ' ;
  If Jointureana Then StCalc:=StCalc+' LEFT JOIN SECTION ON BSE_COMPTE1=S_SECTION ' ;
  If StCalc<>'' Then SQL:=SQL+StCalc ;
  END ;
StCalc:=' WHERE BSE_CODEBAL="'+BalSit+'" ' ;
SQL:=SQL+StCalc ;
{$IFDEF NOVH}
{$ELSE}
If not EstSpecif('51185') Then
  BEGIN
{$ENDIF NOVH}
  StCalc:=' AND BSE_COMPTE1 LIKE "'+Cpte+'%" ' ;
{$IFDEF NOVH}
{$ELSE}
  END Else
  BEGIN
  StCalc:=' AND BSE_COMPTE1 LIKE :CPTE ' ;
  END ;
{$ENDIF NOVH}
SQL:=SQL+StCalc ;
If FiltreSup<>'' Then
  BEGIN
  StCalc:=' AND ('+FiltreSup+') ' ; //Q.SQL.Add(StCalc) ;
  SQL:=SQL+StCalc ;
  END ;
{$IFDEF EAGLCLIENT}
Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase) ;
{$ELSE}
{$IFDEF NOVH}
Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase);
{$ELSE}
If not EstSpecif('51185') Then Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase)
              Else Q:=PrepareSQL(SQL,TRUE) ;
{$ENDIF NOVH}
{$ENDIF}
  Result:=Q ;
END ;

{=============================================================================}
Function PrepareTotCpt(fb : TFichierBase ; SetTyp : SetttTypePiece ; FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot : Boolean ; Multiple : TTypeCalc; stBase : string = '') : TQuery ;
BEGIN
Result:=Nil ;
{$IFDEF NOVH}
{$ELSE}
If EstSpecif('51185') Then Result:=FabricReqCpt1(fb,SetTyp,FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot,Multiple,'','',
                                             '',iDate1900,iDate1900,'','','',NULL,stBase) ;
{$ENDIF NOVH}
END ;

{=============================================================================}
procedure AttribParamsCpt1 ( QX : TQuery ; Cpte : String17 ; Date1,Date2 : TDateTime ; Etab,Devise,Exo : String3 ; VCS : Variant) ;
BEGIN
{$IFNDEF EAGLCLIENT}
QX.ParamByName('CPTE').AsString:=Cpte ;
QX.ParamByName('DD1').AsDateTime:=Date1 ; QX.ParamByName('DD2').AsDateTime:=Date2 ;
If Devise<>'' Then QX.ParamByName('DEV').AsString:=Devise ;
If Etab<>'' Then QX.ParamByName('ETAB').AsString:=Etab ;
If Exo<>'' Then QX.ParamByName('EXO').AsString:=Exo ;
If Not VarIsNull(VCS) Then
  BEGIN
   Case VarType(VCS) of
    varEmpty,varNull  : (*QX.ParamByName('CHAMPSUP').AsString:=''*) ;
    varByte,varSmallint,varInteger  : QX.ParamByName('CHAMPSUP').AsInteger:=VCS ;
    varSingle,varDouble,varCurrency : QX.ParamByName('CHAMPSUP').AsFloat:=VCS ;
    varDate  : QX.ParamByName('CHAMPSUP').AsDateTime:=VCS ;
    else  QX.ParamByName('CHAMPSUP').AsString:=VCS ;
    END ;

  END ;
{$ENDIF}
END ;

{=============================================================================}
procedure AttribParamsCpt1Histo ( QX : TQuery ; Date1,Date2 : tDateTime ; Cpte : String17 ; Etab,Devise,Exo : String3) ;
BEGIN
{$IFNDEF EAGLCLIENT}
QX.ParamByName('CPTE').AsString:=Cpte ;
{QX.ParamByName('DD1').AsDateTime:=Date1 ; QX.ParamByName('DD2').AsDateTime:=Date2 ;
If Devise<>'' Then QX.ParamByName('DEV').AsString:=Devise ;
If Etab<>'' Then QX.ParamByName('ETAB').AsString:=Etab ;
If Exo<>'' Then QX.ParamByName('EXO').AsString:=Exo ;}
{$ENDIF}
END ;

{=============================================================================}
procedure positionneTotaux(Qcalc : TQuery ; Var T : TabTot ; Var TEURO : TabTot ; Dec,DecE : Integer) ;
Var TD,TC,TDE,TCE : Double ;
    STyp,AN,Exo  : String3 ;
    i : Byte ;
    Cpt : String ;
BEGIN
While Not QCalc.EOF Do
  BEGIN
  AN:='N' ;
  Cpt:=QCalc.Fields[0].AsString ;
  Exo:=QCalc.Fields[1].AsString ;
  TD:=Arrondi(QCalc.Fields[2].AsFloat,Dec) ;
  TC:=Arrondi(QCalc.Fields[3].AsFloat,Dec) ;
  TDE:=Arrondi(QCalc.Fields[4].AsFloat,DecE) ;
  TCE:=Arrondi(QCalc.Fields[5].AsFloat,DecE) ;
  AN:=QCalc.Fields[6].AsString ;
  STyp:=QCalc.Fields[7].AsString ;
  i:=6 ;
  If AN<>'N' Then i:=0 Else If STyp='N' Then i:=1 Else If STyp='S' Then i:=2 Else
     If STyp='P' Then i:=3 Else If STyp='U' Then i:=4 Else If STyp='R' Then i:=5
     Else If STyp='I' Then i:=1 ; //FQ 13827 Prise en compte IFRS 09/09/2005 SBO
  If AN='C' Then i:=6 ;
  If (TD<>0) Or (TC<>0) Then
     BEGIN
     T[i].TotDebit:=Arrondi(T[i].TotDebit+TD,Dec) ;
     T[i].TotCredit:=Arrondi(T[i].TotCredit+TC,Dec) ;
     END ;
  If (TDE<>0) Or (TCE<>0) Then
     BEGIN
     TEURO[i].TotDebit:=Arrondi(TEURO[i].TotDebit+TDE,DecE) ;
     TEURO[i].TotCredit:=Arrondi(TEURO[i].TotCredit+TCE,DecE) ;
     END ;

  QCalc.Next ;
  END ;
END ;

procedure positionneTotauxHisto(Qcalc : TQuery ; Var T : TabTot ; Var TEURO : TabTot ; Dec,DecE : Integer) ;
Var TD,TC,TDE,TCE : Double ;
    i : Byte ;
    Cpt : String ;
BEGIN
While Not QCalc.EOF Do
  BEGIN
  Cpt:=QCalc.Fields[0].AsString ;
//  Exo:=QCalc.Fields[1].AsString ;
  TD:=Arrondi(QCalc.Fields[2].AsFloat,Dec) ;
  TC:=Arrondi(QCalc.Fields[3].AsFloat,Dec) ;
  // Les balances de situation étant sytématiquement en euro, on renvoir la contrevaleur avec EuroToPivot.
  TDE:=Arrondi(EuroToPivot(QCalc.Fields[2].AsFloat),DecE) ;
  TCE:=Arrondi(EuroToPivot(QCalc.Fields[3].AsFloat),DecE) ;
  i:=1 ;
  If (TD<>0) Or (TC<>0) Then
     BEGIN
     T[i].TotDebit:=Arrondi(T[i].TotDebit+TD,Dec) ;
     T[i].TotCredit:=Arrondi(T[i].TotCredit+TC,Dec) ;
     END ;
  If (TDE<>0) Or (TCE<>0) Then
     BEGIN
     TEURO[i].TotDebit:=Arrondi(TEURO[i].TotDebit+TDE,DecE) ;
     TEURO[i].TotCredit:=Arrondi(TEURO[i].TotCredit+TCE,DecE) ;
     END ;

  QCalc.Next ;
  END ;
END ;

{=============================================================================}
Function IperOk(N1,N2 : Integer ; Suite : Boolean ; Var M1,M2 : Integer) : Boolean ;
BEGIN
Result:=TRUE ;
M1:=N1 ; M2:=N2 ; If N2<=12 Then Exit ;
If Suite Then
   BEGIN
   M1:=1 ; If N1>12 Then M1:=N1-12 ; M2:=N2-12 ;
   Result:=(N2>12) ;
   END Else
   BEGIN
   M2:=12 ; If N2<12 Then M2:=N2 ; M1:=N1 ;
   Result:=(N1<12) ;
   END ;
END ;

{=============================================================================}
procedure positionneTotauxCum(QcalcCum : TQuery ; DateCum1,DateCum2 : TDateTime ; Var T,TEuro : TabTot ;
                              N1,N2 : Integer ; Dec,DecE : Integer) ;
Var TD,TC : Double ;
    STyp,AN,Exo  : String3 ;
    i,j : Byte ;
    Cpt : String ;
    Suite : Boolean ;
    M1,M2 : Integer ;
BEGIN
If N1<=0 Then Exit ;
If N2<N1 Then Exit ;
//If N2>12 Then Exit ;
While Not QCalcCum.EOF Do
  BEGIN
  AN:='N' ; TD:=0 ; TC:=0 ;
  Cpt:=QCalcCum.Fields[1].AsString ;
  Exo:=QCalcCum.Fields[3].AsString ;
  Suite:=QCalcCum.Fields[4].AsString='X' ;
  If IPerOk(N1,N2,Suite,M1,M2) Then
     BEGIN
     For j:=M1 To M2 Do
       BEGIN
       TD:=TD+Arrondi(QCalcCum.FindField('CU_DEBIT'+IntToStr(j)).AsFloat,V_PGI.OkDecV) ;
       TC:=TC+Arrondi(QCalcCum.FindField('CU_CREDIT'+IntToStr(j)).AsFloat,V_PGI.OkDecV) ;
       END ;
     (*
     AN:=QCalc.Fields[4].AsString ;
     *)
     STyp:=QCalcCum.Fields[7].AsString ;
     i:=6 ;
     If STyp='N' Then i:=1 Else If STyp='S' Then i:=2 Else
        If STyp='P' Then i:=3 Else If STyp='U' Then i:=4 Else If STyp='R' Then i:=5
        Else If STyp = 'I' then i := 1 ; //FQ 13827 Prise en compte IFRS 09/09/2005 SBO

     If (TD<>0) Or (TC<>0) Then
        BEGIN
        T[i].TotDebit:=Arrondi(T[i].TotDebit+TD,V_PGI.OkDecV) ;
        T[i].TotCredit:=Arrondi(T[i].TotCredit+TC,V_PGI.OkDecV) ;
        END ;
     If (M1=1) And (i=1) And (Not Suite) Then
        BEGIN
        T[0].TotDebit:=T[0].TotDebit+Arrondi(QCalcCum.FindField('CU_DEBITAN').AsFloat,V_PGI.OkDecV) ;
        T[0].TotCredit:=T[0].TotCredit+Arrondi(QCalcCum.FindField('CU_CREDITAN').AsFloat,V_PGI.OkDecV) ;
        END ;
     END ;

  QCalcCum.Next ;
  END ;
END ;

{=============================================================================}
procedure CalculHisto(Q : TQuery ; Date1,Date2 : TDateTime ; Cpt : String17 ; Devise,Etab,Exo : String3 ;
                      Var TotCpt,TotCptEuro : TabTot ; Dec,DecE : Integer) ;

BEGIN
AttribParamsCpt1Histo(Q,Date1,Date2,Cpt,Etab,Devise,Exo) ; Q.Open ; PositionneTotauxHisto(Q,TotCpt,TotCptEuro,Dec,DecE) ; Q.Close ;
END ;

{=============================================================================}
procedure ExecuteTotCpt(Var Q : TQuery ; Cpt : String17 ; Date1,Date2 : TDateTime ; Devise,Etab,Exo : String3 ;
                        Var TotCpt,TotCptEuro : TabTot ; SommeTout : Boolean ; Dec,DecE : Integer ;
                        fb : tfichierbase ; SetType : SetttTypePiece ; Multiple : TTypeCalc ; DeviseEnPivot : Boolean; stBase : string = '') ;
BEGIN
FillChar(TotCpt,SizeOf(TotCpt),#0) ;
FillChar(TotCptEuro,SizeOf(TotCptEuro),#0) ;
{$IFDEF NOVH}
{$ELSE}
If  EstSpecif('51185') Then If Q=NIL Then Exit ;
If not EstSpecif('51185') Then
  BEGIN
{$ENDIF NOVH}
  Q:=FabricReqCpt1(fb,SetType,Devise<>'',Etab<>'',Exo<>'',DeviseEnPivot,Multiple,'','',Cpt,Date1,Date2,Etab,Devise,Exo,NULL, stBase) ;
  PositionneTotaux(Q,TotCpt,TotCptEuro,Dec,DecE) ;
  Ferme(Q) ; Q:=Nil ;
{$IFDEF NOVH}
{$ELSE}
  END Else
  BEGIN
  AttribParamsCpt1(Q,Cpt,Date1,Date2,Etab,Devise,Exo,VarNull) ;
  Q.Open ;
  PositionneTotaux(Q,TotCpt,TotCptEuro,Dec,DecE) ;
  Q.Close ;
  END ;
{$ENDIF NOVH}
If SommeTout Then
   BEGIN
   Sommation(TotCpt,FALSE) ; Sommation(TotCptEuro,FALSE) ;
   END ;
END ;

{=============================================================================}
Function FabricReqCpt2 ( fb,fb2 : TFichierBase ; SetTyp : SetttTypePiece ; FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot : Boolean ; Multiple : TTypeCalc ;
                         FiltreSup,ChampSup : String ;
                         Cpte1,Cpte2 : String ; Date1,Date2 : TDateTime ; Etab,Devise,Exo : String ; VCS : Variant; stBase : string = '') : TQuery ;
Var Q   : TQuery ;
    StCalc,SQL : String ;
    SUM_Montant : String ;
    P : String[2] ;
    DEV : String[3] ;
    StD  :String ;
    JointureGene,JointureAux,JointureAna : Boolean ;
BEGIN
//Q:=Nil ;
SQL:='' ;
P:='E_' ;
If (Fb in [fbAxe1..fbAxe5]) Or (Fb2 in [fbAxe1..fbAxe5]) Then P:='Y_' ;
(*
Q:=TQuery.Create(Application) ;
Q.DatabaseName:=DBSOC.DataBaseName ;
Q.SessionName:=DBSOC.SessionName ;
Q.SQL.Clear ;
*)
StCalc:='SELECT ' ;
Case Fb Of
  FbGene : StCalc:=Stcalc+P+'GENERAL, ' ;
  FbAux  : StCalc:=Stcalc+P+'AUXILIAIRE, ';
  {b FP FQ16854}
  //fbAxe1..fbAxe5 : StCalc:=Stcalc+P+'SECTION, ' ;
  fbAxe1..fbAxe5 : StCalc:=Stcalc+TSQLAnaCroise.ChampSection(FbToAxe(Fb))+', ' ;
  {e FP FQ16854}
  fbJal : StCalc:=Stcalc+P+'JOURNAL, ' ;
  END ;
Case Fb2 Of
  FbGene : StCalc:=Stcalc+P+'GENERAL, ' ;
  FbAux  : StCalc:=Stcalc+P+'AUXILIAIRE, ';
  {b FP FQ16854}
  //fbAxe1..fbAxe5 : StCalc:=Stcalc+P+'SECTION, ' ;
  fbAxe1..fbAxe5 : StCalc:=Stcalc+TSQLAnaCroise.ChampSection(FbToAxe(Fb2))+', ' ;
  {e fP FQ16854}
  fbJal : StCalc:=Stcalc+P+'JOURNAL, ' ;
  END ;
//If Multiple<>DeuxSansCumul Then StCalc:=StCalc+P+'EXERCICE, ' ;
StCalc:=StCalc+P+'EXERCICE, ' ;
//Q.SQL.Add(StCalc) ;*
SQL:=SQL+StCalc ;
If EnDevise(FiltreDev,DeviseEnPivot) Then DEV:='DEV' Else Dev:='' ;
SUM_Montant:='sum('+P+'DEBIT'+DEV+'), sum('+P+'CREDIT'+DEV+'), ' ;
SUM_Montant:=SUM_Montant+' 0, 0 ' ;
(*
Case Multiple Of
  DeuxSansCumul : StCalc:=SUM_MONTANT+' FROM ' ;
  Else StCalc:=SUM_MONTANT+', '+P+'ECRANOUVEAU, '+P+'QUALIFPIECE FROM ' ;
  END ;
*)
JointureAna:=False;                {FP FQ16854}
StCalc:=SUM_MONTANT+', '+P+'ECRANOUVEAU, '+P+'QUALIFPIECE FROM ' ;
If P='E_' Then STCalc:=StCalc+'ECRITURE ' Else STCalc:=StCalc+'ANALYTIQ ' ;
If FiltreSup<>'' Then
  BEGIN
  JointureGene:=FALSE ; JointureAux:=FALSE ; JointureAna:=FALSE ;
  If Pos('G_',FiltreSup)>0 Then JointureGene:=TRUE ;
  If Pos('T_',FiltreSup)>0 Then JointureAux:=TRUE ;
  If Pos('S_',FiltreSup)>0 Then JointureAna:=TRUE ;
  If JointureGene Then StCalc:=StCalc+' LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL ' ;
  If JointureAux Then StCalc:=StCalc+' LEFT JOIN GENERAUX ON E_AUXILIAIRE=T_AUXILIAIRE ' ;
  {b FP FQ16854}
  //If JointureAna Then StCalc:=StCalc+' LEFT JOIN SECTION ON Y_SECTION=S_SECTION ' ;
  If JointureAna Then
    begin
    if (Fb in [fbAxe1..fbAxe5]) Then
      StCalc:=StCalc+' LEFT JOIN SECTION ON '+TSQLAnaCroise.ChampSection(FbToAxe(Fb))+'=S_SECTION '
    else if (Fb2 in [fbAxe1..fbAxe5]) Then
      StCalc:=StCalc+' LEFT JOIN SECTION ON '+TSQLAnaCroise.ChampSection(FbToAxe(Fb2))+'=S_SECTION '
    end;
  {e FP FQ16854}
  END ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
{$IFDEF NOVH}
{$ELSE}
If not EstSpecif('51185') Then
  BEGIN
{$ENDIF NOVH}
  Case Fb Of
    FbGene : StCalc:='WHERE '+P+'GENERAL="'+Cpte1+'" AND ' ;
    FbAux  : StCalc:='WHERE '+P+'AUXILIAIRE="'+Cpte1+'" AND ' ;
    {b FP FQ16854}
    //fbAxe1..fbAxe5 : StCalc:='WHERE '+P+'SECTION="'+Cpte1+'" AND ' ;
    fbAxe1..fbAxe5 : StCalc:='WHERE '+TSQLAnaCroise.ChampSection(FbToAxe(Fb))+'="'+Cpte1+'" AND ' ;
    {e FP FQ16854}
    FbJal : StCalc:='WHERE '+P+'JOURNAL="'+Cpte1+'" AND ' ;
    END ;
  Case Fb2 Of
    FbGene : StCalc:=StCalc+P+'GENERAL="'+Cpte2+'" AND ' ;
    FbAux  : StCalc:=StCalc+P+'AUXILIAIRE="'+Cpte2+'" AND ' ;
    {b FP FQ16854}
    //fbAxe1..fbAxe5 : StCalc:=StCalc+P+'SECTION="'+Cpte2+'" AND ' ;
    fbAxe1..fbAxe5 : StCalc:=StCalc+TSQLAnaCroise.ChampSection(FbToAxe(Fb2))+'="'+Cpte2+'" AND ' ;
    {e FP FQ16854}
    FbJal : StCalc:='WHERE '+P+'JOURNAL="'+Cpte2+'" AND ' ;
    END ;
{$IFDEF NOVH}
{$ELSE}
  END Else
  BEGIN
  Case Fb Of
    FbGene : StCalc:='WHERE '+P+'GENERAL=:CPTE1 AND ' ;
    FbAux  : StCalc:='WHERE '+P+'AUXILIAIRE=:CPTE1 AND ' ;
    {b FP FQ16854}
    //fbAxe1..fbAxe5 : StCalc:='WHERE '+P+'SECTION=:CPTE1 AND ' ;
    fbAxe1..fbAxe5 : StCalc:='WHERE '+TSQLAnaCroise.ChampSection(FbToAxe(Fb))+'=:CPTE1 AND ' ;
    {e FP FQ16854}
    FbJal : StCalc:='WHERE '+P+'JOURNAL=:CPTE1 AND ' ;
    END ;
  Case Fb2 Of
    FbGene : StCalc:=StCalc+P+'GENERAL=:CPTE2 AND ' ;
    FbAux  : StCalc:=StCalc+P+'AUXILIAIRE=:CPTE2 AND ' ;
    {b FP FQ16854}
    //fbAxe1..fbAxe5 : StCalc:=StCalc+P+'SECTION=:CPTE2 AND ' ;
    fbAxe1..fbAxe5 : StCalc:=StCalc+TSQLAnaCroise.ChampSection(FbToAxe(Fb2))+'=:CPTE2 AND ' ;
    {e FP FQ16854}
    FbJal : StCalc:='WHERE '+P+'JOURNAL=:CPTE2 AND ' ;
    END ;
  END ;
{$ENDIF NOVH}
{b FP FQ16854}
//If (Fb in [fbAxe1..fbAxe5]) THEN StCalc:=STCalc+'Y_AXE="'+fbToAxe(fb)+'" AND ' Else
//  If (Fb2 in [fbAxe1..fbAxe5]) THEN StCalc:=STCalc+'Y_AXE="'+fbToAxe(fb2)+'" AND ' ;
If (Fb in [fbAxe1..fbAxe5]) THEN
  begin
  StCalc:=STCalc+TSQLAnaCroise.ConditionAxe(fbToAxe(fb))+' AND ';
  if JointureAna then
    StCalc:=STCalc+' AND S_AXE="'+fbToAxe(fb)+'" AND ';
  end
Else If (Fb2 in [fbAxe1..fbAxe5]) THEN
  begin
  StCalc:=STCalc+TSQLAnaCroise.ConditionAxe(fbToAxe(fb2))+' AND ' ;
  if JointureAna then
    StCalc:=STCalc+' AND S_AXE="'+fbToAxe(fb2)+'" AND ';
  end;
{e FP FQ16854}
{$IFDEF NOVH}
{$ELSE}
If not EstSpecif('51185') Then
  BEGIN
{$ENDIF NOVH}
  If FiltreExo Then StCalc:=StCalc+P+'EXERCICE="'+Exo+'" AND ' ;
  StCalc:=StCalc+P+'DATECOMPTABLE>="'+UsDateTime(Date1)+'" AND '+P+'DATECOMPTABLE<="'+UsDateTime(Date2)+'" ' ;
{$IFDEF NOVH}
{$ELSE}
  END Else
  BEGIN
  If FiltreExo Then StCalc:=StCalc+P+'EXERCICE=:EXO AND ' ;
  StCalc:=StCalc+P+'DATECOMPTABLE>=:DD1 AND '+P+'DATECOMPTABLE<=:DD2 ' ;
  END ;
{$ENDIF NOVH}
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
If FiltreDev Then
   BEGIN
   {$IFDEF NOVH}
   StCalc:='AND '+P+'DEVISE="'+Devise+'" ';
   {$ELSE}
   If not EstSpecif('51185') Then StCalc:='AND '+P+'DEVISE="'+Devise+'" '
                 Else StCalc:='AND '+P+'DEVISE=:DEV ' ;
   {$ENDIF NOVH}
//   Q.SQL.Add(StCalc) ;
   SQL:=SQL+StCalc ;
   END ;
If FiltreEtab Then
   BEGIN
   {$IFDEF NOVH}
   StCalc:='AND '+P+'ETABLISSEMENT="'+Etab+'" ';
   {$ELSE}
   If not EstSpecif('51185') Then StCalc:='AND '+P+'ETABLISSEMENT="'+Etab+'" '
                 Else StCalc:='AND '+P+'ETABLISSEMENT=:ETAB ' ;
   {$ENDIF NOVH}
//   Q.SQL.Add(StCalc) ;
   SQL:=SQL+StCalc ;
   END ;
StCalc:=WhereSupp(P,SetTyp) ; //If StCalc<>'' Then Q.SQL.Add(StCalc) ;
If StCalc<>'' Then SQL:=SQL+StCalc ;
If FiltreSup<>'' Then
  BEGIN
  StCalc:=' AND ('+FiltreSup+') ' ; //Q.SQL.Add(StCalc) ;
  SQL:=SQL+StCalc ;
  END ;
If ChampSup<>'' Then
  BEGIN
  StCalc:='' ;
  {$IFDEF NOVH}
  {$ELSE}
  If not EstSpecif('51185') Then
    BEGIN
  {$ENDIF NOVH}
    If Not VarIsNull(VCS) Then
      BEGIN
       Case VarType(VCS) of
        varEmpty,varNull  : (*QX.ParamByName('CHAMPSUP').AsString:=''*) ;
        varByte,varSmallint,varInteger  : StCalc:=' AND ('+ChampSup+'='+IntToStr(VCS)+') ' ;
        varSingle,varDouble,varCurrency : BEGIN
                                          StD:=Formatfloat('###0.00',VCS) ; StD:=FindEtReplace(StD,',','.',TRUE) ;
                                          StCalc:=' AND ('+ChampSup+'='+StD+') '
                                          END ;
        varDate  : StCalc:=' AND ('+ChampSup+'="'+UsDateTime(VCS)+'") ' ;
        else  StCalc:=' AND ('+ChampSup+'="'+VCS+'") ' ;
        END ;
      END ;
    {$IFDEF NOVH}
    {$ELSE}
    END Else
    BEGIN
    StCalc:=' AND ('+ChampSup+'=:CHAMPSUP) ' ; //Q.SQL.Add(StCalc) ;
    END ;
    {$ENDIF NOVH}
  SQL:=SQL+StCalc ;
  END ;
StCalc:='GROUP BY ' ;
Case Fb Of
  FbGene : StCalc:=StCalc+P+'GENERAL,' ;
  FbAux  : StCalc:=StCalc+P+'AUXILIAIRE,' ;
  {b FP FQ16854}
  //fbAxe1..fbAxe5 : StCalc:=StCalc+P+'SECTION,' ;
  fbAxe1..fbAxe5 : StCalc:=StCalc+TSQLAnaCroise.ChampSection(FbToAxe(Fb))+',' ;
  {e FP FQ16854}
  fbJal : StCalc:=Stcalc+P+'JOURNAL, ' ;
  END ;
Case Fb2 Of
  FbGene : StCalc:=StCalc+P+'GENERAL,' ;
  FbAux  : StCalc:=StCalc+P+'AUXILIAIRE,' ;
  {b FP FQ16854}
  //fbAxe1..fbAxe5 : StCalc:=StCalc+P+'SECTION,' ;
  fbAxe1..fbAxe5 : StCalc:=StCalc+TSQLAnaCroise.ChampSection(FbToAxe(Fb2))+',' ;
  {e FP FQ16854}
  fbJal : StCalc:=Stcalc+P+'JOURNAL, ' ;
  END ;
{b FP FQ16854}
//If (Fb in [fbAxe1..fbAxe5]) Or (Fb2 in [fbAxe1..fbAxe5]) Then StCalc:=StCalc+'Y_AXE,' ;
  {$IFDEF EAGLSERVER}
  If (not GetParamSocSecur('SO_CROISAXE', False)) and ((Fb in [fbAxe1..fbAxe5]) Or (Fb2 in [fbAxe1..fbAxe5])) Then
    StCalc:=StCalc+'Y_AXE,'
  else if GetParamSocSecur('SO_CROISAXE', False) and (Fb in [fbAxe1..fbAxe5]) then
    StCalc:=StCalc+'Y_AXE="'+FbToAxe(fb)+'",'
  else if GetParamSocSecur('SO_CROISAXE', False) and (Fb2 in [fbAxe1..fbAxe5]) Then
    StCalc:=StCalc+'Y_AXE="'+FbToAxe(fb2)+'",' ;
{$ELSE}
  If (not VH^.AnaCroisaxe) and ((Fb in [fbAxe1..fbAxe5]) Or (Fb2 in [fbAxe1..fbAxe5])) Then
    StCalc:=StCalc+'Y_AXE,'
  else if VH^.AnaCroisaxe and (Fb in [fbAxe1..fbAxe5]) then
    StCalc:=StCalc+'Y_AXE="'+FbToAxe(fb)+'",'
  else if VH^.AnaCroisaxe and (Fb2 in [fbAxe1..fbAxe5]) Then
    StCalc:=StCalc+'Y_AXE="'+FbToAxe(fb2)+'",' ;
  {$ENDIF EAGLSERVER}
{e FP FQ16854}

StCalc:=StCalc+P+'EXERCICE,'+P+'ECRANOUVEAU,'+P+'QUALIFPIECE' ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
(*
ChangeSQL(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
Result:=Q ;
*)
{$IFDEF EAGLCLIENT}
Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase) ;
{$ELSE}
{$IFDEF NOVH}
Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase);
{$ELSE}
If not EstSpecif('51185') Then Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase)
              Else Q:=PrepareSQL(SQL,TRUE) ;
{$ENDIF NOVH}
{$ENDIF}
Result:=Q ;
END ;

{=============================================================================}
procedure AttribParamsCpt2 ( QX : TQuery ; Cpte1,Cpte2 : String17 ; Date1,Date2 : TDateTime ; Etab,Devise,Exo : String3 ; VCS : Variant) ;
BEGIN
{$IFNDEF EAGLCLIENT}
QX.ParamByName('CPTE1').AsString:=Cpte1 ; QX.ParamByName('CPTE2').AsString:=Cpte2 ;
QX.ParamByName('DD1').AsDateTime:=Date1 ; QX.ParamByName('DD2').AsDateTime:=Date2 ;
If Devise<>'' Then QX.ParamByName('DEV').AsString:=Devise ;
If Etab<>'' Then QX.ParamByName('ETAB').AsString:=Etab ;
If Exo<>'' Then QX.ParamByName('EXO').AsString:=Exo ;
If Not VarIsNull(VCS) Then
  BEGIN
   Case VarType(VCS) of
    varEmpty,varNull  : (*QX.ParamByName('CHAMPSUP').AsString:=''*) ;
    varByte,varSmallint,varInteger  : QX.ParamByName('CHAMPSUP').AsInteger:=VCS ;
    varSingle,varDouble,varCurrency : QX.ParamByName('CHAMPSUP').AsFloat:=VCS ;
    varDate  : QX.ParamByName('CHAMPSUP').AsDateTime:=VCS ;
    else  QX.ParamByName('CHAMPSUP').AsString:=VCS ;
    END ;

  END ;
{$ENDIF}
END ;

{=============================================================================}
procedure AttribParamsCpt2Cum ( QX : TQuery ; Cpte1,Cpte2,Axe1,Axe2 : String17 ; Etab,Devise,Exo : String3) ;
BEGIN
{$IFNDEF EAGLCLIENT}
QX.ParamByName('CPTE1').AsString:=Axe1+Cpte1 ; QX.ParamByName('CPTE2').AsString:=Axe2+Cpte2 ;
If Devise='' Then Devise:=V_PGI.DevisePivot ;
If Devise<>'' Then QX.ParamByName('DEV').AsString:=Devise ;
If Etab<>'' Then QX.ParamByName('ETAB').AsString:=Etab ;
If Exo<>'' Then QX.ParamByName('EXO').AsString:=Exo ;
{$ENDIF}
END ;

{=============================================================================}
procedure positionneTotaux2(Qcalc : TQuery ; Var T,TEuro : TabTot ; Dec,DecE : Integer) ;
Var TD,TC,TDE,TCE : Double ;
    STyp,AN,Exo  : String3 ;
    i : Byte ;
    Cpt1,Cpt2 : String ;
BEGIN
While Not QCalc.EOF Do
  BEGIN
  AN:='N' ;
  Cpt1:=QCalc.Fields[0].AsString ;
  Cpt2:=QCalc.Fields[1].AsString ;
  Exo:=QCalc.Fields[2].AsString ;
  TD:=Arrondi(QCalc.Fields[3].AsFloat,Dec) ;
  TC:=Arrondi(QCalc.Fields[4].AsFloat,Dec) ;
  TDE:=Arrondi(QCalc.Fields[5].AsFloat,DecE) ;
  TCE:=Arrondi(QCalc.Fields[6].AsFloat,DecE) ;
  AN:=QCalc.Fields[7].AsString ;
  STyp:=QCalc.Fields[8].AsString ;
  i:=6 ;
  If AN<>'N' Then i:=0 Else If STyp='N' Then i:=1 Else If STyp='S' Then i:=2 Else
     If STyp='P' Then i:=3 Else If STyp='U' Then i:=4 Else If STyp='R' Then i:=5
     Else If STyp='I' Then i:=1 ; //FQ 13827 Prise en compte IFRS 09/09/2005 SBO
  If (TD<>0) Or (TC<>0) Then
     BEGIN
     T[i].TotDebit:=Arrondi(T[i].TotDebit+TD,Dec) ;
     T[i].TotCredit:=Arrondi(T[i].TotCredit+TC,Dec) ;
     END ;
  If (TDE<>0) Or (TCE<>0) Then
     BEGIN
     TEuro[i].TotDebit:=Arrondi(TEuro[i].TotDebit+TDE,DecE) ;
     TEuro[i].TotCredit:=Arrondi(TEuro[i].TotCredit+TCE,DecE) ;
     END ;
  QCalc.Next ;
  END ;
END ;

{=============================================================================}
procedure positionneTotaux2Cum(QcalcCum : TQuery ; DateCum1,DateCum2 : TDateTime ; Var T,TEuro : TabTot ;
                               N1,N2,Dec,DecE : Integer) ;
Var TD,TC : Double ;
    STyp,AN,Exo  : String3 ;
    i,j : Byte ;
    Cpt1,Cpt2 : String ;
    Suite : Boolean ;
    M1,M2 : Integer ;
BEGIN
If N1<=0 Then Exit ;
If N2<N1 Then Exit ;
//If N2>12 Then Exit ;
While Not QCalcCum.EOF Do
  BEGIN
  AN:='N' ; TD:=0 ; TC:=0 ;
  Cpt1:=QCalcCum.Fields[1].AsString ; Cpt2:=QCalcCum.Fields[1].AsString ;
  Exo:=QCalcCum.Fields[3].AsString ;
  Suite:=QCalcCum.Fields[4].AsString='X' ;
  If IPerOk(N1,N2,Suite,M1,M2) Then
     BEGIN
     For j:=M1 To M2 Do
       BEGIN
       TD:=TD+Arrondi(QCalcCum.FindField('CU_DEBIT'+IntToStr(j)).AsFloat,V_PGI.OkDecV) ;
       TC:=TC+Arrondi(QCalcCum.FindField('CU_CREDIT'+IntToStr(j)).AsFloat,V_PGI.OkDecV) ;
       END ;
     (*
     AN:=QCalc.Fields[4].AsString ;
     *)
     STyp:=QCalcCum.Fields[7].AsString ;
     i:=6 ;
     If STyp='N' Then i:=1 Else If STyp='S' Then i:=2 Else
        If STyp='P' Then i:=3 Else If STyp='U' Then i:=4 Else If STyp='R' Then i:=5
        Else If STyp='I' Then i:=1 ; //FQ 13827 Prise en compte IFRS 09/09/2005 SBO

     If (TD<>0) Or (TC<>0) Then
        BEGIN
        T[i].TotDebit:=Arrondi(T[i].TotDebit+TD,V_PGI.OkDecV) ;
        T[i].TotCredit:=Arrondi(T[i].TotCredit+TC,V_PGI.OkDecV) ;
        END ;
     If (M1=1) And (i=1) And (Not Suite) Then
        BEGIN
        T[0].TotDebit:=T[0].TotDebit+Arrondi(QCalcCum.FindField('CU_DEBITAN').AsFloat,V_PGI.OkDecV) ;
        T[0].TotCredit:=T[0].TotCredit+Arrondi(QCalcCum.FindField('CU_CREDITAN').AsFloat,V_PGI.OkDecV) ;
        END ;
     END ;

  QCalcCum.Next ;
  END ;
END ;

{=============================================================================}
Function FabricReqCpt3 ( SetTyp : SetttTypePiece ; FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot : Boolean ;
                         Cpte : String ; Date1,Date2 : TDateTime ; Etab,Devise,Exo : String; stBase: string) : TQuery ;
Var Q   : TQuery ;
    StCalc,SQL : String ;
    SUM_Montant : String ;
    DEV : String3 ;
BEGIN
{Q:=NIL ;} SQL:='' ;
(*
Q:=TQuery.Create(Application) ;
Q.DatabaseName:=DBSOC.DataBaseName ;
Q.SessionName:=DBSOC.SessionName ;
Q.SQL.Clear ;
*)
StCalc:='SELECT E_GENERAL, E_AUXILIAIRE, E_EXERCICE, ' ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
If EnDevise(FiltreDev,DeviseEnPivot) Then DEV:='DEV' Else Dev:='' ;
SUM_Montant:='sum(E_DEBIT'+DEV+'), sum(E_CREDIT'+DEV+'), 0, 0, ' ;
StCalc:=SUM_MONTANT+'E_ECRANOUVEAU, E_QUALIFPIECE FROM ECRITURE ' ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
{$IFDEF NOVH}
{$ELSE}
If not EstSpecif('51185') Then
  BEGIN
{$ENDIF NOVH}
  StCalc:='WHERE E_GENERAL="'+Cpte+'" AND E_AUXILIAIRE<>"." AND ' ;
  StCalc:=STCalc+'E_DATECOMPTABLE>="'+UsDateTime(Date1)+'" AND E_DATECOMPTABLE<="'+UsDateTime(Date2)+'" ' ;
{$IFDEF NOVH}
{$ELSE}
  END Else
  BEGIN
  StCalc:='WHERE E_GENERAL=:CPTE AND E_AUXILIAIRE<>"." AND ' ;
  StCalc:=STCalc+'E_DATECOMPTABLE>=:DD1 AND E_DATECOMPTABLE<=:DD2 ' ;
  END ;
{$ENDIF NOVH}
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
{$IFDEF NOVH}
{$ELSE}
If not EstSpecif('51185') Then
  BEGIN
{$ENDIF NOVH}
  If FiltreDev Then
     BEGIN
     StCalc:='AND E_DEVISE="'+Devise+'" ' ;
     //Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
  If FiltreEtab Then
     BEGIN
     StCalc:='AND E_ETABLISSEMENT="'+Etab+'" ' ;
     //Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
  If FiltreExo Then
     BEGIN
     StCalc:='AND E_EXERCICE="'+Exo+'" ' ;
     //Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
{$IFDEF NOVH}
{$ELSE}
  END  Else
  BEGIN
  If FiltreDev Then
     BEGIN
     StCalc:='AND E_DEVISE=:DEV ' ;
     //Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
  If FiltreEtab Then
     BEGIN
     StCalc:='AND E_ETABLISSEMENT=:ETAB ' ;
     //Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
  If FiltreExo Then
     BEGIN
     StCalc:='AND E_EXERCICE=:EXO ' ;
     //Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
  END ;
{$ENDIF NOVH}
StCalc:=WhereSupp('E_',SetTyp) ;
//If StCalc<>'' Then Q.SQL.Add(StCalc) ;
If StCalc<>'' Then SQL:=SQL+StCalc ;
StCalc:='GROUP BY E_GENERAL,E_AUXILIAIRE,E_EXERCICE,E_ECRANOUVEAU,E_QUALIFPIECE ' ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
StCalc:='ORDER BY E_GENERAL,E_AUXILIAIRE,E_EXERCICE,E_ECRANOUVEAU,E_QUALIFPIECE' ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
(*
ChangeSQL(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
Result:=Q ;
*)
{$IFDEF EAGLCLIENT}
Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase) ;
{$ELSE}
{$IFDEF NOVH}
Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase);
{$ELSE}
If not EstSpecif('51185') Then Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase)
                          Else Q:=PrepareSQL(SQL,TRUE) ;
{$ENDIF NOVH}
{$ENDIF}
Result:=Q ;
END ;

{=============================================================================}
Function PrepareTotCptSolde(SetTyp : SetttTypePiece ; FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot : Boolean; stBase : string = '') : TQuery ;
BEGIN
Result:=NIL ;
{$IFDEF NOVH}
{$ELSE}
If  EstSpecif('51185') Then Result:=FabricReqCpt3(SetTyp,FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot,'',iDate1900,iDate1900,'','','',stBase) ;
{$ENDIF NOVH}
END ;

{=============================================================================}
procedure AttribParamsCpt3 ( QX : TQuery ; Cpte : String17 ; Date1,Date2 : TDateTime ; Etab,Devise,Exo : String3) ;
BEGIN
{$IFNDEF EAGLCLIENT}
QX.ParamByName('CPTE').AsString:=Cpte ;
QX.ParamByName('DD1').AsDateTime:=Date1 ; QX.ParamByName('DD2').AsDateTime:=Date2 ;
If Devise<>'' Then QX.ParamByName('DEV').AsString:=Devise ;
If Etab<>'' Then QX.ParamByName('ETAB').AsString:=Etab ;
If Exo<>'' Then QX.ParamByName('EXO').AsString:=Exo ;
{$ENDIF}
END ;

{=============================================================================}
procedure CalcTotSolde ( Var TD1,TC1 : Double ; Var TMvt,TSld : TabTot ; Dec : Integer) ;
Var i : Byte ;
    SD,SC : Double ;
begin
SD:=0 ; SC:=0 ;
If TD1>TC1 Then SD:=Arrondi(TD1-TC1,Dec) Else
   If TD1<TC1 Then SC:=Arrondi(TC1-TD1,Dec) ;
i:=1 ;
If (TD1<>0) Or (TC1<>0) Then
   BEGIN
   TMvt[i].TotDebit:=Arrondi(TMvt[i].TotDebit+TD1,Dec) ;
   TMvt[i].TotCredit:=Arrondi(TMvt[i].TotCredit+TC1,Dec) ;
   TSld[i].TotDebit:=Arrondi(TSld[i].TotDebit+SD,Dec) ;
   TSld[i].TotCredit:=Arrondi(TSld[i].TotCredit+SC,Dec) ;
   END ;
TD1:=0 ; TC1:=0 ;
end ;

{=============================================================================}
procedure positionneTotaux3(Qcalc : TQuery ; Var TMvt,TSld,TMvtEuro,TSldEuro : TabTot ; EnSoldeTotal : Boolean ;
                            Dec,DecE : Integer) ;
Var TD,TC,SD,SC,TD1,TC1 : Double ;
    TDE,TCE,SDE,SCE,TD1E,TC1E : Double ;
    STyp,AN,Exo  : String3 ;
    i : Byte ;
    Gen,Aux,OldGen,OldAux : String ;
    PremFois : Boolean ;
BEGIN
{TD:=0 ; TC:=0 ;} TD1:=0 ; TC1:=0 ;
{TDE:=0 ; TCE:=0 ;} TD1E:=0 ; TC1E:=0 ; PremFois:=TRUE ;
While Not QCalc.EOF Do
  BEGIN
  AN:='N' ; OldGen:=Gen ; Oldaux:=Aux ;
  Gen:=QCalc.Fields[0].AsString ;
  Aux:=QCalc.Fields[1].AsString ;
  If PremFois Then BEGIN OldGen:=Gen ; OldAux:=Aux ; END ;
  Exo:=QCalc.Fields[2].AsString ;
  TD:=Arrondi(QCalc.Fields[3].AsFloat,Dec) ;
  TC:=Arrondi(QCalc.Fields[4].AsFloat,Dec) ;
  TDE:=Arrondi(QCalc.Fields[5].AsFloat,DecE) ;
  TCE:=Arrondi(QCalc.Fields[6].AsFloat,DecE) ;
  If EnSoldeTotal Then
     BEGIN
     If ((OldGen<>Gen) Or (OldAux<>Aux)) Then
        BEGIN
        CalcTotSolde(TD1,TC1,TMvt,TSld,Dec) ;
        CalcTotSolde(TD1E,TC1E,TMvtEuro,TSldEuro,DecE) ;
        END ;
     TD1:=Arrondi(TD1+TD,Dec) ; TC1:=Arrondi(TC1+TC,Dec) ;
     TD1E:=Arrondi(TD1E+TDE,DecE) ; TC1E:=Arrondi(TC1E+TCE,DecE) ;
     END Else
     BEGIN
     SD:=0 ; SC:=0 ; SDE:=0 ; SCE:=0 ;
     If TD>TC Then SD:=Arrondi(TD-TC,Dec) Else
        If TD<TC Then SC:=Arrondi(TC-TD,Dec) ;
     If TDE>TCE Then SDE:=Arrondi(TDE-TCE,DecE) Else
        If TDE<TCE Then SCE:=Arrondi(TCE-TDE,DecE) ;
     AN:=QCalc.Fields[7].AsString ;
     STyp:=QCalc.Fields[8].AsString ;
     i:=6 ;
     If AN<>'N' Then i:=0 Else If STyp='N' Then i:=1 Else If STyp='S' Then i:=2 Else
        If STyp='P' Then i:=3 Else If STyp='U' Then i:=4 Else If STyp='R' Then i:=5
        Else If STyp='I' Then i:=1 ; //FQ 13827 Prise en compte IFRS 09/09/2005 SBO
     If (TD<>0) Or (TC<>0) Then
        BEGIN
        TSld[i].TotDebit:=Arrondi(TSld[i].TotDebit+SD,Dec) ;
        TSld[i].TotCredit:=Arrondi(TSld[i].TotCredit+SC,Dec) ;
        TMvt[i].TotDebit:=Arrondi(TMvt[i].TotDebit+TD,Dec) ;
        TMvt[i].TotCredit:=Arrondi(TMvt[i].TotCredit+TC,Dec) ;
        END ;
     If (TDE<>0) Or (TCE<>0) Then
        BEGIN
        TSldEuro[i].TotDebit:=Arrondi(TSldEuro[i].TotDebit+SDE,DecE) ;
        TSldEuro[i].TotCredit:=Arrondi(TSldEuro[i].TotCredit+SCE,DecE) ;
        TMvtEuro[i].TotDebit:=Arrondi(TMvtEuro[i].TotDebit+TDE,DecE) ;
        TMvtEuro[i].TotCredit:=Arrondi(TMvtEuro[i].TotCredit+TCE,DecE) ;
        END ;
     END ;
  PremFois:=FALSE ;
  QCalc.Next ;
  END ;
If EnSoldeTotal Then
   BEGIN
   CalcTotSolde(TD1,TC1,TMvt,TSld,Dec) ;
   CalcTotSolde(TD1E,TC1E,TMvtEURO,TSldEURO,DecE) ;
   END ;
END ;

{=============================================================================}
procedure ExecuteTotCptSolde(Var Q : TQuery ; Cpt1 : String17 ; Date1,Date2 : TDateTime ; Devise,Etab,Exo : String3 ;
                             Var TotCptMvt,TotCptSld,TotCptMvtEuro,TotCptSldEuro : TabTot ; EnSoldeTotal : Boolean ; SommeTout : Boolean ;
                             Dec,DecE : Integer ; MonnaieOpposee : Boolean ; DeviseEnPivot : Boolean ; SetTyp : SetttTypePiece; stBase : string = '') ;
BEGIN
FillChar(TotCptMvt,SizeOf(TotCptMvt),#0) ; FillChar(TotCptSld,SizeOf(TotCptSld),#0) ;
FillChar(TotCptMvtEuro,SizeOf(TotCptMvtEuro),#0) ; FillChar(TotCptSldEuro,SizeOf(TotCptSldEuro),#0) ;
{$IFDEF NOVH}
{$ELSE}
If  EstSpecif('51185') Then If Q=NIL Then Exit ;
If not EstSpecif('51185') Then
  BEGIN
{$ENDIF NOVH}
  Q:=FabricReqCpt3(SetTyp,Devise<>'',etab<>'',exo<>'',DeviseEnPivot,Cpt1,Date1,Date2,Etab,Devise,Exo, stBase) ;
  PositionneTotaux3(Q,TotCptMvt,TotCptSld,TotCptMvtEuro,TotCptSldEuro,EnSoldeTotal,Dec,DecE) ;
  Ferme(Q) ; Q:=Nil ;
{$IFDEF NOVH}
{$ELSE}
  END Else
  BEGIN
  AttribParamsCpt3(Q,Cpt1,Date1,Date2,Etab,Devise,Exo) ;
  Q.Open ;
  PositionneTotaux3(Q,TotCptMvt,TotCptSld,TotCptMvtEuro,TotCptSldEuro,EnSoldeTotal,Dec,DecE) ;
  Q.Close ;
  END ;
{$ENDIF NOVH}
If SommeTout Then
   BEGIN
   Sommation(TotCptMvt,FALSE) ; Sommation(TotCptSld,FALSE) ;
   Sommation(TotCptMvtEuro,FALSE) ; Sommation(TotCptSldEuro,FALSE) ;
   END ;
If MonnaieOpposee Then
  BEGIN
  TotCptMvt:=TotCptMvtEuro ;
  TotCptSld:=TotCptSldEuro ;
  END ;
END ;

Function JalODA ( CodCpt : String3) : String3 ;
Var Q1   : TQuery ;
    Anal : Boolean ;
    Axe : String3 ;
BEGIN
Q1:=OpenSql('Select J_NATUREJAL,J_AXE from JOURNAL Where J_JOURNAL="'+CodCpt+'"',True) ;
Anal:=(Q1.Fields[0].AsString='ODA') Or (Q1.Fields[0].AsString='ANA') ;
Axe:=Q1.Fields[1].AsString ; Ferme(Q1) ; if Not Anal then Axe:='' ;
Result:=Axe ;
END ;

{=============================================================================}
Function FabricReqJal( CodCpt : String3 ; SetTyp : SetttTypePiece ; FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot : Boolean ;
                       Cpte : String ; Date1,Date2 : TDateTime ; Etab,Devise,Exo : String; stBase : string = '') : TQuery ;
Var Q   : TQuery ;
    StCalc,SQL : String ;
    P : String[2] ;
    Anal : Boolean ;
    Axe : String[3] ;
BEGIN
{Q:=NIL ;} SQL:='' ;
Axe:=JalODA(CodCpt) ; Anal:=(Axe<>'') ;
(*
Q:=TQuery.Create(Application) ;
Q.DatabaseName:=DBSOC.DataBaseName ;
Q.SessionName:=DBSOC.SessionName ;
Q.SQL.Clear ;
*)
if Anal then P:='Y_' else P:='E_' ;

StCalc:='SELECT '+P+'JOURNAL, '+P+'EXERCICE, ' ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
If EnDevise(FiltreDev,DeviseEnPivot) then StCalc:='sum('+P+'DEBITDEV), sum('+P+'CREDITDEV), '
                                     else StCalc:='sum('+P+'DEBIT), sum('+P+'CREDIT), ' ;
StCalc:=StCalc + ' 0, 0, ' ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
StCalc:=''+P+'ECRANOUVEAU, '+P+'QUALIFPIECE ' ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
if Anal then StCalc:='FROM ANALYTIQ '
        else StCalc:='FROM ECRITURE ' ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
{$IFDEF NOVH}
{$ELSE}
If not EstSpecif('51185') Then
  BEGIN
{$ENDIF NOVH}
  StCalc:='WHERE '+P+'JOURNAL="'+Cpte+'" AND '+P+'DATECOMPTABLE>="'+UsDateTime(Date1)+'" AND '+P+'DATECOMPTABLE<="'+UsDateTime(Date2)+'" ' ;
  //Q.SQL.Add(StCalc) ;
  SQL:=SQL+StCalc ;

  if FiltreDev then BEGIN StCalc:='AND '+P+'DEVISE="' + Devise + '" ' ; SQL:=SQL+StCalc ; END ;
  if FiltreEtab then BEGIN StCalc:='AND '+P+'ETABLISSEMENT="' + Etab + '" ' ; SQL:=SQL+StCalc ; END ;
  if FiltreExo then BEGIN StCalc:='AND '+P+'EXERCICE="' + Exo + '" ' ; SQL:=SQL+StCalc ; END ;
{$IFDEF NOVH}
{$ELSE}
  END Else
  BEGIN
  StCalc:='WHERE '+P+'JOURNAL="'+Cpte+'" AND '+P+'DATECOMPTABLE>="'+UsDateTime(Date1)+'" AND '+P+'DATECOMPTABLE<="'+UsDateTime(Date2)+'" ' ;
  //Q.SQL.Add(StCalc) ;
  SQL:=SQL+StCalc ;

  if FiltreDev then BEGIN StCalc:='AND '+P+'DEVISE="' + Devise + '" ' ; SQL:=SQL+StCalc ; END ;
  if FiltreEtab then BEGIN StCalc:='AND '+P+'ETABLISSEMENT="' + Etab + '" ' ; SQL:=SQL+StCalc ; END ;
  if FiltreExo then BEGIN StCalc:='AND '+P+'EXERCICE="' + Exo + '" ' ; SQL:=SQL+StCalc ; END ;
  END ;
{$ENDIF NOVH}

if Anal then
  BEGIN
  {b FP FQ16854}
  //StCalc:='AND Y_AXE="'+Axe+'" ' ;
  StCalc:='AND '+TSQLAnaCroise.ConditionAxe(Axe)+' ';
  {e FP FQ16854}
//  Q.SQL.Add(StCalc) ;
  SQL:=SQL+StCalc ;
  END ;
StCalc:=WhereSupp(P,SetTyp) ;
//If StCalc<>'' Then Q.SQL.Add(StCalc) ;
If StCalc<>'' Then SQL:=SQL+StCalc ;
StCalc:='GROUP BY '+P+'JOURNAL,'+P+'EXERCICE,'+P+'QUALIFPIECE,'+P+'ECRANOUVEAU ' ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
(*
ChangeSQL(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
Result:=Q ;
*)
{$IFDEF EAGLCLIENT}
Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase) ;
{$ELSE}
{$IFDEF NOVH}
Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase);
{$ELSE}
If not EstSpecif('51185') Then Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase)
                          Else Q:=PrepareSQL(SQL,TRUE) ;
{$ENDIF NOVH}
{$ENDIF}
Result:=Q ;
END ;

{=============================================================================}
Function PrepareTotJal(CodCpt:String3 ; SetTyp : SetttTypePiece ; FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot : Boolean; stBase : string = '') : TQuery ;
BEGIN
Result:=Nil ;
{$IFDEF NOVH}
{$ELSE}
If  EstSpecif('51185') Then Result:=FabricReqJal(CodCpt,SetTyp,FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot,'',iDate1900,iDate1900,'','','',stBase) ;
{$ENDIF NOVH}
END ;

{=============================================================================}
procedure ExecuteTotJal(Var Q : TQuery ; Cpt : String17 ; Date1,Date2 : TDateTime ; Devise,Etab,Exo : String3 ;
                        Var TotCpt,TotCptEuro : TabTot ; SommeTout : Boolean ; Dec,DecE : Integer ;
                        SetType : SetttTypePiece ; DeviseEnPivot : Boolean) ;
BEGIN
FillChar(TotCpt,SizeOf(TotCpt),#0) ;
FillChar(TotCptEuro,SizeOf(TotCptEuro),#0) ;
{$IFDEF NOVH}
{$ELSE}
If  EstSpecif('51185') Then If Q=NIL Then Exit ;
If not EstSpecif('51185') Then
  BEGIN
{$ENDIF NOVH}
  Q:=FabricReqJal(Cpt,SetType,Devise<>'',Etab<>'',Exo<>'',DeviseEnPivot,Cpt,Date1,Date2,Etab,Devise,Exo) ;
  PositionneTotaux(Q,TotCpt,TotCptEuro,Dec,DecE) ;
  Ferme(Q) ; Q:=Nil ;
  {$IFDEF NOVH}
  {$ELSE}
  END Else
  BEGIN
  AttribParamsCpt1(Q,Cpt,Date1,Date2,Etab,Devise,Exo,VarNull) ;
  Q.Open ;
  PositionneTotaux(Q,TotCpt,TotCptEuro,Dec,DecE) ;
  Q.Close ;
  END ;
  {$ENDIF NOVH}
If SommeTout Then
   BEGIN
   Sommation(TotCpt,FALSE) ; Sommation(TotCptEuro,FALSE) ;
   END ;
END ;


{=======!!!!!!!!!=====================}

{=============================================================================}
Function FabricReqCpt5 ( fb : TFichierBase ; SetTyp : SetttTypePiece ; FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot : Boolean ;
                         PrefixeJointure,ChampJointure,NomChampJointure,ValeurChampJointure : String ;
                         Date1,Date2 : TDateTime ; Etab,Devise,Exo : String; stBase : string) : TQuery ;
Var Q   : TQuery ;
    StCalc,SQL : String ;
    SUM_Montant : String ;
    P : String[2] ;
    TableJointure : String ;
BEGIN
{Q:=NIL ;} SQL:='' ;
If PrefixeJointure='G_' Then TableJointure:='GENERAUX' Else
   If PrefixeJointure='T_' Then TableJointure:='TIERS' Else
      If PrefixeJointure='S_' Then TableJointure:='SECTION' ;
(*
Q:=TQuery.Create(Application) ;
Q.DatabaseName:=DBSOC.DataBaseName ;
Q.SessionName:=DBSOC.SessionName ;
Q.SQL.Clear ;
*)
Case Fb Of
  FbGene : StCalc:='SELECT E_EXERCICE, ' ;
  FbAux  : StCalc:='SELECT E_AUXILIAIRE, E_EXERCICE, ';
  {b FP FQ16854}
  //fbAxe1..fbAxe5 : StCalc:='SELECT Y_SECTION, Y_EXERCICE, ' ;
  fbAxe1..fbAxe5 : StCalc:='SELECT '+TSQLAnaCroise.ChampSection(FbToAxe(Fb))+', Y_EXERCICE, ' ;
  {e FP FQ16854}
  END ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
Case fb Of
  fbGene,fbAux : If EnDevise(FiltreDev,DeviseEnPivot) Then SUM_Montant:='sum(E_DEBITDEV), sum(E_CREDITDEV), '
                                                      Else SUM_Montant:='sum(E_DEBIT), sum(E_CREDIT), ' ;
  fbAxe1..fbAxe5 : If EnDevise(FiltreDev,DeviseEnPivot) Then SUM_Montant:='sum(Y_DEBITDEV), sum(Y_CREDITDEV), '
                                                        Else SUM_Montant:='sum(Y_DEBIT), sum(Y_CREDIT), ' ;
  END ;
Case fb Of
  fbGene,fbAux : StCalc:=SUM_MONTANT+'E_ECRANOUVEAU, E_QUALIFPIECE FROM ECRITURE ' ;
  fbAxe1..fbAxe5 : StCalc:=SUM_MONTANT+'Y_ECRANOUVEAU, Y_QUALIFPIECE FROM ANALYTIQ ' ;
  END ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
(*
Case Fb Of
  FbGene : StCalc:='WHERE E_GENERAL=:CPTE AND E_DATECOMPTABLE>=:DD1 AND E_DATECOMPTABLE<=:DD2 ' ;
  FbAux  : StCalc:='WHERE E_AUXILIAIRE=:CPTE AND E_DATECOMPTABLE>=:DD1 AND E_DATECOMPTABLE<=:DD2 ';
  fbAxe1..fbAxe5 : StCalc:='WHERE Y_SECTION=:CPTE AND Y_AXE="'+fbToAxe(fb)+'" AND Y_DATECOMPTABLE>=:DD1 AND Y_DATECOMPTABLE<=:DD2 ' ;
  END ;
*)
Case Fb Of
  FbGene : StCalc:='Left Join '+TableJointure+' on E_'+ChampJointure+'='+
                    PrefixeJointure+ChampJointure+
                    ' Where '+NomChampJointure+'="'+ValeurChampJointure+'"'  ;
  (*
  FbAux  : StCalc:='WHERE E_AUXILIAIRE=:CPTE AND E_DATECOMPTABLE>=:DD1 AND E_DATECOMPTABLE<=:DD2 ';
  fbAxe1..fbAxe5 : StCalc:='WHERE Y_SECTION=:CPTE AND Y_AXE="'+fbToAxe(fb)+'" AND Y_DATECOMPTABLE>=:DD1 AND Y_DATECOMPTABLE<=:DD2 ' ;
  *)
  END ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
{$IFDEF NOVH}
{$ELSE}
If not EstSpecif('51185') Then
  BEGIN
{$ENDIF NOVH}
    Case Fb Of
      FbGene, FbAux : StCalc:=' AND E_DATECOMPTABLE>="'+UsDateTime(Date1)+'" AND E_DATECOMPTABLE<="'+UsDateTime(Date2)+'" ';
      {b FP FQ16854}
      //fbAxe1..fbAxe5 : StCalc:=' AND Y_AXE="'+fbToAxe(fb)+'" AND Y_DATECOMPTABLE>="'+UsDateTime(Date1)+'" AND Y_DATECOMPTABLE<="'+UsDateTime(Date2)+'" ';
      fbAxe1..fbAxe5 : StCalc:=' AND '+TSQLAnaCroise.ConditionAxe(fbToAxe(fb))+
                       ' AND Y_DATECOMPTABLE>="'+UsDateTime(Date1)+'" AND Y_DATECOMPTABLE<="'+UsDateTime(Date2)+'" ';
      {e FP FQ16854}
    END ;
  {$IFDEF NOVH}
  {$ELSE}
  END Else
  BEGIN
    Case Fb Of
      FbGene : StCalc:=' AND E_DATECOMPTABLE>=:DD1 AND E_DATECOMPTABLE<=:DD2 ' ;
      FbAux  : StCalc:=' AND E_DATECOMPTABLE>=:DD1 AND E_DATECOMPTABLE<=:DD2 ';
      {b FP FQ16854}
      //fbAxe1..fbAxe5 : StCalc:=' AND Y_AXE="'+fbToAxe(fb)+'" AND Y_DATECOMPTABLE>=:DD1 AND Y_DATECOMPTABLE<=:DD2 ' ;
      fbAxe1..fbAxe5 : StCalc:=' AND '+TSQLAnaCroise.ConditionAxe(fbToAxe(fb))+
                         ' AND Y_DATECOMPTABLE>=:DD1 AND Y_DATECOMPTABLE<=:DD2 ' ;
      {e FP FQ16854}
    END ;
  END ;
  {$ENDIF NOVH}

//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
{$IFDEF NOVH}
{$ELSE}
If not EstSpecif('51185') Then
begin
{$ENDIF NOVH}
  If FiltreDev Then
     BEGIN
     Case Fb Of
       FbGene,FbAux : StCalc:='AND E_DEVISE="'+Devise+'" ' ;
       fbAxe1..fbAxe5 : StCalc:='AND Y_DEVISE="'+Devise+'" ' ;
       END ;
  //   Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
  If FiltreEtab Then
     BEGIN
     Case Fb Of
       FbGene,FbAux : StCalc:='AND E_ETABLISSEMENT="'+Etab+'" ';
       fbAxe1..fbAxe5 : StCalc:='AND Y_ETABLISSEMENT="'+Etab+'" ';
       END ;
  //   Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
  If FiltreExo Then
     BEGIN
     Case Fb Of
       FbGene,FbAux : StCalc:='AND E_EXERCICE="'+Exo+'" ';
       fbAxe1..fbAxe5 : StCalc:='AND Y_EXERCICE="'+Exo+'" ';
       END ;
  //   Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
{$IFDEF NOVH}
{$ELSE}
end else
begin
  If FiltreDev Then
     BEGIN
     Case Fb Of
       FbGene,FbAux : StCalc:='AND E_DEVISE=:DEV ' ;
       fbAxe1..fbAxe5 : StCalc:='AND Y_DEVISE=:DEV ' ;
       END ;
  //   Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
  If FiltreEtab Then
     BEGIN
     Case Fb Of
       FbGene,FbAux : StCalc:='AND E_ETABLISSEMENT=:ETAB ' ;
       fbAxe1..fbAxe5 : StCalc:='AND Y_ETABLISSEMENT=:ETAB ' ;
       END ;
  //   Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
  If FiltreExo Then
     BEGIN
     Case Fb Of
       FbGene,FbAux : StCalc:='AND E_EXERCICE=:EXO ' ;
       fbAxe1..fbAxe5 : StCalc:='AND Y_EXERCICE=:EXO ' ;
       END ;
  //   Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
end;
{$ENDIF NOVH}
P:='E_' ; If fb In [fbAxe1..fbAxe5] Then P:='Y_' ;
StCalc:=WhereSupp(P,SetTyp) ;
If StCalc<>'' Then SQL:=SQL+StCalc ; //Q.SQL.Add(StCalc) ;
Case Fb Of
  FbGene : StCalc:='GROUP BY E_EXERCICE,E_ECRANOUVEAU,E_QUALIFPIECE ' ;
  FbAux  : StCalc:='GROUP BY E_AUXILIAIRE,E_EXERCICE,E_ECRANOUVEAU,E_QUALIFPIECE ' ;
  {b FP FQ16854}
  //fbAxe1..fbAxe5 : StCalc:='GROUP BY Y_SECTION,Y_EXERCICE,Y_ECRANOUVEAU,Y_QUALIFPIECE ' ;
  fbAxe1..fbAxe5 : StCalc:='GROUP BY '+TSQLAnaCroise.ChampSection(FbToAxe(Fb))+',Y_EXERCICE,Y_ECRANOUVEAU,Y_QUALIFPIECE ' ;
  {e FP FQ16854}
  END ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
(*
ChangeSQL(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
Result:=Q ;
*)
{$IFDEF EAGLCLIENT}
Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase) ;
{$ELSE}
{$IFDEF NOVH}
Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase);
{$ELSE}
If not EstSpecif('51185') Then Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase)
              Else Q:=PrepareSQL(SQL,TRUE) ;
{$ENDIF NOVH}
{$ENDIF}
Result:=Q ;

END ;

{=============================================================================}
Function PrepareTotCptJointure(fb : TFichierBase ; SetTyp : SetttTypePiece ; FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot : Boolean ;
                               PrefixeJointure,ChampJointure,NomChampJointure,ValeurChampJointure : String; stBase : string = '') : TQuery ;
BEGIN
Result:=NIL ;
{$IFDEF NOVH}
{$ELSE}
If  EstSpecif('51185') Then
 Result:=FabricReqCpt5(fb,SetTyp,FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot,
                       PrefixeJointure,ChampJointure,NomChampJointure,ValeurChampJointure,
                       iDate1900,iDate1900,'','','',stBase) ;
{$ENDIF NOVH}
END ;

{=============================================================================}
procedure AttribParamsCpt5( QX : TQuery ; Date1,Date2 : TDateTime ; Etab,Devise,Exo : String3) ;
BEGIN
{$IFNDEF EAGLCLIENT}
QX.ParamByName('DD1').AsDateTime:=Date1 ; QX.ParamByName('DD2').AsDateTime:=Date2 ;
If Devise<>'' Then QX.ParamByName('DEV').AsString:=Devise ;
If Etab<>'' Then QX.ParamByName('ETAB').AsString:=Etab ;
If Exo<>'' Then QX.ParamByName('EXO').AsString:=Exo ;
{$ENDIF}
END ;

{=============================================================================}
procedure positionneTotaux5(Qcalc : TQuery ; Var T : TabTot ; Dec,DecE : Integer) ;
Var TD,TC : Double ;
    STyp,AN,Exo  : String3 ;
    i : Byte ;
BEGIN
While Not QCalc.EOF Do
  BEGIN
  AN:='N' ;
  Exo:=QCalc.Fields[0].AsString ;
  TD:=Arrondi(QCalc.Fields[1].AsFloat,Dec) ;
  TC:=Arrondi(QCalc.Fields[2].AsFloat,Dec) ;
  AN:=QCalc.Fields[3].AsString ;
  STyp:=QCalc.Fields[4].AsString ;
  i:=6 ;
  If AN<>'N' Then i:=0 Else If STyp='N' Then i:=1 Else If STyp='S' Then i:=2 Else
     If STyp='P' Then i:=3 Else If STyp='U' Then i:=4 Else If STyp='R' Then i:=5
     Else If STyp='I' Then i:=1 ; //FQ 13827 Prise en compte IFRS 09/09/2005 SBO
  If (TD<>0) Or (TC<>0) Then
     BEGIN
     T[i].TotDebit:=Arrondi(T[i].TotDebit+TD,Dec) ;
     T[i].TotCredit:=Arrondi(T[i].TotCredit+TC,Dec) ;
     END ;
  QCalc.Next ;
  END ;
END ;

{=============================================================================}
procedure ExecuteTotCptJointure(Var Q : TQuery ; Date1,Date2 : TDateTime ; Devise,Etab,Exo : String3 ;
                                Var TotCpt : TabTot ; SommeTout : Boolean ; Dec,DecE : Integer ;
                                fb : tfichierBase ; SetTyp : SetttTypePiece ; DeviseEnPivot : Boolean ;
                                PrefixeJointure,ChampJointure,NomChampJointure,ValeurChampJointure : String; stBase : string = '') ;
BEGIN
FillChar(TotCpt,SizeOf(TotCpt),#0) ;
{$IFDEF NOVH}
{$ELSE}
If  EstSpecif('51185') Then If Q=NIL Then Exit ;
If not EstSpecif('51185') Then
  BEGIN
{$ENDIF NOVH}
  Q:=FabricReqCpt5(fb,SetTyp,Devise<>'',Etab<>'',Exo<>'',DeviseEnPivot,
                        PrefixeJointure,ChampJointure,NomChampJointure,ValeurChampJointure,
                        Date1,Date2,Etab,Devise,Exo, stBase) ;
  PositionneTotaux5(Q,TotCpt,Dec,DecE) ;
  Ferme(Q) ; Q:=NIL ;
{$IFDEF NOVH}
{$ELSE}
  END Else
  BEGIN
  AttribParamsCpt5(Q,Date1,Date2,Etab,Devise,Exo) ;
  Q.Open ;
  PositionneTotaux5(Q,TotCpt,Dec,DecE) ;
  Q.Close ;
  END ;
{$ENDIF NOVH}
If SommeTout Then Sommation(TotCpt,FALSE) ;
END ;

function WhatTypeEcr(Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState) : SetttTypePiece ;
//Type ttTypePiece  = (tpReel,tpSim,tpPrev,tpSitu,tpRevi,tpIfrs) ;
BEGIN
Result:=[] ;
if (Valeur<>'TOU') Then
   BEGIN
   if Valeur='NOR' then Result:=[tpReel]  ;
   if Valeur='NSS' then Result:=[tpReel,tpSim,tpSitu] ;
   if Valeur='SSI' then Result:=[tpSim,tpSitu] ;
   if Valeur='PRE' then Result:=[tpPrev] ;
   END Else Result:=[tpReel,tpSim,tpSitu,tpPrev] ;
If Controleur Then
   BEGIN
   Case EtatRevision Of
     cbchecked : Result:=[tpRevi] ;
     cbGrayed  : Result:=Result+[tpRevi] ;
     END ;
   END ;
END ;

function WhatTypeEcr2(St : String ) : SetttTypePiece;
Var SetTyp : SetttTypePiece ;
BEGIN
SetTyp:=[] ;
If Pos('N',St)>0 Then SetTyp:=SetTyp+[tpReel] ;
If Pos('S',St)>0 Then SetTyp:=SetTyp+[tpSim] ;
If Pos('P',St)>0 Then SetTyp:=SetTyp+[tpPrev] ;
If Pos('U',St)>0 Then SetTyp:=SetTyp+[tpSitu] ;
If Pos('R',St)>0 Then SetTyp:=SetTyp+[tpRevi] ;
If Pos('I',St)>0 Then SetTyp:=SetTyp+[tpIfrs] ;
Result:=SetTyp ;
END ;

Function WhatExiste(PreE,PreJ,PreB,Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState ;
                    LDate1,LDate2 : TDateTime ; CodeExo : String ; WhatQualif : Byte ; DEV : String ; Axe : String = '' ;
                    FiltreSup : String = '') : String ;
Var StTable,St,St2 : String ;
    SetType : SetttTypePiece ;
begin
if PreE='Y_' then StTable:='ANALYTIQ' else If PreE='BE_' Then StTable:='BUDECR' Else StTable:='ECRITURE' ;
St:='(exists (Select '+PreE+PreJ+','+PreE+'EXERCICE,'+PreE+'DATECOMPTABLE From '+StTable+' ' ;
St:=St+' Where ('+PreE+PreJ+'=Q.'+PreB+PreJ+' ' ;
If CodeExo<>'' Then St:=St+' And '+PreE+'EXERCICE="'+CodeExo+'" ';
St:=St+' And '+PreE+'DATECOMPTABLE>="'+UsDateTime(LDate1)+'" And '+PreE+'DATECOMPTABLE<="'+UsDateTime(LDate2)+'" ' ;
If PreE='BE_' Then
   BEGIN
   If DEV<>'' Then St:=St+' And '+PreE+'BUDJAL="'+DEV+'" ' ;
   END Else
   BEGIN
   If DEV<>'' Then St:=St+' And '+PreE+'DEVISE="'+DEV+'" ' ;
   END ;
If (PreE='Y_') And (Axe<>'') Then
  BEGIN
  {b FP FQ16854}
  //St:=St+' And '+PreE+'AXE="'+Axe+'" ' ;
  St:=St+' And '+TSQLAnaCroise.ConditionAxe(Axe)+' ';
  {e FP FQ16854}
  END ;
If WhatQualif=0 Then SetType:=WhatTypeEcr(Valeur,Controleur,EtatRevision) Else
   If WhatQualif=1 Then SetType:=WhatTypeEcr2(Valeur) ;
St2:=WhereSupp(PreE,SetType) ;
If St2<>'' Then St:=St+St2 ;
If FiltreSup<>'' Then St:=St+' AND '+FiltreSup+' ' ;
St:=St+')))' ;
Result:=St ;
end ;

Function WhatExisteNat(PreE,PreJ,PreB,PreJ1,Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState ;
                       LDate1,LDate2 : TDateTime ; CodeExo : String ; WhatQualif : Byte ; DEV : String) : String ;
Var StTable,St,St2 : String ;
    SetType : SetttTypePiece ;
begin
if PreE='Y_' then StTable:='ANALYTIQ' else StTable:='ECRITURE' ;
St:='(exists (Select '+PreE+PreJ+','+PreE+'EXERCICE,'+PreE+'DATECOMPTABLE From '+StTable+' ' ;
St:=St+' Where ('+PreE+PreJ+'=Q.'+PreB+PreJ1+' ' ;
St:=St+' And '+PreE+'EXERCICE="'+CodeExo+'" ';
St:=St+' And '+PreE+'DATECOMPTABLE>="'+UsDateTime(LDate1)+'" And '+PreE+'DATECOMPTABLE<="'+UsDateTime(LDate2)+'" ' ;
If DEV<>'' Then
   St:=St+' And '+PreE+'DEVISE="'+DEV+'" ' ;
If WhatQualif=0 Then SetType:=WhatTypeEcr(Valeur,Controleur,EtatRevision) Else
   If WhatQualif=1 Then SetType:=WhatTypeEcr2(Valeur) ;
St2:=WhereSupp(PreE,SetType) ;
If St2<>'' Then St:=St+St2 ;
St:=St+')))' ;
Result:=St ;
end ;

Function WhatExisteMul(PreE,PreJ1,PreJ2,PreB,Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState ;
                       LDate1,LDate2 : TDateTime ; CodeExo,Where1,Where2 : String ; WhatQualif : Byte ; Axe : String = '') : String ;
Var StTable,St,St2 : String ;
    SetType : SetttTypePiece ;
begin
if PreE='Y_' then StTable:='ANALYTIQ' else If PreE='BE_' Then StTable:='BUDECR' Else StTable:='ECRITURE' ;
St:='(exists (Select '+PreE+PreJ1+','+PreE+PreJ2+','+PreE+'EXERCICE,'+PreE+'DATECOMPTABLE From '+StTable+' ' ;
St:=St+' Where ('+PreE+PreJ1+'='+where1+' And '+PreE+PreJ2+'=Q1.'+PreB+PreJ2+' '+Where2+' ' ;
If CodeExo<>'' Then St:=St+' And '+PreE+'EXERCICE="'+CodeExo+'" ';
St:=St+' And '+PreE+'DATECOMPTABLE>="'+UsDateTime(LDate1)+'" And '+PreE+'DATECOMPTABLE<="'+UsDateTime(LDate2)+'" ' ;
If (PreE='Y_') And (Axe<>'') Then
  BEGIN
  {b FP FQ16854}
  //St:=St+' And '+PreE+'AXE="'+Axe+'" ' ;
  St:=St+' And '+TSQLAnaCroise.ConditionAxe(Axe)+' ';
  {e FP FQ16854}
  END ;
If WhatQualif=0 Then SetType:=WhatTypeEcr(Valeur,Controleur,EtatRevision) Else
   If WhatQualif=1 Then SetType:=WhatTypeEcr2(Valeur) ;
St2:=WhereSupp(PreE,SetType) ;
If St2<>'' Then St:=St+St2 ;
St:=St+')))' ;
Result:=St ;
end ;

Function WhatExisteMulNat(PreE,PreJ1,PreJ2,PreB,Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState ;
                       LDate1,LDate2 : TDateTime ; CodeExo,Where1,Where2 : String ; WhatQualif : Byte) : String ;
Var StTable,St,St2 : String ;
    SetType : SetttTypePiece ;
begin
if PreE='Y_' then StTable:='ANALYTIQ' else StTable:='ECRITURE' ;
St:='(exists (Select '+PreE+PreJ1+','+PreE+PreJ2+','+PreE+'EXERCICE,'+PreE+'DATECOMPTABLE From '+StTable+' ' ;
St:=St+' Where ('+PreE+PreJ1+'='+Where1+' And '+PreE+PreJ2+'=Q1.'+Where2+' ' ;
St:=St+' And '+PreE+'EXERCICE="'+CodeExo+'" ';
St:=St+' And '+PreE+'DATECOMPTABLE>="'+UsDateTime(LDate1)+'" And '+PreE+'DATECOMPTABLE<="'+UsDateTime(LDate2)+'" ' ;
If WhatQualif=0 Then SetType:=WhatTypeEcr(Valeur,Controleur,EtatRevision) Else
   If WhatQualif=1 Then SetType:=WhatTypeEcr2(Valeur) ;
St2:=WhereSupp(PreE,SetType) ;
If St2<>'' Then St:=St+St2 ;
St:=St+')))' ;
Result:=St ;
end ;

Function WhereSuppEdtTiers(Valeur : String ; Controleur : Boolean ; EtatRevision : TCheckBoxState) : String ;
Var SetType : SetttTypePiece ;
BEGIN
SetType:=WhatTypeEcr(Valeur,Controleur,EtatRevision) ;
Result:=WhereSupp('E_',SetType) ;
END ;

{=============================================================================}
procedure positionneTotauxUDF(Qcalc : TQuery ; Var T : TabTot ; Var TEURO : TabTot ; Dec,DecE : Integer) ;
Var TD,TC,TDE,TCE : Double ;
    STyp,AN,Exo  : String3 ;
//    Mois : Integer ;
    i : Byte ;
    Cpt : String ;
BEGIN
While Not QCalc.EOF Do
  BEGIN
  AN:='N' ;
  Cpt:=QCalc.Fields[0].AsString ;
  Exo:=QCalc.Fields[1].AsString ;
//  Mois:=QCalc.Fields[2].AsInteger ;
  TD:=Arrondi(QCalc.Fields[3].AsFloat,Dec) ;
  TC:=Arrondi(QCalc.Fields[4].AsFloat,Dec) ;
  TDE:=Arrondi(QCalc.Fields[5].AsFloat,DecE) ;
  TCE:=Arrondi(QCalc.Fields[6].AsFloat,DecE) ;
  AN:=QCalc.Fields[7].AsString ;
  STyp:=QCalc.Fields[8].AsString ;
  i:=6 ;
  If AN<>'N' Then i:=0 Else If STyp='N' Then i:=1 Else If STyp='S' Then i:=2 Else
     If STyp='P' Then i:=3 Else If STyp='U' Then i:=4 Else If STyp='R' Then i:=5
     Else If STyp='I' Then i:=1 ; //FQ 13827 Prise en compte IFRS 09/09/2005 SBO
  If AN='C' Then i:=6 ;
  If (TD<>0) Or (TC<>0) Then
     BEGIN
     T[i].TotDebit:=Arrondi(T[i].TotDebit+TD,Dec) ;
     T[i].TotCredit:=Arrondi(T[i].TotCredit+TC,Dec) ;
     END ;
  If (TDE<>0) Or (TCE<>0) Then
     BEGIN
     TEURO[i].TotDebit:=Arrondi(TEURO[i].TotDebit+TDE,DecE) ;
     TEURO[i].TotCredit:=Arrondi(TEURO[i].TotCredit+TCE,DecE) ;
     END ;

  QCalc.Next ;
  END ;
END ;

{=============================================================================}
procedure positionneTotauxBud(Qcalc : TQuery ; Var T : TabTot ; Var TEURO : TabTot ; Dec,DecE : Integer) ;
Var TD,TC,TDE,TCE : Double ;
    STyp,AN,Exo  : String3 ;
    i : Byte ;
    Cpt : String ;
BEGIN
While Not QCalc.EOF Do
  BEGIN
  AN:='N' ;
  Cpt:=QCalc.Fields[0].AsString ;
  Exo:=QCalc.Fields[1].AsString ;
  TD:=Arrondi(QCalc.Fields[2].AsFloat,Dec) ;
  TC:=Arrondi(QCalc.Fields[3].AsFloat,Dec) ;
  TDE:=Arrondi(QCalc.Fields[4].AsFloat,DecE) ;
  TCE:=Arrondi(QCalc.Fields[5].AsFloat,DecE) ;
  AN:=QCalc.Fields[6].AsString ;
  STyp:=QCalc.Fields[7].AsString ;
  i:=6 ;
  If STyp='N' Then i:=1 Else If STyp='S' Then i:=2 Else
     If STyp='P' Then i:=3 Else If STyp='U' Then i:=4 Else If STyp='R' Then i:=5
     Else If STyp='I' Then i:=1 ; //FQ 13827 Prise en compte IFRS 09/09/2005 SBO
  If (TD<>0) Or (TC<>0) Then
     BEGIN
     T[i].TotDebit:=Arrondi(T[i].TotDebit+TD,Dec) ;
     T[i].TotCredit:=Arrondi(T[i].TotCredit+TC,Dec) ;
     END ;
  If (TDE<>0) Or (TCE<>0) Then
     BEGIN
     TEURO[i].TotDebit:=Arrondi(TEURO[i].TotDebit+TDE,DecE) ;
     TEURO[i].TotCredit:=Arrondi(TEURO[i].TotCredit+TCE,DecE) ;
     END ;

  QCalc.Next ;
  END ;
END ;
{=============================================================================}
Function FabricReqCpt1Bud ( fb : TFichierBase ; SetTyp : SetttTypePiece ; FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot : Boolean ; Multiple : TTypeCalc ;
                            Cpte : String ; Date1,Date2 : TDateTime ; Etab,Devise,Exo : String; stBase : string) : TQuery ;
Var Q   : TQuery ;
    StCalc,SQL : String ;
    SUM_Montant : String ;
BEGIN
{Q:=NIL ;} SQL:='' ;
(*
Q:=TQuery.Create(Application) ;
Q.DatabaseName:=DBSOC.DataBaseName ;
Q.SessionName:=DBSOC.SessionName ;
Q.SQL.Clear ;
*)
Case Fb Of
  FbBudGen : StCalc:='SELECT BE_BUDGENE, BE_EXERCICE, ' ;
  fbBudSec1..fbBudSec5 : StCalc:='SELECT BE_BUDSECT, BE_EXERCICE, ' ;
  END ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
SUM_Montant:='sum(BE_DEBIT), sum(BE_CREDIT), 0, 0, ' ;
StCalc:=SUM_MONTANT+'BE_NATUREBUD, BE_QUALIFPIECE FROM BUDECR ' ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
{$IFDEF NOVH}
{$ELSE}
If not EstSpecif('51185') Then
  BEGIN
{$ENDIF NOVH}
  Case Fb Of
    FbBudGen : StCalc:='WHERE BE_BUDGENE="'+Cpte+'" AND BE_DATECOMPTABLE>="'+UsDateTime(Date1)+'" AND BE_DATECOMPTABLE<="'+UsDateTime(Date2)+'" ' ;
    fbBudSec1..fbBudSec5 : StCalc:='WHERE BE_BUDSECT="'+Cpte+'" AND BE_AXE="'+fbToAxe(fb)+'" AND BE_DATECOMPTABLE>="'+UsDateTime(Date1)+'" AND BE_DATECOMPTABLE<="'+UsDateTime(Date2)+'" ' ;
    END ;
{$IFDEF NOVH}
{$ELSE}
  END Else
  BEGIN
  Case Fb Of
    FbBudGen : StCalc:='WHERE BE_BUDGENE=:CPTE AND BE_DATECOMPTABLE>=:DD1 AND BE_DATECOMPTABLE<=:DD2 ' ;
    fbBudSec1..fbBudSec5 : StCalc:='WHERE BE_BUDSECT=:CPTE AND BE_AXE="'+fbToAxe(fb)+'" AND BE_DATECOMPTABLE>=:DD1 AND BE_DATECOMPTABLE<=:DD2 ' ;
    END ;
  END ;
{$ENDIF NOVH}
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
{$IFDEF NOVH}
{$ELSE}
If not EstSpecif('51185') Then
  BEGIN
{$ENDIF NOVH}
  If FiltreDev Then { Pour Budget : FiltreDev utilisé pour filtre sur journal budgétaire d'unr écriture budgétaire }
     BEGIN
     StCalc:='AND BE_BUDJAL="'+Devise+'" ' ; //Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
  {$IFDEF NOVH}
  {$ELSE}
  END Else
  BEGIN
  If FiltreDev Then { Pour Budget : FiltreDev utilisé pour filtre sur journal budgétaire d'unr écriture budgétaire }
     BEGIN
     StCalc:='AND BE_BUDJAL=:DEV ' ; //Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
  END ;
If not EstSpecif('51185') Then
  BEGIN
  {$ENDIF NOVH}
  If FiltreEtab Then
     BEGIN
     StCalc:='AND BE_ETABLISSEMENT="'+Etab+'" ' ;
     //Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
  {$IFDEF NOVH}
  {$ELSE}
  END Else
  BEGIN
  If FiltreEtab Then
     BEGIN
     StCalc:='AND BE_ETABLISSEMENT=:ETAB ' ;
     //Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
  END ;
If not EstSpecif('51185') Then
  BEGIN
  {$ENDIF NOVH}
  If FiltreExo Then
     BEGIN
     StCalc:='AND BE_EXERCICE="'+Exo+'" ' ;
     //Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
  {$IFDEF NOVH}
  {$ELSE}
  END Else
  BEGIN
  If FiltreExo Then
     BEGIN
     StCalc:='AND BE_EXERCICE=:EXO ' ;
     //Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
  END ;
  {$ENDIF NOVH}

Case Fb Of
  FbBudGen : StCalc:='GROUP BY BE_BUDGENE,BE_EXERCICE,BE_NATUREBUD,BE_QUALIFPIECE ' ;
  fbBudSec1..fbBudSec5 : StCalc:='GROUP BY BE_BUDSECT,BE_EXERCICE,BE_NATUREBUD,BE_QUALIFPIECE ' ;
  END ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
(*
ChangeSQL(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
Result:=Q ;
*)
{$IFDEF EAGLCLIENT}
Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase) ;
{$ELSE}
{$IFDEF NOVH}
  Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase);
{$ELSE}
If EstSpecif('51185') Then
  Q:=PrepareSQL(SQL,TRUE)
Else
  Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase);
{$ENDIF NOVH}
{$ENDIF}
Result:=Q ;
END ;

{=============================================================================}
procedure positionneTotaux2Bud(Qcalc : TQuery ; Var T,TEuro : TabTot ; Dec,DecE : Integer) ;
Var TD,TC,TDE,TCE : Double ;
    STyp,AN,Exo  : String3 ;
    i : Byte ;
    Cpt1,Cpt2 : String ;
BEGIN
While Not QCalc.EOF Do
  BEGIN
  AN:='N' ;
  Cpt1:=QCalc.Fields[0].AsString ;
  Cpt2:=QCalc.Fields[1].AsString ;
  Exo:=QCalc.Fields[2].AsString ;
  TD:=Arrondi(QCalc.Fields[3].AsFloat,Dec) ;
  TC:=Arrondi(QCalc.Fields[4].AsFloat,Dec) ;
  TDE:=Arrondi(QCalc.Fields[5].AsFloat,DecE) ;
  TCE:=Arrondi(QCalc.Fields[6].AsFloat,DecE) ;
  AN:=QCalc.Fields[7].AsString ;
  STyp:=QCalc.Fields[8].AsString ;
  i:=6 ;
  If STyp='N' Then i:=1 Else If STyp='S' Then i:=2 Else
     If STyp='P' Then i:=3 Else If STyp='U' Then i:=4 Else If STyp='R' Then i:=5
     Else If STyp='I' Then i:=1 ; //FQ 13827 Prise en compte IFRS 09/09/2005 SBO
  If (TD<>0) Or (TC<>0) Then
     BEGIN
     T[i].TotDebit:=Arrondi(T[i].TotDebit+TD,Dec) ;
     T[i].TotCredit:=Arrondi(T[i].TotCredit+TC,Dec) ;
     END ;
  If (TDE<>0) Or (TCE<>0) Then
     BEGIN
     TEuro[i].TotDebit:=Arrondi(TEuro[i].TotDebit+TDE,DecE) ;
     TEuro[i].TotCredit:=Arrondi(TEuro[i].TotCredit+TCE,DecE) ;
     END ;
  QCalc.Next ;
  END ;
END ;

{=============================================================================}
Function FabricReqCpt2Bud ( fb,fb2 : TFichierBase ; SetTyp : SetttTypePiece ; FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot : Boolean ; Multiple : TTypeCalc ;
                            Cpte1,Cpte2 : String ; Date1,Date2 : TDateTime ; Etab,Devise,Exo : String; stBase : string='') : TQuery ;
Var Q   : TQuery ;
    StCalc,SQL : String ;
    SUM_Montant : String ;
BEGIN
{Q:=NIL ;} SQL:='' ;
(*
Q:=TQuery.Create(Application) ;
Q.DatabaseName:=DBSOC.DataBaseName ;
Q.SessionName:=DBSOC.SessionName ;
Q.SQL.Clear ;
*)
StCalc:='SELECT ' ;
Case Fb Of
  FbBudGen : StCalc:=Stcalc+'BE_BUDGENE, ' ;
  fbBudSec1..fbBudSec5 : StCalc:=Stcalc+'BE_BUDSECT, ' ;
  END ;
Case Fb2 Of
  FbBudGen : StCalc:=Stcalc+'BE_BUDGENE, ' ;
  fbBudSec1..fbBudSec5 : StCalc:=Stcalc+'BE_BUDSECT, ' ;
  END ;
StCalc:=StCalc+'BE_EXERCICE, ' ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
SUM_Montant:='sum(BE_DEBIT), sum(BE_CREDIT), 0, 0, ' ;
StCalc:=SUM_MONTANT+'BE_NATUREBUD, BE_QUALIFPIECE FROM BUDECR ' ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
{$IFDEF NOVH}
{$ELSE}
If not EstSpecif('51185') Then
  BEGIN
{$ENDIF NOVH}
  Case Fb Of
    FbBudGen : StCalc:='WHERE BE_BUDGENE="'+Cpte1+'" AND ' ;
    fbBudSec1..fbBudSec5 : StCalc:='WHERE BE_BUDSECT="'+Cpte1+'" AND ' ;
    END ;
  Case Fb2 Of
    FbBudGen : StCalc:=StCalc+' BE_BUDGENE="'+Cpte2+'" AND ' ;
    fbBudSec1..fbBudSec5 : StCalc:=StCalc+' BE_BUDSECT="'+Cpte2+'" AND ' ;
    END ;

  If (Fb in [fbBudSec1..fbBudSec5]) THEN StCalc:=STCalc+'BE_AXE="'+fbToAxe(fb)+'" AND ' Else
    If (Fb2 in [fbBudSec1..fbBudSec5]) THEN StCalc:=STCalc+'BE_AXE="'+fbToAxe(fb2)+'" AND ' ;
  If FiltreExo Then StCalc:=StCalc+'BE_EXERCICE="'+Exo+'" AND ' ;
  StCalc:=StCalc+'BE_DATECOMPTABLE>="'+UsDateTime(Date1)+'" AND BE_DATECOMPTABLE<="'+UsDateTime(Date2)+'" ' ;
{$IFDEF NOVH}
{$ELSE}
  END Else
  BEGIN
  Case Fb Of
    FbBudGen : StCalc:='WHERE BE_BUDGENE=:CPTE1 AND ' ;
    fbBudSec1..fbBudSec5 : StCalc:='WHERE BE_BUDSECT=:CPTE1 AND ' ;
    END ;
  Case Fb2 Of
    FbBudGen : StCalc:=StCalc+' BE_BUDGENE=:CPTE2 AND ' ;
    fbBudSec1..fbBudSec5 : StCalc:=StCalc+' BE_BUDSECT=:CPTE2 AND ' ;
    END ;

  If (Fb in [fbBudSec1..fbBudSec5]) THEN StCalc:=STCalc+'BE_AXE="'+fbToAxe(fb)+'" AND ' Else
    If (Fb2 in [fbBudSec1..fbBudSec5]) THEN StCalc:=STCalc+'BE_AXE="'+fbToAxe(fb2)+'" AND ' ;
  If FiltreExo Then StCalc:=StCalc+'BE_EXERCICE=:EXO AND ' ;
  StCalc:=StCalc+'BE_DATECOMPTABLE>=:DD1 AND BE_DATECOMPTABLE<=:DD2 ' ;
  END ;
{$ENDIF NOVH}
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
{$IFDEF NOVH}
{$ELSE}
If not EstSpecif('51185') Then
  BEGIN
{$ENDIF NOVH}
  If FiltreDev Then { Pour Budget : FiltreDev utilisé pour filtre sur journal budgétaire d'unr écriture budgétaire }
     BEGIN
     StCalc:='AND BE_BUDJAL="'+Devise+'" ' ; //Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
  If FiltreEtab Then
     BEGIN
     StCalc:='AND BE_ETABLISSEMENT="'+Etab+'" ' ; //Q.SQL.Add(StCalc) ;   //SG6 19/11/04 FQ 14999 "'+Etab+'" et non pas "'+exo+'"
     SQL:=SQL+StCalc ;
     END ;
{$IFDEF NOVH}
{$ELSE}
  END Else
  BEGIN
  If FiltreDev Then { Pour Budget : FiltreDev utilisé pour filtre sur journal budgétaire d'unr écriture budgétaire }
     BEGIN
     StCalc:='AND BE_BUDJAL=:DEV ' ; //Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
  If FiltreEtab Then
     BEGIN
     StCalc:='AND BE_ETABLISSEMENT=:ETAB ' ; //Q.SQL.Add(StCalc) ;
     SQL:=SQL+StCalc ;
     END ;
  END ;
{$ENDIF NOVH}
StCalc:='GROUP BY ' ;

Case Fb Of
  FbBudGen : StCalc:=StCalc+'BE_BUDGENE, ' ;
  fbBudSec1..fbBudSec5 : StCalc:=StCalc+'BE_BUDSECT, ' ;
  END ;
Case Fb2 Of
  FbBudGen : StCalc:=StCalc+'BE_BUDGENE, ' ;
  fbBudSec1..fbBudSec5 : StCalc:=StCalc+'BE_BUDSECT, ' ;
  END ;
If (Fb in [fbBudSec1..fbBudSec5]) Or (Fb2 in [fbBudSec1..fbBudSec5]) Then StCalc:=StCalc+'BE_AXE,' ;
StCalc:=StCalc+'BE_EXERCICE, BE_NATUREBUD, BE_QUALIFPIECE' ;
//Q.SQL.Add(StCalc) ;
SQL:=SQL+StCalc ;
(*
ChangeSQL(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
Result:=Q ;
*)
{$IFDEF EAGLCLIENT}
Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase);
{$ELSE}
{$IFDEF NOVH}
Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase);
{$ELSE}
If not EstSpecif('51185') Then Q:=OpenSQL(SQL,TRUE,-1,'',False,stBase)
              Else Q:=PrepareSQL(SQL,TRUE) ;
{$ENDIF NOVH}
{$ENDIF}
Result:=Q ;
END ;



Constructor TGCalculCum.Create(FMultiple : TTypeCalc ; Ffb1,Ffb2 : TFichierBase ; FSetTyp : SetttTypePiece ; FFiltreDev,FFiltreEtab,FFiltreExo,FDeviseEnPivot,FEnEuro : Boolean ; FDec,FDecE : Integer ;
//                               FSpeedActif : Boolean = FALSE ; FHisto : Boolean = FALSE ; FFiltreSup : String = '') ;
                               FSpeedActif : Boolean = FALSE ; FBalSit : string = '' ; FFiltreSup : String = '' ;
                               FChampSup : String = '';FBase :string = '') ;
begin
Inherited Create ;
Fillchar(PrepCalc,SizeOf(PrepCalc),#0) ;
Fillchar(ExecCalc,SizeOf(ExecCalc),#0) ;
Multiple:=FMultiple ;
ExecCalc.ValChampSup:=Null ;
PrepCalc.Methode:=Multiple ;
PrepCalc.fb1:=Ffb1 ; PrepCalc.fb2:=Ffb2 ; PrepCalc.SetTyp:=FSetTyp ;
PrepCalc.FiltreDev:=FFiltreDev ; PrepCalc.FiltreEtab:=FFiltreEtab ; PrepCalc.FiltreExo:=FFiltreExo ;
PrePCalc.DeviseEnPivot:=FDeviseEnPivot ;
PrepCalc.Decimale:=FDec ; PrepCalc.DecimaleEuro:=FDecE ; PrepCalc.EnEuro:=FEnEuro ;
PrepCalc.tz1:=tzRien ; PrepCalc.tz2:=tzRien ; //PrepCalc.Histo:=FHisto ;
PrepCalc.BalSit:=FBalSit ;
PrepCalc.FiltreSup:=FFiltreSup ;
PrepCalc.ChampSup:=FChampSup ;
Fillchar(PrepCalc.OkNatLibre,SizeOf(PrepCalc.OkNatLibre),#0) ;
{$IFDEF NOVH}
UseTC:=FALSE;
{$ELSE}
If Not EstSpecif('51187') Then UseTC:=FALSE Else UseTC:=VH^.UseTC ;
{$ENDIF NOVH}
If PrepCalc.BalSit<>'' Then UseTC:=FALSE ;
If FSpeedActif Then
  BEGIN
  PrepCalc.SpeedActif:=TRUE ;
  Exit ;
  END ;
{$IFDEF NOVH}
{$ELSE}
If EstSpecif('51187') Then
  BEGIN
  If UseTC And (VH^.TOBCUM<>NIL) Then Exit ;
  END ;
FstBase := FBase;
If  EstSpecif('51185') Then
  BEGIN
  With PrepCalc Do
    Case Multiple Of
    Deux : BEGIN
           Inc(VH^.NbGC) ;
           Q:=FabricReqCpt2(fb1,fb2,SetTyp,FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot,Multiple,FiltreSup,ChampSup,
                            '','',iDate1900,iDate1900,'','','',NULL,FStBase) ;
           END ;
    Un   : BEGIN
           Inc(VH^.NbGC) ;
  //         If Histo Then Q:=FabricReqHisto1(fb1,SetTyp,FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot,Multiple)
           If FBalSit<>'' Then // CA - 21/11/2001
           begin
            Inc(VH^.NbGC) ;
            QBalSit:=FabricReqHisto1(fb1,SetTyp,FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot,Multiple,FBalSit,'','',FstBase);
           end;
           Q:=FabricReqCpt1(fb1,SetTyp,FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot,Multiple,FiltreSup,ChampSup,
                            '',iDate1900,iDate1900,'','','',NULL, FStBase) ;
           END ;
    UnBud: BEGIN
           Inc(VH^.NbGC) ;
           Q:=FabricReqCpt1Bud(fb1,SetTyp,FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot,Multiple,
                               '',iDate1900,iDate1900,'','','', FStBase) ;
           END ;
    DeuxBud: BEGIN
             Inc(VH^.NbGC) ;
             Q:=FabricReqCpt2Bud(fb1,fb2,SetTyp,FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot,Multiple,
                                 '','',iDate1900,iDate1900,'','','',FStBase) ;
             END ;
    (*
    DeuxSurTableLibre : Q:=FabricReqCptTableLibre2(fb1,fb2,SetTyp,FiltreDev,FiltreEtab,FiltreExo,DeviseEnPivot) ;
    *)
    END ;
  END ;
{$ENDIF NOVH}
end ;

{=============================================================================}
Constructor TGCalculCum.CreateNat(FMultiple : TTypeCalc ; Ffb1,Ffb2 : TFichierBase ;
                      FSetTyp : SetttTypePiece ; FFiltreDev,FFiltreEtab,FFiltreExo,FDeviseEnPivot,FEnEuro : Boolean ;
                      FDec,FDecE : Integer ; Ftz1,Ftz2 : TZoomTable) ;
begin
end ;

{=============================================================================}
Procedure DecodeOkNatLibre(St : String ; Var OkNatLibre : TOkNatLibre) ;
Var Nat : array[1..10] of String ;
    i : Integer ;
BEGIN
Fillchar(OkNatLibre,SizeOf(OkNatLibre),#0) ; If Trim(St)='' then Exit ;
Fillchar(Nat,SizeOf(Nat),#0) ;
i:=1 ; While (St<>'') And (i<=10) do BEGIN Nat[i]:=readTokenSt(St) ; Inc(i) ; END ;
For i:=1 To 10 Do If Not ((Nat[i]='#') or (Nat[i]='---') or (Nat[i]='-')) Then OkNatLibre[i]:=TRUE ;
END ;

Destructor TGCalculCum.Destroy ;
BEGIN
if QBalSit <> nil then begin
  QBalSit.Close;
  {$IFDEF NOVH}
  {$ELSE}
  Dec(VH^.NbGC) ;
  {$ENDIF NOVH}
  QBalSit.Free;
END;   // CA - 21/11/2001
If Q<>NIL Then BEGIN
  Q.Close ;
  {$IFDEF NOVH}
  {$ELSE}
  Dec(VH^.NbGC) ;
  {$ENDIF NOVH}
  Q.Free ;
END ;
Inherited Destroy ;
END ;

Procedure TGCalculCum.InitCalcul(C1,C2,A1,A2,DE,ET,EX : String ; D1,D2 : TDateTime ; ST : Boolean; SurBalSit : boolean = False) ; // CA - 21/11/2001
begin
If PrepCalc.SpeedActif Then
  BEGIN
//  Params.SurQuelCpI:=Code1 ;      Params.SurQuelCpO:=Code2 ;
  ExecCalc.Params.Fb1:=PrepCalc.fb1 ;      ExecCalc.Params.Fb2:=PrepCalc.fb2 ;
  ExecCalc.Params.Date1:=D1 ;              ExecCalc.Params.Date2:=D2 ;
  ExecCalc.Params.Exo:=EX ;                ExecCalc.Params.Etab:=ET ;
  ExecCalc.Params.Devise:=DE ;             ExecCalc.Params.SetTyp:=PrepCalc.SetTyp ;
  ExecCalc.Params.DevPivot:=FALSE ;        ExecCalc.Params.Multiple:=Multiple ;
  ExecCalc.Params.TypePlan:='' ;           ExecCalc.Params.TypeCumul:='' ;
  END ;
With ExecCalc Do
  BEGIN
  Cpt1:=C1 ; Cpt2:=C2 ; Axe1:=A1 ; Axe2:=A2 ; Date1:=D1 ; Date2:=D2 ; Devise:=DE ;
  Etab:=ET ; Exo:=EX ; SommeTout:=ST ;
  BalSit := SurBalSit;  // CA - 21/11/2001
  END ;
end ;

Procedure TGCalculCum.InitCalculVCS(C1,C2,A1,A2,DE,ET,EX : String ; D1,D2 : TDateTime ; ST : Boolean;
                                    VCS : Variant ; SurBalSit : boolean = False) ; // CA - 21/11/2001
begin
If PrepCalc.SpeedActif Then
  BEGIN
//  Params.SurQuelCpI:=Code1 ;      Params.SurQuelCpO:=Code2 ;
  ExecCalc.Params.Fb1:=PrepCalc.fb1 ;      ExecCalc.Params.Fb2:=PrepCalc.fb2 ;
  ExecCalc.Params.Date1:=D1 ;              ExecCalc.Params.Date2:=D2 ;
  ExecCalc.Params.Exo:=EX ;                ExecCalc.Params.Etab:=ET ;
  ExecCalc.Params.Devise:=DE ;             ExecCalc.Params.SetTyp:=PrepCalc.SetTyp ;
  ExecCalc.Params.DevPivot:=FALSE ;        ExecCalc.Params.Multiple:=Multiple ;
  ExecCalc.Params.TypePlan:='' ;           ExecCalc.Params.TypeCumul:='' ;
  END ;
With ExecCalc Do
  BEGIN
  Cpt1:=C1 ; Cpt2:=C2 ; Axe1:=A1 ; Axe2:=A2 ; Date1:=D1 ; Date2:=D2 ; Devise:=DE ;
  Etab:=ET ; Exo:=EX ; SommeTout:=ST ;
  BalSit := SurBalSit;  // CA - 21/11/2001
  If VCSOK(VCS) Then ValChampSup:=VCS Else ValChampSup:=Null ;
  END ;
end ;

Procedure TGCalculCum.ReInitCalcul(C1,C2 : String ; D1,D2 : TDateTime; Exercice : String3 = '';
                                   SurBalSit : boolean = False) ;  // CA - 21/11/2001
begin
With ExecCalc Do
  BEGIN
  If C1<>'' Then Cpt1:=C1 ;
  If C2<>'' Then Cpt2:=C2 ;
  If D1<>0 Then Date1:=D1 ;
  If D2<>0 Then Date2:=D2 ;

  { GC }
  if Exercice <> '' then Exo := Exercice;
  { }
  // CA - 21/11/2001 - récupération de l'indicateur de balance de situation
  BalSit := SurBalSit;

  END ;
end ;

Procedure TGCalculCum.ReInitCalculVCS(C1,C2 : String ; D1,D2 : TDateTime; VCS : Variant ;
                                      Exercice : String3 = '' ; SurBalSit : boolean = False) ;  // CA - 21/11/2001
begin
With ExecCalc Do
  BEGIN
  If C1<>'' Then Cpt1:=C1 ;
  If C2<>'' Then Cpt2:=C2 ;
  If D1<>0 Then Date1:=D1 ;
  If D2<>0 Then Date2:=D2 ;

  { GC }
  if Exercice <> '' then Exo := Exercice;
  { }
  // CA - 21/11/2001 - récupération de l'indicateur de balance de situation
  BalSit := SurBalSit;

  If VCSOK(VCS) Then ValChampSup:=VCS Else ValChampSup:=Null ;
  END ;
end ;

Procedure TGCalculCum.Calcul ;
Var {D11,D12,DCum1,DCum2,}D21,D22 : TdateTime ;
//    N1,N2 : Integer ;
    Avec6 : Boolean ;
begin
{$IFNDEF EAGLSERVER}
If EstSpecif('51187') Then
  BEGIN
  If UseTC And (VH^.TobCum<>NIL) Then
    BEGIN
    ExecCalc.totCpt:=DonneCumul(VH^.TOBCUM,ExecCalc.Cpt1,ExecCalc.Cpt2,Multiple,PrepCalc.Fb1,PrepCalc.Fb2,PrepCalc.SetTyp,
                                ExecCalc.Devise,ExecCalc.Etab,ExecCalc.Exo,PrepCalc.DeviseEnPivot,PrepCalc.EnEuro,
                                ExecCalc.Date1,ExecCalc.Date2,PrepCalc.ChampSup,ExecCalc.ValChampSup,PrepCalc.FiltreSup,FstBase) ;
    If ExecCalc.SommeTout Then
       BEGIN
       Sommation(ExecCalc.TotCpt,FALSE) ; Sommation(ExecCalc.TotCptEuro,FALSE) ;
       END ;
    Exit ;
    END ;
  END ;
{$ENDIF EAGLSERVER}

Case Multiple Of
Deux : BEGIN
       FillChar(ExecCalc.TotCpt,SizeOf(ExecCalc.TotCpt),#0) ;
       FillChar(ExecCalc.TotCptEuro,SizeOf(ExecCalc.TotCptEuro),#0) ;
       {$IFDEF NOVH}
       {$ELSE}
       If  EstSpecif('51185') Then If Q=NIL Then Exit ;
       {$ENDIF NOVH}
       {D11:=0 ; D12:=0 ;} D21:=ExecCalc.Date1 ; D22:=ExecCalc.Date2 ;
       {$IFDEF NOVH}
       {$ELSE}
       If not EstSpecif('51185') Then
         BEGIN
       {$ENDIF NOVH}
         Q:=FabricReqCpt2(prepCalc.fb1,prepCalc.fb2,prepCalc.SetTyp,prepCalc.FiltreDev,prepCalc.FiltreEtab,prepCalc.FiltreExo,prepCalc.DeviseEnPivot,Multiple,prepCalc.FiltreSup,prepCalc.ChampSup,
                          ExecCalc.Cpt1,ExecCalc.Cpt2,D21,D22,ExecCalc.Etab,ExecCalc.Devise,ExecCalc.Exo,ExecCalc.ValChampSup,FStBase) ;
         PositionneTotaux2(Q,ExecCalc.TotCpt,ExecCalc.TotCptEuro,PrepCalc.Decimale,PrepCalc.DecimaleEuro) ;
         Ferme(Q) ; Q:=Nil ;
         {$IFDEF NOVH}
         {$ELSE}
         END Else
         BEGIN
         AttribParamsCpt2(Q,ExecCalc.Cpt1,ExecCalc.Cpt2,ExecCalc.Date1,ExecCalc.Date2,ExecCalc.Etab,ExecCalc.Devise,ExecCalc.Exo,ExecCalc.ValChampSup) ;
         Q.Open ;
         PositionneTotaux2(Q,ExecCalc.TotCpt,ExecCalc.TotCptEuro,PrepCalc.Decimale,PrepCalc.DecimaleEuro) ;
         Q.Close ;
         END ;
         {$ENDIF NOVH}
       If ExecCalc.SommeTout Then
          BEGIN
          Sommation(ExecCalc.TotCpt,FALSE) ; Sommation(ExecCalc.TotCptEuro,FALSE) ;
          END ;
       END  ;
  Un : BEGIN
       FillChar(ExecCalc.TotCpt,SizeOf(ExecCalc.TotCpt),#0) ;
       FillChar(ExecCalc.TotCptEuro,SizeOf(ExecCalc.TotCptEuro),#0) ;
       {$IFDEF NOVH}
       {$ELSE}
       If  EstSpecif('51185') Then If Q=NIL Then Exit ;
       {$ENDIF NOVH}
       {D11:=0 ; D12:=0 ;} D21:=ExecCalc.Date1 ; D22:=ExecCalc.Date2 ;
       If ExecCalc.BalSit Then  // CA - 21/11/2001 : si indicateur BalSit, on calcule sur la balance
          BEGIN
          {$IFDEF NOVH}
          {$ELSE}
          If not EstSpecif('51185') Then
            BEGIN
          {$ENDIF NOVH}
            QBalSit:=FabricReqHisto1(PrepCalc.fb1,PrepCalc.SetTyp,PrepCalc.FiltreDev,PrepCalc.FiltreEtab,PrepCalc.FiltreExo,PrepCalc.DeviseEnPivot,Multiple,PrepCalc.BalSit,ExecCalc.Cpt1,prepCalc.FiltreSup,FStBase);
            PositionneTotauxHisto(QBalSit,ExecCalc.TotCpt,ExecCalc.TotCptEuro,PrepCalc.Decimale,PrepCalc.DecimaleEuro) ;
            Ferme(QBalSit) ; QBalSit:=NIL ;
            {$IFDEF NOVH}
            {$ELSE}
            END Else
            BEGIN
            AttribParamsCpt1Histo(Q,D21,D22,ExecCalc.Cpt1,ExecCalc.Etab,ExecCalc.Devise,ExecCalc.Exo) ;
            QBalSit.Open ;
            PositionneTotauxHisto(QBalSit,ExecCalc.TotCpt,ExecCalc.TotCptEuro,PrepCalc.Decimale,PrepCalc.DecimaleEuro) ;
            QBalSit.Close ;
//            CalculHisto(QBalSit,D21,D22,ExecCalc.Cpt1,ExecCalc.Devise,ExecCalc.Etab,ExecCalc.Exo,ExecCalc.TotCpt,ExecCalc.TotCptEuro,PrepCalc.Decimale,PrepCalc.DecimaleEuro) ;
            END ;
            {$ENDIF NOVH}
          END Else
          BEGIN
          {$IFDEF NOVH}
          {$ELSE}
          If not EstSpecif('51185') Then
            BEGIN
          {$ENDIF NOVH}
            Q:=FabricReqCpt1(prepCalc.fb1,prepCalc.SetTyp,prepCalc.FiltreDev,prepCalc.FiltreEtab,prepCalc.FiltreExo,prepCalc.DeviseEnPivot,Multiple,prepCalc.FiltreSup,prepCalc.ChampSup,
                             ExecCalc.Cpt1,D21,D22,ExecCalc.Etab,ExecCalc.Devise,ExecCalc.Exo,ExecCalc.ValChampSup,FStBase) ;
            PositionneTotaux(Q,ExecCalc.TotCpt,ExecCalc.TotCptEuro,PrepCalc.Decimale,PrepCalc.DecimaleEuro) ;
            Ferme(Q) ; Q:=Nil ;
            {$IFDEF NOVH}
            {$ELSE}
            END Else
            BEGIN
            AttribParamsCpt1(Q,ExecCalc.Cpt1,D21,D22,ExecCalc.Etab,ExecCalc.Devise,ExecCalc.Exo,ExecCalc.ValChampSup) ;
            Q.Open ;
            PositionneTotaux(Q,ExecCalc.TotCpt,ExecCalc.TotCptEuro,PrepCalc.Decimale,PrepCalc.DecimaleEuro) ;
            Q.Close ;
            END ;
            {$ENDIF NOVH}
          END ;
       If ExecCalc.SommeTout Then
          BEGIN
          Avec6:=FALSE ; If tpCloture In PrepCalc.SetTyp Then Avec6:=TRUE ;
          Sommation(ExecCalc.TotCpt,Avec6) ; Sommation(ExecCalc.TotCptEuro,Avec6) ;
          END ;
       END ;
  UnBud : BEGIN
          FillChar(ExecCalc.TotCpt,SizeOf(ExecCalc.TotCpt),#0) ;
          FillChar(ExecCalc.TotCptEuro,SizeOf(ExecCalc.TotCptEuro),#0) ;
          {$IFDEF NOVH}
          {$ELSE}
          If  EstSpecif('51185') Then If Q=NIL Then Exit ;
          {$ENDIF NOVH}
          {D11:=0 ; D12:=0 ;} D21:=ExecCalc.Date1 ; D22:=ExecCalc.Date2 ;
          {$IFDEF NOVH}
          {$ELSE}
          If not EstSpecif('51185') Then
            BEGIN
          {$ENDIF NOVH}
            Q:=FabricReqCpt1Bud(prepCalc.fb1,prepCalc.SetTyp,prepCalc.FiltreDev,prepCalc.FiltreEtab,prepCalc.FiltreExo,prepCalc.DeviseEnPivot,Multiple,
                                ExecCalc.Cpt1,D21,D22,ExecCalc.Etab,ExecCalc.Devise,ExecCalc.Exo,FStBase) ;
            PositionneTotaux(Q,ExecCalc.TotCpt,ExecCalc.TotCptEuro,PrepCalc.Decimale,PrepCalc.DecimaleEuro) ;
            Ferme(Q) ; Q:=Nil ;
            {$IFDEF NOVH}
            {$ELSE}
            END Else
            BEGIN
            AttribParamsCpt1(Q,ExecCalc.Cpt1,D21,D22,ExecCalc.Etab,ExecCalc.Devise,ExecCalc.Exo,ExecCalc.ValChampSup) ;
            Q.Open ;
            PositionneTotauxBud(Q,ExecCalc.TotCpt,ExecCalc.TotCptEuro,PrepCalc.Decimale,PrepCalc.DecimaleEuro) ;
            Q.Close ;
            END ;
            {$ENDIF NOVH}
          END ;
  DeuxBud : BEGIN
            FillChar(ExecCalc.TotCpt,SizeOf(ExecCalc.TotCpt),#0) ;
            FillChar(ExecCalc.TotCptEuro,SizeOf(ExecCalc.TotCptEuro),#0) ;
            {$IFDEF NOVH}
            {$ELSE}
            If  EstSpecif('51185') Then If Q=NIL Then Exit ;
            {$ENDIF NOVH}
            {D11:=0 ; D12:=0 ;} D21:=ExecCalc.Date1 ; D22:=ExecCalc.Date2 ;
            {$IFDEF NOVH}
            {$ELSE}
            If not EstSpecif('51185') Then
              BEGIN
            {$ENDIF NOVH}
              Q:=FabricReqCpt2Bud(prepCalc.fb1,prepCalc.fb2,prepCalc.SetTyp,prepCalc.FiltreDev,prepCalc.FiltreEtab,prepCalc.FiltreExo,prepCalc.DeviseEnPivot,Multiple,
                                  ExecCalc.Cpt1,ExecCalc.Cpt2,D21,D22,ExecCalc.Etab,ExecCalc.Devise,ExecCalc.Exo,FStBase) ;
              PositionneTotaux2Bud(Q,ExecCalc.TotCpt,ExecCalc.TotCptEuro,PrepCalc.Decimale,PrepCalc.DecimaleEuro) ;
              Ferme(Q) ; Q:=Nil ;
              {$IFDEF NOVH}
              {$ELSE}
              END Else
              BEGIN
              AttribParamsCpt2(Q,ExecCalc.Cpt1,ExecCalc.Cpt2,ExecCalc.Date1,ExecCalc.Date2,ExecCalc.Etab,ExecCalc.Devise,ExecCalc.Exo,NULL) ;
              Q.Open ;
              PositionneTotaux2Bud(Q,ExecCalc.TotCpt,ExecCalc.TotCptEuro,PrepCalc.Decimale,PrepCalc.DecimaleEuro) ;
              Q.Close ;
              END ;
              {$ENDIF NOVH}
            END  ;
   END ; { Case }
If {EuroOk And} prepCalc.EnEuro Then ExecCalc.TotCpt:=ExecCalc.TotCptEuro ;
end ;

Function WhereCpt(fb : tFichierBase ; CptIn,CptOut : String) : String ;
Var Where,SQL : String ;
BEGIN
Where:='' ; SQL:='' ; Result:='' ;
Where:=Ent1.AnalyseCompte(CptIn,fb,False,False) ;
if Where<>'' then Sql:=Sql+' AND ('+Where+') ' ;
Where:=AnalyseCompte(CptOut,fb,True,False) ;
if Where<>'' then Sql:=Sql+' AND ('+Where+') ' ;
If SQL<>'' Then Result:=' '+SQL+' ' ;
END ;


Procedure FactoriseComboTypeRub ( C : THValComboBox ) ;
Var Q : TQuery ;
BEGIN
  Q:=OpenSql('Select CO_CODE,CO_LIBELLE from COMMUN Where CO_TYPE="RBT" And '+
             '(CO_CODE="GEN" Or CO_CODE="TIE" Or CO_CODE="G/A" Or CO_CODE="G/T" Or '+
             'CO_CODE="ANA")',True) ;
  C.Items.Clear ; C.Values.Clear ;
  While Not Q.EOF do BEGIN
    C.Values.Add(Q.Fields[0].AsString) ;
    C.Items.Add(Q.Fields[1].AsString) ;
    Q.Next ;
  END ;
  Ferme(Q) ;
END ;
////////////////////////////////////

//CETTE PARTIE A ETE DEPLACE DANS ULIBEXERCICE à L'ENLEVER

{$IFNDEF COMSX}

Function DecodeDateOle(Var St: String ; Var SD : String) : Boolean;
Var Pos1,Pos2 : Integer ;
BEGIN
SD:='' ; Result:=FALSE ;
Pos1:=Pos('(',St) ; Pos2:=Pos(')',St) ;
If (Pos1>0) And (Pos2>0) Then
  BEGIN
  SD:=Copy(St,Pos1+1,Pos2-Pos1-1) ; St[Pos1]:=' ' ; St[Pos2]:=' ' ;
  Result:=TRUE ;
  END ;
END ;

Function TraiteDate(St : String ; Var DD : TDateTime) : Integer ;
Var i : Integer ;
    St1,StD : String ;
    JJ,MM,AA : Word ;
BEGIN
Result:=-13 ; DD:=0 ; If (Length(St)<>8) And (Length(St)<>10) Then Exit ;
For i:=1 To Length(St) Do
  BEGIN
  If (St[i] in ['0'..'9','-','/'])=FALSE Then Exit ;
  If (St[i]='-') Or (St[i]='/') Then St[i]:=';' ;
  END ;
St:=St+';' ;
St1:=ReadTokenSt(St) ; JJ:=StrToInt(St1) ; AA:=0 ; MM:=0 ;
If St<>'' Then BEGIN St1:=ReadTokenSt(St) ; MM:=StrToInt(St1) ; END ;
If St<>'' Then
  BEGIN
  St1:=ReadTokenSt(St) ; AA:=StrToInt(St1) ;
  If AA<80 Then AA:=2000+AA Else If AA<=99 Then AA:=1900+AA ;
  END ;
If (JJ=0) Or (MM=0) Or (AA=0) Then Exit ; If JJ in [1..31]=FALSE Then Exit ;
If MM in [1..12]=FALSE Then Exit ;
StD:=FormatFloat('00',JJ)+'/'+FormatFloat('00',MM)+'/'+FormatFloat('00',AA) ;
If Not IsValidDate(StD) Then Exit ;
DD:=EncodeDate(AA,MM,JJ) ;
Result:=0 ;
END ;

FUNCTION QUELEXODT1(DD : TDateTime) : String ;
Var i : Integer ;
begin
{$IFDEF NOVH}
Result:=GetEnCours.Code ;
If (dd>=GetEnCours.Deb) and (dd<=GetEnCours.Fin) then Result:=GetEnCours.Code else
   If (dd>=GetSuivant.Deb) and (dd<=GetSuivant.Fin) then Result:=GetSuivant.Code Else
      If (dd>=GetPrecedent.Deb) and (dd<=GetPrecedent.Fin) then Result:=GetPrecedent.Code Else
      BEGIN
         For i:=1 To 5 Do
           BEGIN
           If (dd>=GetExoClo[i].Deb) And (dd<=GetExoClo[i].Fin)then BEGIN Result:=GetExoClo[i].Code ; Exit ; END ;
           END ;
      END ;
{$ELSE}
Result:=VH^.EnCours.Code ;
If (dd>=VH^.EnCours.Deb) and (dd<=VH^.EnCours.Fin) then Result:=VH^.EnCours.Code else
   If (dd>=VH^.Suivant.Deb) and (dd<=VH^.Suivant.Fin) then Result:=VH^.Suivant.Code Else
      If (dd>=VH^.Precedent.Deb) and (dd<=VH^.Precedent.Fin) then Result:=VH^.Precedent.Code Else
      BEGIN
         For i:=1 To 5 Do
           BEGIN
           If (dd>=VH^.ExoClo[i].Deb) And (dd<=VH^.ExoClo[i].Fin)then BEGIN Result:=VH^.ExoClo[i].Code ; Exit ; END ;
           END ;
      END ;
{$ENDIF NOVH}
end ;

procedure GetSemaine(St : String ; Var DD1,DD2 : tDateTime ; Var Err : Integer ; Var Exo : TExoDate) ;
var
  i,j,NumSemaine,NumSemaine2,Annee,PremierJourAnnee : integer ; DDT : TDateTime ;
begin
err:=0 ; DD1:=StrToDate(StDate1900) ; DD2:=StrToDate(StDate2099) ; Exo.code:='' ;
i:=pos(':',St) ;                                      if i=0 then begin Err:=-13; Exit ; end ;
system.delete(St,1,i) ;
i:=pos(':',St) ; j:=pos('-',St) ;                     if (i=0) or (j=0) then begin Err:=-13 ; Exit ; end ;
{$I-} NumSemaine:=StrToInt(copy(St,1,j-1)) ; {$I+}    if IOResult<>0 then begin Err:=-13 ; Exit ; end ;
{$I-}NumSemaine2:=StrToInt(copy(St,j+1,i-j-1)) ;{$I+} if IOResult<>0 then begin Err:=-13 ; Exit ; end ;
System.delete(St,1,i) ;
{$I-}Annee:=StrToInt(St) ;{$I+}                       if IOResult<>0 then begin Err:=-13 ; Exit ; end ;
if (Annee<0) or (Annee>9999) then begin Err:=-13 ; Exit ; end ;
if (NumSemaine<1) or (NumSemaine2<1) then begin Err:=-13 ; Exit ; end ;

if NumSemaine>NumSemaine2 then NumSemaine2:=NumSemaine2+52 ;
if Annee<100 then if annee>80 then Annee:=Annee+1900 else Annee:=Annee+2000 ;
PremierJourAnnee:=DayOfWeek(encodedate(Annee,1,1)) ;
if PremierJourAnnee>2 then DDT:=EncodeDate(Annee,1,10-PremierJourAnnee)
                      else if PremierJourAnnee=1 then DDT:=EncodeDate(Annee,1,2)
                                                 else DDT:=EncodeDate(Annee,1,1);
DD1:=DDT+(NumSemaine-1)*7 ;
DD2:=DDT+(NumSemaine2-1)*7+6 ;
if QUELEXODT1(DD1)=QUELEXODT1(DD2) then QuelDateDeExo(QUELEXODT1(DD1),Exo) else  Exo.code:='';
end ;

Function NbMoins(St : String) : Integer ;
Var i,l : Integer ;
begin
l:=0 ; Repeat i:=Pos('-',St) ; If i>0 Then BEGIN St[i]:=' ' ; Inc(l) ; END ; Until i=0 ;
Result:=-1*l ;
end ;

Function NbPer(St : String) : Integer ;
Var l,i : Integer ;
begin
l:=Length(St) ; Result:=0 ;
For i:=1 To L Do
  BEGIN
  If St[i] in ['0'..'9'] Then
     BEGIN
     If (i<L) And (St[i+1] In ['0'..'9']) Then Result:=StrToInt(Copy(St,i,2))
                                          Else Result:=StrToInt(Copy(St,i,1)) ;
     Exit ;
     END ;
  END ;
end ;

Function TraitePeriodique(NoPer,Coeff : Integer ; Exo : TExoDate ; Var DD1,DD2 : TdateTime ) : Integer ;
Var DD,D1,D2 : TdateTime ;
    ii : Integer ;
    OkOk : Boolean ;
begin
Result:=0 ;
If NoPer<>0 Then
   BEGIN
   DD1:=Exo.Deb ;
   If NoPer>1 Then DD1:=PlusMois(Exo.Deb,Coeff*(NoPer-1)) ;
   If DD1>Exo.Fin Then BEGIN Result:=-15 ; Exit ; END Else
      BEGIN
      DD2:=PlusMois(DD1,Coeff)-1 ;
      If DD2>Exo.Fin Then BEGIN Result:=-16 ; Exit ; END ;
      END ;
   END Else
   BEGIN
   DD:=V_PGI.DateEntree ; D1:=Exo.Deb ; D2:=PlusMois(DD1,Coeff)-1 ; ii:=0 ; //OkOk:=FALSE ;
   Repeat
     Inc(ii) ;
     OkOk:=(DD>=D1) And (DD<=D2) ;
     If Not OkOk Then
        BEGIN
        D1:=D2+1 ; D2:=PlusMois(D1,Coeff)-1 ;
        END ;
   Until (ii>20) Or OkOk ;
   If OkOk Then
      BEGIN
      DD1:=D1 ; DD2:=D2 ;
      If DD2>Exo.Fin Then BEGIN Result:=-16 ; Exit ; END ;
      END Else Result:=-15 ;
   END ;
end ;

function WhatDate(St : String ; Var DD1,DD2 : tDateTime ; Var Err : Integer ; Var Exo : TExoDate) : boolean;
Var OnExo,OnPeriode,OnTrimestre,OnSemestre,OnBimestre,OnQuatromestre,OnJour : Boolean ;
    DepuisDebut,JusquaFin : Boolean ;
    NoPer : Integer ;
    Decal : Integer ;
    a,m,j,aa,mm,jj : Word ;
    ii,l : Integer ; 
    MonoExo : Boolean ;
    Sd1,Sd2 : String ;
    DD3 : tDateTime ;
   Label 0 ;
{ Semaine :
W;2;2000 ou W:2-3:00-> Semaine 2 an 2000
W;13;1999 ou w;13;99 -> Semaine 13 an 1999
Renvoie : DD1,DD2 : les 2 dates de la semaine
          Exo : tExo date correspondant à la semaine
}
begin
Err:=0 ; Exo.Code:='' ;
OnExo:=FALSE ; OnPeriode:=FALSE ; OnTrimestre:=FALSE ; OnSemestre:=FALSE ; OnBimestre:=FALSE ;
OnQuatromestre:=FALSE ; DepuisDebut:=FALSE ; JusquaFin:=FALSE ; NoPer:=0 ; Decal:=0 ; OnJour:=FALSE ;
//DD1:=VH^.EnCours.Deb ; DD2:=VH^.EnCours.Fin ;
DD1 := GetEnCours.Deb ; DD2 := GetEnCours.Fin ;

If DecodeDateOLE(St,SD1) And DecodeDateOLE(St,SD2) Then
  BEGIN
  Err:=TraiteDate(Sd1,DD1) ; If Err=0 Then Err:=TraiteDate(Sd2,DD2) ;
  If Err=0 Then
    BEGIN
    If DD1>DD2 Then BEGIN DD3:=DD1 ; DD1:=DD2 ; DD2:=DD3 ; END ;
    QuelExoDate(DD1,DD2,MonoExo,Exo) ;
    If (Not MonoExo) Then Err:=-15 ;
    If (Exo.Code='') Then Err:=-6 ;
    END ;
  Goto 0 ;
  END ;
If (Pos('N',St)>0) or (Pos('n',St)>0) Then OnExo:=TRUE Else
  If (Pos('M',St)>0) Or (Pos('m',St)>0) Then OnPeriode:=TRUE Else
    If (Pos('T',St)>0) Or (Pos('t',St)>0) Then OnTrimestre:=TRUE Else
      If (Pos('B',St)>0) Or (Pos('b',St)>0) Then OnBimestre:=TRUE Else
        If (Pos('S',St)>0) Or (Pos('s',St)>0) Then OnSemestre:=TRUE Else
          If (Pos('Q',St)>0) Or (Pos('q',St)>0) Then OnQuatromestre:=TRUE Else
             If (Pos('J',St)>0) Or (Pos('j',St)>0) Then OnJour:=TRUE Else
                If (Pos('W',St)>0) Or (Pos('w',St)>0) Then BEGIN GetSemaine(St,DD1,DD2,Err,Exo) ; Goto 0 ; END Else Err:=-1 ;
If Err=0 Then
   BEGIN
   If (Pos('<',St)>0) Then DepuisDebut:=TRUE Else
      If (Pos('>',St)>0) Then JusquaFin:=TRUE ;
   If Pos('+',St)>0 Then Decal:=+1 Else If Pos('-',St)>0 Then Decal:=NbMoins(St) ;
   NoPer:=NbPer(St) ;
   END ;
If (Not (OnPeriode or OnJour)) And (DepuisDebut Or JusquaFin) Then Err:=-13 ;

//If Decal=1 Then Exo:=VH^.Suivant ;
//If Decal=0 Then Exo:=VH^.EnCours ;
If Decal = 1 Then Exo := GetSuivant ;
If Decal = 0 Then Exo := GetEnCours ;

If Decal<0 Then
BEGIN
  {$IFDEF EAGLSERVER}
  {JP 04/05/07 : ## NOVH En attendant que CA fasse la modif ; c'est juste pour pouvoir compiler}
  l:=0 ;
  For ii:=5 DownTo 1 Do
  BEGIN
    //If VH^.ExoClo[ii].Code<>'' Then Inc(l) ;
    //If l=Abs(Decal) Then Exo:=VH^.ExoClo[ii] ;
    If GetExoClo[ii].Code<>'' Then Inc(l) ;
    If l=Abs(Decal) Then Exo := GetExoClo[ii] ;
  END ;
  {$ELSE}
  if VH^.Encours.Code = '' then
  begin
    l:=0 ;
    For ii:=5 DownTo 1 Do
    BEGIN
      If GetExoClo[ii].Code<>'' Then Inc(l) ;
      If l=Abs(Decal) Then Exo := GetExoClo[ii] ;
    END ;
  end else
  // CA - 26/04/2007 pour problème ExoClo limité à 5
    QuelDateDeExo(CRelatifVersExercice ( 'N'+IntToStr(Decal)),Exo);
  {$ENDIF EAGLSERVER}
END ;

If OnJour Then
   BEGIN
   QuelExoDate(V_PGI.DateEntree,V_PGI.DateEntree,MonoExo,Exo) ;
   END ;
If Exo.Code=''  Then Err:=-6 ;
If Err=0 Then
   BEGIN
   If OnPeriode Then
      BEGIN
      If NoPer>0 Then
         BEGIN
         DD1:=Exo.Deb ;
         { FQ 15656 - CA - 22/04/2005 : pas le premier mois de l'exercice,
          on se base sur la date de début de mois (cas des exercices qui ne commencent pas au 01 }
         //If NoPer>1 Then DD1:=PlusMois(Exo.Deb,NoPer-1) ;
         If NoPer>1 Then DD1:=PlusMois(DebutDeMois(Exo.Deb),NoPer-1) ;
         If DepuisDebut Then BEGIN DD2:=FinDeMois(DD1) ; DD1:=Exo.Deb ; END Else
            If JusquaFin Then DD2:=Exo.Fin Else
               BEGIN
               DD2:=FinDeMois(DD1) ;
               END ;
         If DD1>Exo.Fin Then Err:=-17 ;
         If DD2>Exo.Fin Then Err:=-18 ;
         END Else
         BEGIN
         DD1:=V_PGI.DateEntree ;
         DecodeDate(DD1,a,m,j) ;
         DecodeDate(Exo.Deb,aa,mm,jj) ;
         DD1:=EncodeDate(aa,m,1) ;
         DD2:=FinDeMois(DD1) ;
         If DepuisDebut Then DD1:=Exo.Deb Else If JusquaFin Then DD2:=Exo.Fin ;
         END ;
      END ;
   If OnExo Then
      BEGIN
      DD1:=Exo.Deb ; DD2:=Exo.Fin ;
      END ;
   If OnTrimestre Then Err:=TraitePeriodique(NoPer,3,Exo,DD1,DD2) ;
   If OnSemestre Then Err:=TraitePeriodique(NoPer,6,Exo,DD1,DD2) ;
   If OnBimestre Then Err:=TraitePeriodique(NoPer,2,Exo,DD1,DD2) ;
   If OnQuatromestre Then Err:=TraitePeriodique(NoPer,4,Exo,DD1,DD2) ;
   If OnJour Then
      BEGIN
      DD1:=V_PGI.DateEntree ; DD2:=V_PGI.DateEntree ;
      If DepuisDebut Then DD1:= Exo.Deb Else If JusquaFin Then DD2:=Exo.Fin ;
      END ;
   END ;
0:If Err<>0 Then Result:=FALSE Else Result:=TRUE ;
end ;



{$ENDIF}


{ TSQLAnaCroise }
{b FP FQ16854}
function TSQLAnaCroise.AxeToSousPlan(NatureCpt: String): Integer;
var
  i: Integer;
begin
  Result := 0;
  for i := 1 to MaxAxe do begin
//    if FAxes[i] then
      Inc(Result);
    if NatureCpt = 'A'+IntToStr(i) then
      break;
  end;
end;

constructor TSQLAnaCroise.Create;
begin
 if GetAnaCroisaxe then LoadInfo;  // ajout me 09/10/2007 pour ne pas parcourir les axes pb lenteur
end;

destructor TSQLAnaCroise.Destroy;
begin
  inherited;
end;

function TSQLAnaCroise.GetConditionAxe(NatureCpt: String): String;
begin
  if (not GetAnaCroisaxe) or (NatureCpt = GetPremierAxe) then
    Result := 'Y_AXE = "'+NatureCpt+'"'
  else
    Result := 'Y_AXE = "'+GetPremierAxe+'" AND '+GetChampSection(NatureCpt)+' <> ""';
end;

function TSQLAnaCroise.GetChampSection(NatureCpt: String): String;
begin
  if (not GetAnaCroisaxe) or (NatureCpt = GetPremierAxe) then
    Result := 'Y_SECTION'
  else
    Result := 'Y_SOUSPLAN'+IntToStr(AxeToSousPlan(NatureCpt));
end;

function TSQLAnaCroise.GetPremierAxe: String;
begin
  Result := 'A' + IntToStr(FPremierAxe);
end;

procedure TSQLAnaCroise.LoadInfo;
var
  i: Integer;
begin
  FPremierAxe := 0;

  for i := 1 to MaxAxe do begin
    FAxes[i] := GetParamSocSecur('SO_VENTILA' + IntToStr(i), False);
    if FAxes[i] then begin
      if (FPremierAxe = 0) then FPremierAxe := i;
    end;
  end;
end;

class function TSQLAnaCroise.ChampSection(NatureCpt: String): String;
var
  Ana: TSQLAnaCroise;
begin
  Ana := TSQLAnaCroise.Create;
  try
    Result := Ana.GetChampSection(NatureCpt);
  finally
    Ana.Free;
  end;
end;

class function TSQLAnaCroise.ConditionAxe(NatureCpt: String): String;
var
  Ana: TSQLAnaCroise;
begin
  Ana := TSQLAnaCroise.Create;
  try
    Result := Ana.GetConditionAxe(NatureCpt);
  finally
    Ana.Free;
  end;
end;
{e FP FQ16854}
end.




