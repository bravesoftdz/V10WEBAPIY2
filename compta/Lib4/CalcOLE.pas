{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 03/05/2007
Modifié le ... : 03/05/2007
Description .. : - FQ 19683 - CA - 03/05/2007 - Prise en compte des 
Suite ........ : rubriques avec codes identiques
Mots clefs ... :
*****************************************************************}
unit CALCOLE;

interface

uses
  Windows,
  SysUtils,
  Classes,
  Forms,
  StdCtrls,
  Hctrls,
  Hent1,
{$IFDEF EAGLCLIENT}
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS}
    dbtables
  {$ELSE}
    uDbxDataSet
  {$ENDIF},
{$ENDIF}

{$IFDEF VER150}
  variants,
{$ENDIF}
  uTob,
  ParamSoc,
  CpteUtil, // WhatExiste
  Ent1,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ELSE}
  tCalcCum,
  {$ENDIF MODENT1}
  CRITEDT  ,UentCommun;

Const EdtPcl : Boolean = FALSE ;

Type TabloExt = Array[1..6] of Double ;
(* 1 pour SM ; 2 pour SC ; 3 pour SD ; 4 pour TC ; 5 pour TD ; 6 pour TotalRub ou QuelCpte*)

//Pour Liaison PUBLIFI
Procedure Get_CumulColl(Cpt,SDate : String ; Var SoldeD,SoldeC : Double) ;

//POUR OLE
//en MultiSociete MS
Function  Get_CumulMS(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif : String; stRegroupement : string ='') : Variant ;
Function  Get_Cumul2MS(Table,Qui1,AvecAno,Etab,Devi,SDate,Collectif : String ; Qui2,BalSit : String ; WhereSup : String ; ChampSup : String ; ValChampSup : Variant; stRegroupement : string ='') : Variant ;
//normal
Function  Get_Cumul(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif : String) : Variant ;
Function  Get_CumulPCL(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif,DeviseEnPivot,EnMonnaieOpposee : String ; LesCptes : tStringList; BalSit : string = '') : Variant ;
Function  Get_CumulDouble(Table,Qui1,Qui2,AvecAno,Etab,Devi,SDate,Collectif : String) : Variant ;
Function Get_Info(Table,Qui,Quoi,Dollar : String) : Variant ;
Function Get_Constante(CodeCons,CodeEtab,Periode : String) : Variant ;
function Get_DateCumul(St,Quoi : String) : Variant;
Function  Get_CumulCorresp(Table,CodeCoresp,Plancoresp,TypMvt,Etab,Devi,
                           SDate,TypCalcul,Signe : String) : Variant ;
Function  Get_Cumul2(Table,Qui1,AvecAno,Etab,Devi,SDate,Collectif : String ; Qui2,BalSit : String ; WhereSup : String ; ChampSup : String ; ValChampSup : Variant) : Variant ;


(*P.FUGIER 08/99 : DEBUT*)
function Get_CumulHisto(Compte, Auxiliaire, Etablissement, Devise, Exercice,
                        DateDeb, DateFin, Periode, Oppose, Options : string) : Variant ;
function Get_CumulBal(Compte, Auxiliaire, Etablissement, Devise, Exercice,
                     DateDeb, DateFin, Periode, Oppose, Options : string ;
                     AvecAno : string=''; TypeBal : string='BDS' ;
                     TypeCumul : string='') : Variant ;
(*P.FUGIER 08/99 : FIN*)


//POUR HALLEY
procedure PREEJB(fb : TFichierBase ; Var PreE,PreJ,PreB : String) ;

//MultiSociété
function GetListeMultiDossier(Groupement : string) : string;
Function GetCumulMS(Typ,Code1,Code2,AvecAno,Etab,Devi,Exo : String ;
                  var Dat11,Dat22 : TDateTime ; Collectif,Detail : Boolean ;
                  LesCompRub : TStringList ; Var TResult : TabloExt ; EnMonnaieOpposee : Boolean ;
                  CptVariant : String = '' ; TypePlan : String = '' ; TypeCumul : String = '' ; DeviseEnPivot : Boolean = FALSE; BalSit : string = '' ;
                  WhereSup : String = '' ; SDate : String =''; stRegroupement : string = '') : Double ; //XMG 26/11/03
Function GetCumul2MS(Typ,Code1,Code2,AvecAno,Etab,Devi,Exo : String ;
                  var Dat11,Dat22 : TDateTime ; Collectif,Detail : Boolean ;
                  LesCompRub : TStringList ; Var TResult : TabloExt ; EnMonnaieOpposee : Boolean ;
                  CptVariant : String = '' ; TypePlan : String = '' ; TypeCumul : String = '' ;
                  DeviseEnPivot : Boolean = FALSE; BalSit : string = '' ;
                  WhereSup : String = '' ; SDate : String =''; stRegroupement : string = '' ) : Double ; //XMG 26/11/03

Function GetCumulAstuce(Typ,Code1,Code2,AvecAno,Etab,Devi,Exo,SCollectif : String ;
                        var Dat11,Dat22 : TDateTime ; Collectif,Detail : Boolean ;
                        LesCompRub : TStringList ; Var TResult : TabloExt ; EnMonnaieOpposee : Boolean ;
                        BalSit : String ; WhereSup : String ; ChampSup : String ; ValChampSup : Variant ;
                        CptVariant : String = '' ;
                        TypePlan : String = '' ; TypeCumul : String = '' ; Sdate : String = ''; stRegroupement : string = '') : Double ; //XMG 26/11/03

Function GetCumul(Typ,Code1,Code2,AvecAno,Etab,Devi,Exo : String ;
                  var Dat11,Dat22 : TDateTime ; Collectif,Detail : Boolean ;
                  LesCompRub : TStringList ; Var TResult : TabloExt ; EnMonnaieOpposee : Boolean ;
                  CptVariant : String = '' ; TypePlan : String = '' ; TypeCumul : String = '' ; DeviseEnPivot : Boolean = FALSE; BalSit : string = '' ;
                  WhereSup : String = '' ; SDate : String =''; stBase : string = '') : Double ; //XMG 26/11/03

Function GetInfoTable(Table,ChampSelect,Lewhere,Dollar : String) : Variant ;
Function GetConstante(CodeCons,CodeEtab,Period : String ) : Variant ;
Function QuelFB(TypRub,C1,C2 : String ; Var fb1,fb2 : TFichierBase ; Var Stax : String ; OnBudget : Boolean) : Byte ;
Procedure QuelFBRub(Var fb1,fb2 : TFichierBase ; TypRub,Stax : String ; Var Stax1,Stax2 : String ;
                    Complex : TTypeCalc ; FromRubrique,OnBudget : Boolean) ;
Function FaitReqPourWhereS(fb1:TFichierBase;Ax,Cpt1,Cex1,CExo:String;D1,D2:TDateTime ;
                           SetTyp : SetttTypePiece ; OnBudget,CalculRealise,OnTableLibre : Boolean ; JalBud : String):String ;
Function FaitReqPourWhereSHisto(fb1:TFichierBase;Ax,Cpt1,Cex1,CExo:String;D1,D2:TDateTime ;
                           SetTyp : SetttTypePiece ; OnBudget,CalculRealise,OnTableLibre : Boolean ; JalBud : String):String ;
Function FaitReqPourWhereM(fb1,fb2 : TFichierBase ; Ax,Cpt1,Cex1,Cpt2,Cex2,CExo,Typrub : String ;
                          D1,D2 : TDateTime ; SetTyp : SetttTypePiece ; OnBudget,CalculRealise,bOnTableLibre : Boolean) : String ;
Procedure SQuelTyp(SetTyp : SetttTypePiece ; Var St : String) ;
Function GetCumulCorresp(Typ,CodeCoresp,Plancoresp,TypMvt,Etab,Devi,Exo : String ;
                         TypCalcul,Signe : String ; Var Dat11,Dat22 : TDateTime) : Double ;
Function CodeAbregeToComplet(Codabr : String) : String ;

function Get_RIBINFO( vStRib, vstQuoi : String ) : String ;
function DateOnly(d : TDateTime) : TDateTime ;

// Pour générateur d'état
Function CP_GetCumulCpteJal(Cpte,Jal,DateDeb,DateFin,QuelSens,QuelSigne,AvecANO : Variant) : Variant ;
Function CPGetCumulCpteJal(Cpte,Jal : String ; DateDeb,DateFin : TDateTime ; QuelSens,QuelSigne,AvecANO : String) : Double ;
Function CPGetCumulCorresp(CodeCoresp,Plancoresp,Etab,Devi,
                             SDate,TypCalcul,QuelSigne : String) : Variant ;
Function QuelTypDeSolde(StCompt,StCode : String ; Lefb : TFichierBase ; Var UneFois : Boolean ) : String ;
function GetSousSections(Axe, Sec : string) : String;

Function GET_TVATPFINFO(TVAouTPF,REGIME,Code,Quoi,Achat,RG : String ) : Variant ;

{$IFNDEF EAGLSERVER}
function CPProcCalcMul(Func,Params,WhereSQL : string ; TT : TDataset ; Total : Boolean) : string ;
{$ENDIF EAGLSERVER}


function GetCumulDC(TypePlan,Compte,Exo : string ; Date1,Date2 : TDateTime ; TypeResult : String) : Variant ;
procedure GetPeriodes(Exo : string ; D1,D2 : TDateTime ; var P1,P2 : integer ; var DD1,DD2,DF1,DF2 : TDateTime) ;
function ComplementCumul(TypePlan,Compte,Exo : string ; D1,D2 : TDateTime ; TypeResult : string) : Extended ;


implementation

(*======= PFUGIER *)
Uses 
      ULibExercice,      
      {$IFDEF MODENT1}
      CPProcMetier,
      CPProcGen,
      {$ENDIF MODENT1}
     {$IFNDEF EAGLSERVER}
     ULibWindows,
    // ZSpeed ,
     {$ELSE}
     eSession,
     {$ENDIF EAGLSERVER}
     R_DE_RUB,
     ZCumul,
     ULibEcriture,
     UtilPGI //XMG 20/04/04
(*======= PFUGIER *)
;

Function OLEErreur(i : Integer) : String ;
BEGIN
Result:='#ERREUR:' ;
Case i Of
  1  : Result:=Result+'Table non renseignée'  ;
  2  : Result:=Result+'Champ non renseigné' ;
  3  : Result:=Result+'Code recherché non renseigné' ;
  4  : Result:=Result+'Table inexistante' ;
  5  : Result:=Result+'Enregistrement non trouvé' ;
  6  : Result:=Result+'Paramètre incorrect' ;
  7  : Result:=Result+'Exercice inexistant' ;
  8  : Result:=Result+'Date de fin hors exercice' ;
  9  : Result:=Result+'Période hors exercice' ;
  10 : Result:=Result+'Rubrique inexistante' ;
  11 : Result:=Result+'Type incorrect' ;
  12 : Result:=Result+'Type de mouvement incorrect' ;
  13 : Result:=Result+'Paramètre "date" incorrect' ;
  14 : Result:=Result+'Libellé non disponible' ;
  15 : Result:=Result+'Identifiant de table libre incorrect' ;
  16 : Result:=Result+'Code non valide ou non renseigné' ;
  17 : Result:=Result+'Cumul comptable ou budgétaire non effectué' ;
  Else Result:=Result+'Non définie '+IntToStr(i) ;
  END ;
END ;

function StDateErreur(err : Integer) : Variant ;
BEGIN
Result:='' ;
Case Err Of
  -1 : Result:=OLEErreur(6) ;
  -6 : Result:=OLEErreur(7) ;
  -13: Result:=OLEErreur(13) ;
  -15,-16 : Result:=OLEErreur(8) ;
  -17,-18 : Result:=OLEErreur(9) ;
  END ;
END ;


Function ToutOk(OnExo,OnPeriode,OnTrimestre,OnSemestre,OnBimestre,OnQuatromestre : Boolean ;
                DepuisDebut,JusquaFin : Boolean ;
                NoPer : Integer ;
                Decal : Integer) : Integer ;
begin
Result:=0 ;
end ;

Function QuelFB(TypRub,C1,C2 : String ; Var fb1,fb2 : TFichierBase ; Var Stax : String ; OnBudget : Boolean) : Byte ;
Var StTemp : String ;
    fbAnal : TFichierBase ;
BEGIN
Result:=0 ;
if C1='' then Exit ;
fbAnal:=fbAxe1 ; StAx:='' ;
if (TypRub='ANA') Or (TypRub='BSE') Or (TypRub='G/A') Or (TypRub='A/G') then
   BEGIN
   if (TypRub='ANA') Or (TypRub='BSE') Or (TypRub='A/G') then StTemp:=Copy(C1,1,2) Else
      BEGIN
      if (C2='') then Exit ;
      StTemp:=Copy(C2,1,2) ;    
      END ;
   if(StTemp='A1')Or(StTemp='A2')Or(StTemp='A3')Or(StTemp='A4')Or(StTemp='A5') then Stax:=StTemp
                                                                               else Exit ;
   if StAx[1]='A' then
      BEGIN
      Case Stax[2] of
           '1' : If OnBudget Then fbAnal:=fbBudSec1 Else fbAnal:=fbAxe1 ;
           '2' : If OnBudget Then fbAnal:=fbBudSec2 Else fbAnal:=fbAxe2 ;
           '3' : If OnBudget Then fbAnal:=fbBudSec3 Else fbAnal:=fbAxe3 ;
           '4' : If OnBudget Then fbAnal:=fbBudSec4 Else fbAnal:=fbAxe4 ;
           '5' : If OnBudget Then fbAnal:=fbBudSec5 Else fbAnal:=fbAxe5 ;
        End ;
      END ;
   END ;
If OnBudget Then
   BEGIN
   If TypRub='GEN' Then fb1:=fbBudGen Else
      If TypRub='BGE' Then fb1:=fbBudGen Else
         If TypRub='ANA' Then fb1:=fbAnal Else
            If TypRub='BSE' Then fb1:=fbAnal Else
               If TypRub='G/A' Then BEGIN If OnBudget Then fb1:=fbBudGen Else fb1:=fbGene ; fb2:=fbAnal ; END Else
                  If TypRub='A/G' Then BEGIN If OnBudget Then fb2:=fbBudGen Else fb2:=fbGene ; fb1:=fbAnal ; END ;
   END Else
   BEGIN
   If TypRub='GEN' Then fb1:=fbGene Else
      If TypRub='TIE' Then fb1:=fbAux Else
         If TypRub='ANA' Then BEGIN fb1:=fbAnal ; END Else
            If TypRub='G/A' Then BEGIN fb1:=fbGene ; fb2:=fbAnal ; END Else
               If TypRub='A/G' Then BEGIN fb2:=fbGene ; fb1:=fbAnal ; END Else
                  If TypRub='G/T' Then BEGIN If OnBudget Then fb1:=fbBudGen Else fb1:=fbGene ; fb2:=fbAux ; END Else
                     If TypRub='T/G' Then BEGIN If OnBudget Then fb2:=fbBudGen Else fb2:=fbGene ; fb1:=fbAux ; END ;
   END ;
Result:=1 ;
END ;


Procedure QuelFBRub(Var fb1,fb2 : TFichierBase ; TypRub,Stax : String ; Var Stax1,Stax2 : String ;
                    Complex : TTypeCalc ; FromRubrique,OnBudget : Boolean) ;
Var fbAnal : TFichierBase ;
{
  Convention d'appel sur TypRub :
  Typ = RUB : Calcul sur rubrique définie dans Halley.
              La fiche rubrique donnera Le type de la rubrique
              avec comme identifiant de type de rubrique :
       TypRub = GEN : Calcul sur compte général
       TypRub = TIE : Calcul sur compte auxiliaire
       TypRub = ANA : Calcul sur section analytique
       TypRub = G/A : Calcul sur couple compte général/section analytique
       TypRub = A/G : Calcul sur couple section analytique/compte général
       TypRub = G/T : Calcul sur couple compte général/compte auxiliaire
       TypRub = T/G : Calcul sur couple compte auxiliaire/compte général
       Si OnBudget=TRUE, cela identifie une rubrique budgétaire

  Typ <> RUB: Calcul Sur une entité ou un couple d'entité.
  Dans ce cas :
       TypRub = GEN : Calcul sur compte général
       TypRub = TIE : Calcul sur compte auxiliaire
       TypRub = ANA : Calcul sur section analytique
       TypRub = BGE : Calcul sur compte budgétaire
       TypRub = BSE : Calcul sur section budgétaire
       TypRub = G/A : Calcul sur couple compte général/section analytique
       TypRub = A/G : Calcul sur couple section analytique/compte général
       TypRub = G/T : Calcul sur couple compte général/compte auxiliaire
       TypRub = T/G : Calcul sur couple compte auxiliaire/compte général
       Si OnBudget=TRUE, cela identifie un calcul sur entité(s) budgétaire(s)
}
BEGIN
fbAnal:=fbAxe1 ; StAx1:='' ; StAx2:='' ;
if Stax<>'' then
   BEGIN
   if StAx[1]='A' then
      BEGIN
      Case Stax[2] of
           '1' : If OnBudget Then fbAnal:=fbBudSec1 Else fbAnal:=fbAxe1 ;
           '2' : If OnBudget Then fbAnal:=fbBudSec2 Else fbAnal:=fbAxe2 ;
           '3' : If OnBudget Then fbAnal:=fbBudSec3 Else fbAnal:=fbAxe3 ;
           '4' : If OnBudget Then fbAnal:=fbBudSec4 Else fbAnal:=fbAxe4 ;
           '5' : If OnBudget Then fbAnal:=fbBudSec5 Else fbAnal:=fbAxe5 ;
        End ;
      END ;
   END ;
If OnBudget Then
   BEGIN
   If TypRub='GEN' Then fb1:=fbBudGen Else
      If TypRub='BGE' Then fb1:=fbBudGen Else
         If TypRub='ANA' Then BEGIN fb1:=fbAnal ; Stax1:=Stax ; Stax2:='' ; END Else
            If TypRub='BSE' Then BEGIN fb1:=fbAnal ; Stax1:=Stax ; Stax2:='' ; END Else
               If TypRub='G/A' Then BEGIN fb1:=fbBudGen ; fb2:=fbAnal ; Stax1:='' ; Stax2:=Stax ; END Else
                  If TypRub='A/G' Then BEGIN fb2:=fbBudGen ; fb1:=fbAnal ; Stax1:=Stax ; Stax2:='' ; END ;
   END Else
   BEGIN
   If TypRub='GEN' Then fb1:=fbGene Else
      If TypRub='TIE' Then fb1:=fbAux Else
         If TypRub='ANA' Then BEGIN fb1:=fbAnal ; Stax1:=Stax ; Stax2:='' ; END Else
            If TypRub='G/A' Then BEGIN fb1:=fbGene ; fb2:=fbAnal ; Stax1:='' ; Stax2:=Stax ; END Else
               If TypRub='A/G' Then BEGIN fb2:=fbGene ; fb1:=fbAnal ; Stax1:=Stax ; Stax2:='' ; END Else
                  If TypRub='G/T' Then BEGIN fb1:=fbGene ; fb2:=fbAux ; END Else
                     If TypRub='T/G' Then BEGIN fb2:=fbGene ; fb1:=fbAux ; END ;
   END ;
END ;

Procedure SQuelTyp(SetTyp : SetttTypePiece ; Var St : String) ;
BEGIN
St:='' ;
If tpReel In SetTyp Then St:=St+'N' ;
If tpSim In SetTyp Then St:=St+'S' ;
If tpPrev In SetTyp Then St:=St+'P' ;
If tpSitu In SetTyp Then St:=St+'U' ;
If tpRevi In SetTyp Then St:=St+'R' ;
If tpIfrs In SetTyp Then St:=St+'I' ;  // Modif IFRS 05/05/2004
If St='' Then St:='NSPURI' ;
END ;

procedure PREEJB(fb : TFichierBase ; Var PreE,PreJ,PreB : String) ;
begin
Case fb of
     fbGene :BEGIN PreE:='E_' ; PreJ:='GENERAL' ; PreB:='G_' ; END ;
     fbAux  :BEGIN PreE:='E_' ; PreJ:='AUXILIAIRE' ; PreB:='T_' ; END ;
     fbBudgen :BEGIN PreE:='BE_' ; PreJ:='BUDGENE' ; PreB:='BG_' ; END ;
     fbBudSec1..fbBudSec5 :BEGIN PreE:='BE_' ; PreJ:='BUDSECT' ; PreB:='BS_' ; END ;
     fbAxe1..fbAxe5 :BEGIN PreE:='Y_' ; PreJ:='SECTION' ; PreB:='S_' ; END ;
  End ;
end ;

Function Retypefb(fb : TFichierBase) : TFichierBase ;
BEGIN
Result:=fb ;
Case fb Of
  fbBudgen  : Result:=fbGene ;
  fbBudSec1 : Result:=fbAxe1 ;
  fbBudSec2 : Result:=fbAxe2 ;
  fbBudSec3 : Result:=fbAxe3 ;
  fbBudSec4 : Result:=fbAxe4 ;
  fbBudSec5 : Result:=fbAxe5 ;
  End ;
END ;
(*
Function CompleteCodeSection(Sect,Jal : String) : String ;
Var QLoc,QLoc1 : TQuery ;
    CodSect,CodeSp,St,StSp,StSp1 : String ;
    Deb,Lon,i : Integer ;
    TabSousPlan : TSousPlanCat ;
BEGIN
Result:=Sect ;
if (Sect='') or (Jal='') then Exit ;
FillChar(TabSousPlan,SizeOf(TabSousPlan),#0) ;
QLoc:=OpenSql('Select BJ_CATEGORIE,BJ_SOUSPLAN From BUDJAL Where BJ_BUDJAL="'+Jal+'"',True) ;
QLoc1:=OpenSql('Select BS_SECTIONRUB,BS_EXCLURUB From BUDSECT Where BS_BUDSECT="'+Sect+'"',True) ;
if QLoc.Fields[0].AsString='' then
   BEGIN
   LesSections:=QLoc1.Fields[0].AsString ;
   if QLoc1.Fields[1].AsString<>'' then LesSections:=LesSections+'<<>>'+QLoc1.Fields[1].AsString ;
   END else
   BEGIN
   TabSousPlan:=SousPlanCat(QLoc.Fields[0].AsString,True) ;
   CodSect:=QLoc1.Fields[0].AsString ; CodeSp:=QLoc.Fields[1].AsString ;
   While CodSect<>'' do
      BEGIN
      St:=ReadTokenSt(CodSect) ;
      if St<>'' then
         BEGIN
         StSp:=CodeSp ;
         for i:=1 to MaxSousPlan do
            BEGIN
            if TabSousPlan[i].Code='' then Break ;
            StSp1:=ReadTokenSt(StSp) ;
            Deb:=TabSousPlan[i].Debut ; Lon:=TabSousPlan[i].Longueur ;
            Insert(StSp1,St,Deb) ; Delete(St,Deb+Lon,Lon) ;
            END ;
         END ;
      if St<>'' then LesSections:=LesSections+St+';' ;
      END ;
   CodSect:=QLoc1.Fields[1].AsString ; CodeSp:=QLoc.Fields[1].AsString ;
   if CodSect<>'' then LesSections:=LesSections+'<<>>' ;
   While CodSect<>'' do
      BEGIN
      St:=ReadTokenSt(CodSect) ;
      if St<>'' then
         BEGIN
         StSp:=CodeSp ;
         for i:=1 to MaxSousPlan do
            BEGIN
            if TabSousPlan[i].Code='' then Break ;
            StSp1:=ReadTokenSt(StSp) ;
            Deb:=TabSousPlan[i].Debut ; Lon:=TabSousPlan[i].Longueur ;
            Insert(StSp1,St,Deb) ; Delete(St,Deb+Lon,Lon) ;
            END ;
         END ;
      if St<>'' then LesSections:=LesSections+St+';' ;
      END ;
   END ;
Ferme(QLoc) ; Ferme(QLoc1) ;
END ;
*)

Function CompleteListeSection(Sect,Jal : String) : String ;
Var QLoc : TQuery ;
    CodSect,CodeSp,St,StSp,StSp1 : String ;
    Deb,Lon,i : Integer ;
    TabSousPlan : TSousPlanCat ;
    LesSections : String ;
BEGIN
Result:=Sect ;
if (Sect='') or (Jal='') then Exit ;
FillChar(TabSousPlan,SizeOf(TabSousPlan),#0) ; LesSections:='' ;
QLoc:=OpenSql('Select BJ_CATEGORIE,BJ_SOUSPLAN From BUDJAL Where BJ_BUDJAL="'+Jal+'"',True) ;
if QLoc.Fields[0].AsString='' then BEGIN Ferme(QLoc) ; Exit ; END Else
   BEGIN
   TabSousPlan:=SousPlanCat(QLoc.Fields[0].AsString,True) ;
   CodSect:=Sect ; CodeSp:=QLoc.Fields[1].AsString ;
   While CodSect<>'' do
      BEGIN
      St:=ReadTokenSt(CodSect) ;
      if St<>'' then
         BEGIN
         StSp:=CodeSp ;
         for i:=1 to MaxSousPlan do
            BEGIN
            if TabSousPlan[i].Code='' then Break ;
            StSp1:=ReadTokenSt(StSp) ;
            Deb:=TabSousPlan[i].Debut ; Lon:=TabSousPlan[i].Longueur ;
            Insert(StSp1,St,Deb) ; Delete(St,Deb+Lon,Lon) ;
            END ;
         END ;
      if St<>'' then LesSections:=LesSections+St+';' ;
      END ;
   Result:=LesSections ;
   END ;
Ferme(QLoc) ;
END ;

Procedure LitFicheBudget(SQL,What,JalBud : String ; Var Cpt,CptEx : String) ;
Var Q : TQuery ;

BEGIN
Cpt:='' ; CptEx:='' ; Q:=OPENSQL(SQL+'"'+What+'"',TRUE) ;
If Not Q.Eof Then BEGIN Cpt:=Q.Fields[0].AsString ; CptEx:=Q.Fields[1].AsString ; END ;
If JalBud<>'' Then Cpt:=CompleteListeSection(Cpt,JalBud) ;
Ferme(Q) ;
END ;

Procedure InitReqPourWhere(TypRub,CodeRub,JalBud : String ; OnBudget,CalculRealise : Boolean ; Var lefb1,lefb2 : TFichierBase ;
                            Var Multiple : TTypeCalc ; Var Cpt1,Cpt1Ex,Cpt2,Cpt2Ex,Devi : String ; Var FiltreDev : Boolean) ;
Var Oldfb1,Oldfb2 : TFichierBase ;
    OldMultiple   : TTypeCalc ;
    Code1,Code2,SQL1,SQL2,Rien : String ;
    PosI : Integer ;
BEGIN
If Not (OnBudget) Then exit ;
If CalculRealise Then
   BEGIN
   Oldfb1:=lefb1 ; Oldfb2:=lefb2 ; OldMultiple:=Multiple ;
   Case OldMultiple Of
     UnBud   : BEGIN lefb1:=Retypefb(Oldfb1) ; Multiple:=Un ; END ;
     DeuxBud : BEGIN lefb1:=Retypefb(Oldfb1) ; lefb2:=Retypefb(Oldfb2) ; Multiple:=Deux ; END ;
     END ;
   Devi:='' ; FiltreDev:=FALSE ;
   If TypRub<>'RUB' Then
      BEGIN
      Code1:=Cpt1 ; Code2:=Cpt2 ;
      SQL1:='Select BG_COMPTERUB, BG_EXCLURUB FROM BUDGENE WHERE BG_BUDGENE=' ;
      SQL2:='Select BS_SECTIONRUB, BS_EXCLURUB FROM BUDSECT WHERE BS_BUDSECT=' ;
      Case Oldfb1 Of
        fbBudGen : LitFicheBudget(SQL1,Code1,'',Cpt1,Cpt1Ex) ;
        fbBudSec1..fbBudsec5 : BEGIN
                               If Length(Code1)>2 Then Delete(Code1,1,2) ;
                               LitFicheBudget(SQL2,Code1,JalBud,Cpt1,Cpt1Ex) ;
//                               If JalBud<>'' Then CompleteCodeSection(Sect,Jal : String ; Var LesSections : String) ;
                               END ;
        END ;
      If OldMultiple=DeuxBud Then
         BEGIN
         Case Oldfb2 Of
           fbBudGen : LitFicheBudget(SQL1,Code2,'',Cpt2,Cpt2Ex) ;
           fbBudSec1..fbBudsec5 : BEGIN
                                  If Length(Code2)>2 Then Delete(Code2,1,2) ;
                                  LitFicheBudget(SQL2,Code2,JalBud,Cpt2,Cpt2Ex) ;
                                  END ;
           END ;
         END ;
      END ;
   END Else If TypRub='RUB' Then
   BEGIN
   Cpt1:='' ; Cpt2:='' ; Cpt1Ex:='' ; Cpt2Ex:='' ;
   SQL1:='Select BG_BUDGENE, BG_EXCLURUB FROM BUDGENE WHERE BG_RUB=' ;
   SQL2:='Select BS_BUDSECT, BS_EXCLURUB FROM BUDSECT WHERE BS_RUB=' ;
   Case Multiple Of
     UnBud   : BEGIN Code1:=Copy(CodeRub,4+Length(JalBud),Length(CodeRub)-3-Length(JalBud)) ; Code2:='' ; END ;
     DeuxBud : BEGIN
               PosI:=Pos(':',CodeRub) ;
               Code1:=Copy(CodeRub,4+Length(JalBud),PosI-4-Length(JalBud)) ;
               Code2:=Copy(CodeRub,PosI+1,Length(CodeRub)-PosI) ;
               END ;
     END ;
   Case Lefb1 Of
     fbBudgen : LitFicheBudget(SQL1,Code1,'',Cpt1,Rien) ;
     fbBudSec1..fbBudSec5 : LitFicheBudget(SQL2,Code1,'',Cpt1,Rien) ;
     END ;
   If Multiple=DeuxBud Then
      BEGIN
      Case Lefb2 Of
        fbBudGen : LitFicheBudget(SQL1,Code2,'',Cpt2,Rien) ;
        fbBudSec1..fbBudsec5 : LitFicheBudget(SQL2,Code2,'',Cpt2,Rien) ;
        END ;
      END ;
   END ;
END ;

Function FaitReqPourWhereS(fb1:TFichierBase;Ax,Cpt1,Cex1,CExo:String;D1,D2:TDateTime ;
                           SetTyp : SetttTypePiece ; OnBudget,CalculRealise,OnTableLibre : Boolean ; JalBud : String):String ;
Var Where,Sql,Tri,St : String ;
    PreE,PreJ,PreB : String ;
    i : Integer ;
BEGIN
SQuelTyp(SetTyp,St) ;
Sql:='Select ' ;
PREEJB(fb1,PreE,PreJ,PreB) ;
Case fb1 of
     fbGene :BEGIN Sql:=Sql+'G_GENERAL From GENERAUX Q' ; Tri:=' Order by G_GENERAL' ; END ;
     fbAux  :BEGIN Sql:=Sql+'T_AUXILIAIRE From TIERS Q' ; Tri:=' Order by T_AUXILIAIRE' ; END ;
     fbBudgen :BEGIN Sql:=Sql+'BG_BUDGENE From BUDGENE Q'  ; Tri:=' Order by BG_BUDGENE' ; END ;
     fbBudSec1..fbBudSec5 :BEGIN Sql:=Sql+'BS_BUDSECT From BUDSECT Q'  ; Tri:=' AND BS_AXE="'+Ax+'" Order by BS_BUDSECT' ; END ;
     fbAxe1..fbAxe5 :BEGIN Sql:=Sql+'S_SECTION From SECTION Q' ; Tri:=' And S_AXE="'+Ax+'" Order by S_SECTION' ; END ;
  End ;
If OnBudget And (Not CalculRealise) Then
   BEGIN
   Case fb1 Of
     fbBudgen :
     begin
         if Length(Cpt1) < GetParamsoc('SO_LGMAXBUDGET') then
            Where:='BG_BUDGENE Like "'+Cpt1+'%" '   // fiche 19662
         else
            Where:='BG_BUDGENE="'+Cpt1+'" ' ;
      end;
     fbBudSec1..fbBudSec5 :
       BEGIN
       if ((Pos('A1',Cpt1)=1) or (Pos('A2',Cpt1)=1) or (Pos('A3',Cpt1)=1) or (Pos('A4',Cpt1)=1) or (Pos('A5',Cpt1)=1)) then Delete(Cpt1,1,2) ;
        if Length(Cpt1) < 5 then
            Where:='BS_BUDSECT Like "'+Cpt1+'%" '   // fiche 19662
        else
            Where:='BS_BUDSECT="'+Cpt1+'" ' ;
       END ;
     END ;
   Sql:=Sql+' Where '+Where+Tri ;
   END Else
   BEGIN
   i:=Length(Sql) ;
   Where:=AnalyseCompte(Cpt1,fb1,False,OnTableLibre,FALSE) ;
   if Where<>'' then Sql:=Sql+' Where '+Where ;
   Where:=AnalyseCompte(Cex1,fb1,True,FALSE,FALSE) ;
   if Where<>'' then Sql:=Sql+' And '+Where ;
   if i<>Length(Sql) then
      BEGIN
      If OnBudget And CalculRealise Then Sql:=Sql+' And '+WhatExiste(PreE,PreJ,PreB,St,False,cbUnchecked,D1,D2,CExo,1,'')+Tri
                                    Else Sql:=Sql+' And '+WhatExiste(PreE,PreJ,PreB,St,False,cbUnchecked,D1,D2,CExo,1,'')+Tri ;
      END ;
   END ;
Result:=Sql ;
END ;

Function FaitReqPourWhereSHisto(fb1:TFichierBase;Ax,Cpt1,Cex1,CExo:String;D1,D2:TDateTime ;
                           SetTyp : SetttTypePiece ; OnBudget,CalculRealise,OnTableLibre : Boolean ; JalBud : String):String ;
Var Where,Sql,Tri,St : String ;
    PreE,PreJ,PreB : String ;
//    i : Integer ;
BEGIN
  SQuelTyp(SetTyp,St) ;
  Sql:='Select ' ;
  PREEJB(fb1,PreE,PreJ,PreB) ;
  Case fb1 of
    fbGene :BEGIN Sql:=Sql+'G_GENERAL From GENERAUX Q' ; Tri:=' Order by G_GENERAL' ; END ;
    fbAux  :BEGIN Sql:=Sql+'T_AUXILIAIRE From TIERS Q' ; Tri:=' Order by T_AUXILIAIRE' ; END ;
    fbBudgen :BEGIN Sql:=Sql+'BG_BUDGENE From BUDGENE Q'  ; Tri:=' Order by BG_BUDGENE' ; END ;
    fbBudSec1..fbBudSec5 :BEGIN Sql:=Sql+'BS_BUDSECT From BUDSECT Q'  ; Tri:=' AND BS_AXE="'+Ax+'" Order by BS_BUDSECT' ; END ;
    fbAxe1..fbAxe5 :BEGIN Sql:=Sql+'S_SECTION From SECTION Q' ; Tri:=' And S_AXE="'+Ax+'" Order by S_SECTION' ; END ;
  end ;
  If OnBudget And (Not CalculRealise) Then
  begin
    Case fb1 Of
      fbBudgen : Where:='BG_BUDGENE="'+Cpt1+'" ' ;
      fbBudSec1..fbBudSec5 :
      begin
        if ((Pos('A1',Cpt1)=1) or (Pos('A2',Cpt1)=1) or (Pos('A3',Cpt1)=1) or (Pos('A4',Cpt1)=1) or (Pos('A5',Cpt1)=1)) then Delete(Cpt1,1,2) ;
        Where:='BS_BUDSECT="'+Cpt1+'" ' ;
      end;
    end;
    Sql:=Sql+' Where '+Where+Tri ;
  end else
  begin
//    i:=Length(Sql) ;
    Where:=AnalyseCompte(Cpt1,fb1,False,OnTableLibre,FALSE) ;
    if Where<>'' then Sql:=Sql+' Where '+Where ;
    Where:=AnalyseCompte(Cex1,fb1,True,FALSE,FALSE) ;
    if Where<>'' then Sql:=Sql+' And '+Where ;
{    if i<>Length(Sql) then
      Sql:=Sql+' And (EXISTS (SELECT HB_EXERCICE FROM HISTOBAL WHERE HB_EXERCICE="'+CExo+'"))'; }
  end;
  Result:=Sql ;
END ;

Function FaitReqPourWhereM(fb1,fb2 : TFichierBase ; Ax,Cpt1,Cex1,Cpt2,Cex2,CExo,Typrub : String ;
                          D1,D2 : TDateTime ; SetTyp : SetttTypePiece ; OnBudget,CalculRealise,bOnTableLibre : Boolean) : String ;
Var Where,WhereEx,Sql,Tri,St : String ;
    Ventil,PreE,PreJ1,PreJ2,PreB,Prax : String ;
    i : Integer ;
BEGIN
SQuelTyp(SetTyp,St) ;

Case fb1 Of
     fbGene : Case fb2 Of
                fbAux : BEGIN
                        Sql:='Select G_GENERAL,T_AUXILIAIRE From GENERAUX P, TIERS Q1 ';
                        Ventil:='' ; WhereEx:='P.G_GENERAL' ; PreE:='E_' ; PreJ1:='GENERAL' ;
                        PreJ2:='AUXILIAIRE'; PreB:='T_' ; Prax:='' ;
                        Tri:=' Order by G_GENERAL,T_AUXILIAIRE ' ;
                        END;
                fbBudSec1..fbBudSec5 : BEGIN
                                       Sql:='Select BG_BUDGENE, BS_BUDSECT From BUDGENE P, BUDSECT Q1 ' ;
                                       Ventil:=' ' ; //' And B_VENTILABLE'+Ax[2]+'="X" ' ;
                                       WhereEx:='P.BG_BUDGENE' ; PreE:='BE_' ; PreJ1:='BUDGENE' ;
                                       PreJ2:='BUDSECT'; PreB:='BS_' ; Prax:=' And BE_AXE="'+Ax+'" ' ;
                                       Tri:=' Order by BG_BUDGENE, BS_BUDSECT' ;
                                       END ;
                fbAxe1..fbAxe5 : BEGIN
                                 Sql:='Select G_GENERAL,S_SECTION From GENERAUX P, SECTION Q1 ';
                                 Ventil:=' And G_VENTILABLE'+Ax[2]+'="X" AND S_AXE="'+Ax+ '"'; // FQ 12447
                                 WhereEx:='P.G_GENERAL' ; PreE:='Y_' ; PreJ1:='GENERAL' ;
                                 PreJ2:='SECTION' ; PreB:='S_' ; Prax:=' And Y_AXE="'+Ax+'" ' ;
                                 Tri:=' Order by G_GENERAL,S_SECTION ' ;
                                 END ;
                END ;
     fbAux  : Case fb2 Of
                fbGene : BEGIN
                         Sql:='Select T_AUXILIAIRE,G_GENERAL From TIERS P, GENERAUX Q1 ';
                         Ventil:='' ; WhereEx:='P.T_AUXILIAIRE' ; PreE:='E_' ; PreJ1:='AUXILIAIRE' ;
                         PreJ2:='GENERAL'; PreB:='G_' ; Prax:='' ;
                         Tri:=' Order by T_AUXILIAIRE,G_GENERAL ' ;
                         END ;
                END ;
     fbBudgen : Case fb2 Of
                  fbBudSec1..fbBudSec5 : BEGIN
                                         Sql:='Select BG_BUDGENE, BS_BUDSECT From BUDGENE P, BUDSECT Q1 ' ;
                                         Ventil:=' ' ; //' And B_VENTILABLE'+Ax[2]+'="X" ' ;
                                         WhereEx:='P.BG_BUDGENE' ; PreE:='BE_' ; PreJ1:='BUDGENE' ;
                                         PreJ2:='BUDSECT'; PreB:='BS_' ; Prax:=' And BE_AXE="'+Ax+'" ' ;
                                         Tri:=' Order by BG_BUDGENE, BS_BUDSECT' ;
                                         END ;
                  END ;
     fbBudSec1..fbBudSec5 : Case fb2 Of
                                 fbBudGen : BEGIN
                                            Sql:='Select BS_BUDSECT,BG_BUDGENE From BUDSECT P, BUDGENE Q1 ' ;
                                            Ventil:=' ' ; //' And B_VENTILABLE'+Ax[2]+'="X" ' ;
                                            WhereEx:='P.BS_BUDSECT' ; PreE:='BE_' ; PreJ1:='BUDSECT' ;
                                            PreJ2:='BUDGENE' ; PreB:='BG_' ; Prax:=' And BE_AXE="'+Ax+'" ' ;
                                            Tri:=' Order by BS_BUDSECT, BG_BUDGENE ' ;
                                            END ;
                              END ;
     fbAxe1..fbAxe5 : Case fb2 Of
                        fbGene : BEGIN
                                 Sql:='Select S_SECTION,G_GENERAL From SECTION P, GENERAUX Q1 ' ;
                                 Ventil:=' And G_VENTILABLE'+Ax[2]+'="X" ' ;
                                 WhereEx:='P.S_SECTION' ; PreE:='Y_' ; PreJ1:='SECTION' ;
                                 PreJ2:='GENERAL' ; PreB:='G_' ; Prax:=' And Y_AXE="'+Ax+'" ' ;
                                 Tri:=' Order by S_SECTION,G_GENERAL ' ;
                                 END ;
                        END ;
  End ;


If OnBudget And (Not CalculRealise) Then
   BEGIN
   Where:='' ; WhereEx:='' ;
   Case fb1 Of
     fbBudgen :BEGIN Where:='BG_BUDGENE="'+Cpt1+'" AND ' ; END ;
     fbBudSec1..fbBudSec5 :
       BEGIN
       if ((Pos('A1',Cpt1)=1) or (Pos('A2',Cpt1)=1) or (Pos('A3',Cpt1)=1) or (Pos('A4',Cpt1)=1) or (Pos('A2',Cpt1)=1)) then Delete(Cpt1,1,2) ;
       Where:='BS_BUDSECT="'+Cpt1+'" AND ' ;
       END ;
     END ;
   Case fb2 Of
     fbBudgen :BEGIN WhereEx:='BG_BUDGENE="'+Cpt2+'" ' ; END ;
     fbBudSec1..fbBudSec5 :
       BEGIN
       if ((Pos('A1',Cpt2)=1) or (Pos('A2',Cpt2)=1) or (Pos('A3',Cpt2)=1) or (Pos('A4',Cpt2)=1) or (Pos('A5',Cpt2)=1)) then Delete(Cpt2,1,2) ;
       WhereEx:='BS_BUDSECT="'+Cpt2+'" ' ;
       END ;
     END ;
   Sql:=Sql+' Where '+Where+WhereEx+Tri ;
   END Else
   BEGIN
   i:=Length(Sql) ;
   Where:=AnalyseCompte(Cpt1,fb1,False,False,FALSE) ;
   if Where<>'' then Sql:=Sql+' Where '+Where ;
   Where:=AnalyseCompte(Cex1,fb1,True,False,FALSE) ;
   if Where<>'' then Sql:=Sql+' And '+Where ;
   Where:=AnalyseCompte(Cpt2,fb2,False,bOnTableLibre,FALSE) ;
   if Where<>'' then Sql:=Sql+' And '+Where ;
   Where:=AnalyseCompte(Cex2,fb2,True,bOnTableLibre,FALSE) ;
   if Where<>'' then Sql:=Sql+' And '+Where ;
   if i<>Length(Sql) then
      Sql:=Sql+' And '+WhatExisteMul(PreE,PreJ1,PreJ2,PreB,St,False,cbUnchecked,D1,D2,CExo,WhereEx,Prax,1)+Ventil+Tri ;
   END ;
Result:=Sql ;
END ;

Function TypDeSoldeMultiple(StCompt,StCode : String ; Lefb : TFichierBase) : String ;
Var St,StTemp,StC1,StC2,StBase : String ;
    i: Integer ;
BEGIN
St:=StCompt ; StBase:='' ;
While St<>'' do
  BEGIN
  StC1:='' ; StC2:='' ;
  StTemp:=ReadTokenSt(St) ;
  if StTemp='' then Continue ;
  i:=Pos(':',StTemp) ;
  if i>0 then
     BEGIN
     StC1:=Copy(StTemp,1,i-1) ; StC2:=Copy(StTemp,i+1,(Pos('(',StTemp))-(i+1)) ;
     StC1:=BourrelaDonc(StC1,Lefb) ; StC2:=BourrelaDonc(StC2,Lefb) ;
     if(StCode>=StC1) And (StCode<=StC2) then BEGIN StBase:=Copy(StTemp,Pos('(',StTemp)+1,2) ; Break ; END ;
     END else
     BEGIN
     StC1:=Copy(StTemp,1,Pos('(',StTemp)-1) ;
     if Length(StC1)<=GetInfoCpta(Lefb).Lg then
        BEGIN
        if(Pos(StC1,StCode)=1)Or(StC1=StCode) then BEGIN StBase:=Copy(StTemp,Pos('(',StTemp)+1,2) ; Break ; END ;
        END ;
     END ;
  END ;
Result:=StBase ;
END ;

Function QuelTypDeSolde(StCompt,StCode : String ; Lefb : TFichierBase ; Var UneFois : Boolean ) : String ;
Var St,StBase,St2 : String ;
    i : Integer ;
    Ok : Boolean ;
BEGIN
St:=StCompt ; Ok:=True ; Result:='' ;
i:=Pos('(',St) ;
if i>0 then BEGIN StBase:=Copy(St,i+1,2) ; Delete(St,1,i+4) ; END else Exit ;
While(Pos('(',St)>0) And Ok do
  BEGIN
  St2:=Copy(St,Pos('(',St)+1,2) ; Delete(St,1,Pos('(',St)+4) ;
  if StBase<>St2 then BEGIN Ok:=False ; UneFois:=False ; Break ; END ;
  END ;
if Ok then BEGIN Result:=StBase ; UneFois:=True ; Exit ; END ;
Result:=TypDeSoldeMultiple(StCompt,StCode,Lefb) ;
END ;

Procedure FaitResultatGlobal(Var ToTal : TabTot ; Var TabResult,TabResultCpte : TabloExt ;
                             AvecAno,St : String ; Deci : Integer ; Signe : String) ;
(* TabResult 1 pour SM ;2 pour SC ;3 pour SD ;4 pour TC ;5 pour TD ;6 pour Total rubrique*)
Var SoldeTampon,SoldeTemp,SoldeTemp2 : Double ;
BEGIN
Case Avecano[1] of
'A':BEGIN
    SoldeTampon:=Arrondi(Total[0].TotDebit-Total[0].TotCredit,Deci) ;
    SoldeTemp:=SoldeTampon ;
    TabResult[1]:=Arrondi(TabResult[1]+SoldeTampon,Deci) ;
    TabResultCpte[1]:=SoldeTampon ;
    if St='SM' then
       BEGIN
       if Signe<>'' then
          if Signe='NEG' then SoldeTampon:=SoldeTampon*(-1) ;
       TabResult[6]:=Arrondi(TabResult[6]+SoldeTampon,Deci) ;
       TabResultCpte[6]:=7 ;
       END ;
    if SoldeTemp<=0 then
       BEGIN
       TabResult[2]:=Arrondi(TabResult[2]+Abs(SoldeTemp),Deci); ;
       TabResultCpte[2]:=Abs(SoldeTemp) ;
       END ;
    if St='SC' then
       BEGIN
       TabResultCpte[6]:=6 ;
       SoldeTemp2:=SoldeTemp ;
       if SoldeTemp2<=0 then
          BEGIN
          if Signe='NEG' then SoldeTemp2:=SoldeTemp2*(-1) ;
          TabResult[6]:=Arrondi(TabResult[6]+SoldeTemp2,Deci) ;
          END ;
       END ;
    if SoldeTemp>=0 then
       BEGIN
       TabResult[3]:=Arrondi(TabResult[3]+Abs(SoldeTemp),Deci) ;
       TabResultCpte[3]:=Abs(SoldeTemp) ;
       END ;
    if St='SD' then
       BEGIN
       TabResultCpte[6]:=5 ;
       SoldeTemp2:=SoldeTemp ;
       if SoldeTemp2>=0 then
          BEGIN
          if Signe='NEG' then SoldeTemp2:=SoldeTemp2*(-1) ;
          TabResult[6]:=Arrondi(TabResult[6]+SoldeTemp2,Deci) ;
          END ;
       END ;
    TabResult[4]:=Arrondi(TabResult[4]+Total[0].TotCredit,Deci) ;
    TabResultCpte[4]:=Total[0].TotCredit ;
    if St='TC' then
       BEGIN
       TabResult[6]:=Arrondi(TabResult[6]+Total[0].TotCredit,Deci) ;
       TabResultCpte[6]:=3 ;
       END ;
    TabResult[5]:=Arrondi(TabResult[5]+Total[0].TotDebit,Deci) ;
    TabResultCpte[5]:=Total[0].TotDebit ;
    if St='TD' then
       BEGIN
       TabResult[6]:=Arrondi(TabResult[6]+Total[0].TotDebit,Deci) ;
       TabResultCpte[6]:=2 ;
       END ;
    END ;
'S':BEGIN
    SoldeTampon:=Arrondi(Total[1].TotDebit-Total[1].TotCredit,Deci) ;
    SoldeTemp:=SoldeTampon ;
    TabResult[1]:=Arrondi(TabResult[1]+SoldeTampon,Deci) ;
    TabResultCpte[1]:=SoldeTampon ;
    if St='SM' then
       BEGIN
       if Signe<>'' then
          if Signe='NEG' then SoldeTampon:=SoldeTampon*(-1) ;
       TabResult[6]:=Arrondi(TabResult[6]+SoldeTampon,Deci) ;
       TabResultCpte[6]:=7 ;
       END ;
    if SoldeTemp<=0 then
       BEGIN
       TabResult[2]:=Arrondi(TabResult[2]+Abs(SoldeTemp),Deci) ;
       TabResultCpte[2]:=Abs(SoldeTemp) ;
       END ;
    if St='SC' then
       BEGIN
       TabResultCpte[6]:=6 ;
       SoldeTemp2:=SoldeTemp ;
       if SoldeTemp2<=0 then
          BEGIN
          if Signe='NEG' then SoldeTemp2:=SoldeTemp2*(-1) ;
          TabResult[6]:=Arrondi(TabResult[6]+SoldeTemp2,Deci) ;
          END ;
       END ;
    if SoldeTemp>=0 then
       BEGIN
       TabResult[3]:=Arrondi(TabResult[3]+Abs(SoldeTemp),Deci) ;
       TabResultCpte[3]:=Abs(SoldeTemp) ;
       END ;
    if St='SD' then
       BEGIN
       TabResultCpte[6]:=5 ;
       SoldeTemp2:=SoldeTemp ;
       if SoldeTemp2>=0 then
          BEGIN
          if Signe='NEG' then SoldeTemp2:=SoldeTemp2*(-1) ;
          TabResult[6]:=Arrondi(TabResult[6]+SoldeTemp2,Deci) ;
          END ;
       END ;
    TabResult[4]:=Arrondi(TabResult[4]+Total[1].TotCredit,Deci) ;
    TabResultCpte[4]:=Total[1].TotCredit ;
    if St='TC' then
       BEGIN
       TabResult[6]:=Arrondi(TabResult[6]+Total[1].TotCredit,Deci) ;
       TabResultCpte[6]:=3 ;
       END ;
    TabResult[5]:=Arrondi(TabResult[5]+Total[1].TotDebit,Deci) ;
    TabResultCpte[5]:=Total[1].TotDebit ;
    if St='TD' then
       BEGIN
       TabResult[6]:=Arrondi(TabResult[6]+Total[1].TotDebit,Deci) ;
       TabResultCpte[6]:=2 ;
       END ;
    END ;
'T':BEGIN
    SoldeTampon:=Arrondi((Total[0].TotDebit-Total[0].TotCredit)+(Total[1].TotDebit-Total[1].TotCredit),Deci) ;
    SoldeTemp:=SoldeTampon ;
    TabResult[1]:=Arrondi(TabResult[1]+SoldeTampon,Deci) ;
    TabResultCpte[1]:=SoldeTampon ;
    if St='SM' then
       BEGIN
       if Signe<>'' then
          if Signe='NEG' then SoldeTampon:=SoldeTampon*(-1) ;
       TabResult[6]:=Arrondi(TabResult[6]+SoldeTampon,Deci) ;
       TabResultCpte[6]:=7 ;
       END ;
    if SoldeTemp<=0 then
       BEGIN
       TabResult[2]:=Arrondi(TabResult[2]+Abs(SoldeTemp),Deci) ;
       TabResultCpte[2]:=Abs(SoldeTemp) ;
       END ;
    if St='SC' then
       BEGIN
       TabResultCpte[6]:=6 ;
       SoldeTemp2:=SoldeTemp ;
       if SoldeTemp2<=0 then
          BEGIN
          if Signe='NEG' then SoldeTemp2:=SoldeTemp2*(-1) ;
          TabResult[6]:=Arrondi(TabResult[6]+SoldeTemp2,Deci) ;
          END ;
       END ;
    if SoldeTemp>=0 then
       BEGIN
       TabResult[3]:=Arrondi(TabResult[3]+Abs(SoldeTemp),Deci) ;
       TabResultCpte[3]:=Abs(SoldeTemp) ;
       END ;
    if St='SD' then
       BEGIN
       TabResultCpte[6]:=5 ;
       SoldeTemp2:=SoldeTemp ;
       if SoldeTemp2>=0 then
          BEGIN
          if Signe='NEG' then SoldeTemp2:=SoldeTemp2*(-1) ;
          TabResult[6]:=Arrondi(TabResult[6]+SoldeTemp2,Deci) ;
          END ;
       END ;
    TabResult[4]:=Arrondi(TabResult[4]+Total[0].TotCredit+Total[1].TotCredit,Deci) ;
    TabResultCpte[4]:=Arrondi(Total[0].TotCredit+Total[1].TotCredit,Deci) ;
    if St='TC' then
       BEGIN
       TabResult[6]:=Arrondi(TabResult[6]+Total[0].TotCredit+Total[1].TotCredit,Deci) ;
       TabResultCpte[6]:=3 ;
       END ;
    TabResult[5]:=Arrondi(TabResult[5]+Total[0].TotDebit+Total[1].TotDebit,Deci) ;
    TabResultCpte[5]:=Arrondi(Total[0].TotDebit+Total[1].TotDebit,Deci) ;
    if St='TD' then
       BEGIN
       TabResult[6]:=Arrondi(TabResult[6]+Total[0].TotDebit+Total[1].TotDebit,Deci) ;
       TabResultCpte[6]:=2 ;
       END ;
    END ;
End ;
END ;

Procedure FaitResultat(Var ToTal : TabTot ; Var TabResult : TabloExt ; AvecAno,St : String ; Deci : Integer ;
                       Signe : String ) ;
Var SoldeTampon : Double ;
BEGIN
//SoldeTampon:=0 ;
Case Avecano[1] of
'A':BEGIN
    if St='SM' then
       BEGIN
       SoldeTampon:=Arrondi(Total[0].TotDebit-Total[0].TotCredit,Deci) ;
       if Signe<>'' then
          if Signe='NEG' then SoldeTampon:=SoldeTampon*(-1) ;
       TabResult[1]:=TabResult[1]+SoldeTampon ;
       END Else
       if St='SC' then
          BEGIN
          SoldeTampon:=Arrondi(Total[0].TotDebit-Total[0].TotCredit,Deci) ;
          if SoldeTampon<=0 then TabResult[2]:=Arrondi(TabResult[2]+Abs(SoldeTampon),Deci) ;
          END else
          if St='SD' then
             BEGIN
             SoldeTampon:=Arrondi(Total[0].TotDebit-Total[0].TotCredit,Deci) ;
             if SoldeTampon>=0 then TabResult[3]:=Arrondi(TabResult[3]+Abs(SoldeTampon),Deci) ;
             END else
             if St='TC' then TabResult[4]:=Arrondi(TabResult[4]+Total[0].TotCredit,Deci) else
                if St='TD' then TabResult[5]:=Arrondi(TabResult[5]+Total[0].TotDebit,Deci) ;
    END ;
'S':BEGIN
    if St='SM' then
       BEGIN
       SoldeTampon:=Arrondi(Total[1].TotDebit-Total[1].TotCredit,Deci) ;
       if Signe<>'' then
          if Signe='NEG' then SoldeTampon:=SoldeTampon*(-1) ;
       TabResult[1]:=TabResult[1]+SoldeTampon ;
       END else
       if St='SC' then
          BEGIN
          SoldeTampon:=Arrondi(Total[1].TotDebit-Total[1].TotCredit,Deci) ;
          if SoldeTampon<=0 then TabResult[2]:=Arrondi(TabResult[2]+Abs(SoldeTampon),Deci) ;
          END else
          if St='SD' then
             BEGIN
             SoldeTampon:=Arrondi(Total[1].TotDebit-Total[1].TotCredit,Deci) ;
             if SoldeTampon>=0 then TabResult[3]:=Arrondi(TabResult[3]+Abs(SoldeTampon),Deci) ;
             END else
             if St='TC' then TabResult[4]:=Arrondi(TabResult[4]+Total[1].TotCredit,Deci) else
                if St='TD' then TabResult[5]:=Arrondi(TabResult[5]+Total[1].TotDebit,Deci) ;
    END ;
'T':BEGIN
    if St='SM' then
       BEGIN
       SoldeTampon:=Arrondi((Total[0].TotDebit-Total[0].TotCredit)+(Total[1].TotDebit-Total[1].TotCredit),Deci) ;
       if Signe<>'' then
          if Signe='NEG' then SoldeTampon:=SoldeTampon*(-1) ;
       TabResult[1]:=TabResult[1]+SoldeTampon ;
       END else
       if St='SC' then
          BEGIN
          SoldeTampon:=Arrondi((Total[0].TotDebit-Total[0].TotCredit)+(Total[1].TotDebit-Total[1].TotCredit),Deci) ;
          if SoldeTampon<=0 then TabResult[2]:=Arrondi(TabResult[2]+Abs(SoldeTampon),Deci) ;
          END else
          if St='SD' then
             BEGIN
             SoldeTampon:=Arrondi((Total[0].TotDebit-Total[0].TotCredit)+(Total[1].TotDebit-Total[1].TotCredit),Deci) ;
             if SoldeTampon>=0 then TabResult[3]:=Arrondi(TabResult[3]+Abs(SoldeTampon),Deci) ;
             END else
             if St='TC' then TabResult[4]:=Arrondi(TabResult[4]+(Total[0].TotCredit+Total[1].TotCredit),Deci)else
                if St='TD' then TabResult[5]:=Arrondi(TabResult[5]+(Total[0].TotDebit+Total[1].TotDebit),Deci) ;
    END ;
End ;
END ;

Procedure SommeLeResultat(Var TabResult : TabloExt) ;
Var i : Integer ;
BEGIN
for i:=1 to 5 do
   if i>1 then
     TabResult[1]:=TabResult[1]+TabResult[i] ;
END ;

Function EstCpteCollectif(Cpte : String) :Boolean ;
Var QLoc : TQuery ;
    St : String ;
BEGIN
QLoc:=OpenSql('Select G_NATUREGENE From GENERAUX Where G_GENERAL="'+Cpte+'"',True) ;
St:=QLoc.Fields[0].AsString ; Ferme(QLoc) ;
if(St='COS') Or (St='COD') Or (St='COF') Or (St='COC')then Result:=True
                                                      else Result:=False ;
END ;

Function QuelTyp(St : String ; Var SetTyp : SetttTypePiece ; Var AvecAno : String ; Var EnMonnaieOpposee : Boolean) : String ;
BEGIN
Result:='SM' ;
AvecANO:='TOU' ; SetTyp:=[] ;
If Pos('N',St)>0 Then SetTyp:=SetTyp+[tpReel] ;
If Pos('S',St)>0 Then SetTyp:=SetTyp+[tpSim] ;
If Pos('P',St)>0 Then SetTyp:=SetTyp+[tpPrev] ;
If Pos('U',St)>0 Then SetTyp:=SetTyp+[tpSitu] ;
If Pos('R',St)>0 Then SetTyp:=SetTyp+[tpRevi] ;
If Pos('I',St)>0 Then SetTyp:=SetTyp+[tpIfrs] ;  // Modif IFRS 05/05/2004
If SetTyp=[] Then SetTyp:=[tpReel] ;
If Pos('-',St)>0 Then AvecAno:='SAN' ;
If Pos('#',St)>0 Then AvecAno:='ANO' ;
If Pos('D',St)>0 Then Result:='TD' ;
If Pos('C',St)>0 Then Result:='TC' ;
If Pos('/',St)>0 Then Result:='SD' ;
If Pos('\',St)>0 Then Result:='SC' ;
If Pos('O',St)>0 Then If Not EnMonnaieOpposee Then EnMonnaieOpposee:=TRUE ;
END ;

Function DecimaleDev(Devi : String) : Byte ;
Var QDev : TQuery ;
BEGIN
if Devi=V_PGI.DevisePivot then
   Result:=V_PGI.OkdecV else
   BEGIN
   if Devi<>'' then
      BEGIN
      QDev:=OpenSql('Select D_DECIMALE From DEVISE Where D_DEVISE="'+Devi+'"',True) ;
      if Not QDev.Eof then Result:= QDev.Fields[0].AsInteger
                      else Result:= V_PGI.OkdecV ;
      Ferme(QDev) ;
      END else Result:= V_PGI.OkdecV ;
   END ;
END ;

Procedure DechiffreTyp(Var Typ : String ; Var OnBudget,CalculRealise : Boolean) ;
(*
'GENERAUX')          -> 'GEN' ;
'TIERS')             -> 'TIE' ;
'SECTION')           -> 'ANA' ;
'RUBRIQUE')          -> 'RUB' ;
'BUDGENE:B')         -> 'BGE' ;
'BUDSECT:B')         -> 'BSE' ;
'BUDGENE')           -> 'BGE' ;
'BUDSECT')           -> 'BSE' ;
'RUBBUD:B')          -> 'RUB' ;
'RUBBUD:R')          -> 'BUDREA' ;
'BUDGENE:R')         -> 'BGEREA' ;
'BUDSECT:R')         -> 'BSEREA' ;
'GENERAUX/TIERS')    -> 'G/T' ;
'TIERS/GENERAUX')    -> 'T/G' ;
'GENERAUX/SECTION')  -> 'G/A' ;
'SECTION/GENERAUX')  -> 'A/G' ;
'BUDSECT/BUDGENE:B') -> 'BUDGET:A/G' ;
'BUDGENE/BUDSECT:B') -> 'BUDGET:G/A' ;
'BUDSECT/BUDGENE:R') -> 'BUDGET:A/GREA' ;
'BUDGENE/BUDSECT:R') -> 'BUDGET:G/AREA' ;
*)
//Var i : Integer ;
BEGIN
OnBudget:=FALSE ; CalculRealise:=FALSE ;
If Typ='RUB' Then Exit ;
If (Typ='BUDREA') Or (Typ='RUBREA')Then
   BEGIN
   OnBudget:=TRUE ; CalculRealise:=TRUE ;
   Typ:='RUB' ;
   Exit ; END ;
If Pos('BUDGET:',Typ)<>0 Then
   BEGIN
   OnBudget:=TRUE ;
   If Pos('REA',Typ)<>0 Then CalculRealise:=TRUE ;
   Typ:=Copy(Typ,8,3) ;
   END Else
   BEGIN
   If (Typ='BGE') Or (Typ='BSE') Then OnBudget:=TRUE ;
   If (Typ='BGEREA') Or (Typ='BSEREA') Then BEGIN Typ:=Copy(Typ,1,3) ; OnBudget:=TRUE ; CalculRealise:=TRUE ; END ;
   If Typ='RUBBUD' Then BEGIN OnBudget:=TRUE ; Typ:='RUB' ; END ;
   END ;
END ;

Function CaVaLeFaire(Exo : String ; D1,D2 : TDateTime ; Var DateBut,DDeb,DFin : TDateTime) : Boolean ;
Var DTemp : TDateTime ;
BEGIN
Result:=False ; DateBut:=0 ; DDeb:=0 ; DFin:=0 ;
if Exo=GetEncours.Code then DateBut:=GetEnCours.DateButoirBud else
   if Exo=GetSuivant.Code then DateBut:=GetSuivant.DateButoirBud else Exit ;
if D1>DateBut then Exit ;
if DebutDeMois(D1)<>D1 then Exit else DDeb:=D1 ;
if D2>=DateBut then DFin:=DateBut else
   BEGIN
   if FindeMois(D2)<>D2 then
      BEGIN
      DTemp:=FinDeMois(PlusMois(D2,-1)) ;
      if DTemp>DDeb then DFin:=DTemp else Exit ;
      END else DFin:=D2 ;
   END ;
Result:=True ;
END ;

Procedure QuelPeriode(DatBut,Dat11,Dat22 : TdateTime ; Exo : String ; Var P1,P2,P11,P12 : Integer) ;
Var UnExo : TExoDate ;
    NbPer,i : Integer ;
    D1,D2 : TDateTime ;
BEGIN
P1:=0 ; P2:=0 ; P11:=0 ; P12:=0 ;
UnExo.Code:=Exo ; RempliExoDate(UnExo) ; NbPer:=UnExo.NombrePeriode ;
if NbPer<=12 then
   BEGIN
   D1:=UnExo.Deb ; D2:=FinDeMois(D1) ;
   for i:=1 to NbPer do
      BEGIN
      if Dat11=D1 then P1:=i ;
      if (D2=DatBut) or (D2=Dat22) then P2:=i ;
      if (P1<>0) And (P2<>0) then Break ;
      D1:=PlusMois(D1,1) ; D2:=FinDeMois(D1) ;
      END ;
   END else
   BEGIN
   D1:=UnExo.Deb ; D2:=FinDeMois(D1) ;
   for i:=1 to 12 do
      BEGIN
      if Dat11=D1 then P1:=i ;
      if (D2=DatBut) or (D2=Dat22) then P2:=i ;
      if (P1<>0) And (P2<>0) then Break ;
      D1:=PlusMois(D1,1) ; D2:=FinDeMois(D1) ;
      END ;
   if (P1<>0) And (P2=0) then P2:=12 ;
   D1:=UnExo.Deb ; D1:=PlusMois(D1,12) ; D2:=FinDeMois(D1) ;
   for i:=13 to NbPer do
      BEGIN
      if Dat11=D1 then P11:=i-12 ;
      if (D2=DatBut) or (D2=Dat22) then P12:=i-12 ;
      if P12<>0 then Break ;
      D1:=PlusMois(D1,1) ; D2:=FinDeMois(D1) ;
      END ;
   if (P11=0) And (P12<>0) then P11:=1 ;
   END ;
END ;

Function FaitRequeteCumul(OnRea : Boolean ;Typ,Exo,Code,Code1,Etab,Devi,St : String) : String ;
Var Sql,St1 : String ;
    UnTyp : String ;
BEGIN
if OnRea then Untyp:='BRE' else UnTyp:='BUD' ;
if Typ<>'' then UnTyp:=Typ ;
Sql:='Select * From CUMULS Where CU_TYPE="'+UnTyp+'" And ' ;
if Code1='' then Sql:=Sql+'CU_COMPTE1="'+Code+'" ' else Sql:=Sql+'CU_COMPTE1>="'+Code+'" And CU_COMPTE1<="'+Code1+'" ' ;
if Exo<>''  then Sql:=Sql+'And CU_EXERCICE="'+Exo+'" ' ;
if Etab<>'' then Sql:=Sql+'And CU_ETABLISSEMENT="'+Etab+'" ' ;
if Devi<>'' then Sql:=Sql+'And CU_DEVQTE="'+Devi+'" ' ;
St1:='' ;
if Pos('N',St)>0 then St1:=St1+'CU_QUALIFPIECE="N" Or ' ;
if Pos('S',St)>0 then St1:=St1+'CU_QUALIFPIECE="S" Or ' ;
if Pos('P',St)>0 then St1:=St1+'CU_QUALIFPIECE="P" Or ' ;
if Pos('U',St)>0 then St1:=St1+'CU_QUALIFPIECE="U" Or ' ;
if Pos('R',St)>0 then St1:=St1+'CU_QUALIFPIECE="R" Or ' ;
if Pos('I',St)>0 then St1:=St1+'CU_QUALIFPIECE="I" Or ' ;  // Modif IFRS 05/05/2004
if St1<>'' then BEGIN Delete(St1,Length(St1)-3,3) ; St1:='('+St1+')' ; Sql:=Sql+' And '+St1 ; END ;
Result:=Sql ;
END ;

Function GetCumulDuCumul(OnRea : Boolean ; DatBut,Dat11,Dat22 : TDateTime ; Exo,Code1,St,Etab,Devi : String ; Var TResult : TabloExt) : Double ;
Var Sql : String ;
    QLoc : TQuery ;
    P1,P2,P11,P12,i : Integer ;
    Debi,Cred : Double ;
BEGIN
//Result:=0 ;
QuelPeriode(DatBut,Dat11,Dat22,Exo,P1,P2,P11,P12) ;
Sql:=FaitRequeteCumul(OnRea,'',Exo,Code1,'',Etab,Devi,St) ;
QLoc:=OpenSql(Sql,True) ; Debi:=0 ; Cred:=0 ;
While Not QLoc.Eof do
    BEGIN
    if QLoc.FindField('CU_SUITE').AsString='-' then
       for i:=P1 to P2 do
           BEGIN
           Debi:=Debi+QLoc.FindField('CU_DEBIT'+IntToStr(i)).AsFloat ;
           Cred:=Cred+QLoc.FindField('CU_CREDIT'+IntToStr(i)).AsFloat ;
           END ;
    if QLoc.FindField('CU_SUITE').AsString='X' then
       for i:=P11 to P12 do
           BEGIN
           Debi:=Debi+QLoc.FindField('CU_DEBIT'+IntToStr(i)).AsFloat ;
           Cred:=Cred+QLoc.FindField('CU_CREDIT'+IntToStr(i)).AsFloat ;
           END ;
    QLoc.Next ;
    END ;
TResult[1]:=Debi ; TResult[2]:=Cred ;
Result:=Debi-Cred ; Ferme(QLoc) ;
END ;

Function DateRubriTotOk(Typ,Exo : String ; D1,D2 : TDateTime ; Var DateBut : TDateTime) : Boolean ;
BEGIN
Result:=False ; DateBut:=0 ;
if Exo=GetEncours.Code then
   BEGIN
   Case Typ[3] of
        'B' : DateBut:=GetEnCours.DateButoirBudgete ;
        'G' : DateBut:=GetEnCours.DateButoirRub ;
        'R' : DateBut:=GetEnCours.DateButoirBud ;
     End ;
   END else
   if Exo=GetSuivant.Code then
      BEGIN
      Case Typ[3] of
          'B' : DateBut:=GetSuivant.DateButoirBudgete ;
          'G' : DateBut:=GetSuivant.DateButoirRub ;
          'R' : DateBut:=GetSuivant.DateButoirBud ;
        End ;      END else Exit ;
if D1>DateBut then Exit ;
if DebutDeMois(D1)<>D1 then Exit ;
if D2>DateBut then Exit ;
if FindeMois(D2)<>D2 then Exit ;
Result:=True ;
END ;

Function CodeAbregeToComplet(Codabr : String) : String ;
Var QLoc : TQuery ;
BEGIN
QLoc:=OpenSql('Select RB_RUBRIQUE From RUBRIQUE Where RB_CODEABREGE="'+Codabr+'"',True) ;
Result:=QLoc.Fields[0].AsString ;
Ferme(QLoc) ;
END ;

Procedure ChercheDebitCredit(P1,P2,P11,P12 : Integer ; Cpt1,TypCum,Exo,Etab,Devi,AvecAno : String ; Var Debit,Credit : Double) ;
Var St,StR,Sql : String ;
    QLoc : TQuery ;
    Ratio : Double ;
    Plus,Moins : Boolean ;
    StCod,StCod1 : String ;
    i : Integer ;
BEGIN
While Cpt1<>'' do
  BEGIN
  St:=ReadTokenSt(Cpt1) ;
  if St='' then Continue ;
  Ratio:=0 ; StCod1:='' ; Plus:=False ; Moins:=False ;
  if Pos('}',St)=1 then Break ;
  if Pos('+',St)=1 then Plus:=True else
     if Pos('-',St)=1 then Moins:=True ;
  if Pos('{',St)=1 then St:=Copy(St,2,Length(St)) ;
  if Pos('!',St)>0 then
     BEGIN
     if Plus or Moins then StCod:=CodeAbregeToComplet(Copy(St,3,Length(St)))
                      else StCod:=CodeAbregeToComplet(Copy(St,2,Length(St))) ;
     END else
     BEGIN
     if Plus or Moins then StCod:=Copy(St,3,Length(St))
                      else StCod:=Copy(St,2,Length(St)) ;
     END ;
  if Pos('<',Cpt1)=1 then
     BEGIN
     St:=ReadTokenSt(Cpt1) ;
     if Pos('!',St)>0 then StCod1:=CodeAbregeToComplet(Copy(St,4,Length(St)))
                      else StCod1:= Copy(St,4,Length(St)) ;
     END ;
  if Pos('[',Cpt1)=1 then
     BEGIN
     StR:=ReadTokenSt(Cpt1) ;
     Ratio:=StrToInt(Copy(StR,Pos('[',StR)+1,Pos('%',StR)-2)) ;
     END ;
  Sql:=FaitRequeteCumul(True,TypCum,Exo,StCod,StCod1,Etab,Devi,AvecAno) ;
  QLoc:=OpenSql(Sql,True) ;
  While Not QLoc.Eof do
    BEGIN
    if QLoc.FindField('CU_SUITE').AsString='-' then
       for i:=P1 to P2 do
           BEGIN
           Debit:=Debit+QLoc.FindField('CU_DEBIT'+IntToStr(i)).AsFloat ;
           Credit:=Credit+QLoc.FindField('CU_CREDIT'+IntToStr(i)).AsFloat ;
           END ;
    if QLoc.FindField('CU_SUITE').AsString='X' then
       for i:=P11 to P12 do
           BEGIN
           Debit:=Debit+QLoc.FindField('CU_DEBIT'+IntToStr(i)).AsFloat ;
           Credit:=Credit+QLoc.FindField('CU_CREDIT'+IntToStr(i)).AsFloat ;
           END ;
    QLoc.Next ;
    END ;
  Ferme(QLoc) ;
  if Ratio<>0 then
     BEGIN
     if Debit<>0 then Debit:=Debit+(Debit/Ratio) ;
     if Credit<>0 then Credit:=Credit+(Credit/Ratio) ;
     END ;
  END ;
END ;

Function GetCumulSurTotRub(Typ,Code1,AvecAno,Etab,Devi,Exo : String ; Dat11,Dat22 : TDateTime ; Var TResult : TabloExt) : Double ;
Var TypCum,Cpt1,Cpt1Ex : String ;
    P1,P2,P11,P12 : Integer ;
    DatBut : TDateTime ;
    QLoc : TQuery ;
    Debit,Credit,DebitEx,CreditEx : Double ;
BEGIN
Result:=0 ;
{$IFDEF NOVH}
Exit;
{$ELSE}
if Not VH^.OnCumEdt then BEGIN Result:=0 ; V_PGI.STOLEErr:=OLEErreur(17) ; Exit ; END ;
if Not DateRubriTotOk(Typ,Exo,Dat11,Dat22,DatBut) then BEGIN Result:=0 ; V_PGI.STOLEErr:=OLEErreur(13) ; Exit ; END ;
QuelPeriode(DatBut,Dat11,Dat22,Exo,P1,P2,P11,P12) ;
QLoc:=OpenSql('Select RB_COMPTE1,RB_EXCLUSION1 From Rubrique Where RB_RUBRIQUE="'+Code1+'" And RB_TYPERUB="TOT"',True) ;
if QLoc.Eof then BEGIN Ferme(QLoc) ; V_PGI.STOLEErr:=OLEErreur(10) ; Exit ; END ;
Cpt1:=QLoc.Fields[0].AsString ; Cpt1Ex:=QLoc.Fields[1].AsString ; Ferme(QLoc) ;
if (Pos('(',Cpt1)<>2) and (Pos('!',Cpt1)<>2) and (Pos('#',Cpt1)<>2) then BEGIN V_PGI.STOLEErr:=OLEErreur(16) ; Exit ; END ;
TypCum:='' ;
if Typ='TOB' then TypCum:='BUD' else
   if Typ='TOR' then TypCum:='BRE' else
      if Typ='TOG' then TypCum:='RUB' ;
Debit:=0 ; Credit:=0 ; DebitEx:=0 ; CreditEx:=0 ;
ChercheDebitCredit(P1,P2,P11,P12,Cpt1,TypCum,Exo,Etab,Devi,AvecAno,Debit,Credit) ;
ChercheDebitCredit(P1,P2,P11,P12,Cpt1Ex,TypCum,Exo,Etab,Devi,AvecAno,DebitEx,CreditEx) ;
if DebitEx<>0 then Debit:=Debit-DebitEx ;
if CreditEx<>0 then Credit:=Credit-CreditEx ;
TResult[1]:=Debit ; TResult[2]:=Credit ;
Result:=Debit-Credit ;
{$ENDIF NOVH}
END ;

Function CptEstVide(St : String) : Boolean ;
Var St1 : String ;
BEGIN
Result:=TRUE ; If St='' Then Exit ;
St1:=ReadTokenSt(St) ;
While St1<>'' Do
  BEGIN
  If Trim(St1)<>'' Then BEGIN Result:=FALSE ;Exit ; END ;
  St1:=ReadTokenSt(St) ;
  END ;
END ;

Procedure RetoucheWhereSup(Var WhereSup : String) ;
BEGIN
WhereSup:=FindEtReplace(WhereSup,'''','"',TRUE) ;
END ;


// récupère la liste des dossiers pour un groupement donné
function GetListeMultiDossier(Groupement : string) : string;
var L : TStringList;
    st,base,dossier : string;
    idx : integer;
begin
result := '';
{$IFNDEF EAGLSERVER}
idx := V_PGI.RequeteMultiDossier.IndexOf(Groupement);
{$ELSE}
idx := LookupCurrentSession.RequeteMultiDossier.IndexOf(Groupement);
{$ENDIF}
if idx=-1 then exit;// big pb !!
L := TStringList.Create;
L.Text := RechDom('YYMULTIDOSSIER',Groupement,true);
if L.Count=2 then
  begin
  result := '';
  st := L[0]; //Liste des dossiers
  while st<>'' do
    begin
    base := ReadTokenSt(st);
    dossier := ReadTokenPipe(base,'|');
    result := result + base + ';';
    end;
  end;
FreeAndNil(L);
end;

Function GetCumulMS(Typ,Code1,Code2,AvecAno,Etab,Devi,Exo : String ;
                  var Dat11,Dat22 : TDateTime ; Collectif,Detail : Boolean ;
                  LesCompRub : TStringList ; Var TResult : TabloExt ; EnMonnaieOpposee : Boolean ;
                  CptVariant : String = '' ; TypePlan : String = '' ; TypeCumul : String = '' ;
                  DeviseEnPivot : Boolean = FALSE; BalSit : string = '' ;
                  WhereSup : String = '' ; SDate : String =''; stRegroupement : string = '' ) : Double ; //XMG 26/11/03
var
    stListeBases : string;
    stBase : string ;
    valCumul : double;
begin
//Rempli la liste des sociétés
stListeBases := trim(GetListeMultiDossier(stRegroupement));
valCumul := 0;
try
  while stListeBases>'' do
  begin
    stBase := ReadTokenSt(stListeBases) ;
    valCumul := valCumul + getcumul(Typ,Code1,Code2,AvecAno,Etab,Devi,Exo,Dat11,Dat22,Collectif,Detail,
                    LesCompRub,TResult,EnMonnaieOpposee,
                    CptVariant,TypePlan,TypeCumul,
                    DeviseEnPivot,BalSit,
                    WhereSup,SDate,stBase);
  end ; //WHILE REGROUPEMENT
  result := ValCumul ;
except
  result := 0 ;
end;
end;

Function GetCumul2MS(Typ,Code1,Code2,AvecAno,Etab,Devi,Exo : String ;
                  var Dat11,Dat22 : TDateTime ; Collectif,Detail : Boolean ;
                  LesCompRub : TStringList ; Var TResult : TabloExt ; EnMonnaieOpposee : Boolean ;
                  CptVariant : String = '' ; TypePlan : String = '' ; TypeCumul : String = '' ;
                  DeviseEnPivot : Boolean = FALSE; BalSit : string = '' ;
                  WhereSup : String = '' ; SDate : String =''; stRegroupement : string = '' ) : Double ; //XMG 26/11/03
var
    stListeBases : string;
    stBase : string ;
    valCumul : double;
begin
//Rempli la liste des sociétés
stListeBases := trim(GetListeMultiDossier(stRegroupement));
valCumul := 0;
try
  while stListeBases>'' do
  begin
    stBase := ReadTokenSt(stListeBases) ;
    valCumul := valCumul + getcumul(Typ,Code1,Code2,AvecAno,Etab,Devi,Exo,Dat11,Dat22,Collectif,Detail,
                    LesCompRub,TResult,EnMonnaieOpposee,
                    CptVariant,TypePlan,TypeCumul,
                    DeviseEnPivot,BalSit,
                    WhereSup,SDate,stBase);
  end ; //WHILE REGROUPEMENT
  result := ValCumul ;
except
  result := 0;
end;

end;


Function GetCumul(Typ,Code1,Code2,AvecAno,Etab,Devi,Exo : String ;
                  var Dat11,Dat22 : TDateTime ; Collectif,Detail : Boolean ;
                  LesCompRub : TStringList ; Var TResult : TabloExt ; EnMonnaieOpposee : Boolean ;
                  CptVariant : String = '' ; TypePlan : String = '' ; TypeCumul : String = '' ;
                  DeviseEnPivot : Boolean = FALSE; BalSit : string = '' ;
                  WhereSup : String = '' ; SDate : String =''; stBase : string = '' ) : Double ; //XMG 26/11/03

Var QRub,QTotCol : TQuery ;
    Sql,S1,Compt1,Compt2,JalBud : String ;
    Multiple : TTypeCalc ;
    TypRub,LAxe,Cpt1,Cpt2,Cpt1Ex,Cpt2Ex,Signe,St,St1 : String ;
    Laxe1,Laxe2 : String ;
    Lefb1,Lefb2 : TFichierBase ;
    FilDev,FilExo,FilEtab,JustUneFois : Boolean ;
    ToTal,TotCptMvt : TabTot ;
    ToTalEURO,TotCptMvtEURO : TabTot ;
    TabResult,TabResultCpte : TabloExt ;
    SetTyp : SetttTypePiece ;
    i,DecDev : Integer ;
    GCalcul : TGCalculCum ;
    OnBudget : Boolean ;
    CalculRealise : Boolean ;
    OnTableLibre : Boolean ;
    ForceModeCalcul : String ;
{$IFNDEF EAGLSERVER}
    j,k : Integer;
    DatBut,D1,D2 : TDateTime ;
    szTemp,szTemp2,szTemp3 : String;

{$ENDIF EAGLSERVER}
{
  Convention d'appel sur Typ :
  Typ = RUB : Calcul sur rubrique définie dans Halley. Seul Code1 représente le code de la
              rubrique à charger. La fiche rubrique donnera Le type de la rubrique
       avec comme identifiant de type de rubrique
       RB_TYPRUB = GEN : Calcul sur compte général
       RB_TYPRUB = TIE : Calcul sur compte auxiliaire
       RB_TYPRUB = ANA : Calcul sur section analytique
       RB_TYPRUB = G/A : Calcul sur couple compte général/section analytique
       RB_TYPRUB = A/G : Calcul sur couple section analytique/compte général
       RB_TYPRUB = G/T : Calcul sur couple compte général/compte auxiliaire
       RB_TYPRUB = T/G : Calcul sur couple compte auxiliaire/compte général
       Si RB_NATRUB='BUD', cela identifie une rubrique budgétaire

    Typ = RUBREA : Calcul sur rubrique budgétaire définie dans Halley pour calcul du réalisé. Seul Code1 représente le code de la
                   rubrique à charger. La fiche rubrique donnera Le type de la rubrique
       avec comme identifiant de type de rubrique
       RB_TYPRUB = GEN : Calcul sur compte général
       RB_TYPRUB = ANA : Calcul sur section analytique
       RB_TYPRUB = G/A : Calcul sur couple compte général/section analytique
       RB_TYPRUB = A/G : Calcul sur couple section analytique/compte général
       RB_NATRUB='BUD' est théoriquement obligatoire

  Typ <> RUB et RUBREA : Calcul Sur une entité ou un couple d'entité.
  Dans ce cas :
       Typ = GEN : Calcul sur compte général
       Typ = TIE : Calcul sur compte auxiliaire
       Typ = ANA : Calcul sur section analytique
       Typ = BGE : Calcul sur compte budgétaire
       Typ = BSE : Calcul sur section budgétaire
       Typ = BGEREA : Calcul sur compte budgétaire pour le réalisé
       Typ = BSEREA : Calcul sur section budgétaire pour le réalisé
       Typ = G/A : Calcul sur couple compte général/section analytique
       Typ = A/G : Calcul sur couple section analytique/compte général
       Typ = G/T : Calcul sur couple compte général/compte auxiliaire
       Typ = T/G : Calcul sur couple compte auxiliaire/compte général
       Typ = BUDGET:G/A : Calcul sur couple compte budgétaire/section budgétaire
       Typ = BUDGET:A/G : Calcul sur couple section budgétaire/compte budgétaire
       Typ = BUDGET:BGE : Calcul sur compte budgétaire
       Typ = BUDGET:BSE : Calcul sur section budgétaire
       Typ = BUDGET:G/AREA : Calcul sur couple compte budgétaire/section budgétaire pour le réalisé
       Typ = BUDGET:A/GREA : Calcul sur couple section budgétaire/compte budgétaire pour le réalisé
       Typ = BUDGET:BGEREA : Calcul sur compte budgétaire pour le réalisé
       Typ = BUDGET:BSEREA : Calcul sur section budgétaire pour le réalisé
       Typ = TOG : Rubrique de totalisation. Calcul sur rubrique comptable
       Typ = TOB : Rubrique de totalisation. Calcul sur rubrique budgétaire
       Typ = TOR : Rubrique de totalisation. Calcul sur rubrique budgétaire réalisé
}
BEGIN


if (Typ='RUB') and (RubriqueDeRubrique.Indexof(Code1)>-1) then Begin //XVI 24/02/2005
   with RubriqueDeRubrique do begin
     PrmAvecAno:=AvecAno ;
     PrmEtab:=Etab ;
     PrmDevi:=Devi ;
     PrmSDate:=SDate ;
     PrmCollectif:=copy('X-',1+ord(Collectif),1) ;
     Result:=GereRDR(Code1) ;
   End ;
   exit ;
End ;
OnTableLibre:= False; RetoucheWhereSup(WhereSup) ;
EnMonnaieOpposee:={EuroOk And} EnMonnaieOpposee ;
if (Typ='TOG') or (Typ='TOB') or (Typ='TOR')then
   BEGIN
   Fillchar(TResult,SizeOf(TResult),#0) ;
   Result:=GetCumulSurTotRub(Typ,Code1,AvecAno,Etab,Devi,Exo,Dat11,Dat22,TResult) ;
   Exit ;
   END ;
OnBudget:=FALSE ; CalculRealise:=FALSE ;
DechiffreTyp(Typ,OnBudget,CalculRealise) ;
V_PGI.StOLEErr:='' ;
// GP 29/04/98 If Detail Then Fillchar(TResult,SizeOf(TResult),#0) ;
Fillchar(TResult,SizeOf(TResult),#0) ;
Result:=0 ; QRub:=NIL ; QTotCol:=NIL ; {V_PGI.OnCumEdt:=TRUE ;} Signe:='' ;
Fillchar(TabResult,Sizeof(TabResult),0) ;
if (Typ='GEN')Or(Typ='TIE')Or(Typ='ANA')Or(Typ='G/A')Or(Typ='G/T')Or(Typ='RUB') Or
   (Typ='BGE') Or (Typ='BSE') Or (Typ='A/G') Or (Typ='T/G') then
  BEGIN
  FilEtab:=(Etab<>'') ; FilDev:=(Devi<>'') ; FilExo:=(Exo<>'') ;
  END else BEGIN V_PGI.STOLEErr:=OLEErreur(11) ; Exit ; END ;
{ détermine l'axe et les fb utilisé pour GCALC Si calcul sur autre chose qu'une rubrique :}
if Typ<>'RUB' then
   if QuelFb(Typ,Code1,Code2,Lefb1,Lefb2,Laxe,OnBudget)=0 then Exit ;

St:=AvecAno ; ForceModeCalcul:=QuelTyp(St,SetTyp,AvecAno,EnMonnaieOpposee) ;
DecDev:=DecimaleDev(Devi) ;

if Typ='RUB' then
   BEGIN
   if Code1='' then BEGIN V_PGI.STOLEErr:=OLEErreur(16) ; Exit ; END ;
   OnBudget:=FALSE ;
//   QRub:=OpenSql( 'Select * From Rubrique Where RB_RUBRIQUE="'+Code1+'"',True,-1,'',False,stBase) ;
// FQ 19683 - CA - 03/05/2007 - Prise en compte des rubriques avec codes identiques
   QRub:=OpenSql( 'Select * From Rubrique Where RB_RUBRIQUE="'+Code1+'"'+
    ' AND RB_DATEVALIDITE>="'+USDateTime(iDate2099)+'"'+
    ' AND (RB_PREDEFINI<>"DOS" OR (RB_PREDEFINI="DOS" AND RB_NODOSSIER="'+V_PGI.NoDossier+'"))'+
    ' ORDER BY RB_NODOSSIER DESC ,RB_PREDEFINI DESC,RB_DATEVALIDITE ASC ' ,True,-1,'',False,stBase) ;
   if QRub.Eof then BEGIN Ferme(QRub) ; V_PGI.STOLEErr:=OLEErreur(10) ; Exit ; END ;
   TypRub:=QRub.FindField('RB_TYPERUB').AsString ;
   Laxe:=QRub.FindField('RB_AXE').AsString ;
   Cpt1:=QRub.FindField('RB_COMPTE1').AsString ;
   Cpt2:=QRub.FindField('RB_COMPTE2').AsString ;
   Cpt1Ex:=QRub.FindField('RB_EXCLUSION1').AsString ;
   Cpt2Ex:=QRub.FindField('RB_EXCLUSION2').AsString ;
   Signe:=QRub.FindField('RB_SIGNERUB').AsString ;
   OnBudget:=QRub.FindField('RB_NATRUB').AsString='BUD' ;
   OnTableLibre:=QRub.FindField('RB_TABLELIBRE').AsString='X' ;
   JalBud:=QRub.FindField('RB_BUDJAL').AsString ;
   Ferme(QRub) ;
(*======= PFUGIER *)


  {$IFNDEF EAGLSERVER}
  // Compte de résultat et SIG analytique
  if VH^.bAnalytique then begin
    TypRub := 'G/A';
    if Copy(CptVariant,1,1) = #22 then begin
      delete(CptVariant,1,1);
      Code2 := ReadTokenSt(CptVariant);
      Laxe := Code2;
      Cpt2 := ReadTokenSt(CptVariant);
      Cpt2Ex:= ReadTokenSt(CptVariant);
      OnTableLibre := True;
    end
    else begin
      Code2 := ReadTokenSt(CptVariant);
      Laxe := Code2;
    end;
  end;

   if OnBudget And (LesCompRub=Nil) And VH^.OnCumEdt then
      if CaVaLeFaire(Exo,Dat11,Dat22,DatBut,D1,D2) then
      BEGIN
        Result:=GetCumulDuCumul(CalculRealise,DatBut,D1,D2,Exo,Code1,St,Etab,Devi,TResult) ;
        if D2<=DatBut then
        BEGIN
          if D2=Dat22 then Exit
                      else Dat11:=DebutDeMois(PlusMois(D2,1)) ;
        END
        else
        Dat11:=DebutDeMois(PlusMois(D2,1)) ;
       END ;
  {$ENDIF EAGLSERVER}

   if(TypRub='A/G') or (TypRub='G/A') or (TypRub='G/T') or (TypRub='T/G') then
      BEGIN
      If OnBudget Then Multiple:=DeuxBud Else Multiple:=Deux ;
      END else If OnBudget Then Multiple:=UnBud Else Multiple:=Un ;
   If (CptVariant<>'') And (Multiple in [Deux,DeuxBud]) And ((CptEstVide(Cpt1) And CptEstVide(Cpt2))=FALSE) Then
     BEGIN
     If CptEstVide(Cpt1) Then Cpt1:=CptVariant ;
     If CptEstVide(Cpt2) Then Cpt2:=CptVariant ;
     END ;
    if Length(Cpt1)<=1 then BEGIN V_PGI.STOLEErr:=OLEErreur(16) ; Exit ; END ;
   QuelFBRub(Lefb1,Lefb2,TypRub,Laxe,Laxe1,Laxe2,Multiple,TRUE,OnBudget) ;
   InitReqPourWhere(Typ,Code1,JalBud,OnBudget,CalculRealise,lefb1,lefb2,Multiple,Cpt1,Cpt1Ex,Cpt2,Cpt2Ex,Devi,FilDev) ;

(*======= PFUGIER *)
(*======= PFUGIER *)

// GP le 30/06/2008 report correction manon
// GCalcul:=TGCalculCum.Create(Multiple,Lefb1,Lefb2,SetTyp,FilDev,FilEtab,FilExo,DeviseEnPivot,EnMonnaieOpposee,DecDev,V_PGI.OkDecE,False,BalSit,'',WhereSup,stBase) ;
                                                                                                                                                 // Fiche 22397
   GCalcul:=TGCalculCum.Create(Multiple,Lefb1,Lefb2,SetTyp,FilDev,FilEtab,FilExo,DeviseEnPivot,EnMonnaieOpposee,DecDev,V_PGI.OkDecE,False,BalSit,WhereSup,SDate,stBase) ;
   GCalcul.InitCalcul('','',Laxe1,Laxe2,Devi,Etab,Exo,Dat11,Dat22,True,BalSit<>'') ;
   if Collectif then QTotCol:=PrepareTotCptSolde(SetTyp,FilDev,FilEtab,FilExo,FALSE) ;
   Case Multiple Of
     Deux     : Sql:=FaitReqPourWhereM(Lefb1,Lefb2,LAxe,Cpt1,Cpt1Ex,Cpt2,Cpt2Ex,Exo,TypRub,Dat11,Dat22,SetTyp,OnBudget,CalculRealise,OnTableLibre) ;
     DeuxBud  : Sql:=FaitReqPourWhereM(Lefb1,Lefb2,LAxe,Cpt1,Cpt1Ex,Cpt2,Cpt2Ex,Exo,TypRub,Dat11,Dat22,SetTyp,OnBudget,CalculRealise,OnTableLibre) ;
//     Un       : Sql:=FaitReqPourWhereS(Lefb1,LAxe,Cpt1,Cpt1Ex,Exo,Dat11,Dat22,SetTyp,OnBudget,CalculRealise,OnTableLibre,JalBud) ;
     Un       :
               begin
                if BalSit <> '' then Sql := FaitReqPourWhereSHisto(Lefb1,LAxe,Cpt1,Cpt1Ex,Exo,Dat11,Dat22,SetTyp,OnBudget,CalculRealise,OnTableLibre,JalBud)
                else Sql:=FaitReqPourWhereS(Lefb1,LAxe,Cpt1,Cpt1Ex,Exo,Dat11,Dat22,SetTyp,OnBudget,CalculRealise,OnTableLibre,JalBud) ;
               end;
     UnBud    : Sql:=FaitReqPourWhereS(Lefb1,LAxe,Cpt1,Cpt1Ex,Exo,Dat11,Dat22,SetTyp,OnBudget,CalculRealise,OnTableLibre,JalBud) ;
   END ;

{$IFNDEF EAGLSERVER}
   // Pour la sélection dans l'analytique d'une fourchette de section
   if VH^.bAnalytique and (CptVariant <> ';') then begin
     if (Pos('(  S_SECTION ',SQL) <> 0) then begin
       J := Pos('(  S_SECTION ',SQL);
       szTemp := copy(SQL,J+13,length(SQL)-J+1);
       K := Pos('Or  S_SECTION ',szTemp);
       szTemp2 := copy(szTemp,K+14,length(szTemp)-K+1);
       szTemp3 := copy(szTemp,1,K-1);
       if ((k<>0) and (pos('LIKE',UpperCase(szTemp3))=0) and (pos('LIKE',UpperCase(szTemp2))=0)) then
          sql := copy(sql,1,J+12)+'>'+szTemp3+'AND S_SECTION <'+szTemp2;
     end;
   end;
{$ENDIF EAGLSERVER}

   QRub:=OpenSql( Sql,True,-1,'',False,stBase) ;
   JustUneFois:=False ;
   While Not QRub.Eof do
      BEGIN
      Compt1:=QRub.Fields[0].AsString ;
      if (Multiple=Deux) Or (Multiple=DeuxBud) then Compt2:=QRub.Fields[1].AsString
                                               else Compt2:='' ;
      If OnBudget And (Not CalculRealise) Then
         BEGIN
         S1:='SM' ; JustUneFois:=TRUE ;
         END Else
         BEGIN
         if(Not JustUneFois) then
            BEGIN
            S1:=QuelTypDeSolde(Cpt1,Compt1,Lefb1,JustUneFois) ;
            If (S1='') And (Multiple in [Deux,DeuxBud]) Then S1:=QuelTypDeSolde(Cpt2,Compt2,Lefb2,JustUneFois) ;
            if S1='' then Break ;
            END ;
         END ;
      if (Multiple=Deux) Or (Multiple=DeuxBud)then
         GCalcul.ReInitCalcul(Compt1,Compt2,Dat11,Dat22,'',BalSit<>'')
      else BEGIN
           if (TypRub='GEN') And Collectif And EstCpteCollectif(Compt1) then
               ExecuteTotCptSolde(QTotCol,Compt1,Dat11,Dat22,Devi,Etab,Exo,ToTal,TotCptMvt,ToTalEURO,TotCptMvtEURO,True,True,DecDev,V_PGI.OkDecE,EnMonnaieOpposee,FALSE,SetTyp)
           else
               GCalcul.ReInitCalcul(Compt1,'',Dat11,Dat22,'',BalSit<>'') ;
           END ;
      GCalcul.Calcul ; Total:=GCalcul.ExecCalc.TotCpt ;
      if Detail then
         BEGIN
// GP le 7/1/99 FaitResultat(Total,TabResult,AvecAno,S1,DecDev,Signe) ;
         FillChar(TabResultCpte,SizeOf(TabResultCpte),0) ;
         FaitResultatGlobal(Total,TResult,TabResultCpte,AvecAno,S1,DecDev,Signe) ;
         END else
         BEGIN
         FillChar(TabResultCpte,SizeOf(TabResultCpte),0) ;
         FaitResultatGlobal(Total,TResult,TabResultCpte,AvecAno,S1,DecDev,Signe) ;
         END ;
      if LesCompRub<>Nil then
         BEGIN
         if Not ((Multiple=Deux) Or (multiple=DeuxBud)) then St1:=Compt1 else St1:=Compt1+';'+Compt2 ;
         for i:=1 to 6 do
            St1:=St1+':'+FloatToStr(TabResultCpte[i]) ;
         If GridPCL Then St1:=St1+':'+Signe ;
         LesCompRub.Add(St1) ;
         END ;
      QRub.Next ;
      END ;
   Ferme(QRub) ;
// GP le 7/1/99   if Detail then SommeLeResultat(TabResult) ;
   GCalcul.Free ; If QTotCol<>NIL Then QTotCol.Free ;
   Result:=TResult[6] ;
   END else
   if(Typ='GEN')Or(Typ='TIE')Or(Typ='ANA')Or(Typ='G/A')Or(Typ='G/T') Or (Typ='A/G')Or(Typ='T/G') Or (Typ='BSE') Or (Typ='BGE') then
      BEGIN
      if(Typ='G/A')Or(Typ='G/T')Or(Typ='A/G')Or(Typ='T/G') then BEGIN If OnBudget Then Multiple:=DeuxBud Else Multiple:=Deux ; END Else
         If (Typ='BSE') Or (Typ='BGE') Then Multiple:=UnBud Else If OnBudget Then Multiple:=UnBud Else Multiple:=Un ;
      Cpt1:=Code1 ; Cpt2:=Code2 ; Cpt1Ex:='' ; Cpt2Ex:='' ;
      If OnBudget Then JalBud:=Devi Else JalBud:='' ;
      InitReqPourWhere(Typ,'',JalBud,OnBudget,CalculRealise,lefb1,lefb2,Multiple,Cpt1,Cpt1Ex,Cpt2,Cpt2Ex,Devi,FilDev) ;
// GP le 30/06/2008 report correction manon
//    GCalcul:=TGCalculCum.Create(Multiple,Lefb1,Lefb2,SetTyp,FilDev,FilEtab,FilExo,DeviseEnPivot,EnMonnaieOpposee,DecDev,V_PGI.OkDecE,False,BalSit,'',WhereSup,stBase) ;
//Fiche 22397
      GCalcul:=TGCalculCum.Create(Multiple,Lefb1,Lefb2,SetTyp,FilDev,FilEtab,FilExo,DeviseEnPivot,EnMonnaieOpposee,DecDev,V_PGI.OkDecE,False,BalSit,WhereSup,SDate,stBase) ;
      GCalcul.InitCalcul('','',Laxe1,Laxe2,Devi,Etab,Exo,Dat11,Dat22,True,BalSit<>'') ;
      if Collectif then QTotCol:=PrepareTotCptSolde(SetTyp,FilDev,FilEtab,FilExo,FALSE) ;
      Case Multiple Of
        Deux,DeuxBud : Sql:=FaitReqPourWhereM(Lefb1,Lefb2,LAxe,Cpt1,Cpt1Ex,Cpt2,Cpt2Ex,Exo,Typ,Dat11,Dat22,SetTyp,OnBudget,CalculRealise,OnTableLibre) ;
//        Un,UnBud     : Sql:=FaitReqPourWhereS(Lefb1,LAxe,Cpt1,Cpt1Ex,Exo,Dat11,Dat22,SetTyp,OnBudget,CalculRealise,FALSE,'') ;
      Un,UnBud     :
        begin
          if BalSit <> '' then Sql := FaitReqPourWhereSHisto(Lefb1,LAxe,Cpt1,Cpt1Ex,Exo,Dat11,Dat22,SetTyp,OnBudget,CalculRealise,FALSE,'')
          else Sql:=FaitReqPourWhereS(Lefb1,LAxe,Cpt1,Cpt1Ex,Exo,Dat11,Dat22,SetTyp,OnBudget,CalculRealise,FALSE,'') ;
        end;
        END ;
      QRub:=OpenSql(Sql,True,-1,'',False,stBase) ;
      While Not QRub.Eof do
         BEGIN
         Compt1:=QRub.Fields[0].AsString ;
         Case Multiple Of
           Deux,DeuxBud :Compt2:=QRub.Fields[1].AsString ;
           else Compt2:='' ;
           END ;
         Case Multiple Of
           Deux,DeuxBud : GCalcul.ReInitCalcul(Compt1,Compt2,Dat11,Dat22,'',BalSit<>'')
           else BEGIN
                if Collectif And (Typ='GEN') And EstCpteCollectif(Compt1) then
                    ExecuteTotCptSolde(QTotCol,Compt1,Dat11,Dat22,Devi,Etab,Exo,ToTal,TotCptMvt,ToTalEURO,TotCptMvtEURO,True,True,DecDev,V_PGI.OkDecE,EnMonnaieOpposee,FALSE,SetTyp)
                else
                    GCalcul.ReInitCalcul(Compt1,'',Dat11,Dat22,'',BalSit<>'') ;
                END ;
           END ;
         GCalcul.Calcul ; Total:=GCalcul.ExecCalc.TotCpt ;

//         FaitResultat(Total,TabResult,AvecAno,ForceModeCalcul,DecDev,Signe) ;

         if Detail or (not(OnBudget And CalculRealise)) then
            BEGIN
            FaitResultat(Total,TabResult,AvecAno,ForceModeCalcul,DecDev,Signe) ;
            TResult:=TabResult ;
            END else
            BEGIN
            FillChar(TabResultCpte,SizeOf(TabResultCpte),0) ;
            FaitResultatGlobal(Total,TResult,TabResultCpte,AvecAno,S1,DecDev,Signe) ;
            if LesCompRub<>Nil then
               BEGIN
               if Not ((Multiple=Deux) Or (multiple=DeuxBud)) then St1:=Compt1 else St1:=Compt1+';'+Compt2 ;
               for i:=1 to 6 do
                  St1:=St1+':'+FloatToStr(TabResultCpte[i]) ;
(* 1 pour SM ; 2 pour SC ; 3 pour SD ; 4 pour TC ; 5 pour TD ; 6 pour TotalRub ou QuelCpte*)
               LesCompRub.Add(St1) ;
               END ;
            TabResult:=TResult ;
            END ;
         QRub.Next ;
         END ;
      Ferme(QRub) ;
      GCalcul.Free ; If QTotCol<>NIL Then QTotCol.Free ;
      If ForceModeCalcul='SD' Then Result:=TabResult[3] Else
        If ForceModeCalcul='SC' Then Result:=TabResult[2] Else
          If ForceModeCalcul='TD' Then Result:=TabResult[5] Else
            If ForceModeCalcul='TC' Then Result:=TabResult[4] Else
              Result:=TabResult[1] ;
      END else
      BEGIN
      Result:=0 ; V_PGI.STOLEErr:=OLEErreur(11) ; Exit ;
      END ;
(*
Result:=TabResult[1] ;
If (Typ='RUB') And Detail Then
   BEGIN
   TResult[1]:=0 ; TResult[2]:=0 ;
   If Signe='POS' Then
      BEGIN
      If Result>0 Then TResult[1]:=Abs(Result) Else If Result<0 Then TResult[2]:=Abs(Result) ;
      END Else
      BEGIN
      If Result<0 Then TResult[1]:=Abs(Result) Else If Result>0 Then TResult[2]:=Abs(Result) ;
      END ;
   END ;
*)


END ;


Function GetConstante(CodeCons,CodeEtab,Period : String ) : Variant ;
Var CodeExo,LeSql,Champ,Suite : String ;
    QLoc : TQuery ;
    Num,i,j : Byte ;
    Err : Integer ;
    DD1,DD2 : TDateTime ;
    Exo : TExoDate ;
    Pm,Pa,NbM : Word ;
BEGIN
Result:='' ;
if (CodeCons='') Or (CodeEtab='') then BEGIN Result:=OLEErreur(3) ; Exit ; END ;
if Pos('$',CodeCons)>0 then
   BEGIN
   Champ:=Copy(CodeCons,2,3) ;
   LeSql:='Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="CON" And CC_CODE="'+Champ+'"' ;
   END else
   BEGIN
   i:=Pos('-',Period) ; j:=Pos('+',Period) ;
   if (i>0) then Num:=StrToInt(Copy(Period,2,i-2)) else
      if (j>0) then Num:=StrToInt(Copy(Period,2,j-2)) else
         Num:=StrToInt(Copy(Period,2,2)) ;
   if Not(Num in[1..24]) then BEGIN Result:=OLEErreur(6) ; Exit ; END ;
   CodeExo:=GetEncours.Code ;
   if (i>0) Or (j>0) then
      BEGIN
      if WhatDate(Period,DD1,DD2,Err,Exo) then CodeExo:=Exo.Code
                                          else BEGIN Result:=StDateErreur(Err) ; Exit ; END ;
      END else Exo.Code:=CodeExo ;
   if Exo.Code=GetPrecedent.Code then BEGIN Exo.Deb:=GetPrecedent.Deb ; Exo.Fin:=GetPrecedent.Fin ; END else
   if Exo.Code=GetEnCours.Code then BEGIN Exo.Deb:=GetEncours.Deb ; Exo.Fin:=GetEncours.Fin ; END else
   if Exo.Code=GetSuivant.Code then BEGIN Exo.Deb:=GetSuivant.Deb ; Exo.Fin:=GetSuivant.Fin ; END else
      BEGIN
      LeSql:='SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+Exo.Code+'"' ;
      QLoc:=OpenSql(LeSql,True) ;
      Exo.Deb:=QLoc.Fields[0].AsDateTime ; Exo.Fin:=QLoc.Fields[1].AsDateTime ;
      Ferme(QLoc) ;
      END ;
   NombrePerExo(Exo,Pm,Pa,NbM) ;
   if NbM<=12 then BEGIN Champ:='CU_DEBIT'+IntToStr(Num) ; Suite:='-' ; END else
      if Num<=12 then BEGIN Champ:='CU_DEBIT'+IntToStr(Num) ; Suite:='X' ; END
                 else BEGIN Champ:='CU_DEBIT'+IntToStr(Num-12) ; Suite:='-' ; END ;
   LeSql:='Select '+Champ+' From CUMULS Where CU_TYPE="CON" And CU_COMPTE1="'+CodeCons+'" '+
          'And CU_EXERCICE="'+CodeExo+'" And CU_ETABLISSEMENT="'+CodeEtab+'" And CU_SUITE="'+Suite+'"' ;
   END ;
QLoc:=OpenSql(LeSql,True) ;
if Not QLoc.Eof then Result:=QLoc.Fields[0].AsVariant  else Result:=OLEErreur(5) ;
Ferme(QLoc) ;
END ;

Function GetInfoTable(Table,ChampSelect,Lewhere,Dollar : String) : Variant ;
Var Q : TQuery ;
    Sql,St1,StChampAxe,AndWhere : String ;
    Axe : String ;
    OkLib : Boolean ;
    fb : TFichierBase ;
    OnTableLibre : Boolean ;
    i : Integer ;
    QuelleTable : String ;
    tz : TZoomTable ;
BEGIN
Result:='' ; St1:='' ; Axe:='' ; fb:=fbGene ; StChampAxe:='' ; AndWhere:='' ;
if(Table='') then BEGIN Result:=OLEErreur(1) ; Exit ; END ;
if(ChampSelect='') then BEGIN Result:=OLEErreur(2) ; Exit ; END ;
if(Lewhere='') then BEGIN Result:=OLEErreur(3) ; Exit ; END ;
OnTableLibre:=FALSE ; i:=Pos('+',Table) ;
If i>0 Then BEGIN OnTableLibre:=TRUE ; Delete(Table,i,1) ; END ;
If OnTableLibre Then
   BEGIN
   QuelleTable:=Copy(ChampSelect,1,3) ;
   ChampSelect:=Copy(ChampSelect,5,Length(ChampSelect)-3) ;
   If (Table='GENERAUX') Or (Table='TIERS') Or (Table='SECTION') Or (Table='BUDSECT') Or (Table='BUDGENE') Or (Table='BUDGENE') //XMG 26/11/03
      Then LeWhere:=BourreLaDoncSurLaTable(QuelleTable,ChampSelect)
      Else BEGIN Result:=OLEErreur(4) ; Exit ; END ;
   tz:=NatureToTz(QuelleTable) ; if tz=tzRien Then BEGIN Result:=OLEErreur(15) ; Exit ; END ;
   St1:=tzToChampNature(tz,TRUE) ;
   END Else
   BEGIN
   If Table='GENERAUX' Then BEGIN St1:='G_GENERAL' ; fb:=fbGene ; LeWhere:=BourreLaDonc(leWhere,fb) ; END Else
      If Table='TIERS' Then BEGIN St1:='T_AUXILIAIRE' ; fb:=fbAux ; LeWhere:=BourreLaDonc(leWhere,fb) ; END Else
         If Table='SECTION' Then
            BEGIN
            St1:='S_SECTION' ; Axe:=Copy(LeWhere,1,2) ;
            LeWhere:=Copy(LeWhere,3,Length(LeWhere)-2) ; fb:=AxeTofb(Axe) ;
            LeWhere:=BourreLaDonc(leWhere,fb) ;
            END Else
            If Table='BUDSECT' Then
               BEGIN
               St1:='BS_BUDSECT' ; Axe:=Copy(LeWhere,1,2) ;
               LeWhere:=Copy(LeWhere,3,Length(LeWhere)-2) ; fb:=AxeTofbBud(Axe) ;
               LeWhere:=BourreLaDonc(leWhere,fb) ;
               END Else
            If Table='STRUCRSE' Then
               BEGIN
               St1:='SS_SOUSSECTION' ; Axe:=Copy(LeWhere,1,2) ;
               LeWhere:=Copy(LeWhere,3,Length(LeWhere)-2) ; fb:=AxeTofbBud(Axe) ;
               StChampAxe:='SS_AXE' ;
               END Else
            If Table='SSSTRUCR' Then
               BEGIN
               St1:='PS_SOUSSECTION' ; Axe:=Copy(LeWhere,1,2) ;
               LeWhere:=Copy(LeWhere,3,Length(LeWhere)-2) ; fb:=AxeTofbBud(Axe) ;
               AndWhere:=LeWhere ;
               LeWhere:=ReadTokenst(AndWhere) ;
               AndWhere:=' AND PS_CODE="'+AndWhere+'" ' ;
               StChampAxe:='PS_AXE' ;
               END Else
               If Table='BUDGENE' Then BEGIN St1:='BG_BUDGENE' ; fb:=fbBudGen ; END Else
                  If Table='JOURNAL' Then St1:='J_JOURNAL' Else
                     If Table='BUDJAL' Then St1:='BJ_BUDJAL' Else
                        If Table='RUBRIQUE' Then St1:='RB_RUBRIQUE' Else
                           If Table='RUBBUD' Then BEGIN Table:='RUBRIQUE' ; St1:='RB_RUBRIQUE' ; END Else
                              If Table='SOCIETE' Then BEGIN
                              // ajout me 07/09/2005 fiche 13482 ST1:='SO_SOCIETE' ; LeWhere:=V_PGI.CodeSociete
                              ST1 := 'SOC_NOM';  LeWhere := ChampSelect;

                              END Else
                                 If Table='ETABLISS' Then St1:='ET_ETABLISSEMENT' Else
                                    If Table='ECRITURE' Then begin
                                       AndWhere:=LeWhere ;
                                       LeWhere:='' ;
                                    End else
                                    BEGIN
                                    V_PGI.STOLEErr:=OLEErreur(4) ;
                                    Result:=OLEErreur(4) ;
                                    Exit ;
                                    END ;
   END ;
if Table='SOCIETE' then   // ajout me 07/09/2005 fiche 13482
   Sql :=  'SELECT SOC_DATA FROM PARAMSOC WHERE '
else
   Sql :='Select '+ChampSelect+' From '+Table+' Where ' ;
if (trim(St1)<>'') and (trim(LeWhere)<>'') then SQL:=SQL+St1+'="'+Lewhere+'"' ; //XMG 24/02/2005
If Fb in [fbAxe1..fbAxe5] Then
   BEGIN
   If StChampAxe='' Then Sql:=Sql+'AND S_AXE="'+AXE+'"'
                    Else Sql:=Sql+'AND '+StChampAxe+'="'+AXE+'"' ;
   END ;
If Fb in [fbBudSec1..fbBudSec5] Then
   BEGIN
   If StChampAxe='' Then Sql:=Sql+'AND BS_AXE="'+AXE+'"'
                    Else Sql:=Sql+'AND '+StChampAxe+'="'+AXE+'"' ;
   END ;
If AndWhere<>'' Then Sql:=Sql+AndWhere;
Q:=OpenSQL(Sql,True);
if Q.Fields[0].Datatype<>ftString then Result:=0 ;
If Not Q.Eof Then
   BEGIN
   Result:=Q.Fields[0].AsVariant ;
   If Dollar='$' Then
      BEGIN
      OkLib:=((Q.Fields[0].DataType=ftString) (* ajout me and (Length(Q.Fields[0].AsString)=4)*)) ;
      If OkLib Then
         BEGIN
         Result:=RechDom(Get_Join(ChampSelect),Q.Fields[0].AsVariant,FALSE) ;
           If Result='Error' Then
           begin
            V_PGI.STOLEErr:=OLEErreur(14);
            Result:=OLEErreur(14) ;
           end;
         END ;
      END ;
   END Else
   BEGIN
   V_PGI.STOLEErr:=OLEErreur(5) ;
   Result:=OLEErreur(5) ;
   END ;
Ferme(Q);
END ;

Procedure VerifCollectif(Var Collectif : String) ;
BEGIN
If (Collectif<>'X') And (Collectif<>'') Then Collectif:='' ;
END ;

Procedure NomTable(Var Table : String ; Var collectif : String ; Var CptVariant : String ; Var TypePlan,TypeCumul : String) ;
Var Pos1,Pos2 : Integer ;
    St : String ;
(*
Type plan (HB_TYPEBAL) : BDS N0A N1A N1C
Type cumul (HB_PLAN) : BAL AUX

Si compte
'GENERAUX:BDS§BAL'
*)
BEGIN
TypePlan:='' ; TypeCumul:='' ;
Pos1:=Pos('(',Table) ; Pos2:=Pos(')',Table) ; CptVariant:='' ;
If (Pos1>0) And (Pos2>0) Then
  BEGIN
  St:=Table ;
  Table:=Copy(St,1,Pos1-1)+Copy(St,Pos2+1,Length(St)-Pos2) ;
  CptVariant:=Copy(St,Pos1+1,Pos2-Pos1-1) ;
  END ;
Pos1:=Pos(':',Table) ; Pos2:=Pos('§',Table) ;
If (Pos1>0) And (Pos2>0) Then
  BEGIN
  St:=Table ;
  Table:=Copy(St,1,Pos1-1) ; TypePlan:=Copy(St,Pos1+1,Pos2-Pos1-1) ; TypeCumul:=Copy(St,Pos2+1,Length(St)-Pos2+1) ;
  END ;
if (Table='GENERAUX')          Then BEGIN Table:='GEN' ; VerifCollectif(Collectif) ; END Else
if (Table='TIERS')             Then BEGIN Table:='TIE' ; VerifCollectif(Collectif) ; END Else
if (Table='SECTION')           Then BEGIN Table:='ANA' ; VerifCollectif(Collectif) ; END Else
if (Table='RUBRIQUE')          Then BEGIN Table:='RUB' ; VerifCollectif(Collectif) ; END Else
If (Table='BUDGENE:B')         Then Table:='BGE' Else
If (Table='BUDSECT:B')         Then Table:='BSE' Else
If (Table='BUDGENE')           Then Table:='BGE' Else
If (Table='BUDSECT')           Then Table:='BSE' Else
If (Table='RUBBUD:B')          Then Table:='RUBBUD' Else
If (Table='RUBBUD:R')          Then Table:='BUDREA' Else
If (Table='BUDGENE:R')         Then Table:='BGEREA' Else
If (Table='BUDSECT:R')         Then Table:='BSEREA' Else
If (Table='GENERAUX/TIERS')    Then BEGIN Table:='G/T' ; VerifCollectif(Collectif) ; END Else
If (Table='TIERS/GENERAUX')    Then BEGIN Table:='T/G' ; VerifCollectif(Collectif) ; END Else
If (Table='GENERAUX/SECTION')  Then BEGIN Table:='G/A' ; VerifCollectif(Collectif) ; END Else
If (Table='SECTION/GENERAUX')  Then BEGIN Table:='A/G' ; VerifCollectif(Collectif) ; END Else
If (Table='BUDSECT/BUDGENE:B') Then Table:='BUDGET:A/G' Else
If (Table='BUDGENE/BUDSECT:B') Then Table:='BUDGET:G/A' Else
If (Table='BUDSECT/BUDGENE:R') Then Table:='BUDGET:A/GREA' Else
If (Table='BUDGENE/BUDSECT:R') Then Table:='BUDGET:G/AREA' ;
END ;

Function GetCumulAstuce(Typ,Code1,Code2,AvecAno,Etab,Devi,Exo,SCollectif : String ;
                        var Dat11,Dat22 : TDateTime ; Collectif,Detail : Boolean ;
                        LesCompRub : TStringList ; Var TResult : TabloExt ; EnMonnaieOpposee : Boolean ;
                        BalSit : String ; WhereSup : String ; ChampSup : String ; ValChampSup : Variant ;
                        CptVariant : String = '' ;
                        TypePlan : String = '' ; TypeCumul : String = '' ; Sdate : String = ''; stRegroupement : string = '') : Double ; //XMG 26/11/03
BEGIN
If (SCollectif<>'X') And (SCollectif<>'') Then BEGIN Devi:=SCollectif ; END ;
if trim(stRegroupement)='' then
    Result:=GetCumul(Typ,Code1,Code2,AvecAno,Etab,Devi,Exo,Dat11,Dat22,Collectif,True,Nil,TResult,EnMonnaieOpposee,CptVariant,TypePlan,TypeCumul,FALSE,BalSit,WhereSup,SDate)
else
    Result:=GetCumulMS(Typ,Code1,Code2,AvecAno,Etab,Devi,Exo,Dat11,Dat22,Collectif,True,Nil,TResult,EnMonnaieOpposee,CptVariant,TypePlan,TypeCumul,FALSE,BalSit,WhereSup,SDate, stRegroupement) ;
END ;

Function  Get_CumulMS(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif : String; stRegroupement : string ='') : Variant ;
Var Err : Integer ;
    DD1,DD2 : TDateTime ;
    Exo : TExoDate ;
    V1 : Variant ;
    T : TabloExt ;
    X : Double ;
    CptVariant : String ;
    TypePlan,TypeCumul : String ;
    LeCumul : TZCumul;
BEGIN
  if (Table='RUBRIQUE') or (Table='GENERAUX') or (Table='TIERS') or (Table='SECTION') then
  begin
    LeCumul := TZCumul.Create;
    try
      LeCumul.InitCriteres(AvecAno,Etab,Devi,SDate,'');
//      LeCumul.Regroupement := trim(GetListeMultiDossier(stRegroupement));
      LeCumul.Regroupement := stRegroupement;
      V1 := LeCumul.GetValeur(Table, Qui, nil);
    finally
      LeCumul.Free;
    end;
  end else
  begin
    NomTable(Table,Collectif,CptVariant,TypePlan,TypeCumul) ;
    if AvecANO='' Then
    BEGIN
      {$IFDEF NOVH}
      V1 := OLEErreur(12);
      {$ELSE}
      If VH^.bByGenEtat Then V1:=0 Else V1:=OLEErreur(12) ;
      {$ENDIF NOVH}
    END
    Else
    BEGIN
      If WhatDate(SDate,DD1,DD2,Err,Exo) Then
      BEGIN
        X:=GetCumulAstuce(Table,Qui,'',AvecAno,Etab,Devi,Exo.Code,Collectif,DD1,DD2,Collectif='X',True,Nil,T,FALSE,'','','',Null,CptVariant,TypePlan,TypeCumul,SDate, stRegroupement) ; //XMG 26/11/03
        {$IFDEF NOVH}
        if X <> Null then V1 := X
                     else V1 := 0;
        {$ELSE}
        If V_PGI.StOLEErr='' Then
          V1:=X
        Else
        BEGIN
          If VH^.bByGenEtat Then V1:=0 Else V1:=V_PGI.StOLEErr ;
        END ;
        {$ENDIF NOVH}
      END
      Else
      BEGIN
        {$IFDEF NOVH}
        V1 := StDateErreur(Err);
        {$ELSE}
        V_PGI.StOLEErr:=StDateErreur(Err) ;
        If VH^.bByGenEtat Then V1:=0 Else V1:=StDateErreur(Err) ;
        {$ENDIF NOVH}
      END ;
    END ;
  end;
  Result:=V1 ;
END ;


Function  Get_Cumul(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif : String) : Variant ;
Var Err : Integer ;
    DD1,DD2 : TDateTime ;
    Exo : TExoDate ;
    V1 : Variant ;
    T : TabloExt ;
    X : Double ;
    CptVariant : String ;
    TypePlan,TypeCumul : String ;
BEGIN
NomTable(Table,Collectif,CptVariant,TypePlan,TypeCumul) ;
if AvecANO='' Then
   BEGIN
   {$IFDEF NOVH}
   V1:=OLEErreur(12);
   {$ELSE}
   If VH^.bByGenEtat Then V1:=0 Else V1:=OLEErreur(12) ;
   {$ENDIF NOVH}
   END Else
   BEGIN
   If WhatDate(SDate,DD1,DD2,Err,Exo) Then
      BEGIN
      X:=GetCumulAstuce(Table,Qui,'',AvecAno,Etab,Devi,Exo.Code,Collectif,DD1,DD2,Collectif='X',True,Nil,T,FALSE,'','','',Null,CptVariant,TypePlan,TypeCumul,SDate,'') ; //XMG 26/11/03
      {$IFDEF NOVH}
      if X <> Null then V1:=X
                   else V1 := 0;
      {$ELSE}
      If V_PGI.StOLEErr='' Then V1:=X Else
        BEGIN
        If VH^.bByGenEtat Then V1:=0 Else V1:=V_PGI.StOLEErr ;
        END ;
      {$ENDIF NOVH}
      END Else
      BEGIN
      {$IFDEF NOVH}
      V1 := StDateErreur(Err);
      {$ELSE}
      V_PGI.StOLEErr:=StDateErreur(Err) ;
      If VH^.bByGenEtat Then V1:=0 Else V1:=StDateErreur(Err) ;
      {$ENDIF NOVH}
      END ;
   END ;
Result:=V1 ;
END ;

Function GetCumulAstucePCL(Typ,Code1,Code2,AvecAno,Etab,Devi,Exo,SCollectif : String ;
                           var Dat11,Dat22 : TDateTime ; Collectif,Detail : Boolean ;
                           LesCompRub : TStringList ; Var TResult : TabloExt ; DeviseEnPivot,EnMonnaieOpposee : Boolean ; CptVariant : String ;
                           TypePlan : String ; TypeCumul : String  ; LesCptes : tStringList; BalSit : string = '') : Double ;
BEGIN
If (SCollectif<>'X') And (SCollectif<>'') Then BEGIN Devi:=SCollectif ; END ;
Result:=GetCumul(Typ,Code1,Code2,AvecAno,Etab,Devi,Exo,Dat11,Dat22,Collectif,True,LesCptes,TResult,EnMonnaieOpposee,CptVariant,TypePlan,TypeCumul,DeviseEnPivot,BalSit) ;
END ;

Function  Get_CumulPCL(Table,Qui,AvecAno,Etab,Devi,SDate,Collectif,DeviseEnPivot,EnMonnaieOpposee : String ; LesCptes : tStringList; BalSit : string ='') : Variant ;
Var Err : Integer ;
    DD1,DD2 : TDateTime ;
    Exo : TExoDate ;
    V1 : Variant ;
    T : TabloExt ;
    X : Double ;
    CptVariant : String ;
    TypePlan,TypeCumul : String ;
    Q : TQuery;
BEGIN
TypePlan:='' ; TypeCumul:='' ; CptVariant:='' ;
NomTable(Table,Collectif,CptVariant,TypePlan,TypeCumul) ;
if AvecANO='' Then
   BEGIN
   {$IFDEF NOVH}
   V1 := OLEErreur(12);
   {$ELSE}
   If VH^.bByGenEtat Then V1:=0 Else V1:=OLEErreur(12) ;
   {$ENDIF NOVH}
   END Else
   BEGIN
   // Si balsit<>'' on n'effectue pas WhatDate sinon plantage car exo peut être absent
   If ((Balsit<>'') or WhatDate(SDate,DD1,DD2,Err,Exo)) Then
      BEGIN
      if Balsit<> '' then
      begin
        Q := OpenSQL ('SELECT * FROM CBALSIT WHERE BSI_CODEBAL="'+Balsit+'"', True);
        if not Q.Eof then
        begin
          DD1 := Q.FindField('BSI_DATE1').AsDateTime;
          DD2 := Q.FindField('BSI_DATE2').AsDateTime;
        end;
        Ferme (Q);
      end;
      X:=GetCumulAstucePcl(Table,Qui,'',AvecAno,Etab,Devi,Exo.Code,Collectif,DD1,DD2,Collectif='X',True,Nil,T,DeviseEnPivot='X',EnMonnaieOpposee='X',CptVariant,TypePlan,TypeCumul,LesCptes,BalSit) ;
      {$IFDEF NOVH}
      if X <> Null then V1 := X
                   else V1 := 0;
      {$ELSE}
      If V_PGI.StOLEErr='' Then V1:=X Else
        BEGIN
        If VH^.bByGenEtat Then V1:=0 Else V1:=V_PGI.StOLEErr ;
        END ;
      {$ENDIF NOVH}
      END Else
      BEGIN
      {$IFDEF NOVH}
      V1 := StDateErreur(Err);
      {$ELSE}
      If VH^.bByGenEtat Then V1:=0 Else V1:=StDateErreur(Err) ;
      {$ENDIF NOVH}
      END ;
   END ;
Result:=V1 ;
END ;


Function  Get_CumulDouble(Table,Qui1,Qui2,AvecAno,Etab,Devi,SDate,Collectif : String) : Variant ;
Var Err : Integer ;
    DD1,DD2 : TDateTime ;
    Exo : TExoDate ;
    V1 : Variant ;
    T : TabloExt ;
    X : Double ;
    CptVariant : String ;
    TypePlan,TypeCumul : String ;
BEGIN
NomTable(Table,Collectif,CptVariant,TypePlan,TypeCumul) ;
if AvecANO='' Then V1:=OLEErreur(12) Else
   BEGIN
   If WhatDate(SDate,DD1,DD2,Err,Exo) Then
      BEGIN
      X:=GetCumulAstuce(Table,Qui1,Qui2,AvecAno,Etab,Devi,Exo.Code,Collectif,DD1,DD2,Collectif='X',True,Nil,T,FALSE,'','','',Null,'',TypePlan,TypeCumul) ;
      {$IFDEF NOVH}
      if X <> Null then V1 := X
                   else V1 := 0;
      {$ELSE}
      If V_PGI.StOLEErr='' Then V1:=X Else V1:=V_PGI.StOLEErr ;
      {$ENDIF NOVH}
      END Else V1:=StDateErreur(Err) ;
   END ;
Result:=V1 ;
END ;

Function  Get_Cumul2(Table,Qui1,AvecAno,Etab,Devi,SDate,Collectif : String ; Qui2,BalSit : String ; WhereSup : String ; ChampSup : String ; ValChampSup : Variant) : Variant ;
Var Err : Integer ;
    DD1,DD2 : TDateTime ;
    Exo : TExoDate ;
    V1 : Variant ;
    T : TabloExt ;
    X : Double ;
    CptVariant : String ;
    TypePlan,TypeCumul : String ;
    CValChampSup : Variant ;
BEGIN
CValChampSup:=Null ;
If ChampSup<>'' Then CValChampSup:=ValChampSup ;
NomTable(Table,Collectif,CptVariant,TypePlan,TypeCumul) ;
if AvecANO='' Then
   BEGIN
   {$IFDEF NOVH}
   V1 := OLEErreur(12);
   {$ELSE}
   If VH^.bByGenEtat Then V1:=0 Else V1:=OLEErreur(12) ;
   {$ENDIF NOVH}
   END Else
   BEGIN
   If WhatDate(SDate,DD1,DD2,Err,Exo) Then
      BEGIN
      If WhereSup<>'' Then WhereSup:=FindetReplace(WhereSup,'''','"',TRUE) ;
      X:=GetCumulAstuce(Table,Qui1,Qui2,AvecAno,Etab,Devi,Exo.Code,Collectif,DD1,DD2,Collectif='X',True,Nil,T,FALSE,BalSit,WhereSup,ChampSup,CValChampSup,CptVariant,TypePlan,TypeCumul) ;
      {$IFDEF NOVH}
      if X <> Null then V1 := X
                   else V1 := 0;
      {$ELSE}
      If V_PGI.StOLEErr='' Then V1:=X Else
        BEGIN
        If VH^.bByGenEtat Then V1:=0 Else V1:=V_PGI.StOLEErr ;
        END ;
      {$ENDIF NOVH}
      END Else
      BEGIN
      {$IFDEF NOVH}
      V1 := StDateErreur(Err);
      {$ELSE}
      If VH^.bByGenEtat Then V1:=0 Else V1:=StDateErreur(Err) ;
      {$ENDIF NOVH}
      END ;
   END ;
Result:=V1 ;
END ;

Function  Get_Cumul2MS(Table,Qui1,AvecAno,Etab,Devi,SDate,Collectif : String ; Qui2,BalSit : String ; WhereSup : String ; ChampSup : String ; ValChampSup : Variant;stRegroupement : string ='') : Variant ;
Var Err : Integer ;
    DD1,DD2 : TDateTime ;
    Exo : TExoDate ;
    V1 : Variant ;
    T : TabloExt ;
    X : Double ;
    CptVariant : String ;
    TypePlan,TypeCumul : String ;
    CValChampSup : Variant ;
BEGIN
CValChampSup:=Null ;
If ChampSup<>'' Then CValChampSup:=ValChampSup ;
NomTable(Table,Collectif,CptVariant,TypePlan,TypeCumul) ;
if AvecANO='' Then
   BEGIN
   {$IFDEF NOVH}
   V1 := OLEErreur(12);
   {$ELSE}
   If VH^.bByGenEtat Then V1:=0 Else V1:=OLEErreur(12) ;
   {$ENDIF NOVH}
   END Else
   BEGIN
   If WhatDate(SDate,DD1,DD2,Err,Exo) Then
      BEGIN
      If WhereSup<>'' Then WhereSup:=FindetReplace(WhereSup,'''','"',TRUE) ;
      X:=GetCumulAstuce(Table,Qui1,Qui2,AvecAno,Etab,Devi,Exo.Code,Collectif,DD1,DD2,Collectif='X',True,Nil,T,FALSE,BalSit,WhereSup,ChampSup,CValChampSup,CptVariant,TypePlan,TypeCumul,'', stRegroupement) ;
      {$IFDEF NOVH}
      if X <> Null then V1 := X
                   else V1 := 0;
      {$ELSE}
      If V_PGI.StOLEErr='' Then V1:=X Else
        BEGIN
        If VH^.bByGenEtat Then V1:=0 Else V1:=V_PGI.StOLEErr ;
        END ;
      {$ENDIF NOVH}
      END Else
      BEGIN
      {$IFDEF NOVH}
      V1 := StDateErreur(Err);
      {$ELSE}
      If VH^.bByGenEtat Then V1:=0 Else V1:=StDateErreur(Err) ;
      {$ENDIF NOVH}
      END ;
   END ;
Result:=V1 ;
END ;



Function Get_Info(Table,Qui,Quoi,Dollar : String) : Variant ;
BEGIN
   if QuelPaysLocalisation=codeISOES then
      Result:=GetInfoTable(Table,Qui,Quoi,Dollar)
   else
      Result:=GetInfoTable(Table,Quoi,Qui,Dollar) ; //XVI 24/02/2005
END ;

Function Get_Constante(CodeCons,CodeEtab,Periode : String) : Variant ;
BEGIN
Result:=GetConstante(CodeCons,CodeEtab,Periode) ;
END ;

function Get_DateCumul(St,Quoi : String) : Variant;
Var err : Integer ;
    Exo : TExoDate ;
    DD1,DD2 : tDateTime ;
begin
Result:='' ;
If WhatDate(St,DD1,DD2,Err,Exo) Then
   BEGIN
   if Quoi='D' then Result:=DD1 ;
   if Quoi='F' then Result:=DD2 ;
   END Else Result:=StDateErreur(Err) ;
end ;

Function  Get_CumulCorresp(Table,CodeCoresp,Plancoresp,TypMvt,Etab,Devi,
                           SDate,TypCalcul,Signe : String) : Variant ;
Var Err : Integer ;
    DD1,DD2 : TDateTime ;
    Exo : TExoDate ;
    V1 : Variant ;
    X : Double ;
BEGIN
if (Table='GENERAUX') then Table:='GEN' else
   if (Table='TIERS') then Table:='TIE' else
      if (Table='SECTION') then Table:='ANA' else
if TypMvt='' then V1:=OLEErreur(12) else
   BEGIN
   if WhatDate(SDate,DD1,DD2,Err,Exo) then
      BEGIN
      X:=GetCumulCorresp(Table,CodeCoresp,Plancoresp,TypMvt,Etab,Devi,Exo.Code,
                         TypCalcul,Signe,DD1,DD2) ;
      if V_PGI.StOLEErr='' then V1:=X else V1:=V_PGI.StOLEErr ;
      END else V1:=StDateErreur(Err) ;
   END ;
Result:=V1 ;
END ;

Function GetCumulCorresp(Typ,CodeCoresp,Plancoresp,TypMvt,Etab,Devi,Exo : String ;
                         TypCalcul,Signe : String ; Var Dat11,Dat22 : TDateTime) : Double ;
Var Table,Champ,ChampCorresp,Axe : String ;
    FilDev,FilExo,FilEtab : Boolean ;
    Multiple : TTypeCalc ;
    Qui,DecDev : Byte ;
    QCpte : TQuery ;
    Sql,WhereAnd : String ;
    SetTyp : SetttTypePiece ;
    Axe1,Axe2,StMvt : String ;
    Lefb1,Lefb2 : TFichierBase ;
    GCalcul : TGCalculCum ;
    ToTal : TabTot ;
    TabResult : TabloExt ;
    EnMonnaieOpposee : Boolean ;
BEGIN
Result:=0 ; V_PGI.StOLEErr:='' ; Qui:=0 ; Multiple:=Un ; Fillchar(TabResult,Sizeof(TabResult),0) ;
if (Typ='') Or (CodeCoresp='') then BEGIN V_PGI.StOLEErr:=OLEErreur(6) ; Exit ; END ;
if (Typ='GEN') Or (Typ='AUX') Or (Typ='ANA') then
    BEGIN
    FilEtab:=(Etab<>'') ; FilDev:=(Devi<>'') ; FilExo:=(Exo<>'') ;
    END else
    BEGIN
    V_PGI.StOLEErr:=OLEErreur(11) ; Exit ;
    END ;
if (Typ='GEN') then Qui:=1 else
   if (Typ='TIE') then Qui:=2 else
      if (Typ='ANA') then Qui:=3 ;
Case Qui of
     1: BEGIN
        Table:='GENERAUX' ; Champ:='G_GENERAL' ; ChampCorresp:='G_CORRESP'+Plancoresp ;
        Axe:='' ; WhereAnd:=' Order By G_GENERAL' ;
        END ;
     2: BEGIN
        Table:='TIERS' ; Champ:='T_AUXILIAIRE' ; ChampCorresp:='T_CORRESP'+Plancoresp ;
        Axe:='' ; WhereAnd:=' Order By T_AUXILIAIRE' ;
        END ;
     3: BEGIN
        Table:='SECTION' ; Champ:='S_SECTION' ; ChampCorresp:='S_CORRESP'+Plancoresp ;
        WhereAnd:=' And S_AXE="'+Copy(CodeCoresp,1,2)+'" Order By S_SECTION' ;
        Axe:=Copy(CodeCoresp,1,2) ;
        END ;
   else BEGIN V_PGI.StOLEErr:=OLEErreur(6) ; Exit ; END ;
  End ;
DecDev:=DecimaleDev(Devi) ;
StMvt:=TypMvt ; QuelTyp(StMvt,SetTyp,TypMvt,EnMonnaieOpposee) ;
QuelFBRub(Lefb1,Lefb2,Typ,Axe,Axe1,Axe2,Multiple,FALSE,FALSE) ;
GCalcul:=TGCalculCum.Create(Multiple,Lefb1,Lefb2,SetTyp,FilDev,FilEtab,FilExo,False,False,DecDev,V_PGI.OkDecE) ;
GCalcul.InitCalcul('','',Axe1,Axe2,Devi,Etab,Exo,Dat11,Dat22,True) ;
Sql:='Select '+Champ+' From '+Table+' Where '+ChampCorresp+'="'+CodeCoresp+'" '+WhereAnd+'' ;
QCpte:=OpenSql(Sql,True) ;
if QCpte.Eof then BEGIN Ferme(QCpte) ; V_PGI.STOLEErr:=OLEErreur(5) ; Exit ; END ;
While Not QCpte.Eof do
      BEGIN
      GCalcul.ReInitCalcul(QCpte.Fields[0].AsString,'',Dat11,Dat22,'') ;
      GCalcul.Calcul ; Total:=GCalcul.ExecCalc.TotCpt ;
      FaitResultat(Total,TabResult,TypMvt,TypCalcul,DecDev,Signe) ;
      QCpte.Next ;
      END ;
SommeLeResultat(TabResult) ; GCalcul.Free ; Ferme(QCpte) ;
Result:=TabResult[1] ; Screen.Cursor:=SyncrDefault ;
END ;

Function inverseMontant(Signe : String ; D,C,R : Double) : Double ;
Var X : Double ;
BEGIN
X:=R ;
If D>C Then
   BEGIN
   if Signe<>'' then if Signe='NEG' then X:=X*(-1) ;
   END Else
If D<C Then
   BEGIN
   if Signe<>'' then if Signe<>'NEG' then X:=X*(-1) ;
   END ;
Result:=X ;
END ;

Function FaitResultatCpt(Var ToTal : TabTot ; AvecAno,St : String ; Deci : Integer ;
                         Signe : String ) : Double ;
Var SoldeTampon,D,C : Double ;

BEGIN
{SoldeTampon:=0 ;} Result:=0 ;
Case AvecANO[1] Of
  'A' : BEGIN D:=Total[0].TotDebit ; C:=Total[0].TotCredit ; END ;
  'S' : BEGIN D:=Total[1].TotDebit ; C:=Total[1].TotCredit ; END ;
  'T' : BEGIN D:=Total[0].TotDebit+Total[1].TotDebit ; C:=Total[0].TotCredit+Total[1].TotCredit ; END ;
  Else BEGIN D:=Total[0].TotDebit+Total[1].TotDebit ; C:=Total[0].TotCredit+Total[1].TotCredit ; END ;
  END ;
if St='SM' then
   BEGIN
   SoldeTampon:=Arrondi(D-C,Deci) ;
   SoldeTampon:=InverseMontant(Signe,D,C,SoldeTampon) ;
   Result:=SoldeTampon ;
   END Else
   if St='SC' then
      BEGIN
      SoldeTampon:=Arrondi(D-C,Deci) ;
      if SoldeTampon<=0 then
         BEGIN
         SoldeTampon:=InverseMontant(Signe,D,C,SoldeTampon) ;
         Result:=SoldeTampon ;
         END ;
      END else
   if St='SD' then
      BEGIN
      SoldeTampon:=Arrondi(D-C,Deci) ;
      if SoldeTampon>=0 then
         BEGIN
         SoldeTampon:=InverseMontant(Signe,D,C,SoldeTampon) ;
         Result:=SoldeTampon ;
         END ;
      END else
   if St='TC' then Result:=Arrondi(C,Deci) else
   if St='TD' then Result:=Arrondi(D,Deci)  ;
END ;

Function CPGetCumulCpteJal(Cpte,Jal : String ; DateDeb,DateFin : TDateTime ; QuelSens,QuelSigne,AvecANO : String) : Double ;
Var GCAlc       : TGCalculCum ;
    QuelTypeEcr : SetttTypePiece ;
    LEtab : String3 ;
    LExo  : TExoDate ;
    MonoExo : Boolean ;
BEGIN
If DateDeb=0 Then DateDeb:=V_PGI.DateEntree ;
If DateFin=0 Then DateFin:=V_PGI.DateEntree ;
//Result:=0 ;
//Exit ;
QuelTypeEcr:=[tpReel] ;
LEtab:='' ;
QuelExoDate(DateDeb,DateFin,MonoExo,LExo) ;
Gcalc:=TGCalculCum.create(CPCpteJal,fbGene,fbJal,QuelTypeEcr,FALSE,LEtab<>'',TRUE,FALSE,FALSE,V_PGI.OkDecV,V_PGI.OkDecE) ;
If Length(Jal)>=4 Then Jal:=Copy(Jal,2,Length(Jal)-1) ; // what a shame !!!
GCalc.initCalcul(Cpte,Jal,'','','',LEtab,LExo.Code,DateDeb,DateFin,TRUE) ;
GCalc.Calcul ;
Result:=FaitResultatCpt(GCalc.ExecCalc.TotCpt,AvecAno,QuelSens,GCalc.PrepCalc.Decimale,QuelSigne) ;
GCalc.Free ;
END ;

Function CP_GetCumulCpteJal(Cpte,Jal,DateDeb,DateFin,QuelSens,QuelSigne,AvecANO : Variant) : Variant ;
Var J,C,QSE,QSI,AA : String ;
    D1,D2 : TDateTime ;
BEGIN
{$IFNDEF EAGLSERVER}
V_PGI.DateEntree:=Now ;
{$ENDIF}

J:=Jal ; C:=Cpte ; D1:=DateDeb ; D2:=DateFin ; QSE:=QuelSens ; QSI:=QuelSigne ; AA:=AvecAno ;
Result:=CPGetCumulCpteJal(Cpte,Jal,D1,D2,QSE,QSI,AA) ;
END ;

Function  CPGetCumulCorresp(CodeCoresp,Plancoresp,Etab,Devi,
                             SDate,TypCalcul,QuelSigne : String) : Variant ;
Var Err : Integer ;
    DD1,DD2 : TDateTime ;
    Exo : TExoDate ;
    V1 : Variant ;
    X : Double ;
    TypMvt,Table : String ;
BEGIN
TypMvt:='N' ; Table:='GENERAUX' ;
if (Table='GENERAUX') then Table:='GEN' else
   if (Table='TIERS') then Table:='TIE' else
      if (Table='SECTION') then Table:='ANA' else
if TypMvt='' then V1:=OLEErreur(12) else
   BEGIN
   if WhatDate(SDate,DD1,DD2,Err,Exo) then
      BEGIN
      X:=GetCumulCorresp(Table,CodeCoresp,Plancoresp,TypMvt,Etab,Devi,Exo.Code,
                         TypCalcul,QuelSigne,DD1,DD2) ;
      if V_PGI.StOLEErr='' then V1:=X else V1:=V_PGI.StOLEErr ;
      END else V1:=StDateErreur(Err) ;
   END ;
Result:=V1 ;
END ;

function DateOnly(d : TDateTime) : TDateTime ;
var dd,mm,yy : word ;
begin
DecodeDate(d,yy,mm,dd) ;
result := EncodeDate(yy,mm,dd) ;
end ;

Function GET_TVATPFINFO(TVAouTPF,REGIME,Code,Quoi,Achat,RG : String ) : Variant ;
var
  IsAchat,IsRG : Boolean ;
Begin
  Result:=#0 ;
  {$IFDEF NOVH}
  {$ELSE}
  IsAchat:=striComp(pchar(Achat),'X')=0 ;
  IsRG:=striComp(pchar(RG),'X')=0 ;
  if stricomp(Pchar(TVAouTPF),'TVA')=0 then begin
     if StriComp(Pchar(Quoi),'TAUX')=0 then
        Result:=TVA2TAUX(Regime,Code,IsAchat)*100 else
     if StriComp(Pchar(Quoi),'CPTE')=0 then
        Result:=TVA2CPTE(Regime,Code,IsAchat) else
     if StriComp(Pchar(Quoi),'ENCAIS')=0 then
        Result:=TVA2ENCAIS(Regime,Code,IsAchat,isRG) else
        ;
  end else
  if stricomp(Pchar(TVAouTPF),'TPF')=0 then begin
     if StriComp(Pchar(Quoi),'TAUX')=0 then
        Result:=TPF2TAUX(Regime,Code,IsAchat)*100 else
     if StriComp(Pchar(Quoi),'CPTE')=0 then
        Result:=TPF2CPTE(Regime,Code,IsAchat) else
     if StriComp(Pchar(Quoi),'ENCAIS')=0 then
        Result:=TPF2ENCAIS(Regime,Code,IsAchat) else
        ;
  end ;
  {$ENDIF NOVH}
End ;
function Get_RIBINFO( vStRib, vstQuoi : String ) : String ;
var lboIsRIB   : Boolean ;
    lStPays    : String ;
    lStCleIBAN : String ;
    lStEtab    : String ;
    lStGuichet : String ;
    lStNum     : String ;
    lStCleRIB  : String ;
    lStDOM     : String ;
Begin
  Result:='' ;
  if (pos(';'+vStQuoi+';',';ETB;GUI;NUM;CLE;DOM;')>0) and (Trim(vStRIB)<>'') then
     Begin
     lStPays:='' ;
     lBoIsRIB:=not RIBEstIBAN(vStRib) ;
     if not lBoIsRIB then
        lBoIsRIB:=IBANtoRIB(Copy(vStRIB,2,length(vStRIB)),vStRIB,lstPays,lstCleIBAN) ;
     if lBoIsRIB then
        Begin
        DecodeRIB(lstEtab,lstGuichet,lstNum,lstCleRIB,lStDOM,vStRIB,lstPays) ;
        if vstQuoi='ETB' then Result:=lStEtab    else
        if vStQuoi='GUI' then Result:=lStGuichet else
        if vStQuoi='NUM' then Result:=lStNum     else
        if vStQuoi='CLE' then Result:=lStCleRIB  else
        if vStQuoi='DOM' then Result:=lStDom     else
           ;
        End ;
     End ;
End ;

function GetSousSections(Axe, Sec : string) : String;
Const
  MaxStruct = 4 ;
var
  i : Integer;
  fb : TFichierBase;
  OkStruct : Boolean;
  NbStruct  : Integer;
  DebStruct : Array[1..MaxStruct] of integer;
  LonStruct : Array[1..MaxStruct] of integer;
  CodStruct : Array[1..MaxStruct] of String3;
  Q : TQuery;
begin
  fb := AxeToFb(Axe);
  OkStruct:=GetInfoCpta(fb).Structure;
  NbStruct:=0 ;

  FillChar(DebStruct,SizeOF(DebStruct),#0);
  FillChar(LonStruct,SizeOF(LonStruct),#0);
  FillChar(CodStruct,SizeOF(CodStruct),#0);

  if OkStruct then begin
    Q:=OpenSQL('SELECT SS_AXE, SS_SOUSSECTION, SS_LIBELLE, SS_DEBUT, SS_LONGUEUR FROM STRUCRSE WHERE SS_AXE="'+Axe+'"' ,True);
    While Not Q.EOF do begin
      Inc(NbStruct);
      DebStruct[NbStruct]:=Q.FindField('SS_DEBUT').AsInteger;
      LonStruct[NbStruct]:=Q.FindField('SS_LONGUEUR').AsInteger;
      CodStruct[NbStruct]:=Q.FindField('SS_SOUSSECTION').AsString;
      Q.Next;
    end;
    Ferme(Q);

    for i:=1 to NbStruct do begin
      Q:=OpenSQL('SELECT PS_CODE, PS_LIBELLE FROM SSSTRUCR WHERE PS_AXE="'+Axe+'" AND PS_SOUSSECTION="'+CodStruct[i]+'" AND PS_CODE="'+Copy(Sec,DebStruct[i],LonStruct[i])+'"', True);
      if Not Q.EOF then Result := Result + '('+Q.Fields[0].AsString+') ' + Q.Fields[1].AsString+'; ';
      Ferme(Q);
    end;
  end;
end;

procedure GetPeriodes(Exo : string ; D1,D2 : TDateTime ; var P1,P2 : integer ; var DD1,DD2,DF1,DF2 : TDateTime) ;
var UnExo : TExoDate ;
    i,NbPer : integer ;
    DP1,DP2 : TDateTime ;
begin
DD1:=0 ; DD2:=0 ; DF1:=0 ; DF2:=0 ; P1:=0 ; P2:=0 ;
UnExo.Code:=Exo ; RempliExoDate(UnExo) ; NbPer:=UnExo.NombrePeriode ;
if D1<>DebutDeMois(D1) then
   begin
   DD1:=D1 ;
   DD2:=FinDeMois(D1) ;
   D1:=DebutDeMois(PlusMois(D1,1)) ;
   end ;
if D2<>FinDeMois(D2) then
   begin
   DF1:=DebutDeMois(D2) ;
   DF2:=D2 ;
   D2:=FinDeMois(PlusMois(D2,-1)) ;
   end ;
DP1:=UnExo.Deb ; DP2:=FinDeMois(DP1) ;
for i:=1 to NbPer do
   begin
   if D1=DP1 then P1:=i ;
   if D2=DP2 then P2:=i ;
   DP1:=DebutDeMois(PlusMois(DP1,1)) ;
   DP2:=FinDeMois(DP1) ;
   end ;
end ;

function ComplementCumul(TypePlan,Compte,Exo : string ; D1,D2 : TDateTime ; TypeResult : string) : Extended ;
// Recherche des écritures sur le complément
var Q : TQuery ;
    FResult, RubCompte : string ;
begin
result:=0 ;
if TypeResult='D' then FResult:='E_DEBIT' else
if TypeResult='C' then FResult:='E_CREDIT' else
if TypeResult='S' then FResult:='E_DEBIT-E_CREDIT' else
// GCO - 21/07/2003 Solution provisoire pour la suppression des champs DEBITEURO, CREDITEURO
//if TypeResult='DE' then FResult:='E_DEBITEURO' else
//if TypeResult='CE' then FResult:='E_CREDIEURO' else
//if TypeResult='SE' then FResult:='E_DEBITEURO-E_CREDITEURO' ;
if TypeResult='DE' then FResult:='0' else
if TypeResult='CE' then FResult:='0' else
if TypeResult='SE' then FResult:='0' ;
// FIN GCO - 21/07/2003

if TypePlan='GEN' then RubCompte:='E_GENERAL' ;
Q:=OpenSQL('SELECT SUM('+FResult+') FROM ECRITURE '+
           'WHERE '+RubCompte+'="'+Compte+'" AND E_EXERCICE="'+Exo+'" AND '+
           'E_DATECOMPTABLE>="'+USDateTime(D1)+'" AND E_DATECOMPTABLE<="'+USDateTime(D2)+'"',TRUE) ;
if not Q.EOF then result:=Q.Fields[0].AsFloat ;
Ferme(Q) ;
end ;

function GetCumulDC(TypePlan,Compte,Exo : string ; Date1,Date2 : TDateTime ; TypeResult : String) : Variant ;
var Q : TQuery ;
    FResult,sSQL : string ;
    Per1,Per2 : integer ;
    DD1,DD2,DF1,DF2 : TDateTime ;
    Montant : Extended ;
begin
Montant:=0 ;
GetPeriodes(Exo,Date1,Date2,Per1,Per2,DD1,DD2,DF1,DF2) ;
if TypeResult='D' then FResult:='CQ_DEBIT' else
if TypeResult='C' then FResult:='CQ_CREDIT' else
if TypeResult='S' then FResult:='CQ_SOLDE' else
// GCO - 21/07/2003 Solution provisoire pour la suppression des champs DEBITEURO, CREDITEURO
//if TypeResult='DE' then FResult:='CQ_DEBITEURO' else
//if TypeResult='CE' then FResult:='CQ_CREDIEURO' else
//if TypeResult='SE' then FResult:='CQ_SOLDEEURO' ;
if TypeResult='DE' then FResult:='0' else
if TypeResult='CE' then FResult:='0' else
if TypeResult='SE' then FResult:='0' ;
// FIN GCO - 21/07/2003

sSQL:='SELECT SUM('+FResult+') FROM CUMULDC WHERE CQ_TYPEPLAN="'+TypePlan+'" ' ;
sSQL:=sSQL+'AND CQ_COMPTE="'+Compte+'" AND CQ_EXERCICE="'+Exo+'" ' ;
sSQL:=sSQL+'AND CQ_PERIODE>='+IntToStr(Per1)+' AND CQ_PERIODE<='+IntToStr(Per2) ;
Q:=OpenSQL(sSQL,True) ;
if not Q.EOF then Montant:=Q.Fields[0].AsFloat ;
Ferme(Q) ;
if (DD1<>0) and (DD2<>0) then Montant:=Montant+ComplementCumul(TypePlan,Compte,Exo,DD1,DD2,TypeResult) ;
if (DF1<>0) and (DF2<>0) then Montant:=Montant+ComplementCumul(TypePlan,Compte,Exo,DF1,DF2,TypeResult) ;
result:=Montant ;
end ;

Procedure Get_CumulColl(Cpt,SDate : String ; Var SoldeD,SoldeC : Double) ;
Var QbalC : TQuery ;
    Err : Integer ;
    DD1,DD2 : TDateTime ;
    Exo : TExoDate ;
    TotCpt1M,TotCpt1S,
    TotCpt2M,TotCpt2S   : TabTot ;
//    TotCpt1MEURO,TotCpt1SEURO,
    TotCpt2MEURO,TotCpt2SEURO   : TabTot ;
BEGIN
SoldeD:=0 ; SoldeC:=0 ;
Fillchar(TotCpt1M,SizeOf(TotCpt1M),#0) ; Fillchar(TotCpt1S,SizeOf(TotCpt1S),#0) ;
Fillchar(TotCpt2M,SizeOf(TotCpt2M),#0) ; Fillchar(TotCpt2S,SizeOf(TotCpt2S),#0) ;
//Fillchar(TotCpt1M,SizeOf(TotCpt1MEURO),#0) ; Fillchar(TotCpt1S,SizeOf(TotCpt1SEURO),#0) ;
Fillchar(TotCpt2M,SizeOf(TotCpt2MEURO),#0) ; Fillchar(TotCpt2S,SizeOf(TotCpt2SEURO),#0) ;

If WhatDate(SDate,DD1,DD2,Err,Exo) Then
   BEGIN
   QBalC:=PrepareTotCptSolde([tpReel],FALSE,FALSE,TRUE,FALSE) ;
   ExecuteTotCptSolde(QBalC,Cpt, Exo.Deb,Exo.Fin,
                      '','',Exo.Code,TotCpt2M,TotCpt2S,TotCpt2MEURO,TotCpt2SEURO,TRUE,TRUE,V_PGI.OkDecV,V_PGI.OkDecE,FALSE,FALSE,[tpReel]) ;
   SoldeD:=TotCpt2S[1].TotDebit ; SoldeC:=TotCpt2S[1].TotCredit ;
   If QBalC<>NIL Then Ferme(QBalC) ;
   END ;
END ;

(*P.FUGIER 08/99 : DEBUT*)
function Get_CumulHisto(Compte, Auxiliaire, Etablissement, Devise, Exercice,
                        DateDeb, DateFin, Periode, Oppose, Options : string) : Variant ;
var QVa : TQuery ; Cols, Where : string ;
//    DateFin2 : TDateTime ;
//    Year, Month, Day : Word ;
begin
Result:=0 ;
if Options='C' then Cols:='HB_CREDIT' else Cols:='HB_DEBIT' ;
if Oppose='X' then Cols:=Cols+'CONTRE' ;
Where:=' HB_TYPEBAL="BDS"'
      +' AND HB_EXERCICE="'+Exercice+'"'
      +' AND HB_DATE1="'+USDateTime(StrToFloat(DateDeb))+'"'
      +' AND HB_DATE2="'+USDateTime(StrToFloat(DateFin))+'"'
      +' AND HB_COMPTE1="'+Compte+'"'
      +' AND HB_COMPTE2="'+Auxiliaire+'"'
      +' AND HB_ETABLISSEMENT="'+Etablissement+'"' ;
QVa:=OpenSQL('SELECT '+Cols+' FROM HISTOBAL WHERE'+Where, TRUE) ;
if not QVa.EOF then Result:=QVA.Fields[0].AsFloat ;
Ferme(QVa) ;
end ;

// TODO : Optimisation + Traitements erreurs paramètres
function Get_CumulBal(Compte, Auxiliaire, Etablissement, Devise, Exercice,
                      DateDeb, DateFin, Periode, Oppose, Options : string ;
                      AvecAno : string; TypeBal : string ; TypeCumul : string) : Variant ;
var QVa : TQuery ;
//    Va : Variant ;
    Montant : Double ;
    {Cols,} Where, Typ : string ;
    bEcrs : Boolean ;
    Dat1, Dat2{, Dat3} : TDateTime ; xResult : TabloExt ;
//    Year, Month, Day : Word ;
begin
Montant:=0 ;
// Traitement des erreurs sur les paramètres
if Exercice='' then begin Result:=Montant ; Exit ; end ;
// Paramètres OK
if (Exercice=GetEncours.Code) or
   (Exercice=GetSuivant.Code) or
   (Exercice=GetPrecedent.Code) then // Exercice N - GetCumul
  begin
  Where:=' E_EXERCICE="'+Exercice+'"'
        +' AND E_QUALIFORIGINE<>"N0A"'
        +' AND E_CREERPAR<>"DET"' ;
  QVa:=OpenSQL('SELECT DISTINCT E_EXERCICE FROM ECRITURE WHERE'+Where, TRUE) ;
  bEcrs:=not QVa.EOF ;
  Ferme(QVa) ;
  if bEcrs then
    begin
    if TypeBal='BDS' then
      begin
      if Auxiliaire<>'' then Typ:='TIE' else Typ:='GEN' ;
      if Options='C' then AvecAno:='C' else AvecAno:='D' ;
      end ;
    Dat1:=StrToFloat(DateDeb) ; Dat2:=StrToFloat(DateFin) ;
    FillChar(xResult, SizeOf(xResult), #0) ;
    GetCumul(Typ, Compte, Auxiliaire, AvecAno, Etablissement, Devise,
             Exercice, Dat1, Dat2, FALSE, FALSE, nil, xResult, Oppose='X') ;
    if Options='C' then Montant:=xResult[4] else Montant:=xResult[5] ;
    end else Montant:=Get_CumulHisto(Compte, Auxiliaire, Etablissement,
                                     Devise, Exercice,DateDeb, DateFin,
                                     Periode, Oppose, Options) ;
  end else // Exercice N-? - CBALSIT
    Montant:=Get_CumulHisto(Compte, Auxiliaire, Etablissement, Devise, Exercice,
                            DateDeb, DateFin, Periode, Oppose, Options) ;
Result:=Montant ;
end ;
(*P.FUGIER 08/99 : FIN*)

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
function CPProcCalcMul(Func,Params,WhereSQL : string ; TT : TDataset ; Total : Boolean) : string ;
var lValeur : Variant;

   function _AffecteCouleur ( vResult : string ) : string;
   begin
     if Pos('C', vResult) > 0 then Result := '#FONT#|ClGreen|#' + vResult
                              else Result := '#FONT#|ClRed|#' + vResult;
   end;

begin
  if Func = 'TTOTALDEBIT'  then begin
    if VH^.CPExoRef.Code = VH^.Suivant.Code then Result := FloatToStr(TT.FindField('T_TOTDEBS').AsFloat)
                                            else Result := FloatToStr(TT.FindField('T_TOTDEBE').AsFloat);
  end
  else
  if Func = 'TTOTALCREDIT' then begin
    if VH^.CPExoRef.Code = VH^.Suivant.Code then Result := FloatToStr(TT.FindField('T_TOTCRES').AsFloat)
                                            else Result := FloatToStr(TT.FindField('T_TOTCREE').AsFloat);
  end
  else
  if Func = 'TSOLDEN' then
  begin
    if VH^.CPExoRef.Code = VH^.Suivant.Code then lValeur := FloatToStr(TT.FindField('T_TOTDEBS').AsFloat - TT.FindField('T_TOTCRES').AsFloat)
                                            else lValeur := FloatToStr(TT.FindField('T_TOTDEBE').AsFloat - TT.FindField('T_TOTCREE').AsFloat);

    Result := AfficheDBCR( lValeur );
    Result := _AffecteCouleur( Result );
    Exit;
  end
  else
  if Func = 'TSOLDEN-1' then
  begin
    if VH^.CPExoRef.Code = VH^.Suivant.Code then lValeur := FloatToStr(TT.FindField('T_TOTDEBE').AsFloat - TT.FindField('T_TOTCREE').AsFloat)
                                            else lValeur := FloatToStr(TT.FindField('T_TOTDEBP').AsFloat - TT.FindField('T_TOTCREP').AsFloat);
    Result := AfficheDBCR( lValeur );
    Result := _AffecteCouleur( Result );
  end;
end;
{$ENDIF EAGLSERVER}

////////////////////////////////////////////////////////////////////////////////

end.



