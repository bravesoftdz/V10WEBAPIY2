unit cum;

interface
  uses SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,
  Utob, CpteUtil, TcalcCum, Ent1, Hent1, HCtrls, dbTables ;

//  Procedure RearrangeLaTobMere(LaTobMere : Tob ; LeMultiple : TTypeCalc ; Ffb1, Ffb2 : TFichierBase) ;
//  Procedure TrieLaTobMere (LaTobMere : Tob ) ;
//  Procedure AlimenteTob(LaFille: tob ; LeMultiple : TTypeCalc ; LeFicBase1,LeFicBase2 : TFichierBase ; Dev,Etab,Exo : String ; D1,D2 : TDateTime ; FFiltresupp, ChampSup : String ; DeviseEnPivot,FEnEuro : Boolean ; FSetTyp : SetttTypePiece ) ;
// ???  Procedure MiseAJourMaman (LaTobMere : Tob ; Date1 : TDatetime ) ;
//  Procedure LesFillesVivantes (LaTobMere : Tob) ;

  Function DonneCumul(LaTobMere : Tob ; Zon1,Zon2 : String ; FMultiple : TTypeCalc ; Ffb1,Ffb2 : TFichierBase ; FSetTyp : SetttTypePiece ;
                      Dev,Etab,Exo : String ; DeviseEnPivot,FEnEuro : Boolean ; D1,D2 :TDatetime ; ChampSup : Variant ; ValChampSup : Variant ; FFiltreSup : String ='' ) : TabTot ;
//  Function CreeLaFille (LaTobMere : Tob ; LeMultiple : TTypeCalc ; LeFicBase1,LeFicBase2 : TFichierbase ; Dev,Etab,Exo : String ; D1,D2 : TDateTime ; FFiltresupp,ChampSup : String ; DeviseEnPivot,FEnEuro : Boolean ; FSetTyp : SetttTypePiece ) : Tob ;
//  Function RechercheOuCreeLaFille(LaTobMere : Tob ; FMultiple : TTypeCalc ; Ffb1,Ffb2 : TFichierBase ; Dev,Etab,Exo : String ; D1,D2 :TDatetime ; DeviseEnPivot,FEnEuro : Boolean ; FSetTyp : SetttTypePiece ; FFiltreSupp, ChampSup : String ) :Tob ;
//  Function VerifieValeur(Zon1,Zon2 : String ; FMultiple : TTypeCalc ; Ffb1,Ffb2 : TFichierBase ; D1,D2 :TDatetime ) : Boolean ;
//  Function CalculTailleTobMere(LaTobMere : Tob ; LeMultiple : TTypeCalc ; LeFicBase1, LeFicBase2 : TFichierBase) : Integer ;


implementation

Uses LicUtil,HMsgBox ;

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
                      FFiltresupp, ChampSup : String ; DeviseEnPivot,FEnEuro : Boolean ; FSetTyp : SetttTypePiece ) ;
Var St,Sum_Montant : String ;
    TobF : tob ;
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
SecondChamps:=False ; Prefixe:='E_' ;

If Ana then Prefixe:='Y_' Else Prefixe:='E_' ;

SecondChamps:=(LeMultiple=Deux) ;

St:='' ; LeOrder:='' ;

If LeFicBase1=FbGene then St:=Prefixe+'GENERAL' ;
If LeFicBase1=FbAux then St:=Prefixe+'AUXILIAIRE';
If LeFicBase1 in [FbAxe1..FbAxe5] then St:=Prefixe+'SECTION' ;

If ChampSup <> '' then S:='SELECT '+ChampSup+', '+St+', ' else S:='SELECT '+St+', ' ;
LeOrder:=St; St:='' ;

If SecondChamps then
  begin
    If LeFicBase2=FbGene then St:=Prefixe+'GENERAL' ;
    If LeFicBase2=FbAux then St:=Prefixe+'AUXILIAIRE';
    If LeFicBase2 in [FbAxe1..FbAxe5] then St:=Prefixe+'SECTION' ;
    LeOrder:=LeOrder+', '+St ;
    S:=S+St+', ' ;
    st:='' ;
  end ;

S:=S+Prefixe+'EXERCICE, ' ;

If (not Ana) then
  If EnDevise(FiltreDev,DeviseEnPivot) Then SUM_Montant:='sum(E_DEBITDEV) as Debit, sum(E_CREDITDEV) as Credit, sum(E_DEBITEURO) as DebitEuro, sum(E_CREDITEURO) as CreditEuro, '
                                       Else SUM_Montant:='sum(E_DEBIT) as Debit, sum(E_CREDIT) as Credit, sum(E_DEBITEURO) as DebitEuro, sum(E_CREDITEURO) as CreditEuro, '
             else
  If EnDevise(FiltreDev,DeviseEnPivot) Then SUM_Montant:='sum(Y_DEBITDEV) as Debit, sum(Y_CREDITDEV) as Credit, sum(Y_DEBITEURO) as DebitEuro, sum(Y_CREDITEURO) as CreditEuro,  '
                                       Else SUM_Montant:='sum(Y_DEBIT) as Debit, sum(Y_CREDIT) as Credit, sum(Y_DEBITEURO) as DebitEuro, sum(Y_CREDITEURO) as CreditEuro, ' ;

If Ana then St:=SUM_MONTANT+'Y_ECRANOUVEAU, Y_QUALIFPIECE FROM ANALYTIQ '
       else St:=SUM_MONTANT+'E_ECRANOUVEAU, E_QUALIFPIECE FROM ECRITURE ' ;
S:=S+St ; st:='' ;

//Le Where
S:=S+' Where '+Prefixe+'DATECOMPTABLE >="'+USDateTime(D1)+'" And '+Prefixe+'DATECOMPTABLE <= "'+UsDateTime(D2)+'" ' ;

If Ana then if (LeFicBase1 in [FbAxe1..fbAxe5]) then st:='AND Y_AXE="'+FbToAxe(LeFicbase1)+'" '
                                                else st:='AND Y_AXE="'+FbToAxe(LeFicBase2)+'" ' ;
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

Q:=OpenSql(S, True) ;

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
Result :=0 ;
Longueur :=0 ; LaTaille :=0 ;

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
      Tf.Free ; Tf:=Nil ; LaTaille:=CalculTailleTobMere(LaTobMere,LeMultiple, Ffb1, Ffb2) ;
      end else
      begin
      Tf.Free ; Tf:=Nil ; LaTaille:=CalculTailleTobMere(LaTobMere,LeMultiple, Ffb1, Ffb2) ;
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
Function CreeLaFille(LaTobMere : Tob ; LeMultiple : TTypeCalc ; LeFicBase1,LeFicBase2 : TFichierBase ; Dev,Etab,Exo : String ; D1,D2 : TDateTime ; FFiltresupp,ChampSup : String ; DeviseEnPivot,FEnEuro : Boolean ; FSetTyp : SetttTypePiece) : Tob ;
var LaFille, LaPetiteFille  : Tob ;
    OnSupprime: Boolean ;
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
AlimenteTob(LaFille,LeMultiple,LeFicBase1,LeFicBase2,Dev,Etab,Exo,D1,D2,FFiltresupp,ChampSup,DeviseEnPivot,FEnEuro, FSetTyp ) ;
Result:=Lafille ;
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 18/06/2002
Modifié le ... :   /  /
Description .. : Recherche la tob des parametres (+données) ou la cree
Mots clefs ... :
*****************************************************************}
Function RechercheOuCreeLaFille(LaTobMere : Tob ;FMultiple : TTypeCalc ; Ffb1,Ffb2 : TFichierBase ; Dev,Etab,Exo : String ; D1,D2 :TDatetime ; DeviseEnPivot,FEnEuro : Boolean ; FSetTyp : SetttTypePiece ; FFiltreSupp, ChampSup : String ) :Tob ;
var LeMultiple,LeFicBase1,LeFicBase2 :String ;
    LaFille, LaPetiteFille : Tob ;
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
  If V_PGI.PassWord=CryptageSt(DayPass(Date)) Then
    BEGIN
    St:='Non Trouvé : dev : '+dev+','+'Etab : '+Etab+', Exo : '+Exo+', Date1 :'+DateToStr(D1)+', Date2 :'+DateToStr(D2) ;
    HShowMessage('0;Non trouvé :;'+St+';E;O;O;O;','','') ;
    END ;
  LaFille:=CreeLaFille(LaTobMere,FMultiple,Ffb1,Ffb2,Dev,Etab,Exo,D1,D2,FFiltresupp,ChampSup,DeviseEnPivot,FEnEuro,FSetTyp ) ;
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
                    Dev,Etab,Exo : String ; DeviseEnPivot,FEnEuro : Boolean ; D1,D2 :TDatetime ; ChampSup : Variant ; ValChampSup : Variant ; FFiltreSup : String ='' ) : TabTot ;
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

LaTob:=RechercheOuCreeLaFille(LaTobMere, FMultiple, Ffb1,Ffb2, Dev,Etab,Exo, D1,D2, DeviseEnPivot,FEnEuro, FSetTyp, FFiltreSup, ChampSup ) ;

{$IFDEF TESTGV}
LaTobMere.saveToFile('c:\Halley\mere.txt',false,true,true) ;
{$ELSE}
{$ENDIF}

If Ana then Prefixe:='Y_' Else Prefixe:='E_' ;

Case Ffb1 Of
  FbGene : Ch1:=Prefixe+'GENERAL' ;
  FbAux  : Ch1:=Prefixe+'AUXILIAIRE';
  fbAxe1..fbAxe5 : Ch1:=Prefixe+'SECTION' ;
  END ;
UnSeulChamp:=(FMultiple=Un) ;

If (not UnSeulChamp) then
  begin
    Case Ffb2 Of
    FbGene : Ch2:=Prefixe+'GENERAL' ;
    FbAux  : Ch2:=Prefixe+'AUXILIAIRE';
    fbAxe1..fbAxe5 : Ch2:=Prefixe+'SECTION' ;
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

(*
{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 20/06/2002
Modifié le ... :   /  /
Description .. : Supprime les tobs filles non utilisées depuis un certain
Suite ........ : temps
Mots clefs ... :
*****************************************************************}
Procedure LesFillesVivantes (LaTobMere : Tob) ;
var LF : TOB ;
    D1,D2 : TDateTime ;
    i:Integer ;
begin
For i:=LaTobMere.Detail.Count-1 downTo 0 do
  begin
  LF:=LaTobMere.detail[i] ;
  D1:=LF.GetValue('HeureAcces') ;
  D2:=Now ;
  if ((D2-D1)>TMaxDuree) then
    begin
    Lf.free ;
    Lf:=Nil ;
    End ;
  end ;
end ;
*)
(*
{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 20/06/2002
Modifié le ... :   /  /
Description .. : Libere des filles de la tob mere si la date passee en
Suite ........ : parametre est inclu dans le bornage des dates des tob filles
Mots clefs ... :
*****************************************************************}
Procedure MiseAJourMaman (LaTobMere : Tob ; Date1 : TDatetime ) ;
var D1,D2 : TDateTime ;
    LF : Tob ;
    i:Integer ;
begin
For i:=LaTobMere.Detail.Count-1 downTo 0 do
  begin
  LF:=LaTobMere.detail[i] ;
  D1:=LF.GetValue('Date1') ;
  D2:=LF.GetValue('Date2') ;
  if ((Date1>=D1) and (Date1<=D2)) then begin Lf.Free ; LF:=Nil ; end ;
  end ;
end ;
*)

end.
