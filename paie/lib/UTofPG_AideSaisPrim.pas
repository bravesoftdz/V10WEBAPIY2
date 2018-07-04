{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 05/08/2002
Modifié le ... :   /  /
Description .. : Unit aide à la saisie des primes : saisie mt ou taux
Suite ........ : et calcul automatique du taux en fct du montant saisi
Mots clefs ... : PAIE;SAISPRIM
*****************************************************************}
unit UTofPG_AideSaisPrim;

interface
uses  StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,Fe_Main,DBGrids,
{$ELSE}
       MaineAgl,
{$ENDIF}
      Grids,HCtrls,HEnt1,HMsgBox,UTOF,UTOB,Vierge,AGLInit;
Type
     TOF_PG_AideSaisPrim = Class (TOF)
       private
       Base,Taux,Montant : Double;
       CodeCalcul        : String;
       Dec               : Integer;
       procedure ExitTaux (Sender: TObject);
       procedure ExitMontant (Sender: TObject);
       public
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnClose                  ; override ;
     END ;

implementation

Uses P5Util,P5Def;

procedure TOF_PG_AideSaisPrim.ExitMontant(Sender: TObject);
begin
// Calcul du taux
if Base = 0 then Base := 1 ;
Montant := Valeur (GetControlText('LEMONTANT'));
if CodeCalcul = '05' then  Taux := (Montant / Base) * 100
   else Taux := Montant / Base;
Taux := ARRONDI (Taux, Dec);
SetControlText ('LETAUX', DoubleToCell (Taux,Dec));
// Recalcul du montant à cause arrondi
if CodeCalcul = '05' then Montant := Base * (Taux / 100)
   else Montant := Base * taux;
Montant := ARRONDI (Montant, 2);
SetControlText ('LEMONTANT', DoubleToCell (Montant,2));
end;

procedure TOF_PG_AideSaisPrim.ExitTaux(Sender: TObject);
begin
Taux := Valeur (GetControlText('LETAUX'));
if Taux = 0 then Taux := 1 ;
if CodeCalcul = '05' then Montant := Base * (Taux / 100)
   else Montant := Base * taux;
Montant := ARRONDI (Montant, 2);
SetControlText ('LEMONTANT', DoubleToCell (Montant,2));
end;

procedure TOF_PG_AideSaisPrim.OnArgument(Arguments: String);
var 
    st,Nom,Lib : String;
    ChT,ChM    : THEdit;
    Mt         : Double;
begin
inherited ;
st:=Trim (Arguments);
Base := Valeur(ReadTokenSt(st));
SetControlText ('LABASE', DoubleToCell (Base,2));
Taux := Valeur(ReadTokenSt(st));
SetControlText ('LETAUX', DoubleToCell (Taux,Dec));
if st <> '' then Nom := ReadTokenSt(st);
CodeCalcul := ReadTokenSt(st);
Dec        := StrToInt(ReadTokenSt(st)); // Nbre de décimales du taux
Lib        := ReadTokenSt(st); // libelle de la colonne de la saisie des primes
if Nom <> '' then
   begin
   Ecran.Caption := Ecran.Caption + Nom;
   UpdateCaption (Ecran);
   end ;
SetControlText ('LBLSAL','Détermination du taux pour la colonne '+Lib);
if CodeCalcul = '05' then Mt := Base * (Taux / 100)
   else Mt := Base * taux;
Mt := ARRONDI (Mt, 2);
SetControlText ('LEMONTANT', DoubleToCell (Mt,2));
ChT := THedit (GetControl ('LETAUX'));
if Cht <> NIL then ChT.OnExit := ExitTaux;
ChM := THedit (GetControl ('LEMONTANT'));
if ChM <> NIL then ChM.OnExit := ExitMontant;
end;

procedure TOF_PG_AideSaisPrim.OnClose;
begin
  inherited;
// on recupère le taux calculé pour la saisie des primes
TFVierge(Ecran).Retour := GetControlText ('LETAUX');
end;

Initialization
registerclasses([TOF_PG_AideSaisPrim]);
end.
