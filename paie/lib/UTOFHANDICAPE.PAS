{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 13/06/2001
Modifié le ... : 30/08/2001
Description .. : Source TOF du multi critère : sélection des  HANDICAPES
Mots clefs ... : PAIE;HANDICAPE
*****************************************************************
PT1 : 22/10/2001 V562 Ajout de la procédure ExitEdit pour affichage du matricule salarié + modification
    des procédures onload et choix : les 2 catégories d'handicapés sont gérés dans le même combo
PT2 05/01/2005 JL V_650 FQ 12778 Supp clause xx_where pour voir tous les salariés
}
Unit UTOFHANDICAPE ;

Interface

Uses StdCtrls, Controls, Classes,  forms, sysutils,  ComCtrls,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,Mul,DBCtrls,
{$ELSE}
     emul,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, UTOF,UTOB,HTB97,
     ParamSoc,HPanel,ParamDat, Dialogs,Windows,PGoutils,PGoutils2,PgCongesPayes,P5Def,
     EntPaie;
Type
  TOF_HANDICAPE = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private

    procedure CHOIX(Sender : Tobject);
    procedure ExitEdit(Sender: TObject);
  end ;

Implementation


procedure TOF_HANDICAPE.OnLoad ;
begin
     Inherited ;
     if GetControlText('CHOIX')='TOUS' then
        SetControlText('XX_WHERE',''); //PT2
     if GetControlText('CHOIX')='HANDICAPE' then
        SetControlText('XX_WHERE','PGH_HANDICAP="X"');
     if GetControlText('CHOIX')='COTOREP' then
        SetControlText('XX_WHERE','PGH_COTOREP="X"');
     if GetControlText('CATEGHANDI')='' then
        SetControlText('XX_WHERECATEGHANDI','');
     if GetControlText('CATEGHANDI')='A' then
        SetControlText('XX_WHERECATEGHANDI','PGH_CATEGHANDI="A"');
     if GetControlText('CATEGHANDI')='B' then
        SetControlText('XX_WHERECATEGHANDI','PGH_CATEGHANDI="B"');
     if GetControlText('CATEGHANDI')='C' then
        SetControlText('XX_WHERECATEGHANDI','PGH_CATEGHANDI="C"');
     if GetControlText('CATEGHANDI')='ADT' then
        SetControlText('XX_WHERECATEGHANDI','PGH_ACCTRAV="X"');
     if GetControlText('CATEGHANDI')='IDP' then
        SetControlText('XX_WHERECATEGHANDI','PGH_INVPEN="X"');
     if GetControlText('CATEGHANDI')='MDG' then
        SetControlText('XX_WHERECATEGHANDI','PGH_MUTILE="X"');
     if GetControlText('CATEGHANDI')='AMG' then
        SetControlText('XX_WHERECATEGHANDI','PGH_ASSMUTILE="X"');
     if GetControlText('PLACEANT')='' then
        SetControlText('XX_WHEREPLACEANT','');
     if GetControlText('PLACEANT')='ATP' then
        SetControlText('XX_WHEREPLACEANT','PGH_PLACEANT="ATP"');
     if GetControlText('PLACEANT')='IMP' then
        SetControlText('XX_WHEREPLACEANT','PGH_PLACEANT="IMP"');
     if GetControlText('PLACEANT')='CAT' then
        SetControlText('XX_WHEREPLACEANT','PGH_PLACEANT="CAT"');
     if GetControlText('PLACEANT')='CDT' then
        SetControlText('XX_WHEREPLACEANT','PGH_PLACEANT="CDT"');
     if GetControlText('PLACEANT')='CFP' then
        SetControlText('XX_WHEREPLACEANT','PGH_PLACEANT="CFP"');
end ;

procedure TOF_HANDICAPE.OnArgument (S : String ) ;
var
{$IFNDEF EAGLCLIENT}
   Combo:THDBValComboBox;
{$ELSE}
   Combo : THValComboBox;
{$ENDIF}
   Num : Integer;
   Defaut:THEdit;

begin
     For Num := 1 to VH_Paie.PGNbreStatOrg do
         begin
         if Num >4 then Break;
         VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
         end;
     VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;
      SetControlText('CHOIX','TOUS');
{$IFNDEF EAGLCLIENT}
     Combo:=THDBValComboBox(GetControl('CHOIX')) ;
{$ELSE}
     Combo:=THValComboBox(GetControl('CHOIX')) ;
{$ENDIF}
     if Combo<>nil then Combo.OnClick:=CHOIX;
     Defaut:=ThEdit(getcontrol('PGH_SALARIE'));
     If Defaut<>nil then Defaut.OnExit:=ExitEdit;
end ;

procedure TOF_HANDICAPE.CHOIX(Sender: Tobject);
begin
     if GetControlText('CHOIX')='TOUS' then
        begin
        setcontrolenabled ('CATEGHANDI',False);
        setcontrolenabled ('PLACEANT',False);
        SetControlText('CATEGHANDI','');
        SetControlText('PLACEANT','');
        end;
     if GetControlText('CHOIX')='HANDICAPE' THEN
        begin
        setcontrolenabled ('CATEGHANDI',True);
        setcontrolenabled ('PLACEANT',False);
        SetControlText('PLACEANT','');
        SetControlProperty('CATEGHANDI','Datatype','PGHANDINONCOTOREP');
        end;
     if GetControlText('CHOIX')='COTOREP' THEN
        begin
        setcontrolenabled ('CATEGHANDI',True);
        setcontrolenabled ('PLACEANT',True);
        SetControlProperty('CATEGHANDI','Datatype','PGCATEGHANDI');
        end;
end;

// PT1
procedure TOF_HANDICAPE.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

Initialization
  registerclasses ( [ TOF_HANDICAPE ] ) ;
end.
