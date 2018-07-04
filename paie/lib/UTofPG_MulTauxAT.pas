{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /    
Description .. : Unit de gestion du multi critère de gestion des taux at de 
Suite ........ : l'établissement
Mots clefs ... : PAIE
*****************************************************************}
unit UTofPG_MulTauxAT;

interface
uses  StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls,HTB97,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,Mul,Fe_Main,DBGrids,
{$ELSE}
       MaineAgl,eMul,
{$ENDIF}
      Grids,HCtrls,HEnt1,EntPaie,HMsgBox,UTOF,UTOB,UTOM,Vierge,P5Util,P5Def,
      AGLInit;
Type
     TOF_PGMULTAUXAT = Class (TOF)
       private
       Date2 : THEdit;
       procedure DateExit(Sender: TObject);
       public
       procedure OnArgument(Arguments : String ) ; override ;
     END ;

implementation


procedure TOF_PGMULTAUXAT.DateExit(Sender: TObject);
begin
if Date2.Text = '' then Date2.Text := DateToStr (StrToDate ('01/01/1999'));
end;

procedure TOF_PGMULTAUXAT.OnArgument(Arguments: String);
begin
inherited ;
Date2:= THEdit (GetControl ('PAT_DATEVALIDITE'));
if Date2 <> NIL then Date2.OnClick := DateExit;
end;

Initialization
registerclasses([TOF_PGMULTAUXAT]);
end.
