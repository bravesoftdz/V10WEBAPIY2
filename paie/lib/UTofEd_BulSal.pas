{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE
*****************************************************************}
unit UTofEd_BulSal;

interface
uses  StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls,HTB97,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBGrids,
{$ELSE}

{$ENDIF}
      Grids,HCtrls,HEnt1,HMsgBox,UTOF,UTOB,UTOM,Vierge,P5Util,P5Def,AGLInit;
Type
     TOF_PGED_BULSAL = Class (TOF)
       private
       CodeSal, Etab, DateD, DateF : String;
       DD, DF : TDateTime;
       public
       procedure OnArgument(Arguments : String ) ; override ;
     END ;

implementation

procedure TOF_PGED_BULSAL.OnArgument(Arguments: String);
var Btn : TToolbarButton97;
    st : String;
{$IFNDEF EAGLCLIENT}
    CSal, CDd, CDf : THDbEdit;
{$ELSE}
    CSal, CDd, CDf : THEdit;
{$ENDIF}
    CEtab : THValComboBox;
begin
inherited ;
st:=Trim (Arguments);
Etab:=ReadTokenSt(st);   // Recup Code Etablissement
CodeSal:=ReadTokenSt(st);// Recup code Salarie
DateD:=ReadTokenSt(st);
DD:=StrToDate(DateD);
DateF:=ReadTokenSt(st);
DF:=StrToDate(DateF);
CEtab:=THValComboBox(GetControl ('PHB_ETABLISSEMENT'));
if CEtab<>NIL then begin CEtab.Value:=Etab; CEtab.enabled :=FALSE; end;
{$IFNDEF EAGLCLIENT}
CSal:=THDbEdit(GetControl ('PHB_SALARIE'));
CDd:=THDbEdit(GetControl ('PHB_DATEDEBUT'));
CDf:=THDbEdit(GetControl ('PHB_DATEFIN'));
{$ELSE}
CSal:=THEdit(GetControl ('PHB_SALARIE'));
CDd:=THEdit(GetControl ('PHB_DATEDEBUT'));
CDf:=THEdit(GetControl ('PHB_DATEFIN'));
{$ENDIF}

if CSal<>NIL then begin CSal.Text:=CodeSal; CSal.enabled :=FALSE; end;
if CDd<>NIL then begin CDd.Text:=DateToStr(DD); CDd.enabled :=FALSE; end;
if CDf<>NIL then begin CDf.Text:=DateToStr(DF); CDf.enabled :=FALSE; end;
if (CEtab <>NIL) AND (CSal<>NIL) AND (CDd<>NIL) AND (CDf<>NIL) then
 begin
 Btn:=TToolbarButton97 (GetControl ('BValider'));
 if Btn <>NIL then begin  Btn.Click; end;
 end;
end;


Initialization
registerclasses([TOF_PGED_BULSAL]);
end.
