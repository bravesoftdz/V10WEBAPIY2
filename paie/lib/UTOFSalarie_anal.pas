{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/09/2001
Modifié le ... : 11/09/2001
Description .. : Unit de gestion pour avoir la liste des salariés sans
Suite ........ : affectation analytique
Mots clefs ... : PAIE;ANALYTIQUE
*****************************************************************}
{
 PT1 11/10/01 PH Raffraichissement de la liste avant lancement du traitement
 PT2 21/05/07 FC V_72 FQ 14112 fetch en CWAS pour prendre en compte la liste complète
}
unit UTOFSalarie_anal;

interface
uses  StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls,HTB97,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,Mul,Fe_Main,DBGrids,
{$ELSE}
       MaineAgl,eMul,
{$ENDIF}
      Hqry,Grids,HCtrls,HEnt1,EntPaie,HMsgBox,UTOF,UTOB,UTOM,Vierge,P5Util,P5Def,
      AGLInit,PgOutils, PgOutils2;

Type
     TOF_Salarie_anal = Class (TOF)
       procedure OnArgument (stArgument: String);        Override;
       procedure ExitEdit(Sender: TObject);
       procedure ActiveWhere(Sender: TObject);
       procedure OnLoad ;                               Override;
   private
       Argument : string;
       Procedure LanceAffSection(Sender: TObject);
    END ;

implementation

procedure TOF_Salarie_anal.OnArgument (stArgument: String);
var Defaut: THEdit;
    Num : Integer;
    BtnValidMul : TToolbarButton97;
begin
// argument = Y  ==> affectation par defaut des ventilations analytiques salaries en fonction des paramètres société
Argument :=stArgument;
For Num := 1 to VH_Paie.PGNbreStatOrg do
 begin
 VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
 end;
Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
If Defaut<>nil then Defaut.OnExit:=ExitEdit;
BtnValidMul:=TToolbarButton97 (GetControl ('BOuvrir'));

if BtnValidMul <> NIL then
 begin
 if Argument = 'Y' then  BtnValidMul.OnClick := LanceAffSection
   else SetControlVisible ('BOuvrir', FALSE);
 end;
end;

procedure TOF_Salarie_anal.LanceAffSection (Sender: TObject);
var     BtnCherche : TToolbarButton97;
begin
// PT1 11/10/01 PH Raffraichissement de la liste avant lancement du traitement
BtnCherche:=TToolbarButton97 (GetControl ('BCherche'));
if BtnCherche <> NIL then BtnCherche.Click ;

if VH_Paie.PGAnalytique then
  begin
{$IFDEF EAGLCLIENT}
  //DEB PT2
  if TFMul(Ecran).Fetchlestous then
    TheMulQ := TOB(Ecran.FindComponent('Q'))
  else
  begin
    PgiBox('Vous n''avez pas de ligne total dans votre liste, #13#10 Traitement impossible ', Ecran.Caption);
    exit;
  end;
  //FIN PT2
{$ELSE}
  TheMulQ := THQuery(Ecran.FindComponent('Q'));
{$ENDIF}
  AglLanceFiche ('PAY','REAFFANAL_AUTO', '', '' , Argument);
  TheMulQ := NIL;
  end;
end;

procedure TOF_Salarie_anal.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;



procedure TOF_Salarie_anal.OnLoad ;
begin
inherited;
ActiveWhere (NIL);
if Argument = 'S' then Ecran.Caption := 'Liste des salariés sans ventilation analytique'
 else if Argument = 'Y' then Ecran.Caption := 'Réaffectation automatique des sections analytiques salariés';
UpdateCaption(TFmul(Ecran));
end;

procedure TOF_Salarie_anal.ActiveWhere(Sender: TObject);
var WW : THEdit;
begin
if Argument = 'S' then
   begin
   WW:=THEdit (GetControl ('XX_WHERE'));
   if WW <> NIL then
   WW.Text := '(NOT exists (SELECT V_NATURE FROM VENTIL WHERE V_NATURE="SA1" AND V_COMPTE=SALARIES.PSA_SALARIE))';
   end;
end;


Initialization
registerclasses([TOF_Salarie_anal]) ;

end.



