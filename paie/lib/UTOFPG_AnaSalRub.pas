{ PT1   26/12/2007 FC V_810 Concept accessibilité fiche salarié
}
unit UTofPG_ANASALRUB;

interface
uses  StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}DBGrids,
{$ENDIF}
      Grids,HCtrls,HEnt1,HMsgBox,UTOF, UTOB, UTOM,  Vierge, P5Util, P5Def, AGLInit, EntPaie,
      PgOutils, Ventil ;

Type
     TOF_PGANASALRUB = Class (TOF)
       private
       Salarie,NomSal : String;
       LaRem  : THEdit;
       Action:String;//PT1
       procedure PGAnaSalRub (Sender: TObject);
       procedure ChangeRem (Sender: TObject);
       public
       procedure OnArgument(Arguments : String ) ; override ;
     END ;

implementation
procedure TOF_PGANASALRUB.ChangeRem(Sender: TObject);
var  LeLibRem  : THLabel;
begin
if LaRem <> NIL then
 begin
 LeLibRem := THLabel (GetControl ('LIBREM'));
 if LeLibRem <> NIL then
 SetControlProperty ('LIBREM', 'Caption', RechDom ('PGREMUNERATION', LaRem.Text, FALSE));
 end;
end;

procedure TOF_PGANASALRUB.OnArgument(Arguments: String);
var    BtnVal : TToolbarButton97;
       LeNom  : THLabel;
begin
inherited ;
Salarie:=ReadTokenSt(Arguments);   // Recup Code Salarie
NomSal:=ReadTokenSt(Arguments);   // Recup Nom Salarie
//DEB PT1
Action:='';
if Arguments <> '' then
  Action:=ReadTokenSt(Arguments);
//FIN PT1
LeNom := THLabel (GetControl ('LIBSAL'));
if LeNom <> NIL then LeNom.Caption := NomSal;
LaRem := THEdit (GetControl ('REM'));
if LaRem <> NIL then LaRem.OnChange := ChangeRem;

BtnVal := TToolbarButton97 (GetControl ('BValider'));
if BtnVal <> NIL then BtnVal.OnClick := PGAnaSalRub;
end;

procedure TOF_PGANASALRUB.PGAnaSalRub(Sender: TObject);
var LaRub : THEdit;
    Q     : TQuery;
    Rubrique   : String;
begin
LaRub := THEdit (GetControl ('REM'));
if LaRub <> NIL then
 begin
 Rubrique := LaRub.Text;
 if Rubrique = '' then exit;
 Q := OpenSql ('SELECT PRM_RUBRIQUE FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_RUBRIQUE="'+Rubrique+'"',TRUE);
 if NOT Q.EOF then
    if Action = 'CONSULTATION' then //PT1
      ParamVentil ('PG',Salarie+';'+Rubrique , '12345', taConsult, False) //PT1
    else
      ParamVentil ('PG',Salarie+';'+Rubrique , '12345', taModif, False)
 else PGIBOX ('Vous devez sélectionner une rubrique', 'Ventilation analytique');
 Ferme (Q);
 end;
end;

Initialization
registerclasses([TOF_PGANASALRUB]);
end.
