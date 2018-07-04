{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 15/03/2002
Modifié le ... :   /  /
Description .. : TOM Gestion des objectifs des stages
Mots clefs ... : PAIE
*****************************************************************
 }
unit UTOMSTAGEOBJECTIF;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,Spin,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,Fe_Main,Fiche,
{$ELSE}
       MaineAgl,eFiche,
{$ENDIF}
      HCtrls,HEnt1,HMsgBox,UTOM,UTOB,HTB97,PgOutils,SaisieList,uTableFiltre ;

Type
     TOM_STAGEOBJECTIF = Class(TOM)
       procedure OnArgument (stArgument : String ) ; override ;
       procedure OnNewRecord  ; override ;
       procedure OnUpdateRecord ; override ;
       private
//       LaNature,LeStage : String;
       TF : TTableFiltre;
       Arg : String;
     END ;

implementation
{ TOM_STAGEOBJECTIF }

procedure TOM_STAGEOBJECTIF.OnArgument(stArgument: String);
var       St : String;
begin
  inherited;
{st:=Trim (StArgument);
LaNature:=ReadTokenSt(St);// Recup de la nature sur laquelle on travaille
LeStage :=ReadTokenSt(St);// recup du code du stage
if LeStage <> '' then SetControlEnabled ('POS_CODESTAGE',FALSE);
SetControlEnabled ('POS_ORDRE',FALSE);
if LaNature <> '' then
   begin
   SetControlEnabled ('POS_NATOBJSTAGE',FALSE);
   end;}
   Arg := ReadTokenPipe(StARgument,';');
   TF  :=  TFSaisieList(Ecran).LeFiltre;
end;

procedure TOM_STAGEOBJECTIF.OnNewRecord;
var Code : String;
begin
  inherited;
  If Arg = 'EMPLOI' then
  begin
        Code := TF.TOBFiltre.GetValue('CC_CODE');
        SetField('POS_NATOBJSTAGE','EMP');
  end
  else
  begin
        Code := TF.TOBFiltre.GetValue('PST_CODESTAGE');
        SetField('POS_NATOBJSTAGE','FOR');
  end;
  SetField('POS_CODESTAGE',Code);
{  SetControlEnabled ('POS_ORDRE',FALSE);
  SetField ('POS_NATOBJSTAGE', LaNature);
  SetField ('POS_CODESTAGE', LeStage);

  If (GetField('POS_NATOBJSTAGE')='') OR (GetField('POS_CODESTAGE') = '') then
     begin
     PgiBox ('Vous devez saisir une nature et un code stage', Ecran.caption);
     if GetField('POS_NATOBJSTAGE')= '' then SetFocusControl ('POS_NATOBJSTAGE');
        if GetField('POS_CODESTAGE')  = '' then SetFocusControl ('POS_CODESTAGE');
     exit;
     end;}
end;

procedure TOM_STAGEOBJECTIF.OnUpdateRecord;
Var QQ      : TQuery ;
    IMax    :integer ;
    St      : String;
begin
If (GetField('POS_NATOBJSTAGE')='') OR (GetField('POS_CODESTAGE') = '') then
     begin
     LastError:=1;
     LastErrorMsg:='Vous devez renseigner la nature et le code du stage';
     exit;
     end;
If (DS.State in [dsInsert]) then
   begin     // increment automatique du numero d'ordre au moment de la creation
   st := 'SELECT MAX(POS_ORDRE) FROM STAGEOBJECTIF WHERE POS_NATOBJSTAGE="'+GetField('POS_NATOBJSTAGE')+
        '" AND POS_CODESTAGE="'+GetField('POS_CODESTAGE')+'"';
   QQ:=OpenSQL(st,TRUE) ;
   if Not QQ.EOF then
     begin
     IMax:=QQ.Fields[0].AsInteger;
     if IMax <> 0 then
       IMax := IMax + 1
     else
       IMax:=1;
     end
     else IMax:=1 ;
   Ferme(QQ) ;
   SetField ('POS_ORDRE', IMax);
   end;     
end;

Initialization
registerclasses([TOM_STAGEOBJECTIF]) ;
end.
