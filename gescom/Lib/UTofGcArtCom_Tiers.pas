unit UTofGcArtCom_Tiers;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF,
{$IFDEF EAGLCLIENT}
      UTob,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul, db,DBGrids,
{$ENDIF}
      HDimension,UTOM;

Type
     TOF_GcArtCom_Tiers = Class (TOF)
        procedure OnLoad ; override ;
     END ;

implementation

procedure TOF_GcArtCom_Tiers.OnLoad();
Var StCaption, Libelle : string;
    Q:TQuery;
    Lab : THLabel;
    Chps : THEdit;
    Top : integer;
begin
stCaption:= GetControlText('GL_TIERS');
Q:=OpenSQL('SELECT T_LIBELLE FROM TIERS WHERE T_TIERS="'+StCaption+'"',True) ;
Libelle:=Q.Findfield('T_LIBELLE').AsString;
Ecran.Caption:='Articles commandés par : '+stCaption +' - '+Libelle;
Ferme(Q) ;
Q:=OpenSQL('SELECT T_NATUREAUXI FROM TIERS WHERE T_TIERS="'+StCaption+'"',True) ;
if Q.Findfield('T_NATUREAUXI').AsString = 'FOU' then
begin
    Ecran.Caption:='Articles commandés chez : '+stCaption +' - '+Libelle;
    SetControlProperty('GL_TIERSLIVRE', 'Visible', False);
    Lab := THLabel(Ecran.FindComponent('TGL_TIERSFACTURE'));
    Lab.Caption := 'Facturation par';
    Lab := THLabel(Ecran.FindComponent('TGL_TIERSPAYEUR'));
    Lab.Caption := 'Payer à';
    Chps := THEdit(Ecran.FindComponent('GL_TIERSFACTURE'));
    SetControlProperty('GL_TIERSPAYEUR', 'Top', Chps.Top);
    Lab := THLabel(Ecran.FindComponent('TGL_TIERSFACTURE'));
    SetControlProperty('TGL_TIERSPAYEUR', 'Top', Lab.Top);
    Chps := THEdit(Ecran.FindComponent('GL_TIERSLIVRE'));
    SetControlProperty('GL_TIERSFACTURE', 'Top', Chps.Top);
    Lab := THLabel(Ecran.FindComponent('TGL_TIERSLIVRE'));
    SetControlProperty('TGL_TIERSFACTURE', 'Top', Lab.Top);
end;
Ferme(Q) ;
end;


Initialization
registerclasses([TOF_GcArtCom_Tiers]);
end.

