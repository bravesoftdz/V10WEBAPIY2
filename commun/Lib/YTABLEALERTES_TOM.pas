{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/10/2006
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : YTABLEALERTES (YTABLEALERTES)
Mots clefs ... : TOM;YTABLEALERTES
*****************************************************************}
Unit YTABLEALERTES_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
{$IFNDEF EAGLCLIENT}
     db,
{$ELSE EAGLCLIENT}
     Utob,
{$ENDIF EAGLCLIENT}
     {$IFDEF EAGLSERVER}
       { TS 8/10/7 : unité inutile en EAGLSERVER }
       qsdf,
     {$ENDIF EAGLSERVER}
     UTOM ;

Type
  TOM_YTABLEALERTES = Class (TOM)
    procedure OnArgument ( S: String )   ; override ;
    procedure OnLoadRecord ; override ;
    procedure AlerteActiveClick(Sender: TObject);
    procedure OnAfterUpdateRecord        ; override ;
    end ;
Implementation

uses
  HDb
  {$IFNDEF EAGLCLIENT}
    ,FichList
  {$ELSE}
    ,eFichList
  {$ENDIF}
  ,YAlertesConst
  ;

procedure TOM_YTABLEALERTES.OnAfterUpdateRecord ;
begin
  Inherited ;
  if isFieldModified('YTA_ACTIF') and ( not TCheckBox(GetControl('YTA_ACTIF')).Checked ) then
    executeSql ('UPDATE YALERTES SET YAL_ACTIVE = "-" WHERE YAL_PREFIXE="'+GetField('YTA_PREFIXE')+'" AND YAL_ACTIVE = "X"');
end;

procedure TOM_YTABLEALERTES.OnArgument ( S: String ) ;
begin
  Inherited ;
  {$IFNDEF GPAO}
    { Alertes - Tables gérées seulement en GP : Wxx et GPO, GEM, GDE, YTS }
    TFFicheListe(Ecran).SetNewRange('', WhereAlertesNotGP);
  {$ENDIF GPAO}

  if Assigned(GetControl('YTA_ACTIF')) then
    TCheckBox(GetControl('YTA_ACTIF')).OnClick:=AlerteActiveClick;
end;

procedure TOM_YTABLEALERTES.AlerteActiveClick(Sender: TObject);
begin
  if (TCheckBox(GetControl('YTA_ACTIF')).Checked = false) and (DS.State <> dsBrowse) then
    begin
    if PgiAsk ('S''il existe des alertes liées à cette table, elles seront désactivées. Confirmez-vous ?','Alertes désactivées') = MrNo then
      TCheckBox(GetControl('YTA_ACTIF')).Checked := true;
    end;
end;

procedure TOM_YTABLEALERTES.OnLoadRecord ;
begin
  Inherited ;
  SetControlVisible ('BINSERT', False);
  SetControlVisible('BDelete', False);
  SetControlVisible('YTA_PREDEFINI',false);
  SetControlVisible('TYTA_PREDEFINI',false);
  SetControlVisible('BLIAISON',false);
end ;

Initialization
  registerclasses ( [ TOM_YTABLEALERTES ] ) ;
end.

