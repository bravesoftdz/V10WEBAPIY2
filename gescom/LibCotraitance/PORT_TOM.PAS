{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 01/04/2011
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PORT (PORT)
Mots clefs ... : TOM;PORT
*****************************************************************}
Unit PORT_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     FichList,
{$else}
     eFiche,
     eFichList,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     Hdb,
     UTob,
     Windows,
     Graphics ;

Type
  TOM_PORT = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    procedure OnLoadAlerte               ; override ;
    procedure OnAfterCommit              ; override ;
    procedure OnAfterRollBack            ; override ;
    procedure OnAfterBeginTrans          ; override ;
  private
    ActionFiche : TActionFiche;
    fTypefrs : string;
    TypePort : THDbValComboBox;
    ColorDisabled,ColorEnabled : Tcolor;
    GPO_FRAISREPARTIS,GPO_FRANCO,GPO_RETENUEDIVERSE : THDBCheckbox;
    GPO_TYPEPORT,GPO_VERROU,GPO_COMPTAARTICLE : THDBValComboBox;
    GPO_PVHT,GPO_PVTTC,GPO_COEFF,GPO_MINIMUM,GPO_MINIMUMTTC : THDBEdit;

    procedure afterFormShow;
    procedure GestionEcranCotraitance;
    procedure DefinieSaisie;
    procedure Controlechamp(Champ, Valeur: string);
    procedure SetEvents (Active : boolean);
    procedure FRANCOClick (Sender : TObject);
    procedure FRAISREPARTISClick (Sender : TObject);
    procedure TYPEPORTChange (Sender : TObject);
    procedure GPO_RETENUEDIVERSEChange (Sender : Tobject);
    procedure FamilleCptaChange (sender : Tobject);
    procedure COLECTIFChange (Sender : Tobject);
  end ;

Implementation

uses DBCtrls;

procedure TOM_PORT.OnNewRecord ;
begin
  Inherited ;
  ThLabel(getControl('TGPO_COLLECTIF')).visible := false;
  ThEdit(getControl('GPO_COLLECTIF')).visible := false;
  ThLabel(getControl('TGPO_COLLECTIFAC')).visible := false;
  ThEdit(getControl('GPO_COLLECTIFAC')).visible := false;

end ;

procedure TOM_PORT.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_PORT.OnUpdateRecord ;
begin
  Inherited ;
  if GPO_RETENUEDIVERSE.Checked then
  begin
    if (ThDbEdit(GetControl('GPO_COLLECTIF')).Text = '') and (ThDbEdit(GetControl('GPO_COLLECTIFAC')).Text = '')  and (THDBValComboBox (GetControl('GPO_COMPTAARTICLE')).Value='') then
    begin
      PgiError('Merci de rensigner un mode de ventilation');
      LastError := 1;
      exit;
    end;

    if (THDBValComboBox(GetControl('GPO_COMPTAARTICLE')).Value <>'') and (THDBValComboBox(GetControl('GPO_FAMILLETAXE1')).Value='') then
    begin
      PgiError('Merci de rensigner une famille de taxe');
      LastError := 2;
      exit;
    end;
  end;


  if fTypefrs = 'CO' then
  begin
  	SetField('GPO_TYPEFRAIS','B00');
  end else
  begin
  	SetField('GPO_TYPEFRAIS','501');
  end;
end ;

procedure TOM_PORT.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_PORT.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_PORT.OnLoadRecord ;
begin
  Inherited ;
  ThEdit(getControl('GPO_COLLECTIF')).Enabled := false;
  ThEdit(getControl('GPO_COLLECTIFAC')).Enabled := false;
  ThvalCombobox(getControl('GPO_COMPTAARTICLE')).enabled := true;
  if ActionFiche <> TaCreat then
  begin
    GPO_FRAISREPARTIS.enabled := false;
    GPO_FRAISREPARTIS.Color := ColorDisabled;
//    SetControlEnabled('GPO_RETENUEDIVERSE',false);
  end else
  begin
    SetControlText('GPO_VERROU','NON');
  end;
  if GPO_FRAISREPARTIS.Checked then
  begin
    SetControlEnabled('GPO_FAMILLETAXE1',false);
    ThvalCombobox(getControl('GPO_FAMILLETAXE1')).Value  := '';
    SetControlEnabled('GPO_RETENUEDIVERSE',false);
    SetControlChecked('GPO_RETENUEDIVERSE',false);
    ThEdit(getControl('GPO_COLLECTIF')).Enabled := false;
    ThEdit(getControl('GPO_COLLECTIF')).Text  := '';
    ThEdit(getControl('GPO_COLLECTIFAC')).Enabled := false;
    ThEdit(getControl('GPO_COLLECTIFAC')).Text  := '';
    ThvalCombobox(getControl('GPO_COMPTAARTICLE')).enabled := false;
    ThvalCombobox(getControl('GPO_COMPTAARTICLE')).Value  := '';
  end;
  if GPO_RETENUEDIVERSE.Checked then
  begin
    ThLabel(getControl('TGPO_COLLECTIF')).visible := true;
    ThLabel(getControl('TGPO_COLLECTIFAC')).visible := true;
    ThEdit(getControl('GPO_COLLECTIF')).visible := true;
    ThEdit(getControl('GPO_COLLECTIF')).Enabled := true;
    ThEdit(getControl('GPO_COLLECTIFAC')).visible := true;
    ThEdit(getControl('GPO_COLLECTIFAC')).Enabled := true;
  end;
end ;

procedure TOM_PORT.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_PORT.Controlechamp (Champ , Valeur : string);
begin
  if Champ = 'ACTION' then
  begin
    if valeur = 'CREATION' then ActionFiche := taCreat
    else if Valeur = 'MODIFICATION' then ActionFiche := TaModif
    else if Valeur = 'CONSULTATION' then ActionFiche := TaConsult;
  end;
end;

procedure TOM_PORT.OnArgument ( S: String ) ;
var Critere,Champ,valeur : string;
    i : Integer;
begin
  Inherited ;
  Critere := S;
  While (Critere <> '') do
  BEGIN
    i:=pos(':',Critere);
    if i = 0 then i:=pos('=',Critere);
    if i <> 0 then
       begin
       Champ:=copy(Critere,1,i-1);
       Valeur:=Copy (Critere,i+1,length(Critere)-i);
       end
    else
       Champ := Critere;
    Controlechamp(Champ, Valeur);
    Critere:=(Trim(ReadTokenSt(S)));
  END;

  ColorDisabled := TGroupBox(GetControl('GB_METHODE')).Color;
  ColorEnabled := TGroupBox(GetControl('GPO_LIBELLE')).Color;
  //
  GPO_FRAISREPARTIS :=  THDBCheckbox(GetControl('GPO_FRAISREPARTIS'));
  GPO_FRANCO :=  THDBCheckbox(GetControl('GPO_FRANCO'));
  GPO_TYPEPORT :=  THDBValComboBox(GetControl('GPO_TYPEPORT'));
  GPO_COMPTAARTICLE :=  THDBValComboBox(GetControl('GPO_COMPTAARTICLE'));
  GPO_VERROU :=  THDBValComboBox(GetControl('GPO_VERROU'));
  GPO_PVHT := THDBEdit(GetControl('GPO_PVHT'));
  GPO_PVTTC  := THDBEdit(GetControl('GPO_PVTTC'));
  GPO_COEFF  := THDBEdit(GetControl('GPO_COEFF'));
  GPO_MINIMUM  := THDBEdit(GetControl('GPO_MINIMUM'));
  GPO_MINIMUMTTC  := THDBEdit(GetControl('GPO_MINIMUMTTC'));
  GPO_RETENUEDIVERSE :=  THDBCheckbox(GetControl('GPO_RETENUEDIVERSE'));

  if Pos('COTRAITANCE',S) > 0 then
  begin
  	fTypefrs := 'CO'; // frais de gestion cotraitance
  end else
  begin
  	fTypefrs := 'PO'; // frais et ports standard
  end;

  if fTypeFrs = 'CO' then
  begin
    GestionEcranCotraitance;
  end
  else
  begin
    THdBValComboBox(getControl('GPO_TYPEPORT')).plus := ' AND (CO_LIBRE LIKE "%PO%")';
  end;

  SetControlVisible('GPO_FRANCO',false);
  SetControlVisible('GPO_MINIMUM',false);
  SetControlVisible('GPO_MINIMUMTTC',false);
  SetControlVisible('TGPO_MINIMUM',false);
  SetControlVisible('TGPO_MINIMUMTTC',false);
  SetControlVisible('GPO_FAMILLETAXE2',false);
  SetControlVisible('TGPO_FAMILLETAXE2',false);

  TFFiche(Ecran).OnAfterFormShow := AfterFormShow;
  SetEvents (True);
end ;


Procedure TOM_PORT.DefinieSaisie;
begin
  SetEvents(False);
  // Init --
  GPO_PVHT.Enabled :=  true;
  GPO_PVHT.Color := ColorEnabled;
  //
  GPO_PVTTC.Enabled :=  true;
  GPO_PVTTC.color := ColorEnabled;
  //
  GPO_COEFF.Enabled := true;
  GPO_COEFF.Color:= ColorEnabled;
  //
  GPO_VERROU.Enabled := true;
  GPO_VERROU.Color:= ColorEnabled;
  // -------
  if GPO_TYPEPORT.Value =  'MT' then
  begin
    GPO_COEFF.Enabled := false;
    GPO_COEFF.Color:= ColorDisabled;
    GPO_VERROU.Enabled := false;
    GPO_VERROU.Color:= ColorDisabled;
    GPO_COEFF.Enabled :=  false;
    GPO_COEFF.Color := ColorDisabled;
    GPO_PVTTC.Enabled := false;
    GPO_PVTTC.Color := ColorDisabled;
  end;
  if GPO_TYPEPORT.Value =   'MTC' then
  begin
    GPO_COEFF.Enabled := false;
    GPO_COEFF.Color:= ColorDisabled;
    GPO_VERROU.Enabled := false;
    GPO_VERROU.Color:= ColorDisabled;
    GPO_COEFF.Enabled :=  false;
    GPO_COEFF.Color := ColorDisabled;
    GPO_PVHT.Enabled := false;
    GPO_PVHT.Color := ColorDisabled;
  end;
  if GPO_TYPEPORT.Value =   'MI' then
  begin
    GPO_PVHT.Enabled := true;
    GPO_PVHT.Color := ColorEnabled;
    GPO_PVTTC.Enabled := false;
    GPO_PVTTC.Color := ColorDisabled;
  end;
  if GPO_TYPEPORT.Value =   'MIC' then
  begin
    GPO_PVHT.Enabled := false;
    GPO_PVHT.Color := ColorDisabled;
    GPO_PVTTC.Enabled := true;
    GPO_PVTTC.Color := ColorEnabled;
  end;
  if (GPO_TYPEPORT.Value =   'HT') or (GPO_TYPEPORT.Value = 'PT') then
  begin
    GPO_PVHT.Enabled := false;
    GPO_PVHT.Color := ColorDisabled;
    GPO_PVTTC.Enabled := false;
    GPO_PVTTC.Color := ColorDisabled;
  end;

  if GPO_FRAISREPARTIS.Checked then
  begin
    ThLabel(getControl('TGPO_COLLECTIF')).visible := true;
    ThLabel(getControl('TGPO_COLLECTIFAC')).visible := true;
    ThEdit(getControl('GPO_COLLECTIF')).visible := true;
    ThEdit(getControl('GPO_COLLECTIFAC')).visible := true;
    SetControlEnabled('GPO_FAMILLETAXE1',false);
    ThvalCombobox(getControl('GPO_FAMILLETAXE1')).Value  := '';
    GPO_TYPEPORT.enabled :=  false;
    GPO_PVHT.enabled := false;
    GPO_PVHT.Color := ColorDisabled;
    GPO_PVTTC.Enabled := false;
    GPO_PVTTC.Color := ColorDisabled;
    ThvalCombobox(getControl('GPO_COLLECTIF')).enabled := false;
    ThvalCombobox(getControl('GPO_COLLECTIFAC')).enabled := false;
  end;
  SetEvents(True);
end;



Procedure TOM_PORT.GestionEcranCotraitance;
begin

  SetControlVisible ('GP_TVA',false);
  SetControlVisible ('GPO_PVHT',false);
  SetControlVisible ('TGPO_PVHT',false);
  SetControlVisible ('GPO_FRAISREPARTIS',false);
  SetControlVisible ('GPO_FRANCO',false);
  SetControlVisible ('GPO_MINIMUM',false);
  SetControlVisible ('TGPO_MINIMUM',false);
  SetControlVisible ('GPO_MINIMUMTTC',false);
  SetControlVisible ('TGPO_MINIMUMTTC',false);
  SetControlVisible ('TGPO_VERROU',false);
  SetControlVisible ('GPO_VERROU',false);
  SetControlVisible ('GPO_RETENUEDIVERSE',false);
  SetControlVisible ('GPO_COLLECTIF',false);
  SetControlVisible ('TGPO_COLLECTIF',false);
  SetControlVisible ('GPO_COLLECTIFAC',false);
  SetControlVisible ('TGPO_COLLECTIFAC',false);

  TypePort  := THDBValComboBox(GetControl('GPO_TYPEPORT'));
  TypePort.plus := ' AND (CO_LIBRE LIKE "%CO%")';

  if TypePort.Value = 'PT' then
  begin
    SetControlEnabled('GPO_COEFF', True);
    SetControlEnabled('GPO_PVTTC', False);
  end else
  begin
    SetControlEnabled('GPO_COEFF', False);
    SetControlEnabled('GPO_PVTTC', True);
  end;

end;


procedure TOM_PORT.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_PORT.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_PORT.OnLoadAlerte;
begin
  Inherited ;
end ;

procedure TOM_PORT.OnAfterBeginTrans;
begin
  Inherited ;
end ;

procedure TOM_PORT.OnAfterCommit;
begin
  Inherited ;
end ;

procedure TOM_PORT.OnAfterRollBack;
begin
  Inherited ;
end ;

procedure TOM_PORT.afterFormShow;
begin
  if fTypefrs = 'CO' then
  begin
    Ecran.Caption := 'Frais de gestion de cotraitance';
    UpdateCaption(ecran);
  end;
  DefinieSaisie;
end;

procedure TOM_PORT.SetEvents(Active: boolean);
begin
  if Active then
  begin
    GPO_FRANCO.OnClick := FRANCOClick;
    GPO_FRAISREPARTIS.OnClick := FRAISREPARTISClick;
    GPO_TYPEPORT.OnChange := TYPEPORTChange;
    GPO_RETENUEDIVERSE.OnClick := GPO_RETENUEDIVERSEChange;
    GPO_COMPTAARTICLE.OnChange := FamilleCptaChange;
    ThEdit(getControl('GPO_COLLECTIF')).OnChange := COLECTIFChange;
    ThEdit(getControl('GPO_COLLECTIFAC')).OnChange := COLECTIFChange;
  end else
  begin
    GPO_FRANCO.OnClick := nil;
    GPO_FRAISREPARTIS.OnClick := nil;
    GPO_TYPEPORT.OnChange := nil;
    GPO_COMPTAARTICLE.OnChange := nil;
    ThEdit(getControl('GPO_COLLECTIF')).OnChange := nil;
    ThEdit(getControl('GPO_COLLECTIFAC')).OnChange := nil;
  end;
end;

procedure TOM_PORT.FRANCOClick(Sender: TObject);
begin
  if GPO_FRANCO.Checked then
  begin
    GPO_MINIMUM.Enabled :=  false;
    GPO_MINIMUM.Color := ColorDisabled;
    GPO_MINIMUMTTC.Enabled := false;
    GPO_MINIMUMTTC.Color := ColorDisabled;
  end else
  begin
    GPO_MINIMUM.Enabled := true;
    GPO_MINIMUM.Color := ColorEnabled;
    GPO_MINIMUMTTC.Enabled := true;
    GPO_MINIMUMTTC.Color := ColorEnabled;
  end;
end;

procedure TOM_PORT.FRAISREPARTISClick(Sender: TObject);
begin
  if GPO_FRAISREPARTIS.Checked then
  begin
    GPO_TYPEPORT.Value := 'HT';
    SetControlText('GPO_PVHT','0.0');
    SetControlText('GPO_PVTTC','0.0');
    SetControlEnabled('GPO_RETENUEDIVERSE',false);
    SetControlChecked('GPO_RETENUEDIVERSE',false);
//    SetControlEnabled('GPO_FAMILLETAXE1',false);
//    ThvalCombobox(getControl('GPO_COMPTAARTICLE')).enabled := false;
  end else
  begin
    SetControlEnabled('GPO_RETENUEDIVERSE',true);
//    ThvalCombobox(getControl('GPO_COMPTAARTICLE')).enabled := true;
//    SetControlEnabled('GPO_FAMILLETAXE1',true);
  end;
  DefinieSaisie;
end;

procedure TOM_PORT.TYPEPORTChange(Sender: TObject);
begin
  DefinieSaisie;
end;

procedure TOM_PORT.GPO_RETENUEDIVERSEChange(Sender: Tobject);
begin
  ThEdit(getControl('GPO_COLLECTIF')).Enabled := GPO_RETENUEDIVERSE.Checked;
  ThEdit(getControl('GPO_COLLECTIFAC')).Enabled := GPO_RETENUEDIVERSE.Checked;
  if GPO_RETENUEDIVERSE.Checked then
  begin
    ThLabel(getControl('TGPO_COLLECTIF')).visible := true;
    ThLabel(getControl('TGPO_COLLECTIFAC')).visible := true;
    ThEdit(getControl('GPO_COLLECTIF')).visible := true;
    ThEdit(getControl('GPO_COLLECTIFAC')).visible := true;
//    ThvalCombobox(getControl('GPO_COMPTAARTICLE')).enabled := false;
//    ThvalCombobox(getControl('GPO_COMPTAARTICLE')).Value  := '';
//    SetControlEnabled('GPO_COMPTAARTICLE',false);
//    SetControlEnabled('GPO_FAMILLETAXE1',false);
//    ThvalCombobox(getControl('GPO_FAMILLETAXE1')).Value  := '';
  end else
  begin
    ThLabel(getControl('TGPO_COLLECTIF')).visible := false;
    ThLabel(getControl('TGPO_COLLECTIFAC')).visible := false;
    ThEdit(getControl('GPO_COLLECTIF')).text := '';
    ThEdit(getControl('GPO_COLLECTIF')).visible := false;
    ThEdit(getControl('GPO_COLLECTIFAC')).text := '';
    ThEdit(getControl('GPO_COLLECTIFAC')).visible := false;
    ThvalCombobox(getControl('GPO_COMPTAARTICLE')).enabled := true;
    SetControlEnabled('GPO_FAMILLETAXE1',true);
  end;
end;

procedure TOM_PORT.FamilleCptaChange(sender: Tobject);
begin
  SetEvents(False);
  if GPO_COMPTAARTICLE.value <> '' then
  begin
    SetField('GPO_COLLECTIF','');
    SetField('GPO_COLLECTIFAC','');
  end;
  SetEvents(true);
end;

procedure TOM_PORT.COLECTIFChange(Sender: Tobject);
begin
  SetEvents(False);
  if (ThEdit(getControl('GPO_COLLECTIF')).Text <> '') or (ThEdit(getControl('GPO_COLLECTIFAC')).Text <> '') then
  begin
    SetField('GPO_FAMILLETAXE1','');
    SetField('GPO_COMPTAARTICLE','');
  end;
  SetEvents(true);
end;

Initialization
  registerclasses ( [ TOM_PORT ] ) ;
end.
