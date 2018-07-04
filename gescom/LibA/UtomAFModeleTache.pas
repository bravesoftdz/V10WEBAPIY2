{***********UNITE*************************************************
Auteur  ...... : AB
Créé le ...... : 18/02/2003
Modifié le ... :   /  /
Description .. : Saisie des Modèles de tâches
Mots clefs ... : TABLE : AFMODELETACHE - Fiche AFTACHEMODELE
******************************************************************}
Unit UtomAFModeleTache;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     dbtables, 
     Fiche, 
     FichList, 
{$ELSE}
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
     UTob,
     UtilGC,AfUtilArticle,UtilTaches,UtilRessource, paramsoc, DicoAf ;

Type
  TOM_AFMODELETACHE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    private
      fRB_MODEGENE1       : TRadioButton;
      fRB_MODEGENE2       : TRadioButton;
      fRB_QUOTIDIENNE     : TRadioButton;
      fRB_NBINTERVENTION  : TRadioButton;
      fRB_HEBDOMADAIRE    : TRadioButton;
      fRB_ANNUELLE        : TRadioButton;
      fRB_MENSUELLE       : TRadioButton;
      fPC_FREQUENCE       : TPageControl;
      fTS_QUOTIDIENNE     : TTAbSheet;
      fTS_NBINTERVENTION  : TTAbSheet;
      fTS_HEBDOMADAIRE    : TTAbSheet;
      fTS_ANNUELLE        : TTAbSheet;
      fTS_MENSUELLE       : TTAbSheet;
      fRB_MOISMETHODE1  : TRadioButton;
      fRB_MOISMETHODE2  : TRadioButton;
      fRB_MOISMETHODE3  : TRadioButton;
      procedure InitNewRegles;
      procedure fRB_MODEGENE1OnClick(Sender: TObject);
      procedure fRB_MODEGENE2OnClick(Sender: TObject);
      procedure fRB_QUOTIDIENNEOnClick(Sender: TObject);
      procedure fRB_NBINTERVENTIONOnClick(Sender: TObject);
      procedure fRB_HEBDOMADAIREOnClick(Sender: TObject);
      procedure fRB_ANNUELLEOnClick(Sender: TObject);
      procedure fRB_MENSUELLEOnClick(Sender: TObject);
      procedure fRB_MOISMETHODE1OnClick(SEnder: TObject);
      procedure fRB_MOISMETHODE2OnClick(SEnder: TObject);
      procedure fRB_MOISMETHODE3OnClick(SEnder: TObject);
      procedure fED_DATEPERIODDEBOnExit(SEnder: TObject);
      procedure fED_DATEPERIODFINOnExit(SEnder: TObject);
    end ;

const                        
      TexteMsgTache: array[1..6] of string 	= (
        {1}        'La date de début est supérieure à la date de fin.'
        {2}        ,'L''article est obligatoire.'
        {3}        ,'L''article n''existe pas.'
        {4}        ,'La durée d''une intervention doit être positive.'
        {5}        ,'La famille de tâche est obligatoire.'
        {6}        ,'Ce modèle de tâche a été utilisé en création de tâche,'+ #13#10 +' Confirmez-vous sa suppression ?'
                   );
Implementation

procedure TOM_AFMODELETACHE.OnArgument ( S: String ) ;
begin
  Inherited ;
  SetControlProperty('AFM_TYPEARTICLE','Plus',PlusTypeArticle) ;
  // traduction champs libres
  GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'AFM_LIBRETACHE', 10, '_');
  GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'AFM_DATELIBRE', 3, '_');
  GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'AFM_CHARLIBRE', 3, '_');
  GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'AFM_VALLIBRE', 3, '_');
  GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'AFM_BOOLLIBRE', 3, '_');

  fPC_FREQUENCE       := TPageControl(GetControl('PC_FREQUENCE'));
  fTS_QUOTIDIENNE     := TTabSheet(GetControl('TS_QUOTIDIENNE'));
  fTS_NBINTERVENTION  := TTabSheet(GetControl('TS_NBINTERVENTION'));
  fTS_HEBDOMADAIRE    := TTabSheet(GetControl('TS_HEBDOMADAIRE'));
  fTS_ANNUELLE        := TTabSheet(GetControl('TS_ANNUELLE'));
  fTS_MENSUELLE       := TTabSheet(GetControl('TS_MENSUELLE'));
  fRB_QUOTIDIENNE     := TRadioButton(GetControl('RB_QUOTIDIENNE'));
  fRB_NBINTERVENTION  := TRadioButton(GetControl('RB_NBINTERVENTION'));
  fRB_HEBDOMADAIRE    := TRadioButton(GetControl('RB_HEBDOMADAIRE'));
  fRB_ANNUELLE        := TRadioButton(GetControl('RB_ANNUELLE'));
  fRB_MENSUELLE       := TRadioButton(GetControl('RB_MENSUELLE'));

  fRB_MODEGENE1 := TRadioButton(GetControl('RB_MODEGENE1'));
  fRB_MODEGENE2 := TRadioButton(GetControl('RB_MODEGENE2'));
  fRB_MODEGENE1.OnClick := fRB_MODEGENE1OnClick;
  fRB_MODEGENE2.OnClick := fRB_MODEGENE2OnClick;

  fRB_MOISMETHODE1 := TRadioButton(GetControl('RB_MOISMETHODE1'));
  fRB_MOISMETHODE2 := TRadioButton(GetControl('RB_MOISMETHODE2'));
  fRB_MOISMETHODE3 := TRadioButton(GetControl('RB_MOISMETHODE3'));
  fRB_MOISMETHODE1.OnClick := fRB_MOISMETHODE1OnClick;
  fRB_MOISMETHODE2.OnClick := fRB_MOISMETHODE2OnClick;
  fRB_MOISMETHODE3.OnClick := fRB_MOISMETHODE3OnClick;

  fTS_QUOTIDIENNE.TabVisible := false;
  fTS_HEBDOMADAIRE.TabVisible := false;
  fTS_NBINTERVENTION.TabVisible := false;
  fTS_ANNUELLE.TabVisible := false;
  fTS_MENSUELLE.TabVisible := false;

  fRB_QUOTIDIENNE.OnClick     := fRB_QUOTIDIENNEOnClick;
  fRB_NBINTERVENTION.OnClick  := fRB_NBINTERVENTIONOnClick;
  fRB_HEBDOMADAIRE.OnClick    := fRB_HEBDOMADAIREOnClick;
  fRB_ANNUELLE.OnClick        := fRB_ANNUELLEOnClick;
  fRB_MENSUELLE.OnClick       := fRB_MENSUELLEOnClick;
  THEdit(GetControl('DATEPERIOD')).OnExit    := fED_DATEPERIODDEBOnExit;
  THEdit(GetControl('DATEPERIOD_')).OnExit    := fED_DATEPERIODFINOnExit;
  if getparamsoc('SO_AFPLANDECHARGE') then
  begin
    setControlEnabled('AFM_UNITETEMPS', False);
    setcontrolvisible('PREGLES',False);
    SetControlVisible('AFM_MODESAISIEPDC', True);
    SetControlVisible('TAFM_MODESAISIEPDC', True);
  end;
  {$IFDEF CCS3}
  if (getcontrol('GBVALEUR') <> Nil) then SetControlVisible ('GBVALEUR', False);
  if (getcontrol('GBTEXTES') <> Nil)  then SetControlVisible ('GBTEXTES', False);
  if (getcontrol('GBDECISION') <> Nil)  then SetControlVisible ('GBDECISION', False);
  if (getcontrol('GBDATES') <> Nil) then SetControlVisible ('GBDATES', False);
  {$ENDIF}
end;

procedure TOM_AFMODELETACHE.OnNewRecord ;
var vStTypeArticle  : String;
    vStArticle      : String;
    vStCodeArticle  : String;
    vStFacturable   : String;
begin
  Inherited ;
  SetField('AFM_PERIODICITE', 'M');
  SetField('AFM_DATEDEBPERIOD', idate1900);
  SetField('AFM_DATEFINPERIOD', idate2099);
  SetField('AFM_MOISSEMAINE', '1');
  SetField('AFM_MODEGENE', '1');
  SetField('AFM_QTEINTERVENT', '1');

  // article par défault
  SetField('AFM_CODEARTICLE', getparamsoc('SO_AFPRESTDEFAUT'));
  vStCodeArticle := getparamsoc('SO_AFPRESTDEFAUT');
  if controleCodeArticle(vStCodeArticle, vStTypeArticle, vStArticle, vStFacturable) then
    begin
      SetField('AFM_TYPEARTICLE', vStTypeArticle);
      SetField('AFM_ARTICLE', vStArticle);
    end
  else
    begin
      SetField('AFM_CODEARTICLE','');
      SetField('AFM_TYPEARTICLE', '');
      SetField('AFM_ARTICLE', '');
    end;

  // pour l'instant, on initialise avec jour comme unité de saisie
  SetField('AFM_UNITETEMPS', 'J');
  SetField('AFM_FAMILLETACHE', getparamsoc('SO_AFFAMILLEDEF'));
  fRB_MODEGENE1.checked := true;
  fRB_QUOTIDIENNE.checked := true;
  fRB_MOISMETHODE1.checked := true;  
  InitNewRegles;
  Setfield ('AFM_MOISMETHODE', '1');
  Setfield ('AFM_MOISFIXE', '1');
  SetField ('AFM_MOISJOURLIB','J1');
  SetField ('AFM_MOISJOURFIXE', '1');  
  SetControlEnabled('AFM_MODESAISIEPDC', getparamsoc('SO_AFPLANDECHARGE'));
  SetField ('AFM_MODESAISIEPDC','QUA');
  fRB_QUOTIDIENNEOnClick (self);
  SetActiveTabSheet('PGeneral');
  SetFocusControl('AFM_MODELETACHE');
end;

procedure TOM_AFMODELETACHE.InitNewRegles;
begin
  THValComboBox(GetControl('AFM_MOISSEMAINE1')).Value := '1';
  THValComboBox(GetControl('AFM_MOISSEMAINE2')).Value := '1';
  SetControlText('AFM_MOISFIXE0', '1');
  SetControlText('AFM_MOISFIXE1', '1');
  SetControlText('AFM_MOISFIXE2', '1');
  TCheckBox(GetControl('AFM_JOUR1H')).checked := false;
  TCheckBox(GetControl('AFM_JOUR2H')).checked := false;
  TCheckBox(GetControl('AFM_JOUR3H')).checked := false;
  TCheckBox(GetControl('AFM_JOUR4H')).checked := false;
  TCheckBox(GetControl('AFM_JOUR5H')).checked := false;
  TCheckBox(GetControl('AFM_JOUR6H')).checked := false;
  TCheckBox(GetControl('AFM_JOUR7H')).checked := false;
  TCheckBox(GetControl('AFM_JOUR1M')).checked := false;
  TCheckBox(GetControl('AFM_JOUR2M')).checked := false;
  TCheckBox(GetControl('AFM_JOUR3M')).checked := false;
  TCheckBox(GetControl('AFM_JOUR4M')).checked := false;
  TCheckBox(GetControl('AFM_JOUR5M')).checked := false;
  TCheckBox(GetControl('AFM_JOUR6M')).checked := false;
  TCheckBox(GetControl('AFM_JOUR7M')).checked := false;
  SetControlEnabled('RB_MODEGENE1', getParamSoc('SO_AFREALPLAN'));
  SetControlEnabled('RB_MODEGENE2', getParamSoc('SO_AFREALPLAN'));

end;

procedure TOM_AFMODELETACHE.OnLoadRecord ;
begin
  Inherited ;
  InitNewRegles;
  if (Getfield('AFM_MODEGENE') = 1) then
    fRB_MODEGENE1.Checked := True // au plus tôt
  else
    fRB_MODEGENE2.Checked := True;// au plus tard
    
  if Getfield('AFM_PERIODICITE') = 'Q' then fRB_QUOTIDIENNE.Checked := true
  else if Getfield('AFM_PERIODICITE') = 'NBI' then fRB_NBINTERVENTION.Checked := true
  else if Getfield('AFM_PERIODICITE') = 'S' then
  begin
    fRB_HEBDOMADAIRE.Checked := true;
    TCheckBox(GetControl('AFM_JOUR1H')).checked := (Getfield('AFM_JOUR1') = 'X');
    TCheckBox(GetControl('AFM_JOUR2H')).checked := (Getfield('AFM_JOUR2') = 'X');
    TCheckBox(GetControl('AFM_JOUR3H')).checked := (Getfield('AFM_JOUR3') = 'X');
    TCheckBox(GetControl('AFM_JOUR4H')).checked := (Getfield('AFM_JOUR4') = 'X');
    TCheckBox(GetControl('AFM_JOUR5H')).checked := (Getfield('AFM_JOUR5') = 'X');
    TCheckBox(GetControl('AFM_JOUR6H')).checked := (Getfield('AFM_JOUR6') = 'X');
    TCheckBox(GetControl('AFM_JOUR7H')).checked := (Getfield('AFM_JOUR7') = 'X');
  end
  else if Getfield('AFM_PERIODICITE') = 'A' then fRB_ANNUELLE.Checked := true
  else if Getfield('AFM_PERIODICITE') = 'M' then fRB_MENSUELLE.Checked := true;
  if Getfield('AFM_MOISMETHODE') =  '1' then
  begin
    fRB_MOISMETHODE1.Checked := true;
    SetControlText('AFM_MOISFIXE0', Getfield('AFM_MOISFIXE'));
  end
  else if Getfield('AFM_MOISMETHODE') = '2' then
  begin
    fRB_MOISMETHODE2.Checked := true;
    THValComboBox(GetControl('AFM_MOISSEMAINE1')).value := Getfield('AFM_MOISSEMAINE');
    SetControlText('AFM_MOISFIXE1', Getfield('AFM_MOISFIXE'));
  end
  else if Getfield('AFM_MOISMETHODE') = '3' then
  begin
    fRB_MOISMETHODE3.Checked := true;
    THValComboBox(GetControl('AFM_MOISSEMAINE2')).value := Getfield('AFM_MOISSEMAINE');
    SetControlText('AFM_MOISFIXE2', Getfield('AFM_MOISFIXE'));
    TCheckBox(GetControl('AFM_JOUR1M')).checked := (Getfield('AFM_JOUR1') = 'X');
    TCheckBox(GetControl('AFM_JOUR2M')).checked := (Getfield('AFM_JOUR2') = 'X');
    TCheckBox(GetControl('AFM_JOUR3M')).checked := (Getfield('AFM_JOUR3') = 'X');
    TCheckBox(GetControl('AFM_JOUR4M')).checked := (Getfield('AFM_JOUR4') = 'X');
    TCheckBox(GetControl('AFM_JOUR5M')).checked := (Getfield('AFM_JOUR5') = 'X');
    TCheckBox(GetControl('AFM_JOUR6M')).checked := (Getfield('AFM_JOUR6') = 'X');
    TCheckBox(GetControl('AFM_JOUR7M')).checked := (Getfield('AFM_JOUR7') = 'X');
  end;
  SetControlText('DATEPERIOD' ,GetField('AFM_DATEDEBPERIOD'));
  SetControlText('DATEPERIOD_',GetField('AFM_DATEFINPERIOD'));
end ;

procedure TOM_AFMODELETACHE.OnUpdateRecord ;
var vStCodeArticle,vStArticle,vStTypeArticle,vStFacturable :string;
begin
  Inherited ;
  SetActiveTabSheet('PGENERAL');
  vStCodeArticle := trim(GetField ('AFM_CODEARTICLE'));
  vStTypeArticle := GetField ('AFM_TYPEARTICLE');
  if strToDate(getControlText('DATEPERIOD')) >  strToDate(getControlText('DATEPERIOD_')) then
  begin
   LastError    := 1;
   LastErrorMsg := TraduitGa (TexteMsgTache [LastError]);
   SetFocusControl ('DATEPERIOD_');
   exit;
  end
  else if vStCodeArticle = '' then
  begin
   LastError    := 2;
   LastErrorMsg := TraduitGa (TexteMsgTache [LastError]);
   SetFocusControl ('AFM_CODEARTICLE');
   exit;
  end
  else if not controleCodeArticle(vStCodeArticle, vStTypeArticle, vStArticle, vStFacturable) then
  begin
   LastError    := 3;
   LastErrorMsg := TraduitGa (TexteMsgTache [LastError]);
   SetFocusControl ('AFM_CODEARTICLE');
   exit;
  end
  else if GetField ('AFM_FAMILLETACHE') = '' then
  begin
   LastError    := 5;
   LastErrorMsg := TraduitGa (TexteMsgTache [LastError]);
   SetFocusControl ('AFM_FAMILLETACHE');
   exit;
  end
  else if (GetField('AFM_QTEINTERVENT') <= 0) and not getparamsoc('SO_AFPLANDECHARGE') then
  begin
   LastError    := 4;
   LastErrorMsg := TraduitGa (TexteMsgTache [LastError]);
   SetFocusControl ('AFM_QTEINTERVENT');
   exit;
  end;

  SetField('AFM_QTEINITUREF', ConversionUnite(getfield('AFM_UNITETEMPS'), getparamsoc('SO_AFMESUREACTIVITE'), valeur(GetField('AFM_QTEINITIALE'))));
  SetField('AFM_ARTICLE', vStArticle);
  SetField('AFM_TYPEARTICLE', vStTypeArticle);
  SetField('AFM_DATEDEBPERIOD', StrToDate(GetControlText('DATEPERIOD')));
  SetField('AFM_DATEFINPERIOD', StrToDate(GetControlText('DATEPERIOD_')));  
  if fRB_HEBDOMADAIRE.Checked then
  begin
    if TCheckBox(GetControl('AFM_JOUR1H')).Checked then Setfield('AFM_JOUR1', 'X') else Setfield('AFM_JOUR1', '-');
    if TCheckBox(GetControl('AFM_JOUR2H')).Checked then Setfield('AFM_JOUR2', 'X') else Setfield('AFM_JOUR2', '-');
    if TCheckBox(GetControl('AFM_JOUR3H')).Checked then Setfield('AFM_JOUR3', 'X') else Setfield('AFM_JOUR3', '-');
    if TCheckBox(GetControl('AFM_JOUR4H')).Checked then Setfield('AFM_JOUR4', 'X') else Setfield('AFM_JOUR4', '-');
    if TCheckBox(GetControl('AFM_JOUR5H')).Checked then Setfield('AFM_JOUR5', 'X') else Setfield('AFM_JOUR5', '-');
    if TCheckBox(GetControl('AFM_JOUR6H')).Checked then Setfield('AFM_JOUR6', 'X') else Setfield('AFM_JOUR6', '-');
    if TCheckBox(GetControl('AFM_JOUR7H')).Checked then Setfield('AFM_JOUR7', 'X') else Setfield('AFM_JOUR7', '-');
  end
  else if fRB_MENSUELLE.Checked then
  begin
    if fRB_MOISMETHODE1.Checked then
      begin
        Setfield('AFM_MOISMETHODE', '1');
        Setfield('AFM_MOISFIXE', GetControlText('AFM_MOISFIXE0'));
      end
    else if fRB_MOISMETHODE2.Checked then
      begin
        Setfield('AFM_MOISMETHODE', '2');
        Setfield('AFM_MOISSEMAINE', THValComboBox(GetControl('AFM_MOISSEMAINE1')).value);
        Setfield('AFM_MOISFIXE', GetControlText('AFM_MOISFIXE1'));
      end
    else if fRB_MOISMETHODE3.Checked then
      begin
        Setfield('AFM_MOISMETHODE', '3');
        Setfield('AFM_MOISSEMAINE', THValComboBox(GetControl('AFM_MOISSEMAINE2')).value);
        Setfield('AFM_MOISFIXE', GetControlText('AFM_MOISFIXE2'));

        if TCheckBox(GetControl('AFM_JOUR1M')).Checked then Setfield('AFM_JOUR1', 'X') else Setfield('AFM_JOUR1', '-');
        if TCheckBox(GetControl('AFM_JOUR2M')).Checked then Setfield('AFM_JOUR2', 'X') else Setfield('AFM_JOUR2', '-');
        if TCheckBox(GetControl('AFM_JOUR3M')).Checked then Setfield('AFM_JOUR3', 'X') else Setfield('AFM_JOUR3', '-');
        if TCheckBox(GetControl('AFM_JOUR4M')).Checked then Setfield('AFM_JOUR4', 'X') else Setfield('AFM_JOUR4', '-');
        if TCheckBox(GetControl('AFM_JOUR5M')).Checked then Setfield('AFM_JOUR5', 'X') else Setfield('AFM_JOUR5', '-');
        if TCheckBox(GetControl('AFM_JOUR6M')).Checked then Setfield('AFM_JOUR6', 'X') else Setfield('AFM_JOUR6', '-');
        if TCheckBox(GetControl('AFM_JOUR7M')).Checked then Setfield('AFM_JOUR7', 'X') else Setfield('AFM_JOUR7', '-');
      end;
  end;
end ;

procedure TOM_AFMODELETACHE.fED_DATEPERIODDEBOnExit(SEnder: TObject);
begin
  if TFFiche(Ecran).TypeAction=taConsult then exit ;
  ForceUpdate;
  SetField('AFM_DATEDEBPERIOD', StrToDate(GetControlText('DATEPERIOD')));
end;

procedure TOM_AFMODELETACHE.fED_DATEPERIODFINOnExit(SEnder: TObject);
begin
  if TFFiche(Ecran).TypeAction=taConsult then exit ;
  ForceUpdate;
  SetField('AFM_DATEFINPERIOD',StrToDate(GetControlText('DATEPERIOD_')));
end;

procedure TOM_AFMODELETACHE.fRB_ANNUELLEOnClick(Sender: TObject);
begin
  if (Getfield ('AFM_PERIODICITE') <> 'A') and not(DS.State in [dsInsert,dsEdit]) then DS.edit;
  Setfield('AFM_PERIODICITE', 'A');
  fRB_HEBDOMADAIRE.checked := not fRB_ANNUELLE.checked;
  fRB_QUOTIDIENNE.checked := not fRB_ANNUELLE.checked;
  fRB_MENSUELLE.checked := not fRB_ANNUELLE.checked;
  fRB_NBINTERVENTION.checked := not fRB_ANNUELLE.checked;
  fPC_FREQUENCE.ActivePage := fTS_ANNUELLE;
  SetControlEnabled('RB_MODEGENE2', false);
  SetControlEnabled('RB_MODEGENE1', false);
end;

procedure TOM_AFMODELETACHE.fRB_HEBDOMADAIREOnClick(Sender: TObject);
begin
  if (Getfield ('AFM_PERIODICITE') <> 'S') and not(DS.State in [dsInsert,dsEdit]) then DS.edit;
  Setfield('AFM_PERIODICITE', 'S');
  fRB_QUOTIDIENNE.checked := not fRB_HEBDOMADAIRE.checked;
  fRB_MENSUELLE.checked := not fRB_HEBDOMADAIRE.checked;
  fRB_ANNUELLE.checked := not fRB_HEBDOMADAIRE.checked;
  fRB_NBINTERVENTION.checked := not fRB_HEBDOMADAIRE.checked;
  fPC_FREQUENCE.ActivePage := fTS_HEBDOMADAIRE;
  SetControlEnabled('RB_MODEGENE1', (GetField('AFM_QTEINITIALE') > 0));
  SetControlEnabled('RB_MODEGENE2', (GetField('AFM_QTEINITIALE') > 0));
end;

procedure TOM_AFMODELETACHE.fRB_MENSUELLEOnClick(Sender: TObject);
begin
  if (Getfield ('AFM_PERIODICITE') <> 'M') and not(DS.State in [dsInsert,dsEdit])  then DS.edit;
  Setfield('AFM_PERIODICITE', 'M');
  fRB_HEBDOMADAIRE.checked := not fRB_MENSUELLE.checked;
  fRB_QUOTIDIENNE.checked := not fRB_MENSUELLE.checked;
  fRB_ANNUELLE.checked := not fRB_MENSUELLE.checked;
  fRB_NBINTERVENTION.checked := not fRB_MENSUELLE.checked;
  fPC_FREQUENCE.ActivePage := fTS_MENSUELLE;
  SetControlEnabled('RB_MODEGENE1', (GetField('AFM_QTEINITIALE') > 0));
  SetControlEnabled('RB_MODEGENE2', (GetField('AFM_QTEINITIALE') > 0));  
end;

procedure TOM_AFMODELETACHE.fRB_NBINTERVENTIONOnClick(Sender: TObject);
begin
  if (Getfield ('AFM_PERIODICITE') <> 'NBI') and not(DS.State in [dsInsert,dsEdit])  then DS.edit;
  Setfield('AFM_PERIODICITE', 'NBI');
  fRB_HEBDOMADAIRE.checked := not fRB_NBINTERVENTION.checked;
  fRB_QUOTIDIENNE.checked := not fRB_NBINTERVENTION.checked;
  fRB_MENSUELLE.checked := not fRB_NBINTERVENTION.checked;
  fRB_ANNUELLE.checked := not fRB_NBINTERVENTION.checked;
  fPC_FREQUENCE.ActivePage := fTS_NBINTERVENTION;
  SetControlEnabled('RB_MODEGENE2', false);
  SetControlEnabled('RB_MODEGENE1', false);  
end;

procedure TOM_AFMODELETACHE.fRB_QUOTIDIENNEOnClick(Sender: TObject);
begin
  if (Getfield ('AFM_PERIODICITE') <> 'Q') and not(DS.State in [dsInsert,dsEdit])  then DS.edit;
  Setfield('AFM_PERIODICITE', 'Q');
  fRB_HEBDOMADAIRE.checked := not fRB_QUOTIDIENNE.checked;
  fRB_MENSUELLE.checked := not fRB_QUOTIDIENNE.checked;
  fRB_ANNUELLE.checked := not fRB_QUOTIDIENNE.checked;
  fRB_NBINTERVENTION.checked := not fRB_QUOTIDIENNE.checked;
  fPC_FREQUENCE.ActivePage := fTS_QUOTIDIENNE;
  SetControlEnabled('RB_MODEGENE1', (GetField('AFM_QTEINITIALE') > 0));
  SetControlEnabled('RB_MODEGENE2', (GetField('AFM_QTEINITIALE') > 0));  
end;

procedure TOM_AFMODELETACHE.fRB_MODEGENE1OnClick(Sender: TObject);
begin
  if (Getfield('AFM_MODEGENE') <> 1) and not(DS.State in [dsInsert,dsEdit]) then DS.edit;
  fRB_MODEGENE2.checked := not fRB_MODEGENE1.Checked;
  Setfield ('AFM_MODEGENE', 1);
end;

procedure TOM_AFMODELETACHE.fRB_MODEGENE2OnClick(Sender: TObject);
begin
  if (Getfield('AFM_MODEGENE') <> 2) and not(DS.State in [dsInsert,dsEdit]) then DS.edit;
  fRB_MODEGENE1.Checked := not fRB_MODEGENE2.Checked;
  Setfield ('AFM_MODEGENE', 2);
end;

procedure TOM_AFMODELETACHE.fRB_MOISMETHODE1OnClick(SEnder: TObject);
begin
  if (Getfield('AFM_MOISMETHODE') <> '1') and not(DS.State in [dsInsert,dsEdit]) then DS.edit;
  fRB_MOISMETHODE2.checked := not fRB_MOISMETHODE1.checked;
  fRB_MOISMETHODE3.checked := not fRB_MOISMETHODE1.checked;
  fRB_MOISMETHODE1.TabStop := fRB_MOISMETHODE1.checked;
  fRB_MOISMETHODE2.TabStop := not fRB_MOISMETHODE1.checked;
  fRB_MOISMETHODE3.TabStop := not fRB_MOISMETHODE1.checked;
  SetControlEnabled('AFM_MOISJOURFIXE', True);
  SetControlEnabled('AFM_MOISFIXE0', True);

  SetControlEnabled('AFM_MOISSEMAINE1', false);
  SetControlEnabled('AFM_MOISJOURLIB', false);
  SetControlEnabled('AFM_MOISFIXE1', false);

  SetControlEnabled('AFM_MOISSEMAINE2', false);
  SetControlEnabled('AFM_MOISFIXE2', false);
  SetControlEnabled('AFM_JOUR1M', false);
  SetControlEnabled('AFM_JOUR2M', false);
  SetControlEnabled('AFM_JOUR3M', false);
  SetControlEnabled('AFM_JOUR4M', false);
  SetControlEnabled('AFM_JOUR5M', false);
  SetControlEnabled('AFM_JOUR6M', false);
  SetControlEnabled('AFM_JOUR7M', false);
end;

procedure TOM_AFMODELETACHE.fRB_MOISMETHODE2OnClick(SEnder: TObject);
begin
  if (Getfield('AFM_MOISMETHODE') <> '2') and not(DS.State in [dsInsert,dsEdit]) then DS.edit;
  fRB_MOISMETHODE1.checked := not fRB_MOISMETHODE2.checked;
  fRB_MOISMETHODE3.checked := not fRB_MOISMETHODE2.checked;
  fRB_MOISMETHODE1.TabStop := not fRB_MOISMETHODE2.checked;
  fRB_MOISMETHODE2.TabStop := fRB_MOISMETHODE2.checked;
  fRB_MOISMETHODE3.TabStop := not fRB_MOISMETHODE2.checked;

  SetControlEnabled('AFM_MOISJOURFIXE', false);
  SetControlEnabled('AFM_MOISFIXE0', false);

  SetControlEnabled('AFM_MOISSEMAINE1', true);
  SetControlEnabled('AFM_MOISJOURLIB', true);
  SetControlEnabled('AFM_MOISFIXE1', true);

  SetControlEnabled('AFM_MOISSEMAINE2', false);
  SetControlEnabled('AFM_MOISFIXE2', false);
  SetControlEnabled('AFM_JOUR1M', false);
  SetControlEnabled('AFM_JOUR2M', false);
  SetControlEnabled('AFM_JOUR3M', false);
  SetControlEnabled('AFM_JOUR4M', false);
  SetControlEnabled('AFM_JOUR5M', false);
  SetControlEnabled('AFM_JOUR6M', false);
  SetControlEnabled('AFM_JOUR7M', false);
end;

procedure TOM_AFMODELETACHE.fRB_MOISMETHODE3OnClick(SEnder: TObject);
begin
  if (Getfield('AFM_MOISMETHODE') <> '3') and not(DS.State in [dsInsert,dsEdit]) then DS.edit;
  fRB_MOISMETHODE1.checked := not fRB_MOISMETHODE3.checked;
  fRB_MOISMETHODE2.checked := not fRB_MOISMETHODE3.checked;
  fRB_MOISMETHODE1.TabStop := not fRB_MOISMETHODE3.checked;
  fRB_MOISMETHODE2.TabStop := not fRB_MOISMETHODE3.checked;
  fRB_MOISMETHODE3.TabStop := fRB_MOISMETHODE3.checked;
  SetControlEnabled('AFM_MOISJOURFIXE', false);
  SetControlEnabled('AFM_MOISFIXE0', false);

  SetControlEnabled('AFM_MOISSEMAINE1', false);
  SetControlEnabled('AFM_MOISJOURLIB', false);
  SetControlEnabled('AFM_MOISFIXE1', false);
                                     
  SetControlEnabled('AFM_MOISSEMAINE2', true);
  SetControlEnabled('AFM_MOISFIXE2', true);
  SetControlEnabled('AFM_JOUR1M', true);
  SetControlEnabled('AFM_JOUR2M', true);
  SetControlEnabled('AFM_JOUR3M', true);
  SetControlEnabled('AFM_JOUR4M', true);
  SetControlEnabled('AFM_JOUR5M', true);
  SetControlEnabled('AFM_JOUR6M', true);
  SetControlEnabled('AFM_JOUR7M', true);
end;

procedure TOM_AFMODELETACHE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_AFMODELETACHE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_AFMODELETACHE.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_AFMODELETACHE.OnDeleteRecord ;
begin
  Inherited ;
  if ExisteSQL ( 'SELECT ATA_AFFAIRE FROM TACHE WHERE ATA_MODELETACHE="'+GetField('AFM_MODELETACHE')+'"') and
  ( PGIAskAF (TexteMsgTache[6],Ecran.Caption) <> mrYes ) then
  begin
   LastError := 6;
   exit;
  end;
  
end ;

procedure TOM_AFMODELETACHE.OnCancelRecord ;
begin
  Inherited ;
end ;


Initialization
  registerclasses ( [ TOM_AFMODELETACHE ] ) ;
end.
