{***********UNITE*************************************************
Auteur  ...... : NA
Créé le ...... : 15/06/2007
Modifié le ... :   /  /
Description .. : Source TOM de la FICHE : JOURNEETYPE ()

Mots clefs ... : UTOMJOURNEETYPE
*****************************************************************
PT1  20/07/2007  FLO  Paramétrage des droits
PT2  10/08/2007  FLO  Recalcul automatique des compteurs lors d'une modification ou d'une suppression
PT3  27/08/2007  FLO  Date de recalcul = date de début de période
}
unit UTOMJourneeType;

interface
uses  classes,sysutils,controls,Dialogs,Graphics, StdCtrls,
      {$IFNDEF EAGLCLIENT}
      db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB,Fiche,DBCtrls,Fe_Main,  //Fiche,
      {$ELSE}
      eFiche, Utob, MaineAgl,
      {$ENDIF}
      UTOM, HCtrls, HTB97,HMSgBox, uTableFiltre,SaisieList,Menus;

Type
TOM_JOURNEETYPE = Class(TOM)
  procedure OnArgument(stArgument: string); override;
  procedure OnNewRecord;                    override;
  procedure OnDeleteRecord;                 override;
  procedure OnUpdateRecord                ; override;
  procedure OnAfterUpdateRecord;            override; //PT2
  procedure OnAfterDeleteRecord;            override; //PT2
  procedure OnLoadRecord;                   override;
  Procedure OnChangeField( F : TField);     override;
  Private
    savcolortype,savtypejour,savhordebplage1, savhorfinplage1,savjourj1plage1,savtypejour1,
    savhordebplage2,savhorfinplage2,savjourj1plage2,savtypejour2,savnotrav,savjourtravferie,
    savnotraveff,savpause,savpaiementpause,savdureetraveff,savpauseeff,savcolorjourtype,savBDhordebplage1,
    savBDhorfinplage1,savBDhordebplage2,savBDhorfinplage2,Savdureeplage1,Savdureeplage2,savdureenotraveff ,
    savdureenotrav,savequivalenttpsplein,savpoidsjour,savtempslibre2,savtempslibre3,savtempslibre4,
    savtempslibre5,savtempslibre6,savtempslibre7 ,Savdureepause   : string;
    Journee     : String;    //PT2
    VerifImpact : Boolean;   //PT2
    DateModif   : TDateTime; //PT3
    procedure OnExitEdit(Sender : TObject);
    procedure Calcultraveff(Sender : TObject);
    Function  PgCalculDureeTime( H1,H2 : Double) : Double;
    procedure ClickPgColor (Sender : TObject);
    procedure ClickPgDefaire (Sender : TObject);
    procedure DupliquerDroitjournee (Sender : TObject);
    procedure DupliquerJournee (Sender : TObject);
    procedure AccesDroitjourneetype (Sender : TObject);
    procedure Sauvegardechamp;
    procedure chargechamp;
  End;

TOM_DROITJOURNEETYPE = Class(TOM)
  procedure OnArgument(stArgument: string); override ;
  procedure OnNewRecord                   ; override ;
  procedure OnUpdateRecord                ; override ;
  procedure OnDeleteRecord                ; override ;  //PT1
  procedure OnAfterUpdateRecord;            override ;  //PT2
  procedure OnAfterDeleteRecord;            override ;  //PT2
  procedure OnCancelRecord                ; override ;  //PT1
  Private
    TF : TTableFiltre;
    Journee : String;    //PT2
    DateModif   : TDateTime; //PT3
  End;



implementation
uses HeureUtil, HEnt1                                       
  ,PGPresence //PT2
  ;



{ TOM_JOURNEETYPE }

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 15/06/2007
Modifié le ... :   /  /    
Description .. : Nouvelle saisie
Mots clefs ... : 
*****************************************************************}
procedure TOM_JOURNEETYPE.OnNewRecord;
begin
  inherited;
setfield('PJO_JOURNEETYPE', '');
SetField('PJO_COLORJOURTYPE','$00804000');
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 15/06/2007
Modifié le ... :   /  /    
Description .. : On argument
Mots clefs ... : 
*****************************************************************}
procedure TOM_JOURNEETYPE.OnArgument(stArgument: string);
var
Edit : THEdit;
Btn  : TToolBarButton97;
pauseeffect : Tcheckbox;
MenuDupli : TPopupMenu;

begin
  inherited;

  Edit := THEdit(GetControl('NOTRAV'));
  if Edit<>nil then Edit.OnExit := OnExitEdit;
  Edit := THEdit(GetControl('NOTRAVEFF'));
  if Edit<>nil then Edit.OnExit := OnExitEdit;
  Edit := THEdit(GetControl('PAUSE'));
  if Edit<>nil then Edit.OnExit := OnExitEdit;
  Edit := THEdit(GetControl('HORDEBPLAGE1'));
  if Edit<>nil then Edit.OnExit := OnExitEdit;
  Edit := THEdit(GetControl('HORFINPLAGE1'));
  if Edit<>nil then Edit.OnExit := OnExitEdit;
  Edit := THEdit(GetControl('HORDEBPLAGE2'));
  if Edit<>nil then Edit.OnExit := OnExitEdit;
  Edit := THEdit(GetControl('HORFINPLAGE2'));
  if Edit<>nil then Edit.OnExit := OnExitEdit;

  Btn := TToolBarButton97(GetControl('BTNPGCOLORJOURTYPE'));
  if Assigned(Btn) then Btn.OnClick := ClickPgColor;

  Btn := TToolBarButton97(GetControl('BTNDEFAIREJOURTYPE'));
  if Assigned(Btn) then Btn.OnClick := ClickPgDefaire;

  Btn := TToolBarButton97(GetControl('DROITJOURNEETYPE'));
  if Assigned(Btn) then Btn.OnClick := AccesDroitjourneetype;

  pauseeffect := TCheckBox(Getcontrol('PJO_PAUSEEFFECTIF'));
  if Assigned(Pauseeffect) then
  pauseeffect.OnExit := calcultraveff;


  MenuDupli := TPopupMenu(GetControl('DUPLICJOURNEE'));
  if MenuDupli <> nil then
  begin
      MenuDupli.Items[0].OnClick := DupliquerDROITJournee;
      MenuDupli.Items[1].OnClick := DupliquerJournee;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 09/08/2007 / PT2
Modifié le ... :   /  /
Description .. : Création ou modification d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_JOURNEETYPE.OnUpdateRecord;
var DD,DF : TDateTime;
begin
  inherited;
     If (LastError = 0) Then
     Begin
          VerifImpact := (ds.State = dsEdit);
          Journee     := GetField('PJO_JOURNEETYPE');
          PresenceDonneMoisCalculActuel (DD,DF);   //PT3
          DateModif   := DD;
     End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 18/06/2007
Modifié le ... :   /  /
Description .. : Suppression journée type
Mots clefs ... :
*****************************************************************}
procedure TOM_JOURNEETYPE.OnDeleteRecord;
var
journeetype : string;
DD,DF : TDateTime;
begin
  inherited;
     journeetype := getfield('PJO_JOURNEETYPE');
     if existesql('SELECT PMO_JOURNEETYPE FROM MODELECYCLE WHERE PMO_JOURNEETYPE = "'+journeetype+'"') then
     begin
          PGIBOX('Suppression impossible: cette journée type est utilisé dans un modèle cycle', Ecran.Caption);
          Lasterror := 1;
          exit;
     end;

     // suppression des droits journée type
     if Lasterror = 0 then ExecuteSQL('DELETE FROM DROITJOURNEETYPE WHERE PDJ_JOURNEETYPE ="' + journeetype +'"');

     If (LastError = 0) Then
     Begin
          Journee := JourneeType; //PT2
          PresenceDonneMoisCalculActuel (DD,DF);   //PT3
          DateModif := DD;
     End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 09/08/2007 / PT2
Modifié le ... :   /  /
Description .. : Suite à la création ou modification d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_JOURNEETYPE.OnAfterUpdateRecord;
begin
  inherited;
     If (LastError = 0) And VerifImpact And (PresenceJourneyIsUsed(Journee)) Then
          CompteursARecalculer(DateModif);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 09/08/2007 / PT2
Modifié le ... :   /  /
Description .. : Suite à la suppression d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_JOURNEETYPE.OnAfterDeleteRecord;
begin
  inherited;
     If (LastError = 0) And (PresenceJourneyIsUsed(Journee)) Then
          CompteursARecalculer(DateModif);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 15/06/2007
Modifié le ... :   /  /    
Description .. : Chargement de la fiche 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JOURNEETYPE.OnLoadRecord;
Var
  EditColor : THEdit;
  {$IFDEF EAGLCLIENT}
  Edit : THEdit;
  {$ELSE}
  Edit : THDBEdit;
  {$ENDIF}
begin
  inherited;

    SetControlText('NOTRAV',FloatToStrTime(GetField('PJO_DUREENOTRAV'),'hh:nn',True));
    SetControlText('NOTRAVEFF',FloatToStrTime(GetField('PJO_DUREENOTRAVEFF'),'hh:nn',True));
    SetControlText('PAUSE',FloatToStrTime(GetField('PJO_DUREEPAUSE'),'hh:nn',True));

    SetControlText('HORDEBPLAGE1',FormatDateTime('hh:mm',GetField('PJO_HORDEBPLAGE1')));
    SetControlText('HORFINPLAGE1',FormatDateTime('hh:mm',GetField('PJO_HORFINPLAGE1')));
    SetControlText('HORDEBPLAGE2',FormatDateTime('hh:mm',GetField('PJO_HORDEBPLAGE2')));
    SetControlText('HORFINPLAGE2',FormatDateTime('hh:mm',GetField('PJO_HORFINPLAGE2')));

    calcultraveff(nil);

  {$IFDEF EAGLCLIENT}
  Edit := THEdit(GetControl('PJO_COLORJOURTYPE'));
  {$ELSE}
  Edit := THDBEdit(GetControl('PJO_COLORJOURTYPE'));
  {$ENDIF}
  If Assigned(Edit) then
    Begin
    if Edit.text<>'' then
      Begin
      Edit.Font.Color := StringToColor(Edit.Text);
      EditColor := THEdit(GetControl('PGCOLORJOURTYPE'));
      if Assigned(EditColor) then
        Begin
        EditColor.Color := StringToColor(Edit.Text);
        EditColor.Enabled := False;
        End;
      End;
    End;

end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 18/06/2007
Modifié le ... :   /  /    
Description .. : Dupliquer journée type
Mots clefs ... : 
*****************************************************************}
procedure TOM_JOURNEETYPE.DupliquerJournee(Sender: TObject);
var
journeetype, mode, champsduplic, journeedup,libelle , appel: string;

Begin
  mode := 'ENT';
  Appel := 'JTYP';
  sauvegardechamp;
  journeetype := getfield('PJO_JOURNEETYPE');
  champsduplic := AglLanceFiche('PAY', 'PARAMPRESENCEDUP', '', '',Appel + ';' + journeetype + ';' + mode);

  if champsduplic <> '' then
  begin

    journeedup := readtokenst(champsduplic);
    libelle := readtokenst(champsduplic);

    (GetControl('BINSERT')  as TToolbarButton97).OnClick(Sender);
    setfield('PJO_JOURNEETYPE', journeedup);
    setfield('PJO_LIBELLE', libelle);
    chargechamp;
    TFFiche(Ecran).BValider.Click;
  end;

end;



{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 18/06/2007
Modifié le ... :   /  /    
Description .. : Dupliquer journée type et les droits de la journée
Mots clefs ... : 
*****************************************************************}
procedure TOM_JOURNEETYPE.DupliquerDroitJournee(Sender: TObject);
var
journeetype, mode, Appel, champsduplic, journeedup, libelle : string;

Begin
  mode := 'DET';
  Appel := 'JTYP';
  sauvegardechamp;
  journeetype := getfield('PJO_JOURNEETYPE');
  champsduplic := AglLanceFiche('PAY', 'PARAMPRESENCEDUP', '', '',Appel + ';' + journeetype + ';' + mode);

  if champsduplic <> '' then
  begin

    journeedup := readtokenst(champsduplic);
    libelle := readtokenst(champsduplic);

    (GetControl('BINSERT')  as TToolbarButton97).OnClick(Sender);
    setfield('PJO_JOURNEETYPE', journeedup);
    setfield('PJO_LIBELLE', libelle);
    chargechamp;
    TFFiche(Ecran).BValider.Click;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 19/06/2007
Modifié le ... :   /  /    
Description .. : Sauvegarde des champs de la  fiche 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JOURNEETYPE.sauvegardechamp;
begin
savcolortype := getcontroltext('PGCOLORJOURTYPE');
savtypejour := getfield('PJO_TYPEJOUR');
savhordebplage1 := getcontroltext('HORDEBPLAGE1');
savhorfinplage1 := getcontroltext('HORFINPLAGE1');
savjourj1plage1 := getfield('PJO_JOURJ1PLAGE1');
savtypejour1 := getfield('PJO_TYPEJOUR1');
savhordebplage2 := getcontroltext('HORDEBPLAGE2');
savhorfinplage2 := getcontroltext('HORFINPLAGE2');
savjourj1plage2 := getfield('PJO_JOURJ1PLAGE2');
savtypejour2 := getfield('PJO_TYPEJOUR2');
savnotrav := getcontroltext('NOTRAV');
savjourtravferie := getfield('PJO_JOURTRAVFERIE');
savnotraveff := getcontroltext('PJO_NOTRAVEFF');
savpause := getcontroltext('PAUSE');
savpaiementpause := getfield('PJO_PAIEMENTPAUSE');
savdureetraveff := getcontroltext('DUREETRAVEFF');
savpauseeff := getfield('PJO_PAUSEEFFECTIF');
savcolorjourtype := getfield('PJO_COLORJOURTYPE');
savBDhordebplage1 := getcontroltext('PJO_HORDEBPLAGE1');
savBDhorfinplage1 := getcontroltext('PJO_HORFINPLAGE1');
savBDhordebplage2 := getcontroltext('PJO_HORDEBPLAGE2');
savBDhorfinplage2 := getcontroltext('PJO_HORFINPLAGE2');
Savdureeplage1 := getfield('PJO_DUREEPLAGE1');
Savdureeplage2 := getfield('PJO_DUREEPLAGE2');
Savdureepause := getfield('PJO_DUREEPAUSE');
savdureenotraveff := getfield('PJO_DUREENOTRAVEFF');
savdureenotrav := getfield('PJO_DUREENOTRAV');
savequivalenttpsplein := getfield('PJO_EQUIVTPSPLEIN');
savpoidsjour := getfield('PJO_POIDSJOUR');
savtempslibre2:= getfield('PJO_TEMPSLIBRE2');
savtempslibre3:= getfield('PJO_TEMPSLIBRE3');
savtempslibre4:= getfield('PJO_TEMPSLIBRE4');
savtempslibre5:= getfield('PJO_TEMPSLIBRE5');
savtempslibre6:= getfield('PJO_TEMPSLIBRE6');
savtempslibre7:= getfield('PJO_TEMPSLIBRE7');
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 19/06/2007
Modifié le ... :   /  /    
Description .. : Chargement des champs de la journée type
Mots clefs ... : 
*****************************************************************}
procedure TOM_JOURNEETYPE.chargechamp;
begin
  setcontroltext('PGCOLORJOURTYPE',savcolortype);
  setfield('PJO_TYPEJOUR',savtypejour );
  setcontroltext('HORDEBPLAGE1',savhordebplage1);
  setcontroltext('HORFINPLAGE1',savhorfinplage1);
  setfield('PJO_JOURJ1PLAGE1',savjourj1plage1);
  setfield('PJO_TYPEJOUR1',savtypejour1);
  setcontroltext('HORDEBPLAGE2',savhordebplage2);
  setcontroltext('HORFINPLAGE2',savhorfinplage2);
  setfield('PJO_JOURJ1PLAGE2',savjourj1plage2);
  setfield('PJO_TYPEJOUR2',savtypejour2 );
  setcontroltext('NOTRAV',savnotrav);
  setfield('PJO_JOURTRAVFERIE',savjourtravferie );
  setcontroltext('PJO_NOTRAVEFF',savnotraveff );
  setcontroltext('PAUSE',savpause );
  setfield('PJO_PAIEMENTPAUSE',savpaiementpause );
  setcontroltext('DUREETRAVEFF',savdureetraveff);
  setfield('PJO_PAUSEEFFECTIF',savpauseeff);
  setfield('PJO_COLORJOURTYPE',savcolorjourtype);
  setcontroltext('PJO_HORDEBPLAGE1',savBDhordebplage1);
  setcontroltext('PJO_HORFINPLAGE1',savBDhorfinplage1);
  setcontroltext('PJO_HORDEBPLAGE2',savBDhordebplage2);
  setcontroltext('PJO_HORFINPLAGE2',savBDhorfinplage2);
  setfield('PJO_DUREEPLAGE1',Savdureeplage1);
  setfield('PJO_DUREEPLAGE2',Savdureeplage2);
  setfield('PJO_DUREEPAUSE',Savdureepause);
  setfield('PJO_DUREENOTRAVEFF',savdureenotraveff);
  setfield('PJO_DUREENOTRAV',savdureenotrav);
  setfield('PJO_EQUIVTPSPLEIN',savequivalenttpsplein );
  setfield('PJO_POIDSJOUR',savpoidsjour);
  setfield('PJO_TEMPSLIBRE2',savtempslibre2 );
  setfield('PJO_TEMPSLIBRE3',savtempslibre3);
  setfield('PJO_TEMPSLIBRE4',savtempslibre4);
  setfield('PJO_TEMPSLIBRE5',savtempslibre5 );
  setfield('PJO_TEMPSLIBRE6',savtempslibre6);
  setfield('PJO_TEMPSLIBRE7',savtempslibre7);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 15/06/2007
Modifié le ... :   /  /    
Description .. : Calcul des durées 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JOURNEETYPE.OnExitEdit(Sender: TObject);
Var
St, Pl : String;
D1,D2,DH : Double;

begin
St := UpperCase(TControl(Sender).Name);
If  (St = 'NOTRAV') or (St = 'NOTRAVEFF') or (St = 'PAUSE') then
  Begin
  DH := StrTimeToFloat(GetControlText(St),True);
  if DH <> GetField('PJO_DUREE'+St) then
    Begin
    If DS.State in [DsBrowse] then DS.Edit;
    SetField('PJO_DUREE'+St, DH);
    End;
  End
else
  If (St = 'HORDEBPLAGE1') or (St = 'HORFINPLAGE1') or (St = 'HORDEBPLAGE2') or (St = 'HORFINPLAGE2') then
    Begin
    Pl := Copy(St,7,Length(st));
    D1 := StrTimeToFloat(GetControlText('HORDEB'+Pl),True);
    D2 := StrTimeToFloat(GetControlText('HORFIN'+Pl),True);
    DH := PgCalculDureeTime(D1,D2);
      //DH := CalculEcartHeure(D1,D2);
    if DH <> GetField('PJO_DUREE'+Pl) then
      Begin
      If DS.State in [DsBrowse] then DS.Edit;
      SetField('PJO_DUREE'+PL,DH);
      end;
    If GetControlText(st) <> GetField('PJO_'+St) then
      Begin
      If DS.State in [DsBrowse] then DS.Edit;
      SetField('PJO_'+St,GetControlText(st));
      End;
    End;

 // calcul durée de travail effective
 calcultraveff(nil);

end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 15/06/2007
Modifié le ... :   /  /    
Description .. : Calcul de la durée de travail effective
Mots clefs ... : 
*****************************************************************}
procedure TOM_JOURNEETYPE.calcultraveff(Sender: TObject);
var
dureetraveff : Double;
pauseeffect : TCheckBox;
begin
  // calcul de la durée de travail effective
  if ds.state <> DSinsert then
  begin
    pauseeffect := TCheckBox(Getcontrol('PJO_PAUSEEFFECTIF'));
//    dureetraveff := 0;
    dureetraveff := getfield('PJO_DUREEPLAGE1') + getfield( 'PJO_DUREEPLAGE2') - getfield('PJO_DUREENOTRAV');
    if not pauseeffect.checked  then
    dureetraveff := dureetraveff - getfield('PJO_DUREEPAUSE');
    setcontroltext('DUREETRAVEFF', FloatToStrTime(dureetraveff,'hh:mm',true));
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 15/06/2007
Modifié le ... :   /  /    
Description .. : on change field
Mots clefs ... : 
*****************************************************************}
procedure TOM_JOURNEETYPE.OnChangeField(F: TField);
{Var
DH,D1,D2 : Double;}
begin
  inherited;
  {If (F.FieldName ='PJO_HORDEBPLAGE1') OR  (F.FieldName = 'PJO_HORFINPLAGE1') then
    Begin
    D1 := StrTimeToFloat(GetControlText('HORDEBPLAGE1'),True); //FormatDateTime('hh:mm',
    D2 := StrTimeToFloat(GetControlText('HORFINPLAGE1'),True);
    if (D1<>0) and (D2<>0) then
      Begin
      DH := CalculEcartHeure(D1,D2);
      if DH <> GetField('PJO_DUREEPLAGE1') then SetField('PJO_DUREEPLAGE1',DH);
      End;
    End;

  If (F.Fieldname = 'PJO_HORDEBPLAGE2') OR  (F.FieldName = 'PJO_HORFINPLAGE2') then
    Begin
    D1 := StrTimeToFloat(GetControlText('HORDEBPLAGE2'),True);
    D2 := StrTimeToFloat(GetControlText('HORFINPLAGE2'),True);
    if (D1<>0) and (D2<>0) then
      Begin
      DH := CalculEcartHeure(D1,D2);
      if DH <> GetField('PJO_DUREEPLAGE2') then SetField('PJO_DUREEPLAGE2',DH);
      End;
    End; }
end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 15/06/2007
Modifié le ... :   /  /    
Description .. : Calcul durée
Mots clefs ... : 
*****************************************************************}
function TOM_JOURNEETYPE.PgCalculDureeTime(H1, H2: Double): Double;
Var
 H : Double;
begin
If H1>H2 then
  Begin
  H := StrTimeToFloat('24:00',True);
  Result := H - H1 + H2 ;
  End
else
  Begin
  Result := H2 - H1;
  End;

end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 15/06/2007
Modifié le ... :   /  /    
Description .. : Click sur la couleur
Mots clefs ... : 
*****************************************************************}
procedure TOM_JOURNEETYPE.ClickPgColor(Sender: TObject);
Var
  Col : String;
  Colors : TColorDialog;
  EditCol : THEdit;
  {$IFDEF EAGLCLIENT}
  Edit : THEdit;
  {$ELSE}
  Edit : THDBEdit;
  {$ENDIF}
begin
Colors:=TColorDialog.Create(Ecran);
if Colors=nil then exit;

if (Colors.Execute) then
  Begin
  Col:=ColorToString(Colors.Color);
  EditCol := THEdit(GetControl('PGCOLORJOURTYPE'));
  {$IFDEF EAGLCLIENT}
  Edit := THEdit(GetControl('PJO_COLORJOURTYPE'));
  {$ELSE}
  Edit := THDBEdit(GetControl('PJO_COLORJOURTYPE'));
  {$ENDIF}
  If Assigned(Edit) then
    Begin
    Edit.Font.Color := Colors.Color;
    if DS.State in [DsBrowse] then Ds.Edit;
    if GetField('PJO_COLORJOURTYPE') <> Col then
       Begin
       SetField('PJO_COLORJOURTYPE',Col);
       if Assigned(EditCol) then EditCol.Color := Colors.Color;
       End;
    End;
  End;
Colors.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 15/06/2007
Modifié le ... :   /  /    
Description .. : Changement de couleur
Mots clefs ... : 
*****************************************************************}
procedure TOM_JOURNEETYPE.ClickPgDefaire(Sender: TObject);
Var
  stVal : String;
  EditCol : THEdit;
  {$IFDEF EAGLCLIENT}
  Edit : THEdit;
  {$ELSE}
  Edit : THDBEdit;
  {$ENDIF}
begin
  stVal := '$00804000';
  EditCol := THEdit(GetControl('PGCOLORJOURTYPE'));
  {$IFDEF EAGLCLIENT}
  Edit := THEdit(GetControl('PJO_COLORJOURTYPE'));
  {$ELSE}
  Edit := THDBEdit(GetControl('PJO_COLORJOURTYPE'));
  {$ENDIF}
  If Assigned(Edit) then
    Begin
    Edit.Font.Color := StringToColor(StVal);
    if DS.State in [DsBrowse] then Ds.Edit;
    if GetField('PJO_COLORJOURTYPE') <> StVal then
      Begin
      SetField('PJO_COLORJOURTYPE',StVal);
      if Assigned(EditCol) then EditCol.Color := StringToColor(StVal);
      End;
    End;



end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 19/06/2007
Modifié le ... :   /  /    
Description .. : Acces à la saisie des droits de la journée type
Mots clefs ... :
*****************************************************************}
procedure TOM_JOURNEETYPE.AccesDroitJourneeType(Sender: TObject);
var
proven, jourtype : string;
begin
proven := 'JTYP';
jourtype := getfield('PJO_JOURNEETYPE');
AglLanceFiche('PAY','DROITJOURTYP1_FSL','','',proven + ';' + jourtype);

end;

{ TOM_DROITJOURNEETYPE }

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 15/06/2007
Modifié le ... :   /  /    
Description .. : On argument droit journée type
Mots clefs ... : 
*****************************************************************}
procedure TOM_DROITJOURNEETYPE.OnArgument(stArgument: string);
begin
  inherited;

     // Si saisie des droits d'une journée, TF = LeFiltre
     // Si paramétrage des droits, TF = Nil
     If (Ecran <> nil ) and (Ecran is TFSaisieList ) And (Ecran.Name<>'DROITJOURNEE_FSL') then //PT1
          TF := TFSaisieList(Ecran).LeFiltre
     Else
          TF := Nil;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 15/06/2007
Modifié le ... : 20/07/2007 /PT1
Description .. : Nouvelle saisie droit journée type
Mots clefs ... : 
*****************************************************************}
procedure TOM_DROITJOURNEETYPE.OnNewRecord;
begin
  inherited;
     if TF<>nil then
     Begin
          SetField('PDJ_JOURNEETYPE',TF.TOBFiltre.GetValue('PJO_JOURNEETYPE'));
          SetControlEnabled('PDJ_JOURNEETYPE',False);
          SetControlEnabled('BTNDROITS', True); //PT1
     End
     Else //PT1
     // Si TF est nul, on est en cas de paramétrage des droits.
     Begin
          SetField('PDJ_JOURNEETYPE', '***');
          SetField('PDJ_QUANTITE1',   0);
          SetField('PDJ_QUANTITE1',   0);
     End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 21/06/2007
Modifié le ... :   /  /
Description .. : Mise à jour droit journée type
Mots clefs ... :
*****************************************************************}
procedure TOM_DROITJOURNEETYPE.OnUpdateRecord ;
var
  droitst : string;
  droitjour: Thvalcombobox;
  DD,DF : TDateTime;
begin
  inherited;
     If TF <> Nil Then  //PT1
     Begin
          droitjour := THValComboBox(GetControl('PDJ_PGDROIT'));
          droitst := droitjour.value;
          if droitst = '' then
          begin
               PGIBOX('Le droit doit être renseigné.', Ecran.Caption);
               Setfocuscontrol('PDJ_PGDROIT');
               lasterror := 1;
               exit;
          end
          Else //PT1
          Begin
               SetField('PDJ_LIBELLE', DroitJour.Text);
          End;
          SetControlEnabled('BTNDROITS', False); //PT1
          if (Ecran.Name = 'DROITJOURTYPE_FSL') then
            if (Lasterror = 0) then Journee := TF.TOBFiltre.GetValue('PJO_JOURNEETYPE'); //PT2
          if (Ecran.Name = 'DROITJOURTYP1_FSL') then
            if (Lasterror = 0) then
            Begin
               Journee := TF.TOBFiltre.detail[0].GetValue('PDJ_JOURNEETYPE'); //PT2
               PresenceDonneMoisCalculActuel (DD,DF);   //PT3
               DateModif := DD;
            End;
     End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 20/07/2007 / PT1
Modifié le ... :   /  /
Description .. : Suppression d'un enregistrement
Mots clefs ... : 
*****************************************************************}
procedure TOM_DROITJOURNEETYPE.OnDeleteRecord;
var
  CodeDROIT : String;
  DD,DF : TDateTime;
begin
  inherited;
     // Contrôle à effectuer uniquement en cas de paramétrage des droits
     If TF = Nil Then
     Begin
          CodeDROIT := GetField('PDJ_PGDROIT');
          if (CodeDROIT = 'HNU') or (CodeDROIT = 'PAJ') or (CodeDROIT = 'PAN') then
          begin
            LastError := 1;
            PGIError (TraduireMemoire('Ces droits sont prédéfinis CEGID. Vous ne pouvez pas les supprimer.'));
          end else If ExisteSQL('SELECT 1 FROM DROITJOURNEETYPE WHERE PDJ_PGDROIT="'+CodeDROIT+'" AND PDJ_JOURNEETYPE<>"***"') Then
          Begin
               PGIError (TraduireMemoire('Suppression impossible. Le droit est utilisé au sein d''une journée type'));
               LastError := 1;
          End;
     End else begin
          if (Ecran.Name = 'DROITJOURTYPE_FSL') then
            if (Lasterror = 0) then Journee := TF.TOBFiltre.GetValue('PJO_JOURNEETYPE'); //PT2
          if (Ecran.Name = 'DROITJOURTYP1_FSL') then
            if (Lasterror = 0) then
            Begin
               Journee := TF.TOBFiltre.detail[0].GetValue('PDJ_JOURNEETYPE'); //PT2
               PresenceDonneMoisCalculActuel (DD,DF);   //PT3
               DateModif := DD;
            End;
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 10/08/2007  / PT2
Modifié le ... :   /  /
Description .. : Suite à la création ou la modification d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_DROITJOURNEETYPE.OnAfterUpdateRecord ;
var DD,DF : TDatetime;
begin
  inherited;
     If (LastError = 0) And (PresenceJourneyIsUsed(Journee)) Then
     begin
        PresenceDonneMoisCalculActuel (DD,DF);
        CompteursARecalculer(DD);
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 10/08/2007  / PT2
Modifié le ... :   /  /
Description .. : Suite à la suppression d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_DROITJOURNEETYPE.OnAfterDeleteRecord;
var DD,DF : TDatetime;
begin
  inherited;
     If (LastError = 0) And (PresenceJourneyIsUsed(Journee)) Then
     begin
        PresenceDonneMoisCalculActuel (DD,DF);
        CompteursARecalculer(DD);
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 20/07/2007
Modifié le ... :   /  /    
Description .. : Annulation de création / modification d'un enregistrement
Mots clefs ... : 
*****************************************************************}
procedure TOM_DROITJOURNEETYPE.OnCancelRecord;
begin
     SetControlEnabled('BTNDROITS', False); //PT1
end;

initialization
  registerclasses([TOM_JOURNEETYPE,TOM_DROITJOURNEETYPE]);

end.
