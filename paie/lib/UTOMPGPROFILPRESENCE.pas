{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 08/03/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PROFILPRESENCE (PROFILPRESENCE)
Mots clefs ... : TOM;PROFILPRESENCE
*****************************************************************
PT1  16/07/2007  FLO  Ajout de contrôles de saisie, d'une duplication par date
PT2  09/08/2007  GGU/FLO  Recalcul automatique des compteurs lors d'une modification ou d'une suppression
}
Unit UTOMPGPROFILPRESENCE ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
{$else}
     MaineAgl,
     UTob,
{$ENDIF}
     HCtrls,
     UTOM,
     HDB,
     Menus;

Type
  TOM_PROFILPRESENCE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnAfterUpdateRecord        ; override; //PT2
    procedure OnAfterDeleteRecord        ; override; //PT2
  private
{$IFNDEF EAGLCLIENT}
    Control_TYPEAFFECT : THDBValComboBox;
    Control_TYPSEUILT1, Control_TYPSEUILT2, Control_TYPSEUILT3, Control_TYPCONTINGHSUP, Control_TYPEPLAFHREAN : THDBValComboBox;
{$ELSE}
    Control_TYPEAFFECT : THValComboBox;
    Control_TYPSEUILT1, Control_TYPSEUILT2, Control_TYPSEUILT3, Control_TYPCONTINGHSUP, Control_TYPEPLAFHREAN : THDBValComboBox;
{$ENDIF}
    MenuDupli : TPopupMenu;
    HadDuplic : Boolean;
//    AccepteQuit : Boolean;
    Profil , libelleprofil     : String;    //PT2
    DateModif   : TDateTime; //PT2
    VerifImpact : Boolean;   //PT2
    Procedure OnTYPEAFFECTExit(Sender : TObject);
    procedure OnSOCDEFAUTExit(Sender : TObject);
    Procedure OnDuplique(Sender : TObject);
    Procedure OnDupliqueAvec(Sender : TObject);
    Procedure OnDupliqueDate(Sender: TObject); //PT1
    procedure Duplique(ChangeDate : boolean; DupliquerLesCompteurs: Boolean);
    Procedure OnShowPlanningClick(Sender : TObject);
    Procedure OnChangeControlDouble(Sender : TObject);
    function VerifSaisie(ControlName: String; LibName : String = ''; PrintMessage : Boolean = False): String;
  end ;

Implementation

Uses HEnt1,
     HMsgBox,
     sysutils,
     HTB97,
     PGPlanningPresenceCycle, PgPlanningOutils
     , PGPresence //PT1
     ;

function lpad(Str: String; Size : Integer; Pad : String = '0') : String;
begin
  result := Str;
  While length(result) < Size do
  begin
    result := Pad + result;
  end;
end;

procedure TOM_PROFILPRESENCE.OnNewRecord ;
begin
  Inherited ;

     // PT1 - Début
     // Affectation de valeurs par défaut
     // Type de profil : Annualisation
     SetField('PPQ_TYPEPROFILPRES', '001');
     // Temps de travail : Temps plein
     SetField('PPQ_TPSTRAVAIL',     'C');
     // Nombre de jours du cylce : 7
     SetField('PPQ_NBJOURSCYCLE',   '7');
     // Vidage des seuils
     SetControlText('PPQ_SEUILT1_', '');
     SetControlText('PPQ_SEUILT2_', '');
     SetControlText('PPQ_SEUILT3_', '');
     SetControlText('PPQ_CONTINGHSUP_', '');
     SetControlText('PPQ_PLAFHREANNU_', '');
     // PT1 - Fin
end ;

procedure TOM_PROFILPRESENCE.OnDeleteRecord ;
begin
  Inherited ;

  //PT1 - Début
  // On ne peut supprimer un profil en cours d'utilisation
  If ExisteSQL ('SELECT 1 FROM PROFILPRESSALARIE WHERE PPZ_PROFILPRES="'+GetControlText('PPQ_PROFILPRES')+'" AND PPZ_DATEVALIDITE>="'+UsDateTime(StrToDate(GetControlText('PPQ_DATEVALIDITE')))+'"') Then
  Begin
     PGIError (TraduireMemoire('Suppression impossible. Le profil est utilisé.'));
     LastError := 1;
  End;

  // Suppression également des compteurs associés
  If LastError = 0 then ExecuteSQL ('DELETE FROM PROFILCOMPTEURPRES WHERE PPV_PROFILPRES="'+GetControlText('PPQ_PROFILPRES')+'" AND PPV_DATEVALIDITE="'+UsDateTime(StrToDate(GetControlText('PPQ_DATEVALIDITE')))+'"');
  //PT1 - Fin

  //PT2 - Début
  If LastError = 0 then
  Begin
     Profil      := GetField('PPQ_PROFILPRES');
     DateModif   := GetField('PPQ_DATEVALIDITE');
     libelleprofil := Getfield('PPQ_LIBELLE');
  End;
  //PT2 - Fin
end ;

procedure TOM_PROFILPRESENCE.OnUpdateRecord ;
var
MsgOng1,MsgOng2,MsgOng3 : String;
begin
  Inherited ;
  //PT1 - Début
  MsgOng1 := '';
  MsgOng2 := '';
  MsgOng3 := '';

  // Onglet Caractéristiques : Clé de l'enregistrement
  MsgOng1 := MsgOng1 + VerifSaisie('PPQ_PROFILPRES');
  MsgOng1 := MsgOng1 + VerifSaisie('PPQ_LIBELLE');
  MsgOng1 := MsgOng1 + VerifSaisie('PPQ_DATEVALIDITE');
  If MsgOng1 <> '' Then MsgOng1 := '#13#10#13#10Onglet "Caractéristiques" :' + MsgOng1;

  // Onglet Heures supplémentaires
  // Si modulation, les limites basses et hautes sont obligatoires
  If GetControlText('PPQ_TYPEPROFILPRES') = '002' Then
  Begin
       MsgOng2 := MsgOng2 + VerifSaisie('PPQ_LIMITEBASMODUL');
       MsgOng2 := MsgOng2 + VerifSaisie('PPQ_LIMITEHTEMODUL');
  End;
  // Si le type des champs suivants a été défini, on vérifie le renseignement des valeurs associées
  MsgOng2 := MsgOng2 + VerifSaisie('PPQ_NBJOURSCYCLE');
  If GetControlText('PPQ_TYPSEUILT1')     <>'' Then MsgOng2 := MsgOng2 + VerifSaisie('PPQ_SEUILT1','Le seuil 1');
  If (GetControlText('PPQ_SEUILT1_')<>'') And (GetControlText('PPQ_SEUILT1_')<>'0') Then MsgOng2 := MsgOng2 + VerifSaisie('PPQ_TYPSEUILT1');
  If GetControlText('PPQ_TYPSEUILT2')     <>'' Then MsgOng2 := MsgOng2 + VerifSaisie('PPQ_SEUILT2','Le seuil 2');
  If (GetControlText('PPQ_SEUILT2_')<>'') And (GetControlText('PPQ_SEUILT2_')<>'0') Then MsgOng2 := MsgOng2 + VerifSaisie('PPQ_TYPSEUILT2');
  If GetControlText('PPQ_TYPSEUILT3')     <>'' Then MsgOng2 := MsgOng2 + VerifSaisie('PPQ_SEUILT3','Le seuil 3');
  If (GetControlText('PPQ_SEUILT3_')<>'') And (GetControlText('PPQ_SEUILT3_')<>'0') Then MsgOng2 := MsgOng2 + VerifSaisie('PPQ_TYPSEUILT3');
  If GetControlText('PPQ_TYPCONTINGHSUP') <>'' Then MsgOng2 := MsgOng2 + VerifSaisie('PPQ_CONTINGHSUP', 'Le contingent d''heures supplémentaires');
  If (GetControlText('PPQ_CONTINGHSUP_')<>'') And (GetControlText('PPQ_CONTINGHSUP_')<>'0') Then MsgOng2 := MsgOng2 + VerifSaisie('PPQ_TYPCONTINGHSUP');
  If GetControlText('PPQ_TYPEPLAFHREAN')  <>'' Then MsgOng2 := MsgOng2 + VerifSaisie('PPQ_PLAFHREANNU', 'Le plafond');
  If (GetControlText('PPQ_PLAFHREANNU_')<>'') And (GetControlText('PPQ_PLAFHREANNU_')<>'0') Then MsgOng2 := MsgOng2 + VerifSaisie('PPQ_TYPEPLAFHREAN');
  If MsgOng2 <> '' Then MsgOng2 := '#13#10#13#10Onglet "Heures supplémentaires" :' + MsgOng2;

  // Onglet Affectation
  MsgOng3 := MsgOng3 + VerifSaisie('PPQ_DATEDEBUTAFF');
  MsgOng3 := MsgOng3 + VerifSaisie('PPQ_TYPEAFFECT');
  MsgOng3 := MsgOng3 + VerifSaisie('PPQ_CYCLEAFFECT');
  If MsgOng3 <> '' Then MsgOng3 := '#13#10#13#10Onglet "Affectation" :' + MsgOng3;
  //PT1 - Fin

  If LastError = 0 Then
  begin

     // Si rien de renseigné dans onglet 2, on affiche un message d'info uniquement
     If (GetControlText('PPQ_SEUILT1_') = '') And (GetControlText('PPQ_SEUILT2_') = '') And (GetControlText('PPQ_SEUILT3_') = '') //PT1
         And (GetControlText('PPQ_CONTINGHSUP_') = '') And (GetControlText('PPQ_PLAFHREANNU_') = '') Then
          PGIInfo (TraduireMemoire('Aucun paramétrage des heures supplémentaires n''a été effectué.'));

     //PT2 - Début
     VerifImpact := (ds.State = dsEdit);
     Profil      := GetField('PPQ_PROFILPRES');
     DateModif   := GetField('PPQ_DATEVALIDITE');
     //PT2 - Fin
      if HadDuplic then
      begin
        verifimpact := true;
        HadDuplic := False;
      end;
  End
  Else
     PGIError(Format(TraduireMemoire('Veuillez renseigner le ou les champ(s) suivant(s) : %s%s%s'),[MsgOng1,MsgOng2,MsgOng3])); //PT1
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 09/08/2007 / PT2
Modifié le ... :   /  /
Description .. : Suite à la suppression d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_PROFILPRESENCE.OnAfterDeleteRecord ;
var
libellemodif,libellemodif2, daterecalcul, paramrecalcul : string;
begin
  Inherited ;
     If (LastError = 0) And (PresenceProfileIsUsed(Profil)) Then
     begin
         Daterecalcul := datetostr(datemodif);
         libellemodif := 'Le profil de présence ' + Profil + ' ' + libelleprofil;
         libellemodif2 := 'a été supprimé';
         paramrecalcul := libellemodif + ';' + libellemodif2 + ';' + daterecalcul;
         AglLanceFiche('PAY','PRESRECALCUL','','',paramrecalcul);
         // CompteursARecalculer(DateModif);
     end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 09/08/2007 / PT2
Modifié le ... :   /  /
Description .. : Suite à la création ou modification d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_PROFILPRESENCE.OnAfterUpdateRecord ;
var
libellemodif,libellemodif2, daterecalcul, paramrecalcul : string;
begin
  Inherited ;
     If (LastError = 0) And VerifImpact And (PresenceProfileIsUsed(Profil)) Then
     begin
         Daterecalcul := datetostr(datemodif);
         libellemodif := 'Le profil de présence ' + Profil + ' ' +
         getfield('PPQ_LIBELLE');
         libellemodif2 := 'a été modifié';
         paramrecalcul := libellemodif + ';' + libellemodif2 + ';' + daterecalcul;
         AglLanceFiche('PAY','PRESRECALCUL','','',paramrecalcul);
         // CompteursARecalculer(DateModif);
     end;
end ;

procedure TOM_PROFILPRESENCE.OnLoadRecord ;
begin
  Inherited ;

  if Control_TYPSEUILT1.Value = 'PER' then
    SetControlText('PPQ_SEUILT1_',GetField('PPQ_SEUILT1'))
  else
  Begin
    If GetField('PPQ_SEUILT1') <> '0' Then
    Begin
         SetControlText('PPQ_SEUILT1_',lpad(GetField('PPQ_SEUILT1'),4));
         SetControlText('LIBELMTNAT1', RechDom('PGELEMENTNAT',lpad(GetField('PPQ_SEUILT1'),4),False)); //PT1
    End
    Else
    Begin
         SetControlText('LIBELMTNAT1', '');
         SetControlText('PPQ_SEUILT1_','');
    End;
  End;
  if Control_TYPSEUILT2.Value = 'PER' then
    SetControlText('PPQ_SEUILT2_',GetField('PPQ_SEUILT2'))
  else
  begin
    If GetField('PPQ_SEUILT2') <> '0' Then
    Begin
         SetControlText('PPQ_SEUILT2_',lpad(GetField('PPQ_SEUILT2'),4));
         SetControlText('LIBELMTNAT2', RechDom('PGELEMENTNAT',lpad(GetField('PPQ_SEUILT2'),4),False)); //PT1
    End
    Else
    Begin
         SetControlText('LIBELMTNAT2', '');
         SetControlText('PPQ_SEUILT2_','');
    End;
  End;
  if Control_TYPSEUILT3.Value = 'PER' then
    SetControlText('PPQ_SEUILT3_',GetField('PPQ_SEUILT3'))
  else
  begin
    If GetField('PPQ_SEUILT3') <> '0' Then
    Begin
         SetControlText('PPQ_SEUILT3_',lpad(GetField('PPQ_SEUILT3'),4));
         SetControlText('LIBELMTNAT3', RechDom('PGELEMENTNAT',lpad(GetField('PPQ_SEUILT3'),4),False)); //PT1
    End
    Else
    Begin
         SetControlText('LIBELMTNAT3', '');
         SetControlText('PPQ_SEUILT3_','');
    End;
  end;
  if Control_TYPCONTINGHSUP.Value = 'PER' then
    SetControlText('PPQ_CONTINGHSUP_',GetField('PPQ_CONTINGHSUP'))
  else
  begin
    If GetField('PPQ_CONTINGHSUP') <> '0' Then
    begin
         SetControlText('PPQ_CONTINGHSUP_',lpad(GetField('PPQ_CONTINGHSUP'),4));
         SetControlText('LIBELMTNAT4', RechDom('PGELEMENTNAT',lpad(GetField('PPQ_CONTINGHSUP'),4),False)); //PT1
    End
    Else
    Begin
         SetControlText('LIBELMTNAT4', '');
         SetControlText('PPQ_CONTINGHSUP_','');
    End;
  end;
  if Control_TYPEPLAFHREAN.Value = 'PER' then
    SetControlText('PPQ_PLAFHREANNU_',GetField('PPQ_PLAFHREANNU'))
  else
  begin
    If GetField('PPQ_PLAFHREANNU') <> '0' Then
    begin
         SetControlText('PPQ_PLAFHREANNU_',lpad(GetField('PPQ_PLAFHREANNU'),4));
         SetControlText('LIBELMTNAT5', RechDom('PGELEMENTNAT',lpad(GetField('PPQ_PLAFHREANNU'),4),False)); //PT1
    End
    Else
    Begin
         SetControlText('LIBELMTNAT5', '');
         SetControlText('PPQ_PLAFHREANNU_','');
    End;
  end;

  OnTYPEAFFECTExit(Self);
  OnSOCDEFAUTExit(Control_TYPSEUILT1);
  OnSOCDEFAUTExit(Control_TYPSEUILT2);
  OnSOCDEFAUTExit(Control_TYPSEUILT3);
  OnSOCDEFAUTExit(Control_TYPCONTINGHSUP);
  OnSOCDEFAUTExit(Control_TYPEPLAFHREAN);
end ;

procedure TOM_PROFILPRESENCE.OnChangeField ( F: TField ) ;
var Enabled : Boolean;
begin
  Inherited ;

  //PT1 - Début
  If F.FieldName = 'PPQ_TYPEPROFILPRES' Then
  Begin
       // Si modulation, on rend saisissable les limites
       Enabled := (GetControlText('PPQ_TYPEPROFILPRES') = '002');
       SetField('PPQ_LIMITEBASMODUL', '0');
       SetField('PPQ_LIMITEHTEMODUL', '0');
       SetControlEnabled('PPQ_LIMITEBASMODUL', Enabled);
       SetControlEnabled('PPQ_LIMITEHTEMODUL', Enabled);
  End;
  //PT1 - Fin
end ;

procedure TOM_PROFILPRESENCE.OnArgument ( S: String ) ;
begin
  Inherited ;
{$IFNDEF EAGLCLIENT}
  Control_TYPEAFFECT := (GetControl('PPQ_TYPEAFFECT')  as THDBValComboBox);
  Control_TYPSEUILT1     := (GetControl('PPQ_TYPSEUILT1')  as THDBValComboBox);
  Control_TYPSEUILT2     := (GetControl('PPQ_TYPSEUILT2')  as THDBValComboBox);
  Control_TYPSEUILT3     := (GetControl('PPQ_TYPSEUILT3')  as THDBValComboBox);
  Control_TYPCONTINGHSUP := (GetControl('PPQ_TYPCONTINGHSUP')  as THDBValComboBox);
  Control_TYPEPLAFHREAN  := (GetControl('PPQ_TYPEPLAFHREAN')  as THDBValComboBox);
{$ELSE}
  Control_TYPEAFFECT := (GetControl('PPQ_TYPEAFFECT')  as THValComboBox);
  Control_TYPSEUILT1     := (GetControl('PPQ_TYPSEUILT1')  as THValComboBox);
  Control_TYPSEUILT2     := (GetControl('PPQ_TYPSEUILT2')  as THValComboBox);
  Control_TYPSEUILT3     := (GetControl('PPQ_TYPSEUILT3')  as THValComboBox);
  Control_TYPCONTINGHSUP := (GetControl('PPQ_TYPCONTINGHSUP')  as THValComboBox);
  Control_TYPEPLAFHREAN  := (GetControl('PPQ_TYPEPLAFHREAN')  as THValComboBox);
{$ENDIF}
  Control_TYPEAFFECT.OnExit := OnTYPEAFFECTExit;
  Control_TYPSEUILT1.OnExit := OnSOCDEFAUTExit;
  Control_TYPSEUILT2.OnExit := OnSOCDEFAUTExit;
  Control_TYPSEUILT3.OnExit := OnSOCDEFAUTExit;
  Control_TYPCONTINGHSUP.OnExit := OnSOCDEFAUTExit;
  Control_TYPEPLAFHREAN.OnExit := OnSOCDEFAUTExit;
  (GetControl('PPQ_SEUILT1_') as THEdit).OnChange := OnChangeControlDouble;
  (GetControl('PPQ_SEUILT2_') as THEdit).OnChange := OnChangeControlDouble;
  (GetControl('PPQ_SEUILT3_') as THEdit).OnChange := OnChangeControlDouble;
  (GetControl('PPQ_CONTINGHSUP_') as THEdit).OnChange := OnChangeControlDouble;
  (GetControl('PPQ_PLAFHREANNU_') as THEdit).OnChange := OnChangeControlDouble;
//  (GetControl('BDUPLIQUER')  as TToolbarButton97).OnClick := OnDuplique;
  MenuDupli := TPopupMenu(GetControl('POPUPDUPLIC'));
  if MenuDupli <> nil then
  begin
      MenuDupli.Items[0].OnClick := OnDupliqueDate; //PT1
      MenuDupli.Items[1].OnClick := OnDupliqueAvec;
      MenuDupli.Items[2].OnClick := OnDuplique;
  end;
  (GetControl('BPLANNING') as TToolbarButton97).OnClick := OnShowPlanningClick;

//  AccepteQuit := True;
end ;

procedure TOM_PROFILPRESENCE.OnTYPEAFFECTExit(Sender: TObject);
begin
  if Control_TYPEAFFECT.Value = 'CYC' then
  begin
    SetControlProperty('PPQ_CYCLEAFFECT','DataType','PGCYCLE');
    SetControlText    ('TPPQ_CYCLEAFFECT','Rythme de travail'); //PT1
    if ds.State <> dsBrowse then SetControlText    ('PPQ_CYCLEAFFECT', '');
  end
  else if Control_TYPEAFFECT.Value = 'MOD' then
  begin
    SetControlProperty('PPQ_CYCLEAFFECT','DataType','PGMODELECYCLE');
    SetControlText    ('TPPQ_CYCLEAFFECT','Cycle'); //PT1
    if ds.State <> dsBrowse then SetControlText    ('PPQ_CYCLEAFFECT', '');
  end;
end;

procedure TOM_PROFILPRESENCE.OnSOCDEFAUTExit(Sender: TObject);
Var
  stValue, stTabletteName, stControlValeurName : String;
  BoPerso : Boolean;
begin
  if      (Sender is THDBValComboBox) then stValue := (Sender as THDBValComboBox).Value
  else if (Sender is THValComboBox)   then stValue := (Sender as THValComboBox).Value;
  BoPerso :=  (stValue = 'PER') Or (stValue = ''); //PT1
//  Si perso : pas de tablette pour la zone valeur, sinon : tablette des elt nat
  if BoPerso then
    stTabletteName := ''
  else
    stTabletteName := 'PGELEMENTNAT';

  if Sender = Control_TYPSEUILT1          then
    stControlValeurName := 'PPQ_SEUILT1_'
  else if Sender = Control_TYPSEUILT2     then
    stControlValeurName := 'PPQ_SEUILT2_'
  else if Sender = Control_TYPSEUILT3     then
    stControlValeurName := 'PPQ_SEUILT3_'
  else if Sender = Control_TYPCONTINGHSUP then
    stControlValeurName := 'PPQ_CONTINGHSUP_'
  else if Sender = Control_TYPEPLAFHREAN  then
    stControlValeurName := 'PPQ_PLAFHREANNU_';

//{$IFNDEF EAGLCLIENT}
//  (GetControl(stControlValeurName) as THDBEdit).ElipsisButton := not BoPerso;
//  (GetControl(stControlValeurName) as THDBEdit).ElipsisAutoHide := False;
//  (GetControl(stControlValeurName) as THDBEdit).DataType := stTabletteName;
//{$ELSE}
  (GetControl(stControlValeurName) as THEdit).ElipsisButton := not BoPerso;
  (GetControl(stControlValeurName) as THEdit).ElipsisAutoHide := False;
  (GetControl(stControlValeurName) as THEdit).DataType := stTabletteName;
//{$ENDIF}
end;

procedure TOM_PROFILPRESENCE.Duplique(ChangeDate : boolean ; DupliquerLesCompteurs: Boolean);
var
  SavPPQ_PROFILPRES, SavPPQ_DATEVALIDITE, SavPPQ_LIBELLE, SavPPQ_TYPEPROFILPRES,
  SavPPQ_TPSTRAVAIL, SavPPQ_LIMITEHTEMODUL, SavPPQ_LIMITEBASMODUL, SavPPQ_TYPSEUILT1,
  SavPPQ_SEUILT1, SavPPQ_TYPSEUILT2, SavPPQ_SEUILT2, SavPPQ_TYPSEUILT3, SavPPQ_SEUILT3,
  SavPPQ_NBJOURSCYCLE, SavPPQ_TYPCONTINGHSUP, SavPPQ_CONTINGHSUP, SavPPQ_TYPEPLAFHREAN,
  SavPPQ_PLAFHREANNU, SavPPQ_DATEREGANNU, SavPPQ_DATEDEBUTAFF, SavPPQ_TYPEAFFECT,
  SavPPQ_CYCLEAFFECT, SavPPQ_JOURRTTLIBRE, SavPPQ_JOURRTTEMPL : String;
  champsduplic, mode, Appel : String;
begin
  HadDuplic := True;
  if ChangeDate Then mode := 'DAT' //PT1
  else if DupliquerLesCompteurs then mode := 'DET' else mode := 'ENT';
  Appel := 'PPRES';
  {On sauve les champs}
  SavPPQ_PROFILPRES     := GetControlText('PPQ_PROFILPRES');
  SavPPQ_DATEVALIDITE   := GetControlText('PPQ_DATEVALIDITE');
  SavPPQ_LIBELLE        := GetControlText('PPQ_LIBELLE');
  SavPPQ_TYPEPROFILPRES := GetControlText('PPQ_TYPEPROFILPRES');
  SavPPQ_TPSTRAVAIL     := GetControlText('PPQ_TPSTRAVAIL');
  SavPPQ_LIMITEHTEMODUL := GetControlText('PPQ_LIMITEHTEMODUL');
  SavPPQ_LIMITEBASMODUL := GetControlText('PPQ_LIMITEBASMODUL');
  SavPPQ_TYPSEUILT1     := GetControlText('PPQ_TYPSEUILT1');
  SavPPQ_SEUILT1        := GetControlText('PPQ_SEUILT1');
  SavPPQ_TYPSEUILT2     := GetControlText('PPQ_TYPSEUILT2');
  SavPPQ_SEUILT2        := GetControlText('PPQ_SEUILT2');
  SavPPQ_TYPSEUILT3     := GetControlText('PPQ_TYPSEUILT3');
  SavPPQ_SEUILT3        := GetControlText('PPQ_SEUILT3');
  SavPPQ_NBJOURSCYCLE   := GetControlText('PPQ_NBJOURSCYCLE');
  SavPPQ_TYPCONTINGHSUP := GetControlText('PPQ_TYPCONTINGHSUP');
  SavPPQ_CONTINGHSUP    := GetControlText('PPQ_CONTINGHSUP');
  SavPPQ_TYPEPLAFHREAN  := GetControlText('PPQ_TYPEPLAFHREAN');
  SavPPQ_PLAFHREANNU    := GetControlText('PPQ_PLAFHREANNU');
  SavPPQ_DATEREGANNU    := GetControlText('PPQ_DATEREGANNU');
  SavPPQ_DATEDEBUTAFF   := GetControlText('PPQ_DATEDEBUTAFF');
  SavPPQ_TYPEAFFECT     := GetControlText('PPQ_TYPEAFFECT');
  SavPPQ_CYCLEAFFECT    := GetControlText('PPQ_CYCLEAFFECT');
//  SavPPQ_JOURRTTLIBRE   := GetControlText('PPQ_JOURRTTLIBRE');//PT1
//  SavPPQ_JOURRTTEMPL    := GetControlText('PPQ_JOURRTTEMPL'); //PT1
  {On demande le nouveau code et le nouveau libellé}

  champsduplic := AglLanceFiche('PAY', 'PARAMPRESENCEDUP', '', '',Appel + ';' + SavPPQ_PROFILPRES + ';' + mode + ';' + SavPPQ_LIBELLE + ';' + SavPPQ_DATEVALIDITE);
  if champsduplic <> '' then
  begin
    {On récupère les valeurs saisies}
    if ChangeDate Then                            //PT1
         SavPPQ_DATEVALIDITE := champsduplic
    Else
    begin
         SavPPQ_PROFILPRES := readtokenst(champsduplic);
         SavPPQ_LIBELLE    := champsduplic;
    end;
    {On cré un nouvel enregistrement}
    (GetControl('BINSERT')  as TToolbarButton97).OnClick(self);
    {On charge les champs et le nouveau code}
    SetField('PPQ_PROFILPRES',    SavPPQ_PROFILPRES );
    SetField('PPQ_DATEVALIDITE',  SavPPQ_DATEVALIDITE );
    SetField('PPQ_LIBELLE',       SavPPQ_LIBELLE );
    SetField('PPQ_TYPEPROFILPRES',SavPPQ_TYPEPROFILPRES );
    SetField('PPQ_TPSTRAVAIL',    SavPPQ_TPSTRAVAIL );
    SetField('PPQ_LIMITEHTEMODUL',SavPPQ_LIMITEHTEMODUL );
    SetField('PPQ_LIMITEBASMODUL',SavPPQ_LIMITEBASMODUL );
    SetField('PPQ_TYPSEUILT1',    SavPPQ_TYPSEUILT1 );
    If SavPPQ_SEUILT1 <> '0' Then
          SetControlText('PPQ_SEUILT1_', SavPPQ_SEUILT1 ) //PT1
    Else
          SetControlText('PPQ_SEUILT1_', '' );
    SetField('PPQ_TYPSEUILT2',    SavPPQ_TYPSEUILT2 );
    If SavPPQ_SEUILT2 <> '0' Then
          SetControlText('PPQ_SEUILT2_', SavPPQ_SEUILT2 ) //PT1
    Else
          SetControlText('PPQ_SEUILT2_', '' ); //PT1
    SetField('PPQ_TYPSEUILT3',    SavPPQ_TYPSEUILT3 );
    If SavPPQ_SEUILT3 <> '0' Then
          SetControlText('PPQ_SEUILT3_', SavPPQ_SEUILT3 ) //PT1
    Else
          SetControlText('PPQ_SEUILT3_', '' ); //PT1
    SetField('PPQ_NBJOURSCYCLE',  SavPPQ_NBJOURSCYCLE );
    SetField('PPQ_TYPCONTINGHSUP',SavPPQ_TYPCONTINGHSUP );
    If SavPPQ_CONTINGHSUP <> '0' Then
          SetControlText('PPQ_CONTINGHSUP_', SavPPQ_CONTINGHSUP ) //PT1
    Else
          SetControlText('PPQ_CONTINGHSUP_', '' ); //PT1
    SetField('PPQ_TYPEPLAFHREAN', SavPPQ_TYPEPLAFHREAN );
    If SavPPQ_PLAFHREANNU <> '0' Then
          SetControlText('PPQ_PLAFHREANNU_', SavPPQ_PLAFHREANNU ) //PT1
    Else
          SetControlText('PPQ_PLAFHREANNU_', '' ); //PT1
    SetField('PPQ_DATEREGANNU',   SavPPQ_DATEREGANNU );
    SetField('PPQ_DATEDEBUTAFF',  SavPPQ_DATEDEBUTAFF );
    SetField('PPQ_TYPEAFFECT',    SavPPQ_TYPEAFFECT );
    SetField('PPQ_CYCLEAFFECT',   SavPPQ_CYCLEAFFECT );
    SetField('PPQ_JOURRTTLIBRE',  SavPPQ_JOURRTTLIBRE );
    SetField('PPQ_JOURRTTEMPL',   SavPPQ_JOURRTTEMPL );
    {On force la sauvegarde}
    (GetControl('BVALIDER') as TToolbarButton97).OnClick(self);
  end;


end;

procedure TOM_PROFILPRESENCE.OnDupliqueDate(Sender: TObject); //PT1
begin
  Duplique(True, False);
end;

procedure TOM_PROFILPRESENCE.OnDuplique(Sender: TObject);
begin
  Duplique(False, False);
end;


procedure TOM_PROFILPRESENCE.OnDupliqueAvec(Sender: TObject);
begin
  Duplique(False, True);
end;

procedure TOM_PROFILPRESENCE.OnChangeControlDouble(Sender: TObject);
var
  tempStr : String;
  Function Num(st : String; FloatAccepted : Boolean) : String;
  begin
    result := '';
    if not IsNumeric(st) then exit;
    if (not FloatAccepted) and (pos(',',st) > 0) then exit;
    if (not FloatAccepted) and (pos('.',st) > 0) then exit;
    result := st;
  end;
  Procedure UpdateControl(Sender : TObject; ControlBase, ControlVisible, ControlType : String; FloatAccepted : Boolean);
  begin
    if Sender = GetControl(ControlVisible) then
    begin
      tempStr := GetControlText(ControlVisible);
      if not (GetControlText(ControlType) = 'PER') then FloatAccepted := False;
      if ( tempStr = '') or (num(tempStr,FloatAccepted) <> '' ) then
      begin
        try
          SetField(ControlBase,tempStr);
        except // Gestion des problèmes de . et , dans les float
          if FloatAccepted then
          if pos(',',tempStr) > 0 then
            tempStr := StringReplace(tempStr, ',', '.', [rfIgnoreCase])
          else
            tempStr := StringReplace(tempStr, '.', ',', [rfIgnoreCase]);
          SetField(ControlBase, tempStr);
        end
      end else
        SetControlText(ControlVisible,'');
    end;
  end;
begin
  UpdateControl(Sender, 'PPQ_SEUILT1', 'PPQ_SEUILT1_', 'PPQ_TYPSEUILT1', True);
  UpdateControl(Sender, 'PPQ_SEUILT2', 'PPQ_SEUILT2_', 'PPQ_TYPSEUILT2', True);
  UpdateControl(Sender, 'PPQ_SEUILT3', 'PPQ_SEUILT3_', 'PPQ_TYPSEUILT3', True);
  UpdateControl(Sender, 'PPQ_CONTINGHSUP', 'PPQ_CONTINGHSUP_', 'PPQ_TYPCONTINGHSUP', True);
  UpdateControl(Sender, 'PPQ_PLAFHREANNU', 'PPQ_PLAFHREANNU_', 'PPQ_TYPEPLAFHREAN', False);
end;

procedure TOM_PROFILPRESENCE.OnShowPlanningClick(Sender: TObject);
begin                                            
  PGPlanningPresenceCycleOpen(Date,(Date+30),'','"'+ GetControlText('PPQ_CYCLEAFFECT')+'"', [piaSemaine, piaMois, piaHeure, piaOutlook], GetControlText('PPQ_TYPEAFFECT'));  
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 17/07/2007
Modifié le ... :   /  /    
Description .. : Vérification de la saisie d'un champ
Suite ........ : Préciser le libellé si besoin (pas de champ libellé associé de la forme T...)
Suite ........ : PrintMessage à true si besoin d'afficher l'erreur de suite, sinon seul le message est retourné
Suite ........ : Passe le LastError directement à 1.
Mots clefs ... : CONTROLE;SAISIE;
*****************************************************************}
function TOM_PROFILPRESENCE.VerifSaisie(ControlName: String; LibName : String = ''; PrintMessage : Boolean = False): String;
Var
     Msg,Value,Mask : String;
     Ctrl  : TControl;
     Champ : TCustomEdit;
     Error : Boolean;
begin
     Error  := False;
     Result := '';

     If (ControlName <> '') Then
     Begin
          Ctrl := GetControl(ControlName);
          If Ctrl = Nil Then Exit;

          If (Ctrl Is THValComboBox) Or (Ctrl Is THMultiValComboBox)
          {$IFNDEF EAGLCLIENT}Or (Ctrl Is THDBValComboBox) Or (Ctrl Is THDBMultiValComboBox){$ENDIF} Then
          Begin
               Champ := TCustomEdit (Ctrl);
               If (Champ <> Nil) And (Trim(Champ.Text) = '') Then
               Begin
                    Error := True;
                    If LibName = '' Then
                         Msg := GetControlText('T'+Champ.Name)
                    Else
                         Msg := LibName;
               End;
          End;

          If (Ctrl Is THEdit) {$IFNDEF EAGLCLIENT} Or (Ctrl Is THDBEdit) {$ENDIF} Then
          Begin
               Champ := TCustomEdit (Ctrl);
               If Champ = Nil Then Exit;

               Value := Trim(Champ.Text);
               If Value = '' Then
               Begin
                    Error := True;
                    If LibName = '' Then
                         Msg := GetControlText('T'+Champ.Name)
                    Else
                         Msg := LibName;
               End
               Else
               Begin
                    // Cas des dates, des numériques et des champs possédant un masque particulier
                    If (IsNumeric(Value)) And (StrToInt(Value) = 0) Then
                         Error := True;
                    If (Value = '__/__/____') Or (Value = StDate1900) Then
                         Error := True;

                    If Not Error Then
                    Begin
                         If Ctrl Is THEdit Then Mask := (Ctrl As THEdit).EditMask;
                         {$IFNDEF EAGLCLIENT}
                         If Ctrl Is THDBEdit Then Mask := (Ctrl As THDBEdit).EditMask;
                         {$ENDIF}
                         If (Mask <> '') Then If Pos(Copy(Mask,Length(Mask),1), Value) > 0 Then Error := True;
                    End;
                    
                    If Error Then
                    Begin
                         If LibName = '' Then
                              Msg := GetControlText('T'+Champ.Name)
                         Else
                              Msg := LibName;
                    End;
               End;
          End;

          If Error Then
          Begin
               If Not PrintMessage Then
                    Result := '#13#10- ' + Msg
               Else
               Begin
                    Result := Msg;
                    PGIError(Format(TraduireMemoire('Vous devez renseigner : %s'),[Msg]));
               End;
               LastError := LastError + 1;
          End;
     End;
end;

Initialization
  registerclasses ( [ TOM_PROFILPRESENCE ] ) ;
end.
