{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 30/08/2001
Modifié le ... :   /  /
Description .. : Gestion des motifs d'absence
Mots clefs ... : PAIE;ABSENCE
*****************************************************************}
{
PT1   : 30/08/2001 SB V562 Affectation predefini et Motif absence en création
PT2   : 22/10/2001 SB V563 Affectation predefini STD au lieu de CEG  Fiche de
                           bug n°337
PT3   : 21/11/2001 PH V563 Valeur par défaut JOURHEURE
PT4   : 30/04/2001 PH V563 Valeur par défaut systématique
Refonte accès V_PGI_env ** V_PGI_env.nodossier remplacé par PgRendNoDossier() **
PT5   : 17/07/2003 SB V_42 FQ 10719 la rubrique associé ne doit être affecté à
                           un seul salarié
PT6   : 06/05/2004 MF V_50 Contrôle champ Type absence renseigné qd versement
                           IJSS =  vrai
PT7   : 12/10/2004 VG V_50 Version S3 - FQ N°11680
PT8   : 30/11/2004 MF V_60 Qd modif coche "versement d' IJSS" la table
                           ABSENCESALARIE est mise à jour.
PT9   : 23/02/2005 MF V_60 Contrôle MEMCHECK : FreeAndNil(tob_rub) avant
                           création
PT10  : 04/04/2005 SB V_60 FQ 11781 Possibilité de saisir rubrique identique
PT11  : 26/04/2005 SB V_60 FQ 12059 Contrôle saisie type alimentation en
                           validation
PT12  : 23/01/2006 SB V_65 FQ 10866 Ajout clause predefini motif d'absence
PT13  : 28/02/2006 SB V_65 Gestion des couleurs du planning
PT14  : 28/02/2006 SB V_65 Ajout zone pour filtrer les motifs d'absences à
                           contrôler existence en saisie
PT15  : 30/02/2006 SB V_65 Gestion nouveaux champs
PT16  : 05/04/2006 Mf V_65 IJSS :  alimentation nbre jours carence par défaut
PT17  : 13/04/2006 SB V_65 Traitement nouveaux champs
PT18  : 19/05/2006 SB V_65 FQ 13158 Ajout message d'alerte rubrique inexistante
PT19  : 23/05/2006 SB V_65 FQ 13193 Controle absence existente
PT20  : 09/10/2006 VG V_70 Ajout contrôle pour Chômage intempéries
PT21  : 08/12/2006 SB V_70 Affectation des couleurs par défaut en création
PT22  : 23/01/2007 GGU V_72 Gestion des motifs de présence dans la table MOTIFABSENCE
PT23  : 05/02/2007 GGS V_72 Journal événements
PT24  : 27/02/2007 MF  V_72 FQ 13936 : mise à jour des NBJCALEND, NBJCARENCE...
PT25  : 26/04/2007 GGS V_72 Revue gestion Trace
PT26  : 07/05/2007 PH  V_72 Concept Paie

}
unit UTOMMotifAbs;

interface
uses StdCtrls, Controls, Classes, sysutils, dialogs, graphics,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}HDB, Fiche,
{$ELSE}
  eFiche,
{$ENDIF}
  ComCtrls, HCtrls, HEnt1, HMsgBox, UTOM, UTOB, HTB97,
  PgOutils, PgOutils2,
  P5Def, PAIETOM; //PT23


type

  TOM_MOTIFABSENCE = class(PGTOM) //PT23  Tom devient PGTOM
    procedure OnArgument(stArgument: string); override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override; { PT12 }
    procedure OnChangeField(F: TField); override;
    procedure OnLoadRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnNewRecord; override;
    procedure OnClose; override;
    procedure initialiseabsence;
    procedure EnabledChampsAbsence;
    procedure EnabledBtnColor(LeMode: Boolean); // PT26

  private
    Loadok, LectureSeule, CEG, STD, DOS, Error, OnFerme: boolean; { PT12 }
    TChampRub: string;
    T_Rub: TOB;
    GestionIJSS: string; //PT8
    DerniereCreate: string; { PT12 }
    GblProfil, GblProfilJ: string; { PT17 }
    IsPresenceMode: Boolean; //PT22
    Trace: TStringList; //PT23
    procedure ClickPgColor(Sender: TObject); { PT13 }
    procedure ClickPgDefaire(Sender: TObject); { PT13 }
// d PT16
    procedure ChangeGestIJSS(Sender: TObject);
    procedure CalculCarenceIJSS(TypeAbs: string);
// f PT16
    procedure InitialiseColor(ControlName: string); { PT17 }
    procedure MajAbsenceSalarie(PmaGestionIJSS, PmaMotifAbs: string; CarenceIJSS: integer); // PT24

    function ControlZoneRubrique(StDif: string): string; { PT17 }
    function ControlCumulPaie(Rub: string): Boolean; { PT17 }

  end;

implementation

{ TOM_MOTIFABSENCE }

procedure TOM_MOTIFABSENCE.OnArgument(stArgument: string);
var Btn: TToolBarButton97;
// d PT16
  CIJSS: TCheckBox;
// f PT16
begin
  inherited; { PT12 }
  //PT7
{$IFDEF CCS3}
  SetControlVisible('PMA_GESTIONIJSS', False);
{$ENDIF}
//PT25  Trace := TStringList.Create ;         //PT23

  //PT22
  // A true si on est sur la fiche des motifs de présence et à false
  // si on est sur la fiche des motifs d'absences
  IsPresenceMode := (Ecran.Name = 'MOTIFPRESENCE');
  //Fin PT22


  //FIN PT7

    { DEB PT13  }
    Btn := TToolBarButton97(GetControl('BTNPGCOLORACTIF'));
    if Assigned(Btn) then Btn.OnClick := ClickPgColor;
    Btn := TToolBarButton97(GetControl('BTNPGCOLORPAY'));
    if Assigned(Btn) then Btn.OnClick := ClickPgColor;
    Btn := TToolBarButton97(GetControl('BTNPGCOLORATT'));
    if Assigned(Btn) then Btn.OnClick := ClickPgColor;
    Btn := TToolBarButton97(GetControl('BTNPGCOLORVAL'));
    if Assigned(Btn) then Btn.OnClick := ClickPgColor;
    Btn := TToolBarButton97(GetControl('BTNPGCOLORANN'));
    if Assigned(Btn) then Btn.OnClick := ClickPgColor;
    Btn := TToolBarButton97(GetControl('BTNPGCOLORREF'));
    if Assigned(Btn) then Btn.OnClick := ClickPgColor;


    Btn := TToolBarButton97(GetControl('BTNDEFAIREACTIF'));
    if Assigned(Btn) then Btn.OnClick := ClickPgDefaire;
    Btn := TToolBarButton97(GetControl('BTNDEFAIREPAY'));
    if Assigned(Btn) then Btn.OnClick := ClickPgDefaire;
    Btn := TToolBarButton97(GetControl('BTNDEFAIREATT'));
    if Assigned(Btn) then Btn.OnClick := ClickPgDefaire;
    Btn := TToolBarButton97(GetControl('BTNDEFAIREVAL'));
    if Assigned(Btn) then Btn.OnClick := ClickPgDefaire;
    Btn := TToolBarButton97(GetControl('BTNDEFAIREANN'));
    if Assigned(Btn) then Btn.OnClick := ClickPgDefaire;
    Btn := TToolBarButton97(GetControl('BTNDEFAIREREF'));
    if Assigned(Btn) then Btn.OnClick := ClickPgDefaire;
    { FIN PT13  }
    // d PT16
    CIJSS := TCheckBox(GetControl('PMA_GESTIONIJSS'));
    if CIJSS <> nil then
      CIJSS.Onclick := ChangeGestIJSS;
    // f PT16


  PaieConceptPlanPaie(Ecran); // PT26
end;


procedure TOM_MOTIFABSENCE.OnUpdateRecord;
var
  Predef: string;
  T: Tob;
begin
  inherited; { DEB PT12 }
  OnFerme := False;
  if (DS.State in [dsInsert]) then
    DerniereCreate := GetField('PMA_MOTIFABSENCE')
  else
    if (DerniereCreate = GetField('PMA_MOTIFABSENCE')) then
      OnFerme := True; // le bug arrive on se casse !!!
{ FIN PT12 }
  Loadok := false;

  if (DS.State = dsinsert) then
  begin { DEB PT12 }
    if ExisteSql('SELECT PMA_MOTIFABSENCE' +
      ' FROM MOTIFABSENCE WHERE' +
      ' ##PMA_PREDEFINI## PMA_MOTIFABSENCE = "' + GetField('PMA_MOTIFABSENCE') + '"') = True then
    begin
      LastError := 1;
      PgiBox('Ce motif d''absence existe déjà . Vous devez le modifier.',
        Ecran.caption);
      SetFocusControl('PMA_MOTIFABSENCE');
      exit;
    end; { FIN PT12 }

    if (GetField('PMA_PREDEFINI') <> 'DOS') then
      SetField('PMA_NODOSSIER', '000000')
    else
      SetField('PMA_NODOSSIER', PgRendNoDossier());
  end;

  Predef := GetField('PMA_PREDEFINI');

  if (Predef = '') and (LastError = 0) then
  begin
    LastError := 1;
    PgiBox('Vous devez renseigner le champ prédéfini.', 'Modifs des absences');
    SetFocusControl('PMA_PREDEFINI');
  end;

  if (not IsPresenceMode) //PT22
    and (LastError = 0) and (Assigned(T_Rub))
    and (GetField('PMA_RUBRIQUE') <> '') then
  begin
    T := T_Rub.FindFirst(['PMA_RUBRIQUE'], [GetField('PMA_RUBRIQUE')], False);
    if Assigned(T) then
      if T.GetValue('PMA_MOTIFABSENCE') <> GetField('PMA_MOTIFABSENCE') then
      begin
        if PGIAsk('Attention! La rubrique en heure est déjà affectée à un' +
          ' autre motif.#13#10' + { PT17 }
          'Confirmez-vous l''enregistrement du motif?',
          Ecran.Caption) = MrNo then
        begin
          LastError := 1;
          exit;
        end;
      end;
  end;

{ DEB PT17 }
  if (not IsPresenceMode) //PT22
    and (LastError = 0) and (Assigned(T_Rub))
    and (GetField('PMA_RUBRIQUEJ') <> '') then
  begin
    T := T_Rub.FindFirst(['PMA_RUBRIQUEJ'], [GetField('PMA_RUBRIQUEJ')], False);
    if Assigned(T) then
      if T.GetValue('PMA_MOTIFABSENCE') <> GetField('PMA_MOTIFABSENCE') then
      begin
        if PGIAsk('Attention! La rubrique en jour est déjà affectée à un' +
          ' autre motif.#13#10' +
          'Confirmez-vous l''enregistrement du motif?',
          Ecran.Caption) = MrNo then
        begin
          LastError := 1;
          exit;
        end;
      end;
  end;
{ FIN PT17 }

  if (not IsPresenceMode) //PT22
    and (LastError = 0)
    and (Getfield('PMA_GESTIONIJSS') = 'X')
    and (Getfield('PMA_TYPEABS') = '') then
  begin
    PgiBox('Attention! Le type d''absence n''est pas renseigné.', Ecran.Caption);
    LastError := 1;
    SetFocusControl('PMA_TYPEABS');
  end;

  if (not IsPresenceMode) //PT22
    and (LastError = 0)
    and (GetField('PMA_RUBRIQUE') <> '')
    and (GetField('PMA_ALIMENT') = '') then
  begin
    PgiBox('Attention! Le type d''alimentation en heure n''est pas renseigné.',
      Ecran.Caption); { PT17 }
    LastError := 1;
    SetFocusControl('PMA_ALIMENT');
  end;

{ DEB PT17 }
  if (not IsPresenceMode) //PT22
    and (LastError = 0)
    and (GetField('PMA_RUBRIQUEJ') <> '')
    and (GetField('PMA_ALIMENTJ') = '') then
  begin
    PgiBox('Attention! Le type d''alimentation en jour n''est pas renseigné.',
      Ecran.Caption);
    LastError := 1;
    SetFocusControl('PMA_ALIMENTJ');
  end;
{ FIN PT17 }

//Rechargement des tablettes
  if (LastError = 0) and (Getfield('PMA_MOTIFABSENCE') <> '') and
    (Getfield('PMA_LIBELLE') <> '') then
    ChargementTablette(TFFiche(Ecran).TableName, '');

  Error := False;

  if (not IsPresenceMode) //PT22
    and (LastError = 0)
    and (GestionIJSS <> GetField('PMA_GESTIONIJSS')) then
// La coche GestionIJSS a été modifiée, il faut maj ABSENCESALARIE
// Toutes les absences ayant ce motif d'absence sont mises à jour(PCN_GESTIONIJSS);
  begin
{   try
      BeginTrans;}

    MajAbsenceSalarie(GetField('PMA_GESTIONIJSS'), GetField('PMA_MOTIFABSENCE'), GetField('PMA_CARENCEIJSS'));

{      ExecuteSql ('UPDATE ABSENCESALARIE SET'+
                  ' PCN_GESTIONIJSS="'+GetField ('PMA_GESTIONIJSS')+'" WHERE'+
                  ' PCN_TYPECONGE = "'+GetField ('PMA_MOTIFABSENCE')+'" AND'+
                  ' PCN_TYPECONGE IN (SELECT PMA_MOTIFABSENCE'+
                  ' FROM MOTIFABSENCE WHERE'+
                  ' ##PMA_PREDEFINI## PMA_GESTIONIJSS="'+GestionIJSS+'")');}
    GestionIJSS := GetField('PMA_GESTIONIJSS');
{      CommitTrans;
   except
      LastError:= 1;
      SetFocusControl ('PMA_GESTIONIJSS');
      Rollback;
      end;         }
  end;

{ DEB PT17 }
  if (not IsPresenceMode) //PT22
    and (LastError = 0)
    and (GetField('PMA_JOURHEURE') = 'JOU')
    and (Getfield('PMA_RUBRIQUEJ') = '') and (GetField('PMA_GERECOMM') = 'X') then
  begin
{ Dans le cas où on génère une ligne de commentaire sans rub. en jour affecté  }
    PgiBox('Attention ! Vous avez coché la gestion d''une ligne de commentaire' +
      ' dans le bulletin sans rubrique en jour liée.#13#10' +
      'Veuillez décocher la gestion ou saisir une rubrique.',
      Ecran.Caption);
    LastError := 1;
    SetFocusControl('PMA_GERECOMM');
  end;

  if (not IsPresenceMode) //PT22
    and (LastError = 0)
    and (GetField('PMA_JOURHEURE') = 'HEU')
    and (Getfield('PMA_RUBRIQUE') = '')
    and (GetField('PMA_GERECOMM') = 'X') then
  begin
{ Dans le cas où on génère une ligne de commentaire sans rub. en heure affecté  }
    PgiBox('Attention ! Vous avez coché la gestion d''une ligne de commentaire' +
      ' dans le bulletin sans rubrique en heure liée.#13#10' +
      'Veuillez décocher la gestion ou saisir une rubrique.',
      Ecran.Caption);
    LastError := 1;
    SetFocusControl('PMA_GERECOMM');
  end;

{ Gestion des zones alimentant le net à payer }
  if (not IsPresenceMode) //PT22
    and (LastError = 0) then
  begin
    if ControlCumulPaie(GetField('PMA_RUBRIQUE')) = True then
      SetField('PMA_ALIMNETH', 'X')
    else
      if GetField('PMA_ALIMNETH') <> '-' then
        SetField('PMA_ALIMNETH', '-');

    if ControlCumulPaie(GetField('PMA_RUBRIQUEJ')) = True then
      SetField('PMA_ALIMNETJ', 'X')
    else
      if GetField('PMA_ALIMNETJ') <> '-' then
        SetField('PMA_ALIMNETJ', '-');

    if (GetField('PMA_ALIMNETH') = 'X') and (GetField('PMA_ALIMNETJ') = 'X') then
    begin
{ Dans le cas où les 2 rubriques alimentent le net à payer }
      PgiBox('Attention ! Vous ne pouvez pas renseigner une rubrique en heure' +
        ' et en jour alimentant le net à payer.', Ecran.Caption);
      LastError := 1;
      SetFocusControl('PMA_RUBRIQUE');
    end;
  end;

{ Gestion maximum }
  if (LastError = 0) then
  begin
    if (GetField('PMA_GESTIONMAXI') = 'X') then
      if (GetField('PMA_JRSMAXI') <= 0) then
      begin
{ Dans le cas où gestion maxi et nombre de jour à 0 }
        PgiBox('Vous avez activé le gestion maximum sans renseigner un' +
          ' nombre de jour positif.', Ecran.Caption);
        LastError := 1;
        SetFocusControl('PMA_JRSMAXI');
      end;

    if (GetField('PMA_TYPEPERMAXI') = 'PER') and ((GetField('PMA_PERMAXI') < 1) or
      (GetField('PMA_PERMAXI') > 12)) then
    begin
{ Dans le cas où gestion maxi, début période et permaxi non comprise entre 1 et 12 }
      PgiBox('La périodicité de la gestion maximum doit être comprise entre 1' +
        ' et 12.', Ecran.Caption);
      LastError := 1;
      SetFocusControl('PMA_PERMAXI');
    end;
    if (GetField('PMA_TYPEPERMAXI') = 'GLI') and
      (GetField('PMA_PERMAXI') < 1) then
    begin
{ Dans le cas où gestion maxi, début période et permaxi non comprise entre 1 et 12 }
      PgiBox('La périodicité de la gestion maximum doit être supérieur à 0.',
        Ecran.Caption);
      LastError := 1;
      SetFocusControl('PMA_PERMAXI');
    end;
  end;
{ FIN PT17 }

//PT20
  if (not IsPresenceMode) //PT22
    and ((LastError = 0)
    and (GetField('PMA_JOURHEURE') <> 'HEU')
    and (Getfield('PMA_TYPEABS') = 'CHO')) then
  begin
{Dans le cas où Type absence = "Chômage Intempéries", Gestion en heure obligatoire}
    PgiBox('Attention ! Pour le chômage intempéries, la gestion en#13#10' +
      'heures est obligatoire.', Ecran.Caption);
    LastError := 1;
    SetFocusControl('PMA_JOURHEURE');
  end;
//FIN PT20
end;


procedure TOM_MOTIFABSENCE.OnLoadRecord;
var
{$IFNDEF EAGLCLIENT}
  PROFILABS: THDBValCombobox;
{$ELSE}
  PROFILABS: THValCombobox;
{$ENDIF}
  Zone: tcontrol;
  Q: TQuery;
begin
  inherited; { PT12 }
  if not (DS.State in [dsInsert]) then DerniereCreate := ''; { PT12 }
  Loadok := false;
  if IsPresenceMode then //PT22
  begin
   // InitialiseColor('PMA_PGCOLORPRE');
    //PT22
    InitialiseColor('PMA_PGCOLORATT');
    InitialiseColor('PMA_PGCOLORVAL');
    InitialiseColor('PMA_PGCOLORANN');
    InitialiseColor('PMA_PGCOLORREF');
  end else begin
{$IFNDEF EAGLCLIENT}
    PROFILABS := THDBValCombobox(getcontrol('PMA_PROFILABS'));
{$ELSE}
    PROFILABS := THValCombobox(getcontrol('PMA_PROFILABS'));
{$ENDIF}
    if PROFILABS <> nil then
      if PROFILABS.value = '' then
        PROFILABS.Itemindex := 0;
  { DEB PT13 }
{ DEB PT17 }
    InitialiseColor('PMA_PGCOLORACTIF');
    InitialiseColor('PMA_PGCOLORPAY');
    InitialiseColor('PMA_PGCOLORATT');
    InitialiseColor('PMA_PGCOLORVAL');
    InitialiseColor('PMA_PGCOLORANN');
    InitialiseColor('PMA_PGCOLORREF');
    GblProfil := GetField('PMA_PROFILABS');
    GblProfilJ := GetField('PMA_PROFILABSJ');
{ FIN PT17 }

    Zone := getcontrol('PMA_TYPEATTEST');
    if Zone <> nil then InitialiseCombo(Zone);
  end;


  Zone := getcontrol('PMA_JOURHEURE');
  if Zone <> nil then InitialiseCombo(Zone);

  AccesPredefini('TOUS', CEG, STD, DOS);
  LectureSeule := FALSE;
  if (Getfield('PMA_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False)); { PT12 }
    EnabledBtnColor (CEG); // PT26
    SetControlEnabled('BDelete', CEG);
  end;
  if (Getfield('PMA_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False)); { PT12 }
    EnabledBtnColor (STD); // PT26
    SetControlEnabled('BDelete', STD);
  end;
  if (Getfield('PMA_PREDEFINI') = 'DOS') then
  begin
    LectureSeule := (DOS = False); // PT26
    PaieLectureSeule(TFFiche(Ecran), False); { PT12 }
    EnabledBtnColor (DOS); // PT26
    SetControlEnabled('BDelete', DOS);
  end;

  SetControlEnabled('BInsert', True);
  SetControlEnabled('PMA_PREDEFINI', False);
  SetControlEnabled('PMA_MOTIFABSENCE', False);

  if DS.State in [dsInsert] then
  begin
    LectureSeule := FALSE;
    PaieLectureSeule(TFFiche(Ecran), False); { PT12 }
    SetControlEnabled('PMA_PREDEFINI', True);
    SetControlEnabled('PMA_MOTIFABSENCE', True);
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDelete', False);
  end;
  if (CEG = FALSE) and (STD = FALSE) then SetControlEnabled('BInsert', False);


  if (not IsPresenceMode) then //PT22
  begin
    { DEB PT5 Chargement du paramétrage }
    Q := OpenSql('SELECT PMA_MOTIFABSENCE,PMA_LIBELLE,PMA_RUBRIQUE,PMA_RUBRIQUEJ FROM MOTIFABSENCE ' + { PT17 }
      'WHERE ##PMA_PREDEFINI## PMA_RUBRIQUE<>"" ', True);
    if not Q.eof then
    begin
  // d PT9
      if (Assigned(T_Rub)) then
        FreeAndNil(T_Rub);
  // f PT9
      T_Rub := TOB.Create('Liste des rubriques motif', nil, -1);
      T_Rub.LoadDetailDB('Liste des rubriques motif', '', '', Q, False);
    end
    else
      T_Rub := nil;
    Ferme(Q);
    { FIN PT5 }
    GestionIJSS := GetField('PMA_GESTIONIJSS'); // PT8
    SetControlEnabled('PMA_CALENDSAL', False); { PT15 }
    if (GetControlText('LBLREM') = '') and (GetField('PMA_RUBRIQUE') <> '') then { PT15 }
      SetControlText('LBLREM', RechDom('PGREMUNERATION', GetField('PMA_RUBRIQUE'), False)); { PT15 }
  end;
end;

procedure TOM_MOTIFABSENCE.initialiseabsence;
var
{$IFNDEF EAGLCLIENT}
  rub: thdbEdit;
{$ELSE}
  rub: thEdit;
{$ENDIF}
begin
  if (not IsPresenceMode) then //PT22
  begin
{$IFNDEF EAGLCLIENT}
    rub := THdbEdit(getcontrol('PMA_RUBRIQUE'));
{$ELSE}
    rub := THEdit(getcontrol('PMA_RUBRIQUE'));
{$ENDIF}
    if rub <> nil then
      Rub.Text := '';

    setfield('PMA_RUBRIQUE', '');
    //setfield('PMA_PREDEFINI','');
    setfield('PMA_ALIMENT', '');
  //  setfield('PMA_GERECOMM', 'X'); { PT17 }
  end;
end;

procedure TOM_MOTIFABSENCE.EnabledChampsAbsence;
var
  profil, rubrique: boolean;
begin
  if (not IsPresenceMode) then //PT22
  begin
  { DEB PT17 }
    { Intégration dans le bulletin en heures }
    rubrique := (getfield('PMA_RUBRIQUE') <> '');
    setcontrolenabled('PMA_ALIMENT', rubrique);
  //  if GetField('PMA_JOURHEURE') = 'HEU' Then  setcontrolenabled('PMA_GERECOMM', rubrique);
    profil := (getfield('PMA_PROFILABS') <> '');
    if profil then
    begin
      SetControlProperty('PMA_RUBRIQUE', 'Datatype', 'PGRUBABSENCE');
      SetControlProperty('PMA_RUBRIQUE', 'Plus', getfield('PMA_PROFILABS'));
    end
    else
    begin
      SetControlProperty('PMA_RUBRIQUE', 'Datatype', 'PGREMUNERATION');
      SetControlProperty('PMA_RUBRIQUE', 'Plus', ' PRM_THEMEREM = "ABS"');
    end;

    { Intégration dans le bulletin en jours }
    rubrique := (getfield('PMA_RUBRIQUEJ') <> '');
    setcontrolenabled('PMA_ALIMENTJ', rubrique);
  //  if GetField('PMA_JOURHEURE') = 'JOU' Then setcontrolenabled('PMA_GERECOMM', rubrique);

    profil := (getfield('PMA_PROFILABSJ') <> '');
    if profil then
    begin
      SetControlProperty('PMA_RUBRIQUEJ', 'Datatype', 'PGRUBABSENCE');
      SetControlProperty('PMA_RUBRIQUEJ', 'Plus', getfield('PMA_PROFILABSJ'));
    end
    else
    begin
      SetControlProperty('PMA_RUBRIQUEJ', 'Datatype', 'PGREMUNERATION');
      SetControlProperty('PMA_RUBRIQUEJ', 'Plus', ' PRM_THEMEREM = "ABS"');
    end;
  { FIN PT17 }
  end;
end;

procedure TOM_MOTIFABSENCE.OnChangeField(F: TField);
var
  Pred: string;
begin
  inherited; { PT12 } { DEB PT15 }
  if IsPresenceMode then //PT22
  begin
   // SetControlEnabled('PGCOLORPRE', False);
    SetControlEnabled('PGCOLORATT', False);
    SetControlEnabled('PGCOLORVAL', False);
    SetControlEnabled('PGCOLORANN', False);
    SetControlEnabled('PGCOLORREF', False);
    //PT22
  end else begin
  { DEB PT17 }
    SetControlEnabled('PGCOLORACTIF', False);
    SetControlEnabled('PGCOLORPAY', False);
    SetControlEnabled('PGCOLORATT', False);
    SetControlEnabled('PGCOLORVAL', False);
    SetControlEnabled('PGCOLORANN', False);
    SetControlEnabled('PGCOLORREF', False);

    SetControlEnabled('PMA_ALIMNETH', False);
    SetControlEnabled('PMA_ALIMNETJ', False);
  { FIN PT17 }
  end;

  if F.Fieldname = 'PMA_GESTIONMAXI' then
  begin
    if (getfield('PMA_GESTIONMAXI') <> 'X') then
    begin
      if GetField('PMA_JRSMAXI') <> 0 then setfield('PMA_JRSMAXI', 0);
      if GetField('PMA_TYPEPERMAXI') <> '' then setfield('PMA_TYPEPERMAXI', '');
      if GetField('PMA_PERMAXI') <> 0 then setfield('PMA_PERMAXI', 0);
      Setcontrolenabled('PMA_JRSMAXI', False);
      Setcontrolenabled('PMA_TYPEPERMAXI', False);
      Setcontrolenabled('PMA_PERMAXI', False);
    end
    else
    begin
      Setcontrolenabled('PMA_JRSMAXI', True);
      Setcontrolenabled('PMA_TYPEPERMAXI', True);
      //PT22 Le champs PERMAXI ne doit pas etre accessible dans tous les cas
      if GetField('PMA_TYPEPERMAXI') = '' then
      begin
        if Getfield('PMA_PERMAXI') <> 0 then setfield('PMA_PERMAXI', 0);
        SetControlEnabled('PMA_PERMAXI', False);
      end
      else
        SetControlEnabled('PMA_PERMAXI', True);
//      Setcontrolenabled('PMA_PERMAXI', True);
      //Fin PT22
    end;
  end;
{ DEB PT17 }
  if (F.Fieldname = 'PMA_TYPEPERMAXI') then
  begin
    if GetField('PMA_TYPEPERMAXI') = '' then
    begin
      if Getfield('PMA_PERMAXI') <> 0 then setfield('PMA_PERMAXI', 0);
      SetControlEnabled('PMA_PERMAXI', False);
    end
    else
      SetControlEnabled('PMA_PERMAXI', True);
  end; { FIN PT15 }

  if not IsPresenceMode then //PT22
  begin
    if F.FieldName = 'PMA_PROFILABS' then
    begin
      if (getfield('PMA_PROFILABS') = '') and (getfield('PMA_PROFILABS') <> GblProfil) then
      begin
        if GetField('PMA_RUBRIQUE') <> '' then setfield('PMA_RUBRIQUE', '');
        if GetField('PMA_ALIMENT') <> '' then setfield('PMA_ALIMENT', '');
  //      if (GetField('PMA_JOURHEURE')='HEU') and (GetField('PMA_GERECOMM') <> 'X') then
  //        setfield('PMA_GERECOMM', 'X');
      end
      else
        if (getfield('PMA_PROFILABS') <> '') then GblProfil := getfield('PMA_PROFILABS');
      EnabledChampsAbsence;
    end;

    if F.FieldName = 'PMA_PROFILABSJ' then
    begin
      if (getfield('PMA_PROFILABSJ') = '') and (getfield('PMA_PROFILABSJ') <> GblProfilJ) then
      begin
        if GetField('PMA_RUBRIQUEJ') <> '' then setfield('PMA_RUBRIQUEJ', '');
        if GetField('PMA_ALIMENTJ') <> '' then setfield('PMA_ALIMENTJ', '');
  //      if (GetField('PMA_JOURHEURE')='JOU') and (GetField('PMA_GERECOMM') <> 'X') then
  //        setfield('PMA_GERECOMM', 'X');
      end
      else
        if (getfield('PMA_PROFILABSJ') <> '') then GblProfilJ := getfield('PMA_PROFILABSJ');
      EnabledChampsAbsence;
    end;
  { FIN PT17 }
  end;


  if (F.FieldName = 'PMA_PREDEFINI') and (DS.State = dsinsert) then
  begin
    Pred := GetField('PMA_PREDEFINI');
    if Pred = '' then exit else Error := False; {PT2 ajout else}
    AccesPredefini('TOUS', CEG, STD, DOS);
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      Error := True;
      LastError := 1;
      PGIBox('Vous ne pouvez créer de motif prédéfini CEGID', 'Accès refusé');
      Pred := '';
      SetControlProperty('PMA_PREDEFINI', 'Value', Pred);
    end;
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      Error := True;
      LastError := 1;
      PGIBox('Vous ne pouvez créer de motif prédéfini Standard', 'Accès refusé');
      Pred := '';
      SetControlProperty('PMA_PREDEFINI', 'Value', Pred);
    end;
    if Pred <> GetField('PMA_PREDEFINI') then SetField('PMA_PREDEFINI', pred);
  end;
{
  if (ds.state in [dsBrowse]) then
  begin
    if LectureSeule then
    begin
      PaieLectureSeule(TFFiche(Ecran), True);
      SetControlEnabled('BDelete', False);
    end;
    SetControlEnabled('BInsert', True);
    if (CEG = FALSE) and (STD = FALSE) then SetControlEnabled('BInsert', False);
    SetControlEnabled('PMA_PREDEFINI', False);
    SetControlEnabled('PMA_MOTIFABSENCE', False);
  end;
}
  if not IsPresenceMode then //PT22
  begin
    if (F.FieldName = 'PMA_RUBRIQUE') then { DEB PT17 }
      TChampRub := ControlZoneRubrique('');

    if (F.FieldName = 'PMA_RUBRIQUEJ') then
      TChampRub := ControlZoneRubrique('J');
  end;
 (*   if (GetField('PMA_RUBRIQUE') <> '') then
    begin
      Q := OpenSql('SELECT PRM_CODECALCUL,PRM_TYPEBASE,PRM_TYPEMONTANT,PRM_TYPETAUX,' +
        'PRM_TYPECOEFF FROM REMUNERATION ' +
        'WHERE PRM_RUBRIQUE="' + GetField('PMA_RUBRIQUE') + '"', True);
      if not Q.eof then
      begin
        CodeCalcul := Q.FindField('PRM_CODECALCUL').Asstring;
        TChampRub := '';
        if (CodeCalcul = '01') then
        begin
          if Q.FindField('PRM_TYPEBASE').AsString = '00' then TChampRub := 'BAS ';
          if Q.FindField('PRM_TYPEMONTANT').AsString = '00' then TChampRub := TChampRub + 'MON ';
        end
        else
          if (CodeCalcul = '02') or (CodeCalcul = '03') or (CodeCalcul = '06') or (CodeCalcul = '07') then
        begin
          if Q.FindField('PRM_TYPEBASE').AsString = '00' then TChampRub := 'BAS ';
          if Q.FindField('PRM_TYPETAUX').AsString = '00' then TChampRub := TChampRub + 'TAU ';
          if Q.FindField('PRM_TYPECOEFF').AsString = '00' then TChampRub := TChampRub + 'COE ';
        end
        else
          if (CodeCalcul = '04') or (CodeCalcul = '05') then
        begin
          if Q.FindField('PRM_TYPEBASE').AsString = '00' then TChampRub := 'BAS ';
          if Q.FindField('PRM_TYPETAUX').AsString = '00' then TChampRub := TChampRub + 'TAU ';
        end
        else
          if CodeCalcul = '08' then
        begin
          if Q.FindField('PRM_TYPEBASE').AsString = '00' then TChampRub := 'BAS ';
          if Q.FindField('PRM_TYPECOEFF').AsString = '00' then TChampRub := TChampRub + 'COE ';
        end;
        if TChampRub <> '' then //PT4 Ajout cond. pour bloque valeur par défaut systématiquement
          if Pos(GetControlText('PMA_ALIMENT'), TChampRub) = 0 then
            SetControlText('PMA_ALIMENT', GetControlText('PMA_ALIMENT'));
      end
      else
        Begin
        TChampRub := '';
        if GetField('PMA_ALIMENT')<>'' then  SetField('PMA_ALIMENT', '');
        End;
      Ferme(Q);
    end;*)
{ FIN PT17 }

  if not IsPresenceMode then //PT22
  begin
    if (F.FieldName = 'PMA_ALIMENT') then
    begin
      if (GetField('PMA_ALIMENT') <> '') and (TChampRub <> '') then
        if (Pos(GetField('PMA_ALIMENT'), TChampRub) = 0) then
        begin
          PgiBox('Pour la rubrique en heure, vous ne pouvez pas renseigner ce type d''alimentation.', Ecran.caption); { PT17 }
          SetField('PMA_ALIMENT', ''); { PT17 }
          SetControlText('PMA_ALIMENT', '');
        end;
    end;
  { DEB PT17 }
    if (F.FieldName = 'PMA_ALIMENTJ') then
    begin
      if (GetField('PMA_ALIMENTJ') <> '') and (TChampRub <> '') then
        if (Pos(GetField('PMA_ALIMENTJ'), TChampRub) = 0) then
        begin
          PgiBox('Pour la rubrique en jour, vous ne pouvez pas renseigner ce type d''alimentation.', Ecran.caption);
          SetField('PMA_ALIMENTJ', '');
          SetControlText('PMA_ALIMENTJ', '');
        end;
  { FIN PT17 }
    end;

    { DEB PT15 }
    if (F.Fieldname = 'PMA_CALENDCIVIL') then
    begin
      if GetField('PMA_CALENDCIVIL') = 'X' then
      begin
        if GetField('PMA_OUVRES') <> '-' then SetField('PMA_OUVRES', '-');
        if GetField('PMA_OUVRABLE') <> '-' then SetField('PMA_OUVRABLE', '-');
        if GetField('PMA_CALENDSAL') <> '-' then SetField('PMA_CALENDSAL', '-');
        SetControlEnabled('PMA_OUVRES', False);
        SetControlEnabled('PMA_OUVRABLE', False);
      end
      else
      begin
        if GetField('PMA_CALENDSAL') <> 'X' then SetField('PMA_CALENDSAL', 'X');
        SetControlEnabled('PMA_OUVRES', True);
        SetControlEnabled('PMA_OUVRABLE', True);
      end;
    end;

    if (F.Fieldname = 'PMA_OUVRES') then
    begin
      if GetField('PMA_OUVRES') = 'X' then
      begin
        if GetField('PMA_CALENDCIVIL') <> '-' then SetField('PMA_CALENDCIVIL', '-');
        if GetField('PMA_OUVRABLE') <> '-' then SetField('PMA_OUVRABLE', '-');
        if GetField('PMA_CALENDSAL') <> '-' then SetField('PMA_CALENDSAL', '-');
        SetControlEnabled('PMA_CALENDCIVIL', False);
        SetControlEnabled('PMA_OUVRABLE', False);
      end
      else
      begin
        if GetField('PMA_CALENDSAL') <> 'X' then SetField('PMA_CALENDSAL', 'X');
        SetControlEnabled('PMA_CALENDCIVIL', True);
        SetControlEnabled('PMA_OUVRABLE', True);
      end;
    end;

    if (F.Fieldname = 'PMA_OUVRABLE') then
    begin
      if GetField('PMA_OUVRABLE') = 'X' then
      begin
        if GetField('PMA_CALENDCIVIL') <> '-' then SetField('PMA_CALENDCIVIL', '-');
        if GetField('PMA_OUVRES') <> '-' then SetField('PMA_OUVRES', '-');
        if GetField('PMA_CALENDSAL') <> '-' then SetField('PMA_CALENDSAL', '-');
        SetControlEnabled('PMA_CALENDCIVIL', False);
        SetControlEnabled('PMA_OUVRES', False);
      end
      else
      begin
        if GetField('PMA_CALENDSAL') <> 'X' then SetField('PMA_CALENDSAL', 'X');
        SetControlEnabled('PMA_CALENDCIVIL', True);
        SetControlEnabled('PMA_OUVRES', True);
      end;
    end;
    { FIN PT15 }
  // d PT16
    if (F.FieldName = 'PMA_GESTIONIJSS') then
    begin
      SetcontrolVisible('PMA_CARENCEIJSS', (getfield('PMA_GESTIONIJSS') = 'X'));
      SetcontrolEnabled('PMA_CARENCEIJSS', (getfield('PMA_GESTIONIJSS') = 'X'));
      SetcontrolVisible('TPMA_CARENCEIJSS', (getfield('PMA_GESTIONIJSS') = 'X'));
      if (getfield('PMA_GESTIONIJSS') = 'X') then
        CalculCarenceIJSS(GetControlText('PMA_TYPEABS'));
    end;
  // f PT16
  end;

  if (ds.state in [dsBrowse]) then
  begin
    if LectureSeule then
    begin
      PaieLectureSeule(TFFiche(Ecran), True); { PT12 }
      SetControlEnabled('BDelete', False);
    end;
    SetControlEnabled('BInsert', True);
    if (CEG = FALSE) and (STD = FALSE) then SetControlEnabled('BInsert', False);
    SetControlEnabled('PMA_PREDEFINI', False);
    SetControlEnabled('PMA_MOTIFABSENCE', False);
  end;

end;

procedure TOM_MOTIFABSENCE.OnDeleteRecord;
begin
  inherited;
  ChargementTablette(TFFiche(Ecran).TableName, ''); { PT12 }
//PT23
  Trace := TStringList.Create; //PT25
  Trace.Add('SUPPRESSION MOTIF D''ABSENCE ' + GetField('PMA_MOTIFABSENCE') + ' ' + GetField('PMA_LIBELLE'));
  CreeJnalEvt('003', '087', 'OK', nil, nil, Trace);
  FreeAndNil(Trace); //PT25  Trace.free;
end;

procedure TOM_MOTIFABSENCE.OnNewRecord;
begin
  inherited;
  SetField('PMA_PREDEFINI', 'STD'); {PT1} {PT2}
  // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
  SetField('PMA_NODOSSIER', PgRendNoDossier()); {PT1}

  if IsPresenceMode then //PT22
  begin
    SetField('PMA_JOURHEURE', 'HEU');
    SetField('PMA_CONTROLMOTIF', '-');
   // SetField('PMA_PGCOLORPRE', '10944422');
    SetField('PMA_TYPEMOTIF', 'PRE');
    SetField('PMA_PGCOLORATT', '10210815');
    SetField('PMA_PGCOLORVAL', '10944422');
    SetField('PMA_PGCOLORANN', '8421631');
    SetField('PMA_PGCOLORREF', '8421631');
    SetField('PMA_PGCOLORACTIF', '10944422');
    SetField('PMA_PGCOLORPAY', '10944422');
  end else begin //Fin PT22
    // PT3 V563 PH  21/11/2001  Valeur par défaut JOURHEURE
    SetField('PMA_JOURHEURE', 'JOU');
    SetField('PMA_CONTROLMOTIF', '-'); { PT14 } { PT19 }
    SetField('PMA_CALENDSAL', 'X'); { PT15 }
    SetControlEnabled('PMA_CALENDSAL', False); { PT15 }

    { DEB PT21 }
    SetField('PMA_PGCOLORACTIF', '10944422');
    SetField('PMA_PGCOLORPAY', '16053248');
    SetField('PMA_PGCOLORATT', '10210815');
    SetField('PMA_PGCOLORVAL', '10944422');
    SetField('PMA_PGCOLORANN', '8421631');
    SetField('PMA_PGCOLORREF', '8421631');
    { FIN PT21 }
    SetField('PMA_TYPEMOTIF', 'ABS'); //PT22
  end;
end;

procedure TOM_MOTIFABSENCE.OnClose;

begin
  inherited;
  if (not IsPresenceMode) and (LastError = 0) and (Assigned(T_Rub)) then FreeAndNil(T_Rub); //PT5
end;
{ DEB PT12 }

procedure TOM_MOTIFABSENCE.OnAfterUpdateRecord;
var
  even: boolean; //PT23
begin
  inherited;
  Trace := TStringList.Create; //PT25
  even := IsDifferent(dernierecreate, PrefixeToTable(TFFiche(Ecran).TableName), 'PMA_MOTIFABSENCE', TFFiche(Ecran).LibelleName, Trace, TFFiche(Ecran)); //PT23  Trace.free;
  FreeAndNil(Trace); //PT25
  if OnFerme then Ecran.Close;
end;
{ FIN PT12 }

{ DEB PT13  }

procedure TOM_MOTIFABSENCE.ClickPgColor(Sender: TObject);
var
  Col, St, StChamp, BtnName: string;
  Colors: TColorDialog;
  EditCol: THEdit;
{$IFDEF EAGLCLIENT}
  Edit: THEdit;
{$ELSE}
  Edit: THDBEdit;
{$ENDIF}
begin
  if Assigned(Sender) then
    if Sender is TToolBarButton97 then
      St := TToolBarButton97(Sender).Name;

  Colors := TColorDialog.Create(Ecran);
  if Colors = nil then exit;

  if (Colors.Execute) then
  begin
    Col := ColorToString(Colors.Color);
    if St = 'BTNPGCOLORACTIF' then StChamp := 'PMA_PGCOLORACTIF' else
      if St = 'BTNPGCOLORPAY' then StChamp := 'PMA_PGCOLORPAY' else
        if St = 'BTNPGCOLORATT' then StChamp := 'PMA_PGCOLORATT' else
          if St = 'BTNPGCOLORVAL' then StChamp := 'PMA_PGCOLORVAL' else
            if St = 'BTNPGCOLORANN' then StChamp := 'PMA_PGCOLORANN' else
              if St = 'BTNPGCOLORREF' then StChamp := 'PMA_PGCOLORREF' ;//else
               // if St = 'BTNPGCOLORPRE' then StChamp := 'PMA_PGCOLORPRE'; //PT22
    BtnName := Copy(StChamp, 5, length(StChamp)); { PT17 }
    EditCol := THEdit(GetControl(BtnName)); { PT17 }
{$IFDEF EAGLCLIENT}
    Edit := THEdit(GetControl(StChamp));
{$ELSE}
    Edit := THDBEdit(GetControl(StChamp));
{$ENDIF}
    if Assigned(Edit) then
    begin
      Edit.Font.Color := Colors.Color;
      if DS.State in [DsBrowse] then Ds.Edit;
      if GetField(StChamp) <> Col then
      begin { DEB PT17 }
        SetField(StChamp, Col);
        if Assigned(EditCol) then EditCol.Color := Colors.Color;
      end; { FIN PT17 }
    end;
  end;
  Colors.Free;
end;
{ FIN PT13  }

{ DEB PT13  }

procedure TOM_MOTIFABSENCE.ClickPgDefaire(Sender: TObject);
var St, StChamp, StVal, BtnName: string;
  EditCol: THEdit;
{$IFDEF EAGLCLIENT}
  Edit: THEdit;
{$ELSE}
  Edit: THDBEdit;
{$ENDIF}
begin
  if Assigned(Sender) then
    if Sender is TToolBarButton97 then
      St := TToolBarButton97(Sender).Name;
  if St = 'BTNDEFAIREACTIF' then begin StChamp := 'PMA_PGCOLORACTIF'; stVal := '10944422' end else
    if St = 'BTNDEFAIREPAY' then begin StChamp := 'PMA_PGCOLORPAY'; stVal := '16053248' end else
      if St = 'BTNDEFAIREATT' then begin StChamp := 'PMA_PGCOLORATT'; stVal := '10210815' end else
        if St = 'BTNDEFAIREVAL' then begin StChamp := 'PMA_PGCOLORVAL'; stVal := '10944422' end else
          if St = 'BTNDEFAIREANN' then begin StChamp := 'PMA_PGCOLORANN'; stVal := '8421631' end else
            if St = 'BTNDEFAIREREF' then begin StChamp := 'PMA_PGCOLORREF'; stVal := '8421631' end ;//else
            //  if St = 'BTNDEFAIREPRE' then begin StChamp := 'PMA_PGCOLORPRE'; stVal := '10944422' end; //PT22
  BtnName := Copy(StChamp, 5, length(StChamp)); { PT17 }
  EditCol := THEdit(GetControl(BtnName)); { PT17 }
{$IFDEF EAGLCLIENT}
  Edit := THEdit(GetControl(StChamp));
{$ELSE}
  Edit := THDBEdit(GetControl(StChamp));
{$ENDIF}
  if Assigned(Edit) then
  begin
    Edit.Font.Color := StringToColor(StVal);
    if DS.State in [DsBrowse] then Ds.Edit;
    if GetField(StChamp) <> StVal then
    begin { DEB PT17 }
      SetField(StChamp, StVal);
      if Assigned(EditCol) then EditCol.Color := StringToColor(StVal);
    end; { FIN PT17 }
  end;
end;
{ FIN PT13  }
// d PT16
{***********A.G.L.***********************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 14/02/2006
Modifié le ... : 14/02/2006
Description .. : procédure ChangeGestIJSS
Suite ........ : Evènement click sur case à cocher gestion IJSS
Suite ........ : Affichage  ou non du nombre de jours de carence IJSS
Mots clefs ... : PAIE;ABSENCE;IJSS
*****************************************************************}

procedure TOM_MOTIFABSENCE.ChangeGestIJSS(Sender: TObject);
begin
  if GetCheckBoxState('PMA_GESTIONIJSS') = CbChecked then
  begin
    SetcontrolVisible('PMA_CARENCEIJSS', true);
    SetcontrolEnabled('PMA_CARENCEIJSS', true);
    SetcontrolVisible('TPMA_CARENCEIJSS', true);
    CalculCarenceIJSS(GetControlText('PMA_TYPEABS'));
  end
  else
  begin
    SetcontrolVisible('PMA_CARENCEIJSS', false);
    SetcontrolEnabled('PMA_CARENCEIJSS', false);
    SetcontrolVisible('TPMA_CARENCEIJSS', false);
  end;
end;

procedure TOM_MOTIFABSENCE.CalculCarenceIJSS(TypeAbs: string);
begin
  if (TypeAbs = 'ATJ') then
      // Accident de trajet
    SetField('PMA_CARENCEIJSS', '1')
  else
    if (TypeAbs = 'ATR') then
      // Accident du travail
      SetField('PMA_CARENCEIJSS', '1')
    else
      if (TypeAbs = 'MAN') then
      // Maladie non professionnelle
        SetField('PMA_CARENCEIJSS', '3')
      else
        if (TypeAbs = 'MAP') then
      // Maladie professionnelle
          SetField('PMA_CARENCEIJSS', '1')
        else
          if (TypeAbs = 'MAT') then
      // Maternité
            SetField('PMA_CARENCEIJSS', '0')
          else
            if (TypeAbs = 'PAT') then
      // Paternité
              SetField('PMA_CARENCEIJSS', '0')
            else
              if (TypeAbs = 'THE') then
      // Mi-temps thérapeutique
                SetField('PMA_CARENCEIJSS', '0');
end;
// f PT16
{ DEB PT17 }

procedure TOM_MOTIFABSENCE.InitialiseColor(ControlName: string);
var
  EditColor: THEdit;
{$IFDEF EAGLCLIENT}
  Edit: THEdit;
{$ELSE}
  Edit: THDBEdit;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
  Edit := THEdit(GetControl(ControlName));
{$ELSE}
  Edit := THDBEdit(GetControl(ControlName));
{$ENDIF}
  if Assigned(Edit) then
  begin
    if Edit.text <> '' then
    begin
      Edit.Font.Color := StringToColor(Edit.Text);
      EditColor := THEdit(GetControl(Copy(ControlName, 5, length(ControlName))));
      if Assigned(EditColor) then
      begin
        EditColor.Color := StringToColor(Edit.Text);
        EditColor.Enabled := False;
      end;
    end;
  end;
end;

function TOM_MOTIFABSENCE.ControlZoneRubrique(StDif: string): string;
var
  Q: TQuery;
  CodeCalcul, st: string;
begin
  if (Trim(GetField('PMA_RUBRIQUE' + StDif)) <> '') then
  begin
    Q := OpenSql('SELECT PRM_CODECALCUL,PRM_TYPEBASE,PRM_TYPEMONTANT,PRM_TYPETAUX,' +
      'PRM_TYPECOEFF FROM REMUNERATION ' +
      'WHERE PRM_RUBRIQUE="' + GetField('PMA_RUBRIQUE' + StDif) + '"', True);
    if not Q.eof then
    begin
      CodeCalcul := Q.FindField('PRM_CODECALCUL').Asstring;
      Result := '';
      if (CodeCalcul = '01') then
      begin
        if Q.FindField('PRM_TYPEBASE').AsString = '00' then Result := 'BAS ';
        if Q.FindField('PRM_TYPEMONTANT').AsString = '00' then Result := Result + 'MON ';
      end
      else
        if (CodeCalcul = '02') or (CodeCalcul = '03') or (CodeCalcul = '06') or (CodeCalcul = '07') then
        begin
          if Q.FindField('PRM_TYPEBASE').AsString = '00' then Result := 'BAS ';
          if Q.FindField('PRM_TYPETAUX').AsString = '00' then Result := Result + 'TAU ';
          if Q.FindField('PRM_TYPECOEFF').AsString = '00' then Result := Result + 'COE ';
        end
        else
          if (CodeCalcul = '04') or (CodeCalcul = '05') then
          begin
            if Q.FindField('PRM_TYPEBASE').AsString = '00' then Result := 'BAS ';
            if Q.FindField('PRM_TYPETAUX').AsString = '00' then Result := Result + 'TAU ';
          end
          else
            if CodeCalcul = '08' then
            begin
              if Q.FindField('PRM_TYPEBASE').AsString = '00' then Result := 'BAS ';
              if Q.FindField('PRM_TYPECOEFF').AsString = '00' then Result := Result + 'COE ';
            end;
      if Result <> '' then //PT4 Ajout cond. pour bloque valeur par défaut systématiquement
        if Pos(GetControlText('PMA_ALIMENT' + StDif), Result) = 0 then
          if GetField('PMA_ALIMENT' + StDif) <> '' then SetField('PMA_ALIMENT' + StDif, '');
    end
    else
    begin
        { DEB PT18 }
      if StDif <> '' then st := 'en jour' else st := 'en heure';
      PgiBox('Rubrique ' + St + ' inexistante. Veuillez saisir une autre rubrique.', Ecran.caption);
      if GetField('PMA_RUBRIQUE' + StDif) <> '' then SetField('PMA_RUBRIQUE' + StDif, '');
      SetFocusControl('PMA_RUBRIQUE' + StDif);
        { FIN PT18 }
      Result := '';
    end;
    Ferme(Q);
  end
  else
    if GetField('PMA_ALIMENT' + StDif) <> '' then SetField('PMA_ALIMENT' + StDif, '');
  SetControlEnabled('PMA_ALIMENT' + StDif, (GetField('PMA_RUBRIQUE' + StDif) <> ''));
end;

function TOM_MOTIFABSENCE.ControlCumulPaie(Rub: string): Boolean;
var
  Q: TQuery;
  St: string;
  Tob_Cum, T_Cum: Tob;
begin
  Result := False;
  St := 'SELECT * FROM CUMULRUBRIQUE WHERE ##PCR_PREDEFINI## ' +
    'PCR_NATURERUB="AAA" AND PCR_RUBRIQUE="' + Rub + '"';
  Q := OpenSql(St, True);
  if not Q.eof then
  begin
    Tob_Cum := TOB.Create('les cumuls', nil, -1);
    Tob_Cum.LoadDetailDB('les cumuls', '', '', Q, False);
  end;
  Ferme(Q);
  if Assigned(Tob_Cum) and (Tob_Cum.Detail.count > 0) then
  begin
    T_Cum := Tob_Cum.FindFirst(['PCR_CUMULPAIE'], ['09'], False);
    if Assigned(T_Cum) then result := True;
  end;
  FreeAndNil(Tob_Cum);
end;
{ FIN PT17 }
// d PT24

procedure TOM_MOTIFABSENCE.MajAbsenceSalarie(PmaGestionIJSS, PmaMotifAbs: string; CarenceIJSS: integer);
var
  Tob_Abs, TAbs: TOB;
  st: string;
  Q: TQuery;
  i: integer;
  WCarence: double;
  WDateFin, WDateDeb: TDateTime;

begin
  Tob_Abs := Tob.Create('Absences IJSS', nil, -1);
{  st := 'SELECT PCN_SALARIE, PCN_DATEDEBUTABS, PCN_DATEFINABS, ' +
        'PCN_LIBELLE, PCN_NBJCARENCE,PCN_GESTIONIJSS ' +}
  st := 'SELECT * ' +
    'FROM ABSENCESALARIE ' +
    'WHERE PCN_TYPECONGE = "' + PmaMotifAbs + '" ' +
    'AND ' +
    'PCN_TYPECONGE IN (SELECT PMA_MOTIFABSENCE ' +
    'FROM MOTIFABSENCE WHERE ' +
    '##PMA_PREDEFINI## PMA_GESTIONIJSS="' + GestionIJSS + '") ' +
    'ORDER BY PCN_SALARIE, PCN_DATEFINABS';
  Q := OpenSql(st, TRUE);
  if not (Q.eof) then
    Tob_Abs.LoadDetailDB('ABSENCESALARIE', '', '', Q, False);
  ferme(Q);

  if Assigned(Tob_Abs) then
  begin
    WDateFin := IDate1900;
    WCarence := 0;

    // détermination de la carence cumulée depuis 12 mois (WCarence)
    for i := 0 to Tob_Abs.Detail.Count - 1 do
    begin
      TAbs := Tob_Abs.detail[i];
      TAbs.PutValue('PCN_GESTIONIJSS', PmaGestionIJSS);

      // calcul du nombre de jours calendaires d'absence
      if TAbs.GetValue('PCN_NBJCALEND') <> (TAbs.GetValue('PCN_DATEFINABS') - TAbs.GetValue('PCN_DATEDEBUTABS') + 1) then
      begin
        TAbs.PutValue('PCN_NBJCALEND',
          TAbs.GetValue('PCN_DATEFINABS') - TAbs.GetValue('PCN_DATEDEBUTABS') + 1);
        TAbs.PutValue('PCN_NBJCARENCE', CarenceIJSS);
        TAbs.PutValue('PCN_IJSSSOLDEE', '-');
      end;

      if (TAbs.GetValue('PCN_DATEDEBUTABS') <> WDateFin + 1) then
        WCarence := TAbs.GetValue('PCN_NBJCARENCE')
      else
        WCarence := WCarence + TAbs.GetValue('PCN_NBJCARENCE');

      if (WCarence > CarenceIJSS) then
        WCarence := CarenceIJSS;

      // calcul du nbre de jours de carence
      WDateDeb := TAbs.GetValue('PCN_DATEDEBUTABS');
      if (WDateDeb = WDateFin + 1) then
      begin
        if (TAbs.GetValue('PCN_NBJCARENCE') <> (CarenceIJSS - WCarence)) then
          TAbs.PutValue('PCN_NBJCARENCE', CarenceIJSS - WCarence);
      end;

      if (TAbs.GetValue('PCN_NBJCALEND') < TAbs.GetValue('PCN_NBJCARENCE')) then
        TAbs.PutValue('PCN_NBJCARENCE', TAbs.GetValue('PCN_NBJCALEND'));

      // calcul du nombre de jours d'IJSS
      if (TAbs.GetValue('PCN_NBJIJSS') <> (TAbs.GetValue('PCN_NBJCALEND') - TAbs.GetValue('PCN_NBJCARENCE'))) then // PT91
        TAbs.PutValue('PCN_NBJIJSS',
          TAbs.GetValue('PCN_NBJCALEND') - TAbs.GetValue('PCN_NBJCARENCE'));

      if (Valeur(TAbs.GetValue('PCN_NBJIJSS')) = 0) then
        TAbs.PutValue('PCN_IJSSSOLDEE', 'X')
      else
        TAbs.PutValue('PCN_IJSSSOLDEE', '-');

      WDateFin := TAbs.GetValue('PCN_DATEFINABS');
    end;
  end;

  if (Tob_Abs <> nil) then
  begin
    // Mise à jour de la table ABSENCESALARIE
    // ================================
    try
      BeginTrans;

      if (Tob_Abs.IsOneModifie) then
        Tob_Abs.InsertOrUpdateDB(TRUE);

      CommitTrans;
    except
      Rollback;
    end;
  end;

  FreeAndNil(Tob_Abs);

end;
// f PT24
// DEB PT26
procedure TOM_MOTIFABSENCE.EnabledBtnColor(LeMode: Boolean);
var TabSht: TTabSheet;
    LeControl : TControl;
    i : Integer;
    LeNom : String;
begin
  TabSht := TTabSheet(Ecran.FindComponent('TBCOMPLEMENT'));
  if TabSht <> nil then
  begin
    for i := 0 to Ecran.ComponentCount - 1 do
    begin
      LeControl := TControl(Ecran.Components[i]);
      if (LeControl is TToolBarButton97) then
      begin
        LeNom := LeControl.Name;
        if (StrLen (PChar(LeNom)) >= 10) then LeNom := Copy (LeNom, 1, 10);
        if (LeNom = 'BTNPGCOLOR') OR (LeNom = 'BTNDEFAIRE') then
          LeControl.Enabled := LeMode;
      end;
    end;
  end;
end;
// FIN PT26
initialization
  registerclasses([TOM_MOTIFABSENCE]);
end.

