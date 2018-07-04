{***********UNITE*************************************************
Auteur  ...... : ph
Créé le ...... : 23/08/2001
Modifié le ... : 30/08/2001
Description .. : Source TOM de la TABLE : VISITEMEDTRAV
Suite ........ : (VISITEMEDTRAV)
Mots clefs ... : PAIE;VISITEMEDTRAV
*****************************************************************
PT- 1 : 09/10/2001 : JL 562: Vérification des dates d'inaptitude et de reclassement par rapport à celle de la visite dans procédure OnUpdateRecord
PT- 2 : 29/10/2001 : JL 562: gestion de l'affichage du code salarié
PT- 3 : 29/10/2001 : JL 562: changement de la gestion de l'enregistrement de la médecine du travail
PT- 4 : 21/11/2001 : JL 562: contrôle : date de visite inférieur àla date de sortie
PT- 5 : 26/03/2002 : JL 571: Ajout procedure pour envoyer e-mail et affichage des dates de mail et convocation
PT- 6 : 26/03/2002 : JL 571: Récupération de l'heure saisie dans THEdit
PT- 7 : 17/12/2002 : PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
                     JL      Et modification du format de l'heure pour le setfield
PT- 8 : 10/09/2003 : JL V_42 FQ 10777 Affichage liste salariés de l'établissement sélectionné ds mul
PT- 9 : 09/10/2003 : JL V_42 FQ 10562 Rafraichissement des données après MAJ date envoi mail
PT-10 : 11/05/2004 : JL V_50 FQ 11293 Correction controle sur modif heures
PT-11 : 28/06/2004 : JL V_50 FQ 11296 Gestion nouveau champ apte sous réserve
PT-12 : 28/07/2004 : JL V_50 FQ 11352 Affichage code médecine si saisie manuel code
PT-13 : 22/04/2005 : PH V_60 Filtrage salarié confidentialité établissement
---- JL 20/03/2006 modification clé annuaire ----
PT-14 : 17/08/2006 : JL V_65 FQ 13391 Modif test date sortie pour cas du 30/12/1899
PT-15 : 10/01/2007 : FCO V_80 FQ 13326 Gestion de la périodicité des visites médicales
PT-16 : 26/01/2007 : FC V_80 Mise en place du filtage habilitation pour les lookuplist
                        pour les critères code salarié uniquement
PT-17 : 31/07/2007 : JL V_80 FQ 14633 Création depuis fiche salarié
PT18  : 02/01/2008 : PH V_80 UpdateCaption  pour titre de la fenêtre
}
unit VISITEMEDTRAV_TOM;

interface

uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fiche, DBCtrls, HDB, FE_Main,
  {$ELSE}
  eFiche, eFichList, UtileAGL, MainEAGL,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOM, UTob, HTB97, EntPaie, PGOutils, MailOL, LookUp,PGOutils2,
  ParamSoc,PGCalendrier,graphics, // PT15
  extctrls; // PT15   util pout TBevel

type
  TOM_VISITEMEDTRAV = class(TOM)
    BValider: TToolbarButton97;
    procedure OnNewRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
  private
    EtablissementMul, medecinecode, HeureEntre: string; // PT 3
    PeriodVisitDefaut : String;  // PT15
    NbZoneLibreCombo : integer;  // PT15
    NbZoneLibreMontant: integer; // PT15
    LeSalarie : String;
    procedure Etab(Sender: Tobject);
    procedure ChangeDate(Sender: Tobject);
    procedure ExitEdit(Sender: TObject); //  PT- 2
    procedure EnvoyerMail(Sender: TObject); //PT- 5
    procedure ChangementHeure(Sender: TObject); //PT- 6
    procedure SalarieElipsisClick(Sender: TObject);
    procedure BConvocClick(Sender: TObject);
    procedure RendVisibleZonesLibres;
  end;

implementation

uses P5Def;

procedure TOM_VISITEMEDTRAV.OnNewRecord;
begin
  inherited;

  SetField('PVM_DATEVISITE', Date);
  // PT7 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
  SetField('PVM_DATEINA', IDate1900);
  SetField('PVM_DATEREC', IDate1900);
  SetControlEnabled('PVM_DATEINA', False);
  SetControlEnabled('PVM_DATEREC', False);
  SetControlEnabled('BMAIL', False);
  SetControlEnabled('BCONVOC', False);
  SetField('PVM_TYPMEDTRAV','ETB');
  //DEBUT PT-17
  If LeSalarie <> '' then
  begin
    SetField('PVM_SALARIE',LeSalarie);
    SetControlEnabled('PVM_SALARIE',False);
    LeSalarie := '';
  end;
  //FIN PT-17
end;

// PT- 1
procedure TOM_VISITEMEDTRAV.OnUpdateRecord;
var
  DateIna, Daterec, DateVisite: TDateTime;
  DatePrevVisitMed: TDateTime;
  Sal, TestDate: string;
  Q: TQuery;
  NbrAjout: integer;   // PT15
  TypAjout: string;    // PT15
  Tob_Ferie: Tob;      // PT15
  i: integer;          // PT15
begin
  inherited;
  DateVisite := GetField('PVM_DATEVISITE');
  if GetControlText('PVM_INAPTE') = 'X' then
  begin
    DateIna := GetField('PVM_DATEINA');
    if DateIna < DateVisite then
    begin
      LastError := 1;
      PGIBox('La date d''inaptitude ne peut être inférieure à la date de la visite', Ecran.Caption);
      Exit;
    end;
  end;
  if GetControlText('PVM_RECLASSE') = 'X' then
  begin
    DateRec := GetField('PVM_DATEREC');
    if DateRec < DateVisite then
    begin
      LastError := 1;
      PGIBox('La date de reclassement ne peut être inférieure à la date de la visite', Ecran.Caption);
      Exit;
    end;
  end;
  //PT- 4
  Sal := GetField('PVM_SALARIE');
  Q := OpenSQL('SELECT PSA_DATESORTIE FROM SALARIES WHERE PSA_SALARIE="' + Sal + '" AND PSA_DATESORTIE IS NOT NULL AND PSA_DATESORTIE>"' + UsDateTime(idate1900) + '"' + //PT-14
    ' AND PSA_DATESORTIE<"' + UsDateTime(DateVisite) + '"', True);
  testDate := '';
  if not Q.eof then TestDate := Q.FindField('PSA_DATESORTIE').AsString; //portageCWAS
  Ferme(Q);
  if TestDate <> '' then
  begin
    LastError := 1;
    PGIBox('La date de visite ne peut être supérieure à la date de sortie du salarié', Ecran.Caption);
    Exit;
  end;
  if GetControlText('HEUREVISITE') <> '' then SetField('PVM_HEUREVISITE', StrToTime(GetControlText('HEUREVISITE'))); //PT- 7
  if DS.State = DsInsert then
  begin
    SetControlEnabled('BMAIL', True);
    SetControlEnabled('BCONVOC', True);
  end;

  // DEB PT15
  // Calculer la date prévisionnelle de la prochaine visite en fonction de la périodicité
  if (GetControlText('PVM_PERIODVISITMED') = '3') then
  begin
    NbrAjout := 1;
    TypAjout := 'A';
  end;
  if (GetControlText('PVM_PERIODVISITMED') = '4') then
  begin
    NbrAjout := 2;
    TypAjout := 'A';
  end;
  if (GetControlText('PVM_PERIODVISITMED') = '5') then
  begin
    NbrAjout := 6;
    TypAjout := 'M';
  end;
  if (GetControlText('PVM_PERIODVISITMED') = '6') then
  begin
    NbrAjout := 3;
    TypAjout := 'M';
  end;
  DatePrevVisitMed := PlusDate(DateVisite,NbrAjout,TypAjout);
  // Vérifier qu'il s'agit d'un jour ouvré.
  if not Assigned(Tob_Ferie) then Tob_Ferie := ChargeTobFerie(DatePrevVisitMed,PlusDate(DatePrevVisitMed,10,'J'));
  i := 10;
  While i > 0 do
    Begin
    if (IfJourFerie(Tob_Ferie,DatePrevVisitMed)) OR (DayOfWeek(DatePrevVisitMed) in [1,7]) then
        Begin
        DatePrevVisitMed := PlusDate(DatePrevVisitMed,1,'J');
        Continue;
        End
    else
      Break;
    End;
  Tob_Ferie.Free;

  SetField ('PVM_DATEVISITMED', DatePrevVisitMed);
  // FIN PT15
end;


procedure TOM_VISITEMEDTRAV.OnAfterUpdateRecord;
begin
  inherited;
  //PT 3 suppression du update
end;

procedure TOM_VISITEMEDTRAV.OnLoadRecord;
var
  Medecine, Sal: string;
  Mail, Convoc: string;
  QMedecine: TQuery;
  DateVisite, DateMail, DateConvoc: TDateTime;
begin
  inherited;
  sal := GetField('PVM_SALARIE');
  DateVisite := GetField('PVM_DATEVISITE');
  DateMail := IDate1900;
  DateConvoc := IDate1900;
  if sal <> '' then
  begin
    QMedecine := OpenSQL('Select PVM_MEDTRAVGU,PVM_MTENVOIMAIL,PVM_MTDATEMAIL,PVM_EDITCONVOC,PVM_DATECONVOC' +
      ' FROM VISITEMEDTRAV WHERE PVM_SALARIE="' + Sal + '" AND PVM_DATEVISITE="' + UsDateTime(Datevisite) + '"', True);
    Medecine := '';
    Mail := '-';
    Convoc := '-';

    if not QMedecine.eof then
    begin
      Medecine := QMedecine.FindField('PVM_MEDTRAVGU').AsString; //portageCWAS // DB2
      Mail := QMedecine.FindField('PVM_MTENVOIMAIL').AsString; //PT- 5
      Convoc := QMedecine.FindField('PVM_EDITCONVOC').AsString;
      DateMail := QMedecine.FindField('PVM_MTDATEMAIL').AsDateTime;
      DateConvoc := QMedecine.FindField('PVM_DATECONVOC').AsDateTime;
    end;
    Ferme(QMedecine);
    // PT 3
    if (Medecine <> '') and (Medecine <> '-1') then
    begin
      SetControlText('MEDTRAV', Medecine);
    end
    else
    begin
      SetControlText('MEDTRAV', '');
    end;
  end;
  MedecineCode := GetControlText('MEDTRAV');

  //PT- 5
  if Mail = 'X' then SetControlCaption('LIBMAIL', 'E-mail envoyé le ' + DateToStr(DateMail))
  else SetControlCaption('LIBMAIL', 'Aucun e-mail envoyé');
  if Convoc = 'X' then SetControlCaption('LIBCONVOC', 'Convocation éditée le ' + DateToStr(DateConvoc))
  else SetControlCaption('LIBCONVOC', 'Aucune convocation éditée');

  if GetControlText('PVM_APTE') = 'X' then
  begin
    SetControlEnabled('PVM_INAPTE', False);
    SetControlEnabled('PVM_RECLASSE', False);
    SetControlChecked('PVM_INAPTE', False);
    SetControlChecked('PVM_RECLASSE', False);
  end;
  if GetControlText('PVM_INAPTE') = 'X' then
  begin
    SetControlEnabled('PVM_APTE', False);
    SetControlEnabled('PVM_RECLASSE', False);
    SetControlChecked('PVM_APTE', False);
    SetControlChecked('PVM_RECLASSE', False);
  end;
  if GetControlText('PVM_RECLASSE') = 'X' then
  begin
    SetControlEnabled('PVM_INAPTE', False);
    SetControlEnabled('PVM_APTE', False);
    SetControlChecked('PVM_INAPTE', False);
    SetControlChecked('PVM_APTE', False);
  end;
  if TFFiche(Ecran).FTypeAction <> taCreat then //PT- 6
  begin
    HeureEntre := FormatDateTime('hh:mm', GetField('PVM_HEUREVISITE'));
    SetControlText('HEUREVISITE', HeureEntre);
    Ecran.Caption := 'Visite médicale de ' + RechDom('PGSALARIE', GetField('PVM_SALARIE'), False);
    UpdateCaption(TFFiche(Ecran));  // PT18
  end;
end;

procedure TOM_VISITEMEDTRAV.OnChangeField(F: TField);
var
  {$IFNDEF EAGLCLIENT}
  ComboTypeVisiteMed: THDBValComboBox;
  ComboPeriodVisiteMed: THDBValComboBox;
  {$ELSE}
  ComboTypeVisiteMed: THValComboBox;
  ComboPeriodVisiteMed: THValComboBox;
  {$ENDIF}
begin
  inherited;
// DEB PT13
  if (F.FieldName = 'PVM_SALARIE') then
  begin
    if (PGRendEtabUser() <> '') And (GetField ('PVM_SALARIE') <> '') then
    begin
      if (RechDom ('PGSALARIE', GetField ('PVM_SALARIE'), FALSE) = '') then
      begin
        PgiError ('Le salarié est inconnu', Ecran.Caption);
        SetField ('PVM_SALARIE', '');
      end;
    end;
  end;
// FIN PT13

  // DEB PT15
  // En fonction du type de visite, on préaffiche la périodicité des visites médicales
  // Elle n'est modifiable que dans 2 cas (embauche et reprise d'activité) et dans ce cas
  // elle est renseignée à partir du paramètre société PGPERIODVISITMED
  if (F.FieldName = 'PVM_TYPEVISITMED') then
  begin
    {$IFNDEF EAGLCLIENT}
      ComboTypeVisiteMed := THDBValComboBox(GetControl('PVM_TYPEVISITMED'));
      if (ComboTypeVisiteMed.Value <> '') then
      begin
        ComboPeriodVisiteMed := THDBValComboBox(GetControl('PVM_PERIODVISITMED'));
        if (ComboTypeVisiteMed.Value = '1') or (ComboTypeVisiteMed.Value = '2') then
        begin
          SetControlVisible('PVM_PERIODVISITMED', true);
          SetControlVisible('TPVM_PERIODVISITMED', true);
          SetControlEnabled('PVM_PERIODVISITMED', true);
          ComboPeriodVisiteMed.Value := PeriodVisitDefaut;
        end
        else
        begin
            SetControlVisible('PVM_PERIODVISITMED', false);
            SetControlVisible('TPVM_PERIODVISITMED', false);
            SetControlEnabled('PVM_PERIODVISITMED', false);
            ComboPeriodVisiteMed.Value := ComboTypeVisiteMed.Value;
        end;
      end
      else
      begin
          SetControlVisible('PVM_PERIODVISITMED', false);
          SetControlVisible('TPVM_PERIODVISITMED', false);
          SetControlEnabled('PVM_PERIODVISITMED', false);
      end;
    {$ELSE}
      ComboTypeVisiteMed := THValComboBox(GetControl('PVM_TYPEVISITMED'));
      if (ComboTypeVisiteMed.Value <> '') then
      begin
        ComboPeriodVisiteMed := THValComboBox(GetControl('PVM_PERIODVISITMED'));
        if (ComboTypeVisiteMed.Value = '1') or (ComboTypeVisiteMed.Value = '2') then
        begin
          SetControlVisible('PVM_PERIODVISITMED', true);
          SetControlVisible('TPVM_PERIODVISITMED', true);
          SetControlEnabled('PVM_PERIODVISITMED', true);
          ComboPeriodVisiteMed.Value := PeriodVisitDefaut;
        end
        else
        begin
            SetControlVisible('PVM_PERIODVISITMED', false);
            SetControlVisible('TPVM_PERIODVISITMED', false);
            SetControlEnabled('PVM_PERIODVISITMED', false);
            ComboPeriodVisiteMed.Value := ComboTypeVisiteMed.Value;
        end;
      end
      else
      begin
          SetControlVisible('PVM_PERIODVISITMED', false);
          SetControlVisible('TPVM_PERIODVISITMED', false);
          SetControlEnabled('PVM_PERIODVISITMED', false);
      end;
    {$ENDIF}


  end;
  // FIN PT15
end;

procedure TOM_VISITEMEDTRAV.OnArgument(S: string);
var
  {$IFNDEF EAGLCLIENT}
  Check1, Check2, Check3, Check4: TDBCheckBox;
  ComboEtab: THDBValComboBox;
  sal: THDBEdit;
  {$ELSE}
  Check1, Check2, Check3, Check4: TCheckBox;
  ComboEtab: THValComboBox;
  sal: THEdit;
  {$ENDIF}
  THMedecine, THHeure: THEdit;
  BMail, BConvoc: TToolBarButton97;
begin
  inherited;
  PeriodVisitDefaut  := GetParamSocSecur('SO_PGPERIODVISITMED',''); // PT15
  NbZoneLibreCombo   := GetParamSocSecur('SO_PGVMNBLIBCMB','');     // PT15
  NbZoneLibreMontant := GetParamSocSecur('SO_PGVMNBLIBDBL','');     // PT15
  EtablissementMul := ReadTokenPipe(S, ';');
  LeSalarie := ReadTokenPipe(S, ';'); //PT-17
  {$IFNDEF EAGLCLIENT}
  sal := THDBEdit(GetControl('PVM_SALARIE'));
  ComboEtab := THDBValComboBox(GetControl('PVM_TYPMEDTRAV'));
  Check1 := TDBCheckBox(GetControl('PVM_APTE'));
  Check2 := TDBCheckBox(GetControl('PVM_INAPTE'));
  Check3 := TDBCheckBox(GetControl('PVM_RECLASSE'));
  Check4 := TDBCheckBox(GetControl('PVM_APTESR'));  // PT-11
  {$ELSE}
  sal := THEdit(GetControl('PVM_SALARIE'));
  ComboEtab := THValComboBox(GetControl('PVM_TYPMEDTRAV'));
  Check1 := TCheckBox(GetControl('PVM_APTE'));
  Check2 := TCheckBox(GetControl('PVM_INAPTE'));
  Check3 := TCheckBox(GetControl('PVM_RECLASSE'));
  Check4 := TCheckBox(GetControl('PVM_APTESR'));  // PT-11
  {$ENDIF}

  if sal <> nil then
  begin
//    sal.Onchange := Etab;
    sal.OnExit := ExitEdit; // PT- 2
    Sal.OnElipsisClick := SalarieElipsisClick;
  end;
  if ComboEtab <> nil then
  begin
    ComboEtab.Onchange := Etab;
  end;
  if check1 <> nil then Check1.OnClick := ChangeDate;
  if check2 <> nil then Check2.onclick := ChangeDate;
  if check3 <> nil then Check3.onclick := ChangeDate;
  if check4 <> nil then Check4.onclick := ChangeDate;
  BMail := TToolBarButton97(GetControl('BMAIL')); // PT- 5
  if BMail <> nil then BMail.OnClick := EnvoyerMail;
  THHeure := THEdit(GetControl('HEUREVISITE')); //PT- 6
  if THHeure <> nil then THHeure.OnExit := ChangementHeure;
  BConvoc := TToolBarButton97(GetControl('BCONVOC'));
  if BConvoc <> nil then BConvoc.OnClick := BConvocClick;

  RendVisibleZonesLibres; //PT15
end;

procedure TOM_VISITEMEDTRAV.Etab(Sender: Tobject);
var
  {$IFNDEF EAGLCLIENT}
  ComboEtab: THDBValComboBox;
  ComboMaj: THDBEdit;
  {$ELSE}
  ComboEtab: THValComboBox;
  ComboMaj: THEdit;
  {$ENDIF}
  salarie, etab, medecine, nommedecine: string;
  Q1, Q2, Q3: TQuery;
  nommed: THLabel;
begin
  {$IFNDEF EAGLCLIENT}
  ComboEtab := THDBValComboBox(GetControl('PVM_TYPMEDTRAV'));
//  ComboMaj := THDBEdit(GetControl('MEDTRAV'));
  {$ELSE}
  ComboEtab := THValComboBox(GetControl('PVM_TYPMEDTRAV'));
//  ComboMaj := THEdit(GetControl('MEDTRAV'));
  {$ENDIF}
  salarie := GetField('PVM_SALARIE');
  Q1 := OpenSQL('Select PSA_ETABLISSEMENT FROM SALARIES WHERE PSA_SALARIE="' + salarie + '"', TRUE);
  Etab := '';
  if not Q1.eof then Etab := Q1.findfield('PSA_ETABLISSEMENT').AsString; //portageCWAS
  Ferme(Q1);
  Q2 := OpenSQL('Select ETB_MEDTRAVGU FROM ETABCOMPL WHERE ETB_ETABLISSEMENT="' + Etab + '"', TRUE);
  medecine := '';
  if not Q2.eof then
    medecine := Q2.FindField('ETB_MEDTRAVGU').AsString; //portageCWAS // DB2
  Ferme(Q2);
  if (ComboEtab <> nil) then
  begin
    if ComboEtab.value = 'ETB' then
    begin
      //PT 3 : Le code n'est récupéré que si il n'est pas = à -1
        SetField('PVM_MEDTRAVGU', medecine);

//      ComboMaj.enabled := FALSE;
       SetControlenabled('PVM_MEDTRAVGU',False);

    end
    else
      //ComboMaj.enabled := True;
      begin
      SetControlenabled('PVM_MEDTRAVGU',tRUE);
      SetControlText('PVM_MEDTRAVGU', '');
     end;
  end;
  NomMedecine := '';
  if Medecine <> '' then
  begin
    Q3 := OpenSQL('SELECT ANN_NOMPER FROM ANNUAIRE WHERE ANN_GUIDPER="' + medecine + '"', TRUE); // DB2
    if not Q3.eof then nommedecine := Q3.FindField('ANN_NOMPER').AsString; //portageCWAS
    Ferme(Q3);
  end;
  nommed := THLabel(GetControl('NOMMED'));
  NOMMED.caption := nommedecine;
end;


procedure TOM_VISITEMEDTRAV.ChangeDate(Sender: Tobject);
begin
  if sender = nil then exit;
  {$IFNDEF EAGLCLIENT}
  if TDBCheckBox(sender).Name = 'PVM_APTE' then
    {$ELSE}
  if TCheckBox(sender).Name = 'PVM_APTE' then
    {$ENDIF}
  begin
    if GetControlText('PVM_APTE') = 'X' then
    begin
      SetControlEnabled('PVM_INAPTE', False);
      SetControlEnabled('PVM_RECLASSE', False);
      SetControlenabled('PVM_DATEINA', False);
      SetControlChecked('PVM_INAPTE', False);
      SetControlChecked('PVM_RECLASSE', False);
      // PT7 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
      SetField('PVM_DATEINA', Idate1900);
      SetControlenabled('PVM_DATEREC', False);
      SetField('PVM_DATEREC', IDate1900);
      SetControlEnabled('PVM_APTESR', False);  // PT-11
      SetControlChecked('PVM_APTESR', False);
    end
    else
    begin
      SetControlEnabled('PVM_INAPTE', True);
      SetControlEnabled('PVM_RECLASSE', True);
      SetControlEnabled('PVM_APTESR', True);           // PT-11
    end;
  end;
  {$IFNDEF EAGLCLIENT}
  if TDBCheckBox(sender).Name = 'PVM_INAPTE' then
    {$ELSE}
  if TCheckBox(sender).Name = 'PVM_INAPTE' then
    {$ENDIF}
  begin
    if GetControlText('PVM_INAPTE') = '-' then
    begin
      SetControlenabled('PVM_DATEINA', False);
      // PT7 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
      SetField('PVM_DATEINA', IDate1900);
    end
    else
    begin
      SetControlenabled('PVM_DATEINA', True);
      SetField('PVM_DATEINA', Date);
    end;
    if GetControlText('PVM_INAPTE') = 'X' then
    begin
      SetControlEnabled('PVM_APTE', False);
      SetControlEnabled('PVM_RECLASSE', False);
      SetControlChecked('PVM_APTE', False);
      SetControlChecked('PVM_RECLASSE', False);
      SetControlEnabled('PVM_APTESR',False); // PT-11
      SetControlChecked('PVM_APTESR',False);
    end
    else
    begin
      SetControlEnabled('PVM_APTE', True);
      SetControlEnabled('PVM_APTESR', True);         // PT-11
      SetControlEnabled('PVM_RECLASSE', True);
    end;
  end;
  {$IFNDEF EAGLCLIENT}
  if TDBCheckBox(sender).name = 'PVM_RECLASSE' then
    {$ELSE}
  if TCheckBox(sender).name = 'PVM_RECLASSE' then
    {$ENDIF}
  begin
    if GetControlText('PVM_RECLASSE') = '-' then
    begin
      SetControlenabled('PVM_DATEREC', False);
      // PT7 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
      SetField('PVM_DATEREC', IDate1900);
    end
    else
    begin
      SetControlenabled('PVM_DATEREC', True);
      SetField('PVM_DATEREC', Date);
    end;
    if GetControlText('PVM_RECLASSE') = 'X' then
    begin
      SetControlEnabled('PVM_INAPTE', false);
      SetControlEnabled('PVM_APTE', false);
      SetControlChecked('PVM_INAPTE', false);
      SetControlChecked('PVM_APTE', False);
      SetControlEnabled('PVM_APTESR',False); // PT-11
      SetControlChecked('PVM_APTESR',False);
    end
    else
    begin
      SetControlEnabled('PVM_INAPTE', True);
      SetControlEnabled('PVM_APTE', True);
      SetControlEnabled('PVM_APTESR', True);         // PT-11
    end;
  end;
  {$IFNDEF EAGLCLIENT}
  if TDBCheckBox(sender).Name = 'PVM_APTESR' then  // PT-11
    {$ELSE}
  if TCheckBox(sender).Name = 'PVM_APTESR' then
    {$ENDIF}
  begin
    if GetControlText('PVM_APTESR') = 'X' then
    begin
      SetControlEnabled('PVM_INAPTE', False);
      SetControlEnabled('PVM_RECLASSE', False);
      SetControlenabled('PVM_DATEINA', False);
      SetControlChecked('PVM_INAPTE', False);
      SetControlChecked('PVM_RECLASSE', False);
      // PT7 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
      SetField('PVM_DATEINA', Idate1900);
      SetControlenabled('PVM_DATEREC', False);
      SetField('PVM_DATEREC', IDate1900);
      SetControlEnabled('PVM_APTE', False);
      SetControlChecked('PVM_APTE', False);
    end
    else
    begin
      SetControlEnabled('PVM_INAPTE', True);
      SetControlEnabled('PVM_RECLASSE', True);
      SetControlEnabled('PVM_APTE', True);
    end;
  end;
end;

// PT- 2
procedure TOM_VISITEMEDTRAV.ExitEdit(Sender: TObject);
var
  edit: thedit;
  QSAL: TQuery;
  NomSalarie, Salarie: string;
begin
  edit := THEdit(Sender);
  if edit <> nil then //AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
    begin
      Salarie := AffectDefautCode(edit, 10);
      SetField('PVM_SALARIE', Salarie);
      edit.text := Salarie;
    end;
  QSAL := OpenSQL('SELECT PSA_LIBELLE FROM SALARIES WHERE PSA_SALARIE="' + edit.text + '"', TRUE);
  nomsalarie := '';
  if not QSAL.eof then nomsalarie := QSAL.FindField('PSA_LIBELLE').AsString; //portageCWAS
  Ferme(QSAL);
  Ecran.caption := 'Saisie visite médicale de ' + NomSalarie;
  UpdateCaption(TFFiche(Ecran));  // PT18
  Etab(edit);
end;

procedure TOM_VISITEMEDTRAV.EnvoyerMail(Sender: TObject); // PT- 5
var
  Texte: HTStrings;
  Titre, Destinataire: string;
  HeureVisite, DateVisite, Adr1, Adr2, Adr3, CPVille, AdrNom: string;
  Q: TQuery;
begin
  AdrNom := '';
  Adr1 := '';
  Adr2 := '';
  Adr3 := '';
  CPVille := '';
  Q := OpenSQL('SELECT ANN_APRUE1,ANN_APRUE2,ANN_APRUE3,ANN_APCPVILLE,ANN_APNOM FROM ANNUAIRE WHERE' +
    ' ANN_GUIDPER="' + GetField('PVM_MEDTRAVGU') + '"', True); // DB2
  if not Q.eof then
  begin
    AdrNom := Q.FindField('ANN_APNOM').AsString; //portageCWAS
    Adr1 := Q.FindField('ANN_APRUE1').AsString;
    Adr2 := Q.FindField('ANN_APRUE2').AsString;
    Adr3 := Q.FindField('ANN_APRUE3').AsString;
    CPVille := Q.FindField('ANN_APCPVILLE').AsString;
  end;
  Ferme(Q);
  Destinataire := '';
  if ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_SALARIE="' + GetField('PVM_SALARIE') + '"') then
  begin
    Q := OpenSQL('SELECT PSE_EMAILPROF FROM DEPORTSAL WHERE PSE_SALARIE="' + GetField('PVM_SALARIE') + '"', true);
    destinataire := '';
    if not Q.eof then Destinataire := Q.FindField('PSE_EMAILPROF').AsString; //portageCWAS
    Ferme(Q);
  end;
  HeureVisite := FormatDateTime('hh "h" mm', GetField('PVM_HEUREVISITE'));
  DateVisite := FormatDateTime('dddd dd mmmm yyyy', GetField('PVM_DATEVISITE'));
  Titre := 'Convocation visite médicale '; //+RechDom('PGTYPVISITEMED',GetField('PGTYPVISITEMED'),false);
  Texte := HTStringList.Create;
  texte.Insert(0, 'Merci de prendre connaissance de votre prochaine date de visite médicale, qui aura lieu le ' + DateVisite + ' à ' + HeureVisite + '.');
  texte.Insert(1, 'Cette visite aura lieu à la médecine du travail suivante :');
  texte.Insert(2, '      ' + AdrNom);
  texte.Insert(3, '      ' + Adr1 + Adr2 + Adr3);
  texte.Insert(4, '      ' + CPVille);
  texte.Insert(5, 'Si vous vous trouvez dans l''impossibilité de vous rendre à ce rendez-vous, merci nous contacter pour nous communiquer vos dates de disponibilité.');
  texte.Insert(6, '');
  texte.Insert(7, 'Cordialement,');
  texte.Insert(8, 'Le Service du Personnel.');
  SendMail(Titre, Destinataire, '', Texte, '', false);
  Texte.free;
  ExecuteSQL('UPDATE VISITEMEDTRAV SET PVM_MTENVOIMAIL="X" ,PVM_MTDATEMAIL="' + UsDateTime(Date) + '" WHERE' +
    ' PVM_SALARIE="' + GetField('PVM_SALARIE') + '"');
  SetControlCaption('LIBMAIL', 'E-mail envoyé le ' + DateToStr(Date));
  RefreshDB; //PT- 9
end;

procedure TOM_VISITEMEDTRAV.ChangementHeure(Sender: TObject); //PT- 6
var
  NouvelHeure: string;
begin
  NouvelHeure := FormatDateTime('hh:mm', StrToTime(GetControlText('HEUREVISITE'))); // PT-10
  if NouvelHeure <> HeureEntre then ForceUpdate;
end;

//DEBUT PT- 8
procedure TOM_VISITEMEDTRAV.SalarieElipsisClick(Sender: TObject);
var
  StEtab: string;
begin
  if EtablissementMul = '' then StEtab := ''
  else StEtab := 'PSA_ETABLISSEMENT="' + EtablissementMul + '"';
  StEtab := RecupClauseHabilitationLookupList(StEtab);  //PT16

  {$IFNDEF EAGLCLIENT}
  LookupList(THDBEdit(Sender), 'Salarié', 'SALARIES', 'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', StEtab, 'PSA_SALARIE', True, -1);
  {$ELSE}
  LookupList(THEdit(Sender), 'Salarié', 'SALARIES', 'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', StEtab, 'PSA_SALARIE', True, -1);
  {$ENDIF}
end;
//FIN PT- 8

procedure TOM_VISITEMEDTRAV.BConvocClick(Sender: TObject);
begin
  AglLanceFiche('PAY', 'SALARIE_RTF', 'OBJET=MEDECINE;PSA_SALARIE=' + GetField('PVM_SALARIE') + ';DATEVISITE=' + DateToStr(GetField('PVM_DATEVISITE')) + ';DATEVISITE_=' +
    DateToStr(GetField('PVM_DATEVISITE')), '', 'MEDECINE');
  RefreshDB; //PT- 9
end;

//DEB PT15
procedure TOM_VISITEMEDTRAV.RendVisibleZonesLibres;
var
{$IFNDEF EAGLCLIENT}
  Combo: THDBValComboBox;
  Montant: THDBEdit;
{$ELSE}
  Combo: THValComboBox;
  Montant: THEdit;
{$ENDIF}
  Cadre: TBevel;
  num: integer;
  numero, libelle: string;
  TLieu: THLabel;
begin

  // Affichage des zones libres combo
  for num := 1 to NbZoneLibreCombo do
  begin
    Numero := InttoStr(num);
    if Num > NbZoneLibreCombo then break;
    Cadre := TBevel(GetControl('BVCOMBO'));
    if cadre <> nil then Cadre.visible := True;
{$IFNDEF EAGLCLIENT}
    Combo := THDbValComboBox(GetControl('PVM_LIBRECMB' + Numero));
{$ELSE}
    Combo := THValComboBox(GetControl('PVM_LIBRECMB' + Numero));
{$ENDIF}
    if Combo <> nil then Combo.Visible := TRUE;
    TLieu := THLabel(GetControl('TPVM_LIBRECMB' + Numero));
    if (TLieu <> nil) then
    begin
      TLieu.Visible := TRUE;
      if Num = 1 then
      begin
        TLieu.Caption := VH_Paie.PgLibCombo1;
      end;
      if Num = 2 then
      begin
        TLieu.Caption := VH_Paie.PgLibCombo2;
      end;
    end;
  end;

  //Affichage des zones libres montant
  for num := 1 to NbZoneLibreMontant do
  begin
    Numero := InttoStr(num);
    if Num > NbZoneLibreMontant then break;
    Cadre := TBevel(GetControl('BVMONTANT'));
    if cadre <> nil then Cadre.visible := True;
{$IFNDEF EAGLCLIENT}
    Montant := THDbEdit(GetControl('PVM_LIBREDBL' + Numero));
{$ELSE}
    Montant := THEdit(GetControl('PVM_LIBREDBL' + Numero));
{$ENDIF}
    if Combo <> nil then Montant.Visible := TRUE;
    TLieu := THLabel(GetControl('TPVM_LIBREDBL' + Numero));
    if (TLieu <> nil) then
    begin
      TLieu.Visible := TRUE;
      if Num = 1 then
      begin
        TLieu.Caption := VH_Paie.PgSalLib1;
      end;
      if Num = 2 then
      begin
        TLieu.Caption := VH_Paie.PgSalLib2;
      end;
      if Num = 3 then
      begin
        TLieu.Caption := VH_Paie.PgSalLib3;
      end;
    end;
  end;
end;
// FIN PT15

initialization
  registerclasses([TOM_VISITEMEDTRAV]);
end.

