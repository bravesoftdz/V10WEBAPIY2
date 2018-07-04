{***********UNITE*************************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 30/09/2002
Modifié le ... :   /  /
Description .. : Génération de calendriers mensuels d' absences salariés
Suite ........ : sous excel
Mots clefs ... : PAIE;ABSENCE
*****************************************************************}
{
PT1 17/05/2004 V_50 SB FQ 11298 Intégration de la notion multi sur jours fériés
PT2 02/03/2006 V_65 SB Modification de la base forfait
PT3 07/04/2006 V_65 SB FQ 12941 Traitement nouveau mode de décompte motif d'absence
PT4 19/06/2006 V_65 SB FQ 13231 Retrait des mvt absences annulées
}
unit UTOFPGCalendrierExcel;

interface

uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  Classes, HCtrls, Dialogs, sysutils, StdCtrls, Windows, Forms, Controls, Graphics,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  HTB97, UTOF, Utob, Paramsoc, ShellAPI, HDebug,
  Hent1, HmsgBox, UtilXls, ed_tools,ULibEditionPaie;

type

  TOF_PGCALENDRIER_EXCEL = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
    procedure OnClose; override;
  private
    { Déclarations privées }
    NomFichier: THEdit;
    Annee: THValComboBox;
    FontEntete: TFontDialog;
    FExcel, FExcelClasseur, FExcelSheet: OleVariant;
    Salarie, NomPrenom, LibEtab, responsable: string;
    Tob_Semaine, Tob_Standard, Tob_JourFerie, Tob_absence, Tob_MoisAbsence, Tob_Histo, Tob_AcquisRtt: Tob;
    Etab, Calend, StandCalend, EtabStandCalend, NbJrTravEtab, Repos1, Repos2, JourHeure: string;
    DebAnnee, FinAnnee, DateEntree, DateSortie, DtCloture: TDateTime;
    GblTotalNb_j, GblTotalCpPris, GblTotalRtt, GblTotalAss, GblMensCpPris, GblMensRtt, GblMensRts, GblMensAss, Jouvres: Double;
    Tob_MotifAbs : Tob;    { PT3 }
    procedure OuvrirClik(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
    procedure GenereCalendrier;
    procedure MiseEnFormeEntete(Col, Row: integer);
    procedure MiseEnFormeMois(Mois, Col, Row: integer);
    procedure RecupAbsenceMoisSalarie(Mois: integer);
    function IsJourAbsence(DateJour: TdateTime; var Lib: string): Double;
    procedure InitialiseTob_Absence(var tob_Abs: Tob);
    procedure MiseEnFormePiedDePage(Col, Row: integer);
    procedure ChargeTob(var Tob_Salarie: Tob);
    procedure IntegreAbsenceHistobulletin;
    procedure LibereTob;
    procedure MyExcelText(CurrentSheet: OleVariant; Ligne, Colonne: Integer; Value: string);
  end;

implementation

uses P5Def, PgCongesPayes, EntPaie, PGoutils2,PGCommun,PGCalendrier;

var
  JOURS: array[1..7] of string = ('Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam', 'Dim');


  { TOF_PGCALENDRIER_EXCEL }

procedure TOF_PGCALENDRIER_EXCEL.OnArgument(Arguments: string);
begin
  inherited;
  NomFichier := THEdit(GetControl('NOMFICHIER'));
  if NomFichier = nil then SetControlEnabled('BValider', False);
  {$IFDEF EAGLCLIENT}
  NomFichier.Text := VH_Paie.PgCheminEagl + '\Calendrier.xls';
  {$ELSE}
  NomFichier.Text := V_PGI.StdPath + '\Calendrier.xls';
  {$ENDIF}

  Annee := THValComboBox(GetControl('ANNEE'));
  if annee = nil then SetControlEnabled('BValider', False);
  InitialiseCombo(TControl(Annee));

  if GetControl('BOUVRIR') <> nil then
    TToolBarButton97(GetControl('BOUVRIR')).OnClick := OuvrirClik;

  VisibiliteChampSalarie('1', GetControl('PSA_TRAVAILN1'), GetControl('TPSA_TRAVAILN1'));
  VisibiliteChampSalarie('1', GetControl('PSA_TRAVAILN1_'), GetControl('TPSA_TRAVAILN1_'));
  SetControlText('TPSA_TRAVAILN1_', 'à');
  VisibiliteChampSalarie('2', GetControl('PSA_TRAVAILN2'), GetControl('TPSA_TRAVAILN2'));
  VisibiliteChampSalarie('2', GetControl('PSA_TRAVAILN2_'), GetControl('TPSA_TRAVAILN2_'));
  SetControlText('TPSA_TRAVAILN2_', 'à');
  VisibiliteChampSalarie('3', GetControl('PSA_TRAVAILN3'), GetControl('TPSA_TRAVAILN3'));
  VisibiliteChampSalarie('3', GetControl('PSA_TRAVAILN3_'), GetControl('TPSA_TRAVAILN3_'));
  SetControlText('TPSA_TRAVAILN3_', 'à');
  VisibiliteChampSalarie('4', GetControl('PSA_TRAVAILN4'), GetControl('TPSA_TRAVAILN4'));
  VisibiliteChampSalarie('4', GetControl('PSA_TRAVAILN4_'), GetControl('TPSA_TRAVAILN4_'));
  SetControlText('TPSA_TRAVAILN4_', 'à');
  VisibiliteStat(GetControl('PSA_CODESTAT'), GetControl('TPSA_CODESTAT'));
  VisibiliteStat(GetControl('PSA_CODESTAT_'), GetControl('TPSA_CODESTAT_'));
  SetControlText('TPSA_CODESTAT_', 'à');
  VisibiliteChampLibreSal('1', GetControl('PSA_LIBREPCMB1'), GetControl('TPSA_LIBREPCMB1'));
  VisibiliteChampLibreSal('1', GetControl('PSA_LIBREPCMB1_'), GetControl('TPSA_LIBREPCMB1_'));
  SetControlText('TPSA_LIBREPCMB1_', 'à');
  VisibiliteChampLibreSal('2', GetControl('PSA_LIBREPCMB2'), GetControl('TPSA_LIBREPCMB2'));
  VisibiliteChampLibreSal('2', GetControl('PSA_LIBREPCMB2_'), GetControl('TPSA_LIBREPCMB2_'));
  SetControlText('TPSA_LIBREPCMB2_', 'à');
  VisibiliteChampLibreSal('3', GetControl('PSA_LIBREPCMB3'), GetControl('TPSA_LIBREPCMB3'));
  VisibiliteChampLibreSal('3', GetControl('PSA_LIBREPCMB3_'), GetControl('TPSA_LIBREPCMB3_'));
  SetControlText('TPSA_LIBREPCMB3_', 'à');
  VisibiliteChampLibreSal('4', GetControl('PSA_LIBREPCMB4'), GetControl('TPSA_LIBREPCMB4'));
  VisibiliteChampLibreSal('4', GetControl('PSA_LIBREPCMB4_'), GetControl('TPSA_LIBREPCMB4_'));
  SetControlText('TPSA_LIBREPCMB4_', 'à');
  TCheckBox(GetControl('PSA_BOOLLIBRE1')).Caption := VH_PAie.PgLibCoche1;
  TCheckBox(GetControl('PSA_BOOLLIBRE1')).visible := (VH_PAie.PgLibCoche1 <> '');

  if GetControl('PSA_SALARIE') <> nil then THEdit(GetControl('PSA_SALARIE')).OnExit := ExitEdit;
  if GetControl('PSA_SALARIE_') <> nil then THEdit(GetControl('PSA_SALARIE_')).OnExit := ExitEdit;

  FontEntete := TFontDialog.Create(Application);

  Tob_MotifAbs := tob.create('tob_virtuelle', nil, -1);                          { PT3 }
  Tob_MotifAbs.loaddetaildb('MOTIFABSENCE', '', 'PMA_MOTIFABSENCE', nil, False); { PT3 }

end;

procedure TOF_PGCALENDRIER_EXCEL.OnUpdate;
begin
  inherited;
  GenereCalendrier;
  if PgiAsk('Traitement terminé. Voulez-vous ouvrir le fichier excel généré?', Ecran.Caption) = MrYes then
    ShellExecute(0, PCHAR('open'), PChar('Excel'), PChar(NomFichier.text), nil, SW_RESTORE);
end;

procedure TOF_PGCALENDRIER_EXCEL.OnClose;
begin
  inherited;
  if Tob_Semaine <> nil then
  begin
    Tob_Semaine.free;
    Tob_Semaine := nil;
  end;
  if Tob_Standard <> nil then
  begin
    Tob_Standard.free;
    Tob_Standard := nil;
  end;
  if Tob_JourFerie <> nil then
  begin
    Tob_JourFerie.free;
    Tob_JourFerie := nil;
  end;
  if Tob_absence <> nil then
  begin
    Tob_absence.free;
    Tob_Semaine := nil;
  end;
  if Tob_MoisAbsence <> nil then
  begin
    Tob_MoisAbsence.free;
    Tob_MoisAbsence := nil;
  end;
  if Tob_Histo <> nil then
  begin
    Tob_Histo.free;
    Tob_Histo := nil;
  end;
  if Tob_AcquisRtt <> nil then
  begin
    Tob_AcquisRtt.free;
    Tob_AcquisRtt := nil;
  end;

  if FontEntete <> nil then FontEntete.Destroy;

  FreeAndNil(Tob_MotifAbs);   { PT3 }
end;

{ Déclarations privées }

procedure TOF_PGCALENDRIER_EXCEL.OuvrirClik(Sender: TObject);
begin
  ShellExecute(0, PCHAR('open'), PChar('Excel'), PChar(NomFichier.text), nil, SW_RESTORE);
end;


procedure TOF_PGCALENDRIER_EXCEL.ExitEdit(Sender: TObject);
var
  Edit: THEdit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(THedit(edit), 10);
end;

procedure TOF_PGCALENDRIER_EXCEL.GenereCalendrier;
var
  WinNew: Boolean;
  i: integer;
  Tob_Sal: Tob;
  //  TT         : TQRProgressForm ; //PORTAGECWAS  hplanning
begin
  Salarie := '';
  if IsValidDate('01/01/' + Annee.text) then
  begin
    DebAnnee := StrToDate('01/01/' + Annee.text);
    FinAnnee := StrToDate('31/12/' + Annee.text)
  end
  else
  begin
    PgiBox('Le format de l''année de réference est erronnée.', Ecran.Caption);
    Exit;
  end;
  if NomFichier.Text = '' then
  begin
    PgiBox('Veuiller renseigner le nom du fichier à générer.', Ecran.Caption);
    Exit;
  end;

  if not OpenExcel(True, FExcel, WinNew) then
  begin
    Exit;
  end;
  try
    FExcelClasseur := OpenWorkBook(NomFichier.Text, FExcel, True);
    if not VarIsEmpty(FExcelClasseur) then
    begin
      InitMoveProgressForm(nil, 'Traitement en cours..', 'Veuillez patienter SVP ...', 100, TRUE, TRUE);
      {Chargement des Tobs à utiliser}
      ChargeTob(Tob_Sal);
      if Tob_Sal <> nil then
      begin
        if not MoveCurProgressForm('Calendrier Annuel') then
        begin
          ExcelClose(FExcel);
          exit;
        end;

        {Parcours de la liste des salariés}
        for i := 0 to Tob_Sal.Detail.Count - 1 do
        begin
          if not MoveCurProgressForm('Calendrier Annuel') then
          begin
            FExcelClasseur.SaveAs(NomFichier.Text);
            Break;
          end;
          Salarie := Tob_Sal.Detail[i].GetValue('PSA_SALARIE');
          NomPrenom := Tob_Sal.Detail[i].GetValue('PSA_LIBELLE') + ' ' + Tob_Sal.Detail[i].GetValue('PSA_PRENOM');
          LibEtab := Tob_Sal.Detail[i].GetValue('ETB_LIBELLE');
          DTcloture := Tob_Sal.Detail[i].GetValue('ETB_DATECLOTURECPN');
          responsable := RechDom('PGTRAVAILN4', Tob_Sal.Detail[i].GetValue('PSA_TRAVAILN4'), False);
          MoveCurProgressForm('Chargement des données du salarié : ' + Salarie + ' ' + RechDom('PGSALARIE', Salarie, False));
          ChargeTobCalendrier(DebAnnee, FinAnnee, Salarie, False, False, True, Tob_Semaine, Tob_Standard, Tob_JourFerie, Tob_Absence, Etab, Calend, StandCalend, EtabStandCalend, { 29/03/2005 SB Ajout param }
            NbJrTravEtab, Repos1, Repos2, JourHeure, DateEntree, DateSortie);
          MoveCurProgressForm('Création classeur Excel du salarié : ' + Salarie + ' ' + RechDom('PGSALARIE', Salarie, False));

          {Création d'un classeur par salarié}
          FExcelSheet := AddSheet(FExcel, FExcelClasseur, Salarie);
          if not VarIsEmpty(FExcelSheet) then
          begin
            MoveCurProgressForm('Mise en forme du classeur Excel du salarié : ' + Salarie + ' ' + RechDom('PGSALARIE', Salarie, False));
            {Mise en forme du classeur Excel du salarié}
            MiseEnFormeEntete(0, 0);
            MiseEnFormeMois(1, 0, 5); //Janvier
            MiseEnFormeMois(2, 17, 5); //Fevrier
            MiseEnFormeMois(3, 0, 17); //Mars
            MiseEnFormeMois(4, 17, 17); //Avril
            MiseEnFormeMois(5, 0, 29); //Mai
            if not MoveCurProgressForm('Mise en forme du classeur Excel du salarié : ' + Salarie + ' ' + RechDom('PGSALARIE', Salarie, False)) then
            begin
              FExcelClasseur.SaveAs(NomFichier.Text);
              Break;
            end;
            MiseEnFormeMois(6, 17, 29); //Juin
            MiseEnFormeMois(7, 0, 41); //Juillet
            MiseEnFormeMois(8, 17, 41); //Aout
            MiseEnFormeMois(9, 0, 53); //Septembre
            MiseEnFormeMois(10, 17, 53); //Octobre
            MiseEnFormeMois(11, 0, 65); //Novembre
            MiseEnFormeMois(12, 17, 65); //Décembre
            //           IntegreAbsenceGrille;
            MiseEnFormePiedDePage(0, 77);
            FExcelClasseur.SaveAs(NomFichier.Text);
          end;
          LibereTob;
        end;
      end;
    end;
  finally
    if Tob_Sal <> nil then
    begin
      Tob_Sal.free;
      Tob_Sal := nil;
    end;
    if Tob_Semaine <> nil then
    begin
      Tob_Semaine.free;
      Tob_semaine := nil;
    end;
    if Tob_Absence <> nil then
    begin
      Tob_Absence.free;
      Tob_Absence := nil;
    end;
    if Tob_JourFerie <> nil then
    begin
      Tob_JourFerie.free;
      Tob_JourFerie := nil;
    end;
    if Tob_Standard <> nil then
    begin
      Tob_Standard.free;
      Tob_Standard := nil;
    end;
    if Tob_MoisAbsence <> nil then
    begin
      Tob_MoisAbsence.free;
      Tob_MoisAbsence := nil;
    end;
    if Tob_Histo <> nil then
    begin
      Tob_Histo.free;
      Tob_Histo := nil;
    end;
    if Tob_AcquisRtt <> nil then
    begin
      Tob_AcquisRtt.free;
      Tob_AcquisRtt := nil;
    end;

    FiniMoveProgressForm;
    ExcelClose(FExcel);
  end;
end;

procedure TOF_PGCALENDRIER_EXCEL.MyExcelText(CurrentSheet: OleVariant; Ligne, Colonne: Integer; Value: string);
begin
  ExcelText(CurrentSheet, Ligne + 1, Colonne + 1, Value);
end;

procedure TOF_PGCALENDRIER_EXCEL.MiseEnFormeEntete(Col, Row: integer);
begin
  GblTotalNb_j := 0;
  Jouvres := 0;
  GblTotalCpPris := 0;
  GblTotalRtt := 0;
  GblTotalAss := 0;
  FontEntete.Font.Color := 12615680;
  FontEntete.Font.Style := [fsBold];
  ExcelAlignRange(FExcelSheet, Col, Row, Col + 33, Row + 5, taCenter);
  ExcelColorRange(FExcelSheet, Row, Col, Row + 5, Col + 33, clWhite, False);
  ExcelFontRange(FExcelSheet, Row, Col, Row + 5, Col + 33, FontEntete.Font);
  {Titre du document}
  FontEntete.Font.Size := 14;
  MyExcelText(FExcelSheet, Row + 2, Col, 'Calendrier annuel forfait jour');
  ExcelMergeRange(FExcelSheet, Row + 2, Col, Row + 2, Col + 33);
  ExcelFontRange(FExcelSheet, Row + 2, Col, Row + 2, Col + 33, FontEntete.Font);
  {Nom Societe}
  MyExcelText(FExcelSheet, Row, Col + 2, GetParamSoc('SO_LIBELLE'));
  {Nom Etablissement}
  MyExcelText(FExcelSheet, Row + 1, Col + 2, LibEtab);
  {Nom Fichier}
  MyExcelText(FExcelSheet, Row, 30, ExtractFileName(NomFichier.Text));
end;


procedure TOF_PGCALENDRIER_EXCEL.MiseEnFormeMois(Mois, Col, Row: integer);
var
  i, j, iColonne, iLigne, iJour, DerCol, DerRow, Jour: integer;
  StMois, lib: string;
  DateDebMois, DateJour: TDateTime;
  YY, MM, JJ: WORD;
  TotalNb_j, Nb_j, Nb_h, J_Abs: Double;
  AM, PM: Boolean;
begin
  TotalNb_j := 0;
  GblMensCpPris := 0;
  GblMensRtt := 0;
  GblMensRts := 0;
  GblMensAss := 0;
  DerCol := Col + 16;
  DerRow := Row + 11;
  for i := Col to DerCol do ExcelColWidth(FExcelSheet, i, 3); //2.14
  FontEntete.Font.Color := clBlack;
  FontEntete.Font.Style := [];
  FontEntete.Font.Size := 8;
  ExcelAlignRange(FExcelSheet, Col, Row, DerCol, DerRow, taCenter);
  ExcelColorRange(FExcelSheet, Row, Col, DerRow, DerCol, clWhite, False);
  { ENTETE : ANNEE }
  StMois := IntToStr(Mois);
  if length(StMois) = 1 then StMois := '0' + StMois;
  MyExcelText(FExcelSheet, Row, Col, RechDom('PGMOIS', StMois, False) + ' ' + Annee.Text);
  { ENTETE : JOURS }
  iColonne := Col;
  iLigne := Row + 2;
  for i := LOW(JOURS) to HIGH(JOURS) do
  begin
    Inc(iColonne, 2);
    MyExcelText(FExcelSheet, Iligne, iColonne, JOURS[I]);
    ExcelColorCell(FExcelSheet, iLigne, iColonne, 12615680, True);
  end;

  { DONNEES }
  if VH_Paie.PGRecupAbsHisto then
    if mois = 1 then IntegreAbsenceHistobulletin;
  RecupAbsenceMoisSalarie(Mois);
  DateDebMois := EnCodeDate(StrToInt(Annee.Text), Mois, 1);
  DecodeDate(FindeMois(DateDebMois), YY, MM, JJ);
  if DayOfWeek(DateDebMois) = 1 then Jour := 7 else Jour := DayOfWeek(DateDebMois) - 1;
  iJour := 1;
  iLigne := Row + 2;
  iColonne := Col + (Jour * 2) - 1;
  ;
  for i := 1 to 6 do
  begin
    Inc(iLIgne);
    for j := Jour to 7 do
    begin
      if Ijour > JJ then Break;
      MyExcelText(FExcelSheet, iLigne, iColonne, IntToStr(iJour));
      ExcelColorRange(FExcelSheet, iLigne, iColonne, iLigne, iColonne + 1, clWhite, True);
      ExcelFontRange(FExcelSheet, iLigne, iColonne, iLigne, iColonne + 1, FontEntete.Font);
      DateJour := EncodeDate(StrToInt(Annee.Text), Mois, Ijour);
      if IfJourNonTravaille(Tob_Semaine, Tob_Standard, Tob_JourFerie, DateJour, Idate1900, Idate1900, NbJrTravEtab, Repos1, Repos2, False, True, Nb_j, Nb_h, AM, PM) = False then
        JOuvres := Jouvres + nb_j; //Pour calcul jours ouvrés calendaire
      if IfJourNonTravaille(Tob_Semaine, Tob_Standard, Tob_JourFerie, DateJour, DateEntree, DateSortie, NbJrTravEtab, Repos1, Repos2, False, True, Nb_j, Nb_h, AM, PM) = True then
      begin
        ExcelColorCell(FExcelSheet, iLigne, iColonne, 13750737, True);
        ExcelColorCell(FExcelSheet, iLigne, iColonne + 1, 13750737, True);
      end
      else
      begin
        J_Abs := IsJourAbsence(DateJour, lib);
        if J_Abs > 0 then
        begin
          MyExcelText(FExcelSheet, iLigne, iColonne + 1, Lib);
          ExcelColorCell(FExcelSheet, iLigne, iColonne + 1, 16310947, True);
          if J_Abs < 1 then TotalNb_j := TotalNb_j + (1 - J_Abs);
        end
        else
          TotalNb_j := TotalNb_j + nb_j;
      end;
      Inc(iColonne, 2);
      Inc(iJour);
    end;
    Jour := 1;
    iColonne := Col + 1;
  end;
  { COMPTEUR ABSENCE RTT RTS PIED DE GRILLE}
  FontEntete.Font.Color := clWhite; //16310947
  FontEntete.Font.Style := [FsBold];
  MyExcelText(FExcelSheet, DerRow - 2, Col, FloatToStr(GblMensRtt) + ' JRTT pris');
  ExcelMergeRange(FExcelSheet, DerRow - 2, Col, DerRow - 2, Col + 2);
  ExcelFontRange(FExcelSheet, DerRow - 2, Col, DerRow - 2, Col + 2, FontEntete.Font);
  ExcelColorRange(FExcelSheet, DerRow - 2, Col, DerRow - 2, Col + 2, 12615680, True);
  GblTotalRtt := GblTotalRtt + GblMensRtt;
  MyExcelText(FExcelSheet, DerRow - 2, Col + 5, FloatToStr(GblMensRts) + ' JRTS pris');
  ExcelMergeRange(FExcelSheet, DerRow - 2, Col + 4, DerRow - 2, Col + 6);
  ExcelFontRange(FExcelSheet, DerRow - 2, Col + 4, DerRow - 2, Col + 6, FontEntete.Font);
  ExcelColorRange(FExcelSheet, DerRow - 2, Col + 4, DerRow - 2, Col + 6, 12615680, True);
  GblTotalRtt := GblTotalRtt + GblMensRts;
  { COMPTEUR ABSENCE PRI PIED DE GRILLE}
  Debug('Pied de grille CP mois ' + IntToStr(mois) + ': ' + FloattoStr(GblMensCpPris));
  MyExcelText(FExcelSheet, DerRow - 2, Col + 10, FloatToStr(GblMensCpPris) + ' CP pris');
  ExcelMergeRange(FExcelSheet, DerRow - 2, Col + 8, DerRow - 2, Col + 10);
  ExcelFontRange(FExcelSheet, DerRow - 2, Col + 8, DerRow - 2, Col + 10, FontEntete.Font);
  ExcelColorRange(FExcelSheet, DerRow - 2, Col + 8, DerRow - 2, Col + 10, 12615680, True);
  GblTotalCpPris := GblTotalCpPris + GblMensCpPris;
  { COMPTEUR JOURS TRAVAILLES PIED DE GRILLE}
  MyExcelText(FExcelSheet, DerRow - 2, Col + 12, FloatToStr(TotalNb_j) + ' Jours');
  ExcelMergeRange(FExcelSheet, DerRow - 2, Col + 12, DerRow - 2, Col + 14);
  ExcelFontRange(FExcelSheet, DerRow - 2, Col + 12, DerRow - 2, Col + 14, FontEntete.Font);
  ExcelColorRange(FExcelSheet, DerRow - 2, Col + 12, DerRow - 2, Col + 14, 12615680, True);
  GblTotalNb_j := GblTotalNb_j + TotalNb_j;
  { COMPTEUR ABS SANS SOLDE PIED DE GRILLE}
  GblTotalAss := GblTotalAss + GblMensAss;

  { FORMAT ENTETE ANNEE}
  FontEntete.Font.Color := clWhite;
  ExcelMergeRange(FExcelSheet, Row, Col, Row, Col + 5);
  ExcelFontRange(FExcelSheet, Row, Col, Row, Col + 5, FontEntete.Font);
  ExcelColorRange(FExcelSheet, Row, Col, Row, Col + 5, 12615680, True);
  { FORMAT ENTETE JOURS}
  iColonne := Col - 1;
  for i := LOW(JOURS) to HIGH(JOURS) do
  begin
    Inc(iColonne, 2);
    ExcelMergeRange(FExcelSheet, Row + 2, Icolonne, Row + 2, iColonne + 1);
    ExcelFontRange(FExcelSheet, Row + 2, iColonne, Row + 2, iColonne + 1, FontEntete.Font);
    ExcelColorRange(FExcelSheet, Row + 2, iColonne, Row + 2, iColonne + 1, 12615680, True);
  end;

end;

procedure TOF_PGCALENDRIER_EXCEL.RecupAbsenceMoisSalarie(Mois: integer);
var
  T, T1, T_MotifAbs: Tob; { PT3 }
  Nb_J, Nb_H: double;
  DateDebutAbs, DateFinAbs, FinMois, DebMois, CalcDebAbs, CalcFinAbs: TDateTime;
  YYDeb, MMDEb, JJDeb, YYFin, MMFin, JJFin: Word;
  DebutDj, FinDj: integer;
begin
  if salarie = '' then exit;
  Debug('Recherche absence mois ' + IntToStr(Mois));
  if Tob_MoisAbsence <> nil then
  begin
    Tob_MoisAbsence.free;
    Tob_MoisAbsence := nil;
  end;
  T := Tob_Absence.FindFirst(['PCN_SALARIE'], [Salarie], False);
  while T <> nil do
  begin
    if Tob_MoisAbsence = nil then Tob_MoisAbsence := Tob.Create('LA maman', nil, -1);
    DateDebutAbs := T.GetValue('PCN_DATEDEBUTABS');
    DateFinAbs := T.GetValue('PCN_DATEFINABS');
    DeCodeDate(DateDebutAbs, YYDeb, MMDeb, JJDeb);
    DeCodeDate(DateFinAbs, YYFin, MMFin, JJFin);
    FinMois := FinDeMois(EncodeDate(YYDeb, MMdeb, 1));
    DebMois := EncodeDate(YYFin, MMFin, 1);
    if T.GetValue('PCN_DEBUTDJ') = 'MAT' then DebutDj := 0 else DebutDj := 1;
    if T.GetValue('PCN_FINDJ') = 'MAT' then FinDj := 0 else FinDj := 1;
    CalcDebAbs := idate1900;
    CalcFinAbs := idate1900;
    //Cas où debut et fin absence dans le même mois
    if (MMDeb = Mois) and (MMFin = Mois) then
    begin
      T1 := Tob.Create('LA fille', Tob_MoisAbsence, -1);
      T1.Dupliquer(T, True, True);
      Debug('Absence mensuel : Mois ' + IntToStr(Mois) + ' ' + T.GetValue('PCN_TYPECONGE') + ' ' + DatetoStr(DateDebutAbs) + ' ' + DatetoStr(DateFinAbs) + ' ' +
        FloattoStr(T.GetValue('PCN_JOURS')));
    end
    else
      //Cas où debut et fin absence pas dans le même mois mais même année
    begin //B Else
      if (YYDeb = StrToInt(Annee.text)) and (YYFin = StrToInt(Annee.text)) and ((MMDeb = Mois) or (MMFin = Mois)) then
      begin
        if MMDeb = Mois then
        begin
          CalcDebAbs := T.GetValue('PCN_DATEDEBUTABS');
          CalcFinAbs := FinMois;
          Findj := 1;
        end
        else
          if MMFin = Mois then
        begin
          CalcDebAbs := DebMois;
          CalcFinAbs := T.GetValue('PCN_DATEFINABS');
          DebutDj := 0;
        end;
      end
        //Cas où debut et fin absence pas dans le même mois ni dans la même année
      else
        if ((YYDeb = StrToInt(Annee.text)) or (YYFin = StrToInt(Annee.text))) and ((MMDeb = Mois) or (MMFin = Mois)) then
      begin
        if (YYDeb = StrToInt(Annee.text)) and (MMDeb = Mois) then
        begin
          CalcDebAbs := T.GetValue('PCN_DATEDEBUTABS');
          CalcFinAbs := FinMois;
          Findj := 1;
        end
        else
          if (YYFin = StrToInt(Annee.text)) and (MMFin = Mois) then
        begin
          CalcDebAbs := DebMois;
          CalcFinAbs := T.GetValue('PCN_DATEFINABS');
          DebutDj := 0;
        end;
      end;
      if (CalcDebAbs <> idate1900) and (CalcFinAbs <> idate1900) then
      begin
        { DEB PT3 }
        //CalculDuree(CalcDebAbs, CalcFinAbs, Salarie, T.GetValue('PCN_ETABLISSEMENT'),
        //T.GetValue('PCN_TYPECONGE'), Nb_J, Nb_H, DebutDj, 1);
        T_MotifAbs := Tob_MotifAbs.FindFirst(['PMA_MOTIFABSENCE'], [T.GetValue('PCN_TYPECONGE')], False);
        CalculNbJourAbsence(CalcDebAbs, CalcFinAbs, Salarie,T.GetValue('PCN_ETABLISSEMENT'),
        T.GetValue('PCN_TYPECONGE'), T_MotifAbs, Nb_J, Nb_H, DebutDj, 1);
        { FIN PT3 }
        T1 := Tob.Create('LE mouvement Eclate', Tob_MoisAbsence, -1);
        T1.Dupliquer(T, True, True);
        T1.PutValue('PCN_DATEFINABS', CalcDebAbs);
        T1.PutValue('PCN_DATEFINABS', CalcFinAbs);
        if DebutDj = 0 then T1.PutValue('PCN_DEBUTDJ', 'MAT') else T1.PutValue('PCN_DEBUTDJ', 'PAM');
        if FinDj = 0 then T1.PutValue('PCN_FINDJ', 'MAT') else T1.PutValue('PCN_FINDJ', 'PAM');
        T1.PutValue('PCN_JOURS', Nb_j);
        T1.PutValue('PCN_HEURES', Nb_h);
        Debug('Absence éclaté : Mois ' + IntToStr(Mois) + ' ' + T1.GetValue('PCN_TYPECONGE') + ' ' + DatetoStr(CalcDebAbs) + ' ' + DatetoStr(CalcFinAbs) + ' ' + FloattoStr(Nb_j));
      end;
    end;
    T := Tob_Absence.FindNext(['PCN_SALARIE'], [Salarie], False);
  end;

  if Tob_MoisAbsence <> nil then
  begin
    GblMensCpPris := Tob_MoisAbsence.Somme('PCN_JOURS', ['PCN_TYPECONGE'], ['PRI'], False);
    GblMensRtt := Tob_MoisAbsence.Somme('PCN_JOURS', ['PCN_TYPECONGE'], ['RTT'], False);
    GblMensRts := Tob_MoisAbsence.Somme('PCN_JOURS', ['PCN_TYPECONGE'], ['RTS'], False);
    GblMensAss := Tob_MoisAbsence.Somme('PCN_JOURS', ['PCN_TYPECONGE'], ['Ass'], False);
  end;
  Debug('GblMensCpPris : ' + FloattoStr(GblMensCpPris));
end;

function TOF_PGCALENDRIER_EXCEL.IsJourAbsence(DateJour: TdateTime; var Lib: string): Double;
var
  NbJ, duree: double;
  DateDebutAbs, DateFinAbs: TDateTime;
  DebutDj, FinDj, Typ: string;
  i: integer;
begin
  result := 0;
  lib := '';
  if Tob_MoisAbsence = nil then exit;
  for i := 0 to Tob_MoisAbsence.Detail.Count - 1 do
  begin
    duree := 0;
    DateDebutAbs := Tob_MoisAbsence.Detail[i].GetValue('PCN_DATEDEBUTABS');
    DateFinAbs := Tob_MoisAbsence.Detail[i].GetValue('PCN_DATEFINABS');
    DebutDj := Tob_MoisAbsence.Detail[i].GetValue('PCN_DEBUTDJ');
    FinDj := Tob_MoisAbsence.Detail[i].GetValue('PCN_FINDJ');
    NbJ := Tob_MoisAbsence.Detail[i].GetValue('PCN_JOURS');
    //Absence journalière
    if (DateJour = DateDebutAbs) and (DateJour = DateFinAbs) then
    begin
      Duree := Nbj;
    end
    else
      //Absence ds période d'absence
      if (DateJour >= DateDebutAbs) and (DateJour <= DateFinAbs) then
    begin
      if (DateJour = DateDebutAbs) and (DebutDj = 'PAM') then duree := 0.5
      else
        if (DateJour = DateDebutAbs) and (DebutDj = 'MAT') then duree := 1
      else
        if (DateJour = DateFinAbs) and (FinDj = 'MAT') then duree := 0.5
      else
        if (DateJour = DatefinAbs) and (FinDj = 'PAM') then duree := 1
      else duree := 1;
    end;
    if duree > 0 then
      if Tob_MoisAbsence.Detail[i].GetValue('PCN_TYPECONGE') = 'PRI' then Typ := 'P'
      else
        if Tob_MoisAbsence.Detail[i].GetValue('PCN_TYPECONGE') = 'Ass' then Typ := 'S'
      else Typ := 'R';
    if Nbj < duree then duree := nbj;
    Result := Result + duree;
    if result > 0 then lib := FloatToStr(Result) + Typ;
    Tob_MoisAbsence.Detail[i].PutValue('PCN_JOURS', Nbj - Duree);
  end;
end;

procedure TOF_PGCALENDRIER_EXCEL.InitialiseTob_Absence(var tob_Abs: Tob);
begin
  tob_Abs := Tob.Create('LA fille', Tob_Absence, -1);
  tob_Abs.AddChampSupValeur('PCN_SALARIE', Salarie);
  tob_Abs.AddChampSupValeur('PCN_ETABLISSEMENT', '');
  tob_Abs.AddChampSupValeur('PCN_TYPECONGE', 'PRI');
  tob_Abs.AddChampSupValeur('PCN_DATEDEBUTABS', idate1900);
  tob_Abs.AddChampSupValeur('PCN_DATEFINABS', idate1900);
  tob_Abs.AddChampSupValeur('PCN_JOURS', 0);
  tob_Abs.AddChampSupValeur('PCN_HEURES', 0);
  tob_Abs.AddChampSupValeur('PCN_DEBUTDJ', 'MAT');
  tob_Abs.AddChampSupValeur('PCN_FINDJ', 'PAM');
end;


procedure TOF_PGCALENDRIER_EXCEL.IntegreAbsenceHistobulletin;
var
  DateDebAbs, DateFinAbs, DateDebut, DateFin: TDateTime;
  Jours, Nb_j, nb_h: double;
  Libelle: string;
  Tob_HistoLibelle, T, T1, T2, T_MotifAbs: Tob;  { PT3 }
  YYdeb, YYFin, MM, JJ: WORD;
begin
  if Tob_Histo = nil then exit;
  Tob_HistoLibelle := nil;
  Debug('Intégration des absences de l''historique du salarié : ' + Salarie);
  T := Tob_Histo.FindFirst(['PHB_SALARIE'], [Salarie], False);
  while T <> nil do
  begin
    if Tob_HistoLibelle = nil then Tob_HistoLibelle := Tob.create('LES LIBELLES ABS', nil, -1);
    if Pos('.', T.GetValue('PHB_RUBRIQUE')) > 0 then
      T.ChangeParent(Tob_HistoLibelle, -1);
    Debug('Duplication des lignes de commentaire');
    T := Tob_Histo.FindNext(['PHB_SALARIE'], [Salarie], False);
  end;

  T := Tob_Histo.FindFirst(['PHB_SALARIE'], [Salarie], False);
  while T <> nil do
  begin
    DateDebut := T.GetValue('PHB_DATEDEBUT');
    DaTeFin := T.GetValue('PHB_DATEFIN');
    Jours := T.GetValue('PHB_BASEREM');
    //Recherche du libellé associé
    T2 := Tob_HistoLibelle.FindFirst(['PHB_SALARIE', 'PHB_DATEDEBUT', 'PHB_DATEFIN'], [Salarie, datedebut, datefin], False);
    while T2 <> nil do
    begin
      DateDebAbs := idate1900;
      DateFinAbs := idate1900;
      Libelle := T2.GetValue('PHB_LIBELLE');
      if IsValidDate(Copy(Libelle, 11, 10)) then DateDebAbs := StrToDate(Copy(Libelle, 11, 10));
      if IsValidDate(Copy(Libelle, 25, 10)) then DateFinAbs := StrToDate(Copy(Libelle, 25, 10));
      DecodeDate(DateDebAbs, YYdeb, MM, JJ);
      DecodeDate(DateFinAbs, YYFin, MM, JJ);
      if (DateDebAbs <> idate1900) and (DateFinAbs <> idate1900) and ((YYdeb = StrToInt(Annee.Text)) or (YYFin = StrToInt(Annee.Text))) then
      begin
        InitialiseTob_Absence(T1);
        T1.PutValue('PCN_ETABLISSEMENT', T.GetValue('PHB_ETABLISSEMENT'));
        T1.PutValue('PCN_SALARIE', Salarie);
        T1.PutValue('PCN_TYPECONGE', 'PRI');
        T1.PutValue('PCN_DEBUTDJ', 'MAT');
        T1.PutValue('PCN_FINDJ', 'PAM');
        T1.PutValue('PCN_DATEDEBUTABS', DateDebAbs);
        T1.PutValue('PCN_DATEFINABS', DateFinAbs);
        { DEB PT3 }
        //CalculDuree(DateDebAbs, DateFinAbs, Salarie, T.GetValue('PHB_ETABLISSEMENT'), 'PRI',
        //Nb_j, nb_h, 0, 1);
        T_MotifAbs := Tob_MotifAbs.FindFirst(['PMA_MOTIFABSENCE'], ['PRI'], False);
        CalculNbJourAbsence(DateDebAbs, DateFinAbs, Salarie, T.GetValue('PHB_ETABLISSEMENT'),
        'PRI', T_MotifAbs, Nb_J, Nb_H, 0, 1);
        { FIN PT3 }
        Debug('Init mvt Date : ' + DateToStr(DateDebAbs) + ' jours : ' + FloatToStr(Nb_j));
        if Nb_j <= Jours then
        begin
          T1.PutValue('PCN_JOURS', Nb_j);
          Jours := Jours - nb_j;
        end
        else
        begin
          T1.PutValue('PCN_JOURS', Jours);
          Break;
        end;
      end;
      T2 := Tob_HistoLibelle.FindNext(['PHB_SALARIE', 'PHB_DATEDEBUT', 'PHB_DATEFIN'], [Salarie, datedebut, datefin], False);
    end;
    T := Tob_Histo.FindNext(['PHB_SALARIE'], [Salarie], False);
  end;
  if Tob_HistoLibelle <> nil then Tob_HistoLibelle.free;
  Debug('Fin Intégration abs historique du salarié : ' + Salarie);
end;

procedure TOF_PGCALENDRIER_EXCEL.MiseEnFormePiedDePage(Col, Row: integer);
var
  Periodei: integer;
  DDebut, Dfin: TDateTime;
  YY, MM, JJ: WORD;
  trouve: boolean;
  Pris, Acquis, Restants, Base, Moisbase, BaseForfait : double; { PT2 }
  St: string;
begin
  Pris := 0;
  Acquis := 0;
  Restants := 0;
  Base := 0;
  Moisbase := 0;
  Trouve := False;
  ExcelAlignRange(FExcelSheet, Col, Row, Col + 31, Row + 14, tacenter);
  ExcelColorRange(FExcelSheet, Row, Col, Row + 14, Col + 33, clWhite, False);
  {PIED DE PAGE DU CLASSEUR GRILLE : INITIALISATION}
  Periodei := CalculPeriode(DTcloture, EnCodeDate(StrToInt(Annee.text), 5, 31));
  while Trouve = False do
  begin
    RechercheDate(DDebut, Dfin, DtCloture, IntToStr(Periodei));
    Decodedate(Dfin, YY, MM, JJ);
    if YY = StrToInt(Annee.text) then trouve := True;
    if (Periodei > 10) or (trouve) then break;
    Inc(Periodei);
  end;
  if Trouve then
    AffichelibelleAcqPri(IntToStr(Periodei), Salarie, idate1900, DFin, Pris, Acquis, Restants, Base, Moisbase, True, False);
  {Identification salarié}
  MyExcelText(FExcelSheet, Row, Col + 1, 'Salarié ' + Salarie + ' ' + NomPrenom);
  ExcelAlignRange(FExcelSheet, Col + 1, Row, Col + 1, Row, taLeftJustify);
  {Identification responsable}
  MyExcelText(FExcelSheet, Row + 1, Col + 1, 'Responsable ' + responsable);
  ExcelAlignRange(FExcelSheet, Col + 1, Row + 1, Col + 1, Row + 1, taLeftJustify);
  {Date Entrée et Date de sortie}
  MyExcelText(FExcelSheet, Row + 3, Col + 1, 'Date d''entrée');
  ExcelMergeRange(FExcelSheet, Row + 3, Col + 1, Row + 4, Col + 4);
  MyExcelText(FExcelSheet, Row + 6, Col + 1, 'Date de sortie');
  ExcelMergeRange(FExcelSheet, Row + 6, Col + 1, Row + 7, Col + 4);

  MyExcelText(FExcelSheet, Row + 3, Col + 5, UsDateTime(DateEntree));
  ExcelMergeRange(FExcelSheet, Row + 3, Col + 5, Row + 4, Col + 9);
  ExcelColorRange(FExcelSheet, Row + 3, Col + 5, Row + 4, Col + 9, 12615680, False);
  ExcelFontRange(FExcelSheet, Row + 3, Col + 5, Row + 4, Col + 9, FontEntete.Font);

  if DateSortie > idate1900 then
    MyExcelText(FExcelSheet, Row + 6, Col + 5, UsDateTime(DateSortie));
  ExcelMergeRange(FExcelSheet, Row + 6, Col + 5, Row + 7, Col + 9);
  ExcelColorRange(FExcelSheet, Row + 6, Col + 5, Row + 7, Col + 9, 12615680, False);
  ExcelFontRange(FExcelSheet, Row + 6, Col + 5, Row + 7, Col + 9, FontEntete.Font);

  {Jours travaillés}
  MyExcelText(FExcelSheet, Row + 9, Col + 1, FloatToStr(GblTotalNb_j) + ' jrs travaillés');
  ExcelMergeRange(FExcelSheet, Row + 9, Col + 1, Row + 9, Col + 5);
  {Jours ouvrés}
  MyExcelText(FExcelSheet, Row + 11, Col + 1, FloatToStr(GblTotalNb_j + GblTotalCpPris + GblTotalRtt + GblTotalAss) + ' jrs ouvrés');
  ExcelMergeRange(FExcelSheet, Row + 11, Col + 1, Row + 11, Col + 5);
  {Base forfaitaire}
  if Jouvres <> 0 then
    Begin
    BaseForfait := StrToFloat(GetControlText('BASEFORFAIT')); { PT2 }
    MyExcelText(FExcelSheet, Row + 10, Col + 8, FloatToStr(Arrondi( BaseForfait * (GblTotalNb_j + GblTotalCpPris + GblTotalRtt + GblTotalAss) / Jouvres, 0)) + ' base forfait');
    End;
  ExcelMergeRange(FExcelSheet, Row + 10, Col + 8, Row + 10, Col + 13);
  {Légende}
  MyExcelText(FExcelSheet, Row + 13, Col + 1, 'Légende : P = Congés Payés    R = Réduction Temps Travail    S = Absence Sans Solde');
  ExcelAlignRange(FExcelSheet, Col + 1, Row + 13, Col + 1, Row + 13, taLeftJustify);
  FontEntete.Font.Color := clBlack;
  FontEntete.Font.Style := [];
  ExcelFontRange(FExcelSheet, Row + 13, Col + 1, Row + 13, Col + 1, FontEntete.Font);

  {Visa Collaborateur}
  MyExcelText(FExcelSheet, Row + 11, Col + 20, 'Visa Collaborateur');
  ExcelAlignRange(FExcelSheet, Col + 20, Row + 11, Col + 20, Row + 11, taLeftJustify);
  ExcelMergeRange(FExcelSheet, Row + 10, Col + 19, Row + 12, Col + 25);
  ExcelColorRange(FExcelSheet, Row + 10, Col + 19, Row + 12, Col + 25, ClWhite, True);
  ExcelMergeRange(FExcelSheet, Row + 10, Col + 26, Row + 12, Col + 31);
  ExcelColorRange(FExcelSheet, Row + 10, Col + 26, Row + 12, Col + 31, ClWhite, True);
  {Jours Rtt}
  MyExcelText(FExcelSheet, Row + 3, Col + 11, 'JRTT Pris');
  ExcelMergeRange(FExcelSheet, Row + 3, Col + 11, Row + 3, Col + 15);
  ExcelAlignRange(FExcelSheet, Col + 11, Row + 3, Col + 15, Row + 3, taLeftJustify);
  ExcelFontRange(FExcelSheet, Row + 3, Col + 11, Row + 3, Col + 15, FontEntete.Font);
  MyExcelText(FExcelSheet, Row + 4, Col + 11, 'sur Année ' + Annee.text);
  ExcelMergeRange(FExcelSheet, Row + 4, Col + 11, Row + 4, Col + 15);
  ExcelAlignRange(FExcelSheet, Col + 11, Row + 4, Col + 15, Row + 4, taLeftJustify);
  ExcelFontRange(FExcelSheet, Row + 4, Col + 11, Row + 4, Col + 15, FontEntete.Font);
  //Abs sans solde
  MyExcelText(FExcelSheet, Row + 6, Col + 11, 'Absence sans solde');
  ExcelMergeRange(FExcelSheet, Row + 6, Col + 11, Row + 6, Col + 15);
  ExcelAlignRange(FExcelSheet, Col + 11, Row + 6, Col + 11, Row + 6, taLeftJustify);
  ExcelFontRange(FExcelSheet, Row + 6, Col + 11, Row + 6, Col + 15, FontEntete.Font);
  MyExcelText(FExcelSheet, Row + 7, Col + 11, 'sur Année ' + Annee.text);
  ExcelMergeRange(FExcelSheet, Row + 7, Col + 11, Row + 7, Col + 15);
  ExcelAlignRange(FExcelSheet, Col + 11, Row + 7, Col + 11, Row + 7, taLeftJustify);
  ExcelFontRange(FExcelSheet, Row + 7, Col + 11, Row + 7, Col + 15, FontEntete.Font);

  {Cp}
  MyExcelText(FExcelSheet, Row, Col + 20, 'Cp Acquis du ' + DateToStr(DDebut));
  ExcelAlignRange(FExcelSheet, Col + 20, Row, Col + 20, Row, taLeftJustify);
  ExcelFontRange(FExcelSheet, Row, Col + 20, Row, Col + 20, FontEntete.Font);

  MyExcelText(FExcelSheet, Row + 1, Col + 20, 'Au ' + DateToStr(DFin));
  ExcelAlignRange(FExcelSheet, Col + 20, Row + 1, Col + 20, Row + 1, taLeftJustify);
  ExcelFontRange(FExcelSheet, Row + 1, Col + 20, Row + 1, Col + 20, FontEntete.Font);

  MyExcelText(FExcelSheet, Row + 3, Col + 20, 'CP Pris du 01/01/' + Annee.text);
  ExcelAlignRange(FExcelSheet, Col + 20, Row + 3, Col + 20, Row + 3, taLeftJustify);
  ExcelFontRange(FExcelSheet, Row + 3, Col + 20, Row + 3, Col + 20, FontEntete.Font);
  if (DateSortie > idate1900) and (DateSortie <= StrToDate('31/12/' + Annee.text)) then
    st := 'Au ' + DateToStr(DateSortie) else st := 'Au 31/12/' + Annee.text;
  MyExcelText(FExcelSheet, Row + 4, Col + 20, St);
  ExcelAlignRange(FExcelSheet, Col + 20, Row + 4, Col + 20, Row + 4, taLeftJustify);
  ExcelFontRange(FExcelSheet, Row + 4, Col + 20, Row + 4, Col + 20, FontEntete.Font);

  MyExcelText(FExcelSheet, Row + 6, Col + 20, 'Solde restant à déduire');
  ExcelMergeRange(FExcelSheet, Row + 6, Col + 20, Row + 6, Col + 24);
  ExcelAlignRange(FExcelSheet, Col + 20, Row + 6, Col + 20, Row + 6, taLeftJustify);
  ExcelFontRange(FExcelSheet, Row + 6, Col + 20, Row + 6, Col + 24, FontEntete.Font);
  MyExcelText(FExcelSheet, Row + 7, Col + 20, 'Sur les jours travaillés');
  ExcelMergeRange(FExcelSheet, Row + 7, Col + 20, Row + 7, Col + 24);
  ExcelAlignRange(FExcelSheet, Col + 20, Row + 7, Col + 20, Row + 7, taLeftJustify);
  ExcelFontRange(FExcelSheet, Row + 7, Col + 20, Row + 7, Col + 24, FontEntete.Font);

  FontEntete.Font.Style := [FsBold];
  FontEntete.Font.Color := clWhite;
  //Pris Rtt
  MyExcelText(FExcelSheet, Row + 3, Col + 16, FloatToStr(Arrondi(GblTotalRtt, 2)) + ' jours');
  ExcelMergeRange(FExcelSheet, Row + 3, Col + 16, Row + 4, Col + 18);
  {Abs sans solde}
  MyExcelText(FExcelSheet, Row + 7, Col + 16, FloatToStr(Arrondi(GblTotalAss, 2)) + ' jours');
  ExcelMergeRange(FExcelSheet, Row + 6, Col + 16, Row + 7, Col + 18);
  //Acquis CP
  MyExcelText(FExcelSheet, Row, Col + 26, FloatToStr(Arrondi(Acquis, 2)) + ' jours');
  ExcelMergeRange(FExcelSheet, Row, Col + 26, Row + 1, Col + 28);
  ExcelColorRange(FExcelSheet, Row, Col + 26, Row + 1, Col + 28, 12615680, False);
  ExcelFontRange(FExcelSheet, Row, Col + 26, Row + 1, Col + 28, FontEntete.Font);
  //Pris CP
  MyExcelText(FExcelSheet, Row + 3, Col + 26, FloatToStr(Arrondi(GblTotalCpPris, 2)) + ' jours');
  ExcelMergeRange(FExcelSheet, Row + 3, Col + 26, Row + 4, Col + 28);
  //Solde CP
  MyExcelText(FExcelSheet, Row + 7, Col + 26, FloatToStr(Arrondi(Acquis - GblTotalCpPris, 2)) + ' jours');
  ExcelMergeRange(FExcelSheet, Row + 6, Col + 26, Row + 7, Col + 28);
  ExcelColorRange(FExcelSheet, Row + 6, Col + 26, Row + 7, Col + 28, 12615680, False);
  ExcelFontRange(FExcelSheet, Row + 6, Col + 26, Row + 7, Col + 28, FontEntete.Font);
end;


procedure TOF_PGCALENDRIER_EXCEL.ChargeTob(var Tob_Salarie: Tob);
var
  Q: TQuery;
  st, StSal, StWhere, StPer: string;
  Edit: THEdit;
  Check: TCheckBox;
begin
  StWhere := '';
  Edit := THEdit(GetControl('PSA_SALARIE'));
  if Edit <> nil then if edit.text <> '' then StWhere := 'AND PSA_SALARIE>="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_SALARIE_'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_SALARIE<="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_ETABLISSEMENT'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_ETABLISSEMENT>="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_ETABLISSEMENT_'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_ETABLISSEMENT<="' + Edit.Text + '"';

  Edit := THEdit(GetControl('PSA_TRAVAILN1'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_TRAVAILN1>="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_TRAVAILN1_'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_TRAVAILN1<="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_TRAVAILN2'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_TRAVAILN2>="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_TRAVAILN2_'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_TRAVAILN2<="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_TRAVAILN3'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_TRAVAILN3>="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_TRAVAILN3_'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_TRAVAILN3<="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_TRAVAILN4'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_TRAVAILN4>="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_TRAVAILN4_'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_TRAVAILN4<="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_CODESTAT'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_CODESTAT>="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_CODESTAT_'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_CODESTAT<="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_LIBREPCMB1'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_LIBREPCMB1>="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_LIBREPCMB1_'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_LIBREPCMB1<="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_LIBREPCMB2'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_LIBREPCMB2>="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_LIBREPCMB2_'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_LIBREPCMB2<="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_LIBREPCMB3'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_LIBREPCMB3>="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_LIBREPCMB3_'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_LIBREPCMB3<="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_LIBREPCMB4'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_LIBREPCMB4>="' + Edit.Text + '"';
  Edit := THEdit(GetControl('PSA_LIBREPCMB4_'));
  if Edit <> nil then if edit.text <> '' then StWhere := StWhere + ' AND PSA_LIBREPCMB4<="' + Edit.Text + '"';

  Check := TCheckBox(GetControl('PSA_BOOLLIBRE1'));
  if Check <> nil then
    if Check.state = CbChecked then StWhere := StWhere + ' AND PSA_BOOLLIBRE1="X"'
    else if Check.state = CbUnChecked then StWhere := StWhere + ' AND PSA_BOOLLIBRE1="-"';


  StSal := 'SELECT PSA_SALARIE FROM SALARIES ' +
    'WHERE PSA_SALARIE<>"" ' + StWhere;
  Q := OpenSql('SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_DATEENTREE,PSA_DATESORTIE,' +
    'PSA_TRAVAILN4,ETB_LIBELLE,ETB_DATECLOTURECPN FROM SALARIES ' +
    'LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' +
    'WHERE PSA_SALARIE<>"" ' + StWhere +
    ' AND (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE="' + UsDateTime(idate1900) + '" OR PSA_DATESORTIE>="' + UsDateTime(EncodeDate(StrToInt(Annee.text), 1, 1)) + '") ' +
    'ORDER BY PSA_ETABLISSEMENT,PSA_SALARIE', True);
  Tob_Salarie := Tob.Create('Liste des salariés', nil, -1);
  Tob_Salarie.LoadDetailDB('SAL_ARIE', '', '', Q, False);
  Ferme(Q);

  if Tob_Salarie.Detail.count > 0 then
  begin
    //Init des jours fériés
    Q := OpenSql('SELECT * FROM JOURFERIE ' +
      'WHERE ##AJF_PREDEFINI## (AJF_JOURFIXE="X" OR AJF_ANNEE=' + Annee.text + ')', True); // DB2 { PT1 }
    Tob_JourFerie := Tob.create('Les jours feriés', nil, -1);
    Tob_JourFerie.LoadDetailDB('JOURFERIE', '', '', Q, False);
    Ferme(Q);
    if VH_Paie.PGRecupAbsHisto then
    begin
      if StrToInt(Annee.text) <= 2001 then
        StPer := 'AND (PCN_DATEDEBUTABS>"05/31/2001" OR PCN_DATEFINABS>"05/31/2001") '
      else StPer := '';
    end
    else StPer := '';
    //Init des CP  et des RTT
    St := 'SELECT PCN_SALARIE,PCN_ETABLISSEMENT,PCN_TYPECONGE,PCN_DATEDEBUTABS,PCN_DATEFINABS,' +
      'PCN_JOURS,PCN_HEURES,PCN_DEBUTDJ,PCN_FINDJ ' +
      'FROM ABSENCESALARIE ' +
      'WHERE PCN_SALARIE IN (' + StSal + ') ' +
      'AND ((PCN_TYPECONGE="PRI" ' + StPer + ') OR PCN_TYPECONGE="RTT" OR PCN_TYPECONGE="RTS" OR PCN_TYPECONGE="Ass") ' +
      'AND PCN_ETATPOSTPAIE <> "NAN" '+ { PT4 } 
      'AND (YEAR(PCN_DATEDEBUTABS)=' + Annee.text + ' or YEAR(PCN_DATEFINABS)=' + Annee.text + ') AND PCN_MVTDUPLIQUE="-" ' +
      'ORDER BY PCN_SALARIE,PCN_DATEDEBUTABS,PCN_DATEFINABS';
    Q := OpenSql(St, True);
    Tob_absence := Tob.create('Les Absences', nil, -1);
    Tob_absence.LoadDetailDB('ABS_SALARIE', '', '', Q, False);
    Ferme(Q);

    //Absence histobulletin spécifique Cegid
    if VH_Paie.PGRecupAbsHisto then
      if StrToInt(Annee.text) <= 2001 then
      begin
        St := 'SELECT PHB_SALARIE,PHB_ETABLISSEMENT,PHB_DATEDEBUT,PHB_DATEFIN,PHB_LIBELLE, ' +
          'PHB_RUBRIQUE,PHB_BASEREM  FROM HISTOBULLETIN ' +
          'WHERE PHB_SALARIE IN (' + StSal + ') ' +
          'AND PHB_NATURERUB="AAA" ' +
          'AND ((PHB_RUBRIQUE="3210" AND PHB_BASEREM<>0) OR (PHB_RUBRIQUE LIKE "3210.%")) ' +
          'AND YEAR(PHB_DATEFIN)=' + Annee.text +
          ' AND PHB_DATEFIN>="01/01/' + Annee.text + '" and PHB_DATEFIN<="05/31/2001" ' +
          'ORDER BY PHB_SALARIE,PHB_DATEDEBUT,PHB_DATEFIN';
        Q := OpenSql(St, True);
        Tob_Histo := Tob.create('Les abs histo', nil, -1);
        Tob_Histo.LoadDetailDB('ABS_HISTO', '', '', Q, False);
        Ferme(Q);
      end;
  end;

end;

procedure TOF_PGCALENDRIER_EXCEL.LibereTob;
begin
  if Tob_Semaine <> nil then
  begin
    Tob_Semaine.free;
    Tob_Semaine := nil;
  end;
  if Tob_Standard <> nil then
  begin
    Tob_Standard.free;
    Tob_Standard := nil;
  end;
  if Tob_MoisAbsence <> nil then
  begin
    Tob_MoisAbsence.free;
    Tob_MoisAbsence := nil;
  end;
end;



initialization
  registerclasses([TOF_PGCALENDRIER_EXCEL]);
end.

