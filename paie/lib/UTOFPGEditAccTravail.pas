{***********UNITE*************************************************
Auteur  ...... : PaiePgi
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Unit gestion attestation Accident du Travail
Mots clefs ... : PAIE;ATTESTATIONS
*****************************************************************}
{
PT1   : 03/12/2001 SB V563 Clé n°SS erronnée
PT2   : 22/04/2002 SB V571 Fiche de bug n°10093 : Code Risque et Medecine du
                           travail
PT3   : 06/06/2002 SB V582 Fiche de bug n°10144 : Confusion variable requête
PT4   : 27/09/2002 VG V585 1 requête non fermée si nouvelle attestation
PT5-1 : 11/12/2002 SB V591 FQ 10311 Reprise des dates d'absence
PT5-2 : 11/12/2002 SB V591 FQ 10230 Calcul salaire à partir du dernier jour travaillé
PT5-3 : 11/12/2002 SB V591 FQ 10325 Calcul Colonne7 part salarial
PT5-4 : 22/01/2003 SB V591 FQ 10230 Remise à blanc colonne 7 non effectué
PT6   : 19/03/2003 SB V595 FQ 10564 Reprise du salaire brut abbatu et non du brut
PT7   : 01/07/2003 PH V_42 Portage CWAS
PT8   : 16/07/2003 SB V_42 FQ 10512 Orthographe (ans)
PT9   : 15/09/2003 SB V_42 FQ 10786 Erreur utilsation champ BanqueCP
PT10  : 09/10/2003 SB V_42 Refonte utilisation d'une tob pour maj table
PT11  : 27/10/2003 SB V_42 FQ 10914 Zone mal initialisé
PT12  : 29/10/2003 SB V_42 Dysfonctionnement MAJ DONNEE en modification
PT13-1: 28/04/2004 SB V_50 FQ 10812 Intégration de la gestion des déclarants
PT13-2: 28/04/2004 SB V_50 FQ 11225 Gestion du mode création
PT14  : 04/05/2004 SB V_50 FQ 11138 Reprise du code risque en cours de validité
PT15  : 05/05/2004 SB V_50 FQ 11220 Ajout du nom de jeune fille
PT16  : 09/06/2004 MF V_50 Le paramétre TypeAbs permet de cocher la case Accident
                           du travail ou maladie professionnelle - Qd il y a subrogation (champ
                          ETB_SUBROGATION) contrôle des champ correspondant.
PT17  : 27/10/2004 PH V_60 Format date ORACLE
PT18  : 15/03/2005 SB V_60 FQ 12057 Recherche code risque erroné
PT19  : 19/09/2005 SB V_65 FQ 12576 Anomalie remplissage tableau en CWAS
PT20  : 19/09/2005 SB V_65 FQ 12326 Refonte calcul colonne 7
---- JL 20/03/2006 modification clé annuaire ----
PT21  : 20/10/2006 SB V_70 FQ 13504 Anomalie conversion double
PT22  : 09/11/2006 SB V_70 Rep FQ 13504 Anomalie maj zone numerique
PT23  : 14/03/2007 VG V_72 BQ_GENERAL n'est pas forcément unique
}
unit UTOFPGEditAccTravail;

interface
uses StdCtrls, Controls, Classes,  sysutils, ComCtrls, UTOF, UTOB,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  EdtREtat,
  {$ELSE}
  UtileAgl,
  {$ENDIF}
  Hctrls, PgEditOutils, HTB97, HPdfviewer, ParamDat, HEnt1, ParamSoc, HMsgBox,
  PgOutilsTreso, Commun;

type
  TOF_PGACCTRAVAIL = class(TOF)
  private
    Matricule, Mode, GblEtab, TypeAbs: string; // PT16
    subrogation: boolean; // PT16
    procedure InitSalarie(Arguments: string);
    procedure InitAccTravail(Mode: string); { PT13-2 }
    procedure UpdateTable(Sender: TObject);
    procedure AffectSalaire(Sender: TObject);
    procedure DesactiveNat(Sender: TObject);
    procedure DesactiveVict(Sender: TObject);
    procedure DesactiveMaintien(Sender: TObject);
    procedure GereMaintien(Sender: TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure AfficheCheck(Sender: TObject);
    procedure GestionSubrogation(Sender: TObject);
    procedure Imprimer(Sender: TObject);
    procedure AffichDeclarant(Sender: TObject); { PT13-1 }
    procedure NewAttest(Sender: TObject); { PT13-2 }
    procedure AffectCodeRisque(Sender: TObject); { PT14 }
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  end;

implementation

uses P5Util, EntPaie;

{ TOF_PGACCTRAVAIL }

procedure TOF_PGACCTRAVAIL.OnArgument(Arguments: string);
var
  btn: TToolBarButton97;
  Edit: THEdit;
  Lb: THLabel;
  Check: TCheckBox;
  i: integer;
 // PDFAccTravail: TPDFPreview;
  NumEdit: THNumEdit;
  DateAbsence: TDateTime;
  St: string;
begin
  inherited;
  Matricule := ReadTokenSt(Arguments);
  Mode := ReadTokenSt(Arguments);
  TypeAbs := ReadTokenSt(Arguments); // PT16

  //DEB PT5-1
  Arguments := Trim(Arguments);
  if IsValidDate(Arguments) then DateAbsence := StrToDate(Arguments) else DateAbsence := idate1900;
  if DateAbsence > idate1900 then
    SetControlText('DATEACCTRAV', DateToStr(DateAbsence));
  //FIN PT5-1
{
  PDFAccTravail := TPDFPreview(GetControl('PDFACCTRAVAIL'));
  if PDFAccTravail <> nil then
  begin
    PDFAccTravail.PDFPath := VH_Paie.PGCheminRech + '\ACCTRAVAIL.pdf'; //Affectation du fichier à recup
  end;
}
  { DEB PT13-2 }
  if not ExisteSql('SELECT PAS_ORDRE FROM ATTESTATIONS WHERE PAS_SALARIE="' + Matricule + '" AND PAS_TYPEATTEST="ACC"') then
    Mode := 'CREATION';
  GblEtab := RechDom('PGSALARIEETAB', Matricule, False);
  { FIN PT13-2 }
  { DEB PT13-1 }
  St := '(PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%' + GblEtab + '%") ' +
    ' AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%ACT%" )  ';
  SetControlProperty('DECLARANT', 'Plus', St);
  { FIN PT13-1 }
  if Matricule <> '' then //PT5-1
  begin
    InitSalarie(Matricule);
    InitAccTravail(Mode); { PT13-2 }
  end;
  btn := TToolBarButton97(GetControl('BValider'));
  if btn <> nil then btn.OnClick := UpdateTable;
  btn := TToolBarButton97(GetControl('BImprimer'));
  if btn <> nil then btn.OnClick := Imprimer;
  btn := TToolBarButton97(GetControl('BInsert')); { PT13-2 }
  if btn <> nil then btn.OnClick := NewAttest;
  //Evenements
  Edit := THEdit(GetControl('DATEACCTRAV'));
  if Edit <> nil then
  begin
    Edit.OnDblClick := DateElipsisclick;
    Edit.OnExit := AffectCodeRisque;
  end; { PT14 }
  Edit := THEdit(GetControl('DATEDERTRAV')); //PT5-2
  if Edit <> nil then
  begin
    Edit.OnDblClick := DateElipsisclick;
    Edit.OnExit := AffectSalaire;
  end;
  Edit := THEdit(GetControl('DATEREPTRAV'));
  if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
  Edit := THEdit(GetControl('DATECONTRAT'));
  if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
  Check := TCheckBox(GetControl('CKNATFRANC'));
  if check <> nil then Check.OnClick := DesactiveNat;
  Check := TCheckBox(GetControl('CKNATCEE'));
  if check <> nil then Check.OnClick := DesactiveNat;
  Check := TCheckBox(GetControl('CKNATAUTRE'));
  if check <> nil then Check.OnClick := DesactiveNat;
  Check := TCheckBox(GetControl('CKVICTOUI'));
  if check <> nil then Check.OnClick := DesactiveVict;
  Check := TCheckBox(GetControl('CKVICTNON'));
  if check <> nil then Check.OnClick := DesactiveVict;
  Check := TCheckBox(GetControl('MAINTIENOUI'));
  if check <> nil then Check.OnClick := DesactiveMaintien;
  Check := TCheckBox(GetControl('MAINTIENNON'));
  if check <> nil then Check.OnClick := DesactiveMaintien;
  Check := TCheckBox(GetControl('MAINTIENINT_'));
  if check <> nil then Check.OnClick := GereMaintien;
  Check := TCheckBox(GetControl('MAINTIENPART_'));
  if check <> nil then Check.OnClick := GereMaintien;

  for i := 1 to 3 do
  begin
    Edit := THEdit(GetControl('DATEECHEANCE' + IntToStr(i)));
    if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
    Edit := THEdit(GetControl('DATEDEBPER' + IntToStr(i)));
    if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
    Edit := THEdit(GetControl('DATEFINPER' + IntToStr(i)));
    if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
    Edit := THEdit(GetControl('DATEVER' + IntToStr(i)));
    if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
    Edit := THEdit(GetControl('DATEDEBVER' + IntToStr(i)));
    if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
    Edit := THEdit(GetControl('DATEFINVER' + IntToStr(i)));
    if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
    Edit := THEdit(GetControl('DATEDEBARRET' + IntToStr(i)));
    if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
    Edit := THEdit(GetControl('DATEFINARRET' + IntToStr(i)));
    if Edit <> nil then Edit.OnDblClick := DateElipsisclick;

    NumEdit := THNumEdit(GetControl('MTSAL' + IntToStr(i)));
    if NumEdit <> nil then NumEdit.OnExit := AfficheCheck;
    NumEdit := THNumEdit(GetControl('AVANT' + IntToStr(i)));
    if NumEdit <> nil then NumEdit.OnExit := AfficheCheck;
    NumEdit := THNumEdit(GetControl('INDEM' + IntToStr(i)));
    if NumEdit <> nil then NumEdit.OnExit := AfficheCheck;
    NumEdit := THNumEdit(GetControl('PARTSAL4' + IntToStr(i)));
    if NumEdit <> nil then NumEdit.OnExit := AfficheCheck;
    NumEdit := THNumEdit(GetControl('SOUMIS' + IntToStr(i)));
    if NumEdit <> nil then NumEdit.OnExit := AfficheCheck;
    NumEdit := THNumEdit(GetControl('BRUTVER' + IntToStr(i)));
    if NumEdit <> nil then NumEdit.OnExit := AfficheCheck;
    NumEdit := THNumEdit(GetControl('PARTSAL12' + IntToStr(i)));
    if NumEdit <> nil then NumEdit.OnExit := AfficheCheck;
    NumEdit := THNumEdit(GetControl('BRUTPERDU' + IntToStr(i)));
    if NumEdit <> nil then NumEdit.OnExit := AfficheCheck;
    NumEdit := THNumEdit(GetControl('PARTSAL18' + IntToStr(i)));
    if NumEdit <> nil then NumEdit.OnExit := AfficheCheck;
  end;

  Edit := THEdit(GetControl('DATEDEBSUBR'));
  if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
  Edit := THEdit(GetControl('DATEFINSUBR'));
  if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
  Edit := THEdit(GetControl('DENOMINATION'));
  if Edit <> nil then
  begin
    Edit.OnDblClick := DateElipsisclick;
    Edit.OnChange := GestionSubrogation;
  end;
  Edit := THEdit(GetControl('DATEFAIT'));
  if Edit <> nil then Edit.OnDblClick := DateElipsisclick;

  { DEB PT13-1 }
  Edit := THEdit(GetControl('DECLARANT'));
  if Edit <> nil then Edit.OnExit := AffichDeclarant;
  { FIN PT13-1 }

  //VALEUR PAR DEFAUT
  check := TCheckBox(GetControl('MAINTIENINT'));
  if check <> nil then Check.Enabled := False;
  check := TCheckBox(GetControl('MAINTIENPART'));
  if check <> nil then Check.Enabled := False;

  Lb := THLabel(GetControl('SOLIBELLE'));
  if Lb <> nil then Lb.Caption := GetParamSoc('SO_LIBELLE');
  Lb := THLabel(GetControl('SOADRESSE'));
  if Lb <> nil then Lb.Caption := GetParamSoc('SO_ADRESSE1') + '  ' +
    GetParamSoc('SO_ADRESSE2') + '   ' + GetParamSoc('SO_ADRESSE3');
  Lb := THLabel(GetControl('SOCPVILLE'));
  if Lb <> nil then Lb.Caption := GetParamSoc('SO_CODEPOSTAL') + '  ' + GetParamSoc('SO_VILLE');
  Lb := THLabel(GetControl('SOTELEPHONE'));
  if Lb <> nil then Lb.Caption := FormatCase(GetParamSoc('SO_TELEPHONE'), 'STR', 2);

SetPlusBanqueCP (GetControl ('DENOMINATION'));                  //PT23
end;

procedure TOF_PGACCTRAVAIL.OnLoad;
begin
  inherited;

end;


procedure TOF_PGACCTRAVAIL.InitSalarie(Arguments: string);
var
  WDate: Word;
  Q, QSal: TQuery;
  Defaut: THLabel;
  CheckFr, CheckCee, CheckAutre: TCheckBox;
  Etab, St: string;
begin
  subrogation := False;
  QSal := OpenSql('SELECT PSA_ETABLISSEMENT,PSA_SALARIE,PSA_LIBELLE,PSA_NOMJF,PSA_PRENOM,PSA_ADRESSE1,PSA_ADRESSE2,' + { PT15 }
    'PSA_ADRESSE3,PSA_CODEPOSTAL,PSA_VILLE,PSA_LIBELLEEMPLOI,PSA_NUMEROSS,' +
    'PSA_DATENAISSANCE,PSA_QUALIFICATION,PSA_DATEENTREE,PSA_DATEANCIENNETE,' +
    'PSA_NATIONALITE,PSA_SEXE, PSA_DATENAISSANCE,ET_ETABLISSEMENT,' +
    'ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,ET_CODEPOSTAL,ET_VILLE,' +
    'ET_TELEPHONE,ET_SIRET,ET_FAX,PAT_CODERISQUE,PAT_ORDREAT,PAT_DATEVALIDITE ' + { PT14 }
    ',ETB_SUBROGATION ' + // PT16
    'FROM SALARIES LEFT JOIN ETABLISS ON PSA_ETABLISSEMENT=ET_ETABLISSEMENT ' +
    'LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' + // PT16
    'LEFT JOIN TAUXAT ON PSA_ETABLISSEMENT=PAT_ETABLISSEMENT AND PSA_ORDREAT=PAT_ORDREAT ' +   //PT2 Recup du code rique { PT18 }
    'WHERE PSA_SALARIE="' + Arguments + '" ORDER BY PAT_DATEVALIDITE DESC,PAT_ORDREAT DESC', TRUE); { PT14 }
  if QSal.eof then
  begin
    PGIBox('Salarié inexistant', Ecran.Caption);
    Ferme(QSal);
    exit;
  end; //PT3 Q remplacé par QSal
  Etab := QSal.FindField('PSA_ETABLISSEMENT').asstring; //PT2 Récup du code étab
  Defaut := THLabel(GetControl('ETLIBELLE'));
  if defaut <> nil then Defaut.caption := QSal.FindField('ET_LIBELLE').asstring;
  Defaut := THLabel(GetControl('ETADRESSE123'));
  if defaut <> nil then Defaut.caption := QSal.FindField('ET_ADRESSE1').asstring + ' ' + QSal.FindField('ET_ADRESSE2').asstring + ' ' + QSal.FindField('ET_ADRESSE3').asstring;
  Defaut := THLabel(GetControl('ETCPVILLE'));
  if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('ET_CODEPOSTAL').asstring, 'STR', 3) + '   ' + QSal.FindField('ET_VILLE').asstring;
  Defaut := THLabel(GetControl('ETTELEPHONE'));
  if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('ET_TELEPHONE').asstring, 'STR', 2);
  Defaut := THLabel(GetControl('ETTELECOPIE'));
  if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('ET_FAX').asstring, 'STR', 2);
  Defaut := THLabel(GetControl('ETSIRET'));
  if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('ET_SIRET').asstring, 'STR', 3);
  Defaut := THLabel(GetControl('CODERISQUE')); //PT2 Recup du code risque
  if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('PAT_CODERISQUE').asstring, 'STR', 3);


  Defaut := THLabel(GetControl('PSALIBELLE'));
  if defaut <> nil then
    Defaut.caption := QSal.FindField('PSA_NOMJF').asstring + '  ' + QSal.FindField('PSA_LIBELLE').asstring + '   ' + QSal.FindField('PSA_PRENOM').asstring; { PT15 }
  Defaut := THLabel(GetControl('PSAADRESSE123'));
  if defaut <> nil then
    Defaut.caption := QSal.FindField('PSA_ADRESSE1').asstring + '  ' + QSal.FindField('PSA_ADRESSE2').asstring + '  ' + QSal.FindField('PSA_ADRESSE3').asstring;
  Defaut := THLabel(GetControl('PSASEXE'));
  if defaut <> nil then Defaut.caption := RechDom('PGSEXE', QSal.FindField('PSA_SEXE').asstring, False);
  Defaut := THLabel(GetControl('PSADATENAISSANCE'));
  if defaut <> nil then Defaut.caption := QSal.FindField('PSA_DATENAISSANCE').asstring;

  Defaut := THLabel(GetControl('PSAEMPLOI'));
  if defaut <> nil then Defaut.caption := RechDom('PGLIBEMPLOI', QSal.FindField('PSA_LIBELLEEMPLOI').asstring, False);
  Defaut := THLabel(GetControl('PSAQUALIF'));
  if defaut <> nil then Defaut.caption := RechDom('PGLIBQUALIFICATION', QSal.FindField('PSA_QUALIFICATION').asstring, False);
  Defaut := THLabel(GetControl('PSADATEEMB'));
  if defaut <> nil then Defaut.caption := QSal.FindField('PSA_DATEENTREE').asstring;
  Defaut := THLabel(GetControl('PSAANC'));
  if defaut <> nil then
    if (QSal.FindField('PSA_DATEANCIENNETE').AsDateTime) > IDate1900 then
    begin
      WDate := AncienneteAnnee(QSal.FindField('PSA_DATEANCIENNETE').AsDateTime, date);
      if WDate > 1 then Defaut.caption := IntToStr(WDate) + ' ans' { DEB PT8 }
      else Defaut.caption := IntToStr(WDate) + ' an'; { FIN PT8 }
    end
    else Defaut.caption := '';

  Defaut := THLabel(GetControl('PSACPVILLE'));
  if defaut <> nil then Defaut.caption := FormatCase(QSal.FindField('PSA_CODEPOSTAL').asstring, 'STR', 2) + '    ' + QSal.FindField('PSA_VILLE').asstring;
  Defaut := THLabel(GetControl('PSASECU1'));
  if defaut <> nil then Defaut.caption := FormatCase(Copy(QSal.FindField('PSA_NUMEROSS').asstring, 1, 13), 'STR', 3);
  Defaut := THLabel(GetControl('PSASECU2'));
  if defaut <> nil then Defaut.caption := FormatCase(Copy(QSal.FindField('PSA_NUMEROSS').asstring, 14, 2), 'STR', 3); //PT1

  CheckFr := TCheckBox(GetControl('CKNATFRANC'));
  CheckCee := TCheckBox(GetControl('CKNATCEE'));
  CheckAutre := TCheckBox(GetControl('CKNATAUTRE'));
  if (QSal.FindField('PSA_NATIONALITE').asstring <> '') and (CheckFr <> nil) and (CheckCee <> nil) and (CheckAutre <> nil) then
  begin
    if (QSal.FindField('PSA_NATIONALITE').asstring) = 'FRA' then
    begin
      CheckFr.Checked := True;
      CheckCee.Enabled := False;
      CheckAutre.Enabled := False;
    end;
    if (QSal.FindField('PSA_NATIONALITE').asstring) <> 'FRA' then
      if (RechDom('PGCEE', (QSal.FindField('PSA_NATIONALITE').asstring), False) <> '') then
      begin
        CheckCee.Checked := True;
        CheckFr.Enabled := False;
        CheckAutre.Enabled := False;
      end;
    if (RechDom('PGCEE', (QSal.FindField('PSA_NATIONALITE').asstring), False) = '') then
    begin
      CheckAutre.Checked := True;
      CheckFr.Enabled := False;
      CheckCee.Enabled := False;
    end;
  end;
  subrogation := QSal.FindField('ETB_SUBROGATION').AsString = 'X'; // PT16
  Ferme(QSal);

  St := ''; //DEB PT2    Affectation de la medecine du travail
  Q := OpenSql('SELECT ANN_NOM1,ANN_ALRUE1,ANN_ALRUE2,ANN_ALRUE3,ANN_ALVILLE,ANN_ALCP ' +
    'FROM ANNUAIRE ' +
    'LEFT JOIN ETABCOMPL ON ETB_MEDTRAVGU=ANN_GUIDPER ' +
    'WHERE ETB_ETABLISSEMENT="' + Etab + '" AND ANN_TYPEPER="MED" ', TRUE);
  if not Q.eof then
    st := Q.FindField('ANN_NOM1').asstring + ' ' + Q.FindField('ANN_ALRUE1').asstring + ' ' +
      Q.FindField('ANN_ALRUE2').asstring + ' ' + Q.FindField('ANN_ALRUE3').asstring + ' ' +
      Q.FindField('ANN_ALCP').asstring + ' ' + Q.FindField('ANN_ALVILLE').asstring;
  SetControlText('MEDNOMADR', St);
  Ferme(Q); //FIN PT2

end;


procedure TOF_PGACCTRAVAIL.InitAccTravail(Mode: string);
var
  QAttes: TQuery;
  Check: TcheckBox;
  Edit: THEdit;
  NumEdit: THNumEdit;
  i: integer;
begin
  { DEB PT13-2 }
  if Matricule = '' then exit;
  if Mode = 'CREATION' then
  begin
    { DEB PT13-1 }
    //QAttes:=OpenSql('SELECT PDA_DECLARANTATTES,PDA_LIBELLE,PDA_PRENOM,PDA_QUALDECLARANT,PDA_AUTRE,PDA_VILLE '+
    QAttes := OpenSql('SELECT PDA_DECLARANTATTES FROM DECLARANTATTEST ' +
      'WHERE (PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%' + GblEtab + '%") ' +
      'AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%ACT%" )  ' +
      'ORDER BY PDA_ETABLISSEMENT DESC', True);
    if not QAttes.eof then
    begin
      SetControlText('DECLARANT', QAttes.FindField('PDA_DECLARANTATTES').AsString);
      AffichDeclarant(nil);
    end;
    Ferme(QAttes);
    { FIN PT13-1 }
    SetControlText('DATEFAIT', DateToStr(Date));
    SetControlText('DATEACCTRAV', '');
    SetControlText('DATEDERTRAV', '');
    SetControlText('DATEREPTRAV', '');
    SetControlChecked('CKMOTIFACCTRAV', False);
    SetControlChecked('CKNONREPRIS', False);
    SetControlText('DATEDEBSUBR', '');
    SetControlText('DATEFINSUBR', '');
    SetControlText('DENOMINATION', '');
    SetControlText('INTITULE', '');
    SetControlChecked('PAIEEURO', False);
    // d PT16
    Check := TcheckBox(GetControl('CKMOTIFACCTRAV'));
    if check <> nil then
      if (TypeAbs = 'ATR') or (TypeAbs = 'ATJ') then Check.Checked := True;
    Check := TcheckBox(GetControl('CKMOTIFMALPROF'));
    if check <> nil then
      if TypeAbs = 'MAP' then Check.Checked := True;
    // f PT16


    for i := 1 to 3 do
    begin
      SetControlText('DATEECHEANCE' + IntToStr(i), ''); //Group A
      SetControlText('DATEDEBPER' + IntToStr(i), '');
      SetControlText('DATEFINPER' + IntToStr(i), '');
      SetControlText('MTSAL' + IntToStr(i), '0');
      SetControlChecked('CKSALEURO' + IntToStr(i), False);
      SetControlText('AVANT' + IntToStr(i), '0');
      SetControlChecked('CKAVANT' + IntToStr(i), False);
      SetControlText('INDEM' + IntToStr(i), '0');
      SetControlChecked('CKINDEM' + IntToStr(i), False);
      SetControlText('PARTSAL4' + IntToStr(i), '0');
      SetControlChecked('CKPARTSAL4' + IntToStr(i), False);
      SetControlText('SOUMIS' + IntToStr(i), '0');
      SetControlChecked('CKSOUMIS' + IntToStr(i), False);
      SetControlText('DEDSUP' + IntToStr(i), '0');
      SetControlText('DATEVER' + IntToStr(i), ''); //Group B
      SetControlText('DATEDEBVER' + IntToStr(i), '');
      SetControlText('DATEFINVER' + IntToStr(i), '');
      SetControlText('BRUTVER' + IntToStr(i), '0');
      SetControlChecked('CKBRUTVER' + IntToStr(i), False);
      SetControlText('PARTSAL12' + IntToStr(i), '0');
      SetControlChecked('CKPARTSAL12' + IntToStr(i), False);
      SetControlText('MOTIFARRET' + IntToStr(i), ''); //GROUP C
      SetControlText('DATEDEBARRET' + IntToStr(i), '');
      SetControlText('DATEFINARRET' + IntToStr(i), '');
      SetControlText('BRUTPERDU' + IntToStr(i), '0');
      SetControlChecked('CKBRUTPERDU' + IntToStr(i), False);
      SetControlText('PARTSAL18' + IntToStr(i), '0');
      SetControlChecked('CKPARTSAL18' + IntToStr(i), False);
    end;

  end
  else
    { FIN PT13-2 }
  begin
    QAttes := OpenSql('SELECT PAS_SALARIE,PAS_TRAVAILTEMP,PAS_DERNIERJOUR,PAS_DATEARRET,PAS_REPRISEARRET,' +
      'PAS_MOTIFARRET,PAS_SITUATION,PAS_NONREPRIS,PAS_REPRISEPARTIEL,PAS_CASGEN,' +
      'PAS_MONTANT,PAS_PLUSDE,PAS_PERIODEDEBUT,PAS_PERIODEFIN,PAS_DECLARLIEU,' +
      'PAS_DECLARDATE,PAS_DECLARPERS,PAS_DECLARQUAL,PAS_SUBDEBUT,PAS_SUBFIN,' +
      'PAS_SUBCOMPTE,PAS_SUBCPTEINT,PAS_SUBMONNAIE FROM  ATTESTATIONS ' +
      'WHERE PAS_SALARIE="' + Matricule + '" AND PAS_TYPEATTEST="ACC" ', True);
    if QAttes.eof then //PORTAGECWAS
    begin
      Ferme(QAttes); //PT4
      exit;
    end;

    Edit := THEdit(GetControl('DATEACCTRAV'));
    if (edit <> nil) and (QAttes.FindField('PAS_DATEARRET').AsDateTime <> IDate1900) then
      Edit.text := DateToStr(QAttes.FindField('PAS_DATEARRET').AsDateTime);
    Edit := THEdit(GetControl('DATEDERTRAV'));
    if (edit <> nil) and (QAttes.FindField('PAS_DERNIERJOUR').AsDateTime <> IDate1900) then
      Edit.text := DateToStr(QAttes.FindField('PAS_DERNIERJOUR').AsDateTime);
    Edit := THEdit(GetControl('DATEREPTRAV'));
    if edit <> nil then Edit.text := QAttes.FindField('PAS_REPRISEARRET').asstring;

    Check := TcheckBox(GetControl('CKMOTIFACCTRAV'));
    if check <> nil then
      if QAttes.FindField('PAS_MOTIFARRET').asstring = 'ACC' then Check.Checked := True;
    Check := TcheckBox(GetControl('CKMOTIFMALPROF'));
    if check <> nil then
      if QAttes.FindField('PAS_MOTIFARRET').asstring = 'MPR' then Check.Checked := True;

    Check := TcheckBox(GetControl('CKNONREPRIS'));
    if check <> nil then Check.Checked := (QAttes.FindField('PAS_NONREPRIS').asstring = 'X'); //PT11
    Edit := THEdit(GetControl('LIEUFAIT'));
    if edit <> nil then Edit.text := QAttes.FindField('PAS_DECLARLIEU').asstring;
    Edit := THEdit(GetControl('DATEFAIT'));
    if (edit <> nil) and (QAttes.FindField('PAS_DECLARDATE').AsDateTime <> IDate1900) then
      Edit.text := DateToStr(QAttes.FindField('PAS_DECLARDATE').AsDateTime);
    Edit := THEdit(GetControl('NOMSIGNE'));
    if edit <> nil then Edit.text := QAttes.FindField('PAS_DECLARPERS').asstring;
    Edit := THEdit(GetControl('QUALITESIGNE'));
    if Edit <> nil then Edit.text := QAttes.FindField('PAS_DECLARQUAL').asstring;
    Edit := THEdit(GetControl('DATEDEBSUBR'));
    if (edit <> nil) and (QAttes.FindField('PAS_SUBDEBUT').AsDateTime <> IDate1900) then
      Edit.text := DateToStr(QAttes.FindField('PAS_SUBDEBUT').AsDateTime);
    Edit := THEdit(GetControl('DATEFINSUBR'));
    if (edit <> nil) and (QAttes.FindField('PAS_SUBFIN').AsDateTime <> IDate1900) then
      Edit.text := DateToStr(QAttes.FindField('PAS_SUBFIN').AsDateTime);
    Edit := THEdit(GetControl('DENOMINATION'));
    if edit <> nil then Edit.text := QAttes.FindField('PAS_SUBCOMPTE').asstring;
    Edit := THEdit(GetControl('INTITULE'));
    if edit <> nil then Edit.text := QAttes.FindField('PAS_SUBCPTEINT').asstring;
    Check := TcheckBox(GetControl('PAIEEURO'));
    if check <> nil then Check.Checked := (QAttes.FindField('PAS_SUBMONNAIE').asstring = 'X');
    Ferme(QAttes);

    //Init Tableau
    QAttes := OpenSql('SELECT PAL_SALARIE,PAL_TYPEATTEST,PAL_ORDRE,PAL_MOIS,' +
      'PAL_ECHEANCE,PAL_DATEDEBUT,PAL_DATEFIN,PAL_SALAIRE,PAL_MONNAIESAL,' + //Groupe A
      'PAL_AVANTNAT,PAL_CKAVANTNAT,PAL_INDEMNITE,PAL_CKINDEMNITE,PAL_PARTSAL4,' +
      'PAL_CKPARTSAL4,PAL_SOUMIS,PAL_CKSOUMIS,PAL_DEDSUP,' +
      'PAL_VERSEMENT,PAL_DEBPERVER,PAL_FINPERVER,PAL_MONTANTVER,PAL_CKMONTANTVER,' + //Groupe B
      'PAL_PARTSAL12,PAL_CKPARTSAL12,' +
      'PAL_MOTIF,PAL_DEBARRET,PAL_FINARRET,PAL_BRUTPERDU,PAL_CKBRUTPERDU, ' + //Groupe C
      'PAL_PARTSAL18,PAL_CKPARTSAL18 ' +
      'FROM ATTSALAIRES WHERE PAL_SALARIE="' + Matricule + '" AND PAL_TYPEATTEST="ACC" ' +
      'ORDER BY PAL_MOIS', True);
    i := 0;
    while not QAttes.Eof do
    begin //B For
      i := i + 1;
      Edit := THEdit(GetControl('DATEECHEANCE' + IntToStr(i))); //Group A
      if (Edit <> nil) and (QAttes.FindField('PAL_ECHEANCE').AsDateTime <> IDate1900) then
        Edit.Text := DateToStr(QAttes.FindField('PAL_ECHEANCE').AsDateTime);
      Edit := THEdit(GetControl('DATEDEBPER' + IntToStr(i)));
      if (Edit <> nil) and (QAttes.FindField('PAL_DATEDEBUT').AsDateTime <> IDate1900) then
        Edit.Text := DateToStr(QAttes.FindField('PAL_DATEDEBUT').AsDateTime);
      Edit := THEdit(GetControl('DATEFINPER' + IntToStr(i)));
      if (Edit <> nil) and (QAttes.FindField('PAL_DATEFIN').AsDateTime <> IDate1900) then
        Edit.text := DateToStr(QAttes.FindField('PAL_DATEFIN').AsDateTime);
      NumEdit := THNumEdit(GetControl('MTSAL' + IntToStr(i)));
      if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_SALAIRE').asFloat;
      Check := TCheckBox(GetControl('CKSALEURO' + IntToStr(i)));
      if Check <> nil then
      begin
        Check.checked := (QAttes.FindField('PAL_MONNAIESAL').asstring = 'X');
        Check.Visible := (QAttes.FindField('PAL_SALAIRE').asFloat <> 0);
      end;
      NumEdit := THNumEdit(GetControl('AVANT' + IntToStr(i)));
      if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_AVANTNAT').asFloat;
      Check := TCheckBox(GetControl('CKAVANT' + IntToStr(i)));
      if Check <> nil then
      begin
        Check.checked := (QAttes.FindField('PAL_CKAVANTNAT').asstring = 'X');
        Check.Visible := (QAttes.FindField('PAL_AVANTNAT').asFloat <> 0);
      end;
      NumEdit := THNumEdit(GetControl('INDEM' + IntToStr(i)));
      if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_INDEMNITE').asFloat;
      Check := TCheckBox(GetControl('CKINDEM' + IntToStr(i)));
      if Check <> nil then
      begin
        Check.checked := (QAttes.FindField('PAL_CKINDEMNITE').asstring = 'X');
        Check.Visible := (QAttes.FindField('PAL_INDEMNITE').asFloat <> 0);
      end;
      NumEdit := THNumEdit(GetControl('PARTSAL4' + IntToStr(i)));
      if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_PARTSAL4').asFloat;
      Check := TCheckBox(GetControl('CKPARTSAL4' + IntToStr(i)));
      if Check <> nil then
      begin
        Check.checked := (QAttes.FindField('PAL_CKPARTSAL4').asstring = 'X');
        Check.Visible := (QAttes.FindField('PAL_PARTSAL4').asFloat <> 0);
      end;
      NumEdit := THNumEdit(GetControl('SOUMIS' + IntToStr(i)));
      if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_SOUMIS').asFloat;
      Check := TCheckBox(GetControl('CKSOUMIS' + IntToStr(i)));
      if Check <> nil then
      begin
        Check.checked := (QAttes.FindField('PAL_CKSOUMIS').asstring = 'X');
        Check.Visible := (QAttes.FindField('PAL_SOUMIS').asFloat <> 0);
      end;
      NumEdit := THNumEdit(GetControl('DEDSUP' + IntToStr(i)));
      if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_DEDSUP').asFloat;
      Edit := THEdit(GetControl('DATEVER' + IntToStr(i))); //Group B
      if (Edit <> nil) and (QAttes.FindField('PAL_VERSEMENT').AsDateTime <> IDate1900) then
        Edit.Text := DateToStr(QAttes.FindField('PAL_VERSEMENT').AsDateTime);
      Edit := THEdit(GetControl('DATEDEBVER' + IntToStr(i)));
      if (Edit <> nil) and (QAttes.FindField('PAL_DEBPERVER').AsDateTime <> IDate1900) then
        Edit.Text := DateToStr(QAttes.FindField('PAL_DEBPERVER').AsDateTime);
      Edit := THEdit(GetControl('DATEFINVER' + IntToStr(i)));
      if (Edit <> nil) and (QAttes.FindField('PAL_FINPERVER').AsDateTime <> IDate1900) then
        Edit.text := DateToStr(QAttes.FindField('PAL_FINPERVER').AsDateTime);
      NumEdit := THNumEdit(GetControl('BRUTVER' + IntToStr(i)));
      if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_MONTANTVER').asFloat;
      Check := TCheckBox(GetControl('CKBRUTVER' + IntToStr(i)));
      if Check <> nil then
      begin
        Check.checked := (QAttes.FindField('PAL_CKMONTANTVER').asstring = 'X');
        Check.Visible := (QAttes.FindField('PAL_MONTANTVER').asFloat <> 0);
      end;
      NumEdit := THNumEdit(GetControl('PARTSAL12' + IntToStr(i)));
      if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_PARTSAL12').asFloat;
      Check := TCheckBox(GetControl('CKPARTSAL12' + IntToStr(i)));
      if Check <> nil then
      begin
        Check.checked := (QAttes.FindField('PAL_CKPARTSAL12').asstring = 'X');
        Check.Visible := (QAttes.FindField('PAL_PARTSAL12').asFloat <> 0);
      end;
      Edit := THEdit(GetControl('MOTIFARRET' + IntToStr(i))); //GROUP C
      if Edit <> nil then Edit.text := QAttes.FindField('PAL_MOTIF').asstring;
      Edit := THEdit(GetControl('DATEDEBARRET' + IntToStr(i)));
      if (Edit <> nil) and (QAttes.FindField('PAL_DEBARRET').AsDateTime <> IDate1900) then
        Edit.Text := DateToStr(QAttes.FindField('PAL_DEBARRET').AsDateTime);
      Edit := THEdit(GetControl('DATEFINARRET' + IntToStr(i)));
      if (Edit <> nil) and (QAttes.FindField('PAL_FINARRET').AsDateTime <> IDate1900) then
        Edit.text := DateToStr(QAttes.FindField('PAL_FINARRET').AsDateTime);
      NumEdit := THNumEdit(GetControl('BRUTPERDU' + IntToStr(i)));
      if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_BRUTPERDU').asFloat;
      Check := TCheckBox(GetControl('CKBRUTPERDU' + IntToStr(i)));
      if Check <> nil then
      begin
        Check.checked := (QAttes.FindField('PAL_CKBRUTPERDU').asstring = 'X');
        Check.Visible := (QAttes.FindField('PAL_BRUTPERDU').asFloat <> 0);
      end;
      NumEdit := THNumEdit(GetControl('PARTSAL18' + IntToStr(i)));
      if NumEdit <> nil then NumEdit.Value := QAttes.FindField('PAL_PARTSAL18').asFloat;
      Check := TCheckBox(GetControl('CKPARTSAL18' + IntToStr(i)));
      if Check <> nil then
      begin
        Check.checked := (QAttes.FindField('PAL_CKPARTSAL18').asstring = 'X');
        Check.Visible := (QAttes.FindField('PAL_PARTSAL18').asFloat <> 0);
      end;
      QAttes.next;
    end;

    Ferme(QAttes);
  end;

end;


procedure TOF_PGACCTRAVAIL.UpdateTable(Sender: TObject);
var
  Check: TcheckBox;
  Edit: THEdit;
  NumEdit: THNumEdit;
  DerjourTrav, DateRepArret, MotifArret, NonRepris: string;
  LieuDecla, DateDecla, PersDecla, QualDecla, DateDebSubr: string;
  DateFinSubr, CompteBq, Intitule, Monnaie, DateAccTrav: string;
  StSal, Motif, SalEuro: string;
  Echeance, DateDebPer, DateFinPer, Avant, CkAvant, Indem, CkIndem, PartSal4, CkPartSal4: string;
  Soumis, CkSoumis, DedSup: string;
  DateVer, DateDebVer, DateFinVer, BrutVer, CkBrutVer, PartSal12, CkPartSal12: string;
  DateDebArret, DateFinArret, BrutPerdu, CkBrutPerdu, PartSal18, CkPartSal18: string;
  ordre, i: integer;
  QRech: TQuery;
  Tob_Mere, T: TOB;
begin
  SetFocusControl('PDFACCTRAVAIL');   { PT22 }
  ordre := 1;
  DerjourTrav := DateToStr(Idate1900); // PT17
  Edit := THEdit(GetControl('DATEDERTRAV'));
  if edit <> nil then if Edit.text <> '' then DerjourTrav := Edit.text;
  DateAccTrav := DateToStr(Idate1900); // PT17
  Edit := THEdit(GetControl('DATEACCTRAV'));
  if edit <> nil then if Edit.text <> '' then DateAccTrav := Edit.text;
  DateRepArret := DateToStr(Idate1900); // PT17
  Edit := THEdit(GetControl('DATEREPTRAV'));
  if edit <> nil then if Edit.text <> '' then DateRepArret := Edit.text;
  MotifArret := '';
  Check := TcheckBox(GetControl('CKMOTIFACCTRAV'));
  if check <> nil then if Check.Checked = True then MotifArret := 'ACC';
  Check := TcheckBox(GetControl('CKMOTIFMALPROF'));
  if check <> nil then if Check.Checked = True then MotifArret := 'MPR';
  NonRepris := '-';
  Check := TcheckBox(GetControl('CKNONREPRIS'));
  if check <> nil then if Check.Checked = True then NonRepris := 'X';
  LieuDecla := '';
  Edit := THEdit(GetControl('LIEUFAIT'));
  if edit <> nil then if Edit.text <> '' then LieuDecla := Edit.text;
  DateDecla := DateToStr(Idate1900); // PT17
  Edit := THEdit(GetControl('DATEFAIT'));
  if edit <> nil then if Edit.text <> '' then DateDecla := Edit.text;
  PersDecla := '';
  Edit := THEdit(GetControl('NOMSIGNE'));
  if edit <> nil then if Edit.text <> '' then PersDecla := Edit.text;
  QualDecla := '';
  Edit := THEdit(GetControl('QUALITESIGNE'));
  if Edit <> nil then if Edit.Text <> '' then QualDecla := Edit.Text;
  DateDebSubr := DateToStr(Idate1900); // PT17
  Edit := THEdit(GetControl('DATEDEBSUBR'));
  if edit <> nil then if Edit.text <> '' then DateDebSubr := Edit.text;
  DateFinSubr := DateTostr (Idate1900); // PT17
  Edit := THEdit(GetControl('DATEFINSUBR'));
  if edit <> nil then if Edit.text <> '' then DateFinSubr := Edit.text;
  CompteBq := '';
  Edit := THEdit(GetControl('DENOMINATION'));
  if edit <> nil then if Edit.text <> '' then CompteBq := Edit.text;
  Intitule := '';
  Edit := THEdit(GetControl('INTITULE'));
  if edit <> nil then if Edit.text <> '' then Intitule := Edit.text;
  Monnaie := '-';
  Check := TcheckBox(GetControl('PAIEEURO'));
  if check <> nil then if Check.Checked = True then Monnaie := 'X';

  // d PT16
  if (subrogation = True) and
    ((DateDebSubr = DateToStr(Idate1900)) or
    (DateFinSubr = DateToStr(Idate1900)) or
    (CompteBq = '') or
    (Intitule = '')) then
    PgiBox('Des informations concernant la subrogation sont absentes!', 'Subrogation');
  // f PT16

  { DEB PT10 Refonte suppression SQL }
  QRech := OpenSql('SELECT PAS_ORDRE FROM ATTESTATIONS ' +
    'WHERE PAS_SALARIE="' + Matricule + '" AND PAS_TYPEATTEST="ACC"', True);
  if not QRech.Eof then ordre := QRech.Fields[0].asinteger;
  Ferme(QRech);
  if ordre = 0 then
  begin
    QRech := OpenSql('SELECT MAX(PAS_ORDRE) FROM ATTESTATIONS ', True);
    if not QRech.EOF then //PORTAGECWAS
      ordre := QRech.Fields[0].asinteger + 1
    else
      ordre := 1;
    Ferme(QRech);
  end;

  Tob_Mere := Tob.create('ATTESTATIONS', nil, -1);
  T := Tob.create('ATTESTATIONS', Tob_Mere, -1);
  if Ordre > 1 then T.SelectDB('"' + Matricule + '";"ACC";"' + IntToStr(Ordre) + '"', nil); //PT12
  if T = nil then exit;
  T.PutValue('PAS_SALARIE', Matricule);
  T.PutValue('PAS_TYPEATTEST', 'ACC');
  T.PutValue('PAS_ORDRE', ordre);
  T.PutValue('PAS_DERNIERJOUR', StrToDate(DerjourTrav));
  T.PutValue('PAS_DATEARRET', StrToDate(DateAccTrav));
  T.PutValue('PAS_REPRISEARRET', StrToDate(DateRepArret));
  T.PutValue('PAS_MOTIFARRET', MotifArret);
  T.PutValue('PAS_NONREPRIS', NonRepris);
  T.PutValue('PAS_DECLARDATE', StrToDate(DateDecla));
  T.PutValue('PAS_DECLARPERS', PersDecla);
  T.PutValue('PAS_DECLARQUAL', QualDecla);
  T.PutValue('PAS_DECLARLIEU', LieuDecla);
  T.PutValue('PAS_SUBDEBUT', StrToDate(DateDebSubr));
  T.PutValue('PAS_SUBFIN', StrToDate(DateFinSubr));
  T.PutValue('PAS_SUBCOMPTE', CompteBq);
  T.PutValue('PAS_SUBCPTEINT', Intitule);
  T.PutValue('PAS_SUBMONNAIE', Monnaie);
  T.InsertOrUpdateDB;
  if Tob_Mere <> nil then Tob_Mere.free;


  Tob_Mere := Tob.create('ATTSALAIRES', nil, -1);
  { FIN PT10 }
  //MAJ Tableau
  for i := 1 to 3 do
  begin //B For
    Echeance := DateToStr(Idate1900); // PT17
    Edit := THEdit(GetControl('DATEECHEANCE' + IntToStr(i)));
    if Edit <> nil then if Edit.Text <> '' then Echeance := Edit.text;
    DateDebPer := DateToStr(Idate1900); // PT17
    Edit := THEdit(GetControl('DATEDEBPER' + IntToStr(i)));
    if Edit <> nil then if Edit.Text <> '' then DateDebPer := Edit.text;
    DateFinPer := DateToStr(Idate1900); // PT17
    Edit := THEdit(GetControl('DATEFINPER' + IntToStr(i)));
    if Edit <> nil then if Edit.Text <> '' then DateFinPer := Edit.text;
    NumEdit := THNumEdit(GetControl('MTSAL' + IntToStr(i)));
    if NumEdit <> nil then StSal := RemplaceVirguleParPoint(FloatToStr(NumEdit.value)); //FloatToStr(NumEdit.Value);
    Check := TCheckBox(GetControl('CKSALEURO' + IntToStr(i)));
    if Check <> nil then if Check.checked = True then SalEuro := 'X' else SalEuro := '-';
    NumEdit := THNumEdit(GetControl('AVANT' + IntToStr(i)));
    if NumEdit <> nil then Avant := RemplaceVirguleParPoint(FloatToStr(NumEdit.value));
    Check := TCheckBox(GetControl('CKAVANT' + IntToStr(i)));
    if Check <> nil then if Check.checked = True then CkAvant := 'X' else CkAvant := '-';
    NumEdit := THNumEdit(GetControl('INDEM' + IntToStr(i)));
    if NumEdit <> nil then Indem := RemplaceVirguleParPoint(FloatToStr(NumEdit.value));
    Check := TCheckBox(GetControl('CKINDEM' + IntToStr(i)));
    if Check <> nil then if Check.checked = True then CkIndem := 'X' else CkIndem := '-';
    NumEdit := THNumEdit(GetControl('PARTSAL4' + IntToStr(i)));
    if NumEdit <> nil then PartSal4 := RemplaceVirguleParPoint(FloatToStr(NumEdit.value));
    Check := TCheckBox(GetControl('CKPARTSAL4' + IntToStr(i)));
    if Check <> nil then if Check.checked = True then CkPartSal4 := 'X' else CkPartSal4 := '-';
    NumEdit := THNumEdit(GetControl('SOUMIS' + IntToStr(i)));
    if NumEdit <> nil then Soumis := RemplaceVirguleParPoint(FloatToStr(NumEdit.value));
    Check := TCheckBox(GetControl('CKSOUMIS' + IntToStr(i)));
    if Check <> nil then if Check.checked = True then CkSoumis := 'X' else CkSoumis := '-';
    NumEdit := THNumEdit(GetControl('DEDSUP' + IntToStr(i)));
    if NumEdit <> nil then DedSup := RemplaceVirguleParPoint(FloatToStr(NumEdit.value));
    DateVer := DateToStr(Idate1900); // PT17
    Edit := THEdit(GetControl('DATEVER' + IntToStr(i))); //Group B
    if Edit <> nil then if Edit.Text <> '' then DateVer := Edit.text;
    DateDebVer := DateToStr(Idate1900); // PT17
    Edit := THEdit(GetControl('DATEDEBVER' + IntToStr(i)));
    if Edit <> nil then if Edit.Text <> '' then DateDebVer := Edit.text;
    DateFinVer := DateToStr(Idate1900); // PT17
    Edit := THEdit(GetControl('DATEFINVER' + IntToStr(i)));
    if Edit <> nil then if Edit.Text <> '' then DateFinVer := Edit.text;
    NumEdit := THNumEdit(GetControl('BRUTVER' + IntToStr(i)));
    if NumEdit <> nil then BrutVer := RemplaceVirguleParPoint(FloatToStr(NumEdit.value));
    Check := TCheckBox(GetControl('CKBRUTVER' + IntToStr(i)));
    if Check <> nil then if Check.checked = True then CkBrutVer := 'X' else CkBrutVer := '-';
    NumEdit := THNumEdit(GetControl('PARTSAL12' + IntToStr(i)));
    if NumEdit <> nil then PartSal12 := RemplaceVirguleParPoint(FloatToStr(NumEdit.value));
    Check := TCheckBox(GetControl('CKPARTSAL12' + IntToStr(i)));
    if Check <> nil then if Check.checked = True then CkPartSal12 := 'X' else CkPartSal12 := '-';
    Motif := '';
    Edit := THEdit(GetControl('MOTIFARRET' + IntToStr(i))); //GROUP C
    if Edit <> nil then if Edit.Text <> '' then Motif := Edit.text;
    DateDebArret := DateToStr(Idate1900); // PT17
    Edit := THEdit(GetControl('DATEDEBARRET' + IntToStr(i)));
    if Edit <> nil then if Edit.Text <> '' then DateDebArret := Edit.text;
    DateFinArret := DateToStr(Idate1900); // PT17
    Edit := THEdit(GetControl('DATEFINARRET' + IntToStr(i)));
    if Edit <> nil then if Edit.Text <> '' then DateFinArret := Edit.text;
    NumEdit := THNumEdit(GetControl('BRUTPERDU' + IntToStr(i)));
    if NumEdit <> nil then BrutPerdu := RemplaceVirguleParPoint(FloatToStr(NumEdit.value));
    Check := TCheckBox(GetControl('CKBRUTPERDU' + IntToStr(i)));
    if Check <> nil then if Check.checked = True then CkBrutPerdu := 'X' else CkBrutPerdu := '-';
    NumEdit := THNumEdit(GetControl('PARTSAL18' + IntToStr(i)));
    if NumEdit <> nil then PartSal18 := RemplaceVirguleParPoint(FloatToStr(NumEdit.value));
    Check := TCheckBox(GetControl('CKPARTSAL18' + IntToStr(i)));
    if Check <> nil then if Check.checked = True then CkPartSal18 := 'X' else CkPartSal18 := '-';
    { DEB PT10 Refonte suppression SQL }
    T := Tob.create('ATTSALAIRES', Tob_Mere, -1);
    if Ordre > 1 then T.SelectDB('"' + IntToStr(Ordre) + '";"' + IntToStr(i) + '"', nil); //PT12
    if T = nil then exit;
    T.PutValue('PAL_SALARIE', Matricule);
    T.PutValue('PAL_TYPEATTEST', 'ACC');
    T.PutValue('PAL_ORDRE', ordre);
    T.PutValue('PAL_MOIS', i);
    T.PutValue('PAL_ECHEANCE', StrToDate(Echeance)); //Groupe A
    T.PutValue('PAL_DATEDEBUT', StrToDate(DateDebPer));
    T.PutValue('PAL_DATEFIN', StrToDate(DateFinPer));
    T.PutValue('PAL_SALAIRE', Valeur(StSal));
    T.PutValue('PAL_MONNAIESAL', SalEuro);
    T.PutValue('PAL_AVANTNAT', Valeur(Avant));  { PT21 }
    T.PutValue('PAL_CKAVANTNAT', CkAvant);
    T.PutValue('PAL_INDEMNITE', Valeur(Indem));  { PT21 }
    T.PutValue('PAL_CKINDEMNITE', CkIndem);
    T.PutValue('PAL_PARTSAL4', Valeur(PartSal4));  { PT21 }
    T.PutValue('PAL_CKPARTSAL4', CkPartSal4);
    T.PutValue('PAL_SOUMIS', Valeur(Soumis));  { PT21 }
    T.PutValue('PAL_CKSOUMIS', CkSoumis);
    T.PutValue('PAL_DEDSUP', Valeur(DedSup));  { PT21 }
    T.PutValue('PAL_VERSEMENT', StrToDate(DateVer)); //Groupe B
    T.PutValue('PAL_DEBPERVER', StrToDate(DateDebVer));
    T.PutValue('PAL_FINPERVER', StrToDate(DateFinVer));
    T.PutValue('PAL_MONTANTVER', Valeur(BrutVer));  { PT21 }
    T.PutValue('PAL_CKMONTANTVER', CkBrutVer);
    T.PutValue('PAL_PARTSAL12', Valeur(PartSal12));   { PT21 }
    T.PutValue('PAL_CKPARTSAL12', CkPartSal12);
    T.PutValue('PAL_MOTIF', Motif); //Groupe C
    T.PutValue('PAL_DEBARRET', StrToDate(DateDebArret));
    T.PutValue('PAL_FINARRET', StrToDate(DateFinArret));
    T.PutValue('PAL_BRUTPERDU', Valeur(BrutPerdu));  { PT21 }
    T.PutValue('PAL_CKBRUTPERDU', CkBrutPerdu);
    T.PutValue('PAL_PARTSAL18', Valeur(PartSal18));  { PT21 }
    T.PutValue('PAL_CKPARTSAL18', CkPartSal18);
  end; //End For
  if Tob_Mere <> nil then
  begin
    if Tob_Mere.detail.count > 0 then
      Tob_Mere.InsertOrUpdateDB;
    Tob_Mere.free;
  end;
  { FIN PT10 }
end;

procedure TOF_PGACCTRAVAIL.DesactiveNat(Sender: TObject);
var
  CheckFr, CheckCee, CheckAutre: TCheckBox;
begin
  CheckFr := TCheckBox(GetControl('CKNATFRANC'));
  CheckCee := TCheckBox(GetControl('CKNATCEE'));
  CheckAutre := TCheckBox(GetControl('CKNATAUTRE'));
  if (CheckFr <> nil) and (CheckCee <> nil) and (CheckAutre <> nil) then
  begin
    if (CheckFr.Checked = True) and (CheckCee.Checked = False) and (CheckAutre.Checked = False) then
    begin
      CheckCee.Checked := False;
      CheckCee.Enabled := False;
      CheckAutre.Checked := False;
      CheckAutre.Enabled := False;
    end;
    if (CheckFr.Checked = False) and (CheckCee.enabled = False) and (CheckAutre.enabled = False) then
    begin
      CheckCee.Enabled := True;
      CheckAutre.Enabled := True;
    end;

    if (CheckCee.Checked = True) and (CheckFr.Checked = False) and (CheckAutre.Checked = False) then
    begin
      CheckFr.Checked := False;
      CheckFr.Enabled := False;
      CheckAutre.Checked := False;
      CheckAutre.Enabled := False;
    end;
    if (CheckCee.Checked = False) and (CheckFr.enabled = False) and (CheckAutre.enabled = False) then
    begin
      CheckFr.Enabled := True;
      CheckAutre.Enabled := True;
    end;

    if (CheckAutre.Checked = True) and (CheckFr.Checked = False) and (CheckCee.Checked = False) then
    begin
      CheckFr.Checked := False;
      CheckFr.Enabled := False;
      CheckCee.Checked := False;
      CheckCee.Enabled := False;
    end;
    if (CheckAutre.Checked = False) and (CheckFr.enabled = False) and (CheckCee.enabled = False) then
    begin
      CheckFr.Enabled := True;
      CheckCee.Enabled := True;
    end;
  end;
end;

procedure TOF_PGACCTRAVAIL.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGACCTRAVAIL.DesactiveVict(Sender: TObject);
var
  checkoui, CheckNon: TCheckBox;
begin
  CheckOui := TCheckBox(GetControl('CKVICTOUI'));
  CheckNon := TCheckBox(GetControl('CKVICTNON'));
  if (CheckOui <> nil) and (CheckNon <> nil) then
  begin
    if (CheckOui.Checked = True) and (CheckNon.Checked = False) then
    begin
      CheckNon.Checked := False;
      CheckNon.Enabled := False;
    end;
    if (CheckOui.Checked = False) and (CheckNon.enabled = False) then CheckNon.Enabled := True;

    if (CheckNon.Checked = True) and (CheckOui.Checked = False) then
    begin
      CheckOui.Checked := False;
      CheckOui.Enabled := False;
    end;
    if (CheckNon.Checked = False) and (CheckOui.enabled = False) then CheckOui.Enabled := True;
  end;
end;

procedure TOF_PGACCTRAVAIL.DesactiveMaintien(Sender: TObject);
var
  checkoui, CheckNon, Int, part: TCheckBox;
begin
  CheckOui := TCheckBox(GetControl('MAINTIENOUI'));
  CheckNon := TCheckBox(GetControl('MAINTIENNON'));
  Int := TCheckBox(GetControl('MAINTIENINT'));
  Part := TCheckBox(GetControl('MAINTIENPART'));

  if (CheckOui <> nil) and (CheckNon <> nil) and (int <> nil) and (Part <> nil) then
  begin
    if (CheckOui.Checked = True) and (CheckNon.Checked = False) then
    begin
      CheckNon.Checked := False;
      CheckNon.Enabled := False;
      Int.Enabled := True;
      Part.Enabled := True;
    end;
    if (CheckOui.Checked = False) and (CheckNon.enabled = False) then
    begin
      CheckNon.Enabled := True;
      Int.Enabled := False;
      Part.Enabled := False;
    end;

    if (CheckNon.Checked = True) and (CheckOui.Checked = False) then
    begin
      CheckOui.Checked := False;
      CheckOui.Enabled := False;
    end;
    if (CheckNon.Checked = False) and (CheckOui.enabled = False) then CheckOui.Enabled := True;
  end;
end;

procedure TOF_PGACCTRAVAIL.GereMaintien(Sender: TObject);
var
  Int, part: TCheckBox;
begin
  Int := TCheckBox(GetControl('MAINTIENINT_'));
  Part := TCheckBox(GetControl('MAINTIENPART_'));
  if (Int <> nil) and (Part <> nil) then
  begin
    if (Int.Checked = True) and (Part.Checked = False) then
    begin
      Part.Checked := False;
      Part.Enabled := False;
    end;
    if (Int.Checked = False) and (Part.enabled = False) then Part.Enabled := True;

    if (Part.Checked = True) and (Int.Checked = False) then
    begin
      Int.Checked := False;
      Int.Enabled := False;
    end;
    if (Part.Checked = False) and (Int.enabled = False) then Int.Enabled := True;
  end;
end;

procedure TOF_PGACCTRAVAIL.AffectSalaire(Sender: TObject);
var
  QRech, QPlus: TQuery;
  Edit, Zone: THedit;
  DerJour: TDateTime;
  UsDateDeb, UsDateFin: string;
  MontZone: THNumEdit;
  i, j: integer;
  CkEuro: TCheckBox;
  Mt :  Double;
begin
  {PT5-2 On prend le dernier jour travaillé au lieu de la date d'accident
  Edit:=THEdit(GetControl('DATEACCTRAV'));}
  Edit := THEdit(GetControl('DATEDERTRAV'));
  if (Edit <> nil) then
    if Edit.Text <> '' then
    begin
      DerJour := StrToDate(Edit.text);
      UsDateDeb := UsDateTime(DebutDeMois(PlusMois(DerJour, -1)));
      UsDateFin := UsDateTime(DerJour);

      QRech := OpenSql('SELECT DISTINCT PPU_DATEDEBUT,PPU_DATEFIN,PPU_PAYELE,' +
        //   'SUM(PPU_CBRUT) AS PPU_CBRUT, SUM(PPU_OCBRUT) AS PPU_OCBRUT,'+ PT6
        'SUM(PPU_CBRUTFISCAL) AS PPU_CBRUTFISCAL, SUM(PPU_OCBRUTFISCAL) AS PPU_OCBRUTFISCAL,' +
        'SUM(PPU_CCOUTSALARIE) AS PPU_CCOUTSALARIE, SUM(PPU_OCCOUTSALARIE) AS PPU_OCCOUTSALARIE ' +
        'FROM PAIEENCOURS ' +
        'WHERE PPU_SALARIE="' + Matricule + '" ' +
        'AND PPU_DATEDEBUT>="' + UsDateDeb + '" AND PPU_DATEFIN<"' + UsDateFin + '" ' +
        'GROUP BY PPU_DATEDEBUT,PPU_DATEFIN,PPU_PAYELE ORDER BY PPU_DATEDEBUT DESC,PPU_DATEFIN DESC', TRUE);
      j := 0;
      QRech.First;                  { PT19 }
      while not QRech.Eof do        { PT19 }
      begin
        j := j + 1;
        Zone := THEdit(GetControl('DATEECHEANCE' + IntToStr(j)));
        if Zone <> nil then Zone.text := QRech.FindField('PPU_PAYELE').asstring;
        Zone := THEdit(GetControl('DATEDEBPER' + IntToStr(j)));
        if Zone <> nil then Zone.text := QRech.FindField('PPU_DATEDEBUT').asstring;
        Zone := THEdit(GetControl('DATEFINPER' + IntToStr(j)));
        if Zone <> nil then Zone.text := QRech.FindField('PPU_DATEFIN').asstring;
        {Gestion de la date de passage à l'euro  }
        if (VH_Paie.PGTenueEuro = FALSE) then
        begin
          //Salaire de base
          MontZone := THNumEdit(GetControl('MTSAL' + IntToStr(j)));
          if MontZone <> nil then
            MontZone.Value := QRech.FindField('PPU_CBRUTFISCAL').asFloat; //PT6
          CkEuro := TCheckBox(GetControl('CKSALEURO' + IntToStr(j)));
          if CkEuro <> nil then
          begin
            CkEuro.Checked := VH_Paie.PGTenueEuro;
            CkEuro.Visible := True;
          end;
          { DEB PT20 }
          QPlus := OpenSql( 'SELECT SUM(PHB_MTSALARIAL) MTSALARIAL '+
                   'FROM HISTOBULLETIN '+
                   'LEFT JOIN COTISATION ON ##PCT_PREDEFINI## PHB_RUBRIQUE = PCT_RUBRIQUE AND PHB_NATURERUB = PCT_NATURERUB '+
                   'WHERE PHB_NATURERUB="COT" AND PCT_ETUDEDROIT = "X" '+
                   'AND PHB_SALARIE="' + Matricule + '" ' +
                   'AND PHB_DATEDEBUT="' + USDateTime(StrToDate(GetControlText('DATEDEBPER' + IntToStr(j)))) + '" '+
                   'AND PHB_DATEFIN="' + USDateTime(StrToDate(GetControlText('DATEFINPER' + IntToStr(j)))) + '" ',True);
         if Not QPlus.eof then
            Begin
            //DEB PT5-3 Colonne 7
            Mt := QRech.FindField('PPU_CCOUTSALARIE').asFloat - QPlus.FindField('MTSALARIAL').asFloat;
            SetControlText('PARTSAL4' + IntToStr(j), FloatToStr(Mt));
            if VH_Paie.PGTenueEuro then
            SetControlText('CKPARTSAL4' + IntToStr(j), 'X')
            else
            SetControlText('CKPARTSAL4' + IntToStr(j), '-');
            SetControlVisible('CKPARTSAL4' + IntToStr(j), True);
            //FIN PT5-3
            End;
          Ferme(QPlus);
          { FIN PT20 }
        end
        else
        begin
          if (VH_Paie.PGDateBasculEuro <= DebutDeMois(PlusMois(DerJour, j - 1))) then
          begin
            //Salaire de base
            MontZone := THNumEdit(GetControl('MTSAL' + IntToStr(j)));
            if MontZone <> nil then
              MontZone.Value := QRech.FindField('PPU_CBRUTFISCAL').asFloat; //PT6
            CkEuro := TCheckBox(GetControl('CKSALEURO' + IntToStr(j)));
            if CkEuro <> nil then
            begin
              CkEuro.Checked := VH_Paie.PGTenueEuro;
              CkEuro.Visible := True;
            end;
            { DEB PT20 }
            QPlus := OpenSql( 'SELECT SUM(PHB_MTSALARIAL) MTSALARIAL '+
                   'FROM HISTOBULLETIN '+
                   'LEFT JOIN COTISATION ON ##PCT_PREDEFINI## PHB_RUBRIQUE = PCT_RUBRIQUE AND PHB_NATURERUB = PCT_NATURERUB '+
                   'WHERE PHB_NATURERUB="COT" AND PCT_ETUDEDROIT = "X" '+
                   'AND PHB_SALARIE="' + Matricule + '" ' +
                   'AND PHB_DATEDEBUT="' + USDateTime(StrToDate(GetControlText('DATEDEBPER' + IntToStr(j)))) + '" '+
                   'AND PHB_DATEFIN="' + USDateTime(StrToDate(GetControlText('DATEFINPER' + IntToStr(j)))) + '" ',True);
            if Not QPlus.eof then
            Begin
            //DEB PT5-3 Colonne 7
            Mt := QRech.FindField('PPU_CCOUTSALARIE').asFloat - QPlus.FindField('MTSALARIAL').asFloat;
            SetControlText('PARTSAL4' + IntToStr(j), FloatToStr(Mt));
            if VH_Paie.PGTenueEuro then
            SetControlText('CKPARTSAL4' + IntToStr(j), 'X')
            else
            SetControlText('CKPARTSAL4' + IntToStr(j), '-');
            SetControlVisible('CKPARTSAL4' + IntToStr(j), True);
            //FIN PT5-3
            End;
          Ferme(QPlus);
          { FIN PT20 }
          end
          else
          begin
            //Salaire de base
            MontZone := THNumEdit(GetControl('MTSAL' + IntToStr(j)));
            if MontZone <> nil then
              MontZone.Value := QRech.FindField('PPU_OCBRUTFISCAL').asFloat; //PT6
            CkEuro := TCheckBox(GetControl('CKSALEURO' + IntToStr(j)));
            if CkEuro <> nil then
            begin
              CkEuro.Checked := False;
              CkEuro.Visible := True;
            end;
            { DEB PT20 }
            QPlus := OpenSql( 'SELECT SUM(PHB_MTSALARIAL) MTSALARIAL '+
                   'FROM HISTOBULLETIN '+
                   'LEFT JOIN COTISATION ON ##PCT_PREDEFINI## PHB_RUBRIQUE = PCT_RUBRIQUE AND PHB_NATURERUB = PCT_NATURERUB '+
                   'WHERE PHB_NATURERUB="COT" AND PCT_ETUDEDROIT = "X" '+
                   'AND PHB_SALARIE="' + Matricule + '" ' +
                   'AND PHB_DATEDEBUT="' + USDateTime(StrToDate(GetControlText('DATEDEBPER' + IntToStr(j)))) + '" '+
                   'AND PHB_DATEFIN="' + USDateTime(StrToDate(GetControlText('DATEFINPER' + IntToStr(j)))) + '" ',True);
            if Not QPlus.eof then
            Begin
            //DEB PT5-3 Colonne 7
            Mt := QRech.FindField('PPU_CCOUTSALARIE').asFloat - QPlus.FindField('MTSALARIAL').asFloat;
            SetControlText('PARTSAL4' + IntToStr(j), FloatToStr(Mt));
            if VH_Paie.PGTenueEuro then
            SetControlText('CKPARTSAL4' + IntToStr(j), 'X')
            else
            SetControlText('CKPARTSAL4' + IntToStr(j), '-');
            SetControlVisible('CKPARTSAL4' + IntToStr(j), True);
            //FIN PT5-3
            End;
            Ferme(QPlus);
            { FIN PT20 }
          end;
        end;
        {FIN Gestion de la date de passage à l'euro}
        if j = 3 then Break;
        QRech.Next;  { PT19 }
      end;
      Ferme(QRech);
      if j < 3 then
        for i := j + 1 to 3 do
        begin
          Zone := THEdit(GetControl('DATEECHEANCE' + IntToStr(i)));
          if Zone <> nil then Zone.text := '';
          Zone := THEdit(GetControl('DATEDEBPER' + IntToStr(i)));
          if Zone <> nil then Zone.text := '';
          Zone := THEdit(GetControl('DATEFINPER' + IntToStr(i)));
          if Zone <> nil then Zone.text := '';
          MontZone := THNumEdit(GetControl('MTSAL' + IntToStr(i)));
          if MontZone <> nil then MontZone.Value := 0;
          CkEuro := TCheckBox(GetControl('CKSALEURO' + IntToStr(i)));
          if CkEuro <> nil then
          begin
            CkEuro.Checked := False;
            CkEuro.Visible := False;
          end;
          //DEB PT5-4 remise à blanc colonne 7
          SetControlText('PARTSAL4' + IntToStr(i), '0');
          SetControlText('CKPARTSAL4' + IntToStr(i), '-');
          //FIN PT5-4
        end;
    end;
end;



procedure TOF_PGACCTRAVAIL.AfficheCheck(Sender: TObject);
var
  NumEdit: THNumEdit;
  check: TCheckBox;
  NomCheck, StI: string;
begin
  NumEdit := THNumEdit(Sender);
  if NumEdit = nil then exit;
  NomCheck := 'CKSALEURO';
  if Pos('MTSAL', NumEdit.name) = 1 then NomCheck := 'CKSALEURO'
  else if Pos('AVANT', NumEdit.name) = 1 then NomCheck := 'CKAVANT'
  else if Pos('INDEM', NumEdit.name) = 1 then NomCheck := 'CKINDEM'
  else if Pos('PARTSAL4', NumEdit.name) = 1 then NomCheck := 'CKPARTSAL4'
  else if Pos('SOUMIS', NumEdit.name) = 1 then NomCheck := 'CKSOUMIS'
  else if Pos('BRUTVER', NumEdit.name) = 1 then NomCheck := 'CKBRUTVER'
  else if Pos('PARTSAL12', NumEdit.name) = 1 then NomCheck := 'CKPARTSAL12'
  else if Pos('BRUTPERDU', NumEdit.name) = 1 then NomCheck := 'CKBRUTPERDU'
  else if Pos('PARTSAL18', NumEdit.name) = 1 then NomCheck := 'CKPARTSAL18';
  StI := Copy(NumEdit.name, Length(NumEdit.name), 1);
  check := TCheckBox(GetControl(NomCheck + StI));
  if Check = nil then exit;
  if NumEdit.value <> 0 then
    check.Visible := True
  else
  begin
    check.Checked := False;
    check.Visible := False;
  end;
  {  For i:=1 to 3 do
    begin
    NumEdit:=THNumEdit(GetControl('MTSAL'+IntToStr(i)));
    check :=TCheckBox(GetControl('CKSALEURO'+IntToStr(i)));
    if (NumEdit<>nil) and (check<>nil) then
       if NumEdit.value<>0 then check.Visible:=True else check.Visible:=False;
    NumEdit:=THNumEdit(GetControl('AVANT'+IntToStr(i)));
    check :=TCheckBox(GetControl('CKAVANT'+IntToStr(i)));
    if (NumEdit<>nil) and (check<>nil) then
       if NumEdit.value<>0 then check.Visible:=True else  check.Visible:=False;
    NumEdit:=THNumEdit(GetControl('INDEM'+IntToStr(i)));
    check :=TCheckBox(GetControl('CKINDEM'+IntToStr(i)));
    if (NumEdit<>nil) and (check<>nil) then
      if NumEdit.value<>0 then check.Visible:=True else  check.Visible:=False;
    NumEdit:=THNumEdit(GetControl('PARTSAL4'+IntToStr(i)));
    check :=TCheckBox(GetControl('CKPARTSAL4'+IntToStr(i)));
    if (NumEdit<>nil) and (check<>nil) then
      if NumEdit.value<>0 then check.Visible:=True else  check.Visible:=False;
    NumEdit:=THNumEdit(GetControl('SOUMIS'+IntToStr(i)));
    check :=TCheckBox(GetControl('CKSOUMIS'+IntToStr(i)));
    if (NumEdit<>nil) and (check<>nil) then
      if NumEdit.value<>0 then check.Visible:=True else  check.Visible:=False;
    NumEdit:=THNumEdit(GetControl('BRUTVER'+IntToStr(i)));
    check :=TCheckBox(GetControl('CKBRUTVER'+IntToStr(i)));
    if (NumEdit<>nil) and (check<>nil) then
      if NumEdit.value<>0 then check.Visible:=True else  check.Visible:=False;
      NumEdit:=THNumEdit(GetControl('PARTSAL12'+IntToStr(i)));
    check :=TCheckBox(GetControl('CKPARTSAL12'+IntToStr(i)));
    if (NumEdit<>nil) and (check<>nil) then
      if NumEdit.value<>0 then check.Visible:=True else  check.Visible:=False;
    NumEdit:=THNumEdit(GetControl('BRUTPERDU'+IntToStr(i)));
    check :=TCheckBox(GetControl('CKBRUTPERDU'+IntToStr(i)));
    if (NumEdit<>nil) and (check<>nil) then
      if NumEdit.value<>0 then check.Visible:=True else  check.Visible:=False;
    NumEdit:=THNumEdit(GetControl('PARTSAL18'+IntToStr(i)));
    check :=TCheckBox(GetControl('CKPARTSAL18'+IntToStr(i)));
    if (NumEdit<>nil) and (check<>nil) then
      if NumEdit.value<>0 then check.Visible:=True else  check.Visible:=False;
    end;  }
end;
procedure TOF_PGACCTRAVAIL.GestionSubrogation(Sender: TObject);
var
  deno: THedit;
  QRech: Tquery;
  StPlus : String;
begin
  //PORTAGECWAS
  SetControlEnabled('BANQUE', False);
  SetControlText('BANQUE', '');
  SetControlEnabled('GUICHET', False);
  SetControlText('GUICHET', '');
  SetControlEnabled('COMPTE', False);
  SetControlText('COMPTE', '');
  SetControlEnabled('CLE', False);
  SetControlText('CLE', '');
  deno := THedit(GetControl('DENOMINATION'));
  if deno <> nil then
    if deno.text <> '' then
    begin
      StPlus:= PGBanqueCP (True);            //PT23

      QRech := opensql('SELECT BQ_ETABBQ,BQ_NUMEROCOMPTE,BQ_CLERIB,BQ_GUICHET ' +
        'FROM BANQUECP WHERE BQ_GENERAL="' + deno.text + '"'+StPlus, TRUE); //PT9
      if not Qrech.eof then //PORTAGECWAS
      begin
        SetControlEnabled('BANQUE', True);
        SetControlText('BANQUE', QRech.FindField('BQ_ETABBQ').asstring);
        SetControlEnabled('GUICHET', True);
        SetControlText('GUICHET', QRech.FindField('BQ_GUICHET').asstring);
        SetControlEnabled('COMPTE', True);
        SetControlText('COMPTE', QRech.FindField('BQ_NUMEROCOMPTE').asstring);
        SetControlEnabled('CLE', True);
        SetControlText('CLE', QRech.FindField('BQ_CLERIB').asstring);
      end;
      Ferme(QRech);
    end;
end;



procedure TOF_PGACCTRAVAIL.Imprimer(Sender: TObject);
var
  Pages: TPageControl;
  {$IFDEF EAGLCLIENT}
  StPages : String;
  {$ENDIF}
begin
  Pages := TPageControl(GetControl('PAGES'));
  if Pages <> nil then
  //PT7   : 01/07/2003 PH V_42 Portage CWAS
  {$IFNDEF EAGLCLIENT}
    LanceEtat('E', 'PAT', 'ACT', True, False, False, Pages, '', '', False);
  {$ELSE}
    StPages := AglGetCriteres(Pages, FALSE);
  LanceEtat('E', 'PAT', 'ACT', True, False, False, nil, '', '', False, 0, StPages);
  {$ENDIF}
end;

{ DEB PT13-1 }
procedure TOF_PGACCTRAVAIL.AffichDeclarant(Sender: TObject);
var
  St: string;
begin
  if GetControlText('DECLARANT') = '' then exit;
  St := RechDom('PGDECLARANTATTEST', GetControlText('DECLARANT'), False);
  SetControlText('NOMSIGNE', RechDom('PGDECLARANTATTEST', GetControlText('DECLARANT'), False));
  SetControlText('LIEUFAIT', RechDom('PGDECLARANTVILLE', GetControlText('DECLARANT'), False));
  St := RechDom('PGDECLARANTQUAL', GetControlText('DECLARANT'), False);
  if St = 'AUT' then SetControlText('QUALITESIGNE', RechDom('PGDECLARANTAUTRE', GetControlText('DECLARANT'), False))
  else SetControlText('QUALITESIGNE', RechDom('PGQUALDECLARANT2', St, False));
end;
{ FIN PT13-1 }
{ DEB PT13-2 }
procedure TOF_PGACCTRAVAIL.NewAttest(Sender: TObject);
begin
  if PGIAsk('Voulez-vous saisir une nouvelle attestation?', Ecran.Caption) = MrYes then
    InitAccTravail('CREATION');
end;
{ FIN PT13-2 }
{ DEB PT14 }
procedure TOF_PGACCTRAVAIL.AffectCodeRisque(Sender: TObject);
var
  DateAcc: TDateTime;
  Q: TQuery;
begin
  if not IsValidDate(GetControlText('DATEACCTRAV')) then exit;
  DateAcc := StrToDate(GetControlText('DATEACCTRAV'));
  Q := OpenSql('SELECT PAT_CODERISQUE,PAT_ORDREAT,PAT_DATEVALIDITE  ' +
    'FROM TAUXAT '+
    'LEFT JOIN SALARIES ON PSA_ETABLISSEMENT=PAT_ETABLISSEMENT AND PSA_ORDREAT=PAT_ORDREAT ' +  { PT18 }
    'WHERE PSA_SALARIE="' + Matricule + '" AND PAT_DATEVALIDITE<="' + USDateTime(DateAcc) + '" ' +
    'ORDER BY PAT_DATEVALIDITE DESC,PAT_ORDREAT DESC', True);
  if not Q.eof then
    SetControlText('CODERISQUE', FormatCase(Q.FindField('PAT_CODERISQUE').asstring, 'STR', 3));
  Ferme(Q);
end;
{ FIN PT14 }
initialization
  registerclasses([TOF_PGACCTRAVAIL]);
end.

