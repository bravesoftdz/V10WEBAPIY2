{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 28/08/2001
Modifié le ... :   /  /
Description .. : Reprise des CP
Mots clefs ... : PAIE;CP
*****************************************************************}
{
PT1-1 : 28/08/2001 SB V547 NBREMOIS positionner à 1, le 0 est interdit
PT1-2 : 28/08/2001 SB V547 Modification de l'affichage de l'arrondi dans la
                           grille
PT2-1 : 07/09/2001 SB V547 Validation  : Controle saisie du nombre de mois
                           Fermeture de la fiche que sur le bouton ferme et non
                           sur le bouton valider
PT- 3 : 19/09/2001 SB V560 Gestion de l'affichage des décimales
PT- 4 : 19/09/2001 SB V562 Edition de la reprise fiche de bug n°285
PT- 5 : 26/11/2001 SB V563 Fiche de bug n°362 : Violation d'acces
PT- 7 : 08/01/2002 SB V571 Rechargement des tobs après la validation..
PT- 8 : 30/04/2002 SB V571 InsertOrUpdateDb insert dans la table un
                           enregistrement à blanc
PT- 9 : 04/09/2002 SB V585 FQ 10002 Edition : les salariés sortis ne sont pas
                           repris
PT-10 : 16/12/2002 SB V591 FQ 10322 Violation d'acces en sauvegarde
PT11  : 12/03/2004 SB V_50 FQ 11162 Encodage de de la date de cloture erroné si
                           fin fevrier
PT12  : 15/03/2004 SB V_50 FQ 11159 Contrôle saisie négative sur reprise acquis,
                           pris CP
PT13-1: 07/04/2004 SB V_50 FQ 11212 En modification, les montants remis à zéro
                           ne sont pas sauvegardés
PT13-2: 07/04/2004 SB V_50 FQ 11212 Si remise à zéro du CPA sur N les consommés
                           sur REP ne sont pas réinitialisés
PT14  : 05/05/2004 SB V_50 FQ 11258 Refonte Sql Edition
PT15  : 18/10/2005 SB V_65 FQ 12499 Ajout Prénom + Date entrée
PT16  : 14/04/2006 SB V_65 FQ 12621 Controle mouvement CP limité au typemvt ='CPA'
PT17  : 19/05/2006 SB V_65 FQ 13155 Refonte de la grille pour intégration des REP d'acs, et d'aca
PT18-1  01/06/2006 SB V_65 FQ 13196 Contôle saisie jour acs, aca
PT18-2  01/06/2006 SB V_65 Refonte contôle + edition
PT18-3  01/06/2006 SB V_65 FQ 13162 Désactivation de l'impression en CWAS
PT19  : 08/10/2007 MF V_80 mise en place traitement des jours de fractionnement
PT20  : 18/10/2007 GGU V_80 FQ 14878 Gestion du copier coller depuis excel
PT21  : 27/12/2007 GGU V_80 FQ 14999 Si la saisie est réalisée à partir d'un liste
                            triée par ordre alphabétique, l'édition de la reprise est erronée
PT22  : 04/09/2008 NA  V_80 FQ 15410 Prendre en compte les acquis suppl ,les jours ancienneté et les jours ACF dans le total des acquis +
                            impute le nombre de jours consommés sur les lignes de reprise correspondantes (pour les périodes N-1 et N)
}
unit UTofRepriseCP;

interface
uses Controls, Classes,sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$IFDEF V530}
  EdtEtat,
{$ELSE}
  EdtREtat,
{$ENDIF}
{$ELSE}
  UtileAGL,
{$ENDIF}
  Grids, HCtrls, HEnt1,  HMsgBox, UTOF, UTOB, Vierge, Paramdat, ed_tools,ULibEditionPaie;

type
  TOF_RepriseCP = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
    procedure Sauvegarde(Salarie, typeconge, Typeimpute : string; N1: boolean;  { PT17 }
      par1, par2, par3, par4, par5: extended);
    procedure Effacement(Salarie, typeconge, Typeimpute: string; N1: Boolean);  { PT17 }
  private
    Etablissement, DateRepEnter, bull: string;
    NbDecimales, sauv, Increment: integer;
    QMul: TQUERY; // Query recuperee du mul
    Modifier, Change, Suppression, DejaCree, Validation: Boolean;
    DateCloture, DateclotureN1, DateReprise, DateClotureN2, DateDebutN1: TDateTime; // Date de debut exercice social
    LaGrille: THGrid; // Grilles de saisie des cumuls et des bases
    Tob_SalConf,LesCp,MereCP, tMaxNoOrdre      : TOB;    // TOB des salaries
    ValOnEnter : string; { PT18-2 }
    function OnSauve: boolean;
    function OkGrille: boolean;
    procedure ValiderClick(Sender: TObject);
    procedure DeleteClick(Sender: TObject);
    //       procedure FermeClick(Sender: TObject);  //PT- 2-1
    procedure GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);  { PT18-2 }
    procedure GrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GrilleCellKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);  //PT20
    procedure DaterepChange(Sender: TObject);
    procedure DateRepclick(Sender: TObject);
    procedure ImprimeEtat(Sender: TObject);
    procedure ChargementTob; //PT- 7
    Function  FormatLibelle(DateEntree : TDateTime; Nom, Prenom : String) : String;
    Function  RendIndexCol ( St : String) : Integer;  { PT17 }
  end;

implementation
uses
  P5Util, PgCalendrier, PgCommun
  , UTofPG_AcompteGrp //PT20
  ;


procedure TOF_RepriseCP.OnArgument(Arguments: string);
var
  F: TFVierge;
  Q: TQuery;
  st, Salarie, TypeMvt, libelle: string;
  i, Ligne, j, iBase, iMois: Integer;  { PT17 }
{$IFDEF EAGLCLIENT}
  t: Integer;
{$ENDIF}
  Abandon, OkOk: Boolean;
  TOBW: TOB;
  BtnValid, BtnDel: TToolbarButton97;
  aa, mm, jj: word;
  datevalidite: Tdatetime;
  NbJouTrav: double;
  DateRep: THEdit;
 DateEntree : String; { PT17 }
begin
  inherited;
  TFVierge(Ecran).BValider.ModalResult := 0;
  Suppression := true;
  Validation  := True; { PT18-2 }
  Tob_SalConf := nil;
  Sauv := 0;
  NbDecimales := 2;
  Etablissement := Trim(ReadTokenSt(Arguments));
  bull := Arguments;
  // recuperation de la query du multicritere
  if not (Ecran is TFVierge) then exit;
  F := TFVierge(Ecran);
  if F <> nil then QMUL := F.FMULQ;
  if QMUL = nil then exit;

  //Gestionnaire d'evenements
  BtnValid := TToolbarButton97(GetControl('BVALIDER'));
  if BtnValid <> nil then BtnValid.OnClick := ValiderClick;
  BtnDel := TToolbarButton97(GetControl('BDELETE'));
  if BtnDel <> nil then BtnDel.OnClick := DeleteClick;
  //BtnFerme:=TToolbarButton97 (GetControl ('BFERME'));
  //if BtnFerme<>NIL then BtnFerme.OnClick := FermeClick ;
  TFVierge(Ecran).BImprimer.OnClick := ImprimeEtat; //PT- 4

  // design des grille et definition des tailles des colonnes
  LaGrille := THGrid(Getcontrol('GRILLE'));
  if LaGrille <> nil then
  begin
{ DEB PT17 }
    LaGrille.ColNames[0]  := 'MATRICULE';
    LaGrille.ColNames[1]  := 'LIBELLE';
    LaGrille.ColNames[2]  := 'DATEENTREE';
    LaGrille.ColNames[3]  := 'BASEN1';
    LaGrille.ColNames[4]  := 'MOISN1';
    LaGrille.ColNames[5]  := 'ACQUISN1';
    LaGrille.ColNames[6]  := 'ACQUISACSN1';
    LaGrille.ColNames[7]  := 'ACQUISACAN1';
// d PT19
    LaGrille.ColNames[8]  := 'ACQUISACFN1'; 
{*    LaGrille.ColNames[8]  := 'PRISN1';
    LaGrille.ColNames[9]  := 'MONTANTN1';
    LaGrille.ColNames[10] := 'BASEN';
    LaGrille.ColNames[11] := 'MOISN';
    LaGrille.ColNames[12] := 'ACQUISN';
    LaGrille.ColNames[13] := 'ACQUISACSN';
    LaGrille.ColNames[14] := 'ACQUISACAN';
    LaGrille.ColNames[15] := 'PRISN';
    LaGrille.ColNames[16] := 'MONTANTN';*}
    LaGrille.ColNames[9]  := 'PRISN1';
    LaGrille.ColNames[10]  := 'MONTANTN1';
    LaGrille.ColNames[11] := 'BASEN';
    LaGrille.ColNames[12] := 'MOISN';
    LaGrille.ColNames[13] := 'ACQUISN';
    LaGrille.ColNames[14] := 'ACQUISACSN';
    LaGrille.ColNames[15] := 'ACQUISACAN';
    LaGrille.ColNames[16] := 'PRISN';
    LaGrille.ColNames[17] := 'MONTANTN';
// f PT19
    LaGrille.ColEditables[RendIndexCol('MATRICULE')] := False;     { DEB PT15 }
    LaGrille.ColEditables[RendIndexCol('LIBELLE')] := False;
    LaGrille.ColEditables[RendIndexCol('DATEENTREE')] := False;
{ FIN PT17 }
    for i := 0 to LaGrille.ColCount-1 do { PT17 }
      LaGrille.ColAligns[i] := taCenter;
    LaGrille.ColAligns[RendIndexCol('LIBELLE')] := taLeftJustify; { FIN PT15 } { PT17 }
  end
  else exit;

  LaGrille.OnCellExit := GrilleCellexit;
  LaGrille.OnCellEnter := GrilleCellenter; { PT18-2 }
  LaGrille.OnKeyDown := GrilleCellKeyDown; //PT20

  Q := opensql('SELECT ETB_DATECLOTURECPN,ETB_NBJOUTRAV,ETB_LIBELLE FROM ETABCOMPL WHERE ETB_ETABLISSEMENT = "' +
    etablissement + '"', TRUE);
  Datecloture := 0;
  NbJouTrav := 0;
  Libelle := '';
  if not Q.eof then
  begin
    DateCloture := Q.findfield('ETB_DATECLOTURECPN').AsVariant;
    NbJouTrav := Q.findfield('ETB_NBJOUTRAV').AsVariant;
    Libelle := Q.findfield('ETB_LIBELLE').AsVariant;
  end;
  Ferme(Q);
  if NbJoutrav = 5 then
    SetControlText('LNBJOUTRAV', 'Etablissement travaillant en jours ouvrés')
  else
    if NbJoutrav = 6 then
    SetControlText('LNBJOUTRAV', 'Etablissement travaillant en jours ouvrables')
  else
    SetControlText('LNBJOUTRAV', '');

  Ecran.Caption := Ecran.Caption + ' de l''établissement ' + etablissement + ' ' + libelle;
  UpdateCaption(TFVierge(Ecran));

  decodedate(DateCloture, aa, mm, jj);
  DateclotureN1 := PGEncodeDateBissextile(AA - 1, MM, JJ); { PT11 }
  DateclotureN2 := PGEncodeDateBissextile(AA - 2, MM, JJ); { PT11 }
  DateDebutN1 := DateClotureN2 + 1;

  ChargementTob; //PT- 7

  st := 'SELECT DISTINCT PCN_SALARIE,MAX(PCN_ORDRE) AS ORDRE ' +
    'FROM ABSENCESALARIE WHERE PCN_TYPEMVT="CPA" AND PCN_ETABLISSEMENT ="' + etablissement + '" ' +  { PT16 }
    'GROUP BY PCN_SALARIE ';
  Q := Opensql(st, True);
  tMaxNoOrdre := Tob.create('Les max ordre', nil, -1);
  tMaxNoOrdre.loaddetaildb('Les max ordre', '', '', Q, False, FALSE);
  Ferme(Q);

  InitMoveProgressForm(nil, 'Chargement des données de la saisie', 'Veuillez patienter SVP ...', 100, TRUE, TRUE);

{$IFNDEF EAGLCLIENT}
  QMul.First;
{$ENDIF}

  LaGrille.RowCount := 2;
  Ligne := 1; // Ligne courante
  Abandon := FALSE;

  SetControlText('TPERIODEN1', 'Période antérieure du ' + datetostr(dateclotureN2 + 1) + ' au ' + datetostr(dateclotureN1)+' (N1)'); { PT17 }
  SetControlText('TPERIODEN' , 'Période en cours du ' + datetostr(dateclotureN1 + 1) + ' au ' + datetostr(datecloture)+' (N)'); { PT17 }

{$IFNDEF EAGLCLIENT}
  while not QMul.EOF do
  begin
    Salarie := QMul.FindField('PSA_SALARIE').AsString;
    Libelle := FormatLibelle(QMul.FindField('PSA_DATEENTREE').AsDateTime,QMul.FindField('PSA_LIBELLE').AsString,QMul.FindField('PSA_PRENOM').AsString); { PT15 }
    DateEntree := DateToStr(QMul.FindField('PSA_DATEENTREE').AsDateTime); { PT17 }
{$ELSE}
  for t := 0 to QMul.Detail.Count - 1 do
  begin
    Salarie := QMul.Detail[t].GetValue('PSA_SALARIE');
    Libelle := FormatLibelle(QMul.Detail[t].GetValue('PSA_DATEENTREE'),QMul.Detail[t].GetValue('PSA_LIBELLE'),QMul.Detail[t].GetValue('PSA_PRENOM')); { PT15 }
    DateEntree := DateToStr(QMul.FindField('PSA_DATEENTREE').AsDateTime); { PT17 }
{$ENDIF}
    LaGrille.Cells[RendIndexCol('MATRICULE'), Ligne] := Salarie;  { PT17 }
    LaGrille.Cells[RendIndexCol('LIBELLE'), Ligne] := Libelle; { PT17 }
    LaGrille.Cells[RendIndexCol('DATEENTREE'), Ligne] := DateEntree;  { PT17 }

    for j := RendIndexCol('BASEN1') to LaGrille.ColCount - 1 do { PT17 }
    begin
      LaGrille.Cells[j, Ligne] := Doubletocell(0, 0);
    end;
    { Récupération des mvts REP }
    TobW := LesCp.findfirst(['PCN_SALARIE', 'PCN_TYPECONGE'], [Salarie, 'REP'], False);
    while TobW <> nil do
    begin { DEB PT17 }
//PT19      if (TobW.GetValue('PCN_TYPEIMPUTE') = 'ACA') OR (TobW.GetValue('PCN_TYPEIMPUTE') = 'ACS') then
      if (TobW.GetValue('PCN_TYPEIMPUTE') = 'ACA') OR (TobW.GetValue('PCN_TYPEIMPUTE') = 'ACS') OR (TobW.GetValue('PCN_TYPEIMPUTE') = 'ACF') then
         Begin
         TobW := LesCp.findnext(['PCN_SALARIE', 'PCN_TYPECONGE'], [Salarie, 'REP'], False);
         Continue;
         End;
      DateValidite := TobW.GetValue('PCN_DATEVALIDITE');
      if DATEVALIDITE <= DaTeClotureN1 then
        Begin
        i := RendIndexCol('ACQUISN1');
        ibase := RendIndexCol('BASEN1');
        iMois := RendIndexCol('MOISN1');
        End
        else
        Begin
        i := RendIndexCol('ACQUISN');
        ibase := RendIndexCol('BASEN');
        iMois := RendIndexCol('MOISN');
        End; { FIN PT17 }
      TypeMvt := TobW.getvalue('PCN_TYPEMVT');
      if Pos(',', FloatToStr(TobW.getvalue('PCN_JOURS'))) = 0 then //DEB1 PT- 3
        LaGrille.Cells[i, Ligne] := TobW.getvalue('PCN_JOURS')
      else //FIN1 PT- 3
        LaGrille.Cells[i, Ligne] := DoubleToCell(TobW.getvalue('PCN_JOURS'), NbDecimales);
      LaGrille.Cells[ibase, Ligne] := DoubleToCell(TobW.getvalue('PCN_BASE'), NbDecimales);  { PT17 }
      if Pos(',', FloatToStr(TobW.getvalue('PCN_NBREMOIS'))) = 0 then //DEB2 PT- 3
        LaGrille.Cells[iMois, Ligne] := TobW.getvalue('PCN_NBREMOIS')  { PT17 }
      else //FIN2 PT- 3
        LaGrille.Cells[iMois, Ligne] := DoubleToCell(TobW.getvalue('PCN_NBREMOIS'), NbDecimales);  { PT17 }
      TobW := LesCp.findnext(['PCN_SALARIE', 'PCN_TYPECONGE'], [Salarie, 'REP'], False);
    end;
    { Récupération des mvts CPA }
    TobW := LesCp.findfirst(['PCN_SALARIE', 'PCN_TYPECONGE'], [Salarie, 'CPA'], False);
    while TobW <> nil do
    begin
      DateValidite := TobW.GetValue('PCN_DATEVALIDITE');
      if DATEVALIDITE <= DaTeClotureN1 then
        Begin  { DEB PT17 }
        i := RendIndexCol('PRISN1');
        ibase := RendIndexCol('MONTANTN1');
        End
        else
        Begin
        i := RendIndexCol('PRISN');
        ibase := RendIndexCol('MONTANTN');
        End; { FIN PT17 }
      if Pos(',', FloatToStr(TobW.getvalue('PCN_JOURS'))) = 0 then //DEB3 PT- 3
        LaGrille.Cells[i, Ligne] := TobW.getvalue('PCN_JOURS')
      else //FIN3 PT- 3
        LaGrille.Cells[i, Ligne] := DoubleToCell(TobW.getvalue('PCN_JOURS'), Nbdecimales);
      LaGrille.Cells[ibase, Ligne] := DoubleToCell(TobW.getvalue('PCN_VALORETENUE'), Nbdecimales);  { PT17 }
      TobW := LesCp.findnext(['PCN_SALARIE', 'PCN_TYPECONGE'], [Salarie, 'CPA'], False);
    end;

    { PT17 Récupération des mvts ACA }
    TobW := LesCp.findfirst(['PCN_SALARIE', 'PCN_TYPECONGE','PCN_TYPEIMPUTE'], [Salarie, 'REP', 'ACA'], False);
    while TobW <> nil do
    begin
      DateValidite := TobW.GetValue('PCN_DATEVALIDITE');
      if DATEVALIDITE <= DaTeClotureN1 then   i := RendIndexCol('ACQUISACAN1')
      else                                    i := RendIndexCol('ACQUISACAN');
      if Pos(',', FloatToStr(TobW.getvalue('PCN_JOURS'))) = 0 then
        LaGrille.Cells[i, Ligne] := TobW.getvalue('PCN_JOURS')
      else
        LaGrille.Cells[i, Ligne] := DoubleToCell(TobW.getvalue('PCN_JOURS'), Nbdecimales);
      TobW := LesCp.findnext(['PCN_SALARIE', 'PCN_TYPECONGE','PCN_TYPEIMPUTE'], [Salarie, 'REP', 'ACA'], False);
    end;

    { PT17 Récupération des mvts ACS }
    TobW := LesCp.findfirst(['PCN_SALARIE', 'PCN_TYPECONGE','PCN_TYPEIMPUTE'], [Salarie, 'REP', 'ACS'], False);
    while TobW <> nil do
    begin
      DateValidite := TobW.GetValue('PCN_DATEVALIDITE');
      if DATEVALIDITE <= DaTeClotureN1 then   i := RendIndexCol('ACQUISACSN1')
      else                                    i := RendIndexCol('ACQUISACSN');
      if Pos(',', FloatToStr(TobW.getvalue('PCN_JOURS'))) = 0 then
        LaGrille.Cells[i, Ligne] := TobW.getvalue('PCN_JOURS')
      else
        LaGrille.Cells[i, Ligne] := DoubleToCell(TobW.getvalue('PCN_JOURS'), Nbdecimales);
      TobW := LesCp.findnext(['PCN_SALARIE', 'PCN_TYPECONGE','PCN_TYPEIMPUTE'], [Salarie, 'REP', 'ACS'], False);
    end;

// d PT19
    { Récupération des mvts ACF }
    TobW := LesCp.findfirst(['PCN_SALARIE', 'PCN_TYPECONGE','PCN_TYPEIMPUTE'], [Salarie, 'REP', 'ACF'], False);
    while TobW <> nil do
    begin
      DateValidite := TobW.GetValue('PCN_DATEVALIDITE');
      if DATEVALIDITE <= DaTeClotureN1 then   i := RendIndexCol('ACQUISACFN1');
//      else                                    i := RendIndexCol('ACQUISACSN');
      if Pos(',', FloatToStr(TobW.getvalue('PCN_JOURS'))) = 0 then
        LaGrille.Cells[i, Ligne] := TobW.getvalue('PCN_JOURS')
      else
        LaGrille.Cells[i, Ligne] := DoubleToCell(TobW.getvalue('PCN_JOURS'), Nbdecimales);
      TobW := LesCp.findnext(['PCN_SALARIE', 'PCN_TYPECONGE','PCN_TYPEIMPUTE'], [Salarie, 'REP', 'ACF'], False);
    end;
// f PT19

{$IFNDEF EAGLCLIENT}
    QMul.NEXT;
{$ENDIF}
    Ligne := Ligne + 1;
    okok := MoveCurProgressForm(IntToStr(Ligne));
    if not OkOk then
    begin
      Abandon := TRUE;
      break;
    end;
    LaGrille.RowCount := LaGrille.RowCount + 1;
  end;
  FiniMoveProgressForm();
  if Abandon = TRUE then Close; // Abandon de la saisie par rubrique
  LaGrille.MontreEdit;
  LaGrille.SynEnabled := True;
  LaGrille.RowCount := LaGrille.RowCount - 1;
  DateReprise := DateClotureN1 + 1;
  DateRep := THEdit(getcontrol('DATEREP'));
  if DateRep <> nil then
  begin
    DateRep.OnChange := DateRepChange;
    DateRep.OnelipsisClick := DateRepclick;
    DateRep.OnClick := DateRepclick;
    DateRep.Text := datetostr(dateClotureN1 + 1);
  end;

{$IFDEF EAGLCLIENT}
SetControlVisible('BIMPRIMER',False);  { PT18-3 }
{$ENDIF}
end;


procedure TOF_RepriseCP.DateRepChange(Sender: TObject);
var
  Daterep: THEdit;
begin
  if Change then
  begin
    Validation := False;
    DateRep := ThEdit(getcontrol('DATEREP'));
    if DateRep = nil then exit;
    if IsValiddate(DateRep.text) then
    begin
      DateReprise := strtodate(DateRep.Text);
      if ((DateReprise <= DateclotureN1) or (DateReprise > DateCloture)) then
      begin
        PgiBox('Attention, La date doit être dans comprise dans l''exercice congés payés courant', 'Reprise des mouvements CP');
        DateRep.text := datetostr(DateclotureN1 + 1);
        DateReprise := strtodate(DateRep.Text);
        setFocuscontrol('DATEREP');
        exit;
      end;
    end;
  end;
end;

procedure TOF_RepriseCP.DateRepclick(Sender: TObject);
var
  Daterep: THEdit;
  Init: word;
var
  key: char;
begin
  Validation := False;
  Change := true;
  DateRep := ThEdit(getcontrol('DATEREP'));
  if DateRep = nil then exit;
  DateRepEnter := DateRep.text;
  Init := HShowMessage('1;Date de Reprise CP;Cette date alimente la date de validité des mouvements en cours.#13#10 Voulez-vous continuer ?;Q;YN;N;N;', '', '');
  if Init = mrNo then
    if Lagrille <> nil then
      Lagrille.setFocus;
  if Init = MrYes then
  begin
    key := '*';
    ParamDate(Ecran, Sender, Key);
  end;
end;

function TOF_RepriseCP.OkGrille: boolean;
const
  Numerique = ['0'..'9'];
  Separateur = [',', '.'];
var
  r: integer;
  AcquisN1, AcquisN, PrisN1, PrisN: double;
begin
  result := true;
  for r := 1 to (Lagrille.RowCount - 1) do
  begin { DEB PT17 Refonte }
// pt22    if LaGrille.Cells[RendIndexCol('ACQUISN1'), r] = '' then AcquisN1 := 0
// pt22    else AcquisN1 := Valeur(LaGrille.Cells[RendIndexCol('ACQUISN1'), r]);
    if ((LaGrille.Cells[RendIndexCol('ACQUISN1'), r] = '') and (LaGrille.Cells[RendIndexCol('ACQUISACAN1'), r] = '')
      AND (LaGrille.Cells[RendIndexCol('ACQUISACSN1'), r] = '') AND (LaGrille.Cells[RendIndexCol('ACQUISACFN1'), r] = ''))
      then  AcquisN1 := 0       // pt22
      else AcquisN1 :=  Valeur(LaGrille.Cells[RendIndexCol('ACQUISN1'), r]) + Valeur(LaGrille.Cells[RendIndexCol('ACQUISACAN1'), r]) +
                        Valeur(LaGrille.Cells[RendIndexCol('ACQUISACSN1'), r]) + Valeur(LaGrille.Cells[RendIndexCol('ACQUISACFN1'), r]);             // pt22
    if LaGrille.Cells[RendIndexCol('PRISN1'), r] = '' then PrisN1 := 0
    else PrisN1 := Valeur(LaGrille.Cells[RendIndexCol('PRISN1'), r]);
// pt22    if LaGrille.Cells[RendIndexCol('ACQUISN'), r] = '' then AcquisN := 0
// pt22    else AcquisN := valeur(LaGrille.Cells[RendIndexCol('ACQUISN'), r]);
   if ((LaGrille.Cells[RendIndexCol('ACQUISN'), r] = '') And (LaGrille.Cells[RendIndexCol('ACQUISACAN'), r] = '')
      AND (LaGrille.Cells[RendIndexCol('ACQUISACSN'), r] = '')) then AcquisN := 0         // pt22
    else AcquisN := valeur(LaGrille.Cells[RendIndexCol('ACQUISN'), r]) + valeur(LaGrille.Cells[RendIndexCol('ACQUISACAN'), r])+
                   valeur(LaGrille.Cells[RendIndexCol('ACQUISACSN'), r]);                 // pt22
    if LaGrille.Cells[RendIndexCol('PRISN1'), r] = '' then PrisN := 0
    else PrisN := valeur(LaGrille.Cells[RendIndexCol('PRISN'), r]);
{ FIN PT17 }
    if ((PrisN1 > AcquisN1) or (PrisN > AcquisN))
      then
    begin
      LastError := 1; //LastErrorMsg:=''; PT-10 Utilisation d'un PgiBox au lieu du LastErrorMsg pour activer le message
      PgiBox('Le nombre de jours pris ne peut être supérieur au nombre de jours acquis ,#10#13 '+
      'contrôlez la saisie du salarié ' + LaGrille.Cells[RendIndexCol('MATRICULE'), r] + ' ' + LaGrille.Cells[RendIndexCol('LIBELLE'), r], Ecran.caption); { PT17 }
      result := false;
      exit;
    end;
  end;
end;

procedure TOF_RepriseCP.DeleteClick(Sender: TObject);
var
  st: string;
  r, i: integer;
begin

  if Bull = 'BULL' then
  begin
    PGIBox('Cette sélection intègre des salariés ayant des bulletins calculés;Suppression impossible', 'Suppression reprise CP');
    exit;
  end;
//  Init := HShowMessage('1;Suppression de tous les Mouvements de Reprise;Vous allez supprimer tous les mouvements de reprise CP de cette sélection.#13#10 Voulez-vous continuer ?;Q;YN;N;N;', '', '');
  if PgiAsk('Confirmez-vous la suppression des mouvements de reprise saisis?',Ecran.caption) = mrNo then { PT17 Refonte }
    exit
  else
  begin
    try
      Suppression := true;
      BeginTrans;
      //     exit;
      for r := 1 to (Lagrille.RowCount - 1) do
      begin
        for i := RendIndexCol('BASEN1') to LaGrille.ColCount - 1  do  { PT17 }
        begin
          Lagrille.Cells[i, r] := '';
        end;
        st := 'DELETE FROM ABSENCESALARIE';
        st := st + ' WHERE (PCN_TYPECONGE = "CPA" OR PCN_TYPECONGE = "REP") AND PCN_DATEVALIDITE <= "' + usdatetime(DaTeCloture) + '"';
        st := st + ' AND PCN_DATEVALIDITE > "' + usdatetime(DaTeClotureN2) + '" and PCN_ETABLISSEMENT = "' + etablissement + '"';
        st := st + ' AND PCN_SALARIE ="' + Lagrille.Cells[RendIndexCol('MATRICULE'), r] + '"'; { PT17 }
        executeSQL(st);
      end;
      //Lescp.Free; PT- 5 On détruit la mère sur le OnClose
      Suppression := false;
      ValiderClick(Sender);
      //     Ecran.Close;
           //MereCp.free; PT- 5 Réporté sur le OnClose
      CommitTrans;
    except
      Rollback;
    end;
  end;
end;

// Fonction de validation de la saisie

procedure TOF_RepriseCP.ValiderClick(Sender: TObject);
var
  Grille: THGrid;
  ACol, ARow: Integer;
  Cancel: Boolean;
begin
  inherited;
  // DEB2 PT- 2-1
  Grille := THGrid(GetControl('GRILLE'));
  if Grille <> nil then
  begin
    Acol := Grille.Col;
    ARow := Grille.Row;
    Cancel := False;
    GrilleCellExit(Grille, ACol, ARow, Cancel);
  end;

  if LastError = 0 then
  begin
    if not Suppression then exit;
    if okgrille then
      OnSauve
    else
      Validation := False;
  end;
end;
// FIN2 PT- 2-1

// fermeture de la forme
//procedure TOF_RepriseCP.FermeClick(Sender: TObject);
//begin
// ValiderClick(Sender);
//end;

// fonction d'ecriture des elements saisis

function TOF_RepriseCP.OnSauve: boolean;
var
  r, i: Integer;
  Salarie, st: string;
  Erreur     : boolean;
  Q: Tquery;
//PT19  Acq, Aca, Acs, Base, Mois, Pris, Mont : Double; { PT17 }
  Acq, Aca, Acs, Acf, Base, Mois, Pris, Mont : Double; { PT17 }
  JconsAc, Jconsaca, jconsacs,jconsacf,reliquat : Double; // pt22
begin
  result := TRUE;
  if Modifier = FALSE then
    exit;
  Modifier := FALSE;
  Erreur := False;
  for r := 1 to (Lagrille.RowCount - 1) do
  begin
    if (Lagrille.Cells[RendIndexCol('ACQUISN1'), r] <> '') and (Lagrille.Cells[RendIndexCol('MOISN1'), r] = '') then { PT17 }
      Erreur := True;
    if (Lagrille.Cells[RendIndexCol('ACQUISN'), r] <> '') and (Lagrille.Cells[RendIndexCol('MOISN'), r] = '') then { PT17 }
      Erreur := True;
    if Erreur then
    begin
      PGIBox('Vous devez renseigner le nombre de mois du salarié "' +
        Lagrille.Cells[RendIndexCol('MATRICULE'), r] + '" !', Ecran.caption); { PT17 }
      LastError := 1;
      Exit;
    end;
  end;

  // on cherche dans la tob un enregistrement qui correspond à la ligne
  // si on le trouve, on le maj sinon on le cree
  st := 'SELECT PSA_SALARIE, PSA_TRAVAILN1, PSA_TRAVAILN2, PSA_TRAVAILN3,' +
    'PSA_TRAVAILN4, PSA_CODESTAT, PSA_CONFIDENTIEL' +
    ' FROM SALARIES ';
  Q := opensql(st, true);
  Tob_SalConf := tob.create('Table des salaries', nil, -1);
  Tob_SalConf.loaddetaildb('Table des salaries', '', '', Q, false);
  Ferme(Q);
  DejaCree := False;
  Increment := 0;
  //Rep:=PGiAsk('Voulez-vous générer un cumul de reprise?','');
  for r := 1 to (Lagrille.RowCount - 1) do
  begin
    // traitement N-1
    Salarie := Lagrille.Cells[0, r];
	{ DEB PT17 }
    Acq  := Valeur(Lagrille.Cells[RendIndexCol('ACQUISN1'), r]);
    Aca  := Valeur(Lagrille.Cells[RendIndexCol('ACQUISACAN1'), r]);
    Acs  := Valeur(Lagrille.Cells[RendIndexCol('ACQUISACSN1'), r]);
    Acf  := Valeur(Lagrille.Cells[RendIndexCol('ACQUISACFN1'), r]); //PT19
    Base := Valeur(Lagrille.Cells[RendIndexCol('BASEN1'), r]);
    Mois := Valeur(Lagrille.Cells[RendIndexCol('MOISN1'), r]);
    Pris := Valeur(Lagrille.Cells[RendIndexCol('PRISN1'), r]);
    Mont := Valeur(Lagrille.Cells[RendIndexCol('MONTANTN1'), r]);
    // deb pt22
   // Si nb jours pris <= nb jours acquis N_1 : Nb jours consommés = nb jours pris sinon nb jours consommés = nb jours acquis
    jconsac := 0;
    if pris <= acq then Jconsac := pris else jconsac := acq;  // pt22
    //  if (Acq <> 0 ) then Sauvegarde(Salarie, 'REP','AC2', True, Acq, Base, Mois, Pris, 0)
    if (Acq <> 0 ) then Sauvegarde(Salarie, 'REP','AC2', True, Acq, Base, Mois, jconsac, 0)
    // fin pt22
    else                Effacement(Salarie, 'REP','AC2', True);
    if (Pris <> 0) then Sauvegarde(Salarie, 'CPA','AC2', True, Pris, 0, 0, 0,Mont)
    else                Effacement(Salarie, 'CPA','AC2', True);
   // deb pt22
   // si nb jours pris - nb jours acquis > nb jours acquis ancienneté : nb jours consommés = nb jours acquis ancienneté
   // si nb jours pris - nb jours acquis > 0 et < nb jours acquis ancienneté : nb jours consommés = nb jours pris - nb jours acquis
    jconsaca :=0;
    reliquat := pris - jconsac;
    if reliquat > aca then jconsaca := aca else if ((reliquat > 0) and ( reliquat <= aca)) then jconsaca := reliquat;
    //if (Aca <> 0) then  Sauvegarde(Salarie, 'REP','ACA', True, Aca, 0, 0, 0, 0)
    if (Aca <> 0) then  Sauvegarde(Salarie, 'REP','ACA', True, Aca, 0, 0, jconsaca, 0)
    // fin pt22
    else                Effacement(Salarie, 'REP','ACA', True);
    // deb pt22
    // si nb jours pris - nb jours acquis - nb jours acquis ancienneté > 0 : nb jours consommés = nb jours pris - nb jours acquis - nb jours acquis anc
    jconsacs := 0;
    reliquat := pris - jconsac - jconsaca;
    if reliquat > acs then jconsacs := acs else if ((reliquat > 0) and (reliquat <= acs)) then jconsacs := reliquat;
    //if (Acs <> 0) then  Sauvegarde(Salarie, 'REP','ACS', True, Acs, 0, 0, 0, 0)
    if (Acs <> 0) then  Sauvegarde(Salarie, 'REP','ACS', True, Acs, 0, 0, jconsacs, 0)
    // fin pt22
    else                Effacement(Salarie, 'REP','ACS', True);
    // deb pt22  même principe avec les jours de fractionnnement
    jconsacf := 0;
    jconsacf := pris - jconsac - jconsaca - jconsacs;
    // d PT19
    //if (Acf <> 0) then  Sauvegarde(Salarie, 'REP','ACF', True, Acf, 0, 0, 0, 0)
    if (Acf <> 0) then  Sauvegarde(Salarie, 'REP','ACF', True, Acf, 0, 0, jconsacf, 0)
    // fin pt22
    else                Effacement(Salarie, 'REP','ACF', True);
// f PT19

  // pt22  même principe que pour N-1
    Acq  := Valeur(Lagrille.Cells[RendIndexCol('ACQUISN'), r]);
    Aca  := Valeur(Lagrille.Cells[RendIndexCol('ACQUISACAN'), r]);
    Acs  := Valeur(Lagrille.Cells[RendIndexCol('ACQUISACSN'), r]);
    Base := Valeur(Lagrille.Cells[RendIndexCol('BASEN'), r]);
    Mois := Valeur(Lagrille.Cells[RendIndexCol('MOISN'), r]);
    Pris := Valeur(Lagrille.Cells[RendIndexCol('PRISN'), r]);
    Mont := Valeur(Lagrille.Cells[RendIndexCol('MONTANTN'), r]);
    // deb pt22
    jconsac := 0;
    if pris <= acq then Jconsac := pris else jconsac := acq;  // pt22
    // if (Acq <> 0 ) then Sauvegarde(Salarie, 'REP','AC2', false, Acq, Base, Mois, Pris, 0)
    if (Acq <> 0 ) then Sauvegarde(Salarie, 'REP','AC2', false, Acq, Base, Mois, Jconsac, 0)
    // fin pt22
    else                Effacement(Salarie, 'REP','AC2', false);
    if Pris <> 0 then   Sauvegarde(Salarie, 'CPA','AC2', false, Pris, 0, 0, 0,Mont)
    else                Effacement(Salarie, 'CPA','AC2', false);
    // deb pt22
    jconsaca :=0;
    reliquat := pris - jconsac;
    if reliquat > aca then jconsaca := aca else if ((reliquat > 0) and ( reliquat <= aca)) then jconsaca := reliquat;
    //  if (Aca <> 0) then  Sauvegarde(Salarie, 'REP','ACA', false, Aca, 0, 0, 0, 0)
    if (Aca <> 0) then  Sauvegarde(Salarie, 'REP','ACA', false, Aca, 0, 0, jconsaca, 0)
    // fin pt22
    else                Effacement(Salarie, 'REP','ACA', false);
    // deb pt22
    jconsacs := 0;
    jconsacs := pris - jconsac - jconsaca;
    //if (Acs <> 0) then  Sauvegarde(Salarie, 'REP','ACS', false, Acs, 0, 0, 0, 0)
    if (Acs <> 0) then  Sauvegarde(Salarie, 'REP','ACS', false, Acs, 0, 0, jconsacs, 0)
    // fin pt22
    else                Effacement(Salarie, 'REP','ACS', false);
	{ FIN PT17 }

  end; // end for
  // if Tob_SalConf <> nil then    //PT- 5 Reporté sur le OnClose
  //     begin Tob_SalConf.free; Tob_SalConf:=nil end;
  try
    BeginTrans;
    //DEB PT- 8  Au lieu du Merecp.InsertOrUpdateDB qui insère un enr à blanc
    for i := 0 to LesCp.Detail.count - 1 do
    begin
      if LesCp.detail[i].Modifie then { PT17 }
      begin
        LesCp.detail[i].SetAllModifie(TRUE);
        LesCp.detail[i].InsertOrUpdateDB;
      end; //FIN  PT- 8 on fait un insertorupdate sur chaque fille modifié..Traitement allegé
    end;
    //Merecp.SetAllModifie (TRUE);  PT- 8  Mise en Commentaire
    //Merecp.InsertOrUpdateDB (TRUE); PT- 8
    //MereCp.free; PT- 5 Réporté sur le OnClose
    CommitTrans;
    Validation := True; // PT- 2-1
  except
    result := FALSE;
    Rollback;
    PGIBox('Une erreur est survenue lors de la validation de la saisie',
      'reprise des CP');
    Validation := False; // PT- 2-1
  end;
  Modifier := FALSE;
  ChargementTob; //PT- 7
end;

{ DEB PT18-2 }
procedure TOF_RepriseCP.GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
Begin
if (Acol=RendIndexCol('MATRICULE')) or (ACol=RendIndexCol('LIBELLE')) or (ACol=RendIndexCol('DATEENTREE')) then exit;   { PT15 } { PT17 }
ValOnEnter := LaGrille.cells[Acol, Arow];
End;
{ FIN PT18-2 }
procedure TOF_RepriseCP.GrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  Nbre: Double;
begin
  if (Acol=RendIndexCol('MATRICULE')) or (ACol=RendIndexCol('LIBELLE')) or (ACol=RendIndexCol('DATEENTREE')) then exit;   { PT15 } { PT17 }
  LastError := 0;
  Modifier := TRUE;
  if (Isnumeric(LaGrille.cells[Acol, Arow])) then //DEB4 PT- 3
  begin
    Nbre := Valeur(LaGrille.Cells[Acol, ARow]);
    if (Nbre <> 0) then
    begin
      if ((ACol = RendIndexCol('BASEN1')) or (ACol = RendIndexCol('MONTANTN1')) { PT17 }
        or (ACol = RendIndexCol('BASEN')) or (ACol = RendIndexCol('MONTANTN')) { PT17 }
        or (Pos('.', LaGrille.cells[Acol, Arow]) > 0)
        or (Pos(',', LaGrille.cells[Acol, Arow]) > 0)) then
        LaGrille.Cells[Acol, Arow] := DoubleToCell(Nbre, NbDecimales);
    end
    else
      LaGrille.cells[Acol, Arow] := '';
  end
  else
    LaGrille.cells[Acol, Arow] := '';
  { PT- 3 Mise en commentaire puis supprimé }

//test nbre de mois supérieur à 0  DEB PT- 1-1
  if ((Acol = RendIndexCol('MOISN1')) or (Acol = RendIndexCol('MOISN'))) and ((LaGrille.Cells[Acol, ARow] <> '')) then { PT17 }
    if Valeur(LaGrille.Cells[Acol, ARow]) <= 0 then
    begin
      LastError := 1; // PT- 2-1
      PGIBox('Vous devez saisir un nombre de mois supérieur à zéro.', 'Nombre de mois :');
      LaGrille.Cells[Acol, ARow] := '';
    end;
  // FIN PT- 1-1
  { DEB PT12 }
  { DEB PT18-1 }
  if ( (Acol = RendIndexCol('ACQUISN1')) or (Acol = RendIndexCol('ACQUISN'))
  or   (Acol = RendIndexCol('ACQUISACSN1')) or (Acol = RendIndexCol('ACQUISACAN1'))
  or   (Acol = RendIndexCol('ACQUISACFN1')) //PT19
  or   (Acol = RendIndexCol('ACQUISACSN')) or (Acol = RendIndexCol('ACQUISACAN'))
  or (Acol = RendIndexCol('PRISN1')) or  (Acol = RendIndexCol('PRISN')) )
  and ((LaGrille.Cells[Acol, ARow] <> '') {or (nbre < 0)}) then
   if Valeur(LaGrille.Cells[Acol, ARow]) < 0 then
    begin
      LastError := 1; // PT- 2-1
      PGIBox('Vous devez saisir un nombre de jour supérieur à zéro.', Ecran.Caption);
      LaGrille.Cells[Acol, ARow] := '';
    end;
  { FIN PT18-1 }
  { FIN PT12 }

  if ValOnEnter <> LaGrille.cells[Acol, Arow] then Validation := False; { PT18-2 Tjrs en fin de procedure } 
end;

//PT20
procedure TOF_RepriseCP.GrilleCellKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = 86) and (ssCtrl in Shift) then
    GrilleCopierColler(LaGrille,3);
end;

procedure TOF_RepriseCP.OnClose;
begin
  if not Suppression then exit;
  if Validation = False then // DEB3 PT- 2-1
  begin
    Sauv := PGIAsk('Voulez-vous sauvegarder votre saisie ?', Ecran.Caption);
    if Sauv = mrNo then exit;
    if Sauv = mrYes then
      ValiderClick(nil);
  end;
  // if sauv = mryes then   mise en commentaire
    //   okgrille;               // DEB3 PT- 2-1
  //DEB PT- 5
  FreeAndNil(Tob_SalConf);
  FreeAndNil(MereCp);
  //FIN PT-5
  FreeAndNil(tMaxNoOrdre);
end;


procedure TOF_RepriseCP.Sauvegarde(Salarie, typeconge, Typeimpute: string; N1: boolean; { PT17 }
  par1, par2, par3, par4, par5: extended);
var
  trouve: boolean;
  DateValidite: tdatetime;
TobW,TN,Tmsal,T : tob;
  Ordre: integer;
begin
  trouve := False;
  Tmsal := Tob_SalConf.findfirst(['PSA_SALARIE'], [Salarie], false);
  if Tmsal = nil then
    exit;
  TobW := LesCp.findfirst(['PCN_SALARIE', 'PCN_TYPECONGE','PCN_TYPEIMPUTE'], { PT17 }
    [Salarie, typeconge,Typeimpute], False);
  while TobW <> nil do
  begin
    DateValidite := TobW.GetValue('PCN_DATEVALIDITE');
    if N1 then
    begin
      if DateValidite <= dateclotureN1 then
      begin
        trouve := true;
        break;
      end;
    end
    else
    begin
      if DateValidite > dateclotureN1 then
      begin
        trouve := true;
        break;
      end;
    end;
    TobW := LesCp.findnext(['PCN_SALARIE', 'PCN_TYPECONGE','PCN_TYPEIMPUTE'], { PT17 }
    [Salarie, typeconge,Typeimpute], False);
    end;
    if trouve then // mise à jour de l'enregistrement
    begin
      TobW.Modifie := True; { PT17 }
      TobW.PutValue('PCN_JOURS', par1);
      TobW.PutValue('PCN_CONFIDENTIEL', Tmsal.getvalue('PSA_CONFIDENTIEL'));
      TobW.PutValue('PCN_DATEPAIEMENT', idate1900); { PT13-2 }
      if typeconge = 'CPA' then //reprise pris
      begin
        if par4 > 0 then
          TobW.PutValue('PCN_DATEPAIEMENT', TobW.getvalue('PCN_DATEVALIDITE'));
        TobW.PutValue('PCN_VALORETENUE', par5); { PT13-1 Supérieur ou non à 0 }
        TobW.PutValue('PCN_BASE', par5); { PT13-1 Supérieur ou non à 0 }
        TobW.putValue('PCN_CODETAPE', 'P');
        TobW.PutValue('PCN_DATEDEBUT', TobW.getvalue('PCN_DATEVALIDITE'));
        TobW.PutValue('PCN_DATEFIN', TobW.getvalue('PCN_DATEVALIDITE'));
      end;
      if typeconge = 'REP' then //reprise acquis
      begin
        TobW.PutValue('PCN_BASE', par2);
        TobW.PutValue('PCN_NBREMOIS', par3);
        if N1 then
          TobW.PutValue('PCN_PERIODECP', 1)
        else
          TobW.PutValue('PCN_PERIODECP', 0);
        TobW.PutValue('PCN_APAYES', par4); { PT13-2 }
      end;
    end
    else // création de l'enregistrement
    begin
      TN := TOB.Create('ABSENCESALARIE', LesCP, -1); //MereCP
      InitialiseTobAbsenceSalarie(TN);
      TN.Modifie := True; { PT17 }
      TN.Putvalue('PCN_TRAVAILN1', tmsal.getvalue('PSA_TRAVAILN1'));
      TN.Putvalue('PCN_TRAVAILN2', tmsal.getvalue('PSA_TRAVAILN2'));
      TN.Putvalue('PCN_TRAVAILN3', tmsal.getvalue('PSA_TRAVAILN3'));
      TN.Putvalue('PCN_TRAVAILN4', tmsal.getvalue('PSA_TRAVAILN4'));
      TN.Putvalue('PCN_CODESTAT', tmsal.getvalue('PSA_CODESTAT'));
      TN.PutValue('PCN_CONFIDENTIEL', tmsal.getvalue('PSA_CONFIDENTIEL'));
      TN.PutValue('PCN_SALARIE', Salarie);
      TN.PutValue('PCN_JOURS', par1);
      TN.PutValue('PCN_TYPEIMPUTE' , Typeimpute); { PT17 }
      if N1 then
        TN.PutValue('PCN_PERIODECP', 1)
      else
        TN.PutValue('PCN_PERIODECP', 0);

      T := tMaxNoOrdre.findfirst(['PCN_SALARIE'], [Salarie], False);
      if (T <> nil) then
        ordre := T.getvalue('ORDRE')
      else
      begin
        T := Tob.create('La fille des max ordre', tMaxNoOrdre, -1);
        T.AddChampSup('PCN_SALARIE', False);
        T.AddChampSup('ORDRE', False);
        T.PutValue('PCN_SALARIE', Salarie);
        ordre := 0;
      end;
      //  if DejaCree=true then increment:=increment+1;  increment
      ordre := ordre + 1;
      T.putvalue('ORDRE', ordre);
    if N1 then
       TN.PutValue('PCN_DATEVALIDITE',dateDebutN1)
    else
       TN.PutValue('PCN_DATEVALIDITE',dateReprise);
      TN.PutValue('PCN_ORDRE', ordre);
      TN.PutValue('PCN_TYPECONGE', typeconge);
      TN.PutValue('PCN_ETABLISSEMENT', etablissement);
      if typeconge = 'REP' then //reprise acquis
      begin
        TN.PutValue('PCN_BASE', par2);
        TN.PutValue('PCN_NBREMOIS', par3);
        TN.PutValue('PCN_SENSABS', '+');
        TN.PutValue('PCN_LIBELLE', 'Reprise acquis ' + floattostr(par1) + 'j');
        TN.PutValue('PCN_APAYES', par4); { PT13-2 }
        if par4 > 0 then
          TN.PutValue('PCN_DATEPAIEMENT', TN.getvalue('PCN_DATEVALIDITE'))
        else
          TN.PutValue('PCN_DATEPAIEMENT', idate1900); { PT13-2 }
      end;
      if typeconge = 'CPA' then // reprise pris
      begin
        TN.PutValue('PCN_SENSABS', '-');
        TN.PutValue('PCN_LIBELLE', 'Reprise pris ' + floattostr(par1) + 'j');
        TN.PutValue('PCN_DATEDEBUT', TN.getvalue('PCN_DATEVALIDITE'));
        TN.PutValue('PCN_DATEPAIEMENT', TN.getvalue('PCN_DATEVALIDITE'));
        TN.PutValue('PCN_DATEFIN', TN.getvalue('PCN_DATEVALIDITE'));
        TN.putValue('PCN_CODETAPE', 'P');
        if par5 > 0 then
          TN.PutValue('PCN_VALORETENUE', par5);
        TN.PutValue('PCN_BASE', par5);
      end;
			{ DEB PT17 }
//PT19     if (Typeimpute = 'ACS') OR (Typeimpute = 'ACA') then
     if (Typeimpute = 'ACS') OR (Typeimpute = 'ACA') OR (Typeimpute = 'ACF')then
       Begin
       if (Typeimpute = 'ACS') then
       TN.PutValue('PCN_LIBELLE', 'Reprise acquis suppl. au ' + datetostr(TN.getvalue('PCN_DATEVALIDITE')))
       else
       if (Typeimpute = 'ACA') then
       TN.PutValue('PCN_LIBELLE', 'Reprise acquis anc. au ' + datetostr(TN.getvalue('PCN_DATEVALIDITE')))
// d PT19
       else
       if (Typeimpute = 'ACF') then
       TN.PutValue('PCN_LIBELLE', 'Reprise acquis Fract. au ' + datetostr(TN.getvalue('PCN_DATEVALIDITE')));
// f PT19       
       End; { FIN PT17 }
    end;
  end;

procedure TOF_RepriseCP.Effacement(Salarie, typeconge, Typeimpute : string; N1: Boolean); { PT17 }
var
  trouve: boolean;
  DateValidite: tdatetime;
  TobW: tob;
begin
  DateValidite := idate1900;
  Trouve := False;
  TobW := LesCp.findfirst(['PCN_SALARIE', 'PCN_TYPECONGE', 'PCN_TYPEIMPUTE'], [Salarie, typeconge, Typeimpute], False); { PT17 }
  while TobW <> nil do
  begin
    DateValidite := TobW.GetValue('PCN_DATEVALIDITE');
    if N1 then
    begin
      if DateValidite <= dateclotureN1 then
      begin
        trouve := true;
        break;
      end;
    end
    else
    begin
      if DateValidite > dateclotureN1 then
      begin
        trouve := true;
        break;
      end;
    end;
    TobW := LesCp.findnext(['PCN_SALARIE', 'PCN_TYPECONGE', 'PCN_TYPEIMPUTE'], [Salarie, typeconge, Typeimpute], False); { PT17 }
  end;
  // Suppresion de l'enregistrement
  if trouve then
  begin
    ExecuteSql('DELETE FROM ABSENCESALARIE WHERE PCN_SALARIE="' + Salarie + '" ' +
      'AND PCN_TYPEMVT="CPA" AND PCN_TYPECONGE="' + typeconge + '"' +
      'AND PCN_TYPEIMPUTE="'+Typeimpute+'" '+  { PT17 }
      'AND PCN_DATEVALIDITE="' + UsDateTime(DateValidite) + '" ');
    TobW.Free;
  end;
end;

{DEB PT- 4}

procedure TOF_RepriseCP.ImprimeEtat(Sender: TObject);
var
  Pages: TPageControl;
  StSql: string;
  i: Integer;

  BoCancel : Boolean;//PT21
  j : Integer;

begin
  //L'édition se sert des indices de la grille et du résultat de la requête triée par matricule salarié.
  //Pour que les données correspondent, il faut trier la grille par matricule salarié.
  LaGrille.SortGrid(0, False); //PT21
  //On remplace les espaces insécables par des espaces normaux, car lors de l'édition, les espaces insécables sont
  //transformés par 'Â '...
  for i := 0 to LaGrille.RowCount -1 do
  begin
    for j := 0 to LaGrille.ColCount -1 do
    begin
      LaGrille.CellValues[j, i] := StringReplace(LaGrille.CellValues[j, i], ' '{Attention, ce caractère n'est pas un espace "normal", mais un espace insécable}, ' ', [rfReplaceAll, rfIgnoreCase	]);
    end;
  end;
  { DEB PT14 }
  if Lagrille.RowCount = 1 then
  begin
    PgiBox('Aucun enregistrement à imprimer.', Ecran.caption);
    exit;
  end;
  Pages := TPageControl(GetControl('PAGES')); { PT14 18/05/2004 }
  if Pages = nil then exit;
  StSql := 'SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_DATEENTREE'+ { PT18-2 }
           ' FROM SALARIES WHERE PSA_SALARIE IN (';
  for i := 1 to Lagrille.RowCount - 1 do
    if LaGrille.Cells[RendIndexCol('MATRICULE'), i] <> '' then  { PT17 }
      StSql := StSql + '"' + LaGrille.Cells[RendIndexCol('MATRICULE'), i] + '",'; { PT17 } { PT18-2 }
  StSql := Copy(StSql, 1, Length(StSql) - 1) + ') ORDER BY PSA_SALARIE';
  { FIN PT14 }
  LanceEtat('E', 'PCG', 'PRE', True, False, False, pages, StSql, '', False);

//  for i := 0 to LaGrille.RowCount -1 do
//  begin
//    acol := 3;
//    arow := i;
//    LaGrille.OnCellExit(self, acol, arow, BoCancel);
//  end;

end;
{FIN PT- 4}

procedure TOF_RepriseCP.ChargementTob;
var
  Q: Tquery;
  St: string;
begin
  FreeAndNil(MereCp);
  MereCp := TOB.Create('Mere des cp', nil, -1);
  LesCP := TOB.Create('ABSENCESALARIE', MereCP, -1);
  st := 'SELECT * ' +
    'FROM ABSENCESALARIE ' +
    'WHERE PCN_TYPEMVT="CPA" AND (PCN_TYPECONGE = "CPA" OR PCN_TYPECONGE = "REP") ' +  { PT16 }
    'AND PCN_DATEVALIDITE <= "' + usdatetime(DaTeCloture) + '"  ' +
    'AND PCN_DATEVALIDITE > "' + usdatetime(DaTeClotureN2) + '" ' +
    'AND PCN_ETABLISSEMENT = "' + etablissement + '" ' +
    'ORDER BY PCN_SALARIE';
  Q := OpenSql(st, TRUE);
  if not Q.EOF then //PT-10 La tob LesCP était chargé avec aucun enr. d'où la violation d'accès
    LesCP.LoadDetailDB('ABSENCESALARIE', '', '', Q, FALSE, FALSE);
  Ferme(Q);
end;

{ DEB PT15 }
Function TOF_RepriseCP.FormatLibelle(DateEntree : TDateTime; Nom, Prenom : String) : String;
Begin
{ PT17 Suppression de la date d'entrée }
Result := UpperCase(Copy(Nom,1,1))+LowerCase(Copy(Nom,2,Length(Nom)))+' '+
          UpperCase(Copy(Prenom,1,1))+LowerCase(Copy(Prenom,2,Length(Prenom)));
//DateToStr(DateEntree)+' - '+
End;
{ FIN PT15 }

{ DEB PT17 }
Function  TOF_RepriseCP.RendIndexCol ( St : String) : Integer;
Var i : integer;
Begin
Result := -1;
For i := 0 to LaGrille.ColCount-1 do
  if LaGrille.ColNames [i] = UpperCase(St) then
    Begin
    result := i;
    Break;
    End;
if Result = -1 then PGibox('Nom de colonne incorrect : ' +St, Ecran.caption);    
End;
{ FIN PT17 }

initialization
  registerclasses([TOF_REPRISECP]);
end.

