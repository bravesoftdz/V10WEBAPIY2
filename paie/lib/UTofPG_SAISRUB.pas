{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 04/09/2001
Modifié le ... :   /  /
Description .. : Unit de gestion de la saisie par rubrique
Mots clefs ... : PAIE;SAISRUB
*****************************************************************}
{
PT1  : 28/02/2002 PH V572 Refonte complète du source pour la mise en place
      de la Multi selection sur les masques associés à l'utilisateur
      On passe de 7 colonnes à 28 (35 au cas où) et on organise la grille en fonction du nbre
      de colonnes gérées dans chaque masque
PT2  : 03/06/2002 V582 PH Gestion historique des évènements
PT3  : 28/10/2002 V582 PH Sauvegarde de la saisie car on a récupéré des infos du mois precedent donc il faut les valider
PT4  : 29/04/2003 V_42 SB FQ 10645 Export sous excel de la grille
PT5  : 23/09/2003 PH V_421 Passage de la clause where de la query en parametre pour ne traiter que les salariés séléctionnés
PT6  : 07/01/2004 PH V_50 Prise en compte des rémunérations de type elt permanent dans la saisie par rubrique
PT7  : 01/06/2004 PH V_50 Design de la grille géré par le HMTRAD
PT8  : 08/06/2004 PH V_50 Gestion d'un libellé associé à chaque cellule
PT9  : 08/06/2004 PH V_50 FQ 10359 Stockage établissement dans la table
PT10 : 23/08/2004 PH V_50 FQ 11428 Titre écran
PT11 : 03/08/2005 PH V_60 FQ 12299 Colonne général sur la colonne mais personnalisable + idem mais avec la valeur de la cellule
PT12 : 03/08/2005 PH V_60 FQ 12306 Accèder aux absences du salarié
PT13 : 13/04/2006 PH V_65 FQ 12490 Initialisation des dates avec iDate1900
PT14 : 10/01/2007 GGU V_70 FQ 13404 Saisie des bulletins complémentaires par rubrique
PT15 : 16/01/0207 GGU V_70 FQ 13404 Unification de la fonction de copier coller depuis Excel avec
       celle de l'unitée UTofPG_AcompteGrp
PT16 : 23/03/2007 FC V_70 FQ 13666 Faire apparaitre les CP dans les absences - modification des paramètres
       passés à ABSENCE_MUL
PT17 : 11/04/2007 GGU V7.02 FQ 14090 Dans l'écran de saisie par rubrique, ajouter en bas d'écran la totalisation des colonnes.
PT18 : 24/07/2007 FC V_72 FQ 14540 ergonomie
PT19 : 14/02/2008 GGU Message de confirmation dans Excel2007 lors de l'export -> Changement de format d'export
PT20 : 18/08/2008 JPP V_80 FQ15439 Désactivation puis réactivation du copier-coller Windows
}
unit UTofPG_SAISRUB;

interface
uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  Windows, StdCtrls, Controls, Classes, Graphics, forms, sysutils, Comctrls, Dialogs,
  HXlsPas,
  HTB97, Grids, HCtrls, HEnt1, EntPaie, HMsgBox, UTOF, UTOB, Vierge, P5Def,
  {$IFNDEF EAGLCLIENT}
  DBCtrls,
  HDB,
  HQry,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  HQuickRP,
  QRGrid,
  PrintDbg,
  Fe_main,
  {$ELSE}
  HQry,
  UtileAgl,
  Maineagl,
  {$ENDIF}
  P5Util,
  HSysMenu,
  AGLInit,
  ed_tools,
  mmsystem,
  PGVisuObjet;

type
  TOF_PG_SAISRUB = class(TOF)
  private
    LesMontants: array[1..35] of double;
    LesLib: array[1..35] of string;
    LesDecimales: array[1..35] of integer;
    LesAReporter: array[1..35] of string;
    LesRub: array[1..35] of string;
    LesTyp: array[1..35] of string;
    LesMasq: array[1..5] of string; // codes des masques traités
    LesDecal: array[1..5] of integer; // nbre de colonnes de décalage (décalge cumulé)
    NbCol: array[1..5] of integer; // nbre de colonnes ds chq masque
    QMul: TQUERY; // Query recuperee du mul
    LeTitre: string; // Code Salarie en traitement et Titre de la forme
    Modifier: Boolean;
    DateDebut, DateFin: TDateTime; // Date de debut exercice social
    LaGrille: THGrid; // Grilles de saisie des cumuls et des bases
    GrilleTotalisation: THGrid; // Grilles de Total des colonnes
    GrilleLib: THGrid; // Grille de stockage des libellés
    GrilleAide : THGrid; // PT13 Grille des aides en lignes des colonnes
    ColLib, LigLib: Integer;
    LeMasque, TOB_Masque: TOB; // TOB du masque de la saisie par rubrique
    TSal, LesRem: TOB; // TOB des salaries et des remunerations avec champ variable
    // PT2 : 03/06/2002 V582 PH Gestion historique des évènements
    MessCopier: string;
    LeLibelle: THEdit;
    TypeSaisieRubrique: string;
    function OnSauve: boolean;
    procedure ValiderClick(Sender: TObject);
    procedure FermeClick(Sender: TObject);
    function NbreDecimale(NbD: Integer): string;
    function RendSaisieRub(Salarie: string; DateDebut, DateFin: TDateTime; Decalage, NbreCols: Integer; LeMasque: TOB): Boolean;
    function RendSaisiePrec(Salarie: string; DateDebut, DateFin: TDateTime; Decalage, NbreCols: Integer; LeMasque: TOB): Boolean;
    procedure GrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure DrawCellule(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
//PT15    procedure GrilleCopierColler;
//PT15    function IdentSalarie(Salarie: string): Integer;
    procedure ImpClik(Sender: TObject);
    procedure ExitLibelle(Sender: TObject);
    procedure BtnExportClick(Sender: TObject); { PT4 }
    function  OnComputeCell(St: string): Variant ;
    procedure BtnDupliqLibClick(Sender: TObject); // PT11
    procedure BtnDupliqColClick(Sender: TObject); // PT11
    procedure BtnAbsClick(Sender: TObject); // PT12
    procedure MajTotalisation; // PT17
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
    procedure onLoad; Override;
  end;

implementation

uses
  UTofPG_AcompteGrp, //PT15
  StrUtils;


procedure TOF_PG_SAISRUB.OnArgument(Arguments: string);
var
  F: TFVierge;
  st, LeType, LaRubrique: string;
  Q: TQuery;
  NumMsq, Salarie, LeWhere: string;
  NbD, i, Ligne, j: Integer;
  AReprendre, Abandon, OkOk: Boolean;
  T1: TOB;
  BtnValid, BtnFerme, BtnImp, BtnExport,BtnDupliqLib, BtnDupliqCol, BtnAbs : TToolbarButton97; // PT11 PT12
  HMTrad: THSystemMenu;
begin
  inherited;
  // PT2 : 03/06/2002 V582 PH Gestion historique des évènements
  MessCopier := '';
  LeLibelle := THEdit(GetControl('LIBELLELIGNE'));
  if LeLibelle = nil then Exit;
  LeLibelle.OnExit := ExitLibelle;
  TSal := nil;
  BtnImp := TToolbarButton97(GetControl('Bimprimer'));
  if BtnImp <> nil then BtnImp.OnClick := ImpClik;
  st := Trim(Arguments);
  //PT14
  if pos(';BCP',st) > 0 then
  begin
    TypeSaisieRubrique := 'BCP';  //"Saisie par rubrique" des Bulletin complémentaire
    // On Vire les 4 derniers caractères qui sont ;XXX avec XXX qui représente le
    // type de la saisie de rubrique (SRB ou BCP)
    Arguments := Trim(LeftStr(Arguments,(Length(Arguments)-4)));
    Ecran.Caption := 'Saisie groupée des bulletins complémentaires';  //PT18
  end
  else if pos(';SRB',st) > 0 then
  begin
    TypeSaisieRubrique := 'SRB'; //"Saisie par rubrique" de la Paie
    Arguments := Trim(LeftStr(Arguments,(Length(Arguments)-4)));
    Ecran.Caption := 'Saisie par rubrique';
  end;
  //Fin PT14
  NumMsq := ReadTokenSt(st); // Recup des masques
  NumMsq := StringReplace(NumMsq, '+', ';', [rfReplaceAll]);
  DateDebut := StrToDate(ReadTokenSt(st)); // Recup de la date de debut
  DateFin := StrToDate(ReadTokenSt(st)); // Recup Date fin
  if st <> '' then LeWhere := ReadTokenSt(st)
  else LeWhere := '';
  // recuperation de la query du multicritere
  if not (Ecran is TFVierge) then exit;
  F := TFVierge(Ecran);
  if F <> nil then
  begin
    {$IFDEF EAGLCLIENT}
    QMUL := THQuery(F.FMULQ).TQ;
    {$ELSE}
    QMUL := F.FMULQ;
    {$ENDIF}
  end;
  if QMUL = nil then exit;
  // PT10 titre écran
  Ecran.Caption := Ecran.Caption + ' du ' +DateToStr(DateDebut)+ ' au '+DateToStr(DateFin)+' Utilisateur : '+ RechDom('TTUTILISATEUR', V_PGI.User, FALSE);
  BtnValid := TToolbarButton97(GetControl('BVALIDER'));
  if BtnValid <> nil then BtnValid.OnClick := ValiderClick;
  BtnFerme := TToolbarButton97(GetControl('BFERME'));
  if BtnFerme <> nil then BtnFerme.OnClick := FermeClick;
  // DEB PT11 Duplication du libellé de la cellule sur l'ensemble de la colonne et idem valeur cellule
  BtnDupliqLib := TToolbarButton97(GetControl('BTNREPORTLIB'));
  if BtnDupliqLib <> nil then BtnDupliqLib.OnClick := BtnDupliqLibClick;
  BtnDupliqCol := TToolbarButton97(GetControl('BTNDUPLIQUER'));
  if BtnDupliqCol <> nil then BtnDupliqCol.OnClick := BtnDupliqColClick;
  // FIN PT11
  BtnAbs := TToolbarButton97(GetControl('BTNABSENCE')); // PT12
  if BtnAbs <> nil then BtnAbs.OnClick := BtnAbsClick;  // PT12
  // design des grille et definition des tailles des colonnes
  GrilleLib := THGrid(Getcontrol('GRILLELIB')); // PT8
  if GrilleLib = nil then exit;
  GrilleAide := THGrid(Getcontrol('GRILLEAIDE')); // PT13
  if GrilleAide = nil then exit;

  LaGrille := THGrid(Getcontrol('GRILLE'));
  GrilleTotalisation := THGrid(Getcontrol('GRILLETOTAL'));//PT17

  if LaGrille <> nil then
  begin
    LaGrille.OnCellExit := GrilleCellexit;
    laGrille.OnKeyDown := KeyDown;
    laGrille.OnCellEnter := GrilleCellEnter;
    LaGrille.PostDrawCell := DrawCellule;
//@@    LaGrille.OnComputeCell := OnComputeCell;    [Erreur] UTofPG_SAISRUB.pas(153): Identificateur non déclaré : 'OnComputeCell'
//@@    LaGrille.CalcInCell := true;   [Erreur] UTofPG_SAISRUB.pas(154): Identificateur non déclaré : 'CalcInCell'
//@@    LaGrille.TraiteMontantNegatif := false;
    {
      LaGrille.ColLengths [0]:=10;
      LaGrille.ColLengths [1]:=37;
      LaGrille.ColLengths [2]:=12;}
    for i := 3 to 37 do
    begin
      LaGrille.ColLengths[i] := 17;
      LaGrille.ColTypes[i]:='K';
    end;
    for i := 0 to 37 do
    begin
      LaGrille.ColAligns[i] := taCenter;
      if i = 1 then LaGrille.ColAligns[i] := taLeftJustify;  //Pour la colonne du nom, on ne veut pas centrer (FQ 14090)
    end;
    if Assigned(GrilleTotalisation) then
    begin
      for i := 0 to GrilleTotalisation.ColCount -1 do
      begin
        if i <= (LaGrille.ColCount -1) then
        begin
          //PT17 On aligne les colones de la totalisation de la même manière que les colones de la saisie
          GrilleTotalisation.ColAligns[i] := LaGrille.ColAligns[i];
          //On met les mêmes largeurs de colone
          GrilleTotalisation.ColLengths[i] := LaGrille.ColLengths[i];
        end;
      end;
    end;
  end
  else exit;
  for i := 1 to 5 do
  begin
    LesMasq[i] := '';
    LesDecal[i] := 0;
    NbCol[i] := 0;
  end;
  TOB_Masque := TOB.Create('Les caractéristiques du masque', nil, -1);
  st := 'SELECT * FROM MASQUESAISRUB ORDER BY PMR_ORDRE';
  Q := OpenSql(st, TRUE);
  TOB_Masque.LoadDetailDB('MASQUESAISRUB', '', '', Q, FALSE, FALSE);
  Ferme(Q);
  for i := 1 to 4 do
  begin
    LesMasq[i] := ReadTokenSt(NumMsq); // codes des masques traités
    if LesMasq[i] = '' then break;
    LeMasque := TOB_Masque.FindFirst(['PMR_ORDRE'], [LesMasq[i]], FALSE);
    // Les décalages sont pour le traitement des colonnes suivantes donc indice +1 par rapport au masque en cours de traitement
    if LeMasque <> nil then
    begin
      LesDecal[i + 1] := LesDecal[i] + LeMasque.getvalue('PMR_NBRECOL');
      NbCol[i + 1] := NbCol[i + 1] + LeMasque.getvalue('PMR_NBRECOL');
    end;
  end;
  if LesMasq[1] = '' then
  begin
    PgiBox('Attention, aucun masque séléctionné', Ecran.caption);
    Close;
  end;
  // recup des remunerations ayant au moins un champ à saisir
  // pour connaitre les nbre de decimales à saisir dans la colonne en fonction du type de champ saisi
  LesRem := TOB.Create('Les Remunérations', nil, -1);
  //PT6
  Q :=
    OpenSql('SELECT * FROM REMUNERATION WHERE ##PRM_PREDEFINI## (PRM_TYPEBASE="00" OR PRM_TYPETAUX="00" OR PRM_TYPECOEFF="00" OR PRM_TYPEMONTANT="00") OR (PRM_TYPEBASE="01" OR PRM_TYPETAUX="01" OR PRM_TYPECOEFF="01" OR PRM_TYPEMONTANT="01")', TRUE);
  //FIN PT6
  LesRem.LoadDetailDB('REMUNERATION', '', '', Q, FALSE, FALSE);
  Ferme(Q);
  i := QMUL.RecordCount;
  InitMoveProgressForm(nil, 'Chargement des données de la saisie', 'Veuillez patienter SVP ...', i, FALSE, TRUE);
  TSal := TOB.Create('Les Salaries', nil, -1);
  St := 'SELECT PSA_SALARIE,PSA_CONFIDENTIEL,PSA_ETABLISSEMENT FROM SALARIES '; // PT9
  if LeWhere <> '' then st := st + LeWhere;
  St := St + ' ORDER BY PSA_SALARIE';
  Q := OpenSql(st, TRUE);
  TSal.LoadDetailDB('SALARIES', '', '', Q, FALSE, FALSE);
  Ferme(Q);
  for i := 1 to 35 do LesAReporter[i] := 'NON';
  LaGrille.CacheEdit;
  LaGrille.SynEnabled := False;
  LaGrille.Cells[0, 0] := 'Matricule';
  LaGrille.Cells[1, 0] := 'Nom';
  LaGrille.Cells[2, 0] := 'Prénom';

  for j := 1 to 4 do
  begin
    if LesMasq[j] = '' then break;
    LeMasque := TOB_Masque.FindFirst(['PMR_ORDRE'], [LesMasq[j]], FALSE);
    if LeMasque <> nil then
    begin
      for i := 1 to 7 do
      begin
        LaRubrique := LeMasque.GetValue('PMR_COL' + IntToStr(i));
        T1 := LesRem.FindFirst(['PRM_RUBRIQUE'], [LaRubrique], FALSE);
        if T1 <> nil then
        begin
          LeType := LeMasque.GetValue('PMR_TYPECOL' + IntToStr(i));
          if LeType = 'BAS' then NbD := T1.GetValue('PRM_DECBASE')
          else if LeType = 'COE' then NbD := T1.GetValue('PRM_DECCOEFF')
          else if LeType = 'TAU' then NbD := T1.GetValue('PRM_DECTAUX')
          else if LeType = 'MON' then NbD := T1.GetValue('PRM_DECMONTANT');
          LesRub[i + LesDecal[j]] := LeMasque.GetValue('PMR_COL' + IntToStr(i));
          LesTyp[i + LesDecal[j]] := LeMasque.GetValue('PMR_TYPECOL' + IntToStr(i));
          LesDecimales[i + LesDecal[j]] := NbD;
          LesAReporter[i + LesDecal[j]] := LeMasque.GetValue('PMR_REPORTCOL' + IntToStr(i)); // Stockage des colonnes à reporter
          LaGrille.Cells[i + 2 + LesDecal[j], 0] := LeMasque.GetValue('PMR_LIBCOL' + IntToStr(i));
          GrilleAide.Cells[i + 2 + LesDecal[j], 0] := LeMasque.GetValue('PMR_AIDECOL' + IntToStr(i));// PT13
        end; // Fin si Rem connue
      end; // Fin Boucle sur nbre de colonnes gérées
    end; // si masque existe
  end; // boucle sur les masques
  // Boucle de chargement de la Grille
  j := 0;
  for i := 1 to 5 do j := j + NbCol[i]; // nbre de colonnes + identification salarie
  LaGrille.ColCount := j + 3;
  GrilleLib.colCount := j + 3; // PT8
  GrilleAide.colCount := j + 3; // PT13
  j := 0;
  QMul.First;
  LaGrille.RowCount := 2;
  Ligne := 1; // Ligne courante
  Abandon := FALSE;
  while not QMul.EOF do
  begin
    Salarie := QMul.FindField('PSA_SALARIE').AsString;
    LaGrille.Cells[0, Ligne] := Salarie;
    LaGrille.Cells[1, Ligne] := QMul.FindField('PSA_LIBELLE').AsString; // Nom
    LaGrille.Cells[2, Ligne] := QMul.FindField('PSA_PRENOM').AsString; // Prénom
    for i := 1 to 35 do
    begin
      LesMontants[i] := 0;
      LesLib[i] := '';
    end;
    for j := 1 to 4 do // boucle sur les masques
    begin
      if LesMasq[j] = '' then break;
      LeMasque := TOB_Masque.FindFirst(['PMR_ORDRE'], [LesMasq[j]], FALSE);
      AReprendre := RendSaisieRub(Salarie, DateDebut, DateFin, LesDecal[j], NbCol[j + 1], LeMasque); // Recup de la saisie déjà effectuée
      if AReprendre = FALSE then
      begin
        RendSaisiePrec(Salarie, DateDebut, DateFin, LesDecal[j], NbCol[j + 1], LeMasque); // Recup de la saisie précédente
        // PT3 : 28/10/2002 V582 PH Sauvegarde de la saisie car on a récupéré des infos du mois precedent donc il faut les valider
        Modifier := TRUE;
      end;
      for i := 1 to LesDecal[j + 1] do
      begin
        if i + 2 + LesDecal[j] > LaGrille.ColCount then break;
        GrilleLib.Cells[i + 2 + LesDecal[j], Ligne] := LesLib[i + LesDecal[j]];
        if AReprendre = TRUE then LaGrille.Cells[i + 2 + LesDecal[j], Ligne] := DoubleToCell(LesMontants[i + LesDecal[j]], LesDecimales[i + LesDecal[j]])
        else // Saisie à reprendre depuis la derniere saisie
        begin
          if LesAReporter[i + LesDecal[j]] = 'OUI' then
          begin // montant à reporter du mois précédent
            LaGrille.Cells[i + 2 + LesDecal[j], Ligne] := DoubleToCell(LesMontants[i + LesDecal[j]], LesDecimales[i + LesDecal[j]]); //
          end
          else LaGrille.Cells[i + 2 + LesDecal[j], Ligne] := '';
        end;
      end; // fin boucle sur les colonnes
    end; // fin boucle sur les masques
    QMul.NEXT;
    Ligne := Ligne + 1;
    okok := MoveCurProgressForm(IntToStr(Ligne));
    if not OkOk then
    begin
      Abandon := TRUE;
      break;
    end;
    LaGrille.RowCount := LaGrille.RowCount + 1;
    GrilleLib.RowCount := GrilleLib.RowCount + 1;
    GrilleAide.RowCount := 1; // PT13
  end; // si query non nulle
  FiniMoveProgressForm();
  SetControlText('LIBELLELIGNE', GrilleLib.Cells[3, 1]);
  SetControlText('LIBELLEAIDE', GrilleAide.Cells[3, 0]); // PT13
  if Abandon = TRUE then Close; // Abandon de la saisie par rubrique
  LaGrille.MontreEdit;
  LaGrille.SynEnabled := True;
  LaGrille.RowCount := LaGrille.RowCount - 1;
  { DEB PT4 }
  BtnExport := TToolbarButton97(GetControl('BEXPORTER'));
  if BtnExport <> nil then BtnExport.OnClick := BtnExportClick;
  { FIN PT4 }
// PT7 : 01/06/2004 PH V_50 Design de la grille géré par le HMTRAD
  HMTrad.ResizeGridColumns(LaGrille);
end;

// Fonction de validation de la saisie
procedure TOF_PG_SAISRUB.ValiderClick(Sender: TObject);
var
  rep: Integer;
begin
  inherited;
  Rep := PGIAsk('Voulez vous sauvegarder votre saisie ?', Ecran.Caption);
  if rep = mrNo then exit;
  if rep = mrCancel then exit;                            
  if rep = mryes then OnSauve;
end;
// fermeture de la forme
procedure TOF_PG_SAISRUB.FermeClick(Sender: TObject);
begin
  ValiderClick(Sender);
  Close;
end;

// fonction d'ecriture des elements saisis
function TOF_PG_SAISRUB.OnSauve: boolean;
var
  st, LeType, St1: string;
  i, j: Integer;
  T_Sais, LeSal, T: TOB;
  Trace: TStringList;
  Err: Boolean;
begin
  // PT2 : 03/06/2002 V582 PH Gestion historique des évènements
  Err := FALSE;

  result := TRUE;
  if Modifier = FALSE then exit;
  try
    BeginTrans;
    Modifier := FALSE;
    T_Sais := TOB.Create('Les lignes de saisie', nil, -1);
    for i := 1 to LaGrille.RowCount - 1 do
    begin
      // Montant := 0;
      // for j := 1 to NbreCol do Montant := Montant + Valeur (LaGrille.Cells [J + 2, i]);
      // if Montant = 0 then continue;  // on stocke es occurences nulles, car il faut pouvoir annuler les lignes saisies
      for j := 1 to LaGrille.colCount - 3 do
      begin
        //  if Valeur (LaGrille.Cells [J + 2, i]) = 0 then continue; // faut pouvoir annuler les lignes saisies
        T := TOB.Create('HISTOSAISRUB', T_Sais, -1);
        if T <> nil then
        begin
          LeSal := TSal.FindFirst(['PSA_SALARIE'], [LaGrille.Cells[0, i]], FALSE);
          if LeSal <> nil then T.PutValue('PSD_CONFIDENTIEL', LeSal.GetValue('PSA_CONFIDENTIEL'));
          T.PutValue('PSD_ORIGINEMVT', TypeSaisieRubrique); //PT14
          T.PutValue('PSD_SALARIE', LaGrille.Cells[0, i]);
          T.PutValue('PSD_ETABLISSEMENT', LeSal.GetValue('PSA_ETABLISSEMENT')); // PT9
          T.PutValue('PSD_DATEDEBUT', DateDebut);
          T.PutValue('PSD_DATEFIN', DateFin);
          T.PutValue('PSD_RUBRIQUE', LesRub[j]);
          T.PutValue('PSD_TYPALIMPAIE', LesTyp[j]);
          LeType := LesTyp[j];
          if LeType = 'MON' then T.PutValue('PSD_MONTANT', Valeur(LaGrille.Cells[J + 2, i]));
          if LeType = 'BAS' then T.PutValue('PSD_BASE', Valeur(LaGrille.Cells[J + 2, i]));
          if LeType = 'TAU' then T.PutValue('PSD_TAUX', Valeur(LaGrille.Cells[J + 2, i]));
          if LeType = 'COE' then T.PutValue('PSD_COEFF', Valeur(LaGrille.Cells[J + 2, i]));
          T.PutValue('PSD_LIBELLE', GrilleLib.Cells[J + 2, i]);
          T.PutValue('PSD_ORDRE', j);
          T.PutValue('PSD_DATEINTEGRAT', IDate1900);//PT13
          T.PutValue('PSD_DATECOMPT', IDate1900);//PT13
        end;
      end; // boucle sur les colonnes
    end; // Boucle sur les lignes de la grille
    T_Sais.SetAllModifie(TRUE);
    T_Sais.InsertOrUpdateDB(TRUE);
    CommitTrans;
  except
    result := FALSE;
    Rollback;
    PGIBox('Une erreur est survenue lors de la validation de la saisie', LeTitre);
    Err := TRUE;
  end;
  // PT2 : 03/06/2002 V582 PH Gestion historique des évènements
  if Err then st1 := 'Erreur de ';
  st := 'Saisie par rubrique';
  Trace := TStringList.Create;
  Trace.add(St);
  if not Err then Trace.add(IntToStr(T_Sais.detail.count - 1) + ' cellules saisies');
  if MessCopier <> '' then
  begin
    Trace.add('Insertion automatique');
    Trace.add(MessCopier);
  end;
  if Err then CreeJnalEvt('001', '010', 'ERR', nil, nil, Trace)
  else CreeJnalEvt('001', '010', 'OK', nil, nil, Trace);
  Trace.free;
  if T_Sais <> nil then T_Sais.free;
  // FIN PT2
  Modifier := FALSE;
end;


function TOF_PG_SAISRUB.NbreDecimale(NbD: Integer): string;
var
  i: Integer;
begin
  result := '';
  if NbD = 0 then exit;
  result := '.';
  for i := 1 to NbD do
  begin
    result := result + '0';
  end;
end;

function TOF_PG_SAISRUB.RendSaisieRub(Salarie: string; DateDebut, DateFin: TDateTime; Decalage, NbreCols: Integer; LeMasque: TOB): Boolean;
var
  Q: TQuery;
  St: string;
  i: Integer;
  Montant: Double;
begin
  result := FALSE;
  st := 'SELECT * FROM HISTOSAISRUB WHERE PSD_SALARIE="' + Salarie + '" AND PSD_DATEDEBUT ="' + USDateTime(DateDebut)
    + '" AND PSD_ORDRE >' + InttoStr(decalage) // DB2
  + ' AND PSD_ORDRE <=' + IntToStr(Decalage + NbreCols) // DB2
  + ' AND PSD_DATEFIN ="' + USDateTime(DateFin) + '" AND PSD_ORIGINEMVT="'+TypeSaisieRubrique+'" ORDER BY PSD_SALARIE,PSD_DATEDEBUT,PSD_DATEFIN,PSD_ORDRE'; //PT14
  Q := OpenSql(St, TRUE);
  while not Q.EOF do
  begin
    i := Q.FindField('PSD_ORDRE').AsInteger;
    Montant := 0;
    if Q.FindField('PSD_TYPALIMPAIE').AsString = 'BAS' then Montant := Q.FindField('PSD_BASE').AsFloat
    else if Q.FindField('PSD_TYPALIMPAIE').AsString = 'MON' then Montant := Q.FindField('PSD_MONTANT').AsFloat
    else if Q.FindField('PSD_TYPALIMPAIE').AsString = 'COE' then Montant := Q.FindField('PSD_COEFF').AsFloat
    else if Q.FindField('PSD_TYPALIMPAIE').AsString = 'TAU' then Montant := Q.FindField('PSD_TAUX').AsFloat;
    if (i > decalage) and (i <= decalage + Nbrecols) then
    begin
      {
        for j := 1 to 7 do
            begin
            if LeMasque.GetValue ('PMR_COL'+IntToStr (j)) = Q.FindField ('PSD_RUBRIQUE').AsString then
              LesMontants [j] := Montant; // Test du numero d'ordre car il represente le numero de colonne
            end;
      }
      if LesRub[i] = Q.FindField('PSD_RUBRIQUE').AsString then
      begin
        LesMontants[i] := Montant;
        LesLib[i] := Q.FindField('PSD_LIBELLE').AsString; //  PT8
      end;
    end;
    result := TRUE;
    Q.next;
  end;
  Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Philippe Dumet
Créé le ...... : 30/05/2001
Modifié le ... :   /  /
Description .. : Saisie par rubriques des éléménts variables
Suite ........ : Utilise le CTRL_C et CTRL_V entre une feuille EXCEL
Suite ........ : et une grille
Suite ........ : Analyse de la tob et rempli la grille en remettant dans les
Suite ........ : bonnes lignes et les bonnes colonnes
Suite ........ : Indique si erreur identification
Mots clefs ... : PAIE;PGSAISRUB
*****************************************************************}
function TOF_PG_SAISRUB.RendSaisiePrec(Salarie: string; DateDebut, DateFin: TDateTime; Decalage, NbreCols: Integer; LeMasque: TOB): Boolean;
var
  Q: TQuery;
  St: string;
  i: Integer;
  Deb, Fin: TDateTime;
  Montant: Double;
begin
  result := FALSE;
  Deb := DebutDeMois(DateDebut) - 1;
  Deb := DebutDeMois(Deb);
  Fin := FindeMois(Deb);
  st := 'SELECT * FROM HISTOSAISRUB WHERE PSD_SALARIE="' + Salarie + '" AND PSD_DATEFIN ="' + USDateTime(FIN)
    + '" AND PSD_ORDRE >' + InttoStr(decalage) // DB2
  + ' AND PSD_ORDRE <=' + IntToStr(Decalage + NbreCols) // DB2
  + ' AND PSD_DATEDEBUT ="' + USDateTime(Deb) + '" AND PSD_ORIGINEMVT="'+TypeSaisieRubrique+'" ORDER BY PSD_SALARIE,PSD_DATEDEBUT,PSD_DATEFIN,PSD_ORDRE';//PT14
  Q := OpenSql(St, TRUE);
  while not Q.EOF do
  begin
    i := Q.FindField('PSD_ORDRE').AsInteger;
    Montant := 0;
    if (LesAReporter[i] = 'OUI') and (i >= decalage) and (i <= decalage + nbrecols) then
    begin
      if Q.FindField('PSD_TYPALIMPAIE').AsString = 'BAS' then Montant := Q.FindField('PSD_BASE').AsFloat
      else if Q.FindField('PSD_TYPALIMPAIE').AsString = 'MON' then Montant := Q.FindField('PSD_MONTANT').AsFloat
      else if Q.FindField('PSD_TYPALIMPAIE').AsString = 'COE' then Montant := Q.FindField('PSD_COEFF').AsFloat
      else if Q.FindField('PSD_TYPALIMPAIE').AsString = 'TAU' then Montant := Q.FindField('PSD_TAUX').AsFloat;
      if i <> 0 then LesMontants[i] := Montant;
      if LesRub[i] = Q.FindField('PSD_RUBRIQUE').AsString then
        result := TRUE;
    end;
    Q.next;
  end;
  Ferme(Q);
end;

procedure TOF_PG_SAISRUB.GrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  Nbre: Double;
begin
  Modifier := TRUE;
  if (LaGrille.Row - 1 = 0) and (ARow = 1) then
  begin
    MajTotalisation;
    exit;
  end;
  if (LaGrille.Cells[Acol, LaGrille.Row - 1] = '') and (ARow <> LaGrille.RowCount - 1) then
  begin
    MajTotalisation;
    exit;
  end;;
  if (IsNumeric(LaGrille.Cells[Acol, LaGrille.Row - 1])) then
  begin
    if ARow = LaGrille.RowCount - 1 then
    begin
      if IsNumeric(LaGrille.Cells[Acol, ARow]) then
      begin
        if LaGrille.Cells[Acol, ARow] <> '' then Nbre := Valeur(LaGrille.Cells[Acol, ARow])
        else Nbre := 0;
        if ARow <> 0 then LaGrille.Cells[Acol, ARow] := DoubleToCell(Nbre, LesDecimales[ACol - 2]);
      end
      else if ARow <> 0 then LaGrille.Cells[Acol, ARow] := '';
    end
    else
    begin
      Nbre := Valeur(LaGrille.Cells[Acol, LaGrille.Row - 1]);
      if LaGrille.Row - 1 <> 0 then LaGrille.Cells[Acol, LaGrille.Row - 1] := DoubleToCell(Nbre, LesDecimales[ACol - 2]);
    end;
  end
  else
  begin
    if LaGrille.Row - 1 <> 0 then LaGrille.Cells[Acol, LaGrille.Row - 1] := DoubleToCell(0, LesDecimales[ACol - 2]);
    if ARow = LaGrille.RowCount - 1 then
    begin
      if IsNumeric(LaGrille.Cells[Acol, ARow]) then
      begin
        Nbre := Valeur(LaGrille.Cells[Acol, ARow]);
        if ARow <> 0 then LaGrille.Cells[Acol, ARow] := DoubleToCell(Nbre, LesDecimales[ACol - 2]);
      end
      else if ARow <> 0 then LaGrille.Cells[Acol, ARow] := '';
    end;
  end;
  MajTotalisation;//PT17
end;


procedure TOF_PG_SAISRUB.OnClose;
begin
  if TSal <> nil then
  begin
    TSal.Free;
    TSal := nil;
  end;
  if TOB_Masque <> nil then
  begin
    TOB_Masque.free;
    TOB_Masque := nil;
  end;
  if LesRem <> nil then
  begin
    LesRem.free;
    LesRem := nil;
  end;
end;

procedure TOF_PG_SAISRUB.KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = 86) and (ssCtrl in Shift) then
  begin
    LaGrille.Options  := LaGrille.Options - [goEditing]; // PT20: Désactivation de l'édition (Copier(CTRL-C)-Coller(CTRL-V))
    GrilleCopierColler(LaGrille,3);
    LaGrille.Options  := LaGrille.Options + [goEditing]; // PT20: Ré-activation de l'édition (Copier(CTRL-C)-Coller(CTRL-V))

    MajTotalisation;//PT17
  end;
end;

{//PT15
procedure TOF_PG_SAISRUB.GrilleCopierColler;
var
  T, TT: TOB;
  i, j, k, NbreCol, LstErrorL, LstErrorP: integer;
  St, Salarie: string;
  LeVar: Variant;
  TitreCol: array[1..35] of string;
begin
  LstErrorL := 0;
  LstErrorP := 0;
  T := TOBLoadFromClipBoard;
  T.Detail[0].Free;

  NbreCol := LaGrille.ColCount - 3; // Car 3 colonnes figees Matricule,Nom,Prenom
  for i := 1 to 35 do TitreCol[i] := '';
  for i := 1 to 35 do
  begin
    if i > NbreCol then break;
    TitreCol[i] := LaGrille.Cells[i + 2, 0];
  end;
  for i := 0 to T.Detail.Count - 1 do
  begin
    TT := T.Detail[i];
    LeVar := TT.GetValue('Matricule');
    Salarie := VarToStr(LeVar);
    k := IdentSalarie(Salarie);
    if k = -1 then LstErrorL := LstErrorL + 1;
    if (k > 0) and (k < LaGrille.RowCount) then
    begin // Boucle sur les lignes de la TOB
      for j := 1 to 35 do
      begin // Boucle sur les colonnes de la grille qui correspondent aux colonnes de la feuilles EXCEL
        st := TitreCol[j];
        LeVar := TT.GetValue(TitreCol[j]);
        if LeVar <> '' then // colonne identifiée
          if IsNumeric(VarToStr(LeVar)) then LaGrille.Cells[j + 2, k] := DoubleToCell(Valeur(VarToStr(LeVar)), 5)
          else LstErrorP := LstErrorP + 1;
      end;
    end;
  end;
  T.Free;
  St := 'Attention, ';
  if LstErrorL > 0 then St := St + IntTOstr(LstErrorL) + ' ligne(s) n''ont pas été intégrées';
  if LstErrorP > 0 then St := St + IntTOstr(LstErrorP) + ' cellule(s) colonne n''ont pas été intégrées';
  if (LstErrorL > 0) then PGIBox(St, Ecran.Caption);
  // PT2 : 03/06/2002 V582 PH Gestion historique des évènements
  MessCopier := St;
end;}

{//PT15
function TOF_PG_SAISRUB.IdentSalarie(Salarie: string): Integer;
var
  i: Integer;
begin
  result := -1;
  for i := 1 to LaGrille.RowCount - 1 do
  begin
    if VH_PAie.PgTypeNumSal = 'NUM' then
    begin
      if Valeur(Salarie) = Valeur(LaGrille.Cells[0, i]) then
      begin
        Result := i;
        break;
      end;
    end
    else
    begin
      if LaGrille.Cells[0, i] = Salarie then
      begin
        Result := i;
        break;
      end;
    end;
  end;
end;}

procedure TOF_PG_SAISRUB.ImpClik(Sender: TObject);
begin
  {$IFNDEF EAGLCLIENT}
  PrintDBGrid(TCustomGrid(LaGrille), nil, Ecran.Caption, '');
  {$ELSE}
  PrintDBGrid('GRILLE', '', Ecran.Caption, '');
  {$ENDIF}
end;


{ DEB PT4 : Export de la grille }
procedure TOF_PG_SAISRUB.BtnExportClick(Sender: TObject);
var
  Grille: ThGrid;
  SD: TSaveDialog;
begin
  Grille := ThGrid(GetControl('GRILLE'));
  if Grille = nil then exit;
  SD := TSaveDialog.Create(nil);
  SD.Filter := 'Fichier Excel (*.xls)|*.xls';
  if SD.Execute then
  begin
    //Fichier Excel (*.xls)|*.xls|  SD.FilterIndex = 1
    if (SD.FilterIndex = 1) and (Pos('.xls', SD.FileName) < 1) then
      SD.FileName := Trim(SD.FileName) + '.xls';
    ExportGrid(Grille, nil, SD.FileName, 2, True); //PT19 code 2 correspond aux fichiers Excel  SD.FilterIndex, True);
  end;
  SD.Destroy;
end;
{ FIN PT4 }
// DEB PT8
procedure TOF_PG_SAISRUB.GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  ColLib := LaGrille.Col;
  LigLib := LAGrille.Row;
  SetControlText('LIBELLELIGNE', GrilleLib.Cells[ColLib, LigLib]);
  SetControlText('LIBELLEAIDE', GrilleAide.Cells[ColLib, 0]); // PT13
end;

procedure TOF_PG_SAISRUB.ExitLibelle(Sender: TObject);
begin
  if (ColLib = 0) and (LigLib = 0) then
  begin
    ColLib := 3;
    LigLib := 1;
  end;
  GrilleLib.Cells[ColLib, LigLib] := GetControlText('LIBELLELIGNE');
end;

procedure TOF_PG_SAISRUB.DrawCellule(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var
  i, j: Integer;
  R: TRect;
begin
  //  if GrilleLib.Cells [ACol, ARow] <> '' then Canvas.Font.Color := clRed;
  if ACol < laGrille.FixedCols then exit;
  if ARow < laGrille.FixedRows then exit;
  if (GrilleLib.Cells[ACol, ARow] <> '') then
  begin
    {
    R := laGrille.CellRect(ACol,ARow);
    Canvas.Brush.Color := clRed;
    Canvas.Rectangle(Rect(R.Left+lagrille.Left,R.Top+lagrille.Top,5,5));
    Canvas.FillRect(Rect(R.Left+lagrille.Left,R.Top+lagrille.Top,5,5));}
    Canvas.Brush.Color := clRed;
    Canvas.Brush.Style := bsSolid;
    Canvas.Pen.Color := clRed;
    Canvas.Pen.Mode := pmCopy;
    Canvas.Pen.Width := 1;
    R := LaGrille.CellRect(ACol, ARow);
    Canvas.Rectangle(R.Right - 5, R.Bottom - 1, R.Right - 2, R.Bottom - 5);
  end;
end;
// FIN PT8

function TOF_PG_SAISRUB.OnComputeCell(St: string): Variant;
begin
  result := TOB(LaGrille.Objects[0,LaGrille.Row-LaGrille.FixedRows]).GetValue(st);
end;
// DEB PT11
procedure TOF_PG_SAISRUB.BtnDupliqLibClick(Sender: TObject);
var i: Integer;
begin
  For i := 1 to LaGrille.RowCount -1 do
     GrilleLib.Cells[LaGrille.Col, i] := GetControlText('LIBELLELIGNE');
end;

procedure TOF_PG_SAISRUB.BtnDupliqColClick(Sender: TObject);
var i : Integer;
    St : String;
begin
  st := LaGrille.Cells[LaGrille.Col, LaGrille.Row];
  For i := 1 to LaGrille.RowCount -1 do
     LaGrille.Cells[LaGrille.Col, i] := LaGrille.Cells[LaGrille.Col, LaGrille.Row];
end;
// FIN PT11
// DEB PT12
procedure TOF_PG_SAISRUB.BtnAbsClick(Sender: TObject);
Var LeSal : String;
begin
  LeSal := LaGrille.Cells [0, LaGrille.Row];
  If LeSal <> '' then   AglLanceFiche('PAY', 'ABSENCE_MUL', '', 'PCN_SALARIE=' + LeSal, LeSal + ';SAISRUB;;' + DateToStr(DateFin));    //PT16 SAISRUB au lieu de BULLETIN
end;
// FIN PT12


// Debut PT17
procedure TOF_PG_SAISRUB.MajTotalisation;
var
  ColIndex, RowIndex, i : Integer;
  Value : Extended;
  Function MajTotalisationStrToFloat(Str: String) : Extended;
  var
    FormatSet : TFormatSettings;
    Finded : Boolean;
  begin
    GetLocaleFormatSettings(0,FormatSet);
    result := 0;
    if (Str <> '') and (IsNumeric(Str)) then
    begin
      Finded := True;
      While Finded do
      begin
        Finded := False;
        i:=1;
        While i <= Length(Str) do
        begin
          if (not (Str[i] in ['0','1','2','3','4','5','6','7','8','9',',','.']))
             and (i <= Length(Str)) then
          begin
            Finded := True;
            Str := StringReplace(Str,Str[i],'',[rfReplaceAll,rfIgnoreCase]);
          end;
          Inc(i);
        end;
      end;
      try
        FormatSet.DecimalSeparator := ',';
        result := StrToFloat(Str,FormatSet);
      except
        FormatSet.DecimalSeparator := '.';
        result := StrToFloat(Str,FormatSet);
      end;
    end;
  end;
begin
  GrilleTotalisation.ColCount := LaGrille.ColCount;
  GrilleTotalisation.FixedCols := LaGrille.FixedCols;
  GrilleTotalisation.ColWidths[0] := LaGrille.ColWidths[0];
  for ColIndex := 1 to LaGrille.ColCount-1 do
  begin
    GrilleTotalisation.ColWidths[ColIndex] := LaGrille.ColWidths[ColIndex];
    Value := 0;
    for RowIndex := LaGrille.FixedRows to LaGrille.RowCount-1 do
    begin
      if (LaGrille.CellValues[ColIndex,RowIndex] <> '') then
      begin
        if (IsNumeric(LaGrille.CellValues[ColIndex,RowIndex])) then
        begin
          try
            Value := Value + MajTotalisationStrToFloat(LaGrille.CellValues[ColIndex,RowIndex]);
          except  //Le transtypage en flottant c'est mal passé
            on E: Exception do begin
              Value := 0;
              Break;
            end;
          end;
        end else begin
          Value := 0;
          Break;
        end;
      end;
    end;
    if Value <> 0 then
      GrilleTotalisation.Cells[ColIndex,0] := FloatToStr(Value)
    else
      GrilleTotalisation.Cells[ColIndex,0] := '';
  end;
end;
//Fin PT17

procedure TOF_PG_SAISRUB.onLoad;
begin
  inherited;
  MajTotalisation; // PT17
end;

initialization
  registerclasses([TOF_PG_SAISRUB]);
end.

