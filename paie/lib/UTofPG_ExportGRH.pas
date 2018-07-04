{***********UNITE*************************************************
Auteur  ...... : Philippe Dumet
Créé le ...... : 12/06/2002
Modifié le ... : 12/06/2002
Description .. :
Suite ........ :
Suite ........ :
Suite ........ :
Mots clefs ... : PAIE;EXPORTGRH
*****************************************************************}
{ Ce module d'exportation permet d'alimenter une table "à plat" de certaines
données concernant le salarié à une date donnée.
L'alimentation des soldes CP et RTT se fait à partir de la table RECAPSAL
qui calcule toutes les informations.
==> penser à contrôler les dates pour ne pas stocker des infos si la table RECAPSAL
n'est pas à jour.
A ce jour, 20 montants calculés sont stockés et les requetes n'attaquent que la table
histobulletin ==> Specif CEGID
Variante à définir paramètrer le nombre, le type, la table d'où l'on veut extraire
des informations

}
unit UTofPG_ExportGRH;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97, Hqry,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Mul, Fe_Main,
{$ELSE}
  MaineAgl, eMul,
{$ENDIF}
  Grids, HCtrls, HEnt1, EntPaie, ed_tools, HMsgBox, UTOF, UTOB, UTOM, Vierge, P5Util,
  P5Def, PGVisuObjet, AGLInit, PgOutils;
type
  TOF_PGExportGRH = class(TOF)
  private
    Date1, Date2: THEdit;
    TSal, TRecap, TTMP: TOB; // Tob utilisees
    TSQL: array[1..10] of TOB;
    TT: TQRProgressForm;
    procedure LanceExport(Sender: TObject);
    procedure DateDebutExit(Sender: TObject);
    procedure DateFinExit(Sender: TObject);
    function RendTypContrat(LeSalarie: string; Quoi: string = ''): string;
    procedure RAZTOB;
  public
    procedure OnArgument(Arguments: string); override;
  end;

implementation

procedure TOF_PGExportGRH.DateDebutExit(Sender: TObject);
var
  DateDeb, DateFin: TDateTime;
begin

  DateDeb := StrToDate(GetControlText('XX_VARIABLEDEB'));
  DateFin := StrToDate(GetControlText('XX_VARIABLEFIN'));
  if DateFin < DateDeb then
  begin
    PGIBox('Attention, la date de début est supérieure à la date de fin', 'Préparation automatique');
    SetFocusControl('XX_VARIABLEDEB');
  end;

end;

procedure TOF_PGExportGRH.DateFinExit(Sender: TObject);
var
  DateDeb, DateFin: TDateTime;
begin

  DateDeb := StrToDate(GetControlText('XX_VARIABLEDEB'));
  DateFin := StrToDate(GetControlText('XX_VARIABLEFIN'));
  if DateFin < DateDeb then
  begin
    PGIBox('Attention, la date de fin est inférieure à la date de début', 'Préparation automatique');
    SetFocusControl('XX_VARIABLEDEB');
  end;

end;

procedure TOF_PGExportGRH.LanceExport(Sender: TObject);
var
  st, LeWhere: string;
  DD, DF: TDateTime;
  i, j: Integer;
  Q: TQuery;
  T1, T2, T3: TOB; // des tob filles
  FLect: TextFile; // Canaux des fichiers
  FileM, Pref, S: string; // Nom des fichiers
const
  Hist: array[1..47] of string = (
    'SALARIE', 'LIBELLE', 'PRENOM', 'DATEENTREE', 'DATESORTIE', 'SUSPENSIONPAIE', 'DATEANCIENNETE', 'HORAIREMOIS',
    'SALAIRETHEO', 'TAUXHORAIRE', 'SALAIREMOIS1', 'SALAIREMOIS2', 'SALAIREMOIS3', 'SALAIREMOIS4', 'SALAIRANN1', 'SALAIRANN2',
    'SALAIRANN3', 'SALAIRANN4', 'PRISEFFECTIF', 'UNITEPRISEFF', 'DATELIBRE1', 'DATELIBRE2', 'DATELIBRE3', 'DATELIBRE4',
    'BOOLLIBRE1', 'BOOLLIBRE2', 'BOOLLIBRE3', 'BOOLLIBRE4', '', '', '', '',
    '', '', '', '', '', '', '', '',
    '', '', '', '', '', '', '');
const
  HistT: array[1..47] of string = (
    'ETABLISSEMENT', 'COEFFICIENT', 'QUALIFICATION', 'CODEEMPLOI', 'TRAVAILN1', 'TRAVAILN2', 'TRAVAILN3', 'TRAVAILN4',
    'CODESTAT', 'INDICE', 'NIVEAU', 'LIBREPCMB1', 'LIBREPCMB2', 'LIBREPCMB3', 'LIBREPCMB4', 'CATDADS',
    'DADSCAT', 'CONDEMPLOI', '', '', '', '', '', '',
    '', '', '', '', '', '', '', '',
    '', '', '', '', '', '', '', '',
    '', '', '', '', '', '', '');
const
  MesSql: array[1..20] of string = (
    ' AND PHB_NATURERUB="AAA" AND PHB_RUBRIQUE="0001"', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '');
begin

  DD := StrToDate(GetControlText('XX_VARIABLEDEB'));
  DF := StrToDate(GetControlText('XX_VARIABLEFIN'));
  LeWhere := RecupWhereCritere(TFMUL(ECran).Pages); // Recup de la clause where generee par le mul
  if LeWhere <> '' then LeWhere := LeWhere + ' AND '
  else LeWhere := ' WHERE ';
  LeWhere := LeWhere + ' (PSA_DATEENTREE <="' + UsDateTime(DF) + '") AND (((PSA_DATESORTIE >="' + UsDateTime(DD) +
    '")) OR (PSA_DATESORTIE IS NULL) OR (PSA_DATESORTIE <= "' + UsDateTime(iDate1900) + '"))';

  st := 'SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_DATEENTREE,PSA_DATESORTIE,PSA_SUSPENSIONPAIE,' +
    'PSA_COEFFICIENT,PSA_QUALIFICATION,PSA_CODEEMPLOI,PSA_DATEANCIENNETE,PSA_HORAIREMOIS,PSA_SALAIRETHEO,PSA_TAUXHORAIRE,' +
    'PSA_SALAIREMOIS1,PSA_SALAIREMOIS2,PSA_SALAIREMOIS3,PSA_SALAIREMOIS4,PSA_SALAIRANN1,PSA_SALAIRANN2,PSA_SALAIRANN3,' +
    'PSA_SALAIRANN4,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_INDICE,PSA_NIVEAU,PSA_PRISEFFECTIF,' +
    'PSA_UNITEPRISEFF,PSA_DATELIBRE1,PSA_DATELIBRE2,PSA_DATELIBRE3,PSA_DATELIBRE4,PSA_BOOLLIBRE1,PSA_BOOLLIBRE2,PSA_BOOLLIBRE3,' +
    'PSA_BOOLLIBRE4,PSA_LIBREPCMB1,PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4,PSA_CATDADS,PSA_DADSCAT,PSA_CONDEMPLOI FROM ' +
    'SALARIES ' + LeWhere + ' ORDER BY PSA_SALARIE';
  Tsal := TOB.create('Mes Salaries', nil, -1);
  Q := OpenSql(st, TRUE);
  Tsal.LoadDetailDB('SALARIES', '', '', Q, FALSE, False);
  Ferme(Q);
  j := TSal.Detail.count - 1;
  InitMoveProgressForm(nil, 'Calcul des informations pour l''export GRH', 'Veuillez patienter SVP ...', j + 25, TRUE, TRUE);

  TRecap := TOB.create('Les RecapCp', nil, -1);
  Q := OpenSql('SELECT PRS_SALARIE,PRS_RESTN,(PRS_CUMRTTACQUIS-PRS_CUMRTTPRIS) SOLDERTT FROM RECAPSALARIES ORDER BY PRS_SALARIE', TRUE);
  TRecap.LoadDetailDB('RECAPSALARIES', '', '', Q, FALSE, False);
  Ferme(Q);
  MoveCurProgressForm('Chargement CP Ok'); // PORTAGECWAS
  // Boucle de chargement des TOB  sur historiques en fonction du SQL
  FileM := VH_Paie.PGCheminRech + '\ReportGrh.txt';
  if not FileExists(FileM) then
  begin
    PgiBox('Fichier inexistant  ' + FileM, Ecran.caption);
    RAZTOB();
    exit;
  end;
  AssignFile(FLect, FileM);
{$I-}Reset(FLect);
{$I+}if IoResult <> 0 then
  begin
    PGIBox('Fichier inaccessible : ' + FileM, 'Abandon du traitement');
    RAZTOB();
    Exit;
  end;
  i := 1;
  while not eof(FLect) do
  begin
{$I-}Readln(FLect, S);
{$I+}if IoResult <> 0 then
    begin
      PGIBox('Erreur de lecture du fichier : ' + FileM, 'Abandon du traitement');
      closeFile(FLect);
      Exit;
    end;
    if S = '' then break;
    Pref := '';
    if Pos('PHB', S) > 0 then Pref := 'PHB'
    else if Pos('PHC', S) > 0 then Pref := 'PHC';
    if Pref = '' then
    begin
      PgiBox('Erreur du contenu de la requête ' + S, Ecran.caption);
      RAZTOB();
      exit;
    end;
    TSQL[i] := TOB.create('La Rq SQL ' + Pref, nil, -1);
    if Pref = 'PHB' then
    begin
      st := 'select PHB_SALARIE,SUM (PHB_MTREM) MONTANT,SUM(PHB_BASEREM) BASE FROM HISTOBULLETIN ' +
        ' WHERE PHB_DATEDEBUT >= "' + UsDateTime(DD) + '" AND PHB_DATEFIN <= "' + UsDateTime(DF) + '" AND PHB_NATURERUB="AAA"' +
        S + ' GROUP BY PHB_SALARIE';
    end
    else
    begin
      st := 'select PHC_SALARIE,SUM (PHC_MONTANT) MONTANT FROM HISTOCUMSAL ' +
        ' WHERE PHC_DATEDEBUT >= "' + UsDateTime(DD) + '" AND PHC_DATEFIN <= "' + UsDateTime(DF) + '" ' +
        S + ' GROUP BY PHC_SALARIE';
    end;
    Q := OpenSql(st, TRUE);
    if Pref = 'PHB' then TSQL[i].LoadDetailDB('HISTOBULLETIN', '', '', Q, FALSE, False)
    else TSQL[i].LoadDetailDB('HISTOCUMSAL', '', '', Q, FALSE, False);
    Ferme(Q);
    MoveCurProgressForm('Chargement histo en cours ' + IntToStr(i)); // PORTAGECWAS
    i := i + 1;
    if i > 20 then break;
  end;
  closeFile(FLect);
  MoveCurProgressForm('Chargement histo Ok'); // PORTAGECWAS

  TTMP := TOB.create('Mon Report', nil, -1);
  T1 := Tsal.FindFirst([''], [''], FALSE);
  while T1 <> nil do
  begin
    T2 := TOB.Create('PGREPORT', TTMP, -1);
    for i := 1 to 47 do
    begin
      if Hist[i] = '' then break;
      T2.PutValue('PTM_' + Hist[i], T1.GetValue('PSA_' + Hist[i]));
    end;
    for i := 1 to 47 do
    begin
      if HistT[i] = '' then break;
      //          T2.PutValue ('PTM_'+HistT [i],RechDom ('PG'+HistT [i], T1.GetValue('PSA_'+HistT [i]),FALSE));
      T2.PutValue('PTM_' + HistT[i], T1.GetValue('PSA_' + HistT[i]));
    end;
    T2.putValue('PTM_DATEFIN', DF);
    T3 := Trecap.FindFirst(['PRS_SALARIE'], [T1.GetValue('PSA_SALARIE')], FALSE);
    if T3 <> nil then
    begin
      T2.putValue('PTM_RESTCPN', T3.getvalue('PRS_RESTN'));
      T2.putValue('PTM_SOLDERTT', T3.getvalue('SOLDERTT'));
    end;
    T2.PutValue('PTM_TYPCONTRAT', RendTypContrat(T1.GetValue('PSA_SALARIE')));
    // Traitement req SQL sur 1 historique
    for i := 1 to 20 do
    begin
      if TSQL[i] = nil then break;
      if Pos('PHB', TSQL[i].NomTable) > 0 then Pref := 'PHB'
      else Pref := 'PHC';
      T3 := TSQL[i].FindFirst([Pref + '_SALARIE'], [T1.GetValue(Pref + '_SALARIE')], FALSE);
      if T3 <> nil then
      begin
        T2.putValue('PTM_MT' + IntToStr(i), T3.getvalue('MONTANT'));
      end;
    end;

    TT.Value := TT.value + 1;
    TT.SubText := 'Traitement du salarié ' + T2.getValue('PTM_SALARIE');
    ;
    T1 := Tsal.FindNext([''], [''], FALSE);
  end;
  try
    BeginTrans;
    MoveCurProgressForm('Sauvegarde en cours ...'); // PORTAGECWAS
    TTMP.SetAllModifie(TRUE);
    TTMP.InsertOrUpdateDB(TRUE);
    MoveCurProgressForm('Traitement terminé ...'); // PORTAGECWAS
    CommitTrans;
  except
    Rollback;
    PGIBox('Une erreur est survenue lors de l''export', Ecran.Caption);
  end;
  RAZTOB();
end;

function TOF_PGExportGRH.RendTypContrat(LeSalarie: string; Quoi: string = ''): string;
var
  Q: TQuery;
begin
  result := '';
  Q := OpenSQL('SELECT PCI_TYPECONTRAT FROM CONTRATTRAVAIL ' + //PT20 On trie par ordre et non par code salarié
    'WHERE PCI_SALARIE = "' + LeSalarie + '" ORDER BY PCI_ORDRE DESC', TRUE);
  if not Q.Eof then
  begin
    if Quoi = '$' then result := RechDom('PGTYPECONTRAT', Q.findfield('PCI_TYPECONTRAT').AsString, FALSE)
    else result := Q.findfield('PCI_TYPECONTRAT').AsString;
  end;
  Ferme(Q);
end;

procedure TOF_PGExportGRH.OnArgument(Arguments: string);
var
  Num: Integer;
  BtnValidMul: TToolbarButton97;
begin
  inherited;
  for Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    if Num > 4 then Break;
    VisibiliteChampSalarie(IntToStr(Num), GetControl('PSA_TRAVAILN' + IntToStr(Num)), GetControl('TPSA_TRAVAILN' + IntToStr(Num)));
    VisibiliteChampLibreSal(IntToStr(Num), GetControl('PSA_LIBREPCMB' + IntToStr(Num)), GetControl('TPSA_LIBREPCMB' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PSA_CODESTAT'), GetControl('TPSA_CODESTAT'));

  BtnValidMul := TToolbarButton97(GetControl('BOuvrir'));
  if BtnValidMul <> nil then BtnValidMul.OnClick := LanceExport;
  Date1 := THEdit(GetControl('XX_VARIABLEDEB'));
  if Date1 <> nil then Date1.OnClick := DateDebutExit;
  Date2 := THEdit(GetControl('XX_VARIABLEFIN'));
  if Date2 <> nil then Date2.OnClick := DateFinExit;
end;

procedure TOF_PGExportGRH.RAZTOB;
var
  i: Integer;
begin
  inherited;
  if TSal <> nil then Tsal.free;
  if TRecap <> nil then TRecap.free;
  if TTMP <> nil then TTMP.free;
  FiniMoveProgressForm(); // PORTAGECWAS
  for i := 1 to 20 do
  begin
    if TSQL[i] <> nil then
    begin
      TSQL[i].free;
      TSQL[i] := nil;
    end;
  end;
end;

initialization
  registerclasses([TOF_PGExportGRH]);
end.

