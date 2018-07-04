{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/08/2003
Modifié le ... :   /  /
Description .. : Gestion : Import/export dossier
Mots clefs ... : PAIE;STANDARD
*****************************************************************}
{
PT1   : 07/06/2004 PH V_50 FQ 11576 Ergonomie message
PT2   : 07/07/2005 PH V_60 FQ 12430 ?? Typage de type var de l'objet TOZ passé en paramètre
PT3   : 04/08/2005 PH V_60 FQ 12244 Non traitement de la table ANNUAIRE en PCL
PT4   : 21/04/2006 PH V_65 FQ 12930 Controle des doublons salariés
PT5   : 14/03/2007 VG V_72 BQ_GENERAL n'est pas forcément unique
PT6   : 03/04/2007 PH V_72 FQ 13797 Alerte sur traitement des paramsoc de tous les produits
PT7   : 14/12/2007 PH V_80 FQ 15026 Compatibilité ORACLE
PT8   : 03/01/2008 PH V_80 FQ 14741 Exporter la totalité des CPs
}

unit UTofPG_GESTIONDOS;

interface
uses
  StdCtrls,
  Controls,
  Classes,
  Graphics,
  forms,
  sysutils,
  ComCtrls,
  HTB97,
{$IFNDEF EAGLCLIENT}
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
  HDebug,
  HCtrls,
  HEnt1,
  HMsgBox,
{$IFNDEF EAGLSERVER}
  UTOF,
{$ENDIF}
  UTOZ,
  Ed_Tools,
  Entpaie,
  UTOB,
  PgOutilsTreso;

{$IFNDEF EAGLSERVER}
type
  TOF_PG_GESTIONDOS = class(TOF)
  private
    procedure LANCECOPIE(Sender: TObject);
    procedure ChangeAction(Sender: TObject);
    procedure ChangeContenu(Sender: TObject);
    procedure ParamGestDOS(Sender: TObject);
  public
    procedure OnArgument(Arguments: string); override;
  end;
{$ENDIF}
implementation

uses
  P5def,
  PgOutils,
  PgOutils2;

procedure PGFermeZipFile(TheTOZ: TOZ);
begin
  TheToz.CloseSession;
end;

procedure PGVideDirectory(FileN: string);
var
  sr: TsearchRec;
  ret: Integer;
  MonFic: string;
begin

  ret := FindFirst(FileN + '\*.TPP', 0 + faAnyFile, sr);
  while ret = 0 do
  begin
    MonFic := FileN + '\' + sr.Name;
    if FileExists(MonFic) then DeleteFile(PChar(MonFic));
    ret := FindNext(sr);
  end;
  sysutils.FindClose(sr);
end;

function RendSQLTablePaie(Trait: string; Etab, Cal, Tiers, Annu, BqCp: Boolean): string;
begin
  result :=
    'SELECT DT_NOMTABLE,DT_PREFIXE,DT_LIBELLE,DT_NUMVERSION FROM DETABLES WHERE (DT_DOMAINE = "P" AND NOT EXISTS (SELECT DH_NOMCHAMP FROM DECHAMPS WHERE DH_NOMCHAMP=DT_PREFIXE||"_PREDEFINI"))';
  result := result + ' OR (DT_NOMTABLE = "CHOIXCOD") OR (DT_NOMTABLE = "PARAMSOC")';
  if Trait = 'E' then // export
  begin
    if Etab then result := result + ' OR (DT_NOMTABLE = "ETABLISS")';
    if Cal then result := result + ' OR (DT_NOMTABLE = "CALENDRIER") OR (DT_NOMTABLE = "JOURFERIE")';
    if Tiers then result := result + ' OR (DT_NOMTABLE = "TIERS") OR (DT_NOMTABLE = "RIB")';
    if Annu then result := result + ' OR (DT_NOMTABLE = "ANNUAIRE")';
    if BqCp then result := result + ' OR (DT_NOMTABLE = "BANQUECP")';
    if VH_Paie.PGAnalytique then result := result + ' OR (DT_NOMTABLE = "VENTIL") OR (DT_NOMTABLE = "VENTANA") OR (DT_NOMTABLE = "AXE") OR (DT_NOMTABLE = "SECTION")';
  end;
  result := result + ' ORDER BY DT_NOMTABLE';
end;

// Fonction qui parcourt les tob pour réaffecter le bon numero de dossier
// en restauration - Uniquement en environnement MULTI

procedure RetraiteNoDossier(T_Mere: TOB);
var
  i, j, k: Integer;
  T1, T2: TOB;
  Prefix, NomT: string;
begin
  Debug('Retraite numéro dossier ' + IntToStr(T_Mere.detail.count));

  for i := 0 to T_Mere.detail.count - 1 do
  begin
    T1 := T_Mere.detail[i];
    for j := 0 to T1.detail.count - 1 do
    begin
      T2 := T1.detail[j];
      NomT := T2.NomTable;
      k := tableToNum(NomT);
      if V_PGI.DETables[k].TypeDP = tdDPSTD then
      begin
        Prefix := TableToPrefixe(NomT);
        Debug('Nom table ' + T2.NomTable + ' préfixe ' + Prefix);
        if T2.NomTable = 'JOURFERIE' then T2.PutValue(Prefix + '_NODOSSIER', PgRendNoDossier());
      end;
    end;
  end;

end;

function PGZipFile(Fichier: string; TheTOZ: TOZ; Ecran: TFORM): Boolean;
begin
  result := false;
  try
    if TheToz.ProcessFile(Fichier, '') then result := TRUE;
  except
    on E: Exception do
    begin
      PgiError('TozError : ' + E.Message, Ecran.caption);
      TheToz.Free;
    end;
  end;
end;
// PT2 La TOZ est passée en VAR ???

function PGCreatZipFile(Archive: string; CodeSession: TModeOpenZip; var TheTOZ: TOZ; Ecran: TFORM): Boolean;
begin
  result := false;
  TheToz := TOZ.Create;
  try
    if TheToz.OpenZipFile(Archive, CodeSession) then
    begin
      if CodeSession = moCreate then
        if TheToz.OpenSession(osAdd) then result := TRUE
        else HShowMessage('0;Erreur;Soit le fichier : ' + Archive + ' n''existe plus, soit la session n''est pas ouverte en ajout.;E;O;O;O', '', '');

      if CodeSession = moOpen then
        if TheToz.OpenSession(osExt) then result := TRUE
        else HShowMessage('0;Erreur;Soit le fichier : ' + Archive + ' n''existe plus, soit la session n''est pas ouverte en ajout.;E;O;O;O', '', '');
    end
    else
    begin
      HShowMessage('0;Erreur;Erreur création du fichier archive : ' + Archive + ' impossible;E;O;O;O', '', '');
      Exit;
    end;
  except
    on E: Exception do
    begin
      PgiError('TozError : ' + E.Message, Ecran.Caption);
      TheToz.Free;
    end;
  end;
end;

procedure PGExportDOS(Ecran: TFORM; CAL, ETAB, MVT, TIERS, ANNU, BQCP, PSOC, MVTSEUL, BINAIRE: Boolean;
  FileN, Lef: string; LaDate: TDateTime; NMois: WORD; CPGLOB: Boolean);
var
  T_Mere, T, T1, T_Fic: TOB;
  St, St1, StPlus, Stq, LeNom, LeMois: string;
  Q: TQuery;
  Rep, i, Nbre, j, NombreDeMois: integer;
  OkOk, ExportStd, PMois: Boolean;
  ExportOk, OkSQL: Boolean;
  MaTable, ChampMvt, MonFic: string;
  TheToz: TOZ;
{$IFNDEF EAGLSERVER}
  Trace: TListBox;
{$ELSE}
  Trace: TStringList;
{$ENDIF}
  LaDateFin, LaDateDeb: TDateTime;
  Annee, Mois, Jour: WORD;
begin
  T_Mere := TOB.Create('Ma TOB', nil, -1);
  T := TOB.Create('La TOB des tables', nil, -1);
  st := RendSQLTablePaie('E', ETAB, CAL, TIERS, ANNU, BQCP);
  Q := OPENSQL(ST, TRUE);
  T.LoadDetailDb('DETABLES', '', '', Q, FALSE);
  Ferme(Q);
{$IFNDEF EAGLSERVER}
  Trace := TListBox(PGGetControl('LSTBXTRACE', Ecran));
  Trace.Items.Add('Les éléments suivants seront analysés et pris en compte s''ils comportent des informations à exporter :');
  for i := 0 to T.detail.count - 1 do
  begin
    T1 := T.Detail[i];
    if T1 <> nil then
    begin
      st := 'SELECT COUNT (*) NBRE FROM ' + T1.GetValue('DT_NOMTABLE');
      Q := OPENSQL(ST, TRUE);
      if not Q.EOF then
      begin
        Nbre := Q.FindField('NBRE').AsInteger;
        if Nbre > 0 then
{$IFNDEF EAGLSERVER}
          Trace.Items.Add('Table ' + T1.GetValue('DT_LIBELLE'));
{$ELSE}
          Trace.Add('Table ' + T1.GetValue('DT_LIBELLE'));
{$ENDIF}
      end;
      Ferme(Q);
    end;
  end;
{$ENDIF}
  TheTOZ := nil;
  if not PGCreatZipFile(FileN + '\' + LeF + '.zip', moCreate, TheTOZ, Ecran) then
  begin
    T.free;
    T_Mere.free;
{$IFDEF EAGLSERVER}
    Trace.Add('Erreur sur fichier archive : ' + FileN + '\' + LeF + '.zip');
    CreeJnalEvt('003', '200', 'ERR', nil, nil, Trace, nil);
{$ENDIF}
    exit;
  end;
{$IFNDEF EAGLSERVER}
  Rep := PgiAsk('Voulez vous continuer le traitement', Ecran.Caption);
  if Rep <> MrYes then
  begin
    T.free;
    T_Mere.free;
    exit;
  end;
  InitMoveProgressForm(nil, 'Lecture des données', 'Veuillez patienter SVP ...', T.detail.count, FALSE, TRUE);
{$ENDIF}
  ExportOk := TRUE;
  for i := 0 to T.detail.count - 1 do
  begin
    T1 := T.Detail[i];
    if T1 <> nil then
    begin
      MaTable := T1.GetValue('DT_NOMTABLE');
      if T1.GetValue('DT_NOMTABLE') = 'PARAMSOC' then
      begin
        if PSOC then st := 'SELECT * FROM PARAMSOC '
        else st := 'SELECT * FROM PARAMSOC WHERE SOC_NOM LIKE "SO_PG%" OR SOC_NOM LIKE "SCO_PG%"';
      end
      else if T1.GetValue('DT_NOMTABLE') = 'CHOIXCOD' then st := 'SELECT CHOIXCOD.* FROM DECOMBOS,CHOIXCOD WHERE DO_DOMAINE="P" AND DO_TYPE=CC_TYPE AND CC_TYPE LIKE "P%"'
      else if T1.GetValue('DT_NOMTABLE') = 'JOURFERIE' then st := 'SELECT * FROM JOURFERIE WHERE ##AJF_PREDEFINI##'
      else if T1.GetValue('DT_NOMTABLE') = 'AXE' then st := 'SELECT * FROM AXE'
      else if T1.GetValue('DT_NOMTABLE') = 'SECTION' then st := 'SELECT * FROM SECTION'
      else if T1.GetValue('DT_NOMTABLE') = 'VENTIL' then st := 'SELECT * FROM VENTIL'
      else if T1.GetValue('DT_NOMTABLE') = 'VENTANA' then st := 'SELECT * FROM VENTANA WHERE YVA_TABLEANA = "PPU"'
      else if T1.GetValue('DT_NOMTABLE') = 'TIERS' then st := 'SELECT * FROM TIERS WHERE T_NATUREAUXI = "SAL"'
      else if T1.GetValue('DT_NOMTABLE') = 'RIB' then st := 'SELECT * FROM RIB LEFT JOIN TIERS ON R_AUXILIAIRE = T_AUXILIAIRE AND T_NATUREAUXI = "SAL"'
      else if T1.GetValue('DT_NOMTABLE') = 'ANNUAIRE' then st := 'SELECT * FROM ANNUAIRE WHERE ANN_FAMPER = "SOC"'
      else if T1.GetValue('DT_NOMTABLE') = 'BANQUECP' then
      begin
        StPlus := PGBanqueCP(True); //PT5

        st := 'SELECT * FROM BANQUECP WHERE (BQ_GENERAL IN (SELECT ETB_RIBSALAIRE FROM ETABCOMPL)) OR (BQ_GENERAL in (SELECT ETB_RIBACOMPTE FROM ETABCOMPL))'
          + ' OR (BQ_GENERAL in (SELECT ETB_RIBDUCSEDI FROM ETABCOMPL)) OR (BQ_GENERAL in (SELECT ETB_RIBFRAIS FROM ETABCOMPL))' + StPlus;
      end
      else st := 'SELECT * FROM ' + T1.GetValue('DT_NOMTABLE');
      ChampMvt := 'SELECT DH_PREFIXE FROM DECHAMPS WHERE DH_NOMCHAMP = "' + T1.GetValue('DT_PREFIXE') + '_DATEDEBUT"';
      if MVTSEUL then st := '';
      if (ExisteSQL(ChampMvt)) and (MVT) or (T1.GetValue('DT_NOMTABLE') = 'VENTANA') then
      begin
        if (T1.GetValue('DT_NOMTABLE') <> 'VENTANA') then st := 'SELECT * FROM ' + T1.GetValue('DT_NOMTABLE')
        else st := 'SELECT * FROM VENTANA WHERE YVA_TABLEANA = "PPU"';
        PMois := TRUE;
      end
      else PMois := FALSE;
      if T1.GetValue('DT_NOMTABLE') = 'VENTANA' then PMois := TRUE;
      if (Copy(T1.GetValue('DT_NOMTABLE'), 1, 4) = 'EXER') then PMois := FALSE; // tables des exercices qui ne sont pas des historiques
      if PMois then NombreDeMois := NMois
      else NombreDeMois := 1;
      for j := 1 to NombreDeMois do
      begin
        if Pmois and ((T1.GetValue('DT_NOMTABLE') <> 'VENTANA') and (T1.GetValue('DT_NOMTABLE') <> 'ABSENCESALARIE')) then
        begin
          LaDateDeb := PLUSMOIS(LaDate, j - 1);
          LaDateFin := FINDEMOIS(LaDateDeb);
          st1 := ' WHERE ' + T1.GetValue('DT_PREFIXE') + '_DATEFIN >= "' + UsDateTime(DebutDeMois(LaDateFin)) +
            '" AND ' + T1.GetValue('DT_PREFIXE') + '_DATEFIN <= "' + UsDateTime(LaDateFin) + '"';
          stq := st + St1;
        end
        else
          if T1.GetValue('DT_NOMTABLE') = 'VENTANA' then
          begin
            DecodeDate(PLUSMOIS(LaDate, j - 1), Annee, Mois, Jour);
            if Mois >= 10 then LeMois := IntToStr(Mois)
            else LeMois := '0' + IntToStr(Mois);
            LeMois := LeMois + IntToStr(Annee);
            stq := st + ' AND YVA_IDENTIFIANT LIKE "%' + LeMois + '%"';
          end
          else
            if T1.GetValue('DT_NOMTABLE') = 'ABSENCESALARIE' then
            begin
              if not CPGLOB then  // PT8
              begin
                LaDateDeb := PLUSMOIS(LaDate, j - 1);
                LaDateFin := FINDEMOIS(LaDateDeb);
                st1 := ' WHERE ' + T1.GetValue('DT_PREFIXE') + '_DATEVALIDITE >= "' + UsDateTime(DebutDeMois(LaDateFin)) +
                  '" AND ' + T1.GetValue('DT_PREFIXE') + '_DATEVALIDITE <= "' + UsDateTime(LaDateFin) + '"';
                stq := st + St1;
              end
              else
              begin
                stq := st; // PT8
              end;
            end
            else stq := st;

        if (stq <> '') then
        begin // DEB PT7
          if (MaTable <> 'JOURFERIE') and (MaTable <> 'RIB') then
            okSQL := ExisteSQL(stq)
          else OkSQL := true;
          if OkSQL then
          begin // FIN PT7
{$IFNDEF EAGLSERVER}
            if not PMois then MoveCurProgressForm('Traitement de la table ' + T1.GetValue('DT_NOMTABLE'))
            else MoveCurProgressForm('Traitement de la table ' + T1.GetValue('DT_NOMTABLE') + ' mois ' + IntToStr(j));
{$ENDIF}
            Q := OPENSQL(stq, TRUE);
            T_Mere := TOB.Create('Ma tob', nil, -1);
            LeNom := T1.GetValue('DT_NOMTABLE') + '_';
            T_Fic := TOB.Create(LeNom, T_Mere, -1);
            T_Fic.AddChampSup('VERSION', FALSE);
            T_Fic.PutValue('VERSION', T1.GetValue('DT_NUMVERSION'));
            T_Fic.LoadDetailDb(T1.GetValue('DT_NOMTABLE'), '', '', Q, FALSE);
            Ferme(Q);
            MonFic := FileN + '\' + T1.GetValue('DT_NOMTABLE') + '.TPP';
            if (LeNom = 'PARAMSOC_') and PSoc then
              MonFic := FileN + '\' + 'PARAMSOC_COMPLET' + '.TPP'; // PT6
            if Pmois then MonFic := MonFic + IntToStr(j);
          //          LeNumFic := LeNumFic + 1;
            if FileExists(MonFic) then DeleteFile(PChar(MonFic));
            if BINAIRE then T_Fic.SaveToBinFile(MonFic, FALSE, TRUE, TRUE, FALSE)
            else T_Fic.SaveToFile(MonFic, FALSE, TRUE, TRUE);
            FreeAndNil(T_Fic);
            FreeAndNil(T_Mere);
            if not PGZipFile(MonFic, TheTOZ, Ecran) then
            begin
              Exportok := FALSE;
              break;
            end;
            if (T1.GetValue('DT_NOMTABLE') = 'ABSENCESALARIE') AND (CPGLOB) then    // PT8
              break; // Pour forcer de traiter en une fois la table absences si CPGLOB
          end; // Fin si requête contient des enregistrements à exporter
        end; // Fin si requête remplie
      end; // Fin de la boucle sur le nombre de mois
    end;
  end;
{$IFNDEF EAGLSERVER}
  MoveCurProgressForm('Fin de traitement des données');
{$ENDIF}
  if not Exportok then PgiError('Le traitement est abandonné', Ecran.caption)
  else PGFermeZipFile(TheTOZ);
  PGVideDirectory(FileN);
{$IFNDEF EAGLSERVER}
  MoveCurProgressForm('Fin de traitement des données');
{$ENDIF}
  if Assigned(TheToz) then TheToz.Free;
  try
{$IFNDEF EAGLSERVER}
    Trace.Items.add('Ecriture du fichier OK');
    CreeJnalEvt('003', '200', 'OK', Trace, nil, nil);
{$ELSE}
    Trace.add('Ecriture du fichier OK');
    CreeJnalEvt('003', '200', 'OK', nil, nil, Trace, nil);
{$ENDIF}

  except
{$IFNDEF EAGLSERVER}
    PGIBox('Une erreur est survenue lors de l''écriture du fichier', Ecran.caption);
    Trace.Items.add('Une erreur est survenue lors de l''écriture du fichier');
    CreeJnalEvt('003', '200', 'ERR', Trace, nil, nil);
{$ELSE}
    Trace.add('Une erreur est survenue lors de l''écriture du fichier');
    CreeJnalEvt('003', '200', 'ERR', nil, nil, Trace, nil);
{$ENDIF}
  end;
{$IFNDEF EAGLSERVER}
  FiniMoveProgressForm();
{$ENDIF}
  T.free;
  T_Mere.free;
{$IFNDEF EAGLSERVER}
  PgiBox('Exportation terminée', Ecran.caption);
{$ENDIF}
end;
// DEB PT4
{$IFNDEF EAGLSERVER}

function VerifDoublonSal(TT, LaTobSal: TOB; Trace: TListbox): boolean;
{$ELSE}

function VerifDoublonSal(TT, LaTobSal: TOB; Trace: TStringList): boolean;
{$ENDIF}
var i: Integer;
  T1, T2: TOB;
  LeSal: string;
begin
  result := FALSE;
  for i := 0 to TT.detail.count - 1 do
  begin
    T1 := TT.detail[i];
    LeSal := T1.GetValue('PSA_SALARIE');
    T2 := LaTobSal.Findfirst(['PSA_SALARIE'], [LeSal], FALSE);
    if T2 <> nil then // DOUBLON de matricule
    begin
      if (T1.GetValue('PSA_NUMEROSS')) <> (T2.GetValue('PSA_NUMEROSS')) then
{$IFNDEF EAGLSERVER}
        Trace.Items.Add('Salarié en doublon avec N° SS différent : ' + LeSal)
      else Trace.Items.Add('Salarié en doublon : ' + LeSal);
{$ELSE}
        Trace.Add('Salarié en doublon avec N° SS différent : ' + LeSal)
      else Trace.Add('Salarié en doublon : ' + LeSal);
{$ENDIF}
      result := TRUE;
    end;
  end;
end;
// FIN PT4

procedure PGImportDOS(Ecran: TFORM; BINAIRE: Boolean; FileN: string);
var
  Ind, rep, rep1, Numvers, ret, LaVersion: Integer;
  T_Mere, T, T1, T2, LaTobSal: TOB;
  Q: TQUERY;
  Repert, LeNom, Libel, st, MonFic: string;
  TheTOZ: TOZ;
  sr: TsearchRec;
{$IFNDEF EAGLSERVER}
  Trace: TListBox;
{$ELSE}
  Trace: TStringList;
{$ENDIF}
  OkOk, AnnuP, Doublon, ParmC: Boolean; // PT3 // PT6
begin
  Ind := Strlen(PChar(FileN));
  while Ind > 0 do
  begin
    if FileN[Ind] = '\' then break;
    Ind := Ind - 1;
  end;
  Repert := Copy(FileN, 1, Ind);
  TheTOZ := nil;
  if (Pos('.zip', FileN) > 0) or (Pos('.ZIP', FileN) > 0) then
  begin
    if not PGCreatZipFile(FileN, moOpen, TheTOZ, Ecran) then
    begin
{$IFNDEF EAGLSERVER}
      PgiError('Traitement abandonné', Ecran.caption);
{$ELSE}
      Trace.Add('Erreur sur fichier archive : ' + FileN);
      CreeJnalEvt('003', '200', 'ERR', nil, nil, Trace, nil);
{$ENDIF}
      exit;
    end
    else PGFermeZipFile(TheTOZ);
  end;
  T := TOB.Create('La TOB des tables', nil, -1);
  ST := RendSQLTablePaie('I', True, True, TRUE, True, TRUE); // On prend tjrs calendriers ou etabliss ou ...
  Q := OPENSQL(ST, TRUE);
  T.LoadDetailDb('DETABLES', '', '', Q, FALSE);
  Ferme(Q);
{$IFNDEF EAGLSERVER}
  Trace := TListBox(PGGetControl('LSTBXTRACE', Ecran));
  InitMoveProgressForm(nil, 'Analyse des données', 'Veuillez patienter SVP ...', 2, FALSE, TRUE);
{$ENDIF}
  ret := FindFirst(Repert + '*.TPP', 0 + faAnyFile, sr);
  Rep := MrYes;
  AnnuP := FALSE; // PT3
  LaTobSal := TOB.Create('Les Salaries présents', nil, -1);
  Q := OpenSQL('SELECT PSA_SALARIE,PSA_NUMEROSS FROM SALARIES', TRUE);
  LaTobSal.LoadDetailDB('SALARIES', '', '', Q, FALSE);
  FERME(Q);
  while (ret = 0) and (rep = mrYes) do
  begin
    MonFic := Repert + sr.Name;
    T_Mere := TOB.Create('Ma TOB', nil, -1);
{$IFNDEF EAGLSERVER}
    MoveCurProgressForm('Analyse des données du fichier ' + sr.Name);
{$ENDIF}
    // DEB PT6
    if POS('PARAMSOC_COMPLET.TPP', MonFic) > 0 then
      ParmC := TRUE
    else ParmC := FALSE;
    // FINPT6
    if BINAIRE then TOBLoadFromBinFile(MonFic, nil, T_Mere)
    else TOBLoadFromFile(MonFic, nil, T_Mere);
    T1 := T_Mere.Detail[0];
    if Assigned(T1) then
    begin
      LeNom := T1.NomTable;
      if LeNom = 'ANNUAIRE' then AnnuP := TRUE; // PT3
      Libel := Copy(LeNom, 1, StrLen(PChar(LeNom)) - 1);
      LaVersion := T1.GetValue('VERSION');
      T2 := T.FindFirst(['DT_NOMTABLE'], [Libel], FALSE);
      if T2 <> nil then NumVers := T2.GetValue('DT_NUMVERSION')
      else NumVers := 0;
      //        Trace.Items.Add('Table '+LeNom);
      if (LaVersion <> NumVers) and (NumVers <> 0) then
      begin
{$IFNDEF EAGLSERVER}
        Trace.Items.Add('Version de table ' + libel + ' incompatible export=' + IntToStr(LaVersion) + ' import=' + IntToStr(NumVers));
{$ELSE}
        Trace.Add('Version de table ' + libel + ' incompatible export=' + IntToStr(LaVersion) + ' import=' + IntToStr(NumVers));
{$ENDIF}
        rep := MrCancel;
      end;
      if (POS('SALARIES', LeNom) > 0) then Doublon := VerifDoublonSal(T1, LaTobSal, Trace); // PT4
    end;
    ret := FindNext(sr);
    FreeAndNil(T_Mere);
  end;
  if Assigned(T_Mere) then FreeAndNil(T_Mere);
  if Assigned(LaTOBSal) then FreeAndNil(LaTOBSal);
  sysutils.FindClose(sr);
  // DEB PT3
{$IFNDEF EAGLSERVER}
  if AnnuP and (V_PGI.ModePcl = '1') then
{ELSE}
    if AnnuP then
{$ENDIF}
    begin
{$IFNDEF EAGLSERVER}
      Rep := PgiAsk('Attention, votre fichier d''import contient des données concernant l''annuaire. #13#10'
        +
        'Voulez-vous continuer le traitement ?', Ecran.Caption);
      if Rep = MrYes then
      begin
        Rep := PgiAsk('Le traitement continue sans intégrer les données de l''annuaire. #13#10'
          +
          'Voulez-vous poursuivre le traitement ?', Ecran.Caption);
      end;
{$ELSE}
      Trace.Add('La table ANNUAIRE ne sera pas traitée.');
{$ENDIF}
    end;
  // FIN PT3
// DEB PT6
{$IFNDEF EAGLSERVER}
  if ParmC then
  begin
    Rep := PgiAsk('Attention, votre fichier d''import contient tous les paramètres société. #13#10'
      +
      'Voulez-vous continuer le traitement ?', Ecran.Caption);
    if Rep = MrYes then
    begin
      Rep := PgiAsk('Le traitement continue avec intégration de tous les paramètres société. #13#10'
        +
        'Voulez-vous poursuivre le traitement ?', Ecran.Caption);
    end;
  end;
{$ELSE}
  Trace.Add('L''import contient tous les paramètres société');
{$ENDIF}
  // FIN PT6
  if Rep = MrYes then
  begin
{$IFNDEF EAGLSERVER}
    Rep := PgiAsk('Aucune anomalie de structure détectée.#13#10' + 'Si une anomalie se produit pendant le traitement, vous devez impérativement restaurer votre sauvegarde.#13#10'
      +
      'Voulez-vous continuer le traitement ?', Ecran.Caption);
{$ELSE}
    Trace.Add('Aucune anomalie de structure détectée.');
{$ENDIF}
  end
  else
  begin
{$IFNDEF EAGLSERVER}
    PgiError('Des anomalies de structure ont été détectées ou bien vous avez abandonné le traitement', Ecran.caption);
{$ELSE}
    Trace.Add('Des anomalies de structure ont été détectées, Abandon du traitement');
{$ENDIF}
  end;
  // DEB PT4
  if Doublon and (Rep = MrYes) then
  begin
    Rep1 := PgiAsk('Attention, vous avez des doublons de matricule salarié.#13#10' + 'Si vous continuez le traitement, les données de paies seront incohérentes.'
      +
      'Voulez-vous abandonner le traitement ?', Ecran.Caption);
    if Rep1 = MrYes then Rep := MrNO;
  end;
  // FIN PT4
  if Rep <> MrYes then
  begin
{$IFNDEF EAGLSERVER}
    FiniMoveProgressForm();
    Trace.Items.Add('Traitement abandonné');
{$ELSE}
    Trace.Add('Traitement abandonné');
{$ENDIF}
    exit;
  end;
  ret := FindFirst(Repert + '*.TPP', 0 + faAnyFile, sr);
  OkOk := TRUE;
  while (ret = 0) and (OkOk) do
  begin
    if Assigned(T_Mere) then FreeAndNil(T_Mere);
    MonFic := Repert + sr.Name;
    T_Mere := TOB.Create('Ma TOB', nil, -1);
    if BINAIRE then TOBLoadFromBinFile(MonFic, nil, T_Mere)
    else TOBLoadFromFile(MonFic, nil, T_Mere);
    T1 := T_Mere.Detail[0];
    // DEB PT3
{$IFNDEF EAGLSERVER}
    if Assigned(T1) and AnnuP and (V_PGI.ModePcl = '1') then
{$ELSE}
    if Assigned(T1) and AnnuP then
{$ENDIF}
    begin
      LeNom := T1.NomTable;
      if LeNom = 'ANNUAIRE' then
      begin
        ret := FindNext(sr);
        if Assigned(T_Mere) then FreeAndNil(T_Mere);
      end;
    end;
    // FIN PT3
{$IFNDEF EAGLSERVER}
    MoveCurProgressForm('Traitement des données du fichier ' + sr.Name);
{$ENDIF}
    Debug('Mode : ' + PgRendModeFonc() + ' Dossier = ' + PgRendNoDossier());
    //      if Uppercase(PgRendModeFonc()) = 'MULTI' then
    RetraiteNoDossier(T_Mere);
    try
      OkOk := T_Mere.InsertOrUpdateDb(TRUE);
{$IFNDEF EAGLSERVER}
      Trace.Items.add('Traitement du fichier ' + sr.Name + ' OK');
{$ELSE}
      Trace.add('Traitement du fichier ' + sr.Name + ' OK');
{$ENDIF}
    except
{$IFNDEF EAGLSERVER}
      Trace.Items.add('Erreur de traitement des données du fichier ' + sr.Name);
      FiniMoveProgressForm();
{$ELSE}
      Trace.add('Erreur de traitement des données du fichier ' + sr.Name);
{$ENDIF}
      FreeAndNil(T_Mere);
      sysutils.FindClose(sr);
{$IFNDEF EAGLSERVER}
      PgiError('Le traitement est interrompu, votre base de données est inutilisable !', Ecran.Caption);
      CreeJnalEvt('003', '201', 'ERR', Trace, nil, nil);
{$ELSE}
      Trace.add('Le traitement est interrompu, votre base de données est inutilisable !');
      CreeJnalEvt('003', '201', 'ERR', nil, nil, Trace, nil);
{$ENDIF}
      OkOk := FALSE;
    end;
    ret := FindNext(sr);
    if Assigned(T_Mere) then FreeAndNil(T_Mere);
  end;
  if Assigned(T_Mere) then FreeAndNil(T_Mere);
{$IFNDEF EAGLSERVER}
  FiniMoveProgressForm();
{$ENDIF}
  sysutils.FindClose(sr);
  if OkOk then
{$IFNDEF EAGLSERVER}
    CreeJnalEvt('003', '201', 'OK', Trace, nil, nil);
{$ELSE}
    CreeJnalEvt('003', '201', 'OK', nil, nil, Trace, nil);
{$ENDIF}
  Repert := Copy(Repert, 1, Strlen(PChar(Repert)) - 1);
  PGVideDirectory(Repert);
{$IFNDEF EAGLSERVER}
  PgiBox('Import terminé', Ecran.caption);
{$ENDIF}
end;

{$IFNDEF EAGLSERVER}

procedure TOF_PG_GESTIONDOS.ChangeAction(Sender: TObject);
var
  Coche: TRadioButton;
begin
  Coche := TRadioButton(GetControl('RDTEXPORT'));
  if Coche = nil then exit;
  if Coche.Checked then
  begin
    SetControlEnabled('IMPORTFIC', FALSE);
    SetControlEnabled('EXPORTFIC', TRUE);
    SetControlEnabled('CHBXETAB', TRUE);
    SetControlEnabled('CHBXJOURFERIES', TRUE);
  end
  else
  begin
    SetControlEnabled('IMPORTFIC', TRUE);
    SetControlEnabled('EXPORTFIC', FALSE);
    SetControlEnabled('CHBXETAB', FALSE);
    SetControlEnabled('CHBXJOURFERIES', FALSE);
  end;
end;

procedure TOF_PG_GESTIONDOS.LANCECOPIE(Sender: TObject);
var
  rep: Integer;
  Q: TQuery;
  ExportStd: Boolean;
  RDT: TRadioButton;
  FileN: string;
  Trace: TListBox;
  Pan: TPageControl;
  Tbsht: TTabSheet;
  CAL, ETAB, MVT, TIERS, ANNU, BQCP, PSOC, MVTSEUL, BINAIRE, CPGLOB: Boolean; // PT8
  NbS, NbE, NbP: Integer;
  LeF: string;
  Annee, Mois, NMois: WORD;
  LaDateDeb: TDateTime;
begin
  Trace := TListBox(GetControl('LSTBXTRACE'));
  if (Trace = nil) then
  begin
    PGIBox('Traitement impossible car le composant trace est indisponible', Ecran.Caption);
    exit;
  end;
  Trace.Clear;
  ExportStd := FALSE;
  RDT := TRadioButton(GetControl('RDTEXPORT'));
  if RDT <> nil then
  begin
    if RDT.Checked then ExportStd := TRUE;
  end;
  if ExportStd then
  begin
    CAL := (GetCheckBoxState('CHBXJOURFERIES') = cbChecked);
    ETAB := (GetCheckBoxState('CHBXETAB') = cbChecked);
    MVT := (GetCheckBoxState('CHBXMVT') = cbChecked);
    TIERS := (GetCheckBoxState('CHBXTIERS') = cbChecked);
    ANNU := (GetCheckBoxState('CHBXANNU') = cbChecked);
    BQCP := (GetCheckBoxState('CHBXBQCP') = cbChecked);
    PSOC := (GetCheckBoxState('CHBXPARAMSOC') = cbChecked);
    MVTSEUL := (GetCheckBoxState('CHBXMVTSEUL') = cbChecked);

    FileN := GetControlText('EXPORTFIC');
    LeF := GetControlText('NOMFIC');

    if FileN = '' then
    begin
      PgiBox('Vous n''avez pas renseigner le réperoire de transfert !', Ecran.caption);
      exit;
    end;
    if LeF = '' then
    begin
      PgiBox('Vous n''avez pas renseigner le nom du fichier archive !', Ecran.caption);
      exit;
    end;
    NOMBREMOIS(StrToDate(GetControlText('DateDebut')), StrToDate(GetControlText('DateFin')), Mois, Annee, NMois);
    if NMois > 12 then
    begin
      PgiBox('Vous ne pouvez pas exporter sur une période supérieure à un an !', Ecran.caption);
      exit;
    end;
    Pan := TPageControl(GetControl('PGCTRL'));
    Tbsht := TTabSheet(GetControl('TBSHTCONTROL'));
    if (Pan <> nil) and (Tbsht <> nil) then Pan.ActivePage := Tbsht;
    LaDateDeb := StrToDate(GetControlText('DATEDEBUT'));
    BINAIRE := (GetCheckBoxState('CHBXBINAIRE1') = cbChecked);
    CPGLOB := (GetCheckBoxState('CHBXCPGLOB') = cbChecked); // PT8
    PGExportDOS(Ecran, CAL, ETAB, MVT, TIERS, ANNU, BQCP, PSOC, MVTSEUL, BINAIRE, FileN, Lef, LaDateDeb, NMois, CPGLOB); // PT8
  end
  else
  begin
    Q := OPENSQL('SELECT COUNT (*) FROM SALARIES', TRUE);
    if not Q.EOF then NbS := Q.Fields[0].AsInteger
    else NbS := 0;
    Ferme(Q);
    Q := OPENSQL('SELECT COUNT (*) FROM PAIEENCOURS', TRUE);
    if not Q.EOF then NbP := Q.Fields[0].AsInteger
    else NbP := 0;
    Ferme(Q);
    Q := OPENSQL('SELECT COUNT (*) FROM ETABLISS', TRUE);
    if not Q.EOF then NbE := Q.Fields[0].AsInteger
    else NbE := 0;
    Ferme(Q);
    if (NbS <> 0) or (NbP <> 0) or (NbE <> 0) then
    begin
      rep := PgiAsk('Attention, vous avez fait ' + IntToStr(NbP) + ' paies pour ' + IntToStr(NbE) + ' établissements concernant ' + IntToStr(NbS) +
        ' salariés.#13#10Vos données seront insérées ou écraseront celles déjà existantes.#13#10' +
        'Vous devez impérativement avoir une sauvegarde de votre base à jour des dernières modifications,#13#10' +
        'Voulez vous continuer le traitement ?', Ecran.caption); // PT1
      if Rep <> MrYes then exit;
    end;
    FileN := GetControlText('IMPORTFIC');
    if FileN = '' then
    begin
      PgiError('Vous n''avez pas renseigné de fichier à importer !', Ecran.caption);
      exit;
    end;
    BINAIRE := (GetCheckBoxState('CHBXBINAIRE') = cbChecked);
    Pan := TPageControl(GetControl('PGCTRL'));
    Tbsht := TTabSheet(GetControl('TBSHTCONTROL'));
    if (Pan <> nil) and (Tbsht <> nil) then Pan.ActivePage := Tbsht;
    PGIMPORTDOS(Ecran, BINAIRE, FileN);
  end;

end;

procedure TOF_PG_GESTIONDOS.OnArgument(Arguments: string);
var
  BtnVal: TToolbarButton97;
  Coche: TRadioButton;
  Mvts: TCheckBox;
begin
  inherited;
  BtnVal := TToolbarButton97(GetControl('BValider'));
  if BtnVal <> nil then BtnVal.OnClick := LANCECOPIE;
  Coche := TRadioButton(GetControl('RDTEXPORT'));
  if Coche <> nil then Coche.OnClick := ChangeAction;
  Coche := TRadioButton(GetControl('RDTIMPORT'));
  if Coche <> nil then Coche.OnClick := ChangeAction;
  Mvts := TCheckBox(GetControl('CHBXMVTSEUL'));
  if Mvts <> nil then Mvts.OnClick := ChangeContenu;
end;

procedure TOF_PG_GESTIONDOS.ChangeContenu(Sender: TObject);
begin
  SetControlChecked('CHBXETAB', FALSE);
  SetControlChecked('CHBXPARAMSOC', FALSE);
  SetControlChecked('CHBXTIERS', FALSE);
  SetControlChecked('CHBXBQCP', FALSE);
  SetControlChecked('CHBXANNU', FALSE);
  SetControlChecked('CHBXJOURFERIES', FALSE);
  SetControlChecked('CHBXMVT', TRUE);
end;

procedure TOF_PG_GESTIONDOS.ParamGestDOS(Sender: TObject);
{$IFDEF EAGLSERVEUR}
var TobParam, T1: TOB;
  LeSQL: string;
  DD, DF: TDateTime;
{$ENDIF}
begin
{$IFDEF EAGLSERVEUR}
  LeSQL := QMUL.SQL[0];
  DD := StrToDate(LblDD.Caption);
  DF := StrToDate(LblDF.Caption);
  TobParam := TOB.create('Ma Tob de Param', nil, -1);
  T1 := TOB.Create('XXX', TobParam, -1);
  T1.AddChampSupValeur('DD', DD);
  T1.AddChampSupValeur('DF', DF);
  T1.AddChampSupValeur('LESQL', LeSQL);
//  LancePRocess ....
  AGLFicheJob(0, taCreat, 'cgiPaieS5', 'PREPAUTO', TobParam);
  FreeAndNil(TOBParam);
{$ENDIF}
end;
{$ENDIF}

initialization
{$IFNDEF EAGLSERVER}
  registerclasses([TOF_PG_GESTIONDOS]);
{$ENDIF}
end.

