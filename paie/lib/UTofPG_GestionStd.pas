{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 13/11/2002
Modifié le ... :   /  /
Description .. : Gestion : Import/export des standards
Mots clefs ... : PAIE;STANDARD
*****************************************************************}
{
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
// **** Refonte accès V_PGI_env ***** V_PGI_env.ModeFonc remplacé par PgRendModeFonc () *****
PT1 : 08/08/2003 : V_421 FQ 10727 PH Prise en compte des types d'organismes
PT2 : 05/09/2003 : V_421 Traitement du numéro en fonction du type de données et non
                    uniquement en focntion de la base dans laquelle on est connecté
PT3 : 07/09/2004 : V500 PH FQ 11424 Suppression des elements standards avant import
PT4 : 04/08/2005 : V_60 PH FQ 12492 Message en cas de suppression des éléments standards
PT5 : 04/08/2005 : V_70 PH FQ 13587 Prise en compte uniquement du choix de l'export
PT6 : 15/11/2006 : V_70 PH Paie 50, on exporte pas les rubriques de paie STD
PT7 : 06/10/2008 : Bug en import "argument incorrect"
}
unit UTofPG_GESTIONSTD;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
  Fe_Main,
{$ELSE}
  MainEagl,
{$ENDIF}
  HDebug, HCtrls, HEnt1, HMsgBox, UTOF, Ed_Tools, Entpaie, UTOB;
type
  TOF_PG_GESTIONSTD = class(TOF)
  private
    Trace: TListBox; // PT3
    procedure LANCECOPIE(Sender: TObject);
    procedure RetraiteNoDossier(T_Mere: TOB);
    procedure ChangeAction(Sender: TObject);
    procedure ChangeCHBXDOS(Sender: TObject);
    procedure ChangeCHBXSTD(Sender: TObject);
    function RendSQLTablePaie(): string;
    procedure TraiteTableSTD; // PT3
  public
    procedure OnArgument(Arguments: string); override;
  end;

implementation

uses P5def, PgOutils2,
  PGVisuObjet;

procedure TOF_PG_GESTIONSTD.ChangeAction(Sender: TObject);
var
  Coche: TRadioButton;
begin
  Coche := TRadioButton(GetControl('RDTEXPORT'));
  if Coche = nil then exit;
  if Coche.Checked then
  begin
    SetControlEnabled('IMPORTFIC', FALSE);
    SetControlEnabled('EXPORTFIC', TRUE);
    SetControlEnabled('CHBXDOSSIER', TRUE);
    SetControlEnabled('CHBXSTD', TRUE);
  end
  else
  begin
    SetControlEnabled('IMPORTFIC', TRUE);
    SetControlEnabled('EXPORTFIC', FALSE);
    SetControlEnabled('CHBXDOSSIER', FALSE);
    SetControlEnabled('CHBXSTD', FALSE);
  end;
end;

procedure TOF_PG_GESTIONSTD.ChangeCHBXDOS(Sender: TObject);
begin
  // **** Refonte accès V_PGI_env ***** V_PGI_env.ModeFonc remplacé par PgRendModeFonc () *****
  if (PgRendModeFonc() = 'MULTI') then
  begin
    if (GetCheckBoxState('CHBXDOSSIER') = cbChecked) then
      SetControlChecked('CHBXSTD', FALSE);
  end;
end;

procedure TOF_PG_GESTIONSTD.ChangeCHBXSTD(Sender: TObject);
begin
  // **** Refonte accès V_PGI_env ***** V_PGI_env.ModeFonc remplacé par PgRendModeFonc () *****
  if (PgRendModeFonc() = 'MULTI') then
  begin
    if (GetCheckBoxState('CHBXSTD') = cbChecked) then
      SetControlChecked('CHBXDOSSIER', FALSE);
  end;
end;

procedure TOF_PG_GESTIONSTD.LANCECOPIE(Sender: TObject);
var
  T_Mere, T, T1, T2: TOB;
  i, rep, NumVers: Integer;
  LaVersion: Integer;
  T_Filles: array[1..100] of TOB; // soit 100 tables DP+STD maxi dans la paie
  Q: TQuery;
  St, LeNom, Libel, st2: string;
  OkOk, ExportStd: Boolean;
  RDT: TRadioButton;
  FileN, St1: string;
  Pan: TPageControl;
  Tbsht: TTabSheet;
  DOS, STD: Boolean;
begin
  // **** Refonte accès V_PGI_env ***** V_PGI_env.ModeFonc remplacé par PgRendModeFonc () *****
  if (PgRendModeFonc() <> 'MONO') and (PgRendModeFonc() <> 'MULTI') then
  begin
    PGIBox('Traitement impossible car votre environnement n''est pas conforme !', Ecran.Caption);
    exit;
  end;
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
    DOS := (GetCheckBoxState('CHBXDOSSIER') = cbChecked);
    STD := (GetCheckBoxState('CHBXSTD') = cbChecked);
    if (not DOS) and (not STD) then
    begin
      PgiBox('Vous n''avez pas sélectionner d''éléments à exporter !', Ecran.caption);
      exit;
    end;
    FileN := GetControlText('EXPORTFIC');
    if FileN = '' then
    begin
      PgiBox('Vous n''avez pas renseigné le nom du fichier à exporter !', Ecran.caption);
      exit;
    end;
    T_Mere := TOB.Create('Ma TOB', nil, -1);
    T := TOB.Create('La TOB des tables', nil, -1);
    //PT-1 : 08/08/2003 : V_421 FQ 10727 PH Prise en compte des types d'organismes
    st := RendSQLTablePaie;
    Q := OPENSQL(ST, TRUE);
    T.LoadDetailDb('DETABLES', '', '', Q, FALSE);
    Ferme(Q);
    Trace.Items.Add('Les éléments suivants seront analysés et pris en compte s''ils comportent des informations à exporter :');
    for i := 1 to T.detail.count - 1 do
    begin
      T1 := T.Detail[i];
      if T1 <> nil then
      begin
        Trace.Items.Add('Table ' + T1.GetValue('DT_LIBELLE'));
      end;
    end;

    Pan := TPageControl(GetControl('PGCTRL'));
    Tbsht := TTabSheet(GetControl('TBSHTCONTROL'));
    if (Pan <> nil) and (Tbsht <> nil) then Pan.ActivePage := Tbsht;
    Rep := PgiAsk('Voulez vous continuer le traitement', Ecran.Caption);
    if Rep <> MrYes then
    begin
      T.free;
      T_Mere.free;
      exit;
    end;
    Tbsht := TTabSheet(GetControl('TPSHTCAR'));
    if (Pan <> nil) and (Tbsht <> nil) then Pan.ActivePage := Tbsht;
    InitMoveProgressForm(nil, 'Lecture des données', 'Veuillez patienter SVP ...', T.detail.count, FALSE, TRUE);
    for i := 1 to T.detail.count - 1 do
    begin
      T1 := T.Detail[i];
      if T1 <> nil then
      begin
        // **** Refonte accès V_PGI_env ***** V_PGI_env.ModeFonc remplacé par PgRendModeFonc () *****
        if PgRendModeFonc() = 'MONO' then // en MONO, on exporte les éléments STD et du dossier 000000
        begin
          st := 'SELECT * FROM ' + T1.GetValue('DT_NOMTABLE') + ' WHERE ';
          // DEB PT5
          St1 := '';
          St2 := '';
          if GetCheckBoxState('CHBXSTD') = cbChecked then
            st1 := T1.getValue('DT_PREFIXE') + '_PREDEFINI ="STD" ';
          if GetCheckBoxState('CHBXDOSSIER') = cbChecked then
            st2 := T1.getValue('DT_PREFIXE') + '_PREDEFINI ="DOS" AND ' + T1.getValue('DT_PREFIXE') + '_NODOSSIER ="000000"';
          if st1 <> '' then
          begin
          st := st + st1;
          if St2 <> '' then st := st + ' OR ( '+ st2+ ' )';
          end
          else st := st + St2;
         // FIN PT5
        end
        else
        begin // Multi, on exporte uniquement les élts du dossier
          st := 'SELECT * FROM ' + T1.GetValue('DT_NOMTABLE') + ' WHERE ';
          if DOS then
          begin
            if STD then
              st := st + ' (';
            st := st + T1.getValue('DT_PREFIXE') + '_PREDEFINI ="DOS" ';
            // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
            st := st + ' AND ' + T1.getValue('DT_PREFIXE') + '_NODOSSIER ="' + PgRendNoDossier() + '"';
            if STD then
              st := st + ') ';
          end;
          if STD then
          begin
            st2 := T1.getValue('DT_PREFIXE') + '_PREDEFINI ="STD"';
            if DOS then
              st := St + ' OR ' + St2
            else
              st := St + ' ' + st2;
          end;
        end;
        //PT-1 : 08/08/2003 : V_421 FQ 10727 PH Prise en compte des types d'organismes
        if T1.GetValue('DT_NOMTABLE') = 'CHOIXCOD' then
          ST := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE = "PTG"';
        // FIN PT-1
        MoveCurProgressForm('Traitement de la table ' + T1.GetValue('DT_NOMTABLE'));
        if ExisteSQL(st) then
        begin
          Q := OPENSQL(ST, TRUE);
          LeNom := T1.GetValue('DT_LIBELLE'); // +'-V='+IntToStr(T1.GetValue ('DT_NUMVERSION'));
          T_Filles[i + 1] := TOB.Create('T_'+LeNom, T_Mere, -1); //PT7
          T_Filles[i + 1].AddChampSup('VERSION', FALSE);
          T_Filles[i + 1].PutValue('VERSION', T1.GetValue('DT_NUMVERSION'));
          T_Filles[i + 1].LoadDetailDb(T1.GetValue('DT_NOMTABLE'), '', '', Q, FALSE);
          Ferme(Q);
        end;
      end;
    end;

    MoveCurProgressForm('Ecriture des données');
    if FileExists(FileN) then DeleteFile(PChar(FileN));
    try
      T_Mere.SaveToBinFile(FileN, FALSE, TRUE, TRUE, TRUE);
      //T_Mere.SaveToFile (FileN, FALSE, TRUE, TRUE);

      Trace.Items.add('Ecriture du fichier OK');
      CreeJnalEvt('003', '200', 'OK', Trace, nil, nil);
    except
      PGIBox('Une erreur est survenue lors de l''écriture du fichier', Ecran.caption);
      Trace.Items.add('Une erreur est survenue lors de l''écriture du fichier');
      CreeJnalEvt('003', '200', 'ERR', Trace, nil, nil);
    end;
    FiniMoveProgressForm();
    T.free;
    T_Mere.free;
    PgiBox('Exportation terminée', Ecran.caption);
  end
  else
  begin
    FileN := GetControlText('IMPORTFIC');
    if FileN = '' then
    begin
      PgiBox('Vous n''avez pas renseigner de fichier à importer !', Ecran.caption);
      exit;
    end;
    rep := PgiAsk('Attention,le transfert du plan de paie est un traitement qui doit être effectué lors de la mise en place' + ' du dossier.#13#10Les éléments importés remplaceront les éléments déjà en place dans ce dossier (rubriques, variables, cumuls, profils, etc.)#13#10 Voulez vous continuer le traitement ?', Ecran.caption);
    if Rep <> MrYes then exit;
    if GetControlText ('CHBSUPSTD') = 'X' then
    begin // DEB PT4
      St := 'Attention ! Cette commande supprimera l''ensemble des données prédéfinies STANDARD de TOUS les dossiers (rémunérations, cotisations, profils, éléments nationaux, .... Confirmez-vous l''abandon du traitement ?';
      rep := PgiAsk(St, Ecran.caption);
      if Rep <> MrNo then exit;
      if (AGLLanceFiche('PAY','COHERENCEREP','','','')<>'X') Then exit;
      // FIN PT4
    end;
    T := TOB.Create('La TOB des tables', nil, -1);
    ST := RendSQLTablePaie;
    Q := OPENSQL(ST, TRUE);
    T.LoadDetailDb('DETABLES', '', '', Q, FALSE);
    Ferme(Q);
    Pan := TPageControl(GetControl('PGCTRL'));
    Tbsht := TTabSheet(GetControl('TBSHTCONTROL'));
    if (Pan <> nil) and (Tbsht <> nil) then Pan.ActivePage := Tbsht;
    if FileExists(FileN) then
    begin
      InitMoveProgressForm(nil, 'Lecture des données', 'Veuillez patienter SVP ...', 2, FALSE, TRUE);
      T_Mere := TOB.Create('Ma TOB', nil, -1);
      TOBLoadFromBinFile(FileN, nil, T_Mere);
      Trace.Items.Add('Les éléments suivants seront pris en compte : ');
      MoveCurProgressForm('Traitement des données, veuillez patienter SVP');
      rep := MrYes;
      if T_Mere.detail.count - 1 < 1 then
      begin
        FiniMoveProgressForm();
        T_Mere.free;
        T.Free;
        PgiBox('Vous n''avez aucun élément à importer', Ecran.Caption);
        exit;
      end;
      for i := 1 to T_Mere.detail.count - 1 do
      begin
        LeNom := T_Mere.Detail[i].NomTable;
        Libel := Copy(LeNom, 1, StrLen(PChar(LeNom)));
        LaVersion := T_Mere.Detail[i].GetValue('VERSION');
        T2 := T.FindFirst(['DT_LIBELLE'], [Libel], FALSE);
        if T2 <> nil then NumVers := T2.GetValue('DT_NUMVERSION')
        else NumVers := 0;
        Trace.Items.Add('Table ' + LeNom);
        if (LaVersion <> NumVers) and (NumVers <> 0) then
        begin
          Trace.Items.Add('Version de table incompatible export=' + IntToStr(LaVersion) + ' import=' + IntToStr(NumVers));
          rep := MrCancel;
        end;
      end;
      if Rep = MrYes then Rep := PgiAsk('Voulez vous continuer le traitement ?', Ecran.Caption)
      else Trace.Items.Add('Traitement abandonné car toutes les tables ne sont pas compatibles');
      if Rep <> MrYes then
      begin
        FiniMoveProgressForm();
        T_Mere.free;
        T.Free;
        exit;
      end;
      OkOk := FALSE;
      try
        // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
        // **** Refonte accès V_PGI_env ***** V_PGI_env.ModeFonc remplacé par PgRendModeFonc () *****
        Debug('Mode : ' + PgRendModeFonc() + ' Dossier = ' + PgRendNoDossier());
        //        if Uppercase(PgRendModeFonc()) = 'MULTI' then
        // DEB PT3
        BeginTrans;
        RetraiteNoDossier(T_Mere);
        STD := (GetCheckBoxState('CHBSUPSTD') = cbChecked);
        if STD then TraiteTableSTD; // Suppression des elements standards
        OkOk := T_Mere.InsertOrUpdateDb(TRUE);
        CommitTrans; // FIN PT3
        Trace.Items.add('Traitement du fichier OK');
        CreeJnalEvt('003', '201', 'OK', Trace, nil, nil);
      except
        RollBack; // PT3
        PGIBox('Une erreur est survenue lors du traitement du fichier', Ecran.caption);
        Trace.Items.add('Une erreur est survenue lors du traitement du fichier');
        CreeJnalEvt('003', '201', 'ERR', Trace, nil, nil);
      end;
      FiniMoveProgressForm();
      T.free;
      T_Mere.free;
      if OkOk then PgiBox('Le traitement est terminé', Ecran.caption)
      else PgiBox('Une erreur est survenue pendant l''insertion des données', Ecran.caption);
    end
    else PgiBox('Le fichier ' + FileN + ' n''existe pas', Ecran.caption);
  end;

end;

procedure TOF_PG_GESTIONSTD.OnArgument(Arguments: string);
var
  BtnVal: TToolbarButton97;
  st: string;
  Coche: TRadioButton;
begin
  inherited;
  BtnVal := TToolbarButton97(GetControl('BValider'));
  if BtnVal <> nil then BtnVal.OnClick := LANCECOPIE;
  {
  if V_PGI_env.ModeFonc ='MULTI' then
     begin
     SetControlProperty ('RDEXPORT', 'Caption', 'Exportation des données de votre dossier');
     SetControlVisible ('CHBXDOSSIER',FALSE);
     end;}
  // **** Refonte accès V_PGI_env ***** V_PGI_env.ModeFonc remplacé par PgRendModeFonc () *****
  if (PgRendModeFonc() = 'MULTI') then
  begin
    st := 'Attention, la recopie de plan de paie dossier à dossier ne doit #13#10' +
      'pas se substituer à l''utilisation du plan de paie standard.#13#10' +
      'Avant toute intervention, vous devez analyser vos plans de paie #13#10' +
      'et utiliser de préférence des rubriques, des cumuls, des variables ...#13#10' +
      'de type standard (STD).';
    PgiBox(St, Ecran.caption);
  end;
  Coche := TRadioButton(GetControl('RDTEXPORT'));
  if Coche <> nil then Coche.OnClick := ChangeAction;
  Coche := TRadioButton(GetControl('RDTIMPORT'));
  if Coche <> nil then Coche.OnClick := ChangeAction;
  Coche := TRadioButton(GetControl('CHBXDOSSIER'));
  if Coche <> nil then Coche.OnClick := ChangeCHBXDOS;
  Coche := TRadioButton(GetControl('CHBXSTD'));
  if Coche <> nil then Coche.OnClick := ChangeCHBXSTD;
{$IFDEF CPS1}
  SetControlVisible ('CHBXSTD', FALSE);
  SetControlEnabled ('CHBXDOSSIER', FALSE);
{$ENDIF}
end;
//PT-1 : 08/08/2003 : V_421 FQ 10727 PH Prise en compte des types d'organismes

function TOF_PG_GESTIONSTD.RendSQLTablePaie: string;
begin
  result :=
    'SELECT DT_NOMTABLE,DT_PREFIXE,DT_LIBELLE,DT_NUMVERSION FROM DETABLES WHERE (DT_DOMAINE = "P" AND EXISTS (SELECT DH_NOMCHAMP FROM DECHAMPS WHERE DH_NOMCHAMP=DT_PREFIXE||"_PREDEFINI"))';
  if PgRendModeFonc() = 'MONO' then result := result + ' OR (DT_NOMTABLE = "CHOIXCOD")';
  result := result + ' ORDER BY DT_NOMTABLE';
end;
// DEB PT3

procedure TOF_PG_GESTIONSTD.TraiteTableSTD;
var
  st: string;
  T, T1: TOB;
  i: integer;
  Q: TQuery;
begin
  try
    St :=
      'SELECT DT_NOMTABLE,DT_LIBELLE,DH_NOMCHAMP FROM DETABLES LEFT JOIN DECHAMPS ON ' +
      'DH_NOMCHAMP=DT_PREFIXE||"_PREDEFINI" WHERE (DT_DOMAINE = "P" AND EXISTS ' +
      '(SELECT DH_NOMCHAMP FROM DECHAMPS WHERE DH_NOMCHAMP=DT_PREFIXE||"_PREDEFINI"))';
    st := st + ' ORDER BY DT_NOMTABLE';
    Q := OPENSQL(ST, TRUE);
    T := TOB.Create ('MES_DETABLES',NIL,-1);
    T.LoadDetailDb('LES_DETABLES', '', '', Q, FALSE);
    Ferme(Q);
    begintrans;
    Trace.Items.Add('Les éléments standards des tables suivantes sont supprimés :');
    for i := 1 to T.detail.count - 1 do
    begin
      T1 := T.Detail[i];
      if T1 <> nil then
      begin
        Trace.Items.Add('Table ' + T1.GetValue('DT_LIBELLE'));
        ExecuteSQl('DELETE FROM ' + T1.GetValue('DT_NOMTABLE') + ' WHERE ' + T1.GetValue('DH_NOMCHAMP') + '="STD"');
      end;
    end;
    CommitTrans;
  except
    Rollback;
    PGIBox('Une erreur est survenue lors de la suppression des éléments standards', 'Suppression des éléments standard');
  end;
  FreeAndNil(T);
end;
// FIN PT3
// Fonction qui parcourt les tob pour réaffecter le bon numero de dossier
// en restauration - Uniquement en environnement MULTI

procedure TOF_PG_GESTIONSTD.RetraiteNoDossier(T_Mere: TOB);
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
        // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
        Debug('Nom table ' + T2.NomTable + ' préfixe ' + Prefix + ' ' + PgRendNoDossier());
        // PT-2 : 05/09/2003 : V_421 Traitement du numéro en fonction du type de données
        if (T2.GetValue(Prefix + '_PREDEFINI') = 'DOS') then T2.PutValue(Prefix + '_NODOSSIER', PgRendNoDossier())
        else T2.PutValue(Prefix + '_NODOSSIER', '000000');
      end;
    end;
  end;

end;

initialization
  registerclasses([TOF_PG_GESTIONSTD]);
end.

