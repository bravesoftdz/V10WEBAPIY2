{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 03/01/2002
Modifié le ... :   /  /
Description .. : Unit de confection du fichier des déclarations prud'homales
Suite ........ :
Mots clefs ... : PAIE;PRUDH
*****************************************************************}
{
PT1-0 : 28/03/2002 VG V571 Modification de la fonction de création
                           d'enregistrement dans la table ENVOISOCIAL
PT1   : 08/04/2002 PH V571 Right join en oracle remplacé par left join et tests
                           complémentaires
PT2   : 22/04/2002 PH V571 Prise en compte cas > 100 etablissements alors
                           suffixe en alpha numérique
PT3   : 22/04/2002 PH V571 Exclusion des stagiaires rémunérés
PT4   : 22/04/2002 SB V571 Ajout clause requête état : liste prud'homale
PT5   : 17/04/2003 VG V_42 Critère de sélection d'une clause WHERE sans libellé
PT6   : 19/05/2006 SB V_65 Appel des fonctions communes pour identifier V_PGI.driver
PT7   : 06/02/2007 FC V_80 Mise en place filtrage des habilitations/populations
}
unit UTofPG_PrepPrudh;

interface
uses Windows, StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
  Dialogs, ed_tools, Mailol,
  {$IFDEF EAGLCLIENT}
  UtileAGL, eFiche,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fiche, Qre, DbGrids, HDB,
  {$IFDEF V530}
  EdtEtat,
  {$ELSE}
  EdtREtat,
  {$ENDIF}
  {$ENDIF}

  Grids, HCtrls, HEnt1, EntPaie, HMsgBox, UTOF, UTOB, UTOM, Vierge, HQry, HStatus,
  AGLInit, PgOutils, Pgoutils2,ShellAPI,
  P5Def //PT7
  ;
type
  TOF_PG_PrepPrudh = class(TOF)
  private
    LesEtab: TOB; // TOB des etablissements
    LaDate: THEdit; // La date à laquelle sont présents les salariés pour la déclaration
    Lesiret, LesiretDos, Lesuffixe: string;
    FEcrt: TextFile;
    LaGrille: THGRID; // La grille de saisie des codes insee
    GrilleMod: Boolean;
    procedure LancePrepPrudh(Sender: TObject);
    procedure TheGrilleExit(Sender: TObject);
    function EcritureEtab(Te: TOB; TraceErr: TListBox): boolean;
    function EcritTotEtab(Nbre: Integer; TraceErr: TListBox): boolean;
    function EcritureSalarie(Ts: TOB; TraceErr, Trace: TListBox): Integer;
    function ControleCodeInsee(LeCode, Action: string): boolean;
    procedure TheCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure MajEnvoiSocial(NbreEnrg: Integer; NomFic: string);
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
  end;

implementation

uses
  HSysMenu,
  ParamSoc,
  PG_OutilsEnvoi,
  Pgcommun;

type
  E020 = record // ENREGISTREMENT DE TYPE ETABLISSEMENT
    siret: string[14];
    suffixe: string[2];
    L1: string[3];
    TypeEnr: string[3];
    SiretEt: string[14];
    Ape: string[4];
    L2: string[56];
    RaisonSoc: string[50];
    CpltAdresse: string[32];
    L3: string[1];
    NumVoie: string[4];
    Bis: string[1];
    L4: string[1];
    Voie: string[26];
    CodeInsee: string[5];
    L5: string[1];
    Commune: string[26];
    CodePostal: string[5];
    L6: string[1];
    BurDist: string[26];
    L7: string[86];
    TypeEmpl: string[2];
    L8: string[201];
  end;
type
  E200 = record // ENREGISTREMENT DE TYPE SALARIE
    siret: string[14];
    suffixe: string[2];
    L1: string[3];
    TypeEnr: string[3];
    L2: string[10];
    NumSS: string[13];
    CleSS: string[2];
    DateNais: string[6];
    DeptNais: string[2];
    L3: string[3];
    LieuNais: string[26];
    L4: string[3];
    Nomnais: string[30];
    Prenom: string[20];
    NomEpouse: string[30];
    CpltAdresse: string[32];
    L5: string[1];
    NumVoie: string[4];
    Bis: string[1];
    L6: string[1];
    Voie: string[26];
    CodeInsee: string[5];
    L7: string[1];
    Commune: string[26];
    CodePostal: string[5];
    L8: string[1];
    BurDist: string[26];
    L9: string[232];
    College: string[1];
    Section: string[1];
    Lieu: string[1];
    L10: string[33];
  end;

type
  E300 = record // ENREGISTREMENT DE TYPE TOTAL ETABLISSEMENT
    siret: string[14];
    suffixe: string[2];
    L1: string[3];
    TypeEnr: string[3];
    L2: string[200];
    L3: string[112];
    NbrePers: string[6]; // nombre de salaries presents dans l'etablissement
    L4: string[224];
  end;

function TOF_PG_PrepPrudh.EcritureEtab(Te: TOB; TraceErr: TListBox): boolean;
var
  S: string;
  E: E020; // enregistrement de type etablissement
begin
  result := TRUE;
  Lesiret := Te.getvalue('ET_SIRET');
  Lesuffixe := Te.getvalue('Suffixe');

  with E do
  begin
    siret := Format_String(Lesiret, 14);
    suffixe := Format_String(Lesuffixe, 2);
    L1 := StringOfChar(' ', sizeof(L1));
    TypeEnr := '020';
    siretEt := siret;
    Ape := Format_String(te.getvalue('ET_APE'), 4);
    L2 := StringOfChar(' ', sizeof(L2));
    RaisonSoc := Format_String(te.getvalue('ET_LIBELLE'), 50);
    CpltAdresse := Format_String(te.getvalue('ET_ADRESSE2'), 32);
    L3 := StringOfChar(' ', sizeof(L3));
    NumVoie := StringOfChar(' ', sizeof(NumVoie));
    Bis := StringOfChar(' ', sizeof(Bis));
    L4 := StringOfChar(' ', sizeof(L4));
    Voie := Format_String(te.getvalue('ET_ADRESSE1'), 26);
    CodeInsee := Format_String(te.getvalue('LeCodeInsee'), 5);
    if not ControleCodeInsee(CodeInsee, 'E') then
      TraceErr.Items.Add('Le code insee n''est pas numérique ou ne comporte pas 5 chiffres');
    L5 := StringOfChar(' ', sizeof(L5));
    Commune := StringOfChar(' ', sizeof(Commune));
    CodePostal := Format_String(te.getvalue('ET_CODEPOSTAL'), 5);
    L6 := StringOfChar(' ', sizeof(L6));
    BurDist := Format_String(te.getvalue('ET_VILLE'), 26);
    L7 := StringOfChar(' ', sizeof(L7));
    TypeEmpl := Copy(te.getvalue('ETB_PRUDH'), 1, 2);
    if TypeEmpl = '00' then TypeEmpl := '  ';
    L8 := StringOfChar(' ', sizeof(L8));
  end;
  S := E.siret + E.suffixe + E.L1 + E.TypeEnr + E.SiretEt + E.Ape + E.L2 + E.RaisonSoc + E.CpltAdresse +
    E.L3 + E.NumVoie + E.Bis + E.L4 + E.Voie + E.CodeInsee + E.L5 + E.Commune + E.CodePostal + E.L6 +
    E.BurDist + E.L7 + E.TypeEmpl + E.L8;
  {$I-}Writeln(FEcrt, S);
  {$I+}if IoResult <> 0 then
  begin
    PGIBox('Erreur d''écriture du fichier prud''hom ligne établissement', 'Abandon du traitement');
    TraceErr.Items.Add('Erreur d''écriture du fichier prud''hom ligne établissement ' + Copy(S, 1, 40));
    Result := FALSE;
  end;
end;

function TOF_PG_PrepPrudh.EcritureSalarie(Ts: TOB; TraceErr, Trace: TListBox): Integer;
var
  S: string;
  E: E200; // enregistrement de type salarie
  JJ, MM, AA: WORD;
  UneDate: TDateTime;
begin
  result := 0;

  with E do
  begin
    siret := Format_String(Lesiret, 14);
    suffixe := Format_String(Lesuffixe, 2);
    L1 := StringOfChar(' ', sizeof(L1));
    TypeEnr := '200';
    L2 := StringOfChar(' ', sizeof(L2));
    NumSS := Format_String(Copy(ts.getvalue('PSA_NUMEROSS'), 1, 13), 13);
    if Copy(NumSS, 1, 2) = '  ' then
    begin
      Trace.Items.Add('Pas de numéro de sécurité sociale pour le salarié : ' + ts.getvalue('PSA_SALARIE'));
      result := 1;
    end;
    CleSS := Format_String(Copy(ts.getvalue('PSA_NUMEROSS'), 14, 2), 2);
    if CleSS = '  ' then
    begin
      Trace.Items.Add('Pas de clé sécurité sociale pour le salarié : ' + ts.getvalue('PSA_LIBELLE') + ' ' + ts.getvalue('PSA_PRENOM'));
      result := 1;
    end;

    UneDate := ts.getvalue('PSA_DATENAISSANCE');
    if (UneDate = 0) then
    begin
      Trace.Items.Add('Pas de date de naissance pour le salarié : ' + ts.getvalue('PSA_LIBELLE') + ' ' + ts.getvalue('PSA_PRENOM'));
      result := 1;
    end;

    DecodeDate(UneDate, AA, MM, JJ);
    DateNais := PGZeroAGauche(IntToStr(JJ), 2) + PGZeroAGauche(IntToStr(MM), 2) +
      PGZeroAGauche(Copy(IntToStr(AA), 3, 2), 2);
    DeptNais := Format_String(ts.getvalue('PSA_DEPTNAISSANCE'), 2);
    if (DeptNais = '  ') and (NumSS <> '') then DeptNais := Format_String(Copy(NumSS, 6, 2), 2);
    if DeptNais = '  ' then
    begin
      Trace.Items.Add('Pas de département de naissance pour le salarié : ' + ts.getvalue('PSA_LIBELLE') + ' ' + ts.getvalue('PSA_PRENOM'));
      result := 1;
    end;
    L3 := StringOfChar(' ', sizeof(L3));
    LieuNais := Format_String(ts.getvalue('PSA_COMMUNENAISS'), 26);
    if ts.getvalue('PSA_COMMUNENAISS') = '' then
    begin
      Trace.Items.Add('Pas de lieu de naissance pour le salarié : ' + ts.getvalue('PSA_LIBELLE') + ' ' + ts.getvalue('PSA_PRENOM'));
      result := 1;
    end;
    L4 := StringOfChar(' ', sizeof(L4));
    NomNais := Format_String(ts.getvalue('PSA_LIBELLE'), 30);
    NomEpouse := StringOfChar(' ', sizeof(NomEpouse));
    Prenom := Format_String(ts.getvalue('PSA_PRENOM'), 20);
    if Ts.getvalue('PSA_SEXE') = 'F' then
    begin
      if Ts.getvalue('PSA_NOMJF') <> '' then
      begin
        NomNais := Format_String(ts.getvalue('PSA_NOMJF'), 30);
        NomEpouse := Format_String(ts.getvalue('PSA_LIBELLE'), 30);
      end;
    end;
    CpltAdresse := Format_String(ts.getvalue('PSA_ADRESSE2'), 32);
    L5 := StringOfChar(' ', sizeof(L5));
    NumVoie := StringOfChar(' ', sizeof(NumVoie));
    Bis := StringOfChar(' ', sizeof(Bis));
    L6 := StringOfChar(' ', sizeof(L6));
    Voie := Format_String(ts.getvalue('PSA_ADRESSE1'), 26);
    CodeInsee := StringOfChar(' ', sizeof(CodeInsee));
    L7 := StringOfChar(' ', sizeof(L7));
    Commune := StringOfChar(' ', sizeof(Commune));
    CodePostal := Format_String(ts.getvalue('PSA_CODEPOSTAL'), 5);
    L8 := StringOfChar(' ', sizeof(L8));
    BurDist := Format_String(ts.getvalue('PSA_VILLE'), 26);
    L9 := StringOfChar(' ', sizeof(L9));
    College := Format_String(ts.getvalue('PSA_PRUDHCOLL'), 1);
    if College = ' ' then
    begin
      TraceErr.Items.Add('Pas de collège éléctoral pour le salarié : ' + ts.getvalue('PSA_LIBELLE') + ' ' + ts.getvalue('PSA_PRENOM'));
      result := 1;
    end;
    Section := Format_String(ts.getvalue('PSA_PRUDHSECT'), 1);
    if Section = ' ' then
    begin
      TraceErr.Items.Add('Pas de section éléctorale pour le salarié : ' + ts.getvalue('PSA_LIBELLE') + ' ' + ts.getvalue('PSA_PRENOM'));
      result := 1;
    end;
    Lieu := Format_String(ts.getvalue('PSA_PRUDHVOTE'), 1);
    if Lieu = ' ' then
    begin
      TraceErr.Items.Add('Pas de lieu de vote pour le salarié : ' + ts.getvalue('PSA_LIBELLE') + ' ' + ts.getvalue('PSA_PRENOM'));
      result := 1;
    end;
    L10 := StringOfChar(' ', sizeof(L10));
  end;
  S := E.siret + E.suffixe + E.L1 + E.TypeEnr + E.L2 + E.NumSS + E.CleSS +
    E.DateNais + E.DeptNais + E.L3 + E.LieuNais + E.L4 + E.NomNais + E.Prenom + E.NomEpouse +
    E.CpltAdresse + E.L5 + E.NumVoie + E.Bis + E.L6 + E.Voie +
    E.CodeInsee + E.L7 + E.Commune + E.CodePostal + E.L8 + E.BurDist + E.L9 +
    E.College + E.Section + E.Lieu + E.L10;
  {$I-}Writeln(FEcrt, S);
  {$I+}if IoResult <> 0 then
  begin
    PGIBox('Erreur d''écriture du fichier prud''hom ligne salarié', 'Abandon du traitement');
    TraceErr.Items.Add('Erreur d''écriture du fichier prud''hom ligne salarié ' + Copy(S, 1, 40));
    Result := 2;
  end;
end;


procedure TOF_PG_PrepPrudh.LancePrepPrudh(Sender: TObject);
var
  nomfic, st: string;
  Pan: TPageControl;
  Tbsht: TTabSheet;
  Trace, TraceErr: TListBox;
  Q: TQuery;
  i, Reponse, Nbre, NbreErr, NbreEnrg, NbreA: Integer;
  F: TextFile;
  okok, Anomalies: Boolean;
  TSal, Te, Ts: TOB;
  Okok1: Integer;
  FileN: string; // Nom des fichiers
  Habilitation:String;//PT7
begin
  Anomalies := FALSE;
  if LaGrille = nil then
  begin
    PgiBox('Attention, erreur grille de saisie établissement', 'Traitement interrompu');
    TraceErr.Items.Add('Erreur récupération code insee établissement');
    exit;
  end;
  NbreEnrg := 0;
  Pan := TPageControl(GetControl('PANELPREP'));
  Tbsht := TTabSheet(GetControl('TBSHTTRACE'));
  Trace := TListBox(GetControl('LSTBXTRACE'));
  TraceErr := TListBox(GetControl('LSTBXERROR'));
  if (Trace = nil) or (TraceErr = nil) then
  begin
    PGIBox('La préparation du fichier ne peut pas être lancée', 'Car les composants trace ne sont pas disponibles');
    exit;
  end;
  if (Pan <> nil) and (Tbsht <> nil) then Pan.ActivePage := Tbsht;
  //DEB PT7
  Habilitation := '';
  if Assigned(MonHabilitation) and (MonHabilitation.LeSQL<>'') then
    Habilitation := ' AND ' + MonHabilitation.LeSQL;
  //FIN PT7
  //impression de la liste prud'homale
  if PGIAsk('Voulez-vous éditer une liste Prud''homale?', Ecran.caption) = MrYes then //PT4 Ajout Clause
    LanceEtat('E', 'PDA', 'PLI', True, False, False, Pan, 'AND PSA_DADSPROF<>"09" AND PSA_DADSPROF<>"10" AND PSA_DADSPROF<>"11"' + Habilitation, '', False); //PT7
  // Recupération des modifications des code Insee saisis car ils ne correspondent pas forcément au code postal
  //LesEtab.getGridDetail (LaGrille,-1,'','LeCodeInsee');
  if GrilleMod then
  begin
    if LaGrille.RowCount > 1 then
    begin
      for i := 1 to LaGrille.RowCount - 1 do
      begin
        Te := LesEtab.Detail[i - 1];
        if Te <> nil then Te.putvalue('LeCodeInsee', LaGrille.Cells[4, i]);
      end;
    end
    else
    begin
      Te := LesEtab.Detail[0];
      if Te <> nil then Te.putvalue('LeCodeInsee', LaGrille.Cells[4, 0]);
    end;
  end;
  NomFic := 'PRUD' + LeSiretDos + '.TXT';
  FileN := '$DAT\' + NomFic;
  FileN := ChangeStdDatPath(FileN);
  if FileExists(FileN) then
  begin
    reponse := HShowMessage('5;;Voulez-vous supprimer votre fichier prud''hom ' + FileN + ';Q;YN;Y;N', '', '');
    if reponse = 6 then DeleteFile(PChar(FileN))
    else exit;
  end;
  AssignFile(FEcrt, FileN);
  {$I-}ReWrite(FEcrt);
  {$I+}if IoResult <> 0 then
  begin
    PGIBox('Fichier inaccessible : ' + FileN, 'Abandon du traitement');
    Exit;
  end;
  St := 'SELECT PSA_SALARIE,PSA_NUMEROSS,PSA_DATENAISSANCE,PSA_DEPTNAISSANCE,PSA_COMMUNENAISS,PSA_NOMJF,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,+PSA_ADRESSE1,PSA_SEXE,' +
    'PSA_ADRESSE2,PSA_CODEPOSTAL,PSA_VILLE,PSA_PRUDHCOLL,PSA_PRUDHSECT,PSA_PRUDHVOTE FROM SALARIES WHERE ((PSA_DATESORTIE IS NULL OR PSA_DATESORTIE <= "' + UsDateTime(Idate1900) +
      '") OR PSA_DATESORTIE >"';
  St := St + USDateTime(StrToDate(GetControlText('LaDate'))) + '") AND PSA_DATEENTREE <="' + USDateTime(StrToDate(GetControlText('LaDate'))) + '"';
  // PT3 22/04/02 PT1 PH Exclusion des stagiaires rémunérés
  St := St + ' AND PSA_DADSPROF <> "09" AND PSA_DADSPROF <> "10" AND PSA_DADSPROF <> "11" ' + Habilitation;
  st := st + ' ORDER BY PSA_ETABLISSEMENT,PSA_SALARIE';
  Q := OpenSql(St, TRUE);
  TSal := TOB.create('Les Salariés', nil, -1);
  TSal.loadDetailDb('SALARIES', '', '', Q, FALSE, False);
  Ferme(Q);
  Trace.Items.Add('Début du traitement ');
  Te := LesEtab.FindFirst([''], [''], FALSE);
  while Te <> nil do
  begin
    Nbre := 0;
    NbreErr := 0;
    NbreA := 0;
    okok := EcritureEtab(Te, TraceErr);
    if okok then
    begin
      Trace.Items.Add('------------------------------------------------------');
      Trace.Items.Add('Etablissement en cours  ' + Te.GetValue('ETB_ETABLISSEMENT') + ' ' + Te.GetValue('ET_LIBELLE'));
      for i := 0 to TSal.Detail.Count - 1 do
      begin
        Ts := TSal.Detail[i];
        if Ts.getvalue('PSA_ETABLISSEMENT') = Te.getvalue('ETB_ETABLISSEMENT') then
        begin
          Okok1 := EcritureSalarie(Ts, TraceErr, Trace);
          if okok1 < 2 then Trace.Items.Add('Salarié traité ' + TS.GetValue('PSA_SALARIE') + ' ' + TS.GetValue('PSA_LIBELLE') + ' ' + TS.GetValue('PSA_PRENOM'));
          if okok1 < 2 then Nbre := Nbre + 1 else NbreErr := NbreErr + 1;
          if okok1 = 1 then
          begin
            NbreA := NbreA + 1;
            Anomalies := TRUE;
          end;
        end;
      end; // fin de traitement des salaries
      EcritTotEtab(Nbre, TraceErr);
      NbreEnrg := NbreEnrg + Nbre;
      Trace.Items.Add('===> ' + IntToStr(Nbre) + ' Salariés traités-' + IntToStr(NbreErr) + ' Salariés rejetés');
      if NbreA > 0 then
        Trace.Items.Add('===> ' + IntToStr(NbreA) + ' Salariés à contrôler comportant des zones non renseignées pour cet établissement');
      Trace.Items.Add('Fin du traitement établissement ' + Te.GetValue('ET_LIBELLE'));
    end;
    Te := LesEtab.FindNext([''], [''], FALSE);
  end; // Fin du while boucle établissement
  CloseFile(FEcrt);
  // Mise à jour de la table envoisocial si pas d'erreur
  if TraceErr.Items.Count = 0 then MajEnvoiSocial(NbreEnrg, NomFic);
  Trace.Items.Add('Fin du traitement ');
  if Anomalies then
  begin
    Tbsht := TTabSheet(GetControl('TBSHTTRACE'));
    if Tbsht <> nil then Pan.ActivePage := Tbsht;
    //Génération d'un fichier de log
    if MessageDlg('Voulez-vous générez le fichier ControlPrudh.log sous le répertoire ' + VH_Paie.PGCheminEagl + '\.', mtConfirmation, [mbYes, mbNo], 0) = mrYes then //PT2
    begin
      if VH_Paie.PGCheminEagl <> '' then FileN := VH_Paie.PGCheminEagl + '\ControlPrudh.log' //PT2
      else FileN := 'C:\ControlPrudh.log';
      if SupprimeFichier(FileN) = False then exit;
      AssignFile(F, FileN);
      {$I-}ReWrite(F);
      {$I+}
      if IoResult <> 0 then
      begin
        PGIBox('Fichier inaccessible : ' + FileN, 'Abandon du traitement');
        Exit;
      end;
      writeln(F, 'Préparation du fichier de contrôle');
      for i := 0 to Trace.Items.Count - 1 do
      begin
        St := Trace.Items.Strings[i];
        writeln(F, St);
      end;
      CloseFile(F);
      PGIInfo('La génération est terminée', Ecran.caption);
      //Ouverture du fichier de control
      if Trace.Items.Count > 0 then
      begin
        if PGIAsk('Voulez-vous éditer la liste des traitements et des erreurs constatées?', Ecran.caption) = MrYes then
          ShellExecute(0, PCHAR('open'), PChar('WordPad'), PChar(FileN), nil, SW_RESTORE);
      end;
    end;
  end;
  if TraceErr.Items.Count > 0 then
  begin
    Tbsht := TTabSheet(GetControl('TBSHTERROR'));
    if Tbsht <> nil then Pan.ActivePage := Tbsht;
    //Génération d'un fichier de log
    if MessageDlg('Voulez-vous générez le fichier Prudh.log sous le répertoire ' + VH_Paie.PGCheminEagl + '\.', mtConfirmation, [mbYes, mbNo], 0) = mrYes then //PT2
    begin
      if VH_Paie.PGCheminEagl <> '' then FileN := VH_Paie.PGCheminEagl + '\Prudh.log' //PT2
      else FileN := 'C:\Prudh.log';
      if SupprimeFichier(FileN) = False then exit;
      AssignFile(F, FileN);
      {$I-}ReWrite(F);
      {$I+}
      if IoResult <> 0 then
      begin
        PGIBox('Fichier inaccessible : ' + FileN, 'Abandon du traitement');
        Exit;
      end;
      writeln(F, 'Préparation prud''hom : Gestion des messages d''erreur.');
      for i := 0 to TraceErr.Items.Count - 1 do
      begin
        St := TraceErr.Items.Strings[i];
        writeln(F, St);
      end;
      CloseFile(F);
      PGIInfo('La génération est terminée', Ecran.caption);
      //Ouverture du fichier d'erreur
      if TraceErr.Items.Count > 0 then
        ShellExecute(0, PCHAR('open'), PChar('WordPad'), PChar(FileN), nil, SW_RESTORE);
    end;
  end;
  TSal.free;
  TSal := nil;
end;

procedure TOF_PG_PrepPrudh.OnArgument(Arguments: string);
var
  st, Princ: string;
  BtnLance: TToolbarButton97;
  Q: TQuery;
  T1: TOB;
  HMTrad: THSystemMenu;
  i, Nbre, Nbre1, Nbre2, Err, Indice, NbreS: Integer;
  Requete, Etablis, Habilitation: string;  //PT7
  j: Integer;                         //PT7
const
  Tableau: array[1..36] of string = ('0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D',
    'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z');
begin
  inherited;
  GrilleMod := FALSE;
  Err := 0;
  Indice := 0;
  LaDate := THedit(GetControl('LaDate'));
  if LaDate <> nil then LaDate.Text := '29/03/2002'
  else Exit;
  LaGrille := THGRID(GetControl('GRILLEETAB'));
  if LaGrille = nil then exit;
  LaGrille.OnCellExit := TheCellExit; // Pour controler la saisie du code insee soit  5 c
  LaGrille.OnExit := TheGrilleExit; //
  Princ := GetParamSoc('SO_ETABLISDEFAUT'); // Etablissement principal defini dans la compta
  if Princ = '' then
  begin
    PGIBox('Vous n''avez défini d''établissement principal dans vos paramètres comptables', Ecran.Caption);
    Exit;
  end;
  LeSiretDos := GetParamSoc('SO_SIRET');
  ForceNumerique(LeSiretDos, LeSiretDos);
  if ControlSiret(LeSiretDos) = False then
  begin
    PGIBox('Le SIRET n''est pas valide', 'Gestion des paramètres société');
    exit;
  end;

  BtnLance := TToolbarButton97(GetControl('BVALIDER'));
  SetControlEnabled('BValider', FALSE);
  if BtnLance <> nil then BtnLance.OnClick := LancePrepPrudh;
  Nbre1 := 0;
  Nbre2 := 0;

  //DEB PT7
  Habilitation := '';
  if Assigned(MonHabilitation) and (MonHabilitation.LeSQL<>'') then
  begin // Champs etablissement seul comme critere de la population
    Requete := MonHabilitation.LesEtab;
    Etablis := ReadTokenSt(Requete);
    j := 0;
    while Etablis <> '' do
    begin
      j := j + 1;
      if Etablis <> '' then
      begin
        if j > 1 then Habilitation := Habilitation + ',';
        Habilitation := Habilitation + '"' + Etablis + '"';
      end;
      Etablis := ReadTokenSt(Requete);
    end;
    if j > 0 then
      Habilitation := ' WHERE ET_ETABLISSEMENT IN (' + Habilitation + ')';
  end;
  //FIN PT7

  // PT1 08/04/02 PT1 PH Right join en oracle remplacé par left join et tests complémentaires
  if (PGisOracle) then { PT6 }
  begin
    st := 'select count (*) Nbre1 from etabcompl';
    Q := OpenSql(st, TRUE);
    if not Q.EOF then Nbre1 := Q.FindField('Nbre1').AsInteger;
    Ferme(Q);
    st := 'select count (*) Nbre2 from etabliss';
    Q := OpenSql(st, TRUE);
    if not Q.EOF then Nbre2 := Q.FindField('Nbre2').AsInteger;
    Ferme(Q);
    if Nbre1 <> Nbre2 then PgiBox('Attention, vous n''avez pas le même nombre d''établissements entre la comptabilité et la paie', Ecran.Caption);
    st := 'SELECT ETB_ETABLISSEMENT,ET_LIBELLE,ETB_PRUDH,CO1.CO_LIBRE as TypeEmpl,ET_SIRET,ET_APE,ET_ADRESSE1,ET_ADRESSE2,ET_CODEPOSTAL,ET_VILLE FROM ETABCOMPL LEFT OUTER JOIN COMMUN CO1 ON CO_TYPE="PPH" AND CO_CODE=ETB_PRUDH '
      +
      'LEFT JOIN ETABLISS ON ET_ETABLISSEMENT=ETB_ETABLISSEMENT ' + Habilitation + ' ORDER BY ETB_ETABLISSEMENT';
  end
  else
  begin
    st := 'SELECT ETB_ETABLISSEMENT,ET_LIBELLE,ETB_PRUDH,CO1.CO_LIBRE as TypeEmpl,ET_SIRET,ET_APE,ET_ADRESSE1,ET_ADRESSE2,ET_CODEPOSTAL,ET_VILLE FROM ETABCOMPL LEFT OUTER JOIN COMMUN CO1 ON CO_TYPE="PPH" AND CO_CODE=ETB_PRUDH '
      +
      'RIGHT OUTER JOIN ETABLISS ON ET_ETABLISSEMENT=ETB_ETABLISSEMENT ' + Habilitation + ' ORDER BY ETB_ETABLISSEMENT';
  end;

  LesEtab := TOB.Create('Les Etab', nil, -1);
  Q := OpenSql(st, TRUE);
  LesEtab.LoadDetailDB('ETABCOMPL', '', '', Q, False);
  LaGrille.RowCount := LesEtab.Detail.count;
  if LaGrille.RowCount = 1 then LaGrille.RowCount := 2;
  Nbre := 1;
  for i := 0 to LesEtab.Detail.count - 1 do
  begin
    T1 := LesEtab.Detail[i];
    if T1.getvalue('ETB_ETABLISSEMENT') = '' then continue;
    T1.addChampSup('Principal', FALSE);
    if T1.getvalue('ETB_ETABLISSEMENT') = Princ then st := 'Oui' else st := 'Non';
    T1.putvalue('Principal', st);
    T1.addChampSup('Suffixe', FALSE);
    if T1.getvalue('ETB_ETABLISSEMENT') = Princ then st := '00'
    else
    begin
      // PT2 22/04/02 PT1 PH Prise en compte cas > 100 etablissements alors suffixe en alpha numérique
      if Nbre < 100 then
      begin
        st := ColleZeroDevant(nbre, 2);
        Nbre := Nbre + 1;
      end
      else
      begin
        if Nbre = 100 then
        begin
          NbreS := 1;
          Indice := 1;
        end;
        st := Tableau[Indice + 10] + Tableau[NbreS];
        NbreS := NbreS + 1;
        Nbre := Nbre + 1;
        if NbreS > 36 then
        begin
          NbreS := 1;
          Indice := Indice + 1;
        end;
      end;
      // FIN PT2
    end;
    T1.putvalue('Suffixe', st);
    T1.addChampSup('LeCodeInsee', FALSE);
    T1.putvalue('LeCodeInsee', T1.getvalue('ET_CODEPOSTAL'));
    try
      if (T1.getvalue('ET_SIRET') = '') or (not ControlSiret(T1.getvalue('ET_SIRET'))) then
      begin
        PgiBox('Attention, le siret de l''établissement comptable ' + T1.getvalue('ETB_ETABLISSEMENT') + ' n''est pas conforme', Ecran.caption);
        Err := 1;
      end;
    except
      PgiBox('Attention, le siret de l''établissement comptable ' + T1.getvalue('ETB_ETABLISSEMENT') + ' n''est pas conforme', Ecran.caption);
      Err := 1;
    end;
    try
      if T1.getValue('ETB_PRUDH') = '' then
      begin
        PgiBox('Attention, le régime de cotisation de l''établissement comptable ' + T1.getvalue('ETB_ETABLISSEMENT') + ' n''est pas renseigné', Ecran.caption);
        Err := 1;
      end;
    except
      PgiBox('Attention, le régime de cotisation de l''établissement comptable ' + T1.getvalue('ETB_ETABLISSEMENT') + ' n''est pas renseigné', Ecran.caption);
      Err := 1;
    end;
  end;
  LesETab.PutGridDetail(THGRID(GetControl('GRILLEETAB')), FALSE, TRUE, 'ET_LIBELLE;TypeEmpl;Principal;Suffixe;LeCodeInsee', TRUE);
  for i := 0 to LaGrille.colCount - 1 do
    LaGrille.ColAligns[i] := taCenter;
  Ferme(Q);
  HMTrad := THSystemMenu(GetControl('HMTrad'));
  HMTrad.ResizeGridColumns(LaGrille);
  if Err = 0 then SetControlEnabled('BValider', TRUE);
end;

procedure TOF_PG_PrepPrudh.TheCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
  ControleCodeInsee(Lagrille.Cells[4, ARow], 'C');
end;
{ 2 types Action : Controle pour :
- affichage erreur en saisie de code insee
- génération d'une erreur fatale afin de ne générer de fichier rejeté
Controle sur le fait que le code insee soit numérique et comporte 5 cracteres
}
function TOF_PG_PrepPrudh.ControleCodeInsee(LeCode, Action: string): boolean;
var
  st: string;
begin
  result := TRUE;
  if (not IsNumeric(LeCode)) or (Length(LeCode) <> 5) then
  begin
    st := 'Le code insee n''est pas numérique ou ne comporte pas 5 chiffres';
    if Action = 'C' then PGIBox(st, Ecran.caption);
    result := FALSE;
  end;
end;
// Fonction ecriture de la tortalisation etablissement ==> nombre de salariés ecrites
function TOF_PG_PrepPrudh.EcritTotEtab(Nbre: Integer; TraceErr: TListBox): boolean;
var
  S: string;
  E: E300; // enregistrement de type total etablissement
begin
  with E do
  begin
    siret := Format_String(Lesiret, 14);
    suffixe := Format_String(Lesuffixe, 2);
    L1 := StringOfChar(' ', sizeof(L1));
    TypeEnr := '300';
    L2 := StringOfChar(' ', sizeof(L2));
    L3 := StringOfChar(' ', sizeof(L3));
    NbrePers := PGZeroAGauche(IntToStr(Nbre), 6);
    L4 := StringOfChar(' ', sizeof(L4));
  end;

  S := E.siret + E.suffixe + E.L1 + E.TypeEnr + E.L2 + E.L3 + E.NbrePers + E.L4;
  {$I-}Writeln(FEcrt, S);
  {$I+}if IoResult <> 0 then
  begin
    PGIBox('Erreur d''écriture du fichier prud''hom total établissement', 'Abandon du traitement');
    TraceErr.Items.Add('Erreur d''écriture du fichier prud''hom total établissement ');
    Result := FALSE;
  end;
end;

procedure TOF_PG_PrepPrudh.MajEnvoiSocial(NbreEnrg: Integer; NomFic: string);
var
st, An: string;
Ordre: Integer;
Q: TQuery;
AA: WORD;
DD, DF: TDateTime;
RaisonSoc: string; // Nom entreprise déclarée
EnregEnvoi: TEnvoiSocial;
begin
AA := StrToInt(Copy(GetControlText('LaDate'), 7, 4));
An := Copy(GetControlText('LaDate'), 7, 4);
An := Copy(An, 1, 1) + Copy(An, 3, 2);
//PT6
// formatage du millésime
St:= 'SELECT CO_CODE'+
     ' FROM COMMUN WHERE'+
     ' CO_TYPE="PGA" AND'+
     ' CO_LIBELLE="'+An+'"';
Q:=OpenSQL(St,TRUE) ;
if Not Q.EOF then
   An := Q.FindField ('CO_CODE').AsString;
Ferme (Q);
//FIN PT6
DD := EncodeDate(AA, 01, 01); // Date debut non utilisée mais remplie
DF := EncodeDate(AA, 12, 31); // Date fin non utilisée mais remplie
RaisonSoc := GetParamSoc('SO_LIBELLE');

St:= 'DELETE FROM ENVOISOCIAL WHERE'+
     ' PES_TYPEMESSAGE= "PRH" AND'+
     ' PES_MILLESSOC = "'+An+'" AND'+
     ' PES_DATEDEBUT = "'+UsDateTime(DD)+'" AND'+
     ' PES_DATEFIN = "'+UsDateTime(Df)+'" AND'+
     ' PES_SIRETDO = "'+LeSiretDos+'"  AND'+
     ' PES_FRACTIONDADS = "Z"';
ExecuteSQL(St);

St := 'SELECT MAX(PES_CHRONOMESS) AS MAXI FROM ENVOISOCIAL';
Q := OpenSQL(St, TRUE);
if not Q.EOF then
   Ordre := Q.FindField('MAXI').AsInteger + 1
else
   Ordre := 1;
Ferme(Q);

ChargeTOBENVOI; // correspond à la création de la TOB mère envoisocial
EnregEnvoi.Ordre := Ordre;
EnregEnvoi.TypeE := 'PRH';
{PT6
EnregEnvoi.Millesime := IntToStr(AA);
EnregEnvoi.Periodicite := An;
}
EnregEnvoi.Millesime:= An;
EnregEnvoi.Periodicite:= 'A';
//FIN PT6
EnregEnvoi.DateD := DD;
EnregEnvoi.DateF := DF;
EnregEnvoi.Inst := 'ZPRU';
EnregEnvoi.Siret := LeSiretDos;
EnregEnvoi.Fraction := 'Z';
EnregEnvoi.Libelle := RaisonSoc;
EnregEnvoi.Size := Arrondi((NbreEnrg * 565) / 1024, 4);
EnregEnvoi.NomFic := NomFic;
EnregEnvoi.Statut := '';
EnregEnvoi.Monnaie := 'EUR';
CreeEnvoi(EnregEnvoi);
LibereTOBENVOI; // Ecrit enreg
end;

procedure TOF_PG_PrepPrudh.TheGrilleExit(Sender: TObject);
begin
  GrilleMod := TRUE;
end;

procedure TOF_PG_PrepPrudh.OnClose;
begin
  inherited;
  if LesEtab <> nil then
  begin
    LesEtab.Free;
    LesEtab := nil;
  end;
end;

initialization
  registerclasses([TOF_PG_PrepPrudh]);
end.

