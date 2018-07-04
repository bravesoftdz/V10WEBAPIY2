unit Echg_Code;

interface

uses
    Sysutils
  , Messages
//*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/
  ,Windows, Classes, Graphics, Controls, Dialogs,
  StdCtrls, Grids, Mask, HTB97, ComCtrls,
  ParamSoc, uiutil, UlibWindows, UTOZ
  , ulanceprocess //  uLanceProcess
{$IFNDEF EAGLCLIENT}
  ,Fe_Main
{$ELSE}
  ,MainEagl
  ,UScriptTob
{$ENDIF}
//*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/
  , Forms
  , Menus
  , FileCtrl         // DirectoryExists
  , HPanel
  , ShellAPI         // DragQueryFile
  , Hctrls           // THCritMaskEdit
  , HmsgBox          // PgiInfo, PgiBox
  , Ent1             // VH : ^LaVariableHalley;
  , HEnt1            // V_PGI
  , UTOB
  , ColMemo
  , ULibEcriture     // CDateParDefautPourSaisie
  , UAssistImport    // ImportDonnees (ComSx)
  , utof
  , utoflookuptob
  , uYFILESTD        // BibleFichier (modes oératoires en pdf)
  , ed_Tools         // ProgressForm
  , ubob
  , LicUtil
  , UExport
  , USaveRestore
  , UDMIMP
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS}, dbtables{$ELSE}, uDbxDataSet, Variants{$ENDIF}
{$ENDIF}
  , UScript
  , UAssistConcept
  , AGLInit // TheTOB
  , cbpPath // TcbpPath.GetCegidUserTempPath
  ;

const
  ConstFicIniDemande = '\Echanges_LSD.INI';
  ConstFicIniRetour = '\Echanges_LSR.INI';
  ConstScriptEtafiConso = 'EE_EtafiCons';
//type

var
  gPP: THPanel;
  gDirEchDuDos: String;
  gScriptEtafiExiste: Boolean = False;
  gTOBScripts: TOB;
  gBALouGLJouAUT: String;

  function CorbeilleActivation: String;
  procedure ImportManuel (pScript, pFicDonnees: String; pHgVars: THGrid);
  procedure ImportAutomatique1;
  procedure FormateGrille (pGrille: THGrid);
  procedure LookUpA (pCme: THCritMaskEdit);
  procedure TraiteMessage (var vMsg: TMsg; var vHandled: Boolean);
  procedure EchgImp_LanceAssistant (pLequel: String);
  procedure VisuFicDonnees (NomCompletFicDonnees: String);
  procedure VisuModeOperatoire (pNomFicModop: String);
  function InitialiseLeBiniou: Boolean;
  procedure LanceComSx;
  function ConvertCISX (pScript, pFicDonnees: String; const pHgVars: THGrid = nil;
   const pIntermediaire: Boolean = False; const Muet: Boolean = False): Boolean;
  function VerifExistOrigineOuScript (Sender: TObject): Boolean;
  procedure TablesModif (pFamilleTable, pScript, pFicDonnees, pFicTable: String);
  procedure TraitePopUp (Sender: TObject; pScript, pFicDonnees: String);
  procedure CompteRenduComAffichage;
  procedure CompteRenduComSauvegarde(SaveDialog: TSaveDialog; ColorMemo: TColorMemo);
  procedure ParametresCharge (BALouGLJouAUT: String);
  procedure ParametresAffichage;
  procedure BoutonValiderClick;
  function PamametresAnnulation: Boolean;
  procedure TrtValueChange (Sender: TObject);
  function ParametresEnregistre: Boolean;
  procedure Kit;
  function TheDateDefaut: Tdate;
  procedure CreeFicIniPourCISX (pAction, pScript, pFicIn, pFicOut: String; pLsVars: TStringList = nil; pOption: String = ''; pFamille: String = '');
  function CopieTableAccess (Muet: Boolean): Integer; // SDA le 17/12/2007

implementation

uses
  EchgImp_Assistant
  , EchgImp_Parametres
  , EchgImp_CompteRendu
  ;

var
  gFicTRA: String;
  gSave_Cursor: TCursor; // Pour sauvegarder le curseur en cours
  // Handle de l'objet d'affichage du nom de fichier à traiter
  //  pour glisser/déposer
  HandleObjNomFich: HWND;

procedure ParametresDisable;
// Désactivation de tous les paramètres
begin
  with fEchgImp_Parametres do
  begin
    labDate.Enabled := False;
    cmeDate.Enabled := False;
    labLibelle.Enabled := False;
    HValCB_Libelle.Enabled := False;
    CB_AvecDateImport.Enabled := False;
    libJournal.Enabled := False;
    Hcme_Journal.Enabled := False;
    libCompte.Enabled := False;
    Hcme_Compte.Enabled := False;
    //Hcme_Etablissement.Enabled := False;
  end;
end;

function ScriptsListeVersTob: Boolean;
var
  i         : Integer;
  TF1       : TOB;
  S         : String;
  Q         : TQuery;
begin
  Result := True;

  //LocCbOrigine.Values.Add ('***');
  //LocCbOrigine.Items.Add ('*** Reconnaissance automatique ***');

  if gTOBScripts <> nil then gTOBScripts.Free;
  gTOBScripts := TOB.Create ('Scripts', nil, -1);
  S := 'select CIS_TABLE1, CIS_CLE, CIS_COMMENT, CIS_TABLE2, CIS_DOMAINE ' +
   ' from CPGZIMPREQ' +
   ' where CIS_NATURE<> "X"' + // pas de délimités
   ' and (CIS_DOMAINE="E" or CIS_DOMAINE="O") ';

  if gBALouGLJouAUT = 'BAL' then
    S := S + ' and CIS_TABLE2="Balance" '
  else if gBALouGLJouAUT = 'GLJ' then
    S := S + ' and CIS_TABLE2="Journal" '
  ;

  S := S + ' order by CIS_TABLE1,CIS_COMMENT,CIS_DOMAINE ';
  Q := OpenSQL (S, FALSE);

  gTOBScripts.LoadDetailDB ('PGZIMPREQ', '', '', Q, TRUE, FALSE);
  Ferme (Q);
  if gTOBScripts.detail.count > 0 then
  begin
    for i := 0 to gTOBScripts.detail.Count -1 do
    begin
      TF1 := gTOBScripts.Detail[i];
      if not gScriptEtafiExiste then
        if TF1.GetValue ('CIS_CLE') = ConstScriptEtafiConso then
          gScriptEtafiExiste := True;
      TF1.AddChampSupValeur ('DomScript', TF1.GetValue ('CIS_DOMAINE') + TF1.GetValue ('CIS_CLE'));
    end;
    if gTOBScripts.detail.Count < 1 then // Cas liste scripts vide
    begin
      PGIError (#13 + 'Liste des scripts CISX vide.' + #13 + #13 +
       'Traitement abandonné.', 'Echanges');
      Result := False;
    end;
  end
  else
  begin
    gTOBScripts := nil ;
    PGIError (#13 + 'Liste des scripts CISX non trouvée.' + #13 + #13 +
     'Traitement abandonné.', 'Echanges');
    Result := False;
  end;
end;

procedure VariablesGrilleVisible (pEtat: Boolean);
begin
  with fEchgImp_Assistant do
  begin
    if pEtat then
    begin
      //PanelH.Height := 1;
      //PanelV.Visible := True;
      hgVariables.Visible := True;
    end
    else
    begin
      //PanelV.Visible := False;
      //PanelH.Height := Round (PanelV.Height / 2);
      hgVariables.Visible := False;
    end;
    hgVariables.Refresh;
    //Refresh;
  end;
end;

function TheDateDefaut: Tdate;
var
  DateDebut, DateFin: TDateTime;
begin
  CDateParDefautPourSaisie (DateDebut, DateFin);
  // Si aucune cloture périodique => utiliser date début d'exercice
  if DateDebut < VH^.EnCours.Deb then DateDebut := VH^.EnCours.Deb;
  Result := DateDebut;
end;

procedure CreeFicTableSiInexist (pNom: String);
var
  FLoc: Integer;
  Nomfich: String;
begin
  NomFich := gDirEchDuDos + '\' + pNom;
  if not FileExists (NomFich) then
  begin
    FLoc := FileCreate (NomFich);
    FileClose (FLoc);
  end;
end;

procedure CurseurSablier;
// Sauvegarde le curseur en cours dans var globale Save_Cursor
//  et affiche le curseur sablier
begin
  gSave_Cursor := Screen.Cursor;
  Screen.Cursor := crHourglass;
end;

procedure RestoreCurseur;
// Restore le curseur sauvegardé par précédement (CurseurSablier)
begin
  //Screen.Cursor := gSave_Cursor;
  Screen.Cursor := crDefault;
end;

procedure CreeFicIniPourCISX (pAction, pScript, pFicIn, pFicOut: String;
 pLsVars: TStringList = nil; pOption: String = ''; pFamille: String = '');
var
  FLoc: TextFile;
  i: Integer;
  S, FicIni, Repertoire, Domaine, Mode: String;
  Chemindos  : string;
begin
  Domaine := copy (pScript, 1, 1);
  Delete (pScript, 1, 1);

  FicINI := gDirEchDuDos + ConstFicIniDemande;
  SysUtils.DeleteFile (FicIni);

  // Effacement du fichier de sortie pour être sûr de ne pas utiliser retours précédents
  if ((pAction <> 'TableConversion') AND (pAction <> 'ListeScripts2')) then
    SysUtils.DeleteFile (pFicOut);

  AssignFile (FLoc, FicIni);
  Rewrite (FLoc);

  WriteLN (FLoc, '[PRODUIT]');
  WriteLN (FLoc, 'APPELANT=' + 'ECHANGES PGI');

  if ((pAction = 'TableConversion') and (Domaine = 'O')) then
    WriteLN (FLoc, 'DOMAINE=E')
  else
  begin
    if ((pAction = 'TRA') and (Domaine = 'O')) then
      WriteLN (FLoc, 'DOMAINE=' + Domaine + ';E')
    else
      WriteLN (FLoc, 'DOMAINE=' + Domaine);
  end;

  if ((pAction = 'Variables') or (Pos ('ListeScripts', pAction) = 1)) then
    Mode := 'RETOUR'
  else if  Pos ('TRA', pAction) = 1 then Mode := 'EXPORT'
  else
  if (pAction = 'RESTOR') then
     WriteLN (FLoc, 'MODE=' + pAction)
  else
     Mode := 'INTERFACE';

  WriteLN (FLoc, 'MODE=' + Mode);

  WriteLN (FLoc, 'OPTION=' + pOption);

  if pFamille <> '' then
    WriteLN (FLoc, 'FAMILLE=' + pFamille);

// ME deb
  if (V_PGI.ModePCL='1') then
     WriteLN (FLoc, 'REPSTD=' + ChangeStdDatPath ('$STD\CISX'))
  else
     WriteLN (FLoc, 'REPSTD=' + ExtractFilePath (Application.ExeName)+'parametre');
// ME fin


(* CA - 05/07/2007 - Remplacement DosPath car n'existe pas en EAGL

  if V_PGI.DosPath[length (V_PGI.DosPath)] <> '\' then Chemindos := V_PGI.DosPath + '\CISX'
    else  Chemindos := V_PGI.DosPath+'CISX' ;

*)
  Chemindos := TcbpPath.GetCegidUserTempPath+'CISX' ;

  WriteLN (FLoc, 'REPDOS=' + Chemindos);
  FicINI := gDirEchDuDos + ConstFicIniDemande;

  if ((pAction = 'TRA') and (Domaine = 'O')) then
  begin
    WriteLN (FLoc, 'LISTECOMMANDE=COMMANDE1;COMMANDE2');
    WriteLN (FLoc, #13 + #10 + '[COMMANDE1]');
  end
  else
    WriteLN (FLoc, #13 + #10 + '[COMMANDE]');
  WriteLN (FLoc, 'MONOTRAITEMENT=1');
  WriteLN (FLoc, 'COMPRESS=0'); WriteLN (FLoc, '');

  if ((pAction = 'TableConversion') and (Domaine = 'O')) then
    WriteLN (FLoc, 'SCRIPT=' + ConstScriptEtafiConso)
  else
    WriteLN (FLoc, 'SCRIPT=' + pScript);

  Repertoire := ExtractFilePath (pFicOut);
  i := Length (Repertoire);
  if Repertoire[i] = '\' then
    System.Delete (Repertoire, i, 1);
  WriteLN (FLoc, 'REPERTOIRE=' + Repertoire);

  if ((pAction = 'TRA') and (Domaine = 'O')) then
    WriteLN (FLoc, 'NOMFICHIER=TraTmp.TRA')
  else
    WriteLN (FLoc, String ('NOMFICHIER=' + ExtractFileName (pFicOut)));
  WriteLN (FLoc, String ('LISTEFICHIER=' + pFicIn + #13 + #10));

  if ((pAction = 'TRA') and (Domaine = 'O')) then
  begin
    WriteLN (FLoc, #13 + #10 + '[COMMANDE2]');
    WriteLN (FLoc, 'MONOTRAITEMENT=1');
    WriteLN (FLoc, 'COMPRESS=0'); WriteLN (FLoc, '');
    WriteLN (FLoc, 'SCRIPT=' + ConstScriptEtafiConso);

    Repertoire := ExtractFilePath (pFicOut);
    i := Length (Repertoire);
    if Repertoire[i] = '\' then
      System.Delete (Repertoire, i, 1);
    WriteLN (FLoc, 'REPERTOIRE=' + Repertoire);

    WriteLN (FLoc, String ('NOMFICHIER=' + ExtractFileName (pFicOut)));
    WriteLN (FLoc, String ('LISTEFICHIER=' + Repertoire + '\TraTmp.TRA' + #13 + #10));
  end;

  if ((Mode = 'EXPORT') and (pAction <> 'TRAinterm')) then
  begin
    WriteLN (FLoc, '[CORRESP]');
    WriteLN (FLoc, 'Corresp1=COMPTE=' + ExtractFilePath (pFicOut) + 'ConvComptes.txt');
    WriteLN (FLoc, 'Corresp2=CODEJOURNAL=' + ExtractFilePath (pFicOut) + 'ConvJournaux.txt' +
     #13 + #10);

    //-----> Variables
    WriteLN (FLoc, '[VARIABLE]');
    if Domaine <> 'X' then
    begin
    if gBALouGLJouAUT = 'BAL' then
    begin
      WriteLN (FLoc, 'VARIABLE1' + '=' + '*DftDat=' + FormatDateTime ('ddmmyyyy',
       GetParamSocSecur ('SO_CPEDIDATEVAL', dateToStr(now))));

      S := GetParamSocSecur ('SO_CPEDILIBELLE', '');
      if GetParamSocSecur ('SO_CPEDILIBELLEDATE', False) then
        S := S + ' ' + FormatDateTime ('dd/mm/yyyy', now);
      WriteLN (FLoc, 'VARIABLE2' + '=' + '*DftLib=' + Copy (S, 1, 35));
      WriteLN (FLoc, 'VARIABLE3' + '=' + '*DftJnl=' + GetParamSocSecur ('SO_CPEDIJOURNAL', ''));
      WriteLN (FLoc, 'VARIABLE4' + '=' + '*DftCpt=' + GetParamSocSecur ('SO_CPEDICOMPTE', ''));
    end
    else if gBALouGLJouAUT = 'GLJ' then
    begin
      WriteLN (FLoc, 'VARIABLE1' + '=' + '*DftDat=' + FormatDateTime ('ddmmyyyy',
       GetParamSocSecur ('SO_CPEDIDATEVALG', dateToStr(now), TRUE)));

      S := GetParamSocSecur ('SO_CPEDILIBELLEG', '');
      if GetParamSocSecur ('SO_CPEDILIBELLEDATEG', False) then
        S := S + ' ' + FormatDateTime ('dd/mm/yyyy', now);
      WriteLN (FLoc, 'VARIABLE2' + '=' + '*DftLib=' + Copy (S, 1, 35));
      WriteLN (FLoc, 'VARIABLE3' + '=' + '*DftJnl=' + GetParamSocSecur ('SO_CPEDIJOURNALG', ''));
      WriteLN (FLoc, 'VARIABLE4' + '=' + '*DftCpt=' + GetParamSocSecur ('SO_CPEDICOMPTEG', ''));
    end
    else if gBALouGLJouAUT = 'AUT' then
    begin
      WriteLN (FLoc, 'VARIABLE1' + '=' + '*DftDat=' + FormatDateTime ('ddmmyyyy',
       GetParamSocSecur ('SO_CPEDIDATEVALA', dateToStr(now), TRUE)));

      S := GetParamSocSecur ('SO_CPEDILIBELLEA', '');
      if GetParamSocSecur ('SO_CPEDILIBELLEDATEA', False) then
        S := S + ' ' + FormatDateTime ('dd/mm/yyyy', now);
      WriteLN (FLoc, 'VARIABLE2' + '=' + '*DftLib=' + Copy (S, 1, 35));
      WriteLN (FLoc, 'VARIABLE3' + '=' + '*DftJnl=' + GetParamSocSecur ('SO_CPEDIJOURNALA', ''));
      WriteLN (FLoc, 'VARIABLE4' + '=' + '*DftCpt=' + GetParamSocSecur ('SO_CPEDICOMPTEA', ''));
    end;

    if pLsVars <> nil then
      for i := 1 to pLsVars.Count do
        WriteLN (FLoc, 'VARIABLE' + intToStr (i + 4) + '=' + pLsVars[I - 1]);
    end
    else
    begin
         if pLsVars <> nil then
            for i := 1 to pLsVars.Count do
              WriteLN (FLoc, 'VARIABLE' + intToStr (i) + '=' + pLsVars[I - 1]);
    end;

  end; // if ((Mode = 'EXPORT') and (pAction <> 'TRAinterm')) then

  CloseFile (FLoc);
end;

procedure FormateGrille (pGrille: THGrid);
begin
  // Mise en forme du THGrid
  with pGrille do
  begin
    if RowCount > 1 then
    begin
      FixedRows := 1;
      TitleCenter := True;
      TitleBold := True;
      ColEditables[0] := False;
      ColEditables[1] := True;
      ColWidths[0] := Round ((pGrille.Width - 23) / 2);
      ColWidths[1] := Round ((pGrille.Width - 23) / 2);
      Hint := 'Cliquer sur la colonne Valeur pour modifier la variable';
      VariablesGrilleVisible (True);
    end
    else
    begin
      VariablesGrilleVisible (False);
    end;
  end;
end;

function CompteRenduNomDuFichierAvant: String;
var
  Stt, Doss: String;
begin
  if not (ctxPCL in V_PGI.PGIContexte) then
  begin
    Stt := ExtractFileName (gFicTRA);
    Doss := ReadTokenPipe (Stt, '.');
    Result := gFicTRA + 'ListeCom' + Doss;
  end
  else
    Result := gFicTRA + 'ListeCom' + V_PGI.NoDossier;
end;

function CompteRenduNomDuFichierApres: String;
var
  S: String;
begin
  S := CompteRenduNomDuFichierAvant;
  Result := ExtractFilePath (S) + '_' + ExtractFileName (S);
end;

procedure CompteRenduComComplete;
var
  S, FileName: String;
  TS1, TS2: TStringList;
begin
  FileName := CompteRenduNomDuFichierAvant;

  TS1 := TStringList.Create;
  TS1.Clear;

  if gBALouGLJouAUT = 'BAL' then S := 'Balances'
  else if gBALouGLJouAUT = 'GLJ' then S := 'Grands-livres - Journaux'
  else S := 'Automatique';
  TS1.Add ('   Rapport d''importation (' + S + ')');

  if FileExists (FileName + '.OK') then
  begin
    FileName := FileName + '.OK';
    TS1.Add ('   Statut : OK');
  end
  else
  begin
    FileName := FileName + '.ERR';
    TS1.Add ('   Statut : ERREUR');
  end;

  TS1.Add ('   Le : ' + DateToStr (now) + ' à '  + TimeToStr (now));
  TS1.Add ('');

  if gBALouGLJouAUT = 'BAL' then S := GetParamSocSecur ('SO_CPEDIORIGINELIB', '')
  else if gBALouGLJouAUT = 'GLJ' then S := GetParamSocSecur ('SO_CPEDIORIGINELIBG', '')
  else S := GetParamSocSecur ('SO_CPEDIORIGINELIBA', '');
  TS1.Add ('   Solution : ' + S);

  if gBALouGLJouAUT = 'BAL' then S := GetParamSocSecur ('SO_CPEDIORIGINE', '')
  else if gBALouGLJouAUT = 'GLJ' then S := GetParamSocSecur ('SO_CPEDIORIGINEG', '')
  else S := GetParamSocSecur ('SO_CPEDIORIGINEA', '');
  TS1.Add ('   Script : ' + S);

  if gBALouGLJouAUT = 'BAL' then S := GetParamSocSecur ('SO_CPEDIFICSOURCE', '')
  else if gBALouGLJouAUT = 'GLJ' then S := GetParamSocSecur ('SO_CPEDIFICSOURCEG', '')
  else S := corbeilleActivation;
  TS1.Add ('   Fichier : ' + S);
  TS1.Add ('');
  TS1.Add ('');

//: String; pHgVars: THGrid

//parametres
//variables demandables
//fichier
//param

  TS2 := nil;
  try
    TS2 := TStringList.Create;

    if FileExists (FileName) then
      TS2.LoadFromFile (FileName)
    else
      TS2.Text :=  #13 + #10 +  #13 + #10 + #13 + #10 + '***** RAPPORT INTROUVABLE ! *****';

    TS1.AddStrings (TS2);
  finally
    TS2.Free;
  end;
  TS1.SaveToFile (ExtractFilePath (FileName) + '_' + ExtractFileName (FileName));
  TS1.Free;
  SysUtils.DeleteFile (FileName);
end;

procedure LanceComSx;
begin
  //ImportDonnees (gFicTRA + ';Minimized', TRUE, nil);
  ImportDonnees (gFicTRA + ';DECOUPELON=99;Minimized', TRUE, nil, '', '', TRUE); // Ajout me TRUE pour signifier mode Echange
  SysUtils.DeleteFile (gDirEchDuDos + '\TraTmp.TRA');
  //SysUtils.DeleteFile (ChangeFileExt (gFicTRA, '.ARC'));
  //SysUtils.DeleteFile (ChangeFileExt (gFicTRA, '.SAV'));
  //SysUtils.DeleteFile (gFicTRA);

  CompteRenduComComplete;
end;

function ConvertCISX (pScript, pFicDonnees: String; const pHgVars: THGrid = nil;
const pIntermediaire: Boolean = False; const Muet: Boolean = False): Boolean;
var
  FicINI: String;
  lsVars: TStringList;
  i: Integer;
begin
  if not FileExists (pFicDonnees) then
  begin
    Result := False;
    IF NOT Muet THEN PGIError (#13 + 'Le fichier ' + #13 + pFicDonnees + #13 +
     'n''existe pas. Traitement abandonné.', 'Echanges');
    EXIT;
  end;

  FicINI := gDirEchDuDos + ConstFicIniDemande;
  lsVars := TStringList.Create;
  lsVars.Clear;
  if pHgVars <> nil then
    for i := 2 to pHgVars.RowCount do
      if pHgVars.Cells[0, i - 1] <> '' then
        lsVars.Add (pHgVars.Cells[2, i - 1] + '=' + pHgVars.Cells[1, i - 1]);

  if pIntermediaire then
    CreeFicIniPourCISX ('TRAinterm', pScript, pFicDonnees, gDirEchDuDos +
     '\TraTmp.TRA', lsVars)
  else
    CreeFicIniPourCISX ('TRA', pScript, pFicDonnees, gFicTRA, lsVars);
  lsVars.Free;

  Result := LanceExportCisx ('/INI=' + FicINI);
end;

procedure TraiteMessage (var vMsg: TMsg; var vHandled: Boolean);
var
  NomDuFichierStr: string;
  NomDuFichier: array[0..255] of char;
begin
  if vMsg.message = WM_DROPFILES then
  begin
    if vMsg.hwnd = HandleObjNomFich then
    begin
      DragQueryFile (vMsg.wParam, 0, NomDuFichier, sizeof (NomDuFichier));// récupération du nom du fichier
      NomDuFichierStr := NomDuFichier; // tansformation du tableau de char en STRING
      DragFinish (vMsg.wParam);
      if not DirectoryExists (NomDuFichierStr) then  // Ne pas traiter si dossier
        SetWindowText (vMsg.hwnd, NomDuFichier);
      vHandled := True;
    end;
  end;
end;

procedure VariableDefautActivation (pName: String);
begin
  with fEchgImp_Parametres do
  begin
    if pName = '*DftDat' then
    begin
      labDate.Enabled := True;
      cmeDate.Enabled := True;
    end else
    if pName = '*DftLib' then
    begin
      labLibelle.Enabled := True;
      HValCB_Libelle.Enabled := True;
      CB_AvecDateImport.Enabled := True;
    end else
    if pName = '*DftJnl' then
    begin
      libJournal.Enabled := True;
      Hcme_Journal.Enabled := True;
    end else
    if pName = '*DftCpt' then
    begin
      libCompte.Enabled := True;
      Hcme_Compte.Enabled := True;
    end;
  end;
end;

function VariablesAfficheListe: Boolean;
var
  i, L: Integer;
  TOBLoc, TF1: TOB;
  GV: THGrid;
begin
  Result := True;

  GV := fEchgImp_Assistant.hgVariables;
  TOBLoc := TOB.Create ('Variables', nil, -1);
  // TF1 := TOB.Create ('', TOBLoc, -1);

  GV.VidePile (False);
  L := 1;
  GV.RowCount := L;
  if TobLoadFromFile (gDirEchDuDos + ConstFicIniRetour, nil, TOBLoc) then
  begin
    // Affichage des variables demandables dans la grille et
    // variables défaut chargées dans fiche fEchgImp_Parametres
    for i := 0 to TOBLoc.detail.Count -1 do
    begin
      TF1 := TOBLoc.Detail[i];
      if (TF1.FieldExists ('DEMANDABLE'))
       and (TF1.GetValue ('DEMANDABLE') = TRUE) then
      begin
        if Copy (TF1.GetValue ('NAME'), 1, 4) = '*Dft' then
          VariableDefautActivation (TF1.GetValue ('NAME'))
        else
        begin
          Inc (L);
          GV.RowCount := L;
          GV.Cells[0, L - 1] := TF1.GetValue ('LIBELLE');
          GV.Cells[1, L - 1] := TF1.GetValue ('TEXT');
          GV.Cells[2, L - 1] := TF1.GetValue ('NAME');
        end;
      end;
    end;
    FormateGrille (GV);
  end
  else
  begin
    PGIError (#13 + 'Liste des variables du script CISX non trouvée.' + #13 + #13 +
     'Traitement abandonné.', 'Echanges');
    //fEchgImp_Assistant.Close;
    Result := False;
  end;
  TOBLoc.Free;
end;

function VariablesDuScriptDemande (pCodeScript: String): Boolean;
begin
  RESULT := FALSE;
  if InitScriptAuto (Copy (pCodeScript, 2, length (pCodeScript))) > 0 then
    Result := True;
end;

procedure VariablesDuScript_ (pCodeScript: String);
begin
  CurseurSablier;
  with fEchgImp_Assistant do
  begin
    ParametresDisable;

    if pos ('O', pCodeScript) = 1 then
      pCodeScript := 'O' + ConstScriptEtafiConso;

    VariablesDuScriptDemande (pCodeScript);
    VariablesAfficheListe;
  end;
  fEchgImp_Assistant.Refresh;
end;

function CopieTableAccess (Muet: Boolean): Integer;
var
  Chemindos, cheminstd,FichierAIntegrer : string;
  CodeRetour: integer;
  zip: TOZ;
begin
  Chemindos := GetWindowsTempPath  + 'CISX';
  if not DirectoryExists (Chemindos) then
    if not CreateDir (Chemindos) then
    begin
      Result := -1;
      IF NOT Muet THEN PGIError (#13 + 'Impossible de créér le répertoire ' + #13 +
       Chemindos + #13 + '. Traitement abandonné.', 'Echanges');
      EXIT;
    end;

  CodeRetour := AGL_YFILESTD_EXTRACT (FichierAIntegrer, 'CISX', 'PGZIMPACCESSD.ZIP',
   'C', 'COMPTA', 'PARAM', '', '', FALSE, V_PGI.LanguePrinc, 'CEG');
  if CodeRetour <> -1 then
  begin
    Result := -2;
    IF NOT Muet THEN PGIError (#13 + 'Erreur d''accès à la base Access. Traitement abandonné.',
     'Echanges');
    EXIT;
  end
  else
  begin
    zip := TOZ.Create;
    zip.OpenZipFile (FichierAIntegrer, moOpen);
    if not zip.OpenSession (osExt) then
    begin
      Result := -3;
      IF NOT Muet THEN PGIError (#13 + 'Echec ouverture zip. Traitement abandonné.',
       'Echanges');
      zip.Free;
      EXIT;
    end
    else
    begin // extraction
      sysutils.DeleteFile (Chemindos + '\PGZIMPACCESSD.MDB');
      zip.SetDirOut (Chemindos);
      zip.CloseSession;
      zip.Free;
    end;
  end;
  Cheminstd := Chemindos;
  InitDB (Cheminstd, Chemindos);
  Result := 1;
end;

function InitialiseLeBiniou: Boolean;
begin
  Result := False;
  if CopieTableAccess (False) <> 1 then EXIT;

(* CA - 05/07/2007 - Remplacement DosPath car n'existe pas en EAGL

  if V_PGI.DosPath[length (V_PGI.DosPath)] <> '\' then
    gDirEchDuDos := V_PGI.DosPath + '\Echanges'
  else gDirEchDuDos := V_PGI.DosPath + 'Echanges';

*)

  gDirEchDuDos := TcbpPath.GetCegidUserTempPath + 'Echanges';

  if Not DirectoryExists (gDirEchDuDos) then
    CreateDir (gDirEchDuDos);

  gFicTRA := gDirEchDuDos + '\PG' + V_PGI.CodeSociete + '.TRA';
  CreeFicTableSiInexist ('ConvComptes.txt');
  CreeFicTableSiInexist ('ConvJournaux.txt');

  fEchgImp_Parametres := TfEchgImp_Parametres.Create (nil);

  ScriptsListeVersTob;

  with fEchgImp_Assistant do
  begin
    hgVariables.VidePile (False);
    hgVariables.ColWidths[2] := -1;

    if gBALouGLJouAUT = 'BAL' then
    begin
      panAuto.Visible := False;
      fEchgImp_Assistant.Caption := 'Importation Balances';
      cmeNomFich.Text := GetParamSocSecur ('SO_CPEDIFICSOURCE', '');
      Hcme_Solutions.Plus := GetParamSocSecur ('SO_CPEDIORIGINE', '');
      Hcme_Solutions.Text := GetParamSocSecur ('SO_CPEDIORIGINELIB', '');

      if IsValidDate (GetParamSocSecur ('SO_CPEDIDATEVAL', idate1900)) then
        cmeDate.Text := GetParamSocSecur ('SO_CPEDIDATEVAL', dateToStr(now))
      else
        cmeDate.Text := dateToStr(now);
    end
    else if gBALouGLJouAUT = 'GLJ' then
    begin
      panAuto.Visible := False;
      panDateBal.Visible := False;
      fEchgImp_Assistant.Caption := 'Importation Grands-livres - Journaux';
      cmeNomFich.Text := GetParamSocSecur ('SO_CPEDIFICSOURCEG', '');
      Hcme_Solutions.Plus := GetParamSocSecur ('SO_CPEDIORIGINEG', '');
      Hcme_Solutions.Text := GetParamSocSecur ('SO_CPEDIORIGINELIBG', '');
    end
    else if gBALouGLJouAUT = 'AUT' then
    begin
      fEchgImp_Assistant.Caption := 'Importation automatique';
      fEchgImp_Assistant.cmeNomFich.Enabled := False;
      fEchgImp_Assistant.cmeNomFich.ElipsisButton := False;

      Hcme_Corbeille.Text := GetParamSocSecur ('SO_CPEDICORBEILLE', '');
      Hcme_Masque.Text := GetParamSocSecur ('SO_CPEDIMASQUE', '');
      cmeNomFich.Text := CorbeilleActivation;
      Hcme_Solutions.Plus := GetParamSocSecur ('SO_CPEDIORIGINEA', '');
      Hcme_Solutions.Text := GetParamSocSecur ('SO_CPEDIORIGINELIBA', '');

      BValider2.Top := BValider.Top;
      BValider2.Left := BValider.Left;
      BValider2.Visible := True;
      BValider.Visible := False;
    end
    else
      EXIT
    ;
    UpdateCaption (fEchgImp_Assistant);

    //if cmeNomFich.Text = '' then TrtValueChange (cmeNomFich);
    //if Hcme_Solutions.Text = '' then TrtValueChange (Hcme_Solutions);

    HandleObjNomFich := cmeNomFich.Handle;
    DragAcceptFiles (HandleObjNomFich, True);  // Pour glisser nom de fichier
  end;

  RestoreCurseur;
  Result := True;
end;

function TesteExistModOp (pNomFicModop: String): Boolean;
var
  FilePath: String;
  Res: Integer;
begin
  FilePath := '';
  Delete (pNomFicModop, 1, 1);
  Res :=  AGL_YFILESTD_EXTRACT (FilePath, 'COMPTA', pNomFicModop +
   '.PDF', 'ECH', 'MOP', '', '', '', FALSE, V_PGI.LanguePrinc, 'CEG' );
  if Res <> -1 then
    Res :=  AGL_YFILESTD_EXTRACT (FilePath, 'COMPTA', pNomFicModop +
     '.PDF', 'ECH', 'MOP', '', '', '', FALSE, V_PGI.LanguePrinc, 'DOS' );
  if Res = -1 then
    Result := True
  else
    Result := False;
end;

function VerifExistOrigineOuScript (Sender: TObject): Boolean;
begin
  Result := False;
  with fEchgImp_Assistant do
  begin
    if Sender = Hcme_Solutions then
    begin
      btnModeOp.Enabled := False;
      btnParametres.Enabled := False;
      btnTables.Enabled := False;
      if Assigned (gTOBScripts) then
        if gTOBScripts.FindFirst (['DomScript'], [Hcme_Solutions.Plus], True) <> nil then
        begin
          Result := True;
          btnParametres.Enabled := True;
          btnTables.Enabled := True;
          if TesteExistModOp (Hcme_Solutions.Plus) then
            btnModeOp.Enabled := True;
        end;
    end
    else if Sender = cmeNomFich then
    begin
      btnVisuFciDonnees.Enabled := False;
      btnTables.Enabled := False;
      if FileExists (cmeNomFich.Text) then
      begin
        Result := True;
        btnVisuFciDonnees.Enabled := True;
        btnTables.Enabled := True;
      end;
    end;

    // désactivation du bouton traiter si fichier ou param inexistant
    //BValider.Enabled  := btnVisuFciDonnees.Enabled and btnModeOp.Enabled;
    BValider.Enabled := btnTables.Enabled;
  end;
end;

function ScriptType (pScript: String): String;
var
  TOBi: TOB;
begin
  Result := '';
  TOBi := gTOBScripts.FindFirst (['DomScript'], [pScript], True);
  if TOBi <> nil then
    Result := TOBi.GetValue ('CIS_TABLE2');
end;

Procedure ScruteFichier (Dossier: string; filtre: string; Attributs: integer;
 var Conteneur: TstringList);
// Retourne, dans une TstringList 'Conteneur', la liste des fichiers du
//  'Dossier'  selon 'filtre' et 'Attributs
var
  FichierTrouve: string;
  Resultat: Integer;
  SearchRec: TSearchRec;
begin
  Conteneur.Clear;     // Voir si bonne methode //************************
  if Dossier[length (Dossier)] = '\' then
    Dossier := copy (Dossier, 1, length (Dossier) -1);
  Resultat := FindFirst (Dossier + '\' + filtre, Attributs, SearchRec);
  while Resultat = 0 do
  begin
    Application.ProcessMessages;
    if ((SearchRec.Attr and faDirectory) <= 0) then // On a trouvé un Fichier (et non un dossier)
    begin
      FichierTrouve := Dossier + '\' + SearchRec.Name;
    end;
      Conteneur.Add (FichierTrouve);
    Resultat := FindNext (SearchRec);
  end;
  SysUtils.FindClose (SearchRec); // libération de la mémoire
end;

function CorbeilleActivation: String;
var
  Corbeille, Masque: String;
  TSL: TstringList;
begin
  Result := '';
  Corbeille := GetParamSocSecur ('SO_CPEDICORBEILLE', '');
  Masque := GetParamSocSecur ('SO_CPEDIMASQUE', '');
  Corbeille := Trim (Corbeille);

  if ((Corbeille = '') or (Masque = '') or (not DirectoryExists (Corbeille))) then
    EXIT;

  TSL := TstringList.Create;

  ScruteFichier (Corbeille, Masque, faAnyFile, TSL);
  if TSL.Count > 0
  then Result := TSL.Strings[0]
  else Result := '';

  TSL.Free;
EXIT;

{ TDirectory
 // uses QFileCtrls
  Corbeille := GetParamSocSecur ('SO_CPEDICORBEILLE', '');
  Masque := GetParamSocSecur ('SO_CPEDIMASQUE', '');

  Corbeille := Trim (Corbeille);

  D := TDirectory.Create (nil);
  D.BeginUpdate;
  D.Location := Corbeille;
  D.FileMask := Masque;
  D.EndUpdate;

  if D.Count > 1
  then Result := D.Location + D.Caption (1)
  else  Result := '';

  D.Free;
}

{  OLD
  Corbeille := GetParamSocSecur ('SO_CPEDICORBEILLE', '');
  Masque := GetParamSocSecur ('SO_CPEDIMASQUE', '');

  if ((Trim (Corbeille) <> '') and (Trim (Masque) <> '') and
   (FileExists (Corbeille + '\' + Masque)))
  then Result := Corbeille + '\' + Masque
  else  Result := '';
}
end;

procedure TrtValueChange (Sender: TObject);
begin
  with fEchgImp_Assistant do
  begin
    if Sender = Hcme_Corbeille then
    begin
      SetParamSoc ('SO_CPEDICORBEILLE', Hcme_Corbeille.Text);
      cmeNomFich.Text := CorbeilleActivation;
    end
    ELSE
    if Sender = Hcme_Masque then
    begin
      SetParamSoc ('SO_CPEDIMASQUE', Hcme_Masque.Text);
      cmeNomFich.Text := CorbeilleActivation;
    end
    ELSE
    if Sender = Hcme_Solutions then
    begin
      if VerifExistOrigineOuScript (Sender) then
      begin
        if gBALouGLJouAUT = 'BAL' then
        begin
          SetParamSoc ('SO_CPEDIORIGINE', Hcme_Solutions.Plus);
          SetParamSoc ('SO_CPEDIORIGINELIB', Hcme_Solutions.Text);
        end
        else if gBALouGLJouAUT = 'GLJ' then
        begin
          SetParamSoc ('SO_CPEDIORIGINEG', Hcme_Solutions.Plus);
          SetParamSoc ('SO_CPEDIORIGINELIBG', Hcme_Solutions.Text);
        end
        else if gBALouGLJouAUT = 'AUT' then
        begin
          SetParamSoc ('SO_CPEDIORIGINEA', Hcme_Solutions.Plus);
          SetParamSoc ('SO_CPEDIORIGINELIBA', Hcme_Solutions.Text);
        end;
      end; // VerifExistOrigineOuScript
      if ScriptType (fEchgImp_Assistant.Hcme_Solutions.Plus) = 'Balance' then
      begin
        cmeDate.Text := dateToStr(now);
        HLabDate.Visible := True;
        cmeDate.Visible := True;
      end
      else
      begin
        HLabDate.Visible := False;
        cmeDate.Visible := False;
      end; // criptType (fEchgImp_Assistant.Hcme_Solutions.Plus) = 'Balance'
      VariablesDuScript_ (Hcme_Solutions.Plus);
    end ELSE
    if Sender = cmeNomFich then
    begin
      if VerifExistOrigineOuScript (Sender) then
      begin
        if length (cmeNomFich.Text) > 70 then
        begin
          PGIError (#13 + 'Le chemin "' + cmeNomFich.Text + '" comporte plus de 70 caractères.' + #13 +
           #13 + 'Déplacez le fichier dans un répertoire plus court.', 'Echanges');
          if gBALouGLJouAUT = 'BAL' then
            SetParamSoc ('SO_CPEDIFICSOURCE', '')
          else if gBALouGLJouAUT = 'GLJ' then
            SetParamSoc ('SO_CPEDIFICSOURCEG', '')
          else if gBALouGLJouAUT = 'AUT' then
            SetParamSoc ('SO_CPEDIFICSOURCEA', '');
        end
        else
        begin
          if gBALouGLJouAUT = 'BAL' then
          begin
            SetParamSoc ('SO_CPEDIFICSOURCE', cmeNomFich.Text);
          end
          else if gBALouGLJouAUT = 'GLJ' then
          begin
            SetParamSoc ('SO_CPEDIFICSOURCEG', cmeNomFich.Text);
          end
          else if gBALouGLJouAUT = 'AUT' then
          begin
            SetParamSoc ('SO_CPEDIFICSOURCEA', cmeNomFich.Text);
          end;
        end;
      end;
    end;
  end;
end;

procedure TablesModif (pFamilleTable, pScript, pFicDonnees, pFicTable: String);
begin
  if FileExists (pFicDonnees) then
  begin
    if pos ('O', pScript) <> 1 then
    begin
      CreeFicIniPourCISX ('TableConversion', pScript, pFicDonnees, gDirEchDuDos +
       '\' + pFicTable, nil, 'CORRESP', pFamilleTable);
      LanceCorrespCisx ('/INI='+ gDirEchDuDos + ConstFicIniDemande);
    end
    else
    begin
      ConvertCISX (pScript, pFicDonnees, nil, True);
      CreeFicIniPourCISX ('TableConversion', pScript, gDirEchDuDos + '\TraTmp.TRA',
       gDirEchDuDos + '\' + pFicTable, nil, 'CORRESP', pFamilleTable);
      LanceCorrespCisx('/INI='+ gDirEchDuDos + ConstFicIniDemande);
    end;
  end
  else
    PGIError (#13 + 'Le fichier de données' + #13 + pFicDonnees + #13 +
     'n''existe pas.' + #13 +'Modification de la table impossible.', 'Echanges');
end;

procedure VisuModeOperatoire (pNomFicModop: String);
var
  FilePath: string;
  CodeRetour: Integer;
begin
  FilePath := '';
  Delete (pNomFicModop, 1, 1);
  CodeRetour := AGL_YFILESTD_EXTRACT (FilePath, 'COMPTA', pNomFicModop + '.PDF', 'ECH', 'MOP', '', '', '', FALSE, V_PGI.LanguePrinc, 'CEG' );
  if Coderetour <> -1 then
    CodeRetour := AGL_YFILESTD_EXTRACT (FilePath, 'COMPTA', pNomFicModop + '.PDF', 'ECH', 'MOP', '', '', '', FALSE, V_PGI.LanguePrinc, 'DOS' );

  if CodeRetour = -1 then
    ShellExecute (Application.Handle, 'Open', PChar('"' + FilePath + '"'), nil,
    PChar(FilePath), SW_SHOWNORMAL)
  else
  begin
    PGIError (#13 + 'Le mode opératoire ' + pNomFicModop + ' n''existe pas.' +
     #13 + '"' + AGL_YFILESTD_GET_ERR (CodeRetour) + '"' + #13 +
     'Visualisation impossible.', 'Echanges');
  end;
end;

procedure VisuFicDonnees (NomCompletFicDonnees: String);
begin
  if FileExists (NomCompletFicDonnees) then
    ShellExecute (Application.Handle, 'Open',  'WordPad.exe' , PChar('"' +
     NomCompletFicDonnees +'"'), PChar(NomCompletFicDonnees), SW_SHOWNORMAL)
  else
    PGIError (#13 + 'Le fichier de données' + #13 + NomCompletFicDonnees + #13 +
     'n''existe pas.' + #13 +'Visualisation impossible.', 'Echanges');
end;

function VariablesVerification (pHgVars: THGrid): Boolean;
var
  i: Integer;
begin
  Result := True;

  i := 2;
  while ((Result) AND (i <= pHgVars.RowCount)) do
  begin
    if pHgVars.Cells[0, i - 1] <> '' then
    begin
      if Trim (pHgVars.Cells[1, i - 1]) = '' then
      begin
        RestoreCurseur;
        PGIError (#13 + 'La variable ' + pHgVars.Cells[0, i - 1] +
         ' n''est pas renseignée .' + #13 + #13 + 'Traitement abandonné.', 'Echanges');
        pHgVars.col := 1;
        pHgVars.Row := i - 1;
        Result := False;
      end;
    end;
    inc (i);
  end;
end;

procedure TraitePopUp (Sender: TObject; pScript, pFicDonnees: String);
var
  NomMenu: String;
begin
  NomMenu := (Sender as TMenuItem).Name;
  if NomMenu = 'mConvJnx' then
    TablesModif ('JOURNAUX', pScript, pFicDonnees, 'ConvJournaux.txt')
  else if NomMenu = 'mConvCptes' then
    TablesModif ('COMPTES', pScript, pFicDonnees, 'ConvComptes.txt')
  ;
end;

procedure EchgImp_LanceAssistant (pLequel: String);
begin
  CurseurSablier;
  fEchgImp_Assistant := TfEchgImp_Assistant.Create (Application);
  gBALouGLJouAUT := pLequel;
  gPP := FindInsidePanel;
  if gPP = nil then
  begin
    try
      fEchgImp_Assistant.ShowModal;
    finally
      fEchgImp_Assistant.Free;
    end;
  end
  else
  begin
    InitInside (fEchgImp_Assistant, gPP);
    fEchgImp_Assistant.Show;
  end;
  RestoreCurseur;
end;

procedure CompteRenduComSuppression;
begin
  if FileExists (CompteRenduNomDuFichierApres + '.ERR') then
    SysUtils.DeleteFile (CompteRenduNomDuFichierApres + '.ERR');
  if FileExists (CompteRenduNomDuFichierApres + '.OK') then
    SysUtils.DeleteFile (CompteRenduNomDuFichierApres + '.OK');
end;

procedure CompteRenduComAffichage;
var
  fEchgImp_CompteRendu: TfEchgImp_CompteRendu;
  FileName: String;
begin
  fEchgImp_CompteRendu := TfEchgImp_CompteRendu.Create (fEchgImp_Assistant);
  fEchgImp_CompteRendu.ColorMemo1.Clear;

  FileName := CompteRenduNomDuFichierAvant;
  FileName := ExtractFilePath (FileName) + '_' + ExtractFileName (FileName);

  if FileExists (FileName + '.OK') then
    FileName := FileName + '.OK'
  else
    FileName := FileName + '.ERR';

  if FileExists (FileName) then
    fEchgImp_CompteRendu.ColorMemo1.Lines.LoadFromFile (FileName)
  else
    fEchgImp_CompteRendu.ColorMemo1.Text :=
      #13 + #10 +  #13 + #10 + #13 + #10 + '***** RAPPORT INTROUVABLE ! *****';

  fEchgImp_CompteRendu.ShowModal;
  fEchgImp_CompteRendu.Free;
end;

procedure CompteRenduComSauvegarde (SaveDialog: TSaveDialog; ColorMemo: TColorMemo);
begin
  SaveDialog.Options := SaveDialog.Options + [ofOverwritePrompt];
  //CM := TColorMemo.Create (SaveDialog.Owner);
  //CM.Lines.Add ('Rapport importation dossier ' + V_PGI.NoDossier);
  SaveDialog.Title := 'Enregistrement du rapport';
  SaveDialog.FileName := 'Rapport importation ' + V_PGI.NoDossier;

  if SaveDialog.Execute then
    ColorMemo.Lines.SaveToFile (SaveDialog.FileName);
    //CM.Lines.SaveToFile (SaveDialog.FileName);
  //CM.Free;
end;

function ParametresObligatoiresVerif (Muet: Boolean): Boolean;
var
  S: String;
begin
  Result := True;
  with fEchgImp_Parametres do
  begin
    if Result then  //---------- Date import
      if cmeDate.Enabled then
        if Trim (cmeDate.Text) = '' then
        begin
          Result := False;
          S := 'La date par défaut n''est pas renseignée.' + S;
          Focuzer := cmeDate;
        end;
    if Result then  //---------- Libellé défaut
      if HValCB_Libelle.Enabled then
        if Trim (HValCB_Libelle.Text) = '' then
        begin
          Result := False;
          S := 'Le libellé par défaut n''est pas renseigné.' + S;
          Focuzer := HValCB_Libelle;
        end;
    if Result then  //---------- Journal défaut
      if Hcme_Journal.Enabled then
      begin
        if Trim (Hcme_Journal.Text) = '' then
        begin
          Result := False;
          S := 'Le journal par défaut n''est pas renseigné.' + S;
        end;
        if not ExisteSQL ('SELECT j_journal FROM journal WHERE j_journal="' +
         Hcme_Journal.Text +'"') then
        begin
          S := 'Le journal par défaut ' + Hcme_Journal.Text + ' n''existe pas.' + S;
          Result := False;
        end;
        if not Result then Focuzer := Hcme_Journal;
      end;

    if Result then  //---------- Compte défaut
      if Hcme_Compte.Enabled then
      begin
        if Trim (Hcme_Compte.Text) = '' then
        begin
          Result := False;
          S := 'Le compte par défaut n''est pas renseigné.' + S;
        end;
        if not ExisteSQL ('SELECT g_general FROM generaux WHERE g_general="' +
         Trim (Hcme_Compte.Text) +'"') then
        begin
          Result := False;
          S := 'Le compte par défaut ' + Hcme_Compte.Text + ' n''existe pas.' + S;
        end;
        if not Result then Focuzer := Hcme_Compte;
      end;

    if Result then  //---------- Etablissement
      if Hcme_Etablissement.Enabled then
      begin
        if ((Trim (Hcme_Etablissement.Text) = '') or
             (not ExisteSQL ('select * from ETABLISS where ET_ETABLISSEMENT="' +
               Hcme_Etablissement.Text +'"'))
        ) then
        begin
          Result := False;
          S := 'L''établissement n''est pas ou mal renseigné.' + S;
          Focuzer := Hcme_Etablissement;
        end;
      end;

    if not Result then
    begin
      PGIError (#13 + S, 'Echanges');
      if Focuzer <> nil then
        // CA - 28/09/2005 - Ajout CanFocus
        try if Focuzer.CanFocus then Focuzer.SetFocus; except; end;  // pour les cas ou la fiche est inactive
    end;
  end;
end;

function ParametresEnregistre: Boolean;
begin
  Result := False;
  if NOT ParametresObligatoiresVerif (False) then EXIT;
  with fEchgImp_Parametres do
  begin
    if gBALouGLJouAUT = 'BAL' then
    begin
      SetParamSoc ('SO_CPEDIEFFACEMENT', cbRAZPrealable.Checked);     // RAZ

      //SetParamSoc ('SO_CPEDIDATEVAL', StrtoDate (cmeDate.Text));    // valeur Date par défaut

      SetParamSoc ('SO_CPEDILIBELLE', HValCB_Libelle.Text);           // Libellé défaut
      SetParamSoc ('SO_CPEDILIBELLEDATE', CB_AvecDateImport.Checked); // Avec date import
      SetParamSoc ('SO_CPEDIJOURNAL', Hcme_Journal.Text);             // Journal défaut
      SetParamSoc ('SO_CPEDICOMPTE', Hcme_Compte.Text);               // Compte défaut
      SetParamSoc ('SO_CPEDIETABLISSEMENT', Hcme_Etablissement.Text); // Etablissement

      SetParamSoc ('SO_CPEDIETABLISSEMENTDEF', Hcme_Etablissement.Text);
    end
    else if gBALouGLJouAUT = 'GLJ' then
    begin
      SetParamSoc ('SO_CPEDIEFFACEMENTG', cbRAZPrealable.Checked);     // RAZ

      SetParamSoc ('SO_CPEDIDATEVALG', StrToDate (cmeDate.Text));      // valeur Date par défaut

      SetParamSoc ('SO_CPEDILIBELLEG', HValCB_Libelle.Text);           // Libellé défaut
      SetParamSoc ('SO_CPEDILIBELLEDATEG', CB_AvecDateImport.Checked); // Avec date import
      SetParamSoc ('SO_CPEDIJOURNALG', Hcme_Journal.Text);             // Journal défaut
      SetParamSoc ('SO_CPEDICOMPTEG', Hcme_Compte.Text);               // Compte défaut
      SetParamSoc ('SO_CPEDIETABLISSEMENTG', Hcme_Etablissement.Text); // Etablissement

      SetParamSoc ('SO_CPEDIETABLISSEMENTDEF', Hcme_Etablissement.Text);
    end
    else if gBALouGLJouAUT = 'AUT' then
    begin
      SetParamSoc ('SO_CPEDIDATEVALA', StrToDate (cmeDate.Text));      // valeur Date par défaut

      SetParamSoc ('SO_CPEDILIBELLEA', HValCB_Libelle.Text);           // Libellé défaut
      SetParamSoc ('SO_CPEDILIBELLEDATEA', CB_AvecDateImport.Checked); // Avec date import
      SetParamSoc ('SO_CPEDIJOURNALA', Hcme_Journal.Text);             // Journal défaut
      SetParamSoc ('SO_CPEDICOMPTEA', Hcme_Compte.Text);               // Compte défaut
      SetParamSoc ('SO_CPEDIETABLISSEMENTA', Hcme_Etablissement.Text); // Etablissement

      SetParamSoc ('SO_CPEDIETABLISSEMENTDEF', Hcme_Etablissement.Text);
    end;
  end;

{$IFDEF EAGLCLIENT}
  AvertirCacheServer ('PARAMSOC') ;
{$ENDIF}

  Result := True;
end;

procedure ParametresCharge (BALouGLJouAUT: String);
var
  S: String;
begin
  with fEchgImp_Parametres do
  begin
    if BALouGLJouAUT = 'BAL' then
    begin
      Caption := 'Paramètres Balances';
      cbRAZPrealable.Checked := GetParamSocSecur ('SO_CPEDIEFFACEMENT', False);

      cmeDate.Text := fEchgImp_Assistant.cmeDate.Text;
      labDate.Enabled := False;
      cmeDate.Enabled := False;

      HValCB_Libelle.Text := GetParamSocSecur ('SO_CPEDILIBELLE', '');
      CB_AvecDateImport.Checked := Boolean (GetParamSocSecur ('SO_CPEDILIBELLEDATE', False));
      Hcme_Journal.Text := GetParamSocSecur ('SO_CPEDIJOURNAL', '');
      Hcme_Compte.Text := GetParamSocSecur ('SO_CPEDICOMPTE', '');

      //S := GetParamSocSecur ('SO_CPEDIETABLISSEMENT', VH^.EtablisDefaut);
      S := GetParamSocSecur ('SO_CPEDIETABLISSEMENT', GetParamSocSecur ('SO_ETABLISDEFAUT', ''));



      if S = '' then
        //S := RechDom ('TTETABLISSEMENT', GetParamSocSecur ('SO_ETABLISDEFAUT', ''), False)
        S := GetParamSocSecur ('SO_ETABLISDEFAUT', '');



      Hcme_Etablissement.Text := S
    end
    else if BALouGLJouAUT = 'GLJ' then
    begin
      Caption := 'Paramètres Grands-livres - Journaux';
      cbRAZPrealable.Checked := GetParamSocSecur ('SO_CPEDIEFFACEMENTG', False);

      if IsValidDate (GetParamSocSecur ('SO_CPEDIDATEVALG', idate1900)) then
        cmeDate.Text := GetParamSocSecur ('SO_CPEDIDATEVALG', dateToStr(now))
      else
        cmeDate.Text := dateToStr(now);

      HValCB_Libelle.Text := GetParamSocSecur ('SO_CPEDILIBELLEG', '');
      CB_AvecDateImport.Checked := Boolean (GetParamSocSecur ('SO_CPEDILIBELLEDATEG', False));
      Hcme_Journal.Text := GetParamSocSecur ('SO_CPEDIJOURNALG', '');
      Hcme_Compte.Text := GetParamSocSecur ('SO_CPEDICOMPTEG', '');

      S := GetParamSocSecur ('SO_CPEDIETABLISSEMENTG', GetParamSocSecur ('SO_ETABLISDEFAUT', ''));
      if S = '' then
        S := GetParamSocSecur ('SO_ETABLISDEFAUT', '');
      Hcme_Etablissement.Text := S
    end
    else if BALouGLJouAUT = 'AUT' then
    begin
      Caption := 'Importation automatique';
      cbRAZPrealable.Enabled := False;
      if IsValidDate (GetParamSocSecur ('SO_CPEDIDATEVALA', idate1900)) then
        cmeDate.Text := GetParamSocSecur ('SO_CPEDIDATEVALA', dateToStr(now))
      else
        cmeDate.Text := dateToStr(now);

      HValCB_Libelle.Text := GetParamSocSecur ('SO_CPEDILIBELLEA', '');
      CB_AvecDateImport.Checked := Boolean (GetParamSocSecur ('SO_CPEDILIBELLEDATEA', False));
      Hcme_Journal.Text := GetParamSocSecur ('SO_CPEDIJOURNALA', '');
      Hcme_Compte.Text := GetParamSocSecur ('SO_CPEDICOMPTEA', '');

      S := GetParamSocSecur ('SO_CPEDIETABLISSEMENTA', GetParamSocSecur ('SO_ETABLISDEFAUT', ''));
      if S = '' then
        S := GetParamSocSecur ('SO_ETABLISDEFAUT', '');
      Hcme_Etablissement.Text := S
    end
    else
    begin
      // Pour ne pas oublier de traiter si on ajoute des extensions dans le futur
      PGIInfo ('ParametresCharge' + #13 + 'Attendu BAL ou GLJ' + #13 +
       'Reçu ' + BALouGLJouAUT);
      EXIT;
    end;
  end;
end;

procedure ParametresAffichage;
begin
  ParametresCharge (gBALouGLJouAUT);
  fEchgImp_Parametres.ShowModal;
end;

procedure ImportManuel (pScript, pFicDonnees: String; pHgVars: THGrid);
var
  S: String;
  Ret: String;
begin
  CurseurSablier;
  CompteRenduComSuppression;

  if not VariablesVerification (pHgVars) then
  begin
    RestoreCurseur;
    EXIT;
  end;

  ParametresCharge (gBALouGLJouAUT);
  if not ParametresObligatoiresVerif (False) then
  begin
    RestoreCurseur;
    ParametresAffichage;
    EXIT;
  end;

  if (((gBALouGLJouAUT = 'BAL') and (GetParamSocSecur ('SO_CPEDIEFFACEMENT', False) = True))
  or  ((gBALouGLJouAUT = 'GLJ') and (GetParamSocSecur ('SO_CPEDIEFFACEMENTG', False) = True)))
  then begin
    S := 'JRL';
    Ret := AglLanceFiche ('CP','IMPORTCOMCONF','', '',
    'Origine du fichier : Echanges PGI; Nature : ' +  S + '; Fichier : ' + pFicDonnees
    + ';Appel : ECHANGES PGI'
    );
  end;

  if Ret = '?' then
  begin
    PGIError ('Incohérence : RAZ exercice non cochée.' + #13 +
     'En contradiction avec le paramètre effacement préalable.' + #13 + #13 +
     'Conseil : Décochez "Effacement préalable" ou cochez "RAZ exercice".',
     'Echanges');
    RestoreCurseur;
    EXIT;
  end
  else if Ret = '-' then
  begin
    PGIError ('Traitement abandonné.', 'Echanges');
    RestoreCurseur;
    EXIT;
  end;

  if ConvertCISX (pScript, pFicDonnees, pHgVars) then
    LanceComSx;

  if FileExists (CompteRenduNomDuFichierApres + '.ERR') then
    CompteRenduComAffichage;

  RestoreCurseur;
end;

function ParametresSiModifies: Boolean;
begin
  Result := False;
  with fEchgImp_Parametres do
  begin
    if gBALouGLJouAUT = 'BAL' then
    begin
      if cbRAZPrealable.Checked <> GetParamSocSecur ('SO_CPEDIEFFACEMENT', False) then
        Result := True else

      if HValCB_Libelle.Text <> GetParamSocSecur ('SO_CPEDILIBELLE', '') then
        Result := True else
      if CB_AvecDateImport.Checked <> Boolean (GetParamSocSecur ('SO_CPEDILIBELLEDATE', False)) then
        Result := True else
      if Hcme_Journal.Text <> GetParamSocSecur ('SO_CPEDIJOURNAL', '') then
        Result := True else
      if Hcme_Compte.Text <> GetParamSocSecur ('SO_CPEDICOMPTE', '') then
        Result := True else
      if Hcme_Etablissement.Text <> GetParamSocSecur ('SO_CPEDIETABLISSEMENT', '') then
        Result := True
      ;
    end
    else if gBALouGLJouAUT = 'GLJ' then
    begin
      if cbRAZPrealable.Checked <> GetParamSocSecur ('SO_CPEDIEFFACEMENTG', False) then
        Result := True else

      //if ((HValCB_Date.Value <> 'LIB') and (cmeDate.Text <> GetParamSocSecur ('SO_CPEDIDATEVALG', dateToStr(now)))) then
      //  Result := True else

      if HValCB_Libelle.Text <> GetParamSocSecur ('SO_CPEDILIBELLEG', '') then
        Result := True else
      if CB_AvecDateImport.Checked <> Boolean (GetParamSocSecur ('SO_CPEDILIBELLEDATEG', False)) then
        Result := True else
      if Hcme_Journal.Text <> GetParamSocSecur ('SO_CPEDIJOURNALG', '') then
        Result := True else
      if Hcme_Compte.Text <> GetParamSocSecur ('SO_CPEDICOMPTEG', '') then
        Result := True else
      if Hcme_Etablissement.Text <> GetParamSocSecur ('SO_CPEDIETABLISSEMENTG', '') then
        Result := True
      ;
    end
    else if gBALouGLJouAUT = 'AUT' then
    begin
      //if ((HValCB_Date.Value <> 'LIB') and (cmeDate.Text <> GetParamSocSecur ('SO_CPEDIDATEVALA', dateToStr(now)))) then
      //  Result := True else

      if HValCB_Libelle.Text <> GetParamSocSecur ('SO_CPEDILIBELLEA', '') then
        Result := True else
      if CB_AvecDateImport.Checked <> Boolean (GetParamSocSecur ('SO_CPEDILIBELLEDATEA', False)) then
        Result := True else
      if Hcme_Journal.Text <> GetParamSocSecur ('SO_CPEDIJOURNALA', '') then
        Result := True else
      if Hcme_Compte.Text <> GetParamSocSecur ('SO_CPEDICOMPTEA', '') then
        Result := True else
      if Hcme_Etablissement.Text <> GetParamSocSecur ('SO_CPEDIETABLISSEMENTA', '') then
        Result := True
      ;
    end;
  end;
end;

procedure LookUpA (pCme: THCritMaskEdit);
var
  TOBi: TOB;
begin
  //TOBi := LookUpTob (nil, gTOBScripts, 'Origine',
  TOBi := LookUpTob (pCme, gTOBScripts, 'Origine',
   'CIS_TABLE1;CIS_COMMENT;CIS_TABLE2', 'Editeur;Libellé;Nature');

  if Assigned (TOBi) then
  begin
    pCme.Plus := TOBi.GetValue('DomScript');
    pCme.Text := TOBi.GetValue('CIS_COMMENT');
  end;      
  //TOBi.free;  // ajout me
end;

function PamametresAnnulation: Boolean;
var
  Ret: Integer;
begin
  Result := True;
  if not ParametresSiModifies then
    EXIT;

  Ret := Application.MessageBox ('Les modifications des paramètres ne sont' +
   ' pas validées' + #13 + #13 + 'Voulez-vous les valider ?', 'Echanges',
   MB_YESNOCANCEL + MB_DEFBUTTON3);

  if Ret = IDCANCEL then
    Result := False
  else if Ret = IDYES then
    fEchgImp_Parametres.BValiderClick (nil);
end;

procedure BoutonValiderClick;
begin
  CurseurSablier;
  fEchgImp_Assistant.BValider.Enabled := False;
  fEchgImp_Assistant.BValider.Refresh;
  fEchgImp_Assistant.btnTables.Enabled := False;
  fEchgImp_Assistant.btnTables.Refresh;

/////////////////////
  if gBALouGLJouAUT = 'BAL' then
    SetParamSoc ('SO_CPEDIDATEVAL', StrToDate (fEchgImp_Assistant.cmeDate.Text));   // valeur Date par défaut
/////////////////////

  if gBALouGLJouAUT = 'AUT' then
  begin
    if VariablesVerification (fEchgImp_Assistant.hgVariables) then
    begin
      ParametresCharge (gBALouGLJouAUT);
      if not ParametresObligatoiresVerif (False) then
      begin
        RestoreCurseur;
        ParametresAffichage;
        //EXIT;
      end;
    end;
  end
  else
    ImportManuel (fEchgImp_Assistant.Hcme_Solutions.Plus,
     fEchgImp_Assistant.cmeNomFich.Text, fEchgImp_Assistant.hgVariables);

  fEchgImp_Assistant.btnTables.Enabled := True;
  fEchgImp_Assistant.BValider.Enabled := True;
  RestoreCurseur;
end;

procedure RESTOR_ScriptCISX (pFicScript: String);
var
  FLoc: TextFile;
  FicIni, FicBobOut: String;
begin
  FicINI := gDirEchDuDos + ConstFicIniDemande;
  SysUtils.DeleteFile (FicIni);

  FicBobOut := gDirEchDuDos + ConstFicIniRetour;
  // Effacement du fichier de sortie pour être sûr de ne pas utiliser retours précédents
  SysUtils.DeleteFile (FicBobOut);

  AssignFile (FLoc, FicIni);
  Rewrite (FLoc);

  WriteLN (FLoc, '[PRODUIT]');
  WriteLN (FLoc, 'APPELANT=ECHANGES PGI');
  WriteLN (FLoc, 'DOMAINE=E');

// ME deb
  //WriteLN (FLoc, 'REPSTD=' + ChangeStdDatPath ('$STD\CISX'));
  WriteLN (FLoc, 'MODE=RESTOR');
  if (V_PGI.ModePCL='1') then
     WriteLN (FLoc, 'REPSTD=' + ChangeStdDatPath ('$STD\CISX'))
  else
     WriteLN (FLoc, 'REPSTD=' + ExtractFilePath (Application.ExeName)+'parametre');
// ME fin

  if V_PGI.DosPath[length (V_PGI.DosPath)] <> '\' then WriteLN (FLoc, 'REPDOS=' + V_PGI.DosPath + '\CISX')
  else  WriteLN (FLoc, 'REPDOS=' + V_PGI.DosPath + 'CISX');


  WriteLN (FLoc, #13 + #10);
  WriteLN (FLoc, '[COMMANDE]');

  {Repertoire := ExtractFilePath (pFicScripts);
  i := Length (Repertoire);
  if Repertoire[i] = '\' then
    System.Delete (Repertoire, i, 1);
  WriteLN (FLoc, 'REPERTOIRE=' + Repertoire);}
  WriteLN (FLoc, 'REPERTOIRE=' + gDirEchDuDos);

  WriteLN (FLoc, String ('NOMFICHIER=' + ExtractFileName (ConstFicIniRetour)));
  WriteLN (FLoc, String ('LISTEFICHIER=' + pFicScript));

  CloseFile (FLoc);
  LanceRestoreCisx('/INI='+FicIni); // ajout me
end;

function IntegreBobMop (pFich: String): Integer;
var
  msg, TitreAlerte: String;
begin
  TitreAlerte := 'Intégration mode opératoire';
  Result := AglIntegreBob (pFich, FALSE, FALSE);
  case Result of
    0: begin
    //   msg := 'Intégration de : ';      // OK
    //   Pgiinfo (msg + sFileImport, TitreAlerte);
    end;
    // 1: begin
    //   msg := 'Intégration déjà effectuée. Fichier ';
    //   Pgiinfo(msg + sFileImport, TitreAlerte);
    // end;
    -1: begin
      msg := 'Erreur d''écriture dans la table YMYBOBS. Fichier ';
      // if V_PGI.SAV then
      PGIInfo (msg + pFich, TitreAlerte);
    end;
    -2: begin
      msg := 'Erreur d''intégration dans la fonction AglImportBob. Fichier ';
      PGIInfo (msg + pFich, TitreAlerte);
    end;
    -3: begin
      msg := 'Erreur de lecture du fichier BOB. Fichier ';
      PGIInfo (msg + pFich, TitreAlerte);
    end;
   // -4: begin     // Erreur inconnue.
    else
    begin
      msg := 'Erreur indéfinie (' + intToStr(Result) + '). Fichier ';
      PGIInfo (msg + pFich, TitreAlerte);
    end;
  end;
end;

procedure Kit;
var
  locOpenDial     : TOpenDialog;
  FNom, Predefini : string;
  msg, TitreAlerte: String;

  function ImportPDF (pFich: String): Integer;
  begin
    if not FileExists (pFich) then
    begin
      TitreAlerte := 'Intégration mode opératoire';
      msg := 'Fichier non trouvé : ';
      PGIInfo (msg + pFich, TitreAlerte);
      Result := -1;
      EXIT;
    end;
    if V_PGI.PassWord = CryptageSt (DayPass (Date)) then Predefini := 'CEG'
    else Predefini := 'DOS';
{    Result := AGL_YFILESTD_IMPORT (pFich, 'COMPTA', ExtractFileName (pFich), 'PDF',
      V_PGI.LanguePrinc, 'ECH', 'MOP', '', '', '', '-', '-', '-', '-', '-',
      Predefini, 'Script Cisx', V_PGI.NoDossier);}

    Result := AGL_YFILESTD_IMPORT (pFich, 'COMPTA', ExtractFileName (pFich), 'PDF',
      'ECH', 'MOP', '', '', '', '-', '-', '-', '-', '-',
      V_PGI.LanguePrinc, Predefini, 'Script Cisx', V_PGI.NoDossier);

    if Result = -1 then Result := 1;
  end;

begin
  locOpenDial := TOpenDialog.Create (nil);
  locOpenDial.Filter :=
   'Paramétrages (*.cix)|*.CIX|Modes opératoires (*.pdf)|*.pdf|Bobs (*.bob)|*.bob';
  locOpenDial.Execute;
  FNom := locOpenDial.FileName;
  locOpenDial.Free;
  if FNom = '' then EXIT;

  if LowerCase (ExtractFileExt (FNom)) = '.pdf' then
  begin
    ImportPDF (FNom);
    if FileExists (ChangeFileExt (FNom, '.cix')) then
      RESTOR_ScriptCISX (ChangeFileExt (FNom, '.cix'));
  end
  else if LowerCase (ExtractFileExt (FNom)) = '.cix' then
  begin
    RESTOR_ScriptCISX (FNom);
    if FileExists (ChangeFileExt (FNom, '.pdf')) then
      ImportPDF (ChangeFileExt (FNom, '.pdf'));
  end
  else if LowerCase (ExtractFileExt (FNom)) = '.bob' then
  begin
    IntegreBobMop (FNom);
  end;
end;

function ImportAutomatique2 (Muet: Boolean): integer;
// Valeurs de retour
//    0 : Pas d'importation automatique dans ce dossier
//    1 : Importation automatique REUSSIE
//   -1 : PAS DE FICHIER à importer
//   -1 : PAS DE FICHIER à importer
var
  NomFich: String;
  Rap: WideString;
  i: integer;
begin
  Result := 0;

  NomFich := CorbeilleActivation;
  if Trim (NomFich) = '' then
    EXIT;


(* CA - 05/07/2007 - Remplacement DosPath car n'existe pas en EAGL

  if V_PGI.DosPath[length (V_PGI.DosPath)] <> '\' then
    gDirEchDuDos := V_PGI.DosPath + '\Echanges'
  else gDirEchDuDos := V_PGI.DosPath + 'Echanges';

*)

  gDirEchDuDos := TcbpPath.GetCegidUserTempPath + 'Echanges';

  if Not DirectoryExists (gDirEchDuDos) then
    CreateDir (gDirEchDuDos);

  gFicTRA := gDirEchDuDos + '\PG' + V_PGI.CodeSociete + '.TRA';
  CompteRenduComSuppression;

  IF NOT Muet THEN InitMoveProgressForm (nil, 'Récupération automatique',
   'Recupération de ' + NomFich + '. Veuillez patienter.', 6, False, True);
  IF NOT Muet THEN MoveCurProgressForm;

  i := CopieTableAccess (True);
  if i <> 1 then
  begin
    Rap := 'Erreur (' + inttostr (i) + ') à la création de la table Access';
    IF NOT Muet THEN PGIError (#13 + Rap + #13 + #13 + 'Traitement abandonné.',
     'Echanges');
    IF NOT Muet THEN FiniMoveProgressForm;
    EXIT;
  end;
  IF NOT Muet THEN MoveCurProgressForm;

  CreeFicTableSiInexist ('ConvComptes.txt');
  CreeFicTableSiInexist ('ConvJournaux.txt');
  IF NOT Muet THEN MoveCurProgressForm;

  fEchgImp_Parametres := TfEchgImp_Parametres.Create (nil);
  IF NOT Muet THEN MoveCurProgressForm;

//////////////////////////
  //if VariablesVerification (fEchgImp_Assistant.hgVariables) then

  // FQ 21831 - CA - 14/11/2007 - Pour initialiser correctement les paramètres à charger.
  ParametresDisable;
  // FQ 21831 - Fin
  ParametresCharge ('AUT');
  if not ParametresObligatoiresVerif (True) then
  begin
    Rap := 'Paramètres erronés';
    IF NOT Muet THEN PGIError (#13 + Rap + #13 + #13 + 'Traitement abandonné.',
     'Echanges');
    IF NOT Muet THEN FiniMoveProgressForm;
    EXIT;
  end;

  //GetParamSocSecur ('SO_CPEDIORIGINELIB', '');

  // on ferme le ProgressForm pour laisser la place à celui de cisx
  IF NOT Muet THEN FiniMoveProgressForm;

//ne traite pas variables
  //if ConvertCISX (GetParamSocSecur ('SO_CPEDIORIGINE', ''), NomFich, pHgVars) then
  if ConvertCISX (GetParamSocSecur ('SO_CPEDIORIGINEA', ''), NomFich, nil) then
    LanceComSx;

  IF NOT Muet AND FileExists (CompteRenduNomDuFichierApres + '.ERR') then
    CompteRenduComAffichage;

  if not FileExists (CompteRenduNomDuFichierApres + '.ERR') then
    SysUtils.RenameFile (NomFich, NomFich + '.OK');

  RestoreCurseur;
end;

procedure ImportAutomatique1;
var
  Save_gBALouGLJouAUT: String;
begin
  Save_gBALouGLJouAUT := gBALouGLJouAUT;
  gBALouGLJouAUT := 'AUT';
  ImportAutomatique2 (FALSE);
  gBALouGLJouAUT := Save_gBALouGLJouAUT;
end;


end.

