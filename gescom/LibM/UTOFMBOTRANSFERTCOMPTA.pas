{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 19/03/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : TRANSFERTCOMPTA ()
Mots clefs ... : TOF;TRANSFERTCOMPTA
*****************************************************************}
unit UTOFMBOTRANSFERTCOMPTA;

interface

uses StdCtrls, Controls, Classes, Vierge, InterCompta, Windows,
  {$IFNDEF EAGLCLIENT}
  db, dbtables,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF
  , MBOExportCompta, FileCtrl, ed_tools, HTB97, ParamSoc, EntGC, UTob {,VoirTob},
  UtilUtilitaires;

type
  TOF_TRANSFERTCOMPTA = class(TOF)
  private
    RapportFinTrt, RegDetDesactive, RgpSurAuxiCentral, ComptaExterne: boolean;
    RapportLst: TListBox;
    RegEnregistre, RgpAuxiCentral, RgpDetEnt, RgpDetPart, RgpRgtTiers, RgpRgtBanque: TCheckBox;
    TobRegrParPiece, TobPieceEqui: TOB;
    GenerationFileOk: Boolean;
    FichierEcr, DateFile: string;
    procedure SetLastError(Num: integer; ou: string);
    function CountNature(Natures: string): integer;
    procedure MiseAJourPiece(Nature, DepotD, DateDeb, DateFin: string);
    procedure NomDuFichierExport;
    procedure MajRapport;
    procedure GenereFile;
    procedure ImportFile;
    function LanceTraitement(LigneDeCommande: string): boolean;
    //procedure IntegreFile ;
    procedure ClickImprimer(Sender: TObject);
    procedure ClickRegRgtBanque(Sender: Tobject);
    procedure ClickRegRgtTiers(Sender: Tobject);

  public
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  end;

const
  // libellés des messages
  TexteMessage: array[1..15] of string = (
    {1}'Aucune pièces trouvées pour la période sélectionnée.'
    {2}, 'Aucune pièce ne correspond aux critères sélectionnés.'
    {3}, 'Tiers inconnu.'
    {4}, 'Impossible de trouver le compte HT de l''article.'
    {5}, 'Impossible de trouver le compte de taxe.'
    {6}, 'Impossible de trouver le compte de remise.'
    {7}, 'Impossible de trouver le compte d''escompte.'
    {8}, 'Impossible de trouver le compte de règlement.'
    {9}, 'Impossible de trouver le compte de ports et Frais.'
    {10}, 'Traitement annulé par l''utilisateur'
    {11}, 'Erreur lors de la génération des lignes.'
    {12}, 'Veuillez confirmer la comptabilisation des écritures déjà extraites.'
    {13}, 'L''emplacement du fichier d''export n''est pas défini dans le paramètrage.'
    {14}, 'Le nom de fichier est incorrect.'
    {15}, 'Le fichier ne s''est pas généré correctement, vous devez relancer le traitement.'
    );

implementation

procedure TOF_TRANSFERTCOMPTA.SetLastError(Num: integer; ou: string);
begin
  if Num = 0 then exit;
  if ou <> '' then SetFocusControl(ou);
  LastError := -1;
  PGIError(TraduireMemoire(TexteMessage[Num]), Ecran.Caption);
end;

procedure TOF_TRANSFERTCOMPTA.OnNew;
begin
  inherited;
end;

procedure TOF_TRANSFERTCOMPTA.OnDelete;
begin
  inherited;
end;

function TOF_TRANSFERTCOMPTA.CountNature(Natures: string): integer;
var St, Nat: string;
  i: integer;
begin
  i := 0;
  Nat := Natures;
  repeat
    St := ReadTokenSt(Nat);
    if St <> '' then Inc(i);
  until Nat = '';
  result := i;
end;

procedure TOF_TRANSFERTCOMPTA.OnUpdate;
var NatureD, DepotD, nature, DateDeb, DateFin, EtatCompta, Options, PieceDem, Piece: string;
  //    LastError : integer;
  TrtOk, RegDetAutorise, bRapport: Boolean;
  TobTmp, TobTmp1: TOB;
  Qry: TQuery;
begin
  inherited;
  // Récupération des paramètres pour remplir la table INTERCOMPTA
  NatureD := GetControlText('PC_NATURE');
  DepotD := GetControlText('PC_ETABLISSEMENT');
  DateDeb := GetControlText('PC_DATE1');
  DateFin := GetControlText('PC_DATE2');
  RapportLst.Clear;
  if GetParamSoc('SO_MBOCHEMINCOMPTA') = '' then
  begin
    LastError := 13;
    SetLastError(LastError, '');
    exit;
  end;
  if (GetControlText('PC_ECRITURES') = '-') then
  begin
    EtatCompta := '"ATT"'
  end else
  begin
    //Demande msg de conf
    if PGIAsk(TraduireMemoire(TexteMessage[12]), Ecran.Caption) = mrNo then
    begin
      LastError := -1;
      exit;
    end;
    EtatCompta := '"ATT","EXP"';
  end;
  if (NatureD = '') or (DepotD = '') or (DateDeb = '') or (DateFin = '') then
  begin
    LastError := 2;
    SetLastError(LastError, '');
    exit;
  end;
  //Création d'une TOB contenant la nature et la nature équivalente (pour traitement des FFO)
  TobPieceEqui := TOB.Create('PARPIECE', nil, -1);
  Qry := OpenSQL('SELECT GPP_NATUREPIECEG, GPP_EQUIPIECE FROM PARPIECE', True);
  TobPieceEqui.LoadDetailDB('PARPIECE', '', '', Qry, False);
  //Création d'une TOB contenant le type de regroupement du parpiece pour chaque type de pièce
  TobRegrParPiece := TOB.Create('Regr piece', nil, -1);
  //Test si détail demandé mais impossible par rapport au paramètrage du ParPiece
  TobTmp1 := nil;
  RegDetDesactive := False;
  PieceDem := NatureD;
  if (GetControlText('REGR_DETENTREPRISE') = 'X') or (GetControlText('REGR_DETPARTICULIER') = '') then
  begin
    RegDetAutorise := False;
    while Length(PieceDem) > 0 do
    begin
      Piece := ReadTokenSt(PieceDem);
      TobTmp := TobRegrParPiece.FindFirst(['LANATURE'], [Piece], True);
      if TobTmp = nil then
      begin
        TobTmp1 := TOB.Create('', TobRegrParPiece, -1);
        TobTmp1.AddChampSupValeur('LANATURE', Piece, False);
      end;
      if GetInfoParPiece(Piece, 'GPP_REGROUPCPTA') = 'AUC' then
      begin
        if TobTmp = nil then
          TobTmp1.AddChampSupValeur('LEREGROUPEMENT', 'NON', False);
        RegDetAutorise := True;
      end else
      begin
        if TobTmp = nil then
          TobTmp1.AddChampSupValeur('LEREGROUPEMENT', 'OUI', False);
      end;
    end;
    if not RegDetAutorise then
    begin
      RegDetDesactive := True;
      SetControlText('REGR_DETENTREPRISE', '-');
      SetControlText('REGR_DETPARTICULIER', '-');
    end;
  end else
  begin
    while Length(PieceDem) > 0 do
    begin
      Piece := ReadTokenSt(PieceDem);
      TobTmp := TobRegrParPiece.FindFirst(['LANATURE'], [Piece], True);
      if TobTmp = nil then
      begin
        TobTmp1 := TOB.Create('', TobRegrParPiece, -1);
        TobTmp1.AddChampSupValeur('LANATURE', Piece, False);
        if GetInfoParPiece(Piece, 'GPP_REGROUPCPTA') = 'AUC' then
          TobTmp1.AddChampSupValeur('LEREGROUPEMENT', 'NON', False)
        else
          TobTmp1.AddChampSupValeur('LEREGROUPEMENT', 'OUI', False);
      end;
    end;
  end;
  Nature := NatureD;
  // Chargement de la table InterCompta
  InitMoveProgressForm(Ecran, Ecran.Caption, TraduireMemoire('Traitement en cours.'), CountNature(NatureD), False, True);
  if not RecupParametre(LastError, Nature, depotD, Datedeb, datefin, EtatCompta, TobPieceEqui) then
  begin
    SetLastError(LastError, '');
    FiniMoveProgressForm;
    exit;
  end;
  //Calcul du nom du fichier .TRA ou .TRT
  FichierEcr := '';
  DateFile := '';
  NomDuFichierExport;
  //Récup des options
  Options := GetControlText('REGR_AUXICENTRALISATION');
  Options := Options + ';' + GetControlText('REGR_DETENTREPRISE');
  Options := Options + ';' + GetControlText('REGR_DETPARTICULIER');
  Options := Options + ';' + GetControlText('REGR_PAIEMDIFFERENT');
  Options := Options + ';' + GetControlText('REGR_RGTBANQUE');
  Options := Options + ';' + GetControlText('REGR_RGTTIERS');
  if RapportFinTrt then
    MajRApport;
  // Export Compta à partir de la table InterCompta
  TrtOk := ExportCompta(LastError, Ecran, Options, RapportFinTrt, RapportLst, TobRegrParPiece, TobPieceEqui);
  if TrtOk then
  begin
    MiseAJourPiece(Nature, DepotD, DateDeb, DateFin);
    //Génération du fichier .TRA ou .TRT
    GenereFile;
    TToolBarButton97(GetControl('BImprimer')).Visible := RapportFinTrt;
    TTabSheet(GetControl('RAPPORT_')).TabVisible := RapportFinTrt;
  end;
  ExecuteSQL('DELETE FROM INTERCOMPTA WHERE GIC_USER="' + V_PGI.User + '"');
  if not GenerationFileOk then
  begin
    LastError := 15;
    SetLastError(LastError, '');
  end;
  //Pas de compta externe, destruction des écritures en simulation et import
  //des écritures "Normales" testées par ComSX
  if not ComptaExterne then
  begin
    if GenerationFileOk then
      ImportFile;
    ExecuteSQL('DELETE FROM ECRITURE WHERE E_QUALIFPIECE="S" AND E_QUALIFORIGINE="GC"');
    ExecuteSQL('DELETE FROM ANALYTIQ WHERE Y_QUALIFPIECE="S" AND Y_CREERPAR="EXP"');
  end else
    //Compta externe, destruction de toutes les écritures (sert de table temporaire)
  begin
    ExecuteSQL('DELETE FROM ECRITURE');
    ExecuteSQL('DELETE FROM ANALYTIQ');
  end;
  TobRegrParPiece.Free;
  TobPieceEqui.Free;
  FiniMoveProgressForm;
  if GenerationFileOk then
    PGIInfo('Transfert terminé.', Ecran.Caption);
end;

procedure TOF_TRANSFERTCOMPTA.MajRapport;
var Cpt, QteOptions: integer;
  TtesNatPces, NatPce, TousEts, Ets, RgpPce: string;

  procedure AddTexte(Texte: string);
  begin
    QteOptions := QteOptions + 1;
    RapportLst.Items.Add(Texte);
  end;

begin
  RapportLst.Items.Add('***** ' + TraduireMemoire('PARAMETRES ') + '**********************************************************');
  TousEts := GetControlText('PC_ETABLISSEMENT');
  Cpt := 0;
  while TousEts <> '' do
  begin
    Cpt := Cpt + 1;
    Ets := ReadTokenSt(TousEts);
    if Cpt = 1 then
      AddTexte(TraduireMemoire('Etablisement(s) : ') + Ets + ' - ' + RechDom('TTETABLISSEMENT', Ets, False))
    else
      AddTexte('                : ' + Ets + ' - ' + RechDom('TTETABLISSEMENT', Ets, False));
  end;
  RapportLst.Items.Add(TraduireMemoire('Periode         : du ' + GetControlText('PC_DATE1') + ' au ' + GetControlText('PC_DATE2')));
  TtesNatPces := GetControlText('PC_NATURE');
  Cpt := 0;
  while TtesNatPces <> '' do
  begin
    Cpt := Cpt + 1;
    NatPce := ReadTokenSt(TtesNatPces);
    if GetInfoParPiece(NatPce, 'GPP_REGROUPCPTA') = 'AUC' then
      RgpPce := TraduireMemoire(' (aucun regroupement)')
    else if GetInfoParPiece(NatPce, 'GPP_REGROUPCPTA') = 'JOU' then
      RgpPce := TraduireMemoire(' (regroupement journalier)')
    else if GetInfoParPiece(NatPce, 'GPP_REGROUPCPTA') = 'MOI' then
      RgpPce := TraduireMemoire(' (regroupement mensuel')
    else if GetInfoParPiece(NatPce, 'GPP_REGROUPCPTA') = 'SEM' then
      RgpPce := TraduireMemoire(' (regroupement hebdomadaire)');
    if Cpt = 1 then
      AddTexte(TraduireMemoire('Piece(s)        : ') + GetInfoParPiece(NatPce, 'GPP_LIBELLE') + RgpPce)
    else
      AddTexte('                  ' + GetInfoParPiece(NatPce, 'GPP_LIBELLE') + RgpPce);
  end;
  if ComptaExterne then
  begin
    if GetControlText('PC_ECRITURES') = 'X' then
      AddTexte(TCheckBox(GetControl('PC_ECRITURES')).Caption);
    AddTexte(TraduireMemoire('Exportation des écritures en fin de traitement dans :'));
    AddTexte(' ' + FichierEcr);
  end;

  {********************************
   REGROUPEMENT A TERMINER
   ********************************
    AddTexte('');
    AddTexte('***** '+TraduireMemoire('OPTIONS DE REGROUPEMENT ')+'*********************************************');
    QteOptions := 0;
    //Sur auxiliaire de centralisation
    if GetControlText('REGR_AUXICENTRALISATION') = 'X' then
      AddTexte('-> ' + TCheckBox(GetControl('REGR_AUXICENTRALISATION')).Caption);
    //Conserver le détail entreprise
    if GetControlText('REGR_DETENTREPRISE') = 'X' then
      AddTexte('-> ' + TCheckBox(GetControl('REGR_DETENTREPRISE')).Caption);
    if (GetControlText('REGR_DETENTREPRISE') = 'X') and (RegDetDesactive)  then
      AddTexte('   ' + TraduireMemoire('A ETE DESACTIVE SUITE AU PARAMETRAGE DES PIECES'));
    //Conserver le détail particulier
    if GetControlText('REGR_DETPARTICULIER') = 'X' then
      AddTexte('-> ' + TCheckBox(GetControl('REGR_DETPARTICULIER')).Caption);
    if (GetControlText('REGR_DETPARTICULIER') = 'X') and (RegDetDesactive)  then
      AddTexte('   ' + TraduireMemoire('A ETE DESACTIVE SUITE AU PARAMETRAGE DES PIECES'));
    //Regrouper si les modes de paiments diffèrent
    if GetControlText('REGR_PAIEMDIFFERENT') = 'X' then
      AddTexte('-> ' + TCheckBox(GetControl('REGR_PAIEMDIFFERENT')).Caption);
    //Règlement, regrouper sur même compte de banque
    if GetControlText('REGR_RGTBANQUE') = 'X' then
      AddTexte('-> ' + TCheckBox(GetControl('REGR_RGTBANQUE')).Caption);
    //Règlement, regrouper sur même compte de tiers
    if GetControlText('REGR_RGTTIERS') = 'X' then
      AddTexte('-> ' + TCheckBox(GetControl('REGR_RGTTIERS')).Caption);
    if QteOptions = 0 then
      AddTexte('Aucunes'); }
end;

procedure TOF_TRANSFERTCOMPTA.NomDuFichierExport;
var Natures: string;
  Extract: integer;
  DateSyst: TdateTime;
  Annee, Mois, Jour, Hre, Min, Sec, mSec: word;

  function TraiteDecode(Num: integer): string;
  begin
    if Num < 10 then
      Result := '0' + IntToStr(Num)
    else
      Result := IntToStr(Num);
  end;

begin
  DateSyst := Now;
  DecodeDate(DateSyst, Annee, Mois, Jour);
  DecodeTime(DateSyst, Hre, Min, Sec, mSec);
  DateFile := IntToStr(Annee) + TraiteDecode(Mois) + TraiteDecode(Jour) + TraiteDecode(Hre) + TraiteDecode(Min);
  Natures := GetControlText('PC_NATURE');
  while pos(';', Natures) > 0 do
  begin
    Extract := pos(';', Natures);
    Natures := copy(Natures, 1, Extract - 1) + '-' + copy(Natures, Extract + 1, length(Natures));
  end;
  //    SO_TYPECOMSX
  FichierEcr := DateFile + '-' + Natures + '.TRA';
  SetControlText('PC_EXPORTFILE', FichierEcr);
end;

procedure TOF_TRANSFERTCOMPTA.GenereFile;
var Command, EMail, DateDeb, DateFin, PathFile, Societe, User, GestionEtab,
  TypeEcriture, FichierTiers, Journal, FichierRapport: string;

  procedure ErreurParam(Champ, Msg: string);
  begin
    if Champ = '' then
      Champ := 'MANQUANT';
    RapportLst.Items.Add(Msg + Champ);
  end;

begin
  GenerationFileOk := True;
  PathFile := GetParamSoc('SO_MBOCHEMINCOMPTA');
  if copy(PathFile, length(Pathfile), 1) <> '\' then
    PathFile := PathFile + '\';
  Societe := V_PGI.CurrentAlias; //Société courante
  User := V_PGI.UserLogin; //User courant
  EMail := ''; //Envoie rapport après traitement
  FichierTiers := 'Tiers.txt'; //Fichier tiers des écritures générées
  if GetParamSoc('SO_MULTIETABCOMPTA') then //Gestion des multi etablissement en compta
    GestionEtab := 'X'
  else
    GestionEtab := '-';
  DateDeb := GetControlText('PC_DATE1'); //Date debut des écritures à extraires
  DateFin := GetControlText('PC_DATE2'); //Date fin des écritures à extraires
  if GetParamSoc('SO_COMPTAEXTERNE') then //Gestion du type ecriture (Normal ou Simulé)
    TypeEcriture := 'TYPE=[N]'  //suivant si l'intégration se fait sur la même base
  else
    TypeEcriture := 'TYPE=[S]';
  Journal := 'JOURNAL=[' + GetParamSoc('SO_EXPJRNX') + ']'; // Journal d'écriture généré
  FichierRapport := DateFile + '-RAPPORT.TXT'; //Nom du fichier d'export
  //Test si pb
  if (Societe = '') or (User = '') or (PathFile = '') or (FichierTiers = '') or (GestionEtab = '') or
    (FichierEcr = '') or (DateDeb = '') or (DateFin = '') or (TypeEcriture = '') or (FichierRapport = '') then
  begin
    RapportLst.Items.Add('***** ' + TraduireMemoire('TRAITEMENT ANNULE ') + '***************************************************');
    RapportLst.Items.Add('***** ' + TraduireMemoire('ERREUR DANS LES PARAMETRES ') + '******************************************');
    ErreurParam(Societe, TraduireMemoire('Nom de la base           : '));
    ErreurParam(User, TraduireMemoire('Utilisateur              : '));
    if (GestionEtab = 'X') then
      ErreurParam(GestionEtab, TraduireMemoire('Multi Etablissements     : OUI'))
    else
      ErreurParam(GestionEtab, TraduireMemoire('Multi Etablissements     : NON'));
    ErreurParam(PathFile, TraduireMemoire('Emplacement des fichiers : '));
    ErreurParam(FichierTiers, TraduireMemoire('Fichier des tiers        : '));
    ErreurParam(FichierEcr, TraduireMemoire('Fichier des écritures    : '));
    ErreurParam(FichierRapport, TraduireMemoire('Fichier du rapport       : '));
    ErreurParam(DateDeb, TraduireMemoire('Date début               : '));
    ErreurParam(DateFin, TraduireMemoire('Date fin                 : '));
    ErreurParam(TypeEcriture, TraduireMemoire('Type écritures           : '));
    ErreurParam(FichierRapport, TraduireMemoire('Fichier du rapport       : '));
    GenerationFileOk := False;
    exit;
  end;
  Command := '/TRF=EXPORT;' + Societe + ';' + User + ';' + EMail + ';' + PathFile + FichierTiers + ';' + GestionEtab + ';S5;JRL;' +
    PathFile + FichierEcr + ';;' + DateDeb + ';' + DateFin + ';' + TypeEcriture + ';' + Journal + ';' + FichierRapport;
  Command := 'C:\PGI00\APP\ComSx.exe ' + Command;
  if not LanceTraitement(Command) then
    GenerationFileOk := False
  else
    GenerationFileOk := True;
end;

procedure TOF_TRANSFERTCOMPTA.ImportFile;
var Command, PathFile, Societe, User, EMail, TypeEcritureGen, TypeEcritureImp: string;
begin
  GenerationFileOk := True;
  PathFile := GetParamSoc('SO_MBOCHEMINCOMPTA'); // Chemin fichier TRA
  if copy(PathFile, length(Pathfile), 1) <> '\' then
    PathFile := PathFile + '\';
  Societe := V_PGI.CurrentAlias; //Société courante
  User := V_PGI.UserLogin; //User courant
  EMail := ''; // Envoie rapport après traitement
  TypeEcritureGen := 'S'; // Gestion du type ecriture générée
  TypeEcritureImp := 'N'; // Gestion du type ecriture importé
  if (PathFile = '') or (FichierEcr = '') or (Societe = '') or (User = '') or
    (TypeEcritureGen = '') or (TypeEcritureImp = '') then
  begin
    RapportLst.Items.Add('***** ' + TraduireMemoire('TRAITEMENT ANNULE ') + '***************************************************');
    RapportLst.Items.Add('***** ' + TraduireMemoire('ERREUR DANS LES PARAMETRES ') + '******************************************');
    RapportLst.Items.Add(TraduireMemoire('Fichier des écritures   : ') + PathFile + FichierEcr);
    RapportLst.Items.Add(TraduireMemoire('Nom de la base          : ') + Societe);
    RapportLst.Items.Add(TraduireMemoire('Utilisateur             : ') + User);
    RapportLst.Items.Add(TraduireMemoire('Type écriture reçues    : ') + TypeEcritureGen);
    RapportLst.Items.Add(TraduireMemoire('Type écriture à générer : ') + TypeEcritureImp);
    GenerationFileOk := False;
    exit;
  end;
  Command := '/TRF=' + PathFile + ';' + FichierEcr + ';IMPORT;' + Societe + ';' + User + ';' + EMail + ';' + TypeEcritureGen + ';' + TypeEcritureImp + '';
  Command := 'C:\PGI00\APP\ComSx.exe ' + Command;
  if not LanceTraitement(Command) then
    GenerationFileOk := False
  else
    GenerationFileOk := True;
end;

function TOF_TRANSFERTCOMPTA.LanceTraitement(LigneDeCommande: string): boolean;
var TSI: TStartupInfo;
  TPI: TProcessInformation;
begin
  Result := True;
  FillChar(TSI, SizeOf(TStartupInfo), 0);
  TSI.cb := SizeOf(TStartupInfo);
  if CreateProcess(nil, PCHAR(LigneDeCommande), nil, nil, False, CREATE_NEW_CONSOLE or NORMAL_PRIORITY_CLASS, nil, nil, TSI, TPI) then
  begin
    while WaitForSingleObject(Tpi.hProcess, 1000) <> WAIT_OBJECT_0 do
    begin
      if WaitForSingleObject(Tpi.hProcess, 1000) = WAIT_OBJECT_0 then
        break;
      Application.ProcessMessages;
    end;
    CloseHandle(TPI.hProcess);
    CloseHandle(TPI.hThread);
  end else
  begin
    Result := False;
  end;
end;

procedure TOF_TRANSFERTCOMPTA.MiseAJourPiece(Nature, DepotD, DateDeb, DateFin: string);
var SQL, SQLL, TempNat, TempDepot, NatureTmp, NatPiece: string;
begin
  NatureTmp := Nature;
  while Nature <> '' do
  begin
    if TempNat = '' then
      TempNat := '"' + ReadTokenSt(Nature) + '"'
    else
      TempNat := TempNat + ',' + '"' + ReadTokenSt(Nature) + '"';
  end;
  while DepotD <> '' do
  begin
    if TempDepot = '' then
      TempDepot := '"' + ReadTokenSt(DepotD) + '"'
    else
      TempDepot := TempDepot + ',' + '"' + ReadTokenSt(DepotD) + '"';
  end;
  //Mise à jour des champs de comptabilisation
  SQL := 'UPDATE PIECE SET GP_ETATCOMPTA="EXP",GP_DATECOMPTA="' + USDATETIME(Date) + '" ' +
    'WHERE GP_NATUREPIECEG IN (' + TempNat + ')' +
    'AND GP_DATEPIECE>="' + USDateTime(StrToDate(DateDeb)) + '" ' +
    'AND GP_DATEPIECE<="' + USDateTime(StrToDate(DateFin)) + '" ' +
    'AND GP_ETABLISSEMENT IN (' + TempDepot + ')';
  ExecuteSQL(SQL);
  //Test s'il faut rendre les pièces/lignes mortes (traitement par nature)
  while NatureTmp <> '' do
  begin
    NatPiece := ReadTokenSt(NatureTmp);
    if GetInfoParPiece(NatPiece, 'GPP_ACTIONFINI') = 'COM' then
    begin
      SQL := 'UPDATE PIECE SET GP_VIVANTE="-" ' +
        'WHERE GP_NATUREPIECEG="' + NatPiece + '" ' +
        'AND GP_DATEPIECE>="' + USDateTime(StrToDate(DateDeb)) + '" ' +
        'AND GP_DATEPIECE<="' + USDateTime(StrToDate(DateFin)) + '" ' +
        'AND GP_ETABLISSEMENT IN (' + TempDepot + ')';
      SQLL := 'UPDATE LIGNE SET GL_VIVANTE="-" ' +
        'WHERE GL_NATUREPIECEG="' + NatPiece + '" ' +
        'AND GL_DATEPIECE>="' + USDateTime(StrToDate(DateDeb)) + '" ' +
        'AND GL_DATEPIECE<="' + USDateTime(StrToDate(DateFin)) + '" ' +
        'AND GL_ETABLISSEMENT IN (' + TempDepot + ')';
      ExecuteSQL(SQL);
      ExecuteSQL(SQLL);
    end;
  end;
end;

procedure TOF_TRANSFERTCOMPTA.OnLoad;
begin
  inherited;
end;

procedure TOF_TRANSFERTCOMPTA.OnArgument(S: string);
var ParamReg: string;
begin
  inherited;
  // Valeur pas défaut
  SetControlText('PC_DATE1', DateToStr(Date));
  SetControlText('PC_DATE2', DateToStr(Date));
  ComptaExterne := (GetParamSoc('SO_COMPTAEXTERNE'));
  SetControlProperty('PC_EXPORTFILE', 'Visible', ComptaExterne);
  SetControlProperty('TPC_EXPORTFILE', 'Visible', ComptaExterne);
  SetControlProperty('PC_ECRITURES', 'Visible', ComptaExterne);
  //  RapportFinTrt := (GetControlText('ACTIVERAPPORT') = 'X');
    //Force l'affichage du rapport
  RapportFinTrt := True;
  RapportLst := TListBox(GetControl('LB_RAPPORT'));
  // Option de regroupement
  RegEnregistre := TCheckBox(GetControl('REGR_ENREG'));
  RgpAuxiCentral := TCheckBox(GetControl('REGR_AUXICENTRALISATION'));
  RgpDetEnt := TCheckBox(GetControl('REGR_DETENTREPRISE'));
  RgpDetPart := TCheckBox(GetControl('REGR_DETPARTICULIER'));
  RgpRgtTiers := TCheckBox(GetControl('REGR_RGTTIERS'));
  RgpRgtBanque := TCheckBox(GetControl('REGR_RGTBANQUE'));
  RgpRgtTiers.Onclick := ClickRegRgtTiers;
  RgpRgtBanque.Onclick := ClickRegRgtBanque;
  TToolBarButton97(GetControl('BImprimer')).OnClick := ClickImprimer;
  //Récupération des paramètres de regroupement
  ParamReg := GetParamSoc('SO_PARAMCOMPTADIFF');
  if ParamReg <> '' then
  begin
    RgpAuxiCentral.Checked := (ReadTokenSt(ParamReg) = 'X');
    RgpDetEnt.Checked := (ReadTokenSt(ParamReg) = 'X');
    RgpDetPart.Checked := (ReadTokenSt(ParamReg) = 'X');
    TCheckBox(GetControl('REGR_PAIEMDIFFERENT')).Checked := (ReadTokenSt(ParamReg) = 'X');
    TCheckBox(GetControl('REGR_RGTBANQUE')).Checked := (ReadTokenSt(ParamReg) = 'X');
    TCheckBox(GetControl('REGR_RGTTIERS')).Checked := (ReadTokenSt(ParamReg) = 'X');
    //TCheckBox(GetControl('PC_ECRITURES')).Checked := (ReadTokenSt(ParamReg) = 'X');
    TCheckBox(GetControl('ACTIVERAPPORT')).Checked := (ReadTokenSt(ParamReg) = 'X');
    if pos('ETS', ParamReg) > 0 then
      ParamReg := copy(ParamReg, pos('ETS=', ParamReg) + 4, length(ParamReg));
    if pos('NAT', ParamReg) > 0 then
    begin
      while copy(ParamReg, 1, 4) <> 'NAT=' do
      begin
        if GetControlText('PC_ETABLISSEMENT') = '' then
          SetControlText('PC_ETABLISSEMENT', ReadTokenSt(ParamReg))
        else
          SetControlText('PC_ETABLISSEMENT', GetControlText('PC_ETABLISSEMENT') + ';' + ReadTokenSt(ParamReg));
      end;
      SetControlText('PC_NATURE', copy(ParamReg, pos('NAT=', ParamReg) + 4, length(ParamReg)));
    end;
  end;
  //Test si auxi de centralisation renseignés
  if (GetParamSoc('SO_GCCLICPTADIFF') = '') and (GetParamSoc('SO_GCCLICPTADIFFPART') = '') and
    (GetParamSoc('SO_GCFOUCPTADIFF') = '') and (GetParamSoc('SO_GCFOUCPTADIFFPART') = '') then
  begin
    RgpSurAuxiCentral := False;
    RgpAuxiCentral.Enabled := False;
    RgpAuxiCentral.Checked := False;
  end else
  begin
    RgpSurAuxiCentral := True;
  end;
  RegEnregistre.Checked := True;
end;

procedure TOF_TRANSFERTCOMPTA.OnClose;
var Enreg: string;
begin
  inherited;
  if RgpAuxiCentral.Checked then
    Enreg := 'X'
  else
    Enreg := '-';
  if RgpDetEnt.Checked then
    Enreg := Enreg + ';' + 'X'
  else
    Enreg := Enreg + ';' + '-';
  if RgpDetPart.Checked then
    Enreg := Enreg + ';' + 'X'
  else
    Enreg := Enreg + ';' + '-';
  if TCheckBox(GetControl('REGR_PAIEMDIFFERENT')).Checked then
    Enreg := Enreg + ';' + 'X'
  else
    Enreg := Enreg + ';' + '-';
  if TCheckBox(GetControl('REGR_RGTBANQUE')).Checked then
    Enreg := Enreg + ';' + 'X'
  else
    Enreg := Enreg + ';' + '-';
  if TCheckBox(GetControl('REGR_RGTTIERS')).Checked then
    Enreg := Enreg + ';' + 'X'
  else
    Enreg := Enreg + ';' + '-';
  if TCheckBox(GetControl('PC_ECRITURES')).Checked then
    Enreg := Enreg + ';' + 'X'
  else
    Enreg := Enreg + ';' + '-';
  if TCheckBox(GetControl('ACTIVERAPPORT')).Checked then
    Enreg := Enreg + ';' + 'X'
  else
    Enreg := Enreg + ';' + '-';
  Enreg := Enreg + ';ETS=' + GetControlText('PC_ETABLISSEMENT');
  Enreg := Enreg + ';NAT=' + GetControlText('PC_NATURE');
  SetParamSoc('SO_PARAMCOMPTADIFF', Enreg);
  RapportLst.free;
end;

procedure TOF_TRANSFERTCOMPTA.ClickRegRgtBanque(Sender: Tobject);
begin
  if (RgpRgtBanque.Checked = False) and (RgpRgtTiers.Checked = True) then
    RgpRgtTiers.Checked := False;
end;

procedure TOF_TRANSFERTCOMPTA.ClickRegRgtTiers(Sender: Tobject);
begin
  if (RgpRgtTiers.Checked = True) and (RgpRgtBanque.Checked = False) then
    RgpRgtBanque.Checked := True;
end;

procedure TOF_TRANSFERTCOMPTA.ClickImprimer(Sender: TObject);
var TobToPrint: TOB;
  Cpt: integer;
begin
  TobToPrint := TOB.Create('', nil, -1);
  for Cpt := 0 to RapportLst.Items.Count - 1 do
    UtilTobCreat(TobToPrint, '', '', RapportLst.Items[Cpt], '');
  UtilTobPrint(TobToPrint, Ecran.Caption);
  TobToPrint.free;
end;

initialization
  registerclasses([TOF_TRANSFERTCOMPTA]);
end.
