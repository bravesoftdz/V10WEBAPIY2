{***********UNITE*************************************************
Auteur  ...... : Franck VAUTRAIN
Créé le ...... : 08/07/2016
Modifié le ... : 08/07/2016
Description .. : Objet servant à gérer excel via pilotage OLE :
Suite......... : dans les cadre de la gestion des métrés
Suite..........:
Mots clefs ... : Excel;OLE;Outils
*****************************************************************}

unit UtilsMetresXLS;

interface

Uses  Windows,
      Messages,
      SysUtils,
      Classes,
      CBPPath,
      ComObj,
      ComCtrls,
      Controls,
      FileCtrl,
      Hctrls,
      HMsgBox,
      HEnt1,
      Variants,
      Graphics,
      uEntCommun,
      {$IFDEF EAGLCLIENT}
      maineagl,
      {$ELSE}
      fe_main,
      db,
      EntGC,
      ActiveX,
      TLHelp32,
      ShellAPI,
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
      {$ENDIF}
      Forms,
      UAuditPerf,
      uTOB;

type  TMetreArt = class

  private
     Public
        vMSExcel      : Variant;  //Instance Excel
        vXLWorkbooks  : Variant;  //ensemble des classeurs de l'instance
        vXLWorkbook   : Variant;  //Classeur
        vXLsheet      : Variant;  //Feuille
        SSheetName    : String;   //Nom de la Feuille
        SRange        : String;   //Tranche de Cellule
        SValue        : String;   //Valeur d'une Cellule
        VCell         : Variant;  //Cellule
        vReadOnly     : Boolean;  //Détermination si ouverture d'un classeur en lecture seule
        //
        FForm         : TForm;
        //
        fNomMachine   : string;   //Nom du Poste Utilisateur
        fRepMetre     : String;   //Répertoire de stockage des fichier métrés par Article Sur le Serveur
        fRepMetreLocal: String;   //Répertoire de stockage Local
        fFileName     : String;   //Nom du fichier XLS associé à l'Article
        fFichierVide  : string;   //Nom du fichier vide
        fRepMetreDoc  : String;   //Répertoire de stockage des fichier métrés par Document Sur Serveur
        fRepMetreDocLocal : String;   //Répertoire de stockage Local
        //
        fTypeArt      : String;   //Type Article (ART, OIUV, PRE,...)
        fArticle      : string;   //code Article
        ArticleXLS    : String;   //Si modification d'un sous-détail pour n'ouvrir que le fichier correspondant au sous-détail sélectionné
        SAfNatProposition : String;
        SAfNatAffaire : String;
        AchatVente    : String;
        //
        UniqueBlo     : Integer;
        IndexFile     : Integer;
        //
        fModeAudit    : Boolean;  //Permet de déclencher un traceur si nous sommes en mode SAV
        fAuditPerf    : TAuditClass;
        //
        OkExcel       : Boolean;
        OkMetreBiblio : Boolean;
        OkMetreDoc    : Boolean;
        OkMetreActif  : Boolean;
        OkMetre       : Boolean;  //Détermine si la saisie des métrés est possible sur un document
        OKDuplication : Boolean;
        OkAValider    : Boolean;
        OkCreation    : Boolean;
        OkClick_Elips : Boolean;
        OkOpenXLS     : Boolean;
        //
        Action        : TActionFiche;
        //
        TTobPiece     : TOB;
        TTObOuvrage   : TOB;
        TTOBMetres    : TOB;
        TTobVariables : TOB;
        TTobVarDoc    : TOB;
        TTobSaisieVar : TOB;
        //

        constructor CreateArt(TypeArt, Article : String);
        constructor CreateDoc;
        destructor  destroy;    override;

        property ModeAudit      : Boolean     read fModeAudit write fModeAudit;
        property AuditPerf      : TAuditClass read fAuditperf write fAuditPerf;

        property RepMetre       : String  read fRepMetre      write fRepMetre;
        property RepMetreLocal  : String  read fRepMetreLocal write fRepMetreLocal;
        property NomMachine     : String  read fNomMachine    write fNomMachine;
        property FileName       : String  read fFileName      write fFileName;
        property FichierVide    : String  read fFichierVide   write fFichierVide;
        //
        property TypeArticle    : String  read FTypeArt       write fTypeArt;
        property Article        : String  read fArticle       write fArticle;
        property AfNatAffaire   : String  Read SAfNatAffaire  write SAfNatAffaire;
        property AfNatProposition :  String  Read SAfNatProposition write SAfNatProposition;
        //
        property Ok_Excel       : Boolean read OkExcel        write OkExcel;
        property Ok_MetreBiblio : Boolean read OkMetreBiblio  write OkMetreBiblio;
        property Ok_MetreDoc    : Boolean read OkMetreDoc     write OkMetreDoc;
        property Ok_MetreActif  : boolean read OkMetreActif   write OkMetreActif;
        property Ok_Metre       : boolean read OkMetre        write OkMetre;
        property Ok_Duplication : boolean read OkDuplication  write OkDuplication;
        property Ok_AValider    : boolean read OkAValider     write OkAValider;
        //
        procedure AccessCelluleXLS(ARange : String);
        procedure AccessWorkSheetXLS(NomFeuille : string);
        procedure ChangeValueCelluleXLS(Arange, AValue : string);
        procedure ChargeFeuilleWithVariables;
        procedure ChargementMetre(CleDoc: R_CleDoc);
        procedure ChargementQtefichierXLSdansTob(TobTravail: TOB);
        procedure ChargementSaisieVarWithVardoc(TypeVar : String; Cledoc : R_CLedoc; OkAffiche :Boolean=True; OkSaisie :Boolean=True);
        procedure ChargementTobVarDocWithTobVarApp(Cledoc : R_CLEDOC; OkAffiche : Boolean=True; OkSaisie : Boolean=True);
        procedure ChargementTobVarDocWithTobVarArt(TOBL, TOBLOuvrage: TOB; OkAffiche : Boolean=True; OkSaisie : Boolean=True);
        procedure ChargementVariablesArt(CodeArt : String);
        procedure ChargementVariablesDoc(CleDoc: R_CleDoc);
        procedure ChargeQteDetailMetre(TObOuvrage, TOBL: TOB);
        procedure CloseWorkbookXLS(vSaveChanges : Boolean);
        procedure CloseXLS;
        procedure ControleDelVardoc(TypeVarDoc, CodeArt: string; TobSaisieDoc: ToB);
        procedure ControleFichierMetresLigne(TOBL, TOBLMETRE: TOB);
        function  ControleMetre     : Boolean;
        function  ControleMetreDoc(Cledoc : String) : Boolean;
        function  CopieVariables(NewArticle: string): boolean;
        procedure CreateEnregTOBOuvrage(TOBART, TOBL: TOB);
        procedure CreateWorkbookXLS(NomClasseur: string);
        Procedure CreateWorkSheetXLS(Before, After, Count, Xltype : OleVariant; NomFeuille : string );
        function  CreationRepertoire(NomRep: String) : Boolean;
        Function  CreationLigneVarDoc(TobLigneVar, TOBL : TOB) : TOB;
        procedure DetruitMetreDoc(TobPiece: Tob);
        function  FichierOuvert(NomFichier: string): Boolean;
        function  FindLigneOuvrageinDoc(TOBD: TOB): TOB;
        function  FormatArticle(Article: String): String;
        function  GetPrefixeTable(TOBL: TOB): string;
        function  IsArticle(TOBL: TOB): boolean;
        function  IsSousDetail(TOBL: TOB): boolean;
        procedure LectureTableVariable(TypeVar : String; Cledoc : R_CLEDOC);
        procedure LibereTob;
        procedure LigneTitreSaisieVariable(Cledoc: R_CLEDOC; Libelle: string;OkAffiche :Boolean=True);
        function  MetreAttached(TOBL: TOB; NumOrdre : Integer): boolean;
        function  MetreResultatExiste: Boolean;
        procedure OpenWorkbookXLS(NomDuFichier : String);
        function  OpenXLS(Ok_visible : Boolean=False; Ok_Controle : Boolean=True) : Boolean;
        function  QteModifie(TOBL: TOB; NumOrdre : Integer; OldQte, NewQte : Double; OK_SsDetail : Boolean=False): boolean;
        Function  ReadValueCelluleXLS(Arange : string) : Variant;
        procedure RecalculsousDetailOuvrage(TobEntOuv, TOBOuvrage: TOB);
        Function  RecupValeurMetreDoc(Nomfichier : string): Double;
        Procedure RenommeWorkSheetXLS(NouveauNom : String);
        procedure RenumerotationTob(TobAtraiter: Tob);
        procedure SauveTob;
        procedure SaveWorkbookXLS;
        procedure SupprimeMetre(TOBTemp: TOb);
        procedure SupprimeVariable(Article: string);
        Procedure TraitementMetre(Cledoc : R_CLEDOC; TobEntOuv, TOBOuvrage, TOBL : Tob);
        Procedure TraitementMetreArticle(TOBL: Tob);
        procedure TraitementMetreModif(Cledoc : R_CLEDOC; TobEntOuv, TOBOuvrage, TOBL: Tob);
        procedure TraitementMetreOuvrage(TOBL, TOBOuv: Tob);
        procedure TraitementVariableCachees(ToblArticle : TOB; Cledoc : R_CLEDOC; TypeArticle : String; OkAffiche : Boolean=True; OkSaisie : Boolean=True);
        procedure ValideMetreDoc(TobPiece: Tob);

        //procedure propre aux ficher XLA
        procedure ActivationDesactivation_BarreOutil(Active: boolean);
        function  AutorisationMetre(NatPiece: string): Boolean;
        procedure CalculMetreDocXLsArt(TOBL : TOB);
        procedure CalculMetreDocXLsOuvrage(TOBL : TOB);
        procedure ColorCellule(Adresse: Variant; CouleurEcriture, CouleurFond: TColor);
        function  ControlePageVide: boolean;
        procedure CopieLocaltoServeur(FileNameO, FileNameD : String);
        procedure CreateProcedure_RecupValeur(NomProcedure : string);
        procedure Creer_Menu_PopUp(CommandBarsName, MenuName, ModuleName, TagName: string; IndexMenu : Integer);
        procedure Delete_Menu_PopUp(CommandBarsName, TagName : string);
        procedure DuplicationFichierXLS(NomOrigine, NomArrive : String);
        procedure FermerMetreXLs;
        function  GereMetreXLS(TobPiece, TOBL, TobOuvrage, TobEntOuv : TOB; Okclick : Boolean = False): Double;
        function  IsRunning(NomApplication : string; StopProcess:Boolean):Boolean;
        procedure OuvrirMetreXLs;
        function  RechercheEtiquette(NomEtiquette : String) : Variant;
        procedure RenommeEtiquette(AncienNom, NouveauNom : String);
        function  SaisieVariables(Cledoc : R_CLEDOC) : boolean;
        procedure SuppressionLigneMetreDoc(TOBL: TOB);
        procedure SupprimeFichierXLS(FichierXls : string);
        procedure SupprimeRepertoireDoc;
        //
        procedure SaisieVarDoc(CLEDOC : R_CLEDOC; TypeVardoc : String);
        procedure KillExcel;

    End;

var TheMetreShare : TMetreArt;

implementation
  Uses  UtilFichiers,
        utilxlsBTP,
        AGLInit,
        //FactUtil,
        FactOuvrage,
        UtilProcesses,
        ParamSoc,
        UCotraitance;

  Const xlChart                 = -4109;
        xlExcel4IntlMacroSheet  = 4;
        xlExcel4MacroSheet      = 3;
        xlWorksheet             = -4167;
        //
        XlsAutomatic            = -4105;
        XlsManual               = -4135;
        XlsSemiAuto             = 2;
        //
        xlCellTypeAllFormatConditions   = -4172; //Cellules de n'importe quel format.
        xlCellTypeAllValidation         = -4174; //Cellules présentant des critères de validation.
        xlCellTypeBlanks                = 4;     //Cellules vides.
        xlCellTypeComments              = -4144; //Cellules contenant des commentaires.
        xlCellTypeConstants             = 2;     //Cellules contenant des constantes.
        xlCellTypeFormulas              = -4123; //Cellules contenant des formules.
        xlCellTypeLastCell              = 11;    //Dernière cellule dans la plage utilisée.
        xlCellTypeSameFormatConditions  = -4173; //Cellules de même format.
        xlCellTypeSameValidation        = -4175; //Cellules présentant les mêmes critères de validation.
        xlCellTypeVisible               = 12;    //Toutes les cellules visibles.
        //
        ErrorDiv                        = -2146826281;  //#Div/0!
        ErrorNA                         = -2146826246;  //#N/A
        ErrorName                       = -2146826259;  //#Name?
        ErrorNull                       = -2146826288;  //#Null!
        ErrorNum                        = -2146826252;  //#Num!
        ErrorRef                        = -2146826265;  //#Ref!
        ErrorValue                      = -2146826273;  //#Value!

constructor TMetreArt.CreateArt (TypeArt, Article : String);
begin

  fTypeArt      := TypeArt;
  fArticle      := FormatArticle(Article);

  IF fArticle = '' then
  begin
    OkMetreBiblio := False;
    OkExcel       := False;
    OKMetreDoc    := False;

    fRepMetre     := '';

    FNomMachine   := '';

    fFileName     := '';
  end
  else
  begin
    OkMetreBiblio := GetParamSocSecur ('SO_BTMETREBIB',   False);
    OkExcel       := GetParamSocSecur ('SO_METRESEXCEL',  False);
    OKMetreDoc    := GetParamSocSecur ('SO_BTMETREDOC',   False);
    
    if OkMetreBiblio or OKMetreDoc then
      fRepMetre   := RecupRepertbase
    else
      fRepMetre   := '';

    FNomMachine   := ComputerName ;

    //Création du nom du fichier Article XLS
    fFileName := fArticle + '.xlsx';
  end;

end;

destructor TMetreArt.destroy;
begin
  inherited;

  vMSExcel      := Unassigned;
  vXLWorkbooks  := Unassigned;
  vXLWorkbook   := Unassigned;
  vXLsheet      := Unassigned;
  VCell         := Unassigned;

end;

function TMetreArt.OpenXLS(Ok_visible : Boolean=False; Ok_Controle : Boolean=True) : Boolean;
begin

  Result    := False;

  vMSExcel  := UnAssigned;

  if Ok_Controle then
  begin
    if IsExcelLaunched then
    BEGIN
      PGIBox ('Vous devez fermer préalablement les instances d''EXCEL', 'Gestion des métrés');
      //ShellExecute(0,PChar('taskkill /F /IM Excel.exe'),'', nil,nil, SW_RESTORE);
      Exit;
    END;
  end;

  try
    vMSExcel := CreateOleObject( 'Excel.Application' );
    Result := True;
  except
    vMSExcel := UnAssigned;
  end;

  if VarType(vMSExcel) = varEmpty then Exit;

  vMSExcel.Visible := Ok_Visible;

end;

// Teste si une application est en cours d'exécution
// StopProcess indique si on termine l'application 'NomApplication'

// Code proposé par Thunder_nico
function TMetreArt.IsRunning(NomApplication : string; StopProcess:Boolean):Boolean;
var
  ProcListExec  : TProcessentry32;
  PrhListExec   : Thandle;
  Continu : Boolean;
  isStarted : Boolean;
  HandleProcessCourant : Cardinal;
  PathProcessCourant : string;
  ProcessCourant :String;

begin
  // Liste des applications en cours d'exécution
  // Initialisation des variables et récuperation de la liste des process
  ProcListExec.dwSize:=sizeof(ProcListExec);
  Continu   := True;
  isStarted := False;
  FreeAndNil(PrhListExec);
  FreeAndnil(HandleProcessCourant);

  Try
    // Récupére la liste des process en cours d'éxécution au moment de l'appel
    PrhListExec :=CreateToolhelp32Snapshot(TH32CS_SNAPALL,0);
    if (PrhListExec <> INVALID_HANDLE_VALUE) then
    begin
      //On se place sur le premier process
      Process32First(PrhListExec,ProcListExec);
      // Tant que le process recherché n'est pas trouvé et qu'il reste
      // des process dans la liste, on parcourt et analyse la liste
      while Continu do
      begin
        ProcessCourant := Uppercase(ExtractFileName(ProcListExec.szExeFile));
        ProcessCourant := ChangeFileExt(ProcessCourant,'');
        isStarted := (ProcessCourant = Uppercase(NomApplication));
        if isStarted then
          begin
            HandleProcessCourant := ProcListExec.th32ProcessID;
            PathProcessCourant := ExtractFilepath(ProcListExec.szExeFile);
            Continu := False;
          end
        else // Recherche le process suivant dans la liste
          Continu := Process32Next(PrhListExec,ProcListExec);
      end;
      //
      if StopProcess then
      begin
        if isStarted then
        begin  // Termine le process en indiquant le code de sortie zéro
          TerminateProcess(OpenProcess(PROCESS_TERMINATE,False,HandleProcessCourant),0);
          Sleep(500); // Laisse le temps au process en cours de suppression de s'arrêter
        end;
      end;
    end;
  Finally
    CloseHandle(PrhListExec); // Libère les ressources
    Result := isStarted;
  end;

end;

Procedure TMetreArt.CreateWorkbookXLS(NomClasseur: string);
begin

  if VarisEmpty(vMSExcel) then Exit;

  vXLWorkbooks := vMSExcel.Workbooks;
  vXLWorkbook  := vXLWorkbooks.Add;

end;

Procedure TMetreArt.OpenWorkbookXLS(NomDuFichier : String);
begin

  if VarisEmpty(vMSExcel) then Exit;

  vXLWorkbooks := vMSExcel.Workbooks;
  
  vXLWorkbook := vXLWorkbooks.Open(NomDuFichier);

end;

Procedure TMetreArt.CreateWorkSheetXLS(Before, After, Count, Xltype : OleVariant; NomFeuille : string );
begin

  //Before : feuille avant laquelle la nouvelle feuille doit être créée
  //After  : feuille après laquelle la nouvelle feuille doit être créée
  //Count  : Nombre de feuille à ajouter
  //Type   : Par défaut xlWorksheet

  if VarisEmpty(vXLWorkbook) then Exit;

  if NomFeuille = '' then Exit;

  vXLsheet      := vXLWorkbook.Worksheets.Add(Before,After,Count,xlType);
  vXLsheet.Name := NomFeuille;

end;

Procedure TMetreArt.AccessWorkSheetXLS(NomFeuille : string);
begin

  if VarisEmpty(vXLWorkbook) then Exit;

  if NomFeuille = '' then Exit;

  vXLsheet := vXLWorkbook.Worksheets[NomFeuille];

end;

Procedure TmetreArt.RenommeWorkSheetXLS(Nouveaunom : string);
begin

  if VarisEmpty(vXLsheet) then Exit;

  if NouveauNom = '' then Exit;

  vXLsheet.name := NouveauNom;

end;

Procedure TMetreArt.AccessCelluleXLS(ARange : String);
begin

  if VarisEmpty(vXLsheet) then Exit;

  VCell := vXLsheet.Range[Arange];

end;

Function TMetreArt.ReadValueCelluleXLS(Arange : string) : Variant;
begin

  if VarisEmpty(vXLsheet) then Exit;

  Result  := '';

  vCell   := vXLsheet.Range[aRange];
  Result  := vCell.Value;

end;

Procedure TMetreArt.ChangeValueCelluleXLS(Arange, AValue : string);
begin

  if VarisEmpty(vXLsheet) then Exit;

  vCell       := vXLsheet.Range[aRange];
  vCell.Value := aValue;

end;

Procedure TMetreArt.SaveWorkbookXLS;
begin

  If VarisEmpty(vMSExcel) then Exit;

  if VarisEmpty(vXLWorkbook) then Exit;

  vXLWorkbook.Save;

end;

Procedure TMetreArt.CloseWorkbookXLS(vSaveChanges : Boolean);
begin

  If VarisEmpty(vMSExcel) then Exit;

  if not FichierOuvert(fRepMetreLocal + fFileName) then exit;

  //==> peut provoquer un message d'erreur l'appel a été rejeté par l'appelé
  //Je suis allé voir sur Internet mais impossible de comprendre d'où vient l'erreur d'autant plus qu'elle n'est pas constante...
  //il faudrait traduire du C# en Delphi et ce n'est pas dans mes compétences.....
  vXLWorkbook.Close(vSaveChanges, fRepMetreLocal + fFileName);

end;

//Test si un fichier est ouvert ou indisponible
function TMetreArt.FichierOuvert(NomFichier : string):Boolean;
// CodeRetour = 0 pas ouvert - 32 : violation de partage - sinon autre erreur
Var CodeRetour  : Integer;
    F           : TextFile;
begin

  FichierOuvert := True ;

  // Test du fichier
  AssignFile(F, NomFichier);

  {$I-}
  Reset(F);

  {$I+}
  CodeRetour := IOResult ;

  Case CodeRetour Of
    // Pas d'erreur
    0 :Begin
         CloseFile(F);
         FichierOuvert := False ;     //ShowMessage('Le fichier existe et n''est pas ouvert.');
       End;
    //L'erreur 32 est une violation de partage
    32:
       Begin
         FichierOuvert := True
       End ;
  Else
    // Autre erreur
    PgiInfo('Erreur ' + IntToStr(CodeRetour), 'Fichier XLS');
  End;

End;

procedure TMetreArt.KillExcel;
begin

  KillExe('EXCEL.EXE');
  sleep (10);
end;

Procedure TMetreArt.CloseXLS;
begin

  If VarisEmpty(vMSExcel) then Exit;

  vMSExcel.Quit;

  vMSExcel      := Unassigned;
  vXLWorkbooks  := Unassigned;
  vXLWorkbook   := Unassigned;
  vXLsheet      := Unassigned;
  VCell         := Unassigned;

  KillExcel;

end;

function TMetreArt.ControleMetre : Boolean;
begin

  Result := false;

  if not OkMetreBiblio Then exit;

  if not OkExcel Then Exit;

  if not OfficeExcelDispo then
  begin
    PGIBox('Office n''est pas installé sur ce poste', 'Gestion des métrés');
    OkExcel := False;
    exit;
  end;

  if fRepMetre = '' then
  begin
    PGIBox('le répertoire métrés Serveur est à blanc dans les paramètres société', 'Gestion des métrés');
    OkExcel := False;
    Exit;
  end;

  if FNomMachine = '' then
  begin
    PGIBox('le Nom de la machine est à blanc', 'Gestion des métrés');
    Exit;
  end;

  fRepMetre := fRepMetre + '\';
  OkExcel := CreationRepertoire(fRepMetre);

  //Contrôle répertoire des fichiers métrés article sur Serveur
  fRepMetre := fRepMetre + 'Bibliotheque' + '\';
  OkExcel := CreationRepertoire(fRepMetre);

  //création du répertoire de sauvegarde des fichiers métrés sur Postre Client
  fRepMetreLocal := 'C:\PGI01\';
  OkExcel := CreationRepertoire(fRepMetreLocal);

  fRepMetreLocal := fRepMetreLocal+ 'Metres\';
  OkExcel := CreationRepertoire(fRepMetreLocal);

  fRepMetreLocal := fRepMetreLocal+ 'Bibliotheque\';
  OkExcel := CreationRepertoire(fRepMetreLocal);

  //Pour une installation TSE on est obligé de gérer le nom de la machine
  fRepMetreLocal := fRepMetreLocal+ FNomMachine + '\';
  OkExcel := CreationRepertoire(fRepMetreLocal);

  //On charge un fichier vide...
  fFichierVide := 'c:\PGI00\STD\ARTVIDE.xlsx';
  if not FileExists(fFichierVide) then
  Begin
    PGIBox('le fichier ' + fFichierVide + ' n''existe pas.', 'Gestion des métrés');
    OkExcel := False;
    exit;
  end;

  Result := true;

end;

Procedure TMetreArt.CopieLocaltoServeur(FileNameO, FileNameD : String);
begin

  if FileNameO = '' then exit;

  if FileNameD = '' then exit;

  //on copie le fichier XLS du Répertoire local vers le répertoire Serveur (???)
  if FileExists(FileNameO) then CopieFichier(FileNameO, FileNameD);

end;


function TMetreArt.FormatArticle(Article: String): String;
begin

  Result := Article;

  if fTypeArt = 'OUV' then Result := trim(Result);

  Result := Replace(Result, '\', '_');
  Result := Replace(Result, '/', '_');
  Result := Replace(Result, '''', '_');
  Result := Replace(Result, '"', '_');

end;

function TMetreArt.CopieVariables(NewArticle: string) : boolean;
var requete : string;
    I       : Integer;
    QQ      : TQuery;
    TobVSce : Tob;
    TobVDes : Tob;
begin

  Result := False;

  if not OkMetreActif then Exit;

  TobVSce := TOB.Create('Var Source', nil, -1);
  TobVDes := TOB.Create('Var Dest', nil, -1);

  Requete := 'SELECT * FROM BVARIABLES WHERE BVA_ARTICLE="' + fArticle + '"';
  QQ := OpenSQL(Requete, True,-1,'',true);

  if not QQ.eof then
  begin
    TobVSce.LoadDetailDB('BVARIABLES', '', '', QQ, False);
    TobVDes.Dupliquer(TobVSce,True,True);
    For I:=0 To TobVDes.detail.count-1 do
    begin
      TobVDes.Detail[I].Putvalue('BVA_ARTICLE', NewArticle);
    end;
    TobVDes.InsertOrUpdateDB;
  end;

  Ferme(QQ);

  TobVSce.free;
  TobVDes.free;

  result := true;

end;

procedure TMetreArt.OuvrirMetreXLs;
begin

  if not OpenXLS(True, False) then Exit;

  OpenWorkbookXLS(fRepMetreLocal + fFileName);

  Delete_Menu_PopUp('cell','Métrés');

  Creer_Menu_PopUp('Cell','Récupération Valeur','RecupValeur','Métrés',1);

  ActivationDesactivation_BarreOutil(False);

  CreateProcedure_RecupValeur('MesArticles');

  vMSExcel.Calculation   := XlsAutomatic;
  vMSExcel.MaxChange     := 0.001;
  vMSExcel.DisplayAlerts := false; //enlève les boîtes de dialogue dans Excel
  vXLWorkbook.PrecisionAsDisplayed := False;

end;

procedure TMetreArt.FermerMetreXLs;
begin

  //Fermeture du classeur
  CloseWorkbookXLS(True);

  //Fermeture d'Excel
  CloseXLS;

end;

Procedure TMetreArt.SupprimeFichierXLS(FichierXls : string);
begin

  if FileExists(FichierXls) then
  begin
    if not DeleteFichier(FichierXls) then PGIBox('Suppression du fichier ' + FichierXls + ' impossible.', 'Gestion des métrés');
  end;

end;

Procedure TMetreArt.DuplicationFichierXLS(NomOrigine, NomArrive : String);
begin

  //Fermeture d'Excel
  CloseXLS;

  //Copie du fichier d'origine sur nouvelle déclinaison dans répertoire du poste de travail
  if FileExists(fRepMetreLocal + NomOrigine) then
  begin
    //Copie du fichier d'origine sur nouvelle déclinaison dans répertoire du poste de travail
    CopieFichier(fRepMetreLocal + NomOrigine, fRepMetreLocal + NomArrive);
    //Puis copie sur le serveur
    CopieFichier(fRepMetreLocal + NomArrive, FrepMetre + NomArrive);
  end;

end;

procedure TMetreArt.Creer_Menu_PopUp(CommandBarsName, MenuName, ModuleName, TagName : string; IndexMenu : Integer);
Var Icbc : Variant;
begin

  if VarIsEmpty(vMSExcel) then Exit;

  Icbc := vMSExcel.CommandBars[CommandBarsName].controls.add(1, IndexMenu, True);
  If IndexMenu = 1 then Icbc.BeginGroup := True;
  Icbc.Caption  := MenuName;
  Icbc.OnAction := ModuleName;
  Icbc.Enabled  := True;
  Icbc.Tag      := TagName;

end;

//Cette procédure va créer une ligne récupération valeur dans le Pop-up d'Excel
Procedure TMetreArt.Delete_Menu_PopUp(CommandBarsName, TagName : string);
Var Icbc : Variant;
    i    : Integer;
    Exit : boolean;
begin

  i := 0;
  Exit := False;

  if VarIsEmpty(vMSExcel) then Exit:=True;

  repeat
    Inc(i);
    if i > vMSExcel.CommandBars[CommandBarsName].controls.count then
      Exit := true
    else
    begin
      Icbc := vMSExcel.CommandBars[CommandBarsName].controls.item[i];
      if VarIsNull(Icbc) then
        continue
      else
      begin
        if IcBc.tag = TagName then icbc.delete;
      end;
    end;
  until Exit;

end;

Procedure TMetreArt.ActivationDesactivation_BarreOutil(Active : boolean);
begin

    if VarIsEmpty(vMSExcel) then Exit;

    vMSExcel.CommandBars['Standard'].Visible            := Active;
    vMSExcel.CommandBars['Stop Recording'].Visible      := Active;
    vMSExcel.CommandBars['Chart Menu Bar'].Visible      := Active;
    vMSExcel.CommandBars['Control Toolbox'].Visible     := Active;
    vMSExcel.CommandBars['Borders'].Visible             := Active;
    vMSExcel.CommandBars['Exit Design Mode'].Visible    := Active;
    vMSExcel.CommandBars['Drawing'].Visible             := Active;
    vMSExcel.CommandBars['Forms'].Visible               := Active;
    vMSExcel.CommandBars['External Data'].Visible       := Active;
    vMSExcel.CommandBars['Chart'].Visible               := Active;
    vMSExcel.CommandBars['Shadow Settings'].Visible     := Active;
    vMSExcel.CommandBars['Picture'].Visible             := Active;
    vMSExcel.CommandBars['3-D Settings'].Visible        := Active;
    vMSExcel.CommandBars['Full Screen'].Visible         := Active;
    vMSExcel.CommandBars['PivotTable'].Visible          := Active;
    vMSExcel.CommandBars['Reviewing'].Visible           := Active;
    vMSExcel.CommandBars['Circular Reference'].Visible  := Active;
    vMSExcel.CommandBars['Web'].Visible                 := Active;
    vMSExcel.CommandBars['WordArt'].Visible             := Active;
    vMSExcel.CommandBars['Visual Basic'].Visible        := Active;

end;

function TMetreArt.RechercheEtiquette(NomEtiquette : string) : Variant;
Var Cellule   : Variant;
    I         : Integer;
    NCell     : String;
begin

  result := '';

  if VarisEmpty(vMSExcel) then Exit;

  if VarIsEmpty(vXLWorkbook) then Exit;

  //if ControlePageVide then Exit;

  for i := 1 to vXLWorkbook.Names.Count do
  begin
    Cellule := vXLWorkbook.Names.Item(i, EmptyParam, EmptyParam);
    if VarIsEmpty(Cellule) then Continue;
    NCell := Cellule.Name;
    if NCell = NomEtiquette then
    begin
      result   := Cellule.RefersToRange.Address;
      vXLsheet := vXLWorkbook.ActiveSheet;
      Break;
    end;
  end;

end;

Procedure TMetreArt.ColorCellule(Adresse : Variant; CouleurEcriture, CouleurFond : TColor);
Var Cellule : Variant;
begin

  if VarIsEmpty(vXLsheet) then Exit;

  Cellule := vXLsheet.Range[Adresse];
  if VarIsEmpty(Cellule) then Exit;

  Cellule.select;
  Cellule.Font.ColorIndex     := CouleurEcriture;
  Cellule.Interior.ColorIndex := Couleurfond;

end;

procedure TMetreArt.RenommeEtiquette(AncienNom, NouveauNom : String);
Var Cellule   : String;
begin

  if VarIsEmpty(vXLsheet) then Exit;

  //On réactivera cela quand on sera lister les cellules de la feuilles pour
  //vérifier si une cellule n'est pas celle qu'il faut erécupérer
  //En attendant il faut forcément au moins une colonne avec une valeur
  //renseignée.

  //if not ControlePageVide then
  //begin
    //Récupère l''adresse de la cellule ou la plage nommée.
    Cellule := vXLsheet.Range[AncienNom].name;
    if Cellule <> '' then
    begin
      //suppression du nom
      vXLsheet.Range[AncienNom].Name.Delete;
      //renomme la plage initiale
      vXLsheet.Range[Cellule].Name := NouveauNom;
    end;
  //end;

end;

Function TMetreArt.ControlePageVide : boolean;
begin

  Result := False;

  if VarIsEmpty(vXLsheet) then Exit;

  //Sinon on vérifie si la feuille a des cellules de rensignées...
  if (vXLsheet.UsedRange.Rows.Count = 1) And (vXLsheet.UsedRange.Columns.Count = 1) then Result := True;

end;

Procedure TMetreArt.CreateProcedure_RecupValeur(NomProcedure : string);
Var Module      : Variant;
    N           : integer;
begin
     
  if VarisEmpty(vMSExcel) then Exit;

  if VarIsEmpty(vXLWorkbook) then Exit;
     
  Module := vMSExcel.VBE.ActiveVBProject.VBComponents.Add(1);
  Module.name := NomProcedure;

  N :=  Module.CodeModule.CountOfLines;

  Module.CodeModule.InsertLines(N+1, 'Private Sub RecupValeur');
  Module.CodeModule.InsertLines(N+2, ' Sauve_Adresse = ActiveCell.Address');
  Module.CodeModule.InsertLines(N+3, '''');
  Module.CodeModule.InsertLines(N+4, '''Controle Existance d''une zone PGI_RESULTAT');
  Module.CodeModule.InsertLines(N+5, ' Set nms = Application.ActiveWorkbook.Names');
  Module.CodeModule.InsertLines(N+6, ' Set wks = Application.ActiveWorkbook.ActiveSheet');
  Module.CodeModule.InsertLines(N+7, ' For I = 1 To nms.Count');
  Module.CodeModule.InsertLines(N+8, '   Adresse = nms(I).RefersToRange.Address');
  Module.CodeModule.InsertLines(N+9, '   If IsEmpty(Adresse) = False Then');
  Module.CodeModule.InsertLines(N+10,'    If nms(I).Name = Application.ActiveWorkbook.ActiveSheet.name & "!PGI_RESULTAT" Then');
  Module.CodeModule.InsertLines(N+11,'       Range(Adresse).Select');
  Module.CodeModule.InsertLines(N+12,'       Selection.Font.ColorIndex = 1');
  Module.CodeModule.InsertLines(N+13,'       Selection.Interior.ColorIndex = 0');
  Module.CodeModule.InsertLines(N+14,'       Exit For');
  Module.CodeModule.InsertLines(N+15,'    End If');
  Module.CodeModule.InsertLines(N+16,'   End If');
  Module.CodeModule.InsertLines(N+17,' Next I');
  Module.CodeModule.InsertLines(N+18,'''');
  Module.CodeModule.InsertLines(N+19,' Range(Sauve_Adresse).Select');
  Module.CodeModule.InsertLines(N+20,'''');
  Module.CodeModule.InsertLines(N+21,' ActiveCell.Name = Application.ActiveWorkbook.ActiveSheet.Name & "!PGI_RESULTAT"');
  Module.CodeModule.InsertLines(N+22,'''');
  Module.CodeModule.InsertLines(N+23,'Selection.Font.ColorIndex = 3');
  Module.CodeModule.InsertLines(N+24,'Selection.Interior.ColorIndex = 6');
  Module.CodeModule.InsertLines(N+25,'''');
  Module.CodeModule.InsertLines(N+26,'Application.Activeworkbook.Save');
  Module.CodeModule.InsertLines(N+27,'Application.Quit');
  Module.CodeModule.InsertLines(N+28,'End Sub');

end;

{***********A.G.L.***********************************************
Auteur  ...... : F.Vautrain
Créé le ...... : 25/07/2016
Modifié le ... :   /  /
Description .. : Partie servant à gérer les métrés dans un document
Suite ........ : c'est object est dépendant d'un certains nombres de
Suite ........ : paramètres et de critères.
Mots clefs ... :
*****************************************************************}

constructor TMetreArt.Createdoc ;
begin

  OkMetre       := False;
  OKDuplication := False;
  OKAValider    := False;
  OkMetreBiblio := False;
  SAfNatAffaire := '';
  SAfNatProposition := '';


  OkExcel       := GetParamSocSecur ('SO_METRESEXCEL', False);
  OKMetreDoc    := GetParamSocSecur ('SO_BTMETREDOC', False);
  //
  SAfNatAffaire     := GetParamSoc  ('SO_AFNATAFFAIRE');
  SAfNatProposition := GetParamSoc  ('SO_AFNATPROPOSITION');

  if OKMetreDoc then
    fRepMetre   := RecupRepertbase
  else
    fRepMetre   := '';

  FNomMachine   := ComputerName ;
  
  TTOBMetres    := TOB.Create('LES METRES DOC',nil,-1);
  TTobVarDoc    := TOB.Create('LES VARIABLES DOC',nil,-1);
  TTobSaisieVar := TOB.Create('SAISIE VARIABLES', nil, -1);
  TTobVariables := TOB.Create('LES VARIABLES',nil,-1);

end;

function TMetreArt.ControleMetreDoc(CleDoc : string) : Boolean;
Var RepSauve  : string;
    NatPiece  : String;
begin

  Result  := False;

  NatPiece := copy(Cledoc,1,Pos('-',Cledoc)-1);

  AchatVente    := GetInfoParPiece(NatPiece, 'GPP_VENTEACHAT');

  if (OkMetreDoc) And (not OkExcel) then Exit;

  //FV1 : 06/01/2017 - FS#2325 - Divers pbs liés aux métrés en saisie de devis
  if not AutorisationMetre(NatPiece)  then exit;

  if not OfficeExcelDispo then
  begin
    PGIBox('Office n''est pas installé sur ce poste', 'Gestion des métrés');
    OkExcel := False;
    exit;
  end;

  if fRepMetre = '' then
  begin
    PGIBox('le répertoire métrés Serveur est à blanc dans les paramètres société', 'Gestion des métrés');
    OkExcel := False;
    Exit;
  end;

  fRepMetreDoc := fRepMetre;

  if FNomMachine = '' then
  begin
    PGIBox('le Nom de la machine est à blanc', 'Gestion des métrés');
    Exit;
  end;

  //Contrôle existence répertoire métrés article sur serveur
  fRepMetre := fRepMetre + '\' + 'Bibliotheque' + '\';
  OkExcel := CreationRepertoire(fRepMetre);
  if Not OkExcel then Exit;

  //Contrôle répertoire des fichiers métrés document sur Serveur
  fRepMetreDoc := fRepMetreDoc + '\' + 'Documents' + '\';
  OkExcel := CreationRepertoire(fRepMetreDoc);
  if OkExcel then
  begin
    if Action <> TaCreat then
    begin
      fRepMetreDoc := fRepMetreDoc + CleDoc +'\';
      OkExcel := CreationRepertoire(fRepMetreDoc);
      if not OkExcel then Exit;
    end;
  end
  else Exit;

  //création du répertoire de sauvegarde des fichiers métrés Document sur Poste Client
  RepSauve := '';

  //Contrôle répertoire de stockage sur poste client
  fRepMetreDocLocal := '';

  fRepMetreDocLocal := 'C:\PGI01\';
  OkExcel := CreationRepertoire(fRepMetreDocLocal);

  fRepMetreDocLocal := fRepMetreDocLocal + 'Metres\';
  OkExcel := CreationRepertoire(fRepMetreDocLocal);

  fRepMetreDocLocal := fRepMetreDocLocal + 'Documents\';
  OkExcel := CreationRepertoire(fRepMetreDocLocal);

  //Pour une installation TSE on est obligé de gérer le nom de la machine
  fRepMetreDocLocal := fRepMetreDocLocal + FNomMachine + '\';
  OkExcel := CreationRepertoire(fRepMetreDocLocal);

  if action = TaCreat then
    fRepMetreDocLocal := fRepMetreDocLocal + 'Nouveau ' + CleDoc +'\'
  else
    fRepMetreDocLocal := fRepMetreDocLocal + CleDoc +'\';

  OkExcel := CreationRepertoire(fRepMetreDocLocal);

  //On charge un fichier vide...
  fFichierVide := 'c:\PGI00\STD\ARTVIDE.xlsx';
  if not FileExists(fFichierVide) then
  Begin
    PGIBox('le fichier ' + fFichierVide + ' n''existe pas.', 'Gestion des métrés');
    OkExcel := False;
    exit;
  end;

  Result  := True;

end;

function TmetreArt.CreationRepertoire(NomRep : String) : Boolean;
begin

  Result := true;

  if not DirectoryExists(NomRep) then
  begin
    if not CreationDir(NomRep) then
    begin
      PGIBox('le répertoire ' + NomRep + ' n''existe pas.', 'Gestion des métrés');
      Result := False;
      Exit;
    end;
  end;

end;

Procedure TMetreArt.SupprimeRepertoireDoc;
Var Rec : TSearchRec;
begin

  //On charge le répertoire document
  if FindFirst (fRepMetreDocLocal, faAnyFile, Rec) = 0 then
  begin
    if (rec.name <> '.') and (rec.name <> '..') then
    begin
      FindClose (Rec);
      exit;
    end;
    while FindNext (Rec) = 0 do
    begin
      if (rec.name <> '.') and (rec.name <> '..') then
      begin
        FindClose (Rec);
        exit;
      end;
    end;
  end;

  FindClose (Rec);

  // Suppression repertoire Ouvrage
  DeleteRepertoire(fRepMetreDocLocal);

  LibereTob;

end;

//Contrôle si la nature de piéce peut gérer ou non les métrés
function TMetreArt.AutorisationMetre(NatPiece : string) : Boolean;
Begin

	result := False;

  // LS -- PROTECTION -- FS#2319 - SES : Message parasite en saisie de CBT.
  if not OKMetreDoc then exit;
  // ------
  //if not ISFromExcel(TTobPiece) then Exit;

  if AchatVente <> 'VEN' then Exit;

  if      NatPiece = SAfNatAffaire      then result := True
  else if NatPiece = SAfNatProposition  then result := True
  else if Pos(NatPiece,'FBT;FBP') > 0   then result := True
  else if NatPiece =   'AFF'            then result := True
  else if NatPiece =   'DAP'            then result := True;
  
end;


//-- On charge l'ensemble des variables associées au document
//-- Variables Application ==> lecture directe du BVA pour récupération des variables de type 'A'
//-- Variables Document
//-- Variables Article
//-- Variables Sous-Détail
procedure TMetreArt.ChargementVariablesDoc(CleDoc : R_CleDoc);
Var QQ        : TQuery;
    StSql     : String;
begin

  if not OkMetre then Exit;

  //Creation des TOB Virtuelles Variable Application, Variables Document et Variables déjà renseignées
  if Assigned(TTobVarDoc) then
    TTobVarDoc.ClearDetail
  else
    TTobVarDoc    := TOB.Create('LES VARIABLES DOC',nil,-1);

  //Lecture de la table des variables Saisie (BTVARDOC)
  StSQL := 'SELECT * FROM BVARDOC WHERE BVD_NATUREPIECE = "' + CleDoc.NaturePiece;
  StSQL := StSQL + '" AND BVD_SOUCHE="' + CleDoc.Souche;
  StSQL := StSQL + '" AND BVD_NUMERO= ' + IntToStr(CleDoc.NumeroPiece);
  StSQL := StSQL + '  AND BVD_INDICE= ' + IntToStr(CleDoc.Indice);

  //Chargement de la TOB
  QQ := OpenSQL(StSQL,False,-1, '', True);
  If not QQ.Eof Then TTobVarDoc.LoadDetailDB('BVARDOC', '', '', QQ, true);

  Ferme (QQ);

end;

procedure TMetreArt.ChargementMetre(CleDoc : R_CleDoc);
Var QQ        : TQuery;
    ind       : Integer;
    StSql     : String;
    Nomfichier: String;
    RepServeur: String;
    RepMetre  : String;
    FicMetre  : String;
begin

  if not OkMetre then Exit;

  //Creation des TOB Virtuelles Variable Application, Variables Document et Variables déjà renseignées
  if Assigned(TTOBMetres) then
    TTOBMetres.ClearDetail
  else
    TTOBMetres:= TOB.Create('LES METRES',nil,-1);

  //Lecture de la table des variables Saisie (BTVARDOC)
  StSQL := 'SELECT * FROM BMETRE WHERE BME_NATUREPIECEG = "' + CleDoc.NaturePiece;
  StSQL := StSQL + '" AND BME_SOUCHE ="' + CleDoc.Souche;
  StSQL := StSQL + '" AND BME_NUMERO = ' + IntToStr(CleDoc.NumeroPiece);
  StSQL := StSQL + '  AND BME_INDICEG= ' + IntToStr(CleDoc.Indice);

  //Chargement de la TOB
  QQ := OpenSQL(StSQL,False,-1, '', True);
  If not QQ.Eof Then TTOBMetres.LoadDetailDB('BMETRE', '', '', QQ, true);

  //On vérifie si le chemin dans Nom Fichier est bien celui du serveur...
  For ind := 0 to TTobMetres.detail.count - 1 do
  begin
    Nomfichier := TTobmetres.detail[Ind].GetValue('BME_NOMDUFICHIER');
    //On vérifie si le fichier existe sur le répertoire serveur - si oui on le copie du serveur sur le local...
    RepMetre    := ExtractFilePath(Nomfichier);
    RepServeur  := Frepmetredoc;
    FicMetre    := ExtractFileName(Nomfichier);

    //Si pas de fichier métré sur ouvrage
    if RepMetre = '' then continue;

    //si le répertoire dans le param soc n'existe pas ou plus on bloc les métrés XLS
    if not DirectoryExists(RepServeur) then
    begin
      OkExcel := False;
      Continue;
    end;

    //Il faut contrôler que Repertoire serveur = répertoire Paramsoc
    if not (RepMetre = RepServeur) then
    begin
      Repmetre   := IncludeTrailingBackslash(RepMetre) + FicMetre;
      RepServeur := IncludeTrailingBackslash(RepServeur) + FicMetre;
      ///on contrôle si il existe pour le copier du serveur vers le local
      if FileExists(RepMetre) then
      begin
        CopieFichier(RepMetre, RepServeur);
        TTobmetres.detail[Ind].PutValue('BME_NOMDUFICHIER', RepServeur);
      end
      else
      Begin
        If FileExists(RepServeur) then
          TTobmetres.detail[Ind].PutValue('BME_NOMDUFICHIER', RepServeur)
        else
          TTobmetres.detail[Ind].PutValue('BME_NOMDUFICHIER', '');
      end;
    end;
  end;

  Ferme (QQ);

end;

function TMetreArt.MetreAttached(TOBL: TOB; NumOrdre : Integer): boolean;
Var TOBTemp : TOB;
    Prefixe : string;
    NatPiece: string;
    Souche  : string;
    Numero  : Integer;
    Indice  : Integer;
    FileName: string;
begin

  result := False;

  if not OkMetre then Exit;

  Prefixe   := GetPrefixeTable(TOBL);
  //
  NatPiece  := TOBL.GetSTRING(Prefixe + 'NATUREPIECEG');
  Souche    := TOBL.GetSTRING(Prefixe + 'SOUCHE');
  Numero    := TOBL.GetValue(Prefixe  + 'NUMERO');
  Indice    := TOBL.GetValue(Prefixe  + 'INDICEG');
  //
  UniqueBlo := 0;
  //
  if prefixe = 'BLO_' then UniqueBlo := TOBL.GetValue(Prefixe + 'UNIQUEBLO');
  //
  TOBTemp := TTOBMetres.FindFirst(['BME_NATUREPIECEG','BME_SOUCHE','BME_NUMERO','BME_INDICEG','BME_NUMORDRE','BME_UNIQUEBLO'],[NatPiece,Souche,Numero,Indice,NumOrdre,Uniqueblo],False);
  //
  if TobTemp <> nil then
  begin
    Filename := TobTemp.getValue('BME_NOMDUFICHIER');
    if FileName = '' then exit;
    if FileExists(FileName) then Result := True;
  end;

end;

Function TMetreArt.GereMetreXLS(TobPiece, TOBL, TobOuvrage, TobEntOuv : TOB; Okclick : Boolean = False) : Double;
Var Prefixe       : string;
    IndiceNomen   : Integer;
    CleDoc        : R_Cledoc;
    SauveQte      : Double;
begin

  try

    IndexFile := 0;

    IndiceNomen := 0;

    OkClick_Elips := Okclick;

    Prefixe := GetPrefixeTable(TOBL);

    TTobPiece     := TOB.Create('LES PIECES', nil, -1);

    TTOBpiece.Dupliquer(TobPiece, true, True);

    TTObOuvrage   := TOB.Create('LES OUVRAGES', nil, -1);
    //
    ArticleXLS    := TOBL.GetString(Prefixe + 'ARTICLE');
    SauveQte      := TOBL.GetDouble(Prefixe + 'QTEFACT');

    //Si nous sommes sur une ligne de sous-détail on charge l'élément principale (Entete) de l'ouvrage
    if IsSousDetail(TOBL) then TobEntOuv := FindLigneOuvrageinDoc(TOBL);
    //
    if TobEntOuv = nil then TobEntOuv := TOBL;
    //
    CleDoc.NaturePiece  := TobEntOuv.GetString('GL_NATUREPIECEG');
    CleDoc.DatePiece    := V_PGI.DateEntree;
    CleDoc.Souche       := TobEntOuv.GetString('GL_SOUCHE');
    CleDoc.NumeroPiece  := TobEntOuv.GetValue('GL_NUMERO');
    CleDoc.Indice       := TobEntOuv.GetValue('GL_INDICEG');
    CleDoc.NumOrdre     := TobEntOuv.GetValue('GL_NUMORDRE');
    CleDoc.UniqueBlo    := TobEntOuv.GetValue('UNIQUEBLO');
    //
    if GetPrefixeTable(TOBL)='BLO_' THEN CleDoc.Uniqueblo := TOBL.GetValue('BLO_UNIQUEBLO');
    //
    if OkClick_Elips then
      TraitementMetreModif(Cledoc, TobEntOuv, tobOuvrage, Tobl)
    else
      TraitementMetre(Cledoc, TobEntOuv, TobOuvrage, TOBL);

    //on retri sur l'index file pour avoir toujpours le même ordre de saisie en création comme en modif
    TTobSaisieVar.Detail.Sort('BVD_INDEXFILE');

    //Saisie de l'ensemble des variable de la ligne
    if not Saisievariables(Cledoc) then
      CloseXLS
    else
    begin
      SourisSablier;
      //Si nom Tobl est Ligne Ouvrage classique
      If TOBL.NomTable = 'LIGNE' then
      begin
        if TOBL.FieldExists('GL_INDICENOMEN') then IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
      end
      else if TOBL.NomTable = 'LIGNEOUV' then IndiceNomen := Tobl.detail.count;
      //
      if (IndiceNomen <> 0) then //Je suis sur une ligne Ouvrage
      begin
        //On charge la TOB provisoire avec les sous-détail de l'ouvrage dans le cas de la modif des valeurs de l'ouvrage car non fait...
        if OkClick_Elips then  RecalculSousDetailOuvrage(TobEntOuv, TOBOuvrage);

        CalculMetreDocXLsOuvrage(TOBL); //Lancement du calcul de la feuille Excel pour un ouvrage prise en compte de l'ensemble des sous-détail....
      end
      else
        CalculMetreDocXLsART(TTOBOuvrage.detail[0]);    //Lancement du calcul de la feuille Excel Pour un Article simple
      SourisNormale;
    end;
    SourisSablier;

    //chargement des valeurs sur l'ensemble des lignes ouvrage...
    ChargeQteDetailMetre(TobOuvrage, TOBL);

    //On vérifie si la quantité calculée est différente de zétro ou de 1 en fonction des paramètres société....
    Result := TOBL.GetDouble(Prefixe + 'QTEFACT');
    //
    if (SauveQte <> 0) and ((Result = 0) OR (Result = 1)) then
    Begin
      Result := SauveQte;
      Result := Arrondi(Result,V_PGI.OkDecQ);
      TOBL.putValue(Prefixe + 'QTEFACT',  FloatToStr(Result));       //StrF00(result, V_PGI.OkDecQ)
      TOBL.putValue(Prefixe + 'QTESTOCK', FloatToStr(Result));       //StrF00(result, V_PGI.OkDecQ)
      TOBL.putValue(Prefixe + 'QTERESTE', FloatToStr(Result));       //StrF00(result, V_PGI.OkDecQ)
    end;

    //Chargement de la valeur sur la ligne si article simple en fonction de numordre et uniqueblo...
    {*
    if TOBL <> nil then
    begin
      if TOBL.NomTable = 'LIGNE' then
        Result  := Arrondi(TOBL.GetValue('GL_QTEFACT'),V_PGI.OkDecQ)
      else
        Result  := Arrondi(TOBL.GetValue('BLO_QTEFACT'),V_PGI.OkDecQ)
    end
    else
      Result    := Arrondi(TTobOuvrage.detail[0].GetValue('QTE'), V_PGI.OkDecQ);

    //Permettre les Qté à zéro quand on vient des métrés
    if not GetParamSocSecur('SO_QTEAZERO', False) then
    begin
      if Result = 0 then Result := 1;
    end;
    //
    *}
    ArticleXLS := '';
  finally
    FreeAndNil(TTobPiece);
    SourisNormale;
  end;

end;

Procedure TMetreArt.RecalculsousDetailOuvrage(TobEntOuv, TOBOuvrage : TOB);
Var FichierLocal    : String;
    FichierServeur  : String;
    FichierXLS      : String;
    ToblOuvrage     : Tob;
    Ind             : Integer;
begin

  TraitementMetreOuvrage(TobEntOuv, TOBOuvrage);

  For Ind := TTObOuvrage.Detail.count -1 downto 0 do
  begin
    ToblOuvrage := TTobOuvrage.detail[ind];
    //
    FichierXLS  := ToblOuvrage.GetValue('FICHIERXLS');
    if fichierXLS <> '' Then
    BEGIN
      //On vérifie si le fichier existe sur le répertoire local
      FichierLocal := IncludeTrailingBackslash(fRepMetreDocLocal) + ExtractFileName(FichierXLS);

      //S'il n'existe pas on vérifie s'il existe sur le répertoire serveur
      if Not fileExists(FichierLocal) then
      begin
        FichierServeur := IncludeTrailingBackslash(fRepMetreDoc) + ExtractFileName(FichierXLS);
        //S'in n'existe pas on ne fait rien
        If not FileExists(FichierServeur) then
          FichierLocal := ''
        else //Il existe sur le serveur
          CopieFichier(FichierServeur, FichierLocal);
      end;
      ToblOuvrage.PutValue('FICHIERXLS', FichierLocal);
    end;
    //
    //Chargement des variables de l'article
    ChargementVariablesArt(ToblOuvrage.GetValue('ARTICLE'));

    //On charge le fichier des variables documents avec le fichier des variables article
    ChargementTobVarDocwithtobVarArt(TobEntOuv, ToblOuvrage);
  end;

end;

procedure TMetreArt.ChargeQteDetailMetre (TObOuvrage, TOBL : TOB);
var Indice : integer;
    TOBLOuvrage : TOB;
    TOBDetail   : TOB;
    ValRet : variant;
begin

  //if Tobouvrage = nil then exit;

  for Indice := 0 to TTObOuvrage.detail.count -1 do
  begin
    //
    TOBLOuvrage := TTObOuvrage.detail[Indice];

    if TOBLOuvrage.detail.count > 0 then ChargeQteDetailMetre (TOBLOuvrage, TOBL);

    //On cherche dans le sous détail les lignes correspondantes aux lignes traitées dans le métrés
    if Tobouvrage <> nil then
      TobDetail := Tobouvrage.FindFirst(['BLO_UNIQUEBLO'],[TOBLOuvrage.GetValue('UNIQUEBLO')], True)
    else
      TobDetail := nil;

    If TobDetail = nil then
    begin
      if TOBL.NomTable = 'LIGNE' then
      begin
        if (TOBLOuvrage.getValue('UNIQUEBLO')=TOBL.GetValue('UNIQUEBLO')) then
          TobDetail := TOBL
        else
          continue;
      end
      else
      begin
        if (TOBLOuvrage.getValue('UNIQUEBLO')=TOBL.GetValue('BLO_UNIQUEBLO')) then
          TobDetail := TOBL
        else
          continue;
      end;
    end;

    ValRet := TOBLOuvrage.GetValue('QTE');

    if not GetParamSocSecur('SO_QTEAZERO', False) then
    begin
      if ValRet = 0 then ValRet := 1;
    end;

    if TobDetail.NomTable = 'LIGNE' then
    begin
       TobDetail.Putvalue('GL_QTERESTE',Arrondi(Valret,V_PGI.OkDecQ));
       TobDetail.Putvalue('GL_QTESTOCK',Arrondi(Valret,V_PGI.OkDecQ));
       TobDetail.Putvalue('GL_QTEFACT', Arrondi(Valret,V_PGI.OkDecQ));
    end
    else
    begin
      TobDetail.Putvalue('BLO_QTEFACT', Arrondi(Valret,V_PGI.OkDecQ));
      if TOBdetail.GetString('BLO_TYPEARTICLE')='PRE' then
      begin
        if (Pos(TobDetail.GetValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne)>0) and
           (not IsLigneExternalise(TobDetail)) then
        begin
          TobDetail.putValue('BLO_TPSUNITAIRE',1);
        end else
        begin
          TobDetail.putValue('BLO_TPSUNITAIRE',0);
        end;
        TobDetail.Putvalue('GA_HEURE',TobDetail.Getvalue('BLO_QTEFACT'));
      end;
    end;
  end;

end;

Procedure TMetreArt.TraitementMetre(Cledoc : R_CLEDOC; TobEntOuv, TOBOuvrage, TOBL : Tob);
var Ind           : Integer;
    Cpt           : Integer;
    TOBLArticle   : TOB;
begin
  //
  TTObOuvrage.clearDetail;

  //On charge la TOB provisoire avec l'entête d'ouvrage
  //on charge dans le cas de la création de l'entête d'ouvrage (article ouvrage) en passant deux fois la même TOB
  //afin de pouvoir gérer par le fait que l'on soit sur un ouvrage ou un article simple
  CreateEnregTOBOuvrage(TobEntOuv, TobEntOuv);

  If TTobSaisieVar <> nil then TTobSaisieVar.ClearDetail;

  //on charge à chaque fois les variables Générales et application....
  //Chargement des variables de l'article
  //CleDoc.NaturePiece  := TobEntOuv.GetString('GL_NATUREPIECEG');
  //CleDoc.DatePiece    := V_PGI.DateEntree;
  //CleDoc.Souche       := TobEntOuv.GetString('GL_SOUCHE');
  //CleDoc.NumeroPiece  := TobEntOuv.GetValue('GL_NUMERO');
  //CleDoc.Indice       := TobEntOuv.GetValue('GL_INDICEG');
  //CleDoc.NumOrdre     := TobEntOuv.GetValue('GL_NUMORDRE');
  //CleDoc.UniqueBlo    := TobEntOuv.GetValue('UNIQUEBLO');

  //On charge la TOB provisoire avec les sous-détail de l'ouvrage...
  if TobOuvrage <> nil then TraitementMetreOuvrage(TobEntOuv, TOBOuvrage);

  //détermination de la ligne Article
  if TTObOuvrage.Detail.count = 1 then
  begin
    ToblArticle := TTObOuvrage.detail[0];
    Cpt := 1;
    //Affichage dans la grille de saisie du variable de la section : Article
    If ToblArticle.GetValue('TYPEARTICLE') <> 'OUV' then LigneTitreSaisieVariable(Cledoc, 'Article : ' + ToblArticle.GetValue('ARTICLE'));
  end
  else
  begin
    Cpt := 0;
    if TOBL.NomTable = 'LIGNE' then
      TOBLArticle := TTObOuvrage.FindFirst(['NUMORDRE','UNIQUEBLO','ARTICLE'],[CleDoc.NumOrdre, TOBL.GetValue('UNIQUEBLO'),TOBL.GetValue('GL_ARTICLE')], True)
    else
      TOBLArticle := TTObOuvrage.FindFirst(['NUMORDRE','UNIQUEBLO','ARTICLE'],[CleDoc.NumOrdre, TOBL.GetValue('BLO_UNIQUEBLO'),TOBL.GetValue('BLO_ARTICLE')], True);
  end;

  //
  For Ind := TTObOuvrage.Detail.count -1 downto 0 do
  begin
    //Traitement du métré de l'article ou élément de l'ouvrage
    TraitementMetreArticle(TTOBOuvrage.Detail[Ind]);

    //Chargement des variables de l'article
    ChargementVariablesArt(fArticle);

    if fFileName = '' then TTObOuvrage.Detail[ind].PutValue('FICHIERXLS', '');

    //Affichage dans la grille de saisie du variable de la section : Sous-détail
    if Cpt = 0 then
    begin
      LigneTitreSaisieVariable(Cledoc, 'Sous-détail');
      Cpt := 1;
    end
    //Affichage dans la grille de saisie du variable de la section : Ouvrage
    else if TTOBOuvrage.Detail[ind].GetValue('TYPEARTICLE')='OUV' then
    begin
      LigneTitreSaisieVariable(Cledoc, 'Ouvrage : ' + TTOBOuvrage.Detail[ind].GetValue('ARTICLE'));
      Cpt := 0;
    end;

    //On charge le fichier des variables documents avec le fichier des variables article
    ChargementTobVarDocwithtobVarArt(TOBLArticle, TTObOuvrage.Detail[Ind]);
    //
  end;

  //chargement des variables lignes dans la tob de saisie des variables
  ChargementSaisieVarwithVardoc('L', Cledoc, True, True);

  Cledoc.NumOrdre := 0;

  //chargement des vaeriables Document dans la tob de saisie des variables
  ChargementSaisieVarwithVardoc('D', Cledoc, True, True);

  //Affichage dans la grille de saisie du variable de la section : Variables Générales
  LectureTableVariable('G', Cledoc);
  if TTobVariables.Detail.count > 0 then
  begin
    LigneTitreSaisieVariable(Cledoc, 'Variables Générales', True);
    ChargementTobVarDocwithtobVarAPP(Cledoc, True, False);
  end;

  //Affichage dans la grille de saisie des variables de la section : Variables Application
  LectureTableVariable('A', Cledoc);
  if TTobVariables.Detail.count > 0 then
  begin
    LigneTitreSaisieVariable(Cledoc, 'Variables Application', True);
    ChargementTobVarDocwithtobVarAPP(Cledoc, True, False);
  end;

  Renumerotationtob(TTobSaisieVar);

end;

Procedure TMetreArt.TraitementMetreModif(Cledoc : R_CLEDOC; TobEntOuv, TOBOuvrage, TOBL : Tob);
Var ToblArticle : TOB;
begin
  //
  OkCreation  := False;
  //
  TTObOuvrage.clearDetail;
  //
  //CleDoc.NaturePiece  := TobEntOuv.GetString('GL_NATUREPIECEG');
  //CleDoc.DatePiece    := V_PGI.DateEntree;
  //CleDoc.Souche       := TobEntOuv.GetString('GL_SOUCHE');
  //CleDoc.NumeroPiece  := TobEntOuv.GetValue('GL_NUMERO');
  //CleDoc.Indice       := TobEntOuv.GetValue('GL_INDICEG');
  //CleDoc.NumOrdre     := TobEntOuv.GetValue('GL_NUMORDRE');
  //
  //if GetPrefixeTable(TOBL)='BLO_' THEN
  //begin
  //  CleDoc.Uniqueblo := TOBL.GetValue('BLO_UNIQUEBLO');
  //end
  //Else
  //  CleDoc.Uniqueblo := TOBL.GetValue('UNIQUEBLO');
  //
  CreateEnregTOBOuvrage(Tobl, TobEntOuv);

  If TTobSaisieVar <> nil then TTobSaisieVar.ClearDetail;

  ToblArticle := TTObOuvrage.detail[0];

  //Traitement du métré de l'article ou élément de l'ouvrage
  TraitementMetreArticle(ToblArticle);

  if fFileName = '' then ToblArticle.PutValue('FICHIERXLS', '');

  //Chargement des variables de l'article
  ChargementVariablesArt(fArticle);

  if TTobVariables.detail.count <> 0 then
  begin
    //entête de ligne... (Ss-Detail ou ligne article)
    if GetPrefixeTable(TOBL)='BLO_' then
       LigneTitreSaisieVariable(Cledoc, 'Sous-Détail : ' + ToblArticle.GetValue('ARTICLE'))
    else
    Begin
      if TOBL.GetValue('GL_INDICENOMEN') <> 0 then
        LigneTitreSaisieVariable(Cledoc, 'Ouvrage : ' + ToblArticle.GetValue('ARTICLE'))
      else
        LigneTitreSaisieVariable(Cledoc, 'Article : ' + ToblArticle.GetValue('ARTICLE'));
    end;

    //On charge les variables associées à la ligne mais on les cache
    //pour qu'elles n'apparaissent pas en saisie des variables en modification...
    TraitementVariableCachees(ToblArticle, Cledoc, ToblArticle.GetValue('TYPEARTICLE'), False, False);

    //On charge le fichier des variables documents avec le fichier des variables article
    if TTobVariables.detail.count <> 0 then ChargementTobVarDocwithtobVarArt(TOBLArticle, TOBLArticle);
  end;

  //On vérifie si dans la tob de saisie les variables lignes ne sont pas déjà chargées
  ChargementSaisieVarwithVardoc('L', Cledoc, True, True);

  //Comment gérer l'ouvrage associé au sous-détail quand en modif de Ss-détail ????
  if GetPrefixeTable(TOBL)='BLO_' then
  begin
    LigneTitreSaisieVariable(Cledoc, 'Ouvrage : ' + TobEntOuv.GetValue('GL_ARTICLE'),False);
    If TobOuvrage.FieldExists('BLO_UNIQUEBLO') then
      Cledoc.UniqueBlo  := TobOuvrage.GetValue('BLO_UNIQUEBLO')
    else
      Cledoc.UniqueBlo  := 0;
    TraitementVariableCachees(TobEntOuv, Cledoc, 'OUV', False, false);
  end;

  Cledoc.NumOrdre   := 0;

  //chargement des variables Document dans la tob de saisie des variables
  ChargementSaisieVarwithVardoc('D', Cledoc, False, false);

  RenumerotationTob(TTobSaisieVar);

  TTobSaisieVar.Detail.Sort('BVD_INDEXFILE');

  //On charge les variables associées à la ligne mais on les cache
  //pour qu'elles n'apparaissent pas en saisie des variables en modification...
  //Affichage dans la grille de saisie du variable de la section : Variables Générales
  LectureTableVariable('G', Cledoc);
  if TTobVariables.Detail.count > 0 then
  begin
    LigneTitreSaisieVariable(Cledoc, 'Variables Générales', False);
    ChargementTobVarDocwithtobVarAPP(Cledoc, False, False);
  end;

  //Affichage dans la grille de saisie des variables de la section : Variables Application
  LectureTableVariable('A', Cledoc);
  if TTobVariables.Detail.count > 0 then
  begin
    LigneTitreSaisieVariable(Cledoc, 'Variables Application',False);
    ChargementTobVarDocwithtobVarAPP(Cledoc, False, False);
  end;
  //
  RenumerotationTob(TTobSaisieVar);

end;

procedure TMetreArt.RenumerotationTob(TobAtraiter : Tob);
Var ind : Integer;
begin

  Indexfile := 0;

  for Ind := 0 to TobAtraiter.detail.count -1 do
  begin
    Inc(Indexfile);
    TobAtraiter.Detail[Ind].PutValue('BVD_INDEXFILE', IndexFile);
  end;

end;

Procedure TMetreArt.TraitementVariableCachees(ToblArticle : TOB; CleDoc : R_Cledoc; TypeArticle : string; OkAffiche : Boolean=True; OKSaisie : Boolean=true);
Var ToblVardoc    : TOB;
    ToblSaisieVar : TOB;
    TobLtempo     : TOB;
    CodeArt       : String;
begin

  CodeArt := CleDoc.NaturePiece + '-' + Cledoc.Souche+ '-' + IntToStr(Cledoc.NumeroPiece) + '-' + intToStr(Cledoc.Indice) + '-' + inttoStr(Cledoc.NumOrdre) + '-' + IntToStr(Cledoc.UniqueBlo);

  //chargement de l'ensemble des variables propres à la ligne en cours
  //If TypeArticle = 'OUV' then
  //  ToblVardoc := TTobVardoc.FindFirst(['BVD_NUMORDRE'],[Cledoc.NumOrdre],False)
  //else
  ToblVardoc := TTobVardoc.FindFirst(['BVD_NUMORDRE','BVD_UNIQUEBLO'],[Cledoc.NumOrdre,Cledoc.UniqueBlo],False);
                 
  while ToblVardoc <> nil do
  Begin
    //on se trouve sur une ligne variable
    if ToblVardoc.GetValue('BVD_ARTICLE') <> '' then
    Begin
      if ToblVardoc.GetValue('BVD_ARTICLE') <> CodeArt then
      begin
        //On contrôle si le libellé de lVarDoc n'existe pas déjà dans la saisie des variables...
        TobLtempo := TTobSaisieVar.FindFirst(['BVD_NUMORDRE','BVD_UNIQUEBLO','BVD_LIBELLE'], [Cledoc.NumOrdre,Cledoc.UniqueBlo, ToblVardoc.GetValue('BVD_LIBELLE')], false);
        if ToblTempo = nil then
        begin
          // Duplication de l'enreg de TOBVardoc (Variables document) dans TobSaisieVar (Saisie Variables)
          ToblSaisieVar := Tob.Create('BVARDOC', TTobSaisieVar, -1);
          ToblSaisieVar.Dupliquer(TobLVarDoc, True, true);
          if OkAffiche then
            ToblSaisieVar.AddChampSupValeur('AFFICHABLE', 'X') //On définit qu'il ne faut pas afficher ces rubriques...
          else
            ToblSaisieVar.AddChampSupValeur('AFFICHABLE', '-');
          if OkSaisie then
            ToblSaisieVar.AddChampSupValeur('SAISISSABLE', 'X') //On définit qu'il ne faut pas Saisir ces rubriques...
          else
            ToblSaisieVar.AddChampSupValeur('SAISISSABLE', '-');
        end;
      end;
    end
    else
    Begin
      //on se trouve sur une ligne Titre
      //If TypeArticle = 'OUV' then
      //  TobLtempo := TTobSaisieVar.FindFirst(['BVD_NUMORDRE','BVD_LIBELLE'], [Cledoc.NumOrdre, ToblVardoc.GetValue('BVD_LIBELLE')], false)
      //else
      TobLtempo := TTobSaisieVar.FindFirst(['BVD_NUMORDRE','BVD_UNIQUEBLO','BVD_LIBELLE'], [Cledoc.NumOrdre,Cledoc.UniqueBlo, ToblVardoc.GetValue('BVD_LIBELLE')], false);
      if ToblTempo = nil then
      begin
        // Duplication de l'enreg de TOBVardoc (Variables document) dans TobSaisieVar (Saisie Variables)
        ToblSaisieVar := Tob.Create('BVARDOC', TTobSaisieVar, -1);
        ToblSaisieVar.Dupliquer(TobLVarDoc, True, true);
        if OkAffiche then
          ToblSaisieVar.AddChampSupValeur('AFFICHABLE', 'X') //On définit qu'il ne faut pas afficher ces rubriques...
        else
          ToblSaisieVar.AddChampSupValeur('AFFICHABLE', '-');
        if OkSaisie then
          ToblSaisieVar.AddChampSupValeur('SAISISSABLE', 'X') //On définit qu'il ne faut pas Saisir ces rubriques...
        else
          ToblSaisieVar.AddChampSupValeur('SAISISSABLE', '-');
      end;
    end;
    //
    //If ToblArticle.GetValue('TYPEARTICLE') = 'OUV' then
    //  ToblVardoc := TTobVardoc.FindNext(['BVD_NUMORDRE'],[Cledoc.NumOrdre],False)
    //else
      ToblVardoc := TTobVardoc.FindNext(['BVD_NUMORDRE','BVD_UNIQUEBLO'],[Cledoc.NumOrdre,Cledoc.UniqueBlo],False);
  end;
  //

end;

Procedure TMetreArt.TraitementMetreArticle(TOBL : TOB);
Var TOBLMETRE   : TOB;
    CodeArt     : string;
    FichierArt  : string;
begin
  //
  fFileName := '';
  //
  UniqueBlo := TOBL.GetValue('UNIQUEBLO');
  //
  fTypeArt  := TOBL.GetValue('TYPEARTICLE');
  FArticle  := TOBL.GetValue('ARTICLE');
  //
  //On Recherche le fichier xlsx associé à la ligne en cours...
  //If UniqueBlo = 0 then
  //  TOBLMETRE := TTOBMetres.FindFirst(['BME_NATUREPIECEG','BME_SOUCHE','BME_NUMERO','BME_INDICEG','BME_NUMORDRE'], [TOBL.GetString('NATUREPIECE'),TOBL.GetString('SOUCHE'),TOBL.GetValue('NUMERO'),TOBL.GetValue('INDICE'),TOBL.GetValue('NUMORDRE')],False)
  //else
  //
  TOBLMETRE := TTOBMetres.FindFirst(['BME_NATUREPIECEG','BME_SOUCHE','BME_NUMERO','BME_INDICEG','BME_NUMORDRE', 'BME_UNIQUEBLO'], [TOBL.GetString('NATUREPIECE'),TOBL.GetString('SOUCHE'),TOBL.GetValue('NUMERO'),TOBL.GetValue('INDICE'),TOBL.GetValue('NUMORDRE'),TOBL.GetValue('UNIQUEBLO')],False);

  If TOBLMETRE = nil then
  begin
    //L'enregistrement n'existe pas il faut donc le créer au niveau de la tob des métrés
    TOBLMETRE := TOB.Create('BMETRE', TTOBMetres, -1);
    TOBLMETRE.PutValue('BME_NATUREPIECEG',  TOBL.GetString('NATUREPIECE'));
    TOBLMETRE.PutValue('BME_SOUCHE',        TOBL.GetString('SOUCHE'));
    TOBLMETRE.PutValue('BME_NUMERO',        TOBL.GetValue('NUMERO'));
    TOBLMETRE.PutValue('BME_INDICEG',       TOBL.GetValue('INDICE'));
    TOBLMETRE.PutValue('BME_NUMORDRE',      TOBL.GetValue('NUMORDRE'));
    TOBLMETRE.PutValue('BME_UNIQUEBLO',     TOBL.GetValue('UNIQUEBLO'));
    //
    TOBLMETRE.PutValue('BME_ARTICLE',       TOBL.GetString('ARTICLE'));
    TOBLMETRE.PutValue('BME_NOMDUFICHIER',  '');
    if OkClick_Elips = True Then
    begin
      OkCreation  := False;
      ArticleXLS  := TOBL.GetString('ARTICLE'); //on est en création mais on a cliquer en saisie sur l'elipsis
    end
    else
    begin
      OkCreation  := True;
      ArticleXLS  := ''; //on est en création on remet donc le code à blanc pour ouvrir si Ouvrage l'ensemble des fichiers...
    end;
  end;

  if TOBLMETRE.GetString('BME_NOMDUFICHIER') <> '' then fFileName := IncludeTrailingBackslash (fRepMetreDoc) + ExtractFileName(TOBLMETRE.GetString('BME_NOMDUFICHIER'));

  ControleFichierMetresLigne(TOBL, TOBLMETRE);

end;

Procedure TMetreArt.ControleFichierMetresLigne(TOBL, TOBLMETRE : TOB);
Var CodeArt     : String;
    FichierArt  : String;
begin

  //On formate le code article pour qu'il puisse être un nom de fichier
  CodeArt := FormatArticle(TOBL.GetValue('ARTICLE'));

  //si le filename est à blanc cela signifie que cette ligne n'a pas de fichier xlsx associées il faut donc le créer
  if fFileName = '' then
  begin
    //Création du nom du fichier Article XLS sur le Poste de travail
    //On crée le fichier métré document en ajoutant le numéro d'ordre (ou de ligne ???) et le uniqueblo si sur sous-détail d'ouvrage poste client
    FichierArt:= fRepMetreDocLocal + CodeArt + '-' + IntToStr(TOBL.GetValue('NUMORDRE')) + '-' + IntToStr(TOBL.GetValue('UNIQUEBLO')) + '.xlsx';
    //On charge le nom du fichier article sur le répertoire bibliothèque du serveur
    fFileName := fRepMetre + CodeArt + '.xlsx';
    //On contrôle si le fichier métré de l'Article existe sur le serveur
    if FileExists(fFileName) then
    begin
      CopieFichier(fFileName, FichierArt); //on le copie sur le poste client
    end
    else
    //si nous n'avons aucun fichier qui correspond à l'article ou au sous-détail on ne fait rien....
    begin
      if OkClick_Elips then
      Begin
        CopieFichier(fFichierVide, FichierArt);  //On copie le fichier vide sur le poste client uniquement en modif si celui-ci n'existe pas...
      end
      else
      begin
        fFileName := '';
        TOBL.PutValue('FICHIERXLS', '');
        TOBLMETRE.PutValue('BME_NOMDUFICHIER', '');
        Exit;
      end;
    end;
  end
  else
  //Cela signifie qu'il existe en principe dans le répertoire du serveur un fichier pour cet article sur ce document à cette ligne
  begin
    FichierArt  := ExtractFileName(fFileName);
    FichierArt  := fRepMetreDocLocal + FichierArt;
    //On vérifie si le fichier existe dans le répertoire local
    if not FileExists(fichierArt) then
    begin
      //si le fichier se trouvant sur le serveur existe
      if FileExists(fFileName) then
      Begin
        CopieFichier(fFileName, FichierArt); //on le copie sur le poste client
      end
      else
      Begin
        CopieFichier(fFichierVide, FichierArt); //sinon on copie le fichier vide sur le poste client
      end;
    end;
  end;

  //On charge le nom du fichier dans le TOBL
  TOBL.PutValue('FICHIERXLS', FichierArt);

  TOBLMETRE.PutValue('BME_NOMDUFICHIER', FichierArt);

end;

Procedure TMetreArt.LigneTitreSaisieVariable(Cledoc : R_CLEDOC; Libelle : string; OkAffiche :Boolean=True);
Var ToblSaisieVar   : TOB;
begin

    ToblSaisieVar  := TTobSaisieVar.Findfirst(['BVD_LIBELLE'],[Libelle], true);
    if ToblSaisieVar = nil then
    begin
      //
      inc(Indexfile);
      //
      ToblSaisieVar := Tob.Create('BVARDOC',     TTobSaisieVar, -1);
      //
      ToblSaisieVar.Putvalue('BVD_NATUREPIECE',  Cledoc.NaturePiece);
      ToblSaisieVar.Putvalue('BVD_SOUCHE',       Cledoc.Souche);
      ToblSaisieVar.Putvalue('BVD_NUMERO',       Cledoc.NumeroPiece);
      ToblSaisieVar.Putvalue('BVD_INDICE',       Cledoc.Indice);
      ToblSaisieVar.Putvalue('BVD_NUMORDRE',     Cledoc.NumOrdre);
      ToblSaisieVar.Putvalue('BVD_UNIQUEBLO',    0);
      ToblSaisieVar.Putvalue('BVD_INDEXFILE',    Indexfile);
      ToblSaisieVar.Putvalue('BVD_ARTICLE',      '');
      ToblSaisieVar.Putvalue('BVD_CODEVARIABLE', '');
      ToblSaisieVar.Putvalue('BVD_LIBELLE',      Libelle);
      ToblSaisieVar.Putvalue('BVD_VALEUR',       '');
    end;

    if OkAffiche then
      ToblSaisieVar.AddChampSupValeur('AFFICHABLE', 'X')
    else
      ToblSaisieVar.AddChampSupValeur('AFFICHABLE', '-');

    ToblSaisieVar.AddChampSupValeur('SAISISSABLE', '-');

end;

Procedure TMetreArt.ChargementTobVarDocWithTobVarApp(Cledoc : R_CLEDOC; OkAffiche : Boolean=true; OkSaisie : Boolean=True);
Var TObLigneVarArt  : TOB;
    ToblSaisieVar   : TOB;
    ToblVarDoc      : TOB;
    i               : Integer;
Begin

  //chargement de la tob Variables du document avec les enreg de la tob variables application
  //Attention une variable est unique par article de par son code sur une ligne du docuement...
  For i := 0 To TTobVariables.detail.count-1 do
  Begin
    Inc(Indexfile);
    TObLigneVarArt := TTobVariables.detail[i];
    ToblVarDoc  := TTobVarDoc.Findfirst(['BVD_CODEVARIABLE'],[TObLigneVarArt.getValue('BVA_CODEVARIABLE')], true);
    if ToblVarDoc = nil then
    begin
      //
      // On créer l'enregistrement de la table VarDOC tout en préparant la tob de saisie TTobSaisieVar....
      ToblVarDoc := Tob.Create('BVARDOC',     TTobVarDoc, -1);
      ToblVarDoc.Putvalue('BVD_NATUREPIECE',  Cledoc.NaturePiece);
      ToblVarDoc.Putvalue('BVD_SOUCHE',       Cledoc.Souche);
      ToblVarDoc.Putvalue('BVD_NUMERO',       Cledoc.NumeroPiece);
      ToblVarDoc.Putvalue('BVD_INDICE',       Cledoc.Indice);
      ToblVarDoc.Putvalue('BVD_NUMORDRE',     0);
      ToblVarDoc.Putvalue('BVD_UNIQUEBLO',    0);
      ToblVarDoc.Putvalue('BVD_INDEXFILE',    Indexfile);
      ToblVarDoc.Putvalue('BVD_ARTICLE',      '');
      ToblVarDoc.Putvalue('BVD_CODEVARIABLE', TObLigneVarArt.getValue('BVA_CODEVARIABLE'));
      ToblVarDoc.Putvalue('BVD_LIBELLE',      TObLigneVarArt.getValue('BVA_LIBELLE'));
      ToblVarDoc.Putvalue('BVD_VALEUR',       TObLigneVarArt.getValue('BVA_VALEUR'));
    end;
    ToblVarDoc.Putvalue('BVD_INDEXFILE',      Indexfile);
    //
    ToblSaisieVar := Tob.Create('BVARDOC', TTobSaisieVar, -1);
    ToblSaisieVar.Dupliquer(ToblVarDoc, True, true);
    if okaffiche then
      ToblSaisieVar.AddChampSupValeur('AFFICHABLE', 'X')
    else
      ToblSaisieVar.AddChampSupValeur('AFFICHABLE', '-');
    if okSaisie then
      ToblSaisieVar.AddChampSupValeur('SAISISSABLE', 'X')
    else
      ToblSaisieVar.AddChampSupValeur('SAISISSABLE', '-');     
    //
  End;

end;

Procedure TMetreArt.ChargementTobVarDocWithTobVarArt(TOBL, TOBLOuvrage : TOB;OkAffiche : Boolean=True; OkSaisie : Boolean=True);
Var TObLigneVarArt  : TOB;
    TobLSVar        : TOB;
    TobLVarDoc      : TOB;
    ToblSaisieVar   : TOB;
    i               : Integer;
begin

  if TOBL = nil then exit;

  For i := 0 To TTobVariables.detail.count-1 do
  begin
    TObLigneVarArt := TTobVariables.detail[i];
    //On recherche d'abord si dans les Variables à saisir la variable Article existe
    TobLSVar := TTobSaisieVar.Findfirst(['BVD_ARTICLE','BVD_CODEVARIABLE'],[TObLigneVarArt.getValue('BVA_ARTICLE'), TObLigneVarArt.getValue('BVA_CODEVARIABLE')], true);
    //si elle existe c'est pas normal on supprime la ligne...
    //Sauf si on a fait un copié-collé !!!!
    if TobLSVar <> nil then
    begin
      if TobLSVar.getValue('BVD_INDEXFILE') = 999999 then
      else
        FreeandNil(ToblSVAR);
    end;

    //On recherche ensuite si dans les Variables document cette variable existe
    TobLVarDoc := TTobVarDoc.Findfirst(['BVD_ARTICLE','BVD_CODEVARIABLE'],[TObLigneVarArt.getValue('BVA_ARTICLE'), TObLigneVarArt.getValue('BVA_CODEVARIABLE')], true);
    //Si elle n'existe pas on crée un ligne de saisie
    if TobLVarDoc = nil then
      TobLVarDoc := CreationLigneVarDoc(TobLigneVarArt, TOBLOuvrage)
    else
    begin
      //on s'assure que l'article se trouvant en ouvrage est le même que l'article de la ligne sélectionnée
      if TOBLOuvrage.GetValue('ARTICLE') = TOBL.getValue('ARTICLE') Then
      Begin
        //Si elle existe on vérifie si c'est le même N° Ordre et le même numero UniqueBLO que la ligne sélectionnée
        TobLVarDoc  := TTobVarDoc.Findfirst(['BVD_NUMORDRE','BVD_UNIQUEBLO','BVD_ARTICLE','BVD_CODEVARIABLE'],[TObL.getValue('NUMORDRE'),TObL.getValue('UNIQUEBLO'),TObL.getValue('ARTICLE'),TObLigneVarArt.getValue('BVA_CODEVARIABLE')], true);
        //Si elle n'existe pas on crée un ligne de saisie
        if TobLVarDoc = nil then TobLVarDoc := CreationLigneVarDoc(TobLigneVarArt, TOBL);
      end
      else
      begin
        TobLVarDoc  := TTobVarDoc.Findfirst(['BVD_NUMORDRE','BVD_UNIQUEBLO','BVD_ARTICLE','BVD_CODEVARIABLE'],[TOBLOuvrage.getValue('NUMORDRE'),TOBLOuvrage.getValue('UNIQUEBLO'),TOBLOuvrage.getValue('ARTICLE'),TObLigneVarArt.getValue('BVA_CODEVARIABLE')], true);
        if TobLVarDoc = nil then TobLVarDoc := CreationLigneVarDoc(TobLigneVarArt, TOBLOuvrage);
      end;
    end;
    //Si l'enregistrement existe déjà on le duplique pour la saisie
    ToblSaisieVar := Tob.Create('BVARDOC', TTobSaisieVar, -1);
    ToblSaisieVar.Dupliquer(TobLVarDoc, True, true);
    //Si la ligne de saisie provient d'un copier-coller il faut penser à récupérer la valeur copiée
    if (Assigned(ToblsVar)) Or (ToblsVar <> nil) then ToblSaisieVar.PutValue('BVD_VALEUR', ToblsVar.GetValue('BVD_VALEUR'));
    //
    if okAffiche then
      ToblSaisieVar.AddChampSupValeur('AFFICHABLE', 'X')
    else
      ToblSaisieVar.AddChampSupValeur('AFFICHABLE', '-');
    //
    if OkSaisie then
      ToblSaisieVar.AddChampSupValeur('SAISISSABLE', 'X')
    else
      ToblSaisieVar.AddChampSupValeur('SAISISSABLE', '-');
  end;
  
end;

Procedure TMetreArt.ChargementSaisieVarWithVardoc(TypeVar : String; Cledoc : R_CLedoc; OkAffiche :Boolean=True; OkSaisie :Boolean=True);
Var CodeArt : string;
    ToblSaisieVar : TOB;
    TobLVarDoc    : TOB;
    Ind           : Integer;
    Nbpassage     : Integer;
begin

  NbPassage := 1;

  //avant d'envoyer à l'écran de saisie des doc il faut faire le ménage en fonction du type passé
  if TypeVar = 'D' then
    CodeArt := CleDoc.Naturepiece + '-' + Cledoc.Souche + '-' + IntToStr(cledoc.NumeroPiece) + '-' + InttoStr(Cledoc.Indice)
  else if TypeVar = 'L' then
    CodeArt := CleDoc.Naturepiece + '-' + Cledoc.Souche + '-' + IntToStr(cledoc.NumeroPiece) + '-' + InttoStr(Cledoc.Indice) + '-' + InttoStr(Cledoc.NumOrdre) + '-' + InttoStr(Cledoc.UniqueBlo);
       
  For ind := 0 to TTObVarDoc.Detail.count -1 do
  begin
    if TTObVarDoc.Detail[ind].GetValue('BVD_ARTICLE') = CodeArt then
    begin
      ToblVardoc := TTObVarDOC.Detail[Ind];
      if nbPassage = 1 then
      begin
        if TypeVar = 'D' then
          LigneTitreSaisieVariable(Cledoc, 'Variables Document', OkAffiche)
        else if TypeVar = 'L' then
          LigneTitreSaisieVariable(Cledoc, 'Variables Lignes', OkAffiche);
      end;
      NbPassage := 0;
      //Si l'enregistrement existe déjà on le duplique pour la saisie
      ToblSaisieVar := Tob.Create('BVARDOC', TTobSaisieVar, -1);
      ToblSaisieVar.Dupliquer(TobLVarDoc, True, true);
      if OkAffiche then
        ToblSaisieVar.AddChampSupValeur('AFFICHABLE', 'X')
      else
        ToblSaisieVar.AddChampSupValeur('AFFICHABLE', '-');
      if OkSaisie then
        ToblSaisieVar.AddChampSupValeur('SAISISSABLE', 'X')
      else
        ToblSaisieVar.AddChampSupValeur('SAISISSABLE', '-');
    end;
  end;

end;


Function TMetreArt.CreationLigneVarDoc(TobLigneVar, TOBL : TOB) : Tob;
Var TobLVarDoc : TOB;
Begin

  // Chargement de la TOB des Variables document avec les enregistrement de la Tob
  TobLVarDoc := Tob.Create('BVARDOC',     TTobVarDoc, -1);
  //
  TobLVarDoc.Putvalue('BVD_NATUREPIECE',  TOBL.GetString('NATUREPIECE'));
  TobLVarDoc.Putvalue('BVD_SOUCHE',       TOBL.GetString('SOUCHE'));
  TobLVarDoc.Putvalue('BVD_NUMERO',       TOBL.GetValue('NUMERO'));
  TobLVarDoc.Putvalue('BVD_INDICE',       TOBL.GetValue('INDICE'));

  If (TobLigneVar.GetString('BVA_TYPE') = 'B') then
  begin
    TobLVarDoc.Putvalue('BVD_NUMORDRE',   TOBL.GetValue('NUMORDRE'));
    TobLVarDoc.Putvalue('BVD_UNIQUEBLO',  TOBL.GetValue('UNIQUEBLO'));
    TobLVarDoc.Putvalue('BVD_ARTICLE',    TobLigneVar.getValue('BVA_ARTICLE'));
  end
  else
  begin
    TobLVarDoc.Putvalue('BVD_NUMORDRE',   0);
    TobLVarDoc.Putvalue('BVD_UNIQUEBLO',  0);
    TobLVarDoc.Putvalue('BVD_ARTICLE',   '');
  end;

  TobLVarDoc.Putvalue('BVD_CODEVARIABLE', TobLigneVar.getValue('BVA_CODEVARIABLE'));
  TobLVarDoc.Putvalue('BVD_LIBELLE',      TobLigneVar.getValue('BVA_LIBELLE'));
  TobLVarDoc.Putvalue('BVD_VALEUR',       TobLigneVar.getValue('BVA_VALEUR'));

  Inc(Indexfile);
  TobLVarDoc.Putvalue('BVD_INDEXFILE',    Indexfile);

  Result := TobLVarDoc;


end;

Procedure TMetreArt.TraitementMetreOuvrage(TOBL, TOBOuv : Tob);
Var Ind         : Integer;
    TOBLigOuv   : TOB;
begin

  if TOBOuv = nil then exit;

  //On lit l'ensemble des sous-détail de l'ouvrage
  For Ind:= 0 to TOBOuv.Detail.count -1 do
  begin
    TOBLigOuv    := TOBOuv.Detail[Ind];
    if TOBLigOuv.detail.count > 1 Then TraitementMetreOuvrage(TOBL, TOBLigOuv);
    CreateEnregTOBOuvrage(TOBLigOuv, TOBL);
  end;

end;

Procedure TMetreArt.CreateEnregTOBOuvrage(TOBART, TOBL : TOB);
Var Prefixe     : string;
    TOBLOuvrage : TOB;
    TOBLMetre   : TOB;
begin

  if TOBART = nil then exit;

  Prefixe := GetPrefixeTable(TOBART);

  //on vérifie si la ligne n'existe pas déjà dans la tob....
  if Prefixe = 'BLO_' then
    TOBLOuvrage := TTObOuvrage.FindFirst(['NUMORDRE','UNIQUEBLO','ARTICLE'],[TOBL.GetValue ('GL_NUMORDRE'),TOBART.GetValue ('BLO_UNIQUEBLO'),TOBART.GetString(Prefixe + 'ARTICLE')], False)
  else
    TOBLOuvrage := TTObOuvrage.FindFirst(['NUMORDRE','ARTICLE'],[TOBL.GetValue ('GL_NUMORDRE'),TOBART.GetString(Prefixe + 'ARTICLE')], False);
  If TOBLOuvrage <> nil then Exit;

  TOBLOuvrage := TOb.Create('LIG OUVRAGE', TTObOuvrage, -1);

  TOBLOuvrage.AddChampSupValeur('NATUREPIECE',  TOBART.GetString(Prefixe + 'NATUREPIECEG'));
  TOBLOuvrage.AddChampSupValeur('SOUCHE',       TOBART.GetString(Prefixe + 'SOUCHE'));
  TOBLOuvrage.AddChampSupValeur('NUMERO',       TOBART.GetValue (Prefixe + 'NUMERO'));
  TOBLOuvrage.AddChampSupValeur('INDICE',       TOBART.GetValue (Prefixe + 'INDICEG'));
  TOBLOuvrage.AddChampSupValeur('NUMORDRE',     TOBL.GetValue ('GL_NUMORDRE'));
  TOBLOuvrage.AddChampSupValeur('UNIQUEBLO', 0);

  if Prefixe = 'BLO_' then
    TOBLOuvrage.putvalue('UNIQUEBLO',  TOBART.GetValue ('BLO_UNIQUEBLO'));

  TOBLOuvrage.AddChampSupValeur('ARTICLE',      TOBART.GetString(Prefixe + 'ARTICLE'));
  TOBLOuvrage.AddChampSupValeur('TYPEARTICLE',  TOBART.GetString(Prefixe + 'TYPEARTICLE'));
  TOBLOuvrage.AddChampSupValeur('FICHIERXLS',   FormatArticle(TOBART.GetString(Prefixe + 'ARTICLE') + '.xlsx'));
  //TOBLOuvrage.AddChampSupValeur('FICHIERXLS',   '');

  //On cherche dans la Tob TTobmetre si le fichier n'est pas renseigné
  TOBLMetre := TTOBMetres.FindFirst(['BME_NATUREPIECEG','BME_SOUCHE','BME_NUMERO','BME_INDICEG','BME_NUMORDRE','BME_UNIQUEBLO'],
                                    [TOBART.GetString(Prefixe + 'NATUREPIECEG'),TOBART.GetString(Prefixe + 'SOUCHE'),TOBART.GetValue (Prefixe + 'NUMERO'),TOBART.GetValue (Prefixe + 'INDICEG'),TOBL.GetValue ('GL_NUMORDRE'), TOBLOuvrage.GetValue('UNIQUEBLO')],False);
  if TOBLMetre <> nil then TOBLOuvrage.PutValue('FICHIERXLS', TOBLMetre.GetValue('BME_NOMDUFICHIER'));

  TOBLOuvrage.AddChampSupValeur('QTE',          TOBART.GetString(Prefixe + 'QTEFACT'));

  TOBLOuvrage.AddChampSupValeur('OKMODIF',      '-');

end;

//Saisie des variables nécessaires  pour calcul des métrés
function TMetreArt.SaisieVariables(Cledoc : R_Cledoc) : boolean;
Var ToblSaisieVar : TOB;
    ToblVardoc    : TOB;
    ToblMetre     : TOB;
    Ind           : Integer;
begin

  Result := False;

  if TTobSaisieVar.Detail.count = 0 then
  Begin
    if OkClick_Elips then
    Begin
      OkOpenXLS := True;
      Result    := true;
    end;
    Exit;
  end
  //FV1 : 05/01/2017 - FS#2325 - Divers pbs liés aux métrés en saisie de devis
  else
  begin
    //Dans le cas d'un sous-détail sans fichier XLs associé il ne faut pas ouvrir
    //la fiche variables en création de la ligne
    if (not OkClick_Elips) then
    begin
      //On recherche dans TTobmetre si Nom du fichier renseigné... Si vide on ne fait rien
      ToblMetre := TTOBMetres.FindFirst(['BME_NATUREPIECEG','BME_SOUCHE','BME_NUMERO','BME_INDICEG','BME_NUMORDRE','BME_UNIQUEBLO'], [Cledoc.NaturePiece, Cledoc.Souche, Cledoc.NumeroPiece, Cledoc.Indice, Cledoc.NumOrdre, Cledoc.UniqueBLO], False);
      If ToblMetre <> nil then
      begin
        if ToblMetre.GetValue('BME_NOMDUFICHIER') = '' then
        Begin
          OkOpenXLS := False;
          result    := False;
          Exit;
        end;
      end;
    end;
  end;

  //Si aucune ligne n'est affichable on ouvre le fichier XLS
  if TTobSaisieVar.FindFirst(['AFFICHABLE'], ['X'], False) = nil then
  begin
    if OkClick_Elips then
    Begin
      OkOpenXLS := True;
      result    := true;
    end;
    Exit;
  end;

  //On s'assure d'avoir au moins un ligne saisissable
  if TTobSaisieVar.FindFirst(['SAISISSABLE'], ['X'], False) <> nil then
  begin
    TTobSaisieVar.AddChampSupValeur ('RESULT','-',false);
    if OkClick_Elips then
      TTobSaisieVar.AddChampSupValeur ('VALIDE','V',false)
    else
      TTobSaisieVar.AddChampSupValeur ('VALIDE','-',false);
    //
    TheTOb := TTobSaisieVar;
    //
    TRY
      AGLLanceFiche('BTP', 'BTSAISIEVAR','','','');
      if (TheTOB <> nil) and (TheTOB.getValue('RESULT')='X') then
      begin
        if Thetob.GetValue('VALIDE') = 'X' then OkOpenXLS := True else OkOpenXLS := False;
        result := true;
        For Ind := 0 to TTobSaisieVar.Detail.count -1 do
        begin
          ToblSaisieVar := TTobSaisieVar.Detail[Ind];
          if ToblSaisieVar.getValue('BVD_CODEVARIABLE') = '' then continue;
          ToblVardoc    := TTobVarDoc.FindFirst(['BVD_NUMORDRE','BVD_UNIQUEBLO','BVD_ARTICLE','BVD_CODEVARIABLE'],
                                                [ToblSaisieVar.getValue('BVD_NUMORDRE'),ToblSaisieVar.getValue('BVD_UNIQUEBLO'),
                                                ToblSaisieVar.getValue('BVD_ARTICLE'), ToblSaisieVar.getValue('BVD_CODEVARIABLE')],true);
          If ToblVardoc = nil then //l'enregistrement n'existe pas au niveau des variables du document on le crée (nouvelle variable)
          begin
            // Chargement d'une grille avec les enregistrement de la Tob
            ToblVardoc := Tob.Create('BVARDOC',      TTobVarDoc, -1);
            ToblVardoc.Putvalue('BVD_NATUREPIECE',   ToblSaisieVar.GetValue('BVD_NATUREPIECE'));
            ToblVardoc.Putvalue('BVD_SOUCHE',        ToblSaisieVar.GetValue('BVD_SOUCHE'));
            ToblVardoc.Putvalue('BVD_NUMERO',        ToblSaisieVar.GetValue('BVD_NUMERO'));
            ToblVardoc.Putvalue('BVD_INDICE',        ToblSaisieVar.GetValue('BVD_INDICE'));
            ToblVardoc.Putvalue('BVD_NUMORDRE',      ToblSaisieVar.GetValue('BVD_NUMORDRE'));
            ToblVardoc.Putvalue('BVD_UNIQUEBLO',     ToblSaisieVar.GetValue('BVD_UNIQUEBLO'));
            ToblVardoc.Putvalue('BVD_INDEXFILE',     ToblSaisieVar.GetValue('BVD_INDEXFILE'));
            ToblVardoc.Putvalue('BVD_ARTICLE',       ToblSaisieVar.getValue('BVD_ARTICLE'));
            ToblVardoc.Putvalue('BVD_CODEVARIABLE',  ToblSaisieVar.getValue('BVD_CODEVARIABLE'));
            ToblVardoc.Putvalue('BVD_LIBELLE',       ToblSaisieVar.getValue('BVD_LIBELLE'));
            ToblVardoc.Putvalue('BVD_VALEUR',        ToblSaisieVar.getValue('BVD_VALEUR'));
          end
          else
          begin
            ToblVardoc.PutValue('BVD_VALEUR',        ToblSaisieVar.GetValue('BVD_VALEUR'));
          end;
        end;
      end;
    FINALLY
      TheTob := nil;
    END;
  end;

end;

procedure TMetreArt.SaisieVarDoc (Cledoc : R_CLEDOC; TypeVarDoc : String);
var TobSaisieDoc : TOB;
    ToblSaisieDoc: TOB;
    Ind          : Integer;
    CodeArt      : String;
begin

  //avant d'envoyer à l'écran de saisie des doc il faut faire le ménage en fonction du type passé
  if TypeVarDoc = 'D' then
    CodeArt := CleDoc.Naturepiece + '-' + Cledoc.Souche + '-' + IntToStr(cledoc.NumeroPiece) + '-' + InttoStr(Cledoc.Indice)
  else //if TypeVarDoc = 'L' then
    CodeArt := CleDoc.Naturepiece + '-' + Cledoc.Souche + '-' + IntToStr(cledoc.NumeroPiece) + '-' + InttoStr(Cledoc.Indice) + '-' + InttoStr(Cledoc.NumOrdre) + '-' + InttoStr(Cledoc.UniqueBlo);

  TobSaisieDoc := Tob.create('Variables DOCUMENT/LIGNE',nil, -1);
  For ind := 0 to TTObVarDoc.Detail.count -1 do
  begin
    if TTObVarDoc.Detail[ind].GetValue('BVD_ARTICLE') = CodeArt then
    begin
      ToblSaisieDoc := Tob.Create('BVARDOC', TobSaisieDoc, -1);
      ToblSaisieDoc.Dupliquer(TTobVarDoc.detail[Ind], True, true);
    end;
  end;
      
  TheTOb := TobSaisiedoc;
  TheTOB.Detail.Sort('BVD_INDEXFILE');

  TRY
    AGLLanceFiche('BTP', 'BTVARDOC','','',TypeVarDoc + ';' + CodeArt);
    if (TheTOB <> nil) then
    Begin
      ControleDelVarDoc(TypeVarDoc, CodeArt, TobSaisieDoc);
      //on vérifie si les variables qui sont dans TobSaisieDoc sont dans TTobsaisiedoc...
      For Ind := 0 to TobSaisieDoc.Detail.count -1 do
      begin
        ToblSaisieDoc := TTobVarDoc.FindFirst(['BVD_ARTICLE', 'BVD_CODEVARIABLE'], [CodeArt, TobSaisieDoc.Detail[Ind].GetValue('BVD_CODEVARIABLE')], False);
        //si non on les crées
        if ToblSaisieDoc = nil then
        begin
          // Chargement d'une grille avec les enregistrement de la Tob
          ToblSaisieDoc := Tob.Create('BVARDOC',      TTobVarDoc, -1);
          ToblSaisieDoc.Putvalue('BVD_NATUREPIECE',   TobSaisieDoc.Detail[Ind].GetValue('BVD_NATUREPIECE'));
          ToblSaisieDoc.Putvalue('BVD_SOUCHE',        TobSaisieDoc.Detail[Ind].GetValue('BVD_SOUCHE'));
          ToblSaisieDoc.Putvalue('BVD_NUMERO',        TobSaisieDoc.Detail[Ind].GetValue('BVD_NUMERO'));
          ToblSaisieDoc.Putvalue('BVD_INDICE',        TobSaisieDoc.Detail[Ind].GetValue('BVD_INDICE'));
          ToblSaisieDoc.Putvalue('BVD_NUMORDRE',      TobSaisieDoc.Detail[Ind].GetValue('BVD_NUMORDRE'));
          ToblSaisieDoc.Putvalue('BVD_UNIQUEBLO',     TobSaisieDoc.Detail[Ind].GetValue('BVD_UNIQUEBLO'));
          ToblSaisieDoc.Putvalue('BVD_INDEXFILE',     TobSaisieDoc.Detail[Ind].GetValue('BVD_INDEXFILE'));
          ToblSaisieDoc.Putvalue('BVD_ARTICLE',       TobSaisieDoc.Detail[Ind].getValue('BVD_ARTICLE'));
          ToblSaisieDoc.Putvalue('BVD_CODEVARIABLE',  TobSaisieDoc.Detail[Ind].getValue('BVD_CODEVARIABLE'));
          ToblSaisieDoc.Putvalue('BVD_LIBELLE',       TobSaisieDoc.Detail[Ind].getValue('BVD_LIBELLE'));
          ToblSaisieDoc.Putvalue('BVD_VALEUR',        TobSaisieDoc.Detail[Ind].getValue('BVD_VALEUR'));
        end
        else //si oui on les recharge pour nouvelles valeurs, nouveaux libellé, ...
        begin
          ToblSaisieDoc.PutValue('BVD_LIBELLE',       TobSaisieDoc.Detail[Ind].GetValue('BVD_LIBELLE'));
          ToblSaisieDoc.PutValue('BVD_VALEUR',        TobSaisieDoc.Detail[Ind].GetValue('BVD_VALEUR'));
        end;
      end;
    end;
  FINALLY
    FreeAndNil(TheTob);
  END;

end;

Procedure TMetreArt.ControleDelVardoc(TypeVarDoc : string; CodeArt : string;TobSaisieDoc : ToB);
Var TobLSaisiedoc : TOB;
    TTobLVarDoc   : TOB;
    TextMsg       : string;
    CodeVariable  : String;
begin

  //on contrôle dans un premier temps si des variables Lignes on été supprimée
  TTobLVarDoc := TTobVarDoc.FindFirst(['BVD_ARTICLE'], [CodeArt], False);
  While TTobLVarDoc <> nil do
  begin
    CodeVariable  := TTobLVarDoc.GetValue('BVD_CODEVARIABLE');
    ToblSaisieDoc := TobSaisieDoc.FindFirst(['BVD_CODEVARIABLE'], [CodeVariable], False);
    Freeandnil(TTObLVarDoc);
    If ToblSaisieDoc = nil then
    begin
      if TypeVarDoc = 'D' then
        textMsg := 'La variable document '
      else
        textMsg := 'La Variable Ligne ';

      TextMsg := TextMsg + CodeVariable + ' a été supprimée. Certains calculs pourraient s''en trouver erronés.' + Chr(10);
      TextMsg := TextMsg + 'Afin de conserver une certaine cohérence dans vos données nous allons la supprimer des variables liées ';
      if TypeVarDoc = 'D' then
        textMsg := TextMsg + ' au document. '
      else
        textMsg := TextMsg + ' à la Ligne. ';
      PGIInfo(TextMsg, 'Suppression Variables');
      //On fait une suppression physique de l'enregistrement...
      ExecuteSQL('DELETE BVARDOC WHERE BVD_ARTICLE="' + CodeArt + '" AND BVD_CODEVARIABLE="' + CodeVariable + '"');
    end;
    TTobLVarDoc := TTobVarDoc.FindNext(['BVD_ARTICLE'], [CodeArt], False);
  end;

end;

procedure TMetreArt.CalculMetreDocXLsArt(TOBL : TOB);
Var Nomfichier  : string;
begin

  if TOBL = nil then exit;

  Nomfichier := TOBL.GetString('FICHIERXLS');
  if NomFichier = '' then Exit;

  //Ouverture de l'application Excel
  if Not OpenXLS(False,False) then Exit;

  vMSExcel.DisplayAlerts := false; //enlève les boîtes de dialogue dans Excel

  //Ouverture du classeur ligne de document
  OpenWorkbookXLS(Nomfichier);

  //On charge la feuille Variables avec la tob TTobVarDoc
  ChargeFeuilleWithVariables;

  //se repositionner sur la feuille N°1
  vXLWorkbook.WorkSheets[1].Select;
  //
  vMSExcel.Calculation   := XlsAutomatic;
  vMSExcel.MaxChange     := 0.001;
  vXLWorkbook.PrecisionAsDisplayed := False;

  vXLWorkbook.Save;

  if OkCreation then
  begin
    ChargementQtefichierXLSdansTob(Tobl);
  end
  else
  begin
    CloseXLS;  
    if OkOpenXLS then FileExecAndWait (GetExcelPath + 'EXCEL.exe "' + Nomfichier + '"');
    if not OpenXLS then Exit;

    //Ouverture du classeur ligne de document
    OpenWorkbookXLS(Nomfichier);
    //
    ChargementQtefichierXLSdansTob(TOBL);
  end;

  CloseXLS;

end;

Procedure TMetreART.CalculMetreDocXLsOuvrage(TOBL : TOB);
Var ind         : Integer;
    Nomfichier  : string;
    TobOuvrage  : TOB;
begin

  TobOuvrage := nil;

  //On calcul d'abord chaque fichier indépendament les uns des autres
  For Ind := 0 to TTObOuvrage.Detail.count -1 do
  begin
    //si le code article est renseigné et que pas le bon enreg on continue sans rien ouvrir ou charger...
    //if (ArticleXLS <> '') And (ArticleXLS=TTObOuvrage.detail[ind].GetString('ARTICLE')) then continue;
    Nomfichier := TTObOuvrage.Detail[Ind].GetString('FICHIERXLS');

    if NomFichier = '' then Continue;

    if Not FileExists(Nomfichier) then Continue;

    //Ouverture de l'application Excel
    if not OpenXLS(False, False) then Exit;

    //Ouverture du classeur ligne de document
    OpenWorkbookXLS(Nomfichier);
    //
    vMSExcel.DisplayAlerts := false; //enlève les boîtes de dialogue dans Excel
    vMSExcel.Calculation   := XlsAutomatic;
    vMSExcel.MaxChange     := 0.001;
    vXLWorkbook.PrecisionAsDisplayed := False;

    //On charge la feuille Variables avec la tob TTobVarDoc
    ChargeFeuilleWithVariables;

    //se repositionner sur la feuille N°1
    vXLWorkbook.WorkSheets[1].Select;
                   
    vXLWorkbook.Save; //As(Nomfichier);

    if OkCreation then
      ChargementQtefichierXLSdansTob(TTObOuvrage.Detail[Ind])
    else
    begin
      CloseXLS;
      if TTObOuvrage.detail[ind].GetValue('TYPEARTICLE') = 'OUV' then
      begin
        TobOuvrage := TTObOuvrage.detail[ind];
        break;
      end
      Else
      begin
        if (ArticleXLS <> '') And (ArticleXLS=TTObOuvrage.detail[ind].GetString('ARTICLE')) And (OkOpenXLS) then FileExecAndWait (GetExcelPath + 'EXCEL.exe "' + Nomfichier + '"');
      end;
      if not OpenXLS then Exit;
      //Ouverture du classeur ligne de document
      OpenWorkbookXLS(Nomfichier);
      //
      ChargementQtefichierXLSdansTob(TTObOuvrage.Detail[Ind]);
    end;
  end;

  if not OkCreation then
  begin
    if ((TOBL.FieldExists('GL_TYPEARTICLE'))  And (TOBL.GetValue('GL_TYPEARTICLE')  = 'OUV')) OR
       ((TOBL.FieldExists('BLO_TYPEARTICLE')) And (TOBL.GetValue('BLO_TYPEARTICLE') = 'OUV')) then
    begin
      if TobOuvrage <> nil then Nomfichier := TobOuvrage.GetString('FICHIERXLS');
      if OkOpenXLS then FileExecAndWait (GetExcelPath + 'EXCEL.exe "' + Nomfichier + '"');
      if not OpenXLS then exit;
      //Ouverture du classeur ligne de document
      OpenWorkbookXLS(Nomfichier);
      //
      ChargementQtefichierXLSdansTob(TobOuvrage);
    end;
  end;

  CloseXLS;

end;

Procedure TMetreArt.ChargementQtefichierXLSdansTob(TobTravail : TOB);
Var QteXls  : Double;
begin

  if tobtravail.FieldExists('FICHIERXLS') then
    QteXls := RecupValeurMetreDoc(TobTravail.GetString('FICHIERXLS'))
  else
    QteXls := RecupValeurMetreDoc(fFileName);

  if TobTravail = nil then exit;

  TobTravail.PutValue('QTE', QteXLS);

  CloseXLS;

end;

procedure TMetreArt.ChargeFeuilleWithVariables;
Var ind     : Integer;
    TOBLVar : Tob;
    Nomfeuille : String;
    Adresse : Variant;
begin

  FreeAndNil(TOBLVar);

  //On s'assure que la feuille Variable n'existe pas
  for ind := 1 To vXLWorkbook.Worksheets.Count do
  begin
    vXLsheet := vXLWorkbook.Worksheets[ind];
    Nomfeuille := vXLWorkbook.Worksheets[ind].name;
    if Nomfeuille = 'VARIABLES' then vXLsheet.Delete;
  end;

  //Création d'une feuilles variables
  CreateWorkSheetXLS(EmptyParam,vXLWorkbook.Worksheets[vXLWorkbook.Worksheets.count],1,xlWorksheet,'VARIABLES');

  //For Ind := 0 To TTobVarDoc.detail.count -1 do
  For Ind := 0 To TTobSaisieVar.detail.count -1 do
  begin
    Try
      TOBLVar := TTobSaisieVar.detail[ind];
      //On se positionne sur la cellule A3
      If TOBLVar.GetString('BVD_CODEVARIABLE') = '' then continue;
      vXLsheet.Cells.Item[Ind+3,1].value  := TOBLVar.GetString('BVD_LIBELLE');
      vXLsheet.cells.Item[Ind+3,2].Name   := TOBLVar.GetString('BVD_CODEVARIABLE');
      vXLsheet.Cells.Item[Ind+3,2].value  := TOBLVar.GetString('BVD_VALEUR');
    except
      on E : Exception do
      begin
        PGIError('Erreur Paramétrage variable : ' + TOBLVar.GetString('BVD_CODEVARIABLE'));
      end;
    end;
  end;

  vXLsheet.Visible := False;

  //A faire uniquement si on ne trouve aucune étiquette PGI_RESULTAT
  if not MetreResultatExiste then
  begin
    vXLWorkbook.Worksheets[1].cells.Item[1,1].value := 'Résultat : ';
    vXLWorkbook.Worksheets[1].cells.Item[1,2].Name := vXLWorkbook.Worksheets[1].name + '!PGI_RESULTAT';
    vXLWorkbook.Worksheets[1].Select;
    Adresse := RechercheEtiquette(vXLWorkbook.Worksheets[1].name + '!PGI_RESULTAT');
    if not VarIsEmpty(Adresse) then ColorCellule(Adresse, 3,6);
  end;

end;

Function TmetreArt.MetreResultatExiste : Boolean;
Var CodeVar : string;
    NomSheet: String;
    Ind     : Integer;
begin

  Result := false;

  for Ind := 1 to vXLWorkbook.Names.Count do
  begin
    //Chargement de la cellule correspondante à une variable
    VCell     := vXLWorkbook.Names.Item(ind, EmptyParam, EmptyParam);
    NomSheet  := vXLWorkbook.Worksheets[1].Name;
    if VarIsEmpty(VCell) then Continue;
    CodeVar   := VCell.Name;
    if CodeVar = NomSheet + '!PGI_RESULTAT' then
    begin
      Result  := True;
      Break;
    end;
  end;

end;

Function TMetreArt.RecupValeurMetreDoc(Nomfichier : string) : Double;
Var ARange  : string;
    CodeVar : string;
    NomSheet: String;
    Ind     : Integer;
    Resultat: Variant;
begin

  Result  := 0;

  if VarisEmpty(vMSExcel) then Exit;

  if VarIsEmpty(vXLWorkbook) then Exit;

  //On se positionne sur la feuille 1 qui est la seule susceptible de permettre la récupération du résultat
  vXLsheet  := vXLWorkbook.Worksheets[1];
  vXLWorkbook.WorkSheets[1].Select;
  NomSheet  := vXLSheet.name;
  //
  Try
    //Resultat:= vXLWorkbook.Worksheets[1].cells.Item[1,2].Value;
    for Ind := 1 to vXLWorkbook.Names.Count do
    begin
      //Chargement de la cellule correspondante à une variable
      VCell := vXLWorkbook.Names.Item(ind, EmptyParam, EmptyParam);
      if VarIsEmpty(VCell) then Continue;
      CodeVar   := VCell.Name;
      if CodeVar = NomSheet + '!PGI_RESULTAT' then
      begin
        //On charge la référence relative de la cellule correspondante à une variable
        Try
          ARange    := VCell.RefersToRange.Address;
          Resultat  := vXLsheet.Range[ARange].value;
          Result    := Resultat;
          break;
        except
          //on E : Exception do
          //begin
          //  PGIError('Erreur dans le fichier des métrés : ' + ExtractFileName(fFileName));
          //end;
        end;
      end;
    end;
    //
    Result  := Resultat;
  except
    on E : Exception do
    begin
      PGIError('Erreur Récupération valeur dans le fichier des métrés : ' + ExtractFileName(Nomfichier));
    end;
  end;

  //

end;

Procedure TMetreart.SauveTob;
Var Ind : Integer;
begin

  //On fait le ménage dans la TOBMetres afin de ne prendre que les enregistrement avec un fichier
  for ind := 0 to TTOBMetres.detail.count -1 do
  begin

  end;

  TTOBMetres.SetAllModifie(True);
  TTOBMetres.InsertOrUpdateDB(true);
  //
  RenumerotationTob(TTobVarDoc);
  //
  TTobVarDoc.SetAllModifie(True);
  TTobVarDoc.InsertOrUpdateDB(true);
  //
  TTobSaisieVar.ClearDetail;
  //
end;

procedure TMetreArt.LibereTob;
begin

  FreeAndNil(TTOBMetres);
  FreeAndNil(TTOBVarDoc);
  FreeAndNil(TTobSaisieVar);

end;

Procedure TMetreArt.DetruitMetreDoc(TobPiece : Tob);
Var StSQL : string;
    ind   : Integer;
    TOBL  : TOB;
    RefArt: string;
    NAture: string;
    Souche: string;
    Numero: Integer;
    Indice: Integer;
begin

  if not OkMetre then Exit;

  if TobPiece = nil then Exit;

  Nature:= TobPiece.GetString('GP_NATUREPIECEG');
  Souche:= TobPiece.GetString('GP_SOUCHE');
  Numero:= TobPiece.GetValue('GP_NUMERO');
  Indice:= TobPiece.GetValue('GP_INDICEG');

  //suppression des enregistrements VARDOC associés à la pièce
  StSQL := 'DELETE BVARDOC WHERE BVD_NATUREPIECE="' + Nature + '" ';
  StSQL := StSQL + 'AND BVD_SOUCHE="' + Souche + '" ';
  StSQL := StSQL + 'AND BVD_NUMERO= ' + IntToStr(Numero) + ' ';
  StSQL := StSQL + 'AND BVD_INDICE= ' + IntToStr(Indice);
  if ExecuteSQL(StSQL) < 0 then PGIError('La suppression des informations dans BVARDOC n''a pas fonctionné', 'Gestion des métrés');

  //suppression des enregistrements BMETRE associés à la pièce
  StSQL := 'DELETE BMETRE WHERE BME_NATUREPIECEG="' + Nature + '" ';
  StSQL := StSQL + 'AND BME_SOUCHE="' + Souche + '" ';
  StSQL := StSQL + 'AND BME_NUMERO= ' + IntToStr(Numero) + ' ';
  StSQL := StSQL + 'AND BME_INDICEG= ' + IntToStr(Indice);
  if ExecuteSQL(StSQL) < 0 then PGIError('La suppression des informations dans BMETRE n''a pas fonctionné', 'Gestion des métrés');

  For Ind := 0 to TobPiece.Detail.Count -1 do
  begin
    TOBL      := TobPiece.Detail[Ind];
    RefArt    := TOBL.GetString('GL_ARTICLE');
    RefArt    := FormatArticle(RefArt)+ '-' + TOBL.GetString('GL_NUMORDRE') + '-' + TOBL.GetString('UNIQUEBLO');
    //Suppression du fichier dans répertoire poste de Travail
    SupprimeFichierXLS(fRepMetreDocLocal + RefArt + '.xlsx');
    //suppression du fichier dans répertoire serveur
    SupprimeFichierXLS(fRepMetreDoc + RefArt + '.xlsx');
  end;

  //Suppression des répertoire local et serveur associés à la pièce
  DeleteRepertoire(fRepMetreDocLocal);
  DeleteRepertoire(fRepMetreDoc);

  LibereTob;

end;

Procedure TMetreArt.ValideMetreDoc(TobPiece : Tob);
Var ind       : Integer;
    TOBL      : TOB;
    StRep     : string;
    FileNameO : string;
    FileNameD : string;
    FileName  : String;
begin

  if TTobMetres = nil then Exit;

  //Si on est en création de devis on crée sur le serveur le répertoire de ce
  //dernier au moment de la validation
  if Action = TaCreat then
  begin
    StRep := TobPiece.GetString('GP_NATUREPIECEG') + '-' +  TobPiece.GetString('GP_SOUCHE') + '-' + IntToStr(TobPiece.GetValue('GP_NUMERO')) + '-' + IntToStr(TobPiece.GetValue('GP_INDICEG'));
    fRepMetreDoc := fRepMetreDoc + StRep +'\';
  end;

  if CreationRepertoire(fRepMetreDoc) then
  begin
    For Ind := 0 to TTobMetres.Detail.Count -1 do
    begin
      TOBL      := TTobMetres.Detail[Ind];
      FileNameO := TOBL.GetString('BME_NOMDUFICHIER');
      FileName  := Extractfilename(FileNameO);
      FileNameD := fRepMetreDoc + Filename;
      //Si le Numéro est à zéro cela signifie que la quantité à été modifiée et le métré supprimé...
      if TOBL.GetValue('BME_NUMERO')=0 then
      begin
        //Suppression du fichier sur le serveur...
        SupprimeFichierXLS(FilenameD);
        TOBL.Free;
        Continue;
      end;
      //si le chemin origine = Chemin départ on fait rien.... Mais ceci est impossible !!!!
      if FileNameO=FileNameD then continue;

      if FileNameO <> '' THEN
      Begin
        //copie du fichier du répertoire poste de Travail vers le répertoire serveur
        CopieLocaltoServeur(FileNameO, FilenameD);
        //
        //On modifie le Nom du fichier dans la TOB BMETRE
        TOBL.PutValue('BME_NOMDUFICHIER', FilenameD);
        //
        //Suppression du fichier dans répertoire poste de Travail si en création devis
        if Action = TaCreat then SupprimeFichierXLS(FileNameO);
      end
      else
        TOBL.PutValue('BME_NOMDUFICHIER', '');
    end;
  end;

  SauveTob;

  //On supprime le répertoire sur le poste de travail
  if Action = TaCreat then
  begin
    DeleteRepertoire(fRepMetreDocLocal);
    LibereTob;
  end;

end;

procedure TMetreArt.ChargementVariablesArt(CodeArt : string);
Var QQ        : TQuery;
    StSql     : String;
begin

  //Creation des TOB Virtuelles Variable Application, Variables Document et Variables déjà renseignées
  if Assigned(TTobVariables) then
    TTobVariables.ClearDetail
  else
    TTobVariables := TOB.Create('LES VARIABLES',nil,-1);

  //Lecture de la table des variables Saisie (BTVARDOC)
  StSQL := 'SELECT * FROM BVARIABLES WHERE BVA_TYPE = "B" ';
  StSQL := StSQL + 'AND BVA_ARTICLE = "' + CodeArt + '"';

  //Chargement de la TOB
  QQ := OpenSQL(StSQL,False,-1, '', True);
  If not QQ.Eof Then TTobVariables.LoadDetailDB('BVARIABLES', '', '', QQ, true);

  Ferme (QQ);

end;

procedure TMetreArt.LectureTableVariable(TypeVar : string; Cledoc : R_CLEDOC);
Var QQ        : TQuery;
    StSql     : String;
    stWhere   : String;
begin

  //Creation des TOB Virtuelles Variable Application, Variables Document et Variables déjà renseignées
  if Assigned(TTobVariables) then
    TTobVariables.ClearDetail
  else
    TTobVariables := TOB.Create('LES VARIABLES',nil,-1);

  //Lecture de la table des variables Saisie (BTVARDOC)
  StSQL := 'SELECT * FROM BVARIABLES WHERE BVA_TYPE = "' + TypeVar + '"';

  if TypeVar = 'D' then
    stwhere := ' AND BVA_ARTICLE="' + Cledoc.NaturePiece + '-' + CleDoc.Souche + '-' + IntToStr(CleDoc.NumeroPiece) + '-' + IntToStr(CleDoc.Indice) + '"'
  else if TypeVar = 'L' then
    stwhere := ' AND BVA_ARTICLE="' + Cledoc.NaturePiece + '-' + CleDoc.Souche + '-' + IntToStr(CleDoc.NumeroPiece) + '-' + IntToStr(CleDoc.Indice) + '-' + IntToStr(CleDoc.NumOrdre) + '-' + IntToStr(CleDoc.UniqueBlo) + '"'
  else
    stwhere := '';

  StSQL := StSQL + StWhere;     

  //Chargement de la TOB
  QQ := OpenSQL(StSQL,False,-1, '', True);
  If not QQ.Eof Then TTobVariables.LoadDetailDB('BVARIABLES', '', '', QQ, true);

  Ferme (QQ);

end;

// Suppression des variables associées à l'ouvrage de référence
procedure TMetreArt.SupprimeVariable(Article : string);
var StSQL : string;
begin

  StSQL := 'DELETE FROM BVARIABLES WHERE BVA_ARTICLE="' + Article + '"';
  if ExecuteSQL(StSQL) < 0 then PGIError('La suppression des Variables article : ' + Article + ' n''a pas fonctionné', 'Gestion des métrés');

end;

Procedure TMetreArt.SuppressionLigneMetreDoc(TOBL : TOB);
Var TobDelVarDoc  : TOB;
    TobDelMetre   : TOB;
    Prefixe       : string;
    Nomfichier    : string;
    NaturePiece   : string;
    Souche        : string;
    Numero        : Integer;
    Indice        : Integer;
    NumOrdre      : Integer;
begin

  //on cherche d'abord le préfixe....
  Prefixe := GetPrefixeTable(TOBL);

  NaturePiece   := TObl.GetValue(Prefixe + 'NATUREPIECEG');
  Souche        := TObl.GetValue(Prefixe + 'SOUCHE');
  Numero        := TObl.GetValue(Prefixe + 'NUMERO');
  Indice        := TObl.GetValue(Prefixe + 'INDICEG');
  NumOrdre      := TObl.GetValue(Prefixe + 'NUMORDRE');

  //suppression des variables de l'article associées à la ligne
  TobDelVarDoc := TTobVarDoc.Findfirst(['BME_NATUREPIECEG','BME_SOUCHE','BME_NUMERO','BME_INDICEG','BME_NUMORDRE'], [NaturePiece, Souche, Numero, Indice, NumOrdre ], True);
  while TobDelVarDoc <> nil do
  begin
    TobDelVarDoc.DeleteDB;
    TobDelVarDoc.free;
    TobDelVarDoc :=  TTobVarDoc.FindNext(['BME_NATUREPIECEG','BME_SOUCHE','BME_NUMERO','BME_INDICEG','BME_NUMORDRE','BVD_NUMORDRE'], [NaturePiece, Souche, Numero, Indice, NumOrdre], True);
  end;

  //Suppression des fichier associés à la ligne
  TobDelMetre := TTOBMetres.Findfirst(['BME_NATUREPIECEG','BME_SOUCHE','BME_NUMERO','BME_INDICEG','BME_NUMORDRE'], [NaturePiece, Souche, Numero, Indice, NumOrdre], True);
  while TobDelMetre <> nil do
  begin
    Nomfichier  := TobDelMetre.GetString('BME_NOMDUFICHIER');
    //suppression du fichier XLS sur repertoire local associé
    SupprimeFichierXLS(Nomfichier);
    //
    TobDelMetre.DeleteDB;
    TobDelMetre.free;
    TobDelMetre :=  TTOBMetres.FindNext(['BME_NATUREPIECEG','BME_SOUCHE','BME_NUMERO','BME_INDICEG','BME_NUMORDRE'], [NaturePiece, Souche, Numero, Indice, NumOrdre], True);
  end;

end;

function TMetreArt.IsArticle (TOBL : TOB) : boolean;
var prefixe : string;
    TheValue: string;
begin

  result := false;

  Prefixe := GetPrefixeTable (TOBL);

  if prefixe = '' then exit;

  TheValue := TOBL.GetValue(prefixe+'TYPELIGNE');

  if (TheValue='ART') or (TheValue='ARV') then result := true;

end;

function TMetreArt.QteModifie(TOBL : TOB; NumOrdre : Integer; OldQte, NewQte : Double; OK_SsDetail : Boolean=False): boolean;
Var TOBTemp       : TOB;
    Prefixe       : string;
    NatPiece      : string;
    Souche        : string;
    Numero        : Integer;
    Indice        : Integer;
    Question      : string;
begin

  result := True;

  if not OkMetre then Exit;

  Question  := '';
  //
  Prefixe   := GetPrefixeTable(TOBL);
  //
  NatPiece  := TOBL.GetValue(Prefixe + 'NATUREPIECEG');
  Souche    := TOBL.GetValue(Prefixe + 'SOUCHE');
  Numero    := TOBL.GetValue(Prefixe + 'NUMERO');
  Indice    := TOBL.GetValue(Prefixe + 'INDICEG');
  UniqueBlo := 0;
  //
  if prefixe = 'BLO_' then UniqueBlo := TOBL.GetValue(Prefixe + 'UNIQUEBLO');
  //
  TOBTemp := TTOBMetres.FindFirst(['BME_NATUREPIECEG','BME_SOUCHE','BME_NUMERO','BME_INDICEG','BME_NUMORDRE','BME_UNIQUEBLO'],[NatPiece,Souche,Numero,Indice,NumOrdre,Uniqueblo],False);
  //
  if TobTemp = nil then Exit;

  //FV1 - 05/01/2017 : FS#2325 - Divers pbs liés aux métrés en saisie de devis
  if TobTemp.GetValue('BME_NOMDUFICHIER') = '' then Exit;

  if Ok_SsDetail then Question := 'Vous allez remplacer la quantitié de : ' + FloatToStr(OldQte) + ' par la quantité suivante : ' + FloatToStr(NewQte) + ' :' + Chr(10);
  Question := Question + 'Confirmez-vous la suppression du fichier métré associé ?';

  //attention : Ceci devra être revu dans le cas de la suppression d'un métré mais qui finalement débouche sur un abandon de saisie....
  if PGIAsk (TraduireMemoire(Question), TraduireMemoire('Modification Qté')) = mrYes then
  begin
    if Ok_SsDetail then
    begin
      if Application.MessageBox(pchar('Attention : la suppression du fichier métré associé est IMMEDIATE et TOTALE !' + Chr(10) + 'Confirmez-vous la suppression du fichier ?'),pchar('Modification Qté Ss-Détail'),MB_YESNO Or MB_ICONQUESTION Or MB_DEFBUTTON2) = MrYes then
        SupprimeMetre(TOBTemp)
      else
        Result := False;
    end
    else
      SupprimeMetre(TOBTemp);
  end
  else
    Result := False;

end;

Procedure TMetreArt.SupprimeMetre(TOBTemp : TOb);
begin

  FileName  := TOBTemp.GetValue('BME_NOMDUFICHIER');

  //Suppression du fichier associé dans le répertoire local...
  SupprimeFichierXLS(FileName);

  FreeAndNil(TobTemp);

end;

function TMetreArt.IsSousDetail (TOBL : TOB) : boolean;
var TheValue : string;
begin

  result := false;

  if TOBL = nil then exit;

  IF TOBL.NomTable='LIGNE' then
    TheValue := TOBL.GetValue('GL_TYPELIGNE')
  else
    TheValue := TOBL.GetValue('BLO_TYPELIGNE');

  if (TheValue='SDV') or (TheValue='SD') then result := true;

end;

function TMetreArt.FindLigneOuvrageinDoc(TOBD: TOB): TOB;
var Inddep      : Integer;
    Indice      : integer;
		indOuvrage  : integer;
    Prefixe     : String;
begin

	result  := nil;

  Prefixe := GetPrefixeTable(TOBD);

  Inddep  := TOBD.GetIndex ;

  IndOuvrage := TOBD.GetValue(Prefixe + 'INDICENOMEN');

  for Indice := Inddep-1 downto 0 do
  begin
    if ISOuvrage(TTOBpiece.detail[Indice]) then
    begin
      if prefixe <> 'BLO_' then
      begin
        if (TTOBpiece.detail[Indice].getValue('GL_INDICENOMEN')=IndOuvrage) then
        begin
          if (IsArticle(TTOBpiece.detail[Indice])) then
          begin
            Result := TTOBpiece.detail[Indice];
            break;
          end;
        end;
      end
      else
      begin
        if (IsArticle(TTOBpiece.detail[Indice])) then
        begin
          Result := TTOBpiece.detail[Indice];
          break;
        end;
      end;
    end;
  end;

end;

function TMetreArt.GetPrefixeTable (TOBL : TOB) : string;
begin
  result := '';
  if TOBL.nomtable = 'PIECE'        then result := 'GP_' else
  if TOBL.nomtable = 'LIGNE'        then result := 'GL_' else
  if TOBL.nomtable = 'LIGNEOUVPLAT' then result := 'BOP_' else
  if TOBL.nomtable = 'LIGNEOUV'     then result := 'BLO_';
end;


end.

