unit MSaveAPI;

interface

uses
    Windows, Classes, Graphics, SysUtils, Forms, { ... }
    Controls, StdCtrls, ExtCtrls, DB, Bde, { ... }
    IniFiles, uMessage, uModeles,USaveDp; { ! }

type
	TRestoreAction = set of ( raAddModel, raAskModel, { ... }
    	raCreateModel, raDontCopy, raUseDefault );  { ! }

    TDocumentDecisiv = class(TComponent)  { Type document ... }
	private  { Declarations privees }
        FFileName       	: string;
        FDateFile       	: TDateTime;
        FIsDossier			: Boolean;  { ! }

    	FOnChecking 		: TNotifyEvent;
        FOnRestoring		: TNotifyEvent;

        procedure SetOnChecking(Value: TNotifyEvent);
        procedure SetOnRestoring(Value: TNotifyEvent);  { ! }

        procedure SetDateFile(Value: string);

        function  GetDateFile: string;

	public  { Declarations publiques }
		constructor Create(aOwner: TComponent); override;

        function  Restore(aOwner: TComponent): Boolean;
        function  Check(aOwner: TComponent): Boolean;

        property  OnChecking: TNotifyEvent read FOnChecking
        	write SetOnChecking;

        property  OnRestoring: TNotifyEvent read FOnRestoring
        	write SetOnRestoring;

	published  { Declarations publiees }
        property DateFile: string read GetDateFile write SetDateFile;
        property FileName: string read FFileName write FFileName;
        property IsDossier: Boolean read FIsDossier write FIsDossier;
    end;

	{------------------------------------------------------------------}

	TDossierDecisiv = class(TComponent)  { Type dossier ... }
	private  { Declarations privees }
		FCodeSociete		: string;  { Code abrege de la societe ... }
        FNomSociete			: string;  { ... nom de la societe ... }
		FDateCloture		: TDateTime; { ... date de cloture ... }
        FDureeExercice		: Smallint;  { ... et duree de l'exercice }

		FLocalisation		: string;  { Localisation du dossier }

        FMoisCloture		: Smallint;  { ... }
        FDestinataire		: string;
        FFormeJuridique 	: string;
        FTaille				: string;  { ! }

        FSiret				: string;  { Numero de SIRET ... }
        FAPE				: string;  { Code APE }
        FCodeNAF			: string;  { Code NAF }
        FCodeFRP			: string;  { Code FRP }
        FCodeIntraCom		: string;  { Code Intra-Communautaire }

        FAdresse1			: string;  { Adresses ... }
        FAdresse2			: string;
        FAdresse3			: string;
        FAdresse4			: string;  { ... }

        FCodePostal			: string;  { ... }
        FVille				: string;
        FPays				: string;
        FTelephone			: string;
        FTelecopie			: string;  { ! }

        FDateCreation		: TDateTime;  { Dates de creation et ... }
        FDateCessation		: TDateTime;  { ... cessation }
        FActivite			: string;  { Activite ... }
        FChgAct				: Boolean; { ... et changement }

        FInsRegMet			: Boolean; { Registre du ... }
        FInsRegCom			: Boolean;  { ... }
        FVilRegCom			: string;  { ... }
        FNumRegCom			: string;  { ... commerce }

        FTauxTVA			: Double;  { TVA ... }
        FProrataTVA			: Double;
        FRegimeTVA			: string;  { ! }

        FRegImp				: string;  { Regimes d'imposition et ... }
        FRegDec				: string;  { ... de declaration }

        FAdrAnc				: string;  { ... }
        FDateDebutEx		: TDateTime; { ! }

        FDateClotureN1		: TDateTime;  { ... }
        FExerciceN1			: Smallint;
        FDateClotureN2 		: TDateTime;
        FExerciceN2			: Smallint;
        FDateClotureN3		: TDateTime;
        FExerciceN3			: Smallint;
        FDateClotureN4		: TDateTime;
        FExerciceN4			: Smallint;  { ! }

        FDateGeneration		: TDateTime; { ... }
        FModeGeneration		: string;

        FMonnaiePrincipale	: string;
        FMonnaieSecondaire	: string;

		FOrigineDossier		: string; { ! }
        FOrigineDate    	: TDateTime;
        FEtatCloture    	: string;
        FDateEffectiveCloture: TDateTime;
        FUserCloture    	: string;

		FProduits			: TList;  { ... }
        FModeles			: TList;
		FModules			: TList;
        FDocuments			: TList;

        FDossierDP          : TDossierDPDecisiv;  { ! }

//		FDMOrganisation 	: TDMOrganisation;
        FOrgDos         	: TList;

        FOnChecking			: TNotifyEvent;
		FOnRestoring		: TNotifyEvent;

		function  GetDateCreation: string; { ... }
        function  GetDateCessation: string;
        function  GetDateDebutEx: string;
        function  GetDateGeneration: string;

		function  GetDateCloture: string;  { ... }
		function  GetDateClotureN1: string;
		function  GetDateClotureN2: string;
		function  GetDateClotureN3: string;
		function  GetDateClotureN4: string; { ! }

        function  GetOrigineDate: string;
        function  GetDateEffectiveCloture: string;

		procedure GetChildren(Proc: TGetChildProc; { Enfants ... }
			Root: TComponent); override;  { ... composant }
        procedure Notification(aComponent: TComponent;
        	Operation: TOperation); override;  { ... }
        procedure SetOnChecking(Value: TNotifyEvent); { ! }
        procedure SetOnRestoring(Value: TNotifyEvent);  { ... }

		procedure SetDateCreation(Value: string); { ... }
		procedure SetDateCessation(Value: string);
        procedure SetDateDebutEx(Value: string);
        procedure SetDateGeneration(Value: string);
        procedure SetOrigineDate(Value: string);
        procedure SetDateEffectiveCloture(Value: string);

		procedure SetDateCloture(Value: string);  { ... }
		procedure SetDateClotureN1(Value: string);
		procedure SetDateClotureN2(Value: string);
		procedure SetDateClotureN3(Value: string);
		procedure SetDateClotureN4(Value: string);  { ! }

	public  { Declarations publiques }
		constructor Create(aOwner: TComponent); override;
		destructor  Destroy; override;  { ! }

		function  RestoreCount: Integer;
//Cad le 10/11/98
        function  Controle: Boolean;
//FIN Cad le 10/11/98
        function  Restore: Boolean;
        function  Check: Boolean;  { ! }

        property  OnChecking: TNotifyEvent read FOnChecking
        	write SetOnChecking;

        property  OnRestoring: TNotifyEvent read FOnRestoring
        	write SetOnRestoring;
        
	published  { Declarations publiees }
		property  CodeSociete: string read FCodeSociete
        	write FCodeSociete;  { ! }

		property  NomSociete: string read FNomSociete
        	write FNomSociete;  { ! }

		property  DateCloture: string read GetDateCloture
			write SetDateCloture;  { ! }

        property  DureeExercice: Smallint read FDureeExercice
        	write FDureeExercice;  { ! }

		property  Localisation: string read FLocalisation
        	write FLocalisation;  { ! }

        property  MoisCloture: Smallint read FMoisCloture
        	write FMoisCloture;  { ! }

        property  Destinataire: string read FDestinataire
        	write FDestinataire;  { ! }

        property  FormeJuridique: string read FFormeJuridique
        	write FFormeJuridique;  { ! }
        property  Taille: string read FTaille write FTaille;

        property  Siret: string read FSiret write FSiret;  { SIRET ... }
        property  APE: string read FAPE write FAPE; { ... APE }
        property  CodeNAF: string read FCodeNAF write FCodeNAF;
        property  CodeFRP: string read FCodeFRP write FCodeFRP;  { ! }

        property  CodeIntraCom: string read FCodeIntraCom
        	write FCodeIntraCom;  { ! }

        property  Adresse1: string read FAdresse1 write FAdresse1;
        property  Adresse2: string read FAdresse2 write FAdresse2;
        property  Adresse3: string read FAdresse3 write FAdresse3;
        property  Adresse4: string read FAdresse4 write FAdresse4;

        property  CodePostal: string read FCodePostal write FCodePostal;
        property  Ville: string read FVille write FVille;
        property  Pays: string read FPays write Fpays;
        property  Telephone: string read FTelephone write FTelephone;
        property  Telecopie: string read FTelecopie write FTelecopie;

        property  DateCreation: string read GetDateCreation  { ... }
        	write SetDateCreation;  { Dates de creation ... }
        property  DateCessation: string read GetDateCessation
        	write SetDateCessation;  { ... et de cessation d'activite }

        property  Activite: string read FActivite write FActivite;
        property  ChgAct: Boolean read FChgAct write FChgAct;

        property  InsRegMet: Boolean read FInsRegMet write FInsRegMet;
        property  InsRegCom: Boolean read FInsRegCom write FInsRegCom;
        property  VilRegCom: string read FVilRegCom write FVilRegCom;
        property  NumRegCom: string read FNumRegCom write FNumRegCom;

        property  TauxTVA: Double read FTauxTVA write FTauxTVA;  { TVA }
        property  ProrataTVA: Double read FProrataTVA write FProrataTVA;
        property  RegimeTVA: string read FRegimeTVA write FRegimeTVA;

        property  RegImp: string read FRegImp write FRegImp;
        property  RegDec: string read FRegDec write FRegDec;

        property  AdrAnc: string read FAdrAnc write FAdrAnc;
        property  DateDebutEx: string read GetDateDebutEx
        	write SetDateDebutEx; { Date de debut exercice }

        property  DateClotureN1: string read GetDateClotureN1
        	write SetDateClotureN1;  { ! }
        property  ExerciceN1: Smallint read FExerciceN1 write FExerciceN1;

        property  DateClotureN2: string read GetDateClotureN2
        	write SetDateClotureN2;  { ! }
        property  ExerciceN2: Smallint read FExerciceN2 write FExerciceN2;

        property  DateClotureN3: string read GetDateClotureN3
        	write SetDateClotureN3;  { ! }
        property  ExerciceN3: Smallint read FExerciceN3 write FExerciceN3;

        property  DateClotureN4: string read GetDateClotureN4
        	write SetDateClotureN4;  { ! }
        property  ExerciceN4: Smallint read FExerciceN4 write FExerciceN4;

        property  DateGeneration: string read GetDateGeneration
        	write SetDateGeneration;  { ! }
        property  ModeGeneration: string read FModeGeneration
        	write FModeGeneration;  { ! }

		property  MonnaiePrincipale: string read FMonnaiePrincipale
        	write FMonnaiePrincipale;  { ! }
        property  MonnaieSecondaire: string read FMonnaieSecondaire
        	write FMonnaieSecondaire;  { ! }

        property  OrigineDossier: string read FOrigineDossier
        	write FOrigineDossier;  { ! }
        property  OrigineDate: string read GetOrigineDate
            write SetOrigineDate;
        property  EtatCloture: string read FEtatCloture
        	write FEtatCloture;
        property  DateEffectiveCloture: string read GetDateEffectiveCloture
            write SetDateEffectiveCloture;
        property  UserCloture: string read FUserCloture
            write FUserCloture;
	end;  { ... fin de definition du type Dossier }

	{------------------------------------------------------------------}

	TModeleDecisiv = class(TComponent)  { Modele ... }
	private { Declarations privees }
		FCodeRubrique		: Integer;
        FNomRubrique		: string;
        FNomModele			: string;
		FCodeModule			: Integer;

        FParent				: Integer;

        FModifiable			: SmallInt; { ... }
        FDefaut				: Boolean;
        FVersion			: SmallInt; { ! }

        FDescription		: string;
		FCommentaire		: string;

		FCodeModele			: Smallint;
        FFieldNames			: string;

		FRestoreAction		: TRestoreAction;
		FTables				: TList;
        FOptions			: TRubriqueOptions;

		FOnChecking			: TNotifyEvent;
		FOnRestoring		: TNotifyEvent;

		function  GetModelValue: Smallint;

        procedure SetNomModele(Value: string);

		procedure GetChildren(Proc: TGetChildProc; { Enfants ... }
			Root: TComponent); override;  { ... composant }
        procedure Notification(aComponent: TComponent;
        	Operation: TOperation); override;  { ... }
        procedure SetOnChecking(Value: TNotifyEvent); { ! }
        procedure SetOnRestoring(Value: TNotifyEvent);  { ... }

	protected  { Declarations protegees }
    	function  GetChildOwner: TComponent; override;

	public  { Declarations publiques }
		constructor Create(aOwner: TComponent); override;
		destructor  Destroy; override;

        function  Restore(aOwner: TComponent): Boolean;
        function  RestoreCount: Integer;
        function  Check(aOwner: TComponent): Boolean;

        property  OnChecking: TNotifyEvent read FOnChecking
        	write SetOnChecking;

        property  OnRestoring: TNotifyEvent read FOnRestoring
        	write SetOnRestoring;

	published  { Declarations publiees }
		property  CodeRubrique: Integer read FCodeRubrique  { ... }
			write FCodeRubrique; { ... numero de rubrique }
        property  NomRubrique: string read FNomRubrique
        	write FNomRubrique;  { ... nom de la rubrique }
        property  Nom: string read FNomModele write SetNomModele;

		property  Modifiable: Smallint read FModifiable write FModifiable;
        property  Defaut: Boolean read FDefaut write FDefaut;
		property  Version: Smallint read FVersion write FVersion;  { ! }

        property  Description: string read FDescription  { ... }
        	write FDescription;  { ! }

        property  Commentaire: string read FCommentaire  { ... }
        	write FCommentaire;  { ! }
	end;

	{------------------------------------------------------------------}

	TModuleDecisiv = class(TComponent)  { Module ... }
	private { Declarations privees }
		FCodeModule			: Integer;
        FOption				: Boolean;
		FVersion			: array[0..1] of DWORD;

		FRestoreAction		: TRestoreAction;
		FTables				: TList;

		FOnChecking			: TNotifyEvent;
		FOnRestoring		: TNotifyEvent;

		function  GetVersion: string;

		procedure GetChildren(Proc: TGetChildProc; { Enfants ... }
			Root: TComponent); override;  { ... composant }
        procedure Notification(aComponent: TComponent;
        	Operation: TOperation); override;  { ... }
		procedure SetVersion(Value: string);  { Type }
        procedure SetOnChecking(Value: TNotifyEvent); { ! }
        procedure SetOnRestoring(Value: TNotifyEvent);  { ... }

    protected  { Declarations protegees }
    	function  GetChildOwner: TComponent; override;

	public  { Declarations publiques }
		constructor Create(aOwner: TComponent); override;
		destructor  Destroy; override;

        function  Restore(aOwner: TComponent): Boolean;
        function  RestoreCount: Integer;
		function  Check(aOwner: TComponent): Boolean;

        property  OnChecking: TNotifyEvent read FOnChecking
        	write SetOnChecking;

        property  OnRestoring: TNotifyEvent read FOnRestoring
        	write SetOnRestoring;

	published  { Declarations publiees }
		property  CodeModule: Integer read FCodeModule write FCodeModule;
        property  Option: Boolean read FOption write FOption;
		property  Version: string read GetVersion write SetVersion;
	end;

	{------------------------------------------------------------------}

    TOrgDosDecisiv = class(TComponent)  { Organisation dossier ... }
	private  { Declarations privees }
        FCatOrg         	: string;  { 'PUB[lic]' ou 'PRI[vee]'    }
        FCodOrg         	: string;  { Code organisation générique }
        FLibCodOrg      	: string;  { Libelle           "         }
        FAbrCodOrg      	: string;  { Libelle abrege    "         }
        FAppCodOrg      	: Boolean; { Applicable  (PUB seulement) }
        FValOrg         	: string;  { Code organisation réelle    }
        FLibValOrg      	: string;  { Libelle organisation réelle }
        FAbrValOrg      	: string;  { Abrege organisation réelle  }

		FOnChecking			: TNotifyEvent;
		FOnRestoring		: TNotifyEvent;

        procedure SetOnChecking(Value: TNotifyEvent);
        procedure SetOnRestoring(Value: TNotifyEvent);  { ! }

	public  { Declarations publiques }
		constructor Create(aOwner: TComponent); override;

        function  Restore(aOwner: TComponent): Boolean;
        function  Check(aOwner: TComponent): Boolean;

        property  OnChecking: TNotifyEvent read FOnChecking
        	write SetOnChecking;

        property  OnRestoring: TNotifyEvent read FOnRestoring
        	write SetOnRestoring;

	published  { Declarations publiees }
        property CatOrg : string read FCatOrg write FCatOrg;
        property CodOrg : string read FCodOrg write FCodOrg;
        property LibCodOrg: string read FLibCodOrg write FLibCodOrg;
        property AbrCodOrg: string read FAbrCodOrg write FAbrCodOrg;
        property AppCodOrg: Boolean read FAppCodOrg write FAppCodOrg;
        property ValOrg: string read FValOrg write FValorg;
        property LibValOrg: string read FLibValOrg write FLibValOrg;
        property AbrValOrg: string read FAbrValOrg write FAbrValOrg;
    end;

	{------------------------------------------------------------------}

	TProduitDecisiv = class(TComponent)  { Produit }
	private { Declarations privees }
		FCodeProduit		: Integer;
		FVersion			: string;

		FOnChecking			: TNotifyEvent;
		FOnRestoring		: TNotifyEvent;

        procedure SetOnChecking(Value: TNotifyEvent); { ... }
        procedure SetOnRestoring(Value: TNotifyEvent);  { ! }

	public  { Declarations publiques }
		constructor Create(aOwner: TComponent); override;

        function  Check(aOwner: TComponent): Boolean;  { ... }
        function  Restore(aOwner: TComponent): Boolean;  { ! }

        property  OnChecking: TNotifyEvent read FOnChecking
        	write SetOnChecking;

        property  OnRestoring: TNotifyEvent read FOnRestoring
        	write SetOnRestoring;  { ! }

	published  { Declarations publiees }
		property  CodeProduit: Integer read FCodeProduit  { ... }
			write FCodeProduit;  { ! }
		property  Version: string read FVersion write FVersion;
	end;

	{------------------------------------------------------------------}

	TTableDecisiv = class(TComponent)  { Type Table ... }
	private { Declarations privees }
        FVersion			: Smallint;

        FOnChecking			: TNotifyEvent;
		FOnRestoring		: TNotifyEvent;

        procedure SetOnChecking(Value: TNotifyEvent);  { ... }
        procedure SetOnRestoring(Value: TNotifyEvent);
        procedure SetDateMaJ(Value: string);  { Date MaJ }

        function  GetDateMaJ: string;
        function  GetTablePath: string;

	public  { Declarations publiques }
        FDateMaJ			: TDateTime;

		constructor Create(aOwner: TComponent); override;
        destructor Destroy; override;

        function  Check(aOwner: TComponent): Boolean;  { ... }
		function  Restore(aOwner: TComponent): Boolean;  { ! }

        property  OnChecking: TNotifyEvent read FOnChecking
        	write SetOnChecking;

        property  OnRestoring: TNotifyEvent read FOnRestoring
        	write SetOnRestoring;

		property TablePath: string read GetTablePath;

	published  { Declarations publiees }
        property  DateMaJ: string read GetDateMaJ  { ... }
        	write SetDateMaJ;  { ! }
    	property  Version: Smallint read FVersion write FVersion;
	end;  { ... fin de definition du type Table }

	{------------------------------------------------------------------}

	TLogFile = class(TFileStream)  { Fichier de LOG }
    private { Declarations privees ... }
		FLogFileName		: string;
    	FIniFile    		: TIniFile;
        FLoggingLevel		: Integer; { ! }

        FWarnings			: Integer;
    	FErrors				: Integer;

    public  { Declarations publiques}
    	constructor Create;
        destructor  Destroy; override;

    	function Write(MsgNo: Integer;  { ... }
        	args: array of const): Longint;

        property LoggingLevel: Integer read FLoggingLevel  { ... }
        	write FLoggingLevel;  { ... niveau de logging }
        property Errors: Integer read FErrors;  { ! }
		property Warnings: Integer read FWarnings;  { ... }
    end;  { ... fin de definition du type fichier de LOG des actions }

	{------------------------------------------------------------------}

    TStepState = ( ssNone, ssExec, ssDone );  { ! }

    TStep = class  { Type phase visuelle ... }
    private { Declarations privees }
        FImage	: TImage;
        FLabel	: TLabel;
    	FState	: TStepState;

        procedure SetState(Value: TStepState);

    public  { Declarations publiques }
        property  State: TStepState read FState write SetState;

    	constructor Create(aImage: TImage; aLabel: TLabel);
    end;  { ... fin de la definition d'une phase de sauvegarde }

	{------------------------------------------------------------------}

    function GetErrorMessage: string; { Type de messages : erreur }
    function GetFileAgeing(Files: string): TDateTime;
    function GetFileVersion(DLLName: string): string;
    function WriteLogFile(MsgNo: Integer; Args: array of const): Boolean;

    function DeleteFiles(Path: string): Boolean;

	function StringToDate(const Value: string): TDateTime;
	function StringToDateTime(const Value: string): TDateTime;

    function DateToString(const DateTime: TDateTime): string;
    function DateTimeToString(const DateTime: TDateTime): string;

var
	BinPath, ModelsPath, TempPath: string;
	LogFile: TLogFile;
    bFilesInUse, bModelInUse, bModuleMDP: Boolean;
    bbalances,bdp:Boolean;

implementation

(*uses MCCConst, uDLLargs, uDmDos, uMCCAPI, uQryMod,
	uOptUTi,URapSsDo;*)
uses MCCConst, uDLLargs, uDmDos, uMCCAPI,
	uOptUTi,URapSsDo;

Const MODULE_MDP = 8;
{**********************************************************************}
{**********************************************************************}

{}
{ Procedures et fonctions globales
{}

function DateToString(const DateTime: TDateTime): string;
begin
	result := FormatDateTime('dd/mm/yyyy', DateTime);
end;

{----------------------------------------------------------------------}

function DateTimeToString(const DateTime: TDateTime): string;
begin
	result := FormatDateTime('dd/mm/yyyy hh:nn:ss', DateTime);
end;

{----------------------------------------------------------------------}

function DeleteFiles(Path: string): Boolean;
var
   	SearchRec	: TSearchRec;
    Found		: Integer;
begin
    Found := FindFirst(Path, faAnyFile, SearchRec);
    result := True;
    Path := ExtractFilePath(Path);

    while Found = 0 do begin  { Parcours des fichiers a supprimer ... }
    	if (SearchRec.Name <> '.') and  { Repertoire courant ... }
        	(SearchRec.Name <> '..') then begin { ... parent }
    		result := DeleteFile(Path+'\'+SearchRec.Name);
        	if not result then break; { Probleme fichier }
        end;  { ... fin de l'effacement d'un fichier ... }
        Found := FindNext(SearchRec);  { ... fichier suivant }
    end; { ... fin du parcours de la liste de fichiers a effacer }
    FindClose(SearchRec); { ... et nettoyage du contexte de recherche }
end;

{----------------------------------------------------------------------}

function GetErrorMessage: string;
begin
   	if GetLastError = NO_ERROR then begin result := ''; exit; end;
   	case GetLastError of  { Gestion des cas d'erreur ... }
		ERROR_WRITE_PROTECT : result := S_E_PROTECT;
        ERROR_NOT_READY		: result := S_E_NO_DISK;
        ERROR_WRONG_DISK	: result := S_E_ALREADY;
        ERROR_GEN_FAILURE	: result := S_E_FAILURE;
    else result := Format(S_E_UNKNOWN, [GetLastError]) end;
    result := result+#13#10#13#10;  { Sauter 2 lignes dans ce cas }
end;

{----------------------------------------------------------------------}

function GetFileAgeing(Files: string): TDateTime;
var
	FileDate	: Integer;  { ... }
    Found 		: Integer;
	Search		: TSearchRec;
begin
	FileDate := 0; result := 0;  { Valeurs initialisees par defaut ... }
	Found := FindFirst(Files, faAnyFile, Search);  { Premier item }
    while Found = 0 do begin { Parcours liste de fichiers ... }
    	if Search.Time > FileDate then FileDate := Search.Time;
    	Found := FindNext(Search); { Item fichier suivant ... }
    end; { ... fin de parcours de la liste de fichiers disque }
    FindClose(Search);  { Liberation du contexte de recherche ... }
    if FileDate > 0 then result := FileDateToDateTime(FileDate); { ... }
end;

{----------------------------------------------------------------------}

function GetFileVersion(DLLName: string): string;
var
	S			: string;
	N			: UINT;
	Handle		: THandle;
	Datas, P	: Pointer;  { ! }
begin
	S := BinPath+DLLName;  { Chemin d'acces aux modules ... }
	N := GetFileVersionInfoSize(PChar(S), Handle);
	if N = 0 then exit; { Pas d'informations ... }
	GetMem(Datas, N); { Allocation du bloc informations }

	try	if not GetFileVersionInfo(PChar(S), Handle, N, Datas) then exit;
		if VerQueryValue(Datas, PChar(S_VERKEY), P, N) then
			S := IntToHex(LoWord(Longint(P^)), 4)+
				IntToHex(HiWord(longint(P^)), 4)  { ! }
		else S := '040904E4';  { Langage et jeu de caracteres ... }

		S := '\StringFileInfo\'+S+'\FileVersion';  { Cle ... }
		if VerQueryValue(Datas, PChar(S), P, N) then
			result := StrPas(P);  { Version module }
	finally FreeMem(Datas); end;  { Nettoyage apres lecture }
end;

{----------------------------------------------------------------------}

procedure GetPaths;
var
    Path		: array[0..255] of Char;
    IniName		: string;
    Section		: string;
    DefaultPath	: string;  { ! }

	{- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }

    procedure AddBackSlash(var S: string);
    begin
    	if S[Length(S)] <> '\' then S := S+'\';
    end;

	{- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }

begin
    IniName := ExtractFilePath(Application.ExeName)+  { Chemin ... }
    	S_DECISIV+'.ini';  { ... et nom du fichier .ini }
    Section := S_DECISIV;  { Section de lecture }
    GetSystemDirectory(Path, 255); { C:\Windows\System ? }
    DefaultPath := ExtractFileDrive(StrPas(Path))+S_DEFAULT;  { ! }

    BinPath := ExtractFilePath(Application.ExeName); { Executables ... }
    ModelsPath := DefaultPath+S_MODELES;  { ... modeles, ... }
    GetTempPath(255, Path); TempPath := StrPas(Path); { ... et travail }

	with TIniFile.Create(IniName) do begin  { Creation du fichier .INI }
		try	ModelsPath := ReadString(Section, 'MODELESPATH', { ... }
        		ModelsPath);  { ... chemin des 'modeles globaux' }
			TempPath := ReadString(Section, 'TEMPPATH', TempPath);
    	finally Free; end; { ... nettoyage de l'objet IniFile cree }
    end;  { ... fin d'utilisation des donnees du fichier de parametres }

    AddBackSlash(BinPath);  { Chemin des executables ... }
	AddBackSlash(ModelsPath); { ... des modeles ... }
    AddBackSlash(TempPath); { ... repertoire de travail }
end;

{----------------------------------------------------------------------}

function StringToDate(const Value: string): TDateTime;
var	S	: string;
begin
	S := ShortDateFormat; ShortDateFormat := 'dd/mm/yyyy';
	try result := StrToDate(Value); { ... }
    except result := 0;  { ! }
    end;  { Par defaut, le 30/12/1899 }
    ShortDateFormat := S; { Cf. parametres regionaux }
end;

{----------------------------------------------------------------------}

function StringToDateTime(const Value: string): TDateTime;
var	S	: string;
begin
	S := ShortDateFormat; ShortDateFormat := 'dd/mm/yyyy';  { ! }
	try result := StrToDateTime(Value);  { ... }
    except result := 0; { ... si erreur }
    end;  { Par defaut, le 30/12/1899 00:00:00 }
    ShortDateFormat := S;  { Cf. parametres regionaux Windows }
end;

{----------------------------------------------------------------------}

function WriteLogFile(MsgNo: Integer; Args: array of const): Boolean;
begin
	if LogFile = nil then  { Creation eventuelle ... }
    	try LogFile := TLogFile.Create;  { ... }
        except result := False; exit end;  { ! }
   	result := (LogFile.Write(MsgNo, Args) > 0);  { ! }
end;

{**********************************************************************}
{**********************************************************************}

{}
{ Classe TOrgDosDecisiv - Organisation pour sauvegarde ou rappel
{}

constructor TOrgDosDecisiv.Create(aOwner: TComponent);
begin
	inherited Create(aOwner);  { Appel des constructeurs parents ... }
end;

{**********************************************************************}

{}
{ Procedures et fonctions publiques }
{}

function TOrgDosDecisiv.Check(aOwner: TComponent): Boolean;
begin
	result := True;  { Organisations non traitees ici }
end;

{----------------------------------------------------------------------}

function TOrgDosDecisiv.Restore(aOwner: TComponent): Boolean;
begin
	try if Assigned(FOnRestoring) then FOnRestoring(self);
    	result := True;  { Non traite dans ce cas }
    except result := False; end;  { ... des fois que ...} 
end;

{**********************************************************************}

procedure TOrgDosDecisiv.SetOnChecking(Value: TNotifyEvent);
begin
	FOnChecking := Value;
end;

{----------------------------------------------------------------------}

procedure TOrgDosDecisiv.SetOnRestoring(Value: TNotifyEvent);
begin
    FOnRestoring := Value;  { Adresse evenement 'OnRestoring' }
end;

{**********************************************************************}
{**********************************************************************}

{}
{ Classe TDocumentDecisiv - type document pour sauvegarde ou rappel
{}

constructor TDocumentDecisiv.Create(aOwner: TComponent);
begin
	inherited Create(aOwner);  { Appel des constructeurs parents ... }
end;

{**********************************************************************}

{}
{ Procedures et fonctions publiques }
{}

function TDocumentDecisiv.Check(aOwner: TComponent): Boolean;
var DocFileName: string;
    dwFileattributes: Integer;
begin
	result := aOwner is TDossierDecisiv;  { Objet proprietaire ... }
    if not result then exit; { ... pas un TDossierDecisiv }
    if Assigned(FOnChecking) then FOnChecking(self);
(*
	with DMDossiers do begin
        DocFileName := MakePathDoc+'\'+FileName;
        if FileExists(DocFileName) and
           ((FDateFile - GetFileAgeing(DocFileName)) < 0) then begin
        	    WriteLogFile(I_W_DOCCHANGED, [FileName]);
                Result := False; Exit;
        end;
        dwFileattributes := GetFileAttributes(PChar(DocFileName));
        if dwFileAttributes and FILE_ATTRIBUTE_READONLY > 0 then
            SetFileAttributes(PChar(DocFileName),
                dwFileAttributes and Not FILE_ATTRIBUTE_READONLY);
    	if Assigned(FOnChecking) then FOnChecking(self);
        result := True;  { Resultat par defaut }
    end;  { ... fin de l'utilisation du module de donnees dossiers }
*)
end;

{----------------------------------------------------------------------}

function TDocumentDecisiv.Restore(aOwner: TComponent): Boolean;
var CodeSociete: string;
    DateCloture: TDateTime;
    Path: string;
    DocFileName: string;
    dwFileAttributes: Integer;
    NewFileName: string;
begin
	result := aOwner is TDossierDecisiv;  { Objet proprietaire ... }
    if not result then exit; { ... pas un TDossierDecisiv }
    if Assigned(FOnRestoring) then FOnRestoring(self);  { ... }
(*
    Path := DMDossiers.MakePathDoc+'\';
    DocFileName := DMDossiers.MakePathDoc+'\'+FileName;
    if FileExists(DocFileName) then begin
        dwFileAttributes := GetFileAttributes(PChar(DocFileName));
        if dwFileattributes and FILE_ATTRIBUTE_READONLY > 0 then
            SetFileAttributes(PChar(DocFileName),
                dwFileAttributes and Not FILE_ATTRIBUTE_READONLY);
            DeleteFile(DocFileName);
    end;
    NewFileName := TempPath+FileName;
    if FileExists(NewFileName) then begin
        dwFileAttributes := GetFileAttributes(PChar(NewFileName));
        if dwFileattributes and FILE_ATTRIBUTE_READONLY > 0 then
            SetFileAttributes(PChar(NewFileName),
                    dwFileAttributes and Not FILE_ATTRIBUTE_READONLY);
        Result := MoveFile(PChar(TempPath+FileName), PChar(DocFileName));
    end;
    if (Not Result) or (Not ISDossier) then Exit;
    try with DMDossiers do begin
            CodeSociete := TDossierDecisiv(aOwner).FCodeSociete;
            DateCloture := TDossierDecisiv(aOwner).FDateCloture;
            if Not DzDosDoc.FindKey([CodeSociete, DateCloture, UpperCase(FileName)]) then begin
                DzDosDoc.Insert;
                try DzDosDocCODE_SOCIETE.Text := CodeSociete;
                    DzDosDocDATE_CLOTURE_N.AsDateTime := DateCloture;
                    DzdosdocNOMDOC.Text := UpperCase(FileName);
                    DzDosDoc.Post; DzDosDoc.FlushBuffers;
                except on E: Exception do begin
                        ShowMsg($04010060, [FileName, E.Message]);
                        if DzDosDoc.State = dsInsert then DzDosDoc.Cancel;
                end end;
            end;
        end;
    except Result := False end;
*)
end;

{**********************************************************************}

{}
{ Procedures et fonctions privees }
{}

{----------------------------------------------------------------------}

function TDocumentDecisiv.GetDateFile: string;
begin
	result := DateTimeToString(FDateFile);
end;

{----------------------------------------------------------------------}

procedure TDocumentDecisiv.SetDateFile(Value: string);
begin
	try FDateFile := StrToDateTime(Value);
	except FDateFile := 0;
	end;  { Par defaut 30/12/1899 }
end;

{----------------------------------------------------------------------}

procedure TDocumentDecisiv.SetOnChecking(Value: TNotifyEvent);
begin
	FOnChecking := Value;
end;

{----------------------------------------------------------------------}

procedure TDocumentDecisiv.SetOnRestoring(Value: TNotifyEvent);
begin
    FOnRestoring := Value;  { Adresse evenement 'OnRestoring' }
end;

{**********************************************************************}
{**********************************************************************}

{}
{ Classe TDossierDecisiv - type dossier pour sauvegarde ou rappel
{}

constructor TDossierDecisiv.Create(aOwner: TComponent);
begin
	inherited Create(aOwner);  { Appel des constructeurs parents ... }
    FProduits := TList.Create; { Liste des produits, ... }
	FModules := TList.Create;  { ... modules, ... }
    FModeles := TList.Create;  { ... modeles, ... }
    FDocuments := TList.Create;  { ... documents associes, ... }
    FOrgDos := TList.Create;  { ... organisation dossier }
//	FDMOrganisation := TDMOrganisation.CreateTable(Nil);
end;

{----------------------------------------------------------------------}

destructor TDossierDecisiv.Destroy;
var	N	: Integer;
begin
	for N := FModeles.Count-1 downto 0 do  { ... }
    	TModeleDecisiv(FModeles[N]).Free;
    FModeles.Free; { ... destruction de la liste }

	inherited Destroy;  { Appel des destructeurs parents ... }
	FProduits.Free; { Liste des produits, ... }
	FModules.Free;  { ... modules, ... }
    FDocuments.Free;  { ... documents, ... }
    FOrgDos.Free; { ... et organisations du dossier }

//	FDMOrganisation.Free;
end;

{**********************************************************************}

{}
{ Procedures et fonctions publiques }
{}

function TDossierDecisiv.Check: Boolean;
var	N	: Integer;
begin
	try if Assigned(FOnChecking) then FOnChecking(self);  { Controles }
    	FLocalisation := LocationCompatibility(Args.UniteDossier);
    	result := True;  { Status, apres controles personnalises }
    except result := False; end;  { Probleme de controle personnalise }

	for N := 0 to FProduits.Count-1 do { Controle des produits ... }
    	if not TProduitDecisiv(FProduits[N]).Check(self) then
        	result := False;  { Code retour si probleme }

	for N := 0 to FModules.Count-1 do { Controle des modules ... }
    	if not TModuleDecisiv(FModules[N]).Check(self) then
        	result := False; { Code retour si probleme }

	for N := 0 to FModeles.Count-1 do { Controle des modeles ... }
    	if not TModeleDecisiv(FModeles[N]).Check(self) then
        	result := False; { Code retour si probleme }

	for N := 0 to FDocuments.Count-1 do { Controle des documents ... }
    	if not TDocumentDecisiv(FDocuments[N]).Check(self) then
        	result := False;  { Code retour si probleme }

	for N := 0 to FOrgDos.Count-1 do { Controle des organisations ... }
    	if not TOrgDosDecisiv(FOrgDos[N]).Check(self) then
        	result := False;  { Code retour si probleme }

    bModuleMdp := Assigned(FDossierDP);
    if not bModuleMdp then exit;
    for N := 0 to FModules.Count-1 do begin { Recherche du module }
        bModuleMdp := TModuleDecisiv(FModules[N]).CodeModule = MODULE_MDP;
        if bModuleMdp then break;
    end;
    if bModuleMdp and not CheckDP(FDossierDP) then result := False;
end;

{----------------------------------------------------------------------}
function TDossierDecisiv.Controle: Boolean;
var
	N				: Integer;
	Path			: string;
begin
    bbalances:=False;
    bdp:=False;
    Path := FLocalisation+FCodeSociete+'\'+  { Chemin du dossier ... }
    	FormatDateTime('yyyymmdd', FDateCloture);  { ! }
    result := True;
    for N := 0 to FModules.Count-1 do begin  { Rappel des modules ... }
        Case TModuleDecisiv(FModules[N]).FCodeModule of
        7:   begin
                 result:=URapSsDo.TestSousDossier(TempPath,bbalances,bdp);
             end;
        6:   begin
                 if Args.CodeModule=6 then
                 result:=URapSsDo.TestSousDossier(TempPath,bbalances,bdp);
             end;

        end;
        if not result then exit;  { Probleme ... }
    end;  { ... fin du rappel des modules de ce dossier }

end;
function TDossierDecisiv.Restore: Boolean;
var
    bExist, bDosNMaj: Boolean; { ... }
	N				: Integer;
	Path			: string;
    EtatDossier		: TEtatsDossier;
begin
    Path := FLocalisation+FCodeSociete+'\'+  { Chemin du dossier ... }
    	FormatDateTime('yyyymmdd', FDateCloture);  { ! }
(*
    if DMDossiers.CreatePath(Path) and  { ... }
    	DMDossiers.CreatePath(Path+'\Doc') then  { ... }
    else begin result := False; exit; end;  { Problemes de creation }
*)
    if Assigned(FOnRestoring) then FOnRestoring(self);
    result := True;
(*
	try with DMDossiers do begin  { Utilisation des donnees dossiers }
    	    if DZDOS.FindKey([FCodeSociete, FDateCloture]) then
        	    DZDOS.Edit  { Mise a jour, ou bien ... }
            else DZDOS.Insert;  { ... creation d'un dossier }

            DZDOSCODE_SOCIETE.Text := FCodeSociete;  { Cle d'acces ... }
            DZDOSDATE_CLOTURE_N.Value := FDateCloture; { ... }
            if Assigned(FOnRestoring) then FOnRestoring(self);
            DZDOS.Post; DZDOS.FlushBuffers;  { Mise a jour et ecriture }

			try EtatDossier := TEtatsDossier(  { Etat du dossier ... }
            	Smallint(DZDOSETAT_CLOTURE.AsInteger));
            except EtatDossier := []; end;
            bDosNMaj := False;  { Etat par defaut ... }
            bExist := DZSOC.FindKey([FCodeSociete]); { Recherche ... }

            if (edCloture in EtatDossier) then begin
            	if bExist then begin
                	if DZSOCDATE_CLOTURE.Value = DZDOSDATE_CLOTURE_N.Value then begin
                    	DZSOC.Edit; DZSOCDATE_CLOTURE.Clear;
                        DZSOC.Post; DZSOC.FlushBuffers;
                    end;
                end else begin
                	DZSOC.Insert; DZSOCCODE_SOCIETE.Text := FCodeSociete;
                    DZSOC.Post; DZSOC.FlushBuffers;
                end;
            end else bDosNMaj := not (edAltere in EtatDossier);

            if bDOSNMaj then begin
                if bExist then DZSOC.Edit
                else DZSOC.Insert;
                DZSOCCODE_SOCIETE.Text := FCodeSociete;
                DZSOCDATE_CLOTURE.Value := FDateCloture;
                DZSOC.Post; DZSOC.FlushBuffers; result := True;
            end;
        end
    except on E:EDatabaseError do begin  { Probleme BDE ... }
		ShowMsg($0401002B, ['DZDOS', E.Message]);
    	result := False; exit; { ... }
    end; end;  { ... fin de mise a jour dossier }
*)
    for N := 0 to FProduits.Count-1 do begin { Rappel des produits ... }
    	result := TProduitDecisiv(FProduits[N]).Restore(self);
        if not result then exit;  { Probleme ... }
    end;  { ... fin du rappel des produits de ce dossier }

    for N := 0 to FModules.Count-1 do begin  { Rappel des modules ... }
        Case TModuleDecisiv(FModules[N]).FCodeModule of
        7:   begin
                 result:=URapSsDo.RappelSousDossier(TempPath);
             end;
        6:   begin
                 if Args.CodeModule=6 then
                 result:=URapSsDo.RappelSousDossier(TempPath)
                 else
           	     result := TModuleDecisiv(FModules[N]).Restore(self);
             end;

        1:   begin
                 if bbalances then
                 begin
                     result:=URapSsDO.RappelEcritures(TempPath,'R');
                     if result then
             	     result := TModuleDecisiv(FModules[N]).Restore(self);
                 end
                 else
                 result:=URapSsDO.RappelEcritures(TempPath,'R');
              end;
        2..3: begin
                 if bbalances then
                 begin
             	     result := TModuleDecisiv(FModules[N]).Restore(self);
                 end
                 else
                 Result:=True;
              end;
        8:    begin
                 if bdp then
          	     result := TModuleDecisiv(FModules[N]).Restore(self)
                 else
                 result:=True;
              end;
        end;
        if not result then exit;  { Probleme ... }
    end;  { ... fin du rappel des modules de ce dossier }

    for N := 0 to FModeles.Count-1 do begin  { Rappel des modeles ... }
    	result := TModeleDecisiv(FModeles[N]).Restore(self);
        if not result then exit;  { Probleme ... }
    end;  { ... fin du rappel des modeles de ce dossier }

   	for N := 0 to FDocuments.Count-1 do begin  { Rappel documents ... }
   		result := TDocumentDecisiv(FDocuments[N]).Restore(self);
        if not result then exit;  { Code retour si probleme }
    end;  { ... fin du rappel des documents associes au dossier }

	{ Organisations non traitees dans ce cas }

    if bModuleMdp then Result := RestoreDP(FDossierDP);
end;

{----------------------------------------------------------------------}

function TDossierDecisiv.RestoreCount: Integer;
var	I, N		: Integer;
    bCountDP	: Boolean;
begin
	result := FProduits.Count+FModules.Count+FDocuments.Count+
                FOrgDos.Count+1; { ! }
    bCountDP := False;
    for N := 0 to FModules.Count-1 do begin { Tables des modules ... }
    	Inc(result, TModuleDecisiv(FModules[N]).RestoreCount);
        if not bCountDP then
            bCountDP := TModuleDecisiv(FModules[N]).CodeModule = MODULE_MDP;
    end;
    if Assigned(FDossierDP) and bCountDP then
        Inc(Result, FDossierDP.RestoreCount);
    for N := 0 to FModeles.Count-1 do begin  { ... modeles }
    	I := TModeleDecisiv(FModeles[N]).RestoreCount; { ... }
        if I > 0 then Inc(result, I+1); { ... modele a rappeler }
    end;  { ... fin de l'estimation du nombre de modeles a rappeler }
end;

{**********************************************************************}

{}
{ Procedures et fonctions privees }
{}

procedure TDossierDecisiv.GetChildren(Proc: TGetChildProc;
	Root: TComponent);  { ! }
var	N	: Integer;
begin
	inherited GetChildren(Proc, Root);  { Appel methodes parentes ... }
	if Root <> self then exit;  { Pas un 'TDossierDecisiv' ... }
	for N := 0 to FProduits.Count-1 do  { Produits ... }
		Proc(TProduitDecisiv(FProduits[N]));  { ... }
	for N := 0 to FModules.Count-1 do  { ... modules ... }
    	Proc(TModuleDecisiv(FModules[N]));  { ... puis modeles }
    for N := 0 to FModeles.Count-1 do Proc(TModeleDecisiv(FModeles[N]));
    for N := 0 to FDocuments.Count-1 do  { Documents associes ... }
    	Proc(TDocumentDecisiv(FDocuments[N]));  { ! }
    for N := 0 to FOrgDos.Count-1 do Proc(TOrgDosDecisiv(FOrgDos[N]));
    if Assigned(FDossierDP) then Proc(FDossierDP); { Dossier Permanent }
end;

{----------------------------------------------------------------------}

function TDossierDecisiv.GetDateCessation: string;
begin
	result := DateToString(FDateCessation);
end;

{----------------------------------------------------------------------}

function TDossierDecisiv.GetDateCloture: string;
begin
	result := DateToString(FDateCloture);
end;

{----------------------------------------------------------------------}

function TDossierDecisiv.GetDateClotureN1: string;
begin
	result := DateToString(FDateClotureN1);
end;

{----------------------------------------------------------------------}

function TDossierDecisiv.GetDateClotureN2: string;
begin
	result := DateToString(FDateClotureN2);
end;

{----------------------------------------------------------------------}

function TDossierDecisiv.GetDateClotureN3: string;
begin
	result := DateToString(FDateClotureN3);
end;

{----------------------------------------------------------------------}

function TDossierDecisiv.GetDateClotureN4: string;
begin
	result := DateToString(FDateClotureN4);
end;

{----------------------------------------------------------------------}

function TDossierDecisiv.GetDateCreation: string;
begin
	result := DateToString(FDateCreation);
end;

{----------------------------------------------------------------------}

function TDossierDecisiv.GetDateDebutEx: string;
begin
	result := DateToString(FDateDebutEx);
end;

{----------------------------------------------------------------------}

function TDossierDecisiv.GetDateGeneration: string;
begin
	result := DateToString(FDateGeneration);
end;

{----------------------------------------------------------------------}

function TDossierDecisiv.GetOrigineDate: string;
begin
	result := DateTimeToString(FOrigineDate);
end;

{----------------------------------------------------------------------}

function TDossierDecisiv.GetDateEffectiveCloture: string;
begin
	result := DateTimeToString(FDateEffectiveCloture);
end;

{----------------------------------------------------------------------}

procedure TDossierDecisiv.Notification(aComponent: TComponent;
	Operation: TOperation);  { ! }
var	List	: TList;
begin
(*    if (aComponent is TProduitDecisiv) then aComponent:=TProduitDecisiv(aComponent);
    if (aComponent is TModuleDecisiv) then aComponent:=TModuleDecisiv(aComponent);
    if (aComponent is TModeleDecisiv) then aComponent:=TModeleDecisiv(aComponent);
    if (aComponent is TOrgDosDecisiv) then aComponent:=TOrgDosDecisiv(aComponent);
    if (aComponent is TDocumentDecisiv) then aComponent:=TDocumentDecisiv(aComponent);

	if (aComponent is TProduitDecisiv) then
	inherited Notification(TProduitDecisiv(aComponent), Operation); List := nil; { ... }
	if (aComponent is TModuleDecisiv) then
	inherited Notification(TModuleDecisiv(aComponent), Operation); List := nil; { ... }
	if (aComponent is TModeleDecisiv) then
	inherited Notification(TModeleDecisiv(aComponent), Operation); List := nil; { ... }
    if (aComponent is TOrgDosDecisiv) then
	inherited Notification(TOrgDosDecisiv(aComponent), Operation); List := nil; { ... }
    if (aComponent is TDocumentDecisiv) then List := FDocuments;  { ! }
	inherited Notification(TDocumentDecisiv(aComponent), Operation); List := nil; { ... }
	if (aComponent is TDossierDPDecisiv) then
	inherited Notification(aComponent, Operation); List := nil; { ... }*)

	inherited Notification(TProduitDecisiv(aComponent), Operation); List := nil; { ... }

    if aComponent is TDossierDPDecisiv then begin
        if Operation = opInsert then
             FDossierDP := TDossierDPDecisiv(aComponent)
        else FDossierDP := Nil;
        Exit;
    end;

	if (aComponent is TProduitDecisiv) then List := FProduits;
	if (aComponent is TModuleDecisiv) then List := FModules;
	if (aComponent is TModeleDecisiv) then List := FModeles;
    if (aComponent is TOrgDosDecisiv) then List := FOrgDos; {!}
    if (aComponent is TDocumentDecisiv) then List := FDocuments;  { ! }

	if List = nil then exit;  { Type d'objet non reconnu ... }
    if Operation = opInsert then
    begin
(*        if (aComponent is TProduitDecisiv) then List.Add(TProduitDecisiv(aComponent));
        if (aComponent is TModuleDecisiv) then List.Add(TModuleDecisiv(aComponent));
        if (aComponent is TModeleDecisiv) then List.Add(TModeleDecisiv(aComponent));
        if (aComponent is TOrgDosDecisiv) then List.Add(TOrgDosDecisiv(aComponent));
        if (aComponent is TDocumentDecisiv) then List.Add(TDocumentDecisiv(aComponent));*)
        List.Add(aComponent);
    end
    else List.Delete(List.IndexOf(aComponent));  { opRemove }
end;

{----------------------------------------------------------------------}

procedure TDossierDecisiv.SetDateCessation(Value: string);
begin
	FDateCessation := StringToDate(Value);
end;

{----------------------------------------------------------------------}

procedure TDossierDecisiv.SetDateCloture(Value: string);
begin
	FDateCloture := StringToDate(Value);
end;

{----------------------------------------------------------------------}

procedure TDossierDecisiv.SetDateClotureN1(Value: string);
begin
	FDateClotureN1 := StringToDate(Value);
end;

{----------------------------------------------------------------------}

procedure TDossierDecisiv.SetDateClotureN2(Value: string);
begin
	FDateClotureN2 := StringToDate(Value);
end;

{----------------------------------------------------------------------}

procedure TDossierDecisiv.SetDateClotureN3(Value: string);
begin
	FDateClotureN3 := StringToDate(Value);
end;

{----------------------------------------------------------------------}

procedure TDossierDecisiv.SetDateClotureN4(Value: string);
begin
	FDateClotureN4 := StringToDate(Value);
end;

{----------------------------------------------------------------------}

procedure TDossierDecisiv.SetDateCreation(Value: string);
begin
	FDateCreation := StringToDate(Value);
end;

{----------------------------------------------------------------------}

procedure TDossierDecisiv.SetDateDebutEx(Value: string);
begin
	FDateDebutEx := StringToDate(Value);
end;

{----------------------------------------------------------------------}

procedure TDossierDecisiv.SetDateGeneration(Value: string);
begin
	FDateGeneration := StringToDate(Value);
end;

{----------------------------------------------------------------------}

procedure TDossierDecisiv.SetOrigineDate(Value: string);
begin
	FOrigineDate := StringToDateTime(Value);
end;

{----------------------------------------------------------------------}

procedure TDossierDecisiv.SetDateEffectiveCloture(Value: string);
begin
	FDateEffectiveCloture := StringToDateTime(Value);
end;

{----------------------------------------------------------------------}

procedure TDossierDecisiv.SetOnChecking(Value: TNotifyEvent);
var	N	: Integer;
begin
    for N := 0 to FProduits.Count-1 do { Liste des produits ... }
    	TProduitDecisiv(FProduits[N]).OnChecking := Value;
    for N := 0 to FModules.Count-1 do { ... module ... }
    	TModuleDecisiv(FModules[N]).OnChecking := Value;
    for N := 0 to FModeles.Count-1 do { ... modele ... }
    	TModeleDecisiv(FModeles[N]).OnChecking := Value;
    for N := 0 to FDocuments.Count-1 do { ... document }
    	TDocumentDecisiv(FDocuments[N]).OnChecking := Value;
    for N := 0 to FOrgDos.Count-1 do { ... document }
    	TOrgDosDecisiv(FOrgDos[N]).OnChecking := Value;
    if Assigned(FDossierDP) then
        FDossierDP.OnChecking := Value;
    FOnChecking := Value;  { Adresse evenement 'OnChecking' ... }
end;

{----------------------------------------------------------------------}

procedure TDossierDecisiv.SetOnRestoring(Value: TNotifyEvent);
var	N	: Integer;
begin
    for N := 0 to FProduits.Count-1 do { Liste des produits ... }
    	TProduitDecisiv(FProduits[N]).OnRestoring := Value;
    for N := 0 to FModules.Count-1 do { ... modules ... }
    	TModuleDecisiv(FModules[N]).OnRestoring := Value;
    for N := 0 to FOrgDos.Count - 1 do  { Organisation }
    	TOrgDosDecisiv(FOrgDos[N]).OnRestoring := Value;
    for N := 0 to FModeles.Count-1 do { ... modeles ... }
    	TModeleDecisiv(FModeles[N]).OnRestoring := Value;
    for N := 0 to FDocuments.Count-1 do  { ... documents }
    	TDocumentDecisiv(FDocuments[N]).OnRestoring := Value;
    if Assigned(FDossierDP) then FDossierDP.OnRestoring := Value;
    FOnRestoring := Value;  { Adresse de l'evenement 'OnRestoring' ... }
end;

{**********************************************************************}
{**********************************************************************}

{}
{ Classe TModeleDecisiv - type modele pour sauvegarde ou rappel
{}

constructor TModeleDecisiv.Create(aOwner: TComponent);
begin
	inherited Create(aOwner); { Appel des constructeurs parents ... }
    FTables := TList.Create;  { Liste des tables et ... }
    FDefaut := False; FModifiable := 1;  { ... }
    FCodeModele := -1; FCodeRubrique := -1; { ... codes }
end;

{----------------------------------------------------------------------}

destructor TModeleDecisiv.Destroy;
var
	CodeSociete	: string;  { ... }
    DateCloture	: TDateTime; { ! }
begin
	if (Owner is TDossierDecisiv) then begin  { Sauvegarde dossier ... }
    	CodeSociete := TDossierDecisiv(Owner).CodeSociete;  { ... }
        DateCloture := TDossierDecisiv(Owner).FDateCloture; { ! }
    end else begin CodeSociete := ''; DateCloture := 0; end;  { ! }
(*
    if Assigned(DMDossiers) then  { Deverrouillage rubrique/modele ... }
        DMDossiers.UnlockModel(CodeSociete, DateCloture, { Dossier }
            CodeRubrique, FCodeModele);  { Rubrique, et modele }
*)
    FTables.Free;  { Effacement de la liste des tables du modele }
	inherited Destroy;  { Appel des destructeurs parents, et enfants }
end;

{**********************************************************************}

{}
{ Procedures et fonctions publiques }
{}

function TModeleDecisiv.Check(aOwner: TComponent): Boolean;
var
	Code, Date		: string;  { ... }
	bExist, bFound	: Boolean; { ! }

	{- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }

    function GetModelCount: Integer;
//  var	No	: Smallint;
    begin
(*
    	with DMDossiers, DZDOSREF do begin  { Donnees modeles ... }
   			result := 0; No := DZREFCODE_MODELE.Value;
			IndexFieldNames := 'CODE_RUBRIQUE';
            SetRange([FCodeRubrique], [FCodeRubrique]);
            First;  { Lecture du premier modele de cette rubrique }

            while not EOF do begin { Parcours des modeles rubrique ... }
            	if ((DZDOSREFCODE_SOCIETE.Text <> Code) or  { ... }
                	(DZDOSREFDATE_CLOTURE_N.Text <> Date)) and
                    (DZDOSREFCODE_MODELE.Value = No) then
                	Inc(result);  { Correspondances ... }
            	Next; { Modele suivant pour la rubrique }
            end; { ... fin de parcours des modeles rubrique }
            CancelRange; IndexFieldNames := ''; { Initialisation }
        end; { ... fin de l'utilisation de la table de modeles dossier }
*)
		result := 0;  { Pas d'autres dossiers sur ce modele global ... }
    end;

	{- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }

    function  CheckModelTables: Boolean;
    var	N	: Integer;
    begin
		result := True;  { Code retour par defaut de la fonction ... }
    	for N := 0 to FTables.Count-1 do begin  { Parcours ... }
        	with TTableDecisiv(FTables[N]) do begin { ... }
        		result := Check(self);  { Controle de la table }
                if result then continue; { Table modele suivante ... }

				if GetLastError = I_E_FILEINUSE then  { Problemes ... }
                	WriteLogFile(I_E_MODELINUSE,  { Deja ouverte }
                    	['modèle '+FNomModele, FNomRubrique]);
                if GetLastError = I_W_NOTABLE then begin { ! }
                   	WriteLogFile(I_E_REFTABLE, [Name, FNomModele]);
                end; { ... fin de verification des principales erreurs }

				if (FNomModele = S_MODDOS) then continue;  { ... }

        		if (GetLastError = I_W_TBLCHANGED) and  { Contenu ... }
            		(GetModelCount > 0) then begin { ... et liens }
       				WriteLogFile(I_W_ASKFORMODEL, { Information }
                    	[FNomModele, FNomRubrique]);  { Message }
        			Include(FRestoreAction, raAskModel);  { ... }
            	end; break;  { ... puis, sortie de la procedure }
            end; { ... fin des controles de coherence d'une table }
        end;  { ... fin du controle des tables pour un modele existant }
    end;

	{- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }

	function  CheckModelVersion: Boolean;
//	var	Delta	: Integer;
    begin
(*
		with DMDossiers do begin { Utilisation des donnees dossier ... }
    		if bFound then Delta := DZDOSREFVERSION.Value  { ! }
        	else if bExist then Delta := DZREFVERSION.Value
        	else Delta := DMREFMODVERSION.Value;  { Rubriques }
        	Dec(Delta, FVersion); { Intervalle eventuel entre versions }

        	if Delta < 0 then begin  { Version modele incompatible ... }
        		WriteLogFile(I_E_REFVERSION,  { ... journalisation }
            		[FNomModele, DMREFMODLIBELLE_RUBRIQUE.Text]);
            	result := False; exit; { Bloquer le rappel dossier }
        	end; { ... fin de traitement des cas de version anterieure }

        	result := (Delta = 0);  { Equivalence ou compatibilite ... }
            if result then exit; { Versions de modele equivalentes }
            WriteLogFile(I_W_REFVERSION,  { Avertissement ... }
            	[FNomModele, DMREFMODLIBELLE_RUBRIQUE.Text]); { ! }
        end;  { ... fin de l'utilisation du module de donnees dossier }
*)
		result := True;  { Pas de difference de version dans ce cas ... }
    end;

	{- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }

	function IsErrorLevel: Boolean;
    begin
    	result := (GetLastError and $F000) = I_ERROR;
    end;

	{- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }

begin
(*
	with DMDossiers do begin { Utilisation des donnees de dossiers ... }
    	result := DMREFMOD.FindKey([FCodeRubrique]); { Lecture }
        if not result then begin  { Rubrique inexistante }
        	WriteLogFile(I_E_NORUBRIQUE, [FNomRubrique]); exit;
        end else if Assigned(FOnChecking) then FOnChecking(self);  { ! }

		FCodeModule := DMREFMODCODE_MODULE.Value; { Module/options ... }
        FOptions := TRubriqueOptions(DMREFMODOPTIONS.Value); { ! }
        FFieldNames := DMREFMODCHAMPS.Text;  { Si rubrique dynamique }

        if aOwner is TDossierDecisiv then begin  { Type de dossier ... }
        	with TDossierDecisiv(aOwner) do begin  { Cle ... }
            	Code := CodeSociete; Date := DateCloture;
            end;  { ... fin de recuperation des donnees }
           	bExist := FindKey([Code, Date]);  { Lecture ... }
        end else begin bExist := False; Code := ''; Date := ''; end;

       	bFound := DZDOSREF.FindKey([Code, Date, FCodeRubrique]); { ... }
        if not bExist and bFound then begin { Lien casse ... }
        	WriteLogFile(I_E_NOLINK, [FCodeRubrique]);
            result := False; exit; { Dans ce cas precis ... }
		end; { ... toujours preferer la fin de fonction a un plantage }

        if FNomModele = S_NONRET then begin  { Modele 'non retenu' ... }
        	Include(FRestoreAction, raDontCopy); exit;  { Non gere }
        end else bExist := ModelExist(FCodeRubrique, FNomModele);
        result := CheckModelVersion;  { Controle de la version ... }
        if not result and IsErrorLevel then exit; { Sortie si probleme }

        if FNomModele = S_MODDOS then begin { Cas 'modele dossier' ... }
        	if bFound then begin  { Modele de dossier resident ... }
            	if not LockModel(Code, StrToDate(Date),  { Cles }
                	FCodeRubrique, 0, mdlSupprimer) then begin
                    WriteLogFile(I_E_MODELINUSE,  { Probleme }
                    	[FNomModele, FNomRubrique]);  { !... }
                    result := False;  { Code fin de fonction }
                end else result := CheckModelTables; { Controle }
            end; exit;  { ... sortie, apres verification du modele }
        end; { ... fin de traitement d'un lien sur un 'modele dossier' }

        if not bExist then begin  { Pas de modele global de ce nom ... }
			if FTables.Count > 0 then begin { Ajout d'un modele ... }
       			WriteLogFile(I_W_ADDMODEL, { Niveau d'information }
           			[FNomModele, DMREFMODLIBELLE_RUBRIQUE.Text]);
				Include(FRestoreAction, raAddModel);  { Ajouter }
            end else begin  { ... lien au modele par defaut ... }
            	WriteLogFile(I_W_USEDEFAULT,  { En information  }
                	[DMREFMODLIBELLE_RUBRIQUE.Text]);  { Modele }
                Include(FRestoreAction, raUseDefault); { Defaut }
            end; { ... fin de la reflexion sur les actions modele }
            result := False; exit; { Sortie de la sous-fonction ... }
        end;  { ... apres envoi d'une notification, vers l'utilisateur }

        if DZREFMODIFIABLE.Value <> 1 then begin  { Non modifiable ... }
        	if DZREFMODIFIABLE.Value <> FModifiable then begin { ! }
            	WriteLogFile(I_W_REFREADONLY, { Fiche modifiee }
                   	[FNomModele,DMREFMODLIBELLE_RUBRIQUE.Text]);
                result := False; { Modele rendu non modifiable }
            end; Include(FRestoreAction, raDontCopy); exit;  { ! }
        end; { ... fin de controle de l'etat non modifiable du modele }

		if not LockModel('', 0, FCodeRubrique, { Verrou rubrique & ... }
        	FCodeModele, mdlSupprimer) then begin { ... code modele }
            WriteLogFile(I_W_MODELINUSE, [FNomModele,FNomRubrique]);
            Include(FRestoreAction, raDontCopy); { Ne pas rappeler }
            result := False; exit;  { Code retour & sortie fonction }
        end; { ... fin du processus de verrouillage d'un modele global }

        if bFound and (DZDOSREFCODE_MODELE.Value <>  { Lien trouve ... }
        	DZREFCODE_MODELE.Value) then begin { ... mais change }
       		WriteLogFile(I_W_REFCHANGED, { Lien modifie ... }
           		[DMREFMODLIBELLE_RUBRIQUE.Text, FNomModele]);
     		result := False; { Etat des controles du modele }
        end; { ... fin de controle du lien modele / dossier }
        if not CheckModelTables then result := False; { Controle }
    end; { ... fin d'utilisation des tables modeles du module dossiers }
*)
	if Assigned(FOnChecking) then FOnChecking(self);
	Include(FRestoreAction, raDontCopy);
	result := True;  { Pas d'erreur ou d'informations generees }
end;

{----------------------------------------------------------------------}

function TModeleDecisiv.Restore(aOwner: TComponent): Boolean;
var
	N		: Integer;
    Path	: string;

	{- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }

	function AttachModel: Boolean;
    begin
        with DMDossiers, TDossierDecisiv(aOwner) do begin  { Lien ... }
    		result := AttachModel(FCodeRubrique, GetModelValue,
            	CodeSociete, StrToDate(DateCloture), { ... }
                FParent); { ... etablissement du lien dossier }
        end;  { ... fin de la creation d'un lien modele - dossier }
    end;

	{- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }

	function Convert(const Text: string): Variant;
    begin
    	try result := Variant(StrToDate(Text));
        except result := Variant(Text);
        end;  { ... fin de la conversion }
    end;

	{- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }

    procedure DeleteTempFiles;
    var	N	: Integer;
    begin
    	for N := 0 to FTables.Count-1 do begin { Effacement tables ... }
    		DeleteFiles(TempPath+TTableDecisiv(FTables[N]).Name+'.*');
        end; { ... fin de l'effacement des tables de travail du modele }
    end;

	{- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }

	function GetIndexOf(Values: Variant): SmallInt;
    var
		List		: TStringList;
    	Count, N	: Integer;
    begin
    	if not VarIsArray(Values) then Count := 1  { Valeur simple ... }
        else Count := VarArrayHighBound(Values, 1)+1;  { Tableau }
        List := TStringList.Create;  { Creation d'une liste de travail }

		try with DMDossiers do  { Acces aux donnees des modeles ... }
        		for result := 0 to LookupTable.Count-1 do begin
					List.Clear; List.CommaText := LookUpTable[result];

                	if ((List.Count <= 0) or (List.Count <> Count)) then
                    	continue; { Nombre de valeurs incompatible }
					if Count > 1 then begin { Valeurs multiples }
                   		for N := 0 to Count-1 do { Parcours ... }
                       		if Values[N] <> Convert(List[N]) then
                           		break; { Pas de correspondance ... }
                	end else N := Ord(Values = Convert(List[0]));  { ! }

                	if N = Count then exit; { Index determine ... }
           		end;  { ... fin de recherche des valeurs }
        finally List.Free end; { Destruction de la liste }
		result := -1; { Valeur de l'index retourne par defaut ... }
    end;

	{- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }

	procedure GetDefaultModel;
	var
		Value	: Integer; { ... }
		Index	: Smallint;  { ! }
	begin
    	if roDynamique in FOptions then begin { Rubrique dynamique ... }
    		try Index := GetIndexOf(DMDossiers.DZDOS[FFieldNames]);
        	except Index := -1; end;  { Probleme de lecture }
    	end else Index := -1; { ... pour les rubriques statiques }
        Value := FCodeRubrique; FNomModele := S_NONRET;  { Par defaut }

		Include(FRestoreAction, raDontCopy); { Ne pas copier le modele }

		with DMDossiers do begin  { Acces aux donnees des modeles ... }
			if not DMREFDEF.FindKey([Value, { Code rubrique ... }
        		DefaultProduct[FCodeModule], Index]) then
            	DZREF.IndexFieldNames := 'CODE_RUBRIQUE' { ... }
            else Value := DMREFDEFCODE_MODELE.Value;  { ... ou modele }

            if DZREF.FindKey([Value]) then begin  { Modele global ... }
            	Value := DZREFCODE_MODELE.Value;  { Code modele }
            	FNomModele := DZREFNOM_MODELE.Text;  { Nom }
            end else Value := DMREFDEFCODE_MODELE.Value;  { ! }
            DZREF.IndexFieldNames := '';  { Retour a la cle primaire }

			case TDefaultOptions(DMREFDEFOPTIONS.Value-1) of  { ... }
				doModeleDossier	: begin  { Modele & parent ... }
                    FNomModele := S_MODDOS; FParent := Value;
                    Include(FRestoreAction, raCreateModel);
				end; { ... fin de gestion du modele dossier }
        	end; { ... fin de traitement des options du modele }
		end; { ... fin de recherche & acces aux donnees des modeles }
	end;

	{- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }

	function QueryModelAction: Boolean;
    var Action	: TModalResult;
    begin
(*
    	with TQueryModelActionDlg.Create(self) do begin { Dialogue ... }
        	try NomRubrique := FNomRubrique; { Nom de rubrique ... }
            	CodeRubrique := FCodeRubrique; { ... et numero }
            	NomModele := FNomModele;  { Nom du modele associe }
                Action := ShowModal;  { Affichage de la boite dialogue }

                case Action of  { Action a realiser sur le modele ... }
                	mrIgnore: Include(FRestoreAction, raDontCopy);
                	mrNo	: begin  { Ajout d'un modele ... }
						Include(FRestoreAction, raAddModel);
                    	FNomModele := NomModele; { Nom ... }
                    end; { ... fin du traitement de ce cas }
                end; { ... fin de traitement des decisions }
                result := Action <> mrCancel;  { Code retour }
            finally Free; end;  { ... destruction de ce dialogue }
        end;  { ... fin de la demande d'une decision de l'utilisateur }
*)
		Include(FRestoreAction, raDontCopy);
		result := True;
    end;

	{- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }

begin
	if not (raDontCopy in FRestoreAction) then { Pas de transfert }
    	if Assigned(FOnRestoring) then FOnRestoring(self);
    result := False;  { Code retour par defaut de la fonction }

	if (raAskModel in FRestoreAction) and not QueryModelAction then
    	exit;  { Rappel annule sur demande de l'utilisateur }
	if (raUseDefault in FRestoreAction) then GetDefaultModel; { ! }
(*
    if (raAddModel in FRestoreAction) then { Ajouter modele global ... }
    	with DMDossiers do begin { Utilisation des donnees modeles }
    		if AddGlobalModel(FCodeRubrique, FNomModele, { ... }
        		False, FModifiable, 0) <> mrOk then exit; {!}
        	try DZREF.Edit; { Description & commentaire ... }
           		DZREFDESCRIPTION.Text := FDescription;  { ! }
           		DZREFCOMMENTAIRE.Value := FCommentaire; { ! }
            	DZREF.Post; { ... mise a jour dans la table }
        	except on E:EDatabaseError do begin  { Probleme }
				ShowMsg($04010018, ['DZREF',E.Message]);  {!}
            	DZREF.Cancel;  { ... annulation mise a jour }
        	end end; { ... fin de traitement des problemes BDE }
        end;  { ... fin d'utilisation du module de donnees modeles }
    result := not (aOwner is TDossierDecisiv) or AttachModel;  { Liens }
*)
    result := True;  { On ne restore pas le modele ... }

    if result and (raCreateModel in FRestoreAction) then begin  { ... }
        result := DMDossiers.CreateModel( { Creation de tables }
        	TDossierDecisiv(aOwner).CodeSociete,  { Codes }
            StrToDate(TDossierDecisiv(aOwner).DateCloture),
            FCodeRubrique, GetModelValue, FParent);  { Modele }
	end; { ... fin de creation des tables d'un nouveau modele dossier }

	if not result or (raDontCopy in FRestoreAction) then begin
		DeleteTempFiles; exit;  { Nettoyage ... }
    end;  { ... pas de transfert - sortie vers l'appelant }

    if (raAddModel in FRestoreAction) and (FTables.Count > 0) then begin
    	Path := TTableDecisiv(FTables[0]).GetTablePath;  { Chemin }
        result := DMDossiers.CreatePath(ExtractFilePath(Path));
        if not result then begin DeleteTempFiles; exit; end;  { ! }
    end; { ... fin de la creation du chemin d'acces a un modele global }

    for N := 0 to FTables.Count-1 do begin  { Transfert des tables ... }
    	result := TTableDecisiv(FTables[N]).Restore(self);
        if not result then exit; { Probleme de transfert }
    end;  { ... fin du transfert des tables vers le modele specifie }
end;

{----------------------------------------------------------------------}

function TModeleDecisiv.RestoreCount: Integer;
begin
	if not (raDontCopy in FRestoreAction) then
		result := FTables.Count
	else result := 0;  { Tables non copiees }
end;

{**********************************************************************}

{}
{ Procedures et fonctions protegees }
{}

function TModeleDecisiv.GetChildOwner: TComponent;
begin
	result := self;
end;

{**********************************************************************}

{}
{ Procedures et fonctions privees }
{}

procedure TModeleDecisiv.GetChildren(Proc: TGetChildProc;
	Root: TComponent);  { ! }
var	N	: Integer;
begin
	inherited GetChildren(Proc, Root);  { Appel methode parente ... }
    for N := 0 to FTables.Count-1 do Proc(TTableDecisiv(FTables[N]));
end;

{----------------------------------------------------------------------}

function TModeleDecisiv.GetModelValue: Smallint;
begin
	result := -1;  { Modele 'non retenu' par defaut ... }
    if FNomModele = S_NONRET then exit;
    if FNomModele = S_MODDOS then begin Inc(result); exit end;

    with DMDossiers, DZREF do begin  { Lecture de la table modeles }
		IndexFieldNames := 'CODE_RUBRIQUE'; { Rubrique ... }
        SetRange([FCodeRubrique], [FCodeRubrique]);
        First;  { Lecture du premier modele de la rubrique }

        while not EOF do  { Parcours des modeles de rubrique ... }
        	if DZREFNOM_MODELE.Text = FNomModele then begin
            	result := DZREFCODE_MODELE.Value; break;
            end else Next; { Modele global suivant ... }
        CancelRange; IndexFieldNames := ''; { Initialiser }
    end;  { ... fin de l'utilisation de la table modeles globaux }
end;

{----------------------------------------------------------------------}

procedure TModeleDecisiv.Notification(aComponent: TComponent;
	Operation: TOperation);
begin
	inherited Notification(aComponent, Operation);  { Methode parente }
    if not (aComponent is TTableDecisiv) then exit; { ! }
    if Operation = opInsert then FTables.Add(aComponent)
    else FTables.Delete(FTables.IndexOf(aComponent));  { opRemove }
end;

{----------------------------------------------------------------------}

procedure TModeleDecisiv.SetNomModele(Value: string);
var	Bookmark	: TBookmark;
begin
(*
	Bookmark := DMDossiers.DZREF.GetBookmark;  { Probleme sinon }
	FNomModele := Value;  { Nom du modele ... }
    try FCodeModele := GetModelValue;
    except FCodeModele := -1; end;  { ! }
    DMDossiers.DZREF.GotoBookmark(Bookmark);  { ... }
*)
	FNomModele := Value;
    FCodeModele := -1;
end;

{----------------------------------------------------------------------}

procedure TModeleDecisiv.SetOnChecking(Value: TNotifyEvent);
var	N	: Integer;
begin
    for N := 0 to FTables.Count-1 do { Liste des tables ... }
    	TTableDecisiv(FTables[N]).OnChecking := Value;
    FOnChecking := Value;  { Adresse evenement 'OnChecking' }
end;

{----------------------------------------------------------------------}

procedure TModeleDecisiv.SetOnRestoring(Value: TNotifyEvent);
var	N	: Integer;
begin
    for N := 0 to FTables.Count-1 do  { Liste des tables ... }
    	TTableDecisiv(FTables[N]).OnRestoring := Value;
    FOnRestoring := Value; { Adresse evenement 'OnRestoring' }
end;

{**********************************************************************}
{**********************************************************************}

{}
{ Classe TModuleDecisiv - type module pour sauvegarde ou rappel
{}

constructor TModuleDecisiv.Create(aOwner: TComponent);
begin
	inherited Create(aOwner);  { Appel des constructeurs parents ... }
	FCodeModule := -1; FVersion[0] := 0; FVersion[1] := 0;
	FTables := TList.Create;  { Creation de la liste des tables }
end;

{----------------------------------------------------------------------}

destructor TModuleDecisiv.Destroy;
begin
	inherited Destroy;  { ... }
	FTables.Free;
end;

{**********************************************************************}

{}
{ Procedures et fonctions publiques }
{}

function TModuleDecisiv.Check(aOwner: TComponent): Boolean;
var	Delta, N	: Integer;

	{- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }

	function CheckFileVersion(sVersion: string): Integer;
    var Last	: array[0..1] of DWORD;
    begin
    	Last[0] := FVersion[0]; Last[1] := FVersion[1];  { Version ... }
        try SetVersion(sVersion); { Conversion interne ... }
			result := Integer(FVersion[0] - Last[0]);
            if result <> 0 then exit; { Differences }
            result := Integer(FVersion[1] - Last[1]);  { ! }
        finally FVersion[0] := Last[0]; FVersion[1] := Last[1]; end;
    end;

	{- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }

begin
	if Assigned(FOnChecking) then FOnChecking(self);  { ... }
    result := True;
(*
	with DMDossiers do begin { Utilisation des donnees de dossiers ... }
		if Assigned(FOnChecking) then FOnChecking(self);  { ... }
    	result := DMMOD.FindKey([FCodeModule]); { Recherche }
        if not result then begin  { Module non installe ... }
        	WriteLogFile(I_E_NOMODULE, [self.Name]); exit;  { ! }
        end;  { ... ecriture d'une entree log, puis sortie de fonction }

        Delta := CheckFileVersion(GetFileVersion(DMMODDLLMOD.Text));
        if Delta < 0 then begin { Version anterieure ... }
        	WriteLogFile(I_E_MODVERSION, [self.Name]);
            result := False; exit;  { ... erreur fatale }
        end;  { ... fin de gestion de version anterieure de module }

        if Delta > 0 then begin  { Compatibilite ascendante ... }
        	WriteLogFile(I_W_MODVERSION, [self.Name]);
            result := False;  { Code retour }
            LogFile.LoggingLevel := I_TABLE;  { ! }
        end;  { ... fin de gestion des compatibilites module }

		for N := 0 to FTables.Count-1 do  { Tables du module ... }
        	with TTableDecisiv(FTables[N]) do begin  { ... }
    			if Check(self) then continue;  { Table suivante ... }

				if (GetLastError = I_E_FILEINUSE) then begin  { ... }
                	WriteLogFile(I_E_FILEINUSE, [Name]);  { ! }
                    result := False; exit;  { ... code retour }
                end;  { ... fin du controle d'ouverture de la table }

                if ((GetLastError = I_W_NOTABLE) and (Delta = 0)) then
                   	WriteLogFile(I_E_MODTABLE, [Name, self.Name]);
               	result := False; { Code retour fonction ... }
    		end; { ... fin de controle des tables du module }
        if LogFile <> nil then LogFile.LoggingLevel := 0;  { ! }
    end;  { ... fin de l'utilisation des tables du module dossiers }
*)
end;

{----------------------------------------------------------------------}

function TModuleDecisiv.Restore(aOwner: TComponent): Boolean;
var	N	: Integer;
begin
	if Assigned(FOnRestoring) then FOnRestoring(self); { Evenement ... }
    result := True; { Initialisation du code retour par defaut }
	for N := 0 to FTables.Count-1 do begin  { Rappel ... }
    	result := TTableDecisiv(FTables[N]).Restore(self);
        if not result then exit;  { ... erreur lors du rappel }
    end; { ... fin du rappel des tables dossier pour cet objet module }
end;

{----------------------------------------------------------------------}

function TModuleDecisiv.RestoreCount: Integer;
begin
	result := FTables.Count;
end;

{**********************************************************************}

{}
{ Procedures et fonctions protegees }
{}

function TModuleDecisiv.GetChildOwner: TComponent;
begin
	result := self;
end;

{**********************************************************************}

{}
{ Procedures et fonctions privees }
{}

procedure TModuleDecisiv.GetChildren(Proc: TGetChildProc;
	Root: TComponent);  { ! }
var	N	: Integer;
begin
	inherited GetChildren(Proc, Root);  { Methodes parentes ... }
	if not (Root is TDossierDecisiv) then exit;  { ! }
	for N := 0 to FTables.Count-1 do Proc(TTableDecisiv(FTables[N]));
end;

{----------------------------------------------------------------------}

function TModuleDecisiv.GetVersion: string;
begin
	result := Format('%d.%d.%d.%d', [  { Numero de version ... }
    	FVersion[0] shr 16, FVersion[0] and $FFFF, { ... }
        FVersion[1] shr 16, FVersion[1] and $FFFF  { ... }
    ]);  { ... fin de la mise au format du numero de version }
end;

{----------------------------------------------------------------------}

procedure TModuleDecisiv.Notification(aComponent: TComponent;
	Operation: TOperation);
begin
	inherited Notification(aComponent, Operation);  { Methode parente }
    if not (aComponent is TTableDecisiv) then exit; { ! }
    if Operation = opInsert then FTables.Add(aComponent)
    else FTables.Delete(FTables.IndexOf(aComponent));  { opRemove }
end;

{----------------------------------------------------------------------}

procedure TModuleDecisiv.SetOnChecking(Value: TNotifyEvent);
var	N	: Integer;
begin
    for N := 0 to FTables.Count-1 do  { Liste des tables ... }
    	TTableDecisiv(FTables[N]).OnChecking := Value;
    FOnChecking := Value;  { Valeur evenement 'OnChecking' }
end;

{----------------------------------------------------------------------}

procedure TModuleDecisiv.SetOnRestoring(Value: TNotifyEvent);
var	N	: Integer;
begin
    for N := 0 to FTables.Count-1 do  { Liste des tables ... }
    	TTableDecisiv(FTables[N]).OnRestoring := Value;
    FOnRestoring := Value; { Adresse evenement 'OnRestoring' }
end;

{----------------------------------------------------------------------}

procedure TModuleDecisiv.SetVersion(Value: string);
var
	W	: array[0..3] of Word;
	S	: string;
    N	: Integer;  { ! }
begin
	for N := Low(W) to High(W) do W[N] := 0; N := 0; S := Value;  { ! }
    while (Pos('.', S) > 0) and (N < High(W)) do begin { ... }
    	try W[N] := StrToInt(Copy(S, 1, Pos('.', S)-1));
        except { Ne pas traiter cette exception } ; end;
        S := Copy(S, Pos('.', S)+1, MaxInt); Inc(N);  { ... }
    end; { ... fin de l'analyse du numero de version du module donne }

    if (Length(S) > 0) and (N <= High(W)) then  { Derniere partie ... }
    	try W[N] := StrToInt(S);  { ... No de construction }
        except { Ne pas traiter l'exception } ; end;
    FVersion[0] := (W[0] shl 16) + W[1]; { dwFileVersionMS }
    FVersion[1] := (W[2] shl 16) + W[3]; { dwFileVersionLS - version }
end;

{**********************************************************************}
{**********************************************************************}

{}
{ Classe TProduitDecisiv - type produit pour sauvegarde ou rappel
{}

constructor TProduitDecisiv.Create(aOwner: TComponent);
begin
	inherited Create(aOwner);  { Appel des constructeurs parents ... }
	FCodeProduit := -1; FVersion := '';  { ... }
end;

{**********************************************************************}

{}
{ Procedures et fonctions publiques }
{}

function TProduitDecisiv.Check(aOwner: TComponent): Boolean;
begin
(*
	with DMDossiers do begin  { Utilisation des donnees dossier ... }
    	result := DMPRO.FindKey([FCodeProduit]);  { Recherche }
        if not result then begin  { Pas installe ... }
        	WriteLogFile(I_E_NOPRODUCT, [self.Name]); exit;
        end;  { ... ecriture d'une entree log, puis fin de fonction }

    	{ Realiser ici le controle de la version du produit ... }
    	if Assigned(FOnChecking) then FOnChecking(self);
    end;  { ... fin de l'utilisation des tables du module dossiers }
*)
    if Assigned(FOnChecking) then FOnChecking(self);
    result := True;
end;

{----------------------------------------------------------------------}

function TProduitDecisiv.Restore(aOwner: TComponent): Boolean;
begin
	result := aOwner is TDossierDecisiv;  { Objet proprietaire ... }
    if not result then exit; { ... pas un TDossierDecisiv }
    if Assigned(FOnRestoring) then FOnRestoring(self);  { ... }
(*
	with DMDossiers, DZDOSPRO, TDossierDecisiv(aOwner) do begin  { ... }
    	Products[FCodeProduit] := True;  { Produit configure ... }
    	if FindKey([CodeSociete, DateCloture, FCodeProduit]) then exit;

		try Insert;  { Tentative d'insertion d'un nouveau produit ... }
           	DZDOSPROCODE_SOCIETE.Text := CodeSociete; { Societe }
            DZDOSPRODATE_CLOTURE.Text := DateCloture; { ! }
            DZDOSPROCODE_PRODUIT.Value := FCodeProduit;
            Post; FlushBuffers; { Ecriture reelle ... }
        except on E:EDatabaseError do begin  { Erreur }
        	result := False;  { Code retour si erreur }
			ShowMsg($0401002B, ['DZDOS', E.Message]); { ! }
        end; end; { ... fin du traitement des exceptions du BDE }
    end; { ... fin de l'ajout du produit dans le dossier proprietaire }
*)
end;

{**********************************************************************}

{}
{ Procedures et fonctions privees }
{}

procedure TProduitDecisiv.SetOnChecking(Value: TNotifyEvent);
begin
	FOnChecking := Value;
end;

{----------------------------------------------------------------------}

procedure TProduitDecisiv.SetOnRestoring(Value: TNotifyEvent);
begin
    FOnRestoring := Value;  { Adresse evenement 'OnRestoring' }
end;

{**********************************************************************}
{**********************************************************************}

{}
{ Classe TTableDecisiv - type table pour sauvegarde ou rappel
{}

constructor TTableDecisiv.Create(aOwner: TComponent);
begin
	inherited Create(aOwner);  { Constructeurs parents ... }
	FDateMaJ := 0; FVersion := -1; { Date de MaJ & version }
end;

{----------------------------------------------------------------------}

destructor TTableDecisiv.Destroy;
begin
	inherited Destroy;  { Destructeurs parents ... }
    DeleteFiles(TempPath+Name+'.*');
end;

{**********************************************************************}

{}
{ Procedures et fonctions publiques }
{}

function TTableDecisiv.Check(aOwner: TComponent): Boolean;
var
	bFound	: Boolean;
	Path	: string;
begin
(*
	with DMDossiers do begin  { Utilisation des donnees dossiers ... }
    	if Assigned(FOnChecking) then FOnChecking(self); { ... }
    	result := True; { Code retour fonction par defaut }
    	DMTBL.IndexFieldNames := 'NOM'; { Nom de la table }
    	bFound := DMTBL.FindKey([self.Name]); { Recherche ... }
        DMTBL.IndexFieldNames := ''; { ... retour a la cle primaire }

        if bFound then begin { Table trouvee dans le dictionnaire ... }
        	if DMTBLVERSION.Value < FVersion then begin  { ... }
            	WriteLogFile(I_E_TBLVERSION, [self.Name]);
                result := False; exit; { ... cas non traitable }
			end;  { ... fin du controle de compatibilite descendante }

        	if DMTBLVERSION.Value > FVersion then begin  { Version ... }
                result := False;  { Message d'avertissement ... }
            	WriteLogFile(I_W_TBLVERSION, [self.Name]);  { ! }
            end;  { ... fin de controle de la compatibilite ascendante }

            Path := GetTablePath;  { Chemin d'acces absolu de la table }
            if FilesInUse(Path) then begin  { Tables ouverte ? }
            	SetLastError(I_E_FILEINUSE);  { Oui ... }
                result := False; exit;  { Code retour & sortie }
            end; { ... fin du controle du mode d'ouverture d'une table }

            if GetFileAgeing(Path) <= FDateMaJ then exit;  { Dates ... }
            SetLastError(I_W_TBLCHANGED);  { Plus de messages }
        end else SetLastError(I_W_NOTABLE); { Non trouvee }
        result := False;  { Message d'information, ou erreur }
    end;  { ... fin de l'utilisation des donnees du module dossiers }
*)
	if Assigned(FOnChecking) then FOnChecking(self);
	result := True;
end;

{----------------------------------------------------------------------}

function TTableDecisiv.Restore(aOwner: TComponent): Boolean;
var
	SearchRec	: TSearchRec;
    Path		: string;
    Found		: Integer;
begin
    if Assigned(FOnRestoring) then FOnrestoring(self); { Evenement ... }
	Path := ExtractFilePath(GetTablePath);  { Chemin ... }
    result := DeleteFiles(Path+Name+'.*');  { Effacement }
	Found := FindFirst(TempPath+Name+'.*', faAnyFile, SearchRec);

    while Found = 0 do begin  { Boucle de transfert des fichiers ... }
    	result := MoveFile(PChar(TempPath+SearchRec.Name),
        	PChar(Path+SearchRec.Name));  { Copie }
        if not result then break; { Probleme }
        Found := FindNext(SearchRec);  { Suivant }
    end; { ... fin de la boucle de transfert de fichiers }
    FindClose(SearchRec);  { ... nettoyage du contexte de recherche }
end;

{**********************************************************************}

{}
{ Procedures et fonctions privees }
{}

function TTableDecisiv.GetDateMaJ: string;
begin
	result := FormatDateTime('dd/mm/yyyy hh:mm:ss', FDateMaJ);
end;

{----------------------------------------------------------------------}

function TTableDecisiv.GetTablePath: string;
var aOwner	: TComponent;
begin
	aOwner := Owner; result := '';  { Initialisation des variables ... }
    while aOwner <> nil do begin  { Recherche proprietaire ... }
    	if (aOwner is TDossierDecisiv) or  { Types ... }
        	(aOwner is TModeleDecisiv) then break;
        aOwner := aOwner.Owner;  { Type proprietaire }
    end;  { ... fin de recherche module/modele proprietaire }
	if aOwner = nil then exit; { Type proprietaire non module/modele }

    if (aOwner is TModeleDecisiv) and  { Cas modeles globaux ... }
    	(TModeleDecisiv(aOwner).Nom <> S_MODDOS) then
    	with TModeleDecisiv(aOwner) do begin
    		if Nom = S_NONRET then exit; {!}
            result := ModelsPath+Nom; { Modeles ... }
    	end  { ... fin de la reconstitution d'un chemin absolu }

    else begin  { Tables ou modeles locaux a un dossier ... }
    	while (aOwner <> nil) and  { Recherche ... }
    		not (aOwner is TDossierDecisiv) do
        	aOwner := aOwner.Owner; { Proprietaire ... }
		if aOwner = nil then exit;  { Aucun dossier proprietaire }

        with TDossierDecisiv(aOwner) do begin { Objet type dossier ... }
        	if Localisation = '' then exit; { Pas de chemin }
        	result := Localisation+FCodeSociete+'\'+  {!}
            	FormatDateTime('yyyymmdd', FDateCloture);
        end; { ... fin de traitement d'un objet dossier }
    end;  { ... fin de la constitution du chemin d'un type }
    result := result+'\'+Name+'.*'; { Nom des fichiers a ajouter }
end;

{----------------------------------------------------------------------}

procedure TTableDecisiv.SetDateMaJ(Value: string);
begin
	try FDateMaJ := StrToDateTime(Value);
	except FDateMaJ := 0;
	end;  { Par defaut 30/12/1899 }
end;

{----------------------------------------------------------------------}

procedure TTableDecisiv.SetOnChecking(Value: TNotifyEvent);
begin
	FOnChecking := Value;
end;

{----------------------------------------------------------------------}

procedure TTableDecisiv.SetOnRestoring(Value: TNotifyEvent);
begin
	FOnRestoring := Value;
end;

{**********************************************************************}
{**********************************************************************}

{}
{ Classe TLogFile - Classe fichier d'enregistrement
{}

constructor TLogFile.Create;
var LogFileName	: array[0..255] of Char;
Const S_ERRFILEMCT	: string = 'ERRMES.MCT';
begin
	if GetTempFileName(PChar(TempPath), 'LOG', 0, LogFileName) = 0 then
    	raise EInOutError.CreateFmt(SInOutError, [InOutRes]);
    FLoggingLevel := 0; FErrors := 0; FWarnings := 0;
    FLogFileName := StrPas(LogFileName);  { Nom ... }
	inherited Create(FLogFileName, fmCreate);  { Creation }
    FIniFile := TIniFile.Create(BinPath+S_ERRFILEMCT);  { Table messages }
end;

{----------------------------------------------------------------------}

destructor TLogFile.Destroy;
begin
    inherited Destroy;  { Destructeurs parents ... }
	FIniFile.Free;  { ... messages ... }
    if FLogFileName <> '' then DeleteFile(FLogFileName);
end;

{**********************************************************************}

{}
{ Procedures et fonctions publiques }
{}

function TLogFile.Write(MsgNo: Integer; Args: array of const): Longint;
var Buffer, S	: string;

	{- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }

    function  Level(MsgNo: Integer): Integer;
    begin
    	result := MsgNo and $0F00;
    end;

	{- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }

	function  Severity(MsgNo: Integer): Integer;
    begin
		if MsgNo < $1000 then begin  { Codes specifiques XceedZip ... }
			if MsgNo < XcdFirstError then result := I_WARNING
            else result := I_ERROR;  { Erreur / information }
    	end else result := MsgNo and $F000; { ... ou autres codes }
    end;

	{- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - }

begin
	result := -1;  { Code retour en cas d'operation non effectuee ... }
	case Severity(MsgNo) of  { Niveau de severite du message ... }
    	I_WARNING :	if Level(MsgNo) > FLoggingLevel then begin
        				Buffer := S_WARNING; Inc(FWarnings);
        			end else exit;  { Ne pas en tenir compte }
        I_ERROR   :	begin Buffer := S_ERROR; Inc(FErrors) end; {!}
    else exit end;  { Retour en cas de niveau d'erreur non interprete }

	S := FIniFile.ReadString(Buffer, IntToHex(MsgNo, 4), S_UNKNOWN);
    Buffer := Format('%s: %s'+CRLF, [Buffer, S]);  { Format }
    SetLastError(MsgNo);  { Derniere erreur consignee ... }
    if S = S_UNKNOWN then Buffer := Format(Buffer, [MsgNo])
    else Buffer := Format(Buffer, Args);  { Parametrage ... }
    result := inherited Write(Pointer(Buffer)^, Length(Buffer));
end;

{**********************************************************************}
{**********************************************************************}

{}
{ Classe TStep - Type phase visuelle de sauvegarde ou rappel
{}

constructor TStep.Create(aImage: TImage; aLabel: TLabel);
begin
    FImage := aImage;
	FLabel := aLabel;
    FState := ssDone;
    SetState(ssNone);
end;

{**********************************************************************}

{}
{ Procedures et fonctions privees }
{}

procedure TStep.SetState(Value: TStepState);
begin
	if Value <> FState then FState := Value else exit;  { ! }
    case FState of  { Selon l'etat de la phase ... }
    	ssNone: begin { Etat indefini ... }
            FImage.Visible := False;
        end;  { ... fin d'etat indefini }

        ssExec: begin  { Traitement de l'etat 'en execution' ... }
        	with FLabel.Font do Style := Style + [fsBold];
            FImage.Visible := False; { Image non visible }
        end;  { ... fin de traitement de l'etat 'en execution' }

        ssDone: begin  { Traitement de la phase 'termine' ... }
        	with FLabel.Font do Style := Style - [fsBold];
            FImage.Visible := True;  { Image visible }
        end; { ... fin de traitement de l'etat 'termine' }
    end; { ... fin de traitement des differents etats d'une phase }

    Application.ProcessMessages;  { Rafraichissement de l'ecran }
end;

{**********************************************************************}
{**********************************************************************}

{}
{ Initialisation et terminaison
{}

initialization	GetPaths;

{**********************************************************************}

end.
