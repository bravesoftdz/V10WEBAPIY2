unit MDispACeg;

// majstruc umajstruct

interface

// ubob

uses Forms, dialogs
{$IFDEF VER150}
  , Classes
{$ENDIF}
	,BTMajStruct
  ,uBTPVerrouilleDossier
  ,RegenEcheCon
  ,UApplication
  ;

procedure InitApplication;

type
  Tgctie = class(TDatamodule)
    procedure FMDispADMCreate(Sender: TObject);
  end;

var
  gctie: Tgctie;

implementation

{$R *.DFM}

uses windows, HEnt1, MenuOLG, sysutils, Messages, HMsgBox, {$IFNDEF VER150}Classes, {$ENDIF}Controls, Hctrls,
  HPanel, UIUtil, ImgList, Tablette, FE_Main, EdtEtat, AglInitPlus, AGLInit,
  MajHalley, CreerSoc, CopySoc, Reseau, Mullog, LicUtil, UtomUtilisat, UserGrp_tom,
  UtilPGI, AssistInitSoc, PGIExec, Reindex, Paramsoc, Ent1,
  UtilSoc, Exercice, Souche_tom, FichComm, Pays, Region, CodePost, Devise_TOM,  Tva, MulJal,
  MulGene, MulTiers, MulSecti, CpAxe_Tom, CPStructure_tof,  VentType, CodeAFB,
  EtbMce, OuvreExo, {CalcEuro, EuroPGI,} SuprGene, SuprJal, QRJDivis, EntPGI,
  UtofExportConfidentialite, Confidentialite_tof, MajSocParLot_TOF, LGCOMPTE_TOF, CORRESP_TOM, {STKMoulinette,}
  MajHalleyUtil, ubob, YYBUNDLE_TOF, uMultiDossier, TomProfilUser,BackupRestore_TOF,traduc,AccesPortail_TOF, MajEnRafale
  ,UtomEtabliss,ChangeVersions,BobGestion,URegenVues

{$IFDEF INTERNAT}
  , TOMIdentificationBAncaire
{$ENDIF}
  ;

function VerifSuperviseur: boolean;
begin
  Result := True;
  if not V_PGI.Superviseur then
  begin
    PGIBox('Vous n''avez pas les droits pour accéder à cette application', 'Attention');
    Result := False;
    FMenuG.ForceClose := True;
    FMenuG.Close;
  end;
end;

function VireMoiCa: string;
begin
  Result := 'SCO_CONSOLIDATION;SCO_SYNAPTIQUE;SCO_MCC;SCO_FISCAPERSO;SCO_TDI;SCO_XISOFLEX;SCO_HRCHR;SCO_KTTRANSPORT;SCO_GPAO';

  //Gisèle MERIEUX Le 29/05/2008 Version 8.1.850.86 Demande N° 2587
  Result := Result + ';SCO_AFFAIRE';
end;

procedure InitCodeProduit;
var
  ii, i: integer;
  stl: HTStringList;
  st, stCombos: string;
begin
  stCombos := 'YYLISTEPRODUITS';
  ii := TTToNum(stCombos);
  st := '';
  if ii > 0 then
  begin
    if (V_PGI.DECombos[ii].Valeurs = nil) then RemplirListe(stCombos, '');
    stl := V_PGI.DeCombos[ii].valeurs;
    if stl = nil then Exit;
    for i := 0 to stl.count - 1 do
    begin
      if V_PGI.CodeProduit <> '' then st := ';';
      V_PGI.CodeProduit := V_PGI.Codeproduit + st + ExtractInfo(stl[i], 1);
    end;
  end;
end;

{$IFDEF MAJBOB}
{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 08/06/2005
Modifié le ... :   /  /
Description .. : fonction qui indique si PgiMajVer est en version MAJBOB
Mots clefs ... :
*****************************************************************}
function AglOkMajBob(var BobName: string): boolean;
var
  i: integer;
  s: string;
begin
  // Par défaut
  BobName := '';

  // contrôle des paramètres
  for i := 0 to ParamCount do
  begin
    s := UpperCase(ParamStr(i));
    if Copy(s, 1, 5) = '/BOB=' then
    begin
      BobName := s;
      ReadTokenPipe(BobName, '=');
      BobName := Trim(BobName);
      Break;
    end;
  end;

  // Le nom du fichier doit être rempli, sinon => rien
  Result := BobName <> '';
end;
{$ENDIF}

{$IFDEF MAJPCL}

// c:\pgi00\app\PGIMAJVER.EXE /USER=CEGID /PASSWORD=46171FC5C3 /DATE=14/09/2005
//    /DOSSIER=DBBLM001@LOCAL_DB0 /MODEPCL=1 /MAJSTRUCTURE=TRUE /MAJSTRUCTAUTO=TRUE /DOSSIERPCL=TRUE
function OkDossierPCL : boolean;
var
  i: integer;
  s: string;
  StDossierPCL : string;
begin
  Result:=false;
  for i := 0 to ParamCount do
  begin
    s := UpperCase(ParamStr(i));
    if Copy(s, 1, 11) = '/DOSSIERPCL' then
    begin
      StDossierPCL := s;
      ReadTokenPipe(stDossierPCL, '=');
      Result := ( stDossierPCL='TRUE' );
      Break;
    end;
  end;
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 05/10/2005
Modifié le ... :   /  /    
Description .. : procédure de lancement de tâche à la connexion à 
Suite ........ : l'application.
Suite ........ : - /OPTIM=PCL permet de lancer la migration vers une base 
Suite ........ : allégée
Mots clefs ... : 
*****************************************************************}
procedure DispatchAuto;
var
  i : integer;
  St, stNom, stValue : string;
begin
  for i:=1 to ParamCount do
  begin
    St := ParamStr(i) ;
    stNom:=UpperCase(Trim(ReadTokenPipe(St,'='))) ;
    stValue:=UpperCase(Trim(St)) ;

    if stNom = '/OPTIM' then
    begin
      if ((V_PGI.DossierPCL) and (UpperCase( stValue ) = 'PCL')) then
      begin
{$IFDEF MAJPCL}
        OptimiseDossierPCL;
{$ENDIF}
        Application.ProcessMessages;
        FMenuG.ForceClose := True ;
        FMenuG.Close ;
        Exit;
      end;
    end;
  end;
end;

procedure Dispatch(Num: Integer; PRien: THPanel; var retourforce, sortiehalley: boolean);
var
{$IFDEF MAJBOB}
  stBobDeMiseAJour: string;
{$ENDIF}
  CodeRetour : integer;
  stMessage, Versionbase : string;
  Ok1,Ok2 : Boolean;
begin
  case Num of
    10:
      begin
(*
        if (V_PGI.MajStructAuto) and (V_PGI.RunWithParams) then
        begin
          FMenuG.ForceClose := True;
          PostMessage(FMenuG.Handle, WM_CLOSE, 0, 0);
          Application.ProcessMessages;
          Exit;
        end
        else
        begin
{$IFDEF BTP}
          if (not ISUtilisable) then
          begin
             stMessage := 'Base de données momentanément indisponible (mise à niveau en cours) ';
             stMessage := stMessage +#13#10 +'Veuillez réessayer ultérieurement';
             PGIInfo(stMessage,'Information de connection');
             FMenuG.ForceClose := True;
             PostMessage(FMenuG.Handle, WM_CLOSE, 0, 0);
             Application.ProcessMessages;
             Exit;
          end;
          VersionBase := RecupVersionbase;
          //
          Coderetour := MajStructure;
          if CodeRetour < 0 then
          begin
      			RetourForce:=True;
            Fmenug.Quitter;
            Exit;
          end;
          //
          TraiteChangementVersions (VersionBase);
          TraitementsBeforeImportBOB (VersionBase);
					// TraiteChangementGamme;
          // PGI_IMPORT_BOB('BAT2008'); ancienne gestion
          Ok1 := (BOB_IMPORT_PCL_STD('BAT2009','Business',False)=1);
          Ok2 := (BOB_IMPORT_PCL_STD('MAJERPCEGIDV9','CEGID',False)=1);
          if (Ok1) or (Ok2) then
          begin
          	TraitementsAfterImportBOB (VersionBase);
						PGIInfo('Mise à jour terminée.#13#10 Une sortie du produit est necéssaire.');
            FMenuG.Quitter;
						Abort;
          end;

{$ENDIF}
          DispatchAuto;   // Ajout CA - 05/10/2005
          VerifSuperviseur;
          InitCodeProduit;
        end;
        MiseAJourEnRafale;
        *)
      end;

    11: ; //Après deconnection
    12:   // permet de faire qques chose avant connexion et seria
      begin
      end;
    13: ;
    15: ;
    16: BEGIN
    		END;
    100: ; // executer depuis le lanceur
    3145: RecopieSoc(PRien, True);
    3220: ReindexSoc;
    {Sociétés et dossiers}
  else HShowMessage('2;?caption?;' + TraduireMemoire('Fonction non disponible : ') + ';W;O;O;O;', TitreHalley, IntToStr(Num));
  end;
end;

procedure DispatchTT(Num: Integer; Action: TActionFiche; Lequel, TT, Range: string);
begin
  case Num of
    1: ;
  end;
end;

procedure AfterChangeModule(NumModule: integer);
begin
end;


procedure InitApplication;
begin
  FMenuG.OnMajAvant := nil;
  FMenuG.OnMajApres := nil;
  FMenuG.IsMajPossible := nil;
  FMenuG.OnMajPendant := nil;
  FMenuG.OnDispatch := Dispatch;
  FMenuG.OnChargeMag := ChargeMagHalley;
  FMenuG.OnChangeModule := AfterChangeModule;
  FMenuG.SetModules([4], [49]);
  FMenuG.bSeria.Visible := False;
  V_PGI.DispatchTT := DispatchTT;
  VH^.OkModCompta := True;
  VH^.OkModBudget := True;
  VH^.OkModImmo := True;
  VH^.OkModEtebac := True;
end;

procedure Tgctie.FMDispADMCreate(Sender: TObject);
begin
  PGIAppAlone := True;
  CreatePGIApp;
end;

procedure InitialisationVPGI;
begin

  Apalatys := 'LSE';
  NomHalley := 'MAJVERBTP';

  V_PGI.NumVersionBase := 998 ;
  TitreHalley := 'Administration sociétés BTP    Base ' + IntToStr(V_PGI.NumVersionBase);
  HalSocIni := 'CEGIDPGI.ini';

  Copyright := '© Copyright ' + Apalatys;
  V_PGI.NumVersion := '9.1' ;
  V_PGI.NumBuild := IntToStr(V_PGI.NumVersionBase) + '.83';
  V_PGI.DateVersion := EncodeDate(2015, 04, 28) ;

  V_PGI.LaSerie := S5;

  V_PGI.OutLook := TRUE;
  V_PGI.OfficeMsg := TRUE;
  V_PGI.ToolsBarRight := TRUE;
  ChargeXuelib;

  // pour forcer mode SAV en Qualité
  if FileExists(ExtractFilePath(Application.ExeName) + 'INI_4B8D120B-1877-4E78-A817-331A23A35209.ini') then V_PGI.SAV := True;

  {$IFDEF XAVIER}
  V_PGI.SAV := True ;
  V_PGI.Debug := True ;
  V_PGI.DebugSQL := True ;
  {$ENDIF}

  V_PGI.AlterTable := True;
  V_PGI.VersionDemo := TRUE;
  V_PGI.VersionReseau := true;
  V_PGI.ImpMatrix := True;
  V_PGI.OKOuvert := FALSE;
  V_PGI.Halley := TRUE;
  V_PGI.MenuCourant := 0;
  //V_PGI.Decla100:=TRUE;

  V_PGI.RazForme := TRUE;



  V_PGI.StandardSurDP := True;
  V_PGI.MajPredefini := True;
  V_PGI.CegidApalatys := False;
  V_PGI.CegidBureau := True;
  V_PGI.IsPgiMajVer := True;
  // V_PGI.MajStructAuto := true ;

  v_pgi.enableTableToView:=False;
  V_PGI.RazForme:=TRUE;
	GetInfoMajApplication;

end;

initialization
  ProcChargeV_PGI := InitialisationVPGI;

end.

