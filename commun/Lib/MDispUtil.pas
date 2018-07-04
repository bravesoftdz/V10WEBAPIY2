unit MDispUtil;

// majstruc umajstruct

interface

// ubob

uses Forms, dialogs
{$IFDEF VER150}
  , Classes
{$ENDIF}
	,BTMajStruct
  ,uBTPVerrouilleDossier
//  ,RegenEcheCon
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
  {MajHalley,} CreerSoc, CopySoc, Reseau, Mullog, LicUtil, UtomUtilisat, UserGrp_tom,
  UtilPGI, {AssistInitSoc,} PGIExec, Reindex, Paramsoc, Ent1,
  UtilSoc, {Exercice,} Souche_tom, FichComm, Pays, Region, CodePost, {Devise_TOM, } Tva, {MulJal,}
  //MulGene, MulTiers, MulSecti, CpAxe_Tom, CPStructure_tof,  VentType, CodeAFB,
  //EtbMce, OuvreExo, {CalcEuro,} EuroPGI, SuprGene, SuprJal, QRJDivis, EntPGI,
  {UtofExportConfidentialite, }Confidentialite_tof, {MajSocParLot_TOF, LGCOMPTE_TOF, CORRESP_TOM,} {STKMoulinette,}
  {MajHalleyUtil,} ubob, {YYBUNDLE_TOF, uMultiDossier, TomProfilUser,BackupRestore_TOF,traduc,AccesPortail_TOF, MajEnRafale}
  {,UtomEtabliss,}ChangeVersions,BobGestion,URegenVues
//  ,UFicheImportExport
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
        if (V_PGI.MajStructAuto) and (V_PGI.RunWithParams) then
        begin
          FMenuG.ForceClose := True;
          PostMessage(FMenuG.Handle, WM_CLOSE, 0, 0);
          Application.ProcessMessages;
          Exit;
        end else
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
          Ok1 := (BOB_IMPORT_PCL_STD('UTILS2009','Business',False)=1);
          if (Ok1) or (Ok2) then
          begin
          	TraitementsAfterImportBOB (VersionBase);
            RegenereVues;
						PGIInfo('Mise à jour terminée.#13#10 Une sortie du produit est necéssaire.');
            FMenuG.Quitter;
						Abort;
          end;
          
{$ENDIF}
          DispatchAuto;   // Ajout CA - 05/10/2005
          VerifSuperviseur;
          InitCodeProduit;
        end;
      end;

    11: ; //Après deconnection
    12: ;  // permet de faire qques chose avant connexion et seria
    13: ;
    15: ;
    16: ;
    100: ; // executer depuis le lanceur

    3220: ReindexSoc;
    4010: AGLLanceFiche('YY', 'YYIMPORTOBJET', '', '', ''); // Import d'objets
    4020: AGLLanceFiche('YY', 'YYMYINDEXES', '', '', ''); // IIndexes supplémentaires

    60201 : FicheUserGrp ;                   // groupe utilisateurs
    60202 : BEGIN FicheUSer(V_PGI.User) ; ControleUsers ; END ;   // Utilisateurs
    60208 : GCLanceFiche_Confidentialite( 'YY','YYCONFIDENTIALITE','','','26;27;60;92;111;145;146;147;148;149;150;160;279;283;284;285;304;319;320;321;322;323;325;327;328;329;331');
    60301: AGLLanceFiche('YY', 'YYJNALEVENT', '', '', ''); // Journal evenements
    // EPURATIONS
    60651 : BEGIN
      if V_PGI.Superviseur then
      begin
        AGLLanceFiche('BTP', 'BTFERMEPIECE_MUL', '', '', ''); // Fermeture/reouverture de pièces
      end else
      begin
        PgiInfo ('Vous n''avez pas les droits nécessaires pour effectuer cette opération');
        exit;
      end;
    END;
    60652 : BEGIN
      if V_PGI.Superviseur then
      begin
        AGLLanceFiche('BTP', 'BTEPUREPIECE_MUL', '', '', ''); // 2puration de pièces
      end else
      begin
        PgiInfo ('Vous n''avez pas les droits nécessaires pour effectuer cette opération');
        exit;
      end;
    END;
    60653 : BEGIN
      if V_PGI.Superviseur then
      begin
        AGLLanceFiche('BTP', 'BTEPUREAFF_MUL', 'AFF_AFFAIRE0=A;AFF_ETATAFFAIRE=TER', '', ''); // 2puration de pièces
      end else
      begin
        PgiInfo ('Vous n''avez pas les droits nécessaires pour effectuer cette opération');
        exit;
      end;
    END;
    60654 : BEGIN
      if V_PGI.Superviseur then
      begin
        AGLLanceFiche('BTP','BTCLOTUREAFF_MUL','AFF_AFFAIRE0=A;AFF_ETATAFFAIRE=ACP','','');  // Clôture de chantier
      end else
      begin
        PgiInfo ('Vous n''avez pas les droits nécessaires pour effectuer cette opération');
        exit;
      end;
    END;																																																 // Antériorite
    60655 : BEGIN
      if V_PGI.Superviseur then
      begin
        AGLLanceFiche('BTP','BTCLOTUREAFF_MUL','AFF_AFFAIRE0=A;AFF_ETATAFFAIRE=TER','','REACTIVE');  // Réactivation de chantier
      end else
      begin
        PgiInfo ('Vous n''avez pas les droits nécessaires pour effectuer cette opération');
        exit;
      end;
    END;																																																 // Antériorite
    60681 : BEGIN
      if not V_PGI.Superviseur then
      begin
        PgiInfo ('Vous n''avez pas les droits nécessaires pour effectuer cette opération');
        exit;
      end;
      AGLLanceFiche('BTP', 'BTEPUREPIECE_MUL', '', '', 'SUPFACTURE;SANSCONTROLE'); // Epuration de factures sans controle
    END;
    60682 : BEGIN
      if not V_PGI.Superviseur then
      begin
        PgiInfo ('Vous n''avez pas les droits nécessaires pour effectuer cette opération');
        exit;
      end;
      AGLLanceFiche('BTP','BTSELFAC_MUL','AFF_ETATAFFAIRE=ACP','','MODIFICATION;STATUT=AFF;REAJUSTEAVANC') ;
    END;
    60683 : BEGIN
      if not V_PGI.Superviseur then
      begin
        PgiInfo ('Vous n''avez pas les droits nécessaires pour effectuer cette opération');
        exit;
      end;
      AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=FBT;GP_VENTEACHAT=VEN','','MODIFICATION;STATUT=AFF;REAJUSTESIT');
    END;
    60684 : BEGIN
      if not V_PGI.Superviseur then
      begin
        PgiInfo ('Vous n''avez pas les droits nécessaires pour effectuer cette opération');
        exit;
      end;
      AGLLanceFiche('BTP','BTREAJPIECE_MUL','GP_NATUREPIECEG=FBT;GP_VENTEACHAT=VEN','','MODIFICATION;STATUT=AFF');
    END;
    65550 : RegenereVues ; // Regénération des vues
    279101:  BEGIN
      if not V_PGI.Superviseur then
      begin
        PgiInfo ('Vous n''avez pas les droits nécessaires pour effectuer cette opération');
        exit;
      end;
      AGLLanceFiche('BTP','BTPIECEVEN_MUL','GP_NATUREPIECEG=BCE','','CONSULTATION');
    END;
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
var
  VireCompta: Boolean;
begin
  V_PGI.VersionDemo := False;
  VireCompta := True;
  V_PGI.LaSerie := S5;
  if ctxPCL in V_PGI.PGIContexte then
  begin
    // maintenant qu'on peut le lancer en direct
  end else
  begin
    // PCS 13/10/2003 modules compta toujours présents.
    VireCompta := False;
  end;
  //
	FMenuG.RemoveGroup(-7211, True);
  FMenuG.RemoveItem(1105);
  FMenuG.RemoveItem(7178);

  if VireCompta then
  begin
    FMenuG.RemoveGroup(-3200, True);
    FMenuG.RemoveGroup(-1705, True);
    FMenuG.RemoveGroup(-1105, True);
    FMenuG.RemoveGroup(-7211, True);
    FMenuG.RemoveGroup(-1104, True);
    FMenuG.RemoveItem(3172);
    FMenuG.RemoveItem(3185);
    FMenuG.RemoveItem(7718);
  end;
  UpdateSeries;
  //if not FileExists(ExtractFilePath(Application.ExeName) + 'INI_E26BCFEE-5C71-498E-BF4C-B5B3400A2CCE.ini') then
  if not V_PGI.DossierPCL then
    FMenuG.RemoveItem(4005);
  // virer paramétrage multi soc si PCL
  if (V_PGI.ModePCL = '1') then
  begin
    FMenuG.RemoveItem(4041);
    FMenuG.RemoveItem(4042);
  end;
  if isOracle then
  begin
    FMenuG.RemoveItem(3131);
    FMenuG.RemoveItem(3132);
  end;
end;


procedure InitApplication;
begin
  FMenuG.OnDispatch := Dispatch;
  FMenuG.OnChargeMag := ChargeMagHalley;
  FMenuG.OnChangeModule := AfterChangeModule;
  FMenuG.SetModules([279], [79]);
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
  NomHalley := 'UTILSBTP';

  V_PGI.NumVersionBase := 998 ;
  TitreHalley := 'Utilitaires BTP    Base ' + IntToStr(V_PGI.NumVersionBase);
  HalSocIni := 'CEGIDPGI.ini';

  Copyright := '© Copyright ' + Apalatys;
  V_PGI.NumVersion := '10.0' ;
  V_PGI.NumBuild := IntToStr(V_PGI.NumVersionBase) + '.155';
  V_PGI.DateVersion := EncodeDate(2018, 05, 04) ;

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
  InitialisationVPGI;

end.

