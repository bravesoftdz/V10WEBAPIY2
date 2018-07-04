unit MDispAdm;

interface

uses Forms;

procedure InitApplication;

type
  Tgctie = class(TDatamodule)
    procedure FMDispADMCreate(Sender: TObject);
  end;

var gctie: Tgctie;

implementation

{$R *.DFM}

uses windows, HEnt1, MenuOLG, sysutils, Messages, HMsgBox, Classes, Controls,Hctrls,
  HPanel, UIUtil, ImgList, Tablette, FE_Main, EdtEtat, AglInitPlus, AGLInit,
  MajHalley, CreerSoc, CopySoc, Reseau, Mullog, LicUtil, UtomUtilisat, UserGrp_tom,
  UtilPGI, AssistInitSoc, PGIExec, PGIEnv, GalOutil, Reindex, Paramsoc, Ent1,
  UtilSoc, Exercice, Souche, FichComm, Pays, Region, CodePost, Devise, Tva, MulJal,
  MulGene, MulTiers, MulSecti, Axe, Structur, VentType, CodeAFB,
  EtbMce, OuvreExo, CalcEuro, EuroPGI, SuprGene, SuprJal, QRJDivis, EntPGI,
  UtofExportConfidentialite,Confidentialite_tof,MajSocParLot_TOF,LGCOMPTE_TOF;

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
  Result := 'SCO_CONSOLIDATION;SCO_SYNAPTIQUE;SCO_MCC;SCO_FISCAPERSO;SCO_TDI;SCO_XISOFLEX;SCO_BTBTP;SCO_HRCHR;SCO_KTTRANSPORT;SCO_GPAO';
end;

Procedure InitCodeProduit ;
var ii,i : integer;
    stl : TStringList;
    st,stCombos : string;
Begin
stCombos:= 'YYLISTEPRODUITS';
ii:=TTToNum(stCombos);
st:='';
if ii>0 then
begin
  if (V_PGI.DECombos[ii].Valeurs = Nil) then RemplirListe (stCombos,'');
  stl:=V_PGI.DeCombos[ii].valeurs;
  If stl=Nil then Exit;
  for i:=0 to stl.count-1 do
  begin
    if V_PGI.CodeProduit <> '' then st:=';';
    V_PGI.CodeProduit:=V_PGI.Codeproduit+st+ExtractInfo(stl[i],1);
  end;
end;
End;

procedure Dispatch(Num: Integer; PRien: THPanel; var retourforce, sortiehalley: boolean);
begin
  case Num of
    10:
      begin
        if (V_PGI.MajStructAuto) and (V_PGI.RunWithParams) then
        begin
          FMenuG.ForceClose := True;
          PostMessage(FMenuG.Handle, WM_CLOSE, 0, 0);
          Application.ProcessMessages;
        end
        else
          VerifSuperviseur;
        InitCodeProduit;
      end;

    11: ; //Après deconnection
    12:
      begin
        if ((V_PGI_ENV <> nil) and (V_PGI_ENV.ModePCL = '1')) then V_PGI.PGIContexte := V_PGI.PGIContexte + [ctxPCL];
      end;
    13: ;
    15: ;
    16: ;
    100: ; // executer depuis le lanceur
    {Sociétés et dossiers}
    3130: GestionSociete(PRien, @InitSocietePGI, nil);        
    3145: RecopieSoc(PRien, True);
    3220: ReindexSoc;
    4010: AGLLanceFiche('YY', 'YYIMPORTOBJET', '', '', ''); // Import d'objets
    4020: AGLLanceFiche('YY', 'YYMYINDEXES', '', '', ''); // IIndexes supplémentaires
    4030: AGLLanceFiche('YY', 'YYJNALEVENT', '', '', ''); // Journal evenements
    4003: YYLanceFiche_MajSocParLot('YY', 'MAJSOCPARLOT', '', '', '');
    {$IFDEF JOEL}
    4004: AGLLanceFiche('YY', 'YYIMPORTOBJETLOT', '', '', '');
    {$ENDIF}
    // majver
    4001: YYLanceFiche_ExportConfidentialite('YY', 'YYEXPORTCONFID', '', '', '');
    4002: YYLanceFiche_ImportConfidentialite('YY', 'YYIMPORTCONFID', '', '', '');

    {Utilisateurs et accès}
    3165: FicheUserGrp;
    3170:
      begin
        FicheUSer(V_PGI.User);
        ControleUsers;
      end;
    3172: AGLLanceFiche('YY', 'PROFILETABL', '', '', 'ETA');
    3180: ReseauUtilisateurs(False);
    3185: VisuLog;
    3195: ReseauUtilisateurs(True);
    4201 :  GCLanceFiche_Confidentialite( 'YY','YYCONFIDENTIALITE','','',''  );
    {Paramètres généraux}
    1104: ParamSociete(False, VireMoiCa, '', '', ChargeSocieteHalley, ChargePageSoc, SauvePageSoc, InterfaceSoc, 1105000);
    1290: FicheExercice('', taModif, 0);
    1300: FicheSouche('CPT');
    1110: FicheEtablissement_AGL(taModif);
    65410: ParamTable('YYDOMAINE', taCreat, 0, PRien);
    1120: ParamTable('ttFormeJuridique', taCreat, 1120000, PRien);
    1122: ParamTable('ttCivilite', taCreat, 1122000, PRien);
    1125: ParamTable('ttFonction', taCreat, 1125000, PRien);
    1135: AglLanceFiche('YY', 'YYPAYS', '', '', ''); // OuvrePays ;
    1140: FicheRegion('', '', False);
    1145: OuvrePostal(PRien);
    1150: FicheDevise('', tamodif, False);
    1165: ParamTable('TTREGIMETVA', taCreat, 1165000, PRien);
    1170: ParamTvaTpf(true);
    1175: ParamTvaTpf(false);
    1185: FicheModePaie_AGL('');
    1190: FicheRegle_AGL('', FALSE, taModif);
    1191: ParamTable('ttQualUnitMesure', taCreat, 1190030, PRien);
    1193: ParamTable('ttNivCredit', taCreat, 1190040, PRien);
    {Structures}
    7211: MulticritereJournal(taModif);
    7112: MulticritereCpteGene(taModif);
    7145: MultiCritereTiers(taModif);
    7118: SuppressionCpteGene; //Suppression des comptes généraux (unité SuprGene)
    7214: SuppressionJournaux; //Suppression des journaux (unité SuprJal)
    7394: JalDivisioPCL; //Édition des journaux divisionnaires (unité QRJDivis)
    3205 : CCLanceFiche_LongueurCompte ; // modif longueur des comptes
    {Analytique}
    1105: FicheAxeAna('');
    7178: MulticritereSection(taModif);
    1450: ParamPlanAnal('');
    1460: ParamVentilType;
    {Banque}
    1705: FicheBanque_AGL('', taModif, 0);
    1720: ParamTable('ttCFONB', taCreat, 0, PRien);
    1725: ParamCodeAFB;
    1730: ParamTable('ttCodeResident', taCreat, 0, PRien);
    1732: ParamTeletrans;
    {Traitements}
    7718: if PrepareOuvreExoSansCompta then OuvertureExo;
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
var VireCompta: Boolean;
begin
  V_PGI.VersionDemo := False;
  VireCompta := True;
  V_PGI.LaSerie := S5;
  if ctxPCL in V_PGI.PGIContexte then
  begin
    // maintenant qu'on peut le lancer en direct
    if V_PGI_Env.NoDossier = '' then V_PGI_Env.NoDossier := '000000';
  end else
  begin
    // PCS 13/10/2003 modules compta toujours présents.
    VireCompta := False;
  end;

  if VireCompta then
  begin
    FMenuG.RemoveGroup(-3200,True) ;
    FMenuG.RemoveGroup(-1705, True);
    FMenuG.RemoveGroup(-1105, True);
    FMenuG.RemoveGroup(-7211, True);
    FMenuG.RemoveGroup(-1104, True);
    FMenuG.RemoveItem(3172);
    FMenuG.RemoveItem(3185);
    FMenuG.RemoveItem(7718);
  end;
  UpdateSeries;
end;

procedure InitApplication;
begin
  FMenuG.OnMajAvant := MAJHalleyAvant;
  FMenuG.OnMajApres := MAJHalleyApres;
  FMenuG.IsMajPossible := MAJHalleyIsPossible;
  FMenuG.OnMajPendant := MajHalleyPendant;
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

  V_PGI.NumVersionBase := 625;

  Apalatys := 'CEGID';
  NomHalley := 'PGIMAJVER';
  TitreHalley := 'Administration PGI     Base ' + IntToStr(V_PGI.NumVersionBase);
  HalSocIni := 'CEGIDPGI.ini';
  Copyright := '© Copyright ' + Apalatys;
  V_PGI.NumVersion := '5.1.3';
  V_PGI.NumBuild := IntToStr(V_PGI.NumVersionBase)+'.0';
  V_PGI.DateVersion := EncodeDate(2004, 11, 10);
  {$IFDEF CCS3}
  V_PGI.LaSerie := S3;
  {$ELSE}
  V_PGI.LaSerie := S5;
  {$ENDIF}
  V_PGI.OutLook := TRUE;
  V_PGI.OfficeMsg := TRUE;
  V_PGI.ToolsBarRight := TRUE;
  ChargeXuelib;

  // pour forcer mode SAV en Qualité
  if FileExists(ExtractFilePath(Application.ExeName)+'INI_4B8D120B-1877-4E78-A817-331A23A35209.ini') then V_PGI.SAV:=True ;

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
end;

initialization
  ProcChargeV_PGI := InitialisationVPGI;

end.
