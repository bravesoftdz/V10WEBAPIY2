unit MDispIE ;

interface

uses
    Windows,
    Messages,
    SysUtils,
    Classes,
    Graphics,
    Controls,
    Forms,
    Dialogs,
    UIUtil,
    ExtCtrls,
    DBTables,
    StdCtrls,
    Mask,
    DBCtrls,
    Db,
    Grids,
    DBGrids,
    Menus,
    ComCtrls,
    HPanel,
    HStatus,
    HMenu97,
    Hctrls,
    HDB,
    HTB97,
    HEnt1,
    Ed_Tools,
    HRotLab,
    hmsgbox,
    HFLabel,
    Courrier,
    MulCour,
    LicUtil,
    HCalc,
    UserChg,
    Agenda,
    SQL,
    HMacro,
    AddDico,
    HCapCtrl,
    IniFiles,
    Prefs,
    Buttons,
    Xuelib,
    About,
    Traduc,
    TradForm,
    Tradmenu,
    Graph,
    FE_Main,
    UIUtil3,
    ImgList,
    Hout,
    TofExpEtafi,
    EdtREtat,
    galoutil,
    EntPGI
    ;

procedure AssignZoom ;
//SG6 01/01/2005 FQ 15162
procedure DispatchTT(Num : Integer; Action : TActionFiche; Lequel,TT,Range : String) ;
Procedure Alerte ;
Procedure Dispatch ( Num : Integer ; PRien : THPanel ; Var RetourForce,FermeHalley : boolean );
Procedure AfterProtec ( sAcces : String ) ;
procedure AfterChangeModule ( NumModule : integer ) ;
procedure InitApplication ;
Procedure InitLaVariablePGI;

type
  TFMDispIE = class(TDatamodule)
    procedure FMenuDispCreate(Sender: TObject);
end ;

Var FMDispIE : TFMDispIE ;

implementation

uses
    MenuOLG,
    Ent1,
    FmtChoix,
    CORRIMP,
    SISCOCOR,
    EcrLibre,
    AnaLibre,
    ExpCpte,
    ImpCpte,
    AssImp,
    ExpMvt,
    AnnulRS,
    CTRSISCO,
    ImpCPTSI,
    ImpMVTSI,
    AssistRS,
    CtdSISCO,
    TImpFic,
    CPSECTION_TOM,
    CPGENERAUX_TOM,
    CPJOURNAL_TOM,
    CPTIERS_TOM,
    IMPCPTA,
    AssistSS,
    ImpCPTSV,
    CpteUtil,
{$IFNDEF SPEC302}
    ImpCegid,
    ImpBds,
{$ENDIF}
    ImpMVTSV,
    Expecc,
    LETTREF,
    UTofCpMulMvt,
    //ConsEcr,
    uTofConsEcr,
    Resulexo,
    ParamSoc,
    UtilSoc,
    RadS3S5,
    UTofMulParamGen
    ;


Var PremFois : boolean ;

{$R *.DFM}


//SG6 01/01/2005 FQ 15162
procedure DispatchTT(Num : Integer; Action : TActionFiche; Lequel,TT,Range : String) ;
begin
  case Num of
    // general
    1,7112 : FicheGene(Nil,'',LeQuel,Action,0);
    // tiers
    2,7145 : FicheTiers(Nil,'',LeQuel,Action,1) ;
    // journal
    4,7211 : FicheJournal(Nil,'',Lequel,Action,0) ;


    // Section
    71781 : FicheSection(nil,'A1',Lequel,Action,1);
    71782 : FicheSection(nil,'A2',Lequel,Action,1);
    71783 : FicheSection(nil,'A3',Lequel,Action,1);
    71784 : FicheSection(nil,'A4',Lequel,Action,1);
    71785 : FicheSection(nil,'A5',Lequel,Action,1);

    end ;
end ;


procedure AssignZoom ;
BEGIN
(*
ProcZoomEdt:=ZoomEdtEtat ;
ProcCalcEdt:=CalcOLEEtat ;
ProcGetVH:=GetVH ;
ProcGetDate:=GetDate ;
*)
If Not Assigned(ProcZoomGene)    Then ProcZoomGene    :=FicheGene    ;
If Not Assigned(ProcZoomTiers)   Then ProcZoomTiers   :=FicheTiers   ;
If Not Assigned(ProcZoomSection) Then ProcZoomSection :=FicheSection ;
If Not Assigned(ProcZoomJal)     Then ProcZoomJal     :=FicheJournal ;
(*
If Not Assigned(ProcZoomCorresp) Then ProcZoomCorresp :=ZoomCorresp  ;
*)
// GCGC If Not Assigned(ProcZoomArticle) Then ProcZoomArticle :=FicheArticle ;
(*
If Not Assigned(ProcZoomNatCpt)  Then ProcZoomNatCpt  :=FicheNatCpte ;
If Not Assigned(ProcZoomBudGen)  Then ProcZoomBudGen  :=FicheBudgene ;
If Not Assigned(ProcZoomBudSec)  Then ProcZoomBudSec  :=FicheBudsect ;
If Not Assigned(ProcZoomRub)     Then ProcZoomRub     :=FicheRubrique ;
*)
END ;

Procedure Alerte ;
Var St : String ;
BEGIN
If (V_PGI.LaSerie=S5) Then St:='S5' Else St:='S3' ;
HShowMessage('0;?caption?;'+TraduireMemoire('Fonction non disponible en '+St)+';W;O;O;O;',TitreHalley,'')
END ;

Procedure Dispatch ( Num : Integer ; PRien : THPanel ; Var RetourForce,FermeHalley : boolean );
var St , StNom, StValeur: string;
    i : integer;
BEGIN
VH^.RecupLTL:=FALSE ;
{$IFDEF RECUPLTL}
VH^.RecupLTL:=TRUE ;
{$ENDIF}
RecupSISCO:=FALSE ; RecupServant:=FALSE ;
if PremFois then BEGIN UpDateSeries ; PremFois:=False ; END ;
FMenuG.bSeria.Visible:=FALSE ;
AssignZoom ;
Case Num of
    10 : BEGIN
           ChargeHalleyUser ;
          if V_PGI.ModePCL='1' then V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxPCL] ;
         END ;
     11 : ; // Après déconnexion
     12 :   begin
              // Interdiction de lancer en direct
              if (Not V_PGI.RunFromLanceur) and (V_PGI.ModePCL='1') then
              begin
                FMenuG.FermeSoc;
                FMenuG.Quitter;
                exit;
              end;
              if ctxPCL in V_PGI.PGIContexte then
              begin
                VH^.ModeSilencieux:=GetFlagAppli('CCS5.EXE');
              end;
            end;
     13 : ;
     15 : ;
     16 : ;
     100 : begin
           if ctxPCL in V_PGI.PGIContexte then
           BEGIN
             for i:=1 to ParamCount do
             begin
               St:=ParamStr(i) ; StNom:=UpperCase(Trim(ReadTokenPipe(St,'='))) ;
               StValeur:=UpperCase(Trim(St)) ;
               if StNom='/TRF' then
               begin
                 SavImo2PGI(StValeur,True);
                 // Dans le cas du transfert, on ferme l'appli.
                 SetFlagAppli('CCS5.EXE',True);
                 FMenuG.ForceClose := True; FMenuG.Close; Exit;
               end;
             end;
             RetourForce:=True ;
           END ;
           end;
   6110 : ChoixFormatImpExp('X','') ;
   6120 : ChoixFormatImpExp('-','') ;
   6130 : ParamCorrespImp('IGE') ;
   6132 : ParamCorrespImp('IAU') ;
   6134 : ParamCorrespImp('IA1') ;
   6136 : if (V_PGI.LaSerie=S7) Or (V_PGI.LaSerie=S5) then ParamCorrespImp('IA2') Else Alerte ;
   6138 : if V_PGI.LaSerie<>S3 then ParamCorrespImp('IA3') else Alerte ;
   6140 : if V_PGI.LaSerie<>S3 then ParamCorrespImp('IA4') else Alerte ;
   6142 : if V_PGI.LaSerie<>S3 then ParamCorrespImp('IA5') else Alerte ;
   6144 : ParamCorrespImp('IPM') ;
   6146 : ParamAuxSISCO ;
   6148 : ParamCorrespImp('IJA') ;

   6210 : ImportCpte('G') ;
   6220 : ImportCpte('X') ;
   6230 : ImportCpte('A') ;
   6232 : if (V_PGI.LaSerie=S7) Or (V_PGI.LaSerie=S5) Then ImportCpte('FBJ') Else Alerte ;
   6235 : if (V_PGI.LaSerie=S7) Or (V_PGI.LaSerie=S5) Then ImportCpte('FBG') Else Alerte ;
   6236 : if (V_PGI.LaSerie=S7) Or (V_PGI.LaSerie=S5) Then ImportCpte('FBS') Else Alerte ;

   6241 : LanceImport('FEC') ;
   6242 : if (V_PGI.LaSerie=S7) Or (V_PGI.LaSerie=S5) Then LanceImport('FOD') Else Alerte ;
   6243 : if (V_PGI.LaSerie=S7) Or (V_PGI.LaSerie=S5) Then LanceImport('FBE') Else Alerte ;
   6250 : if (V_PGI.LaSerie=S7) Or (V_PGI.LaSerie=S5) Then LanceImport('FBA') Else Alerte ;
{$IFNDEF SPEC302}
   6252 : if (V_PGI.LaSerie=S7) Or (V_PGI.LaSerie=S5) Then ImportBds('BDS') Else Alerte ;
{$ENDIF}

   6310 : ExportCpte('G') ;
   6320 : ExportCpte('X') ;
   6330 : ExportCpte('A') ;
   6332 : if (V_PGI.LaSerie=S7) Or (V_PGI.LaSerie=S5) Then ExportCpte('FBJ') Else Alerte ;
   6335 : if (V_PGI.LaSerie=S7) Or (V_PGI.LaSerie=S5) Then ExportCpte('FBG') Else Alerte ;
   6336 : if (V_PGI.LaSerie=S7) Or (V_PGI.LaSerie=S5) Then ExportCpte('FBS') Else Alerte ;

   6341 : ExportMvt('FEC') ;
   6342 : ExportMvt('FOD') ;
   6343 : if (V_PGI.LaSerie=S7) Or (V_PGI.LaSerie=S5) Then ExportMvt('FBE') Else Alerte ;
// N0 10173 le 14/06/2002   6350 : if (V_PGI.LaSerie=S7) Or (V_PGI.LaSerie=S5) Then ExportMvt('FBA') Else Alerte ;
   6350 : ExportMvt('FBA') ;

   6361 : MultiCritereEcrLibre('1') ;
   6362 : MultiCritereEcrLibre('2') ;
   6365 : MultiCritereAnaLibre('1') ;
   6366 : MultiCritereAnaLibre('2') ;

   6368 : ExportEnCours ;
   6369 : CP_LanceFicheExpEtafi;

   6410 : BEGIN RecupSISCO:=TRUE ; LanceAssistRSISCO ; END ;
   6420 : BEGIN RecupSISCO:=TRUE ; If ControleAvantRecupSISCO Then ImportCompteSISCO ; END ;
   6430 : BEGIN RecupSISCO:=TRUE ; CtrlDivSISCO ; RecupSISCO:=FALSE ; END ;
   6440 : BEGIN RecupSISCO:=TRUE ; If ControleAvantRecupSISCO Then ImportMvtSISCO ; END ;
   6441 : BEGIN LettrageParCode ; END ;
   6450 : BEGIN RecupSISCO:=TRUE ; AnnuleRecupSISCO ; END ;
   6460 : BEGIN RecupSERVANT:=TRUE ; LanceAssistRSERVANT ; END ;
   6461 : BEGIN RecupSERVANT:=TRUE ; ImportCompteSERVANT(0) ; END ;
   6462 : BEGIN RecupSERVANT:=TRUE ; ImportCompteSERVANT(1) ; END ;
   6463 : BEGIN RecupSERVANT:=TRUE ; CtrlDivSISCO ; END ;
   6464 : BEGIN RecupSERVANT:=TRUE ; If ControleAvantRecupSISCO Then ImportMvtSERVANT ; END ;
   6465 : BEGIN RecupSERVANT:=TRUE ; AnnuleRecupSISCO ;END ;

   6510 : RevisionADistance(TRUE) ;
// Jusqu'à 6499...
   3255 : RecupSCM ;
// Immobilisations
{$IFNDEF SPEC302}
   2390 : // ImportCegid('I')
   SavImo2PGI('',False);
{$ENDIF}
    //Menu Pop Compta
   25710 : CCLanceFiche_MULGeneraux('P;7112000');  // Comptes généraux
   25720 : CPLanceFiche_MULTiers('P;7145000');     // Comptes auxiliaires
   25730 : CPLanceFiche_MULSection('P;7178000');   // Sections analytiques
   25740 : CPLanceFiche_MULJournal('P;7205000');   // Journaux
   25750 : MultiCritereMvt(taConsult,'N',False) ;
   25755 : OperationsSurComptes('','','') ;
   25760 : ResultatDeLexo ;
   25770 : ParamSociete(False,'SCO_CPREVISION','SCO_PARAMETRESGENERAUX','',RechargeParamSoc,ChargePageSoc,SauvePageSoc,InterfaceSoc,1105000) ;

     else HShowMessage('2;?caption?;'+TraduireMemoire('Fonction non disponible : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)) ;
   end ;
END ;

Procedure AfterProtec ( sAcces : String ) ;
BEGIN
VH^.OkModCompta:=True ; VH^.OkModBudget:=TRUE ; VH^.OkModImmo:=TRUE ;
(*V_PGI.Monoposte:=TRUE ;*) V_PGI.VersionDemo:=FALSE ;
V_PGI.NumMenuPop := 27;    

END ;

procedure AfterChangeModule ( NumModule : integer ) ;
BEGIN
ChargeMenuPop(Integer(hm2),FMenuG.DispatchX) ;
FMenuG.bSeria.Visible:=FALSE ;

  Case NumModule of
    // Révision à distance uniquement en PCL
    6 : begin
      ExecuteSQL('UPDATE MENU SET MN_TAG=6500 WHERE MN_1=6 AND MN_2=7 AND MN_3=0 AND MN_4=0');
      if not(ctxPCL in V_PGI.PGIContexte) then FMenuG.RemoveGroup(6500,True);
      FMenuG.RemoveGroup(6110, True);
      FMenuG.RemoveGroup(6200, True);
      FMenuG.RemoveGroup(6400, True);
      FMenuG.RemoveGroup(3255, True);
      FMenuG.RemoveGroup(6500, True);
      ChargeMenuPop(integer(hm2),FMenuG.DispatchX) ;
    end;
  end;
END ;

procedure InitApplication ;
BEGIN
//ProcZoomEdt:=ZoomEdtEtat ;
//ProcCalcEdt:=CalcOLEEtat ;
ProcGetVH:=GetVH ;
ProcGetDate:=GetDate ;
FMenuG.OnAfterProtec:=AfterProtec ;
FMenuG.OnDispatch:=Dispatch ;
FMenuG.OnChargeMag:=ChargeMagHalley ;
FMenuG.OnMajAvant:=Nil ;
FMenuG.OnMajApres:=Nil ;
FMenuG.SetModules([6],[0]) ;
FMenuG.OnChangeModule:=AfterChangeModule ;
FMenuG.SetPreferences(['Saisies'],False) ;
END ;

procedure TFMDispIE.FMenuDispCreate(Sender: TObject);
begin
InitApplication ;
PremFois:=True ;
end;

Procedure InitLaVariablePGI;
Begin
Apalatys:='CEGID' ;
Copyright:='© Copyright ' + Apalatys ;
V_PGI.OutLook:=TRUE ;
V_PGI.OfficeMsg:=TRUE ;
V_PGI.VersionDemo:=TRUE ;
V_PGI.MenuCourant:=0 ;
V_PGI.NumVersion:='8.1.0' ; V_PGI.NumBuild:='002.001' ; V_PGI.NumVersionBase:=850 ;

// 13584
if ctxPCL in V_PGI.PGIContexte then V_PGI.PortailWeb := 'http://experts.cegid.fr/home.asp'
                               else V_PGI.PortailWeb := 'http://utilisateurs.cegid.fr';

V_PGI.DateVersion:=EncodeDate(2007,07,31);
V_PGI.ImpMatrix := True ;
V_PGI.OKOuvert:=FALSE ;
V_PGI.Halley:=TRUE ;
V_PGI.ParamSocLast:=True ;
V_PGI.NoModuleButtons:=TRUE ;
V_PGI.RAZForme:=TRUE ;
V_PGI.CegidApalatys := False;
V_PGI.CegidBureau := True;
V_PGI.StandardSurDP := True;
V_PGI.MajPredefini:=False ;
V_PGI.BlockMAJStruct:=True ;
V_PGI.EuroCertifiee:=False ;

//SG6 01/01/2005 FQ 15162
V_PGI.DispatchTT:=DispatchTT ;
//Fin Test

If ParamCount>0 Then If ParamStr(1)='ORLIBLANCS' Then PasDeBlanc:=TRUE ;
RenseignelaSerie(ExeCCIMPEXP) ;
ChargeXuelib ;
end;

Initialization
  ProcChargeV_PGI :=  InitLaVariablePGI ;

end.

