unit AssistInitSoc;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, db, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
  HPanel, HEnt1, MajTable,
  {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF} Hqry, Hgauge,
   hCRC, HStatus, LicUtil,HDimension
  {$IFDEF AFFAIRE}
  ,AfInitSoc, ADODB // initialisation metier Affaire
  {$ENDIF AFFAIRE}
  ;

type
  TFAssistInitSoc = class(TFAssist)
    TABGEN: TTabSheet;
    BMetier: TBevel;
    PTITRE: THPanel;
    TABINIT: TTabSheet;
    BInfos: TBevel;
    TFRef: THLabel;
    FRef: THValComboBox;
    GSOCIETE: TGroupBox;
    TSO_SOCIETE: THLabel;
    TSO_LIBELLE: THLabel;
    SO_SOCIETE: TEdit;
    SO_LIBELLE: TEdit;
    GUSERGRP: TGroupBox;
    TUG_GROUPE: THLabel;
    TUG_LIBELLE: THLabel;
    UG_GROUPE: TEdit;
    UG_NIVEAUACCES: THNumEdit;
    UG_NUMERO: THNumEdit;
    UG_LIBELLE: TEdit;
    GUTILISAT: TGroupBox;
    TUS_UTILISATEUR: THLabel;
    TUS_PASSWORD: THLabel;
    TUS_ABREGE: THLabel;
    TUS_LIBELLE: THLabel;
    TUS_PASSWORD2: THLabel;
    US_UTILISATEUR: TEdit;
    US_PASSWORD: TEdit;
    US_ABREGE: TEdit;
    US_GROUPE: TEdit;
    US_SUPERVISEUR: TEdit;
    US_LIBELLE: TEdit;
    FPassWord2: TEdit;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    BGes: TBevel;
    BCpta: TBevel;
    TableDefaut: THValComboBox;
    TabSheet0: TTabSheet;
    BGen: TBevel;
    GBAdmin: TGroupBox;
    USERGRP: TCheckBox;
    UTILISAT: TCheckBox;
    TINTRO: THLabel;
    TABRECAP: TTabSheet;
    BRecap: TBevel;
    Entete: THTable;
    LLigne: THTable;
    GB_DOM: TGroupBox;
    CB_DOM1: TCheckBox;
    CB_DOM2: TCheckBox;
    CB_DOM3: TCheckBox;
    CB_DOM4: TCheckBox;
    DBREF: TDatabase;
    DBSOURCE: TDatabase;
    DBDest: TDatabase;
    EG_Progress: TEnhancedGauge;
    PanelFin: TPanel;
    TTextFin1: THLabel;
    TTextFin2: THLabel;
    TEtat: THLabel;
    HPanel1: THPanel;
    GBTab1: TGroupBox;
    PARPIECE: TCheckBox;
    SOUCHE: TCheckBox;
    GBTT1: TGroupBox;
    CB_DOM5: TCheckBox;
    TTDefaut: THValComboBox;
    TabSheet3: TTabSheet;
    TabSheet5: TTabSheet;
    TabSheet4: TTabSheet;
    BAff: TBevel;
    BMode: TBevel;
    BRCli: TBevel;
    GBTab2: TGroupBox;
    GBTT2: TGroupBox;
    EXERCICE: TCheckBox;
    SUIVCPTA: TCheckBox;
    CORRESP: TCheckBox;
    REFAUTO: TCheckBox;
    GUIDES: TCheckBox;
    RUBRIQUE: TCheckBox;
    RUPTURE: TCheckBox;
    GBTab3: TGroupBox;
    GBTT3: TGroupBox;
    GBTab4: TGroupBox;
    GBTT4: TGroupBox;
    GBTab5: TGroupBox;
    GBTT5: TGroupBox;
    TT_GCCATEGORIETAXE: TCheckBox;
    TT_DIMENSIONS: TCheckBox;
    TT_GCLIBFAMILLE: TCheckBox;
    TT_GCZONELIBRE: TCheckBox;
    TT_AFTRADUCAFFAIRE: TCheckBox;
    TT_AFTTYPEHEURE: TCheckBox;
    TT_AFTTYPERESSOURCE: TCheckBox;
    TT_RTLIBTABLECOMPL: TCheckBox;
    DIMMASQUE: TCheckBox;
    DIMENSION: TCheckBox;
    PTitre1: THPanel;
    PTitre2: THPanel;
    PTitre3: THPanel;
    PTitre4: THPanel;
    PTitre5: THPanel;
    PAYS: TCheckBox;
    REGION: TCheckBox;
    CODEPOST: TCheckBox;
    CODEAFB: TCheckBox;
    DEVISE: TCheckBox;
    MODEPAIE: TCheckBox;
    MODEREGL: TCheckBox;
    MODELES: TCheckBox;
    ETABLISS: TCheckBox;
    MODEDATA: TCheckBox;
    TT_GCCATEGORIEDIM: TCheckBox;
    TProgress: THLabel;
    CHAMPSPRO: TCheckBox;
    PARACTIONS: TCheckBox;
    PARSUSPECT: TCheckBox;
    PARSUSPECTLIG: TCheckBox;
    GroupBox2: TGroupBox;
    TT_RTTYPEPERSPECTIVE: TCheckBox;
    RTRPRLIBPERSPECTIVE: TCheckBox;
    TT_RTRPCLIBTABLE0: TCheckBox;
    TT_RTRPCLIBTABLE1: TCheckBox;
    TT_RTRPCLIBTABLE2: TCheckBox;
    TT_RTRPCLIBTABLE3: TCheckBox;
    TT_RTRPCLIBTABLE4: TCheckBox;
    TT_RTRPCLIBTABLE5: TCheckBox;
    TT_RTRPCLIBTABLE6: TCheckBox;
    TT_RTRPCLIBTABLE7: TCheckBox;
    TT_RTRPCLIBTABLE8: TCheckBox;
    TT_RTRPCLIBTABLE9: TCheckBox;
    TT_RTRPCLIBTABLEA: TCheckBox;
    TT_RTRPCLIBTABLEB: TCheckBox;
    TT_RTRPRLIBPERSPECTIVE1: TCheckBox;
    TT_RTRPRLIBPERSPECTIVE2: TCheckBox;
    TT_RTRPRLIBPERSPECTIVE3: TCheckBox;
    GroupBox3: TGroupBox;
    TT_RTIMPORTANCEACT: TCheckBox;
    RTRPRLIBACTION: TCheckBox;
    TT_RTRPRLIBACTION1: TCheckBox;
    TT_RTRPRLIBACTION2: TCheckBox;
    TT_RTRPRLIBACTION3: TCheckBox;
    TT_RTLIBCHAMPSLIBRES: TCheckBox;
    TT_RTRPRLIBTABLE0: TCheckBox;
    TT_RTRPRLIBTABLE1: TCheckBox;
    TT_RTRPRLIBTABLE2: TCheckBox;
    TT_RTRPRLIBTABLE3: TCheckBox;
    TT_RTRPRLIBTABLE4: TCheckBox;
    TT_RTRPRLIBTABLE5: TCheckBox;
    TT_RTRPRLIBTABLE6: TCheckBox;
    TT_RTRPRLIBTABLE7: TCheckBox;
    TT_RTRPRLIBTABLE8: TCheckBox;
    TT_RTRPRLIBTABLE9: TCheckBox;
    TT_RTRPRLIBTABLE10: TCheckBox;
    TT_RTRPRLIBTABLE11: TCheckBox;
    TT_RTRPRLIBTABLE12: TCheckBox;
    TT_RTRPRLIBTABLE13: TCheckBox;
    TT_RTRPRLIBTABLE14: TCheckBox;
    TT_RTRPRLIBTABLE15: TCheckBox;
    TT_RTRPRLIBTABLE16: TCheckBox;
    TT_RTRPRLIBTABLE17: TCheckBox;
    TT_RTRPRLIBTABLE18: TCheckBox;
    TT_RTRPRLIBTABLE20: TCheckBox;
    TT_RTRPRLIBTABLE19: TCheckBox;
    TT_RTRPRLIBTABLE21: TCheckBox;
    TT_RTRPRLIBTABLE22: TCheckBox;
    TT_RTRPRLIBTABLE23: TCheckBox;
    TT_RTRPRLIBTABLE24: TCheckBox;
    TT_RTRPRLIBTABLE25: TCheckBox;
    TT_RTLIBCLACTPERSP: TCheckBox;
    TT_RTRPRLIBTABMUL1: TCheckBox;
    TT_RTRPRLIBTABMUL2: TCheckBox;
    TT_RTRPRLIBTABMUL3: TCheckBox;
    TT_RTRPRLIBTABMUL4: TCheckBox;
    TT_RTRPRLIBTABMUL5: TCheckBox;
    TT_RTRPRLIBTABMUL6: TCheckBox;
    TT_RTRPRLIBTABMUL7: TCheckBox;
    TT_RTRPRLIBTABMUL8: TCheckBox;
    TT_RTRPRLIBTABMUL9: TCheckBox;
    TT_RTRPRLIBTABMUL0: TCheckBox;
    GroupBox4: TGroupBox;
    TT_RTMOTIFFERMETURE: TCheckBox;
    TT_RTLIBCHPLIBSUSPECTS: TCheckBox;
    TT_RTRSCLIBTABLE0: TCheckBox;
    TT_RTRSCLIBTABLE1: TCheckBox;
    TT_RTRSCLIBTABLE2: TCheckBox;
    TT_RTRSCLIBTABLE3: TCheckBox;
    TT_RTRSCLIBTABLE4: TCheckBox;
    TT_RTRSCLIBTABLE5: TCheckBox;
    TT_RTRSCLIBTABLE6: TCheckBox;
    TT_RTRSCLIBTABLE7: TCheckBox;
    TT_RTRSCLIBTABLE8: TCheckBox;
    TT_RTRSCLIBTABLE9: TCheckBox;
    GroupBox5: TGroupBox;
    TT_RTOBJETOPE: TCheckBox;
    TT_RTMOTIFS: TCheckBox;
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure US_LIBELLEChange(Sender: TObject);
    procedure UG_GROUPEChange(Sender: TObject);
    procedure MODELESClick(Sender: TObject);
    procedure UTILISATClick(Sender: TObject);
    procedure USERGRPClick(Sender: TObject);
    procedure MODEPAIEClick(Sender: TObject);
    procedure MODEREGLClick(Sender: TObject);
    procedure DIMENSIONClick(Sender: TObject);
    procedure TT_DIMENSIONSClick(Sender: TObject);
    // grc
    procedure TT_RTLIBTABLECOMPLClick(Sender: TObject);
    procedure PARSUSPECTClick(Sender: TObject);
    procedure RTRPRLIBPERSPECTIVEClick(Sender: TObject);
    procedure RTRPRLIBACTIONClick(Sender: TObject);
    procedure CHAMPSPROClick(Sender: TObject);
    procedure PARACTIONSClick(Sender: TObject);
    procedure TT_RTLIBCHPLIBSUSPECTSClick(Sender: TObject);



  private
    { Déclarations privées }

    FDestination : String ;
    SocDriver, RDriver, SDriver, DDriver : TDBDriver ;
    SocODBC, RODBC, SODBC, DODBC : Boolean ;

    function  okPage(i_indice : integer) : boolean;
    function  OkDonnees : boolean ;
    procedure ForceChecked(C1,C2 : TCheckBox) ;
    procedure PagePrecedente;
    procedure PageSuivante;
    function  TableSelect(stTable : string) : boolean;

    procedure InitialiseSoc ;
    procedure CreerTables ;
    procedure InitialiseLesTables ;
    procedure RecopieTablettes(Tablette, Base : string);
    function  RechercheChoixType(Tablette,Base : string) : string;
    function  RecupChpDB (DB,Fichier,Champ,Where : String ; var Valeur : Variant ) : Boolean ;
    function  Nbrecords(Table : TDataSet) : Integer ;
    procedure MajRupture ;

    //Spécificités par métier
    procedure SpecifMetier1;
    procedure SpecifMetier2;
    procedure SpecifMetier3;
    procedure SpecifMetier4;
    procedure SpecifMetier5;

  public
    { Déclarations publiques }
  end;

var
  FAssistInitSoc: TFAssistInitSoc;

procedure InitSocietePGI(FDestination : String) ;

implementation

{$R *.DFM}

procedure InitSocietePGI(FDestination : String) ;
var FoAssistInitSoc : TFAssistInitSoc ;
begin
FoAssistInitSoc := TFAssistInitSoc.Create (Application);
Try
   FoAssistInitSoc.FDestination:=FDestination ;
   FoAssistInitSoc.ShowModal;
   Finally
   FoAssistInitSoc.free;
   end;
end;

////////////////////////////////////////////////////////////////////////////////
//*************************Evénements de la form******************************//
////////////////////////////////////////////////////////////////////////////////

procedure TFAssistInitSoc.FormShow(Sender: TObject);
begin
  inherited;
SocDriver := V_PGI.Driver ;
SocODBC   := V_PGI.ODBC ;
V_PGI.StopCourrier := true ;
ChargeDossier(FRef.Items,True) ;                                       
FRef.Items.Insert(0,'<<Aucune>>');
FRef.ItemIndex  := 0 ;
SO_SOCIETE.Text := V_PGI.CodeSociete ;
SO_LIBELLE.Text := V_PGI.NomSociete ;
UG_GROUPE.Text  := V_PGI.Groupe ;
UG_LIBELLE.Text := RechDom('ttUserGroupe',V_PGI.Groupe,False) ;
US_LIBELLE.Text := RechDom('ttUtilisateur',V_PGI.User,False) ;
US_UTILISATEUR.Text := V_PGI.User ;
end;

procedure TFAssistInitSoc.bSuivantClick(Sender: TObject);
begin
  inherited;
if previouspage.Name = 'TABINIT' then
   begin
   if (FRef.Items[FRef.ItemIndex]=FDestination) then
      begin
      RestorePage;
      P.ActivePage := TABINIT ;
      PChange(nil) ;
      FRef.SetFocus ;
      Msg.Execute(8,Caption,'') ;
      Exit ;
      end ;

   if not OkDonnees then
      begin
      Msg.Execute(3,Caption,'') ;
      exit ;
      end ;

   if FPassWord2.Text <> US_PASSWORD.Text then
      begin
      RestorePage;
      P.ActivePage := TABINIT ;
      PChange(nil) ;
      FPassword2.SetFocus ;
      Msg.Execute(10,Caption,'') ;
      exit ;
      end ;
   end;
PageSuivante;
if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
end;

procedure TFAssistInitSoc.bPrecedentClick(Sender: TObject);
begin
  inherited;
PagePrecedente;
if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
end;


procedure TFAssistInitSoc.bFinClick(Sender: TObject);
begin
  inherited;
case Msg.Execute(2,Caption,'') of
  mrCancel : Exit ;
  end ;

InitialiseSoc ;
Msg.Execute(4,Caption,Chr(10)+Msg.Mess[9]) ;
ModalResult:=mrOk ;
end;

procedure TFAssistInitSoc.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
V_PGI.Driver := SocDriver ;
V_PGI.ODBC   := SocODBC ;
V_PGI.StopCourrier := false ;
end;

procedure TFAssistInitSoc.US_LIBELLEChange(Sender: TObject);
begin
  inherited;
US_ABREGE.Text := Copy(US_LIBELLE.Text,1,17) ;
end;

procedure TFAssistInitSoc.UG_GROUPEChange(Sender: TObject);
begin
  inherited;
US_GROUPE.Text := UG_GROUPE.Text ;
end;


procedure TFAssistInitSoc.MODELESClick(Sender: TObject);
begin
  inherited;
MODEDATA.Checked:=MODELES.Checked;
end;


procedure TFAssistInitSoc.UTILISATClick(Sender: TObject);
begin
  inherited;
ForceChecked(UTILISAT,USERGRP) ;
end;

procedure TFAssistInitSoc.USERGRPClick(Sender: TObject);
begin
  inherited;
ForceChecked(UTILISAT,USERGRP) ;
end;

procedure TFAssistInitSoc.MODEPAIEClick(Sender: TObject);
begin
  inherited;
ForceChecked(MODEREGL,MODEPAIE) ;
end;

procedure TFAssistInitSoc.MODEREGLClick(Sender: TObject);
begin
  inherited;
ForceChecked(MODEREGL,MODEPAIE) ;
end;

procedure TFAssistInitSoc.DIMENSIONClick(Sender: TObject);
begin
  inherited;
ForceChecked(DIMENSION,TT_DIMENSIONS) ;
DIMMASQUE.State := DIMENSION.State;
TT_GCCATEGORIEDIM.State := TT_DIMENSIONS.State;
end;

procedure TFAssistInitSoc.TT_DIMENSIONSClick(Sender: TObject);
begin
  inherited;
ForceChecked(DIMENSION,TT_DIMENSIONS) ;
TT_GCCATEGORIEDIM.State := TT_DIMENSIONS.State;
end;

procedure TFAssistInitSoc.TT_RTLIBTABLECOMPLClick(Sender: TObject);
begin
inherited;
TT_RTRPCLIBTABLE0.State := TT_RTLIBTABLECOMPL.State ;
TT_RTRPCLIBTABLE1.State := TT_RTRPCLIBTABLE0.State;
TT_RTRPCLIBTABLE2.State := TT_RTRPCLIBTABLE0.State;
TT_RTRPCLIBTABLE3.State := TT_RTRPCLIBTABLE0.State;
TT_RTRPCLIBTABLE4.State := TT_RTRPCLIBTABLE0.State;
TT_RTRPCLIBTABLE5.State := TT_RTRPCLIBTABLE0.State;
TT_RTRPCLIBTABLE6.State := TT_RTRPCLIBTABLE0.State;
TT_RTRPCLIBTABLE7.State := TT_RTRPCLIBTABLE0.State;
TT_RTRPCLIBTABLE8.State := TT_RTRPCLIBTABLE0.State;
TT_RTRPCLIBTABLE9.State := TT_RTRPCLIBTABLE0.State;
TT_RTRPCLIBTABLEA.State := TT_RTRPCLIBTABLE0.State;
TT_RTRPCLIBTABLEB.State := TT_RTRPCLIBTABLE0.State;
end;

procedure TFAssistInitSoc.RTRPRLIBACTIONClick(Sender: TObject);
begin
inherited;
TT_RTRPRLIBACTION1.State := RTRPRLIBACTION.State ;
TT_RTRPRLIBACTION2.State := TT_RTRPRLIBACTION1.State;
TT_RTRPRLIBACTION3.State := TT_RTRPRLIBACTION1.State;
if ((RTRPRLIBACTION.Checked = True) or (RTRPRLIBPERSPECTIVE.Checked = True))
   and ( TT_RTLIBCHAMPSLIBRES.Checked = False ) then // car TT_RTLIBCLACTPERSP comprise dans TT_RTLIBCHAMPSLIBRES
    TT_RTLIBCLACTPERSP.Checked := True
else
    TT_RTLIBCLACTPERSP.Checked := False;
end;

procedure TFAssistInitSoc.RTRPRLIBPERSPECTIVEClick(Sender: TObject);
begin
inherited;
TT_RTRPRLIBPERSPECTIVE1.State := RTRPRLIBPERSPECTIVE.State ;
TT_RTRPRLIBPERSPECTIVE2.State := TT_RTRPRLIBPERSPECTIVE1.State;
TT_RTRPRLIBPERSPECTIVE3.State := TT_RTRPRLIBPERSPECTIVE1.State;
if ((RTRPRLIBACTION.Checked = True) or (RTRPRLIBPERSPECTIVE.Checked = True))
   and ( TT_RTLIBCHAMPSLIBRES.Checked = False ) then // car TT_RTLIBCLACTPERSP comprise dans TT_RTLIBCHAMPSLIBRES
    TT_RTLIBCLACTPERSP.Checked := True
else
    TT_RTLIBCLACTPERSP.Checked := False;
end;

procedure TFAssistInitSoc.CHAMPSPROClick(Sender: TObject);
begin
inherited;
TT_RTLIBCHAMPSLIBRES.State := CHAMPSPRO.State ;
if TT_RTLIBCHAMPSLIBRES.Checked = True  then // car TT_RTLIBCLACTPERSP comprise dans TT_RTLIBCHAMPSLIBRES
   TT_RTLIBCLACTPERSP.Checked := False
else
   if (RTRPRLIBACTION.Checked = True) or (RTRPRLIBPERSPECTIVE.Checked = True) then
      TT_RTLIBCLACTPERSP.Checked := True
   else
      TT_RTLIBCLACTPERSP.Checked := False;
   
TT_RTRPRLIBTABLE0.State := CHAMPSPRO.State ;
TT_RTRPRLIBTABLE1.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE2.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE3.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE4.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE5.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE6.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE7.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE8.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE9.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE10.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE11.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE12.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE13.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE14.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE15.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE16.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE17.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE18.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE19.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE20.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE21.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE22.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE23.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE24.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABLE25.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABMUL0.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABMUL1.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABMUL2.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABMUL3.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABMUL4.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABMUL5.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABMUL6.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABMUL7.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABMUL8.State := TT_RTRPRLIBTABLE0.State;
TT_RTRPRLIBTABMUL9.State := TT_RTRPRLIBTABLE0.State;
end;

procedure TFAssistInitSoc.PARACTIONSClick(Sender: TObject);
begin
inherited;
if PARACTIONS.Checked = True then
    begin
    TT_RTIMPORTANCEACT.Checked := True;
    RTRPRLIBACTION.Checked := True;
    TT_RTRPRLIBACTION1.Checked := True;
    TT_RTRPRLIBACTION2.Checked := True;
    TT_RTRPRLIBACTION3.Checked := True;
    if TT_RTLIBCHAMPSLIBRES.Checked = False then // car TT_RTLIBCLACTPERSP comprise dans TT_RTLIBCHAMPSLIBRES
        TT_RTLIBCLACTPERSP.Checked := True
    else
        TT_RTLIBCLACTPERSP.Checked := False;
    end;
end;

procedure TFAssistInitSoc.PARSUSPECTClick(Sender: TObject);
begin
inherited;
PARSUSPECTLIG.State := PARSUSPECT.State ;
if PARSUSPECT.Checked = True then
    begin
    TT_RTMOTIFFERMETURE.Checked := True;
    TT_RTLIBCHPLIBSUSPECTS.Checked := True;
    TT_RTRSCLIBTABLE0.Checked := True;
    TT_RTRSCLIBTABLE1.Checked := True;
    TT_RTRSCLIBTABLE2.Checked := True;
    TT_RTRSCLIBTABLE3.Checked := True;
    TT_RTRSCLIBTABLE4.Checked := True;
    TT_RTRSCLIBTABLE5.Checked := True;
    TT_RTRSCLIBTABLE6.Checked := True;
    TT_RTRSCLIBTABLE7.Checked := True;
    TT_RTRSCLIBTABLE8.Checked := True;
    TT_RTRSCLIBTABLE9.Checked := True;
    end;
end;


procedure TFAssistInitSoc.TT_RTLIBCHPLIBSUSPECTSClick(Sender: TObject);
begin
inherited;
TT_RTRSCLIBTABLE0.State := TT_RTLIBCHPLIBSUSPECTS.State ;
TT_RTRSCLIBTABLE1.State := TT_RTRSCLIBTABLE0.State;
TT_RTRSCLIBTABLE2.State := TT_RTRSCLIBTABLE0.State;
TT_RTRSCLIBTABLE3.State := TT_RTRSCLIBTABLE0.State;
TT_RTRSCLIBTABLE4.State := TT_RTRSCLIBTABLE0.State;
TT_RTRSCLIBTABLE5.State := TT_RTRSCLIBTABLE0.State;
TT_RTRSCLIBTABLE6.State := TT_RTRSCLIBTABLE0.State;
TT_RTRSCLIBTABLE7.State := TT_RTRSCLIBTABLE0.State;
TT_RTRSCLIBTABLE8.State := TT_RTRSCLIBTABLE0.State;
TT_RTRSCLIBTABLE9.State := TT_RTRSCLIBTABLE0.State;
end;

////////////////////////////////////////////////////////////////////////////////
//*************Evénement sur la form et contrôle des données******************//
////////////////////////////////////////////////////////////////////////////////

procedure TFAssistInitSoc.ForceChecked(C1,C2 : TCheckBox) ;
begin
if C1.Checked then C2.Checked:=True ;
end ;

procedure TFAssistInitSoc.PageSuivante;
var Onglet : TTabSheet;
    st_NomPage : string;
    i_NumOnglet : integer;
begin
  inherited;
Onglet := P.ActivePage;
st_NomPage := Onglet.Name;
if  isNumeric (Copy (st_NomPage, length (st_NomPage), 1)) then
  begin
  i_NumOnglet := StrToInt (Copy (st_NomPage, length (st_NomPage), 1));
  if not okPage(i_NumOnglet) then
     begin
     RestorePage ;
     Onglet := NextPage ;
     if Onglet = nil then P.SelectNextPage(True)
     else begin
          P.ActivePage := Onglet ;
          PChange(nil) ;
          end ;
     PageSuivante;
     end;
  end;
end;


procedure TFAssistInitSoc.PagePrecedente;
var Onglet : TTabSheet;
    st_NomPage : string;
    i_NumOnglet : integer;
begin
  inherited;
Onglet := P.ActivePage;
st_NomPage := Onglet.Name;
if  isNumeric (Copy (st_NomPage, length (st_NomPage), 1)) then
  begin
  i_NumOnglet := StrToInt (Copy (st_NomPage, length (st_NomPage), 1));
  if not okPage(i_NumOnglet) then
     begin
     RestorePage ;
     Onglet := PreviousPage ;
     if Onglet = nil then P.SelectNextPage(True)
     else begin
          P.ActivePage := Onglet ;
          PChange(nil) ;
          end ;
     PagePrecedente;
     end;
  end;
end;

//Onglet visible si domaine selectionné
function TFAssistInitSoc.okPage(i_indice : integer) : boolean;
var i_ind : integer;
    Domaine : variant;
    Ctrl : TControl;
begin
Result := true;
if (i_indice = 0) and (FRef.ItemIndex = 0) then
   begin
   Result := false;
   exit;
   end;

for i_ind := 0 to GB_DOM.ControlCount -1 do
    begin
    Ctrl := GB_DOM.Controls[i_ind] ;
    if Ctrl is TCheckBox then
       begin
       Domaine := Copy(Ctrl.Name, length (Ctrl.Name), 1);
       if Domaine = i_indice then
          begin
          if (TCheckBox(FindComponent('CB_DOM'+intToStr(i_indice))).State = cbUnChecked)
             or (FRef.ItemIndex = 0) then Result := false;
          exit;
          end;
       end;
    end;
end;

//tous les tedit de tabInit doivent être renseignés
function TFAssistInitSoc.OkDonnees : boolean ;
var CtrlGB,CtrlEdt : TControl ;
    i_ind1,i_ind2 : integer ;
begin
Result := False ;
for i_ind1 := 0 to  TABINIT.ControlCount -1 do
  begin
  CtrlGB := TABINIT.Controls[i_ind1];
  if CtrlGB is TGroupBox then
  for i_ind2 := 0 to TGroupBox(CtrlGB).ControlCount-1 do
      begin
      CtrlEdt := TGroupBox(CtrlGB).Controls[i_ind2] ;
      if CtrlEdt is TEdit then
         if (TEdit(CtrlEdt).Text = '') then
            begin
            RestorePage;
            P.ActivePage := TABINIT ;
            PChange(nil) ;
            TEdit(CtrlEdt).SetFocus ;
            exit ;
            end ;
      end ;
  end ;
Result := True ;
end ;

//true si table ou tablette cochée ainsi que métier correspondant
function TFAssistInitSoc.TableSelect(stTable : string) : boolean;
var stNum : string;
    Ctrl : TComponent;
begin
Result := false;
if FindComponent(StTable) <> nil then
   if TCheckBox(FindComponent(StTable)).Checked then
      begin
      Ctrl := TCheckBox(FindComponent(StTable));
      while not(Ctrl is TTabSheet) and (Ctrl <> nil) do
            Ctrl := Ctrl.GetParentComponent;
      if Ctrl is TTabSheet then
         begin
         stNum := Copy(Ctrl.Name,Length(Ctrl.Name),1);
         if FindComponent('CB_DOM'+stNum) <> nil then
            begin
            if TCheckBox(FindComponent('CB_DOM'+stNum)).Checked then
               Result := true;
            end else Result := true; //TABINIT par exemple
         end;
      end;
end;

////////////////////////////////////////////////////////////////////////////////
//*********************************Gestion bases******************************//
////////////////////////////////////////////////////////////////////////////////

procedure TFAssistInitSoc.InitialiseSoc ;
var stDbSource : string;
begin
//Connection base de référence
ConnecteDB(GetSocRef,DBREF,'DBREF') ;
RDriver := V_PGI.Driver ;
RODBC   := V_PGI.ODBC ;
Entete.DatabaseName := 'DBREF' ;
LLigne.DatabaseName := 'DBREF' ;
DBREF.Connected := True ;
Entete.Open ;
LLigne.Open ;

//Connection base source
if FRef.ItemIndex = 0 then stDbSource := ''
else stDbSource := FRef.Items[FRef.ItemIndex];
ConnecteDB(stDbSource,DBSOURCE,'DBSOURCE') ;
SDriver := V_PGI.Driver ;
SODBC   := V_PGI.ODBC ;
DBSOURCE.Connected := True ;

//Connection base destination
ConnecteDB(FDestination,DBDest,'DBDest') ;
DDriver := V_PGI.Driver ;
DODBC   := V_PGI.ODBC ;
DBDest.Connected := True ;

V_PGI.Driver := SocDriver ;
V_PGI.ODBC   := SocODBC ;

CreerTables ;
end ;

procedure TFAssistInitSoc.CreerTables ;
var StLibTable,StTable,Prefixe : String ;
    Table  : THTable ;
    CtrlGB, Ctrl : TControl ;
    i_ind1, i_ind2, Nb{, VersionTable} : integer ;
    Prem, OkMajRupture, bTableSel   : Boolean ;
begin
Nb := NbRecords(Entete) + TABINIT.ControlCount ;
bTableSel := false;
TEtat.Visible := true;
EG_Progress.Visible := true;
EG_Progress.MaxValue := Nb;
TEtat.Caption := 'Connexion...';
Entete.First ;
OkMajRupture := false ;
While not Entete.EOF do
  begin
  StLibTable := Entete.FindField('DT_LIBELLE').AsString ;
  StTable := Entete.FindField('DT_NOMTABLE').AsString ;
//  VersionTable := Entete.FindField('DT_NUMVERSION').AsInteger ;
  EG_Progress.Progress := EG_Progress.Progress + 1;
  EG_Progress.ProgressText := Msg.Mess[5] + ' ' + StLibTable ;
  Application.ProcessMessages;

// Création de la table
  Prefixe := Entete.FindField('DT_PREFIXE').AsString ;
  LLigne.SetRange([Prefixe,000],[Prefixe,9999]) ;
  DBCreateTable(DBDest,Entete,LLigne,DDriver,False) ;
  if DDriver =  dbMSACCESS then ConnecteTRUEFALSE(DBDest) ;
  if DDriver <> dbMSSQL    then
    {$IFNDEF DBXPRESS}
    DBDest.StartTransaction ;
    {$ELSE}
    DBDest.DbxStartTransaction ;
    {$ENDIF}

// Insertion des données
  //Priorité DBSource si table selectionnée
  //sinon DBRef si renseigné dans les items de TableDefaut
  if stTable <> 'SOUCHE' then
     begin
     if FRef.ItemIndex > 0 then bTableSel := TableSelect(stTable);
     if ((FRef.ItemIndex > 0 ) and (bTableSel))
     or (TableDefaut.Items.IndexOf(StTable) > -1) then
        begin
        EG_Progress.ProgressText := Msg.Mess[6] + ' ' + StLibTable ;
        Application.ProcessMessages;
         if  (FRef.ItemIndex > 0 ) and  (bTableSel) then
             DBCopyTableBM(DBSOURCE,DBDest,StTable,StTable)
         else DBCopyTableBM(DBREF,DBDest,StTable,StTable);
    end;
    if StTable = 'RUPTURE' then OkMajRupture := true ;
    if StTable ='AXE' then
      begin
      Table := THTable.Create(Application) ;
      Table.DataBaseName := DBDest.DataBaseName ;
      Table.TableName := StTable ;
      Table.IndexName := 'X_CLE1' ;
      Table.Open ;
      While not Table.Eof do
        begin
        Table.Edit ;
        Table.FindField('X_SECTIONATTENTE').AsString:='' ;
        Table.FindField('X_GENEATTENTE').AsString:='' ;
        Table.FindField('X_BOURREANA').asString:='0' ;
        Table.Post ;
        Table.Next ;
        end ;
      Table.Close ;
      Table.Free ;
      end ;
    end ;

  if DDriver <> dbMSSQL then
    {$IFNDEF DBXPRESS}
    DBDest.Commit ;
    {$ELSE}
    DBDest.DbxCommit ;
    {$ENDIF}
  Entete.Next ;
  end ;


TEtat.Caption := '';
TProgress.Caption := '';
TProgress.visible := true;
EG_Progress.Visible := false;
//Recherche des tables UTILISAT et USERGRP
Table := THTable.Create(Application) ;
Table.DataBaseName := DBDest.DataBaseName ;
for i_ind1 := 0 to TABINIT.ControlCount -1 do
   begin
   CtrlGB := TABINIT.Controls[i_ind1] ;
   if CtrlGB is TGroupBox then
      begin
      Table.TableName := Copy(CtrlGB.Name,2,length(CtrlGB.Name)-1) ;
      Table.IndexName := CtrlGB.Hint ;
      Table.Open ;
      TEtat.Caption := Msg.Mess[7] ;
      TProgress.caption := Table.TableName ;
      Application.ProcessMessages ;
      Prem := True ;
      if Pos('USERGRP',Table.TableName)>0 then
         begin
         Prem := not Table.FindKey([UG_GROUPE.Text]) ;
         end else if Pos('UTILISAT',Table.TableName) > 0 then
             begin
             Prem := not Table.FindKey([US_UTILISATEUR.Text]) ;
             end ;
      if Prem then
         begin
         Table.Insert ;
         InitNew(Table) ;
         end else Table.Edit ;

      //Insertion des données des controles dans les tables créées
      for i_ind2:=0 to TGroupBox(CtrlGB).ControlCount-1 do
         begin
         Ctrl := TGroupBox(CtrlGB).Controls[i_ind2] ;
         if (Table.FindField(Ctrl.Name) <> Nil) then
           if (Ctrl.Name = 'US_PASSWORD') then
                Table.FindField(Ctrl.Name).Value := CryptageSt(TEdit(Ctrl).Text)
           else Table.FindField(Ctrl.Name).Value := TEdit(Ctrl).Text ;
         end ;
      if Pos('UTILISAT',Table.TableName) > 0 then
          begin
         Table.FindField('US_CRC').AsInteger := GetCRC32ForData(Table) ;
         end;
      Table.Post ;
      Table.Close ;
      end ;
   MoveCur(False) ;
   end ;

Table.Free ;
If OkMajRupture Then MajRupture ;
TEtat.Caption := Msg.Mess[11];
DBCreateAllView(DBDest,TProgress,DDriver) ;
InitialiseLesTables ;
//FiniMove ;
end ;

//nb de tables
function TFAssistInitSoc.Nbrecords(Table : TDataSet) : Integer ;
var Nb : Integer ;
begin
Table.First ;
Nb := 0 ;
While not Table.Eof do
    begin
    Inc(Nb) ;
    Table.Next ;
    end ;
Result := Nb ;
end ;


Procedure TFAssistInitSoc.MajRupture ;
var QDest,QSource : THQuery ;
begin
QSource := THQuery.Create(Application) ;
QSource.DataBaseName := DBSOURCE.DataBaseName ;
QSource.SessionName  := DBSOURCE.SessionName ;

QDest := THQuery.Create(Application) ;
QDest.DataBaseName := DBDest.DataBaseName ;
QDest.SessionName  := DBDest.SessionName ;
QDest.RequestLive  := true ;

QSource.SQL.Clear ;
QSource.SQL.Add('Select * from Choixcod where CC_TYPE="RUG"') ;
V_PGI.Driver := SDriver ;
V_PGI.ODBC   := SODBC ;
ChangeSQL(QSource) ;
QSource.Open ;
V_PGI.Driver := SocDriver ;
V_PGI.ODBC   := SocODBC ;

QDest.SQL.Clear ;
QDest.SQL.Add('Select * From CHOIXCOD where CC_TYPE="'+W_W+'"') ;
V_PGI.Driver := DDriver ;
V_PGI.ODBC   := DODBC ;
ChangeSQL(QDest) ;
QDest.Open ;
V_PGI.Driver := SocDriver ;
V_PGI.ODBC   := SocODBC ;

While not QSource.Eof Do
  begin
  QDest.Insert ;
  InitNew(QDest) ;
  QDest.FindField('CC_TYPE').AsString := QSource.FindField('CC_TYPE').AsString ;
  QDest.FindField('CC_CODE').AsString := QSource.FindField('CC_CODE').AsString ;
  QDest.FindField('CC_LIBELLE').AsString := QSource.FindField('CC_LIBELLE').AsString ;
  QDest.FindField('CC_ABREGE').AsString  := QSource.FindField('CC_ABREGE').AsString ;
  QDest.FindField('CC_LIBRE').AsString   := QSource.FindField('CC_LIBRE').AsString ;
  QDest.Post ;
  QSource.Next ;
  end ;

QSource.Close ;
QDest.Close ;
QSource.SQL.Clear ;
QSource.SQL.Add('DELETE From RUPTURE where RU_NATURERUPT="RUG"') ;
V_PGI.Driver := SDriver ;
V_PGI.ODBC   := SODBC ;
ChangeSQL(QSource) ;
QSource.ExecSQL ;

V_PGI.Driver := SocDriver ;
V_PGI.ODBC   := SocODBC ;
QSource.Free ;
QDest.Free ;
end ;

procedure TFAssistInitSoc.InitialiseLesTables ;
var Val : Variant ;
    Table : THTable ;
    i_ind : integer ;
    CodeD : String ;
    Ctrl  : TComponent;
    QSource : THQuery ;
begin
TEtat.Caption := Msg.Mess[12];
////////////////////////Mise à jour des tablettes ChoixCod
for i_ind := 0 to ComponentCount -1 do
    begin
    if Components[i_ind] is TCheckBox then
       begin
       Ctrl := Components[i_ind];
       if Copy(Ctrl.Name,1,3) = 'TT_' then
          begin
          if (FRef.ItemIndex > 0) and (TableSelect(Ctrl.Name)) then
                  RecopieTablettes(Copy(Ctrl.Name,4,Length(Ctrl.Name)-3),'DBSOURCE')
          else //if (TTDefaut.Items.IndexOf(Ctrl.Name) > -1)then
                  RecopieTablettes(Copy(Ctrl.Name,4,Length(Ctrl.Name)-3),'DBREF');
          end;
       end;
    end;

Val := '' ;
RecupChpDB('DBREF','SOCIETE','SO_VERSIONBASE','',Val) ;
Table:=THTable.Create(Application) ;
Table.DataBaseName:=DBDest.DataBaseName ;
// Societe
Table.TableName := 'SOCIETE' ;
Table.IndexName:='SO_CLE1' ;
Table.Open ;
if Table.Eof then
   begin
   Table.Close ;
   exit ;
   end;
Table.Edit ;
Table.FindField('SO_VERSIONBASE').Value:=Val ;
Table.FindField('SO_SOCIETE').AsString:=SO_SOCIETE.Text ;
Table.FindField('SO_LIBELLE').AsString:=SO_LIBELLE.Text ;
Table.Post ;
Table.Close ;

// SuivDec
Table.TableName:='CHOIXCOD' ;
Table.IndexName := 'CC_CLE1' ;
Table.Open ;
for i_ind:=1 to 5 do
    begin
    if Table.FindKey(['CID','CD'+IntToStr(i_ind)]) then
       Table.Edit
    else Table.Insert ;
    InitNew(Table) ;
    Table.FindField('CC_TYPE').AsString := 'CID' ;
    Table.FindField('CC_CODE').AsString := 'CD'+IntToStr(i_ind) ;
    Table.FindField('CC_LIBELLE').AsString := '' ;
    Table.FindField('CC_ABREGE').AsString  := '' ;
    Table.FindField('CC_LIBRE').AsString   := '' ;
    Table.Post ;
    end ;
Table.Close ;

///////////////Initialisation des catégories d'article
if not TT_GCCATEGORIEDIM.Checked then
   begin
   Table.TableName:='CHOIXCOD' ;
   Table.IndexName := 'CC_CLE1' ;
   Table.Open ;
   for i_ind:=1 to 5 do
       begin
       if Table.FindKey(['DIM','DI'+IntToStr(i_ind)]) then
          Table.Edit
          else Table.Insert ;
       InitNew(Table) ;
       Table.FindField('CC_TYPE').AsString := 'DIM' ;
       Table.FindField('CC_CODE').AsString := 'DI'+IntToStr(i_ind) ;
       Table.FindField('CC_LIBELLE').AsString := 'Dimension '+IntToStr(i_ind) ;
       Table.FindField('CC_ABREGE').AsString  := '' ;
       Table.FindField('CC_LIBRE').AsString   := '' ;
       Table.Post ;
       end ;
   Table.Close ;
   end;

///////////////Récupération des souches dont Natures pièce présentes dans Parpièce

Table.TableName:='SOUCHE' ;
Table.IndexName := 'SH_CLE1' ;
Table.Open ;

QSource := THQuery.Create(Application) ;
if (FRef.ItemIndex > 0) and (TableSelect('SOUCHE')) then
   begin
   QSource.DataBaseName := DBSOURCE.DataBaseName ;
   QSource.SessionName  := DBSOURCE.SessionName ;
   end else
   begin
   QSource.DataBaseName := DBREF.DataBaseName ;
   QSource.SessionName  := DBREF.SessionName ;
   end;

QSource.SQL.Clear ;
QSource.SQL.Add('SELECT * FROM SOUCHE WHERE SH_NATUREPIECEG IN '+
                                    '(SELECT GPP_NATUREPIECEG FROM PARPIECE)') ;
V_PGI.Driver := SDriver ;
V_PGI.ODBC   := SODBC ;
ChangeSQL(QSource) ;

QSource.Open ;
V_PGI.Driver := SocDriver ;
V_PGI.ODBC   := SocODBC ;

While not QSource.Eof Do
  begin
  Table.Insert;
  Table.Edit;
  Table.FindField('SH_TYPE').AsString         := QSource.FindField('SH_TYPE').AsString;
  Table.FindField('SH_NUMDEPART').AsString    := QSource.FindField('SH_NUMDEPART').AsString;
  Table.FindField('SH_SOUCHE').AsString       := QSource.FindField('SH_SOUCHE').AsString;
  Table.FindField('SH_JOURNAL').AsString      := QSource.FindField('SH_JOURNAL').AsString;
  Table.FindField('SH_LIBELLE').AsString      := QSource.FindField('SH_LIBELLE').AsString;
  Table.FindField('SH_MASQUENUM').AsString    := QSource.FindField('SH_MASQUENUM').AsString;
  Table.FindField('SH_ABREGE').AsString       := QSource.FindField('SH_ABREGE').AsString;
  Table.FindField('SH_SOCIETE').AsString      := QSource.FindField('SH_SOCIETE').AsString;
  Table.FindField('SH_NATUREPIECE').AsString  := QSource.FindField('SH_NATUREPIECE').AsString;
  Table.FindField('SH_DATEDEBUT').AsString    := QSource.FindField('SH_DATEDEBUT').AsString;
  Table.FindField('SH_DATEFIN').AsString      := QSource.FindField('SH_DATEFIN').AsString;
  Table.FindField('SH_FERME').AsString        := QSource.FindField('SH_FERME').AsString;
  Table.FindField('SH_ANALYTIQUE').AsString   := QSource.FindField('SH_ANALYTIQUE').AsString;
  Table.FindField('SH_NATUREPIECEG').AsString := QSource.FindField('SH_NATUREPIECEG').AsString;
  Table.FindField('SH_SIMULATION').AsString   := QSource.FindField('SH_SIMULATION').AsString;
  Table.FindField('SH_NUMDEPARTS').AsString   := QSource.FindField('SH_NUMDEPARTS').AsString;
  Table.FindField('SH_NUMDEPARTP').AsString   := QSource.FindField('SH_NUMDEPARTP').AsString;
  Table.FindField('SH_SOUCHEEXO').AsString    := QSource.FindField('SH_SOUCHEEXO').AsString;
  Table.FindField('SH_RESERVEWEB').AsString   := QSource.FindField('SH_RESERVEWEB').AsString;
  Table.Post ;
  QSource.Next ;
  end ;

QSource.Close ;
QSource.Free ;
Table.Close;

/////////////// Si pas recopie des devises, créer au moins la devise principale
if ((FRef.ItemIndex <= 0) or (not Devise.Checked)) then
   begin
   Table.TableName := 'PARAMSOC' ;
   Table.IndexName := 'SOC_CLE1' ;
   Table.Open ;
   if Table.FindKey(['SO_DEVISEPRINC']) then
        CodeD := Table.FindField('SOC_DATA').AsString
   else CodeD := 'FRF' ;
   Table.Close ;
   Table.TableName := 'DEVISE' ;
   Table.IndexName := 'D_CLE1' ;
   Table.Open ;
   Table.Insert ;
   InitNew(Table) ;
   Table.FindField('D_DEVISE').AsString    := CodeD  ;
   Table.FindField('D_LIBELLE').AsString   := 'Francs' ;
   Table.FindField('D_DECIMALE').AsInteger := 2 ;
   Table.FindField('D_QUOTITE').AsFloat    := 1 ;
   Table.FindField('D_MONNAIEIN').AsString := 'X' ;
   Table.FindField('D_FONGIBLE').AsString  := 'X' ;
   Table.FindField('D_PARITEEURO').AsFloat := 6.55957 ;
   Table.Post ;
   Table.Close ;
   Table.TableName:='PARAMSOC' ;
   Table.IndexName:='SOC_CLE1' ;
   Table.Open ;
   if Table.FindKey(['SO_TAUXEURO']) then
      begin
      Table.Edit ;
      Table.FindField('SOC_DATA').AsVariant := 6.55957 ;
      Table.Post ;
      end ;
   if Table.FindKey(['SO_DECVALEUR']) then
      begin Table.Edit ;
      Table.FindField('SOC_DATA').AsVariant := 2 ;
      Table.Post ;
      end ;
   if Table.FindKey(['SO_DEVISEPRINC']) then
      begin
      Table.Edit ;
      Table.FindField('SOC_DATA').AsVariant := CodeD ;
      Table.Post ;
      end ;
   if Table.FindKey(['SO_TENUEEURO']) then
      begin
      Table.Edit ;
      Table.FindField('SOC_DATA').AsVariant := '-' ;
      Table.Post ;
      end ;
   if Table.FindKey(['SO_DATEDEBUTEURO']) then
      begin
      Table.Edit ;
      Table.FindField('SOC_DATA').AsVariant := 36161 ;
      Table.Post ;
      end ;
   if Table.FindKey(['SO_SOCIETE']) then
      begin
      Table.Edit ; Table.FindField('SOC_DATA').AsVariant := SO_SOCIETE.Text ;
      Table.Post ;
      end ;
   if Table.FindKey(['SO_LIBELLE']) then
      begin
      Table.Edit ;
      Table.FindField('SOC_DATA').AsVariant := SO_LIBELLE.Text ;
      Table.Post ;
      end ;
   Table.Close ;
   end ;
Table.Free ;
SpecifMetier1;
SpecifMetier2;
SpecifMetier3;
SpecifMetier4;
SpecifMetier5;
end ;

function TFAssistInitSoc.RecupChpDB (DB,Fichier,Champ,Where : String ; var Valeur : Variant ) : Boolean ;
Var Q : TQuery ;
    SQL : String ;
begin
Q := TQuery.Create(Application) ;
Q.DatabaseName := DB ;
SQL := 'Select ' + Champ + ' from ' + Fichier + Where ;
Q.SQL.Clear ;
Q.SQL.Add(SQL) ;
V_PGI.Driver := RDriver ;
V_PGI.ODBC   := RODBC ;
ChangeSQL(Q) ;
V_PGI.Driver := SocDriver ;
V_PGI.ODBC   := SocODBC ;
Q.RequestLive:= False ;
Q.Open ;
Result := (not Q.EOF) ;
if not Q.Eof then Valeur := Q.Fields[0].Value ;
Q.Close ;
Q.Free ;
end ;


Procedure TFAssistInitSoc.RecopieTablettes(Tablette, Base : string);
var QSource : THQuery ;
    TableDest : THTable ;
    stType : string;
    i_dim : integer;
begin
//mng
if Tablette = 'RTLIBCLACTPERSP' then exit;

if (Tablette = 'TT_GCCATEGORIEDIM') and (FRef.ItemIndex > 0) then exit; // Initialisation dans InitialiseTable
stType := '';
TProgress.Caption := Tablette;
if Tablette = 'DIMENSIONS' then
   begin
   for i_dim := 1 to MaxDimension do
       begin
       if   stType = '' then  stType := '"GG' + intToStr(i_dim) + '"'
       else stType := StType + ' OR CC_TYPE="' + 'GG' + intToStr(i_dim) + '"';
       end;
   end else
   begin
   stType := RechercheChoixType(Tablette,Base);
   if stType = '' then exit;
   stType := '"'+stType+'"';
   end;
QSource := THQuery.Create(Application) ;
if Base = 'DBSOURCE' then
   begin
   QSource.DataBaseName := DBSOURCE.DataBaseName ;
   QSource.SessionName  := DBSOURCE.SessionName ;
   end else
   begin
   QSource.DataBaseName := DBREF.DataBaseName ;
   QSource.SessionName  := DBREF.SessionName ;
   end;

TableDest := THTable.Create(Application) ;
TableDest.DataBaseName := DBDest.DataBaseName ;
TableDest.TableName := 'CHOIXCOD' ;
TableDest.IndexName:='CC_CLE1' ;
TableDest.Open;

QSource.SQL.Clear ;
QSource.SQL.Add('SELECT * FROM CHOIXCOD WHERE CC_TYPE=' + stType) ;
V_PGI.Driver := SDriver ;
V_PGI.ODBC   := SODBC ;
ChangeSQL(QSource) ;

QSource.Open ;
V_PGI.Driver := SocDriver ;
V_PGI.ODBC   := SocODBC ;

While not QSource.Eof Do
  begin
  TableDest.Insert;
  TableDest.Edit;
  TableDest.FindField('CC_TYPE').AsString    := QSource.FindField('CC_TYPE').AsString ;
  TableDest.FindField('CC_CODE').AsString    := QSource.FindField('CC_CODE').AsString ;
  TableDest.FindField('CC_LIBELLE').AsString := QSource.FindField('CC_LIBELLE').AsString ;
  TableDest.FindField('CC_ABREGE').AsString  := QSource.FindField('CC_ABREGE').AsString ;
  TableDest.FindField('CC_LIBRE').AsString   := QSource.FindField('CC_LIBRE').AsString ;
  TableDest.Post ;
  QSource.Next ;
  end ;

QSource.Close ;
TableDest.Close ;
QSource.Free ;
end ;

//Recherche du do_type tablette dans decombos
function TFAssistInitSoc.RechercheChoixType(Tablette,Base : string) : string;
var QSource : THQuery ;
begin
Result := '';
QSource := THQuery.Create(Application) ;
if Base = 'DBSOURCE' then
   begin
   QSource.DataBaseName := DBSOURCE.DataBaseName ;
   QSource.SessionName  := DBSOURCE.SessionName ;
   end else
   begin
   QSource.DataBaseName := DBREF.DataBaseName ;
   QSource.SessionName  := DBREF.SessionName ;
   end;
   QSource.SQL.Clear ;
QSource.SQL.Add('SELECT DO_TYPE FROM DECOMBOS WHERE DO_COMBO="' + Tablette + '"') ;
V_PGI.Driver := SDriver ;
V_PGI.ODBC   := SODBC ;
ChangeSQL(QSource) ;

QSource.Open ;
V_PGI.Driver := SocDriver ;
V_PGI.ODBC   := SocODBC ;
if not QSource.Eof then
    Result := QSource.FindField('DO_TYPE').AsString ;

QSource.Close ;
QSource.Free ;
end;

////////////////////////////////////////////////////////////////////////////////
//************************** Spécificités par métier *************************//
////////////////////////////////////////////////////////////////////////////////


//////////////////////////Gescom
procedure TFAssistInitSoc.SpecifMetier1;
begin

end;

//////////////////////////Compta
procedure TFAssistInitSoc.SpecifMetier2;
begin

end;

//////////////////////////Affaires
procedure TFAssistInitSoc.SpecifMetier3;
begin
// if CB_DOM3.checked then  ...
// MaFonctionMetier( CB_DOM3.Checked ) ;
{$IFDEF AFFAIRE}
AffaireInitSoc(CB_DOM3.Checked,CB_DOM1.Checked);
{$ENDIF AFFAIRE}
end;

//////////////////////////Mode
procedure TFAssistInitSoc.SpecifMetier4;
begin

end;

//////////////////////////Relation clients
procedure TFAssistInitSoc.SpecifMetier5;
begin

end;


end.
