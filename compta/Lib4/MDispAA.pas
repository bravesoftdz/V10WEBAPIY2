unit MDispAA ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  UIUtil, ExtCtrls, {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF}, StdCtrls, Mask, DBCtrls, Db, Grids, DBGrids,
  Menus, ComCtrls, HPanel, HStatus, HMenu97, Hctrls, HDB,UtilPGI,
  HTB97,HEnt1,Ent1,Ed_Tools, HRotLab, hmsgbox, HFLabel, Courrier, MulCour,
  LicUtil,HCalc,UserChg, SQL,HMacro,AddDico, HCapCtrl,IniFiles, Seria,Prefs,
  Buttons,Xuelib,About,Traduc,TradForm,Tradmenu,Graph,
  FE_Main,UIUtil3, ImgList, Hout, EdtEtat, PGIExec, dbgInfo ;

Procedure InitApplication ;
procedure DispatchTT(Num : Integer; Action : TActionFiche; Lequel,TT,Range : String) ;

type
  TFMDispMP = class(TDatamodule)
    end ;

Var FMDispMP : TFMDispMP ;

implementation

uses MenuOLG, CpteUtil ;

{$R *.DFM}

Procedure Dispatch ( Num : Integer ; PRien : THPanel ; Var RetourForce,FermeHalley : boolean );
Var Nature : string ;
BEGIN
MajBeforeDispatch(Num) ;
   case Num of
        10 : BEGIN
             UpdateComboslookUp ; ChargeHalleyUser ;
             END ;
        11 : ; // Après déconnexion
        12 : ; // Avant protec
        13 : ;

   8205  : AGLLanceFiche('CP','CPGENERECROISEAXE','','','') ;
   8210  : AGLLanceFiche('CP','CPCROISEAXE_CUBE','','','') ;
   8215  : AGLLanceFiche('CP','CPCROISEAXE_STAT','','','') ;
   8220  : AGLLanceFiche('CP','CPCROISEAXE_ETAT','','','') ;

     else HShowMessage('2;?caption?;'+TraduireMemoire('Fonction non disponible : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)) ;
   end ;
END ;

Procedure AfterProtec ( sAcces : String ) ;
BEGIN
V_PGI.VersionDemo:=False ;
END ;

procedure AssignLesZoom ;
BEGIN
ProcZoomEdt:=Nil ;
END ;

procedure AfterChangeModule ( NumModule : integer ) ;
BEGIN
UpdateSeries ;
AssignLesZoom ;
if FMenuG.PopupMenu=Nil then FMenuG.PopUpMenu:=ADDMenuPop(Nil,'','') ;
END ;

procedure AssignZoom ;
BEGIN
AssignLesZoom ;
ProcGetVH:=GetVH ;
ProcGetDate:=GetDate ;
END ;

procedure InitApplication ;
BEGIN
AssignZoom ;
FMenuG.OnDispatch:=Dispatch ;
FMenuG.OnChangeModule:=AfterChangeModule ;
FMenuG.OnChargeMag:=ChargeMagHalley ;
FMenuG.OnAfterProtec:=AfterProtec ;
FMenuG.OnMajAvant:=Nil ;
FMenuG.OnMajApres:=Nil ; 
FMenuG.SetModules([8],[]) ;
END ;

procedure DispatchTT(Num : Integer; Action : TActionFiche; Lequel,TT,Range : String) ;
begin
end ;

procedure RenseigneSerie ;
var Buffer: array[0..1023] of Char;
    sWinPath,sIni : String ;
begin
HalSocIni:='CEGIDPGI.INI' ;
NomHalley:='Analytique S5' ;
TitreHalley := 'Analytique avancé S5' ;
V_PGI.LaSerie:=S5 ;
Application.HelpFile:=ExtractFilePath(Application.ExeName)+'CCS5.HLP' ;
END ;

initialization
RenseigneSerie ;
Apalatys:='CEGID' ;
Copyright:='© Copyright ' + Apalatys ;
V_PGI.OutLook:=True ;
V_PGI.VersionDemo:=True ;
V_PGI.VersionReseau:=True ;
V_PGI.MenuCourant:=0 ;
V_PGI.NumVersion:='3.82' ; V_PGI.NumBuild:='535' ; V_PGI.NumVersionBase:=535 ;
V_PGI.DateVersion:=EncodeDate(2001,05,10) ;
V_PGI.VersionDEV:=TRUE ;
V_PGI.ImpMatrix := True ;
V_PGI.OKOuvert:=FALSE ;
V_PGI.Halley:=TRUE ;
//V_PGI.NiveauAcces:=1 ;
V_PGI.OfficeMsg:=True ;
V_PGI.NumMenuPop:=27 ;
V_PGI.Decla100:=TRUE ;
V_PGI.DispatchTT:=DispatchTT ;
V_PGI.ParamSocLast:=False ;
V_PGI.RAZForme:=TRUE ;
V_PGI.NoModuleButtons:=False ;
VH^.GereSousPlan:=FALSE ;
V_PGI.PGIContexte:=[ctxCompta] ;
V_PGI.MajPredefini:=False ;
V_PGI.BlockMAJStruct:=True ;
V_PGI.EuroCertifiee:=False ;

ChargeXuelib ;

end.
