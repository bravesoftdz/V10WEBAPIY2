unit MenuDisp;

interface
Uses HEnt1,MenuOLG,Forms,sysutils,HMsgBox, Classes, Controls, HPanel, UIUtil,ImgList,
     Tablette,FE_Main,EdtEtat,AGLInit ;

Procedure InitApplication ;

type
  TFMenuDisp = class(TDatamodule)
    ImageList1: TImageList;
    end ;

Var FMenuDisp : TFMenuDisp ;

implementation

uses HCtrls, ETransferts, EntGC, Ent1, TestMSMQ ;

{$R *.DFM}

Procedure Dispatch ( Num : Integer ; PRien : THPanel ; var retourforce,sortiehalley : boolean);
BEGIN
   case Num of
     10 : ; //Apres connection
     11 : ; //Après deconnection
     12 : ; //Avant connection ou seria
     13 : ChargeMenuPop(TTypeMenu(-1),FMenuG.Dispatch) ; // avant déconnexion
     1105 :  ;

     59005 : LanceMSMQ ;
//     59005 : TestTarifWEB ;

     59101 : ExportAllTables;         
     59102 : ImportAllTables;
//     59103 : BalanceLesCommandes;
     59103 : LanceMSMQ ;
     59201 : ImporteECommande;
     59301 : AGLLanceFiche('E','EEXPORTARTICLE','','','');
     59302 : AGLLanceFiche('E','EEXPORTTIERS','','','');
     59303 : AGLLanceFiche('E','EEXPORTTARIF','','','');
     59304 : AGLLanceFiche('E','EEXPORTCOMM','','','');
     59305 : AGLLanceFiche('E','EEXPORTPAYS','','','');

     else HShowMessage('2;?caption?;'+TraduireMemoire('Fonction non disponible : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)) ;
     end ;
END ;

Procedure DispatchTT ( Num : Integer ; Action : TActionFiche ; Lequel,TT,Range : String ) ;
BEGIN
   case Num of
     1 : ;
     end ;
END ;


procedure InitApplication ;
BEGIN
//ProcZoomEdt:=ZoomEdtEtat ;
//ProcCalcEdt:=CalcOLEEtat ;
//ProcGetVH:=GetVS1 ;
//ProcGetDate:=GetDate ;
FMenuG.OnDispatch:=Dispatch ;
FMenuG.OnChargeMag:=ChargeMagHalley ;
FMenuG.OnMajAvant:=Nil ;
FMenuG.OnMajApres:=Nil ;
//FMenuG.SetModules([30,31,32,33,34,65,60],[35,18,5,11,28,49]) ;
FMenuG.SetModules([59,60],[32,49]);
V_PGI.DispatchTT:=DispatchTT ;

END ;

initialization
Apalatys:='CEGID' ;                      
NomHalley:='E-Commerce' ;
TitreHalley:='CEGID E-Commerce' ;
HalSocIni:='CEGIDPGI.ini';
Copyright:='© Copyright ' + Apalatys ;
V_PGI.NumVersion:='1.00 ß' ; V_PGI.NumBuild:='0' ;
V_PGI.DateVersion:=EncodeDate(2000,09,01);
V_PGI.LaSerie:=S5 ;

V_PGI.OutLook:=TRUE ;
V_PGI.OfficeMsg:=TRUE ;
V_PGI.ToolsBarRight:=TRUE ;
ChargeXuelib ;
V_PGI.VersionDemo:=TRUE ;
V_PGI.VersionReseau:=False ;
V_PGI.ImpMatrix := True ;
V_PGI.OKOuvert:=FALSE ;
V_PGI.Halley:=TRUE ;
V_PGI.NiveauAcces:=1 ;
V_PGI.MenuCourant:=0 ;
end.

