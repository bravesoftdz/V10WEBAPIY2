unit MDispRC;

interface

Uses HEnt1,MenuOLG,Forms,sysutils,HMsgBox, Classes, Controls, HPanel
     , UIUtil,ImgList,
     FE_Main,EdtEtat,AGLInit , tablette,  ParamSoc, Hctrls
     ,pays, region , codepost, devise , ModifZS , RTParamZL;

Procedure InitApplication ;
procedure AfterChangeModule ( NumModule : integer ) ;

type
  TFMenuDisp = class(TDatamodule)
    end ;

Var FMenuDisp : TFMenuDisp ;

implementation

uses Ent1, UtilPGI, EntGC
            ,ZoomEdtProspect
{admin.}    ,UserGrp,User
            ;
        
{$R *.DFM}

Procedure Dispatch ( Num : Integer ; PRien : THPanel ; var retourforce,sortiehalley : boolean);
var Rafale :Boolean ;
    PCSRecup : string ;
BEGIN
   case Num of
   10 : BEGIN   //  après connexion
        UpdateCombosGC ;
        //AfterChangeModule(65);
        //AfterChangeModule(92);
        END ;
   11 : ; //  après déconnexion
   12 : ; //  avant connexion et séria
   13 : ChargeMenuPop(TTypeMenu(-1),FMenuG.Dispatch) ; // avant déconnexion


   //////// Fiches ///////////
   92101 : AGLLanceFiche('RT','RTPROSPECT_MUL','','','');
   92102 : AGLLanceFiche('YY','YYCONTACTTIERS','','','');
   92103 : AGLLanceFiche('RT','RTOPERATIONS_MUL','','','');
   92104 : AGLLanceFiche('RT','RTACTIONS_MUL','','','');
   92105 : AGLLanceFiche('RT','RTPERSPECTIVE_MUL','','','');
    92106 : AGLLanceFiche('RT','RTGENERE_ACTIONS','','','');

   ///// Editions ///////////////////
   92301 : AGLLanceFiche('RT','RTACTIONS_ETAT2','','','') ;
   92302 : AGLLanceFiche('RT','RTOPERATIONS_ETAT','','','') ;
   92303 : AGLLanceFiche('GC','GCETIQCLI','','','') ;
   92304 : AGLLanceFiche('RT','RTPERSPECTIV_ETAT','','','') ;
   92305 : AGLLanceFiche('RT','RTFICHEPROSPECT','','','') ;
   92306 : AGLLanceFiche('RT','RTOPER_ETAT_COUT','','','') ;

   ////// Traitements///////////////
   92202 : AGLLanceFiche('RT','RTGENERE_ACTIONS','','','');  // lancement opération
   92203 : AGLLanceFiche('RT','RTPROS_MUL_MAILIN','','','') ;

   ///// Analyse ////////

   92402 : AGLLanceFiche('RT','RTANAPROSP','','','');
   92403 : AGLLanceFiche('RT','RTPERSPECTIVE_TV','','','');
   92404 : AGLLanceFiche('RT','RTACTIONS_TV','','','');

   //92202 : PCSRecup:=ModifZoneSerie('PROSPECTS','', Rafale)   ;

   //60399 :  AGLLanceFiche('GC','GCARTICLE_RECH','','','');

    /////// Paramètres - Société //////////
   //65201 : ParamSociete (TRUE,'','',Nil,Nil,Nil,Nil,0);
   //5202 : ParamSociete (FALSE,'','',Nil,Nil,Nil,Nil,0);
   //65203 : FicheEtablissement_AGL(taModif);
   //65204 : FicheUserGrp ;
   //65205 : BEGIN FicheUSer(V_PGI.User) ; ControleUsers ; END ;

   ////// Tablettes //////////
   92911 : ParamTable('RTTYPEACTION',taModif,0,PRien) ;
   92912 : ParamTable('RTTYPEPERSPECTIVE',taModif,0,PRien) ;
   92913 : ParamTable('RTETATACTION',taModif,0,PRien) ;
   92914 : ParamTable('RTETATPERSPECTIVE',taModif,0,PRien) ;
   92915 : ParamTable('RTLIBONGLET',taModif,0,PRien) ;

   92920 : AGLLanceFiche ('GC','GCCOMMERCIAL_MUL','','','') ; //ParamTable('GCREPRESENTANT',taModif,0,PRien) ;

   92930 : begin
           ParamTable('RTLIBCHAMPSLIBRES',taModif,0,PRien) ;
           AvertirTable ('RTLIBCHAMPSLIBRES') ;
           AfterChangeModule(92);
           end ;

   92940 : AGLLanceFiche('RT','RTPARAMCL','','',''); //ParamSaisieTL ();                           
   92951..92960 : ParamTable('RTRPRLIBTABLE'+IntToStr(Num-92951),taCreat,0,PRien) ;
   92971..92980 : ParamTable('RTRPRLIBTABMUL'+IntToStr(Num-92971),taCreat,0,PRien) ;
   92991..92993 : ParamTable('RTRPRLIBACTION'+IntToStr(Num-92990),taCreat,0,PRien) ;
   92996..92998 : ParamTable('RTRPRLIBPERSPECTIVE'+IntToStr(Num-92995),taCreat,0,PRien) ;

   65607 : ParamTable('GCEMPLOIBLOB',taCreat,0,PRien) ;
   65661 : ParamTable('GCZONECOM',taCreat,0,PRien);

   //// Paramètres généraux ///
   65901 : ParamTable('ttFormeJuridique',taCreat,1120000,PRien) ;
   65903 : ParamTable('ttCivilite',taCreat,1122000,PRien) ;
   65902 : ParamTable('ttFonction',taCreat,1125000,PRien) ;
   65904 : OuvrePays ;
   65905 : FicheRegion('','',False) ;
   65906 : OuvrePostal(PRien) ;
   65907 : FicheDevise('',tamodif,False) ;


// outils

//65901 : VerifCodeArticleUnique ;

     else HShowMessage('2;?caption?;'+TraduireMemoire('Fonction non disponible : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)) ;
     end ;

END ;

Procedure DispatchTT ( Num : Integer ; Action : TActionFiche ; Lequel,TT,Range : String ) ;
BEGIN
  case Num of
    20 : AGLLanceFiche('GC','GCTIERS','',Lequel,ActionToString(Action)+';MONOFICHE') ;
    30 : AGLLanceFiche('YY','YYCONTACT','',Lequel,ActionToString(Action)+';MONOFICHE') ;
  end ;
END ;

procedure AfterChangeModule ( NumModule : integer ) ;
Var i : integer ;
BEGIN
Case NumModule of
  65 :  begin
        { Categories de dimensions }
        for i:=1 to 5 do
            FMenuG.RenameItem(65310+i,RechDom('GCCATEGORIEDIM','DI'+IntToStr(i),FALSE)) ;
        { Libellé des codes famille }
        for i:=1 to 3 do
            FMenuG.RenameItem(65620+i,RechDom('GCLIBFAMILLE','LF'+IntToStr(i),FALSE)) ;
        end;
        { Tables libres Prospects }
   { Champs libres Combos }
   92 :  begin
         for i:=1 to 10 do
            begin
         // Tables libres
            FMenuG.RenameItem(92950+i,RechDom('RTLIBCHAMPSLIBRES','CL'+IntToStr(i-1),FALSE)) ;
            FMenuG.RenameItem(92970+i,RechDom('RTLIBCHAMPSLIBRES','ML'+IntToStr(i-1),FALSE)) ;
            end;
         for i:=1 to 3 do
            begin
            // Actions
            FMenuG.RenameItem(92990+i,RechDom('RTLIBCHAMPSLIBRES','AL'+IntToStr(i),FALSE)) ;
            // Perspectives
            FMenuG.RenameItem(92995+i,RechDom('RTLIBCHAMPSLIBRES','PL'+IntToStr(i),FALSE)) ;
            end;
         end;
  END ;

/// Menu Pop général
Case NumModule of
   92 : ChargeMenuPop(hm8,FMenuG.Dispatch) ;
   else ChargeMenuPop(TTypeMenu(NumModule),FMenuG.Dispatch) ;
  END ;

END ;

procedure InitApplication ;
BEGIN
ProcZoomEdt:=RTZoomEdtEtat ;
//ProcCalcEdt:=CalcOLEEtat ;
//ProcGetVH:=GetVS1 ;
//ProcGetDate:=GetDate ;
FMenuG.OnDispatch:=Dispatch ;
FMenuG.OnChargeMag:=ChargeMagHalley ;
FMenuG.OnMajAvant:=Nil ;
FMenuG.OnMajApres:=Nil ;
FMenuG.SetModules([92],[77]) ;
FMenuG.OnChangeModule:=AfterChangeModule ;
FMenuG.SetPreferences(['Pièces'],False) ;
V_PGI.DispatchTT:=DispatchTT ;
END ;

initialization
Apalatys:='CEGID' ;
NomHalley:='CRCS5' ;
TitreHalley:='RELATION CLIENTS S5' ;
HalSocIni:='CRCS5.ini' ;
Copyright:='© Copyright ' + Apalatys ;
V_PGI.OutLook:=TRUE ;
V_PGI.OfficeMsg:=TRUE ;
V_PGI.ToolsBarRight:=TRUE ;
ChargeXuelib ;
V_PGI.VersionDemo:=TRUE ;
V_PGI.MenuCourant:=0 ;
V_PGI.VersionReseau:=False ;
V_PGI.NumVersion:='0.90' ; V_PGI.NumBuild:='6' ;
V_PGI.DateVersion:=EncodeDate(2000,06,20) ;
V_PGI.SAV:=True ;
V_PGI.ImpMatrix := True ;
V_PGI.OKOuvert:=FALSE ;
V_PGI.Halley:=TRUE ;
V_PGI.NiveauAcces:=1 ;
V_PGI.MenuCourant:=0 ;
V_PGI.NumMenuPop:=27 ;
V_PGI.LaSerie:=S5 ;
V_PGI.Decla100:=TRUE;
V_PGI.RazForme:=FALSE;
V_PGI.PGIContexte:=[ctxGescom] ;
end.
