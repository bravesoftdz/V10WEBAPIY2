unit UtilSocAF;

interface

uses Forms,SysUtils,Controls,buttons,Classes,ComCtrls,Graphics,HCtrls,
     stdctrls,Grids, ExtCtrls, Windows, HEnt1,registry,Menus,TraducAffaire,
     WinProcs,shellAPI,HDebug,Hstatus,LicUtil,HMsgBox,UtilPGI,{$IFNDEF AGL570}PGIEnv,{$ENDIF}
     Messages,Hrichedt,HTB97,Dialogs,UTOB,LookUp, Entgc,DicoAF ,AffaireUtil
{$IFDEF EAGLCLIENT}
{$ELSE}
     ,MajTable,DB,DBTables,DBGrids,DBCtrls,HDB,UtilArticle
{$ENDIF}
        ;


Function  InterfaceSocGA ( CC : TControl ) : boolean ;
Function  ChargePageSocGA ( CC : TControl ) : boolean ;
Function  SauvePageSocGA ( CC : TControl ) : boolean ;
Function  SauvePageSocSansverif ( CC : TControl ) : boolean ;
Procedure MarquerPublifi ( Flag : boolean ) ;
Function ExisteEtabliss ( Eta : String ) : boolean ;
Function ExisteArticle ( Art : String ) : boolean ;
Function ExisteProfilGener ( Profil : String ) : boolean ;
Function ExisteUnite ( Uni : String ) : boolean ;
Function ExisteAffaires : boolean ;
Function ExisteActivite : boolean ;
Function ExisteAfCumul : boolean ;


implementation

Uses Ent1, SaisUtil,  ParamSoc, SaisComm, Spin
{$IFDEF GIGI}
     ,AssistGI,HCompte
  {$IFNDEF EAGLCLIENT}
     ,galoutil//,AssistPL
  {$ENDIF}
{$ENDIF}


      ;

{========================= Fonctions utilitaires ======================}
Function GetLaTOB ( CC : TControl ) : TOB ;
BEGIN Result:=TFParamSoc(CC.Owner).LaTOB ; END ;

Function GetTOBSoc ( CC : TControl ) : TOB ;
BEGIN Result:=GetLaTOB(CC).Detail[0] ; END ;

Function GetTOBCpts ( CC : TControl ) : TOB ;
BEGIN Result:=GetLaTOB(CC).Detail[1] ; END ;

Function GetLaForm ( CC : TControl ) : TForm ;
BEGIN Result:=TForm(CC.OWner) ; END ;

Procedure BourreLeCpt ( Cpt : THCritMaskEdit ) ;
Var LeFb : TFichierBase ;
BEGIN
if Cpt.Text='' then Exit ;
LeFb:=CaseFicDataType(Cpt.DataType) ;
if Length(Cpt.Text)<VH^.Cpta[Lefb].Lg then Cpt.Text:=BourreLaDonc(Cpt.Text,Lefb) ;
END ;

Function isFourcCpte ( Nam : String ) : boolean ;
Var Sub : String ;
BEGIN
Result:=True ; Sub:=Copy(Nam,1,9) ;
if Sub='SO_BILDEB' then Exit ; if Sub='SO_BILFIN' then Exit ;
if Sub='SO_CHADEB' then Exit ; if Sub='SO_CHAFIN' then Exit ;
if Sub='SO_PRODEB' then Exit ; if Sub='SO_PROFIN' then Exit ;
if Sub='SO_EXTDEB' then Exit ; if Sub='SO_EXTFIN' then Exit ;
Result:=False ;
END ;

Function ExisteNam ( Nam : String ; FF : TForm ) : Boolean ;
BEGIN Result:=(FF.FindComponent(Nam)<>Nil) ; END ;

Function GetFromNam ( Nam : String ; FF : TForm ) : TControl ;
BEGIN
Result:=TControl(FF.FindComponent(Nam)) ;
if Result=Nil then ShowMessage('Pas trouvé '+Nam) ;
END ;

Procedure SetInvi ( Nam : String ; FF : TForm ) ;
Var CC : TControl ;
BEGIN
CC:=GetFromNam(Nam,FF) ; if CC<>Nil then CC.Visible:=False ;
END ;

Procedure SetVisi ( Nam : String ; FF : TForm ) ;
Var CC : TControl ;
BEGIN
CC:=GetFromNam(Nam,FF) ; if CC<>Nil then CC.Visible:=True ;
END ;

Procedure SetEna ( Nam : String ; FF : TForm ; Ena : boolean ) ;
Var CC : TControl ;
BEGIN
CC:=GetFromNam(Nam,FF) ; if CC<>Nil then CC.Enabled:=Ena ;
END ;

Function RechExisteH ( CC : TControl ) : boolean ;
BEGIN Result:=LookupvalueExist(CC) ; END ;

Function DesEnregs ( TableName : String ) : boolean ;
Var Q : TQuery ;
BEGIN
// PL le 09/01/02 pour optimiser le nombre d'enregistrements lus
Q:=OpenSQL('SELECT * FROM '+TableName,True,1) ;
//Q:=OpenSQL('SELECT * FROM '+TableName,True) ;
Result:=(Not Q.EOF) ;
Ferme(Q) ;
END ;

{============================= Exit des zones =========================}
Procedure ControleBudX ( CC : TControl ) ;
Var FF : TForm ;
    JalBud : THValComboBox ;
    CBud : TCheckBox ;
BEGIN
FF:=GetLaForm(CC) ;
CBud:=TCheckBox(CC) ;
JalBud:=THValComboBox(GetFromNam('SO_JALCTRLBUD',FF)) ;
if CBud.Checked then JalBud.Enabled:=True
                else BEGIN JalBud.Enabled:=False ; JalBud.Value:='' ; END ;
END ;

Procedure DateDebutEuroX ( CC : TControl ) ;
Var FF : TForm ;
    DateDebutEuro : THCritmaskEdit ;
    TauxEuro      : THNumEdit ;
    TTauxEuro : THLabel ;
BEGIN
FF:=GetLaForm(CC) ;
DateDebutEuro:=THCritMaskEdit(CC) ;
TauxEuro:=THNumEdit(GetFromNam('SO_TAUXEURO',FF)) ;
TTauxEuro:=THLabel(GetFromNam('LSO_TAUXEURO',FF)) ;
if TAUXEURO<>Nil then TAUXEURO.Enabled:=(V_PGI.DateEntree<StrToDate(DATEDEBUTEURO.Text)) ;
if TTAUXEURO<>Nil then TTAUXEURO.Enabled:=TAUXEURO.Enabled ;
END ;

Procedure CptBourreX ( CC : TControl ) ;
Var CptBil : THCritmaskEdit ;
BEGIN
CptBil:=THCritMaskEdit(CC) ;
if CptBil<>Nil then BourreLeCpt(CptBil) ;
END ;
   
Procedure PaysX ( CC : TControl ) ;
{$IFDEF AFAIRE}
Var FF : TForm ;
    LePays,LaDiv : THCritMaskEdit ;
{$ENDIF}
BEGIN
{$IFDEF AFAIRE}
FF:=GetLaForm(CC) ;
LePays:=THCritMaskEdit(CC) ;
LaDiv:=THCritMaskEdit(GetFromNam('SO_DIVTERRIT',FF)) ;
// A faire simon PaysRegion(LaPays,LaDiv,False) ;
{$ENDIF}
END ;

function CreerTobModArt:tob;
var Ob_Mod:tob;
begin
OB_MOD := TOB.Create('ARTICLE',Nil,-1);
OB_MOD.PutValue('GA_FAMILLETAXE1','NOR');
OB_MOD.PutValue('GA_DATESUPPRESSION',idate2099);
OB_MOD.PutValue('GA_STATUTART','UNI');
OB_MOD.PutValue('GA_CALCPRIXHT','AUC');
OB_MOD.PutValue('GA_CALCPRIXTTC','AUC');
OB_MOD.PutValue('GA_ACTIVITEREPRISE','F');
OB_MOD.PutValue('GA_REMISEPIED','X');
OB_MOD.PutValue('GA_ESCOMPTABLE','X');
OB_MOD.PutValue('GA_COMMISSIONNABLE','X');
OB_MOD.PutValue('GA_ACTIVITEEFFECT','-');
OB_MOD.PutValue('GA_REMISELIGNE','X');
OB_MOD.PutValue('GA_INVISIBLEWEB','X');
OB_MOD.PutValue('GA_CREERPAR','GEN');
OB_MOD.PutValue('GA_PRIXPOURQTE',1);
OB_MOD.PutValue('GA_CREATEUR',V_PGI.user);
OB_MOD.PutValue ('GA_SOCIETE', GetParamSOc('SO_SOCIETE'));
result:=OB_MOD;
end;

function CreerArticlesAppreciation:boolean;
var  OB , OB_DETAIL,OB_MOD : TOB;
begin
  //mcd 12/02/02 reprendre fct, optimisée...
  // Création de la tob mère
  OB := TOB.Create('un article',nil,-1);
  OB_MOD:=CreerTobMOdArt;

  try
    if Not ExisteSQL('SELECT GA_CODEARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="PRESAPP"' ) then
        begin
        OB_DETAIL := TOB.Create('ARTICLE',OB,-1);
        OB_detail.dupliquer(OB_MOD,false,true);
        OB_DETAIL.PutValue('GA_ARTICLE',CodeArticleUnique2('PRESAPP',''));
        OB_DETAIL.PutValue('GA_CODEARTICLE','PRESAPP');
        OB_DETAIL.PutValue('GA_LIBELLE','Prestation Appréciation');
        OB_DETAIL.PutValue('GA_TYPEARTICLE','PRE');
        OB_DETAIL.PutValue('GA_QUALIFUNITEVTE',GetParamSoc('SO_AFMESUREACTIVITE'));
        OB_DETAIL.PutValue('GA_QUALIFUNITEACT',GetParamSoc('SO_AFMESUREACTIVITE'));
        end;
    if Not ExisteSQL('SELECT GA_CODEARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="FRAPPREC"' ) then
        begin
        OB_DETAIL := TOB.Create('ARTICLE',OB,-1);
        OB_detail.dupliquer(OB_MOD,false,true);
        OB_DETAIL.PutValue('GA_ARTICLE',CodeArticleUnique2('FRAPPREC',''));
        OB_DETAIL.PutValue('GA_CODEARTICLE','FRAPPREC');
        OB_DETAIL.PutValue('GA_LIBELLE','Frais Appréciation');
        OB_DETAIL.PutValue('GA_TYPEARTICLE','FRA');
        end;
    if Not ExisteSQL('SELECT GA_CODEARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="FOURAPP"' ) then
        begin
        OB_DETAIL := TOB.Create('ARTICLE',OB,-1);
        OB_detail.dupliquer(OB_MOD,false,true);
        OB_DETAIL.PutValue('GA_ARTICLE',CodeArticleUnique2('FOURAPP',''));
        OB_DETAIL.PutValue('GA_CODEARTICLE','FOURAPP');
        OB_DETAIL.PutValue('GA_LIBELLE','Fourniture Appréciation');
        OB_DETAIL.PutValue('GA_TYPEARTICLE','MAR');
        end;

    Result:=true;
  finally
    if (OB.Detail.count<>0) then  OB.InsertDB(nil);
       // mcd 25/11/02OB.InsertDBtable(nil);
    OB.Free;
    Ob_Mod.free;  //mcd 25/11/02
  end;
end;


function CreerArticlesBoni:boolean;
var  OB , OB_DETAIL,OB_Mod : TOB;
begin
  // création des 2 articles par défaut pour les boni/mali
  // attention ne sont créer que si table parboni est vide !!!!
  // Création de la tob mère
  OB := TOB.Create('un article',nil,-1);
  OB_MOD:=CreerTobMOdArt;

  try
    if Not ExisteSQL('SELECT GA_CODEARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="PRBONI"' ) then
        begin
        OB_DETAIL := TOB.Create('ARTICLE',OB,-1);
        OB_detail.dupliquer(OB_MOD,false,true);
        OB_DETAIL.PutValue('GA_ARTICLE',CodeArticleUnique2('PRBONI',''));
        OB_DETAIL.PutValue('GA_CODEARTICLE','PRBONI');
        OB_DETAIL.PutValue('GA_LIBELLE','Prestation de boni');
        OB_DETAIL.PutValue('GA_TYPEARTICLE','PRE');
        OB_DETAIL.PutValue('GA_QUALIFUNITEVTE',GetParamSoc('SO_AFMESUREACTIVITE'));
        OB_DETAIL.PutValue('GA_QUALIFUNITEACT',GetParamSoc('SO_AFMESUREACTIVITE'));
        end;
    if Not ExisteSQL('SELECT GA_CODEARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="PRMALI"' ) then
        begin
        OB_DETAIL := TOB.Create('ARTICLE',OB,-1);
        OB_detail.dupliquer(OB_MOD,false,true);
        OB_DETAIL.PutValue('GA_ARTICLE',CodeArticleUnique2('PRMALI',''));
        OB_DETAIL.PutValue('GA_CODEARTICLE','PRMALI');
        OB_DETAIL.PutValue('GA_LIBELLE','Prestation de mali');
        OB_DETAIL.PutValue('GA_TYPEARTICLE','PRE');
        OB_DETAIL.PutValue('GA_QUALIFUNITEVTE',GetParamSoc('SO_AFMESUREACTIVITE'));
        OB_DETAIL.PutValue('GA_QUALIFUNITEACT',GetParamSoc('SO_AFMESUREACTIVITE'));
        end;
    Result:=true;
  finally
    if (OB.Detail.count<>0) then  OB.InsertDB(nil);
    OB.Free;
    Ob_Mod.free;
  end;
end;

function CreerRessBoni:boolean;
var  OB , OB_DETAIL : TOB;
begin
  //Creation de la ressource par defaut pour boni/mali
  // Création de la tob mère
  OB := TOB.Create('une ressource',nil,-1);
  try
    if Not ExisteSQL('SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_RESSOURCE="ASSBONI"' ) then
    begin
      OB_DETAIL := TOB.Create('RESSOURCE',OB,-1);
      OB_DETAIL.PutValue('ARS_RESSOURCE','ASSBONI');
      OB_DETAIL.PutValue('ARS_TYPERESSOURCE','SAL');
      OB_DETAIL.PutValue('ARS_LIBELLE',TraduitGa('Ressource affectation Boni/mali'));
      OB_DETAIL.PutValue ('ARS_UNITETEMPS', VH_GC.AFMesureActivite  );
      OB_DETAIL.PutValue ('ARS_TAUXFRAISGEN1', VH_GC.AFFraisGen1  );
      OB_DETAIL.PutValue ('ARS_TAUXFRAISGEN2', VH_GC.AFFraisGen2  );
      OB_DETAIL.PutValue ('ARS_TAUXCHARGEPAT', VH_GC.AFChargesPat );
      OB_DETAIL.PutValue ('ARS_COEFMETIER'   , VH_GC.AFCoefMetier );
      OB_DETAIL.PutValue ('ARS_ARTICLE'      , VH_GC.AFPrestationRes );
      OB_DETAIL.PutValue ('ARS_FERME', '-');
      OB_DETAIL.PutValue  ('ARS_PAYS', GetParamSoc('SO_GcTiersPays'));
      OB_DETAIL.PutValue  ('ARS_COEFPRPV',1.0);
      If (VH_GC.AFResCalculPR) Then OB_DETAIL.PutValue  ('ARS_CALCULPR','X') Else OB_DETAIL.PutValue ('ARS_CALCULPR','-');
      OB_DETAIL.PutValue  ('ARS_DEBUTDISPO',idate1900);
      OB_DETAIL.PutValue  ('ARS_FINDISPO',idate2099);
      OB_DETAIL.PutValue  ('ARS_CREATEUR', V_PGI.User);
      OB_DETAIL.PutValue ('ARS_UTILISATEUR', V_PGI.User);
      OB_DETAIL.PutValue ('ARS_SOCIETE', GetParamSOc('SO_SOCIETE'));
      OB_DETAIL.PutValue  ('ARS_CREERPAR', 'GEN');
    end;
    Result:=true;
  finally
    if (OB.Detail.count<>0) then OB.InsertDB(nil);
    OB.Free;
  end;
end;


Procedure AppreciationX( CC : TControl ) ;
Var FF : TForm ;
    CMEAppPres, CMEAppFrais, CMEAppFour : THCritMaskEdit ;
    LblAppPres, LblAppFrais, LblAppFour : THLabel;
    CbAppreciation, CbBoniMali, CbFactures: TCheckBox;
BEGIN
//mcd 12/02/02 reprendre toute la fct, tout revu
FF:=GetLaForm(CC) ;
CbAppreciation:=TCheckBox(CC);

CMEAppPres:=THCritMaskEdit(GetFromNam('SO_AFAPPPRES',FF)) ;
CMEAppFrais:=THCritMaskEdit(GetFromNam('SO_AFAPPFRAIS',FF)) ;
CMEAppFour:=THCritMaskEdit(GetFromNam('SO_AFAPPFOUR',FF)) ;
LblAppPres:=THLabel(GetFromNam('LSO_AFAPPPRES',FF)) ;
LblAppFrais:=THLabel(GetFromNam('LSO_AFAPPFRAIS',FF)) ;
LblAppFour:=THLabel(GetFromNam('LSO_AFAPPFOUR',FF)) ;
CbBoniMali:=TCheckBox(GetFromNam('SO_AFAPPAVECBM',FF)) ;
CbFactures:=TCheckBox(GetFromNam('SO_AFAPPPOINT',FF)) ;

{// PL le 21/01/02 : pour blocage appréciation si monnaie de tenue du dossier est le franc
// Si la monnaie de tenue est le franc
if (GetParamSoc('SO_DEVISEPRINC')='FRF') then
    begin
    CbAppreciation.Checked:=false;
    HShowMessage('24;Société;Utilisation des fonctions d''appréciation impossible. Veuillez effectuer la Bascule EURO du dossier;W;O;O;O;','','') ;
    end;
// Fin PL le 21/01/02 } // mcd 03/04/2002 plus de blocage

if CbAppreciation.Checked then
    begin  // si coché, on renseigne les valeurs par défaut
    if (CMeAppPres.text<>'') then exit; // si il y a déjà des valeurs, il ne faut pas changer les coches
                                        // qui ont pu être modifié par l'utilisateur.
    CMEAppPres.Enabled := true;
    CMEAppFrais.Enabled := true;
    CMEAppFour.Enabled := true;
    LblAppPres.Enabled := true;
    LblAppFrais.Enabled := true;
    LblAppFour.Enabled := true;
    CbBoniMali.Enabled := true;
    CbBoniMali.checked := true;
    CreerArticlesAppreciation;   //création des art par défaut si inexistant
    if (CMeAppPres.text='') then CMEAppPres.Text := 'PRESAPP';
    if (CMeAppFrais.text='') then CMEAppFrais.Text := 'FRAPPREC';
    if (CMeAppFour.text='') then CMEAppFour.Text := 'FOURAPP';
    end
else
    begin   // si case non cochée, on efface tout ce qui est dans les zones suivantes
    CMEAppPres.Enabled := false;
    CMEAppFrais.Enabled := false;
    CMEAppFour.Enabled := false;
    LblAppPres.Enabled := false;
    LblAppFrais.Enabled := false;
    LblAppFour.Enabled := false;
    CbBoniMali.Enabled := false;
    CbFactures.Enabled := false;
    CMEAppPres.Text := '';
    CMEAppFrais.Text := '';
    CMEAppFour.Text := '';
    CbBoniMali.checked := false;
    CbFactures.checked := false;
    end;

END ;

Procedure CutOffX( CC : TControl ) ;
Var FF : TForm ;
    CbCutOff: TCheckBox;
    ModeEclat : THValComboBox ;
BEGIN
FF:=GetLaForm(CC) ;
CbCutOff:=TCheckBox(CC);
ModeEclat:=THValComboBox(GetFromNam('SO_AFCUTMODEECLAT',FF)) ;

if not(CbCutOff.Checked) then
    begin   // si case non cochée, on efface tout ce qui est dans les zones suivantes
    ModeEclat.value:='SAN';
    end;

END ;




Procedure DevPrincX ( CC : TControl ) ;
Var FF : TForm ;
    Q  : TQuery ;
    CH : THValComboBox ;
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_DEVISEPRINC',FF) then Exit ;
// Traitement
CH:=THValComboBox(GetFromNam('SO_DEVISEPRINC',FF)) ;
Q:=OpenSQL('SELECT D_DECIMALE FROM DEVISE WHERE D_DEVISE="'+CH.Value+'"',True) ;
If Not Q.EOF then TSpinEdit(GetFromNam('SO_DECVALEUR',FF)).Value:=Q.Fields[0].AsInteger ;
Ferme(Q) ;
END ;

Function InterfaceSocGA ( CC : TControl ) : boolean ;
Var Nam : String ;
FF:TForm;
BEGIN
Result:=True ; Nam:=CC.Name ;
if Nam='SO_DATEDEBUTEURO'           then DateDebutEuroX(CC) else
if isFourcCpte(Nam)                 then CptBourreX(CC) else
if Copy(Nam,1,9)='SO_DEFCOL'        then CptBourreX(CC) else
if Nam='SO_DEVISEPRINC'             then DevPrincX(CC) else
if Nam='SO_CONTROLEBUD'             then ControleBudX(CC) else
if Nam='SO_PAYS'                    then PaysX(CC) else
if Nam='SO_AFGESTIONAPPRECIATION'   then AppreciationX(CC) else
if Nam='SO_AFAPPCUTOFF'             then CutOffX(CC) else
if Nam='SO_TYPEDATEFINACTIVITE'     then
    begin
    FF:=TForm(CC.Owner) ;
    if (thvalcombobox(cc).value='DAF') then
        SetEna('SO_AFDATEFINACT', FF, True)
    else
        begin
        SetParamSoc('SO_AFDATEFINACT', DateToStr(idate2099));
        SetEna('SO_AFDATEFINACT', FF, False);
        end;

    end

else
;
END ;

{============================= Sur Chargement =========================}
Function ExisteEtabliss ( Eta : String ) : boolean ;
BEGIN
Result:=False ; if Eta='' then Exit ;
Result := ExisteSQL('SELECT ET_ETABLISSEMENT FROM ETABLISS WHERE ET_ETABLISSEMENT="'+Eta+'"' ) ;
END ;

Function ExisteUnite ( Uni : String ) : boolean ;
BEGIN
Result:=False ; if Uni='' then Exit ;
Result := ExisteSQL('SELECT GME_MESURE FROM MEA WHERE GME_MESURE="'+Uni+'"') ;
END ;

Function ExisteProfilGener ( Profil : String ) : boolean ;
BEGIN
Result:=False ; if Profil='' then Exit ;
Result := ExisteSQL('SELECT APG_PROFILGENER FROM PROFILGENER WHERE APG_PROFILGENER="'+Profil+'"' ) ;
END ;

Function ExisteArticle ( Art : String ) : boolean ;
BEGIN
Result:=False ; if Art='' then Exit ;
Result := ExisteSQL('SELECT GA_CODEARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="'+Art+'"' ) ;
END ;

Function ExisteAffaires : boolean ;
BEGIN
  Result := ExisteSQL('SELECT AFF_AFFAIRE FROM AFFAIRE' ) ;
END ;

Function ExisteActivite : boolean ;
BEGIN
Result := ExisteSQL('SELECT ACT_TYPEACTIVITE FROM ACTIVITE');
END ;

Function ExisteAfCumul : boolean ;
BEGIN
Result := ExisteSQL('SELECT ACU_TYPEAC FROM AFCUMUL WHERE ACU_TYPEAC="CVE"' ) ;
END ;

Function ExisteMvtsPayeurs ( Jal : String ) : boolean ;
//Var Q : TQuery ;
BEGIN
Result:=False ; if Jal='' then Exit ;
Result := ExisteSQL('SELECT E_JOURNAL FROM ECRITURE WHERE E_JOURNAL="'+Jal+'"') ;
END ;

Function ExisteMvtsCpte ( Cpte : String ) : boolean ;
//Var Q : TQuery ;
BEGIN
Result:=False ; if Cpte='' then Exit ;
Result := ExisteSQL('SELECT E_GENERAL FROM ECRITURE WHERE E_GENERAL="'+Cpte+'"') ;
END ;

Function ExisteMvts : boolean ;
Var Q : TQuery ;
    OkMvt : boolean ;
BEGIN
OkMvt:=False ;
// PL le 09/01/02 pour optimiser le nombre d'enregistrements lus en eagl
if Not OkMvt then BEGIN Q:=OpenSQL('SELECT E_JOURNAL FROM ECRITURE',True,1) ; if Not Q.EOF then OkMvt:=True ; Ferme(Q) ; END ;
if Not OkMvt then BEGIN Q:=OpenSQL('SELECT Y_JOURNAL FROM ANALYTIQ',True,1) ; if Not Q.EOF then OkMvt:=True ; Ferme(Q) ; END ;
if Not OkMvt then BEGIN Q:=OpenSQL('SELECT BE_BUDJAL FROM BUDECR',True,1)   ; if Not Q.EOF then OkMvt:=True ; Ferme(Q) ; END ;
if Not OkMvt then BEGIN Q:=OpenSQL('SELECT EG_TYPE   FROM ECRGUI',True,1)   ; if Not Q.EOF then OkMvt:=True ; Ferme(Q) ; END ;
//if Not OkMvt then BEGIN Q:=OpenSQL('SELECT E_JOURNAL FROM ECRITURE',True) ; if Not Q.EOF then OkMvt:=True ; Ferme(Q) ; END ;
//if Not OkMvt then BEGIN Q:=OpenSQL('SELECT Y_JOURNAL FROM ANALYTIQ',True) ; if Not Q.EOF then OkMvt:=True ; Ferme(Q) ; END ;
//if Not OkMvt then BEGIN Q:=OpenSQL('SELECT BE_BUDJAL FROM BUDECR',True)   ; if Not Q.EOF then OkMvt:=True ; Ferme(Q) ; END ;
//if Not OkMvt then BEGIN Q:=OpenSQL('SELECT EG_TYPE   FROM ECRGUI',True)   ; if Not Q.EOF then OkMvt:=True ; Ferme(Q) ; END ;
Result:=OkMvt ;
END ;

Procedure ShowParamSoc ( CC : TControl ) ;
Var LaTOB,TOBSoc,TOBCpt : TOB ;
    OkMvt : boolean ;
    i,j : integer ;
    NN  : String ;
BEGIN
LaTOB:=GetLaTOB(CC) ;
if LaTOB.Detail.Count<=0 then
   BEGIN
   TOBSoc:=TOB.Create('',LaTOB,-1) ;
   // CHamps Supp
   TOBSoc.AddChampSup('OKECR',False) ; TOBSoc.AddChampSup('OKGEN',False) ; TOBSoc.AddChampSup('OKAUX',False) ;
   TOBSoc.AddChampSup('OKJALATP',False) ; TOBSoc.AddChampSup('OKJALVTP',False) ;
   TOBSoc.AddChampSup('OKECCDEBIT',False) ; TOBSoc.AddChampSup('OKECCCREDIT',False) ;
   // Alimentations
   OkMvt:=ExisteMvts ; if OkMvt then TOBSoc.PutValue('OKECR','X') else TOBSoc.PutValue('OKECR','-') ;
//   OkMvt:=ExisteMvtsCpte(VH^.ECCEuroDebit)  ; if OkMvt then TOBSoc.PutValue('OKECCDEBIT','X') else TOBSoc.PutValue('OKECCDEBIT','-') ;
//   OkMvt:=ExisteMvtsCpte(VH^.ECCEuroCredit) ; if OkMvt then TOBSoc.PutValue('OKECCCREDIT','X') else TOBSoc.PutValue('OKECCCREDIT','-') ;
   OkMvt:=ExisteMvtsPayeurs(VH^.JalATP) ; if OkMvt then TOBSoc.PutValue('OKJALATP','X') else TOBSoc.PutValue('OKJALATP','-') ;
   OkMvt:=ExisteMvtsPayeurs(VH^.JalVTP) ; if OkMvt then TOBSoc.PutValue('OKJALVTP','X') else TOBSoc.PutValue('OKJALVTP','-') ;
   TOBSoc.PutValue('OKGEN',CheckToString(DesEnregs('GENERAUX'))) ;
   TOBSoc.PutValue('OKAUX',CheckToString(DesEnregs('TIERS'))) ;
    // mcd 28/01/03 TOBCpt:=TOB.Create('CPTS',LaTOB,-1) ;
   TOBCpt:=TOB.Create('Les Comptes',LaTOB,-1) ;
   for i:=1 to 4 do
       BEGIN
       Case i of
          1 : NN:='BIL' ; 2 : NN:='CHA' ; 3 : NN:='PRO' ; 4 : NN:='EXT' ;
          END ;
       for j:=1 to 5 do
           BEGIN
           TObCpt.AddChampSup('SO_'+NN+'DEB'+InttoStr(j),False) ;
           TObCpt.AddChampSup('SO_'+NN+'FIN'+InttoStr(j),False) ;
           END ;
       END ;
   END ;
END ;                    

Procedure BlocageDevise ( CC : TControl ) ;
Var FF : TForm ;
    Ena : boolean ;
BEGIN
FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_DEVISEPRINC',FF) then Exit ;
Ena:=(GetTOBSoc(CC).GetValue('OKECR')='-') ;
SetEna('SO_DEVISEPRINC',FF,Ena) ;
SetEna('SO_DECVALEUR',FF,False) ;
END ;

Procedure BlocageECC ( CC : TControl ) ;
Var FF : TForm ;
    Ena : boolean ;
BEGIN
FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_ECCEURODEBIT',FF) then Exit ;
Ena:=(GetTOBSoc(CC).GetValue('OKECCDEBIT')='-')  ; SetEna('SO_ECCEURODEBIT',FF,Ena) ;
Ena:=(GetTOBSoc(CC).GetValue('OKECCCREDIT')='-') ; SetEna('SO_ECCEUROCREDIT',FF,Ena) ;
END ;

Procedure InitDivers ( CC : TControl ) ;
Var FF : TForm ;
    JalBud,RefL,RefP : THValComboBox ;
    OldV : String ;
    HH   : THLabel ;
BEGIN
FF:=GetLaForm(CC) ;
if ExisteNam('SO_TAUXEURO',FF) then
   BEGIN
   DevPrincX(CC) ;
   SetEna('SO_TAUXEURO',FF,(V_PGI.DateEntree<V_PGI.DateDebutEuro)) ;
   SetEna('LSO_TAUXEURO',FF,(V_PGI.DateEntree<V_PGI.DateDebutEuro)) ;
   SetEna('SO_TENUEEURO',FF,False) ;
   if TCheckBox(GetFromNam('SO_TENUEEURO',FF)).Checked then SetEna('SO_TAUXEURO',FF,False) ;
   if VH^.TenueEuro then SetEna('SO_DATEDEBUTEURO',FF,False) ;
   END ;
if ExisteNam('SO_JALCTRLBUD',FF) then
   BEGIN
   JalBud:=THValComboBox(GetFromNam('SO_JALCTRLBUD',FF)) ;
   if JalBud.Vide then JalBud.Items[0]:='' ;
   TCheckBox(GetFromNam('SO_CONTROLEBUD',FF)).Checked:=(JalBud.Value<>'') ;
   if JalBud.Value='' then JalBud.Enabled:=False ;
   END ;
if ExisteNam('SO_CPEXOREF',FF) then if Not (ctxPCL in V_PGI.PGIContexte) then
   BEGIN
   SetInvi('SO_CPEXOREF',FF) ; SetInvi('LSO_CPEXOREF',FF) ;
   END ;
if ExisteNam('SO_LGCPTEGEN',FF) then
   BEGIN
   if GetTOBSoc(CC).GetValue('OKGEN')='X' then SetEna('SO_LGCPTEGEN',FF,False) ;
   if GetTOBSoc(CC).GetValue('OKAUX')='X' then SetEna('SO_LGCPTEAUX',FF,False) ;
   END ;
if ExisteNam('SO_JALATP',FF) then
   BEGIN
   if GetTOBSoc(CC).GetValue('OKJALATP')='X' then SetEna('SO_JALATP',FF,False) ;
   if GetTOBSoc(CC).GetValue('OKJALVTP')='X' then SetEna('SO_JALVTP',FF,False) ;
   END ;
if EstSerie(S5) then
   BEGIN
   if ExisteNam('SO_CPREFLETTRAGE',FF) then
      BEGIN
      RefL:=THValComboBox(GetFromNam('SO_CPREFLETTRAGE',FF)) ;
      OldV:=RefL.Value ;
      RefL.DataType:='CPPARAMREFERENCESS5' ; RefL.Reload ;
      RefL.Value:=OldV ;
      END ;
   if ExisteNam('SO_CPREFPOINTAGE',FF) then
      BEGIN
      RefP:=THValComboBox(GetFromNam('SO_CPREFPOINTAGE',FF)) ;
      OldV:=RefP.Value ;
      RefP.DataType:='CPPARAMREFERENCESS5' ; RefP.Reload ;
      RefP.Value:=OldV ;
      END ;
   if ExisteNam('SO_CPLIBREANAOBLI',FF) then SetInvi('SO_CPLIBREANAOBLI',FF) ;
   END ;
if ExisteNam('LSO_CPRDNOMCLIENT',FF) then if Not V_PGI.Controleur then
   BEGIN
   HH:=THLabel(GetFromNam('LSO_CPRDNOMCLIENT',FF)) ; HH.Caption:='Nom du Cabinet' ;
   HH:=THLabel(GetFromNam('LSO_CPRDEMAILCLIENT',FF)) ; HH.Caption:='E-Mail du Cabinet' ;
   END ;
END ;

Procedure GetCompteFourchette ( CC : TControl ; NN : String ) ;
Var i    : Integer ;
    Nam  : String ;
    CH   : THCritMaskEdit ;
    FF   : TForm ;
BEGIN
FF:=TForm(CC.Owner) ;
for i:=1 to 5 do
    BEGIN
    Nam:='SO_'+NN+'DEB'+IntToStr(i) ; CH:=THCritMaskEdit(FF.FindComponent(Nam)) ;
    if CH<>Nil then GetTOBCpts(CC).PutValue(Nam,CH.Text) ;
    Nam:='SO_'+NN+'FIN'+IntToStr(i) ; CH:=THCritMaskEdit(FF.FindComponent(Nam)) ;
    if CH<>Nil then GetTOBCpts(CC).PutValue(Nam,CH.Text) ;
    END ;
END ;

Procedure ConfigSaisieParamsDefaut ( CC : TControl ) ;
Var FF   : TForm ;
    QQ : TQuery;
BEGIN
// test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_GCTOUTEURO',FF) then Exit ;

if VH^.TenueEuro then begin
                 // en GA-GI on ne peut changer l'option que si toutes les affaire sont passée en EURO .. sinon pb.
// Modif PL le 09/01/02 pour optimiser le nombre d'enregistrements lus en eagl
   QQ := OpenSQL('SELECT  GP_NUMEro From PIECE WHERE (gp_naturepieceg="'
         +GetPAramSoc('SO_AFNATAFFAIRE')+'" or gp_naturepieceg="'
         +GetPAramSoc('SO_AFNATPROPOSITION')+'") and GP_SAISIECONTRE="X"',TRUE,1);
//   QQ := OpenSQL('SELECT  GP_NUMEro From PIECE WHERE (gp_naturepieceg="'
//         +GetPAramSoc('SO_AFNATAFFAIRE')+'" or gp_naturepieceg="'
//         +GetPAramSoc('SO_AFNATPROPOSITION')+'") and GP_SAISIECONTRE="X"',TRUE);
   If Not QQ.EOF then SetEna('SO_GCTOUTEURO', FF, False)
    else SetEna('SO_GCTOUTEURO', FF, True);
   Ferme(QQ);
   end
else    SetEna('SO_GCTOUTEURO', FF, False); // base en francs, on interdit de changer l'option

TraduitForm(FF);
END;

Procedure ConfigSaisieDecen( CC : TControl ) ;
Var FF   : TForm ;
BEGIN
// test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFEACTPERIODIC',FF) then Exit ;
If not(VH_GC.GASaisieDecSeria) then begin
    SetEna('SO_AFEACTPERIODIC', FF, False);
    SetEna('SO_AFEACTPATH', FF, False);
    SetEna('LSO_AFEACTPERIODIC', FF, False);
    SetEna('LSCO_GRPSAISIEDEC', FF, False);
    SetEna('LSO_AFEACTPATH', FF, False);
    SetEna('SO_AFEACTNEWAFF', FF, False);
    SetEna('SO_AFEACTUNPOURTOUS', FF, False);
    end;
// Option pour intégrer des lignes d'eactivite avec suppression d'activite au prealable, non accessible en Scot
// mais quand même visible avec le mot de passe du jour
// PL le 01/08/02
if (ctxScot in V_PGI.PGIContexte) then
    begin
    SetParamsoc('SO_AFEACTIMPSELEC',False);
    if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
        SetEna('SO_AFEACTIMPSELEC', FF, False)
    else
        SetInvi('SO_AFEACTIMPSELEC',FF) ;
    end;
end;

Procedure ConfigSaisieBudget( CC : TControl ) ;
Var FF   : TForm ;
BEGIN
// test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFALIMECHE',FF) then Exit ;
SetInvi('SO_AFBUDZERO',FF) ;

end;


//mcd 12/02/02 nouvelle fct
Procedure ConfigSaisieApprec ( CC : TControl ) ;
Var FF   : TForm ;
    CbAppreciation : TCheckBox;
BEGIN
// test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFGESTIONAPPRECIATION',FF) then Exit ;
CbAppreciation:=TCheckBox(GetFromNam('SO_AFGESTIONAPPRECIATION',FF));

SetEna('SO_AFAPPPOINT', FF, false); //pas gérer à ce jour
SetParamsoc('SO_AFAPPPOINT',False);
//SetParamsoc('SO_AFAPPCUTOFF',False);
If CbAppreciation.checked then begin
    SetEna('SO_AFAPPFRAIS', FF, True);
    SetEna('SO_AFAPPPRES', FF, True);
    SetEna('SO_AFAPPFOUR', FF, True);
    SetEna('SO_AFAPPAVECBM', FF, True);
   end
else begin
    SetEna('LSO_AFAPPFRAIS', FF, false);
    SetEna('LSO_AFAPPPRES', FF, false);
    SetEna('LSO_AFAPPFOUR', FF, false);
    SetEna('SO_AFAPPFRAIS', FF, false);
    SetEna('SO_AFAPPPRES', FF, false);
    SetEna('SO_AFAPPFOUR', FF, false);
    SetEna('SO_AFAPPAVECBM', FF, false);
    end;
TraduitForm(FF);
END;


Procedure ConfigSaisieFactEclat ( CC : TControl ) ;
Var FF   : TForm ;
BEGIN
// test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFFACTPARRES',FF) then Exit ;

SetEna('SO_AFFACTPARRES', FF, false);   // uniquement au menu. des traitement sont à faire sur le chgmt
SetEna('LSO_AFFACTPARRES', FF, false);
If GetParamsoc('SO_AFFACTPARRES')<>'SAN' then begin
    SetEna('SO_AFFACTPRESDEFAUT', FF, True);
    SetEna('SO_AFFACTFRAISDEFAUT', FF, True);
    SetEna('SO_AFFACTFOURDEFAUT', FF, True);
    SetEna('SO_AFFACTRESSDEFAUT', FF, True);
    SetEna('SO_AFFACTMODEPEC', FF, True);
    SetEna('SO_AFFEDATFINACT', FF, True);
   end
else begin
    SetEna('LSO_AFFACTPRESDEFAUT', FF, false);
    SetEna('LSO_AFFACTFRAISDEFAUT', FF, false);
    SetEna('LSO_AFFACTFOURDEFAUT', FF, false);
    SetEna('SO_AFFACTPRESDEFAUT', FF, false);
    SetEna('SO_AFFACTFRAISDEFAUT', FF, false);
    SetEna('SO_AFFACTFOURDEFAUT', FF, false);
    SetEna('SO_AFFACTRESSDEFAUT', FF, false);
    SetEna('SO_AFFEDATFINACT', FF, False);
    SetEna('LSO_AFFEDATFINACT', FF, False);
    SetEna('LSO_AFFACTRESSDEFAUT', FF, false);
    SetEna('LSO_AFFACTMODEPEC', FF, false);
    SetEna('SO_AFFACTMODEPEC', FF, false);
    end;
TraduitForm(FF);
END;

Procedure ConfigSaisiePreferences ( CC : TControl ) ;
Var FF   : TForm ;
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFACOMPTE',FF) then Exit ;


SetEna('SO_AFANNUAIREIMPETIERS', FF, false);
// Si on est dans la base 00, on force lien DP à VRAI, à Faux sinon
if (ctxScot in V_PGI.PGIContexte) then
    begin
    if V_PGI.RunFromLanceur then
        if (V_PGI_Env <> nil) then
            begin
            If (V_PGI_Env.InBaseCommune) then
                begin
                SetParamsoc('SO_AFLIENDP',true);
                SetEna('SO_AFANNUAIREIMPETIERS', FF, true)
                end
            else
                begin
                // si modif voir aussi MDISPGA dispatch 100
                SetParamsoc('SO_AFLIENDP',FALSE);
                end
            end
        else
            SetParamsoc('SO_AFLIENDP',FALSE);

    if V_PGI.SAV then
        SetEna('SO_AFLIENDP', FF, true)
    else
        SetEna('SO_AFLIENDP', FF, False);

    end
else
    begin
    SetParamsoc('SO_AFLIENDP',FALSE);
    SetInvi('SO_AFLIENDP', FF);
    SetParamsoc('SO_AFGESTIONTARIF',TRUE);
    SetInvi('SO_AFGESTIONTARIF', FF);
    SetParamsoc('SO_AFANNUAIREIMPETIERS',FALSE);
    SetInvi('SO_AFANNUAIREIMPETIERS', FF);
    {$IFDEF CCS3}
    SetInvi('SO_AFTYPECONF', FF);
    SetInvi('LSO_AFTYPECONF', FF);
    SetInvi('LSO_AFTYPECONF', FF);
    SetInvi('SO_AFSAISAFFINTERDIT', FF);
    {$ENDIF}
    end;
//if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
//        SetEna('SO_AFFACTPARRES', FF, True)
//      else SetEna('SO_AFFACTPARRES', FF, False);

// On ne doit pas pouvoir modifier l'option de gestion de la première semaine si de l'activite a deja ete saisie
if ExisteActivite then
  SetEna('SO_PREMIERESEMAINE', FF, False)
else
  SetEna('SO_PREMIERESEMAINE', FF, true);


TraduitForm(FF);
END;

Procedure ConfigSaisieDates ( CC : TControl ) ;
Var FF   : TForm ;
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ;
if Not ExisteNam('SO_AFDATEDEBCAB',FF) then Exit ;
if not(ctxScot  in V_PGI.PGIContexte) then
   begin
    SetInvi('SO_AFDATEDEBCAB', FF);
    SetInvi('SO_AFDATEFINCAB', FF);
    SetInvi('LSO_AFDATEDEBCAB', FF);
    SetInvi('LSO_AFDATEFINCAB', FF);
   end ;

{$IFDEF GIGI}
 If V_PGI.RunFromLanceur  then
    if Not GetFlagAppli ('CGIS5.EXE') then    // PL le 17/04/03 : chgt IsAppliActive de AssistPL en GetFlagAppli de galoutil
      begin
      //  PL le 21/11/02 : Traitement particulier à l'assistant de création GI au premier acces
      SetParamSoc('SO_PREMIERESEMAINE', 0);
      if (V_PGI<>nil) then V_PGI.Semaine53 := 0;
      SetVisi('LSO_PREMIERESEMAINE_', FF);
      SetVisi('SO_PREMIERESEMAINE_', FF);
      SetEna('SO_PREMIERESEMAINE_', FF, true);
      end;
{$ENDIF}


TraduitForm(FF);
END;

Procedure ConfigSaisieCutOff ( CC : TControl ) ;
Var FF   : TForm ;
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ;
if Not ExisteNam('SO_AFAPPCUTOFF',FF) then Exit ;

SetEna('SO_AFAPPCUTOFF', FF, True);
SetEna('SO_AFCUTMODEECLAT', FF, True);
SetEna('LSO_AFCUTMODEECLAT', FF, True);
SetEna('SO_AFALIMCUTOFF', FF, True);
SetEna('LSO_DATECUTOFF', FF, True);
SetEna('SO_DATECUTOFF', FF, True);
SetEna('LSO_DATECUTOFFFOUR', FF, True);
SetEna('SO_DATECUTOFFFOUR', FF, True);
   // accès aux params cut off uniquement avec mot de passe du jour ou si rien dans la table afcumul
if ExisteAfCumul then
  begin
//mcd 27/05/03  if  (V_PGI.PassWord <> CryptageSt(DayPass(Date))) then
//    begin
    SetEna('SO_AFCUTMODEECLAT', FF, false);
    SetEna('LSO_AFCUTMODEECLAT', FF, false);
// PL le 05/12/02 : on doit pouvoir alimenter ou pas les cut off    
//    SetEna('SO_AFALIMCUTOFF', FF, false);
    SetEna('LSO_DATECUTOFF', FF, false);
    SetEna('SO_DATECUTOFF', FF, false);
    SetEna('LSO_DATECUTOFFFOUR', FF, false);
    SetEna('SO_DATECUTOFFFOUR', FF, false);
    SetEna('LSO_AFFORMULCUTOFF', FF, false);
//mcd 27/05/03    end;
  end
else
  if (ctxScot  in V_PGI.PGIContexte) then
    begin
    // Dans Scot, la formule est figée par défaut
    SetParamsoc('SO_AFFORMULCUTOFF','[M0]-[M2]');
    end;


// Pour l'instant les champs lies aux cutoff fournisseurs ne sont pas visibles
SetInvi('SO_DATECUTOFFFOUR', FF);
SetInvi('SO_AFFORMULCOFOUR', FF);
SetInvi('LSO_DATECUTOFFFOUR', FF);
SetInvi('LSO_AFFORMULCOFOUR', FF);

// Dans tous les cas la formule n'est modifiable que par le menu 74752 de constitution de la formule du cut off
SetEna('SO_AFFORMULCUTOFF', FF, False);


TraduitForm(FF);
END;

Procedure ConfigSaisieActivite ( CC : TControl ) ;
Var FF   : TForm ;
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFDATEDEBUTACT',FF) then Exit ;

if (ctxScot in V_PGI.PGIContexte) then
  SetInvi('So_AFREMISEMONTANT', FF); // mcd 18/06/03 // PL le 02/10/03 : pour généraliser à toute la GA

if (GetParamSoc('SO_TYPEDATEFINACTIVITE')='DAF') then
    SetEna('SO_AFDATEFINACT', FF, True)
else
    SetEna('SO_AFDATEFINACT', FF, False);

if (ExisteSQL('SELECT ACT_TYPEACTIVITE FROM ACTIVITE')) then SetEna('SO_AFMESUREACTIVITE', FF,False); // mcd 17/10/02 changé par ligne menu

if (ctxScot  in V_PGI.PGIContexte) then
    // Dans Scot
   begin
    if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
         SetEna('SO_AFDATEDEBUTACT', FF, True)
    else SetEna('SO_AFDATEDEBUTACT', FF, False);
   end;
SetInvi('SO_AFUTILTARIFACT', FF); //pas géré pour l'instant
SetInvi('SO_AFMONTANTTVA', FF); //pas géré pour l'instant
{$IFDEF CCS3}
SetInvi('SO_AFVISAACTIVITE', FF); //pas gérer pour l'instant
{$ENDIF}

TraduitForm(FF);
END;

Procedure ConfigSaisieRessource ( CC : TControl ) ;
Var FF   : TForm ;
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFFRAISGEN1',FF) then Exit ;

if (ctxScot  in V_PGI.PGIContexte) then
   begin   // a remettre si création auto tiers sur un osous traitant
   SetInvi('SO_AFFRACINEAUXI', FF);
   SetInvi('LSO_AFFRACINEAUXI', FF);
    // mcd 11/06/03 à remettre quand lien paie OK en GI
   SetInvi('SO_AFLIENPAIEVAR', FF);
   SetInvi('SO_AFLIENPAIEFIC', FF);
   SetInvi('LSO_AFLIENPAIEFIC', FF);
   SetInvi('SO_AFLIENPAIEANA', FF);
   SetInvi('SO_AFLIENPAIEDEC', FF);
   end;

TraduitForm(FF);
END;

Procedure ConfigSaisiePrefAff ( CC : TControl ) ;
Var FF   : TForm ;
C : TControl;
OldV : String ;
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFFLIBFACTAFF',FF) then Exit ;

C := GetFromNam('SO_AFFGENERAUTO', FF);
OldV:=THValComboBox(C).Value ; // GMGM 15/12/2000 bidouille car on perd la .value
{$IFDEF BTP}
THValComboBox(C).Plus :='BTP';
{$ELSE}
THValComboBox(C).Plus :='GA';
{$ENDIF}

THValComboBox(C).Value:=OldV ;
// Impossible de modifier le format de l'exercice s'il existe deja des affaires dans la base
// mcd 21/11/01 ne pas faire ... ne met pas à jour les info de code affaire !!
//SetEna( 'SO_AFFORMATEXER', FF,  Not ExisteAffaires) ;
SetEna( 'SO_AFFORMATEXER', FF,  False);

if (ctxScot  in V_PGI.PGIContexte) then
   // En contexte Scot
   begin
{$IFDEF GIGI}
  SetInvi('LSO_AFMARGEAFFAIRE', FF);
  SetInvi('SO_AFMARGEAFFAIRE', FF);
        // ne pas reforcer les valeurs si on ne vient pas du lanceur !!! mcd 29/03/01
 If V_PGI.RunFromLanceur  then
    if Not GetFlagAppli ('CGIS5.EXE') then   // PL le 17/04/03 : chgt IsAppliActive de AssistPL en GetFlagAppli de galoutil
   //  Traitement particulier à l'assistant au premier acces
        begin
        SetInvi('LSO_AFFORMATEXER', FF);
        SetInvi('SO_AFFORMATEXER', FF);
        SetInvi('LSO_AFPROFILGENER', FF);
        SetInvi('SO_AFPROFILGENER', FF);
        end;

{$ENDIF}
   If (GetParamSoc('SO_AFFORMATEXER') <>'AUC') then begin
      SetInvi('SO_AFALIMGARANTI', FF);
      SetInvi('LSO_AFALIMGARANTI', FF);
      // on voit dasn tous les cas SetInvi('SO_AFALIMCLOTURE', FF);
      // SetInvi('LSO_AFALIMCLOTURE', FF);
      end
   else begin
       SetInvi('SO_AFCALCULFIN', FF);
       SetInvi('LSO_AFCALCULFIN', FF);
       SetInvi('SCO_AFLIBCALCULLIQUID', FF);  //mcd 04/02/02
       SetInvi('SO_AFCALCULLIQUID', FF);
       SetInvi('LSO_AFCALCULLIQUID', FF);
       SetInvi('SO_AFCALCULGARANTI', FF);
       SetInvi('LSO_AFCALCULGARANTI', FF);
       SetInvi('SCO_AFLIBCALGARANTIE', FF);
        end;       
   end
else
   // En contexte Tempo
   begin
   SetInvi('SO_AFFORMATEXER', FF);
   SetInvi('LSO_AFFORMATEXER', FF);
   SetInvi('SO_AFCALCULFIN', FF);
   SetInvi('LSO_AFCALCULFIN', FF);
   SetInvi('SO_AFCALCULLIQUID', FF);
   SetInvi('SCO_AFLIBCALCULLIQUID', FF);  //mcd 02/04/02
   SetInvi('LSO_AFCALCULLIQUID', FF);
   SetInvi('SO_AFCALCULGARANTI', FF);
   SetInvi('LSO_AFCALCULGARANTI', FF);
   SetInvi('SCO_AFLIBCALGARANTIE', FF);
   SetInvi('SO_AFGERELIQUIDE', FF);
   SetInvi('SO_AFLIBLIQFACTAF', FF);
   SetInvi('LSO_AFLIBLIQFACTAF', FF);
   SetInvi('LSO_AFFINFACTURAT', FF);
   SetInvi('SO_AFFINFACTURAT', FF);
   SetInvi('SCO_AFLIBFINFACTURAT', FF);
   {$IFDEF CCS3}
   SetInvi('SO_AFGENERAUTOLIG', FF);
   SetInvi('SO_AFMULTIECHE', FF);
   {$ENDIF}
   end;

// SetInvi('SO_AFBUDZERO', FF);

TraduitForm(FF);
END;

Procedure ConfigSaisieAffReg ( CC : TControl ) ;
Var FF   : TForm ;
BEGIN
// Test de la page
  FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFGERESSAFFAIRE',FF) then Exit ;

  {$IFDEF CCS3}
  SetInvi('SO_AFGERESSAFFAIRE', FF);
  SetInvi('SO_AFCOMPLETE', FF);
  {$ENDIF}

  TraduitForm(FF);
END;

Procedure ConfigSaisieEditions ( CC : TControl ) ;
Var FF   : TForm ;
C : TControl;
OldV : String ;
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFDOCUMENTPRO',FF) then Exit ;

C := GetFromNam('SO_AFDOCUMENTPRO', FF);
OldV:=THValComboBox(C).Value ;
if (THValComboBox(C).Plus<>'ADE') then
    begin
    THValComboBox(C).Plus :='ADE';
    THValComboBox(C).Value:=OldV ;
    //THValComboBox(C).vide:=true;
    THValComboBox(C).Refresh;
    end;

C := GetFromNam('SO_AFETATPRO', FF);
OldV:=THValComboBox(C).Value ;
if (THValComboBox(C).Plus<>'APE') then
    begin
    THValComboBox(C).Plus :='APE';
    THValComboBox(C).Value:=OldV ;
    //THValComboBox(C).vide:=true;
    THValComboBox(C).Refresh;
    end;

C := GetFromNam('SO_AFDOCUMENTAFF', FF);
OldV:=THValComboBox(C).Value ;
if (THValComboBox(C).Plus<>'AFF') then
    begin
    THValComboBox(C).Plus :='AFF';
    THValComboBox(C).Value:=OldV ;
    //THValComboBox(C).vide:=true;
    THValComboBox(C).Refresh;
    end;

C := GetFromNam('SO_AFETATAFF', FF);
OldV:=THValComboBox(C).Value ;
if (THValComboBox(C).Plus<>'AFE') then
    begin
    THValComboBox(C).Plus :='AFE';
    THValComboBox(C).Value:=OldV ;
    //THValComboBox(C).vide:=true;
    THValComboBox(C).Refresh;
    end;

TraduitForm(FF);
END;

Procedure ConfigSaisieLienCompt ( CC : TControl ) ;
Var FF   : TForm ;
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_GCCPTEESCACH',FF) then Exit ;

// Gestion de la pré-saisie par la compta : on grise les champs déjà saisis
(*if IsComptaActivee then
    begin
    SetEna('SO_GCCPTEESCACH',FF,False) ;
    SetEna('SO_GCCPTEREMACH',FF,False) ;
    SetEna('SO_GCCPTEHTACH',FF,False) ;
    SetEna('SO_GCCPTEESCVTE',FF,False) ;
    SetEna('SO_GCCPTEREMVTE',FF,False) ;
    SetEna('SO_GCCPTEHTVTE',FF,False) ;
    SetEna('SO_GCPONTCOMPTABLE',FF,False) ;
    SetEna('SO_GCECARTCONVERT',FF,False) ;
    end;                                  *)

TraduitForm(FF);
END;

Procedure ConfigSaisieDevise ( CC : TControl ) ;
Var FF   : TForm ;
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_DEVISEPRINC',FF) then Exit ;

// Gestion de la pré-saisie par la compta : on grise les champs déjà saisis
{$IFDEF GIGI}
if GetFlagAppli ('CCS5.EXE') then   // PL le 17/04/03 : chgt IsAppliActive de AssistPL en GetFlagAppli de galoutil
    begin
    SetEna('SO_DEVISEPRINC',FF,False) ;
    SetEna('SO_DECVALEUR',FF,False) ;
    SetEna('SO_TENUEEURO',FF,False) ;
    SetEna('SO_DATEDEBUTEURO',FF,False) ;
    SetEna('SO_DATEBASCULE',FF,False) ;
    SetEna('SO_TAUXEURO',FF,False) ;
  (* mcd 22/09/03 suppression zones en 615   SetEna('SO_REGLEEQUILSAIS',FF,False) ;
    SetEna('SO_JALECARTEURO',FF,False) ;
    SetEna('SO_ECCEUROCREDIT',FF,False) ;
    SetEna('SO_ECCEURODEBIT',FF,False) ; *)
    end;
{$ENDIF}
TraduitForm(FF);
END;


Procedure ConfigSaisieComptes ( CC : TControl ) ;
Var FF   : TForm ;
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SCO_GRCOLLDEF',FF) then Exit ;

SetInvi('SCO_GRCARACGEN', FF);
SetInvi('SCO_GRPLANCORRES', FF);
SetInvi('LSCO_GRCARACGEN', FF);
SetInvi('LSCO_GRPLANCORRES', FF);
SetInvi('SO_ETABLISCPTA', FF);
SetInvi('SO_ETABLISDEFAUT', FF);
SetInvi('LSO_ETABLISDEFAUT', FF);
SetInvi('SO_MONTANTNEGATIF', FF);
SetInvi('SO_NUMPLANREF', FF);
SetInvi('LSO_NUMPLANREF', FF);
SetInvi('SCO_TPLAN1', FF);
SetInvi('SCO_TPLAN2', FF);
SetInvi('SO_CORSGE1', FF);
SetInvi('SO_CORSGE2', FF);
SetInvi('SO_CORSAU1', FF);
SetInvi('SO_CORSAU2', FF);
SetInvi('SO_CORSA11', FF);
SetInvi('SO_CORSA12', FF);
SetInvi('SO_CORSA21', FF);
SetInvi('SO_CORSA22', FF);
SetInvi('SO_CORSA31', FF);
SetInvi('SO_CORSA32', FF);
SetInvi('SO_CORSA41', FF);
SetInvi('SO_CORSA42', FF);
SetInvi('SO_CORSA51', FF);
SetInvi('SO_CORSA52', FF);
SetInvi('SO_CORSBU1', FF);
SetInvi('SO_CORSBU2', FF);

(*if V_PGI.RunFromLanceur then
    begin
    C := GetFromNam('SCO_GRCOLLDEF', FF);
    Larg := TBevel(C).width + 30;
    Haut := TBevel(C).height + 30;
    C := GetFromNam('SCO_GRLONG', FF);
    Haut := Haut + TBevel(C).height + 100;
    FF.Width := Larg;
    FF.Height := Haut;
    end;*)

// Gestion de la pré-saisie par la compta : on grise les champs déjà saisis
(*if IsComptaActivee then
    begin
    SetEna('SO_DEFCOLCLI',FF,False) ;
    SetEna('SO_DEFCOLFOU',FF,False) ;
    SetEna('SO_DEFCOLSAL',FF,False) ;
    SetEna('SO_DEFCOLDDIV',FF,False) ;
    SetEna('SO_DEFCOLCDIV',FF,False) ;
    SetEna('SO_DEFCOLDIV',FF,False) ;
    end;*)
SetEna('SO_LGCPTEGEN',FF,False) ;
SetEna('SO_LGCPTEAUX',FF,False) ;
SetEna('SO_LGMAXBUDGET',FF,False) ;
SetEna('SO_BOURREGEN',FF,False) ;
SetEna('SO_BOURREAUX',FF,False) ;

TraduitForm(FF);
END;

Procedure ConfigSaisiePlanning(CC : TControl);
Var FF : TForm ;
BEGIN
  // Test de la page
  FF := TForm(CC.Owner) ; if Not ExisteNam('SO_AFPLANDECHARGE',FF) then Exit ;

//  SetEna('SO_AFPLANHEURE', FF, False);

  SetInvi('SO_AFPLANHEURE', FF);
  SetInvi('LSO_AFPLANHEURE', FF);
  SetInvi('SO_AFGESTIONRAF', FF);
  {$IFDEF CCS3}
  SetInvi('LSCO_AFPLANNINGMOIS',FF);
  SetInvi('SCO_AFPLANNINGMOIS',FF);
  SetInvi('SO_AFPLANDECHARGE',FF);
  SetInvi('SO_AFPDCDEC',FF);
  SetInvi('SO_AFRAFPLANNING',FF);
  {$ENDIF}

  TraduitForm(FF);
END;

Procedure ConfigSaisieGenePlanning(CC : TControl);
Var FF : TForm ;
BEGIN
  // Test de la page
  FF := TForm(CC.Owner) ; if Not ExisteNam('SO_AFGESTIONRESS',FF) then Exit ;
  SetInvi('SO_AFGESTIONRESS',FF);
END;

Procedure ConfigSaisieRevEtVar(CC : TControl);
Var FF : TForm ;
BEGIN
  // Test de la page
  FF := TForm(CC.Owner) ; if Not ExisteNam('SO_AFREVPATH',FF) then Exit ;

  // revision de prix tempo en attendant de developper la fonctionnalité
  SetInvi('SO_AFJUSREVDET', FF);

  TraduitForm(FF);

end;

Procedure ChargeFourchettes ( CC : TControl ; AuCharge : boolean ) ;
Var FF   : TForm ;
BEGIN
if Not V_PGI.RunFromLanceur then exit;
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_BILDEB1',FF) then Exit ;
// Traitement
GetCompteFourchette(CC,'BIL') ;
GetCompteFourchette(CC,'CHA') ;
GetCompteFourchette(CC,'PRO') ;
GetCompteFourchette(CC,'EXT') ;
if AuCharge then GetTOBCpts(CC).SetAllModifie(False) ;
END ;

{$IFDEF CCS3}
Procedure InviFourcheAF ( St : String ; FF : TForm ) ;
Var i,k : integer ;
    Nam,Suff : String ;
BEGIN
for k:=1 to 2 do
    BEGIN
    if k=1 then Suff:='DEB' else Suff:='FIN' ;
    for i:=3 to 5 do
        BEGIN
        Nam:='SO_'+St+Suff+IntToStr(i) ;
        SetInvi(Nam,FF) ; SetInvi('L'+Nam,FF) ;
        END ;
    END ;
END ;
{$ENDIF}

Procedure CacheFonctionS5 ( CC : TControl ) ;
Var FF : TForm ;
BEGIN
FF:=GetLaForm(CC) ;
   if ((EstSerie(S5)) or (EstSerie(S3))) then
   BEGIN
   if ExisteNam('SO_CORSGE2',FF) then
      BEGIN
      SetInvi('SO_CORSGE2',FF) ; SetInvi('SO_CORSAU2',FF) ;
      SetInvi('SO_CORSA12',FF) ; SetInvi('SO_CORSA22',FF) ;
      SetInvi('SO_CORSA32',FF) ; SetInvi('SO_CORSA42',FF) ;
      SetInvi('SO_CORSA52',FF) ; SetInvi('SO_CORSBU2',FF) ;
      SetInvi('SO_CORSA31',FF) ; SetInvi('SO_CORSA41',FF) ;
      SetInvi('SO_CORSA51',FF) ;
      GetFromNam('SO_CORSBU1',FF).Top:=GetFromNam('SO_CORSA31',FF).Top ;
      SetInvi('SCO_TPLAN2',FF) ;
      END ;
   if ExisteNam('SO_BOURRELIB',FF) then SetInvi('SO_BOURRELIB',FF) ;
   if ExisteNam('LSO_BOURRELIB',FF) then SetInvi('LSO_BOURRELIB',FF) ;
   if ExisteNam('SO_DATEREVISION',FF) then
      BEGIN
      SetInvi('SO_DATEREVISION',FF) ; SetInvi('LSO_DATEREVISION',FF) ;
     END ;
   END ;
   ////   GM il faut se poser la quetion , pourquoi utilsocaf et utilsoc...attention redondance
{$IFDEF CCS3}
   // Comptabilité S3 //
if ExisteNam('SO_JALREPBALAN',FF) then
   BEGIN
   SetInvi('SO_JALREPBALAN',FF) ; SetInvi('LSO_JALREPBALAN',FF) ;
   END ;
if ExisteNam('SO_EXTDEB1',FF) then
   BEGIN
   SetInvi('SO_EXTDEB1',FF) ; SetInvi('LSO_EXTDEB1',FF) ;
   END ;
if ExisteNam('SO_EXTFIN1',FF) then
   BEGIN
   SetInvi('SO_EXTFIN1',FF) ; SetInvi('LSO_EXTFIN1',FF) ;
   END ;
if ExisteNam('SO_EXTDEB2',FF) then
   BEGIN
   SetInvi('SO_EXTDEB2',FF) ; SetInvi('LSO_EXTDEB2',FF) ;
   END ;
if ExisteNam('SO_EXTFIN2',FF) then
   BEGIN
   SetInvi('SO_EXTFIN2',FF) ; SetInvi('LSO_EXTFIN2',FF) ;
   END ;
if ExisteNam('SCO_GREXCPTABLE',FF) then
   BEGIN
   SetInvi('SCO_GREXCPTABLE',FF) ;
   END ;
if ExisteNam('LSCO_GREXCPTABLE',FF) then
   BEGIN
   SetInvi('LSCO_GREXCPTABLE',FF) ;
   END ;
if ExisteNam('SO_BILDEB1',FF) then
   BEGIN
   InviFourcheAF('BIL',FF) ; InviFourcheAF('CHA',FF) ; InviFourcheAF('PRO',FF) ;
   END ;
if ExisteNam('SO_CORSGE1',FF) then SetInvi('SO_CORSGE1',FF) ;
if ExisteNam('SO_CORSAU1',FF) then SetInvi('SO_CORSAU1',FF) ;
if ExisteNam('SO_CORSA21',FF) then SetInvi('SO_CORSA21',FF) ;
if ExisteNam('SO_CORSA11',FF) then SetInvi('SO_CORSA11',FF) ;
if ExisteNam('SO_CORSA41',FF) then SetInvi('SO_CORSA41',FF) ;
if ExisteNam('SO_CORSA31',FF) then SetInvi('SO_CORSA31',FF) ;
if ExisteNam('SO_CORSA51',FF) then SetInvi('SO_CORSA51',FF) ;
if ExisteNam('SO_CORSBU1',FF) then SetInvi('SO_CORSBU1',FF) ;
if ExisteNam('SCO_TPLAN1',FF) then SetInvi('SCO_TPLAN1',FF) ;
if ExisteNam('SCO_GRPLANCORRES',FF) then SetInvi('SCO_GRPLANCORRES',FF) ;
if ExisteNam('LSCO_GRPLANCORRES',FF) then SetInvi('LSCO_GRPLANCORRES',FF) ;
if ExisteNam('SO_LGMAXBUDGET',FF) then SetInvi('SO_LGMAXBUDGET',FF)   ;
if ExisteNam('LSO_LGMAXBUDGET',FF) then SetInvi('LSO_LGMAXBUDGET',FF) ;
if ExisteNam('SO_EQUILANAODA',FF) then SetInvi('SO_EQUILANAODA',FF) ;
if ExisteNam('SO_CPLIBREANAOBLI',FF) then SetInvi('SO_CPLIBREANAOBLI',FF) ;
if ExisteNam('SO_DECQTE',FF) then
   BEGIN
   SetInvi('SO_DECQTE',FF)      ; SetInvi('LSO_DECQTE',FF) ;
   END ;
if ExisteNam('SCO_GRBUDGET',FF) then
   BEGIN
   SetInvi('SCO_GRBUDGET',FF)   ; SetInvi('LSCO_GRBUDGET',FF)   ;
   END ;
if ExisteNam('SO_DUPSECTBUD',FF) then SetInvi('SO_DUPSECTBUD',FF)  ;
if ExisteNam('SO_CONTROLEBUD',FF) then SetInvi('SO_CONTROLEBUD',FF) ;
if ExisteNam('SO_SUIVILOG',FF) then SetInvi('SO_SUIVILOG',FF)    ;
if ExisteNam('SO_JALCTRLBUD',FF) then
   BEGIN
   SetInvi('SO_JALCTRLBUD',FF)  ; SetInvi('LSO_JALCTRLBUD',FF)  ;
   END ;
if ExisteNam('SO_JALLOOKUP',FF) then SetInvi('SO_JALLOOKUP',FF)     ;
if ExisteNam('SO_ETABLOOKUP',FF) then SetInvi('SO_ETABLOOKUP',FF) ;
if ExisteNam('SO_CPREFLETTRAGE',FF) then
   BEGIN
   SetInvi('SO_CPREFLETTRAGE',FF) ; SetInvi('LSO_CPREFLETTRAGE',FF) ;
   END ;
if ExisteNam('SO_CPREFPOINTAGE',FF) then
   BEGIN
   SetInvi('SO_CPREFPOINTAGE',FF) ; SetInvi('LSO_CPREFPOINTAGE',FF) ;
   END ;
if ExisteNam('SCO_CPLREFS',FF) then SetInvi('SCO_CPLREFS',FF)      ;
if ExisteNam('SO_OUITP',FF) then SetInvi('SO_OUITP',FF) ;
if ExisteNam('SO_JALVTP',FF) then
   BEGIN
   SetInvi('SO_JALVTP',FF) ; SetInvi('LSO_JALVTP',FF) ;
   END ;
if ExisteNam('SO_JALATP',FF) then
   BEGIN
   SetInvi('SO_JALATP',FF) ; SetInvi('LSO_JALATP',FF) ;
   END ;
if ExisteNam('SO_TAUXCOUTFITIERS',FF) then
   BEGIN
   SetInvi('SO_TAUXCOUTFITIERS',FF) ; SetInvi('LSO_TAUXCOUTFITIERS',FF) ;
   END ;
if ExisteNam('SO_LETEGALC',FF) then SetInvi('SO_LETEGALC',FF) ;
if ExisteNam('SO_LETEGALF',FF) then SetInvi('SO_LETEGALF',FF) ;
if ExisteNam('SO_LETTOLERC',FF) then
   BEGIN
   SetInvi('SO_LETTOLERC',FF) ; SetInvi('LSO_LETTOLERC',FF) ;
   END ;
if ExisteNam('SO_LETTOLERF',FF) then
   BEGIN
   SetInvi('SO_LETTOLERF',FF) ; SetInvi('LSO_LETTOLERF',FF) ;
   END ;
if ExisteNam('SCO_GRTVAENCAIS',FF) then
   BEGIN
   SetInvi('SCO_GRTVAENCAIS',FF) ; SetInvi('LSCO_GRTVAENCAIS',FF) ;
   END ;
if ExisteNam('SO_CPREFPOINTAGE',FF) then
   BEGIN
   SetInvi('SO_CPREFPOINTAGE',FF) ; SetInvi('LSO_CPREFPOINTAGE',FF) ;
   END ;
if ExisteNam('SO_REGLEEQUILSAIS',FF) then SetEna('SO_REGLEEQUILSAIS',FF,False) ;
if ExisteNam('SO_DATEDEBUTEURO',FF) then SetEna('SO_DATEDEBUTEURO',FF,False) ;
   // Gestion Commerciale S3 //
if ExisteNam('SO_GCVENTAXE2',FF) then
   BEGIN
   SetInvi('SO_GCVENTAXE2',FF)    ; SetInvi('SO_GCVENTAXE3',FF) ;
   SetInvi('SO_GCVENTCPTAAFF',FF) ; SetInvi('SO_GCDISTINCTAFFAIRE',FF) ;
   SetInvi('SO_GCCPTAIMMODIV',FF) ;
   END ;
if ExisteNam('SO_GCINVPERM',FF) then
   BEGIN
   SetInvi('SO_GCINVPERM',FF)       ;
   SetInvi('SO_GCJALSTOCK',FF)      ; SetInvi('LSO_GCJALSTOCK',FF) ;
   SetInvi('SO_GCMODEVALOSTOCK',FF) ; SetInvi('LSO_GCMODEVALOSTOCK',FF) ;
   END ;
if ExisteNam('SO_GCFAMHIERARCHIQUE',FF) then
   BEGIN
   // DBR Test existance avant de rendre invisble - Fiche 10745
   if ExisteNam('SO_GCDIMCOLLECTION', FF) then SetInvi('SO_GCDIMCOLLECTION',FF) ;
   if ExisteNam('SO_GCCDUSEIT', FF) then SetInvi('SO_GCCDUSEIT',FF);
   if ExisteNam('SCO_GCJPEGMEMO', FF) then SetInvi('SCO_GCJPEGMEMO',FF) ;
   if ExisteNam('SO_GCCDJPG1', FF) then SetInvi('SO_GCCDJPG1',FF) ;
   if ExisteNam('LSO_GCCDJPG1', FF) then SetInvi('LSO_GCCDJPG1',FF) ;
   if ExisteNam('SO_GCCDJPG2', FF) then SetInvi('SO_GCCDJPG2',FF) ;
   if ExisteNam('LSO_GCCDJPG2', FF) then SetInvi('LSO_GCCDJPG2',FF) ;
   if ExisteNam('SO_GCCDJPG3', FF) then SetInvi('SO_GCCDJPG3',FF) ;
   if ExisteNam('LSO_GCCDJPG3', FF) then SetInvi('LSO_GCCDJPG3',FF) ;
   if ExisteNam('SO_GCCDJPG4', FF) then SetInvi('SO_GCCDJPG4',FF) ;
   if ExisteNam('LSO_GCCDJPG4', FF) then SetInvi('LSO_GCCDJPG4',FF) ;
   if ExisteNam('SO_GCCDMEM1', FF) then SetInvi('SO_GCCDMEM1',FF) ;
   if ExisteNam('LSO_GCCDMEM1', FF) then SetInvi('LSO_GCCDMEM1',FF) ;
   if ExisteNam('SO_GCCDMEM2', FF) then SetInvi('SO_GCCDMEM2',FF) ;
   if ExisteNam('LSO_GCCDMEM2', FF) then SetInvi('LSO_GCCDMEM2',FF) ;
   if ExisteNam('SO_GCCDMEM3', FF) then SetInvi('SO_GCCDMEM3',FF) ;
   if ExisteNam('LSO_GCCDMEM3', FF) then SetInvi('LSO_GCCDMEM3',FF) ;
   if ExisteNam('SO_GCCDMEM4', FF) then SetInvi('SO_GCCDMEM4',FF) ;
   if ExisteNam('LSO_GCCDMEM4', FF) then SetInvi('LSO_GCCDMEM4',FF) ;
   END ;
 // DBR Fiche 11015 - Debut
if ExisteNam ('SCO_GCGRLOGITRANSF', FF) then SetInvi ('SCO_GCGRLOGITRANSF', FF);
if ExisteNam ('LSCO_GCGRLOGITRANSF', FF) then SetInvi ('LSCO_GCGRLOGITRANSF', FF);
SetParamSoc ('SO_GCTRV', False);
if ExisteNam ('SCO_GRPCPTAAUTRE', FF) then SetInvi ('SCO_GRPCPTAAUTRE', FF);
if ExisteNam ('LSCO_GRPCPTAAUTRE', FF) then SetInvi ('LSCO_GRPCPTAAUTRE', FF);
// DBR Fiche 11015 - Fin
if ExisteNam('SO_GCTRV',FF) then SetInvi('SO_GCTRV',FF) ;
if ExisteNam('SCO_RTDIMINFACT',FF) then
   begin
   SetInvi('SCO_RTDIMINFACT',FF) ;
   SetInvi('LSCO_RTDIMINFACT',FF) ;
   SetInvi('SO_RTLARGEURFICHE001',FF) ;
   SetInvi('LSO_RTLARGEURFICHE001',FF) ;
   SetInvi('SO_RTHAUTEURFICHE001',FF) ;
   SetInvi('LSO_RTHAUTEURFICHE001',FF) ;
   SetInvi('SO_RTGESTINFOS001',FF) ;
   end;
if ExisteNam('SCO_RTDIMINFOPE',FF) then
   begin
   SetInvi('SCO_RTDIMINFOPE',FF) ;
   SetInvi('LSCO_RTDIMINFOPE',FF) ;
   SetInvi('SO_RTLARGEURFICHE002',FF) ;
   SetInvi('LSO_RTLARGEURFICHE002',FF) ;
   SetInvi('SO_RTHAUTEURFICHE002',FF) ;
   SetInvi('LSO_RTHAUTEURFICHE002',FF) ;
   SetInvi('SO_RTGESTINFOS002',FF) ;
   end;
if ExisteNam('SO_RTCONFIDENTIALITE',FF) then SetInvi('SO_RTCONFIDENTIALITE',FF) ;
if ExisteNam('SO_GCTRV',FF) then SetInvi('SO_GCTRV',FF) ;
if ExisteNam('SO_GCACHATTTC',FF) then SetInvi('SO_GCACHATTTC',FF) ;
if ExisteNam('SO_GCCPTAIMMODIV',FF) then SetInvi('SO_GCCPTAIMMODIV',FF) ;
if ExisteNam('SO_GCDESACTIVECOMPTA',FF) then SetInvi('SO_GCDESACTIVECOMPTA',FF) ;
if ExisteNam('SO_GCINVPERM',FF)       then SetInvi('SO_GCINVPERM',FF) ;
if ExisteNam('SO_GCJALSTOCK',FF)      then BEGIN SetInvi('SO_GCJALSTOCK',FF) ; SetInvi('LSO_GCJALSTOCK',FF) ; END ;
if ExisteNam('SO_GCMODEVALOSTOCK',FF) then BEGIN SetInvi('SO_GCMODEVALOSTOCK',FF) ; SetInvi('LSO_GCMODEVALOSTOCK',FF) ; END ;
if ExisteNam('SO_GCCPTERGVTE',FF)     then BEGIN SetInvi('SO_GCCPTERGVTE',FF) ; SetInvi('LSO_GCCPTERGVTE',FF) ; END ;
if ExisteNam('SO_GCCPTESTOCK',FF)     then BEGIN SetInvi('SO_GCCPTESTOCK',FF) ; SetInvi('LSO_GCCPTESTOCK',FF) ; END ;
if ExisteNam('SO_GCCPTEVARSTK',FF)    then BEGIN SetInvi('SO_GCCPTEVARSTK',FF) ; SetInvi('LSO_GCCPTEVARSTK',FF) ; END ;
if ExisteNam('SO_GCFICHEEDITPROPO',FF) then SetInvi('SO_GCFICHEEDITPROPO',FF) ;
{$ENDIF}

//////
END ;

Function ChargePageSocGA ( CC : TControl ) : boolean ;
BEGIN
CacheFonctionS5(CC) ;
ShowParamSoc(CC) ;
BlocageDevise(CC) ;
BlocageECC(CC) ;
InitDivers(CC) ;
ChargeFourchettes(CC,True) ;

// Configuration des écran GA (traduction, champs visibles...)
ConfigSaisieComptes(CC) ;
ConfigSaisieParamsDefaut(CC);
ConfigSaisieLienCompt(CC) ;
ConfigSaisieDevise(CC);
ConfigSaisieRessource(CC) ;
ConfigSaisiePrefAff(CC) ;
ConfigSaisieAffReg(CC) ;
ConfigSaisieActivite(CC) ;
ConfigSaisiePreferences(CC) ;
ConfigSaisieEditions(CC);
ConfigSaisieDates(CC);
ConfigSaisieApprec(CC);   //mcd 12/02/02
ConfigSaisieCutOff(CC);   //mcd 11/07/02
ConfigSaisieDecen(CC);   //mcd 11/07/02
ConfigSaisieBudget(CC);   
ConfigSaisieFactEclat(CC);
ConfigSaisiePlanning(CC);
ConfigSaisieGenePlanning(CC);
ConfigSaisieRevEtVar(CC);
Result:=True ;
END ;

{============================= Sur Sauvegarde =========================}
Function GeneAttenteOk ( CC : TControl ) : Boolean ;
Var Q : TQuery ;
    OkOk : Boolean ;
    GeneAtt : String ;
    CptAtt : THCritMaskEdit ;
    FF  : TForm ;
BEGIN
// Test de la page
Result:=True ; FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_GENATTEND',FF) then Exit ;
// Traitement
CptAtt:=THCritMaskEdit(GetFromNam('SO_GENATTEND',FF)) ; GeneAtt:=CptAtt.Text ;
if GeneAtt='' then BEGIN Result:=True ; Exit ; END ;
   // ATTENTIOn si ajout champ, revoir les Fields[x] !!!!
Q:=OpenSql('SELECT G_GENERAL, G_VENTILABLE, G_COLLECTIF, G_LETTRABLE, G_POINTABLE FROM GENERAUX WHERE G_GENERAL="'+GeneAtt+'"',True) ;
if Q.Fields[0].AsString='' then
   BEGIN
   Ferme(Q) ; Exit ;
   END else
   BEGIN
   OkOk:=True ;
   if (Q.Fields[1].AsString='X') And (OkOk) then BEGIN OkOk:=False ; HShowMessage('5;Société;Le compte général d''attente ne peut pas être ventilable;W;O;O;O;','','') ; END ;
   if (Q.Fields[2].AsString='X') And (OkOk) then BEGIN OkOk:=False ; HShowMessage('6;Société;Le compte général d''attente ne peut pas être collectif;W;O;O;O;','','') ; END ;
   if (Q.Fields[3].AsString='X') And (OkOk) then BEGIN OkOk:=False ; HShowMessage('7;Société;Le compte général d''attente ne peut pas être lettrable;W;O;O;O;','','') ; END ;
   if (Q.Fields[4].AsString='X') And (OkOk) then BEGIN OkOk:=False ; HShowMessage('8;Société;Le compte général d''attente ne peut pas être pointable;W;O;O;O;','','') ; END ;
   END ;
Ferme(Q) ;
if Not OkOk then BEGIN if CptAtt.CanFocus then CptAtt.SetFocus ; Result:=False ; Exit ; END ;
END ;

Function DefautCpteOk ( CC : TControl ) : Boolean ;
Var Q : TQuery ;
    i : integer ;
    Sql,LeWhere : String ;
    Cpt,Nam : String ;
    FF  : TForm ;
    LaZone : THCritMaskEdit ;
BEGIN
// Test de la page
Result:=True ; FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_COLDEFCLI',FF) then Exit ;
// Traitement
Result:=True ;
Q:=TQuery.Create(Application) ; Q.DataBaseName:='SOC' ;
for i:=0 to FF.ComponentCount-1 do
   BEGIN
   Nam:=FF.Components[i].Name ; if Copy(Nam,1,9)<>'SO_DEFCOL' then Continue ;
   LaZone:=THCritMaskEdit(FF.Components[i]) ;
   Cpt:=LaZone.Text ; if Cpt='' then Continue ;
   if Nam='SO_DEFCOLCLI'  then LeWhere:='And G_COLLECTIF="X" And G_NATUREGENE="COC"' else
   if Nam='SO_DEFCOLFOU'  then LeWhere:='And G_COLLECTIF="X" And G_NATUREGENE="COF"' else
   if Nam='SO_DEFCOLSAL'  then LeWhere:='And G_COLLECTIF="X" And G_NATUREGENE="COS"' else
   if Nam='SO_DEFCOLDDIV' then LeWhere:='And G_COLLECTIF="X" And (G_NATUREGENE="COC" Or G_NATUREGENE="COD")' else
   if Nam='SO_DEFCOLCDIV' then LeWhere:='And G_COLLECTIF="X" And (G_NATUREGENE="COF" Or G_NATUREGENE="COD")' else
   if Nam='SO_DEFCOLDIV'  then LeWhere:='And G_COLLECTIF="X" And G_NATUREGENE="COD"' ;
   Sql:='SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="'+Cpt+'" '+LeWhere+'' ;
   Q.Sql.Clear ; Q.Sql.Add(Sql) ; ChangeSql(Q) ;
   Q.Open ; if Not Q.EOF then Cpt:=Q.Fields[0].AsString else Cpt:='' ; Q.Close ;
   if Cpt='' then
      BEGIN
      if LaZone.CanFocus then LaZone.SetFocus ;
      HShowMessage('16;Société;Le compte que vous avez choisi n''est pas en accord avec la nature attendue;W;O;O;O;','','') ;
      Result:=False ; Break ;
      END ;
   END ;
Ferme(Q) ;
END ;

Function VerifDateEntreeEuroOk ( CC : TControl ) : Boolean ;
Var FF : TForm ;
    DevPrinc : THValComboBox ;
    DateDeb  : THCritMaskEdit ;
BEGIN
// Test de la page
Result:=True ; FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_DATEDEBUTEURO',FF) then Exit ;
// Traitement
DevPrinc:=THValComboBox(GetFromNam('SO_DEVISEPRINC',FF)) ;
DateDeb:=THCritMaskEdit(GetFromNam('SO_DATEDEBUTEURO',FF)) ;
if StrToDate(DateDeb.Text)<Encodedate(1999,01,01) then BEGIN HShowMessage('20;Société;La date d''entrée en vigueur de l''euro doit être comprise entre le 1er janvier 1999 et le 31 décembre 1999;W;O;O;O;','','') ; Result:=False ; END ;
if StrToDate(DateDeb.Text)>Encodedate(1999,12,31) then
   BEGIN
   if ((DevPrinc.Value<>'') and (V_PGI.DevisePivot<>'') and (Not Devprinc.Enabled) and (Not EstMonnaieIn(DEVPRINC.Value)))
      then else BEGIN HShowMessage('20;Société;La date d''entrée en vigueur de l''euro doit être comprise entre le 1er janvier 1999 et le 31 décembre 1999;W;O;O;O;','','') ; Result:=False ; END ;
   END ;
END ;

function VerifCptesEuro ( CC : TControl ) : boolean ;
Var CptD,CptC : String ;
    Okok,OkVent : Boolean ;
    QQ   : TQuery ;
    FF  : TForm ;
    CHDebit,CHCredit : THCritMaskEdit ;
BEGIN
// Test de la page
Result:=True ; FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_ECCEURODEBIT',FF) then Exit ;
// Traitement
Okok:=True ;
CHDebit:=THCritMaskEdit(GetFromNam('SO_ECCEURODEBIT',FF))   ; CptD:=CHDebit.Text ;
CHCredit:=THCritMaskEdit(GetFromNam('SO_ECCEUROCREDIT',FF)) ; CptC:=CHCredit.Text ;
if ((CptD='') and (CptC='')) then Exit ;
if ((CptD<>'') and (Not RechExisteH(CHDebit))) then
   BEGIN
   HShowMessage('23;Société;Les comptes d''écart de conversion ne sont pas correctement renseignés;W;O;O;O;','','') ;
   Result:=False ; Exit ;
   END ;
if ((CptC<>'') and (Not RechExisteH(CHCredit))) then
   BEGIN
   HShowMessage('23;Société;Les comptes d''écart de conversion ne sont pas correctement renseignés;W;O;O;O;','','') ;
   Result:=False ; Exit ;
   END ;
if CptD<>'' then
   BEGIN
   OkVent:=False ; Okok:=True ;
   QQ:=OpenSQL('SELECT G_VENTILABLE FROM GENERAUX WHERE G_GENERAL="'+CptD+'"',True) ;
   if QQ.EOF then Okok:=False else OkVent:=(QQ.Fields[0].AsString='X') ;
   Ferme(QQ) ;
   if Not Okok then
      BEGIN
      HShowMessage('23;Société;Les comptes d''écart de conversion ne sont pas correctement renseignés;W;O;O;O;','','') ;
      Result:=False ; Exit ;
      END else if OkVent then
      BEGIN
      HShowMessage('24;Société;Les comptes d''écart de conversion ne doivent pas être ventilables;W;O;O;O;','','') ;
      Result:=False ; Exit ;
      END ;
   END ;
if CptC<>'' then
   BEGIN
   OkVent:=False ; Okok:=True ;
   QQ:=OpenSQL('SELECT G_VENTILABLE FROM GENERAUX WHERE G_GENERAL="'+CptC+'"',True) ;
   if QQ.EOF then Okok:=False else OkVent:=(QQ.Fields[0].AsString='X') ;
   Ferme(QQ) ;
   if Not Okok then
      BEGIN
      HShowMessage('23;Société;Les comptes d''écart de conversion ne sont pas correctement renseignés;W;O;O;O;','','') ;
      Result:=False ; Exit ;
      END else if Okvent then
      BEGIN
      HShowMessage('24;Société;Les comptes d''écart de conversion ne doivent pas être ventilables;W;O;O;O;','','') ;
      Result:=False ; Exit ;
      END ;
   END ;
Result:=Okok ;
END ;

function ExisteCptesAttente ( CC : TControl ) : boolean ;
Var FF : TForm ;
BEGIN
// Test de la page
Result:=True ; FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_GENATTEND',FF) then Exit ;
// Traitement
Result:=False ;
if Not RechExisteH(GetFromNam('SO_GENATTEND',FF)) then BEGIN HShowmessage('13;Société;Vous devez renseigner un compte général d''attente;W;O;O;O;','','') ; Exit ; END ;
if Not RechExisteH(GetFromNam('SO_CLIATTEND',FF)) then BEGIN HShowmessage('14;Société;Vous devez renseigner un compte auxiliaire client d''attente;W;O;O;O;','','') ; Exit ; END ;
if Not RechExisteH(GetFromNam('SO_FOUATTEND',FF)) then BEGIN HShowmessage('15;Société;Vous devez renseigner un compte auxiliaire fournisseur d''attente;W;O;O;O;','','') ; Exit ; END ;
Result:=True ;
END ;

Function LongueurCpteOk ( CC : TControl ) : Boolean ;
Var FF : TForm ;
    SS : TSpinEdit ;
BEGIN
// Test de la page
Result:=True ; FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_LGCPTEGEN',FF) then Exit ;
// Traitement
Result:=False ;
SS:=TSpinEdit(GetFromNam('SO_LGCPTEGEN',FF)) ;
if SS.Value<>VH^.Cpta[fbGene].Lg then if DesEnregs('GENERAUX') then
   BEGIN
   HShowMessage('4;Société;Vous ne pouvez pas modifier la longueur des comptes : des comptes sont déjà définis.;W;O;O;O;','','') ;
   if SS.CanFocus then SS.SetFocus ;
   Exit ;
   END ;
SS:=TSpinEdit(GetFromNam('SO_LGCPTEAUX',FF)) ;
if SS.Value<>VH^.Cpta[fbAux].Lg then if DesEnregs('TIERS') then
   BEGIN
   HShowMessage('4;Société;Vous ne pouvez pas modifier la longueur des comptes : des comptes sont déjà définis.;W;O;O;O;','','') ;
   if SS.CanFocus then SS.SetFocus ;
   Exit ;
   END ;
Result:=True ;
END ;

function VerifPariteEuro ( CC : TControl ) : boolean ;
Var StParite : String ;
    Parite   : Integer ;
    FF  : TForm ; 
    CHEuro : THNumEdit ;
BEGIN
// Test de la page
Result:=True ; FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_TAUXEURO',FF) then Exit ;
// Traitement
Result:=True ;
CHEuro:=THNumEdit(GetFromNam('SO_TAUXEURO',FF)) ;
StParite:=Trim(CHEuro.Text) ;
if Pos(V_PGI.SepDecimal,StParite)<>0 then Delete(StParite,Pos(V_PGI.SepDecimal,StParite),1) ;
if StParite<>'' then Parite:=StrToInt(StParite) else Parite:=0 ;
if Parite<=0 then BEGIN HShowMessage('22;Société;La parité fixe par rapport à l''Euro doit être positive et de 6 chiffres significatifs maximum !;W;O;O;O;','','') ; ; Result:=False ; Exit ; END ;
StParite:=IntToStr(Parite) ;
if Length(StParite)>6 then BEGIN HShowMessage('22;Société;La parité fixe par rapport à l''Euro doit être positive et de 6 chiffres significatifs maximum !;W;O;O;O;','','') ; Result:=False ; END ;
if Not Result then if CHEuro.CanFocus then CHEuro.SetFocus ;
END ;

Function CarBourreOk ( CC : TControl ) : Boolean ;
Var CB : TEdit ;
    FF  : TForm ;
    Okok : boolean ;
    CCH  : Set of Char ;
    CH   : Char ;
BEGIN
// Test de la page
Result:=True ; FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_BOURREGEN',FF) then Exit ;
// Traitement
Result:=False ; Okok:=True ;
CCH:=[' ','*','%','?','#',':','_',',','|','"','''',';'] ;
CB:=TEdit(GetFromNam('SO_BOURREGEN',FF)) ;
if (CB.Text='') or (CB.Text=' ') then Okok:=False else
   BEGIN
   CH:=CB.Text[1] ; if CH in CCH then Okok:=False ;
   END ;
if Not Okok then
   BEGIN
   HShowMessage('21;Société;Vous devez renseigner un caractère de bourrage;W;O;O;O;','','') ;
   CB.SetFocus ; Exit ;
   END ;
CB:=TEdit(GetFromNam('SO_BOURREAUX',FF)) ;
if (CB.Text='') or (CB.Text=' ') then Okok:=False else
   BEGIN
   CH:=CB.Text[1] ; if CH in CCH then Okok:=False ;
   END ;
if Not Okok then
   BEGIN
   HShowMessage('21;Société;Vous devez renseigner un caractère de bourrage;W;O;O;O;','','') ;
   CB.SetFocus ; Exit ;
   END ;
Result:=True ;
END ;

Function EtabDefautOk ( CC : TControl ) : Boolean ;
Var CB : THvalComboBox ;
    FF  : TForm ;
BEGIN
// Test de la page
Result:=True ; FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_ETABLISDEFAUT',FF) then Exit ;
// Traitement
CB:=THValComboBox(GetFromNam('SO_ETABLISDEFAUT',FF)) ;
if CB.Value='' then
   BEGIN
   Result:=FALSE ; HShowMessage('17;Société;Vous devez renseigner un établissement par défaut.;W;O;O;O;','','') ;
   if CB.CanFocus then
        CB.SetFocus ;
   END ;
END ;

Function VerifJalTP ( CC : TControl ) : boolean ;
Var FF  : TForm ;
    CBA : THValComboBox ;
    Q   : TQuery ;
    Trouv : boolean ;
BEGIN
  // Test de la page
  Result:=True ; FF:=GetLaForm(CC) ; 
  if Not ExisteNam('SO_JALATP',FF) then Exit ;
  // Traitement
  Result:=False ;
  CBA:=THValComboBox(GetFromNam('SO_JALATP',FF)) ;
  if CBA.Value<>'' then
  BEGIN
    Q:=OpenSQL('SELECT J_MULTIDEVISE FROM JOURNAL WHERE J_JOURNAL="'+CBA.Value+'" AND J_MULTIDEVISE="X"',True) ;
    Trouv:=Not Q.EOF ;
    Ferme(Q) ;
    if Not Trouv then BEGIN HShowMessage('3;Société;Les journaux pour tiers payeurs doivent être multi-devise;W;O;O;O;','','') ; CBA.SetFocus ; Exit ; END ;
  END ;
  CBA:=THValComboBox(GetFromNam('SO_JALVTP',FF)) ;
  if CBA.Value<>'' then
  BEGIN
    Q:=OpenSQL('SELECT J_MULTIDEVISE FROM JOURNAL WHERE J_JOURNAL="'+CBA.Value+'" AND J_MULTIDEVISE="X"',True) ;
    Trouv:=Not Q.EOF ;
    Ferme(Q) ;
    if Not Trouv then BEGIN HShowMessage('3;Société;Les journaux pour tiers payeurs doivent être multi-devise;W;O;O;O;','','') ; CBA.SetFocus ; Exit ; END ;
  END ;
  Result:=True ;
END ;

Function VerifCoord ( CC : TControl ) : boolean ;
Var FF : TForm ;
    CH : TEdit ;
BEGIN
// Test de la page
Result:=True ; FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_LIBELLE',FF) then Exit ;
// Traitement
Result:=False ;
CH:=TEdit(GetFromNam('SO_LIBELLE',FF)) ;
if CH.Text='' then BEGIN HShowMessage('1;Société;Vous devez renseigner un libellé;W;O;O;O;','','') ; CH.SetFocus ; Exit ; END ;
CH:=TEdit(GetFromNam('SO_ADRESSE1',FF)) ;
if CH.Text='' then BEGIN HShowMessage('3;Société;Vous devez renseigner une adresse;W;O;O;O;','','') ; CH.SetFocus ; Exit ; END ;
Result:=True ;
END ;

Function FourchetteVide ( CC : TControl ; NN : String ) : Boolean ;
Var i : Integer ;
    Nam1,Nam2 : String ;
    TT  : TOB ;
BEGIN
Result:=False ; TT:=GetTOBCpts(CC) ;
for i:=1 to 5 do
    BEGIN
    Nam1:='SO_'+NN+'DEB'+IntToStr(i) ; Nam2:='SO_'+NN+'FIN'+IntToStr(i) ;
    if ((TT.GetValue(Nam1)<>'') and (TT.GetValue(Nam2)<>'')) then Exit ;
    END ;
HShowMessage('9;Société;Vous devez renseigner les fourchettes de comptes;W;O;O;O;','','') ;
Result:=True ;
END ;

Function FourchetteOk ( CC : TControl ) : Boolean ;
BEGIN
Result:=False ;
if FourchetteVide(CC,'BIL') then Exit ;
if FourchetteVide(CC,'PRO') then Exit ;
if FourchetteVide(CC,'CHA') then Exit ;
Result:=True ;
END ;

Function ManqueUnCpte ( CC : TControl ) : Boolean ;
Var i,lei : Integer ;
    Nam1,Nam2,NN   : String ;
    Manque  : boolean ;
    TT  : TOB ;
BEGIN
TT:=GetTOBCpts(CC) ; Manque:=False ;
for lei:=1 to 4 do if Not Manque then
    BEGIN
    Case Lei of 1 : NN:='BIL' ; 2 : NN:='CHA' ; 3 : NN:='PRO' ; 4 : NN:='EXT' ; End ;
    for i:=1 to 5 do if Not Manque then if ((i<3) or (NN<>'EXT')) then
        BEGIN
        Nam1:='SO_'+NN+'DEB'+IntToStr(i) ; Nam2:='SO_'+NN+'FIN'+IntToStr(i) ;
        if ((TT.GetValue(Nam1)='') and (TT.GetValue(Nam2)<>'')) then Manque:=True ;
        if ((TT.GetValue(Nam1)<>'') and (TT.GetValue(Nam2)='')) then Manque:=True ;
        END ;
    END ;
Result:=Manque ;
if Result then HShowMessage('10;Société;La fourchette de comptes que vous avez renseignée n''est pas valide : les deux comptes doivent être renseignés;W;O;O;O;','','') ;
END ;

Function CompareTabloCpteOk ( CC : TControl ; NN1,NN2 : String ) : Boolean ;
Var i,j : Integer ;
    TT : TOB ;
    ND1,NF1 : String ;
BEGIN
TT:=GetTOBCpts(CC) ;
Result:=False ; i:=1 ;
While i<5 do
 BEGIN
 ND1:='SO_'+NN1+'DEB'+IntToStr(i) ; NF1:='SO_'+NN1+'FIN'+IntToStr(i) ;
 if (TT.GetValue(ND1)<>'') and (TT.GetValue(NF1)<>'') then
    BEGIN
    j:=1 ;
    While j<5 do
      BEGIN
      if (TT.GetValue('SO_'+NN2+'DEB'+IntToStr(i))<>'') and (TT.GetValue('SO_'+NN2+'FIN'+IntToStr(i))<>'') then
         BEGIN
         if TT.GetValue('SO_'+NN2+'DEB'+IntToStr(j))=TT.GetValue('SO_'+NN1+'DEB'+IntToStr(i)) then Exit ;
         if TT.GetValue('SO_'+NN2+'FIN'+IntToStr(j))=TT.GetValue('SO_'+NN1+'FIN'+IntToStr(i)) then Exit ;
         if TT.GetValue('SO_'+NN2+'DEB'+IntToStr(j))<TT.GetValue('SO_'+NN1+'DEB'+IntToStr(i)) then
            if TT.GetValue('SO_'+NN2+'FIN'+IntToStr(j))>TT.GetValue('SO_'+NN1+'DEB'+IntToStr(i)) then Exit ;
         if TT.GetValue('SO_'+NN2+'DEB'+IntToStr(j))>TT.GetValue('SO_'+NN1+'DEB'+IntToStr(i)) then
            if TT.GetValue('SO_'+NN2+'DEB'+IntToStr(j))<TT.GetValue('SO_'+NN1+'FIN'+IntToStr(i)) then Exit ;
         END ;
      inc(j) ;
      END ;
   END ;
 Inc(i) ;
 END ;
Result:=true ;
END ;

Function ChevauchementCpteFourchetteOk ( CC : TControl ) : Boolean ;
Var okok : boolean ;
BEGIN
OkOk:=True ;
if Not CompareTabloCpteOk(CC,'BIL','PRO') then Okok:=False ;
if Not CompareTabloCpteOk(CC,'BIL','CHA') then Okok:=False ;
if Not CompareTabloCpteOk(CC,'BIL','EXT') then Okok:=False ;
if Not CompareTabloCpteOk(CC,'PRO','CHA') then Okok:=False ;
if Not CompareTabloCpteOk(CC,'PRO','EXT') then Okok:=False ;
if Not CompareTabloCpteOk(CC,'CHA','EXT') then Okok:=False ;
Result:=OkOk ;
if Not OkOk then HShowMessage('11;Société;La fourchette de comptes que vous avez renseignée n''est pas valide : cette fourchette a déjà été choisie;W;O;O;O;','','') ;
END ;

Function VerifiNatureCpteOk ( CC : TControl ; NN : String ) : Boolean ;
Var Q : TQuery ;
    i : Integer ;
    PasOk : Boolean ;
    TT  : TOB ;
    {Nat,}SQL : String ;
    C1,C2 : String ;
BEGIN
Result:=True ; PasOk:=False ; i:=1 ; TT:=GetTOBCpts(CC) ;
While (i<5) and ((i<3) or (NN<>'EXT')) do
   BEGIN
   C1:=TT.GetValue('SO_'+NN+'DEB'+IntToStr(i)) ; C2:=TT.GetValue('SO_'+NN+'FIN'+IntToStr(i)) ;
   if C1<>'' then
      BEGIN
      SQL:='SLECT G_NATUREGENE,G_GENERAL FROM GENERAUX WHERE G_GENERAL>="'+C1+'" AND G_GENERAL<="'+C2+'"' ;
      if NN='BIL' then SQL:=SQL+' AND (G_NATUREGENE="CHA" OR G_NATUREGENE="PRO" OR G_NATUREGENE="EXT")' else
       if NN='CHA' then SQL:=SQL+' AND G_NATUREGENE<>"CHA"' else
        if NN='PRO' then SQL:=SQL+' AND G_NATUREGENE<>"PRO"' else
         if NN='EXT' then SQL:=SQL+' AND G_NATUREGENE<>"EXT"' ;
      Q:=OpenSql(SQL,True) ; if Not Q.EOF then PasOk:=True ; Ferme(Q) ;
      END ;
   Inc(i) ; if PasOk then Break ;
   END ;
if PasOk then
   BEGIN
   HShowMessage('12;Société;La fourchette de comptes que vous avez renseignée n''est pas en accord avec la nature des comptes généraux;W;O;O;O;','','') ;
   Result:=False ;
   END ;
END ;

Function NatureCpteOk ( CC : TControl ) : Boolean ;
BEGIN
Result:=False ;
if Not VerifiNatureCpteOk(CC,'BIL') then Exit ;
if Not VerifiNatureCpteOk(CC,'CHA') then Exit ;
if Not VerifiNatureCpteOk(CC,'PRO') then Exit ;
if Not VerifiNatureCpteOk(CC,'EXT') then Exit ;
Result:=True ;
END ;

Function VerifFourchettes ( CC : TControl ) : boolean ;
Var FF : TForm ;
BEGIN
// Test de la page
Result:=True ; FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_BILDEB1',FF) then Exit ;
// Test de la modif
ChargeFourchettes(CC,False) ;
if Not GetTOBCpts(CC).Modifie then Exit ;
// Traitement
Result:=False ;
if Not FourchetteOk(CC) then Exit ;
if ManqueUnCpte(CC) then Exit ;
if Not ChevauchementCpteFourchetteOk(CC) then Exit ;
if Not NatureCpteOk(CC) then Exit ;
Result:=True ;
END ;

Function GenExiste ( Nam : String ; FF : TForm ) : boolean ;
Var CC : THCritMaskEdit ;
BEGIN
Result:=True ;
CC:=THCritMaskEdit(GetFromNam(Nam,FF)) ;
if CC.Text<>'' then if Not Presence('GENERAUX','G_GENERAL',CC.Text) then
   BEGIN
   HShowMessage('0;Société;Le compte général n''est pas valide;E;O;O;O;','','') ;
   Result:=False ; if CC.CanFocus then CC.SetFocus ;
   END ;
END ;

Function AuxExiste ( Nam : String ; FF : TForm ) : boolean ;
Var CC : THCritMaskEdit ;
BEGIN
Result:=True ;
CC:=THCritMaskEdit(GetFromNam(Nam,FF)) ;
if CC.Text<>'' then if Not Presence('TIERS','T_AUXILIAIRE',CC.Text) then
   BEGIN
   HShowMessage('0;Société;Le compte auxiliaire n''est pas valide;E;O;O;O;','','') ;
   Result:=False ; if CC.CanFocus then CC.SetFocus ;
   END ;
END ;

Function VerifExisteCpts ( CC : TControl ) : boolean ;
Var FF : TForm ;
BEGIN
Result:=False ;
FF:=GetLaForm(CC) ;
if ExisteNam('SO_DEFCOLCLI',FF) then
   BEGIN
   if Not GenExiste('SO_DEFCOLCLI',FF) then Exit ;
   if Not GenExiste('SO_DEFCOLFOU',FF) then Exit ;
   if Not GenExiste('SO_DEFCOLSAL',FF) then Exit ;
   if Not GenExiste('SO_DEFCOLDDIV',FF) then Exit ;
   if Not GenExiste('SO_DEFCOLCDIV',FF) then Exit ;
   if Not GenExiste('SO_DEFCOLDIV',FF) then Exit ;
   END ;
if ExisteNam('SO_GENATTEND',FF) then
   BEGIN
   if Not GenExiste('SO_GENATTEND',FF) then Exit ;
   if Not GenExiste('SO_OUVREBIL',FF) then Exit ;
   if Not GenExiste('SO_FERMEBIL',FF) then Exit ;
   if Not GenExiste('SO_OUVREPERTE',FF) then Exit ;
   if Not GenExiste('SO_FERMEPERTE',FF) then Exit ;
   if Not GenExiste('SO_OUVREBEN',FF) then Exit ;
   if Not GenExiste('SO_FERMEBEN',FF) then Exit ;
   if Not GenExiste('SO_RESULTAT',FF) then Exit ;

   if Not AuxExiste('SO_CLIATTEND',FF) then Exit ;
   if Not AuxExiste('SO_FOUATTEND',FF) then Exit ;
   if Not AuxExiste('SO_SALATTEND',FF) then Exit ;
   if Not AuxExiste('SO_DIVATTEND',FF) then Exit ;
   END ;
if ExisteNam('SO_ECCEURODEBIT',FF) then
   BEGIN
   if Not GenExiste('SO_ECCEURODEBIT',FF) then Exit ;
   if Not GenExiste('SO_ECCEUROCREDIT',FF) then Exit ;
   END ;
Result:=True ;
END ;

{$ifdef AFFAIRE}
Function VerifAfDates ( CC : TControl ) : boolean ;
Var  FF : TForm ;
    DateDeb ,DateFIn : THCritMaskEdit ;
BEGIN
//tEST DATE fIN > DATE D2BUT
Result:=False ;
FF:=GetLaForm(CC) ;
if ExisteNam('SO_AFDATEDEB35',FF) then
begin
DateDeb:=THCritMaskEdit(GetFromNam('SO_AFDATEDEB35',FF)) ;
DateFIn:=THCritMaskEdit(GetFromNam('SO_AFDATEFIN35',FF)) ;

if (STrToDATe(datefin.text) < StrToDate(datedeb.text)) then exit;
if (ctxScot in V_PGI.PGIContexte) then
   BEGIN
    DateDeb:=THCritMaskEdit(GetFromNam('SO_AFDATEDEBCAB',FF)) ;
    DateFIn:=THCritMaskEdit(GetFromNam('SO_AFDATEFINCAB',FF)) ;

    if (STrToDATe(datefin.text) < StrToDate(datedeb.text)) then exit;
   end;

// Pl le 21/11/02 : gestion de la premiere semaine de l'annee
if (V_PGI<>nil) then V_PGI.Semaine53 := GetParamsoc('SO_PREMIERESEMAINE');
end;

Result:=True ;
END ;

Function VerifAfApprec ( CC : TControl ) : boolean ;
Var  FF : TForm ;
     Boni : TCheckBox ;
     SQL : string;
BEGIN
Result:=True ;
FF:=GetLaForm(CC) ;
if ExisteNam('SO_AFGESTIONAPPRECIATION',FF) then
   begin
   Boni:=  TCheckBox(GetFromNam('SO_AFAPPAVECBM',FF));
   if  Boni.checked = true then
      begin
      if ExisteSql('SELECT APB_NUMBONI FROM PARBONI') then exit ; // mcd 31/10/02 suppression select  *
         // il n'y a rien dans les paramètre boni, on en crée 2 par défaut
      CreerArticlesBoni;
      CreerRessBoni;
      Sql :='INSERT INTO PARBONI' +
                    '(APB_NUMBONI,APB_NUMLIG,APB_LIBELLE,APB_SENS, APB_TYPEARTICLE, APB_CODEARTICLE, APB_ARTICLE,'
                    +   ' APB_RESSOURCE, APB_PRIRESS, APB_PRORATA, APB_PRIPRO, APB_PRIRESSAFF, APB_RESSAFF)'
                    +'VALUES (1,1,"Boni","+", "PRE", "PRBONI", "'
                    +   CodeArticleUnique2('PRBONI','')+'" ,"ASSBONI", 3,"X", 1, 2, "AFF_RESPONSABLE")' ;
      Executesql(sql);
      Sql :='INSERT INTO PARBONI' +
          '(APB_NUMBONI,APB_NUMLIG,APB_LIBELLE,APB_SENS, APB_TYPEARTICLE, APB_CODEARTICLE,APB_ARTICLE, '
                    +   ' APB_RESSOURCE, APB_PRIRESS, APB_PRORATA, APB_PRIPRO, APB_PRIRESSAFF, APB_RESSAFF)'
                    +'VALUES (2,2,"Mali","-", "PRE", "PRMALI", "'
                    +   CodeArticleUnique2('PRMALI','')+'" ,"ASSBONI", 3,"X", 1, 2, "AFF_RESPONSABLE")' ;
      Executesql(sql);
      end;
   end;
END ;
{$endif}

{$IFDEF GCGC}
Function VerifExisteCptsGC ( CC : TControl ) : boolean ;
Var FF : TForm ;
BEGIN
Result:=False ;
FF:=GetLaForm(CC) ;
if ExisteNam('SO_GCCPTEESCACH',FF) then
   BEGIN
   if Not GenExiste('SO_GCCPTEESCACH',FF) then Exit ;
   if Not GenExiste('SO_GCCPTEESCVTE',FF) then Exit ;
   if Not GenExiste('SO_GCCPTEREMACH',FF) then Exit ;
   if Not GenExiste('SO_GCCPTEREMVTE',FF) then Exit ;
   if Not GenExiste('SO_GCCPTEHTACH',FF) then Exit ;
   if Not GenExiste('SO_GCCPTEHTVTE',FF) then Exit ;
   END ;
Result:=True ;
END ;

{$ENDIF}

Function SauvePageSocGA ( CC : TControl ) : boolean ;
BEGIN
Result:=False ;
if Not VerifCoord(CC) then Exit ;
if Not LongueurCpteOk(CC) then Exit ;
if Not GeneAttenteOk(CC)  then Exit ;
if Not ExisteCptesAttente(CC) then Exit ;
if Not VerifFourchettes(CC) then Exit ;
if Not DefautCpteOk(CC) then Exit ;
if Not VerifDateEntreeEuroOk(CC) then Exit ;
if not VerifPariteEuro(CC) then Exit ;
if not VerifCptesEuro(CC) then Exit ;
if Not CarBourreOk(CC) then Exit ;
//if Not EtabDefautOk(CC) Then Exit ;
if Not VerifJalTP(CC) Then Exit ;
if Not VerifExisteCpts(CC) then Exit ;
{$IFDEF GCGC}
if Not VerifExisteCptsGC(CC) then Exit ;
{$ENDIF}
{$IFDEF AFFAIRE}
if Not VerifAFDAtes(CC) then Exit ;
if Not VerifAFApprec(CC) then Exit ;
{$ENDIF}
// PL le 21/01/02 : pour blocage appréciation si monnaie de tenue du dossier est le franc
// mcd 12/02/02 if Not VerifApprec(CC) Then Exit ;
// Fin PL le 21/01/02
MarquerPublifi(True) ;
Result:=True ;
END ;

Function SauvePageSocSansVerif ( CC : TControl ) : boolean ;
BEGIN
Result:=False ;
if Not VerifCoord(CC) then Exit ;
if Not LongueurCpteOk(CC) then Exit ;
if Not GeneAttenteOk(CC)  then Exit ;
//if Not ExisteCptesAttente(CC) then Exit ;
if Not VerifFourchettes(CC) then Exit ;
if Not DefautCpteOk(CC) then Exit ;
//if Not VerifDateEntreeEuroOk(CC) then Exit ;
//if not VerifPariteEuro(CC) then Exit ;
//if not VerifCptesEuro(CC) then Exit ;
if Not CarBourreOk(CC) then Exit ;
//if Not EtabDefautOk(CC) Then Exit ;
//if Not VerifJalTP(CC) Then Exit ;
//if Not VerifExisteCpts(CC) then Exit ;
{$IFDEF GCGC}
// if Not VerifExisteCptsGC(CC) then Exit ;
{$ENDIF}
MarquerPublifi(True) ;
Result:=True ;
END ;

Procedure MarquerPublifi ( Flag : boolean ) ;
BEGIN
if Not VH^.LienPublifi then Exit ;
{$IFNDEF SPEC302}
if Flag then SetParamSoc('SO_ACHARGERPUBLIFI','X') else SetParamSoc('SO_ACHARGERPUBLIFI','-') ;
{$ENDIF}
END ;

end.
