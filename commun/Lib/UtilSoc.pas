unit UtilSoc;
 
interface

uses
  Forms,
  SysUtils,
  Controls,
  buttons,
  Classes,
  ComCtrls,
  Graphics,
  HCtrls,
  stdctrls,
  Grids,
  ExtCtrls,
  Windows,
  HEnt1,
  registry,
  Menus,
  {$IFDEF EAGLCLIENT}
    Maineagl,
    UtileAGL,
  {$ELSE EAGLCLIENT}
    MajTable,
    {$IFNDEF EAGLSERVER}
      FE_Main,
    {$ENDIF EAGLSERVER}
    DBGrids,
    DBCtrls,
    DB,
    {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
    HDB,
  {$ENDIF EAGLCLIENT}
  {$IFDEF PGIIMMO}
    uImoParam,
  {$ENDIF PGIIMMO}
  {$IFDEF AMORTISSEMENT}
    ImPlan,
  {$ENDIF AMORTISSEMENT}
  {$IFDEF AFFAIRE}
{$IFNDEF EAGLSERVER}
    TraducAffaire,
{$ENDIF EAGLSERVER}
    UtilArticle,
    DicoBTP,
    {$IFDEF GIGI}
      galoutil,
    {$ENDIF GIGI}
  {$ENDIF AFFAIRE}
  {$IFDEF GCGC}
    EntGC,
  {$ENDIF GCGC}
  WinProcs,
  shellAPI,
  Hstatus,
  LicUtil,
  HMsgBox,
  UtilPGI,
  Messages,
  Hrichedt,
  HTB97,
  Dialogs,
{$IFNDEF EAGLSERVER}
  LookUp,
{$ENDIF EAGLSERVER}
  UTOB,uEntCommun;

{$IFNDEF EAGLSERVER}
Function  InterfaceSoc ( CC : TControl ) : boolean ;
Function  ChargePageSoc ( CC : TControl ) : boolean ;
Function  SauvePageSoc ( CC : TControl ) : boolean ;
Function  SauvePageSocSansverif ( CC : TControl ) : boolean ;
Procedure MarquerPublifi ( Flag : boolean ) ;
procedure planningIntegre(FF: TForm);
{$ENDIF EAGLSERVER}
Procedure TripoteStatus ( StAjout : String = '')  ;
{$IFNDEF EAGLSERVER}
function  ChargePageImmo (CC : TControl) : boolean;
function  SauvePageImmo (CC : TControl) : boolean;
procedure SetLenMaxCompteImmo(CC : TControl; laListe : TStringList ; Len : integer);
Procedure SetLenMaxSt ( Nam : String ; FF : TForm ; Len : integer ; bMajText : boolean = false ; cb : string = '') ;
Procedure AlimZonesProjet ( CC : TControl ) ;
Procedure GereZonesVersions ( CC : TControl ) ;
Procedure FormatZoneCli ( CC : TControl ) ;
Procedure ModifDevPrinc ( CC : TControl ) ;
Function SocTox(CC : TControl) : boolean ;
Function LgNumProjetOk ( CC : TControl ) : boolean;
Function LgNumOperOk ( CC : TControl ) : boolean;
Function GestionMonoDepot( CC : TControl ) : boolean;
Procedure MajGPPEcheances( CC : TControl );
Procedure RTTabHie(CC : TControl;stName : String) ;
Procedure VerifTabHie ( CC : TControl ) ;
Procedure TraiteTabHie ( CC : TControl; stName : string );
Procedure CtrlNbDecimales(CC : TControl; stName : String) ;
procedure BornesVisibles(CC: TControl);
function VerifBorne(CC: TControl; NameMin, NameMax, Quoi: string): boolean;
procedure GereZonesMarche(CC: TControl; Flux: string);
Procedure GereParamSocPieceOrigine(FF: TForm);
Procedure InitGPOFraisAvances( CC : TControl ) ;
{$IFDEF GPAO}
  {$IFNDEF EAGLSERVER}
    Procedure InitWLBContexteOrigine( CC : TControl ) ;
  {$ENDIF EAGLSERVER}
  {$IFDEF AFFAIRE}
    {$IFNDEF PGIMAJVER}
      Procedure InitParamAffairesIndustrielles( CC : TControl );
    {$ENDIF PGIMAJVER}
  {$ENDIF AFFAIRE}
{$ENDIF GPAO}
Procedure VerifGestLienOperPiece ( CC : TControl ) ;
procedure MajTabletteGcTypeFourn( CC : TControl );
procedure MajParamCdeMarche( CC : TControl);
procedure MajClassificationTMDR(CC : TControl); // CCMX-CEGID TMDR Dev N°4722
Function VerifModeleWord ( CC : TControl ) : boolean;
procedure VerifGestYplanningGRC ( CC : TControl );
{$ENDIF EAGLSERVER}
{$ifdef AFFAIRE}
Function ExisteProfilGener ( Profil : String ) : boolean ;
Function ExisteArticle ( Art : String ) : boolean ;
Function ExisteEtabliss ( Eta : String ) : boolean ;
Function ExisteUnite ( Uni : String ) : boolean ;
Function ExisteAffaires : boolean ;
{$endif AFFAIRE}
{$IFDEF GPAOLIGHT}
  {$IFNDEF EAGLSERVER}
	procedure InitParamSocGPAO(FF: TForm);
  {$ENDIF EAGLSERVER}
{$ENDIF GPAOLIGHT}

// GCO - 22/04/2005 - Utilisé dans les projets CCS5 et CCSTD
function  CPEstLiasseModifiee : Boolean;

{$IFNDEF EAGLSERVER}
{$IFDEF GESCOM}
  procedure CdeOuvCacheFonctions(FF: TForm);
  procedure CdeOuvTraiteAutoTransfo(CC: TControl);
{$ENDIF GESCOM}

{$IFDEF GCGC}
  procedure SetPlusUniteParDefaut(CC: TControl);
  procedure SetGCGestUniteModeState(CC: TControl);
  procedure GCChangeGestUniteMode(CC: TControl);
  procedure SetGereEnseigne (CC : TControl);
  procedure SetGereCiblage (CC : TControl);
  procedure SetGereOperation (CC : TControl);
//GP_BUG810_TP_WPLANLIVR_20071008
  procedure SetWPlanLivr(CC: TControl);
{$ENDIF GCGC}

  Function VerifAfActivite ( CC : TControl ) : boolean ;
{$ENDIF EAGLSERVER}
{$IFDEF GIGI}
  //PCH 11-2005 personnalisation paramsoc
  Function VerifAfPerso ( CC : TControl ) : boolean ;
  Procedure CtlPersoParamsocChoix( CC : TControl ) ;
  Procedure CtlPersoParamsocPar1( CC : TControl ) ;
  Procedure CtlPersoParamsocPar2( CC : TControl ) ;
  Function DecochagePersoParamsocChoix( CC : TControl ) : boolean ;
  // PCH 12/12/2006 Gestion personnalisation FACTURE LIQUIDATIVE
  Procedure CtlFactureLiquidative(CC : TControl);
{$ENDIF GIGI}

// ajout me
{$IFNDEF EAGLSERVER}
function VerifExistSynchro ( CC : TControl ) : Boolean;
procedure VerifPdrParametrage(CC:TControl);
procedure GereMasqueCrit    ( FF : TForm );
Function GenExiste ( Nam : String ; FF : TForm; Vide : boolean = false;  Libelle : string = '') : boolean ;

  {$IFDEF COMPTA}
  {$IFNDEF CCMP}
  function TraitePlanDeRevision ( CC : TControl ) : Boolean;
  {$ENDIF}
  {$ENDIF}

{$IFDEF STK}
//  BBI web service
procedure CtrlDecochageGestionEmplacement(CC : TControl);
procedure CtrlDecochageGestionLot(CC : TControl);
procedure CtrlDecochageGestionNoSerie(CC : TControl);
procedure CtrlDecochageGestionNoSerieGrp(CC : TControl);
procedure CtrlDecochageGestionStatutDispo(CC : TControl);
procedure CtrlDecochageGestionStatutFlux(CC : TControl);
procedure CtrlDecochageGestionMarque(CC : TControl);
procedure CtrlDecochageGestionChoixQualite(CC : TControl);
{$ENDIF STK}
//  BBI Fin web service
  procedure GereSoldePiecePrec (CC : TControl);
{$ENDIF !EAGLSERVER}
function IsMasterOnShare (elementpartage : String) : boolean;

procedure CryptePassword(var   Valeur : String);
procedure DeCryptePassword(var Valeur : String);
//pocedure ChkProtectOnExit(Sender: TObject);

implementation

Uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  CPVersion,  
  {$ENDIF MODENT1}
 {$IFDEF COMPTA}
{$IFNDEF CMPGIS35}
  cloture ,
  OuvreExo ,
{$ENDIF}
  {$IFNDEF CCMP}
{$IFNDEF CMPGIS35}
  uLibRevision,
{$ENDIF}
  {$ENDIF}
  {$IFDEF EAGLCLIENT}
  AnulCloServeur ,
  {$ELSE}
{$IFNDEF CMPGIS35}
  AnulClo ,
{$ENDIF}
  {$ENDIF}
  {$ENDIF} 
  Ent1,
  SaisUtil,
  { GC_DBR_CRM10764 }
  Math,
  {$IFNDEF SANSCOMPTA}
    {$IFNDEF PGIIMMO}
      HCompte,
    {$ENDIF PGIIMMO}
  {$ENDIF SANSCOMPTA}
  {$IFDEF STK}
    {$IFNDEF PGIMAJVER}
      StockUtil,
    {$ENDIF PGIMAJVER}
  {$ENDIF STK}
  {$IFDEF ACCESSCM}
    ControleParamSoc,
  {$ENDIF ACCESSCM}
  {$IFDEF GESCOM}
    FactPieceContainer,
    CdeOuv,
  {$ENDIF GESCOM}
  {$IFDEF CHR}
  HRUtils,
  {$ENDIF CHR}
  ParamSoc,
  SaisComm,
  Spin
  {$IFDEF GRC}
  {$IFNDEF GIGI}
{$IFNDEF EAGLSERVER}
  ,AfterImportGEDGRC
{$ENDIF EAGLSERVER}
  {$ENDIF GIGI}
{$IFNDEF EAGLSERVER}
  ,UGedFiles
{$ENDIF EAGLSERVER}
  {$ENDIF GRC}
  {$IFDEF GCGC}
    {$IFNDEF PGIMAJVER}
      {$IFNDEF CMPGIS35}
//      ,yTarifs
      {$ENDIF !CMPGIS35}
    {$ENDIF !PGIMAJVER}
//    ,yTarifsCommun
  {$ENDIF GCGC}
  {$IFNDEF PGIIMMO}
  ,wCommuns
  {$ENDIF PGIIMMO}
  {$IFDEF GPAOLIGHT}
  	{$IFNDEF GPAO}
      {$IFNDEF EAGLSERVER}
    	  ,GpLightSeria
      {$ENDIF EAGLSERVER}
  	{$ELSE !GPAO}
      {$IFDEF AFFAIRE}
        {$IFNDEF PGIMAJVER}
          ,wAffaires
        {$ENDIF !PGIMAJVER}
      {$ENDIF  AFFAIRE}
  	{$ENDIF !GPAO}
    ,GcStkValo
//GP_20080204_TS_GP14791
    ,EntGP
  {$ENDIF GPAOLIGHT}
  //GP_20080107_TS_GC15692 >>>
  {$IFDEF EDI}
    ,EntEDI
  {$ENDIF EDI}
  ,CbpParamsoc
  //GP_20080107_TS_GC15692 <<<
  ,UFonctionsCBP
  ;

{$IFNDEF EAGLSERVER}
Var
  DevisePrinc : String [3];
//  ChkProtect  : TCheckBox;
  {$IFDEF STK}
    RefL : thValComboBox;
  {$ENDIF STK}

{========================= Fonctions utilitaires ======================}
Function GetLaTOB ( CC : TControl ) : TOB ;
BEGIN Result:=TFParamSoc(CC.Owner).LaTOB ; END ;

Function GetTOBSoc ( CC : TControl ) : TOB ;
BEGIN Result:=GetLaTOB(CC).Detail[0] ; END ;

Function GetTOBCpts ( CC : TControl ) : TOB ;
BEGIN Result:=GetLaTOB(CC).Detail[1] ; END ;

Function GetLaForm ( CC : TControl ) : TForm ;
BEGIN Result:=TForm(CC.OWner) ; END ;

Function ExisteNam ( Nam : String ; FF : TForm ) : Boolean ;
BEGIN Result:=(FF.FindComponent(Nam)<>Nil) ; END ;

Function GetFromNam ( Nam : String ; FF : TForm ) : TControl ;
BEGIN
Result:=TControl(FF.FindComponent(Nam)) ;
if ((V_PGI.SAV) and (Result=Nil)) then ShowMessage('Pas trouvé '+Nam) ;
END ;

Procedure SetInvi ( Nam : String ; FF : TForm ) ;
Var CC : TControl ;
BEGIN
CC:=GetFromNam(Nam,FF) ; if CC<>Nil then CC.Visible:=False ;
END ;

Procedure SetEna (Const Nam: String; Const FF: TForm; Const Ena : boolean);
Var
  CC : TControl ;
BEGIN
  CC := GetFromNam(Nam, FF);
  if Assigned(CC) then
    CC.Enabled := Ena;
END ;

Procedure SetVisi ( Nam : String ; FF : TForm ) ;
Var CC : TControl ;
BEGIN
CC:=GetFromNam(Nam,FF) ; if CC<>Nil then CC.Visible:=True ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
// BDU - 30/01/07 - Contrôle de saisie des heures d'affichage du planning
procedure ControleHeure(Champ: String; Fenetre: TForm);
var
  H1,
  H2: THCritMaskEdit;
begin
  if ExisteNam('SO_AFHEUREDEBUTJOUR', Fenetre) and (Champ = 'SO_AFHEUREDEBUTJOUR') and
    ExisteNam('SO_AFAMDEBUT', Fenetre) then
  begin
    H1 := THCritMaskEdit(GetFromNam('SO_AFAMDEBUT', Fenetre));
    H2 := THCritMaskEdit(GetFromNam(Champ, Fenetre));

    //C.B 03/04/2007 l'heure ne peut aller que jusqu'a 23h
    //if H2.Text = '00:00' then
    //  H2.Text := '23:00';

    if H2.Text > H1.Text then
      H2.Text := H1.Text;
  end

  else if ExisteNam('SO_AFHEUREFINJOUR', Fenetre) and (Champ = 'SO_AFHEUREFINJOUR') and
    ExisteNam('SO_AFPMFIN', Fenetre) then
  begin
    H1 := THCritMaskEdit(GetFromNam('SO_AFPMFIN', Fenetre));
    H2 := THCritMaskEdit(GetFromNam(Champ, Fenetre));

    //C.B 03/04/2007 l'heure ne peut aller que jusqu'a 23h
    //if H2.Text = '00:00' then
    //  H2.Text := '23:00';

    if (H2.Text < H1.Text) and (H2.Text <> '00:00') then
      H2.Text := H1.Text;
  end;

  {else if ExisteNam('SO_AFPMFIN', Fenetre) and (Champ = 'SO_AFPMFIN') then
  begin
    H2 := THCritMaskEdit(GetFromNam(Champ, Fenetre));
    if H2.Text = '00:00' then
      H2.Text := '23:00';
  end;}
end; 
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
// BDU - 29/11/07. Nouvelle exportation SDA
procedure GestionExportSDA(Fenetre: TForm);
var
  L: THValComboBox;
  CB1,
  CB2: THCheckBox;
begin
  L := THValComboBox(GetFromNam('SO_AFEXPORTSDAAFFAIRE', Fenetre));
  CB1 := ThCheckBox(GetFromNam('SO_AFEACTUNPOURTOUS', Fenetre));
  CB2 := ThCheckBox(GetFromNam('SO_AFEACTSENDMAIL', Fenetre));
  if CompareText(L.Value, 'MAN') <> 0 then
    CB1.Checked := False;
  if CB1.Checked then
    CB2.Checked := False;
  SetEna('SO_AFEACTUNPOURTOUS', Fenetre, L.Value = 'MAN');
  SetEna('SO_AFEACTSENDMAIL', Fenetre, not CB1.Checked);
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
// BDU - 24/11/06 - 12887, Visa automatique
procedure GestionVisaAutomatique(Fenetre: TForm);
var
  Cb1,
  Cb2: ThCheckBox;
  Lst: THValComboBox;
begin
  Cb1 := ThCheckBox(GetFromNam('SO_AFVISAACTIVITE', Fenetre));
  Cb2 := ThCheckBox(GetFromNam('SO_AFVISAAUTO', Fenetre));
  Lst := THValComboBox(GetFromNam('SO_AFTYPEACTVISA', Fenetre));
  if not Cb1.Checked then
    Cb2.Checked := False;
  if not Cb2.Checked then
    Lst.Text := '';
  SetEna('SO_AFVISAAUTO', Fenetre, Cb1.Checked);
  SetEna('SO_AFTYPEACTVISA', Fenetre, Cb2.Checked);
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{***********A.G.L.***********************************************
Auteur  ...... : BDU
Créé le ...... : 12/12/2006
Modifié le ... : 12/12/2006
Description .. : Nouvelles possibilités d'affichage du planning
Mots clefs ... :
*****************************************************************}
procedure GestionAffichePlanning(Fenetre: TForm);
var
  Cb: ThCheckBox;
  ListeStk,
  Liste: THValComboBox;
  Valeur: String;
  X: Integer;
begin
  Liste := THValComboBox(GetFromNam('SO_AFPLANNINGDEFAUT', Fenetre));
  if Assigned(Liste) then
  begin
    // On conserve la valeur actuelle pour la réaffecter plus tard dans cette procédure
    Valeur := Liste.Values[Liste.ItemIndex];
    // La liste est vidée, elle sera remplie en fonction des cases cochées
    Liste.Items.Clear;
    Liste.Values.Clear;
    // Dans cette liste, nous avons l'ensemble des valeurs possibles
    ListeStk := THValComboBox(GetFromNam('SO_AFPLANNINGDEFAUTSTK', Fenetre));
    if Assigned(ListeStk) then
    begin
      // Le planning à la demi journée est-il coché ?
      Cb := ThCheckBox(GetFromNam('SO_AFDEMIJOURNEE', Fenetre));
      if Assigned(Cb) and Cb.Checked then
      begin
        Liste.Items.Add(ListeStk.Items[ListeStk.Values.IndexOf('003')]);
        Liste.Values.Add(ListeStk.Values[ListeStk.Values.IndexOf('003')]);
      end;
      // Le planning à l'heure est-il coché ?
      Cb := ThCheckBox(GetFromNam('SO_AFPLANNINGHEURE', Fenetre));
      if Assigned(Cb) and Cb.Checked then
      begin
        Liste.Items.Add(ListeStk.Items[ListeStk.Values.IndexOf('004')]);
        Liste.Values.Add(ListeStk.Values[ListeStk.Values.IndexOf('004')]);
      end;
      // Le planning au mois est-il coché ?
      Cb := ThCheckBox(GetFromNam('SO_AFPLANNINGMOIS', Fenetre));
      if Assigned(Cb) and Cb.Checked then
      begin
        Liste.Items.Add(ListeStk.Items[ListeStk.Values.IndexOf('001')]);
        Liste.Values.Add(ListeStk.Values[ListeStk.Values.IndexOf('001')]);
      end;
      // Le planning à la demi-heure est-il coché
      Cb := ThCheckBox(GetFromNam('SO_AFPLANNINGDEMIHEURE', Fenetre));
      if Assigned(Cb) and Cb.Checked then
      begin
        Liste.Items.Add(ListeStk.Items[ListeStk.Values.IndexOf('005')]);
        Liste.Values.Add(ListeStk.Values[ListeStk.Values.IndexOf('005')]);
      end;
      // Le planning au jour est-il coché ?
      Cb := ThCheckBox(GetFromNam('SO_AFPLANNINGJOUR', Fenetre));
      // Si rien n'est encore coché, le planning jour est automatiquement coché par défaut
      if Assigned(Cb) and (Liste.Items.Count = 0) then
        Cb.Checked := True;
      if Assigned(Cb) and Cb.Checked then
      begin
        Liste.Items.Add(ListeStk.Items[ListeStk.Values.IndexOf('002')]);
        Liste.Values.Add(ListeStk.Values[ListeStk.Values.IndexOf('002')]);
      end;
      // Recherche si l'ancienne valeur est toujours disponible
      X := Liste.Values.IndexOf(Valeur);
      if X <> -1 then
        // Re-sélectionne l'ancienne valeur
        Liste.ItemIndex := X
      else
        // On prend la première "par défaut"
        Liste.ItemIndex := 0;
    end;
  end;
{$ifdef AFFAIRE}
  TraduitForm(Fenetre); //mcd 02/01/2008
{$endif AFFAIRE}
end;

// C.B 22/05/2007
procedure GestionTous(Fenetre : TForm);
begin
  THMultiValComboBox(GetFromNam('SO_AFETATRES', Fenetre)).Complete := True;
  THMultiValComboBox(GetFromNam('SO_AFETATINTERDIT', Fenetre)).Complete := True;
  THMultiValComboBox(GetFromNam('SO_AFTYPERESSOURCE', Fenetre)).Complete := True;
end;                                          
{$ENDIF EAGLSERVER}

//debut fct spécfiique paramsco GI/GA
{$ifdef AFFAIRE}
Function ExisteArticle ( Art : String ) : boolean ;
BEGIN
Result:=False ; if Art='' then Exit ;
Result := ExisteSQL('SELECT GA_CODEARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="'+Art+'"' ) ;
END ;

Function ExisteEtabliss ( Eta : String ) : boolean ;
BEGIN
Result:=False ; if Eta='' then Exit ;
Result := ExisteSQL('SELECT ET_ETABLISSEMENT FROM ETABLISS WHERE ET_ETABLISSEMENT="'+Eta+'"' ) ;
END ;

Function ExisteProfilGener ( Profil : String ) : boolean ;
BEGIN
Result:=False ; if Profil='' then Exit ;
Result := ExisteSQL('SELECT APG_PROFILGENER FROM PROFILGENER WHERE APG_PROFILGENER="'+Profil+'"' ) ;
END ;

Function ExisteAffaires : boolean ;
BEGIN
  Result := ExisteSQL('SELECT AFF_AFFAIRE FROM AFFAIRE' ) ;
END ;

Function ExisteActivite : boolean ;
BEGIN
Result := ExisteSQL('SELECT ACT_TYPEACTIVITE FROM ACTIVITE');
END ;

Function ExisteAfCumul (TypeCumul: string): boolean ;
BEGIN
Result := ExisteSQL('SELECT ACU_TYPEAC FROM AFCUMUL WHERE ACU_TYPEAC="'+TypeCumul+'"' ) ;
END ;

Function ExisteUnite ( Uni : String ) : boolean ;
BEGIN
Result:=False ; if Uni='' then Exit ;
Result := ExisteSQL('SELECT GME_MESURE FROM MEA WHERE GME_MESURE="'+Uni+'"') ;
END ;

{$IFNDEF EAGLSERVER}
Procedure ConfigArticle( CC : TControl ) ;
Var FF   : TForm ;
BEGIN       //mcd 16/05/2005 pour cacher info complémentaire en GI
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_GCFAMILLETAXE1',FF) then Exit ;
If CtxScot in V_PGI.PGIContexte then
    begin
    SetInvi('LSCO_RTLARGEURFICHE004', FF);
    SetInvi('SCO_RTLARGEURFICHE004', FF);
    SetInvi('SO_RTLARGEURFICHE004', FF);
    SetInvi('LSO_RTLARGEURFICHE004', FF);
    SetInvi('SO_RTHAUTEURFICHE004', FF);
    SetInvi('LSO_RTHAUTEURFICHE004', FF);
    SetInvi('SO_GCFAMHIERARCHIQUE', FF);
    SetInvi('SO_RTGESTINFOS004', FF);
      //mcd 12/06/07 on cache les autres info non gérées en GI
    SetInvi('SCO_PROFILART', FF);
    SetInvi('LSCO_PROFILART', FF);
    SetInvi('SO_GCPROFILART', FF);
    SetInvi('LSO_GCPROFILART', FF);
    SetInvi('SO_GCPROFILARTAUTO', FF);
    SetInvi('SCO_GCUNITES', FF);
    SetInvi('LSCO_GCUNITES', FF);
    SetInvi('SO_GCDEFQUALIFUNITEGA', FF);
    SetInvi('LSO_GCDEFQUALIFUNITEGA', FF);
    SetInvi('SO_GCDEFUNITEGA', FF);
    SetInvi('LSO_GCDEFUNITEGA', FF);
    SetInvi('SCO_SUPPRESSIONARTICLE', FF);
    SetInvi('LSCO_SUPPRESSIONARTICLE', FF);
    SetInvi('SO_DELETECATALOGASSOCIE', FF);
    SetInvi('SCO_GCDUPNOMEN', FF);
    SetInvi('LSCO_GCDUPNOMEN', FF);
    SetInvi('SO_GCDUPNOMENTYPE', FF);
    SetInvi('LSO_GCDUPNOMENTYPE', FF);
    SetInvi('SO_GCDUPNOMENCONF', FF);
    SetInvi('SO_GCCDUSEIT', FF);
    SetInvi('SCO_GCJPEGMEMO', FF);
    SetInvi('SO_GCCDJPG1', FF);
    SetInvi('LSO_GCCDJPG1', FF);
    SetInvi('SO_GCCDJPG2', FF);
    SetInvi('LSO_GCCDJPG2', FF);
    SetInvi('SO_GCCDJPG3', FF);
    SetInvi('LSO_GCCDJPG3', FF);
    SetInvi('SO_GCCDJPG4', FF);
    SetInvi('LSO_GCCDJPG4', FF);
    SetInvi('LSO_GCCDMEM1', FF);
    SetInvi('SO_GCCDMEM1', FF);
    SetInvi('LSO_GCCDMEM2', FF);
    SetInvi('SO_GCCDMEM2', FF);
    SetInvi('LSO_GCCDMEM3', FF);
    SetInvi('SO_GCCDMEM3', FF);
    SetInvi('LSO_GCCDMEM4', FF);
    SetInvi('SO_GCCDMEM4', FF);
    end;
END;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure ConfigSaisieDecen( CC : TControl );
Var FF : TForm;
BEGIN
  // test de la page
  FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFEACTPERIODIC',FF) then Exit ;
  If not(VH_GC.GASaisieDecSeria) and not(ctxGcAff in V_PGI.PGIContexte) then
  begin
    SetEna('SO_AFEACTPERIODIC', FF, False);
    SetEna('SO_AFEACTPATH', FF, False);
    SetEna('LSO_AFEACTPERIODIC', FF, False);
    SetEna('LSCO_GRPSAISIEDEC', FF, False);
    SetEna('LSO_AFEACTPATH', FF, False);
    SetEna('SO_AFEACTNEWAFF', FF, False);
    SetEna('SO_AFEACTUNPOURTOUS', FF, False);
    SetEna('SO_AFEACTSELART', FF, False);
  end;

  // Option pour intégrer des lignes d'eactivite avec suppression d'activite au prealable, non accessible en Scot
  // mais quand même visible avec le mot de passe du jour
  // PL le 01/08/02
  if (ctxScot in V_PGI.PGIContexte) then
  begin
    SetParamsoc('SO_AFEACTIMPSELEC',False);
    SetInvi('SO_AFEACTIMPSELEC',FF);  //import eactivite avec destruction pas utilisé en GI
    SetInvi('SCO_AFGRPEACT',FF);
    SetInvi('LSCO_AFGRPEACT',FF);

    // C.B 07/12/2004 on laisse le chemin excel
    // on laisse le pas de lien MsProject .. on cache
    SetInvi('SO_AFMSPPATH', FF);
    SetInvi('LSO_AFMSPPATH', FF) ;
    if (not VH_GC.GAPlanningSeria) then
      begin
      SetInvi('LSO_AFEXCELPATH', FF);
      SetInvi('SO_AFEXCELPATH', FF);
      SetInvi('LSCO_MSPROJECT', FF);
      SetInvi('SCO_MSPROJECT', FF);
      end;
  end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure ConfigSaisieBudget( CC : TControl ) ;
Var FF   : TForm ;
BEGIN
// test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFALIMECHE',FF) then Exit ;
SetInvi('SO_AFBUDZERO',FF) ;

end;
{$ENDIF EAGLSERVER}


{$IFNDEF EAGLSERVER}
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
    SetEna('SO_AFAPPANOUVEAU', FF, True);
    SetEna('SO_AFTRANSFERTENCOURS', FF, True);
   end
else begin
    SetEna('LSO_AFAPPFRAIS', FF, false);
    SetEna('LSO_AFAPPPRES', FF, false);
    SetEna('LSO_AFAPPFOUR', FF, false);
    SetEna('SO_AFAPPFRAIS', FF, false);
    SetEna('SO_AFAPPPRES', FF, false);
    SetEna('SO_AFAPPFOUR', FF, false);
    SetEna('SO_AFAPPAVECBM', FF, false);
    SetEna('SO_AFAPPANOUVEAU', FF, false);
    SetEna('SO_AFTRANSFERTENCOURS', FF, false);
    end;
TraduitForm(FF);
END;
{$ENDIF EAGLSERVER}

(*Procedure ConfigSaisieApprec ( CC : TControl ) ;
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
END;*) //Modifié par XMG de la part de Pierre LENORMAND



{$IFNDEF EAGLSERVER}
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
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure ConfigSaisiePreferences ( CC : TControl ) ;
Var FF   : TForm ;
    //PCH 11-2005 Personnalisation paramsoc
    {$IFDEF GIGI}
    CbPersoChoix    : TCheckBox;
    CVCBPersoParam1 : THValComboBox;
    CVCBPersoParam2 : THValComboBox;
    {$ENDIF GIGI}
    //
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFGESTIONCOM',FF) then Exit ;


SetEna('SO_AFANNUAIREIMPETIERS', FF, false);
// Si on est dans la base 00, on force lien DP à VRAI, à Faux sinon
if (ctxScot in V_PGI.PGIContexte) then
    begin
    if V_PGI.RunFromLanceur then
       begin
        If not TableSurAUtreBase('ANNUAIRE') then
            begin
            SetParamsoc('SO_AFLIENDP',true);
            SetEna('SO_AFANNUAIREIMPETIERS', FF, true)
            end
        else
            begin
            // si modif voir aussi MDISPGA dispatch 100
            SetParamsoc('SO_AFLIENDP',FALSE);
            end;
      end else begin
					SetParamsoc('SO_AFLIENDP',FALSE);
          SetInvi('SO_GICREATDOSS', FF);
          SetInvi('SO_GICREATANNUAIRE_FIN', FF);
          end;

    if V_PGI.SAV then
        SetEna('SO_AFLIENDP', FF, true)
    else
        SetEna('SO_AFLIENDP', FF, False);
    if Not Vh_GC.GaPlanningSeria then
      THValComboBox(GetFromNam('SO_AFECRANACCUEIL',FF)).Plus := ' AND CO_CODE <>"PLA"';
    if Not VH_GC.GaPlanChargeSeria then
      THValComboBox(GetFromNam('SO_AFECRANACCUEIL',FF)).Plus := ' AND CO_CODE <>"PC"';
    if (Not Vh_GC.GaPlanningSeria) and (Not VH_GC.GaPlanChargeSeria)then
      THValComboBox(GetFromNam('SO_AFECRANACCUEIL',FF)).Plus := ' AND CO_CODE <>"PC" AND CO_CODE <>"PLA"';
    if not GetParamsocSecur ('SO_AFLIENDP',false) then
    begin
       SetInvi('SO_AFGROUPECLIENT', FF);//mcd 01/06/07
       SetInvi('SO_AFGROUPERESS', FF);//mcd 20/02/2008 15001
    end

    end
else
    begin
    SetParamsoc('SO_AFLIENDP',FALSE);
    SetInvi('SO_AFLIENDP', FF);
    SetInvi('SO_GICREATDOSS', FF);
    SetInvi('SO_GICREATANNUAIRE_FIN', FF);
//    SetParamsoc('SO_AFGESTIONTARIF',TRUE);   // gm 10/05/07 , à priori chgt de place
//    SetInvi('SO_AFGESTIONTARIF', FF);
    SetParamsoc('SO_AFANNUAIREIMPETIERS',FALSE);
    SetInvi('SO_AFANNUAIREIMPETIERS', FF);
    SetInvi('SO_AFGROUPECLIENT', FF);
    SetInvi('SO_AFGROUPERESS', FF);  //mcd 15186
    THValComboBox(GetFromNam('SO_AFECRANACCUEIL',FF)).Plus := ' AND CO_CODE <>"APP"';
    end;

// On ne doit pas pouvoir modifier l'option de gestion de la première semaine si de l'activite a deja ete saisie
if ExisteActivite then
  SetEna('SO_PREMIERESEMAINE', FF, False)
else
  SetEna('SO_PREMIERESEMAINE', FF, true);

//PCH 18-11-05 personnalisation paramsoc
{$IFNDEF GIGI}
  SetInvi('SCO_AFPERSONNALISATION', FF);
  SetInvi('LSCO_AFPERSONNALISATION', FF);
  SetInvi('SO_AFPERSOCHOIX', FF);
  SetInvi('LSO_AFPERSOPARAM1', FF);
  SetInvi('SO_AFPERSOPARAM1', FF);
  SetInvi('LSO_AFPERSOPARAM2', FF);
  SetInvi('SO_AFPERSOPARAM2', FF);
  SetInvi('LSO_AFPERSOPARAM3', FF);
  SetInvi('SO_AFPERSOPARAM3', FF);
{$ELSE GIGI}
{$ifdef OGC}
  SetInvi('SCO_AFPERSONNALISATION', FF);
  SetInvi('LSCO_AFPERSONNALISATION', FF);
  SetInvi('SO_AFPERSOCHOIX', FF);
  SetInvi('LSO_AFPERSOPARAM1', FF);
  SetInvi('SO_AFPERSOPARAM1', FF);
  SetInvi('LSO_AFPERSOPARAM2', FF);
  SetInvi('SO_AFPERSOPARAM2', FF);
  SetInvi('LSO_AFPERSOPARAM3', FF);
  SetInvi('SO_AFPERSOPARAM3', FF);
{$else}
  SetEna('SO_AFPERSOCHOIX', FF, True);
  CbPersoChoix:=TCheckBox(GetFromNam ('SO_AFPERSOCHOIX', FF));
  CVCBPersoParam1 := THValComboBox (GetFromNam ('SO_AFPERSOPARAM1', FF));
  CVCBPersoParam2 := THValComboBox (GetFromNam ('SO_AFPERSOPARAM2', FF));

  SetEna('SO_AFPERSOPARAM1', FF, False);
  SetEna('LSO_AFPERSOPARAM1', FF, False);
  SetEna('SO_AFPERSOPARAM2', FF, False);
  SetEna('LSO_AFPERSOPARAM2', FF, False);
  SetEna('SO_AFPERSOPARAM3', FF, False);
  SetEna('LSO_AFPERSOPARAM3', FF, False);

  // Ne pas autoriser la modification des paramètres soc si une
  // personnalisation de paramètres soc a déjà été faite
  //if GetParamSocSecur('SO_AFPERSOCHOIX','X') = 'X' then
  if CbPersoChoix.Checked then
    if Not ExisteSQL('SELECT APA_NOMPARAMSOC FROM AFPARAMSOC' ) then
      begin
        SetEna('SO_AFPERSOPARAM1', FF, True);
        SetEna('LSO_AFPERSOPARAM1', FF, True);
        if Not (CVCBPersoParam1.value='') then
          begin
            SetEna('SO_AFPERSOPARAM2', FF, True);
            SetEna('LSO_AFPERSOPARAM2', FF, True);
            if Not (CVCBPersoParam2.value='') then
              begin
                SetEna('SO_AFPERSOPARAM3', FF, True);
                SetEna('LSO_AFPERSOPARAM3', FF, True);
              end
          end
      end;
{$ENDIF OGC}
{$ENDIF GIGI}
//
if not (V_PGI.SAV) then
begin
  SetInvi('LSO_AFFICHETIERS', FF);
  SetInvi('SO_AFFICHETIERS', FF);
end;
TraduitForm(FF);
END;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure ConfigSaisieCutOff ( CC : TControl ) ;
Var FF   : TForm ;
    CB : THCheckBox;
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ;
if Not ExisteNam('SO_AFAPPCUTOFF',FF) then Exit ;

SetEna('SO_AFAPPCUTOFF', FF, True);
SetEna('SO_AFCUTMODEECLAT', FF, True);
SetEna('LSO_AFCUTMODEECLAT', FF, True);
SetEna('SO_AFALIMCUTOFF', FF, True);
SetEna('SO_DATECUTOFF', FF, True);
SetEna('LSO_DATECUTOFF', FF, True);
SetEna('SO_AFFORMULCUTOFF', FF, False);
SetEna('LSO_AFFORMULCUTOFF', FF, False);
//Si cocher'Génération des cut-off à zéro' rendre invisible 'Ignorer les valeurs nulles des calculs générés'
CB := THCheckBox (GetFromNam ('SO_AFALIMCUTOFF', FF));
if CB.checked then
  SetInvi('SO_AFCUTOFFNULS', FF);
//Pas d'accès s'il existe des cut-off
//accès après annulation des cut-offs dans administration assistance avec mot du jour
if ExisteAfCumul ('CVE') then
  begin
    SetEna('SO_AFCUTMODEECLAT', FF, false);
    SetEna('LSO_AFCUTMODEECLAT', FF, false);
    SetEna('LSO_DATECUTOFF', FF, false);
    SetEna('SO_DATECUTOFF', FF, false);
  end
else
  if (ctxScot  in V_PGI.PGIContexte) then
  begin
  // Dans Scot, la formule est figée par défaut
  SetParamsoc('SO_AFFORMULCUTOFF','[M0]-[M2]');
  end;

//AB-200406 - Cut-off des achats
SetEna('SO_AFFORMULECUTOFFACH', FF, False);
SetEna('LSO_AFFORMULECUTOFFACH', FF, False);
//Si cocher'Génération des cut-off à zéro' rendre invisible 'Ignorer les valeurs nulles des calculs générés'
CB := THCheckBox (GetFromNam ('SO_AFALIMCUTOFFACH', FF));
if CB.checked then
  SetInvi('SO_AFCUTOFFNULSACH', FF);
//Pas d'accès s'il existe des cut-off
//accès après annulation des cut-offs dans administration assistance avec mot du jour
if ExisteAfCumul ('CAC') then
  begin
    SetEna('SO_AFCUTMODEECLATACH', FF, false);
    SetEna('LSO_AFCUTMODEECLATACH', FF, false);
    SetEna('SO_DATECUTOFFACH', FF, false);
    SetEna('LSO_DATECUTOFFACH', FF, false);
  end
else
  SetParamsoc('SO_AFFORMULECUTOFFACH','[M2]-[M6]');  //formule des cut-off achats non-modifiables

if (ctxScot  in V_PGI.PGIContexte) or (not VH_GC.GAAchatSeria) or
   ((ctxGCAFF in V_PGI.PGIContexte) and not VH_GC.GASeria) then
begin
  SetInvi('SCO_AFGPCUTOFFACCPTA', FF);
  SetInvi('LSCO_AFGPCUTOFFACCPTA', FF);
  SetInvi('SO_AFAPPCUTOFFACH', FF);
  SetInvi('SO_AFALIMCUTOFFACH', FF);
  SetInvi('SO_AFCUTOFFNULSACH', FF);
  SetInvi('SO_DATECUTOFFACH', FF);
  SetInvi('LSO_DATECUTOFFACH', FF);
  SetInvi('SO_AFFORMULECUTOFFACH', FF);
  SetInvi('LSO_AFFORMULECUTOFFACH', FF);
  SetInvi('SO_AFCUTMODEECLATACH', FF);
  SetInvi('LSO_AFCUTMODEECLATACH', FF);
  SetInvi('LSO_AFCUTOFFPASCHA', FF);
  SetInvi('SO_AFCUTOFFCPTAAR', FF);
  SetInvi('SO_AFCUTOFFCPTFAR', FF);
  SetInvi('SO_AFCUTOFFCPTCCA', FF);
  SetInvi('SO_AFCUTOFFTVAAAR', FF);
  SetInvi('SO_AFCUTOFFTVAFAR', FF);

  SetInvi('SO_AFCUTOFFJALCHA', FF);
  SetInvi('SO_AFCUTOFFCPTCHA', FF);
  SetInvi('SO_AFCUTOFFPASCHA', FF);
  SetInvi('LSO_AFCUTOFFCPTAAR', FF);
  SetInvi('LSO_AFCUTOFFCPTFAR', FF);
  SetInvi('LSO_AFCUTOFFCPTCCA', FF);
  SetInvi('LSO_AFCUTOFFTVAAAR', FF);
  SetInvi('LSO_AFCUTOFFTVAFAR', FF);
  SetInvi('LSO_AFCUTOFFJALCHA', FF);
  SetInvi('LSO_AFCUTOFFCPTCHA', FF);
end;


If CtxScot in V_PGI.PGICOntexte then
 begin   //mcd 02/03/2005 pas de génération compta à c ejour.. on cache
  if (GetParamSocSecur ('SO_AFCLIENT', 0) <>8 ) then  //cInClientKPMG =8 pour ne pas inclure AffaireUtil
    begin
    SetInvi('SO_AFCUTOFFCPTFAE', FF);
    SetInvi('SO_AFCUTOFFCPTAAE', FF);
    SetInvi('SO_AFCUTOFFCPTPCA', FF);
    SetInvi('SO_AFCUTOFFTVAFAE', FF);
    SetInvi('SO_AFCUTOFFTVAAAE', FF);
    SetInvi('SO_AFCUTOFFJALPRO', FF);
    SetInvi('SO_AFCUTOFFCPTPRO', FF);
    SetInvi('SO_AFCUTOFFPASPRO', FF);
    SetInvi('LSO_AFCUTOFFCPTFAE', FF);
    SetInvi('LSO_AFCUTOFFCPTAAE', FF);
    SetInvi('LSO_AFCUTOFFCPTPCA', FF);
    SetInvi('LSO_AFCUTOFFTVAFAE', FF);
    SetInvi('LSO_AFCUTOFFTVAAAE', FF);
    SetInvi('LSO_AFCUTOFFJALPRO', FF);
    SetInvi('LSO_AFCUTOFFCPTPRO', FF);
    SetInvi('LSO_AFCUTOFFPASPRO', FF);
    end;
  end;
TraduitForm(FF);
END;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure ConfigComportementActivite ( CC : TControl ) ;
Var FF   : TForm ;
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFRECHTIERSPARNOM',FF) then Exit ;

if (ctxScot in V_PGI.PGIContexte) then
  begin
  THLabel(GetFromNam('SO_AFSAISINTERVAFF',FF)).caption:='Recherche réduite aux intervenants des missions';
	end;

SetInvi('SO_AFTESTPLANNINGJOUR', FF); // en attendant la FQ 12227


TraduitForm(FF);
END;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure ConfigSaisieActivite ( CC : TControl ) ;
Var FF   : TForm ;
    CbFraisKM : TCheckBox;
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFDATEDEBUTACT',FF) then Exit ;

SetInvi('LSO_AFTYPEACTVISA', FF); //mcd 18/12/2006 le libellé ne doit pas se voir
if (ctxScot in V_PGI.PGIContexte) then
  begin
  SetInvi('So_AFREMISEMONTANT', FF); // mcd 18/06/03 // PL le 02/10/03 : pour généraliser à toute la GA
	end;

if (GetParamSoc('SO_TYPEDATEFINACTIVITE')='DAF') then
    SetEna('SO_AFDATEFINACT', FF, True)
else
    SetEna('SO_AFDATEFINACT', FF, False);

if (ExisteSQL('SELECT ACT_TYPEACTIVITE FROM ACTIVITE'))
   or  (ExisteSQL('SELECT APL_TYPEPLA FROM AFPLANNING')) //mcd 29/07/04 ... le planning est aussi concerné
  then SetEna('SO_AFMESUREACTIVITE', FF,False); // mcd 17/10/02 changé par ligne menu

if (ctxScot  in V_PGI.PGIContexte) then
    // Dans Scot
   begin
   SetEna('SO_AFDATEDEBUTACT', FF, False);
   end;
SetInvi('SO_AFUTILTARIFACT', FF); //pas géré pour l'instant

// PL le 05/07/05 : FQ 11639 gestion des frais km
CbFraisKM := TCheckBox (GetFromNam ('SO_AFGESTFRAISKM', FF));
If CbFraisKM.checked then
  begin
    SetEna('SO_AFBOOLLIBARTFKM', FF, True);
    SetEna('SO_AFMTLIBCLIFKM', FF, True);
  end
else
  begin
    SetEna('LSO_AFBOOLLIBARTFKM', FF, false);
    SetEna('LSO_AFMTLIBCLIFKM', FF, false);
    SetEna('SO_AFBOOLLIBARTFKM', FF, false);
    SetEna('SO_AFMTLIBCLIFKM', FF, false);
  end;

{$IFNDEF EAGLSERVER}
  // BDU - 24/11/06 - 12887, Visa automatique
  GestionVisaAutomatique(FF);
{$ENDIF EAGLSERVER}

TraduitForm(FF);
END;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
//PL 18/02/04 nouvelle fct
Procedure ConfigFrais (CC : TControl);
Var FF   : TForm ;
    CbFrais, CbAlimTTC, CbAlimSiZero : TCheckBox;
    CMEFraisJournal, CVCBFraisModeGener, CVCBModeSaisFrais, CVCBPrixGereFraisAct, CVCBFraisTypeEcriture : THValComboBox;
BEGIN
// test de la page
FF := TForm (CC.Owner);
if Not ExisteNam ('SO_AFFRAISCOMPTA', FF) then Exit;
CbFrais := TCheckBox (GetFromNam ('SO_AFFRAISCOMPTA', FF));
// PL le 23/03/07 : FQ 13670 (alimentation par le TTC seulement si zéro) + bonne gestion du comportement grisage/dégrisage
CbAlimTTC := TCheckBox (GetFromNam ('SO_AFALIMAUTOPRIXTTC', FF));
CbAlimSiZero := TCheckBox (GetFromNam ('SO_AFALIMTTCSIZERO', FF));
CVCBFraisModeGener := THValComboBox (GetFromNam ('SO_AFFRAISMODEGENER', FF));
CVCBModeSaisFrais := THValComboBox (GetFromNam ('SO_AFMODESAISFRAIS', FF));
CVCBPrixGereFraisAct := THValComboBox (GetFromNam ('SO_AFPRIXGEREFRAISACT', FF));
CVCBFraisTypeEcriture := THValComboBox (GetFromNam ('SO_AFFRAISTYPEECRITURE', FF));
CMEFraisJournal := THValComboBox (GetFromNam ('SO_AFFRAISJOURNAL', FF));
// Fin PL le 23/03/07

If CbFrais.checked then
  begin
    SetEna('SO_AFFRAISJOURNAL', FF, True);
    SetEna('SO_AFFRAISMODEGENER', FF, True);
    SetEna('SO_AFMODESAISFRAIS', FF, True);
    SetEna('SO_AFPRIXGEREFRAISACT', FF, True);
    SetEna('SO_AFALIMAUTOPRIXTTC', FF, True);
    // PL le 23/03/07 : FQ 13670 (alimentation par le TTC seulement si zéro) + bonne gestion du comportement grisage/dégrisage
    SetEna('SO_AFFRAISTYPEECRITURE', FF, True);
    if (CVCBFraisModeGener.value='') then CVCBFraisModeGener.Value := 'DET';
    if (CVCBModeSaisFrais.value='') then CVCBModeSaisFrais.Value := 'SHT';
    if (CVCBPrixGereFraisAct.value='') then CVCBPrixGereFraisAct.Value := 'PV';
    if (CVCBFraisTypeEcriture.value='') then CVCBFraisTypeEcriture.Value := 'N';
    if (CbAlimTTC <> nil) then
      if not CbAlimTTC.checked then
        begin
          SetEna('SO_AFALIMTTCSIZERO', FF, False);
          CbAlimSiZero.Checked := false;
        end;
    // Fin PL le  23/03/07
  end
else
  begin
    SetEna('LSO_AFFRAISJOURNAL', FF, false);
    SetEna('LSO_AFFRAISMODEGENER', FF, false);
    SetEna('LSO_AFMODESAISFRAIS', FF, false);
    SetEna('LSO_AFPRIXGEREFRAISACT', FF, false);
    // PL le 23/03/07 : FQ 13670 (alimentation par le TTC seulement si zéro) + bonne gestion du comportement grisage/dégrisage
    SetEna('LSO_AFFRAISTYPEECRITURE', FF, false);
    SetEna('SO_AFFRAISJOURNAL', FF, false);
    SetEna('SO_AFFRAISMODEGENER', FF, false);
    SetEna('SO_AFMODESAISFRAIS', FF, false);
    SetEna('SO_AFPRIXGEREFRAISACT', FF, false);
    SetEna('SO_AFALIMAUTOPRIXTTC', FF, false);
    // PL le 23/03/07 : FQ 13670 (alimentation par le TTC seulement si zéro) + bonne gestion du comportement grisage/dégrisage
    SetEna('SO_AFALIMTTCSIZERO', FF, false);
    SetEna('SO_AFFRAISTYPEECRITURE', FF, false);
    CVCBFraisModeGener.value := '';
    CVCBModeSaisFrais.value := '';
    CVCBPrixGereFraisAct.value := '';
    CVCBFraisTypeEcriture.value := '';
    CMEFraisJournal.value := '';
    CbAlimTTC.Checked := false;
    CbAlimSiZero.Checked := false;
    // Fin PL le 23/03/07
  end;
TraduitForm(FF);
END;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure ConfigSaisieImport ( CC : TControl ) ;
Var FF   : TForm ;
BEGIN     //mcd 10/11/2004 pour traduction
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFEACTPERIODIC',FF) then Exit ;
TraduitForm(FF);
  // BDU - 29/11/07. Nouvelle exportation SDA
  GestionExportSDA(FF);
END;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure ConfigSaisieRessource ( CC : TControl ) ;
Var FF   : TForm ;
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFFRAISGEN1',FF) then Exit ;

SetInvi('SO_AFLIENPAIEDEC', FF);
SetInvi('SO_AFRESSCOMMUNE', FF);

if (ctxScot  in V_PGI.PGIContexte) then
   begin   // a remettre si création auto tiers sur un sous traitant
   SetInvi('SO_AFFRACINEAUXI', FF);
   SetInvi('LSO_AFFRACINEAUXI', FF);
    // mcd 17/04/2007 à remettre quand lien paie OK en GI PAIEGI
   SetInvi('SO_AFLIENPAIEVAR', FF);
   SetInvi('SO_AFLIENPAIEFIC', FF);
   SetInvi('LSO_AFLIENPAIEFIC', FF);
   SetInvi('SO_AFLIENPAIEANA', FF);

    // fin mcd 17/04/2007 à remettre quand lien paie OK en GI PAIEGI
   end;
if THCheckBox(GetFromNam ('SO_AFLIENPAIEDEC', FF)).checked then
  SetEna('SO_AFRESSCOMMUNE', FF, True)
else
begin
  SetParamSoc('SO_AFRESSCOMMUNE', False);
  SetEna('SO_AFRESSCOMMUNE', FF, False)
end;
{$ifdef OGC}
SetInVi('SO_AFLIENPAIEDEC', FF);
SetInVi('SO_AFRESSCOMMUNE', FF);
SetInVi('LSCO_GRLIENPAIE', FF);
SetInVi('SCO_GRLIENPAIE', FF);
SetInVi('SO_AFLIENPAIEFIC', FF);
SetInVi('LSO_AFLIENPAIEFIC', FF);
SetInVi('SO_AFLIENPAIEVAR', FF);
SetInVi('SO_AFLIENPAIEANA', FF);
SetInVi('SO_PGLIENRESSOURCE_', FF);
{$endif}
TraduitForm(FF);
END;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure ConfigAfFacture ( CC : TControl ) ;
Var FF   : TForm ;
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFACOMPTE',FF) then Exit ;

// gm non visible en  V7, prévu pour V7.10
   SetInvi('SO_AFDETOPACT', FF);
   SetInvi('SO_AFDATELIGNE', FF); //gm oter pour la V8 suite dde Anne  le 27/06/08

  // PL le 17/05/05 : gestion de la fermeture de mission sur échéance liquidative
  If Ctxscot in V_pgi.PGiCOntexte then
  begin
    SetInvi('SO_AFDATELIGNE', FF); //mcd 17/04/2007
    // PCH 12/12/2006 pour facture liquidative ajout gestion paramsoc
    // SCO_AFFACTLIQFAC, SO_AFFACTLIQUDETCUM et SO_AFLIBFACTLIQU
    if GetParamSocSecur('SO_AFGERELIQUIDE',False) then
    begin
      SetEna ('SO_AFFERMAFFECHLIQ', FF, true);
      SetEna ('SO_AFFACTLIQUDETCUM', FF, true);
      SetEna ('SO_AFLIBFACTLIQU', FF, true);
      SetEna ('LSO_AFLIBFACTLIQU', FF, true);
      if ExisteNam('SO_AFFACTSANSDET',FF) then SetEna ('SO_AFFACTSANSDET', FF, true);  //mcd 14825 18/12/07
      if ExisteNam('SO_AFFACTLIQU_1CUM',FF) then SetEna ('SO_AFFACTLIQU_1CUM', FF, true);  //mcd 14914 23/01/08
    end
    else
    begin
      SetEna ('SO_AFFERMAFFECHLIQ', FF, false);
      SetEna ('SO_AFFACTLIQUDETCUM', FF, false);
      SetEna ('SO_AFLIBFACTLIQU', FF, false);
      SetEna ('LSO_AFLIBFACTLIQU', FF, false);
      if ExisteNam('SO_AFFACTSANSDET',FF) then SetEna ('SO_AFFACTSANSDET', FF, False);  //mcd 14825 18/12/07
      if ExisteNam('SO_AFFACTLIQU_1CUM',FF) then SetEna ('SO_AFFACTLIQU_1CUM', FF, false);  //mcd 14914 23/01/08
    end;
  end
  else
  begin
    SetParamsoc('SO_AFGESTIONTARIF',TRUE);   // gm 10/05/07 , à priori chgt de place
    SetInvi('SO_AFGESTIONTARIF', FF);
    SetInvi('SCO_AFFACTLIQFAC', FF);
    SetInvi('LSCO_AFFACTLIQFAC', FF);
    SetInvi('SO_AFFERMAFFECHLIQ', FF);
    SetInvi('SO_AFFACTLIQUDETCUM', FF);
    SetInvi('SO_AFLIBFACTLIQU', FF);
    SetInvi('LSO_AFLIBFACTLIQU', FF);
    if ExisteNam('SO_AFFACTSANSDET',FF) then SetInvi('SO_AFFACTSANSDET', FF); //mcd 14825 18/12/07
    if ExisteNam('SO_AFFPRTESTAUTRE',FF) then SetInvi('SO_AFFPRTESTAUTRE', FF); //mcd 14825 18/12/07
    if ExisteNam('SO_AFFACTLIQU_1CUM',FF) then SetInvi('SO_AFFACTLIQU_1CUM', FF);  //mcd 14914 23/01/08
  end;
TraduitForm(FF);
end;

Procedure ConfigSaisiePrefAff ( CC : TControl ) ;
Var FF   : TForm ;
  C : TControl;
  OldV : String ;
  {$IFDEF GIGI}
  CbGereLiquide: TCheckBox;
  {$ENDIF}
BEGIN
  // Test de la page
  FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFFLIBFACTAFF',FF) then Exit ;

  C := GetFromNam('SO_AFFGENERAUTO', FF);
  OldV:=THValComboBox(C).Value ; // GMGM 15/12/2000 bidouille car on perd la .value
{$IFDEF BTP}
THValComboBox(C).Plus :='BTP';
{$ELSE}
  THValComboBox(C).Plus :=' AND CO_LIBRE="GA" ';    //mcd 22/08/2005 suite modif ABU
{$ENDIF}
  if ctxScot in V_PGI.PGIContexte then
    begin
    If   Vh_GC.GAFactInfo then THValComboBox(C).Plus :=' AND ((CO_LIBRE="GA" AND CO_CODE<>"CON") OR (CO_LIBRE="GI")) '
    else THValComboBox(C).Plus :=' AND CO_LIBRE="GA" AND CO_CODE<>"CON" ' ;
    end;

  THValComboBox(C).Value:=OldV ;
  // Impossible de modifier le format de l'exercice s'il existe deja des affaires dans la base
  // mcd 21/11/01 ne pas faire ... ne met pas à jour les info de code affaire !!
  //SetEna( 'SO_AFFORMATEXER', FF,  Not ExisteAffaires) ;
  SetEna( 'SO_AFFORMATEXER', FF,  False);

  if (ctxScot  in V_PGI.PGIContexte) then
  begin
  {$IFDEF GIGI}
      //mcd 16/04/2007
    If not  Vh_GC.GAFactInfo then THValComboBox(GetFromNam('SO_AFREPACTFACTLIQ',FF)).Plus := ' AND CO_CODE not like "I%"';
    THValComboBox(GetFromNam('SO_AFTYPGENFACTLIQ',FF)).Plus := ' AND (CO_CODE = "FOR" OR CO_CODE = "POT" OR CO_CODE = "MAN")';
    SetInvi('SO_AFREPRISEENTAFFAIRE', FF);
    SetInvi('LSO_AFREPRISEENTAFFAIRE', FF);
      //fin mcd 16/04/07
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
    // PCH 12/12/2006 pour facture liquidative ajout gestion paramsoc
    CbGereLiquide := TCheckBox(GetFromNam ('SO_AFGERELIQUIDE', FF));
    If (CbGereLiquide.Checked <> True) then
    begin
      SetEna('SO_AFLIBLIQFACTAF', FF, false);
      SetEna('LSO_AFLIBLIQFACTAF', FF, false);
      SetEna('SO_AFTYPGENFACTLIQ', FF, false);
      SetEna('LSO_AFTYPGENFACTLIQ', FF, false);
      SetEna('SO_AFREPACTFACTLIQ', FF, false);
      SetEna('LSO_AFREPACTFACTLIQ', FF, false);
      SetEna('SO_AFPROFACFACTLIQ', FF, false);
      SetEna('LSO_AFPROFACFACTLIQ', FF, false);
    end;
  {$ENDIF}
  end
  else
     // En contexte Tempo
  begin
    SetInvi('SO_AFFORMATEXER', FF);
    SetInvi('LSO_AFFORMATEXER', FF);
    // PCH 12/12/2006 - gestion nouveaux champs personnalisation facture liquidative
    SetInvi('SCO_AFFACTLIQ', FF);
    SetInvi('LSCO_AFFACTLIQ', FF);
    SetInvi('SO_AFGERELIQUIDE', FF);
    SetInvi('SO_AFLIBLIQFACTAF', FF);
    SetInvi('LSO_AFLIBLIQFACTAF', FF);
    SetInvi('SO_AFTYPGENFACTLIQ', FF);
    SetInvi('LSO_AFTYPGENFACTLIQ', FF);
    SetInvi('SO_AFREPACTFACTLIQ', FF);
    SetInvi('LSO_AFREPACTFACTLIQ', FF);
    SetInvi('SO_AFPROFACFACTLIQ', FF);
    SetInvi('LSO_AFPROFACFACTLIQ', FF);
  end;

  // SetInvi('SO_AFBUDZERO', FF);
  // gm non visible en  V7, prévu pour V7.10
  SetInvi('SO_AFADRESSEAFF', FF);

  TraduitForm(FF);
END;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure ConfigSaisieComplAff ( CC : TControl ) ;
Var FF   : TForm ;
 Lib : string;
BEGIN
// Test de la page
  FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFCALCULFIN',FF) then Exit ;
  if (ctxScot  in V_PGI.PGIContexte) then
  begin
   If (GetParamSoc('SO_AFFORMATEXER') <>'AUC') then begin
          // cas gestion exercice.. plus de dates gérées
          //mcd 04/01/2005 revu pour visible garanti renomé en épu et jamais voir clot tech
      THLabel(GetFromNam('LSO_AFCALCULGARANTI',FF)).caption:='Date épuration = Fin exercice client +';
      THLabel(GetFromNam('LSO_AFCALCULFIN',FF)).caption:='Fin de mission = Fin exercice client +';
      SetInvi('LSO_AFALIMGaranti', FF); //mcd 28/02/07 avec CPB7 visible, bien que dessous
      SetInvi('SO_AFALIMGaranti', FF); //mcd 28/02/07 avec CPB7 visible, bien que dessous
      SetInvi('LSO_AFALIMCLOTURE', FF);
      SetInvi('SO_AFALIMCLOTURE', FF);
      SetInvi('SCO_AFLIBALIMCLOTURE', FF);
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
          //mcd 04/01/2005 revu pour visible garanti renomé en épu et jamais voir clot tech
       THLabel(GetFromNam('LSO_AFALIMGARANTI',FF)).caption:='Date épuration = fin de mission +';
       SetInvi('LSO_AFALIMCLOTURE', FF);
       SetInvi('SO_AFALIMCLOTURE', FF);
       SetInvi('SCO_AFLIBALIMCLOTURE', FF);
       end;
   end
else
   // En contexte Tempo
   begin
// BRL 24/11/2014
//   SetInvi('SO_AFCALCULFIN', FF);
//   SetInvi('LSO_AFCALCULFIN', FF);
   SetInvi('SO_AFCALCULLIQUID', FF);
   SetInvi('SCO_AFLIBCALCULLIQUID', FF);  //mcd 02/04/02
   SetInvi('LSO_AFCALCULLIQUID', FF);
   SetInvi('SO_AFCALCULGARANTI', FF);
   SetInvi('LSO_AFCALCULGARANTI', FF);
   SetInvi('SCO_AFLIBCALGARANTIE', FF);

   SetInvi('LSO_AFFINFACTURAT', FF);
   SetInvi('SO_AFFINFACTURAT', FF);
   SetInvi('SCO_AFLIBFINFACTURAT', FF);

   end;

  if CtxScot in v_pgi.PgiContexte then   SetInvi('SO_AFCOMPLETE', FF);

 lib :=RechDom('GCZONELIBRE', 'MR1', False);
 if Copy (lib,1,2) = '.-' then
    begin
    SetInvi('SO_AFRESS1AFF', FF);
    SetInvi('LSO_AFRESS1AFF', FF);
    end
 else THLabel(GetFromNam('LSO_AFRESS1AFF',FF)).caption:='Affectation '+Lib + ' depuis zone';
 lib :=RechDom('GCZONELIBRE', 'MR2', False);
 if Copy (lib,1,2) = '.-' then
    begin
    SetInvi('SO_AFRESS2AFF', FF);
    SetInvi('LSO_AFRESS2AFF', FF);
    end
 else THLabel(GetFromNam('LSO_AFRESS2AFF',FF)).caption:='Affectation '+Lib + ' depuis zone';
 lib :=RechDom('GCZONELIBRE', 'MR3', False);
 if Copy (lib,1,2) = '.-' then
    begin
    SetInvi('SO_AFRESS3AFF', FF);
    SetInvi('LSO_AFRESS3AFF', FF);
    end
 else THLabel(GetFromNam('LSO_AFRESS3AFF',FF)).caption:='Affectation '+Lib + ' depuis zone';

  TraduitForm(FF);
END;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure ConfigSaisiePlanningCommun(CC : TControl);
Var FF : TForm ;
BEGIN
  // Test de la page
  FF := TForm(CC.Owner) ; if Not ExisteNam('SO_AFFONCTION',FF) then Exit ;
  SetInvi('SO_AFPLANHEURE', FF);
  SetInvi('LSO_AFPLANHEURE', FF);
  if (V_PGI.PassWord <> CryptageSt(DayPass(Date)))
  and (ExisteSql ('SELECT APL_TYPEPLA FROM AFPLANNING WHERE APL_TYPEPLA="PC"'))
   then  begin
   SetEna('LSO_AFVALOPR', FF, False);
   SetEna('SO_AFVALOPR', FF, False);
   SetEna('LSO_AFVALOPV', FF, False);
   SetEna('SO_AFVALOPV', FF, False);
   end;
  if (V_PGI.PassWord <> CryptageSt(DayPass(Date)))
  and (ExisteSql ('SELECT ATA_AFFAIRE FROM TACHE'))
   then  begin      //mcd 04/10/04 11543 si tache existe impossible de passer en FAM
   SetEna('LSO_AFLIENARTTACHE', FF, False);
   SetEna('SO_AFLIENARTTACHE', FF, False);
   end;
  // C.B 26/11/2004
  //If CtxScot in V_PGI.pgiContexte then SetInvi ('SO_AFPLANSIMPLE',FF);
  If (CtxScot In V_PGI.PGICOntexte) and (not VH_GC.GAPlanningSeria) then
    begin
    SetEna('SO_AFQUELPLANNING',FF, False); //mcd 27/03/2007 si pas de planning jour, zone figée
    SetEna('SO_AFAVANCEMODELE',FF, False);
    SetEna('SO_AFLIENARTTACHE',FF, False);
    SetEna('SO_AFTYPEARTDEF',FF, False);
    SetEna('SO_AFFAMILLEDEF',FF, False);
    SetEna('SO_AFTYPECONPLA',FF, False); //mcd 13/07/2005
    SetEna('SO_AFMODEPLANNING',FF, False); //mcd 13/07/2005
    SetEna('LSO_AFTYPECONPLA',FF, False); //mcd 13/07/2005
    SetEna('LSO_AFMODEPLANNING',FF, False); //mcd 13/07/2005
    end;
  If (CtxScot In V_PGI.PGICOntexte)  then
    begin   //mcd 04/02/08 toujours planning en manuel en GI pas de cde
    SetEna('SO_AFMODEPLANNING',FF, False);
    SetEna('LSO_AFMODEPLANNING',FF, False);
    end;

  if (not VH_GC.GAPlanningSeria) then
  begin
    SetInvi('SO_AFPLANAFFICH',FF);
    SetInvi('LSO_AFPLANAFFICH',FF);
    SetInvi('SO_AFPLANAFFICHM1',FF);
    SetInvi('SO_AFPLANAFFICHM2',FF);
    SetInvi('LSO_AFPLANAFFICHM1',FF);
    SetInvi('LSO_AFPLANAFFICHM2',FF);
    SetInvi('SO_AFPLANAFFICHS1',FF);
    SetInvi('SO_AFPLANAFFICHS2',FF);
    SetInvi('LSO_AFPLANAFFICHS1',FF);
    SetInvi('LSO_AFPLANAFFICHS2',FF);
    SetInvi('SCO_AFDATEAFFICH',FF);
    SetInvi('SCO_AFDATETEXTE1',FF);
    SetInvi('SCO_AFDATETEXTE2',FF);
    SetInvi('LSCO_AFDATEAFFICH',FF);
  end;

end;

//C.B 04/09/2007
// pas le bureau en ga et gc
Procedure ConfigSaisiePlanningIntegre(CC : TControl);
Var
  FF : TForm ;
  vStListe : THMultiValComboBox;
   
begin
  FF := TForm(CC.Owner); if Not ExisteNam('SO_YPLGA',FF) then Exit ;
  SetInvi('SO_YPLBUREAU', FF);
  SetInvi('LSO_YPLBUREAU', FF);

  // C.B 20/09/2007
  // pour l'instant, on cache les evenements dans l'environnement affaire ou gc
  // pas le cas pour la grc qui fonctionne avec le bureau dans le contexte gi
  if (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte) then
  begin                                                           
    vStListe := THMultiValComboBox(GetFromNam('SO_YPLGA', FF));
    vStListe.plus := 'DT_PREFIXE = "RAI" OR DT_PREFIXE = "PCN"';
    vStListe := THMultiValComboBox(GetFromNam('SO_YPLGRC', FF));
    vStListe.plus := 'DT_PREFIXE = "APL" OR DT_PREFIXE = "PCN"';
    vStListe := THMultiValComboBox(GetFromNam('SO_YPLPAIE', FF));
    vStListe.plus := 'DT_PREFIXE = "RAI" OR DT_PREFIXE = "APL"';
  end;
end;
 
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure ConfigSaisiePlanning(CC : TControl);
Var FF : TForm ;
BEGIN
  // C.B 10/05/2007 SO_AFALIGNREALISE changé de branche
  FF := TForm(CC.Owner); if Not ExisteNam('SO_AFPARSEMAINE',FF) then Exit ;

  // C.B 22/08/2005
  // on bloque les heures de début et de fin de période du planning
  if (V_PGI.PassWord <> CryptageSt(DayPass(Date)))
  and (ExisteSql ('SELECT APL_AFFAIRE FROM AFPLANNING WHERE APL_TYPEPLA="PLA"')) then
  begin
    SetEna('SO_AFAMDEBUT', FF, False);
    SetEna('SO_AFAMFIN', FF, False);
    SetEna('SO_AFPMDEBUT', FF, False);
    SetEna('SO_AFPMFIN', FF, False);
    SetEna('LSO_AFAMDEBUT', FF, False);
    SetEna('LSO_AFAMFIN', FF, False);
    SetEna('LSO_AFPMDEBUT', FF, False);
    SetEna('LSO_AFPMFIN', FF, False);
  end;

  if (ctxaffaire in V_PGI.PGIContexte) then
  Begin
    SetInvi('SO_AFECOLE', FF);
    SetInvi('SO_AFPLANNINGFACTURE', FF); //GA_20080505_GME_GA15100
  end;

  // en attendant oncodedesign  
  SetInvi('SO_AFPARPERIODE', FF);

  if not(VH_GC.GCIfDefCEGID) then
  begin

    SetInvi('SO_AFRESMAX', FF);
    SetInvi('LSO_AFRESMAX', FF);
  end;

  // BDU - 12/12/06
  GestionAffichePlanning(FF);

  //C.B 22/05/2007
  GestionTous(FF);
END;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure ConfigSaisiePDC(CC : TControl);
Var FF : TForm ;
BEGIN
  // Test de la page
  FF := TForm(CC.Owner) ; if Not ExisteNam('SO_AFPLANDECHARGE',FF) then Exit ;
  If (not VH_GC.GAPlanChargeSeria) then
  begin

    // C.B 22/06/2006
    // source inutile
    // C.B 22/08/2005
    // on bloque les heures de début et de fin de période du planning
    {if (V_PGI.PassWord <> CryptageSt(DayPass(Date)))
    and (ExisteSql ('SELECT APL_AFFAIRE FROM AFPLANNING WHERE APL_TYPEPLA="PLA"')) then
    begin
      SetEna('SO_AFAMDEBUT', FF, False);
      SetEna('SO_AFAMFIN', FF, False);
      SetEna('SO_AFPMDEBUT', FF, False);
      SetEna('SO_AFPMFIN', FF, False);
      SetEna('LSO_AFAMDEBUT', FF, False);
      SetEna('LSO_AFAMFIN', FF, False);
      SetEna('LSO_AFPMDEBUT', FF, False);
      SetEna('LSO_AFPMFIN', FF, False);
    end;} 

    SetEna('SO_AFPLANDECHARGE',FF, False);
    SetParamSoc ('SO_AFPLANDECHARGE',false);
    SetEna('SO_AFGESTIONRAF',FF, False);
    SetEna('SO_AFPLAHISTO',FF, False);
    SetEna('SO_AFRAFPLANNING',FF, False);
    SetEna('SO_AFPDCDEC',FF, False);
    SetEna('SO_AFPCSURAFF',FF, False);
    SetEna('SO_AFPDCNBAN',FF, False);
    SetEna('SO_AFPDCNBMOIS',FF, False);
    SetEna('SO_AFFORMULEPC',FF, False);
    SetEna('LSO_AFPDCNBAN',FF, False);
    SetEna('LSO_AFPDCNBMOIS',FF, False);
    SetEna('LSO_AFFORMULEPC',FF, False);
  end;

  If (CtxScot In V_PGI.PGICOntexte) and (not VH_GC.GAPlanningSeria) then
    begin    // dans un 1er temps, ces champs ne sont pas gérés.... même si PC sérialisé
    SetEna('SO_AFPLANDECHARGE',FF, False);
    SetEna('SO_AFGESTIONRAF',FF, False);
    SetEna('SO_AFPLAHISTO',FF, False);
    SetEna('SO_AFRAFPLANNING',FF, False);
   end;
  SetEna('SO_AFFORMULEPC',FF, False);
  if not (GetParamSoc('SO_AFGESTIONRAF')) then SetEna('SO_AFRAFPLANNING',FF, False);
  SetInvi('SO_AFPLAHISTO',FF);
  TraduitForm(FF);
END;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure ConfigSaisieGenePlanning(CC : TControl);
Var FF : TForm ;
BEGIN
  // Test de la page
  FF := TForm(CC.Owner) ; if Not ExisteNam('SO_AFGESTIONRESS',FF) then Exit ;
  SetInvi('SO_AFGESTIONRESS',FF);
  SetInvi('SO_AFRESSREMP',FF);
  SetInvi('LSO_AFRESSREMP',FF);
END;

procedure ConfigSaisieLienPlanningActivite(CC : TControl);
Var FF : TForm ;
begin
  // Test de la page
  FF := TForm(CC.Owner) ; if Not ExisteNam('SO_AFALIGNREALISE',FF) then Exit ;

  // C.B 07/04/2006
  // mis pu un dev a faire plus tard ...
  SetInvi('SO_AFREGLEACT', FF);
   
  // C.B 16/03/2007
  // en attente de dev de on codesign
  SetInvi('LSCO_AFTYPEHEURE', FF);
  SetInvi('SCO_AFTYPEHEURE', FF);
  SetInvi('SO_AFTHMATIN', FF);
  SetInvi('SO_AFTHAPRESMIDI', FF);
  SetInvi('SO_AFTHNUIT', FF);
  SetInvi('LSO_AFTHMATIN', FF);
  SetInvi('LSO_AFTHAPRESMIDI', FF);
  SetInvi('LSO_AFTHNUIT', FF);
end;

{$ENDIF EAGLSERVER}
                                
{$IFNDEF EAGLSERVER}
Procedure ConfigSaisieRevEtVar(CC : TControl);
Var FF : TForm ;
BEGIN
  // Test de la page
  FF := TForm(CC.Owner) ; if Not ExisteNam('SO_AFREVPATH',FF) then Exit ;

  // revision de prix tempo en attendant de developper la fonctionnalité
  SetInvi('SO_AFJUSREVDET', FF);

  //C.B
  //if ctxGCAff in V_PGI.PGIContexte then
 
  TraduitForm(FF);

end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure ConfigParamPiece ( CC : TControl ) ;
Var FF   : TForm ;
BEGIN
  //AB-200510- Pieces dans paramètre Gestion commercial -
  FF:=TForm(CC.Owner) ;
  if Not ExisteNam('SO_COMMISSIONLIGNE',FF) then Exit ;

  if (ctxaffaire in V_PGI.PGIContexte) then
  begin
    SetInvi('SCO_EDITION', FF);
    SetInvi('SCO_GCCDEOUV', FF);
    SetInvi('SO_GCCDEOUV', FF);
    SetInvi('SCO_GCGRPREAPPROMAN', FF);
    SetInvi('SCO_PARAMPIECES', FF);
    SetInvi('SCO_PIECEORIGINE', FF);
    SetInvi('SCO_POIDS', FF);
    SetInvi('SO_COEFFALERTEPB', FF);
    SetInvi('SO_CTRLPOIDSBRUT', FF);
    SetInvi('SO_FACTUREACOMPTE', FF);
    SetInvi('SO_GCCDEOUVAUTOTRANSFO', FF);
    SetInvi('SO_GCCDEOUVELIMINELIGNE0', FF);
    SetInvi('SO_GCCDEOUVLISTESAISIE', FF);
    SetInvi('SO_GCCDEOUVLISTESAISIEACH', FF);
    SetInvi('SO_GCCDEOUVMULTIARTICLE', FF);
    SetInvi('SO_GCCDEOUVNATPIE', FF);
    SetInvi('SO_GCCDEOUVTRANSFOPRE', FF);
    SetInvi('SO_IMPMODELEACOMPTE', FF);
    SetInvi('SO_NATREAPPROMAN', FF);
    SetInvi('SO_OUVREIMPACOMPTE', FF);
    SetInvi('SO_PIECEORIGINEACH', FF);
    SetInvi('SO_PIECEORIGINEVTE', FF);
    SetInvi('SO_POIDSNATUREPIECE', FF);

    SetInvi('LSCO_POIDS', FF);
    SetInvi('LSO_COEFFALERTEPB', FF);
    SetInvi('LSO_POIDSNATUREPIECE', FF);
    SetInvi('LSO_PIECEORIGINEACH', FF);
    SetInvi('LSO_PIECEORIGINEVTE', FF);
    SetInvi('LSCO_PIECEORIGINE', FF);
    SetInvi('LSCO_GCCDEOUV', FF);
    SetInvi('LSO_GCCDEOUVLISTESAISIE', FF);
    SetInvi('LSO_GCCDEOUVLISTESAISIEACH', FF);
    SetInvi('LSCO_EDITION', FF);
    SetInvi('LSO_OUVREIMPACOMPTE', FF);
    SetInvi('LSO_IMPMODELEACOMPTE', FF);
    SetInvi('LSCO_GCGRPREAPPROMAN', FF);
    SetInvi('LSO_GCCDEOUVNATPIE', FF);
    SetInvi('LSO_NATREAPPROMAN', FF);
    SetInvi('SO_INFOSCPLPIECE', FF);
    SetInvi('SO_RTINFOSCOMPL', FF);
    SetInvi('SCO_INFOSCPLPIECE', FF);
    SetInvi('SCO_RTINFOSCOMPL', FF);
    SetInvi('LSCO_INFOSCPLPIECE', FF);
    SetInvi('LSCO_RTINFOSCOMPL', FF);
    SetInvi('SO_RTGESTINFOS008', FF);
    SetInvi('SO_RTGESTINFOS00D', FF);
    SetInvi('SO_RTHAUTEURFICHE008', FF);
    SetInvi('SO_RTHAUTEURFICHE00D', FF);
    SetInvi('SO_RTLARGEURFICHE008', FF);
    SetInvi('SO_RTLARGEURFICHE00D', FF);
    SetInvi('LSO_RTHAUTEURFICHE008', FF);
    SetInvi('LSO_RTHAUTEURFICHE00D', FF);
    SetInvi('LSO_RTLARGEURFICHE008', FF);
    SetInvi('LSO_RTLARGEURFICHE00D', FF);
      //mcd 16/04/07
    SetInvi('LSCO_WCSA_', FF);
    SetInvi('SCO_WCSA_', FF);
    SetInvi('LSO_WCSA_', FF);
    SetInvi('SO_WCSA_', FF);
    SetInvi('LSO_WBSA_', FF);
    SetInvi('SO_WBSA_', FF);
    SetInvi('LSCO_WCSB_', FF);
    SetInvi('SCO_WCSB_', FF);
    SetInvi('LSO_WCSB_', FF);
    SetInvi('SO_WCSB_', FF);
    SetInvi('LSO_GCNEWLIENPIECE', FF);
    SetInvi('SO_GCNEWLIENPIECE', FF);
    SetInvi('LSO_GCOUVSAISIEREGLE', FF);
    SetInvi('SO_GCOUVSAISIEREGLE', FF);
    SetInvi('LSCO_MARCHE', FF);
    SetInvi('SO_CDEMARCHESURMARCHE', FF);
    SetInvi('SO_CDEMARCHESURMARCHEEPUISE', FF);
    SetInvi('SO_CDEMARCHEDEPOT', FF);
    SetInvi('SO_CDEMARCHE', FF);
    SetInvi('SO_CDEMARCHEA', FF);
    SetInvi('SO_CDEMARCHEINIT', FF);
    SetInvi('SO_CDEMARCHEDUP', FF);
    SetInvi('SO_CDEMARCHEV', FF);
    SetInvi('SCO_MARCHEACH', FF);
    SetInvi('SO_CDEMARCHEACH', FF);
    SetInvi('LSO_CDEMARCHEACH', FF);
    SetInvi('SO_CDEMARCHEAPPELACH', FF);
    SetInvi('LSO_CDEMARCHEAPPELACH', FF);
    SetInvi('SO_CDEMARCHEAPPELDATEACH', FF);
    SetInvi('LSO_CDEMARCHEAPPELDATEACH', FF);
    SetInvi('SO_CDEMARCHECTRLDVALACH', FF);
    SetInvi('SCO_MARCHE', FF);
    SetInvi('LSCO_MARCHEACH', FF);
    SetInvi('LSCO_MARCHEVTE', FF);
    SetInvi('SCO_MARCHEVTE', FF);
    SetInvi('SO_CDEMARCHEVTE', FF);
    SetInvi('LSO_CDEMARCHEVTE', FF);
    SetInvi('SO_CDEMARCHEAPPELVTE', FF);
    SetInvi('LSO_CDEMARCHEAPPELVTE', FF);
    SetInvi('SO_CDEMARCHEAPPELDATEVTE', FF);
    SetInvi('LSO_CDEMARCHEAPPELDATEVTE', FF);
    SetInvi('SO_CDEMARCHECTRLDVALVTE', FF);
    SetInvi('SO_GCDEPOTPIECE', FF);
    SetInvi('LSO_GCDEPOTPIECE', FF);
      //fin mcd 16/04/07
  end;
  if (ctxscot in V_PGI.PGIContexte) then
  begin //mcd 13/07/07 14360 giga
    SetInvi('LSCO_SOLDEPIECEPREC', FF);
    SetInvi('SCO_SOLDEPIECEPREC', FF);
    SetInvi('SO_SOLDEPIECEPREC', FF);
    SetInvi('SO_GCINFOLIGSTOPIE', FF);
    SetInvi('SO_SOLDEPRECMODE', FF);
    SetInvi('LSO_SOLDEPRECMODE', FF);
    SetInvi('SO_SOLDEPRECCONFIRM', FF);
    SetInvi('LSO_SOLDEPRECCONFIRM', FF);
    SetInvi('SO_SOLDEPRECACH', FF);
    SetInvi('LSO_SOLDEPRECACH', FF);
    SetInvi('SO_SOLDEPRECVEN', FF);
    SetInvi('LSO_SOLDEPRECVEN', FF);
  end;
  TraduitForm(FF);
END;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure CacheFctGIGA ( CC : TControl ) ;
Var FF   : TForm ;
BEGIN

// Configuration des écran GA (traduction, champs visibles...)
ConfigSaisieParamsDefaut(CC);
ConfigSaisieDevise(CC);
ConfigSaisieRessource(CC) ;
ConfigSaisiePrefAff(CC) ;
ConfigAfFacture(CC) ; //mcd 06/03
ConfigSaisieComplAff(CC) ;
ConfigSaisieActivite(CC) ;
ConfigComportementActivite(CC); //PL le 23/03/06
ConfigSaisiePreferences(CC) ;
ConfigSaisieEditions(CC);
ConfigSaisieDates(CC);
ConfigSaisieApprec(CC);   //mcd 12/02/02
ConfigSaisieCutOff(CC);   //mcd 11/07/02
ConfigSaisieDecen(CC);   //mcd 11/07/02
ConfigSaisieBudget(CC);
ConfigSaisieFactEclat(CC);
ConfigSaisiePlanningCommun(CC);
ConfigSaisiePlanningIntegre(CC);
ConfigSaisiePDC(CC);
ConfigSaisiePlanning(CC);
ConfigSaisieGenePlanning(CC);
ConfigSaisieLienPlanningActivite(CC);
ConfigSaisieRevEtVar(CC);
ConfigSaisieImport(CC); //mcd 10/11/2004
ConfigFrais (CC); // PL le 18/02/04
ConfigArticle (CC); // MCD 16/05/05
ConfigParamPiece (CC); //AB-200510-

if CtxScot in V_Pgi.PgIContexte then
begin //mcd 06/06/07 zones non géres en GI
  FF:=GetLaForm(CC) ;
  if ExisteNam('SO_PALMARESACTIF',FF) then
  begin
    SetInvi('SO_PALMARESACTIF', FF);
    SetInvi('SO_PALMARESTIERS', FF);
    SetInvi('SO_PALMARESFORCE', FF);
    SetInvi('SO_PALMARESENREG', FF);
    SetInvi('SCO_GRPALMARESTRANSPORTEUR', FF);
    SetInvi('LSCO_GRPALMARESTRANSPORTEUR', FF);
  end;
end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
{$ENDIF EAGLSERVER}

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
    begin       //mcd 21/09/2006 suppression Vh_GC  pour passer direct en GetParamsoc
      OB_DETAIL := TOB.Create('RESSOURCE',OB,-1);
      OB_DETAIL.PutValue('ARS_RESSOURCE','ASSBONI');
      OB_DETAIL.PutValue('ARS_TYPERESSOURCE','SAL');
      OB_DETAIL.PutValue('ARS_LIBELLE',TraduitGa('Ressource affectation Boni/mali'));
      OB_DETAIL.PutValue ('ARS_UNITETEMPS', VH_GC.AFMesureActivite  );
      OB_DETAIL.PutValue ('ARS_TAUXFRAISGEN1', Valeur(GetParamSocSecur('SO_AFFRAISGEN1', 0)) );
      OB_DETAIL.PutValue ('ARS_TAUXFRAISGEN2', Valeur(GetParamSocSecur('SO_AFFRAISGEN2', 0)) );
      OB_DETAIL.PutValue ('ARS_TAUXCHARGEPAT', Valeur(GetParamSocSecur('SO_AFCHARGESPAT', 0)));
      OB_DETAIL.PutValue ('ARS_COEFMETIER'   , Valeur(GetParamSocSecur('SO_AFCOEFMETIER', 0)) );
      OB_DETAIL.PutValue ('ARS_ARTICLE'      , GetParamSocSecur('SO_AFPRESTATIONRES', '') );
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

{$IFNDEF EAGLSERVER}
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
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
    HShowMessage('24;Société;Utilisation des fonctions d''appréciation impossible. Veuillez effectuer la Bascule EURO du dossier.;W;O;O;O;','','') ;
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
END;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure CutOffZeroX( CC : TControl );
Var FF : TForm ;
begin
  FF:=GetLaForm(CC) ;
  THCheckBox (GetFromNam ('SO_AFCUTOFFNULS', FF)).visible := not THCheckBox(CC).checked
end;

Procedure CutOffZeroAchX( CC : TControl );
Var FF : TForm ;
begin
  FF:=GetLaForm(CC) ;
  THCheckBox (GetFromNam ('SO_AFCUTOFFNULSACH', FF)).visible := not THCheckBox(CC).checked;
end;

Procedure CutOffX( CC : TControl );
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
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure AFPlanDeChargeX( CC : TControl ) ;
Var
  FF    : TForm ;
  CbPC  : TCheckBox;
  LaListe: THValComboBox;
  PlanDeCharge: Boolean;

BEGIN

  // BDU - 12/12/06. C'est bien gentil, mais ça ne marche pas.
  // On ne peux pas activer/désactiver des zones d'un autre bloc de paramètres
  FF:=GetLaForm(CC);
  PlanDeCharge := False;

  // BDU - 12/12/06. Ancien fonctionnement avec une case à cocher. Conserve la compatibilité provisoirement
  if CC is TCheckBox then
  begin
    CbPC:=TCheckBox(CC);
    PlanDeCharge := CbPC.Checked;
  end

  // BDU - 12/12/06. Nouveau fonctionnement avec une liste
  else if CC is THValComboBox then
  begin
    LaListe := THValComboBox(CC);
    PlanDeCharge := LaListe.Values[LaListe.ItemIndex] = 'PDC';
  end;
  if PlanDeCharge then
  begin  // si coché, on renseigne les valeurs par défaut
    // C.B 19/10/2007 FQ 14665
    // paramsoc suivants changés de branches
    //SetEna('SO_AFLIENPLANACT',FF, False);
    //SetEna('SO_AFALIGNREALISE',FF, False);
    //SetEna('SO_AFPLANNINGETAT',FF, False);
    //SetEna('SO_AFETATINTERDIT',FF, False);
    SetEna('SO_AFPLANAFFICH',FF, False);

    // mcd 08/01/2008 ?? n'existe plus ....SetInvi('LSO_AFDATETEXTE1', FF);

    SetEna('SO_AFPLANAFFICHM1',FF, False);
    SetEna('SO_AFPLANAFFICHM2',FF, False);

    // C.B 28/04/2005 ajout des paramsoc dates planning
    SetEna('SO_AFPLANAFFICHS1',FF, False);
    SetEna('SO_AFPLANAFFICHS2',FF, False);
    SetEna('SCO_AFDATEAFFICH',FF, False);
    SetEna('SCO_AFDATETEXTE1',FF, False);
    SetEna('SCO_AFDATETEXTE2',FF, False);
    SetEna('LSCO_AFDATEAFFICH',FF, False);
  end
  else
  begin
    // C.B  19/10/2007 FQ 14665
    // paramsoc suivants changés de branches
    //SetEna('SO_AFLIENPLANACT',FF, True);
    //SetEna('SO_AFALIGNREALISE',FF, True);
    //SetEna('SO_AFPLANNINGETAT',FF, True);
    //SetEna('SO_AFETATINTERDIT',FF, True);
    SetEna('SO_AFPLANAFFICH',FF, True);
    SetEna('SO_AFPLANAFFICHM1',FF, True);
    SetEna('SO_AFPLANAFFICHM2',FF, True);
                        
    // C.B 28/04/2005 ajout des paramsoc dates planning
    SetEna('SO_AFPLANAFFICHS1',FF, True);
    SetEna('SO_AFPLANAFFICHS2',FF, True);
    SetEna('SCO_AFDATEAFFICH',FF, True);
    SetEna('SCO_AFDATETEXTE1',FF, True);
    SetEna('SCO_AFDATETEXTE2',FF, True);
    SetEna('LSCO_AFDATEAFFICH',FF, True);
  end;
end;
{$ENDIF EAGLSERVER}
 
{$IFNDEF EAGLSERVER}
Procedure AFGestionRAFX( CC : TControl ) ;
Var FF : TForm ;
    CbPC: TCheckBox;
BEGIN
FF:=GetLaForm(CC) ;
CbPC:=TCheckBox(CC);
if CbPC.Checked then
    begin  // si coché, on renseigne les valeurs par défaut
    SetEna('SO_AFRAFPLANNING',FF, True);
    end
else begin
    SetEna('SO_AFRAFPLANNING',FF, False);
    end;
end;
{$ENDIF EAGLSERVER}


{$IFNDEF EAGLSERVER}
// PL le 18/02/04 : nouvelle fct
Procedure FraisX (CC : TControl);
Var FF : TForm;
    CbFrais, CbAlimPrixTTC, CbAlimSiZero : TCheckBox;
    LblFraisJournal, LblFraisModeGener, LblModeSaisFrais, LblPrixGereFraisAct, LblFraisTypeEcriture : THLabel;
    CMEFraisJournal, CVCBFraisModeGener, CVCBModeSaisFrais, CVCBPrixGereFraisAct, CVCBFraisTypeEcriture : THValComboBox;
BEGIN
FF := GetLaForm(CC);
CbFrais := TCheckBox(CC);

CMEFraisJournal := THValComboBox (GetFromNam ('SO_AFFRAISJOURNAL', FF));
CVCBFraisModeGener := THValComboBox (GetFromNam ('SO_AFFRAISMODEGENER', FF));
CVCBModeSaisFrais := THValComboBox (GetFromNam ('SO_AFMODESAISFRAIS', FF));
CVCBPrixGereFraisAct := THValComboBox (GetFromNam ('SO_AFPRIXGEREFRAISACT', FF));
LblFraisJournal := THLabel (GetFromNam ('LSO_AFFRAISJOURNAL', FF));
LblFraisModeGener := THLabel (GetFromNam ('LSO_AFFRAISMODEGENER', FF));
LblModeSaisFrais := THLabel (GetFromNam ('LSO_AFMODESAISFRAIS', FF));
LblPrixGereFraisAct := THLabel (GetFromNam ('LSO_AFPRIXGEREFRAISACT', FF));
CbAlimPrixTTC := TCheckBox(GetFromNam ('SO_AFALIMAUTOPRIXTTC', FF));
// PL le 23/03/07 : FQ 13670 (alimentation par le TTC seulement si zéro) + bonne gestion du comportement grisage/dégrisage
CbAlimSiZero := TCheckBox(GetFromNam ('SO_AFALIMTTCSIZERO', FF));
CVCBFraisTypeEcriture := THValComboBox (GetFromNam ('SO_AFFRAISTYPEECRITURE', FF));
LblFraisTypeEcriture := THLabel (GetFromNam ('LSO_AFFRAISTYPEECRITURE', FF));
// Fin PL le 23/03/07

if CbFrais.Checked then
    begin
    CMEFraisJournal.Enabled := true;
    CVCBFraisModeGener.Enabled := true;
    CVCBModeSaisFrais.Enabled := true;
    CVCBPrixGereFraisAct.Enabled := true;
    CVCBFraisTypeEcriture.Enabled := true;
    LblFraisJournal.Enabled := true;
    LblFraisModeGener.Enabled := true;
    LblModeSaisFrais.Enabled := true;
    LblPrixGereFraisAct.Enabled := true;
    LblFraisTypeEcriture.Enabled := true;
    if (CVCBFraisModeGener.value='') then CVCBFraisModeGener.Value := 'DET';
    if (CVCBModeSaisFrais.value='') then CVCBModeSaisFrais.Value := 'SHT';
    if (CVCBPrixGereFraisAct.value='') then CVCBPrixGereFraisAct.Value := 'PV';
    // PL le 23/03/07 : FQ 13670 (alimentation par le TTC seulement si zéro) + bonne gestion du comportement grisage/dégrisage
    if (CVCBFraisTypeEcriture.value='') then CVCBFraisTypeEcriture.Value := 'N';
		CbAlimPrixTTC.Enabled := true;
    if CbAlimPrixTTC.Checked then
		  CbAlimSiZero.Enabled := true
    else
      begin
      CbAlimSiZero.Enabled := false;
      CbAlimSiZero.Checked := false;
      end;
    // Fin PL le 23/03/07
    end
else
    begin   // si case non cochée, on efface tout ce qui est dans les zones suivantes
    CMEFraisJournal.Enabled := false;
    CVCBFraisModeGener.Enabled := false;
    CVCBModeSaisFrais.Enabled := false;
    CVCBPrixGereFraisAct.Enabled := false;
    CVCBFraisTypeEcriture.Enabled := false;
    LblFraisJournal.Enabled := false;
    LblFraisModeGener.Enabled := false;
    LblModeSaisFrais.Enabled := false;
    LblPrixGereFraisAct.Enabled := false;
    LblFraisTypeEcriture.Enabled := false;
    CMEFraisJournal.value := '';
    CVCBFraisModeGener.value := '';
    CVCBModeSaisFrais.value := '';
    CVCBPrixGereFraisAct.value := '';
    // PL le 23/03/07 : FQ 13670 (alimentation par le TTC seulement si zéro) + bonne gestion du comportement grisage/dégrisage
    CVCBFraisTypeEcriture.value := '';
		CbAlimPrixTTC.Enabled := false;
		CbAlimPrixTTC.Checked := false;
		CbAlimSiZero.Enabled := false;
    CbAlimSiZero.Checked := false;
    // Fin Pl le 23/03/07
    end;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure ActiviteFraisKMX (CC : TControl);
Var FF : TForm;
    CbFraisKM : TCheckBox;
    LblBoolLibArt, LblMtLibCli : THLabel;
    CVCBBoolLibArt, CVCBMtLibCli : THValComboBox;
BEGIN
FF := GetLaForm(CC);

// PL le 05/07/05 : FQ 11639 gestion des frais km
CbFraisKM := TCheckBox(CC);
LblBoolLibArt := THLabel (GetFromNam ('LSO_AFBOOLLIBARTFKM', FF));
LblMtLibCli := THLabel (GetFromNam ('LSO_AFMTLIBCLIFKM', FF));
CVCBBoolLibArt := THValComboBox (GetFromNam ('SO_AFBOOLLIBARTFKM', FF));
CVCBMtLibCli := THValComboBox (GetFromNam ('SO_AFMTLIBCLIFKM', FF));

If CbFraisKM.checked then
  begin
    CVCBBoolLibArt.Enabled := true;
    CVCBMtLibCli.Enabled := true;
    LblBoolLibArt.Enabled := true;
    LblMtLibCli.Enabled := true;
  end
else
  begin
    CVCBBoolLibArt.Enabled := false;
    CVCBMtLibCli.Enabled := false;
    LblBoolLibArt.Enabled := false;
    LblMtLibCli.Enabled := false;
    CVCBBoolLibArt.value := '';
    CVCBMtLibCli.value := '';
  end;

END ;
{$ENDIF EAGLSERVER}

{$endif}
// fin fct GI/GA

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

Function ControlePrefixe ( Prefixe : String; CaractereAutorise : String ) : boolean ;
Var i, j : Integer;
    CarOk: Boolean;
BEGIN
Result:=True;
for i:=1 to Length(Prefixe) do
    begin
    CarOk:=False;
    for j:=1 to Length(CaractereAutorise) do
        begin
        if Prefixe[i]=CaractereAutorise[j] then begin CarOk:=True; break; end;
        end;
    if CarOk = False then begin Result:=False; break; end;
    end;
END ;

Function ControleChrono ( Chrono : String ) : boolean ;
BEGIN
Result:=True;
if (not IsNumeric(Chrono))then
   begin
   HShowMessage('30;Société;Le chrono doit être numérique.;W;O;O;O;','','') ;
   Result:=False;
   end;
END ;

Function ControleLgCode ( LgMin:integer;LgMax:integer;LgCode:integer;LgPrefixe:integer;LgCompteur:integer) : boolean;
BEGIN
Result:=True;
if (LgCode<LgPrefixe+LgMin) then
   begin
   Result := False;
   HShowMessage('26;Société;Longueur minimum du code : '+InttoStr(LgPrefixe+LgMin)+';W;O;O;O;','','')
   end;
if (LgCode>LgMax) then
   begin
   Result := False;
   HShowMessage('27;Société;Longueur maximum du code : '+InttoStr(LgMax)+';W;O;O;O;','','') ;
   Exit;
   end;
if (LgCompteur > (LgCode-LgPrefixe)) then
   begin
   Result := False;
   HShowMessage('29;Société;La longueur du code est trop petite.;W;O;O;O;','','') ;
   Exit;
   end;
END ;


{$IFNDEF EAGLSERVER}
Function RechExisteH ( CC : TControl ) : boolean ;
BEGIN Result:=(RechDomLookUpCombo(CC)<>'') ; END ;
{$ENDIF EAGLSERVER}

Function DesEnregs ( TableName : String ) : boolean ;
BEGIN
Result := ExisteSQL('SELECT 1 FROM '+TableName);
END ;

{$IFNDEF EAGLSERVER}
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
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure CptBourreX ( CC : TControl ) ;
Var CptBil : THCritmaskEdit ;
BEGIN
CptBil:=THCritMaskEdit(CC) ;
if CptBil<>Nil then BourreLeCpt(CptBil) ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
{$ENDIF EAGLSERVER}

//
// Attribution automatique du code Tiers
//
{$IFNDEF EAGLSERVER}
Procedure NumTiersAutoX ( CC : TControl; ForceActive : Boolean ) ;
Var FF : TForm ;
  { GC_DBR_CRM10764_DEB }
  {$IFDEF PGIMAJVER}
    NumTiersAuto : TCheckBox ;
    LgNumTiers   : TSpinEdit ;
    CompteurTiers: TEdit ;
  {$ENDIF PGIMAJVER}
  { GC_DBR_CRM10764_FIN }
    Active       : Boolean ;
    ActiveFormat : boolean;
BEGIN
  FF:=GetLaForm(CC) ;
  if Not ExisteNam('SO_GCNUMTIERSAUTO',FF) and Not ExisteNam('SO_BOURREAUX',FF) then Exit ;
  { GC_DBR_CRM10764_DEB }
  ActiveFormat := false;
  Active := false;
  {$IFNDEF PGIMAJVER}
  if ExisteNam('SO_BOURREAUX',FF) then
  begin
    {$IFDEF GCGC}
    SetEna ('SO_BOURREAUX', FF, False);
    SetEna ('SO_BOURREGEN', FF, False);
    SetEna ('SO_LGMAXBUDGET', FF, False);
    {$ENDIF GCGC}
  end else
  begin
    SetEna('SO_GCNUMTIERSAUTO',FF, false);
    SetEna('SO_GCNUMFOURAUTO',FF, false);
    SetEna('SO_GCFORMATCODETIERS',FF, false);
  end;
  {$ELSE !PGIMAJVER}
  if ExisteNam('SO_GCNUMTIERSAUTO',FF) then
  begin
    SetEna ('SO_GCFORMATCODETIERS', FF, not (EstBaseMultiSoc and TablePartagee ('TIERS')));
    SetEna('SO_GCNUMTIERSAUTO',FF, true);
    SetEna('SO_GCNUMFOURAUTO',FF, true);
    ActiveFormat := TCheckBox(GetFromNam('SO_GCFORMATCODETIERS',FF)).Checked;
    LgNumTiers:=TSpinEdit(GetFromNam('SO_GCLGNUMTIERS',FF));

    if ForceActive=True then Active:=True
    else
    begin
      NumTiersAuto:=TCheckBox(GetFromNam('SO_GCNUMTIERSAUTO',FF)) ;
      SetEna('SO_GCFORMATCODETIERS', FF, NumTiersAuto.Checked and not (EstBaseMultiSoc and TablePartagee ('TIERS'))) ;
      CompteurTiers:=TEdit(GetFromNam('SO_GCCOMPTEURTIERS',FF));
      if NumTiersAuto.Checked then
      begin
        Active := True;
        if LgNumTiers.Value=0 then LgNumTiers.Value:=6;
        if CompteurTiers.Text='' then CompteurTiers.Text:='0';
      end
      else
      begin
        TCheckBox(GetFromNam('SO_GCFORMATCODETIERS',FF)).Checked := false;
        Active := False;
        ActiveFormat := false;
      end;
    end;
    if not ActiveFormat then
    begin
      if Not (ctxChr in V_PGI.PGIContexte) then
      begin
        LgNumTiers.Value:=17;
        THEdit(GetFromNam ('SO_GCPREFIXETIERS', FF)).Text := '';
      end;
    end;
  end;
  {$ENDIF PGIMAJVER}
  { GC_DBR_CRM10764_FIN }
  if ExisteNam('SO_GCNUMTIERSAUTO',FF) then
  begin
    SetEna('SO_GCFORMATCODETIERS',FF,Active) ;
    SetEna('SO_GCPREFIXETIERS',FF,ActiveFormat) ;
    SetEna('LSO_GCPREFIXETIERS',FF,ActiveFormat) ;
    SetEna('SO_GCLGNUMTIERS',FF,ActiveFormat) ;
    SetEna('LSO_GCLGNUMTIERS',FF,ActiveFormat) ;
    SetEna('SO_GCCOMPTEURTIERS',FF,Active) ;
    SetEna('LSO_GCCOMPTEURTIERS',FF,Active) ;
    SetEna ('LSO_GCPREFIXEAUXI', FF, Active);
    SetEna ('SO_GCPREFIXEAUXI', FF, Active);
    SetEna ('LSO_GCPREFIXEAUXIFOU', FF, Active);
    SetEna ('SO_GCPREFIXEAUXIFOU', FF, Active);
  end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure GereAnaParamAff(CC : TControl);
Var FF : TForm ;
    AxeAna, AxeAnaAff : TCheckBox;
begin
  FF := GetLaForm(CC);
  if (not ExisteNam('SO_GCAXEANALYTIQUE', FF)) or (not ExisteNam('SO_GCAXEANAAFF', FF)) then exit;
    AxeAna := TCheckBox(GetFromNam('SO_GCAXEANALYTIQUE',FF));
  AxeAnaAff := TCheckBox(GetFromNam('SO_GCAXEANAAFF',FF));
  if AxeAna.Checked then
  begin
    SetEna('SO_GCAXEANAAFF', FF, True);
  end else
  begin
    SetEna('SO_GCAXEANAAFF', FF, False);
    AxeAnaAff.Checked := false;
  end;
{$IFNDEF STK}
  SetInvi('SO_GCINVPERM', FF);
  SetInvi('SO_GCJALSTOCK', FF);
  SetInvi('LSO_GCJALSTOCK', FF);
  SetInvi('SO_GCTYPECPTASTOCK', FF);
  SetInvi('LSO_GCTYPECPTASTOCK', FF);
  SetInvi('SO_GCMODEVALOSTOCK', FF);
  SetInvi('LSO_GCMODEVALOSTOCK', FF);
  SetInvi('LSCO_GCGRPINVPERM', FF);
  SetInvi('SCO_GCGRPINVPERM', FF);
{$ENDIF STK}
end;

{ GC_JTR_GC15601_Début }
{ Gestion de l'affichage du mot de passe pour forcer la marge mini (police + cacher/montrer) }
procedure GereMdpMargeMini(CC : TControl);
var FF : TForm;
    Gere, Gere1 : boolean;
begin
  FF := GetLaForm(CC);
  if ExisteNam('SO_FORCEMARGEMINI', FF) then
  begin
    THEdit(GetFromNam('SO_MDPMARGEMINI', FF)).PasswordChar := '*';
    Gere := (TCheckBox(GetFromNam('SO_GCBLOQUEMARGE',FF)).Checked);
    if not Gere then
      TCheckBox(GetFromNam('SO_FORCEMARGEMINI',FF)).Checked := false;
    Gere1 := (TCheckBox(GetFromNam('SO_FORCEMARGEMINI',FF)).Checked);
    TCheckBox(GetFromNam('SO_FORCEMARGEMINI', FF)).Enabled := Gere;
    THEdit(GetFromNam('SO_MDPMARGEMINI', FF)).Enabled := Gere1;
    THLabel(GetFromNam('LSO_MDPMARGEMINI', FF)).Enabled := Gere1;
  end;
end;
{ GC_JTR_GC15601_Fin }

procedure GereTypeTVAIntraComm(CC : TControl); // JTR - TVA intracommunautaire
Var FF : TForm ;
    Gere : TCheckBox;
begin
  FF := GetLaForm(CC);
  if (not ExisteNam('SO_GERETVAINTRACOMM', FF))
     or (not ExisteNam('SO_TYPETVAINTRACOMM', FF))
     or (not ExisteNam('SO_PCETVAINTRACOMM', FF)) then exit;
  Gere := TCheckBox(GetFromNam('SO_GERETVAINTRACOMM',FF));
  SetEna('SO_TYPETVAINTRACOMM', FF, Gere.Checked);
  SetEna('SO_PCETVAINTRACOMM', FF, Gere.Checked);
  if not Gere.Checked then
  begin
    THValComboBox(GetFromNam('SO_TYPETVAINTRACOMM', FF)).Value := '';
    THMultiValComboBox(GetFromNam('SO_PCETVAINTRACOMM', FF)).Value := '';
  end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure GereContremarque(CC : TControl);
var FF : TForm;
    Gere, PieceAuto, ForceAuto, PieceAutoVente : Boolean;
    LePlus : string;

    procedure EnaChpEtLib(Lequel : string; Quoi : boolean);
    begin
      SetEna(Lequel , FF, Quoi);
      SetEna('L' + Lequel , FF, Quoi);
    end;

begin
  FF := GetLaForm(CC);
  if (not ExisteNam('SO_GERECONTREMARQUE', FF)) then exit;
  { En attendant de savoir gérer }
  { DBR : ouverture de la finesse de la contremarque - A TESTER }
  (*
  if ExisteNam('SO_GCAFFCONTREM', FF) then
  begin
    SetInvi ('LSO_GCAFFCONTREM', FF);
    SetInvi ('SO_GCAFFCONTREM', FF);
  end;
  SetInvi('SO_GCFORCEQTEAUTOCONTREM', FF);
  *)
  { Fin - En attendant de savoir gérer }
  Gere := (TCheckBox(GetFromNam('SO_GERECONTREMARQUE',FF)).Checked);
  PieceAuto := ((Gere) and (TCheckBox(GetFromNam('SO_GCPCEAUTOCONTREM',FF)).Checked));
  ForceAuto := ((PieceAuto) and (TCheckBox(GetFromNam('SO_GCFORCEAUTOCONTREM',FF)).Checked));
  PieceAutoVente := (Gere and (TCheckBox(GetFromNam('SO_GCAUTOCONTREM',FF)).Checked));
  EnaChpEtLib('SO_GCNATRECCONTREM', PieceAutoVente);
  EnaChpEtLib('SO_GCNATLIVCONTREM', PieceAutoVente);
  SetEna('SO_GCAUTOCONTREM' , FF, Gere);
  SetEna('SO_GCAFFCONTREM', FF, Gere);
  SetEna('SO_MAILCONTREMARQUE', FF, Gere);
  SetEna('SO_GCPCEAUTOCONTREM', FF, Gere);
  { GC_DBR_TSO_DEBUT }
  SetEna('SO_TRADUCCONTREMARQUE', FF, Gere);
  SetEna('SO_AFFAIRECONTREMARQUE', FF, Gere);
  SetEna('SO_ANALYTIQUECONTREMARQUE', FF, Gere);
  SetEna('SO_ADRLIVCONTREMARQUE', FF, Gere);
  SetEna('SO_REPORTPRIXCONTREMARQUE', FF, Gere);
  { GC_DBR_TSO_FIN }
  SetEna('SO_GCFORCEAUTOCONTREM', FF, PieceAuto);
  EnaChpEtLib('SO_GCNATACHAUTOCONTREM', ForceAuto);
  if not Gere then
  begin
    TCheckBox(GetFromNam('SO_MAILCONTREMARQUE', FF)).Checked := false;
    TCheckBox(GetFromNam('SO_GCPCEAUTOCONTREM', FF)).Checked := false;
    TCheckBox(GetFromNam('SO_GCFORCEAUTOCONTREM', FF)).Checked := false;
  end;
  if not PieceAuto then
    TCheckBox(GetFromNam('SO_GCFORCEAUTOCONTREM', FF)).Checked := false;
  if not ForceAuto then
    THValComboBox(GetFromNam('SO_GCNATACHAUTOCONTREM', FF)).Value := '';
  LePlus := 'AND (SELECT GSN_STKTYPEMVT FROM STKNATURE WHERE GSN_QUALIFMVT=GPP_STKQUALIFMVT)="PHY"';
  if ExisteNam('SO_GCNATRECCONTREM', FF) then
    THValComboBox(GetFromNam('SO_GCNATRECCONTREM',FF)).Plus := LePlus;
  if ExisteNam('SO_GCNATLIVCONTREM', FF) then
    THValComboBox(GetFromNam('SO_GCNATLIVCONTREM',FF)).Plus := LePlus;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure GereFarFae(CC : TControl; Sauve : boolean=false);
var FF : TForm;
    Gere : Boolean;
begin
  FF := GetLaForm(CC);
  if (not ExisteNam('SO_GEREFARFAE', FF)) then exit;
  Gere := (TCheckBox(GetFromNam('SO_GEREFARFAE',FF)).Checked);
  SetEna('SO_JALFARFAE_VTE', FF, Gere);
  SetEna('SO_JALFARFAE_ACH', FF, Gere);
  if not Gere then
  begin
    THValComboBox(GetFromNam('SO_JALFARFAE_VTE', FF)).Value := '';
    THValComboBox(GetFromNam('SO_JALFARFAE_ACH', FF)).Value := '';
  end else
  if Sauve then
  begin
    if (THValComboBox(GetFromNam('SO_JALFARFAE_VTE', FF)).Value = '')
       or (THValComboBox(GetFromNam('SO_JALFARFAE_ACH', FF)).Value = '') then
      HShowmessage('31;Société;ATTENTION. Le ou les journaux de gestion des FAR/FAE ne sont pas renseigné(s).;W;O;O;O;','','');
  end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure GereDevisePiece(CC : TControl);
var FF : TForm;
    Gere : boolean;
begin
  FF := GetLaForm(CC);
  if (ExisteNam('SCO_PIECEDEVISE', FF)) then
  begin
    Gere := GetParamSocSecur('SO_GCTAUXNEGOCIE', false);
    SetEna('SCO_PIECEDEVISE', FF, Gere); SetEna('LSCO_PIECEDEVISE', FF, Gere);
    SetEna('SO_TYPETAUXDEVISE', FF, Gere); SetEna('LSO_TYPETAUXDEVISE', FF, Gere);
    if THValComboBox(GetFromNam('SO_TYPETAUXDEVISE', FF)).Value = 'TCH' then
    begin
      SetEna('SO_GENEAUTOCHANCBLOQUE', FF, True);
    end else
    begin
      ThCheckBox(GetFromNam('SO_GENEAUTOCHANCBLOQUE', FF)).Checked := false;
      SetEna('SO_GENEAUTOCHANCBLOQUE', FF, False);
    end;
  end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure GereElimineLigne0(CC : TControl);
var FF : TForm;
begin
  FF := GetLaForm(CC);
  if (ExisteNam('SO_GCELIMINEVENLIGNE0', FF) or ExisteNam('SO_GCELIMINEACHLIGNE0', FF)) then
  begin
    SetEna('SO_GCNATUREVENLIGNE0', FF, TCheckBox(GetFromNam('SO_GCELIMINEVENLIGNE0',FF)).Checked);
    SetEna('SO_GCNATUREACHLIGNE0', FF, TCheckBox(GetFromNam('SO_GCELIMINEACHLIGNE0',FF)).Checked);
  end;
end;
{$ENDIF !EAGLSERVER}

{$IFNDEF EAGLSERVER}
Function PrefixeTiersOk ( CC : TControl ) : boolean;
Var FF : TForm ;
    TEdit_Prefixe : TEdit ;
    Prefixe : String;
const
    CaractereAutorise : String ='ABCDEFGHIJKLMNOPQRSTUVWXYZ' ;
BEGIN
Result:=True;
FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_GCPREFIXETIERS',FF) then Exit ;
TEdit_Prefixe:=TEdit(GetFromNam('SO_GCPREFIXETIERS',FF));
Prefixe:=UpperCase(Trim(TEdit_Prefixe.Text));
TEdit_Prefixe.Text := Prefixe;
Result := ControlePrefixe (Prefixe, CaractereAutorise);
if Result = False then
   begin
   HShowMessage('25;Société;Le préfixe doit être alphabétique.;W;O;O;O;','','') ;
   if TEdit_Prefixe.CanFocus then TEdit_Prefixe.SetFocus ;
   Exit;
   end;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{ GC_DBR_CRM10764_DEB }
Function ControleFormatCodeTiers (FFGC : TForm = nil; FFCP : TForm = nil) : boolean;
var
  CodeTiersMini, CodeTiersMiniB : string;
  lInfoCpta : TInfoCpta ;
  LgNumTiers   : Integer ;
  CompteurTiers: String ;
  Prefixe      : String ;
  bFormatCodeTiers : boolean;
  BourreAux : string;
  LgAuxi : integer;
  LgPrefixeAuxi : integer;

  Function BourreLaDoncLocal (St : String) : string ;
  var
    ll, i : Integer ;
  begin
    Result:=St ;
    ll := Length(Result) ;

    If ll<LgAuxi then for i:=ll+1 to LgAuxi do Result:=Result+BourreAux ;
  end ;

begin
  Result := true;
  if FFGC <> nil then bFormatCodeTiers := TCheckBox(GetFromNam('SO_GCFORMATCODETIERS',FFGC)).Checked
  else bFormatCodeTiers := GetParamSocSecur ('SO_GCFORMATCODETIERS', false);
  if bFormatCodeTiers then
  begin
    if FFCP = nil then
    begin
      lInfoCpta := GetInfoCpta(fbaux) ;
      LgAuxi := lInfoCpta.lg;
      BourreAux := lInfoCpta.Cb;
    end else
    begin
      BourreAux := TEdit(GetFromNam('SO_BOURREAUX',FFCP)).Text;
      lgAuxi := TSpinEdit(GetFromNam('SO_LGCPTEAUX',FFCP)).Value;
    end;

    if FFGC <> nil then
    begin
      LgNumTiers:=TSpinEdit(GetFromNam('SO_GCLGNUMTIERS',FFGC)).Value;
      CompteurTiers:=TEdit(GetFromNam('SO_GCCOMPTEURTIERS',FFGC)).Text;
      Prefixe := TEdit(GetFromNam('SO_GCPREFIXETIERS',FFGC)).Text;
      LgPrefixeAuxi := Min (Length (TEdit (GetFromNam ('SO_GCPREFIXEAUXI', FFGC)).Text), Length (TEdit (GetFromNam ('SO_GCPREFIXEAUXIFOU', FFGC)).Text));
    end else
    begin
      LgNumTiers := GetParamSocSecur ('SO_GCLGNUMTIERS', 0);
      CompteurTiers := GetParamSocSecur ('SO_GCCOMPTEURTIERS', '0');
      Prefixe := GetParamSocSecur ('SO_GCPREFIXETIERS', '');
      LgPrefixeAuxi := Min (Length (GetParamSocSecur ('SO_GCPREFIXEAUXI', '')), Length (GetParamSocSecur ('SO_GCPREFIXEAUXIFOU', '')));
    end;

    if lgNumTiers + LgPrefixeAuxi > LgAuxi + Length (Prefixe) then
    begin
      HShowMessage('30;Société;La longueur maxi du code tiers doit être ' + IntToStr (LgAuxi + Length (Prefixe) - LgPrefixeAuxi) + '.;W;O;O;O;','','');
      Result := false;
    end else
    begin
      if IsNumeric (BourreAux) then
      begin
        if (BourreAux = '0') or (BourreAux = '1') then
          CodeTiersMiniB := BourreLaDoncLocal ('1')
        else CodeTiersMiniB := BourreLaDoncLocal (BourreAux);
        CodeTiersMiniB := copy (CodeTiersMiniB, 1, LgNumTiers - Length (Prefixe));

        CodeTiersMini := BourreLaDoncLocal (CompteurTiers);
        CodeTiersMini := copy (CodeTiersMini, 1, LgNumTiers - Length (Prefixe));
        if CompteurTiers <> CodeTiersMini then
          Result:=False
        else if StrToFloat (CodeTiersMini) < StrToFloat (CodeTiersMiniB) then
          Result:=False;

        if not Result then
          HShowMessage('30;Société;Le chrono mimum doit être ' + CodeTiersMiniB + '.;W;O;O;O;','','') ;
      end;
    end;
  end;
end;
{ GC_DBR_CRM10764_FIN }

Function LgNumTiersOk ( CC : TControl ) : boolean;
Var FF : TForm ;
    NumTiersAuto : TCheckBox ;
    Prefixe      : String ;
    LgNumTiers   : TSpinEdit ;
    CompteurTiers: TEdit ;
    LgCompteur, LgPrefixe : Integer;
BEGIN
  Result:=True;
  FF:=GetLaForm(CC) ;
  if Not ExisteNam('SO_GCNUMTIERSAUTO',FF) then Exit ;
  NumTiersAuto:=TCheckBox(GetFromNam('SO_GCNUMTIERSAUTO',FF)) ;
  if Not NumTiersAuto.Checked then Exit ;

  Prefixe := TEdit(GetFromNam('SO_GCPREFIXETIERS',FF)).Text;
  LgPrefixe := Length(Prefixe);
  LgNumTiers:=TSpinEdit(GetFromNam('SO_GCLGNUMTIERS',FF));
  CompteurTiers:=TEdit(GetFromNam('SO_GCCOMPTEURTIERS',FF));
  CompteurTiers.Text:=Trim(IntToStr(StrToInt64(CompteurTiers.Text)));
  if CompteurTiers.Text='' then CompteurTiers.Text:='0';
  LgCompteur:=Length(CompteurTiers.Text) ;

  Result:=ControleChrono (CompteurTiers.Text);     // Contrôle de numéricité
  if Result=False then
  begin
    if CompteurTiers.CanFocus then CompteurTiers.SetFocus ;
    Exit;
  end;
  Result:=ControleLgCode ( 4, 17, LgNumTiers.Value, LgPrefixe, LgCompteur);
  if Result=False then
  begin
    if LgNumTiers.CanFocus then LgNumTiers.SetFocus ;
  end;
  { GC_DBR_CRM10764 }
  if Result then Result := ControleFormatCodeTiers (FF);
end;
{$ENDIF EAGLSERVER}

//
// Attribution automatique du code Article
//
{$IFNDEF EAGLSERVER}
Procedure CodeArtAutoX ( CC : TControl; ForceActive : Boolean ) ;
Var FF : TForm ;
    NumArtAuto : TCheckBox ;
    LgNumArt   : TSpinEdit ;
    CompteurArt: TEdit ;
    Active     : Boolean ;
BEGIN
FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_GCNUMARTAUTO',FF) then Exit ;
  if ForceActive=True then Active:=True
  else
  begin
    NumArtAuto:=TCheckBox(GetFromNam('SO_GCNUMARTAUTO',FF)) ;
    if NumArtAuto.Checked then
    begin
      Active := True;
      LgNumArt:=TSpinEdit(GetFromNam('SO_GCLGNUMART',FF));
      if Not (ctxChr in V_PGI.PGIContexte) then
        LgNumArt.Value:=18
      else if (LgNumArt.Value=0) then
        LgNumArt.Value:=5;
      CompteurArt:=TEdit(GetFromNam('SO_GCCOMPTEURART',FF));
      if CompteurArt.Text='' then CompteurArt.Text:='0';
    end
    else Active := False;
  end;
  SetEna('SO_GCPREFIXEART',FF,Active) ;
  SetEna('LSO_GCPREFIXEART',FF,Active) ;
  SetEna('SO_GCLGNUMART',FF,Active) ;
  SetEna('LSO_GCLGNUMART',FF,Active) ;
  SetEna('SO_GCCOMPTEURART',FF,Active) ;
  SetEna('LSO_GCCOMPTEURART',FF,Active) ;

  {$IFDEF GPAO}
  // Si numérotation auto -> Saisie de l'article affaire en "manuel" obligatoire
  if (TCheckBox(GetFromNam('SO_GCNUMARTAUTO',FF)).Checked) and (GetParamSocSecur('SO_AFSAISIEART', 'MAN')<>'MAN') then
    SetParamSoc('SO_AFSAISIEART', 'MAN');
  {$ENDIF GPAO}
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Function PrefixeArtOk ( CC : TControl ) : boolean;
Var FF : TForm ;
    TEdit_Prefixe : TEdit ;
    Prefixe : String;
const
{$IFDEF BTP}
    CaractereAutorise : String ='ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%_' ;
{$ELSE}
    CaractereAutorise : String ='ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%' ;
{$ENDIF}
BEGIN
Result:=True;
FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_GCPREFIXEART',FF) then Exit ;
TEdit_Prefixe:=TEdit(GetFromNam('SO_GCPREFIXEART',FF));
Prefixe:=UpperCase(Trim(TEdit_Prefixe.Text));
TEdit_Prefixe.Text := Prefixe;
Result := ControlePrefixe (Prefixe, CaractereAutorise);
if Result = False then
   begin
   HShowMessage('25;Société;Le préfixe doit être alphabétique.;W;O;O;O;','','') ;
   if TEdit_Prefixe.CanFocus then TEdit_Prefixe.SetFocus ;
   Exit;
   end;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Function LgCodeArtOk ( CC : TControl ) : boolean;
Var FF : TForm ;
    NumArtAuto : TCheckBox ;
    Prefixe    : String ;
    LgNumArt   : TSpinEdit ;
    CompteurArt: TEdit ;
    LgCompteur, LgPrefixe : Integer;
BEGIN
Result:=True;
FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_GCNUMARTAUTO',FF) then Exit ;
NumArtAuto:=TCheckBox(GetFromNam('SO_GCNUMARTAUTO',FF)) ;
if Not NumArtAuto.Checked then Exit ;
Prefixe := TEdit(GetFromNam('SO_GCPREFIXEART',FF)).Text;
LgPrefixe := Length(Prefixe);
LgNumArt:=TSpinEdit(GetFromNam('SO_GCLGNUMART',FF));
CompteurArt:=TEdit(GetFromNam('SO_GCCOMPTEURART',FF));
CompteurArt.Text:=Trim(CompteurArt.Text);
if CompteurArt.Text='' then CompteurArt.Text:='0';
LgCompteur:=Length(CompteurArt.Text) ;
Result:=ControleChrono (CompteurArt.Text);        // Contrôle de numéricité
if Result=False then
   begin
   if CompteurArt.CanFocus then CompteurArt.SetFocus ;
   Exit;
   end;
Result:=ControleLgCode ( 4, 18, LgNumArt.Value, LgPrefixe, LgCompteur);
if Result=False then
   begin
   if LgNumArt.CanFocus then LgNumArt.SetFocus ;
   end;
END;
{$ENDIF EAGLSERVER}

//
// Attribution automatique du Code à barres
//
{$IFNDEF EAGLSERVER}
Procedure NumCABAutoX ( CC : TControl ) ;
Var FF : TForm ;
    NumCABAuto : TCheckBox ;
BEGIN
FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_GCNUMCABAUTO',FF) then Exit ;
NumCABAuto:=TCheckBox(GetFromNam('SO_GCNUMCABAUTO',FF)) ;
if NumCABAuto.Checked then
   begin
   // On vérifie que le paramétrage par défaut du code à barres a bien été créé
   If (Not ExisteSQL('Select * from CODEBARRES where (GCB_NATURECAB="SO") and (GCB_IDENTIFCAB="...")')) then
     begin
     AGLLanceFiche('GC','GCCODEBARRES','SO;...','','ACTION=CREATION') ;
     If (Not ExisteSQL('Select * from CODEBARRES where (GCB_NATURECAB="SO") and (GCB_IDENTIFCAB="...")')) then
       NumCABAuto.Checked:=False;
     end;
   end;
END ;
{$ENDIF EAGLSERVER}

//
// Attribution automatique du Code à barres Prestation
//
// DEBUT CCMX-CEGID FQ N° GC14064 ajout SO_GCPRCABAUTO
{$IFNDEF EAGLSERVER}
Procedure PrCABAutoX ( CC : TControl ) ;
Var FF : TForm ;
    PrCABAuto : TCheckBox ;
BEGIN
  FF:=GetLaForm(CC) ;
  if (Not ExisteNam('SO_GCPRCABAUTO',FF)) then Exit ;
  PrCABAuto:=TCheckBox(GetFromNam('SO_GCPRCABAUTO',FF));

  if PrCABAuto.Checked then
  begin
    // On vérifie que le paramétrage par défaut du code à barres a bien été créé
    If (Not ExisteSQL('Select * from CODEBARRES where (GCB_NATURECAB="SO") and (GCB_IDENTIFCAB="---")')) then
    begin
      AGLLanceFiche('GC','GCCODEBARRES','SO;---','','ACTION=CREATION') ;
      If (Not ExisteSQL('Select * from CODEBARRES where (GCB_NATURECAB="SO") and (GCB_IDENTIFCAB="---")')) then
        PrCABAuto.Checked:=False;
    end;
  end;
END ;
{$ENDIF EAGLSERVER}
// FIN CCMX-CEGID FQ N° GC14064 ajout SO_GCPRCABAUTO
// Gestion des familles articles 4 à 8
//
{$IFNDEF EAGLSERVER}
Procedure FamillesArticle4a8 ( CC : TControl ) ;
var FF : TForm ;
    ArtLookOrli : TCheckBox ;
    iFam : integer ;
    procedure UtilSocInsertChoixCode (stType,stCode,stLibelle,stAbrege,stlibre : string);
    var Q : TQuery;
    begin
      Q:=OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="'+stType+'" AND CC_CODE="'+stCode+'"',FALSE) ;
      If Q.EOF Then
        BEGIN
        Q.Insert ;
        Q.FindField('CC_TYPE').AsString:=stType ;
        Q.FindField('CC_CODE').AsString:=stCode ;
        Q.FindField('CC_LIBELLE').AsString:=stLibelle ;
        Q.FindField('CC_ABREGE').AsString:=stAbrege ;
        Q.FindField('CC_LIBRE').AsString:=stLibre ;
        Q.Post ;
        END ;
      Ferme(Q) ;
    end;

BEGIN
FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_ARTLOOKORLI',FF) then Exit ;
ArtLookOrli:=TCheckBox(GetFromNam('SO_ARTLOOKORLI',FF)) ;
if ArtLookOrli.Checked then
  // Insertion des familles article 4 à 8 dans la tablette des libellés des familles GCLIBFAMILLE
  for iFam:=4 to 8 do UtilSocInsertChoixCode ('GLF','LF'+intTostr(iFam),'Famille niveau '+intTostr(iFam),'Famille '+intTostr(iFam),'') ;
END ;
{$ENDIF EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/03/2002
Modifié le ... : 21/03/2002
Description .. : LG* Controle la longueur des zones debit et credit de
Suite ........ : l'onglet
Suite ........ : lettrage gestion
Mots clefs ... :
*****************************************************************}
procedure ControleLgNum ( CC : TControl ) ;
Var v : double;
begin
v:=Valeur(THNumEdit(CC).Text);
if v > 9999 then THNumEdit(CC).Text:='9999' ;
end;
// GP
(* GP le 9/10/2002 : mis en commentaire par sécurité. (IMMP PGI) A Réactiver dès que possible *)
{$IFNDEF EAGLSERVER}
procedure VerifAmorJourSortie ( CC : TControl ) ;
var FF : TForm ;
    CB : tCheckBox ;
BEGIN
If Not (ctxImmo in V_PGI.PGIContexte) then Exit ;
FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_I_CALCULAMORTJOURSORTIE',FF) then Exit ;
CB:=TCheckBox(GetFromNam('SO_I_CALCULAMORTJOURSORTIE',FF)) ;
If CB<>NIL Then
  BEGIN
  If TCheckBox(CC).Checked Then CB.Enabled:=TRUE Else
    BEGIN
    CB.Enabled:=FALSE ;
    CB.Checked:=FALSE ;
    END ;
  END ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure VerifCBLocal ( CC : TControl ) ;
var FF : TForm ;
    SuiviLocal,
    SuiviEtab,
    SuiviCP,
    SuiviLieu,
    SuiviVille,
    SuiviPays : TCheckBox ;
begin
// Nathalie Payrot le 02/08/02
// gestion des checkbox détails localisation en fonction du suivi ou non de la localisation
FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_I_LOCALISATION',FF) then Exit ;
SuiviLocal:=TCheckBox(GetFromNam('SO_I_LOCALISATION',FF)) ;
if not SuiviLocal.Checked then
   begin
     // non suivi de la localisation : décoche du détail (lieu, étab, ville ...)
     // on rend non saisissables ces zones.
     SuiviEtab := TCheckBox(GetFromNam('SO_I_LOCAL_CODEETAB',FF)) ;
     SuiviEtab.Checked:=False;
     SetEna('SO_I_LOCAL_CODEETAB', FF,False ) ;
     SuiviCP := TCheckBox(GetFromNam('SO_I_LOCAL_CODEPOSTAL',FF)) ;
     SuiviCP.Checked:=False;
     SetEna('SO_I_LOCAL_CODEPOSTAL', FF,False ) ;
     SuiviLieu := TCheckBox(GetFromNam('SO_I_LOCAL_CODELIEU',FF)) ;
     SuiviLieu.Checked:=False;
     SetEna('SO_I_LOCAL_CODELIEU', FF,False ) ;
     SuiviVille := TCheckBox(GetFromNam('SO_I_LOCAL_VILLE',FF)) ;
     SuiviVille.Checked:=False;
     SetEna('SO_I_LOCAL_VILLE', FF,False ) ;
     SuiviPays := TCheckBox(GetFromNam('SO_I_LOCAL_PAYS',FF)) ;
     SuiviPays.Checked:=False;
     SetEna('SO_I_LOCAL_PAYS', FF,False ) ;
   end
   else
     begin
     // suivi de la localisation : on rend accessible le détail (lieu, etab, ville ...)
     SetEna('SO_I_LOCAL_CODEETAB', FF,True ) ;
     SetEna('SO_I_LOCAL_CODEPOSTAL', FF,True ) ;
     SetEna('SO_I_LOCAL_CODELIEU', FF,True ) ;
     SetEna('SO_I_LOCAL_VILLE', FF,True ) ;
     SetEna('SO_I_LOCAL_PAYS', FF,True ) ;
     end ;
end ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure VerifFinancementInfoLibre(CC : TControl ) ; // le financement
var FF : TForm ;
    SuiviLibre, FiLibre  : tCheckBox ;
    i : integer ;
begin
FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_I_FISUIVILIBRE',FF) then Exit ;
SuiviLibre:=TCheckBox(GetFromNam('SO_I_FISUIVILIBRE',FF)) ;
if not SuiviLibre.Checked then
  begin
  for i:=1 to 10 do
    begin
    FiLibre := TCheckBox(GetFromNam('SO_I_FILIBRE'+InttoStr(i),FF)) ;
    FiLibre.Checked:=False ;
    SetEna('SO_I_FILIBRE'+IntToStr(i), FF, False ) ;
    SetEna('SO_I_FILIBINT'+IntToStr(i), FF, False ) ;
    SetEna('LSO_I_FILIBINT'+IntToStr(i), FF, False ) ;
    end ;
  end else
  begin
  for i:=1 to 10 do
    begin
    SetEna('SO_I_FILIBRE'+IntToStr(i), FF, True ) ;
    SetEna('SO_I_FILIBINT'+IntToStr(i), FF, True ) ;
    SetEna('LSO_I_FILIBINT'+IntToStr(i), FF, True ) ;
    end ;
  end ;
end ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure VerifCBAutreInfo ( CC : TControl ) ;
var FF : TForm ;
    SuiviAutreInfo,
    SuiviLibre1, SuiviLibre2,
    SuiviLibre3, SuiviLibre4,
    SuiviPointage,
    SuiviEntRep :TCheckBox ;
begin
// Nathalie Payrot le 02/08/02
// gestion des checkbox autres critères libres en fonction du
// suivi ou non des autres infos de la fiche immos

FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_I_AUTREINFO',FF) then Exit ;
SuiviAutreInfo:=TCheckBox(GetFromNam('SO_I_AUTREINFO',FF)) ;
if not SuiviAutreInfo.Checked then
   begin
     SuiviLibre1 := TCheckBox(GetFromNam('SO_I_AUTREINFOCRITERELIBRE1',FF)) ;
     SuiviLibre1.Checked:=False;
     SetEna('SO_I_AUTREINFOCRITERELIBRE1', FF,False ) ;
     SetEna('LSO_I_AUTREINFOLIBRE1LIB', FF,False ) ;
     SetEna('SO_I_AUTREINFOLIBRE1LIB', FF,False ) ;
     SuiviLibre2 := TCheckBox(GetFromNam('SO_I_AUTREINFOCRITERELIBRE2',FF)) ;
     SuiviLibre2.Checked:=False;
     SetEna('SO_I_AUTREINFOCRITERELIBRE2', FF,False ) ;
     SetEna('LSO_I_AUTREINFOLIBRE2LIB', FF,False ) ;
     SetEna('SO_I_AUTREINFOLIBRE2LIB', FF,False ) ;
     SuiviLibre3 := TCheckBox(GetFromNam('SO_I_AUTREINFOCRITERELIBRE3',FF)) ;
     SuiviLibre3.Checked:=False;
     SetEna('SO_I_AUTREINFOCRITERELIBRE3', FF,False ) ;
     SetEna('LSO_I_AUTREINFOLIBRE3LIB', FF,False ) ;
     SetEna('SO_I_AUTREINFOLIBRE3LIB', FF,False ) ;
     SuiviLibre4 := TCheckBox(GetFromNam('SO_I_AUTREINFOCRITERELIBRE4',FF)) ;
     SuiviLibre4.Checked:=False;
     SetEna('SO_I_AUTREINFOCRITERELIBRE4', FF,False ) ;
     SetEna('LSO_I_AUTREINFOLIBRE4LIB', FF,False ) ;
     SetEna('SO_I_AUTREINFOLIBRE4LIB', FF,False ) ;
     SuiviPointage := TCheckBox(GetFromNam('SO_I_AUTREINFOPOINTAGE',FF)) ;
     SuiviPointage.Checked:=False;
     SetEna('SO_I_AUTREINFOPOINTAGE', FF,False ) ;
     SuiviEntRep := TCheckBox(GetFromNam('SO_I_ENTRETIEN',FF)) ;
     SuiviEntRep.Checked:=False;
     SetEna('SO_I_ENTRETIEN', FF,False ) ;
   end
 else
   begin
     SetEna('SO_I_AUTREINFOCRITERELIBRE1', FF,True ) ;
     SetEna('LSO_I_AUTREINFOLIBRE1LIB', FF,True ) ;
     SetEna('SO_I_AUTREINFOLIBRE1LIB', FF,True ) ;
     SetEna('SO_I_AUTREINFOCRITERELIBRE2', FF,True ) ;
     SetEna('LSO_I_AUTREINFOLIBRE2LIB', FF,True ) ;
     SetEna('SO_I_AUTREINFOLIBRE2LIB', FF,True ) ;
     SetEna('SO_I_AUTREINFOCRITERELIBRE3', FF,True ) ;
     SetEna('LSO_I_AUTREINFOLIBRE3LIB', FF,True ) ;
     SetEna('SO_I_AUTREINFOLIBRE3LIB', FF,True ) ;
     SetEna('SO_I_AUTREINFOCRITERELIBRE4', FF,True ) ;
     SetEna('LSO_I_AUTREINFOLIBRE4LIB', FF,True ) ;
     SetEna('SO_I_AUTREINFOLIBRE4LIB', FF,True ) ;
     SetEna('SO_I_AUTREINFOPOINTAGE', FF,True ) ;
     SetEna('SO_I_ENTRETIEN', FF,True ) ;
   end ;

end ;
{$ENDIF EAGLSERVER}

{$IFDEF COMPTA}
//SG6 07/12/2004 Gestion des paramSoc de la comptabilité
procedure GereCompta(FF : TForm);
var
  SO_CROISAXE : TCheckBox;
  i : integer;
begin
  if ctxPCL in V_PGI.PGIContexte then
    if ExisteNam('SO_CPSPECIFIQUES',FF) then
      SetInvi('SO_CPSPECIFIQUES', FF)
  else
    if ExisteNam('SO_CPSPECIFIQUES',FF) then SetVisi('SO_CPSPECIFIQUES', FF) ;

  if ( not ( ctxPCL in V_PGI.PGIContexte ) ) and ExisteNam('SO_CPTVARECDEP',FF) then
   SetInvi('SO_CPTVARECDEP', FF) ;
  {YMO 20/02/2007 Gestion automatique non modifiable en PCL}
  if ( ctxPCL in V_PGI.PGIContexte ) and ExisteNam('SO_AUTOTP',FF) then
   SetInvi('SO_AUTOTP', FF) ;

  if TControl(FF.FindComponent('SO_CROISAXE'))<>nil then
    begin
    SO_CROISAXE:=TCheckBox(GetFromNam('SO_CROISAXE',FF));
    SO_CROISAXE.Enabled := false;
    for i:=1 to MaxAxe do
      begin
      SetEna('SO_VENTILA'+IntToStr(i),FF,false);
      end;
    end ;

  if ExisteNam('SO_JALMULTIETAB',FF) then
    begin
    SetInvi('SO_JALMULTIETAB', FF) ;
    SetInvi('LSO_JALMULTIETAB', FF) ;
    end ;
  if ExisteNam('SCO_SAISIEPARAM',FF) then
    begin
    SetInvi('SCO_SAISIEPARAM', FF) ;
    SetInvi('LSCO_SAISIEPARAM', FF) ;
    end ;

  // GCO - 15/03/2007
  if ExisteNam('SO_CPPLANREVISION', FF) then
    SetEna( 'SO_CPPLANREVISION', FF, not VH^.OkModRIC);

  if ExisteNam('SO_CPREVISBLOQUEGENE', FF) then
    SetEna( 'SO_CPREVISBLOQUEGENE', FF, not VH^.OkModRIC);
  // FIN GCO

{$IFNDEF EAGLSERVER}
  if ExisteNam('SCO_CPMASQUECRIT',FF) then
    GereMasqueCrit( FF ) ;
{$ENDIF EAGLSERVER}

  // GCO - 03/10/2007 - FQ 21570
  // GCO - 29/01/2008 - FQ 22169
  if ExisteNam('SO_OUVREBIL', FF) then
  begin
    THCritMaskEdit(GetFromNam('SO_OUVREBIL', FF)).Plus   := ' AND G_LETTRABLE = "-" AND G_VENTILABLE = "-"';
    THCritMaskEdit(GetFromNam('SO_FERMEBIL', FF)).Plus   := ' AND G_LETTRABLE = "-" AND G_VENTILABLE = "-"';
    THCritMaskEdit(GetFromNam('SO_OUVREPERTE', FF)).Plus := ' AND G_LETTRABLE = "-" AND G_VENTILABLE = "-"';
    THCritMaskEdit(GetFromNam('SO_FERMEPERTE', FF)).Plus := ' AND G_LETTRABLE = "-" AND G_VENTILABLE = "-"';
    THCritMaskEdit(GetFromNam('SO_RESULTAT', FF)).Plus   := ' AND G_LETTRABLE = "-" AND G_VENTILABLE = "-"';
    THCritMaskEdit(GetFromNam('SO_OUVREBEN', FF)).Plus   := ' AND G_LETTRABLE = "-" AND G_VENTILABLE = "-"';
    THCritMaskEdit(GetFromNam('SO_FERMEBEN', FF)).Plus   := ' AND G_LETTRABLE = "-" AND G_VENTILABLE = "-"';
  end;
  // FIN GCO

  // FQ 21634 : SBO 11/10/2007 Dans l'onglet Saisie, rendre non visible l'option Contrôle des quantités analytiques pour PCL
  if ExisteNam('SO_ZCTRLQTE', FF) then
  begin
    if (ctxPCL in V_PGI.PGIContexte) then
      SetInvi('SO_ZCTRLQTE', FF) ;
  end ;
  if ExisteNam('SO_CPANACTRLQTE', FF) then
  begin
    if (ctxPCL in V_PGI.PGIContexte) then
      SetInvi('SO_CPANACTRLQTE', FF) ;
  end ;
  // FQ 21634 Fin

  // FQ 22258 SBO 04/02/2008 : cacher paramsoc sur gestion doc dans la page saisie en mode pge
  if ExisteNam('SCO_CPMASQUESAISIE', FF) then
  begin
    if not (ctxPCL in V_PGI.PGIContexte) then
      begin
      SetInvi('SCO_CPMASQUESAISIE', FF) ;
      SetInvi('LSCO_CPMASQUESAISIE', FF) ;
      SetInvi('SO_CPVIEWERACTIF', FF) ;
      SetInvi('SO_CPVIEWERIMPORTGED', FF) ;
      SetInvi('SO_CPVIEWERPOSITION', FF) ;
      SetInvi('LSO_CPVIEWERPOSITION', FF) ;
      end ;
  end ;
  // FQ 22258

end;
{$ENDIF}

{$IFNDEF EAGLSERVER}
procedure MESEfficience( CC : TControl );
var
  FF : TForm;
  namChpExit : string;
  EffMin : double;
  EffMax : double;
begin
  FF := GetLaForm( CC );
  NamChpExit:=CC.Name;
  EffMin := Valeur(TEdit(GetFromNam('SO_MESEFFICMINI', FF)).Text);
  EffMax := Valeur(TEdit(GetFromNam('SO_MESEFFICMAXI', FF)).Text);
  if (NamChpExit = 'SO_MESEFFICMINI') and ( (EffMin<0) or (EffMin>100) ) then
    TEdit(GetFromNam('SO_MESEFFICMINI', FF)).Text :='0,00';
  if (NamChpExit = 'SO_MESEFFICMAXI') and  ( (EffMax<0) or (EffMax<EffMin) ) then
    TEdit(GetFromNam('SO_MESEFFICMAXI', FF)).Text :='0,00';
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure MESControle( CC : TControl );
var
  FF : TForm;
  namChpExit : string;
  NoChp   : integer;
  PresMaxi: double;
begin
  FF := GetLaForm( CC );
  NamChpExit:=CC.Name;
  NoChp   := ValeurI(TEdit(GetFromNam('SO_MESCHPRESPRES',FF)).Text);
  PresMaxi:= Valeur(TEdit(GetFromNam('SO_MESPRESMAXI',FF)).Text);
  if (NamChpExit = 'SO_MESCHPRESPRES') and  ( (NoChp<0) or (NoChp>3) ) then
    TEdit(GetFromNam('SO_MESCHPRESPRES', FF)).Text :='0';
  if (NamChpExit = 'SO_MESPRESMAXI') and (PresMaxi<0) then
    TEdit(GetFromNam('SO_MESPRESMAXI', FF)).Text :='0,00';
  if ( NamChpExit = 'SO_MESGRPATELIER' ) then
    SetEna('SO_MESGRPAUTONOME', FF, THCheckBox( GetFromNam( 'SO_MESGRPATELIER', FF )).Checked);
end;
{$ENDIF EAGLSERVER}
// GC_20071016_GM_GC15422_DEBUT
{$IFNDEF EAGLSERVER}
Function CtrlLongueur( CC : TControl ): boolean ;
var
  FF : TForm;
  namChpExit : string;
begin
  FF := GetLaForm( CC );
  result := true;
  if Not ExisteNam('SO_NATURECTRLCLOTURE',FF) then Exit ;
  NamChpExit:=CC.Name;
  if TEdit(GetFromNam('SO_NATURECTRLCLOTURE', FF)).Text = '' then
     TEdit(GetFromNam('SO_NATURECTRLCLOTURE', FF)).Text := '<<Aucun>>';
  if Length(TEdit(GetFromNam('SO_NATURECTRLCLOTURE', FF)).Text ) > 70 then
  begin
   ShowMessage('Vous ne pouvez sélectionner, au maximum, que 17 natures de pièces') ;
   if TEdit(GetFromNam('SO_NATURECTRLCLOTURE', FF)).CanFocus then TEdit(GetFromNam('SO_NATURECTRLCLOTURE', FF)).SetFocus ;
   result := false;
   Exit;
  end;
END ;
{$ENDIF EAGLSERVER}
// GC_20071016_GM_GC15422_FIN

{$IFNDEF EAGLSERVER}
Function VerifMES ( CC : TControl ) : boolean ;
Var FF : TForm ;
    EffMin  : double;
    EffMax  : double;
    NoChp   : integer;
    PresMaxi: double;
    Libelle : string;
    Q : TQuery;
  Function VerifAlea (FieldName : string ):boolean;
  begin
    Result :=True;
    if Not ExisteSQL('SELECT QOA_ALEA FROM QOALEA' ) then // Aléa non créé
      exit;
    Result :=False;
    if THValComboBox(GetFromNam(FieldName, FF)).Value='' then
      exit
    else
    begin
      Q:=OpenSQL('SELECT QOF_ALEADEDUCTIBLE FROM QOALEA '
                 + ' LEFT JOIN QOFAMILLEALEA ON QOF_FAMILLEALEA=QOA_FAMILLEALEA '
                 + ' WHERE QOA_ALEA="'+THValComboBox(GetFromNam(FieldName,FF)).Value+'"',True) ;
      try
        if (Not Q.EOF) and (Q.Fields[0].AsString<>'X') then
          exit;
      Finally
        Ferme(Q) ;
      end;
    end;
    Result :=True;
  end;

begin
  Result:=False;
  FF:=GetLaForm(CC) ;
  if ExisteNam('SO_MESTYPEOPER',FF) then
  begin
    if THMultiValComboBox(GetFromNam('SO_MESTYPEOPER', FF)).Text = '' then
    begin
      HShowMessage('30;MES;Vous devez renseigner les types d''opérations autorisés.;W;O;O;O;','','') ;
      exit;
    end;
  end;
  if ExisteNam('SO_MESSUIVIOP',FF) then
  begin
    if THMultiValComboBox(GetFromNam('SO_MESSUIVIOP', FF)).Text = '' then
    begin
      HShowMessage('30;MES;Vous devez renseigner les suivis opérations autorisés.;W;O;O;O;','','') ;
      exit;
    end;
  end;
  if ExisteNam('SO_MESEFFICMINI',FF) Then
  begin
    EffMin  := Valeur(TEdit(GetFromNam('SO_MESEFFICMINI', FF)).Text);
    EffMax  := Valeur(TEdit(GetFromNam('SO_MESEFFICMAXI', FF)).Text);
    if (EffMin<0) or (EffMin>100) then
    begin
      HShowMessage('30;MES;L''efficience mainimale doit être entre 0 % et 100%.;W;O;O;O;','','') ;
      exit;
    end;
    if (EffMax<0) or (EffMax<EffMin) then
    begin
      HShowMessage('30;MES;L''efficience maximale doit être supérieure à l''efficience minimale.;W;O;O;O;','','') ;
      exit;
    end;
  end;
  if ExisteNam('SO_MESCHPRESPRES',FF) Then
  begin
    NoChp   := ValeurI(TEdit(GetFromNam('SO_MESCHPRESPRES',FF)).Text);
    if (NoChp<0) or (NoChp>3) then
    begin
      HShowMessage('30;MES;Le champ libre ressource doit être compris entre 0 et 3.;W;O;O;O;','','') ;
      exit;
    end;
    if (NoChp>=1) and (NoChp<=3) then
    begin
      Libelle := RechDomZoneLibre ('RM'+IntToStr(NoChp), False);
      if libelle ='.-' then libelle :='';
      if libelle='' then
      begin
        HShowMessage('30;MES;Le libellé du champ libre pour le contrôle des présences maxi n''est pas renseigné sur les ressources.;W;O;O;O;','','') ;
        exit;
      end;
    end;
  end;
  if ExisteNam( 'SO_MESGRPAUTONOME', FF ) and (THCheckBox( GetFromNam( 'SO_MESGRPAUTONOME', FF ) ).Checked = True )
  and (THCheckBox( GetFromNam( 'SO_MESGRPATELIER', FF ) ).Checked <> True ) then
  begin
    HShowMessage( '30;MES;Le gestion des groupes autonomes impose la gestion des groupes sur atelier.;W;O;O;O;', '', '' );
    exit;
  end;
  if ExisteNam('SO_MESPRESMAXI',FF) Then
  begin
    PresMaxi:= Valeur(TEdit(GetFromNam('SO_MESPRESMAXI',FF)).Text);
    if PresMaxi<=0 then
    begin
      HShowMessage('30;MES;Le temps de présence maximum journalier n''est pas renseigné.;W;O;O;O;','','') ;
      exit;
    end;
  end;
  if ExisteNam('SO_MESALEATPSPASSE',FF) then
  begin
    if not VerifAlea ('SO_MESALEATPSPASSE') then
    begin
      HShowMessage('30;MES;L''aléa associé aux opérations ''Temps passé'' n''est pas correct.;W;O;O;O;','','') ;
      exit;
    end;
  end;
  if ExisteNam('SO_MESALEAPREPA',FF) then
  begin
    if not VerifAlea ('SO_MESALEAPREPA') then
    begin
      HShowMessage('30;MES;L''aléa associé aux opérations ''Préparation'' n''est pas correct.;W;O;O;O;','','') ;
      exit;
    end;
  end;
  if ExisteNam('SO_MESALEATECHNO',FF) then
  begin
    if not VerifAlea ('SO_MESALEATECHNO') then
    begin
      HShowMessage('30;MES;L''aléa associé aux opérations ''Technologique'' n''est pas correct.;W;O;O;O;','','') ;
      exit;
    end;
  end;
  Result:=True;
  if ExisteNam( 'SO_MESALIMBAC', FF ) then
  begin
    THValComboBox(GetFromNam('SO_MESALIMOPEBAC', FF)).Text := '';
  end;

end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{b FP 21/02/2006}
procedure MAJGestionCompensation(CC : TControl);
var
  FF:        TForm;
  CbGestion: TCheckBox;
BEGIN
  FF        := TForm(CC.Owner) ;
  if Not ExisteNam('SO_CPGESTCOMPENSATION',FF) then Exit ;
  CbGestion := TCheckBox(GetFromNam('SO_CPGESTCOMPENSATION',FF));
  if CbGestion = nil then
    Exit;

  SetEna('SO_CPJALCOMPENSATION',  FF, CbGestion.Checked);
  SetEna('SO_CPPLANCOMPENSATION', FF, CbGestion.Checked);
  if CbGestion.Checked then
    begin
    TComboBox(GetFromNam('SO_CPJALCOMPENSATION', FF)).Style := csDropDownList;
    TComboBox(GetFromNam('SO_CPPLANCOMPENSATION',FF)).Style := csDropDownList;
    end
  else
    begin
    {Permet d'avoir une valeur vide}
    TComboBox(GetFromNam('SO_CPJALCOMPENSATION', FF)).Style := csDropDown;
    TComboBox(GetFromNam('SO_CPPLANCOMPENSATION',FF)).Style := csDropDown;
    TComboBox(GetFromNam('SO_CPJALCOMPENSATION', FF)).Text := '';
    TComboBox(GetFromNam('SO_CPPLANCOMPENSATION',FF)).Text := '';
    end;
end;
{$ENDIF EAGLSERVER}

{$IFDEF COMPTA} //YMO Norme NF 203
{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 07/12/2006
Modifié le ... :   /  /    
Description .. : Conformité avec la norme NF 203 et BOI su 24/01/2006
Mots clefs ... : NORME NF 203 BOI
*****************************************************************}
function CActivConformiteNF203 ( CC : TControl ) : boolean ;
var FF : TForm;
    vStConformeMsg : string;
    vstActivConform : Boolean;

begin
  Result:=True ;
  FF := GetLaForm(CC) ;
  if ExisteNam('SO_CPCONFORMEBOI', FF) then
  begin
    vStActivConform := TCheckBox(GetFromNam('SO_CPCONFORMEBOI', FF)).Checked;

    if vStActivConform<>GetParamSocSecur('SO_CPCONFORMEBOI',false) then
    begin
      if vStActivConform then
        VStConformeMsg:='Activation Conformité NF 203'
      else
        VStConformeMsg:='Désactivation Conformité NF 203';

      { BVE 29.08.07 : Mise en place d'un nouveau tracage }
{$IFNDEF CERTIFNF}
      CPEnregistreLog(VStConformeMsg);
{$ELSE}
      CPEnregistreJalEvent('CBO','Conformité BOI',VStConformeMsg);
{$ENDIF}       
    end;

  end;

end;
{$ENDIF}

{$IFNDEF EAGLSERVER}
Function VerifCompensation(CC : TControl): boolean;
var
  FF:        TForm;
  CbGestion: TCheckBox;
  NumPlan:   Integer;
  QuelPlan:  string;  //Thl
begin
  Result    := True;
  FF        := TForm(CC.Owner) ;
  if Not ExisteNam('SO_CPGESTCOMPENSATION',FF) then Exit ;
  CbGestion := TCheckBox(GetFromNam('SO_CPGESTCOMPENSATION',FF));
  //NumPlan   := 0;
  if CbGestion = nil then
    Exit;

  MAJGestionCompensation(CC);
  if CbGestion.Checked then
    begin
    if Result and (TComboBox(GetFromNam('SO_CPPLANCOMPENSATION',FF)).ItemIndex = -1) then
      begin
      HShowMessage('0;Paramètres société;Le plan de compensation n''est pas renseigné.;W;O;O;O;','','') ;
      Result := False;
      end;
    {b Thl FQ 18643}
    NumPlan := TComboBox(GetFromNam('SO_CPPLANCOMPENSATION',FF)).ItemIndex+1;
    if not GetParamSocSecur('SO_CORSAU' + IntToStr(NumPlan), True, True) then
      begin
      QuelPlan := TComboBox(GetFromNam('SO_CPPLANCOMPENSATION',FF)).Items[NumPlan-1];
      HShowMessage('0;Paramètres société;Le ' + QuelPlan + ' n''est pas renseigné dans paramètres société-comptable.;W;O;O;O;','','') ;
      Result := False;
      end;
    {e Thl}
    if Result and (TComboBox(GetFromNam('SO_CPJALCOMPENSATION',FF)).ItemIndex = -1) then
      begin
      HShowMessage('0;Paramètres société;Le journal de compensation n''est pas renseigné.;W;O;O;O;','','') ;
      Result := False;
      end;

    if Result then
      NumPlan := TComboBox(GetFromNam('SO_CPPLANCOMPENSATION',FF)).ItemIndex+1;
    if Result and
       (IntToStr(NumPlan) <> GetParamSocSecur('SO_CPPLANCOMPENSATION', '')) and
       ExisteSQL('select CR_CORRESP from corresp where CR_TYPE="AU'+IntToStr(NumPlan)+'"') then
      begin
      Result := False;
      if (HShowMessage('0;Paramètres société;Le plan '+IntToStr(NumPlan)+' est déjà utilisé, voulez-vous le supprimer?;Q;YNC;N;C;', '', '') = mrYes) and
         (HShowMessage('0;Paramètres société;La suppression des données du plan sera définitive, voulez-vous continuer?;Q;YNC;N;C;', '', '') = mrYes) then
        begin
        SetParamSoc('SO_CORSAU'+IntToStr(NumPlan), True);
        ExecuteSQL('delete from corresp where CR_TYPE="AU'+IntToStr(NumPlan)+'"');
        ExecuteSQL('update tiers set T_CORRESP'+IntToStr(NumPlan)+'=""');
        Result := True;
        end
      end;
    end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF GCGC}
function ControleStock (FF : TForm; bMessage : boolean = true) : boolean;
{$IFNDEF GPAO}
  Var
    CC : THCritMaskEdit ;
    TCB: TCheckBox;
{$ENDIF !GPAO}
begin
  Result := true;
  {$IFNDEF GPAO}
    TCB := TCheckBox (GetFromNam ('SO_MODEMVTEX',FF)) ;

    SetEna ('SO_GCTIERSMVSTK', FF, TCB.Checked);
    SetEna ('SO_EDITRAPMVTEX', FF, not (TCB.Checked));
    TCheckBox(GetFromNam ('SO_EDITRAPMVTEX', FF)).Checked := not TCB.Checked;

    if TCB.Checked then
    begin
      CC := THCritMaskEdit (GetFromNam ('SO_GCTIERSMVSTK', FF));
      if (CC.Text = '') and bMessage then
      begin
        HShowMessage('0;Stock;Le tiers des mouvements exceptionnels est obligatoire.;E;O;O;O;','','') ;
//        SetEna ('SO_GCTIERSMVSTK', FF, true);
        Result := false;
      end;
    end;
  {$ENDIF !GPAO}
end;
{$ENDIF GCGC}
{$ENDIF EAGLSERVER}

{e FP 21/02/2006}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/03/2002
Modifié le ... :   /  /
Description .. : LG* ajout de la gestion des champs comptes de l'onglet
Suite ........ : lettrage\gestion
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
Function InterfaceSoc ( CC : TControl ) : boolean ;
Var
  NamChpExit : String ;   // Champ que l'on vient de quitter
  NamChpActif: String ;   // Champ sur lequel est positionné le focus
  ChampActif : TWinControl;
  ena : boolean;
  CodeAxe : string;
  {$IFDEF AFFAIRE}  // fct spécif GI/GA
    FF:TForm;
  {$ENDIF AFFAIRE}
  TheValeur : string;
BEGIN
  Result:=True ;

  { Champ exit }
  NamChpExit:=CC.Name;
  
  if NamChpExit = 'SO_BTTYPELAN' then
  begin
    FF:=TForm(CC.Owner) ;
    TheValeur := THValComboBox(GetFromNam('SO_BTTYPELAN', FF)).Value;
    THValComboBox(GetFromNam('SO_BTPLANNINGDEF', FF)).Plus := 'AND HPP_MODEPLANNING ="'+TheValeur+'"';
  end else if NamChpExit='SO_DATEDEBUTEURO'                then DateDebutEuroX(CC)
  else
  {$IFDEF AFFAIRE}  // fct spécif GI/GA
  if NamChpExit='SO_AFGESTIONAPPRECIATION'        then AppreciationX(CC)
  else if NamChpExit='SO_AFPLANDECHARGE'          then AfPlanDeChargeX(CC)
  else if NamChpExit='SO_AFGESTIONRAF'            then AfGestionRAFX(CC)
  else if NamChpExit='SO_AFAPPCUTOFF'             then CutOffX(CC)
  else if NamChpExit='SO_AFALIMCUTOFF'            then CutOffZeroX(CC)
  else if NamChpExit='SO_AFALIMCUTOFFACH'         then CutOffZeroAchX(CC)
  else if NamChpExit='SO_AFFRAISCOMPTA'           then FraisX(CC)
  else if NamChpExit='SO_AFGESTFRAISKM'           then ActiviteFraisKMX(CC) // PL Le 05/07/05 : gestion frais km activité
  else if NamChpExit='SO_TYPEDATEFINACTIVITE'     then
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
  else if NamChpExit='SO_AFLIENPAIEDEC'     then
  begin    //mcd 30/11/2006
    FF:=TForm(CC.Owner) ;
    if THCheckBox(GetFromNam ('SO_AFLIENPAIEDEC', FF)).checked then
      SetEna('SO_AFRESSCOMMUNE', FF, True)
    else
    begin
      SetParamSoc('SO_AFRESSCOMMUNE', False);
      SetEna('SO_AFRESSCOMMUNE', FF, False)
    end;
  end
  else if NamChpExit='SO_BTAXEANALSTOCK' then
  begin
    FF:=TForm(CC.Owner) ;
    CodeAxe := thvalcombobox(cc).value;
    if VH^.LiaisonY2ViaShare then
    begin
      THEdit(GetFromNam('SO_BTCPTANALSTOCK',FF)).DataType := 'TZCSECTION'+Copy(CodeAxe,2,1);;
    end else
    begin
      THEdit(GetFromNam('SO_BTCPTANALSTOCK',FF)).DataType := 'TZSECTION'+Copy(CodeAxe,2,1);;
    end;
  end else if NamChpExit='SO_OPTANALSTOCK' then
  begin
    FF:=TForm(CC.Owner) ;
    Ena := THCheckBox(GetFromNam ('SO_OPTANALSTOCK', FF)).checked;
    if not ena then
    begin
      SetParamSoc('SO_BTAXEANALSTOCK', '');
      SetParamSoc('SO_BTCPTANALSTOCK', '');
      SetParamSoc('SO_BTJNLANALSTOCK', '');
    end;
    SetEna ('SO_BTAXEANALSTOCK',FF,ena) ;
    SetEna ('SO_BTCPTANALSTOCK',FF,ena) ;
    SetEna ('SO_BTJNLANALSTOCK',FF,Ena) ;
  end
  else if NamChpExit='SO_BTMETREDOC' then
  begin
    FF:=TForm(CC.Owner) ;
    if THCheckBox(GetFromNam ('SO_BTMETREDOC', FF)).checked then
    begin
      SetEna('SO_METRESEXCEL', FF, true);
      SetEna('SO_BTMETREBIB', FF, True);
      SetEna('SO_BTREPMETR', FF, True);
    end else
    begin
      SetEna('SO_METRESEXCEL', FF, false);
      SetEna('SO_BTMETREBIB', FF, false);
      SetParamSoc('SO_METRESEXCEL', '-');
      SetParamSoc('SO_BTMETREBIB', '-');
      SetEna('SO_BTREPMETR', FF, false);
    end;
  end
  else if NamChpExit='SO_METRESEXCEL' then
  begin
    FF:=TForm(CC.Owner) ;
    if THCheckBox(GetFromNam ('SO_METRESEXCEL', FF)).checked then
    begin
      SetEna('SO_BTMETREBIB', FF, True);
      SetEna('SO_BTREPMETR', FF, True);
    end else
    begin
      SetEna('SO_BTMETREBIB', FF, false);
      SetParamSoc('SO_BTMETREBIB', '-');
      SetEna('SO_BTREPMETR', FF, false);
    end;
  end
  {$IFDEF GIGI}
  // PCH 11-2005 Gestion personnalisation paramsoc
  else if NamChpExit='SO_AFPERSOCHOIX'  then CtlPersoParamsocChoix(CC)
  else if NamChpExit='SO_AFPERSOPARAM1' then CtlPersoParamsocPar1(CC)
  else if NamChpExit='SO_AFPERSOPARAM2' then CtlPersoParamsocPar2(CC)
  // PCH 12/12/2006 Gestion personnalisation FACTURE LIQUIDATIVE
  else if NamChpExit='SO_AFGERELIQUIDE' then CtlFactureLiquidative(CC)
  {$ENDIF GIGI}

{$IFNDEF EAGLSERVER}
  // BDU - 12/12/06
  else if (NamChpExit = 'SO_AFDEMIJOURNEE') or (NamChpExit = 'SO_AFPLANNINGHEURE') or
    (NamChpExit = 'SO_AFPLANNINGJOUR') or (NamChpExit = 'SO_AFPLANNINGMOIS') or
    (NamChpExit = 'SO_AFPLANNINGDEMIHEURE') then
    GestionAffichePlanning(TForm(CC.Owner))

  else if NamChpExit = 'SO_AFQUELPLANNING' then
    AfPlanDeChargeX(CC)
  // BDU - 30/01/07 - Contrôle des heures d'affichage du planning

  else if (NamChpExit = 'SO_AFHEUREDEBUTJOUR') or (NamChpExit = 'SO_AFHEUREFINJOUR') then
    ControleHeure(NamChpExit, TForm(CC.Owner))

  // BDU - 24/11/06 - 12887, Visa automatique
  else if (NamChpExit = 'SO_AFVISAACTIVITE') or (NamChpExit = 'SO_AFVISAAUTO') then
    GestionVisaAutomatique(TForm(CC.Owner))
  // BDU - 29/11/07. Nouvelle exportation SDA
  else if (NamChpExit = 'SO_AFEXPORTSDAAFFAIRE') or (NamChpExit = 'SO_AFEACTUNPOURTOUS') then
    GestionExportSDA(TForm(CC.OWner))
{$ENDIF EAGLSERVER}

  else
  {$ENDIF AFFAIRE}
  if isFourcCpte(NamChpExit)                           then CptBourreX(CC)
  else if Copy(NamChpExit,1,9)='SO_DEFCOL'             then CptBourreX(CC)
  else if NamChpExit='SO_DEVISEPRINC'                  then DevPrincX(CC)
  else if NamChpExit='SO_CONTROLEBUD'                  then ControleBudX(CC)
  else if NamChpExit='SO_PAYS'                         then PaysX(CC)
  else if (NamChpExit='SO_GCNUMTIERSAUTO') or
          (NamChpExit='SO_GCFORMATCODETIERS')
          {$IFNDEF PGIMAJVER} or (NamChpExit='SO_BOURREAUX') {$ENDIF !PGIMAJVER} then NumTiersAutoX(CC, False)
  else if NamChpExit='SO_GCPREFIXETIERS'               then PrefixeTiersOk(CC)
  else if NamChpExit='SO_GCCOMPTEURTIERS'              then LgNumTiersOk(CC)
  else if NamChpExit='SO_GCNUMARTAUTO'                 then CodeArtAutoX(CC, False)
  else if NamChpExit='SO_GCPREFIXEART'                 then PrefixeArtOk(CC)
  else if NamChpExit='SO_GCCOMPTEURART'                then LgCodeArtOk(CC)
  else if NamChpExit='SO_GCNUMCABAUTO'                 then NumCABAutoX(CC)
  else if NamChpExit='SO_GCPRCABAUTO'                  then PrCABAutoX(CC)   // CCMX-CEGID FQ N° GC14064
  else if NamChpExit='SO_ARTLOOKORLI'                  then FamillesArticle4a8(CC)
  else if NamChpExit='SO_RTPROJGESTION'                then AlimZonesProjet(CC)
  else if (NamChpExit='SO_GESTSTATUTUNIQUE1')
       or (NamChpExit='SO_GESTSTATUTUNIQUE2')          then GereZonesVersions(CC)
  else if NamChpExit='SO_GCFORMATEZONECLI'             then FormatZoneCli(CC)
  else if (NamChpExit='SO_RTPROPTABHIE')
       or (NamChpExit='SO_RTACTTABHIE')                then TraiteTabHie(CC,NamChpExit)
  else if NamChpExit='SO_GCTOXCONFIRM'                 then SocTox(CC)
  else if Copy(NamChpExit,1,12)='SO_LETCHOIXG'         then CptBourreX(CC)
  else if Copy(NamChpExit,1,11)='SO_LETECART'          then ControleLgNum(CC)
  else if NamChpExit='SO_NBJECRAVANT'                  then ControleLgNum(CC)
  else if NamChpExit='SO_NBJECRAPRES'                  then ControleLgNum(CC)
  else if NamChpExit='SO_NBJECHAVANT'                  then ControleLgNum(CC)
  else if NamChpExit='SO_NBJECHAPRES'                  then ControleLgNum(CC)
  else if NamChpExit='SO_LETTOLERC'                    then ControleLgNum(CC)
  else if NamChpExit='SO_LETTOLERF'                    then ControleLgNum(CC)
  else if NamChpExit='SO_I_LOCALISATION'               then VerifCBLocal(CC)
  else if NamChpExit='SO_I_AUTREINFO'                  then VerifCBAutreInfo(CC)
  else if NamChpExit='SO_I_FISUIVILIBRE'               then VerifFinancementInfoLibre(CC)
  else if NamChpExit='SO_ICCCOMPTECAPITAL'             then CptBourreX(CC)
  else if NamChpExit='SO_CPCHARGAVAN'                  then CptBourreX(CC)
  else if NamChpExit='SO_CPINTERCOUR'                  then CptBourreX(CC)
  else if NamChpExit='SO_CPFRAISFINA'                  then CptBourreX(CC)
  else if NamChpExit='SO_CPASSURANCE'                  then CptBourreX(CC)
  else if NamChpExit='SO_GCMULTIDEPOTS'                then GestionMonoDepot(CC)
  else if NamChpExit='SO_GCECHEANCES'                  then MajGPPEcheances(CC)
  else if (NamChpExit='SO_DECQTE_')
       or (NamChpExit='SO_DECPRIX_')                   then CtrlNbDecimales(CC, NamChpExit)
  else if NamChpExit='SO_I_CALCULAMORTEXESORTIE'       then VerifAmorJourSortie(CC)
  else if NamChpExit='SO_LETMODE' then
  begin // FQ 11765
    if (THValComboBox(CC).Value='PME') then
      SetEna('SO_LETCHOIXDEFVALID', TForm(CC.Owner), False)
    else
      SetEna('SO_LETCHOIXDEFVALID', TForm(CC.Owner), True);
  end
  else if NamChpExit = 'SO_CBN_QTE_JOUR'          then BornesVisibles(CC)
  else if NamChpExit = 'SO_QBORNEMINSTOCKCALCUL_' then VerifBorne(CC, NamChpExit, '', 'MIN')
  else if NamChpExit = 'SO_QBORNEMAXSTOCKCALCUL_' then VerifBorne(CC, 'SO_QBORNEMINSTOCKCALCUL_', NamChpExit, 'MAX')
  else if NamChpExit = 'SO_CDEMARCHE'             then GereZonesMarche(CC, 'G') //Global
  else if NamChpExit = 'SO_CDEMARCHEA'            then GereZonesMarche(CC, 'A') //Marché achat
  else if NamChpExit = 'SO_CDEMARCHEV'            then GereZonesMarche(CC, 'V') //Marché vente
  {$IFDEF GESCOM}
    else if NamChpExit = SOCdeOuv then CdeOuvCacheFonctions(GetLaForm(CC))
    else if NamChpExit = SOCdeOuvAutoTransfo then CdeOuvTraiteAutoTransfo(CC)
  {$ENDIF GESCOM}
  {$IFDEF GCGC}
  else if NamChpExit = 'SO_GCDEFQUALIFUNITEGA' then
    SetPlusUniteParDefaut(CC)
  else if NamChpExit = 'SO_GCGESTUNITEMODE' then
    SetGCGestUniteModeState(CC)
  else if NamChpExit = 'SO_GCENSEIGNETAB' then
    SetGereEnseigne(CC)
  else if NamChpExit = 'SO_NUMCIBLAGEAUTO' then
    SetGereCiblage(CC)
  else if NamChpExit = 'SO_RTNUMOPERAUTO' then
    SetGereOperation(CC)
  {$ENDIF GCGC}
  else if (NamChpExit = 'SO_MESEFFICMINI') or (NamChpExit = 'SO_MESEFFICMAXI') then
    MESEfficience( CC )
  else if (NamChpExit = 'SO_MESCHPRESPRES') or (NamChpExit = 'SO_MESPRESMAXI') then
    MESControle( CC )
  else if ( NamChpExit = 'SO_MESGRPATELIER' )then
    MESControle( CC )
  else if NamChpExit = 'SO_GCAXEANALYTIQUE' then
    GereAnaParamAff(CC)
  else if NamChpExit = 'SO_GERETVAINTRACOMM' then
    GereTypeTVAIntraComm(CC)
  else if (NamChpExit = 'SO_GERECONTREMARQUE') or (NamChpExit = 'SO_GCPCEAUTOCONTREM') or (NamChpExit = 'SO_GCFORCEAUTOCONTREM') or
          (NamChpExit = 'SO_GCAUTOCONTREM')then
    GereContremarque(CC)
  else if NamChpExit = 'SO_GEREFARFAE' then
    GereFarFae(CC)
{$IFNDEF EAGLSERVER}
  else if (NamChpExit = 'SO_GCELIMINEVENLIGNE0') or (NamChpExit = 'SO_GCELIMINEACHLIGNE0') then
    GereElimineLigne0(CC)
{$ENDIF EAGLSERVER}
  else if (NamChpExit = 'SO_GCTAUXNEGOCIE') or (NamChpExit = 'SO_TYPETAUXDEVISE') then
    GereDevisePiece(CC)
  // Gestion des champs de paramétrage de la saisie
  else if NamChpExit = 'SO_CPMASQUECRIT1' then
    GereMasqueCrit( GetLaForm(CC) )
  else if NamChpExit = 'SO_CPMASQUECRIT2' then
    GereMasqueCrit( GetLaForm(CC) )
  else if NamChpExit = 'SO_CPMASQUECRIT3' then
    GereMasqueCrit(  GetLaForm(CC) )
{ GC_JTR_GC15601_Début }
  else if (NamChpExit = 'SO_GCBLOQUEMARGE') or (NamChpExit = 'SO_FORCEMARGEMINI') then
    GereMdpMargeMini(CC)
{ GC_JTR_GC15601_Fin }
  else if NamChpExit = 'SO_PDRNATUREPDRGA' then
  begin
    THValComboBox(GetFromNam('SO_PDRTYPEPDRGA', TForm(CC.OWner))).Plus := 'WRT_NATUREPDR="'+THValComboBox(GetFromNam('SO_PDRNATUREPDRGA', TForm(CC.OWner))).Value+'"';
    if (THValComboBox(GetFromNam('SO_PDRNATUREPDRGA', TForm(CC.OWner))).Value='BTH') then
      THValComboBox(GetFromNam('SO_PDRTYPEPDRGA', TForm(CC.OWner))).Value:='ACT'
    else
      THValComboBox(GetFromNam('SO_PDRTYPEPDRGA', TForm(CC.OWner))).Value:='ORT';
  end
  else if NamChpExit='SO_MAJAFFECTATION' then
  begin
    if THCheckBox(GetFromNam ('SO_MAJAFFECTATION', TForm(CC.Owner))).checked then
    begin
      SetVisi('LSO_DUREEMINI', TForm(CC.Owner));
      SetVisi('SO_DUREEMINI', TForm(CC.Owner));
    end else
    begin
      SetInvi('LSO_DUREEMINI', TForm(CC.Owner));
      SetInvi('SO_DUREEMINI', TForm(CC.Owner));
    end;
  end
{$IFDEF GCGC}
  else if NamChpExit = 'SO_MODEMVTEX' then
  begin
    ControleStock (GetLaForm(CC), false);
  end
{$ENDIF GCGC}
  ;


  { Champ actif }
  ChampActif := screen.ActiveControl;
  NamChpActif:=ChampActif.Name;
  if NamChpActif='SO_GCNUMTIERSAUTO'    then NumTiersAutoX(CC, True)
  else if NamChpActif='SO_GCNUMARTAUTO' then CodeArtAutoX(CC, True)
  else if NamChpExit = 'SO_SOLDEPIECEPREC' then
    GereSoldePiecePrec(CC);
	{$IFDEF GPAOLIGHT}
  {$IFNDEF EAGLSERVER}
  	InitParamSocGPAO(GetLaForm(CC));
  {$ENDIF EAGLSERVER}
	{$ENDIF GPAOLIGHT}
END ;
{$ENDIF EAGLSERVER}

{============================= Sur Chargement =========================}
Function ExisteMvtsPayeurs ( Jal : String ) : boolean ;
BEGIN
Result:=False ; if Jal='' then Exit ;
Result := ExisteSQL ('SELECT E_JOURNAL FROM ECRITURE WHERE E_JOURNAL="'+Jal+'"');
END ;

Function ExisteMvtsCpte ( Cpte : String ) : boolean ;
BEGIN
Result:=False ; if Cpte='' then Exit ;
Result := ExisteSQL('SELECT E_GENERAL FROM ECRITURE WHERE E_GENERAL="'+Cpte+'"') ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : JCF
Créé le ...... : 02/08/2001
Modifié le ... : 02/08/2001
Description .. : Modification des règles d'existence des mouvements pour la
Suite ........ : mode.
Suite ........ : Les blocages comptables sont différents des blocages de la
Suite ........ : gescom
Mots clefs ... : PARAMSOC;MODE;AREVOIR
*****************************************************************}
Function ExisteMvts : boolean ;
var
  OkMvt : boolean ;
BEGIN
  OkMvt:=False ;
  if Not OkMvt then OkMvt := ExisteSQL('SELECT 1 FROM ECRITURE');
  if Not OkMvt then OkMvt := ExisteSQL('SELECT 1 FROM ANALYTIQ');
  if Not OkMvt then OkMvt := ExisteSQL('SELECT 1 FROM BUDECR');
  if Not OkMvt then OkMvt := ExisteSQL('SELECT 1 FROM ECRGUI');
  Result:=OkMvt ;
END ;

{$IFNDEF EAGLSERVER}
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
   TOBSoc.SetBoolean('OKGEN', DesEnregs('GENERAUX')) ;
   TOBSoc.SetBoolean('OKAUX', DesEnregs('TIERS')) ;
   TOBCpt:=TOB.Create('CPTS',LaTOB,-1) ;
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
{$ENDIF EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : JCF
Créé le ...... : 02/08/2001
Modifié le ... :   /  /
Description .. : Modification du blocage de la modification de la devise
Suite ........ : dans les paramètres société.
Suite ........ : Réalisé en spécifique pour la MODE
Mots clefs ... : PARAMSOC;EURO;MODE
*****************************************************************}
{$IFNDEF EAGLSERVER}
Procedure BlocageDevise ( CC : TControl ) ;
Var
  FF : TForm ;
  Ena : boolean ;
BEGIN
  FF:=GetLaForm(CC) ;
  if Not ExisteNam('SO_DEVISEPRINC',FF) then Exit ;
  Ena:=(GetTOBSoc(CC).GetValue('OKECR')='-') ;
  SetEna('SO_DEVISEPRINC',FF,Ena) ;
  SetEna('SO_DECVALEUR',FF,False) ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure BlocageECC ( CC : TControl ) ;
Var FF : TForm ;
    Ena : boolean ;
BEGIN
FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_ECCEURODEBIT',FF) then Exit ;
Ena:=(GetTOBSoc(CC).GetValue('OKECCDEBIT')='-')  ; SetEna('SO_ECCEURODEBIT',FF,Ena) ;
Ena:=(GetTOBSoc(CC).GetValue('OKECCCREDIT')='-') ; SetEna('SO_ECCEUROCREDIT',FF,Ena) ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF NETEXPERT}
Procedure InitComptaRevision ( CC : TControl ) ;
Var
  FF : TForm ;
BEGIN
  FF:=GetLaForm(CC) ;
  if ExisteNam('SO_CPLIENGAMME',FF) then
  begin
       if GetParamSocSecur ('SO_CPLIENGAMME', '')= 'S1' then
       begin
            if V_PGI.PassWord=CryptageSt(DayPass(Date)) then
            begin
                 SetEna('SO_CPLIENGAMME', FF, TRUE);
                 SetEna('SO_CPMODESYNCHRO', FF, TRUE);
            end
            else if _IsDossierNetExpert then
            begin
                 SetEna('SO_CPLIENGAMME', FF, false);
                 SetEna('SO_CPMODESYNCHRO', FF, FALSE);
                 SetParamsoc ('SO_FLAGSYNCHRO', 'SYN');
            end;
            if (not V_PGI.SAV) then
            begin
                  if ((GetParamSocSecur ('SO_FLAGSYNCHRO', '')= 'SYN') and (GetparamsocSecur ('SO_CPMODESYNCHRO', '') = TRUE))  or
                  (((GetParamSocSecur ('SO_FLAGSYNCHRO', '')= 'JRL') or (GetParamSocSecur ('SO_FLAGSYNCHRO', 'JRL')= 'BAL')) and  (GetparamsocSecur ('SO_CPMODESYNCHRO', '') = FALSE)) then
                  begin
                    SetEna('SO_CPMODESYNCHRO', FF, FALSE);
                    SetEna('SO_CPLIENGAMME', FF, FALSE);
                  end;
            end;
       end;
       SetInvi ('SO_CPFREQUENCESX', FF);  SetInvi ('LSO_CPFREQUENCESX', FF);
       if ExisteNam('SO_CPJOURNALEXPERT',FF) and (GetParamSocSecur ('SO_CPJOURNALEXPERT', '')= '') then
            THCritMaskEdit(GetFromNam('SO_CPJOURNALEXPERT',FF)).Text := '' ;
  end;
END;
{$ENDIF}
{$ENDIF EAGLSERVER}


{***********A.G.L.***********************************************
Auteur  ...... : JCF
Créé le ...... : 02/08/2001
Modifié le ... :   /  /
Description .. : Modification pour la Mode des règles de modification des
Suite ........ : informations euro.
Suite ........ : Pour règler les dossiers en multi-devises
Mots clefs ... : PARAMSOC;MODE;EURO;MULTIDEVISE
*****************************************************************}
{$IFNDEF EAGLSERVER}
Procedure InitDivers ( CC : TControl ) ;
const
  ORIGINERELEVE   = 'INT';
  CODENEWPOINTAGE = 'RAP';
Var
  FF : TForm ;
  JalBud,RefL,RefP : THValComboBox ;
  OldV : String ;
  HH   : THLabel ;
  CB1,CB2 : tcheckBox ;
BEGIN
  FF:=GetLaForm(CC) ;
  {$IFDEF GIGI}
  {mng 06/03/07 plus de paramsoc alertes
  if ExisteNam('SO_GEREALERTEAFF',FF) then traduitForm(FF) ;}
  {$ENDIF GIGI}
  if ExisteNam('SO_TAUXEURO',FF) then
  BEGIN
    DevPrincX(CC) ;
    SetEna('SO_TAUXEURO',FF,(V_PGI.DateEntree<V_PGI.DateDebutEuro)) ;
    SetEna('LSO_TAUXEURO',FF,(V_PGI.DateEntree<V_PGI.DateDebutEuro)) ;
    SetEna('SO_TENUEEURO',FF,False) ;
    if TCheckBox(GetFromNam('SO_TENUEEURO',FF)).Checked then
      SetEna('SO_TAUXEURO',FF,False) ;
    if VH^.TenueEuro then
      SetEna('SO_DATEDEBUTEURO',FF,False) ;
  END;

  if ExisteNam('SO_JALCTRLBUD',FF) then
  BEGIN
    JalBud:=THValComboBox(GetFromNam('SO_JALCTRLBUD',FF)) ;
    if JalBud.Vide then
      JalBud.Items[0]:='' ;
    TCheckBox(GetFromNam('SO_CONTROLEBUD',FF)).Checked:=(JalBud.Value<>'') ;
    if JalBud.Vide then
      JalBud.Enabled:=False ; // FQ 13130 Vl 11122003
  END;

  if ExisteNam('SO_CPEXOREF',FF) then if Not (ctxPCL in V_PGI.PGIContexte) then
  BEGIN
    SetInvi('SO_CPEXOREF',FF) ; SetInvi('LSO_CPEXOREF',FF) ;
  END;

  if ExisteNam('SO_LGCPTEGEN',FF) then
  BEGIN
    if GetTOBSoc(CC).GetValue('OKGEN')='X' then
      SetEna('SO_LGCPTEGEN',FF,False) ;
    if GetTOBSoc(CC).GetValue('OKAUX')='X' then
      SetEna('SO_LGCPTEAUX',FF,False) ;
  END;

  if ExisteNam('SO_JALATP',FF) then
  BEGIN
    if GetTOBSoc(CC).GetValue('OKJALATP')='X' then
      SetEna('SO_JALATP',FF,False) ;
    if GetTOBSoc(CC).GetValue('OKJALVTP')='X' then
      SetEna('SO_JALVTP',FF,False) ;
  END;

  if EstSerie(S5) then
  BEGIN
    if ExisteNam('SO_CPREFLETTRAGE',FF) then
    BEGIN
      RefL:=THValComboBox(GetFromNam('SO_CPREFLETTRAGE',FF)) ;
      OldV:=RefL.Value ;
      RefL.DataType:='CPPARAMREFERENCESS5' ; RefL.Reload ;
      RefL.Value:=OldV ;
    END;
    if ExisteNam('SO_CPREFPOINTAGE',FF) then
    BEGIN
      RefP:=THValComboBox(GetFromNam('SO_CPREFPOINTAGE',FF)) ;
      OldV:=RefP.Value ;
      RefP.DataType:='CPPARAMREFERENCESS5' ; RefP.Reload ;
      RefP.Value:=OldV ;
    END ;
    if ExisteNam('SO_CPLIBREANAOBLI',FF) then
      SetInvi('SO_CPLIBREANAOBLI',FF) ;
  END;

  if ExisteNam('LSO_CPRDNOMCLIENT',FF) and (Not V_PGI.Controleur) then
  BEGIN
    HH:=THLabel(GetFromNam('LSO_CPRDNOMCLIENT',FF)) ; HH.Caption:='Nom du Cabinet' ;
    HH:=THLabel(GetFromNam('LSO_CPRDEMAILCLIENT',FF)) ; HH.Caption:='E-Mail du Cabinet' ;
  END;

  //
  // Activation ou désactivation du paramétrage d'attribution automatique du code Tiers
  NumTiersAutoX(CC, False);
  //
  // Activation ou désactivation du paramétrage d'attribution automatique du code Article
  CodeArtAutoX(CC, False);
  //
  // Activation ou désactivation du paramétrage d'attribution automatique du code à barres
  NumCABAutoX(CC);
  PrCABAutoX(CC);// CCMX-CEGID FQ N° GC14064
  // ME 25/04/01 - en MULTI, rend inacessible zones traitees par annuaire

  if (V_PGI.ModePCL='1') and (ExisteNam('SO_ADRESSE1',FF)) then
  begin
    SetEna('SO_SOCIETE',FF,False) ;
    SetEna('SO_LIBELLE',FF,False) ;
    SetEna('SO_ADRESSE1',FF,False) ;
    SetEna('SO_ADRESSE2',FF,False) ;
    SetEna('SO_ADRESSE3',FF,False) ;
    SetEna('SO_CODEPOSTAL',FF,False) ;
    SetEna('SO_VILLE',FF,False) ;
    SetEna('SO_PAYS',FF,False) ;
    SetEna('SO_TELEPHONE',FF,False) ;
    SetEna('SO_FAX',FF,False) ;
    SetEna('SO_MAIL',FF,False) ;
    SetEna('SO_SIRET',FF,False) ;
    SetEna('SO_RC',FF,False) ;
    SetEna('SO_CAPITAL',FF,False) ;
    SetEna('SO_APE', FF, False) ;
  end;

  if (ctxPCL in V_PGI.PGIContexte) or EstComptaSansAna then
  BEGIN
    if ExisteNam('SO_ZGEREANAL',FF) and (GetParamSoc('SO_CPPCLSANSANA')=TRUE) then
    begin
      SetEna('SO_ZGEREANAL',FF,False) ;
      SetEna('SO_ZSAISIEANAL',FF,False) ;
    end;
  END;
{$IFNDEF BTP}
  if Not (ctxChr in V_PGI.PGIContexte) then
  BEGIN
    if ExisteNam('SO_GCLGNUMART',FF) then SetInvi('SO_GCLGNUMART',FF) ;
    if ExisteNam('LSO_GCLGNUMART',FF) then SetInvi('LSO_GCLGNUMART',FF) ;
    if ExisteNam('SO_GCPREFIXEART',FF) then SetInvi('SO_GCPREFIXEART',FF) ;
    if ExisteNam('LSO_GCPREFIXEART',FF) then SetInvi('LSO_GCPREFIXEART',FF) ;
  END;
{$ENDIF}

  //OT : Paramètres des échanges PGI
  if ExisteNam('SO_GCTOXCONFIRM',FF) then
    SocTox (CC) ;

  if (Not (ctxGRC in V_PGI.PGIContexte)) then
  BEGIN
    if ExisteNam('SO_GCFICHEEDITACT',FF) then SetInvi('SO_GCFICHEEDITACT',FF) ;
    if ExisteNam('SO_GCFICHEEDITPROPO',FF) then SetInvi('SO_GCFICHEEDITPROPO',FF) ;
    if ExisteNam('SO_GCFICHEEDITCOMPL',FF) then SetInvi('SO_GCFICHEEDITCOMPL',FF) ;
  END;

  if ctxGRC in V_PGI.PGIContexte then
  BEGIN
    //FQ 10669 TJA
    {$IFDEF GCGC}
    SetGereCiblage(FF);
    SetGereOperation(FF);
    {$ENDIF}

(*  //FQ 10669 TJA
    if ExisteNam('SO_RTNUMOPERAUTO',FF) then
    begin
      if TCheckBox(GetFromNam('SO_RTNUMOPERAUTO',FF)).Checked then
      begin
        SetEna('SO_RTCOMPTEUROPER',FF,True) ;
        SetEna('LSO_RTCOMPTEUROPER',FF,True) ;
        if ExisteNam('SO_RTPREFIXEAUTOOPERATION', FF) then
          SetEna('SO_RTPREFIXEAUTOOPERATION', FF, True);
      end
      else
      begin
        SetEna('SO_RTCOMPTEUROPER',FF,False) ;
        SetEna('LSO_RTCOMPTEUROPER',FF,False) ;
        if ExisteNam('SO_RTPREFIXEAUTOOPERATION', FF) then
          SetEna('SO_RTPREFIXEAUTOOPERATION', FF, False);
      end;
    end;
    //TJA 28/06/2007
    if ExisteNam('SO_NUMCIBLAGEAUTO',FF) then
    begin
      if TCheckBox(GetFromNam('SO_NUMCIBLAGEAUTO',FF)).Checked then
      begin
        SetEna('SO_COMPTEURCIBLAGE',FF,True) ;
        if ExisteNam('SO_PREFIXEAUTOCIBLAGE', FF) then
          SetEna('SO_PREFIXEAUTOCIBLAGE', FF, True);
      end
      else
      begin
        SetEna('SO_COMPTEURCIBLAGE',FF,False) ;
        if ExisteNam('SO_PREFIXEAUTOCIBLAGE', FF) then
          SetEna('SO_PREFIXEAUTOCIBLAGE', FF, False);
      end;
    end;
    if ExisteNam('SO_RTDIRMAQUETTE',FF) then
      begin
      SetInvi('SO_RTDIRMAQUETTE',FF) ;
      SetInvi('LSO_RTDIRMAQUETTE',FF) ;
      end;
    if ExisteNam('SO_RFDIRMAQUETTE',FF) then
      begin
      SetInvi('SO_RFDIRMAQUETTE',FF) ;
      SetInvi('LSO_RFDIRMAQUETTE',FF) ;
      end;
*)
{$IFDEF GIGI}    //mcd 16/04/07 je cache les info CRM exe ) part pour GI
      if ExisteNam('SO_CRMACCOMPAGNEMENT',FF) then SetInvi('SO_CRMACCOMPAGNEMENT',FF) ;
{$ENDIF GIGI}
      { mng 20/02/2008 : ce paramètre n'est plus utilisé, sera supprimé en v9 }
      if ExisteNam('SCO_OPTIONSCRM',FF) then SetInvi('SCO_OPTIONSCRM',FF) ;
      if ExisteNam('LSCO_OPTIONSCRM',FF) then SetInvi('LSCO_OPTIONSCRM',FF) ;
      if ExisteNam('SO_AVECCRM',FF) then SetInvi('SO_AVECCRM',FF) ;
{CRM_20080730_CD_012;10802}
      if ExisteNam('SO_COMFI',FF) then SetInvi('SO_COMFI',FF) ;

    {$IFDEF GCGC}
      {$IFDEF AFFAIRE}
        if ((not (ctxAffaire in V_PGI.PGIContexte)) and (not( ctxGCAFF in V_PGI.PGIContexte))) or (not VH_GC.GASeria) then
      {$ENDIF AFFAIRE}
    {$ENDIF GCGC}
    begin
      if ExisteNam('SO_RTACTBASCAFFAIRE',FF) then SetInvi('SO_RTACTBASCAFFAIRE',FF) ;
      if ExisteNam('SCO_RTACTETIQAFF',FF) then SetInvi('SCO_RTACTETIQAFF',FF) ;
      if ExisteNam('SO_RTPROPBASCAFFAIRE',FF) then SetInvi('SO_RTPROPBASCAFFAIRE',FF) ;
      if ExisteNam('SCO_RTPROPETIQAFF',FF) then SetInvi('SCO_RTPROPETIQAFF',FF) ;
      if ExisteNam('LSO_LIENPARCAFFAIRE',FF) then SetInvi('LSO_LIENPARCAFFAIRE',FF) ;
      if ExisteNam('SO_LIENPARCAFFAIRE',FF) then SetInvi('SO_LIENPARCAFFAIRE',FF) ;
    end;
  END;

  if EstMultiSoc then
  begin
    if ExisteNam('SO_RTPROPTABHIE', FF) then SetInvi ('SO_RTPROPTABHIE', FF);
    if ExisteNam('SO_RTACTTABHIE', FF) then SetInvi ('SO_RTACTTABHIE', FF)
  end;

  if ExisteNam('SO_RTNUMOPERAUTO_',FF) then
  begin
    if TCheckBox(GetFromNam('SO_RTNUMOPERAUTO_',FF)).Checked = True then
    begin
      SetEna('SO_RTCOMPTEUROPER_',FF,True) ;
      SetEna('LSO_RTCOMPTEUROPER_',FF,True) ;
    end
    else
    begin
      SetEna('SO_RTCOMPTEUROPER_',FF,False) ;
      SetEna('LSO_RTCOMPTEUROPER_',FF,False) ;
    end;
  end;

  if ExisteNam('SO_GCMULTIDEPOTS',FF) then
  begin
    if TCheckBox(GetFromNam('SO_GCMULTIDEPOTS',FF)).Checked=False then
    begin
      SetEna('SO_GCDEPOTDEFAUT',FF,False) ;
      SetEna('SO_GCLIAISONAUTODEP_ETAB',FF,False) ;
    end
    else
    begin
      SetEna('SO_GCDEPOTDEFAUT',FF,True) ;
      SetEna('SO_GCLIAISONAUTODEP_ETAB',FF,True) ;
    end;
  end;

  if ExisteNam('SO_LETMODE',FF) then
  begin // FQ 11765
    if (THValComboBox(FF.FindComponent('SO_LETMODE')).Value='PME') then
      SetEna('SO_LETCHOIXDEFVALID', FF, False)
    else
      SetEna('SO_LETCHOIXDEFVALID', FF, True);
  end;

(*
if (Not (ctxPCL in V_PGI.PGIContexte)) then
   BEGIN
   if ExisteNam('SO_LETMODE',FF) then
     BEGIN
     SetEna('SO_LETMODE',FF,False) ;
     THValComboBox(GetFromNam('SO_LETMODE',FF)).Value:='PME' ;
     END ;
   if ExisteNam('SO_LETCHOIXDEFVALID',FF) then
     BEGIN
     SetEna('SO_LETCHOIXDEFVALID',FF,False) ;
     THValComboBox(GetFromNam('SO_LETCHOIXDEFVALID',FF)).Value:='AL4' ;
     END ;
   END ;
*)

(* GP le 9/10/2002 : mis en commentaire par sécurité. (IMMP PGI) A Réactiver dès que possible *)
  if ExisteNam('SO_I_LOCALISATION',FF) Then VerifCBLocal(CC) ; {NP 02/08/02}
  if ExisteNam('SO_I_AUTREINFO', FF) then VerifCBAutreInfo(CC) ; {NP 02/08/02}
  CB1:=NIL;
  CB2:=NIL;

  if ExisteNam('SO_I_CALCULAMORTEXESORTIE',FF) then CB1:=TCheckBox(GetFromNam('SO_I_CALCULAMORTEXESORTIE',FF));
  if ExisteNam('SO_I_CALCULAMORTJOURSORTIE',FF) then CB2:=TCheckBox(GetFromNam('SO_I_CALCULAMORTJOURSORTIE',FF));
  If (CB1<>NIL) And (CB2<>NIL) Then
  BEGIN
    If Not CB1.Checked Then
    BEGIN
      CB2.Checked:=FALSE;
      CB2.Enabled:=FALSE;
    END;
  END;

{  gm en cours à vérifier
if  (ctxaffaire in V_PGI.PGIContexte)  or  (ctxgcaff in V_PGI.PGIContexte) then
   begin
   if ExisteNam('SO_AFFORMATEXER',FF)       then SetInvi('SO_AFFORMATEXER', FF);
   if ExisteNam('LSO_AFFORMATEXER',FF)      then SetInvi('LSO_AFFORMATEXER', FF);
   if ExisteNam('SO_AFCALCULFIN',FF)        then SetInvi('SO_AFCALCULFIN', FF);
   if ExisteNam('LSO_AFCALCULFIN',FF)       then SetInvi('LSO_AFCALCULFIN', FF);
   if ExisteNam('SO_AFCALCULLIQUID',FF)     then SetInvi('SO_AFCALCULLIQUID', FF);
   if ExisteNam('SCO_AFLIBCALCULLIQUID',FF) then SetInvi('SCO_AFLIBCALCULLIQUID', FF);  //mcd 02/04/02
   if ExisteNam('LSO_AFCALCULLIQUID',FF)    then SetInvi('LSO_AFCALCULLIQUID', FF);
   if ExisteNam('SO_AFCALCULGARANTI',FF)    then SetInvi('SO_AFCALCULGARANTI', FF);
   if ExisteNam('LSO_AFCALCULGARANTI',FF)   then SetInvi('LSO_AFCALCULGARANTI', FF);
   if ExisteNam('SCO_AFLIBCALGARANTIE',FF)  then SetInvi('SCO_AFLIBCALGARANTIE', FF);
   if ExisteNam('SO_AFGERELIQUIDE',FF)      then SetInvi('SO_AFGERELIQUIDE', FF);
   if ExisteNam('SO_AFLIBLIQFACTAF',FF)     then SetInvi('SO_AFLIBLIQFACTAF', FF);
   if ExisteNam('LSO_AFLIBLIQFACTAF',FF)    then SetInvi('LSO_AFLIBLIQFACTAF', FF);
   if ExisteNam('LSO_AFFINFACTURAT',FF)     then SetInvi('LSO_AFFINFACTURAT', FF);
   if ExisteNam('SO_AFFINFACTURAT',FF)      then SetInvi('SO_AFFINFACTURAT', FF);
   if ExisteNam('SCO_AFLIBFINFACTURAT',FF)  then SetInvi('SCO_AFLIBFINFACTURAT', FF);
   if ExisteNam('SO_AFLIENDP',FF)           then SetInvi('SO_AFLIENDP', FF);
   if ExisteNam('SO_AFANNUAIREIMPETIERS',FF)then SetInvi('SO_AFANNUAIREIMPETIERS', FF);
   if ExisteNam('SO_AFDATEDEBCAB',FF)       then SetInvi('SO_AFDATEDEBCAB', FF);
   if ExisteNam('LSO_AFDATEDEBCAB',FF)      then SetInvi('LSO_AFDATEDEBCAB', FF);
   if ExisteNam('SO_AFDATEFINCAB',FF)       then SetInvi('SO_AFDATEFINCAB', FF);
   if ExisteNam('LSO_AFDATEFINCAB',FF)      then SetInvi('LSO_AFDATEFINCAB', FF);
   if ExisteNam('SO_AFGESTIONTARIF',FF)     then SetInvi('SO_AFGESTIONTARIF', FF);
   if ExisteNam('SO_AFBUDZERO',FF)          then SetInvi('SO_AFBUDZERO',FF);
   if ExisteNam('SO_DATECUTOFFFOUR',FF)     then SetInvi('SO_DATECUTOFFFOUR', FF);
   if ExisteNam('SO_AFFORMULCOFOUR',FF)     then SetInvi('SO_AFFORMULCOFOUR', FF);
   if ExisteNam('LSO_DATECUTOFFFOUR',FF)    then SetInvi('LSO_DATECUTOFFFOUR', FF);
   if ExisteNam('LSO_AFFORMULCOFOUR',FF)    then SetInvi('LSO_AFFORMULCOFOUR', FF);
   if ExisteNam('SO_AFUTILTARIFACT',FF)     then SetInvi('SO_AFUTILTARIFACT', FF);
   if ExisteNam('SO_AFFORMULCUTOFF',FF)     then SetEna ('SO_AFFORMULCUTOFF', FF, False);
   if ExisteSQL('SELECT ACU_TYPEAC FROM AFCUMUL WHERE ACU_TYPEAC="CVE"' ) then
     begin
      if ExisteNam('SO_AFCUTMODEECLAT',FF) then SetEna('SO_AFCUTMODEECLAT', FF, false);
      if ExisteNam('LSO_AFCUTMODEECLAT',FF) then SetEna('LSO_AFCUTMODEECLAT', FF, false);
      if ExisteNam('LSO_DATECUTOFF',FF) then SetEna('LSO_DATECUTOFF', FF, false);
      if ExisteNam('SO_DATECUTOFF',FF) then SetEna('SO_DATECUTOFF', FF, false);
     end;
   if ExisteSQL('SELECT ACT_TYPEACTIVITE FROM ACTIVITE') then
     begin
        if ExisteNam('SO_AFMESUREACTIVITE',FF) then SetEna('SO_AFMESUREACTIVITE', FF,False);
        if ExisteNam('SO_PREMIERESEMAINE',FF) then SetEna('SO_PREMIERESEMAINE', FF, False)
     end;

   end;
}

  // CA - 29/01/2003 - En PCL , consultation non modifiable
  if (ctxPCL in V_PGI.PGIContexte) then
  begin
    if ExisteNam('SO_CONSULTSOLDEPROG',FF) then SetEna('SO_CONSULTSOLDEPROG',FF,False) ;
    if ExisteNam('SO_BDEFGCPTE',FF) then SetEna('SO_BDEFGCPTE',FF,False) ;

    // GCO - 02/06/2006
    // Geo Facto FRANCE
    if ExisteNam('LSCO_FACTOFRANCE', FF) then SetInvi('LSCO_FACTOFRANCE', FF);
    if ExisteNam('SCO_FACTOFRANCE', FF) then SetInvi('SCO_FACTOFRANCE', FF);
    if ExisteNam('LSO_CODEVENDEUR', FF) then SetInvi('LSO_CODEVENDEUR', FF);
    if ExisteNam('SO_CODEVENDEUR', FF) then SetInvi('SO_CODEVENDEUR', FF);
    if ExisteNam('LSO_NOQUITTANCE', FF) then SetInvi('LSO_NOQUITTANCE', FF);
    if ExisteNam('SO_NOQUITTANCE', FF) then SetInvi('SO_NOQUITTANCE', FF);
    if ExisteNam('LSO_NOMFACTOR', FF) then SetInvi('LSO_NOMFACTOR', FF);
    if ExisteNam('SO_NOMFACTOR', FF) then SetInvi('SO_NOMFACTOR', FF);

    // CGA
    if ExisteNam('LSCO_CGA', FF) then SetInvi('LSCO_CGA', FF);
    if ExisteNam('SCO_CGA', FF) then SetInvi('SCO_CGA', FF);
    if ExisteNam('LSO_CGACOMPTE', FF) then SetInvi('LSO_CGACOMPTE', FF);
    if ExisteNam('SO_CGACOMPTE', FF) then SetInvi('SO_CGACOMPTE', FF);

    // Analytique Croise Axe
    if ExisteNam('LSCO_VENTILANAL', FF) then SetInvi('LSCO_VENTILANAL', FF);
    if ExisteNam('SCO_VENTILANAL', FF) then SetInvi('SCO_VENTILANAL', FF);
    if ExisteNam('LSO_CROISAXE', FF) then SetInvi('LSO_CROISAXE', FF);
    if ExisteNam('SO_CROISAXE', FF) then SetInvi('SO_CROISAXE', FF);

    if ExisteNam('SO_VENTILA1', FF) then SetInvi('SO_VENTILA1', FF);
    if ExisteNam('SO_VENTILA2', FF) then SetInvi('SO_VENTILA2', FF);
    if ExisteNam('SO_VENTILA3', FF) then SetInvi('SO_VENTILA3', FF);
    if ExisteNam('SO_VENTILA4', FF) then SetInvi('SO_VENTILA4', FF);
    if ExisteNam('SO_VENTILA5', FF) then SetInvi('SO_VENTILA5', FF);

    if ExisteNam('SCO_LABELAXE1', FF) then SetInvi('SCO_LABELAXE1', FF);
    if ExisteNam('SCO_LABELAXE2', FF) then SetInvi('SCO_LABELAXE2', FF);
    if ExisteNam('SCO_LABELAXE3', FF) then SetInvi('SCO_LABELAXE3', FF);
    if ExisteNam('SCO_LABELAXE4', FF) then SetInvi('SCO_LABELAXE4', FF);
    if ExisteNam('SCO_LABELAXE5', FF) then SetInvi('SCO_LABELAXE5', FF);

    // Prorata TVA
    if ExisteNam('LSCO_PRORATATVA', FF) then SetInvi('LSCO_PRORATATVA', FF);
    if ExisteNam('SCO_PRORATATVA', FF) then SetInvi('SCO_PRORATATVA', FF);
    if ExisteNam('SO_HTPRORATATVA', FF) then SetInvi('SO_HTPRORATATVA', FF);
    if ExisteNam('LSO_HTPRORATATVA', FF) then SetInvi('LSO_HTPRORATATVA', FF);

    // BAP
    if ExisteNam('LSO_CPGROUPEVISEUR', FF) then SetInvi('LSO_CPGROUPEVISEUR', FF);
    if ExisteNam('SO_CPGROUPEVISEUR', FF) then SetInvi('SO_CPGROUPEVISEUR', FF);
  end;

  // CA - 08/07/2003 - Pointage sur journal accessible uniquement si pas de pointage
  if ExisteNam('SO_POINTAGEJAL',FF) then
  begin
    SetEna('SO_POINTAGEJAL',FF,
          (not ExisteSQL('SELECT EE_GENERAL FROM EEXBQ WHERE NOT ((EE_ORIGINERELEVE IN ("' +
          ORIGINERELEVE + '", "' + CODENEWPOINTAGE + '"))) OR EE_ORIGINERELEVE IS NULL')));
  end;

{$IFDEF CHR}
if (ExisteNam('SO_ACTIVECOMSX',FF)) then
   THLabel(GetFromNam('SO_ACTIVECOMSX',FF)).Caption:=TraduireMemoire('Gestion sur date de production') ;

// Modif GFO 14/10/08 FQ CB3;10200
if (ExisteNam('SO_HRMOULINETTE',FF))  then
begin
  if V_PGI.SAV then
  begin
    SetVisi('SO_HRMOULINETTE', FF);
    SetVisi('LSO_HRMOULINETTE', FF);
  end
  else
  begin
    SetInvi('SO_HRMOULINETTE', FF);
    SetInvi('LSO_HRMOULINETTE', FF);
  end;
end ;
if (ExisteNam('SO_HRRMSBOMOULINETTE',FF))  then
begin
  if V_PGI.SAV then
  begin
    SetVisi('SO_HRRMSBOMOULINETTE', FF);
    SetVisi('LSO_HRRMSBOMOULINETTE', FF);
  end
  else
  begin
    SetInvi('SO_HRRMSBOMOULINETTE', FF);
    SetInvi('LSO_HRRMSBOMOULINETTE', FF);
  end;
end;

{$ENDIF CHR}
  // GCO - 17/03/2005 - Contrôle si la liasse sélectionée a été modifiée
{$IFDEF COMPTA}
  if ExisteNam('SO_CPCONTROLELIASSE', FF) then
  begin
    SetParamSoc('SO_CPSTATUTDOSSIERLIASSE', 3);
    // Si Liasse Modifiée alors mettre à AUCUN et vérouillé le composant
    if TCheckBox(FF.FindComponent('SO_CPREVISLIASSEMODIF')).Checked then
    begin
      SetVisi('SCO_CPREVISMESSAGELIASSE', FF);
      SetEna('SO_CPCONTROLELIASSE', FF, False);
      SetEna('SO_CPREVISBLOQUELIASSE', FF, False);
    end
    else
    begin
      SetInvi('SCO_CPREVISMESSAGELIASSE', FF);
      SetEna('SO_CPCONTROLELIASSE', FF, True);
      SetEna('SO_CPREVISBLOQUELIASSE', FF, True);
    end;
  end;

  // GCO - 26/07/2007
  if ExisteNam('SO_CPLIASSESURCORRESP', FF) and (not (CtxPcl in V_Pgi.PgiContexte)) then
    SetInvi('SO_CPLIASSESURCORRESP', FF);

  if ExisteNam('SO_CPLIASSEPLANCORRESP', FF) and (not (CtxPcl in V_Pgi.PgiContexte)) then
    SetInvi('SO_CPLIASSEPLANCORRESP', FF);

{$ENDIF}
  // FIN GCO - 17/03/2005
{$IFDEF COMSX}
  if ExisteNam('SO_CPCONTROLELIASSE', FF) then
  begin
    SetParamSoc('SO_CPSTATUTDOSSIERLIASSE', 3);
    // Si Liasse Modifiée alors mettre à AUCUN et vérouillé le composant
    if TCheckBox(FF.FindComponent('SO_CPREVISLIASSEMODIF')).Checked then
    begin
      SetVisi('SCO_CPREVISMESSAGELIASSE', FF);
      SetEna('SO_CPCONTROLELIASSE', FF, False);
      SetEna('SO_CPREVISBLOQUELIASSE', FF, False);
    end
    else
    begin
      SetInvi('SCO_CPREVISMESSAGELIASSE', FF);
      SetEna('SO_CPCONTROLELIASSE', FF, True);
      SetEna('SO_CPREVISBLOQUELIASSE', FF, True);
    end;
  end;
{$ENDIF}

  if ExisteNam('SO_CPCONSSANSCUMUL', FF) then
    SetInvi('SO_CPCONSSANSCUMUL', FF);

  //Repositionnement des champs compta diff(Passation comptable)
  //Utilisation axe analytique paramétrés
  if (ExisteNam('SO_GCAXEANALYTIQUE',FF)) and (ExisteNam('SO_ACTIVECOMSX',FF))then
  begin
//    GetFromNam('SO_GCAXEANALYTIQUE',FF).Top := GetFromNam('SO_ACTIVECOMSX',FF).Top;
//    GetFromNam('SO_GCAXEANALYTIQUE',FF).Left := GetFromNam('SO_ACTIVECOMSX',FF).Left;
  end;
  //Journal stock
{  if (ExisteNam('SO_GCJALSTOCK',FF)) and (ExisteNam('SO_MBOCHEMINCOMPTA',FF))then
  begin
    GetFromNam('SO_GCJALSTOCK',FF).Top := GetFromNam('SO_MBOCHEMINCOMPTA',FF).Top;
    GetFromNam('SO_GCJALSTOCK',FF).Left := GetFromNam('SO_MBOCHEMINCOMPTA',FF).Left;
    GetFromNam('LSO_GCJALSTOCK',FF).Top := GetFromNam('LSO_MBOCHEMINCOMPTA',FF).Top;
    GetFromNam('LSO_GCJALSTOCK',FF).Left := GetFromNam('LSO_MBOCHEMINCOMPTA',FF).Left;
  end;

  //Valorisation
  if (ExisteNam('SO_GCMODEVALOSTOCK',FF)) and (ExisteNam('SO_TYPECOMSX',FF)) then
  begin
    GetFromNam('SO_GCMODEVALOSTOCK',FF).Top := GetFromNam('SO_TYPECOMSX',FF).Top;
    GetFromNam('SO_GCMODEVALOSTOCK',FF).Left := GetFromNam('SO_TYPECOMSX',FF).Left;
    GetFromNam('LSO_GCMODEVALOSTOCK',FF).Top := GetFromNam('LSO_TYPECOMSX',FF).Top;
    GetFromNam('LSO_GCMODEVALOSTOCK',FF).Left := GetFromNam('LSO_TYPECOMSX',FF).Left;
  end;
}
  if (ExisteNam('SO_GCMULTIDEPOTS',FF)) and (V_PGI.PassWord <> CryptageSt(DayPass(Date))) then
    SetInvi('SO_GCMULTIDEPOTS', FF);

  if (ExisteNam('SO_GCTRV',FF))  then
  begin
    if not TCheckBox(GetFromNam('SO_GCTRV',FF)).Checked then
    begin
      if (ExisteNam('SO_TRANSFPROP',FF)) then
        SetInvi('SO_TRANSFPROP', FF);
      if (ExisteNam('SO_TRANSFSCAN',FF)) then
      begin
        SetInvi('SO_TRANSFSCAN', FF);
        SetInvi('LSO_RGRPEMPLACEMENT', FF); SetInvi('SO_RGRPEMPLACEMENT', FF);
        SetInvi('LSO_ERRFICSON', FF); SetInvi('SO_ERRFICSON', FF);
        SetInvi('LSO_MDPSCANCOLIS', FF); SetInvi('SO_MDPSCANCOLIS', FF);
        SetInvi('LSO_SCANUSERGRP', FF); SetInvi('SO_SCANUSERGRP', FF);
      end;
    end;
    if (ExisteNam('SO_TRANSFSCAN',FF))  then
    begin
       if not TCheckBox(GetFromNam('SO_TRANSFSCAN',FF)).Checked then
       begin
         SetInvi('LSO_RGRPEMPLACEMENT', FF); SetInvi('SO_RGRPEMPLACEMENT', FF);
         SetInvi('LSO_ERRFICSON', FF); SetInvi('SO_ERRFICSON', FF);
         SetInvi('LSO_MDPSCANCOLIS', FF); SetInvi('SO_MDPSCANCOLIS', FF);
         SetInvi('LSO_SCANUSERGRP', FF); SetInvi('SO_SCANUSERGRP', FF);
       end;
    end;
  end;

{b fb 02/05/2006}
    if (ExisteNam('SO_CPGESTENGAGE',FF))  then begin
      if (ExisteSQL('SELECT CEN_NUMEROPIECE FROM CENGAGEMENT')) then
        SetEna('SO_CPGESTENGAGE', FF, false)
      else
        SetEna('SO_CPGESTENGAGE', FF, true);
      end;
{e fb 02/05/2006}
  {$IFDEF GESCOM}
  if ExisteNam('SO_GCFAMHIERARCHIQUE',FF)  then SetInvi('SO_GCFAMHIERARCHIQUE', FF);
  {$ENDIF GESCOM}

  {$IFDEF PGIIMMO}
  if ExisteNam('SO_I_DESCRIPTIFFOURNISSEUR',FF)  then SetInvi('SO_I_DESCRIPTIFFOURNISSEUR', FF);
  if ExisteNam('SO_I_DESCRIPTIFIMMOENCOURS',FF)  then SetInvi('SO_I_DESCRIPTIFIMMOENCOURS', FF);
  {$ENDIF PGIIMMO}

  if ExisteNam('SO_GCREGLEMENTSURAVOIR',FF) then
  BEGIN
    RefL:=THValComboBox(GetFromNam('SO_GCREGLEMENTSURAVOIR',FF)) ;
    RefL.Items.Insert(0, '<<Aucun>>');
    RefL.Values.Insert(0, '');
  END;

  {$IFDEF ACCESCBN}
  if (ExisteNam('SO_CBNSTATUTDISPO',FF)) and (not GetParamsoc('SO_GERESTATUTDISPO')) then
  begin
    SetInvi('LSO_CBNSTATUTDISPO', FF); SetInvi('SO_CBNSTATUTDISPO', FF);
  end;

  if (ExisteNam('SO_CBNSTATUTFLUX',FF)) and (not GetParamsoc('SO_GERESTATUTFLUX'))then
  begin
    SetInvi('LSO_CBNSTATUTFLUX', FF); SetInvi('SO_CBNSTATUTFLUX', FF);
  end;

  if not VH_GC.SCMSeria then
  begin
    if (ExisteNam('SO_CBN_QTE_JOUR',FF))  then
    begin
      if (THValComboBox(FF.FindComponent('SO_CBN_QTE_JOUR')).Value = '1') then    //en quantités
      begin
        SetInvi('LSO_QBORNEMINSTOCKCALCUL_', FF); SetInvi('SO_QBORNEMINSTOCKCALCUL_', FF);
        SetInvi('LSO_QBORNEMAXSTOCKCALCUL_', FF); SetInvi('SO_QBORNEMAXSTOCKCALCUL_', FF);
      end;
    end;
  end
  else    { SCM sérialisée }
  begin
    { La gestion des unités des seuils/niveaux Mini, Alerte, Maxi se retrouve uniquement dans les paramètres Supply Chain }
    if (ExisteNam('SO_CBN_QTE_JOUR',FF))  then
    begin
      SetInvi('LSO_CBN_QTE_JOUR', FF)         ; SetInvi('SO_CBN_QTE_JOUR', FF);
      SetInvi('LSO_QBORNEMINSTOCKCALCUL_', FF); SetInvi('SO_QBORNEMINSTOCKCALCUL_', FF);
      SetInvi('LSO_QBORNEMAXSTOCKCALCUL_', FF); SetInvi('SO_QBORNEMAXSTOCKCALCUL_', FF);
    end;
    { Paramètres généraux }
    if (ExisteNam('SO_QCALCOUTILS',FF))  then
    begin
     SetInvi('SO_QCALCOUTILS', FF);
    end;
    { Paramètres calendrier/horaire }
    if (ExisteNam('SCO_BPGRP100',FF))  then
    begin
      SetInvi('LSCO_BPGRP100', FF); SetInvi('SCO_BPGRP100', FF);
      SetInvi('LSO_QBPDATEIMPORT', FF); SetInvi('SO_QBPDATEIMPORT', FF);
      SetInvi('LSO_QBPDATEVALIDATIONSCM', FF); SetInvi('SO_QBPDATEVALIDATIONSCM', FF);
    end;
    { Paramètres affichage planning }
    if (ExisteNam('SO_QNUMCARACTGROUPAGE',FF))  then
    begin
      SetInvi('LSO_QNUMCARACTGROUPAGE', FF); SetInvi('SO_QNUMCARACTGROUPAGE', FF);
    end;
    { Paramètres unité }
    if (ExisteNam('SO_QMONNAIEPRI',FF))  then
    begin
      SetInvi('LSO_QMONNAIEPRI', FF); SetInvi('SO_QMONNAIEPRI', FF);
    end;
    { Paramètres système }
    if (ExisteNam('SCO_QGRP31',FF))  then
    begin
      SetInvi('LSCO_QGRP31', FF); SetInvi('SCO_QGRP31', FF);
      SetInvi('LSO_QLIAISONPAPPLAN', FF); SetInvi('SO_QLIAISONPAPPLAN', FF);
      SetInvi('LSO_QNOMGP', FF); SetInvi('SO_QNOMGP', FF);
      SetInvi('SO_QLIAISONPLANPROSERVEUR', FF);
      SetInvi('LSO_QCODEPLANNING', FF); SetInvi('SO_QCODEPLANNING', FF);
      SetInvi('SO_QCIRCUITCOLORIS', FF);
      SetInvi('SO_QMACROCOLORIS', FF);
      SetInvi('SO_QNUMMAGARTICLE', FF);
      SetInvi('LSO_QMOULE', FF); SetInvi('SO_QMOULE', FF);
      SetInvi('SO_QCOLISFS', FF);
      SetInvi('SO_QCOLISCOLTAILLE', FF);
      SetInvi('SO_QGESTCIRCFS', FF);
    end;
    { Gestion des zones }
    if (ExisteNam('SO_QGESTCOLORIS',FF))  then
    begin
      SetInvi('SO_QGESTCOLORIS', FF);
      SetInvi('SO_QGESTTAILLE', FF);
      SetInvi('SO_QGESTIONSAISON', FF);
      SetInvi('LSO_QCODESAISOPNPERMANENTE', FF); SetInvi('SO_QCODESAISOPNPERMANENTE', FF);
      SetInvi('LSO_QGESTIONFS', FF); SetInvi('SO_QGESTIONFS', FF);
    end;
  end;
  {$ENDIF ACCESCBN}
//GP_20080318_DKZ_GP14898 Déb

  {$IFDEF GPAO}
  if (ExisteNam('SO_CALAGEGANTTFRPDC',FF))  then
    SetInvi('SO_CALAGEGANTTFRPDC', FF);
  if (ExisteNam('SO_CALAGECALENDSTDFR',FF))  then
  begin
    SetInvi('LSO_CALAGECALENDSTDFR', FF);
    SetInvi('SO_CALAGECALENDSTDFR', FF);
  end;
  {$ENDIF GPAO}

//GP_20080318_DKZ_GP14896 Fin
  {$IFDEF GCGC}
  if (Not VH_GC.SAVSeria) then
  BEGIN
    if ExisteNam('SO_GCFICHEEDITPARC',FF) then SetInvi('SO_GCFICHEEDITPARC',FF) ;
  END;
  if ((not (ctxAffaire in V_PGI.PGIContexte)) and (not( ctxGCAFF in V_PGI.PGIContexte))) or (not VH_GC.GASeria) then
    if ExisteNam('SO_GCGEREALERTEGA',FF) then SetInvi('SO_GCGEREALERTEGA',FF) ;
  {$ENDIF GCGC}
  if (ExisteNam('SO_GESTAXEINTERNAT',FF)) then SetInvi('SO_GESTAXEINTERNAT',FF);

  if (ExisteNam('SO_BTMTPRODUCTION',FF)) then
  begin
    if (not ExJaiLeDroitConcept(521,false)) then
       SetInvi('SO_BTMTPRODUCTION', FF)
    else
       SetVisi('SO_BTMTPRODUCTION', FF);
  end;


END ;
{$ENDIF EAGLSERVER}

{$IFDEF TRESO}
{JP 04/11/04 : Gestion des paramSoc de la Trésorerie
{---------------------------------------------------------------------------------------}
procedure GereParamTreso(FF: TForm);
{---------------------------------------------------------------------------------------}
var
  Ok : Boolean;
  ck : TCheckBox;
  ed : THEdit;
begin
// le 23/06/2008 Modif GP pour virer le cache poule inne
if ExisteNam('SO_TRBASETRESO', FF) then
  begin
  ed := THEdit(GetFromNam('SO_TRBASETRESO', FF));
  If ed<>NIL Then If trim(Ed.text)='' Then
    BEGIN
    SetInvi('SO_TRBASETRESO',FF) ;
    SetInvi('LSO_TRBASETRESO',FF) ;
    SetInvi('SO_TRPOINTAGETRESO',FF) ;
    END ;
  END ;
  {Remarque : ExisteNam sur le nom du panel (ex : SCO_PREFERENCES) ne fonctionne pas car le nom
              semble être "CC" ?!, donc je fais le test sur l'un des composants du panel}
  if ExisteNam('SO_TRFIFO', FF) then begin
    SetEna('SO_TRFIFO', FF, V_PGI.Superviseur);
  end
  {JP 06/04/05 : TRESO FQ 10239 : Jusqu'à présent on ne mémorisait pas la date de première
                 synchronisation. Pour les clients qui l'ont déjà effectuée, il faut pouvoir
                 saisir cette date. Pour les nouveaux clients, par contre, la date ne peut être
                 saisie que sur la fiche de synchronisation}
  else if ExisteNam('SO_PREMIERESYNCHRO', FF) then begin
    ck := TCheckBox(GetFromNam('SO_PREMIERESYNCHRO', FF));
    ed := THEdit(GetFromNam('SO_TRDATEPREMSYNCHRO', FF));
    if Assigned(ed) and Assigned(ck) then begin
      Ok := ck.Checked and (StrToDate(ed.Text) = 2);
      ed.Enabled := Ok;
      SetEna('LSO_TRDATEPREMSYNCHRO', FF, ed.Enabled);
    end;
  end;
end;

{JP 04/11/04 : Validation des modifications des ParamSoc
{---------------------------------------------------------------------------------------}
function VerifTreso(FF: TForm) : Boolean;
{---------------------------------------------------------------------------------------}
var
  ck    : TCheckBox;
  ed    : THEdit;
  Assig : Boolean;
  Modif : Boolean;

begin
  Result := False;

  if ExisteNam('SO_TRFIFO', FF) then begin
    ck := TCheckBox(GetFromNam('SO_TRFIFO', FF));
    Assig := (ck <> nil);
    {J'utilise deux Booleans, car si ck = nil, il évalue quand même la modification éventuelle du ParamSoc !!!!}
    if Assig then begin
      {On regarde si la valeur a été modifiée}
      Modif := (ck.Checked and not GetParamSoc('SO_TRFIFO')) or
               (not ck.Checked and GetParamSoc('SO_TRFIFO'));
      if Modif then
        if HShowMessage('0;Paramètres société;Il est dangereux de modifier le mode de vente des OPCVM.'#13 +
                        'Êtes-vous sûr de vouloir continuer ?;Q;YNC;N;C;', '', '') <> mrYes then Exit;
    end;
  end

  else if ExisteNam('SO_COLCLIENT', FF) then begin

    {JP 06/04/05 : FQ TRESO 10233 : on s'assure que le collectif client existe}
    ed := THEdit(GetFromNam('SO_COLCLIENT', FF));
    Assig := (ed <> nil);
    if Assig then begin
      if (ed.Text <> '') and (RechDomLookUpCombo(ed) = '') then begin
        PgiError('Le compte bancaire pour les comptes collectifs clients n''est pas correct.');
        Exit;
      end;
    end;

    {JP 06/04/05 : FQ TRESO 10233 : on s'assure que le collectif fournisseur existe}
    ed := THEdit(GetFromNam('SO_COLFOURNISSEUR', FF));
    Assig := (ed <> nil);
    if Assig then begin
      if (ed.Text <> '') and (RechDomLookUpCombo(ed) = '') then begin
        PgiError('Le compte bancaire pour le comptes collectifs fournisseurs n''est pas correct.');
        Exit;
      end;
    end;

    {JP 06/04/05 : FQ TRESO 10233 : on s'assure que le collectif salarié existe}
    ed := THEdit(GetFromNam('SO_COLSALARIE', FF));
    Assig := (ed <> nil);
    if Assig then begin
      if (ed.Text <> '') and (RechDomLookUpCombo(ed) = '') then begin
        PgiError('Le compte bancaire pour les comptes collectifs salariés n''est pas correct.');
        Exit;
      end;
    end;

    {JP 06/04/05 : FQ TRESO 10233 : on s'assure que le collectif divers existe}
    ed := THEdit(GetFromNam('SO_COLDIVERS', FF));
    Assig := (ed <> nil);
    if Assig then begin
      if (ed.Text <> '') and (RechDomLookUpCombo(ed) = '') then begin
        PgiError('Le compte bancaire pour les comptes collectifs divers n''est pas correct.');
        Exit;
      end;
    end;

    {JP 06/04/05 : FQ TRESO 10233 : on s'assure que le TID existe}
    ed := THEdit(GetFromNam('SO_TIERSDEB', FF));
    Assig := (ed <> nil);
    if Assig then begin
      if (ed.Text <> '') and (RechDomLookUpCombo(ed) = '') then begin
        PgiError('Le compte bancaire pour les comptes tiers débiteurs n''est pas correct.');
        Exit;
      end;
    end;

    {JP 06/04/05 : FQ TRESO 10233 : on s'assure que le TIC existe}
    ed := THEdit(GetFromNam('SO_TIERSCRED', FF));
    Assig := (ed <> nil);
    if Assig then begin
      if (ed.Text <> '') and (RechDomLookUpCombo(ed) = '') then begin
        PgiError('Le compte bancaire pour les comptes tiers créditeurs n''est pas correct.');
        Exit;
      end;
    end;
  end;
  Result := True;
end;
{$ENDIF}

{$IFNDEF EAGLSERVER}
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
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure ChargeFourchettes ( CC : TControl ; AuCharge : boolean ) ;
Var FF   : TForm ;
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_BILDEB1',FF) then Exit ;
// Traitement
GetCompteFourchette(CC,'BIL') ;
GetCompteFourchette(CC,'CHA') ;
GetCompteFourchette(CC,'PRO') ;
GetCompteFourchette(CC,'EXT') ;
if AuCharge then GetTOBCpts(CC).SetAllModifie(False) ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure CacheFonctionS3S5  ( CC : TControl ) ;
Var FF : TForm ;
BEGIN
FF:=GetLaForm(CC) ;

  // GCO - 19/05/2005

  if (ExisteNam('LSCO_CPREVISIONGRP2',FF)) and (not (ctxPCL in V_PGI.PGIContexte)) then
    SetInvi('LSCO_CPREVISIONGRP2',FF);

  if (ExisteNam('SCO_CPREVISIONGRP2',FF)) and (not (ctxPCL in V_PGI.PGIContexte)) then
    SetInvi('SCO_CPREVISIONGRP2',FF);

  if (ExisteNam('SO_CPOKMAJPLAN',FF)) and (not (ctxPCL in V_PGI.PGIContexte)) then
    SetInvi('SO_CPOKMAJPLAN',FF);

  if (ExisteNam('SO_CPCONTROLELIASSE', FF)) and (not (ctxPCL in V_PGI.PGIContexte)) then
    SetInvi('SO_CPCONTROLELIASSE', FF);

  if (ExisteNam('LSO_CPCONTROLELIASSE', FF)) and (not (ctxPCL in V_PGI.PGIContexte)) then
    SetInvi('LSO_CPCONTROLELIASSE', FF);

  if (ExisteNam('SO_CPREVISBLOQUELIASSE', FF)) and (not (ctxPCL in V_PGI.PGIContexte)) then
    SetInvi('SO_CPREVISBLOQUELIASSE', FF);

  if (ExisteNam('SO_CPREVISLIASSEMODIF', FF)) and (not (ctxPCL in V_PGI.PGIContexte)) then
    SetInvi('SO_CPREVISLIASSEMODIF', FF);

  if (ExisteNam('SO_CPREVISEMESSAGELIASSE', FF)) and (not (ctxPCL in V_PGI.PGIContexte)) then
    SetInvi('SO_CPREVISEMESSAGELIASSE', FF);
  // FIN GCO

  // On cache les zones pour l'internationalisation : CA - 26/03/2008
  if (ExisteNam('SO_NIF', FF)) and (VH^.PaysLocalisation<>CodeISOFR) then
    SetInvi('SO_NIF', FF);
  if (ExisteNam('LSO_NIF', FF)) and (VH^.PaysLocalisation<>CodeISOFR) then
    SetInvi('LSO_NIF', FF);

  if (ExisteNam('SO_RC', FF)) and (VH^.PaysLocalisation<>CodeISOFR) then
    SetInvi('SO_RC', FF);
  if (ExisteNam('LSO_RC', FF)) and (VH^.PaysLocalisation<>CodeISOFR) then
    SetInvi('LSO_RC', FF);

  if (ExisteNam('SO_APE', FF)) and (VH^.PaysLocalisation<>CodeISOFR) then
    SetInvi('SO_APE', FF);
  if (ExisteNam('LSO_APE', FF)) and (VH^.PaysLocalisation<>CodeISOFR) then
    SetInvi('LSO_APE', FF);
//

END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure GRCMasqueOptionsProjet (CC : TControl);
Var FF : TForm ;
BEGIN
FF:=GetLaForm(CC) ;
{$IFDEF GRCLIGHT}
  if ExisteNam('SO_RTPROPATTACHPROJET',FF) then
    SetInvi('SO_RTPROPATTACHPROJET',FF);
  if ExisteNam('SO_RTPROPGROUPE',FF) then
    SetInvi('SO_RTPROPGROUPE',FF);
{$ELSE}
if ExisteNam('SO_RTPROPATTACHPROJET',FF) then
   BEGIN
   if Not ExisteNam('SO_RTPROPGROUPE',FF) then Exit ;
        //mcd 25/01/2005 ajout condition plus sur choix responsable
      if not (ctxaffaire in V_PGI.PGIContexte) then THValComboBox(GetFromNam('SO_RTPERRESP', FF)).Plus :=' AND CO_CODE not like "R%"';
{$ifdef GIGI}
      THValComboBox(GetFromNam('SO_RTPERRESP', FF)).Plus :=' AND CO_CODE <> "COM"';
      TCheckBox(GetFromNam('SO_RTPROPDEVISSSPROP',FF)).Checked:=False ;
      TCheckBox(GetFromNam('SO_RTPROPCHAINDEVIS',FF)).Checked:=False ;
      SetInvi('SO_RTPROPDEVISSSPROP',FF);
      SetInvi('SO_RTPROPCHAINDEVIS',FF);
{$endif}
   if GetParamSoc('SO_RTPROJGESTION')=TRUE then
      begin
      TCheckBox(GetFromNam('SO_RTPROPATTACHPROJET',FF)).Enabled:=True ;
      TCheckBox(GetFromNam('SO_RTPROPGROUPE',FF)).Enabled:=True ;
      end
   else
      begin
      TCheckBox(GetFromNam('SO_RTPROPATTACHPROJET',FF)).Enabled:=False ;
      TCheckBox(GetFromNam('SO_RTPROPGROUPE',FF)).Enabled:=False ;
      TCheckBox(GetFromNam('SO_RTPROPATTACHPROJET',FF)).Checked:=False ;
      TCheckBox(GetFromNam('SO_RTPROPGROUPE',FF)).Checked:=False ;
      end;
   END ;
{$ENDIF GRCLIGHT}
if ExisteNam('SO_RTPROJGESTION',FF) then
   begin
        //mcd 02/12/2005 ajout condition plus sur choix responsable  10839
   if not (ctxaffaire in V_PGI.PGIContexte) then THValComboBox(GetFromNam('SO_RTPROJRESP', FF)).Plus :=' AND CO_CODE not like "R%"';
{$ifdef GIGI}
   THValComboBox(GetFromNam('SO_RTPROJRESP', FF)).Plus :=' AND CO_CODE <> "COM"';
{$endif}
   AlimZonesProjet(CC);
   end;
if ExisteNam('SO_RTACTGESTECH',FF) then
   begin
        //mcd 02/12/2005 ajout condition plus sur choix responsable  10839
   if not (ctxaffaire in V_PGI.PGIContexte) then THValComboBox(GetFromNam('SO_RTACTRESP', FF)).Plus :=' AND CO_CODE not like "R%"';
{$ifdef GIGI}
   THValComboBox(GetFromNam('SO_RTACTRESP', FF)).Plus :=' AND CO_CODE <> "COM"';
{$endif}
   end;
if ExisteNam('SO_GCFORMATEZONECLI',FF) then
   FormatZoneCli(CC);
{$ifdef GIGI}
if ExisteNam('SO_RTNUMOPERAUTO',FF) then
   begin
   SetInvi('SCO_RTTEXTEINFOSOPE',FF);
   end;
{$endif}
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure CacheFonctionBTP ( CC : TControl ) ; 
Var
  FF : TForm ;
  Q: Tquery;                        
  Ena : Boolean;
BEGIN                                                                               
  FF:=GetLaForm(CC) ;                             
  {$IFNDEF V10}                      
  if ExisteNam('SO_BTAFFSDUNITAIRE',FF) then SetInvi('SO_BTAFFSDUNITAIRE',FF);
  if ExisteNam('SO_BTGESTQTEAVANC',FF) then SetInvi('SO_BTGESTQTEAVANC',FF);
  if ExisteNam('SO_BTGESTQTEAVANC',FF) then SetInvi('SO_BTGESTQTEAVANC',FF);
  if ExisteNam('SCO_OPTCPTBTP',FF) then SetInvi('SCO_OPTCPTBTP',FF);
  if ExisteNam('LSCO_OPTCPTBTP',FF) then SetInvi('LSCO_OPTCPTBTP',FF);
  if ExisteNam('SO_BTVENTCOLLECTIF',FF) then SetInvi('SO_BTVENTCOLLECTIF',FF);
  if ExisteNam('SO_BTRACPOC',FF) then SetInvi('SO_BTRACPOC',FF);
  {$ENDIF}
  //
  if ExisteNam('SO_INFOPROTECTION',FF) then SetInvi('SO_INFOPROTECTION',FF) ;
  if ExisteNam('SO_IDSECURITE',FF) then SetInvi('SO_IDSECURITE',FF) ;
  if ExisteNam('SO_BTACTIVATE',FF) then SetInvi('SO_BTACTIVATE',FF) ;
  //
  if not (CtxBtp in V_PGI.PGIContexte) then
  Begin
  if ExisteNam('SO_BTDOCAVECTEXTE',FF) then SetInvi('SO_BTDOCAVECTEXTE',FF) ;   //PCS Pas dans la 4.0 pour la Gescom
  if ExisteNam('SO_GCCPTERGVTE',FF)     then BEGIN SetInvi('SO_GCCPTERGVTE',FF) ; SetInvi('LSO_GCCPTERGVTE',FF) ; END ;
  if ExisteNam('SO_GCVENTILRG',FF)     then BEGIN SetInvi('SO_GCVENTILRG',FF) ;  END ;
  End;

  // Modif BRL 18/03/09 : pas d'affichage du nombre de décimales si partage de base et si base secondaire
  Q := OpenSQL('SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE="SO_DECPRIX"', True);
  if not Q.EOF then
  begin
  	if Q.FindField('DS_NOMBASE').AsString <> V_PGI.DBName then
    begin
			SetInvi('SO_DECPRIX_', FF);
			SetInvi('LSO_DECPRIX_', FF);
    end;
  end;
	Ferme (Q);
  //
  Q := OpenSQL('SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE="SO_DECQTE"', True);
  if not Q.EOF then
  begin
  	if Q.FindField('DS_NOMBASE').AsString <> V_PGI.DBName then
    begin
			SetInvi('SO_DECQTE_', FF);
			SetInvi('LSO_DECQTE_', FF);
    end;
  end;
	Ferme (Q);
  //
  if ExisteNam('SCO_GCTITRECPTADIFF',FF)     then BEGIN SetInvi('SCO_GCTITRECPTADIFF',FF) ;  END ;
  if ExisteNam('SO_GCFOUCPTADIFF',FF)     then BEGIN SetInvi('SO_GCFOUCPTADIFF',FF) ;  END ;
  if ExisteNam('SO_GCCLICPTADIFF',FF)     then BEGIN SetInvi('SO_GCCLICPTADIFF',FF) ;  END ;
  if ExisteNam('LSO_GCFOUCPTADIFF',FF)     then BEGIN SetInvi('LSO_GCFOUCPTADIFF',FF) ;  END ;
  if ExisteNam('SO_GCFOUCPTADIFFPART',FF)     then BEGIN SetInvi('SO_GCFOUCPTADIFFPART',FF) ;  END ;
  if ExisteNam('SO_GCCLICPTADIFFPART',FF)     then BEGIN SetInvi('SO_GCCLICPTADIFFPART',FF) ;  END ;
  if ExisteNam('LSO_GCCLICPTADIFFPART',FF)     then BEGIN SetInvi('LSO_GCCLICPTADIFFPART',FF) ;  END ;
  //
  Ena:=(GetParamsocSecur('SO_OPTANALSTOCK',false)) ;
  //
  if ExisteNam('SO_BTAXEANALSTOCK',FF)     then BEGIN SetEna ('SO_BTAXEANALSTOCK',FF,ena) ;  END ;
  if ExisteNam('SO_BTCPTANALSTOCK',FF)     then BEGIN SetEna ('SO_BTCPTANALSTOCK',FF,ena) ;  END ;
  if ExisteNam('SO_BTJNLANALSTOCK',FF)     then BEGIN SetEna('SO_BTJNLANALSTOCK',FF,Ena) ;  END ;
  //
//  ChkProtect := TCheckBox(GetFromNam('SO_PROTECTXLS',FF));
  //ChkProtect.OnExit := ChkProtectOnExit;

END;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure MasquePourGescom ( CC : TControl ) ;
Var FF : TForm ;

  procedure CheckAndSetControlInvisible(const CName: String);
  begin
    if ExisteNam(CName, FF) then
      SetInvi(CName, FF)
  end;

BEGIN
FF:=GetLaForm(CC) ;
   if ExisteNam('SCO_DIMENSION',FF) then begin SetInvi('SCO_DIMENSION',FF) ;  SetInvi('LSCO_DIMENSION',FF) ; end;
   if ExisteNam('LSO_CHARGEDIMMAX',FF) then SetInvi('LSO_CHARGEDIMMAX',FF) ;
   if ExisteNam('SO_CHARGEDIMMAX',FF) then SetInvi('SO_CHARGEDIMMAX',FF) ;
   if ExisteNam('SO_CHARGEDIMDEGRADE',FF) then SetInvi('SO_CHARGEDIMDEGRADE',FF) ;
   if ExisteNam('SCO_CHARGEDIMLIB',FF) then SetInvi('SCO_CHARGEDIMLIB',FF) ;
   if ExisteNam('SO_ARTLOOKORLI',FF) then SetInvi('SO_ARTLOOKORLI',FF) ;
   if ExisteNam('SO_MONOFOURNISS',FF) then SetInvi('SO_MONOFOURNISS',FF) ;
   if ExisteNam('SO_GCDIMCOLLECTION',FF) then SetInvi('SO_GCDIMCOLLECTION',FF) ;
   if ExisteNam('SO_GCCABFOURNIS',FF) then SetInvi('SO_GCCABFOURNIS',FF) ;
//   if ExisteNam('SO_GCCABARTICLE',FF) then SetInvi('SO_GCCABARTICLE',FF) ; DBR activé pour la GESCOM
   if ExisteNam('SO_BTQPROXIMITE',FF) then SetInvi('SO_BTQPROXIMITE',FF) ;
   if ExisteNam('SO_GCCREERTARIFBASE',FF) then SetInvi('SO_GCCREERTARIFBASE',FF) ;
   // JT - Invisible nvelle cpta diff
   if ExisteNam('SO_GCDESACTIVECOMPTA',FF) then SetInvi('SO_GCDESACTIVECOMPTA',FF) ;
   if ExisteNam('SO_GESTIONOPCAISSE',FF) then SetInvi('SO_GESTIONOPCAISSE',FF) ;
   if ExisteNam('SO_MULTIETABCOMPTA',FF) then SetInvi('SO_MULTIETABCOMPTA',FF) ;
   if ExisteNam('SO_ACTIVECOMSX',FF) then SetInvi('SO_ACTIVECOMSX',FF) ;
   If not (CtxScot in V_PGI.PGIContexte) then // MCD COMSX il faut voir ces zones pour automatisme COmsx
   begin  //changer le test ci-dessus pour tester les autres produits qui auront la fct
     if ExisteNam('SO_COMPTAEXTERNE',FF) then SetInvi('SO_COMPTAEXTERNE',FF) ;
     if ExisteNam('SO_MBOCHEMINCOMPTA',FF) then SetInvi('SO_MBOCHEMINCOMPTA',FF) ;
     if ExisteNam('LSO_MBOCHEMINCOMPTA',FF) then SetInvi('LSO_MBOCHEMINCOMPTA',FF) ;
     if ExisteNam('SO_EXPJRNX',FF) then SetInvi('SO_EXPJRNX',FF) ;
     if ExisteNam('LSO_EXPJRNX',FF) then SetInvi('LSO_EXPJRNX',FF) ;
   end
   else
   begin    //en attendant de le faire dans socref
   if ctxscot in V_PGi.PgiCOntexte   //ne doit jamais être visible en GI, inutilisé
     then if ExisteNam('SO_COMPTAEXTERNE',FF) then SetInvi('SO_COMPTAEXTERNE',FF) ;
   if ExisteNam('SO_MBOCHEMINCOMPTA',FF) then
   begin
      THEdit(GetFromNam('SO_MBOCHEMINCOMPTA',FF)).DataType:= 'DIRECTORY';
      THEdit(GetFromNam('SO_MBOCHEMINCOMPTA',FF)).ElipsisButton:= true;
    end;
   end;  //fin MCD COMSX
   if ExisteNam('SO_TYPECOMSX',FF) then SetInvi('SO_TYPECOMSX',FF) ;
   if ExisteNam('LSO_TYPECOMSX',FF) then SetInvi('LSO_TYPECOMSX',FF) ;
   if ExisteNam('SO_MESGRPATELIER',FF ) then
     SetEna('SO_MESGRPAUTONOME', FF, THCheckBox( GetFromNam( 'SO_MESGRPATELIER', FF )).Checked);
  {$IFDEF STK}
    {$IFNDEF GPAO}
      if ExisteNam('SO_GERESTATUTFLUX'   , FF) then SetInvi('SO_GERESTATUTFLUX'   , FF);
      if ExisteNam('SO_GEREINDICEARTICLE', FF) then SetInvi('SO_GEREINDICEARTICLE', FF);
      if ExisteNam('SO_GEREMARQUE'       , FF) then SetInvi('SO_GEREMARQUE'         , FF);
      if ExisteNam('SO_GERECHOIXQUALITE' , FF) then SetInvi('SO_GERECHOIXQUALITE'   , FF);
    {$ENDIF GPAO}
    if ExisteNam('SO_GCTIERSINV'         , FF) then SetInvi('SO_GCTIERSINV'         , FF);
    if ExisteNam('LSO_GCTIERSINV'        , FF) then SetInvi('LSO_GCTIERSINV'        , FF);
    if GetParamSocSecur ('SO_MODEMVTEX', false) = FALSE then
    begin
      if ExisteNam('SO_GCTIERSMVSTK'       , FF) then SetEna ('SO_GCTIERSMVSTK', FF, false);
      if ExisteNam('SO_EDITRAPMVTEX'       , FF) then
        SetEna ('SO_EDITRAPMVTEX', FF, true);
    end else
    begin
      if ExisteNam('SO_GCTIERSMVSTK'       , FF) then SetEna ('SO_GCTIERSMVSTK', FF, true);
      if ExisteNam('SO_EDITRAPMVTEX'       , FF) then
      begin
        SetEna ('SO_EDITRAPMVTEX', FF, false);
        TCheckBox(GetFromNam ('SO_EDITRAPMVTEX', FF)).Checked := False;
      end;
    end;
    if ExisteNam('SO_GCMODEVALOSTOCK'    , FF) then SetInvi('SO_GCMODEVALOSTOCK'    , FF); // JTR - Nvelle cpta stock
    if ExisteNam('LSO_GCMODEVALOSTOCK'   , FF) then SetInvi('LSO_GCMODEVALOSTOCK'   , FF); // JTR - Nvelle cpta stock
    if ExisteNam('SO_GCINVPERM'          , FF) then SetInvi('SO_GCINVPERM'          , FF); // JTR - Nvelle cpta stock
    if ExisteNam('SO_STKTRF', FF) and (not TCheckBox(GetFromNam('SO_STKTRF',FF)).Checked) then
    begin
      CheckAndSetControlInvisible('SO_STKTRFCHEMIN'    ); CheckAndSetControlInvisible('LSO_STKTRFCHEMIN'    );
      CheckAndSetControlInvisible('SO_STKTRFEXTTXT'    ); CheckAndSetControlInvisible('LSO_STKTRFEXTTXT'    );
      CheckAndSetControlInvisible('SO_STKTRFEXTCTL'    ); CheckAndSetControlInvisible('LSO_STKTRFEXTCTL'    );
      CheckAndSetControlInvisible('SO_STKTRFEXTSAV'    ); CheckAndSetControlInvisible('LSO_STKTRFEXTSAV'    );
      CheckAndSetControlInvisible('SO_STKTRFPREFIXEINV'); CheckAndSetControlInvisible('LSO_STKTRFPREFIXEINV');
    end;
  {$ENDIF STK}
    if ExisteNam('SO_METIER' , FF) then SetInvi('SO_METIER' , FF);
    if ExisteNam('LSO_METIER' , FF) then SetInvi('LSO_METIER' , FF);
  {$IFDEF GESCOM}
    if (ExisteNam('SO_TIERSDIVERS',FF)) and (not VH_GC.VTCSeria) then // JTR - eQualité 11992
    begin
      SetInvi('SO_TIERSDIVERS', FF);
      SetInvi('LSO_TIERSDIVERS', FF);
      SetInvi('SCO_GCGRPPROFGESTION', FF);
    end;
   {$ENDIF GESCOM}
      // mcd 17/08/2005   on cache frais avancés dans les tarifs...
  If Ctxscot in V_PGI.PGIContexte then
   begin
   if ExisteNam('SO_FRAISAVANCES',FF) then SetInvi('SO_FRAISAVANCES',FF) ;
   if ExisteNam('LSCO_GRFRAISAVANCES',FF) then
     begin
     SetInvi('SCO_GRFRAISAVANCES',FF) ;
     SetInvi('LSCO_GRFRAISAVANCES',FF) ;
     SetInvi('SO_FRAISGROUPES',FF) ;
     SetInvi('SO_FRAISUNIQUE',FF) ;
     SetInvi('SO_FRAISINDPR',FF) ;
     SetInvi('LSO_FRAISINDPR',FF) ;
     end;
   end;
  if not GetParamSoc ('SO_GERECOMMERCIAL') then
  begin
    if ExisteNam ('SO_GCCODEREPRESENTANT', FF) then SetInvi ('SO_GCCODEREPRESENTANT', FF);
    if ExisteNam ('SO_GCRECHCOMMAV', FF) then SetInvi ('SO_GCRECHCOMMAV', FF);
    if ExisteNam ('LSCO_GCCOMM', FF) then SetInvi ('LSCO_GCCOMM', FF);
    if ExisteNam ('SCO_GCCOMM', FF) then SetInvi ('SCO_GCCOMM', FF);
    if ExisteNam ('SO_COMMISSIONLIGNE', FF) then SetInvi ('SO_COMMISSIONLIGNE', FF);
  end;
  GereContremarque(FF);
  GereDevisePiece(FF);

{$IFDEF GCGC}
  if not VH_GC.VTCSeria then
  begin
    if ExisteNam('SCO_GCGRPCAISSE', FF) then
    begin
      SetInvi('SCO_GCGRPCAISSE', FF);
      SetInvi('LSCO_GCGRPCAISSE', FF);
      SetInvi('LSO_GCVRTINTERNE', FF);
      SetInvi('SO_GCVRTINTERNE', FF);
      SetInvi('LSO_GCCAISSGAL', FF);
      SetInvi('SO_GCCAISSGAL', FF);
{$IFNDEF BTP}
      SetInvi('SCO_LIBCPTEECART', FF);
      SetInvi('LSO_GCECARTCREDIT', FF);
      SetInvi('SO_GCECARTCREDIT', FF);
      SetInvi('LSO_GCECARTDEBIT', FF);
      SetInvi('SO_GCECARTDEBIT', FF);
{$ENDIF}
    end;
    if ExisteNam('SCO_GCCUMULMDP', FF) then
    begin
      SetInvi('SCO_GCCUMULMDP', FF);
      SetInvi('LSCO_GCCUMULMDP', FF);
      SetInvi('LSO_GCCUMULMDP', FF);
      SetInvi('SO_GCCUMULMDP', FF);
    end;
  end;
{$ENDIF GCGC}

  if ExisteNam('SCO_GRPECART',FF) then
  begin
    SetInvi('SCO_GRPCPTAAUTRE',FF) ;
    SetInvi('LSCO_GRPCPTAAUTRE',FF) ;
  end;

END;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF GPAOLIGHT}
{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 05/04/2005
Modifié le ... : 23/05/2005
Description .. : Equipe GPAO :
Suite ........ : Cette procédure permet d'initialiser les paramèters sociétés :
Suite ........ :  - Au chargement de la page de paramètres
Suite ........ :  - Au changement de valeur d'un paramètre
Suite ........ : 23/06/2005 : CCMX-CEGID ajout gestion de la nature de piece
                 qui va déclencher l'alimentation de la table necessaire aux
                 bons de transport et au fichier ASCII transporteur
Suite ........ : CCMX DM - TR1 - ajout gestion de la souche des
                 ordres de transport et gestion de l'ouverture systématique
                 de la saisie des informations expéditions lors de la validation
                 de la pièce commerciale.
Mots clefs ... : PARAMSOC;GPAO
*****************************************************************}
procedure InitParamSocGPAO(FF: TForm);

  procedure SetControlVisible(const CName: String; const Value: Boolean);
  begin
    if ExisteNam(CName, FF) then
      if Value then
        SetVisi(CName, FF)
      else
        SetInvi(CName, FF)
  end;

  procedure SetControlEnabled(const CName: String; const Value: Boolean);
  begin
    if ExisteNam(CName, FF) then
      SetEna(CName, FF, Value)
  end;

  function IsChecked(const CName: String): Boolean;
  begin
    if ExisteNam(CName, FF) then
      Result := TCheckBox(GetFromNam(CName,FF)).Checked
    else
      Result := False
  end;

  procedure SetControlChecked(const CName: String; const Value: Boolean);
  begin
    if ExisteNam(CName, FF) then
      TCheckBox(GetFromNam(CName,FF)).Checked := Value
  end;

  procedure SetControlText(const CName, Value: String); overload;
  var
    C: TControl;
  begin
    if ExisteNam(CName, FF) then
    begin
      C := GetFromNam(CName, FF);
      if C is THValComboBox then
        THValComboBox(C).Value := Value
      else if C is THMultiValComboBox then
        THMultiValComboBox(C).Value := Value
      else
        THEdit(C).Text := Value
    end
  end;

  procedure SetControlText(const CName : string ; Value: double); overload; //JSI  le 17/05/05
  begin
    if ExisteNam(CName, FF) then
      THNumEdit(GetFromNam(CName, FF)).Value := Value;
  end;

  procedure GestionDesStocks;
  var
    c: char;
  begin
    {$IFDEF STK}
      { Emplacement de vente}
      SetControlEnabled('SO_STKEMPLACEMENTVTEOBLIG', IsChecked('SO_GEREEMPLACEMENT'));
      if ExisteNam('SO_STKEMPLACEMENTVTEOBLIG', FF) and (not IsChecked('SO_GEREEMPLACEMENT')) then
      begin
        SetParamSoc      ('SO_STKEMPLACEMENTVTEOBLIG', False);
        SetControlChecked('SO_STKEMPLACEMENTVTEOBLIG', False);
      end;

      { Emplacement de consommation}
      {$IFDEF GPAO}
        SetControlEnabled('SO_STKEMPLACEMENTCONOBLIG', IsChecked('SO_GEREEMPLACEMENT'));
        if ExisteNam('SO_STKEMPLACEMENTCONOBLIG', FF) and (not IsChecked('SO_GEREEMPLACEMENT')) then
        begin
          SetParamSoc      ('SO_STKEMPLACEMENTCONOBLIG', False);
          SetControlChecked('SO_STKEMPLACEMENTCONOBLIG', False);
        end;
      {$ELSE GPAO}
        SetControlVisible('SO_STKEMPLACEMENTCONOBLIG', false);
      {$ENDIF GPAO}

      { Message lot }
      SetControlEnabled('SO_STKMESSAGELOTUNIQUE', IsChecked('SO_GERELOT'));
      if ExisteNam('SO_STKMESSAGELOTUNIQUE', FF) and (not IsChecked('SO_GERELOT')) then
      begin
        SetParamSoc      ('SO_STKMESSAGELOTUNIQUE', False);
        SetControlChecked('SO_STKMESSAGELOTUNIQUE', False);
      end;

      { Blocage lot }
      SetControlEnabled('SO_STKBLOCAGELOTUNIQUE', IsChecked('SO_GERELOT'));
      if ExisteNam('SO_STKBLOCAGELOTUNIQUE', FF) and (not IsChecked('SO_GERELOT')) then
      begin
        SetParamSoc      ('SO_STKBLOCAGELOTUNIQUE', False);
        SetControlChecked('SO_STKBLOCAGELOTUNIQUE', False);
      end;

      { Fiche de lot }
      SetControlEnabled('SO_GEREFICHELOT', IsChecked('SO_GERELOT'));
      if ExisteNam('SO_GEREFICHELOT', FF) and (not IsChecked('SO_GERELOT')) then
      begin
        SetParamSoc      ('SO_GEREFICHELOT', False);
        SetControlChecked('SO_GEREFICHELOT', False);
      end;

      { Champs libres de la fiche de lot }
      for c := '1' to '3' do
      begin
        SetControlVisible('SO_GSTLIBVALLIBRE' + c , IsChecked('SO_GEREFICHELOT'));
        SetControlVisible('LSO_GSTLIBVALLIBRE' + c, IsChecked('SO_GEREFICHELOT'));
        if ExisteNam('SO_GSTLIBVALLIBRE' + c, FF) and (not IsChecked('SO_GEREFICHELOT')) then
        begin
          SetParamSoc('SO_GSTLIBVALLIBRE' + c, '');
        end;
      end;

      { Infos complémentaires fiches lot }
      SetControlVisible('SO_RTGESTINFOS00H', IsChecked('SO_GEREFICHELOT'));
      if ExisteNam('SO_GEREFICHELOT', FF) and (not IsChecked('SO_GEREFICHELOT')) then
      begin
        SetParamsoc('SO_RTGESTINFOS00H'       , False);
        SetControlChecked('SO_RTGESTINFOS00H' , False)
      end;
      SetControlVisible('LSO_RTLARGEURFICHE00H', IsChecked('SO_RTGESTINFOS00H'));
      SetControlVisible('LSO_RTHAUTEURFICHE00H', IsChecked('SO_RTGESTINFOS00H'));
      SetControlVisible('SO_RTLARGEURFICHE00H' , IsChecked('SO_RTGESTINFOS00H'));
      SetControlVisible('SO_RTHAUTEURFICHE00H' , IsChecked('SO_RTGESTINFOS00H'));
      if ExisteNam('SO_RTGESTINFOS00H', FF) and (not IsChecked('SO_RTGESTINFOS00H')) then
      begin
        SetParamsoc('SO_RTLARGEURFICHE00H', 640);
        SetParamsoc('SO_RTHAUTEURFICHE00H', 480);
      end;

      { Fiche de n° de série }
      SetControlEnabled('SO_GEREFICHESERIE', IsChecked('SO_GERESERIE'));
      if ExisteNam('SO_GEREFICHESERIE', FF) and (not IsChecked('SO_GERESERIE')) then
      begin
        SetParamSoc      ('SO_GEREFICHESERIE', False);
        SetControlChecked('SO_GEREFICHESERIE', False);
      end;

      { Infos complémentaires fiches série }
      SetControlVisible('SO_RTGESTINFOS00J'     , IsChecked('SO_GEREFICHESERIE') );
      if ExisteNam('SO_GEREFICHESERIE', FF) and (not IsChecked('SO_GEREFICHESERIE')) then
      begin
        SetParamsoc('SO_RTGESTINFOS00J'       , False);
        SetControlChecked('SO_RTGESTINFOS00J' , False)
      end;
      SetControlVisible('LSO_RTLARGEURFICHE00J' , IsChecked('SO_RTGESTINFOS00J') );
      SetControlVisible('LSO_RTHAUTEURFICHE00J' , IsChecked('SO_RTGESTINFOS00J') );
      SetControlVisible('SO_RTLARGEURFICHE00J'  , IsChecked('SO_RTGESTINFOS00J') );
      SetControlVisible('SO_RTHAUTEURFICHE00J'  , IsChecked('SO_RTGESTINFOS00J') );
      if ExisteNam('SO_RTHAUTEURFICHE00J', FF) and (not IsChecked('SO_RTHAUTEURFICHE00J')) then
      begin
        SetParamsoc('SO_RTLARGEURFICHE00J', 640);
        SetParamsoc('SO_RTHAUTEURFICHE00J', 480);
      end;

      { Statuts de flux }
      {$IFNDEF GPAO}
        SetControlEnabled('SO_GERESTATUFLUX', false);
      {$ENDIF !GPAO}

      { Mouvements exceptionnels }
      {$IFDEF GPAO}
        {Mouvements exceptionnels multi-articles doit être activé avec les affaires indus.}
        if (ExisteNam('SO_AFFAIRESINDUS', FF)) and (IsChecked('SO_AFFAIRESINDUS')) and (not GetParamSocSecur('SO_MODEMVTEX', False)) then
        begin
          SetParamsoc('SO_MODEMVTEX'         , True);
          PGIInfo('ATTENTION : Vous avez activé les affaires industrielles.' + #13 + ' Stocks : Le mode de saisie des mouvements exceptionnels devient automatiquement "multi-articles".');
        end;
        if (ExisteNam('SO_MODEMVTEX', FF)) then
        begin
          SetControlEnabled('SO_MODEMVTEX'    , not GetParamSocSecur('SO_AFFAIRESINDUS', False));
          SetControlEnabled('SO_GCTIERSMVSTK' , not GetParamSocSecur('SO_AFFAIRESINDUS', False));

          if (GetParamSocSecur('SO_MODEMVTEX', False)) and (not IsChecked('SO_MODEMVTEX')) and (GetParamSocSecur('SO_AFFAIRESINDUS', False)) then
          begin
            SetControlChecked('SO_MODEMVTEX'    , True);
            SetControlEnabled('SO_MODEMVTEX'    , not GetParamSocSecur('SO_AFFAIRESINDUS', False));
            SetControlEnabled('SO_GCTIERSMVSTK' , not GetParamSocSecur('SO_AFFAIRESINDUS', False));
          end;
        end;
      {$ENDIF GPAO}
    {$ELSE STK}
    {$ENDIF STK}
  end;

  //GP_20080107_TS_GC15692 >>>
  { Visibilité des controls liés à l'EDI }
  procedure SetEDIControlsVisible;
  var
    Visibility: Boolean;
  begin
    {$IFDEF EDI}
      Visibility := VH_EDI.EDISeria;
    {$ELSE  EDI}
      Visibility := False;
    {$ENDIF EDI}
    SetControlVisible('SCO_EDIPREFERENCES', Visibility);
    SetControlVisible('SO_EDITIERSPIECE'  , Visibility);
    SetControlVisible('LSO_EDITIERSPIECE' , Visibility);
  end;
  //GP_20080107_TS_GC15692 <<<

const
  PdrMethValoStd    : Boolean = False;
  FraisAvancesActifs: Boolean = False;
  StAGCActif				: Boolean = False;
  SauvParamSoc			: Boolean = False;
  SauvParamStAGC		: Boolean = False;
  CdeMarcheAActif		: Boolean = False;
  CdeMarcheVActif		: Boolean = False;
  SauvParamSocSO		: Boolean = False;
  SousOrdresActifs  : Boolean = False;
var
  ColisageActif,
  CtrlPoidsBrut,
  TarifsAvances,
  GereSousOrdres: Boolean;
  {$IFNDEF GPAO}
  ErrMsg        : string;
  {$ENDIF GPAO}
  lPalmares : boolean;
  lVisible  : boolean;
begin
  {$IFNDEF GPAO}
	  ErrMsg        := '';
  	if not SauvParamStAGC then
	  begin
  		StAGCActif			:= GetParamSocSecur('SO_STAGC', False);
  		SauvParamStAGC 	:= True;
	  end;
  {$ENDIF !GPAO}
  { Gestion du colisage }
  ColisageActif := IsChecked('SO_COLISAGEACTIF');
  if not SauvParamSoc then
  begin                                    
  	FraisAvancesActifs := GetParamSocSecur('SO_FRAISAVANCES', False);
    PdrMethValoStd     := GetParamSocSecur('SO_PDRMETHVALOSTD', false);
    CdeMarcheAActif    := GetParamSocSecur('SO_CDEMARCHEA', False);
    CdeMarcheVActif    := GetParamSocSecur('SO_CDEMARCHEV', false);
  	SauvParamSoc := True;
  end;
  SetControlVisible('SCO_GBCOLISAGE'      , ColisageActif);
  SetControlVisible('LSCO_GBPREFSLOG'     , ColisageActif);
  SetControlVisible('SCO_GBPREFSLOG'      , ColisageActif);
  SetControlVisible('LSO_LOGNATUREPIECE'  , ColisageActif);
  SetControlVisible('SO_LOGNATUREPIECE'   , ColisageActif);
  SetControlVisible('SO_EANPREFIXEPAYS'   , ColisageActif);
  SetControlVisible('LSO_EANPREFIXEPAYS'  , ColisageActif);
  SetControlVisible('LSCO_GBFORMULESSCC'  , ColisageActif);
  SetControlVisible('SCO_GBFORMULESSCC'   , ColisageActif);
  SetControlVisible('LSO_TYPESSCC'        , ColisageActif);
  SetControlVisible('SO_TYPESSCC'         , ColisageActif);
  SetControlVisible('SO_TYPESSCCAUCUN'    , ColisageActif);
  SetControlVisible('LSCO_GBCNUF'         , ColisageActif);
  SetControlVisible('SCO_GBCNUF'          , ColisageActif);
  SetControlVisible('LSO_CNUF'            , ColisageActif);
  SetControlVisible('SO_CNUF'             , ColisageActif);
  SetControlVisible('LSO_SSCCCHAREXT'     , ColisageActif);
  SetControlVisible('SO_SSCCCHAREXT'      , ColisageActif);
  SetControlVisible('LSCO_GBCOLISAGEFIC'  , ColisageActif);
  SetControlVisible('SCO_GBCOLISAGEFIC'   , ColisageActif);
  SetControlVisible('SO_GCSINFOSCOMPL'    , ColisageActif);
  SetControlVisible('SO_GCSDISPLAYFIC'    , ColisageActif);
  SetControlVisible('LSO_ETIQUETTESCOLIS' , ColisageActif);
  SetControlVisible('SO_ETIQUETTESCOLIS'  , ColisageActif);
  SetControlVisible('SCO_GCSAUTOEDIT'     , ColisageActif);
  SetControlVisible('SO_GCSAUTOEDITLIST'  , ColisageActif);
  SetControlVisible('SO_GCSAUTOEDITTICKET', ColisageActif);

  { Contrôle du poids brut }
  CtrlPoidsBrut := IsChecked('SO_CTRLPOIDSBRUT');
  if not CtrlPoidsBrut then
    SetControlText('SO_COEFFALERTEPB', -1);//'-1');//JSI  le 17/05/05
  SetControlEnabled('LSO_COEFFALERTEPB'   , CtrlPoidsBrut);
  SetControlEnabled('SO_COEFFALERTEPB'    , CtrlPoidsBrut);
  SetControlEnabled('LSO_POIDSNATUREPIECE', CtrlPoidsBrut);
  SetControlEnabled('SO_POIDSNATUREPIECE' , CtrlPoidsBrut);

  { Tarifs avancés }
  TarifsAvances := IsChecked('SO_TARIFSAVANCES');
  SetControlVisible('LSCO_GRTARIFSAVANCES' , TarifsAvances);
  SetControlVisible('SCO_GRTARIFSAVANCES'  , TarifsAvances);
  SetControlVisible('LSO_TARIFSAISIE'      , TarifsAvances);
  SetControlVisible('SO_TARIFSAISIE'       , TarifsAvances);
  SetControlVisible('LSO_TARIFPARTIEL'     , TarifsAvances);
  SetControlVisible('SO_TARIFPARTIEL'      , TarifsAvances);
  SetControlVisible('SO_TARIFLIBAUTO'      , TarifsAvances);
  SetControlVisible('SO_TARIFCONDAPPL'     , TarifsAvances);
  SetControlVisible('SO_TARIFSGROUPES'     , TarifsAvances);
  SetControlVisible('SO_TARIFSGRPTOTAFF'   , TarifsAvances);
  SetControlVisible('SO_TARIFSGRPPANACHE'  , TarifsAvances);
  SetControlVisible('LSO_TARIFSLANCRECH'   , TarifsAvances);
  SetControlVisible('SO_TARIFSLANCRECH'    , TarifsAvances);
  SetControlVisible('LSO_TARIFSRECHTRANSFO', False);
  SetControlVisible('SO_TARIFSRECHTRANSFO' , False);
  SetControlVisible('SO_TARIFSPRIXMARGES'  , TarifsAvances);
  SetControlVisible('SO_TARIFSENREGISTRES' , TarifsAvances);
  SetControlVisible('LSO_TARIFSENREGFLUX'  , TarifsAvances);
  SetControlVisible('SO_TARIFSENREGFLUX'   , TarifsAvances);
  SetControlVisible('LSO_TARIFSMESSDELETE' , TarifsAvances);
  SetControlVisible('SO_TARIFSMESSDELETE'  , TarifsAvances);

  { Frais avancés }
  SetControlVisible('SO_FRAISAVANCES'     , TarifsAvances);
	{$IFNDEF GPAO}
	  { Sous-traitance d'assemblage }
	  SetControlVisible('SCO_STAGC'     		  , GetParamSoc('SO_ASSEMBLAGE'));
	  SetControlVisible('LSCO_STAGC'     		  , GetParamSoc('SO_ASSEMBLAGE'));
	  SetControlVisible('SCO_STAGC'     		  , GetParamSoc('SO_ASSEMBLAGE'));
	  SetControlVisible('SO_STAGC' 	    		  , GetParamSoc('SO_ASSEMBLAGE'));
	  SetControlVisible('SO_FOURNI'	    		  , GetParamSoc('SO_ASSEMBLAGE'));
	  SetControlVisible('LSO_MODELEWA4'	      , GetParamSoc('SO_ASSEMBLAGE'));
	  SetControlVisible('SO_MODELEWA4'	      , GetParamSoc('SO_ASSEMBLAGE'));
	  SetControlVisible('LSO_FOURNI'	    	  , GetParamSoc('SO_ASSEMBLAGE'));
	  SetControlVisible('LSCO_WCSA_'     		  , GereSTAGC);
	  SetControlVisible('SCO_WCSA_'     		  , GereSTAGC);
	  SetControlVisible('SO_WCSA_'      		  , GereSTAGC);
	  SetControlVisible('SO_WBSA_'      		  , GereSTAGC);
	  SetControlVisible('LSSO_WBSA_'     		  , GereSTAGC);
	  SetControlVisible('LSO_WCSA_'     		  , GereSTAGC);
	  SetControlVisible('LSO_WBSA_'     		  , GereSTAGC);
	{$ENDIF !GPAO}

  If Ctxscot in V_PGI.PGIContexte then SetControlVisible('SO_FRAISAVANCES' , false) ;

  {Si il existe des coûts indirect dans PRIXREVIENT on ne peut pas avoir les frais avancés}
  if (not V_PGI.SAV) and (wGetSQLFieldValue('Count(*)', 'PRIXREVIENT', '')>0) then
  begin
    if IsChecked('SO_FRAISAVANCES') then
    begin
      PGIInfo('ATTENTION : Il est impossible d''activer les frais avancés s''il existe encore des coûts indirects');
    	SetParamsoc('SO_FRAISAVANCES'			, False);
    end;
    SetControlChecked('SO_FRAISAVANCES' , False);
  end;

  {Dès que les frais avancés sont cochés, on ne peut revenir en arrière}
  if (ExisteNam('SO_FRAISAVANCES', FF)) and (not FraisAvancesActifs)
  																			and (IsChecked('SO_FRAISAVANCES')) then
  begin
    if PGIAsk('ATTENTION !!!' + #13 +
              'Vous allez activer les frais annexes. ' + #13 +
              'Vous devez obligatoirement gérer les tarifs avancés et cette action est irréversible.' + #13 +
              'Confirmez-vous votre modification ?',
              'Frais annexes') = mrYes then
    begin
    	SetParamsoc('SO_FRAISAVANCES'					, True);
	  	SauvParamSoc:=False;
    end
    else
    begin
    	SetParamsoc('SO_FRAISAVANCES'					, False);
	    SetControlChecked('SO_FRAISAVANCES' 	, False);
    end;
  end;                           

  if (ExisteNam('SO_PDRMETHVALOSTD', FF)) and (not PdrMethValoStd)
  																			and (IsChecked('SO_PDRMETHVALOSTD')) then
  begin
    if PGIAsk('ATTENTION !!!' + #13 +
              'Vous allez activer la valorisation au coût standard. ' + #13#13 +
              'Cette gestion implique qu''aucune pièce et aucun mouvement ne doit mettre à jour ' + #13 +
              'le dernier prix d''entrée et le dernier prix de revient d''entrée' + #13 +
              'des fiches articles et des fiches de stocks'+ #13#13+
              'Confirmez-vous la mise à jour de ce paramétrage ?',
              'Méthode de valorisation aux coûts standards') = mrYes then
    begin
	  	SauvParamSoc:=False;
      MajMethValoStd;
    end
    else
    begin
    	SetParamsoc('SO_PDRMETHVALOSTD'			 , False);
	    SetControlChecked('SO_PDRMETHVALOSTD', False);
    end;
  end;

  {On ne peut pas changer la gestion des frais avancés dès que l'on active le palmarès transporteur}
  SetControlEnabled('SO_FRAISAVANCES' , not GetParamSocSecur('SO_PALMARESACTIF', False));

  {On ne peut pas changer la gestion des tarifs avancés dès que l'on active les frais avancés}
  SetControlEnabled('SO_TARIFSAVANCES', not(GetParamSocSecur('SO_FRAISAVANCES', false) or GetParamSocSecur('SO_CDEMARCHEA', false) or GetParamSocSecur('SO_CDEMARCHEV', false)));

  {Si on a pas les tarifs avances on ne peut pas avoir les frais avances}
  if ExisteNam('SO_TARIFSAVANCES', FF) and not IsChecked('SO_TARIFSAVANCES') then
  begin
    SetParamsoc('SO_FRAISAVANCES'      , False);
    SetControlChecked('SO_FRAISAVANCES', False)
  end;

  {Si on a pas les frais avances on ne peut pas avoir le palmarès transporteur}
  if ExisteNam('SO_FRAISAVANCES', FF) and not IsChecked('SO_FRAISAVANCES') then
  begin
    SetParamsoc('SO_PALMARESACTIF'      , False);
    SetControlChecked('SO_PALMARESACTIF', False)
  end;

  SetControlVisible('SO_FRAISGROUPES', (IsChecked('SO_FRAISAVANCES')) and TarifsAvances);
  SetControlVisible('SO_FRAISINDPR'  , (IsChecked('SO_FRAISAVANCES')) and TarifsAvances);
  SetControlVisible('LSO_FRAISINDPR' , (IsChecked('SO_FRAISAVANCES')) and TarifsAvances);

  if ExisteNam('SO_FRAISAVANCES', FF) and not IsChecked('SO_FRAISAVANCES') then
    SetParamsoc('SO_FRAISGROUPES', False);
  if (ExisteNam('SO_FRAISAVANCES', FF)) and (IsChecked('SO_FRAISAVANCES')) then
    SetParamsoc('SO_FRAISUNIQUE', True);

  if ExisteNam('SO_TARIFSENREGISTRES', FF) and not IsChecked('SO_TARIFSENREGISTRES') then
  begin
    SetControlVisible('LSO_TARIFSENREGFLUX', False);
    SetControlVisible('SO_TARIFSENREGFLUX' , False);

    SetControlText('SO_TARIFSENREGFLUX', ''); SetParamsoc('SO_TARIFSENREGFLUX', '');
  end;

  {Gestion du palmares transporteur}
  if ExisteNam('SO_PALMARESACTIF', FF) then
  begin
    lPalmares := IsChecked('SO_PALMARESACTIF');
    SetControlVisible('LSO_PALMARESENREG', False)    ; SetControlVisible('SO_PALMARESENREG' , False);
    SetControlVisible('LSO_PALMARESFORCE', lPalmares); SetControlVisible('SO_PALMARESFORCE' , lPalmares);
    SetControlVisible('LSO_PALMARESTIERS', (IsChecked('SO_FRAISAVANCES')) and TarifsAvances);
    SetControlVisible('SO_PALMARESTIERS' , (IsChecked('SO_FRAISAVANCES')) and TarifsAvances);

    if (not IsChecked('SO_PALMARESACTIF')) then
    begin
      SetControlChecked('SO_PALMARESENREG', False); SetParamsoc('SO_PALMARESENREG', False);
      SetControlChecked('SO_PALMARESFORCE', False); SetParamsoc('SO_PALMARESFORCE', False);
    end;
  end;

  { Stocks }
  GestionDesStocks;

  if (not GetParamSoc('SO_GCPIECEADRESSE')) then
  begin
    SetParamsoc('SO_GCTRANSPORTEURS', False);
    SetParamsoc('SO_FRAISAVANCES'   , False);
  end;

  { Gestion des marchés }
  if ExisteNam('SO_CDEMARCHE', FF) then
  begin
    { Activation Cde Marché implique que les tarifs avancées soient activés }
    if GetParamSocSecur('SO_TARIFSAVANCES', False) then
    begin
      SetControlEnabled('LSCO_MARCHE' , True);
      SetControlEnabled('SCO_MARCHE' , True);
      SetControlEnabled('SO_CDEMARCHE' , True);
      SetControlEnabled('SO_CDEMARCHEA' , IsChecked('SO_CDEMARCHE'));
      SetControlEnabled('SO_CDEMARCHEV' , IsChecked('SO_CDEMARCHE'));
      SetControlEnabled('SO_CDEMARCHEDEPOT' , IsChecked('SO_CDEMARCHE'));
      SetControlEnabled('SO_CDEMARCHESURMARCHEEPUISE' , IsChecked('SO_CDEMARCHE'));
      SetControlEnabled('SO_CDEMARCHEINIT', IsChecked('SO_CDEMARCHE'));
      SetControlEnabled('SO_CDEMARCHEDUP', IsChecked('SO_CDEMARCHE'));
      SetControlEnabled('LSCO_MARCHEACH', IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEA'));
      SetControlEnabled('SCO_MARCHEACH',  IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEA'));
      SetControlEnabled('SO_CDEMARCHEACH'   , IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEA'));
      SetControlEnabled('LSO_CDEMARCHEACH'  , IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEA'));
      SetControlEnabled('SO_CDEMARCHEAPPELACH' , IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEA'));
      SetControlEnabled('LSO_CDEMARCHEAPPELACH', IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEA'));
      SetControlEnabled('SO_CDEMARCHEAPPELDATEACH' , IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEA'));
      SetControlEnabled('LSO_CDEMARCHEAPPELDATEACH', IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEA'));
      SetControlEnabled('SO_CDEMARCHECTRLDVALACH' , IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEA'));
      SetControlEnabled('LSCO_MARCHEVTE', IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEV'));
      SetControlEnabled('SCO_MARCHEVTE',  IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEV'));
      SetControlEnabled('SO_CDEMARCHEVTE', IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEV'));
      SetControlEnabled('LSO_CDEMARCHEVTE', IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEV'));
      SetControlEnabled('SO_CDEMARCHEAPPELVTE' , IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEV'));
      SetControlEnabled('LSO_CDEMARCHEAPPELVTE', IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEV'));
      SetControlEnabled('SO_CDEMARCHEAPPELDATEVTE' , IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEV'));
      SetControlEnabled('LSO_CDEMARCHEAPPELDATEVTE', IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEV'));
      SetControlEnabled('SO_CDEMARCHECTRLDVALVTE' , IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEV'));

      {Dès que la gestion des commandes marché est activée il faut forcer la gestion des conditions spéciales sur les tarifs fournisseurs}
      if (IsChecked('SO_CDEMARCHEA')) and (not CdeMarcheAActif) then
      begin
        if PGIAsk('ATTENTION !!!' + #13 +
                  'Vous allez activer la gestion des commandes marchés achat. ' + #13 +
                  'La gestion des conditions spéciales dans le paramétrage '+ #13 +
                  'des tarifs fournisseurs est obligatoire.' + #13 +
                  'Confirmez-vous ce paramétrage ?',
                  'Gestion des commandes marchés') = mrYes then
        begin
          SetParamsoc('SO_CDEMARCHEA'				, True);
          SauvParamSoc:=False;
        end
        else
        begin
          SetParamsoc('SO_CDEMARCHEA'				, False);
          SetControlChecked('SO_CDEMARCHEA'	, False);
        end
      end
      else if not IsChecked('SO_CDEMARCHEA') then
      begin
        SetParamsoc('SO_CDEMARCHEA'				, False);
        SetControlChecked('SO_CDEMARCHEA'	, False);
      end;
      {Dès que la gestion des commandes marché est activée il faut forcer la gestion des conditions spéciales sur les tarifs fournisseurs}
      if (IsChecked('SO_CDEMARCHEV')) and (not CdeMarcheVActif) then
      begin
        if PGIAsk('ATTENTION !!!' + #13 +
                  'Vous allez activer la gestion des commandes marchés vente. ' + #13 +
                  'La gestion des conditions spéciales dans le paramétrage '+ #13 +
                  'des tarifs clients est obligatoire.' + #13 +
                  'Confirmez-vous ce paramétrage ?',
                  'Gestion des commandes marchés') = mrYes then
        begin
          SetParamsoc('SO_CDEMARCHEV'				, True);
			  	SauvParamSoc:=False;
        end
        else
        begin
			    SetParamsoc('SO_CDEMARCHEV'				, False);
			    SetControlChecked('SO_CDEMARCHEV'	, False);
        end;
      end
      else if not IsChecked('SO_CDEMARCHEV') then
      begin
        SetParamsoc('SO_CDEMARCHEV'				, False);
        SetControlChecked('SO_CDEMARCHEV'	, False);
      end;
    end
    else
    begin
      SetControlEnabled('LSCO_MARCHE' , False);
      SetControlEnabled('SCO_MARCHE' , False);
      SetControlEnabled('SO_CDEMARCHE' , False);
      SetControlEnabled('SO_CDEMARCHEA' , False);
      SetControlEnabled('SO_CDEMARCHEV' , False);
      SetControlEnabled('SO_CDEMARCHEDEPOT' , False);
      SetControlEnabled('SO_CDEMARCHESURMARCHEEPUISE' , False);
      SetControlEnabled('SO_CDEMARCHEINIT' , False);
      SetControlEnabled('SO_CDEMARCHEDUP' , False);
      SetControlEnabled('LSCO_MARCHEACH', False);
      SetControlEnabled('SCO_MARCHEACH',  False);
      SetControlEnabled('SO_CDEMARCHEACH'   , False);
      SetControlEnabled('LSO_CDEMARCHEACH'  , False);
      SetControlEnabled('SO_CDEMARCHEAPPELACH' , False);
      SetControlEnabled('LSO_CDEMARCHEAPPELACH', False);
      SetControlEnabled('SO_CDEMARCHEAPPELDATEACH' , False);
      SetControlEnabled('LSO_CDEMARCHEAPPELDATEACH', False);
      SetControlEnabled('SO_CDEMARCHECTRLDVALACH' , False);
      SetControlEnabled('LSCO_MARCHEVTE', False);
      SetControlEnabled('SCO_MARCHEVTE',  False);
      SetControlEnabled('SO_CDEMARCHEVTE', False);
      SetControlEnabled('LSO_CDEMARCHEVTE', False);
      SetControlEnabled('SO_CDEMARCHEAPPELVTE' , False);
      SetControlEnabled('LSO_CDEMARCHEAPPELVTE', False);
      SetControlEnabled('SO_CDEMARCHEAPPELDATEVTE' , False);
      SetControlEnabled('LSO_CDEMARCHEAPPELDATEVTE', False);
      SetControlEnabled('SO_CDEMARCHECTRLDVALVTE' , False);
    end;
  end;

  {$IFDEF GCGC}
    {$IFNDEF PGIMAJVER}
      if (not IsChecked('SO_FRAISAVANCES')) then
      begin
        SetControlChecked('SO_FRAISGROUPES' , False);
        SetControlText('SO_FRAISINDPR'      , 'NON');
      end
      else
        {Contrôle d'unicité des fras obligatoire en frais avancés}
        SetControlChecked('SO_FRAISUNIQUE', True);

      SetControlEnabled('SO_FRAISGROUPES', IsChecked('SO_FRAISAVANCES'));
      SetControlEnabled('SO_FRAISUNIQUE' , not IsChecked('SO_FRAISAVANCES'));
    {$ENDIF !PGIMAJVER}
  {$ENDIF GCGC}

  { Sous-ordres }
  if ExisteNam('SO_WSOUSORDRES', FF) then
  begin
    if not SauvParamSocSO then
    begin
      SousOrdresActifs := wGereSousOrdres();
      SauvParamSocSO   := True;
    end;
    GereSousOrdres := IsChecked('SO_WSOUSORDRES');
    SetControlEnabled('SCO_WSSOIMPACTEWOB'       , GereSousOrdres);
    SetControlEnabled('SO_WSSOIMPACTEQTEWOBONWOL', GereSousOrdres);
    if not IsChecked('SO_WSOUSORDRES') then
      SetControlChecked('SO_WSSOIMPACTEQTEWOBONWOL', False);
    SetControlEnabled('SCO_WSSOIMPACTEWOL'       , GereSousOrdres);
    SetControlEnabled('SO_WSSOIMPACTEQTEWOLONWOB', GereSousOrdres);
    if not IsChecked('SO_WSOUSORDRES') then
      SetControlChecked('SO_WSSOIMPACTEQTEWOLONWOB', False);
//GP_20080204_TS_GP14791 >>>
    { Valeur changée, si on décide de gérer les sous-ordres : Création d'un index }
    if GereSousOrdres and not SousOrdresActifs then
    begin
      SousOrdresActifs := True;
      wInitProgressForm(FF, TraduireMemoire('Gestion des sous-ordre activée'), '', 1, False, False);
      try
        wMoveProgressForm(TraduireMemoire('Création d''un index'));
        try
          ExecuteSQL('CREATE INDEX WOL_CLETMPSSO ON WORDRELIG (WOL_NATURETRAVAIL, WOL_ORDREPERE)');
        except
          PgiInfo('Cet index existe déjà');
          { L'index existe déjà, on capture l'exception pour l'ignorer.
            Pas de test possible sur l'existence d'un index }
        end
      finally
        wFiniProgressForm();
      end
    end
    { Valeur changée, si on décide de ne plus gérer les sous-ordres : Suppression de l'index }
    else if not GereSousOrdres and SousOrdresActifs
    and (PgiAsk('Voulez-vous supprimer l''index correspondant à cette fonctionnalité ?') = mrYes) then
    begin
      SousOrdresActifs := False;
      wInitProgressForm(FF, TraduireMemoire('Gestion des sous-ordre désactivée'), '', 1, False, False);
      try
        wMoveProgressForm(TraduireMemoire('Suppression de l''index'));
        try
          if IsOracle() then
            ExecuteSQL('DROP INDEX WOL_CLETMPSSO')
          else
            ExecuteSQL('DROP INDEX WORDRELIG.WOL_CLETMPSSO');
        except
          { L'index n'existe pas, on capture l'exception pour l'ignorer.
            Pas de test possible sur l'existence d'un index }
        end
      finally
        wFiniProgressForm();
      end
    end;
  end;
//GP_20080204_TS_GP14791 <<<

  { Préférences système tarifaire : Mode SAV }
  SetControlEnabled('SO_PREFSYSTTARIF', V_PGI.Sav);

  {$IFNDEF GPAO}
    {Sous-traitance d'assemblage}
    if (ExisteNam('SO_STAGC', FF)) and (not STAGCActif) and (IsChecked('SO_STAGC')) then
    begin
      if PGIAsk('Attention !' + #13 +
                ' Vous allez activer la sous-traitance d''achat.' + #13 +
                ' Cette action est irréversible.'+ #13 +
                ' Confirmez-vous ce paramétrage ?',
                'Module Assemblage') = MrYes then
      begin
        if InitTablesStAGC(ErrMsg)='' then
        begin
          SetParamsoc('SO_STAGC', true);
          SauvParamStAGC:=false;
        end
        else
        begin
          PGIError(TraduireMemoire('Problème d''initialisation de la sous-traitance.')
                  + #13 + #13
                  + ' ' + TraduireMemoire(ErrMsg));
          SetParamsoc('SO_STAGC'			, false);
          SetControlChecked('SO_STAGC', false);
        end;;
      end
      else
      begin
        SetParamsoc('SO_STAGC'			, false);
        SetControlChecked('SO_STAGC', false);
      end;
    end;
  {$ENDIF !GPAO}

  { CCMX-CEGID le 23/06/2005 DEBUT }
  SetControlVisible('SO_GCNATPIECEEXPE' , IsChecked('SO_GCTRANSPORTEURS'));
  SetControlVisible('LSO_GCNATPIECEEXPE', IsChecked('SO_GCTRANSPORTEURS'));
  { CCMX-CEGID le 23/06/2005 FIN }
  { CCMX DM - TR1 - DEBUT }
  SetControlVisible('SO_GCSOUCHEPIECEEXPE'  , IsChecked('SO_GCTRANSPORTEURS'));
  SetControlVisible('LSO_GCSOUCHEPIECEEXPE' , IsChecked('SO_GCTRANSPORTEURS'));
  SetControlVisible('SO_GCGEREINFOTRANS'    , IsChecked('SO_GCTRANSPORTEURS'));
  SetControlVisible('LSO_GCGEREINFOTRANS'   , IsChecked('SO_GCTRANSPORTEURS'));
  { CCMX DM - TR1 - FIN }
  { CCMX-CEGID le 30/01/2006 DEBUT }
  SetControlVisible('SO_GCSOUCHEDA'  , IsChecked('SO_GCGESTIONDA'));
  SetControlVisible('LSO_GCSOUCHEDA'  , IsChecked('SO_GCGESTIONDA'));
  //D: FQ15217
  SetControlVisible('SO_GCACCESSPUHTDA'  , IsChecked('SO_GCGESTIONDA'));
  //F: FQ15217
  { CCMX-CEGID le 30/01/2006 FIN }
  { DEBUT CCMX-CEGID TMDR Dev N° 4722} {Evolution}
  SetControlVisible('SO_ETIQUETTESTMDR'  , IsChecked('SO_GCTRANSPORTEURS') AND  IsChecked('SO_COLISAGEACTIF'));
  SetControlVisible('LSO_ETIQUETTESTMDR'  , IsChecked('SO_GCTRANSPORTEURS') AND IsChecked('SO_COLISAGEACTIF'));
  { FIN CCMX-CEGID le TMDR Dev N° 4722 } {Evolution}
  {$IFNDEF GPAO}
    if (ExisteNam('SO_CBNSANSDECALATTPRO', FF)) then
    begin
      TCheckBox(GetFromNam('SO_CBNSANSDECALATTPRO',FF)).Caption := TraduireMemoire('Débrayer le mode de décalage des attendus d''assemblage');
      TCheckBox(GetFromNam('SO_CBNSANSDECALATTPRO',FF)).Enabled := StAGCActif;
    end;
  {$ENDIF !GPAO}

  { Prix de revient }
  if ExisteNam('SO_PDRORRORIGINE', FF) then
  begin
    lVisible := (thValComboBox(GetFromNam('SO_PDRORRORIGINE', FF)).Value<>'B');
    SetControlVisible('SO_PDRORRQTE' , lVisible);
    SetControlVisible('LSO_PDRORRQTE', lVisible);
  end;

  if ExisteNam('SO_PDRORRORIGINECDE', FF) then
  begin
    lVisible := (thValComboBox(GetFromNam('SO_PDRORRORIGINECDE', FF)).Value<>'B');
    SetControlVisible('SO_PDRORRQTECDE' , lVisible);
    SetControlVisible('LSO_PDRORRQTECDE', lVisible);
  end;

  if ExisteNam('SO_SECTIONSPDR', FF) then
  begin
    // Pour l'instant on ne gère que les rubriques d'analyses
    SetControlVisible('SO_SECTIONSPDR' , False);
    SetControlVisible('SO_RUBRIQUESPDR', False);
    // Pas de création automatique des lignes budgétaires si gestion des sections d'analyses
    SetControlEnabled('SO_CREATLIGAUTO'   , not IsChecked('SO_SECTIONSPDR'));
    SetControlEnabled('SO_CONFIRMCREATLIG', not IsChecked('SO_SECTIONSPDR'));
    if IsChecked('SO_SECTIONSPDR') then
    begin
      SetControlChecked('SO_CREATLIGAUTO'   , false);
      SetControlChecked('SO_CONFIRMCREATLIG', false);
    end;
    if (not IsChecked('SO_SECTIONSPDR')) and (not IsChecked('SO_RUBRIQUESPDR')) then
      SetControlChecked('SO_RUBRIQUESPDR', True);
  end;

  {$IFDEF AFFAIRE}
    { Si Numérotaion automatique du code article -> Pas de chois possible pour la saisie de l'article affaire ("manuel" obligatoire)}
    if GetParamSocSecur('SO_GCNUMARTAUTO', False) then
      SetControlText('SO_AFSAISIEART', 'MAN');
    SetControlEnabled('SO_AFSAISIEART'        , not GetParamSocSecur('SO_GCNUMARTAUTO', False) );
    SetControlEnabled('LSO_AFSAISIEART'       , not GetParamSocSecur('SO_GCNUMARTAUTO', False) );
  {$ENDIF AFFAIRE}

  //GP_20080107_TS_GC15692 >>>
  SetEDIControlsVisible();
  {$IFDEF EDI}
    if ExisteNam('SO_EDITIERSPIECE', FF) then
      THEdit(GetFromNam('SO_EDITIERSPIECE', FF)).Plus := 'GP" AND (DH_NOMCHAMP="GP_TIERS" OR DH_NOMCHAMP="GP_TIERSLIVRE" OR '
                                                       +          'DH_NOMCHAMP="GP_TIERSFACTURE" OR DH_NOMCHAMP="GP_TIERSPAYEUR") '
                                                       +          'AND DH_TYPECHAMP="VARCHAR(17)';
                                                       { N.B. : pas de " ouvrante ni fermante car le plus original de la tablette
                                                       est de la forme : DH_PREFIXE="&#@" }
  {$ENDIF EDI}
  //GP_20080107_TS_GC15692 <<<
end;
{$ENDIF GPAOLIGHT}
{$ENDIF EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : XMG
Créé le ...... : 31/07/2003
Modifié le ... :   /  /
Description .. : Adapte la forme aux besoins de la version Espagnole
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
Procedure CacheFonctionsESP ( CC : TControl ; FF : TForm ) ;
Begin
if VH^.PaysLocalisation=COdeISOES then
   Begin
   SetInvi('SO_OUITVAENC',FF) ;
   SetInvi('SO_TVAENCAISSEMENT',FF) ;
   SetInvi('LSO_TVAENCAISSEMENT',FF) ;
   THLabel(GetFromNam('LSCO_GRTVAENCAIS',FF)).Caption:=TraduireMemoire('Paramètres TVA') ;
   End ;
End ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF SAV}
{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc SAUZET
Créé le ...... : 08/09/2004
Modifié le ... :   /  /
Description .. : Modification pour le SAV
Mots clefs ... :
*****************************************************************}
Procedure CacheFonctionsSAV(CC: TControl; FF: TForm);
begin
  if ExisteNam ('SO_QUALIFMVTWPC', FF) then
  begin
    ThMultiValComboBox(GetFromNam('SO_QUALIFMVTWPC', FF)).Complete := true;
    ThMultiValComboBox(GetFromNam('SO_QUALIFMVTWPC', FF)).Aucun := true;
    ThMultiValComboBox(GetFromNam('SO_QUALIFMVTWPC', FF)).plus := 'AND GPP_NATUREPIECEG NOT IN ("FFA","FFO","CCH","FPR","APR")';
  end;
  { mng pas d'intervention si pas GRC}
  if (Not (ctxGRC in V_PGI.PGIContexte)) then
    begin
    if ExisteNam ('SCO_WLIGIINFCOMPL', FF) then SetInvi('SCO_WLIGIINFCOMPL',FF);
    if ExisteNam ('LSCO_WLIGIINFCOMPL', FF) then SetInvi('LSCO_WLIGIINFCOMPL',FF);
    if ExisteNam ('SO_RTGESTINFOS007', FF) then SetInvi('SO_RTGESTINFOS007',FF);
    if ExisteNam ('SO_RTLARGEURFICHE007', FF) then SetInvi('SO_RTLARGEURFICHE007',FF);
    if ExisteNam ('SO_RTHAUTEURFICHE007', FF) then SetInvi('SO_RTHAUTEURFICHE007',FF);
    if ExisteNam ('LSO_RTLARGEURFICHE007', FF) then SetInvi('LSO_RTLARGEURFICHE007',FF);
    if ExisteNam ('LSO_RTHAUTEURFICHE007', FF) then SetInvi('LSO_RTHAUTEURFICHE007',FF);
    end;
  if (ExisteNam('SO_GESTSTATUTUNIQUE1', FF)
       or ExisteNam('SO_GESTSTATUTUNIQUE2', FF)) then
    GereZonesVersions (CC);
end;
{$ENDIF SAV}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF GPAO}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Karine BORGHETTI
Créé le ...... : 07/10/2004
Modifié le ... : 07/10/2004
Description .. : Modification pour la GPAO
Mots clefs ... :
*****************************************************************}
procedure CacheFonctionsNonGP(CC: TControl; FF: TForm);
begin
  if ExisteNam('SO_ARTICLECONFIG', FF) then SetInvi('SO_ARTICLECONFIG' ,FF);
  if ExisteNam('LSO_ARTICLECONFIG',FF) then SetInvi('LSO_ARTICLECONFIG',FF);
  {Méthode de valorisation au coût standard}
  if ExisteNam('LSCO_PDRGRMETHVALO', FF) then SetInvi('LSCO_PDRGRMETHVALO', FF);
  if ExisteNam('SCO_PDRGRMETHVALO' , FF) then SetInvi('SCO_PDRGRMETHVALO' , FF);
  //GP_20071123_CVN_GC15481 Déb
  if ExisteNam('LSCO_PDRMETHVALO'  , FF) then SetInvi('LSCO_PDRMETHVALO'  , FF);
  if ExisteNam('SCO_PDRMETHVALO'   , FF) then SetInvi('SCO_PDRMETHVALO'   , FF);
  //GP_20071123_CVN_GC15481 Fin
  if ExisteNam('SO_PDRMETHVALOSTD' , FF) then SetInvi('SO_PDRMETHVALOSTD' , FF);
  {Consolidation des entêtes de prix de revient : bases techniques et ordres}
  if ExisteNam('LSCO_PDRGRCONSOAVANCE', FF) then SetInvi('LSCO_PDRGRCONSOAVANCE', FF);
  if ExisteNam('SCO_PDRGRCONSOAVANCE' , FF) then SetInvi('SCO_PDRGRCONSOAVANCE' , FF);
  //GP_20071123_CVN_GC15481 Déb
  if ExisteNam('LSCO_PDRCONSOAVANCE'  , FF) then SetInvi('LSCO_PDRCONSOAVANCE'  , FF);
  if ExisteNam('SCO_PDRCONSOAVANCE'   , FF) then SetInvi('SCO_PDRCONSOAVANCE'   , FF);
  //GP_20071123_CVN_GC15481 Fin
  if ExisteNam('SO_PDRPARAMCONSOPDR'  , FF) then SetInvi('SO_PDRPARAMCONSOPDR'  , FF);
  if ExisteNam('SO_PDRCONSOWPLWPL'    , FF) then SetInvi('SO_PDRCONSOWPLWPL'    , FF);
  if ExisteNam('SO_PDRCONSOWPLWPE'    , FF) then SetInvi('SO_PDRCONSOWPLWPE'    , FF);
  {Consolidation des prix de revient d'ordre dans WPL}
  if ExisteNam('LSCO_PDRGRCONSOWPLORD', FF) then SetInvi('LSCO_PDRGRCONSOWPLORD', FF);
  if ExisteNam('SCO_PDRGRCONSOWPLORD' , FF) then SetInvi('SCO_PDRGRCONSOWPLORD' , FF);
  if ExisteNam('SO_PDRCONSOORDWPL'    , FF) then SetInvi('SO_PDRCONSOORDWPL'    , FF);
  {Détail des prix de revient : base technique}
  if ExisteNam('LSCO_PDRGRPDRBTH', FF) then SetInvi('LSCO_PDRGRPDRBTH', FF);
  if ExisteNam('SCO_PDRGRPDRBTH' , FF) then SetInvi('SCO_PDRGRPDRBTH' , FF);
  if ExisteNam('SO_PDRENREGWPL'  , FF) then SetInvi('SO_PDRENREGWPL'  , FF);
  //GP_20071017_CVN_GC15481 Déb
  if ExisteNam('SCO_PDRGRENREGWPL'  , FF) then SetInvi('SCO_PDRGRENREGWPL'  , FF);
  if ExisteNam('LSCO_PDRGRENREGWPL' , FF) then SetInvi('LSCO_PDRGRENREGWPL' , FF);
  if ExisteNam('SCO_PDRGENEREWPL'   , FF) then SetInvi('SCO_PDRGENEREWPL'   , FF);
  if ExisteNam('LSCO_PDRGENEREWPL'  , FF) then SetInvi('LSCO_PDRGENEREWPL'  , FF);
  if ExisteNam('LSO_PDRCALCULSS' , FF) then SetInvi('LSO_PDRCALCULSS' , FF);
  //GP_20071017_CVN_GC15481 Fin
  if ExisteNam('SO_PDRCALCULSS'  , FF) then SetInvi('SO_PDRCALCULSS'  , FF);
  {Configuration par défaut de visualisation des prix de revient depuis la fiche article}
  if ExisteNam('LSCO_PDRGRDEFAUTGA', FF) then SetInvi('LSCO_PDRGRDEFAUTGA', FF);
  if ExisteNam('SCO_PDRGRDEFAUTGA' , FF) then SetInvi('SCO_PDRGRDEFAUTGA' , FF);
  //GP_20071123_CVN_GC15481 Déb
  if ExisteNam('LSCO_PDRDEFAUTGA'  , FF) then SetInvi('LSCO_PDRDEFAUTGA'  , FF);
  if ExisteNam('SCO_PDRDEFAUTGA'   , FF) then SetInvi('SCO_PDRDEFAUTGA'   , FF);
  //GP_20071123_CVN_GC15481 Fin
  if ExisteNam('LSO_PDRNATURETRAGA', FF) then SetInvi('LSO_PDRNATURETRAGA', FF);
  if ExisteNam('SO_PDRNATURETRAGA' , FF) then SetInvi('SO_PDRNATURETRAGA' , FF);
  if ExisteNam('LSO_PDRNATUREPDRGA', FF) then SetInvi('LSO_PDRNATUREPDRGA', FF);
  if ExisteNam('SO_PDRNATUREPDRGA' , FF) then SetInvi('SO_PDRNATUREPDRGA' , FF);
  if ExisteNam('LSO_PDRTYPEPDRGA'  , FF) then SetInvi('LSO_PDRTYPEPDRGA'  , FF);
  if ExisteNam('SO_PDRTYPEPDRGA'   , FF) then SetInvi('SO_PDRTYPEPDRGA'   , FF);
  if ExisteNam('LSO_PDRNBORDREGA'  , FF) then SetInvi('LSO_PDRNBORDREGA'  , FF);
  if ExisteNam('SO_PDRNBORDREGA'   , FF) then SetInvi('SO_PDRNBORDREGA'   , FF);
  if ExisteNam('SCO_PDRNBORDREGA'  , FF) then SetInvi('SCO_PDRNBORDREGA'  , FF);
  //GP_20071123_CVN_GC15481
  if ExisteNam('SCO_PDRNBORDREGA2' , FF) then SetInvi('SCO_PDRNBORDREGA2'  , FF);
end;
{$ENDIF GPAO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF CRM}
Procedure CacheFonctionsCRM(CC: TControl; FF: TForm);
var i : integer;
begin
  if ExisteNam ('LSCO_AFGRVALORES', FF) then
  begin
  SetInvi('SCO_AFGRVALORES',FF); SetInvi('LSCO_AFGRVALORES',FF);
  SetInvi('SO_AFFRAISGEN1',FF); SetInvi('LSO_AFFRAISGEN1',FF);
  SetInvi('SO_AFFRAISGEN2',FF); SetInvi('LSO_AFFRAISGEN2',FF);
  SetInvi('SO_AFCHARGESPAT',FF); SetInvi('LSO_AFCHARGESPAT',FF);
  SetInvi('SO_AFCOEFMETIER',FF); SetInvi('LSO_AFCOEFMETIER',FF);
  SetInvi('SO_AFRESCALCULPR',FF);
  SetInvi('SO_AFFNUMAUXI',FF); SetInvi('LSO_AFFNUMAUXI',FF);
  SetInvi('SO_AFPRESTATIONRES',FF); SetInvi('LSO_AFPRESTATIONRES',FF);
  SetInvi('SO_AFFRACINEAUXI',FF); SetInvi('LSO_AFFRACINEAUXI',FF);
  SetInvi('SO_AFSTANDCALEN',FF); SetInvi('LSO_AFSTANDCALEN',FF);
  SetInvi('SCO_GRLIENPAIE',FF); SetInvi('LSCO_GRLIENPAIE',FF);
  SetInvi('SO_PGLIENRESSOURCE_',FF);
  SetInvi('SO_AFLIENPAIEVAR',FF);
  SetInvi('SO_AFLIENPAIEANA',FF);
  SetInvi('SO_AFLIENPAIEDEC',FF);
  SetInvi('SO_AFLIENPAIEFIC',FF); SetInvi('LSO_AFLIENPAIEFIC',FF);
  end;
  if ExisteNam ('LSCO_QBPGRP20', FF) then
  begin
  SetInvi('SCO_QBPGROUPE1',FF);
  SetInvi('LSCO_QBPGROUPE1',FF);
  for i := 1 to 7 do
    begin
    SetInvi('LSO_QBPVALAFF'+IntToStr(i),FF);
    SetInvi('SO_QBPVALAFF'+IntToStr(i),FF);
    SetInvi('SO_QBPLIBVALAFF'+IntToStr(i),FF);
    end;
  SetInvi('SO_QBPLIENSCM',FF);
  SetInvi('SO_QBPEXPORTSESVALIDE',FF);
  end;
end;
{$ENDIF CRM}
{$ENDIF EAGLSERVER}

{ Crypte le text }
procedure CryptePassword(var Valeur : String);
var s: String[255];
    c: array[0..255] of Byte absolute s;
    i: integer;
begin

 s := Valeur;
 for i := 1 to ord(s[0]) do c[i] := 23 xor c[i];
 Valeur := s;

end;

{ Décrypte le text }
procedure DecryptePassword(var Valeur : String);
var c: array[0..255] of Byte absolute Valeur;
    i: integer;
begin

 for i:=1 to Length(Valeur) do Valeur[i] := Char(23 xor ord(Valeur[i]));

end;

{$IFNDEF EAGLSERVER}
Function ChargePageSoc ( CC : TControl ) : boolean ;
Var FF : TForm ;
    Password : String;
BEGIN

CacheFonctionS3S5(CC) ;
GRCMasqueOptionsProjet(CC);
ShowParamSoc(CC) ;
BlocageDevise(CC) ;
BlocageECC(CC) ;
InitDivers(CC) ;
CacheFonctionBTP(CC) ;
MasquePourGescom(CC) ;
ChargeFourchettes(CC,True) ;
FF:=GetLaForm(CC) ;
{$IFDEF GPAOLIGHT}
	InitParamSocGPAO(FF);
{$ENDIF GPAOLIGHT}
if ExisteNam('SO_DEVISEPRINC',FF)   // ajout JCF
   then DevisePrinc:=THValComboBox(GetFromNam('SO_DEVISEPRINC',FF)).value
   else DevisePrinc:='';
CacheFonctionsESP(CC,FF) ;
{$ifdef AFFAIRE}
CacheFctGIGA(CC);
{$endif}
{$IFDEF SAV}
CacheFonctionsSAV(CC, FF);
{$ENDIF SAV}
GereParamSocPieceOrigine(FF);
{$IFNDEF GPAO}
CacheFonctionsNonGP(CC, FF);
{$ENDIF GPAO}
{$IFDEF TRESO}
GereParamTreso(FF); {JP 04/11/04}
{$ENDIF}
{$IFDEF COMPTA}
GereCompta(FF); //SG6 07/12/2004
{$ENDIF}
GereAnaParamAff(CC);
{$IFDEF GESCOM}
  CdeOuvCacheFonctions(FF);
  // DEBUT CCMX-CEGID FQ N° GC14064 Prestation CAB visible seulement pour GESCOM
  If ExisteNam('LSCO_GRCAB',FF) Then SetVisi('LSCO_GRCAB', FF);
  If ExisteNam('SCO_GRCAB',FF) Then SetVisi('SCO_GRCAB', FF);
  If ExisteNam('SO_GCPRCABAUTO',FF) Then SetVisi('SO_GCPRCABAUTO', FF);
  If ExisteNam('SO_GCPRCABPRESTATION',FF) Then SetVisi('SO_GCPRCABPRESTATION', FF);
{$ELSE}
  If ExisteNam('LSCO_GRCAB',FF) Then SetInVi('LSCO_GRCAB', FF);
  If ExisteNam('SCO_GRCAB',FF) Then SetInVi('SCO_GRCAB', FF);
  If ExisteNam('SO_GCPRCABAUTO',FF) Then SetInVi('SO_GCPRCABAUTO', FF);
  If ExisteNam('SO_GCPRCABPRESTATION',FF) Then SetInVi('SO_GCPRCABPRESTATION', FF);
  // FIN CCMX-CEGID FQ N° GC14064
{$ENDIF GESCOM}

  if ExisteNam('SO_MAJAFFECTATION', FF) then
  begin
    if THCheckBox(GetFromNam ('SO_MAJAFFECTATION',FF)).checked then
    begin
      SetVisi('LSO_DUREEMINI', FF);
      SetVisi('SO_DUREEMINI', FF);
    end else
    begin
      SetInvi('LSO_DUREEMINI', FF);
      SetInvi('SO_DUREEMINI', FF);
    end;
  end;       

GereTypeTVAIntraComm(CC);
GereContremarque(CC);
GereFarFae(CC);
GereElimineLigne0(CC);
{$IFDEF GCGC}
GereDevisePiece(CC);
SetPlusUniteParDefaut(CC);
SetGCGestUniteModeState(CC);
SetGereEnseigne(CC);        // gere les enseignes
//GP_BUG810_TP WPLANLIVR_20071008
SetWPlanLivr(CC);           { Cache le paramétrage par utilisateur }
{$ENDIF}
{$IFDEF NETEXPERT}
InitComptaRevision (CC) ;
{$ENDIF}
{$IFDEF CRM}
CacheFonctionsCRM(CC, FF);
{$ENDIF CRM}
if ExisteNam('SO_TARIFSENREGFLUX',FF) then
begin
  ThMultiValComboBox(GetFromNam('SO_TARIFSENREGFLUX', FF)).Complete := True;
  ThMultiValComboBox(GetFromNam('SO_TARIFSENREGFLUX', FF)).Aucun    := True;
end;
if ExisteNam('SO_TARIFSRECHTRANSFO',FF) then
begin
  ThMultiValComboBox(GetFromNam('SO_TARIFSRECHTRANSFO', FF)).Complete := True;
  ThMultiValComboBox(GetFromNam('SO_TARIFSRECHTRANSFO', FF)).Aucun    := True;
end;
if ExisteNam('SO_CDEMARCHEAPPELACH',FF) then
begin
  ThMultiValComboBox(GetFromNam('SO_CDEMARCHEAPPELACH', FF)).Complete := True;
  ThMultiValComboBox(GetFromNam('SO_CDEMARCHEAPPELACH', FF)).Aucun    := False;
end;
if ExisteNam('SO_MESTYPEOPER',FF) then
  ThMultiValComboBox(GetFromNam('SO_MESTYPEOPER', FF)).Complete := True;
if ExisteNam('SO_MESSUIVIOP',FF) then
  ThMultiValComboBox(GetFromNam('SO_MESSUIVIOP', FF)).Complete := True;
if ExisteNam('SO_CDEMARCHEAPPELVTE',FF) then
begin
  ThMultiValComboBox(GetFromNam('SO_CDEMARCHEAPPELVTE', FF)).Complete := True;
  ThMultiValComboBox(GetFromNam('SO_CDEMARCHEAPPELVTE', FF)).Aucun    := False;
end;
if ExisteNam('SO_PDRTYPEPDRGA',FF) then
begin
  THValComboBox(GetFromNam('SO_PDRTYPEPDRGA', FF)).Plus := 'WRT_NATUREPDR="'+THValComboBox(GetFromNam('SO_PDRNATUREPDRGA', FF)).Value+'"';
end;
{$IFDEF STK}
{ Indicteurs de stock }
if ExisteNam('SO_STKPERIODETAUXR',FF) then
begin
  RefL:=THValComboBox(GetFromNam('SO_STKPERIODETAUXR',FF)) ;
  RefL.Items.Insert(0, TraduireMemoire('<<Aucun>>'));
  RefL.Values.Insert(0, '');
end;
{$ENDIF STK}
{$IFDEF GCGC}
If (ExisteNam('SO_MODULEPGISIDE',FF)) and (not VH_GC.ECSeria) then
  SetInvi('SO_MODULEPGISIDE',FF);
{$ENDIF GCGC}

  //C.B 15/05/2007
  planningIntegre(FF);
if ExisteNam('SO_SOLDEPIECEPREC',FF)  then
    GereSoldePiecePrec(TControl(GetFromNam('SO_SOLDEPIECEPREC',FF)));
// GC_20071016_GM_GC15422_DEBUT
if ExisteNam('SO_NATURECTRLCLOTURE',FF) then
begin
  ThMultiValComboBox(GetFromNam('SO_NATURECTRLCLOTURE', FF)).Complete := false;
  ThMultiValComboBox(GetFromNam('SO_NATURECTRLCLOTURE', FF)).Aucun    := True
end;
// GC_20071016_GM_GC15422_FIN
{ GC_JTR_GC15601_Début }
GereMdpMargeMini(CC);
{ GC_JTR_GC15601_Fin }
Result:=True ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
Q:=OpenSql('Select G_GENERAL, G_VENTILABLE, G_COLLECTIF, G_LETTRABLE, G_POINTABLE FROM GENERAUX WHERE G_GENERAL="'+GeneAtt+'"',True) ;
if Q.Fields[0].AsString='' then
   BEGIN
   Ferme(Q) ; Exit ;
   END else
   BEGIN
   OkOk:=True ;
   if (Q.Fields[1].AsString='X') And (OkOk) then BEGIN OkOk:=False ; HShowMessage('5;Société;Le compte général d''attente ne peut pas être ventilable.;W;O;O;O;','','') ; END ;
   if (Q.Fields[2].AsString='X') And (OkOk) then BEGIN OkOk:=False ; HShowMessage('6;Société;Le compte général d''attente ne peut pas être collectif.;W;O;O;O;','','') ; END ;
   if (Q.Fields[3].AsString='X') And (OkOk) then BEGIN OkOk:=False ; HShowMessage('7;Société;Le compte général d''attente ne peut pas être lettrable.;W;O;O;O;','','') ; END ;
   if (Q.Fields[4].AsString='X') And (OkOk) then BEGIN OkOk:=False ; HShowMessage('8;Société;Le compte général d''attente ne peut pas être pointable.;W;O;O;O;','','') ; END ;
   END ;
Ferme(Q) ;
if Not OkOk then BEGIN if CptAtt.CanFocus then CptAtt.SetFocus ; Result:=False ; Exit ; END ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
   Sql:='Select G_GENERAL FROM GENERAUX Where G_GENERAL="'+Cpt+'" '+LeWhere+'' ;
   Q:=OpenSql(Sql,True) ;
   if Not Q.EOF then Cpt:=Q.Fields[0].AsString else Cpt:='' ;
   Ferme(Q) ;
   if Cpt='' then
      BEGIN
      if LaZone.CanFocus then LaZone.SetFocus ;
      HShowMessage('16;Société;Le compte que vous avez choisi n''est pas en accord avec la nature attendue.;W;O;O;O;','','') ;
      Result:=False ; Break ;
      END ;
   END ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
if StrToDate(DateDeb.Text)<Encodedate(1999,01,01) then BEGIN HShowMessage('20;Société;La date d''entrée en vigueur de l''euro doit être comprise entre le 1er janvier 1999 et le 31 décembre 1999.;W;O;O;O;','','') ; Result:=False ; END ;
if StrToDate(DateDeb.Text)>Encodedate(1999,12,31) then
   BEGIN
   if ((DevPrinc.Value<>'') and (V_PGI.DevisePivot<>'') and (Not Devprinc.Enabled))
      then else BEGIN HShowMessage('20;Société;La date d''entrée en vigueur de l''euro doit être comprise entre le 1er janvier 1999 et le 31 décembre 1999.;W;O;O;O;','','') ; Result:=False ; END ;
   END ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
   HShowMessage('23;Société;Les comptes d''écart de conversion ne sont pas correctement renseignés.;W;O;O;O;','','') ;
   Result:=False ; Exit ;
   END ;
if ((CptC<>'') and (Not RechExisteH(CHCredit))) then
   BEGIN
   HShowMessage('23;Société;Les comptes d''écart de conversion ne sont pas correctement renseignés.;W;O;O;O;','','') ;
   Result:=False ; Exit ;
   END ;
if CptD<>'' then
   BEGIN
   OkVent:=False ; Okok:=True ;
   QQ:=OpenSQL('Select G_VENTILABLE from GENERAUX Where G_GENERAL="'+CptD+'"',True) ;
   if QQ.EOF then Okok:=False else OkVent:=(QQ.Fields[0].AsString='X') ;
   Ferme(QQ) ;
   if Not Okok then
      BEGIN
      HShowMessage('23;Société;Les comptes d''écart de conversion ne sont pas correctement renseignés.;W;O;O;O;','','') ;
      Result:=False ; Exit ;
      END else if OkVent then
      BEGIN
      HShowMessage('24;Société;Les comptes d''écart de conversion ne doivent pas être ventilables.;W;O;O;O;','','') ;
      Result:=False ; Exit ;
      END ;
   END ;
if CptC<>'' then
   BEGIN
   OkVent:=False ; Okok:=True ;
   QQ:=OpenSQL('Select G_VENTILABLE from GENERAUX Where G_GENERAL="'+CptC+'"',True) ;
   if QQ.EOF then Okok:=False else OkVent:=(QQ.Fields[0].AsString='X') ;
   Ferme(QQ) ;
   if Not Okok then
      BEGIN
      HShowMessage('23;Société;Les comptes d''écart de conversion ne sont pas correctement renseignés.;W;O;O;O;','','') ;
      Result:=False ; Exit ;
      END else if Okvent then
      BEGIN
      HShowMessage('24;Société;Les comptes d''écart de conversion ne doivent pas être ventilables.;W;O;O;O;','','') ;
      Result:=False ; Exit ;
      END ;
   END ;
Result:=Okok ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
function ExisteCptesAttente ( CC : TControl ) : boolean ;
Var FF : TForm ;
BEGIN
// Test de la page
Result:=True ; FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_GENATTEND',FF) then Exit ;
// Traitement
Result:=False ;
if Not RechExisteH(GetFromNam('SO_GENATTEND',FF)) then BEGIN HShowmessage('13;Société;Vous devez renseigner un compte général d''attente.;W;O;O;O;','','') ; Exit ; END ;
if Not RechExisteH(GetFromNam('SO_CLIATTEND',FF)) then BEGIN HShowmessage('14;Société;Vous devez renseigner un compte auxiliaire client d''attente.;W;O;O;O;','','') ; Exit ; END ;
if Not RechExisteH(GetFromNam('SO_FOUATTEND',FF)) then BEGIN HShowmessage('15;Société;Vous devez renseigner un compte auxiliaire fournisseur d''attente.;W;O;O;O;','','') ; Exit ; END ;
Result:=True ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
function ExisteCptesHT( CC : TControl ) : boolean ;
Var FF : TForm ;
begin
  Result := True;
  FF := GetLaForm(CC);
  if Not ExisteNam('SO_GCCPTEHTACH',FF) then Exit;
  Result := False;
  if Not RechExisteH(GetFromNam('SO_GCCPTEHTACH',FF)) then
  begin
    HShowmessage('13;Société;Vous devez renseigner un compte HT d''achat par défaut.;W;O;O;O;','','');
    Exit;
  end;
  if Not RechExisteH(GetFromNam('SO_GCCPTEHTVTE',FF)) then
  begin
    HShowmessage('13;Société;Vous devez renseigner un compte HT de vente par défaut.;W;O;O;O;','','');
    Exit;
  end;
  Result := True;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
function ExisteTVAIntracomm(CC : TControl) : boolean;
var FF : TForm;
    LeRegime, Msg : string;
    Qte : integer;
    Qry : TQuery;
begin
  Result := true;
  FF := GetLaForm(CC);
  if not ExisteNam('SO_GERETVAINTRACOMM', FF) then exit;
  if not TCheckBox(GetFromNam('SO_GERETVAINTRACOMM',FF)).Checked then exit;
  Msg := 'TVA intracommunautaire'+#13;
  if THValComboBox(GetFromNam('SO_TYPETVAINTRACOMM', FF)).Value = '' then
  begin
    HShowmessage('13;Société;Vous devez renseigner un régime de TVA intracommunautaire.;W;O;O;O;','','');
    Result := false;
  end else
  if THMultiValComboBox(GetFromNam('SO_PCETVAINTRACOMM', FF)).Text = '' then
  begin
    HShowmessage('13;Société;'+Msg+'Vous devez renseigner des natures de pièces.;W;O;O;O;','','');
    Result := false;
  end else
  if THValComboBox(GetFromNam('SO_TYPETVAINTRACOMM', FF)).Value = GetParamSoc('SO_REGIMEDEFAUT') then
  begin
    HShowmessage('13;Société;'+Msg+'Le régime de TVA intracommunautaire doit être différent du régime de TVA par défaut.;W;O;O;O;','','');
    Result := false;
  end else
  if GetParamSoc('SO_GCFAMILLETAXE1') = '' then
  begin
    HShowmessage('13;Société;'+Msg+'La "Famille de taxe" par défaut des articles doit être renseignée.;W;O;O;O;','','');
    Result := false;
  end;
  // Test s'il reste des codes ou taux non renseignés pour le régime
  if Result then
  begin
    LeRegime := THValComboBox(GetFromNam('SO_TYPETVAINTRACOMM', FF)).Value;
    Qry := OpenSQL('SELECT COUNT(TV_REGIME) AS QTE FROM TXCPTTVA WHERE TV_REGIME="' + LeRegime
                  +'" AND TV_TVAOUTPF = "TX1" AND ((TV_CPTEACH = "") OR (TV_CPTEVTE = ""))' , true);
    if not Qry.Eof then
      Qte := Qry.FindField('QTE').AsInteger
      else
      Qte := 0;
    Ferme(Qry);
    if Qte > 0 then
      HShowmessage('13;Société;'+Msg+'ATTENTION. Certains taux ou comptes ne sont pas renseignés pour le régime "'+ LeRegime +'".;W;O;O;O;','','');
  end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Function verifInfosBTP  ( CC : TControl ) : Boolean ;

	function JournalAnaOk (Code : string) : Boolean;
  begin
		result := ExisteSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_JOURNAL="'+Code+'" AND J_NATUREJAL="ODA"');
  end;

	function CodeSectionAnaOk(Code : string) : boolean;
  begin
		result := ExisteSQL('SELECT S_SECTION FROM SECTION WHERE S_SECTION="'+Code+'"');
  end;

  function ControleCptaPaieDirect (CC : TControl) : Boolean;
  Var FF : TForm ;
      Ena : Boolean;
      CodePaiementDir : string;
  begin
    Result := true;
  	FF:=TForm(CC.Owner);
    if not ExisteNam('SO_BTCPTAPAIEDIRECT',FF) then exit;

    Ena := THCheckBox(GetFromNam ('SO_BTCPTAPAIEDIRECT', FF)).checked;
    if not Ena then Exit;
    CodePaiementDir := THValComboBox(GetFromNam('SO_BTPAIESTDIRECTE', FF)).value;
    if CodePaiementDir = '' then
    begin
      HShowmessage('13;Société;Veuillez renseigner le mode de paiement associé au paiement directe des sous traitants.;W;O;O;O;','','');
      Result := false;
      Exit;
    end;
		if not ExisteSQL('SELECT MP_MODEPAIE FROM MODEPAIE WHERE MP_MODEPAIE="'+CodePaiementDir+'"') then
    begin
      HShowmessage('13;Société;Veuillez renseigner un mode de paiement existant.;W;O;O;O;','','');
      Result := false;
      Exit;
    end;

  end;

Var FF : TForm ;
    Ena : Boolean;
begin
  result := true;
  FF:=TForm(CC.Owner);

  if not ControleCptaPaieDirect (CC) then
  begin
    Result := false;
    Exit;
  end;

	if not ExisteNam('SO_OPTANALSTOCK',FF) then exit;

  Ena := THCheckBox(GetFromNam ('SO_OPTANALSTOCK', FF)).checked;

  if not ena then exit;

  if THValCombobox(GetFromNam('SO_BTAXEANALSTOCK', FF)).Value = '' then
  begin
    HShowmessage('13;Société;Veuillez renseigner l''axe de ventilation;W;O;O;O;','','');
    Result := false;
    Exit;
  end;
  (*
  if THCritMaskEdit(GetFromNam('SO_BTCPTANALSTOCK', FF)).Text = '' then
  begin
    HShowmessage('13;Société;Veuillez renseigner la section analytique de stock.;W;O;O;O;','','');
    Result := false;
    Exit;
  end;
  *)
  //
  if THCritMaskEdit(GetFromNam('SO_BTJNLANALSTOCK', FF)).Text = '' then
  begin
    HShowmessage('13;Société;Veuillez renseigner le journal de ventilation analytique de stock.;W;O;O;O;','','');
    Result := false;
    Exit;
  end;

  if THValComboBox(GetFromNam('SO_BTCPTANALSTOCK', FF)).Text <> '' then
  begin
    if not CodeSectionAnaOk(THCritMaskEdit(GetFromNam('SO_BTCPTANALSTOCK', FF)).Text) then
    begin
      HShowmessage('13;Société;La section analytique de stock est inconnu.;W;O;O;O;','','');
      Result := false;
      Exit;
    end;
  end;

  if not JournalAnaOk(THValComboBox(GetFromNam('SO_BTJNLANALSTOCK', FF)).Value) then
  begin
    HShowmessage('13;Société; Le journal d''enregistrement de suivi analytique inconnu.;W;O;O;O;','','');
    Result := false;
    Exit;
  end;

  //FV1 : 11/06/2014 - FS#978 - SCETEC : en export EXCEL d'analyse, passer la feuille générée en protection.


end;

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

  { GC_DBR_CRM10764_DEB }
  if not ControleFormatCodeTiers (nil, FF) then
  begin
    SS.SetFocus;
    exit;
  end;
  { GC_DBR_CRM10764_FIN }

Result:=True ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
function VerifPariteEuro ( CC : TControl ) : boolean ;
Var StParite, CH : String ;
    Parite, i : Integer ;
    FF  : TForm ;
    CHEuro : THNumEdit ;
BEGIN
// Test de la page
Result:=True ; FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_TAUXEURO',FF) then Exit ;
// Traitement
Result:=True ;
CHEuro:=THNumEdit(GetFromNam('SO_TAUXEURO',FF)) ;
CH:=Trim(CHEuro.Text) ;
For i:=Length(CH) downto 6 do
    BEGIN
    if CH[i] <> '0' then Break else
       BEGIN
       if CH[i] = '0' then CH[i] := ' ';
       END;
    END;
StParite:=Trim(CH);
if Pos(V_PGI.SepDecimal,StParite)<>0 then Delete(StParite,Pos(V_PGI.SepDecimal,StParite),1) ;
if StParite<>'' then Parite:=StrToInt(StParite) else Parite:=0 ;
if Parite<=0 then BEGIN HShowMessage('22;Société;La parité fixe par rapport à l''Euro doit être positive et de 6 chiffres significatifs maximum.;W;O;O;O;','','') ; ; Result:=False ; Exit ; END ;
StParite:=IntToStr(Parite) ;
if Length(StParite)>6 then BEGIN HShowMessage('22;Société;La parité fixe par rapport à l''Euro doit être positive et de 6 chiffres significatifs maximum.;W;O;O;O;','','') ; Result:=False ; END ;
if Not Result then if CHEuro.CanFocus then CHEuro.SetFocus ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
   HShowMessage('21;Société;Vous devez renseigner un caractère de bourrage.;W;O;O;O;','','') ;
   CB.SetFocus ; Exit ;
   END ;
CB:=TEdit(GetFromNam('SO_BOURREAUX',FF)) ;
if (CB.Text='') or (CB.Text=' ') then Okok:=False else
   BEGIN
   CH:=CB.Text[1] ; if CH in CCH then Okok:=False ;
   END ;
if Not Okok then
   BEGIN
   HShowMessage('21;Société;Vous devez renseigner un caractère de bourrage.;W;O;O;O;','','') ;
   CB.SetFocus ; Exit ;
   END ;

  { GC_DBR_CRM10764_DEB }
  if not ControleFormatCodeTiers (nil, FF) then
  begin
    CB.SetFocus;
    exit;
  end;
  { GC_DBR_CRM10764_FIN }

Result:=True ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Function EtabDefautOk ( CC : TControl ) : Boolean ;
Var CB : THvalComboBox ;
    FF  : TForm ;
BEGIN
// Test de la page
Result:=True ; FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_ETABLISDEFAUT',FF) then Exit ;
if (ctxPCL in V_PGI.PGIContexte) and  (ctxStandard in V_PGI.PGIContexte) then exit;
// Traitement
CB:=THValComboBox(GetFromNam('SO_ETABLISDEFAUT',FF)) ;
if CB.Value='' then
   BEGIN
   Result:=FALSE ; HShowMessage('17;Société;Vous devez renseigner un établissement par défaut.;W;O;O;O;','','') ;
   CB.SetFocus ;
   END ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Function VerifJalTP ( CC : TControl ) : boolean ;
Var FF  : TForm ;
    CBA : THValComboBox ;
//    Q   : TQuery ;
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
//   Q:=OpenSQL('SELECT J_MULTIDEVISE FROM JOURNAL WHERE J_JOURNAL="'+CBA.Value+'" AND J_MULTIDEVISE="X"',True) ;
   Trouv:= ExisteSQL('SELECT J_MULTIDEVISE FROM JOURNAL WHERE J_JOURNAL="'+CBA.Value+'" AND J_MULTIDEVISE="X"'); //Not Q.EOF ;
//   Ferme(Q) ;
   if Not Trouv then BEGIN HShowMessage('3;Société;Les journaux pour tiers payeurs doivent être multi-devise.;W;O;O;O;','','') ; CBA.SetFocus ; Exit ; END ;
   END ;
CBA:=THValComboBox(GetFromNam('SO_JALVTP',FF)) ;
if CBA.Value<>'' then
   BEGIN
//   Q:=OpenSQL('SELECT J_MULTIDEVISE FROM JOURNAL WHERE J_JOURNAL="'+CBA.Value+'" AND J_MULTIDEVISE="X"',True) ;
   Trouv:= ExisteSQL('SELECT J_MULTIDEVISE FROM JOURNAL WHERE J_JOURNAL="'+CBA.Value+'" AND J_MULTIDEVISE="X"'); // Not Q.EOF ;
//   Ferme(Q) ;
   if Not Trouv then BEGIN HShowMessage('3;Société;Les journaux pour tiers payeurs doivent être multi-devise.;W;O;O;O;','','') ; CBA.SetFocus ; Exit ; END ;
   END ;
Result:=True ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
if CH.Text='' then
BEGIN
  HShowMessage('1;Société;Vous devez renseigner un libellé.;W;O;O;O;','','') ;
  if CH.CanFocus then CH.SetFocus ;
  Exit ;
END ;
CH:=TEdit(GetFromNam('SO_ADRESSE1',FF)) ;
if CH.Text='' then
BEGIN
  HShowMessage('3;Société;Vous devez renseigner une adresse.;W;O;O;O;','','') ;
  if CH.CanFocus then CH.SetFocus ;
  Exit ;
END ;
//XVI Si localisation Espagnole et la Société est Espagnole,on vérifie le SIRET
if (VH^.PaysLocalisation=CodeISOES) and  (CodeISOduPays(THValComboBox(GetFromNam('SO_PAYS',FF)).Value)=CodeISOES) then Begin
   CH:=TEdit(GetFromNam('SO_SIRET',FF)) ;
   if CH.Text <> VerifNIF_ES(CH.Text) then BEGIN
      HShowMessage('3;Société;Le NIF/CIF saisie est incorrecte.;W;O;O;O;','','') ;
      if CH.CanFocus then CH.SetFocus ;
      Exit ;
   End ;
End ; //XVI 24/02/2005
Result:=True ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
HShowMessage('9;Société;Vous devez renseigner les fourchettes de comptes.;W;O;O;O;','','') ;
Result:=True ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Function FourchetteOk ( CC : TControl ) : Boolean ;
BEGIN
Result:=False ;
if FourchetteVide(CC,'BIL') then Exit ;
if FourchetteVide(CC,'PRO') then Exit ;
if FourchetteVide(CC,'CHA') then Exit ;
Result:=True ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
if Result then HShowMessage('10;Société;La fourchette de comptes que vous avez renseignée n''est pas valide : les deux comptes doivent être renseignés.;W;O;O;O;','','') ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Function VerifiNatureCpteOk ( CC : TControl ; NN : String ) : Boolean ;
Var //Q : TQuery ;
    i : Integer ;
    PasOk : Boolean ;
    TT  : TOB ;
    SQL : String ;
    C1,C2 : String ;
BEGIN
Result:=True ; PasOk:=False ; i:=1 ; TT:=GetTOBCpts(CC) ;
While (i<5) and ((i<3) or (NN<>'EXT')) do
   BEGIN
   C1:=TT.GetValue('SO_'+NN+'DEB'+IntToStr(i)) ; C2:=TT.GetValue('SO_'+NN+'FIN'+IntToStr(i)) ;
   if C1<>'' then
      BEGIN
      SQL:='Select G_NATUREGENE,G_GENERAL from GENERAUX Where G_GENERAL>="'+C1+'" And G_GENERAL<="'+C2+'"' ;
      if NN='BIL' then SQL:=SQL+' AND (G_NATUREGENE="CHA" OR G_NATUREGENE="PRO" OR G_NATUREGENE="EXT")' else
       if NN='CHA' then SQL:=SQL+' AND G_NATUREGENE<>"CHA"' else
        if NN='PRO' then SQL:=SQL+' AND G_NATUREGENE<>"PRO"' else
         if NN='EXT' then SQL:=SQL+' AND G_NATUREGENE<>"EXT"' ;
//      Q:=OpenSql(SQL,True) ; if Not Q.EOF then PasOk:=True ; Ferme(Q) ;
      PasOk:= ExisteSQL(SQL); 
      END ;
   Inc(i) ; if PasOk then Break ;
   END ;
if PasOk then
   BEGIN
   HShowMessage('12;Société;La fourchette de comptes que vous avez renseignée n''est pas en accord avec la nature des comptes généraux.;W;O;O;O;','','') ;
   Result:=False ;
   END ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Function NatureCpteOk ( CC : TControl ) : Boolean ;
BEGIN
Result:=False ;
if Not VerifiNatureCpteOk(CC,'BIL') then Exit ;
if Not VerifiNatureCpteOk(CC,'CHA') then Exit ;
if Not VerifiNatureCpteOk(CC,'PRO') then Exit ;
if Not VerifiNatureCpteOk(CC,'EXT') then Exit ;
Result:=True ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Function GenExiste ( Nam : String ; FF : TForm; Vide : boolean = false;  Libelle : string = '') : boolean ;
Var CC : THCritMaskEdit ;
		messageErr : string;
BEGIN
  Result:=True ;
  if Libelle <> '' then MessageErr := Libelle
                   else MessageErr := 'Le compte général n''est pas valide.';
  CC:=THCritMaskEdit(GetFromNam(Nam,FF)) ;
  if (Vide) and (CC.Text = '') then
  begin
//    HShowMessage('0;Société;Le compte général n''est pas valide.;E;O;O;O;','','') ;
    HShowMessage('0;Société;'+MessageErr+';E;O;O;O;','','') ;
    Result:=False ;
    if CC.CanFocus then CC.SetFocus ;
    exit;
  end;
  if CC.Text<>'' then
  begin
    if Not Presence('GENERAUX','G_GENERAL',CC.Text) then
    begin
//      HShowMessage('0;Société;Le compte général n''est pas valide.;E;O;O;O;','','') ;
    	HShowMessage('0;Société;'+MessageErr+';E;O;O;O;','','') ;
      Result:=False ; if CC.CanFocus then CC.SetFocus ;
    end;
  end;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Function JalExiste ( Nam : String ; FF : TForm; Vide : boolean = false ) : boolean ;
Var CC : THCritMaskEdit ;
BEGIN
  Result:=True ;
  CC:=THCritMaskEdit(GetFromNam(Nam,FF)) ;
  if (Vide) and (CC.Text = '') then
  begin
    HShowMessage('0;Société;Le journal n''est pas valide.;E;O;O;O;','','') ;
    Result:=False ;
    if CC.CanFocus then CC.SetFocus ;
    exit;
  end;
  if CC.Text<>'' then
  begin
    if Not Presence('JOURNAL','J_JOURNAL',CC.Text) then
    begin
      HShowMessage('0;Société;Le journal n''est pas valide.;E;O;O;O;','','') ;
      Result:=False ; if CC.CanFocus then CC.SetFocus ;
    end;
  end;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Function AuxExiste ( Nam : String ; FF : TForm; Vide : boolean = false ) : boolean ;
Var CC : THCritMaskEdit ;
BEGIN
  Result:=True ;
  CC:=THCritMaskEdit(GetFromNam(Nam,FF)) ;
  if (Vide) and (CC.Text = '') then
  begin
    HShowMessage('0;Société;Le compte auxiliaire n''est pas valide.;E;O;O;O;','','') ;
    Result:=False ;
    if CC.CanFocus then CC.SetFocus ;
    exit;
  end;
  if CC.Text<>'' then
  begin
    if Not Presence('TIERS','T_AUXILIAIRE',CC.Text) then
    begin
      HShowMessage('0;Société;Le compte auxiliaire n''est pas valide.;E;O;O;O;','','') ;
      Result:=False ; if CC.CanFocus then CC.SetFocus ;
    end;
  end;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
// TGA Contrôle existence compte à nouveau
Function VerifExisteCptsImmo ( CC : TControl ) : boolean ;
Var FF : TForm ;
BEGIN
Result:=False ;
FF:=GetLaForm(CC) ;
if ExisteNam('SO_IMMOANOUVEAU',FF) then
   if Not GenExiste('SO_IMMOANOUVEAU',FF) then Exit ;
Result:=True ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
// PGR 23/01/2006 Ctl par. d'intégration des écritures de dotation en compta
// PGR 8/02/2006  Ajout test assigned sur les champs  et supression de traduirememoire sur HShowMessage
Function VerifIntegrationEcriture ( CC : TControl ) : boolean ;
var FF : TForm ;
   combo:THValComboBox;
   type_jr : Boolean;
begin
  Result:=False ;
  FF:=GetLaForm(CC) ;

  if ExisteNam('SO_IMMOJALDOTDEF',FF) then
  begin
    //CTl tous les champs obligatoires sauf le libellé dotation
    combo:=THValComboBox(GetFromNam('SO_IMMOJALDOTDEF',FF));

    if assigned(combo) and (combo.Value='') then
    begin
      // TGA 23/06/2005 changement nature du message ( paramêtre E au lieu de W)
      HShowMessage('32;Intégration des écritures;Tous les champs doivent être renseignés.;E;O;O;O;','','') ;
      combo.SetFocus;
      exit;
    end;

    combo:=THValComboBox(GetFromNam('SO_IMMODOTCHOIXDET',FF));

    if assigned(combo) and (combo.Value='') then
    begin
      // TGA 23/06/2005 changement nature du message ( paramêtre E au lieu de W)
      HShowMessage('32;Intégration des écritures;Tous les champs doivent être renseignés.;E;O;O;O;','','') ;
      combo.SetFocus;
      exit;
    end;

    combo:=THValComboBox(GetFromNam('SO_IMMODOTCHOIXTYP',FF));

    if assigned(combo) and (combo.Value='') then
    begin
      // TGA 23/06/2005 changement nature du message ( paramêtre E au lieu de W)
      HShowMessage('32;Intégration des écritures;Tous les champs doivent être renseignés.;E;O;O;O;','','') ;
      combo.SetFocus;
      exit;
    end;

    combo:=THValComboBox(GetFromNam('SO_IMMOJALECHDEF',FF));

    if assigned(combo) and (combo.Value='') then
    begin
      // TGA 23/06/2005 changement nature du message ( paramêtre E au lieu de W)
      HShowMessage('32;Intégration des écritures;Tous les champs doivent être renseignés.;E;O;O;O;','','') ;
      combo.SetFocus;
      exit;
    end;

    combo:=THValComboBox(GetFromNam('SO_IMMOECHCHOIXDET',FF));

    if assigned(combo) and (combo.Value='') then
    begin
      // TGA 23/06/2005 changement nature du message ( paramêtre E au lieu de W)
      HShowMessage('32;Intégration des écritures;Tous les champs doivent être renseignés.;E;O;O;O;','','') ;
      combo.SetFocus;
      exit;
    end;

    combo:=THValComboBox(GetFromNam('SO_IMMOECHCHOIXTYP',FF));

    if assigned(combo) and (combo.Value='') then
    begin
      // TGA 23/06/2005 changement nature du message ( paramêtre E au lieu de W)
      HShowMessage('32;Intégration des écritures;Tous les champs doivent être renseignés.;E;O;O;O;','','') ;
      combo.SetFocus;
      exit;
    end;

    //Ctl journaux
    combo:=THValComboBox(GetFromNam('SO_IMMOJALDOTDEF',FF));

    if assigned(combo) then
    begin;
      if (combo.Value = GetParamSocSecur('SO_JALATP','')) OR (combo.Value = GetParamSocSecur('SO_JALVTP','')) then
      begin
        HShowMessage('31;Journal de dotation;Les journaux liés aux tiers payeurs sont interdits.;W;O;O;O;','','') ;
        combo.SetFocus;
        exit;
      end;

      //PGR - 6/2/2006 Ctl type d'écriture dotation : simulation et situation interdites si journaux en saisie bordereau
      combo:=THValComboBox(GetFromNam('SO_IMMOJALDOTDEF',FF));
      type_jr := (GetColonneSQL('JOURNAL','J_MODESAISIE','J_JOURNAL="'+combo.value+'"')='-');
      combo:=THValComboBox(GetFromNam('SO_IMMODOTCHOIXTYP',FF));

      if (type_jr = false) and (combo.Value <> 'N') then
      begin
        HShowMessage('33;Journal de dotation;Type d''écritures incompatible avec un journal en saisie bordereau.;W;O;O;O;','','') ;
        combo.SetFocus;
        exit;
      end;
    end;
  end;

  Result:=True ;
end ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF GCGC}

Function VerifExisteCptsGC ( CC : TControl ) : boolean ;
Var FF : TForm ;
BEGIN
Result:=False ;
FF:=GetLaForm(CC) ;
if ExisteNam('SO_GCCPTEESCACH',FF) then
  BEGIN
    if Not GenExiste('SO_GCCPTEESCACH',FF,True,'le Compte escompte Achat n''existe pas') then Exit ;
    if Not GenExiste('SO_GCCPTEESCVTE',FF,True,'le Compte escompte Vente n''existe pas') then Exit ;
    if Not GenExiste('SO_GCCPTEREMACH',FF,True,'le Compte Remise Achat n''existe pas') then Exit ;
    if Not GenExiste('SO_GCCPTEREMVTE',FF,True,'le Compte Remise Vente n''existe pas') then Exit ;
    if Not GenExiste('SO_GCCPTEHTACH',FF,True,'le Compte H.T. Achat n''existe pas') then Exit ;
    if Not GenExiste('SO_GCCPTEHTVTE',FF,True,'le Compte H.T. Vente n''existe pas') then Exit ;
    if Not GenExiste('SO_GCCPTEPORTACH',FF,True,'le Compte Port Achat n''existe pas') then Exit ;
    if Not GenExiste('SO_GCCPTEPORTVTE',FF,True,'le Compte Port Vente n''existe pas') then Exit ;
{$IFNDEF BTP}
    if Not AuxExiste('SO_GCFOUCPTADIFF',FF,True) then Exit ;
    if Not AuxExiste('SO_GCCLICPTADIFF',FF,True) then Exit ;
    if Not AuxExiste('SO_GCFOUCPTADIFFPART',FF,True) then Exit ;
    if Not AuxExiste('SO_GCCLICPTADIFFPART',FF,True) then Exit ;
{$ENDIF}
{$IFDEF GCGC}
    if VH_GC.VTCSeria then  //mcd 09/06/06 champ caché si pas sérialisé, donc impossible à mettre OK
    begin
      if Not GenExiste('SO_GCECARTDEBIT',FF,True) then Exit ;
      if Not GenExiste('SO_GCECARTCREDIT',FF,True) then Exit ;
    end;
{$ENDIF GCGC}
{$IFDEF STK}
    if Not GenExiste('SO_GCCPTESTOCK',FF,True) then Exit;
    if Not GenExiste('SO_GCCPTEVARSTK',FF,True) then Exit;
{$ENDIF STK}
{$IFDEF GCGC}
    if (VH_GC.VTCSeria) then
    begin
      if Not GenExiste('SO_GCVRTINTERNE',FF,True) then Exit ;
      if Not JalExiste('SO_GCCAISSGAL',FF,True) then Exit ;
      if Not GenExiste('SO_GCECARTDEBIT',FF,True) then Exit ;
      if Not GenExiste('SO_GCECARTCREDIT',FF,True) then Exit ;
    end;
{$ENDIF GCGC}
  END ;
{$IFDEF STK}  //mcd 06/06/06
if ExisteNam ('SO_GCTIERSMVSTK', FF) then
begin
  if not ControleStock (FF) then exit;
end;
{$ENDIF STK}
Result:=True ;
END ;
{$ENDIF GCGC}
{$ENDIF EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : JCF
Créé le ...... : 02/08/2001
Modifié le ... :   /  /
Description .. : Opérations à réaliser en cas de modification de la devise
Suite ........ : principale du dossier.
Mots clefs ... : MODE;MULTIDEVISE;PARAMSOC
*****************************************************************}
{$IFNDEF EAGLSERVER}
Procedure ModifDevPrinc ( CC : TControl ) ;
Var FF : TForm ;
    CH : THValComboBox ;
    Q  : TQuery ;
BEGIN
FF:=GetLaForm(CC) ;
if ExisteNam('SO_DEVISEPRINC',FF) then
   BEGIN
   CH:=THValComboBox(GetFromNam('SO_DEVISEPRINC',FF));
   if CH.Value <> DevisePrinc then
      BEGIN
      ExecuteSQL ('Update DEVISE SET D_FONGIBLE="-"');
      Q:=OpenSQL('SELECT D_MONNAIEIN FROM DEVISE WHERE D_DEVISE="'+CH.Value+'"',True) ;
      If Not Q.EOF then
         BEGIN
         if Q.Fields[0].AsString='X' then
         ExecuteSQL ('Update DEVISE SET D_FONGIBLE="X" Where D_DEVISE="'+CH.Value+'"') ;
         END;
      Ferme(Q) ;
      END;
   END;
END;
{$ENDIF EAGLSERVER}

{$IFDEF PGIIMMO}
Function VerifPGIIMMO ( CC : TControl ) : boolean ;
Var FF : TForm ;
    CB1,CB2 : tCheckBox ;
BEGIN
Result:=True ;
If Not (ctxImmo in V_PGI.PGIContexte) then Exit ;
FF:=GetLaForm(CC) ;
CB1:=NIL ; CB2:=NIL ;
if ExisteNam('SO_I_CALCULAMORTEXESORTIE',FF) then CB1:=TCheckBox(GetFromNam('SO_I_CALCULAMORTEXESORTIE',FF));
if ExisteNam('SO_I_CALCULAMORTJOURSORTIE',FF) then CB2:=TCheckBox(GetFromNam('SO_I_CALCULAMORTJOURSORTIE',FF));
If (CB1<>NIL) And (CB2<>NIL) Then
  BEGIN
  If Not CB1.Checked Then CB2.Checked:=FALSE ;
  END ;
(*
If CC.Hint='SCO_I_CRITERES' Then MajParamGlobal(TRUE,FF,1);
If CC.Hint='SCO_I_SAISIE' Then MajParamGlobal(TRUE,FF,2);
If CC.Hint='SCO_I_CRITHIERACHIQUE' Then MajParamGlobal(TRUE,FF,3);
*)
If ExisteNam('SO_I_CRIT01',FF) Then MajParamGlobal(TRUE,FF,1);
If ExisteNam('SO_I_SAISIEONTAUX',FF) Then MajParamGlobal(TRUE,FF,2);
If ExisteNam('SO_I_ETABH',FF) Then MajParamGlobal(TRUE,FF,3);
END ;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/03/2002
Modifié le ... : 21/03/2002
Description .. : LG* Test l'existence d'un compte general
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
Function CVerifCompte( CC : THCritMaskEdit ) : boolean ;
BEGIN
Result:=true ;
if CC.Text='' then exit;
if not ExisteSQL('select G_GENERAL from GENERAUX where G_GENERAL="'+CC.Text+'"') then
 BEGIN
  HShowMessage('23;Société;Le compte n''est pas correctement renseigné.;W;O;O;O;','','') ;
  Result:=false ; CC.Text:='' ; if CC.CanFocus then CC.SetFocus ;
 END;
END;
{$ENDIF EAGLSERVER}


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/03/2002
Modifié le ... : 06/06/2002
Description .. : LG* Test l'existence d'un compte general
Suite ........ :
Suite ........ : -06/06/2002- la mise ajour de la devise ne fonctionnait pas
Suite ........ : pour les valeurs avec virgules ( bug 10116 )
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
Function CVerifLettrageGestion( CC : TControl ) : boolean ;
Var FF : TForm ;
BEGIN
// Test de la page
Result:=True ; FF:=GetLaForm(CC) ;
if ExisteNam('SO_LETTOLERC',FF)  then
  ControleLgNum(GetFromNam('SO_LETTOLERC',FF)) ;
if ExisteNam('SO_LETTOLERF',FF) then
  ControleLgNum(GetFromNam('SO_LETTOLERF',FF)) ;

if Not ExisteNam('SO_LETCHOIXGEN',FF) then Exit ; if Not ExisteNam('SO_LETCHOIXGENC',FF) then Exit ;
if Not ExisteNam('SO_LETECARTDEBIT',FF) then Exit ; if Not ExisteNam('SO_LETECARTCREDIT',FF) then Exit ;
result:=CVerifCompte(THCritMaskEdit(GetFromNam('SO_LETCHOIXGEN',FF))) and
        CVerifCompte(THCritMaskEdit(GetFromNam('SO_LETCHOIXGENC',FF)));
// maj de la devise principale

ExecuteSQL ('Update DEVISE SET D_MAXDEBIT='+VariantToSql(THNumEdit(GetFromNam('SO_LETECARTDEBIT',FF)).Value)
             +' , D_MAXCREDIT='+VariantToSql(THNumEdit(GetFromNam('SO_LETECARTCREDIT',FF)).Value)+' Where D_DEVISE="'+GetParamSoc('SO_DEVISEPRINC')+'"') ;

END;
{$ENDIF EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/03/2002
Modifié le ... :   /  /
Description .. : Contrôle de l'onglet ICC
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
Function CVerifCompteICC( CC : TControl ) : boolean ;
Var FF : TForm ;
BEGIN
// Test de la page
Result:=True ; FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_ICCCOMPTECAPITAL',FF) then Exit ;
result:=CVerifCompte(THCritMaskEdit(GetFromNam('SO_ICCCOMPTECAPITAL',FF))) ;
END;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 22/06/2007
Modifié le ... :   /  /    
Description .. : Verification comptes CRE
Mots clefs ... : CRE  FQ20541
*****************************************************************}
Function CVerifCompteCRE( CC : TControl ) : boolean ;
Var FF : TForm ;
BEGIN
Result:=True ; FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_CPCHARGAVAN',FF) then Exit ;
if Not ExisteNam('SO_CPINTERCOUR',FF) then Exit ;
if Not ExisteNam('SO_CPFRAISFINA',FF) then Exit ;
if Not ExisteNam('SO_CPASSURANCE',FF) then Exit ;
If (not CVerifCompte(THCritMaskEdit(GetFromNam('SO_CPCHARGAVAN',FF))))
or (not CVerifCompte(THCritMaskEdit(GetFromNam('SO_CPINTERCOUR',FF))))
or (not CVerifCompte(THCritMaskEdit(GetFromNam('SO_CPFRAISFINA',FF))))
or (not CVerifCompte(THCritMaskEdit(GetFromNam('SO_CPASSURANCE',FF)))) then
  result:=false;
END;
{$ENDIF EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/03/2002
Modifié le ... :   /  /
Description .. : Contrôle sur l'onglet comptabilite\dates
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
Function CVerifDates( CC : TControl ) : boolean ;
Var FF : TForm ;
BEGIN
// Test de la page
Result:=True ; FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_NBJECRAVANT',FF) then Exit ; if Not ExisteNam('SO_NBJECRAPRES',FF) then Exit ;
if Not ExisteNam('SO_NBJECHAVANT',FF) then Exit ; if Not ExisteNam('SO_NBJECHAPRES',FF) then Exit ;
ControleLgNum(GetFromNam('SO_NBJECRAVANT',FF)) ; ControleLgNum(GetFromNam('SO_NBJECRAPRES',FF)) ;
ControleLgNum(GetFromNam('SO_NBJECHAVANT',FF)) ; ControleLgNum(GetFromNam('SO_NBJECHAPRES',FF)) ;
result:=true ;
END;
{$ENDIF EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 28/06/2005
Modifié le ... :   /  /
Description .. : Vérification de la cohérence du nombre de décimales des
Suite ........ : prix et quantités
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
Function CVerifDivers ( CC : TControl ) : boolean;
Var FF : TForm;
    SS : TSpinEdit;
    Chk: TCheckBox;   {FP 19/04/2006 FQ17732}
begin
Result:=True ; FF:=GetLaForm(CC) ;
if not ExisteNam('SO_DECQTE',FF) then exit;
if not ExisteNam('SO_DECPRIX',FF) then exit;
SS:=TSpinEdit(GetFromNam('SO_DECQTE',FF)) ;
if Length(SS.Text) > 1 then SS.Value := 9;
SS:=TSpinEdit(GetFromNam('SO_DECPRIX',FF)) ;
if Length(SS.Text) > 1 then SS.Value := 9;
{b FP 19/04/2006 FQ17732}
Chk := TCheckBox(GetFromNam('SO_CONTROLEBUD',FF));
if Chk.Checked and ExisteSQL('SELECT CRA_CODE FROM CMODELRESTANA') then
  begin
  HShowMessage('0;Paramètres société;Le contrôle en saisie sur le budget est incompatible avec la création d''un modèle de restrictions analytiques.;W;O;O;O;','','');
  Chk.Checked := False;
  THValComboBox(GetFromNam('SO_JALCTRLBUD',FF)).Value:='';
  end;
{e FP 19/04/2006 FQ17732}
end;
{$ENDIF EAGLSERVER}

{$IFDEF COMPTA}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 29/05/2006
Modifié le ... : 14/06/2006
Description .. : LG - FB 10678 - changement du message
Suite ........ : - LG - FB 10678 - qd on annule la cloture provisoire en 
Suite ........ : remets le paramsoc en l'etat
Mots clefs ... : 
*****************************************************************}
function CVerifAnoDyna( CC : TControl ) : boolean ;
var
 FF : TForm ;
 lBoCh : boolean ;
begin
 Result:=True ; FF:=GetLaForm(CC) ;
{$IFNDEF CMPGIS35}
 if not ExisteNam('SO_CPANODYNA',FF) then exit;
 lBoCh := TCheckBox(GetFromNam('SO_CPANODYNA',FF)).Checked ;
 if ( lBoCh <> GetParamSocSecur('SO_CPANODYNA',false)) and lBoCh then
  begin
  {$IFDEF ENTREPRISE}
   ExecuteSQL('UPDATE EXERCICE SET EX_ETATCPTA="OUV" WHERE EX_EXERCICE="'+VH^.EnCours.Code+'"');
    {$IFNDEF EAGLSERVER}
    VH^.EnCours.EtatCpta:='OUV' ;
    {$ENDIF EAGLSERVER}
   AvertirCacheServer( 'EXERCICE' ) ; 
  {$ELSE}
  PGIInfo('Vous venez d''activer les A-nouveaux dynamiques, une clôture provisoire va être effectuée. Validez celle-ci.') ;
  if not ClotureComptable(false,true) then
   TCheckBox(GetFromNam('SO_CPANODYNA',FF)).Checked := false
    else
     if VH^.Suivant.Code = '' then
      begin
       PGIInfo('Vous avez activé les A-nouveaux dynamiques, Voulez ouvrir l''exercice suivant') ;
       OuvertureExo ;
      end ;
  {$ENDIF}
  end ;

  {$IFNDEF ENTREPRISE}
  if ( lBoCh <> GetParamSocSecur('SO_CPANODYNA',false)) and not lBoCh then
   begin
    if not AnnuleclotureComptable(FALSE,true) then
     TCheckBox(GetFromNam('SO_CPANODYNA',FF)).Checked := GetParamSocSecur('SO_CPANODYNA',false) ;
   end ;
  {$ENDIF}


{$ENDIF}

end ;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{$IFDEF COMPTA}
function CInitStatutDossier ( CC : TControl ) : boolean ;
var FF : TForm;
    lStLiasseDossier : string;
    lStPlanRevision : string;
begin
  Result:=True ;
  FF := GetLaForm(CC) ;
  if ExisteNam('SO_CPCONTROLELIASSE',FF) then
  begin
    lStLiasseDossier := THEdit(GetFromNam('SO_CPCONTROLELIASSE',FF)).Text;
    if VH^.BStatutLiasse <> nil then
    begin
      VH^.BStatutLiasse.Visible    := Trim(lStLiasseDossier) <> '';
      VH^.BStatutLiasse.ImageIndex := 3;
    end ;
  end;

  if ExisteNam('SO_CPPLANREVISION',FF) then
  begin
    lStPlanRevision := THEdit(GetFromNam('SO_CPPLANREVISION',FF)).Text;
    if VH^.BStatutRevision <> nil then
    begin
      {$IFNDEF CCMP}
{$IFNDEF CMPGIS35}
      AccesBStatutRevision;
//      VH^.BStatutRevision.Visible    := VH^.OkModRic and (Trim(lStPlanRevision) <> '');
      VH^.BStatutRevision.ImageIndex := 3;
{$ENDIF}
      {$ENDIF}
    end ;
  end;

end;
{$ENDIF}


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Guillaume FONTANA
Créé le ...... : 01/03/2005
Modifié le ... :   /  /
Description .. : Contrôle d'integrité de la saisie des ParamSoc CHR
Mots clefs ... : CHR
*****************************************************************}
{$IFNDEF EAGLSERVER}
{$IFDEF CHR}
function HRVerifZones(CC: TControl) : boolean;
var
  FF : TForm;
  StNom : string;
begin
  Result := True;
  FF:=GetLaForm(CC) ;
  {Initialisation des débiteurs par séjour}
  StNom := 'SO_HRGEREDEBITEURSEJOUR';
  if ExisteNam(StNom,FF) then
  begin
    if (TCheckBox(GetFromNam(StNom,FF)).Checked = True )
      and (GetParamSoc(StNom)=false) then
    begin
      MAJ_HDR_TIERSPAYEUR;
    end;
  end;
end;
{$ENDIF CHR}
{$ENDIF EAGLSERVER}

{*****************************************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/03/2002
Modifié le ... : 21/03/2002
Description .. : LG* Ajout de la gestion de l'onglet lettrage\gestion :
Suite ........ : fonction CVerifCompteLettrage
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
Function SauvePageSoc ( CC : TControl ) : boolean ;
BEGIN
  Result:=False ;
  if not verifInfosBTP(CC) then exit;
  if Not VerifCoord(CC) then Exit ;
  if Not LongueurCpteOk(CC) then Exit ;
  if Not GeneAttenteOk(CC)  then Exit ;
  if Not ExisteCptesAttente(CC) then Exit ;
  if Not ExisteCptesHT(CC) then Exit ;
  if Not VerifFourchettes(CC) then Exit ;
  if Not DefautCpteOk(CC) then Exit ;
  if Not VerifDateEntreeEuroOk(CC) then Exit ;
  if not VerifPariteEuro(CC) then Exit ;
  if not VerifCptesEuro(CC) then Exit ;
  if Not CarBourreOk(CC) then Exit ;
  if Not EtabDefautOk(CC) Then Exit ;
  if Not VerifJalTP(CC) Then Exit ;
  if Not VerifCompensation(CC) then Exit;          {FP 21/02/2006}
  if Not VerifExisteCpts(CC) then Exit ;
  if Not PrefixeTiersOk(CC) then Exit ;
  if Not LgNumTiersOk(CC) then Exit ;
  if Not PrefixeArtOk(CC) then Exit ;
  if Not LgCodeArtOk(CC) then Exit ;
  if Not SocTox(CC) then Exit ;
  if Not CVerifLettrageGestion(CC) then Exit;
  if Not CVerifCompteICC(CC) then Exit;
  if Not CVerifCompteCRE(CC) then Exit;
  if Not CVerifDates(CC) then Exit; // contrôle les nombres saisie dans l'onglet Comptabilite\Dates()
  if Not CVerifDivers(CC) then Exit;
  if not ExisteTVAIntracomm(CC) then exit;
  {$IFDEF COMPTA}
  if not CVerifAnoDyna(CC) then exit;
  if not CInitStatutDossier(CC) then Exit;
  {$ENDIF}
  {$IFDEF PGIIMMO}
    if Not VerifPGIIMMO(CC) then Exit ;
  {$ENDIF PGIIMMO}
  ModifDevPrinc(CC);
  //§{$IFDEF GCGC}
  //if Not VerifExisteCptsGC(CC) then Exit ;
  //GCChangeGestUniteMode(CC);
  //{$ENDIF GCGC}
  {$IFDEF AFFAIRE}
  if Not VerifAFDAtes(CC) then Exit ;
  if Not VerifAFApprec(CC) then Exit ;
  {$ENDIF AFFAIRE}
  {$IFDEF GRC}
  if Not LgNumProjetOk(CC) then Exit ;
  if Not LgNumOperOk(CC) then Exit ;
  VerifTabHie(CC);
  VerifGestLienOperPiece(CC);
  {$IFNDEF GIGI} // MD/MCD 01/07/05 - Ged GRC et GI non compatibles
  if (V_PGI.NumVersionSoc > 700) and (ExisteNam('SO_RTGESTIONGED', GetLaForm(CC))) then
     if TCheckBox(GetLaForm(CC).FindComponent('SO_RTGESTIONGED')).Checked then InitializeGedFiles('', MyAfterImportGRC,gfmdata);
  {$ENDIF GIGI}
  {$ENDIF GRC}
  if Not GestionMonoDepot(CC) then Exit;
  {$IFDEF ACCESCBN}
    if VerifBorne(CC, 'SO_QBORNEMINSTOCKCALCUL_', '', 'MIN') then
    begin
      if not VerifBorne(CC, 'SO_QBORNEMINSTOCKCALCUL_', 'SO_QBORNEMAXSTOCKCALCUL_', 'MAX') then
        Exit;
    end
    else
      Exit;
  {$ENDIF ACCESCBN}
{$IFDEF ACCESSCM}
  if not VerifParamGeneraux(CC) then Exit;
  if not VerifParamCalendHoraire(CC) then Exit;
  if not VerifImportExportGPAO(CC) then Exit;
  if not VerifOptionImport(CC) then Exit;
  if not VerifImportExportHost(CC) then Exit;
  if not VerifAffichagePlan(CC) then Exit;
  if not VerifInfoPlanGP(CC) then Exit;
  if not VerifPlacementAuto(CC) then Exit;
  if not VerifEquivalence(CC) then Exit;
  if not VerifUnite(CC) then Exit;
  if not VerifParamSysteme(CC) then Exit;
  if not VerifGestionDesZones(CC) then Exit;
  if not VerifProfilGroupage(CC) then Exit;
{$ENDIF ACCESSCM}

  if Not VerifAfActivite(CC) then Exit ;   //mcd 12/09/2006
//PCH 11-2005 Personnalisation paramsoc
{$IFDEF GIGI}
  if Not VerifAfPerso(CC) then Exit ;
{$ENDIF GIGI}
                            
{$IFNDEF EAGLSERVER}
  // BDU - 30/01/07 - Contrôle des heures d'affichage des plages
  ControleHeure('SO_AFHEUREDEBUTJOUR', TForm(CC.Owner));
  ControleHeure('SO_AFHEUREFINJOUR', TForm(CC.Owner));
  ControleHeure('SO_AFPMFIN', TForm(CC.Owner));
{$ENDIF EAGLSERVER}

GereAnaParamAff(CC);
GereTypeTVAIntraComm(CC);
GereContremarque(CC);
GereFarFae(CC, true);
GereDevisePiece(CC);
{$IFNDEF EAGLSERVER}
GereElimineLigne0(CC);
{$ENDIF EAGLSERVER}
{ GC_JTR_GC15601_Début }
GereMdpMargeMini(CC);
{ GC_JTR_GC15601_Fin }
{$IFDEF CHR}
HRVerifZones(CC);
{$ENDIF CHR}
  {$IFDEF TRESO}
  if not VerifTreso(GetLaForm(CC)) then Exit; {JP 04/11/04}
  {$ENDIF TRESO}
  {$IFDEF STK}
  if ExisteNam('SO_GERELOT', GetLaForm(CC)) then
  begin
    if not tCheckBox(GetLaForm(CC).FindComponent('SO_GERELOT')).Checked then
      tCheckBox(GetLaForm(CC).FindComponent('SO_GEREFICHELOT')).Checked := false;
  end;
  {$ENDIF STK}

  MarquerPublifi(True) ;
  {$IFDEF GESCOM}
    if ExisteNam(SOCdeOuv, GetLaForm(CC)) then
      CdeOuvTraiteAutoTransfo(GetFromNam(SOCdeOuvAutoTransfo, GetLaForm(CC)));
  {$ENDIF GESCOM}

  {$IFDEF COMPTA}
  {$IFNDEF CCMP}
  // GCO - 14/06/2007
  if not TraitePlanDeRevision( CC ) then Exit;
  {$ENDIF}
  {$ENDIF}

  {$IFDEF COMPTA}
  // ajout me pour la sychronisation
  MAJHistoparam ('PARAMSOC', '');
  if (GetParamSocSecur ('SO_CPJOURNALEXPERT', '')<> '') then   // fiche 10446
     MAJHistoparam ('JOURNAUX', '');
  if not VerifExistSynchro (CC) then exit;
   //YMO 06/12/2006 Norme NF 203
  if not CActivConformiteNF203 (CC) then exit;
  {$ENDIF}


  {$IFDEF COMSX}
  MAJHistoparam ('PARAMSOC', '');
  if (GetParamSocSecur ('SO_CPJOURNALEXPERT', '')<> '') then  // fiche 10573
     MAJHistoparam ('JOURNAUX', '');
  if not VerifExistSynchro (CC) then exit;
  {$ENDIF}

  if not VerifMES( CC ) then Exit;
  InitGPOFraisAvances(CC);

	MajTabletteGcTypeFourn(CC);

	MajParamCdeMarche(CC);

  VerifPdrParametrage(CC);
  MajClassificationTMDR(CC);  // CCMX-CEGID TMDR Dev N° 4722
  // GC_20071016_GM_GC15422_DEBUT
  if not CtrlLongueur(CC) then exit;
  // GC_20071016_GM_GC15422_FIN
  {$IFDEF GPAO}
  InitWLBContexteOrigine(CC);
    {$IFDEF AFFAIRE}
      {$IFNDEF PGIMAJVER}
        InitParamAffairesIndustrielles(CC);
      {$ENDIF PGIMAJVER}
    {$ENDIF AFFAIRE}
  {$ENDIF GPAO}

  {$IFDEF STK}
    CtrlDecochageGestionEmplacement(CC);
    CtrlDecochageGestionLot(CC);
    CtrlDecochageGestionNoSerie(CC);
    CtrlDecochageGestionNoSerieGrp(CC);
    CtrlDecochageGestionStatutDispo(CC);
    CtrlDecochageGestionStatutFlux(CC);
    CtrlDecochageGestionMarque(CC);
    CtrlDecochageGestionChoixQualite(CC);
  {$ENDIF STK}
  if Not VerifModeleWord (CC) then exit;;
  VerifGestYplanningGRC(CC);

  Result:=True ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
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
if Not PrefixeTiersOk(CC) then Exit ;
if Not LgNumTiersOk(CC) then Exit ;
if Not PrefixeArtOk(CC) then Exit ;
if Not LgCodeArtOk(CC) then Exit ;
if Not SocTox(CC) then Exit ;
if not VerifMES( CC ) then Exit;
ModifDevPrinc(CC);
{$IFDEF GCGC}
// if Not VerifExisteCptsGC(CC) then Exit ;
{$ENDIF}
MarquerPublifi(True) ;
InitGPOFraisAvances(CC);
MajTabletteGcTypeFourn(CC);
MajParamCdeMarche(CC);
MajClassificationTMDR(CC);  // CCMX-CEGID TMDR Dev N° 4722
{$IFDEF GPAO}
  InitWLBContexteOrigine(CC);
  {$IFDEF AFFAIRE}
    {$IFNDEF PGIMAJVER}
      InitParamAffairesIndustrielles(CC);
    {$ENDIF PGIMAJVER}
  {$ENDIF AFFAIRE}
{$ENDIF GPAO}
Result:=True ;
END ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure MarquerPublifi ( Flag : boolean ) ;
BEGIN
if ctxPCL in V_PGI.PGIContexte then else if Not VH^.LienPublifi then Exit ;
{$IFNDEF SPEC302}
if Flag then SetParamSoc('SO_ACHARGERPUBLIFI','X') else SetParamSoc('SO_ACHARGERPUBLIFI','-') ;
{$ENDIF}
END ;
{$ENDIF EAGLSERVER}

Procedure TripoteStatus ( StAjout : String = '')  ;
Var St,AJ : String ;
BEGIN
AJ:=VH^.CPStatusBarre ;
if ((AJ='') or (AJ='AUC')) then if StAjout='' then Exit ;
St:=GetDefStatus ;
if AJ='PIV' then St:=St+'   '+RechDom('TTDEVISETOUTES',V_PGI.DevisePivot,False)+' ('+V_PGI.DevisePivot+')' else
if AJ='CON' then St:=St+'   '+DateToStr(V_PGI.DateEntree) else
if AJ='EXO' then St:=St+'   '+RechDom('TTEXERCICE',VH^.Encours.Code,False)+' ('+VH^.Encours.Code+')' else
if AJ='EXR' then St:=St+'   '+RechDom('TTEXERCICE',VH^.CPExoRef.Code,False)+' ('+VH^.CPExoRef.Code+')' else
if AJ='CLO' then St:=St+'   '+DateToStr(VH^.DateCloturePer) else
if AJ='BAS' then St:=St+'   '+GetDBPathName(TRUE) else
if AJ='VER' then St:=St+'   '+V_PGI.NumVersion+' du '+DateToStr(V_PGI.DateVersion) else ;
if StAjout<>'' then St:=St+'    '+StAjout ;
ChgStatus(St) ;
END ;

{$IFNDEF EAGLSERVER}
{==============================================================================}
{======================= FONCTIONS IMMOBILISATIONS ============================}
{==============================================================================}
Procedure SetLenMaxSt ( Nam : String ; FF : TForm ; Len : integer ; bMajText : boolean = false ; cb : string = '') ;
Var CC : TControl ;
BEGIN
  CC:=GetFromNam(Nam,FF) ;
  if (CC<>Nil) then
  begin
    if (CC is THCritMaskEdit) then
    begin
    THCritMaskEdit(CC).MaxLength := Len;
    if bMajText then
      THCritMaskEdit(CC).Text := BourreLaDoncSurLesComptes(THCritMaskEdit(CC).Text,cb);
    end;
  END ;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
function SauvePageImmo (CC : TControl) : boolean;
BEGIN
  Result:=TRUE ;
  // TGA Ajout appel du contrôle cpt immo à nouveau
  if Not VerifExisteCptsImmo(CC) then Exit ;

  // PGR 23/01/2006 Ctl par. d'intégration des écritures de dotation en compta
  if not VerifIntegrationEcriture(CC) then Exit ;

  if CC.Hint <> 'SCO_COEFFDEGRESSIF' then exit;
  {$IFDEF AMORTISSEMENT}
  UpdateBaseImmo;
  {$ENDIF}
  Result:=true;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
function ChargePageImmo (CC : TControl) : boolean;
var
  ListeCh : TStringList;
begin
  ListeCh := TStringList.Create;
  if ListeCh<>nil then
  begin
// liste des controles liés aux fourchettes de comptes immo
    ListeCh.Add('SO_CPTEIMMO');
    ListeCh.Add('SO_CPTEFIN');
    ListeCh.Add('SO_CPTEAMORT');
    ListeCh.Add('SO_CPTEDOT');
    ListeCh.Add('SO_CPTEEXPLOIT');
    ListeCh.Add('SO_CPTEDOTEXC');
    ListeCh.Add('SO_CPTEREPEXC');
    ListeCh.Add('SO_CPTEVACEDEE');
    ListeCh.Add('SO_CPTEDEROG');
    ListeCh.Add('SO_CPTEPROVDER');
    ListeCh.Add('SO_CPTEREPDER');
// liste des controles liés aux comptes CB, loc et depot
    ListeCh.Add('SO_CPTECB');
    ListeCh.Add('SO_CPTELOC');
    ListeCh.Add('SO_CPTEDEPOT');
    SetLenMaxCompteImmo(CC,ListeCh,VH^.Cpta[fbGene].Lg);
    ListeCh.Free;
  end;
Result:=True ;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure SetLenMaxCompteImmo(CC : TControl; laListe : TStringList ; Len : integer);
Var i,iDeb, iFin : integer;
    FF : TForm;
BEGIN
  FF := GetlaForm(CC);
// en fonction du panel affiché : sélection des controls a traiter :
//  iDeb := 0;  iFin := 0; //par defaut SO_COMPTES
  if CC.Hint = 'SCO_COMPTES' then begin iDeb := 0; iFin := 10; end
  else if CC.Hint = 'SCO_FOUR2' then begin iDeb := 11; iFin := 13; end
  else Exit;
//pour chaque control de la liste
  for i:=iDeb to iFin do
  begin
    SetLenMaxSt(laListe[i]+'INF',FF,Len,true,'0') ;
    SetLenMaxSt(laListe[i]+'SUP',FF,Len,true,'9') ;
    ;
  end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure AlimZonesProjet ( CC : TControl ) ;
Var FF : TForm ;
    GestionProjet : TCheckBox ;
BEGIN
FF:=GetLaForm(CC) ;

GestionProjet:=TCheckBox(GetFromNam('SO_RTPROJGESTION',FF)) ;

if Not ExisteNam('SO_RTPROJMULTITIERS',FF) then Exit ;
if Not ExisteNam('SO_RTPROJATTACHTIERS',FF) then Exit ;

if GestionProjet.Checked = True then
   begin
   SetParamSoc('SO_RTPROJGESTION',True);
   TCheckBox(GetFromNam('SO_RTPROJMULTITIERS',FF)).Enabled:=True ;
   TCheckBox(GetFromNam('SO_RTPROJATTACHTIERS',FF)).Enabled:=True;
   TCheckBox(GetFromNam('SO_RTNUMPROJETAUTO',FF)).Enabled:=True;
   if TCheckBox(GetFromNam('SO_RTNUMPROJETAUTO',FF)).Checked = True then
      begin
      SetEna('SO_RTCOMPTEURPROJET',FF,True) ;
      SetEna('LSO_RTCOMPTEURPROJET',FF,True) ;
      end
   else
      begin
      SetEna('SO_RTCOMPTEURPROJET',FF,False) ;
      SetEna('LSO_RTCOMPTEURPROJET',FF,False) ;
      end;
   SetEna('SO_RTPROJRESP',FF,True) ;
   SetEna('LSO_RTPROJRESP',FF,True) ;
   end
else
   begin
   SetParamSoc('SO_RTPROJGESTION',False);
   TCheckBox(GetFromNam('SO_RTPROJMULTITIERS',FF)).Enabled:=False ;
   TCheckBox(GetFromNam('SO_RTPROJMULTITIERS',FF)).Checked:=False ;
   TCheckBox(GetFromNam('SO_RTPROJATTACHTIERS',FF)).Enabled:=False;
   TCheckBox(GetFromNam('SO_RTPROJATTACHTIERS',FF)).Checked:=False;
   TCheckBox(GetFromNam('SO_RTNUMPROJETAUTO',FF)).Enabled:=False;
   TCheckBox(GetFromNam('SO_RTNUMPROJETAUTO',FF)).Checked:=False;
   SetEna('SO_RTCOMPTEURPROJET',FF,False) ;
   SetEna('LSO_RTCOMPTEURPROJET',FF,False) ;
   SetEna('SO_RTPROJRESP',FF,False) ;
   SetEna('LSO_RTPROJRESP',FF,False) ;

   SetParamSoc('SO_RTPROPATTACHPROJET',False);
   SetParamSoc('SO_RTPROPGROUPE',False);
   end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure FormatZoneCli ( CC : TControl ) ;
Var FF : TForm ;
    GestionFormat : TCheckBox ;
BEGIN
  FF:=GetLaForm(CC) ;

  GestionFormat:=TCheckBox(GetFromNam('SO_GCFORMATEZONECLI',FF)) ;

  if GestionFormat.Checked = True then
     begin
     SetParamSoc('SO_GCFORMATEZONECLI',True);
     TCheckBox(GetFromNam('SO_FMTTIERSLIBELLE',FF)).Enabled:=True ;
     TCheckBox(GetFromNam('SO_FMTTIERSPARTIC',FF)).Enabled:=True ;
     TCheckBox(GetFromNam('SO_FMTTIERSTAILLEMOT',FF)).Enabled:=True ;
     TCheckBox(GetFromNam('SO_FMTTIERSPRENOM',FF)).Enabled:=True ;
     TCheckBox(GetFromNam('SO_FMTSOCLIBELLE',FF)).Enabled:=True ;
     TCheckBox(GetFromNam('SO_FMTSOCPRENOM',FF)).Enabled:=True ;
     TCheckBox(GetFromNam('SO_FMTTIERSADR1',FF)).Enabled:=True ;
     TCheckBox(GetFromNam('SO_FMTTIERSADR2',FF)).Enabled:=True ;
     TCheckBox(GetFromNam('SO_FMTTIERSADR3',FF)).Enabled:=True ;
     TCheckBox(GetFromNam('SO_FMTTIERSVILLE',FF)).Enabled:=True ;
     TCheckBox(GetFromNam('SO_FMTTIERSSUPESPACE',FF)).Enabled:=True ;
     end
  else
     begin
     SetParamSoc('SO_GCFORMATEZONECLI',False);
     TCheckBox(GetFromNam('SO_FMTTIERSLIBELLE',FF)).Enabled:=False ;
     TCheckBox(GetFromNam('SO_FMTTIERSPARTIC',FF)).Enabled:=False ;
     TCheckBox(GetFromNam('SO_FMTTIERSTAILLEMOT',FF)).Enabled:=False ;
     TCheckBox(GetFromNam('SO_FMTTIERSPRENOM',FF)).Enabled:=False ;
     TCheckBox(GetFromNam('SO_FMTSOCLIBELLE',FF)).Enabled:=False ;
     TCheckBox(GetFromNam('SO_FMTSOCPRENOM',FF)).Enabled:=False ;
     TCheckBox(GetFromNam('SO_FMTTIERSADR1',FF)).Enabled:=False ;
     TCheckBox(GetFromNam('SO_FMTTIERSADR2',FF)).Enabled:=False ;
     TCheckBox(GetFromNam('SO_FMTTIERSADR3',FF)).Enabled:=False ;
     TCheckBox(GetFromNam('SO_FMTTIERSVILLE',FF)).Enabled:=False ;     
     TCheckBox(GetFromNam('SO_FMTTIERSSUPESPACE',FF)).Enabled:=False ;
     end
end;
{$ENDIF EAGLSERVER}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. Tarcy
Créé le ...... : 30/10/2001
Modifié le ... : 30/10/2001
Description .. : Paramétrage des échanges PGI
Mots clefs ... : FO;BO;
*****************************************************************}
{$IFNDEF EAGLSERVER}
Function SocTox(CC : TControl) : boolean ;
Var FF : TForm ;
begin
  Result := True ;
  FF:=GetLaForm(CC) ;
  if Not ExisteNam('SO_GCTOXCONFIRM',FF) then Exit ;
  if TCheckBox(GetFromNam('SO_GCTOXCONFIRM',FF)).Checked=True then
     begin
     TCheckBox(GetFromNam('SO_GCTOXAFFDONNEES',FF)).Enabled:=True ;
     TCheckBox(GetFromNam('SO_GCTOXAFFDONNEES',FF)).Checked:=True ;
     TCheckBox(GetFromNam('SO_GCTOXAFFERREURS',FF)).Enabled:=True ;
     TCheckBox(GetFromNam('SO_GCTOXAFFERREURS',FF)).Checked:=True ;
     end else
     begin
     TCheckBox(GetFromNam('SO_GCTOXAFFDONNEES',FF)).Checked:=False ;
     TCheckBox(GetFromNam('SO_GCTOXAFFDONNEES',FF)).Enabled:=False ;
     TCheckBox(GetFromNam('SO_GCTOXAFFERREURS',FF)).Checked:=False ;
     TCheckBox(GetFromNam('SO_GCTOXAFFERREURS',FF)).Enabled:=False ;
     end ;
end ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Function LgNumProjetOk ( CC : TControl ) : boolean;
Var FF : TForm ;
    NumProjetAuto : TCheckBox ;
    CompteurProjet: TEdit ;
//    LgCompteur : Integer;
BEGIN
Result:=True;
FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_RTNUMPROJETAUTO',FF) then Exit ;
NumProjetAuto:=TCheckBox(GetFromNam('SO_RTNUMPROJETAUTO',FF)) ;
if Not NumProjetAuto.Checked then Exit ;
CompteurProjet:=TEdit(GetFromNam('SO_RTCOMPTEURPROJET',FF));
CompteurProjet.Text:=Trim(CompteurProjet.Text);
if CompteurProjet.Text='' then CompteurProjet.Text:='0';
//LgCompteur:=Length(CompteurProjet.Text) ;
Result:=ControleChrono (CompteurProjet.Text);     // Contrôle de numéricité
if Result=False then
   begin
   if CompteurProjet.CanFocus then CompteurProjet.SetFocus ;
   Exit;
   end;
END;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Function LgNumOperOk ( CC : TControl ) : boolean;
Var FF : TForm ;
    NumOperAuto : TCheckBox ;
    CompteurOper: TEdit ;
    Suffixe : string;
//    LgCompteur : Integer;
BEGIN
Result:=True;
Suffixe := '';
FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_RTNUMOPERAUTO',FF) then
   if Not ExisteNam('SO_RTNUMOPERAUTO_',FF) then Exit
   else Suffixe := '_';
NumOperAuto:=TCheckBox(GetFromNam('SO_RTNUMOPERAUTO'+Suffixe,FF)) ;
if Not NumOperAuto.Checked then Exit ;
CompteurOper:=TEdit(GetFromNam('SO_RTCOMPTEUROPER'+Suffixe,FF));
CompteurOper.Text:=Trim(CompteurOper.Text);
if CompteurOper.Text='' then CompteurOper.Text:='0';
//LgCompteur:=Length(CompteurOper.Text) ;
Result:=ControleChrono (CompteurOper.Text);     // Contrôle de numéricité
if Result=False then
   begin
   if CompteurOper.CanFocus then CompteurOper.SetFocus ;
   Exit;
   end;
END;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
function GestionMonoDepot( CC : TControl ) : boolean ;
var FF : TForm ;
    Stg : String ;
begin
Result := True;
FF:=GetLaForm(CC) ;
//CT : Paramétrage des dépôts
if ExisteNam('SO_GCMULTIDEPOTS',FF) then
   begin
   if TCheckBox(GetFromNam('SO_GCMULTIDEPOTS',FF)).Checked=False then    //fonctionnement mono-dépôt
      begin
      // Si on est passé de multi-dépôts à mono-dépôt
      if TCheckBox(GetFromNam('SO_GCMULTIDEPOTS',FF)).Checked<>GetParamSoc('SO_GCMULTIDEPOTS') then
         begin
         // Lancement de la fiche du passage entre multi et mono-dépôt
         Stg := AGLLanceFiche('MBO','PASS_MONOMULTIDEP','','','');
         if Stg <> 'TRUE' then TCheckBox(GetFromNam('SO_GCMULTIDEPOTS',FF)).Checked:=True;  // je ne change pas le paramètre société
         end;
      end;
   if TCheckBox(GetFromNam('SO_GCMULTIDEPOTS',FF)).Checked=False then
      begin
      SetEna('SO_GCDEPOTDEFAUT',FF,False) ;
      SetEna('SO_GCLIAISONAUTODEP_ETAB',FF,False) ;
      end
   else
      begin
      SetEna('SO_GCDEPOTDEFAUT',FF,True) ;
      SetEna('SO_GCLIAISONAUTODEP_ETAB',FF,True) ;
      end;
   end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure MajGPPEcheances( CC : TControl );
BEGIN
  ExecuteSQL('Update PARPIECE SET GPP_MODEECHEANCES="' + THValComboBox(CC).Value + '" where GPP_MODEECHEANCES=""');
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure VerifTabHie ( CC : TControl );
Var FF : TForm ;
BEGIN
  FF:=GetLaForm(CC) ;
  if ExisteNam('SO_RTPROPTABHIE',FF) then TraiteTabHie(CC,'SO_RTPROPTABHIE') ;
  if ExisteNam('SO_RTACTTABHIE',FF) then TraiteTabHie(CC,'SO_RTACTTABHIE');
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure VerifPdrParametrage(CC: TControl);
begin
  if ExisteNam('SO_PDRNATUREPDRGA',TForm(CC.OWner)) then
  begin
    {Paramètrage Prix de revient}
    if (THValComboBox(GetFromNam('SO_PDRNATUREPDRGA', TForm(CC.OWner))).value='ORD') then
    begin
      if (pos(THValComboBox(GetFromNam('SO_PDRTYPEPDRGA', TForm(CC.OWner))).value,'ORD/ORT')=0)  or (THValComboBox(GetFromNam('SO_PDRTYPEPDRGA', TForm(CC.OWner))).value='') then
        THValComboBox(GetFromNam('SO_PDRTYPEPDRGA', TForm(CC.OWner))).value := 'ORT';
    end
    else
    begin
      if (pos(THValComboBox(GetFromNam('SO_PDRTYPEPDRGA', TForm(CC.OWner))).value,'ORD/ORT')>0) or (THValComboBox(GetFromNam('SO_PDRTYPEPDRGA', TForm(CC.OWner))).value='') then
        THValComboBox(GetFromNam('SO_PDRTYPEPDRGA', TForm(CC.OWner))).value := 'ACT';
    end;
  end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Function VerifExistSynchro ( CC : TControl ) : Boolean;
Var FF : TForm ;
Val, ValLien    : THValComboBox;
Q               : TQuery;
BEGIN
  Result := TRUE;
  FF:=GetLaForm(CC);
  if ExisteNam('SO_CPMODESYNCHRO',FF) then
  begin
      ValLien := THValComboBox(GetFromNam('SO_CPLIENGAMME',FF));
                                // fiche 18836
      if (ValLien.Value = 'S1') and (TCheckBox(GetFromNam('SO_CPMODESYNCHRO',FF)).checked) then
      begin
                 Val := THValComboBox(GetFromNam('SO_CPPOINTAGESX',FF));
                 if (Val.Value <> 'CLI') and (Val.Value <> 'EXP') then
                 begin
                                  PGIBox ('Attention, vous ne pouvez pas indiquer "Pas de gestion du pointage" dans le cadre d''une liaison avec Business Line. '+#10#13
                                  +' Vous devez indiquer pointage par Expert comptable ou par Client ', 'Paramètres sociétés');
                                  Val.Value := 'CLI'; //fiche 18837
                                  Result := FALSE;
                 end;
      end;
      if (ValLien.Value = 'S1') and not (TCheckBox(GetFromNam('SO_CPMODESYNCHRO',FF)).checked) then
      begin
                 Val := THValComboBox(GetFromNam('SO_CPPOINTAGESX',FF));
                 Val.Value := 'AUC'
      end
  end;
  if ExisteNam('SO_CPJOURNALEXPERT',FF) then // fiche 10532
  begin
    ValLien := THValComboBox(GetFromNam('SO_CPJOURNALEXPERT',FF));
{$ifdef COMPTA}
    if ValLien.text <> '' then
        MAJHistoparam ('JOURNAUX', '');
{$endif}
{$ifdef COMSX}
    if ValLien.text <> '' then
        MAJHistoparam ('JOURNAUX', '');
{$endif}
     if ValLien.text = '<<Tous>>' then
     begin
        Q := OpenSQL('SELECT  J_Journal From Journal Where J_NATUREJAL<>"ANO" AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ODA"', TRUE);
        ValLien.text := '';
        while not Q.EOF do
        begin
        ValLien.text := ValLien.text+';'+ Q.FindField ('J_Journal').asstring;
        Q.next;
        end;
        Ferme (Q);
     end;
  end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure TraiteTabHie ( CC : TControl; stName : string );
Var FF : TForm ;
    CBGestion : TCheckBox ;
    stTablMaitre :string;
begin
  FF:=GetLaForm(CC) ;
  CBGestion:=TCheckBox(GetFromNam(stName,FF)) ;
  if stName = 'SO_RTPROPTABHIE' then
    stTablMaitre:='RTTYPEPERSPECTIVE'
  else if stName = 'SO_RTACTTABHIE' then
          stTablMaitre:='RTTYPEACTION';
  if CBGestion.Checked = true then
    if not ExisteSQL ('Select YDL_CODEHDTLINK from YDATATYPELINKS WHERE YDL_MDATATYPE="'+stTablMaitre+'"') then
       RTTabHie(CC,stName);

  if CBGestion.Checked = false then
    if ExisteSQL ('Select YDL_CODEHDTLINK from YDATATYPELINKS WHERE YDL_MDATATYPE="'+stTablMaitre+'"') then
       RTTabHie(CC,stName);
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure RTTabHie(CC : TControl;stName : String) ;
Var FF : TForm ;
    i : integer;
    stNom,stTablMaitre,stTablesclave,stCle : string;
begin
  FF:=GetLaForm(CC) ;
  if Not ExisteNam(stName,FF) then Exit ;
  if stName = 'SO_RTPROPTABHIE' then
  begin
    stNom:=TraduireMemoire('des propositions');
    stTablMaitre:='RTTYPEPERSPECTIVE';
    stTablesclave:='RTRPRLIBPERSPECTIVE';
  end
  else if stName = 'SO_RTACTTABHIE' then
        begin
          stNom:=TraduireMemoire('des actions');
          stTablMaitre:='RTTYPEACTION';
          stTablesclave:='RTRPRLIBACTION';
        end;

  if CtxScot in V_PGI.PGIContexte then stCle := 'GI' + stTablesclave
  else stCle := 'GC' + stTablesclave;
  if TCheckBox(GetFromNam(stName,FF)).Checked=True then
  begin
    if PGIAsk(TraduireMemoire('Confirmez-vous l''utilisation des tablettes hiérarchiques ')+stNom+' ?',TraduireMemoire('Tablettes hiérarchiques'))=mrYes then
    begin
      for i := 1 to 3 do
      begin
        ExecuteSQL('INSERT INTO YDATATYPELINKS (YDL_CODEHDTLINK,YDL_LIBELLE,YDL_MDATATYPE,YDL_SDATATYPE,YDL_TYPELINK,YDL_PREDEFINI) values ("'+
        stCle+IntToStr(i)+'","'+TraduireMemoire('Table libre ')+IntToStr(i)+' '+stNom+'","'+stTablMaitre+'","'+stTablesclave+IntToStr(i)+
        '","NOR","CEG")');
      end;
    end
    else
        TCheckBox(GetFromNam(stName,FF)).Checked:=False;
  end
  else
  begin
    if PGIAsk(TraduireMemoire('Confirmez-vous la suppression des tablettes hiérarchiques ')+stNom+' ?',TraduireMemoire('Tablettes hiérarchiques'))=mrYes then
    begin
      for i := 1 to 3 do
      begin
        ExecuteSQL('DELETE FROM YDATATYPETREES WHERE YDT_CODEHDTLINK="'+stCle+IntToStr(i)+'"');
        ExecuteSQL('DELETE FROM YDATATYPELINKS WHERE YDL_CODEHDTLINK="'+stCle+IntToStr(i)+'"');
      end;
    end
    else
        TCheckBox(GetFromNam(stName,FF)).Checked:=True;
  end
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure VerifGestLienOperPiece ( CC : TControl );
Var FF : TForm ;
BEGIN
  FF:=GetLaForm(CC) ;
  if ExisteNam('SO_RTLIENOPERATIONPIECE',FF) then
  begin
    if (TCheckBox(GetFromNam('SO_RTLIENOPERATIONPIECE',FF)).Checked=False) and
       (GetParamSocSecur('SO_RTLIENOPERATIONPIECE',False)) then
    begin
      if PGIAsk(TraduireMemoire('Confirmez-vous la suppression du lien avec les opérations commerciales ?'),TraduireMemoire('Pièces'))=mrYes then
      begin
        ExecuteSQL('UPDATE PIECE SET GP_OPERATION = "" WHERE GP_OPERATION<>""');
      end
      else
          TCheckBox(GetFromNam('SO_RTLIENOPERATIONPIECE',FF)).Checked:=True;
    end
  end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure CtrlNbDecimales(CC : TControl; stName : string );
Var FF : TForm;
    NbDecNew : integer;
begin
  FF := GetLaForm(CC);
  NbDecNew := StrToInt(THEdit(GetFromNam(stName, FF)).Text);
  if NbDecNew <> GetParamSoc(stName, True) then
  begin
    if PGIAsk('ATTENTION !!!' + #13 +
              'Ce changement peut entraîner des pertes d''informations' + #13 +
              'Confirmez-vous votre modification ?',
              'Nombre de décimales') = mrNo then
    begin
        THEdit(CC).Text := IntToStr(GetParamSoc(stName, True));
    end;
  end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure BornesVisibles(CC: TControl);
var
  FF       : TForm;
begin
  FF := GetLaForm(CC);
  if      (THValComboBox(CC).Value = '1') then //Quantités
  begin
    SetInvi('LSO_QBORNEMINSTOCKCALCUL_', FF); SetInvi('SO_QBORNEMINSTOCKCALCUL_', FF);
    SetInvi('LSO_QBORNEMAXSTOCKCALCUL_', FF); SetInvi('SO_QBORNEMAXSTOCKCALCUL_', FF);
  end
  else if (THValComboBox(CC).Value = '2') then //Jours
  begin
    SetVisi('LSO_QBORNEMINSTOCKCALCUL_', FF); SetVisi('SO_QBORNEMINSTOCKCALCUL_', FF);
    SetVisi('LSO_QBORNEMAXSTOCKCALCUL_', FF); SetVisi('SO_QBORNEMAXSTOCKCALCUL_', FF);
  end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
function VerifBorne(CC: TControl; NameMin, NameMax, Quoi: string): Boolean;
var
  FF       : TForm;
  ValeurChampMin, ValeurChampMax : integer;
  LeControle : THCritMaskEdit;
begin
  Result := True;
  FF := GetLaForm(CC);
  if      (Quoi = 'MIN') and (FF.Visible) and (ExisteNam(NameMin,FF))
          and (assigned(THEdit(GetFromNam(NameMin, FF)))) then
  begin
    if THEdit(GetFromNam(NameMin, FF)).Visible then
    begin
      ValeurChampMin := ValeurI(THEdit(GetFromNam(NameMin, FF)).Text);
      if not (ValeurChampMin >= 0) then
      begin
        PgiError(TraduireMemoire('La valeur de ce paramètre doit être supérieure ou égal à 0.'));
        LeControle := THCritMaskEdit(GetFromNam(NameMin,FF));
        LeControle.SetFocus;
        Result := False;
      end;
    end;
  end
  else if (Quoi = 'MAX') and (FF.Visible) and (ExisteNam(NameMin,FF)) and (ExisteNam(NameMax,FF))
          and (assigned(THEdit(GetFromNam(NameMin, FF))))
          and (assigned(THEdit(GetFromNam(NameMax, FF)))) then
  begin
    ValeurChampMin := ValeurI(THEdit(GetFromNam(NameMin, FF)).Text);
    ValeurChampMax := ValeurI(THEdit(GetFromNam(NameMax, FF)).Text);
    if (THEdit(GetFromNam(NameMin, FF)).Visible) and (THEdit(GetFromNam(NameMax, FF)).Visible) then
    begin
      if not (ValeurChampMax > ValeurChampMin) then
      begin
        PgiError(TraduireMemoire('La valeur du paramètre maximum doit être strictement supérieure à la valeur du paramètre minimum.'));
        Result := False;
        LeControle := THCritMaskEdit(GetFromNam(NameMax,FF));
        LeControle.SetFocus;
      end;
    end;
  end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure GereZonesMarche(CC: TControl; Flux: string);
var
  FF        : TForm ;
  CbMarche  : TCheckBox;
begin
  FF := GetLaForm(CC);
  CbMarche := TCheckBox(CC);
  if CbMarche.Checked then
  begin
    if      Flux = 'G' then
    begin
      SetEna('SO_CDEMARCHEA',FF,False);
      SetEna('SO_CDEMARCHEV',FF,False);
      SetEna('SO_CDEMARCHEDEPOT',FF,False);
      SetEna('SO_CDEMARCHESURMARCHEEPUISE',FF,False);
      SetEna('SO_CDEMARCHEINIT',FF,False);
      SetEna('SO_CDEMARCHEDUP',FF,False);
      SetEna('LSCO_MARCHEACH',FF,False);
      SetEna('SCO_MARCHEACH',FF,False);
      SetEna('LSO_CDEMARCHEACH',FF,False);
      SetEna('SO_CDEMARCHEACH',FF,False);
      SetEna('LSO_CDEMARCHEAPPELACH',FF,False);
      SetEna('SO_CDEMARCHEAPPELACH',FF,False);
      SetEna('LSO_CDEMARCHEAPPELDATEACH',FF,False);
      SetEna('SO_CDEMARCHEAPPELDATEACH',FF,False);
      SetEna('SO_CDEMARCHECTRLDVALACH',FF,False);
      SetEna('LSCO_MARCHEVTE',FF,False);
      SetEna('SCO_MARCHEVTE',FF,False);
      SetEna('LSO_CDEMARCHEVTE',FF,False);
      SetEna('SO_CDEMARCHEVTE',FF,False);
      SetEna('LSO_CDEMARCHEAPPELVTE',FF,False);
      SetEna('SO_CDEMARCHEAPPELVTE',FF,False);
      SetEna('LSO_CDEMARCHEAPPELDATEVTE',FF,False);
      SetEna('SO_CDEMARCHEAPPELDATEVTE',FF,False);
      SetEna('SO_CDEMARCHECTRLDVALVTE',FF,False);
    end
    else if Flux = 'A' then
    begin
      SetEna('LSCO_MARCHEACH',FF,False);
      SetEna('SCO_MARCHEACH',FF,False);
      SetEna('LSO_CDEMARCHEACH',FF,False);
      SetEna('SO_CDEMARCHEACH',FF,False);
      SetEna('LSO_CDEMARCHEAPPELACH',FF,False);
      SetEna('SO_CDEMARCHEAPPELACH',FF,False);
      SetEna('LSO_CDEMARCHEAPPELDATEACH',FF,False);
      SetEna('SO_CDEMARCHEAPPELDATEACH',FF,False);
      SetEna('SO_CDEMARCHECTRLDVALACH',FF,False);
    end
    else if Flux = 'V' then
    begin
      SetEna('LSCO_MARCHEVTE',FF,False);
      SetEna('SCO_MARCHEVTE',FF,False);
      SetEna('LSO_CDEMARCHEVTE',FF,False);
      SetEna('SO_CDEMARCHEVTE',FF,False);
      SetEna('LSO_CDEMARCHEAPPELVTE',FF,False);
      SetEna('SO_CDEMARCHEAPPELVTE',FF,False);
      SetEna('LSO_CDEMARCHEAPPELDATEVTE',FF,False);
      SetEna('SO_CDEMARCHEAPPELDATEVTE',FF,False);
      SetEna('SO_CDEMARCHECTRLDVALVTE',FF,False);
    end
    ;
  end
  else
  begin
    if      Flux = 'G' then
    begin
      if (ExisteNam('SO_CDEMARCHEA',FF)) and (assigned(TCheckBox(GetFromNam('SO_CDEMARCHEA', FF)))) then
        TCheckBox(GetFromNam('SO_CDEMARCHEA',FF)).State := cbUnchecked;
      if (ExisteNam('SO_CDEMARCHEV',FF)) and (assigned(TCheckBox(GetFromNam('SO_CDEMARCHEV', FF)))) then
        TCheckBox(GetFromNam('SO_CDEMARCHEV',FF)).State := cbUnchecked;
      SetEna('SO_CDEMARCHEA',FF,True);
      SetEna('SO_CDEMARCHEV',FF,True);
      SetEna('SO_CDEMARCHEDEPOT',FF,True);
      SetEna('SO_CDEMARCHESURMARCHEEPUISE',FF,True);
      SetEna('SO_CDEMARCHEINIT',FF,True);
      SetEna('SO_CDEMARCHEDUP',FF,True);
    end
    else if Flux = 'A' then
    begin
      SetEna('LSCO_MARCHEACH',FF,True);
      SetEna('SCO_MARCHEACH',FF,True);
      SetEna('LSO_CDEMARCHEACH',FF,True);
      SetEna('SO_CDEMARCHEACH',FF,True);
      SetEna('LSO_CDEMARCHEAPPELACH',FF,True);
      SetEna('SO_CDEMARCHEAPPELACH',FF,True);
      SetEna('LSO_CDEMARCHEAPPELDATEACH',FF,True);
      SetEna('SO_CDEMARCHEAPPELDATEACH',FF,True);
      SetEna('SO_CDEMARCHECTRLDVALACH',FF,False);
    end
    else if Flux = 'V' then
    begin
      SetEna('LSCO_MARCHEVTE',FF,True);
      SetEna('SCO_MARCHEVTE',FF,True);
      SetEna('LSO_CDEMARCHEVTE',FF,True);
      SetEna('SO_CDEMARCHEVTE',FF,True);
      SetEna('LSO_CDEMARCHEAPPELVTE',FF,True);
      SetEna('SO_CDEMARCHEAPPELVTE',FF,True);
      SetEna('LSO_CDEMARCHEAPPELDATEVTE',FF,True);
      SetEna('SO_CDEMARCHEAPPELDATEVTE',FF,True);
      SetEna('SO_CDEMARCHECTRLDVALVTE',FF,False);
    end
    ;
  end;
end;
{$ENDIF EAGLSERVER}

{***********A.G.L.***********************************************
Auteur ....... : Thierry Petetin
Créé le ...... : 21/01/2005
Description .. : Adapte les combos des ParamSoc lié à la pièce
Description .. : d'origine
*****************************************************************}
{$IFNDEF EAGLSERVER}
Procedure GereParamSocPieceOrigine(FF: TForm);

  procedure SetEmpty(CC: TControl);
  begin
    if Assigned(CC) then
    begin
      THValComboBox(CC).Items.Insert(0, TraduireMemoire('< Aucune >'));
      THValComboBox(CC).Values.Insert(0, '');
    end;
  end;

begin
  if ExisteNam('SO_PIECEORIGINEACH', FF) then
    SetEmpty(GetFromNam('SO_PIECEORIGINEACH', FF));
  if ExisteNam('SO_PIECEORIGINEVTE', FF) then
    SetEmpty(GetFromNam('SO_PIECEORIGINEVTE', FF));
end;
{$ENDIF EAGLSERVER}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/06/2007
Modifié le ... : __/__/____
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
function TraitePlanDeRevision ( CC : TControl ) : Boolean;
var FF : TForm;
    vStPlanRevision : string;
begin
  Result := True;
  FF := GetLaForm(CC) ;

  if ExisteNam('SO_CPPLANREVISION', FF) then
  begin
{$IFNDEF CMPGIS35}
    vStPlanRevision := TEdit(GetFromNam('SO_CPPLANREVISION', FF)).Text;
    if vStPlanRevision <> '' then
    begin
      if vStPlanRevision <> VH^.Revision.Plan then
        Result := AlignementPlanRevision(vStPlanRevision, True);
    end
    else
    begin
      if VH^.Revision.Plan <> '' then
        Result := SupprimeRevision(True);
    end;
{$ENDIF}
  end;
end;
{$ENDIF}
{$ENDIF}
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/03/2005
Modifié le ... :   /  /
Description .. : lStNomLiasse peut avoir les suivantes :
Suite ........ : AUC ( Aucun )
Suite ........ : BIC ( CE1 )
Suite ........ : BNC ( CE3 )
Suite ........ : BA  ( CE5 )
Mots clefs ... :
*****************************************************************}
function CPEstLiasseModifiee : Boolean;
var lStNomLiasse : string;
    lSearchRec : TSearchRec;
    lInSearch : integer;
begin
  Result := False;
  lStNomLiasse := GetParamSoc('SO_CPCONTROLELIASSE');

  if lStNomLiasse <> 'AUC' then
  begin
    lInSearch := FindFirst( V_PGI.DosPath + '\*.*', faDirectory , lSearchRec);
    while (lInSearch = 0) do
    begin
      if lSearchRec.Attr = faDirectory then
      begin
        if ((lStNomLiasse = 'CE1') and (Pos('BIC', lSearchRec.Name) > 0)) OR
           ((lStNomLiasse = 'CE3') and (Pos('BNC', lSearchRec.Name) > 0)) OR
           ((lStNomLiasse = 'CE5') and (Pos('BA',  lSearchRec.Name) > 0)) then
        begin
          if FileExists(V_PGI.DosPath + '\' + lSearchRec.Name +'\*.mdb') then
          begin
            Result := True;
          end;
        end;
      end;
      lInSearch := FindNext( lSearchRec );
    end;
    SysUtils.FindClose( lSearchRec );
  end;
end;

{$IFNDEF EAGLSERVER}
{$IFDEF GESCOM}
procedure CdeOuvCacheFonctions(FF: TForm);
var
  cb: TCheckBox;
begin
  if not ExisteNam(SOCdeOuv, FF) then
    Exit; { Pas sur la bonne page }
  cb := TCheckBox(GetFromNam(SOCdeOuv,FF));
  if not Assigned(cb) then
    Exit;
  SetEna('SO_GCCDEOUVLISTESAISIE', FF, cb.Checked);
  SetEna('SO_GCCDEOUVLISTESAISIEACH', FF, cb.Checked);
  SetEna('SO_GCCDEOUVMULTIARTICLE', FF, cb.Checked);
  SetEna(SOCdeOuvAutoTransfo, FF, cb.Checked);
  SetEna('SO_GCCDEOUVTRANSFOPRE', FF, cb.Checked);
  SetEna('SO_GCCDEOUVELIMINELIGNE0', FF, cb.Checked);
end;
{$ENDIF GESCOM}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF GESCOM}
procedure CdeOuvTraiteAutoTransfo(CC: TControl);
begin
  if Assigned(CC) and (CC is TCheckBox) and TCheckBox(CC).Checked then
  begin
    if ExisteSQL('SELECT 1 FROM WPARAM WHERE WPA_CODEPARAM  = "' + CdeOuvAutoTransfo  + '" AND WPA_UTILISATEUR="'+V_PGI.User+'"')
       and (PGIAsk('Voulez-vous réactiver l''assistant de transformation automatique pour l''utilisateur ?', CdeOuvGetLibellePiece(CdeOuvGetCode(gtcOuverte))) = mrYes) then
    begin
      { Suppression des wParam liés à l'autoTransfo }
      ExecuteSql('DELETE FROM WPARAM WHERE WPA_CODEPARAM  = "' + CdeOuvAutoTransfo  + '" AND WPA_UTILISATEUR="'+V_PGI.User+'"');
    end;
  end;
end;
{$ENDIF GESCOM}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Procedure InitGPOFraisAvances( CC : TControl ) ;
Var
  FF        : TForm;
  Sql       : string;
  TobGPO	  : Tob;
  iGPO      : integer;
begin
  FF := GetLaForm(CC);
  if Assigned(CC) and (ExisteNam('SO_FRAISAVANCES', FF)) and (TCheckBox(GetFromNam('SO_FRAISAVANCES',FF)).Checked) then
  begin
    TobGPO := Tob.Create('GPO', nil, -1);
    try
      Sql:= 'SELECT * FROM PORT WHERE GPO_TYPEFRAIS="" OR GPO_VENTPIECE="" OR GPO_VENTCOMPTA="" OR GPO_REPARTITION="" OR GPO_FACTURABLE=""';
      if TobGPO.LoadDetailDBFromSQL('PORT', Sql) then
      begin
        For iGPO:=0 to TobGPO.Detail.Count-1 do
        begin
          if TobGPO.Detail[iGPO].GetString('GPO_TYPEFRAIS')  ='' then TobGPO.Detail[iGPO].SetString('GPO_TYPEFRAIS'  , '501'); // Frais annexe
          if TobGPO.Detail[iGPO].GetString('GPO_VENTPIECE')  ='' then TobGPO.Detail[iGPO].SetString('GPO_VENTPIECE'  , 'A');
          if TobGPO.Detail[iGPO].GetString('GPO_VENTCOMPTA') ='' then TobGPO.Detail[iGPO].SetString('GPO_VENTCOMPTA' , 'P');
          if TobGPO.Detail[iGPO].GetString('GPO_REPARTITION')='' then TobGPO.Detail[iGPO].SetString('GPO_REPARTITION', 'Q');
          if TobGPO.Detail[iGPO].GetBoolean('GPO_FACTURABLE')=False then TobGPO.Detail[iGPO].SetBoolean('GPO_FACTURABLE' , True);
        end;
      { UpDateRecord }
        TobGPO.AddChampSupValeur('IKC'    , 'M', false);
        TobGPO.AddChampSupValeur('Error'  , '' , false);

        TobGPO.UpdateDB;
      end;
    finally
      TobGPO.Free;
    end;
  end;
end;
{$ENDIF EAGLSERVER}

{$IFDEF GPAO}
{$IFNDEF EAGLSERVER}
Procedure InitWLBContexteOrigine( CC : TControl ) ;
Var
  FF        : TForm;
  Sql: string;
begin
  FF := GetLaForm(CC);
  if Assigned(CC) and (ExisteNam('SO_AFFAIRESINDUS', FF)) and (TCheckBox(GetFromNam('SO_AFFAIRESINDUS',FF)).Checked) then
  begin
    if not ExisteSql('SELECT 1 FROM WAFCONTEXTE WHERE WAC_CONTEXTE="ORIGINE"') then
    begin
      Sql :='INSERT INTO WAFCONTEXTE'
            + '(WAC_CONTEXTE,WAC_LIBELLE,WAC_DATECREATION,WAC_DATEMODIF,WAC_CREATEUR,WAC_UTILISATEUR,WAC_SOCIETE)'
            + ' VALUES ("ORIGINE","'+TraduireMemoire('Situation temps réel')+'","'+UsDateTime(Now)+'","'+UsDateTime(Now)+'","000","000","'+GetParamSOc('SO_SOCIETE')+'")';
      Executesql(sql);
    end;
  end;
end;
{$ENDIF EAGLSERVER}
{$ENDIF GPAO}

{$IFDEF GPAO}
{$IFDEF AFFAIRE}
{$IFNDEF PGIMAJVER}
{$IFNDEF EAGLSERVER}
Procedure InitParamAffairesIndustrielles( CC : TControl ) ;
Var
  FF        : TForm;
begin
  FF := GetLaForm(CC);
  if Assigned(CC) and (ExisteNam('SO_AFFAIRESINDUS', FF)) and (TCheckBox(GetFromNam('SO_AFFAIRESINDUS',FF)).Checked) then
  begin
//    {Mouvements exceptionnels multi-articles doit être activé}
//    SetParamsoc('SO_MODEMVTEX',true);

    wInitParamPrixDeRevientAFF;
    wInitWRRRubriqueNA;
    wInitWRSSectionNA;
  end;
end;
{$ENDIF !EAGLSERVER}
{$ENDIF !PGIMAJVER}
{$ENDIF AFFAIRE}
{$ENDIF GPAO}

{$IFNDEF EAGLSERVER}
Procedure MajParamCdeMarche( CC : TControl ) ;
begin
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Function VerifModeleWord( CC : TControl ) : Boolean;
Var FF : TForm ;
    ModeleWord : TEdit;
begin
  Result:=True;
  FF:=GetLaForm(CC) ;
  if ExisteNam('SO_GCMODELEWORD',FF) then
  begin
  ModeleWord:=TEdit(GetFromNam('SO_GCMODELEWORD',FF));
  if (ModeleWord.Text <> '') and (ModeleWord.Text <> GetParamSocSecur ('SO_GCMODELEWORD','')) and ((not IsNumeric(ModeleWord.Text)) or (not ExisteSQL ('SELECT ADE_LIBELLE FROM AFDOCEXTERNE WHERE ADE_DOCEXETAT = "UTI" AND ADE_DOCEXNATURE = "GCO" AND ADE_DOCEXTYPE = "WOR" AND ADE_DOCEXCODE = '+ ModeleWord.Text))) then
    begin
    HShowMessage('30;Société;Le modèle WORD n''existe pas.;W;O;O;O;','','') ;
    if ModeleWord.CanFocus then ModeleWord.SetFocus ;
    Result:=False;
    Exit;
    end;
  end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure VerifGestYplanningGRC( CC : TControl );
Var FF : TForm ;
    Exist : Boolean;
begin
  FF:=GetLaForm(CC) ;
  if ExisteNam('SO_YPLGRC',FF) then
  begin
    Exist := False;
    if Pos('RAI', THMultiValComboBox(GetFromNam('SO_YPLGA',FF)).Text) > 0 then
      Exist := True
    else if Pos('RAI', THMultiValComboBox(GetFromNam('SO_YPLPAIE',FF)).Text) > 0 then
      Exist := True
    else if Pos('RAI', THMultiValComboBox(GetFromNam('SO_YPLBUREAU',FF)).Text) > 0 then
      Exist := True;
    if Exist = False then // Pas de visu des actions GRC
      begin
      if Pos('RAI', GetParamSocSecur('SO_YPLGA', '')) > 0 then
        Exist := True
      else if Pos('RAI', GetParamSocSecur('SO_YPLPAIE', '')) > 0 then
        Exist := True
      else if Pos('RAI', GetParamSocSecur('SO_YPLBUREAU', '')) > 0 then
        Exist := True;
      if Exist = True then // il y avait visu des actions GRC
        begin
          ExecuteSql ('UPDATE PARACTIONS SET RPA_PLANIFIABLE = "-",'+
            'RPA_DATEMODIF="'+UsDateTime(Now)+'", '+
            'RPA_UTILISATEUR="'+V_PGI.User+'" WHERE RPA_PLANIFIABLE = "X"') ;
        end;
      end;
  end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure MajTabletteGcTypeFourn( CC : TControl );
var
  FF: TForm;
	Q : TQuery;
begin
  FF := GetLaForm(CC);
  if Assigned(CC) and (ExisteNam('SO_GCTRANSPORTEURS', FF)) and (TCheckBox(GetFromNam('SO_GCTRANSPORTEURS',FF)).Checked) then
  begin
    Q:=OpenSQL('SELECT * FROM COMMUN WHERE CO_TYPE="GFT" AND CO_CODE="TRA"',false);
    If Q.EOF Then
    begin
      Q.Insert ;
      Q.FindField('CO_TYPE').AsString		:='GFT';
      Q.FindField('CO_CODE').AsString		:='TRA' ;
      Q.FindField('CO_LIBELLE').AsString:='Transporteur';
      Q.FindField('CO_ABREGE').AsString	:='Transporteur' ;
      Q.FindField('CO_LIBRE').AsString	:='' ;
      Q.Post ;
    end;
    Ferme(Q) ;
  end;
end;
{$ENDIF EAGLSERVER}

// DEBUT CCMX-CEGID TMDR Dev N° 4722 {Evolution}
{$IFNDEF EAGLSERVER}
procedure MajClassificationTMDR(CC : TControl);
var
  FF: TForm;
  sSQL :String;
begin
  FF := GetLaForm(CC);
  if Assigned(CC) and (ExisteNam('SO_GCTRANSPORTEURS', FF)) and (TCheckBox(GetFromNam('SO_GCTRANSPORTEURS',FF)).Checked) then
  begin
    if not ExisteSQL ('Select YDL_CODEHDTLINK from YDATATYPELINKS WHERE YDL_CODEHDTLINK="TMDRCLASSIF"') then
    Begin
      ExecuteSQL('INSERT INTO YDATATYPELINKS (YDL_CODEHDTLINK,YDL_LIBELLE,YDL_MDATATYPE,YDL_SDATATYPE,YDL_TYPELINK,YDL_PREDEFINI,YDL_LCOMMUNE)'
        + ' VALUES ("TMDRCLASSIF","Classification classe danger","GCTMDRCLASSE","GCTMDRCLASSEMENT","EXC","CEG","-")');
    End;
    If Not ExisteSQL('SELECT * FROM YDATATYPETREES WHERE YDT_CODEHDTLINK="TMDRCLASSIF"') Then
    begin
    sSQL := 'INSERT INTO YDATATYPETREES (YDT_CODEHDTLINK,YDT_MCODE,YDT_SCODE,YDT_PREDEFINI)'
          + ' SELECT "TMDRCLASSIF",CO_LIBELLE,CO_CODE,"CEG" FROM COMMUN WHERE CO_TYPE="MD1"';
    executeSQL(sSQL);
    end;
  end;
end;
{$ENDIF EAGLSERVER}
// FIN CCMX-CEGID TMDR Dev N° 4722 {Evolution}

{$IFNDEF EAGLSERVER}
{$IFDEF GCGC}
procedure SetPlusUniteParDefaut(CC: TControl);
var
  F: TForm;
  QualifGA, UniteGA: TControl;
begin
  F := GetLaForm(CC);
  if not ExisteNam('SO_GCDEFUNITEGA', F) then
    Exit; { Pas sur la bonne page }
  UniteGA := GetFromNam('SO_GCDEFUNITEGA', F);
  if Assigned(UniteGA) then
  begin
    QualifGA := GetFromNam('SO_GCDEFQUALIFUNITEGA', F);
    if Assigned(QualifGA) then
    begin
      THValComboBox(QualifGA).Vide := (GetParamSocGcGestUniteMode = '003');
      THValComboBox(QualifGA).VideString := TraduireMemoire('<<Aucun>>');
      THValComboBox(QualifGA).Plus := ' AND (CC_CODE <> "AUC")';
      if THValComboBox(QualifGA).Value <> '' then
        THValComboBox(UniteGA).Plus := '(GME_QUALIFMESURE="' + THValComboBox(QualifGA).Value + '")'
      else
        THValComboBox(UniteGA).Plus := '';
    end;
    SetEna('SO_GCDEFQUALIFUNITEGA', GetLaForm(CC), GetParamSocGcGestUniteMode <> '001');
    SetEna('SO_GCDEFUNITEGA', GetLaForm(CC), True);
  end;
end;
{$ENDIF GCGC}
{$ENDIF EAGLSERVER}

//GP_BUG810_TP WPLANLIVR_20071008 >>>
{$IFNDEF EAGLSERVER}
{$IFDEF GCGC}
procedure SetWPlanLivr(CC: TControl);
var
  F: TForm;

  procedure SetPlus(const Name: String);
  var
    C: TControl;
  begin
    C := GetFromNam(Name, F);
    if Assigned(C) and (C is THMultiValComboBox) then
      THMultiValComboBox(C).Plus :=' AND (GPP_NATUREPIECEG NOT IN ("AFF", "APR", "AVC", "AVS", "CCH", "CCR"'
                                                               + ', "DE", "FFA", "FFO", "FPR", "LCR", "PAF", "PRO"))'
  end;

begin
  F := GetLaForm(CC);
  //GP_20071129_TP_GP14465 >>>
  if ExisteNam('SO_WPLANLIVGRP1', F) then
    SetInvi('SO_WPLANLIVGRP1', F);
  if ExisteNam('SCO_WPLANLIVGRPLIB', F) then
    SetInvi('SCO_WPLANLIVGRPLIB', F);
  //GP_20071129_TP_GP14465 <<<
  if ExisteNam('SO_WPLANLIVRNAT1', F) then
    SetPlus('SO_WPLANLIVRNAT1');
  if ExisteNam('SO_WPLANLIVRNAT2', F) then
    SetPlus('SO_WPLANLIVRNAT2');
end;
{$ENDIF GCGC}
{$ENDIF EAGLSERVER}

//GP_BUG810_TP WPLANLIVR 20071008 <<<
{$IFNDEF EAGLSERVER}
procedure GereZonesVersions (CC : TControl);
Var FF : TForm ;
BEGIN
  FF:=GetLaForm(CC) ;
  if ExisteNam('SCO_LIBSTATUT1',FF) then
     TLabel(GetFromNam('SCO_LIBSTATUT1',FF)).Caption:=RechDom('WSTATUTVERSION','002',False);
  if ExisteNam('SCO_LIBSTATUT2',FF) then
     TLabel(GetFromNam('SCO_LIBSTATUT2',FF)).Caption:=RechDom('WSTATUTVERSION','002',False);
  if ExisteNam('SCO_LIBSTATUT1B',FF) then
     TLabel(GetFromNam('SCO_LIBSTATUT1B',FF)).Caption:=RechDom('WSTATUTVERSION','003',False);
  if ExisteNam('SCO_LIBSTATUT2B',FF) then
     TLabel(GetFromNam('SCO_LIBSTATUT2B',FF)).Caption:=RechDom('WSTATUTVERSION','003',False);
  if ExisteNam('SO_GESTSTATUTUNIQUE1', FF) then
  begin
    if TCheckBox(GetFromNam('SO_GESTSTATUTUNIQUE1',FF)).Checked = false then
        begin
        TComboBox(GetFromNam('SO_STATUTBASCULE1',FF)).Enabled:=False ;
        TComboBox(GetFromNam('LSO_STATUTBASCULE1',FF)).Enabled:=False ;
        TLabel(GetFromNam('SCO_REMPLACER1',FF)).Enabled:=False ;
        TLabel(GetFromNam('SCO_LIBSTATUT2',FF)).Enabled:=False ;
        end
    else
        begin
        TComboBox(GetFromNam('SO_STATUTBASCULE1',FF)).Enabled:=true ;
        TComboBox(GetFromNam('LSO_STATUTBASCULE1',FF)).Enabled:=true ;
        TLabel(GetFromNam('SCO_REMPLACER1',FF)).Enabled:=true ;
        TLabel(GetFromNam('SCO_LIBSTATUT2',FF)).Enabled:=true ;
        end ;
  end;
  if ExisteNam('SO_GESTSTATUTUNIQUE2', FF) then
  begin
    if TCheckBox(GetFromNam('SO_GESTSTATUTUNIQUE2',FF)).Checked = false then
        begin
        TComboBox(GetFromNam('SO_STATUTBASCULE2',FF)).Enabled:=False ;
        TComboBox(GetFromNam('LSO_STATUTBASCULE2',FF)).Enabled:=False ;
        TLabel(GetFromNam('SCO_REMPLACER2',FF)).Enabled:=False ;
        TLabel(GetFromNam('SCO_LIBSTATUT2B',FF)).Enabled:=False ;
        end
    else
        begin
        TComboBox(GetFromNam('SO_STATUTBASCULE2',FF)).Enabled:=true ;
        TComboBox(GetFromNam('LSO_STATUTBASCULE2',FF)).Enabled:=true ;
        TLabel(GetFromNam('SCO_REMPLACER2',FF)).Enabled:=true ;
        TLabel(GetFromNam('SCO_LIBSTATUT2B',FF)).Enabled:=true ;
        end;
  end;
end;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF GCGC}
procedure SetGCGestUniteModeState(CC: TControl);
var
  F: TForm;
begin
  F := GetLaForm(CC);
  if not ExisteNam('SO_GCGESTUNITEMODE', F) then
    Exit { Pas sur la bonne page }
  else
    SetEna('SO_GCGESTUNITEMODE', GetLaForm(CC), False);
end;
{$ENDIF GCGC}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF GCGC}
procedure GCChangeGestUniteMode(CC: TControl);
var
  F: TForm;
  T: Tob;
  Combo: ThValComboBox;

  function GetLastNumEvent: integer;
  var
    TT: Tob;
  begin
    Result := 0;
    TT := Tob.Create('JNALEVENT', nil, -1);
    try
      if TT.LoadDetailFromSQL('SELECT MAX(GEV_NUMEVENT) AS LASTNUMEVENT FROM JNALEVENT') then
        Result := TT.Detail[0].GetInteger('LASTNUMEVENT');
    finally
      TT.Free;
    end;
    Inc(Result);
  end;

begin
  F := GetLaForm(CC);
  if not ExisteNam('SO_GCGESTUNITEMODE', F) then
    Exit { Pas sur la bonne page }
  else
  begin
    Combo := ThValComboBox(GetFromNam('SO_GCGESTUNITEMODE', F));
    if Assigned(Combo) and (Combo.Value <> GetParamSocGcGestUniteMode) then
    begin
      PGIInfo('Le mode de gestion des unités doit être modifié uniquement par le menu : ' + #13
              + '''Administration\Maintenance\Mise à jour\Unités''' + #13 + #13
              + 'Ceci à fin de préserver l''intégrité des données de la base.', 'ATTENTION !!!');
      { Mouchard changement manuel dans le journal }
      T := Tob.Create('JNALEVENT', nil, -1);
      try
        T.SetString('GEV_TYPEEVENT', 'UNI');
        T.SetDateTime('GEV_DATEEVENT', Date);
        T.SetString('GEV_UTILISATEUR', V_PGI.User);
        T.SetInteger('GEV_NUMEVENT', GetLastNumEvent);
        T.SetString('GEV_ETATEVENT', 'ERR');
        T.SetString('GEV_BLOCNOTE', Format(TraduireMemoire('Changement manuel du mode de gestion des unités de %s à %s'), [Combo.Value, GetParamSocGcGestUniteMode]));
        T.InsertDB(nil);
      finally
        T.Free;
      end;
    end;
  end;
end;
{$ENDIF GCGC}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF GIGI}
{***********A.G.L.***********************************************
Auteur  ...... : Philippe Chevron
Créé le ...... : 22/11/2005
Modifié le ... : 22/11/2005
Description .. : A la validation de la grille des préférences affaires :
Suite ........ : contrôles de saisie des paramètres de personnalisation des
Suite ........ : paramètres société.
Suite ........ : (Fonction appelée par SauvePageSoc)
Mots clefs ... :
*****************************************************************}

Function VerifAfPerso ( CC : TControl ) : boolean ;
Var FF : TForm ;
    CbPersoChoix : TCheckBox ;
    CVCBPersoParam1 : THValComboBox;
    CVCBPersoParam2 : THValComboBox;
    CVCBPersoParam3 : THValComboBox;
BEGIN
  Result:=True ;
  FF:=GetLaForm(CC) ;

  if ExisteNam('SO_AFPERSOCHOIX',FF) then
    begin

      CbPersoChoix:=  TCheckBox(GetFromNam('SO_AFPERSOCHOIX',FF));
      CVCBPersoParam1 := THValComboBox (GetFromNam ('SO_AFPERSOPARAM1', FF));
      CVCBPersoParam2 := THValComboBox (GetFromNam ('SO_AFPERSOPARAM2', FF));
      CVCBPersoParam3 := THValComboBox (GetFromNam ('SO_AFPERSOPARAM3', FF));

      if CbPersoChoix.checked = true then
        begin
          if  (CVCBPersoParam1.value='')
          and (CVCBPersoParam2.value='')
          and (CVCBPersoParam3.value='') then
            begin
              //paramètres non renseignés
              Result:=False ;
              PgiError(TraduireMemoire('Il faut renseigner au moins l''un des paramètres de personnalisation.'));
              Exit;
            end;
          //2006/03
          if  (CVCBPersoParam1.value='AB_') then
            begin
              //param1 = "Aucun" interdit
              Result:=False ;
              PgiError(TraduireMemoire('Aucun n''est pas autorisé pour le paramètre 1.'));
              Exit;
            end;

          if  Not (CVCBPersoParam1.value='')
          and Not (CVCBPersoParam2.value='') then
            begin
              if (CVCBPersoParam1.value=CVCBPersoParam2.value) then
                begin
                  //param1 et param2 identiques
                  Result:=False ;
                  PgiError(TraduireMemoire('Les paramètres de personnalisation 1 et 2 ne doivent pas être identiques.'));
                  Exit;
                end
            end;
          if  Not (CVCBPersoParam1.value='')
          and Not (CVCBPersoParam3.value='') then
            begin
              if (CVCBPersoParam1.value=CVCBPersoParam3.value) then
                begin
                  //param1 et param3 identiques
                  Result:=False ;
                  PgiError(TraduireMemoire('Les paramètres de personnalisation 1 et 3 ne doivent pas être identiques.'));
                  Exit;
                end
            end;
          if  Not (CVCBPersoParam2.value='')
          and Not (CVCBPersoParam3.value='') then
            begin
              if ((CVCBPersoParam2.value=CVCBPersoParam3.value)
              and (CVCBPersoParam2.value<>'AB_')) then
                begin
                  //param2 et param3 identiques et différents de "Aucun"
                  Result:=False ;
                  PgiError(TraduireMemoire('Les paramètres de personnalisation 2 et 3 ne doivent pas être identiques.'));
                  Exit;
                end
            end;
        end
      else
        begin
          // choix personnalisation non coché
          Result := DecochagePersoParamsocChoix( CC ) ;
        end;
    end;
END ;
{$ENDIF GIGI}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF GIGI}
{***********A.G.L.***********************************************
Auteur  ...... : Philippe Chevron PCH
Créé le ...... : 12/12/2005
Modifié le ... :
Description .. : contrôles de saisie de la personnalisation de la facture liquidative
Suite ........ : (Fonction appelée par InterfaceSoc)
Mots clefs ... :
*****************************************************************}

Procedure CtlFactureLiquidative ( CC : TControl );
Var FF : TForm ;
    CBGereLiquide : TCheckBox ;
BEGIN
  FF:=GetLaForm(CC) ;

  if ExisteNam('SO_AFGERELIQUIDE',FF) then
  begin
    CBGereLiquide:= TCheckBox(GetFromNam('SO_AFGERELIQUIDE',FF));
    if CBGereLiquide.checked = true then
    begin
      // cochage gestion facture liquidative
      SetEna('SO_AFLIBLIQFACTAF', FF, True);
      SetEna('LSO_AFLIBLIQFACTAF', FF, True);
      SetEna('SO_AFTYPGENFACTLIQ', FF, True);
      SetEna('LSO_AFTYPGENFACTLIQ', FF, True);
      SetEna('SO_AFREPACTFACTLIQ', FF, True);
      SetEna('LSO_AFREPACTFACTLIQ', FF, True);
      SetEna('SO_AFPROFACFACTLIQ', FF, True);
      SetEna('LSO_AFPROFACFACTLIQ', FF, True);
    end
    else
    begin
      // décochage gestion facture liquidative
      SetEna('SO_AFLIBLIQFACTAF', FF, False);
      SetEna('LSO_AFLIBLIQFACTAF', FF, False);
      SetEna('SO_AFTYPGENFACTLIQ', FF, False);
      SetEna('LSO_AFTYPGENFACTLIQ', FF, False);
      SetEna('SO_AFREPACTFACTLIQ', FF, False);
      SetEna('LSO_AFREPACTFACTLIQ', FF, False);
      SetEna('SO_AFPROFACFACTLIQ', FF, False);
      SetEna('LSO_AFPROFACFACTLIQ', FF, False);
    end;
  end;
END ;
{$ENDIF GIGI}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF GIGI}
{***********A.G.L.***********************************************
Auteur  ...... : Philippe Chevron
Créé le ...... : 22/11/2005
Modifié le ... : 22/11/2005
Description .. : Contrôles de saisie du CB de choix de la personnalisation
Suite ........ : des paramètres société.
Suite ........ : (Fonction appelée par InterfaceSoc)
Mots clefs ... :
*****************************************************************}

Procedure CtlPersoParamsocChoix( CC : TControl ) ;
Var FF : TForm ;
    CbPersoChoix    : TCheckBox;
    CVCBPersoParam1 : THValComboBox;
    CVCBPersoParam2 : THValComboBox;
    CVCBPersoParam3 : THValComboBox;

BEGIN
  FF:=GetLaForm(CC) ;
  CbPersoChoix:=TCheckBox(CC);
  CVCBPersoParam1 := THValComboBox (GetFromNam ('SO_AFPERSOPARAM1', FF));
  CVCBPersoParam2 := THValComboBox (GetFromNam ('SO_AFPERSOPARAM2', FF));
  CVCBPersoParam3 := THValComboBox (GetFromNam ('SO_AFPERSOPARAM3', FF));

  if CbPersoChoix.Checked then
    begin
      // choix personnalisation coché
      if ExisteSQL('SELECT APA_NOMPARAMSOC FROM AFPARAMSOC' ) then
        begin
          // saisie de personnalisation déjà effectuée
          // => fermer la saisie sur les paramètres
          SetEna('SO_AFPERSOPARAM1', FF, False);
          SetEna('SO_AFPERSOPARAM2', FF, False);
          SetEna('SO_AFPERSOPARAM3', FF, False);
          SetEna('LSO_AFPERSOPARAM1', FF, False);
          SetEna('LSO_AFPERSOPARAM2', FF, False);
          SetEna('LSO_AFPERSOPARAM3', FF, False);
        end
      else
        begin
          //ouvrir la saisie sur PARAM1
          CVCBPersoParam1.Enabled:=True ;
          SetEna('LSO_AFPERSOPARAM1', FF, True);

          if (CVCBPersoParam1.value='') then
            begin
              // PARAM1 est vide, rendre inaccessibles PARAM2 et PARAM3
              CVCBPersoParam2.Enabled:=False ;
              CVCBPersoParam3.Enabled:=False ;
              CVCBPersoParam2.Value  := '';
              CVCBPersoParam3.Value  := '';
              SetEna('LSO_AFPERSOPARAM2', FF, False);
              SetEna('LSO_AFPERSOPARAM3', FF, False);
            end
          else
            begin
              // PARAM1 est renseigné, ouvrir la saisie sur PARAM2
              CVCBPersoParam2.Enabled:=True ;
              SetEna('LSO_AFPERSOPARAM2', FF, True);
              if ((CVCBPersoParam2.value='')
              or  (CVCBPersoParam2.value='AB_')) then
                begin
                  // PARAM2 est vide ou "Aucun", rendre inaccessible PARAM3
                  TComboBox(GetFromNam('SO_AFPERSOPARAM3',FF)).Enabled:=False ;
                  CVCBPersoParam3.Value := '';
                end
              else
                begin
                  // PARAM2 est renseigné, ouvrir la saisie sur PARAM3
                  CVCBPersoParam3.Enabled:=True ;
                end
            end
        end
    end
  else
    begin
      if DecochagePersoParamsocChoix( CC ) then;
    end;
END;
{$ENDIF GIGI}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF GIGI}
{***********A.G.L.***********************************************
Auteur  ...... : Philippe Chevron
Créé le ...... : 22/11/2005
Modifié le ... :   /  /
Description .. : Contrôles de saisie de la combo du premier paramètre de la
Suite ........ : personnalisation des paramètres société.
Suite ........ : (Fonction appelée par InterfaceSoc)
Mots clefs ... :
*****************************************************************}
Procedure CtlPersoParamsocPar1( CC : TControl ) ;
Var FF : TForm ;
    CVCBPersoParam1 : THValComboBox;
    CVCBPersoParam2 : THValComboBox;
    CVCBPersoParam3 : THValComboBox;
BEGIN
  FF:=GetLaForm(CC) ;
  CVCBPersoParam1 := THValComboBox(CC);
  CVCBPersoParam2 := THValComboBox (GetFromNam ('SO_AFPERSOPARAM2', FF));
  CVCBPersoParam3 := THValComboBox (GetFromNam ('SO_AFPERSOPARAM3', FF));

  if ((CVCBPersoParam1.value = '')
  or  (CVCBPersoParam1.value = 'AB_')) then
    begin
      // PARAM1 est vide ou "Aucun", rendre inaccessibles PARAM2 et PARAM3
      CVCBPersoParam2.Enabled:=False ;
      CVCBPersoParam3.Enabled:=False ;
      CVCBPersoParam2.Value  := '';
      CVCBPersoParam3.Value  := '';
      SetEna('LSO_AFPERSOPARAM2', FF, False);
      SetEna('LSO_AFPERSOPARAM3', FF, False);
    end
  else
    begin
      // ouvrir la saisie sur PARAM2
      CVCBPersoParam2.Enabled:=True ;
      SetEna('LSO_AFPERSOPARAM2', FF, True);
      if ((CVCBPersoParam2.value='')
      or  (CVCBPersoParam2.value = 'AB_')) then
        begin
          // PARAM2 est vide, rendre inaccessible PARAM3
          CVCBPersoParam3.Enabled:=False ;
          CVCBPersoParam3.Value := '';
          SetEna('LSO_AFPERSOPARAM3', FF, False);
        end
      else
        begin
          // PARAM2 est renseigné, ouvrir la saisie sur PARAM3
          CVCBPersoParam3.Enabled:=True ;
          SetEna('LSO_AFPERSOPARAM3', FF, True);
        end
    end;
END;
{$ENDIF GIGI}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF GIGI}
{***********A.G.L.***********************************************
Auteur  ...... : Philippe Chevron
Créé le ...... : 22/11/2005
Modifié le ... :   /  /
Description .. : Contrôles de saisie de la combo du deuxième paramètre de
Suite ........ : la personnalisation des paramètres société.
Suite ........ : (Fonction appelée par InterfaceSoc)
Mots clefs ... :
*****************************************************************}
Procedure CtlPersoParamsocPar2( CC : TControl ) ;
Var FF : TForm ;
    CVCBPersoParam2 : THValComboBox;
    CVCBPersoParam3 : THValComboBox;

BEGIN
  FF:=GetLaForm(CC) ;
  CVCBPersoParam2:=THValComboBox(CC);
  CVCBPersoParam3 := THValComboBox (GetFromNam ('SO_AFPERSOPARAM3', FF));

  if ((CVCBPersoParam2.value='')
  or  (CVCBPersoParam2.value = 'AB_')) then
    begin
      // PARAM2 est vide, rendre inaccessible PARAM3
      CVCBPersoParam3.Enabled:=False ;
      CVCBPersoParam3.Value  := '';
      SetEna('LSO_AFPERSOPARAM3', FF, False);
    end
  else
    begin
      // PARAM2 est renseigné, ouvrir la saisie sur PARAM3
      CVCBPersoParam3.Enabled := True ;
      SetEna('LSO_AFPERSOPARAM3', FF, True);
    end

END;
{$ENDIF GIGI}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF GIGI}
{***********A.G.L.***********************************************
Auteur  ...... : Philippe Chevron
Créé le ...... : 04/01/2006
Modifié le ... :
Description .. : Gestion du décochage du CB de choix de la personnalisation
Mots clefs ... :
*****************************************************************}

Function DecochagePersoParamsocChoix( CC : TControl ) : boolean  ;
Var FF : TForm ;
    CbPersoChoix    : TCheckBox;
    CVCBPersoParam1 : THValComboBox;
    CVCBPersoParam2 : THValComboBox;
    CVCBPersoParam3 : THValComboBox;
BEGIN
  Result:=True ;
  FF:=GetLaForm(CC) ;
  CbPersoChoix    := TCheckBox(GetFromNam('SO_AFPERSOCHOIX',FF));
  CVCBPersoParam1 := THValComboBox (GetFromNam ('SO_AFPERSOPARAM1', FF));
  CVCBPersoParam2 := THValComboBox (GetFromNam ('SO_AFPERSOPARAM2', FF));
  CVCBPersoParam3 := THValComboBox (GetFromNam ('SO_AFPERSOPARAM3', FF));

  // choix personnalisation non coché
  if ExisteSQL('SELECT APA_NOMPARAMSOC FROM AFPARAMSOC' ) then
    begin
      // => personnalisation de paramètres soc déjà faite
      if PGIAsk(TraduireMemoire('Confirmez-vous la suppression de la personnalisation des paramètres société'),
                TraduireMemoire('Paramètres société')) = mrYes then
        begin
          if PGIAsk('Attention ! Vous avez demandé la suppression des paramètres société personnalisés.'
                    + #13 + 'Voulez-vous tout de même les conserver ?',
                    'Paramètres société') = mrYes then
            begin
              CbPersoChoix.Checked := True;
              Result:=False ;
            end;
        end
      else
        begin
          CbPersoChoix.Checked := True;
          Result:=False ;
        end;
      if Result = True then
        begin
          // => Supprimer la personnalisation de paramètres soc déjà faite
          ExecuteSQL('DELETE FROM AFPARAMSOC');
          CVCBPersoParam1.Value := '';
          CVCBPersoParam2.Value := '';
          CVCBPersoParam3.Value := '';
        end
    end
  else
    begin
      CVCBPersoParam1.Value := '';
      CVCBPersoParam2.Value := '';
      CVCBPersoParam3.Value := '';
    end;

  // => rendre inaccessibles les 3 PARAM
  CVCBPersoParam1.Enabled:=False ;
  SetEna('LSO_AFPERSOPARAM1', FF, False);
  CVCBPersoParam2.Enabled:=False ;
  SetEna('LSO_AFPERSOPARAM2', FF, False);
  CVCBPersoParam3.Enabled:=False ;
  SetEna('LSO_AFPERSOPARAM3', FF, False);

END;
{$ENDIF GIGI}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
Function VerifAfActivite ( CC : TControl ) : boolean ;
Var FF : TForm ;
    CbVisa : TCheckBox ;
BEGIN     //mcd 12/09/2006 13310
  Result:=True ;
  FF:=GetLaForm(CC) ;
  if ExisteNam('SO_AFDATEDEBUTACT',FF) then
    begin
      CbVisa:=  TCheckBox(GetFromNam('SO_AFVISAACTIVITE',FF));
      if Cbvisa.Checked= False then
        begin
        if ExisteSql('Select act_etatvisa from activite where act_etatvisa="ATT" and act_typeactivite="REA"') then
          begin
          PgiInfo ('Vous avez des lignes d''activité non visées, vous ne pourrez plus le faire.');
          end;
        end;
    end;
end;

procedure GereMasqueCrit( FF : TForm );
Var CBMasqueCrit1 : THValComboBox;
    CBMasqueCrit2 : THValComboBox;
    CBMasqueCrit3 : THValComboBox;
    lStVal1       : string ;
    lStVal2       : string ;
    lStVal3       : string ;
begin

  CBMasqueCrit1 := THValComboBox (GetFromNam ('SO_CPMASQUECRIT1', FF));
  CBMasqueCrit2 := THValComboBox (GetFromNam ('SO_CPMASQUECRIT2', FF));
  CBMasqueCrit3 := THValComboBox (GetFromNam ('SO_CPMASQUECRIT3', FF));

  // Si des masques de saisie ont été saisis, modif du paramétrage impossible
  if ExisteSQL('SELECT CMS_NUMERO FROM CMASQUESAISIE WHERE CMS_TYPE="SAI"') then
    begin
    CBMasqueCrit1.Enabled := False ;
    CBMasqueCrit2.Enabled := False ;
    CBMasqueCrit3.Enabled := False ;
    Exit ;
    end ;

  lStVal1 := CBMasqueCrit1.Value ;

  // Sur le 2ème crit, on ôte la valeur du 1er
  lStVal2 := CBMasqueCrit2.Value ;
  CBMasqueCrit2.Plus := ' AND CO_CODE<>"' + lStVal1 + '" ' ;
  CBMasqueCrit2.Refresh ;
  if lStVal2 = lStVal1 then
    lStVal2 := '000' ;
  CBMasqueCrit2.Value := lStVal2 ;
  if lStVal2 = '000' then
    begin
    CBMasqueCrit3.Value := '000' ;
    CBMasqueCrit3.Enabled := False ;
    end
  else
    begin
    // Sur le 3ème crit, on ôte la valeur des 2 1er
    lStVal3 := CBMasqueCrit3.Value ;
    CBMasqueCrit3.Enabled := True ;
    CBMasqueCrit3.Plus := ' AND CO_CODE NOT IN ("' + lStVal1 + '", "' + lStVal2 + '") ' ;
    CBMasqueCrit3.Refresh ;
    if (lStVal3 = lStVal1) or (lStVal3 = lStVal2) then
      CBMasqueCrit3.Value := '000' ;
    end ;

end ;
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF GCGC}
{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 24/11/2006
Modifié le ... :   /  /
Description .. : Gestion dynamique des paramètres SO_GCENSEIGNETAB
Suite ........ : et SO_GCENSEIGNECREAT.
Suite ........ : La création d'enseigne est impossible si la gestion des
Suite ........ : tablettes est désactivée
Mots clefs ... :
*****************************************************************}
procedure SetGereEnseigne (CC : TControl);
var
  F               : TForm;
  EnseigneTab     : TCheckBox;
  EnseigneCreat   : TCheckBox;

begin
  F               := GetLaForm(CC);
  if ExisteNam('SO_GCENSEIGNETAB', F) then
  begin
    EnseigneTab     := TCheckBox(GetFromNam('SO_GCENSEIGNETAB', F));
    EnseigneCreat   := TCheckBox(GetFromNam('SO_GCENSEIGNECREAT', F));

    if EnseigneTab.Checked then
    begin
      EnseigneCreat.Enabled     := True;
    end
    else
    begin
      EnseigneCreat.Enabled     := False;
      EnseigneCreat.Checked     := False;
    end;
  end;
End;


//FQ 10669 TJA
{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/10/2007
Modifié le ... :   /  /    
Description .. : Gestion dynamique des paramètres de ciblage automatique.
Mots clefs ... : 
*****************************************************************}
procedure SetGereCiblage (CC : TControl);
var
  F                     : TForm;
  CiblageAuto           : TCheckBox;
  PrefixeCiblage        : TcheckBox;
  CompteurCiblage       : THEdit;

begin
  F                     := GetLaForm(CC);
  if ExisteNam('SO_NUMCIBLAGEAUTO', F) then
  begin
    CiblageAuto         := TCheckBox(GetFromNam('SO_NUMCIBLAGEAUTO', F));
    PrefixeCiblage      := TCheckBox(GetFromNam('SO_PREFIXEAUTOCIBLAGE', F));
    CompteurCiblage     := THEdit(GetFromNam('SO_COMPTEURCIBLAGE', F));

    if CiblageAuto.Checked then
    begin
      CompteurCiblage.Enabled := True;
      PrefixeCiblage.Enabled  := True;
    end
    else
    begin
      CompteurCiblage.Enabled := False;
      PrefixeCiblage.Enabled  := False;
      PrefixeCiblage.Checked  := False;
    end;
  end;
end;


//FQ 10669 TJA
{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/10/2007
Modifié le ... :   /  /    
Description .. : Gestion dynamaique des paramètres de numérotation 
Suite ........ : automatique de l'opération 
Mots clefs ... : 
*****************************************************************}
procedure SetGereOperation (CC : TControl);
var
  F                     : TForm;
  OperAuto              : TCheckBox;
  PrefixeOper           : TCheckBox;
  CompteurOper          : THEdit;

begin
  F                     := GetLaForm(CC);
  if ExisteNam('SO_RTNUMOPERAUTO', F) then
  begin
    OperAuto            := TCheckBox(GetFromNam('SO_RTNUMOPERAUTO', F));
    PrefixeOper         := TCheckBox(GetFromNam('SO_RTPREFIXEAUTOOPERATION', F));
    CompteurOper        := THedit(GetFromNam('SO_RTCOMPTEUROPER', F));

    if OperAuto.Checked then
    begin
      CompteurOper.Enabled  := True;
      PrefixeOper.Enabled   := True;
    end
    else
    begin
      CompteurOper.Enabled  := False;
      PrefixeOper.Enabled   := False;
      prefixeOper.Checked   := False;
    end;
  end;
end;

{$ENDIF GCGC}
{$ENDIF EAGLSERVER}

//  BBI web services
{$IFNDEF EAGLSERVER}
{$IFDEF STK}
{***********A.G.L.***********************************************
Auteur  ...... : CVN
Créé le ...... : 06/04/2007
Modifié le ... :   /  /
Description .. : On controle si il existe un emplacement renseigné dans au
Suite ........ : moins une fiche dispo, si c'est la cas on ne permet pas la
Suite ........ : désactivation de la gestion des emplacements.
Mots clefs ... : GESTION EMPLACEMENT
*****************************************************************}
procedure CtrlDecochageGestionEmplacement(CC : TControl);
var
  FF : TForm;
begin
  FF := GetLaForm(CC);

  if ExisteNam('SO_GEREEMPLACEMENT', FF) then
  begin
    { Si décochage }
    if not ThCheckBox(GetFromNam('SO_GEREEMPLACEMENT', FF)).Checked then
    begin
      { Recherche s'il existe au moins une fiche dispo avec un emplacement }
      if ExisteSQL('SELECT 1 FROM DISPO WHERE GQ_EMPLACEMENT <> "" ' ) then
      begin
        PgiInfo ('Il y a au moins un emplacement renseigné sur une fiche de stock, vous ne pouvez pas désactiver la gestion des emplacements.');
        ThCheckBox(GetFromNam('SO_GEREEMPLACEMENT', FF)).Checked := true;
      end;
    end;
  end;

end;
{$ENDIF STK}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF STK}
{***********A.G.L.***********************************************
Auteur  ...... : CVN
Créé le ...... : 06/04/2007
Modifié le ... : 06/04/2007
Description .. : On controle si il existe au
Suite ........ : moins un article géré par n° de lot, si c'est la cas on ne permet pas la
Suite ........ : désactivation de la gestion des lots.
Mots clefs ... : GESTION LOTS
*****************************************************************}
procedure CtrlDecochageGestionLot(CC : TControl);
var
  FF : TForm;
begin
  FF := GetLaForm(CC);

  if ExisteNam('SO_GERELOT', FF) then
  begin
    { Si décochage }
    if not ThCheckBox(GetFromNam('SO_GERELOT', FF)).Checked then
    begin
      { Recherche s'il existe au moins un article géré par lot }
      if ExisteSQL('SELECT 1 FROM ARTICLE WHERE GA_LOT = "'+ wTrue + '"') then
      begin
        PgiInfo ('Il y a au moins un article géré par lot, vous ne pouvez pas désactiver la gestion des lots.');
        ThCheckBox(GetFromNam('SO_GERELOT', FF)).Checked := true;
      end;
    end;
  end;

end;
{$ENDIF STK}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF STK}
{***********A.G.L.***********************************************
Auteur  ...... : CVN
Créé le ...... : 06/04/2007
Modifié le ... : 06/04/2007
Description .. : On controle si il existe au
Suite ........ : moins un article géré en n° de série, si c'est la cas on ne permet pas la
Suite ........ : désactivation de la gestion des n° de série.
Mots clefs ... : GESTION NUMERO SERIE
*****************************************************************}
procedure CtrlDecochageGestionNoSerie(CC : TControl);
var
  FF : TForm;
begin
  FF := GetLaForm(CC);

  if ExisteNam('SO_GERESERIE', FF) then
  begin
    { Si décochage }
    if not ThCheckBox(GetFromNam('SO_GERESERIE', FF)).Checked then
    begin
      { Recherche s'il existe au moins un article géré en n° de série }
      if ExisteSQL('SELECT 1 FROM ARTICLE WHERE GA_NUMEROSERIE = "'+ wTrue + '"')
        or ExisteSQL('SELECT 1 FROM ARTICLECOMPL WHERE GA2_NUMEROSERIEGR = "'+ wTrue + '"') then
      begin
        PgiInfo ('Il y a au moins un article géré en n° de série, vous ne pouvez pas désactiver la gestion des n° de série.');
        ThCheckBox(GetFromNam('SO_GERESERIE', FF)).Checked := true;
      end;
    end;
  end;

end;
{$ENDIF STK}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF STK}
{***********A.G.L.***********************************************
Auteur  ...... : CVN
Créé le ...... : 06/04/2007
Modifié le ... : 06/04/2007
Description .. : On controle si il existe au
Suite ........ : moins un article géré en n° de série groupés, si c'est la cas on ne permet pas la
Suite ........ : désactivation de la gestion des n° de série groupés.
Mots clefs ... : GESTION NUMERO SERIE
*****************************************************************}
procedure CtrlDecochageGestionNoSerieGrp(CC : TControl);
var
  FF : TForm;
begin
  FF := GetLaForm(CC);

  if ExisteNam('SO_GERESERIEGRP', FF) then
  begin
    { Si décochage }
    if not ThCheckBox(GetFromNam('SO_GERESERIEGRP', FF)).Checked then
    begin
      { Recherche s'il existe au moins un article géré en n° de série groupés}
      if ExisteSQL('SELECT 1 FROM ARTICLECOMPL WHERE GA2_NUMEROSERIEGR = "'+ wTrue + '"') then
      begin
        PgiInfo ('Il y a au moins un article géré en n° de série groupés, vous ne pouvez pas désactiver la gestion des n° de série groupés.');
        ThCheckBox(GetFromNam('SO_GERESERIEGRP', FF)).Checked := true;
      end;
    end;
  end;
end;
{$ENDIF STK}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF STK}
{***********A.G.L.***********************************************
Auteur  ...... : CVN
Créé le ...... : 10/04/2007
Modifié le ... : 10/04/2007
Description .. : On controle si il existe au
Suite ........ : moins un mouvement sur lequel on gère les statuts de disponibilité
Mots clefs ... : GESTION STATUT DE DISPONIBILITE
*****************************************************************}
procedure CtrlDecochageGestionStatutDispo(CC : TControl);
var
  FF : TForm;
begin
  FF := GetLaForm(CC);

  if ExisteNam('SO_GERESTATUTDISPO', FF) then
  begin
    { Si décochage }
    if not ThCheckBox(GetFromNam('SO_GERESTATUTDISPO', FF)).Checked then
    begin
      { Recherche s'il existe au moins un mouvement avec gestion de statut de dispo }
      if ExisteSQL('SELECT 1 FROM STKMOUVEMENT WHERE GSM_STKTYPEMVT="PHY" AND GSM_STATUTDISPO="BLQ" ') then
      begin
        PgiInfo ('Il y a au moins un mouvement avec un statut de disponibilité bloqué, vous ne pouvez pas désactiver la gestion des statuts de disponibilité.');
        ThCheckBox(GetFromNam('SO_GERESTATUTDISPO', FF)).Checked := true;
      end;
    end;
  end;
end;
{$ENDIF STK}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF STK}
{***********A.G.L.***********************************************
Auteur  ...... : CVN
Créé le ...... : 10/04/2007
Modifié le ... : 10/04/2007
Description .. : On controle si il existe au
Suite ........ : moins un mouvement sur lequel on gère les statuts de flux
Mots clefs ... : GESTION STATUT DE FLUX
*****************************************************************}
procedure CtrlDecochageGestionStatutFlux(CC : TControl);
var
  FF : TForm;
begin
  FF := GetLaForm(CC);

  if ExisteNam('SO_GERESTATUTFLUX', FF) then
  begin
    { Si décochage }
    if not ThCheckBox(GetFromNam('SO_GERESTATUTFLUX', FF)).Checked then
    begin
      { Recherche s'il existe au moins un mouvement avec gestion de statut de flux }
      if ExisteSQL('SELECT 1 FROM STKMOUVEMENT WHERE GSM_STKTYPEMVT="PHY" AND GSM_STATUTFLUX <> "STD"') then
      begin
        PgiInfo ('Il y a au moins un mouvement avec un statut de flux renseigné, vous ne pouvez pas désactiver la gestion des statuts de flux.');
        ThCheckBox(GetFromNam('SO_GERESTATUTFLUX', FF)).Checked := true;
      end;
    end;
  end;
end;
{$ENDIF STK}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF STK}
{***********A.G.L.***********************************************
Auteur  ...... : CVN
Créé le ...... : 10/04/2007
Modifié le ... : 10/04/2007
Description .. : On controle si il existe au
Suite ........ : moins un article sur lequel on gère les marques
Mots clefs ... : GESTION MARQUES
*****************************************************************}
procedure CtrlDecochageGestionMarque(CC : TControl);
var
  FF : TForm;
begin
  FF := GetLaForm(CC);

  if ExisteNam('SO_GEREMARQUE', FF) then
  begin
    { Si décochage }
    if not ThCheckBox(GetFromNam('SO_GEREMARQUE', FF)).Checked then
    begin
      { Recherche s'il existe au moins un article avec une gestion des marques }
      if ExisteSQL('SELECT 1 FROM ARTICLE WHERE GA_GMARQUE = "'+ wTrue + '"') then
      begin
        PgiInfo ('Il y a au moins un article avec la gestion par marques, vous ne pouvez pas désactiver la gestion des marques.');
        ThCheckBox(GetFromNam('SO_GEREMARQUE', FF)).Checked := true;
      end;
    end;
  end;
end;
{$ENDIF STK}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFDEF STK}
{***********A.G.L.***********************************************
Auteur  ...... : CVN
Créé le ...... : 10/04/2007
Modifié le ... : 10/04/2007
Description .. : On controle si il existe au
Suite ........ : moins un article sur lequel on gère les choix qualité
Mots clefs ... : GESTION CHOIX QUALITE
*****************************************************************}
procedure CtrlDecochageGestionChoixQualite(CC : TControl);
var
  FF : TForm;
begin
  FF := GetLaForm(CC);

  if ExisteNam('SO_GERECHOIXQUALITE', FF) then
  begin
    { Si décochage }
    if not ThCheckBox(GetFromNam('SO_GERECHOIXQUALITE', FF)).Checked then
    begin
      { Recherche s'il existe au moins un article avec une gestion des choix qualité }
      if ExisteSQL('SELECT 1 FROM ARTICLE WHERE GA_GCHOIXQUALITE = "'+ wTrue + '"') then
      begin
        PgiInfo ('Il y a au moins un article avec la gestion choix qualité, vous ne pouvez pas désactiver la gestion des choix qualité.');
        ThCheckBox(GetFromNam('SO_GERECHOIXQUALITE', FF)).Checked := true;
      end;
    end;
  end;
end;
{$ENDIF STK}
{$ENDIF EAGLSERVER}
//  BBI Fin web services

{***********A.G.L.***********************************************
Auteur  ...... : C.B
Créé le ...... : 15/05/2007
Modifié le ... : 15/05/2007
Description .. : commun aux produits ga, grc, paie et bureau
Mots clefs ... : GESTION CHOIX QUALITE
*****************************************************************}
{$IFNDEF EAGLSERVER}
procedure planningIntegre(FF: TForm);
begin
  if ExisteNam('SO_YPLGA',FF) then
  begin
    THMultiValComboBox(GetFromNam('SO_YPLGA',FF)).Complete := True;
    THMultiValComboBox(GetFromNam('SO_YPLGA',FF)).Aucun := true;
  end;

  if ExisteNam('SO_YPLGRC',FF) then
  begin
    THMultiValComboBox(GetFromNam('SO_YPLGRC',FF)).Complete := True;
    THMultiValComboBox(GetFromNam('SO_YPLGRC',FF)).Aucun := true;
  end;

  if ExisteNam('SO_YPLPAIE',FF) then
  begin
    THMultiValComboBox(GetFromNam('SO_YPLPAIE',FF)).Complete := True;
    THMultiValComboBox(GetFromNam('SO_YPLPAIE',FF)).Aucun := true;
  end;

  if ExisteNam('SO_YPLBUREAU',FF) then
  begin
    THMultiValComboBox(GetFromNam('SO_YPLBUREAU',FF)).Complete := True;
    THMultiValComboBox(GetFromNam('SO_YPLBUREAU',FF)).Aucun := true;
  end;
{$ifdef AFFAIRE}
TraduitForm(FF); //mcd 02/01/2008
{$endif AFFAIRE}
end;

procedure GereSoldePiecePrec (CC : TControl);
var
  Form        : TForm ;
  bGere : boolean;
begin
if ctxscot in v_pgi.pgicontexte then exit; //mcd 13/07/07 14360 GIGA
  Form := GetLaForm(CC);
  bGere := TCheckBox(CC).Checked ;

  SetEna('SO_SOLDEPRECCONFIRM', Form, bGere);
  SetEna('SO_SOLDEPRECMODE'   , Form, bGere);
  SetEna('LSO_SOLDEPRECMODE'  , Form, bGere);
  SetEna('SO_SOLDEPRECVEN'    , Form, bGere);
  SetEna('LSO_SOLDEPRECVEN'   , Form, bGere);
  SetEna('SO_SOLDEPRECACH'    , Form, bGere);
  SetEna('LSO_SOLDEPRECACH'   , Form, bGere);
end;
{$ENDIF EAGLSERVER}
function IsMasterOnShare (elementpartage : String) : boolean;
begin
  if elementpartage <> '' then
    result := (not ExisteSql ('SELECT 1 FROM DESHARE WHERE DS_NOMBASE<>"' + V_PGI.SchemaName + '" AND DS_NOMTABLE="'+elementpartage+'"'))
  else
    result := (not ExisteSql ('SELECT 1 FROM DESHARE WHERE DS_NOMBASE<>"' + V_PGI.SchemaName + '"'));
end;

end.

