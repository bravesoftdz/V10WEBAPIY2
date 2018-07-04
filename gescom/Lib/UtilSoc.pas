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
    FE_Main,
    DBGrids,
    DBCtrls,
    DB,
    {$IFNDEF DBXPRESS} {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF} {$ELSE} uDbxDataSet, {$ENDIF}
    HDB,
  {$ENDIF EAGLCLIENT}
  {$IFDEF PGIIMMO}
    uImoParam,
  {$ENDIF PGIIMMO}
  {$IFDEF AMORTISSEMENT}
    ImPlan,
  {$ENDIF AMORTISSEMENT}
  {$IFDEF AFFAIRE}
    TraducAffaire,
    UtilArticle,
    DicoAf,
    {$IFDEF GIGI}
      galoutil,
    {$ENDIF GIGI}
  {$ENDIF AFFAIRE}
  {$IFDEF GCGC}
    EntGC,
  {$ENDIF GCGC}
  WinProcs,
  shellAPI,
  HDebug,
  Hstatus,
  LicUtil,
  HMsgBox,
  UtilPGI,
  Messages,
  Hrichedt,
  HTB97,
  Dialogs,
  UTOB,
  LookUp
  ;

Function  InterfaceSoc ( CC : TControl ) : boolean ;
Function  ChargePageSoc ( CC : TControl ) : boolean ;
Function  SauvePageSoc ( CC : TControl ) : boolean ;
Function  SauvePageSocSansverif ( CC : TControl ) : boolean ;
Procedure MarquerPublifi ( Flag : boolean ) ;
Procedure TripoteStatus ( StAjout : String = '')  ;
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
Procedure VerifGestLienOperPiece ( CC : TControl ) ;
procedure MajTabletteGcTypeFourn( CC : TControl );

{$ifdef AFFAIRE}
Function ExisteProfilGener ( Profil : String ) : boolean ;
Function ExisteArticle ( Art : String ) : boolean ;
Function ExisteEtabliss ( Eta : String ) : boolean ;
Function ExisteUnite ( Uni : String ) : boolean ;
Function ExisteAffaires : boolean ;
{$endif}
{$IFDEF GPAOLIGHT}
	procedure InitParamSocGPAO(FF: TForm);
{$ENDIF GPAOLIGHT}

// GCO - 22/04/2005 - Utilisé dans les projets CCS5 et CCSTD
function TraitePlanDeRevision ( CC : TControl ) : Boolean;
function  CPEstLiasseModifiee : Boolean;

function  AffecteToutLesCyclesRevision( vStPlanRevision : string = '') : Boolean;
function  AffecteCycleRevisionAuxGene( vStCycleRevision, vStCompte1, vStExclusion1 : string ) : Boolean;
procedure SupprimeCycleRevisionDesGene( vStCycleRevision : string = '' );
procedure MiseAJourPlanRevision;

{$IFDEF GESCOM}
  procedure CdeOuvCacheFonctions(FF: TForm);
  procedure CdeOuvTraiteAutoTransfo(CC: TControl);
{$ENDIF GESCOM}

{$IFDEF GCGC}
  procedure SetPlusUniteParDefaut(CC: TControl);
  procedure SetGCGestUniteModeState(CC: TControl);
  procedure GCChangeGestUniteMode(CC: TControl);
{$ENDIF GCGC}

//PCH 11-2005 personnalisation paramsoc
{$IFDEF GIGI}
  Function VerifAfPerso ( CC : TControl ) : boolean ;
  Procedure CtlPersoParamsocChoix( CC : TControl ) ;
  Procedure CtlPersoParamsocPar1( CC : TControl ) ;
  Procedure CtlPersoParamsocPar2( CC : TControl ) ;
  Function DecochagePersoParamsocChoix( CC : TControl ) : boolean ;
{$ENDIF GIGI}
//

implementation

Uses
 {$IFDEF COMPTA}
  cloture ,
  OuvreExo ,
  {$IFDEF EAGLCLIENT}
  AnulCloServeur ,
  {$ELSE}
  AnulClo ,
  {$ENDIF}
  {$ENDIF}
  Ent1,
  SaisUtil,
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
  ,AfterImportGEDGRC
  {$ENDIF GIGI}
  ,UGedFiles
  {$ENDIF GRC}
  {$IFNDEF BTP}
  {$IFDEF GCGC}
    {$IFNDEF PGIMAJVER}
      {$IFNDEF CMPGIS35}
      ,yTarifs
      {$ENDIF !CMPGIS35}
    {$ENDIF !PGIMAJVER}
//    ,yTarifsCommun
  {$ENDIF GCGC}
  {$ENDIF BTP}
  {$IFNDEF PGIIMMO}
  ,wCommuns
  {$ENDIF PGIIMMO}
  {$IFDEF GPAOLIGHT}
  	{$IFNDEF GPAO}
	  ,GpLightSeria
  	{$ENDIF !GPAO}
    ,GcStkValo
  {$ENDIF GPAOLIGHT}
  ;
Var DevisePrinc : String [3];

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

Function ExisteAfCumul : boolean ;
BEGIN
Result := ExisteSQL('SELECT ACU_TYPEAC FROM AFCUMUL WHERE ACU_TYPEAC="CVE"' ) ;
END ;

Function ExisteUnite ( Uni : String ) : boolean ;
BEGIN
Result:=False ; if Uni='' then Exit ;
Result := ExisteSQL('SELECT GME_MESURE FROM MEA WHERE GME_MESURE="'+Uni+'"') ;
END ;

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
    end;
END;


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
    if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
      SetEna('SO_AFEACTIMPSELEC', FF, False)
    else
    begin
      SetInvi('SO_AFEACTIMPSELEC',FF);
      SetInvi('SCO_AFGRPEACT',FF);
      SetInvi('LSCO_AFGRPEACT',FF);
    end;          
       
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
    //PCH 11-2005 Personnalisation paramsoc
    {$IFDEF GIGI}
    CbPersoChoix    : TCheckBox;
    CVCBPersoParam1 : THValComboBox;
    CVCBPersoParam2 : THValComboBox;
    {$ENDIF GIGI}
    //
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFACOMPTE',FF) then Exit ;


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

    end
else
    begin
    SetParamsoc('SO_AFLIENDP',FALSE);
    SetInvi('SO_AFLIENDP', FF);
    SetInvi('SO_GICREATDOSS', FF);
    SetInvi('SO_GICREATANNUAIRE_FIN', FF);
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
{$ENDIF GIGI}
//

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
SetEna('LSO_DATECUTOFFACH', FF, True); //AB-200406 Cut-off des achats
SetEna('SO_DATECUTOFFACH', FF, True);
   // accès aux params cut off uniquement avec mot de passe du jour ou si rien dans la table afcumul
if ExisteAfCumul then
  begin
    SetEna('SO_AFCUTMODEECLAT', FF, false);
    SetEna('LSO_AFCUTMODEECLAT', FF, false);
    SetEna('LSO_DATECUTOFF', FF, false);
    SetEna('SO_DATECUTOFF', FF, false);
    SetEna('LSO_DATECUTOFFACH', FF, false); //AB-200406 Cut-off des achats
    SetEna('SO_DATECUTOFFACH', FF, false);
    SetEna('LSO_AFFORMULCUTOFF', FF, false);
//mcd 27/05/03    end;
  end
else
  if (ctxScot  in V_PGI.PGIContexte) then
    begin
    // Dans Scot, la formule est figée par défaut
    SetParamsoc('SO_AFFORMULCUTOFF','[M0]-[M2]');
    end;

//AB-200406 - Cut-off des achats
if (ctxScot  in V_PGI.PGIContexte) or (not VH_GC.GAAchatSeria) or
   ((ctxGCAFF in V_PGI.PGIContexte) and not VH_GC.GASeria) then
begin
  SetInvi('SO_AFAPPCUTOFFACH', FF);
  SetInvi('SO_DATECUTOFFACH', FF);
  SetInvi('LSO_DATECUTOFFACH', FF);
  SetInvi('SO_AFFORMULECUTOFFACH', FF);
  SetInvi('LSO_AFFORMULECUTOFFACH', FF);
  SetInvi('SCO_AFGPCUTOFFACCPTA', FF);
  SetInvi('LSCO_AFGPCUTOFFACCPTA', FF);
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
end
else
begin
  SetParamsoc('SO_AFFORMULECUTOFFACH','[M2]-[M6]');  //formule des cut-off achats non-modifiables
  SetEna('SO_AFFORMULECUTOFFACH', FF, False);
  SetEna('LSO_AFFORMULECUTOFFACH', FF, False);
end;

If CtxScot in V_PGI.PGICOntexte then
 begin   //mcd 02/03/2005 pas de génération compta à c ejour.. on cache
  if (GetParamSocSecur ('SO_AFCLIENT', 0) <>8 ) then  //cInClientKPMG =8 pour ne pas inclure AffaireUtil
    begin
    SetInvi('SCO_AFGPCUTOFFCPTA', FF);
    SetInvi('LSCO_AFGPCUTOFFCPTA', FF);
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
// Dans tous les cas la formule n'est modifiable que par le menu 74752 de constitution de la formule du cut off
SetEna('SO_AFFORMULCUTOFF', FF, False);


TraduitForm(FF);
END;

Procedure ConfigComportementActivite ( CC : TControl ) ;
Var FF   : TForm ;
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFRECHTIERSPARNOM',FF) then Exit ;

if (ctxScot in V_PGI.PGIContexte) then
  begin
  THLabel(GetFromNam('SO_AFSAISINTERVAFF',FF)).caption:='Saisie réduite aux intervenants des missions';
	end;

SetInvi('SO_AFTESTPLANNINGJOUR', FF); // en attendant la FQ 12227


TraduitForm(FF);
END;

Procedure ConfigSaisieActivite ( CC : TControl ) ;
Var FF   : TForm ;
    CbFraisKM : TCheckBox;
BEGIN
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFDATEDEBUTACT',FF) then Exit ;
// en attendant de mettre la gestion de valo par type d'article
SetInvi('SO_AFVALOFRAISPV', FF);  SetInvi('LSO_AFVALOFRAISPV', FF);
SetInvi('SO_AFVALOFRAISPR', FF);  SetInvi('LSO_AFVALOFRAISPR', FF);
SetInvi('SO_AFVALOFOURPV', FF);   SetInvi('LSO_AFVALOFOURPV', FF);
SetInvi('SO_AFVALOFOURPR', FF);   SetInvi('LSO_AFVALOFOURPR', FF);

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
    if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
         SetEna('SO_AFDATEDEBUTACT', FF, True)
    else SetEna('SO_AFDATEDEBUTACT', FF, False);
   end;
SetInvi('SO_AFUTILTARIFACT', FF); //pas géré pour l'instant
{$IFDEF CCS3}
SetInvi('SO_AFVISAACTIVITE', FF); //pas gérer pour l'instant
{$ENDIF}

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

TraduitForm(FF);
END;

//PL 18/02/04 nouvelle fct
Procedure ConfigFrais (CC : TControl);
Var FF   : TForm ;
    CbFrais : TCheckBox;
BEGIN
// test de la page
FF := TForm (CC.Owner);
if Not ExisteNam ('SO_AFFRAISCOMPTA', FF) then Exit;
CbFrais := TCheckBox (GetFromNam ('SO_AFFRAISCOMPTA', FF));

If CbFrais.checked then
  begin
    SetEna('SO_AFFRAISJOURNAL', FF, True);
    SetEna('SO_AFFRAISMODEGENER', FF, True);
    SetEna('SO_AFMODESAISFRAIS', FF, True);
    SetEna('SO_AFPRIXGEREFRAISACT', FF, True);
    SetEna('SO_AFALIMAUTOPRIXTTC', FF, True);
  end
else
  begin
    SetEna('LSO_AFFRAISJOURNAL', FF, false);
    SetEna('LSO_AFFRAISMODEGENER', FF, false);
    SetEna('LSO_AFMODESAISFRAIS', FF, false);
    SetEna('LSO_AFPRIXGEREFRAISACT', FF, false);
    SetEna('SO_AFFRAISJOURNAL', FF, false);
    SetEna('SO_AFFRAISMODEGENER', FF, false);
    SetEna('SO_AFMODESAISFRAIS', FF, false);
    SetEna('SO_AFPRIXGEREFRAISACT', FF, false);
    SetEna('SO_AFALIMAUTOPRIXTTC', FF, false);
  end;

TraduitForm(FF);
END;

Procedure ConfigSaisieImport ( CC : TControl ) ;
Var FF   : TForm ;
BEGIN     //mcd 10/11/2004 pour traduction
// Test de la page
FF:=TForm(CC.Owner) ; if Not ExisteNam('SO_AFEACTPERIODIC',FF) then Exit ;
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

if ctxScot in V_PGI.PGIContexte then
  begin
  If Vh_GC.GAFactInfo then   THValComboBox(C).Plus :=' AND ((CO_LIBRE="GA" AND CO_CODE<>"CON") OR (CO_LIBRE="GI")) '
     else  THValComboBox(C).Plus :=' AND CO_LIBRE="GA" AND CO_CODE<>"CON" ' ;
  end;

THValComboBox(C).Value:=OldV ;
// Impossible de modifier le format de l'exercice s'il existe deja des affaires dans la base
// mcd 21/11/01 ne pas faire ... ne met pas à jour les info de code affaire !!
//SetEna( 'SO_AFFORMATEXER', FF,  Not ExisteAffaires) ;
SetEna( 'SO_AFFORMATEXER', FF,  False);
// gm non visible en  V7, prévu pour V7.10
   SetInvi('SO_AFDETOPACT', FF);
   SetInvi('SO_AFGENFAC', FF);
   SetInvi('SO_AFADRESSEAFF', FF);


if (ctxScot  in V_PGI.PGIContexte) then
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
  // PL le 17/05/05 : gestion de la fermeture de mission sur échéance liquidative
   if GetParamSoc('SO_AFGERELIQUIDE') then
    SetEna ('SO_AFFERMAFFECHLIQ', FF, true)
   else
    SetEna ('SO_AFFERMAFFECHLIQ', FF, false);
end
else
   // En contexte Tempo
begin
   SetInvi('SO_AFFORMATEXER', FF);
   SetInvi('LSO_AFFORMATEXER', FF);

   SetInvi('SO_AFGERELIQUIDE', FF);
   SetInvi('SO_AFFERMAFFECHLIQ', FF);
   SetInvi('SO_AFLIBLIQFACTAF', FF);
   SetInvi('LSO_AFLIBLIQFACTAF', FF);
   {$IFDEF CCS3}
   SetInvi('SO_AFGENERAUTOLIG', FF);
   SetInvi('SO_AFMULTIECHE', FF);
   {$ENDIF}
end;

// SetInvi('SO_AFBUDZERO', FF);

TraduitForm(FF);
END;

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
   SetInvi('SO_AFCALCULFIN', FF);
   SetInvi('LSO_AFCALCULFIN', FF);
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

  {$IFDEF CCS3}
  SetInvi('SO_AFGERESSAFFAIRE', FF);
  SetInvi('SO_AFCOMPLETE', FF);
  {$ENDIF}
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
  If (CtxScot In V_PGI.PGICOntexte) and (not VH_GC.GAPlanningSeria)
    and (V_PGI.PassWord <> CryptageSt(DayPass(Date))) then
    begin
    SetEna('SO_AFAVANCEMODELE',FF, False);
    SetEna('SO_AFLIENARTTACHE',FF, False);
    SetEna('SO_AFTYPEARTDEF',FF, False);
    SetEna('SO_AFFAMILLEDEF',FF, False);
    SetEna('SO_AFTYPECONPLA',FF, False); //mcd 13/07/2005
    SetEna('SO_AFMODEPLANNING',FF, False); //mcd 13/07/2005
    SetEna('LSO_AFTYPECONPLA',FF, False); //mcd 13/07/2005
    SetEna('LSO_AFMODEPLANNING',FF, False); //mcd 13/07/2005
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

  // C.B 26/11/2004
  // pas encore développé, invisible
  SetInvi('SO_AFPLANSIMPLE', FF);
end;

Procedure ConfigSaisiePlanning(CC : TControl);
Var FF : TForm ;
BEGIN
  FF := TForm(CC.Owner); if Not ExisteNam('SO_AFALIGNREALISE',FF) then Exit ;
 
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
    SetInvi('SO_AFECOLE', FF);

  if not(VH_GC.GCIfDefCEGID) then
  begin

    SetInvi('SO_AFRESMAX', FF);
    SetInvi('LSO_AFRESMAX', FF);
  end;

  // C.B 07/04/2006
  // mis pu un dev a faire plus tard ...
  SetInvi('SO_AFREGLEACT', FF); 
END;

Procedure ConfigSaisiePDC(CC : TControl);
Var FF : TForm ;
BEGIN
  // Test de la page
  FF := TForm(CC.Owner) ; if Not ExisteNam('SO_AFPLANDECHARGE',FF) then Exit ;
  If (not VH_GC.GAPlanChargeSeria) then
  begin

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

Procedure ConfigSaisieGenePlanning(CC : TControl);
Var FF : TForm ;
BEGIN
  // Test de la page
  FF := TForm(CC.Owner) ; if Not ExisteNam('SO_AFGESTIONRESS',FF) then Exit ;
  SetInvi('SO_AFGESTIONRESS',FF);
  SetInvi('SO_AFRESSREMP',FF);
  SetInvi('LSO_AFRESSREMP',FF);
END;

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
  end;
  TraduitForm(FF);
END;

Procedure CacheFctGIGA ( CC : TControl ) ;
BEGIN
// Configuration des écran GA (traduction, champs visibles...)
ConfigSaisieParamsDefaut(CC);
ConfigSaisieDevise(CC);
ConfigSaisieRessource(CC) ;
ConfigSaisiePrefAff(CC) ;
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
ConfigSaisiePDC(CC);
ConfigSaisiePlanning(CC);
ConfigSaisieGenePlanning(CC);
ConfigSaisieRevEtVar(CC);
ConfigSaisieImport(CC); //mcd 10/11/2004
ConfigFrais (CC); // PL le 18/02/04
ConfigArticle (CC); // MCD 16/05/05
ConfigParamPiece (CC); //AB-200510-
end;
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
Procedure AFPlanDeChargeX( CC : TControl ) ;
Var
  FF    : TForm ;
  CbPC  : TCheckBox;

BEGIN

  FF:=GetLaForm(CC) ;
  CbPC:=TCheckBox(CC);
  if CbPC.Checked then
  begin  // si coché, on renseigne les valeurs par défaut
    SetEna('SO_AFLIENPLANACT',FF, False);
    SetEna('SO_AFALIGNREALISE',FF, False);
    SetEna('SO_AFPLANNINGETAT',FF, False);
    SetEna('SO_AFETATINTERDIT',FF, False);
    SetEna('SO_AFPLANAFFICH',FF, False);

    SetInvi('LSO_AFDATETEXTE1', FF);
    SetInvi('LSO_AFDATETEXTE1', FF);

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
    SetEna('SO_AFLIENPLANACT',FF, True);
    SetEna('SO_AFALIGNREALISE',FF, True);
    SetEna('SO_AFPLANNINGETAT',FF, True);
    SetEna('SO_AFETATINTERDIT',FF, True);
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

// PL le 17/05/05
Procedure AfFactLiquideX (CC : TControl);
Var FF : TForm;
    CbFactLiquid, CbFermMissLiquid : TCheckBox;
BEGIN
FF := GetLaForm(CC);
CbFactLiquid := TCheckBox(CC);

CbFermMissLiquid := TCheckBox (GetFromNam ('SO_AFFERMAFFECHLIQ', FF));


if CbFactLiquid.Checked then
    begin
    CbFermMissLiquid.Enabled := true;
    end
else
    begin 
    CbFermMissLiquid.Enabled := false;
    CbFermMissLiquid.Checked := false;
    end;

END ;

// PL le 18/02/04 : nouvelle fct
Procedure FraisX (CC : TControl);
Var FF : TForm;
    CbFrais, CbAlimPrixTTC : TCheckBox;
    LblFraisJournal, LblFraisModeGener, LblModeSaisFrais, LblPrixGereFraisAct : THLabel;
    CMEFraisJournal, CVCBFraisModeGener, CVCBModeSaisFrais, CVCBPrixGereFraisAct: THValComboBox;
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


if CbFrais.Checked then
    begin
    CMEFraisJournal.Enabled := true;
    CVCBFraisModeGener.Enabled := true;
    CVCBModeSaisFrais.Enabled := true;
    CVCBPrixGereFraisAct.Enabled := true;
    LblFraisJournal.Enabled := true;
    LblFraisModeGener.Enabled := true;
    LblModeSaisFrais.Enabled := true;
    LblPrixGereFraisAct.Enabled := true;
    if (CVCBFraisModeGener.value='') then CVCBFraisModeGener.Value := 'DET';
    if (CVCBModeSaisFrais.value='') then CVCBModeSaisFrais.Value := 'SHT';
    if (CVCBPrixGereFraisAct.value='') then CVCBPrixGereFraisAct.Value := 'PV';
		CbAlimPrixTTC.Enabled := true;
    end
else
    begin   // si case non cochée, on efface tout ce qui est dans les zones suivantes
    CMEFraisJournal.Enabled := false;
    CVCBFraisModeGener.Enabled := false;
    CVCBModeSaisFrais.Enabled := false;
    CVCBPrixGereFraisAct.Enabled := false;
    LblFraisJournal.Enabled := false;
    LblFraisModeGener.Enabled := false;
    LblModeSaisFrais.Enabled := false;
    LblPrixGereFraisAct.Enabled := false;
    CMEFraisJournal.value := '';
    CVCBFraisModeGener.value := '';
    CVCBModeSaisFrais.value := '';
    CVCBPrixGereFraisAct.value := '';
		CbAlimPrixTTC.Enabled := false;
		CbAlimPrixTTC.Checked := false;
    end;

END ;

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


Function RechExisteH ( CC : TControl ) : boolean ;
BEGIN Result:=(RechDomLookUpCombo(CC)<>'') ; END ;

Function DesEnregs ( TableName : String ) : boolean ;
BEGIN
Result := ExisteSQL('SELECT 1 FROM '+TableName);
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

//
// Attribution automatique du code Tiers
//
Procedure NumTiersAutoX ( CC : TControl; ForceActive : Boolean ) ;
Var FF : TForm ;
    NumTiersAuto : TCheckBox ;
    LgNumTiers   : TSpinEdit ;
    CompteurTiers: TEdit ;
    Active       : Boolean ;
BEGIN
FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_GCNUMTIERSAUTO',FF) then Exit ;
if ForceActive=True then Active:=True
else
  begin
  NumTiersAuto:=TCheckBox(GetFromNam('SO_GCNUMTIERSAUTO',FF)) ;
  if NumTiersAuto.Checked then
    begin
    Active := True;
    LgNumTiers:=TSpinEdit(GetFromNam('SO_GCLGNUMTIERS',FF));
    if Not (ctxChr in V_PGI.PGIContexte) then
        LgNumTiers.Value:=17
    else if (LgNumTiers.Value=0) then LgNumTiers.Value:=6;
    CompteurTiers:=TEdit(GetFromNam('SO_GCCOMPTEURTIERS',FF));
    if CompteurTiers.Text='' then CompteurTiers.Text:='0';
    end
    else Active := False;
  end;
SetEna('SO_GCPREFIXETIERS',FF,Active) ;
SetEna('LSO_GCPREFIXETIERS',FF,Active) ;
SetEna('SO_GCLGNUMTIERS',FF,Active) ;
SetEna('LSO_GCLGNUMTIERS',FF,Active) ;
SetEna('SO_GCCOMPTEURTIERS',FF,Active) ;
SetEna('LSO_GCCOMPTEURTIERS',FF,Active) ;
END ;

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

procedure GereContremarque(CC : TControl);
var FF : TForm;
    Gere, GereAuto, GereAutoTr : Boolean;
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
  if ExisteNam('SO_GCAFFCONTREM', FF) then
  begin
    SetInvi ('LSO_GCAFFCONTREM', FF);
    SetInvi ('SO_GCAFFCONTREM', FF);
  end;
  SetInvi('SO_GCFORCEQTEAUTOCONTREM', FF);
  { Fin - En attendant de savoir gérer }
  Gere := (TCheckBox(GetFromNam('SO_GERECONTREMARQUE',FF)).Checked);
  GereAuto := ((Gere) and (TCheckBox(GetFromNam('SO_GCPCEAUTOCONTREM',FF)).Checked));
  GereAutoTr := ((GereAuto) and (TCheckBox(GetFromNam('SO_GCFORCEAUTOCONTREM',FF)).Checked));
  EnaChpEtLib('SO_GCNATRECCONTREM', Gere);
  EnaChpEtLib('SO_GCNATLIVCONTREM', Gere);
  SetEna('SO_GCAUTOCONTREM' , FF, Gere);
  SetEna('SO_GCAFFCONTREM', FF, Gere);
  SetEna('SO_MAILCONTREMARQUE', FF, Gere);
  SetEna('SO_GCPCEAUTOCONTREM', FF, Gere);
  SetEna('SO_GCFORCEAUTOCONTREM', FF, GereAuto);
  EnaChpEtLib('SO_GCNATACHAUTOCONTREM', GereAutoTr);
  if not Gere then
  begin
    TCheckBox(GetFromNam('SO_MAILCONTREMARQUE', FF)).Checked := false;
    TCheckBox(GetFromNam('SO_GCPCEAUTOCONTREM', FF)).Checked := false;
    TCheckBox(GetFromNam('SO_GCFORCEAUTOCONTREM', FF)).Checked := false;
  end;
  if not GereAuto then
    TCheckBox(GetFromNam('SO_GCFORCEAUTOCONTREM', FF)).Checked := false;
  if not GereAutoTr then
    THValComboBox(GetFromNam('SO_GCNATACHAUTOCONTREM', FF)).Value := '';
  LePlus := 'AND (SELECT GSN_STKTYPEMVT FROM STKNATURE WHERE GSN_QUALIFMVT=GPP_STKQUALIFMVT)="PHY"';
  if ExisteNam('SO_GCNATRECCONTREM', FF) then
    THValComboBox(GetFromNam('SO_GCNATRECCONTREM',FF)).Plus := LePlus;
  if ExisteNam('SO_GCNATLIVCONTREM', FF) then
    THValComboBox(GetFromNam('SO_GCNATLIVCONTREM',FF)).Plus := LePlus;
end;

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
END;

//
// Attribution automatique du code Article
//
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
END ;

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

//
// Attribution automatique du Code à barres
//
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

//
// Gestion des familles articles 4 à 8
//
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

  if TControl(FF.FindComponent('SO_CROISAXE'))=nil then exit;
  SO_CROISAXE:=TCheckBox(GetFromNam('SO_CROISAXE',FF));
  SO_CROISAXE.Enabled := false;
  for i:=1 to MaxAxe do
  begin
    SetEna('SO_VENTILA'+IntToStr(i),FF,false);
  end;
end;
{$ENDIF}
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
end;


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
end;

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

Function VerifCompensation(CC : TControl): boolean;
var
  FF:        TForm;
  CbGestion: TCheckBox;
  NumPlan:   Integer;
begin
  Result    := True;
  FF        := TForm(CC.Owner) ;
  if Not ExisteNam('SO_CPGESTCOMPENSATION',FF) then Exit ;
  CbGestion := TCheckBox(GetFromNam('SO_CPGESTCOMPENSATION',FF));
  NumPlan   := 0;
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
{e FP 21/02/2006}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/03/2002
Modifié le ... :   /  /
Description .. : LG* ajout de la gestion des champs comptes de l'onglet
Suite ........ : lettrage\gestion
Mots clefs ... :
*****************************************************************}
Function InterfaceSoc ( CC : TControl ) : boolean ;
Var
  NamChpExit : String ;   // Champ que l'on vient de quitter
  NamChpActif: String ;   // Champ sur lequel est positionné le focus
  ChampActif : TWinControl;
  C : ThValComboBox;
  C1: ThValComboBox;
  {$IFDEF AFFAIRE}  // fct spécif GI/GA
    FF:TForm;
  {$ENDIF AFFAIRE}
BEGIN

  Result:=True ;

  { Champ exit }
  NamChpExit:=CC.Name;
  if NamChpExit='SO_DATEDEBUTEURO'                then DateDebutEuroX(CC)
//FV : Conditionnement du ThvalCombo de l'article par défaut en fonction du
//type d'article dans la page paramètres Appels/Contrats... 27/02/2007
  else if NamChpExit='SO_BTTYPEARTDEF'            Then
     Begin
     C := ThValComboBox(CC);
     C.Value := 'PRE';
     // Recherche du contrôle Article
     ChampActif := CC.Parent ;
     if assigned(ChampActif) then
        begin
        C1 := THValComboBox(ChampActif.FindChildControl('SO_BTPRESTDEFAUT')) ;
        if assigned(C1) then
           //C1.Plus := C.Value;
           C1.Plus := 'PRE';
        end ;
     end
  else
  {$IFDEF AFFAIRE}  // fct spécif GI/GA
  if NamChpExit='SO_AFGESTIONAPPRECIATION'        then AppreciationX(CC)
  else if NamChpExit='SO_AFPLANDECHARGE'          then AfPlanDeChargeX(CC)
  else if NamChpExit='SO_AFGESTIONRAF'            then AfGestionRAFX(CC)
  else if NamChpExit='SO_AFAPPCUTOFF'             then CutOffX(CC)
  else if NamChpExit='SO_AFFRAISCOMPTA'           then FraisX(CC)
  else if NamChpExit='SO_AFGERELIQUIDE'           then AfFactLiquideX(CC)
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

  // PCH 11-2005 Gestion personnalisation paramsoc
  {$IFDEF GIGI}
  else if NamChpExit='SO_AFPERSOCHOIX'  then CtlPersoParamsocChoix(CC)
  else if NamChpExit='SO_AFPERSOPARAM1' then CtlPersoParamsocPar1(CC)
  else if NamChpExit='SO_AFPERSOPARAM2' then CtlPersoParamsocPar2(CC)
  {$ENDIF GIGI}

  else
  {$ENDIF AFFAIRE}
  if isFourcCpte(NamChpExit)                           then CptBourreX(CC)
  else if Copy(NamChpExit,1,9)='SO_DEFCOL'             then CptBourreX(CC)
  else if NamChpExit='SO_DEVISEPRINC'                  then DevPrincX(CC)
  else if NamChpExit='SO_CONTROLEBUD'                  then ControleBudX(CC)
  else if NamChpExit='SO_PAYS'                         then PaysX(CC)
  else if NamChpExit='SO_GCNUMTIERSAUTO'               then NumTiersAutoX(CC, False)
  else if NamChpExit='SO_GCPREFIXETIERS'               then PrefixeTiersOk(CC)
  else if NamChpExit='SO_GCCOMPTEURTIERS'              then LgNumTiersOk(CC)
  else if NamChpExit='SO_GCNUMARTAUTO'                 then CodeArtAutoX(CC, False)
  else if NamChpExit='SO_GCPREFIXEART'                 then PrefixeArtOk(CC)
  else if NamChpExit='SO_GCCOMPTEURART'                then LgCodeArtOk(CC)
  else if NamChpExit='SO_GCNUMCABAUTO'                 then NumCABAutoX(CC)
  else if NamChpExit='SO_ARTLOOKORLI'                  then FamillesArticle4a8(CC)
  else if NamChpExit='SO_RTPROJGESTION'                then AlimZonesProjet(CC)
  else if (NamChpExit='SO_GESTSTATUTUNIQUE1')
       or (NamChpExit='SO_GESTSTATUTUNIQUE2')          then GereZonesVersions(CC)
  else if NamChpExit='SO_GCFORMATEZONECLI'             then FormatZoneCli(CC)
  else if (NamChpExit='SO_RTPROPTABHIE')
       or (NamChpExit='SO_RTACTTABHIE')                then RTTabHie(CC,NamChpExit)
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
  {$ENDIF GCGC}
  else if (NamChpExit = 'SO_MESEFFICMINI') or (NamChpExit = 'SO_MESEFFICMAXI') then
    MESEfficience( CC )
  else if (NamChpExit = 'SO_MESCHPRESPRES') or (NamChpExit = 'SO_MESPRESMAXI') then
    MESControle( CC )
  else if NamChpExit = 'SO_GCAXEANALYTIQUE' then
    GereAnaParamAff(CC)
  else if NamChpExit = 'SO_GERETVAINTRACOMM' then
    GereTypeTVAIntraComm(CC)
  else if (NamChpExit = 'SO_GERECONTREMARQUE') or (NamChpExit = 'SO_GCPCEAUTOCONTREM') or (NamChpExit = 'SO_GCFORCEAUTOCONTREM') then
    GereContremarque(CC)
  else if NamChpExit = 'SO_GEREFARFAE' then
    GereFarFae(CC)
  ;

  { Champ actif }
  ChampActif := screen.ActiveControl;
  NamChpActif:=ChampActif.Name;
  if NamChpActif='SO_GCNUMTIERSAUTO'    then NumTiersAutoX(CC, True)
  else if NamChpActif='SO_GCNUMARTAUTO' then CodeArtAutoX(CC, True);
	{$IFDEF GPAOLIGHT}
  	InitParamSocGPAO(GetLaForm(CC));
	{$ENDIF GPAOLIGHT}
END ;

{============================= Sur Chargement =========================}
Function ExisteMvtsPayeurs ( Jal : String ) : boolean ;
Var Q : TQuery ;
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

{***********A.G.L.***********************************************
Auteur  ...... : JCF
Créé le ...... : 02/08/2001
Modifié le ... :   /  /
Description .. : Modification du blocage de la modification de la devise
Suite ........ : dans les paramètres société.
Suite ........ : Réalisé en spécifique pour la MODE
Mots clefs ... : PARAMSOC;EURO;MODE
*****************************************************************}
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

Procedure BlocageECC ( CC : TControl ) ;
Var FF : TForm ;
    Ena : boolean ;
BEGIN
FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_ECCEURODEBIT',FF) then Exit ;
Ena:=(GetTOBSoc(CC).GetValue('OKECCDEBIT')='-')  ; SetEna('SO_ECCEURODEBIT',FF,Ena) ;
Ena:=(GetTOBSoc(CC).GetValue('OKECCCREDIT')='-') ; SetEna('SO_ECCEUROCREDIT',FF,Ena) ;
END ;

{$IFDEF NETEXPERT}
Procedure InitComptaRevision ( CC : TControl ) ;
Var
  FF : TForm ;
BEGIN
  FF:=GetLaForm(CC) ;
  if ExisteNam('SO_CPLIENGAMME',FF) then
  begin
       if GetParamSoc ('SO_CPLIENGAMME')= 'S1' then
       begin
            if V_PGI.PassWord=CryptageSt(DayPass(Date)) then SetEna('SO_CPLIENGAMME', FF, TRUE)
            else if _IsDossierNetExpert then SetEna('SO_CPLIENGAMME', FF, false);
       end;
       SetInvi ('SO_CPFREQUENCESX', FF);  SetInvi ('LSO_CPFREQUENCESX', FF);

  end;
END;
{$ENDIF}


{***********A.G.L.***********************************************
Auteur  ...... : JCF
Créé le ...... : 02/08/2001
Modifié le ... :   /  /
Description .. : Modification pour la Mode des règles de modification des
Suite ........ : informations euro.
Suite ........ : Pour règler les dossiers en multi-devises
Mots clefs ... : PARAMSOC;MODE;EURO;MULTIDEVISE
*****************************************************************}
Procedure InitDivers ( CC : TControl ) ;
Var
  FF : TForm ;
  JalBud,RefL,RefP : THValComboBox ;
  OldV : String ;
  HH   : THLabel ;
  CB1,CB2 : tcheckBox ;
BEGIN
  FF:=GetLaForm(CC) ;
  {$IFDEF GIGI}
  if ExisteNam('SO_GEREALERTEAFF',FF) then traduitForm(FF) ;
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
    if ExisteNam('SO_GCLGNUMTIERS',FF) then SetInvi('SO_GCLGNUMTIERS',FF);
    if ExisteNam('LSO_GCLGNUMTIERS',FF) then SetInvi('LSO_GCLGNUMTIERS',FF) ;
    if ExisteNam('SO_GCPREFIXETIERS',FF) then SetInvi('SO_GCPREFIXETIERS',FF) ;
    if ExisteNam('LSO_GCPREFIXETIERS',FF) then SetInvi('LSO_GCPREFIXETIERS',FF) ;
    if ExisteNam('SO_GCLGNUMART',FF) then SetInvi('SO_GCLGNUMART',FF) ;
    if ExisteNam('LSO_GCLGNUMART',FF) then SetInvi('LSO_GCLGNUMART',FF) ;
    if ExisteNam('SO_GCPREFIXEART',FF) then SetInvi('SO_GCPREFIXEART',FF) ;
    if ExisteNam('LSO_GCPREFIXEART',FF) then SetInvi('LSO_GCPREFIXEART',FF) ;
  END;
{$ENDIF}
(*
  if ExisteNam('SO_GCNUMCABAUTO',FF) then
    SetInvi('SO_GCNUMCABAUTO',FF) ;
*)

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
    if ExisteNam('SO_RTNUMOPERAUTO',FF) then
    begin
      if TCheckBox(GetFromNam('SO_RTNUMOPERAUTO',FF)).Checked then
      begin
        SetEna('SO_RTCOMPTEUROPER',FF,True) ;
        SetEna('LSO_RTCOMPTEUROPER',FF,True) ;
      end
      else
      begin
        SetEna('SO_RTCOMPTEUROPER',FF,False) ;
        SetEna('LSO_RTCOMPTEUROPER',FF,False) ;
      end;
    end;
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
          (not ExisteSQL('SELECT EE_GENERAL FROM EEXBQ WHERE EE_ORIGINERELEVE<>"INT" OR EE_ORIGINERELEVE IS NULL')));
  end;

{$IFDEF CHR}
if (ExisteNam('SO_ACTIVECOMSX',FF)) then
   THLabel(GetFromNam('SO_ACTIVECOMSX',FF)).Caption:=TraduireMemoire('Gestion sur date de production') ;
if (ExisteNam('SO_HRMOULINETTE',FF)) AND (V_PGI.SAV) then
begin
  SetVisi('SO_HRMOULINETTE', FF);
  SetVisi('LSO_HRMOULINETTE', FF);
end;
{$ENDIF CHR}
  // GCO - 17/03/2005 - Contrôle si la liasse sélectionée a été modifiée
{$IFDEF COMPTA}
  if ExisteNam('SO_CPCONTROLELIASSE', FF) then
  begin
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
  // FIN GCO - 17/03/2005

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
  {$IFDEF GCGC}
  if (Not VH_GC.SAVSeria) then
  BEGIN
    if ExisteNam('SO_GCFICHEEDITPARC',FF) then SetInvi('SO_GCFICHEEDITPARC',FF) ;
  END;
  if ((not (ctxAffaire in V_PGI.PGIContexte)) and (not( ctxGCAFF in V_PGI.PGIContexte))) or (not VH_GC.GASeria) then
    if ExisteNam('SO_GCGEREALERTEGA',FF) then SetInvi('SO_GCGEREALERTEGA',FF) ;
  {$ENDIF GCGC}
END ;

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

{$IFDEF CCS3}
Procedure InviFourche ( St : String ; FF : TForm ) ;
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

Procedure CacheFonctionS3S5  ( CC : TControl ) ;
Var FF : TForm ;
BEGIN
FF:=GetLaForm(CC) ;   
(* JP 16/12/05 : FQ 17202 : S3, S5, S7 et PCL, ici, c'est du pareil au même.
                            Par ailleurs {$IFDEF CCS3} n'est à priori plus utilisée dans les DPR de Régis


  if ((EstSerie(S5)) or (EstSerie(S3))) then
  BEGIN
    if ExisteNam('SO_CORSGE2',FF) then
    BEGIN
      SetInvi('SO_CORSGE2',FF) ; SetInvi('SO_CORSAU2',FF) ;
      SetInvi('SO_CORSA12',FF) ; SetInvi('SO_CORSA22',FF) ;
      SetInvi('SO_CORSA32',FF) ; SetInvi('SO_CORSA42',FF) ;
      SetInvi('SO_CORSA52',FF) ; SetInvi('SO_CORSBU2',FF) ;
//      GetFromNam('SO_CORSBU1',FF).Top:=GetFromNam('SO_CORSA31',FF).Top ;
      SetInvi('SCO_TPLAN2',FF) ;
    END ;
    if ExisteNam('SO_BOURRELIB',FF) then SetInvi('SO_BOURRELIB',FF) ;
    if ExisteNam('LSO_BOURRELIB',FF) then SetInvi('LSO_BOURRELIB',FF) ;
    if ExisteNam('SO_DATEREVISION',FF) then
    BEGIN
      SetInvi('SO_DATEREVISION',FF) ; SetInvi('LSO_DATEREVISION',FF) ;
    END ;
  END ;

{$IFDEF CCS3}
   // Comptabilité S3 //
if ExisteNam('SCO_FACTOFRANCE',FF) then begin
  SetInvi('SCO_FACTOFRANCE',FF) ;
  SetInvi('LSCO_FACTOFRANCE',FF) ;
  SetInvi('SO_CODEVENDEUR',FF) ;
  SetInvi('LSO_CODEVENDEUR',FF) ;
  SetInvi('SO_NOQUITTANCE',FF) ;
  SetInvi('LSO_NOQUITTANCE',FF) ;
  SetInvi('SO_NOMFACTOR',FF) ;
  SetInvi('LSO_NOMFACTOR',FF) ;
  SetInvi('SCO_CGA',FF) ;
  SetInvi('LSCO_CGA',FF) ;
  SetInvi('SO_CGACOMPTE',FF) ;
  SetInvi('LSO_CGACOMPTE',FF) ;
  end;
// FQ 12949
// Onglets TVA
if ExisteNam('SO_HTPRORATATVA',FF) then
   BEGIN
   SetInvi('SO_HTPRORATATVA',FF) ;
   SetInvi('LSO_HTPRORATATVA',FF) ;
   SetInvi('SCO_PRORATATVA',FF) ;
   SetInvi('LSCO_PRORATATVA',FF) ;
   end;
// FQ 11739
// Onglets Dates
if ExisteNam('SO_DATECLOTUREPER',FF) then
   BEGIN
   SetInvi('LSO_DATEREVISION',FF);
   SetInvi('SO_DATEREVISION',FF);

   SetInvi('LSCO_GRPLAGEECR',FF);
   SetInvi('SCO_GRPLAGEECR',FF);
   SetInvi('SO_NBJECRAVANT',FF);
   SetInvi('LSO_NBJECRAVANT',FF);
   SetInvi('SO_NBJECRAPRES',FF);
   SetInvi('LSO_NBJECRAPRES',FF);

   SetInvi('LSCO_GRPLAGEECH',FF);
   SetInvi('SCO_GRPLAGEECH',FF);
   SetInvi('SO_NBJECHAVANT',FF);
   SetInvi('LSO_NBJECHAVANT',FF);
   SetInvi('SO_NBJECHAPRES',FF);
   SetInvi('LSO_NBJECHAPRES',FF);

   SetInvi('LSCO_GRJOURFERM',FF);
   SetInvi('SCO_GRJOURFERM',FF);
   SetInvi('SO_JOURFERMETURE',FF);
   SetInvi('SO_JOURFERMETURE1',FF);
   SetInvi('SO_JOURFERMETURE2',FF);
   SetInvi('SO_JOURFERMETURE3',FF);
   SetInvi('SO_JOURFERMETURE4',FF);
   SetInvi('SO_JOURFERMETURE5',FF);
   SetInvi('SO_JOURFERMETURE6',FF);
   end;
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
   InviFourche('BIL',FF) ; InviFourche('CHA',FF) ; InviFourche('PRO',FF) ;
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
//Sg6 14/01/05 Gestion de l'equilibrage des ODA en S3
//if ExisteNam('SO_EQUILANAODA',FF) then SetInvi('SO_EQUILANAODA',FF) ;
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
   if not (ctxAffaire in V_PGI.PGIContexte)  then  SetInvi('SO_GCVENTCPTAAFF',FF) ;
   SetInvi('SO_GCDISTINCTAFFAIRE',FF) ;
   SetInvi('SO_GCCPTAIMMODIV',FF) ;
   END ;
if ExisteNam('SO_GCINVPERM',FF) then
   BEGIN
   SetInvi('SCO_GCGRPINVPERM',FF)    ; // JTR - eQualité 11901
   SetInvi('SO_GCINVPERM',FF)       ;
   SetInvi('SO_GCJALSTOCK',FF)      ; SetInvi('LSO_GCJALSTOCK',FF) ;
   SetInvi('SO_GCMODEVALOSTOCK',FF) ; SetInvi('LSO_GCMODEVALOSTOCK',FF) ;
   END ;
if ExisteNam('SCO_STKSTANDARD',FF) then
   BEGIN
   SetInvi('SCO_STKSTANDARD',FF); SetInvi('LSCO_STKSTANDARD',FF);
   SetInvi('SO_GEREEMPLACEMENT',FF);       ;
   SetInvi('SO_GERELOT',FF); SetInvi('SO_GERESTATUTDISPO',FF) ;
   SetInvi('SO_GERESERIE',FF);
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
if ExisteNam ('SCO_GCGRLOGITRANSF', FF) then
begin
  SetInvi ('SCO_GCGRLOGITRANSF', FF);
  SetInvi ('LSCO_GCGRLOGITRANSF', FF);
end;
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
if ExisteNam('SCO_GRPCPTAAUTRE',FF) then
  begin
  SetInvi('SCO_GRPCPTAAUTRE',FF) ;
  SetInvi('LSCO_GRPCPTAAUTRE',FF) ;
  end;
if ExisteNam('SO_GCCPTEVARSTK',FF)    then BEGIN SetInvi('SO_GCCPTEVARSTK',FF) ; SetInvi('LSO_GCCPTEVARSTK',FF) ; END ;
if ExisteNam('SO_GCFICHEEDITPROPO',FF) then SetInvi('SO_GCFICHEEDITPROPO',FF) ;
// gm 08/10/04 pas de tarif avancés en S3
if ExisteNam('SO_TARIFSAVANCES',FF) then SetInvi('SO_TARIFSAVANCES',FF);
if ExisteNam('SO_FRAISAVANCES',FF)  then SetInvi('SO_FRAISAVANCES',FF);
{$ENDIF}
      *)

// GCO - 19/05/2005
  if (ExisteNam('SO_CPOKMAJPLAN',FF)) and (not (ctxPCL in V_PGI.PGIContexte)) then
    SetInvi('SO_CPOKMAJPLAN',FF);

  if (ExisteNam('SCO_CPREVISIONGRP2', FF)) then SetInvi('SCO_CPREVISIONGRP2', FF);
  if (ExisteNam('LSCO_CPREVISIONGRP2', FF)) then SetInvi('LSCO_CPREVISIONGRP2', FF);
  if (ExisteNam('SO_CPCONTROLELIASSE', FF)) then SetInvi('SO_CPCONTROLELIASSE', FF);
  if (ExisteNam('LSO_CPCONTROLELIASSE', FF)) then SetInvi('LSO_CPCONTROLELIASSE', FF);
  if (ExisteNam('SO_CPREVISBLOQUELIASSE', FF)) then SetInvi('SO_CPREVISBLOQUELIASSE', FF);
  if (ExisteNam('SO_CPREVISLIASSEMODIF', FF)) then SetInvi('SO_CPREVISLIASSEMODIF', FF);
  if (ExisteNam('SO_CPREVISEMESSAGELIASSE', FF)) then SetInvi('SO_CPREVISEMESSAGELIASSE', FF);
// FIN GCO
END ;

procedure GRCMasqueOptionsProjet (CC : TControl);
Var FF : TForm ;
BEGIN
FF:=GetLaForm(CC) ;
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

Procedure CacheFonctionBTP ( CC : TControl ) ;
Var
  FF : TForm ;
BEGIN
  FF:=GetLaForm(CC) ;
if not (CtxBtp in V_PGI.PGIContexte) then
   Begin
  if ExisteNam('SO_BTDOCAVECTEXTE',FF) then SetInvi('SO_BTDOCAVECTEXTE',FF) ;   //PCS Pas dans la 4.0 pour la Gescom
  if ExisteNam('SO_GCCPTERGVTE',FF)     then BEGIN SetInvi('SO_GCCPTERGVTE',FF) ; SetInvi('LSO_GCCPTERGVTE',FF) ; END ;
  if ExisteNam('SO_GCVENTILRG',FF)     then BEGIN SetInvi('SO_GCVENTILRG',FF) ;  END ;
   End;
  { // Ancienne gestion
	SetInvi('SCO_GRCCPTAAUTRE', FF);
	SetInvi('SO_GCCPTESTOCK', FF);
	SetInvi('SO_GCCPTEVARSTK', FF);
  }
END;

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
   if ExisteNam('SO_COMPTAEXTERNE',FF) then SetInvi('SO_COMPTAEXTERNE',FF) ;
   if ExisteNam('SO_ACTIVECOMSX',FF) then SetInvi('SO_ACTIVECOMSX',FF) ;
   if ExisteNam('SO_MBOCHEMINCOMPTA',FF) then SetInvi('SO_MBOCHEMINCOMPTA',FF) ;
   if ExisteNam('LSO_MBOCHEMINCOMPTA',FF) then SetInvi('LSO_MBOCHEMINCOMPTA',FF) ;
   if ExisteNam('SO_TYPECOMSX',FF) then SetInvi('SO_TYPECOMSX',FF) ;
   if ExisteNam('LSO_TYPECOMSX',FF) then SetInvi('LSO_TYPECOMSX',FF) ;
   if ExisteNam('SO_EXPJRNX',FF) then SetInvi('SO_EXPJRNX',FF) ;
   if ExisteNam('LSO_EXPJRNX',FF) then SetInvi('LSO_EXPJRNX',FF) ;
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
      if ExisteNam('LSO_GCTIERSMVSTK'      , FF) then SetEna ('LSO_GCTIERSMVSTK', FF, false);
      if ExisteNam ('SO_EDITRAPMVTEX', FF) then SetEna ('SO_EDITRAPMVTEX', FF, true);
    end else
    begin
      if ExisteNam ('SO_EDITRAPMVTEX', FF) then SetEna ('SO_EDITRAPMVTEX', FF, false);
      if ExisteNam('SO_GCTIERSMVSTK'       , FF) then SetEna ('SO_GCTIERSMVSTK', FF, true);
      if ExisteNam('LSO_GCTIERSMVSTK'      , FF) then SetEna ('LSO_GCTIERSMVSTK', FF, true);
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
{$IFDEF GCGC}
  if not VH_GC.VTCSeria then
  begin
    if ExisteNam('SCO_GCGRPCAISSE', FF) then
    begin
{$IFNDEF BTP}
      SetInvi('SCO_GCGRPCAISSE', FF);
{$ENDIF}
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
END;

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
  begin
    if ExisteNam(CName, FF) then
      THEdit(GetFromNam(CName, FF)).Text := Value;
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
    {$ELSE STK}
    {$ENDIF STK}
  end;

const
  PdrMethValoStd    : Boolean = False;
  FraisAvancesActifs: Boolean = False;
  StAGCActif				: Boolean = False;
  SauvParamSoc			: Boolean = False;
  SauvParamStAGC		: Boolean = False;
var
  ColisageActif,
  CtrlPoidsBrut,
  TarifsAvances,
  GereSousOrdres: Boolean;
  {$IFNDEF GPAO}
  ErrMsg        : string;
  {$ENDIF GPAO}
  lPalmares : boolean;
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
  SetControlVisible('LSCO_GRTARIFSAVANCES', TarifsAvances);
  SetControlVisible('SCO_GRTARIFSAVANCES' , TarifsAvances);
  SetControlVisible('LSO_TARIFSAISIE'     , TarifsAvances);
  SetControlVisible('SO_TARIFSAISIE'      , TarifsAvances);
  SetControlVisible('LSO_TARIFPARTIEL'    , TarifsAvances);
  SetControlVisible('SO_TARIFPARTIEL'     , TarifsAvances);
  SetControlVisible('SO_TARIFLIBAUTO'     , TarifsAvances);
  SetControlVisible('SO_TARIFCONDAPPL'    , TarifsAvances);
  SetControlVisible('SO_TARIFSGROUPES'    , TarifsAvances);
  SetControlVisible('LSO_TARIFSLANCRECH'  , TarifsAvances);
  SetControlVisible('SO_TARIFSLANCRECH'   , TarifsAvances);
  SetControlVisible('SO_TARIFSPRIXMARGES' , TarifsAvances);
  SetControlVisible('SO_TARIFSENREGISTRES', TarifsAvances);
  SetControlVisible('SO_TARIFSENREGFLUX'  , TarifsAvances);

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
    SetControlEnabled('SO_FRAISAVANCES' , not GetParamSoc('SO_FRAISAVANCES'));
    SetControlChecked('SO_FRAISAVANCES' , False);
  end
  else
	  SetControlEnabled('SO_FRAISAVANCES' 	  , not GetParamSoc('SO_FRAISAVANCES'));

  {Dès que les frais avancés sont cochés, on ne peut revenir en arrière}
  if (ExisteNam('SO_FRAISAVANCES', FF)) and (not FraisAvancesActifs)
  																			and (IsChecked('SO_FRAISAVANCES')) then
  begin
    if PGIAsk(TraduireMemoire('ATTENTION !!!' + #13 +
                              'Vous allez activer les frais annexes. ' + #13 +
                              'Vous devez obligatoirement gérer les tarifs avancés et cette action est irréversible.' + #13 +
                              'Confirmez-vous votre modification ?'),
              TraduireMemoire('Frais annexes')) = mrYes then
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
    if PGIAsk(TraduireMemoire('ATTENTION !!!' + #13 +
                              'Vous allez activer la valorisation au coût standard. ' + #13#13 +
                              'Cette gestion implique qu''aucune pièce et aucun mouvement ne doit mettre à jour ' + #13 +
                              'le dernier prix d''entrée et le dernier prix de revient d''entrée' + #13 +
                              'des fiches articles et des fiches de stocks'+ #13#13+
                              'Confirmez-vous le mise à jour de ce paramétrage ?'),
              TraduireMemoire('Méthode de valorisation aux coûts standards')) = mrYes then
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
  begin
    SetParamsoc('SO_FRAISGROUPES', False);
    SetParamsoc('SO_FRAISUNIQUE', True);
  end;

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
    SetControlVisible('LSO_PALMARESTIERS', lPalmares); SetControlVisible('SO_PALMARESTIERS' , lPalmares);

    if (not IsChecked('SO_PALMARESACTIF')) then
    begin
      SetControlChecked('SO_PALMARESENREG', False); SetParamsoc('SO_PALMARESENREG', False);
      SetControlChecked('SO_PALMARESFORCE', False); SetParamsoc('SO_PALMARESFORCE', False);
      SetControlChecked('SO_PALMARESTIERS', False); SetParamsoc('SO_PALMARESTIERS', False);
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
      SetControlEnabled('LSCO_MARCHEACH', IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEA'));
      SetControlEnabled('SCO_MARCHEACH',  IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEA'));
      SetControlEnabled('SO_CDEMARCHEACH'   , IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEA'));
      SetControlEnabled('LSO_CDEMARCHEACH'  , IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEA'));
      SetControlEnabled('SO_CDEMARCHEAPPELACH' , IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEA'));
      SetControlEnabled('LSO_CDEMARCHEAPPELACH', IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEA'));
      SetControlEnabled('SO_CDEMARCHEAPPELDATEACH' , IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEA'));
      SetControlEnabled('LSO_CDEMARCHEAPPELDATEACH', IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEA'));
      SetControlEnabled('LSCO_MARCHEVTE', IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEV'));
      SetControlEnabled('SCO_MARCHEVTE',  IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEV'));
      SetControlEnabled('SO_CDEMARCHEVTE', IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEV'));
      SetControlEnabled('LSO_CDEMARCHEVTE', IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEV'));
      SetControlEnabled('SO_CDEMARCHEAPPELVTE' , IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEV'));
      SetControlEnabled('LSO_CDEMARCHEAPPELVTE', IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEV'));
      SetControlEnabled('SO_CDEMARCHEAPPELDATEVTE' , IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEV'));
      SetControlEnabled('LSO_CDEMARCHEAPPELDATEVTE', IsChecked('SO_CDEMARCHE') and IsChecked('SO_CDEMARCHEV'));

      {Dès que la gestion des commandes marché est activée il faut forcer la gestion des conditions spéciales sur les tarifs fournisseurs}
      if (ExisteNam('SO_CDEMARCHE', FF)) and (IsChecked('SO_CDEMARCHE')) and (not GetParamSoc('SO_CDEMARCHE')) then
      begin
        if PGIAsk(TraduireMemoire('ATTENTION !!!' + #13 +
                                  'Vous allez activer la gestion des commandes marchés. ' + #13 +
                                  'La gestion des conditions spéciales dans le paramétrage '+ #13 +
                                  'des tarifs fournisseurs est obligatoire.' + #13 +
                                  'Confirmez-vous ce paramétrage ?'),
                  TraduireMemoire('Gestion des commandes marchés')) = mrYes then
        begin
          ExecuteSQL('UPDATE YTARIFSPARAMETRES SET YFO_OKSPECIAL="XXXX" WHERE YFO_FONCTIONNALITE="'+sTarifFournisseur+'" AND YFO_ORIENTATION="TIE" AND YFO_OKSPECIAL="----"');
          ExecuteSQL('UPDATE YTARIFSPARAMETRES SET YFO_OKSPECIAL="X--X" WHERE YFO_FONCTIONNALITE="'+sTarifFournisseur+'" AND YFO_ORIENTATION="ART" AND YFO_OKSPECIAL="----"');
          ExecuteSQL('UPDATE YTARIFSPARAMETRES SET YFO_OKDATE="XXXX" WHERE YFO_FONCTIONNALITE="'+sTarifFournisseur+'" AND YFO_ORIENTATION="TIE" AND YFO_OKDATE="----"');
          ExecuteSQL('UPDATE YTARIFSPARAMETRES SET YFO_OKDATE="X--X" WHERE YFO_FONCTIONNALITE="'+sTarifFournisseur+'" AND YFO_ORIENTATION="ART" AND YFO_OKDATE="----"');
          ExecuteSQL('UPDATE YTARIFSPARAMETRES SET YFO_OKSPECIAL="XXXX" WHERE YFO_FONCTIONNALITE="'+sTarifClient+'" AND YFO_ORIENTATION="TIE" AND YFO_OKSPECIAL="----"');
          ExecuteSQL('UPDATE YTARIFSPARAMETRES SET YFO_OKSPECIAL="X--X" WHERE YFO_FONCTIONNALITE="'+sTarifClient+'" AND YFO_ORIENTATION="ART" AND YFO_OKSPECIAL="----"');
          ExecuteSQL('UPDATE YTARIFSPARAMETRES SET YFO_OKDATE="XXXX" WHERE YFO_FONCTIONNALITE="'+sTarifClient+'" AND YFO_ORIENTATION="TIE" AND YFO_OKDATE="----"');
          ExecuteSQL('UPDATE YTARIFSPARAMETRES SET YFO_OKDATE="X--X" WHERE YFO_FONCTIONNALITE="'+sTarifClient+'" AND YFO_ORIENTATION="ART" AND YFO_OKDATE="----"');
        end;
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
      SetControlEnabled('LSCO_MARCHEACH', False);
      SetControlEnabled('SCO_MARCHEACH',  False);
      SetControlEnabled('SO_CDEMARCHEACH'   , False);
      SetControlEnabled('LSO_CDEMARCHEACH'  , False);
      SetControlEnabled('SO_CDEMARCHEAPPELACH' , False);
      SetControlEnabled('LSO_CDEMARCHEAPPELACH', False);
      SetControlEnabled('SO_CDEMARCHEAPPELDATEACH' , False);
      SetControlEnabled('LSO_CDEMARCHEAPPELDATEACH', False);
      SetControlEnabled('LSCO_MARCHEVTE', False);
      SetControlEnabled('SCO_MARCHEVTE',  False);
      SetControlEnabled('SO_CDEMARCHEVTE', False);
      SetControlEnabled('LSO_CDEMARCHEVTE', False);
      SetControlEnabled('SO_CDEMARCHEAPPELVTE' , False);
      SetControlEnabled('LSO_CDEMARCHEAPPELVTE', False);
      SetControlEnabled('SO_CDEMARCHEAPPELDATEVTE' , False);
      SetControlEnabled('LSO_CDEMARCHEAPPELDATEVTE', False);
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
  GereSousOrdres := IsChecked('SO_WSOUSORDRES');
  SetControlEnabled('SCO_WSSOIMPACTEWOB'       , GereSousOrdres);
  SetControlEnabled('SO_WSSOIMPACTEQTEWOBONWOL', GereSousOrdres);
  if not IsChecked('SO_WSOUSORDRES') then
    SetControlChecked('SO_WSSOIMPACTEQTEWOBONWOL', False);
  SetControlEnabled('SCO_WSSOIMPACTEWOL'       , GereSousOrdres);
  SetControlEnabled('SO_WSSOIMPACTEQTEWOLONWOB', GereSousOrdres);
  if not IsChecked('SO_WSOUSORDRES') then
    SetControlChecked('SO_WSSOIMPACTEQTEWOLONWOB', False);

  { Préférences système tarifaire : Mode SAV }
  SetControlEnabled('SO_PREFSYSTTARIF', V_PGI.Sav);

  {$IFNDEF GPAO}
    {Sous-traitance d'assemblage}
    if (ExisteNam('SO_STAGC', FF)) and (not STAGCActif) and (IsChecked('SO_STAGC')) then
    begin
      if PGIAsk(TraduireMemoire('Activation de la sous-traitance.') + #13 +
                                TraduireMemoire(' Confirmez-vous la mise à jour de la base ?'), TraduireMemoire('Module Assemblage'))= MrYes then
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
  { CCMX-CEGID le 30/01/2006 FIN }
end;
{$ENDIF GPAOLIGHT}

{***********A.G.L.***********************************************
Auteur  ...... : XMG
Créé le ...... : 31/07/2003
Modifié le ... :   /  /
Description .. : Adapte la forme aux besoins de la version Espagnole
Mots clefs ... :
*****************************************************************}
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
  if ExisteNam ('SO_ARTICLECONFIG', FF) then SetInvi('SO_ARTICLECONFIG',FF);
  if ExisteNam ('LSO_ARTICLECONFIG',FF) then SetInvi('LSO_ARTICLECONFIG',FF) ;
end;
{$ENDIF GPAO}

Function ChargePageSoc ( CC : TControl ) : boolean ;
Var FF : TForm ;
    THVCB : THValComboBox;
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
{$ENDIF GESCOM}
GereTypeTVAIntraComm(CC);
GereContremarque(CC);
GereFarFae(CC);
{$IFDEF GCGC}
SetPlusUniteParDefaut(CC);
SetGCGestUniteModeState(CC);
{$ENDIF}
{$IFDEF NETEXPERT}
InitComptaRevision (CC) ;
{$ENDIF}
if ExisteNam('SO_TARIFSENREGFLUX',FF) then
begin
  ThMultiValComboBox(GetFromNam('SO_TARIFSENREGFLUX', FF)).Complete := True;
  ThMultiValComboBox(GetFromNam('SO_TARIFSENREGFLUX', FF)).Aucun    := True;
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

//Gestion de la modif du combo type article dans intervention
if ExisteNam('SO_BTTYPEARTDEF',FF) then
   Begin
   THVCB := THValComboBox(GetFromNam('SO_BTTYPEARTDEF',FF));
   InterfaceSoc(THVCB);
   end;
       
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
   if ((DevPrinc.Value<>'') and (V_PGI.DevisePivot<>'') and (Not Devprinc.Enabled) and (Not EstMonnaieIn(DEVPRINC.Value)))
      then else BEGIN HShowMessage('20;Société;La date d''entrée en vigueur de l''euro doit être comprise entre le 1er janvier 1999 et le 31 décembre 1999.;W;O;O;O;','','') ; Result:=False ; END ;
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
Result:=True ;
END ;

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
if CH.Text='' then BEGIN HShowMessage('1;Société;Vous devez renseigner un libellé.;W;O;O;O;','','') ; CH.SetFocus ; Exit ; END ;
CH:=TEdit(GetFromNam('SO_ADRESSE1',FF)) ;
if CH.Text='' then BEGIN HShowMessage('3;Société;Vous devez renseigner une adresse.;W;O;O;O;','','') ; CH.SetFocus ; Exit ; END ;
//XVI Si localisation Espagnole et la Société est Espagnole,on vérifie le SIRET
if (VH^.PaysLocalisation=CodeISOES) and  (CodeISOduPays(THValComboBox(GetFromNam('SO_PAYS',FF)).Value)=CodeISOES) then Begin
   CH:=TEdit(GetFromNam('SO_SIRET',FF)) ;
   if CH.Text <> VerifNIF_ES(CH.Text) then BEGIN
      HShowMessage('3;Société;Le NIF/CIF saisie est incorrecte.;W;O;O;O;','','') ;
      CH.SetFocus ;
      Exit ;
   End ;
End ; //XVI 24/02/2005
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
HShowMessage('9;Société;Vous devez renseigner les fourchettes de comptes.;W;O;O;O;','','') ;
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
if Result then HShowMessage('10;Société;La fourchette de comptes que vous avez renseignée n''est pas valide : les deux comptes doivent être renseignés.;W;O;O;O;','','') ;
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

Function GenExiste ( Nam : String ; FF : TForm; Vide : boolean = false; Libelle : string = '') : boolean ;
Var CC : THCritMaskEdit ;
		messageErr : string;
    Contenu : string;
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
  Contenu := CC.text;
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
      HShowMessage('32;Intégration des écritures;Tous les champs doivent être renseignés.;W;O;O;O;','','') ;
      combo.SetFocus;
      exit;
    end;

    combo:=THValComboBox(GetFromNam('SO_IMMODOTCHOIXDET',FF));

    if assigned(combo) and (combo.Value='') then
    begin
      HShowMessage('32;Intégration des écritures;Tous les champs doivent être renseignés.;W;O;O;O;','','') ;
      combo.SetFocus;
      exit;
    end;

    combo:=THValComboBox(GetFromNam('SO_IMMODOTCHOIXTYP',FF));

    if assigned(combo) and (combo.Value='') then
    begin
      HShowMessage('32;Intégration des écritures;Tous les champs doivent être renseignés.;W;O;O;O;','','') ;
      combo.SetFocus;
      exit;
    end;

    combo:=THValComboBox(GetFromNam('SO_IMMOJALECHDEF',FF));

    if assigned(combo) and (combo.Value='') then
    begin
      HShowMessage('32;Intégration des écritures;Tous les champs doivent être renseignés.;W;O;O;O;','','') ;
      combo.SetFocus;
      exit;
    end;

    combo:=THValComboBox(GetFromNam('SO_IMMOECHCHOIXDET',FF));

    if assigned(combo) and (combo.Value='') then
    begin
      HShowMessage('32;Intégration des écritures;Tous les champs doivent être renseignés.;W;O;O;O;','','') ;
      combo.SetFocus;
      exit;
    end;

    combo:=THValComboBox(GetFromNam('SO_IMMOECHCHOIXTYP',FF));

    if assigned(combo) and (combo.Value='') then
    begin
      HShowMessage('32;Intégration des écritures;Tous les champs doivent être renseignés.;W;O;O;O;','','') ;
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

{$IFDEF GCGC}
function ControleStock (FF : TForm) : boolean;
Var
  CC : THCritMaskEdit ;
  TCB : TCheckBox;
begin
  Result := true;
  TCB := TCheckBox (GetFromNam ('SO_MODEMVTEX',FF)) ;
  if TCB.Checked then
  begin
    CC := THCritMaskEdit (GetFromNam ('SO_GCTIERSMVSTK', FF));
    if CC.Text = '' then
    begin
      HShowMessage('0;Stock;Le tiers des mouvements exceptionnels est obligatoire.;E;O;O;O;','','') ;
      SetEna ('SO_GCTIERSMVSTK', FF, true);
      Result := false;
    end;
  end;
end;

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
    if Not AuxExiste('SO_GCFOUCPTADIFF',FF,True) then Exit ;
    if Not AuxExiste('SO_GCCLICPTADIFF',FF,True) then Exit ;
    if Not AuxExiste('SO_GCFOUCPTADIFFPART',FF,True) then Exit ;
    if Not AuxExiste('SO_GCCLICPTADIFFPART',FF,True) then Exit ;
    if VH_GC.VTCSeria then  //mcd 09/06/06 champ caché si pas sérialisé, donc impossible à mettre OK
       begin
       if Not GenExiste('SO_GCECARTDEBIT',FF,True) then Exit ;
       if Not GenExiste('SO_GCECARTCREDIT',FF,True) then Exit ;
       end;
    {$IFNDEF BTP}
    if Not GenExiste('SO_GCCPTESTOCK',FF,True) then Exit;
    if Not GenExiste('SO_GCCPTEVARSTK',FF,True) then Exit;
    {$ENDIF}
    if (VH_GC.VTCSeria) then
    begin
      if Not GenExiste('SO_GCVRTINTERNE',FF,True) then Exit ;
      if Not JalExiste('SO_GCCAISSGAL',FF,True) then Exit ;
      if Not GenExiste('SO_GCECARTDEBIT',FF,True) then Exit ;
      if Not GenExiste('SO_GCECARTCREDIT',FF,True) then Exit ;
    end;
  END ;
{$ifdef STK}  //mcd 06/06/06
if ExisteNam ('SO_GCTIERSMVSTK', FF) then
begin
  if not ControleStock (FF) then exit;
end;
{$endif}
Result:=True ;
END ;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : JCF
Créé le ...... : 02/08/2001
Modifié le ... :   /  /
Description .. : Opérations à réaliser en cas de modification de la devise
Suite ........ : principale du dossier.
Mots clefs ... : MODE;MULTIDEVISE;PARAMSOC
*****************************************************************}
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

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/03/2002
Modifié le ... :   /  /
Description .. : Contrôle de l'onglet ICC
Mots clefs ... :
*****************************************************************}
Function CVerifCompteICC( CC : TControl ) : boolean ;
Var FF : TForm ;
BEGIN
// Test de la page
Result:=True ; FF:=GetLaForm(CC) ;
if Not ExisteNam('SO_ICCCOMPTECAPITAL',FF) then Exit ;
result:=CVerifCompte(THCritMaskEdit(GetFromNam('SO_ICCCOMPTECAPITAL',FF))) ;
END;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/03/2002
Modifié le ... :   /  /
Description .. : Contrôle sur l'onglet comptabilite\dates
Mots clefs ... :
*****************************************************************}
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

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 28/06/2005
Modifié le ... :   /  /
Description .. : Vérification de la cohérence du nombre de décimales des
Suite ........ : prix et quantités
Mots clefs ... :
*****************************************************************}
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

{$IFDEF COMPTA}
{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 29/05/2006
Modifié le ... :   /  /    
Description .. : LG - FB 10678 - changement du message
Mots clefs ... : 
*****************************************************************}
function CVerifAnoDyna( CC : TControl ) : boolean ;
var
 FF : TForm ;
 lBoCh : boolean ;
begin
 Result:=True ; FF:=GetLaForm(CC) ;
 if not ExisteNam('SO_CPANODYNA',FF) then exit;
 lBoCh := TCheckBox(GetFromNam('SO_CPANODYNA',FF)).Checked ;
 if ( lBoCh <> GetParamSocSecur('SO_CPANODYNA',false)) and lBoCh then
  begin
  PGIInfo('Vous venez d''activer les A-nouveaux dynamiques, une clôture provisoire va être effectuée. Validez celle-ci.') ;
  if not ClotureComptable(false,true) then
   TCheckBox(GetFromNam('SO_CPANODYNA',FF)).Checked := false
    else
     if VH^.Suivant.Code = '' then
      begin
       PGIInfo('Vous avez activé les A-nouveaux dynamiques, Voulez ouvrir l''exercice suivant') ;
       OuvertureExo ;
      end ;
  end ;

  if ( lBoCh <> GetParamSocSecur('SO_CPANODYNA',false)) and not lBoCh then
  begin
    AnnuleclotureComptable(FALSE,true);
  end ;

end ;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Guillaume FONTANA
Créé le ...... : 01/03/2005
Modifié le ... :   /  /
Description .. : Contrôle d'integrité de la saisie des ParamSoc CHR
Mots clefs ... : CHR
*****************************************************************}
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

{*****************************************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/03/2002
Modifié le ... : 21/03/2002
Description .. : LG* Ajout de la gestion de l'onglet lettrage\gestion :
Suite ........ : fonction CVerifCompteLettrage
Mots clefs ... :
*****************************************************************}
Function SauvePageSoc ( CC : TControl ) : boolean ;
BEGIN
  Result:=False ;
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
  if Not CVerifDates(CC) then Exit; // contrôle les nombres saisie dans l'onglet Comptabilite\Dates()
  if Not CVerifDivers(CC) then Exit;
  if not ExisteTVAIntracomm(CC) then exit;
  {$IFDEF COMPTA}
  if not CVerifAnoDyna(CC) then exit;
  {$ENDIF}
  {$IFDEF PGIIMMO}
    if Not VerifPGIIMMO(CC) then Exit ;
  {$ENDIF PGIIMMO}
  ModifDevPrinc(CC);
  {$IFDEF GCGC}
  if Not VerifExisteCptsGC(CC) then Exit ;
  GCChangeGestUniteMode(CC);
  {$ENDIF GCGC}
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

//PCH 11-2005 Personnalisation paramsoc
{$IFDEF GIGI}
  if Not VerifAfPerso(CC) then Exit ;
{$ENDIF GIGI}
//

GereAnaParamAff(CC);
GereTypeTVAIntraComm(CC);
GereContremarque(CC);
GereFarFae(CC, true);
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
  if not TraitePlanDeRevision( CC ) then Exit;
// ajout me pour la sychronisation
  MAJHistoparam ('PARAMSOC', '');
  {$ENDIF}

  if not VerifMES( CC ) then Exit;
  InitGPOFraisAvances(CC);

	MajTabletteGcTypeFourn(CC);
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
Result:=True ;
END ;

Procedure MarquerPublifi ( Flag : boolean ) ;
BEGIN
if ctxPCL in V_PGI.PGIContexte then else if Not VH^.LienPublifi then Exit ;
{$IFNDEF SPEC302}
if Flag then SetParamSoc('SO_ACHARGERPUBLIFI','X') else SetParamSoc('SO_ACHARGERPUBLIFI','-') ;
{$ENDIF}
END ;

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
{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. Tarcy
Créé le ...... : 30/10/2001
Modifié le ... : 30/10/2001
Description .. : Paramétrage des échanges PGI
Mots clefs ... : FO;BO;
*****************************************************************}
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

Procedure MajGPPEcheances( CC : TControl );
BEGIN
  ExecuteSQL('Update PARPIECE SET GPP_MODEECHEANCES="' + THValComboBox(CC).Value + '" where GPP_MODEECHEANCES=""');
end;

Procedure VerifTabHie ( CC : TControl );
Var FF : TForm ;
BEGIN
  FF:=GetLaForm(CC) ;
  if ExisteNam('SO_RTPROPTABHIE',FF) then TraiteTabHie(CC,'SO_RTPROPTABHIE') ;
  if ExisteNam('SO_RTACTTABHIE',FF) then TraiteTabHie(CC,'SO_RTACTTABHIE');
end;

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
Procedure CtrlNbDecimales(CC : TControl; stName : string );
Var FF : TForm;
    NbDecNew : integer;
begin
  FF := GetLaForm(CC);
  NbDecNew := StrToInt(THEdit(GetFromNam(stName, FF)).Text);
  if NbDecNew <> GetParamSoc(stName, True) then
  begin
    if PGIAsk(TraduireMemoire('ATTENTION !!!' + #13 +
                              'Ce changement peut entraîner des pertes d''informations' + #13 +
                              'Confirmez-vous votre modification ?'),
              TraduireMemoire('Nombre de décimales')) = mrNo then
    begin
        THEdit(CC).Text := IntToStr(GetParamSoc(stName, True));
    end;
  end;
end;

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
      SetEna('LSCO_MARCHEACH',FF,False);
      SetEna('SCO_MARCHEACH',FF,False);
      SetEna('LSO_CDEMARCHEACH',FF,False);
      SetEna('SO_CDEMARCHEACH',FF,False);
      SetEna('LSO_CDEMARCHEAPPELACH',FF,False);
      SetEna('SO_CDEMARCHEAPPELACH',FF,False);
      SetEna('LSO_CDEMARCHEAPPELDATEACH',FF,False);
      SetEna('SO_CDEMARCHEAPPELDATEACH',FF,False);
      SetEna('LSCO_MARCHEVTE',FF,False);
      SetEna('SCO_MARCHEVTE',FF,False);
      SetEna('LSO_CDEMARCHEVTE',FF,False);
      SetEna('SO_CDEMARCHEVTE',FF,False);
      SetEna('LSO_CDEMARCHEAPPELVTE',FF,False);
      SetEna('SO_CDEMARCHEAPPELVTE',FF,False);
      SetEna('LSO_CDEMARCHEAPPELDATEVTE',FF,False);
      SetEna('SO_CDEMARCHEAPPELDATEVTE',FF,False);
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
    end
    ;
  end;
end;

{***********A.G.L.***********************************************
Auteur ....... : Thierry Petetin
Créé le ...... : 21/01/2005
Description .. : Adapte les combos des ParamSoc lié à la pièce
Description .. : d'origine
*****************************************************************}
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

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/04/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TraitePlanDeRevision ( CC : TControl ) : Boolean;
var FF : TForm;
    vStPlanRevision : string;
begin
  Result := True;
  FF := GetLaForm(CC) ;

  if ExisteNam('SO_CPPLANREVISION', FF) then
  begin
    vStPlanRevision := TEdit(GetFromNam('SO_CPPLANREVISION', FF)).Text;

    if vStPlanRevision <> '' then
    begin
      if not AffecteToutLesCyclesRevision( vStPlanRevision ) then
        PgiError('Impossible d''affecter les cycles de révision aux comptes généraux', 'Révision');
    end
    else
      SupprimeCycleRevisionDesGene('');
  end;
end;

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

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/03/2005
Modifié le ... : 22/04/2005
Description .. : Supprime toutes les infos du champ G_CYCLEREVISION en fonction
Suite ........ : du paramètre; Si '' alors on vide le champ G_CYCLEREVISION
Mots clefs ... :
*****************************************************************}
procedure SupprimeCycleRevisionDesGene( vStCycleRevision : string = '');
var lStWhere : string;
begin
  // Mise à blanc du Champ G_CYCLEREVISION car si on restreint le cycle de révision,
  // certains comptes doivent pouvoir appartenir à un autre cycle
  if vStCycleRevision <> '' then
    lStWhere := ' WHERE G_CYCLEREVISION = "' + vStCycleREvision + '"'
  else
    lStWhere := '';

  ExecuteSQL('UPDATE GENERAUX SET G_CYCLEREVISION = ""' + lStWhere );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure MiseAJourPlanRevision;
var lQuery : TQuery;
    lStPlanRevision : string;
begin
  lStPlanRevision := GetParamSocSecur('SO_CPPLANREVISION', '');
  if lStPlanRevision = '' then Exit;

  try
    // Recherche du Plan de révision dans la TABLE CHOIXDPSTD
    lQuery := OpenSQL('SELECT YDS_CODE, YDS_LIBRE FROM CHOIXDPSTD WHERE ' +
                      'YDS_TYPE = "PDR" AND ' +
                      'YDS_CODE = "' + lStPlanRevision + '"', True);

    if not lQuery.Eof then
    begin
      if IsValidDateHeure(lQuery.FindField('YDS_LIBRE').AsString) then
      begin
        if StrToDateTime(GetParamSocSecur('SO_CPDATEMAJPLANREVISION', iDate1900)) <
           StrToDateTime(lQuery.FindField('YDS_LIBRE').AsString) then
        begin
          AffecteToutLesCyclesRevision(lStPlanRevision);
          SetParamSoc('SO_CPDATEMAJPLANREVISION', Now);
        end;
      end;
    end
    else
    begin // Le Plan de révision du PARAMSOC n'existe plus dans la table
      SetParamSoc('SO_CPPLANREVISION', '');
      SetParamSoc('SO_CPREVISBLOQUEGENE', '-');
      SetParamSoc('SO_CPDATEMAJPLANREVISION', idate1900);
      SupprimeCycleRevisionDesGene(''); // On passe G_CYCLEREVISION à ""
    end;
  finally
    Ferme(lQuery);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function AffecteCycleRevisionAuxGene( vStCycleRevision, vStCompte1, vStExclusion1 : string ) : Boolean;
var lStSql             : string;
    lStWhereCompte1    : string;
    lStWhereExclusion1 : string;
begin
  Result := False;
  try
    SupprimeCycleRevisionDesGene( vStCycleRevision );

    lStSql := '';

    lStWhereCompte1 := AnalyseCompte( vStCompte1, fbGene, False, False) ;

    lStWhereExclusion1 := AnalyseCompte( vStExclusion1, fbGene, True, False);

    lStSql := 'UPDATE GENERAUX SET G_CYCLEREVISION = "' + vStCycleRevision + '" WHERE ' +
              lStWhereCompte1;

    if lStWhereExclusion1 <> '' then
      lStSql := lStSql + ' AND ' + lStWhereExclusion1 ;

    ExecuteSql( lStSql );

    Result := True;

  except
    on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : AffecteCycleRevisionAuxGene');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/03/2005
Modifié le ... : 25/08/2005
Description .. :
Mots clefs ... :
*****************************************************************}
function AffecteToutLesCyclesRevision( vStPlanRevision : string = '') : Boolean;
var lQuery : TQuery;
begin
  Result := False;
  BeginTrans;
  try
    try
      ExecuteSQL('UPDATE GENERAUX SET G_CYCLEREVISION = ""');

      // Traitement par défaut avec le PARAMSOC
      if Trim(vStPlanRevision) = '' then
        vStPlanRevision := GetParamSoc('SO_CPPLANREVISION');

      lQuery := OpenSql('SELECT RB_RUBRIQUE, RB_COMPTE1, RB_EXCLUSION1 FROM RUBRIQUE WHERE ' +
                        'RB_NATRUB = "CPT" AND RB_CLASSERUB = "CDR" AND ' +
                        'RB_FAMILLES LIKE "%' + vStPlanRevision + ';%" AND ' +
                        '((RB_PREDEFINI = "CEG") OR (RB_PREDEFINI = "STD")) ' +
                        'ORDER BY RB_RUBRIQUE', True);
                        
      while not lQuery.Eof do
      begin
        if not AffecteCycleRevisionAuxGene( lQuery.FindField('RB_RUBRIQUE').AsString,
                                            lQuery.FindField('RB_COMPTE1').AsString,
                                            lQuery.FindField('RB_EXCLUSION1').AsString) then
          Raise Exception.Create('');

        lQuery.Next;
      end;
      CommitTrans;
      Result := True;
    except
      // Le message d'erreur apparaitra dans AffecteCycleRevisionAuxGene
      RollBack;
    end;
  finally
    Ferme( lQuery );
  end;
end;
////////////////////////////////////////////////////////////////////////////////

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

{$IFDEF GESCOM}
procedure CdeOuvTraiteAutoTransfo(CC: TControl);
begin
  if Assigned(CC) and (CC is TCheckBox) and TCheckBox(CC).Checked then
  begin
    if ExisteSQL('SELECT 1 FROM WPARAM WHERE WPA_CODEPARAM  = "' + CdeOuvAutoTransfo  + '"')
       and (PGIAsk('Voulez-vous supprimer les paramétrages utilisateurs ?', CdeOuvGetLibellePiece(CdeOuvGetCode(gtcOuverte))) = mrYes) then
    begin
      { Suppression des wParam liés à l'autoTransfo }
      ExecuteSql('DELETE FROM WPARAM WHERE WPA_CODEPARAM  = "' + CdeOuvAutoTransfo  + '"');
    end;
  end;
end;
{$ENDIF GESCOM}

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
      THValComboBox(UniteGA).Plus := '(GME_QUALIFMESURE="' + THValComboBox(QualifGA).Value + '" OR GME_QUALIFMESURE="AUC")';
    SetEna('SO_GCDEFQUALIFUNITEGA', GetLaForm(CC), GetParamSocGcGestUniteMode = '002');
    SetEna('SO_GCDEFUNITEGA', GetLaForm(CC), GetParamSocGcGestUniteMode = '002');
  end;
end;
{$ENDIF GCGC}
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

{$IFDEF GCGC}
procedure SetGCGestUniteModeState(CC: TControl);
var
  F: TForm;
begin
  F := GetLaForm(CC);
  if not ExisteNam('SO_GCGESTUNITEMODE', F) then
    Exit { Pas sur la bonne page }
  else
  begin
    {$IFDEF GPAO}
      SetEna('SO_GCGESTUNITEMODE', GetLaForm(CC), V_Pgi.Superviseur);
    {$ELSE}
      SetInvi('SO_GCGESTUNITEMODE', GetLaForm(CC));
    {$ENDIF}
  end;
end;
{$ENDIF GCGC}

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
          if PGIAsk(TraduireMemoire('Attention ! Vous avez demandé la suppression des paramètres société personnalisés.'
                          + #13 + 'Voulez-vous tout de même les conserver ?'),
                    TraduireMemoire('Paramètres société')) = mrYes then
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

end.
