{***********UNITE*************************************************
Auteur  ...... : Paul Chapuis
Créé le ...... : 27/08/2003
Modifié le ... : 27/08/2003
Description .. : Initialisation des variables globales
Suite ........ : Remplace les phases d'initialization de Delphi
Suite ........ :
Suite ........ : 27/08/2003 - initialisation des fonctions script (uses
Suite ........ : AglInitCommun)
Mots clefs ... :
*****************************************************************}
unit EntPGI;

interface

uses
  SysUtils,
  {$IFNDEF EAGLCLIENT}
    {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
    Db,
    {$IFNDEF EAGLSERVER}
      {$IFNDEF ERADIO}
        FE_Main,
      {$ENDIF !ERADIO}
    {$ENDIF EAGLSERVER}
  {$ELSE !EAGLCLIENT}
    MainEagl,
  {$ENDIF !EAGLCLIENT}
  {$IFDEF EAGLSERVER}
    eSession,
  {$ENDIF EAGLSERVER}
  HCtrls,
  uTob;

Type
  TMemoryFonction = procedure;
  tMemoryTob = Class (Tob)
    private
      fIsLoaded,
      FAvertirCache: boolean;
      fStackName,
      FSqlQuery    : string;
      TMProc    : TMemoryFonction;
    protected
      property SqlQuery: String read FSqlQuery write FSqlQuery;
    public
      procedure Load(Const AvertirCache: Boolean = false); virtual;
      procedure ClearDetail;
      procedure AvertirCache;
      Property IsLoaded: boolean Read fIsLoaded write fIsLoaded;
      constructor Create(Const StackName: string; Const ASql: string=''; LaFonction : TMemoryFonction = nil); reintroduce;
    end;
Type LaVariableEntPgi = Class
    Public
     TobTableAlertes,TobAlertes,TobTablesLiees : tMemoryTob;
     Procedure Initialisation;
    end;

{$IFDEF EAGLSERVER}
type TLaVariable = class(TObject)
public
  constructor Create; virtual;
end;
type TVHClass = class of TLaVariable;
{$ENDIF EAGLSERVER}

{$IFDEF EAGLSERVER}
  { permet l'utilisation des variables globales VH }
  function RegisterVHSession(const VHName: String; const VHClass: TVHClass): TObject;
{$ENDIF EAGLSERVER}
Procedure ChargeParamsPGI ;   // Dans ChargeMagHalley (Encore dans Ent1 )
Function OnChangeUserPGI : Boolean;
Procedure InitLaVariableEntPgi ;
Procedure LibereLaVariableEntPgi ;
{$IFNDEF EAGLSERVER}
Procedure ChargeXuelibS7 ( OldNomHalley : string);   // charge registry en reprenant une ancienne clé
{$ENDIF}
{$IFDEF GPAO}
procedure LoadTobWIT;
{$ENDIF GPAO}
Procedure LoadTableAlertes;
Procedure LoadAlertes;
Procedure LoadTablesLiees;

var ProcChargeV_PGI : TProcedure;
    VH_EntPgi : LaVariableEntPgi ;

implementation

Uses  HEnt1
      ,Ent1, M3FP
     {$IFNDEF EAGLSERVER} ,AglInitCommun {$ENDIF}// PCS pour forcer la publication des fonctions script communes
     {$IFDEF GCGC}     ,EntGC  {$ENDIF}
     {$IFDEF PAIEGRH}  ,EntPaie {$ENDIF}
     {$IFDEF GRC}      ,EntRT  {$ENDIF}
     {$IFNDEF GPAO}
     {$IFDEF GPAOLIGHT},EntGP   {$ENDIF GPAOLIGHT}
     {$ENDIF}
     {$IFDEF GPAO}     ,EntGP   {$ENDIF}
     {$IFDEF EDI}      ,EntEDI  {$ENDIF}
     {$IFDEF DP}       ,EntDP   {$ENDIF}
     {$IFDEF CJS5}     ,EntJUR  {$ENDIF}
     {$IFDEF PGIIMMO}  ,EntImo   {$ENDIF}
     {$IFDEF COMPTAAVECSERVANT} ,EntImo   {$ENDIF}
     {$IFDEF GIGI}     ,ParamSoc{$ENDIF}
     {$IFDEF CHR}
     {$IFNDEF HRPGI}
     {$IFDEF FOS5}      ,EntFO      {$ENDIF FOS5}
     {$ENDIF HRPGI}
     {$IFDEF EAGLCLIENT},UFODegrade {$ENDIF EAGLCLIENT}
     {$ENDIF CHR}
     {$IFDEF VEGA}, MenuDisp{$ENDIF} //LMO20061214 cv4 le 29/03/2006
     {$IFDEF CRM}      ,EntGP   {$ENDIF}
     {$IFDEF CHR}      ,EntChr  {$ENDIF CHR} // GF 05/12/2007
     ;

Function OnChangeUserPGI : Boolean;
Begin
Result:=False;
// PCH 13/02/2006
//
{$IFDEF GIGI}
{$IFNDEF EAGLSERVER}
  if (GetParamSocSecur('SO_AFPERSOCHOIX',False)) then
  begin
    // GESTION des PARAMETRES SOCIETE PERSONNALISES
    // Au changement d'utilisateur, il ne faut plus tenir compte de la
    // personnalisation des paramètres société de l'utilisateur précédent.
    // Il faut recharger les paramètres société à partir de la base dans la
    // V_PGI.TobSoc ==>  appel à chargemaghalley
    // Pour le nouvel utilisateur les paramètres société personnalisés sont rechargés
    // Réinitialisation de la TOB des paramètres société en mémoire
    V_PGI.TobSoc.Free;
    V_PGI.TobSoc := Nil;
    chargemaghalley;
    // Enchainement des appels :
    // CHARGEMAGHALLEY (Ent1)
    //   CHARGESOCIETEHALLEY (Ent1)
    //     ChargeInfosSociete (Ent1)
    //     ChargeParamsPGI (EntPGI)
    //       ChargeParamsGC (EntGC) et autres applications
    //         OnChangeUserGC (EntGC)
    //         ChargeParamsocPerso (EntGC) ==> chargement des paramètres personnalisés
  end
  else
  begin
    {$IFDEF GCGC}
      {$IFNDEF DOSSIERPCL}
        V_PGI.TobSoc.Free;
          {$IFNDEF EAGLSERVER}
            V_PGI.TobSoc := Nil;
          {$ENDIF EAGLSERVER}
        chargemaghalley;
      {$ENDIF}
    {$ENDIF}
    {$IFDEF GRC}
      if OnChangeUserGRC then
    {$ENDIF}
  result:= True;
  end;
{$ENDIF EAGLSERVER}
{$ELSE}
  {$IFDEF GCGC}
    {$IFNDEF DOSSIERPCL}
//  BBI Fiche 14645 Changement d'utilisateur dans la GC
      V_PGI.TobSoc.Free;
      {$IFNDEF EAGLSERVER}
      V_PGI.TobSoc := Nil;
      {$ENDIF EAGLSERVER}
      chargemaghalley;
//  BBI Fin Changement d'utilisateur dans la GC
      if OnChangeUserGC then
    {$ENDIF}
  {$ENDIF}
  {$IFDEF GRC}
    if OnChangeUserGRC then
  {$ENDIF}
  result:= True;
{$ENDIF}
  { mng 12/09/07 : pour init si changement d'utilisateur }
  VH_EntPgi.TobTableAlertes.ClearDetail;
  VH_EntPgi.TobAlertes.ClearDetail;
  VH_EntPgi.TobTablesLiees.ClearDetail;
End;

Procedure ChargeParamsPGI ;
Begin
{$IFNDEF MAJCEGID}
//{$IFDEF GCGC}   {$IFNDEF DOSSIERPCL} ChargeParamsGC ; ChargeDescriGC ; ChargeParamsCegid ;{$ENDIF}   {$ENDIF}
{$IFDEF GCGC}  {$IFDEF MAJPCL} if not V_PGI.DossierPcl then begin {$ENDIF} ChargeParamsGC ; ChargeDescriGC ; ChargeParamsCegid ;  {$IFDEF MAJPCL}end;{$ENDIF}  {$ENDIF}
{$IFDEF PAIEGRH} ChargeParamsPaie ;               {$ENDIF}
{$IFDEF GRC}     ChargeParamsRT ;                 {$ENDIF}
{$IFDEF EDI}     ChargeParamsEDI;                 {$ENDIF}
{$IFDEF PGIIMMO} ChargeParamsPGIIMMO ;            {$ENDIF}
{$IFDEF COMPTAAVECSERVANT} ChargeParamsPGIIMMO ;  {$ENDIF}
{$IFDEF CJS5}    ChargeParamsJUR ;                {$ENDIF}
{$IFDEF DP}      ChargeParamsDP ;                 {$ENDIF}
{$IFNDEF HRPGI}
  { mng 12/09/07 : pour init si changement de base }
  VH_EntPgi.TobTableAlertes.ClearDetail;
  VH_EntPgi.TobAlertes.ClearDetail;
  VH_EntPgi.TobTablesLiees.ClearDetail;
{$ENDIF HRPGI}
{$ENDIF MAJCEGID}
End;

Procedure InitVariablesPGI ;
Begin
{$IFNDEF HRPGI}
                                      InitLaVariableHalley;  // Compta pour tout le monde !
                                      InitLaVariableEntPgi;
{$IFDEF GCGC}    {$IFNDEF DOSSIERPCL} InitLaVariableGC;              {$ENDIF DOSSIERPCL} {$ENDIF GCGC}
{$IFDEF EAGLCLIENT}
{$IFDEF CHR}     {$IFDEF FOS5}        InitLaVariableFO;              {$ENDIF FOS5}       {$ENDIF CHR}
{$ENDIF EAGLCLIENT}
{$IFDEF PAIEGRH}                      InitLaVariablePaie;            {$ENDIF PAIEGRH}
{$IFDEF GRC}                          InitLaVariableRT;              {$ENDIF GRC}
{$IFDEF DP}                           InitLaVariableDP;              {$ENDIF DP}
{$IFDEF CJS5}                         InitLaVariableJUR;             {$ENDIF CJS5}
{$IFDEF PGIIMMO}                      InitLaVariablePGIIMMO;         {$ENDIF PGIIMMO}
{$IFDEF COMPTAAVECSERVANT}            InitLaVariablePGIIMMO;         {$ENDIF COMPTAAVECSERVANT}
{$IFDEF VEGA}                         InitVega;               {$ENDIF VEGA} //LMO20061214 cv4 le 29/03/2006
{$IF (Defined(GPAOLIGHT) or Defined(CRM)) and not Defined(EAGLSERVER)}
                                      VH_GP  := LaVariableGP.Create;
{$IFEND (GPAOLIGHT || CRM) && !EAGLSERVER}
{$IF Defined(EDI) and not Defined(EAGLSERVER)}
                                      VH_EDI := LaVariableEDI.Create;
{$IFEND EDI && !EAGLSERVER}
{$IFDEF CHR}                          InitLaVariableCHR;              {$ENDIF CHR}
{$ENDIF HRPGI}
End;

Procedure LibereParamsPGI ;
Begin
LibereLaVariableEntPgi;
{$IFDEF GCGC} {$IFNDEF DOSSIERPCL}  LibereLaVariableGC;         {$ENDIF DOSSIERPCL} {$ENDIF GCGC}
{$IFDEF EAGLCLIENT}
{$IFDEF CHR}  {$IFDEF FOS5}         LibereLaVariableFO;         {$ENDIF FOS5} {$ENDIF CHR}
{$ENDIF EAGLCLIENT}
{$IFDEF PAIEGRH}                    VH_Paie.Free; VH_Paie:=Nil; {$ENDIF PAIEGRH}
{$IFDEF GRC}                        LibereLaVariableRT;         {$ENDIF GRC}
{$IFDEF DP}                         LibereLaVariableDP;         {$ENDIF DP}
{$IFDEF CJS5}                       LibereLaVariableJUR;        {$ENDIF CJS5}
{$IFDEF PGIIMMO}                    LibereLaVariablePGIIMMO;    {$ENDIF PGIIMMO}
{$IFDEF COMPTAAVECSERVANT}          LibereLaVariablePGIIMMO;    {$ENDIF COMPTAAVECSERVANT}
{$IF (Defined(GPAOLIGHT) or Defined(CRM)) and not Defined(EAGLSERVER)}
  if Assigned(VH_GP) then
    FreeAndNil(VH_GP);
{$IFEND (GPAOLIGHT || CRM) && !EAGLSERVER}
{$IF Defined(EDI) and not Defined(EAGLSERVER)}
  if Assigned(VH_EDI) then
    FreeAndNil(VH_EDI);
{$IFEND EDI && !EAGLSERVER}
{$IFDEF CHR}                        LibereLaVariableCHR;        {$ENDIF CHR}
End;

{Procedure InitAGL;
Begin
If assigned(InitProcAGL) then TProcedure(InitProcAgl);
End;}

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 16/09/2005
Modifié le ... : 02/02/2006
Description .. : Permet de charger la clé de registre de la version S7 et de
Suite ........ : l'enregistrer en S5
Suite ........ :
Suite ........ : pas nécessaire en cwas Server
Mots clefs ... : REGISTRE
*****************************************************************}
{$IFNDEF EAGLSERVER}
procedure ChargeXuelibS7 ( OldNomHalley : string);
var SaveNomHalley : string;
begin
// regarde si la clé S5 existe
if GetSynRegKey('SocReference', 'AzErTy', TRUE)= 'AzErTy' then
begin
  // la clé NomHalley n'existe pas
  SaveNomHalley := NomHalley;
  SetNomHalley(oldNomHalley); ;
  // regarde si la clé S7 existe
  if GetSynRegKey('SocReference', 'AzErTy', TRUE) <> 'AzErTy' then
  begin
    // La clé OldNomHalley existe, nous allons la copier
    ChargeXuelib;
    SetNomHalley(SaveNomHalley);
    SauveXuelib;
    exit;
  end;
end;
ChargeXuelib;
end;
{$ENDIF}

{ MemoryTob }
procedure tMemoryTob.AvertirCache;
begin
  FAvertirCache := True;
  AvertirCacheServer(fStackName);
  ClearDetail();
end;

procedure tMemoryTob.ClearDetail;
begin
  inherited;

  IsLoaded := False;
end;

constructor tMemoryTob.Create(Const StackName: string; Const ASql: string = ''; LaFonction : TMemoryFonction = nil);
begin
  inherited Create(StackName, nil, -1);

  fSqlQuery  := ASql;
  fStackName := StackName;
  fIsLoaded  := false;
  TMProc     := LaFonction;
end;

procedure tMemoryTob.Load(Const AvertirCache: Boolean = false);
var
  Q: tQuery;
begin
  if not FAvertirCache then
    FAvertirCache := AvertirCache;

  if FAvertirCache then
    AvertirCacheServer(fStackName);

  if (not isLoaded) or FAvertirCache then
  begin
    if (fSqlQuery = '') and Assigned(TMProc) then
      TMProc
    else if fSqlQuery <> '' then
    begin
      Q := OpenSql(fSqlQuery, True, -1, fStackName);
      try
        if not Q.Eof then
          LoadDetailDB(fStackName, '', '', Q, False);
      finally
        Ferme(Q);
      end;
    end;
    IsLoaded := True;
    FAvertirCache := False;
  end;
end;

{$IFDEF GPAO}
procedure LoadTobWIT;
var
	TobMere, TobTmp: Tob;
  i: Integer;
begin
  { Charge les entêtes }
  VH_GP.TobWIT.LoadDetailDBFromSql('WINITCHAMPTET', 'SELECT * FROM WINITCHAMPTET ORDER BY WIT_NOMTABLE, WIT_STANDARD DESC');
  if VH_GP.TobWIT.Detail.Count > 0 then
  begin
    TobTmp := Tob.Create('OUARZAZATE', nil, -1);
    try
      { Charge les lignes dans TobTMP }
      TobTmp.LoadDetailDBFromSql('WINITCHAMPLIG', 'SELECT * FROM WINITCHAMPLIG ORDER BY WIL_NOMTABLE, WIL_INITCTXTYPE, WIL_INITCTXCODE, WIL_NUMORDRE DESC');
      { Accroche les lignes aux entêtes }
      i := TobTmp.Detail.Count;
      while (i > 0) do
      begin
        i := i - 1;
        TobMere := VH_GP.TobWIT.FindFirst(['WIT_NOMTABLE', 'WIT_INITCTXTYPE', 'WIT_INITCTXCODE'], [TobTmp.Detail[i].GetString('WIL_NOMTABLE'), TobTmp.Detail[i].GetString('WIL_INITCTXTYPE'), TobTmp.Detail[i].GetString('WIL_INITCTXCODE')], False);
        if Assigned(TobMere) then
        begin
          TobTmp.Detail[i].ChangeParent(TobMere, -1);
        end;
      end;
    finally
      TobTmp.Free;
    end;
  end;
end;
{$ENDIF GPAO}

{$IFDEF EAGLSERVER}
{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 02/11/2006
Modifié le ... :   /  /    
Description .. : Cette fonction permet de retrouver et/ou register une 
Suite ........ : variable globale type "VH_" au niveau de la session 
Suite ........ : utilisateur en mode eAglServer. Le principe est le suivant :
Suite ........ : Chaque déclaration "var VH_xx : LaVariableXX" reroutée 
Suite ........ : en define EAGLSERVER vers une fonction du même nom :
Suite ........ : "function VH_xx : LaVariableXX". Cette fonction fait appel à 
Suite ........ : "RegisterVHSession" qui soit
Suite ........ : . Ramène la varaible depuis la session
Suite ........ : . Registre/initialise/ramène la variable au niveau de la 
Suite ........ : session
Mots clefs ... : EAGLSERVER;GLOBALE;POCKET
*****************************************************************}
function RegisterVHSession(const VHName: String; const VHClass: TVHClass): TObject;
var
  Index: Integer;
  Session: TISession;
begin
  Result := nil;
  Session := LookupCurrentSession();
  if Assigned(Session) then
  begin
    Index := Session.UserObjects.IndexOf(VHName);
    if Index >= 0 then
      Result := Session.UserObjects.Objects[Index]
    else
    begin
      Result := VHClass.Create();
      Session.UserObjects.AddObject(VHName, Result);
      if VHName = 'VH_GC' then
      begin
        {$IFDEF GCGC}
        InitLaVariableGC();
        ChargeParamsGC();
        {$ENDIF GCGC}
      end
      else if VHName = 'VH_PAIE' then
      begin
        {$IFDEF PAIEGRH}
          InitLaVariablePaie();
          ChargeParamsPaie();
        {$ENDIF PAIEGRH}
      end
      else if VHName = 'VH_RT' then
      begin
        {$IFDEF GRC}
          InitLaVariableRT();
          ChargeParamsRT();
        {$ENDIF GRC}
      end
      else if VHName = 'VH_CHR' then
      begin
        {$IFDEF CHR}
          InitLaVariableCHR();
        {$ENDIF CHR}
      end
    end
  end
end;
{$ENDIF EAGLSERVER}

Procedure InitialisationAGL ;
Begin
//{$IFDEF GCGC} InitLaVariablePGI ; {$ENDIF}
If assigned(ProcChargeV_PGI) then TProcedure(ProcChargeV_PGI);
InitVariablesPGI ;
End;

{$IFDEF EAGLSERVER}
{ TLaVariable }

constructor TLaVariable.Create;
begin
  inherited Create();
end;
{$ENDIF EAGLSERVER}
Procedure InitLaVariableEntPgi ;
begin
  VH_EntPgi:=LaVariableEntPgi.Create ;
  VH_EntPgi.TobTableAlertes := tMemoryTob.Create('_YTABLEALERTES_','', LoadTableAlertes);
  VH_EntPgi.TobAlertes := tMemoryTob.Create('_YALERTES_','', LoadAlertes);
  VH_EntPgi.TobTablesLiees := tMemoryTob.Create('_YTABLEALERTELIEES_','', LoadTablesLiees);
end;

Procedure LoadTableAlertes;
begin
	VH_EntPgi.TobTableAlertes.LoadDetailDBFromSql('YTABLEALERTES', 'SELECT * FROM YTABLEALERTES');
end;

Procedure LoadAlertes;
var sql : string;
begin
  sql := 'SELECT * FROM YALERTES LEFT JOIN YALERTESUTI ON YAU_ALERTE=YAL_ALERTE'+
  ' LEFT JOIN YALERTESCOND ON YAD_ALERTE=YAL_ALERTE WHERE YAL_ACTIVE="X"'+
  ' AND YAL_DATEECHEANCE >= "'+USDateTime(Date)+'" AND YAL_DATEDEBUT <= "'+USDateTime(Date)+
  '" AND (YAU_UTIGROUPE="'+V_PGI.USER
  +'" OR YAU_UTIGROUPE="'+V_PGI.Groupe+'" OR YAU_UTIGROUPE is NULL OR YAU_TYPE="C" OR YAU_TYPE="D")'
// CRM_20080901_MNG_FQ;010;16643
  +' ORDER BY YAL_ALERTE';
  VH_EntPgi.TobAlertes.LoadDetailDBFromSql('YALERTES', sql);
end;

Procedure LoadTablesLiees;
begin
	VH_EntPgi.TobTablesLiees.LoadDetailDBFromSql('YTABLEALERTELIEES', 'SELECT * FROM YTABLEALERTELIEES');
end;

Procedure LibereLaVariableEntPgi ;
Begin
  if Assigned (VH_EntPgi) then
  begin
    FreeAndNil(VH_EntPgi.TobTableAlertes);
    FreeAndNil(VH_EntPgi.TobAlertes);
    FreeAndNil(VH_EntPgi.TobTablesLiees);
  end;
  FreeAndNil(VH_EntPgi);
end;

procedure LaVariableEntPgi.Initialisation;
begin
end;

{$IFNDEF eAGLServer}
function IsModePCL (Params: array of variant; Nb: Integer) : variant  ;//LM20070529
begin
  result := v_pgi.ModePCL='1' ;
end ;
{$ENDIF}

Initialization
  InitProcAGL := InitialisationAGL ;

{$IFNDEF eAGLServer}
  RegisterAglFunc('IsModePCL', TRUE , 1, IsModePCL);
{$ENDIF}


Finalization
  LibereParamsPGI ;
end.
