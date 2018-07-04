unit UTomOperation;

interface

uses Controls,Classes,AglIsoflex,
{$IFDEF EAGLCLIENT}
     MaineAGL,eFiche,UtileAGL,
{$ELSE}
     Fe_Main,Fiche, db,DBCtrls,
{$ENDIF}
     forms,sysutils,
     menus,
     HCtrls,HEnt1,HMsgBox,UTOM,
     UTob,M3FP,UtilGC,windows  ;

Type
    TOM_OPERATIONS = Class (TOM)
    Private
        { mng_fq012;10859 }
        AsInsert : boolean;
        SOrtgestinfos002,SoRTLienOperation,NoSupprLien: boolean;
        stProduitpgi : string;
        procedure RTOpeAppelParamCL(bCreation : boolean);
        // Gestion Isoflex
        procedure GereIsoflex;
        procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
        procedure MnPiecesLiees_OnClick(Sender: TObject);
  {$IFNDEF EAGLSERVER}
        function TestAlerteOPE (CodeA : String) : boolean;
        procedure ListAlerte_OnClick_ROP(Sender: TObject);
        procedure Alerte_OnClick_ROP(Sender: TObject);
  {$ENDIF EAGLSERVER}
    Public
        procedure OnUpdateRecord ; override ;
        Procedure OnDeleteRecord ; override ;
        procedure OnLoadRecord ; override ;
        procedure OnLoadAlerte  ; override ;
        procedure OnArgument (Arguments : String )  ; override ;
        procedure OnNewRecord ; override ;
        procedure OnAfterUpdateRecord ; override ;
END ;
        function AffecteCodeAuto : String;

const
	// libellés des messages
	TexteMessage: array[1..5] of string 	= (
          {1}        'La date de fin doit être postérieure ou égale à la date de début'
          {2}        ,'Vous devez renseigner la date de début'
          {3}        ,'Vous ne pouvez pas effacer cette opération, elle est actuellement utilisée dans les actions.'
          {4}        ,'Vous ne pouvez pas effacer cette opération, elle est actuellement utilisée dans les propositions.'
          {5}        ,'Vous ne pouvez pas effacer cette opération, elle est actuellement utilisée dans les pièces.'
          );

implementation
Uses UtofRTPiecesOper,ParamSoc,EntRT,UtilConfid,BTPUtil,HrichOle
{$IFNDEF EAGLSERVER}
  ,UtilAlertes,YAlertesConst,EntPgi
{$ENDIF EAGLSERVER}
;
procedure TOM_OPERATIONS.OnArgument (Arguments : String );
var Critere,ChampMul,ValMul : string;
    x : integer;
begin
inherited ;
  AppliqueFontDefaut (THRichEditOle(GetControl('ROP_BLOCNOTE')));
  { mng_fq012;10859 }
  AsInsert := false;
  stproduitpgi := '';
  NoSupprLien := False;
  Repeat
      Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
      if Critere<>'' then
      begin
          x:=pos('=',Critere);
          if x<>0 then
          begin
             ChampMul:=copy(Critere,1,x-1);
             ValMul:=copy(Critere,x+1,length(Critere));
             if ChampMul='PRODUITPGI' then
                stProduitpgi := ValMul;
          end;
          if Critere = 'NOSUPPRLIEN' then NoSupprLien := True;
      end;
  until  Critere='';
  if stproduitpgi = '' then stproduitpgi := 'GRC';
  if stProduitpgi = 'GRF' then
    begin
    Sortgestinfos002:=false;
    TPopupMenu(GetControl('MENUZOOM')).items[2].visible:=false; { infos compl. }
    SetControlVisible ('ROP_OBJETOPE',false);
    SetControlVisible ('TROP_OBJETOPE',false);
    SetControlVisible ('PTABLESLIBRES',false);
    Ecran.Caption:= Traduirememoire('Opération :');
    updatecaption(Ecran);
    SoRTLienOperation:=False;
    end
  else
    begin
    SOrtgestinfos002:=GetParamSocSecur('SO_RTGESTINFOS002',False);
    SetControlVisible ('ROP_OBJETOPEF',false);
    SetControlVisible ('TROP_OBJETOPEF',false);
    SoRTLienOperation:=GetParamSocSecur('SO_RTLIENOPERATIONPIECE',False);
    end;
  SetControlText ('PRODUITPGI',stProduitpgi);
  GereIsoflex;
  if ecran is TFFiche then
    TFFiche(Ecran).OnKeyDown:=FormKeyDown ;
  if Assigned(GetControl('mnPiecesLiees')) then
      TMenuItem(GetControl('mnPiecesLiees')).OnClick := MnPiecesLiees_OnClick;


      
// Pl le 19/05/07 : gestion des champs libres seulement pour KPMG pour l'instant
// Pl le 14/08/07 : FQ10622 on veut tous les nouveaux champs libre maintenant
//  if (GetParamSocSecur ('SO_AFCLIENT', 0) = 8 ) then
    GCMAJChampLibre(TForm(Ecran), False, 'COMBO', 'ROP_ROPTABLELIBRE', 5, '_');
//  else
//    SetControlVisible('PTABLESLIBRES', false);
{$IFNDEF EAGLSERVER}
  if Assigned(GetControl('MnAlerte')) then
    if AlerteActive('ROP') then
      TMenuItem(GetControl('MnAlerte')).OnClick := Alerte_OnClick_ROP
    else
      TMenuItem(GetControl('MnAlerte')).visible:=false;

  if Assigned(GetControl('MnListAlerte')) then
    if AlerteActive('ROP') then
      TMenuItem(GetControl('MnListAlerte')).OnClick := ListAlerte_OnClick_ROP
    else
      TMenuItem(GetControl('MnListAlerte')).visible:=false;

  if Assigned(GetControl('MnGestAlerte')) and Assigned(GetControl('MnAlerte'))
     and Assigned(GetControl('MnListAlerte')) then
         TMenuItem(GetControl('MnGestAlerte')).visible := (TMenuItem(GetControl('MnAlerte')).visible)
          and (TMenuItem(GetControl('MnListAlerte')).visible);
{$ENDIF EAGLSERVER}

end;

procedure TOM_OPERATIONS.OnNewRecord;
begin
inherited;
  { mng_fq012;10859 }
  AsInsert := true;
  if stProduitpgi='GRF' then
   SetField('ROP_PRODUITPGI','GRF')
  else
   SetField('ROP_PRODUITPGI','GRC');
end;

procedure TOM_OPERATIONS.OnLoadAlerte;
begin
  {$IFNDEF EAGLSERVER}
  if (not V_Pgi.SilentMode) and (AlerteActive('ROP')) and (not AfterInserting )  then
     TestAlerteOpe(CodeOuverture+';'+CodeDateAnniv);
  {$ENDIF !EAGLSERVER}
end;

procedure TOM_OPERATIONS.OnLoadRecord;
begin
inherited;
if (GetParamsocSecur('SO_RTNUMOPERAUTO',False)  = TRUE) then
   begin
    SetControlEnabled('ROP_OPERATION', FALSE);
    if ds.state in [dsinsert] then SetFocusControl('ROP_LIBELLE');
   end;
if (SoRTLienOperation = False) or
   (GetField ('ROP_OPERATION') = 'MODELES D''ACTIONS') or
   (DS.State = dsInsert) then
  TMenuItem(GetControl('mnPiecesLiees')).Visible := False
else
  TMenuItem(GetControl('mnPiecesLiees')).Visible := True;

if (stProduitpgi = 'GRF') or
   (GetField ('ROP_OPERATION') = 'MODELES D''ACTIONS') or
   (DS.State = dsInsert) then
  TMenuItem(GetControl('mnAnalyses')).Visible := False
else
  TMenuItem(GetControl('mnAnalyses')).Visible := True;
AppliquerConfidentialite(Ecran,stProduitpgi);
end;

procedure TOM_OPERATIONS.OnUpdateRecord;
var CodeOper, NumChrono : String;
begin
inherited;
if GetField ('ROP_DATEDEBUT') = iDate1900 then
   begin
   Lasterror:=2;
   LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
   SetFocusControl('ROP_DATEDEBUT') ;
   exit;
   end;    
if GetField ('ROP_DATEFIN') < GetField ('ROP_DATEDEBUT') then
   begin
   Lasterror:=1;
   LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
   SetFocusControl('ROP_DATEFIN') ;
   exit;
   end;

{$IFNDEF EAGLSERVER}
  if  (Assigned(Ecran)) and (not V_Pgi.SilentMode) and (AlerteActive(TableToPrefixe(TableName))) then
       if (not Inserting) then
         begin
         if not TestAlerteOpe (CodeModification+';'+CodeModifChamps) then
           begin
           LastError:=99;
           exit;
           end;
         end
       else
         if not TestAlerteOpe (CodeCreation) then
           begin
           LastError:=99;
           exit;
           end;
{$ENDIF !EAGLSERVER}

if GetParamsocSecur('SO_RTNUMOPERAUTO',False) = TRUE then
  begin
  if (ds<>nil) and (ds.state = dsinsert) and (GetField ('ROP_OPERATION')= '') then
     begin
//     CodeOper:=AttribNewCode('OPERATIONS','ROP_OPERATION',0,'',GetParamsocSecur('SO_RTCOMPTEUROPER',''),'');
     CodeOper           := AffecteCodeAuto;   //TJA 27/06/2007 gestion code auto avec préfixage du code user
     SetField ('ROP_OPERATION', CodeOper);
{$IFDEF EAGLCLIENT}
     SetControlText('ROP_OPERATION',CodeOper);
{$ENDIF}
     NumChrono:=ExtraitChronoCode(CodeOper);
     SetParamSoc('SO_RTCOMPTEUROPER', NumChrono) ;
     end;
  end;
end;

procedure TOM_OPERATIONS.OnAfterUpdateRecord;
begin
inherited;
  if (DS<>nil) and (  { mng_fq012;10859 }  AsInsert ) and SOrtgestinfos002 then
    RTOpeAppelParamCL(true);
   AsInsert :=false;
end;

procedure TOM_OPERATIONS.OnDeleteRecord  ;
begin
Inherited ;

if ExisteSQL('SELECT RAC_LIBELLE FROM ACTIONS WHERE RAC_OPERATION="'
     +GetField('ROP_OPERATION')+'" ') then
   BEGIN
   LastError:=3;
   LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
   exit ;
   end ;
if ExisteSQL('SELECT RPE_LIBELLE FROM PERSPECTIVES WHERE RPE_OPERATION="'
     +GetField('ROP_OPERATION')+'" ') then
   BEGIN
   LastError:=4;
   LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
   exit ;
   end ;
if ExisteSQL('SELECT GP_NATUREPIECEG FROM PIECE WHERE GP_OPERATION="'
     +GetField('ROP_OPERATION')+'" ') then
   BEGIN
   LastError:=5;
   LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
   exit ;
   end ;
  if (not V_Pgi.SilentMode) and (not AfterInserting)
    and (AlerteActive (TableToPrefixe(TableName))) then
      if (not TestAlerteOpe(CodeSuppression)) then
        begin
        LastError := 99;
        exit;
        end;

if SOrtgestinfos002 = true then
   ExecuteSQL('DELETE FROM RTINFOS002 where RD2_CLEDATA="'+GetField('ROP_OPERATION')+'"') ;
ExecuteSQL('DELETE FROM ACTIONSGENERIQUES where RAG_OPERATION="'+GetField('ROP_OPERATION')+'"') ;
end;

procedure TOM_OPERATIONS.MnPiecesLiees_OnClick(Sender: TObject);
var stArg : string;
begin
stArg := 'ACTION=MODIFICATION';
if TFFiche(ecran).TypeAction = TaConsult then stArg := 'ACTION=CONSULTATION';
if NoSupprLien = True then stArg := stArg+';NOSUPPRLIEN;'; 
if (GetField('ROP_OPERATION') <> '') then
   RTLanceFiche_RTPiecesOper ('RT', 'RTMULPIECESOPER','GP_OPERATION='+GetField('ROP_OPERATION'), '', stArg);
end;

procedure TOM_OPERATIONS.RTOpeAppelParamCL(bCreation : boolean);
var TobChampsProFille : tob;
  { mng_fq012;10859 }
  bcreat:boolean;
begin
  { mng_fq012;10859 }
  bcreat:=false;
  if ds.state = dsInsert then bcreat:=true;
  VH_RT.TobChampsPro.Load;

  TobChampsProFille:=VH_RT.TobChampsPro.FindFirst(['CO_CODE'], ['2'], TRUE);
  if bCreation then
    begin
    if not Assigned(TobChampsProFille.FindFirst(['RDE_DESC','RDE_OBLIGATOIRE'],['2','X'],TRUE)) then exit;
    end
  else
    if (TobChampsProFille = Nil ) or (TobChampsProFille.detail.count = 0 ) then
      begin
      PGIInfo('Le paramétrage de cette saisie n''a pas été effectué','');
      exit;
      end
    else
      { mng_fq012;10858 }
      if not TFFiche(Ecran).Bouge(TNavigateBtn(nbPost)) then exit ;
  { mng_fq012;10859 }
  if (not bcreat) or (not Assigned(TobChampsProFille.FindFirst(['RDE_DESC','RDE_OBLIGATOIRE'],['2','X'],TRUE))) then
    AglLancefiche('RT','RTPARAMCL','','','FICHEPARAM='+TFFiche(Ecran).Name+';FICHEINFOS='+GetField('ROP_OPERATION')) ;
end;

procedure AGLRTOpeAppelParamCL( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_OPERATIONS) then TOM_OPERATIONS(OM).RTOpeAppelParamCL(false) else exit;
end;

// *****************************************************************************
// ********************** gestion Isoflex **************************************
// *****************************************************************************

procedure TOM_OPERATIONS.GereIsoflex;
var bIso: Boolean;
  MenuIso: TMenuItem;
begin
  MenuIso := TMenuItem(GetControl('mnSGED'));
  bIso := AglIsoflexPresent;
  if MenuIso <> nil then MenuIso.Visible := bIso;
end;

procedure Rop_AppelIsoFlex(parms: array of variant; nb: integer);
var F: TForm;
  Cle1: string;
begin
  F := TForm(Longint(Parms[0]));
  //GP_DS_GP14537_20071226
  if not pos('RTOPERATIONS', UpperCase(F.Name)) = 1 then exit;
  Cle1 := string(Parms[1]);
  AglIsoflexViewDoc(NomHalley, F.Name, 'OPERATIONS', 'ROP_CLE1', 'ROP_OPERATION', Cle1, '');
end;
procedure TOM_OPERATIONS.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
Case Key of
    VK_F6 : {Infos compl.} if (ssAlt in Shift) then
            if (soRtgestinfos002 = True) then RTOpeAppelParamCL(false) ;
    VK_F7 : {actions } if (ssAlt in Shift) then
            begin
            if GetField('ROP_OPERATION') <> '' then
               if stProduitpgi = 'GRF' then AglLancefiche('RT','RFACTIONS_MUL_OPE','RAC_OPERATION='+GetField('ROP_OPERATION'),'','')
               else AglLancefiche('RT','RTACTIONS_MUL_OPE','RAC_OPERATION='+GetField('ROP_OPERATION'),'','') ;
            end;
{$IFNDEF EAGLSERVER}
    81 : {Ctrl + Q - Création d'1 alerte} if (ssCtrl in Shift) then
          begin
          Key:=0 ;
          Alerte_OnClick_ROP(Sender);
          end;
    85 : {Ctrl + U - liste des alertes du tiers} if (ssCtrl in Shift) then
          begin
          Key:=0 ;
          ListAlerte_OnClick_ROP(Sender);
          end;
{$ENDIF EAGLSERVER}

    END;
if (ecran <> nil) then
if ecran is TFFiche then
   TFFiche(ecran).FormKeyDown(Sender,Key,Shift);
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 27/06/2007
Modifié le ... :   /  /    
Description .. : attribution du code automatique de l'opération avec si 
Suite ........ : gestion du préfixage du code user
Mots clefs ... :
*****************************************************************}
function AffecteCodeAuto: String;
begin
  if GetParamSocSecur('SO_RTPREFIXEAUTOOPERATION', False, True) then
    Result              := AttribNewCode('OPERATIONS', 'ROP_OPERATION', 17, V_PGI.User, GetParamsocSecur('SO_RTCOMPTEUROPER', '', True), '')
  else
    Result              := AttribNewCode('OPERATIONS', 'ROP_OPERATION', 0, '',GetParamsocSecur('SO_RTCOMPTEUROPER', '', True), '');
end;

{$IFNDEF EAGLSERVER}
function TOM_OPERATIONS.TestAlerteOpe (CodeA : String) : boolean;
var TOBInfosCompl : tob;
    i : integer;
    F: TForm;
begin
  result:=true;
  { cas comme la duplication ou l'on passe dans le loadalerte alors que les champs ne sont pas renseignés }
  if (GetField('ROP_OPERATION') = '' ) then exit;
  F:=TForm(Ecran);
  if not (F is TfFiche) then exit;
  if assigned( TfFiche(F).TobFinale) then TfFiche(F).TobFinale.free;
  TfFiche(F).TobFinale:=TOB.create ('OPERATIONS',NIL,-1);

  TfFiche(F).TobFinale.GetEcran (TFfiche(Ecran),Nil);

  { si passage du load, on sauvegarde les tobs initiales }
  if pos(CodeOuverture,CodeA) > 0 then
    begin
    if assigned( TfFiche(F).TobOrigine) then TfFiche(F).TobOrigine.free;
    TfFiche(F).TobOrigine:=TOB.create ('OPERATIONS',NIL,-1);
    TfFiche(F).TobOrigine.GetEcran (F,Nil,true);
    end;

  if GetParamSocSecur('SO_RTGESTINFOS002',false) then
    begin
    TOBInfosCompl:= TOB.Create('RTINFOS002', nil, -1);
    if CodeA<>CodeCreation then
      TOBInfosCompl.selectDB ('"'+GetField('ROP_OPERATION')+'"',nil);
    for i := 1 to Pred(TOBInfosCompl.NbChamps) do
      TfFiche(F).TobFinale.AddChampSupValeur(TOBInfosCompl.GetNomChamp(i), TOBInfosCompl.GetValue(TOBInfosCompl.GetNomChamp(i)));
    if (pos(CodeOuverture,CodeA) > 0) then
      begin
      if (pos(CodeOuverture,CodeA) > 0) and assigned( TfFiche(F).TobOrigine) then
        for i := 1 to Pred(TOBInfosCompl.NbChamps) do
          TfFiche(F).TobOrigine.AddChampSupValeur(TOBInfosCompl.GetNomChamp(i), TOBInfosCompl.GetValue(TOBInfosCompl.GetNomChamp(i)));
      end;
    TOBInfosCompl.free;
    end;
  if pos(CodeOuverture,CodeA) > 0 then
    result:=ExecuteAlerteLoad(F,false)
  else
    if pos(CodeSuppression,CodeA) > 0 then
      result:=ExecuteAlerteDelete(F,false)
    else
      result:=ExecuteAlerteUpdate(F,false);
end;

{ GC/GRC : MNG / gestion des alertes }
procedure TOM_OPERATIONS.ListAlerte_OnClick_ROP(Sender: TObject);
begin
if (GetField('ROP_OPERATION') <> '') and(AlerteActive(TableToPrefixe(TableName))) then
   AGLLanceFiche('Y','YALERTES_MUL','YAL_PREFIXE=ROP','','ACTION=CREATION;MONOFICHE;CHAMP=ROP_OPERATION;VALEUR='
      +GetField('ROP_OPERATION')+';LIBELLE='+GetField('ROP_LIBELLE')) ;
end ;

procedure TOM_OPERATIONS.Alerte_OnClick_ROP(Sender: TObject);
begin
if (GetField('ROP_OPERATION') <> '') and(AlerteActive(TableToPrefixe(TableName))) then
   AGLLanceFiche('Y','YALERTES','','','ACTION=CREATION;MONOFICHE;CHAMP=ROP_OPERATION;VALEUR='
   +GetField('ROP_OPERATION')+';LIBELLE='+GetField('ROP_LIBELLE')) ;
VH_EntPgi.TobAlertes.ClearDetail;
end;
{$ENDIF !EAGLSERVER}

Initialization
registerclasses([TOM_OPERATIONS]) ;
RegisterAglProc( 'RTOpeAppelParamCL', True,0,AGLRTOpeAppelParamCL) ;
RegisterAglProc('Rop_AppelIsoFlex', TRUE, 1, Rop_AppelIsoFlex);
end.
