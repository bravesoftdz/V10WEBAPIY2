{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 05/04/2002
Modifié le ... : 09/04/2004
Description .. : Source TOM de la TABLE : RUPTURE (RUPTURE)
Suite ........ : GCO - Passage en EAGL
Mots clefs ... : TOM;RUPTURE
*****************************************************************}
Unit RUPTURE_TOM ;

Interface

uses
  Classes,
  Controls,
  Sysutils,
  Stdctrls,
{$IFDEF EAGLCLIENT}
  Utob,
  //eFiche,
  MaineAGL,    // AGLLanceFiche
  UtileAGL,    // PrintDBGrid
  eFichGrid,   // TFFicheGrid
{$ELSE}
  db,
  dbTables,
  Fe_main,     // AGLLanceFiche
  PrintDBG,    // PrintDBGrid
  FichGrid,    // TFFicheGrid
{$ENDIF}
  HCtrls,   // THSpinEdit
  HEnt1,
  HMsgBox,
  UTOM,
{$IFNDEF CCS3}
  HZoomSp,
{$ENDIF}
  HPanel,   // THPanel
  Htb97,    // TToolBarButton97
  Ent1;

(*
Type
  TOM_RUPTURE = Class (TOM)

    BFirst            : TToolBarButton97;
    BPrev             : TToolBarButton97;
    BNext             : TToolBarButton97;
    BLast             : TToolBarButton97;
    BDefaire          : TToolBarButton97;
    BInsert           : TToolBarButton97;
    BDelete           : TToolBarButton97;
    BImprimer         : TToolBarButton97;
    BRuptanal         : TToolBarButton97;
    BAutomate         : TToolBarButton97;
    BZoomSp           : TToolBarButton97;
    FListe            : THGrid;

    procedure OnNewRecord                ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnCancelRecord             ; override ;
    procedure OnClose                    ; override ;

    procedure OnClickBAutomate ( Sender : TObject );
    procedure OnClickBRuptanal ( Sender : TObject );
    procedure OnClickbZoomSp   ( Sender : TObject );

  private
    gszNature   : string;
    gszPlan     : string;
    gszMode     : string;
    procedure Boutons;
    function  PresenceEnchainements(var St : String) : Boolean ;
    function  GenereRupAnal : Boolean ;

end ;

Const	MsgBox: array[1..11] of string = (
		{1} 'comptes généraux',
		{2} 'comptes auxiliaires',
		{3} 'section axe 1',
		{4} 'section axe 2',
		{5} 'section axe 3',
		{6} 'section axe 4',
		{7} 'section axe 5',
		{8} 'comptes budgétaires',
 		{9} '7;Plans de ruptures des éditions;Aucun enchaînement existe pour ce plan. Vous devez le créer.;W;O;O;O;',
 		{10}'16;Plans de ruptures des éditions;Les ruptures n''ont pas été générées. La modification de l''enchaînement n''est pas validée.;W;O;O;O;',
 		{11}'6;Plans de ruptures des éditions;Désirez-vous enregistrer les modifications;Q;YNC;Y;C;');
 *)
procedure DetailPlanRupture(pszNature,pszPlan : String ; pafMode : TActionFiche);

Implementation

uses CPDRUPTUREAUTO_TOF,
     Rupanal,
     DRupanal,
     DRupGene,
     UtilPgi ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure DetailPlanRupture(pszNature,pszPlan : String ; pafMode : TActionFiche);
begin
  AGLLanceFiche('CP','CPFIGEAGL','','','');
  Exit;
  if pafMode = taModif then
    AGLLanceFiche('CP','CPDRUPTURE','','','ACTION=MODIFICATION;'+pszNature+';'+pszPlan)
  else
  begin
    if PafMode = taConsult then
      AGLLanceFiche('CP','CPDRUPTURE','','','ACTION=CONSULTATION;'+pszNature+';'+pszPlan)
    else
      AGLLanceFiche('CP','CPDRUPTURE','','','ACTION=CREATION;'+pszNature+';'+pszPlan);
  end;
end;
(*
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 09/04/2004
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_RUPTURE.OnArgument(S: String);
var lQuery : TQuery ;
    lStLibelle : string;
begin
  Inherited ;
  gszMode := ReadTokenSt(S);
  gszNature := ReadTokenSt(S);
  gszPlan := ReadTokenSt(S);

//Ajoute le popup menu
//PopUpMenu := ADDMenuPop(PopUpMenu,'','') ;

  try
    lQuery := nil;
    lStlibelle := '';
    try
      // Titre de la feuille
      lQuery := OpenSql('SELECT CC_LIBELLE From CHOIXCOD WHERE ' +
                'CC_TYPE = "' +gszNature+ '" AND CC_CODE="' +gszPlan+ '"',True);

      if not lQuery.Eof then
        lStLibelle := lQuery.FindField('CC_LIBELLE').AsString;

    except
      on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Procedre  : OnArgument (Rupture_TOM)');
    end;

  finally
    Ferme( lQuery );
  end;

  // Titre de l"ecran et HelpContext
  case gszNature[3] of
    'G': begin // Comptes Généraux
           Ecran.Caption := Ecran.Caption+MsgBox[1];
           Ecran.HelpContext := 1370100 ;
         end;
    'T': begin // Comptes Auxiliaires
           Ecran.Caption := Ecran.Caption+MsgBox[2];
           Ecran.HelpContext := 1380100 ;
         end;
    '1': begin // Section Axe 1
           Ecran.Caption:=Ecran.Caption+MsgBox[3];
           Ecran.HelpContext := 1395100 ;
         end;
    '2': begin // Section Axe 2
           Ecran.Caption:=Ecran.Caption+MsgBox[4];
           Ecran.HelpContext := 1395100 ;
         end;
    '3': begin // Section Axe 3
           Ecran.Caption:=Ecran.Caption+MsgBox[5];
           Ecran.HelpContext := 1395100 ;
         end;
    '4': begin // Section Axe 4
           Ecran.Caption:=Ecran.Caption+MsgBox[6];
           Ecran.HelpContext := 1395100 ;
         end;
    '5': begin // Section Axe 5
           Ecran.Caption:=Ecran.Caption+MsgBox[7];
           Ecran.HelpContext := 1395100 ;
         end;
    'B': begin // Comptes Budgétaires
           Ecran.Caption:=Ecran.Caption+MsgBox[8];
           Ecran.HelpContext := 0 ;
         end;
  end;
  Ecran.Caption := Ecran.Caption  + ' : ' + gszPlan + ' ' + lStLibelle;
  UpdateCaption(Ecran);

  // Récupération des composants
  BFirst    := TToolBarButton97(GetControl('BFIRST'   , True));
  BPrev     := TToolBarButton97(GetControl('BPREV'    , True));
  BNext     := TToolBarButton97(GetControl('BNEXT'    , True));
  BLast     := TToolBarButton97(GetControl('BLAST'    , True));
  BDefaire  := TToolBarButton97(GetControl('BDEFAIRE' , True));
  BInsert   := TToolBarButton97(GetControl('BINSERT'  , True));
  BDelete   := TToolBarButton97(GetControl('BDELETE'  , True));
  BImprimer := TToolBarButton97(GetControl('BIMPRIMER', True));
  BRuptanal := TToolBarButton97(GetControl('BRUPTANAL', True));
  BAutomate := TToolBarButton97(GetControl('BAUTOMATE', True));
  BZoomSp   := TToolBarButton97(GetControl('BZOOMSP'  , True));
  FListe    := THGrid(GetControl('FLISTE', True));

  // Branchement des événements
  bRuptanal.OnClick := OnClickbRuptanal;
  bAutomate.OnClick := OnClickbAutomate;
  bZoomSp.OnClick   := OnCLickbZoomSp;

  // Activation des boutons
  Boutons;

  // Longueur Max de saisie dans la colonne CODE

  //if gszNature = 'RUG' then
  //  FListe.ColLengths[3] :=VH^.Cpta[fbGene].Lg
  //else
  //  if gszNature = 'RUT' then
 //     FListe.ColLengths[3] :=VH^.Cpta[fbAux].Lg
  //    else
 //       FListe.ColLengths[3] := VH^.Cpta[AxeToFb('A' + gszNature[3])].Lg ;

end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 09/04/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOM_RUPTURE.OnLoadRecord ;
begin
  Inherited ;

  if gszMode='ACTION=CREATION' then
  begin
    gszMode := '';
    OnClickBAutomate( nil );
  end;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_RUPTURE.OnNewRecord ;
begin
  Inherited ;
  SetField('RU_NATURERUPT', gszNature) ;
  SetField('RU_PLANRUPT', gszPlan) ;
  bAutomate.Enabled := False;
  bRuptanal.Enabled := False;
  bZoomSp.Enabled   := False;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 09/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_RUPTURE.OnCancelRecord ;
begin
  Inherited ;
  Boutons;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 09/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_RUPTURE.OnUpdateRecord ;
var lSt : string;
begin
  inherited;
  lSt := GetField('RU_CLASSE');
  if (Copy(lSt, Length(lSt), 1) <> 'x') then
    SetField('RU_CLASSE', lSt + 'x');
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 09/04/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_RUPTURE.OnAfterUpdateRecord ;
begin
  Inherited ;
  Boutons;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOM_RUPTURE.OnClickbAutomate(Sender: TObject);
var St: string;
begin
    // Restreint la sélection
{$IFDEF EAGLCLIENT}
  TFFicheGrid(Ecran).FRange := 'RU_NATURERUPT='''+gszNature+''''+' AND RU_PLANRUPT='''+gszPlan+'''';
  TFFicheGrid(Ecran).ReloadDB;
{$ELSE}
  DS.Filtered := True;
  DS.Filter := 'RU_NATURERUPT='''+gszNature+''''+' AND RU_PLANRUPT='''+gszPlan+'''';
{$ENDIF}
  exit;

  // Si Général ou Auxiliaire
  if (gszNature='RUT') or (gszNature='RUG') then
  begin
    DetailPlanRuptureAuto( gszNature, gszPlan, TFFicheGrid(Ecran).Caption);
  end
  else // Si Analytique
  begin
    if VH^.Cpta[AxeToFb('A' + gszNature[3])].Structure then
      if PresenceEnChainements(St) then
        GenereRupAnal
     else
       PgiInfo('Aucun enchaînement existe pour ce plan. Vous devez le créer.', Ecran.Caption);
  end;
  DS.Refresh;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 09/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_RUPTURE.OnClickbRuptanal(Sender: TObject);
var CLibreMemo,CLibre : String ;
begin
  PresenceEnChainements(CLibreMemo) ;
  CLibre := CLibreMemo ;
  RuptureAnalytique(gszNature,gszPlan,Clibre) ;

  // S'il y a des changements dans l'enchainement
  if CLibre<>CLibreMemo then
  begin
    // Désirez-vous enregistrer les modifications ?
    if HShowMessage(MsgBox[11],'','') = mrYes then
    begin
      BeginTrans ;
      ExecuteSql('UPDATE CHOIXCOD SET CC_LIBRE="'+CLibre+'" Where CC_TYPE="'+gszNature+'" AND CC_CODE="'+gszPlan+'"') ;
      CommitTrans ;
      if not GenereRupAnal then
      begin
        if CLibreMemo<>'' then
        begin
          // Les ruptures n'ont pas été générées. La modification de l'enchaînement n'est pas validée.
          HShowMessage(MsgBox[10],'','');
          BeginTrans ;
          ExecuteSql('UPDATE CHOIXCOD SET CC_LIBRE="'+CLibreMemo+'" Where CC_TYPE="'+gszNature+'" AND CC_CODE="'+gszPlan+'"') ;
          CommitTrans ;
        end;
      end;
    end;
  end;
end;

{=======================================================================================}
procedure TOM_RUPTURE.OnCLickbZoomSp(Sender: TObject);
{$IFNDEF CCS3}
var
  UnFb : TFichierBase;
  St,St1,StSp : String ;
{$ENDIF}
begin
{$IFNDEF CCS3}
  if gszMode='taConsult' then Exit;

  UnFb:=AxeToFb('A'+gszNature[3]);

  if DS.State=dsInsert then
    St := ''
  else
    St := GetField('RU_CLASSE'); // contient le code

  if ChoisirSousPlan(UnFb,St,True,taModif) then
  begin
    if Pos('?',St)<=0 then
    begin
      if DS.State = dsBrowse then DS.Insert;
      SetField('RU_CLASSE', St+'x');
    end
    else
    begin
      StSp := '' ;
      while St <> '' do
      begin
        St1 := ReadTokenInterro(St) ;
        if St1 <> '' then
          StSp := St1 + ';' ;
      end;
      if DetailPlanRuptureAnal(gszNature,gszPlan,'A'+gszNature[3],StSp) then DS.Refresh;
    end;
  end;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/04/2004
Modifié le ... :   /  /
Description .. : Renvoi True s'il y a des enchainements analytiques
Suite ........ : Mémorise aussi l'enchainements
Mots clefs ... :
*****************************************************************}
function TOM_RUPTURE.PresenceEnchainements(var St : String) : Boolean ;
var Q : TQuery ;
begin
  Q := OpenSql('SELECT CC_LIBRE FROM CHOIXCOD WHERE CC_TYPE="'+gszNature+'" AND CC_CODE="'+gszPlan+'"',True) ;
  St:= Q.FindField('CC_LIBRE').AsString;
  Ferme(Q) ;
  Result := (St <> '') ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 09/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOM_RUPTURE.GenereRupAnal: Boolean;
begin
  Result := False ;
  if DetailPlanRuptureAnal(gszNature,gszPlan,'A'+gszNature[3],'') then
  begin
    DS.Refresh;
    Result := True ;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 09/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_RUPTURE.Boutons;
begin
  bAutomate.Enabled := True;
  bRuptanal.Enabled := ((gszNature <> 'RUT') and (gszNature <> 'RUG'));
  bZoomSp.Enabled   := ((gszNature <> 'RUT') and (gszNature <> 'RUG'));
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_RUPTURE.OnClose;
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOM_RUPTURE ] ) ;  *)
end.
