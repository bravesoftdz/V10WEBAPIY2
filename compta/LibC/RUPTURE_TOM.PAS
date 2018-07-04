{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 05/04/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : RUPTURE (RUPTURE)
Mots clefs ... : TOM;RUPTURE
*****************************************************************}
Unit RUPTURE_TOM ;

Interface

uses Classes, Controls,
{$IFDEF EAGLCLIENT}
	{Utob, eFiche,} MaineAGL,
{$ELSE}
	db, dbTables, Fe_main,
{$ENDIF}
	HCtrls, HEnt1, HMsgBox, UTOM, buttons, Ent1, RupAnal,
	{$IFNDEF CCS3} HZoomSp, {$ENDIF} DRupAnal, DRupture;

procedure DetailPlanRupture(pszNature,pszPlan : String ; pafMode : TActionFiche);

Type
  TOM_RUPTURE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnCancelRecord             ; override ;
  private
		gszNature   : String;
		gszPlan     : String;
    gszMode     : String;
    bRuptanal: TBitBtn;
    bAutomate: TBitBtn;
    bZoomSp: TBitBtn;
    procedure BAutomateClick(Sender: TObject);
    procedure BRuptanalClick(Sender: TObject);
    procedure BZoomSpClick(Sender: TObject);
    function PresenceEnchainements(var St : String) : Boolean ;
    function GenereRupAnal : Boolean ;
    procedure Boutons;
end ;

Const	MsgBox: array[1..11] of string = (
		{1} 'comptes généraux',
		{2}	'comptes auxiliaires',
		{3}	'section axe 1',
		{4}	'section axe 2',
		{5}	'section axe 3',
		{6}	'section axe 4',
		{7}	'section axe 5',
		{8}	'comptes budgétaires',
 		{9}	'7;Plans de ruptures des éditions;Aucun enchaînement existe pour ce plan. Vous devez le créer.;W;O;O;O;',
 		{10}'16;Plans de ruptures des éditions;Les ruptures n''ont pas été générées. La modification de l''enchaînement n''est pas validée.;W;O;O;O;',
 		{11}'6;Plans de ruptures des éditions;Désirez-vous enregistrer les modifications;Q;YNC;Y;C;');

Implementation

{=======================================================================================}
// Détail des plans de ruptures
procedure DetailPlanRupture(pszNature,pszPlan : String ; pafMode : TActionFiche);
begin
	if pafMode = taModif then
		AGLLanceFiche('CP','CPDRUPTURE','','','ACTION=MODIFICATION;'+pszNature+';'+pszPlan)
  else
		AGLLanceFiche('CP','CPDRUPTURE','','','ACTION=CREATION;'+pszNature+';'+pszPlan);
end;

{=======================================================================================}
procedure TOM_RUPTURE.OnNewRecord ;
begin
  Inherited ;
  bAutomate.Enabled := False;
	bRuptanal.Enabled := False;
	bZoomSp.Enabled := False;
end ;

{=======================================================================================}
procedure TOM_RUPTURE.OnUpdateRecord ;
begin
	Inherited ;
	if (DS.Fields.Fields[0].Value = '') then begin
  	DS.Cancel;
    DS.Refresh;
    LastError := -1;
  end;
  Boutons;
end ;

{=======================================================================================}
procedure TOM_RUPTURE.OnAfterUpdateRecord ;
begin
  Inherited ;
  Boutons;
end ;

{=======================================================================================}
procedure TOM_RUPTURE.OnLoadRecord ;
begin
  Inherited ;

  if gszMode='ACTION=CREATION' then begin
    gszMode := '';
  	BAutomateClick(Nil);
  end;
end ;

{=======================================================================================}
procedure TOM_RUPTURE.OnArgument(S: String);
var
	Q : TQuery ;
begin
  Inherited ;
	// Affecte le click des boutons
	bRuptanal := TBitBtn(GetControl('BRUPTANAL')); if (bRuptanal <> nil) then bRuptanal.OnClick := bRuptanalClick;
  bAutomate := TBitBtn(GetControl('BAUTOMATE')); if (bAutomate <> nil) then bAutomate.OnClick := bAutomateClick;
  bZoomSp   := TBitBtn(GetControl('BZOOMSP')); if (bZoomSp <> nil) then bZoomSp.OnClick := bZoomSpClick;

  gszMode := ReadTokenSt(S);
  gszNature := ReadTokenSt(S);
	gszPlan := ReadTokenSt(S);

//  	// Ajoute le popup menu
//	PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;

  // Contexte d'aide
	Case gszNature[3] of
  	'G'     : Ecran.HelpContext:=1370100 ;
    'T'     : Ecran.HelpContext:=1380100 ;
    '1'..'5': Ecran.HelpContext:=1395100 ;
    'B'     : Ecran.HelpContext:=0 ;
   	else      Ecran.HelpContext:=0 ;
	end;

	// Titre de la feuille
	Q:=OpenSql('SELECT CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="'+gszNature+'" AND CC_CODE="'+gszPlan+'"',True) ;
	Case gszNature[3] of
  	'G':Ecran.Caption:=Ecran.Caption+MsgBox[1]+':'+gszPlan+' '+Q.Fields[0].AsString; // comptes généraux
    'T':Ecran.Caption:=Ecran.Caption+MsgBox[2]+':'+gszPlan+' '+Q.Fields[0].AsString; // comptes auxiliaires
    '1':Ecran.Caption:=Ecran.Caption+MsgBox[3]+':'+gszPlan+' '+Q.Fields[0].AsString; // sections axe 1
    '2':Ecran.Caption:=Ecran.Caption+MsgBox[4]+':'+gszPlan+' '+Q.Fields[0].AsString; // sections axe 2
    '3':Ecran.Caption:=Ecran.Caption+MsgBox[5]+':'+gszPlan+' '+Q.Fields[0].AsString; // sections axe 3
    '4':Ecran.Caption:=Ecran.Caption+MsgBox[6]+':'+gszPlan+' '+Q.Fields[0].AsString; // sections axe 4
    '5':Ecran.Caption:=Ecran.Caption+MsgBox[7]+':'+gszPlan+' '+Q.Fields[0].AsString; // sections axe 5
    'B':Ecran.Caption:=Ecran.Caption+MsgBox[8]+':'+gszPlan+' '+Q.Fields[0].AsString; // comptes budgétaires
	end;
	Ferme(Q) ;
  UpdateCaption(Ecran);

	// Restreint la sélection
  DS.Filtered := True;
  DS.Filter := 'RU_NATURERUPT='''+gszNature+''''+' AND RU_PLANRUPT='''+gszPlan+'''';
//	DS.Cancel;

	// Si ce n'est pas une rupture sur axe analytique : cache les boutons
  Boutons;
  //Enchainements des ruptures analytiques
//	bRuptanal.Enabled:=((gszNature<>'RUT') AND (gszNature<>'RUG')) ;
  // Génération automatique des ruptures à partir des sous plans analytiques
//	bZoomSp.Enabled:=((gszNature<>'RUT') AND (gszNature<>'RUG')) ;
end ;

{=======================================================================================}
procedure TOM_RUPTURE.OnCancelRecord ;
begin
  Inherited ;
  Boutons;
end ;

{=======================================================================================}
procedure TOM_RUPTURE.BAutomateClick(Sender: TObject);
var
	St: String;
begin
	// Si Généraux ou Auxiliaire
	if (gszNature='RUT')OR(gszNature='RUG') then begin
  	// Affiche l'autre feuille
    DetailPlanRuptureBis(gszNature,gszPlan,Ecran.Caption);
		end
	// Si Analytique
  else begin
    if VH^.Cpta[AxeToFb('A'+gszNature[3])].Structure then
			if PresenceEnChainements(St) then GenereRupAnal
		                               // Aucun enchaînement existe pour ce plan. Vous devez le créer.
    		                           else HShowMessage(MsgBox[9],'','');
	end;

  // Rafraichit l'affichage
	DS.Refresh;
end;

{=======================================================================================}
procedure TOM_RUPTURE.BRuptanalClick(Sender: TObject);
var
	CLibreMemo,CLibre : String ;
begin
	PresenceEnChainements(CLibreMemo) ;
  CLibre:=CLibreMemo ;
	RuptureAnalytique(gszNature,gszPlan,Clibre) ;

  // S'il y a des changements dans l'enchainement
	if CLibre<>CLibreMemo then begin
    // Désirez-vous enregistrer les modifications ?
    if HShowMessage(MsgBox[11],'','') = mrYes then begin
      BeginTrans ;
      ExecuteSql('UPDATE CHOIXCOD SET CC_LIBRE="'+CLibre+'" Where CC_TYPE="'+gszNature+'" AND CC_CODE="'+gszPlan+'"') ;
      CommitTrans ;
      if not GenereRupAnal then begin
        if CLibreMemo<>'' then begin
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
procedure TOM_RUPTURE.BZoomSpClick(Sender: TObject);
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
  	St:=''
  else
  	St := DS.Fields.Fields[2].Value;	// contient le code

  if ChoisirSousPlan(UnFb,St,True,taModif) then begin
   	if Pos('?',St)<=0 then begin
			if DS.State=dsBrowse then DS.Insert;
      DS.Fields.Fields[2].Value := St+'x';
      end
    else begin
      StSp:='' ;
      While St<>'' do begin
        St1:=ReadTokenInterro(St) ;
        if St1<>'' then StSp:=St1+';' ;
      end;
      if DetailPlanRuptureAnal(gszNature,gszPlan,'A'+gszNature[3],StSp) then DS.Refresh;
    end;
	end;
{$ENDIF}
end;

{=======================================================================================}
// Renvoi True s'il y a des enchainements analytiques
// Mémorise aussi l'enchainements
function TOM_RUPTURE.PresenceEnchainements(var St : String) : Boolean ;
var
	Q : TQuery ;
begin
	Q:=OpenSql('SELECT CC_LIBRE FROM CHOIXCOD WHERE CC_TYPE="'+gszNature+'" AND CC_CODE="'+gszPlan+'"',True) ;
	St:=Q.Fields[0].AsString ;
  Ferme(Q) ;
  Result:=(St<>'') ;
end;

{=======================================================================================}
function TOM_RUPTURE.GenereRupAnal: Boolean;
begin
	Result:=False ;
	if DetailPlanRuptureAnal(gszNature,gszPlan,'A'+gszNature[3],'') then begin
  	DS.Refresh;
   	Result:=True ;
	end;
end;

{=======================================================================================}
procedure TOM_RUPTURE.Boutons;
begin
	bAutomate.Enabled := True;
  bRuptanal.Enabled:=((gszNature<>'RUT') AND (gszNature<>'RUG'));
	bZoomSp.Enabled:=((gszNature<>'RUT') AND (gszNature<>'RUG'));
end;

{=======================================================================================}
Initialization
  registerclasses ( [ TOM_RUPTURE ] ) ;
end.
