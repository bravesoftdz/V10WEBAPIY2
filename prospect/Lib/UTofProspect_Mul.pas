unit UTofProspect_Mul;

interface
Uses
     Controls,
     Classes,
     forms,
     sysutils,
     HCtrls,
     HEnt1,
     HMsgBox,
     UtilGc,
     comctrls,
     HDB,
     BtpUtil,
     ParamSoc,
     M3FP,
{$IFDEF EAGLCLIENT}
     eMul,MainEAGL,
{$ELSE}
     Mul,Fe_Main,
{$ENDIF}
{$Ifdef GIGI}
     EntGC,  ConfidentAffaire,
{$ENDIF}
{$ifdef AFFAIRE}
      UtofAfTraducChampLibre,
{$ENDIF}
{$IFDEF NOMADE}
     EntGC,
{$ENDIF}
     UTOF,UtilSelection,UtilRT,Ent1,(*Tiers360_tof,*)HTB97,variants ;
Type

{$ifdef AFFAIRE}
                //mcd 24/11/2005 pour faire affectation depuis ressource si paramétré
     TOF_Prospect_Mul = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
     TOF_Prospect_Mul = Class (TOF)
{$endif}

     public
        procedure OnArgument(Arguments : String ); override ;
        procedure OnLoad ; override ;
        procedure OnUpdate ; override ;
        procedure OnDisplay                ; override ;
    private
        xx_where,StFiltre : string ;
        TPrescris,TGroupe : THEdit;
        InsertNat	: THEdit;
        Telephone : THEdit;
        CleTel		: THedit;

        Fliste 		: THDbGrid;

        BInsert		: TToolBarButton97;

        Domaine		: string;
		    NomFiche	: String;

        procedure Binsert_OnClick(Sender: TObject);
        procedure FlisteDblClick(Sender: TObject);

        Procedure ChargementMulProspect(Arguments: String);
        procedure ChargementMulRTTiers(Arguments: String);
        procedure TelephoneChange (Sender : Tobject);
        procedure AfterShow;
        procedure B360_OnClick(Sender : tObject);
        procedure constitueXXWHEREPARC;
        procedure MnSGED_OnClick(Sender : tObject);
        procedure BVALIDSELClick (Sender : TObject);
    procedure CodeArticleRech(Sendre: TOBject);
    procedure ScreenKeYDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);

//uniquement en line
//    procedure CacheFonctionsPourLine;
    END ;

Procedure RTLanceFiche_Prospect_Mul(Nat,Cod : String ; Range,Lequel,Argument : string) ;

implementation

Uses
  UtilPGI
  , TiersUtil
  , uTOFComm
  , AglInitGC
  , UFonctionsCBP
  , FormsName
  , CommonTools
  {$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
  , AglIsoflex
  , menus
    {$ENDIF !ERADIO}
  {$ENDIF !EAGLSERVER}
{$IF Defined(NOMADE) or Defined(GRCLIGHT)}
  , ParamSoc
{$IFEND}
  ;

Procedure RTLanceFiche_Prospect_Mul(Nat,Cod : String ; Range,Lequel,Argument : string) ;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_Prospect_Mul.OnArgument(Arguments : String );
var
  F : TForm;
  ChampMul,ValMul,Critere : String;
  x : integer;
  ChgNatTiers : boolean;

  procedure AfficheRepresentant(Oui : boolean);
  begin
    if assigned(GetControl('T_REPRESENTANT')) then
    begin
      SetControlVisible('T_REPRESENTANT', Oui);
      SetControlVisible('TT_REPRESENTANT', Oui);
    end;
  end;

begin
	fMulDeTraitement := true;
inherited ;
	fTableName := 'TIERS';
  F := TForm (Ecran);
  MulCreerPagesCL(F,'NOMFIC=GCTIERS');
  if TToolbarButton97 (GetControl('BVALIDSEL')) <> nil then
  begin
    TToolbarButton97 (GetControl('BVALIDSEL')).OnClick := BVALIDSELClick;
    TToolbarButton97 (GetControl('BVALIDSEL')).Left := THValComboBox (GetControl('BCherche')).Left;
  end;
  xx_where:=RTXXWhereConfident('CON',true);
//  ChgNatTiers := False;
  {$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
      { Isoflex }
      if Assigned(GetControl('MNSGED')) then
        TMenuItem(GetControl('MNSGED')).OnClick := MnSGED_OnClick;
    {$ENDIF !ERADIO}
  {$ENDIF !EAGLSERVER}
  if Assigned(GetControl('B360')) then
    TToolbarButton97(GetControl('b360')).OnClick := B360_OnClick;

  if copy(F.Name,1,14) = 'RTPROSPECT_MUL' then
  begin
  {$IFDEF NOMADE}
    if ctxMode in V_PGI.PGIContexte then
      SetControlText('INSERTNATURE','CLI')
      else
      if GetParamSocSecur('SO_GCPREFIXETIERS','')<>'' then
        SetControlText('INSERTNATURE','PRO')
        else
        SetControlText('INSERTNATURE','0');
  {$ELSE NOMADE}
    if (Arguments <> 'GC') then
      SetControlText('INSERTNATURE','PRO')
      else
      SetControlText('INSERTNATURE','CLI');
  {$ENDIF NOMADE}
{$ifdef BUREAU} //on a maintenant cetet saisie dans le bureau.. pas le même menu
    if V_PGI.MenuCourant = 76 then //on peut créer des prospects ou des clients
{$else}
    // JTR - eQualité 11591
//    if (V_PGI.MenuCourant = 92) or (V_PGI.MenuCourant = 311) then //on peut créer des prospects ou des clients
    if ctxGRC in V_PGI.PGIContexte then //on peut créer des prospects ou des clients
{$endif BUREAU}
    begin
      if (Not AutoriseCreationTiers ('PRO')) and
{$ifdef GIGI}
        (Not (CreatTiersAutre)) and
{$endif}
         (not AutoriseCreationTiers ('CLI')) then
        SetControlVisible('BINSERT',False) ;
    end else
//    if V_PGI.MenuCourant = 30 then // on peut créer que des clients
    begin
      if (Not AutoriseCreationTiers ('CLI')) then
        SetControlVisible('BINSERT',False) ;
    end ;
{     else
    begin
      if (Not ExJaiLeDroitConcept(TConcept(gcProspectCreat),False)) and
         (Not ExJaiLeDroitConcept(TConcept(gcCLICreat),False)) then
        SetControlVisible('BINSERT',False) ;
    end; }
    // Fin JTR
  end;

  if Arguments = 'GC' then
  begin
    SetControlText ('T_NATUREAUXI_','CLI');
  {$IFDEF NOMADE}
    ChgNatTiers := VH_GC.PCPVenteSeria;
  {$ELSE NOMADE}
    ChgNatTiers := False;
  {$ENDIF NOMADE}
    SetControlEnabled ('T_NATUREAUXI_',ChgNatTiers);
    if not ChgNatTiers then
      Ecran.caption := TraduireMemoire('Liste des clients')
      else
      Ecran.caption := TraduireMemoire('Liste des clients et prospects');
    updatecaption(Ecran);
  end;
  if copy(F.Name,1,10) = 'RTTIERS_TL' then
    begin
    Tfmul(Ecran).FiltreDisabled:=true;
    TFmul(Ecran).OnAfterFormShow := AfterShow;
    Critere := Arguments ;
    Repeat
      Critere:=uppercase(ReadTokenSt(Arguments)) ;
      if Critere<>'' then
      begin
        x:=pos('=',Critere);
        if x<>0 then
           begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));
           if ChampMul='FILTRE' then StFiltre := ValMul;
           end;
      end;
    until critere='';
    end;
//uniquement en line
// Domaine	:= 'BTP';
// Nomfiche := 'BTTIERS_S1';
   Domaine	:= 'GC';
	 Nomfiche := 'GCTIERS';

  //RECUPERATION DES ZONES ECRAN
	Fliste := THDbGrid (GetControl('FLISTE'));
  Fliste.OnDblClick := FlisteDblClick;

  BInsert := TToolbarButton97(ecran.FindComponent('BInsert'));
  BInsert.OnClick := Binsert_OnClick;

  InsertNat := THEdit(Ecran.FindComponent('INSERTNATURE'));
  Telephone := THEdit(Ecran.FindComponent('TELEPHONE'));
  CleTel		:= THEdit(Ecran.FindComponent('T_CLETELEPHONE'));
  TPrescris := THEdit(GetControl('T_PRESCRIPTEUR')) ;
  TGroupe := THEdit(GetControl('T_SOCIETEGROUPE')) ;
  if F.Name='RTPROSPECT_MUL' then
  begin
    ChargementMulProspect(Arguments);
    Telephone.OnChange := TelephoneChange;
  end;

  if F.Name='RTTIERS_TL' then ChargementMulRTTiers(Arguments);
  if (GetControl('YTC_RESSOURCE1') <> nil)  then
    begin
    if not (ctxaffaire in V_PGI.PGICONTEXTE) then SetControlVisible ('PRESSOURCE',false)
    else begin
      GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '_');
      if not (ctxscot in V_PGI.PGICOntexte) then
         begin
         SetControlVisible ('T_MOISCLOTURE',false);
         SetControlVisible ('T_MOISCLOTURE_',false);
         SetControlVisible ('TT_MOISCLOTURE',false);
         SetControlVisible ('TT_MOISCLOTURE_',false);
         end;
      end;
    end;
  {$IFDEF GIGI}
   AfficheRepresentant(False);
   if (GetControl('T_ZONECOM') <> nil) then  SetControlVisible('T_ZONECOM',false);
   if (GetControl('TT_ZONECOM') <> nil) then  SetControlVisible('TT_ZONECOM',false);
   if (GetControl('TT_TARIFTIERS') <> nil) then  SetControlVisible('TT_TARIFTIERS',false);
   if (GetControl('T_TARIFTIERS') <> nil) then  SetControlVisible('T_TARIFTIERS',false);
   if (GetControl('T_CONFIDENTIEL') <> nil) then  SetControlVisible('T_CONFIDENTIEL',false);
   SetControlText('T_NatureAuxi','CLI;PRO;NCP');
   SetControlProperty ('T_NATUREAUXI', 'Complete', true);
   SetControlProperty ('T_NATUREAUXI_', 'Complete', true);
   SetControlProperty ('T_NATUREAUXI', 'Datatype', 'TTNATTIERS');
   SetControlProperty ('T_NATUREAUXI_', 'Datatype', 'TTNATTIERS');
   SetControlProperty ('T_NATUREAUXI_', 'Plus', VH_GC.AfNatTiersGRCGI);
   SetControlProperty ('T_NATUREAUXI', 'Plus', VH_GC.AfNatTiersGRCGI);
  {$ENDIF GIGI}
  {$IFDEF GCGC}
  //FQ10467 gestion des commerciaux   // PL le 18/05/07 : on affine le test
  if not (ctxaffaire in V_PGI.PGICONTEXTE) and not GereCommercial then
  begin
   AfficheRepresentant(GereCommercial);
  end;
  {$ENDIF GCGC}
{$IFDEF GRCLIGHT}
  if ( copy(TFMul(Ecran).name,1,17) = 'RTTIERS_PILOTACTC' ) and (not GetParamsocSecur('SO_CRMACCOMPAGNEMENT',False)) then
    begin
    SetControlCaption('TSOPER','Activité');
    SetControlVisible('OPERATION',false);
    SetControlVisible('LOPERATION',false);
    end;
{$ENDIF GRCLIGHT}
//uniquement en line
//  CacheFonctionsPourLine;
  if not JaiLeDroitTag(323210) then
  begin
    TTabSheet (GetControl('TBPARCSAV')).TabVisible := false;
  end;
  if ThEdit(GetControl('CODEARTICLE')) <> nil then
  begin
    ThEdit(GetControl('CODEARTICLE')).OnElipsisClick := CodeArticleRech;
  end;
  Ecran.OnKeyDown := ScreenKeYDown;
  if not Tools.CanInsertedInTable('TIERS'{$IFDEF APPSRV}, '', '' {$ENDIF APPSRV}) then
  begin
    TtoolbarButton97(GetControl('bInsert')).Visible := False;
  end;
  updatecaption(Ecran);
end;

procedure TOF_Prospect_Mul.ScreenKeYDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 120 then
  begin
    TToolbarButton97(GetControl('BVALIDSEL')).Click;
    Key := 0;
  end;
end;


procedure TOF_Prospect_Mul.OnUpdate ;
begin
  Inherited ;
  if Assigned (GetControl('B360')) then { je le met invisible pour le cas de la sélection qui ne retourne rien ..}
    TToolbarButton97(GetControl('B360')).Visible := false;
end ;

procedure TOF_Prospect_Mul.AfterShow;
begin
  Tfmul(Ecran).ForceSelectFiltre(StFiltre , V_PGI.User,false,true);
end ;

procedure TOF_Prospect_Mul.OnLoad;
begin
  inherited;
  if GetControlText ('T_NATUREAUXI_')='CON' then xx_where:='';
  SetControlText('XX_WHERE',xx_where) ;
  if (GetControlText('T_NATUREAUXI') <> 'CLI;PRO') then SetControlEnabled('T_NATUREAUXI_', False);
end;

//uniquement en line
{*
procedure TOF_Prospect_Mul.CacheFonctionsPourLine;
begin
  if TTabSheet(GetControl('TABLESLIBRES')) <> nil then TTabSheet(GetControl('TABLESLIBRES')).TabVisible := false;
  if TTabSheet(GetControl('ZONESLIBRES')) <> nil then TTabSheet(GetControl('ZONESLIBRES')).TabVisible := false;
  if TTabSheet(GetControl('PRESSOURCE')) <> nil then TTabSheet(GetControl('PRESSOURCE')).TabVisible := false;
  if TTabSheet(GetControl('PCOMPLEMENT')) <> nil then TTabSheet(GetControl('PCOMPLEMENT')).TabVisible := false;
  SetControlVisible('TT_ENSEIGNE', False);
  SetControlVisible('T_ENSEIGNE', False);
  SetControlVisible('TT_SOCIETEGROUPE', False);
  SetControlVisible('T_SOCIETEGROUPE', False);
  SetControlVisible('TT_ZONECOM', False);
  SetControlVisible('T_ZONECOM', False);
  SetControlVisible('TT_PRESCRIPTEUR', False);
  SetControlVisible('T_PRESCRIPTEUR', False);
  SetControlVisible('TT_REPRESENTANT', False);
  SetControlVisible('T_REPRESENTANT', False);
  SetControlVisible('T_NATUREAUXI_', False);
  SetControlVisible('TT_NATUREAUXI', False);
end;
*}

procedure TOF_Prospect_Mul.FlisteDblClick(Sender: TObject);
var
  Auxiliaire		: String;
  NatureAuxi		: String;
  StArg					: String;
begin
  //Si pas d'enreg double click interdit ==> FV 05112008
  {$IFDEF EAGLCLIENT}
	if FListe.RowCount = 0  then exit;
	TFMul(ecran).Q.TQ.Seek(TFMul(ecran).FListe.Row-1) ;
  Auxiliaire:=TFMul(ecran).Q.FindField('T_AUXILIAIRE').AsString;
  NatureAuxi:=TFMul(ecran).Q.FindField('T_NATUREAUXI').AsString;
  {$ELSE}
	if FListe.datasource.DataSet.RecordCount = 0  then exit;
  Auxiliaire:=Fliste.datasource.dataset.FindField('T_AUXILIAIRE').AsString;
  NatureAuxi:=Fliste.datasource.dataset.FindField('T_NATUREAUXI').AsString;
  {$ENDIF}
  OpenForm.CliPro(Auxiliaire, NatureAuxi);

(*
  stArg:='ACTION=MODIFICATION;';


  if NatureAuxi = 'PRO' then
  begin
		if Not ExJaiLeDroitConcept(TConcept(bt510),False) then stArg:= 'ACTION=CONSULTATION;';
  end else
 		if Not ExJaiLeDroitConcept(TConcept(bt511),False) then stArg:= 'ACTION=CONSULTATION;';

  if Auxiliaire <> '' then
     AGLLanceFiche ('GC',Nomfiche,'',Auxiliaire,stArg + 'MONOFICHE;T_NATUREAUXI=' + NatureAuxi);
*)     

  if GetParamSocSecur('SO_REFRESHMUL', true) then
     refreshdb;
end;

procedure TOF_Prospect_Mul.Binsert_OnClick(Sender: TObject);
var Auxiliaire		: String;
		NatureAuxi		: String;
    CodeConcept		: TConcept;
begin

  NatureAuxi := GetControlText('T_NATUREAUXI_');

  if NatureAuxi = 'PRO' then
  	CodeConcept := TConcept(gcProspectCreat)
  else if NatureAuxi = 'CLI' then
  	CodeConcept := TConcept(gcCliCreat);

  if ExJaiLeDroitConcept(CodeConcept,False) then
  begin
  	if NatureAuxi <> '' then
    begin
    	AGLLanceFiche (Domaine,Nomfiche,'','','MONOFICHE;ACTION=CREATION;T_NATUREAUXI=' + NatureAuxi);
      if GetParamSocSecur('SO_REFRESHMUL', true) then TToolBarButton97(GetCOntrol('Bcherche')).Click;
    end else
    	PGIBox('Veuillez spécifier la nature Client ou Prospect');
  end else
    PGIBox('Vous n''avez pas les droits d''accès en création');

end;


procedure TOF_Prospect_Mul.ChargementMulProspect(Arguments: String);
Var ChgNatTiers : boolean;
begin

  if (Not ExJaiLeDroitConcept(TConcept(gcProspectCreat),False)) and
     (Not ExJaiLeDroitConcept(TConcept(gcCliCreat),False)) then BINSERT.Visible :=False;

end;

procedure TOF_Prospect_Mul.ChargementMulRTTiers(Arguments: String);
Var ChampMul		: String;
    ValMul			: String;
    Critere 		: String;
    x						: integer;
Begin

  Tfmul(Ecran).FiltreDisabled:=true;
  TFmul(Ecran).OnAfterFormShow := AfterShow;

  Critere := Arguments ;

  Repeat
    Critere:=uppercase(ReadTokenSt(Arguments)) ;
    if Critere<>'' then
       begin
       x:=pos('=',Critere);
       if x<>0 then
          begin
          ChampMul:=copy(Critere,1,x-1);
          ValMul:=copy(Critere,x+1,length(Critere));
          if ChampMul='FILTRE' then StFiltre := ValMul;
          end;
       end;
  until critere='';

end;

procedure TOF_Prospect_Mul.OnDisplay;
begin
  inherited;
  if (GetField('T_AUXILIAIRE') = '') or (GetField('T_TIERS') = '') or (GetField('T_NATUREAUXI') = '') then exit;
  if Assigned (GetControl('B360')) then
    if ( (GetField('T_NATUREAUXI') = 'PRO') and (not ExJaiLeDroitConcept (TConcept(gc360Prospect), false)) ) or
       ( (GetField('T_NATUREAUXI') = 'CLI') and (not ExJaiLeDroitConcept (TConcept(gc360Client), false)) ) or
       ( (GetField('T_NATUREAUXI') <> 'PRO') and (GetField('T_NATUREAUXI') <> 'CLI') ) then
       TToolbarButton97(GetControl('B360')).Visible := false
    else
       TToolbarButton97(GetControl('B360')).Visible := true;
end;

procedure TOF_Prospect_Mul.B360_OnClick(Sender : tObject);
var stArg : String;
begin
  if (GetField('T_AUXILIAIRE') = '') or (VarIsNull(GetField('T_AUXILIAIRE'))) then exit;
  stArg:='ACTION=MODIFICATION';
  if (RTDroitModiftiers(GetField('T_TIERS'))=False) then stArg:= 'ACTION=CONSULTATION';
  AGLLanceFiche('RT','RTTIERS360','','',stArg+';AUXILIAIRE='+GetField('T_AUXILIAIRE')+';TIERS='+GetField('T_TIERS'));
end ;
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TOF_Prospect_Mul.MnSGED_OnClick(Sender : tObject);
Var
  F : TForm;
begin
  F:=TForm(Ecran) ;
  if (F.Name<>'RTPROSPECT_MUL') then Exit;
  AglIsoflexViewDoc(NomHalley, F.Name, 'TIERS', 'T_CLE1', GetParamIsoflex + ',T_AUXILIAIRE', GetField('T_AUXILIAIRE'), '');
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

procedure TOF_Prospect_Mul.TelephoneChange(Sender: Tobject);
begin
  ThEdit(getCOntrol('T_CLETELEPHONE')).Text := CleTelForFind (ThEdit(getCOntrol('TELEPHONE')).Text);
end;

procedure TOF_Prospect_Mul.constitueXXWHEREPARC;
var TheXXWhere : THEdit;
    TheWhere : string;
begin
  TheXXWhere := THEdit(GetControl('XX_WHERE1'));
  if TheXXWhere = nil then Exit;
  TheXXWhere.text := '';
  if (GetControlText('TYPEARTICLE')='') and
     (GetControlText('CODEARTICLE')='') and
     (GetControlText('REFFABRICANT')='') and
     (GetControlText('SERIE')='') and
     (GetControlText('VERSION')='') and
     (GetControlText('DATEFINSERIA')='31/12/2099') and
     (GetControlText('REFCLIENT')='') then Exit;

  TheWhere := ' AND EXISTS (SELECT BP1_ARTICLE FROM PARCTIERS WHERE BP1_TIERS=T_TIERS ';
  if GetControlText('TYPEARTICLE')<>'' then TheWhere := TheWhere+ 'AND BP1_TYPEARTICLE="'+GetControlText('TYPEARTICLE')+'" ';
  if GetControlText('CODEARTICLE')<>'' then TheWhere := TheWhere+ 'AND BP1_CODEARTICLE LIKE "'+GetControlText('CODEARTICLE')+'%" ';
  if GetControlText('REFFABRICANT')<>'' then TheWhere := TheWhere+ 'AND BP1_REFFABRICANT LIKE "'+GetControlText('REFFABRICANT')+'%" ';
  if GetControlText('REFCLIENT')<>'' then TheWhere := TheWhere+ 'AND BP1_REFCLIENT LIKE "'+GetControlText('REFCLIENT')+'%" ';
  if GetControlText('SERIE')<>'' then TheWhere := TheWhere+ 'AND BP1_SERIE LIKE "'+GetControlText('SERIE')+'%" ';
  if GetControlText('VERSION')<>'' then TheWhere := TheWhere+ 'AND BP1_CODEVERSION LIKE "'+GetControlText('VERSION')+'%" ';
  if GetControlText('DATEFINSERIA')<>'31/12/2099' then TheWhere := TheWhere+ 'AND BP1_DATEFINSERIA < "'+USDATETIME(StrToDate(GetControlText('DATEFINSERIA')))+'"';
  TheXXWhere.text := TheWhere + ' AND BP1_ETATPARC="ES")';

end;

procedure TOF_Prospect_Mul.BVALIDSELClick(Sender: TObject);
begin
  constitueXXWHEREPARC;
  BchercheClick(self);
end;

procedure TOF_Prospect_Mul.CodeArticleRech(Sendre: TOBject);
var stWhere,stFiche : string;
		ART : THCritMaskEdit;
    CODEARTICLE : Thedit;
begin
  CODEARTICLE := ThEdit(GetControl('CODEARTICLE'));
	ART := THCritMaskEdit.Create (ecran); ART.Visible := false;
  ART.Text := '';
  //
  stFiche := 'BTARTPARC_RECH';
  if GetControlText('TYPEARTICLE') <> '' then
  begin
	  stWhere := 'GA_TYPEARTICLE = "'+GetControlText('TYPEARTICLE')+'"';
  end else
  begin
	  stWhere := 'GA_TYPEARTICLE IN ("PA1","PA2")';
  end;

  StWhere := 'GA_CODEARTICLE=' + Trim (Copy (ART.Text, 1, 18))+';XX_WHERE=AND '+stWhere;
  DispatchRecherche (ART, 1, '',stWhere, stFiche);
	if ART.Text <> '' then
  	CODEARTICLE.Text := Trim(copy(ART.Text, 0, Length(ART.Text) - 1));
  ART.free;
end;

Initialization
registerclasses([TOF_Prospect_Mul]);

end.
