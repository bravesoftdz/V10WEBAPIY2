unit UTofGCFournisseur_Mul;

interface

Uses Controls,
     Classes,
     forms,
     sysutils,
     HCtrls,
     HEnt1,
     HTB97,
     HMsgBox,
     HDB,
     BtpUtil,
{$IFDEF EAGLCLIENT}
     eMul,MainEAGL,
{$ELSE}
     Mul,Fe_Main,
{$ENDIF}
     Menus,
{$ifdef GIGI}
     UtofAfTraducChampLibre,
{$else}
     uTofComm, //mcd 04/12/2006 UtofAftraducCHampLibre hÃ©rite dÃ©jÃ  de UtofComm
{$ENDIF}
     UTOF,UtilSelection,UtilRT,ParamSoc,UtilGC;
Type

{$ifdef GIGI}
    TOF_Fournisseur_Mul = Class (TOF_AFTRADUCCHAMPLIBRE)
{$else}
    TOF_Fournisseur_Mul = Class (tTOFCOMM)
{$endif}

     public
        procedure OnArgument(Arguments : String ); override ;
        procedure OnLoad ; override ;
        procedure OnClose ; override ;
     private
        xx_where,StFiltre : string ;
        ModuleQualite     : boolean;
        Selection         : boolean;

        Domaine				: string;
		    NomFiche			: String;
        Action        : string;
        BZoom         : TToolbarButton97;
     		Telephone     : THEdit;

        procedure FLISTE_OnDblClick(Sender: TObject);
        procedure BINSERT_OnClick(Sender:TObject);
        procedure AfterShow;
        {$IFDEF QUALITE}
          { Loupe }
          procedure MnLpNonConf_OnClick(Sender: TObject);
          procedure MnNotationFour_OnClick(Sender: TObject);
        {$ENDIF QUALITE}

        procedure PmLoupe_OnPopUp(Sender: TObject);
    	procedure TelephoneChange(Sender: Tobject);
      function GetWhereTypeFourni : string;
    procedure BZoom_Onclick(Sender: Tobject);
   // procedure OnArgument(Arguments: String);

    protected
    	function GetLoupeCtx: string; Override;
    END ;

Procedure RTLanceFiche_Prospect_Mul(Nat,Cod : String ; Range,Lequel,Argument : string) ;

implementation

uses
  wCommuns
  , UtilPGI
{$IFDEF QUALITE}
  ,stdCtrls
  , EntGC, RQTNotations
  , RqNonConf_Tof, rqtNotes_Tof
{$ENDIF}
   ,CbpMCD
   ,CbpEnumerator
  ,TiersUtil
  ,Ent1
  ;

{ EV5 / FQ 15169 (Modification temporaire V8 : reporter dans la vue RFFOURNISSEURS en V9) }
var TabControl: array[1..30] of string;

procedure ChangeControlChamp(Del:boolean);
var i,j,k:integer;
    Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
    FieldList : IEnumerator ;
    Field     : IFieldCOMEx ;

const TabChamp: array [1..30] of string = (
      {1} 'T_APPORTEUR',
      {2} 'T_CODEIMPORT',
      {3} 'T_DATEDERNPIECE',
      {4} 'T_DATEPROCLI',
      {5} 'T_DIVTERRIT',
      {6} 'T_EMAILING',
      {7} 'T_ENSEIGNE',
      {8} 'T_ETATRISQUE',
      {9} 'T_FACTUREHT',
      {10} 'T_FORMEJURIDIQUE',
      {11} 'T_FREQRELEVE',
      {12} 'T_JOURRELEVE',
      {13} 'T_MOISCLOTURE',
      {14} 'T_MOTIFVIREMENT',
      {15} 'T_NATUREECONOMIQUE',
      {16} 'T_NIVEAUIMPORTANCE',
      {17} 'T_NIVEAURISQUE',
      {18} 'T_ORIGINETIERS',
      {19} 'T_PHONETIQUE',
      {20} 'T_PRESCRIPTEUR',
      {21} 'T_PROFIL',
      {22} 'T_PUBLIPOSTAGE',
      {23} 'T_REPRESENTANT',
      {24} 'T_RESIDENTETRANGER',
      {25} 'T_SEXE',
      {26} 'T_SOCIETEGROUPE',
      {27} 'T_TELEPHONE2',
      {28} 'T_TOTALCREDIT',
      {29} 'T_TOTALDEBIT',
      {30} 'T_ZONECOM'
      );
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  
  for k:=1 to 30 do
  begin
		if assigned(mcd.getField(TabChamp[k])) then
    begin
			Field := mcd.getField(TabChamp[k]) as IfieldCOmEx;
      if del then
      begin
        TabControl[k]:=Field.control;
        Field.control:='';
      end
      else Field.control:=TabControl[k];
    end;
  end;
end;
{ EV5 / FQ 15169 (Modification temporaire V8 : reporter dans la vue RFFOURNISSEURS en V9) FIN }


Procedure RTLanceFiche_Prospect_Mul(Nat,Cod : String ; Range,Lequel,Argument : string) ;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_Fournisseur_Mul.OnArgument(Arguments : String );
var F : TForm;
    ChampMul,ValMul,Critere : String;
    x : integer;
begin
	fMulDeTraitement := true;
inherited ;
	fTableName := 'TIERS';
	selection := false;
//
   if pos('SELECTION',Arguments)>0 then BEGIN Action:='SELECTION'; Selection := true; END;
// uniquement en line
// Domaine	:= 'BTP';
// Nomfiche := 'BTFOURNISSEUR_S1';
   Domaine	:= 'GC';
	 Nomfiche := 'GCFOURNISSEUR';
//
   if pos('SELECTION',Arguments)>0 then Action:='SELECTION';
//
  ModuleQualite := pos('SUIVINOTATION',Arguments)>0;
  SetControlVisible('BTLOUPE', ModuleQualite);
  SetControlVisible('BZOOM'  , Not ModuleQualite);

  F := TForm (Ecran);
  if GetParamSocSecur('SO_RTGESTINFOS003',False) = True then
    MulCreerPagesCL(F,'NOMFIC=GCFOURNISSEURS');

  {$ifdef GRC} // mcd 17/08/04 sinon plante si exe GI sans GRC
    xx_where:=RTXXWhereConfident('CONF');
  {$endif}

  if F.Name='RFTIERS_TL' then
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

  if Assigned(GetControl('FLISTE')) then
    ThGrid(GetControl('FLISTE' )).OnDblClick := FLISTE_OnDblClick;
  if Assigned(GetControl('BINSERT')) then
    ThGrid(GetControl('BINSERT' )).OnClick := BInsert_OnClick;
  if Assigned(GetControl('PMLOUPE')) then
    TPopUpMenu(GetControl('PMLOUPE')).OnPopUp := PmLoupe_OnPopUp;

  {$IFDEF QUALITE}
    if (F.Name = 'GCFOURNISSEUR_MUL') then
    begin
      SetControlVisible('SUIVINOTATION', VH_GC.QualiteSeria);
      if VH_GC.QUALITESeria and (GetArgumentBoolean(Arguments, 'SUIVINOTATION')) then
        SetControlChecked('SUIVINOTATION', True);

      if Assigned(GetControl('MNLPNONCONF')) then
        TMenuItem(GetControl('MNLPNONCONF')).OnClick := MnLpNonConf_OnClick;
      if Assigned(GetControl('MNNOTATIONFOUR')) then
        TMenuItem(GetControl('MNNOTATIONFOUR')).OnClick := MnNotationFour_OnClick;
    end;
  {$ENDIF}
  Telephone := THEdit(Ecran.FindComponent('TELEPHONE'));
  Telephone.OnChange := TelephoneChange;

  BZoom     := TToolbarButton97(Ecran.findComponent('BZOOM'));
  Bzoom.OnClick := BZOOM_OnClick;

  GCMAJChampLibre (F, False, 'COMBO', 'YTC_TABLELIBREFOU', 3, '');
  GCMAJChampLibre (F, False, 'EDIT', 'YTC_VALLIBREFOU', 3, '');
  GCMAJChampLibre (F, False, 'EDIT', 'YTC_VALLIBREFOU', 3, '_');
  GCMAJChampLibre (F, False, 'EDIT', 'YTC_DATELIBREFOU', 3, '');
  GCMAJChampLibre (F, False, 'EDIT', 'YTC_DATELIBREFOU', 3, '_');

  // CCMX - DM - 21/09/2005 - DEBUT
//  SetControlVisible('TYTC_TYPEFOURNI' , GetParamSocSecur('SO_GCTRANSPORTEURS', ''));
//  SetControlVisible('YTC_TYPEFOURNI'  , GetParamSocSecur('SO_GCTRANSPORTEURS', ''));
  // CCMX - DM - 21/09/2005 - FIN

  if (Not AutoriseCreationTiers ('FOU')) then
  begin
    SetControlVisible('BINSERT',False) ;
    SetControlVisible('B_DUPLICATION',False) ;
  end;

  { EV5 / FQ 15169 (Modification temporaire V8 : reporter dans la vue RFFOURNISSEURS en V9) }
  ChangeControlChamp(true);
  { EV5 / FQ 15169 (Modification temporaire V8 : reporter dans la vue RFFOURNISSEURS en V9) FIN }
end;

//FV1 : 13/02/2014 - FS#775 - SANITHERM : en saisie pièce achats, si accès au fournisseur, ne pas autoriser la modification.
Procedure TOF_Fournisseur_Mul.BZoom_Onclick(Sender : Tobject);
Var F : TForm;
begin

  F:=TForm(ecran);
  if (TFMul(F).FListe=nil) then exit;

  AGLLanceFiche ('GC', 'GCFOURNISSEUR','',TFMul(F).Fliste.datasource.dataset.FindField('T_AUXILIAIRE').AsString,'ACTION=CONSULTATION;T_NATUREAUXI=FOU');

end;

procedure TOF_Fournisseur_Mul.AfterShow;
begin
  Tfmul(Ecran).ForceSelectFiltre(StFiltre , V_PGI.User,false,true);
end ;

procedure TOF_Fournisseur_Mul.OnLoad;
var
{$IFDEF QUALITE}
  zz_Where: String;
{$ENDIF QUALITE}
  yy_where : string;
begin
  inherited;
  yy_where := '';
  if (GetControltext('TYPEFOURNI')<> '') and (GetControltext('TYPEFOURNI')<> '<<Tous>>') then
  begin
  	yy_where := GetWhereTypeFourni;
  end;
{$IFDEF QUALITE}
  if VH_GC.QualiteSeria and (GetCheckBoxState('SUIVINOTATION') = cbChecked) then
    zz_Where := GetWhereTiersSuivisEnNotation
  else
    zz_Where := '';
  SetControlText('XX_WHERE', xx_where + zz_where+ yy_where);
{$ELSE}
  SetControlText('XX_WHERE', xx_where+yy_where);
{$ENDIF QUALITE}
end;

{$IFDEF QUALITE}
procedure TOF_Fournisseur_Mul.MnLpNonConf_OnClick(Sender: TObject);
begin
  LanceFiche_RtNonConf_Tof('RT', 'RQNONCONF_MUL' 	, 'RQN_TIERS='+GetField('T_TIERS')+';RQN_ETATQNC=0EN', '', 'ACTION='+Action);
end;

procedure TOF_Fournisseur_Mul.MnNotationFour_OnClick(Sender: TObject);
begin
  LanceFiche_QRTNotes(GetField('T_TIERS'), 'ACTION='+Action);
end;
{$ENDIF QUALITE}

procedure TOF_Fournisseur_Mul.FLISTE_OnDblClick(Sender: TObject);
  function GetAction: string;
  begin

{$IFDEF GPAO}
    if not JaiLeDroitGestion('T') then
{$ELSE GPAO}
    if RTDroitModifFou(GetString('T_TIERS'))=False then
{$ENDIF GPAO}
      result:= 'ACTION=CONSULTATION'
    else
      result:= 'ACTION='+action;
  end;
begin
  if Selection then
	begin
  	TFMul(Ecran).Retour := GetString('T_TIERS');
    TFMul(Ecran).Close;
	end
  else
  begin
	  if GetString('T_AUXILIAIRE') <> '' then
      AglLanceFiche(domaine,NomFiche,'', GetString('T_AUXILIAIRE'),GetAction+';MONOFICHE;T_NATUREAUXI='+GetString('T_NATUREAUXI'));
    if GetParamSocSecur('SO_REFRESHMUL', true) then
    	refreshdb;
      //TToolBarButton97(GetCOntrol('Bcherche')).Click;
	end;
end;

function TOF_Fournisseur_Mul.GetLoupeCtx: string;
begin
//  Result := ';QNCNUM='   + GetString('RQE_QNCNUM')
//          + ';TIERS='    + GetString('RQE_TIERS')
end;

procedure TOF_Fournisseur_Mul.PmLoupe_OnPopUp(Sender: TObject);
begin
  if not IsEmpty then
  begin
{$IFDEF QUALITE}
    if Assigned(GetControl('MNLPNONCONF')) then
      TMenuItem(GetControl('MNLPNONCONF')).Visible:= ModuleQualite;
    if Assigned(GetControl('MNNOTATIONFOUR')) then
      TMenuItem(GetControl('MNNOTATIONFOUR')).Visible:= ModuleQualite;
{$ENDIF QUALITE}

    if ModuleQualite then
      inherited PmLoupe_OnPopUp(Sender)
  end
  else
    Abort
end;

procedure TOF_Fournisseur_Mul.BINSERT_OnClick(Sender: TObject);
var NatureAuxi		: String;
begin
//
  NatureAuxi := GetControlText('T_NATUREAUXI');

  if NatureAuxi <> '' then
  Begin
    AGLLanceFiche (Domaine,Nomfiche,'','','MONOFICHE;ACTION=CREATION;T_NATUREAUXI=' + NatureAuxi);
    if GetParamSocSecur('SO_REFRESHMUL', true) then
       TToolBarButton97(GetCOntrol('Bcherche')).Click;
  end;

(*
  BRL 1/09/09 : Mis en commentaire car ne fonctionne pas.
  Récupération des lignes précédentes depuis la V2007

  NatureAuxi := GetControlText('T_NATUREAUXI');

  if ExJaiLeDroitConcept(TConcept(bt511),False) then
  begin
     if NatureAuxi <> '0' then
     begin
      	FParamsLanceFiche := 'T_NATUREAUXI=FOU';
        Inherited;
     end else
		   PGIBox('Impossible de créer un nouveau Fournisseur, (préfixe vide)','');
  end else
  begin
    PGIBox('Vous n''avez pas les droits d''accès en création');
  end;
*)
//
(*
  if RTExisteConfidentF() <> '' then
  begin
    FParamsLanceFiche := 'T_NATUREAUXI=FOU';

    Inherited;
  end
  else
    PGIInfo('Vous n''avez pas les droits d''accÃ¨s en crÃ©ation');
*)
end;

{ EV5 / FQ 15169 (Modification temporaire V8 : reporter dans la vue RFFOURNISSEURS en V9) }
procedure TOF_Fournisseur_Mul.OnClose ;
begin
  Inherited ;
  ChangeControlChamp(false);
end ;
{ EV5 / FQ 15169 (Modification temporaire V8 : reporter dans la vue RFFOURNISSEURS en V9) FIN }

procedure TOF_Fournisseur_Mul.TelephoneChange(Sender: Tobject);
begin
  ThEdit(getCOntrol('T_CLETELEPHONE')).Text := CleTelForFind (ThEdit(getCOntrol('TELEPHONE')).Text);
end;

function TOF_Fournisseur_Mul.GetWhereTypeFourni: string;
var atype,UnType : string;

begin
	result := '';
  atype := GetControlText('TYPEFOURNI');
  repeat
  	UnType := readtokenSt(aType);
    if UnType <> '' then
    begin
    	result := result + ' AND YTC_TYPEFOURNI LIKE "%'+unType+'%"';
    end;
  until untype = '';
end;

Initialization
registerclasses([TOF_Fournisseur_Mul]);
end.
