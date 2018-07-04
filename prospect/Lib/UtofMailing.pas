unit UtofMailing;

{*****************************************************************
Source  utilisé pour les fiches : RTSUSP_MUL_MAILIN
                                  RTPROS_MUL_MAILIN
                                  RTPROS_MAILIN_FIC
                                  RTPRO_MAILIN_CONT
*****************************************************************}

interface

{
	function ConvertDocFile(
		FileSrce, FileDest: String;
		Initialize, Finalize: TNotifyEvent;
		Funct: TOnFuncEvent;
		EGetVar: TGetVarEvent;
		ESetVar: TSetVarEvent;
		GetList: TOnGetListEvent ): Integer;
}

uses  Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HMsgBox,UTOF,HQry,HEnt1,StdCtrls, FileCtrl,menus,HTB97,
      EntGC,Ent1,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
      mul,HDB,Fe_Main,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
{$IFDEF EAGLCLIENT}
{AFAIREEAGL}
{$ELSE}
Doc_Parser,RT_Parser,
{$ENDIF}
{$ifdef AFFAIRE}
      UtofAfTraducChampLibre,
{$ENDIF}

Variants,

      Hstatus,M3FP,UTOB, UtilSelection,Paramsoc,UtilPGI,UtilRT,UtilWord,
      UtilGC,UtilArticle,
//bur 11110 {$ifndef BUREAU}
  UTOFDOCEXTERNE,
//bur 11110 {$endif}

      UtilGA
{$IF Defined(CRM)}
  ,UtofRTSaisieInfosEsker,EskerInterface,Web,AglInit
{$IFEND}
      ;

Type
{$ifdef AFFAIRE}
                //mcd 24/11/2005 pour faire affectation depuis ressource si paramétré
    TOF_Mailing = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
    TOF_Mailing = Class (TOF)
{$endif}
    Private
        TobProsp : TOB ;
        stProduitpgi : string;
        stMaquetteValide : string;
        stNatureDoc : string;
{$IF Defined(CRM)}
        TobInfos : Tob;
        TypeEnvoiEsker : string;
        TypeTraitement : string;
        TobAdr : Tob;
{$IFEND}
        Function MailingValidDoc : Boolean;
        procedure ChangeMaquette;
        procedure ChangeDocument;
        procedure RTPublipostage;
        function MailingValidMail : Boolean;
//bur 11110 {$ifndef BUREAU}
        procedure OpenDoc_OnClick(Sender: TObject);
        procedure MAQUETTE_OnElipsisClick(Sender: TObject);
//bur 11110 {$endif}
        procedure BEMAILING_OnClick(Sender: TObject);
{$IF Defined(CRM)}
        procedure BESKERCOURRIER_OnClick(Sender: TObject);
        procedure BESKEREMAIL_OnClick(Sender: TObject);
        procedure BESKERFAX_OnClick(Sender: TObject);
        procedure BESKERSUIVI_OnClick(Sender: TObject);
        function ChpsObligExiste : Boolean;
        procedure PublipostageCallBack(Index: integer; OutputFileName: hstring; var Cancel: boolean);
        procedure CreatTobInfosMul ;
  private
{$IFEND}
    procedure BRechResponsable(Sender: TObject);
    Public
        procedure OnArgument (Arguments : String ); override ;
        procedure OnLoad ; override ;
        procedure OnClose ; override ;
        procedure EditionMailing(Maquette , DocGenere : string);
        procedure OnGetVar( Sender: TObject; VarName: String; VarIndx: Integer;var Value: variant) ;
{$ifdef gigi}
        procedure bChercheClick(Sender: TObject);
{$endif}
        procedure BVALIDSELClick(Sender: TObject);
        procedure CodeArticleRech(Sendre: TOBject);
        procedure constitueXXWHEREPARC;
        procedure ScreenKeYDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
     END;

function Ctrl_Maquette (stMaquette,TitreEcran,NatureDoc : string;var stMaquetteValide : string) : Boolean;
procedure AGLEditionMailing(parms:array of variant; nb: integer ) ;
Function AGLMailingValidDoc(parms:array of variant; nb: integer ) : variant ;
Function RTLanceFiche_Mailing(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

const
	// libellés des messages
	TexteMessage: array[1..3] of string 	= (
          {1}        'Il manque les champs E-Mail (contact et tiers) dans la liste'
          {2}        ,'Il manque les champs Fax (contact et tiers) dans la liste'
          {3}        ,'Envoi n° %s'
          );

implementation

uses ed_tools,EntRt,AglInitGC,uTILRESSOURCE;

procedure AGLEditionMailing(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_Mailing) then TOF_Mailing(TOTOF).EditionMailing(string(Parms[1]),string(Parms[2])) else exit;

end;

Function AGLMailingValidDoc( parms: array of variant; nb: integer ) : variant;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
result:=TOF_Mailing(TOTOF).MailingValidDoc;
end;

{$ifdef gigi}
//mcd 29/11/2005 pour passer outre ce qui est fait dans latof,
// car pas même vue en GI
procedure TOF_Mailing.BChercheClick(Sender: TObject);
var stWhere : string;
begin
  if not Assigned(GetControl('EXISTE')) then Exit; //mcd 11/05/2006 pour les fiche qui ne l'ont pas (12939)
  if TCheckBox(GetControl('EXISTE')).state=cbgrayed then
  begin
  	stwhere:=''
  end else
  begin
    if TCheckBox(GetControl('EXISTE')).checked=False then stwhere:='not '
                                                     else  stwhere:='';

    stwhere:=stwhere+'exists (select RAC_NUMACTGEN from ACTIONS  where ';
    // PL le 29/08/07 : la vue n'est pas la même sur les écrans GI, on doit donc ajuster la clause where en fonction de l'écran
    if (TFMul(Ecran).name = 'RTPROS_MUL_MAILIN') then
    stwhere:=stwhere+'((RAC_AUXILIAIRE=AFDPTIERS.T_AUXILIAIRE)'
    else
    if (TFMul(Ecran).name = 'RTPRO_MAILIN_CONT') then
    stwhere:=stwhere+'((RAC_AUXILIAIRE=AFDPCONTTIERS.T_AUXILIAIRE)';

    stwhere:=stwhere+ ' AND RAC_OPERATION="'
            +GetControlText ('OPERATION')+'" AND (RAC_NUMACTGEN='
            +GetControlText('NUMACTGEN')+')))';
  end;
  SetControlText ('XX_WHERE',stwhere);
  TFMul(Ecran).BCherCheClick(self) ;
end;
{$endif}

procedure TOF_Mailing.OnArgument (Arguments : String );
var    F : TForm;
       x,LaVersionWord : integer;
       CC : THValComboBox;
       CCe : Thedit;
begin
	fMulDeTraitement := true;
inherited ;

  if TToolbarButton97 (GetControl('BVALIDSEL')) <> nil then
  begin
    TToolbarButton97 (GetControl('BVALIDSEL')).OnClick := BVALIDSELClick;
    TToolbarButton97 (GetControl('BVALIDSEL')).Left := THValComboBox (GetControl('BCherche')).Left;
  end;
  if not JaiLeDroitTag(323210) then
  begin
    TTabSheet (GetControl('TBPARCSAV')).TabVisible := false;
  end;
  if ThEdit(GetControl('CODEARTICLE')) <> nil then
  begin
    ThEdit(GetControl('CODEARTICLE')).OnElipsisClick := CodeArticleRech;
  end;
  Ecran.OnKeyDown := ScreenKeYDown;

  //fTableName := 'AFFAIRE';
  if (pos('GRF', Arguments) <> 0) then
    stProduitpgi := 'GRF'
  else
    stProduitpgi := 'GRC';
  //
  F := TForm (Ecran);
  //
  if stProduitpgi = 'GRF' then
  begin
    if GetParamSocSecur('SO_RTGESTINFOS003',False) = True then MulCreerPagesCL(F,'NOMFIC=GCFOURNISSEURS');
  end
  else if (F.name <> 'RTSUSP_MUL_MAILIN') and (F.name <> 'BTMULCONTREN') then MulCreerPagesCL(F,'NOMFIC=GCTIERS');
  //
  if assigned (ThEdit(GetControl('RAC_INTERVENANT'))) then
  begin
    SetControlText ('RAC_INTERVENANT',VH_RT.RTResponsable);
  end;
  //
  if (F.name = 'RTPRO_MAILIN_CONT') and (GetParamSocSecur('SO_RTGESTINFOS006',False) = True) then MulCreerPagesCL(F,'NOMFIC=YYCONTACT');

  if (pos('FIC', Arguments) <> 0) then
   begin
   if (F.name = 'RTPRO_MAILIN_CONT') then
      begin
      Ecran.Caption:= TraduireMemoire('Export pour Mailing par contact');
      updatecaption(Ecran);
      //      TFMul(F).Q.Liste:='RTCONTMAILINGFIC';
      TFMul(F).SetDBListe('RTCONTMAILINGFIC');
      SetControlVisible('PDOCUMENT',FALSE) ;
      SetControlVisible('BCREATDOC',FALSE) ;
      SetControlVisible('BMAILINGDOC',FALSE) ;
      SetControlVisible('BEMAILING',FALSE) ;
      Ecran.HelpContext:=111000333 ;
      end;
   end
  else
    begin
    if stProduitpgi = 'GRF' then
      begin
      //SetControlText('MAQUETTE',GetParamsocSecur('SO_RFDIRMAQUETTE','C:\PGI00\STD\MAQUETTE')+'\*.doc');
        SetControlText('DOCGENERE',GetParamsocSecur('SO_RFDIRSORTIE','')+'\DOCUMENTFRS.doc');
      end
    else
      begin
      //SetControlText('MAQUETTE',GetParamsocSecur('SO_RTDIRMAQUETTE','C:\PGI00\STD\MAQUETTE')+'\*.doc');
        SetControlText('DOCGENERE',GetParamsocSecur('SO_RTDIRSORTIE','')+'\MAILING.doc');
      end;
    end;

  if (F.name = 'RTPRO_MAILIN_CONT') then
   begin
   // LaVersionWord := GetWordVersion;
   LaVersionWord := 9;
   if LaVersionWord<10 then
      begin
      SetControlVisible('ENVOIEMAIL',FALSE) ;
      SetControlVisible('TOBJETMAIL',FALSE) ;
      SetControlVisible('OBJETMAIL',FALSE) ;
      end
   else if GetControlText('ENVOIEMAIL')<>'X' then SetControlEnabled('OBJETMAIL',FALSE) ;
   end;

  SetControlEnabled('BMAILINGDOC',FALSE) ;
  SetControlEnabled('BEMAILING',FALSE) ;
  
  if stProduitPgi = 'GRF' then stNatureDoc := 'MLF' else stNatureDoc := 'MLC';

  // Paramétrage des libellés des familles, stat. article et dimensions
  if  (TFMul(Ecran).name = 'RTLIGN_MUL_MAILIN') or
      (TFMul(Ecran).name = 'RTLIGN_MAILIN_FIC') or
      (TFMul(Ecran).name = 'RFLIGN_MUL_MAILIN') or
      (TFMul(Ecran).name = 'RFLIGN_MAILIN_FIC') then
  begin
  for x:=1 to 3 do ChangeLibre2('TGA_FAMILLENIV'+InttoStr(x),Ecran);
  ChangeLibre2('TGA_COLLECTION',Ecran);
  MajChampsLibresArticle(TForm(Ecran));
  FiltreComboTypeArticle(THMultiValComboBox(GetControl('TYPEARTICLE')));
  end;

  if (GetControl('YTC_RESSOURCE1') <> nil)  then
  begin
  if not (ctxaffaire in V_PGI.PGICONTEXTE) then SetControlVisible ('PRESSOURCE',false)
    else
    begin
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

{$Ifdef GIGI}
 if (GetControl('T_REPRESENTANT') <> nil) then  SetControlVisible('T_REPRESENTANT',false);
 if (GetControl('TT_REPRESENTANT') <> nil) then  SetControlVisible('TT_REPRESENTANT',false);
 if (GetControl('T_ZONECOM') <> nil) then  SetControlVisible('T_ZONECOM',false);
 if (GetControl('TT_ZONECOM') <> nil) then  SetControlVisible('TT_ZONECOM',false);
 if (f.name <> 'BTMULCONTREN') then
 begin
   if (F.name <> 'RTSUSP_MUL_MAILIN') then
     begin
     SetControlText('T_NatureAuxi','');    //on efface les valeurs CLI et PO, car NCP en plus
     SetControlProperty ('T_NATUREAUXI', 'Complete', true);
     SetControlProperty ('T_NATUREAUXI_', 'Complete', true);
     SetControlProperty ('T_NATUREAUXI', 'Datatype', 'TTNATTIERS');
     SetControlProperty ('T_NATUREAUXI_', 'Datatype', 'TTNATTIERS');
     SetControlProperty ('T_NATUREAUXI_', 'Plus', VH_GC.AfNatTiersGRCGI);
     SetControlProperty ('T_NATUREAUXI', 'Plus', VH_GC.AfNatTiersGRCGI);
     end
   else begin
    if (GetControl('RSU_REPRESENTANT') <> nil) then  SetControlVisible('RSU_REPRESENTANT',false);
    if (GetControl('TRSU_REPRESENTANT') <> nil) then  SetControlVisible('TRSU_REPRESENTANT',false);
    if (GetControl('RSU_ZONECOM') <> nil) then  SetControlVisible('RSU_ZONECOM',false);
    if (GetControl('TRSU_ZONECOM') <> nil) then  SetControlVisible('TRSU_ZONECOM',false);
    end;
 end;
  //MCD 04/07/2005
 if (F.name = 'RTPROS_MUL_MAILIN') then
  begin //on peut avoir des info du DP
  //mcd 20/09/2005 agl 148 TfMul(Ecran).dbliste := 'AFDPTIERSMAILING';
  //mcd 20/09/2005 agl 148 if TfMul(Ecran).Q <> NIL then TfMul(Ecran).Q.Liste  := 'AFDPTIERSMAILING';
  TfMul(Ecran).Setdbliste ('AFDPTIERSMAILING');
  MulCreerPagesDP(F);
  end;
 if (F.name = 'RTPRO_MAILIN_CONT') then
  begin //on peut avoir des info du DP
  //mcd 20/09/2005 agl 148 TfMul(Ecran).dbliste := 'AFMULCONTMAILING';
  //mcd 20/09/2005 agl 148 if TfMul(Ecran).Q <> NIL then TfMul(Ecran).Q.Liste  := 'AFMULCONTMAILING';
  TfMul(Ecran).Setdbliste ('AFMULCONTMAILING');
  MulCreerPagesDP(F);
  end;
  If TfMul(Ecran).name <>'RTSUSP_MUL_MAILIN' then     //mcd 20/06/06 12939  pour suspect, pas de changement de liste
      TToolBarButton97(GetControl('BCHERCHE')).OnClick := bChercheClick; //mcd 29/11/2005

{$endif}

//bur 11110{$ifndef BUREAU}
  if Assigned(GetControl('OpenDoc')) then
    TMenuItem(GetControl('OpenDoc')).OnClick := OpenDoc_OnClick;
  if Assigned(GetControl('MAQUETTE')) then
    THEDIT(GetControl('MAQUETTE')).OnElipsisClick := MAQUETTE_OnElipsisClick;
(*//bur 11110 {$else}    //dans le bureau, on envuet pas les maquetts. sinon pb avec affaireOle
  SetControlVisible ('MAQUETTE',false);
  TMenuItem(GetControl('OpenDoc')).visible := false;
{$endif} *)
  if Assigned(GetControl('BEMAILING')) then
    TToolBarButton97(GetControl('BEMAILING')).OnClick := BEMAILING_OnClick;
{$IF Defined(CRM)}
  if (F.name = 'RTPROS_MUL_MAILIN') or (F.name = 'RTPRO_MAILIN_CONT') or
     (F.name = 'RTLIGN_MUL_MAILIN') then
    begin
    if GetParamSocSecur ('SO_RTGESTIONFLYDOC',False) = True then
      begin
      SetControlVisible ('BESKERSUIVI',True);
      SetControlVisible ('BESKERCOURRIER',True);
      SetControlVisible ('BESKEREMAIL',True);
      SetControlVisible ('BESKERFAX',True);
      SetControlEnabled('BESKERCOURRIER',False) ;
      SetControlEnabled('BESKEREMAIL',False) ;
      SetControlEnabled('BESKERFAX',False) ;
      end;
    if Assigned(GetControl('BESKERCOURRIER')) then
       TToolbarButton97(GetControl('BESKERCOURRIER')).OnClick := BESKERCOURRIER_OnClick;
    if Assigned(GetControl('BESKEREMAIL')) then
       TToolbarButton97(GetControl('BESKEREMAIL')).OnClick := BESKEREMAIL_OnClick;
    if Assigned(GetControl('BESKERFAX')) then
       TToolbarButton97(GetControl('BESKERFAX')).OnClick := BESKERFAX_OnClick;
    if Assigned(GetControl('BESKERSUIVI')) then
       TToolbarButton97(GetControl('BESKERSUIVI')).OnClick := BESKERSUIVI_OnClick;
    end;
{$IFEND}

//

  // gestion Etablissement (BTP)
	CC:=THValComboBox(GetControl('BTBETABLISSEMENT')) ;
	if CC<>Nil then
  begin
  	PositionneEtabUser(CC) ;
    if not VH^.EtablisCpta then
    begin
    	if THLabel(GetControl('TBTB_ETABLISSEMENT')) <> nil then THLabel(GetControl('TBTB_ETABLISSEMENT')).Visible := false;
			CC.visible := false;
    end;
	end;

  // gestion Domaine (BTP)
	CC:=THValComboBox(GetControl('AFF_DOMAINE')) ;
	if CC<>Nil then PositionneDomaineUser(CC) ;

	CCE:=ThEdit(GetControl('RAC_INTERVENANT')) ;
	if CCE<>Nil then
  begin
  	PositionneResponsableUser(CCE) ;
    CCE.OnElipsisClick := BRechResponsable;
	end;

end;

procedure TOF_Mailing.BRechResponsable(Sender: TObject);
Var QQ  : TQuery;
    SS  : THCritMaskEdit;
begin

  if GetParamSocSecur('SO_AFRECHRESAV', True) then
  begin
    SS := THCritMaskEdit.Create(nil);
    GetRessourceRecherche(SS,'ARS_RESSOURCE=' + ThEdit(Sender).text + ';TYPERESSOURCE=SAL', '', '');
    if (SS.Text <> THEdit(Sender).text) then
    begin
      if SS.text = '' then ss.text := ThEdit(Sender).text;
    end;
    ThEdit(Sender).text  := SS.Text;
    SS.Free;
  end
  else
    GetRessourceRecherche(ThEdit(Sender),'ARS_TYPERESSOURCE="SAL"', '', '');

end;

procedure TOF_Mailing.OnLoad;
var xx_where,StWhere,StListeActions,StJoint,ListeCombos,Confid : string;
    DateDeb,DateFin : TDateTime;
    i: integer;
begin
inherited ;
if (TFMul(Ecran).name = 'RTSUSP_MUL_MAILIN') then exit;
if (TFMul(Ecran).name = 'BTMULCONTREN') then exit;
StWhere := '';
if (GetControlText('SANSSELECT') <> 'X') then
   begin
   DateDeb := StrToDate(GetControlText('DATEACTION'));
   DateFin := StrToDate(GetControlText('DATEACTION_'));
   if GetControlText('SANS') = 'X' then StWhere := 'NOT ';
   if (TFMul(Ecran).name = 'RTPRO_MAILIN_CONT') then
      begin
    {$ifdef GIGI} //MCD 04/07/2005
      StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = AFDPCONTTIERS.T_AUXILIAIRE';
      StWhere:=StWhere + ' AND RAC_NUMEROCONTACT = AFDPCONTTIERS.C_NUMEROCONTACT';
    {$else}
      StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = RTCONTACTSTIERS.T_AUXILIAIRE';
      StWhere:=StWhere + ' AND RAC_NUMEROCONTACT = RTCONTACTSTIERS.C_NUMEROCONTACT';
    {$ENDIF GIGI}
      end
   else
       begin
       if stProduitpgi = 'GRF' then StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = RFFOURNISSEURS.T_AUXILIAIRE'
 {$ifdef GIGI} //MCD 04/07/2005
      // PL le 05/09/07 : en contexte GI la vue derrière la liste est AFDPTIERS
       //else StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = AFDPCONTTIERS.T_AUXILIAIRE';
       else StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = AFDPTIERS.T_AUXILIAIRE';
  {$else}
       else StWhere:=StWhere + 'EXISTS (SELECT RAC_NUMACTION FROM ACTIONS WHERE (RAC_AUXILIAIRE = RTTIERS.T_AUXILIAIRE';
 {$ENDIF GIGI}
       end;
   if (GetControlText('TYPEACTION') <> '') and (GetControlText('TYPEACTION') <> TraduireMemoire('<<Tous>>')) then
      begin
      StListeActions:=FindEtReplace(GetControlText('TYPEACTION'),';','","',True);
      StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
      StWhere:=StWhere + ' AND RAC_TYPEACTION in '+StListeActions;
      end;
   if GetControlText('RESPONSABLE') <> '' then
      begin
      ListeCombos := GetControlText('RESPONSABLE');
      if copy(ListeCombos,length(ListeCombos),1) <> ';' then ListeCombos := ListeCombos + ';';
      ListeCombos:=FindEtReplace(ListeCombos,';','","',True);
      ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
      StWhere:=StWhere + ' AND RAC_INTERVENANT in '+ListeCombos;
      end;
    for i:=1 to 3 do
        begin
        if (GetControlText('TABLELIBRE'+intToStr(i)) <> '') and (GetControlText('TABLELIBRE'+intToStr(i)) <> TraduireMemoire('<<Tous>>')) then
            begin
            ListeCombos:=FindEtReplace(GetControlText('TABLELIBRE'+intToStr(i)),';','","',True);
            ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
            if stProduitpgi = 'GRF' then StWhere := StWhere + ' AND RAC_TABLELIBREF' +intToStr(i) +' in ' + ListeCombos
            else StWhere := StWhere + ' AND RAC_TABLELIBRE' +intToStr(i) +' in ' + ListeCombos;
            end;
        end;
   if GetControlText('OPERATIONACT') <> '' then
      begin
      StWhere:=StWhere + ' AND RAC_OPERATION = "'+GetControlText('OPERATIONACT')+'"';
      end;
   if (GetControlText('ETATACTION') <> '') and (GetControlText('ETATACTION') <> TraduireMemoire('<<Tous>>')) then
      begin
      ListeCombos:=FindEtReplace(GetControlText('ETATACTION'),';','","',True);
      ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
      StWhere:=StWhere + ' AND RAC_ETATACTION in '+ListeCombos;
      end;
   if (GetControlText('NIVIMP') <> '') and (GetControlText('NIVIMP') <> TraduireMemoire('<<Tous>>')) then
      begin
      ListeCombos:=FindEtReplace(GetControlText('NIVIMP'),';','","',True);
      ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
      StWhere:=StWhere + ' AND RAC_NIVIMP in '+ListeCombos;
      end;
   StWhere := StWhere + ' AND RAC_DATEACTION >= "'+UsDateTime(DateDeb) +'" AND RAC_DATEACTION <= "'+UsDateTime(DateFin)+'"))';
   if (GetControlText('XX_WHERE') = '') then
      SetControlText('XX_WHERE',StWhere)
   else
       begin
       xx_where := GetControlText('XX_WHERE');
       xx_where := xx_where + ' and (' + StWhere + ')';
       SetControlText('XX_WHERE',xx_where) ;
       end;
   end;
{ sélection sur critères lignes/articles }
if (TFMul(Ecran).name = 'RTLIGN_MUL_MAILIN') or (TFMul(Ecran).name = 'RTLIGN_MAILIN_FIC') or
   (TFMul(Ecran).name = 'RFLIGN_MUL_MAILIN') or (TFMul(Ecran).name = 'RFLIGN_MAILIN_FIC') then
    begin
       StWhere := '';
       DateDeb := StrToDate(GetControlText('DATEPIECE'));
       DateFin := StrToDate(GetControlText('DATEPIECE_'));
       if GetControlText('PASDEPIECE') = 'X' then StWhere := 'NOT ';
       if stProduitpgi = 'GRF' then StWhere:=StWhere + 'EXISTS (SELECT GL_NUMERO FROM LIGNE WHERE (GL_TIERS = RFFOURNISSEURS.T_TIERS'
       else StWhere:=StWhere + 'EXISTS (SELECT GL_NUMERO FROM LIGNE WHERE (GL_TIERS = RTTIERS.T_TIERS';
       // NATUREPIECEG,ETABLISSEMENT,DEPOT,REPRESENTANT,ARTICLE,TYPEARTICLE,DOMAINE,
       if (GetControlText('ETABLISSEMENT') <> '') and (GetControlText('ETABLISSEMENT') <> TraduireMemoire('<<Tous>>')) then
          begin
          StListeActions:=FindEtReplace(GetControlText('ETABLISSEMENT'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_ETABLISSEMENT in '+StListeActions;
          end;
       if (GetControlText('NATUREPIECEG') <> '') and (GetControlText('NATUREPIECEG') <> TraduireMemoire('<<Tous>>')) then
          begin
          StListeActions:=FindEtReplace(GetControlText('NATUREPIECEG'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_NATUREPIECEG in '+StListeActions;
          end;
       if (GetControlText('DEPOT') <> '') and (GetControlText('DEPOT') <> TraduireMemoire('<<Tous>>')) then
          begin
          StListeActions:=FindEtReplace(GetControlText('DEPOT'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_DEPOT in '+StListeActions;
          end;
       if (GetControlText('TYPEARTICLE') <> '') and (GetControlText('TYPEARTICLE') <> TraduireMemoire('<<Tous>>')) then
          begin
          StListeActions:=FindEtReplace(GetControlText('TYPEARTICLE'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_TYPEARTICLE in '+StListeActions;
          end;
       if (GetControlText('DOMAINE') <> '') and (GetControlText('DOMAINE') <> TraduireMemoire('<<Tous>>')) then
          begin
          StListeActions:=FindEtReplace(GetControlText('DOMAINE'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_DOMAINE in '+StListeActions;
          end;
       // COLLECTION, FAMILLENIV1,FAMILLENIV2,FAMILLENIV3
       if (GetControlText('COLLECTION') <> '') and (GetControlText('COLLECTION') <> TraduireMemoire('<<Tous>>')) then
          begin
          StListeActions:=FindEtReplace(GetControlText('COLLECTION'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_COLLECTION in '+StListeActions;
          end;
       if (GetControlText('FAMILLENIV1') <> '') and (GetControlText('FAMILLENIV1') <> TraduireMemoire('<<Tous>>')) then
          begin
          StListeActions:=FindEtReplace(GetControlText('FAMILLENIV1'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_FAMILLENIV1 in '+StListeActions;
          end;
       if (GetControlText('FAMILLENIV2') <> '') and (GetControlText('FAMILLENIV2') <> TraduireMemoire('<<Tous>>')) then
          begin
          StListeActions:=FindEtReplace(GetControlText('FAMILLENIV2'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_FAMILLENIV2 in '+StListeActions;
          end;
       if (GetControlText('FAMILLENIV3') <> '') and (GetControlText('FAMILLENIV3') <> TraduireMemoire('<<Tous>>')) then
          begin
          StListeActions:=FindEtReplace(GetControlText('FAMILLENIV3'),';','","',True);
          StListeActions:='("'+copy(StListeActions,1,Length(StListeActions)-2)+')';
          StWhere:=StWhere + ' AND GL_FAMILLENIV3 in '+StListeActions;
          end;
       if GetControlText('ARTICLE') <> '' then
          begin
          StWhere:=StWhere + ' AND GL_CODEARTICLE = "'+GetControlText('ARTICLE')+'"';
          end;
       if (stProduitpgi <> 'GRF') and (GetControlText('REPRESENTANT') <> '') then
          begin
          StWhere:=StWhere + ' AND GL_REPRESENTANT = "'+GetControlText('REPRESENTANT')+'"';
          end;
       for i:=1 to 9 do
          begin
          if (GetControlText('LIBREART'+intToStr(i)) <> '') and (GetControlText('LIBREART'+intToStr(i)) <> TraduireMemoire('<<Tous>>')) then
              begin
              ListeCombos := GetControlText('LIBREART'+intToStr(i));
              if copy(ListeCombos,length(ListeCombos),1) <> ';' then ListeCombos := ListeCombos + ';';
              ListeCombos:=FindEtReplace(ListeCombos,';','","',True);
              ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
              StWhere:=StWhere + ' AND GL_LIBREART'+intToStr(i)+' in '+ListeCombos;
              end;
          end;
       if (GetControlText('LIBREARTA') <> '') and (GetControlText('LIBREARTA') <> TraduireMemoire('<<Tous>>')) then
          begin
          ListeCombos := GetControlText('LIBREARTA');
          if copy(ListeCombos,length(ListeCombos),1) <> ';' then ListeCombos := ListeCombos + ';';
          ListeCombos:=FindEtReplace(ListeCombos,';','","',True);
          ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
          StWhere:=StWhere + ' AND GL_LIBREARTA in '+ListeCombos;
          end;

       StWhere := StWhere + ' AND GL_DATEPIECE >= "'+UsDateTime(DateDeb) +'" AND GL_DATEPIECE <= "'+UsDateTime(DateFin)+'"))';
       if (GetControlText('XX_WHERE') = '') then
          SetControlText('XX_WHERE',StWhere)
       else
           begin
           xx_where := GetControlText('XX_WHERE');
           xx_where := xx_where + ' and (' + StWhere + ')';
           SetControlText('XX_WHERE',xx_where) ;
           end;
    end;

StWhere := '';
if (GetControlText('COURRIER') <> 'X') then
   begin
   if stProduitpgi = 'GRF' then
     begin
     if GetControlText('FAX') = 'X' then StWhere := '((T_FAX <> "") or (C_FAX <> ""))';
     if GetControlText('NOTFAX') = 'X' then StWhere := '(((T_FAX = "") or (T_FAX is null)) and ((C_FAX = "") or (C_FAX is null))) ';
     end
   else
     begin
     if (TFMul(Ecran).name <> 'RTLIGN_MAILIN_FIC') and (TFMul(Ecran).name <> 'RTLIGN_MUL_MAILIN') then
       begin
       if GetControlText('FAX') = 'X' then StWhere := '((T_FAX <> "") or (C_FAX <> "")) and T_PARTICULIER <> "X"';
       if GetControlText('NOTFAX') = 'X' then StWhere := '(((T_FAX = "") or (T_FAX is null)) and ((C_FAX = "") or (C_FAX is null))) or T_PARTICULIER = "X"';
       end
     else
       begin
       if GetControlText('FAX') = 'X' then StWhere := '(T_FAX <> "") and T_PARTICULIER <> "X"';
       if GetControlText('NOTFAX') = 'X' then StWhere := '((T_FAX = "") or (T_FAX is null)) or T_PARTICULIER = "X"';
       end;
     end;

   if (TFMul(Ecran).name = 'RTPRO_MAILIN_CONT') then
      begin
      if GetControlText('MAIL') = 'X' then StWhere := '(C_RVA <> "")';
      if GetControlText('NOTMAIL') = 'X' then StWhere := '(C_RVA = "")'
      end;
   if (GetControlText('XX_WHERE') = '') then
      SetControlText('XX_WHERE',StWhere)
   else
       begin
       xx_where := GetControlText('XX_WHERE');
       xx_where := xx_where + ' and (' + StWhere + ')';
       SetControlText('XX_WHERE',xx_where) ;
       end;
   end;

  if stProduitpgi <> 'GRF' then Confid:='CON' else Confid:='CONF';
  if (GetControlText('XX_WHERE') = '') then
      SetControlText('XX_WHERE',RTXXWhereConfident(Confid))
  else
      begin
      xx_where := GetControlText('XX_WHERE');
      xx_where := xx_where + RTXXWhereConfident(Confid);
      SetControlText('XX_WHERE',xx_where) ;
      end;

if (TFMul(Ecran).name = 'RTPROS_MUL_MAILIN') or (TFMul(Ecran).name = 'RTPROS_MAILIN_FIC') or
   (TFMul(Ecran).name = 'RTLIGN_MUL_MAILIN') or (TFMul(Ecran).name = 'RTLIGN_MAILIN_FIC') or
   (TFMul(Ecran).name = 'RFPROS_MUL_MAILIN') or (TFMul(Ecran).name = 'RFPROS_MAILIN_FIC') or
   (TFMul(Ecran).name = 'RFLIGN_MUL_MAILIN') or (TFMul(Ecran).name = 'RFLIGN_MAILIN_FIC') then
    begin
    StJoint := '';
    StJoint := StJoint + 'C_FERME <> "X" and ';
    if (GetCheckBoxState ('PRINCIPAL') <> cbGrayed) then
       StJoint := StJoint + 'C_PRINCIPAL = "' + GetControlText ('PRINCIPAL') + '" and ';
    if (GetCheckBoxState ('PUBLIPOSTAGE') <> cbGrayed) then
       StJoint := StJoint + 'C_PUBLIPOSTAGE = "' + GetControlText ('PUBLIPOSTAGE') + '" and ';
    if (GetControlText('FONCTIONCODEE') <> '') and (GetControlText('FONCTIONCODEE') <> TraduireMemoire('<<Tous>>')) then
       begin
//       ListeCombos:=FindEtReplace(GetControlText('FONCTIONCODEE'),';','","',True);
//       ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
//       StJoint := StJoint + 'C_FONCTIONCODEE in ' + ListeCombos + ' and ';
         StJoint := StJoint + 'C_FONCTIONCODEE = "' + GetControlText ('FONCTIONCODEE') + '" and ';
       end;
    if (GetControlText('SERVICECODE') <> '') and (GetControlText('SERVICECODE') <> TraduireMemoire('<<Tous>>')) then
       begin
         StJoint := StJoint + 'C_SERVICECODE = "' + GetControlText ('SERVICECODE') + '" and ';
       end;
    For i:=1 to 10 do
        begin
        if i < 10 then
           begin
           if (GetControlText('LIBRECONTACT'+intToStr(i)) <> '') and (GetControlText('LIBRECONTACT'+intToStr(i)) <> TraduireMemoire('<<Tous>>')) then
              begin
    //            ListeCombos:=FindEtReplace(GetControlText('LIBRECONTACT'+intToStr(i)),';','","',True);
    //            ListeCombos:='("'+copy(ListeCombos,1,Length(ListeCombos)-2)+')';
    //            StJoint := StJoint + 'C_LIBRECONTACT' +intToStr(i) +' in ' + ListeCombos + ' and ';
              StJoint := StJoint + 'C_LIBRECONTACT' +intToStr(i) +' = "' + GetControlText('LIBRECONTACT'+intToStr(i)) + '" and ';
              end;
           end else
           begin
           if (GetControlText('LIBRECONTACTA') <> '') and (GetControlText('LIBRECONTACTA') <> TraduireMemoire('<<Tous>>')) then
              begin
              StJoint := StJoint + 'C_LIBRECONTACTA' +' = "' + GetControlText('LIBRECONTACTA') + '" and ';
              end;
           end;
        end;
    SetControlText ('XX_JOIN',StJoint);
    end;
end;

procedure TOF_Mailing.EditionMailing(Maquette , DocGenere : string);
var  F : TFMul;
{$IFDEF EAGLCLIENT}
       L : THGrid;
{$ELSE}
       Ext : String;
       L : THDBGrid;
{$ENDIF}
//     Q : THQuery;
     i : integer;
     codeprospect,stClauseWhere,StSelect,StSelect2 : string;
     QQ : TQuery ;
     Pages:TPageControl;
begin
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
   begin
   PGIInfo('Aucun élément sélectionné','');
   exit;
   end;
if PGIAsk('Confirmez-vous le traitement ?','')<>mrYes then exit ;

L:= F.FListe;
//Q:= F.Q;
Pages:=F.Pages;

StSelect := 'Select T_TIERS,T_AUXILIAIRE,T_NATUREAUXI,T_LIBELLE,T_ADRESSE1,T_ADRESSE2,T_ADRESSE3';
StSelect := StSelect + ',T_VILLE,T_CODEPOSTAL,T_TELEPHONE,T_SIRET,T_FAX,T_JURIDIQUE,T_PAYS';
StSelect := StSelect + ',T_TELEX,T_ZONECOM,T_REPRESENTANT,T_APE,T_PRENOM,T_SOCIETEGROUPE,T_PRESCRIPTEUR';
StSelect := StSelect + ',T_PARTICULIER,C_NOM,C_PRENOM,C_CIVILITE,C_FONCTION,C_SERVICE';
StSelect := StSelect + ',RPR_RPRLIBBOOL0,RPR_RPRLIBBOOL1,RPR_RPRLIBBOOL2,RPR_RPRLIBBOOL3';
StSelect := StSelect + ',RPR_RPRLIBBOOL4,RPR_RPRLIBBOOL5,RPR_RPRLIBBOOL6,RPR_RPRLIBBOOL7';
StSelect := StSelect + ',RPR_RPRLIBBOOL8,RPR_RPRLIBBOOL9';
StSelect := StSelect + ',RPR_RPRLIBTABLE0,RPR_RPRLIBTABLE1,RPR_RPRLIBTABLE2,RPR_RPRLIBTABLE3,RPR_RPRLIBTABLE4';
StSelect := StSelect + ',RPR_RPRLIBTABLE5,RPR_RPRLIBTABLE6,RPR_RPRLIBTABLE7,RPR_RPRLIBTABLE8,RPR_RPRLIBTABLE9';
StSelect := StSelect + ',RPR_RPRLIBTABLE10,RPR_RPRLIBTABLE11,RPR_RPRLIBTABLE12,RPR_RPRLIBTABLE13,RPR_RPRLIBTABLE14';
StSelect := StSelect + ',RPR_RPRLIBTABLE15,RPR_RPRLIBTABLE16,RPR_RPRLIBTABLE17,RPR_RPRLIBTABLE18,RPR_RPRLIBTABLE19';
StSelect := StSelect + ',RPR_RPRLIBTABLE20,RPR_RPRLIBTABLE21,RPR_RPRLIBTABLE22,RPR_RPRLIBTABLE23,RPR_RPRLIBTABLE24';
StSelect := StSelect + ',RPR_RPRLIBTABLE25';
StSelect := StSelect + ',RPR_RPRLIBVAL0,RPR_RPRLIBVAL1,RPR_RPRLIBVAL2,RPR_RPRLIBVAL3,RPR_RPRLIBVAL4';
StSelect := StSelect + ',RPR_RPRLIBVAL5,RPR_RPRLIBVAL6,RPR_RPRLIBVAL7,RPR_RPRLIBVAL8,RPR_RPRLIBVAL9';
StSelect := StSelect + ',RPR_RPRLIBDATE0,RPR_RPRLIBDATE1,RPR_RPRLIBDATE2,RPR_RPRLIBDATE3,RPR_RPRLIBDATE4';
StSelect := StSelect + ',RPR_RPRLIBDATE5,RPR_RPRLIBDATE6,RPR_RPRLIBDATE7,RPR_RPRLIBDATE8,RPR_RPRLIBDATE9';
StSelect := StSelect + ',RPR_RPRLIBTEXTE0,RPR_RPRLIBTEXTE1,RPR_RPRLIBTEXTE2,RPR_RPRLIBTEXTE3,RPR_RPRLIBTEXTE4';
StSelect := StSelect + ',RPR_RPRLIBTEXTE5,RPR_RPRLIBTEXTE6,RPR_RPRLIBTEXTE7,RPR_RPRLIBTEXTE8,RPR_RPRLIBTEXTE9';
StSelect := StSelect + ',YTC_TABLELIBRETIERS1,YTC_TABLELIBRETIERS2,YTC_TABLELIBRETIERS3';
StSelect := StSelect + ',YTC_TABLELIBRETIERS4,YTC_TABLELIBRETIERS5,YTC_TABLELIBRETIERS6';
StSelect := StSelect + ',YTC_TABLELIBRETIERS7,YTC_TABLELIBRETIERS8,YTC_TABLELIBRETIERS9';
StSelect := StSelect + ',YTC_TABLELIBRETIERSA,YTC_BOOLLIBRE1,YTC_BOOLLIBRE2,YTC_BOOLLIBRE3';
StSelect := StSelect + ',YTC_VALLIBRE1,YTC_VALLIBRE2,YTC_VALLIBRE3,YTC_DATELIBRE1,YTC_DATELIBRE2,YTC_DATELIBRE3';
StSelect := StSelect + ',YTC_TEXTELIBRE1,YTC_TEXTELIBRE2,YTC_TEXTELIBRE3';
StSelect := StSelect + ' FROM RTTIERS';
StSelect := StSelect + ' left join CONTACT on C_AUXILIAIRE=T_AUXILIAIRE and C_TYPECONTACT = "T" and C_PRINCIPAL = "X" ';


TobProsp:=Tob.create('_Les Prospects',nil,-1);

if L.AllSelected then
   begin
   stClauseWhere:=RecupWhereCritere(Pages);
   StSelect := StSelect + stClauseWhere;
   QQ:=OpenSQL(StSelect,TRUE) ;
   if Not QQ.EOF then TobProsp.LoadDetailDB('','','',QQ,False,True);
   Ferme(QQ) ;
   L.AllSelected:=False;
   end else
   begin
   InitMove(L.NbSelected,'');
   for i:=0 to L.NbSelected-1 do
      begin
      MoveCur(False);
      L.GotoLeBookmark(i);
      Tob.create('',TobProsp,-1);
      codeprospect:=TFmul(Ecran).Q.FindField('T_TIERS').asstring ;
      StSelect2 := StSelect + ' WHERE T_TIERS ="'+codeprospect+'"';
      QQ:=OpenSQL(StSelect2,TRUE) ;
      if Not QQ.EOF then TobProsp.detail[i].selectDB('',QQ,false);
      Ferme(QQ) ;
      end;
   L.ClearSelected;
   end;
FiniMove;

{$IFDEF EAGLCLIENT}
{AFAIREEAGL}
{$ELSE}
Ext := ExtractFileExt(Maquette);
if Ext <> '' then
   begin
   if Ext = '.doc' then ConvertDocFile(Maquette,DocGenere,Nil,Nil,Nil,OnGetVar,Nil,Nil)
   else ConvertRTFFile(Maquette,DocGenere,Nil,Nil,Nil,OnGetVar,Nil,Nil) ;
   end;
{$ENDIF}
TobProsp.free;
TobProsp:=Nil;
end ;


procedure TOF_Mailing.OnGetVar( Sender: TObject; VarName: String; VarIndx: Integer;var Value: variant) ;
var TTOB : tob ;
BEGIN
{if VarName='NBPROSPECT' then Value:=TobProsp.Detail.Count else
 if Pos('T_',VarName)>0 then Value:=TobProsp.Detail[VarIndx-1].GetValue(VarName) else
    Value:=TobProsp.Detail[VarIndx-1].Detail[0].GetValue(VarName) ;  }
if VarName='NBPROSPECT' then
   begin
      Value:=TobProsp.Detail.Count;
      exit;
   end;
{if (Pos('T_',VarName)< 0) and (Pos('RPR_',VarName)<0) then exit else
   if Pos('BLOCNOTE',VarName)>0
      then  begin BlobToFile(VarName,TobProsp.Detail[VarIndx-1].GetValue(VarName)); Value:=''; end else
            if ChampToType( VarName)= 'COMBO' then
               begin
               if Pos('T_',VarName)>0 then Value:=RechDom(ChampToTT( VarName),TobProsp.Detail[VarIndx-1].GetValue(VarName),false) else
                Value:=RechDom(ChampToTT( VarName),TobProsp.Detail[VarIndx-1].Detail[0].GetValue(VarName),false)
               end else
               begin
               if Pos('T_',VarName)>0 then Value:=TobProsp.Detail[VarIndx-1].GetValue(VarName) else
                    Value:=TobProsp.Detail[VarIndx-1].Detail[0].GetValue(VarName) ;
               end;    }
{if Pos('YTC_',VarName)>0 then  TTob:=TobProsp.Detail[VarIndx-1].Detail[0].Detail[0]
else if Pos('RPR_',VarName)>0 then  TTob:=TobProsp.Detail[VarIndx-1].Detail[0]
else if ((Pos('T_',VarName)>0) or (Pos('C_',VarName)>0)) then TTob:=TobProsp.Detail[VarIndx-1]
else exit ;
if TTOB.GetValue(VarName) = NULL then TTOB.PutValue(VarName,'');
if Pos('BLOCNOTE',VarName)>0
      then  begin BlobToFile(VarName,TTob.GetValue(VarName)); Value:=''; end else
            if ChampToType( VarName)= 'COMBO' then Value:=RechDom(ChampToTT( VarName),TTob.GetValue(VarName),false)
               else Value:=TTOB.GetValue(VarName) ;  }
if (Pos('T_',VarName)<= 0) and (Pos('RPR_',VarName)<=0) and (Pos('YTC_',VarName)<=0) and (Pos('C_',VarName)<=0) then exit;
TTob:=TobProsp.Detail[VarIndx-1];
if TTOB.GetValue(VarName) = NULL then
   begin
   if ((Pos('RPRLIBVAL',VarName)>0) or (Pos('VALLIBRE',VarName)>0)) then TTOB.PutValue(VarName,0)
   else if ((Pos('RPRLIBDATE',VarName)>0) or (Pos('DATELIBRE',VarName)>0)) then TTOB.PutValue(VarName,IDate1900)
   else TTOB.PutValue(VarName,'');
   end;
if Pos('BLOCNOTE',VarName)>0
      then  begin BlobToFile(VarName,TTob.GetValue(VarName)); Value:=''; end else
            if ChampToType( VarName)= 'COMBO' then Value:=RechDom(ChampToTT( VarName),TTob.GetValue(VarName),false)
               else Value:=TTOB.GetValue(VarName) ;  
END ;

Function TOF_Mailing.MailingValidDoc : Boolean;
var stDocWord : string;
begin
Result:=False;
stDocWord:=getControlText('DOCGENERE');
{stMaquette:=getControlText('MAQUETTE');
if (stMaquette='') or (pos('\*.',stMaquette)<>0) then begin PGIBox('Vous devez choisir une maquette', ecran.caption); exit; end;
if not FileExists(stMaquette) then begin PGIBox('La maquette n''existe pas', ecran.caption); exit; end;   }
if (stDocWord='') then begin PGIBox('Vous devez mentionner un document', ecran.caption); exit; end;
if not DirectoryExists(ExtractFilePath(stDocWord)) then begin PGIBox(Format(TraduireMemoire('Le répertoire %s n''existe pas'),[ExtractFilePath(stDocWord)]), ecran.caption); exit; end;
if FileExists(stDocWord) then
   if PGIAsk('Ce fichier existe déjà, voulez-vous l''écraser ?',ecran.caption)=mryes then DeleteFile(stDocWord ) else exit;
if (TFMul(Ecran).name = 'RTPRO_MAILIN_CONT') then
   begin
   if GetControlText('ENVOIEMAIL') = 'X' then
     begin
     if GetControlText('OBJETMAIL') = '' then begin PGIBox('Vous devez renseigner l''objet du message', ecran.caption); exit; end;
    //   StSql := TFMul(Ecran).Q.sql.Text;
    //   if (pos('C_RVA',StSql) = 0) or (pos('C_RVA',StSql) > pos('FROM',StSql)) then begin PGIBox('Vous devez ajouter le champ E-mail dans la liste', ecran.caption); exit; end;
     if GetControlText('MAIL') <> 'X' then begin PGIBox('Vous devez sélectionner les contacts avec e-mail (Onglet compléments)', ecran.caption); exit; end;
//     if GetControlText('C_EMAILING') <> 'X' then begin PGIBox('Vous devez sélectionner les contacts avec e-mail opt-in', ecran.caption); exit; end;
     end;
   end;
Result:=True;
end;


procedure TOF_Mailing.ChangeMaquette;
var OkOk : Boolean;
begin
if StMaquetteValide <> '' then DeleteFile(pchar(StMaquetteValide));
StMaquetteValide := '';
OkOk := Ctrl_Maquette (GetControlText('MAQUETTE'),Ecran.Caption,stNatureDoc,StMaquetteValide);
//SetControlEnabled('BMAQUETTE',  OkOk) ;
//SetControlEnabled('BOUVRIR',  OkOk) ;
SetControlEnabled('BMAILINGDOC',  OkOk) ;
SetControlEnabled('BEMAILING',  OkOk) ;
{$IF Defined(CRM)}
SetControlEnabled('BESKERCOURRIER',OkOk and GetParamsocSecur('SO_RTGESTIONFLYDOC',False)) ;
SetControlEnabled('BESKEREMAIL',OkOk and GetParamsocSecur('SO_RTGESTIONFLYDOC',False)) ;
SetControlEnabled('BESKERFAX',OkOk and GetParamsocSecur('SO_RTGESTIONFLYDOC',False)) ;
{$IFEND}
end;

procedure TOF_Mailing.ChangeDocument;
var stDocument: string;
    OkOk : boolean ;
begin
OkOk:=False;
stDocument:=getControlText('DOCGENERE');
if (stDocument<>'') and (pos('\*.',stDocument)=0) and  FileExists(stDocument) then OkOk:=true;
SetControlEnabled('BDOCUMENT',OkOk) ;
end;

procedure TOF_Mailing.RTPublipostage;
var StTitre,stDocument,stMaquette,ArgEmail: string;
    DS : THquery;
    TT : TOB;
begin
stDocument:=getControlText('DOCGENERE');

if (TFMul(Ecran).name = 'RTPRO_MAILIN_CONT') then
   begin
   if GetControlText('ENVOIEMAIL') = 'X' then
      begin
      {$IFDEF EAGLCLIENT}
      if TFMul(Ecran).Fetchlestous then
      begin
      {$ENDIF}
      ArgEmail := 'EMAIL;-;ADRESSEEMAIL;';
      ArgEmail := ArgEmail + GetControlText('OBJETMAIL');
      DS:=THquery(TFMul(Ecran).Q) ;
      StTitre :=  TFMul(Ecran).Q.titres ;// GetTitreGrid (TFMul(Ecran).FListe);
      LancePublipostage(ArgEmail,stMaquette,stDocument,Nil,StTitre,DS,True);
  {$IFDEF EAGLCLIENT}
      end;
  {$ENDIF}
      exit;
      end;
   end;

{$IFDEF EAGLCLIENT}
if TFMul(Ecran).Fetchlestous then
   begin
{$ENDIF} //EAGLCLIENT
   DS:=THquery(TFMul(Ecran).Q) ;
   StTitre :=  TFMul(Ecran).Q.titres ;
   LancePublipostage('FILE',stMaquetteValide,stDocument,Nil,sttitre,DS,true);
{$IFDEF EAGLCLIENT}
   end;
{$ENDIF}
end;


function TOF_Mailing.MailingValidMail : Boolean;
begin
{$IFDEF EAGLCLIENT}
Result := False;
{$ENDIF EAGLCLIENT}
{stMaquette:=getControlText('MAQUETTE');
if (stMaquette='') or (pos('\*.',stMaquette)<>0) then begin PGIBox('Vous devez choisir une maquette', ecran.caption); exit; end;
if not FileExists(stMaquette) then begin PGIBox('La maquette n''existe pas', ecran.caption); exit; end; }

{$IFDEF EAGLCLIENT}
if Not TFMul(Ecran).Fetchlestous then exit;
{$ENDIF}
Result := True;
end;

//bur 11110 {$ifndef BUREAU}
procedure TOF_Mailing.OpenDoc_OnClick(Sender: TObject) ;
var OkOk : Boolean;
begin
  AFLanceFiche_DocExterneMul('', 'TYPE=WOR;NATURE='+stNatureDoc+';FORM='+intToStr(longint(ecran)));
  if (GetControlText('MAQUETTE') <> '') and (stMaquetteValide <> '') and (not FileExists(stMaquetteValide)) then
    begin
    StMaquetteValide := '';
    OkOk := Ctrl_Maquette (GetControlText('MAQUETTE'),Ecran.Caption,stNatureDoc,StMaquetteValide);
    SetControlEnabled('BMAILINGDOC',  OkOk) ;
    SetControlEnabled('BEMAILING',  OkOk) ;
{$IF Defined(CRM)}
    SetControlEnabled('BESKERCOURRIER',OkOk and GetParamsocSecur('SO_RTGESTIONFLYDOC',False)) ;
    SetControlEnabled('BESKEREMAIL',OkOk and GetParamsocSecur('SO_RTGESTIONFLYDOC',False)) ;
    SetControlEnabled('BESKERFAX',OkOk and GetParamsocSecur('SO_RTGESTIONFLYDOC',False)) ;
{$IFEND}
    end;
end ;
//bur 11110{$endif}

procedure TOF_Mailing.BEMAILING_OnClick(Sender: TObject);
var StTitre : string;
    DS : THquery;
begin
{$IFDEF EAGLCLIENT}
if Not TFMul(Ecran).Fetchlestous then exit;
{$ENDIF}
StTitre :=  TFMul(Ecran).Q.titres ;
DS := THquery(TFMul(Ecran).Q) ;
LancePublipostage('EMAIL',StMaquetteValide,'',Nil,StTitre,DS,False);
end;

//bur 11110 {$ifndef BUREAU}
procedure TOF_Mailing.MAQUETTE_OnElipsisClick(Sender: TObject);
var NoMaquette : string;
begin
  NoMaquette := AFLanceFiche_DocExterneMulSelec('ADE_DOCEXETAT=UTI', 'TYPE=WOR;NATURE='+stNatureDoc+';SELECTION');
  SetControlText('MAQUETTE',NoMaquette)
end;
//bur 11110 {$endif}

procedure TOF_Mailing.OnClose ;
begin
if StMaquetteValide <> '' then DeleteFile(pchar(StMaquetteValide));
end;


{$IF Defined(CRM)}
function TOF_Mailing.ChpsObligExiste : Boolean;
begin
  Result := False;
  if TypeEnvoiEsker = 'MAIL' then
    begin
    if assigned(TFmul(Ecran).Q.FindField('C_RVA') ) and
      assigned(TFmul(Ecran).Q.FindField('T_EMAIL') ) then Result := True;
    end;
  if TypeEnvoiEsker = 'FAX' then
    begin
    if assigned(TFmul(Ecran).Q.FindField('C_FAX') ) and
      assigned(TFmul(Ecran).Q.FindField('T_FAX') ) then Result := True;
    end;
end;

procedure TOF_Mailing.BESKERCOURRIER_OnClick(Sender: TObject) ;
var Infos : string;
    StTitre : string;
    DS : THquery;
begin
  if TFMul(Ecran).Q.RecordCount = 0 then exit;
  TypeEnvoiEsker := 'COURRIER';
  if MailingValidDoc() = False then Exit;
{$IFDEF EAGLCLIENT}
  if Not TFMul(Ecran).Fetchlestous then exit;
{$ENDIF}
  if LoginEsker = False then Exit;
  TobAdr:=Tob.create('LesAdresses',Nil , -1);
// Préparation de la Pré-Visualisation
  TFMul(Ecran).Q.First;
  TobAdr.AddChampSupValeur ('ADR_JURIDIQUE',TFmul(Ecran).Q.FindField('T_JURIDIQUE').asstring,False);
  TobAdr.AddChampSupValeur ('ADR_LIBELLE',TFmul(Ecran).Q.FindField('T_LIBELLE').asstring,False);
  TobAdr.AddChampSupValeur ('ADR_SUITELIBELLE',TFmul(Ecran).Q.FindField('T_PRENOM').asstring,False);
  TobAdr.AddChampSupValeur ('ADR_ADRESSE1',TFmul(Ecran).Q.FindField('T_ADRESSE1').asstring,False);
  TobAdr.AddChampSupValeur ('ADR_ADRESSE2',TFmul(Ecran).Q.FindField('T_ADRESSE2').asstring,False);
  TobAdr.AddChampSupValeur ('ADR_ADRESSE3',TFmul(Ecran).Q.FindField('T_ADRESSE3').asstring,False);
  TobAdr.AddChampSupValeur ('ADR_CODEPOSTAL',TFmul(Ecran).Q.FindField('T_CODEPOSTAL').asstring,False);
  TobAdr.AddChampSupValeur ('ADR_VILLE',TFmul(Ecran).Q.FindField('T_VILLE').asstring,False);
  TobAdr.AddChampSupValeur ('ADR_PAYS',TFmul(Ecran).Q.FindField('T_PAYS').asstring,False);
  TobAdr.AddChampSupValeur ('C_NOM',TFmul(Ecran).Q.FindField('C_NOM').asstring,False);
  TobAdr.AddChampSupValeur ('C_PRENOM',TFmul(Ecran).Q.FindField('C_PRENOM').asstring,False);
  TypeTraitement := 'PREVISU';
  StTitre :=  TFMul(Ecran).Q.titres ;
  DS := THquery(TFMul(Ecran).Q) ;
  LancePublipostage('SEPARATE',StMaquetteValide,GetControlText('DOCGENERE'),Nil,StTitre,DS,True,'',PublipostageCallBack);
  TobInfos:=Tob.create('LesInfos',Nil , -1);
  TobInfos.AddChampSupValeur ('TYPECOURRIER','M',False);
  Infos := RTLanceFiche_RTSaisieInfosEsker ('RT','RTSAISIECOURRIER','','','Document='+GetControlText('DOCGENERE')+';TobAdr='+intToStr(longint(TobAdr))+';TobInfos='+intToStr(longint(TobInfos)));
  if Infos <> '' then
  begin
    TobProsp:=Tob.create('_Les Prospects',Nil , -1);
    CreatTobInfosMul;
    InitMoveProgressForm(nil,Ecran.caption, 'Traitement en cours, veuillez patienter...', TFMul(Ecran).Q.RecordCount, True, True) ;
    TypeTraitement := 'REEL';
    LancePublipostage('SEPARATE',StMaquetteValide,GetControlText('DOCGENERE'),Nil,StTitre,DS,True,'',PublipostageCallBack);
    FiniMoveProgressForm ;
    FreeAndNil (TobProsp);
  end;
  LogoutEsker;
  FreeAndNil (TobInfos);
  FreeAndNil (TobAdr);
end ;

procedure TOF_Mailing.BESKEREMAIL_OnClick(Sender: TObject) ;
var Infos,Chp : string;
    StTitre : string;
    DS : THquery;
    Q : TQuery;
begin
  if TFMul(Ecran).Q.RecordCount = 0 then exit;
  TypeEnvoiEsker := 'MAIL';
  if MailingValidDoc() = False then Exit;
  if ChpsObligExiste = False then
    begin
    PGIBox (TexteMessage[1],ecran.caption);
    Exit;
    end;
{$IFDEF EAGLCLIENT}
  if Not TFMul(Ecran).Fetchlestous then exit;
{$ENDIF}
  if LoginEsker = False then Exit;
  TobAdr:=Tob.create('LesAdresses',Nil , -1);
  TobInfos:=Tob.create('LesInfos',Nil , -1);
  TobInfos.AddChampSupValeur ('TYPECOURRIER','M',False);
  Infos := RTLanceFiche_RTSaisieInfosEsker ('RT','RTSAISIEEMAIL','','','TobInfos='+intToStr(longint(TobInfos)));
  if Infos <> '' then
  begin
    Q:=OpenSQL('Select ARS_EMAIL FROM RESSOURCE WHERE ARS_UTILASSOCIE="'+V_PGI.User+'"',true);
    if Not Q.EOF then Chp:=Q.FindField('ARS_EMAIL').asstring;
    ferme(Q);
    TobInfos.AddChampSupValeur ('EMAILEMETTEUR',Chp,False);
    TobAdr.AddChampSup ('NOMCONTACT',False);
    TobAdr.AddChampSup ('NOMSOCCONTACT',False);
    TobAdr.AddChampSup ('EMAILCONTACT',False);
    TobProsp:=Tob.create('_Les Prospects',Nil , -1);
    CreatTobInfosMul;
    StTitre :=  TFMul(Ecran).Q.titres ;
    DS := THquery(TFMul(Ecran).Q) ;
    InitMoveProgressForm(nil,Ecran.caption, 'Traitement en cours, veuillez patienter...', TFMul(Ecran).Q.RecordCount, True, True) ;
    LancePublipostage('SEPARATE',StMaquetteValide,GetControlText('DOCGENERE'),Nil,StTitre,DS,True,'',PublipostageCallBack);
    FiniMoveProgressForm ;
    FreeAndNil (TobProsp);
  end;
  LogoutEsker;
  FreeAndNil (TobInfos);
  FreeAndNil (TobAdr);
end ;

procedure TOF_Mailing.BESKERFAX_OnClick(Sender: TObject) ;
var Infos : string;
    StTitre : string;
    DS : THquery;
    Chp : string;
begin
  if TFMul(Ecran).Q.RecordCount = 0 then exit;
  TypeEnvoiEsker := 'FAX';
  if MailingValidDoc() = False then Exit;
  if ChpsObligExiste = False then
    begin
    PGIBox (TexteMessage[2],ecran.caption);
    Exit;
    end;
{$IFDEF EAGLCLIENT}
  if Not TFMul(Ecran).Fetchlestous then exit;
{$ENDIF}
  if LoginEsker = False then Exit;
// Préparation de la Pré-Visualisation
  TypeTraitement := 'PREVISU';
  StTitre :=  TFMul(Ecran).Q.titres ;
  DS := THquery(TFMul(Ecran).Q) ;
  LancePublipostage('SEPARATE',StMaquetteValide,GetControlText('DOCGENERE'),Nil,StTitre,DS,True,'',PublipostageCallBack);
  TobInfos:=Tob.create('LesInfos',Nil , -1);
  TobInfos.AddChampSupValeur ('TYPECOURRIER','M',False);
  TobAdr:=Tob.create('LesAdresses',Nil , -1);
  TFMul(Ecran).Q.First;
  TobAdr.AddChampSup ('NOMCONTACT',False);
  TobAdr.AddChampSup ('NOMSOCCONTACT',False);
  TobAdr.AddChampSup ('FAXCONTACT',False);
  TobAdr.SetString ('NOMCONTACT',TFmul(Ecran).Q.FindField('C_PRENOM').asstring + ' ' + TFmul(Ecran).Q.FindField('C_NOM').asstring);
  TobAdr.SetString ('NOMSOCCONTACT',TFmul(Ecran).Q.FindField('T_LIBELLE').asstring);
  if TFmul(Ecran).Q.FindField('C_FAX').asstring <> '' then Chp := TFmul(Ecran).Q.FindField('C_FAX').asstring
  else Chp := TFmul(Ecran).Q.FindField('T_FAX').asstring;
  TobAdr.SetString ('FAXCONTACT',CleTelephone (Chp));
  Infos := RTLanceFiche_RTSaisieInfosEsker ('RT','RTSAISIEFAX','','','Document='+GetControlText('DOCGENERE')+';TobAdr='+intToStr(longint(TobAdr))+';TobInfos='+intToStr(longint(TobInfos)));
  if Infos <> '' then
  begin
    TobProsp:=Tob.create('_Les Prospects',Nil , -1);
    CreatTobInfosMul;
    InitMoveProgressForm(nil,Ecran.caption, 'Traitement en cours, veuillez patienter...', TFMul(Ecran).Q.RecordCount, True, True) ;
    TypeTraitement := 'REEL';
    LancePublipostage('SEPARATE',StMaquetteValide,GetControlText('DOCGENERE'),Nil,StTitre,DS,True,'',PublipostageCallBack);
    FiniMoveProgressForm ;
    FreeAndNil (TobProsp);
  end;
  LogoutEsker;
  FreeAndNil (TobInfos);
  FreeAndNil (TobAdr);
end ;

procedure TOF_Mailing.PublipostageCallBack(Index: integer; OutputFileName: hstring; var Cancel: boolean);
begin
if TypeTraitement = 'PREVISU' then
  begin
  Cancel := True;
  exit;
  end;
if Not MoveCurProgressForm(format(TexteMessage[3],[IntToStr(Index+1)])) then
  begin
  Cancel := True;
  exit;
  end
else
  begin
  Cancel := false;
  if FileExists(OutputFileName) then
    begin
    if TypeEnvoiEsker = 'COURRIER' then
      begin
      TobAdr.SetString ('ADR_JURIDIQUE',TobProsp.detail[index].GetString('ADR_JURIDIQUE'));
      TobAdr.SetString ('ADR_LIBELLE',TobProsp.detail[index].GetString('ADR_LIBELLE'));
      TobAdr.SetString ('ADR_SUITELIBELLE',TobProsp.detail[index].GetString('ADR_SUITELIBELLE'));
      TobAdr.SetString ('ADR_ADRESSE1',TobProsp.detail[index].GetString('ADR_ADRESSE1'));
      TobAdr.SetString ('ADR_ADRESSE2',TobProsp.detail[index].GetString('ADR_ADRESSE2'));
      TobAdr.SetString ('ADR_ADRESSE3',TobProsp.detail[index].GetString('ADR_ADRESSE3'));
      TobAdr.SetString ('ADR_CODEPOSTAL',TobProsp.detail[index].GetString('ADR_CODEPOSTAL'));
      TobAdr.SetString ('ADR_VILLE',TobProsp.detail[index].GetString('ADR_VILLE'));
      TobAdr.SetString ('ADR_PAYS',TobProsp.detail[index].GetString('ADR_PAYS'));
      TobAdr.SetString ('C_NOM',TobProsp.detail[index].GetString('C_NOM'));
      TobAdr.SetString ('C_PRENOM',TobProsp.detail[index].GetString('C_PRENOM'));
      SendEsker (0,TobInfos,OutputFileName,TobAdr);
      end else
      if TypeEnvoiEsker = 'MAIL' then
        begin
        if TobProsp.detail[index].GetString('EMAIL') <> '' then
          begin
          TobAdr.SetString ('NOMCONTACT',TobProsp.detail[index].GetString('C_NOM'));
          TobAdr.SetString ('NOMSOCCONTACT',TobProsp.detail[index].GetString('ADR_LIBELLE'));
          TobAdr.SetString ('EMAILCONTACT',TobProsp.detail[index].GetString('EMAIL'));
          SendEsker (2,TobInfos,OutputFileName,TobAdr);
          end;
        end else
        if TypeEnvoiEsker = 'FAX' then
          begin
          if TobProsp.detail[index].GetString('FAX') <> '' then
            begin
            TobAdr.SetString ('NOMCONTACT',TobProsp.detail[index].GetString('C_PRENOM') + ' ' + TobProsp.detail[index].GetString('C_NOM'));
            TobAdr.SetString ('NOMSOCCONTACT',TobProsp.detail[index].GetString('ADR_LIBELLE'));
            TobAdr.SetString ('FAXCONTACT',TobProsp.detail[index].GetString('FAX'));
            SendEsker (1,TobInfos,OutputFileName,TobAdr);
            end;
          end;
    DeleteFile(OutputFileName )
    end;
  end;
end;

procedure TOF_Mailing.CreatTobInfosMul ;
var Q : THQuery;
    TobAdrCli : Tob;
    Chp : string;
begin
  Q:= TFMul(Ecran).Q;
  InitMove(Q.RecordCount,'');
  Q.First;
  while Not Q.EOF do
    begin
    MoveCur(False);
    TobAdrCli := Tob.create('',TobProsp,-1);
    TobAdrCli.AddChampSupValeur ('ADR_JURIDIQUE',TFmul(Ecran).Q.FindField('T_JURIDIQUE').asstring,False);
    TobAdrCli.AddChampSupValeur ('ADR_LIBELLE',TFmul(Ecran).Q.FindField('T_LIBELLE').asstring,False);
    TobAdrCli.AddChampSupValeur ('ADR_SUITELIBELLE',TFmul(Ecran).Q.FindField('T_PRENOM').asstring,False);
    TobAdrCli.AddChampSupValeur ('ADR_ADRESSE1',TFmul(Ecran).Q.FindField('T_ADRESSE1').asstring,False);
    TobAdrCli.AddChampSupValeur ('ADR_ADRESSE2',TFmul(Ecran).Q.FindField('T_ADRESSE2').asstring,False);
    TobAdrCli.AddChampSupValeur ('ADR_ADRESSE3',TFmul(Ecran).Q.FindField('T_ADRESSE3').asstring,False);
    TobAdrCli.AddChampSupValeur ('ADR_CODEPOSTAL',TFmul(Ecran).Q.FindField('T_CODEPOSTAL').asstring,False);
    TobAdrCli.AddChampSupValeur ('ADR_VILLE',TFmul(Ecran).Q.FindField('T_VILLE').asstring,False);
    TobAdrCli.AddChampSupValeur ('ADR_PAYS',TFmul(Ecran).Q.FindField('T_PAYS').asstring,False);
    TobAdrCli.AddChampSupValeur ('C_NOM',TFmul(Ecran).Q.FindField('C_NOM').asstring,False);
    TobAdrCli.AddChampSupValeur ('C_PRENOM',TFmul(Ecran).Q.FindField('C_PRENOM').asstring,False);
    if TypeEnvoiEsker = 'MAIL' then
      begin
      if TFmul(Ecran).Q.FindField('C_RVA').asstring <> '' then Chp := TFmul(Ecran).Q.FindField('C_RVA').asstring
      else Chp := TFmul(Ecran).Q.FindField('T_EMAIL').asstring;
      TobAdrCli.AddChampSupValeur ('EMAIL',Chp,False);
      end;
    if TypeEnvoiEsker = 'FAX' then
      begin
      if TFmul(Ecran).Q.FindField('C_FAX').asstring <> '' then Chp := TFmul(Ecran).Q.FindField('C_FAX').asstring
      else Chp := TFmul(Ecran).Q.FindField('T_FAX').asstring;
      TobAdrCli.AddChampSupValeur ('FAX',CleTelephone (Chp),False);
      end;
    Q.Next;
    end;
end;

procedure TOF_Mailing.BESKERSUIVI_OnClick(Sender: TObject) ;
var sHttp : String;
begin
  sHttp := 'https://as1.ondemand.esker.com/ondemand/webaccess/40/documents.aspx';
  LanceWeb(sHttp,False);
end;

{$IFEND}


function Ctrl_Maquette (stMaquette,TitreEcran,NatureDoc: string;var stMaquetteValide : string) : Boolean;
var sOrigine : string;
    TOBDoc : TOB;
begin
Result:=False;
if (stMaquette='') or (not IsNumeric(stMaquette)) then begin PGIBox('Vous devez choisir une maquette', TitreEcran); exit; end;
TOBDoc := TOB.Create('AFDOCEXTERNE', nil, -1);
TOBDoc.LoadDetailDBFromSQL('AFDOCEXTERNE', 'SELECT * FROM AFDOCEXTERNE WHERE ADE_DOCEXNATURE = "'+NatureDoc+'" AND ADE_DOCEXTYPE = "WOR" AND ADE_DOCEXETAT = "UTI" AND ADE_DOCEXCODE = ' + stMaquette);
if TOBDoc.Detail.count = 0 then stMaquette := 'NONDISPO'
else stMaquette := TOBDoc.Detail[0].GetString('ADE_FICHIER');
if stMaquette = '' then
begin
  PGIBox('Aucun fichier associé au document ', TitreEcran);
  FreeAndNil(TOBDoc);
  exit;
end;
if stMaquette = 'NONDISPO' then
begin
  PGIBox('Aucun modèle disponible avec ce numéro ', TitreEcran);
  FreeAndNil(TOBDoc);
  exit;
end;
if stMaquette[1] <> '@' then
begin
  PGIBox('Le fichier n''est pas en base', TitreEcran);
  FreeAndNil(TOBDoc);
  exit;
end;
System.Delete(stMaquette, 1, 1);
if TOBDoc.Detail[0].GetString('ADE_DOSSIER') = '$STD' then
  sOrigine := 'CEG'
else if TOBDoc.Detail[0].GetString('ADE_DOSSIER') = '$DAT' then
  sOrigine := 'STD'
else
  sOrigine :='DOS' ;
stMaquette := ExtractDocBaseFile(ExtractFileName(stMaquette), TOBDoc.Detail[0].GetString('ADE_DOCEXTYPE') + TOBDoc.Detail[0].GetString('ADE_DOCEXNATURE'), sOrigine, FALSE, NatureDoc);
FreeAndNil(TOBDoc);
if not FileExists(stMaquette) then begin PGIBox('La maquette n''existe pas', TitreEcran); exit; end;
StMaquetteValide := stMaquette;
Result := True;
end;

Procedure AGLRTMailing_ChangeMaquette( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFmul) then ToTof:=TFmul(F).LaTof else exit;
  TOF_Mailing(totof).ChangeMaquette;
end;
Procedure AGLRTMailing_ChangeDocument( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFmul) then ToTof:=TFmul(F).LaTof else exit;
  TOF_Mailing(totof).ChangeDocument;
end;

Procedure AGLRTPublipostage( parms: array of variant; nb: integer ) ;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFmul) then ToTof:=TFmul(F).LaTof else exit;
  TOF_Mailing(totof).RTPublipostage;
end;

Function AGLMailingValidMail( parms: array of variant; nb: integer ) : variant;
var  F : TForm ; ToTof : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFmul) then ToTof:=TFmul(F).LaTof else exit;
  result:=TOF_Mailing(totof).MailingValidMail;
end;

Function AGLRTGetTitreGrid( parms: array of variant; nb: integer ) : variant ;
var  F : TForm ;
begin
  Result:='';
  F:=TForm(Longint(Parms[0])) ;
//  if (F is TFmul) then result:=getTitreGrid(TFmul(F).Fliste) else exit;
  if (F is TFmul) then result:=TFMul(F).Q.titres else exit;
end;

Function RTLanceFiche_Mailing(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;


procedure TOF_Mailing.constitueXXWHEREPARC;
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

procedure TOF_Mailing.BVALIDSELClick(Sender: TObject);
begin
  constitueXXWHEREPARC;
  BchercheClick(self);
end;

procedure TOF_Mailing.CodeArticleRech(Sendre: TOBject);
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


procedure TOF_Mailing.ScreenKeYDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 120 then
  begin
    TToolbarButton97(GetControl('BVALIDSEL')).Click;
    Key := 0;
  end;
end;

Initialization
registerclasses([TOF_Mailing]);
RegisterAglProc('EditionMailing',TRUE,2,AGLEditionMailing);
RegisterAglFunc('MailingValidDoc', TRUE , 0, AGLMailingValidDoc);
RegisterAglProc( 'RTMailing_ChangeMaquette', TRUE ,0 , AGLRTMailing_ChangeMaquette);
RegisterAglProc( 'RTMailing_ChangeDocument', TRUE , 0, AGLRTMailing_ChangeDocument);
RegisterAglFunc( 'RTGetTitreGrid', TRUE ,0 , AGLRTGetTitreGrid);
RegisterAglProc( 'RTPublipostage', TRUE ,0 , AGLRTPublipostage);
RegisterAglFunc( 'MailingValidMail', TRUE ,0 , AGLMailingValidMail);
end.
