unit UTOMAnn;
// TOM de la table ANNUAIRE

// DWI 310108 Rajout d'une commande pre-*processeur "annuaire_seul"
//     qui permet d'isoler les unites pour ne gerer que les fiches "annuaire".


interface

uses
  StdCtrls, Controls, Classes, db, forms, sysutils,
{$IFDEF VER150}
  Variants,
{$ENDIF}
  {$IFNDEF DBXPRESS}dbtables,
  {$ELSE}
     {$ifndef EAGLCLIENT}
     uDbxDataSet,
     {$ENDIF}
  {$ENDIF}
  ComCtrls,
  messages, windows,
{$IFDEF EAGLCLIENT}
  MaineAgl, eFiche, eCodePost, utilEagl,
{$ELSE}
  FE_Main, DBCtrls, HDB, Fiche, CodePost, EdtrEtat,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOM, HRichEdt, HRichOLE,
  Dialogs, UTOB, HTB97,
  Web, MailOL, ParamSoc,
  dpTomPerso,
  EventDecla, //LM20070213
  LittleTom
  ;

//const
//  TexteRtfDefaut : String = '{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}'+
//                            '{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}'+
//                            '{\f3\fswiss\fcharset1 Arial;}}'+
//                            '{\colortbl\red0\green0\blue0;\red255\green0\blue0;}'+
//                            '\deflang1036\pard\plain\f3\fs16\par }';



type
  TOM_Annuaire = class(TOM)
    procedure OnArgument(stArgument: string); override;
    procedure OnNewRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnLoadRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnClose; override;

  private
    ZoneCJ        : THEdit;
    ZoneAss1      : THEdit;
    ZoneAss2      : THEdit;
    TOBAnnuBis    : TOB;
    TOBBlob       : TOB;
    bDuringLoad   : Boolean; // pour éviter les OnChange pendant les onload
    bVoirLiens    : Boolean; // pour ne pas lire MENU constamment
    TypeFich, Action, Onglet, Appli: string;
    sGuidPer_c       : String;
    sInfosDossier_c, sDepuisDossier_c: string;
    sCodeDosRef_c : string;
    sJurForme_c   : string;
    sCodeDoublon_c: string; //BM 30/10/2002
    bChangeBlob   : Boolean;
    OldLibelle, OldEmail: string; // valeurs avant modif
    NoDossier     : string;
    bEwsActif     : Boolean;
    bDossierEws   : Boolean;

    sDORFormeAsso_c: string;
    sDFIRegleFisc_c: string;
    sDORSectionBNC_c: string;

    iCapNbTitre_c : integer;

    DPP_ : TLittleTOM ; //LM20070213
    EvDecla : TEventDecla ;//LM20070213
    codeNafIni : string; //, siretIni : string ; 

    DefTabSheet : string ;

    procedure AfficheRegimeMatrimonial(sSituFam_p : string);    
    procedure majcontact;
    procedure ChangeAnnubis(Sender: TObject);
    procedure ChangeBlob(Sender: TObject);
    procedure ChangeCOMPLT(Sender: TObject);
    procedure ChangeNOADM(Sender: TObject);
    procedure BcontactClick(Sender: TObject);
    function  CalculAgePersonne(DateNaiss: TDateTime): string;
    procedure ChargeBlob;
    procedure EnregBlob;
    procedure TraiteChangeTypePer;
    procedure FiltreTypePer;
    //LM20070516 procedure AfficheInfoLienPersonne(GuidPer, ZoneGuidPer, ZoneLibelle: string);
    procedure LiensMajDetention;//(iCapNbTitre_p : integer);
   //LM20070516  procedure ChoisirPersonne(sGuidPer_p, sZoneGuidPer_p, sZoneLibelle_p: string);
    procedure ElipsisPerAssClick(Sender: TObject);
    //LM procedure PerAss_OnExit(Sender: TObject);
    procedure BTSUP_OnClick(Sender: TObject);
    procedure TraiteMajInfoDossier;
    procedure ChargeAnnuBis;
    procedure SauveAnnuBis;
    procedure ChargeCaption;
    procedure CreationChampBlob(TableBlob, GuidPer, NomChamp, Lib: string; T: TOB);
    procedure AfficheRCS;
    procedure AfficheRM;
    procedure BTIERS_OnClick(Sender: TObject);
    procedure BLIENS_OnClick(Sender: TObject);
    procedure BINTERVENANTS_OnClick(Sender: TObject);
    procedure BEVENEMENT_OnClick(Sender: TObject);
    procedure BDELETE_OnClick(Sender: TObject);
    procedure BImprimer_OnClick(Sender: TObject);
    procedure AfficheOngletPPouPMouORG;
    function  CoherenceFormeJuridique(sGuidPerDos_p: string; sAnnForme_p: string; var sJurForme_p: string): boolean;
    function  ChangementFormeJuridique(sAncForme_p, sNouForme_p: string): boolean;
    function  EnregJurExiste(sGuidPerDos_p: string): boolean;
    procedure BValiderClick(Sender: TObject);
    procedure ANN_CHGTADR_OnClick(Sender: TObject);
    procedure AfficheChgtAdr;
    procedure RCSClick(Sender: TObject);
    procedure RMClick(Sender: TObject);
    procedure BEMAIL_OnClick(Sender: TObject);
    procedure ChangeAdresse1;
    procedure ChangeAdresseOld;
    procedure ANN_VOIENO_OnExit(Sender: TObject);
    procedure ANN_VOIETYPE_OnExit(Sender: TObject);
    procedure ANN_VOIENOCOMPL_OnExit(Sender: TObject);
    procedure ANN_VOIENOM_OnExit(Sender: TObject);
    procedure BPERSONNE_OnClick(Sender: TObject);
    procedure ANN_SEXE_OnClick(Sender: TObject);
    procedure ANN_ALCP_OnExit(Sender: TObject);
    procedure ANN_PPPM_OnClick(Sender: TObject);
    procedure ANN_FAMPER_OnClick(Sender: TObject);
    procedure ANN_TYPEPER_OnClick(Sender: TObject);
    procedure ANN_SITUFAM_OnClick(Sender: TObject);

    function  DossierRattache(SGuidPer_p: String): string;
    procedure FiltreCivilite(sCodeSexe_p: string);
    procedure OnFormKeyDown(MySender_p: TObject; var MyKey_p: Word; MyShift_p: TShiftState);

    function  HasQuote (strValue:string):boolean;

    // $$$ JP 28/09/05 - gestion saisie adresse old
    procedure ANB_OLDVOIENO_OnExit(Sender: TObject);
    procedure ANB_OLDVOIETYPE_OnExit(Sender: TObject);
    procedure ANB_OLDVOIENOCOMPL_OnExit(Sender: TObject);
    procedure ANB_OLDVOIENOM_OnExit(Sender: TObject);

    procedure ChargeRegleFisc;
    procedure EnregRegleFisc;

    procedure OnClick_BForme(Sender: TObject);

    //PGR 22/06/2006 Ajout Base Winner dans l'onglet Divers
    procedure UpdateBaseWinner;
    procedure ChangeBaseWinner(Sender: TObject);

    procedure ChangeAdrFiscale (Sender: TObject);//LM20070201
    procedure CtrlChangeAdrFiscale(Sender: TObject);//LM20070201
    procedure SauveDDP_;//LM20070213
    procedure ForceModidier ;//LM20070315
    //LM20070511 procedure GereOnChange ;
    procedure handlerChangeField (sender:TObject) ;//LM20070415

    // MD 20071207 - FQ 11887 désactivé : procedure ANN_PAYSNAISChange (sender:TObject) ;//20070425
    procedure ANN_CODENAFOnExit (sender:TObject) ;//20070425
    procedure ANN_CODENAFOnEnter (sender:TObject) ;//20070511

    procedure VoirSCI ;

    procedure AvertirCodeNaf;
  end;




////////////// IMPLEMENTATION ///////////////
implementation

uses AnnOutils, galOutil, DpJurOutils, YNewMessage, EntDP, galSystem,
{$IFDEF EWS}
     UtileWS,
{$ENDIF}

     ulibwindows, //PGR 22/06/2006 Ajout Base Winner dans l'onglet Divers
{$ifndef Annuaire_seul}
     UtomContact, //mcd 12/2005   pour saisie contact
     TiersUtil,   //mcd 12/2005
     utilPgi,
{$ENDIF}

     DPTofAnnuSel, CLASSAnnuLien,
     UTOFYFormeInsee,
     uSatUtil,
{$IFDEF BUREAU}
     m3fp, // $$$ JP 13/08/07
{$ENDIF}
     TypInfo//LMO20061005
     , TntClasses,
     uToolsRTF ;

/////////////////////////////////////////////////////

{ TOM_Annuaire }

/////////////////////////////////////////////////////

procedure TOM_Annuaire.OnChangeField(F: TField);
var S (*,sFamPer_l*): string;
  QQ: TQuery;
begin
  // on ne fait rien pendant le OnLoadRecord
  if bDuringLoad then exit;

  if (DS.state <> dsBrowse) and (F.FieldName = 'ANN_TYPECIV') then
  begin
    S := GetField('ANN_TYPECIV');
    QQ := OpenSQL('SELECT JTC_CIVCOURTE,JTC_CIVABREGE,JTC_CIVCOUR FROM JUTYPECIVIL WHERE JTC_TYPECIV="' + S + '"', TRUE);
    if not QQ.EOF then
    begin
      SetField('ANN_CV', QQ.FindField('JTC_CIVCOURTE').AsString);
      SetField('ANN_CVA', QQ.FindField('JTC_CIVABREGE').AsString);
      SetField('ANN_CVL', QQ.FindField('JTC_CIVCOUR').AsString);
    end;
    Ferme(QQ);
  end
  else if (F.FieldName = 'ANN_DATENAIS') then
    SetControlText('CALCULAGE', CalculAgePersonne(GetField('ANN_DATENAIS')))
  else if (DS.state <> dsBrowse) and (F.FieldName = 'ANN_TYPEPER') then
    TraiteChangeTypePer
  else if F.FieldName = 'ANN_FAMPER' then
    FiltreTypePer
  else if F.FieldName = 'ANN_PERASS1GUID' then
    AfficheInfoLienPersonne(self, GetField('ANN_PERASS1GUID'), 'ANN_PERASS1GUID', 'TANN_PERASS1GUID') //LM20070516
  else if F.FieldName = 'ANN_PERASS2GUID' then
    AfficheInfoLienPersonne(self, GetField('ANN_PERASS2GUID'), 'ANN_PERASS2GUID', 'TANN_PERASS2GUID');//LM20070516

  if (F.FieldName = 'ANN_RCS') or (F.FieldName = 'ANN_SIREN') or (F.FieldName = 'ANN_RCSVILLE') then
    AfficheRCS;
  if (F.FieldName = 'ANN_RM') or (F.FieldName = 'ANN_SIREN') or (F.FieldName = 'ANN_RMDEP') then
    AfficheRM;

  if (DS.state <> dsBrowse) and (F.FieldName = 'ANN_PAYS') then
  begin
    S := GetField('ANN_PAYS');
    QQ := OpenSQL('SELECT PY_LANGUE,PY_DEVISE FROM PAYS WHERE PY_PAYS="' + S + '"', TRUE);
    if not QQ.EOF then
    begin
      SetField('ANN_NATIONALITE', S);
      SetField('ANN_LANGUE', QQ.FindField('PY_LANGUE').AsString);
      SetField('ANN_DEVISE', QQ.FindField('PY_DEVISE').AsString);
    end;
    Ferme(QQ);
  end;

  if (F.FieldName = 'ANN_PPPM') then
    AfficheOngletPPouPMouORG
  else if (F.FieldName = 'ANN_FAMPER') then
    AfficheOngletPPouPMouORG
  else if (F.FieldName = 'ANN_GUIDCJ') then
  begin
    if GetField('ANN_GUIDCJ') <> NULL then
      AfficheInfoLienPersonne(self, GetField('ANN_GUIDCJ'), 'ANN_GUIDCJ', 'TANN_GUIDCJ');//LM20070516
  end;

  if (F.FieldName = 'ANN_CAPITAL') or (F.FieldName = 'ANN_CAPNBTITRE') then
  begin
      if DS.State <> dsBrowse then
      begin
         if GetField('ANN_CAPNBTITRE') = 0 then
            SetField('ANN_CAPVN', 0)
         else
            SetField('ANN_CAPVN', GetField('ANN_CAPITAL') / GetField('ANN_CAPNBTITRE'));
      end;
  end;

  {$IFDEF EAGLCLIENT}
  if (F.FieldName = 'ANN_FORME') and (DS.State <> dsBrowse) then
    SetControlText('ANN_FORMEPP', GetField('ANN_FORME'));

  if (F.FieldName = 'ANN_CODEINSEE') and (DS.State <> dsBrowse) then
    SetControlText('ANN_CODEINSEEPP', GetField('ANN_CODEINSEE'));

  {$ENDIF EAGLCLIENT}
end;


procedure TOM_Annuaire.OnLoadRecord;
var
  T        : TDateTime;
  visib,TiersValid : Boolean;
{$ifdef EAGLCLIENT}
  i : integer ;
  s, ZoneDoublee,
{$endif}
  sVoieNoCompl_l, sVoieType_l : string;
  Q1       : TQuery;
begin
  bDuringLoad := true;

  // MD : boutons de "liens" actifs uniquement si on a validé la fiche
  if DS.State <> dsInsert then
  begin
    SetControlEnabled('BCONTACTS', True);
    if bVoirLiens then
    begin
      SetControlEnabled('BLIENS', True);
      SetControlEnabled('BINTERVENANTS', True);
    end;
    SetControlEnabled('BEVENEMENT', True); //mcd 10744
    SetControlEnabled('BDOCEVEN', True);
    iCapNbTitre_c := GetField('ANN_CAPNBTITRE');
  end;

  if TypeFich = 'TIERS' then
  begin
    SetControlEnabled('ANN_NATIONALITE', FALSE);
    SetControlEnabled('ANN_LANGUE', FALSE);
    SetControlEnabled('ANN_DEVISE', FALSE);
    SetControlEnabled('ANN_FAX', FALSE);
    SetControlEnabled('ANN_MOISCLOTURE', FALSE);
    SetControlEnabled('ANN_NOMPER', FALSE);
    SetControlEnabled('ANN_FORME', FALSE);
    SetControlEnabled('ANN_FORMEPP', FALSE); //LM20070125
    SetControlEnabled('ANN_SITEWEB', FALSE);
    SetControlEnabled('ANN_EMAIL', FALSE);
    if GetCOntrol('ANN_NATUREAUXI') <> nil then SetControlEnabled('ANN_NATUREAUXI', FALSE);
    SetControlEnabled('ANN_TEL1', FALSE);
    SetControlEnabled('ANN_MINITEL', FALSE);
    SetControlEnabled('ANN_CODENAF', FALSE);
    // mcd 15/04/2005 pour permettre modif car inexistant dans tiers
    //  SetControlEnabled('ANN_VOIENO', FALSE);
    //   SetControlEnabled('ANN_VOIENOCOMPL', FALSE);
    //   SetControlEnabled('ANN_VOIETYPE', FALSE);
    //   SetControlEnabled('ANN_VOIENOM', FALSE);
    SetControlEnabled('ANN_ALRUE1', FALSE);
    SetControlEnabled('ANN_ALRUE2', FALSE);
    SetControlEnabled('ANN_ALRUE3', FALSE);
    SetControlEnabled('ANN_ALCP', FALSE);
    SetControlEnabled('ANN_ALVILLE', FALSE);
    SetControlEnabled('ANN_PAYS', FALSE);
    SetControlEnabled('ANN_NOM1', FALSE);
    SetControlEnabled('ANN_NOM2', FALSE);
    //mcd 23/03/2005 nouveaux champs en synchro
    SetControlEnabled('ANN_ENSEIGNE', FALSE);
    SetControlEnabled('ANN_SEXE', FALSE);
    SetControlEnabled('ANN_DATENAIS', FALSE);
    if GetControl('BTIERS') <> nil then SetControlVisible('BTIERS', FALSE);
  end;

  // ouverture automatique du bloc note si case cochée.
  if (GetField('ANN_NOTEOUV') = 'X') then
  begin
    TPageControl(Getcontrol('Pages')).ActivePage := TTabSheet(GetControl('PBLOCNOTE'));
    TTabSheet(GetControl('PBLOCNOTE')).Visible := true;
  end;

   AfficheRegimeMatrimonial(GetField('ANN_SITUFAM'));
  AfficheRCS;
  AfficheRM;
  AfficheChgtAdr;

  // MD TiersAssoc := '';
  sGuidPer_c := '';

  // --- tronc commun
  ChargeAnnuBis;
  ChargeBlob;
  setControlText('OLE_ACTIVITEPM', getControlText('OLE_ACTIVITE')) ;
  setControlText('OLE_ACTIVITE2PM', getControlText('OLE_ACTIVITE2')) ;

  //LM20070125 ChangeCodeNaf; //mcd 10743 à faire dans tous les cas
  FiltreTypePer; //md
  AfficheOngletPPouPMouORG; //md
  ChargeRegleFisc; //BM

  // Ouverture d'un onglet par défaut
  if (Onglet<>'') and (GetControl(Onglet)<>Nil) then
    TPageControl(GetControl('Pages')).ActivePage := TTabSheet(GetControl(Onglet));

  T := TDateTime(GetField('ANN_DATENAIS'));
  SetControlText('CALCULAGE', CalculAgePersonne(T));

  AfficheInfoLienPersonne(self, GetField('ANN_PERASS1GUID'), 'ANN_PERASS1GUID', 'TANN_PERASS1GUID'); //LM20070516
  AfficheInfoLienPersonne(self, GetField('ANN_PERASS2GUID'), 'ANN_PERASS2GUID', 'TANN_PERASS2GUID'); //LM20070516
  AfficheInfoLienPersonne(self, GetField('ANN_GUIDCJ'), 'ANN_GUIDCJ', 'TANN_GUIDCJ');                //LM20070516

  // MD 6/3/01 - Zones "N° admin" et complémt, en bas onglet "Personne morale"
  visib := Instring(GetField('ANN_FAMPER'),['FIS', 'SOC', 'OGA']);
  SetControlVisible('ORGANISME', visib);
  SetControlText('ANN_COMPLT', GetField('ANN_COMPLTNOADMIN'));
  SetControlText('ANN_NOADM', GetField('ANN_NOADMIN'));

  sJurForme_c := ''; // MD 05/08/02
  CoherenceFormeJuridique(GetField('ANN_GUIDPER'), GetField('ANN_FORME'), sJurForme_c);
  //MD20070709 sinon réplic ann_forme->jur_forme marchait pas ! if sJurForme_c = '' then sJurForme_c := GetField('ANN_FORME');

  // Si un dossier est rattaché, infos visuelles
  NoDossier := DossierRattache(GetField('ANN_GUIDPER'));
//  SetControlVisible('PFAMTYPE', (NoDossier = '')) ;//LMO20061005
  SetControlVisible('PFAMTYPE', not (AnnuaireLanceeDepuisSynthese) or V_PGI.SAV);// CAT 20071207 FQ 11886

  // Stocke les anciennes valeurs de certaines zones qui peuvent changer
  bDossierEws := (bEwsActif) and (NoDossier <> '') and (GetField('ANN_GUIDPER') <> '')
    and ExisteSQL('SELECT 1 FROM DOSSIER WHERE DOS_GUIDPER="' + GetField('ANN_GUIDPER') + '" AND DOS_EWSCREE="X"');
  if bDossierEws then
  begin
    OldLibelle := Copy(Trim(GetField('ANN_NOM1') + ' ' + GetField('ANN_NOM2')), 1, 50);
    OldEmail := GetField('ANN_EMAIL');
  end;

  // performances pour la synchro éventuelle des paramsoc
{$IFDEF DP}
  // $$$ JP 03/01/06 - utile pour les autres que DP???
  if (NoDossier <> '') and DBExists ('DB' + NoDossier) then
     DBSetAutoClose ('DB' + NoDossier, FALSE);
     // (ne peut pas être fait au moment de la synchro, car
     // "ALTER DATABASE statement not allowed within multi-statement transaction")
{$ENDIF}

  if (GetControl('ANN_NATUREAUXI') <> nil) and (GetField('ANN_NATUREAUXI') <> '') and (Getfield('ANN_TIERS') <> '')
    then SetControlEnabled('ANN_NATUREAUXI', false) ; //mcd 16/11/2006 mise ;
(* mcd 16/11/2006 11150
  else if ctxgrc in V_PGI.PGICOntexte then SetControlEnabled('ANN_NATUREAUXI', true); //mcd 23/05/05
fin mcd 16/11/2006 *)

{$ifndef Annuaire_seul}

  if GetControl('BTIERS') <> nil then
  begin
    //gha 11/2007 - FQ 11861.
    TiersValid := (GetField('ANN_TIERS') <> '') and ExistT(GetField('ANN_TIERS'));
    //gha 11/2007 - FQ 11859.
    if getControl('TCODECLIENT') <> nil then
      SetControlVisible('TCODECLIENT', TiersValid);
    if getControl('ANN_TIERS') <> nil then
      SetControlVisible('ANN_TIERS', TiersValid);

    if (GetField('ANN_TIERS') <> '') and (TypeFich <> 'TIERS') and (TiersValid) then
      SetControlVisible('BTIERS', TRUE)
    else
      SetControlVisible('BTIERS', FALSE);
    //mcd 12/2005, on cache bouton Tiers si pas dans la base commune
    //sinon on va aller dans le tiers de la base dossier !!!!
    if not((V_PGI.DbName=V_PGI.DefaultSectionDbName) or (V_PGI.DefaultSectionName='')) then
      SetControlVisible ('BTIERS',false);
  end;

{$endif}


  //mcd 06/06/2005 on affiche le calcul de l'adresse
  sVoieNoCompl_l := RechDom('YYVOIENOCOMPL', GetControlText('ANN_VOIENOCOMPL'), TRUE);
  sVoieType_l := RechDom('YYVOIETYPE', GetControlText('ANN_VOIETYPE'), TRUE);

  if uppercase(sVoieNoCompl_l) = 'ERROR' then
    sVoieNoCompl_l := '';
  if uppercase(sVoieType_l) = 'ERROR' then
    sVoieType_l := '';

  SetCOntrolCaption('ADRESSE', GetControlText('ANN_VOIENO') + ' '
    + sVoieNoCompl_l + ' ' + sVoieType_l + ' ' + GetControlText('ANN_VOIENOM'));

  //PGR 22/06/2006 Ajout Base Winner dans l'onglet Divers
{$ifndef Annuaire_seul}
  if (GetParamSocSecur ('SO_MDLIENWINNER', False)) and (GetControl('DOS_WINSTALL')<>Nil)
  and (GetField('ANN_GUIDPER')<>'') then
  begin
    // si il y a réellement un dossier associé à la fiche annuaire qu'on ouvre
    if ExisteSQL('SELECT 1 FROM DOSSIER WHERE DOS_GUIDPER="'+GetField('ANN_GUIDPER')+'"') then
    begin
      SetControlVisible('LABELWINNER', True);
      SetControlVisible('DOS_WINSTALL', True);
      THEdit(GetControl('DOS_WINSTALL')).OnChange := ChangeBaseWinner;

      Q1 := OpenSql('SELECT DOS_WINSTALL FROM DOSSIER WHERE DOS_GUIDPER="' + GetField('ANN_GUIDPER') + '"', True);
      if not Q1.Eof then
        SetControlText('DOS_WINSTALL', Q1.FindField('DOS_WINSTALL').Asstring);
      Ferme(Q1);

      if (FileExists (GetWindowsTempPath+'\cegideweenvexp.ena')) OR (V_PGI.SAV) then
      begin
        SetControlEnabled('LABELWINNER', True);
        SetControlEnabled('DOS_WINSTALL', True);
      end;
    end;
  end;
{$endif}

  if GetField('ANN_GUIDPER')<>'' then
    DPP_.ChargeEnreg(GetField('ANN_GUIDPER')); //LM20070213
  setActiveTabSheet('GBADRLEGALE') ;//LM20070611
  // $$$ JP 24/04/06

  // MD 20071207 - FQ 11887 désactivé : ANN_PAYSNAISChange(nil) ;
  VoirSCI ;

  {$ifdef EAGLCLIENT}
  ZoneDoublee:= 'ANN_MOISCLOTUREPP,ANN_SIRENPP,ANN_ENSEIGNEPP,ANN_CLESIRETPP,ANN_RCSVILLEPP,ANN_RCSDATEPP,' +
                'ANN_RCSGESTPP,ANN_RMDEPPP,ANN_RMANNEEPP,ANN_FORMEPP,ANN_CODEINSEEPP,ANN_RCSPP,ANN_RMPP,'+
                'ANN_CODENAFPP, ANN_CODENAF2PP';  //LM20070524

  s := '*' ; i:=1 ;//+LM20070524
  while s<>'' do
  begin
    s:=trim(gtfs(ZoneDoublee, ',', i));
    if s<>'' then setcontrolText(s, getField(copy(s, 1, length(s)-2)) ) ;
    inc(i) ;
  end ;            //-LM20070524
  {$endif}

  if DefTabSheet<>'' then SetActiveTabSheet(DefTabSheet); //LM20070528

  // FQ 11699
  AvertirCodeNaf;

  // siretIni := GetField('ANN_SIREN') + GetField('ANN_CLESIRET') ;
  bDuringLoad := false;
end;


procedure TOM_Annuaire.OnNewRecord;
begin
  bDuringLoad := True;
  bChangeBlob := True;

  if DPP_ <>Nil then DPP_.NewEnreg; //LM20070213      +mcd 11/06/07 si fct appelé depuis une fct de création TOM plante

  if GetControl('BCONTACTS') <> nil then
    SetControlEnabled('BCONTACTS', False);
  if bVoirLiens and (GetControl('BLIENS') <> nil) then
    SetControlEnabled('BLIENS', False);
  if bVoirLiens and (GetControl('BINTERVENANTS') <> nil) then
    SetControlEnabled('BINTERVENANTS', False);
  if GetControl('BEVENEMENT') <> nil then
    SetControlEnabled('BEVENEMENT', False);
  if GetControl('BDOCEVEN') <> nil then
    SetControlEnabled('BDOCEVEN', False);
  if GetControl('BPERSONNE') <> nil then
    SetControlEnabled('BPERSONNE', False);

  // BM : 30/05/06 : Par défaut en création Personne morale
  SetField('ANN_PPPM', 'PM');
//  SetField('ANN_PPPM', 'PP');
  SetField('ANN_SEXE', 'M');
  SetField('ANN_RCSDATE', iDate1900);
  SetField('ANN_CAPDEV', 'EUR');
  SetField('ANN_CODEPER', -2); // MD 05/04/06 - pour pb majstruct d'un ancien dossier
  // lorsque la DB0 a déjà des fiches avec guidper rempli et ann_codeper à 0 =>
  // ces fiches étaient prises en compte comme des clés étrangères valides...
  If TobAnnubis <> Nil then ChargeAnnuBis;  //mcd 11/06/07 ajout test fct Nil
  If TobBlob <> Nil then ChargeBlob;  //mcd 11/06/07 ajout test fct Nil

  OldLibelle := '';
  OldEmail := '';
  NoDossier := '';
  SetField('ANN_GUIDPER', '') ;//LMO20061005 ???
  //RP0, 27/11/07, FQ11231 : Alimenter la famille en fonction du contexte
  if (sDepuisDossier_c = 'FIS') or (sDepuisDossier_c = 'SOC') then
    SetField('ANN_FAMPER', sDepuisDossier_c);

  bDuringLoad := false;
end;


procedure TOM_Annuaire.OnUpdateRecord;
var
  St,stdoublon: string;
  Siret: string;
  strNatureAuxi  :string; // $$$ JP 05/05/06 - FQ 106937
  sPPPM, pf, strEns         :string; // $$$ JP 05/05/06 - FQ 10637
begin
  // MD
  if (ds<>nil) then   //mcd 13/06/07 pour Ok si appel via Veriftob, il n'y a pas d'écran
    if DS.State = dsInsert then SetField('ANN_GUIDPER', AglGetGuid());

  if inString(GetField ('ANN_PPPM'),['PP','PM']) then
  begin
    if GetField ('ANN_PPPM')='PP' then pf := 'PP' ;
    if (GetControltext('ANN_SIREN' + pf ) <> '') and (GetControltext('ANN_CLESIRET' + pf ) <> '') then //LM20070502
    begin
      Siret := GetControltext('ANN_SIREN' + pf ) + GetControltext('ANN_CLESIRET' + pf); //LM20070502
      Siret := DP_SupprimerEspace(Siret, SUPPPARTOUT);
      if not CoherenceSiret(Siret, Length(Siret)) then
      begin
        ErreurChamp('ANN_SIREN' + pf, GetField ('ANN_PPPM'), Self, 'Numéro siren ou clé siret incorrect'); // MD
        exit;
      end;
    end;
  end ;

  // MD : il le faut pour afterupdate
  sGuidPer_c := GetField('ANN_GUIDPER');
  if GetField('ANN_PAYS') = '' then
     SetField('ANN_PAYS','FRA') ;


{$ifndef Annuaire_seul}

  // MB Utilisation de la fonction TiersUtil.VerifDoublonSiretPays
  VerifDoublonSiretPays(Self,GetField('ANN_SIREN')+GetField('ANN_CLESIRET'),GetField('ANN_PAYS'),GetField('ANN_NATUREAUXI'),'',sGuidPer_c,'ANB_DOUBLON',LastError,stDoublon);
  if LastError > 1 then
  begin
       ErreurChamp ('ANB_DOUBLON', 'PGeneral', Self, 'Ce numéro de Siret existe déjà sur la fiche annuaire : '+stdoublon);
       exit;
  end;
  // En cas de réponse Non sur la proposition d'accepter quand même "Ce code Siret existe déjà pour la fiche ..."
  if LastError = 1 then exit ;
{$endif}

  // $$$ JP 10/08/05 - certains champs ne doivent pas avoir de " (pb ensuite pour les màj de paramsoc, et autres)
  if HasQuote (GetField ('ANN_NOM1')) = TRUE then
  begin
       ErreurChamp ('ANN_NOM1', 'PGeneral', Self, 'Veuillez ne pas saisir de guillemet dans certaines zones');
       exit;
  end;
  if HasQuote (GetField ('ANN_NOM2')) = TRUE then
  begin
       ErreurChamp ('ANN_NOM2', 'PGeneral', Self, 'Veuillez ne pas saisir de guillemet dans certaines zones');
       exit;
  end;
  if HasQuote (GetField ('ANN_NOMPER')) = TRUE then
  begin
       ErreurChamp ('ANN_NOMPER', 'PGeneral', Self, 'Veuillez ne pas saisir de guillemet dans certaines zones');
       exit;
  end;
  if HasQuote (GetField ('ANN_ALRUE1')) = TRUE then
  begin
       ErreurChamp ('ANN_ALRUE1', 'PGeneral', Self, 'Veuillez ne pas saisir de guillemet dans certaines zones');
       exit;
  end;
  if HasQuote (GetField ('ANN_ALRUE2')) = TRUE then
  begin
       ErreurChamp ('ANN_ALRUE2', 'PGeneral', Self, 'Veuillez ne pas saisir de guillemet dans certaines zones');
       exit;
  end;
  if HasQuote (GetField ('ANN_ALRUE3')) = TRUE then
  begin
       ErreurChamp ('ANN_ALRUE3', 'PGeneral', Self, 'Veuillez ne pas saisir de guillemet dans certaines zones');
       exit;
  end;
  if HasQuote (GetField ('ANN_ALVILLE')) = TRUE then
  begin
       ErreurChamp ('ANN_ALVILLE', 'PGeneral', Self, 'Veuillez ne pas saisir de guillemet dans certaines zones');
       exit;
  end;
  if HasQuote (GetField ('ANN_PAYS')) = TRUE then
  begin
       ErreurChamp ('ANN_PAYS', 'PGeneral', Self, 'Veuillez ne pas saisir de guillemet dans certaines zones');
       exit;
  end;
  if HasQuote (GetField ('ANN_TEL1')) = TRUE then
  begin
       ErreurChamp ('ANN_TEL1', 'PGeneral', Self, 'Veuillez ne pas saisir de guillemet dans certaines zones');
       exit;
  end;
  if HasQuote (GetField ('ANN_TEL2')) = TRUE then
  begin
       ErreurChamp ('ANN_TEL2', 'PGeneral', Self, 'Veuillez ne pas saisir de guillemet dans certaines zones');
       exit;
  end;
  if HasQuote (GetField ('ANN_EMAIL')) = TRUE then
  begin
       ErreurChamp ('ANN_EMAIL', 'PGeneral', Self, 'Veuillez ne pas saisir de guillemet dans certaines zones');
       exit;
  end;
  if HasQuote (GetField ('ANN_RCSVILLE')) = TRUE then
  begin
       ErreurChamp ('ANN_RCSVILLE', 'PGeneral', Self, 'Veuillez ne pas saisir de guillemet dans certaines zones');
       exit;
  end;
  if HasQuote (GetField ('ANN_PAYSNAIS')) = TRUE then
  begin
       ErreurChamp ('ANN_PAYS', 'PGeneral', Self, 'Veuillez ne pas saisir de guillemet dans certaines zones');
       exit;
  end;

{$ifndef Annuaire_seul}
  // $$$ JP 05/07/04 - màj CLETELEPHONE
  SetField('ANN_CLETELEPHONE', CleTelephone(GetField('ANN_TEL1')));
  SetField('ANN_CLETELEPHONE2', CleTelephone(GetField('ANN_TEL2'))); // socref 699 ou 700
{$endif}

{#### Ces contrôles ne sont jamais appelés => voir si besoin de les activer uniquement si bureau ?
  FormGen := THDBValCombobox(GetControl('ANN_FORMEGEN')).value;
  form := THDBValCombobox(GetControl('ANN_FORME')).value;
  if FormGen = '' then
  begin
    ErreurChamp('ANN_FORMEGEN', 'PGeneral', Self, 'La forme juridique est obligatoire.'); // MD
    exit;
  end;
  if ((FormGen = 'SC') or (FormGen = 'DIV') or (FormGen = 'SAGEN')) and (Form = '') then
  begin
    TPageControl(GetControl('Pages')).ActivePage := TTabSheet(GetControl('PGENERAL'));
    setfocusControl('ANN_FORME');
    LastError := 1;
    PGIInfo('La précision est obligatoire !', TitreHalley);
    exit;
  end;

  if GetField('ANN_NOM1') = '' then
  begin
    ErreurChamp('ANN_NOM1', 'PGeneral', Self, 'La dénomination sociale est obligatoire'); // MD
    exit;
  end;

  if (FormGen <> 'SC') and (FormGen <> 'DIV') and (FormGen <> 'SAGEN') then
    SetField('ANN_FORME', FormGen);

  }
  sPPPM := getField('ANN_PPPM') ;

  // Tests sur CODENAF désactivés pour les organismes fiscaux & OGA
  // (qui d'ailleurs n'ont pas l'onglet personne morale ni physique !)
  // #### voir si on doit autoriser la saisie du CodeNaf sur un OGA
  if ( GetField('ANN_FAMPER')<>'FIS' ) and ( GetField('ANN_FAMPER')<>'OGA' ) then
  begin
    St := GetField('ANN_CODENAF');
    if GetControl('TAVERTCODENAF')=Nil then // FQ 11699
    begin
      if St = '' then
      begin
        ErreurChamp('ANN_CODENAF' + iif(sPPPM='PP','PP',''), getField('ANN_PPPM'), Self, 'Le code NAF n''est pas saisi'); // MD
        LastError := 0; // MD20070626 car ErreurChamp le met à 1, et cela signifie validation bloquée en WA !!
        //LM20070524 exit;
      end;
    end;

    if st<>'' then
    begin
      St := DP_SupprimerEspace(St, SUPPFIN);
      if length(St) < 4 then
      begin
        ErreurChamp('ANN_CODENAF' + iif(sPPPM='PP','PP',''), sPPPM, Self, 'Le code NAF doit comporter au moins 4 caractères'); // MD
        exit;
      end;
    end ;
  end;

  if GetField('ANN_NOMPER') = '' then
  begin
    ErreurChamp('ANN_NOMPER', 'PGeneral', Self, 'Le nom abrégé doit être renseigné'); //LM20070524
    exit;
  end;

  if GetField('ANN_NOM1') = '' then
  begin
    ErreurChamp('ANN_NOM1', 'PGeneral', Self, 'La dénomination doit être renseignée');
    exit;
  end;

  // $$$ JP 05/05/06 - FQ 10637: si enseigne en tablette, il faut que le libellé saisie existe
  if (GetField ('ANN_ENSEIGNE') <> '') and (GetParamSocSecur ('SO_GCENSEIGNETAB', FALSE)) then
  begin
       strNatureAuxi := GetField ('ANN_NATUREAUXI');
       if (strNatureAuxi = 'CLI') or (strNatureAuxi = 'PRO') or (strNatureAuxi = 'CON')
{$IFDEF GIGI}
          or (strNatureAuxi = 'NCP')
{$ENDIF}
       then
       begin
            strEns := RechDom ('RTENSEIGNE', GetField ('ANN_ENSEIGNE'), FALSE);
            if (strEns = '') or (strEns = 'Error') then
            begin
                 ErreurChamp ('ANN_ENSEIGNE', 'Activites', Self, 'Veuillez sélectionner une enseigne dans la liste des enseignes');
                 exit;
            end;
       end;
  end;

  //  SetField('ANN_NOPER', 0);
  if TobAnnubis<>Nil then //mcd 13/06/07 si appel depusi veriftob
    if TOBAnnuBis.GetValue('MODIF') = TRUE then SauveAnnuBis;

  // Maj de la forme générique à partir de la forme détaillée
  SetField('ANN_FORMEGEN', RendFormeGen(GetField('ANN_FORME')));

  If TobBlob <>Nil then EnregBlob;    //mcd 09/07/07 ajout test blob à vide poru Ok veriftob
  EnregRegleFisc;

  TraiteMajInfoDossier; // MD 05/08/02 déplacé ici car traitement même si DP !
  LiensMajDetention;

  // Informe si changement de nomper
  if ds<>Nil then   //mcd 13/06/2007 pour Ok via veriftob
    TFFiche(Ecran).Retour := '1;' + sInfosDossier_c + 'JUR_NOMPER=' + GetField('ANN_NOMPER');

  // MD NCX - 28/12/01
  ExecuteSQL('UPDATE ANNULIEN SET ANL_NOMPER="' + GetField('ANN_NOMPER') + '" WHERE ANL_GUIDPER="' + sGuidPer_c +'"');
  //mcd 12/2005 table annuinterloc n'existe plus et ce champs est suppriler
  //mcd 12/2005 ExecuteSQL('UPDATE ANNUINTERLOC SET ANI_NOMPER="' + GetField('ANN_NOMPER') + '" WHERE ANI_GUIDPER="' + GuidPer + '"');
  // #### Aie Juri : risque de conflit sur l'enreg JURIDIQUE ouvert dans la fiche sous-jacente :
  // ExecuteSQL('UPDATE JURIDIQUE SET JUR_NOMPER="' + GetField('ANN_NOMPER') + '" WHERE JUR_GUIDPERDOS="' + GuidPer + '"');

  //PGR 22/06/2006 Ajout Base Winner dans l'onglet Divers
  UpdateBaseWinner;

  // MD 14/01/05 - Si existe un dossier rattaché eWS, met à jour ses infos
    //mcd 13/06/07 ajout ds<>nil pour Ok si appel via Veriftob, il n'y a pas d'écran
  if (ds<>nil)  and (DS.State <> dsInsert) and bDossierEws then // si dossier existant, actif sur eWS
  begin
    St := Copy(Trim(GetField('ANN_NOM1') + ' ' + GetField('ANN_NOM2')), 1, 50);
    // on teste si réellement on a des modifs, car le chargement des objets COM eWS est lourd !
    if (St <> OldLibelle) or (GetField('ANN_EMAIL') <> OldEmail) then
{$IFDEF EWS}
      EwsModifieClient(NoDossier, St, GetField('ANN_EMAIL'));
{$ELSE}
      PGIInfo('Cette fiche annuaire ' + St + ' est liée à un dossier qui possède un espace eWS.' + #13#10
        + 'Il est conseillé d''utiliser une application PGI compatible eWS pour que le dossier (' + NoDossier + ')' + #13#10
        + 'soit mis à jour sur eWS lorsqu''il y a une mise à jour de sa fiche annuaire.');
{$ENDIF}
    // pour le cas où on remodifie sans quitter la fiche
    OldEmail := GetField('ANN_EMAIL');
    OldLibelle := St;
  end;
  if DPP_<>nil then SauveDDP_ ; //mcd 13/06/07 pour Ok veriftob

  if inString (getField('ANN_CODEINSEE'), ['6541','6542','6540']) then else
    setField ('ANN_TYPESCI','');
  if (getField('ANN_CAPDEV')= '') and (getField('ANN_DEVISE')<>'') then
    setField ('ANN_CAPDEV', getField('ANN_DEVISE'));

end;

procedure TOM_Annuaire.SauveDDP_;
begin
  if DPP_.IsModifie then
  begin
    NextPrevControl(Ecran);
    DPP_.tb_.GetEcran(Ecran);//LM20070829...
    //+LM20070620
    DPP_.put('DPP_EXPLOITDOM', iif( (GetControlText('ANN_PPPM')='PM') ,
                                    iif(trim(DPP_.get('DPP_VILLE'))='' , '-', 'X') ,
                                    iif(trim(DPP_.get('DPP_VILLE'))='' , 'X', '-')
                                   )) ;
    (*
    if trim(DPP_.get('DPP_VILLE'))='' then //LM20070511
      DPP_.put('DPP_EXPLOITDOM', 'X') //Tjs ?? LM20070213
    else
      DPP_.put('DPP_EXPLOITDOM', '-') ;//LM20070511
    -LM20070620 *)

    DPP_.Sauve (GetField('ANN_GUIDPER'));//LM20070213
  end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_Annuaire.TraiteMajInfoDossier;
var
  OBDossiers_l: TOB;
  dAnnCapital_l, dJurCapital_l: double;
  sMessage_l, sRequete_l, sInfosDossier_l, sValChamp_l, sNomChamp_l, sAnnNomper_l,
    sAnnGuidPer_l, sAnnForme_l, sJurForme_l, sJurMoisClo_l, sAnnMoisClo_l,
    sAnnCapDev_l, sJurCapDev_l, sJurNbDroitsVote_l, sJurNbTitresClot_l, sAnnCapNbTitre_l,
    sJurCapVn_l, sAnnCapVn_l, sJurNomDos_l, sJurNomper_l : string;
  bChgtFormOk : Boolean;
  nOBInd_l: integer;
begin
   // Récupération des informations de l'annuaire
  dAnnCapital_l := GetField('ANN_CAPITAL');
  sAnnCapDev_l := GetField('ANN_CAPDEV');
  sAnnMoisClo_l := GetField('ANN_MOISCLOTURE');
  sAnnForme_l := GetField('ANN_FORME');
  sAnnGuidPer_l := GetField('ANN_GUIDPER');
  sAnnCapNbTitre_l := GetField('ANN_CAPNBTITRE');
  sAnnCapVn_l := GetField('ANN_CAPVN');
  sAnnNomper_l := GetField('ANN_NOMPER');
  dJurCapital_l := 0;

  bChgtFormOk := False;
  if sAnnForme_l <> sJurForme_c then
  begin
    bChgtFormOk := ChangementFormeJuridique(sJurForme_c, sAnnForme_l);
    if bChgtFormOk then
    begin
      sJurForme_c := sAnnForme_l;
      ModeEdition(DS);
    end;
  end;

  if sDepuisDossier_c = 'JUR' then
  begin
    if sInfosDossier_c = '' then
      exit;

      // Récupération des informations du dossier transmis
    while sInfosDossier_c <> '' do
    begin
      sValChamp_l := ReadTokenSt(sInfosDossier_c);
      sNomChamp_l := ReadTokenPipe(sValChamp_l, '=');
      if sNomChamp_l = 'JUR_NOMDOS' then
        sJurNomDos_l := sValChamp_l
      else if sNomChamp_l = 'JUR_FORME' then
        sJurForme_l := sValChamp_l
      else if sNomChamp_l = 'JUR_CAPDEV' then
        sJurCapDev_l := sValChamp_l
      else if sNomChamp_l = 'JUR_CAPITAL' then
        dJurCapital_l := StrToFloat(sValChamp_l)
      else if (sNomChamp_l = 'JUR_NBDROITSVOTE') then
        sJurNbDroitsVote_l := sValChamp_l
      else if (sNomChamp_l = 'JUR_NBTITRESCLOT') then
        sJurNbTitresClot_l := sValChamp_l
      else if sNomChamp_l = 'JUR_MOISCLOTURE' then
        sJurMoisClo_l := sValChamp_l
      else if sNomChamp_l = 'JUR_VALNOMINCLOT' then
        sJurCapVn_l := sValChamp_l
      else if sNomChamp_l = 'JUR_NOMPER' then
        sJurNomper_l := sValChamp_l;
    end;

      // Traitement du dossier transmis
    sMessage_l := '';
    if sAnnForme_l <> sJurForme_l then
    begin
      if bChgtFormOk then
      begin
        sMessage_l := sMessage_l + '- La forme juridique a changé.#13#10';
        sInfosDossier_l := 'JUR_FORME=' + sAnnForme_l + ';';
      end;
    end;
      // Devise
    if sAnnCapDev_l <> sJurCapDev_l then
    begin
      sMessage_l := sMessage_l + '- La devise a changé.#13#10';// +
//        '  Pensez à vérifier la répartition des titres et des droits de vote entre les associés.#13#10';
      sInfosDossier_l := sInfosDossier_l + 'JUR_CAPDEV=' + sAnnCapDev_l + ';';
    end;
      // Capital
    if dAnnCapital_l <> dJurCapital_l then
    begin
      sMessage_l := sMessage_l + '- Le montant du capital a changé.#13#10';// +
//        '  Pensez à vérifier la répartition des titres et des droits de vote entre les associés.#13#10';
      sInfosDossier_l := sInfosDossier_l + 'JUR_CAPITAL=' + FloatToStr(dAnnCapital_l) + ';';
    end;
    // Titres
    if sAnnCapNbTitre_l <> sJurNbTitresClot_l then
    begin
      sMessage_l := sMessage_l + ' - Le nombre de titre(s) composant le capital a changé.#13#10';// +
//        '  Pensez à vérifier le nombre total des droits de vote,#13#10' +
//        '  et vérifier les fiches de tous les associés.#13#10';

      sInfosDossier_l := sInfosDossier_l + 'JUR_NBTITRESCLOT=' + sAnnCapNbTitre_l + ';';
   //      if sJurNbDroitsVote_l = JurNbTitresClot_l then
      sInfosDossier_l := sInfosDossier_l + 'JUR_NBDROITSVOTE=' + sAnnCapNbTitre_l + ';';
    end;

      // Mois de cloture
    if sAnnMoisClo_l <> sJurMoisClo_l then
    begin
      sMessage_l := sMessage_l + '- Le mois de clôture a changé.#13#10';
      sInfosDossier_l := sInfosDossier_l + 'JUR_MOISCLOTURE=' + sAnnMoisClo_l + ';';
    end;
      // Titres
      // Valeur nominale des titres
    if sAnnCapVn_l <> sJurCapVn_l then
    begin
      sMessage_l := sMessage_l + '- La valeur nominale des titres a changé.#13#10';
      sInfosDossier_l := sInfosDossier_l + 'JUR_VALNOMINCLOT=' + sAnnCapVn_l + ';';
    end;
    // JUR_NOMPER
    if sJurNomper_l <> sAnnNomper_l then
    begin
      sMessage_l := sMessage_l + '- Le nom abrégé de la personne a changé.#13#10';
      sInfosDossier_l := sInfosDossier_l + 'JUR_NOMPER=' + sAnnNomper_l + ';';
    end;
      // Il y a eu des changements
    if (sInfosDossier_l <> '') then
    begin
      sMessage_l := 'Souhaitez-vous mettre à jour les informations du dossier ' +
        sJurNomDos_l + ':#13#10' +
        sMessage_l;

      if (PGIAsk(sMessage_l, TitreHalley) = mrYes) then
      begin
        sInfosDossier_c := sInfosDossier_l;
        ModeEdition(DS);
      end
      else
        sInfosDossier_c := '';
    end;
  end;

  if (sDepuisDossier_c = 'JUR') then //or (sInfosDossier_c = '') then
    Exit;

  // Traitement des dossiers juridiques rattachés possibles
  sRequete_l := 'select JUR_GUIDPERDOS, JUR_TYPEDOS, JUR_NOORDRE, JUR_CODEDOS, ' +
    '       JUR_NOMDOS, JUR_FORME, JUR_NOMPER, ' +
    '       JUR_CAPITAL, JUR_CAPDEV, JUR_MOISCLOTURE, JUR_NBTITRESCLOT, ' +
    '       JUR_VALNOMINCLOT, JUR_NBDROITSVOTE ' +
    'from   JURIDIQUE ' +
    'where  JUR_GUIDPERDOS = "' + sAnnGuidPer_l+'"';

  // Recherche des dossiers correspondants : on ne met pas à jour le dossier
  // passé en paramètre depuis la fiche JURIDIQUE
  if sCodeDosRef_c <> '' then
    sRequete_l := sRequete_l + '  and  JUR_CODEDOS <> "' + sCodeDosRef_c + '"';

  OBDossiers_l := TOB.Create('JURIDIQUE', nil, -1);
//  QRYDossiers_l := OpenSQL(sRequete_l, false);
//  OBDossiers_l.LoadDetailDB('JURIDIQUE', '', '', QRYDossiers_l, false);
  OBDossiers_l.LoadDetailDBFromSQL('JURIDIQUE', sRequete_l);
//  Ferme(QRYDossiers_l);

   // Recherche des dossiers correspondants : on ne met pas à jour le dossier
   // passé en paramètre depuis la fiche JURIDIQUE
  for nOBInd_l := 0 to OBDossiers_l.Detail.Count - 1 do
  begin
    if sDepuisDossier_c = '' then
    begin
      sJurForme_l := OBDossiers_l.Detail[nOBInd_l].GetValue('JUR_FORME');
      dJurCapital_l := OBDossiers_l.Detail[nOBInd_l].GetValue('JUR_CAPITAL');
      sJurCapDev_l := OBDossiers_l.Detail[nOBInd_l].GetValue('JUR_CAPDEV');
      sJurMoisClo_l := OBDossiers_l.Detail[nOBInd_l].GetValue('JUR_MOISCLOTURE');
      sJurNbTitresClot_l := OBDossiers_l.Detail[nOBInd_l].GetValue('JUR_NBTITRESCLOT');
      sJurNbDroitsVote_l := OBDossiers_l.Detail[nOBInd_l].GetValue('JUR_NBDROITSVOTE');
      sJurCapVn_l := OBDossiers_l.Detail[nOBInd_l].GetValue('JUR_VALNOMINCLOT');
      sJurNomper_l := OBDossiers_l.Detail[nOBInd_l].GetValue('JUR_NOMPER');
    end;

    if sAnnForme_l <> sJurForme_l then
    begin
      if bChgtFormOk then
        OBDossiers_l.Detail[nOBInd_l].PutValue('JUR_FORME', sAnnForme_l);
    end;
      // Capital
    if dAnnCapital_l <> dJurCapital_l then
      OBDossiers_l.Detail[nOBInd_l].PutValue('JUR_CAPITAL', dAnnCapital_l);

      // Devise
    if sAnnCapDev_l <> sJurCapDev_l then
      OBDossiers_l.Detail[nOBInd_l].PutValue('JUR_CAPDEV', sAnnCapDev_l);
      // Mois de cloture
    if sAnnMoisClo_l <> sJurMoisClo_l then
      OBDossiers_l.Detail[nOBInd_l].PutValue('JUR_MOISCLOTURE', sAnnMoisClo_l);
      // Titres
    if sAnnCapNbTitre_l <> sJurNbTitresClot_l then
    begin
      OBDossiers_l.Detail[nOBInd_l].PutValue('JUR_NBTITRESCLOT', sAnnCapNbTitre_l);
//            if sJurNbDroitsVote_l = JurNbTitresClot_l then
      OBDossiers_l.Detail[nOBInd_l].PutValue('JUR_NBDROITSVOTE', sAnnCapNbTitre_l);
    end;
      // Valeur nominale des titres
    if sAnnCapVn_l <> sJurCapVn_l then
      OBDossiers_l.Detail[nOBInd_l].PutValue('JUR_VALNOMINCLOT', sAnnCapVn_l);

    if sAnnNomper_l <> sJurNomper_l then
      OBDossiers_l.Detail[nOBInd_l].PutValue('JUR_NOMPER', sAnnNomper_l);

    if OBDossiers_l.Detail[nOBInd_l].Modifie then
      OBDossiers_l.Detail[nOBInd_l].UpdateDB;
  end;

  OBDossiers_l.Free;
end;

procedure TOM_Annuaire.ChargeBlob;
var
  T: TOB;
begin
  TOBBlob.ClearDetail;
  T := TOB.Create('LIENSOLE', TOBBlob, -1);
  CreationChampBlob('ANN', GetField('ANN_GUIDPER'), 'OLE_ACTIVITE', 'Activité', T);
  T := TOB.Create('LIENSOLE', TOBBlob, -1);
  CreationChampBlob('ANN', GetField('ANN_GUIDPER'), 'OLE_ACTIVITE2', 'Activité 2', T);//LM20070712
  T := TOB.Create('LIENSOLE', TOBBlob, -1);
  CreationChampBlob('ANN', GetField('ANN_GUIDPER'), 'OLE_REGMATTXT', 'Régime matrimonial', T);
  T := TOB.Create('LIENSOLE', TOBBlob, -1);
  CreationChampBlob('ANN', GetField('ANN_GUIDPER'), 'OLE_BLOCNOTES', 'Bloc notes', T);
end;


procedure TOM_Annuaire.CreationChampBlob(TableBlob, GuidPer, NomChamp, Lib: string; T: TOB);
//Attention : à synchroniser avec dpJurOutilsBlob.BlobGetCode
var
    ChSql, EmploiBlob, RangBlob : string;
    RSql     : TQuery;
begin
  if instring(NomChamp, ['OLE_ACTIVITE', 'OLE_ACTIVITEPM']) then //LM20070315
    begin
    EmploiBlob := 'ACT';
    RangBlob   := '1';
    end
  else if instring(NomChamp, ['OLE_ACTIVITE2', 'OLE_ACTIVITE2PM']) then 
    begin
    EmploiBlob := 'ACT';
    RangBlob   := '5';
    end
  else if NomChamp = 'OLE_BLOCNOTES' then
    begin
    EmploiBlob := 'BLO';
    RangBlob   := '2';
    end
  else if NomChamp = 'OLE_REGMATTXT' then
    begin
    EmploiBlob := 'REG';
    RangBlob   := '3';
    end
  else if NomChamp = 'OLE_OBJETSOC' then
    begin
    EmploiBlob := 'OBJ';
    RangBlob   := '4';
    end

  else

    // SORTIE !
    exit;

  ChSql := 'SELECT LO_TABLEBLOB,LO_QUALIFIANTBLOB,LO_EMPLOIBLOB,LO_IDENTIFIANT,LO_RANGBLOB,LO_LIBELLE,LO_PRIVE,LO_DATEBLOB,LO_OBJET,LO_DATECREATION,LO_DATEMODIF,LO_CREATEUR,LO_UTILISATEUR' +
    ' FROM ##DP##.LIENSOLE' +
    ' WHERE LO_TABLEBLOB="' + TableBlob + '" AND LO_IDENTIFIANT="' + GuidPer + '" AND LO_RANGBLOB=' + RangBlob ;
  RSql := OpenSql(ChSql, True);

  // test sur <>'' pour le cas d'une fiche en création
  // if (GuidPer<>'') and (T.SelectDB('"'+TableBlob+'";"'+GuidPer+'";'+RangBlob,nil)) then
  //  StringToRich(TRichEdit(GetControl(nomch)),T.GetValue('LO_OBJET'))
  if (GuidPer <> '') and (not Rsql.Eof) then
  begin
    T.PutValue('LO_TABLEBLOB', RSql.FindField('LO_TABLEBLOB').AsString);
    T.PutValue('LO_QUALIFIANTBLOB', RSql.FindField('LO_QUALIFIANTBLOB').AsString);
    T.PutValue('LO_EMPLOIBLOB', RSql.FindField('LO_EMPLOIBLOB').AsString);
    T.PutValue('LO_IDENTIFIANT', RSql.FindField('LO_IDENTIFIANT').AsString);
    T.PutValue('LO_RANGBLOB', RSql.FindField('LO_RANGBLOB').AsInteger);
    T.PutValue('LO_LIBELLE', RSql.FindField('LO_LIBELLE').AsString);
    T.PutValue('LO_PRIVE', RSql.FindField('LO_PRIVE').AsString);
    T.PutValue('LO_DATEBLOB', RSql.FindField('LO_DATEBLOB').AsDateTime);
    T.PutValue('LO_OBJET', RSql.FindField('LO_OBJET').AsString);
    StringToRich(TRichEdit(GetControl(NomChamp)), RSql.FindField('LO_OBJET').AsString);
  end
  else
  begin
    T.PutValue('LO_TABLEBLOB', TableBlob);
    T.PutValue('LO_QUALIFIANTBLOB', 'DAT');
    T.PutValue('LO_EMPLOIBLOB', EmploiBlob); // MD 24/02/06
    T.PutValue('LO_IDENTIFIANT', GetField(TableBlob + '_GUIDPER'));
    T.PutValue('LO_RANGBLOB', RangBlob);
    T.PutValue('LO_LIBELLE', Lib);
    T.PutValue('LO_PRIVE', '-');
    T.PutValue('LO_DATEBLOB', GetField('ANN_DATEMODIF'));
    T.PutValue('LO_OBJET', '');
    StringToRich(TRichEdit(GetControl(NomChamp)), '');
  end;
  Ferme(RSql);
end;
//FIN EPZ 25/10/00

procedure TOM_Annuaire.EnregBlob;
var
    Indice: Integer;
    ChSql, Contenu: string;
    T: TOB;
begin
  //--- On affecte le guidper aux enregistrement des blobs
  if bChangeBlob then
  begin
    if (Appli = 'JUR') then
    begin
       for Indice := 0 to TOBBLob.Detail.Count - 1 do
       begin
         TOBBLob.Detail[Indice].PutValue('LO_IDENTIFIANT', GetField('ANN_GUIDPER'));
         TOBBLob.Detail[Indice].InsertOrUpdateDB;
       end;
    end
    else
    begin
       for Indice := 0 to TOBBLob.Detail.Count - 1 do
       begin
         T := TOBBLob.Detail[Indice];
         T.PutValue('LO_IDENTIFIANT', GetField('ANN_GUIDPER'));

         //--- On remplace les guillemets sinon erreur lors des insert SQL
         Contenu := T.GetValue('LO_OBJET');
         EntreCote(Contenu, False);
         Contenu := Trim(Contenu);

         //--- Maj de la base
         ChSql := 'SELECT 1 FROM ##DP##.LIENSOLE WHERE LO_TABLEBLOB="' + T.GetValue('LO_TABLEBLOB') + '" AND LO_IDENTIFIANT="' + T.GetValue('LO_IDENTIFIANT') + '" AND LO_RANGBLOB=' + IntToStr(T.GetValue('LO_RANGBLOB'));

         // MD 20/07/2006 - Cette syntaxe sur les requêtes est OK en MsSQL et Oracle,
         // mais ne marche pas en DB2, donc on la réserve aux cas Pcl Multidossier
         if not((V_PGI.DbName=V_PGI.DefaultSectionDbName) or (V_PGI.DefaultSectionName='')) then
           begin
           if ExisteSQL(Chsql) then
             ChSql := 'UPDATE ##DP##.LIENSOLE SET LO_TABLEBLOB="' + T.GetValue('LO_TABLEBLOB') + '", LO_QUALIFIANTBLOB="' + T.GetValue('LO_QUALIFIANTBLOB') + '", LO_EMPLOIBLOB="' + T.GetValue('LO_EMPLOIBLOB') + '",' +
               'LO_IDENTIFIANT="' + T.GetValue('LO_IDENTIFIANT') + '", LO_RANGBLOB=' + IntToStr(T.GetValue('LO_RANGBLOB')) + ', LO_LIBELLE="' + T.GetValue('LO_LIBELLE') + '", LO_PRIVE="' + T.GetValue('LO_PRIVE') +
               '", LO_DATEBLOB="' + UsDateTime(T.GetValue('LO_DATEBLOB')) + '", LO_OBJET="' + Contenu + '", LO_DATECREATION="' + UsDateTime(T.GetValue('LO_DATECREATION')) + '", LO_DATEMODIF="' + UsDateTime(T.GetValue('LO_DATEMODIF')) + //JL 15/03/2005 Modifs format date pour oracle
               '", LO_CREATEUR="' + T.GetValue('LO_CREATEUR') + '", LO_UTILISATEUR="' + T.GetValue('LO_UTILISATEUR') + '"' +
               ' WHERE LO_TABLEBLOB="' + T.GetValue('LO_TABLEBLOB') + '" AND LO_IDENTIFIANT="' + T.GetValue('LO_IDENTIFIANT') + '" AND LO_RANGBLOB=' + IntToStr(T.GetValue('LO_RANGBLOB'))
           else
             ChSql := 'INSERT INTO ##DP##.LIENSOLE (LO_TABLEBLOB, LO_QUALIFIANTBLOB, LO_EMPLOIBLOB, LO_IDENTIFIANT, LO_RANGBLOB, LO_LIBELLE, LO_PRIVE, LO_DATEBLOB, LO_OBJET, LO_DATECREATION, LO_DATEMODIF, LO_CREATEUR, LO_UTILISATEUR) ' +
               'VALUES ("' + T.GetValue('LO_TABLEBLOB') + '","' + T.GetValue('LO_QUALIFIANTBLOB') + '","' + T.GetValue('LO_EMPLOIBLOB') + '","' + T.GetValue('LO_IDENTIFIANT') + '",' + IntToStr(T.GetValue('LO_RANGBLOB')) + ',"' + T.GetValue('LO_LIBELLE') + '","' +
               T.GetValue('LO_PRIVE') + '","' + UsDateTime(T.GetValue('LO_DATEBLOB')) + '","' + Contenu + '","' + UsDateTime(T.GetValue('LO_DATECREATION')) + '","' + UsDateTime(T.GetValue('LO_DATEMODIF')) + '","' + T.GetValue('LO_CREATEUR') + '","' + T.GetValue('LO_UTILISATEUR') + '")';

           ExecuteSql(ChSql);
           end
         // mise à jour directe car on est déjà sur la bonne base
         else
           T.InsertOrUpdateDB;
       end;
    end;
  end;

  bChangeBlob := False;
end;

function TOM_Annuaire.CalculAgePersonne(DateNaiss: TDateTime): string;
var
  A1, M1, J1, A2, M2, J2, NbMois: Word;
begin
  Result := ' ';
  if (DateNaiss = iDate1900) or (DateNaiss > Date) then
    exit;

  DecodeDate(DateNaiss, A1, M1, J1);
  DecodeDate(Date, A2, M2, J2);
  NbMois := 12 * (A2 - A1) + (M2 - M1);
  if J2 < J1 then
    NbMois := NbMois - 1;

  Result := Format('%d', [NbMois div 12]);
end;


procedure TOM_Annuaire.TraiteChangeTypePer;
var
  Q: TQuery;
  sSelect, sFrom, sTypePer_l: string;
  //LM vue:boolean ;
begin
  sTypePer_l := GetField('ANN_TYPEPER');
  sSelect := 'select JTP_DEFAUTLIB1, JTP_DEFAUTLIB2, JTP_DEFAUTLIB3, JTP_DEFAUTNOM, JTP_FAMPER ';
  sFrom := 'from JUTYPEPER where JTP_TYPEPER = "' + sTypePer_l + '"';
  Q := OpenSQL(sSelect + sFrom, true);

  if not Q.EOF then
  begin
    if GetField('ANN_NOM1') = '' then
      SetField('ANN_NOM1', Q.FindField('JTP_DEFAUTLIB1').AsString);
    if GetField('ANN_NOM2') = '' then
      SetField('ANN_NOM2', Q.FindField('JTP_DEFAUTLIB2').AsString);
    if GetField('ANN_NOM3') = '' then
      SetField('ANN_NOM3', Q.FindField('JTP_DEFAUTLIB3').AsString);
    if GetField('ANN_NOMPER') = '' then
      SetField('ANN_NOMPER', Q.FindField('JTP_DEFAUTNOM').AsString);
    SetField('ANN_FAMPER', Q.FindField('JTP_FAMPER').AsString);
(* mcd 16/11/2006 11150
    if (GetControl('ANN_NATUREAUXI') <> nil) and (GetField('ANN_NATUREAUXI') = '') then
      SetField('ANN_NATUREAUXI', Q.FindField('JTP_NATUREAUXI').AsString);  fin mcd *)
  end;
  Ferme(Q);

end;


procedure TOM_Annuaire.FiltreTypePer;
begin
   if GetField('ANN_FAMPER') <> '' then
      THDBValCombobox(GetControl('ANN_TYPEPER')).Plus := 'JTP_FAMPER = "' + GetField('ANN_FAMPER') + '"'
   else
      THDBValCombobox(GetControl('ANN_TYPEPER')).Plus := '';
   THDBValCombobox(GetControl('ANN_TYPEPER')).Reload ;
   AvertirCodeNaf; // FQ 11699
end;


procedure TOM_Annuaire.AfficheRCS;
var
  T: string;
begin
  if Boolean(GetField('ANN_RCS') = 'X') then
    T := string(GetField('ANN_SIREN')) + ' RCS ' + string(GetField('ANN_RCSVILLE'))
  else
    T:='' ;
  SetControlText('TANN_RCSNO', T) ;
  SetControlText('TANN_RCSNOPP', T) ;
end;


procedure TOM_Annuaire.AfficheRM;
var
  T: string;
begin
  if Boolean(GetField('ANN_RM') = 'X') then
    T := string(GetField('ANN_SIREN')) + ' RM ' + string(GetField('ANN_RMDEP'))
  else
    T:='' ;
  SetControlText('TANN_RMNO', T);
  SetControlText('TANN_RMNOPP', T);
end;

{LM20070516
procedure TOM_Annuaire.AfficheInfoLienPersonne(GuidPer, ZoneGuidPer, ZoneLibelle: string);
var
  NomPer, Libelle, TypePer: string;
begin
  TraiteChoixPersonne(GuidPer, Libelle, NomPer, TypePer);
  if GetControl(ZoneLibelle + 'SOV') <> nil then
    SetControlText(ZoneLibelle + 'SOV', Libelle);
  SetControlText(ZoneLibelle, Libelle);
end;
}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MD
Créé le ...... : 04/11/2005
Modifié le ... :   /  /
Description .. : Arguments de lancement (tous facultatifs) :
Suite ........ : 1. ACTION=...
Suite ........ : 2. TypeFich :
Suite ........ :       '' pour ouvrir fiche YY ANNUAIRE
Suite ........ :       'DP' pour ouvrir DP FICHESIGNALE => obsolète
Suite ........ :       'TIERS' pour ouvrir YY ANNUAIRE depuis les tiers
Suite ........ : 3. Onglet : nom de l'onglet à ouvrir directement
Suite ........ : 4. CodeDosRef : code dossier juridique de référence
Suite ........ : 5. DepuisDossier : 'JUR' pour savoir d'où on vient
Suite ........ : 6. InfosDossier : NomPer du dossier juridique
Mots clefs ... :
*****************************************************************}
procedure TOM_Annuaire.OnArgument(stArgument: string);
var ctrl,ctrlRef : TControl ;
    st, stAdrFis : string ;
    i:integer ;
    ads : tdataset ;
begin
  inherited; // traite ACTION=, mais ne l'enlève pas de stArgument
  bDuringLoad := True;
  bChangeBlob := False;

  DefTabSheet := ArgumentToTabsheet (stArgument) ;//LM20070528

  //+LM20070213
  EvDecla := TEventDecla.create(Ecran) ; //empile les evts existants de la fiche décla
  ads:=ds ;
  DPP_:=TLittleTOM.Create(Ecran, 'DPPERSO', self, @ads);
  DPP_.ExclusVerifModif := 'DPP_EXPLOITDOM' ;

  stAdrFis :='DPP_VOIENO;DPP_VOIENOM;DPP_VOIETYPE;DPP_VOIENOCOMPL;' ;
  for i:=1 to 4 do //pour la construction d'adresse1
    EvDecla.Rebranche(gtfs(stAdrFis, ';', i), 'OnExit', ChangeAdrFiscale); //Affecte le OnExit du contrôle (le OnExit initial existe tjs)
  ctrl:=getControl('TabAdresseFisc');
  if ctrl<>nil then //pour le contrôle de modif sur l'adresse fiscale
    for i:=0 to TTabSheet(ctrl).ControlCount-1 do
    begin
      st :=TWinControl(ctrl).Controls[i].Name ;
      if (ChampToNum(st)>-1) and (pos(st+';', stAdrFis)=0) then //On ne réaffecte pas d'evt à ce qui a été attribué juste avant
        EvDecla.Rebranche(st, 'OnExit', CtrlChangeAdrFiscale);
    end ;
  //-LM20070213
  // MD 20071207 - FQ 11887 désactivé EvDecla.Rebranche('ANN_PAYSNAIS', 'OnExit', ANN_PAYSNAISChange);
  EvDecla.Rebranche('ANN_CODENAFPP', 'OnExit', ANN_CODENAFOnExit);
  EvDecla.Rebranche('ANN_CODENAF2PP', 'OnExit', ANN_CODENAFOnExit);
  EvDecla.Rebranche('ANN_CODENAFPP', 'OnEnter', ANN_CODENAFOnEnter);
  EvDecla.Rebranche('ANN_CODENAF2PP', 'OnEnter', ANN_CODENAFOnEnter);
  EvDecla.Rebranche('ANN_CODENAF', 'OnExit', ANN_CODENAFOnExit);//LM20070712
  EvDecla.Rebranche('ANN_CODENAF2', 'OnExit', ANN_CODENAFOnExit);//LM20070712

  // eWS actif dans les paramètres d'environnement
  bEwsActif := ExisteSQL('SELECT 1 FROM ##DP##.PARAMSOC WHERE SOC_NOM="SO_NE_EWSACTIF" AND SOC_DATA="X"');
  bDossierEws := False;

  GereDroitsConceptsAnnuaire(Self, Ecran);
  bVoirLiens := JaiLeDroitConceptBureau(ccVoirLesLiens);

  TToolbarButton97(GetControl('BCONTACTS')).OnClick := BContactClick;  //mcd 12/2005

  // Bouton spécifique à la fiche annuaire
  SetControlVisible('BLIENS', (TypeFich <> 'TIERS') and bVoirLiens); //EPZ 22/06/00
  if AnnuaireLanceeDepuisSynthese then
    SetControlVisible('BINTERVENANTS', (TypeFich <> 'TIERS') and bVoirLiens); //MD 11/07/07

  Action := ReadTokenSt(stArgument);
  if (Action = 'ACTION=CONSULTATION') or (TFFiche(Ecran).FTypeAction = taConsult) then
    // ... le 2ème test nécessaire car état modifié suite à GereDroitsConceptsAnnuaire
  begin
    if GetControl('BCOPIEADRSLEGALE') <> nil then
      // pas d'accès au bouton de modif de l'adresse postale
      SetControlEnabled('BCOPIEADRSLEGALE', False);
      SetControlEnabled('BFORME', false);
      SetControlEnabled('BFORME2', false);//LM20070125
      SetControlEnabled('BWEB', false);
      SetControlEnabled('BEMAIL', false);
      SetControlEnabled('BTSUPPERASS1GUID', false);
      SetControlEnabled('BTSUPPERASS2GUID', false);
      SetControlEnabled('BTSUPGUIDCJ', false);
  end;

  //mcd 12/2005, on cache bouton contact si pas dans la base commune
  if not((V_PGI.DbName=V_PGI.DefaultSectionDbName) or (V_PGI.DefaultSectionName='')) then
    SetControlVisible('BCONTACTS',false);

  // TypeFich='' pour ouvrir fiche YY ANNUAIRE
  //          'DP' pour ouvrir DP FICHESIGNALE => obsolète
  //          'TIERS' pour ouvrir YY ANNUAIRE depuis les tiers
  TypeFich := ReadTokenSt(stArgument);
  Onglet := ReadTokenSt(stArgument);
  // appel par application juri ou selection lanceur dos
  Appli := ReadTokenSt(stArgument);
  sCodeDosRef_c := ReadTokenSt(stArgument); //EPZ/NCX 14/12/2000
  sDepuisDossier_c := ReadTokenSt(stArgument);
  sInfosDossier_c := stArgument;

  TOBBlob := TOB.Create('La table des blobs', nil, -1); //EPZ 25/10/00
  //TOBBlob := TOB.Create('LIENSOLE',nil,-1);
  TOBAnnuBis := TOB.Create('ANNUBIS', nil, -1);
  ZoneCJ := THEdit(GetControl('TANN_GUIDCJ'));
  ZoneCJ.OnElipsisClick := ElipsisPerAssClick;

  //mcd 06/06/2005 .. déplacé, pour faire dans tous les cas
  //if GetControl('ANN_CHGTADR') <> nil then TCheckBox(GetControl('ANN_CHGTADR')).OnClick := ANN_CHGTADR_OnClick;
  THEdit(GetControl('ANN_VOIENO')).OnExit := ANN_VOIENO_OnExit;
  THEdit(GetControl('ANN_VOIENOM')).OnExit := ANN_VOIENOM_OnExit;
  THValComboBox(GetControl('ANN_VOIETYPE')).OnExit := ANN_VOIETYPE_OnExit;
  THValComboBox(GetControl('ANN_VOIENOCOMPL')).OnExit := ANN_VOIENOCOMPL_OnExit;
  if GetControl('BPERSONNE') <> nil then TToolBarButton97(GetControl('BPERSONNE')).OnClick := BPERSONNE_OnClick;

  TButton(GetControl('BTIERS')).OnClick := BTIERS_OnClick; //mcd 30/05/05
  ZoneAss1 := THEdit(GetControl('TANN_PERASS1GUID'));
  ZoneAss2 := THEdit(GetControl('TANN_PERASS2GUID'));
  ZoneAss1.OnElipsisClick := ElipsisPerAssClick;
  ZoneAss2.OnElipsisClick := ElipsisPerAssClick;
  ZoneAss2.ReadOnly := true;
  THEdit(GetControl('TANN_GUIDCJ')).OnElipsisClick := ElipsisPerAssClick;
  // BM 18/01/2006 : RAZ représentant personne morale, personne physique, conjoint
  if GetControl('BTSUPPERASS1GUID') <> nil then
  begin
    TToolBarButton97(GetControl('BTSUPPERASS1GUID')).OnClick := BTSUP_OnClick;
    TToolBarButton97(GetControl('BTSUPPERASS2GUID')).OnClick := BTSUP_OnClick;
    TToolBarButton97(GetControl('BTSUPGUIDCJ')).OnClick := BTSUP_OnClick;
  end;

  THRichEditOle(GetControl('OLE_REGMATTXT')).OnChange := ChangeBlob;
  THRichEditOle(GetControl('OLE_BLOCNOTES')).OnChange := ChangeBlob;
  THEdit(GetControl('ANN_COMPLT')).OnChange := ChangeCOMPLT;
  THEdit(GetControl('ANN_NOADM')).OnChange := ChangeNOADM;
  THDBValComboBox(GetControl('ANN_FAMPER')).OnClick := ANN_FAMPER_OnClick;
  THDBValComboBox(GetControl('ANN_TYPEPER')).OnClick := ANN_TYPEPER_OnClick;
  THDBValComboBox(GetControl('ANN_SITUFAM')).OnClick := ANN_SITUFAM_OnClick;

  ChargeCaption;

  // On n'accède qu'à la forme détaillée
  if GetControl('ANN_FORMEGEN') <> nil then
  begin
    SetControlVisible('ANN_FORMEGEN', false);
    SetControlVisible('LabelForme', false);
    SetControlText('TANN_FORME', 'Forme');
  end;

  EvDecla.Rebranche('ANN_RCS',  'OnClick', RCSClick);  //LM20070627
  EvDecla.Rebranche('ANN_RCSPP','OnClick', RCSClick);  //LM20070627
  EvDecla.Rebranche('ANN_RM',   'OnClick', RMClick);
  EvDecla.Rebranche('ANN_RMPP', 'OnClick', RMClick);

  if GetParamsoc('SO_GCCONTROLESIREN') <> 'DOU' then
  begin
     if GetControl('ANB_DOUBLON') <> nil then
        SetControlVisible ('ANB_DOUBLON', False);
     if GetControl('ANP_TABLELIBRE') <> nil then
        SetControlVisible ('ANP_TABLELIBRE', False);
  end ;


  // Accès limité
  (*+LM20070424
  if (TypeFich='TIERS') or (TypeFich = 'DP') then
  begin
    SetControlVisible('BDELETE', false); //mcd 05/07/2005 10752
    SetControlVisible('BEVENEMENT', false); //mcd 10744
    SetControlVisible('BINSERT', False); // MD 20/06/01
  end;*)

  {+GUH - 12/2007 Mise en commentaire FQ 11858.}
  // setControlVisible('bInsert',    Ecran.parent=nil ) ;
  // setControlVisible('bDelete',    Ecran.parent=nil ) ;
  {-GUH - 12/2007 }

  setControlVisible('bEvenement', Ecran.parent=nil ) ;
  //-LM20070424

{$IFDEF EAGLCLIENT}
  // MD 03/02/06 - La suppression peut faire appel à SuppressionDossier
  // de AnnOutils si la fiche est reliée à un dossier, mais pour l'instant
  // la suppression physique de la base nécessite le plug-in de VH_DP
  {$IFNDEF DP}
  SetControlVisible('BDELETE', False);
  {$ENDIF}
{$ENDIF}

  TToolBarButton97(GetControl('BEMAIL')).OnClick := BEMAIL_OnClick;
  THRichEditOle(GetControl('OLE_ACTIVITE')).onChange := ChangeBlob;
  THRichEditOle(GetControl('OLE_ACTIVITEPM')).onChange := ChangeBlob;//LM20070315
  THRichEditOle(GetControl('OLE_ACTIVITE2')).onChange := ChangeBlob;
  THRichEditOle(GetControl('OLE_ACTIVITE2PM')).onChange := ChangeBlob;

  THEdit(GetControl('ANN_ALCP')).OnExit := ANN_ALCP_OnExit;

  // Appel depuis le menu Annuaire du cabinet : ;DOS
  // #### mais par contre l'appel depuis le bouton identité ne met pas ;DOS ! why ?
  // #### du coup BLIENS n'ouvre pas DP LIENPERSONNE mais JUR ANNULIEN_LST
  if Appli = 'DOS' then
  begin
    TButton(GetControl('BLIENS')).OnClick := BLIENS_OnClick;
    TButton(GetControl('BEVENEMENT')).OnClick := BEVENEMENT_OnClick;
  end;
  // MD 11/07/07
  if GetControl('BINTERVENANTS')<>Nil then
    TButton(GetControl('BINTERVENANTS')).OnClick := BINTERVENANTS_OnClick;

  TButton(GetControl('BDELETE')).OnClick := BDELETE_OnClick;

  TButton(GetControl('BIMPRIMER')).OnClick := BIMPRIMER_OnClick;

  // BM 31/07/02
  TButton(GetControl('BVALIDER')).OnClick := BValiderClick;

  THDbValComboBox(GetControl('ANN_PPPM')).OnClick := ANN_PPPM_OnClick;
  THDbValComboBox(GetControl('ANN_SEXE')).OnClick := ANN_SEXE_OnClick;

  //mcd 19/05/05
  if Getcontrol('ANN_NATUREAUXI') <> nil then
  begin
    SetControlVisible('ANN_NATUREAUXI', false);
    SetControlVisible('TANN_NATUREAUXI', false);
(* mcd 16/11/2006 11150
{$IFDEF GRC}
    if ctxGRC in V_PGI.PGICOntexte then
    begin
      SetControlVisible('ANN_NATUREAUXI', true);
      SetControlVisible('TANN_NATUREAUXI', true);
    end;
{$ENDIF} fin mcd 16/11/2006*)

  end;

  // $$$ JP 28/09/05 - gestion saisie adresse old
  THEdit (GetControl ('ANB_OLDVOIENO')).OnExit := ANB_OLDVOIENO_OnExit;
  THEdit (GetControl ('ANB_OLDVOIENOM')).OnExit := ANB_OLDVOIENOM_OnExit;
  THValComboBox (GetControl('ANB_OLDVOIETYPE')).OnExit := ANB_OLDVOIETYPE_OnExit;
  THValComboBox (GetControl('ANB_OLDVOIENOCOMPL')).OnExit := ANB_OLDVOIENOCOMPL_OnExit;

  // $$$ JP 10/06/2005 - montrer les boutons CTI si cti actif
{$IFDEF BUREAU}
  SetControlEnabled('BTEL1', VH_DP.ctiAlerte <> nil);
  SetControlEnabled('BTEL2', VH_DP.ctiAlerte <> nil);
{$ELSE}
  SetControlVisible('BTEL1', FALSE);
  SetControlVisible('BTEL2', FALSE);
{$ENDIF}

  // $$$ JP 05/05/06 - FQ 10637: tablette des enseignes
  if GetParamSocSecur ('SO_GCENSEIGNETAB', FALSE) then
  begin
       SetControlProperty ('ANN_ENSEIGNE', 'DataType', 'RTENSEIGNE');
       SetControlProperty ('ANN_ENSEIGNE', 'ElipsisButton', TRUE);
  end;

  TToolBarButton97(GetControl('BFORME')).OnClick := OnClick_BForme;
  TToolBarButton97(GetControl('BFORME2')).OnClick := OnClick_BForme;

  EvDecla.Rebranche(Ecran.Name, 'OnKeyDown', OnFormKeyDown); //LM20070502 Ecran.OnKeyDown := OnFormKeyDown;

  //+LM20070125
  ctrlRef := getControl('pages') ;
  if ctrlRef<>nil then
  begin
    ctrl := getControl('TNODOSSIER') ; if ctrl<>nil then ctrl.parent := TWinControl(ctrlRef) ;
    ctrl := getControl('NODOSSIER') ; if ctrl<>nil then ctrl.parent := TWinControl(ctrlRef) ;
    //gha 11/2007 FQ 11859
    ctrl := getControl('TCODECLIENT') ; if ctrl<>nil then ctrl.parent := TWinControl(ctrlRef) ;
    ctrl := getControl('ANN_TIERS') ; if ctrl<>nil then ctrl.parent := TWinControl(ctrlRef) ;
  end ;
  //-LM20070125

  //+LM20070511 GereOnChange ;
  //contourne : nom de champ incompatible Web Access ( donnée non affichée )
  EvDecla.Rebranche('ANN_MOISCLOTUREPP', 'OnExit', handlerChangeField) ;
  EvDecla.Rebranche('ANN_SIRENPP', 'OnExit', handlerChangeField) ;
  EvDecla.Rebranche('ANN_ENSEIGNEPP', 'OnExit', handlerChangeField) ;
  EvDecla.Rebranche('ANN_CLESIRETPP', 'OnExit', handlerChangeField) ;
  EvDecla.Rebranche('ANN_RCSVILLEPP', 'OnExit', handlerChangeField) ;
  EvDecla.Rebranche('ANN_RCSDATEPP', 'OnExit', handlerChangeField) ;
  EvDecla.Rebranche('ANN_RCSGESTPP', 'OnExit', handlerChangeField) ;
  EvDecla.Rebranche('ANN_RMDEPPP', 'OnExit', handlerChangeField) ;
  EvDecla.Rebranche('ANN_RMANNEEPP', 'OnExit', handlerChangeField) ;
  EvDecla.Rebranche('ANN_FORMEPP', 'OnExit', handlerChangeField) ;
  EvDecla.Rebranche('ANN_CODEINSEEPP', 'OnExit', handlerChangeField) ;
  EvDecla.Rebranche('ANN_RCSPP', 'OnExit', handlerChangeField) ;
  EvDecla.Rebranche('ANN_RMPP', 'OnExit', handlerChangeField) ;
  //ANN_CODENAFPP: nom de champ incompatible Web Access ( donnée non affichée )
  //ANN_CODENAF2PP: nom de champ incompatible Web Access ( donnée non affichée )
  //-LM20070511


  bDuringLoad := false ; //LM20070213 il n'y était pas. Normal?
  //MD20070405 Oui, car on comptait rester en bDuringLoad jusqu'à la fin du onloadrecord
  //mais ça n'a pas d'incidence, le onloadrecord le repasse à true,
  //et le vrai pb est qu'on passe dans OnChangeField même APRES le OnLoadRecord...
end;


// $$$ JP 13/08/07
{$IFDEF BUREAU}
procedure DPAGLFonctionCTI (parms:array of variant; nb:integer);
var
   strTelephone   :string;
begin
  if string (Parms [0]) = 'MAKECALL' then
  begin
       strTelephone := Trim (string (Parms [1]));
       if strTelephone <> '' then
          if PgiAsk ('Appeler le ' + strTelephone + ' ?') = mrYes then
             VH_DP.ctiAlerte.MakeCall (strTelephone);
  end;
end;
{$ENDIF BUREAU}

{*****************************************************************
Auteur ....... : BM
Date ......... : 25/10/02
Fonction ..... : ElipsisPerAssClick
Description .. : Affectation d'une personne
Paramètres ... : L'objet
*****************************************************************}

procedure TOM_Annuaire.ElipsisPerAssClick(Sender: TObject);
//LM20070516 var sChampCode_l, sChampNom_l, sValeurCode_l, sValeurNom_l: string;
begin
  if sender=nil then exit ;
  cElipsisPerAssClick(Self, Sender) ;

  (*LM20070516
  ModeEdition(DS);

  sChampNom_l := (Sender as THEdit).Name;
  sChampCode_l := Copy(sChampNom_l, 2, Length(sChampNom_l));
  sValeurNom_l := GetControlText(sChampNom_l);

  ChoisirPersonne(sValeurCode_l, sChampCode_l, sChampNom_l);
  *)
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 04/11/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
(*LM
procedure TOM_Annuaire.PerAss_OnExit(Sender: TObject);
var
  sChampCode_l, sChampNom_l, sValeurCode_l, sValeurNom_l, sValeurNomSov_l: string;
begin
  if not bDuringLoad then exit;

  ModeEdition(DS);

  sChampNom_l := (Sender as THEdit).Name;
  sChampCode_l := Copy(sChampNom_l, 2, Length(sChampNom_l));
  sValeurCode_l := GetField(sChampCode_l);
  sValeurNom_l := GetControlText(sChampNom_l);
  sValeurNomSov_l := GetControlText(sChampNom_l + 'SOV');

  if sValeurNom_l = '' then
    SetField(sChampCode_l, '')
  else if sValeurNom_l <> sValeurNomSov_l then
    ChoisirPersonne(sValeurCode_l, sChampCode_l, sChampNom_l);
end;
*)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 18/01/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_Annuaire.BTSUP_OnClick(Sender: TObject);
var
    sNom_l : string;
begin
  ModeEdition(DS);

  sNom_l := (Sender as TToolbarButton97).Name;
  sNom_l := Copy(sNom_l, 6, Length(sNom_l));

  SetField('ANN_' + sNom_l, '');//LM20070315
  SetControlText('TANN_' + sNom_l, '');
  SetControlText('TANN_' + sNom_l + 'SOV', '');

  if Copy(sNom_l, 1, 6) = 'PERASS' then
  begin
    sNom_l := Copy(sNom_l, 1, 7) + 'QUAL';
    SetField('ANN_' + sNom_l, '');
  end;
end;

{LM20070516
procedure TOM_Annuaire.ChoisirPersonne(sGuidPer_p, sZoneGuidPer_p, sZoneLibelle_p: string);
var
  sValeurCode_l, sValeurNom_l, sSansDossier_l : string;
begin
  sValeurNom_l := GetControlText(sZoneLibelle_p);
  if Appli = 'JUR' then
     sSansDossier_l := 'X'
  else
     sSansDossier_l := '-';

//  sValeurCode_l := AGLLanceFiche('YY', 'ANNUAIRE_SEL', 'ANN_NOMPER=' + sValeurNom_l, '', ';;' + sSansDossier_l);
  sValeurCode_l := LancerAnnuSel ('ANN_NOMPER=' + sValeurNom_l, '', ';;' + sSansDossier_l);

  if (sValeurCode_l <> '') and (sValeurCode_l <> GetControlText(sZoneGuidPer_p)) then
  begin
    ModeEdition(DS);
    SetField(sZoneGuidPer_p, sValeurCode_l);
  end;
end;
}

procedure TOM_Annuaire.OnClose;
begin
  inherited;

  // Rétablissement après la synchro éventuelle des paramsoc
  // (rq : non adapté si pas monofiche, car dans ce cas plusieurs NoDossier parcourus)
{$IFDEF DP}
  // $$$ JP 03/01/06 - utile pour les autres que DP???
  if (NoDossier <> '') and DBExists ('DB' + NoDossier) then
     DBSetAutoClose ('DB' + NoDossier, TRUE); 
     // (ne pouvait être fait au moment de la synchro, car
     // "ALTER DATABASE statement not allowed within multi-statement transaction")
{$ENDIF}

  if TOBAnnuBis <> nil then
    TOBAnnuBis.Free;
  if TOBBlob <> nil then
  begin
    TOBBlob.ClearDetail;
    TOBBlob.Free;
  end;

  // MD20070705 libérés automatiquement par le owner (car ce sont des TComponent)
  // DPP_.Free;//LM20070213
  // EvDecla.Free ;//LM20070213
  TFFiche(Ecran).Retour := TFFiche(Ecran).Retour + sCodeDoublon_c;
end;


procedure TOM_Annuaire.SauveAnnuBis;
begin
  TOBAnnuBis.PutValue('ANB_GUIDPER', GetField('ANN_GUIDPER'));
  TOBAnnuBis.GetEcran(Ecran);
  TOBAnnuBis.InsertOrUpdateDB(true);
end;


procedure TOM_Annuaire.ChargeAnnuBis;
begin
  TOBAnnuBis.ClearDetail;
  TOBAnnuBis.SelectDB('"' + GetField('ANN_GUIDPER') + '"', nil);
  TOBAnnuBis.AddChampSup('MODIF', TRUE);
  TOBAnnuBis.PutValue('MODIF', FALSE);
  TOBAnnuBis.PutEcran(Ecran);
  // ChangeAnnubis permet de détecter qu'il y a une modif dans ANNUBIS
  // même si on ne modifie aucune zones de la table ANNUAIRE de la tom
  if GetControl('ANB_GESTIONFISC') <> nil then TCheckbox(GetControl('ANB_GESTIONFISC')).OnClick := ChangeAnnubis;
  if GetControl('ANB_GESTIONPATRIM') <> nil then TCheckbox(GetControl('ANB_GESTIONPATRIM')).OnClick := ChangeAnnubis;
  if GetControl('ANB_COTISETNS') <> nil then TCheckbox(GetControl('ANB_COTISETNS')).OnClick := ChangeAnnubis;
  THEdit(GetControl('ANB_DATELIBRE1')).OnChange := ChangeAnnubis;
  THEdit(GetControl('ANB_DATELIBRE2')).OnChange := ChangeAnnubis;
  THEdit(GetControl('ANB_DATELIBRE3')).OnChange := ChangeAnnubis;
  THEdit(GetControl('ANB_CHARLIBRE1')).OnChange := ChangeAnnubis;
  THEdit(GetControl('ANB_CHARLIBRE2')).OnChange := ChangeAnnubis;
  THEdit(GetControl('ANB_CHARLIBRE3')).OnChange := ChangeAnnubis;
  THValComboBox(GetControl('ANB_CHOIXLIBRE1')).OnChange := ChangeAnnubis;
  THValComboBox(GetControl('ANB_CHOIXLIBRE2')).OnChange := ChangeAnnubis;
  THValComboBox(GetControl('ANB_CHOIXLIBRE3')).OnChange := ChangeAnnubis;
  TCheckbox(GetControl('ANB_BOOLLIBRE1')).OnClick := ChangeAnnubis;
  TCheckbox(GetControl('ANB_BOOLLIBRE2')).OnClick := ChangeAnnubis;
  TCheckbox(GetControl('ANB_BOOLLIBRE3')).OnClick := ChangeAnnubis;
  THNumEdit(GetControl('ANB_VALLIBRE1')).OnChange := ChangeAnnubis;
  THNumEdit(GetControl('ANB_VALLIBRE2')).OnChange := ChangeAnnubis;
  THNumEdit(GetControl('ANB_VALLIBRE3')).OnChange := ChangeAnnubis;
  //LM20070516 THValComboBox(GetControl('ANB_DOUBLON')).OnChange := ChangeAnnubis;
  //mcd 06/06/2005 ajout champa dresse, plus ci-dessus, fait dans tous les cas
  // BM 08/06/2005 : contrôle de l'existence des champs sur la fiche sinon scratch!!!!
  if GetControl('ANB_OLDRUE1') <> nil then
    THEdit(GetControl('ANB_OLDRUE1')).OnChange := ChangeAnnubis;

  if GetControl('ANB_OLDRUE2') <> nil then
    THEdit(GetControl('ANB_OLDRUE2')).OnChange := ChangeAnnubis;

  if GetControl('ANB_OLDRUE3') <> nil then
    THEdit(GetControl('ANB_OLDRUE3')).OnChange := ChangeAnnubis;

  if GetControl('ANB_OLDVILLE') <> nil then
    THEdit(GetControl('ANB_OLDVILLE')).OnChange := ChangeAnnubis;

  if GetControl('ANB_OLDCP') <> nil then
    THEdit(GetControl('ANB_OLDCP')).OnChange := ChangeAnnubis;
  //CM 15/06/2005 : ajout des nouveaux champs
  if GetControl('ANB_OLDVOIENO') <> nil then
    THEdit(GetControl('ANB_OLDVOIENO')).OnChange := ChangeAnnubis;

  if GetControl('ANB_OLDVOIENOCOMPL') <> nil then
    THEdit(GetControl('ANB_OLDVOIENOCOMPL')).OnChange := ChangeAnnubis;

  if GetControl('ANB_OLDVOIETYPE') <> nil then
    THEdit(GetControl('ANB_OLDVOIETYPE')).OnChange := ChangeAnnubis;

  if GetControl('ANB_OLDVOIENOM') <> nil then
    THEdit(GetControl('ANB_OLDVOIENOM')).OnChange := ChangeAnnubis;
end;


procedure TOM_Annuaire.ChangeAnnubis(Sender: TObject);
// pour dire qu'on a changé une valeur dans ANNUBIS, qui n'est
// pas la table principale de la tom
begin
  // on ne fait rien pendant le OnLoadRecord
  if bDuringLoad then exit;

  TOBAnnuBis.PutValue('MODIF', TRUE);
  ModeEdition(DS);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 18/12/2003
Modifié le ... : 18/12/2003
Description .. : pour dire qu'on a changé une valeur dans ANN_COMPLT,
Suite ........ : qui n'est pas un champ de la tom
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TOM_Annuaire.ChangeCOMPLT(Sender: TObject);
begin
  // on ne fait rien pendant le OnLoadRecord
  if bDuringLoad then exit;

  if GetField('ANN_COMPLTNOADMIN') = GetControlText('ANN_COMPLT') then exit;
  ModeEdition(DS);

  SetField('ANN_COMPLTNOADMIN', GetControlText('ANN_COMPLT'));
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 12/05/2006
Modifié le ... :
Description .. : pour dire qu'on a changé une valeur dans ANN_NOADM,
Suite ........ : qui n'est pas un champ de la tom
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TOM_Annuaire.ChangeNOADM(Sender: TObject);
begin
  // on ne fait rien pendant le OnLoadRecord
  if bDuringLoad then exit;

  if GetField('ANN_NOADMIN') = GetControlText('ANN_NOADM') then exit;
  ModeEdition(DS);

  SetField('ANN_NOADMIN', GetControlText('ANN_NOADM'));
end;

procedure TOM_Annuaire.BContactClick(Sender: TObject);
var sttmp,stTmp1, stTmp2 ,auxi: string;
begin    //mcd 12/2005
{$ifndef Annuaire_seul}
  stTmp2:=';TITRE='+GetField('ANN_NOM1')+';ACTION=MODIFICATION';
  stTmp2:= StTmp2+';ANNUAIRE=TRUE'; //mcd 11677 Bureau.
  if TFFiche(Ecran).fTypeAction = taConsult  then   stTmp2:=';ACTION=CONSULTATION';
  //mcd 16/06/06 pour le cas ou le tiers n'existe pas dans la base
  // il ne faut pas ramener tous les contacts (en passant sttmp1=T et blanc)
  if GetField('ANN_TIERS') <> '' then
    begin
    Auxi := TiersAuxiliaire (GetField('ANN_TIERS'),False); //mcd 16/06/06
    if (Auxi <> '')
    then
      begin
      stTmp := 'TYPE=T;TIERS='+ GetField('ANN_TIERS')+';TYPE2=' +GetField('ANN_NATUREAUXI');
      stTmp1 := 'T;'+ Auxi ;
      end
{$IFDEF DP}
    else
    // MBrun 20/12/2007 - KPMG
    // si le tiers est renseigné mais qu'on ne trouve pas l'auxiliaire (table tiers vide),
    // il faut quand même passer une clé auxiliaire (= au code tiers chez kpmg)
    if VH_DP.SeriaKPMG then
      begin
      stTmp := 'TYPE=T;TIERS='+ GetField('ANN_TIERS')+';TYPE2=' +GetField('ANN_NATUREAUXI');
      stTmp1 := 'T;'+ GetField('ANN_TIERS') ;
      end
{$ENDIF}
    else
      begin
      stTmp := 'TYPE=ANN';
      stTmp1:='ANN;'+ GetField('ANN_GUIDPER');   //mcd 16/06/06 pour ne pas avoir tous les tiers,
      end;                                       //si onpassse T + blanc. essaie avec ann et guidper, mais il y a de grande chance que cela n'existe pas
    end
  else
    begin
    stTmp := 'TYPE=ANN';
    stTmp1 := 'ANN;'+ GetField('ANN_GUIDPER');
    end;
  AglLanceFiche('YY','YYCONTACT',stTmp1,'',stTmp+';GUIDPER='+GetField('ANN_GUIDPER')+';ALLCONTACT'+stTmp2) ;
{$endif}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 29/01/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_Annuaire.ChangeBlob(Sender: TObject);
var
  T: TOB;
  GuidPer, RangBlob: String;
  leTexte, laZone: string;
begin
  // on ne fait rien pendant le OnLoadRecord
  if bDuringLoad then exit;
  bChangeBlob := True;

  GuidPer := GetField('ANN_GUIDPER');
  laZone := (Sender as THRicheditOle).Name;
  if instring(laZone , ['OLE_ACTIVITE', 'OLE_ACTIVITEPM']) then//LM20070315
    RangBlob   := '1'
  else if instring(laZone , ['OLE_ACTIVITE2', 'OLE_ACTIVITE2PM']) then 
    RangBlob   := '5'
  else if laZone = 'OLE_BLOCNOTES' then
    RangBlob   := '2'
  else if laZone = 'OLE_REGMATTXT' then
    RangBlob   := '3'
  else if laZone = 'OLE_OBJETSOC' then
    RangBlob   := '4'
  else
    exit;    // SORTIE !

  T := TOBBlob.FindFirst(['LO_TABLEBLOB', 'LO_IDENTIFIANT', 'LO_RANGBLOB'], ['ANN', GuidPer, RangBlob], true);
  if T = nil then exit;

  //  leTexte := GetControlText(laZone);
  leTexte := RichToString(Sender as THRicheditOle); //EPZ 17/05/01
  T.PutValue('LO_OBJET', leTexte);

  //+LM20070315
  bDuringLoad :=true ;
  if laZone = 'OLE_ACTIVITE' then
    setControlText('OLE_ACTIVITEPM', getControlText('OLE_ACTIVITE'))
  else if laZone = 'OLE_ACTIVITEPM' then
    setControlText('OLE_ACTIVITE', getControlText('OLE_ACTIVITEPM')) ;

  if laZone = 'OLE_ACTIVITE2' then
    setControlText('OLE_ACTIVITE2PM', getControlText('OLE_ACTIVITE2'))
  else if laZone = 'OLE_ACTIVITE2PM' then
    setControlText('OLE_ACTIVITE2', getControlText('OLE_ACTIVITE2PM')) ;

  bDuringLoad :=false;
  //-LM20070315

  ForceModidier ;//LM20070315
end;

//EPZ 25/10/00

procedure TOM_Annuaire.ChargeCaption;
// affiche les libellés des zones libres (onglet Divers de la fiche annuaire)
// rq : table ANNUPARAM aurait pu être des paramsoc car un seul enreg
var TOBCaption: TOB;
begin
  TOBCaption := TOB.Create('ANNUPARAM', nil, -1);
  TOBCaption.SelectDB('', nil);
  // renseigne les labels ANP_DATELIBREx, ANP_MONTANTLIBREx, ANP_TEXTELIBREx
  TOBCaption.PutEcran(Ecran);
  SetControlText('ANP_TABLELIBRE1', RechDom('YYANNUPARAM', 'AN1', false));
  SetControlText('ANP_TABLELIBRE2', RechDom('YYANNUPARAM', 'AN2', false));
  SetControlText('ANP_TABLELIBRE3', RechDom('YYANNUPARAM', 'AN3', false));
  SetControlCaption('ANB_BOOLLIBRE1', TOBCaption.GetValue('ANP_COCHELIBRE1'));
  SetControlCaption('ANB_BOOLLIBRE2', TOBCaption.GetValue('ANP_COCHELIBRE2'));
  SetControlCaption('ANB_BOOLLIBRE3', TOBCaption.GetValue('ANP_COCHELIBRE3'));
  // cache les zones libres non utilisées
  SetControlVisible('ANB_CHOIXLIBRE1', GetControlText('ANP_TABLELIBRE1') <> '');
  SetControlVisible('ANB_CHOIXLIBRE2', GetControlText('ANP_TABLELIBRE2') <> '');
  SetControlVisible('ANB_CHOIXLIBRE3', GetControlText('ANP_TABLELIBRE3') <> '');
  SetControlVisible('ANB_DATELIBRE1', TOBCaption.GetValue('ANP_DATELIBRE1') <> '');
  SetControlVisible('ANB_DATELIBRE2', TOBCaption.GetValue('ANP_DATELIBRE2') <> '');
  SetControlVisible('ANB_DATELIBRE3', TOBCaption.GetValue('ANP_DATELIBRE3') <> '');
  SetControlVisible('ANB_CHARLIBRE1', TOBCaption.GetValue('ANP_TEXTELIBRE1') <> '');
  SetControlVisible('ANB_CHARLIBRE2', TOBCaption.GetValue('ANP_TEXTELIBRE2') <> '');
  SetControlVisible('ANB_CHARLIBRE3', TOBCaption.GetValue('ANP_TEXTELIBRE3') <> '');
  SetControlVisible('ANB_BOOLLIBRE1', TOBCaption.GetValue('ANP_COCHELIBRE1') <> '');
  SetControlVisible('ANB_BOOLLIBRE2', TOBCaption.GetValue('ANP_COCHELIBRE2') <> '');
  SetControlVisible('ANB_BOOLLIBRE3', TOBCaption.GetValue('ANP_COCHELIBRE3') <> '');
  SetControlVisible('ANB_VALLIBRE1', TOBCaption.GetValue('ANP_MONTANTLIBRE1') <> '');
  SetControlVisible('ANB_VALLIBRE2', TOBCaption.GetValue('ANP_MONTANTLIBRE2') <> '');
  SetControlVisible('ANB_VALLIBRE3', TOBCaption.GetValue('ANP_MONTANTLIBRE3') <> '');
  TOBCaption.Free;
end;


procedure TOM_Annuaire.BTIERS_OnClick(Sender: TObject);
begin
  // MCD-MD tout refait pour lien DP-GI
  if GetField('ANN_TIERS') <> '' then
    // MD 22/11/00 - AGLLanceFiche('AFF', 'AFTIERS', 'T_TIERS='+GetField('ANN_TIERS'),'','T_NATUREAUXI=CLI;ORIGINE=DP;MONOFICHE')
    V_PGI.DispatchTT(8, taConsult, GetField('ANN_TIERS'), '', 'ORIGINE=DP;MONOFICHE')
  else
    PGIInfo('Si vous voulez un tiers associé, utilisez la gestion interne pour le créer.', TitreHalley);
end;


procedure TOM_Annuaire.BLIENS_OnClick(Sender: TObject);
begin
  // Dans le script de la fiche, il y a :
  //  ouvrefiche('JUR','ANNULIEN_LST','ANL_GUIDPER='+GetChamp('ANN_GUIDPER'),'',';'+GetChamp('ANN_NOMPER'))
  AglLanceFiche('DP', 'LIENPERSONNE', 'ANL_GUIDPER=' + GetField('ANN_GUIDPER'),'', GetField('ANN_GUIDPER') + ';DOS')
end;


procedure TOM_Annuaire.BINTERVENANTS_OnClick(Sender: TObject);
begin
  AglLanceFiche('DP','LIENINTERVENANT','ANL_GUIDPERDOS='+GetField('ANN_GUIDPER')+';ANL_FONCTION=DIV',
    '', GetField('ANN_GUIDPER')+';DIV') ;
end;


procedure TOM_Annuaire.BEVENEMENT_OnClick(Sender: TObject);
var
    GuidPerdos, NumDos: string;
    Q1: TQuery;
begin
  GuidPerdos := GetField('ANN_GUIDPER');

  NumDos := '';
  Q1 := OpenSql('SELECT DOS_NODOSSIER FROM DOSSIER WHERE DOS_GUIDPER="' + GuidPerdos + '"', True);
  if not Q1.Eof then
    NumDos := Q1.FindField('DOS_NODOSSIER').Asstring;
  Ferme(Q1);

  AglLanceFiche('JUR', 'EVENEMENT_MUL', 'JEV_GUIDPER='+GuidPerDos, '', 'DOS');
end;


procedure TOM_Annuaire.OnAfterUpdateRecord;
begin
  inherited;
  majcontact; //mcd 08/03/2006 10835
  // MD-MCD
  SynchroniseParamsoc(sGuidPer_c);

{$IFNDEF Annuaire_seul}
// dwi ceic 070208
  if GetField('ANN_TIERS') <> '' then
    SynchroniseTiers(TRUE, sGuidPer_c, GetField('ANN_TIERS'));
{$ENDIF Annuaire_seul}

(* mcd 16/11/2006 11150 on ne crée plus le tiers en automatique, même si GRC
{$IFDEF BUREAU}
  if CtxGrc in V_PGI.PgiCOntexte then
  begin
    if (GetField('ANN_TIERS') = '') and (GetField('ANN_NATUREAUXI') <> '') then
    begin
      DpCreerTiers(GuidPer, GetControlText('ANN_NATUREAUXI'), GetCOntrolText('ANN_NOM1'), GetCOntrolText('ANN_NOMPER'));
    end;
  end; // fin GRC
{$ENDIF}  fin mcd 16/11/2006*)

  // MD : pour quand on valide un nouvel enreg
  if GetControl('BCONTACTS') <> nil then
    SetControlEnabled('BCONTACTS', True);
  if bVoirLiens and (GetControl('BLIENS') <> nil) then
    SetControlEnabled('BLIENS', True);
  if bVoirLiens and (GetControl('BINTERVENANTS') <> nil) then
    SetControlEnabled('BINTERVENANTS', True);
  if GetControl('BEVENEMENT') <> nil then
    SetControlEnabled('BEVENEMENT', True);
  if GetControl('BDOCEVEN') <> nil then
    SetControlEnabled('BDOCEVEN', True);

  if GetControl('BPERSONNE') <> nil then
    SetControlEnabled('BPERSONNE', True);
end;


procedure TOM_Annuaire.BIMPRIMER_OnClick (Sender : Tobject);
var GuidPer, ChSql : String;
begin
 GuidPer:=GetField ('ANN_GUIDPER');

 ChSql:='SELECT * FROM ANNUAIRE '+
        'LEFT JOIN DOSSIER ON DOS_GUIDPER=ANN_GUIDPER '+
        'LEFT JOIN ANNUBIS ON ANB_GUIDPER=ANN_GUIDPER '+
        'LEFT JOIN DPPERSO ON DPP_GUIDPER=ANN_GUIDPER '+
        'LEFT JOIN ANNUPARAM ON ANP_CODE="-" OR ANP_CODE="X" '+
        'LEFT JOIN TIERSCOMPL ON YTC_TIERS=ANN_TIERS '+
        'WHERE ANN_GUIDPER="'+GuidPer+'"';

 LanceEtat ('E','DPG','AN3',True,False,False,nil,ChSql,'Annuaire',False);
end;

procedure TOM_Annuaire.BDELETE_OnClick(Sender: TObject);
var GuidPer, Texte: string;
begin
  GuidPer := GetField('ANN_GUIDPER');
  // suppression personne, avec vérifications
  if not TestSupprimePersonne(GuidPer) then
    exit;

  Texte := 'Vous allez supprimer définitivement cette personne.#13#10Confirmez vous l''opération ?';

  if PGIAsk(Texte, TitreHalley) <> mrYes then
    exit; // SORTIE

  // suppression DP
  if SupprimeInfoDp(GuidPer) then
  begin
    // suppression de l'enreg en cours dans la fiche (vu qu'on a redirigé le bdelete)
    SupprimeResteAnnuaire(GuidPer);
//    RefreshDB;
      Ecran.Close;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 10/06/2003
Modifié le ... :
Description .. : utilisée dans OnChangeField (donc inutile dans
Suite ........ : OnLoadRecord)
Mots clefs ... :
*****************************************************************}
procedure TOM_Annuaire.AfficheOngletPPouPMouORG;
var
  sAnnFamPer_l, sAnnPpPm_l  :string;
  sAnnSexe,s          :string; // $$$ JP 24/04/06
  //LM bOrganismeVisible_l       :boolean;
begin
  sAnnPpPm_l   := GetField ('ANN_PPPM');
  sAnnFamPer_l := GetField ('ANN_FAMPER');
  sAnnSexe     := GetField ('ANN_SEXE');

  if (sAnnPpPm_l = 'PM') then
    begin
    SetControlText('TANN_NOM1','Dénomination');
    TTabSheet(GetControl('TABADRESSEFISC')).Caption := 'Principal établissement'; // FQ 11568
    SetControlProperty('BCOPIEADRSFISCALE', 'Hint', 'Cliquer pour créer l''adresse courrier à partir de l''adresse du principal établissement');
    SetControlCaption('EXPLICATIONADRLEGALE', 'Il s''agit de l''adresse du siège social.'); // FQ 11568
    end
  else
    begin
    SetControlText('TANN_NOM1','Nom') ; // FQ 11688
    SetControlText('TANN_NOM3','Nom de jeune fille'); // FQ 11688
    TTabSheet(GetControl('TABADRESSEFISC')).Caption := 'Adresse professionnelle'; // FQ 11568
    SetControlProperty('BCOPIEADRSFISCALE', 'Hint', 'Cliquer pour créer l''adresse courrier à partir de l''adresse professionnelle');
    SetControlCaption('EXPLICATIONADRLEGALE', 'Il s''agit de l''adresse du domicile de l''exploitant.'); // FQ 11568
    end;

  if (sAnnPpPm_l = 'PM') then
    s := 'En cas seulement d''utilisation de cette adresse, cette adresse sera reprise dans le cadre "Identification du destinataire" de la déclaration fiscale. '+
         '"L''adresse légale" sera alors reprise dans le cadre "Adresse du siège social si elle est différente du principal établissement".'
  else
    s := 'En cas seulement d''utilisation de cette adresse, cette adresse sera reprise dans le cadre "Identification du destinataire" de la déclaration fiscale. '+
         '"L''adresse légale" sera alors reprise dans le cadre "Adresse du domicile de l''exploitant" (situé à gauche du précédent).' ;
  setControlText('ExplicationAdrFiscale', s) ;

  SetControlVisible('ANN_SEXE', sAnnPpPm_l = 'PP');
  SetControlVisible('TANN_SEXE', sAnnPpPm_l = 'PP');
  // Onglet Personne morale
  SetControlVisible('PM', (sAnnPpPm_l = 'PM')
    //LM20070620 marche pas and (not AnnuaireLanceeDepuisSynthese)  //LM20070611 on masque l'onglet PM hors annuaire central
    //FQ11232 and (not InString(sAnnFamPer_l, ['FIS', 'OGA']))
    and (not InString(sAnnFamPer_l, ['FIS']))
    );
  // Onglet Personne physique
  SetControlVisible('PP', sAnnPpPm_l = 'PP');

  SetControlVisible('TANN_PPPM', not (sAnnFamPer_l='FIS') ); //LM fq11232 SetControlVisible('ANN_PPPM', not InString(sAnnFamPer_l, ['FIS', 'OGA']));
  SetControlVisible('TANN_PPPM', not (sAnnFamPer_l ='FIS') ); //LM fq11232  SetControlVisible('TANN_PPPM', not InString(sAnnFamPer_l, ['FIS', 'OGA']) );

  SetControlVisible('TANN_NOM2', (sAnnPpPm_l = 'PP'));
  SetControlVisible('ANN_NOM3',  (sAnnPpPm_l = 'PP') and (sAnnSexe='F'));
  SetControlVisible('TANN_NOM3', (sAnnPpPm_l = 'PP') and (sAnnSexe='F'));
  SetControlVisible('ANN_NOM4',  (sAnnPpPm_l = 'PP'));
  SetControlVisible('TANN_NOM4', (sAnnPpPm_l = 'PP'));

  SetControlVisible('GBCOURRIER', true); //LM20070502 //(sAnnPpPm_l = 'PP')); //LM20070125
  SetControlVisible('ANB_GESTIONFISC', (sAnnPpPm_l = 'PP'));
  SetControlVisible('ANB_GESTIONPATRIM', (sAnnPpPm_l = 'PP'));
  SetControlVisible('ANB_COTISETNS', (sAnnPpPm_l = 'PP'));
//  SetControlVisible('PFAMTYPE', (nodossier = '') or v_pgi.SAV);  //-LMO20061005
  SetControlVisible('PFAMTYPE', not (AnnuaireLanceeDepuisSynthese) or v_pgi.SAV) ; // CAT 20071207 FQ 11886

  ANN_FAMPER_OnClick(nil);//LM20070502 SetControlProperty('TABADRESSEFISC','TabVisible', not InString(sAnnFamPer_l, ['FIS', 'OGA']) );  //-LM20070213

  if sAnnPpPm_l = 'PM' then
  begin
    FiltreCivilite('');
    // MD 12/12/06 - Pas de comportement spécifique Juri
    //    bOrganismeVisible_l := (Appli <> 'JUR');
    //    bOrganismeVisible_l := (Application.ExeName <> 'CJS5.EXE');

    if {bOrganismeVisible_l and } ((sAnnFamPer_l = 'FIS') or (sAnnFamPer_l = 'OGA')) then
         {or(sAnnFamPer_l = 'SOC'))}
    begin
      if sAnnFamPer_l = 'OGA' then
      begin
        SetControlCaption('GRPBOXADMIN', 'Identification du cabinet auprès de l''OGA');
        SetControlCaption('GBNOADMIN', 'N° d''agrément du CGA');
        SetControlProperty('ANN_NOADM', 'MaxLength', 6);
        // EditMask inutilisables en Web Access (pas du tout même comportement qu'en 2 tiers)
        // SetControlProperty('ANN_NOADM', 'EditMask', '! |  a  |  a  |  a  |  a  |  a  |  a  |;0;_');
        SetControlCaption('GBCOMPLTNOADMIN', 'Identif. cabinet');
        // SetControlProperty('ANN_NOIDENTIF', 'EditMask', '! | a | a |;0;_');
        SetControlProperty('ANN_NOIDENTIF', 'MaxLength', 2);
        SetControlProperty('ANN_COMPLT', 'MaxLength', 12);
        // SetControlProperty('ANN_COMPLT', 'EditMask', '! |  a  |  a  |  a  |  a  |  a  |  a  |  a  |  a  |  a  |  a  |  a  |  a  |;0;_');
      end
      else
      begin
        SetControlCaption('GRPBOXADMIN', 'Organisme Fiscal');
        SetControlCaption('GBNOADMIN', 'Recette / CDIR');
        SetControlProperty('ANN_NOADM', 'MaxLength', 7);
        // SetControlProperty('ANN_NOADM', 'EditMask', '! |  a  |  a  |  a  |  a  |  a  |  a  |  a  |;0;_');
        SetControlCaption('GBCOMPLTNOADMIN', 'Inspection / IFU');
        // SetControlProperty('ANN_NOIDENTIF', 'EditMask', '! | a | a |;0;_');
        SetControlProperty('ANN_NOIDENTIF', 'MaxLength', 2);
        SetControlProperty('ANN_COMPLT', 'MaxLength', 3);
        // SetControlProperty('ANN_COMPLT', 'EditMask', '! |  a  |  a  |  a  |;0;_');
        SetControlText('ANN_COMPLT', Copy(GetControlText('ANN_COMPLT'), 1, 3));
      end;
    end;
    setControlVisible('GBNOIDENTIF',(sAnnFamPer_l = 'FIS'));//LM20070611

  end
  else if sAnnPpPm_l = 'PP' then
  begin
    FiltreCivilite (sAnnSexe); // $$$ JP 24/04/06 GetField('ANN_SEXE'));
  end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 02/08/02
Fonction ..... : CoherenceFormeJuridique
Description .. : Vérifie la cohérence de la forme juridique entre le dossier et l'annuaire
Paramètres ... : Code du dossier
  Forme juridique de l'annuaire
  Forme juridique du dossier (à déterminer)
Renvoie ...... : true si cohérent, false sinon
*****************************************************************}

function TOM_Annuaire.CoherenceFormeJuridique(sGuidPerDos_p: string; sAnnForme_p: string; var sJurForme_p: string): boolean;
var
  qQuery_l: TQuery;
begin
  qQuery_l := OpenSQL('select JUR_FORME from JURIDIQUE where JUR_GUIDPERDOS = "' + sGuidPerDos_p + '" and JUR_TYPEDOS = "STE"', true);
  if not qQuery_l.EOF then
    sJurForme_p := qQuery_l.Fields[0].AsString;
  Ferme(qQuery_l);

  if sJurForme_p <> sAnnForme_p then
    result := true
  else
    result := false;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 24/07/02
Fonction ..... : ChangementFormeJuridique
Description .. : Traitement effectué en cas de changement
Paramètres ... : Ancienne forme
  Nouvelle forme
Renvoie ...... : true si le changement s'est bien passé, false sinon
*****************************************************************}

function TOM_Annuaire.ChangementFormeJuridique(sAncForme_p, sNouForme_p: string): boolean;
var sOk_l: string;
begin
  result := false;
  if EnregJurExiste(GetField('ANN_GUIDPER')) then
  begin
    if PGIAsk('Attention : la forme juridique de la société a changé.#13#10' +
      'Or cette société fait l''objet d''un dossier "juridique".#13#10' +
      'Voulez-vous modifier immédiatement ce dossier "juridique" et#13#10' +
      'les fonctions des personnes éventuellement rattachées?',
      Ecran.Caption) = mrYes then
    begin
      if PGIAsk('Attention : le changement de forme juridique du dossier#13#10' +
        'va s''accompagner d''une adaptation des liens des personnes#13#10' +
        'rattachées (propriétaires de titres, dirigeants, etc...).#13#10' +
        'Ce traitement ne doit être réalisé que lorsque toutes les étapes#13#10' +
        'de la transformation auront été réalisées.#13#10' +
        'Voulez-vous réellement modifier la forme juridique de cette société#13#10' +
        'et les fonctions des personnes qui lui sont rattachées?',
        Ecran.Caption) = mrYes then
      begin
        sOk_l := AGLLanceFiche('JUR', 'TRANSFORM', '', '', GetField('ANN_GUIDPER') + ';' + GetField('ANN_NOMPER') + ';' + sAncForme_p + ';' + sNouForme_p);
        if sOk_l = '1' then
        begin
          PGIInfo('La forme juridique du dossier ' + GetField('ANN_NOMPER')
            + ' a été modifiée,#13#10' +
            'ainsi que les liens des personnes rattachées.#13#10' +
            'Il conviendra cependant de vérifier dans le dossier les informations#13#10' +
            'concernant les propriétaires de titres ainsi que les mandataires, pour#13#10' +
            's''assurer que ces informations correspondent à la nouvelle forme #13#10' +
            'juridique (dates de nomination et d''expiration...)',
            Ecran.Caption);
          result := true;
        end;
      end;
    end;
  end
  else
  begin
    result := true;
  end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 25/07/02
Fonction ..... : EnregJurExiste
Description .. : Vérifie l'existence d'un enregsitrement dans la table JURIDIQUE
  et / ou d'enregistrements dans la table ANNULIEN
Paramètres ... : Code personne
Renvoie ...... : true si existe, false sinon
*****************************************************************}

function TOM_Annuaire.EnregJurExiste(sGuidPerDos_p: string): boolean;
begin
  Result := False;
  if ExisteSQL('select JUR_GUIDPERDOS from JURIDIQUE where JUR_GUIDPERDOS = "'
    + sGuidPerDos_p + '" and JUR_TYPEDOS = "STE"')
    or ExisteSQL('select ANL_GUIDPERDOS from ANNULIEN where ANL_GUIDPERDOS = "'
    + sGuidPerDos_p + '" and ANL_TYPEDOS = "STE"') then
    Result := True;
end;

procedure TOM_Annuaire.BValiderClick(Sender: TObject);
var siret : string ;
begin
{  if (DS.State <> dsInsert) and (GetField('ANN_FORME') <> sJurForme_c) then
  begin
    DS.Edit;
  end;}

  siret := GetField('ANN_SIREN') + GetField('ANN_CLESIRET') ;
//   sCodeDoublon_c := '';
  if (DS.State = dsInsert) then //LM>MD : Pourquoi ce n'est pas fait dans le OnUpdateRecord ?
  //MD>LM : parce qu'une fois la fiche enregistrée, on ne veut plus contrôler, et surtout pas avec elle-même !
  // or ((siretIni <> siret) and (siret<>'')) then    //LM20070511
  //MD20070627 FQ11441 - test désactivé : on vérifie les doublons une seule fois à l'insertion, pas en modif
  begin
    if sCodeDoublon_c = '' then
    begin
       //sCodeDoublon_c := ListerDoublons( GetField('ANN_PPPM'), GetControlText('ANN_NOM1'), GetControlText('ANN_NOM2'), GetControlText('ANN_NOMPER'));

       //--- Vérification homonyme sur le siret
      if (GetField('ANN_SIREN') <> '') and (GetField('ANN_CLESIRET') <> '') then
        sCodeDoublon_c := VerifierDoublon('DOUBLONSIRET', GetField('ANN_SIREN'), GetField('ANN_CLESIRET'), '', '', '', '');

       //--- Vérification homonyme sur le nom de recherche
      if (sCodeDoublon_c = '') and ((GetField('ANN_NOM1') <> '') or (GetField('ANN_NOMPER') <> '')) then
        sCodeDoublon_c := VerifierDoublon('DOUBLONS', '', '', GetField('ANN_PPPM'), GetControlText('ANN_NOM1'), GetControlText('ANN_NOM2'), GetControlText('ANN_NOMPER'));
    end;

    if (sCodeDoublon_c <> '') and (sCodeDoublon_c <> '-1') then
    begin
         // On charge les données correspondantes
      TToolbarButton97(GetControl('BDEFAIRE')).OnClick(Sender);
      if TFFiche(Ecran).FMulQ <> nil then
      begin
        DS.Locate('ANN_GUIDPER', sCodeDoublon_c, []);
        sCodeDoublon_c := '';
      end
      else
      begin
        TToolbarButton97(GetControl('BFERME')).OnClick(Sender);
        sCodeDoublon_c := ';ANN_GUIDPER=' + sCodeDoublon_c;
      end;
//         exit;
    end

      // On conserve l'homonyme : sera validé au prochain click sur bouton valider
    else if sCodeDoublon_c = '' then
    begin
      sCodeDoublon_c := '-1';
    end

    else if sCodeDoublon_c = '-1' then
    begin
      TFFiche(Ecran).BValiderClick(Sender);
      sCodeDoublon_c := '';
    end;
    exit;
  end;

  TFFiche(Ecran).BValiderClick(Sender);

end;

procedure TOM_Annuaire.ANN_CHGTADR_OnClick(Sender: TObject);
var actif: Boolean;
begin
  // on ne fait rien pendant le OnLoadRecord
  if bDuringLoad then exit;

{$IFDEF EAGLCLIENT}
  actif := TCheckBox(GetControl('ANN_CHGTADR')).checked;
{$ELSE}
  actif := TDBCheckBox(GetControl('ANN_CHGTADR')).checked;
{$ENDIF}
  if not actif then
  begin
    SetControlText('ANB_OLDRUE1', '');
    SetControlText('ANB_OLDRUE2', '');
    SetControlText('ANB_OLDRUE3', '');
    SetControlText('ANB_OLDCP', '');
    SetControlText('ANB_OLDVILLE', '');
    //CM 15/06/2005 : ajout des nouveaux champs
    SetControlText('ANB_OLDVOIENO', '');
    SetControlText('ANB_OLDVOIENOCOMPL', '');
    SetControlText('ANB_OLDVOIETYPE', '');
    SetControlText('ANB_OLDVOIENOM', '');
  end;
  AfficheChgtAdr;
end;

procedure TOM_Annuaire.AfficheChgtAdr;
//var actif: Boolean;
begin
  if GetCOntrol('ANN_CHGTADR') = nil then exit;
{$IFDEF EAGLCLIENT}
  //actif := TCheckBox(GetControl('ANN_CHGTADR')).checked;
{$ELSE}
  //actif := TDBCheckBox(GetControl('ANN_CHGTADR')).checked;
{$ENDIF}

  //setControlVisible('BVOIRCHGADRESSE', actif) ; //LMO20061005

  //setControlVisible('GBCHGADRESSE', actif) ;
  (*+LMO20061005 : la suite n'est plus utile
  SetControlEnabled('ANB_OLDCP', actif);
  SetControlEnabled('ANB_OLDRUE1', actif);
  SetControlEnabled('TANB_OLDRUE1', actif);
  SetControlEnabled('ANB_OLDRUE2', actif);
  SetControlEnabled('TANB_OLDRUE2', actif);
  SetControlEnabled('ANB_OLDRUE3', actif);
  SetControlEnabled('ANB_OLDCP', actif);
  SetControlEnabled('TANB_OLDCP', actif);
  SetControlEnabled('ANB_OLDVILLE', actif);
  //CM 15/06/2005 : ajout des nouveaux champs
  SetControlEnabled('ANB_OLDVOIENO', actif);
  SetControlEnabled('ANB_OLDVOIENOCOMPL', actif);
  SetControlEnabled('ANB_OLDVOIETYPE', actif);
  SetControlEnabled('ANB_OLDVOIENOM', actif);
  -LMO20061005*)
end;

procedure TOM_Annuaire.RCSClick(Sender: TObject);
var actif: Boolean;
    prefixe : string ;
begin
  // on ne fait rien pendant le OnLoadRecord
  if bDuringLoad or (sender=nil) then exit;
  Application.ProcessMessages ;

  if copy(TControl(sender).Name , length(TControl(sender).Name)+1-2, 2)='PP' then Prefixe := 'PP' ;
  Actif := GetControlText('ANN_RCS' + prefixe)='X';

  SetControlEnabled('ANN_RCSDATE' + prefixe , actif);
  SetControlEnabled('ANN_RCSVILLE' + prefixe, actif);
  SetControlEnabled('ANN_RCSGEST'+prefixe, actif);

  if not actif then
  begin
    SetField('ANN_RCSVILLE' + prefixe, '');
    SetField('ANN_RCSDATE' + prefixe , iDate1900);
    SetControlText('TANN_RCSNO' + prefixe, '');
    SetField('ANN_RCSGEST'+prefixe, '');
  end;
end;


procedure TOM_Annuaire.RMClick(Sender: TObject);
var actif: Boolean;
  dat: TDateTime;
  annee, mois, jour: Word;
  prefixe : string ;
begin
  // on ne fait rien pendant le OnLoadRecord
  if bDuringLoad then exit;
  if copy(TControl(sender).Name , length(TControl(sender).Name)+1-2, 2)='PP' then Prefixe := 'PP' ;
  Actif := GetControlText('ANN_RM' + prefixe)='X';

  SetControlEnabled('ANN_RMDEP' + prefixe, actif);
  SetControlEnabled('ANN_RMANNEE' + prefixe, actif);

  if not actif then
  begin
    SetField('ANN_RMDEP', '');
    SetField('ANN_RMANNEE', '');
    SetControlText('TANN_RMNO', '');
  end
  else
  begin
    // récupère l'année d'inscription RCS si existe
    if (GetField('ANN_RCSDATE') <> iDate1900)
      and (GetField('ANN_RMANNEE') = '') then
    begin
      dat := GetField('ANN_RCSDATE');
      DecodeDate(dat, annee, mois, jour);
      SetField('ANN_RMANNEE', IntToStr(annee));
    end;
  end;
end;


procedure TOM_Annuaire.BEMAIL_OnClick(Sender: TObject);
begin
  // fonctionnera même si messagerie non sérialisée :
  // SendMail('',GetControlText('ANN_EMAIL'),'',nil,'',FALSE);
  // ShowNewMessage(0, 0, '', GetControlText('ANN_EMAIL') );
{$IFDEF DP}
  if (V_PGI.MailMethod = mmSMTP) and (VH_DP.SeriaMessagerie) then
    ShowNewMessage('', '', '', GetControlText('ANN_EMAIL'), NoDossier)
  else
{$ENDIF}
    SendMail('', GetControlText('ANN_EMAIL'), '', nil, '', False);
end;


procedure TOM_Annuaire.ChangeAdresse1;
var adresse, sep, sVoieNoCompl_l, sVoieType_l : string;
begin
  if bDuringLoad then exit;

  adresse := '';

  sVoieNoCompl_l := RechDom('YYVOIENOCOMPL', GetControlText('ANN_VOIENOCOMPL'), TRUE);
  sVoieType_l := RechDom('YYVOIETYPE', GetControlText('ANN_VOIETYPE'), TRUE);

  if uppercase(sVoieNoCompl_l) = 'ERROR' then
    sVoieNoCompl_l := '';
  if uppercase(sVoieType_l) = 'ERROR' then
    sVoieType_l := '';

  if GetControlText('ANN_VOIENO') <> '0' then
    adresse := GetControlText('ANN_VOIENO');
  if adresse <> '' then sep := ' ' else sep := '';
  //if THDBValComboBox(GetControl('ANN_VOIENOCOMPL')).Value<>'' then
  if GetControlText('ANN_VOIENOCOMPL') <> '' then
    adresse := adresse + sep + sVoieNoCompl_l;
  if adresse <> '' then sep := ' ' else sep := '';
  //if THDBValComboBox(GetControl('ANN_VOIETYPE')).Value<>'' then
  if GetControlText('ANN_VOIETYPE') <> '' then
    adresse := adresse + sep + sVoieType_l;
  if adresse <> '' then sep := ' ' else sep := '';
  if GetControlText('ANN_VOIENOM') <> '' then
    adresse := adresse + sep + GetControlText('ANN_VOIENOM');
  if GetField('ANN_ALRUE1') <> adresse then SetField('ANN_ALRUE1', adresse);
  if GetControl('Adresse') <> nil then SetControlCaption('Adresse', adresse);
end;

procedure TOM_Annuaire.ChangeAdresseOld;
var
   adresse, sVoieNoCompl_l, sVoieType_l   :string;
begin
     if bDuringLoad then exit;

     adresse := '';
  sVoieNoCompl_l := RechDom('YYVOIENOCOMPL', GetControlText('ANB_OLDVOIENOCOMPL'), TRUE);
  sVoieType_l := RechDom('YYVOIETYPE', GetControlText('ANB_OLDVOIETYPE'), TRUE);

  if uppercase(sVoieNoCompl_l) = 'ERROR' then
    sVoieNoCompl_l := '';
  if uppercase(sVoieType_l) = 'ERROR' then
    sVoieType_l := '';

     if GetControlText('ANB_OLDVOIENO') <> '0' then
        adresse := GetControlText('ANB_OLDVOIENO');
     if GetControlText('ANB_OLDVOIENOCOMPL') <> '' then
        adresse := adresse + ' ' + sVoieNoCompl_l;
     if GetControlText('ANB_OLDVOIETYPE') <> '' then
        adresse := adresse + ' ' + sVoieType_l;
     if GetControlText('ANB_OLDVOIENOM') <> '' then
        adresse := adresse + ' ' + GetControlText('ANB_OLDVOIENOM');
     adresse := Trim (adresse);
     if GetField ('ANB_OLDRUE1') <> adresse then
     begin
          SetField ('ANB_OLDRUE1', adresse);
          SetControlText ('ANB_OLDRUE1', adresse);
     end;
end;

procedure TOM_Annuaire.ANN_VOIENO_OnExit(Sender: TObject);
begin
  ChangeAdresse1;
end;

procedure TOM_Annuaire.ANN_VOIETYPE_OnExit(Sender: TObject);
begin
  ChangeAdresse1;
end;

procedure TOM_Annuaire.ANN_VOIENOCOMPL_OnExit(Sender: TObject);
begin
  ChangeAdresse1;
end;

procedure TOM_Annuaire.ANN_VOIENOM_OnExit(Sender: TObject);
begin
  ChangeAdresse1;
end;

procedure TOM_Annuaire.BPERSONNE_OnClick(Sender: TObject);
begin
  SauveDDP_;//LM20070213  Bof, bof, bof...
  AGLLanceFiche('DP', 'FICHPERSONNE', GetField('ANN_GUIDPER'), GetField('ANN_GUIDPER'), '');
end;

procedure TOM_Annuaire.ANN_ALCP_OnExit(Sender: TObject);
begin
     // $$$ JP 03/05/06 - FQ 10903
     if Trim (GetControlText ('ANN_ALVILLE')) = '' then
{$IFDEF EAGLCLIENT}
        SelectCodePostalVille (TCustomEdit(GetControl('ANN_ALCP')), TCustomEdit(GetControl('ANN_ALVILLE')), True);
{$ELSE}
        SelectCodePostalVille (TDBEdit(GetControl('ANN_ALCP')), TDBEdit(GetControl('ANN_ALVILLE')), True);
{$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 10/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_Annuaire.ANN_PPPM_OnClick(Sender: TObject);
var
  sFamPer_l: string;
  strPPPM: string;
begin
  // on ne fait rien pendant le OnLoadRecord
  if bDuringLoad then exit;

     // Force la modif dans la base, car radiogroup pas réactif !
     // $$$ JP 17/11/2004 - GetControlText pas à jour, à priori ItemIndex à l'air de l'être
{$IFDEF EAGLCLIENT}
  //LM20070125 if THRadioGroup(Sender).ItemIndex = 0 then strPPPM := 'PM' else strPPPM := 'PP';
  strPPPM := THValComboBox(sender).value ; //LM20070125
{$ELSE}
  //LM20070125 if THDBRadioGroup(Sender).ItemIndex = 0 then strPPPM := 'PM' else strPPPM := 'PP';
  strPPPM := THDbValComboBox(sender).value ; //LM20070125
{$ENDIF}

//   if GetControlText('ANN_PPPM')<>GetField('ANN_PPPM') then
  if strPPPM <> GetField('ANN_PPPM') then
  begin
    bDuringLoad := True; // Evite le déclenchement du onclick en eAGL
    SetField('ANN_PPPM', strPPPM); //GetControlText('ANN_PPPM'));
    bDuringLoad := False;
  end;

  sFamPer_l := GetControlText('ANN_FAMPER');
  if (GetField('ANN_PPPM') = 'PP') and
    ((sFamPer_l = 'FIS') or (sFamPer_l = 'SOC') or (sFamPer_l = 'OGA')) then
  begin
    PGIInfo('Un organisme ' + RechDom('JUFAMPER', sFamPer_l, false) + ' ne peut être qu''une personne morale', TitreHalley);
    SetField('ANN_PPPM', 'PM');
  end;

  AfficheOngletPPouPMouORG;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 22/10/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_Annuaire.ANN_SEXE_OnClick(Sender: TObject);
var sCodeSexe_l, sAnnPpPm_l : string;
begin
  sCodeSexe_l := THValComboBox(GetControl('ANN_SEXE')).Value; //LM20070125 sCodeSexe_l := THRadioGroup(GetControl('ANN_SEXE')).Value;
  FiltreCivilite(sCodeSexe_l);
  sAnnPpPm_l   := GetField ('ANN_PPPM');
  SetControlVisible('ANN_NOM3',  (sAnnPpPm_l = 'PP') and (sCodeSexe_l='F'));
  SetControlVisible('TANN_NOM3', (sAnnPpPm_l = 'PP') and (sCodeSexe_l='F'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 05/11/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_Annuaire.FiltreCivilite(sCodeSexe_p: string);
begin
// tout revu... un seul champ pour OK en Eagl ...
(*   if sCodeSexe_p = 'M' then
   begin
      SetControlVisible( 'ANN_TYPECIV', false );
      SetControlVisible( 'ANN_TYPECIVM', true );
      SetControlVisible( 'ANN_TYPECIVF', false );
   end
   else if sCodeSexe_p = 'F' then   // Femmes
   begin
      SetControlVisible( 'ANN_TYPECIV', false );
      SetControlVisible( 'ANN_TYPECIVM', false );
      SetControlVisible( 'ANN_TYPECIVF', true );
   end
   else
   begin
      SetControlVisible( 'ANN_TYPECIV', true );
      SetControlVisible( 'ANN_TYPECIVM', false );
      SetControlVisible( 'ANN_TYPECIVF', false );
   end; *)
  if sCodeSexe_p = 'M' then THDBValCombobox(GetControl('ANN_TYPECIV')).Plus := 'JTC_SEXE="M"'
  else if sCodeSexe_p = 'F' then THDBValCombobox(GetControl('ANN_TYPECIV')).Plus := 'JTC_SEXE="F"'
  else THDBValCombobox(GetControl('ANN_TYPECIV')).Plus := ''

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 10/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_Annuaire.ANN_FAMPER_OnClick(Sender: TObject);
var sFamPer_l: string;
begin
  sFamPer_l := GetControlText('ANN_FAMPER');
  //LM20070502 SetField('ANN_FAMPER', sFamPer_l); //LM20070125 LM>MD pourquoi?
  if InString (sFamPer_l, ['FIS', 'SOC', 'OGA'])
   and (getField('ANN_PPPM')<>'PM') then //LM20070511
    SetField('ANN_PPPM', 'PM');

  SetControlVisible('GRPBOXADMIN', InString (sFamPer_l, ['FIS', 'OGA']) ); //LM20070125
  SetControlVisible('TANN_ALRESID', not InString (sFamPer_l, ['FIS', 'OGA']) ); //LM20070516
  SetControlVisible('ANN_ALRESID', not InString (sFamPer_l, ['FIS', 'OGA']) ); //LM20070516
  SetControlVisible('TANN_ALBAT', not InString (sFamPer_l, ['FIS', 'OGA']) ); //LM20070516
  SetControlVisible('ANN_ALBAT', not InString (sFamPer_l, ['FIS', 'OGA']) ); //LM20070516
  SetControlVisible('TANN_ALESC', not InString (sFamPer_l, ['FIS', 'OGA']) ); //LM20070516
  SetControlVisible('ANN_ALESC', not InString (sFamPer_l, ['FIS', 'OGA']) ); //LM20070516
  SetControlVisible('TANN_ALETA', not InString (sFamPer_l, ['FIS', 'OGA']) ); //LM20070516
  SetControlVisible('ANN_ALETA', not InString (sFamPer_l, ['FIS', 'OGA']) ); //LM20070516
  SetControlVisible('TANN_ALNOAPP', not InString (sFamPer_l, ['FIS', 'OGA']) ); //LM20070516
  SetControlVisible('ANN_ALNOAPP', not InString (sFamPer_l, ['FIS', 'OGA']) ); //LM20070516

  SetControlVisible('GBINFOINTER', not InString (sFamPer_l, ['FIS', 'OGA']) ); //LM20070125
  SetControlProperty('TABADRESSEFISC','TabVisible', not InString (sFamPer_l, ['FIS', 'OGA', 'SOC', 'JUR'])) ;//LM20070502

  FiltreTypePer;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 02/12/2005
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOM_Annuaire.ANN_TYPEPER_OnClick(Sender: TObject);
var
    sTypePer_l : string;
begin
  sTypePer_l := GetControlText('ANN_TYPEPER');
  SetField('ANN_TYPEPER', sTypePer_l);

  TraiteChangeTypePer;
  ANN_FAMPER_OnClick(nil);
end;

procedure TOM_Annuaire.ANN_SITUFAM_OnClick(Sender: TObject);
var
   sSituFam_l :string ;
begin
   sSituFam_l := GetControlText('ANN_SITUFAM') ;
   if sSituFam_l <> 'MAR' then
      SetField('ANN_REGMAT', '');
   AfficheRegimeMatrimonial(sSituFam_l);
end;

procedure TOM_Annuaire.AfficheRegimeMatrimonial(sSituFam_p : string);
var
   bVisible_l : boolean;
begin
  SetControlVisible('ANN_REGMAT', (sSituFam_p = 'MAR'));
  SetControlVisible('TANN_REGMAT', (sSituFam_p = 'MAR'));

  bVisible_l := (sSituFam_p <> 'CEL') and (sSituFam_p <> '');
  SetControlVisible('TANN_GUIDCJ', bVisible_l);
  SetControlVisible('BTSUPGUIDCJ', bVisible_l);
  SetControlVisible('TTANN_GUIDCJ', bVisible_l);
end;

// $$$ JP 05/05/06 - FQ 10637
// cf utomtiers.TOM_TIERS.RTListeEnseigne
{procedure TOM_Annuaire.ANN_ENSEIGNE_ElipsisClick (Sender:TObject);
begin
end;}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 10/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_Annuaire.DossierRattache(SGuidPer_p: String): string;
var
  QRYDossier_l: TQuery;
  bVoirNoDossier: Boolean;
begin
  Result := '';
  QRYDossier_l := OpenSQL('SELECT DOS_NODOSSIER FROM DOSSIER WHERE DOS_GUIDPER = "' + SGuidPer_p + '"', true);
  if not QRYDossier_l.EOF then
    Result := QRYDossier_l.FindField('DOS_NODOSSIER').AsString
  else
    Result := '';
  Ferme(QRYDossier_l);
  if (Ecran.Name = 'ANNUAIRE') then SetControlText('NODOSSIER', Result);
  bVoirNoDossier := (Ecran.Name='ANNUAIRE') and (Result<>'') and (Not AnnuaireLanceeDepuisSynthese);
  SetControlVisible('NODOSSIER', bVoirNoDossier);
  SetControlVisible('TNODOSSIER', bVoirNoDossier);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BM
Créé le ...... : 30/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_Annuaire.OnFormKeyDown(MySender_p: TObject; var MyKey_p: Word; MyShift_p: TShiftState);
begin
  if (MyKey_p = VK_DELETE) and (MyShift_p = [ssCtrl]) then // Ctrl + suppr : supprimer
  begin
    MyKey_p := 0;
    if (TButton(GetControl('BDELETE')) <> nil) and
      (TButton(GetControl('BDELETE')).Visible) and
      (TButton(GetControl('BDELETE')).Enabled) then
      TButton(GetControl('BDELETE')).OnClick(nil);
  end
  else if (MyKey_p = vk_nouveau) and (MyShift_p = [ssCtrl]) then // Ctrl + N : nouveau
  begin
    MyKey_p := 0;
    if (TButton(GetControl('BINSERT')) <> nil) and
      (TButton(GetControl('BINSERT')).Visible) and
      (TButton(GetControl('BINSERT')).Enabled) then
      TButton(GetControl('BINSERT')).OnClick(nil);
  end
  // $$$ JP 12/05/06 - FQ 10903+mail CJ: F5 <=> dble click dans champ code postal
  else if (MyKey_p = VK_F5) and (Ecran.ActiveControl.Name = 'ANN_ALCP') then
{$IFDEF EAGLCLIENT}
        SelectCodePostalVille (TCustomEdit(GetControl('ANN_ALCP')), TCustomEdit(GetControl('ANN_ALVILLE')), True)
{$ELSE}
        SelectCodePostalVille (TDBEdit(GetControl('ANN_ALCP')), TDBEdit(GetControl('ANN_ALVILLE')), True)
{$ENDIF}
  else
    EvDecla.Exec(Ecran.Name, 'OnKeyDown', MySender_p, MyKey_p, MyShift_p); // TFFiche(Ecran).FormKeyDown(MySender_p, MyKey_p, MyShift_p);
end;

function TOM_Annuaire.HasQuote (strValue:string):boolean;
begin
     Result := Pos ('"', strValue) > 0;
end;

procedure TOM_Annuaire.ANB_OLDVOIENO_OnExit(Sender: TObject);
begin
  ChangeAdresseOld;
end;

procedure TOM_Annuaire.ANB_OLDVOIETYPE_OnExit(Sender: TObject);
begin
  ChangeAdresseOld;
end;

procedure TOM_Annuaire.ANB_OLDVOIENOCOMPL_OnExit(Sender: TObject);
begin
  ChangeAdresseOld;
end;

procedure TOM_Annuaire.ANB_OLDVOIENOM_OnExit(Sender: TObject);
begin
  ChangeAdresseOld;
end;

procedure TOM_Annuaire.ChargeRegleFisc;
var Q : TQuery ;
begin
   if ds.State = dsInsert then exit ; //LM20070412

   sDFIRegleFisc_c  := '';
   sDORSectionBNC_c := '';
   sDORFormeAsso_c  := '';

   // left join pour tenir compte de l'absence des enregistrements
   Q := OpenSql('select DOS_NODOSSIER, DFI_REGLEFISC, DOR_SECTIONBNC, DOR_FORMEASSO '
              + 'from DOSSIER left join DPORGA on DOR_GUIDPER=DOS_GUIDPER '
              + 'left join DPFISCAL on DFI_GUIDPER=DOS_GUIDPER '
              + 'where DOS_GUIDPER = "'+GetField('ANN_GUIDPER')+'"', true) ;
   if Not Q.Eof then
     begin
     sDFIRegleFisc_c  := Q.FindField('DFI_REGLEFISC').AsString;
     sDORSectionBNC_c := Q.FindField('DOR_SECTIONBNC').AsString;
     sDORFormeAsso_c  := Q.FindField('DOR_FORMEASSO').AsString;
     end;
   Ferme(Q);

   // #### pour affichage dans fiche annuaire, mais zones invisibles => les supprimer ?
   SetControlText ('DFI_REGLEFISC',  sDFIRegleFisc_c);
   SetControlText ('DOR_SECTIONBNC', sDORSectionBnc_c);
   SetControlText ('DOR_FORMEASSO',  sDORFormeAsso_c);
end;

procedure TOM_Annuaire.EnregRegleFisc;
// Ces zones ont pu être modifiées suite au changement de la forme juridique
var T : TOB;
begin
   T := TOB.Create('DPFISCAL', Nil, -1);
   T.SelectDB('"' + sGuidPer_c + '"', nil);
   T.PutValue('DFI_GUIDPER', sGuidPer_c);
   if T.GetString('DFI_REGLEFISC') <> sDFIRegleFisc_c then
     begin
     T.PutValue('DFI_REGLEFISC', sDFIRegleFisc_c);
     T.InsertOrUpdateDB;
     end;
   T.Free;

   T := TOB.Create('DPORGA', Nil, -1);
   T.SelectDB('"' + sGuidPer_c + '"', nil);
   T.PutValue('DOR_GUIDPER', sGuidPer_c);
   if (T.GetString('DOR_SECTIONBNC') <> sDORSectionBNC_c) or (T.GetString('DOR_FORMEASSO') <> sDORFormeAsso_c) then
      begin
      T.PutValue('DOR_SECTIONBNC', sDORSectionBNC_c);
      T.PutValue('DOR_FORMEASSO', sDORFormeAsso_c);
      T.InsertOrUpdateDB;
      end;
   T.Free;
end;

//mcd 08/03/2006 10835
procedure TOM_Annuaire.majcontact;
var
    TOB_CON: TOB;
    iYear, iMonth, iDay : word;
begin
{$ifndef Annuaire_seul}
  inherited ;
  // MD 20/07/2006 - Si connecté à une base dossier en config multi-dossier,
  // on ne saura pas mettre à jour le contact dans la DB0, ni récupérer le bon code
  // auxiliaire issu du tiers de la DB0 (cf fonction TiersAuxiliaire)
  if not((V_PGI.DbName=V_PGI.DefaultSectionDbName) or (V_PGI.DefaultSectionName='')) then
    exit;

  if GetField('ANN_PPPM') <>'PP' then exit ;
  TOB_CON:=TOB.Create('CONTACT', NIL, -1);  // Création de la TOB des contacts
  if GetField('ANN_TIERS') <> '' then
    begin
    TOB_CON.PutValue('C_TYPECONTACT','T');
    TOB_CON.PutValue('C_AUXILIAIRE',TiersAUxiliaire(GetField('ANN_TIERS'),false)) ;
    end
  else
    begin
    TOB_CON.PutValue('C_TYPECONTACT','ANN');
    TOB_CON.PutValue('C_AUXILIAIRE',GetField('ANN_GUIDPER')) ;
    end;
  TOB_CON.PutValue('C_GUIDPER',GetField('ANN_GUIDPER')) ;
  TOB_CON.PutValue('C_TIERS',GetField('ANN_TIERS')) ;
  TOB_CON.PutValue('C_NUMEROCONTACT','1') ;
  TOB_CON.PutValue('C_NATUREAUXI',GetField ('ANN_NATUREAUXI')) ;
  TOB_CON.PutValue('C_NOM',GetField('ANN_NOM1')) ;
  TOB_CON.PutValue('C_PRENOM',GetField('ANN_NOM2')) ;
  TOB_CON.PutValue('C_TELEPHONE',GetField('ANN_TEL1')) ;
  TOB_CON.PutValue('C_TELEX', GetField ('ANN_TEL2')); // $$$ JP 27/10/06 FQ 11111 //TOB_CON.PutValue('C_TELEX',GetField('ANN_MINITEL'));
  TOB_CON.PutValue('C_FAX', GetField ('ANN_FAX')); // $$$ JP 27/10/06 FQ 11111
  TOB_CON.PutValue('C_RVA',GetField('ANN_EMAIL')) ;
  DecodeDate (GetField('ANN_DATENAIS'), iYear, iMonth, iDay);
  TOB_CON.PutValue('C_JOURNAIS',iday) ;
  TOB_CON.PutValue('C_MOISNAIS',imonth) ;
  TOB_CON.PutValue('C_ANNEENAIS',iyear) ;
  TOB_CON.PutValue('C_SEXE',GetField('ANN_SEXE')) ;
  TOB_CON.PutValue('C_PRINCIPAL','X') ;
  TOB_CON.PutValue('C_CLETELEPHONE',CleTelephone (GetField('ANN_TEL1'))) ;
  TOB_CON.PutValue('C_CLETELEX',CleTelephone (GetField('ANN_TEL2'))) ;// $$$ JP 27/10/06 FQ 11111
  TOB_CON.PutValue('C_CLEFAX',CleTelephone (GetField('ANN_FAX'))) ;// $$$ JP 27/10/06 FQ 11111
  TOB_CON.InsertOrUpdateDB (FALSE);
  TOB_CON.free;
{$endif}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 22/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_Annuaire.OnClick_BForme(Sender: TObject);
Var ads : tdataset ;
begin
ads:=ds ;
   if ChangeFormeJuridique (self, 'ANN_GUIDPER', @ADS) then
      VoirSCI ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 30/05/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_ANNUAIRE.LiensMajDetention;//(iCapNbTitre_p : integer);
var
   OBLiens_l : TOB;
   sReq_l, sTitre_l : string;
   iInd_l : integer;
   iTTNBPP_l, iTTNBUS_l, iTTNBNP_l, iTTNBTOTUS_l, iTTNBTOTPP_l : integer;
   iVoixAgo_l, iVoixAge_l, iFenetre_l : integer;
   dPctVoix_l, dPctBenef_l, dPctCap_l : double;
begin
   if iCapNbTitre_c = GetField('ANN_CAPNBTITRE') then exit;
   iCapNbTitre_c := GetField('ANN_CAPNBTITRE');

   sReq_l := 'select * from ANNULIEN ' +
             'where ANL_GUIDPERDOS = "' + sGuidPer_c + '" ' +
             '  AND ANL_NOORDRE = 1 ' +
             '  AND ANL_TYPEDOS = "STE" ' +
//           '  AND (ANL_TTPCTBENEF <> 0 OR ANL_TTPCTCAP <> 0) ' +
             'ORDER BY ANL_TRI, ANL_NOMPER';


   OBLiens_l := TOB.Create('ANNULIEN', NIL, -1);
   OBLiens_l.LoadDetailDBFromSQL('ANNULIEN', sReq_l);

   for iInd_l := 0 to OBLiens_l.Detail.count - 1 do
   begin
      iFenetre_l := ANLOngletChoix(OBLiens_l.Detail[iInd_l].GetString('ANL_FONCTION'), sTitre_l);
      if iFenetre_l = 2 then
      begin
         iTTNBPP_l    := OBLiens_l.Detail[iInd_l].GetInteger('ANL_TTNBPP');
         iTTNBUS_l    := OBLiens_l.Detail[iInd_l].GetInteger('ANL_TTNBUS');
         iTTNBNP_l    := OBLiens_l.Detail[iInd_l].GetInteger('ANL_TTNBNP');
         iVoixAgo_l   := OBLiens_l.Detail[iInd_l].GetInteger('ANL_VOIXAGO');
         iVoixAge_l   := OBLiens_l.Detail[iInd_l].GetInteger('ANL_VOIXAGE');
         iTTNBTOTUS_l := OBLiens_l.Detail[iInd_l].GetInteger('ANL_TTNBTOTUS');
         iTTNBTOTPP_l := OBLiens_l.Detail[iInd_l].GetInteger('ANL_TTNBTOT');
         dPctBenef_l  := OBLiens_l.Detail[iInd_l].GetDouble('ANL_TTPCTBENEF');
         dPctCap_l    := OBLiens_l.Detail[iInd_l].GetDouble('ANL_TTPCTCAP');

         ANLVoixCalcule(iTTNBPP_l, iTTNBUS_l, iTTNBNP_l,
                        iVoixAgo_l, iVoixAge_l, iTTNBTOTUS_l, iTTNBTOTPP_l);

         dPctVoix_l := ANLTotauxVoixCalcule(iCapNbTitre_c,
                        iVoixAgo_l + iVoixAge_l);

         ANLTotauxPCTCalcule(iCapNbTitre_c,
                        iTTNBTOTUS_l, iTTNBTOTPP_l, dPctBenef_l, dPctCap_l);

         OBLiens_l.Detail[iInd_l].PutValue('ANL_VOIXAGO', iVoixAgo_l);
         OBLiens_l.Detail[iInd_l].PutValue('ANL_VOIXAGE', iVoixAge_l);
         OBLiens_l.Detail[iInd_l].PutValue('ANL_TTNBTOTUS', iTTNBTOTUS_l);
         OBLiens_l.Detail[iInd_l].PutValue('ANL_TTNBTOT', iTTNBTOTPP_l);
         OBLiens_l.Detail[iInd_l].PutValue('ANL_TTPCTVOIX', dPctVoix_l);
         OBLiens_l.Detail[iInd_l].PutValue('ANL_TTPCTBENEF', dPctBenef_l);
         OBLiens_l.Detail[iInd_l].PutValue('ANL_TTPCTCAP', dPctCap_l);
      end;
   end;
   OBLiens_l.UpdateDB;
   OBLiens_l.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PGR
Créé le ...... : 23/06/2006
Modifié le ... :
Description .. : MAJ Base WINNER dans la table DOSSIER,
Suite ........ : qui n'est pas un champ de la tom
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TOM_Annuaire.UpdateBaseWinner;
var
 GuidPerdos, BaseWinner: string;
 Q1: TQuery;
begin
  // on ne fait rien pendant le OnLoadRecord
  if bDuringLoad then exit;

  if (GetParamSocSecur ('SO_MDLIENWINNER', False)) and (GetControl('DOS_WINSTALL')<>Nil) then
  begin
    if GetControlVisible('DOS_WINSTALL') then
    begin
      GuidPerdos := GetField('ANN_GUIDPER');
      BaseWinner := '';
      Q1 := OpenSql('SELECT DOS_WINSTALL FROM DOSSIER WHERE DOS_GUIDPER="'+GuidPerdos+'"', True);

      if not Q1.Eof then
      begin
        BaseWinner := Q1.FindField('DOS_WINSTALL').Asstring;

        if BaseWinner <> GetControlText('DOS_WINSTALL') then
           ExecuteSql('UPDATE DOSSIER SET DOS_WINSTALL = "'+Uppercase(GetControlText('DOS_WINSTALL'))+'" WHERE DOS_GUIDPER="'+GuidPerdos+'"' );
      end;

      Ferme(Q1);
    end;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : PGR
Créé le ...... : 23/06/2006
Modifié le ... :   /  /    
Description .. : Passe en mode édition sur modif de la zone Base Winner
Mots clefs ... :
*****************************************************************}
procedure TOM_Annuaire.ChangeBaseWinner(Sender: TObject);
begin
  // on ne fait rien pendant le OnLoadRecord
  if bDuringLoad then exit;
  ModeEdition(DS);
  // modif factice pour avoir le message de validation
  SetField('ANN_APPAYS', GetField('ANN_APPAYS'));
end;


procedure TOM_Annuaire.ChangeAdrFiscale(Sender: TObject);  //LM20070201
var adr_:tob ;
    adresse : string ;
begin
  if (sender=nil) or bDuringLoad  then exit ;
  EvDecla.Exec(TControl(sender).Name, 'OnExit', sender); //Lance l'exécution du OnExit initial
  adr_:=tob.create('DPPERSO', nil, -1 );
  adr_.getEcran(Ecran);
  Adresse := adr_.g('DPP_ADRESSE1');
  DPP_changeAdresse(adr_) ;
  SetControlText('DPP_ADRESSE1', adr_.g('DPP_ADRESSE1'));
  if Adresse <> adr_.g('DPP_ADRESSE1') then ForceModidier ;//LM20070315
  adr_.free ;
end ;

procedure TOM_Annuaire.CtrlChangeAdrFiscale(Sender: TObject);//LM20070201
begin
  if (sender=nil) or bDuringLoad then exit ;
  EvDecla.Exec(TControl(sender).Name, 'OnExit', sender);
  if DPP_.IsModifie then
  begin
    ModeEdition (DS) ;
    setField('ANN_GUIDPER', GetField('ANN_GUIDPER')) ;
  end ;
end ;

procedure TOM_Annuaire.ForceModidier ;//LM20070315
begin
  if DS<>nil then
  begin
    ModeEdition (DS) ;
    setField('ANN_GUIDPER', GetField('ANN_GUIDPER')) ;
  end ;
end ;

{LM20070511 
procedure TOM_Annuaire.GereOnChange ; //LM20070415
var i : integer ;
    FieldName : variant ;
    nom : string ;
begin
  (*
  scanne toutes zones DB (cad DataField existe) de la fiche Annuaire
  si le nom de la zone est différent du nom du champ,
      on affecte le OnExit du nom de la zone sur handlerChangeField
  Contrainte : toutes les zones DB avec Name<>FieldName ne peuvent gérer un OnExit pour le moment...
  *)
exit ; //attention : revoir las ann_codenaf*.onExit
  for i:=0 to Ecran.ComponentCount-1 do
  begin
    nom:= upperCase(Ecran.Components[i].name) ;
    if (getPropertyVal(Ecran.Components[i], 'DataField', FieldName)) then
    begin
      FieldName :=uppercase(trim(FieldName)) ;
      if (FieldName<>'') and (FieldName<>nom) then
      begin
        if isPropertyExists(Ecran.Components[i], 'OnExit') then
          EvDecla.Rebranche(nom, 'OnExit', handlerChangeField)
        else if isPropertyExists(Ecran.Components[i], 'OnChange') then
          EvDecla.Rebranche(nom, 'OnChange', handlerChangeField)
        else
          EvDecla.Rebranche(nom, 'OnClick', handlerChangeField);
        //debug:=debug+nom +#13#10 ;
      end ;
    end ;
  end ;
end ;
}

procedure TOM_Annuaire.handlerChangeField (sender:TObject) ;//LM20070511  LM20070415
var FldName : variant ;
begin
  if bDuringLoad or (sender=nil) then exit ;
  FldName := TControl(sender).Name ;
  FldName:=copy(FldName, 1, length(FldName)-2) ;//On affiche le 'PP' en, fin de zone ce qui donne le nom du champ
  if getControlText(TControl(sender).Name) <> getField(FldName) then
  begin
    //pgiBox('Reporte la modif de ' + TControl(sender).Name + ' sur ' + FieldName );
    ModeEdition(DS);
    setField(FldName, getControlText(TControl(sender).Name) ) ;
  end ;
end ;

{// MD 20071207 - FQ 11887 désactivé :
procedure TOM_Annuaire.ANN_PAYSNAISChange (sender:TObject) ;//20070425
var b : Boolean ;
begin
  b:= inString(getControltext('ANN_PAYSNAIS') , ['','FRA']) ;
  SetControlEnabled('ANN_SIRENPP', b);
  SetControlEnabled('ANN_CLESIRETPP', b);

  SetControlEnabled('ANN_CODENAFPP', b);
  SetControlEnabled('ANN_CODENAF2PP', b);

  SetControlEnabled('GBRCSPP', b);
  SetControlEnabled('GBPPRM', b);
end ; }

procedure TOM_Annuaire.ANN_CODENAFOnEnter (sender:TObject) ;//20070425
begin
  if sender = nil then exit ;
  codeNafIni := GetControlText(TControl(sender).name) ;
end ;

procedure TOM_Annuaire.ANN_CODENAFOnExit (sender:TObject) ;//20070425
//var s, act, lib  : string ;
//    sl : HTStrings ;
begin

  //if not (sender is THDbEdit) then exit ;
  if sender = nil then exit ;
  if champToNum(TControl(sender).name)<0 then //=nom de zone <> nom de champ  LM20070712
    handlerChangeField (sender) ; //LM20070511

  ChangeCodeNaf(self, TControl(sender), codeNafIni)  ;//LM20070516
  AvertirCodeNaf; // FQ 11699

  {LM20070516
  s := GetControlText(TControl(sender).name) ; //code naf
  if pos('2', TControl(sender).name)>0 then act:='OLE_ACTIVITE2' else act:='OLE_ACTIVITE' ;

  sl:=HTStringList.Create;
  RichToStrings(TRichEdit(getControl(act)), sl) ;

  if ((s<>'') and (Trim(CbpRTFtoText(sl.Text, false, false, 0))=''))
   or (codeNafIni<>s) then  //LM20070511
  begin //si code naf renseigné & activité vide => activité = CO_LIBRE provenant de YYCODENAF
    lib := rechDom('YYCODENAF', s, false, '', true) ;
    if lib<>'' then setControlText(act, lib) ;
  end ;
  sl.free ; }
end ;


procedure TOM_Annuaire.VoirSCI ;
var b:boolean ;
begin
  b:=inString (getField('ANN_CODEINSEE'), ['6541','6542','6540']) ;
  setControlVisible('ANN_TYPESCI', b) ;
  setControlVisible('TANN_TYPESCI', b) ;

  {LM20070516
  if (not bDuringLoad) and b then
  begin
    if getField('ANN_CODEINSEE')= '6540' then      setField('ANN_TYPESCI', 'LOC')
    else if getField('ANN_CODEINSEE')= '6541' then setField('ANN_TYPESCI', 'CVT')
    else if getField('ANN_CODEINSEE')= '6542' then setField('ANN_TYPESCI', 'ATT');
    //nb : le raz est fait si nécessaire sur le OnUpdateRecord
  end ;
  }
end ;


procedure TOM_Annuaire.AvertirCodeNaf;
// FQ 11699
var
    visib : Boolean;
begin
  if GetControl('TAVERTCODENAF')=Nil then exit;
  visib := False;

  // Tests sur CODENAF désactivés pour les organismes fiscaux & OGA
  // (qui d'ailleurs n'ont pas l'onglet personne morale ni physique !)
  // #### voir si on doit autoriser la saisie du CodeNaf sur un OGA
  if ( GetField('ANN_FAMPER')<>'FIS' ) and ( GetField('ANN_FAMPER')<>'OGA' ) and ( GetField('ANN_CODENAF') = '' ) then
    visib := True;

  SetControlVisible('TAVERTCODENAF', visib);
  SetControlVisible('TAVERTCODENAFPP', visib);
end;

initialization
   registerclasses([TOM_Annuaire]);

{$IFDEF BUREAU}
   RegisterAglProc ('FonctionCTI', FALSE, 2, DPAGLFonctionCTI);
{$ENDIF}

end.
