unit dpTOMFiscal;
// TOM de la table DPFISCAL
interface
uses
     StdCtrls,Controls,Classes,db,forms,sysutils, ComCtrls,
     HCtrls,HEnt1,HMsgBox,
     HDB, // => effectue un cast TDBCheckBox = TcheckBox en eagl
     UTOM,UTOB,HTB97,Paramsoc,
{$IFDEF EAGLCLIENT}
     eFiche,
     MaineAGL,
     UtileAGL,
{$ELSE}
     Fiche,
     FE_Main,
     DBCtrls,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     extctrls, //LMO20060901
     Spin,Dialogs,
     EventDecla, //LM20070308
     GalSystem
     ;

Const cnLgNoIntraCom_f = 9;

Type
  TOM_DPFISCAL = Class (TOM)
    procedure OnArgument (stArgument : String ) ; override ;
    procedure OnNewRecord  ; override ;
    procedure OnLoadRecord ; override ;
    procedure OnUpdateRecord  ; override ;
    procedure OnAfterUpdateRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnClose                    ; override ;
    procedure OnChangeCA12(Sender: TObject);

  private
    EvDecla : TEventDecla ;//LM20070302
    sGuidPer_c       : string;
    DateCreation     : TDateTime; //= JUR_DEBACT
    MoisCloture     : word; //LM20070511
    TabVisible       : String;
    sSocName_c, sNoDossier_c : string;
    actif            : Boolean;
    visib            : Boolean;
    bCharge : boolean; // pour éviter les OnChange pendant les onload
    ANL_Lst : Tob ;
    DistribDividIni : boolean ;
    ann_forme, ann_pppm : string ;
    DefTabSheet :  string ;
    bModifLiensFisc : Boolean;
    procedure _setField(champ : string; v : variant ) ;//LM20070329
    procedure Initialisation( sGuidPer_p : string );
    procedure ElipIntegClick(Sender: TObject);
    procedure ElipsisAdhesionClick(Sender: TObject);
    procedure ElipsisUnLienClick(fonction, zone : String);
    procedure DFI_TETEGROUPE_OnClick(Sender: TObject);
    procedure BTGROUPEMEREClick(Sender: TObject);
    procedure DFI_IMPODIR_OnClick(Sender: TObject);
    procedure DFI_REGIMFISCDIR_OnClick(Sender: TObject);
    procedure DFI_OPTIONAUTID_OnClick(Sender: TObject);
    procedure DFI_OPTIONRSS_OnClick(Sender: TObject);
    procedure DFI_EXONERE_OnClick(Sender: TObject);
    procedure DFI_INTEGRAFISC_OnClick(Sender: TObject);
    procedure DFI_OPTIONREPORT_OnClick(Sender: TObject);
    procedure DFI_OPTIONRDSUP_OnClick(Sender: TObject);
    procedure DFI_OPTIONRISUP_OnClick(Sender: TObject);
    procedure DFI_EXIGIBILITE_OnClick(Sender: TObject);
    procedure DFI_OPTIONEXIG_OnClick(Sender: TObject);
    procedure DFI_DECLARATION_OnClick(Sender: TObject);
    procedure DFI_IMPOINDIR_OnClick(Sender: TObject);
    procedure DFI_AUTRESTAXES_OnClick(Sender: TObject);
    procedure DFI_EXONERETP_OnClick(Sender: TObject);
    procedure DFI_DEGREVTREDUC_OnClick(Sender: TObject);
    procedure DFI_ADHEREOGA_OnClick(Sender: TObject);
    //LM20070315 procedure DFI_DEMATERIATDFC_OnClick(Sender: TObject);
    procedure DFI_REGIMFUSION_OnClick(Sender: TObject);
    procedure DFI_DROITSAPPORT_OnClick(Sender: TObject);
    procedure BINTEGRATION_OnClick(Sender: TObject);
    procedure BLIENOGA_OnClick(Sender: TObject);
    procedure DFI_REGIMEINSPECT_OnChange(Sender: TObject);
    procedure DFI_TAXEPROF_OnClick(Sender : TObject);
    procedure GereExoneration;
    procedure GereIntegraFisc;
    procedure GereExonereTp;
    procedure GereDeclaration;
    procedure GereAdhesionOGA;
    procedure GereTDFC (sender : TObject); //LM20070315
    procedure GereExigibilite;
    procedure GereDroitsApport;
    procedure GereRegFusion;
    procedure GereDegrevmtReduc;
    procedure GereAutresTaxes;
    procedure GereImpoIndir;
    procedure GereOptionCA12;     // $$$ JP 03/05/06
    procedure GereCasParticulier; // $$$ JP 03/05/06
    procedure GereImpoDir (sender:TObject) ;//LM20070315
    procedure GereOptionRDSup;
    procedure GereOptionRSS;
    procedure GereOptionReport;
    // $$$ JP 03/05/06 procedure GereOptionRISup;
    procedure GereResultatFiscal;
    procedure GereExoImpositionDirecte;
    procedure InitImpoDir;
    procedure InitOptions;
    function  IsRDIValide: Boolean;
    function  GetGuidPer: String;
    function  ChangeImpoDir: Boolean;
    function  ChangeIntegraFisc: Boolean;
    function  ChargeNOINTRACOMM( sGuidPer_p, sDfiNoIntraComm_p : string ) : string;
    function  ControleNOINTRACOMM : string;
    procedure InfoSocMere (guidper : string) ; //LMO20060901
    //LM20070315 procedure SelectAnnuaire (Sender: TObject);//LMO20060901
    {procedure SelectGED (Sender: TObject) ; //LMO20060901
    procedure RAZLookup (Sender: TObject);//LMO20060901
    procedure showGED (Sender: TObject);//LMO20060901}
    procedure DFI_REGLEFISCclick (Sender: TObject);//LM20070308

    //+LM20070315
    //++Signataire
    procedure Signataire_Init ;
    procedure Signataire_Sauve (prefixe : string) ;
    procedure Signataire_IsModif(sender:TObject);

    procedure AfficheINDouSAL (sender:TObject) ;
    procedure EffaceINDouSAL(sender:TObject) ;

    procedure AfficheSalarie (codeSal, prefixe : string) ;
    procedure SelectSalarie (sender:TObject) ;
    procedure SelectAnnuaire(sender:TObject) ;
    //--Signataire

    procedure DFI_TVADECLAChange (sender:TObject) ;
    procedure initJourDepotTVA ;
    //-LM20070315

    //procedure LiensAnn(sender:TObject) ;
    procedure LibelleOGA ;
    procedure DFI_TAXEVEHICSOnClick (sender:TObject) ;
    procedure DFI_TAXESALAIRESOnClick (sender:TObject) ;
    procedure getAnn_forme  ;
    function  MoisClotureDecale : boolean ;//LM20070511

    procedure EltTP (ok:boolean) ;//LM20070511

    procedure BLIENS_OnClick(sender:TObject); //MD 20070524
    function ongletEnCours:string;
    procedure LiensDAS2(sender:TObject);
    procedure RechercheSIEetMajFicheFiscale;
    procedure MajDpfiscalFromLienSIE(sTypePer, sGuidPerDos, sGuidPer, sNoAdhesion, sNoAdhCompl : String);

  end  ;


///////////// IMPLEMENTATION /////////////
implementation
uses
     dpOutils, AnnOutils, galOutil,
     UDossierSelect,
     DpJurOutils,
     UGedViewer, EntDP,
     usatUtil, choisir, //LM20070315
     ULibWindows, dpTofAnnusel,
     UTomYYAnnulien
     ;

const Signataire = 'LF;CSL;EXP;TVA;R19;D1003;CMVA;PVA;' ; //LM20070315 à gérer en tablette...
//LM20071105 : Remarques sur les signataires :
// - si des signataires sont à ajouter, il faut les mettre en fin de chaîne (pour ANL_FONCTION)
// - Il faut que les 3 premier caractères soit discriminent (champs combo)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/06/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOM_DPFISCAL.OnNewRecord;
var
   sDfiNoIntraComm_l : string;
begin
   inherited;
   bCharge:=True;
   // FQ 11768 : en cas d'utilisation de la tom par code, le sGuidPer_c est vide
   if sGuidPer_c<>'' then
     SetField('DFI_GUIDPER', sGuidPer_c);

   SetField('DFI_TAUXTVA1', 5.5);
   SetField('DFI_TAUXTVA2', 19.6);
   SetField('DFI_TAUXTVA3', 20.6);
   SetField('DFI_IMPOINDIR', 'T');
   SetField('DFI_IMPODIR', 'IS');
   SetField('DFI_REGLEFISC', 'BIP');
   SetField('DFI_REGIMFISCDIR', 'RN');
   SetField('DFI_ACTIVTAXEPRO', 'GEN');
   SetField('DFI_TYPETAXEPRO', 'GEN');  //LM20070315
   SetField('DFI_TAXEPROF', '-');
   SetField('DFI_TVARBTCE', '-');
   SetField('DFI_TVATERR', 'EF');

   if (ds<>nil) and (Ecran.Name <> 'DPCOMFISCAL') and (sGuidPer_c<>'') then
   begin
      sDfiNoIntraComm_l := ChargeNOINTRACOMM(sGuidPer_c, GetField('DFI_NOINTRACOMM'));
      if sDfiNoIntraComm_l <> '' then SetField('DFI_NOINTRACOMM', sDfiNoIntraComm_l);
   end;

  bCharge := False;
end;


procedure TOM_DPFISCAL.OnArgument(stArgument: String);
// reçoit un nom de TabSheet pour afficher directement tel onglet
var Action, sign : String; //LM20070315
    ctrl : TControl ; //LMO200600901
    i : integer ; //LM20070315
    bVoirLiens : Boolean;

    procedure setSignataire_IsModif (pref:string) ;      //LM20070315
    begin
      EvDecla.Rebranche(pref + '_QUALI', 'OnExit', Signataire_IsModif);
      EvDecla.Rebranche(pref + '_PRENOM', 'OnExit', Signataire_IsModif);
      EvDecla.Rebranche(pref + '_NOM', 'OnExit', Signataire_IsModif);
      EvDecla.Rebranche(pref , 'OnExit', Signataire_IsModif);
    end ;

begin
   inherited;
   bCharge := True;
   bModifLiensFisc := False;

   DefTabSheet := ArgumentToTabsheet (stArgument) ;//LM20070528

   EvDecla := TEventDecla.Create(Ecran);      //LM20070308
   ANL_Lst:=Tob.Create('ANNULIEN', nil, -1); //LM20070315

   Action     := ReadTokenSt(stArgument);
   TabVisible := ReadTokenSt(stArgument);
   sGuidPer_c := ReadTokenSt(stArgument);
   if sGuidPer_c = '' then
      sGuidPer_c := TFFiche(Ecran).FLequel;

   Initialisation(sGuidPer_c);
   Ecran.Caption := 'Fiscal : ' + GetNomPer(sGuidPer_c);
   UpdateCaption(Ecran);
   if V_PGI.RunFromLanceur then
   begin
      sSocName_c := V_PGI.DBName;
      sNoDossier_c := V_PGI.NoDossier;
   end
   else
   begin
      sSocName_c := VH_Doss.DBSocName;
      sNoDossier_c := VH_Doss.NoDossier;
   end;

  // Annuelle/Mensuelle/Trimestrielle
  THDBValComboBox(GetControl('DFI_PERIODIIMPIND')).Plus :=
    'AND (CO_CODE="ANN" OR CO_CODE="MEN" OR CO_CODE="TRI")';
    // Autre / Obligation Cautionnée / Virement Bancaire
  THDBValComboBox(GetControl('DFI_MODEPAIEFISC')).Plus :=
    'AND (CO_CODE="AUT" OR CO_CODE="OBL" OR CO_CODE="VIB")';

  THEdit(GetControl('TINTEGRATION')).OnElipsisClick := ElipIntegClick;
  THEdit(GetControl('TADHESION')).OnElipsisClick := ElipsisAdhesionClick;
  TDBCheckBox(GetControl('DFI_TETEGROUPE')).Onclick := DFI_TETEGROUPE_OnClick;
  TToolbarButton97(GetControl('BTGROUPEMERE')).Onclick := BTGROUPEMEREClick;
  //LM>MD 20070315 TToolbarButton97(GetControl('BINTEGRATION')).OnClick := BINTEGRATION_OnClick;
  TToolbarButton97(GetControl('BLIENOGA')).OnClick := BLIENOGA_OnClick;
//  THDBEdit(GetControl('DFI_REGIMEINSPECT')).OnElipsisClick := DFI_REGIMEINSPECT_OnChange;
  THDBEdit(GetControl('DFI_REGIMEINSPECT')).OnExit := DFI_REGIMEINSPECT_OnChange;

  THDBValComboBox(GetControl('DFI_IMPODIR')).OnClick := DFI_IMPODIR_OnClick;
  THDBValComboBox(GetControl('DFI_REGIMFISCDIR')).OnClick := DFI_REGIMFISCDIR_OnClick;
  TDBCheckBox(GetControl('DFI_OPTIONAUTID')).OnClick := DFI_OPTIONAUTID_OnClick;
  TDBCheckBox(GetControl('DFI_OPTIONRSS')).OnClick := DFI_OPTIONRSS_OnClick;
  TDBCheckBox(GetControl('DFI_EXONERE')).OnClick := DFI_EXONERE_OnClick;
  TDBCheckBox(GetControl('DFI_INTEGRAFISC')).OnClick := DFI_INTEGRAFISC_OnClick;
  TDBCheckBox(GetControl('DFI_OPTIONREPORT')).OnClick := DFI_OPTIONREPORT_OnClick;
  TDBCheckBox(GetControl('DFI_OPTIONRDSUP')).OnClick := DFI_OPTIONRDSUP_OnClick;
  TDBCheckBox(GetControl('DFI_OPTIONRISUP')).OnClick := DFI_OPTIONRISUP_OnClick;
  THDBValComboBox(GetControl('DFI_EXIGIBILITE')).OnClick := DFI_EXIGIBILITE_OnClick;
  TDBCheckBox(GetControl('DFI_OPTIONEXIG')).OnClick := DFI_OPTIONEXIG_OnClick;
  TDBCheckBox(GetControl('DFI_DECLARATION')).OnClick := DFI_DECLARATION_OnClick;
  THDBValComboBox(GetControl('DFI_IMPOINDIR')).OnClick := DFI_IMPOINDIR_OnClick;
  TDBCheckBox(GetControl('DFI_AUTRESTAXES')).OnClick := DFI_AUTRESTAXES_OnClick;
  TDBCheckBox(GetControl('DFI_EXONERETP')).OnClick := DFI_EXONERETP_OnClick;
  if GetControl('DFI_DEGREVTREDUC')<>Nil then
    TDBCheckBox(GetControl('DFI_DEGREVTREDUC')).OnClick := DFI_DEGREVTREDUC_OnClick;
  TDBCheckBox(GetControl('DFI_ADHEREOGA')).OnClick := DFI_ADHEREOGA_OnClick;
  //LM20070315 TDBCheckBox(GetControl('DFI_DEMATERIATDFC')).OnClick := DFI_DEMATERIATDFC_OnClick;
  EvDecla.Rebranche('DFI_ISDECLA','OnClick', GereTDFC);//LM20070315
  TDBCheckBox(GetControl('DFI_REGIMFUSION')).OnClick := DFI_REGIMFUSION_OnClick;
  TDBCheckBox(GetControl('DFI_DROITSAPPORT')).OnClick := DFI_DROITSAPPORT_OnClick;
  //BM : 20/06/2004
  TDBCheckBox(GetControl('DFI_TAXEPROF')).OnClick := DFI_TAXEPROF_OnClick;
  // 20/06/01
  SetControlVisible('BINSERT', False);
  //PGR 07/07/2006 Traitement champ Inspection
  if GetParamSocSecur ('SO_MDLIENWINNER', False) then
  begin
    SetControlVisible('TDFI_NOINSPECTION', True);
    SetControlVisible('DFI_NOINSPECTION', True);
  end;

  (*LM20070315
  TRadioButton(GetControl('RBANNEE')).OnClick := OnChangeCA12;
  TRadioButton(GetControl('RBEXERCICE')).OnClick := OnChangeCA12;
  *)
  SetControlVisible ('DFI_DECLARATION', not vh_dp.group ) ;
  setcontrolVisible ('CBCA12DECALE', (getField('DFI_REGIMFISCDIR')='RS')
                                 and (getField('DFI_OPTIONRISUP')<>'X')
                                 and MoisClotureDecale ) ;//LM20070511
  ctrl:= getcontrol('CBCA12DECALE') ;
  if ctrl<>nil then TCheckBox(ctrl).OnClick := OnChangeCA12 ;
  //LM20070315

  EvDecla.Rebranche('DFI_REGLEFISC', 'OnClick', DFI_REGLEFISCclick);
  EvDecla.Rebranche('DFI_TVADECLA', 'OnChange', DFI_TVADECLAChange);
  EvDecla.Rebranche('DFI_ACTDIFF', 'OnClick', GereImpoDir);
  EvDecla.Rebranche('DFI_OPTIONRISUP', 'OnClick', GereImpoDir);
  EvDecla.Rebranche('bIntegration', 'OnClick', BINTEGRATION_OnClick);

  EvDecla.Rebranche('BSie', 'OnClick', BLIENS_OnClick);//LM20070611 +LM20070404 : à "repointer" une fois la refonte de liens développée
  EvDecla.Rebranche('BDAS2CENTRE', 'OnClick', LiensDAS2);

  bVoirLiens := JaiLeDroitConceptBureau(ccVoirLesLiens);
  //MD20070711
  if Not VH_DP.Group and bVoirLiens then
  begin
    SetControlVisible('BLIENS', True);
    EvDecla.Rebranche('BLIENS', 'OnClick', BLIENS_OnClick);
  end;

  //Gestion spécifique des droits sur les liens de la fiche
  SetControlVisible('BSIE', bVoirLiens);
  SetControlVisible('BFRP', False); // Bouton non branché donc inutile ! // bVoirLiens);
  SetControlVisible('BDAS2CENTRE', bVoirLiens);

  // DEBUT FQ 11942
  //--- Onglet Imposition sur le Résultat , déclaration de liasse fiscale et de résultat
  SetControlVisible('TLF_QUALI',False);
  SetControlVisible('LF_QUALI',False);
  SetControlVisible('LF_NOADMIN',False);
  SetControlVisible('LF_GUIDPER',False);
  SetControlVisible('TLF',False);
  SetControlVisible('TLF_NOM',False);
  SetControlVisible('TLF_PRENOM',False);
  SetControlVisible('LF',False);
  SetControlVisible('LF_NOM',False);
  SetControlVisible('LF_PRENOM',False);
  SetControlVisible('LF_BSAL',False);
  SetControlVisible('LF_BLIEN',False);
  SetControlVisible('LF_DEL',False);

  //--- Onglet Taxe sur la valeur ajoutée, Déclaration
  SetControlVisible('TTVA_QUALI',False);
  SetControlVisible('TVA_QUALI',False);
  SetControlVisible('TVA_NOADMIN',False);
  SetControlVisible('TVA_GUIDPER',False);
  SetControlVisible('TTVA',False);
  SetControlVisible('TTVA_NOM',False);
  SetControlVisible('TTVA_PRENOM',False);
  SetControlVisible('TVA',False);
  SetControlVisible('TVA_NOM',False);
  SetControlVisible('TVA_PRENOM',False);
  SetControlVisible('TVA_BSAL',False);
  SetControlVisible('TVA_BLIEN',False);
  SetControlVisible('TVA_DEL',False);

  //--- Onglet Taxe sur la valeur ajoutée, Demande de remboursement
  SetControlVisible('TR19_QUALI',False);
  SetControlVisible('R19_QUALI',False);
  SetControlVisible('R19_NOADMIN',False);
  SetControlVisible('R19_GUIDPER',False);
  SetControlVisible('TR19',False);
  SetControlVisible('TR19_NOM',False);
  SetControlVisible('TR19_PRENOM',False);
  SetControlVisible('R19',False);
  SetControlVisible('R19_NOM',False);
  SetControlVisible('R19_PRENOM',False);
  SetControlVisible('R19_BSAL',False);
  SetControlVisible('R19_BLIEN',False);
  SetControlVisible('R19_DEL',False);
  // FIN FQ 11942

  //+LM20070315
  sign:='*' ; i:=1 ;
  while sign<>'' do
  begin
    sign := trim(gtfs(signataire,';',i)) ;
    if sign='' then break ;
    EvDecla.Rebranche(sign + '_QUALI', 'OnClick', AfficheINDouSAL);
    EvDecla.Rebranche(sign + '_BSAL',  'OnClick', SelectSalarie);
    EvDecla.Rebranche(sign + '_BLIEN', 'OnClick', SelectAnnuaire);
    EvDecla.Rebranche(sign + '_DEL',   'OnClick', EffaceINDouSAL);
    setSignataire_IsModif(sign) ;
    inc(i);
  end ;

  setControlvisible('DFI_ACTIVFISC', (V_PGI.ModePCL='1') ) ;
  //SetControlProperty('TSCOMPLEMENTTVA','TabVisible', (V_PGI.ModePCL='1')) ;
  setControlvisible('GBLIENSERVANTISSIMMO', false (*(V_PGI.ModePCL<>'1') en attendant *) ) ;

  initJourDepotTVA ;
  //-LM20070315

  SetActiveTabSheet('PGENERAL') ;//LM20070611 SetActiveTabSheet('TSIMPOSITIONINDIRECTE') ;

  EvDecla.rebranche('DFI_TAXEVEHICSOC', 'Onclick', DFI_TAXEVEHICSOnClick) ;
  EvDecla.rebranche('DFI_TAXESALAIRES', 'Onclick', DFI_TAXESALAIRESOnClick) ;

  // #### MD 20070524 - En attendant refonte des liens
  if GetControl('BLIENS')<>Nil then
  begin
    TToolbarButton97(GetControl('BLIENS')).OnClick := BLIENS_OnClick;
    TToolbarButton97(GetControl('BLIENS')).Hint := 'Liens organismes fiscaux';
  end;

  if vh_dp.group then setControlvisible ('TSCOMPLEMENTTVA', false);

  // MD20070709 - FQ 11597
  SetControlProperty('DFI_NOINTRACOMM', 'EditMask', '');
  SetControlProperty('DFI_NOINTRACOMM', 'MaxLength', 13);

  bCharge:=False;
end;

procedure TOM_DPFISCAL.OnLoadRecord;
var
    QQ      : TQuery;
    wA, wJ : word;
    Expreced : integer;
    sPPPM_l, Page : string;
    ctrl : TControl ;
begin
  bCharge:=True;

  Expreced := 0;
  Page := ongletEnCours ;

  (*+LM20070315
  QQ := OpenSQL ('SELECT DOR_TENUEEUROCPTA, DOR_DUREEPREC, DOR_DATEFINEX '
    + 'FROM DPORGA WHERE DOR_GUIDPER="'+GetGuidPer+'"',True);*)
  //+LM20070329 QQ := OpenSQL ('select DOR_TENUEEUROCPTA, JUR_DUREEEXPREC, JUR_DATEFINEX, JUR_DEBACT, JUR_CPTESDIV ' +
  QQ := OpenSQL ('select JUR_DUREEEXPREC, JUR_DATEFINEX, JUR_DEBACT, JUR_CPTESDIV ' +
                 'from JURIDIQUE ' +
                 'where JUR_GUIDPERDOS = "'+GetGuidPer+'"',True);
  //-LM20070329
  //-LM20070315

  if not QQ.Eof then
  begin
    DateCreation := QQ.FindField('JUR_DEBACT').AsDateTime;
    Expreced := QQ.FindField('JUR_DUREEEXPREC').AsInteger;

    if IsValidDate( QQ.FindField('JUR_DATEFINEX').AsString ) then
      DecodeDate(strtodate(QQ.FindField ('JUR_DATEFINEX').AsString), wA , MoisCloture, wJ)
    else
      wA := 0;

    {$IFDEF EAGLCLIENT}
    DistribDividIni:=(QQ.FindField('JUR_CPTESDIV').AsString = 'X');
    {$ELSE}
    DistribDividIni:=QQ.FindField('JUR_CPTESDIV').AsBoolean ;
    {$ENDIF}
  end
  else
  begin //LM20070315
    DateCreation := idate1900 ;
    DistribDividIni:= false ;
  end ;

  Ferme (QQ);
  if Expreced = 0 then
    begin
    SetControlVisible('DFI_OLDBENEFFISC', FALSE);
    SetControlVisible('TDFI_OLDBENEFFISC', FALSE);
    SetControlVisible('DFI_OLDREINTEGR', FALSE);
    SetControlVisible('DFI_OLDREDUC', FALSE);
    SetControlVisible('DFI_OLDCA', FALSE);
    SetControlVisible('DFI_ANNEECIVILE', FALSE);
    SetControlVisible('TDFI_OLDCA', FALSE);
    end;

  // valeurs dupliquées pour info
//  SetControlText('LIBREGIMEFISC',THDBValCombobox(GetControl('DFI_REGIMFISCDIR')).Text);
  SetControlText('DFI_ANNEECIVILE',IntToStr(wa));

  // zones visibles/invisibles
  GereExoneration;
  GereImpoDir (nil);//LM20070315
  GereExigibilite;
  GereImpoIndir;
  GereDeclaration;
  GereAutresTaxes;
  GereDegrevmtReduc;
  GereAdhesionOGA;
  GereTDFC(nil);
  GereRegFusion;
  GereDroitsApport;
  GereExonereTp;
  DFI_REGIMFISCDIR_OnClick (nil) ;//LM20070315

  // récup noms d'organismes liés
  SetControlText('TINTEGRATION', RecupNomAnnuLien (GetGuidPer, 'TIF'));
  SetControlText('TADHESION', RecupNomAnnuLien (GetGuidPer, 'OGA'));

  // GRPCA12 zone gérée par pg
  (*//+LM20070315
  if GetField('DFI_TYPECA12')='12' then // à l'année civile
    TRadioButton(GetControl('RBANNEE')).Checked := True
  else if GetField('DFI_TYPECA12')='12E' then // à l'exercice comptable
    TRadioButton(GetControl('RBEXERCICE')).Checked := True ;
  *)
  ctrl:= getcontrol('CBCA12DECALE') ;
  if ctrl<>nil then TCheckBox(ctrl).checked := (getField('DFI_TYPECA12')='12E') ;
  //-LM20070315

  // Conditionnement du champ impot sur la fortune
  sPPPM_l := GetValChamp( 'ANNUAIRE', 'ANN_PPPM', 'ANN_GUIDPER = "'+GetField('DFI_GUIDPER')+'"');

  if sPPPM_l = 'PM' then SetField('DFI_IMPSOLFORTUNE', '-');

  //LM20070315 SetControlEnabled('DFI_IMPSOLFORTUNE', (sPPPM_l = 'PP'));
  SetControlVisible('GRPBOXISF', (sPPPM_l = 'PP')); //FQ 11674
  SetControlEnabled('GRPBOXTVS', (sPPPM_l <> 'PP'));//

  //LM20050315 SetControlVisible('TSTAXEPROF', TDBCheckBox(GetControl('DFI_TAXEPROF')).checked);

  // $$$JP 20/10/2003: interdire modification du n° frp si un CDI existe dans la liste des intervenants
  // BM 27/01/2003 : interdire tout court
  //PGR 07/07/2006 Frp modifiable si Winner
  if GetParamSocSecur ('SO_MDLIENWINNER', True) = false then
    SetControlEnabled ('DFI_NOFRP', FALSE);
{  if ExisteSQL ('SELECT ANL_CODEPER FROM ANNULIEN WHERE ANL_GUIDPERDOS="' + GetGuidPer + '" AND ANL_TYPEDOS="DP" AND ANL_FONCTION="FIS" AND ANL_TYPEPER="CDI"') = TRUE then
      SetControlEnabled ('DFI_NOFRP', FALSE)
  else
      SetControlEnabled ('DFI_NOFRP', TRUE);}

//  DFI_REGIMEINSPECT_OnChange(Nil);

  if vh_dp.Group then //LMO20060901
  begin
    SetControlChecked('CBTS', (getField('DFI_IMPOINDIR')<>'T'));
    //SetControlChecked('CBTVA', ??);
    InfoSocMere(getField('DFI_GUIDPER') ) ;
  end ;

  Signataire_init ;//LM20070315
  DFI_TAXESALAIRESOnClick (nil);
  DFI_TAXEVEHICSOnClick (nil);
  DFI_TVADECLAChange (nil) ;//FQ11558
  getAnn_Forme ;

  if DefTabSheet<>'' then
    SetActiveTabSheet(DefTabSheet) //LM20070528
  else if Page = 'PGeneral' then //FQ11511
    SetFocusControl( 'DFI_REGLEFISC' ) //LM20070611
  else if Page = 'TSIMPOSITIONINDIRECTE' then
    SetFocusControl( 'DFI_IMPOINDIR' )
  else if Page = 'TSAUTRESIMPOTS' then
    SetFocusControl('DFI_TAXEPROF')
  else if Page = 'TSTAXEPROF' then
    SetFocusControl( 'DFI_ACPTEJUIN' )
  else if Page = 'TSAUTRELEMENT' then
    SetFocusControl( 'DFI_ADHEREOGA' );

  VerrouilleSiBaseComptable(self, GetGuidPer) ;//LM20070620
  bCharge:=False;
end;


procedure TOM_DPFISCAL.OnUpdateRecord;
// Tous les contrôles de données
var
   sign, valchp, ann_nationalite  : String;
   coh : boolean;
   i : integer ;
   Dt : TDateTime ;
   tvaOk:boolean ;
   bMajFormatSignataire : boolean ;//LM20071105
begin
  if (ds<>Nil) and (Ecran.Name='DPCOMFISCAL') then exit;

  if (GetField('DFI_REGLEFISC') = '') then
    begin
    ErreurChamp('DFI_REGLEFISC', 'PGeneral', Self, 'La catégorie fiscale doit être renseignée.');
    exit;
    end;

  // ------- Onglet Général (Imposition directe) ------
  if (GetField('DFI_NOFRP') <> '') then
    begin
    // 04/02/01 Coh := CoherenceFrp (GetField('DFI_NOFRP'));
    Coh := CtrlCodeFRP(False, GetField('DFI_NOFRP'));
    if Coh = FALSE then
      begin
      ErreurChamp('DFI_NOFRP', 'PGeneral', Self, 'Numéro FRP incorrect');
      exit;
      end;
    end;
  if ds<>Nil then //mcd 14/06/07 pour ok Veriftob
  begin
  if (TDBCheckBox(GetControl('DFI_OPTIONRDSUP')).checked) and
   ((StrToDate(GetField('DFI_DATEOPTRDSUP')) < DateCreation)
   or (StrToDate(GetField('DFI_DATEOPTRDSUP')) = iDate1900))then
    begin
    ErreurChamp('DFI_DATEOPTRDSUP', 'PGeneral', Self, 'Cette date ne peut être inférieure à la date de création') ;
    exit;
    end;
  if (TDBCheckBox(GetControl('DFI_OPTIONRSS')).checked) and
   ((StrToDate(GetField('DFI_DATEOPTRSS')) < DateCreation)
   or (StrToDate(GetField('DFI_DATEOPTRSS')) = iDate1900)) then
    begin
    ErreurChamp('DFI_DATEOPTRSS', 'PGeneral', Self, 'Cette date ne peut être inférieure à la date de création') ;
    exit;
    end;
  if (TDBCheckBox(GetControl('DFI_OPTIONREPORT')).checked) AND
   ((StrToDate(GetField('DFI_DATEOPTREPORT')) < DateCreation)
   or (StrToDate(GetField('DFI_DATEOPTREPORT')) = iDate1900)) then
    begin
    ErreurChamp('DFI_DATEOPTREPORT', 'PGeneral', Self, 'Date ne peut être inférieure à la date de création') ;
    exit;
    end;
    end;
  if GetField('DFI_BENEFFISC') = Null then
    SetField('DFI_BENEFFISC',0.0);
  if GetField('DFI_REINTEGR') = Null then
    SetField('DFI_REINTEGR', 0.0);
  if GetField('DFI_REDUC') = Null then
    SetField('DFI_REDUC', 0.0);

  if ((GetField('DFI_BENEFFISC') <> 0.0) or (GetField('DFI_REINTEGR') <> 0.0)
   or (GetField('DFI_REDUC') <> 0.0)) then
    begin
    if GetField('DFI_DATEMAJBENEF') = Null then
      begin
      ErreurChamp('DFI_DATEMAJBENEF', 'PGeneral', Self, 'Date de modification est obligatoire') ;
      exit;
      end;
    if StrtoDate(GetField('DFI_DATEMAJBENEF')) = iDate1900 then
      begin
      ErreurChamp('DFI_DATEMAJBENEF', 'PGeneral', Self, 'Date de modification est obligatoire') ;
      exit;
      end;
    if GetField('DFI_ORIGMAJBENEF') = '' then
      begin
      ErreurChamp('DFI_ORIGMAJBENEF', 'PGeneral', Self, 'Origine de l''information est obligatoire') ;
      exit;
      end;
    end;

  // ------- Onglet Imposition indirecte -------
  if ds <>nil then //mcd 14/06/07 pour ok veriftob
  begin
  if (TDBCheckBox(GetControl('DFI_OPTIONEXIG')).checked) and
   (StrToDate(GetField('DFI_DATEOPTEXIG')) < DateCreation) then
    begin
    ErreurChamp('DFI_DATEOPTEXIG', 'TSIMPOSITIONINDIRECTE', Self, 'Date ne peut être inférieure à la date de création') ;
    exit;
      end;
    if (TDBCheckBox(GetControl('DFI_OPTIONEXIG')).checked) and
    (StrToDate(GetField('DFI_DATEOPTEXIG')) < DateCreation) then
      begin
      ErreurChamp('DFI_DATEOPTEXIG', 'TSIMPOSITIONINDIRECTE', Self, 'Date ne peut être inférieure à la date de création') ;
      exit;
      end;
  end;
  if GetField('DFI_TAUXTVA1') > 100 then
    begin
    ErreurChamp('DFI_TAUXTVA1', 'TSIMPOSITIONINDIRECTE', Self, 'Taux ne peut être supérieur à 100%') ;
    exit;
    end;
  if GetField('DFI_TAUXTVA2') > 100 then
    begin
    ErreurChamp('DFI_TAUXTVA2', 'TSIMPOSITIONINDIRECTE', Self, 'Taux ne peut être supérieur à 100%') ;
    exit;
    end;
  if GetField('DFI_TAUXTVA3') > 100 then
    begin
    ErreurChamp('DFI_TAUXTVA3', 'TSIMPOSITIONINDIRECTE', Self, 'Taux ne peut être supérieur à 100%') ;
    exit;
    end;

  // déclaration cochée : tout le panel obligatoire
  if (ds<>Nil )then  //mcd 14/06/07 pour Ok veriftob
  begin
   if ((V_PGI.ModePCL='1') and (TDBCheckBox(GetControl('DFI_DECLARATION')).checked))
    or (vh_dp.Group) then //LM20070502
   begin
      if GetField('DFI_PERIODIIMPIND')='' then
      begin
         ErreurChamp('DFI_PERIODIIMPIND', 'TSIMPOSITIONINDIRECTE', Self, 'La périodicité est obligatoire');
         exit;
      end;
      valchp := ControleNOINTRACOMM;
      if valchp = '' then
      begin
         ErreurChamp('DFI_NOINTRACOMM', 'TSIMPOSITIONINDIRECTE', Self, 'Le numéro intracommunautaire est obligatoire');
         exit;
      end;
      if Length( valchp )>13 then
      begin
         ErreurChamp('DFI_NOINTRACOMM', 'TSIMPOSITIONINDIRECTE', Self, 'La longueur du numéro intracommunautaire est de 13 caractères maxi');
         exit;
      end;

      ann_nationalite := GetValChamp('ANNUAIRE', 'ANN_NATIONALITE', 'ANN_GUIDPER = "'+GetField('DFI_GUIDPER')+'"' ); //LM20070315
      if (ann_nationalite='FRA') and (copy(getField('DFI_NOINTRACOMM'),1,2)<>'FR') then //LM20070315
      begin
         ErreurChamp('DFI_NOINTRACOMM', 'TSIMPOSITIONINDIRECTE', Self, 'Le n° intracommunautaire devrait commencer par FR.');
         exit;
      end;

      if (ann_nationalite<>'FRA') and (getField('DFI_TVATERR') = 'EF') then //LM20070315
      begin
         ErreurChamp('DFI_TVATERR', 'TSIMPOSITIONINDIRECTE', Self, 'La territorialité de la TVA est incohérente.');
         exit;
      end;

      if GetField('DFI_MODEPAIEFISC')='' then
      begin
         ErreurChamp('DFI_MODEPAIEFISC', 'TSIMPOSITIONINDIRECTE', Self, 'Le mode de règlement est obligatoire');
         exit;
      end;
      valchp := VarToStr(GetField('DFI_JOURDECLA'));
      if (valchp='')
       or (Not Isnumeric(valchp))
        or (Not (StrToInt(valchp) in [5,15,16,17,19,20,21,23,24])) then
      begin
         ErreurChamp('DFI_JOURDECLA', 'TSIMPOSITIONINDIRECTE', Self, 'Le jour de dépot de la déclaration doit être 5, 15, 16, 17, 19, 20, 21, 23, ou 24') ;
         exit;
      end;

      if (getControltext('DFI_IMPOINDIR') <> 'E') and (GetField('DFI_REGIMEINSPECT')='') then
      begin
         ErreurChamp('DFI_REGIMEINSPECT', 'TSIMPOSITIONINDIRECTE', Self, 'Le régime de TVA est obligatoire');
         exit;
      end;

      valchp := RechDOM('LTVATYPE', GetField('DFI_REGIMEINSPECT'), False);
      if UpperCase( valchp ) = 'ERROR' then
      begin
         ErreurChamp('DFI_REGIMEINSPECT', 'TSIMPOSITIONINDIRECTE', Self, 'Le régime d''inspection est incorrect') ;
         exit;
      end;
      // obligation d'avoir un RDI valide
      If Not IsRDIValide then
      begin
         PgiInfo('Attention : vous devez préciser un service des impôts (SIE) !');
         // #### On ne fait plus la saisie à cet endroit : trop de pbs ! (voir FQ 11768, 11995...)
         {
         ErreurChamp('', 'TSIMPOSITIONINDIRECTE', Self, 'Vous devez préciser une recette des impôts ou un SIE') ;
         // on affiche que les recettes des impôts
         // #### dommage que ça ne préselectionne pas le ANL_TYPEPER à RDI dans le mul !
         AGLLanceFiche('DP', 'LIENINTERVENANT', 'ANL_GUIDPERDOS='+GetGuidPer+';ANL_FONCTION=FIS;ANL_TYPEPER=RDI','',GetGuidPer+';FIS');
         // pb : on ouvre les liens alors que la fiche est en cours de modif, donc on flaggue...
         bModifLiensFisc := True;
         If Not IsRDIValide then
            exit;
         }
      end;
      // mise à jour du dossier
      UpdateParamSocDossier('SO_ZGERETVA', 'X');
   end
   else
    begin
    // mise à jour du dossier
    UpdateParamSocDossier('SO_ZGERETVA', '-');
    end; // fin cas Déclaration cochée
   end; // fin ds<>nil

  // On contrôle le jour, même si déclaration non cochée
  valchp := VarToStr(GetField('DFI_JOURDECLA'));
  if (ds<>Nil) and (Isnumeric(valchp)) and ( (StrToInt(valchp)<>0) or TDBCheckBox(GetControl('DFI_DECLARATION')).checked)
   and (Not (StrToInt(valchp) in [5,15,16,17,19,20,21,23,24])) then
  begin
     ErreurChamp('DFI_JOURDECLA', 'TSIMPOSITIONINDIRECTE', Self, 'Le jour de dépot de la déclaration doit être 5, 15, 16, 17, 19, 20, 21, 23, ou 24') ;
     exit;
  end;

  // ------- Onglet Taxe Professionnelle -------
  if ds<>nil then //mcd 14/06/07 pour Ok veriftob
  if GetControl('DFI_DEGREVTREDUC')<>Nil then
     if (TDBCheckBox(GetControl('DFI_DEGREVTREDUC')).checked) and
     ((GetField('DFI_MTTDEGREVTRED') = 0.0) or (GetField('DFI_MTTDEGREVTRED') = Null)) then
      begin
      ErreurChamp('DFI_MTTDEGREVTRED','TSTAXEPROF', Self, 'Montant du dégrèvement obligatoire') ;
      exit;
      end;
  if StrToDate(GetField('DFI_DATEFINEXOTP')) < StrToDate(GetField('DFI_DATEDEBEXOTP')) then
    begin
    ErreurChamp('DFI_DATEFINEXOTP', 'TSTAXEPROF', Self, 'Date de fin d''exonération est inférieure à la date de début') ;
    exit;
    end;

  //Dates d'options ..
  if (StrToDate(GetField('DFI_DATEFINEXO')) < StrToDate(GetField('DFI_DATEDEBEXO'))) then  //LM20070315
    begin
    ErreurChamp('DFI_DATEFINEXO', 'PGeneral', Self, 'La date de fin d''exonération est inférieure à la date de début d''exonération ') ;
    exit;
    end;

  Dt := StrToDate(GetField('DFI_DATEDEBEXO')) ;
  if (Dt<>idate1900) and (DateCreation > Dt) then  //LM20070315
    begin
    ErreurChamp('DFI_DATEDEBEXO', 'PGeneral', Self, 'La date de début d''exonération est inférieure à la date de début d''activité') ;
    exit;
    end;

  Dt := StrToDate(GetField('DFI_DATECONVTDFC')) ;
  if (Dt<>idate1900) and (DateCreation > Dt) then  //LM20070315
    begin
    ErreurChamp('DFI_DATECONVTDFC', 'PGeneral', Self, 'La date d''option TDFC est inférieure à la date de début d''activité') ;
    exit;
    end;

  Dt := StrToDate(GetField('DFI_DATEADHOGA')) ;
  if (Dt<>idate1900) and (DateCreation > Dt) then  //LM20070315
    begin
    ErreurChamp('DFI_DATEADHOGA', 'PGeneral', Self, 'La date d''adhésion OGA est inférieure à la date de début d''activité') ;
    exit;
    end;

  Dt := StrToDate(GetField('DFI_DATECONVEDITVA')) ;
  if (Dt<>idate1900) and (DateCreation > Dt) then  //LM20070315
    begin
    ErreurChamp('DFI_DATECONVEDITVA', 'TSIMPOSITIONINDIRECTE', Self, 'La date d''option EDI-TVA est inférieure à la date de début d''activité') ;
    exit;
    end;

   if (ann_pppm = 'PP') and (GetField('DFI_IMPODIR') = 'IS') then //LM20070315
   begin
    ErreurChamp('DFI_IMPODIR', 'PGeneral', Self, 'L''imposition est incohérente pour une personne physique.') ;
    exit ;
   end ;

  if ((GetField('DFI_REGIMFISCDIR') = 'FOR') and (GetField('DFI_REGLEFISC') <> 'BA'))
   or ((GetField('DFI_REGIMFISCDIR') = 'MIC') and  (GetField('DFI_REGLEFISC') = 'BA'))
    or ((GetField('DFI_REGIMFISCDIR') = 'FOR') and (GetField('DFI_REGLEFISC') = 'BA') and (GetField('DFI_IMPODIR') = 'IS'))
   then //LM20070315
  begin
    ErreurChamp('DFI_REGIMFISCDIR', 'PGeneral', Self, 'Le régime d''imposition est incohérent pour la catégorie fiscale.') ;
    exit ;
  end ;

  if ((GetField('DFI_REGLEFISC') = 'BNC') and (not InString(GetField('DFI_IMPODIR'),['IR','TF'])))
   or   ((GetField('DFI_IMPODIR')='IS') and (not InString(GetField('DFI_REGIMFISCDIR'), ['RN', 'RS']))) //LM20070502
    or ((getField('DFI_REGIMFISCDIR')='FOR') and (getField('DFI_REGLEFISC')='BA') and (getField('DFI_IMPODIR')='IS')) //LM20070502
    then //LM20070315
  begin
    ErreurChamp('DFI_IMPODIR', 'PGeneral', Self, 'L''imposition est incohérente pour le régime.') ;
    exit ;
  end ;

  tvaOk := true ;//+LM20070611
  if inString(GetField('DFI_REGIMEINSPECT'), ['EM', 'ET', 'ESM', 'AM']) then
    if (GetField('DFI_REGIMFISCDIR') = 'RN')
     or ((GetField('DFI_REGIMFISCDIR') = 'RS') and (GetField('DFI_OPTIONRISUP')='X')) then else tvaOk := false ;
  if inString(GetField('DFI_REGIMEINSPECT'), ['RM', 'RT']) then
    if inString(GetField('DFI_REGIMFISCDIR'), ['MIC', 'RS']) (* + 'RS' LM20070912 *) then else tvaOk := false ;
  if inString(GetField('DFI_REGIMEINSPECT'), ['AET']) then
    if inString(GetField('DFI_REGLEFISC'), ['BA']) then else tvaOk := false ;
  if inString(GetField('DFI_REGIMEINSPECT'), ['EST']) then
    if (GetField('DFI_REGIMFISCDIR') = 'RN')
     or ((GetField('DFI_REGIMFISCDIR') = 'RS') and (GetField('DFI_OPTIONRISUP')<>'X'))  then else tvaOk := false ;
  if inString(GetField('DFI_REGIMEINSPECT'), ['RSM', 'RST']) then
    if inString(GetField('DFI_REGIMFISCDIR'), ['RS']) and (GetField('DFI_OPTIONRISUP')<>'X')  then else tvaOk := false ;
  if inString(GetField('DFI_REGIMEINSPECT'), ['ST']) then
    if ((GetField('DFI_REGIMFISCDIR') = 'RS') and (GetField('DFI_OPTIONRISUP')<>'X'))
     or ((GetField('DFI_REGIMFISCDIR') = 'MIC') and (GetField('DFI_OPTIONRSS')='X')) then else tvaOk := false ;
  if inString(GetField('DFI_REGIMEINSPECT'), ['AST']) then
    if inString(GetField('DFI_REGLEFISC'), ['BA'])
     and inString(GetField('DFI_REGIMFISCDIR'), ['RS'])
      and (GetField('DFI_OPTIONRISUP')<>'X')  then else tvaOk := false ;
  if inString(GetField('DFI_REGIMEINSPECT'), ['ADT']) then
    if (GetField('DFI_REGLEFISC') = 'BA') then else tvaOk := false ;
  if inString(GetField('DFI_REGIMEINSPECT'), ['EOM', 'EOT']) then
    if ((GetField('DFI_REGLEFISC') = 'BNC') and (GetField('DFI_REGIMFISCDIR') = 'RN'))
     or ((GetField('DFI_REGIMFISCDIR') = 'RS') and (GetField('DFI_OPTIONRISUP')='X'))  then else tvaOk := false ;
  if not tvaOk then
  begin
    ErreurChamp('DFI_REGIMEINSPECT', 'TSIMPOSITIONINDIRECTE', Self, 'Le régime de tva est incohérent.') ;
    //exit ;
  end ; //-LM20070611


  if (GetField('DFI_REGLEFISC')='BIP') and (not inString(GetField('DFI_IMPODIR'), ['IS', 'IR']))
   or (GetField('DFI_REGLEFISC')='RP') and (inString(GetField('DFI_IMPODIR'), ['IS', 'IR'])) then //LM20070315
    ErreurChamp('DFI_REGLEFISC', 'pGeneral', Self, 'La catégorie fiscale est incohérente.') ;

  if (GetField('DFI_EXONERATION')='X') and isValidDate(GetField('DFI_DATEDEBEXO'))
   and (strToDate(GetField('DFI_DATEDEBEXO'))>iDate1900) then //LM20070315
    ErreurChamp('DFI_EXONERATION', 'pGeneral', Self, 'La date de début d''éxonération est incohérente.') ;

  if (GetField('DFI_EXONERATIONTP')='X') and isValidDate(GetField('DFI_DATEDEBEXOTP'))
   and (strToDate(GetField('DFI_DATEDEBEXOTP'))>iDate1900) then //LM20070315
    ErreurChamp('DFI_EXONERATIONTP', 'TSTAXEPROF', Self, 'La date de début d''éxonération est incohérente.') ;

  //+LM20070315
  bMajFormatSignataire:=false ;
  if Anl_lst <>nil then //mcd 20/12/07 sinon plante si appel hors saisie
  begin
    for i:=0 to ANL_Lst.detail.count-1 do  //LM20071105
    begin
      if ANL_Lst.detail[i].g('ANL_NOORDRE')=1 then
      begin
        bMajFormatSignataire:=true ;
        break ;
      end ;
    end ;
  end;
  sign:='*' ; i:=1 ;
  if ds<>Nil then //mcd 14/06/07 pour OK veriftob
  while sign<>'' do
  begin
    sign := trim(gtfs(signataire,';',i)) ;
    if sign='' then break ;
    Signataire_Sauve (sign) ;
    inc(i);
  end ;
  if bMajFormatSignataire then //LM20071105
      executeSQL('delete from ANNULIEN ' +
                 'where ANL_GUIDPERDOS = "' + GetGuidPer + '" ' +
                 'and ANL_TYPEDOS = "DP" '+
                 'and ANL_FONCTION="SGN" and ANL_NOORDRE=1 and ANL_DECLARATION<>""') ;

  //-LM20070315

  if DistribDividIni<>StrToBool_(GetField('DFI_DISTRIBDIVID')) then    //LM20070315
    executeSql('update JURIDIQUE set JUR_CPTESDIV="' + GetField('DFI_DISTRIBDIVID') + '" ' +
               'where JUR_GUIDPERDOS="' + GetField('DFI_GUIDPER') + '" AND JUR_TYPEDOS="STE" AND JUR_NOORDRE=1')  ;//LM>MD JUR_TYPEDOS & JUR_NOORDRE vraiment utile?

  if (GetField('DFI_IMPOINDIR') = 'E') then //LM20070329
  begin
    _SetField('DFI_EXIGIBILITE', '');
    _SetField('DFI_TAUXTVA1', 0);
    _SetField('DFI_TAUXTVA2', 0);
    _SetField('DFI_TAUXTVA3', 0);
  end;


  if not (GetField('DFI_IMPODIR')='IR') and (InString(GetField('DFI_REGIMFISCDIR'), ['RN', 'RS'])) then  //LM20070425
    setField ('DFI_OPTIONAUTID', '-');
  if not (getField('DFI_REGIMFISCDIR')='FOR') and (getField('DFI_REGLEFISC')='BA') then
      setField ('DFI_BAFORFAIT', '');

  if getField('DFI_EXIGIBILITE') <> 'TE' then
  begin
    setField('DFI_OPTIONEXIG', '-') ;
    setField('DFI_DATEOPTEXIG',iDate1900) ;
  end ;

  setField('DFI_DEMATERIATDFC', iif(inString(getControltext('DFI_ISDECLA'),['010','020']), 'X', '-'));

  // Récup infos DPFISCAL issues du lien SIE/RDI/CDI
  if bModifLiensFisc then RechercheSIEetMajFicheFiscale;
  bModifLiensFisc := False;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/06/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_DPFISCAL.OnAfterUpdateRecord;
begin
 inherited;

 //--- FQ 11782 Gestion de la TVA sur les encaissements
 if (DBExists ('DB'+VH_DOSS.NoDossier)) then
  begin
   if (GetField('DFI_EXIGIBILITE')<>'TE') then
    ExecuteSQL ('UPDATE DB'+VH_DOSS.NoDossier+'.DBO.PARAMSOC SET SOC_DATA="-" WHERE SOC_NOM="SO_OUITVAENC"')
   else
    begin
     if (GetField('DFI_OPTIONEXIG')<>'X') then
      ExecuteSQL ('UPDATE DB'+VH_DOSS.NoDossier+'.DBO.PARAMSOC SET SOC_DATA="X" WHERE SOC_NOM="SO_OUITVAENC"')
     else
      ExecuteSQL ('UPDATE DB'+VH_DOSS.NoDossier+'.DBO.PARAMSOC SET SOC_DATA="-" WHERE SOC_NOM="SO_OUITVAENC"');
    end;
  end;


// FQ 11768 : nième tentative d'éviter le carnage en désactivant la proc :
//   if IsInside(Ecran) then
//      ReloadTomInsideAfterInsert(TFFiche(Ecran), DS, ['DFI_GUIDPER'], [sGuidPer_c]);
end;


procedure TOM_DPFISCAL.OnChangeField(F: TField);
begin
  if Ecran.Name='DPCOMFISCAL' then exit;

(*  if (F.FieldName = 'DFI_REGIMFISCDIR') and (DS.State in [dsInsert,dsEdit]) then
    SetControlText('LIBREGIMEFISC',THDBValCombobox(GetControl('DFI_REGIMFISCDIR')).Text);*)

  if (F.FieldName = 'DFI_REGIMFISCDIR') and (GetField('DFI_REGIMFISCDIR') = '') then
    THDBValCombobox(GetControl('DFI_REGIMFISCDIR')).Value:='RN';

  if (F.FieldName = 'DFI_REGIMFISCDIR') then
    if (GetField('DFI_REGIMFISCDIR')='RN') then
      begin
      TDBCheckBox(GetControl('DFI_OPTIONRDSUP')).checked := FALSE;
      TDBCheckBox(GetControl('DFI_OPTIONRISUP')).checked := FALSE;
      TDBCheckBox(GetControl('DFI_OPTIONRSS')).checked := FALSE;
      end;

  if (F.FieldName = 'DFI_ACTIVFISC') and (GetField('DFI_ACTIVFISC') = '') then
    THDBValCombobox(GetControl('DFI_ACTIVFISC')).Value:='AUT';
  if (F.FieldName = 'DFI_DROITSAPPORT') and (DS.State in [dsInsert,dsEdit])
   and (not (TDBCheckBox(GetControl('DFI_DROITSAPPORT')).checked)) then
    SetField('DFI_DATEAPPORT', iDate1900);
  if (F.FieldName = 'DFI_REGIMFUSION') and (DS.State in [dsInsert,dsEdit])
   and (not (TDBCheckBox(GetControl('DFI_REGIMFUSION')).checked)) then
    SetField('DFI_DATEREGFUS', iDate1900);
  if (F.FieldName = 'DFI_DEMATERIATDFC') and (DS.State in [dsInsert,dsEdit])
   and (not (TDBCheckBox(GetControl('DFI_DEMATERIATDFC')).checked)) then
    SetField('DFI_DATECONVTDFC', iDate1900);
  if (F.FieldName = 'DFI_ADHEREOGA') and (DS.State in [dsInsert,dsEdit])
   and (not (TDBCheckBox(GetControl('DFI_ADHEREOGA')).checked)) then
    begin
    SetField('DFI_DATEADHOGA', iDate1900);
    SetField('DFI_NOADHOGA', '');
    end;
  if (F.FieldName = 'DFI_DROITSAPPORT') and (DS.State in [dsInsert,dsEdit])
   and (not (TDBCheckBox(GetControl('DFI_DROITSAPPORT')).checked)) then
    SetField('DFI_DATEAPPORT', iDate1900);

  if (F.FieldName = 'DFI_EXONERETP') and (DS.State in [dsInsert,dsEdit])
   and (not (TDBCheckBox(GetControl('DFI_EXONERETP')).checked)) then
    begin
    SetField('DFI_DATEDEBEXOTP', iDate1900);
    SetField('DFI_DATEFINEXOTP', iDate1900);
    SetControlEnabled ('DFI_DATEDEBEXOTP', FALSE);
    SetControlEnabled ('DFI_DATEFINEXOTP', FALSE);
    SetControlEnabled ('DFI_EXONERATIONTP', FALSE);
    end;
  if (F.FieldName = 'DFI_EXONERETP') and (DS.State in [dsInsert,dsEdit])
   and ((TDBCheckBox(GetControl('DFI_EXONERETP')).checked)) then
    begin
    SetControlEnabled ('DFI_DATEDEBEXOTP', TRUE);
    SetControlEnabled ('DFI_DATEFINEXOTP', TRUE);
    SetControlEnabled ('DFI_EXONERATIONTP', TRUE);
    end;

  if (F.FieldName = 'DFI_AUTRESTAXESA') and (DS.State in [dsInsert,dsEdit])
   and (not (TDBCheckBox(GetControl('DFI_AUTRESTAXES')).checked)) then
    SetField('DFI_TYPETAXETVA', '');

  if (F.FieldName = 'DFI_OPTIONREPORT') and (DS.State in [dsInsert,dsEdit])
   and (not (TDBCheckBox(GetControl('DFI_OPTIONREPORT')).checked)) then
    SetField('DFI_DATEOPTREPORT', iDate1900);
  if (F.FieldName = 'DFI_OPTIONRSS') and (DS.State in [dsInsert,dsEdit])
   and (not (TDBCheckBox(GetControl('DFI_OPTIONRSS')).checked)) then
    SetField('DFI_DATEOPTRSS', iDate1900);
  if (F.FieldName = 'DFI_EXONERE') and (DS.State in [dsInsert,dsEdit])
   and (not (TDBCheckBox(GetControl('DFI_EXONERE')).checked)) then
    begin
    SetField('DFI_DATEDEBEXO', iDate1900);
    SetField('DFI_DATEFINEXO', iDate1900);
    end;
  if (F.FieldName = 'DFI_OPTIONRDSUP') and (DS.State in [dsInsert,dsEdit])
   and (not (TDBCheckBox(GetControl('DFI_OPTIONRDSUP')).checked)) then
    SetField('DFI_DATEOPTRDSUP', iDate1900);

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 29/06/2004
Modifié le ... :
Description .. : Corrige des valeurs sur l'enregistrement demandé dans
Suite ........ : FLequel du AglLanceFiche, avant même son chargement.
Suite ........ :
Suite ........ : MD : devrait être fait dans OnLoadRecord mais comme la
Suite ........ : fiche ne permet pas de scroller sur d'autres enregistrements,
Suite ........ : on corrige uniqt l'enreg en cours, avant même de le charger
Suite ........ : (cela évite le message "voulez-vous enregistrer les
Suite ........ : modifications" alors qu'on n'a rien modifié.)
Mots clefs ... :
*****************************************************************}
procedure TOM_DPFISCAL.Initialisation( sGuidPer_p : string );
var
   sDfiNoIntraComm_l : string;
   OBFiscal_l : TOB;
   bNew_l : boolean;
begin
   if sGuidPer_p='' then exit;
   THLabel (Getcontrol ('TDFI_TYPETAXEPRO')).Visible:=False;
   THDBValCombobox(GetControl('DFI_TYPETAXEPRO')).Visible:=False;

   // on ne fait rien sur un enregistrement pas encore créé
   if TFFiche(Ecran).TypeAction in [taCreat..taCreatOne] then exit;

   OBFiscal_l := TOB.Create('DPFISCAL', nil, -1);
   OBFiscal_l.LoadDetailDBFromSQL('DPFISCAL', 'SELECT * ' +
                                  'FROM DPFISCAL WHERE DFI_GUIDPER = "' + sGuidPer_p + '"');
   if OBFiscal_l.Detail.Count = 0 then
   begin
      OBFiscal_l.Free;
      exit;
   end;

   // initialisation des zones
   if (OBFiscal_l.Detail[0].GetValue('DFI_TAUXTVA1') = 0.0) then
      OBFiscal_l.Detail[0].PutValue('DFI_TAUXTVA1', 5.5);

   if (OBFiscal_l.Detail[0].GetValue('DFI_TAUXTVA2') = 0.0) then
      OBFiscal_l.Detail[0].PutValue('DFI_TAUXTVA2', 19.6);

   if (OBFiscal_l.Detail[0].GetValue('DFI_TAUXTVA3') = 0.0) then
      OBFiscal_l.Detail[0].PutValue('DFI_TAUXTVA3', 20.6);

   if (OBFiscal_l.Detail[0].GetValue('DFI_IMPOINDIR') = '') then
      OBFiscal_l.Detail[0].PutValue('DFI_IMPOINDIR', 'T');

   if (OBFiscal_l.Detail[0].GetValue('DFI_IMPODIR') = '') then
      OBFiscal_l.Detail[0].PutValue('DFI_IMPODIR', 'IS');

   if (OBFiscal_l.Detail[0].GetValue('DFI_REGIMFISCDIR') = '') then
      OBFiscal_l.Detail[0].PutValue('DFI_REGIMFISCDIR', 'RN');

   if (OBFiscal_l.Detail[0].GetValue('DFI_ACTIVTAXEPRO') = '') then
      OBFiscal_l.Detail[0].PutValue('DFI_ACTIVTAXEPRO', 'GEN');

   bNew_l := (OBFiscal_l.Detail[0].IsFieldModified('DFI_TAUXTVA1')) and
             (OBFiscal_l.Detail[0].IsFieldModified('DFI_TAUXTVA2')) and
             (OBFiscal_l.Detail[0].IsFieldModified('DFI_TAUXTVA3')) and
             (OBFiscal_l.Detail[0].IsFieldModified('DFI_IMPOINDIR')) and
             (OBFiscal_l.Detail[0].IsFieldModified('DFI_IMPODIR')) and
             (OBFiscal_l.Detail[0].IsFieldModified('DFI_REGIMFISCDIR')) and
             (OBFiscal_l.Detail[0].IsFieldModified('DFI_ACTIVTAXEPRO'));

   // Initialisation par défaut pour taxe professionnelle
   if bNew_l then
      OBFiscal_l.Detail[0].PutValue('DFI_TAXEPROF', 'X');

   if Ecran.Name <>'DPCOMFISCAL' then
   begin
      sDfiNoIntraComm_l := ChargeNOINTRACOMM( sGuidPer_p, OBFiscal_l.Detail[0].GetValue('DFI_NOINTRACOMM') );
      if sDfiNoIntraComm_l <> '' then
         OBFiscal_l.Detail[0].PutValue('DFI_NOINTRACOMM', sDfiNoIntraComm_l);
   end;

   if OBFiscal_l.Detail[0].Modifie then
      OBFiscal_l.Detail[0].UpdateDB;

   OBFiscal_l.Free;
end;

procedure TOM_DPFISCAL.ElipIntegClick(Sender: TObject);
// clic sur l'elipsis de la zone affichant la tête de groupe
begin
  ElipsisUnLienClick('TIF', 'TINTEGRATION');
end;


procedure TOM_DPFISCAL.ElipsisAdhesionClick(Sender: TObject);
begin
  ElipsisUnLienClick('OGA', 'TADHESION');
end;


procedure TOM_DPFISCAL.ElipsisUnLienClick(fonction, zone : String);
// clic sur un elipsis de la zone affichant un lien (TIF ou OGA)
var NomPer : String;
    Q1     : TQuery;
    Nb     : Integer;
    GuidPer : String;
begin
  NomPer := ''; // nom de la personne liée
  GuidPer := ''; // guid personne

  // Inactif si on n'a pas les droits "voir les liens"
  if not JaiLeDroitConceptBureau(ccVoirLesLiens) then exit;

  // #### peut mieux faire !
  ModeEdition(DS) ;

  // nb de liens de ce type
  Nb := NbEnreg('ANNULIEN', '*', 'WHERE ANL_GUIDPERDOS="'+GetGuidPer+'"'
   +' AND ANL_FONCTION="'+fonction+'"');

  // si un seul lien
  if (Nb=1) then
  begin
    // récup du lien
    Q1 := OpenSQL('SELECT ANL_NOMPER, ANL_GUIDPER FROM ANNULIEN WHERE ANL_GUIDPERDOS="'
      +GetGuidPer+'" AND ANL_FONCTION="'+fonction+'"', True);
    if not Q1.eof then
    begin
      NomPer := Q1.Fields[0].AsString;
      GuidPer := Q1.Fields[1].AsString;
    end;
    Ferme(Q1);
    // affiche directement sa fiche annuaire
    AGLLanceFiche('YY','ANNUAIRE',GuidPer,GuidPer,'ACTION=MODIFICATION')  ;
  end
  else
    // sinon la liste des liens pour faire un choix
    NomPer := AGLLanceFiche('DP','LIENINTERVENANT','ANL_GUIDPERDOS='+GetGuidPer
     +';ANL_FONCTION='+fonction,'',GetGuidPer+';'+fonction);

  // lien choisi
  if (NomPer <> '') then
    SetControlText(zone, NomPer)
  else
    // aucun choisi, on récupère le 1er trouvé (si possible)
    SetControlText(zone, RecupNomAnnuLien(GetGuidPer, fonction));
end;


procedure TOM_DPFISCAL.DFI_TETEGROUPE_OnClick(Sender: TObject);
var Msg : String;
begin
  // on coche tête de groupe alors qu'on avait des liens
  if (TDBCheckBox(GetControl('DFI_TETEGROUPE')).checked)
   and (GetControlText('TINTEGRATION') <> '') then
    begin
    Msg := 'Cette opération entraînera la suppression de tous les liens société mère. Voulez-vous supprimer ?';
    if PGIAsk(Msg, TitreHalley)=mrYes then
      begin
      ExecuteSQL('Delete from ANNULIEN where ANL_GUIDPERDOS="' + GetGuidPer +  '" and ANL_FONCTION="TIF"');
      SetControlText ('TINTEGRATION', '');
      end
    else
      TDBCheckBox(GetControl('DFI_TETEGROUPE')).checked := FALSE;
    end;
  GereIntegraFisc;
end;


procedure TOM_DPFISCAL.BTGROUPEMEREClick(Sender: TObject);
// affiche la liste des liens société mère
begin
  // On commence par enregistrer car les liens peuvent vouloir exploiter les infos de la fiche !
//  if not TFFiche(Ecran).Bouge(TNavigateBtn(NbPost)) then exit;
  if DS.State<>dsBrowse then
  begin
    if TFFiche(Ecran).EnregOk then
      TFFiche(Ecran).BValider.Click
    else
      exit;
  end;

  AGLLanceFiche('DP','LIENINTERVENANT','ANL_GUIDPERDOS='+GetGuidPer+';ANL_FONCTION=TIF','',GetGuidPer+';TIF');
  // Affichage 1ère société mère trouvée
  SetControlText('TINTEGRATION', RecupNomAnnuLien (GetField('DFI_GUIDPER'), 'TIF'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 13/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_DPFISCAL.ChargeNOINTRACOMM( sGuidPer_p, sDfiNoIntraComm_p : string ) : string;
var
   sPaysCle_l, sSiren_l, sNoIntraComm_l, sSirenIntra_l : string;
begin
   result := '';
   sPaysCle_l := StrRAdd( Copy( sDfiNoIntraComm_p, 1, 4 ), ' ', 4 );
   if Copy( sDfiNoIntraComm_p, 1, 4 ) <> '' then exit;
   sSirenIntra_l := Copy( sDfiNoIntraComm_p, 5, Length(sDfiNoIntraComm_p) - 4 );
   sSiren_l := GetValChamp( 'ANNUAIRE', 'ANN_SIREN', 'ANN_GUIDPER = "'+sGuidPer_p+'"');

   if ( sSiren_l <> '' ) and ( sSiren_l <> sSirenIntra_l ) then
      sSirenIntra_l := sSiren_l;

   sSirenIntra_l := StrRAdd( sSirenIntra_l, ' ', cnLgNoIntraCom_f );;
   sNoIntraComm_l := sPaysCle_l + sSirenIntra_l;

   if sDfiNoIntraComm_p <> sNoIntraComm_l then
      result := sNoIntraComm_l;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 13/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_DPFISCAL.ControleNOINTRACOMM : string;
var
   sCodePays_l, sCleInfo_l, sSiren_l : string;
begin
   sCodePays_l := Trim( Copy( GetField('DFI_NOINTRACOMM' ), 1, 2 ));
   sCleInfo_l := Trim( Copy( GetField('DFI_NOINTRACOMM' ), 3, 2 ));
   sSiren_l := Trim( Copy( GetField('DFI_NOINTRACOMM' ), 5, cnLgNoIntraCom_f ));
   result := sCodePays_l + sCleInfo_l + sSiren_l;
end;


procedure TOM_DPFISCAL.GereExoneration;
begin
  actif := TDBCheckBox(GetControl('DFI_EXONERE')).checked;
  SetControlEnabled('DFI_DATEDEBEXO', actif);
  SetControlEnabled('DFI_DATEFINEXO', actif);
  SetControlEnabled('DFI_EXONERATION', actif);
  SetControlEnabled('TDFI_DATEDEBEXO', actif);
  SetControlEnabled('TDFI_DATEFINEXO', actif);
  SetControlEnabled('TDFI_EXONERATION', actif);
end;

procedure TOM_DPFISCAL.GereIntegraFisc;
begin
  If (Not JaiLeDroitConceptBureau(ccModifAnnuaireOuLiens) and Not JaiLeDroitConceptBureau(ccCreatAnnuaireOuLiens)
     and ((AnnuaireLanceeDepuisSynthese and not JaiLeDroitConceptBureau(ccModifAnnuaireDossier)) or not AnnuaireLanceeDepuisSynthese)) // {+GUH - 01/2008 FQ 11858}
  or Not JaiLeDroitConceptBureau(ccVoirLesLiens) then
    begin
    SetControlEnabled('TINTEGRATION', False);
    SetControlEnabled('BTGROUPEMERE', False);
    SetControlEnabled('BINTEGRATION', False);
    end
  else
    begin
    actif := TDBCheckBox(GetControl('DFI_INTEGRAFISC')).checked;

    SetControlEnabled('DFI_TETEGROUPE', actif);
    SetControlEnabled('BTGROUPEMERE', actif and (getControltext('DFI_TETEGROUPE')='-' )); //FQ11462
    if vh_dp.Group then SetControlEnabled('CBSOCINTEGREE', actif);//LMO20060901

    actif := getControltext('DFI_TETEGROUPE') = 'X'  ;
    SetControlVisible('BINTEGRATION', actif);//FQ11462 SetControlEnabled('BINTEGRATION', actif);
    SetControlEnabled('TINTEGRATION', actif);

    if TDBCheckBox(GetControl('DFI_INTEGRAFISC')).checked then
      begin
      actif := (TDBCheckBox(GetControl('DFI_TETEGROUPE')).checked);
      SetControlEnabled('BINTEGRATION', actif);
      SetControlEnabled('TINTEGRATION', Not actif);
      //LM20070425 ?? SetControlEnabled('BTGROUPEMERE', Not actif);
      end;
    if vh_dp.group then //LMO20060901
      begin
      SetControlEnabled('GBFILIALE', GetControlEnabled('TINTEGRATION'));
      SetControlChecked('CBSOCINTEGREE', not (getControlText('DFI_TETEGROUPE')='X'));
      SetControlEnabled('GRPBOXINTEGRATION', GetControlEnabled('BINTEGRATION'));
      end ;
    end;

end;

procedure TOM_DPFISCAL.GereExonereTp;
begin
  actif := TDBCheckBox(GetControl('DFI_EXONERETP')).checked;
  SetControlEnabled('DFI_DATEDEBEXOTP', actif);
  SetControlEnabled('DFI_DATEFINEXOTP', actif);
  SetControlEnabled('DFI_EXONERATIONTP', actif);
  SetControlEnabled('TDFI_DATEDEBEXOTP', actif);
  SetControlEnabled('TDFI_DATEFINEXOTP', actif);
  SetControlEnabled('TDFI_EXONERATIONTP', actif);
end;

procedure TOM_DPFISCAL.GereDeclaration;
begin
   actif := (GetControlText('DFI_IMPOINDIR') <> 'E');

   // $$$ JP 03/05/06 - FQ 11020 - on cache
   SetControlVisible ('GRPBOXDECLA', actif);

   SetControlEnabled ('DFI_DECLARATION', actif );//LM20070502
   if not actif then
      SetControlChecked( 'DFI_DECLARATION', false );

   // $$$ JP 03/05/06 - FQ 11020: inutile puisque groupe caché
{   SetControlEnabled('TDFI_PERIODIIMPIND', actif);
   SetControlEnabled('DFI_PERIODIIMPIND', actif);

   SetControlEnabled('TDFI_NOINTRACOMM', actif);
   SetControlEnabled('DFI_NOINTRACOMM', actif);

   SetControlEnabled('TDFI_JOURDECLA', actif);
   SetControlEnabled('DFI_JOURDECLA', actif);

   SetControlEnabled('TDFI_MODEPAIEFISC', actif);
   SetControlEnabled('DFI_MODEPAIEFISC', actif);

   SetControlEnabled('TDFI_REGIMEINSPECT', actif);
   SetControlEnabled('DFI_REGIMEINSPECT', actif);}
end;

procedure TOM_DPFISCAL.GereAdhesionOGA;
begin
  actif := TDBCheckBox(GetControl('DFI_ADHEREOGA')).checked;
  If (Not JaiLeDroitConceptBureau(ccModifAnnuaireOuLiens) and Not JaiLeDroitConceptBureau(ccCreatAnnuaireOuLiens))
  or Not JaiLeDroitConceptBureau(ccVoirLesLiens) then
    begin
    SetControlEnabled('TADHESION', False);
    SetControlEnabled('BLIENOGA', False);
    end
  else
    begin
    SetControlEnabled('TADHESION', actif);
    SetControlEnabled('BLIENOGA', actif);
    end;

  //SetControlEnabled('TADHESION', actif);
  //SetControlEnabled('BLIENOGA', actif);
  SetControlEnabled('DFI_NOADHOGA', actif);
  SetControlEnabled('TDFI_NOADHOGA', actif);
  SetControlEnabled('DFI_DATEADHOGA', actif);
  SetControlEnabled('TDFI_DATEADHOGA', actif);

  if (not actif) and presenceComplexe('ANNULIEN', ['ANL_GUIDPERDOS','ANL_FONCTION'], ['=','='], [GetGuidPer,'OGA'], ['S','S'])
   and (pgiAsk('Confirmez vous la suppression du lien avec l''organisme de gestion agréé?') = mrYes) then //LM20070315
    ExecuteSQL('delete from ANNULIEN where ANL_GUIDPERDOS = "' +GetGuidPer+ '" and ANL_FONCTION="OGA"');

end ;

procedure TOM_DPFISCAL.GereTDFC (sender:TObject); //LM20070315
begin
  actif := inString(getControltext('DFI_ISDECLA'),['010','020'] ) ;//<> EFI et Cerfa LM20070315 TDBCheckBox(GetControl('DFI_DEMATERIATDFC')).checked;
  SetControlEnabled('DFI_DATECONVTDFC', actif);
  SetControlEnabled('TDFI_DATECONVTDFC', actif);
end;

procedure TOM_DPFISCAL.GereExigibilite;
begin
  visib := (GetControlText('DFI_EXIGIBILITE')='TE');
  SetControlVisible('DFI_OPTIONEXIG', visib);
  SetControlVisible('DFI_DATEOPTEXIG', visib);
  SetControlVisible('TDFI_DATEOPTEXIG', visib);

  actif := (visib) and (TDBCheckBox(GetControl('DFI_OPTIONEXIG')).checked);
  SetControlEnabled('DFI_DATEOPTEXIG', actif);
  SetControlEnabled('TDFI_DATEOPTEXIG', actif);

end;

procedure TOM_DPFISCAL.GereDroitsApport;
begin
  actif := TDBCheckBox(GetControl('DFI_DROITSAPPORT')).checked;
  SetControlEnabled('DFI_DATEAPPORT', actif);
  SetControlEnabled('TDFI_DATEAPPORT', actif);
end;

procedure TOM_DPFISCAL.GereRegFusion;
begin
  actif := TDBCheckBox(GetControl('DFI_REGIMFUSION')).checked;
  SetControlEnabled('DFI_DATEREGFUS', actif);
  SetControlEnabled('TDFI_DATEREGFUS', actif);
end;

procedure TOM_DPFISCAL.GereDegrevmtReduc;
begin
  actif := False;
  if GetControl('DFI_DEGREVTREDUC')<>Nil then
    actif := TDBCheckBox(GetControl('DFI_DEGREVTREDUC')).checked;
  SetControlEnabled('DFI_MTTDEGREVTRED', actif);
end;

procedure TOM_DPFISCAL.GereAutresTaxes;
begin
  visib := TDBCheckBox(GetControl('DFI_AUTRESTAXES')).checked;
  SetControlVisible('DFI_TYPETAXETVA', visib);
  SetControlVisible('TDFI_TYPETAXETVA', visib);
end;

procedure TOM_DPFISCAL.GereImpoIndir;
begin

  visib := (GetControlText('DFI_IMPOINDIR') <> 'E');

  setControlEnabled('GBTVADECLARATION', visib) ;
  DFI_TAXESALAIRESOnClick(nil);//LM20070511 
  SetControlVisible('DFI_EXIGIBILITE', visib);
  SetControlVisible('TDFI_EXIGIBILITE', visib);

  SetControlVisible('DFI_TAUXTVA1', visib);
  SetControlVisible('TDFI_TAUXTVA1', visib);
  SetControlVisible('LBPCTUSUEL', visib);

  SetControlVisible('DFI_TAUXTVA2', visib);
  SetControlVisible('TDFI_TAUXTVA2', visib);
  SetControlVisible('LBPCTNORMAL', visib);

  SetControlVisible('DFI_TAUXTVA3', visib);
  SetControlVisible('TDFI_TAUXTVA3', visib);
  SetControlVisible('LBPCTAUTRE', visib);

  if not visib then
  begin
    SetField('DFI_EXIGIBILITE', '');
    SetField('DFI_TAUXTVA1', 0);
    SetField('DFI_TAUXTVA2', 0);
    SetField('DFI_TAUXTVA3', 0);
  end;

  GereExigibilite;
  GereDeclaration;

  // $$$ JP 02/05/06 - FQ 11020
  GereOptionCA12;
  GereCasParticulier;

  visib := inString(GetControlText('DFI_IMPOINDIR'), ['T', 'P']);
  SetControlVisible('GBTAUXTVA', visib) ;
  SetControlVisible('GRPBOXCA' , visib) ;
  SetControlVisible('GRPBOXCASPART' , (GetControlText('DFI_IMPOINDIR') = 'P') ) ;
  SetControlProperty( 'TSCOMPLEMENTTVA','TabVisible',
                      (V_PGI.ModePCL='1')
                      and (visib or (GetControlText('DFI_IMPOINDIR') = 'P') )) ;
end;

procedure TOM_DPFISCAL.GereImpoDir (sender:TObject) ;//LM20070315
var RegimeInspect, regimp, regfisc, catFisc: String;
begin

  // affiche les bonnes cases à cocher d'options selon le régime fiscal
  regimp := GetControlText('DFI_IMPODIR');
  regfisc := GetControlText('DFI_REGIMFISCDIR');
  catFisc := getControlText('DFI_REGLEFISC');
  RegimeInspect := getControlText('DFI_REGIMEINSPECT');

  // Option autre impo dir
  visib := (regimp='IR') and InString(regfisc,['RN', 'RS']) ;//LM20070315 and (regfisc='RN');
  SetControlVisible('DFI_OPTIONAUTID', visib);
  SetControlEnabled('DFI_OPTIONAUTID', visib);

  SetControlEnabled('GBLIASSEETRESULTAT', (regimp<>'NI'));//LM20070425

  // Option TF ou NI
  visib := (regimp <> 'TF') and (regimp <> 'NI');
  SetControlVisible('DFI_REGIMFISCDIR', visib);
  SetControlVisible('TDFI_REGIMFISCDIR', visib);
  if not visib then
  begin
    SetField('DFI_REGIMFISCDIR', 'RN');
    regfisc := 'RN';
  end;

  // Option NI
  visib := (regimp <> 'NI');
  SetControlVisible('TDFI_ACTIVFISC', visib);
  SetControlVisible('DFI_ACTIVFISC', visib);
  SetControlVisible('TDFI_NOFRP', visib);
  SetControlVisible('DFI_NOFRP', visib);
  if not visib then
  begin
    SetField('DFI_ACTIVFISC', '');
    SetField('DFI_NOFRP', '');
  end;
  GereResultatFiscal;
  GereExoImpositionDirecte;

  // Option RN
  visib := (regfisc='RS');
  SetControlVisible('DFI_OPTIONRDSUP', visib or (regfisc='MIC') );//LM20070627
  SetControlVisible('DFI_DATEOPTRDSUP', visib);
  SetControlVisible('TDFI_DATEOPTRDSUP', visib);
  GereOptionRDSup;

  // Option RSS
  // $$$ JP 03/05/06 - FQ 11020
  visib := (regimp = 'IR') and (regfisc = 'MIC');//LM20070611 FQ11456 retour v7! (regimp = 'IR') and (inString(regfisc, ['RS', 'MIC'])) ;//LM20070315 on tourne en rond ! and (regfisc = 'MIC'); //((regfisc='MIC') or (regfisc='RS'));
  SetControlVisible('DFI_OPTIONRSS', visib);
  SetControlVisible('DFI_DATEOPTRSS', visib);
  SetControlVisible('TDFI_DATEOPTRSS', visib);
  GereOptionRSS;

  // Option report

  //FQ11459 tourne en rond => reportV7 visib :=  ((Instring(regimp,['IS', 'IR']) and TDBCheckBox(GetControl('DFI_OPTIONAUTID')).checked)) ;
  visib := ((regimp='IR') and (regfisc='RN') and (TDBCheckBox(GetControl('DFI_OPTIONAUTID')).checked))
            or ((regimp='IS') and (regfisc<>'MIC'));
  SetControlVisible('DFI_OPTIONREPORT', visib);
  SetControlVisible('DFI_DATEOPTREPORT', visib);
  SetControlVisible('TDFI_DATEOPTREPORT', visib);
  GereOptionReport;

  // Intégration fiscale (comme report mais en RN)
  visib := (regimp='IS') and (regfisc='RN');//LM20070315
  SetControlVisible('GRPBOXINTEGRATION', visib);
  if vh_dp.group then SetControlVisible('GBFILIALE', visib);//LMO20060901
  GereIntegraFisc;

  // Option RN dans imposition indirecte
  GereOptionCA12;
  // $$$ JP 03/05/06 GereOptionRISup; // Type de déclaration CA12

  // Type de forfait bénéfices agricoles
  visib := (regfisc='FOR') and (catFisc='BA');//LM20070315
  SetControlVisible('DFI_BAFORFAIT', visib);
  SetControlVisible('TDFI_BAFORFAIT', visib);

  SetControlVisible('DFI_REGETRANGER', catFisc='BIP'); //LM20070315

  visib := (regimp='IS') and (getControlText('DFI_OPTIONAUTID') = 'X') ;
  SetControlVisible('DFI_CONTREVENUSLOC', visib ) ;//LM20070315
  SetControlVisible('DFI_DISTRIBDIVID', (regimp='IS') ) ;//LM20070611 FQ11463 LM20070315

  LibelleOGA ;

  (*+LM20070425
  SetControlVisible('DFI_ACTDIFF', inString(regfisc , ['RN', 'RSI'])
                                  and (getControlText('DFI_OPTIONRISUP')='X') ) ;//LM20070315
  SetControlVisible('DFI_AUTRESTAXES', inString(regfisc , ['RN', 'RSI'])
                                  and (getControlText('DFI_OPTIONRISUP')='X') ) ;//LM20070315
  *)
  SetControlVisible('DFI_ACTDIFF', inString(RegimeInspect , ['EM', 'ET']) ) ;
  SetControlVisible('DFI_AUTRESTAXES', inString(RegimeInspect , ['EM', 'ET', 'RM', 'RS', 'RT']) ) ;
  //-20070425  la bonne, c quoi ?

  SetControlVisible('TDFI_TYPETAXEPRO', catFisc='BNC') ;//LM20070315
  SetControlVisible('DFI_TYPETAXEPRO', catFisc='BNC') ;//LM20070315

  setcontrolVisible ('CBCA12DECALE',  (regfisc='RS')
                                  and (GetControlText('DFI_OPTIONRISUP')<>'X')
                                  and MoisClotureDecale ) ; //LM20071105
end;

procedure TOM_DPFISCAL.InitImpoDir;
// purger les zones quand on change de régime fiscal
begin
  // pourrait être factorisé avec InitOptions...
  SetField('DFI_OPTIONRISUP', '-');
  SetField('DFI_OPTIONRDSUP', '-');
  SetField('DFI_DATEOPTRDSUP', idate1900);
  SetField('DFI_OPTIONRSS', '-');
  SetField('DFI_DATEOPTRSS', idate1900);
  SetField('DFI_OPTIONREPORT', '-');
  SetField('DFI_DATEOPTREPORT', idate1900);
end;

procedure TOM_DPFISCAL.GereResultatFiscal;
begin
  SetControlVisible('GRPBOXRESULT', visib);

  if not visib then
  begin
    SetField('DFI_BENEFFISC', 0);
    SetField('DFI_REINTEGR', 0);
    SetField('DFI_REDUC', 0);
    SetField('DFI_DATEMAJBENEF', iDate1900);
    SetField('DFI_ORIGMAJBENEF', '');
    SetField('DFI_OLDBENEFFISC', 0);
    SetField('DFI_OLDREINTEGR', 0);
    SetField('DFI_OLDREDUC', 0);
  end;
end;

procedure TOM_DPFISCAL.GereExoImpositionDirecte;
begin
  SetControlVisible('GRPBOXEXONERATION', visib);

  if not visib then
  begin
    SetField('DFI_DATEDEBEXO', iDate1900);
    SetField('DFI_DATEFINEXO', iDate1900);
    SetField('DFI_EXONERATION', '');
  end;
end;

procedure TOM_DPFISCAL.GereOptionRDSup ;
begin
  actif := TDBCheckBox(GetControl('DFI_OPTIONRDSUP')).checked;
  SetControlEnabled('DFI_DATEOPTRDSUP', actif);
  SetControlEnabled('TDFI_DATEOPTRDSUP', actif);
end;

// $$$ JP 03/05/06 - FQ 11020
procedure TOM_DPFISCAL.GereOptionCA12;
begin
     visib := (GetControlText ('DFI_IMPOINDIR') <> 'E') and (GetControlText ('DFI_IMPODIR') = 'IR') and
              (((GetControlText ('DFI_REGIMFISCDIR') = 'MIC') and (TDBCheckBox (GetControl ('DFI_OPTIONRSS')).Checked)) or
               ((GetControlText ('DFI_REGIMFISCDIR') = 'RS')  and (not TDBCheckBox (GetControl ('DFI_OPTIONRISUP')).Checked)));
     SetControlVisible ('GRPCA12', visib);
     (*+LM20070315
     visib := (GetControlText ('DFI_IMPOINDIR') <> 'E') and
              ((GetControlText ('DFI_IMPODIR') = 'IR') or (GetControlText ('DFI_IMPODIR') = 'IS')) and
              (GetControlText ('DFI_REGIMFISCDIR') = 'RS'); *)
     SetControlVisible('DFI_OPTIONRISUP', InString(GetControlText ('DFI_REGIMFISCDIR'), ['RS', 'MIC']));//LM20070627
     //-LM20070315

end;

// $$$ JP 03/05/06 - FQ 11020
procedure TOM_DPFISCAL.GereCasParticulier;
begin

  (*visib := GetControlText ('DFI_IMPOINDIR') <> 'E';
  SetControlVisible ('GRPBOXCASPART', visib);
  if visib = TRUE then
  begin
    visib := GetControlText ('DFI_IMPOINDIR') = 'P';

    SetControlVisible('DFI_PRORATATVA', visib);
    SetControlVisible('DFI_DATEPRORTVA', visib);
    SetControlVisible('DFI_PRORTVAREVIS', visib);
    SetControlVisible('TDFI_PRORATATVA', visib);
    SetControlVisible('TDFI_DATEPRORTVA', visib);
    SetControlVisible('TDFI_PRORTVAREVIS', visib);
    SetControlVisible('TPC1', visib);
    SetControlVisible('TPC2', visib);
  end;
  -LM20070315 *)
end;

procedure TOM_DPFISCAL.GereOptionRSS;
begin
  actif := TDBCheckBox(GetControl('DFI_OPTIONRSS')).checked;
  SetControlEnabled('DFI_DATEOPTRSS', actif);
  SetControlEnabled('TDFI_DATEOPTRSS', actif);

  // $$$ JP 02/05/06 - FQ 11020
  GereOptionCA12
end;

procedure TOM_DPFISCAL.GereOptionReport;
begin
  actif := TDBCheckBox(GetControl('DFI_OPTIONREPORT')).checked;
  SetControlEnabled('DFI_DATEOPTREPORT', actif);
  SetControlEnabled('TDFI_DATEOPTREPORT', actif);
end;

{procedure TOM_DPFISCAL.GereOptionRISup;
var regfisc: String;
begin
  regfisc := GetControlText('DFI_REGIMFISCDIR');
  visib := (regfisc='RS') and (Not TDBCheckBox(GetControl('DFI_OPTIONRISUP')).checked);
  SetControlVisible('GRPCA12', visib);
end;}

procedure TOM_DPFISCAL.InitOptions;
begin
  SetField('DFI_OPTIONRISUP', '-');
  SetField('DFI_OPTIONRSS', '-');
  SetField('DFI_DATEOPTRDSUP', idate1900);
  SetField('DFI_DATEOPTRSS', idate1900);
end;

////////// EVENEMENTS CLICK ///////////

procedure TOM_DPFISCAL.DFI_IMPODIR_OnClick(Sender: TObject);
begin
  if (GetControlText('DFI_IMPODIR') <> GetField('DFI_IMPODIR')) then
    if Not ChangeImpoDir then
      SetControlText('DFI_IMPODIR', GetField('DFI_IMPODIR'));

  GereImpoDir (sender);
end;

procedure TOM_DPFISCAL.DFI_REGIMFISCDIR_OnClick(Sender: TObject);
begin
  if (GetControlText('DFI_REGIMFISCDIR') <> GetField('DFI_REGIMFISCDIR')) then
    if Not ChangeImpoDir then
      SetcontrolText('DFI_REGIMFISCDIR', GetField('DFI_REGIMFISCDIR'));

  SetControlText('TVAREGIMPOT', rechDom('DPREGIMFISCDIR', getControltext('DFI_REGIMFISCDIR'), false) ) ;  //réplication d'info

  GereImpoDir (sender);
end;

procedure TOM_DPFISCAL.DFI_OPTIONAUTID_OnClick(Sender: TObject);
begin
  GereImpoDir (sender);
end;

procedure TOM_DPFISCAL.DFI_OPTIONRSS_OnClick(Sender: TObject);
begin
  GereOptionRSS;
end;

procedure TOM_DPFISCAL.DFI_EXONERE_OnClick(Sender: TObject);
begin
  GereExoneration;
end;

procedure TOM_DPFISCAL.DFI_INTEGRAFISC_OnClick(Sender: TObject);
begin
  // si on décoche intégration fiscale alors qu'on a des liens sté mère
  if (Not TDBCheckBox(GetControl('DFI_INTEGRAFISC')).Checked)
   and (GetField('DFI_INTEGRAFISC')='X') then
    if Not ChangeIntegraFisc then
      // modif refusée, on recoche
      TDBCheckBox(GetControl('DFI_INTEGRAFISC')).Checked := True;

  GereIntegraFisc;
end;

procedure TOM_DPFISCAL.DFI_OPTIONREPORT_OnClick(Sender: TObject);
begin
  GereOptionReport;
end;

procedure TOM_DPFISCAL.DFI_OPTIONRDSUP_OnClick(Sender: TObject);
begin
  GereOptionRDSup;
  // #### bizarre ou incomplet
  if (GetControlText('DFI_REGIMFISCDIR')='RS')
   and (TDBCheckBox(GetControl('DFI_OPTIONRDSUP')).checked) then
    InitOptions;

  GereImpoDir (sender); // #### ????
end;

procedure TOM_DPFISCAL.DFI_OPTIONRISUP_OnClick(Sender: TObject);
begin
     GereOptionCA12;
     //GereOptionRISup;
end;

procedure TOM_DPFISCAL.DFI_EXIGIBILITE_OnClick(Sender: TObject);
begin
  GereExigibilite;
end;

procedure TOM_DPFISCAL.DFI_OPTIONEXIG_OnClick(Sender: TObject);
begin
  GereExigibilite;
end;

procedure TOM_DPFISCAL.DFI_DECLARATION_OnClick(Sender: TObject);
begin
  GereDeclaration;
end;

procedure TOM_DPFISCAL.DFI_IMPOINDIR_OnClick(Sender: TObject);
begin
  // #### purger zones dépendantes
//  if (GetControlText('DFI_IMPOINDIR')<> GetField('DFI_IMPOINDIR')) then
//    InitImpoIndir();
  GereImpoIndir;
end;

procedure TOM_DPFISCAL.DFI_AUTRESTAXES_OnClick(Sender: TObject);
begin
  GereAutresTaxes;
end;

procedure TOM_DPFISCAL.DFI_EXONERETP_OnClick(Sender: TObject);
begin
  GereExonereTp;
end;

procedure TOM_DPFISCAL.DFI_DEGREVTREDUC_OnClick(Sender: TObject);
begin
  GereDegrevmtReduc;
end;

procedure TOM_DPFISCAL.DFI_ADHEREOGA_OnClick(Sender: TObject);
begin
  GereAdhesionOGA;
end;

(*LM20070315
procedure TOM_DPFISCAL.DFI_DEMATERIATDFC_OnClick(Sender: TObject);
begin
  GereTDFC;
end;
*)

procedure TOM_DPFISCAL.DFI_REGIMFUSION_OnClick(Sender: TObject);
begin
  GereRegFusion;
end;

procedure TOM_DPFISCAL.DFI_DROITSAPPORT_OnClick(Sender: TObject);
begin
  GereDroitsApport;
end;

procedure TOM_DPFISCAL.BINTEGRATION_OnClick(Sender: TObject);
begin
  // On commence par enregistrer car les liens peuvent vouloir exploiter les infos de la fiche !
//  if not TFFiche(Ecran).Bouge(TNavigateBtn(NbPost)) then exit; // FQ 11588
  if DS.State<>dsBrowse then
  begin
    if TFFiche(Ecran).EnregOk then
      TFFiche(Ecran).BValider.Click
    else
      exit;
  end;

  if TDBCheckBox(GetControl('DFI_TETEGROUPE')).checked then
    // affiche la liste des filiales, car on est tête de groupe
    // AGLLanceFiche('DP','LIENTETEGROUPE','ANL_GUIDPERDOS='+GetGuidPer+';ANL_FONCTION=FIL','',GetGuidPer+';FIL')
    AGLLanceFiche('DP','LIENTETEGROUPE','ANL_GUIDPERDOS='+GetGuidPer+';ANL_FONCTION=FIL;ANL_GRPFISCAL=X','',GetGuidPer+';FIL;X')
  else
    // affiche la liste des stés mère, mais normalement jamais ce cas,
    // car BINTEGRATION désactivé quand on n'est pas tête de groupe
    AGLLanceFiche('DP','LIENINTERVENANT','ANL_GUIDPERDOS='+GetGuidPer+';ANL_FONCTION=TIF','',GetGuidPer+';TIF');
end;

procedure TOM_DPFISCAL.BLIENOGA_OnClick(Sender: TObject);
begin
  // On commence par enregistrer car les liens peuvent vouloir exploiter les infos de la fiche !
//  if not TFFiche(Ecran).Bouge(TNavigateBtn(nbPost)) then exit;
  if DS.State<>dsBrowse then
  begin
    if TFFiche(Ecran).EnregOk then
      TFFiche(Ecran).BValider.Click
    else
      exit;
  end;

  AGLLanceFiche('DP','LIENINTERCGE','ANL_GUIDPERDOS='+GetGuidPer+';ANL_FONCTION=OGA','',GetGuidPer+';OGA');
  // Récupération text TADHESION
  SetControlText('TADHESION', RecupNomAnnuLien (GetGuidPer, 'OGA'));
end;

function TOM_DPFISCAL.IsRDIValide: Boolean;
// recherche un RDI valide par rapport à l'exercice en cours
begin
  Result := ExisteSQL('SELECT 1 FROM ANNULIEN WHERE ANL_GUIDPERDOS="'
      + GetGuidPer + '" AND (ANL_TYPEPER="RDI" or ANL_TYPEPER="SIE")');

  // #### réflexion avec CJ non terminée
// VH^.EnCours.Deb et .Fin => fourchette de l'ex. en cours en TDateTime
{  Result := ExisteSQL('SELECT 1 FROM ANNULIEN WHERE ANL_GUIDPER="' + GetField('DFI_GUIDPER') + '" ' +
  'AND ANL_TYPEPER="RDI" AND ('
   + ' (ANL_EFDEB="01/01/1900" AND ANL_EFFIN="01/01/1900") '
   + 'OR (ANL_EFDEB >= ??? AND ANL_EFFIN ...=) '
}
end;

function TOM_DPFISCAL.GetGuidPer: String;
begin
  Result := GetField('DFI_GUIDPER');
end;


procedure TOM_DPFISCAL.OnChangeCA12(Sender: TObject);

begin
  if bCharge then exit ; // on ne fait rien sur les données pendant le chgmt d'un enreg
  ModeEdition(DS);

  if TRadioButton(GetControl('cbCA12DECALE')).Checked then SetField('DFI_TYPECA12', '12E')
                                                      else SetField('DFI_TYPECA12', '12')  ;//LM20070315

  (* //LM20070315
  if TRadioButton(GetControl('RBANNEE')).Checked then
    SetField('DFI_TYPECA12', '12') // à l'année civile
  else if TRadioButton(GetControl('RBEXERCICE')).Checked then
    SetField('DFI_TYPECA12', '12E'); // à l'exercice fiscal
  *)
end;


function TOM_DPFISCAL.ChangeImpoDir: Boolean;
var integra: Boolean;
    regimp, regfisc: String;
begin
  Result := True; // chgt de régime autorisé

  regimp := GetControlText('DFI_IMPODIR');
  regfisc := GetControlText('DFI_REGIMFISCDIR');

  // si on passe à un régime sans intégration fiscale, faut décocher
  integra := ((regimp='IR') and (regfisc='RN') and (TDBCheckBox(GetControl('DFI_OPTIONAUTID')).checked))
          or ((regimp='IS') and (regfisc='RN'));
  if (Not integra) and (TDBCheckBox(GetControl('DFI_INTEGRAFISC')).Checked) then
    begin
    // cela doit déclencher son OnClick qui gère les cqs du décochage
    TDBCheckBox(GetControl('DFI_INTEGRAFISC')).Checked := False;
    // si la modif a été refusée
    if TDBCheckBox(GetControl('DFI_INTEGRAFISC')).Checked then
      Result := False
    else
      InitImpoDir();
    end
  else
    InitImpoDir();
end;

function TOM_DPFISCAL.ChangeIntegraFisc: Boolean;
var Msg : String;
begin
  Result := True; // chgt d'intégration fiscale autorisé

  // si on décoche intégration alors qu'on a des liens sté mère ou qu'on est tête de grp
  if (Not TDBCheckBox(GetControl('DFI_INTEGRAFISC')).Checked) then
    begin
    // tête de groupe ne doit pas rester coché => on décoche
    if TDBCheckBox(GetControl('DFI_TETEGROUPE')).Checked then
      begin
      SetControlChecked('DFI_TETEGROUPE', False);
      // décoche la case "Intégration fiscale" sur tous les liens filiales
      ExecuteSQL('update ANNULIEN set ANL_GRPFISCAL="-" where ANL_GUIDPERDOS = "'+GetGuidPer
       + '" and ANL_FONCTION="FIL"');
      end
    // pas tête de groupe, mais des liens existent vers une tête de gr
    else if GetControlText('TINTEGRATION')<>'' then
      begin
      Msg := 'Cette opération entraînera la suppression de tous les liens société mère. Voulez-vous supprimer ?';
      if PGIAsk(Msg, TitreHalley)=mrYes then
        begin
        //LM20070425 ExecuteSQL('delete from ANNULIEN where ANL_CODEPERDOS = "'
        ExecuteSQL('delete from ANNULIEN where ANL_GUIDPERDOS = "' +GetGuidPer+ '" and ANL_FONCTION="TIF"');
        {+GHA - 12/2007}
        // Mise de la table annulien sur le lien filiale (FIL)
        ExecuteSQL('Update ANNULIEN Set ANL_GRPFISCAL ="-",'+
                   'ANL_TXDETDIRECT = 0, ANL_TXDETINDIRECT =0 '+
                   'where ANL_GUIDPER = "' +GetGuidPer+ '" and ANL_FONCTION="FIL"');
        {-GHA - 12/2007}
        SetControlText ('TINTEGRATION', '');
        end
      else
        Result := False;
      end;
    end;
    if vh_dp.Group then //LMO20060901
    begin
      SetControlChecked('CBSOCINTEGREE', not (getControlText('DFI_TETEGROUPE')='X'));
      SetControlEnabled('GBFILIALE', (getControlText('CBSOCINTEGREE')='X')) ;
    end ;
end;


procedure TOM_DPFISCAL.DFI_REGIMEINSPECT_OnChange(Sender: TObject);
var txt: String;
begin
   inherited;
   if GetControlText('DFI_REGIMEINSPECT') = '' then exit;
   txt := RechDOM('LTVATYPE', GetControlText('DFI_REGIMEINSPECT'), False);
   if UpperCase(txt) = 'ERROR' then
   begin
      ErreurChamp('DFI_REGIMEINSPECT', 'TSIMPOSITIONINDIRECTE', Self, 'Le régime d''inspection est incorrect') ;
//      txt := '';
   end;
//   SetControlText('DFI_REGIMEINSPECT', txt);

  txt := getControlText('DFI_REGIMEINSPECT');
  if InString(txt, ['ET', 'RT', 'AET', 'EST', 'RST', 'EOT', 'ST', 'AST', 'ADT']) then
    setControltext('DFI_PERIODIIMPIND', 'TRI')
  else if InString(txt, ['EM', 'RM', 'ESM', 'RSM', 'EOM']) then
    setControltext('DFI_PERIODIIMPIND', 'MEN' ) ; //LM20070315

end;

// BM : 22/06/2004
procedure TOM_DPFISCAL.DFI_TAXEPROF_OnClick(Sender : TObject);
var bReinit : Boolean;
begin
  EltTp (GetControlText('DFI_TAXEPROF')='X') ; //LM20070511

   if (DS.State = dsBrowse) then
      Exit;

   // Si on a décoché
   if not TDBCheckBox(GetControl('DFI_TAXEPROF')).Checked then
   begin
      // en création, on agit sans poser de question
      if DS.State = dsInsert then
        bReinit := True
      // sinon on demande
      else if PGIAsk('Vous avez décoché l''option "Imposable à la TAXE PROFESSIONNELLE".' + #13#10 +
              ' Les données concernant la Taxe professionnelle vont être réinitialisées.' + #13#10 +
              ' Confirmez-vous?', Ecran.Name) = mrYes then
          bReinit := True
      else
          bReinit := False;
      if bReinit then
      begin
          SetField('DFI_ACPTEJUIN', '-');
          SetField('DFI_DECLA1003R', '-');
          SetField('DPTYPETAXEPRO', '');
          SetField('DFI_ACTIVTAXEPRO', '');
          SetField('DFI_COTISMIN', '-');
          SetField('DFI_ALLEGETRANS', '-');
          SetField('DFI_ABATTEFIXE', '-');
          SetField('DFI_CAMIONSCARS', '-');
          SetField('DFI_DEGREVTREDUC', '-');
          SetField('DFI_MTTDEGREVTRED', '');
          SetField('DFI_EXONERETP', '-');
          SetField('DFI_DATEDEBEXOTP', iDate1900);
          SetField('DFI_DATEFINEXOTP', idate1900);
          SetField('DFI_EXONERATIONTP', '');
      end
      else
      begin
         SetField('DFI_TAXEPROF', 'X');
      end;
   end;
   //SetControlVisible('TSTAXEPROF', TDBCheckBox(GetControl('DFI_TAXEPROF')).Checked);
end;

procedure TOM_DPFISCAL.InfoSocMere (guidper : string) ; //LMO20060901
var q : tquery ;
begin
  q:=opensql('select ANL_EFDEB, ANL_EFFIN, ANL_TXDETTOTAL from ANNULIEN '+
              'where ANL_GUIDPERDOS = "' + getField('DFI_GUIDPER') + '"', true) ;
  if not q.eof then
  begin
    setControlText('EDDTINGP', DateToStr(q.FindField('ANL_EFDEB').asDateTime));
    setControlText('EDDTOUTGP', DateTostr(q.FindField('ANL_EFFIN').asDateTime));
    setControlText('ANL_TXDETTOTAL', FloatToStrF(q.FindField('ANL_TXDETTOTAL').asfloat,ffNumber, 15, 2));
  end ;
  ferme(q);
end ;

{LM20070315
procedure TOM_DPFISCAL.SelectAnnuaire (Sender: TObject) ; //LMO20060901
var ann_ : tob ;
    q : tquery ;
    champ, st: string ;
begin
  if sender=nil then exit ;
  champ:=TControl(sender).Name ;
  if champ='DFI_2065SIGNEPARID' then
  begin
    ann_:=tob.create('ANNUAIRE', nil, -1) ;
    st := getControltext(champ) ;
    if st='' then st:='1=2' else st :=  'ANN_GUIDPER="' + st + '"';
    q:=openSql('select ANN_CV, ANN_NOM1, ANN_NOM2, ANN_TEL1, ANN_EMAIL from ANNUAIRE where '+ st, true) ;
    ann_.SelectDB('', q) ;
    ferme(q) ;
    ann_.AddChampSup('ANN_NOM', false);
    ann_.p('ANN_NOM', ann_.g('ANN_NOM1') + ' '  + ann_.g('ANN_NOM2'));
    ann_.PutEcran(Ecran) ;
    ann_.free;
  end;
end ;
}

{
procedure TOM_DPFISCAL.SelectGED (Sender: TObject) ; //LMO20060901
var q : tquery ;
    st, champ : string ;
begin
  if sender=nil then exit ;

  champ:=TControl(sender).Name ;
  if champ='DFI_DOCOPTIONGPGED' then
  begin
    st := getControltext(champ) ;
    if st='' then st:='1=2' else st :=  'YDO_DOCGUID="' + st + '"';
    q:=openSql('select YDO_LIBELLEDOC from YDOCUMENTS where '+ st, true) ;
    if not q.eof then st := q.FindField('YDO_LIBELLEDOC').asString
                 else st := '';
    ferme(q) ;
    setControltext(champ+'_', st) ;//Hypothèse : le controle qui affiche le libellé est = nom du db + _
  end;
end ;

procedure TOM_DPFISCAL.RAZLookup (Sender: TObject);//LMO20060901
var nom : string ;
begin
  if sender=nil then exit ;
  nom := TControl(sender).name ;
  nom := copy(nom,4,length(nom)-3) ;//hypothèse : nom de la gomme = RAZ + nom du champ(=nom zone)
  setControltext(nom,'') ;
  setField(nom,#0);
end ;

procedure TOM_DPFISCAL.showGED (Sender: TObject);//LMO20060901
var nom : string ;
begin
  if sender=nil then exit ;
  nom := TControl(sender).name ;
  nom := copy(nom,4,length(nom)-3) ;//hypothèse : nom de l'aperçu = GED + nom du champ(=nom zone)
  showGedViewer(getControltext(nom), true) ;
end ;
}
procedure TOM_DPFISCAL.OnClose ;
begin
  inherited ;
  // EvDecla.Free ; //LM20070308  MD20070705 assuré par le owner du TComponent
  FreeAndNil(ANL_Lst) ; //LM20070315
end ;

procedure TOM_DPFISCAL.DFI_REGLEFISCclick (Sender: TObject);//LM20070308
begin
  //LM20070611 fq11455 SetControlText('TVAREGIMPOT', rechDom('DPREGLEFISC', getControltext('DFI_REGLEFISC'), false) ) ;  //réplication d'info
  LibelleOGA ;
  GereImpoDir(sender);
end ;

procedure TOM_DPFISCAL.LibelleOGA ;
var catFisc : string ;
begin
  catFisc := getControlText('DFI_REGLEFISC');
  if catFisc='BNC' then
    setcontrolProperty('DFI_ADHEREOGA', 'caption', 'Adhésion à une AGA')
  else if catFisc='BIP' then
    setcontrolProperty('DFI_ADHEREOGA', 'caption', 'Adhésion à un CGA')
  else if catFisc='BA' then
    setcontrolProperty('DFI_ADHEREOGA', 'caption', 'Adhésion à un CGAA')
  else
    setcontrolProperty('DFI_ADHEREOGA', 'caption', 'Adhésion à un organisme de gestion agréée') ; //LM20070315

end ;


procedure TOM_DPFISCAL.AfficheINDouSAL (sender:TObject) ;  //LM20070315
var cas, nomRef : string ;
begin
  if sender=nil then exit ;
  nomRef := gtfs(Tcontrol(sender).Name, '_', 1 ) ;
  if pos('ValComboBox' , Tcontrol(sender).ClassName)=0 then exit ;
  cas:=THValcombobox(sender).Value ;

  //setcontrolVisible(Nomref + '_CIVI',  cas='SAL' ) ; setcontrolVisible('T' + Nomref + '_CIVI',  cas='SAL' ) ;
  setcontrolVisible(Nomref + '_PRENOM',cas='SAL' ) ; setcontrolVisible('T' + Nomref + '_PRENOM',cas='SAL' ) ;
  setcontrolVisible(Nomref + '_NOM',   cas='SAL' ) ; setcontrolVisible('T' + Nomref + '_NOM',   cas='SAL' ) ;
  setcontrolVisible(Nomref + '_BSAL',  cas='SAL' ) ;

  setcontrolVisible(Nomref , InString(cas,['IND']) ) ;
  setcontrolVisible('T' + Nomref , InString(cas,['IND']) ) ;
  setcontrolVisible(Nomref + '_BLIEN', InString(cas,['IND']) ) ;

  setcontrolVisible(Nomref + '_DEL',  cas<>'' ) ;

  // Pas de droits en modif
  if Not JaiLeDroitConceptBureau(ccModifAnnuaireOuLiens) or Not JaiLeDroitConceptBureau(ccVoirLesLiens) then
    begin
    setcontrolVisible(Nomref + '_BSAL',  False ) ;
    setcontrolVisible(Nomref + '_BLIEN', False ) ;
    setcontrolVisible(Nomref + '_DEL',  False ) ;
    end
  // ou bien droits plus détaillés
  else
    begin
    // Pas de création
    if Not JaiLeDroitConceptBureau(ccCreatAnnuaireOuLiens) then
      begin
      setcontrolVisible(Nomref + '_BSAL',  False ) ;
      setcontrolVisible(Nomref + '_BLIEN', False ) ;
      end;
    // Pas de suppression
    if Not JaiLeDroitConceptBureau(ccSupprAnnuaireOuLiens) then
      setcontrolVisible(Nomref + '_DEL',  False ) ;
    //
    end;

end ;

procedure TOM_DPFISCAL.SelectSalarie (sender:TObject) ;  //LM20070315
var codeSal, nomRef : string ;
begin
  if sender = nil then exit ;
  if ChoisirParTablette ('PGSALARIE', '', codeSal) then
  begin
    nomRef := gtfs(TControl(sender).name, '_', 1) ;
    AfficheSalarie(codesal, nomRef) ;
    Signataire_IsModif (sender);
  end ;

end ;

procedure TOM_DPFISCAL.SelectAnnuaire(sender:TObject) ;  //LM20070315
var nomRef, guidper, sParam1_l : string ;
begin
  // A faire avec ANNULIEN_TRI ensuite
  if sender = nil then exit ;

  nomRef := gtfs(TControl(sender).name, '_', 1) ;
  if nomRef = 'CSL' then
  	 sParam1_l := 'ANN_TYPEPER=CFIS'
  else if nomRef = 'EXP' then
    sParam1_l := 'ANN_TYPEPER=EXPCPT'
  else
    sParam1_l := '';

  guidper := LancerAnnuSel(sParam1_l, '', '');
  if (guidper <> '') and (guidper <> GetControlText(nomRef + '_GUIDPER')) then
//  if ChoisirParTablette ('ANNUAIRE', '', guidper) then
  begin
    setControlText(nomRef + '_GUIDPER', guidper ) ;
    setControlText(nomRef, rechDom('ANNUAIRE', guidper, false) ) ;
    Signataire_IsModif (sender);
  end ;
end ;


procedure TOM_DPFISCAL.AfficheSalarie (codeSal, prefixe : string) ;  //LM20070315
var nom, prenom : string ;
    q:TQuery ;
begin
  if codeSal<>'' then
  begin
    q:=openSql('select PSA_PRENOM, PSA_LIBELLE from SALARIES ' +
               'where PSA_SALARIE = "' + codeSal+ '"', true) ;
    if not q.Eof then
    begin
      prenom := q.findField('PSA_PRENOM').asstring ;
      nom := q.findField('PSA_LIBELLE').asstring ;
    end ;
    ferme(q) ;
  end ;
  setControlText(prefixe+'_NOADMIN', codeSal) ;
  setControlText(prefixe+'_PRENOM', prenom) ;
  setControlText(prefixe+'_NOM', nom) ;
end ;


procedure TOM_DPFISCAL.EffaceINDouSAL(sender:TObject) ;  //LM20070315
var nomRef : string ;
begin
  // A faire avec ANNULIEN_TRI ensuite
  if sender = nil then exit ;

  nomRef := gtfs(TControl(sender).name, '_', 1) ;
  setControlText(nomRef, '' ) ;
  setControlText(nomRef + '_GUIDPER', '' ) ;
  setControlText(nomRef + '_NOADMIN', '' ) ;
  setControlText(nomRef + '_PRENOM', '' ) ;
  setControlText(nomRef + '_NOM', '' ) ;

  Signataire_IsModif (sender);

end ;


procedure TOM_DPFISCAL.Signataire_Init ;//LM20070315
var sign, s, st : string ;
//    q : TQuery ;
    i : integer ;
    Lst_, ann_, ann_Dupli : Tob ; //LM20071105

    procedure Affiche_ (prefixe : string);
    var ANL_, ANN_ : Tob ;
        st : string ;
    begin
      ANL_ := ANL_Lst.FindFirst(['ANL_DECLARATION'], [copy(prefixe,1, 3)], false) ;
      if ANL_<> nil then
      begin
        setcontroltext(prefixe + '_GUIDPER' , ANL_.getString('ANL_GUIDPER') );
        //+LM20070820
        if ANL_.Detail.count>0 then ANN_ := ANL_.Detail[0] ;
        if ANL_.getString('ANL_INFO')<>'' then st := 'SAL' else st := 'IND' ;
        setcontroltext(prefixe + '_QUALI'   , st );
        setcontroltext(prefixe + '_NOADMIN' , ANL_.getString('ANL_INFO') );
        // setcontroltext(prefixe + '_PRENOM'  , ANN_.getString('ANN_NOM2') );
        if st='SAL' then
        begin
          s := ANL_.getString('ANL_NOTELIEN') ;
          setcontroltext(prefixe + '_PRENOM'     , gtfs(s,'^',1) ) ;
          setcontroltext(prefixe + '_NOM'     , gtfs(s,'^',2) ) ;
        end
        else if ANN_<> nil then //LM20071105
          setcontroltext(prefixe, ANN_.getString('ANN_NOM1') );
        //-LM20070820

      end;

    end ;

begin
  //Lien de type signataire
  s:= 'select ANL_DECLARATION, ANL_GUIDPER, ANL_NOMPER, ANL_INFO, ANL_NOTELIEN, ANL_NOORDRE ' +
      'from ANNULIEN ' +
      'where ANL_GUIDPERDOS = "' + GetGuidPer + '" and ANL_DECLARATION <> "" ' ;
//  q:=openSql(s, true) ;
  ANL_Lst.LoadDetailDBFromSql('ANNULIEN', s) ;
//  ferme(q) ;

  //on y associe la fiche annuaire correspondante
  if ANL_Lst.Detail.count>0 then
  begin
    st:='' ;
    for i:= 0 to ANL_Lst.Detail.count-1 do
    begin
      s:=ANL_Lst.detail[i].G('ANL_GUIDPER') ;
      if st<>'' then st := st + ',' ;
      st:=st + '"' + s + '"' ;
    end ;
    s:= 'select ANN_GUIDPER, ANN_NOM1, ANN_NOM2, ANN_NOADMIN ' + //Plus?? nom1=nom, nom2=prénom
        'from ANNUAIRE  ' +
        'where ANN_GUIDPER in (' + st + ') ' ;
//    q:=openSql(s, true) ;
    Lst_:=tob.create('ANNUAIRE', nil, -1) ;
//    Lst_.LoadDetailDB('ANNUAIRE', '','', q, true) ;
    Lst_.LoadDetailDBFromSQL('ANNUAIRE', s) ;
//    ferme(q) ;

    for i:= 0 to ANL_Lst.Detail.count-1 do //au moins 1 détail (ann_) par lien (anl_)
    begin
      s:=ANL_Lst.detail[i].GetString('ANL_GUIDPER') ;
      ann_ := Lst_.FindFirst(['ANN_GUIDPER'], [s], true) ;
      //LM20071105  if ann_<> nil then ann_.ChangeParent(ANL_Lst.Detail[i], -1) ;
      if ann_<> nil then //LM20071105
      begin
        ann_Dupli := tob.Create('ANNUAIRE', nil, -1) ;
        ann_Dupli.dupliquer(ann_, false, true);
        ann_Dupli.changeParent(ANL_Lst.Detail[i],-1) ;
      end ;
      if st<>'' then st := st + ',' ;
      st:=st + '"' + s + '"' ;
    end ;

    (*LM20071105 A supprimer
    if lst_.detail.count >0 then
    begin //Ca devrait être vide => pb

    end ;*)
    Lst_.Free ;

  end ;

  //+LM20070315
  sign:='*' ; i:=1 ;
  while sign<>'' do
  begin
    sign := trim(gtfs(signataire,';',i)) ;
    if sign='' then break ;
    Affiche_(sign) ;
    AfficheINDouSAL (getControl(sign+'_QUALI')) ;
    inc(i);
  end ;
  //-LM20070315

end ;

procedure TOM_DPFISCAL.Signataire_Sauve (prefixe : string);
var anl_, anlDb_ : Tob ;
    NoAdmin, guid, st : string ;
    majVerAFaire : boolean ;
    NoOrdre : integer ;
begin
{
Les signataires sont stockés dans ANNULIEN : on stocke une référence à la fiche annuaire qui comporte l'information
Il sont identifiés :
- le fct standard des liens ANL_GUIDPER, et ANL_GUIDPERDOS
- ANL_FONCTION = 'SGN'
- ANL_DECLARATION qui comporte le type de déclaration - cf const Signataire
- ANL_NOORDRE, sur la base de la constante Signataire, pour avoir un critère discriminant sur l'index d'ANNULIEN
}
  NoAdmin := getControlText(prefixe+'_NOADMIN') ;// Distingue les cas Salarié/Indépendant  - si vide alors IND sinon SAL
  if getControltext(prefixe + '_QUALI')  <> 'SAL' then NoAdmin:=''  //LM20070611
  else if ((getControltext(prefixe + '_QUALI')  = 'SAL') and (NoAdmin='')) then NoAdmin:='0';

  // guidper de l'intervenant (ne pas confondre avec GetGuidPer qui renvoie le guidper du
  guid := getControlText(prefixe+'_GUIDPER') ;
  if guid='' then guid:=AglGetGuid() ;

  anl_ := ANL_Lst.FindFirst(['ANL_DECLARATION'], [copy(prefixe,1,3)], true) ;
  if (anl_=nil)  then
  begin
    anl_:=tob.create('ANNULIEN', ANL_Lst, -1) ;
    //ann_:=tob.create('ANNUAIRE', anl_, -1) ;
  end;
  //else
  //  ann_ :=anl_.detail[0] ; //si info lien alors tjs info annuaire  : 1 anl_ => 1 ann_

  (* LM20070622 pas de fiche annuaire
  if (getControltext(prefixe + '_QUALI')  = 'SAL')
   and (getControltext(prefixe + '_NOM')<>'') then //"EAI" fiche psa Salarie>ann Annuaire quand le nom est saisi
  begin
    ann_.p('ANN_CODEPER', -2);
    ann_.p('ANN_GUIDPER', guid);
    ann_.p('ANN_NOADMIN', NoAdmin);
    ann_.p('ANN_NOM1', getControltext(prefixe + '_NOM') );
    ann_.p('ANN_NOM2', getControltext(prefixe + '_PRENOM'));
    ann_.InsertOrUpdateDB ;//LM>MD : suffisant comme champs renseignés?
  end ;
  *)

  if ((getControltext(prefixe + '_QUALI')  = 'IND') and (getControlText(prefixe+'_GUIDPER')<>''))
   or ((getControltext(prefixe + '_QUALI')  = 'SAL') and (getControltext(prefixe + '_NOM')<>'') ) then
  begin
    anl_.p('ANL_GUIDPER', guid);
    anl_.p('ANL_DECLARATION', prefixe);
    anl_.p('ANL_GUIDPERDOS', GetGuidPer );
    anl_.p('ANL_TYPEDOS', 'DP');//LM>MD ?

    // de uTomYYAnnulien
    NoOrdre := Postfs (signataire, ';', prefixe)+1 ; //LM20071105 commence à 2 pour gérer le pseudo majver
    anl_.p('ANL_NOORDRE', NoOrdre);//LM20071105
    anl_.p('ANL_FONCTION', 'SGN');  //LM20071105
    anl_.p('ANL_CODEDOS', GetDosjuri (GetGuidPer));
    anl_.p('ANL_AFFICHE', 'Signataire déclaration');
    //+LM20070820
    anl_.p('ANL_INFO', NoAdmin);
    if getControltext(prefixe + '_QUALI')  = 'IND' then
      anl_.p('ANL_NOMPER', GetNomPer(guid) )
    else
    begin
      anl_.p('ANL_NOMPER', getControltext(prefixe + '_NOM') );
      anl_.p('ANL_NOTELIEN', getControltext(prefixe + '_PRENOM') + '^' + getControltext(prefixe + '_NOM') );
    end ;
    //-LM20070820

    //index : ANL_GUIDPERDOS,ANL_TYPEDOS,ANL_NOORDRE,ANL_FONCTION,ANL_GUIDPER
    anlDb_:=tob.create('ANNULIEN', nil, -1) ;
    anlDb_.Dupliquer(anl_, false, true );
    //anlDb_.InsertOrUpdateDB ;
    if presenceComplexe('ANNULIEN',
                        ['ANL_GUIDPERDOS', 'ANL_TYPEDOS', 'ANL_DECLARATION', 'ANL_FONCTION', 'ANL_NOORDRE' ],
                        ['=','=','=', '=', '='],
                        [anl_.g('ANL_GUIDPERDOS'), 'DP', Copy(Prefixe,1,3), 'SGN',  intTostr(NoOrdre)],
                        ['S', 'S', 'S', 'S', 'I']) then     //LM20071105
      ExecuteSql('delete ANNULIEN '+
                 'where ANL_DECLARATION="' + copy(prefixe,1,3) + '" ' +
                 'and ANL_GUIDPERDOS = "' + GetGuidPer + '"') ;
    anlDb_.insertDb (nil) ;
    anlDb_.free ;


    if inString(prefixe, ['EXP','CSL']) then
    begin
      anl_.p('ANL_FONCTION', 'DIV');
      if prefixe = 'EXP' then anl_.p('ANL_TYPEPER', 'EXPCPT' )
      else if prefixe = 'CSL' then anl_.p('ANL_TYPEPER', 'CFIS' );
      anl_.p('ANL_DECLARATION', '');
      anl_.p('ANL_AFFICHE', RechDom('JUTYPEFONCTAFF', anl_.g('ANL_TYPEPER'), false));
      anl_.InsertOrUpdateDB;
    end ;
  end
  else
  begin //Les liens sont désactivés => on supprime
    if inString(prefixe, ['EXP','CSL']) then //Cas du lien normal comptable ou conseil - anl_Fonction=DIV et anl_typeper=...
    begin
      if prefixe = 'EXP' then st:='EXPCPT'
      else if prefixe = 'CSL' then st:='CFIS' ;
      st:= ' or (ANL_TYPEPER ="' + st + '" and ANL_FONCTION="DIV") ' ;
    end
    else
      st:= '' ;
    ExecuteSql('delete ANNULIEN '+
               'where (ANL_DECLARATION="' + copy(prefixe,1,3) + '"  ' + st + ' ) ' +
               'and ANL_GUIDPERDOS = "' + GetGuidPer + '"') ;//LM20070622
  end ;
end ;


procedure TOM_DPFISCAL.Signataire_IsModif(sender:TObject);//LM20070315
var n,p,s : string ;
    modif : boolean ;
    anl_ : tob ;
begin
  if bCharge or (sender=nil) then exit ;
  if ANL_Lst=Nil then exit; // FQ 11711

  n:=TControl(sender).Name;
  p:=gtfs(n, '_', 1) ;if p='' then p:=n ;
  s:=gtfs(n, '_', 2) ;


  anl_ := ANL_Lst.FindFirst(['ANL_DECLARATION'], [copy(p,1,3)], true) ;
  if (anl_=nil) then
    modif:=(getcontroltext(n)<>'')  //Il n'y avait rien de renseigné
  else if (s='') or (s='BLIEN') then//Zone Désignation=> type INDépendant ou Select annuaire
    modif := getcontroltext(n)<>anl_.GetString('ANL_GUIDPER')
  else if (s='QUALI') then          //Chgt IND<>SALarié
    modif := ( (anl_.Detail.count=0) and (getcontroltext(n)<>'') )
              or ((anl_.Detail.count>0) and (getcontroltext(p + '_NOADMIN')<> anl_.detail[0].GetString('ANL_NOADMIN'))) //LM20070820
  else if (s='NOM') or (s='BSAL') then            //Chgt nom salarié
    modif := getcontroltext(n) <> gtfs(anl_.getString('ANL_NOTELIEN') ,'^',2) //LM20070820
  else if (s='PRENOM') then         //Chgt prénom salarié
    modif := getcontroltext(n) <> gtfs(anl_.getString('ANL_NOTELIEN') ,'^',1) //LM20070820
  else if (s='DEL') then            //Efface
    modif:=true
  else
    Modif:=false ;

  if Modif then
  begin
    ForceUpdate ;
    setField('DFI_GUIDPER', getField('DFI_GUIDPER')) ;
  end ;
end;

procedure TOM_DPFISCAL.DFI_TVADECLAChange (sender:TObject) ;    //LM20070315
begin
  actif :=getControlText('DFI_TVADECLA')<>'PCE' ; //Attention si modif de tablette DPISDECLA / DPTVAEDI
  setControlEnabled('DFI_DATECONVEDITVA', actif) ;
  setControlEnabled('DFI_NOFRP', actif) ;
end ;

procedure TOM_DPFISCAL.initJourDepotTVA ; //LM20070315
var i :integer ;
    cb: THValcomboBox ;
    ctrl : TControl ;
    st,w : string ;
begin
  ctrl := getControl('DFI_JOURDECLA') ;
  if ctrl=nil then exit ;
  cb:=THValcomboBox (ctrl) ;
  st:='5,15,16,17,19,20,21,23,24' ;
  w:='*'; i:=1;
  while w<>'' do
  begin
    w:=gtfs(st,',',i);
    cb.items.add(w) ;
    cb.values.add(w) ;
    inc(i);
  end ;
end ;

procedure TOM_DPFISCAL._setField(champ : string ; v : variant) ;   //LM20070329
begin
  if getField(champ)<>v then setField(champ, v)
end ;

{
procedure TOM_DPFISCAL.LiensAnn(sender:TObject) ;//LM20070404 : en attendant la refonte des liens
var q:TQuery ;
begin
  q:=opensql('select JUR_TYPEDOS, JUR_NOORDRE, JUR_FORME from JURIDIQUE '+
             'where JUR_GUIDPERDOS = "' + sGuidPer_c + '"' ,true) ;
  if not q.eof then
    AGLLanceFiche('JUR','ANNULIEN_TRI', '', '',
                       sGuidPer_c + ';' +
                       q.findField('JUR_TYPEDOS').asString + ';' +
                       q.findField('JUR_NOORDRE').asString + ';' +
                       q.findField('JUR_FORME').asString + ';' +
                       v_pgi.NoDossier + ';' +
                       '');
  ferme(q) ;
end ;
}
procedure TOM_DPFISCAL.DFI_TAXEVEHICSOnClick (sender:TObject) ;
begin
  Application.ProcessMessages ;
  GroupBoxEnabled(Ecran, getControl('GRPBOXTVS'), 'DFI_TAXEVEHICSOC', getControlText('DFI_TAXEVEHICSOC')='X'); //LM20070511 setcontrolEnabled('DFI_TVSPAIEMENT', getControlText('DFI_TAXEVEHICSOC')='X') ;
end ;

procedure TOM_DPFISCAL.DFI_TAXESALAIRESOnClick (sender:TObject) ;
var b:boolean ;
begin
  Application.ProcessMessages ;
  b := (getControlText('DFI_IMPOINDIR') <> 'T') and (getControlText('DFI_TAXESALAIRES')='X') ;
  GroupBoxEnabled( Ecran, getControl('GRPBOXTS'), 'DFI_TAXESALAIRES', b); //LM20070511 setControlEnabled('GRPBOXTS', (GetControlText('DFI_IMPOINDIR') <> 'T') ) ;
end ;

procedure TOM_DPFISCAL.getAnn_forme  ;
var q:TQuery ;
begin
  q:=opensql('select JUR_FORME, ANN_PPPM from JURIDIQUE '+
             'left outer join ANNUAIRE on ANN_GUIDPER = JUR_GUIDPERDOS ' +
             'where JUR_GUIDPERDOS = "' + sGuidPer_c + '"' ,true) ;


  if not q.eof then
  begin
    ann_forme := q.findField('JUR_FORME').asString ;
    ann_pppm := q.findField('ANN_PPPM').asString ;
  end ;
  ferme(q) ;
end ;

function TOM_DPFISCAL.MoisClotureDecale : boolean;//LM20070511
begin
  result:=MoisCloture<>12
end ;

procedure TOM_DPFISCAL.EltTP (ok:boolean) ;//LM20070511
begin
  setcontrolEnabled ('GRPBOXTP1', ok) ;
  setcontrolEnabled ('GBOXTPSIGNATAIRE', ok) ;
  setcontrolEnabled ('GRPBOXTPEXONERATION', ok) ;
  setcontrolEnabled ('GBLIENSERVANTISSIMMO', ok) ;
  setcontrolEnabled ('DFI_EXONERETP', ok) ;
end ;


procedure TOM_DPFISCAL.BLIENS_OnClick(sender: TObject);
begin
  // On commence par enregistrer car les liens peuvent vouloir exploiter les infos de la fiche !
  // if not TFFiche(Ecran).Bouge(TNavigateBtn(nbPost)) then exit;
  if DS.State<>dsBrowse then
  begin
    if TFFiche(Ecran).EnregOk then
      TFFiche(Ecran).BValider.Click
    else
      exit;
  end;

  if GetGuidPer<>'' then
  begin
    // FQ 11768
    bModifLiensFisc := True;
    // FQ 11894
    ModeEdition(DS);
    SetField('DFI_GUIDPER', GetField('DFI_GUIDPER'));

    // AGLLanceFiche('DP','LIENINTERVENANT','ANL_GUIDPERDOS='+GetGuidPer+';ANL_FONCTION=FIS','',GetGuidPer+';FIS');
    // FQ 11995 : on passe la tom fiscale pour gestion des FRP
    LanceFicheLien('DP','LIENINTERVENANT','ANL_GUIDPERDOS='+GetGuidPer+';ANL_FONCTION=FIS','',GetGuidPer+';FIS', Self);
  end;
end;

function TOM_DPFISCAL.ongletEnCours:string;
var p : THPageControl2;
begin
  p:=ThPageControl2(getControl('Pages')) ;
  result:= p.activepage.name ;
end ;

procedure TOM_DPFISCAL.LiensDAS2(sender:TObject);
begin
  // On commence par enregistrer car les liens peuvent vouloir exploiter les infos de la fiche !
//  if not TFFiche(Ecran).Bouge(TNavigateBtn(nbPost)) then exit;
  if DS.State<>dsBrowse then
  begin
    if TFFiche(Ecran).EnregOk then
      TFFiche(Ecran).BValider.Click
    else
      exit;
  end;

  AGLLanceFiche('DP', 'LIENINTERVENANT', 'ANL_GUIDPERDOS='+GetGuidPer+';ANL_FONCTION=SOC;','',GetGuidPer+';SOC');
end ;


procedure TOM_DPFISCAL.RechercheSIEetMajFicheFiscale;
var
    Q: TQuery;
    sTypePer, sGuidPer, sNoAdhesion, sNoAdhCompl: String;
begin
  // Recherche le lien fiscal le plus récent
  // Rq: sGuidPer_c est le guidper du dossier (propriété de la classe TOM_DPFISCAL)
  Q := OpenSQL('SELECT ANL_TYPEPER, ANL_GUIDPER, ANL_NOADHESION, ANL_NODHCOMPL FROM ANNULIEN'
  + ' WHERE ANL_GUIDPERDOS="'+ sGuidper_c + '"'
  + ' AND ANL_TYPEDOS="DP" AND ANL_FONCTION="FIS"'
  // $$$ JP 24/04/06 - prendre le SIE en priorité donc DESC
  + ' AND (ANL_TYPEPER="CDI" OR ANL_TYPEPER="RDI" OR ANL_TYPEPER="SIE") ORDER BY ANL_TYPEPER DESC, ANL_NOORDRE', True);
  if not Q.Eof then
    begin
    sTypePer    := Q.FindField('ANL_TYPEPER').AsString;
    sGuidPer    := Q.FindField('ANL_GUIDPER').AsString;
    sNoAdhesion := Q.FindField('ANL_NOADHESION').AsString;
    sNoAdhCompl := Q.FindField('ANL_NODHCOMPL').AsString;
    end;
  Ferme(Q);
  // Mettre à jour le N° FRP (DFI_NOFRP) et N° de dossier (DFI_NOINSPECTION)
  MajDpfiscalFromLienSIE(sTypePer, sGuidPer_c, sGuidPer, sNoAdhesion, sNoAdhCompl) ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 20/09/2007
Modifié le ... : 20/09/2007
Description .. : Mise à jour de la fiche fiscale (si possible)
Suite ........ : à partir d'un lien de type CDI/RDI/SIE :
Suite ........ : met à jour no d'inspection, numéro FRP
Suite ........ : (voir OnUpdateRecordDP de UTOMYYAnnulien)
Mots clefs ... : 
*****************************************************************}
procedure TOM_DPFISCAL.MajDpfiscalFromLienSIE(sTypePer, sGuidPerDos, sGuidPer, sNoAdhesion, sNoAdhCompl : String);
// sTypePer    = CDI/RDI/SIE
// sGuidPerDos = guidper du dossier
// sGuidPer    = guidper de l'organisme fiscal (CDI/RDI/SIE) rattaché
//              (si plusieurs : privilégier le SIE, ou le plus récent)
// sNoAdhesion = no de dossier auprès de l'organisme = ANL_NOADHESION
// sNoAdhCompl = complément au no d'adhésion = ANL_NODHCOMPL
var
    TobDpFiscal : TOB;
    Q : TQuery;
    bIIndExo    :boolean;
    sRecette, sInspection : String;
begin
  if (sGuidPerDos='') or (sGuidPer='') then exit;

  // Récup des infos de l'organisme
  Q := OpenSQL('select ANN_NOADMIN, ANN_NOIDENTIF, ANN_COMPLTNOADMIN, ANN_FAMPER ' +
               'from ANNUAIRE ' +
               'where ANN_GUIDPER = "' + sGuidPer + '"', TRUE);
  if not Q.Eof then
  begin
     sRecette := Q.FindField('ANN_NOADMIN').AsString;
     // sCDI := Q.FindField('ANN_NOIDENTIF').AsString;
     sInspection := Q.FindField('ANN_COMPLTNOADMIN').AsString;
  end;
  Ferme(Q);

  try
     TobDpFiscal := TOB.Create('DPFISCAL', nil, -1);
     if TobDpFiscal.SelectDB ('"' + sGuidPerDos + '"', nil, TRUE) then
     begin

         bIIndExo := TobDpFiscal.GetValue ('DFI_IMPOINDIR') = 'E';

         // $$$ JP 27/04/06 - FQ 10964 - contrôle et màj CA3 uniquement si exonéré impots indirects
         if bIIndExo = FALSE then
         begin
              if sNoAdhesion = '' then
              begin
                   // ErreurChamp('ANL_NOADHESION', 'TSPAIE', Self, 'Le n° de dossier du code FRP est obligatoire');
                   exit;
              end;
              if sNoAdhCompl = '' then
              begin
                   // ErreurChamp('ANL_NOADHESION', 'TSPAIE', Self, 'La clé du code FRP est obligatoire');
                   exit;
              end;
         end;

         // $$$ JP 27/04/06 - FQ 10964 - contrôle de la clé uniquement si quelque chose de saisi
         if (ControleCleFRP (sRecette, sNoAdhesion, sNoAdhCompl) = FALSE) or ((sNoAdhCompl = '') and (sNoAdhesion <> '')) then
            begin
                 // ErreurChamp( 'ANL_NODHCOMPL', 'TSPAIE', Self, 'La clé du code FRP est incorrecte' );
                 exit;
            end;

         // RDI trouvé => met à jour coordonnées + n° Recette et CDI
         try
             // $$$ JP 27/04/06 - si RDI ou SIE (le SIE=CDI+RDI à priori)
             if (sTypePer = 'RDI') or (sTypePer = 'SIE') then
             begin
                  // Mise à jour du N° FRP
                  // $$$ jP 27/04/06 - ne pas màj si impoindir = 'E'
                  if bIIndExo = FALSE then
                  begin
                       // TobDpFiscal.PutValue ('DFI_NOFRP', sRecette + sNoAdhesion + sNoAdhCompl );
                       // TobDpFiscal.InsertOrUpdateDB(False);
                       ModeEdition(DS) ;
                       SetField('DFI_NOFRP', sRecette + sNoAdhesion + sNoAdhCompl);
                  end;
             end;

             // $$$ JP 27/04/06 - si CDI ou SIE
             if (sTypePer = 'CDI') or (sTypePer = 'SIE') then
             begin
                  // Mise à jour du N° inspection // ajout SMA
                  // TobDpFiscal.PutValue ('DFI_NOINSPECTION', sInspection ); // ajout SMA
                  // TobDpFiscal.InsertOrUpdateDB(False);
                  ModeEdition(DS);
                  SetField('DFI_NOINSPECTION', sInspection);
             end;
         except
               on E:Exception do PGIInfo('Erreur lors de la mise à jour des infos fiscales#13#10 ('+E.Message+')', TitreHalley);
         end;
     end;

  finally
         TobDpFiscal.Free;
  end;
end;


Initialization
  registerclasses([TOM_DPFISCAL]) ;
end.

