{***********UNITE*************************************************
Auteur  ...... : ???
Créé le ...... : 17/04/2000
Modifié le ... : 27/06/2000
Description .. : TOM de la fiche Tiers
Suite ........ :
Suite ........ : argument : T_NATUREAUXI=CLI si tiers client (ou FOU
Suite ........ : pour fournisseurs)
Mots clefs ... : TIERS;CLIENT;FOURNISSEUR
*****************************************************************}
unit UTomTiers;

// au 07/11/00 les zones :
// T_PAYEUR et T_FACTURE contienne en fichier la zone T_AUXILIAIRE
// T_PRESCRIPTUER, T_GROUPE ,T_APPORTEUR contienne en fichier  la zone T_TIERS
// dans tous les cas de figure, le bouton ellipsis fait une recherche sur T_TIERS

interface
uses  Extctrls,StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HCrc,HMsgBox,UTOM,UTob,HTB97,utilPGI,UtilGC,
{$IFDEF EAGLCLIENT}
      eFiche,eFichList,Maineagl,UtileAGL,Spin,
{$ELSE}
      Fiche,FichList,db,dbctrls,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,HDB,AglIsoflex,
{$IFDEF V530}
      EdtEtat,
{$ELSE}
      EdtREtat,
{$ENDIF}
{$ENDIF}
{$IFDEF DP}
      DpUtil,galoutil,
{$Endif}
      AglInit,Graphics,ent1,M3FP,LookUp,EntGc,menus,TiersUtil,aglInitGC,
      Windows,HRichOLE,Factutil,DicoBTP,messages,

{$IFNDEF CCS3}
      UtilConfid,
{$ENDIF}
{$IFDEF GRC}
      UtilRT,
{$IFDEF CTI}
      CtiGrc,UtilCti,CtiPcb,
{$ENDIF CTI}
{$ENDIF GRC}
      FactAdresse,
      wCommuns,
      UtofRTAnnuaire,
      ETudes,EtudesUtil,BtpUtil,uEntCommun;

Type


     TOM_TIERS = Class (TOM)
      private
         TiersCreation, client_particulier, FormaterLesZones : boolean;
         stNatureAuxi,StSuspect : String;
         ModifLot : boolean;
         StSQL : string;
         LgCode : Byte ;
         TobZonelibre : TOB;
         TobContact : TOB;
         TOBTIERSBTP : TOB;
         Old_Payeur, Old_Facture, Old_Livre : String;
         TE_LivreEnter, TE_FactureEnter: String;
         TE_PayeurEnter: String;
         EtatRisque : string;
         TOBProspect,TOBSuspect,TOBSuspectCompl : TOB;
         CritereSuspect : Boolean;
         num_contact : integer;    //CHR
         TransProCli : boolean;
{$IFDEF CTI}
         AppelCtiOk,SerieCTI,AppelCti,CtiCCA : Boolean;
         AppelantCTI : string;
         CtiHeureDebCom,CtiHeureFinCom : TDateTime;
         WhereUpdate : String;
         BPHONE,BTELBUR,BTELDOM,BTELPORT,BAttente,BStop : TToolBarButton97;
{$ENDIF}
         TiersToDuplic: String;
         DuplicFromMenu: Boolean;
         OldTiersOnCreat : string; // JT - eQualité 10384
{$IFDEF BTP}
    		 TOBEtude: TOB;
{$ENDIF}
         //FV1 : Gestion des options Terminaux portables
         CodeBarre      : THEdit;
         TypeCodeBarre  : THValComboBox;
         BCodeBarre     : TToolbarButton97;

         function  GestionFournisseur : boolean;
         procedure BCODEBARREOnClick(Sender : Tobject);
         function  RechercheTiers (Tiers : string; NatureTiers, stPayeur : string;var LibelleTiers : string) : boolean;

	     procedure AfterFormShow;
         procedure AfficheContactTiers;
         procedure SetLastError (Num : integer; ou : string );
         procedure SetArguments(StSQL : string);
         procedure SetBoutonEnabled(Etat : boolean);
         procedure AffecteTAuxiliaire(Const StAuxiliaire : string; Toujours : boolean);
         procedure AfficheEcranClientMode(cli_part : boolean);
         Procedure RechDPbis( Var CodePer : integer; Tiers, NatAuxi : String) ; //MCD 07/08/00
         procedure FormatelesZones(NomZone : string);
         procedure majcontact;
         Procedure ZonesLibresIsModified ;
         Procedure TestModifZoneLibre (Champ : String);
         Procedure TestModifZoneLibre2 (Champ : String);
         Procedure TestModifFActure;
         Procedure TestModifPayeur ;
         procedure TestModifLivre;
         procedure TOBCopieChamp(FromTOB, ToTOB : TOB);// GRC
         Function LimiteVersDemo_Mode : Boolean;
         procedure AddRemoveItemFromPopup(stPopup,stItem : string; bVisible : boolean);     //PCS
         procedure Duplication (CleDuplic: string);
         procedure ReInitDuplication(TobForm : TOB);

         function  RisqueTiers: String;
         procedure VoirEncours ;
         procedure VoirDetailEncours ;
         procedure GereIsoflex;
         procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
         Function  ChampsObligatoires : boolean;
         Procedure MajAffichageKardexEntreprise ;      // CHR
         Procedure RecupContact;      // CHR
         Procedure MajContact_chr ;      // CHR
         procedure tobToEcran (TobForm : TOB);
         procedure SuspectToEcran ; //GRC
         procedure RTTransProCli ;
         procedure RTListeEnseigne ;
         procedure TE_LIVREOnEnter(Sender: TObject);
         procedure TE_LIVREOnExit(Sender: TObject);
         procedure SetPlusNumAdresseLiv(Sender: TObject);
         procedure TE_FACTUREOnEnter(Sender: TObject);
         procedure TE_FactureOnExit(Sender: TObject);
//  BBI Fiche 11227
         procedure TE_PAYEUROnEnter(Sender: TObject);
         procedure TE_PAYEUROnExit(Sender: TObject);
//  BBI Fin Fiche 11227
         procedure SetPlusNumAdresseFac(Sender: TObject);
         procedure T_Email_OnDblclick(Sender : tObject);
         procedure T_RVA_OnDblclick(Sender : tObject);
         procedure T_TIERSOnExit(Sender : tObject);  // JT - eQualité 10384
         procedure BTPBordereaux;
         {$IFDEF EDI}
         procedure MnTiersEDI_OnClick(Sender: TObject);
         {$ENDIF}
    		function  GetInfoTiers: string;
    		procedure GestionBoutonGED;
{$IFDEF GRC}
 {$IFNDEF GIGI} // $$$ JP 09/08/07: pas de CTI en GI (PCL), car fait dans le Bureau exclusivement
         procedure BTELBUR_OnClick(Sender: TObject);
         procedure BTELDOM_OnClick(Sender: TObject);
         procedure BTELPORT_OnClick(Sender: TObject);
    		 procedure SetEtatBoutonCti;
    procedure BAttenteClick(Sender: TObject);
    procedure BPhoneClick(Sender: TObject);
    procedure BStopClick(Sender: TObject);
 {$ENDIF GIGI}
{$ENDIF GRC}
    procedure BVisuParcClick(Sender: TObject);
    procedure ChargeCodeBarre;
    procedure ParamInterfaceRexelClick(Sender: TObject);
    procedure RTListeCollectif;
      public
{$IFDEF CTI}
         procedure RTCTIAfficheMessageCTI;
         procedure RTCTIDecrocheAppel ;
         procedure RTCTIRaccrocheAppel ;
         procedure RTCTINumeroter ;
         procedure RTCTIAppelAttente ;
         procedure RTCTIRepriseAppel ;
{$ENDIF}
         procedure OnChangeField (F : TField)  ; override ;
         procedure OnUpdateRecord  ; override ;
         procedure OnAfterUpdateRecord  ; override ;
         procedure OnLoadRecord  ; override ;
         procedure OnNewRecord  ; override ;
         procedure OnDeleteRecord  ; override ;
         procedure OnCancelRecord  ; override ;
         procedure OnClose   ; override ;
         procedure OnArgument (Arguments : String )  ; override ;
     private
         Origine : String;    //permet de connaitre d'où est appelée la fiche (util pour DP)
         stContact : string;
         BTCONTRAT: TToolbarButton97;
         BTAPPELS : TToolbarButton97;

         DerniereCreate,CodeTiersDP : string;
         OnFerme : Boolean ;
         HandleSeria : Thandle;

         procedure BtAppels_OnClick(Sender: TObject);
         procedure BtContrat_OnClick(Sender: TObject);

         //PAUL OldInsert : boolean;  // permet de boucler sur insertion nouveau client
         Procedure ModifParticulier;
         Procedure AffIndisponible (AffMes : Boolean);
         Procedure MajIndisponible;
         Procedure YTC_INCOTERMEXIT(Sender : tObject);
         procedure DevisClick (Sender : Tobject);
         procedure BPLAN_OnClick(Sender: TObject);
         procedure BITI_OnClick(Sender: TObject);
//         procedure GoSeriaLSE (Sender : TObject);
//uniquement en line
//       Procedure CacheZonesNONLine;

     END ;

     TOM_TiersPiece = Class (TOM)
     private
       procedure SetLastError (Num : integer; ou : string );
     public
       procedure OnArgument (Arguments : String ); override ;
       procedure OnNewRecord  ; override ;
       procedure OnChangeField (F : TField)  ; override ;
       procedure OnUpdateRecord  ; override ;
       function  ReferenceCirculaire : Boolean;
     END ;

const
	// libellés des messages
  TexteMessage: array[1..64] of string  = (
(*
          {1}         'Le code doit être renseigné'
          {2}        ,'La raison sociale doit être renseignée'
          {3}        ,'Le client payeur spécifié n''existe pas ou n''est pas un client payeur'
          {4}        ,'Le code du représentant n''existe pas'
          {5}        ,'Le code du tiers facturé n''existe pas ou n''est pas un client'
          {6}        ,'Le code de l''apporteur n''existe pas'
          {7}        ,'Le tiers n''est pas un client'
          {8}        ,'Le destinataire de l''adresse est obligatoire'
          {9}        ,'Le code postal de l''adresse est obligatoire'
          {10}       ,'La ville de l''adresse est obligatoire'
          {11}       ,'Les natures de pièces suivantes et courantes doivent être différentes'
          {12}       ,'Vous ne pouvez choisir ce type de pièce suivante du fait de la circularité : XXX'
          {13}       ,'Le plafond autorisé ne doit pas être inférieur au crédit accordé'
          {14}       ,'Le montant de visa doit etre supérieur à zéro'
          {15}       ,'Le compte collectif doit être renseigné' + #13 + 'Veuillez le renseigner dans les paramètres comptables'
          {16}       ,'La nature d''auxiliaire n''est pas définie'
          {17}       ,'Une fiche existe déjà avec ce code tiers'
          {18}       ,'Le mode de règlement doit être renseigné'
          {19}       ,'Le régime fiscal doit être renseigné'
          {20}       ,'L''adresse doit être renseignée'
          {21}       ,'Une fiche existe déjà avec ce code auxiliaire'
          {22}       ,'Le mois de clôture à une valeur incorrecte'   // AFfaire
          {23}       ,'Le compte Auxiliaire doit être renseigné '
          {24}       ,'Le compte Collectif n''existe pas'
          {25}       ,'Le code groupe n''existe pas'
          {26}       ,'Le code prescripteur n''existe pas'
          {27}       ,'Problème de création de la fiche Suspect'  // GRC
          {28}       ,'Version de démonstration, vous n''êtes pas autorisé à créer plus de 10 tiers'
          {29}       ,'Le numéro de Siret ou Siren n''est pas correct '
          {30}       ,'Ce numéro de Siret existe déjà sur la fiche tiers : '// GRC
          {31}       ,'Vous ne pouvez créer qu''une seule adresse de facturation' // Affaire
          {32}       ,'Version de démonstration : le nombre de tiers excède le nombre autorisé.'
          {33}       ,'Champ obligatoire : '
          {34}       ,'Ce numéro de Siren existe déjà sur la fiche tiers : ' // mng : 34
          {35}       ,'Date d''expiration inférieure à la date de délivrance' //CHR
          {36}       ,'La saisie du champs suivant est obligatoire : '
          {37}       ,'Le code du tiers livré n''existe pas ou n''est pas un client'
          {38}       ,'Cette région n''appartiens pas à ce pays'
          {39}       ,'Le fournisseur payé spécifié n''existe pas'
          {40}       ,'L''emetteur de factures spécifié n''existe pas'
          {41}       ,'Le code NIF est obligatoire'
*)
          {1}  'Le code doit être renseigné'
          {2}  ,'La raison sociale doit être renseignée'
          {3}  ,'Le client payeur spécifié n''existe pas ou n''est pas un client payeur'
          {4}  ,'Le code du représentant n''existe pas'
          {5}  ,'Le code du tiers facturé n''existe pas ou n''est pas un client'
          {6}  ,'Le code de l''apporteur n''existe pas'
          {7}  ,'Le tiers n''est pas un client'
          {8}  ,'Le destinataire de l''adresse est obligatoire'
          {9}  ,'Le code postal de l''adresse est obligatoire'
          {10} ,'La ville de l''adresse est obligatoire'
          {11} ,'Les natures de pièces suivantes et courantes doivent être différentes'
          {12} ,'Vous ne pouvez choisir ce type de pièce suivante du fait de la circularité : XXX'
          {13} ,'Le plafond autorisé ne doit pas être inférieur au crédit accordé'
          {14} ,'Le montant de visa doit etre supérieur à zéro'
          {15} ,'Le compte collectif doit être renseigné#13Veuillez le renseigner dans les paramètres comptables'
          {16} ,'La nature d''auxiliaire n''est pas définie'
          {17} ,'Une fiche existe déjà avec ce code tiers'
          {18} ,'Le mode de règlement doit être renseigné'
          {19} ,'Le régime fiscal doit être renseigné'
          {20} ,'L''adresse doit être renseignée'
          {21} ,'Une fiche existe déjà avec ce code auxiliaire'
          {22} ,'Le mois de clôture à une valeur incorrecte'   // AFfaire
          {23} ,'Le compte Auxiliaire doit être renseigné '
          {24} ,'Le compte Collectif n''existe pas'
          {25} ,'Le code groupe n''existe pas'
          {26} ,'Le code prescripteur n''existe pas'
          {27} ,'Problème de création de la fiche Suspect'  // GRC
          {28} ,'Version de démonstration, vous n''êtes pas autorisé à créer plus de 10 tiers'
          {29} ,'Le numéro de Siret ou Siren n''est pas correct '
          {30} ,'Ce numéro de Siret existe déjà sur la fiche tiers : '// GRC
          {31} ,'Vous ne pouvez créer qu''une seule adresse de facturation' // Affaire
          {32} ,'Version de démonstration : le nombre de tiers excède le nombre autorisé.'
          {33} ,'Champ obligatoire : '
          {34} ,'Ce numéro de Siren existe déjà sur la fiche tiers : ' // mng : 34
          {35} ,'Date d''expiration inférieure à la date de délivrance' //CHR
          {36} ,'La saisie du champ suivant est obligatoire : '
          {37} ,'Le code du tiers livré n''existe pas ou n''est pas un client'
          {38} ,'Cette région n''appartient pas à ce pays'
          {39} ,'Le fournisseur payé spécifié n''existe pas'
          {40} ,'L''emetteur de factures spécifié n''existe pas'
          {41} ,'Le code NIF est obligatoire'
          {42} ,'Un code ressource libre est inexistant'
          {43} ,'Un code assistant libre est inexistant'
          {44} ,'Ce tiers n''est pas un tiers payeur'
          {45} ,'Le jour de naissance n''est pas correct'
          {46} ,'L''année de naissance doit être saisie sur 4 caractères'
          {47} ,'L''adresse de livraison n''existe pas'
          {48} ,'L''adresse de facturation n''existe pas'
          {49} ,'Le commercial est obligatoire'
          {50} ,'Cette enseigne n''existe pas'
          {51} ,'Ce transporteur n''existe pas'
          {52} ,'Ce fournisseur transporteur est utilisé dans les tarifs.' + #13 + ' Vous ne pouvez supprimer le type de fournisseur "transporteur".'
          {53} ,'Ce fournisseur transporteur est affecté à un tiers.' + #13 + ' Vous ne pouvez supprimer le type de fournisseur "transporteur".'
          {54} ,'Ce fournisseur transporteur est affecté à une adresse.' + #13 + ' Vous ne pouvez supprimer le type de fournisseur "transporteur".'
          {55} ,'Le compte n''est pas un collectif fournisseur.'
          {56} ,'Le compte n''est pas un collectif client.'
          {57} ,'La prestation de transport est incorrecte .'
          {58} ,'Le code frais renseigné est inconnu.'
          {59} ,'Les caractères spéciaux dans la codification sont interdits !'
          {60}, 'Le code barre et le qualifiant code barre doivent être renseignés  '
          {61}, 'Le nombre de caractères est incorrect'
          {62}, 'Le code barre est incorrect'
          {63}, 'Le code barre contient des caractères non autorisés'
          {64}, 'Le code barre doit être numérique'
                          );

	// Nom des champs à formater
	ChpAFormater: array[1..6] of string 	= (
          {1}        'T_LIBELLE'
          {2}       ,'T_PRENOM'
          {3}       ,'T_ADRESSE1'
          {4}       ,'T_ADRESSE2'
          {5}       ,'T_ADRESSE3'
          {6}       ,'T_VILLE'
                     );
{$IFDEF CTI}
  ModeAppelSortant : integer = 1 ; { pour forcer dans le cas d'une numérotation infructueuse : erreur connexion}
{$ENDIF}

function Tiers_MyAfterImport (Sender : TObject): string ;
procedure Tiers_GestionBoutonGed (Sender : TObject) ;

implementation

uses
  ParamSoc
  , uTomAdresses
  , SaisieList
  {$IFDEF EDI}
  , EDITiers
  {$ENDIF}
  , Tarifs
  , MailOl
  , Shellapi
  , Web
  , EntRt
  , CtiAlerte
  , UseriaLSe
  , UTilFonctionCalcul
  , UtilArticle
  , ConfidentAffaire
  , UFonctionsCBP
  , FormsName
  , CommonTools
  , HSysMenu
  ;

/////////////////////////////////////////////
// ****** TOM TIERS ***********************
/////////////////////////////////////////////

procedure TOM_TIERS.SetLastError (Num : integer; ou : string );
begin
if ou<>'' then SetFocusControl(ou);
LastError:=Num;
LastErrorMsg:=TexteMessage[LastError];
end ;

/////////////////////////////////////////////////////////////////////////////

procedure TOM_TIERS.SetArguments(StSQL : string);
var Critere,ChampMul,ValMul : string ;
    x,y : integer ;
    Ctrl : TControl;
    Fiche : TFFiche;
begin
DS.Edit;
Fiche := TFFiche(ecran);
Repeat
    //Critere:=uppercase(Trim(ReadTokenPipe(StSQL,'|'))) ;
    Critere:=Trim(ReadTokenPipe(StSQL,'|')) ;

    if Critere<>'' then
        begin
        x:=pos('=',Critere);
        if x<>0 then
           begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));
           y := pos(',',ValMul);
           if y<>0 then ValMul:=copy(ValMul,1,length(ValMul)-1);
           if copy(ValMul,1,1)='"' then ValMul:=copy(ValMul,2,length(ValMul));
           if copy(ValMul,length(ValMul),1)='"' then ValMul:=copy(ValMul,1,length(ValMul)-1);
           SetField(ChampMul,ValMul);
           Ctrl:=TControl(Fiche.FindComponent(ChampMul));
           if Ctrl=nil then exit;
{$IFDEF EAGLCLIENT}
           if (Ctrl is TCustomCheckBox) or (Ctrl is THValComboBox) Or (Ctrl is TCustomEdit) then
               begin
               TEdit(Ctrl).Font.Color:=clRed;
               SetControlText(ChampMul,ValMul);
               end
           else if Ctrl is TSpinEdit then TSpinEdit(Ctrl).Font.Color:=clRed
              //mcd 25/01/01 ajout test TH pour zones YTC de tierscompl
           else if (Ctrl is TCheckBox) or (Ctrl is THValComboBox) Or (Ctrl is THEdit)Or (Ctrl is THNumEdit)then
              begin
              TSpinEdit(Ctrl).Font.Color:=clRed;
              SetControlText(ChampMul,ValMul);
              end;
{$ELSE}
           if (Ctrl is TDBCheckBox) or (Ctrl is THDBValComboBox) Or (Ctrl is THDBEdit) then TEdit(Ctrl).Font.Color:=clRed
           else if Ctrl is THDBSpinEdit then THDBSpinEdit(Ctrl).Font.Color:=clRed
              //mcd 25/01/01 ajout test TH pour zones YTC de tierscompl
           else if (Ctrl is TCheckBox) or (Ctrl is THValComboBox) Or (Ctrl is THEdit)Or (Ctrl is THNumEdit)then
              begin
              THDBSpinEdit(Ctrl).Font.Color:=clRed;
              SetControlText(ChampMul,ValMul);
              end;
{$ENDIF}
           end;
        end;
until  Critere='';
end;

Procedure TOM_TIERS.YTC_INCOTERMEXIT(Sender : tObject);
var sIncoterm : string;
Begin
  sIncoterm := RechDom('GCINCOTERM', GetControlText('YTC_INCOTERM') , False);
  If sIncoterm = 'Error' then sIncoterm := '';
  SetControlCaption('TYTC_INCOTERM', sIncoterm);
End;

/////////////////////////////////////////////////////////////////////////////

procedure TOM_TIERS.OnArgument (Arguments : String );
var
  Critere,ChampMul,ValMul : string;
  i : integer;
  kk : integer ;
  pop : TPopupMenu ;
  stmp,st,stAuxi : string;
  tsArgsReception : TStringList;
  cc : ThEdit;
  Ctrl: TControl;
  CanInsertChoixCod : boolean;
  {$IFNDEF EAGLCLIENT}
  Confidentiel : TDBCheckBox;
  {$ENDIF EAGLCLIENT}
begin
  inherited ;
    TOBTIERSBTP := TOB.Create ('BTIERS',nil,-1);
    //
    //FV1 : 25/06/2015 - FS#1619 - MULTIPHONE NETCOM - la fonction de conversion des suspects en prospects
    if assigned(Getcontrol('T_BLOCNOTE')) then AppliqueFontDefaut (THRichEditOle(GetControl('T_BLOCNOTE')));
  if TmenuItem(GetControl('MnInterfRexel')) <> nil then
  begin
    TmenuItem(GetControl('MnInterfRexel')).OnClick := ParamInterfaceRexelClick;
  end;

   if Assigned(GetControl('BPLAN')) then
     ThEdit(GetControl('BPLAN')).OnClick := BPLAN_OnClick;

   if Assigned(GetControl('BITI')) then
     ThEdit(GetControl('BITI')).OnClick := BITI_OnClick;

  if (GetControl('BTCONTRAT') <> nil) then
     Begin
	   BtContrat 		 := TToolbarButton97(ecran.FindComponent('BTCONTRAT'));
  	 BtContrat.onclick:= BtContrat_OnClick;
     End;

  if (GetControl('BTAPPELS') <> nil) then
     Begin
	   BtAppels 		    := TToolbarButton97(ecran.FindComponent('BTAPPELS'));
 	   BtAppels.onclick := BtAppels_OnClick;
		 end;

{$IFDEF CTI}
	WhereUpdate:='';
	AppelCti:=false;
	CtiRaccrocherFiche:=False ;
	CtiCCA:=False;
	AppelantCTI:='';
{$ENDIF}

	FormaterLesZones:=GetParamsoc('SO_GCFORMATEZONECLI');

  // récupère l'argument ORIGINE si existe
	sTmp := StringReplace(Arguments, ';', chr(VK_RETURN), [rfReplaceAll]);
	tsArgsReception := TStringList.Create;
	tsArgsReception.Text := sTmp;
	origine := tsArgsReception.Values['ORIGINE'];
	tsArgsReception.Free;
	TobZonelibre:=TOB.Create ('TIERSCOMPL', Nil, -1);
	TobContact := Tob.create('CONTACT',nil,-1);

	{récup des arguments }
	i := pos('MODIFLOT',Arguments);
	ModifLot := i<>0;
	TOBProspect:=Nil;
	TOBSuspect:=Nil;
	TOBSuspectCompl:=Nil;
	CritereSuspect:=false;
	TiersCreation := True;
	TiersToDuplic := '';
	DuplicFromMenu := False;

  if (Arguments = '') then exit;

	if ModifLot then
  	 begin
  	 TFfiche(Ecran).MonoFiche:=true;
  	 StSQL := copy(Arguments,i+9,length(Arguments));
  	 StNatureAuxi:='CLI';     //mcd 27/02/01 pour OK si modif depuis fourn
  	 if (TFfiche(Ecran).name = 'GCFOURNISSEUR') then  StNatureAuxi:='FOU';
  	 end
	else
  	 begin
     TiersToDuplic := GetArgumentValue(Arguments, 'DUPLICATION');
  	 if TiersToDuplic <> '' then
        begin
        DuplicFromMenu := True;
        SetControlEnabled('BDUPLICATION',FALSE);
        end;
     stNatureAuxi := GetArgumentValue(Arguments, 'T_NATUREAUXI');
     StAuxi := GetArgumentValue(Arguments, 'T_AUXILIAIRE');
     SetField('T_AUXILIAIRE',stAuxi);
{$IFDEF CTI}
     WhereUpdate := GetArgumentValue(Arguments,'UPDCTI');
     // ce doit etre le dernier argument car il risque de comporter des ";"
     if Arguments <> '' then WhereUpdate := WhereUpdate+Arguments;
		 AppelantCTI:= GetArgumentValue(Arguments,'APPELANT');
{$ENDIF}
     //gestion des critères !!!
{$IFDEF CTI}
//     CritereSuspect := GetArgumentBoolean(Arguments, 'SUSPECT');
		 CritereSuspect := (Pos('SUSPECT',Arguments)>0);
     stContact := GetArgumentValue(Arguments,'CONTACT');
     SetControlVisible ('BCONTACT',  not GetArgumentBoolean(Arguments, 'ANNUAIRE'));
     //
     SerieCTI:=GetArgumentBoolean(Arguments, 'SERIECTI');
     if (GetArgumentBoolean(Arguments,'DECROCHE')) and (CtiHeureDeb<>0) then
        begin
        AppelCtiOk:=True;
        CtiHeureDebCom:=CtiHeureDeb;
        end;
     if (GetArgumentBoolean(Arguments,'COUPLAGE')) or
        (GetArgumentBoolean(Arguments,'NUMEROTER')) then
        RTFormCti:=Ecran;
     AppelCti:=GetArgumentBoolean(Arguments,'APPEL');
     CtiCCA:=GetArgumentBoolean(Arguments,'CCA');
     {cas particulier du CCA ou si montée de fiche, c que l'appel est ok }
	   if AppelCti and CtiCCA then AppelCtiOk:=True;
{$ENDIF}
     end;

  // Test droit d'accès en création

  if not assigned(TFfiche(Ecran)) then exit;
  
	if (stNatureAuxi = '') and (TFfiche(Ecran).name = 'GCFOURNISSEUR') then
  	 StNatureAuxi:='FOU';

	if stNatureAuxi = 'CON' then
  	 begin
     //SetControlVisible ('T_EMAILING',false);
     SetControlVisible ('PREGLEMENT',false);
     SetControlVisible ('T_CONFIDENTIEL',false);
     SetControlVisible ('TT_CONFIDENTIEL',false);
     //SetControlVisible ('T_PUBLIPOSTAGE',false);
     SetControlVisible ('GB_TARIFICATION' , False);
     SetControlVisible ('T_PARTICULIER' , False);
     SetControlVisible ('FEUVERT' , False);
     SetControlVisible ('PCONDITION',false);
     SetControlVisible ('TSCOMPLEMENT',false);
     AddRemoveItemFromPopup ('POPM','MNPRESCRIS',False);
     AddRemoveItemFromPopup ('POPM','MNFILLE',False);
     end;

  if stNatureAuxi = 'FOU' then
     begin
    if Assigned(GetControl('BT1_CODEBARRE'))       Then CodeBarre      := THEdit(GetControl('BT1_CODEBARRE'));
    if Assigned(GetControl('BT1_QUALIFCODEB')) Then TypeCodeBarre  := THValComboBox(GetControl('BT1_QUALIFCODEB'));
    if Assigned(GetControl('BCODEBARRE'))        Then BCodeBarre     := TToolbarButton97(GetControl('BCODEBARRE'));
    //
    SetControlProperty('T_FACTURE','DataType',   'TZTFOURN');
     SetControlProperty('T_PAYEUR', 'DataType', 'TZTPAYEURFOU');
     AddRemoveItemFromPopup ('POPM','MNVTECLI',False);  // mcd 05/03/01 ne marche pas pour l'instant ...
     AddRemoveItemFromPopup ('POPM','MNTARIFGRP',False);  // JT eQualité 10331
    //
    BCodeBarre.OnClick    := BCODEBARREOnclick;
    //
    CodeBarre.Visible     := False;
    TypeCodeBarre.Visible := False;
    //BCodeBarre.Visible    := False;
    SetControlVisible('TT_CODEBARRE',   False);

     end;

	if ((StNatureAuxi<>'CLI') and (StNatureAuxi<>'PRO')) then
  	 BEGIN
    SetControlVisible('T_REPRESENTANT',  False);
    SetControlVisible('TT_REPRESENTANT', False);
    SetControlVisible('GBREPRESENTANT',  False);
     END ;

  // debut Affaire
  SetControlVisible ('BDP',False);

{$IFDEF EAGLCLIENT}
  Ctrl := GetControl('T_EMAIL');
  if Assigned(Ctrl) and (Ctrl is thEdit) then ThEdit(Ctrl).OnDblclick := T_Email_OnDblclick; { GPAO1_TIERS}
  Ctrl := GetControl('T_RVA');
  if assigned(Ctrl) and (Ctrl is ThEdit) then ThEdit(Ctrl).OnDblclick := T_RVA_OnDblclick;
{$ELSE}
  Ctrl := GetControl('T_EMAIL');
  if Assigned(Ctrl) and (Ctrl is thDBEdit) then ThDBEdit(Ctrl).OnDblclick := T_Email_OnDblclick; { GPAO1_TIERS}
  Ctrl := GetControl('T_RVA');
  if Assigned(Ctrl) and (Ctrl is ThDBEdit) then ThDBEdit(Ctrl).OnDblclick := T_RVA_OnDblclick; { GPAO1_TIERS}
{$ENDIF}
	SetControlVisible ('GBCLOTURE', False);
	SetControlVisible ('TT_MOISCLOTURE' , False);
	SetControlVisible ('T_MOISCLOTURE' , False);

If (ctxAffaire in V_PGI.PGIContexte) then
begin
    if (StNatureAuxi ='FOU') then
    begin
        SetControlVisible ('GBEXPORT',False);
        SetControlVisible ('GBEXPEDITION' , FALSE);
        SetControlVisible ('T_SOUMISTPF' , FALSE);
        If ctxSCot in V_PGI.PGIContexte  then
            begin
            // specif pour lien DP
                if (GetParamsoc('SO_AFLIENDP') =true)  then
                begin     // cas SCOT avec lien DP
{$IFDEF DP}
                    DPInit(Ecran);
{$endif}
                    SetControlVisible ('T_APE',False);
                    SetControlVisible ('TT_APE',False);
                    SetControlVisible ('BDP',True);
                end;
                if (GetParamsoc('SO_AFLIENDP') =true) and (origine = 'DP') then
                begin    //cas lien DP et SCOT et appel depuis le DP
{$IFDEF DP}
                    DpGrise(Ecran);        // grise les zones communes
{$endif}
                    SetControlVisible ('BDP',FAlse);
                    SetControlVisible ('BINSERT' , FALSE);
                end;
            end;
        end
    else
    begin
      if (stNatureAuxi = 'CLI') or (stNatureAuxi = 'PRO') then
        CanInsertChoixCod := Tools.CanInsertedInTable('CHOIXCOD'{$IFDEF APPSRV}, '', '' {$ENDIF APPSRV})
      else
        CanInsertChoixCod := True;
      cc:=THEdit(GetControl('T_MODEREGLE'));
      SetControlProperty('T_MODEREGLE', 'DataTypeParametrable', TRUE);
      i := cc.width;
      SetControlProperty('T_MODEREGLE', 'Width', i-20);
      cc:=THEdit(GetControl('T_ORIGINETIERS'));
      SetControlProperty('T_ORIGINETIERS', 'DataTypeParametrable', CanInsertChoixCod);
      i := cc.width;
      SetControlProperty('T_ORIGINETIERS', 'Width', i-20);
      cc:=THEdit(GetControl('T_SECTEUR'));
      SetControlProperty('T_SECTEUR', 'DataTypeParametrable', CanInsertChoixCod);
      i := cc.width;
      SetControlProperty('T_SECTEUR', 'Width', i-20);
      cc:=THEdit(GetControl('T_TARIFTIERS'));
      SetControlProperty('T_TARIFTIERS', 'DataTypeParametrable', TRUE);
      i := cc.width;
      SetControlProperty('T_TARIFTIERS', 'Width', i-20);
      cc:=THEdit(GetControl('T_COMPTATIERS'));
      SetControlProperty('T_COMPTATIERS', 'DataTypeParametrable', TRUE);
      i := cc.width;
      SetControlProperty('T_COMPTATIERS', 'Width', i-20);
      SetControlVisible ('GB_RESSOURCE' , True);
      SetControlVisible ('GBEXPEDITION' , FALSE);
      SetControlVisible ('GB_RELEVE' , FALSE);
      SetControlVisible ('T_SOUMISTPF' , FALSE);
                   // si pas de gestion apporteur et commissionnement
    {if Not(VH_GC.AFGestionCom) then begin
      SetControlVisible ('T_APPORTEUR' , False);
      SetControlVisible ('TT_APPORTEUR' , False);
      SetControlVisible ('T_LIBELLEAPP' , False);
      SetControlVisible ('T_COEFCOMMA' , False);
      SetControlVisible ('TT_COEFCOMMA' , False);
      end;  }

      If ctxTempo in V_PGI.PGIContexte  then SetControlVisible ('GBEXPORT',False);

      If ctxSCot in V_PGI.PGIContexte  then
      begin
             // pas de gestion representatn pour l'instant
          SetControlVisible ('GBREPRESENTANT', False);
          SetControlVisible ('T_REPRESENTANT' , FALSE);
          SetControlVisible ('TT_REPRESENTANT' , FALSE);
          SetControlVisible ('T_PRESCRIPTEUR' , False);
          SetControlVisible ('TT_PRESCRIPTEUR' , False);
          SetControlVisible ('TT_PRESCRIPTEUR_' , False);
          SetControlVisible ('T_PUBLIPOSTAGE' , False);
          SetControlVisible ('T_SEXE' , False);
          SetControlVisible ('TT_SEXE' , False);
          SetControlVisible ('T_ZONECOM' , False);
          SetControlVisible ('GBCLOTURE', True);
          SetControlVisible ('TT_MOISCLOTURE' , True);
          SetControlVisible ('T_MOISCLOTURE' , True);
          SetControlProperty('T_MOISCLOTURE', 'EditorEnabled', False); //mcd 01/10/03 oblige passer par ascenseur pour résourdre cas si saisie > 12 pas pris en compte equalité 10372
          SetControlVisible ('T_CONSO' , True);
          SetControlVisible ('TT_CONSO' , True);
          SetControlVisible ('GBPARTICULIER', FALSE);
          SetControlVisible ('GBEXPORT',False);
          SetControlVisible ('T_PROFIL',False);
          SetControlVisible ('TT_PROFIL',False);
          SetControlVisible ('TT_ENSEIGNE',False);
          SetControlVisible ('TE_LIVRE',False);
          SetControlVisible ('TLIB_LIVRE',False);
          SetControlVisible ('TYTC_NADRESSELIV',False);
          SetControlVisible ('YTC_NADRESSELIV',False);
        // specif pour lien DP
          if (GetParamsoc('SO_AFLIENDP') =true)  then
          begin     // cas SCOT avec lien DP
{$IFDEF DP}
              DPInit(Ecran);
{$endif}
              SetControlVisible ('T_APE',False);
              SetControlVisible ('TT_APE',False);
              SetControlVisible ('BDP',True);
          end;
          if (GetParamsoc('SO_AFLIENDP') =true) and (origine = 'DP') then
          begin    //cas lien DP et SCOT et appel depuis le DP
{$IFDEF DP}
              DpGrise(Ecran);        // grise les zones communes
{$endif}
              SetControlVisible ('BDP',FAlse);
              SetControlVisible ('BINSERT' , FALSE);
          end;
          SetCOntrolText('XXcodeper','0');
      end;
    end;
end
else   // fin Affaire
begin
//   if (ctxGCAFF in V_PGI.PGIContexte) then
    if VH_GC.GASeria then      // gm le 29/9/03 sinon on retrouve les ressoures dans la GC avec
                               // l'affaire non sérailisée
        SetControlVisible ('GB_RESSOURCE' , True)
    else
        SetControlVisible ('GB_RESSOURCE' , False);
{$IFNDEF CHR}
    if not VH_GC.GCIfDefCEGID then
    begin
{$IFDEF EAGLCLIENT}
      if not (ctxGCAFF in V_PGI.PGIContexte) then
          TWinControl(GetControl('T_BLOCNOTE')).Align:=alBottom;
{$ELSE}
      if not (ctxGCAFF in V_PGI.PGIContexte) then
          THDBRichEditOLE(GetControl('T_BLOCNOTE')).Align:=alBottom;
{$ENDIF}
    end;
{$ENDIF}
    if stNatureAuxi = 'CON' then // mng 30-05-02
        begin
        SetControlVisible ('GB_STATISTIQUE' , False);
        SetControlVisible ('GB_VALEURS' , False);
        SetControlVisible ('GB_DATES' , False);
        SetControlVisible ('GB_TEXTES' , False);
        SetControlVisible ('GB_DECISIONS' , False);
        SetControlVisible ('BRIB' , False);
        SetControlVisible ('BMEMO' , False);
        SetControlVisible ('BCONTACT' , False);
        SetControlVisible ('BNEWMAIL' , False);
        SetControlVisible ('T_EMAIL' , False);
        SetControlVisible ('TT_EMAIL' , False);
        SetControlVisible ('GB_RESSOURCE' , False);
        //SetControlProperty ('PNINFOHAUT' ,'Height',0 );
        //SetControlProperty ('PNINFOHAUT' ,'Width',0 );
        SetControlProperty ('PNINFOHAUTGAUCHE' ,'Height',0 );
        SetControlProperty ('PNINFOHAUTGAUCHE' ,'Width',0 );
        SetControlProperty ('PNINFOHAUTDROIT' ,'Height',0 );
        SetControlProperty ('PNINFOHAUTDROIT' ,'Width',0 );
        //TPanel(GetControl('PNINFOHAUT')).Align:=alNone;
        TPanel(GetControl('PNINFOHAUTGAUCHE')).Align:=alNone;
        TPanel(GetControl('PNINFOHAUTDROIT')).Align:=alNone;
        {$IFDEF EAGLCLIENT}
          TWinControl(GetControl('T_BLOCNOTE')).Parent:=TPanel(GetControl('PNINFOHAUT'));
          TWinControl(GetControl('T_BLOCNOTE')).Align:=alClient;
        {$ELSE EAGLCLIENT}
          THDBRichEditOLE(GetControl('T_BLOCNOTE')).Parent:=TPanel(GetControl('PNINFOHAUT'));
          THDBRichEditOLE(GetControl('T_BLOCNOTE')).Align:=alClient;
        {$ENDIF EAGLCLIENT}
        TLabel(GetControl('TT_PAYS1')).Parent:=TPanel(GetControl('PNINFOBAS'));
        THValComboBox(GetControl('T_ZONECOM')).Parent:=TPanel(GetControl('PNINFOBAS'));
        TGroupBox(GetControl('GB_INFO')).Parent:=TPanel(GetControl('PNINFOBAS'));
        end;
    end;

      // pour rendre visible ou non la zone famille comptable
  SetControlVisible ('T_COMPTATIERS' , GetParamSocSecur('SO_GCVENTCPTATIERS', False));
  SetControlVisible ('TT_COMPTATIERS', GetParamSocSecur('SO_GCVENTCPTATIERS', False));

if (TFfiche(Ecran).name = 'HRCLIENTS') then   //CHR
    BEGIN
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'YTC_TABLELIBRETIERS', 10, '');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_VALLIBRE', 3, '');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_DATELIBRE', 3, '');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_TEXTELIBRE', 3, '');
    END;

//¨Paramétrage des libellés des tables libres
//if (TFfiche(Ecran).name = 'GCTIERS') or (TFfiche(Ecran).name = 'GCTIERSFO') then  // DBR fiche 10106
if (Copy (TFFiche(Ecran).name, 1, 7) = 'GCTIERS') or (Copy (TFfiche(Ecran).name, 1, 9) = 'GCTIERSFO') then
    BEGIN
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'YTC_TABLELIBRETIERS', 10, '');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_VALLIBRE', 3, '');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_DATELIBRE', 3, '');
    GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'YTC_BOOLLIBRE', 3, '');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_TEXTELIBRE', 3, '');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '');
    If ctxMode in V_PGI.PGIContexte then
        begin
        SetControlVisible ('T_SOCIETEGROUPE' , False);
        SetControlVisible ('TT_SOCIETEGROUPE' , False);
        SetControlVisible ('T_SOCIETEGROUPE_' , False);
        end;
    if VH_GC.GCIfDefCEGID then
    begin
      for kk:=1 to 9 do SetControlProperty('YTC_TABLELIBRETIERS'+IntToStr(kk),'DataTypeParametrable',False) ;
      SetControlProperty('YTC_TABLELIBRETIERSA','DataTypeParametrable',False) ;
    end;
    if JaiLeDroitTag(323210) then
    begin
      SetControlVisible ('BVISUPARC',  true);
      TToolBarButton97(GetControl('BVISUPARC')).onclick := BVisuParcClick;
    end;
    END;
if (TFfiche(Ecran).name = 'GCFOURNISSEUR') then
    BEGIN
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'YTC_TABLELIBREFOU', 3, '');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_VALLIBREFOU', 3, '');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_DATELIBREFOU', 3, '');
    END;

if not (ctxGRC in V_PGI.PGIContexte) then
    begin
    pop :=TPopupMenu(GetControl('POP2M') ) ;
    if ( Copy(Ecran.Name,1,7) = 'GCTIERS') then
        for i:=0 to pop.items.count-1 do
            begin
            st:=uppercase(pop.items[i].name);
            if (st='MNPROSPECT') or (st='MNACTIONS' ) or (st='MNPROJETS') or (st='MNPROPOSITIONS')  or (st='MNPROSPCOMPL') then  pop.items[i].visible:=false;
            end;
    end ;
// GRC
if (ctxGRC in V_PGI.PGIContexte) then
    begin
    SetControlProperty ('T_APPORTEUR', 'DataType', 'RTTIERSPRO');
    SetControlProperty('T_SOCIETEGROUPE', 'DataType', 'RTTIERSPROTIERS');
    if ( Copy(Ecran.Name,1,7) = 'GCTIERS') then
        begin
        if stNatureAuxi = 'CON' then
            begin
            AddRemoveItemFromPopup ('POP2M','MNACTIONS',False);
            AddRemoveItemFromPopup ('POP2M','MNPROJETS',False);
            AddRemoveItemFromPopup ('POP2M','MNPROPOSITIONS',False);
            AddRemoveItemFromPopup ('POP2M','MNTIERSPIECE',False);
            AddRemoveItemFromPopup ('POP2M','MNPROSPECT',False);
            AddRemoveItemFromPopup ('POP2M','MNPROSPCOMPL',False);
            AddRemoveItemFromPopup ('POP2M','MNREFERENCEMENT',False);
            AddRemoveItemFromPopup ('POP2M','MNTIERSFRAIS',False);
            AddRemoveItemFromPopup ('POP2M','MNTIERSEDI',False);

            AddRemoveItemFromPopup ('POPM','MNPIECECOURS',False);
            AddRemoveItemFromPopup ('POPM','MNARTICLECOM',False);
            AddRemoveItemFromPopup ('POPM','MNENCOURS',False);
            AddRemoveItemFromPopup ('POPM','MNTARIF',False);
            AddRemoveItemFromPopup ('POPM','MNTARIFFAMILLE',False);
            AddRemoveItemFromPopup ('POPM','MNAFFAIRE',False);
            AddRemoveItemFromPopup ('POPM','MNVISUPIECE',False);
            end
        else
          begin
            if GetParamSoc('SO_RTPROJGESTION') = False then
              AddRemoveItemFromPopup ('POP2M','MNPROJETS',False);
{$IFDEF CCS3}
            AddRemoveItemFromPopup ('POPM','MNAFFAIRE',False);
            AddRemoveItemFromPopup ('POP2M','MNPROPOSITIONS',False);
            AddRemoveItemFromPopup ('POP2M','MNPROSPCOMPL',False);
{$ENDIF}
// confidentialité , concept Modif & Creation Prospects
            if (TFFiche(Ecran).fTypeAction<>taConsult) and (StNatureAuxi = 'PRO')then
            BEGIN
              if (Not ExJaiLeDroitConcept(TConcept(bt510),False)) and (TFFiche(Ecran).fTypeAction<>taCreat) then  //confidentialité , concept Modif Prospect
              begin
              	TFFiche(Ecran).fTypeAction := taConsult;
                SetControlVisible('BDUPLICATION',False) ;
              end;

              if Not ExJaiLeDroitConcept(TConcept(gcProspectCreat),False) then  //confidentialité , concept Creation Prospect
                BEGIN
                SetControlVisible('BINSERT',False) ;
                SetControlVisible('BDUPLICATION',False) ;
                END ;
              END ;
            end;
          end;
    if CritereSuspect then StSuspect := stauxi;// GRC-Code Suspect pour Bascule Suspect en Prospect
    end;

if Assigned(GetControl('TE_LIVRE')) then
begin
   ThEdit(GetControl('TE_LIVRE')).OnEnter := TE_LIVREOnEnter;
   ThEdit(GetControl('TE_LIVRE')).OnExit := TE_LIVREOnExit;
end;
if Assigned(GetControl('TE_FACTURE')) then
begin
   ThEdit(GetControl('TE_FACTURE')).OnEnter := TE_FACTUREOnEnter;
   ThEdit(GetControl('TE_FACTURE')).OnExit := TE_FACTUREOnExit;
end;
if Assigned(GetControl('TE_PAYEUR')) then
begin
   ThEdit(GetControl('TE_PAYEUR')).OnEnter := TE_PAYEUROnEnter;
   ThEdit(GetControl('TE_PAYEUR')).OnExit := TE_PAYEUROnExit;
end;
  {$IFDEF GRC}
   {$IFNDEF GIGI} // $$$ JP 09/08/07: pas de CTI en GI (PCL), car fait dans le Bureau exclusivement
     if Assigned(GetControl('BTELPORT')) then
       TToolBarButton97(GetControl('BTELPORT')).OnClick := BTELPORT_OnClick;

     if Assigned(GetControl('BTELBUR')) then
       TToolBarButton97(GetControl('BTELBUR')).OnClick := BTELBUR_OnClick;

     if Assigned(GetControl('BTELDOM')) then
       TToolBarButton97(GetControl('BTELDOM')).OnClick := BTELDOM_OnClick;
   {$ENDIF GIGI}
  {$ENDIF GRC}

// JT - eQualité 10384
ThEdit(GetControl('T_TIERS')).OnExit := T_TIERSOnExit;
//uniquement en line
//Thedit(GetControl('T_TIERS')).MaxLength := getparamsoc('SO_LGCPTEAUX');
//Thedit(GetControl('T_AUXILIAIRE')).MaxLength := getparamsoc('SO_LGCPTEGEN');
// JT eQualité 11030
{$IFNDEF EAGLCLIENT}
if GetControl('T_CONFIDENTIEL') is TDBCheckBox then
begin
  Confidentiel := TDBCheckBox(GetControl('T_CONFIDENTIEL'));
  Confidentiel.ValueChecked := '1';
  Confidentiel.ValueUnChecked := '0';
end;
{$ENDIF EAGLCLIENT}

//gestion des boutons appels et contrats
ThEdit(GetControl('T_TIERS')).OnExit := T_TIERSOnExit;

//SetControlVisible('BCOURRIER',(ctxGRC in V_PGI.PGIContexte));
{$IFDEF CTI}
//  SetControlVisible('BPHONE',True);
{$ELSE CTI}
//  SetControlVisible('BPHONE',False);
{$ENDIF CTI}

GereIsoflex; // Affichage du bouton de gestion documentaire

pop :=TPopupMenu(GetControl('POP2M') ) ;
if (Ecran.Name = 'GCTIERS') then
    for i:=0 to pop.items.count-1 do
        begin
{$IFDEF CCS3}
        if (uppercase(pop.items[i].name)='MNREFERENCEMENT') then pop.items[i].enabled:=false;
{$ELSE}
        if (uppercase(pop.items[i].name)='MNREFERENCEMENT') AND ((TFFiche(Ecran).FtypeAction=taConsult) or (stNatureAuxi<>'CLI')) then
           pop.items[i].enabled:=false;
{$ENDIF}
        end;

          // pour affaire, on supprime les pop up voulu
  If (ctxAffaire in V_PGI.PGIContexte) And (stNatureAuxi = 'CLI')then
     begin
       AddRemoveItemFromPopup ('POP2M','MNREFERENCEMENT',False);
       AddRemoveItemFromPopup ('POPM','MNARTICLECOM',False);
       AddRemoveItemFromPopup ('POPM','MNART',False);
       AddRemoveItemFromPopup ('POPM','MNVTECLI',False);
       If ctxscot in V_PGI.PGIContexte then
        begin
        AddRemoveItemFromPopup ('POP2M','MNPROSPECT',False);
        AddRemoveItemFromPopup ('POP2M','MNACTIONS',False);
        AddRemoveItemFromPopup ('POP2M','MNPROJETS',False);
        AddRemoveItemFromPopup ('POPM','MNPRESCRIS',False);
        AddRemoveItemFromPopup ('POPM','MNTARIF',False);
        AddRemoveItemFromPopup ('POPM','MNTARIFFAMILLE',False);
        AddRemoveItemFromPopup ('POPM','MNVISUPIECE',False);
        AddRemoveItemFromPopup ('POPM','MNVOIRPIECE',False);
        AddRemoveItemFromPopup ('POP2M','MNTIERSFRAIS',False);
        end ;
   end
else begin   // on supprimer popu spcéif affaire
   pop :=TPopupMenu(GetControl('POPM') ) ;
  for i:=0 to pop.items.count-1 do
      begin
      st:=uppercase(pop.items[i].name);
      if ((st='MNPROPO') or (st='MNTB')) then pop.items[i].visible:=false;
      //if (st='MNAFFAIRE') and Not(ctxGCAFF in V_PGI.PGIContexte) then pop.items[i].visible:=false;
      end;
   end;
{if stNatureAuxi <> 'FOU' then
   begin
   pop := TPopupMenu(Ecran.FindComponent('POPM'));
       for i := 0 to Pop.Items.Count - 1 do
       begin
            if uppercase(Pop.Items[i].Name) = 'MNCATALOGUE' then break;
       end;
    if i < Pop.Items.Count then Pop.items[i].Visible := False;
    end;  }

if stNatureauxi = 'FOU' then // spécificités sur les fournisseurs
begin
   //if not (ctxMode in V_PGI.PGIContexte) then SetControlVisible('BCODEBARRE', False);
   //
   //SetControlVisible('BCODEBARRE', True);
   //
   if ctxFO in V_PGI.PGIContexte then
   begin
       SetControlVisible('BPOPMENU', False);
       SetControlVisible('BPOP2MENU', False);
   end ;
   if ctxAffaire in V_PGI.PGIContexte then
   begin
     //SetControlVisible('BCODEBARRE', False);
     If ctxscot in V_PGI.PGIContexte then
     begin
       AddRemoveItemFromPopup ('POP2M','MNREFERENCEMENT',False);
       AddRemoveItemFromPopup ('POP2M','MNPROSPECT',False);
       AddRemoveItemFromPopup ('POP2M','MNACTIONS',False);
       AddRemoveItemFromPopup ('POP2M','MNPERSPECTIVE',False);
       AddRemoveItemFromPopup ('POPM','MNARTICLECOM',False);
       AddRemoveItemFromPopup ('POPM','MNART',False);
       AddRemoveItemFromPopup ('POPM','MNTARIF',False);
       AddRemoveItemFromPopup ('POPM','MNTARIFFAMILLE',False);
     end else
     begin
       AddRemoveItemFromPopup ('POP2M','MNREFERENCEMENT',False);
       AddRemoveItemFromPopup ('POPM','MNART',False);
     end;
   end;
{$IFDEF BTP}
end else if stNatureauxi = 'CLI' then
begin
	if Assigned(TMenuItem (GetCOntrol('mnDevis1'))) then
  begin
    TMenuItem (GetCOntrol('mnDevis1')).OnClick := DevisClick;
  end;
{$ENDIF}
{ mng CTI }
{$IFDEF CTI}
  BPhone := TToolbarButton97(ecran.FindComponent('BPhone'));
  BPhone.OnClick := BPhoneClick;

  BStop := TToolbarButton97(ecran.FindComponent('BStop'));
  BStop.onClick := BStopClick;

  BAttente := TToolbarButton97(ecran.FindComponent('BATTENTE'));
  BAttente.onClick := BAttenteClick;

  BTELDOM := TToolbarButton97(ecran.FindComponent('BTELDOM'));
  if Assigned(GetControl('BTELDOM')) then
    TToolBarButton97(GetControl('BTELDOM')).OnClick := BTELDOM_OnClick;
  BTELBUR := TToolbarButton97(ecran.FindComponent('BTELBUR'));
  if Assigned(GetControl('BTELBUR')) then
    TToolBarButton97(GetControl('BTELBUR')).OnClick := BTELBUR_OnClick;
  BTELPORT := TToolbarButton97(ecran.FindComponent('BTELPORT'));
  if Assigned(GetControl('BTELPORT')) then
    TToolBarButton97(GetControl('BTELPORT')).OnClick := BTELPORT_OnClick;
  //
  if not GetParamSocSecur('SO_RTCTIGESTION', False) then
  begin
    BPhone.Visible := GetParamSocSecur('SO_RTCTIGESTION', False);
    BAttente.Visible:= GetParamSocSecur('SO_RTCTIGESTION', False);
    BStop.Visible:= GetParamSocSecur('SO_RTCTIGESTION', False);
  end else
  begin
  	SetEtatBoutonCti;
    If TFFiche(Ecran).fTypeAction = TaModif then
    Begin
       BTELDOM.Visible := true;
       BTELBUR.Visible := true;
       BTELPORT.Visible := true;
    end;
  end;
{$ENDIF}
end;

  // confidentialité , concept Creation Clients
  //FV1 : 04/11/2013 - FS#694 - POUCHAIN (V8) : création fourn. impossible si utilisateur n'a pas le droit de modifier un client
  if TFFiche(Ecran).fTypeAction=TaCreat then
  BEGIN
    if ((StNatureAuxi = 'PRO') and (Not ExJaiLeDroitConcept(TConcept(gcProspectCreat),False))) or
       ((StNatureAuxi = 'CLI') and (Not ExJaiLeDroitConcept(TConcept(gcCLICreat),False))) Or
       ((StNatureAuxi = 'FOU') and (Not ExJaiLeDroitConcept(TConcept(gcFouCreat),False))) then
    BEGIN
      PGIBox('Vous n''avez pas les droits d''accès nécessaires !', 'Création impossible');
      SetControlVisible('BINSERT',False) ;
      SetControlVisible('BDUPLICATION',False) ;
      TFFiche(Ecran).fTypeAction := TaConsult;
    END ;
  end;

  if TFFiche(Ecran).fTypeAction=TaModif then
  BEGIN
    if ((StNatureAuxi = 'PRO') and (Not ExJaiLeDroitConcept(TConcept(bt510),False))) or
       ((StNatureAuxi = 'CLI') and (Not ExJaiLeDroitConcept(TConcept(bt511),False))) Or
       ((StNatureAuxi = 'FOU') and (Not ExJaiLeDroitConcept(TConcept(bt518),False))) then
    Begin
      PGIBox('Vous n''avez pas les droits d''accès nécessaires !', 'Modification impossible');
      TFFiche(Ecran).fTypeAction := TaConsult;
    end;
  end;
  if TFFiche(Ecran).fTypeAction=TaConsult then SetControlVisible ('BDUPLICATION',false);

  if Not ExJaiLeDroitConcept(TConcept(gcAdminECommerce),False) then
  begin
    SetControlVisible('T_PASSWINTERNET',False) ;
    SetControlVisible('TT_PASSWINTERNET',False) ;
  end ;

if (ecran <> Nil)then TFFiche(Ecran).OnKeyDown:=FormKeyDown ;

{$IFDEF BTP}
if not (ctxGRC in V_PGI.PGIContexte) then
begin
  AddRemoveItemFromPopup ('POPM','MNTB',False);
  AddRemoveItemFromPopup ('POPM','MNPROPO',False);
  AddRemoveItemFromPopup ('POPM','MNJUSTIF',False);
  AddRemoveItemFromPopup ('POP2M','MNPROJETS',False);
  AddRemoveItemFromPopup ('POP2M','MNACTIONS',False);
  AddRemoveItemFromPopup ('POP2M','MNPROPOSITIONS',False);
  AddRemoveItemFromPopup ('POP2M','MNREFERENCEMENT',False);
  AddRemoveItemFromPopup ('POP2M','MNTIERSPIECE',False);
  AddRemoveItemFromPopup ('POP2M','MNPROSPECT',False);
  AddRemoveItemFromPopup ('POP2M','MNPROSPCOMPL',False);
end;

//if stNatureauxi = 'FOU' then
//    SetControlVisible ('BPOP2MENU',False);
SetControlVisible ('GB_RESSOURCE' , False);
SetControlVisible ('T_PROFIL' , False);
SetControlVisible ('TT_PROFIL' , False);
AddRemoveItemFromPopup ('POP2M','MNTIERSFRAIS',False);
{$ENDIF}

{$IFDEF CHR}
AddRemoveItemFromPopup ('POPM','MNSGED',False);
AddRemoveItemFromPopup ('POPM','MNTB',False);
AddRemoveItemFromPopup ('POPM','MNPROPO',False);
AddRemoveItemFromPopup ('POPM','MNJUSTIF',False);
AddRemoveItemFromPopup ('POPM','MNVTECLI',False);
AddRemoveItemFromPopup ('POPM','MNART',False);
AddRemoveItemFromPopup ('POPM','MNTARIF',False);
AddRemoveItemFromPopup ('POPM','MNGROUPE',False);
AddRemoveItemFromPopup ('POPM','MNENCOURS',False);
If not (CtxGRC in V_PGI.PgiContexte) then
begin
  {$IFDEF CCS5}
  AddRemoveItemFromPopup ('POP2M','MNPROPOSITIONS',False);
  AddRemoveItemFromPopup ('POP2M','MNPROSPCOMPL',False);
  {$ENDIF}
  AddRemoveItemFromPopup ('POP2M','MNPROJETS',False);
  AddRemoveItemFromPopup ('POP2M','MNREFERENCEMENT',False);
  AddRemoveItemFromPopup ('POP2M','MNPROSPECT',False);
end;
if stNatureauxi = 'FOU' then
  SetControlVisible ('BPOP2MENU',False);
SetControlVisible ('GB_RESSOURCE' , False);
SetControlVisible ('T_PROFIL' , False);
SetControlVisible ('TT_PROFIL' , False);
{$ENDIF}


// AIde en ligne
if stNatureauxi = 'CON' then TFFiche(Ecran).HelpContext:=111000316 ;

SetActiveTabSheet('PGeneral');
{$IFDEF EDI}
  if Assigned(GetControl('MnTiersEDI')) then
    TMenuItem(GetControl('MnTiersEDI')).OnClick := MnTiersEDI_OnClick;
{$ENDIF}

if (not GetParamSoc('SO_PREFSYSTTARIF')) then
  AddRemoveItemFromPopup ('POPM','MNTARIFFAMILLE',False);

if (sTNatureAuxi='CLI') then
begin
  if Assigned(GetControl('YTC_TARIFSPECIAL')) then SetControlProperty('YTC_TARIFSPECIAL', 'Plus', 'YTP_FONCTIONNALITE="'+sTarifClient+'"');
  if Assigned(GetControl('YTC_COMMSPECIAL' )) then SetControlProperty('YTC_COMMSPECIAL' , 'Plus', 'YTP_FONCTIONNALITE="'+sCommissionClient+'"');
end
else if  (sTNatureAuxi='FOU') then
begin
  if Assigned(GetControl('YTC_TARIFSPECIAL')) then SetControlProperty('YTC_TARIFSPECIAL', 'Plus', 'YTP_FONCTIONNALITE="'+sTarifFournisseur+'"');
  if Assigned(GetControl('YTC_COMMSPECIAL' )) then SetControlProperty('YTC_COMMSPECIAL' , 'Plus', 'YTP_FONCTIONNALITE="'+sCommissionFournisseur+'"');
end;
if Assigned(GetControl('YTC_INCOTERM')) then
  thedit(GetControl('YTC_INCOTERM')).Onexit := YTC_INCOTERMEXIT;

//uniquement en line
//	CacheZonesNONLine;

TFFiche(ecran).OnAfterFormShow := AfterFormShow;

end;  // OnArgument


// Debut CHR
Procedure TOM_TIERS.MajAffichageKardexEntreprise;
var Bouton : TradioGroup;
BEGIN

    // cette fct est appelé depuis le script TIERS si modif le Kardex => Entreprise et reciproquement
  Bouton := TradioGroup(GetControl ('T_PARTICULIER'));
  If (Bouton.Itemindex = 0) Then
     AfficheEcranClientMode(TRUE)
  Else
      AfficheEcranClientMode(FALSE);
END;

Procedure TOM_TIERS.RecupContact;   //CHR
Var QLoc : TQuery;
BEGIN

     THEdit(GetControl('C_TELEPHONE')).enabled := False;
     THEdit(GetControl('C_NOM')).text := '';
     THEdit(GetControl('C_PRENOM')).text := '';
     THEdit(GetControl('C_TELEPHONE')).text := '';
     THEdit(GetControl('C_RVA')).text := '';

     Qloc := Nil ;
     if (client_particulier = True) and (THEdit(GetControl('C_NOM')).text = '') then
     begin
          Qloc:=OpenSQL('select C_NOM, C_PRENOM, C_TELEPHONE, C_PRINCIPAL, C_CIVILITE, C_NUMEROCONTACT from CONTACT where C_AUXILIAIRE="'+GetField('T_AUXILIAIRE')+'" and C_TYPECONTACT="T" and C_PRINCIPAL <>"X" AND C_FERME <> "X"',True) ;
          if (not Qloc.eof) then
          begin
               THEdit(GetControl('C_NOM')).text := QLoc.Fields[4].AsString + ' ' + QLoc.Fields[0].AsString + ' ' + QLoc.Fields[1].AsString;
               THEdit(GetControl('C_TELEPHONE')).text := QLoc.Fields[2].AsString;
               num_contact :=  QLoc.Fields[5].AsInteger;
               if (THEdit(GetControl('C_NOM')).text <> '') then
               THEdit(GetControl('C_TELEPHONE')).enabled := True;
          end;
     end;
     if (client_particulier <> True) and (THEdit(GetControl('C_NOM')).text = '') then
     begin
          Qloc:=OpenSQL('select C_NOM, C_PRENOM, C_TELEPHONE, C_PRINCIPAL, C_CIVILITE, C_NUMEROCONTACT from CONTACT where C_AUXILIAIRE="'+GetField('T_AUXILIAIRE')+'" and C_TYPECONTACT="T" and C_PRINCIPAL = "X" AND C_FERME <> "X"',True) ;
          if (not Qloc.eof) then
          begin
               THEdit(GetControl('C_NOM')).text := QLoc.Fields[4].AsString + ' ' + QLoc.Fields[0].AsString + ' ' + QLoc.Fields[1].AsString;
               THEdit(GetControl('C_TELEPHONE')).text := QLoc.Fields[2].AsString;
               num_contact :=  QLoc.Fields[5].AsInteger;
               if (THEdit(GetControl('C_NOM')).text <> '') then
                  THEdit(GetControl('C_TELEPHONE')).enabled := True;
          end;
     end;
     if Qloc <> nil then Ferme (QLoc) ;

END;


Procedure TOM_TIERS.MajContact_chr;      //CHR
begin
inherited ;

if (THEdit(GetControl('C_NOM')).text <> '') then
Begin
     ExecuteSQL('Update CONTACT SET C_TELEPHONE="'+THEdit(GetControl('C_TELEPHONE')).text+'" where C_AUXILIAIRE="'+GetField('T_AUXILIAIRE')+'" and C_TYPECONTACT="T" and C_NUMEROCONTACT=' + IntToStr(Num_Contact) +'') ;
End;

End;
// fin CHR

/////////////////////////////////////////////////////////////////////////////
function TOM_TIERS.RechercheTiers (Tiers : string; NatureTiers, stPayeur : string;
                                    var libelleTiers : string) : boolean;
var Q : TQuery;
    stSql,stNature,st,stVirgule : string;
begin
Result := True;
stNature:='';
stVirgule:='';
if Naturetiers<>'' then
  begin
  stNature:='(';
  repeat
  st:=Trim(ReadTokenSt(NatureTiers)) ;
  if st<>'' then
     begin
     stNature:=stNature+stVirgule+'"'+st+'"';
     stVirgule:=',';
     end;
  until st='';
     stNature:=stNature+')';
  end;

if Tiers <> '' then
    begin
    if stNature = '' then
        begin
        stSql := 'SELECT T_LIBELLE FROM TIERS WHERE T_TIERS="'+ Tiers +'"';
        end else
        begin
        stSql := 'SELECT T_LIBELLE FROM TIERS WHERE T_TIERS="'+ Tiers +
                    '" AND T_NATUREAUXI in ' + stNature ;
        end;
    if stPayeur <> '' then
        begin
        stSql := stSql + ' AND T_ISPAYEUR="' + stPayeur + '"';
        end;
    Q := OpenSql (stSql, True);
    if not Q.Eof then
        begin
        LibelleTiers := Q.Findfield('T_LIBELLE').AsString;
        end else Result := False;
    Ferme (Q);
    end;
end;

/////////////////////////////////////////////////////////////////////////////

procedure TOM_TIERS.AffecteTAuxiliaire(const StAuxiliaire : string; Toujours : Boolean);
Var Tampon : String;
begin

   //Fv1 : 18/09/2013 - FS#660 - DELABOUDINIERE : En duplication de client si attribution auto du code, Auxiliaire non renseigné.
   if ds<>nil then
   begin
      if (Ds.State=dsEdit) And (TFFiche(Ecran).fTypeAction = TaModif) then Exit;
   end;

   LgCode:=VH^.Cpta[fbAux].Lg ;
   Tampon:=Uppercase(Trim(StAuxiliaire));

   // FV1 : 09-10-2012 : Pas de modification du compte auxiliaire si tjrs = false (if (Toujours = True) or (Length(Tampon)<>LgCode) then)
   if (Toujours) Or (Length(Tampon)<>LgCode) then
   begin
    // BRL : 27/02/2015 : caractères spéciaux interdits en compta !!!
    Tampon:=SupprimeCaracteresSpeciaux(Tampon, false, false, false, true);

    if Length(Tampon)>LgCode then
    begin
      Tampon:=Copy(Tampon,1,LgCode);
    end else
    begin
    if Length(Tampon)<LgCode then
      Tampon:=BourreLaDonc(Tampon,fbAux);
    end;
    SetField('T_AUXILIAIRE',Tampon);
   end;

end;

/////////////////////////////////////////////////////////////////////////////

procedure TOM_TIERS.AddRemoveItemFromPopup(stPopup,stItem : string; bVisible : boolean);     //PCS
var pop : TPopupMenu ;
    i : integer;
    st : string;
begin
pop :=TPopupMenu(GetControl(stPopup) ) ;
if pop<> nil  then
   for i:=0 to pop.items.count-1 do
       begin
       st:=uppercase(pop.items[i].name);
       if st=stItem then  begin pop.items[i].visible:=bVisible; break; end;
       end;
end ;

procedure TOM_TIERS.Duplication (CleDuplic: string);
var TobForm : TOB;
    StTable,  StCleDuplic : String;
    //TagFicheRech : integer ;
    G_CodeTiers : THCritMaskEdit;
    QQ : TQuery ;
    stPlus : string;
    sIncoterm : String;
{$IFDEF GRC}
    stConfid : string;
{$ENDIF}
begin
//Duplic:= True ;
StCleDuplic:=CleDuplic;
{$IFDEF EAGLCLIENT}
StTable:='TIERS' ;
{$ELSE}
StTable := GetTableNameFromDataSet(DS);
{$ENDIF}
//TagFicheRech := 1; // pour DispatchRecherche
TobForm := TOB.Create (StTable, Nil, -1);
if TobForm <> Nil then
    begin
    if STCleDuplic = '' then
        begin
        if (DS.State=dsInsert) then  // si nouvelle fiche  recherche record à dupliquer
            begin
            // StCleDuplic:=DispatchRecherche( TagFicheRech, '');    // recherche de la fiche à dupliquer
            G_CodeTiers := THCritMaskEdit.Create (nil);
{$IFDEF GRC}
            if ctxGRC in V_PGI.PgiContexte then
               begin
{$IFDEF BTP}
               stConfid:='';
{$ELSE}
               if (ctxAffaire in V_PGI.PGIContexte) or (ctxMode in V_PGI.PGIContexte) then
                 stConfid := ''
               else
                 stConfid := RTXXWhereConfident('CON');
{$ENDIF}
               if stNatureauxi='FOU' then stPlus:='((T_NATUREAUXI="FOU") OR (T_NATUREAUXI="CLI" '+ stConfid+'))'
                  else if stNatureauxi='CON' then stPlus:='(T_NATUREAUXI="CON")'
                  else stPlus:='(((T_NATUREAUXI="CLI") OR (T_NATUREAUXI="PRO")) '+ stConfid+') OR (T_NATUREAUXI="FOU") ' ;
               end
            else
{$ENDIF}
            if stNatureauxi='FOU' then stPlus:='(T_NATUREAUXI="FOU" OR T_NATUREAUXI="CLI")'
               else if stNatureauxi='CON' then stPlus:='(T_NATUREAUXI="CON")'
               else stPlus:='(T_NATUREAUXI="CLI") OR (T_NATUREAUXI="PRO") OR (T_NATUREAUXI="FOU") ' ;
            DispatchRecherche (G_CodeTiers, 2, stPlus, '', '');
            StCleDuplic := G_CodeTiers.Text;
            QQ:=OPENSQL('SELECT * From TIERS WHERE T_TIERS="'+StCleDuplic+'"',True);
            if not TobForm.SelectDB ('', QQ) then StCleDuplic:='' ;
            Ferme(QQ);
            G_CodeTiers.Free;
            end else
            begin
            if not TobForm.SelectDB ('',TFFiche(Ecran).QFiche) then StCleDuplic:='' ;
            stCleDuplic:='*' ;
            TFFiche (Ecran).Bouge (NbInsert);
            end;
        end else
        begin
        if not TobForm.SelectDB ('"' + StCleDuplic + '"', Nil) then StCleDuplic:='' ;
        end;
    if StCleDuplic <> '' then
        begin
        ReInitDuplication(TobForm);
        TobToEcran(TobForm);
        ModifParticulier ;
         // TobForm.PutEcran (F); A mettre quand fonctionnement ok (bloc note affiché)
                             //mcd 05/03/02 pour tables libres
        if not TobZonelibre.SelectDB('"'+TiersAuxiliaire (stCleDuplic,false)+'"',Nil )
               then TobZonelibre.InitValeurs;
        TobZonelibre.PutEcran(TFfiche(Ecran));
        sIncoterm := RechDom('GCINCOTERM', TobZoneLibre.getValue('YTC_INCOTERM') , False);
        If sIncoterm = 'Error' then sIncoterm := '';
        SetControlCaption('TYTC_INCOTERM', sIncoterm);
        TobZoneLibre.SetAllModifie(True);
        end;
    end;

TobForm.Free;
SetActiveTabSheet (TFFiche(Ecran).pages.Pages[0].name);
SetField('T_AUXILIAIRE','');
SetControlText ('T_TIERS', ''); // DBR - suite fiche 10655 - pour avoir en duplication le message code obligatoire
SetFocusControl ('T_TIERS');
SetControlEnabled('BDUPLICATION',FALSE);
end;

procedure TOM_TIERS.tobToEcran (TobForm : TOB);
var i : integer ;
    CC : TControl ;
    CC_RO : Boolean ;
begin
        for i := 1 to TobForm.NbChamps do
            begin
            CC:=TControl(Ecran.findcomponent(TobForm.GetNomChamp(i))) ;
{$IFDEF EAGLCLIENT}
            CC_RO:=((CC is TCustomEdit) and (TEdit(CC).ReadOnly)) ;
{$ELSE}
            CC_RO:=((CC is TEdit) and (TEdit(CC).ReadOnly)) or ((CC is TDBEdit) and (TDBEdit(CC).ReadOnly)) ;
{$ENDIF}
            if (CC<>nil) and (CC.enabled and ( not CC_RO) and cc.Visible) then SetField (TobForm.GetNomChamp(i), TobForm.GetValeur (i));
            end;
end;

/////////////////////////////////////////////////////////////////////////////
procedure TOM_TIERS.SuspectToEcran ; // GRC- Met à l'écran infos provenant fiche suspect
var TM : TOM_TIERS;
begin
    TOBSuspect := TOB.Create('SUSPECTS', nil, -1);
    if not TobSuspect.selectDB ('"'+stSuspect+'"',Nil,False) then
       begin
       TOBSuspect.free;
       SetLastError(27, 'T_AUXILIAIRE');
       exit;
       end;
    TOBProspect:= TOB.Create('TIERS', nil, -1);
    TOBProspect.InitValeurs;
    TOBProspect.AddChampSup('TE_FACTURE',False);
    TOBProspect.AddChampSup('TE_PAYEUR',False);
    TM:=TOM_TIERS(CreateTOM('TIERS',Nil,False,False) );
    TM.stNatureAuxi := 'PRO';
    TM.InitTob(TOBProspect);
    TM.Free;
    TOBCopieChamp(TOBSuspect, TOBProspect);
//    if  not GetParamsoc('SO_GCNUMTIERSAUTO') then
//        TOBProspect.PutValue('T_TIERS', TobSuspect.GetValue('RSU_SUSPECT'));
    TOBProspect.putvalue('T_NATUREAUXI','PRO');
    TOBProspect.putvalue('T_DATEINTEGR', V_PGI.DateEntree);
    //TOBProspect.putvalue('T_PUBLIPOSTAGE','X');
    TOBSuspectCompl := TOB.Create('SUSPECTSCOMPL', nil, -1);
{$IFDEF GRC}
    if TOBSuspectCompl.selectDB ('"'+stSuspect+'"',Nil,False) then
        RTCorrespondSuspectProspect('RSC',TOBSuspectCompl, TobZonelibre);
{$ENDIF}
    if (TOBProspect.getvalue('T_PARTICULIER')='X') then client_particulier := true
    else client_particulier := false;
    AfficheEcranClientMode(client_particulier);
    TobToEcran(TOBProspect);
{$IFDEF EAGLCLIENT}
    TFFiche(Ecran).codename :='';
{$ENDIF}
    SetControlEnabled('BDUPLICATION',FALSE);
    TOBProspect.free;TOBProspect := NIL;
end;

procedure TOM_TIERS.ReInitDuplication(TobForm : TOB);
begin
TobForm.PutValue('T_TOTALDEBIT', 0);
TobForm.PutValue('T_TOTALCREDIT', 0);
TobForm.PutValue('T_DATEDERNMVT', IDate1900);
TobForm.PutValue('T_DEBITDERNMVT', 0);
TobForm.PutValue('T_CREDITDERNMVT', 0);
TobForm.PutValue('T_NUMDERNMVT', 0);
TobForm.PutValue('T_LIGNEDERNMVT', 0);
TobForm.PutValue('T_DERNLETTRAGE', '');
TobForm.PutValue('T_TOTDEBP', 0);
TobForm.PutValue('T_TOTCREP', 0);
TobForm.PutValue('T_TOTDEBE', 0);
TobForm.PutValue('T_TOTCREE', 0);
TobForm.PutValue('T_TOTDEBS', 0);
TobForm.PutValue('T_TOTCRES', 0);
TobForm.PutValue('T_TOTDEBANO', 0);
TobForm.PutValue('T_TOTCREANO', 0);
TobForm.PutValue('T_TOTDEBANON1', 0);
TobForm.PutValue('T_TOTCREANON1', 0);
TobForm.PutValue('T_DATEDERNRELEVE', IDate1900);
TobForm.PutValue('T_DATEDERNPIECE', IDate1900);
TobForm.PutValue('T_NUMDERNPIECE', 0);
TobForm.PutValue('T_TOTDERNPIECE', 0);
TobForm.PutValue('T_DATECREATION', Date);       // ajout PCS 28052002
TobForm.PutValue('T_DATEMODIF', Date);          // ajout PCS 28052002
TobForm.PutValue('T_UTILISATEUR', V_PGI.User);   // ajout PCS 28052002
end;

{$IFDEF MODE}
/////////////////////////////////////////////////////////////////////////////
function ChangeBase ( Prefixe, Valeur : String ; TailleMax : Integer ) : String ;
Var Ind  : Byte ;
    iVal : Longint ;
BEGIN
if IsNumeric(Valeur) then
   BEGIN
   // conversion d'une valeur numérique en base décimale en base 36
   Result := '' ;
   iVal := StrToInt(Valeur) ;
   while iVal <> 0 do
      BEGIN
      Ind := iVal MOD 36 ;
      if Ind < 10 then Result := Chr(Ord('0') +Ind) + Result
                  else Result := Chr(Ord('A') +Ind -10) + Result ;
      iVal := iVal DIV 36 ;
      END ;
   Result := Prefixe + StringOfChar('0', TailleMax -Length(Prefixe) -Length(Result)) + Result ;
   END else
   BEGIN
   Result := Prefixe + Valeur ;
   END ;
END ;
{$ENDIF}

procedure TOM_TIERS.OnChangeField(F: TField);
var tampon    : string;
    tampon1   : string;
    i : integer;
    QR : TQUERY;
    CodeTiers : string;
    Etat      : string;
begin
Inherited;

{$IFDEF CTI}
  if CtiErreurConnexion then RTCTIAfficheMessageCTI;
  if CtiRaccrocherFiche then SetControlText('INFOSCTI','le correspondant a raccroché');
{$ENDIF}

  if (ds.state = dsinsert) then
   begin

  //FV1 : 10/12/2014 - FS#1175 - MULTIPHONE NETCOM - Interdire les caractères spéciaux dans les codes tiers
  if (F.FieldName = 'T_TIERS') and (GetField('T_TIERS')<>'') then
  begin
    CodeTiers := GetField('T_TIERS');
    if SupprimeCaracteresSpeciaux(CodeTiers,'',False) then
    begin
      SetField('T_TIERS', CodeTiers);
      SetLastError (59, 'T_TIERS');
      exit;
    end;
  end;

//   if (F.FieldName = 'T_TIERS') and (Presence('TIERS','T_TIERS',GetField ('T_TIERS'))) then
   if (F.FieldName = 'T_TIERS') and (PresenceComplexe('TIERS',['T_NATUREAUXI','T_TIERS'],['=','='],[stNatureAuxi,GetField ('T_TIERS')],['S','S'])) then
   begin
     PgiBox (TexteMessage[17]);
     SetField('T_TIERS','');
     SetFocusControl ('T_TIERS');
     exit ;
   end ;

   // prise en compte nouvelle option pour racine compte ..
{$IFDEF MODE}
   if (F.FieldName = 'T_TIERS') and (GetField('T_AUXILIAIRE')='') and (GetField('T_TIERS')<>'') then
      begin
      if (stNatureAuxi = 'FOU') then AffecteTAuxiliaire(Trim(GetParamSoc('SO_GCPREFIXEAUXIFOU'))+GetField('T_TIERS'), True)
      else if (stNatureAuxi = 'CLI') or (stNatureAuxi = 'PRO') then
         begin
         Tampon := Trim(GetParamSoc('SO_GCPREFIXEAUXI'))+GetField('T_TIERS') ;
         if Length(Tampon)>VH^.Cpta[fbAux].Lg then
           begin
           Tampon := ChangeBase(Trim(GetParamSoc('SO_GCPREFIXEAUXI')), GetField('T_TIERS'), VH^.Cpta[fbAux].Lg) ;
           end ;
         AffecteTAuxiliaire(Tampon, True) ;
         end
      else
        AffecteTAuxiliaire(GetField('T_TIERS'), True) ;
      end;
{$ELSE}
      if (F.FieldName = 'T_TIERS') and (GetField('T_AUXILIAIRE')='') and (GetField('T_TIERS')<>'') then
      begin
      if (stNatureAuxi = 'FOU') then
        AffecteTAuxiliaire(Trim(GetParamSoc('SO_GCPREFIXEAUXIFOU'))+GetField('T_TIERS'), True)
      else if (stNatureAuxi = 'CLI') or (stNatureAuxi = 'PRO') then
        AffecteTAuxiliaire(Trim(GetParamSoc('SO_GCPREFIXEAUXI'))+GetField('T_TIERS'), True)
      else
        AffecteTAuxiliaire(GetField('T_TIERS'), True)
                   end;
{$ENDIF}

   if (F.FieldName = 'T_AUXILIAIRE') and (GetField('T_AUXILIAIRE')<>'') then AffecteTAuxiliaire(GetField('T_AUXILIAIRE'), False);
   end;

  if (F.FieldName = 'T_COLLECTIF')  then
  begin
    Tampon := GetField('T_COLLECTIF');
    tampon1:=Trim(tampon);
    if Length(tampon1)>VH^.Cpta[fbGene].Lg then tampon1:=Copy(Tampon1,1,VH^.Cpta[fbGene].Lg) ;
    if Length(tampon1)<VH^.Cpta[fbGene].Lg then
    begin
      for i:=Length(tampon1) to VH^.Cpta[fbGene].Lg do tampon1:=tampon1+VH^.Cpta[fbGene].Cb ;
      end ;
    if Tampon<>tampon1 then SetField('T_COLLECTIF',tampon1);
  end;

if (F.FieldName = 'T_ISPAYEUR') then
begin
  if GetField ('T_ISPAYEUR')='X' then
  begin
    SetControlEnabled('TT_PAYEUR',FALSE);
    if StNatureAuxi='CLI' then
    begin
      SetCOntrolText('TE_PAYEUR','');
      SetControlProperty('TE_PAYEUR','Enabled', False);
    end
  end
  else
  begin
    SetControlEnabled('TT_PAYEUR',TRUE);
      if StNatureAuxi='CLI' then
        SetControlProperty('TE_PAYEUR','Enabled', True);
  end;
end;

{$IFNDEF CHR}
if (DS.State in [dsInsert,dsEdit]) and (F.FieldName = 'T_SOCIETEGROUPE') then
    begin
    if  (GetField ('T_SOCIETEGROUPE')<>'') and (GetField ('T_SOCIETEGROUPE') <> GetField ('T_TIERS')) then
        begin
        AddRemoveItemFromPopup ('POPM','MNSOCMERE',True);
        AddRemoveItemFromPopup ('POPM','MNGROUPE',True);
        end else
        begin
        AddRemoveItemFromPopup ('POPM','MNSOCMERE',False);
        AddRemoveItemFromPopup ('POPM','MNGROUPE', False);
        end;
    end;
if (DS.State in [dsInsert,dsEdit]) and (F.FieldName = 'T_PRESCRIPTEUR') then
    begin
    if (GetField ('T_PRESCRIPTEUR')<>'') and (GetField ('T_PRESCRIPTEUR') <> GetField ('T_TIERS'))
       then AddRemoveItemFromPopup ('POPM','MNPRESCRIPTEUR',True)
       else AddRemoveItemFromPopup ('POPM','MNPRESCRIPTEUR',False);
    end;
{$ENDIF}
if (F.FieldName = 'T_RELEVEFACTURE') then
   begin
  if (GetField('T_RELEVEFACTURE') = 'X') then
     begin
     SetControlProperty('T_FREQRELEVE','Enabled',TRUE);SetControlProperty('T_FREQRELEVE','Color',ClWindow);
     SetControlProperty('T_JOURRELEVE','Enabled',TRUE);SetControlProperty('T_JOURRELEVE','Color',ClWindow);
     end
  else
     begin
     SetControlProperty('T_FREQRELEVE','Enabled',FALSE);SetControlProperty('T_FREQRELEVE','Color',clBtnFace);
     SetControlProperty('T_JOURRELEVE','Enabled',FALSE);SetControlProperty('T_JOURRELEVE','Color',clBtnFace);
     end;
   end;
if ((F.FieldName = 'T_PAYS') and (GetField('T_NATIONALITE')='')) then
   SetField('T_NATIONALITE',GetField('T_PAYS'));
// grc
if not CritereSuspect then
if ((F.FieldName = 'T_PAYS') and (GetField('T_LANGUE')='')) then   begin
  Qr := OpenSql('SELECT PY_LANGUE FROM PAYS WHERE PY_PAYS="'+GetField('T_PAYS')+'"', False);
  if not Qr.Eof then SetField('T_LANGUE', Qr.Findfield('PY_LANGUE').AsString);
  Ferme(Qr);
   end;

if F.FieldName='T_MULTIDEVISE' then
  begin
  if GetField('T_MULTIDEVISE')='X' then
    begin
    SetField('T_DEVISE','');
    SetControlEnabled('T_DEVISE',FALSE);
    end else
    begin
    SetControlEnabled('T_DEVISE',TRUE);
    if GetField('T_DEVISE')='' then SetField('T_DEVISE', V_PGI.DevisePivot);
    end;
  end;

  if F.FieldName='T_FERME' then
   begin
   if (GetField('T_FERME')='X') and ( ds.State <> dsBrowse ) then SetField ('T_DATEFERMETURE', V_PGI.DateEntree);
   AffIndisponible (False);
   end;

  if FormaterLesZones and ((F.FieldName='T_LIBELLE') or (F.FieldName='T_PRENOM') or  (F.FieldName='T_ADRESSE1') or (F.FieldName='T_ADRESSE2') or (F.FieldName='T_ADRESSE3') or (F.FieldName='T_VILLE')) then
   FormatelesZones(F.FieldName);

  if (F.FieldName = 'T_PAYS') and (GetField('T_PAYS')<>'') then
    SetControlProperty('T_REGION','PLUS','RG_PAYS="'+GetControlText('T_PAYS')+'"');

end;

// formatage des zones du client //
procedure TOM_TIERS.FormatelesZones(NomZone : string);
var texte : string;
    Ind : integer ;
begin

  if NomZone = '' then
  Begin
  for Ind := Low(ChpAFormater) to High(ChpAFormater) do
    Begin
    texte := VerifAppliqueFormat(ChpAFormater[Ind], GetField (ChpAFormater[Ind]), client_particulier);
    if (texte <> GetField (ChpAFormater[Ind])) then SetField(ChpAFormater[Ind], texte);
    End ;
  End
  else
  Begin
  for Ind := Low(ChpAFormater) to High(ChpAFormater) do
    Begin
    if ChpAFormater[Ind] = NomZone then
      Begin
      texte := VerifAppliqueFormat(ChpAFormater[Ind], GetField (ChpAFormater[Ind]), client_particulier);
      if Length(Texte)>35 then Texte:=Copy(Texte,1,35);
      if (texte <> GetField (ChpAFormater[Ind])) then SetField(ChpAFormater[Ind], texte);
      Break ;
      End ;
    End ;
  End ;

end;

/////////////////////////////////////////////////////////////////////////////
procedure TOM_TIERS.OnAfterUpdateRecord;
var codeper: integer;
begin
Inherited;
        // Appel fct pour tiers  -> DP , sauf si fiche appeler depuis fiche annuaire
if (stNatureAuxi = 'CLI') or (stNatureAuxi = 'FOU')then begin
if (ctxscot in V_PGI.PGIContexte) and (GetParamsoc('SO_AFLIENDP') = true) and (origine<>'DP')
        then  begin
        codeper:=   StrToInt(GetControlText('XXCodeper'));
{$IFDEF DP}
        DpSynchro(FALSE,codeper, CodeTiersDP, stNatureAuxi,Nil);      //MCD 07/08/00
{$IFDEF EAGLCLIENT}
{$ELSE}
        SynchroniseParamSoc (codeper);           //mcd 11/10/00
{$ENDIF}
{$endif}
        SetControlText ('XXCODEPER',IntToSTR(codeper));
        end;
if (ctxChr in V_PGI.PGIContexte) then MajContact_chr; //CHR
end;

// Mise à jour du premier contact de manière automatique SD
majcontact;

If OnFerme then Ecran.Close;
{PAUL   if (not (ctxscot in V_PGI.PGIContexte)) and
   (not (ctxmode in V_PGI.PGIContexte)) and
   (not (ctxaffaire in V_PGI.PGIContexte)) then
    begin
    if OldInsert then
        begin
        SetActiveTabSheet('PGeneral') ;
        TFFiche(Ecran).Bouge(nbInsert) ;
        end;
    end;}
end;

// SD : création ou modification du contact numéro 1 pour les clients particuliers.
procedure TOM_TIERS.majcontact;
var TOB_CON:TOB;
begin
inherited ;

  if client_particulier = False then exit
  else if ctxscot in V_PGI.PGIContexte then exit;

  TOB_CON:=TOB.Create('CONTACT', NIL, -1);

  // Création de la TOB des contacts
  //  Le contact numéro 1 est modifié lors de la modif de la fiche du client particulier
  TOB_CON.PutValue('C_TYPECONTACT','T') ;
  TOB_CON.PutValue('C_AUXILIAIRE',GetField('T_AUXILIAIRE')) ;
  TOB_CON.PutValue('C_TIERS',GetField('T_TIERS')) ;
  TOB_CON.PutValue('C_NUMEROCONTACT','1') ;
  TOB_CON.PutValue('C_NATUREAUXI','CLI') ;
  TOB_CON.PutValue('C_CIVILITE',GetField('T_JURIDIQUE')) ;
  TOB_CON.PutValue('C_NOM',GetField('T_LIBELLE')) ;
  TOB_CON.PutValue('C_PRENOM',GetField('T_PRENOM')) ;
  TOB_CON.PutValue('C_TELEPHONE',GetField('T_TELEPHONE')) ;
  TOB_CON.PutValue('C_TELEX',GetField('T_TELEX')) ;
  TOB_CON.PutValue('C_FAX',GetField('T_FAX')) ;
  TOB_CON.PutValue('C_RVA',GetField('T_RVA')) ;
  TOB_CON.PutValue('C_JOURNAIS',GetField('T_JOURNAISSANCE')) ;
  TOB_CON.PutValue('C_MOISNAIS',GetField('T_MOISNAISSANCE')) ;
  TOB_CON.PutValue('C_ANNEENAIS',GetField('T_ANNEENAISSANCE')) ;
  TOB_CON.PutValue('C_SEXE',GetField('T_SEXE')) ;
  TOB_CON.PutValue('C_PRINCIPAL','X') ;
  TOB_CON.PutValue('C_CLETELEPHONE',GetField('T_CLETELEPHONE')) ;
  TOB_CON.PutValue('C_CLEFAX',GetField('T_CLEFAX')) ;

  //FV1 : 03/11/2015 - FS#1771 - CLOSSUR : pb de mise à jour fiche contact sur client type particulier
  TOB_CON.SetAllModifie(True);

  TOB_CON.InsertOrUpdateDB(False);

  TOB_CON.free;
  
end;

Function TOM_TIERS.LimiteVersDemo_Mode : Boolean;
var QQ :TQuery ;
    SQL : string ;
    Nbr : integer;
begin
Result := False;
if (ds<>nil) and not(DS.State in [dsInsert]) then exit;
if (ctxMode in V_PGI.PGIContexte) and (V_PGI.VersionDemo = True) then
  begin
  SQL:='SELECT COUNT(*) FROM TIERS' ;
  QQ:=OpenSQL(SQL,True) ;
  if Not QQ.EOF then Nbr:=QQ.Fields[0].AsInteger else Nbr:=10;
  Ferme(QQ) ;
  if (Nbr >= 10) then Result := True;
  end;
end;

procedure TOM_TIERS.OnUpdateRecord;
var CodeTiers, NumChrono, LibelleTiers, NatureTiers, ret, stNat : String;
    Nbr : Integer;
    Qr : TQuery;
    ExisteTiers,ExisteOblig : boolean;
    TOB_CON :TOB;
   // ChainageInfosCompl : Boolean; // grc
{$IFDEF MODE}
    Tampon : String;
{$ENDIF}
	LePays : string;
    Etat   : string;
begin
  Inherited;

  {$IFDEF GRC}
  if stNatureAuxi <>'CON' then
  Begin
    if  (ctxGRC in V_PGI.PGIContexte) and
        (ds<>nil)                     and not
        RTControleModifTiers(TFFiche(Ecran),nil,(DS.State<>dsInsert)) then
    begin
      LastError :=1;
      exit;
    end;
  end;
  {$ENDIF}

  if (ds<>nil) and (DS.State in [dsInsert]) then
  begin
    if V_PGI.VersionDemo then
        begin
        Qr:=OpenSQL('SELECT COUNT(*) FROM TIERS',True) ;
      if Not Qr.EOF then
        Nbr:=Qr.Fields[0].AsInteger
      else
        Nbr:=0;
        Ferme(Qr) ;
        if (Nbr >= 100) then
            begin
            LastError:=32;
            LastErrorMsg:=TexteMessage[LastError] ;
            exit;
            end;
        end;
    end;

  If LimiteVersDemo_Mode then begin SetLastError(28, 'T_LIBELLE'); exit ; end ;

  //  memorisation de ds.state=insert pour boucle sur insertion
  //OldInsert := (DS.State=dsInsert);
    // Mise à jour pour consolidation
  if (GetField ('T_CREERPAR')= '')  then SetField ('T_CREERPAR', 'SAI');
  if (GetField ('T_DEVISE')= '')    then SetField ('T_DEVISE', V_PGI.devisepivot);
  if (GetField ('T_LETTRABLE')= '') then SetField ('T_LETTRABLE', 'X');
  if (GetField ('T_PAYS')= '')      then SetField ('T_PAYS', GetParamSoc('SO_GcTiersPays'));

  If (GetField ('T_PAYS') <> '') and (GetField ('T_REGION') <> '') and (not existesql('SELECT RG_PAYS FROM REGION WHERE RG_PAYS = "'+string(GetField ('T_PAYS'))+'" AND RG_REGION = "'+string(GetField ('T_REGION'))+'"')) then
  Begin
  SetLastError(38, 'T_REGION');
  exit ;
  end ;

  { Code NIF non obligatoire pour l'instant mais peut-être un jour ...}
{
if (GetField ('T_NIF')='') and (GetField('T_PARTICULIER')='-') then
begin
  SetLastError(41, 'T_NIF');
  exit ;
end ;
}

  If (GetField ('T_NATUREAUXI')= '') then begin SetLastError(16, ''); exit ;end ;

// ==== gestion Tiers Clients ou Prospects ====================================================================
  if (GetField ('T_NATUREAUXI') = 'CLI') or
     (GetField ('T_NATUREAUXI') = 'PRO') or
     (GetField ('T_NATUREAUXI') = 'CON') then  // GRC
   begin
   if (GetField ('T_LIBELLE')= '') then begin SetLastError(2, 'T_LIBELLE'); exit ; end;
   if not (ctxMode in V_PGI.PGIContexte) then
      begin
      //If (GetField ('T_ADRESSE1')= '') then begin SetLastError(20, 'T_ADRESSE1'); exit ; end ;  { A PARALMETRER }
      If (GetField ('T_CODEPOSTAL')= '') then begin SetLastError(9, 'T_CODEPOSTAL'); exit ; end ;
      If (GetField ('T_VILLE')= '') then begin SetLastError(10, 'T_VILLE'); exit ; end ;
      end;

   SetField('T_PHONETIQUE',phoneme(Getfield('T_LIBELLE')));

   //Chargement du N° de téléphone sans Car. Spéciaux
   if GetField('T_TELEPHONE') <> '' then
	    SetField('T_CLETELEPHONE',CleTelephone(GetField('T_TELEPHONE')))
   Else
	    SetField('T_CLETELEPHONE', '');
    //
   if GetField('T_FAX') <> '' then
	    SetField('T_CLEFAX',CleTelephone(GetField('T_FAX')))
   Else
	    SetField('T_CLEFAX', '');
    //
   if GetField('T_TELEPHONE2') <> '' then
	    SetField('T_CLETELEPHONE2',CleTelephone(GetField('T_TELEPHONE2')))
   Else
	    SetField('T_CLETELEPHONE2', '');
    //
   if GetField('T_TELEPHONE3') <> '' then
	    SetField('T_CLETELEPHONE3',CleTelephone(GetField('T_TELEPHONE3')))
   Else
	    SetField('T_CLETELEPHONE3', '');

   //**************************** Gestion des chronos automatiques *********************
   SetField('T_TIERS',Trim(GetField('T_TIERS')));
   if (GetParamsoc('SO_GCNUMTIERSAUTO') = FALSE ) and (GetField ('T_TIERS')= '') then
      begin
      SetLastError(1, 'T_TIERS');
      exit ;
      end;
    //
    if (ds<>nil) then if client_particulier = TRUE then
      SetField ('T_PARTICULIER','X')
    else
      SetField ('T_PARTICULIER','-');
    //
   if (GetParamsoc('SO_GCFORMATEZONECLI') = TRUE )then FormatelesZones('');

   if GetParamsoc('SO_GCNUMTIERSAUTO') = TRUE then
      begin
      if (GetField ('T_TIERS')= '') then
         begin
{$IFDEF NOMADE}  //Si nomade, code tiers avec préfixe obligatoire
         CodeTiers:=AttribNewCode('TIERS','T_TIERS',GetParamsoc('SO_GCLGNUMTIERS'),trim(GetParamsoc('SO_GCPREFIXETIERS')),GetParamsoc('SO_GCCOMPTEURTIERS'),'');
{$ELSE}
         if ((ctxMode in V_PGI.PGIContexte) or (ctxChr in V_PGI.PGIContexte)) then
            CodeTiers:=AttribNewCode('TIERS','T_TIERS',GetParamsoc('SO_GCLGNUMTIERS'),trim(GetParamsoc('SO_GCPREFIXETIERS')),GetParamsoc('SO_GCCOMPTEURTIERS'),'')
         else
            CodeTiers:=AttribNewCode('TIERS','T_TIERS',-1,'',GetParamsoc('SO_GCCOMPTEURTIERS'),'');
{$ENDIF}//Nomade
         SetField ('T_TIERS', CodeTiers);
         if (ds<>nil) and (DS.State=dsInsert) then
           begin
{$IFDEF EAGLCLIENT}
           SetControlText('T_TIERS',CodeTiers);
{$ENDIF}
           end;
         NumChrono:=ExtraitChronoCode(CodeTiers);
         SetParamSoc('SO_GCCOMPTEURTIERS', NumChrono) ;
         end;
      end;

   if (ds<>nil) and (DS.State=dsInsert) and Presence('TIERS','T_TIERS',GetField ('T_TIERS')) then
      begin
      SetLastError(17, 'T_TIERS');
      exit ;
      end ;

// JT - eQualité 10384
   if ((OldTiersOnCreat <> '') and (OldTiersOnCreat <> GetField('T_TIERS')) OR DuplicFromMenu )and
      (GetField('T_AUXILIAIRE') <> '') then
   begin
     if PGIAsk('Le code Tiers a été changé.'+Chr(13)+
               'Voulez-vous réaffecter le code auxiliaire ?'+Chr(13)+
               '(ancien auxiliaire : '+GetField('T_AUXILIAIRE')+')',
        TFFiche(ecran).Caption) = mrYes then
         SetField('T_AUXILIAIRE', '');
   end;

   // Gestion du compte auxiliaire
{$IFDEF MODE}
   if (GetField ('T_AUXILIAIRE')= '') then
      begin
      Tampon := Trim(GetParamSoc('SO_GCPREFIXEAUXI'))+GetField('T_TIERS') ;
      if Length(Tampon)>VH^.Cpta[fbAux].Lg then
        begin
        Tampon := ChangeBase(Trim(GetParamSoc('SO_GCPREFIXEAUXI')), GetField('T_TIERS'), VH^.Cpta[fbAux].Lg) ;
        end ;
      AffecteTAuxiliaire(Tampon, True) ;
      end else AffecteTAuxiliaire(GetField('T_AUXILIAIRE'), False);
{$ELSE}
   SetField('T_AUXILIAIRE',Uppercase(Trim(GetField('T_AUXILIAIRE'))));
   if (GetField ('T_AUXILIAIRE')= '') then AffecteTAuxiliaire(Trim(GetParamSoc('SO_GCPREFIXEAUXI'))+GetField('T_TIERS'), True)
                                      else AffecteTAuxiliaire(GetField('T_AUXILIAIRE'), False);
{$ENDIF}



   if (ds<>nil) and (DS.State=dsInsert) then
      begin
{$IFDEF EAGLCLIENT}
      SetControlText('T_AUXILIAIRE',GetField ('T_AUXILIAIRE'));
{$ENDIF}
      if Presence('TIERS','T_AUXILIAIRE',GetField ('T_AUXILIAIRE')) then
         begin
         SetLastError(21, 'T_AUXILIAIRE');
         exit ;
         end ;
      end;

  if (StNatureAuxi='CLI') or (StNatureAuxi='PRO') then
  begin
    if (GetControlText ('TE_FACTURE')= '') then SetField('T_FACTURE',GetField('T_AUXILIAIRE'))
    else begin
        Qr := OpenSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+GetControlText('TE_FACTURE')+'"', False);
        if not Qr.Eof then SetField('T_FACTURE', Qr.Findfield('T_AUXILIAIRE').AsString) else begin SetLastError(5, 'TE_FACTURE');Ferme(Qr); exit ; end;
        Ferme(Qr);
        end;
   if GetControlText ('TE_PAYEUR')=GetField ('T_TIERS') then SetControlText('TE_PAYEUR', '');
   if (GetControlText ('TE_PAYEUR')<>'') then
      begin
      Qr := OpenSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+GetControlText('TE_PAYEUR')+'"', False);
      if not Qr.Eof then SetField('T_PAYEUR', Qr.Findfield('T_AUXILIAIRE').AsString);
      Ferme(Qr);
      end else
      begin     //mcd 05/03/02
      if (old_payeur <>'') then SetField('T_PAYEUR','');
      end;

    if (GetControlText('TE_LIVRE') = '') then
      TobZoneLibre.PutValue('YTC_TIERSLIVRE', GetField('T_AUXILIAIRE'))
    else
    begin
      Qr := OpenSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="' + GetControlText('TE_LIVRE') + '"', False);
      if not Qr.Eof then
         TobZoneLibre.PutValue('YTC_TIERSLIVRE', Qr.Findfield('T_AUXILIAIRE').AsString)
      else
      begin
         SetLastError(37, 'TE_LIVRE');
         Ferme(Qr);
         EXIT;
      end;
      Ferme(Qr);
    end;
  end;
   if stNatureAuxi ='CLI'  then NatureTiers:= stNatureAuxi+';PRO'
       else if stNatureAuxi ='PRO' then NatureTiers:= stNatureAuxi+';CLI'
           else NatureTiers:= stNatureAuxi;
   if  (GetField ('T_SOCIETEGROUPE')<>'') and (GetField ('T_SOCIETEGROUPE') <> GetField ('T_TIERS')) then
      begin
      ExisteTiers := RechercheTiers (getfield ('T_SOCIETEGROUPE'), NatureTiers, '',LibelleTiers);
      if not ExisteTiers then  begin SetLastError(25, 'T_SOCIETEGROUPE'); exit ; end ;
      end;
   if (GetField ('T_PRESCRIPTEUR')<>'') and (GetField ('T_PRESCRIPTEUR') <> GetField ('T_TIERS')) then
      begin
      ExisteTiers := RechercheTiers (getfield ('T_PRESCRIPTEUR'), NatureTiers, '',
                                   LibelleTiers);
      if not ExisteTiers then  begin SetLastError(26, 'T_PRESCRIPTEUR'); exit ; end ;
      end;
   if (GetField ('T_APPORTEUR') <> GetField ('T_AUXILIAIRE')) then
      begin
      ExisteTiers := RechercheTiers (getfield ('T_APPORTEUR'), '', '',LibelleTiers);
      if not ExisteTiers then  begin SetLastError(6, 'T_APPORTEUR'); exit ; end ;
      end;
   if (GetField('T_CREDITPLAFOND') < GetField ('T_CREDITACCORDE')) then
      begin
      SetLastError(13, 'T_CREDITPLAFOND');
      exit;
      end;

   if (GetField('T_REGIMETVA')='') then SetField('T_REGIMETVA',VH^.RegimeDefaut);
   if (GetField('T_COLLECTIF')='') then SetField('T_COLLECTIF',VH^.DefautCli);
   if (GetField('T_MODEREGLE')='') then SetField('T_MODEREGLE',GetParamSoc('SO_GcModeRegleDefaut'));
      // mcd 19/07/01, en liaison avec le DP/jur, il ne faut pas effacer le code siret, même sur un
      // particulier. L'info peut exister...
   if (ctxscot in V_PGI.PGIContexte) and (GetParamsoc('SO_AFLIENDP') =true)  then
   else if (GetField('T_PARTICULIER')='X') and (GetField('T_SIRET') <> '') then SetField('T_SIRET','') ;
   { mng : déplacement du controle Siret qui doit également se faire sur les fournisseurs}
   end;
   // Fin de gestion des clients

   if ( GetField('T_SIRET') <> '' ) and
      ( Pos(GetField('T_PAYS'),GetParamSoc('SO_GCPAYSSIRET')) <> 0 ) then
   begin
     if not VerifSiret( GetField('T_SIRET')) then  begin SetLastError(29, 'T_SIRET'); exit ; end;
     { mng : ajout condition sur la nature : exemple SIC même tiers en client et fournisseur }
     if ( GetField('T_NATUREAUXI')='CLI' ) or ( GetField('T_NATUREAUXI')='PRO' ) then
        stNat:=' and ( (T_NATUREAUXI="PRO") or (T_NATUREAUXI="CLI"))';
     if ( GetField('T_NATUREAUXI')='CON' ) then
        stNat:=' and (T_NATUREAUXI="CON")';
     if ( GetField('T_NATUREAUXI')='FOU' ) then
        stNat:=' and (T_NATUREAUXI="FOU")';

     Qr := OpenSql('SELECT T_TIERS,T_AUXILIAIRE FROM TIERS WHERE T_SIRET LIKE "'+GetField('T_SIRET')+
         '%" and T_AUXILIAIRE<>"'+GetField('T_AUXILIAIRE')+'"'+stNat, False);
     if not Qr.Eof then
          begin
          // mng : si Siret, refus doublon, si Siren, suivant ParamSoc
          if length(GetField('T_SIRET'))<>9 then
              begin
              LastError:=30;
              LastErrorMsg:=TexteMessage[LastError]+Qr.Findfield('T_TIERS').AsString;
              Ferme(Qr);
              exit;
              end;
          if GetParamsoc('SO_GCCONTROLESIREN') = 'BLO' then
              begin
              LastError:=34;
              LastErrorMsg:=TexteMessage[LastError]+Qr.Findfield('T_TIERS').AsString;
              Ferme(Qr);
              exit;
              end;
          if GetParamsoc('SO_GCCONTROLESIREN') = 'AVE' then
              begin
              if (PGIAsk('Ce code Siren existe déjà pour le tiers ' +
                  Qr.Findfield('T_TIERS').AsString + chr(13) +
                  ' Confirmez-vous ce code ?', TFFiche(ecran).Caption)<>mrYes) then
                  begin
                  Ferme(Qr);
                  LastError:=1; LastErrorMsg:='';
                  exit;
                  end;
              end
          else if GetParamsoc('SO_GCCONTROLESIREN') = 'ALD' then
              begin
              if (PGIAsk('Ce code Siren est déjà utilisé' + chr(13) +
                  ' Voulez-vous consulter la liste des doublons ?', TFFiche(ecran).Caption)=mrYes) then
                  begin
                  ret:=AGLLanceFiche('RT','RTDOUBLONSSIREN','T_SIRET='+GetField('T_SIRET'),'','');
                  if ret<>'ok' then
                      begin
                      Ferme(Qr);
                      LastError:=1; LastErrorMsg:='';
                      exit;
                      end  ;
                  end;
              end
          else
              begin
              LastError:=34;
              LastErrorMsg:=TexteMessage[LastError]+Qr.Findfield('T_TIERS').AsString;
              Ferme(Qr);
              exit;
              end;
          end;
     Ferme(Qr);
   end;
   TOBTIERSBTP.GetEcran(ecran); 
   TOBTIERSBTP.SetAllModifie(true);
   TOBTIERSBTP.InsertOrUpdateDB(false);
     
   // Gestion des champs obligatoires Clients et Prospect en attendant mieux....)
   // grc : pas de contôle sur les concurrents
   // mcd 20/09/02 ce test est déplacé, était mis au dessus end précdent ==> n'étais pas fait pour FOU
  if GetField ('T_NATUREAUXI') <> 'CON' then
        if ChampsObligatoires then exit;

  {Contrôle n° adresse livraison}
  if (Valeur(GetControlText('YTC_NADRESSELIV')) <> 0) and (not LookupvalueExist(thedit(getControl('YTC_NADRESSELIV')))) then
  begin
    SetLastError(47, 'YTC_NADRESSELIV');
    exit;
  end;

  {Contrôle n° adresse facturation}
  if (Valeur(GetControlText('YTC_NADRESSEFAC')) <> 0) and (not LookupvalueExist(thedit(getControl('YTC_NADRESSEFAC')))) then
  begin
   SetLastError(48, 'YTC_NADRESSEFAC');
   exit;
  end;

// gestion Tiers Fournisseurs ====================================================================

  if (StNatureAuxi='FOU') then
  begin
    if not Gestionfournisseur then exit;
      end;
// FIN Ajout LS

if (GetField ('T_NATUREAUXI') <> 'CLI') and (GetField ('T_NATUREAUXI') <> 'PRO')
    and (GetField ('T_NATUREAUXI') <> 'CON') then // GRC
   begin
   If (GetField ('T_LIBELLE')= '') then begin SetLastError(2, 'T_LIBELLE'); exit ; end;
   if not (ctxMode in V_PGI.PGIContexte) then
      begin
      If (GetField ('T_ADRESSE1')= '') then   begin SetLastError(20, 'T_ADRESSE1'); exit ; end ;
      If (GetField ('T_CODEPOSTAL')= '') then  begin SetLastError(9, 'T_CODEPOSTAL'); exit ; end ;
      If (GetField ('T_VILLE')= '') then   begin SetLastError(10, 'T_VILLE'); exit ; end ;
      end;

  // (Fonctionnement identique à ce qui est fait / TE_FACTURE et TE_PAYEUR pour le client)
{   if (GetControlText ('T_FACTURE')= '') then SetField('T_FACTURE',GetField('T_AUXILIAIRE'))
   else begin         }
   if (GetControlText ('T_NATUREAUXI') <> 'FOU') then
   begin
     if (GetControlText ('T_FACTURE')<> '') then
     begin
          Qr := OpenSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+GetControlText('TE_FACTURE')+'"', False);
          if not Qr.Eof then SetField('T_FACTURE', Qr.Findfield('T_AUXILIAIRE').AsString) else begin SetLastError(40, 'T_FACTURE');Ferme(Qr); exit ; end;
          Ferme(Qr);
     end;
     if GetControlText ('T_PAYEUR')=GetField ('T_TIERS') then SetControlText('T_PAYEUR', '') ;
     if (GetControlText ('T_PAYEUR')<>'') then
     begin
        Qr := OpenSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+GetControlText('T_PAYEUR')+'" AND T_ISPAYEUR="X"' , False);
        if not Qr.Eof then SetField('T_PAYEUR', Qr.Findfield('T_AUXILIAIRE').AsString) else begin SetLastError(39, 'T_PAYEUR');Ferme(Qr); exit ; end;
        Ferme(Qr);
     end else
     begin     //mcd 05/03/02
        if (old_payeur <>'') then SetField('T_PAYEUR','');
     end;
   end;
            //**************************** Gestion des chronos automatiques *********************
   SetField('T_TIERS',Trim(GetField('T_TIERS')));
   if GetField ('T_NATUREAUXI') = 'FOU' then
      begin
      if GetParamsoc('SO_GCNUMFOURAUTO') = TRUE then // changer en SO_GCNUMTIERSAUTO si pas parametrage propre au fournisseur
         begin
         if (ds<>nil) and (ds.state = dsinsert) and (GetField ('T_TIERS')= '') then
            begin
            if ((ctxMode in V_PGI.PGIContexte) or (ctxChr in V_PGI.PGIContexte)) then
               CodeTiers:=AttribNewCode('TIERS','T_TIERS',GetParamsoc('SO_GCLGNUMTIERS'),trim(GetParamsoc('SO_GCPREFIXETIERS')),GetParamsoc('SO_GCCOMPTEURTIERS'),'')
               else
               CodeTiers:=AttribNewCode('TIERS','T_TIERS',-1,'',GetParamsoc('SO_GCCOMPTEURTIERS'),'');
            SetField ('T_TIERS', CodeTiers);
            NumChrono:=ExtraitChronoCode(CodeTiers);
            SetParamSoc('SO_GCCOMPTEURTIERS', NumChrono) ;
            end;
         end;
      end;
   If (GetField ('T_TIERS')= '') then begin SetLastError(1, 'T_TIERS'); exit ; end;
   if (ds<>nil) and (DS.State=dsInsert) and Presence('TIERS','T_TIERS',GetField ('T_TIERS')) then begin SetLastError(17, 'T_TIERS'); exit ; end ;

            // Gestion du compte auxiliaire
   SetField('T_AUXILIAIRE',Uppercase(Trim(GetField('T_AUXILIAIRE'))));
   if (ds<>nil) and (DS.State=dsInsert) then
      begin
      if (GetField ('T_AUXILIAIRE')= '') then AffecteTAuxiliaire(Trim(GetParamSoc('SO_GCPREFIXEAUXIFOU'))+GetField('T_TIERS'), True)
                                         else AffecteTAuxiliaire(GetField('T_AUXILIAIRE'),False);
      if Presence('TIERS','T_AUXILIAIRE',GetField ('T_AUXILIAIRE')) then
         begin
         SetLastError(21, 'T_AUXILIAIRE');
         exit ;
         end ;
      end;

   if (GetField('T_CREDITPLAFOND') < GetField ('T_CREDITACCORDE')) then  begin SetLastError(13, 'T_CREDITPLAFOND'); exit ; end;

   if (GetField('T_REGIMETVA')='') then SetField('T_REGIMETVA',VH^.RegimeDefaut);
   if (GetField('T_COLLECTIF')='') then SetField('T_COLLECTIF',VH^.DefautFou);
   if (GetField('T_MODEREGLE')='') then SetField('T_MODEREGLE',GetParamSoc('SO_GcModeRegleDefaut'));

   end;               // Fin de gestion des "non-clients" : fournisseurs ?
// Fin gestion Tiers Fournisseurs ====================================================================

   // on regarde si collectif existe dans fiche
  if Not Presence('GENERAUX','G_GENERAL',GetField('T_COLLECTIF')) then
  begin
  SetLastError(24, 'T_COLLECTIF');
  exit;
  end;
  // debut Affaire
  if (ctxSCOT in V_PGI.PGIContexte) and (GetField ('T_NATUREAUXI') = 'CLI')  then
   begin              // ctrl mois clôture
   if ((GetField ('T_MOISCLOTURE')< 1) or (GetField ('T_MOISCLOTURE')> 12))
      then begin SetLastError(22, 'T_MOSICLOTURE'); exit ; end;
   end;

  SetField ('T_DATEMODIF', V_PGI.DateEntree);
  if ModifLot then TFFiche(ecran).BFermeClick(nil);
  if (ds<>nil) and (DS.State=dsInsert) then
   begin
   SetBoutonEnabled(TRUE);
   SetControlEnabled('T_TIERS',FALSE);
   SetControlEnabled('T_AUXILIAIRE',FALSE);
   end;

  // Modif due à l'ajout d'une seconde table pour la gestion des zones libres
  If (GetField('T_AUXILIAIRE')='') then begin SetLastError(23, 'T_AUXILIAIRE'); exit ; end;
  if (ds<>nil) then
  begin
  TobZoneLibre.GetEcran (TFfiche(Ecran),Nil);
  TobZoneLibre.PutValue ('YTC_AUXILIAIRE', GetField('T_AUXILIAIRE')) ;
  TobZoneLibre.PutValue ('YTC_TIERS', GetField('T_TIERS')) ;
  TobZoneLibre.InsertOrUpdateDB (FALSE);
  TobZoneLibre.SetAllModifie(False);
  end;

  // grc: si 1 info compl obligatoire, chainage écran saisie infos sinon enregistrement table prospects initialisé
  if (ctxGRC in V_PGI.PGIContexte) and (DS<>nil) and (DS.State=dsInsert) and
     ( (GetField ('T_NATUREAUXI') = 'CLI') or (GetField ('T_NATUREAUXI') = 'PRO') ) then
  begin
    ExisteOblig:=ExisteSQL('SELECT RCL_OBLIGATOIRE FROM CHAMPSPRO WHERE RCL_OBLIGATOIRE="X" ');
    if ( not ExisteOblig ) or ( CritereSuspect ) then
    begin
      TOBProspect:= TOB.Create('PROSPECTS', nil, -1);
      TOBProspect.InitValeurs;
      TOBProspect.PutValue ('RPR_AUXILIAIRE',GetField('T_AUXILIAIRE'));
{$IFDEF GRC}
      if CritereSuspect then
        RTCorrespondSuspectProspect('RSC',TOBSuspectCompl, TOBProspect);
{$ENDIF}
      TOBProspect.InsertOrUpdateDB (FALSE);
      TOBProspect.free;
      TOBProspect := NIL;
    end;
    if ExisteOblig then
       AGLLanceFiche('RT','RTPARAMCL','','','FICHEPARAM='+TFfiche(Ecran).name+';FICHEINFOS='+GetField('T_AUXILIAIRE')+';NATUREAUXI='+GetField('T_NATUREAUXI'));

    if CritereSuspect then   // la fiche suspect a une date de basculement suspect -> prospect
    begin
      TobSuspect.putvalue('RSU_FERME','X');
      TobSuspect.putvalue('RSU_DATEFERMETURE',V_PGI.DateEntree);
      TobSuspect.putvalue('RSU_DATESUSPRO',V_PGI.DateEntree);
      TobSuspect.UpdateDB (FALSE);
      if ( trim (TobSuspect.Getvalue('RSU_CONTACTNOM')) <> '' ) then
      begin
        TOB_CON:=TOB.Create('CONTACT', NIL, -1);  //Création du contact principal à partir de la fiche suspect
        TOB_CON.PutValue('C_TYPECONTACT','T') ;
        TOB_CON.PutValue('C_AUXILIAIRE',GetField('T_AUXILIAIRE')) ;
        if Client_particulier then
        begin
          TOB_CON.PutValue('C_NUMEROCONTACT','2');
          TOB_CON.PutValue('C_PRINCIPAL','-')
        end else
        begin
          TOB_CON.PutValue('C_NUMEROCONTACT','1') ;
          TOB_CON.PutValue('C_PRINCIPAL','X') ;
        end;
        TOB_CON.PutValue('C_NATUREAUXI','PRO') ;
        TOB_CON.PutValue('C_CIVILITE',TobSuspect.Getvalue('RSU_CONTACTCIVILITE')) ;
        TOB_CON.PutValue('C_NOM',TobSuspect.Getvalue('RSU_CONTACTNOM')) ;
        TOB_CON.PutValue('C_PRENOM',TobSuspect.Getvalue('RSU_CONTACTPRENOM')) ;
        TOB_CON.PutValue('C_FONCTIONCODEE',TobSuspect.Getvalue('RSU_CONTACTFONCTION')) ;
        //TOB_CON.PutValue('C_PUBLIPOSTAGE','X') ;
        TOB_CON.PutValue('C_PUBLIPOSTAGE',TobSuspect.Getvalue('RSU_CONTACTPUBLI')) ;
        TOB_CON.PutValue('C_EMAILING',TobSuspect.Getvalue('RSU_CONTACTEMLG')) ;
        TOB_CON.PutValue('C_TELEPHONE',TobSuspect.Getvalue('RSU_CONTACTTELEPH')) ;
        TOB_CON.PutValue('C_RVA',TobSuspect.Getvalue('RSU_CONTACTRVA')) ;
        TOB_CON.InsertOrUpdateDB (FALSE);
        TOB_CON.free;
      end;
      TOBSuspectCompl.free;
      TOBSuspectCompl := NIL;
    end;
  end;

  if (ds<>nil) and (Ecran is TFFiche) then
  begin
    case Tools.CaseFromString(TFFiche(Ecran).Name, [frm_ThirdCliPro, 'GCTIERSFO', 'HRCLIENTS']) of
      {} 0..2 : TFFiche(Ecran).Retour := GetField('T_TIERS');
    else
      TFFiche(Ecran).Retour :=  '';
    end;
  end;


  CodeTiersDP:=GetField('T_TIERS');

  OnFerme:=False;
  If (ds<>nil) then
    begin
    if (DS.State in [dsInsert])
       then DerniereCreate := GetField('T_AUXILIAIRE')
       else if (DerniereCreate = GetField('T_AUXILIAIRE')) then OnFerme:=True; // le bug arrive on se casse !!!
    end;
  {$IFDEF GRC}
  if TransProCli then
      begin
      RTMajNatureContacts(GetField('T_AUXILIAIRE'),GetField('T_NATUREAUXI'));
      TransProCli:=false;
      end;
  {$ENDIF}
end;

Function  TOM_TIERS.ChampsObligatoires : boolean;
var NomChamp,TypeTiers: string;
begin
   result:=false;
   NomChamp:='';
   if VH_GC.GCIfDefCEGID then
   begin
     if (GetField ('T_REPRESENTANT')= '')  then NomChamp:='T_REPRESENTANT'
     else if (GetField ('T_COMPTATIERS')= '')  then  NomChamp:='T_COMPTATIERS'
     else if (GetField ('T_REGIMETVA')= '')  then  NomChamp:='T_REGIMETVA'
     else if ( (GetField ('T_APE')= '') and (client_particulier = false ) )  then  NomChamp:='T_APE'
     else if (GetControlText ('YTC_TABLELIBRETIERS1')= '')  then  NomChamp:='YTC_TABLELIBRETIERS1'
     else if (GetControlText ('YTC_TABLELIBRETIERS2')= '')  then  NomChamp:='YTC_TABLELIBRETIERS2'
     else if (GetControlText ('YTC_TABLELIBRETIERS3')= '')  then  NomChamp:='YTC_TABLELIBRETIERS3'
     else if (GetControlText ('YTC_TABLELIBRETIERS4')= '')  then  NomChamp:='YTC_TABLELIBRETIERS4'
     ;
     if NomChamp<>'' then
        begin
        SetFocusControl(NomChamp) ;
        LastError:=33; LastErrorMsg:=TexteMessage[LastError]+champToLibelle(NomChamp);
        result:=True;
        end;
   end
else
begin
  {$IFNDEF CCS3}
  TypeTiers:=GetField('T_NATUREAUXI') ;
  NomChamp:=VerifierChampsObligatoires(Ecran,TypeTiers);
  if NomChamp='' then
    if (TypeTiers='CLI') or (TypeTiers='PRO')
       then NomChamp:=NomChamp +VerifierChampsObligatoires(Ecran,TypeTiers,'TIERSSUITE',False) // mcd pour tiers compl
       else if TypeTiers='FOU' then NomChamp:=NomChamp +VerifierChampsObligatoires(Ecran,TypeTiers,'FOURSUITE',False); // mcd pour tiers compl
  if NomChamp<>'' then
        begin
        NomChamp:=ReadTokenSt(NomChamp);
        SetFocusControl(NomChamp) ;
        LastError:=36; LastErrorMsg:=TexteMessage[LastError]+champToLibelle(NomChamp);
        result:=True;
        end;

  {$ENDIF}
end;
end;
//////////////////////////////////////////////////////////////////////

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Pascal Gautherot
Créé le ...... : 16/09/2003
Modifié le ... :   /  /
Description .. : Fonction de réaffichage des données du contacte principal
Suite ........ : dans la fiche tiers. Appellée d'une part par le OnLoad, et
Suite ........ : d'autre part par la fonction registrée AGLAfficheContactTiers
Mots clefs ... : CONTACT TIERS
*****************************************************************}
Procedure TOM_TIERS.AfficheContactTiers;
var
    sReqCont : String;
    qContact : TQuery;
begin
  sReqCont := 'SELECT C_NOM, C_PRENOM, C_TELEPHONE, C_RVA FROM CONTACT WHERE C_TYPECONTACT = "T" AND C_AUXILIAIRE = "'+STring(GetField('T_AUXILIAIRE'))+'" AND C_PRINCIPAL = "X"';
  qContact := OpenSql(sReqCont,False);
  if not TobContact.SelectDB('',qContact,True) then
    TobContact.InitValeurs;
  TobContact.PutEcran(TFfiche(Ecran));
  Ferme(qContact);
end;

procedure TOM_TIERS.OnLoadRecord;
Var QLoc      : TQuery;
    staction  : TActionFiche ;
    Codeper   : integer;
    QContrat  : TQuery;
    StSQL     : String;
    //TobLoc  : TOB;
    sIncoterm : string;
    QQ : Tquery;
begin
inherited;
  //
  QQ := OpenSql ('SELECT * FROM BTIERS WHERE BT1_AUXILIAIRE="'+GetField('T_AUXILIAIRE')+'"',true,1,'',true);
  if not QQ.eof then
  begin
    TOBTIERSBTP.SelectDb('',QQ);
  end else
  begin
    TOBTIERSBTP.InitValeurs(false);
    TOBTIERSBTP.SetString('BT1_AUXILIAIRE',GetField('T_AUXILIAIRE'));
  end;
  ferme (QQ);
  //
  //thValComboBox(GetControl('T_REGION')).PLUS := 'RG_PAYS="'+GetControlText('T_PAYS')+'"';
  SetControlProperty('T_REGION','PLUS','RG_PAYS="'+GetControlText('T_PAYS')+'"');
  TransProCli:=False;
  if (stNatureAuxi = 'PRO') and (DS.State <> dsInsert)
   and ( ExJaiLeDroitConcept(TConcept(gcCLICreat),False) ) then SetControlVisible ('BCLIENT',True);
  SetControlEnabled('T_AUXILIAIRE',(DS.State in [dsInsert]));
  //Non en line
  If (ctxCHR in V_PGI.PGIContexte) or ((ctxGRC in V_PGI.PGIContexte) and
   ((stNatureAuxi = 'CLI')or (stNatureAuxi = 'PRO') or (stNatureAuxi = 'FOU')) and
   (Not(DS.State in [dsInsert]))) Then
   SetControlVisible ('BCOURRIER',True)
  else
   SetControlVisible ('BCOURRIER',false);


  //if ModifLot then SetArguments(StSQL);

  LgCode:=VH^.Cpta[fbAux].Lg ;
  AffIndisponible (True);                // Gestion des affichages relatifs à T_FERME
  SetControlText('XXCODEPER','0');
  SetControlEnabled('BDUPLICATION',TiersCreation) ;
 //AFFAIRE debut
        // Appel fct pour DP -> tiers
  if (stNatureAuxi = 'CLI')or (stNatureAuxi = 'FOU') then
  begin
    if (ctxscot in V_PGI.PGIContexte) and (GetParamsoc('SO_AFLIENDP') = true) then
    begin
      if (Getfield('T_TIERS') <>'') then
      begin
        codeper:=StrToInt(GetcontrolText('XXCODEPER'));
        {$IFDEF DP}
        DpSynchro(True, codeper, Getfield('T_TIERS'), stNatureAuxi,Self);  // MCD 07/08/00 modif
        {$endif}
        SetControlText('XXCODEPER',IntTostr(codeper));
        staction :=   TActionFiche(GetControl ('typeAction' ));
        end
      else
      begin
        codeper := 0 ;
        staction := TaModif ; // ou TaConsult ???
      end ;
      if (codeper = 0)  and (staction<>TaConsult) and (origine <> 'DP') then
      begin
              DS.Edit;
              //RechDp( CodePer, stNatureAuxi,self);   //MCD 07/08/00
              RechDpBis( CodePer,GetField('T_TIERS'), stNatureAuxi);   //MCD 07/08/00
              SetControlText('XXCODEPER',IntTostr(codeper));
              end
      end;
  end;
 //AFFAIRE fin

  // zone n'existe plus if GetField('T_CONFIDENTIEL')='1' then SetControlText('CB_CONFIDENTIEL','X') else SetControlText('CB_CONFIDENTIEL','-');

  if GetField ('T_NATUREAUXI') = 'CLI' then
    begin
    if GetField('T_MULTIDEVISE') = 'X' then
        begin
      QLoc := OpenSql ('SELECT GP_DATEPIECE FROM PIECE WHERE GP_TIERS="' +  GetField('T_TIERS') + '"', True);
        if not QLoc.Eof then
            begin
            SetControlEnabled ('T_MULTIDEVISE', False);
      end
      else
            begin
            Ferme (Qloc);
        QLoc := OpenSql('SELECT E_DATECOMPTABLE FROM ECRITURE WHERE E_AUXILIAIRE="' + GetField('T_AUXILIAIRE') + '"', True);
            if not QLoc.Eof then
        Begin
                SetControlEnabled ('T_MULTIDEVISE', False);
                end;
            end;
        Ferme (QLoc);
        end;
    //SetControlProperty ('T_FACTURE', 'DataType', 'GCTIERSCLI');
    If Pos('Client :',Ecran.Caption)=0 then Ecran.Caption:='Client :';
    end;

  (* Gestion du tiers de facturation et payeur pour les fournisseurs
  if GetField ('T_NATUREAUXI') <> 'FOU' then
  begin
  *)

  if (GetField ('T_PAYEUR')<>'') then
      begin
    QLoc := OpenSql('SELECT T_TIERS FROM TIERS WHERE T_AUXILIAIRE="' + GetField('T_PAYEUR') + '"', False);
      if not QLoc.Eof then
          begin
          Old_Payeur := QLoc.Findfield('T_TIERS').AsString;
          SetCOntrolText('TE_PAYEUR', Old_Payeur);
          end;
      Ferme(QLoc);
  end
  else
  begin     //mcd 05/03/01
          old_payeur:='';
          SetCOntrolText('TE_PAYEUR','');
      end;

  if (GetField ('T_FACTURE')<>'') then
      begin
    if GetField('T_FACTURE') = GetField('T_AUXILIAIRE') then
    begin
            Old_Facture :=GetField('T_TIERS');   // a mettre obligatoiremeent avant setcontroltext, car setcontrol passe par onchenge de la zone
            SetCOntrolText('TE_FACTURE',GetField('T_TIERS')) ;
            end
    else
    begin
      QLoc := OpenSql('SELECT T_TIERS FROM TIERS WHERE T_AUXILIAIRE="' + GetField('T_FACTURE') + '"', False);
          if not QLoc.Eof then
              begin
              Old_Facture:= QLoc.Findfield('T_TIERS').AsString;
              SetCOntrolText('TE_FACTURE', Old_Facture);
        //Old_Facture:= QLoc.Fields[0].AsString;
        //SetCOntrolText('TE_FACTURE', QLoc.Fields[0].AsString);
              end;
          Ferme(QLoc);
          end;
      end;

    { BUG : 10368 }
  if (GetField ('T_NATUREAUXI') = 'FOU') then
    Thedit(GetControl('T_FACTURE')).UpdateLibelle { GPAO }
  else
    Thedit(GetControl('TE_FACTURE')).UpdateLibelle; { GPAO }
  (* end; *)

  // Activation ou désactivation du bouton de paramétrage du code à barres Article dans la fiche fournisseur
  if GetField ('T_NATUREAUXI') = 'FOU' then
   begin
   if ctxMode in V_PGI.PGIContexte then
    begin
      if GetParamsoc('SO_GCCABFOURNIS')=True then
        SetControlEnabled('BCODEBARRE', True)
      else
        SetControlEnabled('BCODEBARRE', False);
   end;
  End;

  // prospect : nature prospect
  //if GetField ('T_NATUREAUXI') = 'PRO' then
  if (ctxGRC in V_PGI.PGIContexte) and ( ( GetField ('T_NATUREAUXI') = 'PRO') or ( GetField ('T_NATUREAUXI') = 'CLI') ) then
  begin
  SetControlProperty ('TE_FACTURE', 'DataType', 'RTTIERSPRO');
  SetControlProperty ('TE_LIVRE', 'DataType', 'RTTIERSPRO');
  //If Pos('Prospect :',Ecran.Caption)=0 then Ecran.Caption:='Prospect :' + Ecran.Caption ;
  end;

  if DuplicFromMenu then
  begin
  Duplication(TiersToDuplic);
  TiersToDuplic := '';
  end;

  if ds.state in [dsinsert] then
   begin
   SetControlEnabled('T_TIERS', true);
{$IFDEF EAGLCLIENT} // mng 30-05-02
   TFFiche(Ecran).QFiche.CurrentFille.PutEcran(TFFiche(Ecran)) ;
{$ENDIF}
   SetFocusControl('T_TIERS') ;
   end
  else
    SetControlEnabled('T_TIERS', FALSE);

  If GetField ('T_NATUREAUXI') = 'FOU' then
  begin
          if (GetParamsoc('SO_GCNUMFOURAUTO')  = TRUE) then
             begin
             SetControlEnabled('T_TIERS', FALSE); //mcd 15/03/00
             SetFocusControl('T_JURIDIQUE') ; // mng 30-05-02
             end;
          end
  else
  begin
          if (GetParamsoc('SO_GCNUMTIERSAUTO')  = TRUE) then
             begin
             SetControlEnabled('T_TIERS', FALSE); //mcd 05/03/00
             SetFocusControl('T_JURIDIQUE') ; // mng 30-05-02
             end;
          end;

	Client_particulier := FALSE;

  //if (GetField ('T_NATUREAUXI') = 'CLI') or (GetField ('T_NATUREAUXI') = 'PRO') or (GetField ('T_NATUREAUXI') = 'CON') Or then  //GRC

	if (GetControl('T_EMAIL') <> nil) then
     begin
   	 client_particulier := (GetField('T_PARTICULIER')='X');
	 	 if (ctxChr in V_PGI.PGIContexte) then  //CHR
     	  begin
      	RecupContact;
      	end;
     if (GetControl('BNEWMAIL') <> nil) then SetControlEnabled('BNEWMAIL', GetField('T_EMAIL')<>'');
  	 end;

  AfficheEcranClientMode(Client_particulier);

// Modif due à l'ajout d'une seconde table pour la gestion des zones libres
{TobZonelibre:=TOB.Create ('Zones libres', Nil, -1);
if TobZonelibre<>nil then
    begin
    Qloc:=OpenSQL('select * from TIERSCOMPL where YTC_AUXILIAIRE="'+GetField('T_AUXILIAIRE')+'"',True) ;
    if QLoc.Eof then
       begin
       TOBLoc := TOB.Create ('TIERSCOMPL', TobZonelibre, -1);
       TOBLoc.PutValue ('YTC_AUXILIAIRE', GetField('T_AUXILIAIRE')) ;
       TOBLoc.PutValue ('YTC_TIERS', GetField('T_TIERS')) ;
       end else
       begin
       TobZonelibre.LoadDetailDB('TIERSCOMPL','','',Qloc, False) ;
       end;
    Ferme(QLoc) ;
    if TFfiche(Ecran)<>nil then TobZonelibre.detail[0].PutEcran(TFfiche(Ecran));
    end;
}

if  (GetField ('T_SOCIETEGROUPE')<>'') and (GetField ('T_SOCIETEGROUPE') <> GetField ('T_TIERS')) then
    begin
    AddRemoveItemFromPopup ('POPM','MNSOCMERE',True);
    AddRemoveItemFromPopup ('POPM','MNGROUPE',True);
    end else
    begin
    AddRemoveItemFromPopup ('POPM','MNSOCMERE',False);
    AddRemoveItemFromPopup ('POPM','MNGROUPE', False);
    end;
if (GetField ('T_PRESCRIPTEUR')<>'') and (GetField ('T_PRESCRIPTEUR') <> GetField ('T_TIERS'))
   then AddRemoveItemFromPopup ('POPM','MNPRESCRIPTEUR',True)
   else AddRemoveItemFromPopup ('POPM','MNPRESCRIPTEUR',False);

if (critereSuspect) and (ds.state = dsinsert) then
        SuspectToEcran   // GRC- Met à l'écran infos provenant fiche suspect
else if (not TobZonelibre.SelectDB('"'+GetField('T_AUXILIAIRE')+'"',Nil )) then
        TobZonelibre.InitValeurs;

if (StNatureAuxi = 'CLI') or (StNatureAuxi = 'PRO') then 
begin
  SetControlText('TE_LIVRE', '');
  if (TobZoneLibre.GetValue('YTC_TIERSLIVRE') <> '') then
  begin
    if TobZoneLibre.GetValue('YTC_TIERSLIVRE') = GetField('T_AUXILIAIRE') then
    begin
      Old_Livre := GetField('T_TIERS');
      SetControlText('TE_LIVRE', GetField('T_TIERS')) ;
    end
    else
    begin
      QLoc := OpenSql('SELECT T_TIERS FROM TIERS WHERE T_AUXILIAIRE="' + TobZoneLibre.GetValue('YTC_TIERSLIVRE') + '"', False);
      if not QLoc.Eof then
      begin
        Old_Livre := QLoc.Findfield('T_TIERS').AsString;
        SetControlText('TE_LIVRE', Old_Livre);
      end;
      Ferme(QLoc);
    end;
  end
  else
  begin
    SetControlProperty('YTC_NADRESSELIV', 'ENABLED', False);  
  end;
  Thedit(GetControl('TE_LIVRE')).UpdateLibelle;
end;  

TobZonelibre.PutEcran(TFfiche(Ecran));
sIncoterm := RechDom('GCINCOTERM', TobZoneLibre.getValue('YTC_INCOTERM') , False);
If sIncoterm = 'Error' then sIncoterm := '';
SetControlCaption('TYTC_INCOTERM', sIncoterm);


If Not(DS.State in [dsInsert]) Then  DerniereCreate := '';

{$IFNDEF CCS3}
AppliquerConfidentialite(Ecran,stNatureAuxi);
{$ENDIF}

if ModifLot then SetArguments(StSQL);

{$IFDEF NOMADE}
SetControlEnabled('T_REPRESENTANT', False);
SetControlVisible ('GBREPRESENTANT', False);
{$ENDIF}

if (GetField ('T_NATUREAUXI') = 'CLI') then
begin
  if Assigned(GetControl('MnTiersEDI')) then
  {$IFDEF EDI}
    TMenuItem(GetControl('MnTiersEDI')).Visible := True;
  {$ELSE}
    TMenuItem(GetControl('MnTiersEDI')).Visible := False;
  {$ENDIF}
  SetPlusNumAdresseLiv(Self);
  SetPlusNumAdresseFac(Self);
end;

AfficheContactTiers;

  // Activation ou désactivation du bouton d'affichage des contrats
  if (GetField ('T_NATUREAUXI') = 'CLI') then
   Begin
      StSQL := 'SELECT AFF_AFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE0="I" AND AFF_TIERS="' + GetField ('T_TIERS') + '"';
      QContrat := OpenSQL(StSQL, true);
      If QContrat.Eof then
      BtContrat.Visible := False
      else
      BtContrat.Visible := True;
      Ferme(QContrat);
  end
  else
    BTCONTRAT.Visible := False;


// Activation ou désactivation du bouton d'affichage des Appels
  if (GetField ('T_NATUREAUXI') = 'CLI') then
   Begin
    //Non en line
    //FV1 : 22/12/2015 - FS#1820 - VEODIS : Pb en saisie d'appel depuis fiche client
    if not(CreationAffaireAutorise) then BTAPPELS.Visible := false else BTAPPELS.Visible := True;
  End
  else
    BTAPPELS.Visible := false;

{$IFDEF CCS3}
  if (getcontrol('GB_RESSOURCE') <> Nil) then SetControlVisible ('GB_RESSOURCE', False);
{$ENDIF}
{$IFDEF EAGLCLIENT} // JT eQualité 11030
if GetControl('T_CONFIDENTIEL') is TCheckBox then
  TCheckBox(GetControl('T_CONFIDENTIEL')).Visible := False;
{$ENDIF}
{$IFDEF BTP}
    //
    TOBEtude := TOB.create ('Les Etudes du tiers', nil, -1) ; //mcd 23/01/03 chgmt nom tob
    // SELECT * : nombre de champs et d'enreg restreint
    QLoc := opensql ('SELECT * FROM BDETETUDE WHERE BDE_AFFAIRE="" AND BDE_TIERS="' + GetField ('T_TIERS') + '" AND BDE_NATUREAUXI="'+GetField ('T_NATUREAUXI') + '"', false) ;
    TOBEtude.LoadDetailDB ('BDETETUDE', '', '', QLoc, false, true) ;
    ferme (QLoc) ;
{$ENDIF}
end;

/////////////////////////////////////////////////////////////////////////////

procedure TOM_TIERS.SetBoutonEnabled(Etat : boolean);
begin
SetControlEnabled('BPOPMENU',Etat);
SetControlEnabled('BPOP2MENU',Etat);
SetControlEnabled('BADRESSE',Etat);
SetControlEnabled('BCONTACT',Etat);
SetControlEnabled('BRIB',Etat);
SetControlEnabled('BMEMO',Etat);
end;

/////////////////////////////////////////////////////////////////////////////

procedure TOM_TIERS.OnNewRecord;
Var QQ : TQuery;
    Nbr : integer ;
{$IFDEF NOMADE}
    stRepresentant : string ;
{$ENDIF}
{$IFDEF CTI}
    StNotel : string;
{$ENDIF}
begin
Inherited ;
Old_Facture :='' ;
Old_Payeur :=''  ;
Old_Livre   :='';
if StNatureAuxi = 'CLI' then 
begin
  SetControlText('TE_FACTURE','') ;
  SetControlText('TE_PAYEUR','') ;
  SetControlText('TE_LIVRE', '');
  SetPlusNumAdresseLiv(Self);
  SetPlusNumAdresseFac(Self);
end;

if V_PGI.VersionDemo then
  begin
  QQ:=OpenSQL('SELECT COUNT(*) FROM TIERS',True) ;
  if Not QQ.EOF then Nbr:=QQ.Fields[0].AsInteger else Nbr:=0;
  Ferme(QQ) ;
  if (Nbr >= 100) then PgiBox (TexteMessage[32], 'Fiche Tiers');
  end;
SetField ('T_NATUREAUXI', stNatureAuxi);
SetField ('T_DATECREATION', V_PGI.DateEntree); SetField ('T_DATEOUVERTURE', V_PGI.DateEntree);
SetField ('T_CREERPAR', 'SAI'); SetField ('T_DEVISE', V_PGI.devisepivot);
SetField ('T_LETTRABLE', 'X');
if stNatureAuxi='FOU' then SetField('T_COLLECTIF',VH^.DefautFou) ;
if (stNatureAuxi='CLI') then    SetField('T_COLLECTIF',VH^.DefautCli) ;
if (stNatureAuxi='FOU') or (stNatureAuxi='CLI') or(stNatureAuxi='PRO') then
begin
	SetField('T_EURODEFAUT','X')
end;
if stNatureAuxi='PRO' then SetField('T_COLLECTIF',VH^.DefautCli) ; // prospect
if stNatureAuxi='CON' then SetField('T_COLLECTIF',VH^.DefautCli) ; // GRC
SetField('T_REGIMETVA',VH^.RegimeDefaut);
SetField ('T_PAYS', GetParamSoc('SO_GcTiersPays'));
SetField ('T_MODEREGLE', GetParamSoc('SO_GcModeRegleDefaut'));
SetField ('T_TVAENCAISSEMENT', GetParamSoc('SO_TVAENCAISSEMENT'));

SetControlEnabled ('T_TIERS', TRUE);
TiersCreation := True;
if stNatureAuxi='' then PGIBox('La nature d''auxiliaire n''est pas définie' ,TFFiche(ecran).Caption)
                   else if stNatureAuxi='FOU' then SetField('T_FACTUREHT','X')       //JCF
                   else if VH_GC.GCDefFactureHT then SetField('T_FACTUREHT','X') ;   //PCS

SetBoutonEnabled(FALSE);
LgCode:=VH^.Cpta[fbAux].Lg ;
SetField('T_CONFIDENTIEL','0') ;
SetControlChecked('T_PUBLIPOSTAGE',True);
SetField('T_PUBLIPOSTAGE','X');

Client_particulier := FALSE;
If not (ctxAffaire in V_PGI.PGIContexte) and (GetField ('T_NATUREAUXI') = 'CLI') then
    begin
    if (GetParamsoc('SO_GCPARTICULIER') = TRUE) then
       begin
       SetField('T_PARTICULIER','X');
       client_particulier := True;
       end else
       begin
       SetField('T_PARTICULIER','-');
       client_particulier := False;
       end;
    SetField('T_MODEREGLE',VH_GC.GCModeRegleDefaut);
    end;

if GetField ('T_COLLECTIF') = '' then
    begin
    PgiBox (TexteMessage[15], 'Fiche Tiers');
    end;

{$IFDEF NOMADE}
stRepresentant := GetColonneSQL('COMMERCIAL','GCL_COMMERCIAL','GCL_UTILASSOCIE="'+V_PGI.User+'"') ;
SetField ('T_REPRESENTANT', stRepresentant);
{$ENDIF}
{$IFDEF CTI}
if AppelantCTI<>'' then
   begin
   SetField ('T_TELEPHONE',StNotel);
   end;
{$ENDIF}
end;

procedure TOM_TIERS.OnCancelRecord;
var sIncoterm : string; 
begin
Inherited ;
  if TFfiche(Ecran)<>nil then
  begin
    TobZonelibre.PutEcran(TFfiche(Ecran));
    sIncoterm := RechDom('GCINCOTERM', TobZoneLibre.getValue('YTC_INCOTERM') , False);
    if sIncoterm = 'Error' then sIncoterm := '';
    SetControlCaption('TYTC_INCOTERM', sIncoterm);
  end;
end;


procedure TOM_Tiers.AfficheEcranClientMode(cli_part : boolean);
var libel : string;
    pop : TPopupMenu;
    i : integer;
begin
client_particulier := cli_part;
if client_particulier = TRUE then
    begin
    libel := TraduireMemoire ('Client : ') + GetField('T_JURIDIQUE') + GetField('T_LIBELLE');
    Ecran.Caption:=libel;
    If ctxSCot in V_PGI.PGIContexte  then begin
      SetControlVisible ('T_JOURNAISSANCE', False);
      SetControlVisible ('T_MOISNAISSANCE', False);
      SetControlVisible ('T_ANNEENAISSANCE', False);
      SetControlVisible ('T_DATNAISSANCE', False);
      SetControlVisible ('SEP_MOISNAISSANCE', False);
      SetControlVisible ('SEP_ANNEENAISSANCE', False);
      SetControlVisible ('TT_PRENOM',True);  //mcd 03/10/03 mise en true
      SetControlVisible ('T_SIRET', true);
      SetControlVisible ('TT_SIRET', True);
       end
    else begin
      SetControlVisible ('T_JOURNAISSANCE', TRUE);
      SetControlVisible ('T_MOISNAISSANCE', TRUE);
      SetControlVisible ('T_ANNEENAISSANCE', TRUE);
      SetControlVisible ('T_DATNAISSANCE', TRUE);
      SetControlVisible ('SEP_MOISNAISSANCE', TRUE);
      SetControlVisible ('SEP_ANNEENAISSANCE', TRUE);
      SetControlVisible ('TT_PRENOM', TRUE);
      SetControlVisible ('T_SIRET', FALSE);
      SetControlVisible ('TT_SIRET', FALSE);
      end;
    if ctxMode in V_PGI.PGIContexte then
    begin
      SetControlVisible ('T_SOCIETEGROUPE', False);
      SetControlVisible ('TT_SOCIETEGROUPE', False);
    end;
    SetControlVisible ('T_CONFIDENTIEL', False);
    SetControlVisible ('TT_CONFIDENTIEL', False);
    SetControlVisible ('T_PASSWINTERNET', False);
    SetControlVisible ('TT_PASSWINTERNET', False);
    if ctxMode in V_PGI.PGIContexte then
    begin
      SetControlVisible ('T_RVA', False);
      SetControlVisible ('TT_RVA', False);
    end;
    SetControlVisible ('T_ENSEIGNE', False);
    SetControlVisible ('TT_ENSEIGNE', False);
    SetControlVisible ('T_FORMEJURIDIQUE', False);
    SetControlVisible ('TT_FORMEJURIDIQUE', False);
    SetControlVisible ('GBIDENTIFIANT', FALSE);
    SetControlVisible ('T_NIF', FALSE);
    SetControlVisible ('TT_NIF', FALSE);
    SetControlVisible ('T_EAN', FALSE);
    SetControlVisible ('TT_EAN', FALSE);
    SetControlVisible ('T_APE', FALSE);
    SetControlVisible ('TT_APE', FALSE);
    SetControlVisible ('TT_APE_', FALSE);
    SetControlProperty ('T_JURIDIQUE', 'DataType', 'TTCIVILITE');
//    TTAbSheet(GetControl('PREGLEMENT')).TabVisible := FALSE;
//    SetControlVisible ('GB_TARIFICATION', TRUE);
    SetControlText ('TT_LIBELLE', TraduireMemoire('&Nom'));
    SetControlText ('TT_TELEPHONE', TraduireMemoire('&Tél domicile'));
    SetControlText ('TT_FAX', TraduireMemoire('Tél bureau'));
    SetControlText ('TT_TELEX', TraduireMemoire('Tél portable'));
    {If not(ctxChr in V_PGI.PGIContexte) then       // CHR
       SetControlText ('TT_RVA', TraduireMemoire('&E-mail'));}  //suite réunion, T_RVA est systématiquement utilisé pour le site Web, y compris pour un particulier.
    SetControlText ('TT_JURIDIQUE', TraduireMemoire('&Civilité'));
    If not(ctxSCot in V_PGI.PGIContexte) then begin
       SetControlVisible ('T_SEXE', TRUE);
       SetControlVisible ('TT_SEXE', TRUE);
       SetControlVisible ('GBPARTICULIER', TRUE); 
       end;
{    if GetField ('T_NATUREAUXI') <> 'CON' then    // mng
        SetControlVisible ('BNEWMAIL', TRUE); }
    if (ctxCHR in V_PGI.PGIContexte) then    //CHR
    begin
      SetControlText ('TT_LIBELLE', TraduireMemoire('Nom'));
      SetControlText ('TT_TELEPHONE', TraduireMemoire('Tél domicile'));
      SetControlText ('TT_FAX', TraduireMemoire('Tél bureau'));
      SetControlText ('TT_TELEX', TraduireMemoire('Tél portable'));
      SetControlText ('TT_JURIDIQUE', TraduireMemoire('Civilité'));
      SetControlVisible ('IDENTITE' , True);
      SetControlVisible ('T_PRENOM', TRUE);
      SetControlVisible ('T_SOCIETEGROUPE', True);
      SetControlVisible ('TT_SOCIETEGROUPE', True);
      SetControlVisible ('T_CONFIDENTIEL', True);
      SetControlVisible ('TT_CONFIDENTIEL', True);
    end;

    pop :=TPopupMenu(GetControl('POPM') ) ;
    for i:=0 to pop.items.count-1 do
        begin
        if (ctxMode in V_PGI.PGIContexte) then      // particuliers pour la MODE
            begin
            if (pop.items[i].name='MnPieceCours') then pop.items[i].visible:=False;
            if (pop.items[i].name='MnArticleCom') then pop.items[i].visible:=False;
            if (pop.items[i].name='MnJustif') then pop.items[i].visible:=False;
            if (pop.items[i].name='MnVteCli') then pop.items[i].visible:=True;
            if (pop.items[i].name='MnArt') then pop.items[i].visible:=True;
            if (pop.items[i].name='MnWeb') then pop.items[i].enabled:=False;
            end else
            begin                                     // particuliers pour GESCOM
            if (pop.items[i].name='MnPieceCours') and (stNatureAuxi<>'CON') then
               pop.items[i].visible:=True;
            if not (ctxAffaire in V_PGI.PGIContexte) then
               if (pop.items[i].name='MnArticleCom') and (stNatureAuxi<>'CON') then
               pop.items[i].visible:=True;
            // Modifs JT 09/09/2003 - eQualité n° 10331
            //if (pop.items[i].name='MnJustif') then pop.items[i].visible:=True;
            if (pop.items[i].name='MnJustif') then pop.items[i].visible:=False;
            if (pop.items[i].name='MnVteCli') then pop.items[i].visible:=False;
            if (pop.items[i].name='MnArt') then pop.items[i].visible:=False;
            if (pop.items[i].name='MnWeb') then pop.items[i].enabled:=False;
            end;
        end;
    if (ctxMode in V_PGI.PGIContexte) then SetControlVisible('BPOP2MENU', False) ;
    end else
    begin
    libel := TraduireMemoire('Entreprise : ') + GetField('T_JURIDIQUE') + ' '+ GetField('T_LIBELLE');
    Ecran.Caption:=libel;
    If ctxSCot in V_PGI.PGIContexte  then begin
      SetControlVisible ('T_JOURNAISSANCE', False);
      SetControlVisible ('T_MOISNAISSANCE', False);
      SetControlVisible ('T_ANNEENAISSANCE', False);
      SetControlVisible ('T_DATNAISSANCE', False);
      SetControlVisible ('SEP_MOISNAISSANCE', False);
      SetControlVisible ('SEP_ANNEENAISSANCE', False);
      SetControlVisible ('TT_PRENOM',False);
      SetControlVisible ('T_SIRET', true);
      SetControlVisible ('TT_SIRET', True);
       end
    else begin
      SetControlVisible ('T_JOURNAISSANCE', FALSE);
      SetControlVisible ('T_MOISNAISSANCE', FALSE);
      SetControlVisible ('T_ANNEENAISSANCE', FALSE);
      SetControlVisible ('T_DATNAISSANCE', FALSE);
      SetControlVisible ('SEP_MOISNAISSANCE', FALSE);
      SetControlVisible ('SEP_ANNEENAISSANCE', FALSE);
      SetControlVisible ('TT_PRENOM', FALSE);
      SetControlVisible ('T_SIRET', TRUE);
      SetControlVisible ('TT_SIRET', TRUE);
      end;
//Non en line
    SetControlVisible ('T_SOCIETEGROUPE', True);
    SetControlVisible ('TT_SOCIETEGROUPE', True);
    if (stNatureAuxi<>'CON') then
    begin
      SetControlVisible ('T_CONFIDENTIEL', True);
      SetControlVisible ('TT_CONFIDENTIEL', True);
    end;
    SetControlVisible ('T_PASSWINTERNET', True);
    SetControlVisible ('TT_PASSWINTERNET', True);
    SetControlVisible ('T_ENSEIGNE', True);
    SetControlVisible ('TT_ENSEIGNE', True);
    SetControlVisible ('T_RVA', True);
    SetControlVisible ('TT_RVA', True);
    If ctxSCot in V_PGI.PGIContexte  then
    begin // mcd 19/09/03 nongéré dans un premier temps. attente réflexion avec enseigne annuaire
      SetControlVisible ('T_ENSEIGNE', False);
      SetControlVisible ('TT_ENSEIGNE', False);
    end;
    SetControlVisible ('T_FORMEJURIDIQUE', True);
    SetControlVisible ('TT_FORMEJURIDIQUE', True);
    SetControlVisible ('GBIDENTIFIANT', TRUE);
    SetControlVisible ('T_NIF', TRUE);
    SetControlVisible ('TT_NIF', TRUE);
    SetControlVisible ('T_EAN', TRUE);
    SetControlVisible ('TT_EAN', TRUE);
    SetControlVisible ('T_APE', TRUE);
    SetControlVisible ('TT_APE', TRUE);
    SetControlVisible ('TT_APE_', TRUE);
    // grc
    if not CritereSuspect then
       SetControlProperty ('T_JURIDIQUE', 'DataType', 'TTFORMEJURIDIQUE');
//    TTAbSheet(GetControl('PREGLEMENT')).TabVisible := TRUE;
//    SetControlVisible ('GB_TARIFICATION', FALSE);
    SetControlText ('TT_LIBELLE', TraduireMemoire('&Raison sociale'));
    SetControlText ('TT_TELEPHONE', TraduireMemoire('&Téléphone'));
      SetControlText ('TT_TELEX', TraduireMemoire('Tél portable'));
    SetControlText ('TT_FAX', TraduireMemoire('&Fax'));
//    SetControlText ('TT_TELEX', TraduireMemoire('&Minitel'));
    SetControlText ('TT_RVA', TraduireMemoire('&Site Web'));
  //mcd 25/04/01  SetControlText ('TT_JURIDIQUE', '&Forme juridique');
    SetControlText ('TT_JURIDIQUE', TraduireMemoire('&Abréviat. postale'));
    SetControlVisible ('T_SEXE', FALSE);
    SetControlVisible ('TT_SEXE', FALSE);
    SetControlVisible ('GBPARTICULIER', FALSE);
//    SetControlVisible ('BNEWMAIL', FALSE);
    if (ctxCHR in V_PGI.PGIContexte) then    //CHR
    begin
         SetControlText ('TT_TELEPHONE2', TraduireMemoire('Autres'));
//         SetControlVisible ('BNEWMAIL', TRUE);
         SetControlVisible ('TT_COMPTATIERS' , True);
         SetControlVisible ('T_COMPTATIERS' , True);
         SetControlVisible ('T_PRENOM', FALSE);
         SetControlText ('TT_LIBELLE', TraduireMemoire('Raison sociale'));
         SetControlText ('TT_TELEPHONE', TraduireMemoire('Téléphone'));
         SetControlText ('TT_FAX', TraduireMemoire('Fax'));
         SetControlText ('TT_TELEX', TraduireMemoire('Minitel'));
         SetControlText ('TT_JURIDIQUE', TraduireMemoire('Forme juridique'));
    end;
    pop :=TPopupMenu(GetControl('POPM') ) ;
    for i:=0 to pop.items.count-1 do
        begin
        // FQ Mode  10112 AC
        if (ctxMode in V_PGI.PGIContexte) then      // Fournisseur pour la MODE
            begin
            if (pop.items[i].name='MnPieceCours') then pop.items[i].visible:=True;
            if (pop.items[i].name='MnArticleCom') then pop.items[i].visible:=True;
            if (pop.items[i].name='MnJustif') then pop.items[i].visible:=False;
            if (pop.items[i].name='MnVteCli') then pop.items[i].visible:=False;
            if (pop.items[i].name='MnArt') then pop.items[i].visible:=False;
            if (pop.items[i].name='MnWeb') then pop.items[i].enabled:=True;
            end else
            begin                                     // particuliers pour GESCOM
            if (pop.items[i].name='MnPieceCours') and (stNatureAuxi<>'CON') then pop.items[i].visible:=True;
            if not (ctxAffaire in V_PGI.PGIContexte) then
               if (pop.items[i].name='MnArticleCom') and (stNatureAuxi<>'CON') then pop.items[i].visible:=True;
            //Modifs JT 09/09/2003 - eQualité n° 10331
            //if (pop.items[i].name='MnJustif') then pop.items[i].visible:=True;
            if (pop.items[i].name='MnJustif') then pop.items[i].visible:=False;
            if (pop.items[i].name='MnVteCli') then pop.items[i].visible:=False;
            if (pop.items[i].name='MnArt') then pop.items[i].visible:=False;
            if (pop.items[i].name='MnWeb') then pop.items[i].enabled:=True;
            end;
        end;
    if (ctxMode in V_PGI.PGIContexte) then SetControlVisible('BPOP2MENU', True) ;
    end;
if (ctxChr in V_PGI.PGIContexte) then
  begin      // particuliers pour CHR
  SetControlVisible('BPOP2MENU', True);
      for i:=0 to pop.items.count-1 do
      begin
          if (pop.items[i].name='MnPieceCours') and (stNatureAuxi<>'CON') then pop.items[i].visible:=True;
          if (pop.items[i].name='MnArticleCom') then pop.items[i].visible:=False;
          if (pop.items[i].name='MnJustif') then pop.items[i].visible:=True;
          if (pop.items[i].name='MnVteCli') then pop.items[i].visible:=False;
          if (pop.items[i].name='MnArt') then pop.items[i].visible:=False;
          if (pop.items[i].name='MnWeb') then pop.items[i].enabled:=True;
          if (pop.items[i].name='MnEncours') then pop.items[i].enabled:=False;
          if (pop.items[i].name='MnTarif') then pop.items[i].enabled:=False;
          if (pop.items[i].name='mnPrescripteur') then pop.items[i].enabled:=False;
      end;
    end;
{$IFDEF BTP}
  pop :=TPopupMenu(GetControl('POPM') ) ;
  for i:=0 to pop.items.count-1 do
  begin
    if (pop.items[i].name='MnPieces') then pop.items[i].visible:=False;
    if (pop.items[i].name='MnTarifAvance') then pop.items[i].visible:=False;
    if (pop.items[i].name='MnWeb') then pop.items[i].visible:=False;
    if (pop.items[i].name='MnAnalyses') then pop.items[i].visible:=False;
    if (pop.items[i].name='MnSavParc') then pop.items[i].visible:=False;
    if (pop.items[i].name='MnTb') then pop.items[i].visible:=false;
    if (pop.items[i].name='mnAnalyses') then pop.items[i].visible:=false;
    if (pop.items[i].name='MnSAVParc') then pop.items[i].visible:=false;
    if (pop.items[i].name='mnVisuPiece') then pop.items[i].visible:=false;
  end;
{$ELSE}
  pop :=TPopupMenu(GetControl('POPM') ) ;
  for i:=0 to pop.items.count-1 do
  begin
    if (pop.items[i].name='MnBordereaux') then pop.items[i].visible:=False;
  end;
{$ENDIF}
TFFiche(Ecran).Refresh;
end;
//**************************** specifique MODE *********************

/////////////////////////////////////////////
// ****** TOM TiersPiece ********************
/////////////////////////////////////////////
procedure TOM_TiersPiece.OnArgument (Arguments : String );
begin
inherited ;
if CtxScot in V_PGI.PgiContexte then
   begin    //mcd 17/07/03 zones non gérées en GI
   SetControlVisible  ('GTP_RELIQUAT',False);
   SetControlVisible  ('TGTP_RELIQUAT',False);
   SetControlVisible  ('GTP_REGROUPE',False);
   SetControlVisible  ('GTP_SUSPENSION',False);
   end;
end;
procedure TOM_TiersPiece.SetLastError (Num : integer; ou : string );
begin
if ou<>'' then SetFocusControl(ou);
LastError:=Num;
LastErrorMsg:=TexteMessage[LastError];
end ;

/////////////////////////////////////////////////////////////////////////////

procedure TOM_TiersPiece.OnNewRecord ;
begin
     SetField('GTP_REGROUPE','-') ;
     SetField('GTP_SUSPENSION','-') ;
     SetField('GTP_VISA','-') ;
end;

/////////////////////////////////////////////////////////////////////////////

procedure TOM_TiersPiece.OnChangeField (F : TField)  ;
begin
inherited;
If (F.FieldName='GTP_NATUREPIECEG') and (F.Value<>'') then
If  GetField('GTP_LIBELLE')='' then
    BEGIN
    SetField('GTP_LIBELLE',RechDom('GCNATUREPIECEG',GetField('GTP_NATUREPIECEG'),FALSE));
    END ;
if (F.FieldName = 'GTP_MONTANTVISA') and (F.Value < 0.0) then begin SetLastError(14, 'GTP_MONTANTVISA'); exit ; end ;
end;

/////////////////////////////////////////////////////////////////////////////

procedure TOM_TiersPiece.OnUpdateRecord ;
begin
inherited;
if (GetField('GTP_NATURESUIVANTE') = GetField('GTP_NATUREPIECEG')) then
    begin SetLastError(11, 'GTP_NATURESUIVANTE'); exit ; end ;
if (GetField ('GTP_MONTANTVISA') < 0.0) then  begin SetLastError(14, 'GTP_MONTANTVISA'); exit ; end ;
If ReferenceCirculaire then exit;
end;

/////////////////////////////////////////////////////////////////////////////

function TOM_TiersPiece.ReferenceCirculaire : Boolean;
Var NatSuiv, RefCirc: String ;
    erreur : boolean ;
    Q_Nature : TQuery ;
    i_ind, nbenreg : integer;
    T : array of string ;
begin
Result:=True ;
erreur:=False ;
nbenreg:=1;
Setlength(T,nbenreg);
T[0]:=GetField('GTP_NATUREPIECEG');
NatSuiv:=GetField('GTP_NATURESUIVANTE');
Repeat
    for i_ind:=0 to high(T) do
        if T[i_ind]=NatSuiv then erreur:=True ;
    inc(nbenreg) ;
    Setlength(T,nbenreg) ;
    T[high(T)]:=NatSuiv ;
    Q_Nature:=OpenSQL('SELECT GTP_NATURESUIVANTE FROM TIERSPIECE ' +
                      'WHERE GTP_TIERS="'+GetField('GTP_TIERS')+
                      '" AND GTP_NATUREPIECEG="'+ NatSuiv +'"',True);
    If not Q_Nature.Eof then NatSuiv:=Q_Nature.FindField('GTP_NATURESUIVANTE').AsString
    else NatSuiv:='' ;
    Ferme(Q_Nature) ;
until (NatSuiv='') or (erreur) ;
    if erreur then
    begin
    for i_ind:=0 to high(T) do
        begin
        RefCirc:=RefCirc+T[i_ind] ;
        if i_ind<high(t) then RefCirc:=RefCirc+' -> ';
        end ;
    SetLastError(12, 'GTP_NATURESUIVANTE');
    i_ind := pos ('XXX', LastErrorMsg);
    if i_ind > 0 then
            LastErrorMsg := StringReplace (LastErrorMsg,'XXX',RefCirc, [rfIgnoreCase]);
    Result := False;
    end;
end;

procedure TOM_TIERS.OnDeleteRecord;
begin
  inherited;
  {IFDEF CHR}
  if ExisteSQL('select hdr_dosres from hrdosres where hdr_tiers="'+Getfield('T_TIERS')+'"')then
  begin
       LastError:=100 ;
       LastErrorMsg:='Tiers utilisé, Suppression impossible';
       exit;
  end;

  if ExisteSQL('select gp_naturepieceg from piece where gp_tiers="'+Getfield('T_TIERS')+'"') then
  begin
       LastError:=100 ;
       LastErrorMsg:='Tiers utilisé, Suppression impossible';
       exit;
  end;

  if ExisteSQL('select htd_typdos from hrtypdos where htd_tiersdefaut="'+Getfield('T_TIERS')+'"') then
  begin
       LastError:=100 ;
       LastErrorMsg:='Tiers utilisé, Suppression impossible';
       exit;
  end;

{ENDIF}
end;

// debut Affaire
procedure TOM_TIERS.OnClose ;
begin
inherited;
if HandleSeria <> 0 then FreeLibrarySeriaLSE (HandleSeria);
{$IFDEF CTI}
// Appels en série : on sort sans avoir raccroché : on empêche de sortir
if (AppelCtiOk) and (SerieCTI) and (not CtiRaccrocherFiche) then
    begin
    LastError := 1;
    exit;
    end
else
    begin
    // Serie Appels CTI : pas d'Appel, on génère l'action
    { mng 31/08/2004 : finalement non, on considère que femer la fiche n'est pas
      un appel sortant non abouti, donc pas d'action générée
    if (GetParamSoc('SO_RTCTIGESTION')) and (SerieCTI) and (not AppelCtiOk) then
      begin
        CtiModeAppel:=1;
        RTCTIGenereAction(AppelCtiOk,GetField('T_TIERS'),GetField('T_AUXILIAIRE'),CtiHeureDebCom,CtiHeureFinCom-CtiHeureDebCom,'',CtiModeAppel,0);
      end;}
    // Appels Sortant seul : on sort sans avoir raccroché : on mémorise l'heure de début de Communication
    if (AppelCtiOk) and (not SerieCTI) and (not CtiRaccrocherFiche) then
        CtiHeureDeb:=CtiHeureDebCom;
    end;
if RTFormCti<> Nil then RTFormCti:=Nil;
{$ENDIF}
TobContact.free; 
TobZonelibre.free;
if TobSuspect<>nil then begin TobSuspect.free ; TobSuspect:=nil ; end ;
if TOBSuspectCompl<>nil then begin TOBSuspectCompl.free ; TOBSuspectCompl:=nil ; end ;
{$IFDEF BTP}
if TOBEtude <> nil then TOBEtude.free;
{$ENDIF}
  TOBTIERSBTP.free;
end;
// fin Affaire

Procedure TOM_TIERS.RechDPBis(Var CodePer : integer; Tiers, NatAuxi : String) ;
var Q1,Q2  : TQuery;
StCodePer : string;
begin
//MCD 07/08/00 Mis dans ce soure en attendant que SetField ne soit plus
// Protected dans L'AGL. Dans ce cas, sera fait en extérieur !!!!
If (NatAuxi <> 'CLI') and (NatAuxi <> 'FOU') and (NatAuxi <> 'PRO') and (NatAuxi <> 'CON') // GRC
then exit;
    // on interdit de ne pas sélectionner de code personne
    // En cas de création, obligation de le faire dans le mul annuaire
repeat
    StCodePer:=AglLanceFiche('YY', 'ANNUAIRE_SEL','ANN_NOMPER='+GetField('T_ABREGE'),'','TIERS');
    until StCodePer <> '';

if StCodePer <>'' then
    begin
    CodePer :=StrtoInt(StCodePer);
        //ATTENTION ALigner avec Synchronise de AnnOutils si modif
    Q1 := OpenSQL('select * from ANNUAIRE where ANN_CODEPER = '
        + StCodePer, TRUE);
    if not Q1.eof then
        begin
        SetField ('T_NATIONALITE',Q1.FindField('ANN_NATIONALITE').Asstring);
        SetField ('T_LANGUE',Q1.FindField('ANN_LANGUE').Asstring);
        SetField ('T_DEVISE',Q1.FindField('ANN_DEVISE').Asstring);
        SetField ('T_FAX',Q1.FindField('ANN_FAX').Asstring);
        SetField ('T_MOISCLOTURE',Q1.FindField('ANN_MOISCLOTURE').Asstring);
        SetField ('T_ABREGE',Q1.FindField('ANN_NOMPER').Asstring);
        SetField ('T_RVA',Q1.FindField('ANN_SITEWEB').Asstring);
        SetField ('T_TELEPHONE',Q1.FindField('ANN_TEL1').Asstring);
        SetField ('T_TELEX',Q1.FindField('ANN_MINITEL').Asstring);
        SetField ('T_APE',Q1.FindField('ANN_CODENAF').Asstring);
        SetField ('T_SIRET',Q1.FindField('ANN_SIREN').Asstring+Q1.FindField('ANN_CLESIRET').Asstring);
        SetField ('T_LIBELLE',Q1.FindField('ANN_NOM1').Asstring);
        if (Q1.FindField('ANN_PPPM').Asstring = 'PP') then Setfield('T_PARTICULIER','X')
           else setfield('T_PARTICULIER','-');
        SetField ('T_PRENOM',Q1.FindField('ANN_NOM2').Asstring);
        SetField ('T_ADRESSE1',Q1.FindField('ANN_ALRUE1').Asstring);
        SetControlText ('T_ADRESSE1',Q1.FindField('ANN_ALRUE1').Asstring);
        SetField ('T_ADRESSE2',Q1.FindField('ANN_ALRUE2').Asstring);
        SetField ('T_ADRESSE3',Q1.FindField('ANN_ALRUE3').Asstring);
        SetField ('T_VILLE',Q1.FindField('ANN_ALVILLE').Asstring);
        SetField ('T_PAYS',Q1.FindField('ANN_PAYS').Asstring);
        SetField ('T_CODEPOSTAL',Q1.FindField('ANN_ALCP').Asstring);
        // T_JURIDIQUE et ANN_FORME
        Q2 := OpenSQL('select JFJ_CODEDP from JUFORMEJUR where JFJ_FORME = '
          + '"'+Q1.FindField('ANN_FORME').Asstring+'"', TRUE);
        if not Q2.eof then begin
           Setfield('T_FORMEJURIDIQUE', Q2.FindField('JFJ_CODEDP').Asstring);
           end;
        Ferme (Q2);
        end;
    Ferme (Q1);
    end;
end;

// fin Affaire

Procedure TOM_TIERS.ModifParticulier;
var Bouton : TradioGroup;
BEGIN
// cette fct est appelé depuis le script TIERS
Bouton := TradioGroup(GetControl ('T_PARTICULIER'));

//FV1 - 29/08/2017 - FS#2652 - LAFOSSE : pb en duplication de fournisseur
if bouton = nil then exit;

if (Bouton.Itemindex = 0) then
   begin
   SetField('T_PARTICULIER','X');
   AfficheEcranClientMode(TRUE);
   end else
   begin
   SetField('T_PARTICULIER','-');
   AfficheEcranClientMode(FALSE);
   end;
END;

Procedure TOM_TIERS.AffIndisponible (AffMes : boolean);
BEGIN
if (GetField('T_FERME')='X') then
    begin
    if (AffMes = TRUE) then PGIBox('le tiers choisi est fermé','Attention!');
    SetControlVisible('T_DATEFERMETURE',True);
    SetControlVisible('TT_DATEFERMETURE',True);
    SetControlVisible('T_DATEOUVERTURE',False);
    SetControlVisible('TT_DATEOUVERTURE',False);
    end else
    begin
    SetControlVisible('T_DATEFERMETURE',False);
    SetControlVisible('TT_DATEFERMETURE',False);
    SetControlVisible('T_DATEOUVERTURE',True);
    SetControlVisible('TT_DATEOUVERTURE',True);
    end;
END;

Procedure TOM_TIERS.MajIndisponible;
var Bouton : TCheckBox;
BEGIN
// cette fct est appelé depuis le script TIERS
Bouton := TCheckBox(GetControl ('T_FERME'));
if (Bouton.Checked = True) then SetField('T_FERME','X') else SetField('T_FERME','-');
AffIndisponible (False);
END;

Procedure TOM_TIERS.GestionBoutonGED;
BEGIN
if Assigned(GetControl('BDOCGEDEXIST')) then
   begin
   if (GetParamSoc('SO_RTGESTIONGED')) and (ExisteSQL('SELECT RTD_DOCID FROM RTDOCUMENT WHERE RTD_TIERS="'+GetField('T_TIERS')+'"')) then SetControlVisible ('BDOCGEDEXIST',True)
   else SetControlVisible ('BDOCGEDEXIST',False);
   end;
END;

procedure AGLDuplication( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS) then TOM_TIERS(OM).Duplication('') else exit;
end;

procedure AGLModifParticulier( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS)
    then TOM_TIERS(OM).ModifParticulier
    else exit;
end;

procedure AGLmajcontact( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS)
    then TOM_TIERS(OM).majcontact
    else exit;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Pascal Gautherot
Créé le ...... : 16/09/2003
Modifié le ... :   /  /
Description .. : Procédure permetant, depuis le scripts, de rafraichire les
Suite ........ : données du contact principal affichées dasn la fiche tiers.
Suite ........ : Appel AfficheContactTiers
Mots clefs ... : CONTACT TIERS
*****************************************************************}
procedure AGLAfficheContactTiers( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS) then TOM_TIERS(OM).AfficheContactTiers else exit;
end;

Procedure TOM_TIERS.ZonesLibresIsModified ;
BEGIN
TobZoneLibre.GetEcran (TFfiche(Ecran),Nil);
if TobZoneLibre.IsOneModifie and  not(DS.State in [dsInsert,dsEdit])then
    begin
    DS.edit; // pour passer DS.state en mode dsEdit
{$IFDEF EAGLCLIENT}
    TFFiche(Ecran).QFiche.CurrentFille.Modifie:=true;
{$ELSE}
    SetField ('T_TIERS', GetControlText ('T_TIERS'));
{$ENDIF}
    end;
END;

Procedure TOM_TIERS.TestModifZoneLibre (champ : STring);
BEGIN
    // cette fct est appelé depuis le script TIERS si modif zone libre
{if ( GetControlText(champ) <> TobZonelibre.Getvalue (champ )) then
  if not(DS.State in [dsInsert,dsEdit])
    then begin
    DS.edit; // pour passer DS.state en mode dsEdit
    SetField ('T_TIERS', GetControlText ('T_TIERS'));
    end;}
TobZoneLibre.GetEcran (TFfiche(Ecran),Nil);
if TobZoneLibre.IsOneModifie and  not(DS.State in [dsInsert,dsEdit])then
    begin
    DS.edit; // pour passer DS.state en mode dsEdit
    SetField ('T_TIERS', GetControlText ('T_TIERS'));
    end;
END;

Procedure TOM_TIERS.TestModifZoneLibre2 (champ : STring);
BEGIN
    // cette fct est appelé depuis le script TIERS si modif zone date
if ( StrtoDate(GetControlText(champ)) <> TobZonelibre.Getvalue (champ )) then
  if not(DS.State in [dsInsert,dsEdit])
    then begin
    DS.edit; // pour passer DS.state en mode dsEdit
    SetField ('T_TIERS', GetControlText ('T_TIERS'));
    end;
END;
Procedure TOM_TIERS.TestModifFacture ;
BEGIN
    // cette fct est appelé depuis le script TIERS si modif zone ressource
if ( GetControlText('TE_FACTURE') <> Old_Facture) then
  if not(DS.State in [dsInsert,dsEdit])
    then begin
    DS.edit; // pour passer DS.state en mode dsEdit
    SetField ('T_TIERS', GetControlText ('T_TIERS'));
    end;
END;

Procedure TOM_TIERS.TestModifLivre;
begin
   // cette fct est appelé depuis le script TIERS si modification de la zone tiers livré
   if (GetControlText('TE_LIVRE') <> Old_Livre) then
   begin
      { Raz de l'adresse de livraison }
      TobZonelibre.PutValue('YTC_NADRESSELIV', 0);
      SetControlText('YTC_NADRESSELIV', '0');
      { Passe le DataSet en mode Edit }
      if not(DS.State in [dsInsert,dsEdit]) then
      begin
         DS.Edit;
         SetField ('T_TIERS', GetControlText ('T_TIERS'));
      end;
      { Propriétées Plus de l'adresse de livraison }
   end;
end;

Procedure TOM_TIERS.TestModifPayeur ;
BEGIN
    // cette fct est appelé depuis le script TIERS si modif zone ressource
if ( GetControlText('TE_PAYEUR') <> Old_Payeur) then
  begin
  if not(DS.State in [dsInsert,dsEdit]) then
     begin
     DS.edit; // pour passer DS.state en mode dsEdit
     SetField ('T_TIERS', GetControlText ('T_TIERS'));
     end;

  if GetControlText('TE_PAYEUR')='' then
      begin
      SetControlEnabled('T_ISPAYEUR',TRUE);
      end else
      begin
      SetControlChecked('T_ISPAYEUR',FALSE);
      SetControlEnabled('T_ISPAYEUR',FALSE);
      end;
  end;

{    if GetControlText('TE_PAYEUR')='' then
          begin
          SetControlEnabled('T_ISPAYEUR',TRUE);
          end
    else begin
          if GetControlText ('TE_PAYEUR')=GetField ('T_TIERS') then SetControlText('TE_PAYEUR', '');
          ExisteTiers := RechercheTiers(getControlText('TE_PAYEUR'),stNatureAuxi, 'X', LibelleTiers);
          if ExisteTiers then
            begin
            SetControlChecked('T_ISPAYEUR',FALSE);
            SetControlEnabled('T_ISPAYEUR',FALSE);
            end else
            begin
            SetControlEnabled('T_ISPAYEUR',TRUE);
            SetLastError(3, 'T_PAYEUR');
            exit;
            end;
         end;
    end;   }
END;

procedure TOM_TIERS.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
Case Key of
    VK_F12 : {encours} if ((ssAlt in Shift) and (ssShift in Shift)) then begin; VoirDetailEncours ; end
                       else if (ssAlt in Shift) then begin Key:=0 ; VoirEncours; end;
    VK_F7 : {Actions} if (ssAlt in Shift) then
                       begin Key:=0 ; AGLLanceFiche('RT','RTACTIONS_TIERS','RAC_TIERS='+GetField('T_TIERS'),'',ActionToString(TFFiche(Ecran).FtypeAction)+';NOCHANGEPROSPECT') ;  end;
    VK_F9 : {Soldes}  if (ssAlt in Shift) then BEGIN Key:=0 ; CalculSoldesAuxi(GetField('T_AUXILIAIRE')) ; END ;
END;
if (ecran <> nil) then TFFiche(ecran).FormKeyDown(Sender,Key,Shift);
end;

// <<<<<<<<<<<<<< Fonctions pour script AGL >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

procedure AGLZonesLibresIsModified( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS)
    then TOM_TIERS(OM).ZonesLibresIsModified
    else exit;
end;

procedure AGLTestModifZonelibre( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS)
    then TOM_TIERS(OM).TestModifZoneLibre(Parms[1])
    else exit;
end;

procedure AGLTestModifZonelibre2( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS)
    then TOM_TIERS(OM).TestModifZoneLibre2(Parms[1])
    else exit;
end;

procedure AGLTestModifLivre( parms: array of variant; nb: integer ) ;
var
   F: TForm;
   OM: TOM;
begin
   F := TForm(Longint(Parms[0])) ;
   if (F is TFFiche) then
   begin
      OM := TFFiche(F).OM;
      if (OM is TOM_TIERS) then
         TOM_TIERS(OM).TestModifLivre();
   end;
end;

procedure AGLTestModifFacture( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS)
    then TOM_TIERS(OM).TestModifFacture()
    else exit;
end;

procedure AGLTestModifPayeur( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS)
    then TOM_TIERS(OM).TestModifPayeur()
    else exit;
end;

procedure AGLMajIndisponible( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS) then TOM_TIERS(OM).MajIndisponible else exit;
end;

// GRC
procedure TOM_TIERS.TOBCopieChamp(FromTOB, ToTOB : TOB);
var i_pos,i_ind1: integer;
    FieldNameTo,FieldNameFrom,St:string;
    PrefixeTo,PrefixeFrom : string;
begin
PrefixeFrom := TableToPrefixe (FromTOB.NomTable);
PrefixeTo := TableToPrefixe (ToTOB.NomTable);
for i_ind1 := 1 to FromTOB.NbChamps do
  begin
    FieldNameFrom := FromTOB.GetNomChamp(i_ind1);
    St := FieldNameFrom ;
    i_pos := Pos ('_', St) ;
    Delete (St, 1, i_pos-1) ;
    FieldNameTo := PrefixeTo + St ;
    if ToTOB.FieldExists(FieldNameTo) then
      ToTOB.PutValue(FieldNameTo, FromTOB.GetValue(FieldNameFrom))
  end;
end;


// version provisoire du calcul du risque client
function  TOM_TIERS.RisqueTiers: String;
var RisqueTiers, P1, P2 : double;
    TobTiers: TOB;
begin
if stNatureAuxi = 'CON' then begin result:=''; exit; end;
if (ds<>nil) and (DS.State in [dsInsert]) then begin result:='V'; exit; end;
// L'état du risque calculé est stocké dans EtatRisque. La fonction retourne l'état réelle qui a pu être forcé
EtatRisque:='V' ;
// DCA - FQ MODE 10855 - Eviter requête inutile
P1:=getField('T_CREDITACCORDE') ;
P2:=getField('T_CREDITPLAFOND') ;
if (P1>0) or (P2>0) then
  begin
  TOBTiers:=Tob.create('TIERS',Nil,-1);
  TOBTiers.selectDB ('',TFFiche(Ecran).QFiche,true) ;
  RisqueTiers:=RisqueTiersGC(TOBTiers)+RisqueTiersCPTA(TOBTiers,V_PGI.DateEntree) ;
  if ((P1>0 ) and (RisqueTiers > P1)) then EtatRisque:='O' ;
  if ((P2>0 ) and (RisqueTiers > P2)) then EtatRisque:='R' ;
  TobTiers.free;
  end;
Result:=EtatRisque;
if getField('T_ETATRISQUE')<>'' then
   begin
   Result:=getField('T_ETATRISQUE');
   end;
end;

procedure TOM_TIERS.VoirEncours;
var TobTiers : TOB ;
    VOR, OkOk : string ;
    Action : TActionFiche ;
begin
Action:=TFFiche(Ecran).FtypeAction;
TOBTiers:=Tob.create('TIERS',Nil,-1);
TOBTiers.selectDB ('',TFFiche(Ecran).QFiche,True) ;
CalculSoldesAuxi(TOBTiers.GetValue('T_AUXILIAIRE')) ;
TobTiers.LoadDB(True) ; // Pour recharger les données calculées
TheTob:=TobTiers;
OkOk:=AglLanceFiche('GC', 'GCENCOURS','','',ActionToSTring(Action));
if (OkOk='OK') and (Action<>taConsult) then
   begin
   VOR:=tobTiers.getvalue('T_ETATRISQUE');
   if VOR<>'' then setControlText('RISQUE',VOR) else setControlText('RISQUE',EtatRisque);
   setControlText('T_ETATRISQUE',VOR) ;
   if not(DS.State in [dsInsert,dsEdit])then
     begin
     DS.edit;
     SetField('T_ETATRISQUE',VOR);
     end;
   end;
TobTiers.free;
end;

procedure TOM_TIERS.VoirDetailEncours;
var TobTiers : TOB ;
    Action : TActionFiche ;
begin
Action:=TFFiche(Ecran).FtypeAction;
TOBTiers:=Tob.create('TIERS',Nil,-1);
TOBTiers.selectDB ('',TFFiche(Ecran).QFiche,True) ;
CalculSoldesAuxi(TOBTiers.GetValue('T_AUXILIAIRE')) ;
TobTiers.LoadDB(True) ; // Pour recharger les données calculées
TheTob:=TobTiers;
AGLLanceFiche ('GC','GCENCOURSGC','','',ActionToString(Action)) ;
TobTiers.free;
end;

//******************************************************************************
// ************************ Gestion Isoflex ************************************
// *****************************************************************************
procedure TOM_TIERS.GereIsoflex;
var bIso : Boolean ;
    MenuIso : TMenuItem ;
begin
MenuIso := TMenuItem(GetControl('mnSGED')) ;
{$IFNDEF EAGLCLIENT}
bIso := AglIsoflexPresent ;
{$ELSE}
bIso := False ;
{$ENDIF}
if MenuIso <> Nil then MenuIso.Visible := bIso ;
end;

procedure TOM_Tiers_AppelIsoFlex ( parms: array of variant; nb: integer ) ;
{$IFNDEF EAGLCLIENT}
var  F : TForm ;
     Tiers : string;
{$ENDIF EAGLCLIENT}
begin
{$IFNDEF EAGLCLIENT}
F:=TForm(Longint(Parms[0])) ;
if (F.Name<>'GCTIERS') and (F.Name<>'GCFOURNISSEUR') then exit;
Tiers:=string(Parms[1]);
AglIsoflexViewDoc(NomHalley,F.Name,'TIERS','T_CLE2','T_TIERS',Tiers, '');
{$ENDIF}
end;



Function  AGLVORRisqueTiers( parms: array of variant; nb: integer ) : variant ;
var  F : TForm ;
     OM : TOM ;
begin
Result:='V';
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS) then result:=TOM_TIERS(OM).RisqueTiers else exit;
end;

Procedure  AGLGCVoirEncours( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS) then TOM_TIERS(OM).VoirEncours else exit;
end;

Function AGLPCSPhoneme( parms: array of variant; nb: integer ) : variant ;
begin
result:=RTPhonemeSearch(string(Parms[0]));
end;

procedure AGLRTTiersImprime( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     Tiers, stWhere, Modele : string;
     NatureAuxi,StArgument : string;
     OptionEdit : boolean;
begin
F:=TForm(Longint(Parms[0])) ;
//uniquement en line
//TFFiche(F).BImprimerClick(Nil) ;

Tiers:=string(Parms[1]);
NatureAuxi:=string(Parms[2]);
stWhere:=' T_TIERS="'+Tiers+'"';
Modele:=GetParamSoc('SO_GCETATFICHETIERS');
OptionEdit:=GetParamSoc('SO_GCFICHEOPTIONEDIT');
//if (Modele<>'') and (not(ctxaffaire in V_PGI.PGIContexte)) and (NatureAuxi<>'CON') then  //AB-20031021
if (Modele<>'') and (pos(NatureAuxi,'CON;FOU')=0) then
   begin
   if OptionEdit then
      begin
      StArgument := '';
      if (not(ctxGRC in V_PGI.PGIContexte)) then StArgument := 'NOTGRC';
{   StOptionEdit := AGLLanceFiche ('RT','RTPARAMEDITFICHE','','',StArgument);
   if StOptionEdit <> '' then
       begin
       StDateAct := ReadTokenSt(StOptionEdit);
       StEditAct := ReadTokenSt(StOptionEdit);
       StDatePropo := ReadTokenSt(StOptionEdit);
       StEditPropo := ReadTokenSt(StOptionEdit);
       StEditBN := ReadTokenSt(StOptionEdit);
       StEditCont := ReadTokenSt(StOptionEdit);
       StEditEnc := ReadTokenSt(StOptionEdit);
       StEditEncG := ReadTokenSt(StOptionEdit);
    //   LanceEtat('E','RPF',Modele,True,False,False,Nil,stWhere,'',False,0,'XX_DATEACTION='+usdateTime(StrToDate(StDateAct))+'`XX_EDITACT='+StEditAct+'`XX_DATEPROPO='+usdateTime(StrToDate(StDatePropo))+'`XX_EDITPROPO='+StEditPropo+'`XX_EDITBN='+StEditBN+'`XX_EDITCONT='+StEditCont+'`XX_EDITENC='+StEditEnc+'`XX_EDITENCG='+StEditEncG);
       LanceEtat('E','RPF',Modele,True,False,False,Nil,stWhere,'',False,0,'XX_DATEACTION='+StDateAct+'`XX_EDITACT='+StEditAct+'`XX_DATEPROPO='+StDatePropo+'`XX_EDITPROPO='+StEditPropo+'`XX_EDITBN='+StEditBN+'`XX_EDITCONT='+StEditCont+'`XX_EDITENC='+StEditEnc+'`XX_EDITENCG='+StEditEncG);
       end; }
       AGLLanceFiche ('RT','RTPARAMEDITFICHE','T_TIERS='+Tiers,'',StArgument);
       end
   else LanceEtat('E','RPF',Modele,True,False,False,Nil,stWhere,'',False);
   end
else
    TFFiche(F).BImprimerClick(Nil) ;
end;

 //CHR
Function AGLRechercheClient (Parms : Array of variant; nb: integer) : variant;
var F           : TForm;
    G_Client : THCritMaskEdit;
    Q:TQuery;
BEGIN
F := TForm (Longint (Parms[0]));
G_Client := THCritMaskEdit (F.FindComponent (string (Parms [1])));
DispatchRecherche(G_Client, 2, '',
                   'T_TIERS=' + Trim (Copy (G_Client.Text, 1, 18)), '');
Q:=OpenSQL('Select T_TIERS from tiers where T_AUXILIAIRE="'+G_Client.text+'"',true);
if not Q.EOF  then G_Client.Text:=Q.findfield('T_TIERS').AsString;
ferme(Q);
Result := G_Client.Text;
END;

procedure AGLMajContact_chr( parms: array of variant; nb: integer ) ;//CHR
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS)
    then TOM_TIERS(OM).MajContact_chr
    else exit;
end;
procedure AGLMajAffichageKardexEntreprise( parms: array of variant; nb: integer ) ; //CHR
var  F : TForm ;
     OM : TOM ;
begin
    // fct qui publie la fct de ctrl modification ressource
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS)
    then TOM_TIERS(OM).MajAffichageKardexEntreprise
    else exit;
end;
procedure AGLRecupContact( parms: array of variant; nb: integer ) ;  //CHR
var  F : TForm ;
     OM : TOM ;
begin
    // fct qui publie la fct d'affichage des contacts
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS)
    then TOM_TIERS(OM).RecupContact
    else exit;
end;

procedure AGLRTTransProCli( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS) then TOM_TIERS(OM).RTTransProCli() else exit;
end;

{$IFDEF CTI}
procedure TOM_TIERS.RTCTIDecrocheAppel ;
begin
Fo_PBC.DecrocherClick (TForm (Ecran)) ;
CtiHeureDebCom := Time;
AppelCtiOk:=True;
Fo_PBC.AfficheCoordonnees (true,GetField('T_TIERS'),GetField('T_LIBELLE'),GetField('T_NATUREAUXI'),GetField('T_AUXILIAIRE'),'','') ;
CtiNumAct:=RTCTIGenereAction(AppelCtiOk,GetField('T_TIERS'),GetField('T_AUXILIAIRE'),CtiHeureDebCom,0,WhereUpdate,CtiModeAppel,0);
end;

procedure TOM_TIERS.RTCTIRaccrocheAppel ;
begin
CtiHeureFinCom:=Time;
SetControlProperty('BPHONE','Down',False);
Fo_PBC.RaccrocherClick (TForm (Ecran)) ;
// Serie ou Appel Sortant Seul : on génère l'action
// même sur appel entrant if (SerieCTI) or (CtiModeAppel=1) then
//RTCTIGenereAction(AppelCtiOk,GetField('T_TIERS'),GetField('T_AUXILIAIRE'),CtiHeureDebCom,CtiHeureFinCom-CtiHeureDebCom,WhereUpdate,CtiModeAppel);
RTCTIMajDureeAction (GetField('T_AUXILIAIRE'),CtiNumAct,CtiHeureFinCom-CtiHeureDebCom);
CtiHeureDeb:=0;
AppelCti:=False;
end;
procedure TOM_TIERS.RTCTINumeroter ;
begin
if GetField('T_CLETELEPHONE') <> '' then
    begin
    if not CtiCCA then
       Fo_PBC.RTCTINumeroteClient (TForm (Ecran),GetField('T_CLETELEPHONE'),GetField('T_TIERS'),GetField('T_LIBELLE'),GetField('T_NATUREAUXI'),GetField('T_AUXILIAIRE')) 
    else
       Fo_CCA.RTCTINumeroteClient (TForm (Ecran),GetField('T_CLETELEPHONE'),GetField('T_TIERS'),GetField('T_LIBELLE'),GetField('T_NATUREAUXI'),GetField('T_AUXILIAIRE')) ;
    if CtiErreurConnexion=True then
        begin
        RTCTIAfficheMessageCTI;
        SetControlProperty('BPHONE','Down',False);
        CtiErreurConnexion:=False;
        SetControlVisible('BATTENTE',False);
        if (GetParamSoc('SO_RTCTIGESTION')) and (not AppelCtiOk) then
          RTCTIGenereAction(AppelCtiOk,GetField('T_TIERS'),GetField('T_AUXILIAIRE'),CtiHeureDebCom,CtiHeureFinCom-CtiHeureDebCom,'',ModeAppelSortant,0);
        end
    else
        begin
        if not CtiCCA then
          begin
          SetControlProperty('BPHONE','Hint','Raccrocher');
          SetControlProperty('BPHONE','Down',True);
          SetControlText('INFOSCTI','Communication en cours');
          end
        else
          SetControlVisible('BATTENTE',False);
        AppelCtiOk:=True;
        CtiHeureDebCom := Time;
        if not CtiCCA then
           Fo_PBC.AfficheCoordonnees (false,GetField('T_TIERS'),GetField('T_LIBELLE'),GetField('T_NATUREAUXI'),GetField('T_AUXILIAIRE'),'','') ;
        CtiNumAct:=RTCTIGenereAction(AppelCtiOk,GetField('T_TIERS'),GetField('T_AUXILIAIRE'),CtiHeureDebCom,0,WhereUpdate,CtiModeAppel+CCAModeAppel,0);
        end;
    end
else
    begin
    PGIBox('Renseigner l''indicatif du pays dans les coordonnées, puis lancer l''utilitaire d''initialisation de la clé téléphone' ,'clé téléphone vide'); 
    if not CtiCCA then
      SetControlProperty('BPHONE','Down',False);
    end;
end;

procedure TOM_TIERS.RTCTIAppelAttente ;
begin
Fo_PBC.AttenteClick (TForm (Ecran)) ;
end;
procedure TOM_TIERS.RTCTIRepriseAppel ;
begin
Fo_PBC.RepriseClick (TForm (Ecran)) ;
end;


procedure TOM_TIERS.RTCTIAfficheMessageCTI;
var MessageCti: string;
begin
MessageCti := 'Communication Impossible';
case TResultatAppel(CtiResultatAppel) of
    raOccupe  : MessageCti:=MessageCti+' : corresp. occupé';
    raRefus   : MessageCti:=MessageCti+' : Appel refusé';
    raAbort   : MessageCti:=MessageCti+' : Appel interrompu';
    end;
SetControlText('INFOSCTI',MessageCti);
end;


procedure AGLRTCTIDecrocheAppel( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS) then TOM_TIERS(OM).RTCTIDecrocheAppel() else exit;
end;

procedure AGLRTCTINumeroter( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS) then TOM_TIERS(OM).RTCTINumeroter() else exit;
end;

procedure AGLRTCTIRaccrocheAppel( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS) then TOM_TIERS(OM).RTCTIRaccrocheAppel() else exit;
end;

procedure AGLRTCTIAppelAttente( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS) then TOM_TIERS(OM).RTCTIAppelAttente() else exit;
end;

procedure AGLRTCTIRepriseAppel( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS) then TOM_TIERS(OM).RTCTIRepriseAppel() else exit;
end;
{$ENDIF}

procedure TOM_TIERS.RTTransProCli ;
begin
ForceUpdate;
SetField('T_NATUREAUXI','CLI') ;
{$IFDEF EAGLCLIENT}
SetControlText('TT_NATURAUXI','Client') ;
{$ENDIF}
stNatureAuxi:= 'CLI';
SetField('T_DATEPROCLI', V_PGI.DateEntree);
OnLoadRecord;
TransProCli:=true;
SetControlEnabled('BDUPLICATION',FALSE);
SetControlEnabled('BRIB',FALSE);
SetControlEnabled('BMEMO',FALSE);
SetControlEnabled('BDP',FALSE);
SetControlEnabled('BCOURRIER',FALSE);
end;

procedure AGLRTListeCollectif( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS) then TOM_TIERS(OM).RTListeCollectif() else exit;
end;

procedure AGLRTListeEnseigne( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS) then TOM_TIERS(OM).RTListeEnseigne() else exit;
end;

procedure AGLBTPBordereaux( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
  if (OM is TOM_TIERS) then TOM_TIERS(OM).BTPBordereaux() else exit;
end;

procedure TOM_TIERS.RTListeCollectif ;
var ChaineAct : THEdit;
    Enseigne : THEdit;
    stWhere :string;
begin
  ChaineAct:= THEdit.Create (Nil);
  Enseigne := THEdit(GetControl('T_COLLECTIF'));
  ChaineAct.Parent:=Enseigne.Parent;
  ChaineAct.visible := False;
  ChaineAct.Top := Enseigne.top;
  ChaineAct.Left := Enseigne.left+Enseigne.width;
  stWhere := 'G_COLLECTIF = "X"';
{$IFNDEF EAGLCLIENT}
  ChaineAct.text :=  Getfield ('T_COLLECTIF') ;
{$ENDIF}
{$IFNDEF EAGLCLIENT}
  if (LookupList(ChaineAct,'Collectif Tiers','GENERAUX','DISTINCT G_GENERAL','',stWhere,'',True,0,'',tlLocate )) then
{$ELSE}
  if (LookupList(ChaineAct,'Collectif Tiers','GENERAUX','DISTINCT G_GENERAL','',stWhere,'',True,0,'',tlDefault )) then
{$ENDIF}
  begin
    DS.edit;
    setfield ('T_COLLECTIF', ChaineAct.Text);
  end;
  ChaineAct.Destroy ;
end;

procedure TOM_TIERS.RTListeEnseigne ;
var ChaineAct : THEdit;
    Enseigne : THEdit;
    stWhere :string;
begin
  ChaineAct:= THEdit.Create (Nil);
  Enseigne := THEdit(GetControl('T_ENSEIGNE'));
  ChaineAct.Parent:=Enseigne.Parent;
  ChaineAct.visible := False;
  ChaineAct.Top := Enseigne.top;
  ChaineAct.Left := Enseigne.left+Enseigne.width;
  stWhere := 'T_ENSEIGNE<>"" AND T_ENSEIGNE LIKE "'+GetControlText('T_ENSEIGNE')+'%"';
{$IFNDEF EAGLCLIENT}
  ChaineAct.text :=  Getfield ('T_ENSEIGNE') ;
{$ENDIF}
{$IFNDEF EAGLCLIENT}
  if (LookupList(ChaineAct,'Enseigne','TIERS','DISTINCT T_ENSEIGNE','',stWhere,'',True,0,'',tlLocate )) then
{$ELSE}
  if (LookupList(ChaineAct,'Enseigne','TIERS','DISTINCT T_ENSEIGNE','',stWhere,'',True,0,'',tlDefault )) then
{$ENDIF}
  begin
    DS.edit;
    setfield ('T_ENSEIGNE',ChaineAct.Text);
  end;
  ChaineAct.Destroy ;
end;

procedure TOM_TIERS.TE_LIVREOnEnter(Sender: TObject);
begin
   { Sauve le code tiers livré }
   TE_LivreEnter := GetControlText('TE_LIVRE');
end;

procedure TOM_TIERS.TE_LIVREOnExit(Sender: TObject);
begin
  {GPAO_V500_008 Début}
  // Normalement, cette fonction est appellée dans le script sur le OnChange.
  // Mais il y a une erreur dans le script, c'est fait sur YTC_TIERSLIVRE_OnChange
  // et non sur TE_LIVRE_OnChange => On n'y passe jamais. Comme la Socref est figée, je fait un appel ici.
  TestModifLivre;
  {GPAO_V500_008 Fin}
   { Si le code tiers livré à changé on remet le N° d'adresse de livraison à 0 }
   if TE_LivreEnter <> GetControlText('TE_LIVRE') then
   begin
      if GetControlText('TE_LIVRE') <> '' then
      begin 
        //SetControlText('YTC_NADRESSELIV', IntToStr(GetNumAdresseFromTiers('', GetControlText('TE_LIVRE'), taLivr)));
        SetControlText('YTC_NADRESSELIV', GetNumAdresseFromTiers('', GetControlText('TE_LIVRE'), taLivr));
        SetControlProperty('YTC_NADRESSELIV', 'ENABLED', True);  
      end
      else
      begin
        SetControlText('YTC_NADRESSELIV', '0');
        SetControlProperty('YTC_NADRESSELIV', 'ENABLED', False);  
      end;
   end;
   SetPlusNumAdresseLiv(Self);
end;

procedure TOM_TIERS.SetPlusNumAdresseLiv(Sender: TObject);
{ Mise à jour de la propriétée .Plus dans le contrôle du N° d'adresse de livraison }
begin
   SetControlProperty('YTC_NADRESSELIV', 'Plus', 'ADR_TYPEADRESSE="TIE"'
                                               + ' AND ADR_REFCODE="' + GetControlText('TE_LIVRE') + '"'
                                               + ' AND ADR_LIVR="X"'
                                               );
end;

procedure TOM_TIERS.TE_FACTUREOnEnter(Sender: TObject);
begin
   { Sauve le code tiers facturé }
   TE_FactureEnter := GetControlText('TE_FACTURE');
end;

procedure TOM_TIERS.TE_FactureOnExit(Sender: TObject);
var Qr : TQuery;
begin
//  BBI Fiche 11227
  if stNatureAuxi = 'FOU' then
  begin
    if GetControlText('TE_FACTURE') <> TE_FactureEnter then
      DS.Edit;
    if (GetControlText('TE_FACTURE') = '') then SetField('T_FACTURE',GetField('T_AUXILIAIRE'))
    else
    begin
      Qr := OpenSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+GetControlText('TE_FACTURE')+'"', False);
      try
        if not Qr.Eof then
          SetField('T_FACTURE', Qr.Findfield('T_AUXILIAIRE').AsString)
        else
        begin
          SetLastError(5, 'TE_FACTURE');
          exit ;
        end;
        THEdit(GetControl('T_FACTURE')).Text := Qr.Findfield('T_AUXILIAIRE').AsString;
      finally
        Ferme(Qr);
      end;
    end;
  end;
   { Si le code tiers facturé à changé on remet le N° d'adresse de livraison à 0 }
   if TE_FactureEnter <> GetControlText('TE_FACTURE') then
   begin
      if GetControlText('TE_FACTURE') <> '' then
      begin
        //SetControlText('YTC_NADRESSEFAC', IntToStr(GetNumAdresseFromTiers('', GetControlText('TE_FACTURE'), taFact)));
        SetControlText('YTC_NADRESSEFAC', GetNumAdresseFromTiers('', GetControlText('TE_FACTURE'), taFact));
        SetControlProperty('YTC_NADRESSEFAC', 'ENABLED', True);
      end
      else
      begin
        SetControlText('YTC_NADRESSEFAC', '0');
        SetControlProperty('YTC_NADRESSEFAC', 'ENABLED', False);
      end
   end;
   { Réinitialise le plus du where du N° d'adresse }
   SetPlusNumAdresseFac(Self);
end;

//  BBI Fiche 11227
procedure TOM_TIERS.TE_PAYEUROnEnter(Sender: TObject);
begin
   { Sauve le code tiers facturé }
   TE_PayeurEnter := GetControlText('TE_PAYEUR');
end;

procedure TOM_TIERS.TE_PAYEUROnExit(Sender: TObject);
var Qr : TQuery;
begin
  if stNatureAuxi = 'FOU' then
  begin
    if GetControlText('TE_PAYEUR') <> TE_PayeurEnter then
      DS.Edit;
    if GetControlText('TE_PAYEUR') = GetField('T_TIERS') then SetControlText('TE_PAYEUR', '');
    if (GetControlText('TE_PAYEUR') <> '') then
    begin
      If LookupvalueExist(THEdit(getControl('TE_PAYEUR'))) then  // On ne peut saisire qu'un code tiers qui peut être choisis
      begin
        Qr := OpenSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+GetControlText('TE_PAYEUR')+'"', False);
        try
          if not Qr.Eof then SetField('T_PAYEUR', Qr.Findfield('T_AUXILIAIRE').AsString);
          THEdit(GetControl('T_PAYEUR')).Text := Qr.Findfield('T_AUXILIAIRE').AsString;
        finally
          Ferme(Qr);
        end;
      end
      Else
      begin
        SetLastError(44,'TE_PAYEUR');
        Exit;
      end;
    end
    else
    begin     //mcd 05/03/02
       if (old_payeur <>'') then SetField('T_PAYEUR','');
    end;
  end;
end;
//  BBI Fin Fiche 11227

procedure TOM_TIERS.SetPlusNumAdresseFac(Sender: TObject);
{ Mise à jour de la propriétée .Plus dans le contrôle du N° d'adresse de facturation }
begin
   SetControlProperty('YTC_NADRESSEFAC', 'Plus', 'ADR_TYPEADRESSE="TIE"'
                                               + ' AND ADR_REFCODE="' + GetControlText('TE_FACTURE') + '"'
                                               + ' AND ADR_FACT="X"'
                                               );
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 31/03/2003
Modifié le ... :   /  /
Description .. : Appel de la fiche EDITiers pour paramètres de configuration
Suite ........ : EDI
Mots clefs ... :
*****************************************************************}
{$IFDEF EDI}
procedure TOM_TIERS.MnTiersEDI_OnClick(Sender: TObject);
begin
  EDICallMulETS(GetField('T_TIERS'), ActionToString(TFFiche(Ecran).FTypeAction), True);
end;
{$ENDIF}

procedure TOM_TIERS.T_Email_OnDblclick(Sender : tObject);
begin
  SendMail('',GetControlText('T_EMAIL'),'',nil,'',False);
end;

procedure TOM_TIERS.T_RVA_OnDblclick(Sender : tObject);
var
  sHttp: String;
begin
  sHttp := GetControlText('T_RVA');
  if Pos('HTTP://', UpperCase(sHttp)) = 0 then
    sHttp := 'http://'+sHttp;
  LanceWeb(sHttp,True);
end;

// JT - eQualité 10384
procedure TOM_TIERS.T_TIERSOnExit(Sender : tObject);
begin
  if OldTiersOnCreat = '' then
    OldTiersOnCreat := GetField('T_TIERS');
end;

//Gestion d'affichage de la liste des Appels
procedure TOM_TIERS.BtAppels_OnClick(Sender: TObject);
var StSql : string;
	 QQ : TQuery;
   AppelsExistants : boolean;
Begin
{$IFDEF CTI}
	AppelsExistants := false;
  StSQL := 'SELECT AFF_AFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE0="W" AND AFF_TIERS="' + GetField ('T_TIERS') + '"';
  QQ := OpenSQL(StSQL, true);
  If not QQ.Eof then AppelsExistants := true;
  Ferme(QQ);
  if AppelsExistants then
  begin
    if GetParamSoc('SO_BTAFFAIRERECHLIEU') then
    	AGLLanceFiche('BTP', 'BTMULAPPELSINT','AFF_AFFAIRE0=W','','AFF_TIERS=' + GetField('T_TIERS')+ ';NUMTEL='+AppelantCTI+ ';APPELANT='+stContact + ';ETAT=ECO;STATUT=APP')
    else
    	AglLanceFiche('BTP','BTMULAPPELS','AFF_AFFAIRE0=W','AFF_TIERS=' + GetField('T_TIERS'),'AFF_TIERS=' + GetField('T_TIERS') + ';NUMTEL='+AppelantCTI+ ';APPELANT='+stContact+';ETAT=ECO;STATUT=APP'); // Affaires
  end else
  begin
    //FV1 : 22/12/2015 - FS#1820 - VEODIS : Pb en saisie d'appel depuis fiche client
    if not(CreationAffaireAutorise) then
      Exit
    else
  	AGLLanceFiche('BTP', 'BTAPPELINT','','','CODETIERS=' + GetField('T_TIERS')+';NUMTEL='+AppelantCTI+ ';APPELANT='+stContact+';ACTION=CREATION')
  end;
	SetEtatBoutonCti;
{$ENDIF}
end;

//Gestion d'affichage de la liste des contrats
procedure TOM_TIERS.BtContrat_OnClick(Sender: TObject);
Var StChamps    : String;
    StArgument  : String;
begin



    StChamps := 'AFF_AFFAIRE0=I';
    StChamps := stChamps + ' ;AFF_STATUTAFFAIRE=INT';
    StChamps := stChamps + ' ;AFF_AVENANT=';

    StArgument := 'AFF_TIERS=' + GetField('T_TIERS');
    StArgument := StArgument + ';ETAT=ENC';
    StArgument := Stargument + ';STATUT=INT';

    AGLLanceFiche('BTP', 'BTCONTRAT_MUL', StChamps, '', StArgument);

    //AglLanceFiche('BTP','BTAFFAIRE_MUL','AFF_AFFAIRE0=I','',STATUT=INT');   // Affaires
end;

function Tiers_MyAfterImport (Sender: TObject) : string;
var  OM : TOM ;
begin
result := '';
if sender is TFFICHE then OM := TFFICHE(Sender).OM else exit;
if (OM is TOM_TIERS) then result := TOM_TIERS(OM).GetInfoTiers else exit;
end;

// suppression des bouton magneto alors que monofiche = false
procedure Tom_TIERS.AfterFormShow;
begin
//

  TToolbarButton97(GetControl('BLast')).visible := false;
  TToolbarButton97(GetControl('BFirst')).visible := false;
  TToolbarButton97(GetControl('BNext')).visible := false;
  TToolbarButton97(GetControl('BPrev')).visible := false;

end;

function TOM_TIERS.GetInfoTiers : string;
begin
if (ds<>nil) and not(DS.State in [dsInsert]) and ((GetField ('T_NATUREAUXI')= 'CLI') or (GetField ('T_NATUREAUXI')= 'PRO')) then result := 'Tiers='+GetField('T_TIERS') ;
end;

procedure Tiers_GestionBoutonGED (Sender: TObject);
var  OM : TOM ;
begin
if sender is TFFICHE then OM := TFFICHE(Sender).OM else exit;
if (OM is TOM_TIERS) then TOM_TIERS(OM).GestionBoutonGED else exit;
end;

procedure TOM_TIERS.BTPBordereaux;
begin
	GestionDetailEtude (GetField ('T_NATUREAUXI'),GetField('T_TIERS'),'',false,TFFiche(Ecran).fTypeAction ,TOBEtude,TatBordereaux);   // Bordereaux de prix
end;

//uniquement en line
{*
procedure TOM_TIERS.CacheZonesNONLine;
begin
//
  SetcontrolVisible ('TSCOMPLEMENT', False);
  SetcontrolVisible ('PREGLEMENT', False);
  SetcontrolVisible ('PINFORMATION', False);
//
  if stNatureAuxi = 'FOU' then SetcontrolVisible ('PCONDITION', False);
//
	SetControlVisible ('T_TARIFTIERS',false);
	SetControlVisible ('TT_TARIFTIERS',false);
	SetControlVisible ('GBREPRESENTANT',false);
	SetControlVisible ('T_PRESCRIPTEUR',false);
	SetControlVisible ('TT_PRESCRIPTEUR_',false);
	SetControlVisible ('T_APPORTEUR',false);
	SetControlVisible ('TT_APPORTEUR',false);
	SetControlVisible ('T_COEFCOMMA',false);
	SetControlVisible ('TT_COEFCOMMA',false);
	SetControlVisible ('TE_PAYEUR',false);
	SetControlVisible ('TT_PAYEUR',false);
	SetControlVisible ('GBCONTACT',false);
	SetControlVisible ('GBPUBLIPOSTAGE',false);
	SetControlVisible ('TYTC_NADRESSELIV',false);
	SetControlVisible ('YTC_NADRESSELIV',false);
	SetControlVisible ('TYTC_NADRESSEFAC',false);
	SetControlVisible ('YTC_NADRESSEFAC',false);
	SetControlVisible ('BPOPMENU',false);
	SetControlVisible ('BPOP2MENU',false);
	SetControlVisible ('BCONTACT',false);
	SetControlVisible ('BCOURRIER',false);
	SetControlVisible ('PINFORMATION',false);
//
	SetControlEnabled('T_FACTUREHT',false);
	SetControlVisible('T_REGION',false);
	SetControlVisible('T_PAYS',false);
	SetControlVisible('TT_PAYS',false);
	SetControlVisible('T_ENSEIGNE',false);
	SetControlVisible('TT_ENSEIGNE',false);
	SetControlVisible('T_SOCIETEGROUPE',false);
	SetControlVisible('TT_SOCIETEGROUPE',false);
	SetControlVisible('LBGROUPE',false);
	SetControlVisible('T_CONFIDENTIEL',false);
	SetControlVisible('TT_CONFIDENTIEL',false);
	SetControlVisible('T_LANGUE',false);
	SetControlVisible('TT_LANGUE',false);
	SetControlVisible('GB_TARIFICATION',false);
	SetControlVisible('GB_GESTION',false);
	SetControlVisible('GBEXPORT',false);
	SetControlVisible('GBLOGISTIQUE',false);
	SetControlVisible('GBRELANCE',false);
	SetControlVisible('T_PASSWINTERNET',false);
	SetControlVisible('TT_PASSWINTERNET',false);
//
	if GetParamSocSecur('SO_GCDESACTIVECOMPTA',true) then
		 begin
	   SetControlEnabled('GB_COMPTA',false);
		 end;

	SetControlVisible('GB_CREDIT',false);
	SetControlVisible('BTAPPELS',false);

end;
*}

procedure TOM_TIERS.DevisClick(Sender: Tobject);
begin
	AGLLanceFiche('BTP','BTDEVIS_MUL','GP_NATUREPIECEG=DBT;GP_VENTEACHAT=VEN;GP_TIERS='+GetField('T_TIERS'),'','MODIFICATION') ;
end;

procedure TOM_TIERS.BITI_OnClick(Sender: TObject);
begin
  LanceGoogleMaps (GetControlText('T_ADRESSE1'),GetControlText('T_VILLE'),false);
end;

procedure TOM_TIERS.BPLAN_OnClick(Sender: TObject);
begin
  LanceGoogleMaps (GetControlText('T_ADRESSE1'),GetControlText('T_VILLE'),true);
end;

{$IFDEF GRC}
 {$IFNDEF GIGI} // $$$ JP 09/08/07: pas de CTI en GI (PCL), car fait dans le Bureau exclusivement


procedure TOM_TIERS.BTELBUR_OnClick(Sender: TObject);
begin
{$IFNDEF BUREAU}
  FonctionCTI ('MAKECALL', GetControlText ('T_FAX'), 'T_AUXILIAIRE=' + GetControlText ('T_AUXILIAIRE')); // $$$ JP 13/08/07 FQ 10691 GetControlText ('T_TELEX'));
{$ENDIF BUREAU}
end;

procedure TOM_TIERS.BTELDOM_OnClick(Sender: TObject);
begin
{$IFNDEF BUREAU}
  FonctionCTI ('MAKECALL', GetControlText ('T_TELEPHONE'), 'T_AUXILIAIRE=' + GetControlText ('T_AUXILIAIRE'));
{$ENDIF BUREAU}
end;

procedure TOM_TIERS.BTELPORT_OnClick(Sender: TObject);
begin
{$IFNDEF BUREAU}
  FonctionCTI ('MAKECALL', GetControlText ('T_TELEX'), 'T_AUXILIAIRE=' + GetControlText ('T_AUXILIAIRE')); // $$$ JP 13/08/07 FQ 10691 GetControlText ('T_GetControlText ('T_FAX'));
{$ENDIF BUREAU}
end;

procedure TOM_TIERS.BPhoneClick(Sender: TObject);
begin
  SetControlProperty('INFOSCTI','Caption','Communication en cours');
  FonctionCTI ('GETCALL','');
  BPHONE.Down := True;
  BPHONE.enabled := false;
  BAttente.Visible:= True;
  BStop.Visible:= True;
end;

procedure TOM_TIERS.BStopClick (Sender : TObject);
begin
  SetControlProperty('INFOSCTI','Caption','Communication terminée');
  FonctionCTI ('CALLBYE','');
  BPHONE.Down := false;
  BPHONE.enabled := true;
  BAttente.Visible:= false;
  BStop.Visible:= false;
	SetEtatBoutonCti;
end;

procedure TOM_TIERS.BAttenteClick (Sender : TObject);
begin
  if BATTENTE.hint = 'Mise en attente' then
  begin
  SetControlProperty('INFOSCTI','Caption','Mise en attente');
  FonctionCTI ('CALLWAIT','');
  BATTENTE.hint := 'Reprise de l''appel';
  end else
  begin
  SetControlProperty('INFOSCTI','Caption','Communication en cours');
  FonctionCTI ('CALLWAIT','');
  BATTENTE.hint := 'Mise en attente';
  end;
end;
procedure TOM_TIERS.SetEtatBoutonCti;
begin
	if VH_RT.CtiAlerte <> nil then
  begin
  BPhone.Visible := (VH_RT.ctiAlerte.GetState<>TCtiDisable)
  									and (VH_RT.ctiAlerte.GetState<>TctiDecroche)
                    and (VH_RT.ctiAlerte.GetState<>TctiNothingToDo)   ;
  BAttente.Visible:= (VH_RT.ctiAlerte.GetState<>TCtiDisable)
  										and (VH_RT.ctiAlerte.GetState=TctiDecroche);
  BStop.Visible:= (VH_RT.ctiAlerte.GetState<>TCtiDisable)
  								and (VH_RT.ctiAlerte.GetState=TctiDecroche) ;
  end;
end;
{$ENDIF}
{$ENDIF}

procedure TOM_TIERS.BVisuParcClick(Sender: TObject);
var CodeTiers : string;
		ActionT : string;
begin
	if TFFiche(Ecran).fTypeAction<>taConsult then ActionT := 'ACTION=MODIFICATION'
  																				 else ActionT := 'ACTION=CONSULTATION';
	CodeTiers := getField('T_TIERS');
  if CodeTiers  <> '' then
     AGLLanceFiche ('BTP','BTPARCTIER','',CodeTiers,ActionT+';TIERS='+CodeTiers+';MONOFICHE');

end;

procedure AGLRTCTIETATBOUTON( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_TIERS)
    then
{$IFDEF GRC}
 {$IFNDEF GIGI} // $$$ JP 09/08/07: pas de CTI en GI (PCL), car fait dans le Bureau exclusivement
    TOM_TIERS(OM).SetEtatBoutonCti
{$ENDIF}
{$ENDIF}
    else exit;
end;

(*
procedure TOM_TIERS.GoSeriaLSE(Sender: TObject);
type TAppSeria = procedure (Server,DataBase ,CodeClient : PChar) ;  stdcall;
var  AppSeria : TAppSeria;
begin
	@AppSeria := GetProcAddress(HandleSeria,'SetVerrouClient');
  if Assigned (AppSeria) then PGIInfo('OKOK');
end;
*)

//FV1 : 09/03/2015

procedure TOM_TIERS.BCODEBARREOnClick(Sender : Tobject);
var trouve    : Boolean;
    Codetiers : string;
    Nature    : String;
    Action    : string;
begin

  Codetiers := Getfield('T_TIERS');
  //
  if not ExisteSQL('SELECT BCB_IDENTIFCAB FROM BTCODEBARRE WHERE BCB_NATURECAB="' + stNatureAuxi + '" and BCB_IDENTIFCAB="'+Codetiers+'"') then
    Action := 'ACTION=CREATION'
  else
    Action := 'ACTION=MODIFICATION';

  AGLLanceFiche('BTP','BTCODEBARRE','','', Action + ';ID=' + stNatureAuxi + ';CODE='+ CodeTiers );

  ChargeCodeBarre;

end;

Procedure TOM_TIERS.ChargeCodeBarre;
var StSQL  : string;
    QQ     : TQuery;
begin

  //Au retour on vérifie si un code principal et on l'affiche
  StSQL := 'SELECT * FROM BTCODEBARRE ';
  StSQL := StSQL + ' WHERE BCB_NATURECAB = "' + stNatureAuxi + '" ';
  StSQL := StSQL + ' AND   BCB_IDENTIFCAB= "' +  Getfield('T_TIERS') + '" ';
  StSQL := StSQL + ' AND   BCB_CABPRINCIPAL= "X" ';

  QQ:=OpenSql(StSQL, True) ;

  if not QQ.Eof then
  begin
    ThEdit(Getcontrol('BT1_CODEBARRE')).Text :=  QQ.FindField('BCB_CODEBARRE').AsString;
    SetControlText ('BT1_QUALIFCODEB', QQ.FindField('BCB_QUALIFCODEBARRE').AsString);
    if not (DS.State in [dsInsert, dsEdit]) then DS.edit;
 end;

  Ferme(QQ) ;

end;

Function TOM_TIERS.GestionFournisseur : boolean;
Var QR          : TQuery;
    LePays      : string;
    TypeCB      : string;
    STCodeBarre : string;
begin
  Result := false;
  SetField('T_FACTUREHT','X');       //JCF

  // -- Ajout LS Pour gestion du tiers de facturation et de paiment
  if (GetControlText ('TE_FACTURE')= '') then
    SetField('T_FACTURE',GetField('T_AUXILIAIRE'))
  else
  begin
    Qr := OpenSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+GetControlText('TE_FACTURE')+'"', False);
    try
      if not Qr.Eof then
        SetField('T_FACTURE', Qr.Findfield('T_AUXILIAIRE').AsString) else
      begin
        SetLastError(5, 'TE_FACTURE');
        Exit;
      end;
    finally
      Ferme(Qr);
    end;
  end;

  if (GetControlText('TE_LIVRE') = '') then
    TobZoneLibre.PutValue('YTC_TIERSLIVRE', '')
  else
  begin
    Qr := OpenSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="' + GetControlText('TE_LIVRE') + '"', False);
    try
      if not Qr.Eof then
        TobZoneLibre.PutValue('YTC_TIERSLIVRE', Qr.Findfield('T_AUXILIAIRE').AsString)
      else
      begin
        SetLastError(37, 'TE_LIVRE');
        Exit;
      end;
    finally
      Ferme(Qr);
    end;
  end;

  if GetControlText ('TE_PAYEUR')=GetField ('T_TIERS') then SetControlText('TE_PAYEUR', '');

  if (GetControlText ('TE_PAYEUR')<>'') then
  begin
    If LookupvalueExist(thedit(getControl('TE_PAYEUR'))) then  // On ne peut saisir qu'un code tiers qui peut être choisis
    begin
      Qr := OpenSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+GetControlText('TE_PAYEUR')+'"', False);
      try
        if not Qr.Eof then SetField('T_PAYEUR', Qr.Findfield('T_AUXILIAIRE').AsString);
      finally
        Ferme(Qr);
      end;
    end
    Else
    begin
      SetLastError(44,'TE_PAYEUR');
      Exit;
    end;
  end
  else if (old_payeur <>'') then SetField('T_PAYEUR','');

  // JTR - TVA Intracommunautaire
  if (VH_GC.GereTVAIntraComm) and (GetField('T_PAYS') <> '') then
  begin
    LePays := GetField('T_PAYS');
    Qr := OpenSql('SELECT PY_MEMBRECEE FROM PAYS WHERE PY_PAYS = "' + LePays + '"', true);
    try
      if (Qr.FindField('PY_MEMBRECEE').AsString = 'X') and
         (LePays <> GetParamSoc('SO_GCTIERSPAYS'))     and
         (GetField('T_REGIMETVA') <> VH_GC.TypeTVAIntraComm) then
      begin
        if PGIAsk('Le régime fiscal ne correspond pas au pays. Voulez-vous continuer ?', 'TVA intracommunautaire') <> mrYes then
        begin
          LastError := 1;
          LastErrorMsg := '';
          SetFocusControl ('T_REGIMETVA');
          exit;
        end;
      end;
    finally
      Ferme(Qr);
    end;
  end;
  // Fin JTR
  Result := true;

end;


procedure TOM_TIERS.ParamInterfaceRexelClick(Sender: TObject);
begin
  AGLLanceFiche('BTP','BTPARECHREXEL','','','AUXILIAIRE='+GetControltext('T_AUXILIAIRE'));
end;

Initialization
//=== TOM =================
registerclasses([TOM_TIERS]) ;
registerclasses([TOM_TIERSPIECE]) ;
RegisterAglProc( 'AfficheContactTiers', TRUE , 0, AGLAfficheContactTiers);
RegisterAglProc( 'ModifParticulier', TRUE , 0, AGLModifParticulier);   // Mode
RegisterAglProc( 'Majlescontacts', TRUE , 0, AGLmajcontact); // Mode
RegisterAglProc( 'ZonesLibresIsModified', TRUE , 0, AGLZonesLibresIsModified);
RegisterAglProc( 'TestModifZoneLibre', TRUE , 1, AGLTestModifZoneLibre);
RegisterAglProc( 'TestModifZoneLibre2', TRUE , 1, AGLTestModifZoneLibre2);
RegisterAglProc( 'TestModifLivre', TRUE , 0, AGLTestModifLivre);
RegisterAglProc( 'TestModifFacture', TRUE , 0, AGLTestModifFacture);
RegisterAglProc( 'TestModifPayeur', TRUE , 0, AGLTestModifPayeur);
RegisterAglProc( 'MajIndisponible', TRUE , 0, AGLMajIndisponible);
RegisterAglProc( 'DuplicationTiers', TRUE , 0, AGLDuplication);
RegisterAglFunc( 'VORRisqueTiers', TRUE , 0, AGLVORRisqueTiers);
RegisterAglProc( 'GCVoirEncours', TRUE , 0, AGLGCVoirEncours);
RegisterAglFunc( 'PCSPHONEME', False, 1, AGLPCSPHONEME);
RegisterAglProc( 'RTTiersImprime', TRUE , 0, AGLRTTiersImprime);
RegisterAglProc( 'TOM_Tiers_AppelIsoFlex',TRUE,1,TOM_Tiers_AppelIsoFlex);
RegisterAglProc( 'MajAffichageKardexEntreprise', TRUE , 0, AglMajAffichageKardexEntreprise);   // CHR
RegisterAglProc( 'RecuplesContacts', TRUE , 1, AglRecupContact);   // CHR
RegisterAglProc( 'MajContacts_chr', TRUE , 0, AGLMajContact_Chr); // CHR
RegisterAglFunc( 'RechercheClient', TRUE , 1, AGLRechercheClient);  //CHR
RegisterAglProc( 'RTTransProCli', TRUE , 0, AGLRTTransProCli);
RegisterAglProc( 'RTListeEnseigne', TRUE , 0, AGLRTListeEnseigne);
RegisterAglProc( 'RTListeCollectif', TRUE , 0, AGLRTListeCollectif);
RegisterAglProc( 'BTPBordereaux', TRUE , 0, AGLBTPBordereaux);
{$IFDEF CTI}
RegisterAglProc( 'RTCTIDecrocheAppel', TRUE , 0, AGLRTCTIDecrocheAppel);
RegisterAglProc( 'RTCTIRaccrocheAppel', TRUE , 0, AGLRTCTIRaccrocheAppel);
RegisterAglProc( 'RTCTINumeroter', TRUE , 0, AGLRTCTINumeroter);
RegisterAglProc( 'RTCTIAppelAttente', TRUE , 0, AGLRTCTIAppelAttente);
RegisterAglProc( 'RTCTIRepriseAppel', TRUE , 0, AGLRTCTIRepriseAppel);
RegisterAglProc( 'RTCTIETATBOUTON', TRUE , 0, AGLRTCTIETATBOUTON);
{$ENDIF}
end.
