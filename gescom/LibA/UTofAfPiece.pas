unit UTofAfPiece;

interface

uses
  {$IFDEF EAGLCLIENT}
  Maineagl,efiche,
  {$ELSE}
  MajTable, Fe_Main, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fiche,
  {$ENDIF}
  {$IFNDEF CCS3}
      UtilConfid,
  {$ENDIF}
  {$IFDEF BTP}
  BTPUtil,TiersUtil,CalcOleGenericBTP,AglInitGc,FactTvaMilliem,
  {$ENDIF}
  M3FP, StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
  HCtrls, HEnt1, HMsgBox, UTOB, UTOF, AglInit, Agent, EntGC,
  AffaireUtil,  FactAdresse, factcomm,
  DicoBTP, HTB97, FactUtil, UAFO_Ressource, Vierge, UtilPGI, LookUp,
  UTOFAFBASECODEAFFAIRE, UTOFAFTRADUCCHAMPLIBRE, utilgc, Graphics, paramsoc,uEntCommun;

type
  TOF_AFPiece = class(TOF_AFTRADUCCHAMPLIBRE)
  private
  	ModifAvanc : boolean;
  	TypeCom : string;
    Apporteur     : THEdit;
    BAdrRefresh   : TToolbarButton97;
    TOBAdresses   : TOB;
    TobAdr_Old    : TOB;
    EnteteAffaire : Boolean;
    EnteteAffPiece: Boolean;
    Particulier   : Boolean;
    Cp_Adr_Ident  : Boolean;
    RibEnc        : string;
    toto          : string;
    {$IFDEF BTP}
    TOBRepart : TOB;
    GPTiersLivre  : THEdit;
    GPTiersFac    : THedit;
    NADRESSELIV : string;
    AccesModeFac  : Boolean;
    DateSituation : String;
    TiersFacture  : String;
    TiersLivre    : string;
    Action        : TActionFiche;
    //
    Commercial    : THCritMaskEdit;
    LblCommercial : THLabel;
    OldCommercial : String;
    //
    OkPieceAdresse: Boolean;
    //
    procedure AffichageMarge;
    procedure SetAdrLivraisonModifie;
		procedure SelectionAdresseLivAffaire (Sender : Tobject);
    procedure SelectionAdresseLivTiers (Sender : Tobject);
    procedure PutEcranTobPieceAdresseLivraison(TobAdresseLivraison: tob; ForceAdresse : boolean=false);
		function NombreAdressesLiv (Affaire : string) : integer;
    procedure SetInfosContact(TypeAdr : string; NumeroContact: integer; Telephone: string);
    procedure GP_TIERSLIVRE_OnElipsisClick(Sender: TObject);
    procedure GP_TIERSLIVRE_OnEnter(Sender: TObject);
    procedure GP_TIERSLIVRE_OnExit(Sender: TObject);
    procedure GP_TIERSFAC_OnElipsisClick(Sender: TObject);
    procedure RechercheAdresseLivraison;
    procedure GP_NADRESSELIV_OnEnter(Sender: TObject);
    procedure GP_NADRESSELIV_OnExit(Sender: TObject);
    Function ExisteAdresseLiv ( sTiersLivre : string) : boolean;
    Procedure NumeroContactLivClick (Sender : Tobject);
    Procedure NumeroContactClick (Sender : Tobject);
    procedure SetComponents;
    procedure ChangeAdresseLivr (Sender : TObject);
    procedure GP_TIERSFACExit (Sender : Tobject);
    procedure RechercheAdressefact;
    procedure PutEcranTobPieceAdressefac(TobAdresseFact: tob;ForceAdresse: boolean=false);
    {$ENDIF}
    procedure ApporteurExit(Sender: Tobject);
    function GetCodeAuxi(Prefixe: string): string;
//uniquement en line
//    procedure SetShowForLine;
    procedure SetModeRegle;
    procedure ControleChamp(Champ, Valeur: Variant);
    procedure ChargementAdresseTobEcran(TypeAdresse, Prefixe1, Prefixe2: string; TOBLAdresses: TOB);
    procedure ChargementAdresseEcranTOB(TypeAdresse, Prefixe1, Prefixe2: string; TOBLAdresses: TOB);
    //
    procedure GP_REPRESENTANTElipsisClick(Sender: TObject);
    procedure GP_REPRESENTANTExit(Sender: TObject);

  public
    procedure OnArgument(stArgument: string); override;
    procedure OnCancel; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnClose; override;
    procedure SelectAdresseTiers;
    function IsAdresseModifie: boolean;
    procedure TrtAdresseSpecifAffaire;
  end;

type
  TOF_AFLignePiece = class(TOF_AFBASECODEAFFAIRE)
  private
  	TypeCom : string;
    LaRessource, CodeTiers, AffGenerauto: string;
    PieceAffaire, bVente: Boolean;
    Action: TActionFiche;
    procedure AffecteRessourceLigne;
    procedure RechFormuleVar(Sender: TObject); //Affaire-ONYX
  public
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;
    procedure OnCancel; override;
    procedure OnLoad; override;
    procedure OnArgument(stArgument: string); override;
    procedure OnUpdate; override;
    procedure OnClose; override;
    procedure FormuleOnChange(SEnder: TObject); //Affaire-ONYX
  end;

const
  TexteMsgCompl: array[1..9] of string = (
    {1}'Vous devez renseigner un code tiers livré valide',
    {2}'Vous devez renseigner un code tiers facturé valide',
    {3}'Vous devez renseigner un code commercial valide',
    {4}'Vous devez renseigner une ressource valide',
    {5}'Vous devez renseigner un apporteur valide',
    {6}'Changement de tiers interdit, l''affaire saisie n''est pas sur le tiers initial',
    {7}'Code affaire non valide',
    {8}'Vous devez renseigner un code tiers payeur valide',
    {9}'Vous devez renseigner unu ressource libre valide'
    );

procedure AFLanceFiche_ComplementPiece(Argument: string);
procedure AFLanceFiche_ComplementLigne(Argument: string);
procedure AFLanceFiche_Mul_Piece(range, Argument: string);
procedure AFLanceFiche_Mul_Edition_Piece(Argument: string);
procedure AFLanceFiche_Mul_Piece_Provisoire(Argument: string);
function AFLanceFiche_Date_ValidPiecePro(Argument: string): variant;
procedure AFLanceFiche_PIecePourUneAffaire(Argument: string);
procedure AFLanceFiche_DuplicPiece(Range, Argument: string);

implementation
/////////////////////////////////////////////////////////////////////////////
uses HrichOle;
procedure TOF_AFPIECE.ApporteurExit(Sender: Tobject);
begin
  if (Apporteur.text <> '') then
    if not (LookupValueExist(Apporteur)) then
    begin
      PGIBoxAF(TexteMsgCompl[5], Ecran.Caption);
      SetFocusControl('GP_APPORTEUR');
      TForm(Ecran).ModalResult := 0;
      Exit;
    end;
end;

procedure TOF_AFPiece.OnArgument(stArgument: string);
Var Critere     : String;
    Stnat       : string;
    i           : integer;
    bVisible    : Boolean;
    Combo       : THValComboBox;
  TOBAdresse : TOB;
  Cledoc : R_CLEDOC;
    //
  	Champ       : String;
    Valeur      : String;
    X           : integer;
begin
  inherited;
  AccesModeFac := true;

  AppliqueFontDefaut (THRichEditOle(GetControl('GP_BLOCNOTE')));

  ModifAvanc := false;
  // mcd 10/12/01 !! init fait en début et non dans la boucle ... posait pb
  EnteteAffaire := False; // gm 05/09/02 enteteaffaire si on arrive depuis les lignes de l'affaire
  EnteteAffPiece := False; // enteteaffaffaire si on arrive depuis l'affaire   directement

  Particulier := False;

  OkPieceAdresse  := GetParamSocSecur('SO_GCPIECEADRESSE', False);

  Critere := uppercase(Trim(ReadTokenSt(StArgument)));
  while Critere <> '' do
  begin
    x := pos('=', Critere);
    if x <> 0 then
    begin
      Champ  := copy(Critere, 1, x - 1);
      Valeur := copy(Critere, x + 1, length(Critere));
      end
      else
      Champ  := Critere;
    ControleChamp(Champ, Valeur);
    Critere:= uppercase(Trim(ReadTokenSt(StArgument)));
  end;

  // gm 10/01/03 plus de tiers payeur personnalisable en GA
  // je le laisse pour ALGOE car utilisé dans les editions d'adresse (cf NLB) 24/01/03
  if GetParamSoc('SO_AFCLIENT') <> cInClientAlgoe then
  begin
    SetControlVisible('TIERSPAYEUR', False);
    SetControlVisible('TGP_TIERSPAYEUR', False);
  end;

  // Modif BTP
  Apporteur := THEdit(Ecran.FindComponent('GP_APPORTEUR'));
  Apporteur.OnExit := ApporteurExit;

  Commercial                := THCritMaskEdit(Ecran.FindComponent('GP_REPRESENTANT'));
  Commercial.OnExit         := GP_REPRESENTANTExit;
  Commercial.OnElipsisClick := GP_REPRESENTANTElipsisClick;
  LblCommercial             := THLabel(Ecran.FindComponent('LIBREPRESENTANT'));

  {$IFDEF BTP}
	GPTiersLivre := THEdit(GetControl('GP_TIERSLIVRE'));
  GPTiersLivre.onEnter := GP_TIERSLIVRE_OnEnter;
  GPTiersLivre.onExit := GP_TIERSLIVRE_OnExit;
  GPTiersLivre.onElipsisClick := GP_TIERSLIVRE_OnElipsisClick;

  THEdit(GetCOntrol('GP_NADRESSELIV')).onElipsisClick := SelectionAdresseLivTiers;
	THEdit(GetCOntrol('GP_NADRESSELIV')).onEnter := GP_NADRESSELIV_OnEnter;
	THEdit(GetCOntrol('GP_NADRESSELIV')).onExit := GP_NADRESSELIV_OnExit;

	THEdit(GetCOntrol('ADL_NUMEROCONTACT')).OnElipsisClick  := NumeroContactLivClick;
  if THEdit(GetCOntrol('ADR_NUMEROCONTACT')) <> nil then
		THEdit(GetCOntrol('ADR_NUMEROCONTACT')).OnElipsisClick  := NumeroContactClick;
  //
  GPTiersFac := THEdit(GetControl('GP_TIERSFACTURE'));
  GPTiersFac.onExit := GP_TIERSFACExit;
  GPTiersFac.OnElipsisClick := GP_TIERSFAC_OnElipsisClick;
  //
  if ModifAvanc then
  begin
     TcheckBox(GetCOntrol('GP_APPLICFGST')).enabled := false;
     TcheckBox(GetCOntrol('GP_APPLICFCST')).enabled := false;
  end;
{$ENDIF}
  // --

  if EnteteAffaire or EnteteAffPiece then // certains compléments d'entête de pièce sont caché si appel depuis l'affaire
  begin
    //on ne rends pas dispo ni GMARGE ni GENTETE
    //FV1 : 12/12/2014 - FS#1350 - ACTUACOM - Non reprise de l'adresse de facturation en génération d'échéance
    SetControlProperty('GENTETE','Enabled', False);
    SetControlProperty('GMARGE' ,'Enabled', False);
    //SetControlVisible('GENERAL', False);
    SetControlVisible('COMMENTAIRE', False);
    SetControlVisible('GBPIECE', False);
    SetControlVisible('GBAFFAIRE', False);
    // mcd 18/10/01 SetControlVisible ('GBRESSOURCE',False);
    SetControlvisible('BADRESSETIERS', False);
    SetControlText('GP_REFINTERNE1', LaTob.getvalue('GP_REFINTERNE'));
    for i := 1 to 10 do
    begin
      if i <> 10 then
        Combo := THValComboBox(GetControl('GP_LIBRETIERS' + intTostr(i)))
      else
        Combo := THValComboBox(GetControl('GP_LIBRETIERSA')); // gm 05/09/02

      if Combo <> nil then bVisible := Combo.Visible else bVisible := False;
      SetControlVisible('ALIGNETIERS' + intTostr(i), bVisible);
    end;
    if VH_GC.GCIfDefCEGID then
       SetControlEnabled('GP_TARIFTIERS', False);
  end
  else // Pas entête de pièce de l'affaire
  begin
    SetControlVisible('TALIGNETIERS', False);
    SetControlVisible('TALIGNETIERS1', False);
    SetControlVisible('GP_REFINTERNE1', false);
    SetControlVisible('TGP_REFINTERNE1', false);
    for i := 1 to 10 do
    begin
      SetControlVisible('ALIGNETIERS' + intTostr(i), false);
    end;
    SetControlVisible('ALIGNEREPRESENTANT', false);
  end;

  if EnteteAffaire then
  begin
    Ecran.caption := 'Conditions de règlement';
    SetControlEnabled('GP_REGIMETAXE', False);
  end;

  if Particulier then
  begin
    SetControlProperty('ADR_JURIDIQUE', 'DataType', 'TTCIVILITE');
  end;
  if ctxscot in V_PGI.Pgicontexte then
  begin
    SetControlVisible('GP_REPRESENTANT', false);
    SetControlVisible('TGP_REPRESENTANT', false);
    SetControlVisible('ALIGNEREPRESENTANT', false);
  end;

  {$IFDEF EAGLCLIENT} //GISE
  SetControlEnabled('GP_ESCOMPTE', False);
  SetControlEnabled('GP_REMISEPIED', False);
  {$ENDIF}
  // 14/12/01 mcd ces zones ne sont pas géres .. on cache
  SetControlVisible('GP_QUALIFESCOMPTE', False);
  SetControlVisible('TGP_QUALIFESCOMPTE', False);
  {$IFDEF BTP}
  SetControlProperty('MODEFACTURATION', 'Plus', 'BTP');
//  SetControlVisible('INFORMATIONS', False);
  SetControlVisible('GBRESSOURCE', False);
  if (Ecran.Name = 'BTCOMPLPIECE') or (Ecran.Name = 'BTCOMPLPIECE_S1') then
  begin
    Stnat := GetInfoParPiece(LaTOB.GetValue('GP_NATUREPIECEG'),'GPP_VENTEACHAT');
  	Thedit(GetControl('LAFFAIRE')).text := '';
    if (Stnat = 'ACH') then
    begin
      TTabSheet(Ecran.FindComponent('LIVRAISON')).Caption := 'Adresse Fournisseur';
//      SetControlVisible('_GERE_EN_STOCK', false);
      SetControlVisible('LIVRECEPFOUR', false);
      SetControlVisible('GP_DATELIVRAISON', true);
      SetControlVisible('TGP_DATELIVRAISON', true);
      TTabSheet(Ecran.FindComponent('OPTIONS')).TabVisible  := false;
      SetControlText('GBADRESSE','Coordonnées du fournisseur');
    end else
    begin
    	if LaTob.GetValue('GP_AFFAIRE') <> '' then THEdit(GetControl('LAFFAIRE')).text := BTPCodeAffaireAffiche(LaTob.GetValue('GP_AFFAIRE'));
//      TTabSheet(Ecran.FindComponent('OPTIONS')).Visible := true;

    	if NombreAdressesLiv (LaTOB.getValue('GP_AFFAIRE')) > 1 then
      begin
      	SetControlVisible ('ADRESSELIVAFF',True);
      	SetControlVisible ('LNUMLIVRAFF',True);
        THedit (GetControl('ADRESSELIVAFF')).OnElipsisClick := SelectionAdresseLivAffaire;
      end;
    end;


//uniquement en line
//    SetControlText('GBTAXE', 'Regime de taxe du document dépendant du détail');
    SetControlText('LGB_FAMILLETAXE', 'Dépendant du détail');

    if LaTOB.GetValue('GP_AFFAIREDEVIS') <> '' then
       begin
       SetControlVisible('GACOMPTE', true);
       end;
    //
    if STNat = 'ACH' then
    begin
      THEdit(GetCOntrol('ADR_JURIDIQUE')).OnExit := ChangeAdresseLivr;
      THEdit(GetCOntrol('ADR_LIBELLE')).OnExit := ChangeAdresseLivr;
    	THEdit(GetCOntrol('ADR_LIBELLE2')).OnExit := ChangeAdresseLivr;
      THEdit(GetCOntrol('ADR_ADRESSE1')).OnExit := ChangeAdresseLivr;
      THEdit(GetCOntrol('ADR_ADRESSE2')).OnExit := ChangeAdresseLivr;
      THEdit(GetCOntrol('ADR_ADRESSE3')).OnExit := ChangeAdresseLivr;
      THEdit(GetCOntrol('ADR_CODEPOSTAL')).OnExit := ChangeAdresseLivr;
      THEdit(GetCOntrol('ADR_VILLE')).OnExit := ChangeAdresseLivr;
      THEdit(GetCOntrol('ADR_PAYS')).OnExit := ChangeAdresseLivr;
      TGroupBox(GetCOntrol('GBLIVRE')).Enabled := false;
    end;
    //
  end;
  {$ENDIF}
  if OkPieceAdresse then //mcd 14/02/03 pas de téléphone dans pieceAdresse
  begin
    SetControlVisible('ADR_TELEPHONE', False);
    SetControlVisible('TADF_TELEPHONE', False);
  end;
  {$IFNDEF CCS3}
  if (latob.getvalue('gp_naturepieceg') = VH_GC.AFNatProposition)  then
    AppliquerConfidentialite(Ecran,'PRO')
  else
    if (latob.getvalue('gp_naturepieceg') = VH_GC.AFNatAffaire)  then
      AppliquerConfidentialite(Ecran,'AFF')
    else
      AppliquerConfidentialite(Ecran,GetInfoParPiece(latob.getvalue('gp_naturepieceg'),'GPP_VENTEACHAT'));
  {$ENDIF}
SetControlEnabled('GP_TARIFTIERS', False);  //mcd 21/10/03 alignement sur GC ou ce champ n'est pas accessible (si modif par reporté sur les lignes ..)
   {$IFDEF CCS3}
  if (getcontrol('GBPIECE') <> Nil) then SetControlVisible ('GBPIECE', False);
  if (getcontrol('GBRESSOURCE') <> Nil) then SetControlVisible ('GBRESSOURCE', False);
  {$ENDIF}
TOBAdr_Old := TOB.Create ('SAUVE LES ADRESSES',nil,-1);
{$IFDEF BTP}
	TOBRepart := TOB(TOB(LaTOB.data).Data);
	SetComponents;
{$ENDIF}

end;

procedure TOF_AFPiece.GP_REPRESENTANTExit(Sender: TObject);
Var Lib1, Lib2 : String;
    st : string;
begin

  st := Commercial.Text;

  EntreCote(st, true);

  Commercial.Text := st;

  if (Commercial.Text <> '') and (Commercial.Text <> OldCommercial) then
  begin
  	if not OKRepresentant (Commercial.text,TypeCom) then
    begin
      PGIBoxAF(TexteMsgCompl[3], Ecran.Caption);
      Commercial.SetFocus;
      Exit;
    end;
  end;

  LblCommercial.Caption := '';
  if Commercial.text <> '' then
  begin
    LibelleCommercial(Commercial.Text, Lib1, Lib2, False);
    LblCommercial.Caption := Lib1 + ' ' + Lib2;
    LblCommercial.Visible := True;
  end;

end;

function ExisteRessource(CodeRess: string): boolean;
begin
  Result := ExisteSql('SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_RESSOURCE="' + CodeRess + '"');
end;


procedure TOF_AFPiece.GP_REPRESENTANTElipsisClick(Sender: TObject);
begin

  if GetRepresentantEntete(Commercial, 'Recherche Commercial', TypeCom, nil) then
  begin
    if Commercial.Text <> OldCommercial then
    begin
      OldCommercial := Commercial.Text;
    end;
  end;

end;

Procedure TOF_AFPiece.ControleChamp(Champ, Valeur : Variant);
Var CD: R_CleDoc;
    St : String;
Begin

  if Champ = 'MODIFSITUATION' then ModifAvanc := true;

  if Champ = 'ENTETEAFFAIRE' then EnteteAffaire := true;

  if Champ = 'ENTETEAFFAIREPIECE' then EnteteAffpiece := true;

  if Champ = 'PARTICULIER' then Particulier := true;

{
  if champ = 'ACTION=CREATION' then
    AccesModeFac := False // pas d'accès au mode de facturation en creation
  else
  begin
    if ((LaTOB.GetValue('GP_AFFAIREDEVIS') = '')  or        // pas d'accès au mode de facturation si pas d'affaire liée au devis
        (LaTOB.GetValue('GP_AFFAIRE') = ''))      then      // pas d'accès au mode de facturation si pas d'affaire
      AccesModeFac := False
    else if (LaTob.GetValue('GP_DEVENIRPIECE') <> '') then  // pas d'accès au mode de facturation si devis déjà facturé
    begin
      AccesModeFac := False;
    end else
      AccesModeFac := True;
  end;
}
  if (LaTOB.GetValue('GP_AFFAIRE') = '') then      // pas d'accès au mode de facturation si pas d'affaire
    AccesModeFac := False
  else if (LaTob.GetValue('GP_DEVENIRPIECE') <> '') then  // pas d'accès au mode de facturation si devis déjà facturé
  begin
    AccesModeFac := False;
    // vérification existence pièce :
    St := LaTob.GetValue('GP_DEVENIRPIECE');
    DecodeRefPiece(St, CD);
    if Not ExistePiece(CD) then AccesModeFac := True;
  end else
    AccesModeFac := True;

end;

procedure TOF_AFPiece.OnCancel;
begin
  inherited;
  TOBAdresses.Free;
end;

procedure TOF_AFPiece.OnLoad;
var TOBAdr: TOB;
  sRIB, Etab, Guichet, Numero, Cle, Dom, stAlign: string;
  i: integer;
  fPaAdr: Ttabsheet; // panel contenant l'adresse
  i_ind : integer;
  Titre : string;
  {$IFDEF BTP}
  TypeFacturation: string;
  Q: TQuery;
  Req: string;
  {$ENDIF}
  OneTOB : TOB;
begin
  inherited;
  TypeCom := GetInfoParPiece(LaTOB.GetValue('GP_NATUREPIECEG'), 'GPP_TYPECOMMERCIAL');

  // Gestion des adresses de facturation

  TobAdresses := TOB(LaTob.data);
  (* mcd 14/02/03if TobAdresses.Detail.count >= 2 then
     begin
     TOBAdr := TOBAdresses.Detail[1];
     TobAdr_Old := Tob.Create('',nil,-1); TobAdr_Old.Dupliquer(TobAdr,False,True);
     TOBAdr.PutEcran(Ecran);
     end; *)

  if TobAdresses.Detail.count <> 0 then // mcd 06/05/2003 pour la saisie affaire, pas d'adresse!....
  begin
    if LaTOB.GetValue('GP_NUMADRESSEFACT') = LaTOB.GetValue('GP_NUMADRESSELIVR') then CP_ADR_IDENT := True;
    //mcd 14/03/02 pour traiter adresse sur Adresse ou PieceAdresse
    if OkPieceAdresse then
    begin
      // BRL 17/11/03 FQ N°10206
      if TOBAdresses.Detail.Count = 1 then
      begin
        TOBAdr := TOB.Create('PIECEADRESSE', TOBAdresses, -1);
        TOBAdr.Dupliquer(TOBAdresses.Detail[0], False, True);
      end;
      OneTOB := TOBAdresses.Detail[0];
      // gestion des adresses de Livraison sur piece d'achat et de vente
      ChargementAdresseTobEcran('LIV', 'ADL', 'GPA', OneTOB);
      // gestion des adresses de Facturation sur piece d'achat et de vente
      OneTOB := TOBAdresses.Detail[1];
      ChargementAdresseTobEcran('FAC', 'ADR', 'GPA', OneTOB);
    end
    else
    begin
      // gestion des adresses de Livraison sur piece d'achat et de vente
      ChargementAdresseTobEcran('LIV', 'ADL', 'ADR', TOBAdresses.Detail[0]);
      // gestion des adresses de Facturation sur piece d'achat et de vente
      ChargementAdresseTobEcran('FAC', 'ADR', 'ADR', TOBAdresses.Detail[0]);
    end;
  end;

  (* MODIF LS ----> archi faux
  // gm 01/04/04 sauvegarde des valeurs avant modif
  //  TobAdr_Old := Tob.Create('ADRESSES', nil, -1);
  //  TOBAdr_Old.GetEcran(Ecran, nil);
  *)
  TOBADR_Old.Dupliquer (TOBAdresses,true,true); // sauvegarde des adresses origine

  {$IFDEF BTP}
  //SetControlText('GP_NADRESSELIV',IntToStr(GetNumAdresseFromTiers('', GetControlText('GP_TIERSLIVRE'), taLivr)));
  SetControlText('GP_NADRESSELIV', GetNumAdresseFromTiers('', GetControlText('GP_TIERSLIVRE'), taLivr));
	if ExisteAdresseLiv (LaTOB.GetValue('GP_TIERSLIVRE')) then
  begin
     THEdit(GetControl('LNUMLIVR')).visible := true;
     THEdit(GetControl('GP_NADRESSELIV')).visible := true;
  end;
  {$ENDIF}

  // Gestion du RIB
  sRIB := LaTOB.GetValue('GP_RIB');
  RibEnc := sRIB;
  if (sRIB <> '') then
  begin
    DecodeRIB(Etab, Guichet, Numero, Cle, Dom, sRIB);
    SetcontrolText('ETABRIB', Etab);
    SetcontrolText('GUICHETRIB', Guichet);
    SetcontrolText('NUMERORIB', Numero);
    SetcontrolText('CLERIB', Cle);
    SetcontrolText('DOMRIB', Dom);
  end;

  // Gestion alignement client
  stAlign := LaTob.GetValue('GP_MAJLIBRETIERS');
  if stAlign = '' then stAlign := 'XXXXXXXXXX';
  for i := 1 to 10 do
  begin
    if TCheckBox(GetControl('ALIGNETIERS' + intTostr(i))) <> Nil then
      SetControlText('ALIGNETIERS' + intTostr(i), Copy(stAlign, i, 1));
  end;

  if TCheckBox(GetControl('ALIGNEREPRESENTANT')) <> Nil then
    SetControlText('ALIGNEREPRESENTANT', Copy(stAlign, 11, 1));

  // Alim champs payeur
  if LaTob.GetValue('GP_TIERSPAYEUR') <> LaTob.GetValue('GP_TIERS') then
    SetControlText('TIERSPAYEUR', LaTob.GetValue('GP_TIERSPAYEUR'));
  if LaTob.GetValue('GP_APPORTEUR') <> '' then
    SetControlText('LIBAPPORTEUR', RechDom('AFAPPORTEUR', LaTob.GetValue('GP_APPORTEUR'), False));

  {$IFDEF BTP}
  // gestion du destinataire de livraison fournisseur
  if LaTOB.GetValue('GP_IDENTIFIANTWOT')=-1 then TcheckBox(getCOntrol('LIVRECEPFOUR')).Checked := true;

  // Gestion du mode de facturation
  if LaTOB.GetValue('GP_AFFAIREDEVIS') = '' then
  begin
    if LaTOB.GetValue('AFF_GENERAUTO') <> '' then TypeFacturation := LaTOB.GetValue('AFF_GENERAUTO')
    else TypeFacturation := 'DIR';
  end else
    TypeFacturation := RenvoieTypeFact(LaTOB.GetValue('GP_AFFAIREDEVIS'));

  SetcontrolText('MODEFACTURATION', TypeFacturation);
  if AccesModeFac = False then SetControlEnabled('MODEFACTURATION', False);
  
{
  if AccesModeFac = True then
    SetcontrolText('MODEFACTURATION', TypeFacturation)
  else
  begin
    if LaTOB.GetValue('GP_AFFAIRE') = '' then
      SetcontrolText('MODEFACTURATION', 'DIR')
    else
      SetcontrolText('MODEFACTURATION', TypeFacturation);

    SetControlEnabled('MODEFACTURATION', False);
  end;
}

  // FQ 12073
  if (Pos(LaTOB.getValue('GP_NATUREPIECEG'),'FBT;FBP')>0) then SetControlEnabled('MODEFACTURATION', False);
  // ---
  THSpinEdit(GetControl('RENDUACOMPTE')).value := GetMiniAvancAcompte(LaTOB.GetValue('GP_AFFAIREDEVIS'));

  // Traitement pour les situations
  if ((TypeFacturation = 'AVA') or (TypeFacturation = 'DAC')) and
  	 (Pos(LaTOB.GetValue('GP_NATUREPIECEG'),'FBT;FBP;BAC;DAC')>0) then
  begin
    SetControlVisible('DATESIT', True);
    SetControlEnabled('DATESIT', True);
    SetControlVisible('TDATESIT', True);
    SetControlEnabled('BValider', True);

    // récupération numéro et date de situation
    Req := 'SELECT BST_NUMEROSIT, BST_DATESIT FROM BSITUATIONS WHERE BST_NATUREPIECE="' + LaTob.GetValue('GP_NATUREPIECEG') + '" AND BST_SOUCHE="' +
      LaTob.GetValue('GP_SOUCHE') + '" AND BST_NUMEROFAC=' + IntToStr(LaTob.GetValue('GP_NUMERO'));
    Q := OpenSQL(Req, TRUE);
    if not Q.EOF then
    begin
      Q.Findfirst;
      SetControlText('TDATESIT', 'Situation N° ' + IntToStr(Q.Fields[0].AsInteger) + ' du ');
      SetControlText('DATESIT', FormatDateTime('dd/mm/yyyy', Q.Fields[1].AsDateTime));
      DateSituation := GetControlText('DATESIT');
    end;
    Ferme(Q);
  end;

  TiersFacture := GetControlText('GP_TIERSFACTURE');

  AffichageMarge;

  if LaTOB.getValue('GP_NATUREPIECEG')='DBT' Then TCheckBox(GetCOntrol('AFF_OKSIZERO')).visible := true;
  {$ENDIF} // BTP

  if LaTob.GetValue('GP_VENTEACHAT') = 'ACH' then
  begin
    SetControlText('TGP_TIERS','Fournisseur');
    // BBI, correction fiche 10374
    SetControlProperty('GP_TIERSLIVRE','DATATYPE','GCTIERSFOURN'); //Modif du 12/06/01
    SetControlProperty('GP_TIERSFACTURE','DATATYPE','GCTIERSFOURN');
    // BBI, fin correction fiche 10374
    for i_ind:=1 to 3 do SetControlProperty('GP_LIBRETIERS'+IntToStr(i_ind),'DATATYPE','GCLIBREFOU'+IntToStr(i_ind));
    for i_ind:=4 to 9 do
    begin
      SetControlVisible('TGP_LIBRETIERS'+IntToStr(i_ind),False);
      SetControlVisible('GP_LIBRETIERS'+IntToStr(i_ind),False);
    end ;
    SetControlVisible('TGP_LIBRETIERSA',False);
    SetControlVisible('GP_LIBRETIERSA',False);
    for i_ind:=1 to 3 do
    begin
      GCTitreZoneLibre('YTC_TABLELIBREFOU'+IntToStr(i_ind),Titre);
      if Titre = '' then
      begin
      SetControlVisible('TGP_LIBRETIERS'+IntToStr(i_ind),False);
      SetControlVisible('GP_LIBRETIERS'+IntToStr(i_ind),False);
      end else SetControlCaption('TGP_LIBRETIERS'+IntToStr(i_ind), Titre) ;
    end;


    SetControlProperty('GBCLIENT', 'CAPTION','Fournisseur') ;
  end;
  for i_ind:=1 to 3 do
  begin
    GCTitreZoneLibre('TGP_DATELIBREPIECE'+IntToStr(i_ind),Titre);
    if Titre = '' then
    begin
    SetControlVisible('TGP_DATELIBREPIECE'+IntToStr(i_ind),False);
    SetControlVisible('GP_DATELIBREPIECE'+IntToStr(i_ind),False);
    end else SetControlCaption('TGP_DATELIBREPIECE'+IntToStr(i_ind), Titre) ;
  end;

  if (Pos(laTOB.GEtVAlue('GP_NATUREPIECEG'),'FBT;FBP')>0) and (Pos(RenvoieTypeFact(laTOB.GetValue('GP_AFFAIREDEVIS')),'DAC;AVA')>0 ) then
  begin
    //SetCOntrolEnabled ('TAXES',False); : modif BRL 15/12/2015
    SetCOntrolEnabled ('MODEFACTURATION',False);
    //SetCOntrolEnabled ('GP_REGIMETAXE',False); : modif BRL 15/12/2015
  end;

if LaTOB.GetValue('ISDEJAFACT')='X' then
begin
	SetCOntrolEnabled ('TAXES',False);
	SetCOntrolEnabled ('MODEFACTURATION',False);
	SetCOntrolEnabled ('GP_REGIMETAXE',False);
end;
//uniquement en line
//SetShowForLine;

end;


Procedure TOF_AFPiece.ChargementAdresseTobEcran(TypeAdresse, Prefixe1, Prefixe2 : string; TOBLAdresses : TOB);
Begin
  // gestion des adresses de livraisons sur piece d'achat et de vente
  if TOBLAdresses.GetValue(Prefixe2 + '_LIBELLE') <> '' then
    SetControlText(Prefixe1 + '_LIBELLE',  TOBLAdresses.GetValue(Prefixe2 + '_LIBELLE'));
  //
  SetControlText(Prefixe1 + '_LIBELLE2',      TOBLAdresses.GetValue(Prefixe2 + '_LIBELLE2'));
  SetControlText(Prefixe1 + '_JURIDIQUE',     TOBLAdresses.GetValue(Prefixe2 + '_JURIDIQUE'));
  SetControlText(Prefixe1 + '_ADRESSE1',      TOBLAdresses.GetValue(Prefixe2 + '_ADRESSE1'));
  SetControlText(Prefixe1 + '_ADRESSE2',      TOBLAdresses.GetValue(Prefixe2 + '_ADRESSE2'));
  SetControlText(Prefixe1 + '_ADRESSE3',      TOBLAdresses.GetValue(Prefixe2 + '_ADRESSE3'));
  SetControlText(Prefixe1 + '_CODEPOSTAL',    TOBLAdresses.GetValue(Prefixe2 + '_CODEPOSTAL'));
  SetControlText(Prefixe1 + '_VILLE',         TOBLAdresses.GetValue(Prefixe2 + '_VILLE'));
  SetControlText(Prefixe1 + '_PAYS',          TOBLAdresses.GetValue(Prefixe2 + '_PAYS'));
  SetControlText(Prefixe1 + '_NUMEROCONTACT', TOBLAdresses.GetValue(Prefixe2 + '_NUMEROCONTACT'));
  //
  SetInfosContact(TypeAdresse, TOBLAdresses.GetValue(Prefixe2 + '_NUMEROCONTACT'), '');

End;

procedure TOF_AFPiece.OnUpdate;
var TraiteLivre, TraiteRepresentant: boolean;
  Etab, Guichet, Numero, Cle, Dom, sRIB, CleCalcul, stAlign: string;
  i: integer;
  {$IFDEF BTP}
  CodReg, Req, TypeFac: string;
  QQ : Tquery;
  {$ENDIF}
  {$IFNDEF CCS3}
  NomChamp,Stnat  : string;
  {$ENDIF}
begin
  inherited;
{*
  if IsAdresseModifie then
   begin
   Cp_Adr_Ident:=False; //mcd 30/10/03 l'adresse à été modifer à la main, il fat forcer qu'elle est # de l'adresse origine
   if OkPieceAdresse then SetControlText('GP_NUMADRESSEFACT', IntToStr(2))
    else SetControlText('GP_NUMADRESSEFACT', IntToStr(-2));
   end;
*}
  TraiteLivre := True;
  TraiteRepresentant := True;
  if ctxScot in V_PGI.PGIContexte then
  begin
    TraiteLivre := False;
    TraiteRepresentant := False;
  end;
  if (ctxAffaire in V_PGI.PGIContexte) and
      (GetControl ('GP_TIERSLIVRE') <> Nil) and  (GetControlText('GP_TIERSLIVRE') = '')
      then traiteLivre := False;
  if (ctxAffaire in V_PGI.PGIContexte) and (GetControlText('GP_REPRESENTANT') = '')
    then traiteRepresentant := False;

  {$IFNDEF CCS3}
  if (latob.getvalue('GP_NATUREPIECEG') = VH_GC.AFNatProposition)  then
    Stnat := 'PRO'
  else
    if (latob.getvalue('GP_NATUREPIECEG') = VH_GC.AFNatAffaire)  then
      Stnat := 'AFF'
    else
      Stnat := GetInfoParPiece(LaTOB.GetValue('GP_NATUREPIECEG'),'GPP_VENTEACHAT');

  NomChamp:=VerifierChampsObligatoires(Ecran,Stnat);
  if NomChamp<>'' then
    begin
    NomChamp:=ReadTokenSt(NomChamp);
    SetFocusControl(NomChamp) ;
    PGIBox('La saisie du champs suivant est obligatoire : '+champToLibelle(NomChamp),Ecran.Caption);
    TForm(Ecran).ModalResult:=0;
    Exit ;
    end;
  {$ENDIF}

  if (TraiteLivre)  and
  (GetControl ('GP_TIERSLIVRE') <> Nil) and (GetControlText('GP_TIERSLIVRE') <> '') then
  begin
    if not ExisteSQL('SELECT T_TIERS FROM TIERS WHERE T_TIERS="' + GetControlText('GP_TIERSLIVRE') + '"') then
    begin
      PGIBoxAf(TexteMsgCompl[1], Ecran.Caption);
      SetActiveTabSheet('GENERAL');
      SetFocusControl('GP_TIERSLIVRE');
      TForm(Ecran).ModalResult := 0;
      Exit;
    end;
  end;

  // Vérif du tiers facturé
  if GetControlText('GP_TIERSFACTURE') = '' then SetControlText('GP_TIERSFACTURE', LaTob.GetValue('GP_TIERS'));
  if not (GetControlText('GP_TIERSFACTURE') = LaTob.GetValue('GP_TIERS')) then
  begin
    if not ExisteSQL('SELECT T_TIERS FROM TIERS WHERE T_TIERS="' + GetControlText('GP_TIERSFACTURE') + '"') then
    begin
      PGIBoxAf(TexteMsgCompl[2], Ecran.Caption);
      SetActiveTabSheet('INFORMATIONS');
      SetFocusControl('GP_TIERSFACTURE');
      TForm(Ecran).ModalResult := 0;
      Exit;
    end;
  end;
  if not (GetControlText('GP_TIERSFACTURE') = TiersFacture) then
  begin
{$IFDEF BTP}
		SetModeRegle;
    TiersFacture := GetControlText('GP_TIERSFACTURE');
{$ENDIF}
  end;
  //gm 05/09/02 bug , si on modifie le Tiers Facturé, son adresse n'est pas prise en compte
  // en saisie de facture
  // a ne faire que si on arrive depuis la saisie de facture    (pas depuis affaire)
  (* -----------> la je tombe par terre --- :\
  if not (EnteteAffaire) and not (EnteteAffPiece) then
  begin
    if not (IsAdresseModifie) then // gm 01/04/04
      TrtAdresseSpecifAffaire;
  end;
  *)
  LaTob.PutValue('GP_TIERSFACTURE', GetControlText('GP_TIERSFACTURE'));
  // Vérif du tiers payeur
  if GetControlText('TIERSPAYEUR') <> '' then SetControlText('GP_TIERSPAYEUR', GetControlText('TIERSPAYEUR'))
  else SetControlText('GP_TIERSPAYEUR', LaTob.GetValue('GP_TIERS'));
  if not (GetControlText('GP_TIERSPAYEUR') = LaTob.GetValue('GP_TIERS')) then
  begin
    if not ExisteSQL('SELECT T_TIERS FROM TIERS WHERE T_TIERS="' + GetControlText('GP_TIERSPAYEUR') + '"') then
    begin
      PGIBoxAf(TexteMsgCompl[8], Ecran.Caption);
      SetActiveTabSheet('INFORMATIONS');
      SetFocusControl('GP_TIERSPAYEUR');
      TForm(Ecran).ModalResult := 0;
      Exit;
    end
  end;
  LaTob.PutValue('GP_TIERSPAYEUR', GetControlText('GP_TIERSPAYEUR'));

  // Contrôle du commercial
  if (TraiteRepresentant) and (Commercial.text <> '') then
  begin
  	Req :='SELECT GCL_COMMERCIAL FROM COMMERCIAL WHERE GCL_COMMERCIAL="' + GetControlText('GP_REPRESENTANT') + '"';
    if TypeCom <> '' then req := Req + ' AND GCL_TYPECOMMERCIAL="'+TypeCom+'"';
//    if not ExisteSQL('SELECT GCL_COMMERCIAL FROM COMMERCIAL WHERE GCL_COMMERCIAL="' + GetControlText('GP_REPRESENTANT') + '" AND GCL_TYPECOMMERCIAL="REP" ')
    if not ExisteSQL(req) then
    begin
      PGIBoxAF(TexteMsgCompl[3], Ecran.Caption);
      SetActiveTabSheet('INFORMATIONS');
      SetFocusControl('GP_REPRESENTANT');
      TForm(Ecran).ModalResult := 0;
      Exit;
    end;
  end;
  // Contrôle de l'apporteur d'affaire
  if (GetControlText('GP_APPORTEUR') <> '') then
    if not (LookupValueExist(THEdit(Getcontrol('GP_APPORTEUR')))) then
    begin
      PGIBoxAF(TexteMsgCompl[5], Ecran.Caption);
      SetFocusControl('GP_APPORTEUR');
      TForm(Ecran).ModalResult := 0;
      Exit;
    end;
  // Contrôle de la ressource saisie à la piece
  if (THEdit(Getcontrol('GP_RESSOURCE')) <> Nil) and (GetControlText('GP_RESSOURCE') <> '') then
    if not (LookupValueExist(THEdit(Getcontrol('GP_RESSOURCE')))) then
    begin
      PGIBoxAF(TexteMsgCompl[4], Ecran.Caption);
      SetFocusControl('GP_RESSOURCE');
      TForm(Ecran).ModalResult := 0;
      Exit;
    end;
  if (GetControlText('GP_TIERSSAL1') <> '') then
    if not (LookupValueExist(THEdit(Getcontrol('GP_TIERSSAL1')))) then
    begin
      PGIBoxAF(TexteMsgCompl[9], Ecran.Caption);
      SetFocusControl('GP_TIERSSAL1');
      TForm(Ecran).ModalResult := 0;
      Exit;
    end;
  if (GetControlText('GP_TIERSSAL2') <> '') then
    if not (LookupValueExist(THEdit(Getcontrol('GP_TIERSSAL2')))) then
    begin
      PGIBoxAF(TexteMsgCompl[9], Ecran.Caption);
      SetFocusControl('GP_TIERSSAL2');
      TForm(Ecran).ModalResult := 0;
      Exit;
    end;
  if (GetControlText('GP_TIERSSAL3') <> '') then
    if not (LookupValueExist(THEdit(Getcontrol('GP_TIERSSAL3')))) then
    begin
      PGIBoxAF(TexteMsgCompl[9], Ecran.Caption);
      SetFocusControl('GP_TIERSSAL3');
      TForm(Ecran).ModalResult := 0;
      Exit;
    end;
  // Gestion du RIB
  Etab := GetcontrolText('ETABRIB');
  Guichet := GetcontrolText('GUICHETRIB');
  Numero := GetcontrolText('NUMERORIB');
  Cle := GetcontrolText('CLERIB');
  Dom := GetcontrolText('DOMRIB');

  sRIB := EncodeRIB(Etab, Guichet, Numero, Cle, Dom);
  if (Trim(sRib) <> '') then
  begin
    if RIBEnc <> sRIB then
    begin
      CleCalcul := VerifRib(Etab, Guichet, Numero);
      if (CleCalcul <> Cle) then
      begin
        if PGIAskAF(Format('Clé Rib calculée: %2s  - Clé Rib saisie: %2s. Souhaitez vous conserver la clé Rib saisie', [CleCalcul, Cle]),
          'Calcul de la clé RIB') <> mrYes then
        begin
          sRIB := EncodeRIB(Etab, Guichet, Numero, CleCalcul, Dom);
        end;
      end;
    end;
    LaTOB.PutValue('GP_RIB', sRIB);
  end;
  // Gestion alignement client
  stAlign := '';
  for i := 1 to 10 do
  begin
    if TCheckBox(GetControl('ALIGNETIERS' + intTostr(i))) <> Nil then
      stalign := stAlign + GetControlText('ALIGNETIERS' + intTostr(i));
  end;
  if TCheckBox(GetControl('ALIGNEREPRESENTANT')) <> Nil then
    stalign := stAlign + GetControlText('ALIGNEREPRESENTANT');
  LaTob.PutValue('GP_MAJLIBRETIERS', stAlign);

  (*// Traitement de l'adresse Facturation
  if TobAdresses.Detail.count >= 2 then
     begin
     TOBAdresses.Detail[1].GetEcran (TForm(Ecran));
     IsAdresseModifie;
     end;   *)
{$IFDEF BTP}
  if GetParamSoc('SO_GCPIECEADRESSE') then
  begin //mcd 14/02/03
  	if TOBAdresses.detail.count = 0 then TOB.Create ('PIECEADRESSE',TOBAdresses,-1);
    TOBAdresses.Detail[0].PutValue('GPA_LIBELLE', GetControlText('ADL_LIBELLE'));
    TOBAdresses.Detail[0].PutValue('GPA_LIBELLE2', GetControlText('ADL_LIBELLE2'));
    TOBAdresses.Detail[0].PutValue('GPA_JURIDIQUE', GetControlText('ADL_JURIDIQUE'));
    TOBAdresses.Detail[0].PutValue('GPA_ADRESSE1', GetControlText('ADL_ADRESSE1'));
    TOBAdresses.Detail[0].PutValue('GPA_ADRESSE2', GetControlText('ADL_ADRESSE2'));
    TOBAdresses.Detail[0].PutValue('GPA_ADRESSE3', GetControlText('ADL_ADRESSE3'));
    TOBAdresses.Detail[0].PutValue('GPA_CODEPOSTAL', GetControlText('ADL_CODEPOSTAL'));
    TOBAdresses.Detail[0].PutValue('GPA_VILLE', GetControlText('ADL_VILLE'));
    TOBAdresses.Detail[0].PutValue('GPA_PAYS', GetControlText('ADL_PAYS'));
    TOBAdresses.Detail[0].PutValue('GPA_NUMEROCONTACT', GetControlText('ADL_NUMEROCONTACT'));
  end else
  begin
  	if TOBAdresses.detail.count = 0 then TOB.Create ('ADRESSES',TOBAdresses,-1);
    TOBAdresses.Detail[0].PutValue('ADR_LIBELLE', GetControlText('ADL_LIBELLE'));
    TOBAdresses.Detail[0].PutValue('ADR_LIBELLE2', GetControlText('ADL_LIBELLE2'));
    TOBAdresses.Detail[0].PutValue('ADR_JURIDIQUE', GetControlText('ADL_JURIDIQUE'));
    TOBAdresses.Detail[0].PutValue('ADR_ADRESSE1', GetControlText('ADL_ADRESSE1'));
    TOBAdresses.Detail[0].PutValue('ADR_ADRESSE2', GetControlText('ADL_ADRESSE2'));
    TOBAdresses.Detail[0].PutValue('ADR_ADRESSE3', GetControlText('ADL_ADRESSE3'));
    TOBAdresses.Detail[0].PutValue('ADR_CODEPOSTAL', GetControlText('ADL_CODEPOSTAL'));
    TOBAdresses.Detail[0].PutValue('ADR_VILLE', GetControlText('ADL_VILLE'));
    TOBAdresses.Detail[0].PutValue('ADR_PAYS', GetControlText('ADL_PAYS'));
    TOBAdresses.Detail[0].PutValue('ADR_TELEPHONE', GetControlText('ADL_TELEPHONE'));
    TOBAdresses.Detail[0].PutValue('ADR_CONTACT', GetControlText('ADL_CONTACT'));
    TOBAdresses.Detail[0].PutValue('ADR_NUMEROCONTACT', GetControlText('ADL_NUMEROCONTACT'));
  end;

  if TobAdresses.Detail.count <> 0 then
  Begin
    if OkPieceAdresse then
    begin
  		if TOBAdresses.detail.count = 1 then TOB.Create ('PIECEADRESSE',TOBAdresses,-1);
      //
      ChargementAdresseEcranTob('FAC', 'GPA', 'ADR',  TOBAdresses.Detail[1]);
      //
      if LaTOB.GetValue('GP_NUMADRESSEFACT') = LaTOB.GetValue('GP_NUMADRESSELIVR') then
      begin
        LaTOB.PutValue('GP_NUMADRESSEFACT', +2);
        CP_ADR_IDENT := false;
      end;
    end
    else
    begin
  	if TOBAdresses.detail.count = 1 then TOB.Create ('ADRESSES',TOBAdresses,-1);
      //
      ChargementAdresseEcranTob('LIV', 'ADR', 'ADR',  TOBAdresses.Detail[1]);
      //
      if LaTOB.GetValue('GP_NUMADRESSEFACT') = LaTOB.GetValue('GP_NUMADRESSELIVR') then
      begin
        LaTOB.PutValue('GP_NUMADRESSEFACT', -2);
        CP_ADR_IDENT := false;
      end;
    end;
  end
  else
  begin
    if OkPieceAdresse then
    begin //mcd 14/02/03
      //Chargement adresse de Livraison
      TOB.Create ('PIECEADRESSE',TOBAdresses,-1);
      ChargementAdresseEcranTob('LIV', 'GPA', 'ADL',  TOBAdresses.Detail[0]);
    end
    else
    begin
      TOB.Create ('ADRESSES',TOBAdresses,-1);
      ChargementAdresseEcranTob('LIV', 'ADR', 'ADL',  TOBAdresses.Detail[0]);
  end;
  end;
{$ENDIF}

  LaTOB.GetEcran(TForm(Ecran));

  if CP_ADR_IDENT then
  begin
    LaTOB.PutValue('GP_NUMADRESSEFACT', LaTOB.GetValue('GP_NUMADRESSELIVR'));
{$IFNDEF BTP}
    TOBAdresses.Detail[1].Dupliquer(TOBAdresses.Detail[0], False, True);
{$else}
    TOBAdresses.Detail[0].Dupliquer(TOBAdresses.Detail[1], False, True);
{$endif}
    if OkPieceAdresse then TOBAdresses.Detail[1].PutValue('GPA_TYPEPIECEADR', '002');
  end
  else
  begin
    if TobAdresses.Detail.count <> 0 then
    Begin
      if OkPieceAdresse then
      begin
        ChargementAdresseEcranTob('FAC', 'GPA', 'ADR',  TOBAdresses.Detail[1]);
        //
        if LaTOB.GetValue('GP_NUMADRESSEFACT') = LaTOB.GetValue('GP_NUMADRESSELIVR') then
        begin
          LaTOB.PutValue('GP_NUMADRESSEFACT', 2);
          CP_ADR_IDENT := false;
        end;
      end
      else
      begin
        ChargementAdresseEcranTob('FAC', 'ADR', 'ADR',  TOBAdresses.Detail[1]);
        //
        if LaTOB.GetValue('GP_NUMADRESSEFACT') = LaTOB.GetValue('GP_NUMADRESSELIVR') then
        begin
          LaTOB.PutValue('GP_NUMADRESSEFACT', -2);
          CP_ADR_IDENT := false;
        end;
      end;
    end;
  end;

  // trt ref interne gm 18/03/2002
  if EnteteAffaire or EnteteAffPiece then
    LaTob.PutValue('GP_REFINTERNE', GetControlText('GP_REFINTERNE1'));

  {$IFDEF BTP}

  //  Maj mode de facturation pour l'affaire du devis
  if (LaTob.GetValue('GP_VENTEACHAT') = 'VEN') then
  begin
    MajTypeFact (LaTob, GetcontrolText('MODEFACTURATION'));
  end;

  // Traitement code regroupement si modification mode de facturation
  TypeFac := RenvoieTypeFact(LaTOB.GetValue('GP_AFFAIREDEVIS'));
  if GetcontrolText('MODEFACTURATION') <> TypeFac then
  begin
    // Si le nouveau mode de facturation saisi est différent de celui de l'affaire principale,
    // plus de regroupement possible sur facture.
    // s'il est identique le code regroupement doit être le même que celui de l'affaire principale
    if GetcontrolText('MODEFACTURATION') <> RenvoieTypeFact(LaTOB.GetValue('GP_AFFAIRE')) then
      CodReg := 'AUC'
    else
      CodReg := RenvoieCodeReg(LaTOB.GetValue('GP_AFFAIRE'));
    // maj nouveau mode de facturation dans l'affaire du devis
    Req := 'UPDATE AFFAIRE SET AFF_REGROUPEFACT = ' + '"' + CodReg + '"' +
      ' WHERE AFF_AFFAIRE =' + '"' + LaTOB.GetValue('GP_AFFAIREDEVIS') + '"';
    ExecuteSQL(Req);
  end;

  if TcheckBox(GetControl('LIVRECEPFOUR')).Checked then LaTOB.PutValue('GP_IDENTIFIANTWOT',-1)
                                                   else LaTOB.PutValue('GP_IDENTIFIANTWOT',0);

  (* EN ATTENTE POUR BTP
    if ((GetcontrolText ('MODEFACTURATION') = 'AVA') or (GetcontrolText ('MODEFACTURATION') = 'DAC')) AND
     (THSpinEdit(GetControl('RENDUACOMPTE')).value <> GetMiniAvancAcompte(LaTOB.GetValue('GP_AFFAIREDEVIS'))) then
    begin
    // maj De l'avancemenement mini pour declenchement restitution de l'acompte
    Req := 'UPDATE AFFAIRE SET AFF_MINAVANCACC = '+'"'+ inttostr(THSpinEdit(GetControl('RENDUACOMPTE')).value) + '" WHERE AFF_AFFAIRE ='+'"'+LaTOB.GetValue('GP_AFFAIREDEVIS')+'"';
    ExecuteSQL(Req);
    end;
  *)

  // Maj si date de situation modifiée
  if (Pos(LaTOB.GetValue('GP_NATUREPIECEG'),'FBT;FBP')>0) and
    ((TypeFac = 'AVA') or (TypeFac = 'DAC')) and
    (GetcontrolText('DATESIT') <> DateSituation) then
  begin
    DateSituation := GetControlText('DATESIT');
    Req := 'UPDATE BSITUATIONS SET BST_DATESIT =' + '"' + usdatetime(StrToDate(DateSituation)) + '" WHERE BST_NATUREPIECE="' + LaTob.GetValue('GP_NATUREPIECEG')
      + '" AND BST_SOUCHE="' + LaTob.GetValue('GP_SOUCHE') + '" AND BST_NUMEROFAC=' + IntToStr(LaTob.GetValue('GP_NUMERO'));
    ExecuteSQL(Req);
  end;
  {$ENDIF}
  // Récup Ecran Tobpiece + passage à la facture
  TheTob := LaTOB;
end;

Procedure TOF_AFPiece.ChargementAdresseEcranTob(TypeAdresse, Prefixe1, Prefixe2 : string; TOBLAdresses : TOB);
Begin

  TOBLAdresses.PutValue(Prefixe1 + '_LIBELLE',        GetControlText(Prefixe2 + '_LIBELLE'));
  TOBLAdresses.PutValue(Prefixe1 + '_LIBELLE2',       GetControlText(Prefixe2 + '_LIBELLE2'));
  TOBLAdresses.PutValue(Prefixe1 + '_JURIDIQUE',      GetControlText(Prefixe2 + '_JURIDIQUE'));
  TOBLAdresses.PutValue(Prefixe1 + '_ADRESSE1',       GetControlText(Prefixe2 + '_ADRESSE1'));
  TOBLAdresses.PutValue(Prefixe1 + '_ADRESSE2',       GetControlText(Prefixe2 + '_ADRESSE2'));
  TOBLAdresses.PutValue(Prefixe1 + '_ADRESSE3',       GetControlText(Prefixe2 + '_ADRESSE3'));
  TOBLAdresses.PutValue(Prefixe1 + '_CODEPOSTAL',     GetControlText(Prefixe2 + '_CODEPOSTAL'));
  TOBLAdresses.PutValue(Prefixe1 + '_VILLE',          GetControlText(Prefixe2 + '_VILLE'));
  TOBLAdresses.PutValue(Prefixe1 + '_PAYS',           GetControlText(Prefixe2 + '_PAYS'));
  TOBLAdresses.PutValue(Prefixe1 + '_TELEPHONE',      GetControlText(Prefixe2 + '_TELEPHONE'));
  TOBLAdresses.PutValue(Prefixe1 + '_CONTACT',        GetControlText(Prefixe2 + '_CONTACT'));
  //
  TOBLAdresses.PutValue(Prefixe1 + '_NUMEROCONTACT',  GetControlText(Prefixe2 + '_NUMEROCONTACT'));

end;


procedure TOF_AFPiece.OnClose;
begin
  TobAdr_Old.Free;
end;

{$IFDEF BTP}

procedure TOF_AFPiece.AffichageMarge;
var Pourcent, Montant, Coef, PxValo, NewMontant: double;
  IndiceMontant: Integer;
  EnHt: boolean;
begin // fct utilisée par BTP uniquement au 25/02/02
  PxValo := 0;
  Pourcent := 0;
  Montant := 0;
  Coef := 0;
  ENHt := LaTob.getValue('GP_FACTUREHT') = 'X';
  if EnHt then
  begin
    IndiceMontant := LaTob.GetNumChamp('GP_TOTALHT');
  end else
  begin
    IndiceMontant := LaTob.GetNumChamp('GP_TOTALTTC');
  end;

  NewMontant := LaTob.GetValeur(IndiceMontant);
  PxValo := LaTob.GetValue('GP_MONTANTPR');

  if PxValo <> 0 then
  begin
    Coef := arrondi(NewMontant / PxValo,4);
    Montant := arrondi(NewMontant - PxValo,V_PGI.okdecV);
  end;
  if GetParamSocSecur('SO_BTGESTIONMARQ', False) then
  begin
    if NewMontant <> 0 then Pourcent := arrondi(((NewMontant - PxValo) / NewMontant) * 100, 2)
                       else Pourcent := 0;
  end else
  begin
    if PXValo <> 0 then Pourcent := arrondi(((NewMontant - PxValo) / PxValo) * 100, 2)
                   else Pourcent := 0;
  end;

  THNumEdit(GetControl('CMARGE')).Masks.PositiveMask := '#0.0000';
  SetControlProperty('PMARGE', 'value', Pourcent);
  SetControlProperty('CMARGE', 'value', Coef);
  SetControlProperty('MMARGE', 'value', Montant);
end;
{$ENDIF}

function TOF_AFPiece.IsAdresseModifie: boolean;
var TobAdr : TOB;
    bModif : Boolean;
begin

  Result := false;

  //comment sait-on que noius sommes sur l'adresse de facturation ou l'adresse de livraison (???????)
  //Après vérification on le sait pas... c'est magique !!!
  //Sauf qu'on est nul en magie et du coup on contrôle si l'adresse de facturation n'aurait pas été modifiée en la comparant à l'adresse de livraison...
  //des génies je vous dit !!!!!!!!
  if OkPieceAdresse then
  begin
    //contrôle de l'adresse Facture
    if (TobAdr_Old.detail[1].GetValue('GPA_JURIDIQUE')  <> GetcontrolText('ADR_JURIDIQUE'))   then Result := true else
    if (TobAdr_Old.detail[1].GetValue('GPA_LIBELLE')    <> GetcontrolText('ADR_LIBELLE'))     then Result := true else
    if (TobAdr_Old.detail[1].GetValue('GPA_LIBELLE2')   <> GetcontrolText('ADR_LIBELLE2'))    then Result := true else
    if (TobAdr_Old.detail[1].GetValue('GPA_ADRESSE1')   <> GetcontrolText('ADR_ADRESSE1'))    then Result := true else
    if (TobAdr_Old.detail[1].GetValue('GPA_ADRESSE2')   <> GetcontrolText('ADR_ADRESSE2'))    then Result := true else
    if (TobAdr_Old.detail[1].GetValue('GPA_ADRESSE3')   <> GetcontrolText('ADR_ADRESSE3'))    then Result := true else
    if (TobAdr_Old.detail[1].GetValue('GPA_CODEPOSTAL') <> GetcontrolText('ADR_CODEPOSTAL'))  then Result := true else
    if (TobAdr_Old.detail[1].GetValue('GPA_VILLE')      <> GetcontrolText('ADR_VILLE'))       then Result := true else
    if (TobAdr_Old.detail[1].GetValue('GPA_PAYS')       <> GetcontrolText('ADR_PAYS'))        then Result := true;
  end
  else
  begin
    if (TobAdr_Old.detail[1].GetValue('ADR_JURIDIQUE')  <> GetcontrolText('ADR_JURIDIQUE'))   then Result := true else
    if (TobAdr_Old.detail[1].GetValue('ADR_LIBELLE')    <> GetcontrolText('ADR_LIBELLE'))     then Result := true else
    if (TobAdr_Old.detail[1].GetValue('ADR_LIBELLE2')   <> GetcontrolText('ADR_LIBELLE2'))    then Result := true else
    if (TobAdr_Old.detail[1].GetValue('ADR_TELEPHONE')  <> GetcontrolText('ADR_TELEPHONE'))   then Result := true else
    if (TobAdr_Old.detail[1].GetValue('ADR_ADRESSE1')   <> GetcontrolText('ADR_ADRESSE1'))    then Result := true else
    if (TobAdr_Old.detail[1].GetValue('ADR_ADRESSE2')   <> GetcontrolText('ADR_ADRESSE2'))    then Result := true else
    if (TobAdr_Old.detail[1].GetValue('ADR_ADRESSE3')   <> GetcontrolText('ADR_ADRESSE3'))    then Result := true else
    if (TobAdr_Old.detail[1].GetValue('ADR_CODEPOSTAL') <> GetcontrolText('ADR_CODEPOSTAL'))  then Result := true else
    if (TobAdr_Old.detail[1].GetValue('ADR_VILLE')      <> GetcontrolText('ADR_VILLE'))       then Result := true else
    if (TobAdr_Old.detail[1].GetValue('ADR_PAYS')       <> GetcontrolText('ADR_PAYS'))        then Result := true;
  end;

end;

procedure TOF_AFPiece.SelectAdresseTiers;
var
  rep, TiersAdr: string;
  TobAdr: TOB;
  Q: TQuery;
begin

  if  GetControlText('GP_TIERSFACTURE') <> '' then
    TiersAdr:=GetControlText('GP_TIERSFACTURE')
  else
    TiersAdr:=GetControlText('GP_TIERS');

  rep := AGLLanceFiche('AFF', 'AFADRTIERS_MUL', '', '', 'TIERS=' + TiersAdr + ';ACTION=CONSULTATION');
  if (rep <> '') then
  begin
    Cp_Adr_Ident:=False; //mcd 22/10/2003 si on récupère une autre adresse, il faut  considérer quelle est # de l'adresse d'origine (0)= Livraison
    TobAdr := Tob.Create('ADRESSES', nil, -1);
    Q := nil;
    try
      // un seul enrgt, on peut tout prendre
      Q := OpenSQL('SELECT * FROM ADRESSES WHERE ADR_NUMEROADRESSE="' + rep + '"', True);
      if not TobAdr.SelectDB('', Q) then exit;
    finally
      ferme(Q);
    end;
    ChargementAdresseTobEcran('FAC', 'ADR', 'ADR', TobAdr);

    {*
    SetControlText('ADR_JURIDIQUE', TobAdr.GetValue('ADR_JURIDIQUE'));
    SetControlText('ADR_LIBELLE', TobAdr.GetValue('ADR_LIBELLE'));
    SetControlText('ADR_LIBELLE2', TobAdr.GetValue('ADR_LIBELLE2'));
    SetControlText('ADR_ADRESSE1', TobAdr.GetValue('ADR_ADRESSE1'));
    SetControlText('ADR_ADRESSE2', TobAdr.GetValue('ADR_ADRESSE2'));
    SetControlText('ADR_ADRESSE3', TobAdr.GetValue('ADR_ADRESSE3'));
    SetControlText('ADR_CODEPOSTAL', TobAdr.GetValue('ADR_CODEPOSTAL'));
    SetControlText('ADR_VILLE', TobAdr.GetValue('ADR_VILLE'));
    SetControlText('ADR_PAYS', TobAdr.GetValue('ADR_PAYS'));
    SetControlText('ADR_TELEPHONE', TobAdr.GetValue('ADR_TELEPHONE'));
    *}

    //mcd 07/05/03  ajout test
    if OkPieceAdresse then SetControlText('GP_NUMADRESSEFACT', IntToStr(2))
    else SetControlText('GP_NUMADRESSEFACT', IntToStr(-2));

    if LaTob.GetValue('GP_VENTEACHAT') = 'ACH' then
      PutEcranTobPieceAdresseLivraison( TobAdr,True );

    TobAdr.Free; //GISE
  end;
end;

/////////////////////////////////////////////////////////////////////////////

procedure TOF_AFLignePiece.FormuleOnChange(SEnder: TObject); //Affaire-ONYX
var
  vSt: string;
begin
  if THEdit(sender).Name = 'GL_FORCODE1' then
  begin
    vSt := RechDom('AFTREVFORMULE', GetControlText('GL_FORCODE1'), False);
    if vSt <> 'Error' then
      SetControlText('LIBELLEFORMULE1', vSt)
    else
      SetControlText('LIBELLEFORMULE1', '');
  end
  else
  begin
    vSt := RechDom('AFTREVFORMULE', GetControlText('GL_FORCODE2'), False);
    if vSt <> 'Error' then
      SetControlText('LIBELLEFORMULE2', vSt)
    else
      SetControlText('LIBELLEFORMULE2', '');
  end;
end;

procedure TOF_AFLignePiece.OnCancel;
begin
  inherited;
end;

procedure TOF_AFLignePiece.OnLoad;
var Nature, TiersAff: string;
begin
  inherited;
  TypeCom := GetInfoParPiece(LaTOB.GetValue('GL_NATUREPIECEG'), 'GPP_TYPECOMMERCIAL');

  if ctxAffaire in V_PGI.PGIContexte then
  begin
    ChargeCleAffaire(nil, THEDIT(GetControl('GL_AFFAIRE1')), THEDIT(GetControl('GL_AFFAIRE2')),
      THEDIT(GetControl('GL_AFFAIRE3')), THEDIT(GetControl('GL_AVENANT')), nil, taCreat, GetControlText('GL_AFFAIRE'), False);
  end;
  LaRessource := LaTOB.GetValue('GL_RESSOURCE');
  Nature := LaTOB.GetValue('GL_NATUREPIECEG');
  CodeTiers := LaTOB.GetValue('GL_TIERS');
  if (Nature = VH_GC.AFNatAffaire) or (Nature = VH_GC.AFNatProposition) then PieceAffaire := true
  else PieceAffaire := false;
  if PieceAffaire then
  begin
    SetControlEnabled('GL_AFFAIRE1', False);
    SetControlEnabled('GL_AFFAIRE2', False);
    SetControlEnabled('GL_AFFAIRE3', False);
    setControlVisible('BSELECTAFF1', False);
  end else
  begin
    SetControlEnabled('GL_FORCODE1', False);
    SetControlEnabled('GL_FORCODE2', False);
  end;
  if (LaTOB.GetValue('GL_AFFAIRE') <> '') and not (bVente) then
  begin
    TiersAff := GetChampsAffaire(LaTOB.GetValue('GL_AFFAIRE'), 'AFF_TIERS');
    SetControlText('TIERSAFFAIRE', tiersAff);
  end;
  if VH_GC.GCIfDefCEGID then
  begin
    SetControlVisible('GL_PMRP', False);
    SetControlVisible('GL_DPR', True);
  end;

  {$IFNDEF CCS3}
  if (latob.getvalue('GL_NATUREPIECEG') = VH_GC.AFNatProposition)  then
    AppliquerConfidentialite(Ecran,'PRO')
  else
    if (latob.getvalue('GL_NATUREPIECEG') = VH_GC.AFNatAffaire)  then
      AppliquerConfidentialite(Ecran,'AFF')
    else
      AppliquerConfidentialite(Ecran,GetInfoParPiece(latob.getvalue('GL_NATUREPIECEG'),'GPP_VENTEACHAT'));
  {$ENDIF}
end;

procedure TOF_AFLignePiece.OnUpdate;
var Aff, Aff1, Aff2, Aff3, Aff4, Tiers: THEdit;
  NbAff: integer;
  Part0, Part1, Part2, Part3, Part4: string;
  {$IFNDEF CCS3}
  Stnat,NomChamp : String;
  {$ENDIF}
  Req : String;
begin
  inherited;

  {$IFNDEF CCS3}
  if (latob.getvalue('GL_NATUREPIECEG') = VH_GC.AFNatProposition)  then
    Stnat := 'PRO'
  else
    if (latob.getvalue('GL_NATUREPIECEG') = VH_GC.AFNatAffaire)  then
      Stnat := 'AFF'
    else
      Stnat := GetInfoParPiece(LaTOB.GetValue('GL_NATUREPIECEG'),'GPP_VENTEACHAT');

  NomChamp:=VerifierChampsObligatoires(Ecran,stnat);
  if NomChamp<>'' then
  begin
    NomChamp:=ReadTokenSt(NomChamp);
    SetFocusControl(NomChamp) ;
    PGIBox('La saisie du champs suivant est obligatoire : '+champToLibelle(NomChamp),Ecran.Caption);
    TForm(Ecran).ModalResult:=0;
    Exit ;
  end;
  {$ENDIF}
  // Contrôle de la ressource saisie à la ligne
  if (GetControlText('GL_RESSOURCE') <> '') then
    if not (LookupValueExist(THEdit(Getcontrol('GL_RESSOURCE')))) then
    begin
      PGIBoxAF(TexteMsgCompl[4], Ecran.Caption);
      SetFocusControl('GL_RESSOURCE');
      TForm(Ecran).ModalResult := 0;
      Exit;
    end;
  // Contrôle de l'apporteur d'affaire
  if (GetControlText('GL_APPORTEUR') <> '') then
    if not (LookupValueExist(THEdit(Getcontrol('GL_APPORTEUR')))) then
    begin
      PGIBoxAF(TexteMsgCompl[5], Ecran.Caption);
      SetFocusControl('GL_APPORTEUR');
      TForm(Ecran).ModalResult := 0;
      Exit;
    end;
  // Contrôle du représentant
  if (GetControlText('GL_REPRESENTANT') <> '') then
    // pb tablette...  if Not(LookupValueExist(THEdit(Getcontrol('GL_REPRESENTANT')))) then
  	Req :='SELECT GCL_COMMERCIAL FROM COMMERCIAL WHERE GCL_COMMERCIAL="' + GetControlText('GL_REPRESENTANT') + '"';
    if TypeCom <> '' then req := Req + ' AND GCL_TYPECOMMERCIAL="'+TypeCom+'"';
//    if not ExisteSQL('SELECT GCL_COMMERCIAL FROM COMMERCIAL WHERE GCL_COMMERCIAL="' + GetControlText('GL_REPRESENTANT') + '" AND GCL_TYPECOMMERCIAL="REP" ')
    if not ExisteSQL(Req) then
    begin
      PGIBoxAF(TexteMsgCompl[3], Ecran.Caption);
      SetFocusControl('GL_REPRESENTANT');
      TForm(Ecran).ModalResult := 0;
      Exit;
    end;
  if not PieceAffaire then
  begin
    // Contrôle de l'affaire
    Aff := THEdit(GetControl('GL_AFFAIRE'));
    Aff1 := THEdit(GetControl('GL_AFFAIRE1'));
    Aff2 := THEdit(GetControl('GL_AFFAIRE2'));
    Aff3 := THEdit(GetControl('GL_AFFAIRE3'));
    Aff4 := THEdit(GetControl('GL_AVENANT'));
    if bVente then Tiers := THEdit(GetControl('GL_TIERS'))
    else Tiers := THEdit(GetControl('TIERSAFFAIRE'));
    // controle du tiers associé à l'affaire
    if (Codetiers <> Tiers.Text) and (bVente) then
    begin
      SetFocusControl('GL_AFFAIRE1');
      PGIBoxAF(TexteMsgCompl[6], Ecran.Caption);
      SetFocusControl('GL_AFFAIRE1');
      TForm(Ecran).ModalResult := 0;
      Exit;
    end;

    if (Aff1.Text = '') and (Aff2.Text = '') and (Aff3.Text = '') then
      NbAff := -1 // pas de test sur l'affaire
    else
      NbAff := TestCleAffaire(Aff, Aff1, Aff2, Aff3, Aff4, TIERS, 'A', True, False, False);
    if NbAff = 0 then
    begin
      //mcd 04/07/02 déjà message dans TestCleaffaire PGIBoxAF(TexteMsgCompl[7],Ecran.Caption);
      SetFocusControl('GL_AFFAIRE1');
      TForm(Ecran).ModalResult := 0;
      Exit;
    end
    else if NbAff > 1 then //mcd ajout il faut obligaotirement sélecionner un code complet
    begin
      PGIBoxAF(TexteMsgCompl[7], Ecran.Caption);
      SetFocusControl('GL_AFFAIRE1');
      TForm(Ecran).ModalResult := 0;
      Exit;
    end
    else
    begin // mcd 03/07/02 pour si code affaire existe, bien alimenté tous les champs de l'affaire correctment
      {$IFDEF BTP}
      BTPCodeAffaireDecoupe(Aff.text, Part0, Part1, Part2, Part3, Part4, Tacreat, False);      
      {$ELSE}
      CodeAffaireDecoupe(Aff.text, Part0, Part1, Part2, Part3, Part4, Tacreat, False);
      {$ENDIF}
      Latob.Putvalue('GL_AFFAIRE1', Part1);
      Latob.Putvalue('GL_AFFAIRE2', Part2);
      Latob.Putvalue('GL_AFFAIRE3', Part3);
      Latob.Putvalue('GL_AVENANT', Part4);
      Latob.Putvalue('GL_AFFAIRE', Aff.text);
    end;
  end;
  TheTob := LaTOB;
end;

procedure TOF_AFLignePiece.OnArgument(stArgument: string);
var TLab: THLabel;
  TTab: TTabSheet;
  BRechForm: TToolBarButton97;
  BFormuleVar: boolean;
  Critere, Champ, valeur: string;
  x: integer;
begin
  inherited;
  bChangetiers := not (bVente);
  BFormuleVar := False;
  Critere := (Trim(ReadTokenSt(stArgument)));
  while (Critere <> '') do
  begin
    if Critere <> '' then
    begin
      X := pos('=', Critere);
      if x <> 0 then
      begin
        Champ := copy(Critere, 1, X - 1);
        Valeur := Copy(Critere, X + 1, length(Critere) - X);
        if (Champ = 'ACTION') and (valeur = 'CREATION') then Action := taCreat
        else if (Champ = 'ACTION') and (valeur = 'MODIFICATION') then Action := taModif
        else if (Champ = 'ACTION') and (valeur = 'CONSULTATION') then Action := taConsult
        else if (Champ = 'GENERAUTO') then AffGenerauto := Valeur;
      end
      else if Critere = 'FORMULEVAR' then BFormuleVar := true;
    end;
    Critere := (Trim(ReadTokenSt(stArgument)));
  end;
  if ctxScot in V_PGI.PGIContexte then
  begin
    SetControlVisible('TGL_REPRESENTANT', False);
    SetControlVisible('GL_REPRESENTANT', False);
    SetControlVisible('GL_COMMISSIONR', False);
    //mcd 22/10/03 ?? maintenant on gère les tarifs SetControlVisible('GBTARIF', False);
  end;
  SetcontrolVisible('TIERSAFFAIRE', not (bVente));
  SetcontrolVisible('LIBTIERSAFFAIRE', not (bVente));
  SetcontrolVisible('BEFFACEAFF1', not (bVente));
  if not (bVente) then SetControlText('TGL_TIERS', '&Fournisseur');

  //fait dans tof ancetre GCMAJChampLibre(TForm(Ecran), False, 'COMBO', 'GL_LIBREART', 10, '');
  TLab := THLabel(Ecran.FindComponent('TGL_FAMILLEART1'));
  if tlab <> nil then SetCOntrolText('TGL_FAMILLEART1', RechDom('GCLIBFAMILLE', 'LF1', false));
//Non en line
  TLab := THLabel(Ecran.FindComponent('TGL_FAMILLEART2'));
  if tlab <> nil then SetCOntrolText('TGL_FAMILLEART2', RechDom('GCLIBFAMILLE', 'LF2', false));
  TLab := THLabel(Ecran.FindComponent('TGL_FAMILLEART3'));
  if tlab <> nil then SetCOntrolText('TGL_FAMILLEART3', RechDom('GCLIBFAMILLE', 'LF3', false));
{*
  //THLabel(Ecran.FindComponent('TGL_FAMILLEART2')).visible := false;
  //THValCOmboBox(Ecran.FindComponent('FAMILLETAXE2')).visible := false;
  //THLabel(Ecran.FindComponent('TGL_FAMILLEART3')).visible := false;
  //THValCOmboBox(Ecran.FindComponent('FAMILLETAXE3')).visible := false;
*}
  TTab := TTabSheet(Ecran.FindComponent('PFACTURATION'));
  if (TTab <> nil) then
  begin
    TTab.TabVisible := true;
    if GetParamSoc('SO_AFVARIABLES') then //Affaire-ONYX
    begin
      SetControlVisible('TGL_FORMULEVAR', true);
      SetControlVisible('GL_FORMULEVAR', true);
      SetControlVisible('BFORMULEVAR', true);
      BRechForm := TToolBarButton97(GetControl('BFORMULEVAR'));
      if (BRechForm <> nil) then BRechForm.OnClick := RechFormuleVar;
    end else
    begin
      SetControlVisible('TGL_FORMULEVAR', false);
      SetControlVisible('GL_FORMULEVAR', false);
      SetControlVisible('BFORMULEVAR', false);
    end;
    if (BFormuleVar) then
      TPageControl(GetControl('PAGE')).ActivePage := TTab;

    if GetParamSoc('SO_AFGENERAUTOLIG') and (AffGenerauto = 'POT') then //Génération ligne d'affaire
    begin
      SetControlVisible('GL_GENERAUTO', true);
      SetControlVisible('TGL_GENERAUTO', true);
      SetControlProperty('GL_GENERAUTO', 'Plus', ' AND (CO_CODE="POT" OR CO_CODE="ACT" OR CO_CODE="MAN") ');
    end else
    begin
      SetControlVisible('GL_GENERAUTO', false);
      SetControlVisible('TGL_GENERAUTO', false);
      SetControlText('GL_GENERAUTO', 'MAN');
    end;
  end;

  //Affaire-Revision de prix
  THEdit(GetControl('GL_FORCODE1')).onChange := FormuleOnChange;
  THEdit(GetControl('GL_FORCODE2')).onChange := FormuleOnChange;

  // elements variables et revision de prix
  if not getParamsoc('SO_AFREVISIONPRIX') then
    SetControlVisible('GB_REVISIONPRIX', False);
  if not getParamsoc('SO_AFVARIABLES') then
    SetControlVisible('GB_VARIABLES', False);
SetControlEnabled('GL_TARIFTIERS', False);  //mcd 22/10/03 alignement sur GC ou ce champ n'est pas accessible en 5.0
//uniquement en line
{*
  TCheckBox(GetCOntrol('APPLICCOEFFG')).visible := false;
  TCheckBox(GetCOntrol('GL_TENUESTOCK')).Checked := false;
  TGroupBox(GetCOntrol('GSTOCK')).visible := false;
  TCheckBox(GetCOntrol('LIVRAISONFOUR')).visible := false;
  TCheckBox(GetCOntrol('LIVRAISONFOUR')).Checked := false;
  TGroupBox(GetCOntrol('G_COMM1')).visible := false;
  TCheckBox(GetCOntrol('BLOQUETARIF')).visible := false;
  TCheckBox(GetCOntrol('FROMBORDEREAU')).visible := false;
  THLabel(GetCOntrol('TGL_PUTTCDEV')).visible := false;
  THNumEdit(GetCOntrol('GL_PUTTCDEV')).visible := false;
  THLabel(GetCOntrol('TGL_PUTTCNET')).visible := false;
  THNumEdit(GetCOntrol('GL_PUTTC')).visible := false;
*}
end;

{***************************************************************************
Auteur  ...... : AB
Créé le ...... : 14/04/2003
Modifié le ... : 14/04/2002
Description .. : Affectation et saisie de la formule des éléments variables ONYX
Suite .........: AFormuleVar
******************************************************************************}

procedure TOF_AFLignePiece.RechFormuleVar(Sender: TObject);
var StFormule: string;
begin
  StFormule := GetControltext('GL_FORMULEVAR');
  if StFormule = '' then
  begin
    if Action = taConsult then exit;
    StFormule := AGLLanceFiche('AFF', 'AFORMULEVAR_RECH', '', StFormule, '');
    AvertirTable('AFORMULEVAR');
    THValCombobox(GetControl('GL_FORMULEVAR')).reload;
    if StFormule <> '' then
      SetControltext('GL_FORMULEVAR', StFormule);
  end else
    AglLanceFiche('AFF', 'AFORMULEVAR', '', '', 'AVF_FORMULEVAR=' + StFormule + ';' + ActionToString(action));
end;

procedure TOF_AFLignePiece.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit);
begin
  // ***Fait avant le on_argument (on_argument de la tof ancêtre ) **
  if LaTob <> nil then
  begin
    if LaTob.FieldExists('GL_NATUREPIECEG') then
    begin
      if GetInfoParPiece(LaTob.GetValue('GL_NATUREPIECEG'), 'GPP_VENTEACHAT') = 'ACH' then
        bVente := False else bVente := True;
    end;
  end;
  ///******************

  Aff := THEdit(GetControl('GL_AFFAIRE'));
  Aff1 := THEdit(GetControl('GL_AFFAIRE1'));
  Aff2 := THEdit(GetControl('GL_AFFAIRE2'));
  Aff3 := THEdit(GetControl('GL_AFFAIRE3'));
  Aff4 := THEdit(GetControl('GL_AVENANT'));
  if bVente then Tiers := THEdit(GetControl('GL_TIERS'))
  else Tiers := THEdit(GetControl('TIERSAFFAIRE'));
end;

procedure TOF_AFLignePiece.OnClose;
begin

end;

procedure TOF_AFLignePiece.AffecteRessourceLigne;
var Ressource: TAFO_Ressource;
  CodeRes: string;
  CoutPR: double;
begin
  if not VH_GC.GCIfDefCEGID then
  begin
    CodeRes := GetControlText('GL_RESSOURCE');
    if (CodeRes = '') or (LaTOB = nil) then exit;
    if CodeRes = LaRessource then exit;

    Ressource := TAFO_Ressource.Create(CodeRes);
    CoutPR := Ressource.PRRessource(LaTOB.GetValue('GL_DATEPIECE'), nil, nil, LaTOB.GetValue('GL_CODEARTICLE'), LaTOB.GetValue('GL_AFFAIRE'));
    SetControlText('GL_PMRP', FloatToStr(CoutPR));
    LaRessource := CodeRes;
    Ressource.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 05/09/2002
Modifié le ... : 05/09/2002
Description .. : Gestion des adresses pour l'adresse de facturation
Suite ........ : Attention, en GA si l'affaire a une adresse de facturation,
Suite ........ : celle-ci prime sur l'adresse du CAF
Mots clefs ... : ADRESSE;GIGA
*****************************************************************}

procedure TOF_AFPiece.TrtAdresseSpecifAffaire;
var QQ: TQUERY;
  Req, caff: string;
  TobAdrClifac, Tobd: TOB;
begin

  // gm 05/09/02
  // Si le client à facturer a changé : on vérifie si l'affaire a une adresse de facturation
  // et dans ce cas , on ne fait rien
  // sinon on récupère l'adresse du client Facturé
  caff := LaTob.GetValue('GP_AFFAIRE');
  req := 'SELECT ADR_NUMEROADRESSE FROM ADRESSES where adr_refcode="' + caff + '"';
  req := req + ' AND ADR_TYPEADRESSE="AFA"';
  QQ := OpenSQL(Req, TRUE);
  if not QQ.EOF then
  begin
    Ferme(QQ);
    exit;
  end;
  Ferme(QQ);

  // mcd 28/03/03 TobAdrClifac := Tob.Create('ADRESSES',nil,-1);
  TobAdrClifac := Tob.Create('LES ADRESSES', nil, -1);
  if OkPieceAdresse then
  begin
    TOB.Create('PIECEADRESSE', TobAdrClifac, -1);
  end
  else
  begin
    TOB.Create('ADRESSES', TobAdrClifac, -1);
  end;
  GetAdrFromCode(TobAdrClifac.Detail[0], GetControlText('GP_TIERSFACTURE'));

  if (TobAdrclifac.detail.count > 0) then
  begin
    Tobd := Tobadrclifac.detail[0];
    //mcd 28/03/03 oubli de cette modif pour adresses
    if OkPieceAdresse then
    begin
      ChargementAdresseTOBEcran('FAC', 'ADR', 'GPA', Tobd);
      {*
      SetControlText('ADR_LIBELLE', Tobd.getvalue('GPA_libelle'));
      SetControlText('ADR_LIBELLE2', Tobd.getvalue('GPA_libelle2'));
      SetControlText('ADR_ADRESSE1', Tobd.getvalue('GPA_ADRESSE1'));
      SetControlText('ADR_ADRESSE2', Tobd.getvalue('GPA_ADRESSE2'));
      SetControlText('ADR_ADRESSE3', Tobd.getvalue('GPA_ADRESSE3'));
      SetControlText('ADR_CODEPOSTAL', Tobd.getvalue('GPA_CODEPOSTAL'));
      SetControlText('ADR_VILLE', Tobd.getvalue('GPA_VILLE'));
      SetControlText('ADR_PAYS', Tobd.getvalue('GPA_PAYS'));
      *}
      if LaTOB.GetValue('GP_NUMADRESSEFACT') = LaTOB.GetValue('GP_NUMADRESSELIVR') then
      begin
        LaTOB.PutValue('GP_NUMADRESSEFACT', 2);
        CP_ADR_IDENT := false;
      end;
    end
    else
    begin
      ChargementAdresseTOBEcran('FAC', 'ADR', 'ADR', Tobd);
      {*
      SetControlText('ADR_LIBELLE', Tobd.getvalue('ADR_libelle'));
      SetControlText('ADR_LIBELLE2', Tobd.getvalue('ADR_libelle2'));
      SetControlText('ADR_ADRESSE1', Tobd.getvalue('ADR_ADRESSE1'));
      SetControlText('ADR_ADRESSE2', Tobd.getvalue('ADR_ADRESSE2'));
      SetControlText('ADR_ADRESSE3', Tobd.getvalue('ADR_ADRESSE3'));
      SetControlText('ADR_CODEPOSTAL', Tobd.getvalue('ADR_CODEPOSTAL'));
      SetControlText('ADR_VILLE', Tobd.getvalue('ADR_VILLE'));
      SetControlText('ADR_PAYS', Tobd.getvalue('ADR_PAYS'));
      *}
      if LaTOB.GetValue('GP_NUMADRESSEFACT') = LaTOB.GetValue('GP_NUMADRESSELIVR') then
      begin
        LaTOB.PutValue('GP_NUMADRESSEFACT', -2);
        CP_ADR_IDENT := false;
      end;
    end;
  end;
  TobAdrClifac.free;
end;

// ********************Fonctions du script des lignes  ************************

procedure AGLAffecteRessourceLigne(parms: array of variant; nb: integer);
var F: TForm;
  LaTof: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then LaTof := TFVierge(F).Latof else laTof := nil;
  if (LaTof is TOF_AFLignePiece) then TOF_AFLignePiece(LaTof).AffecteRessourceLigne;
end;

procedure AGLAdrTiersPourPiece(parms: array of variant; nb: integer);
var F: TForm;
  LaTof: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then LaTof := TFVierge(F).Latof else laTof := nil;
  if (LaTof is TOF_AFPiece) then TOF_AFPiece(LaTof).SelectAdresseTiers();
end;

{$IFDEF BTP}
procedure AGLSetAdrLivraisonModifie(parms: array of variant; nb: integer);
var F: TForm;
  LaTof: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then LaTof := TFVierge(F).Latof else laTof := nil;
  if (LaTof is TOF_AFPiece) then TOF_AFPiece(LaTof).SetAdrLivraisonModifie();
end;
{$ENDIF}

procedure AFLanceFiche_ComplementPiece(Argument: string);
begin
//uniquement en line
//  AGLLanceFiche('BTP', 'BTCOMPLPIECE_S1', '', '', Argument);
	    AGLLanceFiche('BTP', 'BTCOMPLPIECE', '', '', Argument);
//  AGLLanceFiche('AFF', 'AFCOMPLPIECE', '', '', Argument);
//
end;

procedure AFLanceFiche_ComplementLigne(Argument: string);
begin
//uniquement en line
//	AGLLanceFiche('BTP', 'BTCOMPLLIGNE_S1', '', '', Argument);
  AGLLanceFiche('BTP', 'BTCOMPLLIGNE', '', '', Argument);
end;

procedure AFLanceFiche_Mul_Piece(Range, Argument: string);
begin
  AGLLanceFiche('AFF', 'AFPIECE_MUL', range, '', Argument);
end;

procedure AFLanceFiche_Mul_Edition_Piece(Argument: string);
begin
  //AGLLanceFiche('AFF', 'AFPIECEEDIT_MUL', '', '', Argument);
  AGLLanceFiche ('GC','GCEDITDOCDIFF_MUL','','',Argument);
end;

procedure AFLanceFiche_Mul_Piece_Provisoire(Argument: string);
Var Statut 	: string;
		Critere	: string;
    champ		: string;
    Valeur	: String;
    SArg		: string;
    X				: Integer;
    Vivante : string;
begin
    SArg := Argument;
    Vivante := 'X';
	  Critere:=(Trim(ReadTokenSt(SArg)));
	  While (Critere <>'') do
    BEGIN
       if Critere<>'' then
       BEGIN
       X:=pos(':',Critere);
       if x<>0 then
          begin
          Champ:=copy(Critere,1,X-1);
          Valeur:=Copy (Critere,X+1,length(Critere)-X);
          end
       else
          Begin
       		X:=pos('=',Critere);
       		if x<>0 then
          	 begin
          	 Champ:=copy(Critere,1,X-1);
          	 Valeur:=Copy (Critere,X+1,length(Critere)-X);
          	 end
       		else
	        	 Champ:= Critere;
          end;
       If Champ = 'STATUT' then Statut := valeur else
       if champ = 'VIVANTE' then Vivante := valeur;
       END;
    	 Critere:=(Trim(ReadTokenSt(SArg)));
    END;

{$IFDEF BTP}
    if Statut = 'APP' then
 	    AGLLanceFiche('BTP', 'BTPIECEPRO_MUL', 'AFF_AFFAIRE0=W;GP_VIVANTE='+Vivante, '', Argument)
    else if Statut = 'GRP' then
 	    AGLLanceFiche('BTP', 'BTPIECEPRO_MUL', 'GP_VIVANTE='+Vivante, '', Argument)
    else if Statut = 'INT' then
    begin
	    if vivante <> '' then AGLLanceFiche('BTP', 'BTPIECEPRO_MUL', 'GP_VIVANTE='+Vivante, '', Argument)
      else AGLLanceFiche('BTP', 'BTPIECEPRO_MUL', '', '', Argument);
    end else if Statut = 'AFF' then
      AGLLanceFiche('BTP', 'BTPIECEPRO_MUL', 'AFF_AFFAIRE0=A;GP_VIVANTE='+Vivante, '', Argument);
{$ELSE}
  	  AGLLanceFiche('BTP', 'BTPIECEPRO_MUL', '', '', Arg);
{$ENDIF}

end;

function AFLanceFiche_Date_ValidPiecePRO(Argument: string): variant;
begin
  result := AGLLanceFiche('AFF', 'AFPIECEPRO_VAL', '', '', Argument);
end;


procedure AFLanceFiche_DuplicPiece(Range, Argument: string);
begin
  AGLLanceFiche('AFF', 'AFDUPLICPIECE_MUL', range, '', Argument);
end;

procedure AFLanceFiche_PiecePourUneAffaire(Argument: string);
begin
  AGLLanceFiche('AFF', 'AFPIECEF_MUL', '', '', Argument);
end;

{$IFDEF BTP}
procedure TOF_AFPiece.GP_TIERSLIVRE_OnEnter(Sender : TObject);
begin
TiersLivre := THEdit(GetControl('GP_TIERSLIVRE')).Text;
end;

procedure TOF_AFPiece.GP_TIERSLIVRE_OnExit(Sender : TObject);
begin
	if TiersLivre <> THEdit(GetControl('GP_TIERSLIVRE')).Text then
  begin
    //SetControlText('GP_NADRESSELIV', IntToStr(GetNumAdresseFromTiers('', GetControlText('GP_TIERSLIVRE'), taLivr)));
    SetControlText('GP_NADRESSELIV', GetNumAdresseFromTiers('', GetControlText('GP_TIERSLIVRE'), taLivr));
    SetControlProperty('ADL_NUMEROCONTACT', 'Plus', GetCodeAuxi('ADL'));
    RechercheAdresseLivraison;
    if ExisteAdresseLiv (LaTOB.GetValue('GP_TIERSLIVRE')) then
    begin
       THEdit(GetControl('LNUMLIVR')).visible := true;
       THEdit(GetControl('GP_NADRESSELIV')).visible := true;
    end else
    begin
       THEdit(GetControl('LNUMLIVR')).visible := true;
       THEdit(GetControl('GP_NADRESSELIV')).visible := true;
    end;
  end;
end;

procedure TOF_AFPiece.GP_TIERSFACExit (Sender : Tobject);
BEGIN
  if TiersFacture <> THEdit(GetControl('GP_TIERSFACTURE')).Text then
  begin
    SetControlText('ADR_TELEPHONEFAC' , '');
    SetControlText('ADR_CONTACT','');
    RechercheAdressefact;
		SetModeRegle;
    Tiersfacture := THEdit(GetControl('GP_TIERSFACTURE')).Text;
  end;
END;

procedure TOF_AFPiece.GP_NADRESSELIV_OnEnter(Sender : TObject);
begin
NADRESSELIV := THEdit(GetControl('GP_NADRESSELIV')).Text;
end;

procedure TOF_AFPiece.GP_NADRESSELIV_OnExit(Sender : TObject);
begin
	if NADRESSELIV <> THEdit(GetControl('GP_NADRESSELIV')).Text then
  begin
    SetControlProperty('ADL_NUMEROCONTACT', 'Plus', GetCodeAuxi('ADL'));
    RechercheAdresseLivraison;
  end;
end;

procedure TOF_AFPiece.GP_TIERSLIVRE_OnElipsisClick(Sender : TObject);
begin
  if THEdit(GetCOntrol('GP_VENTEACHAT')).Text = 'VEN' then
  begin
    DispatchRecherche(THEdit(GetControl('GP_TIERSLIVRE')),2,'T_NATUREAUXI="CLI"','','');
  end
  else if THEdit(GetCOntrol('GP_VENTEACHAT')).Text = 'ACH' then
  begin
    DispatchRecherche(THEdit(GetControl('GP_TIERSLIVRE')),2,'T_NATUREAUXI="FOU"','','');
  end
  else
  begin
    DispatchRecherche(THEdit(GetControl('GP_TIERSLIVRE')),2,'','','');
  end;
end;

procedure TOF_AFPiece.GP_TIERSFAC_OnElipsisClick(Sender : TObject);
Begin

  if THEdit(GetCOntrol('GP_VENTEACHAT')).Text = 'VEN' then
  begin
    DispatchRecherche(THEdit(GetControl('GP_TIERSFACTURE')),2,'T_NATUREAUXI="CLI"','','');
  end
  else if THEdit(GetCOntrol('GP_VENTEACHAT')).Text = 'ACH' then
  begin
    DispatchRecherche(THEdit(GetControl('GP_TIERSFACTURE')),2,'T_NATUREAUXI="FOU"','','');
  end
  else
  begin
    DispatchRecherche(THEdit(GetControl('GP_TIERSFACTURE')),2,'','','');
  end;

end;

procedure TOF_AFPiece.SetAdrLivraisonModifie;
begin
  CP_ADR_IDENT := false;
end;

procedure TOF_AFPiece.SelectionAdresseLivTiers (Sender : Tobject);
var CodeTiers,CodeAuxi,Range,params,Valeur : string;
	  BadrOk : boolean;
    TOBAdresse : TOB;
begin
	TOBAdresse := TOB.Create ('ADRESSES',nil,-1);
  TRY
    CodeTiers := THEdit(GetControl('GP_TIERSLIVRE')).Text;
    CodeAuxi:=TiersAuxiliaire(CodeTiers,false);
    Range :='ADR_REFCODE=' + CodeTiers + ';ADR_LIVR=X';
    Params := 'TYPEADRESSE=TIE;NATUREAUXI=CLI'
            +';YTC_TIERSLIVRE=' + CodeTiers
            +';CLI=' + CodeAuxi
            +';TIERS=' + CodeTiers
            +';ACTION=MODIFICATION';

    Valeur := AglLanceFiche('GC', 'GCADRESSES', Range, '', Params);
    if Valeur <> '' then
    begin
      bAdrOk := THEdit(GetControl('GP_TIERSLIVRE')).Focused;
      TOBAdresse.PutValue('ADR_NUMEROADRESSE', StrToInt (Valeur));
      TOBAdresse.PutValue('ADR_TYPEADRESSE', 'TIE');
      TOBAdresse.LoadDB();
      PutEcranTobPieceAdresseLivraison( TobAdresse,True );
      THEdit(GetControl('GP_NADRESSELIV')).text := TOBAdresse.getValue('ADR_NADRESSE');
      CP_ADR_IDENT := false;
    end;
  FINALLY
  	TOBAdresse.free;
  END;
end;

procedure TOF_AFPiece.SelectionAdresseLivAffaire (Sender : Tobject);
var select,Valeur,Req : string;
		QQ : TQuery;
    TOBTemp : TOB;
begin
  Select := 'ADR_REFCODE = "' + LaTOB.GetValue('GP_AFFAIRE')+'" AND ADR_TYPEADRESSE = "INT"';
  Valeur := AGLLanceFiche('GC','GCADRESSES_MUL', '', '', Select);
  if Valeur <> '' then
  begin
    TobTemp := TOB.Create('ADRESSES', nil, -1);
    Req := 'SELECT * FROM ADRESSES WHERE ADR_NUMEROADRESSE='+Valeur+' AND ADR_TYPEADRESSE="INT"';
    QQ := OpenSql (Req,True);
    if not QQ.eof then
    begin
      TOBTemp.selectDb ('',QQ,true);
      PutEcranTobPieceAdresseLivraison( TobTemp,True );
    	CP_ADR_IDENT := false;
    end;
    ferme (QQ);
    TOBTemp.free;
  end;
end;


//Adresse de Facturation
procedure TOF_AFPiece.PutEcranTobPieceAdressefac( TobAdresseFact : tob; ForceAdresse : boolean=false);
var
  sIncoterm : string;
begin
  if OkPieceAdresse and (not ForceAdresse) then
  begin
    ChargementAdresseTobEcran('FAC', 'ADR', 'GPA', TobAdresseFact);
    {*
    if TobAdresseFact.GetValue('GPA_LIBELLE')<>'' then SetControlText('ADR_LIBELLE',TobAdresseFact.GetValue('GPA_LIBELLE'));
    SetControlText('ADR_JURIDIQUE', TobAdresseFact.GetValue('GPA_JURIDIQUE'));
    SetControlText('ADR_LIBELLE2' , TobAdresseFact.GetValue('GPA_LIBELLE2')    );
    SetControlText('ADR_ADRESSE1'  ,TobAdresseFact.GetValue('GPA_ADRESSE1')    );
    SetControlText('ADR_ADRESSE2'  ,TobAdresseFact.GetValue('GPA_ADRESSE2')    );
    SetControlText('ADR_ADRESSE3'  ,TobAdresseFact.GetValue('GPA_ADRESSE3')    );
    SetControlText('ADR_CODEPOSTAL',TobAdresseFact.GetValue('GPA_CODEPOSTAL')  );
    SetControlText('ADR_VILLE'     ,TobAdresseFact.GetValue('GPA_VILLE')       );
    SetControlText('ADR_PAYS'      ,TobAdresseFact.GetValue('GPA_PAYS')        );
    SetControlText('ADR_NUMEROCONTACT',TobAdresseFact.GetValue('GPA_NUMEROCONTACT'));
    SetInfosContact('FAC',TobAdresseFact.GetValue('GPA_NUMEROCONTACT'),TobAdresseFact.GetValue('GPA_TELEPHONE'));
    *}
  end
  else
  begin
    ChargementAdresseTobEcran(' FAC', 'ADR', 'ADR', TobAdresseFact);
    {*
    if TobAdresseFact.GetValue('ADR_LIBELLE')<>'' then SetControlText('ADR_LIBELLE',TobAdresseFact.GetValue('ADR_LIBELLE'));
    SetControlText('ADR_JURIDIQUE', TobAdresseFact.GetValue('ADR_JURIDIQUE'));
    SetControlText('ADR_LIBELLE2' , TobAdresseFact.GetValue('ADR_LIBELLE2')    );
    SetControlText('ADR_ADRESSE1'  ,TobAdresseFact.GetValue('ADR_ADRESSE1')    );
    SetControlText('ADR_ADRESSE2'  ,TobAdresseFact.GetValue('ADR_ADRESSE2')    );
    SetControlText('ADR_ADRESSE3'  ,TobAdresseFact.GetValue('ADR_ADRESSE3')    );
    SetControlText('ADR_CODEPOSTAL',TobAdresseFact.GetValue('ADR_CODEPOSTAL')  );
    SetControlText('ADR_VILLE'     ,TobAdresseFact.GetValue('ADR_VILLE')       );
    SetControlText('ADR_PAYS'      ,TobAdresseFact.GetValue('ADR_PAYS')        );
    SetControlText('ADR_NUMEROCONTACT',TobAdresseFact.GetValue('ADR_NUMEROCONTACT'));
    *}
  end;
end;

//Adresse de Livraison
procedure TOF_AFPiece.PutEcranTobPieceAdresseLivraison( TobAdresseLivraison : tob; ForceAdresse : boolean=false);
var
  sIncoterm : string;
begin
	SetControlText('ADL_TELEPHONE' ,'');
	SetControlText('ADL_CONTACT'   ,'');
  if OkPieceAdresse and (not ForceAdresse) then
  begin
    ChargementAdresseTobEcran('LIV', 'ADL', 'GPA', TobAdresseLivraison);
    {*
    if TobAdresseLivraison.GetValue('GPA_LIBELLE')<>'' then SetControlText('ADL_LIBELLE',TobAdresseLivraison.GetValue('GPA_LIBELLE'));
    SetControlText('ADL_JURIDIQUE', TobAdresseLivraison.GetValue('GPA_JURIDIQUE'));
    SetControlText('ADL_LIBELLE2' , TobAdresseLivraison.GetValue('GPA_LIBELLE2')    );
    SetControlText('ADL_ADRESSE1'  ,TobAdresseLivraison.GetValue('GPA_ADRESSE1')    );
    SetControlText('ADL_ADRESSE2'  ,TobAdresseLivraison.GetValue('GPA_ADRESSE2')    );
    SetControlText('ADL_ADRESSE3'  ,TobAdresseLivraison.GetValue('GPA_ADRESSE3')    );
    SetControlText('ADL_CODEPOSTAL',TobAdresseLivraison.GetValue('GPA_CODEPOSTAL')  );
    SetControlText('ADL_VILLE'     ,TobAdresseLivraison.GetValue('GPA_VILLE')       );
    SetControlText('ADL_PAYS'      ,TobAdresseLivraison.GetValue('GPA_PAYS')        );
    SetControlText('ADL_NUMEROCONTACT'   ,TobAdresseLivraison.GetValue('GPA_NUMEROCONTACT'));

    SetInfosContact('LIV',TobAdresseLivraison.GetValue('GPA_NUMEROCONTACT'),TobAdresseLivraison.GetValue('GPA_TELEPHONE'));
    *}
  end
  else
  begin
    ChargementAdresseTobEcran('LIV', 'ADL', 'ADR', TobAdresseLivraison);
    {*
    if TobAdresseLivraison.GetValue('ADR_LIBELLE')<>'' then SetControlText('ADL_LIBELLE',TobAdresseLivraison.GetValue('ADR_LIBELLE'));
    SetControlText('ADL_JURIDIQUE', TobAdresseLivraison.GetValue('ADR_JURIDIQUE'));
    SetControlText('ADL_LIBELLE2' , TobAdresseLivraison.GetValue('ADR_LIBELLE2')    );
    SetControlText('ADL_ADRESSE1'  ,TobAdresseLivraison.GetValue('ADR_ADRESSE1')    );
    SetControlText('ADL_ADRESSE2'  ,TobAdresseLivraison.GetValue('ADR_ADRESSE2')    );
    SetControlText('ADL_ADRESSE3'  ,TobAdresseLivraison.GetValue('ADR_ADRESSE3')    );
    SetControlText('ADL_CODEPOSTAL',TobAdresseLivraison.GetValue('ADR_CODEPOSTAL')  );
    SetControlText('ADL_VILLE'     ,TobAdresseLivraison.GetValue('ADR_VILLE')       );
    SetControlText('ADL_PAYS'      ,TobAdresseLivraison.GetValue('ADR_PAYS')        );
    SetControlText('ADL_NUMEROCONTACT'   ,TobAdresseLivraison.GetValue('ADR_NUMEROCONTACT'));
    SetControlText('ADRESSELIVAFF',TobAdresseLivraison.GetValue('ADR_NADRESSE'));
    SetInfosContact('LIV',TobAdresseLivraison.GetValue('ADR_NUMEROCONTACT'),TobAdresseLivraison.GetValue('ADR_TELEPHONE'));
    *}
  end;
end;

procedure TOF_AFPiece.SetInfosContact(TypeAdr : string; NumeroContact : integer; Telephone : string);
var
  Req : string;
  QQ : TQuery;
  TOBContact : TOB;
begin
	TOBContact := TOB.create ('CONTACT',nil,-1);
  if NumeroContact <> 0 then
  begin
  	if TypeAdr = 'LIV' then
    begin
      if GetContact (TOBContact,GetControltext('GP_TIERSLIVRE'),NumeroContact) then
      begin
        SetControlText('ADL_TELEPHONE' , TOBContact.GetValue('C_TELEPHONE'));
        SetControlText('ADL_CONTACT'   ,rechDom('TTCIVILITE',TOBCOntact.getValue('C_CIVILITE'),false) + ' ' +TOBContact.GetValue('C_PRENOM') + ' '+TOBContact.GetValue('C_NOM') );
      end;
    end else
    begin
      if GetContact (TOBContact,GetControltext('GP_TIERSFACTURE'),NumeroContact) then
      begin
        SetControlText('ADR_TELEPHONEFAC' , TOBContact.GetValue('C_TELEPHONE'));
        SetControlText('ADR_CONTACT'   ,rechDom('TTCIVILITE',TOBCOntact.getValue('C_CIVILITE'),false) + ' ' +TOBContact.GetValue('C_PRENOM') + ' '+TOBContact.GetValue('C_NOM') );
      end;
    end;
  end else
  begin
  	if TypeAdr = 'LIV' then
    begin
  		SetControlText('ADL_TELEPHONE'   ,Telephone);
    end else
    begin
  		SetControlText('ADR_TELEPHONE'   ,Telephone);
    end;
  end;
  TOBContact.free;
end;

function TOF_AFPiece.NombreAdressesLiv(Affaire: string): integer;
var Req : String;
		QQ : Tquery;
begin
	result := 0;
  Req := 'SELECT COUNT(*) FROM ADRESSES WHERE '+
  			 '(ADR_REFCODE = "' + LaTOB.GetValue('GP_AFFAIRE')+'" AND ADR_TYPEADRESSE = "INT")';
  QQ := OpenSql (REQ,True);
  if not QQ.eof then result := QQ.Fields[0].asInteger;
  ferme (QQ);
end;


Function TOF_AFPiece.GetCodeAuxi(Prefixe: string) : string;
var
  Suffixe: string;
begin
  if Prefixe = 'ADL' then
    Suffixe := 'TIERSLIVRE'
  else Suffixe := 'TIERSFACTURE';

  Result :=TiersAuxiliaire(GetControlText('GP_' + Suffixe),false);
end;

procedure TOF_AFPiece.RechercheAdressefact;
var
   TobAdresse : tob;
begin
   //TobAdresse  est une Tob liée à la table ADRESSES
   //TobAdresses est une Tob liée à la table PIECEADRESSE (ou à la table ADRESSES : ancien fonctionnement)
   TobAdresse := Tob.create('_ADRESSES_',nil,-1);
   Try
      GetTobAdresseFromAdresses(TobAdresse, 'TIE', GetControlText('GP_TIERSFACTURE'), 0, 'Facturation');
      //Si on trouve une adresse
      if (TobAdresse.Detail.count<>0) then
      begin
      	GetTobPieceAdresseFromTobAdresses( TobAdresses.Detail[1], TobAdresse.Detail[0]);
      end else // sinon on reprend l'adresse du tiers
      begin
      	GetAdrFromCode(TOBAdresses.Detail[1], GetControlText('GP_TIERSFACTURE')) ;
      end;
      PutEcranTobPieceAdresseFac( TobAdresses.Detail[1] ,false);
   Finally
      TobAdresse.Free;
   end;
end;

procedure TOF_AFPiece.RechercheAdresseLivraison;
var
   TobAdresse : tob;
begin
   //TobAdresse  est une Tob liée à la table ADRESSES
   //TobAdresses est une Tob liée à la table PIECEADRESSE (ou à la table ADRESSES : ancien fonctionnement)
   TobAdresse := Tob.create('_ADRESSES_',nil,-1);
   Try
      GetTobAdresseFromAdresses(TobAdresse, 'TIE', GetControlText('GP_TIERSLIVRE'), StrToInt(GetControlText('GP_NADRESSELIV')), 'Livraison');
      //Si on trouve une adresse
      if (TobAdresse.Detail.count<>0) then
      begin
      	GetTobPieceAdresseFromTobAdresses( TobAdresses.Detail[0], TobAdresse.Detail[0]);
      end
      else // sinon on reprend l'adresse du tiers
      begin
      	GetAdrFromCode(TOBAdresses.Detail[0], GetControlText('GP_TIERSLIVRE')) ;
      end;
      PutEcranTobPieceAdresseLivraison( TobAdresses.Detail[0] );
   Finally
      TobAdresse.Free;
   end;
end;

Function TOF_AFPiece.ExisteAdresseLiv ( sTiersLivre : string) : boolean;
var Srequete : string;
begin
	result := false;
  sRequete := 'SELECT * FROM ADRESSES WHERE (ADR_TYPEADRESSE="TIE") and (ADR_REFCODE="' + sTiersLivre + '")';
  sRequete := sRequete + ' and (ADR_LIVR="X")';

  if ExisteSQL(sRequete) then result := true;
end;


procedure TOF_AFPiece.SetComponents;
var Indice : integer;
		UneTaxe : TOB;
begin
	for Indice := 0 to VH_GC.TOBParamTaxe.detail.count -1 do
  begin
    Unetaxe  := VH_GC.TOBParamTaxe.detail[Indice];
    // non paramétré
    THLabel (GetControl('TGB_FAMILLETAXE'+inttostr(Indice+1))).Visible := false;
    THValComboBox (GetControl('FAMILLETAXE'+inttostr(Indice+1))).Visible := false;
    if UneTaxe.getValue('BPT_TYPETAXE') <> '' then
    begin
      if UneTaxe.getValue('BPT_TYPETAXE') = 'TVA' Then
      begin
      //
      	if (Indice > 0) and ((TOBRepart = nil)  or (TOBRepart.detail.count = 0 )) then continue;
      //
      	THValComboBox (GetControl('FAMILLETAXE'+inttostr(Indice+1))).DataType := 'TTTVA';
      	THValComboBox (GetControl('FAMILLETAXE'+inttostr(Indice+1))).Visible := True;
      	THLabel (GetControl('TGB_FAMILLETAXE'+inttostr(Indice+1))).Caption := UneTaxe.getValue('BPT_LIBELLE');
      	THLabel (GetControl('TGB_FAMILLETAXE'+inttostr(Indice+1))).visible := true;
        if (TOBRepart <> nil)  and (TOBRepart.detail.count > 0 ) then
        begin
          THNumEdit(GetCOntrol('MILLIEME'+InttoStr(indice+1))).Value :=  GetMillieme('TX'+IntToStr(Indice+1),TOBRepart);
          if THNumEdit(GetCOntrol('MILLIEME'+InttoStr(indice+1))).Value > 0 then
          begin
            THNumEdit(GetCOntrol('MILLIEME'+InttoStr(indice+1))).Visible := true;
            THLabel(GetCOntrol('MILLI'+InttoStr(indice+1))).Visible := true;
          end;
        end;
      end else
      begin if UneTaxe.getValue('BPT_TYPETAXE') = 'TPF' then
//Non en line
      	THValComboBox (GetControl('FAMILLETAXE'+inttostr(Indice+1))).DataType := 'TTTPF';
      	THValComboBox (GetControl('FAMILLETAXE'+inttostr(Indice+1))).Visible := True;
      	THLabel (GetControl('TGB_FAMILLETAXE'+inttostr(Indice+1))).Caption := UneTaxe.getValue('BPT_LIBELLE');
      	THLabel (GetControl('TGB_FAMILLETAXE'+inttostr(Indice+1))).visible := true;
      end;
      if (TOBRepart <> nil) and (TOBRepart.detail.count > 0) then
         THValComboBox (GetControl('FAMILLETAXE'+inttostr(Indice+1))).Enabled := false;
    end;
  end;
//uniquement en line
//  SetControlProperty('MODEFACTURATION', 'Plus', 'BTP" AND CO_CODE<>"DAC');

end;

{$ENDIF}

procedure TOF_AFPiece.NumeroContactLivClick(Sender: Tobject);
var retour,CodeAuxi,CodeTiers,Natureauxi,Action : string;
begin
  CodeTiers := THEdit(GetControl('GP_TIERSLIVRE')).Text;
	CodeAuxi:=TiersAuxiliaire(CodeTiers,false);
  if THEdit(GetCOntrol('GP_VENTEACHAT')).Text = 'VEN' then
  begin
    NATUREAUXI:='CLI';
  end
  else if THEdit(GetCOntrol('GP_VENTEACHAT')).Text = 'ACH' then
  begin
    NATUREAUXI:='FOU';
  end;
  if TFFiche(ecran).FTypeAction <> taConsult then
  begin
  	Action := 'ACTION=MODIFICATION';
  end else
  begin
  	Action := 'ACTION=CONSULTATION';
  end;

  retour := AglLanceFIche('YY','YYCONTACT','T;'+CodeAuxi,'',Action+ ';TYPE=T;'+'TYPE2='+Natureauxi+';TITRE='+GetControlText('ADL_LIBELLE')+';TIERS='+CodeTiers+';ALLCONTACT' );
  if retour <> '' then
  begin
  	THEdit(GetCOntrol('ADL_NUMEROCONTACT')).Text := ReadTokenSt(Retour);
    SetInfosContact ('LIV',StrToInt(THEdit(GetCOntrol('ADL_NUMEROCONTACT')).text),'');
  end;
end;

procedure TOF_AFPiece.NumeroContactClick(Sender: Tobject);
var retour,CodeAuxi,CodeTiers,Natureauxi,Action : string;
begin
  CodeTiers := THEdit(GetControl('GP_TIERSFACTURE')).Text;
  if CodeTiers = '' then CodeTiers := THEdit(GetControl('GP_TIERS')).Text;
	CodeAuxi:=TiersAuxiliaire(CodeTiers,false);
  if THEdit(GetCOntrol('GP_VENTEACHAT')).Text = 'VEN' then
  begin
    NATUREAUXI:='CLI';
  end
  else if THEdit(GetCOntrol('GP_VENTEACHAT')).Text = 'ACH' then
  begin
    NATUREAUXI:='FOU';
  end;
  if TFFiche(ecran).FTypeAction <> taConsult then
  begin
  	Action := 'ACTION=MODIFICATION';
  end else
  begin
  	Action := 'ACTION=CONSULTATION';
  end;

  retour := AglLanceFIche('YY','YYCONTACT','T;'+CodeAuxi,'',Action+ ';TYPE=T;'+'TYPE2='+Natureauxi+';TITRE='+GetControlText('ADR_LIBELLE')+';TIERS='+CodeTiers+';ALLCONTACT' );
  if retour <> '' then
  begin
  	THEdit(GetCOntrol('ADR_NUMEROCONTACT')).Text := ReadTokenSt(Retour);
    SetInfosContact ('FAC',StrToInt(THEdit(GetCOntrol('ADR_NUMEROCONTACT')).text),'');
  end;
end;

procedure TOF_AFPiece.ChangeAdresseLivr(Sender: TObject);
var NOmchamp : string;
begin
  NomChamp := THEDit(Sender).Name;

  if NomChamp = 'ADR_JURIDIQUE' then THEDit(GEtcontrol('ADL_JURIQUE')).Text := THEDit(Sender).text else
  if NomChamp = 'ADR_LIBELLE' then THEDit(GEtcontrol('ADL_LIBELLE')).Text := THEDit(Sender).text else
  if NomChamp = 'ADR_LIBELLE2' then THEDit(GEtcontrol('ADL_LIBELLE2')).Text := THEDit(Sender).text else
  if NomChamp = 'ADR_ADRESSE1' then THEDit(GEtcontrol('ADL_ADRESSE1')).Text := THEDit(Sender).text else
  if NomChamp = 'ADR_ADRESSE2' then THEDit(GEtcontrol('ADL_ADRESSE2')).Text := THEDit(Sender).text else
  if NomChamp = 'ADR_ADRESSE3' then THEDit(GEtcontrol('ADL_ADRESSE3')).Text := THEDit(Sender).text else
  if NomChamp = 'ADR_CODEPOSTAL' then THEDit(GEtcontrol('ADL_CODEPOSTAL')).Text := THEDit(Sender).text else
  if NomChamp = 'ADR_VILLE' then THEDit(GEtcontrol('ADL_VILLE')).Text := THEDit(Sender).text else
  if NomChamp = 'ADR_PAYS' then THEDit(GEtcontrol('ADL_PAYS')).Text := THEDit(Sender).text;
end;

//uniquement en line
{*
procedure TOF_AFPiece.SetShowForLine;
var TC : TCheckBox;
    TV : THValCOmboBox;
    TL : THLabel;
begin

  TC := TCheckBox(GetCOntrol('GP_FACTUREHT'));
  if TC <> nil then TC.visible := false;

  TV := THValComboBox(GEtCOntrol('GP_TARIFTIERS'));
  IF TV <> nil Then TV.visible := false;

  TL := THLabel(GEtCOntrol('TGP_TARIFTIERS'));
  IF TL <> nil Then TL.visible := false;

  TTabSheet(GetControl('INFORMATIONS')).TabVisible :=false;

  TTabSheet(GetControl('TAXES')).TabVisible :=false;

  TTabSheet(GetControl('COMMENTAIRE')).TabVisible :=false;
	TTabSheet(GetControl('OPTIONS')).TabVisible :=false;

  THEdit(GetControl('GP_APPORTEUR')).Visible :=false;
  THEdit(GetControl('GP_REPRESENTANT')).Visible :=false;
  THEdit(GetControl('ADL_NUMEROCONTACT')).visible := false;

  THLabel(GetControl('TGP_APPORTEUR')).Visible :=false;
  THLabel(GetControl('TGP_REPRESENTANT')).Visible :=false;

  TCheckBox(GetControl('LIVRECEPFOUR')).Visible :=false;
  TCheckBox(GetControl('LIVRECEPFOUR')).checked := false;
  TCheckBox(GetControl('GP_APPLICFGST')).checked := True;
  TCheckBox(GetControl('GP_APPLICFGST')).visible := false;
  TCheckBox(GetControl('_GERE_EN_STOCK')).visible := false;

  if LaTOB.GetValue('GP_NATUREPIECEG') <> 'DBT' Then TTabSheet(GetCOntrol('OPTIONS')).tabVisible := false;

end;
*}

procedure TOF_AFPiece.SetModeRegle;
var Req : string;
		QQ : TQuery;
begin
  Req := 'SELECT T_MODEREGLE FROM TIERS WHERE T_TIERS="' + GetControlText('GP_TIERSFACTURE') + '"';
  QQ := OpenSql (REQ,True);
  if not QQ.eof then
  begin
    SetControlText('GP_MODEREGLE',QQ.Fields[0].asString);
    LaTob.PutValue('GP_MODEREGLE',QQ.Fields[0].asString);
  end;
  ferme (QQ);
end;

initialization
  registerclasses([TOF_AFPiece, TOF_AFLignePiece]);
  RegisterAglProc('AffecteRessourceLigne', True, 1, AGLAffecteRessourceLigne);
  RegisterAglProc('AFAdrTiersPourPiece', True, 1, AGLAdrTiersPourPiece);
{$IFDEF BTP}
  RegisterAglProc('SetAdrLivraisonModifie', True, 0, AGLSetAdrLivraisonModifie);
{$ENDIF}
end.
