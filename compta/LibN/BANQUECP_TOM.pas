{ Unité : Source TOM de la TABLE : BANQUECP
--------------------------------------------------------------------------------------
    Version    |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
                18/03/03   BPY  Création de l'unité
 7.09.001.007   02/03/07   JP   FQ 19724 : Test sur l'utilisation des comptes dans la Tréso étendu à la compta
 8.00.001.014   11/05/07   JP   FQ TRESO 10462 : à l'enregistrement BQ_NODOSSIER est vidé quand
                                on n'est pas en multi dossiers
 8.00.001.018   07/06/07   JP   FQ 19591 : gestion des modèles de bordereaux
 8.00.001.024   05/07/07   JP   FQ 20319 : Calcul de l'IBAN avnt le OnUpdateRecord 
--------------------------------------------------------------------------------------}
Unit BANQUECP_TOM;

//================================================================================
// Interface
//================================================================================
Interface

uses {$IFDEF VER150} variants,{$ENDIF}
    StdCtrls,
    Controls,
    Classes,
{$IFDEF EAGLCLIENT}
    MaineAGL,
    eFiche,
    UTob,
{$ELSE}
    FE_Main,
    Fiche,
    db,
    HDB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
    sysutils,
    ComCtrls,
    HCtrls,
    Buttons,
    Ent1,
    HEnt1,
    HMsgBox,
    UTOM,
    LookUp,
    utilPGI,
    CPCODEPOSTAL_TOF,
    CPREGION_TOF,
    AglInit {JP 23/09/03 = > ActionToString}
    , UTomBanque
    , TomAgence
    , UlibIdentBancaire
    , FichComm
  ,UentCommun    
    ;

//==================================================
// Externe
//==================================================
procedure FicheBanqueCP(CpteGeneral : string ; Comment : TActionFiche ; QuellePage : Integer; Dossier : string = '');
{JP 25/10/06 : Avec le DispatchTT, c'est le code que l'on reçoit ...}
procedure FicheBanqueCPCode(Code : string ; Comment : TActionFiche);

//==================================================
// Definition de class
//==================================================
Type
    TOM_BANQUECP = Class (TOM)
        procedure OnNewRecord              ; override ;
        procedure OnDeleteRecord           ; override ;
        procedure OnUpdateRecord           ; override ;
        procedure OnAfterUpdateRecord      ; override ;
        procedure OnLoadRecord             ; override ;
        procedure OnChangeField(F: TField) ; override ;
        procedure OnArgument(S: String)    ; override ;
        procedure OnClose                  ; override ;
        procedure OnCancelRecord           ; override ;
    private
        procedure InitBordereau; {JP 07/06/07 : FQ 19591}
        // Reinitialisation de l'affichage :
        procedure AfficheEcran;
        procedure AfficheSpecifES;  
        // evenement
        procedure BDleteClick(Sender : TObject); {FQ 15585}
        procedure OnClickBMCbanq(Sender: TObject);
        procedure OnChangeBQ_PAYS(Sender: TObject);
        procedure OnChangeBQ_TYPEPAYS(Sender: TObject);
        procedure Verif_CODEIBAN;
        procedure OnKeyPressBQ_NUMEROCOMPTE(Sender: TObject; var Key: Char);
        procedure OnChangeBQ_DEVISE(Sender: TObject);
        procedure OnDblClickBQ_CODEPOSTAL(Sender: TObject);
        procedure OnDblClickBQ_DIVTERRIT(Sender: TObject);
        procedure OnDblClickBQ_VILLE(Sender: TObject);
        // gestionn des jours de fermeture
        procedure LitJourFermeture;
        procedure EcritJourFermeture;
        // control des chanps
        procedure IsMouvementer;
        function RIBRenseigner: Boolean;
        function CodExist : Boolean;
        function VerifiRibBic : Boolean;
        function OkPourChangementDevise(OldCode,NewCode : String) : Boolean;
        {$IFDEF TRESO}
        procedure GeneralElipsisClick(Sender: TObject);
        {$ENDIF TRESO}

      procedure BAgenceOnClick(Sender: TObject);
      procedure ConstruireRib(Sender : TObject);
      function  GetCodeGuichet(Chp, Valeur : string) : string;
      function  ValideGen    : Boolean;
      function  ValideAgence : Boolean;
      function  ValideBanque : Boolean;
      function  TesteNature : Boolean;
      function  TesteDevise : Boolean;
      procedure OnDossierChange;
      procedure OnNatureChange;
    public
      {A True si EstMultiSoc and EstComptaTreso}
      EstTresoMultiSoc : Boolean;
      FNoDossier : string;
      {$IFDEF EAGLCLIENT}
      BQ_GENERAL : THEdit;
      {$ELSE}
      BQ_GENERAL : THDBEdit;
      {$ENDIF EAGLCLIENT}

      FSuppressionClick : TNotifyEvent; {FQ 15585 : pour la surchage du bouton de suppression}
        // svg de parametre de lancement
        LeQuel : string;
        Mode : TActionFiche;
        LaPage : integer;
        // mouvement du compte
        AvecMvt : boolean;
        // est en court de chargement ?
        gbChargement : boolean;
        {JP 23/09/03 : C'est un peu redondant avec gbChargement, mais en EAGL, je ne m'en sors pas.
                       OnChangeBQ_PAYS est exécuté trois fois et fait sauter le code IBAN}
        EnSortie     : Boolean;
        Pages : TPageControl;

        OkMajTotPointe : boolean;
        OldCodeDev : string;
        CodeISO : string;
        codeTypePays : string;
        {$IFDEF EAGLCLIENT}
        CodePays : THValComboBox;
        {$ELSE}
        CodePays : THDBValComboBox;
        {$ENDIF}

        FParamPays : Boolean; {JP 05/07/07 : FQ 20319}
        procedure MajCodeIban;{JP 05/07/07 : FQ 20319}
    end;
                           

//================================================================================
// Implementation
//================================================================================
Implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPVersion,
  CPTypeCons,
  {$ENDIF MODENT1}
  Commun, ParamSoc,
  Constantes,
  HTB97; //XMG 14/07/03

//==================================================
// Definition des Constante
//==================================================
// les tag de pages ;)
const PIdent    : integer = 0;
const PCoordo   : integer = 1;
const PCommunic : integer = 2;
const PDelai    : integer = 3;
const PInfo     : integer = 4;

//==================================================
// Definition des variable
//==================================================
var
    MESS : array [0..20] of string = (
        '0;Comptes bancaires;Voulez-vous enregistrer les modifications?;Q;YNC;Y;C;',
        '1;Comptes bancaires;Confirmez-vous la suppression de l'#39'enregistrement?;Q;YNC;N;C;',
        '2;Comptes bancaires;Vous devez renseigner un code !;E;O;O;O;',
        '3;Comptes bancaires;Vous devez renseigner un libellé !;E;O;O;O;',
        '4;Comptes bancaires;Vous devez renseigner le code banque du RIB;E;O;O;O;',
        '5;Comptes bancaires;Vous devez renseigner le code guichet du RIB;E;O;O;O;',
        '6;Comptes bancaires;Vous devez renseigner le numéro de compte du RIB;E;O;O;O;',
        '7;Comptes bancaires;Vous devez renseigner la clé du RIB;E;O;O;O;',
        '8;Comptes bancaires;Vous devez renseigner le RIB ou le code BIC du compte;E;O;O;O;',
        'L''enregistrement est inaccessible',
        '10;Comptes bancaires;Le code que vous avez saisi existe déja. Vous devez le modifier;E;O;O;O;',
        '11;Comptes bancaires;La clé RIB est erronnée. Souhaitez-vous la recalculer ?;Q;YN;Y;C;',
        '12;Comptes bancaires;Vous ne pouvez pas affecter cette devise sur le compte;E;O;O;O;',
        '13;Comptes bancaires;La domiciliation comporte des caractères interdits !;E;O;O;O;',
        '14;Comptes bancaires;Le compte général rattaché au compte est incorrect !;E;O;O;O;',
        '15;Comptes bancaires;La banque n''existe pas !;E;O;O;O;',
        '16;Comptes bancaires;L''agence n''existe pas !;E;O;O;O;',
        '17;Comptes bancaires;La clé du code IBAN est erronée. Voulez-vous la recalculer ?;Q;YN;Y;C;',
        '18;Comptes bancaires;Vous devez renseigner le code IBAN;E;O;O;O;',
        '19;Comptes bancaires;Vous devez renseigner le type d''identification.',
        '20;Comptes bancaires;La longueur minimale du code IBAN est de 10 caractères.'
        );

{25/10/06 : Par le DispatchTT, on reçoit le code et non le général => pas besoin du NoDossier.
            Quant à QuellePage il ne sert à rien ...
{---------------------------------------------------------------------------------------}
procedure FicheBanqueCPCode(Code : string ; Comment : TActionFiche);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche('CP', 'CPBANQUECP', '', Code, 'ACTION=' + TAToStr(Comment) + ';' + TAToStr(Comment) + ';;0;;');
end;

procedure FicheBanqueCP(CpteGeneral : String ; Comment : TActionFiche ; QuellePage : Integer; Dossier : string = '');
var
    Q : TQuery;// TOB; JP 30/07/03 préférable en 2/3
    code,action : string;
begin
    action := 'CREATION';
    // si modification
    if (Comment in [taConsult,taModif,taModifEnSerie]) then
    begin
        // recup du code du RIB
        if Dossier = '' then
          Q := OpenSQL('SELECT BQ_CODE FROM BANQUECP WHERE BQ_GENERAL = "' + CpteGeneral + '"', True)
        else
          {JP 31/07/06 : mise en place du NoDossier}
          Q := OpenSQL('SELECT BQ_CODE FROM BANQUECP WHERE BQ_GENERAL = "' + CpteGeneral + '" AND BQ_NODOSSIER = "' + Dossier + '"', True);
        if not Q.EOF then
          Code := Q.Fields[0].AsString;
        Ferme(Q);
        { FQ 15967 - CA - 23/06/2005 - Mise à jour du mode consultation }
        if Comment=taConsult then
          action := 'CONSULTATION'
        {JP 01/08/06 : si le code n'existe pas et que l'on est en consultation, il ne me parait
                       pas mal de passer en création, sinon on se positionne sur le mauvais enregistrement}
        else if Code = '' then begin
          Action  := 'CREATION';
          Comment := taCreat;
        end
        else
          Action := 'MODIFICATION';
    end;
    // lancement de la fiche
    AGLLanceFiche('CP','CPBANQUECP','',code,'ACTION=' + action + ';' + TAToStr(Comment) + ';' + CpteGeneral + ';' +
                  IntToStr(QuellePage) + ';' + Dossier + ';');
end;

//==================================================
// Evenements par default de la TOM
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.OnNewRecord;
var
  S1 : string;
begin
    Inherited;

    SetControlEnabled('BQ_BANQUE',true);
    {JP 03/08/05 : FQ 16126
    SetControlEnabled('BQ_CODE',true);}

    SetField('BQ_GENERAL',Lequel);
    SetField('BQ_JOURFERMETUE','000000X');
    SetField('BQ_PAYS',CodePaysDeIso(VH^.PaysLocalisation)) ; //codepaysdeiso(CodeISOFR)) ; //'FR'));
    SetField('BQ_LANGUE',V_PGI.LanguePrinc);
    SetField('BQ_DEVISE',V_PGI.DevisePivot);
    SetField('BQ_RAPPAUTOREL','X');
    SetField('BQ_RAPPROAUTOLCR','X');
    SetField('BQ_ECHEREPPRELEV','X');
    SetField('BQ_ECHEREPLCR','X');

    if EstMultiSoc then begin
      SetField('BQ_NATURECPTE', tcb_Bancaire);
      SetField('BQ_NODOSSIER', FNoDossier); {JP 31/07/06}
      S1 := GetInfosFromDossier('DOS_NODOSSIER', FNoDossier, 'DOS_SOCIETE');
      SetField('BQ_SOCIETE', S1); {JP 19/09/06}
    end
    else begin
      SetField('BQ_NODOSSIER', V_PGI.NoDossier); {JP 08/11/06 : C'est préférable pour PCL plutôt que CODEDOSSIERDEF}
      SetField('BQ_SOCIETE', V_PGI.CodeSociete); {JP 19/09/06}
    end;

    LookUpValueExist(GetControl('BQ_GENERAL'));
    SetControlProperty('BQ_LETTRELCR','ItemIndex','0');

    LitJourFermeture;

    gbChargement := false;
    // ici l'AGL set les zone de l'ecran avec les variable du DataSet !
    // Affichage :
    // FQ 19465 BV 11/01/2007
    codeISO := codeISODuPays(GetControlText('BQ_PAYS'));
    // END FQ 19465
    //AfficheEcran;

    OnChangeBQ_PAYS(nil);

end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.OnDeleteRecord;
begin

  if EstComptaTreso and
    ExisteSQL('SELECT TE_GENERAL FROM ' + GetTableDossier(GetParamSocSecur('SO_TRBASETRESO', ''), 'TRECRITURE') +
              ' WHERE TE_GENERAL = "' + GetField('BQ_CODE') + '"') then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Ce compte est utilisé en trésorerie, il n''est pas possible de le supprimer.');
  end;

  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.OnUpdateRecord;
var RibCalcul : String ;
    sType,stParam,szIban : string;
    StrType : String ;
begin
    Inherited;
    LastError := 1;

    if (not (DS.state in [dsEdit,dsInsert])) then exit;  
    sType := quelType(codeISO,codeTypePays);

    if (GetControlText('BQ_CODE') = '') then
    begin
        HShowMessage(MESS[2],'','');
        Pages.ActivePageIndex := PIdent;
        SetFocusControl('BQ_CODE');
        LastError := 1;
        exit;
    end;

    if (GetControlText('BQ_LIBELLE') = '') then
    begin
        HShowMessage(MESS[3],'','');
        Pages.ActivePageIndex := PIdent;
        SetFocusControl('BQ_LIBELLE');
        LastError := 1;
        exit;
    end;

    if not ValideGen then begin
        HShowMessage(MESS[14],'','');
        Pages.ActivePageIndex := PIdent;
        SetFocusControl('BQ_GENERAL');
        LastError := 1;
        exit;
    end;

    if not ValideAgence then begin
        HShowMessage(MESS[16],'','');
        Pages.ActivePageIndex := PIdent;
        SetFocusControl('BQ_AGENCE');
        LastError := 1;
        exit;
    end;

    if not ValideBanque then begin
        HShowMessage(MESS[15],'','');
        Pages.ActivePageIndex := PIdent;
        SetFocusControl('BQ_BANQUE');
        LastError := 1;
        exit;
    end;

    if (ExisteCarInter(GetControlText('BQ_DOMICILIATION'))) then
    begin
        HShowMessage(MESS[13],'','') ;
        Pages.ActivePageIndex := PIdent;
        SetFocusControl('BQ_DOMICILIATION');
        LastError := 1;
        exit;
    end;
    //SDA le 14/03/2007
    if GetControlVisible('BQ_TYPEPAYS') = true then
    begin
     StrType := GetControlText('BQ_TYPEPAYS');
     if StrType = '' then
     begin
      HShowMessage(MESS[19],'','');
      SetFocusControl('BQ_TYPEPAYS');
      LastError := 1;
      exit;
     end;
     //SetControlText('BQ_TYPEPAYS', CodePays.Values[CodePays.itemindex]);
     SetField('BQ_TYPEPAYS', CodePays.Values[CodePays.itemindex]);
    end;
    //Fin SDA le 14/03/2007
    if EstTresoMultiSoc and (GetControlText('BQ_NATURECPTE') = tcb_Courant) and
       (GetControlText('BQ_CODECIB') = '') then begin
      Pages.ActivePageIndex := PIdent;
      SetFocusControl('BQ_CODECIB');
      LastErrorMsg := TraduireMemoire('Veuillez saisir un CIB pour votre compte courant');
      LastError := 1;
      Exit;
    end;

    if IsTresoMultiSoc and (GetControlText('BQ_NODOSSIER') = '') then begin
      SetFocusControl('BQ_NODOSSIER');
      LastErrorMsg := TraduireMemoire('Veuillez choisir un dossier');
      LastError := 1;
      Exit;
    end
    {JP 11/05/07 : FQ TRESO 10462 : Nécessaire en cwas au moins, car la tablette sur dossiers étant
                   vide le champ est vider}
    else if not IsTresoMultiSoc and (GetField('BQ_NODOSSIER') <> V_PGI.NoDossier) then
      SetField('BQ_NODOSSIER', V_PGI.NoDossier);

    if EstTresoMultiSoc and (GetControlText('BQ_NATURECPTE') = '') then begin
      Pages.ActivePageIndex := PIdent;
      SetFocusControl('BQ_NATURECPTE');
      LastErrorMsg := TraduireMemoire('Veuillez affecter une nature au compte');
      LastError := 1;
      Exit;
    end;

    if EstTresoMultiSoc and not TesteNature then begin
      Pages.ActivePageIndex := PIdent;
      SetFocusControl('BQ_NATURECPTE');
      LastErrorMsg := TraduireMemoire('Vous ne pouvez pas affecter cette nature à ce type de compte');
      LastError := 1;
      Exit;
    end;

    if not TesteDevise then begin
      Pages.ActivePageIndex := PIdent;
      SetFocusControl('BQ_DEVISE');
      LastErrorMsg := TraduireMemoire('Vous devez affecter la devise pivot aux comptes courants et titres');
      LastError := 1;
      Exit;
    end;

    if ((VH^.CtrlRIB) and ((codeISO = CodeISOFR) or (codeISO = CodeISOES)) and (RIBRenseigner)) then //XMG 14/07/03
    begin
        RibCalcul := VerifRib(GetControlText('BQ_ETABBQ'),GetControlText('BQ_GUICHET'),GetControlText('BQ_NUMEROCOMPTE'),CodeISO); //XMG 14/07/03
        if (RibCalcul <> Trim(GetControlText('BQ_CLERIB'))) Then
        begin
            if (HShowMessage(MESS[11],'','') <> mrYes) Then
            begin
                SetFocusControl('BQ_CLERIB');
                SetField('BQ_CODEIBAN',''); //XMG 14/07/03
                exit;
            end
            else
            begin
                {SetControlText('BQ_CLERIB',RibCalcul); JP 19/03/04, le SetField s'avère plus efficace}
                SetField('BQ_CLERIB',RibCalcul);
                SetField('BQ_CODEIBAN',''); //XMG 14/07/03
            end;
        end;
    end;

    if (DS.state in [dsInsert]) then
    begin
        if (CodExist) then exit;
        if (not VerifiRibBic) then exit;
    end
    else if (DS.state in [dsEdit]) then
    begin
        if (AvecMvt and (OldCodeDev <> GetControlText('BQ_DEVISE'))) Then
        begin
            OkMajTotPointe := false;
            If (OkPourChangementDevise(OldCodeDev,GetControlText('BQ_DEVISE'))) Then OkMajTotPointe := true
            Else
            begin
                HShowMessage(MESS[12],'','');
                Pages.ActivePageIndex := PIdent;
                SetFocusControl('BQ_DEVISE');
                LastError := 1;
                exit;
            end;
        end;
    end;

  // Verification RIB
  if (sType = 'RIB') or
  ((sType = '') and (codeISO = CodeISOFR) and (VH^.PaysLocalisation = CodeISOFR)) or
  ((sType = '') and (codeISO = CodeISOES) and (VH^.PaysLocalisation = CodeISOES)) then
  begin
    // FQ 19489 BV 11/01/2007
    if (codeISO <> CodeISOFR) then
    begin
    	// On verifie l'IBAN car un des codes du RIB peut avoir été modifié sans qu'on modifie l'IBAN
	    stParam := calcRIB(codeISO,GetControlText('BQ_ETABBQ'),GetControlText('BQ_GUICHET'),GetControlText('BQ_NUMEROCOMPTE'),GetControlText('BQ_CLERIB'));
    	if stParam <> '' then
      	szIban := calcIBAN(codeISO, stParam);
    end;
    stParam := GetControlText('BQ_CODEIBAN');
    if (stParam <> szIban) or (sType='RIB') or (sType='') then
  	begin
      SetControlText('BQ_CODEIBAN','');
      Verif_CODEIBAN;
    end;
    // END FQ 19489
    stParam := GetControlText('BQ_CODEIBAN');
      //SDA le 14/03/2007
    if length(GetControlText('BQ_CODEIBAN')) < 10 then
    begin
       HShowMessage(MESS[20],'','');
       SetFocusControl('BQ_CODEIBAN');
       LastError := 1;
       Exit;
    end;
    //Fin SDA le 14/03/2007
    if ErreurDansIban(stParam) then
      if HShowMessage(MESS[17],'','') = mrYes then
      begin
        SetControlText('BQ_CODEIBAN','');
        Verif_CODEIBAN;
      end;
  end
  else if (sType = 'IBA') or (sType = '') then
  begin
    // Verification de l'IBAN
    stParam := GetControlText('BQ_CODEIBAN');
    if Trim(stParam) = '' then
    begin
       HShowMessage(MESS[18],'','');
       SetFocusControl('BQ_CODEIBAN');
       LastError := 1;
       Exit;
    end;
      //SDA le 14/03/2007
    if length(GetControlText('BQ_CODEIBAN')) < 10 then
    begin
       HShowMessage(MESS[20],'','');
       SetFocusControl('BQ_CODEIBAN');
       LastError := 1;
       Exit;
    end;
    //Fin SDA le 14/03/2007
    if ErreurDansIban(stParam) then
      if HShowMessage(MESS[17],'','') = mrYes then
      begin
        Verif_CODEIBAN;
      end;
  end;
    {JP 23/09/03 : Pour éviter les traitement de PaysChange}
    EnSortie := True;

    EcritJourFermeture;
    LastError := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.OnAfterUpdateRecord;
begin
    Inherited;
    {$IFNDEF EAGLCLIENT}
    // SBO : 30/06/2003 : pour forcer la fermeture de la fiche à la validation du copte géné
    // uniquement en 2tiers, bidouille en attente refonte de la fiche......
    if Mode = taCreatOne then
      TFFiche(Ecran).BFermeClick(nil) ;
    {$ENDIF EAGLCLIENT}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.OnLoadRecord;
begin
    gbChargement := true;

    Inherited;
    SetControlText('EDev',GetControlText('BQ_DEVISE'));

    IsMouvementer;

    SetControlEnabled('BQ_GENERAL',((Lequel = '') and (GetField('BQ_GENERAL') = '')));
    {Le comportement de BQ_NODOSSIER est le même que BQ_GENERAL car les deux ensembles constituent
     une sorte de clef primaire BIS}
    SetControlEnabled('BQ_NODOSSIER', ((Lequel = '') and (GetField('BQ_GENERAL') = '')));
    if GetField('BQ_NATURECPTE') = '' then
      SetField('BQ_NATURECPTE', tcb_Bancaire);

    SetControlEnabled('BQ_BANQUE',((Not AvecMvt) or (GetField('BQ_BANQUE') = '')));
    {JP 03/08/05 : FQ 16126 : BQ_CODE constituant l'index primaire, il n'est pas normal de pouvoir
                   le modifier. Par ailleurs, ce champs n'étant utiliser nulle part, il ne me semble
                   pas esssentiel de pouvoir le modifier.
    SetControlEnabled('BQ_CODE',(not AvecMvt));}

    // Gestion des erreurs dues au pays non renseigné
    if (GetField('BQ_PAYS') = '') then if (RibRenseigner) then
    begin
        SetField('BQ_PAYS',CodePaysDeIso(VH^.PaysLocalisation)) ; //CodePaysDeIso(CodeISOFR)) ; //'FR'));
        PGIBox('Votre code pays a été modifié. Mise à jour effectué', 'Attention');
    end;

    { FQ 19473 BV 11/01/2007
    if VH^.PaysLocalisation<>CodeISOES then
    Begin      }
       OnChangeBQ_PAYS(nil);
       Verif_CODEIBAN;
	{    End ; //XVI 24/02/2005 }

    LitJourFermeture;

    gbChargement := false;
    // ici l'AGL set les zone de l'ecran avec les variable du DataSet !  *)
    // Affichage :
    // Gestion differente selon si on change de pays ou non.
    if (codeISO <> codeISODuPAys(GetControlText('BQ_PAYS'))) then
      OnChangeBQ_PAYS(nil)
    else
      AfficheEcran;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.OnChangeField(F: TField);
begin
  Inherited;
  if F.FieldName = 'BQ_PAYS' then OnChangeBQ_PAYS(nil) else
  if F.FieldName = 'BQ_TYPEPAYS' then OnChangeBQ_TYPEPAYS(nil) else
  if VH^.PaysLocalisation=CodeISOES then
  begin
     if F.FieldName = 'BQ_DEVISE' then OnChangeBQ_DEVISE(nil);
  End
  else if F.FieldName = 'BQ_NODOSSIER'  then OnDossierChange
  else if F.FieldName = 'BQ_NATURECPTE' then OnNatureChange
  {JP 05/07/07 : FQ 20319 : Calcul de l'IBAN en Direct}
  else if (F.FieldName = 'BQ_CODEIBAN')     or
          (F.FieldName = 'BQ_ETABBQ')       or
          (F.FieldName = 'BQ_GUICHET')      or
          (F.FieldName = 'BQ_NUMEROCOMPTE') or
          (F.FieldName = 'BQ_CLERIB')       then MajCodeIban
  else Begin
    if (F.FieldName = 'BQ_BANQUE') then begin
      {JP 18/12/03 : On filtre les agences en fonction de la banque}
      {$IFDEF EAGLCLIENT}
      THValComboBox(GetControl('BQ_AGENCE')).Plus := 'TRA_BANQUE = "' + GetControlText('BQ_BANQUE') + '"';
      THEdit(GetControl('BQ_CODECIB')).Plus := 'TCI_BANQUE = "' + GetControlText('BQ_BANQUE') + '"';
      {$ELSE}
      THDBValComboBox(GetControl('BQ_AGENCE')).Plus := 'TRA_BANQUE = "' + GetControlText('BQ_BANQUE') + '"';
      THDBEdit(GetControl('BQ_CODECIB')).Plus := 'TCI_BANQUE = "' + GetControlText('BQ_BANQUE') + '"';
      {$ENDIF}
    end;
  End ; //XVI 24/02/2005
end;

procedure TOM_BANQUECP.OnArgument(S: String);
var
    S1,S2 : string;
    BQ_NUMEROCOMPTE,BQ_CODEPOSTAL,BQ_DIVTERRIT,BQ_VILLE : THEdit;
    BQ_CODEIBAN(*,BQ_CLERIB*) : THEdit;
    BQ_DEVISE(*,BQ_PAYS*) : THValComboBox;
    BMCbanq : TBitBtn;
begin
    Inherited;
    // On sauvegarde les positions des THDBEdit par defaut afin de les réutiliser plus tard.
    SaveAffichageDefaut(ecran,'BQ_');
    {$IFDEF EAGLCLIENT}
      BQ_GENERAL := THEdit(GetControl('BQ_GENERAL'));
      CodePays := THValComboBox(GetControl('BQ_TYPEPAYS', true));
    {$ELSE}
      BQ_GENERAL := THDBEdit(GetControl('BQ_GENERAL'));
      CodePays := THDBValComboBox(GetControl('BQ_TYPEPAYS', true));
    {$ENDIF}

    gbChargement := true;

    Ecran.HelpContext := 500006;
    S1 := UpperCase(S);
    S2 := ReadTokenSt(S1); // mode pour l'AGL
    S2 := ReadTokenSt(S1); Mode := StrToTA(S2);     // mode
    S2 := ReadTokenSt(S1); LeQuel := S2;            // compte
    S2 := ReadTokenSt(S1); LaPage := StrToInt(S2);  // page
    S2 := ReadTokenSt(S1); FNoDossier := S2;
    if FNoDossier = '' then FNoDossier := V_PGI.NoDossier;

//XMG 26/11/03 Début
    // recup des control
    BMCbanq := TBitBtn(GetControl('BMCbanq',true)); if (assigned(BMCbanq)) then BMCbanq.OnClick := OnClickBMCbanq;
    pages := TPageControl(GetControl('Pages',true)); if (assigned(pages)) then Pages.ActivePageIndex := PIdent;


    BQ_NUMEROCOMPTE := THEdit(GetControl('BQ_NUMEROCOMPTE',true));
    BQ_CODEPOSTAL := THEdit(GetControl('BQ_CODEPOSTAL',true));
    BQ_DIVTERRIT := THEdit(GetControl('BQ_DIVTERRIT',true));
    BQ_VILLE := THEdit(GetControl('BQ_VILLE',true));

    {JP 03/03/04 : FQ Trésorerie 10007}
//    SetControlVisible('PDELAI', False); {JP 14/05/04 : Il y a une zone nécessaire sur cette onglet : en attendant de la déplacer !!}

    if (assigned(BQ_NUMEROCOMPTE)) then BQ_NUMEROCOMPTE.OnKeyPress := OnKeyPressBQ_NUMEROCOMPTE;
    if (assigned(BQ_CODEPOSTAL)) then BQ_CODEPOSTAL.OnDblClick := OnDblClickBQ_CODEPOSTAL;
    if (assigned(BQ_DIVTERRIT)) then BQ_DIVTERRIT.OnDblClick := OnDblClickBQ_DIVTERRIT;
    if (assigned(BQ_VILLE)) then BQ_VILLE.OnDblClick := OnDblClickBQ_VILLE;
    if VH^.PaysLocalisation<>CodeISOEs then
    Begin
        BQ_CODEIBAN := THEdit(GetControl('BQ_CODEIBAN',true));
        //BQ_PAYS := THValComboBox(GetControl('BQ_PAYS',true));
        BQ_DEVISE := THValComboBox(GetControl('BQ_DEVISE',true));
        //BQ_CLERIB := THEdit(GetControl('BQ_CLERIB',true));

        //if (assigned(BQ_CODEIBAN)) then BQ_CODEIBAN.OnEnter := OnEnterBQ_CODEIBAN;
        if (assigned(BQ_CODEIBAN)) then BQ_CODEIBAN.OnKeyPress := OnKeyPressBQ_NUMEROCOMPTE;
        // if (assigned(BQ_PAYS)) then BQ_PAYS.OnChange := OnChangeBQ_PAYS;
        if (assigned(codePays)) then codePays.OnChange := OnChangeBQ_TYPEPAYS;
        if (assigned(BQ_DEVISE)) then BQ_DEVISE.OnChange := OnChangeBQ_DEVISE;
        //if (assigned(BQ_CLERIB)) then BQ_CLERIB.OnExit := OnEnterBQ_CODEIBAN;
    End ; //XVI 24/02/2005
    AvecMvt := false;
    OldCodeDev := '';

    ChangeMask(THNumEdit(GetControl('BQ_DERNSOLDEFRS',true)),V_PGI.OkDecV,V_PGI.SymbolePivot);
    ChangeMask(THNumEdit(GetControl('BQ_DERNSOLDEDEV',true)),V_PGI.OkDecV,V_PGI.SymbolePivot);

    {Dans la tréso, on gère la notion d'Agence, d'où l'ajout de la combo TRADOMICIL et du bouton
     BAGENCE. Lors du chargement du champ BQ_DOMICILIATION, on met à jour la combo.}
    SetControlVisible('BQ_DOMICILIATION', not EstComptaTreso);
    SetControlVisible('BQ_AGENCE', EstComptaTreso);
    SetControlVisible('BAGENCE', EstComptaTreso);
    TBitBtn(GetControl('BAGENCE')).OnClick := BAgenceOnClick;

    EstTresoMultiSoc := EstMultiSoc and EstComptaTreso and EstTablePartagee('BANQUECP');
    {$IFDEF TRESO}
    SetControlVisible('BQ_NODOSSIER' , IsTresoMultiSoc);
    SetControlVisible('TBQ_NODOSSIER', IsTresoMultiSoc);
      {$IFDEF EAGLCLIENT}
      if IsTresoMultiSoc then THValComboBox(GetControl('BQ_NODOSSIER')).Plus := 'DOS_NODOSSIER ' + FiltreNodossier
                         else THValComboBox(GetControl('BQ_NODOSSIER')).Plus := 'DOS_NODOSSIER = "' + V_PGI.NoDossier + '"';
      {$ELSE}
      if IsTresoMultiSoc then THDBValComboBox(GetControl('BQ_NODOSSIER')).Plus := 'DOS_NODOSSIER ' + FiltreNodossier
                         else THDBValComboBox(GetControl('BQ_NODOSSIER')).Plus := 'DOS_NODOSSIER = "' + V_PGI.NoDossier + '"';
      {$ENDIF EAGLCLIENT}
    {$ELSE}
    SetControlVisible('BQ_NODOSSIER' , False);
    SetControlVisible('TBQ_NODOSSIER', False);
    {$ENDIF TRESO}
    SetControlVisible('GBTRESO', EstTresoMultiSoc);
    {$IFDEF EAGLCLIENT}
    {27/07/06 : Pour le moment on ne gère pas les comptes titres}
    THValComboBox(GetControl('BQ_NATURECPTE')).Plus := 'AND CO_CODE <> "' + tcb_Titre + '"';
    {$ELSE}
    THDBValComboBox(GetControl('BQ_NATURECPTE')).Plus := 'AND CO_CODE <> "' + tcb_Titre + '"';
    {$ENDIF EAGLCLIENT}

    {$IFDEF TRESO}
    SetControlProperty('PCOMMUNIC', 'TABVISIBLE', False);
    SetControlProperty('PDELAI'   , 'TABVISIBLE', False);
    {JP 11/05/07 : FQ TRESO 10463 : Pour cacher le bouton zoom}
    BQ_GENERAL.OnElipsisClick := GeneralElipsisClick;
    {$ELSE}
    if EstTresoMultiSoc then begin
      //BQ_GENERAL.DataType := 'TRGENERALBQEDIV';
      BQ_GENERAL.DataType := 'TZGENERAL';
      S1 := GetInfosFromDossier('DOS_NODOSSIER', FNoDossier, 'DOS_NOMBASE');
      {Un compte général ne peut être associé qu'à un compte bancaire par dossier}
      BQ_GENERAL.Plus := '(G_NATUREGENE = "BQE" AND G_GENERAL NOT IN (SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_NODOSSIER = "' + FNoDossier + '")) OR ' +
                         '(G_NATUREGENE = "DIV" AND G_GENERAL IN (SELECT CLS_GENERAL FROM ' + GetTableDossier(S1, 'CLIENSSOC') + ') AND ' +
                         'G_GENERAL NOT IN (SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_NODOSSIER = "' + FNoDossier + '"))';
    end
    else begin
      BQ_GENERAL.DataType := 'TZGBANQUE';
      {Un compte général ne peut être associé qu'à un compte bancaire}
      BQ_GENERAL.Plus := 'G_GENERAL NOT IN (SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_NODOSSIER = "' + FNoDossier + '")';
    end;
    {$ENDIF TRESO}

    SetControlVisible('BINSERT', False);

    THValComboBox(GetControl('BQ_BANQUE')).OnChange := ConstruireRib;
    {Lors du changement de valeur de la combo, on met à jour le champ BQ_DOMICILIATION}
    THValComboBox(GetControl('BQ_AGENCE')).OnChange := ConstruireRib;

    {JP 27/09/05 : FQ 15585 surcharge du Click du bouton de suppression}
    FSuppressionClick := TToolbarButton97(GetControl('BDELETE')).OnClick;
    TToolbarButton97(GetControl('BDELETE')).OnClick := BDleteClick;

    //SDA le 04/01/2007 version belge
    if VH^.PaysLocalisation=CodeISOBE then
    begin
      SetControlVisible('BQ_RELEVEETRANGER', False);
      SetControlVisible('TBQ_DESTINATAIRE', False);
      SetControlVisible('BQ_DESTINATAIRE', False);
    end;
    //Fin SDA le 04/01/2007

  {JP 07/06/07 : FQ 19591 : Gestion des modèles de bordereau : documents ou états}
  InitBordereau;  
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.OnClose;
begin
    Inherited;
    DestroySaveAffichage ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.OnCancelRecord;
begin
  Inherited;
end;

//==================================================
// Autres Evenements
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.OnClickBMCbanq(Sender: TObject);
var
  {$IFDEF EAGLCLIENT} {JP 09/02/04}
  BQ_BANQUE : THValComboBox;
  {$ELSE}
  BQ_BANQUE : THDBValComboBox;
  {$ENDIF}
  Laction : TActionFiche;
begin
    SetFocusControl('BQ_CODE');
  {$IFDEF EAGLCLIENT} {JP 09/02/04}
  BQ_BANQUE := THValComboBox(GetControl('BQ_BANQUE',true));
  {$ELSE}
  BQ_BANQUE := THDBValComboBox(GetControl('BQ_BANQUE',true));
  {$ENDIF}
    if (not assigned(BQ_BANQUE)) then exit;

    if (BQ_BANQUE.Value = '') then Laction := taCreatOne
    { FQ 15967 - CA - 23/06/2005 - Utilisation du mode courant de consultation }
    else Laction := Mode;

    if EstComptaTreso then
      TRLanceFiche_Banques('TR','TRCTBANQUE', '', BQ_BANQUE.Value, ActionToString(LAction))
    else
      FicheBanque_AGL(BQ_BANQUE.Value,Laction,0);

    BQ_BANQUE.Reload;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.OnChangeBQ_PAYS(Sender: TObject);
begin
    if (EnSortie) then exit;

    codeISO := codeISODuPays(GetControlText('BQ_PAYS'));
    if trim(codeISO) <> '' then
      CodePays.Plus := ' AND YIB_PAYSISO="' + codeISO + '"'
    else
      CodePays.Plus := '';
    CodePays.Refresh ;

    // Affichage :
    AfficheEcran;
    if ds.State in [dsInsert, dsEdit] then
    begin
      SetField('BQ_ETABBQ', '');
      SetField('BQ_GUICHET', '');
      SetField('BQ_NUMEROCOMPTE', '');
      SetField('BQ_CLERIB', '');
      //SetField('BQ_TYPEPAYS', '');  // FQ 19459 BV 11/01/2007
      { FQ 19469 BV 01/02/2007
      if (((codeISO = codeISOES) and (VH^.PaysLocalisation = CodeISOES)) or ((codeISO = CodeISOFR) and not(existeParamPays(codeISO,codeTypePays))))  or (quelType(codeISO,GetControlText('BQ_TYPEPAYS')) = 'BBA') then
        SetField('BQ_CODEIBAN', '')
      else
        SetField('BQ_CODEIBAN', codeISO)
        }
      if
      (quelType(codeISO,GetControlText('BQ_TYPEPAYS')) = 'IBA') or
         ((quelType(codeISO,GetControlText('BQ_TYPEPAYS')) = '') and
         (((codeISO <> codeISOES) and (VH^.PaysLocalisation <> CodeISOES)) or ((codeISO <> codeISOFR) and (VH^.PaysLocalisation <> CodeISOFR)))) then
         SetField('BQ_CODEIBAN', codeISO)
      else
         SetField('BQ_CODEIBAN', '');
    end;


    if VH^.PaysLocalisation=CodeISOES then
    Begin
       if (ds.state in [dsInsert,dsEdit]) then Begin
         SetField('BQ_ETABBQ','');
         SetField('BQ_GUICHET','');
         SetField('BQ_NUMEROCOMPTE','');
         SetField('BQ_CLERIB','');
       end ;
    End ; //XVI 24/02/2005
  // On met a jour si besoin le rib:
  ConstruireRib(GetControl('BQ_BANQUE'));
  ConstruireRib(GetControl('BQ_AGENCE'));
  OnChangeBQ_TYPEPAYS(nil);
    //SDA le 15/03/2007
    if (ds.State in [dsInsert]) or
{$IFDEF EAGLCLIENT}
   (THVALComboBox(GetControl('BQ_TYPEPAYS')).text = '') then
{$ELSE}
   (THDBVALComboBox(GetControl('BQ_TYPEPAYS')).text = '') then
{$ENDIF}
{$IFDEF EAGLCLIENT}
  if THVALComboBox(GetControl('BQ_TYPEPAYS')).visible = true then
  begin
   IF (THVALComboBox(GetControl('BQ_TYPEPAYS')).text = '') then
      THVALComboBox(GetControl('BQ_TYPEPAYS')).ItemIndex := 0;
   //SetControlText('BQ_TYPEPAYS', CodePays.Values[CodePays.itemindex]); //SDA le 27/12/2007
   SetField('BQ_TYPEPAYS', CodePays.Values[CodePays.itemindex]); //SDA le 27/12/2007
  end;
{$ELSE}
  if THDBVALComboBox(GetControl('BQ_TYPEPAYS')).visible = true then
  begin
   if (THDBVALComboBox(GetControl('BQ_TYPEPAYS')).text = '') then
      THDBVALComboBox(GetControl('BQ_TYPEPAYS')).ItemIndex := 0;
   //SetControlText('BQ_TYPEPAYS', CodePays.Values[CodePays.itemindex]); //SDA le 27/12/2007
   SetField('BQ_TYPEPAYS', CodePays.Values[CodePays.itemindex]); //SDA le 27/12/2007
  end;
{$ENDIF}
    //SDA le 15/03/2007

end;

procedure TOM_BANQUECP.OnChangeBQ_TYPEPAYS(Sender: TObject);
var stTypePays : string;
begin
  stTypePays := GetControlText('BQ_TYPEPAYS');
  codeIso := codeISODuPays(GetControlText('BQ_PAYS'));

  if codeIso = '' then exit;
  if CodeTypePays <> stTypePays then
  begin
    CodeTypePays := stTypePays ;
    AfficheEcran; // SDA le 27/12/2007
    if ds.State in [dsInsert, dsEdit] then
    begin
      // Affichage :
      // Affiche(ecran,codeISO,CodeTypePays,'BQ_',3);
      //SDA le 27/12/2007 AfficheEcran;
      SetField('BQ_ETABBQ', '');
      SetField('BQ_GUICHET', '');
      SetField('BQ_NUMEROCOMPTE', '');
      SetField('BQ_CLERIB', '');
      //if (((codeISO = codeISOES) and (VH^.PaysLocalisation = CodeISOES)) or (codeISO = CodeISOFR)) or (quelType(codeISO,GetControlText('BQ_TYPEPAYS')) = 'BBA') then
      //RR 15/03/2007 FQ 19833 .... à contre coeur...
      if (quelType(codeISO,GetControlText('BQ_TYPEPAYS')) = 'BBA') then
        SetField('BQ_CODEIBAN', '')
      else
        SetField('BQ_CODEIBAN', codeISO);
    end;
  end;
  // On met a jour si besoin le rib:
  ConstruireRib(GetControl('BQ_BANQUE'));
  ConstruireRib(GetControl('BQ_AGENCE'));
  // MAJ du libelle
  //THLabel(ecran.FindComponent('LBQ_TYPEPAYS')).Caption := GetLibelleParamPays(codeIso,CodeTypePays);

end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.Verif_CODEIBAN;
var
    szPays,szIban,sType,sParam,sIban : string;
begin 
  sIban := GetControlText('BQ_CODEIBAN');
  szPays := Copy(codeISO,1,2);
  // FQ 19475 BV 11/01/2007 if VH^.PaysLocalisation=CodeISOES then 
  if szPays = CodeISOES then
  begin
     if ds.State in [dsInsert,dsEdit] then begin
       szPays := Copy(codeISO,1,2);
       // Pas de code iban ou iban incorrect : Le calcul
       szIban := calcIBAN(szPays,calcRIB(szPays,GetControlText('BQ_ETABBQ'),GetControlText('BQ_GUICHET'),GetControlText('BQ_NUMEROCOMPTE'),GetControlText('BQ_CLERIB')));
       if ((szPays = CodeISOFR) or (szPays = CodeISOES)) and (szIban<>GetField('BQ_CODEIBAN')) then begin
          SetField('BQ_CODEIBAN',szIban);
       end;
     End ;
  End else
  Begin
    // Au chargement : Ne fait rien
    if (gbChargement) {or (trim(GetControlText('BQ_CODEIBAN'))<>'')} then exit;
    // Pas de code iban ou iban incorrect : Le calcul
    // Gestion selon les differents type de compte :
    sType := quelType(codeISO,codeTypePays);
    if (sType = 'RIB') or
      ((sType = '') and (codeISO = CodeISOFR) and (VH^.PaysLocalisation = CodeISOFR) ) or
      ((sType = '') and (codeISO = CodeISOES) and (VH^.PaysLocalisation = CodeISOES) ) then
      // RIB
	    sParam := calcRIB(szPays, GetField('BQ_ETABBQ'), GetField('BQ_GUICHET'), GetField('BQ_NUMEROCOMPTE'), GetField('BQ_CLERIB'))
		else if (sType = 'IBA') or (sType = '') then
    begin
	 	 	// IBAN
   		if length(sIban) > 5 then
      	sParam := copy(sIban ,5,length(sIban) - 4);
	   	SetControlText('BQ_CODEIBAN','')
		end;
    if sParam <> '' then
	  	szIban := calcIBAN(szPays, sParam);
    if ((Trim(GetControlText('BQ_CODEIBAN')) = '') or
		   (((szPays = CodeISOFR) or (szPays = CodeISOES)) and
       (szIban <> Trim(GetControlText('BQ_CODEIBAN'))))) then //XMG 14/07/03
		begin
	   {JP 19/03/04 : j'ai l'impression que j'aurais mieux fait de m'abstenir
   		SetControlText('BQ_CODEIBAN',szIban); {JP 23/09/03 : Nécessaire au moins en EAGL}
    	SetField('BQ_CODEIBAN',szIban);
		end;
	end ; //XVI 24/02/2005
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.OnDblClickBQ_CODEPOSTAL(Sender: TObject);
begin
    VerifCodePostal(nil,THEdit(GetControl('BQ_CODEPOSTAL')),THEdit(GetControl('BQ_VILLE')),true);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.OnDblClickBQ_VILLE(Sender: TObject);
begin
    VerifCodePostal(nil,THEdit(GetControl('BQ_CODEPOSTAL')),THEdit(GetControl('BQ_VILLE')),false);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.OnChangeBQ_DEVISE(Sender: TObject);
begin
    SetControlText('EDev',GetControlText('BQ_DEVISE'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.OnDblClickBQ_DIVTERRIT(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
    PaysRegion(THValComboBox(GetControl('BQ_PAYS')),THEdit(GetControl('BQ_DIVTERRIT')),true) ;
{$ELSE}
    PaysRegion(THDBValComboBox(GetControl('BQ_PAYS')),THDBEdit(GetControl('BQ_DIVTERRIT')),true) ;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.OnKeyPressBQ_NUMEROCOMPTE(Sender: TObject; var Key: Char);
begin
    // Authorise uniquement les caractères A-Z et 0-9 et Backspace
    if not(Key in ['a'..'z','A'..'Z','0'..'9',#8]) then Key:=#0 ;
end;

//==================================================
// Autres fonctions de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.LitJourFermeture;
var
    i : Byte;
    C : TCheckBox;
    St : String[7];
begin
    St := GetField('BQ_JOURFERMETUE');
    for i := 1 to 7 do
    begin
        C := TCheckBox(GetControl('Cbj'+IntToStr(i),true));
        C.Checked := (St[i] = 'X');
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.EcritJourFermeture;
var
    i : Byte;
    C : TCheckBox;
    St : String[7];
begin
    St := '';
    for i := 1 to 7 do
    begin
        C := TCheckBox(GetControl('Cbj'+IntToStr(i),true));
        if (C.Checked) then St := St + 'X'
        else St := St + '0';
    end;
    SetField('BQ_JOURFERMETUE',St);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.IsMouvementer;
begin
    if (DS.State = dsInsert) then AvecMvt := False
    else
    begin
        //AvecMvt := ExisteSQL('Select G_GENERAL FROM GENERAUX WHERE G_GENERAL="' + GetField('BQ_GENERAL') + '" And Exists(Select E_GENERAL FROM ECRITURE Where E_GENERAL="' + GetField('BQ_GENERAL') + '")');
        //SG6 28.02.05
        AvecMvt := ExisteSQL('Select E_GENERAL FROM ECRITURE Where E_GENERAL="' + GetField('BQ_GENERAL') + '"');
        OldCodeDev := GetField('BQ_DEVISE');
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_BANQUECP.RIBRenseigner : Boolean;
begin
  Result := ((Trim(GetField('BQ_ETABBQ')) <> '') and (Trim(GetField('BQ_GUICHET')) <> '') and (Trim(GetField('BQ_NUMEROCOMPTE')) <> ''));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_BANQUECP.OkPourChangementDevise(OldCode,NewCode : String) : Boolean;
begin
    if (NewCode = V_PGI.DevisePivot) or (NewCode = V_PGI.DeviseFongible) then
    begin
        result := true;
        exit;
    end;

    if (OldCode <> V_PGI.DevisePivot) and (OldCode <> V_PGI.DeviseFongible) then
    begin
        result := false;
        exit;
    end;

    result := not ExisteSQL('SELECT E_GENERAL FROM ECRITURE WHERE E_GENERAL="' + GetControlText('BQ_GENERAL') + '" AND E_DEVISE<>"' + OldCode + '" AND E_DEVISE<>"' + NewCode + '"');
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_BANQUECP.CodExist : Boolean;
begin
    result := False;
    if (Presence('BANQUECP','BQ_CODE',GetControlText('BQ_CODE'))) then
    begin
        HShowMessage(MESS[10],'','');
        Pages.ActivePageIndex := PIdent;
        SetFocusControl('BQ_CODE');
        result := true;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_BANQUECP.VerifiRibBic : Boolean;
var
   TypeCompte : string;
begin
    result := false;
    TypeCompte := quelType(GetControlText('BQ_PAYS'),GetControlText('BQ_TYPEPAYS'));
    // Cas RIB FQ 19583 BV 01/02/2007
    if (TypeCompte = 'RIB') or
       ((TypeCompte = '') and (codeISO = CodeISOFR) and (VH^.PaysLocalisation = CodeISOFR)) or
       ((TypeCompte = '') and (codeISO = CodeISOES) and (VH^.PaysLocalisation = CodeISOES)) then
    begin
        if (GetControlText('BQ_ETABBQ') = '') then
        begin
            HShowMessage(MESS[4],'','');
            Pages.ActivePageIndex := PIdent;
            SetFocusControl('BQ_ETABBQ');
            exit;
        end;
        if (GetControlText('BQ_GUICHET') = '') then
        begin
            HShowMessage(MESS[5],'','');
            Pages.ActivePageIndex := PIdent;
            SetFocusControl('BQ_GUICHET');
            exit;
        end;
        if (GetControlText('BQ_NUMEROCOMPTE') = '') then
        begin
            HShowMessage(MESS[6],'','');
            Pages.ActivePageIndex := PIdent;
            SetFocusControl('BQ_NUMEROCOMPTE');
            exit;
        end;
        if (GetControlText('BQ_CLERIB') = '') then
        begin
            HShowMessage(MESS[7],'','');
            Pages.ActivePageIndex := PIdent;
            SetFocusControl('BQ_CLERIB');
            exit;
        end;
    end
    else
    begin
        if (GetControlText('BQ_CODEBIC') <> '') then
        begin
            result := true;
            exit;
        end
        else if ((GetControlText('BQ_ETABBQ') = '') or (GetControlText('BQ_GUICHET') = '') or (GetControlText('BQ_NUMEROCOMPTE') = '') or (GetControlText('BQ_CLERIB') = '')) and
                (codeISO <> VH^.PaysLocalisation) then    // FQ 19617 BV 01/02/2007
        begin
            HShowMessage(MESS[8],'','');
            Pages.ActivePageIndex := PIdent;
            SetFocusControl('BQ_CODEBIC');
            exit;
        end;
    end;

    result := true;
end;


{JP 24/09/03 : recherche du code guichet dans la table agence en fonction d'un champ donné
{---------------------------------------------------------------------------------------}
function TOM_BANQUECP.GetCodeGuichet(Chp, Valeur : string) : string;
{---------------------------------------------------------------------------------------}
var
  SQL : String ;
  Q   : TQuery ;
begin
  Result := '';
  SQL    := 'SELECT TRA_GUICHET FROM AGENCE WHERE ' + chp + ' = "' + Valeur + '"';
  Q      := OpenSQL(SQL, True);
  if not Q.EOF then Result := Q.Fields[0].AsString;
  Ferme(Q);
end;

{JP 23/09/03 : constitution du rib et de l'iban lors du changemen des champs constitutifs
{---------------------------------------------------------------------------------------}
procedure TOM_BANQUECP.ConstruireRib(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Q      : TQuery;  
  Banque : string;
begin
  if (TComponent(Sender).Name = 'BQ_BANQUE') then begin
    if GetControlText('BQ_BANQUE') <> '' then begin
      Q := Opensql('SELECT PQ_ETABBQ FROM BANQUES WHERE PQ_BANQUE = "' + GetControlText('BQ_BANQUE') + '"', True);
      if not Q.EOF then Banque := Q.Fields[0].AsString;
      Ferme(Q);
    end;
    {On filtre les agences en fonction de la banque}
    {$IFDEF EAGLCLIENT}
    THValComboBox(GetControl('BQ_AGENCE')).Plus := 'TRA_BANQUE = "' + GetControlText('BQ_BANQUE') + '"';
    {$ELSE}
    THDBValComboBox(GetControl('BQ_AGENCE')).Plus := 'TRA_BANQUE = "' + GetControlText('BQ_BANQUE') + '"';
    {$ENDIF}
  end

  else if (TComponent(Sender).Name = 'BQ_AGENCE') and EstComptaTreso {FQ 15933 : JP 24/05/05} then begin
    {$IFDEF EAGLCLIENT}
    {JP 19/10/05 : FQ 16888 : Le comportement en eAgl est différént de 2/3 et exécute le ConstruireRib
                   après le chargement de l'enregistrement, ce qui met en mode modification et grise le
                   bouton supprimé}
    if not gbChargement then begin
    {$ENDIF EAGLCLIENT}
      {La combo a changé : on met à jour le champ BQ_DOMICILIATION}
      SetField('BQ_DOMICILIATION', THValComboBox(GetControl('BQ_AGENCE')).Text);
      {Si le code Iso est FR, on à jour le code guichet, en fonction de l'agence choisie}
      //RR if ((VH^.PaysLocalisation=CodeISOEs) and (ds.state in [dsInsert,dsEdit]) and ((codeISO = CodeISOFR) or (codeISO = CodeISOES))) or
      if (ds.state in [dsInsert,dsEdit]) or (not gbChargement ) then //XVI 24/02/2005
      begin
        FParamPays := ExisteParamPays(codeISO, CodeTypePays);
        if (FParamPays
           and ( quelType(codeISO,CodeTypePays) = 'RIB' )
           and (chercheLONG('GUICHET',codeISO,CodeTypePays) > 0 ))
           or (not FParamPays
           and (((codeISO = CodeISOFR) and (VH^.PaysLocalisation = CodeISOFR)) or ((codeISO = CodeISOES) and (VH^.PaysLocalisation = CodeISOES)))) then
           SetField('BQ_GUICHET', GetCodeGuichet('TRA_AGENCE', GetControlText('BQ_AGENCE')));
      end;
    {$IFDEF EAGLCLIENT}
    end;
    {$ENDIF EAGLCLIENT}
  end;

  //if ((VH^.PaysLocalisation=CodeISOEs) and (not (ds.state in [dsInsert,dsEdit])) or ((codeISO <> CodeISOFR) and (codeISO <> CodeISOES) and (codeISO <> ''))) or
  //   ((VH^.PaysLocalisation<>CodeISOES) and (gbChargement or ((codeISO <> CodeISOFR) and (codeISO <> '')))) then Exit; //XVI 24/02/2005
  //if (ds.state in [dsInsert,dsEdit]) or (not gbChargement ) then exit ;
  if not (ds.state in [dsInsert,dsEdit]) or gbChargement then exit ;

  {JP 18/05/06 : FQ 17722 : Si l'on n'est pas en Mode Treso, banque a de fortes chances d'être vide}
  if (TComponent(Sender).Name = 'BQ_BANQUE') and (Banque <> '') and EstComptaTreso  or
     (TComponent(Sender).Name = 'BQ_BANQUE') and (Banque <> '') and (VH^.PaysLocalisation = CodeIsoBE) then //SDA le 27/12/2007
  begin
     if (existeParamPays(codeISO,CodeTypePays)
           and ( quelType(codeISO,CodeTypePays) = 'RIB' )
           and (chercheLONG('ETABBQ',codeISO,CodeTypePays) > 0 ))
           or (not(existeParamPays(codeISO,CodeTypePays))
           and (((codeISO = CodeISOFR) and (VH^.PaysLocalisation = CodeISOFR)) or ((codeISO = CodeISOES) and (VH^.PaysLocalisation = CodeISOES)))) then
         {Si le code Iso est FR, on met à jour le code banque, en fonction de la banque choisie}
          SetField('BQ_ETABBQ', Banque);
  end;
end;

{JP 24/09/03 : vérifie la validité du compte général saisi.
  1/ Il ne doit pas être vide
  2/ Il ne doit pas être utilisé par un autre compte bancaire
  3/ Il doit appartenir à la table GENERAUX
  4/ S'il s'agit d'un compte divers, il doit appartenir à CLIENSSOC
{---------------------------------------------------------------------------------------}
function TOM_BANQUECP.ValideGen : Boolean;
{---------------------------------------------------------------------------------------}
{$IFDEF TRESO}
var
  Q : TQuery;
  S : string;
{$ENDIF}
begin
  Result := False;

  if Trim(GetControlText('BQ_GENERAL')) = '' then Exit;

  if EstMultiSoc then begin
    if ExisteSQL('SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_GENERAL = "' + GetControlText('BQ_GENERAL') + '"' +
                 ' AND BQ_CODE <> "' + GetControlText('BQ_CODE') + '" AND BQ_NODOSSIER = "' + FNoDossier + '"') then Exit;

    {$IFDEF TRESO}
    {On s'assure qu'il s'agit bien d'un compte de banque, de caisse ou divers}
    Q := OpenSQL('SELECT G_NATUREGENE FROM GENERAUX WHERE G_GENERAL = "' + GetControlText('BQ_GENERAL') + '" ' +
                  'AND G_NATUREGENE IN ("CAI", "BQE", "DIV")', True);
    if not Q.EOF then begin
      {S'il s'agit d'un compte divers, on s'assure qu'il appartient bien à CLIENSSOC}
      if Q.FindField('G_NATUREGENE').AsString = 'DIV' then begin
        S := VarToStr(GetField('BQ_NODOSSIER'));
        S := GetInfosFromDossier('DOS_NODOSSIER', S, 'DOS_NOMBASE');
        if (S = '') or (
          not ExisteSQL('SELECT CLS_GENERAL FROM ' + GetTableDossier(S, 'CLIENSSOC') +
                         ' WHERE CLS_GENERAL = "' + GetControlText('BQ_GENERAL') + '"')) then begin
          Ferme(Q);
          Exit;
        end;
      end;
      Ferme(Q);
    end
    else begin
      Ferme(Q);
      Exit;
    end;
    {$ENDIF}
  end
  else
  begin
    if ExisteSQL('SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_GENERAL = "' + GetControlText('BQ_GENERAL') + '"' +
                 ' AND BQ_CODE <> "' + GetControlText('BQ_CODE') + '"') then Exit;

    {$IFDEF TRESO} {JP 25/03/04 : En compta, en création de compte généraux, le compte général n'existe
                                  pas encore inséré dans la table au moment de la création du RIB}
    {JP 25/02/04 : La requête renvoyé par LookUpValueExist ne correspond pas au besoin
                   if not LookUpValueExist(GetControl('BQ_GENERAL')) then Exit;}
      if not ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = "' + GetControlText('BQ_GENERAL') + '" ' +
                       'AND G_NATUREGENE IN ("CAI", "BQE")') then Exit;
    {$ENDIF}
  end;

  Result := True;
end;

{JP 26/09/03 : On s'assure que l'agence est renseignée}
{---------------------------------------------------------------------------------------}
function TOM_BANQUECP.ValideAgence : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := False;
  if Trim(GetControlText('BQ_DOMICILIATION')) = '' then Exit;
  Result := True;
end;

{JP 26/09/03 : On s'assure que la banque est renseignée}
{---------------------------------------------------------------------------------------}
function TOM_BANQUECP.ValideBanque : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := False;
  if Trim(GetControlText('BQ_BANQUE')) = '' then Exit;
  Result := True;
end;

{JP 26/09/03 : Ouverture de la fiche des agences bancaires}
{---------------------------------------------------------------------------------------}
procedure TOM_BANQUECP.BAgenceOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_Agence('TR','TRAGENCE','',GetControlText('BQ_AGENCE'),'');
  {On raffraîchit la tablette}
  {$IFDEF EAGLCLIENT}
  THValComboBox(GetControl('BQ_AGENCE')).ReLoad;
  {$ELSE}
  THDbValComboBox(GetControl('BQ_AGENCE')).ReLoad;
  {$ENDIF}
end;

{27/09/05 : FQ 15585 : Surcharge de la suppression, pour empécher la fermeture de la fiche et la
            mettre ensuite en mode création
{---------------------------------------------------------------------------------------}
procedure TOM_BANQUECP.BDleteClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  MonoF : Boolean;
begin
  {Lequel a un rôle sur l'activation de certaines zones s'il est renseigné. Comme on va
   ensuite passer en mode insertion, on le vide pour que toutes les zones soient actives}
  Lequel := '';

  {Pour que la fiche ne se ferme pas après la suppression}
  MonoF := TFFiche(Ecran).MonoFiche;
  TFFiche(Ecran).MonoFiche := False;

  {Reprise du traitement implémenté dans l'unité Fiche.Pas}
  if Assigned(FSuppressionClick) then FSuppressionClick(Sender);
  TFFiche(Ecran).MonoFiche := MonoF;

  {Insertion d'un nouvel en registrement : valable pour la compta, sinon un rib d'attente est créé}
  TFFiche(Ecran).BInsert.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_BANQUECP.OnDossierChange;
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  FNoDossier := GetField('BQ_NODOSSIER');
  {Mise à jour du filtre sur les comptes généraux}
  S := GetInfosFromDossier('DOS_NODOSSIER', FNoDossier, 'DOS_NOMBASE');
  BQ_GENERAL.Plus := '(G_NATUREGENE = "BQE" AND G_GENERAL NOT IN (SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_NODOSSIER = "' + FNoDossier + '")) OR ' +
                     '(G_NATUREGENE = "DIV" AND G_GENERAL IN (SELECT CLS_GENERAL FROM ' + GetTableDossier(S, 'CLIENSSOC') + ') AND ' +
                     'G_GENERAL NOT IN (SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_NODOSSIER = "' + FNoDossier + '"))';
end;

{---------------------------------------------------------------------------------------}
procedure TOM_BANQUECP.OnNatureChange;
{---------------------------------------------------------------------------------------}
begin
  if (GetControlText('BQ_NATURECPTE') = tcb_Titre) or (GetControlText('BQ_NATURECPTE') = tcb_Courant) then
    SetField('BQ_DEVISE', V_PGI.DevisePivot);
end;

{---------------------------------------------------------------------------------------}
function TOM_BANQUECP.TesteDevise : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := (GetControlText('BQ_NATURECPTE') = tcb_Titre) or (GetControlText('BQ_NATURECPTE') = tcb_Courant);
  if Result then Result := GetControlText('BQ_DEVISE') = V_PGI.DevisePivot
            else Result := True;
end;

{---------------------------------------------------------------------------------------}
function TOM_BANQUECP.TesteNature : Boolean;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  {Un compte bancaire doit être de nature bancaire et un compte divers de nature titre ou courant}
  Q := OpenSQL('SELECT G_NATUREGENE FROM GENERAUX WHERE G_GENERAL = "' +
               GetControlText('BQ_GENERAL') + '"', True);
  try
    if Q.EOF then
      Result := (GetControlText('BQ_NATURECPTE') = tcb_Bancaire) {Par défaut}
    else begin
      if Q.FindField('G_NATUREGENE').AsString = 'DIV' then
        Result := (GetControlText('BQ_NATURECPTE') = tcb_Titre) or (GetControlText('BQ_NATURECPTE') = tcb_Courant)
      else
        Result := (GetControlText('BQ_NATURECPTE') = tcb_Bancaire);
    end;
  finally
    Ferme(Q);
  end;
end;
{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 21/11/2006
Modifié le ... :   /  /    
Description .. : Fonction permettant de replacer les champs Banque,
Suite ........ : Guichet,Cle, Numéro de compte à leur taille et leurs
Suite ........ : positions initiales dans la fiche
Mots clefs ... : 
*****************************************************************}
procedure TOM_BANQUECP.AfficheEcran;
begin
  // Permet d'activer ou non le type pays
  AfficheTYPEPAYS(ecran,ds,'BQ_',codeISO) ;

  if (codeISO = CodeISOES) and (VH^.PaysLocalisation = CodeISOES) then
    // Affichage Specifique pour l'Espagne
    AfficheSpecifES
  else
    // Affichage Standard
    Affiche(ecran,codeISO,codeTypePays,'BQ_',3);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/11/2006
Modifié le ... :   /  /
Description .. : Affichage Specifique pour l'espagne
Mots clefs ... :
*****************************************************************}
procedure TOM_BANQUECP.AfficheSpecifES;
var
{$IFNDEF EAGLCLIENT}
  ZoneRIB, ZoneCP : THDBEdit;
{$ELSE}
  ZoneRIB, ZoneCP : THEdit;
{$ENDIF}
  LblZoneRib : THLabel;
  ZoneLeft, OrdreTab : Integer;
begin

  // On fait un affichage defaut pour s'assurer que les champs ont bien le formatage attendu.
  AfficheDefault(ecran,'BQ_',codeISO, true,2) ;
  // On desactive l'IBAN FQ 19465 BV 11/01/2007
  SetControlEnabled('BQ_CODEIBAN',false);
  if (codeISO = CodeISOFR) or (CodeISO=CodeISOES) or (CodeISO='') {??} then
  begin
    // enable des control
    {$IFDEF EAGLCLIENT}
      ZoneRIB:=THEdit(GetControl('BQ_ETABBQ')) ;
    {$ELSE}
      ZoneRIB:=THDBEdit(GetControl('BQ_ETABBQ')) ;
    {$ENDIF EAGLCLIENT}
    if Assigned(ZoneRIB) then
    Begin
      ZoneRib.Enabled:=(TFFiche(Ecran).TypeAction<>taConsult) and (not EstComptaTreso) ; //XVI 24/02/2005
      ZoneRib.EditMask:=FormatZonesRib(CodeISO,'BQ') ;
      ZoneRib.MaxLength := 4 ;
      {$IFNDEF EAGLCLIENT}
      //  ZoneRib.Field.EditMask:=FormatZonesRib(CodeISO,'BQ') ;
      {$ENDIF}
    End ;

    {$IFDEF EAGLCLIENT}
      ZoneRIB:=THEdit(GetControl('BQ_GUICHET')) ;
    {$ELSE}
      ZoneRIB:=THDBEdit(GetControl('BQ_GUICHET')) ;
    {$ENDIF EAGLCLIENT}
    if Assigned(ZoneRIB) then
    Begin
      ZoneRib.Enabled:=(TFFiche(Ecran).TypeAction<>taConsult) and (not EstComptaTreso) ; //XVI 24/02/2005 ;
      ZoneRib.EditMask:=FormatZonesRib(CodeISO,'GU') ;
      ZoneRib.MaxLength := 4 ;
      {$IFNDEF EAGLCLIENT}
      //  ZoneRib.Field.EditMask:=FormatZonesRib(CodeISO,'GU') ;
      {$ENDIF}
    End ;

    {$IFDEF EAGLCLIENT}
      ZoneRIB:=THEdit(GetControl('BQ_NUMEROCOMPTE')) ;
    {$ELSE}
      ZoneRIB:=THDBEdit(GetControl('BQ_NUMEROCOMPTE')) ;
    {$ENDIF EAGLCLIENT}
    if Assigned(ZoneRIB) then
    Begin
      ZoneRib.Enabled:=(TFFiche(Ecran).TypeAction<>taConsult) ;
      ZoneRib.EditMask:=FormatZonesRib(CodeISO,'CP') ;
      ZoneRib.MaxLength := 10 ;
      {$IFNDEF EAGLCLIENT}
      //  ZoneRib.Field.EditMask:=FormatZonesRib(CodeISO,'CP') ;
      {$ENDIF}
    End ;

    {$IFDEF EAGLCLIENT}
      ZoneRIB:=THEdit(GetControl('BQ_CLERIB')) ;
    {$ELSE}
      ZoneRIB:=THDBEdit(GetControl('BQ_CLERIB')) ;
    {$ENDIF EAGLCLIENT}
    if Assigned(ZoneRIB) then
    Begin
      ZoneRib.Enabled:=(TFFiche(Ecran).TypeAction<>taConsult) ;
      ZoneRib.EditMask:=FormatZonesRib(CodeISO,'DC') ;
      ZoneRib.MaxLength := 2 ;
      {$IFNDEF EAGLCLIENT}
      //  ZoneRib.Field.EditMask:=FormatZonesRib(CodeISO,'DC') ;
      {$ENDIF}
    End ;

    // preset des control
    if (ds.state in [dsInsert,dsEdit]) then
    begin
      if (quelType(codeISO,GetControlText('BQ_TYPEPAYS')) <> 'BBA') then
        SetField('R_CODEIBAN',CodeISO);
    end;
    {$IFDEF EAGLCLIENT}
    ZoneRIB:=THEdit(GetControl('BQ_CLERIB')) ;
    ZoneCP:=THEdit(GetControl('BQ_NUMEROCOMPTE')) ;
    {$ELSE}
    ZoneRIB:=THDBEdit(GetControl('BQ_CLERIB')) ;
    ZoneCP:=THDBEdit(GetControl('BQ_NUMEROCOMPTE')) ;
    {$ENDIF EAGLCLIENT}
    if (Assigned(ZoneCP)) and (Assigned(ZOneRIB)) then
    Begin
      {$IFDEF EAGLCLIENT}
        With THEdit(GetControl('BQ_GUICHET')) do Begin
       {$ELSE}
        With THDBEdit(GetControl('BQ_GUICHET')) do Begin
      {$ENDIF EAGLCLIENT}
        ZoneLeft:=Left+Width+20 ;
        OrdreTab:=TabOrder ;
      End ;
       if CodeISO=CodeISOES then
       begin //Si Espagne, on tourne la position des champs....
        ZoneRIB.Left:=ZoneLeft ;
        ZoneRIB.TabOrder:=OrdreTab+1 ;
        ZoneCP.Left:=ZoneRIB.Left+ZoneRIB.Width+20 ;
        ZoneCP.TabOrder:=ZoneRIB.TabOrder+1 ;
      End
       else
      begin //Pour tout autre Pays on re-établir l'affichage Français....
         ZoneCP.Left:=ZoneLeft ;
         ZoneCP.TabOrder:=OrdreTab+1 ;
         ZoneRIB.Left:=ZoneCP.Left+ZoneCP.Width+20 ;
         ZoneRIB.TabOrder:=ZoneCP.TabOrder+1 ;
      End ;
      lblZoneRib:=THLabel(GetControl('TBQ_NUMEROCOMPTE')) ;
      lblZoneRIB.Left:=LblZoneRIB.FocusControl.Left ;
      lblZoneRib:=THLabel(GetControl('TBQ_CLERIB')) ;
      lblZoneRIB.Left:=LblZoneRIB.FocusControl.Left ;
    end
    else
    Begin
      if ((codeISO = 'FR') or (codeISO = '')) then
      begin
        {JP 25/03/04 : Suppression de la directive TRSYNCHRO}
        SetControlEnabled('BQ_ETABBQ', not EstComptaTreso);
        SetControlEnabled('BQ_GUICHET',not EstComptaTreso);
        // enable des control
        SetControlEnabled('BQ_NUMEROCOMPTE',true);
        SetControlEnabled('BQ_CLERIB',true);
        SetControlEnabled('BQ_CODEIBAN',false);
        // preset des control
        if (not gbChargement) then // On est pas en train de changer un enreg
        begin
          //SetControlText('BQ_CODEIBAN',codepaysdeiso('FR'));
          if (quelType(codeISO,GetControlText('BQ_TYPEPAYS')) <> 'BBA') then
            SetField('BQ_CODEIBAN',CodeISO) ; //codepaysdeiso('FR')); //XMG 14/07/03
        end;
      end
      else
      begin
        // enable des control
        SetControlEnabled('BQ_ETABBQ',false);
        SetControlEnabled('BQ_GUICHET',false);
        SetControlEnabled('BQ_NUMEROCOMPTE',false);
        SetControlEnabled('BQ_CLERIB',false);
        SetControlEnabled('BQ_CODEIBAN',true);
      end;
      // preset des control
      if (not gbChargement) then // On est pas en train de changer un enreg
      Begin
        SetField('BQ_ETABBQ','');
        SetField('BQ_GUICHET','');
        SetField('BQ_NUMEROCOMPTE','');
        SetField('BQ_CLERIB','');
        if  (quelType(codeISO,GetControlText('BQ_TYPEPAYS')) <> 'BBA') then
          SetField('BQ_CODEIBAN',codeISO);
      end;
    end;
  end;
end;

{JP 07/06/07 : FQ 19591 : Gestion des modèles de bordereau : documents ou états
{---------------------------------------------------------------------------------------}
procedure TOM_BANQUECP.InitBordereau;
{---------------------------------------------------------------------------------------}
var
  bEtat : Boolean;
begin
  bEtat := GetParamSocSecur('SO_CPDOCAVECETAT', True);
  {$IFDEF EAGLCLIENT}
  if bEtat then begin
    (GetControl('BQ_LETTREPRELV') as THValComboBox).DataType := 'CPMODELESENCADECA';
    (GetControl('BQ_LETTRELCR'  ) as THValComboBox).DataType := 'CPMODELESENCADECA';
    (GetControl('BQ_LETTREVIR'  ) as THValComboBox).DataType := 'CPMODELESENCADECA';
    (GetControl('BQ_LETTRECHQ'  ) as THValComboBox).DataType := 'CPMODELESENCADECA';
    (GetControl('BQ_LETTREPRELV') as THValComboBox).Plus := 'CLP';
    (GetControl('BQ_LETTRELCR'  ) as THValComboBox).Plus := 'CLT';
    (GetControl('BQ_LETTREVIR'  ) as THValComboBox).Plus := 'CLV';
    (GetControl('BQ_LETTRECHQ'  ) as THValComboBox).Plus := 'CLC';
  end else begin
    (GetControl('BQ_LETTREPRELV') as THValComboBox).DataType := 'TTMODELELETTREPRE';
    (GetControl('BQ_LETTRELCR'  ) as THValComboBox).DataType := 'TTMODELELETTRETRA';
    (GetControl('BQ_LETTREVIR'  ) as THValComboBox).DataType := 'TTMODELELETTREVIR';
    (GetControl('BQ_LETTRECHQ'  ) as THValComboBox).DataType := 'TTMODELELETTRECHQ';
  end;
  {$ELSE}
  if bEtat then begin
    (GetControl('BQ_LETTREPRELV') as THDBValComboBox).DataType := 'CPMODELESENCADECA';
    (GetControl('BQ_LETTRELCR'  ) as THDBValComboBox).DataType := 'CPMODELESENCADECA';
    (GetControl('BQ_LETTREVIR'  ) as THDBValComboBox).DataType := 'CPMODELESENCADECA';
    (GetControl('BQ_LETTRECHQ'  ) as THDBValComboBox).DataType := 'CPMODELESENCADECA';
    (GetControl('BQ_LETTREPRELV') as THDBValComboBox).Plus := 'CLP';
    (GetControl('BQ_LETTRELCR'  ) as THDBValComboBox).Plus := 'CLT';
    (GetControl('BQ_LETTREVIR'  ) as THDBValComboBox).Plus := 'CLV';
    (GetControl('BQ_LETTRECHQ'  ) as THDBValComboBox).Plus := 'CLC';
  end else begin
    (GetControl('BQ_LETTREPRELV') as THDBValComboBox).DataType := 'TTMODELELETTREPRE';
    (GetControl('BQ_LETTRELCR'  ) as THDBValComboBox).DataType := 'TTMODELELETTRETRA';
    (GetControl('BQ_LETTREVIR'  ) as THDBValComboBox).DataType := 'TTMODELELETTREVIR';
    (GetControl('BQ_LETTRECHQ'  ) as THDBValComboBox).DataType := 'TTMODELELETTRECHQ';
  end;
  {$ENDIF EAGLCLIENT}
end;

{$IFDEF TRESO}
{JP 11/05/07 : FQ TRESO 10462 : comme on n'a pas la fiche des généraux en Tréso, il faut cacher les
               boutons zoom et nouveau => on fait le lookup à la main
{---------------------------------------------------------------------------------------}
procedure TOM_BANQUECP.GeneralElipsisClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Plus : string;
  Doss : string;
begin
  Plus := 'G_FERME <> "X" AND ';
  if EstTresoMultiSoc then begin
    {On récupère les comptes bancaires et les comptes courants}
    Doss := GetInfosFromDossier('DOS_NODOSSIER', FNoDossier, 'DOS_NOMBASE');
    Plus := Plus + '(G_NATUREGENE = "BQE" AND G_GENERAL NOT IN (SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_NODOSSIER = "' + FNoDossier + '")) OR ' +
            '(G_NATUREGENE = "DIV" AND G_GENERAL IN (SELECT CLS_GENERAL FROM ' + GetTableDossier(Doss, 'CLIENSSOC') + ') AND ' +
            'G_GENERAL NOT IN (SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_NODOSSIER = "' + FNoDossier + '"))';
  end
  else
    {Un compte général ne peut être associé qu'à un compte bancaire}
    Plus := Plus + 'G_NATUREGENE = "BQE" AND G_GENERAL NOT IN (SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_NODOSSIER = "' + FNoDossier + '")';

  LookupList(BQ_GENERAL, TraduireMemoire('Liste des généraux'), 'GENERAUX', 'G_GENERAL', 'G_LIBELLE',
             Plus, 'G_GENERAL', True, 0); {0 permet de cacher les deux boutons}
end;
{$ENDIF TRESO}

{JP 05/07/07 : FQ 20319 : Calcul de l'IBAN en Direct, avant le OnUpdateRecord
{---------------------------------------------------------------------------------------}
procedure TOM_BANQUECP.MajCodeIban;
{---------------------------------------------------------------------------------------}
var
  szPays, szIban : string;
begin
  if not(DS.State in [dsInsert, dsEdit]) then Exit;
  {05/07/07 : Si on a mis en place un paramétrage des RIB / IBAN / BBAN => on sort}
  if FParamPays then Exit;
  
  szPays := Copy(codeISO,1,2);
  {Calcul de l'IBAN}
  szIban := calcIBAN(szPays,calcRIB(szPays,GetControlText('BQ_ETABBQ'),GetControlText('BQ_GUICHET'),GetControlText('BQ_NUMEROCOMPTE'),GetControlText('BQ_CLERIB')));
  if ((szPays = CodeISOFR) or (szPays = CodeISOES)) and (szIban <> GetField('BQ_CODEIBAN')) then 
    SetField('BQ_CODEIBAN',szIban);
end;

//================================================================================
// Initialization
//================================================================================
Initialization
    registerclasses([TOM_BANQUECP]);
end.
