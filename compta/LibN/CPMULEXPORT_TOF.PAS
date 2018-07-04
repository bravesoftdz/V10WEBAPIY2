{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 26/04/2004
Modifié le ... : 16/05/2007
Description .. : Source TOF de la FICHE : CPMULEXPORT ()
Mots clefs ... : TOF;CPMULEXPORT
*****************************************************************}
{-------------------------------------------------------------------------------------
  Version    |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
               26/04/04  VL   Création de l'unité
               16/06/05  JP   VCOM 1000 BECM
               16/06/05       FQ 16005 : Zoom Pièce
 6.50.001.005  16/06/05  JP   FQ 15991 : VerifEcheNegatives
 6.50.001.005  20/06/05  JP   FQ 15500 : Gestion des montants en devises
               20/06/05       Utilisation d'un seul modèle d'export pour GeFactofrance
 6.50.001.005  21/06/05  JP   FQ 13696 : le numéro de quittance est géré dans le code et non plus
                              dans le générateur de fichier afin d'être utilisé dans les états
 8.00.001.015  14/05/07  JP   FQ 18132,  DISTINCT sur la Référence
 8.00.001.015  16/05/07  JP   FQ 19142, Incrémentation du numéro de quittance pour GEFACTOFRANCE
 8.10.006.006  07/02/07  JP   FQ 22047 : Disctinction entre le nombre d'avoirs et de factures
--------------------------------------------------------------------------------------}

unit CPMULEXPORT_TOF;

interface

uses
    StdCtrls,
    {$IFDEF VER150} Variants, {$ENDIF}    
    Windows,
    Controls,
    Classes,
    uTob,
{$IFDEF EAGLCLIENT}
    eMul,
    MaineAGL,
    UtileAGL, // LanceEtat
{$ELSE}
    db,
    {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
    DBGrids,
    Grids,
    HDB,
    mul,
    FE_Main,
    EdtREtat, // LanceEtat
{$ENDIF}
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    HEnt1,
    HMsgBox,
    HTB97,
    Ent1,
    HQry,
    Saisie, // TrouveEtLanceSaisie
    SaisBor, // LanceSaisieFolio
    HPanel,
    UTOF,
    CPTIERS_TOM,
    EdtEtat,
    LookUp,
    uLibCFONB; // SDA le 17/12/2007 version belge

procedure CP_LanceFicheExport(Arg : string); {FQ 15333}

type
    TOF_CPMULEXPORT = class(TOF)
        procedure OnNew; override;
        procedure OnDelete; override;
        procedure OnUpdate; override;
        procedure OnLoad; override;
        procedure OnArgument(S : string); override;
        procedure OnDisplay; override;
        procedure OnClose; override;
        procedure OnCancel; override;
    private
        ObjEXPORT  : TObjCFONB; //SDA le 17/12/2007
        GetInfoParamSoc : boolean; //SDA le 17/12/2007
        Libelle,Adresse1,Adresse2,CodePostal,CodeNif,NumeroTVA,Ville : string; // SDA le 17/12/2007
        E_EXERCICE : THValComboBox;
        E_DATECOMPTABLE, E_DATECOMPTABLE_, E_NUMCFONB, E_JOURNAL : THEdit;
        cboModeleEtat, T_NATUREAUXI, CBOMODELEGFF : THValComboBox;
        Q : THQuery;
        DATE1 : THEdit;
        bCanUpdate : Boolean;

{$IFDEF EAGLCLIENT}
        FListe : THGrid;
{$ELSE}
        FListe : THDBGrid;
{$ENDIF}
        {FQ 15333 : Getter sur le type d'export}
        function GetTypeExport : string;

        // parametrage des etat/fichier
        procedure BPARAMEXPORTClick(Sender : TObject);
        procedure BPARAMBORDERAUClick(Sender : TObject);

        procedure E_EXERCICEChange(Sender : TObject);

        procedure FListeDblClick(Sender : TObject);

        procedure BOuvrirClick(Sender : TObject);

        procedure DoExport(T : Tob);
        procedure DoExportCGA(T : Tob);

        // selection des modeles
        procedure cboModeleEtatOnChange(Sender : TObject);

        // Export CGA
        procedure DateOnExit(Sender : TObject);
        function GetMinMaxCompte(szTable, szCol, szDefaut : string) : string;
        function GetWhereColl(szNatureTiers : string) : string;

        procedure OnElipsisClickE_NumCfonB(Sender : TObject);
        procedure AuxiElipsisClick(Sender : TObject);

        procedure CreateObjEXPORT; // SDA le 17/12/2007
        procedure ChargeTypesExport(strPays : string); // SDA le 17/12/2007
        function  GetValeur     (NomChamp : string) : Variant; // SDA le 17/12/2007
        procedure GetInfosParamSoc (var TOBFille : TOB); // SDA le 17/12/2007
        function  ChargeMaTobTraitement (NomListe: string; TobVue : Tob) : Tob; // SDA le 21/01/2008


    public
      {JP 02/05/05 : FQ 15333 : Pour savoir s'il s'agit d'un appel client ou fournisseur}
      ClientOk : Boolean;
      {Du fait de l'extension du nombre de menus, Régis suggère de mettre une Combo plutôt que
       de rajouter des radios à RGPMODELEEXPORT}
      CbType   : THValComboBox;

      {FQ 15334 : Gestion de la banque prévisionnelle pour les exports VCOM}
      CbBanque : THValComboBox;
      NumCompte : string;
      CodeIban  : string;
      CodeRib   : string;
      {FQ 15513 : Référence du donneur d'ordre}
      RefDonneur : string;

      {FQ 15537 : Cumul des débits/Crédits des lignes sélectionnées}
      TotalCredit : Double;
      TotalDebit  : Double;

      {FQ 15537 : Caption d'origine de l'écran}
      CaptionOrigine : string;

      {JP 11/05/05 : FQ 15537 : Surcharge de la sélection total pour l'affichage des soldes}
      FBSelectAllClick : TNotifyEvent;

      {JP 10/05/05 : Nettoyage du OnArgument : Initialisation des composants}
      procedure InitComposants;
      {JP 10/05/05 : Nettoyage du OnArgument : Initialisation des évènements}
      procedure InitEvenements;
      {JP 10/05/05 : Nettoyage du OnArgument : Initialisation des variables et des contrôles}
      procedure InitVariables;

      {JP 20/06/05 : Ajout d'une combo pour GE FactoFrance}
      procedure GeFactoChange(Sender : TObject);
      {JP 20/09/05 : FQ 16627 : Gestion de la combo type d'export GEFACTOFRANCE}
      procedure MajZonesGeFactoFrance;
      {Changement du type d'export : CGA, VCOM ...}
      procedure CbTypeOnChange(Sender : TObject);
      {FQ 15334 : Au changement de la banque prévisionnelle pour les exports VCOM}
      procedure CbBanquePreviChange(Sender : TObject);
      {Vérification qu'il n'y a pas d'échénaces négatives}
      function  VerifEcheNegatives(var T : TOB) : Boolean;
      {JP 10/05/05 : FQ 15537 : Gestion du solde des lignes sélectionnées dans le caption de la fiche}
      procedure FListeFlipSelection(Sender: TObject);
      {JP 10/05/05 : FQ 15537 : Affichage du solde des lignes sélectionnées dans le caption de la fiche}
      procedure AfficheSoldeSelection;
      {JP 10/05/05 : FQ 15537 : Calcul du solde des lignes sélectionnées }
      function CalculSoldeSelection : Double;
      {Boite de dialogue et génération du fichier ASCII}
      function  LancerDialog(var T : TOB) : Boolean;
      {JP 09/05/05 : FQ 15552 et 15738 : Mise à jour du champ E_CFONBOK après la génération du fichier}
      procedure MajE_CFONBOK(T : TOB);
      {JP 09/05/05 : FQ 15513 : Affichage de la fenêtre de saisie des références du donneur d'ordre}
      function GererRefVCOM(var T : TOB) : Boolean;
      {YMO}
      function GererRefGEF(var T : TOB) : String;
      {JP 10/05/05 : FQ 15537 : Réinitialisation des variables avec ou sans le Caption}
      procedure InitGestionSolde(AvecCaption : Boolean = True);
      {JP 11/05/05 : FQ 15537 : Surcharge de la sélection total pour l'affichage des soldes}
      procedure BSelectAllClick(Sender : TObject);
      {JP 21/09/05 : FQ 16626 : Affectation de la DBListe du Mul}
      procedure SetLaListe(NomListe : string);
      {JP 22/09/05 : FQ 16623 - 16625 : Récupération des états par défaut pour GEFactoFrance}
      procedure RecupEtatParDefaut(var BOR, QUIT : string);
      {JP 23/09/05 : Pour filtrer la tob en fonction de la nature de pièce pour les impressions}
      procedure SetTobImpression(tDepart : TOB; var tEtat : TOB; Nat : string);

      property TypeExport : string read GetTypeExport;
    end;

implementation

uses
    {$IFDEF MODENT1}
    CPProcGen,
    ULibExercice,
    CPTypeCons,
    {$ENDIF MODENT1}
    Dialogs,
    mc_erreur,
    LP_Base,
    LP_View,
    LP_PARAM,
    ParamSoc,
    Messages, UTobDebug,
    UtilPgi, {CleTelephone}
    ExtCtrls, {TPanel}
    Commun, {SetPlusBanqueCp}
    CPREFDONNEURORD_TOF,{JP 09/05/05 : FQ 15513 : CPRefDonneurOrd}
    UTofMulParamGen, {13/04/07 YMO F5 sur Auxiliaire }
  {$IFDEF CISXPGI}
    Echg_Code,
  {$ENDIF CISXPGI}
    CPSAISIEPIECE_TOF; //SDA le 17/12/2007

const
  {FQ 15333 : différentes valeurs de la tablette CPEXPORTCCMP}
  cex_DIV  = 'DIV';
  cex_VCOM = 'VCO';
  cex_CGA  = 'EGA'; {E et non C pour l'ordre dans la tablette qui est par ordre alphabétique}
  cex_GEF  = 'GEF';

//////////////////////////////////////////////////////////////////////////////////////////

procedure CP_LanceFicheExport(Arg : string);{FQ 15333}
begin
    AGLLanceFiche('CP', 'CPMULEXPORT', '', '', Arg);
end;

//////////////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPMULEXPORT.OnNew;
begin
    inherited;
end;

//////////////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPMULEXPORT.OnDelete;
begin
    inherited;
end;

//////////////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPMULEXPORT.OnUpdate;
begin
  inherited;
  {JP 10/05/05 : FQ 15537 : Réinitialisation des variables et du Caption de la fiche}
  InitGestionSolde;
end;

//////////////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPMULEXPORT.OnLoad;
var
    szExoV8, szDate, szWhere, szGroupBy, szOrderBy, szSelect : string;
    szCommun1, szCommun2, szAuxi, szAuxi_, szGen, szGen_ : string;
    dtDate : TDateTime;
begin
  inherited;

  //SDA le 17/12/2007 version belge
  {$IFDEF CISXPGI}
  if VH^.PaysLocalisation = CodeIsoDuPays('BEL') then
  begin
    SetLaListe('CPLISTECISXCFONB');
    exit;
  end;
  {$ENDIF}
  //Fin SDA le 17/12/2007

  {Valeurs par défaut}
  SetControlText('XX_WHERE', '');
  SetControlText('XX_GROUPBY', '');
  SetControlText('XX_ORDERBY', '');
  SetControlText('XX_SELECT', '');

  {Export CGA}
  if (TypeExport = cex_CGA) then
  begin

    szGroupBy := 'E_JOURNAL, E_ETABLISSEMENT, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_QUALIFPIECE, E_NATUREPIECE, E_GENERAL, E_AUXILIAIRE, E_DEVISE, E_MODESAISIE, E_MODEPAIE, E_LIBELLE, ';
    szGroupBy := szGroupBy + 'E_REFINTERNE, E_REFEXTERNE, E_REFLIBRE, T_AUXILIAIRE, T_NATUREAUXI, T_LIBELLE, T_ENSEIGNE, T_SIRET, T_NIF, T_LIBELLE, T_ADRESSE1, T_ADRESSE2, T_ADRESSE3, T_CODEPOSTAL, T_VILLE, T_SECTEUR, ';
    szGroupBy := szGroupBy + 'T_PAYS, T_TELEPHONE, T_FAX, T_PAYEUR, T_RVA, T_DEVISE, T_CREDITDEMANDE, T_DATEMODIF, G_NATUREGENE, J_NATUREJAL, R_PAYS, R_DOMICILIATION, R_ETABBQ, R_GUICHET, R_NUMEROCOMPTE, R_CLERIB';
    SetControlText('XX_GROUPBY', szGroupBy);

    // Nouveau debiteur !
    if (cboModeleEtat.Value = 'CG1') then
    begin
      {JP 21/09/05 : FQ 16626 : Affectation de la DBListe du Mul}
      SetLaListe('CPEXPORTDEB');

      szWhere := 'T_AUXILIAIRE IS NOT NULL';
      szGroupBy := 'T_AUXILIAIRE, T_LIBELLE, T_NATUREAUXI, T_CODEPOSTAL, T_VILLE, T_DATEMODIF, T_CREDITDEMANDE, T_SIRET, T_ENSEIGNE, T_ADRESSE1, T_ADRESSE2, T_ADRESSE3, ';
      szGroupBy := szGroupBy + 'T_PAYS, T_TELEPHONE, T_DEVISE, T_NIF, T_FAX, T_RVA, T_PAYEUR, R_PAYS, R_DOMICILIATION, R_ETABBQ, R_GUICHET, R_NUMEROCOMPTE, R_CLERIB';
      SetControlText('XX_GROUPBY', szGroupBy);
      SetControlText('XX_WHERE', szWhere);
    end

    // Balance âgée
    else if (cboModeleEtat.Value = 'CG2') then
    begin
      {JP 21/09/05 : FQ 16626 : Affectation de la DBListe du Mul}
      SetLaListe('CPEXPORTBALAGEE');

      // Rempli les comptes auxiliaires et généraux si vides
      if (GetControlText('T_AUXILIAIRE') = '') then SetControlText('T_AUXILIAIRE', GetMinMaxCompte('TIERS', 'MIN(T_AUXILIAIRE)', '0'));
      if (GetControlText('T_AUXILIAIRE_') = '') then SetControlText('T_AUXILIAIRE_', GetMinMaxCompte('TIERS', 'MAX(T_AUXILIAIRE)', 'ZZZZZZZZZZZZZZZZZ'));

      if (GetControlText('E_GENERAL') = '') then SetControlText('E_GENERAL', GetMinMaxCompte('GENERAUX', 'MIN(G_GENERAL)', '0'));
      if (GetControlText('E_GENERAL_') = '') then SetControlText('E_GENERAL_', GetMinMaxCompte('GENERAUX', 'MAX(G_GENERAL)', '0'));

      // Rempli le where et le select
      szExoV8 := '01/01/1900';
      if VH.ExoV8.Code <> '' then szExoV8 := USDATETIME(VH.ExoV8.Deb)
      else szExoV8 := '01/01/1900';

      szAuxi := GetControlText('T_AUXILIAIRE');
      szAuxi_ := GetControlText('T_AUXILIAIRE_');
      szGen := GetControlText('E_GENERAL');
      szGen_ := GetControlText('E_GENERAL_');

      dtDate := StrToDate(DATE1.Text);

      szDate := USDATE(DATE1);

      szCommun1 := '(SELECT (SUM(E_DEBIT)- SUM(E_CREDIT)) FROM ECRITURE WHERE E_AUXILIAIRE=CPEXPORT.T_AUXILIAIRE AND E_GENERAL>="' + szGen + '" AND E_GENERAL<="' + szGen_ + '" AND ';
      szCommun2 := '((E_ETATLETTRAGE<>"TL" AND E_ETATLETTRAGE<>"RI") OR (E_ETATLETTRAGE="TL" AND E_DATEPAQUETMAX>"' + szDate + '")) AND E_ECRANOUVEAU<>"CLO" AND E_ECRANOUVEAU<>"OAN" AND E_QUALIFPIECE="N" AND ' +
          'E_DATECOMPTABLE<="' + szDate + '" AND E_DATECOMPTABLE>="' + szExoV8 + '"';

      szWhere := szCommun1 + szCommun2 + ')<>0 AND T_AUXILIAIRE>="' + szAuxi + '" AND T_AUXILIAIRE<="' + szAuxi_ + '"';
      SetControlText('XX_WHERE', szWhere);

      szSelect := 'DISTINCT(T_AUXILIAIRE), T_LIBELLE, T_NATUREAUXI, E_DEVISE AS T_DEVISE, "' + Date1.Text + '" AS C4, ';
      // Solde du compte à la date de situation
      szSelect := szSelect + szCommun1 + 'E_ETATLETTRAGE<>"TL" AND E_ECRANOUVEAU<>"CLO" AND E_ECRANOUVEAU<>"OAN" AND E_QUALIFPIECE="N" AND E_DATECOMPTABLE<="' + szDate + '" AND E_DATECOMPTABLE>="' + szExoV8 + '") C5, ';
      // Nombre de pièces de nature Facture de type N sur le mois de la date de situation
      szSelect := szSelect + '(SELECT (SUM(E_DEBIT)- SUM(E_CREDIT)) FROM ECRITURE WHERE E_AUXILIAIRE=CPEXPORT.T_AUXILIAIRE AND E_GENERAL>="' + szGen + '" AND E_GENERAL<="' + szGen_ + '" AND E_QUALIFPIECE="N" AND E_PERIODE="' + IntToStr(GetPeriode(dtDate)) + '" AND E_NATUREPIECE="FC" AND E_TYPEMVT="TTC") C6, ';
      // Non échu >= 30 jours          //COUNT(E_NATUREPIECE)
      szSelect := szSelect + szCommun1 + szCommun2 + ' AND E_DATEECHEANCE>="' + USDateTime(dtDate + 30) + '") C7, ';
      // Non échu < 30 jours
      szSelect := szSelect + szCommun1 + szCommun2 + ' AND E_DATEECHEANCE> "' + USDateTime(dtDate) + '" AND E_DATEECHEANCE<"' + USDateTime(dtDate + 30) + '") C8, ';
      // Echu de 1 à 30 jours
      szSelect := szSelect + szCommun1 + szCommun2 + ' AND E_DATEECHEANCE>="' + USDateTime(dtDate - 30) + '" AND E_DATEECHEANCE<="' + USDateTime(dtDate) + '") C9, ';
      // Echu de 31 à 60 jours
      szSelect := szSelect + szCommun1 + szCommun2 + ' AND E_DATEECHEANCE>="' + USDateTime(dtDate - 60) + '" AND E_DATEECHEANCE<"' + USDateTime(dtDate - 30) + '") C10, ';
      // Echu de 61 à 90 jours
      szSelect := szSelect + szCommun1 + szCommun2 + ' AND E_DATEECHEANCE>="' + USDateTime(dtDate - 90) + '" AND E_DATEECHEANCE<"' + USDateTime(dtDate - 60) + '") C11, ';
      // Echu > 90 jours
      szSelect := szSelect + szCommun1 + szCommun2 + ' AND E_DATEECHEANCE< "' + USDateTime(dtDate - 90) + '") C12 ';
      SetControlText('XX_SELECT', szSelect);
    end

    // Ecritures non soldées
    else if (cboModeleEtat.Value = 'CG4') then
    begin
      {JP 21/09/05 : FQ 16626 : Affectation de la DBListe du Mul}
      SetLaListe('CPEXPORTCGA');
      szWhere := 'AND ((E_ETATLETTRAGE<>"TL") OR (E_ETATLETTRAGE="TL" AND E_DATEPAQUETMAX>"' + UsDateTime(StrToDate(DATE1.text)) + '"))';
      SetControlText('XX_WHERE', szWhere);
    end
    else
      {JP 21/09/05 : FQ 16626 : Affectation de la DBListe du Mul}
      SetLaListe('CPEXPORTCGA');
  end

  {Export GEFACTOFRANCE}
  else if (TypeExport = cex_GEF ) then
  begin
    {JP 20/09/05 : FQ 16627 : Gestion de la combo type d'export GEFACTOFRANCE}
    MajZonesGeFactoFrance;
    // Demande de garantie
    if (cboModeleEtat.Value = 'GE4') then
    begin
      {JP 21/09/05 : FQ 16626 : Affectation de la DBListe du Mul}
      SetLaListe('CPEXPORTDGA');

      szWhere := 'T_AUXILIAIRE IS NOT NULL';

      szGroupBy := 'T_AUXILIAIRE, T_LIBELLE, T_NATUREAUXI, T_CODEPOSTAL, T_VILLE, T_CREDITDEMANDE, T_SIRET, T_ENSEIGNE, T_ADRESSE1, T_ADRESSE2, T_PAYS, T_TELEPHONE, T_DEVISE, R_PAYS, R_DOMICILIATION, R_ETABBQ, R_GUICHET, R_NUMEROCOMPTE, R_CLERIB';
      SetControlText('XX_GROUPBY', szGroupBy);
      SetControlText('XX_WHERE', szWhere);
    end
    else
    begin
      {JP 21/09/05 : FQ 16626 : Affectation de la DBListe du Mul
       YMO 03/04/2006 Affectation de la liste VCOM pour les champs obligatoires dont E_NUMCFON
       sans changer la liste CPEXPORT pour la SOCREF740
      SetLaListe('CPEXPORTVCOM');

      JP 23/05/06 : FQ 17890 : Non, il faut laisser CPMULEXPORT}
      SetLaListe('CPEXPORT');

      {JP 20/06/05 : Les auxiliaires sont filtrés en fonction du type de contrats}
      szWhere := 'T_AUXILIAIRE IS NOT NULL AND ' + THEdit(GetControl('T_AUXILIAIRE')).Plus;

      {JP 22/09/05 : FQ 16623 : si on ne demande aucune impression ou les deux, on traite
                     les avoirs et les factures clients. Sinon pour le bordereau d'avoirs
                     on filtre sur la nature AC et pour les quittances sur FC}
      if (CBOMODELEGFF.Value = '') or (CBOMODELEGFF.Value = 'BAQ') then
        szWhere := 'E_NATUREPIECE IN ("FC", "AC") AND ' + szWhere
      else if (Copy(CBOMODELEGFF.Value, 1, 2) = 'BO') then
        szWhere := 'E_NATUREPIECE = "AC" AND ' + szWhere
      else if (Copy(CBOMODELEGFF.Value, 1, 2) = 'QI') then
        szWhere := 'E_NATUREPIECE = "FC" AND ' + szWhere;

      SetControlText('XX_WHERE', szWhere);
      szOrderBy := 'E_AUXILIAIRE, E_REFINTERNE';
      SetControlText('XX_ORDERBY', szOrderBy);
    end;
  end

  {JP 10/05/05 : Export VCOM : On limite la sélection aux créditeurs divers, fournisseurs, et
                 Divers à la demande de OG même si cela va renvoyer des écritures clients}
  else if TypeExport = cex_VCOM then begin
    {JP 21/09/05 : FQ 16626 : Affectation de la DBListe du Mul}
    SetLaListe('CPEXPORTVCOM');
    SetControlText('XX_WHERE', 'T_NATUREAUXI IN ("AUC", "FOU", "DIV")');
  end
  else
    {JP 21/09/05 : FQ 16626 : Affectation de la DBListe du Mul}
    SetLaListe('CPEXPORT');
end;

//////////////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPMULEXPORT.OnArgument(S : string);
begin
  inherited;

  //SDA le 17/12/2007 version belge
  {$IFDEF CISXPGI}
  if VH^.PaysLocalisation = CodeIsoDuPays('BEL') then
  begin
    //Gestion liste du MUL
    SetLaListe('CPLISTECISXCFONB');
    //Gestion éléments écran
    //Zone de saisie et étiquette associée visibles
    THLabel(GetControl('LBBANQUE')).Visible := True;
    THValComboBox(GetControl('CBBANQUE')).Visible := True;
    //Modification caption modèle
    THLabel(GetControl('LBLMODELEEXPORT')).Caption := TraduireMemoire ('Type export');
    //Modification datatype modèle
    ChargeTypesExport('BEL');
    //Modification datatype type
    THValComboBox(GetControl('CBTYPE')).DataType := 'CPEXPORTFICTEXTE';
    THValComboBox(GetControl('CBTYPE')).ItemIndex := 0;
    // Bouton paramétrage export rendu invisible
    TToolBarButton97(GetControl('BPARAMEXPORT')).Visible := False;
    //InitDB export belge
    CopieTableAccess(false);
    //Initialisation des composants
    InitComposants;
    //Initialisation des évènements
    InitEvenements;
    //Titre de l'écran
    CaptionOrigine := Ecran.Caption;
    InitGestionSolde(False);
    exit;
  end;
  {$ENDIF}
  //Fin SDA le 17/12/2007

  ClientOk := ReadTokenSt(S) = 'CLI';
  {JP 21/09/05 : FQ 16626 : Affectation de la DBListe du Mul}
  if ClientOk then SetLaListe('CPEXPORTCGA')
              else SetLaListe('CPEXPORTVCOM');

  {Initialisation des composants}
  InitComposants;
  {Initialisation des évènements}
  InitEvenements;
  {Initialisation des variables et des contrôles}
  InitVariables;

  LibellesTableLibre(TTabSheet(GetControl('PZLIBRE', True)), 'TT_TABLE', 'T_TABLE', 'T');
  {FQ 15333 : Mise à jour de l'écran en fonction du type d'export}
  CbTypeOnChange(CbType);

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    THEdit(GetControl('T_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;
    THEdit(GetControl('T_AUXILIAIRE_', true)).OnElipsisClick:=AuxiElipsisClick;
  end;
end;

//////////////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPMULEXPORT.OnClose;
begin
    inherited;
end;

//////////////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPMULEXPORT.OnDisplay();
begin
    inherited;
end;

//////////////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPMULEXPORT.OnCancel();
begin
    inherited;
end;

//////////////////////////////////////////////////////////////////////////////////////////

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 29/06/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOF_CPMULEXPORT.cboModeleEtatOnChange(Sender : TObject);
begin
    // Export CGA
    if (TypeExport = cex_CGA) then
    begin
        // Nouveau Debiteur
        if (cboModeleEtat.Value = 'CG1') or (cboModeleEtat.Value = 'CG2') then
        begin
            SetControlText('E_EXERCICE', '');
            SetControlText('E_DATECOMPTABLE', stDate1900);
            SetControlText('E_DATECOMPTABLE_', stDate2099);
        end
        else begin
            SetControlText('E_EXERCICE', VH^.Entree.Code);
            SetControlText('E_DATECOMPTABLE', DateToStr(V_PGI.DateEntree));
            SetControlText('E_DATECOMPTABLE_', DateToStr(V_PGI.DateEntree));
        end;
    end

    // Export GE FactoFrance
    else if (TypeExport = cex_GEF) then
    begin
      {Demande de garantie}
      if (cboModeleEtat.Value = 'GE4') then begin
        SetControlText('E_EXERCICE', '');
        SetControlText('E_DATECOMPTABLE', stDate1900);
        SetControlText('E_DATECOMPTABLE_', stDate2099);
        {22/09/05 : FQ 16623/16625 : Pas d'édition}
        CBOModeleGFF.ItemIndex := -1;
        CBOModeleGFF.Enabled := False;
      end
      {Quittance subrogative}
      else begin
        SetControlText('E_EXERCICE', VH^.Entree.Code);
        SetControlText('E_DATECOMPTABLE', DateToStr(V_PGI.DateEntree));
        SetControlText('E_DATECOMPTABLE_', DateToStr(V_PGI.DateEntree));
        {22/09/05 : FQ 16623/16625}
        CBOModeleGFF.ItemIndex := 0;
        CBOModeleGFF.Enabled := True;
      end;

      {JP 20/09/05 : FQ 16627 : Gestion de la combo type d'export GEFACTOFRANCE}
      MajZonesGeFactoFrance;
    end

    {FQ 15333 Export VCOM et autres}
    else if (TypeExport = cex_VCOM) then begin
      SetControlText('E_EXERCICE', VH^.Entree.Code);
      SetControlText('E_DATECOMPTABLE', DateToStr(V_PGI.DateEntree));
      SetControlText('E_DATECOMPTABLE_', DateToStr(V_PGI.DateEntree));
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 29/06/2004
Modifié le ... :   /  /
Description .. : Fiche n°13690 => rendre les etat parametrable
Mots clefs ... :
*****************************************************************}

procedure TOF_CPMULEXPORT.BPARAMBORDERAUClick(Sender : TObject);
var
    lStOldEtat : string;
begin
    // Pour forcer la maj de la sélection de l'état,
    //  ( mal mis à jour si clique direct sur bouton de paramétrage ... )
    NextPrevControl(Ecran);
    {JP 22/09/05 : FQ 16623 - 16625 : sur tous et aucun, on sort
     JP 14/05/07 : < 1 car on n'a plus Tous !!!}
    if CBOMODELEGFF.ItemIndex < 1 then Exit;

    if (not ExJaiLeDroitConcept(ccParamEtat, true)) then exit;

    // Edition Etat / Document
    lStOldEtat := CBOMODELEGFF.Value;

    EditEtat('E', 'GFF', lStOldEtat, true, nil, '', '');

    // Rechargement de la Combo
    CBOMODELEGFF.ReLoad;
    if (lStOldEtat <> '') then CBOMODELEGFF.Value := lStOldEtat;
end;

//////////////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPMULEXPORT.BPARAMEXPORTClick(Sender : TObject);
var
    lStOldEtat : string;
begin
    // Pour forcer la maj de la sélection de l'état,
    //  ( mal mis à jour si clique direct sur bouton de paramétrage ... )
    NextPrevControl(Ecran);

    if not ExJaiLeDroitConcept(ccParamEtat, True) then Exit;

    // Edition Etat / Document
    lStOldEtat := cboModeleEtat.Value;

    Param_LPTexte(nil, 'F', 'FIC', lStOldEtat, True);
    LibereSauvGardeLP;

    // Rechargement de la Combo
    cboModeleEtat.ReLoad;
    if lStOldEtat <> '' then cboModeleEtat.Value := lStOldEtat;
end;

//////////////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPMULEXPORT.E_EXERCICEChange(Sender : TObject);
begin
    ExoToDates(E_EXERCICE.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_);
    if (E_EXERCICE.Value = '') then
    begin
        E_DATECOMPTABLE.Text := stDate1900;
        E_DATECOMPTABLE_.Text := stDate2099;
    end;
end;

//////////////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPMULEXPORT.FListeDblClick(Sender : TObject);
var
{$IFDEF EAGLCLIENT}
    TOBListe, TOBLigne : TOB;
{$ENDIF}
    sMode : string;
    lMulQ : TQuery; // SDa le 17/12/2007
begin

    //SDA le 17/12/2007
    if VH^.PaysLocalisation = CodeIsoDuPays('BEL') then
    begin
      {$IFDEF EAGLCLIENT}
      lMulQ := TFMul(Ecran).Q.TQ;
      lMulQ.Seek(FListe.Row - 1);
      {$ELSE}
      lMulQ := TFMul(Ecran).Q;
      {$ENDIF}
  
      if ((GetField('J_MODESAISIE') <> '-') and (GetField('J_MODESAISIE') <> '')) then
        LanceSaisieFolio(lMulQ, taConsult)
      else
        TrouveEtLanceSaisieParam(lMulQ, taConsult, GetField('E_QUALIFPIECE'), False, '');
    end;
    //Fin SDA le 17/12/2007

    // Export CGA : Balance âgée ET Nouveaux débiteurs OU Export GEFF : Demande de garantie => zoom sur AUXILIAIRE
    {FQ 16626 : à partir de l'Agl 5.8.1.147}
  //  if ((TFMul(Ecran).Q.Liste = 'CPEXPORTBALAGEE') or (TFMul(Ecran).Q.Liste = 'CPEXPORTDGA') or (TFMul(Ecran).Q.Liste = 'CPEXPORTDEB')) then
    if ((TFMul(Ecran).DBListe = 'CPEXPORTBALAGEE') or (TFMul(Ecran).DBListe = 'CPEXPORTDGA') or (TFMul(Ecran).DBListe = 'CPEXPORTDEB')) then
    begin
        FicheTiers(nil, '', GetField('T_AUXILIAIRE'), taConsult, 0);
    end
        // Autre cas => Zoom sur pieces            {FQ 16005 : JP 16/06/05 : Zoom VCOM}
    else if (TFMul(Ecran).DBListe = 'CPEXPORT') or (TypeExport = cex_VCOM) or (TFMul(Ecran).DBListe = 'CPEXPORTCGA') then
    begin
{$IFDEF EAGLCLIENT}
        if (Q.TQ.Detail.Count = 0) then Exit;

        sMode := GetField('E_MODESAISIE');
        if ((sMode <> '') and (sMode <> '-')) then PGIInfo('Saisie bordereau non implémentée en mode CWAS', Ecran.Caption)
        else
        begin
            TOBListe := TOB.Create('Liste Ecriture', nil, -1);
            TOBLigne := Q.TQ.Detail[FListe.row - 1];
            TOB.Create('ECRITURE', TOBListe, -1);
            TOBListe.Detail[0].Dupliquer(TOBLigne, False, True);
            TrouveEtLanceSaisie(TOBListe, taConsult, 'N');
        end;
{$ELSE}
        if ((Q.EOF) and (Q.BOF)) then Exit;

        sMode := GetField('E_MODESAISIE');
        if ((sMode <> '') and (sMode <> '-')) then LanceSaisieFolio(Q, taConsult)
        else TrouveEtLanceSaisie(Q, taConsult, 'N');
{$ENDIF}
    end;
end;

//////////////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPMULEXPORT.BOuvrirClick(Sender : TObject);
var
    T : Tob;
    i : Integer;
    SommeA : Double; {JP 23/09/05 : FQ 16623 - 16625}
    SommeF : Double; {JP 23/09/05 : FQ 16623 - 16625}

    // Ajoute une fille à la tob T avec les champs présents dans la liste

    function AddFille : TOB;
    var
        TF : Tob;
        {$IFNDEF EAGLCLIENT}
        i : Integer;
        {$ENDIF}
    begin
        TF := Tob.Create('ECRITURE_', T, -1);
        {$IFDEF EAGLCLIENT}
        // Renseigne les valeurs
        TF.Dupliquer(Q.TQ.CurrentFille,true,true);
        {$ELSE}
        // Renseigne les valeurs
        for i := 0 to Q.FieldCount - 1 do TF.AddChampSupValeur(Q.Fields.Fields[i].FieldName, Q.Fields.Fields[i].AsString, False);
        {$ENDIF}
        // Export CGA
        if (TypeExport = cex_CGA) then TF.AddChampSupValeur('DATE1', Date1.Text, False);

        Result := TF;
    end;
var
  OldDevise : string;
  Quitt     : Integer;
  LigneA    : Integer; {JP 14/12/07 : FQ 22047}
  LigneF    : Integer; {JP 14/12/07 : FQ 22047}
  n, p      : Integer;
  strWhere, strSql  : hstring; //SDA le 21/01/2008
  MyTobVue  : Tob; //SDA le 21/01/2008
  QQ : TQuery; //SDA le 21/01/2008
begin
    inherited;
    //SDA le 17/12/2007
    //Contrôle zones écran
    {$IFDEF CISXPGI}
    if VH^.PaysLocalisation = CodeIsoDuPays('BEL') then
    begin
      if GetControlText('MODELEEXPORT') = '' then begin
        PGIError('Merci de sélectionner un type d''export.', Ecran.Caption);
        Exit;
      end;
      if GetControlText('CBBANQUE') = '' then begin
        PGIError('Merci de sélectionner une banque.', Ecran.Caption);
        Exit;
      end;
    end;
    {$ENDIF}
    //Fin SDA le 17/12/2007
    if (cboModeleEtat.ItemIndex < 0) then
    begin
        PGIError('Vous devez renseigner un modèle d''export.', Ecran.Caption);
        Exit;
    end;

    if (not FListe.AllSelected) and (FListe.NbSelected = 0) then
    begin
        PGIError('Vous devez sélectionner au moins une ligne.', Ecran.Caption);
        Exit;
    end;

    //SDA le 17/12/2007 version belge
    //La banque doit être sélectionnée dans le cadre d'un traitement belge
    {$IFDEF CISXPGI}
    if VH^.PaysLocalisation = CodeIsoDuPays('BEL') then
    begin
      //Création objet export
      if Assigned(ObjEXPORT) then FreeAndNil(ObjEXPORT);
        CreateObjEXPORT;

      //SDA le 21/01/2008 remaniement des virements belges
      strWhere := TFMul(Ecran).Q.Criteres; //Récupération de la clause where
      MyTobVue := TOB.Create('CPEXPORTCFONBCISX', nil, -1); // Création TOB
      try
        try
          //attention, ne pas modifier l'ordre des champs de la requête,
          //si nouveaux champs les positionner en fin de liste des champs
          strSql:='SELECT BQ_NUMEMETVIR, BQ_ETABBQ, BQ_GUICHET, '+
                  'BQ_NUMEROCOMPTE, BQ_CLERIB, BQ_CODEIBAN, ' +
                  'BQ_DEVISE, E_LIBELLE, E_REFINTERNE, ' +
                  'E_RIB, E_DEBIT, E_DEVISE, E_AUXILIAIRE, ' +
                  'E_GENERAL, E_NUMEROPIECE, G_LIBELLE, ' +
                  'G_PAYS, R_DOMICILIATION, R_CODEBIC, ' +
                  'R_CODEIBAN, T_LIBELLE, T_ADRESSE1, ' +
                  'T_ADRESSE2, T_CODEPOSTAL, T_VILLE, ' +
                  'T_PAYS, E_NUMLIGNE, E_DATECOMPTABLE, ' +
                  'E_JOURNAL, E_EXERCICE, E_QUALIFPIECE, ' +
                  'E_NUMECHE, G_NATUREGENE, E_MODEPAIE ' +
                  'FROM CPEXPORTCFONBCISX ';
          strSql := strSql + ' WHERE ' + strWhere;
          QQ:=OpenSql(strSql,True);
        if not QQ.Eof then
          //Alimentation de la tob
          MyTobVue.LoadDetailDB('CPEXPORTCFONBCISX', '', '', QQ, False);
        Except
        end;
      finally
        Ferme(QQ);
      end;

      n := 0; // SDA le 21/01/2008
      TFMul(Ecran).Q.First;
      if FListe.AllSelected then
        while not TFMul(Ecran).Q.EOF do begin
          {Récupération du paramétrage de la grille et de valeurs de la ligne courante}
          ChargeMaTobTraitement(TFMul(Ecran).FNomFiltre, MyTobVue.Detail[n]); //SDA le 21/01/2008
          n := n + 1; // SDA le 21/01/2008
          TFMul(Ecran).Q.Next;
       end
      else
        //SDA le 21/01/2008
        if MyTobVue.Detail.Count > 0 then
        begin
          for n := 0 to FListe.nbSelected - 1 do begin
            FListe.GotoLeBookmark(n);
            {$IFDEF EAGLCLIENT}
            TFMul(Ecran).Q.TQ.Seek(FListe.Row - 1);
            {$ENDIF}
            {Récupération des éléments de la ligne sélectionnée}
            ChargeMaTobTraitement(TFMul(Ecran).FNomFiltre, MyTobVue.Detail[n]);
        end;
        //Fin SDA le 21/01/2008
      end;

      //Fin SDA le 21/01/2008

      //Récup. ligne(s) sélectionnée(s)
      (* SDA le 21/01/2008 TFMul(Ecran).Q.First;
      if FListe.AllSelected then
        while not TFMul(Ecran).Q.EOF do begin
          {Récupération du paramétrage de la grille et de valeurs de la ligne courrante}
          NewTobFromDBL(TFMul(Ecran).FNomFiltre);
          TFMul(Ecran).Q.Next;
       end
      else
      for n := 0 to FListe.nbSelected - 1 do begin
        FListe.GotoLeBookmark(n);
        {$IFDEF EAGLCLIENT}
        TFMul(Ecran).Q.TQ.Seek(FListe.Row - 1);
        {$ENDIF}
        {Récupération du paramétrage de la grille et de valeurs de la ligne courrante}
        NewTobFromDBL(TFMul(Ecran).FNomFiltre);
      end; *)
      //Traitement export
      ObjEXPORT.Pays := 'BEL';
      ObjEXPORT.Script := GetControlText('MODELEEXPORT');
      ObjEXPORT.Banque := GetControlText('CBBANQUE');
      ObjEXPORT.LanceTraitement;
      //Libération de l'objet
      FreeAndNil(MyTobVue); // SDA le 21/01/2008
      exit;
    end;
    {$ENDIF}
    //Fin SDA le 17/12/2007

    // Export CGA
    if (TypeExport = cex_CGA) then
        DateOnExit(nil);

    try
        T := Tob.Create('ECRITURE_', nil, -1);

        if not FListe.AllSelected then
        begin
            for i := 0 to FListe.NbSelected - 1 do
            begin
                FListe.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
                Q.TQ.Seek(FListe.Row-1);
{$ENDIF}
                AddFille;
            end;
        end
        else
        begin
            Q.First;
            while not Q.EOF do
            begin
                AddFille;
                Q.Next;
            end;
        end;

        // GE Factofrance (Pour l'édition des quittances)
        if (TypeExport = cex_GEF) then
        begin
          //Sg6 18/01/05 : FQ 13696 rajout du tri sur E_DEVISE pour rupture sur E_DEVISE
          //SG6 17.03.05 FQ 15498
          if (not (cboModeleEtat.Value = 'GE4')) then T.Detail.Sort('E_DEVISE;E_AUXILIAIRE;E_REFINTERNE');
          SommeA := 0.0;{JP 23/09/05 : FQ 16623 - 16625}
          SommeF := 0.0;{JP 23/09/05 : FQ 16623 - 16625}
          LigneA := 0; {JP 14/12/07 : FQ 22047}
          LigneF := 0; {JP 14/12/07 : FQ 22047}
          OldDevise := '';
          n := 0;
          Quitt := ValeurI(GetParamSocSecur('SO_NOQUITTANCE',0));
          for i := 0 to T.Detail.Count - 1 do begin
            {Changement de devises}
            if (OldDevise <> T.Detail[i].GetString('E_DEVISE')) then begin
              {JP 16/05/07 : FQ 19142 : si on est en Monodevise, il faut quand même mettre à jour
                             le numéro de quittance !!! on sort ce morceau de IF ci-dessous}
              Inc(Quitt);
              T.Detail[i].AddChampSupValeur('QUITT', Quitt);
              {On est pas sur le premier enregistrement}
              if OldDevise <> '' then begin
                for p := n to i - 1 do begin
                  {Demande de Garantie}
                  if cboModeleEtat.Value = 'GE4' then begin
                    T.Detail[p].AddChampSupValeur('SOMME', SommeA);
                    T.Detail[p].AddChampSupValeur('LIGNE', LigneA);
                  end
                  {Gestion des avoirs et des factures}
                  else begin
                    if UpperCase(T.Detail[p].GetString('E_NATUREPIECE'))= 'AC' then begin
                      T.Detail[p].AddChampSupValeur('SOMME', SommeA);
                      {JP 14/12/07 : FQ 22047 : Distinction du nombre d'avoirs et de factures}
                      T.Detail[p].AddChampSupValeur('LIGNE', LigneA);
                    end
                    else begin
                      T.Detail[p].AddChampSupValeur('SOMME', SommeF);
                      {JP 14/12/07 : FQ 22047 : Distinction du nombre d'avoirs et de factures}
                      T.Detail[p].AddChampSupValeur('LIGNE', LigneF);
                    end;
                  end;

                  T.Detail[p].AddChampSupValeur('QUITT', Quitt);
                end;
              end;

              SommeA := 0.0;{JP 23/09/05 : FQ 16623 - 16625}
              SommeF := 0.0;{JP 23/09/05 : FQ 16623 - 16625}
              LigneA := 0; {JP 14/12/07 : FQ 22047}
              LigneF := 0; {JP 14/12/07 : FQ 22047}
              OldDevise := T.Detail[i].GetString('E_DEVISE');
              n := i;
              {Inc(Quitt); JP 16/05/07 : FQ 19142}
            end;

            {JP 20/06/05 : FQ 15500 : on gère tout en devises, comme cela cela marche dans tous les cas}
//            Somme := Somme + T.Detail[i].GetDouble('E_DEBITDEV') - T.Detail[i].GetDouble('E_CREDITDEV');
            if cboModeleEtat.Value = 'GE4' then begin
              SommeA := SommeA + T.Detail[i].GetDouble('E_CREDITDEV') - T.Detail[i].GetDouble('E_DEBITDEV');
              Inc(LigneA);
            end
            {Gestion des avoirs et des factures}
            else begin
              if UpperCase(T.Detail[i].GetString('E_NATUREPIECE'))= 'AC' then begin
                SommeA := SommeA + T.Detail[i].GetDouble('E_CREDITDEV') - T.Detail[i].GetDouble('E_DEBITDEV');
                {JP 14/12/07 : FQ 22047 : Distinction du nombre d'avoirs et de factures}
                Inc(LigneA);
              end
              else begin
                SommeF := SommeF + T.Detail[i].GetDouble('E_DEBITDEV') - T.Detail[i].GetDouble('E_CREDITDEV');
                {JP 14/12/07 : FQ 22047 : Distinction du nombre d'avoirs et de factures}
                Inc(LigneF);
              end;
            end;

          end;

          for p := n to T.Detail.Count - 1 do begin
            {Demande de Garantie}
            if cboModeleEtat.Value = 'GE4' then begin
              T.Detail[p].AddChampSupValeur('SOMME', SommeA);
              T.Detail[p].AddChampSupValeur('LIGNE', LigneA); {JP 14/12/07 : FQ 22047}
            end
            {Gestion des avoirs et des factures}
            else begin
              if UpperCase(T.Detail[p].GetString('E_NATUREPIECE'))= 'AC' then begin
                T.Detail[p].AddChampSupValeur('SOMME', SommeA);
                {JP 14/12/07 : FQ 22047 : Distinction du nombre d'avoirs et de factures}
                T.Detail[p].AddChampSupValeur('LIGNE', LigneA);
              end
              else begin
                T.Detail[p].AddChampSupValeur('SOMME', SommeF);
                {JP 14/12/07 : FQ 22047 : Distinction du nombre d'avoirs et de factures}
                T.Detail[p].AddChampSupValeur('LIGNE', LigneF);
              end;
            end;

            T.Detail[p].AddChampSupValeur('QUITT', Quitt);
          end;

          {JP 21/06/05 : FQ 15498 : le numéro de quittance n'est plus géré dans le générateur de fichier}
          SetParamSoc('SO_NOQUITTANCE', Quitt)
        end;

        DoExport(T);
    finally
        T.Free;
    end;
end;

//////////////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPMULEXPORT.DoExport(T : Tob);
var
  Bor   : string;{JP 22/09/05 : FQ 16625}
  Quit  : string;{JP 22/09/05 : FQ 16625}
  tEtat : TOB;   {JP 22/09/05 : FQ 16625}
begin
  {C'est un export CGA : On demande de choisir un répertoire}
  if (TypeExport = cex_CGA) then
    DoExportCGA(T)

  else begin
    //TobDebug(T);
    {C'est un export GEFACTO : On précise un nom de fichier et on imprime}
    if (TypeExport = cex_GEF) then begin
      {BPY le 28/07/2004 : Fiche de bug n° 13691 => trie de la tob !}
      {SG6 18/01/05 : FQ 13696 rajout du tri sur E_DEVISE pour rupture sur E_DEVISE}
      if (not (cboModeleEtat.Value = 'GE4')) then T.Detail.Sort('E_DEVISE;E_AUXILIAIRE;E_REFINTERNE');

      {JP 04/05/05 : Boite de dialogue et génération du fichier ASCII}
      if not LancerDialog(T) then
          Exit
      else
      if (cboModeleEtat.Value <> 'GE4') then
      // YMO 03/04/2006 Pour GEFACTOFRANCE, on tope les écritures (Tâche Devt3225 / FQ16977)
          MajE_CFONBOK(T);

      {JP 22/09/05 : FQ 16625 : si ItemIndex = 1, on imprime les deux états}
      if (CBOMODELEGFF.ItemIndex = 1) then begin
        RecupEtatParDefaut(Bor, Quit);
        {Récupération des avoirs clients}
        SetTobImpression(T, tEtat, 'AC');
        try
          {Impression du bordereau d'avoirs}
          {JP 22/05/06 : FQ 18138 : on s'assure qu'il y a des données dans la TOB}
          if tEtat.Detail.Count > 0 then
            LanceEtatTob('E', 'GFF', Bor , tEtat, True, False, False, nil, '', Ecran.Caption, False);
          {Récupération des factures Clients}
          SetTobImpression(T, tEtat, 'FC');
          {Impression de la quittance subrogative}
          {JP 22/05/06 : FQ 18138 : on s'assure qu'il y a des données dans la TOB}
          if tEtat.Detail.Count > 0 then
            LanceEtatTob('E', 'GFF', Quit, tEtat, True, False, False, nil, '', Ecran.Caption, False);
        finally
          if Assigned(tEtat) then FreeAndNil(tEtat);
        end;
      end
      {BPY le 28/06/2004 => fiche n°13693 : <<Aucun>> dans le choix des etat
                           (ce qui donne ">0" au lieux de ">=0" ici
       JP 22/09/05 : FQ 16625 : Maintenant il faut que l'ItemIndex soit supérieur à 1}
      else if (CBOMODELEGFF.ItemIndex > 1) then
        LanceEtatTob('E', 'GFF', CBOMODELEGFF.Value, T, True, False, False, nil, '', Ecran.Caption, False);
    end

    {JP 04/05/05 : Export VCOM}
    else if (TypeExport = cex_VCOM) then begin
      {Vérification qu'il n'y a pas d'échéance négative}
      if VerifEcheNegatives(T) and
         {Boite de dialogue et génération du fichier ASCII}
         LancerDialog(T) then
        MajE_CFONBOK(T);
    end

    else begin
        {JP 04/05/05 : Boite de dialogue et génération du fichier ASCII}
        if not LancerDialog(T) then Exit;
    end;
  end;
end;

//////////////////////////////////////////////////////////////////////////////////////////
// Export CGA
//////////////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPMULEXPORT.DoExportCGA(T : Tob);
var
    Err : TMC_ERR;
    szNumOrdre, szDir, szFileName : string;
    TE : THEdit;
    sr : TSearchRec;
    NumOrdre, Len : Integer;

    procedure UpdateNumOrdre;
    var
        i : Integer;
        sz : string;
    begin
        sz := Copy(sr.Name, Len + 1, 3);
        if IsNumeric(sz) then
        begin
            i := StrToInt(sz);
            if i >= NumOrdre then NumOrdre := i + 1;
        end;
    end;
begin
    NumOrdre := 1;
    if (cboModeleEtat.Value = 'CG1') then szFileName := 'NA' else // Nouveaux débiteurs
        if (cboModeleEtat.Value = 'CG2') then szFileName := 'NB' else // Balance âgée
            if (cboModeleEtat.Value = 'CG3') then szFileName := 'NE' else // Remise de facture
                if (cboModeleEtat.Value = 'CG4') then szFileName := 'NC' // Ecritures non soldées
                else szFileName := 'NX';
    szFileName := szFileName + UpperCase(GetParamSocSecur('SO_CGACOMPTE',''));
    Len := Length(szFileName);

    try
        TE := THEdit.Create(Ecran);
        TE.Parent := Ecran;
        TE.Left := -500; // Pour ne pas voir le contrôle s'afficher à l'écran
        TE.DataType := 'DIRECTORY';
        TE.Perform(WM_LBUTTONDBLCLK, 0, 0);
        if (TE.Text <> '') then
        begin
            szDir := IncludeTrailingBackslash(TE.Text) + '*.*';
            // Recherche si un fichier Existe déjà (pour le n° d'ordre)
            if FindFirst(szDir, faAnyFile, sr) = 0 then
            begin
                if (Copy(UpperCase(sr.Name), 1, Len) = szFileName) then UpdateNumOrdre;
                while FindNext(sr) = 0 do
                begin
                    if (Copy(UpperCase(sr.Name), 1, Len) = szFileName) then UpdateNumOrdre;
                end;
            end;
            FindClose(sr);

            szNumOrdre := IntToStr(NumOrdre);
            while Length(szNumOrdre) < 3 do
                szNumOrdre := '0' + szNumOrdre;
            szFileName := IncludeTrailingBackslash(TE.Text) + szFileName + szNumOrdre + '.TXT';

            ExporteLP('F', 'FIC', cboModeleEtat.Value, '', szFileName, False, False, T, Err);
        end;
    finally
        FreeAndNil(TE);
    end;
end;

//////////////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPMULEXPORT.DateOnExit(Sender : TObject);
begin
    Date1.Text := DateToStr(FinDeMois(StrToDate(Date1.Text)));
end;

//////////////////////////////////////////////////////////////////////////////////////////

function TOF_CPMULEXPORT.GetMinMaxCompte(szTable, szCol, szDefaut : string) : string;
var
    Q : TQuery;
    szWhere : string;
begin
    // Condition dépend de la nature auxiliaire
    if szTable = 'GENERAUX' then
        if T_NATUREAUXI.Value <> '' then szWhere := GetWhereColl(T_NATUREAUXI.Value)
        else szWhere := ' WHERE G_COLLECTIF="X"'
    else
        if T_NATUREAUXI.Value <> '' then
            szWhere := ' WHERE T_NATUREAUXI="' + T_NATUREAUXI.Value + '"';

    // Requête
    Q := OpenSQL('SELECT ' + szCol + ' CODE FROM ' + szTable + szWhere, True);
    if not Q.Eof then Result := Q.FindField('CODE').asString
    else Result := szDefaut;
    Ferme(Q);
end;

//////////////////////////////////////////////////////////////////////////////////////////

function TOF_CPMULEXPORT.GetWhereColl(szNatureTiers : string) : string;
begin
    result := ' WHERE G_COLLECTIF="X"';
    if szNatureTiers = 'CLI' then result := result + ' AND G_NATUREGENE="COC"' else
        if szNatureTiers = 'FOU' then result := result + ' AND G_NATUREGENE="COF"' else
            if szNatureTiers = 'SAL' then result := result + ' AND G_NATUREGENE="COS"' else
                if szNatureTiers = 'DIV' then result := result + ' AND G_NATUREGENE="COD"' else
                    if szNatureTiers = 'AUC' then result := result + ' AND (G_NATUREGENE="COC" OR G_NATUREGENE="COD")' else
                        if szNatureTiers = 'AUD' then result := result + ' AND (G_NATUREGENE="COF" OR G_NATUREGENE="COD")';
end;

{Gestion du OnChange sur la combo qui détermine sur quel type d'export on travaille
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEXPORT.CbTypeOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Il n'y a que pour CGA que la liste n'est pas paramétrable : SG6 17.03.05 FQ 15504 }
  SetControlEnabled('BPARAMLISTE', TypeExport <> cex_CGA);
  {FQ 15333 : remplacement du RadioGroup par une combo}
       if TypeExport = cex_CGA  then cboModeleEtat.Plus := 'FIC" AND MO_CODE LIKE "CG%'
  else if TypeExport = cex_GEF  then cboModeleEtat.Plus := 'FIC" AND MO_CODE LIKE "GE%'
  else if TypeExport = cex_VCOM then cboModeleEtat.Plus := 'FIC" AND MO_CODE LIKE "VC%'
  else if TypeExport = cex_DIV  then cboModeleEtat.Plus := 'FIC" AND MO_CODE NOT LIKE "CG%" AND MO_CODE NOT LIKE "GE%" AND MO_CODE NOT LIKE "VC%';

  SetControlVisible ('PNLBALAGEE', TypeExport = cex_CGA);
  SetControlVisible ('PNLGFF'    , TypeExport = cex_GEF);
  SetControlProperty('TSVCOM', 'TABVISIBLE', (TypeExport = cex_VCOM) or (TypeExport = cex_GEF));

  // YMO 03/04/2006 Pour GEFACTOFRANCE, on tope les écritures (Tâche Devt3225 / FQ16977)
  // Ici les informations de l'onglet à l'origine pour VCOM sont cachées
  If TypeExport = cex_GEF then
  begin
    SetControlProperty('TSVCOM', 'CAPTION', 'Exports');
    SetControlProperty('E_CFONBOK', 'CAPTION', 'Pièces exportées');
    SetControlProperty('E_CFONBOK', 'LEFT', '15');
    SetControlProperty('E_CFONBOK', 'TOP', '8');
    SetControlProperty('TE_NUMCFONB', 'LEFT', '15');
    SetControlProperty('TE_NUMCFONB', 'TOP', '34');
    SetControlProperty('E_NUMCFONB', 'LEFT', '75');
    SetControlProperty('E_NUMCFONB', 'TOP', '30');
    SetControlProperty('E_BANQUEPREVI', 'VISIBLE', False);
    SetControlProperty('TE_BANQUEPREVI', 'VISIBLE', False);
    SetControlProperty('CBACTION', 'VISIBLE', False);
    SetControlProperty('LBACTION', 'VISIBLE', False);
    SetControlProperty('CBENVOI', 'VISIBLE', False);
    SetControlProperty('LBENVOI', 'VISIBLE', False);
    SetControlProperty('CBNATECO', 'VISIBLE', False);
    SetControlProperty('LBNATECO', 'VISIBLE', False);
  end;

  {JP 20/06/05 : Maintenant les 3 exports GE FactoFrance sont fusionnés en un seul. J'ai rajouté une ComboBox
                 permettant le choix du type d'export, de manière à filtrer les tiers, en fonction du choix,
                 sur la devise ou sur le pays}
  SetControlVisible('CBGEFACTO'     , TypeExport  = cex_GEF);
  SetControlVisible('E_NATUREPIECE' , TypeExport <> cex_GEF);
  if TypeExport = cex_GEF then begin
    SetControlCaption ('TE_NATUREPIECE', TraduireMemoire('&Type de contrat')); {JP 31/01/08 : FQ 22324}
    THValComboBox(GetControl('CBGEFACTO')).ItemIndex := 0;
    {Pour filtrer les auxiliaires}
    GeFactoChange(GetControl('CBGEFACTO'));
  end
  else
    SetControlCaption('TE_NATUREPIECE' , TraduireMemoire('&Nature pièce')); {JP 31/01/08 : FQ 22324}

  {JP 10/05/05 : FQ 15334 : Les écritures doivent être filtrées sur une banque prévisionnelle}
  if TypeExport = cex_VCOM then
    CbBanque.ItemIndex := 0;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPMULEXPORT.GetTypeExport : string;
{---------------------------------------------------------------------------------------}
begin
  Result := CbType.Value;
end;

{Vérification qu'il n'y a pas d'échénaces négatives
{---------------------------------------------------------------------------------------}
function TOF_CPMULEXPORT.VerifEcheNegatives(var T : TOB) : Boolean;
{---------------------------------------------------------------------------------------}

    {Calcul du total des opérations de l'échéance : ne peut être fait
     dans le générateur de fichiers car l'enregistrement d'échéance est
     en entête et les factures dans le détail : FQ 15541 et 15545
    {----------------------------------------------------------------}
    procedure MajTotauxEcheances(Deb, Fin : Integer; Valeur : string);
    {----------------------------------------------------------------}
    var
      n : Integer;
    begin
      for n := Deb to Fin do
        T.Detail[n].AddChampSupValeur('MNTECHE', Valeur);
    end;

var
  OldEche  : TDateTime;
  OldTiers : string;
  OldInd   : Integer;
  DerInd   : Integer;
  n        : Integer;
  Mnt      : Double;
begin
  Result  := True;
  OldEche := iDate1900;
  Mnt     := 0;
  OldInd  := 0;
  DerInd  := 0;
  {On trie la tob sur les deux champs de rupture du test}
  T.Detail.Sort('E_AUXILIAIRE;E_DATEECHEANCE');
  {Boucle sur les lignes à traiter}
  for n := 0 to T.Detail.Count - 1 do begin
    DerInd := n;
    {Si on change d'échéance ou de tiers ...}
    if (OldEche <> T.Detail[n].GetDateTime('E_DATEECHEANCE')) or
       (OldTiers <> T.Detail[n].GetString('E_AUXILIAIRE')) then begin
      {... on s'assure que le montant de l'échéance est positif}
      if Mnt < 0 then begin
        {On ne génère pas de fichier s'il y a un montant négatif}
        HShowMessage('0;' + Ecran.Caption + ';Il y a des échéances négatives.'#13#13 +
                     'Assurez-vous que vous n''avez pas des avoirs dont la date d''échéance'#13 +
                     'est différente de celle de la facture concernée :'#13 +
                     'Tiers : ' + OldTiers + ' - Échéance : ' + DateToStr(OldEche) + #13#13 +
                     'Le traitement est abandonné.;W;O;O;O;', '', '');
        Result := False;
        Exit;
      end
      else begin
        {On stocke le cumul des opérations sur une échéance : cf FQ 15541 et 15545}
        MajTotauxEcheances(OldInd, n - 1, StrfMontant(Mnt, 16, 2, '', False));

        {Réinitialisation des variables}
        OldEche  := T.Detail[n].GetDateTime('E_DATEECHEANCE');
        OldTiers := T.Detail[n].GetString('E_AUXILIAIRE');
        OldInd   := n;
        {FQ 15991 : JP 16/06/05 : j'avais inversé l'opération, ce qui n'était pas "TOP"
         Mnt := T.Detail[n].GetDouble('E_DEBIT') - T.Detail[n].GetDouble('E_CREDIT');}
        Mnt := T.Detail[n].GetDouble('E_CREDIT') - T.Detail[n].GetDouble('E_DEBIT');
      end;
    end
    else
      {On cumule les montants de l'échéance}
      {FQ 15991 : JP 16/06/05 : j'avais inversé l'opération, ce qui n'était pas "TOP"
       Mnt := Mnt + T.Detail[n].GetDouble('E_DEBIT') - T.Detail[n].GetDouble('E_CREDIT');}
      Mnt := Mnt + T.Detail[n].GetDouble('E_CREDIT') - T.Detail[n].GetDouble('E_DEBIT');
  end;

  {... on s'assure que le montant de l'échéance est positif}
  if Mnt < 0 then begin
    {On ne génère pas de fichier s'il y a un montant négatif}
    HShowMessage('0;' + Ecran.Caption + ';Il y a des échéances négatives.'#13#13 +
                 'Assurez-vous que vous n''avez pas des avoirs dont la date d''échéance'#13 +
                 'est différente de celle de la facture concernée :'#13 +
                 'Tiers : ' + OldTiers + ' - Échéance : ' + DateToStr(OldEche) + #13#13 +
                 'Le traitement est abandonné.;W;O;O;O;', '', '');
    Result := False;
  end
  else
    {On stocke le cumul des opérations pour la dernière échénace}
    MajTotauxEcheances(OldInd, DerInd{- 1 JP 19/07/05 FQ 15991}, StrfMontant(Mnt, 16, 2, '', False));
end;

{Boite de dialogue et génération du fichier ASCII
{---------------------------------------------------------------------------------------}
function TOF_CPMULEXPORT.LancerDialog(var T : TOB) : Boolean;
{---------------------------------------------------------------------------------------}
var
  SD  : TSaveDialog;
  Err : TMC_ERR;
begin
  {Création et paramétrage de la boite de dialogue}
  SD := TSaveDialog.Create(Application);
  try
    SD.DefaultExt := 'TXT';
    SD.Filter := 'Fichier Texte (*.txt)|*.txt';
    SD.FilterIndex := 1;
    SD.Title := 'Export';
    Result := SD.Execute;
    Result := Result and GererRefVCOM(T) and (GererRefGEF(T)<>'@stop');
    {Si on a validé la boite de dialoguer et, si en VCOM, les références du donneur d'ordre ont été saisies...}
    if Result then begin
      {... Lancement de la génération du fichier}
      ExporteLP('F', 'FIC', cboModeleEtat.Value, '', SD.FileName, True, False, T, Err);
      {Gestion des éventuelles erreurs}
      if Err.Code <> 0 then begin
        PgiError(MC_MsgErrDefaut(Err.Code));
        Result := False;
      end;
    end;
  finally
    if Assigned(SD) then FreeAndNil(SD);
  end;
end;

{JP 09/05/05 : FQ 15552 et 15738 : Mise à jour du champ E_CFONBOK après la génération du fichier
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEXPORT.MajE_CFONBOK(T : TOB);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  F : TOB;
begin
  for n := 0 to T.Detail.Count - 1 do begin
    F := T.Detail[n];
    ExecuteSQL('UPDATE ECRITURE SET E_CFONBOK = "X", E_NUMCFONB = "' + RefDonneur + '", ' +
               'E_DATEMODIF = "'     + USDateTime(Now) + '" WHERE ' +
               'E_EXERCICE = "'      + F.GetString('E_EXERCICE') + '" AND ' +
               'E_JOURNAL = "'       + F.GetString('E_JOURNAL')  + '" AND ' +
               'E_DATECOMPTABLE = "' + USDateTime(F.GetDateTime('E_DATECOMPTABLE')) + '" AND ' +
               'E_NUMEROPIECE = '    + F.GetString('E_NUMEROPIECE') + ' AND ' +
               'E_NUMLIGNE = '       + F.GetString('E_NUMLIGNE')    + ' AND ' +
               'E_NUMECHE = '        + F.GetString('E_NUMECHE')     + ' AND ' +
               'E_QUALIFPIECE = "'   + F.GetString('E_QUALIFPIECE') + '"');

  end;
end;

{FQ 15513 : Affichage de la fenêtre de saisie des références du donneur d'ordre
{---------------------------------------------------------------------------------------}
function TOF_CPMULEXPORT.GererRefVCOM(var T : TOB) : Boolean;
{---------------------------------------------------------------------------------------}
var
  n    : Integer;
  Nat  : string;
  Mpa  : string;
  EuOk : Boolean;
begin
  {Ne concerne que les références VCOM}
  Result := TypeExport <> cex_VCOM;
  if not Result then begin
    {FQ 15405 : Gestion de l'action de paiement}
    Result := GetControlText('CBACTION') <> '';
    if not Result then
      HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner une action de paiement.'#13 +
                                          'Le traitement est abandonné.;W;O;O;O;', '', '');

    if Result then begin
      {JP 16/06/05 : VCOM 1000 : Gestion du mode d'envoi}
      Result := GetControlText('CBENVOI') <> '';
      if not Result then
        HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner un mode d''envoi.'#13 +
                                            'Le traitement est abandonné.;W;O;O;O;', '', '');
    end;

    {FQ 15513 Gestion des référence du donneur d'ordre}
    if Result then begin
      RefDonneur := CPRefDonneurOrd(cex_VCOM);
      Result := Trim(RefDonneur) <> '';
      if not Result then
        HShowMessage('0;' + Ecran.Caption + ';Les références du donneur d''ordre sont obligatoires.'#13 +
                                            'Le traitement est abandonné.;W;O;O;O;', '', '')
      {En VCOM 1000, les deux champs qui utilisent RefDonneur sont facultatifs ...}
      else if RefDonneur = cex_1000 then
        {On remet donc RefDonneur à vide}
        RefDonneur := '';
    end;

    if Result then begin
      {Si pour une raison ou une autre on n'était pas passé dans CbBanquePreviChange}
      if (NumCompte = '') and (CodeIban = '') then
        CbBanquePreviChange(cbBanque);

      Nat := Trim(GetControlText('CBNATECO'));
      MPa := Copy(GetControlText('CBENVOI'), 1, 1);
      {On remplit la Tob avec les champs indépendants de la table}
      for n := 0 to T.Detail.Count - 1 do begin
        T.Detail[n].AddChampSupValeur('REFDONNEUR'   , RefDonneur);
        T.Detail[n].AddChampSupValeur('NUMCOMPTE'    , NumCompte);
        T.Detail[n].AddChampSupValeur('BQ_CODEIBAN'  , CodeIban);
        T.Detail[n].AddChampSupValeur('BQ_RIB'       , CodeRib);
        T.Detail[n].AddChampSupValeur('ACTION'       , GetControlText('CBACTION'));
        {16/06/05 : VCOM 1000 : Gestion du mode d'envoi entre la banque et le fournisseur}
        T.Detail[n].AddChampSupValeur('FAXMAIL'      , MPa);

        {16/06/05 : VCOM 1000 : Gestion des virements dans la zone euro :
                    Il faut le Bic dès lors que la banque du fournisseur n'est pas en France
                    Il faut la nature économique, si le montant est supérieur à 12500€}
        EuOk := Copy(T.Detail[n].GetString('R_CODEIBAN'), 1, 2) <> CodeISOFR;
        if EuOk then begin
          if T.Detail[n].GetDouble('MNTECHE') >= 12500 then begin
            T.Detail[n].AddChampSupValeur('BICNATECO', 'N');
            {16/06/05 : VCOM 1000 : Par défaut, on prend la nature économique de la table RIB, sauf si
                        la combo CBNATECO n'est pas vide}
            if Nat <> '' then T.Detail[n].SetString('R_NATECO', Nat);
          end
          else
            T.Detail[n].AddChampSupValeur('BICNATECO', 'B');
        end
        else
          T.Detail[n].AddChampSupValeur('BICNATECO', 'R');

        {16/06/05 : VCOM 1000 : il est recommandé d'utiliser (at) et c'est accepté en VCOM 400}
        if MPa = 'M' then
          T.Detail[n].SetString('T_RVA', FindEtReplace(T.Detail[n].GetString('T_RVA'), '@', '(at)', False))
        {16/06/05 : VCOM 1000 : Si par fax, on supprime les séparateurs et autres caractères non-numériques}
        else if MPa = 'F' then
          T.Detail[n].SetString('T_FAX', CleTelephone(T.Detail[n].GetString('T_FAX'), False));

        {Suppression des caractères non-numériques du N° de téléphone}
        T.Detail[n].SetString('T_TELEPHONE', CleTelephone(T.Detail[n].GetString('T_TELEPHONE'), False));

      end;
    end;
  end;
end;


{---------------------------------------------------------------------------------------}
{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 03/04/2006
Modifié le ... :   /  /
Description .. : Pour GEFACTOFRANCE, on tope les écritures (Tâche
Suite ........ : Devt3225 / FQ16977),
Suite ........ : tout ciomme c'est fait pour l'export VCOM en suivi
Suite ........ : fournisseurs
Suite ........ :
Suite ........ : La référence n'est pas obligatoire
Mots clefs ... :
*****************************************************************}
function TOF_CPMULEXPORT.GererRefGEF(var T : TOB) : String;
{---------------------------------------------------------------------------------------}
var
  n    : Integer;
begin
  {Ne concerne que les références GEFACTOFRANCE}
  If (TypeExport = cex_GEF) And (cboModeleEtat.Value <> 'GE4') then
  begin

    {YMO Gestion des référence (utilisation de l'écran créé pour VCOM
    le terme "donneur d'ordre" ne sert su'en suivi foournisseurs}
    RefDonneur := CPRefDonneurOrd(cex_GEF);

    If Trim(RefDonneur) <> '@stop' then
    Begin
      if Trim(RefDonneur) <> '' then
      begin
        {On remplit la Tob avec la référence passée en paramètre}
        for n := 0 to T.Detail.Count - 1 do
        begin
          T.Detail[n].AddChampSupValeur('REFDONNEUR'   , RefDonneur);
        end;
      end
      else
      begin
      //YMO 04/07/2006 FQ 17819 On ne bloque pas, comme pour VCOM
      HShowMessage('0;' + Ecran.Caption + ';Attention les pièces sélectionnées vont être exportées sans référence ,confirmez-vous ?', '', '');

      end;
    End;
  end;
  //12/06/2006 YMO FQ18356 On différencie la sortie à vide de la sortie par la croix
  Result := Trim(RefDonneur);
end;

{FQ 15334 : Au changement de la banque prévisionnelle pour les exports VCOM
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEXPORT.CbBanquePreviChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  QQ : TQuery;
begin
  {Dans le cas d'un export VCOM ...}
  if (TypeExport = cex_VCOM) and (cbBanque.Value <> '') then begin
    {... On récupère les informations bancaires du compte sur lequel est fait le virement ...}
    QQ := OpenSQL('SELECT BQ_ETABBQ, BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CLERIB, BQ_CODEIBAN FROM BANQUECP ' +
                  'WHERE BQ_GENERAL = "' + cbBanque.Value +'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"', True); // 19/10/2006 YMO Multisociétés
    try
      {... Et on les stocke dans des zones invisibles qui serviront de critères
       dans le générateur de fichiers d'export}
      if not QQ.EOF then begin
        NumCompte := QQ.FindField('BQ_NUMEROCOMPTE').AsString;
        CodeIban  := QQ.FindField('BQ_CODEIBAN'    ).AsString;
        {CodeRib   := QQ.FindField('BQ_ETABBQ'      ).AsString +
                     QQ.FindField('BQ_GUICHET'     ).AsString +
                     QQ.FindField('BQ_NUMEROCOMPTE').AsString +
                     QQ.FindField('BQ_CLERIB'      ).AsString;}
        CodeRib   := Format('%5s%5s%11s%2s', [QQ.FindField('BQ_ETABBQ'      ).AsString,
                                              QQ.FindField('BQ_GUICHET'     ).AsString,
                                              QQ.FindField('BQ_NUMEROCOMPTE').AsString,
                                              QQ.FindField('BQ_CLERIB'      ).AsString]);

      end;
    finally
      Ferme(QQ);
    end;
  end;
end;

{Sélection d'une ligne
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEXPORT.FListeFlipSelection(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if VH^.PaysLocalisation <> CodeIsoDuPays('BEL') then // SDA le 17/12/2007 version belge
    AfficheSoldeSelection
  //SDA le 17/12/2007 version belge
  else
  begin
    Ecran.Caption := CaptionOrigine;
    UpdateCaption(Ecran);
  end;
  //Fin SDA le 17/12/2007
end;

{FQ 15537 : Affichage du solde de la sélection dans le caption de la fiche
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEXPORT.AfficheSoldeSelection;
{---------------------------------------------------------------------------------------}
var
  Solde : Double;
begin
  {Calcul du solde}
  Solde := CalculSoldeSelection;
  {Constitution du titre de la fiche avec le solde}
  if Solde < 0 then
    Ecran.Caption := CaptionOrigine + ' (Cumul de la sélection : ' + StrFMontant(-1 * Solde, 0, V_PGI.OkDecV, '', True) + ' C)'
  else
    Ecran.Caption := CaptionOrigine + ' (Cumul de la sélection : ' + StrFMontant(Solde, 0, V_PGI.OkDecV, '', True) + ' D)';
  {Mise à jour du caption}
  UpdateCaption(Ecran);
  {Si AllSelected est à True, on réinitialise les variables à zéro, car la prochaine manipulation
   videra la sélection}
  if FListe.AllSelected then begin
    TotalCredit := 0;
    TotalDebit  := 0;
  end;
end;

{FQ 15537 : Calcul du solde de la sélection
{---------------------------------------------------------------------------------------}
function TOF_CPMULEXPORT.CalculSoldeSelection : Double;
{---------------------------------------------------------------------------------------}
begin
  if not FListe.AllSelected and (FListe.nbSelected > 0) then begin
    {$IFDEF EAGL}
    if FListe.IsSelected(FListe.Row) then begin
    {$ELSE}
    if FListe.IsCurrentSelected then begin
    {$ENDIF EAGL}
      {JP 09/11/05 : FQ 16872 : Pour les exports CGA (CG3 et CG4, les autres n'ont pas de cumuls}
      if (TFmul(Ecran).FNomFiltre = 'CPEXPORTCGA') then begin
        TotalCredit := TotalCredit + Valeur(VarToStr(GetField('C9')));
        TotalDebit  := TotalDebit  + Valeur(VarToStr(GetField('C8')));
      end
      else begin
        TotalCredit := TotalCredit + Valeur(VarToStr(GetField('E_CREDIT')));
        TotalDebit  := TotalDebit  + Valeur(VarToStr(GetField('E_DEBIT')));
      end;
    end
    else begin
      {JP 09/11/05 : FQ 16872 : Pour les exports CGA (CG3 et CG4, les autres n'ont pas de cumuls}
      if (TFmul(Ecran).FNomFiltre = 'CPEXPORTCGA') then begin
        TotalCredit := TotalCredit - Valeur(VarToStr(GetField('C9')));
        TotalDebit  := TotalDebit  - Valeur(VarToStr(GetField('C8')));
      end
      else begin
        TotalCredit := TotalCredit - Valeur(VarToStr(GetField('E_CREDIT')));
        TotalDebit  := TotalDebit  - Valeur(VarToStr(GetField('E_DEBIT')));
      end;
    end;
  end;
  Result := TotalDebit - TotalCredit;
end;

{Initialisation des composants
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEXPORT.InitComposants;
{---------------------------------------------------------------------------------------}
begin
  {JP 04/0505 : Nouvelle gestion du type d'export}
  CbType := THValComboBox(GetControl('CBTYPE', True));

  {FQ 15334 :Gestion de la banque prévisionnelle pour les exports VCOM}
  cbBanque := THValComboBox(GetControl('E_BANQUEPREVI', True));
  {JP 29/01/07 : gestion du partage de BanqueCP : Uilisation d'une fonction générique}
  SetPlusBanqueCp(cbBanque);

  cboModeleEtat := THValComboBox(GetControl('MODELEEXPORT', True));

  DATE1 := THEdit(GetControl('DATE1', True));

  T_NATUREAUXI := THValComboBox(Getcontrol('T_NATUREAUXI', True));
  CBOMODELEGFF := THValComboBox(Getcontrol('CBOMODELEGFF', True));
  E_EXERCICE   := THValComboBox(Getcontrol('E_EXERCICE'  , True));

  E_DATECOMPTABLE  := THEdit(Getcontrol('E_DATECOMPTABLE' , True));
  E_DATECOMPTABLE_ := THEdit(Getcontrol('E_DATECOMPTABLE_', True));

  E_NUMCFONB  := THEdit(Getcontrol('E_NUMCFONB' , True));
  E_JOURNAL  := THEdit(Getcontrol('E_JOURNAL' , True));

  E_NUMCFONB.OnElipsisClick := OnElipsisClickE_NUMCFONB;

  Q := THQuery(GetControl('Q', True));
  FListe := TFMul(Ecran).FListe;
end;

{Initialisation des évènements
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEXPORT.InitEvenements;
{---------------------------------------------------------------------------------------}
begin
  CbType.OnChange := CbTypeOnChange;
  CbBanque.OnChange := CbBanquePreviChange;

  TToolBarButton97(Getcontrol('BPARAMEXPORT'  , True)).OnClick := BPARAMEXPORTClick;
  TToolBarButton97(Getcontrol('BPARAMBORDERAU', True)).OnClick := BPARAMBORDERAUClick;
  TToolBarButton97(GetControl('BOUVRIR'       , True)).OnClick := BOuvrirClick;

  cboModeleEtat.OnChange := cboModeleEtatOnChange;

  DATE1.OnExit := DateOnExit;
  E_EXERCICE.OnChange := E_EXERCICEChange;

  FListe.OnDblClick := FListeDblClick;
  FListe.OnFlipSelection := FListeFlipSelection;
  {JP 20/06/05 : Ajout d'une combo pour GEFACTOFRANCE}
  THValComboBox(GetControl('CBGEFACTO')).OnChange := GeFactoChange;

  {FQ 15537 : Surcharge de la sélection, pour l'affichage du solde dans le titre de la fenêtre}
  FBSelectAllClick := TFMul(Ecran).bSelectAll.OnClick;
  TFMul(Ecran).bSelectAll.OnClick := BSelectAllClick;
end;

{Initialisation des variables et des contrôles
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEXPORT.InitVariables;
{---------------------------------------------------------------------------------------}
begin
  if ClientOk then CbType.Plus := 'CLI'
              else CbType.Plus := 'FOU';
  CbType.ItemIndex := 0;

  E_EXERCICE.Value := VH^.Entree.Code;
  E_DATECOMPTABLE.Text := DateToStr(V_PGI.DateEntree);
  E_DATECOMPTABLE_.Text := DateToStr(V_PGI.DateEntree);

  DATE1.Text := DateToStr(DebutDeMois(Date()) - 1);
  CBOMODELEGFF.ItemIndex := 0;

  bCanUpdate := False;
  {JP 10/05/05 : FQ 15537 : Affichage du solde dans le caption de la fiche}
  CaptionOrigine := Ecran.Caption;
  InitGestionSolde(False);
end;

{JP 10/05/05 : FQ 15537 : Réinitialisation des variables avec ou sans le Caption
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEXPORT.InitGestionSolde(AvecCaption : Boolean);
{---------------------------------------------------------------------------------------}
begin
  TotalCredit := 0;
  TotalDebit  := 0;
  if AvecCaption then begin
    Ecran.Caption := CaptionOrigine;
    UpdateCaption(Ecran);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEXPORT.BSelectAllClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  PCumul : TPanel;
  CumC   : string;
  CumD   : string;
begin
  {On commence par exécuté le code hérité}
  FBSelectAllClick(Sender);


  PCumul := TPanel( GetControl('PCumul', True));

  {JP 09/11/05 : FQ 16872 : Pour les exports CGA (CG3 et CG4, les autres n'ont pas de cumuls}
  if (TFmul(Ecran).FNomFiltre = 'CPEXPORTCGA') then begin
    if FListe.AllSelected then begin
      TotalCredit := 0;
      TotalDebit  := 0;
      TFMul(Ecran).Q.First;
      while not TFMul(Ecran).Q.EOF do begin
        TotalCredit := TotalCredit + Valeur(VarToStr(GetField('C9')));
        TotalDebit  := TotalDebit  + Valeur(VarToStr(GetField('C8')));
        TFMul(Ecran).Q.Next;
      end;
    end
    else begin
      TotalDebit  := 0;
      TotalCredit := 0;
    end;
  end
  else begin
    CumC := '__QRYPCumul_E_CREDIT';
    CumD := '__QRYPCumul_E_DEBIT';
  end;

  {JP 09/11/05 : FQ 16872 : Pour les exports CGA (CG3 et CG4, les autres n'ont pas de cumuls}
  if not (TFmul(Ecran).FNomFiltre = 'CPEXPORTCGA') then begin
    {Si les cumuls ne sont pas affichés, on sort}
    if (PCumul.FindChildControl(CumC)  = nil) or
       (PCumul.FindChildControl(CumD) = nil) then Exit ;
  end;

  {On vient de tout sélectionner}
  if FListe.AllSelected then begin
    if not (TFmul(Ecran).FNomFiltre = 'CPEXPORTCGA') then begin
      TotalDebit  := Valeur(THNumEdit(PCumul.FindChildControl(CumD)).Text);
      TotalCredit := Valeur(THNumEdit(PCumul.FindChildControl(CumC)).Text);
    end;
  end
  {On vient de tout désélectionner}
  else begin
    TotalDebit  := 0;
    TotalCredit := 0;
  end;
  AfficheSoldeSelection;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEXPORT.GeFactoChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  P : string;
begin
  Q := OpenSQL('SELECT PY_PAYS FROM PAYS WHERE PY_CODEISO2 = "' + CodeISOFR + '"', True);
  if not Q.EOF then P := Q.FindField('PY_PAYS').AsString
               else P := 'FRA'; {Sans garantie !!!}
  Ferme(Q);

  case GetControlText('CBGEFACTO')[1] of
    'F' : begin {Mono-Devise Euro France}
            THEdit(GetControl('T_AUXILIAIRE' )).Plus := 'T_PAYS = "' + P + '" AND T_DEVISE = "EUR"';
            THEdit(GetControl('T_AUXILIAIRE_')).Plus := 'T_PAYS = "' + P + '" AND T_DEVISE = "EUR"';
            {La quittance n'est pas la même pour les "exports"}
            {20/09/05 : FQ 16628 : en fait c'est inutile, car les deux états étaient identiques
            THValComboBox(GetControl('CBOMODELEGFF')).Plus := 'AND (MO_CODE LIKE "QI%" OR MO_CODE LIKE "BO%")';}
          end;
    'M' : begin {Multi-Devises France}
            THEdit(GetControl('T_AUXILIAIRE' )).Plus := 'T_PAYS = "' + P + '"';
            THEdit(GetControl('T_AUXILIAIRE_')).Plus := 'T_PAYS = "' + P + '"';
            {La quittance n'est pas la même pour les "exports"}
            {20/09/05 : FQ 16628 : en fait c'est inutile, car les deux états étaient identiques
             THValComboBox(GetControl('CBOMODELEGFF')).Plus := 'AND (MO_CODE LIKE "QI%" OR MO_CODE LIKE "BO%")';}
          end;
    'E' : begin {Multi-Devises Export}
            THEdit(GetControl('T_AUXILIAIRE' )).Plus := 'T_PAYS <> "' + P + '"';
            THEdit(GetControl('T_AUXILIAIRE_')).Plus := 'T_PAYS <> "' + P + '"';
            {La quittance n'est pas la même pour les "exports"}
            {20/09/05 : FQ 16628 : en fait c'est inutile, car les deux états étaient identiques
            THValComboBox(GetControl('CBOMODELEGFF')).Plus := 'AND (MO_CODE LIKE "QE%" OR MO_CODE LIKE "BO%")';}
          end;
  end;
end;

{JP 20/09/05 : FQ 16627 : Gestion de la combo type d'export GEFACTOFRANCE}
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEXPORT.MajZonesGeFactoFrance;
{---------------------------------------------------------------------------------------}
var
  cb : THValComboBox;
begin
  cb := THValComboBox(GetControl('CBGEFACTO'));
  {Demande de garantie}
  if (cboModeleEtat.Value = 'GE4') then begin
    {20/09/05 : FQ 16627 : Pour les demandes de garanties, on n'a pas besoin de choisir le type d'export}
    cb.ItemIndex := -1;
    cb.Enabled   := False;
  end
  {Quittance subrogative}
  else begin
    {20/09/05 : FQ 16627 : Pour les quittances, on a besoin de choisir le type d'export}
    if not cb.Enabled then begin
      cb.Enabled   := True;
      cb.ItemIndex := 0;
    end;
  end;
end;

{JP 20/09/05 : Il semble qu'il y ait des modifs d'agl conséquentes qui ne permette pas de
               faire Q.FLISTE := NomListe ni DBLISTE := NomListe comme recommandé par l'agl
               C'est la seule solution que j'ai trouvé sur recommandation d'Alain Rastoul.
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEXPORT.SetLaListe(NomListe : string);
{---------------------------------------------------------------------------------------}
begin
  TFMul(Ecran).SetDBListe(NomListe);
end;

{JP 22/09/05 : FQ 16623 - 16625 : Récupération des états par défaut pour GEFactoFrance}
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEXPORT.RecupEtatParDefaut(var BOR, QUIT : string);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  s : string;
begin
  BOR  := '';
  QUIT := '';
  {Recherche des états par défaut}
  Q := OpenSQL('SELECT MO_CODE FROM MODELES WHERE MO_TYPE = "E" AND MO_NATURE = "GFF" ' +
               'AND MO_DEFAUT = "X"', True);
  try
    while not Q.EOF do begin
      s := Q.FindField('MO_CODE').AsString;
           if Copy(s, 1, 2) = 'QI' then QUIT := s
      else if Copy(s, 1, 2) = 'BO' then BOR := s;
      Q.Next;
    end;
  finally
    Ferme(Q);
  end;

  {S'il n'y a pas d'état par défaut, on récupère les états de la SocRef}
  if QUIT = '' then QUIT := 'QI1';
  if BOR  = '' then BOR  := 'BOA';
end;

{JP 23/09/05 : Pour filtrer la tob en fonction de la nature de pièce pour les impressions
               C'est peut-être inutile, mais je n'arrive pas avec le LanceEtatTob à ajouter
               une clause Where sur la nature de pièce, que ce soit dans l'état ou dans la
               fonction de lancement.
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEXPORT.SetTobImpression(tDepart : TOB; var tEtat : TOB; Nat : string);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  T : TOB;
begin
  if not Assigned(tEtat) then tEtat := TOB.Create('ECRITURE_', nil, -1)
                         else tEtat.Detail.Clear;
  for n := 0 to tDepart.Detail.Count -1 do begin
    if UpperCase(tDepart.Detail[n].GetString('E_NATUREPIECE')) = Nat then begin
      T  := TOB.Create('ECRITURE_', tEtat, -1);
      T.Dupliquer(tDepart.Detail[n], True, True);
    end;
  end;
end;

{YMO 20/04/06 : FQ 17824 : Sélection sur le numéro de référence
 JP 14/05/07 : FQ 18132 : Ajout du DISTINCT
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEXPORT.OnElipsisClickE_NumCfonB(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  rest : string ;
begin

  rest := RecupWhereCritere(TFMul(Ecran).Pages);
  if Pos('ORDER BY', rest) > 0 then     // on enlève le WHERE et le ORDER BY
    rest:=copy(rest, 6, pos('ORDER BY', rest) - 6)
  else
    rest:=copy(rest,6,length(rest)-6);
  {JP 14/05/07 : exclusion des champs Null}
  rest:= rest + ' AND E_NUMCFONB <> "" AND E_NUMCFONB IS NOT NULL';

  {JP 14/05/07 : FQ 18132 : gestion du DISTINCT}
  LookUpList((Sender as THEdit), TraduireMemoire('Référence'), 'CPEXPORT', 'DISTINCT E_NUMCFONB', '',
  {JP 08/06/07 : FQ 18134 : On cache le titre (False), car le libellé dans le MCD est moyen !}
             Rest, 'E_NUMCFONB', False, 0, '');
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULEXPORT.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

//SDA le 17/12/2007
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEXPORT.CreateObjEXPORT;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF CISXPGI}
  objEXPORT := TObjCISX       .Create('oooo', nil, -1);

  ObjEXPORT.TypeExport := cfonb_CISX;
  objEXPORT.SetString('COMPTEBQ', GetControlText('CBBANQUE'));
  //objEXPORT.SetString('DOCUMENT', GetControlText('CBOMODELEBOR'));
  objEXPORT.SetString('MODELEFICHIER', GetControlText('MODELEEXPORT'));
  {$ENDIF CISXPGI}
end;
//Fin SDA le 17/12/2007

//SDA le 17/12/2007
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULEXPORT.ChargeTypesExport(strPays : string);
{---------------------------------------------------------------------------------------}
Var
  SQL : string;
  Combo : THValComboBox;
begin
  // Ne pouvant pas passer par une tablette on gère le chargement du THValComboBox.
  SQL := 'SELECT CIS_CLE,CIC_LIBELLE ' +
         'FROM CPGZIMPREQ, CPGZIMPCORRESP ' +
         'WHERE CIS_CLE<>"" AND ' +
         'CIS_CLE = CIC_IDENTIFIANT AND ' +
         'CIS_NATURE<> "X" AND ' +
         'CIS_TABLE3="' + strPays + '" AND ' +
         'CIS_CLEPAR="EXPORT" ' +
         'ORDER BY CIC_LIBELLE';
  Combo := THValComboBox(GetControl('MODELEEXPORT'));
  ChargeTHValComboBox(SQL,Combo);
  SetFocusControl('MODELEEXPORT');
end;
//Fin SDA le 17/12/2007

//SDA le 17/12/2007
{Récupère la valeur d'un champ de la grille
{---------------------------------------------------------------------------------------}
function TOF_CPMULEXPORT.GetValeur(NomChamp : string) : Variant;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(TFMul(Ecran).Q.FindField(NomChamp)) then
    Result := TFMul(Ecran).Q.FindField(NomChamp).AsVariant
  else
    Result := '';
end;
//Fin SDA le 17/12/2007

//SDA le 17/12/2007 version belge - récup procédure de BVER
procedure TOF_CPMULEXPORT.GetInfosParamSoc (var TOBFille : TOB);
var
   SQL : String;
   Q : TQuery;
begin
    if not(GetInfoParamSoc) then
    begin
      { Récupération des valeurs }
         SQL := 'SELECT SOC_NOM, SOC_DATA FROM PARAMSOC WHERE '+
                'SOC_NOM = "SO_LIBELLE" OR '+
                'SOC_NOM = "SO_ADRESSE1" OR '+
                'SOC_NOM = "SO_ADRESSE2" OR '+
                'SOC_NOM = "SO_CODEPOSTAL" OR '+
                'SOC_NOM = "SO_NIF" OR '+
                'SOC_NOM = "SO_RVA" OR '+
                'SOC_NOM = "SO_VILLE"';

        Q :=  OpenSQL(SQL,True);
        try
          while not(Q.Eof) do
          begin
             if Q.FindField('SOC_NOM').AsString = 'SO_LIBELLE' then
                Libelle := Q.FindField('SOC_DATA').AsString
             else if Q.FindField('SOC_NOM').AsString = 'SO_ADRESSE1' then
                Adresse1 := Q.FindField('SOC_DATA').AsString
             else if Q.FindField('SOC_NOM').AsString = 'SO_ADRESSE2' then
                Adresse2 := Q.FindField('SOC_DATA').AsString
             else if Q.FindField('SOC_NOM').AsString = 'SO_CODEPOSTAL' then
                CodePostal := Q.FindField('SOC_DATA').AsString
             else if Q.FindField('SOC_NOM').AsString = 'SO_NIF' then
                CodeNif := Q.FindField('SOC_DATA').AsString  
             else if Q.FindField('SOC_NOM').AsString = 'SO_RVA' then
                NumeroTVA := Q.FindField('SOC_DATA').AsString 
             else if Q.FindField('SOC_NOM').AsString = 'SO_VILLE' then
                Ville := Q.FindField('SOC_DATA').AsString;
             Q.Next;
          end
        finally
          Ferme(Q);
        end;
        GetInfoParamSoc:=true;
    end;
    TOBFille.AddChampSupValeur('SO_LIBELLE',Libelle);
    TOBFille.AddChampSupValeur('SO_ADRESSE1',Adresse1);
    TOBFille.AddChampSupValeur('SO_ADRESSE2',Adresse2);
    TOBFille.AddChampSupValeur('SO_CODEPOSTAL',CodePostal);
    TOBFille.AddChampSupValeur('SO_NIF',CodeNif);   
    TOBFille.AddChampSupValeur('SO_TVA',NumeroTVA);
    TOBFille.AddChampSupValeur('SO_VILLE',Ville);
end;


function TOF_CPMULEXPORT.ChargeMaTobTraitement (NomListe: string; TobVue : Tob) : Tob;
var
  NomChamp     : string;
  Liste_Champs_Export : string;
begin
  Result := TOB.Create(NomListe, ObjEXPORT.TobTraitement, -1);
  //ATTENTION, si un champ supplémentaire doit être ajouté, il faut IMPERATIVEMENT
  //le mettre en FIN DE LISTE
  Liste_Champs_Export  := 'BQ_NUMEMETVIR;BQ_ETABBQ;BQ_GUICHET;' +
                          'BQ_NUMEROCOMPTE;BQ_CLERIB;BQ_CODEIBAN;' +
                          'BQ_DEVISE;E_LIBELLE;E_REFINTERNE;' +
                          'E_RIB;E_DEBIT;E_DEVISE;E_AUXILIAIRE;' +
                          'E_GENERAL;E_NUMEROPIECE;G_LIBELLE;' +
                          'G_PAYS;R_DOMICILIATION;R_CODEBIC;' +
                          'R_CODEIBAN;T_LIBELLE;T_ADRESSE1;' +
                          'T_ADRESSE2;T_CODEPOSTAL;T_VILLE;' +
                          'T_PAYS;E_NUMLIGNE;E_DATECOMPTABLE;' +
                          'E_JOURNAL;E_EXERCICE;E_QUALIFPIECE;' +
                          'E_NUMECHE;G_NATUREGENE;E_MODEPAIE';
  NomChamp := ReadTokenSt(Liste_Champs_Export);
  while not (Trim(NomChamp) = '') do begin
    Result.AddChampSupValeur(NomChamp, GetValeur(NomChamp));
    NomChamp := ReadTokenSt(Liste_Champs_Export);
  end;
  { Cas Spécifique envoi vers CISX }
  if ObjEXPORT.TypeExport = cfonb_CISX then
  begin
       { Rajout des infos de la table ParamSoc }
       GetInfosParamSoc(Result);
  end;
end;



initialization
    registerclasses([TOF_CPMULEXPORT]);

end.
