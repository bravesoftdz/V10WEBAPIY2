{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 30/10/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : FEMPRUNT (FEMPRUNT)
Mots clefs ... : TOM;FEMPRUNT
*****************************************************************}
Unit TOMEMPRUNT;

Interface

Uses StdCtrls,
     Controls,
     Windows,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ELSE}
     UtileAGL,
{$ENDIF}
{$IFDEF VER150}
     Variants,
{$ENDIF}
     grids,
     graphics,
     forms,
     sysutils,
     ComCtrls,
     extctrls,
     HCtrls,
     HEnt1,
{$IFNDEF EAGLCLIENT}
     FE_Main,
     Fiche,
     FichList,
 {$ELSE}
     MaineAGL,
     eFiche,
     eFichList,
 {$ENDIF}
     HMsgBox,
     UTOM,
     Ent1,
     Spin,
     HPanel,
     HTB97,
     buttons,
     HDB,
     UObjetEmprunt,
     UTob;

Type

  TOM_FEMPRUNT = Class (TOM)
  private
    grEcheances       : THGrid;
    Pages         : TPageControl;
//    pnBoutons         : THPanel;
    pnEmprunt         : THPanel;

    btCalculer        : TBitBtn;
    btSimuler         : TBitBtn;
    btAjuster         : TBitBtn;
    btColonnes        : TBitBtn;
    btDeviseOrigine   : TBitBtn;

    btTauxVariables   : TBitBtn;
    btOrganisme       : TBitBtn;

    laRaison          : THLabel;
    laBaseAssur       : THLabel;
    laDifAss          : THLabel;

    pnTotaux          : THPanel;
    pnEcheances       : THPanel;
    pnStatut          : THPanel;

    bImprimer  : TToolBarButton97;
    bExporter  : TToolBarButton97;

    FMyOnKeyPress : TKeyPressEvent ;
    FMkOnKeyPress : TKeyPressEvent ;

    EMP_EMPPERIODE    : THValComboBox;
    EMP_TOTINT        : THNumEdit;
    EMP_TOTASS        : THNumEdit;
    EMP_TOTAMORT      : THNumEdit;
    EMP_TOTVERS       : THNumEdit;
    EMP_TOTINTASS     : THNumEdit;

{FQ21519 03/10/2007  YMO Pb d'affichage en WA}
{$IFDEF EAGLCLIENT}
    EMP_CAPITAL       : THEdit;
    EMP_DUREE         : THEdit;
    EMP_VALASSUR      : THEdit;
    EMP_LIBEMPRUNT    : THEdit;
    EMP_DATEDEBUT     : THEdit;
    EMP_NUMCOMPTE     : THEdit;
    EMP_CPTINTERCOUR  : THEdit;
    EMP_CPTCHARGAVAN  : THEdit;
    EMP_CPTFRAISFINA  : THEdit;
    EMP_CPTASSURANCE  : THEdit;
{$ELSE}
    EMP_CAPITAL       : THDBEdit;
    EMP_DUREE         : THDBEdit;
    EMP_VALASSUR      : THDBEdit;
    EMP_LIBEMPRUNT    : THDBEdit;
    EMP_DATEDEBUT     : THDBEdit;
    EMP_NUMCOMPTE     : THDBEdit;
    EMP_CPTINTERCOUR  : THDBEdit;
    EMP_CPTCHARGAVAN  : THDBEdit;
    EMP_CPTFRAISFINA  : THDBEdit;
    EMP_CPTASSURANCE  : THDBEdit;
{$ENDIF}


    fOmEmprunt : TObjetEmprunt;

    pnEmpruntWidth, pnEmpruntHeight : Integer;

    fInOldRow       : Integer;
    fInOldCol       : Integer;
    fStOldCellValue : String;

    fStOldDateDebut : String;

    fObClignoTimer  : TTimer;
    fBoClignoOn     : Boolean;

    fBoSaisirMontant : Boolean;

    fEnCreation : Boolean;  {Permet de gérer le focus perturbé par le OnComboChange en création}

    procedure InitControls;
    procedure SaisirMontant(pStFieldName : String);
    procedure AppliquerMonnaieSaisie;
    Function  InfoDevise(pStCodeDevise : String; Var pStLibelle, pStSymbole : String) : String;
    Function  Confirmer(pBoWithPrompt : Boolean) : Boolean;
    Function  SaisieValide(pBoWithMessage : Boolean) : Boolean;
    Function  Calculer(pBoWithPrompt : Boolean) : Boolean;
    Procedure EcheancesVersGrille;
    Procedure EcranVersEmprunt;
    Procedure EmpruntVersEcran;
    Procedure GrilleVersEcheances;
    procedure FocaliserGrille;
    procedure ChangeStatutEmprunt(pInStatut : TStatutEmprunt);
    procedure Clignoter(pBoStatus : Boolean);
    procedure ClignoTimer(Sender: TObject);
    function  RemplirComboOrganismes : Boolean;

    Procedure Modifier;
    Procedure Imprimer(pBoListeExport : Boolean = False);

    Procedure RepercuterModifTableau(pInCol : Integer; pInRow : Integer; pStOldValue : String);
    Procedure ModifierTotal(pInCol : Integer; pInRow : Integer; pRdDelta : Double);

    Function  GetDeviseEmprunt : String;
    Procedure SetDeviseEmprunt(pValue : String);
    Function  GetDeviseSaisie : String;
    Procedure SetDeviseSaisie(pValue : String);

    procedure bImprimerClick(Sender : TObject);
    procedure bExporterClick(Sender : TObject);

    procedure btCalculerClick(Sender: TObject);
    procedure btAjusterClick(Sender: TObject);
    procedure btSimulerClick(Sender: TObject);
    procedure btDeviseOrigineClick(Sender: TObject);

    Procedure btTauxVariablesClick(Sender: TObject);
    Procedure btOrganismeClick(Sender: TObject);

    procedure OnCtrlKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure OnDureeExit(Sender: TObject);
    procedure OnCompteExit(Sender: TObject);
    procedure OnDebutEnter(Sender: TObject);
//    procedure OnDebutExit(Sender: TObject);
    procedure OnLibelleExit(Sender: TObject);
    procedure OnComboChange(Sender : TObject);

    procedure OnMontantKeyPress(Sender: TObject; var Key: Char);
    procedure VerifGene(Compte : String ; Existe, Seize : boolean ) ;
    procedure OnPageControlChange(Sender: TObject);
    procedure OnPageControlChanging(Sender: TObject; var AllowChange: Boolean);
    procedure grEcheancesGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    procedure FormResize(Sender: TObject);
    procedure grEcheancesSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);

  public
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;

    Procedure PutValue(pStFieldName : String; pVaValue : Variant);

    Property DeviseEmprunt : String Read GetDeviseEmprunt Write SetDeviseEmprunt;
    Property DeviseSaisie  : String Read GetDeviseSaisie  Write SetDeviseSaisie;
  end ;



const
   // Pages
   cInEmpPage = 0;
   cInEchPage = 1;



procedure NouvelEmprunt(pStCompte : String = ''; pRdCapital : Double = 0);
procedure SaisieEmprunt(pInNumEmprunt : Integer);

Implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPTypeCons,
  {$ENDIF MODENT1}
  TOFEMPIMP,
  TOFEMPCHOIXIMP,
  TOFEMPMONTANT,
  TOFEMPAJUST,
  TOFTYPESIM,
  TOFEMPCST,
  TOFTAUXVAR,
  TOFEMPMONORIG,
  dpTofAnnuSel,
  uTomAnn, // ATTENTION : pour ne pas oublier la TOM ANNUAIRE :INDISPENSABLE
  UCreCommun,
  Lookup,
  paramsoc;



procedure NouvelEmprunt(pStCompte : String = ''; pRdCapital : Double = 0);
var
   lStArgs : String;
begin
//   lStArgs := 'ACTION=CREATION;SELECT EMP_CODEEMPRUNT FROM FEMPRUNT ORDER BY EMP_CODEEMPRUNT;' +
   lStArgs := 'ACTION=CREATION;' + pStCompte;
   if pRdCapital > 0 then
      lStArgs := lStArgs + ';' + FloatToStr(pRdCapital);

   AGLLanceFiche('FP','FEMPRUNT','','',lStArgs);
end;


procedure SaisieEmprunt(pInNumEmprunt : Integer);
begin
{
   AGLLanceFiche('FP',
                 'FEMPRUNT',
                 '',
                 IntToStr(pInNumEmprunt),
                 'ACTION=MODIFICATION;SELECT EMP_CODEEMPRUNT FROM FEMPRUNT ORDER BY EMP_CODEEMPRUNT');
}
   AGLLanceFiche('FP',
                 'FEMPRUNT',
                 '',
                 IntToStr(pInNumEmprunt),
                 '');
end;

// ----------------------------------------------------------------------------
// Nom    : OnArgument
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.OnArgument ( S: String ) ;
var
//  lStNumEmprunt : String;
  lStNumCompte  : String;
  lStMontant    : String;
begin
  Inherited ;

  // Création des objets
  fOmEmprunt := TObjetEmprunt.CreerEmprunt;
//  fOmEmprunt.DateFinExercice := VH^.EnCours.Fin;

  fObClignoTimer          := TTimer.create(Nil);
  fObClignoTimer.Interval := 500;
  fObClignoTimer.OnTimer  := ClignoTimer;

  fEnCreation:=False;
  // Devises
  gSymbDevNull    := '---';
  gCodeDevEuro    := 'EUR';
  gCodeDevFranc   := 'FRF';
  InfoDevise(gCodeDevEuro,gSymbDevEuro,gSymbDevEuro);
  InfoDevise(gCodeDevFranc,gSymbDevFranc,gSymbDevFranc);

  gCodeDevDossier := V_PGI.DevisePivot;
  if gCodeDevDossier = '' then
     gCodeDevDossier := gCodeDevEuro;

  // Init des controles
  InitControls;

  // Traitement des arguments
//  ReadTokenSt(S);
//  ReadTokenSt(S);
  ReadTokenSt(S);
  lStNumCompte  := ReadTokenSt(S);
  lStMontant    := ReadTokenSt(S);

  //
  if lStNumCompte <> '' then
  begin
     SetControlEnabled('EMP_NUMCOMPTE', False);
     PutValue('EMP_NUMCOMPTE', lStNumCompte);
  end;

  //
  if lStMontant <> '' then
     PutValue('EMP_CAPITAL', Valeur(lStMontant));

  // La saisie de l'organisme liée à l'annuaire n'est disponible qu'en PCL
  SetControlVisible('BTORGANISME',ctxPCL in V_PGI.PGICOntexte);
  SetControlVisible('TEMP_ORGPRETEUR',ctxPCL in V_PGI.PGICOntexte);
  SetControlVisible('EMP_GUIDORGPRETEUR',ctxPCL in V_PGI.PGICOntexte);

  {FQ21516 03/10/2007  YMO Pb d'affichage en WA}
  setcontrolvisible('bImprimer', true);
end ;

// ----------------------------------------------------------------------------
// Nom    : OnClose
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.OnClose ;
begin
  fOmEmprunt.Free;
  fObClignoTimer.Free;
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/03/2007
Modifié le ... :   /  /
Description .. : Vérifie la saisie du numero de compte
Mots clefs ... :
*****************************************************************}
procedure TOM_FEMPRUNT.VerifGene(Compte : String ; Existe, Seize : boolean ) ;
var LeCompte : String;
begin

  if Existe and (Trim(GetControlText(Compte))='') then
  begin
    PgiInfo('Veuillez renseigner un compte.','Vérification du compte saisi');
    SetFocusControl(Compte);
    Abort;
  end;

  LeCompte:=THEdit(GetControl(Compte)).Text;
  {FQ20541  22.06.07  YMO}
  THEdit(GetControl(Compte)).Text:=BourreEtLess( LeCompte, fbGene ) ;

  if Trim(GetControlText(Compte))<>'' then
  begin

    if not Presence('GENERAUX', 'G_GENERAL', GetControlText(Compte)) then
    begin
      PgiInfo('Le compte général sélectionné n''existe pas.','Vérification du compte saisi');
      SetFocusControl(Compte);
      Abort;
    end;

    {FQ20803  21.06.07  YMO}
    if Seize And (Copy(GetControlText(Compte),1,2)<>'16') then
    begin
      PgiInfo('Le compte général sélectionné doit avoir ''16'' comme racine','Vérification du compte saisi');
      SetFocusControl(Compte);
      Abort;
    end;

  end;

end;

// ----------------------------------------------------------------------------
// Nom    : OnUpdateRecord
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.OnUpdateRecord ;
begin
  // Détermination du numéro d'emprunt si nécéssaire
  if GetField('EMP_CODEEMPRUNT') <= 0 then
  begin
     fOmEmprunt.ValeurPK := fOmEmprunt.EmpruntSuivant;
     PutValue('EMP_CODEEMPRUNT', fOmEmprunt.GetValue('EMP_CODEEMPRUNT'));
  end;

  //Vérification comptes
  VerifGene('EMP_NUMCOMPTE', true, true);
  VerifGene('EMP_CPTINTERCOUR', GetField('EMP_STATUT') <> seSimule, false);
  VerifGene('EMP_CPTCHARGAVAN', GetField('EMP_STATUT') <> seSimule, false);
  VerifGene('EMP_CPTFRAISFINA', GetField('EMP_STATUT') <> seSimule, false);
  VerifGene('EMP_CPTASSURANCE', GetField('EMP_STATUT') <> seSimule, false);

  //Date Par défaut
  PutValue('EMP_DERDATEGENEFF', iDate1900);

  // Lancement du calcul ou récupération de la saisie des éhéances
  if Integer(GetField('EMP_STATUT')) in [Ord(seVide), Ord(seModifie)] then
     Calculer(False)
  else
     if Integer(GetField('EMP_STATUT')) = Ord(seSaisi) then
        GrilleVersEcheances;

  // Enregistrement entête
  Inherited ;

  // Enregistrement des échéances et des taux variables
  fOmEmprunt.DBSupprimeEtRemplaceEcheances;
end ;

// ----------------------------------------------------------------------------
// Nom    : OnAfterUpdateRecord
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

// ----------------------------------------------------------------------------
// Nom    : OnLoadRecord
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.OnLoadRecord ;
begin
  Inherited ;
  DeviseEmprunt := GetField('EMP_DEVISE');
  DeviseSaisie  := GetField('EMP_DEVISESAISIE');
  fOmEmprunt.DBChargerEmprunt(GetField('EMP_CODEEMPRUNT'));
  EcheancesVersGrille;
  Pages.ActivePageIndex := cInEmpPage;
  SetFocusControl('EMP_LIBEMPRUNT');
  ChangeStatutEmprunt(GetField('EMP_STATUT'));

  //
  if GetField('EMP_ORGPRETEUR') = Null then
     PutValue('EMP_ORGPRETEUR',-1);

  if GetField('EMP_GUIDORGPRETEUR') = Null then
     PutValue('EMP_GUIDORGPRETEUR',-2);

  fStOldDateDebut := GetControlText('EMP_DATEDEBUT');
  SetFocusControl('EMP_LIBEMPRUNT');
end ;

// ----------------------------------------------------------------------------
// Nom    : OnNewRecord
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.OnNewRecord ;
begin
  Inherited ;
  fOmEmprunt.Init;
  fOmEmprunt.PutValue('EMP_DEVISE', gCodeDevDossier);
  fOmEmprunt.PutValue('EMP_DEVISESAISIE', '');

  {renseignement des comptes auto à partir des paramsoc (partie ICC/CRE)}
  fOmEmprunt.PutValue('EMP_CPTINTERCOUR', GetParamSocSecur ('SO_CPINTERCOUR', False));
  fOmEmprunt.PutValue('EMP_CPTCHARGAVAN', GetParamSocSecur ('SO_CPCHARGAVAN', False));
  fOmEmprunt.PutValue('EMP_CPTFRAISFINA', GetParamSocSecur ('SO_CPFRAISFINA', False));
  fOmEmprunt.PutValue('EMP_CPTASSURANCE', GetParamSocSecur ('SO_CPASSURANCE', False));

//  DeviseEmprunt := gCodeDevDossier;
//  DeviseSaisie  := '';
  fEnCreation:=True;

  EmpruntVersEcran;
  Pages.ActivePageIndex := cInEmpPage;
  SetFocusControl('EMP_LIBEMPRUNT');
  ChangeStatutEmprunt(GetField('EMP_STATUT'));
  SetControlEnabled('EMP_CODEEMPRUNT',False);
  end ;

// ----------------------------------------------------------------------------
// Nom    : OnDeleteRecord
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.OnDeleteRecord ;
begin
  Inherited ;
  fOmEmprunt.DBSupprimerEcheances;
  Ecran.Close;
end ;

// ----------------------------------------------------------------------------
// Nom    : OnChangeField
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.OnChangeField ( F: TField ) ;
var
  lBoVisible : Boolean;
//  lInMinValue, lInMaxValue : Integer;
  lRdTaux : Double;
begin
  Inherited ;

  // Change le statut de l'emprunt
  if ( (F.FieldName = 'EMP_DATEDEBUT') or
       (F.FieldName = 'EMP_CAPITAL') or
       (F.FieldName = 'EMP_EMPPERIODE') or
       (F.FieldName = 'EMP_DUREE') or
       (F.FieldName = 'EMP_EMPTYPETAUX') or
       (F.FieldName = 'EMP_TAUXAN') or
       (F.FieldName = 'EMP_RAISON') or
       (F.FieldName = 'EMP_EMPBQEQ') or
       (F.FieldName = 'EMP_EMPTYPEASSUR') or
       (F.FieldName = 'EMP_VALASSUR') or
       (F.FieldName = 'EMP_EMPBASEASSUR') or
       (F.FieldName = 'EMP_DIFREMB') or
       (F.FieldName = 'EMP_DIFINT') or
       (F.FieldName = 'EMP_DIFASS') or
       (F.FieldName = 'EMP_EMPTYPEVERS') or
       (F.FieldName = 'EMP_AMORTCST') or
       (F.FieldName = 'EMP_VERSCST') or
       (F.FieldName = 'EMP_AMORTTXINI') or
       (F.FieldName = 'EMP_ACTUANNU') ) and
     (DS.State in [dsEdit, dsInsert]) then
     if Integer(GetField('EMP_STATUT')) in [Ord(seCalcule),Ord(seVide)] then
        ChangeStatutEmprunt(seModifie);

  // Répercutions du changement du type de taux
  if F.FieldName = 'EMP_EMPTYPETAUX' then
  begin
     lBoVisible := (GetField('EMP_EMPTYPETAUX') = ttArythmetique) or
                   (GetField('EMP_EMPTYPETAUX') = ttGeometrique);
     SetControlVisible('EMP_RAISON', lBoVisible);
     SetControlVisible('laRaison', lBoVisible);
     if not lBoVisible then PutValue('EMP_RAISON', 0);

     SetControlEnabled('EMP_AMORTTXINI', GetField('EMP_EMPTYPETAUX') = ttVariable);
     SetControlVisible('btTauxVariables', GetField('EMP_EMPTYPETAUX') = ttVariable);
     SetControlEnabled('btSimuler', GetField('EMP_EMPTYPETAUX') <> ttVariable);
  end;

  // Cache la saisie de l'assurance si aucune assurance
  if F.FieldName = 'EMP_EMPTYPEASSUR' then
  begin
     lBoVisible := GetField('EMP_EMPTYPEASSUR') <> taAucune;
     SetControlVisible('EMP_VALASSUR', lBoVisible);
     If lBoVisible then SetFocusControl('EMP_VALASSUR');  {FQ20157  31.05.07  YMO}
     SetControlVisible('EMP_EMPBASEASSUR', lBoVisible);
//     SetControlVisible('laBaseAssur', lBoVisible);
     SetControlVisible('EMP_DIFASS', lBoVisible);
     SetControlVisible('laDifAss', lBoVisible);

     lBoVisible := GetField('EMP_EMPTYPEASSUR') = taTaux;
     SetControlVisible('EMP_EMPBASEASSUR', lBoVisible );
     SetControlVisible('laBaseAssur', lBoVisible );

     if GetField('EMP_EMPTYPEASSUR') = taMontant then
     begin
//         EMP_VALASSUR.NumericType := ntGeneral;
         EMP_VALASSUR.DisplayFormat := '# ##0.00';
         FMyOnKeyPress := EMP_VALASSUR.OnKeyPress;
         EMP_VALASSUR.OnKeyPress := OnMontantKeyPress;
     end
     else
     begin
//         EMP_VALASSUR.NumericType := ntPercentage;
         EMP_VALASSUR.DisplayFormat := '##0.0000';
       //  EMP_VALASSUR.OnKeyPress  := Nil; {FQ20157  31.05.07  YMO}
         EMP_VALASSUR.Hint        := '';
     end;

     {FQ21141 23.07.07  YMO}
     FMkOnKeyPress := EMP_CAPITAL.OnKeyPress;
     EMP_CAPITAL.OnKeyPress := OnMontantKeyPress;

     if GetField('EMP_EMPTYPEASSUR') = taTaux then
     begin
        grEcheances.ColWidths[cColTxAssurance]  := cInLargColTaux;
        grEcheances.ColWidths[cColTxGlobal]     := cInLargColTaux;
        grEcheances.ColLengths[cColTxAssurance] := 0;
     end
     else
     begin
        grEcheances.ColWidths[cColTxAssurance]  := -1;
        grEcheances.ColWidths[cColTxGlobal]     := -1;
        grEcheances.ColLengths[cColTxAssurance] := -1;
     end;

//     PutValue('EMP_VALASSUR', 0);
     AppliquerMonnaieSaisie;
  end;

  // Réinitialise la devise d'origine si on met les montants à 0
  if ( (F.FieldName = 'EMP_EMPTYPEASSUR') or (F.FieldName = 'EMP_CAPITAL') or ((F.FieldName = 'EMP_VALASSUR') and (GetField('EMP_EMPTYPEASSUR') = taMontant)) ) and
     (GetField('EMP_CAPITAL') = 0) and ((GetField('EMP_VALASSUR') = 0) or (GetField('EMP_EMPTYPEASSUR') <> taMontant)) then
  begin
     DeviseSaisie := '';
  end;

  // Conversion en monaie d'origine en Hint pour les ctrls des montants (capital et assurance)
  if (F.FieldName = 'EMP_CAPITAL') or ((F.FieldName = 'EMP_VALASSUR') and (GetField('EMP_EMPTYPEASSUR') = taMontant)) then
     if (DeviseSaisie <> DeviseEmprunt) and (DeviseSaisie <> '') then
     begin
        if DeviseSaisie = gCodeDevEuro then
           lRdTaux := 1/V_PGI.TauxEuro
        else
           lRdTaux := V_PGI.TauxEuro;
        Try
           SetControlProperty(F.FieldName, 'Hint', fOmEmprunt.FormaterMontant(GetField(F.FieldName) * lRdTaux) + ' ' + gSymbDevSaisie);
        except
           SetControlProperty(F.FieldName, 'Hint', '');
        end;
     end
     else
        SetControlProperty(F.FieldName, 'Hint', '');

  // Active la checkbox EMP_ACTUANNU si différé
  if (F.FieldName = 'EMP_DIFREMB') or (F.FieldName = 'EMP_DIFINT') or (F.FieldName = 'EMP_DIFASS') then
     SetControlEnabled('EMP_ACTUANNU', (GetField('EMP_DIFREMB') <> 0) or
                                       (GetField('EMP_DIFINT') <> 0) or
                                       (GetField('EMP_DIFASS') <> 0) );

  {FQ19824 YMO Date debut>=Date contrat 31.05.07}
  if (F.FieldName = 'EMP_DATECONTRAT') and (GetField('EMP_DATEDEBUT') < GetField('EMP_DATECONTRAT')) then
       PutValue('EMP_DATEDEBUT', GetField('EMP_DATECONTRAT'));

end ;

// ----------------------------------------------------------------------------
// Nom    : OnCancelRecord
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.OnCancelRecord ;
begin
  Inherited ;
  OnLoadRecord;
end ;


// ----------------------------------------------------------------------------
// Nom    : InitControls
// Date   : 07/12/2001
// Auteur : D. ZEDIAR
// Objet  : Initiallisation des controles de la fenêtre
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.InitControls;
begin
   // Récup des ctrls
   grEcheances       := THGrid(GetControl('grEcheances'));
   Pages             := TPageControl(GetControl('Pages'));
   pnEmprunt         := THPanel(GetControl('pnEmprunt'));
   pnTotaux          := THPanel(GetControl('pnTotaux'));
   pnEcheances       := THPanel(GetControl('pnEcheances'));
   pnStatut          := THPanel(GetControl('pnStatut'));

   pnEmpruntWidth  := pnEmprunt.Width;
   pnEmpruntHeight := pnEmprunt.Height;

   bExporter  := TToolBarButton97(GetControl('bExport'));
   bImprimer  := TToolBarButton97(GetControl('bImprimer'));

   btCalculer      := TBitBtn(GetControl('btCalculer'));
   btColonnes      := TBitBtn(GetControl('btColonnes'));
   btSimuler       := TBitBtn(GetControl('btSimuler'));
   btAjuster       := TBitBtn(GetControl('btAjuster'));
   btDeviseOrigine := TBitBtn(GetControl('btDeviseOrigine'));

   btTauxVariables := TBitBtn(GetControl('btTauxVariables'));
   btOrganisme     := TBitBtn(GetControl('btOrganisme'));

   laRaison        := THLabel(GetControl('laRaison'));
   laBaseAssur     := THLabel(GetControl('laBaseAssur'));
   laDifAss        := THLabel(GetControl('laDifAss'));

   EMP_EMPPERIODE    := THValComboBox(GetControl('EMP_EMPPERIODE'));

{$IFDEF EAGLCLIENT}
   EMP_NUMCOMPTE     := THEdit(GetControl('EMP_NUMCOMPTE'));
   EMP_CPTINTERCOUR  := THEdit(GetControl('EMP_CPTINTERCOUR'));
   EMP_CPTCHARGAVAN  := THEdit(GetControl('EMP_CPTCHARGAVAN'));
   EMP_CPTFRAISFINA  := THEdit(GetControl('EMP_CPTFRAISFINA'));
   EMP_CPTASSURANCE  := THEdit(GetControl('EMP_CPTASSURANCE'));
   EMP_CAPITAL       := THEdit(GetControl('EMP_CAPITAL'));
   EMP_DUREE         := THEdit(GetControl('EMP_DUREE'));
   EMP_VALASSUR      := THEdit(GetControl('EMP_VALASSUR'));
   EMP_LIBEMPRUNT    := THEdit(GetControl('EMP_LIBEMPRUNT'));
   EMP_DATEDEBUT     := THEdit(GetControl('EMP_DATEDEBUT'));
{$ELSE}
   EMP_NUMCOMPTE     := THDBEdit(GetControl('EMP_NUMCOMPTE'));
   EMP_CPTINTERCOUR  := THDBEdit(GetControl('EMP_CPTINTERCOUR'));
   EMP_CPTCHARGAVAN  := THDBEdit(GetControl('EMP_CPTCHARGAVAN'));
   EMP_CPTFRAISFINA  := THDBEdit(GetControl('EMP_CPTFRAISFINA'));
   EMP_CPTASSURANCE  := THDBEdit(GetControl('EMP_CPTASSURANCE'));
   EMP_CAPITAL       := THDBEdit(GetControl('EMP_CAPITAL'));
   EMP_DUREE         := THDBEdit(GetControl('EMP_DUREE'));
   EMP_VALASSUR      := THDBEdit(GetControl('EMP_VALASSUR'));
   EMP_LIBEMPRUNT    := THDBEdit(GetControl('EMP_LIBEMPRUNT'));
   EMP_DATEDEBUT     := THDBEdit(GetControl('EMP_DATEDEBUT'));
{$ENDIF}

   EMP_TOTINT        := THNumEdit(GetControl('EMP_TOTINT'));
   EMP_TOTASS        := THNumEdit(GetControl('EMP_TOTASS'));
   EMP_TOTAMORT      := THNumEdit(GetControl('EMP_TOTAMORT'));
   EMP_TOTVERS       := THNumEdit(GetControl('EMP_TOTVERS'));
   EMP_TOTINTASS     := THNumEdit(GetControl('EMP_TOTINTASS'));

   SetControlProperty('EMP_TAUXAN','DisplayFormat','##0.0000');

   //
   SetActiveTabSheet('PGeneral');

   TShape(GetControl('Shape1')).Brush.Color := clBtnFace;
   TShape(GetControl('Shape2')).Brush.Color := clBtnFace;
   TShape(GetControl('Shape3')).Brush.Color := clBtnFace;
   TShape(GetControl('Shape1')).Pen.Color   := clBtnShadow;
   TShape(GetControl('Shape2')).Pen.Color   := clBtnShadow;
   TShape(GetControl('Shape3')).Pen.Color   := clBtnShadow;

   Ecran.OnResize := FormResize;

   bImprimer.OnClick   := bImprimerClick;
   bExporter.OnClick   := bExporterClick;

   // Init de la grille des échéances
   grEcheances.Options := [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goEditing,
                           goColSizing, goTabs, goThumbTracking, goTabs, goAlwaysShowEditor];
   // Init de la grille
   grEcheances.DefaultRowHeight := 18;
   grEcheances.ColCount         := cNbColEcheances;
   grEcheances.RowCount         := 2;
   grEcheances.RowCount         := 2;
   grEcheances.FixedCols        := 1;
   grEcheances.FixedRows        := 1;
   grEcheances.TitleBold        := True;
   grEcheances.TitleCenter      := True;

   grEcheances.GetCellCanvas := grEcheancesGetCellCanvas;

   grEcheances.ColWidths[cColNumEch]        := 36;
   grEcheances.ColWidths[cColDate]          := cInLargColDate;
   grEcheances.ColWidths[cColTxInteret]     := cInLargColTaux;
   grEcheances.ColWidths[cColTxAssurance]   := cInLargColTaux;
   grEcheances.ColWidths[cColTxGlobal]      := cInLargColTaux;
   grEcheances.ColWidths[cColAmortissement] := cInLargColMnt;
   grEcheances.ColWidths[cColInteret]       := cInLargColMnt;
   grEcheances.ColWidths[cColAssurance]     := cInLargColMnt;
   grEcheances.ColWidths[cColTotInteret]    := cInLargColMnt;
   grEcheances.ColWidths[cColVersement]     := cInLargColMnt;
   grEcheances.ColWidths[cColSolde]         := cInLargColMnt;
   grEcheances.ColWidths[cColReport]        := -1;
   grEcheances.ColWidths[cColDette]         := -1;
   grEcheances.ColWidths[cColCumulA]        := -1;
   grEcheances.ColWidths[cColCumulI]        := -1;
   grEcheances.ColWidths[cColCumulM]        := -1;
   grEcheances.ColWidths[cColTypeLigne]     := -1;
   grEcheances.ColWidths[cColModif]         := -1;

   grEcheances.ColEditables[cColReport]       := False;
   grEcheances.ColEditables[cColDette]        := False;
   grEcheances.ColEditables[cColCumulA]       := False;
   grEcheances.ColEditables[cColCumulI]       := False;
   grEcheances.ColEditables[cColCumulM]       := False;
   grEcheances.ColEditables[cColTypeLigne]    := False;
   grEcheances.ColEditables[cColModif]        := False;
   grEcheances.ColEditables[cColNumEch]       := False;
   grEcheances.ColEditables[cColDate]         := False;
   grEcheances.ColEditables[cColTxGlobal]     := False;
   grEcheances.ColEditables[cColTotInteret]   := False;

   grEcheances.ColLengths[cColReport]       := -1;
   grEcheances.ColLengths[cColDette]        := -1;
   grEcheances.ColLengths[cColCumulA]       := -1;
   grEcheances.ColLengths[cColCumulI]       := -1;
   grEcheances.ColLengths[cColCumulM]       := -1;
   grEcheances.ColLengths[cColTypeLigne]    := -1;
   grEcheances.ColLengths[cColModif]        := -1;
   grEcheances.ColLengths[cColNumEch]       := -1;
   grEcheances.ColLengths[cColDate]         := -1;
   grEcheances.ColLengths[cColTxGlobal]     := -1;
   grEcheances.ColLengths[cColTotInteret]   := -1;

   // Evenements
   btCalculer.OnClick      := btCalculerClick;
   btAjuster.OnClick       := btAjusterClick;
   btSimuler.OnClick       := btSimulerClick;
   btDeviseOrigine.OnClick := btDeviseOrigineClick;

   EMP_CAPITAL.OnKeyDown  := OnCtrlKeyDown;
   EMP_VALASSUR.OnKeyDown := OnCtrlKeyDown;

   EMP_DUREE.OnExit       := OnDureeExit;
   EMP_EMPPERIODE.OnExit  := OnDureeExit;
   EMP_DATEDEBUT.OnEnter  := OnDebutEnter;
//   EMP_DATEDEBUT.OnExit   := OnDebutExit;

   EMP_LIBEMPRUNT.OnExit   := OnLibelleExit;
   EMP_NUMCOMPTE.OnExit    := OnCompteExit;
   EMP_CPTINTERCOUR.OnExit    := OnCompteExit;
   EMP_CPTCHARGAVAN.OnExit    := OnCompteExit;
   EMP_CPTFRAISFINA.OnExit    := OnCompteExit;
   EMP_CPTASSURANCE.OnExit    := OnCompteExit;

   THValComboBox(GetControl('EMP_EMPTYPEASSUR')).OnChange := OnComboChange;
   THValComboBox(GetControl('EMP_EMPTYPETAUX')).OnChange  := OnComboChange;

   btTauxVariables.OnClick := btTauxVariablesClick;
   btOrganisme.OnClick     := btOrganismeClick;

   Pages.OnChange   := OnPageControlChange;
   Pages.OnChanging := OnPageControlChanging;

   SetControlEnabled('EMP_CODEEMPRUNT',False);

   grEcheances.OnSelectCell := grEcheancesSelectCell;

   RemplirComboOrganismes;

   // taille de certains labels
{   for lInCpt := 1 to 5 do
   begin
      laLabel := THLabel(GetControl('LALABEL'+inttostr(lInCpt)));
      if Assigned(laLabel) then
      begin
         lInRight := laLabel.Left + laLabel.Width;
         laLabel.AutoSize := True;
         laLabel.Left := lInRight - laLabel.Width;
      end;
   end;
}

end;


// Ajoute l'élément '-1;Aucun' à la combo des organismes prêteurs
function  TOM_FEMPRUNT.RemplirComboOrganismes : Boolean;
begin
   with THValComboBox(GetControl('EMP_GUIDORGPRETEUR')) do
   begin
      ReLoad;
      Result := Values.Count > 0;
      Values.Insert(0,'-1');
      Items.Insert(0,'<<Aucun>>');
      if not Result then
         Value := '-1';
   end;
end;


// ----------------------------------------------------------------------------
// Nom    : PutValue
// Date   : 06/11/2002
// Auteur : D. ZEDIAR
// Objet  : Met à jour l'objet emprunt et l'e dataset en passant en mode edit
//          si nécéssaire
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TOM_FEMPRUNT.PutValue(pStFieldName : String; pVaValue : Variant);
begin
   if GetField(pStFieldName) <> pVaValue then
   begin
      Modifier;
      SetField(pStFieldName,pVaValue);
   end;
   fOmEmprunt.PutValue(pStFieldName,pVaValue);
end;


procedure TOM_FEMPRUNT.bImprimerClick(Sender : TObject);
begin
   Imprimer;
end;


procedure TOM_FEMPRUNT.bExporterClick(Sender : TObject);
begin
  Imprimer(True);
end;

procedure TOM_FEMPRUNT.btCalculerClick(Sender: TObject);
begin
   Calculer(True);
//   if not Calculer(True) then Exit;
end;

// Calcul théorique
Function TOM_FEMPRUNT.Calculer(pBoWithPrompt : Boolean) : Boolean;
var
   lRdMontant            : Double;
   lInNbEcheancesSaisies : Integer;
   lObEmprunACalculer    : TObjetEmprunt;
   lInCpt1, lInCpt2      : Integer;
   lObEch1,lObEch2       : TOB;
   lStNomChamp           : String;
   lRdAs                 : Double;
   lRdAm                 : Double;
   lRdIn                 : Double;
   lRdVe                 : Double;
   lRdCumulA             : Double;
   lRdCumulI             : Double;
   lRdCumulV             : Double;
   lStMessage            : String;

   Procedure MinorerReport(pStNomChamp : String);
   begin
      if lObEmprunACalculer.GetValue(pStNomChamp) < lInNbEcheancesSaisies then
         lObEmprunACalculer.PutValue(pStNomChamp,0)
      else
         lObEmprunACalculer.PutValue(pStNomChamp,lObEmprunACalculer.GetValue(pStNomChamp) - lInNbEcheancesSaisies);
   end;
begin
   Result := False;

   Pages.OnChange := Nil;

   // Controles divers
   if not SaisieValide(pBoWithPrompt) then Exit;

   // Ecran vers objet emprunt
   EcranVersEmprunt;

   // L'utilisateur veut-il conserver les pemières échéances saisie
   lObEmprunACalculer := fOmEmprunt;
   lInNbEcheancesSaisies := 0;
   While grEcheances.Cells[cColModif,lInNbEcheancesSaisies+1] = 'X' do Inc(lInNbEcheancesSaisies);
   if lInNbEcheancesSaisies > 0 then
   begin
      if ( (GetField('EMP_STATUT') = seSaisi) or
           (GetField('EMP_STATUT') = seSimule) ) and
         (not pBoWithPrompt) then Exit;

      lStMessage := 'Le tableau d''amortissement a été ajusté manuellement.' + chr(13) +
                    'Souhaitez-vous calculer l''emprunt à partir de l''échéance n° ' + IntTostr(lInNbEcheancesSaisies+1) + ' ?';

      case PGIAskCancel(lStMessage,TitreHalley) of
            mrYes :
            begin
               lObEmprunACalculer := TObjetEmprunt.CreerEmprunt;
               lObEmprunACalculer.Assigner(fOmEmprunt);
               lObEmprunACalculer.PutValue('EMP_CODEEMPRUNT',fOmEmprunt.GetValue('EMP_CODEEMPRUNT'));
               lObEmprunACalculer.PutValue('EMP_CAPITAL', Valeur(grEcheances.Cells[cColSolde,lInNbEcheancesSaisies]));
               lObEmprunACalculer.PutValue('EMP_DATEDEBUT', fOmEmprunt.Echeances.Detail[lInNbEcheancesSaisies].GetValue('ECH_DATE'));
               lObEmprunACalculer.PutValue('EMP_DUREE', fOmEmprunt.GetValue('EMP_DUREE') - fOmEmprunt.EMP_NBMOISECH * lInNbEcheancesSaisies);

               MinorerReport('EMP_DIFREMB');
               MinorerReport('EMP_DIFINT');
               MinorerReport('EMP_DIFASS');
            end;
            mrCancel : Exit;
      end;
   end
   else
      if not Confirmer(pBoWithPrompt) then Exit;

   // Saisie du montant de l'amortissement constant
   if lObEmprunACalculer.GetValue('EMP_EMPTYPEVERS') = tvVariable then
   begin
      if pBoWithPrompt then
      begin
         lRdMontant := SaisieMontantConstant(True,lObEmprunACalculer.CalcAmortCst);
         if lRdMontant <= 0 then
         begin
            if lObEmprunACalculer <> fOmEmprunt then lObEmprunACalculer.Free;
            Exit;
         end;
         lObEmprunACalculer.PutValue('EMP_AMORTCST', lRdMontant);
      end
      else
         lObEmprunACalculer.PutValue('EMP_AMORTCST', fOmEmprunt.CalcAmortCst);
   end;

   //  Calcul
   lObEmprunACalculer.Simulation(trVersement);

   // Traitement d'un emprunt partiellement saisi / partiellement calculé
   if lObEmprunACalculer <> fOmEmprunt then
   begin
      // Transfert des échéances de l'emprunt calculé vers l'emprunt principal
      for lInCpt1 := lInNbEcheancesSaisies to fOmEmprunt.Echeances.Detail.Count - 1 do
      begin
         lObEch1 := fOmEmprunt.Echeances.Detail[lInCpt1];
         lObEch2 := lObEmprunACalculer.Echeances.Detail[lInCpt1-lInNbEcheancesSaisies];
         For lInCpt2 := 1 to lObEch1.NbChamps do
         begin
             lStNomChamp := lObEch1.GetNomChamp(lInCpt2);
             lObEch1.PutValue( lStNomChamp,
                               lObEch2.GetValue(lStNomChamp) );
         end;
         lObEch1.PutValue( 'ECH_CODEEMPRUNT', fOmEmprunt.GetValue('EMP_CODEEMPRUNT'));
      end;
      lObEmprunACalculer.Free;


      // Recalcul des totaux
      fOmEmprunt.PutValue('EMP_TOTINT', 0);
      fOmEmprunt.PutValue('EMP_TOTASS', 0);
      fOmEmprunt.PutValue('EMP_TOTAMORT', 0);
      PutValue('EMP_TOTVERS', 0);
      lRdCumulA  := 0;
      lRdCumulI  := 0;
      lRdCumulV  := 0;
      for lInCpt1 := 0 to fOmEmprunt.Echeances.Detail.Count - 1 do
      begin
         lObEch1  := fOmEmprunt.Echeances.Detail[lInCpt1];
         lRdAs  := lObEch1.GetValue('ECH_ASSURANCE');
         lRdAm  := lObEch1.GetValue('ECH_AMORTISSEMENT');
         lRdIn  := lObEch1.GetValue('ECH_INTERET');
         lRdVe  := lRdIn + lRdAs + lRdAm;

         fOmEmprunt.PutValue('EMP_TOTINT',   fOmEmprunt.GetValue('EMP_TOTINT')   + lRdIn );
         fOmEmprunt.PutValue('EMP_TOTASS',   fOmEmprunt.GetValue('EMP_TOTASS')   + lRdAs );
         fOmEmprunt.PutValue('EMP_TOTAMORT', fOmEmprunt.GetValue('EMP_TOTAMORT') + lRdAm );
         fOmEmprunt.PutValue('EMP_TOTVERS',  fOmEmprunt.GetValue('EMP_TOTVERS')  + lRdVe );

         lRdCumulA := lRdCumulA + lRdAm;
         lRdCumulI := lRdCumulI + lRdIn + lRdAs;
         lRdCumulV := lRdCumulV + lRdVe;

         lObEch1.PutValue('ECH_CUMULA', lRdCumulA );
         lObEch1.PutValue('ECH_CUMULI', lRdCumulI );
         lObEch1.PutValue('ECH_CUMULV', lRdCumulV );
      end;
   end;

   // Init des champs totaux dans l'écran de l'emprunt
   Modifier;
   SetField('EMP_TOTINT',   fOmEmprunt.GetValue('EMP_TOTINT'));
   SetField('EMP_TOTASS',   fOmEmprunt.GetValue('EMP_TOTASS'));
   SetField('EMP_TOTAMORT', fOmEmprunt.GetValue('EMP_TOTAMORT'));
   SetField('EMP_TOTVERS',  fOmEmprunt.GetValue('EMP_TOTVERS'));
   SetField('EMP_AMORTCST', fOmEmprunt.GetValue('EMP_AMORTCST'));
   SetField('EMP_VERSCST',  fOmEmprunt.GetValue('EMP_VERSCST'));
   SetField('EMP_NBECHTOT', fOmEmprunt.GetValue('EMP_NBECHTOT'));

   // Rafraichissement du tableau d'amortissement
   EcheancesVersGrille;

   // Focalise la grille
   if pBoWithPrompt then
      FocaliserGrille;

   // Tableau d'amortissement au statut calculé ou saisi
   if grEcheances.Cells[cColModif,1] = 'X' then
      ChangeStatutEmprunt(seSaisi)
   else
      ChangeStatutEmprunt(seCalcule);

   //
   Pages.OnChange := OnPageControlChange;
end;

// ----------------------------------------------------------------------------
// Nom    : FocaliserGrille
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  : donne le focus à la grille des échéances
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.FocaliserGrille;
begin
   if Pages.ActivePageIndex <> cInEchPage then
   begin
      Pages.ActivePageIndex := cInEchPage;
      if Assigned(Pages.OnChange) then
         Pages.OnChange(Self);
   end;

   SetFocusControl('grEcheances');
end;


// ----------------------------------------------------------------------------
// Nom    : ChangeStatutEmprunt
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  : Change le statut des du tableau d'amortissement (calculé, saisi, ...)
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TOM_FEMPRUNT.ChangeStatutEmprunt(pInStatut : TStatutEmprunt);
var
   lStCaption : String;
begin
   PutValue('EMP_STATUT', Ord(pInStatut));
   pInStatut := GetField('EMP_STATUT');

   // Libellé en fonction du statut des échéances
   case pInStatut of
      seVide    : lStCaption := '';
      seCalcule : lStCaption := 'Tableau d''amortissement calculé';
      seModifie : lStCaption := 'Tableau d''amortissement non à jour';
      seSaisi   : lStCaption := 'Tableau d''amortissement ajusté manuellement';
      seSimule  : lStCaption := 'Tableau d''amortissement simulé';
   end;
   THPanel(Getcontrol('PNSTATUT')).Caption := lStCaption;

   // Passe en mode modification lorsque le tableau d'amortissement est saisi
   // manuellement
   if pInStatut = seSaisi then
      Modifier;

   // Fait clignoter le bouton 'calcul théorique' si les échéances ne sont plus
   // à jour par rapport à ce qui a été saisi dans l'entête
   Clignoter(pInStatut = seModifie)
end;


// ----------------------------------------------------------------------------
// Nom    : btAjusterClick
// Date   : 06/11/2002
// Auteur : D. ZEDIAR
// Objet  : Ajustement proportionnel
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.btAjusterClick(Sender: TObject);
var
   lRdMontant : Double;
begin
   if fOmEmprunt.Echeances.Detail.Count = 0 then exit;

   // Calcul du montant de base de l'ajustement (intérets + assurance)
   lRdMontant := GetField('EMP_TOTINT');
   if GetField('EMP_EMPTYPEASSUR') = taTaux then
      lRdMontant := lRdMontant + GetField('EMP_TOTASS');

   // Saisie du montant de l'ajustement
   lRdMontant := SaisieAjustement(lRdMontant);

   if lRdMontant <= 0 then Exit;

   if not Confirmer(True) then Exit;

   // Récupère les données de l'écran si nécéssaire
   EcranVersEmprunt;

   // Ajustement
   if not fOmEmprunt.Ajustement(lRdMontant) then
      Exit;

   // Rafraichissement de l'écran
   EcheancesVersGrille;

   ChangeStatutEmprunt(seSaisi);

   FocaliserGrille;
end;

// ----------------------------------------------------------------------------
// Nom    : btSimulerClick
// Date   : 18/12/2001
// Auteur : D. ZEDIAR
// Objet  : Simulation
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.btSimulerClick(Sender: TObject);
var
   lInTypeRecherche : TTypeRecherche;
   lRdMontant       : Double;
begin
   // Confirmation
   if not Confirmer(True) then Exit;

   //
   EcranVersEmprunt;

   // Saisie du versement ou de l'amortissement constant
   if GetField('EMP_EMPTYPEVERS') = tvVariable then
      lRdMontant := SaisieMontantConstant(True,fOmEmprunt.CalcAmortCst)
   else
      lRdMontant := SaisieMontantConstant(False,fOmEmprunt.CalcVersCst);

   if lRdMontant <= 0 then Exit;

   if GetField('EMP_EMPTYPEVERS') = tvVariable then
      PutValue('EMP_AMORTCST',lRdMontant)
   else
      PutValue('EMP_VERSCST',lRdMontant);

   // Demande le type de la simulation
   lInTypeRecherche := ChoixTypeSim(GetField('EMP_EMPTYPETAUX'), GetField('EMP_EMPTYPEVERS'));

   if lInTypeRecherche = trAucune then Exit;

   // Lancement de la simulation
   fOmEmprunt.Simulation(lInTypeRecherche);

   //
   EmpruntVersEcran;

   ChangeStatutEmprunt(seSimule);

   // Affichage du résultat
   case lInTypeRecherche of
      trCapital   : PGIInfo('Capital calculé : ' + fOmEmprunt.FormaterMontant(GetField('EMP_TOTAMORT') - fOmEmprunt.TotalReport) + ' ' + gSymbDevEmprunt, TitreHalley);
      trDuree     : PGIInfo('Duree calculée : ' + Inttostr(GetField('EMP_NBECHTOT')) + ' mois', TitreHalley);
      trTaux      : PGIInfo('Taux annuel calculé : ' + fOmEmprunt.FormaterMontant(GetField('EMP_TAUXAN')) + ' %', TitreHalley);
      trVersement : if not ((lInTypeRecherche = trVersement) and (GetField('EMP_EMPTYPEVERS') = tvVariable)) then
                       PGIInfo('Versement calculé : ' + fOmEmprunt.FormaterMontant(fOmEmprunt.VersementMaximum) + ' ' + gSymbDevEmprunt, TitreHalley);
   end;

   //
   FocaliserGrille;

end;


// ----------------------------------------------------------------------------
// Nom    : btDeviseOrigineClick
// Date   : 07/11/2002
// Auteur : D. ZEDIAR
// Objet  : Affichage de l'emprunt en devise d'origine
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.btDeviseOrigineClick(Sender: TObject);
begin
   EcranVersEmprunt;
   EmpruntEnMonaieOrigine(fOmEmprunt);
end;



// ----------------------------------------------------------------------------
// Nom    : EcranVersEmprunt
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  : Ecran vers OM emprunt
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TOM_FEMPRUNT.EcranVersEmprunt;
begin
   fOmEmprunt.GetEcran(Ecran,pnEmprunt);
   // Récupère le contenu de la grille si l'utilisateur a saisi le tableau
   // d'amortissement
   if GetField('EMP_STATUT') = Ord(seSaisi) then
      GrilleVersEcheances
end;


// ----------------------------------------------------------------------------
// Nom    : EmpruntVersEcran
// Date   : 04/11/2002
// Auteur : D. ZEDIAR
// Objet  : OM emprunt vers écran
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TOM_FEMPRUNT.EmpruntVersEcran;
var
   lInOldStatus : Integer;
   lInCpt : Integer;
begin
   lInOldStatus := fOmEmprunt.GetValue('EMP_STATUT');

   Modifier;
   for lInCpt := 1 to fOmEmprunt.NbChamps do
      SetField(fOmEmprunt.GetNomChamp(lInCpt), fOmEmprunt.GetValue(fOmEmprunt.GetNomChamp(lInCpt)));
   DeviseEmprunt := fOmEmprunt.GetValue('EMP_DEVISE');
   DeviseSaisie  := fOmEmprunt.GetValue('EMP_DEVISESAISIE');
   EcheancesVersGrille;

   ChangeStatutEmprunt(TStatutEmprunt(lInOldStatus));
end;


// ----------------------------------------------------------------------------
// Nom    : GrilleVersEcheances
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  : Grille vers échéances
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TOM_FEMPRUNT.GrilleVersEcheances;
var
  lInRow    : Integer;
  lInIndexe : Integer;
  lObEch    : Tob;
begin

  lInIndexe := 0;
  For lInRow := 1 to grEcheances.RowCount - 1 do
  begin
     // Ignore les lignes de totalisation de fin d'exercice
     if grEcheances.Cells[cColTypeLigne, lInRow] = cStLigneTotal then
        Continue;

     lObEch := fOmEmprunt.Echeances.Detail[lInIndexe];

     // Récupère uniquement les colonnes saisissables (toutes désormais)
//     fOmEmprunt.ECH_DATEEXERCICE.Value  := fOmEmprunt.ProchaineDateExercice(fOmEmprunt.ECH_DATE.AsDate);
     lObEch.PutValue('ECH_TAUXINT',        Valeur(grEcheances.Cells[cColTxInteret,lInRow]) );
     lObEch.PutValue('ECH_TAUXASS',        Valeur(grEcheances.Cells[cColTxAssurance,lInRow]) );
     lObEch.PutValue('ECH_INTERET',        Valeur(grEcheances.Cells[cColInteret,lInRow]) );
     lObEch.PutValue('ECH_ASSURANCE',      Valeur(grEcheances.Cells[cColAssurance,lInRow]) );
     lObEch.PutValue('ECH_VERSEMENT',      Valeur(grEcheances.Cells[cColVersement,lInRow] ));
     lObEch.PutValue('ECH_AMORTISSEMENT',  Valeur(grEcheances.Cells[cColAmortissement,lInRow] ));
     lObEch.PutValue('ECH_SOLDE',          Valeur(grEcheances.Cells[cColSolde,lInRow] ));
     lObEch.PutValue('ECH_MODIF',          grEcheances.Cells[cColModif,lInRow] );

     lInIndexe := lInIndexe + 1;
  end;

  // récupération des totaux
  PutValue('EMP_TOTINT',   EMP_TOTINT.Value);
  PutValue('EMP_TOTASS',   EMP_TOTASS.Value);
  PutValue('EMP_TOTAMORT', EMP_TOTAMORT.Value);
  PutValue('EMP_TOTVERS',  EMP_TOTVERS.Value);

end;

// ----------------------------------------------------------------------------
// Nom    : EcheancesVersGrille
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  : Echéances vers grille
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TOM_FEMPRUNT.EcheancesVersGrille;
begin
   grEcheances.OnSelectCell := Nil;

   fOmEmprunt.EcheancesVersGrille(grEcheances);

   //
   grEcheances.RowHeights[0] := pnStatut.Height;

   // Totaux de l'emprunt
   EMP_TOTINT.Value    := fOmEmprunt.GetValue('EMP_TOTINT');
   EMP_TOTASS.Value    := fOmEmprunt.GetValue('EMP_TOTASS');
   EMP_TOTAMORT.Value  := fOmEmprunt.GetValue('EMP_TOTAMORT');
   EMP_TOTVERS.Value   := fOmEmprunt.GetValue('EMP_TOTVERS');
   EMP_TOTINTASS.Value := fOmEmprunt.GetValue('EMP_TOTINT') + fOmEmprunt.GetValue('EMP_TOTASS');

   // Se place sur la première cellule saisissable
   grEcheances.Col := cColTxInteret;
   grEcheances.Row := 1;
   fInOldCol := grEcheances.Col;
   fInOldRow := grEcheances.Row;
   fStOldCellValue := grEcheances.Cells[grEcheances.Col,grEcheances.Row];

   //
   grEcheances.OnSelectCell := grEcheancesSelectCell;
end;


// ----------------------------------------------------------------------------
// Nom    : OnCtrlKeyDown
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  : CtrlKeyDown
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.OnCtrlKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
   lStFieldName : String;
begin
   fBoSaisirMontant := False;
   lStFieldName := UpperCase(TCustomEdit(Sender).Name);

   // gestion de la saisie des montants en devise d'origine
   if (lStFieldName = 'EMP_CAPITAL') or
      ((lStFieldName = 'EMP_VALASSUR') and (GetField('EMP_EMPTYPEASSUR') = taMontant)) then
   begin
      // Suppression su montant si Suppr + readonly
{$IFDEF EAGLCLIENT}
      if (Key = VK_DELETE) and THEdit(Sender).ReadOnly then
{$ELSE}
      if (Key = VK_DELETE) and THDBEdit(Sender).ReadOnly then
{$ENDIF}
      begin
         DeviseSaisie := DeviseEmprunt;
      end
      else
         // saisie du montant si readonly ou ctrl + E
         if ( DeviseEmprunt <> DeviseSaisie) and
{$IFDEF EAGLCLIENT}
        ( (THEdit(Sender).ReadOnly) or
{$ELSE}
        ( (THDBEdit(Sender).ReadOnly) or
{$ENDIF}
               ((Shift = [ssCtrl]) and (Key in [Ord('e'),Ord('E')])) ) then
            fBoSaisirMontant := True;
   end;
end;


// ----------------------------------------------------------------------------
// Nom    : OnMontantKeyPress
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.OnMontantKeyPress(Sender: TObject; var Key: Char);
begin
  {FQ20157  30.05.07  YMO}
  If (THEdit(Sender).name='EMP_VALASSUR') and (Key ='.') then FMyOnKeyPress(Sender,Key);
  {FQ21141  23.07.07  YMO}
  If (THEdit(Sender).name='EMP_CAPITAL') and (Key ='.') then FMkOnKeyPress(Sender,Key);

  if fBoSaisirMontant then
{$IFNDEF EAGLCLIENT}
     SaisirMontant(THDbEdit(Sender).DataField)
{$ELSE}
     SaisirMontant(THEDIT(Sender).name)
{$ENDIF}
  else
     DeviseSaisie := DeviseEmprunt;

end;


// ----------------------------------------------------------------------------
// Nom    : Clignoter
// Date   : 18/12/2001
// Auteur : D. ZEDIAR
// Objet  : Fair clignoter un control
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TOM_FEMPRUNT.Clignoter(pBoStatus : Boolean);
begin
   if V_PGI.ModeTSE then Exit;

   if pBoStatus then
   begin
      fBoClignoOn := False;
      fObClignoTimer.Interval := 500;
   end;

   fObClignoTimer.Enabled := pBoStatus;

   if not pBoStatus then
   begin
      btCalculer.Font.Style := [];
      btCalculer.Font.Color := clBlack;
   end;
end;



// Timer du clignotement
procedure TOM_FEMPRUNT.ClignoTimer(Sender: TObject);
begin
   fBoClignoOn := not fBoClignoOn;

   if fBoClignoOn then
      fObClignoTimer.Interval := 300
   else
      fObClignoTimer.Interval := 500;

   if fBoClignoOn then
   begin
      btCalculer.Font.Style := [fsBold];
      btCalculer.Font.Color := clBlack;
   end
   else
   begin
      btCalculer.Font.Style := [fsBold];
      btCalculer.Font.Color := clDkGray;
   end;

end;


// ----------------------------------------------------------------------------
// Nom    : SaisirMontant
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  : sisie des montants en devise d'origine
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.SaisirMontant(pStFieldName : String);
var
   lRdMontant : Double;
begin
   lRdMontant := SaisieMontant(GetField(pStFieldName));
   if lRdMontant > 0 then
   begin
      if DeviseSaisie = '' then
      begin
         if DeviseEmprunt = gCodeDevEuro then
            DeviseSaisie := gCodeDevFranc
         else
            DeviseSaisie := gCodeDevEuro;
      end;
      PutValue(pStFieldName, lRdMontant);
   end;
end;

// ----------------------------------------------------------------------------
// Nom    : GetDeviseEmprunt / SetDeviseEmprunt
// Date   : 18/12/2001
// Auteur : D. ZEDIAR
// Objet  : Propriété DeviseEmprunt
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Function  TOM_FEMPRUNT.GetDeviseEmprunt : String;
begin
   Result := GetField('EMP_DEVISE');
end;


Procedure TOM_FEMPRUNT.SetDeviseEmprunt(pValue : String);
begin
   // Met à jour les variables globales
   PutValue('EMP_DEVISE', pValue);
   gCodeDevEmprunt := pValue;
   if pValue = '' then
   begin
      gLibDevEmprunt  := gSymbDevNull;
      gSymbDevEmprunt := '';
   end
   else
   begin
      InfoDevise(pValue,gLibDevEmprunt,gSymbDevEmprunt);
      gLibDevEmprunt := UpperCase(Copy(gLibDevEmprunt,1,1)) + Copy(gLibDevEmprunt,2,1000)
   end;

   // Met à jour le symbole monaitaire des ctrls
   SetControlCaption('EMP_DEVISE', gLibDevEmprunt);
   SetControlProperty('EMP_CAPITAL',   'CurrencySymbol', gSymbDevEmprunt);
   SetControlProperty('EMP_VALASSUR',  'CurrencySymbol', gSymbDevEmprunt);
   SetControlProperty('EMP_TOTINT',    'CurrencySymbol', gSymbDevEmprunt);
   SetControlProperty('EMP_TOTASS',    'CurrencySymbol', gSymbDevEmprunt);
   SetControlProperty('EMP_TOTAMORT',  'CurrencySymbol', gSymbDevEmprunt);
   SetControlProperty('EMP_TOTINTASS', 'CurrencySymbol', gSymbDevEmprunt);
   SetControlProperty('EMP_TOTVERS',   'CurrencySymbol', gSymbDevEmprunt);
end;


// ----------------------------------------------------------------------------
// Nom    : GetDeviseSaisie / SetDeviseSaisie
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  : Propriété DeviseSaisie
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Function  TOM_FEMPRUNT.GetDeviseSaisie : String;
begin
   Result := GetField('EMP_DEVISESAISIE');
end;

Procedure TOM_FEMPRUNT.SetDeviseSaisie(pValue : String);
begin
   // Met à jour les variables
   PutValue('EMP_DEVISESAISIE', pValue);
   gCodeDevSaisie := pValue;
   if pValue = '' then
   begin
      gLibDevSaisie  := gSymbDevNull;
      gSymbDevSaisie := '';
   end
   else
   begin
      InfoDevise(pValue,gLibDevSaisie,gSymbDevSaisie);
      gLibDevSaisie := UpperCase(Copy(gLibDevSaisie,1,1)) + Copy(gLibDevSaisie,2,1000)
   end;

   // Mise à jour des contrôles
   SetControlCaption('EMP_DEVISESAISIE', gLibDevSaisie);

   // Maj de l'apparence des ctrls concernés par la devise de saisie (montants)
   AppliquerMonnaieSaisie;

end;


// ----------------------------------------------------------------------------
// Nom    : AppliquerMonnaieSaisie
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  : Maj de l'apparence des ctrls concernés par la devise de saisie
//         (montants)
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.AppliquerMonnaieSaisie;
begin
   if (DeviseSaisie <> DeviseEmprunt) and
      (DeviseSaisie <> '') then
   begin
      SetControlProperty('EMP_CAPITAL','ReadOnly',True);
      SetControlProperty('EMP_VALASSUR','ReadOnly',GetField('EMP_EMPTYPEASSUR') = taMontant);
      SetControlEnabled('btDeviseOrigine',True);
   end
   else
   begin
      SetControlProperty('EMP_CAPITAL','ReadOnly',False);
      SetControlProperty('EMP_VALASSUR','ReadOnly',False);
      SetControlEnabled('btDeviseOrigine',False);
   end;
end;


// ----------------------------------------------------------------------------
// Nom    : InfoDevise
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  : Renvoie des informations sur une devise
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Function TOM_FEMPRUNT.InfoDevise(pStCodeDevise : String; Var pStLibelle, pStSymbole : String) : String;
var
   lObTob : TOB;
   lStSQL : String;
begin
   // Chargement de la devises
   lStSQL := 'Select D_LIBELLE, D_SYMBOLE From DEVISE Where D_DEVISE = "' + pStCodeDevise + '"';
   lObTob := Tob.Create('',Nil,-1);
   try
      lObTob.LoadDetailFromSQL(lStSQL);
      if lObTob.Detail.Count = 0 then
      begin
         pStLibelle := gSymbDevNull;
         pStSymbole := '';
      end
      else
      begin
         pStLibelle := lObTob.Detail[0].GetValue('D_LIBELLE');
         pStSymbole := lObTob.Detail[0].GetValue('D_SYMBOLE');
      end;
   finally
      lObTob.Free;
   end;
end;



// ----------------------------------------------------------------------------
// Nom    : Confirmer
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  : Demande de confirmation avant écrasement du tableau d'ammortissement
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Function TOM_FEMPRUNT.Confirmer(pBoWithPrompt : Boolean) : Boolean;
var
   lStMessage       : String;
   lInStatutEmprunt : TStatutEmprunt;
begin
   Result := True;
   lInStatutEmprunt := GetField('EMP_STATUT');

   if (lInStatutEmprunt = seSaisi) or
      (lInStatutEmprunt = seSimule) then
   begin
      if pBoWithPrompt then
      begin
         if lInStatutEmprunt = seSaisi then
            lStMessage := 'Le tableau d''amortissement a été ajusté manuellement.' + chr(13)
         else
            lStMessage := 'Le tableau d''amortissement est issu d''une simulation.' + chr(13);
         if PGIAsk(lStMessage + 'Etes-vous certain de vouloir écraser les modifications ?',TitreHalley) = mrNo then
            Result := False;
      end
      else
         Result := False;
   end;
end;


// ----------------------------------------------------------------------------
// Nom    : SaisieValide
// Date   : 30/10/2002
// Auteur : D. ZEDIAR
// Objet  : tèste la validité de la saisie de l'emprunt
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Function TOM_FEMPRUNT.SaisieValide(pBoWithMessage : Boolean) : Boolean;
begin
   Result := True;

   // Capital > 0
   if GetField('EMP_CAPITAL') <= 0 then
   begin
      if pBoWithMessage then
      begin
         PGIInfo('La saisie du capital emprunté est obligatoire',TitreHalley);
         SetFocusControl('EMP_CAPITAL');
      end;
      result := False;
      Exit;
   end;

   // Taux annuel > 0
   if GetField('EMP_TAUXAN') < 0 then
   begin
      if pBoWithMessage then
      begin
         PGIInfo('La saisie du taux annuel est obligatoire',TitreHalley);
         SetFocusControl('EMP_TAUXAN');
      end;
      result := False;
      Exit;
   end;
end;



// ----------------------------------------------------------------------------
// Nom    : OnDureeExit
// Date   : 04/11/2002
// Auteur : D. ZEDIAR
// Objet  : Traitement de la durée
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.OnDureeExit(Sender: TObject);
var
//   lInNbMoisEch : Integer;
   lInDuree     : Integer;
   lStPeriode   : String;
   lInMinValue  : Integer;
   lInMaxValue  : Integer;
begin
     // Valaurs min et max
     lStPeriode := GetField('EMP_EMPPERIODE');
     lInMinValue := 1;
     if lStPeriode = tpTrimestre then
        lInMinValue := 3
     else
        if lStPeriode = tpSemestre then
           lInMinValue := 4
        else
           if lStPeriode = tpAnnee then
              lInMinValue := 12;

     lInMaxValue := cInMaxEch * trunc(lInMinValue);
     lInDuree := GetField('EMP_DUREE');

     if lInDuree < lInMinValue then
        PutValue('EMP_DUREE', lInMinValue);

     if lInDuree > lInMaxValue then
        PutValue('EMP_DUREE', lInMaxValue);

     // Arrondi le nombre d'échéances selon la période
     lInDuree := GetField('EMP_DUREE');
     if lInDuree mod lInMinValue > 0 then
     begin
        lInDuree := (Trunc(lInDuree / lInMinValue) + 1) * lInMinValue;
        PutValue('EMP_DUREE',lInDuree);
     end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 08/06/2007
Modifié le ... :   /  /
Description .. : Comptes complétés
Mots clefs ... : FQ20541
*****************************************************************}
procedure TOM_FEMPRUNT.OnCompteExit(Sender: TObject);
var
  LeCompte : String;
begin
  LeCompte:=THEdit(Sender).Text;
  If Trim(LeCompte)<> '' then THEdit(Sender).Text:=BourreEtLess( LeCompte, fbGene ) ;
 {
  try
    lQuery := OpenSQL('SELECT * FROM GENERAUX WHERE G_GENERAL = "' + LeCompte + '"', True);
    if lQuery.Eof then
    begin
      {$IFDEF EAGLCLIENT
      LookupList(THEdit(Sender),TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE','G_GENERAL LIKE "16%"','G_GENERAL', True, -1);
      {$ELSE
      LookupList(THDBEdit(Sender),TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE','G_GENERAL LIKE "16%"','G_GENERAL', True, -1);
      {$ENDIF
    end;

  finally
    Ferme( lQuery );
  end;
  }
end;

procedure TOM_FEMPRUNT.OnDebutEnter(Sender: TObject);
begin
   fStOldDateDebut := GetControlText('EMP_DATEDEBUT');
end;

{FQ20540  31.05.07 YMO
procedure TOM_FEMPRUNT.OnDebutExit(Sender: TObject);
begin
   if fStOldDateDebut <> GetControlText('EMP_DATEDEBUT') then
      try
         PutValue('EMP_DATECONTRAT', StrToDate(GetControlText('EMP_DATEDEBUT')));
      except
      end;
end;
}

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/03/2007
Modifié le ... :   /  /
Description .. : Pour gérer le mode création, le -OnComboChange-
Description .. : doit laisser le focus sur le libellé
Mots clefs ... :
*****************************************************************}
procedure TOM_FEMPRUNT.OnLibelleExit(Sender: TObject);
begin
  fEnCreation:=False;
end;

// ----------------------------------------------------------------------------
// Nom    : OnComboChange
// Date   : 04/11/2002
// Auteur : D. ZEDIAR
// Objet  : Permet une réaction "immédiate" au changement de valeurs de
//          certaines combo (pour faire apparaitre ou disparaitre certains
//          controles)
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.OnComboChange(Sender : TObject);
begin
   if fEnCreation then
     SetFocusControl('EMP_LIBEMPRUNT')
   else
     SetFocusControl(THValComboBox(Sender).Name);
end;


// ----------------------------------------------------------------------------
// Nom    : btTauxVariablesClick
// Date   : 04/11/2002
// Auteur : D. ZEDIAR
// Objet  : Saisie des taux variables
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TOM_FEMPRUNT.btTauxVariablesClick(Sender: TObject);
begin
   if SaisieTauxVariables(fOmEmprunt) then
      if Integer(GetField('EMP_STATUT')) in [Ord(seCalcule),Ord(seVide)] then
         ChangeStatutEmprunt(seModifie)
      else
         Modifier;
end;

// ----------------------------------------------------------------------------
// Nom    : btOrganismeClick
// Date   : 10/12/2001
// Auteur : D. ZEDIAR
// Objet  : lancement de la fenêtre de l'annuaire
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TOM_FEMPRUNT.btOrganismeClick(Sender: TObject);
var
   lStOrgPreteur : String;
   lStOrgSelectionne : String;
begin
   lStOrgPreteur := GetField('EMP_GUIDORGPRETEUR');
  // lStOrgSelectionne := AGLLanceFiche('YY','ANNUAIRE_SEL','XX_WHERE=ANN_TYPEPER##"BNQ"','','');
  // FQ 19523 - CA 26/03/2007
   lStOrgSelectionne := LancerAnnuSel('ANN_TYPEPER=BNQ','','');

{   if lStOrgPreteur = '-1' then
      AGLLanceFiche('YY','ANNUAIRE','ANN_TYPEPER=="BNQ"','','')
   else
      AGLLanceFiche('YY','ANNUAIRE_SEL','',lStOrgPreteur,'');}

   // Recharge la combo des organismes prêteurs
   if RemplirComboOrganismes then
   begin
      if lStOrgSelectionne = '' then
      //   PutValue('EMP_ORGPRETEUR', ValeurI(lStOrgPreteur))
         PutValue('EMP_GUIDORGPRETEUR', lStOrgPreteur)
      else
       //  PutValue('EMP_ORGPRETEUR', ValeurI(lStOrgSelectionne));
         PutValue('EMP_GUIDORGPRETEUR', lStOrgSelectionne);
   end
     else
      PutValue('EMP_GUIDORGPRETEUR', ''); {CA 26/04/2007}
end;


// ----------------------------------------------------------------------------
// Nom    : OnPageControlChange
// Date   : 18/12/2001
// Auteur : D. ZEDIAR
// Objet  : Changement de page
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.OnPageControlChange(Sender: TObject);
begin
   // Lancement automatique du calcul théorique si passage dans la page 2
   if (Pages.ActivePageIndex = cInEchPage) then
      if Integer(GetField('EMP_STATUT')) in [Ord(seVide), Ord(seModifie)] then
         Calculer(True);

   // Etat des boutons en fonction de l'onglet sélectionné
   Case Pages.ActivePageIndex of
      cInEmpPage :
          SetFocusControl('EMP_DATEDEBUT');
      cInEchPage :
          SetFocusControl('grEcheances.SetFocus');
   end;

end;



procedure TOM_FEMPRUNT.OnPageControlChanging(Sender: TObject; var AllowChange: Boolean);
begin
   // Vérif de la validité de la saisie avant de passer au tableau d'amortissement
   AllowChange := true;
   if (Pages.ActivePageIndex = 0) then
   begin
      SetFocusControl('EMP_LIBEMPRUNT');
      if Integer(GetField('EMP_STATUT')) in [Ord(seVide), Ord(seModifie)] then
         AllowChange := SaisieValide(True);
   end;
end;


// ----------------------------------------------------------------------------
// Nom    : grEcheancesGetCellCanvas
// Date   : 18/12/2001
// Auteur : D. ZEDIAR
// Objet  : Style de la cellule
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.grEcheancesGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin

   // Cellules fixes en gras (titres)
   if (ACol = 0) or (ARow = 0) then
   begin
      Canvas.Font.Name  := 'Arial';
      Canvas.Font.Size  := 10;
      Canvas.Font.Style := [fsBold];
   end;

   if (ARow > 0) and (ACol <> 0) then
      // Lignes modifiées en rouge
      if grEcheances.Cells[cColModif,ARow] = 'X' then
         Canvas.Font.Color := cInModifColor
      else
         if grEcheances.Cells[cColTypeLigne,ARow] = cStLigneTotal then
         begin
            // Totaux d'exercice en bleu
            Canvas.Brush.Color := cInFondExColor;
            Canvas.Font.Style  := [fsBold];
         end;

end;



// ----------------------------------------------------------------------------
// Nom    : FormResize
// Date   : 04/11/2002
// Auteur : D. ZEDIAR
// Objet  : Centrage du block de saisie de l'emprunt
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.FormResize(Sender: TObject);
var
   lIn : Integer;
begin
   lIn := Pages.Width;
   if pnEmpruntWidth + 10 > lIn then
      pnEmprunt.Left := 5
   else
      pnEmprunt.Left := trunc((lIn - pnEmpruntWidth) / 2);

   lIn := Pages.Height;
   if pnEmpruntHeight + 10 > lIn - 30 then
      pnEmprunt.Top  := 5
   else
      pnEmprunt.Top  := trunc((lIn - pnEmpruntHeight - 30) / 2);
end;


// ----------------------------------------------------------------------------
// Nom    : grEcheancesSelectCell
// Date   : 18/12/2001
// Auteur : D. ZEDIAR
// Objet  : gestion du changement de cellule et des lignes et colonnes en
//          lecture seul
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOM_FEMPRUNT.grEcheancesSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
begin
   // Modif d'un montant -> répercutions sur le rèste du tableau d'amortissement
   if (fInOldRow <> -1) and (fInOldCol <> -1) then
      if fStOldCellValue <> grEcheances.Cells[fInOldCol,fInOldRow] then
      begin
         // Formate la valeur saisie
         if grEcheances.Col in [cColTxAssurance, cColTxAssurance] then
            grEcheances.Cells[fInOldCol,fInOldRow] := fOmEmprunt.Formatertaux(Valeur(grEcheances.Cells[fInOldCol,fInOldRow]))
         else
            grEcheances.Cells[fInOldCol,fInOldRow] := fOmEmprunt.Formatermontant(Valeur(grEcheances.Cells[fInOldCol,fInOldRow]));

         ChangeStatutEmprunt(seSaisi);

         // Flag de modification de la ligne
         if grEcheances.Cells[cColModif,fInOldRow] = '-' then
         begin
            grEcheances.Cells[cColModif,fInOldRow] := 'X';
            grEcheances.Invalidate;
//            Modifier;
         end;

         // Recalcul des totaux
         RepercuterModifTableau(fInOldCol,fInOldRow,fStOldCellValue);

      end;

   //
   fInOldRow       := aRow;
   fInOldCol       := aCol;
   fStOldCellValue := grEcheances.Cells[fInOldCol,fInOldRow];
end;

// ----------------------------------------------------------------------------
// Nom    : RepercuterModifTableau
// Date   : 04/11/2002
// Auteur : D. ZEDIAR
// Objet  : Modification en cascade du tableau d'amortissement suite à une
//          saisie de l'utilisateur dans la grille
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TOM_FEMPRUNT.RepercuterModifTableau(pInCol : Integer; pInRow : Integer; pStOldValue : String);
var
   lRdDelta  : Double;
   lStValeur : String;
begin
   // Delta de la modif apportée par l'utilisateur
   lRdDelta := Valeur(grEcheances.Cells[pInCol, pInRow]) - Valeur(pStOldValue);

   case pInCol of

      cColInteret, cColassurance :
         begin
            // Modif de l'intéret -> modif du total int + ass
            lStValeur := grEcheances.Cells[cColTotInteret,pInRow];
            grEcheances.Cells[cColTotInteret,pInRow] := fOmEmprunt.FormaterMontant(Valeur(lStValeur) + lRdDelta);
            RepercuterModifTableau(cColTotInteret,pInRow,lStValeur);
            ModifierTotal(pInCol,pInRow,lRdDelta);
         end;

      cColTxInteret, cColTxAssurance :
         begin
            lStValeur := grEcheances.Cells[cColTxGlobal,pInRow];
            grEcheances.Cells[cColTxGlobal,pInRow] := fOmEmprunt.FormaterMontant(Valeur(lStValeur) + lRdDelta);
         end;

      cColTotInteret, cColAmortissement :
         begin
            // Modif du total de l'intéret ou de l'amortieement -> modif du versement
            lStValeur := grEcheances.Cells[cColVersement,pInRow];
            {FQ19522 25/04/07 YMO Répercussion sur le versement}
            grEcheances.Cells[cColVersement,pInRow] := fOmEmprunt.FormaterMontant(Valeur(lStValeur) + lRdDelta);
            RepercuterModifTableau(cColVersement,pInRow,lStValeur);
            ModifierTotal(pInCol,pInRow,lRdDelta);
         end;

      cColVersement :
         begin
            ModifierTotal(pInCol,pInRow,lRdDelta);
            Calculer(True); //YMOO
         end;
   end;
end;


// ----------------------------------------------------------------------------
// Nom    : ModifierTotal
// Date   : 04/11/2002
// Auteur : D. ZEDIAR
// Objet  : Modif d'un total
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TOM_FEMPRUNT.ModifierTotal(pInCol : Integer; pInRow : Integer; pRdDelta : Double);
var
   lInCpt   : Integer;
   lInRowEx : Integer;
begin

   // Modif des totaux généraux
   Case pInCol of
      cColInteret :
      begin
         EMP_TOTINT.Value    := EMP_TOTINT.Value + pRdDelta;
         EMP_TOTINTASS.Value := EMP_TOTINTASS.Value + pRdDelta;
      end;
      cColassurance :
      begin
         EMP_TOTASS.Value    := EMP_TOTASS.Value + pRdDelta;
         EMP_TOTINTASS.Value := EMP_TOTINTASS.Value + pRdDelta;
      end;
      cColAmortissement :
         EMP_TOTAMORT.Value := EMP_TOTAMORT.Value + pRdDelta;
      cColVersement :
         EMP_TOTVERS.Value := EMP_TOTVERS.Value + pRdDelta;
   end;

   // Modif des totaux d'exercices dans la grille
   // recherche de la ligne de total
   lInRowEx := 0;
   for lInCpt := pInRow to grEcheances.RowCount - 1 do
      if grEcheances.Cells[cColTypeLigne,lInCpt] = cStLigneTotal then
      begin
         lInRowEx := lInCpt;
         Break;
      end;
   // Ligne de total trouvée -> modif du total
   if lInRowEx > 0 then
      grEcheances.Cells[pInCol,lInRowEx] := fOmEmprunt.FormaterMontant(Valeur(grEcheances.Cells[pInCol,lInRowEx]) + pRdDelta);

end;


// ----------------------------------------------------------------------------
// Nom    : Imprimer
// Date   : 18/12/2001
// Auteur : D. ZEDIAR
// Objet  : Impression de l'emprunt
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TOM_FEMPRUNT.Imprimer(pBoListeExport : Boolean = False);
var
   lInCodeEmprunt : Integer;
   lOmEmpDev      : TObjetEmprunt;
   lRdTaux        : Double;
   lStChoix       : String;
begin
   if not EstSerialise then Exit;

   // Enregistrement de l'emprunt si nécéssaire
   if DS.State <> dsBrowse then
   begin
//      PGIBox('Vous devez enregistrer l''emprunt avant de l''imprimer.',TitreHalley);
//      Exit;
      TToolbarButton97(Getcontrol('bValider')).Click;
      if DS.State <> dsBrowse then
         Exit;
   end;

   // Tableau vide ?
   if fOmEmprunt.Echeances.Detail.Count = 0 then
   begin
      PGIBox('Le tableau d''amortissement est vide.',TitreHalley);
      Exit;
   end;

   //
   lInCodeEmprunt := GetField('EMP_CODEEMPRUNT');
   lOmEmpDev      := Nil;
   lStChoix       := '';

   // Devise de l'impression ?
   // Si la devise de saisie (devise d'origine) est différente de la devise de l'emprunt,
   // L'utilisateur peut choisir entre l'une ou l'autre
   if (DeviseSaisie <> DeviseEmprunt) and (DeviseSaisie <> '') then
   begin
      // Choix de la devise
      lStChoix := ChoixDeviseImpression(gSymbDevEmprunt,gSymbDevSaisie);
      if lStChoix = '' then
         Exit;

      if lStChoix = 'ORIGINE' then
      begin
         // Impression en devise d'origine :
         // On va créer un emprunt temporaire dans la devise d'origine

         // Détermine le code de l'emprunt en devise d'origine
{         if GetField('EMP_CODEEMPRUNT') <= 0 then
         begin
            fOmEmprunt.ValeurPK := fOmEmprunt.EmpruntSuivant;
            PutValue('EMP_CODEEMPRUNT', fOmEmprunt.GetValue('EMP_CODEEMPRUNT'));
         end;      }
         lInCodeEmprunt := - GetField('EMP_CODEEMPRUNT');

         // Création et calcul de l'emprunt en devise d'origine
         lOmEmpDev := TObjetEmprunt.Creeremprunt;
         EcranVersEmprunt;
         lOmEmpDev.Assigner(fOmEmprunt);
         lOmEmpDev.ValeurPK := lInCodeEmprunt;
         lOmEmpDev.PutValue('EMP_DEVISE', DeviseSaisie);
         lOmEmpDev.PutValue('EMP_DEVISESAISIE', '@');

         // Conversion des montants dans la devise d'origine
         if DeviseSaisie = gCodeDevEuro then
            lRdTaux := 1/V_PGI.TauxEuro
         else
            lRdTaux := V_PGI.TauxEuro;
         lOmEmpDev.PutValue('EMP_CAPITAL', lOmEmpDev.GetValue('EMP_CAPITAL') * lRdTaux);
         if lOmEmpDev.GetValue('EMP_EMPTYPEASSUR') = taMontant then
            lOmEmpDev.PutValue('EMP_VALASSUR', lOmEmpDev.GetValue('EMP_VALASSUR') * lRdTaux);

         // simulation et enregistrement
         if lOmEmpDev.GetValue('EMP_EMPTYPEVERS') = tvVariable then
            lOmEmpDev.PutValue('EMP_AMORTCST', lOmEmpDev.CalcAmortCst);

         lOmEmpDev.Simulation(trVersement);
         lOmEmpDev.PutValue('EMP_CODEEMPDEV', GetField('EMP_CODEEMPRUNT') );

         // Enregistrement de l'emprunt en devise d'origine
         if transactions(lOmEmpDev.DBSupprimeEtRemplaceEmprunt,1) <> oeOk then
         begin
            lOmEmpDev.Free;
            Exit;
         end;
      end;
   end;

   // Impression
   ImprimeEmprunt(lInCodeEmprunt,'',pBoListeExport); {YMO 18/04/07 Ajout libellé}

   // suppression de l'emprunt en devise d'origine si nécéssaire
   if (lStChoix = 'ORIGINE') and (Assigned(lOmEmpDev)) then
   begin
      transactions(lOmEmpDev.DBSupprimerEmprunt,1);
      lOmEmpDev.Free;
   end;

end;

// ----------------------------------------------------------------------------
// Nom    : Modifier
// Date   : 18/12/2001
// Auteur : D. ZEDIAR
// Objet  : Passe en mode modif ou création
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TOM_FEMPRUNT.Modifier;
begin
   if not (DS.State in [dsEdit, dsInsert]) then
      ForceUpdate;
end;



Initialization
  registerclasses ( [ TOM_FEMPRUNT ] ) ;
end.

