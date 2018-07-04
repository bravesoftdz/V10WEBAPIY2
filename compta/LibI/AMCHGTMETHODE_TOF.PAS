
{***********UNITE*************************************************
Auteur  ...... : PGR
Créé le ...... : 09/09/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AMCHGTMETHODE ()
Mots clefs ... : TOF;AMCHGTMETHODE
*****************************************************************}
// PGR -16/12/2005 FQ17204 Calcul ancien montant des amortissements
// FQ 17215 - TGA 21/12/2005 - GetParamSoc => GetParamSocSecur
// FQ 17282 - PGR 09/01/2006 - initialisation pointeur pour appel éclatement
// FQ 17298 - PGR 12/01/2006 - Ctl taux éco < taux fiscal
// FQ 17326 - PGR 19/01/2006 - Calcul quantité par rapport à valeur à éclater et inversement
// PGR 20/01/2006 - Remplacement Imeclate par Imeclate2
// FQ 17569 - MBO 16/05/2006 - Modification pour gestion date début d'amortissement
// FQ 18119 - BTY - 05/06    - Positionner les indicateurs de modif de la compta dans PARAMSOC
// FQ 18394 - BTY - 06/06 Passer le n° d'ordre en série attendu par l'éclatement
// FQ 18876 - MVG - 09/06 Récupération du PARAMSOC SO_IMMOANOUVEAU
// BTY 10/06 - Bloquer l'opération si l'immo possède une prime d'équipement
// MBO 11/06 - Bloquer l'opération si l'immo possède une subvention d'investissement
Unit AMCHGTMETHODE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     maineagl,
{$ENDIF}
{$IFDEF VER150}
   Variants,
{$ENDIF}
     ParamSoc,
    // PGR 20/01/2006 - Remplacement Imeclate par Imeclate2
    // Imeclate,
     ImEnt,
     ImPlan,
     forms,
     Outils,
     sysutils,
     ComCtrls,
     Dialogs,
     HCtrls,
     HEnt1,
     HTB97,
     HMsgBox,
     HRichOle,
     UTOM,
     UTOB,
     UTOF,
     ImOuPlan,
     Math,
     Vierge;

Type
  TOF_AMCHGTMETHODE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure OnAfterFormShowEcran;

    private
    { Variables }
    fmethode: String;
    fmontantht: double;
    fblocnote: THRichEditOLE;
    fImmoDejaDeprecie : Boolean;
    ffiscal : String;
    fquantite : integer;
    famortanteco: double;
    famortantfisc: double;
    ftauxfisc: double;
    EcoEgalService: Boolean;
    FiscEgalService: Boolean;
    MethodeFisc: string;
    procedure OnClickChEclatement(Sender: TObject);
    procedure OnClickChCondition(Sender: TObject);
    procedure OnChangeValeurAEclater(Sender: TObject);
    // FQ 17326 - PGR 19/01/2006 - Calcul quantité par rapport à valeur à éclater et inversement
    procedure OnExitValeurAEclater(Sender: TObject);
    procedure OnExitQuantite(Sender: TObject);
    procedure OnExitDuree(Sender: TObject);
    procedure DureeDejaAmortie;
    procedure OnbFermeClick(Sender: TObject);
  end ;

{***********UNITE*************************************************
Auteur  ...... : PGR
Créé le ...... : 09/09/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AMCHGTMETHODE ()
Mots clefs ... : TOF;AMCHGTMETHODE
*****************************************************************}
function AMLanceFiche_FicheChgtMethode(Stcode: string): Integer ;

Implementation
// PGR 20/01/2006 - Remplacement Imeclate par Imeclate2
uses
 Imeclate2;

const
 HM2: array[0..17] of string =
 {1} ('%d mois - %d jour(s)',
 {2}  'Vous devez renseigner la date d''opération dans l''exercice.',
 {3}  'La date d''opération doit être supérieure ou égale à la plus haute date d''opération déjà enregistrée.',
 {4}  'Vous devez choisir au moins un changement de méthode.',
 {5}  'Le montant Valeur à éclater doit être inférieur au Hors Taxe.',
 {6}  'Le code immobilisation existe déjà.',
 {7}  'Vous devez choisir une méthode d''amortissement.',
 {8}  'Le montant Valorisation doit être supérieur ou égal au montant Valeur à éclater',
 {9}  'La date ''A compter du'' doit être supérieure ou égale à la date d''achat.',
 {10} 'Voulez-vous enregistrer cette opération ?',
 {11} 'L''enregistrement du changement des conditions d''amortissement a échoué.',
 {12} 'L''immobilisation a déjà été dépréciée, il n''est plus possible de modifier rétrospectivement le plan d''amortissement économique.',
 {13} '%d mois',
 {14} 'Vous devez renseigner une quantité inférieure à la quantité d''origine et non nulle.',
 {15}'Veuillez définir le compte de report à nouveau dans les fourchettes de comptes.',
 {16}'Le taux d''amortissement économique ne peut être supérieur au taux fiscal.',
     'Les antérieurs économiques ne doivent pas être supérieurs aux antérieurs fiscaux.',
 {17}'Changement de méthode non autorisé,' +
      ' l''immobilisation a été financée au moyen d''une prime d''équipement ou d''une subvention d''investissement.'
     );


function AMLanceFiche_FicheChgtMethode(Stcode: string): Integer;
var
  PlanAmort : TPlanAmort;
  QPlan : TQuery;
  Prime : double;
  SBV : double;
begin
  // FQ 17282 - PGR 09/01/2006 - initialisation pour éviter les warnings
  result := 0;
  
  QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+Stcode+'"', FALSE) ;
  PlanAmort:=TPlanAmort.Create(true) ;

  try
    PlanAmort.Charge(QPlan);
    PlanAmort.Recupere(Stcode,QPlan.FindField('I_PLANACTIF').AsString);
  finally
    PlanAmort.Free;
    Prime := QPlan.FindField('I_SBVPRI').AsFloat;
    SBV := QPlan.FindField('I_SBVMT').AsFloat;
    Ferme(Qplan);
  end;

  // Appel Changement de méthode si présence du cpte de report à nouveau
  // MVG FQ 18876 GetParamSocSecur('SO_IMMOANOUVEAU','')

  if GetParamSocSecur('SO_IMMOANOUVEAU',false) = '' then
    pgibox(HM2[14])
  else if  Prime <> 0 then     // BTY 10/06 Prime d'équipement
    pgibox(HM2[17])
  else if SBV <> 0 then   // mbo 03.11.06 pour subvention
    pgibox(HM2[17])
  else
    result:=StrToInt(AGLLanceFiche('AM', 'AMCHGTMETHODE', '', '', Stcode));
end;

procedure TOF_AMCHGTMETHODE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AMCHGTMETHODE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AMCHGTMETHODE.OnDisplay ;
begin
  Inherited ;
end ;

{***********UNITE*************************************************
Auteur  ...... : PGR
Créé le ...... : 09/09/2005
Modifié le ... :   /  /
Description .. : OnUpdate
Mots clefs ... : TOF;AMCHGTMETHODE
*****************************************************************}
procedure TOF_AMCHGTMETHODE.OnUpdate ;
var ordre:integer;
    date_saisie:TdateTime;
    date_achat:TdateTime;
    montant:Double;
    aeclater:Double;
    valorisation:Double;
    T1: TOB;
    taux: double;
    QPlan : TQuery;
    PlanAmort : TPlanAmort;
    PlanNew : TPlanAmort;
   // PGR 20/01/2006 - Remplacement Imeclate par Imeclate2
   // Eclate : TFEclatement;
    PremMois, PremAnnee, NbMois: Word;
    PlanActifAv: integer;
    PlanActifAp: integer;
    PlanAmortAntEco : double;
    PlanAmortAntFisc : double;
    immoATraiter : String;
    DotationExe : Double;
    ExcepExe : Double;
    DateOpe : TDateTime;
    methode: String;
    dureeeco : integer;
    miseenservice : TdateTime;
    repriseeco : Double;
    reprisefisc : Double;
    dureereprise : integer;
    DotSansOper : Double;
    amortantecoeclat : double;
    amortantfisceclat : double;
    AntEco : double;
    AntFiscal: double;
    OldEco: TDateTime;
    OldFis: TDateTime;
    OldAntEco: double;
    OldAntFisc:double;
    tauxEco : double;
    FFOrdreSerie : integer;  // FQ 18394
begin
  Inherited ;
  ecran.ModalResult := mrYes;
  // FQ 17282 - PGR 09/01/2006 - initialisation pour éviter les warnings
  taux:= 0;
  aeclater:= 0;
  valorisation:= 0;
  PlanAmortAntFisc := 0;

  if (GetControlText('CH_CONDITION') <> 'X') and (GetControlText('CH_ECLATEMENT') <> 'X') then
  begin
    pgibox(HM2[3]);
    ecran.ModalResult := mrNone;
    //LastError = -1;
    //LastErrorMsg := '';
    SetFocusControl('CH_ECLATEMENT');
    exit;
  end;

  //Contrôle date opération dans l'exercice et >= à la dernière opération

  if valeur(GetControlText('CH_DATEOP'))= 0 then
  begin
    pgibox(HM2[1]);
    ecran.ModalResult := mrNone;
    SetFocusControl('CH_DATEOP');
    exit;
  end
  else
  begin
    date_saisie := StrToDate(GetControlText('CH_DATEOP'));
    // la date doit appartenir à l'exercice en cours
    if (date_saisie >  VHImmo^.Encours.Fin) OR ( date_saisie < VHImmo^.Encours.Deb) then
    begin
      pgibox(HM2[1]);
      ecran.ModalResult := mrNone;
      SetFocusControl('CH_DATEOP');
      exit;
    end
    //la date d'opération doit être >= date dernière opération
    else if ExisteSQL('SELECT IL_DATEOP FROM IMMOLOG WHERE IL_IMMO="'+GetControlText('CHI_CODEIMMO')+'" AND IL_DATEOP>"'+USDATETIME(date_saisie)+'"') then
    begin
      pgibox(HM2[2]);
      ecran.ModalResult := mrNone;
      SetFocusControl('CH_DATEOP');
      exit;
    end;

  end;

  //Contrôle bloc éclatement préalable

  if GetControlText('CH_ECLATEMENT') = 'X' then
  begin
    //Contrôle Valeur à éclater < montant HT

    montant := Valeur(GetControlText('CH_VALEURECLAT'));

    If (montant >= fmontantht) or (montant=0) then
    begin
      pgibox(HM2[4]);
      ecran.ModalResult := mrNone;
      SetFocusControl('CH_VALEURECLAT');
      exit;
    end;

    //Contrôle Code immobilisation inexistant

    if ExisteSQL('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+GetControlText('CH_CODEIMMO')+'"') then
    begin
      pgibox(HM2[5]);
      ecran.ModalResult := mrNone;
      SetFocusControl('CH_CODEIMMO');
      exit;
    end;

    //Contrôle Valorisation >= valeur à éclater

    aeclater := Valeur(GetControlText('CH_VALEURECLAT'));
    valorisation := Valeur(GetControlText('CH_VALORISATION'));

    if (valorisation < aeclater) then
    begin
      pgibox(HM2[7]);
      ecran.ModalResult := mrNone;
      SetFocusControl('CH_VALORISATION');
      exit;
    end;

    //Contrôle Quantité < quantité origine, si > 1

    if ((fquantite > 1) AND (Valeur(GetControlText('CH_QUANTITE')) >= fquantite)) OR
      (Valeur(GetControlText('CH_QUANTITE')) < 1) then
    begin
      pgibox(HM2[13]);
      ecran.ModalResult := mrNone;
      SetFocusControl('CH_QUANTITE');
      exit;
    end;

  end;

  //Contrôle bloc Changement des conditions d'amortissement

  if GetControlText('CH_CONDITION') = 'X' then
  begin

     //Changement des conditions d'amortissement impossible si immo déjà dépréciée

    if (fImmoDejaDeprecie = TRUE) then
    begin
      pgibox(HM2[11]);
      ecran.ModalResult := mrNone;
      SetFocusControl('CH_CONDITION');
      exit;
    end;

    //Contrôle A compter du >= date d'achat

    if valeur(GetControlText('CH_ACOMPTER'))= 0 then
    begin
      pgibox(HM2[8]);
      ecran.ModalResult := mrNone;
      SetFocusControl('CH_ACOMPTER');
      exit;
    end
    else
    begin
      date_saisie := StrToDate(GetControlText('CH_ACOMPTER'));
      date_achat := StrToDate(GetControlText('CHI_DATEACHAT'));

      if (date_saisie) < (date_achat) then
      begin
        pgibox(HM2[8]);
        ecran.ModalResult := mrNone;
        SetFocusControl('CH_ACOMPTER');
        exit;
      end;
    end;

    //Contrôle Méthode

    if GetControlText('CH_METHODE') = '' then
    begin
      pgibox(HM2[6]);
      ecran.ModalResult := mrNone;
      SetFocusControl('CH_METHODE');
      exit;
    end;

    //Calcul taux
    taux := 0;

    if ((GetControlText('CH_METHODE') <> '')
      and (GetControlText('CH_METHODE') <> 'NAM')
      and (valeur(GetControlText('CH_ACOMPTER'))<> 0)
      and (not VarIsNull(GetControlText('CH_DUREE')))) then
    begin
      taux := GetTaux(GetControlText('CH_METHODE')
         , StrToDate(GetControlText('CH_ACOMPTER'))
         , StrToDate(GetControlText('CH_ACOMPTER'))
         , StrToInt(GetControlText('CH_DUREE')));
    end;

    SetControlText('CH_TAUX',FormatFloat('#,##0.00',taux));

    //FQ 17298 - PGR 12/01/2006 - Ctl taux éco < taux fiscal
    //Contrôle taux < taux fiscal si existant
    {if (ffiscal <> '') and
     (taux >= ftauxfisc) then
    begin
      pgibox(HM2[15]);
      ecran.ModalResult := mrNone;
      SetFocusControl('CH_DUREE');
      exit;
    end;}

    // mbo fq 17569
  if methodeFisc <> '' then
  begin
    QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+GetControlText('CHI_CODEIMMO')+'"', FALSE) ;
    Oldeco:= QPlan.FindField('I_DATEDEBECO').AsDateTime;
    Oldfis:= QPlan.FindField('I_DATEDEBFIS').AsDateTime;
    OldAntEco := QPlan.FindField('I_REPRISEECO').AsFloat;
    OldAntFisc := QPlan.FindField('I_REPRISEFISCAL').AsFloat;
    methode := QPlan.FindField('I_METHODEECO').AsString;
    dureeeco := QPlan.FindField('I_DUREEECO').AsInteger;
    dureereprise := QPlan.FindField('I_DUREEREPRISE').AsInteger;
    tauxEco := QPlan.FindField('I_TAUXECO').AsFloat;

    QPlan.Edit;
    PlanNew:=TPlanAmort.Create(true) ;
    try
      QPlan.FindField('I_DATEDEBECO').AsDateTime := StrToDate(GetControlText('CH_ACOMPTER'));
      QPlan.FindField('I_METHODEECO').AsString := GetControlText('CH_METHODE');

       if GetControlText('CH_METHODE') = 'NAM' then
       begin
         QPlan.FindField('I_DUREEECO').AsInteger := 0 ;
         QPlan.FindField('I_TAUXECO').AsFloat := 0;
         QPlan.FindField('I_DUREEREPRISE').AsString := '';
       end
       else
       begin
         QPlan.FindField('I_DUREEECO').AsInteger := StrToInt(GetControlText('CH_DUREE')) ;
         QPlan.FindField('I_TAUXECO').AsFloat := taux;
         NOMBREMOIS(StrToDate(GetControlText('CH_ACOMPTER')),VHImmo^.Encours.Deb,PremMois, PremAnnee,  NbMois);
         QPlan.FindField('I_DUREEREPRISE').AsString := IntToStr(MaxIntValue([0,NbMois - 1]));
       end;

      if (FiscEgalService) and (MethodeFisc <> 'DEG') then
         QPlan.FindField('I_DATEDEBFIS').AsDateTime := StrToDate(GetControlText('CH_ACOMPTER'));
      QPlan.Post;

      PlanNew.Charge(QPlan);
      PlanNew.calculReprises;

      Qplan.Edit;
      QPlan.FindField('I_REPRISEECO').AsFloat := PlanNew.AmortEco.Reprise;
      QPlan.FindField('I_REPRISEFISCAL').AsFloat := PlanNew.AmortFisc.Reprise;
      Qplan.Post;

      PlanNew.Recalcul(Qplan);
      PlanNew.GetCumulsDotExercice(VHImmo^.Encours.Deb,AntEco, AntFiscal, False, true, True);

      Qplan.Edit;
      QPlan.FindField('I_DATEDEBECO').AsDateTime := OldEco;
      QPlan.FindField('I_DATEDEBFIS').AsDateTime := oldFis;
      QPlan.FindField('I_REPRISEECO').AsFloat:= oldAntEco;
      QPlan.FindField('I_REPRISEFISCAL').AsFloat:= OldAntFisc;
      QPlan.FindField('I_TAUXECO').AsFloat:= TauxEco;
      QPlan.FindField('I_METHODEECO').AsString:= methode;
      QPlan.FindField('I_DUREEECO').AsInteger:= dureeeco;
      QPlan.FindField('I_DUREEREPRISE').AsInteger:= dureereprise;

      QPlan.Post;

    finally
      Plannew.free;
      Ferme (QPlan);
    end;
    if Anteco > AntFiscal then
      begin
         pgibox(HM2[16]);
         ecran.ModalResult := mrNone;
         SetFocusControl('CH_ACOMPTER');
         exit;
      end;
   end;
  end;

  // Mise à jour des tables liée à éclatement préalable

  if GetControlText('CH_ECLATEMENT') = 'X' then
  begin
    // FQ 18394 Passer le n° d'ordre en série
    FFOrdreSerie :=  TrouveNumeroOrdreSerieLogSuivant;
    //FQ 17282 - PGR 09/01/2006 - initialisation pointeur pour appel éclatement
    // PGR 20/01/2006 - Remplacement Imeclate par Imeclate2
    //Eclate:=TFEclatement.Create(Application) ;
    //Eclate.New_enregistreEclatement(GetControlText('CHI_CODEIMMO'),
    enregistreEclatement_2(GetControlText('CHI_CODEIMMO'),
                                     fmontantht,
                                     valeur(GetControlText('CHI_QTE')),
                                     GetControlText('CH_CODEIMMO'),
                                     aeclater,
                                     valorisation,
                                     GetControlText('CH_DESIGNATION'),
                                     valeur(GetControlText('CH_QUANTITE')),
                                     GetControlText('CH_DATEOP'),
                                     false,
                                     '',
                                     FFOrdreSerie);  // FQ 18394
    //Eclate.Free;

    // PGR -16/12/2005 FQ17204 Calcul ancien montant des amortissements
    //Récupération antérieur mère après éclatement
    QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+GetControlText('CHI_CODEIMMO')+'"', FALSE) ;

    if QPlan.Eof then
    begin
      Ferme(Qplan);
      PGIBox(HM2[10]);
      TFVIERGE(ECRAN).retour := IntToStr(MrNo);
      exit
    end;

    PlanAmort:=TPlanAmort.Create(true) ;
    try
      PlanAmort.Charge(QPlan);
      PlanAmort.Recupere(GetControlText('CHI_CODEIMMO'),QPlan.FindField('I_PLANACTIF').AsString);
      PlanAmort.GetCumulsDotExercice(VHImmo^.Encours.Deb,amortantecoeclat, amortantfisceclat, False, False, True);
      amortantecoeclat := amortantecoeclat + QPlan.FindField('I_REPRISEECO').AsFloat;
      amortantfisceclat := amortantfisceclat + QPlan.FindField('I_REPRISEFISCAL').AsFloat;
    finally
      PlanAmort.Free;
    end;

  end;

  // Mise à jour des tables liée à Changement de conditions d'amortissement

  if GetControlText('CH_CONDITION') = 'X' then
  begin

   // Mise à jour des tables IMMO,IMMOAMOR et IMMOLOG
   BEGINTRANS ;
   try
    //Récupération immo en cours ou immo fille si éclatement préalable
    if GetControlText('CH_ECLATEMENT') = 'X' then
      immoATraiter := GetControlText('CH_CODEIMMO')
    else immoATraiter := GetControlText('CHI_CODEIMMO');

    // Mise à jour de IMMO
    QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+immoATraiter+'"', FALSE) ;

    if QPlan.Eof then
    begin
      Ferme(Qplan);
      PGIBox(HM2[10]);
      TFVIERGE(ECRAN).retour := IntToStr(MrNo);
      exit
    end;

    //Sauvegarde anciennes données

    methode := QPlan.FindField('I_METHODEECO').AsString;
    dureeeco := QPlan.FindField('I_DUREEECO').AsInteger;
    dureereprise := QPlan.FindField('I_DUREEREPRISE').AsInteger;
    miseenservice := QPlan.FindField('I_DATEAMORT').AsDateTime;
    // Fq 17569 mbo ajout des 2 dates début d'amortissement
    OldEco := QPlan.FindField('I_DATEDEBECO').AsDateTime;
    OldFis := QPlan.FindField('I_DATEDEBFIS').AsDateTime;
    repriseeco := QPlan.FindField('I_REPRISEECO').AsFloat;
    reprisefisc := QPlan.FindField('I_REPRISEFISCAL').AsFloat;
    PlanAmort:=TPlanAmort.Create(true) ;
    try
      PlanAmort.Charge(QPlan);
      PlanAmort.Recupere(immoATraiter,QPlan.FindField('I_PLANACTIF').AsString);
      DotationExe := PlanAmort.GetDotationExercice(VHImmo^.EnCours.Fin,PlanAmort.AmortEco,false);
      ExcepExe := PlanAmort.GetExcepExercice(VHImmo^.Encours.Fin, DateOpe, true, true);
      DotSansOper := DotationExe - ExcepExe;
      methodefisc := PlanAmort.AmortFisc.methode;

   finally
      PlanAmort.Free;
    end;

    QPlan.Edit;

    QPlan.FindField('I_OPERATION').AsString:= 'X' ;
    QPlan.FindField('I_OPECHANGEPLAN').AsString:= 'X' ;

    // fq 17569 mbo on force la date mise en service avec nvelle uniquement si elle était
    // égale à la date début d'amort éco
    QPlan.FindField('I_DATEDEBECO').AsDateTime := StrToDate(GetControlText('CH_ACOMPTER'));
    if EcoEgalService then
       QPlan.FindField('I_DATEAMORT').AsDateTime := StrToDate(GetControlText('CH_ACOMPTER'));
    if (FiscEgalService) and (MethodeFisc <> 'DEG') then
       QPlan.FindField('I_DATEDEBFIS').AsDateTime := StrToDate(GetControlText('CH_ACOMPTER'));

    QPlan.FindField('I_METHODEECO').AsString := GetControlText('CH_METHODE');

    if GetControlText('CH_METHODE') = 'NAM' then
    begin
      QPlan.FindField('I_DUREEECO').AsInteger := 0 ;
      QPlan.FindField('I_TAUXECO').AsFloat := 0;
      QPlan.FindField('I_DUREEREPRISE').AsString := '';
    end
    else
    begin
      QPlan.FindField('I_DUREEECO').AsInteger := StrToInt(GetControlText('CH_DUREE')) ;
      QPlan.FindField('I_TAUXECO').AsFloat := taux;
      NOMBREMOIS(StrToDate(GetControlText('CH_ACOMPTER')),VHImmo^.Encours.Deb,PremMois, PremAnnee,  NbMois);
      QPlan.FindField('I_DUREEREPRISE').AsString := IntToStr(MaxIntValue([0,NbMois - 1]));
    end;

    QPlan.Post;

    // Maj du plan d'amortissement IMMOAMOR
    QPlan.Edit;
    PlanAmort:=TPlanAmort.Create(true) ;
    try
      PlanAmort.Charge(QPlan);
      PlanAmort.CalculReprises;
      PlanAmortAntEco := PlanAmort.AmortEco.Reprise;
      QPlan.FindField('I_REPRISEECO').AsFloat := PlanAmort.AmortEco.Reprise;
      QPlan.FindField('I_DATEDERMVTECO').AsDateTime := PlanAmort.GetDateFinAmortEx(PlanAmort.AmortEco);

      if ffiscal <> '' then
      begin
        PlanAmortAntFisc := PlanAmort.AmortFisc.Reprise;
        QPlan.FindField('I_REPRISEFISCAL').AsFloat := PlanAmort.AmortFisc.Reprise;
        QPlan.FindField('I_DATEDERNMVTFISC').AsDateTime := PlanAmort.GetDateFinAmortEx(PlanAmort.AmortFisc);
      end;

      PlanAmort.Recupere(immoATraiter,QPlan.FindField('I_PLANACTIF').AsString);
      PlanActifAv := PlanAmort.NumSeq;
      PlanAmort.SetTypeOpe('CDA');
      PlanAmort.Recalcul(Qplan);
      PlanAmort.Sauve;

      // Attribuer à l'immo le n° du nouveau plan d'amortissement
      Qplan.FindField('I_PLANACTIF').AsInteger := PlanAmort.NumSeq;
      Qplan.Post;

      PlanActifAp := PlanAmort.NumSeq;
    finally
      PlanAmort.Free;
      Ferme(Qplan);
    end;

    // Mise à jour de IMMOLOG

    Ordre := TrouveNumeroOrdreLogSuivant(immoATraiter);

    T1:=TOB.Create('IMMOLOG',nil,-1) ;
    try
      T1.PutValue('IL_IMMO',immoATraiter) ;
      T1.PutValue('IL_LIBELLE','Anciennes conditions d''amortissement') ;
      T1.PutValue('IL_TYPEOP','CDA') ;
      T1.PutValue('IL_DATEOP',StrToDate(GetControlText('CH_DATEOP')));
      T1.PutValue('IL_ORDRE',ordre);
      T1.PutValue('IL_ORDRESERIE',-1);
      T1.PutValue('IL_PLANACTIFAV',PlanActifAv);
      T1.PutValue('IL_PLANACTIFAP',PlanActifAp);
      T1.PutValue('IL_METHODEECO',methode);
      T1.PutValue('IL_DUREEECO',dureeeco);
      T1.PutValue('IL_DATEOPREELLE',miseenservice);
      T1.PutValue('IL_DUREE',dureereprise);
      // PGR -16/12/2005 FQ17204 Calcul ancien montant des amortissements
      // Antérieur = Ant mère de départ - Ant mère après éclatement
      T1.PutValue('IL_REPRISEECO',famortanteco-amortantecoeclat);
      T1.PutValue('IL_CUMANTCESECO',repriseeco);
      T1.PutValue('IL_REPRISEFISC',PlanAmortAntEco);
      T1.PutValue('IL_MONTANTDOT',DotSansOper);
      // ajout mbo fq 17569
      T1.PutValue('IL_CODEMUTATION', DateTostr(OldEco));
      T1.PutValue('IL_CODEECLAT', DateToStr(OldFis));

      if ffiscal <> '' then
      begin
        T1.PutValue('IL_REVISIONREPFISC',PlanAmortAntFisc);
        // PGR -16/12/2005 FQ17204 Calcul ancien montant des amortissements
        T1.PutValue('IL_REVISIONREPECO',famortantfisc-amortantfisceclat);
        T1.PutValue('IL_CUMANTCESFIS',reprisefisc);
      end;

      T1.PutValue('IL_BLOCNOTE', RichToString (fblocnote));
      T1.PutValue('IL_TYPEMODIF',AffecteCommentaireOperation('DPA'));
      T1.InsertDB(nil) ;
    finally
      T1.Free ;
    end;

    COMMITTRANS ;
   except
    PGIBox(HM2[10]);
    ROLLBACK ;
   end ;
  end;

  // 05/06 FQ 18119 Positionner les indicateurs de modif
  VHImmo^.ChargeOBImmo := True;
  ImMarquerPublifi (True);

  TFVIERGE(ECRAN).retour := IntToStr(MrYes);
end;

procedure TOF_AMCHGTMETHODE.OnLoad ;
begin
  Inherited ;
end ;

{***********UNITE*************************************************
Auteur  ...... : PGR
Créé le ...... : 09/09/2005
Modifié le ... :   /  /
Description .. : OnArgument
Mots clefs ... : TOF;AMCHGTMETHODE
*****************************************************************}
procedure TOF_AMCHGTMETHODE.OnArgument (S : String ) ;
var
    PlanAmort : TPlanAmort;
    QPlan : TQuery;

begin
  Inherited ;
  { Evènements }
  TCheckBox(GetControl('CH_ECLATEMENT')).OnClick := OnClickchEclatement;
  TCheckBox(GetControl('CH_CONDITION')).OnClick := OnClickchCondition;
  THEdit(GetControl('CH_VALEURECLAT')).OnChange := OnChangeValeurAEclater;
  // FQ 17326 - PGR 19/01/2006 - Calcul quantité par rapport à valeur à éclater et inversement
  THEdit(GetControl('CH_VALEURECLAT')).OnExit := OnExitValeurAEclater;
  THEdit(GetControl('CH_QUANTITE')).OnExit := OnExitQuantite;
  TCheckBox(GetControl('CH_DUREE')).OnExit := OnExitDuree;
  TToolbarButton97(GetControl('BFerme')).Onclick := OnbFermeClick;

  TFVIERGE(ECRAN).retour := IntToStr(MrNo);
  TFVierge(Ecran).OnAfterFormShow := OnAfterFormShowEcran;

  // Assignation du bloc note
  fblocnote := THRichEditOLE(GetControl('CH_BLOCNOTE'));

  // Lecture de l'immo en cours

  QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+S+'"', FALSE) ;

  SetControlCaption('CHI_CODEIMMO',QPlan.FindField('I_IMMO').AsString);
  SetControlCaption('CHI_DESIGNATION',QPlan.FindField('I_LIBELLE').AsString);
  SetControlCaption('CHI_DATEACHAT',QPlan.FindField('I_DATEPIECEA').AsString);
  SetControlCaption('CHI_BASEAMORT',StrFMontant(QPlan.FindField('I_BASEECO').AsFloat,13,V_PGI.OkDecV,'',True));
  SetControlCaption('CHI_CPTE',QPlan.FindField('I_COMPTEIMMO').AsString);
  SetControlCaption('CHI_QTE',IntToStr(QPlan.FindField('I_QUANTITE').AsInteger));
  SetControlCaption('CHI_METHODE',RechDom('TIMETHODEIMMO',QPlan.FindField('I_METHODEECO').AsString,False));
  SetControlCaption('CHI_TAUX',FormatFloat('#,##0.00',QPlan.FindField('I_TAUXECO').AsFloat));
  SetControlCaption('CHI_DUREE', Format(HM2[12], [QPlan.FindField('I_DUREEECO').AsInteger]));
  SetControlCaption('CHI_SERVICE',QPlan.FindField('I_DATEAMORT').AsString);
  SetControlText('CHI_DUREEAMORTIE', '');
  // ajout mbo fq 17569
  SetControlCaption('CHI_DEBECO',QPlan.FindField('I_DATEDEBECO').AsString);
  If (QPlan.FindField('I_DATEDEBECO').AsDateTime = QPlan.FindField('I_DATEAMORT').AsDateTime) then
     EcoEgalService := true
  else
     EcoEgalService := false;

  If (QPlan.FindField('I_DATEDEBFIS').AsDateTime = QPlan.FindField('I_DATEAMORT').AsDateTime) then
     FiscEgalService := true
  else
     FiscEgalService := false;
  
  MethodeFisc := QPlan.FindField('I_METHODEFISC').AsString;

  SetControlEnabled('CH_VALEURECLAT',False);
  SetControlEnabled('CH_CODEIMMO',False);
  SetControlEnabled('CH_VALORISATION',False);
  SetControlEnabled('CH_QUANTITE',False);
  SetControlEnabled('CH_DESIGNATION',False);
  SetControlEnabled('CH_METHODE',False);
  SetControlEnabled('CH_DUREE',False);
  SetControlEnabled('CH_TAUX',False);
  SetControlEnabled('CH_ACOMPTER',False);

  fmethode := QPlan.FindField('I_METHODEECO').AsString;
  fmontantht := QPlan.FindField('I_MONTANTHT').AsFloat;
  fquantite := QPlan.FindField('I_QUANTITE').AsInteger;
  ffiscal  := QPlan.FindField('I_METHODEFISC').AsString;
  // FQ 17298 - PGR 12/01/2006 - Ctl taux éco < taux fiscal
  ftauxfisc := QPlan.FindField('I_TAUXFISC').AsFloat;


  THValComboBox(GetControl('CH_METHODE')).Items.Delete(THValComboBox(GetControl('CH_METHODE')).Items.IndexOf('Variable'));
  if ffiscal <> '' then
    THValComboBox(GetControl('CH_METHODE')).Items.Delete(THValComboBox(GetControl('CH_METHODE')).Items.IndexOf('Non amortissable'));

  //calcul durée déjà amortie
  DureeDejaAmortie;

  PlanAmort:=TPlanAmort.Create(true) ;

  try
    PlanAmort.Charge(QPlan);
    PlanAmort.Recupere(GetControlText('CHI_CODEIMMO'),QPlan.FindField('I_PLANACTIF').AsString);

    //Immo dépréciée sur N-1 et N ?
    if (PlanAmort.GestDeprec = TRUE) OR (PlanAmort.DureeRest = TRUE) then
      fImmoDejaDeprecie := TRUE;

    // PGR -16/12/2005 FQ17204 Calcul ancien montant des amortissements
    PlanAmort.GetCumulsDotExercice(VHImmo^.Encours.Deb,famortanteco, famortantfisc, False, False, True);
    famortanteco := famortanteco + QPlan.FindField('I_REPRISEECO').AsFloat;
    famortantfisc := famortantfisc + QPlan.FindField('I_REPRISEFISCAL').AsFloat;
  finally
    PlanAmort.Free;
    Ferme(Qplan);
  end;

  SetFocusControl('CH_DATEOP');
end ;

procedure TOF_AMCHGTMETHODE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AMCHGTMETHODE.OnCancel () ;
begin
  Inherited ;
end ;

{***********UNITE*************************************************
Auteur  ...... : PGR
Créé le ...... : 09/09/2005
Modifié le ... :   /  /
Description .. : MAJ Valorisation en même temps que Valeur à éclater
Mots clefs ... : TOF;AMCHGTMETHODE
*****************************************************************}
procedure TOF_AMCHGTMETHODE.OnChangeValeurAEclater(Sender: TObject);

begin
  SetControlText('CH_VALORISATION', GetControlText('CH_VALEURECLAT'));
end;


{***********UNITE*************************************************
Auteur  ...... : PGR
Créé le ...... : 09/09/2005
Modifié le ... :   /  /
Description .. : Click sur éclatement préalable
Mots clefs ... : TOF;AMCHGTMETHODE
*****************************************************************}
procedure TOF_AMCHGTMETHODE.OnClickChEclatement(Sender: TObject);

begin
  If GetControlText('CH_ECLATEMENT') = 'X' then
  begin
    SetControlEnabled('CH_VALEURECLAT',True);
    SetControlEnabled('CH_CODEIMMO',True);
    SetControlEnabled('CH_VALORISATION',True);
    if fquantite > 1 then
      SetControlEnabled('CH_QUANTITE',True)
    else SetControlText('CH_QUANTITE', '1');
    SetControlEnabled('CH_DESIGNATION',True);
    SetControlText('CH_DESIGNATION', GetControlText('CHI_DESIGNATION'));
    SetControlText('CH_CODEIMMO', NouveauCodeImmo);
  end
  else
  begin
    SetControlEnabled('CH_VALEURECLAT',False);
    SetControlEnabled('CH_CODEIMMO',False);
    SetControlEnabled('CH_VALORISATION',False);
    SetControlEnabled('CH_QUANTITE',False);
    SetControlEnabled('CH_DESIGNATION',False);
    SetControlText('CH_VALEURECLAT', StrFMontant(0.00,15,V_PGI.OkDecV,'',True));
    SetControlText('CH_CODEIMMO', '');
    SetControlText('CH_VALORISATION', StrFMontant(0.00,15,V_PGI.OkDecV,'',True));
    SetControlText('CH_QUANTITE', '');
    SetControlText('CH_DESIGNATION', '');
  end;
end;

{***********UNITE*************************************************
Auteur  ...... : PGR
Créé le ...... : 09/09/2005
Modifié le ... :   /  /
Description .. : Click sur changement des conditions d'amortissement
Mots clefs ... : TOF;AMCHGTMETHODE
*****************************************************************}
procedure TOF_AMCHGTMETHODE.OnClickChCondition(Sender: TObject);
begin
  If GetControlText('CH_CONDITION') = 'X' then
  begin
    SetControlEnabled('CH_METHODE',True);
    SetControlEnabled('CH_DUREE',True);
    SetControlEnabled('CH_ACOMPTER',True);

    { mbo - fq 17569 if fmethode = 'DEG' then
       SetControlText('CH_ACOMPTER',GetControlText('CHI_DATEACHAT'))
    else  SetControlText('CH_ACOMPTER',GetControlText('CHI_SERVICE'));}
    SetControlText('CH_ACOMPTER',GetControlText('CHI_DEBECO'));
  end
  else
  begin
    SetControlEnabled('CH_METHODE',False);
    SetControlEnabled('CH_DUREE',False);
    SetControlEnabled('CH_ACOMPTER',False);
    SetControlText('CH_METHODE', '');
    SetControlText('CH_DUREE', '1');
    SetControlText('CH_ACOMPTER', '');
    SetControlText('CH_TAUX', '0,00');
  end;
end;

{***********UNITE*************************************************
Auteur  ...... : PGR
Créé le ...... : 09/09/2005
Modifié le ... :   /  /
Description .. : Calcul Taux
Mots clefs ... : TOF;AMCHGTMETHODE
*****************************************************************}
procedure TOF_AMCHGTMETHODE.OnExitDuree(Sender: TObject);
var
  taux: double;
begin

  if ((GetControlText('CH_METHODE') <> '')
     and (valeur(GetControlText('CH_ACOMPTER'))<> 0)
     and (not VarIsNull(GetControlText('CH_DUREE')))) then
  begin
    taux := GetTaux(GetControlText('CH_METHODE')
         , StrToDate(GetControlText('CH_ACOMPTER'))
         , StrToDate(GetControlText('CH_ACOMPTER'))
         , StrToInt(GetControlText('CH_DUREE')));
    SetControlText('CH_TAUX',FormatFloat('#,##0.00',taux));
  end
  else taux := 0;

  SetControlText('CH_TAUX',FormatFloat('#,##0.00',taux));
end;

{***********UNITE*************************************************
Auteur  ...... : PGR - FQ17326
Créé le ...... : 17/01/2006
Modifié le ... :   /  /
Description .. : Calcul Quantité en fonction du montant à éclater
Mots clefs ... : TOF;AMCHGTMETHODE
*****************************************************************}
procedure TOF_AMCHGTMETHODE.OnExitValeurAEclater(Sender: TObject);
var
  aeclater: double;
  newquantite: double;
begin
  aeclater :=Valeur(GetControlText('CH_VALEURECLAT'));

  if (aeclater = fmontantht) or (fquantite = 1) then newquantite:=Arrondi(fquantite,0)
                                                else newquantite:=Arrondi(aeclater/(fmontantht/fquantite),0);

  SetControlText('CH_QUANTITE',FormatFloat('##0',newquantite));
end;

{***********UNITE*************************************************
Auteur  ...... : PGR - FQ17326
Créé le ...... : 17/01/2006
Modifié le ... :   /  /
Description .. : Calcul montant à éclater en fonction de la quantité
Mots clefs ... : TOF;AMCHGTMETHODE
*****************************************************************}
procedure TOF_AMCHGTMETHODE.OnExitQuantite(Sender: TObject);
var
  newaeclater: double;
  quantite: double;
begin
  quantite := Arrondi(Valeur(GetControlText('CH_QUANTITE')),0);
  if quantite < 1 then quantite := 1;
  SetControlText('CH_QUANTITE',FormatFloat('##0',quantite));
  newaeclater := Arrondi(quantite*(fmontantht/fquantite),V_PGI.OkDecV);
  SetControlText('CH_VALEURECLAT',StrFMontant(newaeclater,13,V_PGI.OkDecV,'',True));
end;

{***********UNITE*************************************************
Auteur  ...... : PGR
Créé le ...... : 09/09/2005
Modifié le ... :   /  /
Description .. : OnAfterFormShowEcran
Mots clefs ... : TOF;AMCHGTMETHODE
*****************************************************************}
procedure TOF_AMCHGTMETHODE.OnAfterFormShowEcran;
begin
  SetControlProperty('CH_VALEURECLAT', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('CH_VALORISATION', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
end;

{***********UNITE*************************************************
Auteur  ...... : PGR
Créé le ...... : 09/09/2005
Modifié le ... :   /  /
Description .. : Calcul durée déjà amortie
Mots clefs ... : TOF;AMCHGTMETHODE
*****************************************************************}
procedure TOF_AMCHGTMETHODE.DureeDejaAmortie;
var
  PremMois, PremAnnee, NbMois: Word;
  nJour: integer;
  date_immo: TdateTime;

begin
  //si méthode est dégressif : différence entre date d'achat et date fin ex.précédent
  //sinon                    : différence entre date de mise en service et date fin ex.précédent
  if fmethode = 'DEG' then
    date_immo := StrToDate(GetControlText('CHI_DATEACHAT'))
  else date_immo := StrToDate(GetControlText('CHI_SERVICE'));

  if (date_immo >= VHImmo^.Encours.Deb) or (VarIsNull(date_immo)) or (fmethode = 'NAM') then
  begin
    SetControlText('CHI_DUREEAMORTIE', '');
    SetControlText('CHI_LIDUREEAMORTIE', '');
    exit;
  end;

  NOMBREMOIS(date_immo, VHImmo^.Encours.Deb, PremMois, PremAnnee,NbMois);
  nJour := NombreJour360(date_immo,VHImmo^.Encours.Deb-1);

  if nJour < 0 then nJour := 0;

  SetControlCaption('CHI_DUREEAMORTIE', Format(HM2[0], [NbMois - 1, nJour]));
end;

{***********UNITE*************************************************
Auteur  ...... : PGR
Créé le ...... : 09/09/2005
Modifié le ... :   /  /
Description .. : Click sur fermeture
Mots clefs ... : TOF;AMCHGTMETHODE
*****************************************************************}
procedure TOF_AMCHGTMETHODE.OnbFermeClick(Sender: TObject);

begin

  if (GetControlText('CH_CONDITION') <> 'X') and (GetControlText('CH_ECLATEMENT') <> 'X') then
  begin
    // Pas de saisie on sort
    Ecran.ModalResult := mrNo;
  end
  else
  begin
    // Action Echap ou Fermer par la croix
    case PGIAskCancel(HM2[9],Ecran.Caption) of
      mrYes:
        // contrôle validation et sortie
        TToolbarButton97(GetControl('BValider')).Click;
      mrNo:
        // Sortir
        Ecran.ModalResult := mrNo;
      else
        // Rester dans l'écran
        Ecran.ModalResult := mrNone;
    end;
  end;
end;



Initialization
  registerclasses ( [ TOF_AMCHGTMETHODE ] ) ;
end.

