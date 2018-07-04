{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 09/02/2005
Modifié le ... : 09/02/2005
Description .. : Gestion des sorties
Mots clefs ... :
Modif ........ : FQ 16717 - mbo 19.09.2005 - pb suite saisie d'une cession avec présence d'une dot except°
                 Le champ I_opechangeplan est réinitialisé à tort
Modif ........ : FQ 15280 - mbo 30.09.2005 - dans le cas ou pas de plan fiscal zone il_dotcessfis était fausse
                 TGA 04/04/2006 maj I_REPRISEDEPCEDEE
Suite ........ : FQ 17629 - BTY 04/06   Choix du calcul de la PM Value
Suite......... : MBO - 09/2006 - modif pour gestion prime d'équipement
Suite......... : BTY - 10/06 - FQ 18937 PBs en cession dans certains cas :
Suite......... :   . mauvaises valeurs de PValue / VNC
Suite......... :   . absence d'enreg dans IMMOLOG
Suite......... : Causes PutValue trop complexes et stockage du BlocNotes non standard
Suite......... : MBO - 10/2006 - modif pour gestion Subvention d'équipement
Suite......... : MBO - 24/10/2006 - stockage des antérieurs saisis subvention ds immolog lors de la cession
Suite......... : MBO - 20/03/2007 - FQ 17512 - Prise en compte des antérieurs dérogatoire et réintégration fiscale
Suite......... : MBO - 02/10/2007 - ajout param true à TraiteSortie pour nveau param 'amort du jour de sortie'
Suite..........: BTY - 18/09/07 - Nouveaux paramètres immos : Ne pas amortir le jour de cession si SO_AMORTFSORTIE est décoché
*****************************************************************}
unit AmSortie;

interface

uses
  { Delphi }
  classes
  , SysUtils
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  {$IFNDEF DBXPRESS} ,dbtables {$ELSE} ,uDbxDataSet {$ENDIF}
  {$ENDIF}
  { AGL }
  , uTOB
  , HCtrls
  , HEnt1
  { Immobilisations }
  , ImEnt
  , ImPlan
  , Outils
  //, UtobDebug
  ;

type
  TAmSortie = class
    private
      fCodeImmo   : string;
      fTobImmo    : TOB;
      fTobLog     : TOB;
      fPlanAmort  : TPlanAmort;
      fDateOp     : TDateTime;
      fMotif      : string;
      fModeCalc   : string;
      fPrix       : double;
      fTVA        : double;
      fMontantExc : double;
      fTypeExc    : string;
      fBlocNote   : string ;  // FQ 18937 HTStringList;
      fOrdreSerie : integer;
      fbSortie    : boolean;
      fRegleCession : string; // BTY 04/06 FQ 17629
      fAmortJsortie : string; // MBO 30/10/2007 - FQ 21754
      procedure Enregistre;
      procedure Annule;
      procedure SetBlocnote(const Value: string ); // FQ 18937 HTStringList);
      procedure SauveParametreSortie;
    public
      constructor Create ( stCodeImmo : string );
      destructor  Destroy; override;
      procedure   Charge ( stCodeImmo : string );

      // BTY 04/06 FQ 17629 Paramètre supplémentaire = règle de calcul de la PMVAlue
      //procedure   Init ( DateOp : TDateTime; stMotif, stModeCalc : string; Prix : double; TVA : double );
      //mbo - fq 21754 -30/10/2007 - ajout du param amortJsortie
      procedure   Init ( DateOp : TDateTime; stMotif, stModeCalc : string; Prix : double; TVA : double ; RegleCession, AmortJsortie : string);

      procedure   InitExceptionnel ( MontantExc : double; TypeExc : string );
      function    Execute : boolean;
    published
      property Blocnote : string write SetBlocnote; // FQ 18937  HTStringList write SetBlocnote;
      property EstSortie : boolean read fbSortie;
      property DateOp : TDateTime read fDateOp;
      property Motif : string read fMotif;
      property ModeCalc : string read fModeCalc;
      property Prix : double read fPrix;
      property TVA : double read fTVA;
      property MontantExc : double read fMontantExc;
      property TypeExc : string read fTypeExc;
      property RegleCession : string read fRegleCession;   // BTY 04/06 FQ 17629
      property AmortJsortie : string read fAmortJsortie;   // MBO 30/10/2007 - FQ 21754
  end;

implementation
uses ParamSoc, StdCtrls;  // 18/09/07 Nouveaux paramètres immos (jour de sortie)

{ TAmSortie }

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 09/02/2005
Modifié le ... :   /  /
Description .. : Annulation de la sortie en cours
Mots clefs ... :
*****************************************************************}
procedure TAmSortie.Annule;
var
  Plan : TPlanAmort;
begin
  { Mémorisation des paramètres de sortie }
  SauveParametreSortie;
  { Suppression du plan d'amortissement }
  ExecuteSQL('DELETE FROM IMMOAMOR WHERE IA_IMMO="'+fTobImmo.GetValue('I_IMMO')+
                 '" AND IA_NUMEROSEQ='+IntToStr(fTobImmo.GetValue('I_PLANACTIF'))) ;
  { Mise à jour de la fiche }
  fTobImmo.PutValue('I_VNC',fTobLog.GetValue('IL_VNC'));
  if fTobLog.GetValue('IL_QTECEDEE') <> 0 then
    fTobImmo.PutValue('I_QUANTITE',fTobLog.GetValue('IL_QTECEDEE'))
  else fTobImmo.PutValue('I_QUANTITE',1);
  fTobImmo.PutValue('I_QTCEDE',0);
  fTobImmo.PutValue('I_MONTANTHT',fTobLog.GetValue('IL_VOCEDEE')-fTobLog.GetValue('IL_TVARECUPERABLE')+fTobLog.GetValue('IL_TVARECUPEREE'));
  fTobImmo.PutValue('I_TVARECUPERABLE',fTobLog.GetValue('IL_TVARECUPERABLE'));
  fTobImmo.PutValue('I_TVARECUPEREE',fTobLog.GetValue('IL_TVARECUPEREE'));
  if (fTobLog.GetValue('IL_BASETAXEPRO')<>0) then
    fTobImmo.PutValue('I_BASETAXEPRO',fTobLog.GetValue('IL_BASETAXEPRO'))
  else fTobImmo.PutValue('I_BASETAXEPRO',fTobImmo.GetValue('I_MONTANTHT'));
  if (fTobLog.GetValue('IL_BASEECOAVMB')<>0) then
    fTobImmo.PutValue('I_BASEECO',fTobLog.GetValue('IL_BASEECOAVMB'))
  else fTobImmo.PutValue('I_BASEECO',fTobImmo.GetValue('I_MONTANTHT'));
  if (fTobLog.GetValue('IL_BASEFISCAVMB')<>0) then
    fTobImmo.PutValue('I_BASEFISC',fTobLog.GetValue('IL_BASEFISCAVMB'))
  else fTobImmo.PutValue('I_BASEFISC',fTobImmo.GetValue('I_MONTANTHT'));
  fTobImmo.PutValue('I_BASEAMORFINEXO',fTobLog.GetValue('IL_VOCEDEE'));
  fTobImmo.PutValue('I_REPRISEECO',fTobLog.GetValue('IL_REPRISEECO'));
  fTobImmo.PutValue('I_REPCEDECO',0);
  fTobImmo.PutValue('I_REPRISEFISCAL',fTobLog.GetValue('IL_REPRISEFISC'));
  fTobImmo.PutValue('I_REPCEDFISC',0);
  // BTY 04/06 FQ 17629 RAZ Règle de calcul de PMVAlue
  fTobImmo.PutValue('I_REGLECESSION', 'NOR');

  // MBO 20/03/2007 FQ 17512
  fTobImmo.PutValue('I_REPRISEDR',fTobImmo.GetValue('I_REPRISEFDRCEDEE'));
  fTobImmo.PutValue('I_REPRISEFEC',fTobImmo.GetValue('I_REPRISEFECCEDEE'));
  fTobImmo.PutValue('I_REPRISEFDRCEDEE',0);
  fTobImmo.PutValue('I_REPRISEFECCEDEE',0);

  if (fTobImmo.GetValue('IL_TYPEEXC')<> '') and (fTobImmo.GetValue('I_MONTANTEXC')=0) then
    fTobImmo.PutValue('I_OPECHANPLAN','-');
  fTobImmo.PutValue('I_OPECESSION','-');
  fTobImmo.PutValue('I_PLANACTIF',fTobLog.GetValue('IL_PLANACTIFAV'));
  Plan:=TPlanAmort.Create(true) ;
  Plan.ChargeTOB ( fTOBImmo);
  fTobImmo.PutValue('I_DATEDERMVTECO',Plan.GetDateFinAmortEx(Plan.AmortEco));
  fTobImmo.PutValue('I_DATEDERNMVTFISC',Plan.GetDateFinAmortEx(Plan.AmortFisc));
  Plan.free ;
  fTobImmo.UpdateDB;

  { Suppression du Log }
  ExecuteSQL ('DELETE FROM IMMOLOG WHERE IL_IMMO="'+fTobImmo.GetValue('I_IMMO')+'" AND IL_TYPEOP="CES"');

  { Rechargement des données }
  Charge ( fTobImmo.GetValue('I_IMMO') );
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 14/02/2005
Modifié le ... :   /  /
Description .. : Chargement des informations de l'immobilisation concernée
Suite ........ : par la sortie
Mots clefs ... :
*****************************************************************}
procedure TAmSortie.Charge(stCodeImmo: string);
var Q : TQuery;
begin
  { Si des éléments ont déjà été chargé, on les détruit }
  FreeAndNil (fTobLog);
  FreeAndNil (fPlanAmort);
  FreeAndNil (fTobImmo);

  fbSortie := False;
  fCodeImmo   := stCodeImmo;
  { Chargement de l'enregistrement de la table IMMO }
  Q := OpenSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+fCodeImmo+'"',True);
  try
    fTobImmo := TOB.Create ('IMMO', nil, -1);
    fTobImmo.SelectDB ('',Q);
    { Chargement du plan d'amortissement }
    fPlanAmort := TPlanAmort.Create(True);
    fPlanAmort.Charge (Q);
    fPlanAmort.Recupere(fCodeImmo,fTobImmo.GetValue('I_PLANACTIF'));
  finally
    Ferme (Q);
  end;
  { Chargement du LOG si l'immobilisation est déjà sortie }
  Q := OpenSQL ('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+fCodeImmo+'" ORDER BY IL_ORDRE DESC',True);
  try
    if not Q.Eof then
    begin
      Q.First;
      if (Q.FindField('IL_TYPEOP').AsString='CES') then
      begin
        fTobLog := TOB.Create ('IMMOLOG', nil, -1);
        fTobLog.SelectDB ('',Q);
      end;
    end;
  finally
    Ferme (Q);
  end;
  fbSortie := (fTobLog <> nil);
  { Mémorisation des paramètres de sortie }
  if fbSortie then SauveParametreSortie;
end;

constructor TAmSortie.Create( stCodeImmo: string );
begin
  fCodeImmo   := stCodeImmo;
  fMontantExc := 0;
  fTypeExc    := '';
  fBlocNote   := '' ; // FQ 18937 HTStringList.Create;
  fOrdreSerie := TrouveNumeroOrdreSerieLogSuivant;
  fbSortie    := False;
  fRegleCession := 'NOR'; // BTY 04/06 FQ 17629

  // mbo - fq 21754 - 30/10/2007
  fAmortJsortie := 'OUI';
  if (GetParamSocSecur('SO_AMORTJSORTIE', True) = cbUnChecked) then
      fAmortJsortie := 'NON';

  if stCodeImmo <> '' then Charge (fCodeImmo);
end;

destructor TAmSortie.Destroy;
begin
  // FQ 18937 FreeAndNil (fBlocNote);
  FreeAndNil (fTobLog);
  FreeAndNil (fPlanAmort);
  FreeAndNil (fTobImmo);
  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 09/02/2005
Modifié le ... : 09/02/2005
Description .. : Enregistrement de la sortie en base
Mots clefs ... :
*****************************************************************}
procedure TAmSortie.Enregistre;
begin
  { Fiche }
  fTobImmo.UpdateDB;
  { Plan d'amortissement }
  fPlanAmort.Sauve;
  { Log }
  fTobLog.InsertDB(nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 09/02/2005
Modifié le ... : 14/02/2005
Description .. : Lancement de la procédure (calcul des éléments)
Suite ........ : Cette fonction renvoie Vrai si l'enregistrement s'est bien
Suite ........ : passe et Faux sinon.
Suite ........ : BTY 04/06 FQ 17629
Suite ........ : MBO 30/10/2007 - FQ 21754 - ajout paramètre à fonction traitesortie
Mots clefs ... :
*****************************************************************}
function TAmSortie.Execute : boolean;
var
  CumulAntEco, CumulAntFisc : double;
  dDotEco,dCesEco,dExcEco : double;
  CommentaireOpe : string;
  PValue, VNC : double;
begin
  { Si déjà une sortie en cours, on annule la sortie }
  if fbsortie then
  begin
   fbSortie := not (Transactions(Annule,2) = oeOk);
   Result := not fbSortie;
  end else
  begin
    { Mise à jour du plan d'amortissement }
    fPlanAmort.TypeOpe := 'CES';
    // ajout du dernier param pour prise en cpte du param 'amort du jour de cession'
    fPlanAmort.TraiteSortie (fDateOp, fModeCalc, fMontantExc, fTypeExc, fAmortJsortie);

    { Création de l'enregistrement dans IMMOLOG }

    fTobLog := TOB.Create ('IMMOLOG',nil,-1);
    fTobLog.PutValue('IL_IMMO',fCodeImmo);
    // BTY 04/06 FQ 17629 Prise en compte de la règle de calcul de la PMValue
    //fTobLog.PutValue('IL_LIBELLE',RechDom('TIOPEAMOR', 'CES', FALSE)+' '+DateToStr(fDateop));
    if (fRegleCession='CT') then
        CommentaireOpe := ' (répartition +/- value à CT)'
    else if (fRegleCession='LT') then
        CommentaireOpe := ' (répartition +/- value à LT)'
    else if (fRegleCession='RSD') then
        CommentaireOpe := ' (répart. +/- ignorant la durée)'
    else
        CommentaireOpe := '';
    fTobLog.PutValue('IL_LIBELLE',RechDom('TIOPEAMOR', 'CES', FALSE)+' '+DateToStr(fDateop)+CommentaireOpe);
    fTobLog.PutValue('IL_TYPEMODIF',AffecteCommentaireOperation('CES'));
    fTobLog.PutValue('IL_DATEOP',fDateOp);
    fTobLog.PutValue('IL_TYPEOP','CES');
    fTobLog.PutValue('IL_MOTIFCES',fMotif);
    fTobLog.PutValue('IL_CALCCESSION',fModeCalc);
    fTobLog.PutValue('IL_VOCEDEE',fTobImmo.GetValue('I_MONTANTHT'));
    fTobLog.PutValue('IL_QTECEDEE',fTobImmo.GetValue('I_QUANTITE'));
    fTobLog.PutValue('IL_MONTANTCES',fPrix);
    fTobLog.PutValue('IL_TVAAREVERSER',fTVA);
    fTobLog.PutValue('IL_TVARECUPEREE',fTobImmo.GetValue('I_TVARECUPEREE'));
    fTobLog.PutValue('IL_TVARECUPERABLE',fTobImmo.GetValue('I_TVARECUPERABLE'));
    fTobLog.PutValue('IL_REPRISEECO',fTobImmo.GetValue('I_REPRISEECO'));
    fTobLog.PutValue('IL_REPRISEFISC',fTobImmo.GetValue('I_REPRISEFISCAL'));

    // FQ 18937  Mauvais résultats de PValue/VNC/BN dans certains cas
    // Causes PutValue trop complexes pour PValue / VNC + écriture non standard du BN
    //fTobLog.PutValue('IL_PVALUE', fPrix - ( fPlanAmort.GetVNCAvecMethode(fPlanAmort.AmortEco, VHImmo^.Encours.Fin)));
    //fTobLog.PutValue('IL_VNC',  fPlanAmort.GetVNCAvecMethode(fPlanAmort.AmortEco, VHImmo^.Encours.Fin) );
    //fTobLog.PutValue('IL_BLOCNOTE',fBlocNote.Text);
    PValue := fPrix - ( fPlanAmort.GetVNCAvecMethode(fPlanAmort.AmortEco, VHImmo^.Encours.Fin));
    fTobLog.PutValue('IL_PVALUE', Arrondi(PValue, V_PGI.OkDecV));
    VNC := fPlanAmort.GetVNCAvecMethode(fPlanAmort.AmortEco, VHImmo^.Encours.Fin);
    fTobLog.PutValue('IL_VNC', Arrondi(VNC, V_PGI.OkDecV) );
    fTobLog.PutValue('IL_BLOCNOTE', fBlocNote);
    // fin FQ 18937
    fTobLog.PutValue('IL_PLANACTIFAV',fTobImmo.GetValue('I_PLANACTIF'));
    fTobLog.PutValue('IL_PLANACTIFAP',fTobImmo.GetValue('I_PLANACTIF')+1);
    fTobLog.PutValue('IL_ORDRE',TrouveNumeroOrdreLogSuivant(fCodeImmo));
    fTobLog.PutValue('IL_ORDRESERIE',fOrdreSerie);
    fPlanAmort.GetCumulsDotExercice(VHImmo^.Encours.Deb,CumulAntEco,CumulAntFisc, True, true, False);
    fTobLog.PutValue('IL_DOTCESSECO',Arrondi(fPlanAmort.GetDotationGlobale(1,VHImmo^.Encours.Fin,dDotEco,dCesEco,dExcEco),V_PGI.OKDecV));
    // 30/09/2005 - ajout mbo FQ 15280 -
    // fTobLog.PutValue('IL_DOTCESSFIS',Arrondi(fPlanAmort.GetDotationGlobale(2,VHImmo^.Encours.Fin,dDotEco,dCesEco,dExcEco),V_PGI.OKDecV));
    if fPlanAmort.AmortFisc.Methode<>'' then
       fTobLog.PutValue('IL_DOTCESSFIS',Arrondi(fPlanAmort.GetDotationGlobale(2,VHImmo^.Encours.Fin,dDotEco,dCesEco,dExcEco),V_PGI.OKDecV))
    else
       fTobLog.PutValue('IL_DOTCESSFIS',0);

    fTobLog.PutValue('IL_CUMANTCESECO',Arrondi(CumulAntEco,V_PGI.OKDecV));
    fTobLog.PutValue('IL_CUMANTCESFIS',Arrondi(CumulAntFisc,V_PGI.OKDecV));
    fTobLog.PutValue('IL_MONTANTEXC',fMontantExc);
    fTobLog.PutValue('IL_TYPEEXC',fTypeExc);
    fTobLog.PutValue('IL_BASEECOAVMB',fTobImmo.GetValue('I_BASEECO'));
    fTobLog.PutValue('IL_BASEFISCAVMB',fTobImmo.GetValue('I_BASEFISC'));
    fTobLog.PutValue('IL_BASETAXEPRO',fTobImmo.GetValue('I_BASETAXEPRO'));

    // ajout mbo 24.10.06 stockage des antérieurs subvention pour pouvoir les restituer en cas d'annulation sortie
    fTobLog.PutValue('IL_MONTANTAVMB',fTobImmo.GetValue('I_CORRECTIONVR'));

    // 18/09/07 Nouveaux paramètres immos
    {
    if fTobImmo.GetValue('I_METHODEECO')='LIN' then
    begin
      fTobLog.PutValue('IL_CODECB', 'OUI');
      if (GetParamSocSecur('SO_AMORTJSORTIE', True) = cbUnChecked) then
          fTobLog.PutValue('IL_CODECB', 'NON');
    end;
    }
    // 18/09/07 Fin Nouveaux paramètres immos

    // fq 21754 - mbo - 30/10/2007
    if fTobImmo.GetValue('I_METHODEECO')='LIN' then
      fTobLog.PutValue('IL_CODECB', fAmortJsortie);

    { Mise à jour de la fiche  }
    fTobImmo.PutValue('I_OPERATION','X');
    fTobImmo.PutValue('I_OPECESSION','X');
    if fTypeExc<>'' then  fTobImmo.PutValue('I_OPECHANGEPLAN','X');
    // mbo - FQ 16717 - 19.09.2005 - else fTobImmo.PutValue('I_OPECHANGEPLAN','-');
    fTobImmo.PutValue('I_MONTANTHT',0);
    fTobImmo.PutValue('I_VNC',fTobLog.GetValue('IL_VNC'));

    // modif FQ 12437 mbo 24.08.05 fTobImmo.PutValue('I_BASETAXEPRO',0);
    fTobImmo.PutValue('I_BASETAXEPRO',fTobImmo.GetValue('I_BASETAXEPRO'));

    fTobImmo.PutValue('I_BASEECO',0);
    fTobImmo.PutValue('I_BASEFISC',0);
    fTobImmo.PutValue('I_QTCEDE',fTobImmo.GetValue('I_QUANTITE'));
    fTobImmo.PutValue('I_QUANTITE',0);
    fTobImmo.PutValue('I_BASEAMORFINEXO',0);
    fTobImmo.PutValue('I_PLANACTIF',fTobImmo.GetValue('I_PLANACTIF')+1);
    fTobImmo.PutValue('I_REPCEDECO',fTobImmo.GetValue('I_REPRISEECO'));
    fTobImmo.PutValue('I_REPCEDFISC',fTobImmo.GetValue('I_REPRISEFISCAL'));
    fTobImmo.PutValue('I_REPRISEECO',0);
    fTobImmo.PutValue('I_REPRISEFISCAL',0);
    fTobImmo.PutValue('I_DATEDERMVTECO', fDateOp);

    // TGA 04/04/2006
    fTobImmo.PutValue('I_REPRISEDEPCEDEE',fTobImmo.GetValue('I_REPRISEDEP'));
    fTobImmo.PutValue('I_REPRISEDEP',0);
    fTobImmo.PutValue('I_DATECESSION', fDateOp);

    // BTY 04/06 FQ 17629 Nouveau paramètre = règle de calcul de la PMValue
    fTobImmo.PutValue('I_REGLECESSION',fRegleCession);

    if fPlanAmort.Fiscal then fTobImmo.PutValue('I_DATEDERNMVTFISC',fDateOp);

    // MBO 09/2006 - gestion de la prime d'équipement
    if fTobImmo.GetValue('I_SBVPRI') <> 0 then
    begin
       fTobImmo.PutValue('I_REPRISEUOCEDEE',fTobImmo.GetValue('I_REPRISEUO'));
       fTobImmo.PutValue('I_SBVPRIC',fTobImmo.GetValue('I_SBVPRI'));
       fTobImmo.PutValue('I_REPRISEUO',0);
       fTobImmo.PutValue('I_SBVPRI',0);
    end;

    // MBO 10/2006 - Gestion de la subvention d'équipement
    if fTobImmo.GetValue('I_SBVMT') <> 0 then
    begin
       fTobImmo.PutValue('I_CORRECTIONVR', 0); // on remet à 0 antérieurs saisis (sauvegarde dans immolog avant)
       fTobImmo.PutValue('I_SBVMTC',fTobImmo.GetValue('I_SBVMT'));
       fTobImmo.PutValue('I_SBVMT',0);
    end;

    //MBO 20/03/2007 - FQ 17512 - Gestion des antérieurs dérogatoire / réintégration fiscale
    fTobImmo.PutValue('I_REPRISEFDRCEDEE',fTobImmo.GetValue('I_REPRISEDR'));
    fTobImmo.PutValue('I_REPRISEFECCEDEE',fTobImmo.GetValue('I_REPRISEFEC'));
    fTobImmo.PutValue('I_REPRISEFEC',0);
    fTobImmo.PutValue('I_REPRISEDR',0);

    fTobImmo.UpdateDateModif;

    fbSortie := (Transactions(Enregistre,2) = oeOk);
    Result := fbSortie;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 09/02/2005
Modifié le ... :   /  /
Description .. : Initialisation de la procédure de sortie
Mots clefs ... :
*****************************************************************}
//procedure TAmSortie.Init(DateOp: TDateTime; stMotif, stModeCalc: string;
//  Prix, TVA: double);
procedure TAmSortie.Init(DateOp: TDateTime; stMotif, stModeCalc: string; Prix, TVA: double ; RegleCession, AmortJsortie : string);
begin
  fDateOp     := DateOp;
  fMotif      := stMotif;
  fModeCalc   := stModeCalc;
  fPrix       := Prix;
  fTVA        := TVA;
  fRegleCession := RegleCession;  // BTY 04/06 FQ 17629
  fAmortJsortie := AmortJsortie;  // mbo 30/10/2007 - fq 21754
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 14/02/2005
Modifié le ... :   /  /
Description .. : Intialisation de l'exceptionnel éventuel
Mots clefs ... :
*****************************************************************}
procedure TAmSortie.InitExceptionnel(MontantExc: double; TypeExc: string);
begin
  fMontantExc := MontantExc;
  fTypeExc := TypeExc;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 27/10/2005
Modifié le ... :   /  /
Description .. : Enregistrement des paramètres de sortie pour annulation
Suite ........ : future
Mots clefs ... :
*****************************************************************}
procedure TAmSortie.SauveParametreSortie;
begin
  if fTobLog = nil then exit;
  fDateOp     := fTobLog.GetValue('IL_DATEOP');
  fMotif      := fTobLog.GetValue('IL_MOTIFCES');
  fModeCalc   := fTobLog.GetValue('IL_CALCCESSION');
  fPrix       := fTobLog.GetValue('IL_MONTANTCES');
  fTVA        := fTobLog.GetValue('IL_TVAAREVERSER');
  fMontantExc := fTobLog.GetValue('IL_MONTANTEXC');
  fTypeExc    := fTobLog.GetValue('IL_TYPEEXC');
  fOrdreSerie := fTobLog.GetValue('IL_ORDRESERIE');

  //fq 21754 - récup du paramètre de calcul de cession : prise en cpte du jour de cession
  fAmortJsortie := fTobLog.GetValue('IL_CODECB');

  // BTY 04/06 FQ 17629 Paramètre supplémentaire = Règle de calcul de la PMValue
  if fTobImmo <> nil then
     fRegleCession := fTobImmo.GetValue('I_REGLECESSION');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 14/02/2005
Modifié le ... :   /  /
Description .. : Mise à jour du bloc note
Mots clefs ... :
*****************************************************************}
// FQ 18937 Passer le BlocNotes par une string, pour compatibilité CWAS
{procedure TAmSortie.SetBlocnote(const Value: HTStringList);
begin
 fBlocNote.Assign(Value); // FQ 18937
end; }
procedure TAmSortie.SetBlocnote(const Value: string );
begin
 fBlocNote := Value;
end;

end.
