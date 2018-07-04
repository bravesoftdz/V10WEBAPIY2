{***********UNITE*************************************************
Auteur  ...... : PAIE MF
Créé le ...... : 30/03/2007
Modifié le ... : 30/03/2007
Description .. : Procédures non interactives d'initialisation et d'impression
Suite ........ : des DUCS.
Suite ........ : Pour traiter les états chaînés et le Processserver
Mots clefs ... : PAIE; DUCS; PROCESS SERVER; ETATSCHAINES
*****************************************************************}
{
 PT1 : 02/07/2007 : V_72   MF FQ 14473 : alimentation des type de bordereau quand non renseignés
 PT4 : 20/06/2008 : V_810  MF CWAS : on vide MaTob et MATobAT
}

unit PgDucsOutils;

interface
uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  Classes,
{$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$IFNDEF EAGLSERVER}
  Fiche,
  PgEdtEtat,
{$ENDIF}
  {$IFDEF V530}
  EdtEtat,
  {$ELSE}
  EdtREtat,
  {$ENDIF}
{$ELSE}
  eFiche,UtilEagl,
{$ENDIF}
  EntPaie,
  HCtrls,
  Hent1,
  hmsgbox,
  P5Def,
  P5Util,
  ParamSoc,
  PgOutils,
  PgOutils2,
  sysutils,
  Utob,
  ULibEditionPaie;

Type TDucsDeclarant = record
  Emetteur        : string;
  RaisonSoc       : string;
  ContactDecla    : string;
  TelDeclarant    : string;
  FaxDeclarant     : string;
end;

Type TDucsParametres = record
  Etab            : string;
  Organisme       : string;
  Regroupt        : string;
  NatureDucs      : string;
  DucsDoss        : Boolean;
  PaieGroupe      : Boolean;
end;

Type TDucsVariables = record
  WSiret          : string;
  WApe            : string;
  WNumInt         : string;
  WGpInt          : string;
  InstEnCours     : string;
  CodifEnCours    : string;
  NoMere          : integer;
  CodifPosDebut   : Integer;
  CodifLongEdit   : Integer;
  CondSpec        : Boolean;
  Rupture         : Boolean;
  RupSiret        : Boolean;
  RupApe          : Boolean;
  RupNumInt       : Boolean;
  RupGpInt        : Boolean;
  LigneOptique    : string;
  BaseArr         : string;
  WBaseArr        : string;
  MontArr         : string;
  WMontArr        : string;
  JExigible       : Integer;
  JLimDepot       : Integer;
  JReglement      : Integer;
  DatePaye        : TDateTime;
  EtSiret         : string;
  EtApe           : string;
  NumDucs         : integer;
  EtabEnCours     : string;
  SiretEnCours    : string;
  ApeEnCours      : string;
  OrganismeEnCours: string;
  NumIntEnCours   : string;
  GpIntEnCours    : string;
  IndNeant2       : boolean;
  StWhere         : string;
  StWhereDet      : string;
end;

Type TDucsEffectifs = record
    EffAppH       : double;
    EffAppF       : double;
    EffCadH       : double;
    EffCadF       : double;
    EffTot        : double;
    EffHom        : double;
    EffFem        : double;
    EffApp        : double;
    EffCAd        : double;
    EffOuv        : double;
    EffEtam       : double;
    EffCDIH       : double;
    EffCDIF       : double;
    Eff65H        : double;
    Eff65F        : double;
    EffProfH      : double;
    EffProfF      : double;
    EffCDDH       : double;
    EffCDDF       : double;
    EffCNEH       : double;
    EffCNEF       : double;
    EffCadH22     : double;
    EffCadH23     : double;
    EffCadH24     : double;
    EffCadH25     : double;
    EffCadF22     : double;
    EffCadF23     : double;
    EffCadF24     : double;
    EffCadF25     : double;
end;

Type TDucsOrganisme = record
    PogNumInterne  : string;
    PogNature      : string;
    PogBoolST      : Boolean;
    PogSTPos       : Integer;
    PogSTLong      : Integer;
    PogPeriod1     : string;
    PogPeriod2     : string;
    PogTypBordereau: string;
    PogTypDecl1    : Boolean;
    PogTypDecl2    : Boolean;
    PogLGOptique   : string;
    PogCPInfo      : string;
    PogCentrePayeur: string;
end;


procedure DucsInit(PeriodEtat,DebPer, FinPer, NatDeb, NatFin, EtabDeb,EtabFin : string);
{$IFNDEF EAGLSERVER}
procedure DucsEdit (PeriodEtat,DebPer, FinPer, NatDeb, NatFin, EtabDeb,EtabFin : string; Neant : boolean;var StSQL : string);
{$ENDIF}
procedure CalculPeriode (var PeriodEtat, DebPer, FinPer : string);
procedure InfosDeclarant (var DucsDeclarant : TDucsDeclarant);
procedure ChargeDucsAfaire (PeriodEtat, NatDeb, NatFin, EtabDeb,EtabFin : string;var TDucs : TOB);
procedure ChargeListeInstit (DucsParametres : TDucsParametres;  var Tinst : TOB);
procedure InitVarTrait (DucsParametres : TDucsParametres; Debper, FinPer : string; var DucsVariables : TDucsVariables;var DucsEffectifs : TDucsEffectifs;TDucs, Tinst : TOB; var TRupMere : TOB);
procedure RechRupture(DucsParametres : TDucsparametres; DucsVariables : TDucsVariables; TInst :TOB; var TRupMere : TOB );
procedure ChargeCot(DucsParametres : TDucsParametres; DucsVariables : TDucsVariables; DebPer, FinPer : string; TInst : TOB; var TCot3 : TOB; var ErrRupGpInt: integer;var  TraceE :TStringList; IndTraceE : boolean);
procedure MajDucs(DebPer, FinPer, PeriodEtat :string; DucsParametres: TDucsParametres; DucsDeclarant : TDucsDeclarant; DucsEffectifs : TDucsEffectifs; TCot3 : TOB; var DucsOrganisme : TDucsOrganisme; var DucsVariables : TDucsVariables;var  TraceE :TStringList; IndTraceE : boolean; TrupMere, Tinst : TOB);
procedure ChargeWheres(DebPer, FinPer : string; var DucsVariables : TDucsvariables; DucsParametres : TDucsParametres;TRupMere : TOB);
procedure CalculEffectifs(DebPer, FinPer : string; DucsVariables :TDucsvariables;DucsParametres : TDucsParametres;var DucsEffectifs : TDucsEffectifs;TInst : TOB; PogTypBordereau, PogNature : string);
procedure AlimDetail(DebPer, FinPer : string; var TOB_FilleLignes, TOB_FilleTete: TOB; var ligne: Integer; DucsParametres:TDucsParametres; DucsVariables:TDucsVariables;PogNumInterne : String);
procedure LibereTobs(var TInst,TCot3,TRupMere : TOB);

function  fctRupture(var DucsVariables : TDucsVariables; TDetail: TOB; Ind: integer): Boolean;
function  Process_DucsInit(PeriodEtat,DebPer, FinPer, NatDeb, NatFin, EtabDeb,EtabFin : string): TOB;
// d PT1
function  AffectTypeBordereau( NatureOrg, Periodicite : string; DucsDossier : boolean):string;
// f PT1

implementation
var
    AlsaceMoselle : boolean;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE MF
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : initialisation des DUCS : alimentation DUCSENTETE et 
Suite ........ : DUCSDETAIL
Mots clefs ... : PAIE; DUCS
*****************************************************************}
procedure DucsInit (PeriodEtat,DebPer, FinPer, NatDeb, NatFin, EtabDeb,EtabFin : string);
var
//  DebPer, FinPer                         : string;
  DucsDeclarant                          : TDucsDeclarant;
  DucsParametres                         : TDucsParametres;
  DucsVariables                          : TDucsVariables;
  DucsEffectifs                          : TDucsEffectifs;
  DucsOrganisme                          : TDucsOrganisme;
  TDetDucs,TDucs, TInst,TRupMere, TCot3  : TOB;
  i,ErrRupGpInt,IndOR                    : integer;
  Trace, TraceE                          : TStringList;
  IndTraceE                              : boolean;
begin

  Trace := TStringList.Create;
  TraceE := TStringList.Create;
  IndTraceE := false;

  // Calcul des dates de période et de la périodicité
    CalculPeriode (PeriodEtat, DebPer, FinPer);

  Trace.Add ('DUCS : Traitement non interactif d''initialisation des déclarations du '+ DebPer + ' au '+ FinPer);
  Trace.Add ('');

  // Récupération des informations Déclarant
  InfosDeclarant (DucsDeclarant);
  if (DucsDeclarant.Emetteur = '') then
  begin
    if (not IndTraceE) then
    begin
      TraceE.add ('Anomalies de traitement');
      TraceE.add ('--------------------------------------------------------');
      IndTraceE := True;
    end;
    TraceE.add ('Aucun émetteur par défaut n''est renseigné. Information obligatoire pour la Ducs EDI');
  end;

  // Chargement de la TOB TDucs des DUCS à initialiser
  ChargeDucsAfaire(PeriodEtat, NatDeb, NatFin, EtabDeb,EtabFin, TDucs );

  DucsVariables.StWhereDet := 'WHERE (';
  DucsVariables.StWhere := 'WHERE (';
  IndOR := 0;

  for i := 0 to TDucs.Detail.Count-1 do
  begin
    TDetDucs := TDucs.Detail[i];
    if (TDetDucs.GetValue('POG_DUCSDOSSIER') = 'X') then
      DucsParametres.DucsDoss := True
    else
      DucsParametres.DucsDoss := False;
    if (TDetDucs.GetValue('POG_PAIEGROUPE') = 'X') then
      DucsParametres.PaieGroupe := True
    else
      DucsParametres.PaieGroupe := False;

    DucsParametres.Regroupt := TDetDucs.GetValue('POG_REGROUPEMENT');
    DucsParametres.Etab := TDetDucs.GetValue('POG_ETABLISSEMENT');
    DucsParametres.Organisme := TDetDucs.GetValue('POG_ORGANISME');
    DucsParametres.NatureDucs := TDetDucs.GetValue('POG_NATUREORG');

    {On établit la liste des institutions rattachées
     à l'organisme destinataire}
    ChargeListeInstit(DucsParametres, Tinst);

    {Initialisation variables de traitement}
    InitVarTrait(DucsParametres,Debper,FinPer,DucsVariables,DucsEffectifs,TDucs,Tinst,TRupMere);

    {Mémorisation en vue d'édition des critères de sélection des différentes
     DUCS initialisé}
    if (IndOR <> 0) then
    begin
      DucsVariables.StWhere := 'WHERE (';
      DucsVariables.StWhereDet := 'WHERE (';
    end;

    {Initialisation des strings "where" en fonctions des ruptures
     paramètrées en vue de la suppression des ducs déjà existantes
     pour les mêmes critères}
    ChargeWheres(DebPer, FinPer, DucsVariables, DucsParametres, TRupMere);

    IndOR := 1;

    {Liste des cotisations regroupées
     par codification , par tauxat, par taux global et
     par condition de rupture
     pour chaque salarié par base identique pour une période.}
    ErrRupGpInt := 0;
    ChargeCot(DucsParametres,DucsVariables,DebPer, FinPer,TInst,TCot3,ErrRupGpInt,TraceE,IndTraceE);
    if (ErrRupGpInt <> 0) then break;

    {Mise à jour des lignes détail à partir de la TOB TCot  }
    MajDucs(DebPer,FinPer,PeriodEtat,DucsParametres, DucsDeclarant,DucsEffectifs,TCot3,DucsOrganisme,DucsVariables, TraceE,IndTraceE,TrupMere,Tinst);
    if(TCot3.Detail.Count <> 0) then
      Trace.Add ('Etablissement '+ DucsParametres.Etab + ':' +
               'Déclaration '+ TDetDucs.GetValue('POG_LIBELLE'))
    else
      Trace.Add ('Etablissement '+ DucsParametres.Etab + ':' +
               'Déclaration '+ TDetDucs.GetValue('POG_LIBELLE')+' *** Néant ***');

    { free des TOBs utilisées}
    LibereTobs(TInst, TCot3,TRupMere);

  end;
  Trace.Add ('');
  Trace.Add ('Fin du traitement non interactif d''initialisation des DUCS');
  CreeJnalEvt('001', '020', 'OK', nil, nil, Trace, TraceE);

  Trace.Free;
  TraceE.Free;

end;

{$IFNDEF EAGLSERVER}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE MF
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : Impression des Ducs : traitement interactif : menu 42302 et 
Suite ........ : états chaînés
Mots clefs ... : PAIE; DUCS; ETATSCHAINES
*****************************************************************}
procedure DucsEdit (PeriodEtat,DebPer, FinPer, NatDeb, NatFin, EtabDeb,EtabFin : string; Neant : boolean; var StSQL : string );
var
  StWhereEd, StWhereBis                         : string;
  Trace, TraceE                                 : TStringList;
  IndTraceE                                     : boolean;
  OldSilentMode,OldNoPrintDialog                : boolean;

begin
  Trace := TStringList.Create;
  TraceE := TStringList.Create;
  IndTraceE := false;

  OldSilentMode    := V_Pgi.SilentMode;
  OldNoPrintDialog := V_Pgi.NoPrintDialog;

  // Calcul des dates de période et de la périodicité
  CalculPeriode (PeriodEtat, DebPer, FinPer);

  Trace.Add ('DUCS : Traitement d''impression des déclarations du '+ DebPer + ' au '+ FinPer);
  Trace.Add ('');

  V_PGI.NoPrintDialog := True; // pas d'affichage de la fenêtre d'impression
  V_Pgi.SilentMode := True;

  StWhereEd := 'WHERE (';
  StWhereEd := StWhereEd + '(PDU_DATEDEBUT >="' + UsDateTime(StrToDate(DebPer)) + '")'+
                           ' AND (PDU_DATEDEBUT <="' + UsDateTime(StrToDate(FinPer)) +
                            '")'+
                            ' AND (PDU_DATEFIN >="' + UsDateTime(StrToDate(DebPer)) + '")'+
                            ' AND (PDU_DATEFIN <= "' + UsDateTime(StrToDate(FinPer)) + '") AND ';
  if ((EtabDeb <> '') and (EtabFin <> '')) then
    StWhereEd := StWhereEd + 'PDU_ETABLISSEMENT >= "'+ EtabDeb+'" AND '+
                           'PDU_ETABLISSEMENT <= "'+ EtabFin+'" AND ';

  if ((NatDeb <> '') and (NatFin <> '')) then
    StWhereEd := StWhereEd + 'POG_NATUREORG >= "'+ NatDeb+'" AND '+
                             'POG_NATUREORG <= "'+ Copy(NatFin,1,1)+'99" AND ';
  if (Neant) then
  begin
  // IMPRESSION DUCS NEANT
  StWhereBis := 'AND ((PDU_PAIEMENT = "' + UsDateTime(IDate1900) + '") AND ' +
                '(PDU_NBSALFPE = 0) ' +
                'AND (PDU_NBSALHAP = 0))';

  StSQL := 'SELECT DUCSENTETE.*,' +
           'ET_ETABLISSEMENT, ET_LIBELLE, ET_ADRESSE1, ET_ADRESSE2, ET_ADRESSE3, ' +
           'ET_CODEPOSTAL, ET_VILLE, ET_TELEPHONE, ET_FAX, ' +
           'POG_ETABLISSEMENT, POG_ORGANISME, POG_LIBELLE, POG_NATUREORG, ' +
           'POG_ADRESSE1, POG_ADRESSE2, POG_ADRESSE3, POG_CODEPOSTAL, ' +
           'POG_VILLE , POG_LONGEDITABLE, ' +
           'POG_PERIODICITDUCS, POG_AUTREPERIODUCS, POG_PERIODCALCUL, ' +
           'POG_AUTPERCALCUL, ' +
           'POG_POSTOTAL,POG_LONGTOTAL ' +
           'FROM DUCSENTETE ' +
           'LEFT JOIN ETABLISS ON PDU_ETABLISSEMENT = ET_ETABLISSEMENT ' +
           'LEFT JOIN ORGANISMEPAIE ON PDU_ETABLISSEMENT = POG_ETABLISSEMENT ' +
           'AND PDU_ORGANISME = POG_ORGANISME ' +
           StWhereEd + ') ' + StWhereBis;

//  FunctPGDebEditDucs(); {raz n° page}

//  LanceEtat('E', 'PDU', 'DUN', false, False, False, nil, StSQL, '', False);

//  FunctPGFinEditDucs(); {raz AncNumPagSys}
  end
  else
  begin
  {IMPRESSION DUCS NON NEANT}
  StWhereBis := 'AND ((PDU_PAIEMENT <> "' + UsDateTime(IDate1900) + '") OR ' +
                '(PDU_NBSALFPE <> 0) ' +
                'OR (PDU_NBSALHAP <> 0))';

  StSQL := 'SELECT DUCSENTETE.*,PDD_CODIFICATION,PDD_CODIFEDITEE, ' +
           'PDD_LIBELLE, PDD_BASECOTISATION,PDD_TAUXCOTISATION, PDD_MTCOTISAT, ' +
           'PDD_EFFECTIF, PDD_TYPECOTISATION, PDD_INSTITUTION, ' +
           'PDD_LIBELLESUITE, ' +
           'PDD_REGIMEALSACE, '+
           'ET_ETABLISSEMENT, ET_LIBELLE, ET_ADRESSE1, ET_ADRESSE2, ET_ADRESSE3, ' +
           'ET_CODEPOSTAL, ET_VILLE, ET_TELEPHONE, ET_FAX, ' +
           'POG_ETABLISSEMENT, POG_ORGANISME, POG_LIBELLE, POG_NATUREORG, ' +
           'POG_ADRESSE1, POG_ADRESSE2, POG_ADRESSE3, POG_CODEPOSTAL, ' +
           'POG_PERIODICITDUCS, POG_AUTREPERIODUCS, POG_PERIODCALCUL, ' +
           'POG_AUTPERCALCUL, ' +
           'POG_POSTOTAL,POG_LONGTOTAL, ' +
           'POG_VILLE , POG_LONGEDITABLE,POG_BASETYPARR,' +
           'POG_MTTYPARR ' +
           'FROM DUCSENTETE ' +
           'LEFT JOIN ETABLISS ON PDU_ETABLISSEMENT = ET_ETABLISSEMENT ' +
           'LEFT JOIN ORGANISMEPAIE ON PDU_ETABLISSEMENT = POG_ETABLISSEMENT ' +
           'AND PDU_ORGANISME = POG_ORGANISME ' +
           'LEFT JOIN DUCSDETAIL ON  PDU_ETABLISSEMENT =  PDD_ETABLISSEMENT ' +
           'AND PDU_ORGANISME = PDD_ORGANISME AND PDU_DATEDEBUT = PDD_DATEDEBUT ' +
           'AND PDU_DATEFIN = PDD_DATEFIN AND PDU_NUM = PDD_NUM ' +
           StWhereEd + ') ' + StWhereBis + ' ' +
           'ORDER BY PDU_ETABLISSEMENT,PDU_ORGANISME,PDU_DATEDEBUT,PDU_DATEFIN,' +
           'PDU_NUM,PDD_INSTITUTION,PDD_CODIFICATION';

//  FunctPGDebEditDucs(); {raz n° page}
//  V_PGI.QRPDFMerge := '';
//  if (V_PGI.QRPDFQueue <> '') and FileExists(V_PGI.QRPDFQueue) then
//    V_PGI.QRPDFMerge := V_PGI.QRPDFQueue;
//  LanceEtat('E', 'PDU', 'DUC', true, False, False, nil, StSQL, '', False);
//  V_PGI.NoPrintDialog := True; // pas d'affichage de la fenêtre d'impression


//  FunctPGFinEditDucs(); {raz AncNumPagSys}
  end;

  V_Pgi.SilentMode := OldSilentMode ;
  V_Pgi.NoPrintDialog := OldNoPrintDialog  ;

  Trace.Add ('');
  if (Neant) then
    Trace.Add ('Fin du traitement d''impression des DUCS Néant')
  else
    Trace.Add ('Fin du traitement d''impression des DUCS');
  CreeJnalEvt('001', '021', 'OK', nil, nil, Trace, TraceE);

  Trace.Free;
  TraceE.Free;

end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE MF
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : Calcul de la période et des dates période
Mots clefs ... : PAIE; DUCS
*****************************************************************}
procedure  CalculPeriode (var PeriodEtat, DebPer, FinPer : string);
var
  ExerPerEncours, MoisE, AnneeE: string;
  DebExer, FinExer: TDateTime;
  DebTrim, FinTrim, DebSem, FinSem : TDateTime;
  NbreMois, NbreJour : integer;

begin
  if (DebPer = '') or (FinPer = '') or (DebPer = '  /  /    ') or (FinPer = '  /  /    ') then
  begin
  { Date par défaut (qd PeriodeEtat = ''):
    La date proposée est le trimestre en cours si la période
    en cours correspond à une fin de trimestre, sinon c'est le mois en cours.}
  if RendExerSocialEnCours(MoisE, AnneeE, ExerPerEncours, DebExer, FinExer) = True then
  begin
    RendPeriodeEnCours(ExerPerEnCours, DebPer, FinPer);
    RendTrimestreEnCours(StrToDate(DebPer), DebExer, FinExer, DebTrim, FinTrim, DebSem, FinSem);
    if ((PeriodEtat = '') or (PeriodEtat = 'T')) then
    begin
      if FindeMois(StrToDate(DebPer)) = FinTrim then
      // trimestrielle
      begin
        DebPer :=  DateToStr(DebTrim);
        FinPer :=  DateToStr(FinTrim) ;
      end;
    end
    else
    if (PeriodEtat = 'A') then
    // Annuelle
    begin
      DebPer :=  DateToStr(DebExer);
      FinPer :=  DateToStr(DebExer);
    end;
  end;
  end;

  { On détermine la périodicité des DUCS à initialiser
  Mensuelle, Trimestrielle, Semestrielle, Annuelle ou Quelconque
  Remarque : Il faut que la période corresponde à un mois ou un trimestre
  ou un semestre de l'exercice (Par exemple Février/mars/avril n'est pas un
  trimestre) sinon elle sera considérée comme quelconque (PeriodEtat = '')}

  if (PeriodEtat = '') then
  begin
    DiffMoisJour(StrToDate(DebPer), StrToDate(FinPer), NbreMois, NbreJour);
    NbreMois := NbreMois + 1;
    PeriodEtat := 'M';
    if (NbreMois = 3) then PeriodEtat := 'T';
    if (NbreMois = 6) then PeriodEtat := 'S';
    if (NbreMois = 12) then PeriodEtat := 'A';
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PAIE MF
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : Récupération des informations concernant le déclarant
Mots clefs ... : PAIE; DUCS
*****************************************************************}
procedure InfosDeclarant (var DucsDeclarant : TDucsDeclarant);
var
  Q: Tquery;
begin
  { récupération des infos déclarant }
  DucsDeclarant.Emetteur := GetParamSocSecur('SO_PGEMETTEUR','');

  if (DucsDeclarant.Emetteur<>'') then
  begin
    Q := OpenSQL('SELECT PET_EMETTSOC,PET_RAISONSOC,PET_CONTACTDUCS,' +
                 'PET_TELDUCS,PET_FAXDUCS FROM EMETTEURSOCIAL' +
                 ' where PET_EMETTSOC="' + DucsDeclarant.Emetteur + '"', True);
    if not Q.EOF then
    begin
      DucsDeclarant.RaisonSoc := copy(Q.FindField('PET_RAISONSOC').AsString, 1, 35);
      DucsDeclarant.ContactDecla := Q.FindField('PET_CONTACTDUCS').AsString;
      DucsDeclarant.TelDeclarant := Q.FindField('PET_TELDUCS').AsString;
      DucsDeclarant.FaxDeclarant := Q.FindField('PET_FAXDUCS').AsString;
    end;
    ferme(Q);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE MF
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : Création de la TOB TDucs en fonctions des critères de 
Suite ........ : sélection
Mots clefs ... : PAIE; DUCS
*****************************************************************}
procedure ChargeDucsAfaire( PeriodEtat, NatDeb, NatFin, EtabDeb,EtabFin : string;var TDucs : TOB);
var
  st: string;
  QRechDucs: TQuery;

begin
  st := ' ';
  st := 'SELECT POG_AUTPERCALCUL,POG_AUTREPERIODUCS,POG_BASETYPARR,' +
        'POG_CAISSEDESTIN,POG_DUCEXIGIBILITE,POG_DUCLIMITEDEPOT,' +
        'POG_DUCSDOSSIER,POG_DUCSREGLEMENT,POG_ETABLISSEMENT,POG_INSTITUTION,' +
        'POG_LONGEDITABLE,POG_LONGTOTAL,POG_LONGTOTALE,POG_MTTYPARR,' +
        'POG_NATUREORG,POG_NUMINTERNE,POG_ORGANISME,POG_PERIODCALCUL,' +
        'POG_PERIODICITDUCS,POG_POSDEBUT,POG_POSTOTAL,POG_REGROUPEMENT,' +
        'POG_RUPTAPE,POG_RUPTGROUPE,POG_RUPTNUMERO,POG_RUPTSIRET,'+
        'POG_SOUSTOTDUCS,POG_CONDSPEC,POG_BASETYPARR,POG_MTTYPARR ' +
        'FROM ORGANISMEPAIE WHERE ';

  // Sélection d'établissements
  if (EtabDeb <> '') and (EtabFin <> '') then
    st := st + '(POG_ETABLISSEMENT >="' + EtabDeb + '" ' +
               'AND POG_ETABLISSEMENT <="' + EtabFin + '")';

  // Sélection de nature de DUCS
  if (NatDeb <> '') and (NatFin <> '') then
  begin
    if (EtabDeb <> '') and (EtabFin <> '') then
      st := st + ' AND ';
    st := st + '(POG_NATUREORG >= "' + NatDeb + '" ' +
      'AND POG_NATUREORG <="' + NatFin + '")';
  end;

  if ((EtabDeb <> '') and (EtabFin <> '')) or
    ((NatDeb <> '') and (NatFin <> '')) then
    st := st + ' AND ';

  // sélection des organismes
  // - pour lesquels la périocicité correspond
  // - qui sont "caisse  destinataire"
  // - qui ne sont pas "ducs dossier" et "caisse destinataire" non cochée
  st := st + '((POG_PERIODICITDUCS="' + PeriodEtat + '") OR ' +
             '((POG_AUTREPERIODUCS<>"") AND (POG_AUTREPERIODUCS="' +
             PeriodEtat + '"))) ' +
             'AND' +
             '((POG_CAISSEDESTIN="X") OR ' +
             '((POG_CAISSEDESTIN<>"X") AND (POG_DUCSDOSSIER<>"X")))';

  QRechDucs := OpenSql(st, TRUE);
  TDucs := TOB.Create('Les DUCS', nil, -1);
  TDucs.LoadDetailDB('ORGANISMEPAIE', '', '', QRechDucs, False);
  Ferme(QRechDucs);
end; {fin ChargeDucsAfaire}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE MF
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : On établit la TOB  TInst des organismes ayant même code 
Suite ........ : de regroupement
Mots clefs ... : PAIE; DUCS
*****************************************************************}
procedure ChargeListeInstit(DucsParametres : TDucsParametres; var Tinst : TOB);
var
  st: string;
  QRechInstit: TQuery;
begin
  st := ' ';
  if (DucsParametres.Regroupt <> '') then
  begin
    if (DucsParametres.DucsDoss = TRUE) then
    {Ducs Dossier}
    begin
      st := 'SELECT POG_ETABLISSEMENT, POG_ORGANISME,POG_INSTITUTION, ' +
        'POG_NUMINTERNE ' +
        'FROM ORGANISMEPAIE' +
        ' WHERE POG_REGROUPEMENT ="' + DucsParametres.Regroupt + '"' +
        ' AND POG_NATUREORG ="' + DucsParametres.NatureDucs + '"';
    end
    else
    { Ducs Etablissement }
    begin
      st := 'SELECT POG_ETABLISSEMENT, POG_ORGANISME,POG_INSTITUTION, ' +
        'POG_NUMINTERNE ' +
        'FROM ORGANISMEPAIE' +
        ' WHERE POG_REGROUPEMENT="' + DucsParametres.Regroupt + '"' +
        ' AND POG_ETABLISSEMENT="' + DucsParametres.Etab + '"' +
        ' AND POG_NATUREORG ="' + DucsParametres.NatureDucs + '"';
    end;
  end
  else
  begin
    st := 'SELECT POG_ETABLISSEMENT, POG_ORGANISME, POG_INSTITUTION, ' +
      'POG_NUMINTERNE ' +
      'FROM ORGANISMEPAIE' +
      ' WHERE POG_ORGANISME="' + DucsParametres.Organisme + '"' +
      ' AND POG_ETABLISSEMENT="' + DucsParametres.Etab + '"';
  end;

  QRechInstit := OpenSql(st, TRUE);
  TInst := TOB.Create('Les institutions', nil, -1);
  TInst.LoadDetailDB('TABLEORGANISMEPAIE', '', '', QRechInstit, False);
  Ferme(QRechInstit);
end; {fin ChargeListeInstit}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE MF
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : Initialisation des variables de traitement
Mots clefs ... : PAIE; DUCS
*****************************************************************}
procedure InitVarTrait(DucsParametres : TDucsParametres; Debper, FinPer : string; var DucsVariables : TDucsVariables;var DucsEffectifs : TDucsEffectifs;TDucs, Tinst : TOB; var TRupMere : TOB);
var
  QQ                                  : Tquery;
  StQQ                                : string;
  TD                                  : TOB;

begin
  { Traitement des ruptures}
  DucsVariables.WSiret := '';
  DucsVariables.WApe := '';
  DucsVariables.WNumInt := '';
  DucsVariables.WGpInt := '';

  StQQ := '';

  DucsVariables.InstEnCours := '';

  DucsEffectifs.EffTot := 0;
  DucsEffectifs.EffHom := 0;
  DucsEffectifs.EffFem := 0;
  DucsEffectifs.EffApp := 0;
  DucsEffectifs.EffAppH  := 0;
  DucsEffectifs.EffAppF  := 0;
  DucsEffectifs.EffCad := 0;
  DucsEffectifs.EffCadh := 0;
  DucsEffectifs.EffCadF := 0;
  DucsEffectifs.EffEtam := 0;
  DucsEffectifs.EffOuv := 0;

  {Récupératon des paramètres Organisme}
  TD := TDucs.FindFirst(['POG_ETABLISSEMENT', 'POG_ORGANISME'],
    [DucsParametres.Etab, DucsParametres.Organisme], TRUE);

  {Paramètres d'édition de la codification}
  DucsVariables.CodifPosDebut := TD.GetValue('POG_POSDEBUT');
  DucsVariables.CodifLongEdit := TD.GetValue('POG_LONGEDITABLE');

  {Booléen condition spéciale de cotisation alimentation auto O/N}
  if (TD.GetValue('POG_CONDSPEC') =  'X') then
     DucsVariables.CondSpec :=True
  else
     DucsVariables.CondSpec := False;

  {Paramètres pour le traitement des ruptures}
  DucsVariables.NoMere := 0;
  DucsVariables.RupSiret := FALSE;
  DucsVariables.RupApe := FALSE;
  DucsVariables.RupNumInt := FALSE;
  DucsVariables.RupGpInt := FALSE;
  DucsVariables.Rupture := FALSE;

  if (TD.GetValue('POG_RUPTGROUPE') = 'X') then
  begin
    DucsVariables.RupGpint := TRUE;
    if (DucsVariables.NoMere = 0) then DucsVariables.NoMere := 1;
    DucsVariables.Rupture := TRUE;
  end;

  {Quand Mono établissement on ne traite pas les ruptures Siret, Ape, N° interne}
  if (TD.GetValue('POG_DUCSDOSSIER') = 'X') then
  begin
    if (TD.GetValue('POG_RUPTAPE') = 'X') then
    begin
      DucsVariables.RupApe := TRUE;
      if (DucsVariables.NoMere = 0) then DucsVariables.NoMere := 3;
      DucsVariables.Rupture := TRUE;
    end;

    if (TD.GetValue('POG_RUPTSIRET') = 'X') then
    begin
      DucsVariables.RupSiret := TRUE;
      if (DucsVariables.NoMere = 0) then DucsVariables.NoMere := 4;
      DucsVariables.Rupture := TRUE;
    end;
  end;

  if (TD.GetValue('POG_RUPTNUMERO') = 'X') then
  begin
    DucsVariables.RupNumInt := TRUE;
    if (DucsVariables.NoMere = 0) then DucsVariables.NoMere := 2;
    DucsVariables.Rupture := TRUE;
  end;

  { Paramètres d'arrondi des bases et montants}
  DucsVariables.BaseArr := TD.GetValue('POG_BASETYPARR');
  DucsVariables.MontArr := TD.GetValue('POG_MTTYPARR');
  DucsVariables.WBaseArr := DucsVariables.BaseArr;
  DucsVariables.WMontArr := DucsVariables.MontArr;

  {paramètres de calcul des dates}
  DucsVariables.JExigible  := TD.GetValue('POG_DUCEXIGIBILITE');
  DucsVariables.JLimDepot := TD.GetValue('POG_DUCLIMITEDEPOT');
  DucsVariables.JReglement := TD.GetValue('POG_DUCSREGLEMENT');

  {On établit les listes des valeurs possibles pour chaque critère de rupture}
  if (DucsVariables.Rupture = TRUE) then
    RechRupture(DucsParametres,DucsVariables, TInst,TRupMere);


  StQQ := '';
  StQQ := '((PPU_ETABLISSEMENT ="' + DucsParametres.Etab + '") AND ' +
    '(PPU_DATEDEBUT >="' + UsDateTime(StrToDate(DebPer)) + '") AND ' +
    '(PPU_DATEFIN <= "' + UsDateTime(StrToDate(FinPer)) + '")) ';
  QQ := OpenSQL('SELECT MAX(PPU_PAYELE) FROM PAIEENCOURS WHERE ' +
    StQQ, True);
  if not QQ.EOF then
  begin
    DucsVariables.DatePaye := QQ.Fields[0].AsDateTime;
  end;
  Ferme(QQ);
  StQQ := '';

end; {fin InitVarTrait}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE MF
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : Création des différentes TOB utilisées pour gérer les 
Suite ........ : ruptures
Mots clefs ... : PAIE; DUCS
*****************************************************************}
procedure RechRupture(DucsParametres : TDucsparametres; DucsVariables : TDucsVariables; TInst :TOB; var TRupMere : TOB );
var
  st, Organ                             : string;
  QRech                                 : TQuery;
  TobExiste, LesInst, TRupFille         : TOB;
  NbInst, i                             : Integer;
  NomChamp                              : array[1..3] of string;
  ValeurChamp                           : array[1..3] of variant;

begin
  TRupFille := nil;
  
  st := ' ';
  {Création de la TOB utilisée en fonction du paramètrage des ruptures}
  case DucsVariables.NoMere of
    1: begin
         { TOB - groupe interne}
        st := 'SELECT distinct(PGI_GROUPE),PGI_INSTITUTION, PSA_DADSPROF, ' +
          'ET_ETABLISSEMENT, ET_SIRET, ET_APE, POG_ORGANISME, ' +
          'POG_NUMINTERNE,' +
          'POG_REGROUPEMENT,POG_INSTITUTION, ' +
          'POG_CAISSEDESTIN, '+
          'POG_LGOPTIQUE,POG_NATUREORG ' +
          ',ETB_REGIMEALSACE '+
          'FROM HISTOBULLETIN ' +
          'LEFT JOIN ETABLISS ON ET_ETABLISSEMENT=PHB_ETABLISSEMENT ' +
          'LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT=PHB_ETABLISSEMENT ' +
          'LEFT JOIN SALARIES ON PSA_SALARIE=HISTOBULLETIN.PHB_SALARIE ' +
          'LEFT JOIN GROUPEINTERNE ON PGI_CATEGORIECRC=SALARIES.PSA_DADSPROF ' +
          'LEFT JOIN ORGANISMEPAIE ON POG_INSTITUTION=PGI_INSTITUTION ' +
          'AND POG_ORGANISME=PHB_ORGANISME AND POG_ETABLISSEMENT=PHB_ETABLISSEMENT ' +
          'WHERE PSA_DADSPROF<>"" AND ';

        if (DucsParametres.Regroupt <> '') then { Ducs regroupement d'organisme}
        begin
          for NbInst := 0 to TInst.Detail.Count - 1 do
          begin
            LesInst := TInst.Detail[NbInst];
            Organ := LesInst.GetValue('POG_ORGANISME');
            if (NbInst <> 0) then
            begin
              if (Pos('POG_ORGANISME="' + Organ + '"', st) = 0) then
                st := st + ' OR '
            end
            else
              st := st + '(';

            if (Pos('POG_ORGANISME="' + Organ + '"', st) = 0) then
              st := st + 'POG_ORGANISME="' + Organ + '"';
          end;
          st := st + ')';
        end
        else
          st := st + 'POG_ORGANISME="' + DucsParametres.Organisme + '" ';

        st := st + ' AND POG_RUPTGROUPE= "X"';
        if (DucsParametres.DucsDoss = TRUE) then
          st := st + ' AND POG_DUCSDOSSIER="X" '
        else
          st := st + ' AND ET_ETABLISSEMENT="' + DucsParametres.Etab + '"';

        st := st + ' ORDER BY ';
         {order by dans le même ordre que pour TCot2}
        if (DucsVariables.RupSiret = True) then
          st := st + 'ET_SIRET,';
        if (DucsVariables.RupApe = True) then
          st := st + 'ET_APE,';
        if (DucsVariables.RupNumInt = true) then
          st := st + 'POG_NUMINTERNE,';
        if (DucsVariables.RupSiret = False) then
          st := st + ' PGI_GROUPE, ET_SIRET'
        else
          st := st + ' PGI_GROUPE';

        QRech := OpenSql(st, TRUE);
      TRupMere := TOB.Create('Les groupes internes', nil, -1);
      end;

    2: begin
         { TOB - N° interne}
        st := 'SELECT DISTINCT(POG_NUMINTERNE), ET_ETABLISSEMENT, ET_SIRET, ' +
          'ET_APE, POG_ORGANISME,' +
          'POG_REGROUPEMENT,POG_INSTITUTION, ' +
          'POG_CAISSEDESTIN, '+
          'POG_LGOPTIQUE,POG_NATUREORG ' +
          ',ETB_REGIMEALSACE '+
          'FROM ETABLISS ' +
          'LEFT JOIN ORGANISMEPAIE ' +
          'ON ET_ETABLISSEMENT=ORGANISMEPAIE.POG_ETABLISSEMENT ' +
          'LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT = ET_ETABLISSEMENT '+
          'WHERE ';

        if (DucsParametres.Regroupt <> '') then
         { Ducs regroupement d'organisme}
        begin
          for NbInst := 0 to TInst.Detail.Count - 1 do
          begin
            LesInst := TInst.Detail[NbInst];
            Organ := LesInst.GetValue('POG_ORGANISME');
            if (NbInst <> 0) then
            begin
              if (Pos('POG_ORGANISME="' + Organ + '"', st) = 0) then
                st := st + ' OR '
            end
            else
              st := st + '(';
            if (Pos('POG_ORGANISME="' + Organ + '"', st) = 0) then
              st := st + 'POG_ORGANISME="' + Organ + '"';
          end;
          st := st + ')';
        end
        else
          st := st + 'POG_ORGANISME="' + DucsParametres.Organisme + '" ';

        st := st + ' AND POG_RUPTNUMERO= "X" AND POG_NUMINTERNE <> ""';

        if (DucsParametres.DucsDoss = TRUE) then
          st := st + ' AND POG_DUCSDOSSIER="X" '
        else
          st := st + ' AND POG_ETABLISSEMENT="' + DucsParametres.Etab + '"';

        if (DucsVariables.RupSiret = True) then
          st := st + 'ET_SIRET,';
        if (DucsVariables.RupApe = True) then
          st := st + 'ET_APE,';

        if (DucsVariables.RupSiret = False) then
          st := st + ' ORDER BY POG_NUMINTERNE, ET_SIRET'
        else
          st := st + ' ORDER BY POG_NUMINTERNE';

        QRech := OpenSql(st, TRUE);
        TRupMere := TOB.Create('Les numéros internes', nil, -1);
      end;

    3: begin
         {TOB - Code Ape}
        st := 'SELECT DISTINCT(ET_APE), ET_ETABLISSEMENT, ET_SIRET, POG_ORGANISME, POG_NUMINTERNE,' +
          'POG_REGROUPEMENT,POG_INSTITUTION, ' +
          'POG_CAISSEDESTIN, '+
          'POG_LGOPTIQUE,POG_NATUREORG ' +
          ',ETB_REGIMEALSACE '+ 
          'FROM ETABLISS ' +
          'LEFT JOIN ORGANISMEPAIE ' +
          'ON ET_ETABLISSEMENT=ORGANISMEPAIE.POG_ETABLISSEMENT ' +
          'LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT = ET_ETABLISSEMENT  ' +  
          ' WHERE ';

        if (DucsParametres.Regroupt <> '') then
         {Ducs regroupement d'organisme}
        begin
          for NbInst := 0 to TInst.Detail.Count - 1 do
          begin
            LesInst := TInst.Detail[NbInst];
            Organ := LesInst.GetValue('POG_ORGANISME');

            if (NbInst <> 0) then
            begin
              if (Pos('POG_ORGANISME="' + Organ + '"', st) = 0) then
                st := st + ' OR ';
            end
            else
              st := st + '(';

            if (Pos('POG_ORGANISME="' + Organ + '"', st) = 0) then
              st := st + 'POG_ORGANISME="' + Organ + '"';
          end;
          st := st + ')';
        end
        else
          st := st + 'POG_ORGANISME ="' + DucsParametres.Organisme + '"';

        st := st + ' AND POG_DUCSDOSSIER="X" AND POG_RUPTAPE= "X" ' +
          ' ORDER BY ET_APE';
        QRech := OpenSql(st, TRUE);
        TRupMere := TOB.Create('Les ape', nil, -1);
      end;

    4: begin
         {TOB - N° Siret}
        st := 'SELECT DISTINCT(ET_SIRET),ET_ETABLISSEMENT, ET_APE, ' +
          'POG_ORGANISME, POG_NUMINTERNE,' +
          'POG_REGROUPEMENT,POG_INSTITUTION, ' +
          'POG_CAISSEDESTIN, '+
          'POG_LGOPTIQUE,POG_NATUREORG ' +
          ',ETB_REGIMEALSACE '+
          'FROM ETABLISS ' +
          'LEFT JOIN ORGANISMEPAIE ' +
          'ON ET_ETABLISSEMENT=ORGANISMEPAIE.POG_ETABLISSEMENT ' +
          'LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT=ET_ETABLISSEMENT ' +
          'WHERE '; // PT15-1

        if (DucsParametres.Regroupt <> '') then
         {Ducs regroupement d'organisme}
        begin
          for NbInst := 0 to TInst.Detail.Count - 1 do
          begin
            LesInst := Tinst.Detail[NbInst];
            Organ := LesInst.GetValue('POG_ORGANISME');
            if (NbInst <> 0) then
            begin
              if (Pos('POG_ORGANISME="' + Organ + '"', st) = 0) then
                st := st + ' OR ';
            end
            else
              st := st + '(';

            if (Pos('POG_ORGANISME="' + Organ + '"', st) = 0) then
              st := st + 'POG_ORGANISME="' + Organ + '"';
          end;
          st := st + ')';
        end
        else
          st := st + 'POG_ORGANISME ="' + DucsParametres.Organisme + '"';

        st := st + ' AND POG_DUCSDOSSIER="X" AND POG_RUPTSIRET= "X"' +
          ' ORDER BY ET_SIRET';

        QRech := OpenSql(st, TRUE);
        TRupMere := TOB.Create('Les siret', nil, -1);
      end;
  end;

  {Recherche si élément pas déjà en TOB}
  NomChamp[1] := '';
  ValeurChamp[1] := '';
  NomChamp[2] := '';
  ValeurChamp[2] := '';
  NomChamp[3] := '';
  ValeurChamp[3] := '';

  while not QRech.EOF do
  begin
    if (DucsVariables.RupGpInt = TRUE) then
    begin
      NomChamp[1] := 'PGI_GROUPE';
      ValeurChamp[1] := QRech.FindField('PGI_GROUPE').AsString;
    end;
    if (DucsVariables.RupNumInt = TRUE) then
    begin
      if (NomChamp[1] = '') then
      begin
        NomChamp[1] := 'POG_NUMINTERNE';
        ValeurChamp[1] := QRech.FindField('POG_NUMINTERNE').AsString;
      end
      else
      begin
        NomChamp[2] := 'POG_NUMINTERNE';
        ValeurChamp[2] := QRech.FindField('POG_NUMINTERNE').AsString;
      end;
    end;
    if (DucsVariables.RupApe = TRUE) then
    begin
      if (NomChamp[2] = '') then
      begin
        if (NomChamp[1] = '') then
        begin
          NomChamp[1] := 'ET_APE';
          ValeurChamp[1] := QRech.FindField('ET_APE').AsString;
        end
        else
        begin
          NomChamp[2] := 'ET_APE';
          ValeurChamp[2] := QRech.FindField('ET_APE').AsString;
        end;
      end
      else
      begin
        NomChamp[3] := 'ET_APE';
        ValeurChamp[3] := QRech.FindField('ET_APE').AsString;
      end;
    end;

    if (DucsVariables.RupSiret = TRUE) then
    begin
      if (NomChamp[2] = '') then
      begin
        if (NomChamp[1] = '') then
        begin
          NomChamp[1] := 'ET_SIRET';
          ValeurChamp[1] := QRech.FindField('ET_SIRET').AsString;
        end
        else
        begin
          NomChamp[2] := 'ET_SIRET';
          ValeurChamp[2] := QRech.FindField('ET_SIRET').AsString;
        end;
      end
      else
      begin
        NomChamp[3] := 'ET_SIRET';
        ValeurChamp[3] := QRech.FindField('ET_SIRET').AsString;
      end;
    end;

    TobExiste := TRupMere.FindFirst(NomChamp, ValeurChamp, TRUE);
    if (TobExiste = nil) then
    {Pas déjà en TOB - Création d'une TOB fille}
    begin
      TRupFille := TOB.Create('TABLEETABLISS', TRupMere, -1);
      TRupFille.InitValeurs;
      if (DucsVariables.RupGpInt = true) then
      begin
        TRupFille.AddChampSup('PGI_GROUPE', FALSE);
        TRupFille.PutValue('PGI_GROUPE', QRech.FindField('PGI_GROUPE').AsString);
      end;
      TRupFille.AddChampSup('POG_NUMINTERNE', FALSE);
      TRupFille.PutValue('POG_NUMINTERNE', QRech.FindField('POG_NUMINTERNE').AsString);
      TRupFille.AddChampSup('ET_APE', FALSE);
      TRupFille.PutValue('ET_APE', QRech.FindField('ET_APE').AsString);
      TRupFille.AddChampSup('POG_ETABLISSEMENT', FALSE);
      TRupFille.PutValue('POG_ETABLISSEMENT', QRech.FindField('ET_ETABLISSEMENT').AsString);
      TRupFille.AddChampSup('ET_SIRET', FALSE);
      TRupFille.PutValue('ET_SIRET', QRech.FindField('ET_SIRET').AsString);
      TRupFille.AddChampSup('POG_ORGANISME', FALSE);
      TRupFille.PutValue('POG_ORGANISME', QRech.FindField('POG_ORGANISME').AsString);

      {Formatage ligne optique (elle peut être déjà renseignée au niveau
       du paramètrage de l'organisme}
      DucsVariables.LigneOptique := '';
      if (QRech.FindField('POG_LGOPTIQUE').AsString = '') then
      begin
        if (QRech.FindField('POG_NATUREORG').AsString = '100') then
        {URSSAF}
        begin
          DucsVariables.LigneOptique := Copy(QRech.FindField('POG_NUMINTERNE').AsString, 1,
            length(QRech.FindField('POG_NUMINTERNE').AsString));
          if (length(QRech.FindField('POG_NUMINTERNE').AsString) < 30) then
          begin
            for i := length(QRech.FindField('POG_NUMINTERNE').AsString) + 1 to 30 do
            begin
              DucsVariables.LigneOptique := DucsVariables.LigneOptique + '0';
            end;
          end;
        end;
        if (QRech.FindField('POG_NATUREORG').AsString = '200') then
        {ASSEDIC}
        begin
          DucsVariables.LigneOptique := 'S2' + Copy(QRech.FindField('POG_NUMINTERNE').AsString, 1,
            length(QRech.FindField('POG_NUMINTERNE').AsString));
          if ((length(QRech.FindField('POG_NUMINTERNE').AsString) + 2) < 30) then
          begin
            for i := length(QRech.FindField('POG_NUMINTERNE').AsString) + 3 to 30 do
            begin
              DucsVariables.LigneOptique := DucsVariables.LigneOptique + '0';
            end;
          end;
        end;
      end
      else
      begin
        DucsVariables.LigneOptique := QRech.FindField('POG_LGOPTIQUE').AsString;
        if (length(QRech.FindField('POG_LGOPTIQUE').AsString) < 30) then
        begin
          for i := length(QRech.FindField('POG_LGOPTIQUE').AsString) + 1 to 30 do
          begin
            DucsVariables.LigneOptique := DucsVariables.LigneOptique + '0';
          end;
        end;
      end;

      TRupFille.AddChampSup('POG_LGOPTIQUE', FALSE);
      TRupFille.PutValue('POG_LGOPTIQUE', DucsVariables.LigneOptique);
    end {fin if (TobExiste = NIL)}
    else
    begin
      // TobExiste <> NIL : élément déjà existant en TOB (cas pls Gps Internes avec
      // même code et ducs dossier) on récupère les infos de la caisse destinataire)
      if (QRech.FindField('POG_CAISSEDESTIN').AsString = 'X') then
      begin
        TRupFille.PutValue('POG_NUMINTERNE', QRech.FindField('POG_NUMINTERNE').AsString);
        TRupFille.PutValue('ET_APE', QRech.FindField('ET_APE').AsString);
        TRupFille.PutValue('POG_ETABLISSEMENT', QRech.FindField('ET_ETABLISSEMENT').AsString);
        TRupFille.PutValue('ET_SIRET', QRech.FindField('ET_SIRET').AsString);
        TRupFille.PutValue('POG_ORGANISME', QRech.FindField('POG_ORGANISME').AsString);
      end;
    end;
    TRupFille.AddChampSup('ETB_REGIMEALSACE', FALSE);
    TRupFille.PutValue('ETB_REGIMEALSACE', QRech.FindField('ETB_REGIMEALSACE').AsString);
    QRech.Next;

    NomChamp[1] := '';
    ValeurChamp[1] := '';
    NomChamp[2] := '';
    ValeurChamp[2] := '';
    NomChamp[3] := '';
    ValeurChamp[3] := '';
  end; {fin while not QRech.EOF}

//  TRupMere.SaveToFile('c:\tmp\TrupMere', False, TRUE, TRUE);

  Ferme(QRech);
end; { fin RechRupture}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE MF
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : ChargeCot : Chargement de la TOB TCot3 des cotisations
Suite ........ : regroupées par codification , par tauxat, par taux global et
Suite ........ : pour chaque salarié par base identique pour une période
Mots clefs ... : PAIE; DUCS
*****************************************************************}
procedure ChargeCot(DucsParametres : TDucsParametres; DucsVariables : TDucsVariables; DebPer, FinPer : string; TInst : TOB; var TCot3 : TOB; var ErrRupGpInt: integer;var  TraceE :TStringList; IndTraceE : boolean);
var
  QRechCot: TQuery;
  LesInst, TDetCot, TCot2Fille, TCot, TDetCod: TOB;
  StOrder, StOrder2, StOrder3, StOrder4, WsiretEnCours: string;
  WRubLu, WRubEnCours, WSalLu, WSalEnCours, WsiretLu: string;
  WNoIntLu, WNoIntEnCours, WApeLu, WApeEnCours, WGpLu: string;
  WCodEnCours, WCodLu, WOrgEnCours, WOrgLu, WGpEnCours: string;
  WInstLu, WInstEnCours, st, Organ, WcodPrec: string;
  WTxATLu, WTxATEncours, WTxGlLu, WTxGlEnCours: double;
  WBaseEnCours, WBaselu, TxGlobal, MtGlobal: double;
  WDateLu, WDateEnCours: TDateTime;
  WDateLuFin, WDateFinEnCours: TDateTime;
  IndFree, Personnalisee, WPrecExiste: Boolean;
  I, NbInst: integer;
  WMtGlobalLu, WMtGlobalEnCours : double;
  IndicD, IndicP, NbCodif                : integer;
  TauxglobalEnCours, TauxGlobalLu        : double;
  CodifW , SalW                          : string;
  DateDebW                               : TdateTime;
  TauxAtW                                : double;
  TCotA, TCotAFille,TCot3Fille           : TOB;
  TCot2Bis,TCot2BisFille                 : TOB;
  DateFinW                               : TdateTime;
  TCot2                                  : TOB;
  Tcot3FilleBis                          : tob;
  WBaseMaj                               : double;
  IndCotRegul                            : boolean;
  RegulNegative                          : boolean;
  SalEnCours,SalLu,WCodifEnCours         : string;
  BaseEnCours, Baselue,TauxAtEnCours     : double;
  TCotABis, TCotABisFille,TCotABisFilleP : TOB;
  INDIC                                  : integer;

begin
  WcodPrec := '';
  WPrecExiste := false;

  Personnalisee := False;

  {CHARGEMENT DE LA REQUETE DE SELECTION DES COTISATIONS A TRAITER}
  st := ' ';
  st := 'SELECT PHB_SALARIE,PDF_CODIFICATION,PDF_CODIFALSACE,PHB_RUBRIQUE,PDF_PREDEFINI,' +
         'PHB_COTREGUL,' +
        'PDF_NODOSSIER,PHB_BASECOT,PHB_TAUXAT,' +
        'PHB_DATEFIN, ' +
        'PHB_DATEDEBUT,PHB_ORGANISME, ';

  st := st + 'POG_POSTOTAL, POG_LONGTOTAL, ';
  st := st + 'POG_INSTITUTION, ';
  st := st + 'ETB_REGIMEALSACE, ';

  {traitement des ruptures : il faut récupérer les champs sur lesquels une
   rupture est possible}
  if (DucsVariables.RupSiret = TRUE) then
    st := st + 'ET_SIRET, ';
  if (DucsVariables.RupApe = TRUE) then
    st := st + 'ET_APE, ';
  if (DucsVariables.RupNumInt = TRUE) then
    st := st + 'POG_NUMINTERNE, ';
  if (DucsVariables.RupGpInt = TRUE) then
    st := st + 'PGI_GROUPE, ';

  {Il est nécessaire de récuperer le n° d'institution pour pouvoir gérer
   les sous-totaux par institution}
  st := st + 'SUM(DISTINCT(PHB_TAUXSALARIAL+PHB_TAUXPATRONAL)) AS TAUXGLOBAL, ' +
             'SUM(DISTINCT(PHB_MTSALARIAL+PHB_MTPATRONAL)) AS MTGLOBAL ' +
             'FROM HISTOBULLETIN ' +
              'LEFT JOIN DUCSAFFECT ON PDF_RUBRIQUE=HISTOBULLETIN.PHB_RUBRIQUE ';

  { Traitement ruptures Siret ou APE:}
  if (DucsVariables.RupSiret = TRUE) or (DucsVariables.RupApe = TRUE) then
    st := st + ' LEFT JOIN ETABLISS ON ' +
               'ET_ETABLISSEMENT=HISTOBULLETIN.PHB_ETABLISSEMENT ';

  st := st + ' LEFT JOIN ETABCOMPL ON ' +
             'ETB_ETABLISSEMENT=HISTOBULLETIN.PHB_ETABLISSEMENT ';
  st := st + ' LEFT JOIN ORGANISMEPAIE ON ' +
             'POG_ORGANISME=HISTOBULLETIN.PHB_ORGANISME ' +
             ' AND POG_ETABLISSEMENT=HISTOBULLETIN.PHB_ETABLISSEMENT ';

  {Traitement rupture sur Groupe interne}
  if (DucsVariables.RupGpInt = TRUE) then
  begin
    st := st + ' LEFT JOIN SALARIES ON PSA_SALARIE= HISTOBULLETIN.PHB_SALARIE ';
    st := st + ' LEFT JOIN GROUPEINTERNE ON PGI_CATEGORIECRC = ' +
      'SALARIES.PSA_DADSPROF ' +
      'AND PGI_INSTITUTION=ORGANISMEPAIE.POG_INSTITUTION ';
  end;

    st := st + ' WHERE ##PDF_PREDEFINI## AND PHB_NATURERUB="COT" AND ' +
               'PHB_DATEFIN >="' + UsDateTime(StrToDATe(DebPer)) + '" AND ' +
               'PHB_DATEFIN<="' + UsDateTime(StrToDATe(FinPer)) + '" AND  ' +
               '((ETB_REGIMEALSACE = "X" AND  '+
               'PDF_CODIFALSACE IS NOT NULL AND PDF_CODIFALSACE <> "" AND ' +
               'PDF_CODIFALSACE <> "       " ) OR ' +
               '(ETB_REGIMEALSACE <> "X" AND  '+
               'PDF_CODIFICATION IS NOT NULL AND PDF_CODIFICATION <> "" AND ' +
               'PDF_CODIFICATION <> "       ")) AND ';

  if (DucsParametres.DucsDoss <> TRUE) then
  {Ducs établissement}
    st := st + 'PHB_ETABLISSEMENT="' + DucsParametres.Etab + '" AND '

  else
  {Ducs dossier
   exclure les établissements qui ne font pas partie de la Ducs Dossier}
    st := st + ' POG_DUCSDOSSIER="X" AND ';

  if (DucsParametres.Regroupt <> '') then
  {Ducs regroupement d'organisme, il est nécessaire de récupérer les
   cotisations des différents caisses participant au groupe}
  begin
    for NbInst := 0 to TInst.Detail.Count - 1 do
    begin
      LesInst := TInst.Detail[NbInst];
      Organ := LesInst.GetValue('POG_ORGANISME');
      if (NbInst <> 0) then
      begin
        if (Pos('PHB_ORGANISME="' + Organ + '"', st) = 0) then
          st := st + ' OR '
      end
      else
        st := st + '(';

      if (Pos('PHB_ORGANISME="' + Organ + '"', st) = 0) then
        st := st + 'PHB_ORGANISME="' + Organ + '"';
    end;
    st := st + ')';
  end
  else
    st := st + 'PHB_ORGANISME="' + DucsParametres.Organisme + '" ';

   st := st + ' AND (PHB_BASECOT <> 0 OR PHB_MTSALARIAL <> 0 OR ' +
               'PHB_MTPATRONAL <> 0) ' +
               'GROUP BY PHB_BASECOT,PHB_TAUXAT,PHB_DATEDEBUT,PHB_DATEFIN,PHB_SALARIE,' +
               'PDF_CODIFICATION,PDF_CODIFALSACE,PHB_RUBRIQUE,PDF_PREDEFINI,PDF_NODOSSIER,' +
               'PHB_ORGANISME ';

  st := st + ',POG_INSTITUTION ';

  {traitement des ruptures : il faut regrouper les cotisations en fonctions des
   ruptures possibles}
  if (DucsVariables.RupSiret = TRUE) then
    st := st + ',ET_SIRET ';
  if (DucsVariables.RupApe = TRUE) then
    st := st + ',ET_APE ';
  if (DucsVariables.RupNumInt = TRUE) then
    st := st + ',POG_NUMINTERNE ';
  if (DucsVariables.RupGpInt = TRUE) then
    st := st + ',PGI_GROUPE ';

  st := st + ',POG_POSTOTAL, POG_LONGTOTAL ';
  st := st + ',ETB_REGIMEALSACE '; 
  st := st + ',PHB_COTREGUL ';  
  st := st + 'ORDER BY ';
  StOrder := '';
  StOrder4 := '';

  {traitement des ruptures}
  if (DucsVariables.RupNumInt = TRUE) then
  begin
    st := st + 'POG_NUMINTERNE,';
    StOrder := StOrder + 'POG_NUMINTERNE;';
    StOrder4 := StOrder4 + 'NUMINTERNE;';
  end;
  if (DucsVariables.RupSiret = TRUE) then
  begin
    st := st + 'ET_SIRET,';
    StOrder := StOrder + 'ET_SIRET;';
    StOrder4 := StOrder4 + 'SIRET;';
  end;
  if (DucsVariables.RupApe = TRUE) then
  begin
    st := st + 'ET_APE,';
    StOrder := StOrder + 'ET_APE;';
    StOrder4 := StOrder4 + 'APE;';
  end;
  if (DucsVariables.RupGpInt = TRUE) then
  begin
    st := st + 'PGI_GROUPE, ';
    StOrder := StOrder + 'PGI_GROUPE;';
    StOrder4 := StOrder4 + 'GROUPE;';
  end;

  StOrder2 := StOrder;
  StOrder3 := StOrder;

   st := st + 'PHB_ORGANISME,PDF_CODIFICATION,PDF_CODIFALSACE,PHB_RUBRIQUE,PDF_PREDEFINI,' +
              'PDF_NODOSSIER,PHB_TAUXAT,TAUXGLOBAL,PHB_SALARIE';

   StOrder := StOrder + 'PHB_ORGANISME;PDF_CODIFICATION;PDF_CODIFALSACE;PHB_RUBRIQUE;' +
                        'PDF_PREDEFINI;PDF_NODOSSIER;' +
                        'PHB_TAUXAT;TAUXGLOBAL;PHB_SALARIE';

  StOrder2 := StOrder2 + 'PHB_ORGANISME;' +
                         'PHB_TAUXAT;TAUXGLOBAL;MTGLOBAL;PHB_SALARIE;' +
                         'PHB_DATEDEBUT;PHB_DATEFIN;PRED;PDF_CODIFICATION;PDF_CODIFALSACE';
    StOrder3 := StOrder3 + 'PHB_ORGANISME;PDF_CODIFICATION;PDF_CODIFALSACE;PHB_TAUXAT;PHB_SALARIE;' +
                            'PHB_RUBRIQUE;PHB_BASECOT;PHB_DATEDEBUT';
  StOrder4 := StOrder4 + 'ORGANISME';

  if (DucsParametres.Regroupt <> '') then
  begin
    st := st + ',POG_INSTITUTION';
    StOrder := StOrder + ';POG_INSTITUTION';
    StOrder2 := StOrder2 + ';POG_INSTITUTION';
    StOrder3 := StOrder3 + ';POG_INSTITUTION';
    StOrder4 := StOrder4 + ';INSTITUTION';
  end;

  StOrder4 := StOrder4 + ';CODIFICATION;TAUXAT;SALARIE;TAUXGLOBAL;' +	
    'BASECOT;DATEDEBUT';

  QRechCot := OpenSql(st, TRUE);
  TCot := TOB.Create('Les Cotisations', nil, -1);
  TCot.LoadDetailDB('TABHISTOBULLETIN', '', '', QRechCot, False);
  Ferme(QRechCot);

  if (DucsVariables.RupGpInt = TRUE) then
  begin
    ErrRupGpInt := 0;
    for I := 0 to TCot.Detail.Count - 1 do
    begin
      TDetCod := TCot.Detail[I];
      if (TDetCod.GetValue('POG_INSTITUTION') = '') or
        (TDetCod.GetValue('POG_INSTITUTION') = NULL) then
      begin
        ErrRupGpInt := 1;
        break;
      end;
      if (TDetCod.GetValue('PGI_GROUPE') = '') or
        (TDetCod.GetValue('PGI_GROUPE') = NULL) then
      begin
        ErrRupGpInt := 2;
        break;
      end;
    end;
    if (ErrRupGpInt <> 0) then
    begin
      if (ErrRupGpInt = 1) then
        PGIError('Organisme sans code institution.', 'Rupture sur groupe interne impossible');
      if (ErrRupGpInt = 2) then
        PGIError('Certains groupes internes ne sont pas créés.', 'Rupture sur groupe interne impossible');
      exit;
    end;
  end;

  {On épure les codifications non utiles au dossier}
  for I := 0 to TCot.Detail.Count - 1 do
  {On attribue une priorité aux affectations}
  begin
    TDetCod := TCot.Detail[I];
    TDetCod.AddChampSup('PRED', FALSE);
    if (TDetCod.GetValue('PDF_PREDEFINI') = 'DOS') then
      TDetCod.PutValue('PRED', 1);
    if (TDetCod.GetValue('PDF_PREDEFINI') = 'STD') then
      TDetCod.PutValue('PRED', 2);
    if (TDetCod.GetValue('PDF_PREDEFINI') = 'CEG') then
      TDetCod.PutValue('PRED', 3);
  end;

  TCot.Detail.Sort('PHB_RUBRIQUE;' + StOrder2);

  TDetCod := TCot.FindFirst([''], [''], TRUE);

  WRubEnCours := '';
  WSalEnCours := '';
  WsiretEnCours := '';
  WNoIntEnCours := '';
  WApeEnCours := '';
  WGpEnCours := '';
  WInstEnCours := '';
  WTxATEncours := 0.00;
  WTxGlEnCours := 0.00;
  WDateEnCours := IDate1900;
  WDateFinEnCours := IDate1900;
  WMtGlobalEncours := 0.00;
  for I := 0 to Tcot.Detail.Count - 1 do
  {On épure les affectations qui ne seront pas utilisées}
  begin
    WRubLu := TDetCod.GetValue('PHB_RUBRIQUE');
    WSalLu := TDetCod.GetValue('PHB_SALARIE');
    if (DucsVariables.RupGpInt = True) or
      ((DucsParametres.Regroupt <> '') and
      (DucsVariables.RupGpInt = False)) then
      if (TDetCod.GetValue('POG_INSTITUTION') <> Null) then
        WInstLu := TDetCod.GetValue('POG_INSTITUTION');
    if (DucsVariables.RupSiret = TRUE) then
      if (TDetCod.GetValue('ET_SIRET') <> Null) then
        WsiretLu := TDetCod.GetValue('ET_SIRET');
    if (DucsVariables.RupNumInt = TRUE) then
      if (TDetCod.GetValue('POG_NUMINTERNE') <> Null) then
        WNoIntLu := TDetCod.GetValue('POG_NUMINTERNE');
    if (DucsVariables.RupApe = TRUE) then
      if (TDetCod.GetValue('ET_APE') <> Null) then
        WApeLu := TDetCod.GetValue('ET_APE');
    if (DucsVariables.RupGpInt = TRUE) then
      if (TDetCod.GetValue('PGI_GROUPE') <> Null) then
        WGpLu := TDetCod.GetValue('PGI_GROUPE');
        
    WTxATLu := Arrondi(TDetCod.GetValue('PHB_TAUXAT'), 5);
    WTxGlLu := Arrondi(TDetCod.GetValue('TAUXGLOBAL'), 5);
    WDateLu := TDetCod.GetValue('PHB_DATEDEBUT');
    WDateLuFin := TDetCod.GetValue('PHB_DATEFIN');
    WMtGlobalLu := TDetCod.GetValue('MTGLOBAL');
    IndFree := TRUE;
    if (WRubLu <> WRubEnCours) then IndFree := False;
    if (WSalLu <> WSalEnCours) then IndFree := False;
    if (WInstLu <> WInstEnCours) then IndFree := False;
    if (WTxATLu <> WTxATEnCours) then IndFree := False;
    if (WTxGlLu <> WTxGlEnCours) then IndFree := False;
    if (WMtGlobalLu <> WMtGlobalEnCours) then IndFree := False;
    if (WDateLu <> WDateEnCours) then IndFree := False;
    if (WDateLuFin <> WDateFinEnCours) then IndFree := False;
    if (DucsVariables.RupSiret = TRUE) and (WSiretLU <> WsiretEnCours) then IndFree := False;
    if (DucsVariables.RupNumInt = TRUE) and (WNoIntLu <> WNoIntEnCours) then IndFree := False;
    if (DucsVariables.RupApe = TRUE) and (WApeLu <> WApeEnCours) then IndFree := False;
    if (DucsVariables.RupGpInt = TRUE) and (WGpLu <> WGpEnCours) then IndFree := False;

    if (IndFree = True) then
      TDetCod.Free;

    WRubEnCours := WRubLu;
    WSalEnCours := WSalLu;
    WInstEnCours := WInstLu;
    WsiretEnCours := WsiretLu;
    WNoIntEnCours := WNoIntLu;
    WApeEnCours := WApeLu;
    WGpEnCours := WGpLU;
    WTxATEncours := WTxATLu;
    WTxGlEnCours := WTxGlLu;
    WDateEnCours := WDateLu;
    WMtGlobalEnCours := WMtGlobalLu ;
    WDateFinEnCours := WDateLuFin;

    TDetCod := TCot.FindNext([''], [''], TRUE);
  end;

  {Création de la TOB Finale - cumul TAUXGLOBAL et MTGLOBAL
   par salarié, date de début de période, base, codification}
  TCot.Detail.Sort(StOrder3);
  TCot2 := TOB.Create('les codif', nil, -1);

  if (TCot.Detail.Count <> 0) then
    TDetCod := TCot.Detail[0];
  if (TDetCod <> nil) then
  begin
    WSalEnCours := TDetCod.GetValue('PHB_SALARIE');
    WRubEnCours := '';  // PT87 FQ15199

    if (TDetCod.GetValue('POG_INSTITUTION') <> Null) then
      WInstEnCours := TDetCod.GetValue('POG_INSTITUTION');

    WDateEnCours := TDetCod.GetValue('PHB_DATEDEBUT');
    WDateFinEnCours := TDetCod.GetValue('PHB_DATEFIN');
    if (TDetCod.GetValue('ETB_REGIMEALSACE')='X') then
      WCodEnCours := TDetCod.GetValue('PDF_CODIFALSACE')
    else
      WCodEnCours := TDetCod.GetValue('PDF_CODIFICATION');
    WBaseEnCours := Arrondi(TDetCod.GetValue('PHB_BASECOT'), 5);
    WTxATEnCours := Arrondi(TDetCod.GetValue('PHB_TAUXAT'), 5);
    WOrgEnCours := TDetCod.GetValue('PHB_ORGANISME');

    if (DucsVariables.RupSiret = True) then
      if (TDetCod.GetValue('ET_SIRET') <> Null) then
        WsiretEnCours := TDetCod.GetValue('ET_SIRET');
    if (DucsVariables.RupNumInt = True) then
      if (TDetCod.GetValue('POG_NUMINTERNE') <> Null) then
        WNoIntEnCours := TDetCod.GetValue('POG_NUMINTERNE');
    if (DucsVariables.RupApe = True) then
      if (TDetCod.GetValue('ET_APE') <> Null) then
        WApeEnCours := TDetCod.GetValue('ET_APE');
    if (DucsVariables.RupGpInt = True) then
      if (TDetCod.GetValue('PGI_GROUPE') <> Null) then
        WGpEnCours := TDetCod.GetValue('PGI_GROUPE');

    TxGlobal := 0.00;
    MtGlobal := 0.00;

    WBaseMaj := 0.0;
    RegulNegative := False;
    I := 0;
    IndCotRegul := false; 
    while (I < TCot.Detail.Count) do
    begin
      TDetCod := TCot.Detail[I];
      if (TDetCod.GetValue('PHB_COTREGUL') =  'REG')  then
        IndCotRegul := true;
      if (TDetCod.GetValue('ETB_REGIMEALSACE') =  'X') then
        AlsaceMoselle := true
      else
        AlsaceMoselle := false;

      I := I + 1;
      WSalLu := TDetCod.GetValue('PHB_SALARIE');
      WRubLu := TDetCod.GetValue('PHB_RUBRIQUE');

      if (TDetCod.GetValue('POG_INSTITUTION') <> Null) then
        WInstLu := TDetCod.GetValue('POG_INSTITUTION');

      WDateLu := TDetCod.GetValue('PHB_DATEDEBUT');
      WDateLuFin := TDetCod.GetValue('PHB_DATEFIN');
      if (AlsaceMoselle) then
        WCodLu := TDetCod.GetValue('PDF_CODIFALSACE')
      else
        WCodLu := TDetCod.GetValue('PDF_CODIFICATION');
      WBaseLu := Arrondi(TDetCod.GetValue('PHB_BASECOT'), 5);

      if (DucsVariables.RupSiret = True) then
        if (TDetCod.GetValue('ET_SIRET') <> Null) then
          WsiretLu := TDetCod.GetValue('ET_SIRET');
      if (DucsVariables.RupNumInt = True) then
        if (TDetCod.GetValue('POG_NUMINTERNE') <> Null) then
          WNoIntLu := TDetCod.GetValue('POG_NUMINTERNE');
      if (DucsVariables.RupApe = True) then
        if (TDetCod.GetValue('ET_APE') <> Null) then
          WApeLu := TDetCod.GetValue('ET_APE');
      if (DucsVariables.RupGpInt = True) then
        if (TDetCod.GetValue('PGI_GROUPE') <> Null) then
          WGpLu := TDetCod.GetValue('PGI_GROUPE');

      if (((DucsVariables.RupSiret = True) and (WsiretLu = WsiretEnCours)) or
        (DucsVariables.RupSiret = False)) and
        (((DucsVariables.RupApe = True) and (WApeLu = WApeEnCours)) or
        (DucsVariables.RupApe = False)) and
        (((DucsVariables.RupNumInt = True) and (WNoIntLu = WNoIntEnCours)) or
        (DucsVariables.RupNumInt = False)) and
        (((DucsVariables.RupGpInt = True) and (WGpLu = WGpEnCours)) or
        (DucsVariables.RupGpInt = False)) and
        (WSalLu = WSalEnCours) and
        (WInstLu = WInstEnCours) and
        (WCodLu = WCodEnCours) and
        ((WBaseLu = WBaseEnCours) or (WRubLu = WRubEnCours)) then
      begin
        if (WRubLu <> WRubEnCours) then
        begin
        	TxGlobal := TxGlobal + Arrondi(TDetCod.GetValue('TAUXGLOBAL'), 5);
          WBaseMaj := 0.00;
        end;
        MtGlobal := MtGlobal + Arrondi(TDetCod.GetValue('MTGLOBAL'), 5);
        if (WRubLu = WRubEnCours) then
        begin
        if (WBaseMaj = 0.00) then
            WBaseMaj:=  WBaseEnCours;
          WBaseMaj :=  WBaseMaj  + WBaseLu;

        if (WbaseEnCours = WbaseLu *-1) then
          RegulNegative := True;
        end;

        WRubEnCours := TDetCod.GetValue('PHB_RUBRIQUE');
      end
      else
      begin
        WRubEnCours := '';
        I := I - 1;
        TCot2Fille := TOB.Create('', TCot2, -1);
        TCot2Fille.AddChampSupValeur('SALARIE', '', FALSE);
        TCot2Fille.AddChampSupValeur('CODIFICATION', '', FALSE);
        TCot2Fille.AddChampSupValeur('BASECOT', 0.00, FALSE);
        TCot2Fille.AddChampSupValeur('TAUXAT', 0.00, FALSE);
        TCot2Fille.AddChampSupValeur('DATEDEBUT', iDate1900, FALSE);
        TCot2Fille.AddChampSupValeur('DATEFIN', iDate1900, FALSE);
        TCot2Fille.AddChampSupValeur('ORGANISME', '', FALSE);
        TCot2Fille.AddChampSupValeur('INSTITUTION', '', FALSE);
        TCot2Fille.AddChampSupValeur('TAUXGLOBAL', 0.00, FALSE);
        TCot2Fille.AddChampSupValeur('MTGLOBAL', 0.00, FALSE);
        TCot2Fille.AddChampSupValeur('SIRET', '', FALSE);
        TCot2Fille.AddChampSupValeur('APE', '', FALSE);
        TCot2Fille.AddChampSupValeur('NUMINTERNE', '', FALSE);
        TCot2Fille.AddChampSupValeur('GROUPE', '', FALSE);
        TCot2Fille.AddChampSupValeur('REGIMEALSACE', '-', FALSE);

        if (TDetCod.GetValue('POG_POSTOTAL') = 0) and
          (TDetCod.GetValue('POG_LONGTOTAL') = 0) then
        { ducs standardisée - on a besoin du n° institution}
        // TCot2Fille.AddChampSup('POG_INSTITUTION',FALSE)
        else
          {ducs personnalisée}
          Personnalisee := True;

        TCot2Fille.PutValue('SALARIE', WSalEnCours);
        TCot2Fille.PutValue('CODIFICATION', WCodEnCours);
        TCot2Fille.PutValue('BASECOT', Arrondi(WBaseEnCours, 5));
        if (WBaseMaj <> 0.0) or (RegulNegative)  then 
        begin
          TCot2Fille.PutValue('BASECOT', Arrondi(WBaseMaj, 5));
          WBaseMaj := 0.0;
        end;
        TCot2Fille.PutValue('DATEDEBUT', WDateEnCours);
        TCot2Fille.PutValue('DATEFIN', StrToDATe(FinPer));
        if (Copy(WCodEnCours, 1, 6) <> WcodPrec) then
        begin
          WCodPrec := Copy(WCodEnCours, 1, 6);
          if (((AlsaceMoselle) and
               (ExisteSQL('SELECT PDP_CODIFALSACE FROM DUCSPARAM ' +
                          'WHERE ##PDP_PREDEFINI## PDP_TYPECOTISATION="A" AND ' +
                          'PDP_CODIFALSACE LIKE "' +
                          Copy(WCodEnCours, 1, 6) + '%"'))) or
              ((not AlsaceMoselle) and
               (ExisteSQL('SELECT PDP_CODIFICATION FROM DUCSPARAM ' +
                          'WHERE ##PDP_PREDEFINI## PDP_TYPECOTISATION="A" AND ' +
                          'PDP_CODIFICATION LIKE "' +
                           Copy(WCodEnCours, 1, 6) + '%"')))) then
          begin
            TCot2Fille.PutValue('TAUXAT', Arrondi(WTxATEnCours, 5));
            WPrecExiste := True;
          end
          else
          begin
            TCot2Fille.PutValue('TAUXAT', 0.00);
            WPrecExiste := False
          end;
        end
        else
        begin
          if (WPrecExiste = True) then
            TCot2Fille.PutValue('TAUXAT', Arrondi(WTxATEnCours, 5))
          else
            TCot2Fille.PutValue('TAUXAT', 0.00)
        end;
        TCot2Fille.PutValue('ORGANISME', WOrgEnCours);

        {ducs standardisée - on a besoin du n° institution}
        TCot2Fille.PutValue('INSTITUTION', WInstEnCours);

        TCot2Fille.PutValue('TAUXGLOBAL', Arrondi(TxGlobal, 5));
        TCot2Fille.PutValue('MTGLOBAL', Arrondi(MtGlobal, 5));
        TCot2Fille.PutValue('REGIMEALSACE',TDetCod.GetValue('ETB_REGIMEALSACE')); 

        if (DucsVariables.RupSiret = True) then TCot2Fille.PutValue('SIRET', WsiretEnCours);
        if (DucsVariables.RupApe = True) then TCot2Fille.PutValue('APE', WApeEnCours);
        if (DucsVariables.RupNumInt = True) then TCot2Fille.PutValue('NUMINTERNE', WNoIntEnCours);
        if (DucsVariables.RupGpInt = True) then TCot2Fille.PutValue('GROUPE', WGpEnCours);

        WSalEnCours := TDetCod.GetValue('PHB_SALARIE');
        if (DucsVariables.RupGpInt = True) or
          ((DucsParametres.Regroupt <> '') and (DucsVariables.RupGpInt = False)) then
          if (TDetCod.GetValue('POG_INSTITUTION') <> Null) then
            WInstEnCours := TDetCod.GetValue('POG_INSTITUTION');
        WDateEnCours := TDetCod.GetValue('PHB_DATEDEBUT');
        WDateFinEnCours := TDetCod.GetValue('PHB_DATEFIN');
        if (alsaceMoselle) then
          WCodEnCours := TDetCod.GetValue('PDF_CODIFALSACE')
        else
          WCodEnCours := TDetCod.GetValue('PDF_CODIFICATION');
        WBaseEnCours := Arrondi(TDetCod.GetValue('PHB_BASECOT'), 5);
        WTxATEnCours := Arrondi(TDetCod.GetValue('PHB_TAUXAT'), 5);
        WOrgEnCours := TDetCod.GetValue('PHB_ORGANISME');
        RegulNegative := False;

        if (DucsVariables.RupSiret = True) then
          if (TDetCod.GetValue('ET_SIRET') <> Null) then
            WsiretEnCours := TDetCod.GetValue('ET_SIRET');
        if (DucsVariables.RupNumInt = True) then
          if (TDetCod.GetValue('POG_NUMINTERNE') <> Null) then
            WNoIntEnCours := TDetCod.GetValue('POG_NUMINTERNE');
        if (DucsVariables.RupApe = True) then
          if (TDetCod.GetValue('ET_APE') <> Null) then
            WApeEnCours := TDetCod.GetValue('ET_APE');
        if (DucsVariables.RupGpInt = True) then
          if (TDetCod.GetValue('PGI_GROUPE') <> Null) then
            WGpEnCours := TDetCod.GetValue('PGI_GROUPE');
        TxGlobal := 0.00;
        MtGlobal := 0.00;
      end;
    end;

    { dernier élément}
    TCot2Fille := TOB.Create('', TCot2, -1);
    TCot2Fille.AddChampSupValeur('SALARIE', '', FALSE);
    TCot2Fille.AddChampSupValeur('CODIFICATION', '', FALSE);
    TCot2Fille.AddChampSupValeur('BASECOT', 0.00, FALSE);
    TCot2Fille.AddChampSupValeur('TAUXAT', 0.00, FALSE);
    TCot2Fille.AddChampSupValeur('DATEDEBUT', iDate1900, FALSE);
    TCot2Fille.AddChampSupValeur('DATEFIN', iDate1900, FALSE);
    TCot2Fille.AddChampSupValeur('ORGANISME', '', FALSE);
    TCot2Fille.AddChampSupValeur('INSTITUTION', '', FALSE);
    TCot2Fille.AddChampSupValeur('TAUXGLOBAL', 0.00, FALSE);
    TCot2Fille.AddChampSupValeur('MTGLOBAL', 0.00, FALSE);
    TCot2Fille.AddChampSupValeur('SIRET', '', FALSE);
    TCot2Fille.AddChampSupValeur('APE', '', FALSE);
    TCot2Fille.AddChampSupValeur('NUMINTERNE', '', FALSE);
    TCot2Fille.AddChampSupValeur('GROUPE', '', FALSE);
    TCot2Fille.AddChampSupValeur('REGIMEALSACE', '-', FALSE);

    TCot2Fille.PutValue('SALARIE', WSalEnCours);
    TCot2Fille.PutValue('CODIFICATION', WCodEnCours);
    TCot2Fille.PutValue('BASECOT', Arrondi(WBaseEnCours, 5));
    TCot2Fille.PutValue('DATEDEBUT', WDateEnCours);

    if (WBaseMaj <> 0.0) and (not RegulNegative)  then
    begin
      TCot2Fille.PutValue('BASECOT', Arrondi(WBaseMaj, 5));
      WBaseMaj := 0.0;
    end;
        TCot2Fille.PutValue('DATEFIN', StrToDATe(FinPer));

    if (Copy(WCodEnCours, 1, 6) <> WcodPrec) then
    begin
      WCodPrec := Copy(WCodEnCours, 1, 6);
      if (((AlsaceMoselle) and
           (ExisteSQL('SELECT PDP_CODIFALSACE FROM DUCSPARAM ' +
                      'WHERE ##PDP_PREDEFINI## PDP_TYPECOTISATION="A" AND ' +
                      'PDP_CODIFALSACE LIKE "' +
                      Copy(WCodEnCours, 1, 6) + '%"'))) or
          ((not AlsaceMoselle) and
           (ExisteSQL('SELECT PDP_CODIFICATION FROM DUCSPARAM ' +
                      'WHERE ##PDP_PREDEFINI## PDP_TYPECOTISATION="A" AND ' +
                      'PDP_CODIFICATION LIKE "' +
                      Copy(WCodEnCours, 1, 6) + '%"')))) then
      begin
        TCot2Fille.PutValue('TAUXAT', Arrondi(WTxATEnCours, 5));
      end
      else
      begin
        TCot2Fille.PutValue('TAUXAT', 0.00);
      end;
    end
    else
    begin
      if (WPrecExiste = True) then
        TCot2Fille.PutValue('TAUXAT', Arrondi(WTxATEnCours, 5))
      else
        TCot2Fille.PutValue('TAUXAT', 0.00);
    end;

    TCot2Fille.PutValue('ORGANISME', WOrgEnCours);
    TCot2Fille.PutValue('INSTITUTION', WInstEnCours);
    TCot2Fille.PutValue('TAUXGLOBAL', Arrondi(TxGlobal, 5));
    TCot2Fille.PutValue('MTGLOBAL', Arrondi(MtGlobal, 5));
    TCot2Fille.PutValue('REGIMEALSACE',TDetCod.GetValue('ETB_REGIMEALSACE'));

    if (DucsVariables.RupSiret = True) then TCot2Fille.PutValue('SIRET', WsiretEnCours);
    if (DucsVariables.RupApe = True) then TCot2Fille.PutValue('APE', WApeEnCours);
    if (DucsVariables.RupNumInt = True) then TCot2Fille.PutValue('NUMINTERNE', WNoIntLu);
    if (DucsVariables.RupGpInt = True) then TCot2Fille.PutValue('GROUPE', WGpEnCours);

  end;

  TCot2.Detail.Sort(StOrder4);
  IndicD := 0;
  IndicP := 0;
  TauxGlobalEnCours := 0;

  SalEnCours := '';
  BaseEnCours := 0;
  WMtGlobalEnCours := 0;
  WCodifEnCours := '';
  TauxAtEnCours := 0;

  TCotA := TOB.Create ('Les codifications type AT', nil, -1);
  TCot2bis := TOB.Create ('Les codifications non AT', nil, -1);
  // on éclate TCot2 en 2 TOB :
  //    - TCotA, tob des rubrique de type "présentation par taux AT"
  //    - TCot2Bis, tob des autres rubriques

  for i := 0 to TCot2.Detail.Count-1 do
  begin
    TCot2Fille := TCot2.Detail[i];
    if (ARRONDI(TCot2Fille.GetValue('TAUXAT'),5) <> 0) then
    // si Taux AT renseigné - TOB TCotA des rubriques de type "présentation par taux AT"
    begin
      Baselue := ARRONDI(TCot2Fille.GetValue('BASECOT'),5);
      CodifW := TCot2Fille.GetValue('CODIFICATION');
      SalW :=  TCot2Fille.GetValue('SALARIE');
      DateDebW := TCot2Fille.GetValue('DATEDEBUT');
      DateFinW := TCot2Fille.GetValue('DATEFIN');
      TauxAtW :=  TCot2Fille.GetValue('TAUXAT');
      WMtGlobalLu   := ARRONDI(TCot2Fille.GetValue('MTGLOBAL'),5);
      TauxGlobalLu := ARRONDI(TCot2Fille.GetValue('TAUXGLOBAL'),5);

      if ((SalW = SalEnCours) or (SalEnCours = '')) and
         ((CodifW = WCodifEnCours) or (WCodifEnCours = ''))then
      begin
      // Pour un salarié, une codification --> Cumul des Bases à Taux identiques
        if (TauxGlobalLu = TauxGlobalEnCours) or (TauxGlobalEnCours = 0) then
        begin
          BaseEnCours := BaseEnCours + Baselue;
          TauxGlobalEnCours := TauxGlobalLu;
          WMtGlobalEnCours := WMtGlobalEnCours + WMtGlobalLu;
          WCodifEnCours := CodifW;
          SalEnCours := SalW;
          TauxAtEnCours := TauxAtW;
        end
        else
        begin
          if (SalEnCours <> '') and  (WCodifEnCours <> '') and (TauxAtEnCours <> 0) then
          // Test pour traiter le cas où l'élément précédent a un taux à 0
          // (c.a.d mis en Tob TCot2Bis) et que cependant sa codification est
          // de type "présentation par taux AT"
          begin
            // Changement de taux : mise en TOB de l'élément en cours
            TCotAFille := TOB.Create ('',TCotA,-1);
            TCotAFille.AddChampSupValeur ('SALARIE',SalEnCours);
            TCotAFille.AddChampSupValeur ('CODIFICATION',WCodifEnCours);
            TCotAFille.AddChampSupValeur ('BASECOT',BaseEnCours);
            TCotAFille.AddChampSupValeur ('TAUXAT',TauxAtEnCours);
            TCotAFille.AddChampSupValeur ('DATEDEBUT',DateDebW);
            TCotAFille.AddChampSupValeur ('DATEFIN',TCot2Fille.GetValue('DATEFIN'));
            TCotAFille.AddChampSupValeur ('ORGANISME',TCot2Fille.GetValue('ORGANISME'));
            TCotAFille.AddChampSupValeur ('INSTITUTION',TCot2Fille.GetValue('INSTITUTION'));
            TCotAFille.AddChampSupValeur ('TAUXGLOBAL',TauxGlobalEnCours);
            TCotAFille.AddChampSupValeur ('MTGLOBAL',WMtGlobalEnCours);
            TCotAFille.AddChampSupValeur ('SIRET',TCot2Fille.GetValue('SIRET'));
            TCotAFille.AddChampSupValeur ('APE',TCot2Fille.GetValue('APE'));
            TCotAFille.AddChampSupValeur ('NUMINTERNE',TCot2Fille.GetValue('NUMINTERNE'));
            TCotAFille.AddChampSupValeur ('GROUPE',TCot2Fille.GetValue('GROUPE'));
            TCotAFille.AddChampSupValeur ('REGIMEALSACE',TCot2Fille.GetValue('REGIMEALSACE'));
          end;

          BaseEnCours :=  Baselue;
          TauxGlobalEnCours := TauxGlobalLu;
          WMtGlobalEnCours := WMtGlobalLu;
          WCodifEnCours := CodifW;
          TauxAtEnCours := TauxAtW;
        end;
      end
      else
      begin
        // Changement de salarié ou de codification : mise en TOB de l'élément en cours
        TCotAFille := TOB.Create ('',TCotA,-1);
        TCotAFille.AddChampSupValeur ('SALARIE',SalEnCours);
        TCotAFille.AddChampSupValeur ('CODIFICATION',WCodifEnCours);
        TCotAFille.AddChampSupValeur ('BASECOT',BaseEnCours);
        TCotAFille.AddChampSupValeur ('TAUXAT',TauxAtEnCours);
        TCotAFille.AddChampSupValeur ('DATEDEBUT',DateDebW);
        TCotAFille.AddChampSupValeur ('DATEFIN',TCot2Fille.GetValue('DATEFIN'));
        TCotAFille.AddChampSupValeur ('ORGANISME',TCot2Fille.GetValue('ORGANISME'));
        TCotAFille.AddChampSupValeur ('INSTITUTION',TCot2Fille.GetValue('INSTITUTION'));
        TCotAFille.AddChampSupValeur ('TAUXGLOBAL',TauxGlobalEnCours);
        TCotAFille.AddChampSupValeur ('MTGLOBAL',WMtGlobalEnCours);
        TCotAFille.AddChampSupValeur ('SIRET',TCot2Fille.GetValue('SIRET'));
        TCotAFille.AddChampSupValeur ('APE',TCot2Fille.GetValue('APE'));
        TCotAFille.AddChampSupValeur ('NUMINTERNE',TCot2Fille.GetValue('NUMINTERNE'));
        TCotAFille.AddChampSupValeur ('GROUPE',TCot2Fille.GetValue('GROUPE'));
        TCotAFille.AddChampSupValeur ('REGIMEALSACE',TCot2Fille.GetValue('REGIMEALSACE'));

        BaseEnCours :=  Baselue;
        TauxGlobalEnCours := TauxGlobalLu;
        SalEnCours := SalW;
        WMtGlobalEnCours := WMtGlobalLu;
        WCodifEnCours := CodifW;
        TauxAtEnCours := TauxAtW;
      end;
    end
    else
    // taux AT = 0 - TOB TCot2Bis des autres rubriques
    begin
      if (WCodifEnCours <> '') and (  SalEnCours <>'') then
      // TOB TCotA : mise en Tob du dernier élément (élt en cours)
      begin
        TCotAFille := TOB.Create ('',TCotA,-1);
        TCotAFille.AddChampSupValeur ('SALARIE',SalEnCours);
        TCotAFille.AddChampSupValeur ('CODIFICATION',WCodifEnCours);
        TCotAFille.AddChampSupValeur ('BASECOT',BaseEnCours);
        TCotAFille.AddChampSupValeur ('TAUXAT',TauxAtEnCours);
        TCotAFille.AddChampSupValeur ('DATEDEBUT',DateDebW);
        TCotAFille.AddChampSupValeur ('DATEFIN',TCot2Fille.GetValue('DATEFIN'));
        TCotAFille.AddChampSupValeur ('ORGANISME',TCot2Fille.GetValue('ORGANISME'));
        TCotAFille.AddChampSupValeur ('INSTITUTION',TCot2Fille.GetValue('INSTITUTION'));
        TCotAFille.AddChampSupValeur ('TAUXGLOBAL',TauxGlobalEnCours);
        TCotAFille.AddChampSupValeur ('MTGLOBAL',WMtGlobalEnCours);
        TCotAFille.AddChampSupValeur ('SIRET',TCot2Fille.GetValue('SIRET'));
        TCotAFille.AddChampSupValeur ('APE',TCot2Fille.GetValue('APE'));
        TCotAFille.AddChampSupValeur ('NUMINTERNE',TCot2Fille.GetValue('NUMINTERNE'));
        TCotAFille.AddChampSupValeur ('GROUPE',TCot2Fille.GetValue('GROUPE'));
        TCotAFille.AddChampSupValeur ('REGIMEALSACE',TCot2Fille.GetValue('REGIMEALSACE'));

        SalEnCours := '';
        BaseEnCours := 0;
        WMtGlobalEnCours := 0;
        WCodifEnCours := '';
        TauxAtEnCours := 0;

      end;
      TCot2BisFille := TOB.Create ('',TCot2Bis,-1);
      TCot2BisFille.AddChampSupValeur ('SALARIE',TCot2Fille.GetValue('SALARIE'));
      TCot2BisFille.AddChampSupValeur ('CODIFICATION',TCot2Fille.GetValue('CODIFICATION'));
      TCot2BisFille.AddChampSupValeur ('BASECOT',TCot2Fille.GetValue('BASECOT'));
      TCot2BisFille.AddChampSupValeur ('TAUXAT',TCot2Fille.GetValue('TAUXAT'));
      TCot2BisFille.AddChampSupValeur ('DATEDEBUT',TCot2Fille.GetValue('DATEDEBUT'));
      TCot2BisFille.AddChampSupValeur ('DATEFIN',TCot2Fille.GetValue('DATEFIN'));
      TCot2BisFille.AddChampSupValeur ('ORGANISME',TCot2Fille.GetValue('ORGANISME'));
      TCot2BisFille.AddChampSupValeur ('INSTITUTION',TCot2Fille.GetValue('INSTITUTION'));
      TCot2BisFille.AddChampSupValeur ('TAUXGLOBAL',TCot2Fille.GetValue('TAUXGLOBAL'));
      TCot2BisFille.AddChampSupValeur ('MTGLOBAL',TCot2Fille.GetValue('MTGLOBAL'));
      TCot2BisFille.AddChampSupValeur ('SIRET',TCot2Fille.GetValue('SIRET'));
      TCot2BisFille.AddChampSupValeur ('APE',TCot2Fille.GetValue('APE'));
      TCot2BisFille.AddChampSupValeur ('NUMINTERNE',TCot2Fille.GetValue('NUMINTERNE'));
      TCot2BisFille.AddChampSupValeur ('GROUPE',TCot2Fille.GetValue('GROUPE'));
      TCot2BisFille.AddChampSupValeur ('REGIMEALSACE',TCot2Fille.GetValue('REGIMEALSACE'));
    end;
  end;
//TCotA.SaveToFile('c:\tmp\TCotA', False, TRUE, TRUE);
  if (TCotA.Detail.count <> 0) then
  begin
    TCotA.Detail.Sort('TAUXAT;SALARIE;CODIFICATION;TAUXGLOBAL;BASECOT;DATEDEBUT;');

    NbCodif := 0;

    TCotABis := TOB.Create ('Les codifications type AT', nil, -1);
    TauxGlobalEnCours := 0;
    SalEnCours := '';
    BaseEnCours := 0;
    TauxAtW := 0;
    WMtGlobalLu := 0;
    for i := 0 to TCotA.Detail.Count-1 do
    begin
      TCotAFille := TCotA.Detail[i];
      Baselue := ARRONDI(TCotAFille.GetValue('BASECOT'),5);
      TauxGlobalLu := ARRONDI(TCotAFille.GetValue('TAUXGLOBAL'),5);
      SalW := TCotAFille.GetValue('SALARIE');
      TauxAtW :=  TCotAFille.GetValue('TAUXAT');
      WMtGlobalLu := ARRONDI(TCotAFille.GetValue('MTGLOBAL'),5);
      CodifW := TCotAFille.GetValue('CODIFICATION');

      if ((SalW = SalEnCours) or (SalEnCours = '')) and ((CodifW = WCodifEnCours) or (WCodifEnCours = '')) then
      // Pour un salarié, une codification --> Cumul des Taux à Bases identiques
      begin
        if (Baselue = BaseEnCours) or (BaseEnCours = 0) then
        begin
          BaseEnCours := Baselue;
          TauxGlobalEnCours := TauxGlobalEnCours + TauxGlobalLu;
          SalEnCours := SalW;
          TauxAtEnCours := TauxAtW;
          WMtGlobalEnCours := WMtGlobalLu;
          WCodifEnCours := CodifW;
        end
        else
        begin
          // Changement de base : mise en TOB de l'élément en cours
          TCotABisFille := TOB.Create ('',TCotABis,-1);
          Nbcodif := NbCodif+1;
          TCotABisFille.AddChampSupValeur ('SALARIE',SalEnCours);
          TCotABisFille.AddChampSupValeur ('CODIFICATION',WCodifEnCours);
          TCotABisFille.AddChampSupValeur ('BASECOT',BaseEnCours);
          TCotABisFille.AddChampSupValeur ('TAUXAT',TauxAtEnCours);
          TCotABisFille.AddChampSupValeur ('DATEDEBUT',DateDebW);
          TCotABisFille.AddChampSupValeur ('DATEFIN',TCotAFille.GetValue('DATEFIN'));
          TCotABisFille.AddChampSupValeur ('ORGANISME',TCotAFille.GetValue('ORGANISME'));
          TCotABisFille.AddChampSupValeur ('INSTITUTION',TCotAFille.GetValue('INSTITUTION'));
          TCotABisFille.AddChampSupValeur ('TAUXGLOBAL',TauxGlobalEnCours);
          TCotABisFille.AddChampSupValeur ('MTGLOBAL',WMtGlobalEnCours);
          TCotABisFille.AddChampSupValeur ('SIRET',TCotAFille.GetValue('SIRET'));
          TCotABisFille.AddChampSupValeur ('APE',TCotAFille.GetValue('APE'));
          TCotABisFille.AddChampSupValeur ('NUMINTERNE',TCotAFille.GetValue('NUMINTERNE'));
          TCotABisFille.AddChampSupValeur ('GROUPE',TCotAFille.GetValue('GROUPE'));
          TCotABisFille.AddChampSupValeur ('INDIC',0);
          TCotABisFille.AddChampSupValeur ('REGIMEALSACE',TCotAFille.GetValue('REGIMEALSACE'));

          BaseEnCours :=  Baselue;
          TauxGlobalEnCours := TauxGlobalLu;
          WMtGlobalEnCours := WMtGlobalLu;
          WCodifEnCours := CodifW;
          TauxAtEnCours := TauxAtW;
         end;
      end
      else
      begin
        // Changement de salarié ou de codification : mise en TOB de l'élément en cours
        TCotABisFille := TOB.Create ('',TCotABis,-1);
        Nbcodif := NbCodif+1;
        TCotABisFille.AddChampSupValeur ('SALARIE',SalEnCours);
        TCotABisFille.AddChampSupValeur ('CODIFICATION',WCodifEnCours);
        TCotABisFille.AddChampSupValeur ('BASECOT',BaseEnCours);
        TCotABisFille.AddChampSupValeur ('TAUXAT',TauxAtEnCours);
        TCotABisFille.AddChampSupValeur ('DATEDEBUT',DateDebW);
        TCotABisFille.AddChampSupValeur ('DATEFIN',TCotAFille.GetValue('DATEFIN'));
        TCotABisFille.AddChampSupValeur ('ORGANISME',TCotAFille.GetValue('ORGANISME'));
        TCotABisFille.AddChampSupValeur ('INSTITUTION',TCotAFille.GetValue('INSTITUTION'));
        TCotABisFille.AddChampSupValeur ('TAUXGLOBAL',TauxGlobalEnCours);
        TCotABisFille.AddChampSupValeur ('MTGLOBAL',WMtGlobalEnCours);
        TCotABisFille.AddChampSupValeur ('SIRET',TCotAFille.GetValue('SIRET'));
        TCotABisFille.AddChampSupValeur ('APE',TCotAFille.GetValue('APE'));
        TCotABisFille.AddChampSupValeur ('NUMINTERNE',TCotAFille.GetValue('NUMINTERNE'));
        TCotABisFille.AddChampSupValeur ('GROUPE',TCotAFille.GetValue('GROUPE'));
        TCotABisFille.AddChampSupValeur ('INDIC',0);
        TCotABisFille.AddChampSupValeur ('REGIMEALSACE',TCotAFille.GetValue('REGIMEALSACE'));

        BaseEnCours :=  Baselue;
        TauxGlobalEnCours := TauxGlobalLu;
        SalEnCours  := SalW;
        WMtGlobalEnCours := WMtGlobalLu;
        WCodifEnCours := CodifW;
        TauxAtEnCours := TauxAtW;
      end;
    end;
    // TOB TCotABis : mise en Tob du dernier élément (élt en cours)
    TCotABisFille := TOB.Create ('',TCotABis,-1);
    Nbcodif := NbCodif+1;
    TCotABisFille.AddChampSupValeur ('SALARIE',SalEnCours);
    TCotABisFille.AddChampSupValeur ('CODIFICATION',WCodifEnCours);
    TCotABisFille.AddChampSupValeur ('BASECOT',BaseEnCours);
    TCotABisFille.AddChampSupValeur ('TAUXAT',TauxAtEnCours);
    TCotABisFille.AddChampSupValeur ('DATEDEBUT',DateDebW);
    TCotABisFille.AddChampSupValeur ('DATEFIN',TCotAFille.GetValue('DATEFIN'));
    TCotABisFille.AddChampSupValeur ('ORGANISME',TCotAFille.GetValue('ORGANISME'));
    TCotABisFille.AddChampSupValeur ('INSTITUTION',TCotAFille.GetValue('INSTITUTION'));
    TCotABisFille.AddChampSupValeur ('TAUXGLOBAL',TauxGlobalEnCours);
    TCotABisFille.AddChampSupValeur ('MTGLOBAL',WMtGlobalEnCours);
    TCotABisFille.AddChampSupValeur ('SIRET',TCotAFille.GetValue('SIRET'));
    TCotABisFille.AddChampSupValeur ('APE',TCotAFille.GetValue('APE'));
    TCotABisFille.AddChampSupValeur ('NUMINTERNE',TCotAFille.GetValue('NUMINTERNE'));
    TCotABisFille.AddChampSupValeur ('GROUPE',TCotAFille.GetValue('GROUPE'));
    TCotABisFille.AddChampSupValeur ('INDIC',0);
    TCotABisFille.AddChampSupValeur ('REGIMEALSACE',TCotAFille.GetValue('REGIMEALSACE'));

    TCotABis.Detail.Sort('CODIFICATION;TAUXAT;TAUXGLOBAL;BASECOT;DATEDEBUT;');

    INDIC := 0;
    TauxAtEnCours := 0;
    TauxGlobalEnCours := 0;
    WCodifEnCours := '';

    for i := 0 to TCotABis.Detail.Count-1 do
    begin
      TCotABisFille := TCotABis.Detail[i];
      TauxAtW := TCotABisFille.GetValue('TAUXAT');
      TauxGlobalLu := Arrondi(TCotABisFille.GetValue('TAUXGLOBAL'),5);
      CodifW := TCotABisFille.GetValue('CODIFICATION');
      SalW := TCotABisFille.GetValue('SALARIE');
      if (copy(CodifW,7,1) = 'D') then
      begin
        if ((CodifW = WCodifEnCours) or (WCodifEnCours= '')) and
           ((TauxGlobalLu = TauxGlobalEnCours) or (TauxGlobalEnCours=0))and
           ((TauxAtW = TauxAtEnCours) or (TauxAtEnCours=0)) then
        begin
          TCotABisFille.PutValue ('INDIC',INDIC);
        end
        else
        begin
          INDIC := INDIC+1;
          TCotABisFille.PutValue ('INDIC',INDIC);
        end;
        TauxAtEnCours := TauxAtW;
        TauxGlobalEnCours := TauxGlobalLu;
        WCodifEnCours := CodifW;
      end;
    end;

    for i := 0 to TCotABis.Detail.Count-1 do
    begin
      TCotABisFille := TCotABis.Detail[i];
      TauxAtW := TCotABisFille.GetValue('TAUXAT');
      SalW := TCotABisFille.GetValue('SALARIE');
      CodifW := TCotABisFille.GetValue('CODIFICATION');
      if (copy(CodifW,7,1) = 'D') then
      begin
        TCotABisFilleP :=  TCotABis.FindFirst(['CODIFICATION','SALARIE','TAUXAT'],
                                         [copy(CodifW,1,6)+'P', SalW, TauxAtW], TRUE);
        while (TCotABisFilleP <> nil)  do
        begin
          TCotABisFilleP.AddChampSupValeur('INDIC',TCotABisFille.GetValue ('INDIC'));

          TCotABisFilleP :=  TCotABis.FindNext(['CODIFICATION','SALARIE','TAUXAT'],
                             [copy(CodifW,1,6)+'P', SalW, TauxAtW], TRUE);
        end
      end;
    end;
  end; // fin TCotA.Detail.Count <> 0

  // Création de la TOB TCot3 (détail des rubriques par codification et salarié)
  TCot3 := TOB.Create('les codifs à éditer', nil, -1);
  for i := 0 to NbCodif-1 do
  // on récupère d'abord les rubriquee de type "présentation par taux AT" (TCotA)
  begin
    TCotABisFille := TCotABis.Detail[i];

    TCot3Fille := TOB.Create ('',TCot3,-1);

    TCot3Fille.AddChampSupValeur ('SALARIE',TCotABisFille.GetValue('SALARIE'));
    TCot3Fille.AddChampSupValeur ('CODIFICATION',TCotABisFille.GetValue('CODIFICATION'));
    TCot3Fille.AddChampSupValeur ('BASECOT',TCotABisFille.GetValue('BASECOT'));
    TCot3Fille.AddChampSupValeur ('TAUXAT',TCotABisFille.GetValue('TAUXAT'));
    TCot3Fille.AddChampSupValeur ('DATEDEBUT',TCotABisFille.GetValue('DATEDEBUT'));
    TCot3Fille.AddChampSupValeur ('DATEFIN',TCotABisFille.GetValue('DATEFIN'));
    TCot3Fille.AddChampSupValeur ('ORGANISME',TCotABisFille.GetValue('ORGANISME'));
    TCot3Fille.AddChampSupValeur ('INSTITUTION',TCotABisFille.GetValue('INSTITUTION'));
    TCot3Fille.AddChampSupValeur ('TAUXGLOBAL',TCotABisFille.GetValue('TAUXGLOBAL'));
    TCot3Fille.AddChampSupValeur ('MTGLOBAL',TCotABisFille.GetValue('MTGLOBAL'));
    TCot3Fille.AddChampSupValeur ('SIRET',TCotABisFille.GetValue('SIRET'));
    TCot3Fille.AddChampSupValeur ('APE',TCotABisFille.GetValue('APE'));
    TCot3Fille.AddChampSupValeur ('NUMINTERNE',TCotABisFille.GetValue('NUMINTERNE'));
    TCot3Fille.AddChampSupValeur ('GROUPE',TCotABisFille.GetValue('GROUPE'));
    TCot3Fille.AddChampSupValeur ('INDIC',TCotABisFille.GetValue('INDIC'));
    TCot3Fille.AddChampSupValeur ('REGIMEALSACE',TCotABisFille.GetValue('REGIMEALSACE'));

  end;

  for i:= 0 to TCot2Bis.Detail.Count-1 do
  // on récupère ensuite les rubriquee de type "cas normal" (TCot2Bis)
  begin
      TCot2BisFille := TCot2Bis.Detail[i];

    TCot3Fille := TOB.Create ('',TCot3,-1);
    TCot3Fille.AddChampSupValeur ('SALARIE',TCot2BisFille.GetValue('SALARIE'));
    TCot3Fille.AddChampSupValeur ('CODIFICATION',TCot2BisFille.GetValue('CODIFICATION'));
    TCot3Fille.AddChampSupValeur ('BASECOT',TCot2BisFille.GetValue('BASECOT'));
    TCot3Fille.AddChampSupValeur ('TAUXAT',TCot2BisFille.GetValue('TAUXAT'));
    TCot3Fille.AddChampSupValeur ('DATEDEBUT',TCot2BisFille.GetValue('DATEDEBUT'));
    TCot3Fille.AddChampSupValeur ('DATEFIN',TCot2BisFille.GetValue('DATEFIN'));
    TCot3Fille.AddChampSupValeur ('ORGANISME',TCot2BisFille.GetValue('ORGANISME'));
    TCot3Fille.AddChampSupValeur ('INSTITUTION',TCot2BisFille.GetValue('INSTITUTION'));
    TCot3Fille.AddChampSupValeur ('TAUXGLOBAL',TCot2BisFille.GetValue('TAUXGLOBAL'));
    TCot3Fille.AddChampSupValeur ('MTGLOBAL',TCot2BisFille.GetValue('MTGLOBAL'));
    TCot3Fille.AddChampSupValeur ('SIRET',TCot2BisFille.GetValue('SIRET'));
    TCot3Fille.AddChampSupValeur ('APE',TCot2BisFille.GetValue('APE'));
    TCot3Fille.AddChampSupValeur ('NUMINTERNE',TCot2BisFille.GetValue('NUMINTERNE'));
    TCot3Fille.AddChampSupValeur ('GROUPE',TCot2BisFille.GetValue('GROUPE'));
    TCot3Fille.AddChampSupValeur ('INDIC',999);
    TCot3Fille.AddChampSupValeur ('REGIMEALSACE',TCot2BisFille.GetValue('REGIMEALSACE'));

// Le TAUX AT est à zéro : exite t'il dans TCotA, tob des rubriques de
// type "présentation par taux AT", une codification identique.
// Correspond au cas où le salarié n'a de cotisation AT.

    TCotABisFille := TCotA.FindFirst (['CODIFICATION','DATEDEBUT','DATEFIN'],
                                       [TCot2BisFille.GetValue('CODIFICATION'),
                                        TCot2BisFille.GetValue('DATEDEBUT'),
                                        TCot2BisFille.GetValue('DATEFIN')], TRUE);
    if TCotABisFille <> nil then
    begin
        Tcot3FilleBis :=  TCot3.FindFirst (['CODIFICATION','DATEDEBUT','DATEFIN'],
                                       [TCot2BisFille.GetValue('CODIFICATION'),
                                        TCot2BisFille.GetValue('DATEDEBUT'),
                                        TCot2BisFille.GetValue('DATEFIN')], TRUE);
        if Tcot3FilleBis <> nil then
          TCot3Fille.AddChampSupValeur ('INDIC',Tcot3FilleBis.GetValue('INDIC'));
    end;

   end;
    StOrder := '';
    if (DucsVariables.RupSiret = True)  then   StOrder := StOrder + 'SIRET;';
    if (DucsVariables.RupApe = True)    then   StOrder := StOrder + 'APE;';
    if (DucsVariables.RupNumInt = True) then StOrder := StOrder + 'NUMINTERNE;';
    if (DucsVariables.RupGpInt = True)  then  StOrder := StOrder + 'GROUPE;';

    Storder := StOrder+'INDIC;CODIFICATION;TAUXAT;TAUXGLOBAL;SALARIE;BASECOT;DATEDEBUT' ;

    TCot3.Detail.Sort(StOrder);
//    TobDebug(tcot3);
//    TCot3.SaveToFile('c:\tmp\tcot3', False, TRUE, TRUE);
  FreeAndNil(TCotA);
  FreeAndNil(TCot2);
  FreeAndNil(TCot2Bis);
  FreeAndNil(TCotABis);

  if (IndCotRegul) then
  begin
    if (not IndTraceE) then
    begin
      TraceE.add ('Anomalies de traitement');
      TraceE.add ('--------------------------------------------------------');
      IndTraceE := True;
      end;

      TraceE.Add('Etablissement '+ DucsParametres.Etab +#13#10+
                 ' Des rubriques de régularisation de cotisation .R ont été calculées'+#13#10+
                 ' dans les bulletins de la période sélectionnée et peuvent par conséquent entrainer'+#13#10+
                 ' une rupture sur les codifications DUCS');
  end;

  if TCot <> nil then
  begin
    TCot.Free;
  end;
end; { fin ChargeCot}

//Mise à jour des tables DUCSENTETE et DUCSDETAIL
//On Supprime la DUCS (entete et détail) si existe déjà
procedure MajDucs(DebPer, FinPer, PeriodEtat :string; DucsParametres: TDucsParametres; DucsDeclarant : TDucsDeclarant; DucsEffectifs : TDucsEffectifs; TCot3 : TOB; var DucsOrganisme : TDucsOrganisme; var DucsVariables : TDucsVariables;var  TraceE :TStringList; IndTraceE : boolean; TrupMere, Tinst : TOB);
var
  TDetail, TOB_Lignes, TOB_FilleTete, TOB_FilleLignes                 : TOB;
  TAT, TobSTExiste,  TEffM, TEffMFille, TRup                          : TOB;
  BaseCodif, BaseLigne, TauxEnCours, ATEnCours, TauxLu, ATLu          : double;
  MontForfait, MontCodif, MontST, MontLigne                           : double;
  Effectif, ligne, i, NoAT, NoATP, NoInstEnCours                      : Integer;
  PogSTPos, PogSTLong, PosDeb, LongEdit, j, NbJour                    : Integer;
  PogBoolST, PogTypDecl1, PogTypDecl2, IndRup                         : Boolean;
  Personnalisee, IndCalcSt                                            : Boolean;
  QDucsParam, QInst, QEtab, QOrg, Qdetail, QQ, QDucsParamAT           : TQuery;
  aa, mm, jj                                                          : Word;
  DatExigible, DatLimDepot, DatReglt                                  : TDateTime;
  CodifEnCours, CodifLue,CodifST, SalEnCours, SalLu, StrDelDetail     : string;
  OrgLu, OrgEnCours, NoInst, LibInst, StrDelTete                      : string;
  PogPeriod1, PogPeriod2, PogNature, Cod100, AbregePeriod, StQQ       : string;
  Monnaie, LigneOptique, BaseArrCod, MontArrCod, TypeCotis            : string;
  CodifAT, Predef, ZEDE, WSiretLu, WNoIntLu, WApeLu, WGpLu            : string;
  NumMsa, CodInstLu, CodInstEnCours                                   : string;
  PogCpInfo, PogCentrePayeur, CSC, LGOptiqueEncours                   : string;
  NomChamp                                             : array[1..2] of string;
  ValeurChamp                                          : array[1..2] of variant;
  ValChST                                              : array[1..1] of variant;
  NomChampST                                           : array[1..1] of string;
  Anomalie                                             : string;

{$IFDEF EAGLCLIENT}
  KK, SortDos                                                         : Integer;
  MaTob, MaTobAt                                                      : TOB;
{$ENDIF}

  PogTypBordereau1, PogTypBordereau2, PogDucsDoss : string; //PT1


begin
  Anomalie      := '';
  IndRup        := False;
  ZEDE          := 'ZZZZZZZ';
  PosDeb        := 0;
  LongEdit      := 0;
  NoInstEnCours := 0;

  DucsOrganisme.PogBoolST     := FALSE;
  DucsOrganisme.PogSTPos      := 0;
  DucsOrganisme.PogSTLong     := 0;

  {Récupération informations établissement}
  DucsOrganisme.PogNumInterne := 'Non Renseigné';
  DucsVariables.EtSiret       := 'Non Renseigné';
  DucsVariables.EtApe         := 'Inco';

  QEtab := OpenSql('SELECT ET_SIRET,ET_APE' +
                   ' FROM ETABLISS ' +
                   ' WHERE ET_ETABLISSEMENT = "' + DucsParametres.Etab + '"', True);
  if not QEtab.eof then
  begin
    ForceNumerique(QEtab.FindField('ET_SIRET').AsString, DucsVariables.EtSiret);
    DucsVariables.EtApe := QEtab.FindField('ET_APE').AsString;
  end;
  Ferme(QEtab);

  {Récupération informations organisme}
  QOrg := OpenSql('SELECT POG_NUMINTERNE, POG_NATUREORG,' +
                  'POG_SOUSTOTDUCS, POG_POSTOTAL, POG_LONGTOTAL,' +
                  'POG_PERIODICITDUCS, POG_AUTREPERIODUCS,' +
                  'POG_PERIODCALCUL, POG_AUTPERCALCUL,' +
                  'POG_LGOPTIQUE' +
                  ',POG_CPINFO,POG_CENTREPAYEUR' +
                  ',POG_TYPEPERIOD,POG_TYPEAUTPERIOD'+
                  ' FROM ORGANISMEPAIE ' +
                  ' WHERE POG_ETABLISSEMENT= "' + DucsParametres.Etab + '" AND ' +
                  ' POG_ORGANISME = "' + DucsParametres.Organisme + '"', True);
  if not QOrg.eof then
  begin
    DucsOrganisme.PogNumInterne := copy(QOrg.FindField('POG_NUMINTERNE').AsString, 1, 30);
    DucsOrganisme.PogNature := QOrg.FindField('POG_NATUREORG').AsString;
    DucsOrganisme.PogBoolST := QOrg.FindField('POG_SOUSTOTDUCS').AsString = 'X';
    DucsOrganisme.PogSTPos := QOrg.FindField('POG_POSTOTAL').AsInteger;
    DucsOrganisme.PogSTLong := QOrg.FindField('POG_LONGTOTAL').AsInteger;
    DucsOrganisme.PogPeriod1 := QOrg.FindField('POG_PERIODICITDUCS').AsString;
    DucsOrganisme.PogPeriod2 := QOrg.FindField('POG_AUTREPERIODUCS').AsString;
    DucsOrganisme.PogTypDecl1 := QOrg.FindField('POG_PERIODCALCUL').AsString = 'X';
    DucsOrganisme.PogTypDecl2 := QOrg.FindField('POG_AUTPERCALCUL').AsString = 'X';
    DucsOrganisme.PogLGOptique := QOrg.FindField('POG_LGOPTIQUE').AsString;
    DucsOrganisme.PogCPInfo := QOrg.FindField('POG_CPINFO').AsString;
    DucsOrganisme.PogCentrePayeur := QOrg.FindField('POG_CENTREPAYEUR').AsString;

// d PT1
    PogTypBordereau1 := QOrg.FindField('POG_TYPEPERIOD').AsString;
    PogTypBordereau2 := QOrg.FindField('POG_TYPEAUTPERIOD').AsString;
    if (DucsParametres.DucsDoss) then
      PogDucsDoss := 'X'
    else
      PogDucsDoss := '-';

   if (PogTypBordereau1 = '') then
   begin
     try
     begintrans;
      PogTypBordereau1 := AffectTypeBordereau (PogNature, PogPeriod1,
                                         DucsParametres.DucsDoss);
      ExecuteSQL('UPDATE ORGANISMEPAIE SET POG_TYPEPERIOD = "'+PogTypBordereau1+'" '+
                 'WHERE POG_NATUREORG = "'+PogNature+'" AND '+
                 'POG_TYPEPERIOD = "" AND '+
                 'POG_PERIODICITDUCS = "'+PogPeriod1+'" AND '+
                 'POG_DUCSDOSSIER ="'+PogDucsDoss+'"');

     Committrans;
     except
     Rollback;
     end;
   end;
   if (PogTypBordereau2 = '') then
   begin
     try
     begintrans;
      PogTypBordereau2 := AffectTypeBordereau (PogNature, PogPeriod2,
                                         DucsParametres.DucsDoss);
      ExecuteSQL('UPDATE ORGANISMEPAIE SET POG_TYPEAUTPERIOD = "'+PogTypBordereau2+'" '+
                 'WHERE POG_NATUREORG = "'+PogNature+'" AND '+
                 'POG_TYPEAUTPERIOD ="" AND '+
                 'POG_AUTREPERIODUCS = "'+PogPeriod2+'" AND '+
                 'POG_DUCSDOSSIER ="'+PogDucsDoss+'"');
     Committrans;
     except
     Rollback;
     end;
   end;
// f PT1

   // récupération du type de bordereau
   if (PeriodEtat = DucsOrganisme.PogPeriod1) then
// PT1     DucsOrganisme.PogTypBordereau := QOrg.FindField('POG_TYPEPERIOD').AsString
     DucsOrganisme.PogTypBordereau := PogTypBordereau1
   else
     if (PeriodEtat = DucsOrganisme.PogPeriod2) then
// PT1       DucsOrganisme.PogTypBordereau := QOrg.FindField('POG_TYPEAUTPERIOD').AsString
       DucsOrganisme.PogTypBordereau := PogTypBordereau2
     else
       DucsOrganisme.PogTypBordereau := '';
  end;
  Ferme(QOrg);

  {ducs personnalisée}
  personnalisee := False;
  if (PogNature = '300') and
     (PogBoolST = True) and
     (PogSTPos <> 0) and
     (PogSTLong <> 0) then
    personnalisee := True;

  if PogNature = '600' then
  {MSA  : Récup du N° MSA}
  begin
    QEtab := OpenSql('SELECT ETB_NUMMSA' +
                      ' FROM ETABCOMPL ' +
                      ' WHERE ETB_ETABLISSEMENT = "' + DucsParametres .Etab + '"', True);
    if not QEtab.eof then
    begin
      ForceNumerique(QEtab.FindField('ETB_NUMMSA').AsString, NumMsa);
    end;
    DucsVariables.EtSiret := copy(NumMsa, 1, 14);
    Ferme(QEtab);
  end;

  {Formatage code période abrégé}
  DecodeDate(StrToDate(FinPer), aa, mm, jj);
  if (aa < 2000) then
    aa := aa - 1900
  else
    aa := aa - 2000;

  if (VH_Paie.PGDecalage = TRUE) or (VH_Paie.PGDecalagePetit = TRUE) then
  {Décalage de paie}
  begin
    if (mm = 12) then
      aa := aa + 1;
    if (aa = 100) then
      aa := 0;
    mm := mm + 1;
    if (mm = 13) then
      mm := 1;
  end;

  AbregePeriod := '0000';

  if (PeriodEtat = 'M') then
    AbregePeriod := FormatFloat('00', aa) +
                    FormatFloat('0', ((mm - 1) div 3) + 1) +
                    FormatFloat('0', (((mm - 1) - (((mm - 1) div 3) * 3) + 1)));

  if (PeriodEtat = 'T') then
    AbregePeriod := FormatFloat('00', aa) +
                    FormatFloat('0', ((mm - 1) div 3) + 1) + '0';
  if (PeriodEtat = 'A') then
    AbregePeriod := FormatFloat('00', aa) + '00';

  { Calcul Date exigibilité }
  DatExigible := PlusDate(DebutdeMois(StrToDate(Finper)),1,'M');
  if (DucsVariables.JExigible = 30) or (DucsVariables.JExigible = 31) then
    DatExigible := FinDeMois(DatExigible)
  else
  begin
     DecodeDate (FindeMois(DatExigible), aa, mm, jj);
     if (mm = 2) and ((DucsVariables.JExigible = 28) or (DucsVariables.JExigible = 29)) then
       DatExigible := FinDeMois(DatExigible)
     else
     begin
       if (DucsVariables.JExigible > 31) then
       // après la fin de mois suivant
       begin
         DatExigible := FinDeMois(DatExigible) ;
         DucsVariables.JExigible := DucsVariables.JExigible-31+1;
       end;
       DatExigible := PlusDate(DatExigible,DucsVariables.JExigible-1,'J');
     end;
  end;

  { Calcul Date limite de dépôt }
  DatLimDepot := PlusDate(DebutdeMois(StrToDate(Finper)),1,'M');
  if (DucsVariables.JLimDepot = 30) or (DucsVariables.JLimDepot = 31) then
    DatLimDepot := FinDeMois(DatLimDepot)
  else
  begin
     DecodeDate (FindeMois(DatLimDepot), aa, mm, jj);
     if (mm = 2) and ((DucsVariables.JLimDepot = 28) or (DucsVariables.JLimDepot = 29)) then
       DatLimDepot := FinDeMois(DatLimDepot)
     else
     begin
       if (DucsVariables.JLimDepot > 31) then
       // après la fin de mois suivant
       begin
         DatLimDepot := FinDeMois(DatLimDepot) ;
         DucsVariables.JLimDepot := DucsVariables.JLimDepot-31+1;
       end;
       DatLimDepot := PlusDate(DatLimDepot,DucsVariables.JLimDepot-1,'J');
     end;
  end;

  {Calcul date de règlement}
  DatReglt := PlusDate(DebutdeMois(StrToDate(Finper)),1,'M');
  if (DucsVariables.JReglement = 30) or (DucsVariables.JReglement = 31) then
    DatReglt := FinDeMois(DatReglt)
  else
  begin
     DecodeDate (FindeMois(DatReglt), aa, mm, jj);
     if (mm = 2) and ((DucsVariables.JReglement = 28) or (DucsVariables.JReglement = 29)) then
       DatReglt := FinDeMois(DatReglt)
     else
     begin
       if (DucsVariables.JReglement > 31) then
       // après la fin de mois suivant
       begin
         DatReglt := FinDeMois(DatReglt) ;
         DucsVariables.JReglement := DucsVariables.JReglement-31+1;
       end;
       DatReglt := PlusDate(DatReglt,DucsVariables.JReglement-1,'J');
     end;
  end;

  if (DucsVariables.DatePaye = 0) then
    DucsVariables.DatePaye := IDate1900;

  {Récup. de la monnaie de tenue}
  if VH_Paie.PGTenueEuro = FALSE then
    Monnaie := 'FRF' // Franc
  else
    Monnaie := 'EUR'; // Euro

  {Formatage ligne optique (elle peut être déjà renseignée au niveau
   du paramètrage de l'organisme)}
  LigneOptique := '';
  if (DucsOrganisme.PogLgOptique = '') then
  begin
    if (DucsOrganisme.PogNature = '100') then
    { URSSAF }
    begin
      LigneOptique := Copy(DucsOrganisme.PogNumInterne, 1, length(DucsOrganisme.PogNumInterne));
      if (length(DucsOrganisme.PogNumInterne) < 30) then
      begin
        for i := length(DucsOrganisme.PogNumInterne) + 1 to 30 do
        begin
          LigneOptique := LigneOptique + '0';
        end;
      end;
    end;

    if (DucsOrganisme.PogNature = '200') then
    { ASSEDIC }
    begin
      LigneOptique := 'S2' + Copy(DucsOrganisme.PogNumInterne, 1, length(DucsOrganisme.PogNumInterne));
      if ((length(DucsOrganisme.PogNumInterne) + 2) < 30) then
      begin
        for i := length(DucsOrganisme.PogNumInterne) + 3 to 30 do
        begin
          LigneOptique := LigneOptique + '0';
        end;
      end;
    end;
  end
  else
  begin
    LigneOptique := DucsOrganisme.PogLgOptique;
    if (length(DucsOrganisme.PogLgOptique) < 30) then
    begin
      for i := length(DucsOrganisme.PogLgOptique) + 1 to 30 do
      begin
        LigneOptique := LigneOptique + '0';
      end;
    end;
  end;

  {ALIMENTATION TABLES DUCS}
  BaseCodif   := 0.00;
  MontCodif   := 0.00;
  MontForfait := 0.00;
  MontST      := 0.00;
  IndCalcSt   := false; { ducs personnalisée}
  NoAT        := -1;

  DucsVariables.NumDucs     := 0;
  ligne       := 0;
  DucsVariables.EtabEnCours := DucsParametres.Etab;
  DucsVariables.OrganismeEnCours  := DucsParametres.Organisme;

  TDetail := TCot3.FindFirst([''], [''], TRUE);
  { Gestion des ruptures}
  if (DucsVariables.Rupture = TRUE) then
  begin
    TRup := TRupMere.FindFirst([''], [''], TRUE);
    if (TRup <> nil) then
    begin
      IndRup := False;
      if (TRup.GetValue('ET_SIRET') <> Null) then
        DucsVariables.SiretEnCours := TRup.GetValue('ET_SIRET');

      if (TRup.GetValue('ET_APE') <> Null) then
        DucsVariables.ApeEnCours := TRup.GetValue('ET_APE');

      DucsVariables.EtabEnCours := TRup.GetValue('POG_ETABLISSEMENT');
      DucsVariables.OrganismeEnCours := TRup.GetValue('POG_ORGANISME');

      if (TRup.GetValue('POG_NUMINTERNE') <> Null) then
        DucsVariables.NumIntEnCours := copy(TRup.GetValue('POG_NUMINTERNE'), 1, 30);

      if (TRup.GetValue('POG_LGOPTIQUE') <> Null) then
        LGOptiqueEncours := TRup.GetValue('POG_LGOPTIQUE');

      if (DucsVariables.RupGpInt = TRUE) then
      begin
        if (TRup.GetValue('PGI_GROUPE') <> Null) then
          DucsVariables.GpIntEnCours := TRup.GetValue('PGI_GROUPE');
      end;
    end;
  end; { fin if (Rupture = TRUE)}

  if (Tdetail <> nil) then
  begin
    if (DucsVariables.RupSiret = TRUE) then
      if (TDetail.GetValue('SIRET') <> Null) then
        DucsVariables.Wsiret := TDetail.GetValue('SIRET');
    if (DucsVariables.RupApe = TRUE) then
      if (TDetail.GetValue('APE') <> Null) then
        DucsVariables.WApe := TDetail.GetValue('APE');
    if (DucsVariables.RupNumInt = TRUE) then
      if (TDetail.GetValue('NUMINTERNE') <> Null) then
        DucsVariables.WNumInt := TDetail.GetValue('NUMINTERNE');
    if (DucsVariables.RupGpInt = TRUE) then
      if (TDetail.GetValue('GROUPE') <> Null) then
        DucsVariables.WGpInt := TDetail.GetValue('GROUPE');
  end; { fin if (Tdetail <> NIl)}

  while ((fctRupture(DucsVariables,TDetail, 0) = TRUE) or (DucsVariables.IndNeant2 = TRUE)) and
        ((DucsVariables.Rupture = False) or ((DucsVariables.Rupture = TRUE) and
         (TRup <> nil))) do
  begin
  {Alimentation DUCSENTETE}
    if (DucsVariables.NumDucs = 0) then
    begin
      TOB_Lignes := TOB.Create('La Déclaration', nil, -1);
    end;
    {Recherche s'il n'existe pas déjà une ducs avec la même clé
     dans ce cas on incrémente le n° de ducs pour ne pas rencontrer de
     problème de clé duppliquée. (cas de l'utilisation des ruptures)}
    while ExisteSQL('SELECT PDU_ETABLISSEMENT,PDU_ORGANISME,PDU_DATEDEBUT,' +
      'PDU_DATEFIN,PDU_NUM FROM DUCSENTETE WHERE ' +
      'PDU_ETABLISSEMENT="' + DucsVariables.EtabEnCours + '" AND ' +
      'PDU_ORGANISME="' + DucsParametres.Organisme + '" AND ' +
      'PDU_NUM=' + IntToStr(DucsVariables.NumDucs) + '') do 
    begin
      DucsVariables.NumDucs := DucsVariables.NumDucs + 1;
      ligne := 0;
    end;

//  Centre Payeur
    QOrg := OpenSql('SELECT' +
      ' POG_CPINFO,POG_CENTREPAYEUR' +
      ' FROM ORGANISMEPAIE ' +
      ' WHERE POG_ETABLISSEMENT= "' + DucsVariables.EtabEnCours + '" AND ' +
      ' POG_ORGANISME = "' + DucsParametres.Organisme + '"', True);
    if not QOrg.eof then
    begin
      PogCPInfo := QOrg.FindField('POG_CPINFO').AsString;
      PogCentrePayeur := QOrg.FindField('POG_CENTREPAYEUR').AsString;
    end;
    Ferme(QOrg);

    TOB_FilleTete := TOB.create('DUCSENTETE', TOB_Lignes, -1);
    TOB_FilleTete.PutValue('PDU_ETABLISSEMENT', DucsVariables.EtabEnCours);
    TOB_FilleTete.PutValue('PDU_ORGANISME', DucsPArametres.Organisme);
    TOB_FilleTete.PutValue('PDU_DATEDEBUT', StrToDate(DebPer));
    TOB_FilleTete.PutValue('PDU_DATEFIN', StrToDate(FinPer));
    TOB_FilleTete.PutValue('PDU_NUM', DucsVariables.NumDucs);

    if ((DucsVariables.Rupture = TRUE) and (DucsVariables.SiretEnCours = '')) or
      ((DucsVariables.Rupture <> TRUE) and (DucsVariables.EtSiret = '')) then
      Anomalie := 'Non Renseigné';

    if (DucsVariables.Rupture = TRUE) then
      TOB_FilleTete.PutValue('PDU_SIRET', DucsVariables.SiretEnCours)
    else
      TOB_FilleTete.PutValue('PDU_SIRET', DucsVariables.EtSiret);

    if (DucsVariables.Rupture = TRUE) then
      TOB_FilleTete.PutValue('PDU_APE', DucsVariables.ApeEnCours)
    else
      TOB_FilleTete.PutValue('PDU_APE', DucsVariables.EtApe);

    if (DucsVariables.RupGpInt = TRUE) then
      TOB_FilleTete.PutValue('PDU_GROUPE', DucsVariables.GpIntEnCours)
    else
      TOB_FilleTete.PutValue('PDU_GROUPE', '');

    if (DucsVariables.Rupture = TRUE) then
      TOB_FilleTete.PutValue('PDU_NUMERO', DucsVariables.NumIntEnCours)
    else
      TOB_FilleTete.PutValue('PDU_NUMERO',DucsOrganisme.PogNumInterne);

    TOB_FilleTete.PutValue('PDU_ABREGEPERIODE', AbregePeriod);
    TOB_FilleTete.PutValue('PDU_DATEEXIGIBLE', DatExigible);
    TOB_FilleTete.PutValue('PDU_DATELIMDEPOT', DatLimDepot);
    TOB_FilleTete.PutValue('PDU_DATEREGLEMENT', DatReglt);

    if (DucsVariables.IndNeant2 = TRUE) then
      TOB_FilleTete.PutValue('PDU_PAIEMENT', IDate1900)
    else
      TOB_FilleTete.PutValue('PDU_PAIEMENT', DucsVariables.DatePaye);

    TOB_FilleTete.PutValue('PDU_CESSATION', IDate1900);
    TOB_FilleTete.PutValue('PDU_CONTINUATION', IDate1900);
    TOB_FilleTete.PutValue('PDU_SUSPENSION', '-');
    TOB_FilleTete.PutValue('PDU_MAINTIENT', '-');
    TOB_FilleTete.PutValue('PDU_DECLARANT', DucsDeclarant.RaisonSoc);
    TOB_FilleTete.PutValue('PDU_DECLARANTSUITE', DucsDeclarant.ContactDecla);
    TOB_FilleTete.PutValue('PDU_TELEPHONEDECL', DucsDeclarant.TelDeclarant);
    TOB_FilleTete.PutValue('PDU_FAXDECLARANT', DucsDeclarant.FaxDeclarant);
    if (DucsVariables.Rupture = TRUE) then
      TOB_FilleTete.PutValue('PDU_LIGNEOPTIQUE', LGOptiqueEnCours)
    else
      TOB_FilleTete.PutValue('PDU_LIGNEOPTIQUE', LigneOptique);
    TOB_FilleTete.PutValue('PDU_ACOMPTES', 0);
    TOB_FilleTete.PutValue('PDU_REGULARISATION', 0);
    TOB_FilleTete.PutValue('PDU_MONNAIETENUE', Monnaie);

    CalculEffectifs(DebPer,FinPer,DucsVariables,DucsParametres,DucsEffectifs,TInst,DucsOrganisme.PogTypBordereau, DucsOrganisme.PogNature );

    if (DucsVariables.IndNeant2 = TRUE) then
      TOB_FilleTete.PutValue('PDU_NBSALFPE', 0)
    else
      TOB_FilleTete.PutValue('PDU_NBSALFPE', DucsEffectifs.EffTot);

    TOB_FilleTete.PutValue('PDU_NBSALHAP', 0);
    TOB_FilleTete.PutValue ('PDU_NBSALQ920',DucsEffectifs.EffHom);
    TOB_FilleTete.PutValue ('PDU_NBSALQ921',DucsEffectifs.EffFem);
    TOB_FilleTete.PutValue ('PDU_NBSALQ922',DucsEffectifs.EffCDIH);
    TOB_FilleTete.PutValue ('PDU_NBSALQ923',DucsEffectifs.EffCDIF);
    TOB_FilleTete.PutValue ('PDU_NBSALQ924',DucsEffectifs.EffAppH);
    TOB_FilleTete.PutValue ('PDU_NBSALQ925',DucsEffectifs.EffAppF);
    TOB_FilleTete.PutValue ('PDU_NBSALQ926',DucsEffectifs.Eff65H);
    TOB_FilleTete.PutValue ('PDU_NBSALQ927',DucsEffectifs.Eff65F);
    TOB_FilleTete.PutValue ('PDU_NBSALQ928',DucsEffectifs.EffCadH);
    TOB_FilleTete.PutValue ('PDU_NBSALQ929',DucsEffectifs.EffCadF);
    TOB_FilleTete.PutValue ('PDU_NBSALQ934',DucsEffectifs.EffProfH);
    TOB_FilleTete.PutValue ('PDU_NBSALQ935',DucsEffectifs.EffProfF);
    TOB_FilleTete.PutValue ('PDU_NBSALQ936',DucsEffectifs.EffCDDH);    // effectif 930 du CdC
    TOB_FilleTete.PutValue ('PDU_NBSALQ937',DucsEffectifs.EffCDDF);    // effectif 931 du CdC
    TOB_FilleTete.PutValue ('PDU_NBSALQ962',DucsEffectifs.EffCNEH);     // effectif 932 du CdC
    TOB_FilleTete.PutValue ('PDU_NBSALQ963',DucsEffectifs.EffCNEF);     // effectif 933 du CdC
    TOB_FilleTete.PutValue('PDU_TOTHOMMES',DucsEffectifs.EffHom);
    TOB_FilleTete.PutValue('PDU_TOTFEMMES',DucsEffectifs.EffFem);
    TOB_FilleTete.PutValue('PDU_TOTAPPRENTI',DucsEffectifs.EffApp);
    TOB_FilleTete.PutValue('PDU_TOTOUVRIER',DucsEffectifs.EffOuv);
    TOB_FilleTete.PutValue('PDU_TOTETAM',DucsEffectifs.EffEtam);
    TOB_FilleTete.PutValue('PDU_TOTCADRES',DucsEffectifs.EffCad);
    TOB_FilleTete.PutValue('PDU_DATEPREVEL', IDate1900);
    TOB_FilleTete.PutValue('PDU_EMETTSOC',DucsDeclarant.Emetteur);
    TOB_FilleTete.PutValue ('PDU_NBSALQ944',0);
    TOB_FilleTete.PutValue ('PDU_NBSALQ945',0);
    TOB_FilleTete.PutValue ('PDU_NBSALQ960',0);
    TOB_FilleTete.PutValue ('PDU_NBSALQ961',0);
    TOB_FilleTete.PutValue ('PDU_NBSALQ964',0);
    TOB_FilleTete.PutValue ('PDU_NBSALQ965',0);
    TOB_FilleTete.PutValue ('PDU_NBSALQ966',0);
    TOB_FilleTete.PutValue ('PDU_NBSALQ967',0);
    TOB_FilleTete.PutValue ('PDU_ECARTZE1','-');
    TOB_FilleTete.PutValue ('PDU_ECARTZE2','-');
    TOB_FilleTete.PutValue ('PDU_ECARTZE3','-');
    TOB_FilleTete.PutValue ('PDU_ECARTZE4','-');
    TOB_FilleTete.PutValue ('PDU_ECARTZE5','-');
    TOB_FilleTete.PutValue ('PDU_ECARTZE6','-');
    TOB_FilleTete.PutValue ('PDU_ECARTZE7','-');
    TOB_FilleTete.PutValue ('PDU_ECARTZE8','-');
    if (DucsParametres.DucsDoss = True) then
      TOB_FilleTete.PutValue('PDU_DUCSDOSSIER', 'X')
    else
      TOB_FilleTete.PutValue('PDU_DUCSDOSSIER', '-');

// Centre Payeur
    if (PogNature <> '300') then
      TOB_FilleTete.PutValue('PDU_CENTREPAYEUR', '')
    else
    begin
      if (DucsOrganisme.PogCpInfo = '') or (DucsOrganisme.PogCpInfo = '001') or (DucsOrganisme.PogCpInfo = '002') then
      { type info = libre ou n° affiliation ou n° interne }
        TOB_FilleTete.PutValue('PDU_CENTREPAYEUR', DucsOrganisme.PogCentrePayeur)
      else
        if (DucsOrganisme.PogCpInfo = '003') then
        { type info = siret }
        begin
          if (DucsVariables.Rupture = TRUE) then
            TOB_FilleTete.PutValue('PDU_CENTREPAYEUR', DucsVariables.SiretEnCours)
          else
            TOB_FilleTete.PutValue('PDU_CENTREPAYEUR', DucsVariables.EtSiret);
        end
        else
          if (DucsOrganisme.PogCpInfo = '004') then
          { type info = ape }
          begin
            if (DucsVariables.Rupture = TRUE) then
              TOB_FilleTete.PutValue('PDU_CENTREPAYEUR', DucsVariables.ApeEnCours)
            else
              TOB_FilleTete.PutValue('PDU_CENTREPAYEUR', DucsVariables.EtApe);
          end
          else
            if (DucsOrganisme.PogCpInfo = '005') then
            { type info = groupe interne }
            begin
              if (DucsVariables.RupGpInt = TRUE) then
                TOB_FilleTete.PutValue('PDU_CENTREPAYEUR', DucsVariables.GpIntEnCours)
              else
                TOB_FilleTete.PutValue('PDU_CENTREPAYEUR', '');
            end;
    end;

    TOB_FilleTete.PutValue ('PDU_TYPBORDEREAU',DucsOrganisme.PogTypBordereau);

//?????????    if (Tdetail = nil) or (IndNeant2 = TRUE) then
    { il s'agit d'une Ducs Néant}
// ?????????      IndNeant := TRUE;

  {Alimentation DUCSDETAIL}
  {-----------------------}
  {
    Traitement des ruptures
  }
    while (TDetail <> nil) and
      (DucsVariables.IndNeant2 = FALSE) and
      (fctRupture(DucsVariables,TDetail, 0) = TRUE) do
    begin
    {arrondi : récup. paramétrage organisme}
      DucsVariables.BaseArr := DucsVariables.WBaseArr;
      DucsVariables.MontArr := DucsVariables.WMontArr;

    {affectation des variables de traitement des lignes de détail}
      CodifLue := TDetail.GetValue('CODIFICATION');
      CodifEnCours := TDetail.GetValue('CODIFICATION');
// loi fillon
      if (TDetail.GetValue('REGIMEALSACE') =  'X') then
        AlsaceMoselle := true
      else
        AlsaceMoselle := false;

      if (AlsaceMoselle) then
        QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, PDP_LIBELLESUITE,' +
        'PDP_TYPECOTISATION, PDP_BASETYPARR,PDP_MTTYPARR, PDP_CONDITION' +
        ' FROM DUCSPARAM ' +
        ' WHERE ##PDP_PREDEFINI## PDP_CODIFALSACE = "' + CodifEnCours + '"' +
        ' ORDER BY PDP_PREDEFINI', True)
      else
        QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, PDP_LIBELLESUITE,' +
        'PDP_TYPECOTISATION, PDP_BASETYPARR,PDP_MTTYPARR, PDP_CONDITION' +
        ' FROM DUCSPARAM ' +
        ' WHERE ##PDP_PREDEFINI## PDP_CODIFICATION = "' + CodifEnCours + '"' +
        ' ORDER BY PDP_PREDEFINI', True);
      Predef := '';
{$IFDEF EAGLCLIENT}
      SortDos := 0;
      KK := 0;
{$ENDIF}
      while not QDucsParam.eof do
      begin
        Predef := QDucsparam.FindField('PDP_PREDEFINI').AsString;
        if (Predef = 'DOS') then
        begin
{$IFDEF EAGLCLIENT}
          SortDos := 1;
{$ENDIF}
          break;
        end;
{$IFDEF EAGLCLIENT}
        KK := KK + 1;
{$ENDIF}
        QDucsParam.Next;
      end;
{$IFDEF EAGLCLIENT}
      if SortDos = 1 then
        MaTob := QDucsParam.Detail[QDucsParam.Detail.count - 1]
      else
        MaTob := QDucsParam.Detail[KK - 1];
{$ENDIF}
      if (Predef <> '') then
      begin
{$IFNDEF EAGLCLIENT}
        TypeCotis := QDucsParam.FindField('PDP_TYPECOTISATION').AsString;
{$ELSE}
        TypeCotis := MaTob.GetValue('PDP_TYPECOTISATION');
{$ENDIF}
      end;
//    ferme (QDucsParam);
      TauxLu := Arrondi(TDetail.GetValue('TAUXGLOBAL'), 5);

// loi fillon
      if (TypeCotis = 'M') then
        TauxLu := 0.00;
// loi TEPA
      if (TypeCotis = 'C') then       // base & montant
        TauxLu := 0.00;
      if (TypeCotis = 'D') then       // Quantité en base & mt
        TauxLu := 0.00;

      TauxEnCours := Arrondi(TDetail.GetValue('TAUXGLOBAL'), 5);
// loi fillon
      if (TypeCotis = 'M') then
        TauxEnCours := 0.00;
// loi TEPA
      if (TypeCotis = 'C') then       // base & montant
        TauxEnCours := 0.00;

      if (TypeCotis = 'D') then       // quantité en base & mt
        TauxEnCours := 0.00;

      ATLu := Arrondi(TDetail.GetValue('TAUXAT'), 5);
      ATEnCours := Arrondi(TDetail.GetValue('TAUXAT'), 5);
      BaseLigne := Arrondi(TDetail.GetValue('BASECOT'), 5);
// loi fillon
      if (TypeCotis = 'M') then
        BaseLigne := 0.00;

      MontLigne := Arrondi(TDetail.GetValue('MTGLOBAL'), 5);
      SalLu := TDetail.GetValue('SALARIE');
      OrgLu := TDetail.GetValue('ORGANISME');
      if (personnalisee = False) then
        if (TDetail.GetValue('INSTITUTION') <> Null) then
          CodInstLu := TDetail.GetValue('INSTITUTION');
      if (DucsVariables.RupSiret = True) then
        if (TDetail.GetValue('SIRET') <> Null) then
          WSiretLu := TDetail.GetValue('SIRET');
      if (DucsVariables.RupNumInt = TRUE) then
        if (TDetail.GetValue('NUMINTERNE') <> Null) then
          WNoIntLu := TDetail.GetValue('NUMINTERNE');
      if (DucsVariables.RupApe = TRUE) then
        if (TDetail.GetValue('APE') <> Null) then
          WApeLu := TDetail.GetValue('APE');
      if (DucsVariables.RupGpInt = TRUE) then
        if (TDetail.GetValue('GROUPE') <> Null) then
          WGpLu := TDetail.GetValue('GROUPE');
      OrgEnCours := TDetail.GetValue('ORGANISME');
        if (TDetail.GetValue('INSTITUTION') <> Null) then
          CodInstEnCours := TDetail.GetValue('INSTITUTION');

      if (PogNature = '300') then
    {IRC : récup de l'institution encours }
        if (copy(IntToStr(NoInstEnCours), 1, 2) <> copy(CodifLue, 1, 2)) or
          ((NoInstEnCours = 0) and
          (copy(IntToStr(NoInstEnCours), 1, 1) = copy(CodifLue, 1, 1))) then
        begin
          if IsNumeric(copy(CodifLue, 1, 2)) then
            NoInstEnCours := StrToInt(copy(CodifLue, 1, 2))
          else
          begin
            Anomalie := 'erreur d''affectation';
          end;
        end;
      Effectif := 0;
      FreeAndNil(TEffM);
      SalEnCours := ' ';
      if (AlsaceMoselle) then
        QDucsParamAT := OpenSQL('SELECT PDP_PREDEFINI,PDP_CODIFALSACE FROM DUCSPARAM ' +
                      'WHERE ##PDP_PREDEFINI## PDP_TYPECOTISATION="A" AND ' +
                      'PDP_CODIFALSACE LIKE "' + Copy(CodifEnCours, 1, 6) + '%"' +
                      ' ORDER BY PDP_PREDEFINI',
                      True)
      else
        QDucsParamAT := OpenSQL('SELECT PDP_PREDEFINI,PDP_CODIFICATION FROM DUCSPARAM ' +
                      'WHERE ##PDP_PREDEFINI## PDP_TYPECOTISATION="A" AND ' +
                      'PDP_CODIFICATION LIKE "' + Copy(CodifEnCours, 1, 6) + '%"' +
                      ' ORDER BY PDP_PREDEFINI',
                      True);
      CodifAT := '';
      while not QDucsParamAT.Eof do
      begin
        if (AlsaceMoselle) then
          CodifAT := QDucsparamAT.FindField('PDP_CODIFALSACE').AsString
        else
          CodifAT := QDucsparamAT.FindField('PDP_CODIFICATION').AsString;
        Predef := QDucsparamAT.FindField('PDP_PREDEFINI').AsString;
        if (Predef = 'DOS') then break;
        QDucsParamAT.Next;
      end;
      Ferme(QDucsParamAT);

      while ((IndRup = False) and
        (CodifLue = CodifEnCours) and
        (ARRONDI(TauxLu, 5) = ARRONDI(TauxEnCours, 5)) and
        (((CodifAT <> '') and (ARRONDI(ATLu, 5) = ARRONDI(ATEnCours, 5)) and
        (Copy(CodifEnCours, 1, 6) = Copy(CodifAT, 1, 6))) or
        (Copy(CodifEnCours, 1, 6) <> Copy(CodifAT, 1, 6))) and
        (TDetail <> nil)) do
    {pour une même codification: traitement d'une ligne}
      begin
    {(Cumuls)}
        BaseCodif := BaseCodif + BaseLigne;
        MontCodif := MontCodif + MontLigne;
        MontForfait := MontLigne;
        if (SalEnCours <> SalLu) then
      {(Incrémentation effectif)}
        begin
          if (TypeCotis = 'M') then
          begin
            if (TEffM = nil) then
              TEffM := TOB.Create('effectif forfait', nil, -1);

            TEffMFille := TEffM.FindFirst(['SALARIE'], [SalLu], TRUE);
            if (TEffMFille = nil) then
            begin
              TEffMFille := TOB.Create('', TEffM, -1);
              TEffMFille.AddChampSupValeur('SALARIE', SalLu, FALSE);
            end;
            if (MontLigne <> 0) then //PT46 FQ 11011
              effectif := TEffM.Detail.Count;
          end
          else
            Effectif := Effectif + 1;
          SalEnCours := SalLu;
        end;

      {ligne de cotisation suivante}
        TDetail := TCot3.FindNext([''], [''], TRUE);
        if (TDetail <> nil) then
      {affectation des variables de traitement des lignes de détail}
        begin
          CodifLue := TDetail.GetValue('CODIFICATION');
          TauxLu := Arrondi(TDetail.GetValue('TAUXGLOBAL'), 5);
// loi fillon
          if (TypeCotis = 'M') then
            TauxLu := 0.00;
// loi TEPA
          if (TypeCotis = 'C') then     //  base & montant
             TauxLu := 0.00;

          if (TypeCotis = 'D') then     // quantité & montant
            TauxLu := 0.00;

          ATLu := Arrondi(TDetail.GetValue('TAUXAT'), 5);
          BaseLigne := Arrondi(TDetail.GetValue('BASECOT'), 5);
// loi fillon
          if (TypeCotis = 'M') then
            BaseLigne := 0.00;

          MontLigne := Arrondi(TDetail.GetValue('MTGLOBAL'), 5);
          Sallu := TDetail.GetValue('SALARIE');
          OrgLu := TDetail.GetValue('ORGANISME');
          if (personnalisee = False) then
            if (TDetail.GetValue('INSTITUTION') <> Null) then
              CodInstLu := TDetail.GetValue('INSTITUTION');
          IndRup := false;
          if (DucsVariables.RupSiret = True) then
          begin
            if (TDetail.GetValue('SIRET') <> Null) then
              WSiretLu := TDetail.GetValue('SIRET');
            if (WSiretLu <> DucsVariables.SiretEnCours) then IndRup := True;
          end;
          if (DucsVariables.RupNumInt = TRUE) then
          begin
            if (TDetail.GetValue('NUMINTERNE') <> Null) then
              WNoIntLu := TDetail.GetValue('NUMINTERNE');
            if (WNoIntLu <> DucsVariables.NumIntEnCours) then IndRup := True;
          end;
          if (DucsVariables.RupApe = TRUE) then
          begin
            if (TDetail.GetValue('APE') <> Null) then
              WApeLu := TDetail.GetValue('APE');
            if (WApeLu <> DucsVariables.ApeEnCours) then IndRup := True;
          end;
          if (DucsVariables.RupGpInt = TRUE) then
          begin
            if (TDetail.GetValue('GROUPE') <> Null) then
              WGpLu := TDetail.GetValue('GROUPE');
            if (WGpLu <> DucsVariables.GpIntEnCours) then IndRup := True;
          end;
        end;
      end; { fin pour une même codification: traitement d'une ligne}

    {Mise à jour  les lignes}
      if (CodifEnCours = Copy(CodifAT, 1, 6) + 'D') then
    {Traitement TAUX AT}
      begin
        NoAT := NoAT + 1;
        if (NoAT > 9) then NoAT := 0; { Au cas où + de 10 tx AT Différent}
        if (AlsaceMoselle) then
          QDucsParamAT := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, PDP_LIBELLESUITE,' +
                                  'PDP_TYPECOTISATION' +
                                  ' FROM DUCSPARAM ' +
                                  ' WHERE ##PDP_PREDEFINI## PDP_CODIFALSACE = "' + CodifAT + '" ' +
                                  'ORDER BY PDP_PREDEFINI', True)
        else
          QDucsParamAT := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, PDP_LIBELLESUITE,' +
                                  'PDP_TYPECOTISATION' +
                                  ' FROM DUCSPARAM ' +
                                  ' WHERE ##PDP_PREDEFINI## PDP_CODIFICATION = "' + CodifAT + '" ' +
                                  'ORDER BY PDP_PREDEFINI', True);
        Predef := '';
{$IFDEF EAGLCLIENT}
        SortDos := 0;
        KK := 0;
{$ENDIF}
        while not QDucsParamAT.eof do
        begin
          Predef := QDucsparamAT.FindField('PDP_PREDEFINI').AsString;
          if (Predef = 'DOS') then
          begin
{$IFDEF EAGLCLIENT}
            SortDos := 1;
{$ENDIF}
            break;
          end;
{$IFDEF EAGLCLIENT}
          KK := KK + 1;
{$ENDIF}
          QDucsParamAT.Next;
        end;

        if (Predef <> '') then
        begin
          AlimDetail(DebPer,FinPer,TOB_FilleLignes, TOB_FilleTete, ligne, DucsParametres, DucsVariables,DucsOrganisme.PogNumInterne);
          TOB_FilleLignes.PutValue('PDD_CODIFICATION', Copy(CodifAT, 1, 2) +
            IntToStr(NoAT) + Copy(CodifAT, 4, 4));
          Cod100 := Copy(CodifAT, 1, 2) + IntToStr(NoAT) + Copy(CodifAT, 4, 4);
{$IFDEF EAGLCLIENT}
          if SortDos = 1 then MaTobAT := QDucsParamAT.Detail[QDucsParam.Detail.count - 1]
          else MaTobAT := QDucsParamAT.Detail[KK - 1];
          TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
            MaTobAT.GetValue('PDP_TYPECOTISATION'));
          TOB_FilleLignes.PutValue('PDD_LIBELLE',
            MaTobAT.GetValue('PDP_LIBELLE'));
          TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
            MaTobAT.GetValue('PDP_LIBELLESUITE'));
{$ELSE}
          TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
            QDucsParamAT.FindField('PDP_TYPECOTISATION').AsString);
          TOB_FilleLignes.PutValue('PDD_LIBELLE',
            QDucsParamAT.FindField('PDP_LIBELLE').AsString);
          TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
            QDucsParamAT.FindField('PDP_LIBELLESUITE').AsString);
{$ENDIF}
          TOB_FilleLignes.PutValue('PDD_EFFECTIF', 0);
          TOB_FilleLignes.PutValue('PDD_BASECOTISATION', 0.00);
          TOB_FilleLignes.PutValue('PDD_TAUXCOTISATION', Arrondi(AtEnCours, 5));
          TOB_FilleLignes.PutValue('PDD_MTCOTISAT', 0.00);
          TOB_FilleLignes.PutValue('PDD_COMURBAINE', '    ');
          TOB_FilleLignes.PutValue('PDD_CONDITION', 'SCS');
          TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
            copy(Cod100, DucsVariables.CodifPosDebut, DucsVariables.CodifLongEdit));
        end;
// d PT4
{$IFDEF EAGLCLIENT}
        if (Assigned(MaTobAT)) then
          FreeandNil (MaTobAT);
{$ENDIF}
// f PT4
        Ferme(QDucsParamAT);
      end; {If (CodifEnCours = '1A0100D')}
    {fin traitement Taux AT}

      if (Predef <> '') then
      begin
        if (MontCodif <> 0.00) or
          ((MontCodif = 0.00) and
           ((Arrondi(AtEnCours, 5) <> 0.00) and
{$IFNDEF EAGLCLIENT}
          (QDucsParam.FindField('PDP_TYPECOTISATION').AsString = 'M')) or
          (QDucsParam.FindField('PDP_TYPECOTISATION').AsString = 'I'))
{$ELSE}
          (MaTob.GetValue('PDP_TYPECOTISATION') = 'M')) or
          (MaTob.GetValue('PDP_TYPECOTISATION') = 'I'))
{$ENDIF}
        then
      {Une ligne n'est éditée que si le mt est différent de zéro ou s'il
       est à zéro et que le taux est différent de 0 et
       s'il s'agit d'une ligne "M"ontant ou  'I"ntitulé}
        begin
          AlimDetail(DebPer,FinPer,TOB_FilleLignes, TOB_FilleTete, ligne, DucsParametres, DucsVariables,DucsOrganisme.PogNumInterne);
          if (Copy(CodifEnCours, 1, 6) = Copy(CodifAT, 1, 6)) then
        {CAS GENERAL URSSAF}
          begin
            if (Copy(CodifEnCours, 7, 1) = 'D') then
          { 1AxyyyD }
              TOB_FilleLignes.PutValue('PDD_CODIFICATION',
                Copy(CodifEnCours, 1, 2) +
                IntToStr(NoAT) +
                Copy(CodifEnCours, 4, 7))
            else
            begin
          {1AxyyyP}
              TAT := TOB_FilleTete.FindFirst(['PDD_TAUXCOTISATION',
                'PDD_TYPECOTISATION',
                  'PDD_ETABLISSEMENT',
                  'PDD_ORGANISME',
                  'PDD_NUM',
                  'PDD_DATEDEBUT',
                  'PDD_DATEFIN'],
                  [ATEnCours,
                'A',
                  DucsVariables.EtabEnCours,
                  DucsParametres.organisme,
                  DucsVariables.NumDucs,
                  StrToDate(DebPer),
                  StrToDate(FinPer)],
                  TRUE);

              if (TAT <> nil) then
                if (copy(TAT.GetValue('PDD_CODIFICATION'), 3, 1)) = '' then
                  NoATP := 0
                else
                  NoATP := NOAT;
              TOB_FilleLignes.PutValue('PDD_CODIFICATION',
                Copy(CodifEnCours, 1, 2) +
                IntToStr(NoATP) +
                Copy(CodifEnCours, 4, 7));
            end;
          end
          else
            if (PogNature = '300') then
         {Codif IRC : La codif est modifée en fct de la rupture
                             sur institution }
            begin
              NomChamp[1] := 'PDD_CODIFICATION';
              ValeurChamp[1] := ColleZeroDevant(NoInstEnCours, 2) + Copy(CodifEnCours, 3, 5);
              NomChamp[2] := 'PDD_NUMORDRE';
              ValeurChamp[2] := ligne;

              TobSTExiste := TOB_FilleTete.FindFirst(NomChamp, ValeurChamp, TRUE);
              while (TobStExiste <> nil) do
              begin
                ValChST[1] := ColleZeroDevant(NoInstEnCours + 1, 2) + 'ZZZZZ';
                NomChampST[1] := 'PDD_CODIFICATION';
                if (Personnalisee = False) or
                  ((Personnalisee = True) and
                  (TOB_FilleTete.FindFirst(NomChampST, ValChST, TRUE) <> nil)) then
                begin
                  NoInstEnCours := NoInstEnCours + 1;
                  ValeurChamp[1] := ColleZeroDevant(NoInstEnCours, 2) + Copy(CodifEnCours, 3, 5);
                  TobSTExiste := TOB_FilleTete.FindFirst(NomChamp, ValeurChamp, TRUE);
                end
                else
                begin
                  ligne := ligne + 1;
                  ValeurChamp[2] := ligne;
                  TOB_FilleLignes.PutValue('PDD_NUMORDRE', ligne);
                  TobSTExiste := TOB_FilleTete.FindFirst(NomChamp, ValeurChamp, TRUE);
                end;
              end;
              TOB_FilleLignes.PutValue('PDD_CODIFICATION',
                ColleZeroDevant(NoInstEnCours, 2) + Copy(CodifEnCours, 3, 5))
            end
            else
         {autres natures de DUCS}
              TOB_FilleLignes.PutValue('PDD_CODIFICATION', CodifEnCours);
{$IFDEF EAGLCLIENT}
          if SortDos = 1 then MaTob := QDucsParam.Detail[QDucsParam.Detail.count - 1]
          else MaTob := QDucsParam.Detail[KK - 1];
          TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
            MaTob.GetValue('PDP_TYPECOTISATION'));
          TOB_FilleLignes.PutValue('PDD_LIBELLE',
            MaTob.GetValue('PDP_LIBELLE'));
          TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
            MaTob.GetValue('PDP_LIBELLESUITE'));
           // Type de cotisation
          TypeCotis := MaTob.GetValue('PDP_TYPECOTISATION');
          CSC := 'SCS';
          if (DucsVariables.CondSpec) then
            CSC := MaTob.GetValue('PDP_CONDITION');
          if (CSC = '') then
            CSC := 'SCS';

           // arrondi de la codification
          BaseArrCod := MaTob.GetValue('PDP_BASETYPARR');
          MontArrCod := MaTob.GetValue('PDP_MTTYPARR');
{$ELSE}
          TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
            QDucsParam.FindField('PDP_TYPECOTISATION').AsString);
          TOB_FilleLignes.PutValue('PDD_LIBELLE',
            QDucsParam.FindField('PDP_LIBELLE').AsString);
          TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
            QDucsParam.FindField('PDP_LIBELLESUITE').AsString);

           {Type de cotisation}
          TypeCotis := QDucsParam.FindField('PDP_TYPECOTISATION').AsString;
          CSC := 'SCS';
          if (DucsVariables.CondSpec) then
            CSC := QDucsParam.FindField('PDP_CONDITION').AsString;
          if (CSC = '') then
            CSC := 'SCS';

           {arrondi de la codification}
          BaseArrCod := QDucsParam.FindField('PDP_BASETYPARR').AsString;
          MontArrCod := QDucsParam.FindField('PDP_MTTYPARR').AsString;
{$ENDIF}
// d PT4
{$IFDEF EAGLCLIENT}
          if (Assigned(MaTob)) then
            FreeandNil (MaTob);
{$ENDIF}
// f PT4
          Ferme(QDucsParam);

         {détermination de l'arrondi appliqué}
          if (BaseArrCod <> '') then DucsVariables.BaseArr := BaseArrCod;
          if (MontArrCod <> '') then DucsVariables.MontArr := MontArrCod;

         {arrondi de la base}
          if (DucsVariables.BaseArr <> '') and (DucsVariables.BaseArr <> 'P') then
          begin
            BaseCodif := Int(BaseCodif);
            MontForfait := Int(MontForfait);
            if (DucsVariables.BaseArr = 'S') then BaseCodif := BaseCodif + 1;
          end;
          if (DucsVariables.BaseArr = 'P') then
          begin
            BaseCodif := Arrondi(BaseCodif, 0);
            MontForfait := Arrondi(MontForfait, 0);
          end;

         {Calcul montant de la cotisation}
          if (TypeCotis = 'T') then
         {Base * Taux}
            MontCodif := (BaseCodif * TauxEnCours) / 100.0;
          if (TypeCotis = 'Q') then
          begin
         {Quantité (Base * Qté)}
            MontCodif := (MontForfait * Effectif);
            BaseCodif := MontForfait;
          end;

         {Arrondi du Montant}
          if (DucsVariables.MontArr  <> '') and (DucsVariables.MontArr  <> 'P') then
          begin
            MontCodif := Int(MontCodif);
            if (DucsVariables.MontArr  = 'S') then MontCodif := MontCodif + 1;
          end;
          if (DucsVariables.MontArr  = 'P') then
          begin
            MontCodif := Arrondi(MontCodif, 0);
          end;

         { Cumul SOUS TOTAL }
          MontSt := MontST + MontCodif;
          IndCalcSt := true;
          TOB_FilleLignes.PutValue('PDD_EFFECTIF', Effectif);
          TOB_FilleLignes.PutValue('PDD_BASECOTISATION',
            Arrondi(BaseCodif, 5));

          if (TypeCotis <> 'Q') then
            TOB_FilleLignes.PutValue('PDD_TAUXCOTISATION', Arrondi(TauxEnCours, 5))
          else
            TOB_FilleLignes.PutValue('PDD_TAUXCOTISATION', Effectif);

          TOB_FilleLignes.PutValue('PDD_MTCOTISAT', Arrondi(MontCodif, 5));
          TOB_FilleLignes.PutValue('PDD_COMURBAINE', '    ');
          TOB_FilleLignes.PutValue('PDD_CONDITION', CSC);
          TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
            copy(CodifEnCours, DucsVariables.CodifPosDebut, DucsVariables.CodifLongEdit));
          if (PogNature = '300') then
            TOB_FilleLignes.PutValue('PDD_INSTITUTION', CodInstEnCours)
          else
            TOB_FilleLignes.PutValue('PDD_INSTITUTION', '');

          NoInst := CodInstEnCours;
        end { fin if MonCodif<>0}
        else
          Ferme(QDucsParam);
      end;

    {TRAITEMENT DES TYPES DE CODIFICATION PARTICULIERS}
    {SOUS TOTAL}
      if (PogBoolST = TRUE) then
      begin
        if (MontSt <> 0.00) or ((PogNature = '300') and
          (PogSTPos <> 0) and (PogSTLong <> 0)) then
        begin
          if (PogNature = '200') or
            (PogNature = '600') or
            (PogNature = '300') then
        {ASSEDIC ou MSA  ou IRC personnalisée}
          begin
            if (Copy(CodifLue, PogSTPos, PogStLong) <>
              Copy(CodifEnCours, PogSTPos, PogStLong)) then
            begin
              if (PogNature = '200') then
            {ASSEDIC}
                CodifST := Copy(CodifEnCours, 1, 5) + 'ZZ';
              if (PogNature = '600') then
            {MSA}
                CodifST := Copy(CodifEnCours, 1, 4) + 'ZZZ';
              if (PogNature = '300') then
            {IRC personnalisée}
                if (PogSTPos = 1) then
                begin
                  CodifST := Copy(CodifEnCours, 1, PogSTLong) +
                    Copy(ZEDE, 1, (7 - PogSTLong));
                end
                else
                begin
                  CodifST := Copy(CodifEnCours, 1, PogSTPos - 1) +
                    Copy(CodifEnCours, PogSTPos, PogSTLong) +
                    Copy(ZEDE, 1, (7 - (PogStPos + PogSTLong - 1)));
                end;
              if (AlsaceMoselle) then
                QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, ' +
                                      'PDP_LIBELLESUITE,' +
                                      'PDP_TYPECOTISATION' +
                                      ' FROM DUCSPARAM ' +
                                      ' WHERE ##PDP_PREDEFINI## ' +
                                      'PDP_CODIFALSACE = "' + CodifST + '"' +
                                      ' ORDER BY PDP_PREDEFINI', True)
              else
                QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, ' +
                                      'PDP_LIBELLESUITE,' +
                                      'PDP_TYPECOTISATION' +
                                      ' FROM DUCSPARAM ' +
                                      ' WHERE ##PDP_PREDEFINI## ' +
                                      'PDP_CODIFICATION = "' + CodifST + '"' +
                                      ' ORDER BY PDP_PREDEFINI', True);
              Predef := '';
{$IFDEF EAGLCLIENT}
              SortDos := 0;
              KK := 0;
{$ENDIF}
              while not QDucsParam.eof do
              begin
                Predef := QDucsparam.FindField('PDP_PREDEFINI').AsString;
                if (Predef = 'DOS') then
                begin
{$IFDEF EAGLCLIENT}
                  SortDos := 1;
{$ENDIF}
                  break;
                end;
{$IFDEF EAGLCLIENT}
                KK := KK + 1;
{$ENDIF}
                QDucsParam.Next;
              end;
              if (Predef <> '') then
              begin
                ValChST[1] := CodifST;
                NomChampST[1] := 'PDD_CODIFICATION';
                TOB_FilleLignes := TOB_FilleTete.FindFirst(NomChampST, ValChST, TRUE);
                if (Personnalisee = TRUE) and
                  (TOB_FilleLignes <> nil) then
                begin
                  TOB_FilleLignes.PutValue('PDD_MTCOTISAT',
                    TOB_FilleLignes.GetValue('PDD_MTCOTISAT') +
                    Arrondi(MontST, 5));
                end
                else
                begin
                  AlimDetail(DebPer,FinPer,TOB_FilleLignes, TOB_FilleTete, ligne,DucsParametres, DucsVariables,DucsOrganisme.PogNumInterne);
                  TOB_FilleLignes.PutValue('PDD_CODIFICATION', CodifST);
{$IFDEF EAGLCLIENT}
                  if SortDos = 1 then MaTob := QDucsParam.Detail[QDucsParam.Detail.count - 1]
                  else MaTob := QDucsParam.Detail[KK - 1];
                  TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
                    MaTob.GetValue('PDP_TYPECOTISATION'));
                  TOB_FilleLignes.PutValue('PDD_LIBELLE',
                    MaTob.GetValue('PDP_LIBELLE'));
                  TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
                    MaTob.GetValue('PDP_LIBELLESUITE'));
{$ELSE}
                  TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
                    QDucsParam.FindField('PDP_TYPECOTISATION').AsString);
                  TOB_FilleLignes.PutValue('PDD_LIBELLE',
                    QDucsParam.FindField('PDP_LIBELLE').AsString);
                  TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
                    QDucsParam.FindField('PDP_LIBELLESUITE').AsString);
{$ENDIF}
// d PT4
{$IFDEF EAGLCLIENT}
                  if (Assigned(MaTob)) then
                    FreeandNil (MaTob);
{$ENDIF}
// f PT4
                  Ferme(QDucsParam);
                  TOB_FilleLignes.PutValue('PDD_EFFECTIF', 0);
                  TOB_FilleLignes.PutValue('PDD_BASECOTISATION', 0.00);
                  TOB_FilleLignes.PutValue('PDD_TAUXCOTISATION', 0.00);
                  TOB_FilleLignes.PutValue('PDD_MTCOTISAT', Arrondi(MontST, 5));
                  TOB_FilleLignes.PutValue('PDD_COMURBAINE', '    ');
                  TOB_FilleLignes.PutValue('PDD_CONDITION', 'SCS');
                  if (PogNature = '200') then
              { ASSEDIC }
                    TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
                      copy(CodifST, 5, 1) + ' ');
                  if (PogNature = '600') then
              { MSA }
                    TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
                      copy(CodifST, 3, 2));

                  if (PogNature = '300') then
                  begin
              {IRC personnalisée}
                    if (PogSTPos < 3) then
                    begin
                      PosDeb := 3;
                      LongEdit := PogSTLong - 2;
                    end
                    else
                    begin
                      PosDeb := PogStPos;
                      LongEdit := PogStLong;
                    end;

                    TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
                      copy(CodifST, PosDeb, LongEdit));
                    TOB_FilleLignes.PutValue('PDD_INSTITUTION',
                      CodInstEnCours);
                  end;
                end;
              {RAZ Cumul Sous Total}
                MontST := 0.00;
                IndCalcST := false;
              end; {if(Predef <> '')}
            end; { if (Copy(CodifLue,PogSTPos,PogStLong)...}
          end; {if (PogNature = '200') OR (PogNature = '600')}

          if (PogNature = '300') and (((PogSTPos = 0) and (PogSTLong = 0))) then
        {IRC  standardisée}
          begin
            if ((((DucsVariables.RupSiret = True) and (WsiretLu <> DucsVariables.siretEnCours)) or
              ((DucsVariables.RupApe = True) and (WApeLu <> DucsVariables.ApeEnCours)) or
              ((DucsVariables.RupNumInt = True) and (WNoIntLu <> DucsVariables.NumIntEnCours)) or
              ((DucsVariables.RupGpInt = True) and (WGpLu <> DucsVariables.GpIntEnCours))) or
              ((Orglu <> OrgEnCours) or
              (CodInstLu <> CodInstEnCours)) or
              ((CodifLue <> CodifEnCours) and
              (Copy(CodifLue, 1, 1) = '8') and
              (Copy(CodifEnCours, 1, 1) <> '8'))) then
            begin
              NomChamp[1] := 'PDD_CODIFICATION';
              ValeurChamp[1] := ColleZeroDevant(NoInstEnCours, 2) + 'ZZZZZ';
              NomChamp[2] := '';
              ValeurChamp[2] := '';

              TobSTExiste := TOB_FilleTete.FindFirst(NomChamp, ValeurChamp, TRUE);
              while (TobStExiste <> nil) do
              begin
                NoInstEnCours := NoInstEnCours + 1;
                ValeurChamp[1] := ColleZeroDevant(NoInstEnCours, 2) + 'ZZZZZ';
                TobSTExiste := TOB_FilleTete.FindFirst(NomChamp, ValeurChamp, TRUE);
              end;
              CodifST := ColleZeroDevant(NoInstEnCours, 2) + 'ZZZZZ';

              AlimDetail(DebPer,FinPer,TOB_FilleLignes, TOB_FilleTete, ligne, DucsParametres, DucsVariables,DucsOrganisme.PogNumInterne);
              TOB_FilleLignes.PutValue('PDD_CODIFICATION', CodifST);
              TOB_FilleLignes.PutValue('PDD_TYPECOTISATION', 'S');

             {recup Code Institution et Libellé}
              QInst := OpenSql('SELECT PIP_ABREGE ' +
                ' FROM INSTITUTIONPAYE ' +
                ' WHERE PIP_INSTITUTION = "' +
                NoInst + '"', True);

              if not QInst.eof then
              begin
                LibInst := 'ST ' + NoInst + ' ' +
                  Copy(QInst.FindField('PIP_ABREGE').AsString, 1, 7);
                TOB_FilleLignes.PutValue('PDD_LIBELLE', LibInst);
              end;
              Ferme(QInst);

              TOB_FilleLignes.PutValue('PDD_EFFECTIF', 0);
              TOB_FilleLignes.PutValue('PDD_BASECOTISATION', 0.00);
              TOB_FilleLignes.PutValue('PDD_TAUXCOTISATION', 0.00);
              TOB_FilleLignes.PutValue('PDD_MTCOTISAT', Arrondi(MontST, 5));
              TOB_FilleLignes.PutValue('PDD_COMURBAINE', '    ');
              TOB_FilleLignes.PutValue('PDD_CONDITION', 'SCS');
             {Le code n'est pas édité}
// ci-après les lignes seront à activer si le terme "AGFF" doit apparaître sur
// la ligne de Sous-total
//             if (Copy(CodifEnCours,1,1) = '8') then
//               TOB_FilleLignes.PutValue ('PDD_CODIFEDITEE','8')
//             else
              TOB_FilleLignes.PutValue('PDD_CODIFEDITEE', ' ');
              if (PogNature = '300') then
                TOB_FilleLignes.PutValue('PDD_INSTITUTION', NoInst)
              else
                TOB_FilleLignes.PutValue('PDD_INSTITUTION', '');
             {RAZ Cumul Sous Total}
              MontST := 0.00;
              IndCalcSt := false;
            end; { if (Orglu <> OrgEnCours)}
          end; { if (PogNature = '300')}

          if (PogNature = '700') then
        {BTP}
          begin
            if ((PogPeriod1 = PeriodEtat) and (PogTypDecl1 = TRUE)) or
              ((PogPeriod2 = PeriodEtat) and (PogTypDecl2 = TRUE)) then
          {Déclaration avec calcul}
            begin
              if (Copy(CodifLue, PogSTPos, PogStLong) <>
                Copy(CodifEnCours, PogSTPos, PogStLong)) then
              begin
                CodifST := Copy(CodifEnCours, 1, 4) + 'ZZZ';
                if (AlsaceMoselle) then
                  QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, ' +
                                        'PDP_LIBELLESUITE,' +
                                        'PDP_TYPECOTISATION' +
                                        ' FROM DUCSPARAM ' +
                                        ' WHERE ##PDP_PREDEFINI##' +
                                        ' PDP_CODIFALSACE = "' + CodifST + '"' +
                                        ' ORDER BY PDP_PREDEFINI', True)
                else
                  QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, ' +
                                        'PDP_LIBELLESUITE,' +
                                        'PDP_TYPECOTISATION' +
                                        ' FROM DUCSPARAM ' +
                                        ' WHERE ##PDP_PREDEFINI##' +
                                        ' PDP_CODIFICATION = "' + CodifST + '"' +
                                        ' ORDER BY PDP_PREDEFINI', True);
                Predef := '';
{$IFDEF EAGLCLIENT}
                SortDos := 0;
                KK := 0;
{$ENDIF}
                while not QDucsParam.eof do
                begin
                  Predef := QDucsparam.FindField('PDP_PREDEFINI').AsString;
                  if (Predef = 'DOS') then
                  begin
{$IFDEF EAGLCLIENT}
                    SortDos := 1;
{$ENDIF}
                    break;
                  end;
{$IFDEF EAGLCLIENT}
                  KK := KK + 1;
{$ENDIF}
                  QDucsParam.Next;
                end;

                if (Predef <> '') then
                begin
                  AlimDetail(DebPer,FinPer,TOB_FilleLignes, TOB_FilleTete, ligne, DucsParametres, DucsVariables,DucsOrganisme.PogNumInterne);
                  TOB_FilleLignes.PutValue('PDD_CODIFICATION', CodifST);
{$IFDEF EAGLCLIENT}
                  if SortDos = 1 then MaTob := QDucsParam.Detail[QDucsParam.Detail.count - 1]
                  else MaTob := QDucsParam.Detail[KK - 1];
                  TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
                    MaTob.GetValue('PDP_TYPECOTISATION'));
                  TOB_FilleLignes.PutValue('PDD_LIBELLE',
                    MaTob.GetValue('PDP_LIBELLE'));
                  TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
                    MaTob.GetValue('PDP_LIBELLESUITE'));
{$ELSE}
                  TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
                    QDucsParam.FindField('PDP_TYPECOTISATION').AsString);
                  TOB_FilleLignes.PutValue('PDD_LIBELLE',
                    QDucsParam.FindField('PDP_LIBELLE').AsString);
                  TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
                    QDucsParam.FindField('PDP_LIBELLESUITE').AsString);
{$ENDIF}
// d PT4
{$IFDEF EAGLCLIENT}
                 if (Assigned(MaTob)) then
                    FreeandNil (MaTob);
{$ENDIF}
// f PT4
                  Ferme(QDucsParam);
                  TOB_FilleLignes.PutValue('PDD_EFFECTIF', 0);
                  TOB_FilleLignes.PutValue('PDD_BASECOTISATION', 0.00);
                  TOB_FilleLignes.PutValue('PDD_TAUXCOTISATION', 0.00);
                  TOB_FilleLignes.PutValue('PDD_MTCOTISAT', Arrondi(MontST, 5));
                  TOB_FilleLignes.PutValue('PDD_COMURBAINE', '    ');
                  TOB_FilleLignes.PutValue('PDD_CONDITION', 'SCS');
                  TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
                    copy(CodifST, 3, 2));
                {RAZ Cumul Sous Total}
                  MontST := 0.00;
                  IndCalcSt := False;
                end; { if (Predef <> '')}
              end; {if (Copy(CodifLue,PogS}
            end; {if ((PogPeriod1 = PeriodEtat)}
          end; {if (PogNature = '700'}
        end; {if (MontST <> 0)}
      end; {if (PogBoolST = TRUE)}

      if (CodifLue <> CodifEnCours) then ligne := 0;
      BaseCodif := 0.00;
      MontCodif := 0.00;
    end; { fin while (tDetail....}
    FreeAndNil(TEffM); // PT45 FQ 10874

    if (TDetail = nil) and (DucsVariables.Rupture = False) then
    {il n'y a plus de détail}
      fctRupture(DucsVariables,TDetail, 1);

  {deb traitement dernière ligne de sous total}

  {SOUS TOTAL}
    if (PogBoolST = TRUE) and (IndCalcSt = true) then
    begin
      if (MontSt <> 0.00) or ((PogNature = '300') and
        (PogSTPos <> 0) and (PogSTLong <> 0)) then
      begin
        if (PogNature = '200') or
          (PogNature = '600') or
          (PogNature = '300') then
      {ASSEDIC ou MSA  ou IRC personnalisée}
        begin
          if (PogNature = '200') then
        {ASSEDIC}
            CodifST := Copy(CodifEnCours, 1, 5) + 'ZZ';

          if (PogNature = '600') then
        {MSA}
            CodifST := Copy(CodifEnCours, 1, 4) + 'ZZZ';

          if (PogNature = '300') then
        {IRC personnalisée}
            if (PogSTPos = 1) then
            begin
              CodifST := Copy(CodifEnCours, 1, PogSTLong) +
                Copy(ZEDE, 1, (7 - PogSTLong));
            end
            else
            begin
              CodifST := Copy(CodifEnCours, 1, PogSTPos - 1) +
                Copy(CodifEnCours, PogSTPos, PogSTLong) +
                Copy(ZEDE, 1, (7 - (PogStPos + PogSTLong - 1)));
            end;

          if (AlsaceMoselle) then
            QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, PDP_LIBELLESUITE,' +
                                  'PDP_TYPECOTISATION' +
                                  ' FROM DUCSPARAM ' +
                                  ' WHERE ##PDP_PREDEFINI## PDP_CODIFALSACE = "' +
                                   CodifST + '"' +
                                  ' ORDER BY PDP_PREDEFINI', True)
          else
            QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, PDP_LIBELLESUITE,' +
                                  'PDP_TYPECOTISATION' +
                                  ' FROM DUCSPARAM ' +
                                  ' WHERE ##PDP_PREDEFINI## PDP_CODIFICATION = "' +
                                   CodifST + '"' +
                                  ' ORDER BY PDP_PREDEFINI', True);
          Predef := '';
{$IFDEF EAGLCLIENT}
          SortDos := 0;
          KK := 0;
{$ENDIF}
          while not QDucsParam.eof do
          begin
            Predef := QDucsparam.FindField('PDP_PREDEFINI').AsString;
            if (Predef = 'DOS') then
            begin
{$IFDEF EAGLCLIENT}
              SortDos := 1;
{$ENDIF}
              break;
            end;
{$IFDEF EAGLCLIENT}
            KK := KK + 1;
{$ENDIF}
            QDucsParam.Next;
          end;
          if (Predef <> '') then
          begin
            ValChST[1] := CodifST;
            NomChampST[1] := 'PDD_CODIFICATION';
            TOB_FilleLignes := TOB_FilleTete.FindFirst(NomChampST, ValChST, TRUE);
            if (Personnalisee = TRUE) and
              (TOB_FilleLignes <> nil) then
            begin
              TOB_FilleLignes.PutValue('PDD_MTCOTISAT',
                TOB_FilleLignes.GetValue('PDD_MTCOTISAT') +
                Arrondi(MontST, 5));
            end
            else
            begin
              AlimDetail(DebPer,FinPer,TOB_FilleLignes, TOB_FilleTete, ligne, DucsParametres, DucsVariables,DucsOrganisme.PogNumInterne);
              TOB_FilleLignes.PutValue('PDD_CODIFICATION', CodifST);
{$IFDEF EAGLCLIENT}
          If SortDos = 1 then
            MaTob := QDucsParam.Detail [QDucsParam.Detail.count-1]
          else
            MaTob := QDucsParam.Detail [KK-1] ;
              TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
                MaTob.GetValue('PDP_TYPECOTISATION'));
              TOB_FilleLignes.PutValue('PDD_LIBELLE',
                MaTob.GetValue('PDP_LIBELLE'));
              TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
                MaTob.GetValue('PDP_LIBELLESUITE'));
{$ELSE}
              TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
                QDucsParam.FindField('PDP_TYPECOTISATION').AsString);
              TOB_FilleLignes.PutValue('PDD_LIBELLE',
                QDucsParam.FindField('PDP_LIBELLE').AsString);
              TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
                QDucsParam.FindField('PDP_LIBELLESUITE').AsString);
{$ENDIF}
// d PT4
{$IFDEF EAGLCLIENT}
              if (Assigned(MaTob)) then
                FreeandNil (MaTob);
{$ENDIF}
// f PT4
              Ferme(QDucsParam);
              TOB_FilleLignes.PutValue('PDD_EFFECTIF', 0);
              TOB_FilleLignes.PutValue('PDD_BASECOTISATION', 0.00);
              TOB_FilleLignes.PutValue('PDD_TAUXCOTISATION', 0.00);
              TOB_FilleLignes.PutValue('PDD_MTCOTISAT', Arrondi(MontST, 5));
              TOB_FilleLignes.PutValue('PDD_COMURBAINE', '    ');
              TOB_FilleLignes.PutValue('PDD_CONDITION', 'SCS');

              if (PogNature = '200') then
          {ASSEDIC}
                TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
                  copy(CodifST, 5, 1) + ' ');
              if (PogNature = '600') then
          {MSA}
                TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
                  copy(CodifST, 3, 2));

              if (PogNature = '300') then
              begin
          {IRC personnalisée}
                if (PogSTPos < 3) then
                begin
                  PosDeb := 3;
                  LongEdit := PogSTLong - 2;
                end
                else
                begin
                  PosDeb := PogStPos;
                  LongEdit := PogStLong;
                end;

                TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
                  copy(CodifST, PosDeb, LongEdit));
                TOB_FilleLignes.PutValue('PDD_INSTITUTION', CodInstEnCours);
              end;
            end;
          {RAZ Cumul Sous Total}
            MontST := 0.00;
            IndCalcSt := false;
          end; { if (Predef <> '')}
        end; { if (PogNature = '200') OR (PogNature = '600')}


        if (PogNature = '300') and (((PogSTPos = 0) and (PogSTLong = 0))) then
      {IRC  standardisée}
        begin
          NomChamp[1] := 'PDD_CODIFICATION';
          ValeurChamp[1] := ColleZeroDevant(NoInstEnCours, 2) + 'ZZZZZ';
          NomChamp[2] := '';
          ValeurChamp[2] := '';

          TobSTExiste := TOB_FilleTete.FindFirst(NomChamp, ValeurChamp, TRUE);
          while (TobStExiste <> nil) do
          begin
            NoInstEnCours := NoInstEnCours + 1;
            ValeurChamp[1] := ColleZeroDevant(NoInstEnCours, 2) + 'ZZZZZ';
            TobSTExiste := TOB_FilleTete.FindFirst(NomChamp, ValeurChamp, TRUE);
          end;
          CodifST := ColleZeroDevant(NoInstEnCours, 2) + 'ZZZZZ';
          AlimDetail(DebPer,FinPer,TOB_FilleLignes, TOB_FilleTete, ligne, DucsParametres, DucsVariables,DucsOrganisme.PogNumInterne);
          TOB_FilleLignes.PutValue('PDD_CODIFICATION', CodifST);
          TOB_FilleLignes.PutValue('PDD_TYPECOTISATION', 'S');
        {recup Code Institution et Libellé}
          QInst := OpenSql('SELECT PIP_ABREGE' +
            ' FROM INSTITUTIONPAYE ' +
            ' WHERE PIP_INSTITUTION = "' + NoInst + '"', True);
          if not QInst.eof then
          begin
            LibInst := 'ST ' + NoInst + ' ' +
              Copy(QInst.FindField('PIP_ABREGE').AsString, 1, 7);
            TOB_FilleLignes.PutValue('PDD_LIBELLE', LibInst);
          end;
          Ferme(QInst);

          TOB_FilleLignes.PutValue('PDD_EFFECTIF', 0);
          TOB_FilleLignes.PutValue('PDD_BASECOTISATION', 0.00);
          TOB_FilleLignes.PutValue('PDD_TAUXCOTISATION', 0.00);
          TOB_FilleLignes.PutValue('PDD_MTCOTISAT', Arrondi(MontST, 5));
          TOB_FilleLignes.PutValue('PDD_COMURBAINE', '    ');
          TOB_FilleLignes.PutValue('PDD_CONDITION', 'SCS');
        {Le code n'est pas édité}
// ci-après les lignes seront à activer si le terme "AGFF" doit apparaître sur
// la ligne de Sous-total
//        if (Copy(CodifEnCours,1,1)='8') then
//          TOB_FilleLignes.PutValue ('PDD_CODIFEDITEE','8')
//        else
          TOB_FilleLignes.PutValue('PDD_CODIFEDITEE', ' ');
          if (PogNature = '300') then
            TOB_FilleLignes.PutValue('PDD_INSTITUTION', NoInst)
          else
            TOB_FilleLignes.PutValue('PDD_INSTITUTION', '');
        {RAZ Cumul Sous Total}
          MontST := 0.00;
          IndCalcSt := false;
        end; { if (PogNature = '300')}

        if (PogNature = '700') then
      {BTP}
        begin
          if ((PogPeriod1 = PeriodEtat) and (PogTypDecl1 = TRUE)) or
            ((PogPeriod2 = PeriodEtat) and (PogTypDecl2 = TRUE)) then
        {Déclaration avec calcul}
          begin
            CodifST := Copy(CodifEnCours, 1, 4) + 'ZZZ';
            if (AlsaceMoselle) then
              QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, PDP_LIBELLESUITE,' +
                                    'PDP_TYPECOTISATION' +
                                    ' FROM DUCSPARAM ' +
                                    ' WHERE ##PDP_PREDEFINI## PDP_CODIFALSACE= "' + CodifST + '"' +
                                    ' ORDER BY PDP_PREDEFINI', True)
            else
              QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, PDP_LIBELLESUITE,' +
                                    'PDP_TYPECOTISATION' +
                                    ' FROM DUCSPARAM ' +
                                    ' WHERE ##PDP_PREDEFINI## PDP_CODIFICATION = "' + CodifST + '"' +
                                    ' ORDER BY PDP_PREDEFINI', True);
            Predef := '';
{$IFDEF EAGLCLIENT}
            SortDos := 0;
            KK := 0;
{$ENDIF}
            while not QDucsParam.eof do
            begin
              Predef := QDucsparam.FindField('PDP_PREDEFINI').AsString;
              if (Predef = 'DOS') then
              begin
{$IFDEF EAGLCLIENT}
                SortDos := 1;
{$ENDIF}
                break;
              end;
{$IFDEF EAGLCLIENT}
             KK := KK + 1;
{$ENDIF}
              QDucsParam.Next;
            end;

            if (Predef <> '') then
            begin
              AlimDetail(DebPer,FinPer,TOB_FilleLignes, TOB_FilleTete, ligne, DucsParametres, DucsVariables,DucsOrganisme.PogNumInterne);
              TOB_FilleLignes.PutValue('PDD_CODIFICATION', CodifST);
{$IFDEF EAGLCLIENT}
              if SortDos = 1 then MaTob := QDucsParam.Detail[QDucsParam.Detail.count - 1]
              else MaTob := QDucsParam.Detail[KK - 1];
              TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
                MaTob.GetValue('PDP_TYPECOTISATION'));
              TOB_FilleLignes.PutValue('PDD_LIBELLE',
                MaTob.GetValue('PDP_LIBELLE'));
              TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
                MaTob.GetValue('PDP_LIBELLESUITE'));
{$ELSE}
              TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
                QDucsParam.FindField('PDP_TYPECOTISATION').AsString);
              TOB_FilleLignes.PutValue('PDD_LIBELLE',
                QDucsParam.FindField('PDP_LIBELLE').AsString);
              TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
                QDucsParam.FindField('PDP_LIBELLESUITE').AsString);
{$ENDIF}
// d PT4
{$IFDEF EAGLCLIENT}
              if (Assigned(MaTob)) then
                FreeandNil (MaTob);
{$ENDIF}
// f PT4
              Ferme(QDucsParam);
              TOB_FilleLignes.PutValue('PDD_EFFECTIF', 0);
              TOB_FilleLignes.PutValue('PDD_BASECOTISATION', 0.00);
              TOB_FilleLignes.PutValue('PDD_TAUXCOTISATION', 0.00);
              TOB_FilleLignes.PutValue('PDD_MTCOTISAT', Arrondi(MontST, 5));
              TOB_FilleLignes.PutValue('PDD_COMURBAINE', '    ');
              TOB_FilleLignes.PutValue('PDD_CONDITION', 'SCS');
              TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
                copy(CodifST, 3, 2));
            {RAZ Cumul Sous Total}
              MontST := 0.00;
              IndcalcSt := false;
            end; {if (Predef <> '')}
          end; {if ((PogPeriod1 = PeriodEtat)}
        end; {if (PogNature = '700'}
      end; {if (MontSt <> 0)}
    end; {if (PogBoolST = TRUE)}
    {fin traitement dernière ligne détail}

    if (DucsVariables.Rupture = TRUE) then
    begin
      TRup := TRupMere.FindNext([''], [''], TRUE);
      if (TRup <> nil) then
      begin
        IndRup := False;
        if (TRup.GetValue('ET_SIRET') <> Null) then
          DucsVariables.SiretEnCours := TRup.GetValue('ET_SIRET');

        if (TRup.GetValue('ET_APE') <> Null) then
          DucsVariables.ApeEnCours := TRup.GetValue('ET_APE');

        DucsVariables.EtabEnCours := TRup.GetValue('POG_ETABLISSEMENT');
        DucsVariables.OrganismeEnCours := TRup.GetValue('POG_ORGANISME');

        if (TRup.GetValue('POG_NUMINTERNE') <> Null) then
          DucsVariables.NumIntEnCours := copy(TRup.GetValue('POG_NUMINTERNE'), 1, 30);

        if (TRup.GetValue('POG_LGOPTIQUE') <> Null) then
          LGOptiqueEncours := TRup.GetValue('POG_LGOPTIQUE');

        if (DucsVariables.RupGpInt = TRUE) then
        begin
          if (TRup.GetValue('PGI_GROUPE') <> Null) then
            DucsVariables.GpIntEnCours := TRup.GetValue('PGI_GROUPE');
        end;
      end;
    end;
  { traitement des ruptures }
    DucsVariables.NumDucs := DucsVariables.NumDucs + 1;
    ligne := 0;
    NoAT := -1;

  end; {fin while (fctRupture = true)}

  if (Anomalie = '') then
  begin
  {MISE A JOUR DES TABLES DUCSENTETE ET DUCSDETAIL}
  {-----------------------------------------------}
    try
      begintrans;
      if TOB_Lignes <> nil then
        TOB_Lignes.SetAllModifie(TRUE);
      StrDelDetail := 'DELETE FROM DUCSDETAIL ';
      if (DucsVariables.Rupture = FALSE) then
        StrDelDetail := StrDelDetail + 'WHERE ' +
          '(PDD_ETABLISSEMENT ="' + DucsParametres.etab + '") AND ' +
          '(PDD_ORGANISME ="' + DucsParametres.organisme + '") AND ' +
          '(PDD_DATEDEBUT ="' + UsDateTime(StrToDate(DebPer)) + '") AND ' +
          '(PDD_DATEFIN = "' + UsDateTime(StrToDate(FinPer)) + '")'
      else
        StrDelDetail := StrDelDetail + DucsVariables.StWhereDet + ')';
      DucsVariables.StWhereDet := 'WHERE (';

      ExecuteSQL(StrDelDetail);

      StrDelTete := 'DELETE FROM DUCSENTETE ';
      if (DucsVariables.Rupture = FALSE) then
        StrDelTete := StrDelTete + 'WHERE ' +
          '(PDU_ETABLISSEMENT ="' + DucsParametres.etab + '") AND ' +
          '(PDU_ORGANISME ="' + DucsParametres.organisme + '") AND ' +
          '(PDU_DATEDEBUT ="' + UsDateTime(StrToDate(DebPer)) + '") AND ' +
          '(PDU_DATEFIN = "' + UsDateTime(StrToDate(FinPer)) + '")'
      else
        StrDelTete := StrDelTete + DucsVariables.StWhere + ')';
      DucsVariables.StWhere := 'WHERE (';
      ExecuteSQL(StrDelTete);

      TOB_Lignes.InsertDB(nil, FALSE);
      Committrans;
    except
      Rollback;
      if (not IndTraceE) then
      begin
        TraceE.add ('Anomalies de traitement');
        TraceE.add ('--------------------------------------------------------');
        IndTraceE := True;
      end;
      TraceE.Add('Etablissement '+ DucsParametres.Etab +'! Erreur maj table DUCSDETAIL. Traitement abandonné');
    end;
  end
  else
  begin
    if (Anomalie = 'Non Renseigné') then
    begin
      if (not IndTraceE) then
      begin
        TraceE.add ('Anomalies de traitement');
        TraceE.add ('--------------------------------------------------------');
        IndTraceE := True;
      end;
      TraceE.Add ('Le Siret de l''établissement '+DucsParametres.Etab+' n''est pas renseigné. Traitement abandonné');
    end;

    if (Anomalie = 'erreur d''affectation') then
    begin
      if (not IndTraceE) then
      begin
        TraceE.add ('Anomalies de traitement');
        TraceE.add ('--------------------------------------------------------');
        IndTraceE := True;
      end;
      TraceE.Add ('Etablissement '+ DucsParametres.Etab+' : Vérifier l''affectation des rubriques pour cet organisme. Traitement abandonné');
    end;
  end;

  if TOB_Lignes <> nil then
  begin
    TOB_Lignes.Free;
  end;
end; { fin MajDucs}

// Mise en forme des "Where"
procedure ChargeWheres(DebPer, FinPer : string; var DucsVariables : TDucsvariables; DucsParametres : TDucsParametres;TRupMere : TOB);
var
  TRup                          : TOB;
begin
  {  Gestion des ruptures }
  if (DucsVariables.rupture = TRUE) then
  begin
    TRup := TRupMere.FindFirst([''], [''], TRUE);
    while (TRup <> nil) do
    begin
      if (TRup.GetValue('ET_SIRET') <> Null) then
        DucsVariables.SiretEnCours := TRup.GetValue('ET_SIRET');
      if (TRup.GetValue('POG_ETABLISSEMENT') <> Null) then
        DucsVariables.EtabEnCours := TRup.GetValue('POG_ETABLISSEMENT');
      if (TRup.GetValue('ET_APE') <> Null) then
        DucsVariables.ApeEnCours := TRup.GetValue('ET_APE');
// deb PT15-1
      if (DucsParametres.Regroupt <> '') then
        DucsVariables.OrganismeEnCours := DucsParametres.organisme
      else
        DucsVariables.OrganismeEnCours := TRup.GetValue('POG_ORGANISME');

      { en-tête }
      DucsVariables.StWhere := DucsVariables.StWhere +
        '((PDU_ETABLISSEMENT ="' + DucsVariables.EtabEnCours + '") AND ' +
        '(PDU_ORGANISME ="' + DucsVariables.organismeEnCours + '") AND ' +
        '(PDU_DATEDEBUT ="' + UsDateTime(StrToDate(DebPer)) + '") AND ' +
        '(PDU_DATEFIN = "' + UsDateTime(StrToDate(FinPer)) + '")) ';

      { détail}
      DucsVariables.StWhereDet := DucsVariables.StWhereDet +
        '((PDD_ETABLISSEMENT ="' + DucsVariables.EtabEnCours + '") AND ' +
        '(PDD_ORGANISME ="' + DucsVariables.organismeEnCours + '") AND ' +
        '(PDD_DATEDEBUT ="' + UsDateTime(StrToDate(DebPer)) + '") AND ' +
        '(PDD_DATEFIN = "' + UsDateTime(StrToDate(FinPer)) + '")) ';

      TRup := TRupMere.FindNext([''], [''], TRUE);
      if (TRup <> nil) then
      {il y a d'autres Siret}
      begin
        DucsVariables.StWhere := DucsVariables.StWhere + 'OR ';
        DucsVariables.StWhereDet := DucsVariables.StWhereDet + 'OR ';
      end;
    end; { finwhile (TRup <> NIL) }
  end; {fin if (rupture = TRUE)}

  if (DucsVariables.Rupture <> TRUE) then
  { Cas sans rupture}
  begin
    {en-tête}
    DucsVariables.StWhere := DucsVariables.StWhere +
      '((PDU_ETABLISSEMENT ="' + DucsParametres.Etab + '") AND ' +
      '(PDU_ORGANISME ="' + DucsParametres.organisme + '") AND ' +
      '(PDU_DATEDEBUT ="' + UsDateTime(StrToDate(DebPer)) + '") AND ' +
      '(PDU_DATEFIN = "' + UsDateTime(StrToDate(FinPer)) + '")) ';

    { détail}
    DucsVariables.StWhereDet := DucsVariables.StWhereDet +
      '((PDD_ETABLISSEMENT ="' + DucsParametres.Etab + '") AND ' +
      '(PDD_ORGANISME ="' + DucsParametres.organisme + '") AND ' +
      '(PDD_DATEDEBUT ="' + UsDateTime(StrToDate(DebPer)) + '") AND ' +
      '(PDD_DATEFIN = "' + UsDateTime(StrToDate(FinPer)) + '")) ';
  end; { fin if (Rupture <> TRUE) }
end; { fin ChargeWheres}

//Calcul des effectifs
// Effectifs des salariés concernés par la déclaration: recherche sur HISTOBULLETIN
// Effectif des salariés présents en fin de période : recherche sur SALARIES
procedure CalculEffectifs(DebPer, FinPer : string; DucsVariables :TDucsvariables;DucsParametres : TDucsParametres;var DucsEffectifs : TDucsEffectifs;TInst : TOB; PogTypBordereau, PogNature : string);
var
  LesInst                                   : TOB;
  NbInst                                    : Integer;
  StQQ, StRupt, StJoinRupt, Organ, PogNumInt: string;
  QQ                                        : Tquery;
  StJoin                                    : string;
  TEffectifs , LeSal                        : TOB;
  I                                         : integer;
  DateNaiss                                 : TDateTime;
  WEff                                      : double;  
begin
  STQQ := '';
  StJoin := '';

  {CALCUL DU NBRE DE SALARIES CONCERNES PAR LA DECLARATION}
  {(ayant des cotisations pour la période)                }
  {-------------------------------------------------------}
  if (DucsParametres.DucsDoss = False) then
    { ducs établissement }
    StQQ := 'PHB_ETABLISSEMENT="' + DucsParametres.Etab + '" AND '
  else
    { ducs dossier}
    if (DucsVariables.RupNumInt = FALSE) and
       (DucsVariables.RupGpInt = FALSE) then
    begin
    { On ne tient compte que des établissements pour lesquels "Ducs Dossier" est
      coché sur l'organisme traité
      dans ce cas là, le Left Join n'est pas fait par la suite }
      StJoin := 'LEFT JOIN ORGANISMEPAIE ON PHB_ETABLISSEMENT = POG_ETABLISSEMENT ';
      StQQ := 'POG_DUCSDOSSIER = "X" AND ';
    end;
  StQQ := StQQ + 'PHB_NATURERUB="COT" AND ';

  if (DucsParametres.Regroupt <> '') then
  {Ducs regroupement d'organismes }
  begin
    for NbInst := 0 to TInst.Detail.Count - 1 do
    begin
      LesInst := TInst.Detail[NbInst];
      Organ := LesInst.GetValue('POG_ORGANISME');
      if (NbInst <> 0) then
      begin
        if (Pos('PHB_ORGANISME="' + Organ + '"', stQQ) = 0) then
          stQQ := stQQ + ' OR ';
      end
      else
        stQQ := stQQ + '(';

      if (Pos('PHB_ORGANISME="' + Organ + '"', stQQ) = 0) then
        stQQ := stQQ + 'PHB_ORGANISME="' + Organ + '"';
    end;
    stQQ := stQQ + ')';
  end
  else
  {Ducs sans regroupement d'organismes}
    StQQ := StQQ + 'PHB_ORGANISME="' + DucsParametres.Organisme + '" ';

  StQQ := StQQ + 'AND ' +
    'PHB_DATEFIN >="' + UsDateTime(StrToDate(DebPer)) + '" AND ' + //PT47
    'PHB_DATEFIN <="' + UsDateTime(StrToDate(FinPer)) + '" AND ' +
    '(PHB_BASECOT <> 0 OR PHB_MTSALARIAL <> 0 OR PHB_MTPATRONAL <> 0)';

  if (DucsVariables.Rupture = FALSE) then
  {PAS DE RUPTURE}
  begin
    QQ := OpenSQL('SELECT Count(DISTINCT PHB_SALARIE) AS NOMBRE FROM HISTOBULLETIN ' + StJoin + 'WHERE ' + StQQ, True);
  end
  else
  {EXISTENCE DE RUPTURE }
  begin
    StRupt := '';
    StJoinRupt := '';
    if (DucsVariables.RupSiret = TRUE) then
    begin
      StJoinRupt := StJoinRupt +
        ' LEFT JOIN ETABLISS ON PHB_ETABLISSEMENT=ET_ETABLISSEMENT ';
      StRupt := 'AND ET_SIRET="' + DucsVariables.SiretEnCours + '" ';
    end;
    if (DucsVariables.RupApe = TRUE) then
    begin
      StJoinRupt := StJoinRupt +
        ' LEFT JOIN ETABLISS ON PHB_ETABLISSEMENT=ET_ETABLISSEMENT ';
      StRupt := 'AND ET_APE="' + DucsVariables.ApeEnCours + '" ';
    end;
    if (DucsVariables.RupNumInt = TRUE) then
    begin
      StJoinRupt := StJoinRupt +
        ' LEFT JOIN ORGANISMEPAIE ' +
        'ON PHB_ETABLISSEMENT=POG_ETABLISSEMENT AND ' +
        'PHB_ORGANISME=POG_ORGANISME';
      if (DucsVariables.RupSiret = TRUE) or (DucsVariables.RupApe = TRUE) then
        StRupt := StRupt + ' AND '
      else
        StRupt := 'AND ';

      StRupt := StRupt + 'POG_NUMINTERNE="' + DucsVariables.NumIntEnCours + '"';
    end;

    if (DucsVariables.RupGpInt = TRUE) then
    begin
      StJoinRupt := StJoinRupt +
        ' LEFT JOIN SALARIES ' +
        'ON PHB_SALARIE=PSA_SALARIE ' +
        'LEFT JOIN GROUPEINTERNE ' +
        'ON PGI_CATEGORIECRC=PSA_DADSPROF ';
      if (DucsVariables.RupNumInt = FALSE) then
        StJoinRupt := StJoinRupt +
          ' LEFT JOIN ORGANISMEPAIE ' +
          'ON PHB_ORGANISME=POG_ORGANISME ';
      if (DucsVariables.RupSiret = TRUE) or
         (DucsVariables.RupApe = TRUE) or
         (DucsVariables.RupNumInt = TRUE) then
        StRupt := StRupt + ' AND '
      else
        StRupt := ' AND ';

      StRupt := StRupt + 'PGI_GROUPE="' + DucsVariables.GpIntEnCours + '" ';

      if (DucsParametres.DucsDoss = TRUE) then
        StRupt := StRupt + 'AND POG_DUCSDOSSIER="X"';
    end;

    QQ := OpenSQL('SELECT Count(DISTINCT PHB_SALARIE) AS NOMBRE FROM HISTOBULLETIN ' +
      StJoin + ' ' + StJoinRupt + ' WHERE ' + StQQ + StRupt, True);
  end; { fin EXISTENCE DE RUPTURE}
  if not QQ.EOF then // PortageCWAS
    DucsEffectifs.EffTot := QQ.FindField('NOMBRE').AsInteger;
  Ferme(QQ);

 {FIN -CALCUL DU NOMBRE DE SALARIES CONCERNES PAR LA DECLARATION}
  {--------------------------------------------------------------}

  {CALCUL DES EFFECTIFS PRESENTS EN FIN DE PERIODE}
  {-----------------------------------------------}
  StQQ := '';
  StJoin := '';

  if (DucsParametres.DucsDoss = False) then
    { ducs établissement }
    StQQ := 'PSA_ETABLISSEMENT="' + DucsParametres.Etab + '" AND '
  else
    { ducs dossier}
    if (DucsVariables.RupNumInt = FALSE) and (DucsVariables.RupGpInt = FALSE) then
    begin
    { On ne tient compte que des établissements pour lesquels "Ducs Dossier" est
      coché sur l'organisme traité
      dans ce cas là, le Left Join n'est pas fait par la suite }
      StJoin := 'LEFT JOIN ORGANISMEPAIE ON PSA_ETABLISSEMENT = POG_ETABLISSEMENT ';
      StQQ := 'POG_DUCSDOSSIER="X" AND POG_ORGANISME ="' + DucsParametres.Organisme + '" AND';
    end;

  StQQ := StQQ + '(PSA_DATESORTIE="' + UsDateTime(IDate1900) + '"' +
    ' OR  PSA_DATESORTIE IS NULL ' +
    ' OR PSA_DATESORTIE>="' + UsDateTime(StrToDate(FinPer)) + '")' +
    ' AND PSA_DATEENTREE<="' + UsDateTime(StrToDate(FinPer)) + '"';

  if (DucsVariables.Rupture = FALSE) then
  {PAS DE RUPTURE}
  begin
    StRupt := '' ;
    StJoinRupt := StJoin;
  end {FIN PAS DE RUPTURE}
  else
  {EXISTENCE DE RUPTURE }
  begin
    StRupt := '';
    StJoinRupt := '';
    if (DucsVariables.RupSiret = TRUE) then
    begin
      StJoinRupt := StJoinRupt +
        ' LEFT JOIN ETABLISS ON PSA_ETABLISSEMENT=ET_ETABLISSEMENT ';
      StRupt := 'AND ET_SIRET="' + DucsVariables.SiretEnCours + '" ';
    end;
    if (DucsVariables.RupApe = TRUE) then
    begin
      StJoinRupt := StJoinRupt +
        ' LEFT JOIN ETABLISS ON PSA_ETABLISSEMENT=ET_ETABLISSEMENT ';
      StRupt := 'AND ET_APE="' + DucsVariables.ApeEnCours + '" ';
    end;
    if (DucsVariables.RupNumInt = TRUE) then
    begin
      StJoinRupt := StJoinRupt +
        ' LEFT JOIN ORGANISMEPAIE ' +
        'ON PSA_ETABLISSEMENT=POG_ETABLISSEMENT AND ';
      if (DucsVariables.RupSiret = TRUE) or (DucsVariables.RupApe = TRUE) then
        StRupt := StRupt + ' AND '
      else
        StRupt := 'AND ';
      if (DucsParametres.Regroupt <> '') then
      {Ducs regroupement d'organismes}
      begin
        for NbInst := 0 to Tinst.Detail.Count - 1 do
        begin
          LesInst := Tinst.Detail[NbInst];
          if (LesInst.GetValue('POG_NUMINTERNE') <> Null) then
            PogNumInt := LesInst.GetValue('POG_NUMINTERNE');
          if (NbInst <> 0) then
            StRupt := StRupt + ' OR '
          else
            StRupt := StRupt + '(';
          StRupt := StRupt + 'POG_NUMINTERNE="' + PogNumInt + '"';
        end;
        StRupt := StRupt + ')';
      end
      else
        StRupt := StRupt + 'POG_NUMINTERNE="' + DucsVariables.NumIntEnCours + '"';
    end;

    if (DucsVariables.RupGpInt = TRUE) then
    begin
      StJoinRupt := StJoinRupt +
        ' LEFT JOIN GROUPEINTERNE ' +
        'ON PSA_DADSPROF=PGI_CATEGORIECRC ';
      if (DucsVariables.RupNumInt = False) then
        StJoinRupt := StJoinRupt +
          'LEFT JOIN ORGANISMEPAIE ' +
          'ON PSA_ETABLISSEMENT = POG_ETABLISSEMENT ';
      if (DucsVariables.RupSiret = TRUE) or
         (DucsVariables.RupApe = TRUE) or
         (DucsVariables.RupNumInt = TRUE) then
        StRupt := StRupt + ' AND '
      else
        StRupt := 'AND ';

      StRupt := StRupt + 'PGI_GROUPE="' + DucsVariables.GpIntEnCours + '"  ' +
        'AND POG_ORGANISME="' + DucsVariables.OrganismeEnCours + '" ';

      if (DucsParametres.DucsDoss = TRUE) then
        StRupt := StRupt + 'AND POG_DUCSDOSSIER="X"';
    end;
    StJoinRupt := StJoin + StJoinRupt;
  end;{FIN EXISTENCE DE RUPTURE}

   StQQ := StQQ +' and ((pci_debutcontrat <= "' + UsDateTime(StrToDate(FinPer)) + '" and '+
                 'pci_fincontrat >= "' + UsDateTime(StrToDate(DebPer)) + '" or '+
                 'pci_fincontrat = "' + UsDateTime(IDate1900) + '") or '+
                 '(pci_typecontrat is  NULL) or (psa_salarie not in '+
                 '(select pci_Salarie from contrattravail where '+
                 '(pci_debutcontrat <= "' + UsDateTime(StrToDate(FinPer)) + '" and '+
                 'pci_fincontrat >= "' + UsDateTime(StrToDate(DebPer)) + '" or '+
                 'pci_fincontrat = "' + UsDateTime(IDate1900) + '"))))';

  if PogNature <> '600' then
    StQQ := StQQ + ' and psa_salarie not in (select pse_salarie from deportsal where '+
                 'pse_msa="X")'
  else
    StQQ := StQQ + ' and psa_salarie in (select pse_salarie from deportsal where '+
                 'pse_msa="X")';

  QQ := OpenSQL('SELECT DISTINCT (PSA_SALARIE),PSA_SEXE,PSA_CATDADS,PSA_DATENAISSANCE,'+
                'PSA_PRISEFFECTIF, PSA_UNITEPRISEFF,'+
                'PSA_DADSPROF,PSA_PROFIL, PCI_TYPECONTRAT FROM SALARIES '+
                  StJoinRupt+
                  ' left join contrattravail on psa_salarie = pci_salarie WHERE '+
                  StQQ+StRupt,True) ;
  TEffectifs := TOB.Create('Les effectifs', NIL, -1);
  TEffectifs.LoadDetailDB('SAL','','',QQ,False);
    Ferme(QQ);

  if (TEffectifs.Detail.Count <> 0) then
  begin
    DateNaiss := PlusMois(StrToDate(FinPer),-780);

    for I := 0 to TEffectifs.Detail.Count-1 do
    begin
      LeSal := TEffectifs.Detail[I];
      if (LeSal.GetValue('PSA_DATENAISSANCE') <=  DateNaiss) then
        LeSal.AddChampSupValeur('+65ANS','X')
      else
         LeSal.AddChampSupValeur('+65ANS','-');
    QQ := OpenSQL('SELECT PPS_PROFIL FROM PROFILSPECIAUX WHERE PPS_CODE ="'+
                   LeSal.GetValue('PSA_SALARIE')+'"',True);
    if QQ.Eof then
      LeSal.AddChampSupValeur('PROFILSPECIAL','')
    else
      LeSal.AddChampSupValeur('PROFILSPECIAL',QQ.Fields[0].AsString);
    Ferme(QQ);
    end;

    if (PogTypBordereau = '915') or (PogTypBordereau = '916') then
    // Pour le TR URSSAF un salarié à temps partiel est pris en compte au
    // prorata de son temps de travail (PSA_UNITEPRISEFF)
    begin
      {nbre d'hommes}
      WEff := TEffectifs.Somme('PSA_UNITEPRISEFF',['PSA_SEXE','PSA_PRISEFFECTIF'],['M','X'],TRUE, TRUE);
      DucsEffectifs.EffHom := Int(Weff);
      if (DucsEffectifs.EffHom <> Weff) then DucsEffectifs.EffHom := DucsEffectifs.EffHom+1;
      {nbre de femmes}
      WEff := TEffectifs.Somme('PSA_UNITEPRISEFF',['PSA_SEXE','PSA_PRISEFFECTIF'],['F','X'],TRUE, FALSE);
      DucsEffectifs.EffFem := Int(Weff);
      if (DucsEffectifs.EffFem <> Weff) then DucsEffectifs.EffFem := DucsEffectifs.EffFem+1;
      {Nbre d'apprentis}
      Weff:= TEffectifs.Somme('PSA_UNITEPRISEFF',['PSA_CATDADS','PSA_PRISEFFECTIF'],['003','X'],TRUE, TRUE);
      DucsEffectifs.EffApp := Int(Weff);
      if (DucsEffectifs.EffApp <> Weff) then DucsEffectifs.EffApp := DucsEffectifs.EffApp+1;
    end
    else
    begin
      {nbre d'hommes}
      DucsEffectifs.EffHom := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_PRISEFFECTIF'],['M','X'],TRUE, TRUE);
      {nbre de femmes}
      DucsEffectifs.EffFem := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_PRISEFFECTIF'],['F','X'],TRUE, TRUE);
      {Nbre d'apprentis}
      DucsEffectifs.EffApp:= TEffectifs.Somme('PSA_CATDADS',['PSA_CATDADS','PSA_PRISEFFECTIF'],['003','X'],TRUE, TRUE);
    end;
    {Nbre d'apprentis hommes}
    DucsEffectifs.EffAppH:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_CATDADS','PSA_PRISEFFECTIF'],['M','003','X'],TRUE, TRUE);
    {Nbre d'apprentis femmes}
    DucsEffectifs.EffAppF:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_CATDADS','PSA_PRISEFFECTIF'],['F','003','X'],TRUE, TRUE);
    {Nbre de cadres  (cadre ou dirigeant)}
    DucsEffectifs.EffCad := TEffectifs.Somme('PSA_CATDADS',['PSA_CATDADS','PSA_PRISEFFECTIF'],['001','X'],TRUE, TRUE)+
              TEffectifs.Somme('PSA_CATDADS',['PSA_CATDADS','PSA_PRISEFFECTIF'],['002','X'],TRUE, TRUE);
    {Nbre de cadres hommes (cadre ou dirigeant)}
    DucsEffectifs.EffCadH := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF','PSA_PRISEFFECTIF'],['M','22','X'],TRUE, TRUE)+
               TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF','PSA_PRISEFFECTIF'],['M','23','X'],TRUE, TRUE)+
               TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF','PSA_PRISEFFECTIF'],['M','24','X'],TRUE, TRUE)+
               TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF','PSA_PRISEFFECTIF'],['M','25','X'],TRUE, TRUE);
    {Nbre de cadres femmes (cadre ou dirigeant)}
    DucsEffectifs.EFFCadF := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF','PSA_PRISEFFECTIF'],['F','22','X'],TRUE, TRUE)+
               TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF','PSA_PRISEFFECTIF'],['F','23','X'],TRUE, TRUE)+
               TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF','PSA_PRISEFFECTIF'],['F','24','X'],TRUE, TRUE)+
               TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF','PSA_PRISEFFECTIF'],['F','25','X'],TRUE, TRUE);
    {Nbre d'ouvriers}
    DucsEffectifs.EffOuv := TEffectifs.Somme('PSA_CATDADS',['PSA_CATDADS','PSA_PRISEFFECTIF'],['004'],TRUE, TRUE);

    {nbre d'Etam }
    DucsEffectifs.EffEtam := TEffectifs.Somme('PSA_CATDADS',['PSA_CATDADS','PSA_PRISEFFECTIF'],['005','X'],TRUE, TRUE);

    {nbre salariés hommes de + de 65 ans}
    DucsEffectifs.Eff65H := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','+65ANS','PSA_PRISEFFECTIF'],['M','X','X'],TRUE, TRUE);

    {nbre salariés femmes de + de 65 ans}
    DucsEffectifs.Eff65F := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','+65ANS','PSA_PRISEFFECTIF'],['F','X','X'],TRUE, TRUE);

    {Nbre d'hommes en contrat de professionnalisation}
//PT62    EffCesH:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_PROFIL'],['M','023'],TRUE, TRUE);
    DucsEffectifs.EffProfH:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PROFILSPECIAL','PSA_PRISEFFECTIF'],['M','296','X'],TRUE, TRUE);

    {Nbre de femmes titulaires d'un CES}
//PT62    EffCesF:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_PROFIL'],['F','023'],TRUE, TRUE);
    DucsEffectifs.EffProfF:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PROFILSPECIAL','PSA_PRISEFFECTIF'],['F','296','X'],TRUE, TRUE);

    {Nbre d'hommes en CDI}
//  PT62  EffCDIH:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PCI_TYPECONTRAT'],['M','CDI'],TRUE, TRUE);
    DucsEffectifs.EffCDIH:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_PRISEFFECTIF'],
                               ['M','CDI','-','X'],
                               TRUE, TRUE);
    DucsEffectifs.EffCadH22:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_DADSPROF','PSA_PRISEFFECTIF'],
                               ['M','CDI','-','22','X'],
                               TRUE, TRUE);
    DucsEffectifs.EffCadH23:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_DADSPROF','PSA_PRISEFFECTIF'],
                               ['M','CDI','-','23','X'],
                               TRUE, TRUE);
    DucsEffectifs.EffCadH24:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_DADSPROF','PSA_PRISEFFECTIF'],
                               ['M','CDI','-','24','X'],
                               TRUE, TRUE);
    DucsEffectifs.EffCadH25:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_DADSPROF','PSA_PRISEFFECTIF'],
                               ['M','CDI','-','25','X'],
                               TRUE, TRUE);
    DucsEffectifs.EffCDIH := DucsEffectifs.EffCDIH -
                             DucsEffectifs.EffCadH22 -
                             DucsEffectifs.EffCadH23 -
                             DucsEffectifs.EffCadH24 -
                             DucsEffectifs.EffCadH25;

    {Nbre de femmes en CDI}
    DucsEffectifs.EffCDIF:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_PRISEFFECTIF'],
                               ['F','CDI','-','X'],
                               TRUE, TRUE);
    DucsEffectifs.EffCadF22:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_DADSPROF','PSA_PRISEFFECTIF'],
                               ['F','CDI','-','22','X'],
                               TRUE, TRUE);
    DucsEffectifs.EffCadF23:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_DADSPROF','PSA_PRISEFFECTIF'],
                               ['F','CDI','-','23','X'],
                               TRUE, TRUE);
    DucsEffectifs.EffCadF24:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_DADSPROF','PSA_PRISEFFECTIF'],
                               ['F','CDI','-','24','X'],
                               TRUE, TRUE);
    DucsEffectifs.EffCadF25:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_DADSPROF','PSA_PRISEFFECTIF'],
                               ['F','CDI','-','25','X'],
                               TRUE, TRUE);
    DucsEffectifs.EffCDIF := DucsEffectifs.EffCDIF -
                             DucsEffectifs.EffCadF22 -
                             DucsEffectifs.EffCadF23 -
                             DucsEffectifs.EffCadF24 -
                             DucsEffectifs.EffCadF25;

    {Nbre d'hommes en CDD}
    DucsEffectifs.EffCDDH:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PCI_TYPECONTRAT','PSA_PRISEFFECTIF'],['M','CCD','X'],TRUE, TRUE);

    {Nbre de femmes en CDD}
    DucsEffectifs.EffCDDF:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PCI_TYPECONTRAT','PSA_PRISEFFECTIF'],['F','CCD','X'],TRUE, TRUE);

    {Nbre d'hommes en CNE}
    DucsEffectifs.EffCNEH:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PCI_TYPECONTRAT','PSA_PRISEFFECTIF'],['M','CNE','X'],TRUE, TRUE);

    {Nbre de femmes en CNE}
    DucsEffectifs.EffCNEF:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PCI_TYPECONTRAT','PSA_PRISEFFECTIF'],['F','CNE','X'],TRUE, TRUE);
  end;
  FreeAndNIl(Teffectifs);
  { FIN - CALCUL DES EFFECTIFS PRESENTS EN FIN DE PERIODE}
end; { fin Calculer effectifs}

// Création des TOB filles (lignes détail) + alimentation de
// champs, commune quelque soit le type de ligne.
procedure AlimDetail(DebPer, FinPer : string; var TOB_FilleLignes, TOB_FilleTete: TOB; var ligne: Integer; DucsParametres:TDucsParametres; DucsVariables:TDucsVariables; PogNumInterne : String);
begin

  TOB_FilleLignes := TOB.create('DUCSDETAIL', TOB_FilleTete, -1);

  ligne := ligne + 1;

  TOB_FilleLignes.PutValue('PDD_ETABLISSEMENT', DucsVariables.EtabEnCours);
  if (DucsParametres.Regroupt <> '') then
    TOB_FilleLignes.PutValue('PDD_ORGANISME', DucsParametres.Organisme)
  else
    TOB_FilleLignes.PutValue('PDD_ORGANISME', DucsVariables.OrganismeEnCours);

  TOB_FilleLignes.PutValue('PDD_DATEDEBUT', StrToDate(DebPer));
  TOB_FilleLignes.PutValue('PDD_DATEFIN', StrToDate(FinPer));
  TOB_FilleLignes.PutValue('PDD_NUM', DucsVariables.NumDucs);

  if (DucsVariables.Rupture = TRUE) then
    TOB_FilleLignes.PutValue('PDD_SIRET', DucsVariables.SiretEnCours)
  else
    TOB_FilleLignes.PutValue('PDD_SIRET', DucsVariables.EtSiret);

  if (DucsVariables.Rupture = TRUE) then
    TOB_FilleLignes.PutValue('PDD_APE', DucsVariables.ApeEnCours)
  else
    TOB_FilleLignes.PutValue('PDD_Ape', DucsVariables.EtApe);

  if (DucsVariables.RupGpInt = TRUE) then
    TOB_FilleLignes.PutValue('PDD_GROUPE', DucsVariables.GpIntEnCours)
  else
    TOB_FilleLignes.PutValue('PDD_GROUPE', '');

  if (DucsVariables.Rupture = TRUE) then
    TOB_FilleLignes.PutValue('PDD_NUMERO', DucsVariables.NumIntEnCours)
  else
    TOB_FilleLignes.PutValue('PDD_NUMERO', PogNumInterne);

  TOB_FilleLignes.PutValue('PDD_NUMORDRE', ligne);

  if (AlsaceMoselle) then
    TOB_FilleLignes.PutValue('PDD_REGIMEALSACE', 'X')
  else
    TOB_FilleLignes.PutValue('PDD_REGIMEALSACE', '-')
end; { fin AlimDetail }

procedure LibereTobs(var TInst,TCot3,TRupMere : TOB);
begin
  if TInst <> nil then
  begin
    TInst.Free;
  end;

  FreeAndNil(TCot3);

  if (TRupMere <> nil) then
  begin
    TRupMere.Free;
    TRupMere := nil;
  end;
end; { fin LibereTobs}

{ Traitement des ruptures }
function fctRupture(var DucsVariables : TDucsVariables; TDetail: TOB; Ind: integer): Boolean;
var
  Siretlu, Apelu, NumIntlu, GpIntlu: string;

begin
  Siretlu := '';
  Apelu := '';
  NumIntlu := '';
  GpIntlu := '';

  Result := TRUE;
  DucsVariables.IndNeant2 := FALSE;

  if (Ind = 1) then
  {fin des lignes détail}
  begin
    DucsVariables.WSiret := 'findetail';
  end
  else
    if (TDetail = nil) and (DucsVariables.WSiret <> 'findetail') then
    {Pas de ligne détail - ducs néant}
    begin
      DucsVariables.IndNeant2 := TRUE;
      exit;
    end;

  if (DucsVariables.WSiret = 'findetail') then
  {fin du traitement}
  begin
    Result := FALSE;
    DucsVariables.IndNeant2 := FALSE;
    exit;
  end
  else
  begin
    if (DucsVariables.RupSiret = TRUE) then
      if (TDetail.GetValue('SIRET') <> Null) then
        Siretlu := TDetail.GetValue('SIRET');

    if (DucsVariables.RupApe = TRUE) then
      if (TDetail.GetValue('APE') <> Null) then
        Apelu := TDetail.GetValue('APE');

    if (DucsVariables.RupNumInt = TRUE) then
      if (TDetail.GetValue('NUMINTERNE') <> Null) then
        NumIntlu := TDetail.GetValue('NUMINTERNE');

    if (DucsVariables.RupGpInt = TRUE) then
      if (TDetail.GetValue('GROUPE') <> Null) then
        GpIntlu := TDetail.GetValue('GROUPE');

    {Y-a-t'il rupture de Siret ?}
    if (DucsVariables.RupSiret = TRUE) and
       (DucsVariables.WSiret <> Siretlu) then
    begin
      Result := FALSE;
      DucsVariables.WSiret := Siretlu;
    end;

    if (DucsVariables.RupSiret = TRUE) then
      if (DucsVariables.WSiret <> DucsVariables.SiretEnCours) then
      {il n'y a aucune ligne pour le Siret en cours : c'est une ducs néant}
        DucsVariables.IndNeant2 := TRUE
      else
        DucsVariables.IndNeant2 := FALSE;

    {Y-a-t'il rupture de code Ape}
    if (DucsVariables.RupApe = TRUE) and (DucsVariables.WApe <> Apelu) then
    begin
      Result := FALSE;
      DucsVariables.WApe := Apelu;
    end;

    if (DucsVariables.RupApe = TRUE) then
      if (DucsVariables.WApe <> DucsVariables.ApeEnCours) then
      {il n'y a aucune ligne pour le code Ape en cours : c'est une ducs néant}
        DucsVariables.IndNeant2 := TRUE
      else
        DucsVariables.IndNeant2 := FALSE;

    {Y-a-t'il rupture de n° interne}
    if (DucsVariables.RupNumInt = TRUE) and
       (DucsVariables.WNumInt <> NumIntlu) then
    begin
      Result := FALSE;
      DucsVariables.WNumInt := NumIntlu;
    end;

    if (DucsVariables.RupNumInt = TRUE) and (DucsVariables.IndNeant2 = FALSE) then
      if (DucsVariables.WNumInt <> DucsVariables.NumIntEnCours) then
      {il n'y a aucune ligne pour le n° interne en cours : c'est une ducs néant}
        DucsVariables.IndNeant2 := TRUE
      else
        DucsVariables.IndNeant2 := FALSE;

    if (DucsVariables.RupGpInt = TRUE) and (DucsVariables.WGpInt <> GpIntlu) then
    begin
      Result := FALSE;
      DucsVariables.WGpInt := GpIntlu;
    end;
    if (DucsVariables.RupGpInt = TRUE) and (DucsVariables.IndNeant2 = FALSE) then
      if (DucsVariables.WGpInt <> DucsVariables.GpIntEnCours) then
      {il n'y a aucune ligne pour le groupe interne en cours : c'est une ducs néant}
        DucsVariables.IndNeant2 := TRUE
      else
        DucsVariables.IndNeant2 := FALSE;
  end;

  if (DucsVariables.WSiret = '') then
    DucsVariables.WSiret := Siretlu;
  if (DucsVariables.WApe = '') then
  DucsVariables.WApe := Apelu;
  if (DucsVariables.WNumInt = '') then
    DucsVariables.WNumInt := NumIntlu;
  if (DucsVariables.WGpInt = '') then
    DucsVariables.WGpInt := GpIntlu;
end; { fin fctRupture }

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE MF
Créé le ...... : 02/04/2007
Modifié le ... :   /  /    
Description .. : Fonction appelée par CgiPaieS5 (process server)
Suite ........ : Lance DucsInit (Initialisation des Ducs)
Suite ........ : Renvoie une TOB
Mots clefs ... : PROCESSSERVER
*****************************************************************}
function Process_DucsInit(PeriodEtat,DebPer, FinPer, NatDeb, NatFin, EtabDeb,EtabFin : string): TOB;
{$IFDEF EAGLSERVER}
var TobResult: TOB;
{$ENDIF}
begin
{$IFDEF EAGLSERVER}
  TOBResult := TOB.Create('DucsInitPROCESS', nil, -1);
  DucsInit(PeriodEtat,DebPer, FinPer, NatDeb, NatFin, EtabDeb,EtabFin);
  TOBResult.AddChampSupValeur('OKTRAITE', TRUE);
  result := TOBResult;
{$ENDIF}
end;

// d PT1
Function  AffectTypeBordereau( NatureOrg, Periodicite : string; DucsDossier : boolean):string;
var TypePeriod : string;
begin
  TypePeriod := '';
  result := '';
    if (NatureOrg = '100') then
    // ACOSS
    begin
      if ((Periodicite = 'M') or (Periodicite = 'T')) and (not DucsDossier) then
        TypePeriod := '913';    // BRC d'un établissement
      if ((Periodicite = 'M') or (Periodicite = 'T')) and (DucsDossier) then
        TypePeriod := '914';    // BRC de plusieurs établissements
      if (Periodicite = 'A') and (not DucsDossier) then
        TypePeriod := '915';    // TR d'un établissement
      if (Periodicite = 'A') and (DucsDossier) then
        TypePeriod := '916'     // TR de plusieurs établissements
    end
    else
    if (NatureOrg = '200') then
    // UNEDIC
    begin
      if ((Periodicite = 'M') or (Periodicite = 'T')) and (not DucsDossier) then
        TypePeriod := '920';    // ADV d'un établissement
      if ((Periodicite = 'M') or (Periodicite = 'T')) and (DucsDossier) then
        TypePeriod := '921';    // ADV de plusieurs établissements
      if (Periodicite = 'A') and (not DucsDossier) then
        TypePeriod := '922';     // DRA d'un établissement
      if (Periodicite = 'A') and (DucsDossier) then
        TypePeriod := '923'     // DRA de plusieurs établissements
    end
    else
    if (NatureOrg = '300') then
    // IRC
    begin
      if (Periodicite = 'M') then
        TypePeriod := '931';    // Bordereau mensuel
      if (Periodicite = 'T') then
        TypePeriod := '930';    // Bordereau trimestriel
      if (Periodicite = 'A') then
        TypePeriod := '932';    // Bordereau annuel
    end;
  Result := TypePeriod;
end;
// f PT1

end.


