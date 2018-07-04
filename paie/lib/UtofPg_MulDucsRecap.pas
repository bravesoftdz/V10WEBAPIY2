{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 26/03/2003
Modifié le ... : 03/04/2003
Description .. : Edition de DUCS Récapitulatives : Destiné aux ducs dossier
Suite ........ : afin d'obtenir un état tous établissements (toutes ducs)
Suite ........ : confondus.
Suite ........ : Le traitement est effectué à partir des informations stockées
Suite ........ : dans DUCSENTETE et DUCSDETAIL. (La Ducs n'est pas
Suite ........ : calculée à partir de HISTOBULLETIN)
Mots clefs ... : PAIE , PGDUCS
*****************************************************************}
{
 PT1 : 15/09/2003 :  V_421 MF  Correction FQ 10788 : Traitement des raccourcis
                               clavier (Dates de période)

 PT2 : 18/09/2003 : V_421 MF   PGIBox remplace MessageAlerte
 PT3 : 02/10/2003 : V_421 MF   Mise au point CWAS
 PT4 : 12/02/2006 : V_702 MF   FQ 13070 : Traitement des codification Alsace-Moselle
}
unit UtofPg_MulDucsRecap;

interface
uses
        {$IFDEF VER150}
        Variants,
        {$ENDIF}
{$IFDEF EAGLCLIENT}
        UtileAGL,eMul,
{$ELSE}
        HDB,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Mul,EdtREtat,
{$ENDIF}
        Classes,UTOF,HCtrls,HTB97,
//unused        Hqry,
        PgOutils,PgOutils2,sysutils,HEnt1,
        ParamDat,hmsgbox,P5Util,UTob,ComCtrls,ed_tools,HStatus,ULibEditionPaie;
Type
    TParamOrg = record
      Etab                      : String;
      Organisme                 : String;
      Pdu_Siret                 : String;
      Pdu_Ape                   : String;
      Pdu_Groupe                : String;
      Pdu_Numero                : String;
      Pdu_AbregePeriode         : String;
      Pdu_DateExigible          : TDateTime;
      Pdu_DateLimDepot          : TDateTime;
      Pdu_DateReglement         : TDateTime;
      Pdu_Paiement              : TDateTime;
      Pdu_Cessation             : TDateTime;
      Pdu_Continuation          : TDateTime;
      Pdu_suspension            : Boolean;
      Pdu_Maintient             : Boolean;
      Pdu_Declarant             : String;
      Pdu_DeclarantSuite        : String;
      Pdu_TelephoneDecl         : String;
      Pdu_FaxDeclarant          : String;
      Pdu_LigneOptique          : String;
      Pdu_Acompte               : Double;
      Pdu_Regularisation        : Double;
      Pdu_MonnaieTenue          : String;
      Pdu_Nbsal_Fpe             : Integer;
      Pdu_Nbsal_Hap             : Integer;
      Pdu_Nbsal_Q920            : Integer;
      Pdu_Nbsal_Q921            : Integer;
      Pdu_Nbsal_Q922            : Integer;
      Pdu_Nbsal_Q923            : Integer;
      Pdu_Nbsal_Q924            : Integer;
      Pdu_Nbsal_Q925            : Integer;
      Pdu_Nbsal_Q926            : Integer;
      Pdu_Nbsal_Q927            : Integer;
      Pdu_Nbsal_Q928            : Integer;
      Pdu_Nbsal_Q929            : Integer;
      Pdu_Nbsal_Q934            : Integer;
      Pdu_Nbsal_Q935            : Integer;
      Pdu_Nbsal_Q936            : Integer;
      Pdu_Nbsal_Q937            : Integer;
      Pdu_TotHommes             : Integer;
      Pdu_TotFemmes             : Integer;
      Pdu_TotApprenti              : Integer;
      Pdu_TotOuvrier            : Integer;
      Pdu_TotEtam               : Integer;
      Pdu_TotCadres             : Integer;
      Pdu_EmettSoc              : String;
      Pdu_DucsDossier           : Boolean;
      Et_Libelle                : String;
      Et_Adresse1               : String;
      Et_Adresse2               : String;
      Et_Adresse3               : String;
      Et_CodePostal             : String;
      Et_Ville                  : String;
      Et_Telephone              : String;
      Et_Fax                    : String;
      Pog_Libelle               : String;
      Pog_NatureOrg             : String;
      Pog_Adresse1              : String;
      Pog_Adresse2              : String;
      Pog_Adresse3              : String;
      Pog_CodePostal            : String;
      Pog_PeriodicitDucs        : String;
      Pog_AutPerioDucs          : String;
      Pog_PeriodCalcul          : Boolean;
      Pog_AutPerCalcul          : Boolean;
      Pog_Ville                 : String;
      Pog_LongEditable          : Integer;
      Pog_BaseTypArr            : String;
      Pog_MtTypArr              : String;
      Pog_SousTotDucs           : Boolean;
      Pog_LongTotal             : Integer;
      Pog_PosTotal              : integer;
    end;

type
    TOF_PGMULDUCSRECAP = Class (TOF)
    public
      procedure OnArgument(Arguments : String ) ; override ;
      procedure OnLoad;  override ;

    private
      DateDeb,DateFin : THEdit;
      DebPer,FinPer,NatDeb,PeriodEtat : string;
      DebTrim,FinTrim,DebExer,FinExer,DebSem,FinSem : TDateTime;
      NbreMois : integer;
      WW : THEdit;
      procedure ActiveWhere (Okok : Boolean);
      procedure DateElipsisclick(Sender: TObject);
      procedure Change(Sender: TObject);
      procedure ChercheClick(Sender: TObject);
      procedure CalculerClick(Sender: TObject);
      procedure ChargeOrganismes(NatureDucs : String;var TOrg : TOB);
      procedure ChargeWheres(var StWhereEd : String; TOrg : TOB);
      procedure ChargeDucs(StWhereEd : string;var TDucs : TOB);
      procedure CreationTOBEtat(TDucs : TOB;var TOBEtat, TOBEtatFille, TOBEtatNeant : TOB);
      procedure RecupParamOrganisme(Nature : String;TD : TOB;
                                    var POrg : TParamOrg);
      procedure ChargeLignesATraiter(var POrg : TParamOrg; TD : TOB;var TDL : TOB;
                                     var EtabEnCours : string; var Neant : boolean);
      procedure MajTobEtat (var TDL : TOB;var TOBEtat : TOB; POrg : TParamOrg;
                            NumOrdre : integer);
      procedure ConstructionTOBEtat(TDL : TOB;var TOBEtat :TOB;
                                    CasTraitement, NumOrdre : integer;
                                    POrg : TParamOrg);
      procedure CreationTOBFille (TOrig : TOB; var TOBDest, TOBDestFille : TOB;
                                  var MontST : double;
                                  Neant : boolean);
      procedure CreationTOBEtatST(CasTraitement : integer;TDLFille : TOB;var TOBEtat,TOBEtatFille : TOB;
                                  CodificationLue,CodificationEnCours,
                                  InstitutionEnCours : String;
                                  POrg : TParamOrg;var Numordre : integer;
                                  var MontST : double);
      procedure MajTobEtatFille (TDLFille : TOB;var TOBEtat,TOBEtatFille : TOB;
                                 var MontSt : double;CasTraitement : integer);
      procedure CreationTOBEtatAT(TDL : TOB;var TOBEtat : TOB;
                                  var MontSt : double);
      procedure MajTobEtatNeant (var TDL : TOB;
                                 var TOBEtatNeant : TOB;
                                 POrg : TParamOrg;
                                 NumOrdre : integer);
    END;

implementation
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 03/04/2003
Modifié le ... : 03/04/2003
Description .. : Récupération de l'établissement par défaut.
Suite ........ : Détermination de la Date par défaut
Suite ........ : Initialisations diverses
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.OnArgument(Arguments: String);
var
     ExerPerEncours,MoisE,AnneeE : string;
     Ouvrir, Cherche             : TToolbarButton97;
begin

inherited ;

  WW:=THEdit (GetControl ('XX_WHERE'));
// d PT1
  DateDeb:= THEdit(getcontrol('XX_VARIABLED'));
  DateFin:= THEdit(getcontrol('XX_VARIABLED_'));
// f PT1
  if (DateDeb <> NIL) and (DateFin <> NIL) then
  begin
    DateDeb.OnElipsisClick := DateElipsisclick;
    DateDeb.OnExit:=Change;
    DateFin.OnElipsisClick := DateElipsisclick;
    DateFin.OnExit:=Change;
  end;

{ Date par défaut : La date proposée est le trimestre en cours si la période
  en cours correspond à une fin de trimestre, sinon c'est le mois en cours.}
  if RendExerSocialEnCours(MoisE,AnneeE,ExerPerEncours,DebExer,FinExer)=True then
  begin
    RendPeriodeEnCours(ExerPerEnCours,DebPer,FinPer);
    RendTrimestreEnCours(StrToDate(DebPer),DebExer,FinExer,DebTrim,FinTrim,DebSem,FinSem);
    if FindeMois(StrToDate(DebPer)) <> FinTrim then
    { mois en cours}
    begin
      if  DateDeb <> NIL then DateDeb.text:=DateToStr(StrToDate(DebPer));
      if  DateFin <> NIL then DateFin.text:=DateToStr(FindeMois(StrToDate(DebPer)));
    end
    else
    { trimestre en cours}
    begin
      if  DateDeb <> NIL then DateDeb.text:=DateToStr(DebTrim);
      if  DateFin <> NIL then DateFin.text:=DateToStr(FinTrim);
    end;
  end;

  Ouvrir := TToolbarButton97(GetControl('BOUVRIR'));
  if Ouvrir <> NIL then
  begin
    Ouvrir.Visible := True;
    Ouvrir.OnClick := CalculerClick;
  end;

  Cherche := TToolbarButton97(GetControl('BCHERCHE'));
  if Cherche <> NIL then
  begin
    Cherche.OnClick := ChercheClick;
  end;
end;  { fin OnArgument}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 03/04/2003
Modifié le ... : 03/04/2003
Description .. : OnLoad : Chargement de la fiche
Suite ........ : On détermine la périodicité des DUCS à initialiser
Suite ........ : On lance l'ActiveWhere
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.OnLoad;
var
   Okok : Boolean;
   NbreJour : Integer;
begin
inherited ;
{ Récup. des critères de sélection}
// d PT1
  DebPer := GetControlText('XX_VARIABLED');
  FinPer := GetControlText('XX_VARIABLED_');
// f PT1
  NatDeb := GetControlText('POG_NATUREORG');

  { On détermine la périodicité des DUCS à initialiser
  Mensuelle, Trimestrielle, Semestrielle, Annuelle ou Quelconque
  Remarque : Il faut que la période corresponde à un mois ou un trimestre
  ou un semestre de l'exercice (Par exemple Février/mars/avril n'est pas un
  trimestre) sinon elle sera considérée comme quelconque (PeriodEtat = '')}
  DiffMoisJour (StrToDate(DebPer),StrToDate(FinPer),NbreMois,NbreJour);
  NbreMois := NbreMois + 1;

  if (NbreMois <> 1) and
     (NbreMois <> 3) and
     (NbreMois <> 6) and
     (NbreMois <> 12) then
  begin
    MessageAlerte('La période sélectionnée n''est pas un mois, '+
                  '#13#10 ni un trimestre, ni un semestre, ni une année ');
  end;

  PeriodEtat := 'M';
  if (NbreMois = 3) then PeriodEtat := 'T';
  if (NbreMois = 6) then PeriodEtat := 'S';
  if (NbreMois = 12) then PeriodEtat := 'A';

  Okok := TRUE;
  ActiveWhere (Okok);
end;  { fin OnLoad}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 26/03/2003
Modifié le ... :   /  /    
Description .. : procédure de mise en forme des critères de sélection
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.ActiveWhere(Okok : Boolean);
var
   St : STring;

begin
  WW.Text := '';
  st := '';
  St := '((POG_PERIODICITDUCS="'+PeriodEtat+'") OR '+
        '((POG_AUTREPERIODUCS<>"") AND (POG_AUTREPERIODUCS="'+PeriodEtat+'"))) '+
        'AND'+
        '((POG_CAISSEDESTIN="X") AND (POG_DUCSDOSSIER="X"))';

  if St <> '' then WW.Text := st;
end;  {fin ActiveWhere}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 26/03/2003
Modifié le ... :   /  /
Description .. : Ouverture calendrier quand clck sur date
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.DateElipsisclick(Sender: TObject);
var
   key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;  {fin DateElipsisclick}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 26/03/2003
Modifié le ... :   /  /    
Description .. : Contrôle de la validité des critères "date" en cas de
Suite ........ : modification
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.Change(Sender: TObject);
begin
  if not IsValidDate(GetControlText('XX_VARIABLED'))then // PT1
  {Pour générer message erreur si date erronnée}
  begin
    PGIBox('La date de début est erronée.',Ecran.caption);
    SetControlText('XX_VARIABLED',DatetoStr(Date));     // PT1
  end;
  if not IsValidDate(GetControlText('XX_VARIABLED_'))then       // PT1
  {Pour générer message erreur si date erronnée}
  begin
    PGIBox('La date de fin est erronée.',Ecran.caption);
    SetControlText('XX_VARIABLED_',DatetoStr(Date));    // PT1
  end;
end;  {fin Change}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 26/03/2003
Modifié le ... :   /  /
Description .. : procédure activée quand "Mouette bleue"
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.ChercheClick(Sender: TObject);
begin
  if not IsValidDate(GetControlText('XX_VARIABLED'))then // PT1
  {Pour générer message erreur si date erronnée }
  begin
   PGIBox('La date de début est erronée.',Ecran.caption);
   SetControlText('XX_VARIABLED',DatetoStr(Date)); // PT1
  end;
  if not IsValidDate(GetControlText('XX_VARIABLED_'))then // PT1
  {Pour générer message erreur si date erronnée}
  begin
   PGIBox('La date de fin est erronée.',Ecran.caption);
   SetControlText('XX_VARIABLED_',DatetoStr(Date));  // PT1
  end;
  TFMul(Ecran).BChercheClick(Nil);
end; {fin ChercheClick}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 26/03/2003
Modifié le ... : 03/04/2003
Description .. : Procédure activée après sélection des DUCS à traiter et
Suite ........ : validation (mouette verte)
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.CalculerClick(Sender: TObject);
Var

{$IFDEF EAGLCLIENT}
   Liste : THGrid;
{$ELSE}
   Liste : THDBGrid;
{$ENDIF}
   Pages                                             : TPageControl;
   i , IndOR                                         : integer;
   NatureDucs                                        : String;
   Regroupt,Etab, Organisme                          : String;
   StWhereEd                                         : String;
   TOrg, TDucs, TOBEtat, TOBEtatFille,TOBEtatNeant   : TOB;
begin
  StWhereEd := '';
  Pages := TPageControl(GetControl('Pages'));
  i := 0;
  IndOR := 0;
  TOBEtat := NIL;
  TOBEtatNeant := NIL; 
  
{$IFDEF EAGLCLIENT}
  Liste := THGrid(GetControl('FListe'));
{$ELSE}
  Liste := THDBGrid(GetControl('FListe'));
{$ENDIF}

  if Liste <> NIL then
    if (Liste.NbSelected=0) and (not Liste.AllSelected) then
    begin
      PgiBox('Aucun élément sélectionné',Ecran.Caption); // PT2
      exit;
    end;

  StWhereEd := 'WHERE ('+
            '((PDU_DATEDEBUT ="'+UsDateTime(StrToDate(DebPer))+'") AND '+
             '(PDU_DATEFIN = "'+UsDateTime(StrToDate(FinPer))+'")) AND (';

  { Sélection des DUCS à Editer}
  if Liste <> NIL then
  begin
    if (Liste.AllSelected=TRUE) then
    begin
      InitMoveProgressForm (NIL,'Edition en cours',
                                'Veuillez patienter SVP ...',i,FALSE,TRUE);
      InitMove(TFmul(Ecran).Q.RecordCount,'');
// d PT3
{$IFDEF EAGLCLIENT}
      if (TFMul(Ecran).bSelectAll.Down) then
         TFMul(Ecran).Fetchlestous;
{$ENDIF}
 // f PT3
      TFmul(Ecran).Q.First;
      while not TFmul(Ecran).Q.EOF do
      begin
        MoveCur(False);
        Regroupt:=TFmul(Ecran).Q.FindField('POG_REGROUPEMENT').asstring;
        Etab:=TFmul(Ecran).Q.FindField('POG_ETABLISSEMENT').asstring;
        Organisme:=TFmul(Ecran).Q.FindField('POG_ORGANISME').asstring;
        NatureDucs := TFmul(Ecran).Q.FindField('POG_NATUREORG').asstring; // PT33-1

        {On établit la liste des organismes rattachées
         à l'organisme destinataire (même nature & ducs dossier cochée}
        ChargeOrganismes(NatureDucs,Torg);

        { Mémorisation en vue d'édition et de mise à jour des critères de
          sélection des différentes DUCS initialisées}
        if (IndOR <> 0) then
        begin
          StWhereEd := StWhereEd + ' OR ';
        end;

        {Initialisation StWhereEd" en fonctions des organismes de TOrg}
        ChargeWheres(StWhereEd , TOrg);

        IndOR := 1;

        { free des TOBs utilisées}
        TOrg.free;

        MoveCurProgressForm(Organisme);
        TFmul(Ecran).Q.Next;
      end;

      Liste.AllSelected:=False;
      FiniMove;
      FiniMoveProgressForm;
      TFMul(Ecran).bSelectAll.Down := False;

    end
    else
    begin
      InitMoveProgressForm (NIL,'Initialisation en cours',
                            'Veuillez patienter SVP ...',i,FALSE,TRUE);
      InitMove(Liste.NbSelected,'');
      for i:=0 to Liste.NbSelected-1 do
      begin
        Liste.GotoLeBOOKMARK(i);
// d PT3
{$IFDEF EAGLCLIENT}
          TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row-1) ;
{$ENDIF}
// f PT3

        MoveCur(False);
        Regroupt:=TFmul(Ecran).Q.FindField('POG_REGROUPEMENT').asstring;
        Etab:=TFmul(Ecran).Q.FindField('POG_ETABLISSEMENT').asstring;
        Organisme:=TFmul(Ecran).Q.FindField('POG_ORGANISME').asstring;
        NatureDucs := TFmul(Ecran).Q.FindField('POG_NATUREORG').asstring;

        {On établit la liste des organismes rattachées
         à l'organisme destinataire (même nature & ducs dossier cochée}
        ChargeOrganismes(NatureDucs,Torg);

        {Mémorisation des critères de sélection des différentes DUCS }
        if (IndOR <> 0) then
        begin
          StWhereEd := StWhereEd + 'OR ';
        end;

        {Initialisation StWhereEd" en fonctions des organismes de TOrg}
        ChargeWheres(StWhereEd , TOrg);

        IndOR := 1;

        { free des TOBs utilisées}
        TOrg.free;

        MoveCurProgressForm(Organisme);
      end;

      Liste.ClearSelected;
      FiniMove;
      FiniMoveProgressForm;
    end;
  end;
  StWhereEd := StWhereEd + ')) ';

  {Chargement des enregistrements DUCENTETE et DUCSDETAIL}
  ChargeDucs(StWhereEd, TDucs);

  if (TDucs.Detail.Count <> 0) then
  begin
   {Construction des tob utilisées pour l'édition TOBEtat & TOBEtatNeant}
   CreationTOBEtat(TDucs, TOBEtat,TOBEtatFille,TOBEtatNeant);

   {lancement de l'édition}
   if (TobEtatNeant <> NIL) then
   begin
     LanceEtatTOB('E','PDU','DUN',TobEtatNeant,True,False,False,Pages,'','',False);
     TOBEtatNeant.free;
     TOBEtatNeant := NIl;
   end;
  
   {lancement de l'édition}
   if (TobEtat <> NIL) then
   begin
{   if V_PGI.Debug=True then
   PGVisuUnObjet( TobEtat,'','') ;}
     LanceEtatTOB('E','PDU','DUR',TobEtat,True,False,False,Pages,'','',False);
     TOBEtat.free;
     TOBEtat := NIL;
   end;
  end;

  TDucs.free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 26/03/2003
Modifié le ... :   /  /    
Description .. : Pour un élément de la liste sélectionnée recherche pour les
Suite ........ : différents établissements  des organismes de même nature
Suite ........ : rattachés à l'organisme destinataire.
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.ChargeOrganismes(NatureDucs :String;var TOrg : TOB);
var
   st           : string;
   QOrg         : TQuery;
begin
  st := ' ';

  st := 'SELECT POG_ETABLISSEMENT, POG_ORGANISME,POG_DUCSDOSSIER,POG_NATUREORG,'+
        'POG_CAISSEDESTIN, POG_SOUSTOTDUCS, POG_LONGTOTAL '+
        'FROM ORGANISMEPAIE '+
        'WHERE POG_DUCSDOSSIER = "X" AND POG_NATUREORG ="'+NatureDucs+'"';

  QOrg := OpenSql(st,TRUE);
  TOrg := TOB.Create('Les organismes', NIL, -1);
  TOrg.LoadDetailDB('TABLEORGANISMEPAIE','','',QOrg,False);
  Ferme(QOrg);

end;  {fin ChargeOragnismes}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 26/03/2003
Modifié le ... :   /  /
Description .. : Mise en forme du "WHERE" utilisé dasn la requête
Suite ........ : d'extraction des lignes de DUCS.
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.ChargeWheres(var StWhereEd :  String; TOrg : TOB);
var
  EtabEnCours,organismeEnCours  : String ;
  NbOrg                         : integer;
  LesOrg                        : TOB;
begin
  for NbOrg := 0 to TOrg.Detail.Count-1 do
  begin
    LesOrg := TOrg.Detail[NbOrg];
    EtabEnCours := LesOrg.GetValue('POG_ETABLISSEMENT');
    organismeEnCours := LesOrg.GetValue('POG_ORGANISME');
    if (NbOrg <> 0) then
      StWhereEd := StWhereEd + ' OR ';
    StWhereEd := StWhereEd +
             '((PDU_ETABLISSEMENT ="'+ EtabEnCours +'") AND '+
             '(PDU_ORGANISME ="'+ organismeEnCours+'")) ';
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 26/03/2003
Modifié le ... : 03/04/2003
Description .. : Création de la TOB TDucs : liste des lignes de Ducs
Suite ........ : répondant aux critères de sélection.
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.ChargeDucs(StWhereEd : string; var TDucs : TOB);
var
  st            : string;
  QDucs         : TQuery;
  Nblignes      : integer;
begin
   st := 'SELECT DUCSENTETE.*,PDD_CODIFICATION,PDD_CODIFEDITEE, '+
          'PDD_LIBELLE, PDD_BASECOTISATION,PDD_TAUXCOTISATION, PDD_MTCOTISAT, '+
          'PDD_EFFECTIF, PDD_TYPECOTISATION, PDD_INSTITUTION, '+
          'PDD_LIBELLESUITE, '+
          'PDD_REGIMEALSACE, '+ // PT4
          'ET_ETABLISSEMENT, ET_LIBELLE, ET_ADRESSE1, ET_ADRESSE2, ET_ADRESSE3, '+
          'ET_CODEPOSTAL, ET_VILLE, ET_TELEPHONE, ET_FAX, '+
          'POG_ETABLISSEMENT, POG_ORGANISME, POG_LIBELLE, POG_NATUREORG, '+
          'POG_ADRESSE1, POG_ADRESSE2, POG_ADRESSE3, POG_CODEPOSTAL, '+
          'POG_PERIODICITDUCS, POG_AUTREPERIODUCS, POG_PERIODCALCUL, '+
          'POG_AUTPERCALCUL, '+
          'POG_POSTOTAL,POG_LONGTOTAL, '+
          'POG_VILLE , POG_LONGEDITABLE,POG_BASETYPARR,'+
          'POG_MTTYPARR '+
          'FROM DUCSENTETE '+
          'LEFT JOIN ETABLISS ON PDU_ETABLISSEMENT = ET_ETABLISSEMENT '+
          'LEFT JOIN ORGANISMEPAIE ON PDU_ETABLISSEMENT = POG_ETABLISSEMENT '+
          'AND PDU_ORGANISME = POG_ORGANISME '+
          'LEFT JOIN DUCSDETAIL ON  PDU_ETABLISSEMENT =  PDD_ETABLISSEMENT '+
          'AND PDU_ORGANISME = PDD_ORGANISME AND PDU_DATEDEBUT = PDD_DATEDEBUT '+
          'AND PDU_DATEFIN = PDD_DATEFIN AND PDU_NUM = PDD_NUM '+
          StWhereEd+
          'ORDER BY POG_NATUREORG,POG_CAISSEDESTIN DESC,PDU_ORGANISME,PDU_ETABLISSEMENT,PDU_DATEDEBUT,PDU_DATEFIN,'+
          'PDU_NUM,PDD_INSTITUTION,PDD_CODIFICATION' ;
  QDucs := OpenSql(st,TRUE);
  TDucs := TOB.Create('Les ducs', NIL, -1);
  TDucs.LoadDetailDB('TABLEORGANISMEPAIE','','',QDucs,False);
  Ferme(QDucs);
//  NbLignes := TDucs.Detail.Count;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 26/03/2003
Modifié le ... :   /  /    
Description .. : Création de la TOBEtat qui servira à l'édition.
Suite ........ : 
Suite ........ : La sélection de  plusieurs Ducs de différentes natures et les
Suite ........ : différents paramètrages possibles au niveau organisme
Suite ........ : (Avec ou sans Sous-Total, Standardisée ou
Suite ........ : Personnalisée...) oblige à l'utilisation de TOBs intermédiaires
Suite ........ : avant la création de la TOB Définitive.
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.CreationTOBEtat(TDucs : TOB; var TOBEtat, TOBEtatFille, TOBEtatNeant : TOB);
var
  TD, TDL                                       : TOB;
  Nature, NatureEnCours, EtabEnCours            : String;
  POrg                                          : TParamOrg;
  PremierPassage, Neant                         : Boolean;
  NumOrdre,I                                    : integer;

begin
  Nature := '';
  NatureEnCours := '';
  EtabEnCours := '';
  PremierPassage := True;
  Numordre := 0;
  Neant := True;

{if V_PGI.Debug=True then
   PGVisuUnObjet( TDucs,'','');}
   

  TD :=  TDucs.FindFirst([''],[''], TRUE);

  for I:= 0 to TDucs.Detail.Count-1 do
  begin
    Nature := TD.GetValue('POG_NATUREORG');
    if (Nature =  NatureEnCours) then
    begin
     ChargeLignesATraiter(POrg,TD, TDL, EtabEnCours, Neant);
    end
    else
    begin
     if (PremierPassage = False) then
     begin
       if (Neant = True) then
         MajTobEtatNeant(TDL, TOBEtatNeant, POrg, NumOrdre)
       else
         MajTobEtat (TDL, TOBEtat, POrg, NumOrdre);
       EtabEnCours := '';
       Neant := True;
       if (TDL <> NIL) then
       begin
         TDL.free;
         TDL := NIL;
       end;
     end;
     EtabEnCours := '';
     PremierPassage := False;
     NatureEnCours := Nature;
     RecupParamOrganisme(Nature,TD,POrg);
     ChargeLignesATraiter(POrg,TD, TDL,EtabEnCours, Neant);
     NumOrdre := 0;
    end;

    TD :=  TDucs.FindNext([''],[''], TRUE);
  end;

  if (Neant = True) then
    MajTobEtatNeant(TDL, TOBEtatNeant, POrg, NumOrdre)
  else
    MajTobEtat (TDL, TOBEtat, POrg, NumOrdre);

  if (TDL <> NIL) then
    begin
      TDL.free;
      TDL := NIL;
    end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 03/04/2003
Modifié le ... : 03/04/2003
Description .. : Récupération des caractéristiques ORGANISMEPAIE
Suite ........ : de la caisse destinataire
Suite ........ : Récupération des caractéristiques , ETABLISSEMENT,
Suite ........ : ORGANISME de la caisse  destinataire
Mots clefs ... : PAIE PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.RecupParamOrganisme(Nature : String; TD : TOB; var POrg : TParamOrg);
var
  TOrg,TOrgDest                          : TOB;
begin
  { Récupération des caractéristiques ORGANISMEPAIE de la caisse destinataire}
  ChargeOrganismes(Nature,Torg);
  TOrgDest := Torg.FindFirst (['POG_CAISSEDESTIN'],['X'], TRUE);
  if (TOrgDest <> NIL) then
    begin
      POrg.Etab := TOrgDest.GetValue('POG_ETABLISSEMENT');
      POrg.Organisme := TOrgDest.GetValue('POG_ORGANISME');
      POrg.Pog_SousTotDucs := TOrgDest.GetValue('POG_SOUSTOTDUCS')  = 'X';
      POrg.Pog_LongTotal := TOrgDest.GetValue('POG_LONGTOTAL');
    end;
  Torg.free;

  { Récupération des caractéristiques , ETABLISSEMENT, ORGANISME de la caisse
    destinataire}
  if (TD <> NIL) then
    begin
      POrg.Pdu_Siret := TD.GetValue('PDU_SIRET');
      POrg.Pdu_Ape  := TD.GetValue('PDU_APE');
      POrg.Pdu_Groupe := TD.GetValue('PDU_GROUPE');
      POrg.Pdu_Numero := TD.GetValue('PDU_NUMERO');
      POrg.Pdu_AbregePeriode := TD.GetValue('PDU_ABREGEPERIODE');
      POrg.Pdu_DateExigible := TD.GetValue('PDU_DATEEXIGIBLE');
      POrg.Pdu_DateLimDepot := TD.GetValue('PDU_DATELIMDEPOT');
      POrg.Pdu_DateReglement := TD.GetValue('PDU_DATEREGLEMENT');
      POrg.Pdu_Paiement := TD.GetValue('PDU_PAIEMENT');
      POrg.Pdu_Cessation := TD.GetValue('PDU_CESSATION');
      POrg.Pdu_Continuation := TD.GetValue('PDU_CONTINUATION');
      POrg.Pdu_suspension := TD.GetValue('PDU_SUSPENSION') = 'X';
      POrg.Pdu_Maintient := TD.GetValue('PDU_MAINTIENT') = 'X';
      POrg.Pdu_Declarant := TD.GetValue('PDU_DECLARANT');
      POrg.Pdu_DeclarantSuite := TD.GetValue('PDU_DECLARANTSUITE');
      POrg.Pdu_TelephoneDecl := TD.GetValue('PDU_TELEPHONEDECL');
      POrg.Pdu_FaxDeclarant := TD.GetValue('PDU_FAXDECLARANT');
      POrg.Pdu_LigneOptique := TD.GetValue('PDU_LIGNEOPTIQUE');
      POrg.Pdu_Acompte := TD.GetValue('PDU_ACOMPTES');
      POrg.Pdu_Regularisation := TD.GetValue('PDU_REGULARISATION');
      POrg.Pdu_MonnaieTenue := TD.GetValue('PDU_MONNAIETENUE');
      POrg.Pdu_Nbsal_Fpe := TD.GetValue('PDU_NBSALFPE');
      POrg.Pdu_Nbsal_Hap := TD.GetValue('PDU_NBSALHAP');
      POrg.Pdu_Nbsal_Q920 := TD.GetValue('PDU_NBSALQ920');
      POrg.Pdu_Nbsal_Q921 := TD.GetValue('PDU_NBSALQ921');
      POrg.Pdu_Nbsal_Q922 := TD.GetValue('PDU_NBSALQ922');
      POrg.Pdu_Nbsal_Q923 := TD.GetValue('PDU_NBSALQ923');
      POrg.Pdu_Nbsal_Q924 := TD.GetValue('PDU_NBSALQ924');
      POrg.Pdu_Nbsal_Q925 := TD.GetValue('PDU_NBSALQ925');
      POrg.Pdu_Nbsal_Q926 := TD.GetValue('PDU_NBSALQ926');
      POrg.Pdu_Nbsal_Q927 := TD.GetValue('PDU_NBSALQ927');
      POrg.Pdu_Nbsal_Q928 := TD.GetValue('PDU_NBSALQ928');
      POrg.Pdu_Nbsal_Q929 := TD.GetValue('PDU_NBSALQ929');
      POrg.Pdu_Nbsal_Q934 := TD.GetValue('PDU_NBSALQ934');
      POrg.Pdu_Nbsal_Q935 := TD.GetValue('PDU_NBSALQ935');
      POrg.Pdu_Nbsal_Q936 := TD.GetValue('PDU_NBSALQ936');
      POrg.Pdu_Nbsal_Q937 := TD.GetValue('PDU_NBSALQ937');
      POrg.Pdu_TotHommes := TD.GetValue('PDU_TOTHOMMES') ;
      POrg.Pdu_TotFemmes := TD.GetValue('PDU_TOTFEMMES') ;
      POrg.Pdu_TotApprenti := TD.GetValue('PDU_TOTAPPRENTI');
      POrg.Pdu_TotOuvrier := TD.GetValue('PDU_TOTOUVRIER');
      POrg.Pdu_TotEtam := TD.GetValue('PDU_TOTETAM');
      POrg.Pdu_TotCadres :=TD.GetValue('PDU_TOTCADRES');
      POrg.Pdu_EmettSoc := TD.GetValue('PDU_EMETTSOC');
      POrg.Pdu_DucsDossier := TD.GetValue('PDU_DUCSDOSSIER') = 'X';
      POrg.Et_Libelle := TD.GetValue('ET_LIBELLE');
      POrg.Et_Adresse1 := TD.GetValue('ET_ADRESSE1');
      POrg.Et_Adresse2 := TD.GetValue('ET_ADRESSE2');
      POrg.Et_Adresse3 := TD.GetValue('ET_ADRESSE3');
      POrg.Et_CodePostal := TD.GetValue('ET_CODEPOSTAL');
      POrg.Et_Ville := TD.GetValue('ET_VILLE');
      POrg.Et_Telephone := TD.GetValue('ET_TELEPHONE');
      POrg.Et_Fax := TD.GetValue('ET_FAX');
      POrg.Pog_Libelle := TD.GetValue('POG_LIBELLE');
      POrg.Pog_NatureOrg := TD.GetValue('POG_NATUREORG');
      POrg.Pog_Adresse1 := TD.GetValue('POG_ADRESSE1');
      POrg.Pog_Adresse2 := TD.GetValue('POG_ADRESSE2');
      POrg.Pog_Adresse3 := TD.GetValue('POG_ADRESSE3');
      POrg.Pog_CodePostal := TD.GetValue('POG_CODEPOSTAL');
      POrg.Pog_PeriodicitDucs := TD.GetValue('POG_PERIODICITDUCS');
      POrg.Pog_AutPerioDucs := TD.GetValue('POG_AUTREPERIODUCS');
      POrg.Pog_PeriodCalcul := TD.GetValue('POG_PERIODCALCUL') = 'X';
      POrg.Pog_AutPerCalcul := TD.GetValue('POG_AUTPERCALCUL') = 'X';
      POrg.Pog_PosTotal := TD.GetValue('POG_POSTOTAL');
      POrg.Pog_LongTotal := TD.GetValue('POG_LONGTOTAL');
      POrg.Pog_Ville:= TD.GetValue('POG_VILLE');
      POrg.Pog_LongEditable := TD.GetValue('POG_LONGEDITABLE');
      POrg.Pog_BaseTypArr := TD.GetValue('POG_BASETYPARR');
      POrg.Pog_MtTypArr := TD.GetValue('POG_MTTYPARR');
    end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 26/03/2003
Modifié le ... :   /  /
Description .. : Création de la TOB TDL qui se limite aux lignes d'une Ducs.
Suite ........ :
Suite ........ : Au fur et à mesure de son alimentation le cumull des
Suite ........ : champs assimilés à des totaux alimentant  l'en-tête est tenu.
Suite ........ : (effectifs, acomptes, régularisation...)
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.ChargeLignesATraiter(var POrg : TParamOrg; TD : TOB; var TDL : TOB;var EtabEnCours : string; var Neant : boolean);
var
  TDLFille                              : TOB;
  etab                                  : string;

begin
  if (TDL = NIL)then
  begin
   TDL:=TOB.Create('les lignes DUCS',NIL,-1) ;
  end;
  etab := TD.GetValue('PDU_ETABLISSEMENT');
  TDLFille:=TOB.Create('',TDL,-1) ;
  TDLFille.AddChampSupValeur('PDU_ETABLISSEMENT',POrg.Etab,FALSE) ;
  TDLFille.AddChampSupValeur('PDU_ORGANISME',POrg.Organisme,FALSE) ;
  TDLFille.AddChampSupValeur('PDU_DATEDEBUT',StrToDate(DebPer) ,FALSE) ;
  TDLFille.AddChampSupValeur('PDU_DATEFIN',StrToDate(FinPer),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_SIRET',POrg.Pdu_Siret,FALSE) ;
  TDLFille.AddChampSupValeur('PDU_APE',POrg.Pdu_Ape,FALSE) ;
  TDLFille.AddChampSupValeur('PDU_GROUPE',POrg.Pdu_Groupe,FALSE) ;
  TDLFille.AddChampSupValeur('PDU_NUMERO',POrg.Pdu_Numero,FALSE) ;
  TDLFille.AddChampSupValeur('PDU_ABREGEPERIODE',POrg.Pdu_AbregePeriode,FALSE) ;
  TDLFille.AddChampSupValeur('PDU_DATEEXIGIBLE',POrg.Pdu_DateExigible,FALSE) ;
  TDLFille.AddChampSupValeur('PDU_DATELIMDEPOT',POrg.Pdu_DateLimDepot,FALSE) ;
  TDLFille.AddChampSupValeur('PDU_DATEREGLEMENT',POrg.Pdu_DateReglement,FALSE) ;
  TDLFille.AddChampSupValeur('PDU_PAIEMENT',POrg.Pdu_Paiement,FALSE) ;
  TDLFille.AddChampSupValeur('PDU_CESSATION',POrg.Pdu_Cessation,FALSE) ;
  TDLFille.AddChampSupValeur('PDU_CONTINUATION',POrg.Pdu_Continuation,FALSE) ;
  if (POrg.Pdu_suspension = true) then
    TDLFille.AddChampSupValeur('PDU_SUSPENSION','X',FALSE)
  else
    TDLFille.AddChampSupValeur('PDU_SUSPENSION','-',FALSE);
  if (POrg.Pdu_Maintient = true) then
     TDLFille.AddChampSupValeur('PDU_MAINTIENT','X',FALSE)
  else
     TDLFille.AddChampSupValeur('PDU_MAINTIENT','-',FALSE);
  TDLFille.AddChampSupValeur('PDU_DECLARANT',POrg.Pdu_Declarant,FALSE) ;
  TDLFille.AddChampSupValeur('PDU_DECLARANTSUITE',POrg.Pdu_DeclarantSuite,FALSE) ;
  TDLFille.AddChampSupValeur('PDU_TELEPHONEDECL',POrg.Pdu_TelephoneDecl,FALSE) ;
  TDLFille.AddChampSupValeur('PDU_FAXDECLARANT',POrg.Pdu_FaxDeclarant,FALSE) ;
  TDLFille.AddChampSupValeur('PDU_LIGNEOPTIQUE',POrg.Pdu_LigneOptique,FALSE) ;
  TDLFille.AddChampSupValeur('PDU_ACOMPTES',TD.GetValue('PDU_ACOMPTES'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_REGULARISATION',TD.GetValue('PDU_REGULARISATION'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_MONNAIETENUE',POrg.Pdu_MonnaieTenue,FALSE) ;
  TDLFille.AddChampSupValeur('PDU_NBSALFPE',TD.GetValue('PDU_NBSALFPE'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_NBSALHAP',TD.GetValue('PDU_NBSALHAP'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_NBSALQ920',TD.GetValue('PDU_NBSALQ920'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_NBSALQ921',TD.GetValue('PDU_NBSALQ921'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_NBSALQ922',TD.GetValue('PDU_NBSALQ922'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_NBSALQ922',TD.GetValue('PDU_NBSALQ922'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_NBSALQ923',TD.GetValue('PDU_NBSALQ923'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_NBSALQ924',TD.GetValue('PDU_NBSALQ924'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_NBSALQ925',TD.GetValue('PDU_NBSALQ925'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_NBSALQ926',TD.GetValue('PDU_NBSALQ926'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_NBSALQ927',TD.GetValue('PDU_NBSALQ927'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_NBSALQ928',TD.GetValue('PDU_NBSALQ928'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_NBSALQ929',TD.GetValue('PDU_NBSALQ929'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_NBSALQ934',TD.GetValue('PDU_NBSALQ934'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_NBSALQ935',TD.GetValue('PDU_NBSALQ935'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_NBSALQ936',TD.GetValue('PDU_NBSALQ936'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_NBSALQ937',TD.GetValue('PDU_NBSALQ937'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_TOTHOMMES',TD.GetValue('PDU_TOTHOMMES'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_TOTFEMMES',TD.GetValue('PDU_TOTFEMMES'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_TOTAPPRENTI',TD.GetValue('PDU_TOTAPPRENTI'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_TOTOUVRIER',TD.GetValue('PDU_TOTOUVRIER'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_TOTETAM',TD.GetValue('PDU_TOTETAM'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_TOTCADRES',TD.GetValue('PDU_TOTCADRES'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_DATEPREVEL',TD.GetValue('PDU_DATEPREVEL'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_NUM',TD.GetValue('PDU_NUM'),FALSE) ;
  TDLFille.AddChampSupValeur('PDU_EMETTSOC',POrg.Pdu_EmettSoc,FALSE) ;
  if (POrg.Pdu_DucsDossier = True) then
    TDLFille.AddChampSupValeur('PDU_DUCSDOSSIER','X',FALSE)
  else
    TDLFille.AddChampSupValeur('PDU_DUCSDOSSIER','-',FALSE);
  TDLFille.AddChampSupValeur('PDD_CODIFICATION',TD.GetValue('PDD_CODIFICATION'),FALSE) ;
  TDLFille.AddChampSupValeur('PDD_CODIFEDITEE',TD.GetValue('PDD_CODIFEDITEE'),FALSE) ;
  TDLFille.AddChampSupValeur('PDD_LIBELLE',TD.GetValue('PDD_LIBELLE'),FALSE) ;
  TDLFille.AddChampSupValeur('PDD_BASECOTISATION',TD.GetValue('PDD_BASECOTISATION'),FALSE) ;
  TDLFille.AddChampSupValeur('PDD_TAUXCOTISATION',TD.GetValue('PDD_TAUXCOTISATION'),FALSE) ;
  TDLFille.AddChampSupValeur('PDD_MTCOTISAT',TD.GetValue('PDD_MTCOTISAT'),FALSE) ;
  TDLFille.AddChampSupValeur('PDD_EFFECTIF',TD.GetValue('PDD_EFFECTIF'),FALSE) ;
  TDLFille.AddChampSupValeur('PDD_TYPECOTISATION',TD.GetValue('PDD_TYPECOTISATION'),FALSE) ;
  TDLFille.AddChampSupValeur('PDD_INSTITUTION',TD.GetValue('PDD_INSTITUTION'),FALSE) ;
  TDLFille.AddChampSupValeur('PDD_LIBELLESUITE',TD.GetValue('PDD_LIBELLESUITE'),FALSE) ;
  TDLFille.AddChampSupValeur('ET_ETABLISSEMENT',POrg.Etab,FALSE) ;
  TDLFille.AddChampSupValeur('ET_LIBELLE',POrg.Et_Libelle,FALSE) ;
  TDLFille.AddChampSupValeur('ET_ADRESSE1',POrg.Et_Adresse1,FALSE) ;
  TDLFille.AddChampSupValeur('ET_ADRESSE2',POrg.Et_Adresse2,FALSE) ;
  TDLFille.AddChampSupValeur('ET_ADRESSE3',POrg.Et_Adresse3,FALSE) ;
  TDLFille.AddChampSupValeur('ET_CODEPOSTAL',POrg.Et_CodePostal,FALSE) ;
  TDLFille.AddChampSupValeur('ET_VILLE',POrg.Et_Ville,FALSE) ;
  TDLFille.AddChampSupValeur('ET_TELEPHONE',POrg.Et_Telephone,FALSE) ;
  TDLFille.AddChampSupValeur('ET_FAX',POrg.Et_Fax,FALSE) ;
  TDLFille.AddChampSupValeur('POG_ETABLISSEMENT',POrg.Etab,FALSE) ;
  TDLFille.AddChampSupValeur('POG_ORGANISME',POrg.Organisme,FALSE) ;
  TDLFille.AddChampSupValeur('POG_LIBELLE',POrg.Pog_Libelle,FALSE) ;
  TDLFille.AddChampSupValeur('POG_NATUREORG',POrg.Pog_NatureOrg,FALSE) ;
  TDLFille.AddChampSupValeur('POG_ADRESSE1',POrg.Pog_Adresse1,FALSE) ;
  TDLFille.AddChampSupValeur('POG_ADRESSE2',POrg.Pog_Adresse2,FALSE) ;
  TDLFille.AddChampSupValeur('POG_ADRESSE3',POrg.Pog_Adresse3,FALSE) ;
  TDLFille.AddChampSupValeur('POG_CODEPOSTAL',POrg.Pog_CodePostal,FALSE) ;
  TDLFille.AddChampSupValeur('POG_PERIODICITDUCS',POrg.Pog_PeriodicitDucs,FALSE) ;
  TDLFille.AddChampSupValeur('POG_AUTREPERIODUCS',POrg.Pog_AutPerioDucs,FALSE) ;
  if (POrg.Pog_PeriodCalcul = True) then
    TDLFille.AddChampSupValeur('POG_PERIODCALCUL','X',FALSE)
  else
    TDLFille.AddChampSupValeur('POG_PERIODCALCUL','-',FALSE);
  if (POrg.Pog_AutPerCalcul = True) then
    TDLFille.AddChampSupValeur('POG_AUTPERCALCUL','X',FALSE)
  else
    TDLFille.AddChampSupValeur('POG_AUTPERCALCUL','-',FALSE);

  TDLFille.AddChampSupValeur('POG_POSTOTAL',POrg.Pog_PosTotal,FALSE) ;
  TDLFille.AddChampSupValeur('POG_LONGTOTAL',POrg.Pog_LongTotal,FALSE) ;
  TDLFille.AddChampSupValeur('POG_VILLE',POrg.Pog_Ville,FALSE) ;
  TDLFille.AddChampSupValeur('POG_LONGEDITABLE',POrg.Pog_LongEditable,FALSE) ;
  TDLFille.AddChampSupValeur('POG_BASETYPARR',POrg.Pog_BaseTypArr,FALSE) ;
  TDLFille.AddChampSupValeur('POG_MTTYPARR',POrg.Pog_MtTypArr,FALSE) ;
  TDLFille.AddChampSupValeur('LETTRE','',FALSE);

  if (Etab <> EtabEnCours) and  (EtabEnCours <> '') then
  begin
    POrg.Pdu_Acompte := POrg.Pdu_Acompte + TD.GetValue('PDU_ACOMPTES');
    POrg.Pdu_Regularisation := POrg.Pdu_Regularisation + TD.GetValue('PDU_REGULARISATION');
    POrg.Pdu_Nbsal_Fpe := POrg.Pdu_Nbsal_Fpe + TD.GetValue('PDU_NBSALFPE');
    POrg.Pdu_Nbsal_Hap := POrg.Pdu_Nbsal_Hap + TD.GetValue('PDU_NBSALHAP');
    POrg.Pdu_Nbsal_Q920 := POrg.Pdu_Nbsal_Q920 + TD.GetValue('PDU_NBSALQ920');
    POrg.Pdu_Nbsal_Q921 := POrg.Pdu_Nbsal_Q921 + TD.GetValue('PDU_NBSALQ921');
    POrg.Pdu_Nbsal_Q922 := POrg.Pdu_Nbsal_Q922 + TD.GetValue('PDU_NBSALQ922');
    POrg.Pdu_Nbsal_Q923 := POrg.Pdu_Nbsal_Q923 + TD.GetValue('PDU_NBSALQ923');
    POrg.Pdu_Nbsal_Q924 := POrg.Pdu_Nbsal_Q924 + TD.GetValue('PDU_NBSALQ924');
    POrg.Pdu_Nbsal_Q925 := POrg.Pdu_Nbsal_Q925 + TD.GetValue('PDU_NBSALQ925');
    POrg.Pdu_Nbsal_Q926 := POrg.Pdu_Nbsal_Q926 + TD.GetValue('PDU_NBSALQ926');
    POrg.Pdu_Nbsal_Q927 := POrg.Pdu_Nbsal_Q927 + TD.GetValue('PDU_NBSALQ927');
    POrg.Pdu_Nbsal_Q928 := POrg.Pdu_Nbsal_Q928 + TD.GetValue('PDU_NBSALQ928');
    POrg.Pdu_Nbsal_Q929 := POrg.Pdu_Nbsal_Q929 + TD.GetValue('PDU_NBSALQ929');
    POrg.Pdu_Nbsal_Q934 := POrg.Pdu_Nbsal_Q934 + TD.GetValue('PDU_NBSALQ934');
    POrg.Pdu_Nbsal_Q935 := POrg.Pdu_Nbsal_Q935 + TD.GetValue('PDU_NBSALQ935');
    POrg.Pdu_Nbsal_Q936 := POrg.Pdu_Nbsal_Q936 + TD.GetValue('PDU_NBSALQ936');
    POrg.Pdu_Nbsal_Q937 := POrg.Pdu_Nbsal_Q937 + TD.GetValue('PDU_NBSALQ937');
    Porg.Pdu_TotHommes := Porg.Pdu_TotHommes + TD.GetValue('PDU_TOTHOMMES');
    POrg.Pdu_TotFemmes := Porg.Pdu_TotFemmes + TD.GetValue('PDU_TOTFEMMES');
    POrg.Pdu_TotApprenti := POrg.Pdu_TotApprenti + TD.GetValue('PDU_TOTAPPRENTI');
    POrg.Pdu_TotCadres := POrg.Pdu_TotCadres + TD.GetValue('PDU_TOTCADRES');
    POrg.Pdu_TotEtam := POrg.Pdu_TotEtam + TD.GetValue('PDU_TOTETAM');
    POrg.Pdu_TotOuvrier := POrg.Pdu_TotOuvrier + TD.GetValue('PDU_TOTOUVRIER');
  end;
  EtabEnCours := etab;
  if (TDLFille.GetValue('PDD_CODIFICATION') <> NULL) then
     Neant := False
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 26/03/2003
Modifié le ... :   /  /
Description .. : Préparation de la TOB TOBEtat utilisée pour l'édition.
Suite ........ :
Suite ........ : Différents cas à traiter selon le type de paramètrage.
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.MajTobEtat (var TDL : TOB; var TOBEtat : TOB; POrg : TParamOrg; NumOrdre : integer);
var
  i, CasTraitement                     : integer;
  TDLFille                             : TOB;
begin
  CasTraitement := 0;

  {élimination des éléments provenat de ducs Néant}
  TDLFille := TDL.FindFirst([''],[''], True);
  for I := 1 to TDL.Detail.Count do
  begin
  if (TDLFille.GetValue('PDD_CODIFICATION') = NULL) then
    TDLFille.free;
  TDLFille := TDL.FindNext([''],[''], True);
  end;

  if (TOBEtat = NIL) then
  begin
    TOBEtat := TOB.Create('Ducs à éditer', NIL, -1);
  end;

  if (POrg.Pog_SousTotDucs = True) then
  {Ducs avec lignes de sous-total}
  begin
    if (POrg.Pog_LongTotal = 0) and (POrg.Pog_PosTotal = 0) then
    {Cas correspondant à une DUCS IRC standardisée}
    begin
      TDL.Detail.Sort('PDD_INSTITUTION;PDD_CODIFICATION;PDD_TAUXCOTISATION;PDU_NUM');
      CasTraitement := 1;
    end;
    if (POrg.Pog_LongTotal <> 0) and (POrg.Pog_PosTotal <> 0) then
    {Cas correspondant à une DUCS IRC personnalisée ou MSA }
    begin
      TDL.Detail.Sort('PDD_CODIFICATION;PDD_TAUXCOTISATION;PDU_NUM');
      CasTraitement := 2;
    end;

  end
  else
  {Ducs sans ligne de sous-total}
  {Cas DUCS URSSAF              }
  begin
      TDL.Detail.Sort('PDU_NUM;PDD_CODIFICATION;PDD_TAUXCOTISATION');
      CasTraitement := 3;
  end;

  ConstructionTOBEtat(TDL, TOBEtat,CasTraitement, NumOrdre, POrg);

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 03/04/2003
Modifié le ... :   /  /
Description .. : Construction de la TOBEtat
Mots clefs ... : PAIE PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.ConstructionTOBEtat(TDL : TOB; var TOBEtat : TOB; CasTraitement, NumOrdre : integer; POrg : TParamOrg);
var
  CodificationLue, CodificationEnCours, TauxLu, TauxEnCours     : String;
  InstitutionLue, InstitutionEnCours, LettreLue                 : String;
  TOBEtatFille,  TDLFille                                       : TOB;
  MontST                                                        : double;
  I                                                             : integer;
begin
  if (casTraitement = 3) then
  {Cas de type URSSAF - Traitement TAux AT ou autre sans st}
  begin
    CreationTOBEtatAT(TDL, TOBEtat, MontSt);
    TDL.Detail.Sort('PDD_CODIFICATION;PDD_TAUXCOTISATION;PDU_NUM');
  end;

  {Cas général}
  TDLFille := TDL.FindFirst([''],[''], True);
  if (TDLFille <> NIL) and
     (TDLFille.GetValue('PDD_CODIFICATION') <> NULL) then
  begin
    CodificationLue := TDLFille.GetValue ('PDD_CODIFICATION');
    TauxLu := TDLFille.GetValue ('PDD_TAUXCOTISATION');
    InstitutionLue := TDLFille.GetValue ('PDD_INSTITUTION');
    LettreLue := TDLFille.GetValue('LETTRE');
  end;

  for I := 1 to TDL.Detail.Count do
  begin
//    if (not((CasTraitement = 2) and (TDLFille.GetValue('PDD_TYPECOTISATION') = 'S'))) and
    if ( TDLFille.GetValue ('PDD_CODIFICATION') <> Null) then
    begin

      if (((CasTraitement = 1) and
              (Codificationlue = CodificationEnCours) and
              (TauxLu = TauxEnCours) and
              (Institutionlue = InstitutionEnCours)) or
             ((CasTraitement = 2) and
              (Codificationlue = CodificationEnCours) and
              (TauxLu = TauxEnCours)) or
             ((CasTraitement = 3) and
              (Codificationlue = CodificationEnCours) and
              (TauxLu = TauxEnCours) and
              (LettreLue = ''))) then
      begin
        MajTobEtatFille (TDLFille, TOBEtat, TOBEtatFille, MontSt, CasTraitement);
      end;
      if (((CasTraitement = 1) and
           (((Codificationlue = CodificationEnCours) and
             (TauxLu <> TauxEnCours) and
             (Institutionlue = InstitutionEnCours)) or
             ((Codificationlue <> CodificationEnCours) and
              (TDLFille.GetValue('PDD_TYPECOTISATION') <> 'S')))) or
           ((CasTraitement = 2) and
           (((Codificationlue = CodificationEnCours) and
             (TauxLu <> TauxEnCours)) or
             (Codificationlue <> CodificationEnCours))) or
           ((CasTraitement = 3) and
           (((Codificationlue = CodificationEnCours) and
             (TauxLu <> TauxEnCours)) or
            (Codificationlue <> CodificationEnCours)) and
            (LettreLue = ''))) then
      begin
        if (CasTraitement = 2) and
           (TDLFille.GetValue('PDD_TYPECOTISATION') = 'S') then
        begin
          CreationTOBEtatST(CasTraitement,TDLFille,TOBEtat,TOBEtatFille,CodificationLue,CodificationEnCours,InstitutionEnCours,POrg,Numordre, MontST);
        end;

        NumOrdre := NumOrdre + 1;
//      TOBEtatFille.PutValue('PDD_NUMORDRE', NumOrdre);

        CreationTOBFille (TDLFille, TOBEtat, TOBEtatFille, MontSt, False);

        CodificationEnCours := TDLFille.GetValue ('PDD_CODIFICATION');
        TauxEnCours := TDLFille.GetValue ('PDD_TAUXCOTISATION');
        InstitutionEnCours := TDLFille.GetValue ('PDD_INSTITUTION');

      end;

      if (CasTraitement = 1) and  (TDLFille.GetValue('PDD_TYPECOTISATION') = 'S') then
      begin
        CreationTOBEtatST(CasTraitement,TDLFille,TOBEtat,TOBEtatFille,CodificationLue,CodificationEnCours,InstitutionEnCours,POrg,Numordre,  MontST);
      end;
    end;

    TDLFille := TDL.FindNext([''],[''], TRUE);
    if (TDLFille <> NIL) and
       (TDLFille.GetValue('PDD_CODIFICATION') <> NULL) then
    begin
      CodificationLue := TDLFille.GetValue ('PDD_CODIFICATION');
      TauxLu := TDLFille.GetValue ('PDD_TAUXCOTISATION');
      InstitutionLue := TDLFille.GetValue ('PDD_INSTITUTION');
      LettreLue := TDLFille.GetValue('LETTRE');
    end;
  end;
  TOBEtatFille := TOBEtat.FindFirst(['PDU_ETABLISSEMENT','PDU_ORGANISME'],
                                      [Porg.Etab,Porg.Organisme], True);
  TOBEtatFille.PutValue ('PDU_ACOMPTES', Porg.Pdu_Acompte) ;
  TOBEtatFille.PutValue ('PDU_REGULARISATION',POrg.Pdu_Regularisation) ;
  TOBEtatFille.PutValue ('PDU_NBSALFPE',POrg.Pdu_Nbsal_Fpe) ;
  TOBEtatFille.PutValue ('PDU_NBSALHAP',POrg.Pdu_Nbsal_Hap) ;
  TOBEtatFille.PutValue ('PDU_NBSALQ920',POrg.Pdu_Nbsal_Q920) ;
  TOBEtatFille.PutValue ('PDU_NBSALQ921',POrg.Pdu_Nbsal_Q921) ;
  TOBEtatFille.PutValue ('PDU_NBSALQ922',POrg.Pdu_Nbsal_Q922) ;
  TOBEtatFille.PutValue ('PDU_NBSALQ923',POrg.Pdu_Nbsal_Q923) ;
  TOBEtatFille.PutValue ('PDU_NBSALQ924',POrg.Pdu_Nbsal_Q924) ;
  TOBEtatFille.PutValue ('PDU_NBSALQ925',POrg.Pdu_Nbsal_Q925) ;
  TOBEtatFille.PutValue ('PDU_NBSALQ926',POrg.Pdu_Nbsal_Q926) ;
  TOBEtatFille.PutValue ('PDU_NBSALQ927',POrg.Pdu_Nbsal_Q927) ;
  TOBEtatFille.PutValue ('PDU_NBSALQ928',POrg.Pdu_Nbsal_Q928) ;
  TOBEtatFille.PutValue ('PDU_NBSALQ929',POrg.Pdu_Nbsal_Q929) ;
  TOBEtatFille.PutValue ('PDU_NBSALQ934',POrg.Pdu_Nbsal_Q934) ;
  TOBEtatFille.PutValue ('PDU_NBSALQ935',POrg.Pdu_Nbsal_Q935) ;
  TOBEtatFille.PutValue ('PDU_NBSALQ936',POrg.Pdu_Nbsal_Q936) ;
  TOBEtatFille.PutValue ('PDU_NBSALQ937',POrg.Pdu_Nbsal_Q937) ;
  TOBEtatFille.PutValue ('PDU_TOTHOMMES',POrg.Pdu_TotHommes) ;
  TOBEtatFille.PutValue ('PDU_TOTFEMMES',POrg.Pdu_TotFemmes) ;
  TOBEtatFille.PutValue ('PDU_TOTAPPRENTI',POrg.Pdu_TotApprenti) ;
  TOBEtatFille.PutValue ('PDU_TOTOUVRIER',POrg.Pdu_TotOuvrier) ;
  TOBEtatFille.PutValue ('PDU_TOTETAM',POrg.Pdu_TotEtam) ;
  TOBEtatFille.PutValue ('PDU_TOTCADRES',POrg.Pdu_TotCadres) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 10/04/2003
Modifié le ... :   /  /    
Description .. : Création d'une fille de la TOB finale utilisée pour l'édition.
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.CreationTOBFille (TOrig : TOB; var TOBDest, TOBDestFille : TOB; var MontST : double; Neant : boolean);
        begin

  TOBDestFille := TOB.Create('',TOBDest,-1) ;
  TOBDestFille.AddChampSupValeur('PDU_ETABLISSEMENT',TOrig.GetValue('PDU_ETABLISSEMENT'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_ORGANISME',TOrig.GetValue('PDU_ORGANISME'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_DATEDEBUT',TOrig.GetValue('PDU_DATEDEBUT') ,FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_DATEFIN',TOrig.GetValue('PDU_DATEFIN'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_SIRET',TOrig.GetValue('PDU_SIRET'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_APE',TOrig.GetValue('PDU_APE'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_GROUPE',TOrig.GetValue('PDU_GROUPE'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_NUMERO',TOrig.GetValue('PDU_NUMERO'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_ABREGEPERIODE',TOrig.GetValue('PDU_ABREGEPERIODE'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_DATEEXIGIBLE',TOrig.GetValue('PDU_DATEEXIGIBLE'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_DATELIMDEPOT',TOrig.GetValue('PDU_DATELIMDEPOT'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_DATEREGLEMENT',TOrig.GetValue('PDU_DATEREGLEMENT'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_PAIEMENT',TOrig.GetValue('PDU_PAIEMENT'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_CESSATION',TOrig.GetValue('PDU_CESSATION'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_CONTINUATION',TOrig.GetValue('PDU_CONTINUATION'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_SUSPENSION',TOrig.GetValue('PDU_SUSPENSION'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_MAINTIENT',TOrig.GetValue('PDU_MAINTIENT'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_DECLARANT',TOrig.GetValue('PDU_DECLARANT'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_DECLARANTSUITE',TOrig.GetValue('PDU_DECLARANTSUITE'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_TELEPHONEDECL',TOrig.GetValue('PDU_TELEPHONEDECL'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_FAXDECLARANT',TOrig.GetValue('PDU_FAXDECLARANT'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_LIGNEOPTIQUE',TOrig.GetValue('PDU_LIGNEOPTIQUE'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_ACOMPTES',TOrig.GetValue('PDU_ACOMPTES'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_REGULARISATION',TOrig.GetValue('PDU_REGULARISATION'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_MONNAIETENUE',TOrig.GetValue('PDU_MONNAIETENUE'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_NBSALFPE',TOrig.GetValue('PDU_NBSALFPE'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_NBSALHAP',TOrig.GetValue('PDU_NBSALHAP'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_NBSALQ920',TOrig.GetValue('PDU_NBSALQ920'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_NBSALQ921',TOrig.GetValue('PDU_NBSALQ921'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_NBSALQ922',TOrig.GetValue('PDU_NBSALQ922'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_NBSALQ922',TOrig.GetValue('PDU_NBSALQ922'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_NBSALQ923',TOrig.GetValue('PDU_NBSALQ923'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_NBSALQ924',TOrig.GetValue('PDU_NBSALQ924'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_NBSALQ925',TOrig.GetValue('PDU_NBSALQ925'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_NBSALQ926',TOrig.GetValue('PDU_NBSALQ926'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_NBSALQ927',TOrig.GetValue('PDU_NBSALQ927'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_NBSALQ928',TOrig.GetValue('PDU_NBSALQ928'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_NBSALQ929',TOrig.GetValue('PDU_NBSALQ929'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_NBSALQ934',TOrig.GetValue('PDU_NBSALQ934'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_NBSALQ935',TOrig.GetValue('PDU_NBSALQ935'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_NBSALQ936',TOrig.GetValue('PDU_NBSALQ936'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_NBSALQ937',TOrig.GetValue('PDU_NBSALQ937'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_TOTHOMMES',TOrig.GetValue('PDU_TOTHOMMES'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_TOTFEMMES',TOrig.GetValue('PDU_TOTFEMMES'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_TOTAPPRENTI',TOrig.GetValue('PDU_TOTAPPRENTI'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_TOTOUVRIER',TOrig.GetValue('PDU_TOTOUVRIER'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_TOTETAM',TOrig.GetValue('PDU_TOTETAM'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_TOTCADRES',TOrig.GetValue('PDU_TOTCADRES'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_DATEPREVEL',TOrig.GetValue('PDU_DATEPREVEL'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_NUM',0,FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_EMETTSOC',TOrig.GetValue('PDU_EMETTSOC'),FALSE) ;
  TOBDestFille.AddChampSupValeur('PDU_DUCSDOSSIER',TOrig.GetValue('PDU_DUCSDOSSIER'),FALSE) ;
  if (Neant = False) then
  begin
    TOBDestFille.AddChampSupValeur('PDD_CODIFICATION',TOrig.GetValue('PDD_CODIFICATION'),FALSE) ;
    TOBDestFille.AddChampSupValeur('PDD_CODIFEDITEE',TOrig.GetValue('PDD_CODIFEDITEE'),FALSE) ;
    TOBDestFille.AddChampSupValeur('PDD_LIBELLE',TOrig.GetValue('PDD_LIBELLE'),FALSE) ;
    TOBDestFille.AddChampSupValeur('PDD_BASECOTISATION',TOrig.GetValue('PDD_BASECOTISATION'),FALSE) ;
    TOBDestFille.AddChampSupValeur('PDD_TAUXCOTISATION',TOrig.GetValue('PDD_TAUXCOTISATION'),FALSE) ;
    TOBDestFille.AddChampSupValeur('PDD_MTCOTISAT',TOrig.GetValue('PDD_MTCOTISAT'),FALSE) ;
    TOBDestFille.AddChampSupValeur('PDD_EFFECTIF',TOrig.GetValue('PDD_EFFECTIF'),FALSE) ;
    TOBDestFille.AddChampSupValeur('PDD_TYPECOTISATION',TOrig.GetValue('PDD_TYPECOTISATION'),FALSE) ;
    TOBDestFille.AddChampSupValeur('PDD_INSTITUTION',TOrig.GetValue('PDD_INSTITUTION'),FALSE) ;
    TOBDestFille.AddChampSupValeur('PDD_LIBELLESUITE',TOrig.GetValue('PDD_LIBELLESUITE'),FALSE) ;
  end;
  TOBDestFille.AddChampSupValeur('ET_ETABLISSEMENT',TOrig.GetValue('ET_ETABLISSEMENT'),FALSE) ;
  TOBDestFille.AddChampSupValeur('ET_LIBELLE',TOrig.GetValue('ET_LIBELLE'),FALSE) ;
  TOBDestFille.AddChampSupValeur('ET_ADRESSE1',TOrig.GetValue('ET_ADRESSE1'),FALSE) ;
  TOBDestFille.AddChampSupValeur('ET_ADRESSE2',TOrig.GetValue('ET_ADRESSE2'),FALSE) ;
  TOBDestFille.AddChampSupValeur('ET_ADRESSE3',TOrig.GetValue('ET_ADRESSE3'),FALSE) ;
  TOBDestFille.AddChampSupValeur('ET_CODEPOSTAL',TOrig.GetValue('ET_CODEPOSTAL'),FALSE) ;
  TOBDestFille.AddChampSupValeur('ET_VILLE',TOrig.GetValue('ET_VILLE'),FALSE) ;
  TOBDestFille.AddChampSupValeur('ET_TELEPHONE',TOrig.GetValue('ET_TELEPHONE'),FALSE) ;
  TOBDestFille.AddChampSupValeur('ET_FAX',TOrig.GetValue('ET_FAX'),FALSE) ;
  TOBDestFille.AddChampSupValeur('POG_ETABLISSEMENT',TOrig.GetValue('POG_ETABLISSEMENT'),FALSE) ;
  TOBDestFille.AddChampSupValeur('POG_ORGANISME',TOrig.GetValue('POG_ORGANISME'),FALSE) ;
  TOBDestFille.AddChampSupValeur('POG_LIBELLE',TOrig.GetValue('POG_LIBELLE'),FALSE) ;
  TOBDestFille.AddChampSupValeur('POG_NATUREORG',TOrig.GetValue('POG_NATUREORG'),FALSE) ;
  TOBDestFille.AddChampSupValeur('POG_ADRESSE1',TOrig.GetValue('POG_ADRESSE1'),FALSE) ;
  TOBDestFille.AddChampSupValeur('POG_ADRESSE2',TOrig.GetValue('POG_ADRESSE2'),FALSE) ;
  TOBDestFille.AddChampSupValeur('POG_ADRESSE3',TOrig.GetValue('POG_ADRESSE3'),FALSE) ;
  TOBDestFille.AddChampSupValeur('POG_CODEPOSTAL',TOrig.GetValue('POG_CODEPOSTAL'),FALSE) ;
  TOBDestFille.AddChampSupValeur('POG_PERIODICITDUCS',TOrig.GetValue('POG_PERIODICITDUCS'),FALSE) ;
  TOBDestFille.AddChampSupValeur('POG_AUTREPERIODUCS',TOrig.GetValue('POG_AUTREPERIODUCS'),FALSE) ;
  TOBDestFille.AddChampSupValeur('POG_PERIODCALCUL',TOrig.GetValue('POG_PERIODCALCUL'),FALSE) ;
  TOBDestFille.AddChampSupValeur('POG_AUTPERCALCUL',TOrig.GetValue('POG_AUTPERCALCUL'),FALSE) ;
  TOBDestFille.AddChampSupValeur('POG_POSTOTAL',TOrig.GetValue('POG_POSTOTAL'),FALSE) ;
  TOBDestFille.AddChampSupValeur('POG_LONGTOTAL',TOrig.GetValue('POG_LONGTOTAL'),FALSE) ;
  TOBDestFille.AddChampSupValeur('POG_VILLE',TOrig.GetValue('POG_VILLE'),FALSE) ;
  TOBDestFille.AddChampSupValeur('POG_LONGEDITABLE',TOrig.GetValue('POG_LONGEDITABLE'),FALSE) ;
  TOBDestFille.AddChampSupValeur('POG_BASETYPARR',TOrig.GetValue('POG_BASETYPARR'),FALSE) ;
  TOBDestFille.AddChampSupValeur('POG_MTTYPARR',TOrig.GetValue('POG_MTTYPARR'),FALSE) ;
  if (Neant = False) then
  begin
    TOBDestFille.AddChampSupValeur('LETTRE',TOrig.GetValue('LETTRE'),FALSE);
    if (TOBDestFille.GetValue('PDD_TYPECOTISATION') <> 'S') then
      MontSt := MontST +  TOBDestFille.GetValue('PDD_MTCOTISAT');
  end;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 10/04/2003
Modifié le ... :   /  /    
Description .. : Mise à jour d'une fille de la TOB finale utilisée pour l'édition.
Suite ........ : Tenue des cumuls BAse , Effectif, Montant
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.MajTobEtatFille (TDLFille : TOB; var TOBEtat, TOBEtatFille : TOB; var MontSt : double; CasTraitement : integer);
begin
  TOBEtatFille.PutValue('PDD_BASECOTISATION',
                        TOBEtatFille.GetValue('PDD_BASECOTISATION')+
                        TDLFille.GetValue('PDD_BASECOTISATION')) ;
  TOBEtatFille.PutValue('PDD_EFFECTIF',
                        TOBEtatFille.GetValue('PDD_EFFECTIF')+
                        TDLFille.GetValue('PDD_EFFECTIF')) ;
  TOBEtatFille.PutValue('PDD_MTCOTISAT',
                        TOBEtatFille.GetValue('PDD_MTCOTISAT')+
                        TDLFille.GetValue('PDD_MTCOTISAT')) ;
  MontSt := MontST +  TDLFille.GetValue('PDD_MTCOTISAT');

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 10/04/2003
Modifié le ... :   /  /
Description .. : traitement des lignes de Sous total
Suite ........ : formatage de la ligne en fonction du cas traité
Suite ........ : et création d'une fille de la TOB finale
Mots clefs ... : PAIE PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.CreationTOBEtatST(CasTraitement : integer;TDLFille : TOB;var TOBEtat, TOBEtatFille : TOB; CodificationLue,CodificationEnCours,InstitutionEnCours : String; POrg : TParamOrg; var Numordre : integer;var MontST : double);
var
  CodifST, Predef, NoInst, LibInst           : String;
  QDucsParam, QInst                          : TQuery;
  PosDeb, LongEdit,NoInstEnCours             : integer;
  NomChamp                                   : array[1..2] of string;
  ValeurChamp                                : array[1..2] of variant;
  TobSTExiste                                : TOB;

begin
  if (CasTraitement = 2) then
  {Traitement ST quand Longueur et position de ST renseignée}
  {cas IRC personnalisée, MSA, BTP ou autre                 }
  begin
    if (Copy(CodificationLue,POrg.Pog_PosTotal,POrg.Pog_LongTotal) <>
        Copy(CodificationEnCours,POrg.Pog_PosTotal,POrg.Pog_LongTotal)) then
      {Rupture pour sous totalisation}
    begin
      if (POrg.Pog_NatureOrg = '200') then
        {ASSEDIC}
        CodifST := Copy(CodificationEnCours,1,5)+'ZZ';
      if (POrg.Pog_NatureOrg = '600') then
        {MSA}
        CodifST := Copy(CodificationEnCours,1,4)+'ZZZ';
      if (POrg.Pog_NatureOrg = '300') then
        {IRC personnalisée}
        if (POrg.Pog_PosTotal =  1) then
          CodifST := Copy(CodificationEnCours,1,POrg.Pog_LongTotal) +
                     Copy('ZZZZZZZ',1,(7-POrg.Pog_LongTotal))
        else
          CodifST := Copy(CodificationEnCours,1,POrg.Pog_PosTotal-1) +
                     Copy(CodificationEnCours,POrg.Pog_PosTotal,POrg.Pog_PosTotal)+
                     Copy('ZZZZZZZ',1,(7-POrg.Pog_PosTotal));

      CreationTOBFille (TDLFille, TOBEtat, TOBEtatFille, MontST, False);

      QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, '+
                            'PDP_LIBELLESUITE,'+
                            'PDP_TYPECOTISATION'+
                            ' FROM DUCSPARAM '+
                            ' WHERE ##PDP_PREDEFINI## '+
                            'PDP_CODIFICATION = "'+CodifST+'"'+
                            ' ORDER BY PDP_PREDEFINI',True);
      Predef := '';
      while  not QDucsParam.eof do
      begin
        Predef :=  QDucsparam.FindField('PDP_PREDEFINI').AsString;
        if (Predef = 'DOS') then break;
          QDucsParam.Next;
      end;
      if (Predef <> '') then
      begin
        TOBEtatFille.PutValue ('PDD_CODIFICATION',CodifST);
        TOBEtatFille.PutValue ('PDD_TYPECOTISATION',
                      QDucsParam.FindField('PDP_TYPECOTISATION').AsString);
        TOBEtatFille.PutValue ('PDD_LIBELLE',
                      QDucsParam.FindField('PDP_LIBELLE').AsString);
        TOBEtatFille.PutValue ('PDD_LIBELLESUITE',
                      QDucsParam.FindField('PDP_LIBELLESUITE').AsString);
        Ferme(QDucsParam);
        TOBEtatFille.PutValue ('PDD_EFFECTIF',0);
        TOBEtatFille.PutValue ('PDD_BASECOTISATION',0.00);
        TOBEtatFille.PutValue ('PDD_TAUXCOTISATION',0.00);
        TOBEtatFille.PutValue ('PDD_MTCOTISAT',Arrondi(MontST,5));
        TOBEtatFille.PutValue ('PDD_COMURBAINE','    ');
        if (POrg.Pog_NatureOrg = '200') then
          { ASSEDIC }
          TOBEtatFille.PutValue ('PDD_CODIFEDITEE',
                              copy(CodifST,5,1)+' ');
        if (POrg.Pog_NatureOrg ='600') then
          { MSA }
          TOBEtatFille.PutValue ('PDD_CODIFEDITEE',
                              copy(CodifST,3,2));
        if (POrg.Pog_NatureOrg='300') then
        begin
          {IRC personnalisée}
          if (POrg.Pog_PosTotal < 3) then
          begin
            PosDeb := 3;
            LongEdit := POrg.Pog_LongTotal-2;
          end
          else
          begin
            PosDeb := POrg.Pog_PosTotal;
            LongEdit := POrg.Pog_LongTotal;
          end;
          TOBEtatFille.PutValue ('PDD_CODIFEDITEE',
                              copy(CodifST, PosDeb, LongEdit));
          TOBEtatFille.PutValue ('PDD_INSTITUTION',
                              InstitutionEnCours);
        end;
        NumOrdre := NumOrdre + 1;
        TOBEtatFille.PutValue('PDD_NUMORDRE', NumOrdre);
      end;
    end;
    MontST := 0.0;
  end;
  if (CasTraitement = 1) then
  {Traitement ST quand Longueur et position de ST = 0}
  {cas IRC standardisée                              }
  begin
    NoInstEnCours := StrToInt(copy(CodificationLue,1,2));
    NomChamp[1] := 'PDD_CODIFICATION';
    ValeurChamp[1] := IntToStr(NoInstEnCours)+'ZZZZZ';
    NomChamp[2] := 'PDD_INSTITUTION';
    ValeurChamp[2] := InstitutionEnCours;

        TobSTExiste := TOBEtat.FindFirst(NomChamp,ValeurChamp, TRUE);
    if (TobStExiste = NIL) then
    { Ligne St existe déjà
      qd on traite une Ducs Dossier on peut avoir des codifications
      concernant le même institution sur déclarations (2 etab) différentes,
      il ne faut générer qu'un seul Sous Total}
    begin
      NoInstEnCours := StrToInt(copy(CodificationEnCours,1,2));
      NoInst := CodificationEnCours;

      NomChamp[1] := 'PDD_CODIFICATION';
      ValeurChamp[1] := IntToStr(NoInstEnCours)+'ZZZZZ';
      NomChamp[2] := '';
      ValeurChamp[2] := '';

      TobSTExiste := TOBEtat.FindFirst(NomChamp,ValeurChamp, TRUE);
      while (TobStExiste <> NIL) do
      begin
        NoInstEnCours := NoInstEnCours + 1;
        ValeurChamp[1] := IntToStr(NoInstEnCours)+'ZZZZZ';
        TobSTExiste:=TOBEtat.FindFirst(NomChamp,ValeurChamp, TRUE);
      end;

      CodifST := IntToStr(NoInstEnCours)+'ZZZZZ';
      CreationTOBFille (TDLFille, TOBEtat, TOBEtatFille, MontST, False);

      TOBEtatFille.PutValue ('PDD_CODIFICATION',CodifST);
      TOBEtatFille.PutValue ('PDD_TYPECOTISATION','S');

      {recup Code Institution et Libellé}
      QInst := OpenSql('SELECT PIP_ABREGE '+
                       ' FROM INSTITUTIONPAYE '+
                       ' WHERE PIP_INSTITUTION = "'+
                        NoInst+'"',True);

      if not QInst.eof then
      begin
       LibInst := 'ST '+NoInst+' '+
                  Copy(QInst.FindField('PIP_ABREGE').AsString,1,7);
       TOBEtatFille.PutValue ('PDD_LIBELLE',LibInst);
       TOBEtatFille.PutValue ('PDD_INSTITUTION', NoInst);
      end;
      Ferme(QInst);
      TOBEtatFille.PutValue ('PDD_MTCOTISAT',Arrondi(MontST,5));
      MontST := 0.0;
    end;
  end;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 10/04/2003
Modifié le ... :   /  /    
Description .. : Traitement des lignes associées à un taux AT pour l'Urssaf
Suite ........ : et création des filles de la TOB Finale
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.CreationTOBEtatAT(TDL : TOB;var TOBEtat : TOB; var MontST : double);
var
  TDLFille,TOBA,TOBAFille,TOBD,TOBDFille,TOBP,TOBPFille,TOBEtatFille : TOB;
  CodifLue, CodifEnCours,CodifAt, LettreEnCours                      : String;
  TxA,TxD,TxP,EffectA,BaseA,MtA,EffectD,BaseD,MtD,EffectP,BaseP,MtP  : double;
  I,Lettre                                                           : integer;
Const
  Alphabet : array[1..26] of Char = ('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');
begin
  EffectA := 0.00;
  BaseA := 0.00;
  MtA := 0.00;
  EffectD := 0.00;
  BaseD := 0.00;
  MtD := 0.00;
  EffectP := 0.00;
  BaseP := 0.00;
  MtP := 0.00;

  Lettre := 0;

  TDLFille := TDL.FindFirst([''],[''], True);
  if (TDLFille <> NIL) and
     (TDLFille.GetValue('PDD_CODIFICATION') <> NULL) then
  begin
    CodifLue := TDLFille.GetValue('PDD_CODIFICATION');
    CodifEnCours := TDLFille.GetValue('PDD_CODIFICATION');
  end;

  for I := 1 To TDL.Detail.Count-1 do
  begin
    if (TDLFille.GetValue('PDD_CODIFICATION') <> NULL) then
    begin
      if (Copy(CodifLue,1,6) <> copy(CodifEnCours,1,6)) then
      begin
        CodifEnCours := TDLFille.GetValue('PDD_CODIFICATION');
      end;
      if (TDLFille.GetValue('PDD_TYPECOTISATION') = 'A') then
      begin
        if (TOBA = NIL) then
          TOBA := TOB.Create('Les TAux AT', NIL, -1);

        Lettre := Lettre+1;

        CreationTOBFille(TDLFille,TOBA ,TOBAFille, MontST, False);
        TOBAFille.PutValue('LETTRE',Alphabet[Lettre]);
        TDLFille.PutValue('LETTRE',Alphabet[Lettre]);
        CodifAt := TDLFille.GetValue('PDD_CODIFICATION');
      end
      else
      begin
        if (Copy(CodifLue,1,6) = Copy(CodifEnCours,1,6)) and
           (Copy(CodifAT,1,6) = Copy(CodifEnCours,1,6)) then
        begin
          if (Copy(CodifLue,7,1) = 'D') then
          {déplafonné}
          begin
            if (TOBD = NIL) then
              TOBD := TOB.Create('Les Déplafonnés', NIL, -1);

            CreationTOBFille(TDLFille,TOBD,TOBDFille, MontST, False);
            TOBDFille.PutValue('LETTRE',Alphabet[Lettre]);
            TDLFille.PutValue('LETTRE',Alphabet[Lettre]);
          end;

          if (Copy(CodifLue,7,1) = 'P') then
          {plafonné}
          begin
            if (TOBP = NIL) then
              TOBP := TOB.Create('Les Déplafonnés', NIL, -1);

            CreationTOBFille(TDLFille,TOBP,TOBPFille, MontST, False);
            TOBPFille.PutValue('LETTRE',Alphabet[Lettre]);
            TDLFille.PutValue('LETTRE',Alphabet[Lettre]);
          end;
        end;
      end;
    end;
    TDLFille := TDL.Findnext([''],[''], True);
    if (TDLFille.GetValue('PDD_CODIFICATION') <> NULL) then
      CodifLue := TDLFille.GetValue('PDD_CODIFICATION');
  end;

  TxA := 0.00;
  TxD := 0.00;
  TxP := 0.00;

  if (TOBAFille <> NIL) then
    TOBAFille := TOBA.FindFirst([''],[''],True);
  if (TOBDFille <> NIL) then
    TOBDFille := TOBD.FindFirst([''],[''],True);
  if (TOBPFille <> NIL) then
    TOBPFille := TOBP.FindFirst([''],[''],True);

  if (TOBAFille <> NIL) then
    LettreEnCours := TOBAFille.GetValue('LETTRE');
  if (TOBDFille <> NIL) then
  begin
    if (LettreEnCours > TOBDFille.GetValue('LETTRE')) then
      LettreEnCours := TOBDFille.GetValue('LETTRE');
  end;
  if (TOBPFille <> NIL) then
  begin
    if (LettreEnCours > TOBPFille.GetValue('LETTRE')) then
      LettreEnCours := TOBPFille.GetValue('LETTRE');
  end;

  for I := 1 to TOBA.Detail.Count do
  begin
    if ((TOBAFille <> NIL) and
        (((TxA = TOBAFille.GetValue('PDD_TAUXCOTISATION')) and
         (LettreEnCours = TOBAFille.GetValue('LETTRE'))) or
         (LettreEnCours <> TOBAFille.GetValue('LETTRE')))) and
       ((TOBDFille <> NIL) and
        (((TxD = TOBDFille.GetValue('PDD_TAUXCOTISATION')) and
         (LettreEnCours = TOBDFille.GetValue('LETTRE'))) or
         (LettreEnCours <> TOBDFille.GetValue('LETTRE')))) and
       ((TOBPFille <> NIL) and
        (((TxP = TOBPFille.GetValue('PDD_TAUXCOTISATION')) and
         (LettreEnCours = TOBPFille.GetValue('LETTRE'))) or
         (LettreEnCours <> TOBPFille.GetValue('LETTRE')))) then
    begin
      if (LettreEnCours = TOBAFille.GetValue('LETTRE')) then
      begin
        EffectA := EffectA + TOBAFille.GetValue('PDD_EFFECTIF');
        BaseA := BaseA + TOBAFille.GetValue('PDD_BASECOTISATION');
        MtA := MtA + TOBAFille.GetValue('PDD_MTCOTISAT');
      end;

      if (LettreEnCours = TOBDFille.GetValue('LETTRE')) then
      begin
        EffectD := EffectD + TOBDFille.GetValue('PDD_EFFECTIF');
        BaseD := BaseD + TOBDFille.GetValue('PDD_BASECOTISATION');
        MtD := MtD + TOBDFille.GetValue('PDD_MTCOTISAT');
      end;

      if (LettreEnCours = TOBPFille.GetValue('LETTRE')) then
      begin
        EffectP := EffectP + TOBPFille.GetValue('PDD_EFFECTIF');
        BaseP := BaseP + TOBPFille.GetValue('PDD_BASECOTISATION');
        MtP := MtP + TOBPFille.GetValue('PDD_MTCOTISAT');
      end;

      if (LettreEnCours = TOBAFille.GetValue('LETTRE')) then TOBAFille.free;
      if (LettreEnCours = TOBDFille.GetValue('LETTRE')) then TOBDFille.free;
      if (LettreEnCours = TOBPFille.GetValue('LETTRE')) then TOBPFille.free;
    end
    else
    begin
      TxA := 0;
      TxD := 0;
      TxP := 0;
      if (TOBAFille <> NIL) and (LettreEnCours = TOBAFille.GetValue('LETTRE'))then
        TxA := TOBAFille.GetValue('PDD_TAUXCOTISATION');
      if (TOBDFille <> NIL) and (LettreEnCours = TOBDFille.GetValue('LETTRE')) then
        TxD := TOBDFille.GetValue('PDD_TAUXCOTISATION');
      if (TOBPFille <> NIL) and (LettreEnCours = TOBPFille.GetValue('LETTRE'))then     
        TxP := TOBPFille.GetValue('PDD_TAUXCOTISATION');
      if (TxA <> 0) or (TxD <> 0) or (TxP <> 0) then
      begin
        if ((TOBAFille <> NIL) and
           (LettreEnCours = TOBAFille.GetValue('LETTRE'))) then
          CreationTOBFille(TOBAFille,TOBEtat,TOBEtatFille,MontST, False);
        if ((TOBDFille <> NIL) and
           (LettreEnCours = TOBDFille.GetValue('LETTRE'))) then
          CreationTOBFille(TOBDFille,TOBEtat,TOBEtatFille,MontST, False);
        if ((TOBPFille <> NIL) and
           (LettreEnCours = TOBPFille.GetValue('LETTRE'))) then
          CreationTOBFille(TOBPFille,TOBEtat,TOBEtatFille,MontST, False);
      end;
    end;
    if ((TOBAFille <> NIL) and
        (LettreEnCours = TOBAFille.GetValue('LETTRE'))) then
      TOBAFille := TOBA.FindNext([''],[''],True);
    if (TOBAFille <> NIL) then
    begin
      EffectA := TOBAFille.GetValue('PDD_EFFECTIF');
      BaseA := TOBAFille.GetValue('PDD_BASECOTISATION');
      MtA := TOBAFille.GetValue('PDD_MTCOTISAT');
    end;

    if ((TOBDFille <> NIL) and
        (LettreEnCours = TOBDFille.GetValue('LETTRE'))) then
      TOBDFille := TOBD.FindNext([''],[''],True);
    if (TOBDFille <> NIL) then
    begin
      EffectD := TOBDFille.GetValue('PDD_EFFECTIF');
      BaseD := TOBDFille.GetValue('PDD_BASECOTISATION');
      MtD := TOBDFille.GetValue('PDD_MTCOTISAT');
    end;

    if ((TOBPFille <> NIL) and
        (LettreEnCours = TOBPFille.GetValue('LETTRE'))) then
      TOBPFille := TOBP.FindNext([''],[''],True);
    if (TOBPFille <> NIL) then
    begin
      EffectP := TOBPFille.GetValue('PDD_EFFECTIF');
      BaseP := TOBPFille.GetValue('PDD_BASECOTISATION');
      MtP := TOBPFille.GetValue('PDD_MTCOTISAT');
    end;

    if (TOBAFille <> NIL) then
      LettreEnCours := TOBAFille.GetValue('LETTRE');
    if (TOBDFille <> NIL) then
    begin
      if (LettreEnCours > TOBDFille.GetValue('LETTRE')) then
        LettreEnCours := TOBDFille.GetValue('LETTRE');
    end;
    if (TOBPFille <> NIL) then
    begin
      if (LettreEnCours > TOBPFille.GetValue('LETTRE')) then
        LettreEnCours := TOBPFille.GetValue('LETTRE');
    end;

  end;
  TOBA.free;
  TOBD.free;
  TOBP.free;
  MontST := 0.0;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 10/04/2003
Modifié le ... :   /  /    
Description .. : Traitement d'une Ducs Néant.
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
procedure TOF_PGMULDUCSRECAP.MajTobEtatNeant (var TDL : TOB; var TOBEtatNeant : TOB; POrg : TParamOrg; NumOrdre : integer);
var
  TOBEtatFille, TDLFille        : TOB;
  MontSt                        : double;
begin
  TDLFille := TDL.FindFirst([''],[''],True);

  if (TOBEtatNeant = NIL) then
  begin
    TOBEtatNeant := TOB.Create('Ducs Néant à éditer', NIL, -1);
  end;
  MontSt := 0.00;
  CreationTOBFille (TDLFille, TOBEtatNeant, TOBEtatFille, MontST, True);

end;

Initialization
registerclasses([TOF_PGMULDUCSRECAP]);
end.
