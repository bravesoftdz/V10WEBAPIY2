// mbo 25.10.05 ajout dans methode TMethodeAmort du plan traité (éco ou fiscal)
//              incidence sur fonctions copie et affecteMethodeAmort
// TGA 27/01/2006 Ajout de la propriété RepriseDep à TMethodeAmort
//              incidence sur fonctions copie et affecteMethodeAmort
// mbo 10/09/2007 - fq 21405
{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 09/10/2006 - nouveautés de la version 7
Modifié le ... :   /  /
Description .. : les 2 derniers éléments de la méthode sont
Suite ........ : Ces 2 éléments servent pour les méthodes PRI et SBV
Suite ........ : Pour les méthodes ECO et FISC ils sont à false
Suite ........ :
Suite ........ : creation = est positionné à true lors de la saisie de la prime
Suite ........ :                 ou de la subvention
Suite ........ : ReprisePrem = ne sert que pour la subvention
Suite ........ :               si = true ---> la reprise est cumulée sur la première
Suite ........ :                             dotation calculée
Suite ........ :              si false ---> la première dotation est calculée
Suite ........ :                            normalement et les antérieurs saisis
Suite ........ :                            sont bien à part

Suite ........ : Ajout d'une propriété = plan de référence qui va servir pour le calcul
Suite ........ : des dotations de la subvention en mode dégressif
Suite ........ : les valeurs : 0 = pour les méthodes eco, fisc et pri et SBV en non amortissable
Suite ........ :               1 = plan de référence = éco
Suite ........ :               2 = plan de référence = fiscal
Mots clefs ... :
*****************************************************************}

unit ImPlanMeth;

interface
uses
  HEnt1, AmType;
type

  TMethodeAmort = class
    Base         : double;
    Duree        : Integer;
    Methode      : string;
    Taux         : double;
    Reprise      : double;
    RepriseCedee : double;
    Revision     : double;
    DateFinAmort : TDateTime;
    Exceptionnel : double;
    BaseDebutExo : array [0..MAX_LIGNEDOT] of double;
    TableauDot   : array [0..MAX_LIGNEDOT] of double;
    TabCession   : array [0..MAX_LIGNEDOT] of double;
    BaseDebExoCes: array [0..MAX_LIGNEDOT] of double;
    //ajout mbo 25.10.05 pour savoir si on est plan éco ou plan fiscal (typeplan défini dans outils.pas)
    PlanUtil : string;
    RepriseDep : double;  //ajout tga 27.01.06
    Creation : boolean;   // ajout mbo 03.10.06  pour PRI
    ReprisePrem : boolean;  // ajout mbo 09.10.06 pour SBV
    PlanRef : integer;     // ajout mbo 25.10.06 pour SBV
  private
    { Déclarations privées }

  public
    { Déclarations publiques }
    // mbo 25.10.05 procedure copie(Source: TMethodeAmort) ;
    procedure copie(Source: TMethodeAmort; Tplan: string) ;
    procedure RazMethodeAmort ;
    function GetCumulDotExercice(ImmoClot : boolean;DateRef : TDateTime;TabDate : array of TDateTime;AvecCession : boolean): double;
    function GetCumulDotationCesMet(TabDate: array of TDateTime;tDateRef : TDateTime) : double;

    // mbo 25.10.05 procedure AffecteMethodeAmort(TabDate : array of TDateTime;wDateFinAmort: TDateTime;wRevision,wExceptionnel,wReprise,wRepriseCedee,wBase,wTaux: double;wDuree : integer;wMethode: string;wClot: boolean);
    //procedure AffecteMethodeAmort(TabDate : array of TDateTime;wDateFinAmort: TDateTime;wRevision,wExceptionnel,wReprise,wRepriseCedee,wBase,wTaux: double;wDuree : integer;wMethode: string;wClot: boolean;Tplan:string);
    // tga 27/01/06 Ajout reprisedep - ajout mbo 03.10.05 creation - ajout mbo 09.10.06 reprisePrem
    procedure AffecteMethodeAmort(TabDate : array of TDateTime;wDateFinAmort: TDateTime;
                                  wRevision,wExceptionnel,wReprise,wRepriseCedee,wBase,wTaux: double;
                                  wDuree : integer;wMethode: string;wClot: boolean;Tplan:string;
                                  wreprisedep:Double; wCreation:boolean; wReprisePrem:boolean;
                                  wPlanRef:integer);

    procedure ModifieMethodeAmort(wDateFinAmort: TDateTime;wRevision,wBase,wTaux: double;wDuree : integer;
                                  wMethode: string;wClot: boolean);

    procedure TraiteDotationCloture(TabDate : array of TDateTime;var IndDot : integer; var SommeDot : double;bCloture : boolean);
  end;

implementation

uses ImEnt;

procedure TMethodeAmort.copie(Source: TMethodeAmort; Tplan: string) ;
var i: integer ;
begin
  DateFinAmort:=Source.DateFinAmort;
  Revision    :=Source.Revision;
  Exceptionnel:=Source.Exceptionnel;
  Reprise     :=Source.Reprise;
  RepriseCedee:=Source.RepriseCedee;
  Base        :=Source.Base;
  Duree       :=Source.Duree;
  Methode     :=Source.Methode;
  Taux        :=Source.Taux;
  PlanUtil    := Tplan;               // ajout mbo 25.10.05
  RepriseDep  := Source.RepriseDep;   // TGA 27/01/2006
  Creation    := Source.Creation;     // ajout mbo 03.10.06
  ReprisePrem := Source.ReprisePrem;
  PlanRef     := Source.PlanRef;      // ajout mbo 25.10.06

  for i:=low(BaseDebutExo)  to high(BaseDebutExo)  do BaseDebutExo[i]:=Source.BaseDebutExo[i];
  for i:=low(BaseDebExoCes) to high(BaseDebExoCes) do BaseDebExoCes[i]:=Source.BaseDebExoCes[i];
  for i:=low(TableauDot)    to high(TableauDot)    do TableauDot[i]:=Source.TableauDot[i];
  for i:=low(TabCession)    to high(TabCession)    do TabCession[i]:=Source.TabCession[i];

end ;

// Initialisation d'une méthode - Valeurs par défaut
procedure TMethodeAmort.RazMethodeAmort;
begin
//  FillChar(self,Sizeof(self),#0) ;
  Base         :=0.0 ;
  Duree        :=0 ;
  Methode      :='' ;
  Taux         :=0.0 ;
  Reprise      :=0.0 ;
  RepriseDep   :=0.0 ;
  RepriseCedee :=0.0 ;
  Revision     :=0.0 ;
  DateFinAmort :=iDate1900 ;
  Exceptionnel :=0.0 ;
  PlanUtil     := '';     // ajout mbo 25.10.05
  creation     := false;  // ajout mbo 03.10.06 pour PRI
  ReprisePrem  := false;  // ajout mbo 09.10.06 pour SBV
  PlanRef      := 0;      // ajout mbo 25.10.06 pour SBV

  FillChar(BaseDebutExo,Sizeof(BaseDebutExo),#0) ;
  FillChar(TableauDot,Sizeof(TableauDot ),#0) ;
  FillChar(TabCession,Sizeof(TabCession),#0) ;
  FillChar(BaseDebExoCes,Sizeof(BaseDebExoCes),#0) ;

end;

// Affectation des données d'une méthode d'amortissement
procedure TMethodeAmort.AffecteMethodeAmort
                        (TabDate : array of TDateTime;wDateFinAmort: TDateTime;wRevision,wExceptionnel,
                        wReprise,wRepriseCedee,wBase,wTaux: double;wDuree : integer;
                        wMethode: string;wClot: boolean;Tplan: string;wRepriseDep:Double;
                        wcreation:boolean; wReprisePrem:boolean ; wPlanRef:integer);
var i,Ind : integer; SommeDot: double ;
begin
  Base         :=wBase;
  Duree        :=wDuree;
  Methode      :=wMethode;
  Taux         :=wTaux/100;
  Revision     :=wRevision;
  DateFinAmort :=wDateFinAmort;
  Exceptionnel :=wExceptionnel;
  Reprise      :=wReprise;
  RepriseCedee :=wRepriseCedee;
  Ind          := 0;
  PlanUtil     := Tplan;        // ajout mbo 25.10.05
  RepriseDep   :=wRepriseDep ;  // ajout Tga 27/01/2006
  Creation     := wCreation;    // ajout mbo 03.10.06
  ReprisePrem  := wReprisePrem;
  PlanRef      := wPlanRef;     // ajout mbo 25.10.06

  if wClot then TraiteDotationCloture(TabDate,Ind,SommeDot,false);
  for i:=Ind to MAX_LIGNEDOT do TableauDot[i]:=0.00;
  end;

procedure TMethodeAmort.ModifieMethodeAmort(wDateFinAmort: TDateTime;wRevision,wBase,wTaux: double;wDuree : integer;
                                            wMethode: string;wClot: boolean);
var i : integer;
begin
  Base        :=wBase;
  Duree       :=wDuree;
  Methode     :=wMethode;
  Taux        :=wTaux;
  Revision    :=wRevision;
  DateFinAmort:=wDateFinAmort;
  //  Reprise := On laisse valeur courante
  //  Exceptionnel := On laisse valeur courante
  DateFinAmort := iDate1900;
  if not wClot then for i:=0 to MAX_LIGNEDOT do TableauDot[i]:=0.00;
end;

function TMethodeAmort.GetCumulDotExercice(ImmoClot : boolean;DateRef : TDateTime;TabDate : array of TDateTime;AvecCession : boolean): double;
var SommeDot,Cumul : double;  Indice : integer;
begin
  Cumul:=0.00; SommeDot:=0.00; Indice:=0;
  if ImmoClot then TraiteDotationCloture(TabDate,Indice,SommeDot,false);
  while (TabDate[Indice]<>iDate1900) and (TabDate[Indice]<=DateRef) do
  begin
     Cumul:=Cumul+TableauDot[Indice];
     if (AvecCession) then Cumul:=Cumul+TabCession[Indice];
     inc(Indice);
  end;
  result:=Cumul+SommeDot;
end;

procedure  TMethodeAmort.TraiteDotationCloture(TabDate : array of TDateTime;var IndDot : integer; var SommeDot : double;bCloture : boolean);
var dtFinCumul : TDateTime;
begin
  if bCloture then dtFinCumul:=VHImmo^.Encours.Fin+1 else dtFinCumul:=VHImmo^.Encours.Deb;
  while (TabDate[IndDot]<>iDate1900) and (TabDate[IndDot]<dtFinCumul) do
  begin

    // mbo - 05.09.07 - fq 21405 SommeDot:=SommeDot+TableauDot[IndDot];
    SommeDot:=SommeDot+TableauDot[IndDot]+ TabCession[IndDot] ;
    inc(IndDot) ;
  end;
end;

function TMethodeAmort.GetCumulDotationCesMet(TabDate: array of TDateTime;tDateRef : TDateTime) : double;
var  Indice: integer;  MontantDot: double;
begin
  Indice:=0; MontantDot:=0.00;
  while (TabDate[Indice]<=tDateRef) and (TabCession[Indice]<>0.00) do
  begin
    MontantDot:=MontantDot+TabCession[Indice];
    inc(Indice) ;
  end;
  result:=MontantDot;
end;

end.
