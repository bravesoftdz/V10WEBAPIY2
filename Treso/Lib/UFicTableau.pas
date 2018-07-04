unit UFicTableau;
{Ancêtre des "faux muls" TRPREVANNUEL, TRPREVMENSUEL, DETAILSUIVI, TRSUIVISOLDE, TRCONTROLESOLDE
--------------------------------------------------------------------------------------
  Version    |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
   0.91        19/12/03  JP   Création de l'unité
 1.2X.000.000  16/04/04  JP   Regroupement de certains traitements des descendants
 1.50.001.001  12/05/04  JP   Gestion des comptes non mouvementés dans la gestion par solde
 6.0x.xxx.xxx  29/07/04  JP   Mise en place du suivi des commissions cf COMMISS
 6.00.018.001  12/10/04  JP   Mise en place du suivi bancaire
 6.2X.001.001  04/11/04  JP   Mise en place de la fiche de suivi des portefeuille d'OPCVM
 6.30.001.002  08/03/05  JP   FQ 10217 : gestion de la date d'opération lors de l'appel de la saisie.
                              FQ 10215 : Modification de la fonction GetPosition
 6.30.001.006  04/04/05  JP   FQ 10240 : Modification du suivi bancaire : on supprime l'option par
                              date d'opération et on met par date de rapprochement à la place.
 6.30.001.007  18/04/05  JP   FQ 10237 : Gestion de la périodicité sur les fiches de suivi
 6.30.001.011  23/05/05  JP   FQ 10253 : correction de la requête pour le suivi bancaire sous Oracle
                              FQ 10254 : modification de la périodicité de traitement
 6.50.001.002  01/06/05  JP   FQ 10266 : Suite à la FQ 10237, l'appel à CommanderDetail dans le suivi du
                              trésorier est incorrect, car la variable code n'est plus passé en paramètre
 6.50.001.021  03/10/05  JP   Remplacement de l'alias DEVISE par ADEVISE dans les requête car l'agl prend
                              DEVISE pour une table et avec le MultiSociétés le préfixe avec la base.
 6.51.001.001  28/11/05  JP   FQ 10317 : Gestion des soldes initiaux n'est pas correcte si on n'est pas
                              en affichage quotidien
 7.00.001.001  16/01/06  JP   FQ 10328 : Refonte de la gestion des devises : pour le moment ne concerne que
                              les suivis de solde et bancaire. Pour le suivi de Tréso, il me semble mieux
                              d'attendre que la fiche soit uniformisée avec les autres et pour les OPCVM
                              d'attendre la demande si le besoin s'en fait sentir
 7.00.001.001  27/01/06  JP   Modification de la gestion des filtres
 7.00.001.001  24/05/06  JP   FQ 10343 : Gestion des soldes initiaux dans la donne devise !!!!
 7.09.001.001  11/08/06  JP   Création de la fiche suivi de groupes et gestion des regroupements
 7.09.001.001  10/10/06  JP   FQ 10376 : fonction GetDateStr pour afficher le jour devant la Date
 7.09.001.001  07/11/06  JP   FQ 10256 : Déplacement des soldes de réinitialisation au matin
 7.09.001.005  02/07/07  JP   Le filtre sur les dossiers doit aussi se faire pour les bases autonomes dans un regroupement
 8.01.001.001  23/11/06  JP   Mise en place du Controle des soldes avec la compta et EEXBQ
 8.01.001.001  08/12/06  JP   Gestion de l'impression en eAGL : UobjEtats
 8.00.001.021  20/06/07  JP   FQ 10480 : gestion des concepts
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}

interface

uses
  StdCtrls, Controls, Classes, Menus, Windows,
  {$IFDEF EAGLCLIENT}
  eMul, MaineAGL,
  {$ELSE}
  Mul, FE_Main, PrintDBG,
  {$ENDIF}
  {$IFDEF TRCONF}
  ULibConfidentialite,
  {$ENDIF TRCONF}
  Forms, SysUtils, HCtrls, HEnt1, UTOF, HTB97, HMsgBox, HXlsPas, Grids,
  Spin, Graphics, UTob, Constantes;

type
  FICTAB = class (TOF)
    procedure OnUpdate              ; override;
    procedure OnArgument(S : string); override;
    procedure OnLoad                ; override;
    procedure OnDisplay             ; override;
    procedure OnClose               ; override;
  private
    {A Vrai, si l'on choisit plus d'une nature}
    FMultiNature : Boolean;
    {$IFDEF TRCONF}
    FObjConf : TObjConfidentialite;

    function  GetObjConf : TObjConfidentialite; 
    {$ENDIF TRCONF}

    function  GetMultiNature : Boolean;
    {08/03/05 : FQ 10217 : renvoie la date "courante" en fonction de la colonne active}
    function  GetCurDate : string;
    {18/04/05 : FQ 10237 : gestion de la périodicité}
    function  GetPeriodicite : TypePeriod;
    function  GetDateDepart  : TDateTime;
    function  GetNbColonne   : Integer;
    function  GetDateFin     : TDateTime;
    procedure SetNbColonne (Value : Integer);
    procedure SetDateDepart(Value : TDateTime);
    {27/01/06 : surcharge du chargement du filtre pour pouvoir intervenir après}
    procedure InitSelectFiltre(T : TOB);
    {08/12/06 : Impression en eAGL}
    procedure ImprimerClick;
  protected
    PtClick      : TPoint; {Coordonnées dans la grille de la souris}
    clUltraLight : TColor;
    clUltraCreme : TColor;
    clRougeMat   : TColor;
    clVertMat    : TColor;
    clCremePale  : TColor;
    clBleuLight  : TColor;
    clOrangePale : TColor;
    clGrisTexte  : TColor;
    FormKeyDown  : TKeyEvent;
    TypeDesc     : TDescendant; {Pour déterminer quelle est la fiche descendante appelante}
    LargeurCol   : Integer; {Largeur des colonnes pour le OnDrawCell}
    cbBanque     : THMultiValComboBox;
    FFormResize  : TNotifyEvent;
    FBChercClick : TNotifyEvent;
    {18/04/05 : FQ 10237 : Gestion de la périodicité}
    rbMensu    : TRadioButton;
    rbQuoti    : TRadioButton;
    rbHebdo    : TRadioButton;
    seInterval : THSpinEdit;
    edDepart   : THEdit;
    {12/09/06 : Pour les budgets : s'il n'est pas coché, présentation comparative}
    ckGlobal   : TCheckBox;

    procedure OnFormResize   (Sender : TObject); virtual;
    procedure BChercheClick  (Sender : TObject); virtual;
    procedure BImprimerClick (Sender : TObject); virtual;
    procedure bExportClick   (Sender : TObject); virtual;
    procedure bGraphClick    (Sender : TObject); virtual;
    procedure GridOnDblClick (Sender : TObject); virtual;
    procedure PopupSimulation(Sender : TObject); virtual;
    procedure NoDossierChange(Sender : TObject); virtual;
    procedure FormOnKeyDown  (Sender : TObject; var Key : Word; Shift : TShiftState); virtual;
    procedure OnDrawCell     (Sender : TObject; Col, Row : Integer; Rect : TRect; State : TGridDrawState); virtual;
    procedure DessinerCells  (Sender : TObject; Col, Row : Integer; Rect : TRect; State : TGridDrawState); virtual;
    procedure GetPosition;
    function  IsCelulleReinit(Col, Row : Integer) : Boolean; virtual;
    procedure FormAfterShow; virtual;{27/01/06}
    {27/01/06 : Pointeur, afin de pouvoir intervenir à la fin du chargement des filtres}
    procedure MySelectFiltre; virtual;
    {Clause where sur les banques et les comptes}
    function  GetWhereBqCpt : string; virtual;
    {Initialisation des contrôles des fiches de suivi lors de l'appel depuis une autre fiche}
    procedure InitControls; virtual;
    {Filtre la combo des dossiers sur les seuls dossiers Tréso du regroupement}
    procedure FiltreDossierTreso;
    {20/06/07 : FQ 10480 : Gestion du concept la création / Suppresion des flux}
    procedure GereDroitsCreation;
  public
    Societe    : string;
    Devise     : string;
    Nature     : string;
    LibNature  : string;

    {Liste contenant les cellules, dans un affichage par solde, qui n'ont pas d'écriture correspondante}
    ListeSolde : TStringList;
    {Liste avertissant qu'il y a un solde de réinitialisation}
    ListeRouge : TStringList;
    {Liste contenant tous les comptes}
    ListeCpte  : TOB;

    TobRappro, {05/04/05 : FQ 10240 : Pour le suivi bancaire, on utilise une seconde "TobListe"}
    TobListe,
    TobGrid    : TOB;
    rbDateOpe  : TRadioButton;
    Grid       : THGrid;
    cbDevise   : THValComboBox;
    InArgument : Boolean;
    sDateOpe   : string;
    {Indice de la colonne contenant les soldes forcés}
    ColRouge   : Integer;
    TREcritureModified : string;

    {28/11/05 : FQ 10317 : Liste contenant les opérations des premiers jours de l'année en
                affichage hebdomadaire}
    lOpeDebAnnee : TStringList;
    {16/01/06 : FQ 10328 : Gestion des devises : Nom du champ de la table TRECRITURE à
                partir duquel seront effectués les calculs TE_MONTANT ou TE_MONTANTDEV}
    ChpMontant : string;
    {27/01/06 : Si à True, on peut passer CritModified à True}
    CanSetModified : Boolean;

    {10/10/06 : FQ 10376 : affiche le jour devant la date}
    function GetDateStr(aDate : TDateTime) : string;
    {Permet d'ajouter dans la TobGrid des comptes qui n'y figureraient pas car non mouvementés}
    procedure SupprimeCptePresent(Bqe, Cpte : string); 
    {Permet d'ajouter dans la TobGrid des comptes qui n'y figureraient pas car non mouvementés}
    procedure SupprimeCptePresentMS(Cpte : string);

    procedure CreerListeCompte(ParBanque : Boolean = False);
    {Pour le suivi de solde qui est utilisé dans les fiches TofDetailSuivi et TRSuiviSolde_Tof}
    procedure ParSolde      ;
    {12/10/04 : Pour le suivi bancaire}
    procedure ParPointage   ;
    procedure MajComboDevise(Dev : string);
    {Constitution de la requête du OnUpdate}
    function GetRequete(Cl1 : string) : string;
    {Pour le remplissage des colonnes CODE et LIBELLE de la grille}
    function RetourneNomLib (Cpte, Nat : string) : string;
    function IsNbColOk : Boolean; virtual;
    {Calcul le solde initial pour un compte}
    function CalculSoldeInit(Gen, Dev : string; dt : TDateTime) : Double;
    {Constitue l'argument de l'appel à la fiche du detail des flux}
    function  CommandeDetail(CurRow : HTStrings; var TT : TOB) : Boolean; virtual;
    procedure CritereClick  (Sender : TObject); virtual;
    procedure MouetteClick  ;

    {16/01/06 : FQ 10328 : A True si on Travail sur TE_MONTANT, A False sur TE_MONTANTDEV}
    function  IsMntPivot     : Boolean;
    {18/04/05 : FQ 10237 : Retourne le titre de la colonne à une date donnée en fonction de la périodicité}
    function  GetTitreFromDate(CurDate : TDateTime) : string;
    {Retourne le titre de la colonne en fonction de l'indice}
    function  RetourneCol     (Col : Integer ) : string;
    {Retourne la date de début de la colonne en fonction de l'indice de la colonne.
     Si FinPeriode est à True, on retourne la date de fin}
    function  ColToDate       (Col : Integer; FinPeriode : Boolean = False) : TDateTime;
    {Retourne l'indice de la colonne SUIVANTE correspondant à la date}
    function  DateToCol       (Dt : TDateTime) : Integer;
    {JP 28/11/05 : FQ 10317 : Gestion des soldes de départ si début de millésime}
    procedure GereSoldeDepart(var TobG : Tob; Gen, Dev : string; Tau : Double);
    function  IsColAvecSoldeInit(Col : Integer) : Boolean;
    procedure SetlOpeDebAnnee(Gen : string; Mnt : Double; Dt : TDateTime);
    function  GetlOpeDebAnnee(Gen : string) : Double;
    {28/11/05 : FQ 10317 : la colonne courante est-elle celle des soldes de réinitialisation}
    function  HasSoldeReinit (Col : Integer) : Boolean;
    {24/01/06 : FQ 10328 : Affecte ChpMontant en fonction de la diversité des devises dans la sélection}
    procedure GereLesDevises; virtual;

    {16/08/06 : Génèré la clause where sur les dossiers appartenant au regroupement choisi :
                par défaut on travaille sur TE_NODOSSIER, mais cela peut être BQ_NODOSSIER ...}
    function  GetWhereDossier(Chp : string = 'TE_NODOSSIER') : string;
    {10/08/07 : Gestion des filtres sur les confidentialité}
    function  GetWhereConfidentialite : string; virtual;
    {16/08/06 : Retourne True si l'on est sur une fiche gérant les regroupements}
    function  IsAvecRegroupement : Boolean;
    {12/09/06 : Pour les budgets non globalisés, les natures ne correspondent pas aux valeurs de
                TE_NATURE, mais sont recalculées afin de pouvoir être distingués dans le DrawCell}
    function  GetNatureToBudget  (TypeFlux, Nat : string) : string;
    {12/09/06 : Cheminement inverse de ci-dessus : retourne la nature pour TRECRITURE}
    function  GetNatureFromBudget(TypeFlux, CodeFlux, Nat : string) : string;

    class function FicLance (Dom, Fiche, Range, Lequel, Arguments : string) : string;

    property MultiNature : Boolean read GetMultiNature write FMultiNature;
    {Clause where sur les banques et les comptes}
    property WhereBqCpt : string read GetWhereBqCpt;
    {08/03/05 FQ 10217 : Date courante pour la saisie de flux en fonction de la colonne active}
    property DateCourante : string read GetCurDate;
    {18/04/05 : FQ 10237 : Gestion de la périodicité}
    property Periodicite : TypePeriod read GetPeriodicite;
    property DateFin     : TDateTime  read GetDateFin;
    property DateDepart  : TDateTime  read GetDateDepart write SetDateDepart;
    property NbColonne   : Integer    read GetNbColonne  write SetNbColonne;
    {16/01/06 : FQ 10328 : A True si on Travail sur TE_MONTANT, A False sur TE_MONTANTDEV}
    property MntPivot    : Boolean    read IsMntPivot default True;
    {$IFDEF TRCONF}
    property ObjConf     : TObjConfidentialite read GetObjConf write FObjConf;
    {$ENDIF TRCONF}
  end;

const
  {Colonnes dans la grille}
  COL_TYPE    = 0;
  COL_CODE    = 1;
  COL_NATURE  = 2;
  COL_LIBELLE = 3;
  COL_DATEDE  = 4;
  {$IFDEF EAGLCLIENT}
  COL_IMPMAX  = 10;
  {$ENDIF}
  NBJOURBIMENS = 15;

implementation

uses
  TofDetailFlux, Commun, TRGRAPHSUIVI_TOF, UProcGen, UProcSolde, AglInit,
  uTobDebug, Filtre, UtilPgi, UObjEtats;

{---------------------------------------------------------------------------------------}
class function FICTAB.FicLance(Dom, Fiche, Range, Lequel, Arguments : string) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
var
  C1 : TColor;
begin
  inherited;
  ChpMontant := 'TE_MONTANT'; {16/01/06 : FQ 10328}
  CanSetModified := False;

  Ecran.HelpContext := 150;
  {Largeur par défaut}
  LargeurCol := 100;
  {Couleur claire (voir TomCalendrier)}
  if ColorToRGB(clBtnHighLight) <> ColorToRGB(clWindow) then
    clUltraLight := clBtnHighLight
  else begin
    C1 := ColorToRGB(clBtnFace);
    {Couleur intermédiaire entre clBtnFace et clWhite}
    clUltraLight := RGB(((C1 and $000000FF) + $000000FF) shr 1,
                        ((C1 and $0000FF00) + $0000FF00) shr 9,
                        ((C1 and $00FF0000) + $00FF0000) shr 17);
  end;

  {Couleur intermédiaire entre le blanc et le jaune}
  clUltraCreme := $00D6F3F2;
  clCremePale  := $00E1F0F2;
  clRougeMat   := $00A8A9CE;
  clVertMat    := $00B7C8B0;
  clBleuLight  := $00FFF2F4;
  clOrangePale := $00CADBFF;
  clGrisTexte  := $00535353;

  ListeSolde := TStringList.Create;
  ListeRouge := TStringList.Create;
  ListeRouge.Duplicates := dupIgnore;
  ListeRouge.Sorted := True;

  lOpeDebAnnee := TStringList.Create; {28/11/05 : FQ 10317}

  FormKeyDown := Ecran.OnKeyDown;
  Ecran.OnKeyDown := FormOnKeyDown;

  Grid := THGrid(GetControl('GRID'));
  Grid.OnDrawCell := OnDrawCell;
  Grid.OnDblClick := GridOnDblClick;
  Grid.FixedCols := COL_DATEDE;
  Grid.ColWidths[COL_TYPE] := -1;
  Grid.ColWidths[COL_CODE] := -1;
  Grid.ColWidths[COL_NATURE] := -1;
  Grid.ColWidths[COL_LIBELLE] := 150;

  cbDevise := THValComboBox(GetControl('CBDEVISE'));
  if not (TypeDesc in [dsCommission, dsOPCVM]) then begin
    rbDateOpe := TRadioButton(GetControl('DATEOPE'));
    rbDateOpe.OnClick := CritereClick;
  end;

  {Dans dsDetail, dsSolde, dsTreso}
  if not (TypeDesc in [dsOPCVM, dsControle]) then
    cbBanque := THMultiValComboBox(GetControl('BANQUE'));

  {05/04/05 : FQ 10240 : Pour constituer la TobGrid, on utilise maintenant deux Tobs :
              - TobListe sur la date de Valeur / Opération
              - TobRappro sur la date de rapprochement}
  if TypeDesc in [dsBancaire, dsControle] then
    TobRappro := TOB.Create('', nil, -1);

  TobListe := TOB.Create('', nil, -1);
  TobGrid  := Tob.Create('', nil, -1);

  TREcritureModified := '';
  FBChercClick := TToolbarButton97(GetControl('BCHERCHE')).OnClick;
  TToolbarButton97(GetControl('BCHERCHE')).OnClick := BChercheClick;
  {$IFDEF EAGLCLIENT}
  {En CWAS, le bouton Chercher essaie de travailler sur la DBListe de la forme ;
   or ici, elle n'est pas renseignée}
  SetControlVisible('BPREV', False);
  SetControlVisible('BNEXT', False);
  {$ENDIF}
  TToolbarButton97(GetControl('BIMPRIMER')).OnClick := BImprimerClick;
  TToolbarButton97(GetControl('BEXPORT'  )).OnClick := bExportClick;

  if TypeDesc <> dsControle then begin
    TToolbarButton97(GetControl('BGRAPH'   )).OnClick := bGraphClick;
    TPopupMenu(GetControl('POPUPMENU')).Items[0].OnClick := GridOnDblClick;
    TPopupMenu(GetControl('POPUPMENU')).Items[1].OnClick := PopupSimulation;
    if TypeDesc = dsOPCVM then begin
      TPopupMenu(GetControl('POPUPMENU')).Items[0].Caption := TraduireMemoire('Afficher le détail des mouvements');
      TPopupMenu(GetControl('POPUPMENU')).Items[1].Caption := TraduireMemoire('Créer un nouvel OPCVM');
    end;
  end;

  {18/04/05 : FQ 10237 : Gestion de la périocité}
  if TypeDesc in [dsTreso, dsDetail, dsSolde, dsMensuel, dsCommission, dsBancaire, dsOPCVM, dsGroupe, dsControle] then begin
    seInterval := THSpinEdit(GetControl('INTERVAL'));
    seInterval.OnChange := CritereClick;
    if not(TypeDesc in [dsMensuel]) then begin
      if TypeDesc <> dsControle then begin
        rbMensu := TRadioButton(GetControl('RBMENSU'));
        rbQuoti := TRadioButton(GetControl('RBQUOTI'));
        rbMensu.OnClick := CritereClick;
        rbQuoti.OnClick := CritereClick;
      end;
      seInterval.MaxValue := 185;
    end;

    edDepart := THEdit(GetControl('DATEDE'));
    edDepart.OnChange := CritereClick;

    if not (TypeDesc in [dsMensuel, dsTreso, dsControle]) then begin
      rbHebdo := TRadioButton(GetControl('RBHEBDO'));
      rbHebdo.OnClick := CritereClick;
    end;
  end;

  TFMul(Ecran).OnAfterFormShow := FormAfterShow;
  {Pas de saisie dans le suivi de groupes car il s'agit de cumul}
  if not (TypeDesc in [{dsGroupe,} dsControle]) then
    ADDMenuPop(TPopupMenu(GetControl('POPUPMENU')), '', '');

  FFormResize := Ecran.OnResize;
  Ecran.OnResize := OnFormResize;
  if TypeDesc = dsBancaire then
    {Affectation de la liste ad-hoc}
    TFMul(Ecran).SetDBListe('TRSUIVIBANCAIRE');

  if TypeDesc in [dsPeriodique, dsMensuel] then
    {12/09/06 : Gestion du comparatif budgétaire}
    ckGlobal   := TCheckBox(GetControl('CKGLOBALISE'));

  if IsAvecRegroupement then begin
    SetControlVisible('REGROUPEMENT', IsTresoMultiSoc);
    SetControlVisible('TREGROUPEMENT', IsTresoMultiSoc);

    if IsTresoMultiSoc and (TypeDesc in [dsSolde, dsGroupe, dsBancaire, dsTreso]) then begin
       {Si détail en affichage en solde, alors TypeDesc = dsSolde, mais il n'y a pas de combo REGROUPEMENT}
      if (Assigned(GetControl('REGROUPEMENT'))) then begin
        THMultiValComboBox(GetControl('REGROUPEMENT')).OnChange := NoDossierChange;
        THMultiValComboBox(GetControl('REGROUPEMENT')).Plus := 'DOS_NODOSSIER ' + FiltreNodossier;
      end;
    end
    else if IsTresoMultiSoc and Assigned(GetControl('REGROUPEMENT')) then begin
      THValComboBox(GetControl('REGROUPEMENT')).OnChange := NoDossierChange;
      THValComboBox(GetControl('REGROUPEMENT')).Plus := 'DOS_NODOSSIER ' + FiltreNodossier;
    end;
  end;

  {Filtre la combo des dossiers sur les seuls dossiers Tréso du regroupement}
  FiltreDossierTreso;

  {Pour le moment, seule l'appel de la fiche de suivi de solde est appelé avec TheTob depuis une autre
   fiche. Le detail de suivi est appelé depuis les virements et la fiche de suivi avec une string}
  if TypeDesc in [dsSolde] then
    if Assigned(LaTob) then TFMul(Ecran).FiltreDisabled := True;

  {20/06/07 : FQ 10480 : Gestion du concept la création / Suppresion des flux}
  GereDroitsCreation;
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.OnLoad;
{---------------------------------------------------------------------------------------}
begin
  InitControls;

  if TypeDesc in [dsDetail, dsMensuel, dsSolde, dsTreso, dsCommission, dsBancaire, dsOPCVM, dsGroupe, dsControle] then begin
    if seInterval.Value <= 0 then seInterval.Value := 3;
  end;

  inherited;
  ChpMontant := 'TE_MONTANT';
  {Affecte ChpMontant en fonction de la sélection}
  GereLesDevises;
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.OnUpdate;
{---------------------------------------------------------------------------------------}
begin
  LibereListe(lOpeDebAnnee, False); {28/11/05 : FQ 10317}
  inherited;
  {En fonction du type de date}
  if not (TypeDesc in [dsCommission, dsOPCVM]) then begin
    if rbDateOpe.Checked then sDateOpe := 'TE_DATECOMPTABLE'
                         else sDateOpe := 'TE_DATEVALEUR';
  end;

  {20/04/05 : FQ 10237 : redimension des colonnes en fonction de la périodicité}
  if TypeDesc in [dsDetail, dsSolde, dsCommission, dsBancaire, dsOPCVM, dsGroupe] then begin
    if Periodicite = tp_1 then LargeurCol := 100
                          else LargeurCol := 137;
  end;

  Grid.VidePile(False);
  {On vide les deux tobs}
  TobListe.ClearDetail;
  TobGrid .ClearDetail;
  {05/04/05 : FQ 10240 : Ajout d'une deuxième tob pour le suivi bancaire contenant les écritures rapprochées}
  if Assigned(TobRappro) then TobRappro.ClearDetail;
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.OnDisplay;
{---------------------------------------------------------------------------------------}
begin
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.OnClose;
{---------------------------------------------------------------------------------------}
begin
  TFMul(Ecran).Retour := TREcritureModified;
  if Assigned(TobListe) then FreeAndNil(TobListe);
  if Assigned(TobGrid ) then FreeAndNil(TobGrid);
  if Assigned(TobRappro)  then FreeAndNil(TobRappro); {05/04/05 : FQ 10240}
  if Assigned(ListeSolde) then FreeAndNil(ListeSolde);
  if Assigned(ListeRouge) then FreeAndNil(ListeRouge);
  if Assigned(ListeCpte)  then FreeAndNil(ListeCpte);
  if Assigned(lOpeDebAnnee) then LibereListe(lOpeDebAnnee, True); {28/11/05 : FQ 10317}

  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.BChercheClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if not IsNbColOk then Exit;
  FBChercClick(Sender);
  (*
  {$IFDEF EAGLCLIENT}
  OnLoad;
  OnUpdate;
  {$ELSE}
  FBChercClick(Sender);
  {$ENDIF}
  *)
  TFMul(Ecran).CritModified := False;
end;

{Imprime la grille rajoutée au lieu de la DB grille
{---------------------------------------------------------------------------------------}
procedure FICTAB.BImprimerClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  ImprimerClick
end;

{Exporte la grille rajoutée au lieu de la DB grille
{---------------------------------------------------------------------------------------}
procedure FICTAB.bExportClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if ExJaiLeDroitConcept(ccExportListe, True) and TFMul(Ecran).SD.Execute then
    ExportGrid(Grid, Nil, TFMul(Ecran).SD.FileName, TFMul(Ecran).SD.FilterIndex, True);
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.CritereClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {FQ 10237 : gestion du libellé en fonction de la périodicité}
  {FQ 10254 : modification des maxima pour les différents types de périodes}
  if Sender is TRadioButton then begin
    if (Sender as TRadioButton).Name = 'RBMENSU' then begin
      SetControlText('LBPREC', 'mois suivants');
      seInterval.MaxValue := 24;
    end
    else if (Sender as TRadioButton).Name = 'RBHEBDO' then begin
      SetControlText('LBPREC', 'semaines suivantes');
      seInterval.MaxValue := 52;
    end
    else if (Sender as TRadioButton).Name = 'RBQUOTI' then begin
      if Periodicite = tp_1 then begin
        SetControlText('LBPREC', 'jours suivants');
        seInterval.MaxValue := 185;
      end
      else begin
        SetControlText('LBPREC', 'quinzaines suivantes');
        seInterval.MaxValue := 52;
      end;
    end;
  end;

  if V_PGI.AutoSearch and not InArgument then // Si MulAutoSearch...
    TFMul(Ecran).ChercheClick;
  {$IFDEF EAGLCLIENT}
  {$ELSE}
//  TFMul(Ecran).CritModified := True;
  {$ENDIF EAGLCLIENT}
  {27/01/06 : Pour éviter un double passage dans la recherche en ouverture de fiche
              s'il y a un filtre}
  if CanSetModified then
    TFMul(Ecran).CritModified := True;
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.bGraphClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Titre  : string;
  sSolde : string;
  sDev   : string;
  n, p   : Integer;
  T, D   : TOB;
  sGraph : string;
begin
  if TypeDesc in [dsSolde, dsDetail] then begin
    {20/11/03 : s'il y a au moins deux éléments de sélectionnés}
    if (Pos(',', GetClauseIn(cbBanque.Value)) > 0) or (cbBanque.Value = '') then
      Titre := 'Évolution des soldes bancaires'
    else
      Titre := 'Évolution du solde de la banque : ' + Grid.Cells[COL_LIBELLE, 1];
    if rbDateOpe.Checked then sSolde := 'Les montants sont calculés en date d''opération'
                         else sSolde := 'Les montants sont calculés en date de valeur';
    sDev  := 'Les montants sont éxprimés en ' + cbDevise.Items[cbDevise.ItemIndex];

    {Constitution de la chaine qui servira aux paramètres de "LanceGraph"}
    sGraph := 'DATE;';

    {Constitution de la tob des flux}
    T := Tob.Create('$$$', nil, -1);
    {Les lignes contenant les soldes des comptes ont une nature à "T"}
    for p := 1 to Grid.RowCount - 1 do
      if Grid.Cells[COL_NATURE, p] = 'T' then
        sGraph := sGraph + Grid.Cells[COL_LIBELLE, p] + ';';

    for n := COL_DATEDE to Grid.ColCount - 1 do begin
      D := Tob.Create('$$$', T, -1);
      D.AddChampSup('DATE', False);
      D.PutValue('DATE', Grid.Cells[n, 0]);
      for p := 1 to Grid.RowCount - 1 do begin
        {Les lignes contenant les soldes des comptes ont une nature à "T"}
        if Grid.Cells[COL_NATURE, p] = 'T' then begin
          D.AddChampSup(Grid.Cells[COL_LIBELLE, p], False);
          D.PutValue(Grid.Cells[COL_LIBELLE, p], Grid.Cells[n, p]);
        end;
      end;
    end;

    {!!! Mettre sGraph en fin à cause des ";" qui traine dans la chaine}
    TRLanceFiche_GraphSuivi(Titre + ';' + sDev + ';' + sSolde + ';' + sGraph, T);
    T.ClearDetail;
    FreeAndNil(T);

  end;
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.GridOnDblClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Code    : string;
  CurRow  : HTStrings;
  General : string;
  TT      : TOB;
begin
  GetPosition;
  {On est sur une colonne Date}
  if (PtClick.X >= COL_DATEDE) and (PtClick.Y > 0) then begin
    CurRow := Grid.Rows[PtClick.Y];
    Code := CurRow[COL_CODE];

    {30/04/04 : Totaux recettes et dépenses dans les deux écrans du budget}
    if CurRow[COL_TYPE] = '£' then Exit;

    {On est sur une ligne détail}
    if Code <> '' then begin
      TT := InitTobPourDetail;
      try
        {Fonction à surcharger qui constitue l'argument de l'appel à la fiche du detail des flux}
        if not CommandeDetail(CurRow, TT) then Exit;
        {Récupération du mois}
        if Periodicite = tp_30 then
          TT.SetString('MOIS', ' pour le mois de ' + AnsiLowerCase(Grid.Cells[PtClick.X, 0]))
        else if Periodicite <> tp_1 then
          TT.SetString('MOIS', ' pour la période ' + AnsiLowerCase(Grid.Cells[PtClick.X, 0]))
        else
          TT.SetString('MOIS', ' du ' + Grid.Cells[PtClick.X, 0]);

        {01/06/05 : FQ 10266 : En ayant enlevé un paramètre à la fonction CommanderDetail pour la FQ 10237,
                    "Code", n'était plus renseigné avec les arguments nécessaires à l'appel du détail}
        if TypeDesc = dsTreso then Code := TT.GetString('ARGTRESO')
                              else Code := '';
        TheTob := TT;
        {Appel de la fiche}
        Code := TRLanceFiche_DetailFlux('TR', 'TRDETAILFLUX', '', '', Code);

        if Code <> '' then begin
          TREcritureModified := Code;
          if (Code <> 'R') and (Code <> 'N') then
            RecalculSolde(General, Code, 0, True);
          {Rafraichissement en cas de modifications dans TREcriture}
          MouetteClick;
        end;
      finally
        if Assigned(TT) then FreeAndNil(TT);
      end;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.MouetteClick;
{---------------------------------------------------------------------------------------}
begin
  TFMul(Ecran).ChercheClick;
end;

{Met dans PtClick les coordonnées de la case cliquée
{---------------------------------------------------------------------------------------}
procedure FICTAB.GetPosition;
{---------------------------------------------------------------------------------------}
begin
  {08/03/05 : FQ 10215 : Le code ci-dessous est peut-être plus brillant que l'appel aux
              propriétés Col et Row mais il marche surtout beaucoup moins bien !
  GetCursorPos(Pt);
  Pt := Grid.ScreenToClient(Pt);
  Grid.MouseToCell(Pt.X, Pt.Y, PtClick.X, PtClick.Y);}
  PtClick.X := Grid.Col;
  PtClick.Y := Grid.Row;
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.MajComboDevise(Dev : string);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  n := cbDevise.Values.IndexOf(Dev);
  if n = -1 then n := cbDevise.Values.IndexOf(V_PGI.DevisePivot);
  if n > -1 then cbDevise.ItemIndex := n
            else cbDevise.ItemIndex := 0;
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.FormOnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
{---------------------------------------------------------------------------------------}
var
  I : Integer;

  {-----------------------------------------------------------------------}
  function Reached(Col, Row: Integer): Boolean;
  {-----------------------------------------------------------------------}
  begin
    {True si on n'est pas sur une ligne de solde ou de total final}
    Result := (Grid.Cells[Col, Row] <> '') and
              (Grid.Cells[COL_NATURE, Row] <> na_Total) and
              (Grid.Cells[COL_NATURE, Row] <>  '') and
              (Grid.Cells[COL_TYPE, Row] <> '£'  ) and
              (Grid.Cells[COL_TYPE, Row] <> '+'  ) and
              (Grid.Cells[COL_TYPE, Row] <> '-');

    if TypeDesc in [dsBancaire, dsSolde] then
      Result := Result and (Grid.Cells[COL_TYPE, Row] <> '*')
    else if TypeDesc = dsBancaire then
      Result := Result and (ListeRouge.IndexOf(Grid.Cells[Col, 0] + '|' +
                                               Grid.Cells[COL_CODE, Row] + '|' +
                                               Grid.Cells[COL_NATURE, Row] + '|') > -1)
    else if TypeDesc = dsControle then
      Result := Result and (
                (ListeRouge.IndexOf(Grid.Cells[Col, 0] + '|T|') > -1) or
                (ListeRouge.IndexOf(Grid.Cells[Col, 0] + '|$|') > -1));
  end;

begin
  if not (TypeDesc in [dsPeriodique, dsTreso]) and (Shift = [ssCtrl]) then begin
    case Key of
      VK_RIGHT : for I := Grid.Col + 1 to Grid.ColCount - 1 do
                   if Reached(I, Grid.Row) then	begin
                      Grid.Col := I;
                      Key := 0;
                      Break;
                    end;

      VK_LEFT :	for I := Grid.Col - 1 downto COL_DATEDE do
                  if Reached(I, Grid.Row) then	begin
                    Grid.Col := I;
                    Key := 0;
                    Break;
                  end;
      VK_DOWN : for I := Grid.Row + 1 to Grid.RowCount - 3 do
                  if Reached(Grid.Col, I) then begin
                    Grid.Row := I;
                    Break;
                  end;
      VK_UP : for I := Grid.Row - 1 downto 2 do
                if Reached(Grid.Col, I) then begin
                  Grid.Row := I;
                  Break;
                end;
    end;
  end;

  FormKeyDown(Sender, Key, Shift); // Pour touches standard AGL
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.OnDrawCell(Sender: TObject; Col, Row: Integer; Rect: TRect; State: TGridDrawState);
{---------------------------------------------------------------------------------------}
var
  S, sType : string;
  N        : Integer;
begin
  S := Grid.Cells[Col, Row];
  with Grid.Canvas do begin
    Font.Style := [];
    Font.Color := clWindowText;
    Brush.Color := clWindow;

    {Ligne des titres}
    if Row = 0 then begin
      Brush.Color := clBtnFace;
      FillRect(Rect);
      DrawFrameControl(Handle, Rect, DFC_BUTTON, DFCS_BUTTONPUSH);
      N := DT_CENTER or DT_VCENTER or DT_SINGLELINE;
      {Elimine le resize automatique des colonnes !}
      if Grid.ColWidths[Col] < LargeurCol then Grid.ColWidths[Col] := LargeurCol;
    end

    {Lignes de valeurs}
    else begin
      {La fiche de suivi du trésorier fonctionne différemment}
      if TypeDesc in [dsTreso, dsCommission, dsOpcvm, dsGroupe] then
        DessinerCells(Sender, Col, Row, Rect, State)
      {Pour les autres fiches}
      else begin
        sType := Grid.Cells[COL_TYPE, Row];
        {Sur type flux ou Banque}
        if sType = '' then begin
          Font.Style := [fsBold];
          Brush.Color := clUltraLight;
        end;

        if gdSelected in State then begin
          Brush.Color := clHighLight;
          Font.Color := clHighLightText;
        end

        {On est sur des soldes}
        else if sType = '*' then begin
          Brush.Color := clBtnFace;
          if TypeDesc = dsMensuel then begin
            Font.Style := [fsBold];
            if (Col >= COL_DATEDE) then begin
                   if (Valeur(S) > 0) then Font.Color := clGreen {Solde positif}
              else if (Valeur(S) < 0) then Font.Color := clRed; {Solde négatif}
            end;
          end;
        end
        else
          DessinerCells(Sender, Col, Row, Rect, State);
      end;{else if TypeDesc}

      FillRect(Rect);

      if Col = COL_LIBELLE then
        N := DT_LEFT or DT_VCENTER or DT_SINGLELINE {Pas de relief pour FixedCol}
      else
        N := DT_RIGHT or DT_VCENTER or DT_SINGLELINE;
    end; {else if Row = 0}

    InflateRect(Rect, -2, -2);
    DrawText(Handle, PChar(S), Length(S), Rect, N);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.DessinerCells(Sender: TObject; Col, Row: Integer; Rect: TRect; State: TGridDrawState);
{---------------------------------------------------------------------------------------}
var
  S, sType : string;
  ch : string;
begin
  {La fiche de suivi du trésorier est traitée dans la fiche hérité}
  if TypeDesc in [dsTreso, dsCommission, dsOpcvm, dsGroupe, dsControle] then Exit;

  S := Grid.Cells[Col, Row];
  sType := Grid.Cells[COL_TYPE, Row];

  with Grid.Canvas do begin
    if TypeDesc in [dsPeriodique, dsMensuel] then begin {COMMISS}
      {On est sur des totaux}
      if (sType = '+') or (sType = '-') then begin
        if V_PGI.NumAltCol = 0 then Brush.Color := clInfoBk
                               else Brush.Color := AltColors[V_PGI.NumAltCol];
        Font.Style := [fsBold];
      end

      {On est sur les totaux recettes et depenses}
      else if sType = '£' then begin
        if Grid.Cells[COL_CODE, Row] = '+' then Brush.Color := clVertMat {Vert pale}
                                           else Brush.Color := clRougeMat;{Rouge pale}
        Font.Style := [fsBold];
      end
      {10/11/06 : On est sur le solde initial ou final du budget périodique}
      else if (sType = '¥') and (TypeDesc = dsPeriodique) then begin
        Brush.Color := clBtnFace;
        Font.Style := [fsBold];
      end;

      {25/08/06 : gestion du comparatif budgétaire}
      if not (GetCheckBoxState('CKGLOBALISE') = cbChecked) then begin
        {Sur des écritures budgétaires}
        if (Grid.Cells[COL_NATURE, Row] = na_Simulation) or
           (Grid.Cells[COL_NATURE, Row] = na_PrevSimul) then
          Brush.Color := clCremePale;
      end;

      if IsCelulleReinit(Col, Row) then Font.Color := clBlue;
    end

    else if TypeDesc in [dsDetail, dsSolde, dsBancaire] then begin
      {On est sur des totaux}
      if (sType = '+') or (sType = '-') then begin
        if V_PGI.NumAltCol = 0 then Brush.Color := clInfoBk
                               else Brush.Color := AltColors[V_PGI.NumAltCol];
        {Solde général, par solde}
        if ((TypeDesc in [dsSolde, dsBancaire]) and (sType = '-')) or
           ((TypeDesc = dsDetail) and (sType = '-') and (TraduireMemoire('Solde général') = Grid.Cells[COL_LIBELLE, Row]))then begin
          Font.Style := [fsBold];
          Brush.Color := clBtnFace;
        end;
      end

      {Par solde, les lignes des comptes on comme type le libellé de leur banque}
      {08/06/04 : le problème c'est qu'un libellé peut être sur moins de 4 caractères : BNP}
      else if TypeDesc in [dsSolde, dsBancaire] then begin
        if ((Grid.Cells[COL_NATURE, Row] = na_Total) and (TypeDesc = dsSolde)) or
           ((Grid.Cells[COL_NATURE, Row] = na_Estime) and (TypeDesc = dsBancaire)) then begin
          Brush.Color := clUltraCreme;
          if (Col >= COL_DATEDE) then begin
            if (Valeur(S) >= 0) then Font.Color := clGreen {Solde positif}
                                else Font.Color := clRed; {Solde négatif}
          end;
        end;
        ch := Grid.Cells[Col, 0] + '|' + Grid.Cells[COL_CODE, Row] + '|' + Grid.Cells[COL_NATURE, Row] + '|';
        if (TypeDesc = dsBancaire) and (ListeRouge.IndexOf(ch) > -1) then
          Font.Style := [fsBold];
      end;

      if TypeDesc = dsDetail then begin
        {si par flux et au premier janvier}
        if (Grid.Cells[1, Row] = CODEREGULARIS) and HasSoldeReinit(Col) then {28/11/05 : FQ 10317}
          Font.Color := clBlue;
      end
      else begin
        if (TypeDesc = dsBancaire) and HasSoldeReinit(Col) then {28/11/05 : FQ 10317}
          Font.Color := clBlue
        {Si l'on est sur la colonne du premier janvier, que le compte a un montant de réinitialisation
         et que l'on est sur une ligne de solde par compte}
        else if HasSoldeReinit(Col) and {28/11/05 : FQ 10317}
           (ListeRouge.IndexOf(Grid.Cells[COL_CODE, Row]) > -1) and
           (Grid.Cells[COL_NATURE, Row] = na_Total )then Font.Color := clBlue;
      end;
    end;
  end; {With Canvas}
end;

{---------------------------------------------------------------------------------------}
function FICTAB.IsCelulleReinit(Col, Row : Integer) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := False;
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.PopupSimulation(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
//
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.NoDossierChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  if TypeDesc in [dsPeriodique, dsMensuel] then begin
    THValComboBox(GetControl('GENERAL')).Plus := FiltreBanqueCp(THValComboBox(GetControl('GENERAL')).DataType, '', GetControlText('REGROUPEMENT'));
    THValComboBox(GetControl('GENERAL')).ItemIndex := -1;
  end
  else if TypeDesc in [dsSolde, dsBancaire] then begin
    if THMultiValComboBox(Sender).Tous then s := ''
                                       else s := GetControlText('REGROUPEMENT');
    THMultiValComboBox(GetControl('MCCOMPTE')).Plus := FiltreBanqueCp(THMultiValComboBox(GetControl('MCCOMPTE')).DataType, '', s);
    s := THMultiValComboBox(GetControl('MCCOMPTE')).Plus;
    SetControlText('MCCOMPTE', '');
  end
  else if TypeDesc = dsControle then begin
    THEdit(GetControl('GENERAL')).Plus := FiltreBanqueCp(THEdit(GetControl('GENERAL')).DataType, '', s);
    SetControlText('GENERAL', '');
  end;
end;

{---------------------------------------------------------------------------------------}
function FICTAB.CommandeDetail(CurRow : HTStrings; var TT : TOB) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := False;
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.ParPointage;
{---------------------------------------------------------------------------------------}
var
  aTob,
  TobL,
  TobP,
  TobG : TOB;
  lBanque : string; {20/11/03}
  MntTmp  : Double; {20/11/03}
  CodeBqe : string; {12/05/04}
  txConv  : Double;
  {LÉGENDE DES TYPES : "LibBanque" = Ligne contenant les différents soldes d'un compte bancaire
                       ""          = Ligne contenant le libellé d'une banque
                       "*"         = Ligne de séparation entre deux banques
                       "+"         = Solde total d'une banque
                       "-"         = Solde général
   LÉGENDE DES NATURES : "P", "B", "E" = Soldes par Type (non pointé, pointé, total)}

    {------------------------------------------------------------------}
    procedure CreeTobDetail(TypeFlux, CodeFlux, LibelleFlux, Nat : string; X2 : Boolean);
    {------------------------------------------------------------------}
    var
      n   : Integer;
      Mnt : Double;
      TR  : TOB;
      Titre : string;
      sDate : string;
      First : Boolean;
      Col   : Integer;
      IniOk : Boolean;
    begin
      {Tob contenant le solde général}
      TobG := TOB.Create('', TobGrid, -1);
      sDate := DateToStr(DateDepart);

      TobG.AddChampSupValeur('TYPE', TypeFlux);
      TobG.AddChampSupValeur('CODE', CodeFlux);
      if X2 then begin
        TobG.AddChampSupValeur('NATURE' , na_Estime);
        TobG.AddChampSupValeur('LIBELLE', RetourneNomLib(LibelleFlux, na_Estime));
      end
      else begin
        TobG.AddChampSupValeur('NATURE' , Nat);
        TobG.AddChampSupValeur('LIBELLE', LibelleFlux);
      end;

      {Tob contenant le solde des écritures pointées}
      if X2 then begin
        TobP := TOB.Create('', TobGrid, -1);
        TobP.AddChampSupValeur('TYPE', TypeFlux);
        TobP.AddChampSupValeur('CODE', CodeFlux);
        TobP.AddChampSupValeur('NATURE' , na_Pointe);
        TobP.AddChampSupValeur('LIBELLE', RetourneNomLib(LibelleFlux, na_Pointe));
      end;
      {Ajout des colonnes dates}
      for n := 0 to NbColonne - 1 do begin
        TobG.AddChampSupValeur(RetourneCol(n), '');
        if X2 then
          TobP.AddChampSupValeur(RetourneCol(n), '');
      end;

      {15/04/05 : FQ 10240 : On stocke le contenu de la TobRappro}
      if X2 then begin
        {08/11/06 : Deux variables pour éviter de passer x fois dans le if HasSoldeReinit, notamment
         si affichage mensuel}
        Col := -1;
        IniOk := False;

        Mnt := GetSoldePointe(CodeFlux, sDate, Nature, not rbDateOpe.Checked) * txConv;
        {On stocke le solde initial pour le cas où il n'y aurait pas de rapprochement
         sur le premier jour de la période}
        TobP.SetDouble(GetTitreFromDate(DateDepart), Mnt);
        First := False;
        {TobDebug(TobRappro);}
        for n := 0 to TobRappro.Detail.Count - 1 do begin
          TR := TobRappro.Detail[n];
          {Si on n'est pas sur le compte courant}
          if TR.GetString('TE_GENERAL') <> CodeFlux then begin
            {Quand on a traité le compte, on sort ... il est inutile de continuer la boucle}
            if First then Break
                     else Continue;
          end;

          {On est sur le bon compte bancaire}
          Titre := GetTitreFromDate(TR.GetDateTime('TE_DATERAPPRO'));
          //SetlOpeDebAnnee
          {28/11/05 : FQ 10317 : Si on est sur une colonne avec solde de réinitialisation
           08/11/06 : FQ 10256 : Le DateToCol et préférable au n !!!}
          if HasSoldeReinit(COL_DATEDE + DateToCol(TR.GetDateTime('TE_DATERAPPRO')) - 1) and
            {08/11/06 : si on n'est pas encore passé ici ou si en mensuel, on est 12 mois plus tard}
            (not IniOk or (Col <> (COL_DATEDE + DateToCol(TR.GetDateTime('TE_DATERAPPRO')) - 1))) then begin
            {Si la date est un 01/01, GetSoldePointe renvoie le solde au matin du 01/01}
            if (Periodicite = tp_7) then
              Mnt := GetSoldePointe(CodeFlux, DateToStr(DebutAnnee(ColToDate(DateToCol(TR.GetDateTime('TE_DATERAPPRO')) + 1))), Nature, not rbDateOpe.Checked) * txConv
            else
              Mnt := GetSoldePointe(CodeFlux, DateToStr(DebutAnnee(TR.GetDateTime('TE_DATERAPPRO'))), Nature, not rbDateOpe.Checked) * txConv;

            TobP.SetDouble(Titre, Mnt);
            {08/11/06 : pour ne pas repasser ici, sauf en affichage mensuel si on est 12 mois plus tard}
            Iniok := True;
            Col   := COL_DATEDE + DateToCol(TR.GetDateTime('TE_DATERAPPRO')) - 1;
          end;

          Mnt := Mnt + TR.GetDouble('MONTANT') * txConv;
          TobP.SetDouble(Titre, Mnt);
          {on mémorise qu'il y a des écritures à cette date}
          ListeRouge.Add(UpperCase(Titre + '|' + CodeFlux + '|' + na_Pointe + '|'));

          First := True;
        end;
      end
    end;

    {-------------------------------------------------------------------------}
    procedure CalculerNonPointe(TT : TOB);
    {-------------------------------------------------------------------------}
    var
      TA, TB : TOB;
      n : Integer;
      D : string;
      M : Double;
    begin
      TA := TobGrid.FindFirst(['CODE', 'NATURE'], [TT.GetString('CODE'), na_Estime], False);
      TB := TobGrid.FindFirst(['CODE', 'NATURE'], [TT.GetString('CODE'), na_Pointe], False);
      if not Assigned(TA) or not Assigned(TB) then Exit;
      for n := 0 to NbColonne - 1 do begin
        D := RetourneCol(n);
        {Report éventuel des soldes précédents si la journée n'est pas mouvementée}
        if n > 0 then begin
          if TA.GetString(D) = '' then TA.SetDouble(D, TA.GetDouble(RetourneCol(n - 1)));
          if TB.GetString(D) = '' then TB.SetDouble(D, TB.GetDouble(RetourneCol(n - 1)));
        end;

        M := TA.GetDouble(D) - TB.GetDouble(D);
        TT.SetDouble(D, M);
      end;
    end;

    {-------------------------------------------------------------------------}
    procedure SommerBanque;
    {-------------------------------------------------------------------------}
    var
      N   : Integer;
      S   : string;
      Mnt : Double;
    begin
      MntTmp := 0.0;
      {Cumule chaque colonne date pour le type flux courant}
      for N := 0 to NbColonne - 1 do begin
        S := RetourneCol(n);

        {Calcul des totaux par banque ou par compte}
        Mnt := TobGrid.Somme(S, ['TYPE', 'NATURE'], [lBanque, na_Estime], False);
        if Mnt <> 0 then TobG.PutValue(S, Mnt)
                    else TobG.PutValue(S, MntTmp);
        MntTmp := Mnt;
      end;
    end;

    {-------------------------------------------------------------------------}
    procedure CreeTobTotal;
    {-------------------------------------------------------------------------}
    begin
      CreeTobDetail('+', lBanque, TraduireMemoire('Total banque'), '', False);
      SommerBanque;
    end;

    {12/05/04 : Ajout des comptes de la banque précédante non mouvementés sur la périodes
                mais appartenant à la sélection}
    {-------------------------------------------------------------------------}
    procedure GererComptesAbsents;
    {-------------------------------------------------------------------------}
    var
      Montant  : Double;
      n        : Integer;
    begin
      if not Assigned(ListeCpte)  then Exit;
      
      aTob := ListeCpte.FindFirst(['BANQUE'], [CodeBqe], True);
      while aTob <> nil do begin
        {Création de la tob qui va contenir les soldes du compte non mouvementé}
        CreeTobDetail(lBanque, aTob.GetString('COMPTE'), aTob.GetString('LIBELLE'), '', True);

        {Récupère le solde général du compte}
        Montant := CalculSoldeInit(aTob.GetString('COMPTE'), aTob.GetString('ADEVISE'), DateDepart); {***}
        Montant := Montant * TxConv;
        for n := 0 to NbColonne - 1 do begin
          {Si on est en début d'année, on récupère le solde forcé}
          if COL_DATEDE + n = ColRouge then
            Montant := CalculSoldeInit(aTob.GetString('COMPTE'), aTob.GetString('ADEVISE'), DateDepart) * TxConv;
          {JP 29/11/05 : FQ 10317}
          TobG.SetDouble(RetourneCol(n){DateToStr(DateDepart + n)}, Montant);
        end;

        {Récupère le solde des écritures pointées du compte}
        Montant := GetSoldePointe(aTob.GetString('COMPTE'), DateToStr(DateDepart), Nature, not rbDateOpe.Checked);
        Montant := Montant * TxConv;
        for n := 0 to NbColonne - 1 do begin
          {Si on est en début d'année, on récupère le solde forcé}
          if COL_DATEDE + n = ColRouge then
            Montant := TobG.GetDouble(RetourneCol(n));
          TobP.SetDouble(RetourneCol(n), Montant);
        end;

        {Création de la tob contenant le solde des écritures non pointées}
        CreeTobDetail(lBanque, aTob.GetString('COMPTE'), RetourneNomLib(aTob.GetString('LIBELLE'), na_Prevision), na_Prevision, False);
        {Calcul des soldes non pointés par soustraction du solde pointé au solde estimé}
        CalculerNonPointe(TobG);
        {On supprime ce compte, car il est possible que des banques n'est aucun compte mouvementé.
         A la fin de la boucle, on bouclera sur les comptes qui restent la Tob ListeCpte}
        SupprimeCptePresent(CodeBqe, aTob.GetString('COMPTE'));
        aTob := ListeCpte.FindNext(['BANQUE'], [CodeBqe], True);
      end;
    end;

    {FQ 10240 : Gestion des dates de rapprochement en option Date d'opération
    {-------------------------------------------------------------------------}
    function ATraiter(Dt : TDateTime) : Boolean;
    {-------------------------------------------------------------------------}
    begin
      Result := (Dt >= DateDepart) and (Dt <= DateFin) {FQ 10317 and (Dt <> DebutAnnee(Dt))};
    end;

var
  n, i      : Integer;
  General   : string;
  S         : string;
  Montant   : Double;
  MntReInit : Double;
  Operation : Double;
  LibCpte   : string;
  DevTmp    : string;
  Milles1Ok : Boolean; {FQ 10317}
  Milles2Ok : Boolean; {FQ 10317}
begin
  {On vide la liste zones ayant des opérations}
  ListeRouge.Clear;
  TobL := nil;
  {01/12/05 : FQ 10317}
  Milles1Ok := False;
  Milles2Ok := False;
  MntReInit := 0;
  
  {Définition de la colonne du millésime}
  ColRouge := -1;
  for n := 0 to NbColonne - 1 do
    if IsColAvecSoldeInit(n) then ColRouge := COL_DATEDE + n;{29/11/05 : FQ 10317}

  General := '';
  lBanque := '';
  Montant := 0;

  txConv := RetContreTaux(V_PGI.DateEntree, V_PGI.DevisePivot, Devise);

  for I := 0 to TobListe.Detail.Count - 1 do begin
    TobL := TobListe.Detail[I];

    S := TobL.GetString('LIBANQUE');
    {Changement de Banque}
    if S <> lBanque then begin
      if lBanque <> '' then begin
        {Création de la tob qui va contenir les soldes des écritures non pointées du compte précédent}
        CreeTobDetail(lBanque, General, RetourneNomLib(LibCpte, na_Prevision), na_Prevision, False);
        {Calcul des soldes non pointés par soustraction du solde pointé au solde estimé}
        CalculerNonPointe(TobG);

        {12/05/04 : On ajoute les éventuels comptes non mouvementés de la banque en cours}
        GererComptesAbsents;
        {Mise à jour du code de la banque pour GererComptesAbsents}
        CodeBqe := TobL.GetString('BQ_BANQUE');

        {Création de la tob contenant le solde total de la banque}
        CreeTobTotal;
        {Création d'une tob fille vide pour séparer les banques}
        CreeTobDetail('*', '', '', '', False);
      end;

      {Mise à jour du code de la banque pour GererComptesAbsents}
      CodeBqe := TobL.GetString('BQ_BANQUE');

      lBanque := S;
      {Création de la tob contenant le libellé de la banque}
      CreeTobDetail('', '', lBanque, '', False);
      General := '';
      {01/12/05 : FQ 10317}
      Milles1Ok := True;
      Milles2Ok := True;
    end;

    S := TobL.GetString('TE_GENERAL');

    {Nouveau compte}
    if General <> S then begin
      {12/05/04 : On supprime les comptes mouvementés}
      SupprimeCptePresent(CodeBqe, S);

      {Si on n'est pas sur la première ligne ...}
      if General <> '' then begin
        {Création de la tob qui va contenir les soldes des écritures non pointées du compte précédent}
        CreeTobDetail(lBanque, General, RetourneNomLib(LibCpte, na_Prevision), na_Prevision, False);
        {Calcul des soldes non pointés par soustraction du solde pointé au solde estimé}
        CalculerNonPointe(TobG);
      end;

      {Mise à jour du compte bancaire et de son libellé}
      General := S;
      LibCpte := TobL.GetString('LIBCPT');
      {28/11/05 : Il est préférable que la devise soit renseignée au moment du CalculSoldeInit}
      DevTmp := TobL.GetValue('ADEVISE');
      {On crée les tobs qui vont contenir les soldes du nouveau compte}
      CreeTobDetail(lBanque, General, LibCpte, '', True);

      {Récupère le solde du nouveau compte }
      Montant  := CalculSoldeInit(General, DevTmp, DateDepart) * txConv;

      {Affectation des soldes initiaux sur le premier jour pour le cas où il n'y aurait pas d'écriture
       le premier jour de la période de sélection}
      TobG.SetDouble(RetourneCol(0), Montant);

      {01/12/05 : FQ 10317}
      Milles1Ok := True;
      Milles2Ok := True;
    end;

    S := GetTitreFromDate(TobL.GetDateTime(sDateOpe));

    {Dans la fourchette d'affichage}
    if ATraiter(TobL.GetDateTime(sDateOpe)) then begin
      {FQ 10317 : dans le buget mensuel, il peut y avoir deux colonnes de réinitialisation}
      if (Periodicite = tp_30) and Milles2Ok then begin
        {S'il y a deux colonnes de réinitialisation}
        if ColRouge - 12 >= COL_DATEDE then begin
          if (DateToCol(TobL.GetDateTime(sDateOpe)) - 1 + COL_DATEDE) = ColRouge then begin
            Milles1Ok := True;
            Milles2Ok := False;
          end;
        end
        else
          Milles2Ok := False;
      end;

      {Si c'est la première fois que l'on est sur la colonne  de réinitialisation avec montant}
      if (Milles1Ok) and
         HasSoldeReinit(COL_DATEDE + DateToCol(TobL.GetDateTime(sDateOpe)) - 1) then begin
        {04/05/06 : FQ 10343 : Gestions des soldes initiaux : GetSoldeMillesime
                    Suite au problème rencontré avec THermocompact, je mets en dernier paramètre IsMntPivot
                    plutôt que True, car à la différence de la 6.53, txConv n'est plus calculé lorsque l'on
                    travaille avec TE_MONTANTDEV, or en mettant true, on demande un montant en euro alors que
                    tous les autres montants sont en devise. Le problème doit être le même dans ParPointage }
        if Periodicite = tp_1 then
          MntReInit := GetSoldeMillesime(General, DateToStr(DebutDeMois(TobL.GetDateTime(sDateOpe))), Nature, not rbDateOpe.Checked, False, IsMntPivot) * txConv
        else if Periodicite = tp_30 then
          MntReInit := GetSoldeMillesime(General, DateToStr(DebutDeMois(TobL.GetDateTime(sDateOpe))), Nature, not rbDateOpe.Checked, True, IsMntPivot) * txConv
        else if (Periodicite = tp_7) and (TobL.GetDateTime(sDateOpe) = DebutAnnee(TobL.GetDateTime(sDateOpe))) then
          MntReInit := GetSoldeMillesime(General, DateToStr(DebutDeMois(TobL.GetDateTime(sDateOpe))), Nature, not rbDateOpe.Checked, True, IsMntPivot) * txConv;

        {Si tp_7, tester la date d'opé avant de forcer montant et Milles1Ok}
        if (Periodicite in [tp_1, tp_30]) or
           {Le mode hebdomadaire a la particularité de pouvoir avoir le 01/01 au milieu de la période d'une colonne}
           ((Periodicite = tp_7) and (TobL.GetDateTime(sDateOpe) = DebutAnnee(TobL.GetDateTime(sDateOpe)))) then begin
          Montant := MntReInit;
          Milles1Ok := False;
        end;

        {En affichage quotidien, on force le solde au soir du 01/01 et on passe aux écritures du 02/01}
        if Periodicite = tp_1 then begin
          TobG.SetDouble(S, Montant);
          ListeRouge.Add(UpperCase(S + '|' + General + '|' + na_Estime + '|'));
          Continue;
        end;
      end;

      {En afficchage quotidien quotidien, sur le 01/01 et qu'on a déjà récupéré le solde du 01/01 au soir ...}
      if (Periodicite = tp_1) and not Milles1Ok and
        (TobL.GetDateTime(sDateOpe) = DebutAnnee(TobL.GetDateTime(sDateOpe))) then Continue;

      Operation := TobL.GetDouble('MONTANT') * txConv;

      {Si premier passage sur colonne de réinitialisation, il faut :
      Montant := MntReInit; MntReInit = 0.001
      Peut-être faire le test sur TobL.GetDateTime(sDateOpe) et non sur HasSoldeReinit
      => ajouter un boolean
      => dans les 3 périodicités récupérer GetSoldeMillesime au matin}
      Montant := Montant + Operation;

      TobG.SetDouble(S, Montant);

      ListeRouge.Add(UpperCase(S + '|' + General + '|' + na_Estime + '|'));

      {15/04/05 : FQ 10240 : si la date de rappro est supérieure à "sDateOpe" ou égale à iDate1900,
                  c'est que l'écriture n'est pas (ou pas encore)  pointée}
      if (TobL.GetDateTime(sDateOpe) < TobL.GetDateTime('TE_DATERAPPRO')) or
         (TobL.GetDateTime('TE_DATERAPPRO') = iDate1900) then
        ListeRouge.Add(UpperCase(S + '|' + General + '|' + na_Prevision + '|'));

    end;
  end;{Boucle For}

  if lBanque <> '' then begin
    {Création de la tob qui va contenir les soldes des écritures non pointées du compte précédent}
    CreeTobDetail(lBanque, General, RetourneNomLib(LibCpte, na_Prevision), na_Prevision, False);
    {Calcul des soldes non pointés par soustraction du solde pointé au solde estimé}
    CalculerNonPointe(TobG);

    {On ajoute les éventuels comptes non mouvementés de la banque en cours}
    GererComptesAbsents;
    {Mise à jour du code de la banque pour GererComptesAbsents}
    CodeBqe := TobL.GetString('BQ_BANQUE');

    {Pour la dernière banque}
    CreeTobTotal;
  end;

  if Assigned(ListeCpte) then
    {12/05/04 : ajout des banques n'ayant que des comptes non mouvementés}
    while ListeCpte.Detail.Count > 0 do begin
      CodeBqe := ListeCpte.Detail[0].GetString('BANQUE');
      lBanque := ListeCpte.Detail[0].GetString('LIBBANQUE');

      {Création d'une tob fille vide pour séparer les banques}
      CreeTobDetail('*', '', '', '', False);
      {Création de la tob contenant le libellé de la banque}
      CreeTobDetail('', '', ListeCpte.Detail[0].GetString('LIBBANQUE'), '', False);
      {On ajoute les comptes de la banque}
      GererComptesAbsents;
      {Pour la dernière banque}
      CreeTobTotal;
    end;

  {Totaux généraux}
  CreeTobDetail('-', '', TraduireMemoire('Solde général'), '', False);
  MntTmp := 0.0;
  for n := 0 to NbColonne - 1 do begin
    S := RetourneCol(n);
    Montant := TobGrid.Somme(S, ['TYPE'], ['+'], False);
    if Montant <> 0 then TobG.PutValue(S, Montant)
                    else TobG.PutValue(S, MntTmp);
    if Montant <> 0 then MntTmp := Montant;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.ParSolde;
{---------------------------------------------------------------------------------------}
var
  aTob,
  TobL,
  TobG : TOB;
  lBanque  : string; {20/11/03}
  MntTmp   : Double; {20/11/03}
  CodeBqe  : string; {12/05/04}
  txConv   : Double;
  {LÉGENDE DES TYPES : "LibBanque" = Ligne contenant le solde d'un compte bancaire
                       ""          = Ligne contenant le libellé d'une banque
                       "*"         = Ligne de séparation entre deux banques
                       "+"         = Solde total d'une banque
                       "-"         = Solde général
   LÉGENDE DES NATURES : "P", "R", "S" = Cumuls des opérations par nature
                         "T"           = Ligne de soldes bancaires}

    {------------------------------------------------------------------}
    procedure CreeTobDetail(TypeFlux, CodeFlux, LibelleFlux, Nat : string);
    {------------------------------------------------------------------}
    var
      n : Integer;
    begin
      TobG := TOB.Create('', TobGrid, -1);
      TobG.AddChampSupValeur('TYPE', TypeFlux);
      TobG.AddChampSupValeur('CODE', CodeFlux);
      TobG.AddChampSupValeur('NATURE'  , Nat);
      TobG.AddChampSupValeur('LIBELLE', LibelleFlux);
      for n := 0 to NbColonne - 1 do {Ajout des colonnes dates}
        TobG.AddChampSupValeur(RetourneCol(n), '');
    end;

    {JP 19/03/04 : Le traitement de SommerBanque a été sorti pour pouvoir être
                   exécuté plus tard
    {-------------------------------------------------------------------------}
    procedure CreeTobTotal;
    {-------------------------------------------------------------------------}
    begin
      CreeTobDetail('+', lBanque{''}, TraduireMemoire('Total banque'), '');
    end;

    {-------------------------------------------------------------------------}
    procedure SommerBanque(BqeOk : Boolean);
    {-------------------------------------------------------------------------}
    var
      N   : Integer;
      S   : string;
      Mnt : Double;
      UsD : string;
    begin
      MntTmp := 0.0;
      {Cumule chaque colonne date pour le type flux courant}
      for N := 0 to NbColonne - 1 do begin
        S := RetourneCol(n);

        {Gestion du solde forcé au premier janvier, dans la mesure où l'on n'est pas sur un total banque}
        if IsColAvecSoldeInit(n) and not BqeOk then begin {JP 28/11/05 : FQ 10317 : IsColAvecSoldeInit}
          if Periodicite = tp_7 then UsD := USDateTime(DebutAnnee(ColToDate(n + 1)))
                                else UsD := USDateTime(DebutAnnee(ColToDate(n)));

          if ExisteSQL('SELECT TE_GENERAL FROM TRECRITURE WHERE TE_DATEVALEUR = "' + UsD +
                       '" AND TE_GENERAL = "' + lBanque + '" AND TE_QUALIFORIGINE = "' + CODEREINIT + '"') then begin
            {JP 28/11/05 : FQ 10317 : chaque type d'affichage demande une gestion spécifique du solde d'initialisation
                           Quotidien : on récupère le solde du 01/01 au soir et on le met dans la colonne du 01/01
                           Hebdomadaire : récupération du solde du 01/01 au matin (ColToDate(n + 1) pour être sûr de
                                          prendre la bonne année auquel on ajoute les opérations comprises entre le
                                          01/01 et la fin de la semaine contenant le 01/01 (GetlOpeDebAnnee)
                           Mensuel : récupération du solde au 01/01 au matin auquel on rajoute toutes les opérations
                                     du mois de janvier}
            {04/05/06 : FQ 10343 : Gestions des soldes initiaux : GetSoldeMillesime
                        Suite au problème rencontré avec THermocompact, je mets en dernier paramètre IsMntPivot
                        plutôt que True, car à la différence de la 6.53, txConv n'est plus calculé lorsque l'on
                        travaille avec TE_MONTANTDEV, or en mettant true, on demande un montant en euro alors que
                        tous les autres montants sont en devise. Le problème doit être le même dans ParPointage }
            if Periodicite = tp_1 then
              Mnt := GetSoldeMillesime(lBanque, DateToStr(ColToDate(n)), Nature, not rbDateOpe.Checked, False, IsMntPivot) * txConv
            else if Periodicite = tp_7 then
              Mnt := GetSoldeMillesime(lBanque, DateToStr(ColToDate(n + 1)), Nature, not rbDateOpe.Checked, True, IsMntPivot) * txConv
            else
              Mnt := GetSoldeMillesime(lBanque, DateToStr(ColToDate(n)), Nature, not rbDateOpe.Checked, True, IsMntPivot) * txConv;

            {Il y a un solde forcé}
            if StrFPoint(Mnt) <> '0.001' then begin
              {Pour mettre le solde en rouge dans la grille ***}
              ListeRouge.Add(TobG.GetString('CODE'));

              {JP 28/11/05 : FQ 10317 : GetlOpeDebAnnee renvoie 0 sauf en affichage hebdomadaire}
              Mnt := Mnt + GetlOpeDebAnnee(lBanque);
              {Sauf en affichage mensuel, on force le solde ainsi calculé}
              if Periodicite <> tp_30 then TobG.SetDouble(S, Mnt);
              MntTmp := Mnt;
              {JP 28/11/05 : FQ 10317 : en affichage mensuel, on ajoute au solde récupéré toutes les opérations du mois}
              if Periodicite <> tp_30 then Continue;
            end;
          end;
        end;

        {Calcul des totaux par banque ou par compte}
        if BqeOk then Mnt := TobGrid.Somme(S, ['TYPE', 'NATURE'], [lBanque, na_Total], False) {Solde par banque}
                 else Mnt := MntTmp + TobGrid.Somme(S, ['CODE'], [lBanque], False); {Solde par compte}
        if Mnt <> 0 then TobG.SetDouble(S, Mnt)
                    else TobG.SetDouble(S, MntTmp);
        MntTmp := Mnt;
      end;
    end;

    {12/05/04 : Ajout des comptes de la banque précédante non mouvementés sur la périodes
                mais appartenant à la sélection}
    {-------------------------------------------------------------------------}
    procedure GererComptesAbsents;
    {-------------------------------------------------------------------------}
    begin
      if not Assigned(ListeCpte) then Exit;

      aTob := ListeCpte.FindFirst(['BANQUE'], [CodeBqe], True);
      while aTob <> nil do begin
        {Création de la tob qui va contenir les soldes du compte non mouvementé}
        CreeTobDetail(lBanque, aTob.GetString('COMPTE'), RetourneNomLib(aTob.GetString('LIBELLE'), na_Total), na_Total);
        {28/11/05 : FQ 10317 : Récupération et affectation du solde de départ}
        GereSoldeDepart(TobG, aTob.GetString('COMPTE'), aTob.GetString('ADEVISE'), txConv);
        {On supprime ce compte, car il est possible que des banques n'est aucun compte mouvementé.
         A la fin de la boucle, on bouclera sur les comptes qui restent la Tob ListeCpte}
        SupprimeCptePresent(CodeBqe, aTob.GetString('COMPTE'));
        aTob := ListeCpte.FindNext(['BANQUE'], [CodeBqe], True);
      end;
    end;

var
  I, N    : Integer;
  General : string;
  S, Code : string;
  Montant : Double;
  OldNat  : string; {16/04/04}
  LibCpte : string;
  DevTmp  : string;
begin
  {On vide la liste des soldes vides}
  ListeRouge.Clear;
  ColRouge := -1;

  Code    := '';
  General := '';
  lBanque := '';
  OldNat  := '';
  TobL    := nil;

  {16/01/06 : FQ 10328 : On ne gère le taux de conversion que si on travaille avec TE_MONTANT}
  txConv := 1;
  if ChpMontant = 'TE_MONTANT' then
    txConv := RetContreTaux(V_PGI.DateEntree, V_PGI.DevisePivot, Devise);

  for I := 0 to TobListe.Detail.Count - 1 do begin
    TobL := TobListe.Detail[I];

    S := TobL.GetString('LIBANQUE');
    {Changement de Banque}
    if S <> lBanque then begin
      if lBanque <> '' then begin
        {Création de la tob qui va contenir les soldes du compte précédent}
        CreeTobDetail(lBanque, General, RetourneNomLib(LibCpte, na_Total), na_Total);
        {28/12/05 : FQ 10317 : Gestion du solde au premier jour de la période de traitement}
        GereSoldeDepart(TobG, General, DevTmp, txConv);

        {12/05/04 : On ajoute les éventuels comptes non mouvementés de la banque en cours}
        GererComptesAbsents;
        {Mise à jour du code de la banque pour GererComptesAbsents}
        CodeBqe := TobL.GetString('BQ_BANQUE');

        {Création de la tob contenant le solde total de la banque}
        CreeTobTotal;
        {Création d'une tob fille vide pour séparer les banques}
        CreeTobDetail('*', '', '', '');
      end;

      {Mise à jour du code de la banque pour GererComptesAbsents}
      CodeBqe := TobL.GetString('BQ_BANQUE');

      lBanque := S;
      {Création de la tob contenant le libellé de la banque}
      CreeTobDetail('', '', lBanque, '');
      General := '';
    end;

    S := TobL.GetString('TE_GENERAL');

    {Nouveau compte}
    if General <> S then begin
      {12/05/04 : On supprime les comptes mouvementés}
      SupprimeCptePresent(CodeBqe, S);

      {Si on n'est pas sur la première ligne ...}
      if General <> '' then begin
        {... on crée la tob qui va contenir les soldes du compte précédent}
        CreeTobDetail(lBanque, General, RetourneNomLib(LibCpte, na_Total), na_Total);
        {28/12/05 : FQ 10317 : Gestion du solde au premier jour de la période de traitement}
        GereSoldeDepart(TobG, General, DevTmp, txConv);
      end;

      {Mise à jour de la variable sur la nature}
      OldNat := TobL.GetString('TE_NATURE');
      {Mise à jour du compte bancaire et de son libellé}
      General := S;

      {06/12/05 : remonté et non plus à la fin du For}
      DevTmp := TobL.GetString('ADEVISE'); {***}

      LibCpte := TobL.GetString('LIBCPT');
      {Création de la ligne de cumul des opérations pour le nouveau compte et la nature en cours}
      CreeTobDetail(lBanque, S, RetourneNomLib(LibCpte, OldNat), OldNat);
  //    {28/12/05 : FQ 10317 : Gestion du solde au premier jour de la période de traitement}
//      GereSoldeDepart(TobG, General, DevTmp, txConv);
    end;

    S := TobL.GetString('TE_NATURE');
    if S <> OldNat then begin
      OldNat := S;
      {Création de la ligne de cumul des opérations pour le nouveau compte et la nature en cours}
      CreeTobDetail(lBanque, General, RetourneNomLib(LibCpte, OldNat), OldNat);
    end;

    S := GetTitreFromDate(TobL.GetDateTime(sDateOpe));
    {Dans la fourchette d'affichage}
    if (TobL.GetDateTime(sDateOpe) >= DateDepart) and (TobL.GetDateTime(sDateOpe) <= DateFin) then begin
      Montant := TobL.GetDouble('MONTANT') * txConv;
      TobG.SetDouble(S, Montant + TobG.GetDouble(S));
      SetlOpeDebAnnee(General, Montant, TobL.GetDateTime(sDateOpe));
    end;
  end;{Boucle For}

  if lBanque <> '' then begin
    {On crée la tob qui va contenir les soldes du dernier compte.}
    CreeTobDetail(lBanque, General, RetourneNomLib(LibCpte, na_Total), na_Total);
    {28/12/05 : FQ 10317 : Gestion du solde au premier jour de la période de traitement}
    GereSoldeDepart(TobG, General, DevTmp, txConv);
    {12/05/04 : On ajoute les éventuels comptes non mouvementés de la banque en cours}
    GererComptesAbsents;
    {Mise à jour du code de la banque pour GererComptesAbsents}
    CodeBqe := TobL.GetString('BQ_BANQUE');

    {Pour la dernière banque}
    CreeTobTotal;
  end;

  if Assigned(ListeCpte)  then
    {12/05/04 : ajout des banques n'ayant que des comptes non mouvementés}
    while ListeCpte.Detail.Count > 0 do begin
      CodeBqe := ListeCpte.Detail[0].GetString('BANQUE');
      lBanque := ListeCpte.Detail[0].GetString('LIBBANQUE');

      {Création d'une tob fille vide pour séparer les banques}
      CreeTobDetail('*', '', '', '');
      {Création de la tob contenant le libellé de la banque}
      CreeTobDetail('', '', ListeCpte.Detail[0].GetString('LIBBANQUE'), '');
      {On ajoute les comptes de la banque}
      GererComptesAbsents;
      {Pour la dernière banque}
      CreeTobTotal;
    end;

  {16/04/04 : Gestion des soldes des comptes, on part du solde initial auquel on ajoute les opérations du jours}
  for I := 1 to TobGrid.Detail.Count-1 do begin
    TobG := TobGrid.Detail[I];
    if (TobG.GetString('NATURE') <> na_Total) then Continue;
    lBanque := TobG.GetString('CODE');
    SommerBanque(False);
  end;

  {JP 19/03/04 : Totaux par banque. Auparavant cette opération était réalisée dans CreeTobTotal,
                 donc avant le traitement ci-dessus. On pouvait donc se retrouver dans la situation
                 suivante : Cpt1 ----  15000    ----     2000
                            Cpt2 ----  10000    -2000    -----
                    Total banque 2500  27500    -2000    2000
                    au lieu de   2500  27500    25500   27500}
  for n := 0 to TobGrid.Detail.Count - 1 do begin
    TobG := TobGrid.Detail[n];
    lBanque := TobG.GetString('CODE');
    if TobG.GetString('TYPE') = '+' then
      SommerBanque(True);
  end;

  {Totaux généraux}
  CreeTobDetail('-', '', TraduireMemoire('Solde général'), '');
  MntTmp := 0.0;
  for n := 0 to NbColonne - 1 do begin
    {Définition de la colonne du millésime}
    if IsColAvecSoldeInit(n) then ColRouge := COL_DATEDE + n; {28/11/05 : FQ 10317}
    S := RetourneCol(n);
    Montant := TobGrid.Somme(S, ['TYPE'], ['+'], False);
    if Montant <> 0 then TobG.SetDouble(S, Montant)
                    else TobG.SetDouble(S, MntTmp);
    if Montant <> 0 then MntTmp := Montant;
  end;
end;

{Est utilisé pour le suivi des soldes}
{---------------------------------------------------------------------------------------}
function FICTAB.GetRequete(Cl1 : string) : string;
{---------------------------------------------------------------------------------------}
var
  TmpREI : string;
  ch     : string; {23/05/05 : FQ 10253}
  whConf : string;
begin
  TmpREI := '';
  {S'il y a un filtre sur les natures, il ne faut pas qu'il s'applique sur les écritures de réinitialisation}
  if Trim(Nature) <> '' then begin
    TmpREI := Copy(Nature, Pos('AND', Nature) + 4, Length(Nature));
    TmpREI := ' AND ((' + TmpREI + ') OR TE_QUALIFORIGINE = "' + CODEREINIT + '") '
  end;

  whConf := GetWhereConfidentialite;

  ch := GetWhereDossier;
  if Ch <> '' then
    TmpREI := TmpREI + ' AND ' + Ch;

  {Cl1 : filtre sur les dates}
  if TypeDesc = dsSolde then                                                   {16/01/06 : FQ 10328}
    Result := 'SELECT TE_GENERAL, ' + sDateOpe + ', MAX(TE_DEVISE) ADEVISE, SUM(' + ChpMontant + ') MONTANT,' +
          ' TE_NATURE, PQ_LIBELLE LIBANQUE, BQ_LIBELLE LIBCPT, BQ_BANQUE FROM TRECRITURE' +
          ' LEFT JOIN BANQUECP ON TE_GENERAL = BQ_CODE' +
          ' LEFT JOIN BANQUES ON BQ_BANQUE = PQ_BANQUE' + Cl1 + TmpRei + whConf +
          ' GROUP BY TE_NATURE, ' + sDateOpe + ', TE_GENERAL, BQ_LIBELLE, PQ_LIBELLE, BQ_BANQUE' +
          ' ORDER BY PQ_LIBELLE, TE_GENERAL, TE_NATURE, ' + sDateOpe

  else if TypeDesc = dsGroupe then begin
    Result := 'SELECT BQ_NATURECPTE, BQ_NODOSSIER, ' + sDateOpe + ', SUM(TE_MONTANT) MONTANT,' +
          ' TE_NATURE, BQ_BANQUE, DOS_LIBELLE, DOS_NOMBASE FROM TRECRITURE' +
          ' LEFT JOIN BANQUECP ON TE_GENERAL = BQ_CODE' +
          ' LEFT JOIN DOSSIER ON DOS_NODOSSIER = BQ_NODOSSIER ' + Cl1 + TmpRei + whConf +
          ' GROUP BY BQ_NODOSSIER, DOS_LIBELLE, DOS_NOMBASE, BQ_NATURECPTE, ' + sDateOpe + ', TE_NATURE, BQ_BANQUE' +
          ' ORDER BY BQ_NODOSSIER, BQ_NATURECPTE, ' + sDateOpe
  end

  else begin
    {FQ 10253 : Pour éviter de se retrouver deux fois avec TE_DATERAPPRO dans la requête,
               ce qui ne passe pas sous Oracle}
    if sDateOpe = 'TE_DATERAPPRO' then ch := ''
                                  else ch := sDateOpe + ',';
    {24/08/06 : sur le suivi bancaire, on ne travaille que sur les comptes bancaires}
    TmpRei := TmpRei + 'AND ' + BQCLAUSEWHERE;

                                                                     {16/01/06 : FQ 10328}
    Result := 'SELECT TE_GENERAL, ' + ch + ' MAX(TE_DEVISE) ADEVISE, SUM(' + ChpMontant + ') MONTANT,' +
          ' TE_NATURE, PQ_LIBELLE LIBANQUE, BQ_LIBELLE LIBCPT, BQ_BANQUE, TE_DATERAPPRO FROM TRECRITURE' +
          ' LEFT JOIN BANQUECP ON TE_GENERAL = BQ_CODE' +
          ' LEFT JOIN BANQUES ON BQ_BANQUE = PQ_BANQUE' + Cl1 + TmpRei + whConf +
          ' GROUP BY TE_NATURE, ' + ch + ' TE_GENERAL, TE_DATERAPPRO, BQ_LIBELLE, PQ_LIBELLE, BQ_BANQUE' +
          {' GROUP BY TE_NATURE, ' + sDateOpe + ', TE_GENERAL, TE_DATERAPPRO, BQ_LIBELLE, PQ_LIBELLE, BQ_BANQUE' +
          {08/11/06 : FQ 10256 : Ajout de la clef valeur pour que sur les écriture rapprochées le 01/01,
                      l'initialisation soit la première du jour}
          ' ORDER BY PQ_LIBELLE, TE_GENERAL, ' + sDateOpe {+ ', TE_CLEVALEUR'};
  end;
end;

{Fonction utilisée lors de la constitution des tob d'affichage
{---------------------------------------------------------------------------------------}
function FICTAB.RetourneNomLib(Cpte, Nat : string) : string;
{---------------------------------------------------------------------------------------}
begin
       if Nat = na_Realise    then Result := Cpte + ' (Réal.)'
  else if Nat = na_Prevision  then Result := Cpte + ' (Prév.)'
  else if Nat = na_Simulation then Result := Cpte + ' (Simu.)'
  else if Nat = na_Total      then Result := Cpte + ' (Solde)'
  else if Nat = na_Pointe     then Result := Cpte + ' (Rapp.)' {13/10/04 : solde bancaire}
  else if Nat = na_Estime     then Result := Cpte + ' (Esti.)' {13/10/04 : solde bancaire}
  else Result := Cpte;
end;


{---------------------------------------------------------------------------------------}
function FICTAB.GetMultiNature : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := (Trim(Nature) = '') or (Pos('<<', Nature) > 0) or {Pas de filtre sur les natures}
            ((Pos('"' + na_Prevision  + '"', Nature) > 0) and (Pos('"' + na_Realise    + '"', Nature) > 0)) or
            ((Pos('"' + na_Prevision  + '"', Nature) > 0) and (Pos('"' + na_Simulation + '"', Nature) > 0)) or
            ((Pos('"' + na_Simulation + '"', Nature) > 0) and (Pos('"' + na_Realise    + '"', Nature) > 0));
end;

{---------------------------------------------------------------------------------------}
function FICTAB.CalculSoldeInit(Gen, Dev : string; dt : TDateTime) : Double;
{---------------------------------------------------------------------------------------}
var
  S : string;
  D : TDateTime;
begin
  {Récupère le solde du compte précédent}
  if (Nature = '') then begin
    if rbDateOpe.Checked then
      Result := GetSolde(Gen, DateToStr(dt - 1), S)
    else
      Result := GetSoldeValeur(Gen, DateToStr(dt - 1), S);

    {Commentaire du 28/11/05 : Cela demanderait à être repensé : si on demande une fiche de suivi au 01/01
                               et une au 02/01, le soldes sont différents car la date de départ pour la
                               conversion en devise n'est pas la même ....}
    if dt = DebutAnnee(dt) then D := Dt
                           else D := V_PGI.DateEntree;
    if MntPivot then {16/01/06 : FQ 10328}
      Result := Result * RetPariteEuro(Dev, D); {***}
  end
  else
    {À la différence de GetSolde(Valeur), GetSoldeInit exécute une requête avec une inégalité stricte => pas de -1
     16/01/06 : FQ 10328 : On ne récupère plus nécessairement le solde en devise pivot}
    Result := GetSoldeInit(Gen, DateToStr(dt), Nature, not rbDateOpe.Checked, MntPivot); {***}
end;

{Supprime de la tob liste les comptes / OPCVM mouvementés car ils sont traités par la requête principale
{---------------------------------------------------------------------------------------}
procedure FICTAB.SupprimeCptePresent(Bqe, Cpte : string);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  if not Assigned(ListeCpte) then Exit;
  T := ListeCpte.FindFirst(['BANQUE', 'COMPTE'], [Bqe, Cpte], True);
  if T <> nil then FreeAndNil(T);
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.SupprimeCptePresentMS(Cpte : string);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  if not Assigned(ListeCpte) then Exit;
  T := ListeCpte.FindFirst(['COMPTE'], [Cpte], True);
  if T <> nil then FreeAndNil(T);
end;

{Permet d'ajouter dans la TobListe les comptes qui n'y figureraient pas car non mouvementés
 dans le but d'obtenir des soldes généraux cohérents (27/12/04 : même principe pour les OPCVM)
{---------------------------------------------------------------------------------------}
procedure FICTAB.CreerListeCompte(ParBanque : Boolean = False);
{---------------------------------------------------------------------------------------}
var
  W, S, C : string;
  whConf  : string;
begin
  {Ne concerne que les écrans travaillant par soldes et non par flux}
  if not (TypeDesc in [dsSolde, dsDetail, dsTreso, dsBancaire, dsGroupe]) then Exit;
  {Création / libération de la liste}
  if not Assigned(ListeCpte) then ListeCpte := TOB.Create('****', nil, -1)
                             else ListeCpte.ClearDetail;

  C := WhereBqCpt;
  W := GetWhereDossier('BQ_NODOSSIER');
  if C <> '' then begin
    if W <> '' then W := C + ' AND ' + W
               else W := C;
  end;

  whConf := GetWhereConfidentialite;

  if IsTresoMultiSoc and (TypeDesc in [dsSolde, dsDetail, dsGroupe]) then begin
    if TypeDesc = dsGroupe then
      {Pas de notion de compte dans le suivi de groupe}
      S := 'SELECT DISTINCT BQ_NODOSSIER CODDOSSIER, BQ_NATURECPTE, "" BANQUE, "" COMPTE, '
    else
      S := 'SELECT BQ_NODOSSIER CODDOSSIER, BQ_BANQUE BANQUE, BQ_LIBELLE LIBELLE, BQ_CODE COMPTE, BQ_DEVISE ADEVISE, PQ_LIBELLE LIBBANQUE, ';

    {$IFDEF TRCONF}
    if W = '' then W := ObjConf.GetWhereSQLConf(tyc_Banque)
              else W := W + whConf;
    {$ENDIF TRCONF}

    S := S + 'DOS_LIBELLE LIBDOSSIER, BQ_NATURECPTE FROM BANQUECP ' +
             'LEFT JOIN BANQUES ON PQ_BANQUE = BQ_BANQUE ' +
             'LEFT JOIN DOSSIER ON DOS_NODOSSIER = BQ_NODOSSIER ' +
             'WHERE ' + W + ' ORDER BY CODDOSSIER';
  end
  else begin
    if ParBanque then
      {$IFDEF TRCONF}
      S := 'SELECT DISTINCT PQ_BANQUE BANQUE, PQ_LIBELLE LIBBANQUE, "" LIBELLE, "" COMPTE, "" ADEVISE FROM BANQUES ' +
           'LEFT JOIN BANQUECP ON BQ_BANQUE = PQ_BANQUE WHERE ' + WhereBqCpt + WhConf + ' ORDER BY PQ_BANQUE'
      {$ELSE}
      S := 'SELECT PQ_BANQUE BANQUE, PQ_LIBELLE LIBBANQUE, "" LIBELLE, "" COMPTE, "" ADEVISE FROM BANQUES WHERE ' + WhereBqCpt + ' ORDER BY PQ_BANQUE'
      {$ENDIF TRCONF}
    else begin
      if W <> '' then W := 'AND ' + W;
      W := W + WhConf;
      if IsTresoMultiSoc and (TypeDesc = dsBancaire) then W := W + 'AND BQ_BANQUE NOT IN ("' + CODECOURANTS + '", "' + CODETITRES + '")';

      S := 'SELECT BQ_BANQUE BANQUE, BQ_LIBELLE LIBELLE, BQ_CODE COMPTE, BQ_DEVISE ADEVISE, PQ_LIBELLE LIBBANQUE ' +
           'FROM BANQUECP, BANQUES WHERE PQ_BANQUE = BQ_BANQUE ' + W + ' ORDER BY BQ_BANQUE';
    end;
  end;
  ListeCpte.LoadDetailFromSQL(S);
end;

{Récupère les clauses where sur les comptes et les banques pour la fonction CreerListeCompte
{---------------------------------------------------------------------------------------}
function FICTAB.GetWhereBqCpt : string;
{---------------------------------------------------------------------------------------}
begin
  Result := '';
end;

{Filtre la combo des dossiers sur les seuls dossiers Tréso du regroupement}
{---------------------------------------------------------------------------------------}
procedure FICTAB.FiltreDossierTreso;
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  if IsAvecRegroupement and IsTresoMultiSoc then begin
    s := 'DOS_NOMBASE ' + FiltreBaseTreso(False);
    if (TypeDesc in [dsSolde, dsGroupe, dsBancaire, dsTreso]) and
       {Si détail en affichage en solde, alors TypeDesc = dsSolde, mais il n'y a pas de combo REGROUPEMENT}
       (Assigned(GetControl('REGROUPEMENT'))) then
      THMultiValComboBox(GetControl('REGROUPEMENT')).Plus := s
    else if Assigned(GetControl('REGROUPEMENT')) then
      THValComboBox(GetControl('REGROUPEMENT')).Plus := s;
  end;
  {Pour filtrer les comptes bancaires}
  NoDossierChange(GetControl('REGROUPEMENT'));
end;

{Initialisation des controls si TheTob / LaTof.LaTob contient leur valeur
{---------------------------------------------------------------------------------------}
procedure FICTAB.InitControls;
{---------------------------------------------------------------------------------------}
begin
//(dsPeriodique, dsMensuel, dsSolde, dsTreso, dsCommission, dsBancaire, dsOPCVM, dsGroupe);
  {Si on est passé par TheTob et que l'on n'est pas sur le détail de suivi qui a son propre
   mode d'appel, utilisé depuis les  Virements et la fiche de suivi}
  if Assigned(LaTOB) and not(TypeDesc in [dsDetail, dsControle]) then begin
    TFMul(Ecran).FiltreDisabled := False;
    if TypeDesc in [dsSolde, dsBancaire, dsGroupe] then
      SetControlText('BANQUE', LaTob.GetString('BANQUE'));

    if IsAvecRegroupement then SetControlText('REGROUPEMENT', LaTob.GetString('REGROUP'));

    SetControlText('MCNATURE', LaTob.GetString('NATURE'));

    if not (TypeDesc in [dsCommission, dsOPCVM]) then begin
      if TypeDesc <> dsBancaire then begin
        if LaTob.GetString('DATEOPE') = 'TE_DATECOMPTABLE' then SetControlChecked('DATEOPE', True)
                                                           else SetControlChecked('DATEVAL', True);
      end;
    end;

    if not (TypeDesc in [dsPeriodique, dsMensuel]) then begin
      case LaTob.GetInteger('PERIODICITE') of
        0, 2 : SetControlChecked('RBQUOTI', True);
        1    : SetControlChecked('RBHEBDO', True);
        3    : SetControlChecked('RBMENSU', True);
      end;

      SetControlText('INTERVAL', LaTob.GetString('INTERVAL'));
      SetControlText('DATEDE'  , LaTob.GetString('DATEDEB'));

      if not (TypeDesc in [dsBancaire, dsControle]) then
        MajComboDevise(LaTob.GetString('DEVISE'));
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.OnFormResize(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  {Exécute le code d'origine}
  FFormResize(Sender);
  {Pour remettre les colonnes telles que définies dans le OnArgument}
  for n := COL_DATEDE to Grid.ColCount - 1 do
    Grid.ColWidths[n] := LargeurCol;
  {Maintenant, on peut redessiner la grille}
  Grid.Invalidate;
end;

{08/03/05 : FQ 10217 : renvoie la date "courante" en fonction de la colonne active
 20/04/05 : FQ 10237 : Nouvelle gestion des dates avec la gestion de la périodicité
{---------------------------------------------------------------------------------------}
function FICTAB.GetCurDate : string;
{---------------------------------------------------------------------------------------}
begin
  Result := DateToStr(V_PGI.DateEntree);
  if TypeDesc in [dsTreso, dsDetail, dsSolde, dsBancaire, dsCommission, dsGroupe] then begin
    if rbDateOpe.Checked or (TypeDesc = dsCommission) then
      Result := DateToStr(ColToDate(Grid.Col - COL_DATEDE));
  end;
end;

{10/10/06 : FQ 10376 : affiche le jour devant la date}
{---------------------------------------------------------------------------------------}
function FICTAB.GetDateStr(aDate : TDateTime) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := FormatDateTime('ddd dd/mm/yy', aDate);
end;

{18/04/05 : FQ 10237 : Getter pour connaître la périodicité demandée
{---------------------------------------------------------------------------------------}
function FICTAB.GetPeriodicite : TypePeriod;
{---------------------------------------------------------------------------------------}
begin
  if TypeDesc in [dsDetail, dsSolde, dsTreso, dsCommission, dsBancaire, dsOPCVM, dsGroupe] then begin
    if rbMensu.Checked then
      Result := tp_30
    else if not (TypeDesc in  [dsTreso]) and rbHebdo.Checked then
      Result := tp_7
    else if TypeDesc = dsCommission then Result := tp_15
                                    else Result := tp_1;
  end
  else if TypeDesc = dsMensuel then
    Result := tp_30
  else if TypeDesc = dsPeriodique then
    Result := tp_15
  else
    Result := tp_1;
end;

{---------------------------------------------------------------------------------------}
function FICTAB.GetDateDepart : TDateTime;
{---------------------------------------------------------------------------------------}
begin
  if edDepart.Text = '/  /' then DateDepart := V_PGI.DateEntree - 1;
  if Periodicite = tp_30 then
    Result := DebutDeMois(StrToDate(edDepart.Text))
  else
    Result := StrToDate(edDepart.Text);
end;

{---------------------------------------------------------------------------------------}
function FICTAB.GetNbColonne : Integer;
{---------------------------------------------------------------------------------------}
begin
  Result := seInterval.Value;
  if Result <= 0 then begin
    Result := 24;
    NbColonne := Result;
  end;
end;

{18/04/05 : On limite le nombre de colonnes pour que le temps d'affichage reste correcte
 23/05/05 : FQ 10254 : On traite sur un semestre pour une fiche quotidienne
                                 sur un an pour une fiche hebdomadaire
                                 sur deux ans pour une fiche mensuelle
{---------------------------------------------------------------------------------------}
function FICTAB.IsNbColOk : Boolean;
{---------------------------------------------------------------------------------------}
begin
  if TypeDesc in [dsDetail, dsSolde, dsCommission, dsBancaire, dsOPCVM, dsGroupe, dsControle] then begin
    if Periodicite = tp_1 then begin
      Result := (seInterval.Value <= 185) or
                (HShowMessage('0;' + Ecran.Caption + ';Vous ne pouvez pas traiter plus de 185 jours.'#13 +
                     'Voulez-vous poursuivre sur les 185 premiers jours ?;Q;YNC;N;C;', '', '') = mrYes);
      if Result and (seInterval.Value > 185) then
        NbColonne := 185;
    end
    else if Periodicite = tp_30 then begin
      Result := (NbColonne <= 24) or
                (HShowMessage('0;' + Ecran.Caption + ';Vous ne pouvez pas traiter plus de 24 mois.'#13 +
                     'Voulez-vous poursuivre sur les 24 premiers mois ?;Q;YNC;N;C;', '', '') = mrYes);
      if Result and (seInterval.Value > 24) then
        NbColonne := 24;
    end
    else begin
      Result := (NbColonne <= 52) or
                (HShowMessage('0;' + Ecran.Caption + ';Vous ne pouvez pas traiter plus de 52 périodes.'#13 +
                     'Voulez-vous poursuivre sur les 52 premieres périodes ?;Q;YNC;N;C;', '', '') = mrYes);
      if Result and (seInterval.Value > 52) then
        NbColonne := 52;
    end;
  end
  else
    Result := True;
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.SetDateDepart(Value : TDateTime);
{---------------------------------------------------------------------------------------}
begin
  SetControlText('DATEDE', DateToStr(Value));
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.SetNbColonne(Value : Integer);
{---------------------------------------------------------------------------------------}
begin
  seInterval.Value := Value;
end;

{---------------------------------------------------------------------------------------}
function FICTAB.GetDateFin : TDateTime;
{---------------------------------------------------------------------------------------}
begin
   case Periodicite of
     tp_1  : Result := DateDepart + NbColonne - 1;
     tp_7  : Result := DateDepart + NbColonne * 7 - 1;
     tp_15 : Result := DateDepart + NbColonne * NBJOURBIMENS - 1;
     else    Result := PlusDate(DateDepart, NbColonne, 'M') - 1;//+ 1; 06/12/05 A priori c'est mieux
   end;
end;

{Retourne la date de début de la colonne en fonction de l'indice de la colonne.
 Si FinPeriode est à True, on retourne la date de fin
{---------------------------------------------------------------------------------------}
function FICTAB.ColToDate(Col : Integer; FinPeriode : Boolean = False): TDateTime;
{---------------------------------------------------------------------------------------}
var
  d : TDateTime;
begin
  Result := DateDepart;
  if FinPeriode then
    case Periodicite of
      tp_1  : Result := DateDepart + Col;
      tp_7  : Result := DateDepart + 7 * (Col + 1);     //-1 ???
      tp_15 : Result := DateDepart + NBJOURBIMENS * (Col + 1); //-1 ???
      tp_30 : begin
                d := MoisToDate(RetourneMois(DateToStr(DateDepart), Col));
                Result := FinDeMois(d);
              end;
    end
  else
    case Periodicite of
      tp_1  : Result := DateDepart + Col;
      tp_7  : Result := DateDepart + 7 * Col;
      tp_15 : Result := DateDepart + NBJOURBIMENS * Col;
      tp_30 : begin
                d := MoisToDate(RetourneMois(DateToStr(DateDepart), Col));
                Result := d;
              end;
    end;
end;

{Retourne l'indice de la colonne -- SUIVANTE -- correspondant à la date (cf. Commentaire)
{---------------------------------------------------------------------------------------}
function FICTAB.DateToCol(Dt : TDateTime) : Integer;
{---------------------------------------------------------------------------------------}
begin
  Result := 0;
  {Remarque : on fait plus un (+1) pour le moment car cette fonction est appelé uniquement
              dans le suivi des OPCVM où on ne veut pas la colonne de la date mais celle
              qui suit ...
              Maintenant, la fonction est aussi utilisé Dans UFicTableau =>
                  PENSER À RETRANCHER 1, SI L'ON VEUT L'INDICE DE LA COLONNE DE LA DATE}
  case Periodicite of
    tp_1  : Result := Round(Dt - DateDepart) + 1;
    tp_7  : Result := Trunc(Dt - DateDepart) div 7 + 1;
    tp_15 : Result := Trunc(Dt - DateDepart) div NBJOURBIMENS + 1;
    tp_30 : Result := GetNbMoisEntre(DateDepart, Dt) + 1;
  end;
end;

{Retourne le titre de la colonne en fonction de l'indice
{---------------------------------------------------------------------------------------}
function FICTAB.RetourneCol(Col : Integer): string;
{---------------------------------------------------------------------------------------}
var
  dt : TDateTime;
begin
  case Periodicite of
    tp_1  : begin
              dt := DateDepart + Col; //+ 1
              Result := GetTitreFromDate(dt);
            end;
    tp_7  : begin
              dt := DateDepart + 7 * Col; //+ 1
              Result := GetTitreFromDate(dt);
            end;
    tp_15 : begin
              dt := DateDepart + NBJOURBIMENS * Col;//+ 1
              Result := GetTitreFromDate(dt);
            end;
    tp_30 : begin
              Result := RetourneMois(DateToStr(DateDepart), Col);
            end;
  end;
end;

{FQ 10237 : Retourne le titre de la colonne à une date donnée en fonction de la périodicité}
{---------------------------------------------------------------------------------------}
function FICTAB.GetTitreFromDate(CurDate : TDateTime) : string;
{---------------------------------------------------------------------------------------}
var
  d : TDateTime;
  p : Integer;
  s : string;
begin
  case Periodicite of
    tp_1  : Result := GetDateStr(CurDate); {10/10/06 : FQ 10376 DateToStr(CurDate);}
    tp_7  : begin
              {On récupère le nb de jours entre la date en cours et la date de départ}
              d := CurDate - DateDepart;
              {On calcule le nb de semaine que cela représente}
              p := Trunc(d) div 7;
              {Nom de la colonne sous la forme 01/01/04 - 08/01/04}
              DateTimeToString(Result, 'dd/mm/yy', DateDepart + 7 * p); // + 1
              DateTimeToString(s     , 'dd/mm/yy', DateDepart + 7 * (p + 1) - 1);
              Result := 'Du ' + Result + ' au ' + s;
            end;
    tp_15 : begin
              {On récupère le nb de jours entre la date en cours et la date de départ}
              d := CurDate - DateDepart;
              {On calcule le nb de semaine que cela représente}
              p := Trunc(d) div NBJOURBIMENS;
              {Nom de la colonne sous la forme 01/01/04 - 08/01/04}
              DateTimeToString(Result, 'dd/mm/yy', DateDepart + NBJOURBIMENS * p); // + 1
              DateTimeToString(s     , 'dd/mm/yy', DateDepart + NBJOURBIMENS * (p + 1) - 1);
              Result := 'Du ' + Result + ' au ' + s;
            end;
    tp_30 : begin
              Result := DateToMois(DateToStr(CurDate));
            end;
  end;
end;

{La colonne est-elle celle qui contient le 01/01
{---------------------------------------------------------------------------------------}
function FICTAB.IsColAvecSoldeInit(Col : Integer) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  case Periodicite of
   tp_1  : Result := ColToDate(Col) = DebutAnnee(ColToDate(Col));
   tp_7  : Result := Between(DebutAnnee(ColToDate(Col, True) - 1), ColToDate(Col), ColToDate(Col, True) - 1);
   tp_30 : Result := DebutAnnee(ColToDate(Col)) = ColToDate(Col);
  else
    Result := False;
  end;
end;

{Récupère le solde du prémier jour de la période de traitement, sauf ci ce jour est le 01/01
{---------------------------------------------------------------------------------------}
procedure FICTAB.GereSoldeDepart(var TobG : Tob; Gen, Dev : string; Tau : Double);
{---------------------------------------------------------------------------------------}
var
  Montant : Double;
  DebutOk : Boolean;
begin
  {La premère colonne contient-elle la date de début d'exercice ?}
  case Periodicite of
   tp_1  : DebutOk := DateDepart = DebutAnnee(DateDepart);
   tp_7  : DebutOk := Between(DebutAnnee(DateDepart + 7), DateDepart, DateDepart + 7);
   tp_30 : DebutOk := DebutAnnee(DateDepart) = DebutDeMois(DateDepart);
  else
    DebutOk := False;
  end;

  {Dans ce cas, la gestion du solde initial sera faite dans le SommerBanque}
  if DebutOk then Exit;
  {Récupère le solde du compte précédent}
  Montant := CalculSoldeInit(Gen, Dev, DateDepart) * tau; {***}
  {On stocke le solde du compte au premier jour de la période de traitement}
  TobG.SetDouble(GetTitreFromDate(DateDepart), Montant)
end;

{28/11/05 : FQ 10317 : En suivi hebdomadaire, le 01/01/AA peut tomber en milieu de semaine =>
            La gestion du solde initial peut tomber au milieu d'une période : il faut donc
            récupérer le total des opérations postérieures au 01/01/AA pour les ajouter au
            solde forcé !
{---------------------------------------------------------------------------------------}
function FICTAB.GetlOpeDebAnnee(Gen : string) : Double;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  Result := 0;
  if (Periodicite = tp_7) or (TypeDesc = dsPeriodique) then begin
    n := lOpeDebAnnee.IndexOf(Gen);
    if n > -1 then Result := TObjNombre(lOpeDebAnnee.Objects[n]).Nombre;
  end;
end;

{28/11/05 : FQ 10317 : En suivi hebdomadaire, le 01/01/AA peut tomber en milieu de semaine =>
            La gestion du solde initial peut tomber au milieu d'une période : il faut donc
            mémoriser le total des opérations postérieures au 01/01/AA pour les ajouter au
            solde forcé !
{---------------------------------------------------------------------------------------}
procedure FICTAB.SetlOpeDebAnnee(Gen : string; Mnt : Double; Dt : TDateTime);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  O : TObjNombre;
begin
  {Si présentation hebdomadaire, Si la colonne de la date contient le 01/01 et ...
   Remarque : DateToCol - 1, car DateToCol renvoie la colonne suivante de celle correspondant à la Date}
  if ((Periodicite = tp_7) and IsColAvecSoldeInit(DateToCol(Dt) - 1) and
     {... si la date est > 01/01 et si le premier janvier est bien celui voulu et non celui de l'année précédente}
     (Dt >= DebutAnnee(Dt)) and (DebutAnnee(Dt) >= DateDepart))
     {... Ou bien on est sur le solde périodique}
     or (TypeDesc = dsPeriodique) then begin
    n := lOpeDebAnnee.IndexOf(Gen);
    if n  = -1 then begin
      O := TObjNombre.Create;
      O.Nombre := 0;
      lOpeDebAnnee.AddObject(Gen, O);
    end
    else
      O := TObjNombre(lOpeDebAnnee.Objects[n]);

    O.Nombre := O.Nombre + Mnt;
  end;
end;

{28/11/05 : FQ 10317 : Je viens de m'apercevoir qu'en plus d'un mauvais calcul des soldes forcés,
            je n'en avais pas correctement géré l'affichage : en effet, en affichage mensuel, il
            est possible d'avoir deux débuts de millésime, pour un seul Integer (ColRouge) =>
            J'affine donc le test sur ColRouge lors du DrawCell
{---------------------------------------------------------------------------------------}
function FICTAB.HasSoldeReinit(Col : Integer) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := ColRouge = Col;
  {Si on est en affichage mensuel et que l'on demande plus de 12 mois ...}
  if not Result and (Periodicite = tp_30) and (NbColonne > 12) then
    {En Théorie (!), s'il y a 2 colonnes "rouges", ColRouge Contient l'indice le plus élevé}
    Result := Col = (ColRouge - 12);
end;

{16/01/06 : FQ 10328 : pour savoir si on traite les données dans la devise pivot
            ou dans celle sélectionnée
{---------------------------------------------------------------------------------------}
function FICTAB.IsMntPivot : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := ChpMontant = 'TE_MONTANT';
end;

{24/01/06 : FQ 10328 : Affecte ChpMontant en fonction de la diversité des devises dans la sélection
{---------------------------------------------------------------------------------------}
procedure FICTAB.GereLesDevises;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  D : string;
  S : string;
  B : Boolean;
begin
  {16/01/06 : FQ 10328 : Gestion des devises}
  {Cela serait à faire dans le suivi de treso, mais l'idéal serait d'uniformiser la fiche avec les autres !
   15/11/06 : Fait pour la fiche de suivi de Trésorerie}
  if TypeDesc in [dsSolde, dsTreso, dsBancaire] then begin
    {Affiche-t-on un message d'avertissement}
    B := False;
    {Chargements de la listes de comptes bancaires ainsi que de leur devise}
    CreerListeCompte;

    {On récupère la devise d'affichage}
    D := GetControlText('CBDEVISE');
    if (D = '') then D := V_PGI.DevisePivot;

    for n := 0 to ListeCpte.Detail.Count - 1 do begin
      B := D <> ListeCpte.Detail[n].GetString('ADEVISE');
      if B then Break;
    end;

    if B then begin
      D := RechDom('TTDEVISE', D, False);
      s := 'Tous les comptes ne sont pas en ' + D + '.'#13;
      s := s + 'Tous les montants seront affichés en ' + D + #13;
      s := s + 'après une conversion au taux de change du jour.';
      HShowMessage('0;' + Ecran.Caption + ';' + s + ';I;O;O;O;', '', '');
    end
    else
      {La devise des comptes sélectionnées et la devise sont les mêmes, on peut travailler sur TE_MONTANTDEV}
      ChpMontant := 'TE_MONTANTDEV';
  end
  else if TypeDesc in [dsGroupe, dsControle] then
    {Chargements de la listes de comptes bancaires}
    CreerListeCompte;

end;

{---------------------------------------------------------------------------------------}
function FICTAB.GetWhereDossier(Chp : string = 'TE_NODOSSIER') : string;
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  Result := '';

  {02/07/02 : Le filtre sur les NoDossier doit se faire aussi dans le cas d'une base autonome dans un regroupement
  if IsAvecRegroupement and IsTresoMultiSoc then begin}
  if IsAvecRegroupement and EstMultiSoc then begin
    s := GetControlText('REGROUPEMENT');
    if (TypeDesc in [dsSolde, dsbancaire, dsGroupe, dsTreso]) and
       {Si détail en affichage en solde, alors TypeDesc = dsSolde, mais il n'y a pas de combo REGROUPEMENT}
       (Assigned(GetControl('REGROUPEMENT'))) then
      if THMultiValComboBox(GetControl('REGROUPEMENT')).Tous then s := '';

    if (s <> '') then begin
      {Sur ces deux fiches, il s'agit d'un MultiVal}
      if TypeDesc in [dsGroupe, dsSolde, dsBancaire, dsTreso] then begin
        s := GetClauseIn(s);
        if s <> '' then
          Result := ' ' + Chp + ' IN (' + s + ') ';
      end
      else
        Result := ' ' + Chp + ' = "' + s + '" ';
    end
    else
      {Le regroupement tréso ne correspond pas forcément au regroupement Multi sociétés}
      Result :=  ' ' + Chp + ' ' + FiltreNodossier;
  end;
end;

{---------------------------------------------------------------------------------------}
function FICTAB.IsAvecRegroupement : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := TypeDesc in [dsGroupe, dsSolde, dsBancaire, dsOpcvm,
                         dsTreso, dsPeriodique, dsMensuel];
end;

{12/09/06 : Pour les budgets non globalisés, les natures ne correspondent pas aux valeurs de
            TE_NATURE, mais sont recalculées afin de pouvoir être distingués dans le DrawCell
{---------------------------------------------------------------------------------------}
function FICTAB.GetNatureToBudget(TypeFlux, Nat : string) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := '';
  if TypeDesc in [dsPeriodique, dsMensuel] then begin
      {Sur les écritures avec rubriques, il y a les trois natures P, R, S : les P et R vont ensembles}
      if (TypeFlux = CODETEMPO) then begin
        if (Nat = na_Realise) or (Nat = na_Prevision) then
          Result := na_PrevReal
        else
          Result := na_Simulation;
      end
      {Sur les écritures avec Flux, les prévisionnelles sont considérées comme des budgétaires}
      else begin
        if (Nat = na_Prevision) then Result := na_PrevSimul
                                else Result := na_Realise;
      end;
  end;
end;

{12/09/06 : Cheminement inverse de ci-dessus : retourne la nature pour TRECRITURE
{---------------------------------------------------------------------------------------}
function FICTAB.GetNatureFromBudget(TypeFlux, CodeFlux, Nat : string) : string;
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  Result := '';
  if TypeDesc in [dsPeriodique, dsMensuel] then begin
    {Il s'agit de écritures à partir des rubriques}
    if ((TypeFlux = 'TRP') or (TypeFlux = 'TRN')) and not
       ExisteSQL('SELECT TFT_FLUX FROM FLUXTRESO WHERE TFT_FLUX = "' + CodeFlux + '"') then begin
      if Nat = na_Simulation then
        Result := 'AND TE_NATURE = "' + na_Simulation + '" '
      else begin
        {Les écritures "réelles" sur les rubriques peuvent avoir deux natures : prévisionnelle et réalisée :
         il faut donc regarder s'il y a un filtre dans la combo de nature}
        s := GetControlText('MCNATURE');
        if THMultiValComboBox(GetControl('MCNATURE')).Tous or
           (s = '') or ((Pos(na_Realise, s) > 0) and (Pos(na_Prevision, s) > 0)) then
          Result := 'AND (TE_NATURE IN ("' + na_Realise + '", "' + na_Prevision + '")) '
        else if Pos(na_Realise, s) > 0 then
          Result := 'AND TE_NATURE = "' + na_Realise + '" '
        else
          Result := 'AND TE_NATURE = "' + na_Prevision + '" '
      end;
    end

    {Il s'agit d'écritures à partir de flux}
    else begin
      if Nat = na_PrevSimul then Result := 'AND TE_NATURE = "' + na_Prevision + '" '
                            else Result := 'AND TE_NATURE = "' + na_Realise + '" ';
    end;
  end;
end;

{En eAgl, le fait de passer CritModified relance un deuxième chargement en cas de présence
 d'un filtre : lors du chargement du filtre, CritereClick est exécuté sur chaque OnChange
 ou Onclick des zones => CritModified = True => on relance la recherche. Le seul moyen que
 j'ai trouvé de contourner le problème est d'intervenir en fin de chargement du filtre d'où
 la réimplémentation de ListeFiltre.OnSelect
{---------------------------------------------------------------------------------------}
procedure FICTAB.FormAfterShow;
{---------------------------------------------------------------------------------------}
begin
  TFMul(Ecran).ListeFiltre.OnSelect := InitSelectFiltre;
  {En cas d'appel de la fiche avec TheTob contenant les filtres à mettre à la fiche}
//  InitControls;
end;

{Réimplémentation de ListeFiltre.OnSelect afin de pouvoir intervenir à la fin du traitement
{---------------------------------------------------------------------------------------}
procedure FICTAB.InitSelectFiltre(T : TOB);
{---------------------------------------------------------------------------------------}
var
  Lines   : HTStringList;
  n       : Integer;
  stChamp : string;
  stVal   : string;
  E       : TFMul;
begin
  if V_PGI.AGLDesigning then Exit;
  E := TFMul(Ecran);
  if T = nil then begin
    CanSetModified := True;
    Exit;
  end;
  Lines := HTStringList.Create;

  for n := 0 to T.Detail.Count - 1 do begin
    stChamp := T.Detail[n].GetValue('N');
    stVal := T.Detail[n].GetValue('V');
    Lines.Add(stChamp + ';' + stVal);
  end;

  VideFiltre(E.FFiltres, E.Pages, False);
  ChargeCritMemoire(Lines, E.Pages);
  Lines.Free;

  {Traitement à effectuer après chargement du filtre}
  MySelectFiltre;
end;

{27/01/06 : On permet d'affecter CritModified à True dans CritereClick
{---------------------------------------------------------------------------------------}
procedure FICTAB.MySelectFiltre;
{---------------------------------------------------------------------------------------}
begin
  CanSetModified := True;
end;

{---------------------------------------------------------------------------------------}
procedure FICTAB.ImprimerClick;
{---------------------------------------------------------------------------------------}
var
  n       : Integer;
  ObjEtat : TObjEtats;
begin
  ObjEtat := TObjEtats.Create(Ecran.Caption, Grid);
  try
    ObjEtat.MajTitre(0, 'Libellé');
    for n := 1 to ObjEtat.NbColonne - 1 do
      ObjEtat.MajAlign(n, ali_Droite);
    ObjEtat.Imprimer
  finally
    FreeAndNil(ObjEtat);
  end;
end;

{20/06/07 : FQ 10480 : Gestion du concept la création / Suppresion des flux
{---------------------------------------------------------------------------------------}
procedure FICTAB.GereDroitsCreation;
{---------------------------------------------------------------------------------------}
begin
  if not (TypeDesc in [dsControle, dsOPCVM, dsGroupe]) then
    CanCreateEcr(TPopupMenu(GetControl('POPUPMENU')).Items[1])
  else if TypeDesc = dsGroupe then
    TPopupMenu(GetControl('POPUPMENU')).Items[1].Visible := False;
end;

{10/08/07 : Gestion des filtres sur les confidentialité
{---------------------------------------------------------------------------------------}
function FICTAB.GetWhereConfidentialite : string;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRCONF}
  Result := ObjConf.GetWhereSQLConf(tyc_Banque);
  if Result <> '' then Result := ' AND (' + Result  + ') ';
  {$ELSE}
  Result := '';
  {$ENDIF TRCONF}
end;

{$IFDEF TRCONF}
{---------------------------------------------------------------------------------------}
function FICTAB.GetObjConf : TObjConfidentialite;
{---------------------------------------------------------------------------------------}
begin
  if not Assigned(FObjConf) then
    FObjConf := TObjConfidentialite.Create(V_PGI.User);
  Result := FObjConf;
end;
{$ENDIF TRCONF}

initialization
  RegisterClasses([FICTAB]);

end.

