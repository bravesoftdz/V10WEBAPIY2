unit PlanUtil;

interface

uses {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Graphics,
  forms,
  HPlanning,
  StdCtrls,
  ExtCtrls,
  Controls,
  HCtrls,
  HStatus,
  MsgUtil,
  SysUtils,
  Hent1,
  dialogs,
  HMsgBox,
  HRichOLE,
  Paramsoc,
  HeureUtil,
  uJoursFeries,
	UAFO_Ferie,
  Windows,
  AglMail,
  MailOl,
  UDateUtils,
  BTPUtil,
  UTOB,
  CbpMCD,
  CbpEnumerator
  ;

type
  THPlanningBTP = Class(THPlanning)
  public
    FamRes          : String; //Famille ressource
    CodeOnglet      : String; //Permet de définir quelles information est dans l'onglet
    TypePlanning    : String; //PSF : Sous-Famille - PTA : Type Action
    AppelsTraites   : Boolean; // Gestion des appels traités ou pas
  end;

  RecordPlanning = record
    TobItems          : TOB;
    TobEtats          : TOB;
    TobRes            : TOB;
    TobCols           : TOB;
    TobRows           : TOB;
    TobEvents         : TOB;
    TobPeriodeDivers  : TOB;
  end;


  Procedure AffichageDateHeure(TobItem : Tob; Sens : String);
  procedure AfficheLFixe(P: THPlanningBTP; T: Tob);

  Procedure ChargeCadencement(var R: RecordPlanning; P: THPlanningBTP; T: TOB);
  Function  ChargeColonneFixe(TypRes : String; var P: THPlanningBTP) : String;
  Function  ChargeColonnePAR (var P: THPlanningBTP) : String;
  function  ChargeColonnePFM (TypRes : String; var P: THPlanningBTP) : String;  //Planning Famille Matériel
  function  ChargeColonnePLA (TypRes : String; var P: THPlanningBTP) : String;  //Planning Intervention
  Function  ChargeColonnePMA (TypRes : String; var P: THPlanningBTP) : String;  //Planning Matériel
  function  ChargeColonnePPA (TypRes : String; var P: THPlanningBTP) : String;  //Planning Type Action Matériel
  Function  ChargeColonnePSF (TypRes : String; var P: THPlanningBTP) : String;  //Planning Sous-Type Ressource
  Function  ChargeColonnePTR (TypRes : String; var P: THPlanningBTP) : String;  //Planning Type Ressource
  Function  ChargeColonnePFO (TypRes : String; var P: THPlanningBTP) : String;  //Planning Fonction Ressource
  Procedure ChargeEvent(var R: RecordPlanning; var P: THPlanningBTP);
  Procedure ChargeIcone(TypeAction : String; TobItem : TOB);
  procedure ChargeItemPlanning  (var R: RecordPlanning; var P: THPlanningBTP; T: TOB; DateEnCours: TDateTime);
  Procedure ChargementEtat      (var R: RecordPlanning; T: TOB; DateEnCours: TDateTime; ModeGestionPlanning: Boolean);
  procedure ChargeParamPlanning (var R: RecordPlanning; var P: THPlanningBTP; T: TOB; DateEnCours: TDateTime; ModegestionPlanning: Boolean);
  Function  ChargeRequeteAbsSalaries (P : THPlanningBTP; DateDebut, DateFin : TdateTime) : String;
  Function  ChargeRequeteActionsGRC  (P : THPlanningBTP; DateDebut, DateFin : TdateTime) : String;
  Function  ChargeRequeteEventMat    (P : THPlanningBTP; DateDebut, DateFin : TdateTime) : String;
  Function  ChargeRequeteFamilleMat  (P : THPlanningBTP; DateDebut, DateFin : TdateTime) : String;
  Function  ChargeRequeteIntervention(P : THPlanningBTP; DateDebut, DateFin : TdateTime) : String;
  Function  ChargeRequeteRessourceChantier(P : THPlanningBTP; DateDebut, DateFin : TdateTime) : String;
  Function  ChargeRequeteTypeAction  (P : THPlanningBTP; DateDebut, DateFin : TdateTime) : String;
  Function  ChargeRequeteUnifie      (P : THPlanningBTP; DateDebut, DateFin : TdateTime) : String;
  Procedure CompleteTobItem(TobItem : Tob);
  Procedure ControleCadencement(Cadencement : String;R: RecordPlanning;P: THPlanningBTP; T : TOB);
  Procedure CreationZoneTobItem(TobTmp : Tob);

  Procedure RechercheContenuItem(TypeEnreg : String; TobItem : TOb);
  Procedure RechercheContenuHint(TypeEnreg : String; TobItem : TOb);

  procedure TransformeEvenementEnItem(TobItems, TOBEvent : TOB);



Procedure GestionAffichageZoneItem(Numliste : String);
Procedure GestionAffichageZoneHint(Numliste : String);

procedure MajParamplanning;
//procedure RecupChamps(var ListeChamps: array of string; iTable: integer);
procedure TobFree(var T: TOB);

function AffecteDataType(NomChamp: string): string;
function AnalyseChamp(Tobligne: TOB; ListeChamps: string; Formats: string): string;
Function ChargeFamres   : String;
Function ChargeFamMat   : String;
Function ChargeFonction : String;

function ChargeStrByTablette(CodeTablette,ZoneARecuperer: String): String;
function LibelleChamp(CodeTablette: string): string;
//function MajItem(Liste: string; var Format: string; ZoneARecup : String): string;
function TabletteAssociee(NomTable: string): string;
function GestionDateFinPourModif (DateDeFin : TdateTime) : TdateTime;
function GestionDateFinPourEnreg (DateDeFin : TdateTime) : TdateTime;
function CalculDateDebut (DateDebut : TdateTime) : TdateTime;
function CalculDateFin (DateFin : TdateTime) : TdateTime; Overload

function ControleHoraires (var DateDebEvt,DateFinEvt : TdateTime; ModeOption: string; Duree : double; CodeRessource : string) : boolean;
function GetDescriptifCourt ( TOBD : TOB) : string;

function  UpdateActionGRC (TobLigItem : TOB): boolean;
procedure AddChampsSupItems (TOBItem : TOB);


//procedure EnvoieEltEmailFromGRC (Auxiliaire: string;NumAction: integer; TypeEnvoie : string; NewDate : TdateTime;newdelai : integer; InfoFrom : boolean=false);
//procedure EnvoieEltEmailFromAction ( TOBElt : TOB; TypeEnvoie : string; NewDate : TdateTime;newdelai : integer; InfoFrom : boolean=false; NewRess : string='');
procedure EnvoieEmailFromIntervention (Sender : TObject; TobItem : TOB);

var

  ModeSaisie: Boolean = True; // True On peut travailler, False on consulte
  NumLigneTO: Integer;
  NumYield: Integer;

  GestAffiche   : AffGrille;

  TobParametres : Tob;
  TobIcone	: Tob;

  Cadencement: string;
  CtrlCal    : Boolean;

implementation

uses  uBtpEtatPlanning ,
      DateUtils,
      TntClasses,
      Classes,
      UtilsParc,
      DB,
      UtilsMail,
      UtilsRapport,
      UtilActionPlanning,
      UtilsParcPlanning;

var
  ContenuItem: string;
  FormatItem: string;
  ContenuHint: string;
  FormatHint: string;

procedure TobFree(var T: TOB);
begin
  if T <> nil then
  begin
    T.Free;
    T := nil;
  end;
end;


procedure ChargeParamPlanning(var R: RecordPlanning; var P: THPlanningBTP; T: TOB; DateEnCours: TDateTime; ModeGestionPlanning: Boolean);
var
  NomDuChamp  : array[1..9] of string;
  LibDuChamp  : array[1..9] of string;
  TypeDuChamp : array[1..9] of string;
  TailleChamp : array[1..9] of Integer;
  Formegraphique: Variant;
  SQL         : string;
  Req         : string;
  TypRes      : String;
  NomHpp      : string;
  StrTempo    : String;
  ZoneRes     : String;
  LibZone     : String;
  QRessource  : TQuery;
  QEtats      : TQuery;
  Boucle      : integer;
  I           : integer;
  NbCF        : integer;
  NbCT        : integer;
  Ind         : integer;
  TailleColFixe: Integer;
  IconesExiste: Boolean;
  Regroupement: Boolean;
begin

  If T <> nil then TobParametres := T;

  //Remise à zéro des tob Etats et Ressources
  if R.Tobres = nil then
  begin
    R.TobRes := Tob.Create('THE RESSOURCES', nil, -1);
  end else
  begin
    R.TobRes.clearDetail;
    R.ToBres.InitValeurs(false); 
  end;

  NumLigneTO  := -1;
  NumYield    := -1;

  // Barre de patiente
  InitMove(10, 'Chargement du planning en cours');

  //Mise en place de la gestion du surbooking et des Week-end à revoir
  with P do
  begin
    SurBooking      := True;
    ActiveSaturday  := True;
    ActiveSunday    := True;
  end;

  // MAJ regroupement de date
  //P.ActiveLigneGroupeDate := TobParametres.Getvalue('HPP_AFFDATEGROUP') = 'X';

  if pos(P.TypePlanning, 'PMA;PFM;PPA') > 0 then
    P.Famres := ChargeFamMat
  else if pos(P.TypePlanning, 'PFO') > 0 then
    P.Famres := ChargeFonction
  else
    P.Famres := ChargeFamRes;

  // Changement date d'un séjour sur planning possible O/N
  P.MoveHorizontal := True;

  // Chargement de la forme graphique
  FormeGraphique := TobParametres.Getvalue('HPP_FORMEGRAPHIQUE');

  if (Formegraphique = 'PGF') then
    P.FormeGraphique := pgFleche
  else if (FormeGraphique = 'PGB') then
    P.FormeGraphique := pgBerceau
  else if (FormeGraphique = 'PGL') then
    P.FormeGraphique := pgLosange
  else if (FormeGraphique = 'PGE') then
    P.FormeGraphique := pgEtoile
  else if (FormeGraphique = 'PGA') then
    P.FormeGraphique := pgRoundRect
  else if (FormeGraphique = 'PGI') then
    P.FormeGraphique := pgBisot
  else if (FormeGraphique = 'PGR') then
    P.FormeGraphique := pgRectangle
  else
    P.FormeGraphique := pgFleche;

  //Cadencement
  ChargeCadencement(R, P, T);
  //
  //Paramètre de controle du calendrier
  CtrlCal := T.GetBoolean('HPP_CONTROLCAL');

  MoveCur(FALSE);

  // Gestion des couleurs du planning
  P.ColorSelection := StringToColor(T.Getvalue('HPP_COULSELECTION'));
  P.ColorBackground := StringToColor(T.Getvalue('HPP_COULEURFOND'));
  if T.Getvalue('HPP_COULEURSAMEDI') <> '' then
    P.ColorOfSaturday := StringToColor(T.Getvalue('HPP_COULEURSAMEDI'));
  if T.Getvalue('HPP_COULDIMANCHE') <> '' then
    P.ColorOfSunday := StringToColor(T.Getvalue('HPP_COULDIMANCHE'));
  if T.Getvalue('HPP_COULLUNDI') <> '' then
    P.ColorOfMonday := StringToColor(T.Getvalue('HPP_COULLUNDI'));
  if T.Getvalue('HPP_COULMARDI') <> '' then
    P.ColorOfTuesday := StringToColor(T.Getvalue('HPP_COULMARDI'));
  if T.Getvalue('HPP_COULMERCREDI') <> '' then
    P.ColorOfWednesday := StringToColor(T.Getvalue('HPP_COULMERCREDI'));
  if T.Getvalue('HPP_COULJEUDI') <> '' then
    P.ColorOfThuesday := StringToColor(T.Getvalue('HPP_COULJEUDI'));
  if T.Getvalue('HPP_COULVENDREDI') <> '' then
    P.ColorOfFriday := StringToColor(T.Getvalue('HPP_COULVENDREDI'));
  if T.Getvalue('HPP_COULVEILJF') <> '' then
  begin
    P.ColorVeilleJoursFeries := StringToColor(T.Getvalue('HPP_COULVEILJF'));
    P.GestionVeilleJoursFeries := True;
  end;
  if T.Getvalue('HPP_COULJFERIE') <> '' then
  begin
    P.ColorJoursFeries := StringToColor(T.Getvalue('HPP_COULJFERIE'));
    P.GestionJoursFeriesActive := True;
  end;

  //Lignes du planning
  if pos(P.TypePlanning, 'PMA;PFM;PPA') > 0 then
    P.RowFieldID      := 'R_CODEMATERIEL'
  else if P.TypePlanning ='PAR' then
    P.RowFieldID      := 'R_AFFAIRE'
  else
    P.RowFieldID      := 'R_RESSOURCE';

  P.RowFieldReadOnly  := 'R_RO';
  P.RowFieldColor     := 'R_COLOR';

  MoveCur(FALSE);

  // MAJ de la police des colonnes
  P.FontNameColRowFixed := T.Getvalue('HPP_FONTCOLONNE');

  // Encadrement
  P.FrameOn := True;

  // Champ des items
  P.ChampdateDebut  := 'BPL_DATEAFFDEB';
  P.ChampDateFin    := 'BPL_DATEAFFFIN';
  P.MultiLine       := False;
  P.ChampLibelle    := 'LIBELLE';

  GestionAffichageZoneItem(T.Getvalue('HPP_CONTENUITEM'));
  if pos('#13#10', ContenuItem) <> 0 then P.MultiLine := True;

  GestionAffichageZoneHint(T.Getvalue('HPP_CONTENUHINT'));

  // Champ des Type Actions
  P.EtatChampCode             := 'BTA_BTETAT';
  P.EtatChampLibelle          := 'BTA_LIBELLE';
  P.EtatChampBackGroundColor  := 'BTA_COULEURFOND';
  P.EtatChampFontColor        := 'BTA_COULEUR';
  P.EtatChampFontName         := 'BTA_FONTE';
  P.EtatChampFontSize         := 'BTA_FONTESIZE';
  P.EtatChampFontStyle        := 'BTA_FONTESTYLE';
  P.EtatChampIcone            := 'BTA_NUMEROICONE';

  MoveCur(FALSE);

  // Gestion des évenements (Salons, concert, ...)
  P.ChampCodeEvenement        := 'HEV_EVENEMENT';
  P.ChampLibelleEvenement     := 'HEV_LIBELLE';
  P.ChampDateDebutEvenement   := 'HEV_DATEDEBUT';
  P.ChampDateDeFinEvenement   := 'HEV_DATEFIN';
  P.ChampCouleurEvenement     := 'HEV_COULEUR';
  P.ChampStyleEvenement       := 'HEV_STYLE';

  MoveCur(FALSE);

  // Activation bulles d'aide
  P.ShowHint := ModeGestionPlanning;

  // Ordre de tri des colonnes.
  NbCF := T.getinteger('HPP_NBCOLDIVERS');

  //Gestion des Sous-TypeRessource
  TypRes := T.getvalue('HPP_PLANNINGTYPETD');

  //chargement de la ou des colonnes fixes du planning
  Req := ChargeColonneFixe(TypRes, P);

  //chargement des "ressources"
  QRessource  := OpenSQL(req, True,-1,'',true);
  R.TobRes.LoadDetailDB('', '', '', Qressource, True);
  ferme(Qressource);

  IconesExiste := False;

  if R.TobRes.Detail.Count = 0 then
  begin
    if Pos(P.TypePlanning, 'PMA;PFM') > 0 then
      PGIError('Le Code Famille ' + P.CodeOnglet + ' ne dispose d''aucune ressource', 'Erreur Planning')
    else if Pos(P.TypePlanning, 'PAR,PLA;PTA;PTR;PRA') > 0 then
      PGIError('Le Type Ressource ' + P.CodeOnglet + ' ne dispose d''aucune ressource', 'Erreur Planning')
    else if P.TypePlanning = 'PPA' then
      PGIError('Le type d''action ' + P.CodeOnglet + ' ne dispose d''aucune ressource', 'Erreur Planning')
    else if P.TypePlanning = 'PSF' then
      PGIError('Le sous-type ressource ' + P.CodeOnglet + ' ne dispose d''aucune ressource', 'Erreur Planning')
    else if P.TypePlanning = 'PFO' then
      PGIError('La fonction ' + P.CodeOnglet + ' ne dispose d''aucune ressource', 'Erreur Planning')
    else
      PGIError('aucune ressource a afficher', 'Erreur Planning');
    P.parent.Visible := false;

    Exit;
  end;

  MoveCur(FALSE);

  ind := 0;

  //Affichage du libellé pour un tri par sous famille ressource en fonction parametrage Planning
  if TypRes = 'X' then
  Begin
    NomDuChamp[1]  := 'BTR_LIBELLE';
    LibDuChamp[1]  := 'Ss-Type Ressource';
    TypeduChamp[1] := 'VARCHAR(35)';
    TailleChamp[1] := 150;
    Ind := 1;
  end;

  // Récupératon des libéllés des champs pour les colonnes d'entete
  for boucle := 1 to NbCf do
  begin
    StrTempo := '';
    NomHpp := 'HPP_LIBCOLENT';
    //
    if pos(P.TypePlanning, 'PMA;PFM;PPA') > 0 then
    begin
      ZoneRes := LectLibCol('CC', 'BMA', T.getvalue(NomHpp + IntToStr(Boucle + 3)), 'CC_LIBRE');
      LibZone := LectLibCol('CC', 'BMA', T.Getvalue(NomHpp + IntToStr(Boucle + 3)), 'CC_LIBELLE');
    end
    else if pos(P.TypePlanning, 'PAR') > 0 then
    begin
      ZoneRes := LectLibCol('CC', 'BAR', T.getvalue(NomHpp + IntToStr(Boucle + 3)), 'CC_LIBRE');
      LibZone := LectLibCol('CC', 'BAR', T.Getvalue(NomHpp + IntToStr(Boucle + 3)), 'CC_LIBELLE');
    end
    else
    begin
      ZoneRes := LectLibCol('CC', 'BLI', T.getvalue(NomHpp + IntToStr(Boucle + 3)), 'CC_LIBRE');
      LibZone := LectLibCol('CC', 'BLI', T.Getvalue(NomHpp + IntToStr(Boucle + 3)), 'CC_LIBELLE');
    end;
    //
    TailleColFixe := StrToInt(format('%2.2d', [Integer(T.Getvalue('HPP_TAILLECOLENT' + IntToStr(Boucle+3)))]));
    //
    for i := 1 to Boucle do
    Begin
      StrTempo := StrTempo + NomDuChamp[i];
    end;
    //
    if pos(ZoneRes, StrTempo) = 0 then
    Begin
      Ind := Ind + 1;
      NomDuChamp[Ind]  := ZoneRes;
      LibDuChamp[Ind]  := LibZone;
      TypeduChamp[Ind] := ChampToType(ZoneRes);
      TailleChamp[Ind] := TailleColFixe;
    end
  end;

  MoveCur(FALSE);

  P.DisplayEnteteColFixed := True;
  P.SortOnColFixed        := True;

  if pos(P.TypePlanning, 'PMA;PFM;PPA') > 0 then
    P.ResChampID := 'BMA_CODEMATERIEL'
  else if pos(P.TypePlanning, 'PAR') > 0 then
    P.ResChampID := 'AFF_AFFAIRE'
  else
    P.ResChampID := 'ARS_RESSOURCE';

  // Initalisation des colonnes d'entete
  for Boucle := 1 to Ind do //NbCf do
  begin
    if Boucle = 1 then
    Begin
      P.TokenFieldColFixed  := NomDuChamp[Boucle];
      P.TokenSizeColFixed   := IntToStr(TailleChamp[Boucle]);
      P.TokenAlignColFixed  := 'C';
      P.TokenFieldColEntete := LibDuChamp[Boucle];
    end
    else
    begin
      P.TokenFieldColFixed  := P.TokenFieldColFixed + ';' + NomDuChamp[Boucle];
      P.TokenSizeColFixed   := P.TokenSizeColFixed + ';' + IntToStr(TailleChamp[Boucle]);
      P.TokenAlignColFixed  := P.TokenAlignColFixed + ';C';
      P.TokenFieldColEntete := P.TokenFieldColEntete + '; ' + LibDuChamp[Boucle];
    end;
  end;

  // la précédente tob des Ressources n'a pas été vidé il faut donc la vider
  P.TobRes := R.TobRes;

  //Chargement des types d'action pour affichage des items
  ChargementEtat(R, T, DateEnCours, ModeGestionPlanning);
  MajFontStyle(R.TobEtats);
  P.TobEtats := R.TobEtats;

  MoveCur(FALSE);

  // Icones
  if T.Getvalue('HPP_AFFICHEICONE') = 'X' then P.ChampIcone := 'ICONETYPDOS';

  // MAJ Taille des colonnes largeurs
  P.ColSizeData   := T.Getvalue('HPP_TAILLECOLONNE');
  P.ColSizeEntete := T.Getvalue('HPP_TAILLECOLENTET');

  // MAJ Taille des colonnes largeurs
  P.RowSizeData   := T.Getvalue('HPP_HAUTLIGNEDATA');
  P.RowSizeEntete := T.Getvalue('HPP_HAUTLIGNEENT');

  if pos(P.TypePlanning, 'PMA;PFM;PPA') > 0 then
    P.ChampLineID  := 'BPL_MATERIEL'
  else if pos(P.TypePlanning, 'PRA') > 0 then
    P.ChampLineID  := 'BPL_RESSOURCE'
  else if pos(P.TypePlanning, 'PAR') > 0 then
    P.ChampLineID  := 'BPL_AFFAIRE'
  else
    P.ChampLineID  := 'BPL_RESSOURCE';

  P.ChampEtat    := 'BPL_BTETAT';

  if (ModegestionPlanning) then P.ChampHint := 'HINT';

  ChargeItemPlanning(R, P, T, DateEnCours);

end;

//Chargement des informations propres à la famille ressource si planning Standard
Function ChargeFonction : String;
var QRecupFonction: Tquery;
    Req           : String;
    LibFonction   : String;
begin

  if TobParametres.Getvalue('HPP_VISUTYPERES') = 'X' then
     Result := 'GEN'
  else
     if TobParametres.getvalue('HPP_FAMRES') = '' then
        Result := 'GEN'
     else
        Result := TobParametres.getvalue('HPP_FAMRES');

  //Ceci n'est pas la requete initiale de plus elle fait planter je la vire !!!!!
  // Req := 'SELECT AFO_LIBELLE FROM FONCTION '
  //'LEFT JOIN BRESSOURCE ON BFO_RESSOURCE=AFO_RESSOURCE ';

  //Récupération des information de la famille de Planning
  Req := 'SELECT AFO_LIBELLE FROM FONCTION lEFT JOIN BFONCTION ON AFO_FONCTION=BFO_FONCTION ';

  //Modif pour planning Général
  if Result <> 'GEN' then
     Req := Req + 'WHERE BFO_FONCTION="' + Result + '" AND BFO_GEREPLANNING = "X"'
  else
     Req := Req + 'WHERE BFO_GEREPLANNING = "X"';

  QRecupFonction := OpenSQL(Req, True,-1,'',true);

  if not QRecupFonction.eof then
  begin
    LibFonction := QRecupFonction.findfield('AFO_LIBELLE').Asstring;
  end;

  ferme(QRecupFonction);

end;

//Chargement des informations propres à la famille ressource si planning Standard
Function ChargeFamres : String;
var QRecupFamRes  : Tquery;
    Req           : String;
    LibFamRes     : String;
Begin

  if TobParametres.Getvalue('HPP_VISUTYPERES') = 'X' then //==> Planning Général
     Result := 'GEN'
  else
     if TobParametres.getvalue('HPP_FAMRES') = '' then
        Result := 'GEN'
     else
        Result := TobParametres.getvalue('HPP_FAMRES');

  //Récupération des information de la famille de Planning
  Req := 'SELECT HFR_LIBELLE,HFR_HRPERIODEDEBUT,HFR_HRPERIODEFIN FROM HRFAMRES ';

  //Modif pour planning Général
  if Result <> 'GEN' then
     Req := Req + 'WHERE HFR_FAMRES="' + Result + '"';

  QRecupFamRes := OpenSQL(Req, True,-1,'',true);

  if Result <> 'GEN' then
  begin
    if not QRecupFamRes.eof then
    begin
      LibFamRes := QRecupFamRes.findfield('HFR_LIBELLE').Asstring;
    end;
  end;

  ferme(QRecupFamRes);

end;

//Chargement des informations propres à la famille Matériel si planning Standard
Function ChargeFamMat : String;
var QRecupFamMat  : Tquery;
    Req           : String;
    LibFamMat     : String;
Begin

  if TobParametres.Getvalue('HPP_VISUTYPERES') = 'X' then
     Result := 'GEN'
  else
  begin
    if TobParametres.getvalue('HPP_FAMRES') = '' then
    begin
      if TobParametres.getvalue('HPP_FAMMATGERE') = '' then
        Result := 'GEN'
      else
        Result := FormatMultiValComboforSQL(TobParametres.getvalue('HPP_FAMMATGERE'));
    end
    else
    begin
      Result :=  TobParametres.getvalue('HPP_FAMRES');
      //Récupération des information de la famille de Planning
      Req := 'SELECT BFM_LIBELLE FROM BTFAMILLEMAT ';
      //Modif pour planning Général
      Req := Req + 'WHERE BFM_CODEFAMILLE ="' + Result + '" AND BFM_NONGEREPLANNING="X"';
      Result := '"' + result + '"';
      QRecupFamMat := OpenSQL(Req, True,-1,'',true);
      if not QRecupFamMat.eof then LibFamMat := QRecupFamMat.findfield('BFM_LIBELLE').Asstring;
      ferme(QRecupFamMat);
    end;
  end;

end;


function ConstitueRequeteAbsenceConges (P :THPlanningBTP;DateD,DateF : TdateTime) : string ;
begin
	result := '( '+ // 1
            '(PCN_DATEDEBUTABS>="' + UsDateTime(DateD) + '" AND PCN_DATEDEBUTABS<="' + UsDateTime(DateF) + '") OR '+
						'(PCN_DATEFINABS>="'   + USDateTime(DateD) + '" AND PCN_DATEFINABS<="'   + USDateTime(DateF) + '") OR ' +
						'(PCN_DATEDEBUTABS<="' + UsDateTime(DateD) + '" AND PCN_DATEFINABS>="'   + UsDateTime(DateF) + '")' +
            ')'; //1
end;

procedure EnvoieEltEmailFromGRC (Auxiliaire: string;NumAction: integer; TypeEnvoie : string; NewDate : TdateTime;newdelai : integer; InfoFrom : boolean=false);
var maRequete : string;
		QQ : Tquery;
    Sujet : HString;
    Corps : HTStringList;
    ResultMailForm : TResultMailForm;
    copie,fichiers,Destinataire : string;
begin
	Corps := hTStringList.Create;
//
	Destinataire := '';
  QQ := OpenSql ('SELECT RAI_RESSOURCE,ARS_EMAIL FROM ACTIONINTERVENANT '+
             		 'LEFT JOIN RESSOURCE ON ARS_RESSOURCE=RAI_RESSOURCE '+
  							 'WHERE '+
  							 'RAI_AUXILIAIRE="'+Auxiliaire+'" '+
                 'AND RAI_NUMACTION='+inttoStr(NumAction),true,-1,'',true);
  if not QQ.eof then
  begin
  	QQ.First;
    repeat
    	if  QQ.FindField('ARS_EMAIL').AsString <> '' then
      begin
        if destinataire = '' then Destinataire := QQ.FindField('ARS_EMAIL').AsString
                             else Destinataire := destinataire +';'+ QQ.FindField('ARS_EMAIL').AsString;
      end;
      QQ.Next;
    until QQ.eof;
  end;
  ferme (QQ);
//
  Marequete:='SELECT RAC_NUMACTION,RAC_AUXILIAIRE,RAC_INTERVENANT,RAC_LIBELLE,RAC_DATEACTION,RAC_HEUREACTION,'+
  					 'RAC_DUREEACTION,ARS_EMAIL,T_LIBELLE,'+
  					 'T_ADRESSE1,T_ADRESSE2,T_CODEPOSTAL,T_VILLE FROM RTVACTIONS '+
             'LEFT JOIN RESSOURCE ON ARS_RESSOURCE=RAC_INTERVENANT '+
             'WHERE RAC_NUMACTION='+IntToStr(NumAction)+' AND RAC_AUXILIAIRE="'+Auxiliaire+'"';
  TRY
  	QQ := OpenSql (MaRequete,True,1,'',true);
    if not QQ.eof then
    begin
      if QQ.findField ('ARS_EMAIL').AsString ='' then exit;
//      Destinataire := QQ.findField ('ARS_EMAIL').AsString;
      if TypeEnvoie = 'S' then
      begin
        Sujet := 'Annulation de votre action '+IntToStr(QQ.findField('RAC_NUMACTION').AsInteger);
        Corps.Add('Votre action N°'+IntToStr(QQ.findField('RAC_NUMACTION').AsInteger)+' a été annulée');
        Corps.Add('');
        Corps.Add('Pour information :');
        Corps.Add('Action :'+QQ.findField('RAC_LIBELLE').AsString +' pour le client '+QQ.findField('T_LIBELLE').AsString);
        Corps.Add('Adresse :');
        if QQ.findField('T_ADRESSE1').AsString <> '' then
        begin
          Corps.add('         '+QQ.findField('T_ADRESSE1').AsString);
        end;
        if QQ.findField('T_ADRESSE2').AsString <> '' then
        begin
          Corps.Add('         '+QQ.findField('T_ADRESSE2').AsString);
        end;
        if (QQ.findField('T_CODEPOSTAL').AsString <> '') or (QQ.findField('T_VILLE').AsString<>'') then
        begin
        	Corps.Add('         '+QQ.findField('T_CODEPOSTAL').AsString+' '+QQ.findField('T_VILLE').AsString);
        end;
      end else if (Typeenvoie = 'M') or (TypeEnvoie='R') then
      begin
        Sujet := 'Modification de votre action '+IntToStr(QQ.findField('RAC_NUMACTION').AsInteger);
        Corps.Add('Votre action N°'+IntToStr(QQ.findField('RAC_NUMACTION').AsInteger)+' a été modifiée');
        Corps.Add('');
        Corps.Add('Pour information :');
        Corps.Add('Action :'+QQ.findField('RAC_LIBELLE').AsString +' pour le client '+QQ.findField('T_LIBELLE').AsString);
        Corps.Add('');
        if not InfoFrom then
        begin
          Corps.Add('prévue le :'+DateTimeToStr(QQ.findField('RAC_DATEACTION').AsDateTime)+' à '+TimeToStr(QQ.findField('RAC_HEUREACTION').AsDateTime));
          Corps.Add('d''une durée de :'+LibelleDuree(QQ.findField('RAC_DUREEACTION').Asinteger,false));
        end else
        begin
          Corps.Add('prévue le :'+DateTimeToStr(NewDate)+' à '+TimeToStr(NewDate));
          Corps.Add('d''une durée de :'+LibelleDuree(newdelai,false));
        end;
        Corps.Add('Adresse :');
        if QQ.findField('T_ADRESSE1').AsString <> '' then
        begin
          Corps.add('         '+QQ.findField('T_ADRESSE1').AsString);
        end;
        if QQ.findField('T_ADRESSE2').AsString <> '' then
        begin
          Corps.Add('         '+QQ.findField('T_ADRESSE2').AsString);
        end;
        if (QQ.findField('T_CODEPOSTAL').AsString <> '') or (QQ.findField('T_VILLE').AsString<>'') then
        begin
        	Corps.Add('         '+QQ.findField('T_CODEPOSTAL').AsString+' '+QQ.findField('T_VILLE').AsString);
        end;
        Corps.Add('');
        if not InfoFrom then
        begin
          Corps.Add('à été décalée au :'+DateToStr(NewDate)+' à '+TimeToStr(NewDate));
          Corps.Add('pour une durée de :'+LibelleDuree(newdelai,false));
        end else
        begin
          Corps.Add('à été décalée au :'+DateToStr(QQ.findField('RAC_DATEACTION').AsDateTime)+' à '+TimeToStr(QQ.findField('RAC_HEUREACTION').AsDateTime));
          Corps.Add('pour une durée de :'+LibelleDuree(QQ.findField('RAC_DUREEACTION').Asinteger,false));
        end;
      end else
      begin
      	exit; // non pris en charge
      end;

      ResultMailForm := AglMailForm(Sujet,Destinataire,copie,Corps,fichiers,false);
      if ResultMailForm = rmfOkButNotSend then SendMail(Sujet,Destinataire,Copie,Corps,fichiers,True,1, '', '');

    end;
  FINALLY
  	ferme (QQ);
    Corps.Free;
  End;
end;

function UpdateActionGRC (TobLigItem : TOB): boolean;
var MaRequete : string;
    NbHeure   : double;
    HeureDeb  : TTime;
    Auxiliaire: string;
    NumAction : Integer;
begin

	result  := true;

  if not IsModifiable (TobLigItem) then
  begin
    Result := true;
    exit;
  end;

  Auxiliaire      := TobLigItem.GetValue('BPL_TIERS');
  NumAction       := TobLigItem.GetValue('BPL_IDEVENT');

  if not GestionDateItem(TobLigItem) then Exit;

  HeureDeb := StrToTime(TobLigItem.GetValue('BPL_HEUREDEB'));
  NbHeure  := TobLigItem.GetValue('BPL_DUREE');

  if GetParamSocSecur ('SO_BTAVERTIRENMODIF',true) then
  begin
    EnvoieEltEmailFromGRC (Auxiliaire,NumAction,'M',TobLigItem.GetValue('BPL_DATEDEB'),round(Nbheure));
  end;

  Marequete:='UPDATE ACTIONS SET RAC_DATEACTION="'+USDateTime(TobLigItem.GetValue('BPL_DATEDEB'))+'",'+
  					 'RAC_HEUREACTION="'   + USDateTime(heureDeb)+'",' +
             'RAC_DUREEACTION='    + IntToStr(round(Nbheure))+
             'WHERE RAC_NUMACTION='+ IntToStr(NumAction)+
             '  AND RAC_AUXILIAIRE="'+Auxiliaire+'"';

  if ExecuteSql (Marequete) <= 0 then result := false;

end;

Function ChargeRequeteUnifie(P : THPlanningBTP; DateDebut, DateFin : TdateTime) : String;
begin

  Result := 'SELECT * FROM BTITEMS WHERE';
  //
  //Si on est sur un planning Famille Matériel
  if p.TypePlanning = 'PFM' then
    Result := Result + ' CODEFAMILLE="' + P.CodeOnglet + '" '
  else if p.typePlanning = 'PPA' then
    Result := Result + ' TYPEACTION="PMA" AND ACTION="' + P.CodeOnglet + '" '
  else
    Result := Result + ' TYPEACTION="INT" AND ACTION="' + P.CodeOnglet + '" ';

  if (pos(p.TypePlanning, 'PMA;PFM;PPA') = 0) and (not P.AppelsTraites) then
    Result := Result + 'AND (((AFFAIRE <> "") AND (CODEETAT="AFF")) OR (AFFAIRE="")) ';
  //
  Result := Result + '   ((DATEDEB>="'  + UsDateTime(Datedebut) + '" ';
  Result := Result + 'AND  DATEDEB<="'  + UsDateTime(Datefin)   + '") ';
  Result := Result + ' OR (DATEFIN>="'  + USDateTime(DateDebut) + '" ';
  Result := Result + 'AND  DATEFIN<="'  + USDateTime(DateFin)   + '") ';
  Result := Result + ' OR (DATEDEB<="'  + UsDateTime(DateDebut) + '" ';
  Result := Result + 'AND  DATEFIN>="'  + UsDateTime(DateFin)   + '")) ';
  Result := Result + 'AND ASSOSRES="X" ';
  //
  //
  Result := Result + 'ORDER BY DATEDEB';

end;

procedure ChargeItemPlanning(var R: RecordPlanning; var P: THPlanningBTP; T: TOB; DateEnCours: TDateTime);
var Qitem		  : TQuery;
    QIcone    : TQuery;
    TobEvent	: Tob;
    DateDebut	: TDateTime;
    datefin	  : TDateTime;
    MaRequete	: string;
    Civilite	: string;
    indice    : integer;
    Icone		  : integer;
    LibTypAct : String;
  	HeureDeb,HeureFin	: TDateTime;
    QQ : TQuery;
begin
  TobIcone := TOB.Create ('BTETAT',nil,-1);

  QQ := openSql('SELECT BTA_BTETAT,BTA_TYPEACTION,BTA_NUMEROICONE, BTA_LIBELLE FROM BTETAT '+
                'WHERE BTA_ASSOSRES   = "X" ORDER BY BTA_BTETAT,BTA_TYPEACTION',true,-1,'',true);
  TOBICone.LoadDetailDB('BTETAT','','',QQ,false);
  Ferme (QQ);
  TRY
    if P.tobres = nil then
    begin
      P.parent.Visible := false;
      Exit;
    end;

    P.ACtivate := False;
    if R.TobItems = nil then
    begin
      R.TobItems := Tob.Create('EVENEMENTS PLANNING', nil, -1);
    end else
    begin
      R.TobItems.ClearDetail;
      R.TobItems.InitValeurs(false);
    end;

    P.Activate      := False;
    P.IntervalDebut := DateEnCours - (T.GetValue('HPP_INTERVALLEDEB'));
    P.IntervalFin   := DateEnCours + (T.GetValue('HPP_INTERVALLEFIN'));
    P.DateOfStart   := DateEnCours;
    Civilite        := '';

    //Chargement des données (items) comprises dans l'intervalle choisi
    // Calcul de la date de debut et de fin
    //HeureDeb  := StrToTime('00:00:00');
    //HeureFin  := StrToTime('23:59:59');
    //Quite à prendre une heure autant utiliser celles du paramètre société...
    HeureDeb := GetParamSocSecur('SO_BTAMDEBUT', '08:30');
    HeureFin := GetParamSocSecur('SO_BTPMFIN',   '17:30');

    DateDebut := DateEncours - T.getvalue('HPP_INTERVALLEDEB');
    DateFin   := T.getvalue('HPP_INTERVALLEFIN') + DateEncours;

    DateDebut := Trunc(DateDebut); // + HeureDeb;
    DateFin   := Trunc(DateFin);   // + HeureFin;

    AfficheLFixe(P, T);

    //Lectures des tables évènement en fonction des type de Planning
    if (p.TypePlanning = 'PMA') then
      MaRequete := ChargeRequeteEventMat(P, DateDebut, DateFin)
    else if p.TypePlanning = 'PFM' then
      MaRequete := ChargeRequeteFamilleMat(P, DateDebut, DateFin)
    else if p.TypePlanning = 'PPA' then
      MaRequete := ChargeRequeteTypeAction(P, DateDebut, DateFin)
    else if p.TypePlanning = 'PRA' then
      MaRequete := ChargeRequeteRessourceChantier(P, DateDebut, DateFin)
    else if p.TypePlanning = 'PAR' then
      MaRequete := ChargeRequeteRessourceChantier(P, DateDebut, DateFin)
    else if p.TypePlanning = 'PFO' then
      MaRequete := ChargeRequeteRessourceChantier(P, DateDebut, DateFin)
    else
      MaRequete := ChargeRequeteIntervention(P, DateDebut, DateFin);

    QItem := OpenSQL(MaRequete, True,-1,'',true);

    TobEvent := Tob.Create('LES EVENEMENTS', nil, -1);
    TRY
      TobEvent.LoadDetailDB('BTITEMS', '', '', QItem, True);

      Ferme(QItem);
      P.TobItems := nil;

      if (T.getValue('HPP_FAMRES')='SAL') or (T.getValue('HPP_FAMRES')='') then
      begin
        if (T.GetBoolean('HPP_AFFABSSAL')) then
        begin
          MaRequete := ChargeRequeteAbsSalaries (P,DateDebut,DateFin);
          QItem := OpenSQL(MaRequete, True,-1,'',true);
          TobEvent.LoadDetailDB('BTITEMS', '', '', QItem, True);
          Ferme(Qitem);
        end;
        if (T.GetBoolean('HPP_AFFACTGRC')) then
        begin
          MaRequete := ChargeRequeteActionsGRC (P,DateDebut,DateFin);
          QItem := OpenSQL(MaRequete, True,-1,'',true);
          TobEvent.LoadDetailDB('BTITEMS', '', '', QItem, True);
          Ferme(Qitem);
        end;
      end;

      //On affiche les événements Parc/Matériel en plus de ceux du planning initial
      If pos(P.TypePlanning, 'PFM;PMA;PPA') = 0 then
      Begin
        if (T.GetBoolean('HPP_AFFEVTMAT')) then
        Begin
          MaRequete := ChargeRequeteEventMat(P, DateDebut, DateFin);
          QItem := OpenSQL(MaRequete, True,-1,'',true);
          TobEvent.LoadDetailDB('BTITEMS', '', '', QItem, True);
          Ferme(Qitem);
        end;
      end;
      //
      //On affiche les événements Chantier en plus de ceux du planning initial
      If pos(P.TypePlanning, 'PAR;PRA;PFO') = 0 then
      begin
        if (T.GetBoolean('HPP_AFFEVTCHA')) then
        Begin
          MaRequete := ChargeRequeteRessourceChantier(P, DateDebut, DateFin);
          QItem := OpenSQL(MaRequete, True,-1,'',true);
          TobEvent.LoadDetailDB('BTITEMS', '', '', QItem, True);
          Ferme(Qitem);
        end;
      end;
      //
      //On affiche les événement Intervention en plus de ceux du planning initial
      If pos(P.TypePlanning, 'PLA;PSF;PTA;PTR') = 0 then
      begin
        if (T.GetBoolean('HPP_AFFEVTINT')) then
        Begin
          MaRequete := ChargeRequeteIntervention(P, DateDebut, DateFin);
          QItem := OpenSQL(MaRequete, True,-1,'',true);
          TobEvent.LoadDetailDB('BTITEMS', '', '', QItem, True);
          Ferme(Qitem);
        end;
      end;

      //chargement de la table des items avec les informations des tables events...
      for Indice := 0 to  TobEvent.Detail.count -1 do
      begin
        TransformeEvenementEnItem(R.TobItems, TobEvent.detail[Indice]);
      end;

      if R.TOBItems.Detail.count > 1 then R.TOBItems.detail.Sort('BPL_DATEAFFDEB');
    FINALLY
      TobEvent.free;
      P.TobItems := R.TobItems;

      // Gestion et chargement des evenements (Salon, congrès,...)
      //ChargeEvent(R, P);

      P.DisplayOptionLier               := False;
      P.DisplayOptionSuppressionLiaison := False;
      P.DisplayOptionLiaison            := False;

      FINIMOVE;

      P.Activate := True;
    END;
  FINALLY
    TobIcone.free;
  END;

end;

function AnalyseChamp(Tobligne: TOB; ListeChamps: string; Formats: string): string;
var
  ChampEnCours: string;
  FormatEnCours: string;
  CharTmp: string;
  TypeChamp: string;
  ZHeure   : String;
  ZDate    : String;
begin
  result := '';

  while ListeChamps <> '' do
  begin
    ChampEncours := ReadTokenst(ListeChamps);
    FormatEnCours := ReadTokenst(Formats);
    if Copy(ChampEncours, 0, 1) <> '(' then
    begin
      if Tobligne.FieldExists(ChampEncours) then
        if (pos('HEURE', ChampEncours) <> 0) then
           Begin
           Zheure := TimeTostr(Tobligne.getvalue(ChampEnCours));
           Zheure := FormatDateTime('hh:mm', StrToTime(zheure));
           Result := Result + Zheure + ' ';
           end
        else if pos('DATE', ChampEncours) <> 0 then
          Result := Result + DateTimeTostr(Tobligne.getvalue(ChampEnCours)) + ' '
        else if pos('DUREE', ChampEncours)<> 0 then
          Result := Result + IntTostr(Tobligne.getvalue(ChampEnCours)) + ' '
        else if FormatEncours = '$' then
        begin
          CharTmp := AffecteDataType(ChampEncours);
          CharTmp := RechDom(CharTmp, Tobligne.getvalue(ChampEnCours), False);

          if CharTmp = 'Error' then
            CharTmp := '';

          Result := Result + CharTmp + ' ';
        end
        else
        begin
          typeChamp := ChampToType(ChampEnCours);
          if typechamp = 'INTEGER' then
            Result := Result + inttostr(Tobligne.getvalue(ChampEnCours)) + ' '
          else if typechamp = 'DOUBLE' then
            Result := Result + floattostr(Tobligne.getvalue(ChampEnCours)) + ' '
          else
            Result := Result + Tobligne.getvalue(ChampEnCours) + ' ';
        end;
    end
    else
      // on retire les paranthèses
      Result := Result + Copy(ChampEncours, 3, Length(ChampEncours) - 4) + ' ';
  end;

end;

function LibelleChamp(CodeTablette: string): string;
begin

  Result := 'ARS_LIBELLE';

  if CodeTablette = 'H21' then
    Result := 'HTR_TYPRES'
  else if CodeTablette = 'H22' then
    Result := 'HTR_ABREGE'
  else if CodeTablette = 'H23' then
    Result := 'HTR_TYPRES'
  else if CodeTablette = 'H24' then
    Result := 'HTR_RESSOURCE';

end;

function TabletteAssociee(NomTable: string): string;
begin

  Result := '';

  if NomTable = 'HTR_FAMRES' then
    Result := 'HRFAMRES'
  else if NomTable = 'ARS_TYPERESSOURCE' then
    Result := 'AFTTYPERESSOURCE';

end;


//permet de savoir si des lignes fixes sont attribuées au planning : Titre, Totaux, ...
procedure AfficheLFixe(P: THPlanningBTP; T: Tob);
var
  NbCol: integer;
  FontTo: Tfont;
  TauxOccupation: array of Hstring;
  Attribue: Boolean;
  FaireInit: Boolean;
  procedure InitTableau;
  begin
    FontTO.Name := T.Getvalue('HPP_FONTCOLONNE');
    NbCol := trunc(P.IntervalFin) - trunc(P.IntervalDebut) + 1;
    SetLength(TauxOccupation, NBcol);
  end;
begin

  FaireInit := True;

  FontTO := TFont.Create();

  try
    if (T.FieldExists('HPP_VISULIGNETRF')) and (T.GetValue('HPP_VISULIGNETRF') = 'X') then
    begin
      InitTableau;
      FaireInit := False;
      NumYield := 999;
    end;

    if (T.FieldExists('HPP_VISULIGNETO')) and (T.GetValue('HPP_VISULIGNETO') = 'X') then
    begin
      if FaireInit then InitTableau;
      // Gestion sur attribué ou dispo
      Attribue := T.Getvalue('HPP_MODEGESTION') = 'ATT';
      //ToutesResDispo(T.getvalue('HPP_FAMRES'), P.IntervalDebut, P.IntervalFin, TauxOccupation, Attribue);
      if (NumLigneTO = 0) then
        P.UpdateLigneFixe(NumLigneTO, TauxOccupation)
      else
        NumLigneTO := P.AjoutLigneFixe(FontTO, T.Getvalue('HPP_COULEURFOND'), TauxOccupation);
    end;

  finally
    FreeAndNil(FontTO);
  end;

end;

function AffecteDataType(NomChamp: string): string;
begin
  Result := '';
  Result := Get_Join(NomChamp);
end;

procedure MajParamplanning;
begin

  //Mise à jour des paramètres planning mode Graphique
  if (ExisteSQL('SELECT HPP_PARAMPLANNING FROM HRPARAMPLANNING WHERE HPP_MODEPLANNING=""')) then
     begin
     ExecuteSQL('UPDATE HRPARAMPLANNING SET HPP_MODEPLANNING="1DE" WHERE HPP_PLANNINGGRAPH="X"');
     ExecuteSQL('UPDATE HRPARAMPLANNING SET HPP_MODEPLANNING="4NR" WHERE HPP_PLANNINGGRAPH="-"');
     ExecuteSQL('UPDATE HRPARAMPLANNING SET HPP_MODEPLANNING="6NP" WHERE HPP_PLANNINGGRAPH=" "');
     end;

end;

(*
procedure RecupChamps(var ListeChamps: array of string; iTable: integer);
var iChamp: integer;
    Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
    FieldList : IEnumerator ;
    Field     : IFieldCOM ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  for iChamp := 1 to high(V_PGI.DeChamps[iTable]) do
    begin
    ListeChamps[iChamp] := V_PGI.DEChamps[iTable, iChamp].Nom;
    end;

  ListeChamps[Length(ListeChamps) - 1] := '';
  ListeChamps[Length(ListeChamps) - 2] := '';
  ListeChamps[Length(ListeChamps) - 3] := 'ICONETYPDOS';
  ListeChamps[Length(ListeChamps) - 4] := 'HINT';
  ListeChamps[Length(ListeChamps) - 5] := 'LIBELLE';

end;
*)

function ChargeStrByTablette(CodeTablette,ZoneARecuperer: String): String;
var StReq : String;
    QQ    : TQuery;
    TTab  : Tob;
    I			: Integer;
begin

    Result := '';

    StReq := 'SELECT ' + ZoneARecuperer + ' FROM CHOIXCOD WHERE CC_TYPE="' + CodeTablette + '"';

    // Lecture de la tablette passée en paramètre
    QQ := OpenSQL(StReq, True,-1,'',true);
    TTab := Tob.Create('Tob Tablette', nil, -1);
    TTab.LoadDetailDB('CHOIXCOD', '', '', QQ, False);
    Ferme(QQ);

    for I := 0 to TTab.Detail.Count-1 do
    Begin
        Result := Result + TTab.detail[i].GetValue(ZoneARecuperer) + ',';
    end;

    if result = '' then 
       Result := '*'
    else
       Result := Copy(result, 0, length(result)-1);

end;

function GestionDateFinPourModif (DateDeFin : TdateTime) : TdateTime;
Var DateFinE	: TDateTime;
  	HeureFin	: TDateTime;
    Delta : integer;
Begin

  // Gestion de la date et l'heure
  DateFinE := StrToDate(DateToStr(DateDeFin));
  HeureFin := StrTotime(TimetoStr(DateDeFin));

  if Cadencement = '001' then // gestion par 1/4 Heure
  begin
    HeureFin := IncMinute (HeureFin, 15);
  end else if Cadencement = '002' then // gestion par 1/2 Heure
  begin
    HeureFin := IncMinute (HeureFin,30);
  end else if Cadencement = '003' then // gestion par heure
  begin
    HeureFin := IncHour (HeureFin, 1);
  end else if Cadencement = '004' then // gestion par 1/2 Journée
  Begin
    if heureFin > GetFinMatinee then
    begin
    	Delta :=  MinutesBetween(heureFin,GetFinApresMidi);
    end else
    begin
    	Delta := MinutesBetWeen(HeureFin,GetFinMatinee);
    end;
    HeureFin := incminute (HeureFin,Delta);
  end else if Cadencement = '005' then // gestion par Journée
  begin
  	// pas de decalage dans ce cas la
  end else if Cadencement = '008' then // gestion par Période
  begin
     {TobTmp.PutValue('BEP_DATEFIN', TobTmp.GetValue('BEP_DATEFIN'));}
  end;
  result := DateFinE + heureFin;
end;

function GestionDateFinPourEnreg (DateDeFin : TdateTime) : TdateTime;
Var
  	DateFinE	: TDateTime;
  	HeureFin	: TDateTime;
    Delta : integer;
Begin

  // Gestion de la date et l'heure (Decalage pour la gestion de l'affichage)
  DateFinE := StrToDate(DateToStr(DateDeFin));
  HeureFin := StrTotime(TimetoStr(DateDeFin));
  //
  if Cadencement = '001' then // gestion par 1/4 Heure
  begin
     HeureFin := IncMinute (HeureFin, -15);
  end else if Cadencement = '002' then // gestion par 1/2 Heure
  begin
     HeureFin := IncMinute (HeureFin,-30);
  end else if Cadencement = '003' then // gestion par heure
  begin
     HeureFin := IncHour (HeureFin, - 1);
  end else if Cadencement = '004' then // gestion par 1/2 Journée
  Begin
    if heureFin > GetFinMatinee then
    begin
    	Delta :=  MinutesBetween(heureFin,GetFinApresMidi);
    end else
    begin
    	Delta := MinutesBetWeen(HeureFin,GetFinMatinee);
    end;
     HeureFin := incminute (HeureFin,Delta*(-1));
  end else if Cadencement = '005' then // gestion par Journée
  begin
  	// pas de decalage dans ce cas la
  end else if Cadencement = '008' then // gestion par Période
  begin
     {TobTmp.PutValue('BEP_DATEFIN', TobTmp.GetValue('BEP_DATEFIN'));}
  end;
  result := DateFinE + heureFin;
end;

function CalculDateDebut (DateDebut : TdateTime) : TdateTime;
Var DateDebE	  : TDateTime;
  	HeureDeb    : TdateTime;
    HeureDebAM  : TdateTime;
    HeureFinAM  : TDateTime;
    HeureDebPM  : Tdatetime;
    HeureFinPM  : TDateTime;
    HeureDebJr  : TDateTime;
Begin

  // Gestion de la date et l'heure (Decalage pour la gestion de l'affichage)
  DateDebE    := StrToDate(DateToStr(DateDebut));
  HeureDeb    := StrTotime(TimetoStr(DateDebut));
  //Si jamais il y a un calendrier sur la ressource celui-ci ne sera pas pris en compte ici...
  HeureDebAm  := GetDebutMatinee;
  HeureFinAm  := GetFinMatinee;
  HeureDebPm  := GetDebutApresMidi;
  HeurefinPm  := GetfinApresMidi;
  HeureDebJr  := HeureDebAm;


  if (pos(Cadencement,'001;002;003')>0) then
  begin
  	 if (heureDeb < heureDebAM) then HeureDeb := heureDebAM;
     if (heureDeb > heureFinPM) then HeureDeb := heureFinPM;
     //if (heureDeb > HeureFinAM) and (heureDeb < HeureDebPm) then heureDeb := HeureDebPm;
  end
  else if cadencement = '005' then  // gestion par journée
  begin
    HeureDeb := HeureDebJr;
  end
  else if Cadencement = '004' then // gestion par 1/2 Journée
  Begin
    if heureDeb > HeureFinAm then
    begin
    	heureDeb := heureDebpm;
    end
    else
    begin
    	heureDeb := HeureDebAm;
    end;
  end;

  result := DateDebE + HeureDeb;

end;

function CalculDateFin (DateFin : TdateTime) : TdateTime;
Var DateFinE	  : TDateTime;
  	HeureFin    : TdateTime;
    HeureDebAM  : TdateTime;
    HeureDebPM  : TdateTime;
    HeureFinPM  : Tdatetime;
    HeureFinAM  : TDateTime;
    HeureFinJr  : TDateTime;
Begin

  // Gestion de la date et l'heure (Decalage pour la gestion de l'affichage)
  DateFinE := StrToDate(DateToStr(DateFin));
  HeureFin := StrTotime(TimetoStr(DateFin));

  //Si jamais il y a un calendrier sur la ressource celui-ci ne sera pas pris en compte ici...
  HeureDebAm := GetDebutMatinee;
  HeureFINAm := GetFinMatinee;
  HeureDebPm := GetDebutApresMidi;
  HeureFINPM := GetFinApresMidi;
  
  HeureFINJr := HeureFINPM;

  //
  if (pos(Cadencement,'001;002;003')>0) then
  begin
    if (HeureFin < HeureDebAm) Then Heurefin := HeureDebAm;
    if (heureFin > HeureFinPM) Then HeureFin := HeureFinPM;
    //if (heureFin < HeureDebPM) and (heureFin > HeureFinAm) then
    //  HeureFin := HeureFinAM
    //else
    //Begin
    if (Cadencement='001') then
      HeureFin  := HeureFin + (15 *(1/1440))
    else if  (Cadencement='002') then
      HeureFin  := HeureFin + (30 * (1/1440))
    else if  (Cadencement='003') then
      HeureFin  := HeureFin + 1/24;
    //end;
  end
  else if cadencement = '005' then  // gestion par journée
  begin
    Heurefin := HeureFinJr;
  end
  else if Cadencement = '004' then // gestion par 1/2 Journée
  Begin
    if heureFin > HeureFinAM then
    begin
    	heureFin := heureFinPM;
    end
    else
    begin
    	heureFin := HeureFinAm;
    end;
  end;

  result := DateFinE + HeureFin;

end;


function ControleHoraires (var DateDebEvt,DateFinEvt : TdateTime; ModeOption: string; Duree : double; CodeRessource : string) : boolean;
Var DateDeb,DateFin,HeureDeb,heureFin,HeureDebAM,HeureFinAm,heureDebPm,heureFinPM	: TDateTime;
begin
	result := true;
  //
  DateDeb := StrToDate(DateToStr(DateDebEvt));
  DateFin := StrToDate(DateToStr(DateFinEvt));
  HeureDeb := StrToTime(TimeToStr(DateDebEvt));
  HeureFin := StrToTime(TimeToStr(DateFinEvt));
  //
  if IsDateClosed (DateDeb) then
  begin
  	pgiInfo ('Attention : La date de début est un jour non travaillé');
    result := false;
    exit;
  end;

  if ModeOption <> 'D' then
  begin
    if IsDateClosed (DateFin) then
    begin
      pgiInfo ('Attention : La date de fin est un jour non travaillé');
      result := false;
      exit;
    end;
  end;
  //
  HeureDebAm := GetDebutMatinee;
  HeureDebPm := GetDebutApresMidi;
  HeureFinAm := GetFinMatinee;
  HeureFinPm := GetFinApresMidi;

  if (HeureDeb < heureDebAm) or
  	 (HeureDeb > heureFinPm) or
     ((HeureDeb > HeureFinAm) and
     (HeureDeb < heureDebPM)) then
  begin
  	if (HeureDeb < heureDebAM) then
    begin
    	heureDeb := HeureDebAM;
      DateDebEvt := DateDeb + heureDeb;
    end else if (heureDeb < HeureDebPM ) and (heureDeb > HeureFinAm) then
    begin
    	heureDeb := HeureDebPM;
      DateDebEvt := DateDeb + heureDeb;
    end else if (heureDeb > HeureFinPM ) then
    begin
    	heureDeb := HeureDebAM;
      DateDebEvt := IncDay(DateDeb,1) + heureDeb;
      DateDeb := StrToDate(DateToStr(DateDebEvt));
    end;
  end;
  //
  if ModeOption = 'D' then
  begin
  	DateFinEvt := AjouteDuree (DateDebEvt,round(Duree));
  	DateFin := StrToDate(DateToStr(DateFinEvt));
  	HeureFin := StrToTime(TimeToStr(DateFinEvt));
  end;
  //
  if (HeureFin < heureDebAm) or
  	 (HeureFin > heureFinPm) or
     ((HeureFin > HeureFinAm) and
     (HeureFin <= heureDebPM)) then
  begin
  	if (heureFin > heureFinAm) and (HeureFin <= heureDebPM) then
    begin
    	heureFin := HeureFinAM;
      DateFinEvt := DateFin + heureFin;
    end else if (heureFin < heureDebAm) then
    begin
    	heureFin := HeureDebAM;
      DateFinEvt := IncDay (DateFin,-1) + heureFin;
    end else if (heureFin > heureFinPM) then
    begin
    	heureFin := HeureDebAM;
      DateFinEvt := IncDay (DateFin,1) + heureFin;
    end;
  end;

end;

function GetDescriptifCourt ( TOBD : TOB) : string;
var
  Memo: Tmemo;
  Panel: TPanel;
begin
  Panel := TPanel.Create(nil);
  try
    Panel.Visible := False;
    Panel.ParentWindow := GetDesktopWindow;
    Memo := Tmemo.Create(Panel);
    Memo.Parent := Panel;
    Memo.Visible := False;
    memo.text :=  TOBD.GetValue('BEP_BLOCNOTE');
    result := memo.lines.Strings [0]+'#13#10';
    if (memo.lines.Strings [1] <> '') and (memo.lines.Strings [1]<>'#13#10') then
    begin
    	result := result + memo.lines.Strings [1];
    end;
    if (memo.lines.Strings [2] <> '') and (memo.lines.Strings [2]<>'#13#10') then
    begin
    	result := result +'#13#10...';
    end;
    Memo.Free;
  finally
    Panel.Free;
  end

end;

function GetDescriptifTotal ( TOBD : TOB) : string;
var
  Memo: Tmemo;
  Panel: TPanel;
begin
  Panel := TPanel.Create(nil);
  try
    Panel.Visible := False;
    Panel.ParentWindow := GetDesktopWindow;
    Memo := Tmemo.Create(Panel);
    Memo.Parent := Panel;
    Memo.Visible := False;
    memo.text :=  TOBD.GetValue('BEP_BLOCNOTE');
    result := memo.lines.Strings [0]+'#13#10';
    if (memo.lines.Strings [1] <> '') and (memo.lines.Strings [1]<>'#13#10') then
    begin
    	result := result + memo.lines.Strings [1];
    end;
    if (memo.lines.Strings [2] <> '') and (memo.lines.Strings [2]<>'#13#10') then
    begin
    	result := result +'#13#10...';
    end;
    Memo.Free;
  finally
    Panel.Free;
  end

end;


procedure EnvoieEltEmailFromAction ( TOBElt : TOB; TypeEnvoie : string; NewDate : TdateTime;newdelai : integer; InfoFrom : boolean=false; NewRess : string='');
var maRequete : string;
		QQ : Tquery;
    Sujet : HString;
    Corps : HTStringList;
    ResultMailForm : TResultMailForm;
    copie,fichiers,Destinataire : string;
    Memo: Tmemo;
    Panel: TPanel;
    indice : integer;
    TOBL : TOB;
begin
	Corps := hTStringList.Create;
  Marequete:='SELECT ARS_EMAIL FROM RESSOURCE '+
             'WHERE ARS_RESSOURCE="'+TOBELT.getValue('BEP_RESSOURCE')+'"';
  Panel := TPanel.Create(nil);
  Panel.Visible := False;
  Panel.ParentWindow := GetDesktopWindow;
  Memo := Tmemo.Create(Panel);
  Memo.Parent := Panel;
  Memo.Visible := False;
  memo.text :=  TOBElt.GetValue('BEP_BLOCNOTE');
  TRY
  	QQ := OpenSql (MaRequete,True,1,'',true);
    if not QQ.eof then
    begin
      if QQ.findField ('ARS_EMAIL').AsString ='' then exit;
      Destinataire := QQ.findField ('ARS_EMAIL').AsString;
			if TypeEnvoie = 'C' then   // création
      begin
        Sujet := 'Nouveau(nouvelle) '+Rechdom('BTTYPEACTION',TOBELt.getValue('BEP_BTETAT'),false);
        Corps.add ('Prévue le : '+DateToStr(TOBElt.getValue('BEP_DATEDEB'))+' à '+ TimeToStr(TOBElt.getValue('BEP_HEUREDEB')));
        Corps.Add('d''une durée de :'+LibelleDuree(TOBElt.getValue('BEP_DUREE')));
        Corps.Add('');
        Corps.add ('détail :');
        for indice := 0 to memo.Lines.Count -1 do
        begin
        	Corps.Add(memo.lines.Strings [Indice]);
        end;
        Corps.Add('');
      end else if TypeEnvoie = 'S' then // suppression
      begin
        Sujet := 'Annulation de votre '+Rechdom('BTTYPEACTION',TOBELt.getValue('BEP_BTETAT'),false);
        Corps.Add('Votre '+Rechdom('BTTYPEACTION',TOBELt.getValue('BEP_BTETAT'),false)+' a été annulé(e)');
        Corps.add('Prévue le : '+DateToStr(TOBElt.getValue('BEP_DATEDEB'))+' à '+ TimeToStr(TOBElt.getValue('BEP_HEUREDEB'))+
        				 'est annulée');
      end else if (Typeenvoie = 'M') or (TypeEnvoie='R') then // etirement, reduction, modification, déplacement
      begin
      	if (NewRess <> '') and (NewRess <> TOBELt.getValue('BEP_RESSOURCE')) then
        begin
        	if InfoFrom then
          begin
          	// préparation des infos pour l'ancienne ressource -> Annulation de son action
          	TOBL := TOB.Create ('BTEVENPLAN',nil,-1);
            TRY
              TOBL.dupliquer (TOBElt,false,true);
              TOBL.putValue('BEP_DATEDEB',newDate);
              TOBL.putValue('BEP_HEUREDEB',StrToTime(TimeToStr(newDate)));
              TOBL.putValue('BEP_RESSOURCE',NewRess);
              EnvoieEltEmailFromAction (TOBL,'S',Idate1900,0); // envoie de l'annulation
              EnvoieEltEmailFromAction (TOBElt,'C',Idate1900,0); // envoie de l'info a la nouvelle ressource
            FINALLY
            	TOBL.free;
            END;
          end else
          begin
          	// préparation des infos pour la nouvelle ressource
          	TOBL := TOB.Create ('BTEVENPLAN',nil,-1);
            TRY
              TOBL.dupliquer (TOBElt,false,true);
              TOBL.putValue('BEP_DATEDEB',newDate);
              TOBL.putValue('BEP_HEUREDEB',StrToTime(TimeToStr(newDate)));
              TOBL.putValue('BEP_RESSOURCE',NewRess);
              EnvoieEltEmailFromAction (TOBL,'C',Idate1900,0); // envoie de l'info a la nouvelle ressource
              EnvoieEltEmailFromAction (TOBElt,'S',Idate1900,0); // envoie de l'annulation
            FINALLY
            	TOBL.free;
            END;
          end;
          exit;
        end else
        begin
          Sujet := 'Modification de '+Rechdom('BTTYPEACTION',TOBELt.getValue('BEP_BTETAT'),false);
          Corps.Add('Votre '+Rechdom('BTTYPEACTION',TOBELt.getValue('BEP_BTETAT'),false)+' a été modifié(e)');
          Corps.Add('');
          Corps.Add('Pour information :');
          for indice := 0 to memo.Lines.Count -1 do
          begin
            Corps.Add(memo.lines.Strings [Indice]);
          end;
          Corps.Add('');
          if not InfoFrom then
          begin
            Corps.add('Prévue le : '+DateToStr(TOBElt.getValue('BEP_DATEDEB'))+' à '+ TimeToStr(TOBElt.getValue('BEP_HEUREDEB')));
            Corps.Add('d''une durée de :'+LibelleDuree(TOBElt.getValue('BEP_DUREE')));
          end else
          begin
            Corps.Add('prévue le :'+DateTimeToStr(NewDate)+' à '+TimeToStr(NewDate));
            Corps.Add('d''une durée de :'+LibelleDuree(newdelai,false));
          end;
          Corps.Add('');
          if not InfoFrom then
          begin
            Corps.Add('à été décalée au :'+DateToStr(NewDate)+' à '+TimeToStr(NewDate));
            Corps.Add('pour une durée de :'+LibelleDuree(newdelai,false));
          end else
          begin
            Corps.add('à été décalée au :'+DateToStr(TOBElt.getValue('BEP_DATEDEB'))+' à '+ TimeToStr(TOBElt.getValue('BEP_HEUREDEB')));
            Corps.Add('pour une durée de :'+LibelleDuree(TOBElt.getValue('BEP_DUREE')));
          end;
        end;
      end else exit;

      ResultMailForm := AglMailForm(Sujet,Destinataire,copie,Corps,fichiers,false);
      if ResultMailForm = rmfOkButNotSend then SendMail(Sujet,Destinataire,Copie,Corps,fichiers,True,1, '', '');

    end;
  FINALLY
    Panel.Free;
  	ferme (QQ);
    Corps.Free;
  End;
end;

procedure EnvoieEmailFromIntervention (Sender : TObject; TobItem : TOB);
var EnvoiMail   : TGestionMail;
    Memo        : THRichEditOLE;
    descriptif  : TMemo;
    iInd        : Integer;
    LineSave    : String;
begin

  EnvoiMail := TGestionMail.Create(Sender);
  //
  EnvoiMail.Sujet := 'Confirmation intervention du ' + DateToStr(TobItem.GetValue('BPL_DATEDEB')) + ' à ' + TimeToStr(TobItem.GetValue('BPL_HEUREDEB'));

  EnvoiMail.Corps := hTStringList.Create;
  EnvoiMail.Corps.Clear ; //Descriptif de l'affaire

  Memo := THRichEditOLE.create (Application);
  Memo.Name := 'CorpsMail';
  Memo.Parent := Application.MainForm;
  StringToRich(Memo, TobItem.GetValue('BPL_DESCRIPTIF'));
  For iInd := 0 to Memo.Lines.count do
  begin
    LineSave  := Memo.lines[iInd];
    EnvoiMail.corps.Add(LineSave);
  end;
  FreeAndNil(Memo);

  EnvoiMail.Copie         := '';
  EnvoiMail.TypeContact   := 'RES';
  EnvoiMail.Fournisseur   := '';
  EnvoiMail.FichierSource := '';
  EnvoiMail.FichierTempo  := '';
  EnvoiMail.Fichiers      := '';
  EnvoiMail.TypeDoc       := '';
  EnvoiMail.Tiers         := '';

  EnvoiMail.Contact       := '';
  EnvoiMail.Destinataire  := TobItem.getValue('BPL_EMAIL');
  EnvoiMail.QualifMail    := 'PLA';
  EnvoiMail.TobRapport    := TobItem;

  if EnvoiMail.QualifMail = '' then
    EnvoiMail.GestionParam  := False
  else
    EnvoiMail.GestionParam  := True;

  EnvoiMail.CopySourceSurTempo;

  EnvoiMail.AppelEnvoiMail;

  FreeAndNil(EnvoiMail);

  {*
  TRY
			if TypeEnvoie = 'C' then   // création
      begin
        Sujet := 'Nouveau(nouvelle) '+Rechdom('BTTYPEACTION',TOBELt.getValue('BEP_BTETAT'),false);
        Corps.add ('Prévue le : '+DateToStr(TOBElt.getValue('BEP_DATEDEB'))+' à '+ TimeToStr(TOBElt.getValue('BEP_HEUREDEB')));
        Corps.Add('d''une durée de :'+LibelleDuree(TOBElt.getValue('BEP_DUREE')));
        Corps.Add('');
        Corps.add ('détail :');
        for indice := 0 to memo.Lines.Count -1 do
        begin
        	Corps.Add(memo.lines.Strings [Indice]);
        end;
        Corps.Add('');
      end else if TypeEnvoie = 'S' then // suppression
      begin
        Sujet := 'Annulation de votre '+Rechdom('BTTYPEACTION',TOBELt.getValue('BEP_BTETAT'),false);
        Corps.Add('Votre '+Rechdom('BTTYPEACTION',TOBELt.getValue('BEP_BTETAT'),false));
        Corps.add('Prévue le : '+DateToStr(TOBElt.getValue('BEP_DATEDEB'))+' à '+ TimeToStr(TOBElt.getValue('BEP_HEUREDEB')));
        Corps.Add('a été annulé(e)');
        Corps.add ('détail :');
        for indice := 0 to memo.Lines.Count -1 do
        begin
        	Corps.Add(memo.lines.Strings [Indice]);
        end;
        Corps.Add('');
      end else if (Typeenvoie = 'M') or (TypeEnvoie='R') then // etirement, reduction, modification, déplacement
      begin
      	if (NewRess <> '') and (NewRess <> TOBELt.getValue('BEP_RESSOURCE')) then
        begin
        	if InfoFrom then
          begin
          	// préparation des infos pour l'ancienne ressource -> Annulation de son action
          	TOBL := TOB.Create ('BTEVENPLAN',nil,-1);
            TRY
              TOBL.dupliquer (TOBElt,false,true);
              TOBL.putValue('BEP_DATEDEB',newDate);
              TOBL.putValue('BEP_HEUREDEB',StrToTime(TimeToStr(newDate)));
              TOBL.putValue('BEP_RESSOURCE',NewRess);
              EnvoieEltEmailFromAction (TOBL,'S',Idate1900,0); // envoie de l'annulation
              EnvoieEltEmailFromAction (TOBElt,'C',Idate1900,0); // envoie de l'info a la nouvelle ressource
            FINALLY
            	TOBL.free;
            END;
          end else
          begin
          	// préparation des infos pour la nouvelle ressource
          	TOBL := TOB.Create ('BTEVENPLAN',nil,-1);
            TRY
              TOBL.dupliquer (TOBElt,false,true);
              TOBL.putValue('BEP_DATEDEB',newDate);
              TOBL.putValue('BEP_HEUREDEB',StrToTime(TimeToStr(newDate)));
              TOBL.putValue('BEP_RESSOURCE',NewRess);
              EnvoieEltEmailFromAction (TOBL,'C',Idate1900,0); // envoie de l'info a la nouvelle ressource
              EnvoieEltEmailFromAction (TOBElt,'S',Idate1900,0); // envoie de l'annulation
            FINALLY
            	TOBL.free;
            END;
          end;
          exit;
        end else
        begin
          Sujet := 'Modification de '+Rechdom('BTTYPEACTION',TOBELt.getValue('BEP_BTETAT'),false);
          Corps.Add('Votre '+Rechdom('BTTYPEACTION',TOBELt.getValue('BEP_BTETAT'),false)+' a été modifié(e)');
          Corps.Add('');
          Corps.Add('Pour information :');
          for indice := 0 to memo.Lines.Count -1 do
          begin
            Corps.Add(memo.lines.Strings [Indice]);
          end;
          Corps.Add('');
          if not InfoFrom then
          begin
            Corps.add('Prévue le : '+DateToStr(TOBElt.getValue('BEP_DATEDEB'))+' à '+ TimeToStr(TOBElt.getValue('BEP_HEUREDEB')));
            Corps.Add('d''une durée de :'+LibelleDuree(TOBElt.getValue('BEP_DUREE')));
          end else
          begin
            Corps.Add('prévue le :'+DateTimeToStr(NewDate)+' à '+TimeToStr(NewDate));
            Corps.Add('d''une durée de :'+LibelleDuree(newdelai,false));
          end;
          Corps.Add('');
          if not InfoFrom then
          begin
            Corps.Add('à été décalée au :'+DateToStr(NewDate)+' à '+TimeToStr(NewDate));
            Corps.Add('pour une durée de :'+LibelleDuree(newdelai,false));
          end else
          begin
            Corps.add('à été décalée au :'+DateToStr(TOBElt.getValue('BEP_DATEDEB'))+' à '+ TimeToStr(TOBElt.getValue('BEP_HEUREDEB')));
            Corps.Add('pour une durée de :'+LibelleDuree(TOBElt.getValue('BEP_DUREE')));
          end;
        end;
      end else exit;

      ResultMailForm := AglMailForm(Sujet,Destinataire,copie,Corps,fichiers,false);
      if ResultMailForm = rmfOkButNotSend then SendMail(Sujet,Destinataire,Copie,Corps,fichiers,True,1, '', '');

  FINALLY
    Panel.Free;
  	ferme (QQ);
    Corps.Free;
  End;
  *}


end;

procedure AddChampsSupItems (TOBItem : TOB);
begin
  TOBItem.AddChampSupValeur('MODIFIABLE',   '-');

  TOBItem.AddChampSupValeur('RESPRINCIPALE','-');
  TOBItem.AddChampSupValeur('EQUIPESEP',    '-');
  TOBItem.AddChampSupValeur('OBLIGATOIRE',  'X');
  TOBItem.AddChampSupValeur('HEURETRAV',    '-');
  TOBItem.AddChampSupValeur('CREAPLANNING', '-');
  //
  TOBItem.AddChampSupValeur('LIBELLE',      '');
  TOBItem.AddChampSupValeur('HINT',         '');
  TOBItem.AddChampSupValeur('ICONETYPDOS',  '');
  TOBItem.AddChampSupValeur('AUXILIAIRE',   '');
  TOBItem.AddChampSupValeur('BLOCNOTE',     '');
end;



{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 10/02/2016
Modifié le ... :   /  /
Description .. : Gestion de l'affichage des items dans le corps du planning
Suite ........ : Chargement de la tob commune (BTEVENTPLAN)
Suite ........ : Chargement des zones supp.
Mots clefs ... :
*****************************************************************}
procedure  TransformeEvenementEnItem(TobItems, TOBEvent : TOB);
Var DateD : TDateTime;
    DateF : TDateTime;
    HeureD: TDateTime;
    Heuref: TDateTime;
    Duree : Integer;
    DD,DF : TdateTime;
    TobItem : TOB;
    Lib   : String;
begin

 	TOBItem := TOB.Create ('BTEVENEMENTPLA',  TobItems, -1);

  //A revoir le moment venu...
  AddChampsSupItems (TOBItem);

  TOBItem.PutValue('AUXILIAIRE',   TOBEvent.GetValue('AUXILIAIRE'));
  TOBItem.PutValue('BLOCNOTE',     TOBEvent.GetString('BLOCNOTE'));

  TOBItem.PutValue('BPL_ORIGINEITEM',       TOBEvent.GetString('ORIGINEITEM'));
  TOBItem.PutValue('BPL_IDEVENT',           TOBEvent.GetString('IDEVENT'));
  TOBItem.putValue('BPL_AFFAIRE',           TOBEvent.GetString('CODEAFFAIRE'));
  TOBItem.putValue('BPL_LIBAFFAIRE',        TOBEvent.GetString('LIBAFFAIRE'));
  TOBItem.putValue('BPL_TIERS',             TOBEvent.GetString('CODETIERS'));
  TOBItem.putValue('BPL_NOMTIERS',          TOBEvent.GetString('LIBTIERS'));
  TOBItem.putValue('BPL_JURIDIQUE',         TOBEvent.GetString('NOMJURIDIQUE'));
  TOBItem.putValue('BPL_NOM1',              TOBEvent.GetString('NOM1'));
  TOBItem.putValue('BPL_NOM2',              TOBEvent.GetString('NOM2'));
  TOBItem.putValue('BPL_ADRESSE1',          TOBEvent.GetString('ADR1'));
  TOBItem.putValue('BPL_ADRESSE2',          TOBEvent.GetString('ADR2'));
  TOBItem.putValue('BPL_ADRESSE3',          TOBEvent.GetString('ADR3'));
  TOBItem.putValue('BPL_CODEPOSTAL',        TOBEvent.GetString('CP'));
  TOBItem.putValue('BPL_VILLE',             TOBEvent.GetString('VILLE'));


  //FV1 - 24/04/2017 : FS#2509 - TEAM RESEAUX - Les plannings ne sont plus utilisables et bloquent l'application
  Try
    HeureD := StrToTime(TOBEvent.GetValue('HEUREDEB'));
  Except
    HeureD := GetDebutMatinee;
  end;

  Try
    Heuref := StrToTime(TOBEvent.GetValue('HEUREFIN'));
  except
    Heuref := GetFinJournee;
  end;


  TOBItem.putValue('BPL_DATEDEB',           TOBEvent.GetValue('DATEDEB'));
  TOBItem.putValue('BPL_HEUREDEB',          HeureD); //TOBEvent.GetValue('HEUREDEB'));
  TOBItem.putValue('BPL_DATEFIN',           TOBEvent.GetValue('DATEFIN'));
  TOBItem.putValue('BPL_HEUREFIN',          Heuref); //TOBEvent.GetValue('HEUREFIN'));

  TOBItem.putvalue('BPL_DUREE',             TOBEvent.GetValue('DUREE'));

  DateD   := StrToDate(DateToStr(TOBEvent.getValue('DATEDEB')));
  DateF   := StrToDate(DateToStr(TOBEvent.getValue('DATEFIN')));

  //chargement des événements absences de la paye pour les salariés dans la période
  if TOBEvent.GetString('ORIGINEITEM') = 'ABSPAIE' then
  begin
    //DateD := TOBEvent.getValue('PCN_DATEDEBUTABS');
    if TOBEvent.getValue('PCN_DEBUTDJ')='MAT' then
      heureD := GetDebutMatinee
    else
      heureD := GetDebutApresMidi;
    if TOBEvent.getValue('PCN_FINDJ')='MAT' then
      heureF := GetFinMatinee
    else
      heureF := GetFinApresMidi;
    //
    TOBItem.putValue('BPL_HEUREDEB',     HeureD);
    TOBItem.putValue('BPL_HEUREFIN',     HeureF);
    TobEvent.putvalue('BPL_HEUREAFFDEB', HeureD);
    TobEvent.putvalue('BPL_HEUREAFFFIN', HeureF);
    //
    TobEvent.putvalue('LIBACTION',   RechDom('PGTYPEMVT',TOBEvent.GetString('PCN_TYPEMVT'),False));
    TOBEvent.PutValue('LIBEVENEMENT',       TOBEvent.getValue('LIBACTION') + '#13#10' + RechDom('PGMOTIFABSENCE', TOBEvent.getValue('PCN_TYPECONGE'),false));
  end
  //chargement des événements Action GRC
  else if TOBEvent.GetString('ORIGINEITEM') = 'ACT-GRC' then
  begin
    if TOBEvent.GetValue('CODEETAT')='PRE' then Tobitem.putValue('MODIFIABLE','X');

    heureD := StrToTime(TimeToStr(TOBEvent.getValue('HEUREDEB')));

    Duree  := TOBEvent.getValue('DUREE');

    DateF  := AjouteDuree(DateD+heureD , Duree) ;
    HeureF :=StrToTime(TimeToStr(DateF));

    TOBItem.putValue('BPL_DATEFIN',      DateF);
    TOBItem.putValue('BPL_HEUREFIN',     HeureF);

    TobEvent.putvalue('BPL_DATEAFFDEB',  TOBEvent.getValue('DATEDEB'));
    TobEvent.putvalue('BPL_DATEAFFFIN',  DateF);
    TobEvent.putvalue('BPL_HEUREAFFDEB', TOBEvent.getValue('HEUREDEB'));
    TobEvent.putvalue('BPL_HEUREAFFFIN', HeureF);
    //
    TobEvent.putvalue('LIBACTION',   rechDom('RTTYPEACTIONALL', TOBEvent.getValue('RAC_TYPEACTION'),false));
    TobEvent.putvalue('LIBEVENEMENT',    TOBEvent.getValue('LIBACTION') + '#13#10' + TOBEvent.getValue('LIBTIERS') + ' (' + TOBEvent.getValue('VILLE') + ')#13#10' + TOBEvent.getValue('LIBACTION'));
  end
  else
  begin
    HeureD  := GetDebutMatinee;
    HeureF  := GetFinApresMidi;
  end;

  TOBItem.putValue('BPL_BTETAT',            TOBEvent.GetString('ACTION'));
  TOBItem.putValue('BPL_LIBACTION',         TOBEvent.GetString('LIBACTION'));
  TOBItem.putValue('BPL_GEREPLANNING',      TOBEvent.GetString('ASSOSRES'));

  TOBItem.putValue('BPL_RESSOURCE',         TOBEvent.GetString('CODERESSOURCE'));
  TOBItem.putValue('BPL_NOMRESSOURCE',      TOBEvent.GetString('LIBRESSOURCE'));
  TOBItem.putValue('BPL_EQUIPERESS',        TOBEvent.GetString('EQUIPE'));

  if TOBItem.GetString('BPL_RESSOURCE') = '' then
  Begin
    TOBItem.putValue('BPL_RESSOURCE',       TOBEvent.GetString('MATERIEL'));
    TOBItem.putValue('BPL_NOMRESSOURCE',    TOBEvent.GetString('LIBMATERIEL'));
    TOBItem.putValue('BPL_EQUIPERESS',      '');
  end;

  TOBItem.PutValue('BPL_MATERIEL',          TOBEvent.GetString('MATERIEL'));
  TOBItem.PutValue('BPL_LIBMATERIEL',       TOBEvent.GetString('LIBMATERIEL'));

  if TOBItem.GetString('BPL_MATERIEL') = '' then
  Begin
    TOBItem.putValue('BPL_RESSOURCE',       TOBEvent.GetString('CODERESSOURCE'));
    TOBItem.putValue('BPL_NOMRESSOURCE',    TOBEvent.GetString('LIBRESSOURCE'));
    TOBItem.putValue('BPL_EQUIPERESS',      TOBEvent.GetString('EQUIPE'));
  end;

  TOBItem.PutValue('BPL_FAMILLEMAT',        TOBEvent.GetString('FAMILLEMAT'));
  if TOBEvent.GetString('LIBFAMILLEMAT') = '' then
  begin
    If TOBEvent.GetString('ORIGINEITEM') <> 'PARCMAT' then
    begin
      Lib := RechDom('AFTTYPERESSOURCE',    TOBEvent.GetString('FAMILLEMAT'), True);
      TOBItem.PutValue('BPL_LIBFAMILLEMAT', Lib);
    end;
  end
  else
  begin
    If TOBEvent.GetString('ORIGINEITEM') = 'CHANTIER' then
      TOBItem.AddChampSupValeur('IDAFFECT', TOBEvent.GetString('LIBFAMILLEMAT'))
    else
      TOBItem.PutValue('BPL_LIBFAMILLEMAT', TOBEvent.GetString('LIBFAMILLEMAT'));
  end;

  TOBItem.PutValue('BPL_CODEETAT',          TOBEvent.GetString('CODEETAT'));
  TOBItem.PutValue('BPL_LIBEVENEMENT',      TOBEvent.GetString('LIBACTION'));
  TOBItem.PutValue('BPL_LIBEVENTREA',       TOBEvent.GetString('LIBREALISE'));

  TOBItem.putvalue('BPL_DUREE',             TOBEvent.GetValue('DUREE'));

  TOBItem.putvalue('BPL_FONCTION',          TOBEvent.GetString('CODEFONCTION'));
  TOBItem.putvalue('BPL_LIBFONCTION',       TOBEvent.GetString('LIBFONCTION'));
  //
  TOBItem.putvalue('BPL_NUMPHASE',          TOBEvent.GetInteger('NUMPHASE'));
  TOBItem.putvalue('BPL_LIBPHASE',          TOBEvent.GetString('LIBPHASE'));
  //
  TOBItem.putvalue('BPL_EMAIL',             TOBEvent.GetString('EMAIL'));
  TOBItem.putvalue('BPL_STANDCALEN',        TOBEvent.GetString('STANDCALEN'));
  TOBItem.putvalue('BPL_CALENSPECIF',       TOBEvent.GetString('CALENSPECIF'));
  TOBItem.putvalue('BPL_TYPERESSOURCE',     TOBEvent.GetString('TYPERESSOURCE'));
  TOBItem.putvalue('BPL_STYPERESSOURCE',    TOBEvent.GetString('STYPERESSOURCE'));
  TOBItem.putvalue('BPL_TYPEACTION',        TOBEvent.GetString('TYPEACTION'));

  TOBItem.putvalue('BPL_TYPEMVT',           TOBEvent.GetString('TYPEMVT'));
  TOBItem.putvalue('BPL_TYPECONGE',         TOBEvent.GetString('TYPECONGE'));
  TOBItem.putvalue('BPL_SALARIE',           TOBEvent.GetString('SALARIE'));
  //
  TOBItem.AddChampSupValeur('BPL_DESCRIPTIF',     TOBEvent.GetString('DESCRIPTIF'));
  TOBItem.AddChampSupValeur('BPL_NUMEROADRESSE',  TOBEvent.GetValue('NUMEROADRESSE'));
  TobItem.AddChampSupValeur('BPL_CONTROLCAL',     CtrlCal);
  //FV1 : 03/04/2018 - FS#3032 - ECHAFAUDAGE SERVICE : Ajouter date début et date fin du chantier
  DateD   := StrToDate(DateToStr(TOBEvent.getValue('DATEDEBAFF')));
  DateF   := StrToDate(DateToStr(TOBEvent.getValue('DATEFINAFF')));
  TOBItem.putValue('BPL_DATEDEBAFF',      DateD);
  TOBItem.putValue('BPL_DATEFINAFF',      DateF);
  //
  CompleteTobItem(TobItem)
  //
end;

Procedure CompleteTobItem(TobItem : Tob);
Var Icone    : Integer;
    QIcone   : TQuery;
    LibTypAct: String;
    TypeEnreg: String;
    BlocNote : Variant;
Begin

  CreationZoneTobItem(TobItem);

  if TobItem.GetValue('BPL_ORIGINEITEM') ='INTERV' then
    TypeEnreg := 'INT'
  else if TobItem.GetValue('BPL_ORIGINEITEM') ='ACT-GRC' then
    TypeEnreg := 'INT'
  else if TobItem.GetValue('BPL_ORIGINEITEM') ='ABSPAIE' then
    TypeEnreg := 'INT'
  else if TobItem.GetValue('BPL_ORIGINEITEM') ='PARCMAT' then
    TypeEnreg := 'PMA'
  else if TobItem.GetValue('BPL_ORIGINEITEM') ='CHANTIER' then
    TypeEnreg := 'PCA'
  else
    TypeEnreg := '';
  //
  ChargeIcone(TypeEnreg, TobItem);
  //
  //Vieille astuce pleine de puces pour palier au fait que d'un côté c'est INT et de l'autre PLA
  if TypeEnreg = 'INT' Then TypeEnreg := '';

  //On charge le contenu de l'item
  RechercheContenuItem(TypeEnreg, TobItem);
  TobItem.PutValue('LIBELLE', AnalyseChamp(TobItem, ContenuItem, FormatItem));

  //Oncharge le contenu de la bulle d'aide
  RechercheContenuHint(TypeEnreg, TobItem);
  TobItem.PutValue('HINT', AnalyseChamp(TobItem, ContenuHint, FormatHint));

  if TobItem.GetString('BPL_CODEETAT') = 'REA' then
    TobItem.AddChampSupValeur('MODIFIABLE','-')
  else
    TobItem.AddChampSupValeur('MODIFIABLE','X');

  // on decale la date de fin
  AffichageDateHeure(TobItem, '-');

end;

Procedure CreationZoneTobItem(TobTmp : Tob);
begin

  // Contrôle si l'ensemble des zones obligatoires sont présentes
  if not TobTmp.FieldExists('LIBELLE')          then 	TobTmp.AddChampSupValeur('LIBELLE', '');
  if not TobTmp.FieldExists('HINT')             then  TobTmp.AddChampSupValeur('HINT', '');
  if not Tobtmp.FieldExists('ICONETYPDOS')      then  TobTmp.AddChampSupValeur('ICONETYPDOS', '');
  
  if not TobTmp.FieldExists('BPL_DATEAFFDEB')   then 	TobTmp.AddChampSupValeur('BPL_DATEAFFDEB',  '');
  if not TobTmp.FieldExists('BPL_DATEAFFFIN')   then 	TobTmp.AddChampSupValeur('BPL_DATEAFFFIN',  '');
  if not TobTmp.FieldExists('BPL_HEUREAFFDEB')  then 	TobTmp.AddChampSupValeur('BPL_HEUREAFFDEB', '');
  if not TobTmp.FieldExists('BPL_HEUREAFFFIN')  then 	TobTmp.AddChampSupValeur('BPL_HEUREAFFFIN', '');

end;


Procedure AffichageDateHeure(TobItem : Tob; Sens : String);
Var DateDebE	: TDatetime;
  	HeureDeb	: TDateTime;
  	DateFinE	: TDatetime;
  	HeureFin  : TDatetime;
    HeureFinT	: TDateTime;
Begin

  DateDebE := TobItem.GetValue('BPL_DATEDEB');
  DateFinE := TobItem.GetValue('BPL_DATEFIN');
  HeureDeb := TobItem.GetValue('BPL_HEUREDEB');
  HeureFin := TobItem.GetValue('BPL_HEUREFIN');
  //
  //if heuredeb = heureFin then HeureFin := IncHour(HeureFin);
  //
  //On Fait ça afin d'avoir pour comme borne de début et de fin la zone
  //si 1 heure alors 10h00-10h00 = 10h00-11h00 à l'écran...
  //Pas logique mais c'est comme ça....

  if Cadencement = '001' then // gestion par 1/4 Heure
  begin
     if sens = '-' then
       HeureFinT := IncMinute (HeureFin,-15)
     else
       HeureFinT := IncMinute (HeureFin, 15);
  end else if Cadencement = '002' then // gestion par 1/2 Heure
  begin
     if sens = '-' then
       HeureFinT := IncMinute (HeureFin,-30)
     else
       HeureFinT := IncMinute (HeureFin, 30);
  end else if Cadencement = '003' then // gestion par heure
  begin
     if sens = '-' then
       HeureFinT := IncHour (HeureFin,  -1)
     else
       HeureFinT := IncHour (HeureFin,   1);
  end else if Cadencement = '004' then // gestion par 1/2 Journée
  Begin
     if sens = '-' then
       HeureFinT := incminute (HeureFin,-72)
     Else
       HeureFinT := IncHour (HeureFin,   72);
  end else if Cadencement = '005' then // gestion par Journée
  begin
  	// pas de decalage dans ce cas la
  end else if Cadencement = '008' then // gestion par Période
  begin
     {TobTmp.PutValue('BEP_DATEFIN', TobTmp.GetValue('BEP_DATEFIN'));}
  end;

  if HeureFinT <= heureDeb then HeureFinT := HeureDeb;

  //TobItem.PutValue('BPL_DATEDEB', Trunc(StrToDate(DateDebe)) + HeureDeb);
  //TobItem.PutValue('BPL_DATEFIN', Trunc(StrToDate(DateFine)) + HeureFin);
  //TobItem.PutValue('BPL_HEUREDEB', HeureDeb);
  //TobItem.PutValue('BPL_HEUREFIN', HeureFin);
  //
  TobItem.PutValue('BPL_DATEAFFDEB',  DateTimeToStr(Trunc(DateDebe) + HeureDeb));
  TobItem.PutValue('BPL_DATEAFFFIN',  DateTimeToStr(Trunc(DateFine) + HeureFinT));
  TobItem.PutValue('BPL_HEUREAFFDEB', HeureDeb);
  TobItem.PutValue('BPL_HEUREAFFFIN', HeureFinT);
    
end;

{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 17/02/2016
Modifié le ... :   /  /    
Description .. : Chargement des évènements, des informations de cadencement
Mots clefs ... : 
*****************************************************************}
procedure ChargeEvent(var R: RecordPlanning; var P: THPlanningBTP);
Var DateDebE	: string;
  	HeureDeb	: TDateTime;
  	DateFinE	: String;
  	HeureFin,HeureFinT	: TDateTime;
  	Qitem: TQuery;
  	MaRequete: string;
  	Indice : integer;
begin
  if R.TobEvents = nil then
  begin
    TobFree(R.TobEvents);
  end else
  begin
    R.TobEvents := TOB.Create('la tob', nil, -1);
  end;

  P.TobEvenements := nil;

  MaRequete := 'SELECT * FROM HREVENEMENT WHERE ((HEV_DATEDEBUT<="' + UsDateTime(P.IntervalFin)
    + '" AND HEV_DATEDEBUT>="' + UsDateTime(P.IntervalDebut) + '") OR '
    + '(HEV_DATEFIN<="' + UsDateTime(P.IntervalFin) + '" AND HEV_DATEFIN>="' + UsDateTime(P.IntervalDebut) + '") OR '
    + '(HEV_DATEFIN>="' + UsDateTime(P.IntervalFin) + '" AND HEV_DATEDEBUT<="' + UsDateTime(P.IntervalDebut) + '"))'
    + ' AND (HEV_EVENEMENT NOT LIKE "!%")';

  // Les items totaux
  QItem := OpenSQL(MaRequete, True,-1,'',true);

  R.TobEvents.LoadDetailDB('HREVENEMENT', '', '', Qitem, False, True);
  for Indice := 0 to R.TobEvents.detail.count -1 do
  begin
    DateDebE := DateToStr(R.TobEvents.detail[Indice].GetValue('HEV_DATEDEBUT'));
    DateFinE := DateToStr(R.TobEvents.detail[Indice].GetValue('HEV_DATEFIN'));
    HeureDeb := StrToTime('00:00:00');
    HeureFin := StrToTime('23:59:59');
    R.TobEvents.detail[Indice].PutValue('HEV_DATEDEBUT', Trunc(StrToDate(DateDebe)) + HeureDeb);
    R.TobEvents.detail[Indice].PutValue('HEV_DATEFIN',   Trunc(StrToDate(DateFine)) + HeureFin);
  end;

  Ferme(QItem);

  P.GestionEvenements := False;

  if (R.TobEvents.detail.count > 0) then
  begin
    P.GestionEvenements := True;
    P.TobEvenements := R.TobEvents;
  end;

end;

Procedure ChargeCadencement(Var R: RecordPlanning; P: THPlanningBTP; T: TOB);
Var   FormatDate  : string;
begin

  cadencement             := T.Getvalue('HPP_CADENCEMENT');

  // MAJ regroupement de date
  P.ActiveLigneGroupeDate := T.Getvalue('HPP_AFFDATEGROUP') = 'X';

  P.interval              := piNone;

  ControleCadencement(Cadencement, R, P, TobParametres);

end;

Procedure ControleCadencement(Cadencement : String; R: RecordPlanning; P: THPlanningBTP; T : TOB);
Var StSQL         : String;
    PeriodeDebut  : string;
    PeriodeFin    : string;
    FormatDate    : string;
begin

  PeriodeDebut  := '000';
  PeriodeFin    := '999';

    //Mise à jour du format de la date
  if Cadencement = '001' then
  begin
    P.Interval    := piQuartHeure;
    FormatDate    := T.Getvalue('HPP_FORMATDATECOL0');
    P.DateFormat  := 'hh:mm';
  end
  else if Cadencement = '002' then
  begin
    P.Interval    := piDemiHeure;
    FormatDate    := T.Getvalue('HPP_FORMATDATECOL0');
    P.DateFormat  := 'hh:mm';
  end
  else if (Cadencement = '003') then
  begin
    P.Interval    := piHeure;
    FormatDate    := T.Getvalue('HPP_FORMATDATECOL0');
    P.dateFormat  := 'hh:mm';
  end
  else if (Cadencement = '004') then
  Begin
    P.ActiveLigneGroupeDate := False;
    P.interval := piDemiJour;
  end
  else if (Cadencement = '005') then
  begin
    P.Interval    := piJour;
    FormatDate    := T.Getvalue('HPP_FORMATDATECOL0');
    P.DateFormat  := Rechdom('HRFORMATDATE', FormatDate, True);
  end
  else if (Cadencement = '006') then
  begin
    P.Interval    := piSemaine;
    P.DateFormat  := 'dd/mm/yyyy';
  end
  else if (Cadencement = '007') then
  begin
    P.Interval    := pimois;
    P.DateFormat  := 'mmmmm';
  end
  else if cadencement = '008' then  //Lecture de la tablette Service de CHR (???)
  begin
    P.ActiveLigneGroupeDate := False;
    P.interval              := piDivers;
    if (not assigned(R.TobPeriodeDivers)) then
    begin
      R.TobPeriodeDivers    := TOB.Create('la tob', nil, -1);
      StSQL := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="HSE" ';
      StSQL := StSQL + '  AND CC_CODE>="' + PeriodeDebut + '"';
      StSQL := StSQL + '  AND CC_CODE<="' + PeriodeFin + '"';
      R.TobPeriodeDivers.LoadDetailFromSQL(StSql);
    end else
    begin
      R.TobPeriodeDivers.ClearDetail;
      R.TobPeriodeDivers.InitValeurs(false); 
    end;
    P.TobPeriodeDivers      := R.TobPeriodeDivers;
    P.ChampLibelleDivers    := 'CC_LIBELLE';
    P.ChampNumPeriodeDebut  := 'NUMEROHRPERIODEDEBUT';
    P.ChampNumPeriodeFin    := 'NUMEROHRPERIODEFIN';
  end
  else if cadencement = '009' then
  begin
    P.ActiveLigneGroupeDate := False;
    P.interval              := piNone;
    P.DateFormat            := 'ddd dd/mm' ;
  end;

  if (P.ActiveLigneGroupeDate) then
  begin
    if pos(cadencement,'001;002;003;004') > 0 then
      P.CumulInterval := PciJour
    else if cadencement = '005' then
      P.CumulInterval := pciSemaine
    else if cadencement = '006' then
      P.CumulInterval := pciMois
    else if cadencement = '007' then
      P.CumulInterval := pciTrimestre
    else
      P.CumulInterval := pciNone;
  end;

  if pos(cadencement,'001;002;003;004;005') > 0 then
  begin
    if T.Getvalue('HPP_HEUREDEBUT')=0 then
       P.JourneeDebut := GetParamSoc('SO_HEUREDEB')
    else
       P.JourneeDebut := T.Getvalue('HPP_HEUREDEBUT');
    if T.Getvalue('HPP_HEUREFIN')=0 then
       P.JourneeFin := GetParamSoc('SO_HEUREFIN')
    else
       P.JourneeFin := T.Getvalue('HPP_HEUREFIN');
  end;

end;

Procedure ChargementEtat(var R: RecordPlanning; T: TOB; DateEnCours: TDateTime; ModeGestionPlanning: Boolean);
Var StSQL  :  string;
    QEtats : TQuery;
begin
  if R.TobEtats = nil then
  begin
    R.TobEtats := TOB.Create ('Les etats', Nil, -1);
  end else
  begin
    R.TobEtats.ClearDetail;
    R.TobEtats.InitValeurs(false); 
  end;

  //if P.TobEtats <> nil then P.TobEtats.ClearDetail;


  StSQL := 'SELECT * FROM BTETAT WHERE BTA_ASSOSRES="X" ORDER BY BTA_BTETAT';

  QEtats   := OpenSQL(StSQL, True,-1,'',true);
  R.TobEtats.LoadDetailDB('BTETAT', '', '', QEtats, False);

  //P.TobEtats := R.TobEtats;

  Ferme(QEtats);

end;

Procedure ChargeIcone(TypeAction : String; TobItem : TOB);
Var StSQl     : string;
    QIcone    : TQuery;
    Icone     : Integer;
    LibTypAct : String;
    TOBI : TOB;
begin
  TOBI := TOBIcone.findFirst(['BTA_BTETAT','BTA_TYPEACTION'],[TobItem.getValue('BPL_BTETAT'),TypeAction],true);
  if TOBI <> nil then
  begin
    Icone      := TOBIcone.GetInteger('BTA_NUMEROICONE');
    LibTypAct  := TOBIcone.GetString('BTA_LIBELLE');
  end else
  begin
    Icone     := -1;
    LibTypAct := '';

    StSQL := 'SELECT BTA_NUMEROICONE, BTA_LIBELLE FROM BTETAT ';
    StSQl := StSQL + ' WHERE BTA_BTETAT     = "'  + TobItem.getValue('BPL_BTETAT');
    StSQl := StSQL + '"  AND BTA_TYPEACTION = "'  + TypeAction;
    StSQl := StSQl + '"  AND BTA_ASSOSRES   = "X" ORDER BY BTA_BTETAT';

    // Gestion des icones selon type d'évènement
    QIcone   := OpenSQL(StSQL, True,-1,'',true);

    if not QIcone.Eof then
    begin
      Icone      := QIcone.FindField('BTA_NUMEROICONE').AsInteger;
      LibTypAct  := QIcone.findfield('BTA_LIBELLE').AsString;
    end;
  end;
  TobItem.PutValue('ICONETYPDOS', Icone);

  Ferme(QIcone);

end;
{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 10/02/2016
Modifié le ... :   /  /
Description .. : Gestion des affichage des Items et des bulles info
Mots clefs ... :
*****************************************************************}
Procedure RechercheContenuItem(TypeEnreg : String; TobItem : TOb);
Var STSQL : String;
    QQ    : TQuery;
Begin

  //Lecture de la TABLETTE BTCONTENUITEM
  StSQL := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE = "BCI" AND CC_LIBRE="' + TypeEnreg + '"';
  QQ    := OpenSQL(StSQL,False,1,'',True);
  if not QQ.eof then
  begin
    GestionAffichageZoneItem(QQ.FindField('CC_CODE').AsString);
    TobItem.AddChampSupValeur('CONTENUITEM', ContenuItem);
    TobItem.AddChampSupValeur('FORMATITEM',  FormatItem);
    if pos('#13#10', ContenuItem) <> 0 then TobItem.AddChampSupValeur('MULTILINE', 'X');
  end;

  Ferme(QQ);

end;

Procedure RechercheContenuHint(TypeEnreg : String; TobItem : TOb);
Var STSQL : String;
    QQ    : TQuery;
Begin

  //Lecture de la TABLETTE BTCONTENUHINT
  StSQL := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE = "BCH" AND CC_LIBRE="' + TypeEnreg + '"';
  QQ    := OpenSQL(StSQL,False,1,'',True);
  if not QQ.eof then
  begin
    GestionAffichageZoneHint(QQ.FindField('CC_CODE').AsString);
    TobItem.AddChampSupValeur('CONTENUHINT', ContenuHint);
    TobItem.AddChampSupValeur('FORMATHINT',  FormatHint);
  end;

  Ferme(QQ);

end;


Procedure GestionAffichageZoneItem(Numliste : String);
begin

  //gestions des zones à afficher dans l'items
  GestAffiche.FormatGapp := '';

  ContenuItem := MajItem(RechDom('BTCONTENUITEM',Numliste, true), GestAffiche,'C');

  FormatItem := GestAffiche.FormatGapp;

  if ContenuItem = '' then
  begin
    ContenuItem := 'BPL_BTETAT;("#13#10");BPL_DUREE';
    FormatItem := '$;';
  end;

end;

Procedure GestionAffichageZoneHint(Numliste : String);
Begin

  //gestions des zones à afficher dans la bulle d'aide (hint)
  FormatHint  := '';

  ContenuHint := MajItem(RechDom('BTCONTENUHINT', Numliste, true), GestAffiche, 'C');

  FormatHint  := GestAffiche.FormatGapp;

  If Contenuhint = '' then
  begin
    ContenuHint := 'BPL_BTETAT;BPL_LIBACTION';
    FormatHint  := '$;;$;;';
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 17/02/2016
Modifié le ... :   /  /
Description .. : Requetes de chargement des colonnes fixes du planning
                 Procedure de chargement de la commande SQL
                 Cette fonction permet également de déterminé les élèments de tri.
Mots clefs ... :
*****************************************************************}
Function ChargeColonneFixe(TypRes : String; var P: THPlanningBTP) : String;
Var ZoneTri       : String;
    ZoneRes       : String;
    SQLOrder      : String;
    I             : integer;
    NbCT          : integer;
    Ind           : integer;
    Regroupement  : Boolean;
    Condition     : String;
begin

  NbCT := TobParametres.getinteger('HPP_NBLIGDIVERS') + 1;

  if TypRes = 'X' then NbCT := NbCT+1;

  Result := '';
  
  I := 1;

  //Tri en fonction des types de ressource associés à la famille de ressource
  if TypRes = 'X' then
  Begin
    P.SetLinkColFixed(0, True); //---> voir nouvel AGL
    ZoneTri := 'ARS_CHAINEORDO';
    I := 2
  end;

  //On gère ici l'affichage des colonnes fixes (à gauche) en fonction du type de planning
  //Une requête par type de planning....
  if P.TypePlanning = 'PMA' then
    Result := ChargeColonnePMA(TypRes, P)
  else if Pos(P.TypePlanning, 'PLA;PTA;PRA') > 0 then
    Result := ChargeColonnePLA(TypRes, P)
  else if P.TypePlanning = 'PFM' then
    Result := ChargeColonnePFM(TypRes, P)
  else if P.TypePlanning = 'PPA' then
    Result := ChargeColonnePPA(TypRes, P)
  else if P.TypePlanning = 'PSF' then
    Result := ChargeColonnePSF(TypRes, P)
  else if P.TypePlanning = 'PTR' then
    Result := ChargeColonnePTR(TypRes, P)
  else if P.TypePlanning = 'PFO' then
    Result := ChargeColonnePFO(TypRes, P)
  else if P.TypePlanning = 'PAR' then
    Result := ChargeColonnePAR(P)
  else
    Result := ChargeColonnePLA(TypRes, P);

  //Voir comment on pourrait gérer le fait de n'avoir que les fonctions en fonction peut-être du modèle de planning (???)
  //if pos('AFO_LIBELLE',Result) > 0 then Result := Result + ' AND AFO_GEREPLANNING ="X" ';

  //gestion du tri en fonction des paramètres récupérés au dessus
  //Tri en fonction des paramétrage du client
  for ind := I to NbCT do
  begin
    i := ind;
    //
    if TypRes = 'X' then I := I - 1;
    //
    if TobParametres.GetValue('HPP_REGCOL' + IntToStr(I))= 'X' then
       Regroupement := true
    else
       Regroupement := False;
    //
    if TypRes = 'X' then
       P.SetLinkColFixed(I, Regroupement)
    else
       P.SetLinkColFixed(I-1, Regroupement);
    //
    if pos(P.TypePlanning, 'PMA;PFM;PPA') > 0 then
      ZoneTri := LectLibCol('CC', 'BMA', TobParametres.getvalue('HPP_LIBCOL' + IntToStr(I)), 'CC_LIBRE')
    else if P.TypePlanning = 'PAR' then
      ZoneTri := LectLibCol('CC', 'BAR', TobParametres.getvalue('HPP_LIBCOL' + IntToStr(I)), 'CC_LIBRE')
    else
      zoneTri := LectLibCol('CC', 'BLI', TobParametres.getvalue('HPP_LIBCOL' + IntToStr(I)), 'CC_LIBRE');
    //
    if (ZoneTri <> '') then
    begin
      if (SQLOrder <> '') then
        SQLOrder := SQLOrder + ',' + ZoneTri
      else
        SQLOrder := ZoneTri;
    end;
  end;

  if (SQLOrder = '') then
  begin
    if      (P.TypePlanning = 'PMA') then SQLOrder := 'BMA_CODEMATERIEL'
    else if (P.TypePlanning = 'PFM') then SQLOrder := 'BMA_CODEFAMILLE'
    else if (P.TypePlanning = 'PPA') then SQLOrder := 'BTA_BTETAT'
    else if (P.TypePlanning = 'PAR') then SQLOrder := 'AFF_AFFAIRE'
    else if (P.TypePlanning = 'PTR') then SQLOrder := 'ARS_TYPERESSOURCE'
    else if (P.TypePlanning = 'PFO') then SQLOrder := 'ARS_FONCTION1'
    else                                  SQLOrder := 'ARS_RESSOURCE';
  end;

  if TobParametres.GetBoolean('HPP_PLANNINGTYPETD') then
  begin
    if Pos(SQLOrder,'CHAINEORDO') < 0 then SQLOrder := 'ARS_CHAINEORDO,' + SQLOrder;
  end;

  Result := Result + ' ORDER BY ' + SQLOrder;

end;

Function ChargeColonnePMA(TypRes : string; var P: THPlanningBTP) : String;
Var ZoneRes : String;
begin

  ZoneRes := ChargeStrByTablette('BMA', 'CC_LIBRE');

  Result  := 'SELECT "PARCMAT" AS ORIGINEITEM, ';
  Result  := Result + ZoneRes + ' FROM BTMATERIEL ';
  Result  := Result + 'LEFT JOIN BTFAMILLEMAT ON BMA_CODEFAMILLE=BFM_CODEFAMILLE ';

  //Gestion des conditions de selection des enregistrement
  Result := Result + ' WHERE BMA_FERME="-" ';

  if P.Famres <> 'GEN' then Result := Result + ' AND BMA_CODEFAMILLE IN (' + p.FamRes + ')';

  Result := Result + ' AND BFM_NONGEREPLANNING="X"';

end;

Function ChargeColonnePLA(TypRes : String; var P: THPlanningBTP) : String;
Var ZoneREs : String;
begin

  ZoneRes := ChargeStrByTablette('BLI', 'CC_LIBRE');
  ZoneRes := ZoneRes + ', ARS_STANDCALEN, ARS_CALENSPECIF';

  //chargement de l'état de la ressource (Type d'action)
  if Pos(P.TypePlanning, 'PTA;PRA') > 0 then
    Result := 'SELECT "CHANTIER" AS ORIGINEITEM, '
  else
    Result := 'SELECT "INTERV" AS ORIGINEITEM, ';

  //gestion de la jointure si planning sur sous famille ressource
  if TypRes = 'X' then
  Begin
    Result := Result + ZoneRes + ',BTR_LIBELLE FROM RESSOURCE ';
    Result := Result + ' LEFT JOIN BTTYPERES ON ARS_CHAINEORDO=BTR_TYPRES';

    //Gestion du libellé de la ressource
    if pos('AFO_LIBELLE',ZoneRes) > 0 then
    Begin
      Result := Result + ' LEFT JOIN FONCTION  ON ARS_FONCTION1=AFO_FONCTION ';
      Result := Result + ' LEFT JOIN BFONCTION ON AFO_FONCTION =BFO_FONCTION ';
    end;

    //Gestion Affichage du sous-type ressource au Planning
    Result := Result + ' WHERE ARS_RESSOURCE in ';
    Result := Result + '(SELECT ARS_RESSOURCE FROM RESSOURCE LEFT JOIN BTTYPERES ON ARS_CHAINEORDO=BTR_TYPRES ';
    Result := Result + 'WHERE BTR_GEREPLANNING="X" AND  ARS_FERME="-") OR ARS_CHAINEORDO="" AND ';
  end
  else
  begin
    Result := Result + ZoneRes + ' FROM RESSOURCE ';
    //Gestion du libellé du Sous-type Ressource
    if pos('BTR_LIBELLE',ZoneRes) > 0 Then Result := Result + ' LEFT JOIN BTTYPERES ON ARS_CHAINEORDO=BTR_TYPRES ';
    //Gestion du libellé de la Fonction Ressource
    if pos('AFO_LIBELLE',ZoneRes) > 0 then
    Begin
      Result := Result + ' LEFT JOIN FONCTION  ON ARS_FONCTION1=AFO_FONCTION ';
      Result := Result + ' LEFT JOIN BFONCTION ON AFO_FONCTION =BFO_FONCTION ';
    end;
    Result := Result + ' WHERE ';
  end;

  //Gestion des conditions de selection des enregistrement
  Result := Result + 'ARS_FERME="-" ';
  //FV1 - 03/04/2018 : FS#3033 - ECHAFAUDAGE SERVICE : Ne pas afficher tous les salariés dans le planning
  if pos('BTR_LIBELLE',ZoneRes) > 0 then Result := Result + ' AND BTR_GEREPLANNING="X" ';
  if pos('AFO_LIBELLE',ZoneRes) > 0 then Result := Result + ' AND BFO_GEREPLANNING="X" ';
  //
  if P.Famres <> 'GEN' then
    Result := Result + ' AND ARS_TYPERESSOURCE = "' + p.FamRes + '" '
  Else
    Result := Result + ' AND ARS_TYPERESSOURCE in ("INT", "SAL", "LOC", "ST", "MAT")';


end;

Function ChargeColonnePFM(TypRes : String; var P: THPlanningBTP) : String;
Var ZoneRes : String;
    Condition : string;
begin

  ZoneRes := ChargeStrByTablette('BMA', 'CC_LIBRE');
  ZoneRes := ZoneRes + ', ARS_STANDCALEN, ARS_CALENSPECIF';
  //
  Result  := 'SELECT "PARCMAT" AS ORIGINEITEM, ';
  Result  := Result + ZoneRes + ' FROM BTMATERIEL ';
  Result  := Result + ' LEFT JOIN BTFAMILLEMAT ON (BMA_CODEFAMILLE=BFM_CODEFAMILLE) ';
  Result  := Result + ' LEFT JOIN RESSOURCE    ON (BMA_RESSOURCE=ARS_RESSOURCE) ';
  Result  := Result + ' WHERE BMA_FERME="-" ';

  //Gestion des conditions de selection des enregistrement
  Condition := LectLibCol('CO', 'BMP', p.TypePlanning,'CO_LIBRE') + '="' + P.CodeOnglet + '"';
  if Condition <> '' then Result := Result + ' AND BMA_CODEFAMILLE' + Condition;

end;

Function ChargeColonnePPA(TypRes : String; var P: THPlanningBTP) : String;
Var ZoneRes   : String;
    Condition : string;
begin

  ZoneRes := ChargeStrByTablette('BMA', 'CC_LIBRE');
  ZoneRes := ZoneRes + ', ARS_STANDCALEN, ARS_CALENSPECIF';
  //
  Result := 'SELECT DISTINCT "PARCMAT" AS ORIGINEITEM, ';
  Result := Result + Zoneres + ' FROM BTMATERIEL ';
  Result := Result + ' LEFT JOIN BTFAMILLEMAT ON (BMA_CODEFAMILLE=BFM_CODEFAMILLE) ';
  Result := Result + ' LEFT JOIN RESSOURCE    ON (BMA_RESSOURCE=ARS_RESSOURCE) ';
  Result := Result + ' LEFT JOIN BTEVENTMAT   ON (BEM_CODEMATERIEL=BMA_CODEMATERIEL) ';
  Result := Result + ' LEFT JOIN BTETAT       ON (BEM_BTETAT=BTA_BTETAT) ';
  Result := Result + 'WHERE BMA_FERME="-" ';
  Result := Result + '  AND BFM_NONGEREPLANNING="X"';

  ///gestion de la famille renseignée au niveau du paramétrage Planning
  if P.Famres <> 'GEN' then Result := Result + ' AND BMA_CODEFAMILLE In (' + P.FamRes + ')';

  //Gestion des conditions de selection des enregistrement
  Condition := LectLibCol('CO', 'BMP', p.TypePlanning,'CO_LIBRE') + '="' + P.CodeOnglet + '"';
  if Condition <> '' then Result := Result + ' AND BTA_BTETAT' + Condition;

end;

Function ChargeColonnePSF(TypRes : String; var P: THPlanningBTP) : String;
Var ZoneRes   : String;
    Condition : string;
begin

  ZoneRes := ChargeStrByTablette('BLI', 'CC_LIBRE');
  ZoneRes := ZoneRes + ', ARS_STANDCALEN, ARS_CALENSPECIF';

  if TypRes = 'X' then
  Begin
    Result := 'SELECT "INTERV" AS ORIGINEITEM, ';
    Result := Result + ZoneRes + ',BTR_LIBELLE ';
    Result := Result + '  FROM RESSOURCE LEFT JOIN BTTYPERES ON ARS_CHAINEORDO=BTR_TYPRES';
    //Gestion du libellé de la ressource
    if pos('AFO_LIBELLE',ZoneRes) > 0 then
    Begin
      Result := Result + ' LEFT JOIN FONCTION ON ARS_FONCTION1=AFO_FONCTION ';
      Result := Result + ' LEFT JOIN BFONCTION ON BFO_FONCTION=AFO_FONCTION ';
    end;
    Result := Result + ' WHERE ARS_RESSOURCE in '; //Gestion Affichage du sous-type ressource au Planning
    Result := Result + '(SELECT ARS_RESSOURCE FROM RESSOURCE LEFT JOIN BTTYPERES ON ARS_CHAINEORDO=BTR_TYPRES ';
    Result := Result + 'WHERE BTR_GEREPLANNING="X" AND  ARS_FERME="-") AND ';
  end
  else
  begin
    Result := 'SELECT "INTERV" AS ORIGINEITEM, ';
    Result := Result + ZoneRes + ' FROM RESSOURCE ';
    //Gestion du Libellé de la fonction de la ressource
    if pos('AFO_LIBELLE',ZoneRes) > 0 then
    Begin
      Result := Result + ' LEFT JOIN FONCTION ON ARS_FONCTION1=AFO_FONCTION ';
      Result := Result + ' LEFT JOIN BFONCTION ON BFO_FONCTION=AFO_FONCTION ';
    end;
    Result := Result + ' WHERE ';
  end;

  //Gestion des conditions de selection des enregistrement
  Result := Result + 'ARS_FERME="-" ';
  if pos('AFO_LIBELLE',ZoneRes) > 0 then   Result := Result + ' AND BFO_GEREPLANNING="X" ';

  if P.Famres <> 'GEN' then Result := Result + ' AND ARS_TYPERESSOURCE = "' + P.FamRes + '" ';

  Condition := LectLibCol('CO', 'BMP', p.TypePlanning,'CO_LIBRE') + '="' + P.CodeOnglet + '" ';
  Result := Result + ' AND ' + Condition;

end;

Function ChargeColonnePTR(TypRes : String; var P: THPlanningBTP) : String;
Var ZoneRes   : String;
    Condition : string;
    SsTypeRessource : string;
begin

  SsTypeRessource := TobParametres.GetValue('HPP_PLANNINGTYPETD');

  ZoneRes := ChargeStrByTablette('BLI', 'CC_LIBRE');
  ZoneRes := ZoneRes + ', ARS_STANDCALEN, ARS_CALENSPECIF';

  Result := 'SELECT "INTERV" AS ORIGINEITEM, ';
  Result := Result + ZoneRes + ',CC_LIBELLE ';

  if SsTypeRessource='X' then Result := Result + ', BTR_LIBELLE';

  Result := Result + '  FROM RESSOURCE ';
  Result := Result + '  LEFT JOIN CHOIXCOD  ON ARS_TYPERESSOURCE=CC_CODE';

  if pos('AFO_LIBELLE',ZoneRes) > 0 then
  Begin
    Result := Result + ' LEFT JOIN FONCTION  ON ARS_FONCTION1=AFO_FONCTION ';
    Result := Result + ' LEFT JOIN BFONCTION ON BFO_FONCTION=AFO_FONCTION ';
  end;

  if SsTypeRessource='X' then Result := Result + '  LEFT JOIN BTTYPERES ON ARS_CHAINEORDO=BTR_TYPRES';

  Result := Result + ' WHERE ';

  //Gestion des conditions de selection des enregistrement
  Result := Result + 'ARS_FERME="-" ';

  if pos('AFO_LIBELLE',ZoneRes) > 0 then Result := Result + 'AND BFO_GEREPLANNING="X" ';
  if SsTypeRessource='X'            then Result := Result + ' AND BTR_GEREPLANNING="X" ';

  Condition := LectLibCol('CO', 'BMP', p.TypePlanning,'CO_LIBRE') + '="' + P.CodeOnglet + '"';
  Result := Result + ' AND ' + Condition;

end;

Function ChargeColonnePFO(TypRes : String; var P: THPlanningBTP) : String;
Var ZoneRes   : String;
    Condition : string;
    SsTypeRessource : string;
begin

  //SsTypeRessource := TobParametres.GetValue('HPP_PLANNINGTYPETD');

  ZoneRes := ChargeStrByTablette('BLI', 'CC_LIBRE');
  ZoneRes := ZoneRes + ', ARS_STANDCALEN, ARS_CALENSPECIF';

  Result := 'SELECT "CHANTIER" AS ORIGINEITEM, ' + ZoneRes;

  //if SsTypeRessource='X' then Result := Result + ', BTR_LIBELLE';

  Result := Result + '  FROM FONCTION ';
  Result := Result + '  LEFT JOIN BFONCTION ON AFO_FONCTION=BFO_FONCTION  ';
  Result := Result + '  LEFT JOIN RESSOURCE ON ARS_FONCTION1=AFO_FONCTION';

  //if SsTypeRessource='X' then
  Result := Result + '  LEFT JOIN BTTYPERES ON ARS_CHAINEORDO=BTR_TYPRES';

  Result := Result + ' WHERE ';

  //Gestion des conditions de selection des enregistrement
  Result := Result + 'ARS_FERME="-" ';

  //FV1 : 03/04/2018 - FS#3033 - ECHAFAUDAGE SERVICE : Ne pas afficher tous les salariés dans le planning
  Result := Result + 'AND BFO_GEREPLANNING="X" ';
  //

  ///gestion de la famille renseignée au niveau du paramétrage Planning
  if P.Famres <> 'GEN' then Result := Result + ' AND BFO_FONCTION IN (' + P.FamRes + ') ';

  Condition := LectLibCol('CO', 'BMP', p.TypePlanning,'CO_LIBRE') + '="' + P.CodeOnglet + '" ';
  Result := Result + ' AND ' + Condition;

end;     

Function ChargeColonnePAR(var P: THPlanningBTP) : String;
Var ZoneRes   : String;
begin

  ZoneRes := ChargeStrByTablette('BAR', 'CC_LIBRE');

  Result := 'SELECT ' + ZoneRes + '  FROM AFFAIRE ';

  //Gestion des conditions de selection des enregistrement
  Result := Result + 'WHERE AFF_ETATAFFAIRE="ACP" AND AFF_AFFAIRE0="A"';

end;

{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 17/02/2016
Modifié le ... :   /  /
Description .. : Gestion des différentes requêtes utilisées pour charger les
Suite ........ : items dans le planning
Mots clefs ... :
*****************************************************************}
Function ChargeRequeteIntervention(P : THPlanningBTP; DateDebut, DateFin : TdateTime) : String;
begin

  Result := 'SELECT "INTERV"            AS ORIGINEITEM, ';
  Result := Result + 'BEP_CODEEVENT     AS IDEVENT, ';
  Result := Result + 'BEP_AFFAIRE       AS CODEAFFAIRE, ';
  Result := Result + 'AFF_LIBELLE       AS LIBAFFAIRE, ';
  Result := Result + 'BEP_TIERS         AS CODETIERS, ';
  Result := Result + 'T_AUXILIAIRE      AS AUXILIAIRE, ';
  Result := Result + 'T_LIBELLE         AS LIBTIERS, ';
  Result := Result + 'ADR_JURIDIQUE     AS NOMJURIDIQUE, ';
  Result := Result + 'ADR_LIBELLE       AS NOM1, ';
  Result := Result + 'ADR_LIBELLE2      AS NOM2, ';
  Result := Result + 'ADR_ADRESSE1      AS ADR1, ';
  Result := Result + 'ADR_ADRESSE2      AS ADR2, ';
  Result := Result + 'ADR_ADRESSE3      AS ADR3, ';
  Result := Result + 'ADR_CODEPOSTAL    AS CP,  ';
  Result := Result + 'ADR_VILLE         AS VILLE, ';
  Result := Result + 'BEP_DATEDEB       AS DATEDEB, ';
  Result := Result + 'BEP_DATEFIN       AS DATEFIN, ';
  Result := Result + 'BEP_HEUREDEB      AS HEUREDEB, ';
  Result := Result + 'BEP_HEUREFIN      AS HEUREFIN, ';
  Result := Result + 'BEP_DUREE         AS DUREE, ';
  Result := Result + 'BEP_BTETAT        AS ACTION, ';
  Result := Result + 'BTA_LIBELLE       AS LIBACTION, ';
  Result := Result + 'BTA_ASSOSRES      AS ASSOSRES, ';
  Result := Result + 'BEP_RESSOURCE     AS CODERESSOURCE, ';
  Result := Result + 'ARS_LIBELLE       AS LIBRESSOURCE, ';
  Result := Result + 'ARS_EQUIPERESS    AS EQUIPE, ';
  Result := Result + 'BEP_RESSOURCE     AS MATERIEL, ';
  Result := Result + 'ARS_LIBELLE       AS LIBMATERIEL, ';
  Result := Result + 'ARS_TYPERESSOURCE AS FAMILLEMAT, ';
  Result := Result + '""                AS LIBFAMILLEMAT, ';
  Result := Result + 'AFF_ETATAFFAIRE   AS CODEETAT, ';
  Result := Result + '""                AS LIBEVENEMENT, ';
  Result := Result + '""                AS LIBREALISE, ';
  Result := Result + 'BEP_BLOCNOTE      AS BLOCNOTE, ';
  Result := Result + 'ARS_FONCTION1     AS CODEFONCTION, ';
  Result := Result + 'AFO_LIBELLE       AS LIBFONCTION, ';
  Result := Result + '"0"               AS NUMPHASE, ';
  Result := Result + '""                AS LIBPHASE, ';
  Result := Result + 'ARS_EMAIL         AS EMAIL, ';
  Result := Result + 'ARS_STANDCALEN    AS STANDCALEN, ';
  Result := Result + 'ARS_CALENSPECIF   AS CALENSPECIF, ';
  Result := Result + 'ARS_TYPERESSOURCE AS TYPERESSOURCE, ';
  Result := Result + 'ARS_CHAINEORDO    AS STYPERESSOURCE, ';
  Result := Result + 'BTR_LIBELLE       AS LIBSTYPERESSOURCE, ';
  Result := Result + '""                AS TYPEACTION, ';
 	Result := Result + '""                AS TYPEMVT, ';
  Result := Result + '""                AS TYPECONGE, ';
  Result := Result + 'ARS_SALARIE       AS SALARIE, ';
  Result := Result + 'AFF_DESCRIPTIF    AS DESCRIPTIF, ';
  Result := Result + 'BEP_NUMEROADRESSE AS NUMEROADRESSE, ';
  //FV1 - 03/04/2018 : FS#3032 - ECHAFAUDAGE SERVICE : Ajouter date début et date fin du chantier
  Result := Result + 'AFF_DATEDEBUT     AS DATEDEBAFF, ';
  Result := Result + 'AFF_DATEFIN       AS DATEFINAFF ';
  Result := Result + 'FROM BTEVENPLAN ';
  Result := Result + 'LEFT JOIN BTETAT    ON (BTA_BTETAT=BEP_BTETAT) ';
  Result := Result + 'LEFT JOIN AFFAIRE   ON (AFF_AFFAIRE=BEP_AFFAIRE) ';
  Result := Result + 'LEFT JOIN TIERS     ON (T_TIERS=BEP_TIERS) ';
  Result := Result + 'LEFT JOIN ADRESSES  ON (ADR_REFCODE=AFF_AFFAIRE) ';
  Result := Result + 'LEFT JOIN RESSOURCE ON (ARS_RESSOURCE=BEP_RESSOURCE) ';
  Result := Result + 'LEFT JOIN FONCTION  ON (AFO_FONCTION=ARS_FONCTION1) ';
  Result := Result + 'LEFT JOIN BFONCTION ON (BFO_FONCTION=AFO_FONCTION) ';

  Result := Result + 'LEFT JOIN BTTYPERES ON (BTR_TYPRES=ARS_CHAINEORDO) ';


  if p.TypePlanning = 'PTA' then
    Result := Result + 'WHERE BEP_BTETAT="' + P.CodeOnglet + '" AND BTA_TYPEACTION="INT" AND'
  else if p.TypePlanning = 'PSF' then
    Result := Result + 'WHERE ARS_CHAINEORDO="' + P.CodeOnglet + '" AND '
  else if p.TypePlanning = 'PTR' then
    Result := Result + 'WHERE ARS_TYPERESSOURCE="' + P.CodeOnglet + '" AND '
  else if p.TypePlanning = 'PFO' then
    Result := Result + 'WHERE ARS_FONCTION1="' + P.CodeOnglet + '" AND '
  else
    Result := Result + 'WHERE ';

  Result := Result + '  ((BEP_DATEDEB>="' + UsDateTime(Datedebut) + '" ';
  Result := Result+ 'AND BEP_DATEDEB<="'  + UsDateTime(Datefin) + '") ';
  Result := Result+ 'OR (BEP_DATEFIN>="'  + USDateTime(DateDebut) + '" ';
  Result := Result+ 'AND BEP_DATEFIN<="'  + USDateTime(DateFin) + '") ';
  Result := Result+ 'OR (BEP_DATEDEB<="'  + UsDateTime(DateDebut) + '" ';
  Result := Result+ 'AND BEP_DATEFIN>="'  + UsDateTime(DateFin) + '")) ';
  Result := Result+ 'AND BTA_ASSOSRES="X" ';

  if pos(p.TypePlanning, 'PFO;PRA') = 0 then
  begin
    if not (P.AppelsTraites) then Result := Result + 'AND (((BEP_AFFAIRE <> "") AND (AFF_ETATAFFAIRE="AFF")) OR (BEP_AFFAIRE="")) ';
  end;

  Result := Result+ ' ORDER BY BEP_DATEDEB';

end;

Function ChargeRequeteActionsGRC (P :THPlanningBTP;DateDebut,DateFin : TdateTime) : String;
var DateD   : TDateTime;
    DateF   : TdateTime;
begin

	DateD := StrToDate(DateTOStr(DateDebut));
	DateF := StrToDate(DateTOStr(DateFin));

  Result := 'SELECT "ACT-GRC"           AS ORIGINEITEM, ';
  Result := Result + 'RAC_NUMACTION     AS IDEVENT, ';
  Result := Result + 'RAC_AFFAIRE       AS CODEAFFAIRE, ';
  Result := Result + 'AFF_LIBELLE       AS LIBAFFAIRE, ';
  Result := Result + 'RAC_TIERS         AS CODETIERS, ';
  Result := Result + 'RAC_AUXILIAIRE    AS AUXILIAIRE, ';
  Result := Result + 'T_LIBELLE         AS LIBTIERS, ';
  Result := Result + '""                AS NOMJURIDIQUE, ';
  Result := Result + 'C_PRENOM          AS NOM1, ';
  Result := Result + 'C_NOM             AS NOM2, ';
  Result := Result + 'T_ADRESSE1        AS ADR1, ';
  Result := Result + 'T_ADRESSE2        AS ADR2, ';
  Result := Result + 'T_ADRESSE3        AS ADR3, ';
  Result := Result + 'T_CODEPOSTAL      AS CP,  ';
  Result := Result + 'T_VILLE           AS VILLE, ';
  Result := Result + 'RAC_DATEACTION    AS DATEDEB, ';
  Result := Result + 'RAC_DATEACTION    AS DATEFIN, ';
  Result := Result + 'RAC_HEUREACTION   AS HEUREDEB, ';
  Result := Result + 'RAC_HEUREACTION   AS HEUREFIN, ';
  Result := Result + 'RAC_DUREEACTION   AS DUREE, ';
  Result := Result + '"ACT-GRC"         AS ACTION, ';
  Result := Result + 'BTA_LIBELLE       AS LIBACTION, ';
  Result := Result + 'BTA_ASSOSRES      AS ASSOSRES, ';
  Result := Result + 'RAI_RESSOURCE     AS CODERESSOURCE, ';
  Result := Result + 'ARS_EQUIPERESS    AS EQUIPE, ';
  Result := Result + 'ARS_LIBELLE       AS LIBRESSOURCE, ';
  Result := Result + 'RAC_INTERVENANT   AS MATERIEL, ';
  Result := Result + 'ARS_LIBELLE       AS LIBMATERIEL, ';
  Result := Result + 'ARS_TYPERESSOURCE AS FAMILLEMAT, ';
  Result := Result + '""                AS LIBFAMILLEMAT, ';
  Result := Result + 'RAC_ETATACTION    AS CODEETAT, ';
  Result := Result + '""                AS LIBEVENEMENT, ';
  Result := Result + '""                AS LIBREALISE, ';
  Result := Result + 'RAC_BLOCNOTE      AS BLOCNOTE, ';
  Result := Result + 'ARS_FONCTION1     AS CODEFONCTION, ';
  Result := Result + 'AFO_LIBELLE       AS LIBFONCTION, ';
  Result := Result + '"0"               AS NUMPHASE, ';
  Result := Result + '""                AS LIBPHASE, ';
  Result := Result + 'ARS_EMAIL         AS EMAIL, ';
  Result := Result + 'ARS_STANDCALEN    AS STANDCALEN, ';
  Result := Result + 'ARS_CALENSPECIF   AS CALENSPECIF, ';
  Result := Result + 'ARS_TYPERESSOURCE AS TYPERESSOURCE, ';
  Result := Result + 'ARS_CHAINEORDO    AS STYPERESSOURCE, ';
  Result := Result + 'BTR_LIBELLE       AS LIBSTYPERESSOURCE, ';
  Result := Result + 'RAC_TYPEACTION    AS TYPEACTION, ';
	Result := Result + '""                AS TYPEMVT, ';
  Result := Result + '""                AS TYPECONGE, ';
  Result := Result + 'ARS_SALARIE       AS SALARIE, ';
  Result := Result + 'AFF_DESCRIPTIF    AS DESCRIPTIF, ';
  Result := Result + '"0"               AS NUMEROADRESSE, ';
  //FV1 - 03/04/2018 : FS#3032 - ECHAFAUDAGE SERVICE : Ajouter date début et date fin du chantier
  Result := Result + 'AFF_DATEDEBUT     AS DATEDEBAFF, ';
  Result := Result + 'AFF_DATEFIN       AS DATEFINAFF, ';
  Result := Result + 'RAC_CHAINAGE, RAC_NUMLIGNE, RAC_GESTRAPPEL, RAC_DELAIRAPPEL, RAC_DATERAPPEL, RAC_NUMEROCONTACT ';
  Result := Result + 'FROM RTVACTIONS ';
  Result := Result + 'LEFT JOIN ACTIONINTERVENANT ON RAI_AUXILIAIRE=RAC_AUXILIAIRE AND RAI_NUMACTION=RAC_NUMACTION ';
  Result := Result + 'LEFT JOIN AFFAIRE           ON (AFF_AFFAIRE=RAC_AFFAIRE) ';
  Result := Result + 'LEFT JOIN RESSOURCE         ON RAI_RESSOURCE=ARS_RESSOURCE ';
  Result := Result + 'LEFT JOIN FONCTION          ON (ARS_FONCTION1=AFO_FONCTION) ';
  Result := Result + 'LEFT JOIN BFONCTION         ON (AFO_FONCTION=BFO_FONCTION) ';
  Result := Result + 'LEFT JOIN BTTYPERES         ON (ARS_CHAINEORDO=BTR_TYPRES) ';
  Result := result + 'LEFT JOIN BTETAT            ON BTA_BTETAT="ACT-GRC" ';
  Result := Result + 'WHERE ';

  //RAC_NUMACTION =' + IntToStr(NumAction) + ' AND RAC_AUXILIAIRE="' + Auxiliaire + '" AND ';

  //NB : Géré ici les affaire Réalisée/en cours ou les deux...
  if (not P.AppelsTraites) then
  	Result := Result + 'RAC_ETATACTION="PRE" '
  else
  	Result := Result + 'RAC_ETATACTION in ("PRE","REA","NRE") ';

  if DateDebut <> IDate1900 then
  begin
    Result  := Result + '  AND RAC_DATEACTION >= "'+ UsDateTime(DateDebut) + '" ';
    Result  := Result + '  AND RAC_DATEACTION <= "'+ UsDateTime(DateFin)   + '" ';
  end;

  Result := Result + '  AND BTA_ASSOSRES="X" ';

  Result    := Result + 'ORDER BY RAC_DATEACTION';

end;

Function ChargeRequeteAbsSalaries (P :THPlanningBTP;DateDebut,DateFin : TdateTime) : String;
var DateD   : TDateTime;
    DateF   : TdateTime;
begin

	DateD := StrToDate(DateTOStr(DateDebut));
	DateF := StrToDate(DateTOStr(DateFin));

  Result := 'SELECT "ABSPAIE"           AS ORIGINEITEM, ';
  Result := Result + '0                 AS IDEVENT, ';
  Result := Result + '""                AS CODEAFFAIRE, ';
  Result := Result + '""                AS LIBAFFAIRE, ';
  Result := Result + '""                AS CODETIERS, ';
  Result := Result + '""                AS LIBTIERS, ';
  Result := Result + '""                AS NOMJURIDIQUE, ';
  Result := Result + '""                AS NOM1, ';
  Result := Result + '""                AS NOM2, ';
  Result := Result + '""                AS ADR1, ';
  Result := Result + '""                AS ADR2, ';
  Result := Result + '""                AS ADR3, ';
  Result := Result + '""                AS CP,  ';
  Result := Result + '""                AS VILLE, ';
  Result := Result + 'PCN_DATEDEBUTABS  AS DATEDEB, ';
  Result := Result + 'PCN_DATEFINABS    AS DATEFIN, ';
  Result := Result + '0                 AS HEUREDEB, ';
  Result := Result + '0                 AS HEUREFIN, ';
  Result := Result + 'PCN_JOURS         AS DUREE, ';
  Result := Result + '"ABSPAIE"         AS ACTION, ';
  Result := Result + 'BTA_LIBELLE       AS LIBACTION, ';
  Result := Result + 'BTA_ASSOSRES      AS ASSOSRES, ';
  Result := Result + 'ARS_RESSOURCE     AS CODERESSOURCE, ';
  Result := Result + 'ARS_LIBELLE       AS LIBRESSOURCE, ';
  Result := Result + 'ARS_EQUIPERESS    AS EQUIPE, ';
  Result := Result + 'ARS_RESSOURCE     AS MATERIEL, ';
  Result := Result + 'ARS_LIBELLE       AS LIBMATERIEL, ';
  Result := Result + 'ARS_TYPERESSOURCE AS FAMILLEMAT, ';
  Result := Result + '""                AS LIBFAMILLEMAT, ';
  Result := Result + '""                AS CODEETAT, ';
  Result := Result + 'PCN_LIBELLE       AS LIBEVENEMENT, ';
  Result := Result + '""                AS LIBREALISE, ';
  Result := Result + '""                AS BLOCNOTE, ';
  Result := Result + 'ARS_FONCTION1     AS CODEFONCTION, ';
  Result := Result + 'AFO_LIBELLE       AS LIBFONCTION, ';
  Result := Result + '"0"               AS NUMPHASE, ';
  Result := Result + '""                AS LIBPHASE, ';
  Result := Result + 'ARS_EMAIL         AS EMAIL, ';
  Result := Result + 'ARS_STANDCALEN    AS STANDCALEN, ';
  Result := Result + 'ARS_CALENSPECIF   AS CALENSPECIF, ';
  Result := Result + 'ARS_TYPERESSOURCE AS TYPERESSOURCE, ';
  Result := Result + 'ARS_CHAINEORDO    AS STYPERESSOURCE, ';
  Result := Result + 'BTR_LIBELLE       AS LIBSTYPERESSOURCE, ';
  Result := Result + '"INT"             AS TYPEACTION, ';
 	Result := Result + 'PCN_TYPEMVT       AS TYPEMVT, ';
  Result := Result + 'PCN_TYPECONGE     AS TYPECONGE, ';
  Result := Result + 'PCN_SALARIE       AS SALARIE, ';
  Result := Result + '""                AS DESCRIPTIF, ';
  Result := Result + '"0"               AS NUMEROADRESSE, ';
  //FV1 - 03/04/2018 : FS#3032 - ECHAFAUDAGE SERVICE : Ajouter date début et date fin du chantier
  Result := Result + '"01/01/1900"      AS DATEDEBAFF, ';
  Result := Result + '"31/12/2099"      AS DATEFINAFF ';
  Result := Result + ' FROM ABSENCESALARIE ';
  Result := Result + ' LEFT JOIN RESSOURCE ON ARS_SALARIE=PCN_SALARIE ';
  Result := Result + ' LEFT JOIN FONCTION  ON (ARS_FONCTION1=AFO_FONCTION) ';
  Result := Result + ' LEFT JOIN BFONCTION ON (AFO_FONCTION=BFO_FONCTION) ';
  Result := Result + ' LEFT JOIN BTTYPERES ON (ARS_CHAINEORDO=BTR_TYPRES) ';
  Result := result + ' LEFT JOIN BTETAT    ON BTA_BTETAT="ABSPAIE" ';
  Result := Result + 'WHERE (ARS_RESSOURCE <> '') ';
  Result := Result + '  AND (' + ConstitueRequeteAbsenceConges (P,DateD,DateF) + ')';
  Result := Result + '  AND (PCN_DATEANNULATION="' + USDATETIME(Idate1900) + '") ';
  Result := Result + '  AND (PCN_VALIDRESP<>"")';
  Result := Result + '  AND BTA_ASSOSRES="X" ';
  Result := Result + 'ORDER BY PCN_DATEDEBUTABS';

end;

//Requete Planning Parc/matériel simple...
Function ChargeRequeteEventMat(P : THPlanningBTP; DateDebut, DateFin : TdateTime) : String;
begin
  //
  Result := 'SELECT "PARCMAT"           AS ORIGINEITEM, ';
  Result := Result + 'BEM_IDEVENTMAT    AS IDEVENT, ';
  Result := Result + 'BEM_AFFAIRE       AS CODEAFFAIRE, ';
  Result := Result + 'AFF_LIBELLE       AS LIBAFFAIRE, ';
  Result := Result + 'BEM_TIERS         AS CODETIERS, ';
  Result := Result + 'T_AUXILIAIRE      AS AUXILIAIRE, ';
  Result := Result + 'T_LIBELLE         AS LIBTIERS, ';
  Result := Result + 'ADR_JURIDIQUE     AS NOMJURIDIQUE, ';
  Result := Result + 'ADR_LIBELLE       AS NOM1, ';
  Result := Result + 'ADR_LIBELLE2      AS NOM2, ';
  Result := Result + 'ADR_ADRESSE1      AS ADR1, ';
  Result := Result + 'ADR_ADRESSE2      AS ADR2, ';
  Result := Result + 'ADR_ADRESSE3      AS ADR3, ';
  Result := Result + 'ADR_CODEPOSTAL    AS CP,  ';
  Result := Result + 'ADR_VILLE         AS VILLE, ';
  Result := Result + 'BEM_DATEDEB       AS DATEDEB, ';
  Result := Result + 'BEM_DATEFIN       AS DATEFIN, ';
  Result := Result + 'BEM_HEUREDEB      AS HEUREDEB, ';
  Result := Result + 'BEM_HEUREFIN      AS HEUREFIN, ';
  Result := Result + 'BEM_NBHEURE       AS DUREE, ';
  Result := Result + 'BEM_BTETAT        AS ACTION, ';
  Result := Result + 'BTA_LIBELLE       AS LIBACTION, ';
  Result := Result + 'BTA_ASSOSRES      AS ASSOSRES, ';
  Result := Result + 'BMA_RESSOURCE     AS CODERESSOURCE, ';
  Result := Result + 'ARS_LIBELLE       AS LIBRESSOURCE, ';
  Result := Result + 'ARS_EQUIPERESS    AS EQUIPE, ';
  Result := Result + 'BEM_CODEMATERIEL  AS MATERIEL, ';
  Result := Result + 'BMA_LIBELLE       AS LIBMATERIEL, ';
  Result := Result + 'BFM_CODEFAMILLE   AS FAMILLEMAT, ';
  Result := Result + 'BFM_LIBELLE       AS LIBFAMILLEMAT, ';
  Result := Result + 'BEM_CODEETAT      AS CODEETAT, ';
  Result := Result + 'BEM_LIBACTION     AS LIBEVENEMENT, ';
  Result := Result + 'BEM_LIBREALISE    AS LIBREALISE, ';
  Result := Result + 'BEM_LIBACTION     AS BLOCNOTE, ';
  Result := Result + 'ARS_FONCTION1     AS CODEFONCTION, ';
  Result := Result + 'AFO_LIBELLE       AS LIBFONCTION, ';
  Result := Result + '"0"               AS NUMPHASE, ';
  Result := Result + '""                AS LIBPHASE, ';
  Result := Result + 'ARS_EMAIL         AS EMAIL, ';
  Result := Result + 'ARS_STANDCALEN    AS STANDCALEN, ';
  Result := Result + 'ARS_CALENSPECIF   AS CALENSPECIF, ';
  Result := Result + 'ARS_TYPERESSOURCE AS TYPERESSOURCE, ';
  Result := Result + 'ARS_CHAINEORDO    AS STYPERESSOURCE, ';
  Result := Result + 'BTR_LIBELLE       AS LIBSTYPERESSOURCE, ';
  Result := Result + 'BTA_TYPEACTION    AS TYPEACTION, ';
 	Result := Result + '""                AS TYPEMVT, ';
  Result := Result + '""                AS TYPECONGE, ';
  Result := Result + 'ARS_SALARIE       AS SALARIE, ';
  Result := Result + 'AFF_DESCRIPTIF    AS DESCRIPTIF, ';
  Result := Result + '"0"               AS NUMEROADRESSE, ';
  //FV1 - 03/04/2018 : FS#3032 - ECHAFAUDAGE SERVICE : Ajouter date début et date fin du chantier
  Result := Result + 'AFF_DATEDEBUT     AS DATEDEBAFF, ';
  Result := Result + 'AFF_DATEFIN       AS DATEFINAFF ';
  Result := Result + 'FROM BTEVENTMAT ';
  Result := Result + 'LEFT JOIN BTMATERIEL   ON (BMA_CODEMATERIEL=BEM_CODEMATERIEL) ';
  Result := Result + 'LEFT JOIN BTFAMILLEMAT ON (BFM_CODEFAMILLE=BMA_CODEFAMILLE) ';
  Result := Result + 'LEFT JOIN BTETAT       ON (BTA_BTETAT=BEM_BTETAT) ';
  Result := Result + 'LEFT JOIN RESSOURCE    ON (ARS_RESSOURCE=BMA_RESSOURCE) ';
  Result := Result + 'LEFT JOIN FONCTION     ON (ARS_FONCTION1=AFO_FONCTION) ';
  Result := Result + 'LEFT JOIN BFONCTION    ON (AFO_FONCTION=BFO_FONCTION) ';  
  Result := Result + 'LEFT JOIN BTTYPERES    ON (ARS_CHAINEORDO=BTR_TYPRES) ';
  Result := Result + 'LEFT JOIN AFFAIRE      ON (AFF_AFFAIRE=BEM_AFFAIRE) ';
  Result := Result + 'LEFT JOIN TIERS        ON (T_AUXILIAIRE=BEM_TIERS) ';
  Result := Result + 'LEFT JOIN ADRESSES     ON (ADR_REFCODE=T_AUXILIAIRE) ';
  //
  if p.TypePlanning = 'PTA' then
    Result := Result + ' WHERE BEM_BTETAT="' + P.CodeOnglet + '" AND '
  else
    Result := Result + ' WHERE ';
  //
  Result := Result + '   ((BEM_DATEDEB>="'  + UsDateTime(Datedebut) + '" ';
  Result := Result + 'AND  BEM_DATEDEB<="'  + UsDateTime(Datefin)   + '") ';
  Result := Result + ' OR (BEM_DATEFIN>="'  + USDateTime(DateDebut) + '" ';
  Result := Result + 'AND  BEM_DATEFIN<="'  + USDateTime(DateFin)   + '") ';
  Result := Result + ' OR (BEM_DATEDEB<="'  + UsDateTime(DateDebut) + '" ';
  Result := Result + 'AND  BEM_DATEFIN>="'  + UsDateTime(DateFin)   + '")) ';
  Result := Result + 'AND BTA_ASSOSRES="X" ';
  //
  Result := Result + 'ORDER BY BEM_DATEDEB';

end;

//Requete Planning par famille matériel...
Function ChargeRequeteFamilleMat(P : THPlanningBTP; DateDebut, DateFin : TdateTime) : String;
begin
  //
  Result := 'SELECT "PARCMAT"           AS ORIGINEITEM, ';
  Result := Result + 'BEM_IDEVENTMAT    AS IDEVENT, ';
  Result := Result + 'BEM_AFFAIRE       AS CODEAFFAIRE, ';
  Result := Result + 'AFF_LIBELLE       AS LIBAFFAIRE, ';
  Result := Result + 'BEM_TIERS         AS CODETIERS, ';
  Result := Result + 'T_AUXILIAIRE      AS AUXILIAIRE, ';
  Result := Result + 'T_LIBELLE         AS LIBTIERS, ';
  Result := Result + 'ADR_JURIDIQUE     AS NOMJURIDIQUE, ';
  Result := Result + 'ADR_LIBELLE       AS NOM1, ';
  Result := Result + 'ADR_LIBELLE2      AS NOM2, ';
  Result := Result + 'ADR_ADRESSE1      AS ADR1, ';
  Result := Result + 'ADR_ADRESSE2      AS ADR2, ';
  Result := Result + 'ADR_ADRESSE3      AS ADR3, ';
  Result := Result + 'ADR_CODEPOSTAL    AS CP,  ';
  Result := Result + 'ADR_VILLE         AS VILLE, ';
  Result := Result + 'BEM_DATEDEB       AS DATEDEB, ';
  Result := Result + 'BEM_DATEFIN       AS DATEFIN, ';
  Result := Result + 'BEM_HEUREDEB      AS HEUREDEB, ';
  Result := Result + 'BEM_HEUREFIN      AS HEUREFIN, ';
  Result := Result + 'BEM_NBHEURE       AS DUREE, ';
  Result := Result + 'BEM_BTETAT        AS ACTION, ';
  Result := Result + 'BTA_LIBELLE       AS LIBACTION, ';
  Result := Result + 'BTA_ASSOSRES      AS ASSOSRES, ';
  Result := Result + 'BMA_RESSOURCE     AS CODERESSOURCE, ';
  Result := Result + 'ARS_LIBELLE       AS LIBRESSOURCE, ';
  Result := Result + 'ARS_EQUIPERESS    AS EQUIPE, ';
  Result := Result + 'BEM_CODEMATERIEL  AS MATERIEL, ';
  Result := Result + 'BMA_LIBELLE       AS LIBMATERIEL, ';
  Result := Result + 'BFM_CODEFAMILLE   AS FAMILLEMAT, ';
  Result := Result + 'BFM_LIBELLE       AS LIBFAMILLEMAT, ';
  Result := Result + 'BEM_CODEETAT      AS CODEETAT, ';
  Result := Result + 'BEM_LIBACTION     AS LIBEVENEMENT, ';
  Result := Result + 'BEM_LIBREALISE    AS LIBREALISE, ';
  Result := Result + 'BEM_LIBACTION     AS BLOCNOTE, ';
  Result := Result + 'ARS_FONCTION1     AS CODEFONCTION, ';
  Result := Result + 'AFO_LIBELLE       AS LIBFONCTION, ';
  Result := Result + '"0"               AS NUMPHASE, ';
  Result := Result + '""                AS LIBPHASE, ';
  Result := Result + 'ARS_EMAIL         AS EMAIL, ';
  Result := Result + 'ARS_STANDCALEN    AS STANDCALEN, ';
  Result := Result + 'ARS_CALENSPECIF   AS CALENSPECIF, ';
  Result := Result + 'ARS_TYPERESSOURCE AS TYPERESSOURCE, ';
  Result := Result + 'ARS_CHAINEORDO    AS STYPERESSOURCE, ';
  Result := Result + 'BTR_LIBELLE       AS LIBSTYPERESSOURCE, ';
  Result := Result + 'BTA_TYPEACTION    AS TYPEACTION, ';
 	Result := Result + '""                AS TYPEMVT, ';
  Result := Result + '""                AS TYPECONGE, ';
  Result := Result + 'ARS_SALARIE       AS SALARIE, ';
  Result := Result + 'AFF_DESCRIPTIF    AS DESCRIPTIF, ';
  Result := Result + '"0"               AS NUMEROADRESSE, ';
  //FV1 - 03/04/2018 : FS#3032 - ECHAFAUDAGE SERVICE : Ajouter date début et date fin du chantier
  Result := Result + 'AFF_DATEDEBUT     AS DATEDEBAFF, ';
  Result := Result + 'AFF_DATEFIN       AS DATEFINAFF ';
  Result := Result + 'FROM BTEVENTMAT ';
  Result := Result + 'LEFT JOIN BTMATERIEL   ON (BMA_CODEMATERIEL=BEM_CODEMATERIEL) ';
  Result := Result + 'LEFT JOIN BTFAMILLEMAT ON (BFM_CODEFAMILLE=BMA_CODEFAMILLE) ';
  Result := Result + 'LEFT JOIN BTETAT       ON (BTA_BTETAT=BEM_BTETAT) ';
  Result := Result + 'LEFT JOIN RESSOURCE    ON (ARS_RESSOURCE=BEM_RESSOURCE) ';
  Result := Result + 'LEFT JOIN FONCTION     ON (ARS_FONCTION1=AFO_FONCTION) ';
  Result := Result + 'LEFT JOIN BFONCTION    ON (AFO_FONCTION=BFO_FONCTION) ';  
  Result := Result + 'LEFT JOIN BTTYPERES    ON (ARS_CHAINEORDO=BTR_TYPRES) ';
  Result := Result + 'LEFT JOIN AFFAIRE      ON (AFF_AFFAIRE=BEM_AFFAIRE) ';
  Result := Result + 'LEFT JOIN TIERS        ON (T_AUXILIAIRE=BEM_TIERS) ';
  Result := Result + 'LEFT JOIN ADRESSES     ON (ADR_REFCODE=AFF_AFFAIRE) ';
  //
  Result := Result + ' WHERE   BFM_CODEFAMILLE="' + P.CodeOnglet + '" ';

  Result := Result + '   AND ((BEM_DATEDEB>="'    + UsDateTime(Datedebut) + '" ';
  Result := Result + '   AND   BEM_DATEDEB<="'    + UsDateTime(Datefin) + '") ';
  Result := Result + '    OR  (BEM_DATEFIN>="'    + USDateTime(DateDebut) + '" ';
  Result := Result + '   AND   BEM_DATEFIN<="'    + USDateTime(DateFin) + '") ';
  Result := Result + '    OR  (BEM_DATEDEB<="'    + UsDateTime(DateDebut) + '" ';
  Result := Result + '   AND   BEM_DATEFIN>="'    + UsDateTime(DateFin) + '")) ';
  Result := Result + '   AND   BTA_ASSOSRES="X" ';
  //
  Result := Result + ' ORDER BY BEM_DATEDEB';

end;

//Requete Planning par Type d'action
Function ChargeRequeteTypeAction(P : THPlanningBTP; DateDebut, DateFin : TdateTime) : String;
begin

  //
  Result := 'SELECT "PARCMAT"           AS ORIGINEITEM, ';
  Result := Result + 'BEM_IDEVENTMAT    AS IDEVENT, ';
  Result := Result + 'BEM_AFFAIRE       AS CODEAFFAIRE, ';
  Result := Result + 'AFF_LIBELLE       AS LIBAFFAIRE, ';
  Result := Result + 'BEM_TIERS         AS CODETIERS, ';
  Result := Result + 'T_AUXILIAIRE      AS AUXILIAIRE, ';
  Result := Result + 'T_LIBELLE         AS LIBTIERS, ';
  Result := Result + 'ADR_JURIDIQUE     AS NOMJURIDIQUE, ';
  Result := Result + 'ADR_LIBELLE       AS NOM1, ';
  Result := Result + 'ADR_LIBELLE2      AS NOM2, ';
  Result := Result + 'ADR_ADRESSE1      AS ADR1, ';
  Result := Result + 'ADR_ADRESSE2      AS ADR2, ';
  Result := Result + 'ADR_ADRESSE3      AS ADR3, ';
  Result := Result + 'ADR_CODEPOSTAL    AS CP,  ';
  Result := Result + 'ADR_VILLE         AS VILLE, ';
  Result := Result + 'BEM_DATEDEB       AS DATEDEB, ';
  Result := Result + 'BEM_DATEFIN       AS DATEFIN, ';
  Result := Result + 'BEM_HEUREDEB      AS HEUREDEB, ';
  Result := Result + 'BEM_HEUREFIN      AS HEUREFIN, ';
  Result := Result + 'BEM_NBHEURE       AS DUREE, ';
  Result := Result + 'BEM_BTETAT        AS ACTION, ';
  Result := Result + 'BTA_LIBELLE       AS LIBACTION, ';
  Result := Result + 'BTA_ASSOSRES      AS ASSOSRES, ';
  Result := Result + 'BMA_RESSOURCE     AS CODERESSOURCE, ';
  Result := Result + 'ARS_LIBELLE       AS LIBRESSOURCE, ';
  Result := Result + 'ARS_EQUIPERESS    AS EQUIPE, ';
  Result := Result + 'BEM_CODEMATERIEL  AS MATERIEL, ';
  Result := Result + 'BMA_LIBELLE       AS LIBMATERIEL, ';
  Result := Result + 'BFM_CODEFAMILLE   AS FAMILLEMAT, ';
  Result := Result + 'BFM_LIBELLE       AS LIBFAMILLEMAT, ';
  Result := Result + 'BEM_CODEETAT      AS CODEETAT, ';
  Result := Result + 'BEM_LIBACTION     AS LIBEVENEMENT, ';
  Result := Result + 'BEM_LIBREALISE    AS LIBREALISE, ';
  Result := Result + 'BEM_LIBACTION     AS BLOCNOTE, ';
  Result := Result + 'ARS_FONCTION1     AS CODEFONCTION, ';
  Result := Result + 'AFO_LIBELLE       AS LIBFONCTION, ';
  Result := Result + '"0"               AS NUMPHASE, ';
  Result := Result + '""                AS LIBPHASE, ';
  Result := Result + 'ARS_EMAIL         AS EMAIL, ';
  Result := Result + 'ARS_STANDCALEN    AS STANDCALEN, ';
  Result := Result + 'ARS_CALENSPECIF   AS CALENSPECIF, ';
  Result := Result + 'ARS_TYPERESSOURCE AS TYPERESSOURCE, ';
  Result := Result + 'ARS_CHAINEORDO    AS STYPERESSOURCE, ';
  Result := Result + 'BTR_LIBELLE       AS LIBSTYPERESSOURCE, ';
  Result := Result + 'BTA_TYPEACTION    AS TYPEACTION, ';
 	Result := Result + '""                AS TYPEMVT, ';
  Result := Result + '""                AS TYPECONGE, ';
  Result := Result + 'ARS_SALARIE       AS SALARIE, ';
  Result := Result + 'AFF_DESCRIPTIF    AS DESCRIPTIF, ';
  Result := Result + '"0"               AS NUMEROADRESSE, ';
  //FV1 - 03/04/2018 : FS#3032 - ECHAFAUDAGE SERVICE : Ajouter date début et date fin du chantier
  Result := Result + 'AFF_DATEDEBUT     AS DATEDEBAFF, ';
  Result := Result + 'AFF_DATEFIN       AS DATEFINAFF ';
  Result := Result + 'FROM BTEVENTMAT ';
  Result := Result + 'LEFT JOIN BTMATERIEL   ON (BMA_CODEMATERIEL=BEM_CODEMATERIEL) ';
  Result := Result + 'LEFT JOIN BTFAMILLEMAT ON (BFM_CODEFAMILLE=BMA_CODEFAMILLE) ';
  Result := Result + 'LEFT JOIN BTETAT       ON (BTA_BTETAT=BEM_BTETAT) ';
  Result := Result + 'LEFT JOIN RESSOURCE    ON (ARS_RESSOURCE=BEM_RESSOURCE) ';
  Result := Result + 'LEFT JOIN FONCTION     ON (ARS_FONCTION1=AFO_FONCTION) ';
  Result := Result + 'LEFT JOIN BFONCTION    ON (AFO_FONCTION=BFO_FONCTION) ';  
  Result := Result + 'LEFT JOIN BTTYPERES    ON (ARS_CHAINEORDO=BTR_TYPRES) ';
  Result := Result + 'LEFT JOIN AFFAIRE      ON (AFF_AFFAIRE=BEM_AFFAIRE) ';
  Result := Result + 'LEFT JOIN TIERS        ON (T_AUXILIAIRE=BEM_TIERS) ';
  Result := Result + 'LEFT JOIN ADRESSES     ON (ADR_REFCODE=AFF_AFFAIRE) ';
  //
  Result := Result + ' WHERE   BTA_TYPEACTION="PMA" AND BTA_BTETAT="' + P.CodeOnglet + '" ';
  //
  Result := Result + '   AND ((BEM_DATEDEB>="'    + UsDateTime(Datedebut) + '" ';
  Result := Result + '   AND   BEM_DATEDEB<="'    + UsDateTime(Datefin) + '") ';
  Result := Result + '    OR  (BEM_DATEFIN>="'    + USDateTime(DateDebut) + '" ';
  Result := Result + '   AND   BEM_DATEFIN<="'    + USDateTime(DateFin) + '") ';
  Result := Result + '    OR  (BEM_DATEDEB<="'    + UsDateTime(DateDebut) + '" ';
  Result := Result + '   AND   BEM_DATEFIN>="'    + UsDateTime(DateFin) + '")) ';
  Result := Result + '   AND   BTA_ASSOSRES="X" ';

  Result := Result + ' ORDER BY BEM_DATEDEB';

end;

Function ChargeRequeteRessourceChantier(P : THPlanningBTP; DateDebut, DateFin : TdateTime) : String;
begin
  //
  Result := 'SELECT "CHANTIER"          AS ORIGINEITEM, ';
  Result := Result + 'BEC_IDEVENTCHA    AS IDEVENT, ';
  Result := Result + 'BEC_AFFAIRE       AS CODEAFFAIRE, ';
  Result := Result + 'AFF_LIBELLE       AS LIBAFFAIRE, ';
  Result := Result + 'BEC_TIERS         AS CODETIERS, ';
  Result := Result + 'T_AUXILIAIRE      AS AUXILIAIRE, ';
  Result := Result + 'T_LIBELLE         AS LIBTIERS, ';
  Result := Result + 'ADR_JURIDIQUE     AS NOMJURIDIQUE, ';
  Result := Result + 'ADR_LIBELLE       AS NOM1, ';
  Result := Result + 'ADR_LIBELLE2      AS NOM2, ';
  Result := Result + 'ADR_ADRESSE1      AS ADR1, ';
  Result := Result + 'ADR_ADRESSE2      AS ADR2, ';
  Result := Result + 'ADR_ADRESSE3      AS ADR3, ';
  Result := Result + 'ADR_CODEPOSTAL    AS CP,  ';
  Result := Result + 'ADR_VILLE         AS VILLE, ';
  Result := Result + 'BEC_DATEDEB       AS DATEDEB, ';
  Result := Result + 'BEC_DATEFIN       AS DATEFIN, ';
  Result := Result + 'BEC_HEUREDEB      AS HEUREDEB, ';
  Result := Result + 'BEC_HEUREFIN      AS HEUREFIN, ';
  Result := Result + 'BEC_DUREE         AS DUREE, ';
  Result := Result + 'BEC_BTETAT        AS ACTION, ';
  Result := Result + 'BTA_LIBELLE       AS LIBACTION, ';
  Result := Result + 'BTA_ASSOSRES      AS ASSOSRES, ';
  Result := Result + 'BEC_RESSOURCE     AS CODERESSOURCE, ';
  Result := Result + 'ARS_LIBELLE       AS LIBRESSOURCE, ';
  Result := Result + 'ARS_EQUIPERESS    AS EQUIPE, ';
  Result := Result + 'BEC_RESSOURCE     AS MATERIEL, ';
  Result := Result + 'ARS_LIBELLE       AS LIBMATERIEL, ';
  Result := Result + 'ARS_TYPERESSOURCE AS FAMILLEMAT, ';
  Result := Result + 'BEC_IDAFFECT      AS LIBFAMILLEMAT, ';
  Result := Result + 'AFF_ETATAFFAIRE   AS CODEETAT, ';
  Result := Result + 'AFF_LIBELLE       AS LIBEVENEMENT, ';
  Result := Result + 'AFF_LIBELLE       AS LIBREALISE, ';
  Result := Result + 'AFF_DESCRIPTIF    AS BLOCNOTE, ';
  Result := Result + 'ARS_FONCTION1     AS CODEFONCTION, ';
  Result := Result + 'AFO_LIBELLE       AS LIBFONCTION, ';
  Result := Result + 'BEC_NUMPHASE      AS NUMPHASE, ';
  Result := Result + 'BEC_LIBPHASE      AS LIBPHASE, ';
  Result := Result + 'ARS_EMAIL         AS EMAIL, ';
  Result := Result + 'ARS_STANDCALEN    AS STANDCALEN, ';
  Result := Result + 'ARS_CALENSPECIF   AS CALENSPECIF, ';
  Result := Result + 'ARS_TYPERESSOURCE AS TYPERESSOURCE, ';
  Result := Result + 'ARS_CHAINEORDO    AS STYPERESSOURCE, ';
  Result := Result + 'BTR_LIBELLE       AS LIBSTYPERESSOURCE, ';
  Result := Result + 'BTA_TYPEACTION    AS TYPEACTION, ';
 	Result := Result + '""                AS TYPEMVT, ';
  Result := Result + '""                AS TYPECONGE, ';
  Result := Result + 'ARS_SALARIE       AS SALARIE, ';
  Result := Result + 'AFF_DESCRIPTIF    AS DESCRIPTIF, ';
  Result := Result + '"0"               AS NUMEROADRESSE, ';
  //FV1 - 03/04/2018 : FS#3032 - ECHAFAUDAGE SERVICE : Ajouter date début et date fin du chantier
  Result := Result + 'AFF_DATEDEBUT     AS DATEDEBAFF, ';
  Result := Result + 'AFF_DATEFIN       AS DATEFINAFF ';
  Result := Result + 'FROM BTEVENTCHA ';
  Result := Result + 'LEFT JOIN BTETAT    ON (BTA_BTETAT=BEC_BTETAT) ';
  Result := Result + 'LEFT JOIN AFFAIRE   ON (AFF_AFFAIRE=BEC_AFFAIRE) ';
  Result := Result + 'LEFT JOIN TIERS     ON (T_TIERS=BEC_TIERS) ';
  Result := Result + 'LEFT JOIN ADRESSES  ON (ADR_REFCODE=AFF_AFFAIRE) ';
  Result := Result + 'LEFT JOIN RESSOURCE ON (ARS_RESSOURCE=BEC_RESSOURCE)';
  Result := Result + 'LEFT JOIN FONCTION  ON (ARS_FONCTION1=AFO_FONCTION) ';
  Result := Result + 'LEFT JOIN BFONCTION    ON (AFO_FONCTION=BFO_FONCTION) ';
  Result := Result + 'LEFT JOIN BTTYPERES ON (ARS_CHAINEORDO=BTR_TYPRES) ';

  //
  Result := Result + 'WHERE ';
  //
  if P.TypePlanning = 'PFO' then  Result := Result + ' ARS_FONCTION1="' + P.CodeOnglet + '" AND ';

  Result := Result + '    ((BEC_DATEDEB>="'  + UsDateTime(Datedebut) + '" ';
  Result := Result + '  AND BEC_DATEDEB<="'  + UsDateTime(Datefin) + '") ';
  Result := Result + '   OR (BEC_DATEFIN>="' + USDateTime(DateDebut) + '" ';
  Result := Result + '  AND BEC_DATEFIN<="'  + USDateTime(DateFin) + '") ';
  Result := Result + '   OR (BEC_DATEDEB<="' + UsDateTime(DateDebut) + '" ';
  Result := Result + '  AND BEC_DATEFIN>="'  + UsDateTime(DateFin) + '")) ';
  Result := Result + '  AND BTA_ASSOSRES="X" ';

  Result := Result + ' ORDER BY BEC_DATEDEB';

end;

end.


