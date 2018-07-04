unit CPCUMULPERIODIQUE_TOF;
{-------------------------------------------------------------------------------------
    Version    |   Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
  6.50.001.005  22/06/05    JP    Création de l'unité : Migration en eAGL de QRCumulG
                                  La gestion des Journaux est reprise de PCL dans TofEdJal.Pas
  6.50.001.005  23/06/05    JP    La gestion des états chainés n'est pas reprise car elle n'est
                                  pas utilisée (après de mande à GC)
  6.50.001.006  01/07/05    JP    Gestion des caractères Joker : Les fonction de base sont
                                  définies dans TofMeth
  7.00.001.001  15/11/05   SBO    Uniformisation des traitements avec la classe TOF_CPBALANCE
                                  Limitant le nombre de requêtes éxecutées et remontant un grand
                                  nombre de procédure dans la classe mère.
--------------------------------------------------------------------------------------}

interface

uses
    StdCtrls, Controls, Classes,
    {$IFDEF EAGLCLIENT}
    MaineAGL, eQRS1,
    {$ELSE}
    FE_Main, QRS1,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
    {$ENDIF}
    SysUtils, ComCtrls, HCtrls, Ent1, HEnt1, UTOF, UTOB, HTB97, CRITEDT,
    {$IFDEF MODENT1}
    CPTypeCons,
    {$ENDIF MODENT1}
    CPBALANCE_TOF;


type
  TOF_CPCUMULPERIODIQUE = class(TOF_CPBALANCE)
    procedure OnNew                               ; override;
    procedure OnUpdate                            ; override;
    procedure OnLoad                              ; override;
    procedure OnArgument(S : string)              ; override;
    procedure SetTypeBalance                      ; override ;

    // EVT FICHE
    procedure NatureCptChanged(Sender : TObject)  ; override;
    procedure CompteOnExit(Sender: TObject)       ; override;

    // TRAITEMENT
    procedure RemplirEdtBalance                   ; override; // Remplissage de la table temporaire
    procedure UpdateCumulsCEDTBALANCE             ; override ; // MAJ des totaux de la table temporaire
    function GetSQLCumul(vInPer, vInCol: Integer): String; override; {FP 14/06/2006 FQ16854}

    {Construit la requête de constitution de la liste des comptes / sections cibles}
    function  GenererInsertCPTPourPer  ( CodPer, LibPer : string )  : string;
    function  GetConditionSQLEcr                                    : String ; override ;

    {Génération de la requête de l'état}
    function  GenererRequeteBAL                                     : string ; override ;

  private
    {Propriétés et Méthodes reprises de QRCumulG}
    fbEdt : TFichierBase;

    procedure GenPerExo;
    procedure ChangeDateCompta;

  protected
    {Propriétés et Méthodes propres à la version CWAS}
    procedure InitControles;
    procedure InitEvenements;
    procedure GereAffichage;
    procedure MajControlZoom;
    procedure SetHelpContext;
    procedure SetfbEdt(NomFiltre : string);
    procedure SetTablettes;
    procedure SetNatureEtat;

  public
    {Contrôles et évènements}
    FSoldeProg : TCheckBox;
    {Zones visibles contenant les périodes}
    FDateCpta1 : THValComboBox;
    FDateCpta2 : THValComboBox;
    {Zones invisibles contenant les dates de début et de fin de périodes}
    FDate1     : THValComboBox;
    FDate2     : THValComboBox;

    procedure ExoOnChange    (Sender : TObject); override;
    procedure FDateCptaChange(Sender : TObject);
    procedure FSoldeProgClick(Sender : TObject);
  end;

{
procedure CumulPeriodique(Lefb : TFichierBase);
procedure CumulPeriodiqueZoom(Crit : TCritEdt; Lefb : TFichierBase);
}
procedure CPLanceFiche_CumulPeriodique ( Args : String  = 'QRS1CUMGEN' ) ;
procedure CPLanceFiche_CumulPeriodiqueZoom ( Args : String  = 'QRS1CUMGEN' ) ;


implementation


uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
  {$ENDIF MODENT1}
  HStatus, HMsgBox, UtilPGI, ULibExercice, uLibWindows, ed_Tools, HQry, TofMeth,
  AGLInit, LicUtil;

// =========================================================================

procedure CPLanceFiche_CumulPeriodique ( Args : String ) ;
begin
{
  Doit être appelé avec le nom du filtre correspondant à l'état
    Generaux    : Args := 'QRS1CUMGEN';
    Auxiliaire  : Args := 'QRS1CUMAUX';
    Section     : Args := 'QRS1CUMSEC';
  Pour dire que l'appel est fait par le menu et qu'il ne faut pas traiter aCritEdt, on ajoute NON
}
  AGLLanceFiche('CP', 'CPCUMULPERIODIQUE', '', '', Args + ';NON' );
end;

// =========================================================================

procedure CPLanceFiche_CumulPeriodiqueZoom ( Args : String ) ;
begin
{
  Doit être appelé avec le nom du filtre correspondant à l'état
    fbGene : s := 'QRS1CUMGEN;';
    fbAux  : s := 'QRS1CUMAUX;';
    fbSect : s := 'QRS1CUMSEC;';

  Pour dire que l'appel est fait par zoom et qu'il faut traiter aCritEdt, on ajoute OUI
}
  AGLLanceFiche('CP', 'CPCUMULPERIODIQUE', '', '', Args + ';OUI' );
end;

// =========================================================================

(*
{---------------------------------------------------------------------------------------}
procedure CumulPeriodique(Lefb : TFichierBase);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  case Lefb of
    {Nom du filtre de l'état}
    fbGene : s := 'QRS1CUMGEN;';
    fbAux  : s := 'QRS1CUMAUX;';
    fbSect : s := 'QRS1CUMSEC;';
  end;
  {Pour dire que l'appel est fait par le menu et qu'il ne faut pas traiter aCritEdt}
  s := s + 'NON;';
  AGLLanceFiche('CP', 'CPCUMULPERIODIQUE', '', '', s);
end;

{Pour le moment inutilisé, mais sait-on jamais !!!
{---------------------------------------------------------------------------------------}
procedure CumulPeriodiqueZoom(Crit : TCritEdt; Lefb : TFichierBase);
{---------------------------------------------------------------------------------------}
var
  s : string;
  o : ClassCritEdt;
begin
  {La gestion du zoom est faite dans le OnNew de CPBALANCE_TOF.PAS}
  o := ClassCritEdt.Create;
  o.CritEdt := Crit;
  TheData := o;

  case Lefb of
    {Nom du filtre de l'état}
    fbGene : s := 'QRS1CUMGEN;';
    fbAux  : s := 'QRS1CUMAUX;';
    fbSect : s := 'QRS1CUMSEC;';
  end;
  {Pour dire que l'appel est fait par zoom et qu'il faut traiter aCritEdt}
  s := s + 'OUI;';

  AGLLanceFiche('CP', 'CPCUMULPERIODIQUE', '', '', s);
end;
 *)

{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
var
  FFiltre : string;
  ZoomOk  : Boolean;
begin

  FFiltre := ReadTokenSt(S);
  ZoomOk  := ReadTokenSt(S) = 'OUI';

  // Mise à jour de la variable fbEdt, opur le type de cumul
  SetfbEdt(FFiltre);

  // Ajout type balance
  TypeBal := balCumul ;

  inherited;

  {Mise à jour du nom des filtres si on n'est pas dans le cas d'une impression sur "zoom"}
  if not ZoomOk then
    TFQRS1(Ecran).FNomFiltre := FFiltre;

  {Mise à jour du HelpContext en fonction de fbEdt}
  SetHelpContext;
  {Mise à jour des tablettes de combos et des Edit en fonction du type de cumul}
  SetTablettes;
  {Gestion de la nature de l'état et du ParamEtat}
  SetNatureEtat;

  {Récupération des contrôles du QRS1}
  InitControles;
  {Branchement des évènements}
  InitEvenements;
  {Gestion de l'affichage en fonction de l'état}
  GereAffichage;

  {Affecte les valeurs des contrôles s'il s'agit d'un zoom}
  if ZoomOk then MajControlZoom;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.OnLoad;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  FSoldeProgClick(FSoldeProg);

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if not TestJoker(CompteDe.Text) then
    {Auto-remplissage si comptes non renseignés}
    case fbEdt of
      fbSect : begin
                 if (Trim(CompteDe.Text) = '') then
                   CompteDe.Text := GetMinCompte('SECTION', 'S_SECTION', NatureCpt.Value );
                 if (Trim(CompteA.Text) = '') then
                   CompteA.Text := GetMaxCompte('SECTION', 'S_SECTION', NatureCpt.Value );
               end;
      fbAux  : begin
                 if (Trim(CompteDe.Text) = '') then
                   CompteDe.Text := GetMinCompte('TIERS', 'T_AUXILIAIRE', NatureCpt.Value );
                 if (Trim(CompteA.Text) = '') then
                   CompteA.Text := GetMaxCompte('TIERS', 'T_AUXILIAIRE', NatureCpt.Value );
               end;
      fbGene : begin
                 if (Trim(CompteDe.Text) = '') then
                   CompteDe.Text := GetMinCompte('GENERAUX', 'G_GENERAL', NatureCpt.Value );
                 if (Trim(CompteA.Text) = '') then
                   CompteA.Text := GetMaxCompte('GENERAUX', 'G_GENERAL', NatureCpt.Value );
               end;
    end;

  {Dans CEDTBALANCE, il n'y a les données que d'un seul état par User. La clause Where se limite
   donc au User}
  SetControlText('XX_WHERE', V_PGI.User);

  // Patch car Cri_xxx sur un radio bouton ne fonctionne pas (post SocRef 721 )
  SetControlText('AFFICHAGECODE', Affichage.value ) ;

  SetControlText('EXOSQL', CRelatifVersExercice(Exercice.value));
  { FQ 19829 BVE 28.05.07 }
  if fbEdt = fbSect then
     SetControlText('QFSQL', ' AND Y_AXE = "' + GetControlText('NATURECPT') + '" AND ' + RecupWhereQualifPiece )
  else
     SetControlText('QFSQL', ' AND ' + RecupWhereQualifPiece );
  { END FQ 19829 }

  {b FP 14/06/2006 FQ16854}
  SetControlText('CHPSECTION', TSQLAnaCroise.ChampSection(NatureCpt.Value));
  {e FP 14/06/2006 FQ16854}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.OnNew;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if fbEdt = fbSect then NatureCpt.ItemIndex := 0;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.OnUpdate;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{Eventuellement si l'on branche le zoom sur les fiches Tiers, sections et généraux
 (cf.CumulPeriodiqueZoom) et que ce qui est fait dans CPBALANCE_TOF s'avérait insuffisant
{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.MajControlZoom;
{---------------------------------------------------------------------------------------}
begin
(*
  {Comptes ou section}
  CompteDe.Text := aCritEdt.Cpt1;
  CompteA.Text  := aCritEdt.Cpt2;
  {Exercice}
  Exercice.Value := CExerciceVersRelatif(aCritEdt.Exo.Code);
  {Fourchette de dates de l'édition}
  DateComptaDe.Text := DateToStr(aCritEdt.Date1);
  DateComptaA.Text := DateToStr(aCritEdt.Date2);
  {Devise}
  Devise.Value := aCritEdt.DeviseSelect;
  {Établissement}
  Etab.Value := aCritEdt.Etab;
  {Spécifiques ...}
  if fbEdt = fbAux then
    NatureCpt.Value := aCritEdt.NatureCpt {Nature de compte}

  else if fbEdt = fbSect then
    NatureCpt.Value := aCritEdt.Cum.Axe {Axe}
    *)
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.SetfbEdt(NomFiltre : string);
{---------------------------------------------------------------------------------------}
begin
       if Pos('CUMGEN', NomFiltre) > 0 then fbEdt := fbGene
  else if Pos('CUMAUX', NomFiltre) > 0 then fbEdt := fbAux
  else if Pos('CUMSEC', NomFiltre) > 0 then fbEdt := fbSect;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.SetHelpContext;
{---------------------------------------------------------------------------------------}
begin
  case fbEdt of
    fbSect : Ecran.HelpContext := 7480000;
    fbAux  : Ecran.HelpContext := 7479000;
    fbGene : Ecran.HelpContext := 7478000;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.SetNatureEtat;
{---------------------------------------------------------------------------------------}
begin
  case fbEdt of
    fbSect : TFQRS1(Ecran).NatureEtat := 'CPS';
    fbAux  : TFQRS1(Ecran).NatureEtat := 'CPA';
    fbGene : TFQRS1(Ecran).NatureEtat := 'CPG';
  end;
  {Gestion du paramétrage des états}
  TFQRS1(Ecran).ChoixEtat := True;
  TFQRS1(Ecran).ParamEtat := V_PGI.Superviseur or (V_PGI.Password = CryptageSt(DayPass(Date)));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.SetTablettes;
{---------------------------------------------------------------------------------------}
begin
  case fbEdt of
    fbSect : begin
               CompteDe .DataType := 'TZSECTION';
               CompteA  .DataType := 'TZSECTION';
               {Pour les sections on ne travaille que sur un seul axe}
               NatureCpt.Vide       := False;
               NatureCpt.DataType := 'TTAXE';
               NatureCpt.Refresh;
             end;
    fbAux  : begin
               CompteDe .DataType := 'TZTTOUS';
               CompteA  .DataType := 'TZTTOUS';
               NatureCpt.DataType := 'TTNATTIERSCPTA';
             end;
    fbGene : begin
               CompteDe .DataType := 'TZGENERAL';
               CompteA  .DataType := 'TZGENERAL';
               NatureCpt.DataType := 'TTNATGENE';
             end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.CompteOnExit(Sender : TObject);
{---------------------------------------------------------------------------------------}
var CptType : TFichierBase;
begin
  if Trim(THEdit(Sender).Text) = '' then Exit;
  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if HasJoker(Sender) then Exit;

  CptType := fbEdt ;
  if CptType = fbSect
    then CptType := AxeToFb( NatureCpt.Value ) ; // Pour les sections le CompleteAuto est fonction de l'axe

  {Complétion auto du numéro de compte si possible}
  if not CompleteAuto(Sender, CptType) then
    THEdit(Sender).ElipsisClick(Sender);
end;

{Restreint la tablette des comptes en fonction de la nature de compte sélectionnée
{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.NatureCptChanged(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin

  if fbEdt = fbSect
    then NatureAxeChanged( Sender )
    else inherited  NatureCptChanged( Sender ) ;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.InitControles;
{---------------------------------------------------------------------------------------}
begin
  FSoldeProg := TCheckBox(GetControl('CKSOLDEPROG'));
  FDateCpta1 := THValComboBox(GetControl('FDATECPTA1'));
  FDateCpta2 := THValComboBox(GetControl('FDATECPTA2'));
  FDate1     := THValComboBox(GetControl('FDATE1'));
  FDate2     := THValComboBox(GetControl('FDATE2'));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.InitEvenements;
{---------------------------------------------------------------------------------------}
begin
  FDateCpta1.OnChange := FDateCptaChange;
  FDateCpta2.OnChange := FDateCptaChange;
  FSoldeProg.OnClick  := FSoldeProgClick;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.GereAffichage;
{---------------------------------------------------------------------------------------}
begin
  case fbEdt of
    fbGene : begin
               SetControlCaption('TCOMPTE', 'Comptes &généraux de');
               Ecran.Caption := 'Cumuls périodiques des comptes généraux';
             end;
    fbAux  : begin
               SetControlCaption('TCOMPTE', 'Comptes &auxiliaires de');
               Ecran.Caption := 'Cumuls périodiques des comptes auxiliaires';
             end;
    fbSect : begin
               SetControlCaption('TCOMPTE'  , '&Sections de');
               SetControlCaption('TNATURECPT', '&Axe');
               if EstSerie(S3) then begin
                 {On cache l'axe en S3}
                 SetControlVisible('NATURECPT' , False);
                 SetControlVisible('TNATURECPT', False);
               end;
               Ecran.Caption := 'Cumuls périodiques des sections analytiques';
             end;
  end; {case fbEdt}

  UpdateCaption(Ecran);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.RemplirEdtBalance;
{---------------------------------------------------------------------------------------}
var n : Integer ;
begin
  {1. Effacement des enregistrements presents pour le user}
  DeleteCedtBalance;

  {2. Insertion de la liste des comptes/sections - Périodes cibles}
  for n := FDateCpta1.ItemIndex to FDateCpta2.ItemIndex do
    ExecuteSql(  GenererInsertCPTPourPer ( FDate1.Values[n], FDateCpta1.Items[n] ) );

  {3. Update des totaux / soldes}
  UpdateCumulsCEDTBALANCE;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.ExoOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Création des périodes comptables pour un exercice donné}
  GenPerExo;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.FDateCptaChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {Mise à jour des champs dates comptables qui sont invisibles}
  ChangeDateCompta;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.FSoldeProgClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Solde progressif est un critère pour les formules de l'état}
  if FSoldeProg.Checked then SetControlText('SOLDEPROG', 'X')
                        else SetControlText('SOLDEPROG', '-');
end;

{Création des périodes comptables pour un exercice donné
{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.GenPerExo;
{---------------------------------------------------------------------------------------}
var
  DateExo : TExoDate;
  i       : Integer;
  Annee,
  pMois,
  NbMois  : Word;
  DD      : TdateTime;
  D1, D2  : string;
BEGIN
  NbMois := 0;
  {On vide les quatre ComboBox pour commencer}
  FDateCpta1.Items.Clear;
  FDateCpta2.Items.clear;
  FDate1    .Items.Clear;
  FDate2    .Items.Clear;
  FDateCpta1.Values.Clear;
  FDateCpta2.Values.Clear;
  FDate1    .Values.Clear;
  FDate2    .Values.Clear;

  {On récupère les dates de debut et de fin de l'exercice en cours}
  QuelDateDeExo(CRelatifVersExercice(Exercice.value), DateExo);
  {On récupère le nombre de mois sur l'exercice}
  NombrePerExo(DateExo, pMois, Annee, NbMois);

  {Remplissage des combos}
  for i := 0 to NbMois - 1 do begin
    DD := PlusMois(DateExo.Deb, i);
    D1 := FormatDateTime('mmmm yyyy', DD);
    D2 := FormatDateTime('mmmm yyyy', DD);
    {Stockage du libellé}
    FDateCpta1.Items.Add(FirstMajuscule(D1));
    FDateCpta2.Items.Add(FirstMajuscule(D2));
    {Stockage de l'indice courant}
    FDateCpta1.Values.Add(IntTostr(i));
    FDateCpta2.Values.Add(IntTostr(i));
    {Stockage des dates de début et de fin de période}
    FDate1.Items.Add(DateToStr(DebutdeMois(DD)));
    FDate2.Items.Add(DateToStr(FindeMois(DD)));
    {Stockage des "codes" périodes tels qu'ils figurent dans les tables
     ECRITURE et ANALYTIQ afin de constituer les clauses where des requêtes}
    FDate1.Values.Add(IntToStr(GetPeriode(DD)));
    FDate2.Values.Add(IntToStr(GetPeriode(DD)));
  end;

  {On positionne les ComboBox}
  FDateCpta1.ItemIndex := 0;
  FDateCpta2.ItemIndex := NbMois - 1;
  FDate1.ItemIndex := FDateCpta1.ItemIndex;
  FDate2.ItemIndex := FDateCpta2.ItemIndex;

  {Mise à jour des champs dates comptables qui sont invisibles}
  ChangeDateCompta;
end;

{Mise à jour des critères qui serviront à exécutés les requêtes portant sur les soldes
 initiaux dans l'état
{---------------------------------------------------------------------------------------}
procedure TOF_CPCUMULPERIODIQUE.ChangeDateCompta;
{---------------------------------------------------------------------------------------}
begin
  {Récupération des dates correspondant aux début et fin de période de traitement}
  SetControlText('DATECOMPTABLE' , FDate1.Items[FDateCpta1.ItemIndex]);
  SetControlText('DATECOMPTABLE_', FDate2.Items[FDateCpta2.ItemIndex]);
end;

{constitution de la requête de l'état à proprement parlé
{---------------------------------------------------------------------------------------}
function TOF_CPCUMULPERIODIQUE.GenererRequeteBAL : string;
{---------------------------------------------------------------------------------------}
begin
(*
  if MultiSoc and (GetControlText('MULTIDOSSIER') <> '') then
    Result := 'SELECT SYSDOSSIER, CED_COMPTE CED_' + ChampCpt + ', CED_NATURE, CED_LIBELLE'
  else *)
  Result := 'SELECT CED_COMPTE CED_' + ChampCpt + ', CED_NATURE, CED_LIBELLE';
  Result := Result + ', CED_DEBIT1, CED_CREDIT1'
                   + ', CED_DEBIT2, CED_CREDIT2'
                   + ', CED_COMPTE2, CED_LIBELLE2 ';

  Result := Result + ' FROM CEDTBALANCE ';
  {La clause where ne porte que sur le User, car un utilisateur ne peut avoir que l'état
   courant dans la table CEDTBALANCE}
  Result := Result + ' WHERE CED_USER = "' + V_PGI.User + '"';

  Result := Result + ' ORDER BY CED_COMPTE, CED_COMPTE2';
end;

procedure TOF_CPCUMULPERIODIQUE.SetTypeBalance;
begin
 case fbEdt of
    fbSect, fbAxe1..fbAxe5 : begin
             TableEcr    := 'ANALYTIQ' ;
             PfEcr       := 'Y' ;
             TableCpt    := 'SECTION' ;
             PfCpt       := 'S' ;
             ChampCpt    := 'SECTION' ;
             ChampNatCpt := 'S_AXE' ;
             end ;

    fbAux  : begin
             TableEcr    := 'ECRITURE' ;
             PfEcr       := 'E' ;
             TableCpt    := 'TIERS' ;
             PfCpt       := 'T' ;
             ChampCpt    := 'AUXILIAIRE' ;
             ChampNatCpt := 'T_NATUREAUXI' ;
             end ;

    fbGene : begin
             TableEcr    := 'ECRITURE' ;
             PfEcr       := 'E' ;
             TableCpt    := 'GENERAUX' ;
             PfCpt       := 'G' ;
             ChampCpt    := 'GENERAL' ;
             ChampNatCpt := 'G_NATUREGENE' ;
             end ;

  end;

end;

function TOF_CPCUMULPERIODIQUE.GenererInsertCPTPourPer( CodPer, LibPer: string ) : string ;
begin
  // ===============================================
  // ==== DEBUT DE LA CLAUSE INSERT des comptes ====
  // ===============================================
  result := GetBaseInsertCpt ;

  // ============================
  // ==== SELECT des comptes ====
  // ============================
  // Champs
  result := result + 'SELECT "' + V_PGI.User + '", ' + PfCpt + '_' + ChampCpt + ', ' ;
  result := result + '"", ' + PfCpt + '_LIBELLE';
  // Montants
  result := result + ',0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0';
  // Rupture ?
  if XXRupture.Text <> ''
    then result := result + ', ' + XXRupture.Text
    else result := result + ', ""';
  // Collectif, Compte et libelle combinatoire
  result := result + ', "-", "' + CodPer + '", "' + LibPer + '"' ;

  // ================
  // ==== TABLES ====
  // ================
  result := result + ' FROM ' + GetTablePourBase( TableCpt );

  // Table Ecriture si besoin
  if ( ModeSelect.value <> 'FOU') then
    result := result + ' LEFT JOIN ' + GetTablePourBase(TableEcr) + ' ON ' + PfEcr + '_' + ChampCpt + '=' + PfCpt + '_' + ChampCpt ;

  // ===============
  // ==== WHERE ====
  // ===============
  result := result + ' WHERE ' + GetConditionSQL( True, 1 ) ;

  // Group by
  result := result + ' GROUP BY ' + PfCpt + '_' + ChampCpt + ', ' + PfCpt + '_LIBELLE' ;
  if XXRupture.Text <> '' then result := result + ', ' + XXRupture.Text;

  // Mode de sélection non soldé
  if (ModeSelect.value = 'NSL') then
    Result := Result + ' HAVING SUM(' + PfEcr + '_CREDIT) <> SUM(' + PfEcr + '_DEBIT)';

  {b FP 14/06/2006 FQ16854}
  // traduction requête pour multi-axe
  TSQLAnaCroise.TraduireRequete( NatureCpt.Value , result ) ;
  {e FP 14/06/2006}
end;

function TOF_CPCUMULPERIODIQUE.GetConditionSQLEcr: String;
begin
  result := inherited GetConditionSQLEcr ;
end;

procedure TOF_CPCUMULPERIODIQUE.UpdateCumulsCEDTBALANCE;
Var lStReq      : String ;
    i           : Integer ;
    lStPeriode  : String ;
begin

  try  // finally

    try  // except

      // -----------------------------------------
      // Chargement de la liste des comptes cibles sur les périodes sélectionnées
      for i := FDateCpta1.ItemIndex to FDateCpta2.ItemIndex do   {FQ19134 30.07.07  YMO}
        begin

        lStPeriode := FDate1.Values[ i ] ;

        // Champ Débit
        lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT1 = ( '
                + GetSQLCumul( 1, 1 ) + ' AND ' + PfEcr + '_PERIODE=' + lStPeriode + ' AND ' + PfEcr + '_ECRANOUVEAU="N" )'
                + ' WHERE CED_USER="'+V_PGI.User+'" AND CED_COMPTE2="' + lStPeriode + '"' ;
        ExecuteSQL( lStReq ) ;
        // Le problème des requêtes précédentes est que les champs sont mis à NULL si le compte n'est pas mouvementé sur la période donnée
        ExecuteSQL('UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT1 = 0 WHERE CED_DEBIT1 IS NULL AND CED_USER="' + V_PGI.User + '" AND CED_COMPTE2="' + lStPeriode + '"' ) ;

        // Champ Crédit
        lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_CREDIT1 = ( '
                + GetSQLCumul( 1, 2 ) + ' AND ' + PfEcr + '_PERIODE=' + lStPeriode + ' AND ' + PfEcr + '_ECRANOUVEAU="N" )'
                + ' WHERE CED_USER="'+V_PGI.User+'" AND CED_COMPTE2="' + lStPeriode + '"' ;
        ExecuteSQL( lStReq ) ;
        // Le problème des requêtes précédentes est que les champs sont mis à NULL si le compte n'est pas mouvementé sur la période donnée
        ExecuteSQL('UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_CREDIT1 = 0 WHERE CED_CREDIT1 IS NULL AND CED_USER="' + V_PGI.User + '" AND CED_COMPTE2="' + lStPeriode + '"' ) ;

        end ;

    // Try Except
    except
      on E : Exception do
      begin
        PgiError( E.Message, Ecran.Caption );
      end;
    end;
  // Try Finally
  finally
//    FiniMoveProgressForm ;
  end ;

end;

{b FP 14/06/2006 FQ16854}
function TOF_CPCUMULPERIODIQUE.GetSQLCumul(vInPer, vInCol: Integer): String;
begin
  result := inherited GetSQLCumul( vInPer, vInCol ) ;
  { FQ 19829 BVE 28.05.07 }
  if fbEdt = fbSect then
     result := result + ' AND Y_AXE = "' + GetControlText('NATURECPT') + '" ';
  { END FQ 19829 }

  // SBO traduction requête pour multi-axe
  TSQLAnaCroise.TraduireRequete( NatureCpt.Value , result ) ;
end;
{e FP 14/06/2006 FQ16854}
initialization
    RegisterClasses([TOF_CPCUMULPERIODIQUE]);

end.

