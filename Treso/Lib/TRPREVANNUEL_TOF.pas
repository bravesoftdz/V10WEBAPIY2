{Source de la tof TRPREVANNUEL
--------------------------------------------------------------------------------------
  Version    |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
   0.91        06/11/03  JP   création de l'unité
 1.50.000.000  22/04/04  JP   Nouvelle gestion des natures
 6.01.001.001  03/08/04  JP   Gestion des titres de colonnes FQ 10101
 6.60.001.001  20/12/05  JP   FQ 10289 : Gestion des saisies négatives en Tréso (demande de Thermo compac)
                              Il ne faut plus travailler sur les valeurs absolues mais multiplier par -1
                              pour les flux de types dépenses / débit
 7.09.001.001  25/08/06  JP   1/ Gestion du Multi sociétés
                              2/ Gestion du comparatif budgétaire
 8.01.001.009  28/03/07  JP   FQ 10423 : modifications des libellés
 8.10.001.004  08/08/07  JP   Gestion des confidentialités : Modification de la requête et ajout de BANQUECP
 8.10.001.010  17/09/07  JP   FQ 10405 : on empêche le double click sur le solde initial et final
 8.10.004.001  08/11/07  JP   FQ 10534 : ajout du test sur la classe Trésorerie
--------------------------------------------------------------------------------------}

unit TRPREVANNUEL_TOF;

interface

uses
  StdCtrls, Controls, Classes, UFicTableau,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  SysUtils, HCtrls, HEnt1, UTOB, Grids, UProcGen;

type
  TOF_TRPREVANNUEL = class (FICTAB)
    procedure OnUpdate              ; override;
    procedure OnArgument(S : string); override;
  protected
    sTypeMnt  : string;

    cbGeneral  : THValComboBox;
    cbNature   : THMultiValComboBox;
    EdPerDed1  : THEdit;
    EdPerFin1  : THEdit;
    EdPerDed2  : THEdit;
    EdPerFin2  : THEdit;
    {Méthodes}
    function  Afficher      : Integer;
    function  PrepareRequete(UnOk : Boolean) : string;
    function  GetNomColonne (Col  : Byte   ) : string; {FQ 10101}
    procedure ChargeTobListe(aTob : TOB; Chp : string);

    {Evènements}
    procedure GeneralOnChange(Sender : TObject);
    procedure PopupSimulation(Sender : TObject); override;
    procedure bGraphClick    (Sender : TObject); override;
    function  IsCelulleReinit(Col, Row : Integer) : Boolean; override;
    function  IsColRouge(Col : Integer; Ligne : Char = '"') : Boolean;
    {10/11/06 : Pour la gestion des soldes de réinitialisation : on mémorise toutes les opérations
                comprises entre}
  public
    function  IsNbColOk   : Boolean; override;
    function  CommandeDetail(CurRow : HTStrings; var TT : TOB) : Boolean; override;
  end;

function TRLanceFiche_TRPREVANNUEL(Dom, Fiche, Range, Lequel, Arguments : string) : string;

implementation

uses
  ExtCtrls, AGLInit, Commun, HMsgBox, TRGRAPHLIBRE_TOF, TRSAISIEFLUX_TOF,
  Constantes, UProcSolde, uTobDebug;

{---------------------------------------------------------------------------------------}
function TRLanceFiche_TRPREVANNUEL(Dom, Fiche, Range, Lequel, Arguments : string) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := FICTAB.FicLance(Dom, Fiche, Range, Lequel, Arguments);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRPREVANNUEL.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  Ecran.HelpContext := 150;
  InArgument := True;
  TypeDesc := dsPeriodique;
  inherited;
  LargeurCol := 180;

  TRadioButton(GetControl('DATEVAL')).OnClick := CritereClick;
  cbGeneral := THValComboBox(GetControl('GENERAL'));
  cbGeneral.OnChange := GeneralOnChange;
  EdPerDed1 := THEdit(GetControl('PERUN'));
  EdPerFin1 := THEdit(GetControl('PERUN_'));
  EdPerDed2 := THEdit(GetControl('PERDEUX'));
  EdPerFin2 := THEdit(GetControl('PERDEUX_'));

  {Initialisation de la devise}
  Devise := V_PGI.DevisePivot;

  cbGeneral.DataType := tcp_Tous;
  cbGeneral.Plus := FiltreBanqueCp(tcp_Tous, '', '');

  InArgument := False;

  {Initialisation des dates}
  SetControlText('PERUN', DebutTrimestre(PlusDate(V_PGI.DateEntree, -3, 'M')));
  SetControlText('PERUN_', FinTrimestre(PlusDate(V_PGI.DateEntree, -3, 'M')));
  SetControlText('PERDEUX', DebutTrimestre(V_PGI.DateEntree));
  SetControlText('PERDEUX_', FinTrimestre(V_PGI.DateEntree));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRPREVANNUEL.GeneralOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Gen : string;
begin
  Gen := GetControlText('GENERAL');
  if Trim(Gen) <> '' then Devise := RetDeviseCompte(Gen)
                     else Devise := V_PGI.DevisePivot;
  if Devise = '' then Devise := V_PGI.DevisePivot;
  AssignDrapeau(TImage(GetControl('IDEV')), Devise);
  SetControlCaption('DEV', Devise);
  CritereClick(nil);
end;

{28/10/04 : Cette propriété teste si le nombre de colonnes n'est pas trop important
{---------------------------------------------------------------------------------------}
function TOF_TRPREVANNUEL.IsNbColOk : Boolean;
{---------------------------------------------------------------------------------------}
begin
  {Ici pas de problèmes de colonnes à dessiner}
  Result := inherited IsNbColOk;
end;

{---------------------------------------------------------------------------------------}
function TOF_TRPREVANNUEL.CommandeDetail(CurRow : HTStrings; var TT : TOB) : Boolean;
{---------------------------------------------------------------------------------------}
var
  Deb : string;
  Fin : string;
  Lib : string;
begin
  Result := inherited CommandeDetail(CurRow, TT);

  {Code flux}
  if CurRow[PtClick.X] = '' then Exit; {Pas de montant}

  {17/09/07 : FQ 10405 : on empêche le double click sur le solde initial et final}
  if StrToChr(CurRow[COL_CODE]) in ['§', '¤'] then Exit;

  TT.SetString('GENERAL', cbGeneral.Value);
  TT.SetString('CODE', CurRow[COL_CODE]);
  TT.SetString('DATEOPE', sDateOpe);
  TT.SetInteger('APPEL', Ord(dsMensuel));
  {25/08/06 : Si on est en globalisé, on récupère la filtre sur la nature de la requête ...}
  if ckGlobal.Checked then TT.SetString('NATURE', Nature)
  {... sinon, on constitue le filtre en fonction de la ligne de sélection}
                      else TT.SetString('NATURE', GetNatureFromBudget(CurRow[COL_TYPE], CurRow[COL_CODE], CurRow[COL_NATURE]));

  if rbDateOpe.Checked then Lib := Lib + ' en date d''opération'
                       else Lib := Lib + ' en date de valeur';
  TT.SetString('LIBELLE', CurRow[COL_LIBELLE] + Lib);

  {Récupération des bornes de datation}
  if PtClick.X = COL_DATEDE then begin
    Deb := EdPerDed1.Text;
    Fin := EdPerFin1.Text;
  end else begin
    Deb := EdPerDed2.Text;
    Fin := EdPerFin2.Text;
  end;

  TT.SetDateTime('DATEDEB', StrToDate(Deb));
  TT.SetDateTime('DATEFIN', StrToDate(Fin));
  TT.SetString('REGROUP', GetControlText('REGROUPEMENT'));
  Result := True;
end;

{Lancement de l'écran de saisie des écritures de simulation
{---------------------------------------------------------------------------------------}
procedure TOF_TRPREVANNUEL.PopupSimulation(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  General : string;
begin
  inherited;
  GetPosition;
  {On est sur une colonne Date}
  if (PtClick.X >= COL_DATEDE) and (PtClick.Y > 0) then begin
    General := cbGeneral.Value; // ?
    if PtClick.X = COL_DATEDE then
      TRLanceFiche_SaisieFlux('TR', 'TRSAISIEFLUX', '', '', General + ';LIBRE;' + EdPerDed1.Text + ';' + EdPerFin1.Text + ';')
    else
      TRLanceFiche_SaisieFlux('TR', 'TRSAISIEFLUX', '', '', General + ';LIBRE;' + EdPerDed2.Text + ';' + EdPerFin2.Text + ';');

    MouetteClick;
  end;

  (* Pour le moment on clique sur la mouette. Si avec le remplissage de la table TRECRITURE cela
     devenait fastidieux, on mettrait à jour directement la grille
  {Création de la tob contenant les infos de base de la navoulle écriture}
  aTob  := TOB.Create('*-*', nil, -1);
  aTob.AddChampSup('TYPEFLUX');
  aTob.AddChampSup('CODEFLUX');
  aTob.AddChampSup('LIBELLE');
  aTob.AddChampSup('MONTANT');
  {Tob sur laquelle on se trouve}
  TobGrid.FindFirst(['CODEFLUX'], [Grid.Cells[1, PtClick.Y]], True);

  TheTob := aTob;
  *)
end;

{Constitution de la tob contenant les données avant de la mettre dans la grille
{---------------------------------------------------------------------------------------}
function TOF_TRPREVANNUEL.Afficher : Integer;
{---------------------------------------------------------------------------------------}
var
  TobL,
  TobG     : TOB;
  Sens     : string;
  TypeFlux : string;
  Mois     : string;
  DtD1, DtD2 : TDateTime;
//  DtF1, DtF2 : TDateTime;

    {-------------------------------------------------------------------------}
    function RecupDateDep(Premier, Init : Boolean) : string;
    {-------------------------------------------------------------------------}
    var
      aDate : TDateTime;
    begin
      Result := DateToStr(V_PGI.DateEntree);
      if Premier then aDate := StrToDate(GetControlText('PERUN'))
                 else aDate := StrToDate(GetControlText('PERDEUX'));
      if Init then Result := DateToStr(aDate)
              else Result := DateToStr(aDate - 1);
    end;

    {-------------------------------------------------------------------------}
    procedure CreeTobDetail(TypeFlux, CodeFlux, LibelleFlux, Nat : string);
    {-------------------------------------------------------------------------}
    begin
      TobG := TOB.Create('', TobGrid, -1);
      TobG.AddChampSupValeur('TYPEFLUX', TypeFlux);
      TobG.AddChampSupValeur('CODEFLUX', CodeFlux);
      TobG.AddChampSupValeur('NATURE'  , Nat);
      TobG.AddChampSupValeur('LIBELLE', LibelleFlux);
      {Ajout des colonnes dates avec montant}
      TobG.AddChampSupValeur(GetNomColonne(1), '');
      TobG.AddChampSupValeur(GetNomColonne(2), '');
    end;

    {-------------------------------------------------------------------------}
    procedure CreeTobTotal(Demi : Boolean = False);
    {-------------------------------------------------------------------------}
    var
      n   : Integer;
      S   : string;
      Mnt : Double;
    begin
      if Sens = 'C' then S := '+'
                    else S := '-';

      {Création des tobs total recettes/dépenses}
      if Demi then begin
        if S = '+' then CreeTobDetail('£', '+', TraduireMemoire('Total des encaissements'), '')
                   else CreeTobDetail('£', '-', TraduireMemoire('Total des décaissements'), '');
      end
      else
        {Création de la tob total d'un type de flux}
        CreeTobDetail(S, '', TraduireMemoire('Total'), '');

      {Cumule chaque colonne date pour le type flux courant}
      for n := 1 to 2 do begin
        Mois := GetNomColonne(n);
        {Cumul dépenses/recettes}
        if Demi then begin
          if Sens = 'C' then Mnt := TobGrid.Somme(Mois, ['TYPEFLUX'], ['+'], False)
                        else Mnt := TobGrid.Somme(Mois, ['TYPEFLUX'], ['-'], False);
        end
        else
          {Cumul d'un type de flux}
          Mnt := TobGrid.Somme(Mois, ['TYPEFLUX'], [TypeFlux], False);

        TobG.PutValue(Mois, Mnt);
      end;
    end;
            
    {-------------------------------------------------------------------------}
    function PutSoldeInit(n : Byte; var TobS : TOB; Dt : TDateTime) : Double;
    {-------------------------------------------------------------------------}
    begin
      {09/11/06 : GetSoldeMillesime ne gère pas le filtre sur les dossiers}
      if (GetControlText('REGROUPEMENT') <> '') and (cbGeneral.Value = '') then
        Result := GetSoldeInitDossier(GetControlText('REGROUPEMENT'), DateToStr(DebutAnnee(Dt)), Nature, not rbDateOpe.Checked)
      else
        {Récupération du solde forcé - opérations du 01/01}
        Result := GetSoldeMillesime(cbGeneral.Value, DateToStr(DebutAnnee(Dt)), Nature, not rbDateOpe.Checked, True);
      {Mise à jour du solde, sachant que les soldes initiaux ne peuvent être que qu'en début de période}
      TobS.SetDouble(GetNomColonne(n), Result);
      {Si Dt est un date de fin de période, on met result à zéro
       13/11/06 : Pour le moment, on ne met pas de réinitialisation sur les totaux
      if ((Dt = DtF1 + 1) and (n = 1)) or (Dt = DtF2 + 1) then Result := 0.0001;}
    end;

var
  S, CodeFlux : string;
  n, p        : Integer;
  Montant     : Double;
  CreerOk     : Boolean;
  OldNature   : string;
begin
  DtD1 := StrToDate(EdPerDed1.Text);
  DtD2 := StrToDate(EdPerDed2.Text);
//  DtF1 := StrToDate(EdPerFin1.Text);
  //DtF2 := StrToDate(EdPerFin2.Text);

    TobListe.Detail.Sort('TTL_SENS;TFT_TYPEFLUX;') ;
  CreeTobDetail('¥', '¤', 'Solde initial', '');

  TypeFlux := '';
  Sens := 'C';
  Result := 0;
  CreerOk := False;

  {JP 29/04/04 : On regarde s'il y a des flux négatifs autre que la réinitialisation}
  TobG := TobListe.FindFirst(['TTL_SENS'], ['D'], True);
  while TobG <> nil do begin
    if (TobG.GetString('TFT_TYPEFLUX') <> CODEREGULARIS) then begin
      CreerOk := True;
      TobG := nil;
      Break;
    end;
    TobG := TobListe.FindNext(['TTL_SENS'], ['D'], True);
  end;

  {********************************************************************}
  {******************  REMPLISSAGE DE LA TOB GRILLE  ******************}
  {********************************************************************}
  for p := 0 to TobListe.Detail.Count - 1 do begin
    TobL := TobListe.Detail[p];

    S := TobL.GetValue('TFT_TYPEFLUX');
    {Changement de type flux}
    if S <> TypeFlux then begin

      {Changement de sens}
      if Sens <> TobL.GetValue('TTL_SENS') then begin
        {Il y a quelquechose avant et soit on a des dépenses autres que la réinitialisation
                                      soit le premier janvier est le premier jour de la seconde période}
        if (TypeFlux <> '') and (CreerOk or (DateToMois(EdPerDed2.Text) = DateToMois(DateToStr(DebutAnnee(DtD2))))) then begin
          CreeTobTotal;
          {Création de la tob total des recettes}
          CreeTobTotal(True);
          Sens := 'D';
          {Création d'une tob fille vide pour séparer les sens}
          CreeTobDetail('*', '', '', '');
          Result := TobGrid.Detail.Count; {Position pour mise en forme (=TobG.GetIndex+1)}
        end
        else if (CreerOk or (DateToMois(EdPerDed2.Text) = DateToMois(DateToStr(DebutAnnee(DtD2))))) then
          Sens := 'D';
      end
      else if (TypeFlux <> '') then begin
        if CodeFlux <> CODEREGULARIS then CreeTobTotal;
      end;

      TypeFlux := S;
      if TobL.GetValue('TTL_LIBELLE') = CODETEMPO then begin
        if Sens = 'C' then CreeTobDetail('', '', 'Encaissements comptables', '')
                      else CreeTobDetail('', '', 'Décaissements comptables', '');
      end
      else
        CreeTobDetail('', '', TobL.GetValue('TTL_LIBELLE'), '');
    end;
    //tobdebug(TobL);
    if ckGlobal.Checked then OldNature := ''
                        else OldNature := TobL.GetValue('TE_NATURE');

    CodeFlux := TobL.GetValue('TE_CODEFLUX');
    {Si l'algo est correct, il y une TOB Fille par flux : On crée les Tob Filles de la grille}
    CreeTobDetail(TypeFlux, CodeFlux, TobL.GetValue('TFT_LIBELLE'), OldNature);

    {JP 20/12/05 : FQ 10289 : Pour pouvoir gérer les écritures saisies négativement en Tréso, on ne peut
                   travailler sur les valeur absolue (cf suivi mensuel : demande de Thermo Compac)
    Montant := Abs(Valeur(TobL.GetString(GetNomColonne(1))));}
    if Sens = 'C' then Montant := Valeur(TobL.GetString(GetNomColonne(1)))
                  else Montant := Valeur(TobL.GetString(GetNomColonne(1))) * -1;
    TobG.PutValue(GetNomColonne(1), Montant);
    if Sens = 'C' then Montant := Valeur(TobL.GetString(GetNomColonne(2)))
                  else Montant := Valeur(TobL.GetString(GetNomColonne(2))) * -1;
    TobG.PutValue(GetNomColonne(2), Montant);
  end; {Boucle for TobListe}

  {Création de la tob fille qui contiendra le total des opérations de la période pour le dernier flux}
  if CreerOk and not (StrToChr(TobG.GetString('TYPEFLUX')) in ['+', '-']) then
    CreeTobTotal;

  {********************************************************************}
  {******************  GESTION DES SOLDES TOTAUX **********************}
  {********************************************************************}

  {Création de la tob total pour le total des dépenses}
  CreeTobTotal(True);

  {Création de la dernière tob fille qui contiendra le solde total}
  CreeTobDetail('¥', '§', TraduireMemoire('Solde final'), '');

  {Calcul des soldes Pour chaque période}
  for n := 1 to 2 do begin
    Mois := GetNomColonne(n);
    TobL := TobGrid.Detail[0];

    {Si le début de millésime figure dans la fourchette de date}
    if IsColRouge(n) then begin
      if n = 1 then begin
        Montant := PutSoldeInit(n, TobL, DtD1);
        {13/11/06 : Pour le moment, on ne met pas de réinitialisation sur les totaux
        if DebutAnnee(DtF1 + 1) = (DtF1 + 1) then Montant := PutSoldeInit(n, TobG, DtF1 + 1);}
      end
      else begin
        Montant := PutSoldeInit(n, TobL, DtD2);
        {13/11/06 : Pour le moment, on ne met pas de réinitialisation sur les totaux
        if DebutAnnee(DtF2 + 1) = (DtF2 + 1) then Montant := PutSoldeInit(n, TobG, DtF2 + 1);}
      end;
      {Si l'on est sur une période de fin, on ne fait pas la somme des totaux}
      if StrFPoint(Montant) = '0.0001' then Continue;
    end
    else begin
      {Récupération du solde de la veille du début de chacune des périodes}
      if cbGeneral.Value = '' then begin
        {09/11/06 : Si on travaille sur un dossier, sans filtre sur un compte}
        if GetControlText('REGROUPEMENT') <> '' then
          Montant := GetSoldeInitDossier(GetControlText('REGROUPEMENT'), RecupDateDep(n = 1, True), Nature, not rbDateOpe.Checked)
        else
          Montant := GetSoldeInit(cbGeneral.Value, RecupDateDep(n = 1, True), Nature, not rbDateOpe.Checked);
      end
      else
      if Nature = '' then begin
        if rbDateOpe.Checked then
          Montant := GetSolde(cbGeneral.Value, RecupDateDep(n = 1, False), S)
        else
          Montant := GetSoldeValeur(cbGeneral.Value, RecupDateDep(n = 1, False), S);
      end
      else
        Montant := GetSoldeInit(cbGeneral.Value, RecupDateDep(n = 1, True), Nature, not rbDateOpe.Checked);

      {Solde du début de période, affecté à la première tob fille}
      TobGrid.Detail[0].PutValue(Mois, Montant);
    end;

    {Calcul du solde à la fin du jour}
    Montant := Montant + TobGrid.Somme(Mois, ['TYPEFLUX'], ['+'], False) - TobGrid.Somme(Mois, ['TYPEFLUX'], ['-'], False);
    {On affecte le solde à la dernière tob fille}
    TobG.PutValue(Mois, Montant);
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_TRPREVANNUEL.PrepareRequete(UnOk : Boolean) : string;
{---------------------------------------------------------------------------------------}
var
  Clause : string;
  ChpGlo : string;
begin
  if cbGeneral.Value <> '' then begin
    Clause := 'AND TE_GENERAL="' + cbGeneral.Value + '" ';
    sTypeMnt := 'TE_MONTANTDEV';
  end
  else
    sTypeMnt := 'TE_MONTANT';

  Clause := Clause + Nature;
  {25/08/06 : Gestion du filtre sur le dossier}
  if IsTresoMultiSoc and (GetControlText('REGROUPEMENT') <> '') then
    Clause := Clause + ' AND ' + GetWhereDossier;

  if UnOk then
    Clause := Clause + 'AND ' + sDateOpe + ' >= "' + UsDateTime(StrToDate(EdPerDed1.Text)) +
                     '" AND ' + sDateOpe + ' <= "' + UsDateTime(StrToDate(EdPerFin1.Text)) + '" '
  else
    Clause := Clause + 'AND ' + sDateOpe + ' >= "' + UsDateTime(StrToDate(EdPerDed2.Text)) +
                     '" AND ' + sDateOpe + ' <= "' + UsDateTime(StrToDate(EdPerFin2.Text)) + '" ';

  Clause := Clause + GetWhereConfidentialite;

  {Gestion des montants globalisés}
  if ckGlobal.Checked then ChpGlo := ''
                      else ChpGlo := ', TE_NATURE';

  Result := 'SELECT TRECRITURE.TE_CODEFLUX, SUM(TRECRITURE.' + sTypeMnt + ') AS TE_MONTANT' + ChpGlo +
        ', FLUXTRESO.TFT_TYPEFLUX, MAX(FLUXTRESO.TFT_LIBELLE) AS TFT_LIBELLE,' +
        'TYPEFLUX.TTL_SENS, MAX(TYPEFLUX.TTL_LIBELLE) AS TTL_LIBELLE ' +
        'FROM TRECRITURE, FLUXTRESO, TYPEFLUX, BANQUECP ' +
        'WHERE TE_CODEFLUX = TFT_FLUX AND TFT_TYPEFLUX = TTL_TYPEFLUX AND BQ_CODE = TE_GENERAL ' + Clause +
        ' GROUP BY TTL_SENS,TFT_TYPEFLUX, TE_CODEFLUX ' + ChpGlo +
        ' UNION ALL ' +
        'SELECT  TRECRITURE.TE_CODEFLUX, SUM(TRECRITURE.' + sTypeMnt + ') AS TE_MONTANT' + ChpGlo +
        ', RUBRIQUE.RB_TYPERUB TFT_TYPEFLUX, MAX(RUBRIQUE.RB_LIBELLE) AS TFT_LIBELLE,' +
        'RUBRIQUE.RB_SIGNERUB TTL_SENS, "' + CODETEMPO + '" AS TTL_LIBELLE ' +
        'FROM TRECRITURE, RUBRIQUE, BANQUECP ' +
        {08/11/07 : FQ 10534 : ajout du test sur la classe Trésorerie}
        'WHERE TE_CODEFLUX = RB_RUBRIQUE AND RB_CLASSERUB = "TRE" AND BQ_CODE = TE_GENERAL ' + Clause +
        ' GROUP BY RB_SIGNERUB, RB_TYPERUB, TE_CODEFLUX ' + ChpGlo +
        {11/04/05 : FQ 10238 : on trie sur les codes flux -- 28/06/05 : et éventuellement par nature}
        ' ORDER BY TTL_SENS, TFT_TYPEFLUX, TE_CODEFLUX' + ChpGlo;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRPREVANNUEL.ChargeTobListe(aTob : TOB; Chp : string);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  T : TOB;
  S : string;
begin
  for n := 0 to aTob.Detail.Count - 1 do begin
    {On n'affiche plus les flux de réinitialisation}
    if aTob.Detail[n].GetString('TE_CODEFLUX') = CODEREGULARIS then Continue;

    if not ckGlobal.Checked then begin
      S := GetNatureToBudget(aTob.Detail[n].GetString('TTL_LIBELLE'), aTob.Detail[n].GetString('TE_NATURE'));
      T := TobListe.FindFirst(['TE_CODEFLUX', 'TE_NATURE'], [aTob.Detail[n].GetString('TE_CODEFLUX'), S], True);
    end
    else
      T := TobListe.FindFirst(['TE_CODEFLUX'], [aTob.Detail[n].GetValue('TE_CODEFLUX')], True);

    {Si la tob existe déjà, c'est que l'on traite la deuxième période}
    if (T <> nil) then
      T.SetDouble(Chp, T.GetDouble(Chp) + aTob.Detail[n].GetDouble('TE_MONTANT'))
    else begin
      T := TOB.Create('***', TobListe, -1);
      T.AddChampSupValeur('TE_CODEFLUX', aTob.Detail[n].GetValue('TE_CODEFLUX'));
      T.AddChampSupValeur('TFT_TYPEFLUX', aTob.Detail[n].GetValue('TFT_TYPEFLUX'));
      T.AddChampSupValeur('TFT_LIBELLE', aTob.Detail[n].GetValue('TFT_LIBELLE'));
      T.AddChampSupValeur('TTL_SENS', aTob.Detail[n].GetValue('TTL_SENS'));
      T.AddChampSupValeur('TTL_LIBELLE', aTob.Detail[n].GetValue('TTL_LIBELLE'));
      if not ckGlobal.Checked then
        T.AddChampSupValeur('TE_NATURE', S);
      {JP 05/05/04 : l'algo était mal conçu et on pouvait se retrouver avec une TobListe sans le champ
                     PERIODE 1 ou PERIODE 2. Maintenant on crée les deux champs ensembles}
      T.AddChampSup(GetNomColonne(1), True);
      T.AddChampSup(GetNomColonne(2), True);
      T.SetDouble(Chp, aTob.Detail[n].GetDouble('TE_MONTANT'));
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRPREVANNUEL.OnUpdate;
{---------------------------------------------------------------------------------------}
var
  Q      : TQuery;
  SQL    : string;
  I      : Integer;
  NSepar : Integer;
  aTob   : TOB;
begin
  if (Trim(EdPerDed1.Text) = '/  /') or (Trim(EdPerDed2.Text) = '/  /') or
     (Trim(EdPerFin1.Text) = '/  /') or (Trim(EdPerFin2.Text) = '/  /') then begin
    HShowMessage('0;Budget périodique;Veuillez saisir les dates des périodes.;W;O;O;O;', '', '');
    Exit;
  end;

  {10/11/06 : Vu avec OG ce jour : on empèche l'utilisateur de chosir une période à cheval sur un millésime}
  if (StrToDate(EdPerDed1.Text) < DebutAnnee(StrToDate(EdPerFin1.Text))) and (StrToDate(EdPerFin1.Text) >= DebutAnnee(StrToDate(EdPerFin1.Text))) or
     (StrToDate(EdPerDed2.Text) < DebutAnnee(StrToDate(EdPerFin2.Text))) and (StrToDate(EdPerFin2.Text) >= DebutAnnee(StrToDate(EdPerFin2.Text))) then begin
    HShowMessage('0;' + Ecran.Caption + ';Les périodes ne peuvent être à cheval sur deux millésimes.;W;O;O;O;', '', '');
    EdPerDed1.SetFocus;
    Exit;
  end;
  inherited;
  {Maj des devises}
  InArgument := True; {Pour ne pas lancer une recherche et faire une boucle sans fin}
  GeneralOnChange(cbGeneral);
  InArgument := False;

  if (GetControlText('MCNATURE') <> '') and (Pos('<<', GetControlText('MCNATURE')) = 0) then
    Nature := 'AND TE_NATURE IN (' + GetClauseIn(GetControlText('MCNATURE')) + ') '
  else
    Nature := '';

  {Requête de la première période}
  aTob := TOB.Create('1', nil, 0);
  try
    SQL := PrepareRequete(True);
    Q := OpenSQL(SQL, True);
    aTob.LoadDetailDB('1', '', '', Q, False, True);
    ChargeTobListe(aTob, GetNomColonne(1));
    Ferme(Q);

    {Requête de la deuxième période}
    aTob.ClearDetail;
    SQL := PrepareRequete(False);
    Q := OpenSQL(SQL, True);
    aTob.LoadDetailDB('2', '', '', Q, False, True);
    ChargeTobListe(aTob, GetNomColonne(2));
    Ferme(Q);
  finally
    FreeAndNil(aTob);
  end;

  NSepar := Afficher;

  Grid.ColCount := COL_DATEDE + 2;
  {On cache les colonnes d Code et type de flx}
  Grid.ColWidths[COL_TYPE] := -1;
  Grid.ColWidths[COL_NATURE] := -1;
  Grid.ColWidths[COL_CODE] := -1;
  Grid.ColWidths[COL_LIBELLE] := 150;
  {Formatage des montant : '#,##0.00'; Euros}
  SQL := StrFMask(CalcDecimaleDevise(Devise), '', True);
  {Formatage des dates}
  for I := COL_DATEDE to COL_DATEDE + 1 do begin
    Grid.ColWidths[I] := 150;
    Grid.ColFormats[I] := SQL;
    Grid.ColAligns[I] := taRightJustify;
  end;

  TobGrid.PutGridDetail(Grid, True, True, '', True);

  Grid.Cells[COL_LIBELLE, 0] := '';

  if NSepar > 0 then // Un séparateur est présent
    Grid.RowHeights[NSepar] := 3;

  Grid.Col := COL_DATEDE;
  Grid.Row := Grid.FixedRows;
  Grid.SetFocus;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRPREVANNUEL.bGraphClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Titre  : string;
  sSolde : string;
  sDev   : string;
  sGraph : string;
  sDate1 : string;
  sDate2 : string;
begin
//  inherited;
  if (cbGeneral.Value = '') or ( Pos('>>', cbGeneral.Value) > 0) then begin
    Titre := 'Évolution comparative des flux sur deux périodes';
    sDev  := 'Les montants sont éxprimés en ' + RechDom('TTDEVISE', V_PGI.DevisePivot, False);
  end
  else begin
    Titre := 'Mouvements sur le compte : ' + cbGeneral.Items[cbGeneral.ItemIndex];
    sDev  := 'Les montants sont éxprimés en ' + RechDom('TTDEVISE', GetControlText('DEV'), False);
  end;

  if rbDateOpe.Checked then sSolde := 'Les montants sont calculés en date d''opération'
                       else sSolde := 'Les montants sont calculés en date de valeur';

  sDate1 := 'Du ' + GetControlText('PERUN') + ' au ' + GetControlText('PERUN_');
  sDate2 := 'Du ' + GetControlText('PERDEUX') + ' au ' + GetControlText('PERDEUX_');

  {Constitution de la chaine qui servira aux paramètres de "LanceGraph"}
  sGraph := 'TE_CODEFLUX;' + GetNomColonne(1) + ';' + GetNomColonne(2) + ';';

  TheTob := TobListe;
  {!!! Mettre sGraph en fin à cause des ";" qui traine dans la chaine}
  TRLanceFiche_HistoLibre(Titre + ';' + sDev + ';' + sSolde + ';' + sDate1 + ';' + sDate2 + ';' + sGraph);
end;

{---------------------------------------------------------------------------------------}
function TOF_TRPREVANNUEL.IsCelulleReinit(Col, Row : Integer) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited IsCelulleReinit(Col, Row);
  if Grid.Cells[COL_Type, Row] = '¥' then
    if StrToChr(Grid.Cells[COL_CODE, Row]) in ['§', '¤'] then
      Result := IsColRouge(Col - COL_DATEDE + 1, StrToChr(Grid.Cells[COL_CODE, Row]));
end;

{FQ 10101 : Nom des colonnes PERIODE 1 et Période 2 au lieu d'avoir les fourchettes de dates
{---------------------------------------------------------------------------------------}
function TOF_TRPREVANNUEL.GetNomColonne(Col : Byte) : string;
{---------------------------------------------------------------------------------------}
begin
  if Col = 1 then Result := 'Du ' + EdPerDed1.Text + ' au ' + EdPerFin1.Text
             else Result := 'Du ' + EdPerDed2.Text + ' au ' + EdPerFin2.Text;
end;

{---------------------------------------------------------------------------------------}
function TOF_TRPREVANNUEL.IsColRouge(Col : Integer; Ligne : Char = '"') : Boolean;
{---------------------------------------------------------------------------------------}
var
  DtDeb : TDateTime;
//  DtFin : TDateTime;
begin
  Result := False;
  if not Between(Col, 1, 2) then Exit;

  if Col = 1 then begin
    DtDeb := StrToDate(EdPerDed1.Text);
  //  DtFin := StrToDate(EdPerFin1.Text);
  end else begin
    DtDeb := StrToDate(EdPerDed2.Text);
    //DtFin := StrToDate(EdPerFin2.Text);
  end;

       if Ligne = '¤' then Result := (DtDeb = DebutAnnee(DtDeb))
  {13/11/06 : Pour le moment, on ne met pas de réinitialisation sur les totaux     
  else if Ligne = '§' then Result := ((DtFin + 1) = DebutAnnee(DtFin + 1))}
  {02/02/07 : je ne sais pas trop ce que j'ai voulu faire, mais cela fausse tout si une des deux
              périodes a le 31/12/xx comme fin de période !!!
  else if Ligne = '"' then Result := (DtDeb = DebutAnnee(DtDeb)) or ((DtFin + 1) = DebutAnnee(DtFin + 1));}
  else if Ligne = '"' then Result := (DtDeb = DebutAnnee(DtDeb));
end;

initialization
  RegisterClasses([TOF_TRPREVANNUEL]);

end.

