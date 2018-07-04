{-------------------------------------------------------------------------------------
  Version   |  Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
               12/02/02  BT   Création de l'unité
   0.91        04/11/03  JP   Pour l'affichage en solde on travaille maintenant sur le Code Banque
                              et non plus sur EtabBq cf TypeSoldeClick, GeneralOnChange, OnUpdate
   0.91        20/11/03  JP   On change encore tout : pour les soldes, on peut avoir une sélection
                              de banques !! cf 20/11/03
 1.05.001      19/03/04  JP   Problème dans l'affichage par solde : on faisait les totaux par banque
                              avant d'avoir remplir les soldes pour les jours sans mouvement.
 1.2X.000      21/04/04  JP   Un certain nombre de fonctionnalités ont été remontées dans l'ancêtre
                              avec la nouvelle gestion des natures et de l'échéancier
 6.30.001.002  08/03/05  JP   FQ 10217 : gestion de la date d'opération lors de l'appel de la saisie.
                              ATTENTION : si on met en place le suivi hebdomadaire, il faudra prendre
                              des précautions !!!! cf. UFicTableau.GetCurDate
 6.30.001.007  18/04/05  JP   FQ 10237 : Gestion de la périodicité sur les fiches de suivi
 6.51.001.002  05/12/05  JP   FQ 10317 : Gestion des soldes forcés pour les périodicité tp_7 et tp_30
 6.60.001.001  20/12/05  JP   FQ 10289 : Gestion des saisies négatives en Tréso (demande de Thermo compac)
                              Il ne faut plus travailler sur les valeurs absolues mais multiplier par -1
                              pour les flux de types dépenses / débit
 7.00.001.001  16/01/06  JP   FQ 10328 : Refonte de la gestion des devises
 7.05.001.001  20/10/06  JP   Mise en place des filres multi sociétés
 8.10.001.004  08/08/07  JP   Gestion des confidentialités : Modification de la requête et ajout de BANQUECP
--------------------------------------------------------------------------------------}

unit TofDetailSuivi ;

interface

uses
  StdCtrls, Controls, Classes, UFicTableau,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  SysUtils, HCtrls, HEnt1, UTOB, Grids;

type
  TOF_DETAILSUIVI = class(FICTAB)
    procedure OnUpdate              ; override;
    procedure OnArgument(S : string); override;
  protected
    cbGeneral  : THValComboBox;
    cbCompte   : THMultiValComboBox;
    rbFlux     : TRadioButton;

    {JP 15/01/04 : mise en rapport des critères de date de la fiche de suivi et du détail du suivi}
    fs_Date : string;
    fs_Jour : Integer;

    {Méthodes}
    function  ParFlux        : Integer;
    {Evènements}
    procedure GeneralOnChange(Sender : TObject);
    procedure TypeSoldeClick (Sender : TObject);
    procedure PopupSimulation(Sender : TObject); override;
    function  IsCelulleReinit(Col, Row : Integer) : Boolean; override;
    {Clause where sur les banques et les comptes}
    function  GetWhereBqCpt : string; override;
  public
    function  CommandeDetail(CurRow : HTStrings; var TT : TOB) : Boolean; override;
  end ;

function TRLanceFiche_DetailSuivi(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string): String;

Implementation

uses
  ExtCtrls, Commun, TRSAISIEFLUX_TOF, UProcGen, Constantes, UProcSolde;

{---------------------------------------------------------------------------------------}
function TRLanceFiche_DetailSuivi(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string): String;
{---------------------------------------------------------------------------------------}
begin
  Result := FICTAB.FicLance(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILSUIVI.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
var
  rbSolde   : TRadioButton;
  Mode, Dev,
  Cpt       : string;
begin
  TypeDesc := dsDetail;
  inherited ;
  LargeurCol := 100;

  fs_Date := ReadTokenSt(S);
  fs_Jour := StrToInt(ReadTokenSt(S));

  InArgument := True;
  TRadioButton(GetControl('DATEVAL')).OnClick := CritereClick;
  cbGeneral := THValComboBox(GetControl('GENERAL'));
  cbGeneral.OnChange := GeneralOnChange;
  cbGeneral.Plus := FiltreBanqueCp(cbGeneral.DataType, '', '');

  cbDevise.OnChange := GeneralOnChange;
  cbCompte := THMultiValComboBox(GetControl('MCCOMPTE'));
  cbCompte.Plus := FiltreBanqueCp(cbCompte.DataType, '', '');
  cbBanque.OnChange := GeneralOnChange;

  {On récupère les valeurs événtuellement passées en argument, à savoir :
   Choix d'affichage ; Societe ; Devise ; Compte ; Nature}
  Mode := ReadTokenSt(S);{Flux ou Solde demandé (si filtre par défaut -> écrasé !)}
  Societe := ReadTokenSt(S);

  rbFlux  := TRadioButton(GetControl('FLUX')); {Par flux ou solde}
  rbSolde := TRadioButton(GetControl('SOLDE'));
  rbFlux .OnClick := TypeSoldeClick;
  rbSolde.OnClick := TypeSoldeClick;

  {Si la devise est vide, on se met en Euro}
  Dev := ReadTokenSt(S);
  if Dev = '' then Dev := V_PGI.DevisePivot;

  {Définition de la combo devise}
  cbDevise.DataType := 'TTDEVISETOUTES';
//  G := 'D_DEVISE IN (SELECT TE_DEVISE FROM TRECRITURE';
  //cbDevise.Plus := G + ')';

  {Initialisation de la devise}
  Devise := Dev;
  MajComboDevise(Dev);

  {Gestion de la combo des comptes}
  Cpt := ReadTokenSt(S);

  cbGeneral.DataType := tcp_Tous;

  {Si le compte est vide, on prend le premier par défaut}
  if Cpt = '' then begin
    cbGeneral.ItemIndex := 0;
    Cpt := GetControlText('GENERAL');
  end
  else
    SetControlText('GENERAL', Cpt);

  {On récupère le type de flux}
  Nature := S;

  InArgument := False;
  if Mode = 'F' then rbFlux.Checked  := True
                else rbSolde.Checked := True;

  {Pour une raison que j'ignore, le faire plutôt provoque une violation d'accès en CWAS}
  DateDepart := StrToDate(fs_Date);
  NbColonne := fs_Jour;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILSUIVI.GeneralOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  DevISO : string;
begin
  if UpperCase(TComponent(Sender).Name) = 'GENERAL' then begin
    {Si on est dans un affichage par Flux, on récupère la devise du compte}
    DevISO := GetControlText('GENERAL');
    Devise := RetDeviseCompte(DevISO);
    AssignDrapeau(TImage(GetControl('IDEV1')), Devise);
    SetControlCaption('DEV', Devise);
  end
  else if UpperCase(TComponent(Sender).Name) = 'BANQUE' then begin
    DevIso := GetClauseIn(cbBanque.Value);
    {Si on est dans un affichage par solde, on filtre les comptes en fonction des banques choisies}
    if Trim(DevIso) <> '' then
      cbCompte.Plus := 'BQ_BANQUE IN(' + DevIso + ')';
    {Par défaut on récupère la devise du premier compte trouvé}
    if cbCompte.Items.Count > 0 then begin
      Devise := RetDeviseCompte(cbCompte.Values[0]);
      {Mise à jour de la combo des devises}
      MajComboDevise(Devise);
    end;
  end
  else
    Devise := cbDevise.Value;

  if Devise = '' then Devise := V_PGI.DevisePivot;
  AssignDrapeau(TImage(GetControl('IDEV')), Devise);

  CritereClick(edDepart);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILSUIVI.TypeSoldeClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  SoldeSimple : Boolean;
begin
  SoldeSimple := not rbFlux.Checked and not IsTresoMultiSoc;
  SetControlVisible('MCCOMPTE'    , SoldeSimple); //not rbFlux.Checked
  SetControlVisible('IDEV'        , SoldeSimple);
  SetControlVisible('CBDEVISE'    , SoldeSimple);
  SetControlVisible('SELECTIONDEV', SoldeSimple);
  SetControlVisible('SELECTION1'  , SoldeSimple);
  SetControlVisible('DEV'         , not SoldeSimple);// rbFlux.Checked
  SetControlVisible('IDEV1'       , not SoldeSimple);
  if rbFlux.Checked then begin
    TypeDesc := dsDetail;
    SetControlCaption('SELECTION', 'Compte');
    {20/11/03 : On rend le multival invisible pour mettre à la place le combo
    S := 'SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_BANQUE = "' + cbGeneral.Value + '"';
    cbGeneral.DataType := 'TRBANQUECP';}
    cbGeneral.Visible := True;
    cbBanque .Visible := False;
    cbGeneral.OnChange(cbGeneral);
  end
  else begin
    TypeDesc := dsSolde;
    if IsTresoMultiSoc then begin
      SetControlCaption('SELECTION', 'Compte');
      cbGeneral.Visible := True;
      cbBanque .Visible := False;
      cbGeneral.OnChange(cbGeneral);
    end
    else begin
      {20/11/03 : On rend le combo invisible pour mettre à la place le multival
      cbGeneral.DataType := 'TTBANQUE';//'TTETABBQ';}
      cbGeneral.Visible := False;
      cbBanque .Visible := True;
      SetControlCaption('SELECTION', 'Banque'); //BQ_ETABBQ
      cbBanque.OnChange(cbBanque);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_DETAILSUIVI.IsCelulleReinit(Col, Row : Integer) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited IsCelulleReinit(Col, Row);
end;

{---------------------------------------------------------------------------------------}
function TOF_DETAILSUIVI.CommandeDetail(CurRow : HTStrings; var TT : TOB) : Boolean;
{---------------------------------------------------------------------------------------}
var
  Lib : string;
begin
  Result := inherited CommandeDetail(CurRow, TT);
  {Pas de montant}
  if CurRow[PtClick.X] = '' then Exit;

  TT.SetString('DATEOPE', sDateOpe);
  TT.SetDateTime('DATEDEB', ColToDate(PtClick.X - COL_DATEDE));
  TT.SetDateTime('DATEFIN', ColToDate(PtClick.X - COL_DATEDE, True));
  TT.SetInteger('PERIODICITE', Ord(Periodicite));

  Lib := 'Flux ' + CurRow[COL_LIBELLE] + ' ';
  if rbDateOpe.Checked then Lib := Lib + 'en date d''opération'
                       else Lib := Lib + 'en date de valeur';
  TT.SetString('LIBELLE', Lib);

  {Code flux}
  if rbFlux.Checked then begin
    TT.SetString('GENERAL', cbGeneral.Value);
    TT.SetInteger('APPEL', Ord(dsDetail));
    TT.SetString('NATURE', Nature);
    TT.SetString('CODE', CurRow[COL_CODE]);
  end
  {Code banque}
  else begin
    {Si on est sur une ligne de totalisation, on sort}
    if (CurRow[COL_TYPE] = '') or (CurRow[COL_TYPE] = '+') or (CurRow[COL_TYPE] = '-') or
       (CurRow[COL_TYPE] = '*') or (CurRow[COL_NATURE] = na_Total) then Exit;

    TT.SetString('CODE', CurRow[COL_TYPE]);
    TT.SetString('GENERAL', CurRow[COL_CODE]);
    if CurRow[COL_CODE] = '' then Exit; {Sur banque}
    TT.SetString('LIBNAT', RechDom('TRNATURE', CurRow[COL_NATURE], False));
    TT.SetString('NATURE', 'AND TE_NATURE = "' + CurRow[COL_NATURE] + '"');
    TT.SetInteger('APPEL', Ord(TypeDesc));
  end;

  Result := True;
end;

{JP 27/10/03 : Lancement de l'écran de saisie des écritures de simulation
{---------------------------------------------------------------------------------------}
procedure TOF_DETAILSUIVI.PopupSimulation(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  General : string;
begin
  inherited;
  if rbFlux.Checked then
    General := cbGeneral.Value
  {Code banque}
  else
    General := Grid.Cells[COL_CODE, Grid.Row];
  {08/03/05 : FQ 10217 : se positionner sur la date d'opération, si on est en date d'opération}
  TRLanceFiche_SaisieFlux('TR', 'TRSAISIEFLUX', '', '', General + ';;' + DateCourante + ';');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 13/02/2002
Modifié le ... :   /  /
Description .. : Constitue la sélection selon les critères du panneau et ceux
Suite ........ : passés en paramètre pour remplir une Tob qui sert a
Suite ........ : l'affichage de la grille
Mots clefs ... : GRID
*****************************************************************}
function TOF_DETAILSUIVI.ParFlux : Integer;
var
  TobL,
  TobG,
  TobT     : TOB;
  Sens     : string;
  TypeFlux : string;
  {LÉGENDE DES TYPES : "TypeFlux" = Ligne contenant les opérations
                       ""          = Ligne contenant les libellés
                       "*"         = Ligne de séparation entre recettes et dépense et solde total
                       "+"         = total recette
                       "-"         = Total dépense}

    {-------------------------------------------------------------------------}
    procedure CreeTobDetail(TypeFlux, CodeFlux, LibelleFlux, Nat : string);
    {-------------------------------------------------------------------------}
    var
      N : Integer;
    begin
      TobG := TOB.Create('', TobGrid, -1);
      TobG.AddChampSupValeur('TYPEFLUX', TypeFlux);
      TobG.AddChampSupValeur('CODEFLUX', CodeFlux);
      TobG.AddChampSupValeur('NATURE'  , Nat);
      TobG.AddChampSupValeur('LIBELLE', LibelleFlux);
      {Ajout des colonnes dates avec montant}
      for N := 0 to NbColonne - 1 do
        TobG.AddChampSupValeur(RetourneCol(n), '');
    end;

    {-------------------------------------------------------------------------}
    procedure CreeTobTotal;
    {-------------------------------------------------------------------------}
    var
      N   : Integer;
      S   : string;
      Mnt : Double;
    begin
      if Sens = 'C' then S := '+'
                    else S := '-';
      CreeTobDetail(S, '', TraduireMemoire('Total'), '');
      {Cumule chaque colonne date pour le type flux courant}
      for n := 0 to NbColonne - 1 do begin
        S := RetourneCol(n);
        Mnt := TobGrid.Somme(S, ['TYPEFLUX'], [TypeFlux], False);
        TobG.SetDouble(S, Mnt);
      end;
    end;

var
  S, CodeFlux : string;
  I           : Integer;
  Montant     : Double;
  MntTmp      : Double;
  GererMilles : Boolean;
  CreerOk     : Boolean; {JP 03/03/04}
  DateTmp     : TDateTime;
  Mnt31_12    : Double;
  Mnt01_01    : Double;
  MntReinit   : Double;
begin
  CreeTobDetail('*', '', 'Solde initial', '');

  ColRouge := -1;

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

  for I := 0 to TobListe.Detail.Count - 1 do begin
    TobL := TobListe.Detail[I];

    S := TobL.GetValue('TFT_TYPEFLUX');
    {Changement de type flux}
    if S <> TypeFlux then begin
      {S'il on n'est pas sur sur le premier enregistrement et si la précédente ligne n'était pas
       une ligne de réinitialisation, on crée une ligne "total", pour le précédent type de flux}
      if (TypeFlux <> '') and (CodeFlux <> CODEREGULARIS) then
        CreeTobTotal;

      {Changement de sens}
      if Sens <> TobL.GetValue('TTL_SENS') then begin
        Sens := 'D'; // Pour total
        {Il y a quelquechose avant}
        if TypeFlux <> '' then begin
          {Création d'une tob fille vide pour séparer les sens}
          CreeTobDetail('*', '', '', '');
          Result := TobGrid.Detail.Count; {Position pour mise en forme (=TobG.GetIndex+1)}
        end;
      end;

      CodeFlux := TobL.GetString('TE_CODEFLUX');
      {JP 03/03/04 : Les écriture de réinitialisation son gérées à la fin}
      if CodeFlux = CODEREGULARIS then Continue;

      TypeFlux := S;
      if TobL.GetValue('TTL_LIBELLE') = CODETEMPO then begin
        if Sens = 'C' then CreeTobDetail('', '', 'Recette comptable', '')
                      else CreeTobDetail('', '', 'Dépense comptable', '');
      end
      else
        CreeTobDetail('', '', TobL.GetValue('TTL_LIBELLE'), '');
    end
    else
      CodeFlux := TobL.GetString('TE_CODEFLUX');

    TobG := TobGrid.FindFirst(['CODEFLUX'], [CodeFlux], True);
    {Il sagit d'un nouveau flux, on crée une tob fille de la TobGrid}
    if (TobG = nil) then
      {21/04/04 : Le R sert à signifier qu'il s'agit d'une ligne de détail, pour le KeyDown}
      CreeTobDetail(TypeFlux, CodeFlux, TobL.GetValue('TFT_LIBELLE'), 'R');

    {JP 20/12/05 : FQ 10289 : pour gérer les saisies négatives dans la Trésorerie, il ne faut pas travailler
                   sur les valeurs absolues mais en fonctions du sens du flux => si les montants sont négatifs
                   c'est qu'il y a eu des saisies en Tréso négative ; si le problème se produit avec
    Montant := Abs(TobL.GetValue('TE_MONTANTDEV'));}
    if Sens = 'C' then Montant := Valeur(TobL.GetString('TE_MONTANTDEV'))
                  else Montant := Valeur(TobL.GetString('TE_MONTANTDEV')) * -1;

    S := GetTitreFromDate(TobL.GetDateTime(sDateOpe));
    {Dans la fourchette d'affichage}
    if (TobL.GetDateTime(sDateOpe) >= DateDepart) and (TobL.GetDateTime(sDateOpe) <= DateFin) then begin
      if TobG.GetString(S) <> '' then
        Montant := Montant + TobG.GetDouble(S);
      {06/12/05 : FQ 10317 : Cumul éventuel des opérations comprises en le 01/01 et la fin de la
                  première semaine de l'année pour le suivi hébdomadaire. C'est le seul cas où le
                  01/01 peut être "au milieu de la période d'une colonne"}
      SetlOpeDebAnnee('DetailFlux', TobL.GetDouble('TE_MONTANTDEV'), TobL.GetDateTime(sDateOpe));
      TobG.PutValue(S, Montant);
    end;
  end;{Boucle For}

  {********************************************************************}
  {******************  GESTION DES RÉINITIALISATIONS ******************}
  {********************************************************************}

  GererMilles := False;
  {Gestion de l'éventuelle présence d'une écriture de réinitialisation}
  TobL := TobListe.FindFirst(['TE_CODEFLUX'], [CODEREGULARIS], True);
  if TobL <> nil then begin
    DateTmp := TobL.GetDateTime(sDateOpe);
    {2 Possibilités : - le premier jour est le 01/01, dans ce cas il n'y a rien à faire car GetSoldeInit va renvoyer
                        comme solde initial le Montant forcé - les opérations du jour et de la nature demandée =>
                        la première colonne sera   01/01
                                                 Montant forcé - les opérations du jour
                                                 les opérations du jour
                                                 Montant Forcé

                      - Si le 01/01 se situe en milieu :     31/12         |    01/01
                                                         Solde reporté 30  | Solde reporté 31
                                                         Opérations du J   | Opérations du J
                                                         Solde Calculé 31  | Solde Forcé
                        Il faut donc rajouter le 01/01 une ligne de compensation dont le montant est
                           M := Solde Forcé - Opérations du J - Solde reporté 31}
    if (DateDepart <> DebutAnnee(DateDepart)) and
       (DateTmp > DateDepart) and (DateTmp < DateFin) then begin
      {On se contente de créer les tobs, le calcul des montants sera fait en même temps que le calcul total
       08/06/04 : Ajout du test sur le TYPEFLUX pour ne pas créer une tob total si la dernière en est une}
      if CreerOk and not (StrToChr(TobG.GetString('TYPEFLUX')) in ['+', '-']) then CreeTobTotal;
      TypeFlux := TobL.GetValue('TFT_TYPEFLUX');
      CreeTobDetail('', '', 'Réinitialisation', '');
      CreeTobDetail(TypeFlux, CODEREGULARIS, TobL.GetValue('TFT_LIBELLE'), '');
      if CodeFlux = CODEREGULARIS then CreeTobTotal;/////
      GererMilles := True;
    end;
  end; {if TobL <> nil }

  {Création de la tob fille qui contiendra le total du dernier flux traiter}
  if CodeFlux <> CODEREGULARIS then CreeTobTotal;//////
  TobT := TobG;

  {********************************************************************}
  {******************  GESTION DES SOLDES TOTAUX **********************}
  {********************************************************************}

  {Création de la dernière tob fille qui contiendra le solde total}
  CreeTobDetail('*', '', TraduireMemoire('Solde final'), '');

  if (Nature = '') then begin
    if rbDateOpe.Checked then
      Montant := GetSolde(cbGeneral.Value, DateToStr(DateDepart - 1), S)
    else
      Montant := GetSoldeValeur(cbGeneral.Value, DateToStr(DateDepart - 1), S);
  end
  else{À la différence de GetSolde(Valeur), GetSoldeInit exécute une requête avec une inégalité stricte => pas de -1}
    Montant := GetSoldeInit(cbGeneral.Value, DateToStr(DateDepart - 1), Nature, not rbDateOpe.Checked);

  {Calcul des soldes des dates de la fourchette de sélection}
  for i := 0 to NbColonne - 1 do begin
    S := RetourneCol(i);

    {Si le début de millésime figure dans la fourchette de date}
    if IsColAvecSoldeInit(i) then begin
      {Pour afficher le solde en rouge vif !!}
      ColRouge := i + COL_DATEDE;
      {Calcul du montant de réinitialisation}
      if GererMilles then begin
        {Affectation du report du solde de la veille}
        TobGrid.Detail[0].SetDouble(S, Montant);

        {JP 05/12/05 : FQ 10317 : Calcul de la date de début d'année}
        if Periodicite = tp_7 then DateTmp := DebutAnnee(ColToDate(i, True))
                              else DateTmp := DebutAnnee(ColToDate(i));

        {JP 05/12/05 : FQ 10317 : Récupération du solde forcé}
        if Periodicite = tp_1 then
          Mnt01_01 := GetSoldeMillesime(cbGeneral.Value, DateToStr(DateTmp), Nature, not rbDateOpe.Checked, False)
        else
          Mnt01_01 := GetSoldeMillesime(cbGeneral.Value, DateToStr(DateTmp), Nature, not rbDateOpe.Checked, True);

        {Récupération la tob Reinitialisation}
        TobL := TobGrid.FindFirst(['TYPEFLUX'], [CODEREGULARIS], True);
        {JP 05/12/05 : FQ 10317 : on affecte directement le solde forcé en affichage quotidien}
        if Periodicite = tp_1 then begin
          {Récupération des opérations du 01/01}
          MntTmp := Mnt01_01 - TobGrid.Somme(S, ['TYPEFLUX'], ['+'], False) - TobGrid.Somme(S, ['TYPEFLUX'], ['-'], False);

          {Récupération du solde du 31}
          MntReinit := - MntTmp + Montant;
          {Le solde à reporter pour la suite de la boucle est le solde de réinitialisation}
          Montant := Mnt01_01;
        end
        //////////////
        else begin
          if Periodicite = tp_7 then begin
            {Récupération des opérations de la colonne du 01/01 strictement postérieures au 01/01}
            MntTmp := GetlOpeDebAnnee('DetailFlux');
            {Calcul du solde "théorique" en fin de période s'il n'y avait pas de réinitialisation}
            Mnt31_12 := Montant + TobGrid.Somme(S, ['TYPEFLUX'], ['+'], False) - TobGrid.Somme(S, ['TYPEFLUX'], ['-'], False);
          end
          else begin
            {Somme des opérations du mois de janvier}
            MntTmp := TobGrid.Somme(S, ['TYPEFLUX'], ['+'], False) - TobGrid.Somme(S, ['TYPEFLUX'], ['-'], False);
            {Calcul du solde "théorique" en fin de période s'il n'y avait pas de réinitialisation}
            Mnt31_12 := Montant + MntTmp;
          end;

          {Calcul du solde "constaté" en fin de période avec la réinitialisation}
          Montant  := Mnt01_01 + MntTmp;
          {Calcul de la compensation du à la réinitialisation}
          MntReinit := Montant - Mnt31_12;
        end;

        {Affectation de la compensation due à la réinitilisation}
        TobL.SetDouble(S, MntReinit * -1);
        {Affectation de la tob Total Reinitialisation}
        TobT.SetDouble(S, TobL.GetDouble(S));
        {Affectation du solde à la fin de la période de réinitilisation}
        TobG.SetDouble(S, Montant);
        Continue;
      end;
    end;

    {Solde de la veille, affecté à la première tob fille}
    TobGrid.Detail[0].SetDouble(S, Montant);
    {Calcul du solde à la fin du jour}
    Montant := Montant + TobGrid.Somme(S, ['TYPEFLUX'], ['+'], False) - TobGrid.Somme(S, ['TYPEFLUX'], ['-'], False);
    {On affecte le solde à la dernière tob fille}
    TobG.SetDouble(S, Montant);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILSUIVI.OnUpdate ;
{---------------------------------------------------------------------------------------}
var
  Q      : TQuery;
  SQL    : string;
  I, J   : Integer;
  NSepar : Integer;
  sCpt,
  Tmp1   : string;
  TmpREI : string;
  wBanque: string;
  wDate  : string;
  wConf  : string;
begin
  inherited ;

  wDate  := sDateOpe + ' >= "' + UsDateTime(DateDepart) + '" AND ' + sDateOpe + ' <= "' + UsDateTime(DateFin) + '" ';

  if rbFlux.Checked then begin
    {A DEPLACER DANS L'ANCETRE ??!!}
    {S'il y a un filtre sur les natures, il ne faut pas qu'il s'applique sur les écritures de réinitialisation}
    if Trim(Nature) <> '' then begin
      Tmp1 := Copy(Nature, Pos('AND', Nature) + 4, Length(Nature));
      TmpREI := ' AND ((' + Tmp1 + ') OR TE_QUALIFORIGINE = "' + CODEREINIT + '") '
    end
    else
      TmpREI := Nature;

    wConf := GetWhereConfidentialite;

    SQL := 'SELECT TRECRITURE.'+sDateOpe+',TRECRITURE.TE_CODEFLUX,TRECRITURE.TE_MONTANTDEV,' +
       'FLUXTRESO.TFT_FLUX,FLUXTRESO.TFT_TYPEFLUX,FLUXTRESO.TFT_LIBELLE, ' +
       'TYPEFLUX.TTL_SENS,TYPEFLUX.TTL_LIBELLE ' +
       'FROM TRECRITURE, FLUXTRESO, TYPEFLUX, BANQUECP ' +
       'WHERE TE_CODEFLUX = TFT_FLUX AND TFT_TYPEFLUX = TTL_TYPEFLUX AND BQ_CODE = TE_GENERAL ' +
       'AND TE_GENERAL="' + cbGeneral.Value + '" '+ TmpREI + ' AND ' + wDate + wConf + 
       ' UNION ALL ' + {27/01/05 : FQ 10198 "ALL" pour bien prendre tous les enregistrements}
       'SELECT TRECRITURE.' + sDateOpe + ', TRECRITURE.TE_CODEFLUX, TRECRITURE.TE_MONTANTDEV,' +
       'RUBRIQUE.RB_RUBRIQUE TFT_FLUX, RUBRIQUE.RB_TYPERUB TFT_TYPEFLUX, RUBRIQUE.RB_LIBELLE TFT_LIBELLE,' +
       'RUBRIQUE.RB_SIGNERUB TTL_SENS, "' + CODETEMPO + '" TTL_LIBELLE ' +
       'FROM TRECRITURE, RUBRIQUE, BANQUECP ' +
       'WHERE TE_CODEFLUX = RB_RUBRIQUE AND BQ_CODE = TE_GENERAL  ' +
       'AND TE_GENERAL="' + cbGeneral.Value + '" '+ TmpREI + ' AND ' + wDate + wConf +
       ' ORDER BY TTL_SENS, TFT_TYPEFLUX, TE_CODEFLUX'; {11/04/05 : FQ 10238 : on trie sur les codes flux}
  end
  else begin
    {12/05/04 : Éventuelle création de la liste des comptes
     24/01/06 : Fait dans le OnLoad suite à la FQ 10328
    CreerListeCompte;}
    if not IsTresoMultiSoc then begin
      {Constitution du filtre sur les comptes}
      if cbCompte.Value <> '' then begin
        sCpt := ' AND TE_GENERAL IN (';
        Tmp1 := GetClauseIn(cbCompte.Value);
        sCpt := sCpt + Tmp1 + ') ';
      end;

      {20/11/03 : Gestion de la nouvelle clause where sur les banques}
      if cbBanque.Value <> '' then begin
        wBanque := 'AND BQ_BANQUE IN (';
        Tmp1 := GetClauseIn(cbBanque.Value);
        wBanque := wBanque + Tmp1 + ') AND PQ_BANQUE IN (' + Tmp1 + ') ';
      end
      else
        wBanque := '';
      Tmp1 := wBanque + sCpt;
    end
    else
      Tmp1 := ' AND TE_GENERAL = "' + cbGeneral.Value + '" ';

    Tmp1 := ' WHERE ' + wDate + Tmp1;
    SQL := GetRequete(Tmp1);
  end;
  Q := OpenSQL(SQL, True);
  TobListe.LoadDetailDB('1', '', '', Q, False, True);
  Ferme(Q);

  NSepar := 0;
  if rbFlux.Checked then NSepar := ParFlux
                    else ParSolde;

  Grid.ColCount := COL_DATEDE + NbColonne;

  {Formatage des montant : '#,##0.00'; Euros}
  SQL := StrFMask(CalcDecimaleDevise(Devise), '', True);
  {Formatage des dates}
  for I := COL_DATEDE to COL_DATEDE + NbColonne - 1 do
    Grid.ColFormats[I] := SQL;

  {Affichage dans la grille}
  TobGrid.PutGridDetail(Grid, True, True, '', True);

  Grid.Cells[COL_LIBELLE, 0] := '';

  {Gestion du séparateur de ligne par flux : entre dépenses et recettes}
  if NSepar > 0 then
    Grid.RowHeights[NSepar] := 3;
  {Gestion du séparateur de ligne par solde entre chaque banque}
  if not rbFlux.Checked then
    for i := 0 to Grid.RowCount - 1 do
      if Grid.Cells[COL_TYPE, i] = '*' then Grid.RowHeights[i] := 3;

  SQL := 'Détail de la fiche de suivi';
  if not MultiNature then begin
         if Pos('"' + na_Realise    + '"', Nature) > 0 then LibNature := ' réalisés'
    else if Pos('"' + na_Simulation + '"', Nature) > 0 then LibNature := ' de simulation'
                                                       else LibNature := ' prévisionnels';
    SQL := SQL + ' (flux' + LibNature + ')';
  end
  else
    LibNature := '';

  Ecran.Caption := TraduireMemoire(SQL);
  UpdateCaption(Ecran);

  if rbFlux.Checked then
    {Recherche du premier flux pour positionnement}
    for I := COL_DATEDE to COL_DATEDE + NbColonne - 1 do
      for J := 2 to Grid.RowCount - 3 do
        {Lignes sauf soldes (''), total (+, -)}
        if (Grid.Cells[I, J] <> '') and
           (Grid.Cells[COL_TYPE, J] <> '+') and
           (Grid.Cells[COL_TYPE, J] <> '-') then
        begin
          Grid.Col := I;
          Grid.Row := J;
          Grid.SetFocus;
          Exit;
        end;

  Grid.Col := COL_DATEDE;
  Grid.Row := Grid.FixedRows;
  Grid.SetFocus;
end ;

{---------------------------------------------------------------------------------------}
function TOF_DETAILSUIVI.GetWhereBqCpt : string;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited GetWhereBqCpt;
  if not IsTresoMultiSoc then begin
    if cbCompte.Value <> '' then begin
      Result := 'BQ_CODE IN (';
      Result := Result + GetClauseIn(cbCompte.Value);
      Result := Result + ')';
    end;

    if cbBanque.Value <> '' then begin
      if Result <> '' then Result := 'AND BQ_BANQUE IN ('
                      else Result := 'BQ_BANQUE IN (';
      Result := Result + GetClauseIn(cbBanque.Value);
      Result := Result + ')';
    end;
  end
  else
    Result := 'BQ_CODE = "' + cbGeneral.Value + '"';
end;

initialization
  RegisterClasses([TOF_DETAILSUIVI]);

end.

