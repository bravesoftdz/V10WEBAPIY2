{Source de la tof TRPREVMENSUEL
--------------------------------------------------------------------------------------
  Version    |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
   0.91        06/11/03  JP   création de l'unité
 1.50.000.000  22/04/04  JP   Nouvelle gestion des natures (OnUpdate, ModifierLibelle)
 6.50.001.019  20/09/05  JP   FQ 10289 : Mauvaise gestion du signe des montants.
 6.60.001.001  20/12/05  JP   FQ 10289 : On revient en arrière et on modifie les autres suivis
                              par flux et non l'inverse (modification du 20/09/05 pour uniformiser
                              avec le suivi périodique)
 7.09.001.001  25/08/06  JP   1/ Gestion du Multi sociétés
                              2/ Gestion du comparatif budgétaire
 8.01.001.009  28/03/07  JP   FQ 10423 : modifications des libellés
 8.10.001.004  08/08/07  JP   Gestion des confidentialités : Modification de la requête et ajout de BANQUECP
 8.10.004.001  08/11/07  JP   FQ 10534 : ajout du test sur la classe Trésorerie
--------------------------------------------------------------------------------------}

unit TRPREVMENSUEL_TOF;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  StdCtrls, Controls, Classes, UFicTableau,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  SysUtils, Spin, HCtrls, HEnt1, UTOB, Grids, UProcGen;

type
  TOF_TRPREVMENSUEL = class (FICTAB)
    procedure OnUpdate              ; override;
    procedure OnArgument(S : string); override;
  protected
    sTypeMnt  : string;
    RecCptOk  : Boolean;
    DepCptOk  : Boolean;

    cbGeneral  : THValComboBox;
    cbNature   : THMultiValComboBox;

    TableRouge : array [0..30] of Char;

    {Méthodes}
    function  Afficher      (NbCol: Integer; DateDe: TDateTime) : Integer;
    function  FourchetteDate(dtDeb : string; DebOk : Boolean; NbCol : Integer) : TDateTime;
    {Evènements}
    procedure GeneralOnChange(Sender : TObject);
    procedure PopupSimulation(Sender : TObject); override;
    procedure bGraphClick    (Sender : TObject); override;
    function  IsCelulleReinit(Col, Row : Integer) : Boolean; override;
  public
    function  IsNbColOk   : Boolean; override;
    function  CommandeDetail(CurRow : HTStrings; var TT : TOB) : Boolean; override;
  end ;

function TRLanceFiche_TRPREVMENSUEL(Dom, Fiche, Range, Lequel, Arguments : string) : string;

implementation

uses
  ExtCtrls, AGLInit, Commun, TRGRAPHMENSUEL_TOF, TRSAISIEFLUX_TOF, Constantes,
  UProcSolde, HMsgBox;

{---------------------------------------------------------------------------------------}
function TRLanceFiche_TRPREVMENSUEL(Dom, Fiche, Range, Lequel, Arguments : string) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := FICTAB.FicLance(Dom, Fiche, Range, Lequel, Arguments);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRPREVMENSUEL.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  TypeDesc := dsMensuel;
  inherited;
  LargeurCol := 120;

  InArgument := True;
  TRadioButton(GetControl('DATEVAL')).OnClick := CritereClick;
  cbGeneral := THValComboBox(GetControl('GENERAL'));
  cbGeneral.OnChange := GeneralOnChange;

  {Initialisation de la devise}
  Devise := V_PGI.DevisePivot;

  cbGeneral.DataType := tcp_Tous;
  cbGeneral.Plus := FiltreBanqueCp(tcp_Tous, '', '');

  Nature := S;

  InArgument := False;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRPREVMENSUEL.GeneralOnChange(Sender: TObject);
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
function TOF_TRPREVMENSUEL.IsNbColOk : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited IsNbColOk;
  if Result then begin
    Result := (seInterval.Value <= 24) or
              (HShowMessage('0;' + Ecran.Caption + ';Vous ne pouvez pas traiter plus de 24 mois.'#13 +
                   'Voulez-vous poursuivre sur les 24 premieres mois ?;Q;YNC;N;C;', '', '') = mrYes);
    if Result and (seInterval.Value > 24) then
      seInterval.Value := 24;
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_TRPREVMENSUEL.CommandeDetail(CurRow : HTStrings; var TT : TOB) : Boolean;
{---------------------------------------------------------------------------------------}
var
  Deb : string;
  Fin : string;
  Lib : string;
begin
  Result := inherited CommandeDetail(CurRow, TT);

  {Code flux}
  if CurRow[PtClick.X] = '' then Exit; {Pas de montant}

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
  Deb := DateToStr(DebutDeMois(PlusMois(StrToDate(GetControlText('DATEDE')), PtClick.X - COL_DATEDE)));
  Fin := DateToStr(FinDeMois(StrToDate(Deb)));

  TT.SetDateTime('DATEDEB', StrToDate(Deb));
  TT.SetDateTime('DATEFIN', StrToDate(Fin));
  TT.SetString('REGROUP', GetControlText('REGROUPEMENT'));
  Result := True;
end;

{Lancement de l'écran de saisie des écritures de simulation
{---------------------------------------------------------------------------------------}
procedure TOF_TRPREVMENSUEL.PopupSimulation(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  General : string;
  DateM   : string;
begin
  inherited;
  GetPosition;
  {On est sur une colonne Date}
  if (PtClick.X >= COL_DATEDE) and (PtClick.Y > 0) then begin
    General := cbGeneral.Value; // ?
  end;

  DateM := DateToStr(MoisToDate(Grid.Cells[PtClick.X, 0]));
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
  TRLanceFiche_SaisieFlux('TR', 'TRSAISIEFLUX', '', '', General + ';MENS;' + DateM + ';');
  MouetteClick;
end;

{Constitution de la tob contenant les données avant de la mettre dans la grille
{---------------------------------------------------------------------------------------}
function TOF_TRPREVMENSUEL.Afficher(NbCol: Integer; DateDe: TDateTime): Integer;
{---------------------------------------------------------------------------------------}
var
  TobL,
  TobG,
  TobD,
  TobT     : TOB;
  Sens     : string;
  TypeFlux : string;
  Mois     : string;
  DateDeb  : TDateTime;
  DateFin  : TDateTime;

    {-------------------------------------------------------------------------}
    procedure CreeTobDetail(TypeFlux, CodeFlux, LibelleFlux, Nat : string);
    {-------------------------------------------------------------------------}
    var
      n : Integer;
    begin
      TobG := TOB.Create('', TobGrid, -1);
      TobG.AddChampSupValeur('TYPEFLUX', TypeFlux);
      TobG.AddChampSupValeur('CODEFLUX', CodeFlux);
      TobG.AddChampSupValeur('NATURE'  , Nat);
      TobG.AddChampSupValeur('LIBELLE' , LibelleFlux);
      {Ajout des colonnes dates avec montant}
      for n := 0 to NbCol - 1 do
        TobG.AddChampSupValeur(RetourneMois(DateToStr(DateDe), n ), '');
    end;

    {-------------------------------------------------------------------------}
    procedure CreeTobTotal(Demi : Boolean = False);
    {-------------------------------------------------------------------------}
    var
      N   : Integer;
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
      for N := 0 to NbCol - 1 do begin
        S := DateToStr(PlusMois(DateDe, N));
        Mois := DateToMois(S);
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

var
  S, CodeFlux : string;
  I           : Integer;
  Montant     : Double;
  MntTmp      : Double;
  GererMilles : Boolean;
  CreerOk     : Boolean; {JP 03/03/04}
  OldNature   : string;
begin
  {Par défaut on considère qu'il n'y a pas de d'imports comptables dans le résultat de la requête}
  RecCptOk := False;
  DepCptOk := False;
  CreerOk  := False;

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

  {Création de la Tob contenant les soldes en début de mois}
  CreeTobDetail('*', '', 'Solde initial', '');

  TypeFlux := '';
  Sens := 'C';
  Result := 0;

  DateDeb := FourchetteDate(DateToStr(DateDe), True, NbCol);
  DateFin := FourchetteDate(DateToStr(DateDe), False, NbCol);

  for I := 0 to 30 do TableRouge[I] := '-';

  {$IFDEF TOB_DEBUG}
  TobDebug(TobListe);
  {$ENDIF}

  {********************************************************************}
  {******************  REMPLISSAGE DE LA TOB GRILLE  ******************}
  {********************************************************************}

  for I := 0 to TobListe.Detail.Count - 1 do begin
    TobL := TobListe.Detail[I];

    S := TobL.GetValue('TFT_TYPEFLUX');
    {Changement de type flux}
    if S <> TypeFlux then begin

      {Changement de sens}
      if Sens <> TobL.GetValue('TTL_SENS') then begin
        {Il y a quelquechose avant et soit on a des dépenses autres que la réinitialisation
                                      soit le premier janvier n'est pas le premier jour de
                                           traitement
                                      soit le premier janvier est le premier jour mais on
                                           travaille sur plus de douze mois}
        if (TypeFlux <> '') and (CreerOk or (DateDeb <> DebutAnnee(DateDeb)) or (seInterval.Value > 12)) then begin
          CreeTobTotal;
          {Création de la tob total des recettes}
          CreeTobTotal(True);
          Sens := 'D';
          {Création d'une tob fille vide pour séparer les sens}
          CreeTobDetail('*', '', '', '');
          Result := TobGrid.Detail.Count; {Position pour mise en forme (=TobG.GetIndex+1)}
        end
        else if (CreerOk or (DateDeb <> DebutAnnee(DateDeb))) then
          Sens := 'D';
      end
      else if (TypeFlux <> '') then begin
        if CodeFlux <> CODEREGULARIS then CreeTobTotal;
      end;

      if TobL.GetValue('TE_CODEFLUX') = CODEREGULARIS then begin
        {08/06/04 : J'espère que c'est la dernière modification pour ces réinitialisations !}
        CodeFlux := TobL.GetValue('TE_CODEFLUX');
        Continue;
      end;

      TypeFlux := S;
      if TobL.GetValue('TTL_LIBELLE') = CODETEMPO then begin
        if Sens = 'C' then begin
          CreeTobDetail('', '', 'Encaissements comptables', '');
          RecCptOk := True;
        end
        else begin
          CreeTobDetail('', '', 'Décaissements comptables', '');
          DepCptOk := True;
        end;
      end
      else
        CreeTobDetail('', '', TobL.GetValue('TTL_LIBELLE'), '');
    end;

    CodeFlux := TobL.GetValue('TE_CODEFLUX');
    if ckGlobal.Checked then begin
      TobG := TobGrid.FindFirst(['CODEFLUX'], [CodeFlux], True);
      OldNature := '';
    end
    else begin
      OldNature := GetNatureToBudget(TobL.GetValue('TTL_LIBELLE'), TobL.GetValue('TE_NATURE'));
      TobG := TobGrid.FindFirst(['CODEFLUX', 'NATURE'], [CodeFlux, OldNature], True)
    end;

    {Il sagit d'un nouveau flux, on crée une tob fille de la TobGrid}
    if TobG = nil then
      {JP 22/04/04 : La nature ne sert que pour le KeyDown dans l'immédiat}
      CreeTobDetail(TypeFlux, CodeFlux, TobL.GetValue('TFT_LIBELLE'), OldNature);

    {Montant := Abs(TobL.GetValue(sTypeMnt));
     JP 29/04/04 : Cela donne des erreurs avec les imports comptables tant que l'on ne gérera pas
                   le sens dans les rubriques. Pour que les montants au debit soient de manière
                   générale positifs dans la TobGrid, on tient compte du solde lors du cumul (cf. ci-dessous)

     JP 15/09/05 : FQ 10289 :Je ne sais pas pourquoi j'ai mis le commentaire ci-dessus, toujours
                   est-il qu'aujourd'hui, il faut travailler avec les valeurs absolues
     JP 21/12/05 : FQ 10289 : Finalement, comme on accepte la saisie d'écriture négative en Tréso (demande de Thermo Compac)
                   il faut gérer le sens : c'est le suivi mensuel qui était correct, la fiche de suivi et
                   le suivi périodique devant être modifiés
     Montant := Abs(TobL.GetValue(sTypeMnt));}
     Montant := TobL.GetValue(sTypeMnt);

    S := TobL.GetValue(sDateOpe);
    {On récupère le titre de la colonne correspondant à la date de l'écriture}
    Mois := DateToMois(S);

    {Dans la fourchette d'affichage}
    if (StrToDate(S) >= DateDeb) and (StrToDate(S) < DateFin) then begin
      {JP 29/04/04 : En attendant une gestion du signe des rubriques, on cumule en tenant compte du sens.
                     Ainsi, si le signe d'une écriture ne correspond pas à son sens, il sera pris en compte

       JP 15/09/05 : FQ 10289 : Je ne sais pas pourquoi j'ai mis le commentaire ci-dessus, toujours
                     est-il qu'aujourd'hui, il faut travailler avec les valeurs absolues
       JP 20/12/05 : FQ 10289 : on revient en arrière (demande de Thermo Compac)
       Montant := Montant + Abs(Valeur(VarToStr(TobG.GetValue(Mois))));}
       if Sens = 'C' then Montant := Montant + Valeur(VarToStr(TobG.GetValue(Mois)))
                     else Montant := Valeur(VarToStr(TobG.GetValue(Mois))) - Montant;

      TobG.PutValue(Mois, Montant);
    end;
  end;{Boucle For}

  {********************************************************************}
  {******************  GESTION DES RÉINITIALISATIONS ******************}
  {********************************************************************}

  GererMilles := False;
  {Gestion de l'éventuelle présence d'une écriture de réinitialisation}
  TobL := TobListe.FindFirst(['TE_CODEFLUX'], [CODEREGULARIS], True);
  repeat
    if TobL <> nil then begin
      S := TobL.GetValue(sDateOpe);
      {2 Possibilités : - le premier jour est le 01/01, dans ce cas il n'y a rien à faire car GetSoldeInit va renvoyer
                          comme solde initial le Montant forcé - les opérations du jour et de la nature demandée =>
                          la première colonne sera   01/01
                                                   Montant forcé - les opérations du jour
                                                   les opérations du jour
                                                   Montant Forcé

                        - Si le 01/01 se situe en milieu :     12          |    01
                                                         Solde reporté 11  | Solde reporté 12
                                                         Opérations du M   | Opérations du M
                                                         Solde Calculé 12  | Solde Forcé + ope [02/01, 31/01]
                          Il faut donc rajouter sur 01 une ligne de compensation dont le montant est
                                 M := Solde Forcé - Opérations du 01/01 - Solde reporté 12}
      if (DateDeb <> DebutAnnee(DateDeb)) and
         (StrToDate(S) >= DateDeb)       and
         (StrToDate(S) <= DateFin) then begin
        {On se contente de créer les tobs, le calcul des montants sera fait en même temps que le calcul total}
        {JP 03/03/04 : pour éviter de créer une tob s'il n'y a aucun débit
           08/06/04 : Ajout du test sur le TYPEFLUX pour ne pas créer une tob total si la dernière en est une}
        if CreerOk and not (StrToChr(TobG.GetString('TYPEFLUX')) in ['+', '-']) then CreeTobTotal;
        Sens := 'D'; {Dans le cas où il n'y aurait que la ligne d'initilisation}
        TypeFlux := TobL.GetValue('TFT_TYPEFLUX');
        CreeTobDetail('', '', 'Réinitialisation', '');
        CreeTobDetail(TypeFlux, CODEREGULARIS, TobL.GetValue('TFT_LIBELLE'), '');
        GererMilles := True;
        if not CreerOk then CreeTobTotal; {09/11/06}
        Break; {JP 27/02/04 : Si plusieurs exercices, on ne crée qu'une tob}
      end;
    end;
    TobL := TobListe.FindNext(['TE_CODEFLUX'], [CODEREGULARIS], True);
  until TobL = nil;

  if CreerOk and not (StrToChr(TobG.GetString('TYPEFLUX')) in ['+', '-']) then
    CreeTobTotal;
  TobT := TobG;

  {********************************************************************}
  {******************  GESTION DES SOLDES TOTAUX **********************}
  {********************************************************************}

  {Création de la tob total pour le total des dépenses}
  CreeTobTotal(True);
  TobD := TobG;

  {Création de la dernière tob fille qui contiendra le solde total}
  CreeTobDetail('*', '', TraduireMemoire('Solde final'), '');

  {Récupération du solde à la veille du premier mois traiter}
  if cbGeneral.Value = '' then begin
    {09/11/06 : Si on travaille sur un dossier, sans filtre sur un compte}
    if GetControlText('REGROUPEMENT') <> '' then
      Montant := GetSoldeInitDossier(GetControlText('REGROUPEMENT'), DateToStr(DateDeb), Nature, not rbDateOpe.Checked)
    else
      {À la différence de GetSolde(Valeur), GetSoldeInit exécute une requête avec une inégalité stricte => pas de -1}
      Montant := GetSoldeInit(cbGeneral.Value, DateToStr(DateDeb), Nature, not rbDateOpe.Checked);
  end
  else if Nature = '' then begin
    if rbDateOpe.Checked then
      Montant := GetSolde(cbGeneral.Value, DateToStr(DateDeb - 1), S)
    else
      Montant := GetSoldeValeur(cbGeneral.Value, DateToStr(DateDeb - 1), S);
  end
  else
    {À la différence de GetSolde(Valeur), GetSoldeInit exécute une requête avec une inégalité stricte => pas de -1}
    Montant := GetSoldeInit(cbGeneral.Value, DateToStr(DateDeb), Nature, not rbDateOpe.Checked);

  {Calcul des soldes des dates de la fourchette de sélection}
  for I := 0 to NbCol - 1 do begin
    S := DateToStr(PlusMois(DateDe, I));
    Mois := DateToMois(S);

    {Si le début de millésime figure dans la fourchette de date}
    if (Mois = DateToMois(DateToStr(DebutAnnee(PlusMois(DateDe, I))))) then begin
      {Pour afficher le solde en rouge vif !!}
        {Si premier mois de l'année et que l'on est pas sur la première colonne, on récupère le solde forcé.
         Remarque : si I = 0, le problème du solde forcé est géré Dans les GetSoldexxx}
      if I > 0 then TableRouge[I + COL_DATEDE] := 'X';
      {Calcul du montant de réinitialisation}
      if GererMilles then begin
        {09/11/06 : GetSoldeMillesime ne gère pas le filtre sur les dossiers}
        if (GetControlText('REGROUPEMENT') <> '') and (cbGeneral.Value = '') then
          MntTmp := GetSoldeInitDossier(GetControlText('REGROUPEMENT'), DateToStr(DebutAnnee(StrToDate(S))), Nature, not rbDateOpe.Checked)
        else
          {Récupération du solde forcé}
          MntTmp := GetSoldeMillesime(cbGeneral.Value, DateToStr(DebutAnnee(StrToDate(S))), Nature, not rbDateOpe.Checked, True);
        {Récupération du solde du 31}
        MntTmp := (Montant - MntTmp);
        {Affectation de la tob Reinitialisation}
        TobL := TobGrid.FindFirst(['TYPEFLUX'], [CODEREGULARIS], True);
        TobL.PutValue(Mois, MntTmp);
        {Affectation de la tob Total Reinitialisation}
        TobT.PutValue(Mois, MntTmp);
        {Il faut remettre à jour le total des dépenses}
        if Assigned(TobD) then begin
          MntTmp := TobGrid.Somme(Mois, ['TYPEFLUX'], ['-'], False);
          TobD.PutValue(Mois, MntTmp);
        end;
      end;
    end;

    {Solde du mois précédent, affecté à la première tob fille}
    TobGrid.Detail[0].PutValue(Mois, Montant);
    {Calcul du solde à la fin du jour}
    Montant := Montant + TobGrid.Somme(Mois, ['TYPEFLUX'], ['+'], False) - TobGrid.Somme(Mois, ['TYPEFLUX'], ['-'], False);
    {On affecte le solde à la dernière tob fille}
    TobG.PutValue(Mois, Montant);
  end;
  {$IFDEF TOB_DEBUG}
  TobDebug(TobGrid);
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
function TOF_TRPREVMENSUEL.FourchetteDate(dtDeb : string; DebOk : Boolean; NbCol : Integer) : TDateTime;
{---------------------------------------------------------------------------------------}
begin
  if DebOk then Result := DebutDeMois(StrToDate(dtDeb))
           else Result := FinDeMois(PlusMois(StrToDate(dtDeb), NbCol));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRPREVMENSUEL.OnUpdate ;
{---------------------------------------------------------------------------------------}
var
  Q      : TQuery;
  SQL    : string;
  NbCol  : Integer; {Nombre de colonne date affichées}
  DateDe : TDateTime;
  I      : Integer;
  NSepar : Integer;
  Clause : string;
  ChpGlo : string;
begin
  inherited ;
  {Maj des devises}
  InArgument := True; {Pour ne pas lancer une recherche et faire une boucle sans fin}
  GeneralOnChange(cbGeneral);
  InArgument := False;

  DateDe := StrToDate(GetControlText('DATEDE'));

  if cbGeneral.Value <> '' then begin
    Clause := 'AND TE_GENERAL="' + cbGeneral.Value + '" ';
    sTypeMnt := 'TE_MONTANTDEV';
  end
  else
    sTypeMnt := 'TE_MONTANT';

  if (GetControlText('MCNATURE') <> '') and (Pos('<<', GetControlText('MCNATURE')) = 0) then
    {Par contre si filtre sur la nature, il ne faut pas exclure les montants forcés}
    Nature := 'AND (TE_NATURE IN (' + GetClauseIn(GetControlText('MCNATURE')) + ') OR TE_QUALIFORIGINE = "' + CODEREINIT + '")'
  else
    Nature := '';

  NbCol := seInterval.Value;
  Clause := Clause + Nature + ' AND ' + sDateOpe + ' >= "' + UsDateTime(DebutDeMois(DateDe) - 1) + '" ' +
                              ' AND ' + sDateOpe + ' <= "' + UsDateTime(PlusDate(DateDe, NbCol, 'M') + 1) + '" ';

  {25/08/06 : Gestion du filtre sur le dossier}
  if IsTresoMultiSoc and (GetControlText('REGROUPEMENT') <> '') then
    Clause := Clause + ' AND ' + GetWhereDossier;

  Clause := Clause + GetWhereConfidentialite;

  {Gestion des montants globalisés}
  if ckGlobal.Checked then ChpGlo := ''
                      else ChpGlo := ', TE_NATURE';

  SQL := 'SELECT TRECRITURE.'+sDateOpe+',TRECRITURE.TE_CODEFLUX, TRECRITURE.' + sTypeMnt + ChpGlo + ',' +
       'FLUXTRESO.TFT_FLUX,FLUXTRESO.TFT_TYPEFLUX,FLUXTRESO.TFT_LIBELLE, TRECRITURE.TE_CLEVALEUR, ' +
       'TYPEFLUX.TTL_SENS,TYPEFLUX.TTL_LIBELLE ' +
       'FROM TRECRITURE, FLUXTRESO, TYPEFLUX, BANQUECP ' +
       'WHERE TE_CODEFLUX = TFT_FLUX AND TFT_TYPEFLUX = TTL_TYPEFLUX AND BQ_CODE = TE_GENERAL ' + Clause +
       //'WHERE TE_CODEFLUX=TFT_FLUX AND TFT_TYPEFLUX=TTL_TYPEFLUX AND TE_CODEFLUX <> "' + CODEREGULARIS + '" ' + Clause +
       ' UNION ALL ' +
       'SELECT TRECRITURE.' + sDateOpe + ', TRECRITURE.TE_CODEFLUX, TRECRITURE.' + sTypeMnt + ChpGlo + ',' +
       'RUBRIQUE.RB_RUBRIQUE TFT_FLUX, RUBRIQUE.RB_TYPERUB TFT_TYPEFLUX, RUBRIQUE.RB_LIBELLE TFT_LIBELLE, TRECRITURE.TE_CLEVALEUR, ' +
       'RUBRIQUE.RB_SIGNERUB TTL_SENS, "' + CODETEMPO + '" TTL_LIBELLE ' +
       'FROM TRECRITURE, RUBRIQUE, BANQUECP ' +
        {08/11/07 : FQ 10534 : ajout du test sur la classe Trésorerie}
       'WHERE TE_CODEFLUX = RB_RUBRIQUE AND RB_CLASSERUB = "TRE" AND BQ_CODE = TE_GENERAL ' + Clause +
       {11/04/05 : FQ 10238 : on trie sur les codes flux -- 28/06/05 : et éventuellement par nature}
       ' ORDER BY TTL_SENS, TFT_TYPEFLUX, TE_CODEFLUX' + ChpGlo;

  Q := OpenSQL(SQL, True);
  TobListe.LoadDetailDB('1', '', '', Q, False, True);
  Ferme(Q);

  NSepar := Afficher(NbCol, DateDe);

  Grid.ColCount := COL_DATEDE + nbCol;

  {Formatage des montant : '#,##0.00'; Euros}
  SQL := StrFMask(CalcDecimaleDevise(Devise), '', True);
  {Formatage des dates}
  for I := COL_DATEDE to COL_DATEDE + NbCol - 1 do
    Grid.ColFormats[I] := SQL;

  TobGrid.PutGridDetail(Grid, True, True, '', True);

  Grid.Cells[COL_LIBELLE, 0] := '';

  if NSepar > 0 then // Un séparateur est présent
    Grid.RowHeights[NSepar] := 3;

  Grid.Col := COL_DATEDE;
  Grid.Row := Grid.FixedRows;
  Grid.SetFocus;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRPREVMENSUEL.bGraphClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Titre  : string;
  sSolde : string;
  sDev   : string;
  n, p   : Integer;
  T, D   : TOB;
  sGraph : string;
  dTmp   : Double;
begin
  inherited;
  if rbDateOpe.Checked then sSolde := 'Les montants sont calculés en date d''opération'
                       else sSolde := 'Les montants sont calculés en date de valeur';
  if (cbGeneral.Value = '') or ( Pos('>>', cbGeneral.Value) > 0) then begin
    Titre := 'Budget mensuel général';
    sDev  := 'Les montants sont éxprimés en ' + RechDom('TTDEVISE', V_PGI.DevisePivot, False);
  end
  else begin
    Titre := 'Budget mensuel du compte : ' + cbGeneral.Items[cbGeneral.ItemIndex];
    sDev  := 'Les montants sont éxprimés en ' + RechDom('TTDEVISE', GetControlText('DEV'), False);
  end;

  {Constitution de la chaine qui servira aux paramètres de "LanceGraph"}
  sGraph := 'MOIS;RECCPT;DEPCPT;RECTRE;DEPTRE;SOLDE';
  {Constitution de la tob du graph}
  T := Tob.Create('$$$', nil, -1);

  for n := COL_DATEDE to Grid.ColCount - 1 do begin
    {Les abscisses contiennent les mois qui sont placés en première série}
    D := Tob.Create('$$$', T, -1);
    D.AddChampSupValeur('MOIS', Grid.Cells[n, 0]);
    for p := 2 to Grid.RowCount - 2 do begin
      {On est sur un total recettes ou dépenses}
      if Grid.Cells[0, p] = '£' then begin
        if Grid.Cells[1, p] = '+' then begin
          {S'il y a des imports comptables}
          if RecCPtOk then dTmp :=  Valeur(Grid.Cells[n, p - 1])
                      else dTmp := 0.0;
          {Le total des imports de compta sont justes au-dessus du total}
          D.AddChampSupValeur('RECCPT', dTmp);
          {Le total de la tréso est donc egal au total - les imports compta}
          D.AddChampSupValeur('RECTRE', Valeur(Grid.Cells[n, p]) - dTmp);
        end
        else begin
          {S'il y a des imports comptables}
          if DepCPtOk then dTmp :=  Valeur(Grid.Cells[n, p - 1]) * -1
                      else dTmp := 0.0;
          {Le total des imports de compta sont justes au-dessus du total}
          D.AddChampSupValeur('DEPCPT', dTmp);
          {Le total de la tréso est donc egal au total - les imports compta}
          D.AddChampSupValeur('DEPTRE', (Valeur(Grid.Cells[n, p]) - dTmp) * -1);
        end;
      end;
    end;
    D.AddChampSupValeur('SOLDE', Grid.Cells[n,Grid.RowCount - 1]);
  end;

  TheTOB := T;
  {!!! Mettre sGraph en fin à cause des ";" qui traine dans la chaine}
  TRLanceFiche_GraphMensu(Titre + ';' + sDev + ';' + sSolde + ';' + sGraph);
  T.ClearDetail;
  FreeAndNil(T);
end;

{---------------------------------------------------------------------------------------}
function TOF_TRPREVMENSUEL.IsCelulleReinit(Col, Row : Integer) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  inherited IsCelulleReinit(Col, Row);
  Result := (Grid.Cells[1, Row] = CODEREGULARIS) and (TableRouge[Col] = 'X');
end;

initialization
  RegisterClasses([TOF_TRPREVMENSUEL]);

end.

