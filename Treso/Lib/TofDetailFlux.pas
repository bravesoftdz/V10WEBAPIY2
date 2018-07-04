{-------------------------------------------------------------------------------------
   Version    |  Date  | Qui | Commentaires
--------------------------------------------------------------------------------------
               20/03/02  BT   Création de l'unité
 1.20.001.001  08/03/04  JP   Autoriser la suppression d'écritures non réalisées FQ 10010
 1.2X.000.001  15/04/04  JP   Nouvelle gestion des natures (R, S, P, "" ?)
 1.90.xxx.xxx  25/06/04  JP   Création de l'impression de la Grille FQ 10100
 1.90.xxx.xxx  29/06/04  JP   Affichage de la date Valeur/Opération dans la Grille FQ 10100
 6.xx.xxx.xxx  20/07/04  JP   Gestion des commissions dans la réalisation et l'intégration
 6.xx.xxx.xxx  30/07/04  JP   Appel depuis la fiche de suivi des commissions
 6.00.018.002  14/10/04  JP   Gestion de l'affichage des écritures pointées pour le solde bancaire
 6.30.001.006  05/04/05  JP   FQ 10240 : Correction de la requête sur les écritures du suivi bancaire
 6.30.001.007  18/04/05  JP   FQ 10237 : Gestion de la périodicité sur les fiches de suivi
                              Ménage dans le OnArgument qui devenait illisible
 6.50.001.018  14/09/05  JP   FQ 10282 : Ajout d'une protection lors de la définition du range pour
                              appeler la tom TRECRITURE et le lancement des traitements
 6.50.001.020  29/09/05  JP   FQ 10296 : On ne tenait pas compte du général passé en paramètre lors de l'appel
                              de la fiche car il y avait un doublon entre les variables General et FGeneral
 6.50.001.021  03/10/05  JP   FQ 10277 : En détail des commissions, on remplace le général par
                              le code flux.
 6.51.001.001  30/11/05  JP   FQ 10303 : Sur le suivi des commission, la jointure sur TE_NUMEROPIECE est de trop
 7.01.001.001  23/03/06  JP   Gestion de TE_NODOSSIER
 7.09.001.001  18/09/06  JP   Mise en place de l'intégration Multi-sociétés
 7.09.001.001  31/10/06  JP   FQ 10338 : Dans le prévisionnel on affiche les non pointées et les pointées
                              ultérieurement => affichage de la date de rappro
 7.06.001.002  11/01/07  JP   FQ 10399 : plantage sur le changement des dates de valeurs (Clevaleur n'est plus
                              la dernière colonne depuis qu'il y a le NoDossier) +
                              On désactive le bouton si on est en date de valeur, comme c'est fait pour le popup
                              car en date d'opération, on n'a pas la clef dans la requête
 8.01.001.001  12/12/06  JP   Ajout du contrôle des soldes
 8.00.001.019  10/06/07  JP   FQ 10432 : gestion de l'appel depuis les échelles d'intérêts
 8.00.001.021  20/06/07  JP   FQ 10480 : gestion des concepts
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
 8.10.001.013  09/10/07  JP   FQ 10522 : Corrections du contrôle des soldes.
                              branchement de l'objet état pour l'impression de l'écran.
--------------------------------------------------------------------------------------}
unit TofDetailFlux ;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  Controls, Classes, Commun, Constantes, HQry, HMsgBox,
  {$IFDEF EAGLCLIENT}
  MaineAGL, UtileAgl,
  {$ELSE}
  FE_Main, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} EdtREtat,
  {$ENDIF}
  Forms, SysUtils, HCtrls, HEnt1, UTOF, Grids, UTOB, Menus, ParamSoc,
  ULibPieceCompta, HStatus;

type
  TOF_DETAILFLUX = class (TOF)
    procedure OnArgument (S : string); override;
    procedure OnClose                ; override;
  private
    FAppelant  : TDescendant;
    FPeriode   : TypePeriod;
    FDateDeb   : TDateTime;
    FDateFin   : TDateTime;
    FDateModif : TDateTime;
    FGeneral   : string;
    FNature    : string;
    FDateOpe   : string;
    RequeteSQL : string;
    TobArgument: TOB;

    function  SetFVariables     : Boolean;
    function  GetFiltreDossier  : string;
    procedure SetAppelant       (Value : string);
    procedure SetPeriode        (Value : string);
    procedure GererDetailTreso  (Arg : string);
    procedure InitVariable      ;
    procedure PrepareRequete    ;
    procedure PrepareRequeteCtrl; {12/12/06 : Contrôle des soldes}
  protected
    TREcritureModified : string;
    PopupMenu : TPopUpMenu;
    Grid      : THGrid;
    FinMois   : string;
    {29/09/05 : FQ 10296 : c'etait en doublon avec FGeneral
     General   : string; {Compte général sur lequel on est}
    Regle     : TReglesInteg;
    TobGrid   : TOB;
    TobGen    : TobPieceCompta;
    EchelleOk : Boolean; {10/06/07 : FQ 10432}
    WhereConf : string; {10/08/07 : Gestion des confidentialités}

    procedure OnPopUpMenu           (Sender : TObject);
    procedure PopupRealiseOnClick   (Sender : TObject);
    procedure PopupChangeDateOnClick(Sender : TObject);
    procedure ToutSelectionner      (Sender : TObject);
    procedure ToutDeSelectionner    (Sender : TObject);
    {JP 16/09/03 : Saisie d'écritures de simulation}
    procedure PopupSimulation       (Sender : TObject);
    procedure BImprimerClick        (Sender : TObject);
    procedure ListeOnDbleClick      (Sender : TObject);
    procedure ListeOnDbleClickCP    (Sender : TObject); {12/12/06 : Affichage de la ligne comptable}
    procedure ListeOnDbleClickBQ    (Sender : TObject); {09/10/07 : Affichage de la ligne EEXBQLIG}
    {JP 08/03/04 : Suppression des écritures non réalisées}
    procedure BDeleteClick          (Sender : TObject);
    {JP 20/11/03 : pour interdir la sélection des écritures réalisées}
    procedure ListeOnBeforeFlip     (Sender : TObject ; ARow : Integer ; Var Cancel : Boolean);

    {JP 28/08/03 : Intégration en comptabilité}
    procedure IntegreCompta;
    procedure PrepareEcriture(NumTransac, NumPiece : string);

    {20/04/05 : Méthodes appelées dans le Onargument}
    procedure GestionControls  ;
    procedure GestionBoutons   (VisibleOk : Boolean = True); {10/06/07 : FQ 10432 : ajout du paramètre}
    procedure PrepareGrille    (Dev : string);
    procedure PrepareGrilleCtrl(Dev : string);
    procedure MajCaptionForm   ;
  end;

function TRLanceFiche_DetailFlux(Dom, Fiche, Range, Lequel, Arguments: string): string;
function InitTobPourDetail : TOB;

implementation

uses
  {$IFDEF TRCONF}
  ULIbConfidentialite,
  {$ENDIF TRCONF}
  HTB97, ParamDat, Vierge, TRSAISIEFLUX_TOF, UProcCommission, UProcEcriture, UProcGen,
  UProcSolde, UtilPgi, TRLIGNEECRCOMPTA_TOF, UlibPointage, uObjEtats, EEXBQLIG_TOM, AglInit;

const
  COL_DATEAFF  = 0;
  COL_NATURE   = 1;
  COL_COMPTE   = 2; // Colonne Compte / Code flux
  COL_NUMPIECE = 3;
  COL_TRANSAC  = 4;
  COL_CIB      = 5;
  COL_LIBELLE  = 6;
  COL_DEVISE   = 7;
  COL_MONTANT  = 8;
  COL_NUMLIGNE = 10;

{---------------------------------------------------------------------------------------}
function TRLanceFiche_DetailFlux(Dom, Fiche, Range, Lequel, Arguments: string): string;
{---------------------------------------------------------------------------------------}
begin
  Result := AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
function InitTobPourDetail : TOB;
{---------------------------------------------------------------------------------------}
begin
  Result := TOB.Create('$ùùù', nil, -1);
  Result.AddChampSupValeur('CODE'    , '');
  Result.AddChampSupValeur('GENERAL' , '');
  Result.AddChampSupValeur('NATURE'  , '');
  Result.AddChampSupValeur('APPEL'   , '');
  Result.AddChampSupValeur('PERIODICITE', '');
  Result.AddChampSupValeur('DATEDEB' , iDate1900);
  Result.AddChampSupValeur('DATEFIN' , iDate1900);
  Result.AddChampSupValeur('DATEOPE' , '');
  Result.AddChampSupValeur('DATEAFF' , '');
  Result.AddChampSupValeur('LIBELLE' , '');
  Result.AddChampSupValeur('LIBNAT'  , '');
  Result.AddChampSupValeur('MOIS'    , '');
  Result.AddChampSupValeur('TYPE'    , '');
  Result.AddChampSupValeur('DEVISE'  , '');
  Result.AddChampSupValeur('REGROUP' , '');
  Result.AddChampSupValeur('INTERVAL', '');
  Result.AddChampSupValeur('BANQUE'  , '');
  Result.AddChampSupValeur('NOMBASE' , ''); {12/12/06 : contrôle des soldes}
  Result.AddChampSupValeur('LEFTJOIN', ''); {10/06/07 : FQ 10432}
  Result.AddChampSupValeur('WHERE'   , ''); {10/06/07 : FQ 10432}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.OnArgument (S : string ) ;
{---------------------------------------------------------------------------------------}
var
  S1 : string;
begin
  inherited;
  Ecran.HelpContext := 150;
  {Initialisation des variables de traitement}
  InitVariable;
  {Récupération de la fiche appelante}
  SetAppelant(ReadTokenSt(S));
  {20/04/05 : Gestion des Contrôles autres que les boutons}
  GestionControls;

  {Récupération du type de périodicité}
  {Le suivi du trésorier est trop différent pour être gérer d'une manière standard}
  if FAppelant = dsTreso then begin
    GererDetailTreso(S);
    Exit;
  end
  else begin
    SetPeriode(TobArgument.GetString('PERIODICITE'));
    SetFVariables;
  end;

  {10/06/07 : FQ 10432 : gestion de l'appel depuis les échelles d'intérêts}
  EchelleOk := ReadTokenSt(S) = 'ECHINT';

  {Préparation de la requête}
  PrepareRequete;
  TobGrid.LoadDetailDBFromSQL('1', RequeteSQL);
  {Théoriquement le résultat de l requête ne renvoie qu'une même devise}
  if TobGrid.Detail.Count > 0 then
    S1 := TobGrid.Detail[0].GetString('TE_DEVISE');

  TobGrid.PutGridDetail(Grid, False, False, '', True);
  {20/04/05 : Dessin de la grille}
  PrepareGrille(S1);
  {20/04/05 : Gestion des boutons
   10/06/07 : FQ 10432 : gestion de l'appel depuis les échelles d'intérêts}
  EchelleOk := ReadTokenSt(S) = 'ECHINT';
  GestionBoutons(not EchelleOk);
  {20/04/05 : Mise à jour tu titre de la forme}
  MajCaptionForm;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.OnClose ;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(TobGrid) then FreeAndNil(TobGrid);
  if Assigned(TobGen ) then FreeAndNil(TobGen);

  TFVierge(Ecran).Retour := TREcritureModified;

  if FDateModif < iDate2099 then
    TFVierge(Ecran).Retour := DateToStr(FDateModif - 1); // La veille, pour recalcul des soldes
  inherited;
end ;

{Prépare le menu popup
{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.OnPopupMenu(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if PopupMenu.Items.Count = 0 then Exit;

  {Les simulations de virements ne peuvent être passées qu'en "V" lors de la génération du virement lié}
  PopupMenu.Items[0].Enabled := (Grid.nbSelected > 0   )  and
                                (Pos('"' + na_Realise   + '"', FNature) = 0) and
                                {JP 15/04/04 : les écritures de l'échéancier sont modifiables
                                (Pos('"' + na_ImpCompta + '"', Nature) = 0) and}
                                {Il faut les droits pour injecter en compta}
                                V_PGI.Superviseur; // ?
  PopupMenu.Items[1].Enabled := (Grid.nbSelected > 0) and (FDateOpe = 'TE_DATEVALEUR');
  {JP 08/03/04 : FQ 10010 : mise en place de la suppression}
  PopupMenu.Items[3].Enabled := (Grid.nbSelected > 0);
end;

{21/07/04 : on ne traite plus que les écritures de saisie ; ancien code en fin d'unité
 Passe les lignes d'écritures sélectionnées à réalisé
 Rappel     : Les lignes d'écritures sont définies par un Numéro de transaction,
              un numéro de pièce et un muméro de ligne
 Les étapes :
    1/ Il faut que les écritures de la transaction dont le numéro de ligne est inférieur à
       celui que de la ligne en cours soient déjà réalisées
    2/ Si c'est le cas, on "réalise" la ligne en cours
    3/ On regarde si la transaction de le ligne en cours a encore des écritures non réalisées
    4/ Si ce n'est pas le cas, on dénoue la transaction et on propose l'intégration en compta
{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.PopupRealiseOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  I : Integer;
  N : string;
  Ok : Boolean;
begin
  if Grid.nbSelected = 0 then begin
    MessageAlerte('Veuillez sélectionner des écritures.');
    Exit;
  end;

  Ok := False;

  if trShowMessage(Ecran.Caption, 20, '', '') = mrYes then begin
    InitMove(Grid.nbSelected, 'Réalisation des écritures');
    for I := Grid.RowCount - 1 downto 1 do begin
      Movecur(False);
      {On ne traite que les lignes sélectionnées qui sont prévisionnelles}
      if Grid.IsSelected(I) and (Grid.Cells[COL_NATURE, I] = na_Prevision) and
         {21/07/04 : on ne traite plus que les écritures de saisie}
         (Pos(TRANSACSAISIE, UpperCase(Grid.Cells[COL_TRANSAC, I])) >= 0) then begin

        N := TOB(Grid.Objects[0, I]).GetValue('TE_NUMLIGNE');

        {On passe toute l'écriture à réalisé, opération et commissions
         14/09/05 : FQ 10282 : Ajout de Trim et de FloatToStr(Valeur()), car il semble qu'il y ait
                               maintenant un formatage automatique des zones}
        UpdatePieceStr(TOB(Grid.Objects[0, I]).GetString('TE_NODOSSIER'), Trim(Grid.Cells[COL_TRANSAC, I]),
                       FloatToStr(ValeurI(Grid.Cells[COL_NUMPIECE, I])), '', 'TE_NATURE', na_Realise);

        {Intégration en comptabilité}
        if (GetParamSocSecur('SO_TRINTEGAUTO', False) = True)  or OK or
           (HShowMessage('0;Intégration en comptabilité;Voulez-vous intégrer' +
                        ' en comptabilité les lignes de la transaction ?;Q;YN;N;N;', '', '') = mrYes) then begin
          Ok := True;

          {De même, on intègre toute la pièce en comptabilité
           14/09/05 : FQ 10282 : Ajout de Trim et de FloatToStr(Valeur()), car il semble qu'il y ait
                                 maintenant un formatage automatique des zones}
          PrepareEcriture(Trim(Grid.Cells[COL_TRANSAC, I]), FloatToStr(ValeurI(Grid.Cells[COL_NUMPIECE, I])));

          {Si les écritures affichées n'étaient pas filtrées}
          if (FNature = '') or (Pos(na_Realise, FNature) > 0) then
            {On passe la nature à réalisé}
            Grid.Cells[COL_NATURE, I] := na_Realise
          else
            {Sinon, on supprime la ligne}
            Grid.DeleteRow(I);

          TREcritureModified := na_Realise;
        end;
      end;
    end;
    FiniMove;
    if Ok and AutoriseFonction(dac_Integration) then IntegreCompta;
    Grid.ClearSelected;
  end;
end;

{Mise à jour de la date uniquement dans le cas d'un traitement par date de valeur
 Il faut aussi mettre à jour la CleValeur qui est composée de la date et d'un numéro unique
{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.PopupChangeDateOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  C     : Char;
  aDate : THEdit;
  I     : Integer;
  D1, D2: TDateTime;
  Clef  : string;
  dtTemp: string;
begin
  if Grid.nbSelected = 0 then begin
    MessageAlerte('Veuillez sélectionner des écritures.');
    Exit;
  end;

  dtTemp := TobArgument.GetString('DATEDEB');
  D2 := StrToDate(dtTemp);

  aDate := THEdit.Create(Ecran);
  // aDate.OpeType := otDate; ?
  aDate.Parent := Grid; // Grid.ClientOrigin
  aDate.Left := 100;
  aDate.Top := 12;
  aDate.Visible := False;
  aDate.Text := dtTemp;
  C := ' ';
  try
    ParamDate(Ecran, aDate, C);
    if aDate.Text <> dtTemp then {123456}
      if trShowMessage(Ecran.Caption, 20, '', '') = mrYes then
        for I := Grid.RowCount-1 downto 1 do
          {On balaie les lignes selectionnées}
          if Grid.IsSelected(I) then begin
            D1 := StrToDate(aDate.Text);
            {11/01/07 : FQ 10399 : ColCount - 2, car maintenant ColCount - 1 est pour no Dossier}
            Clef := RetourneCleEcriture(D1, StrToInt(Copy(Grid.Cells[Grid.ColCount - 2, I], 7, 7)));
            ExecuteSQL('UPDATE TRECRITURE SET TE_DATEVALEUR = "' + UsDateTime(D1) + '", TE_CLEVALEUR = "' +
                       Clef + '" WHERE TE_CLEVALEUR = "' + Grid.Cells[Grid.ColCount - 2, I] + '"');
            {20/11/06 : éventuel recalcul des soldes si la date change de millésime}
            GereSoldeInitClef(Grid.Cells[Grid.ColCount - 2, I], Clef,
                             TobArgument.GetString('GENERAL'), Valeur(Grid.Cells[COL_MONTANT + 1, I]), False);  
            {Pour éviter des effets de bords, on déselectionne la ligne}
            Grid.FlipSelection(I);
            {On efface la ligne si elle n'appartient plus à la fourchette de dates}
            if not Between(StrToDate(aDate.Text), FDateDeb, FDateFin) then
              Grid.DeleteRow(I);
            {Sauve la date mini de modif du solde}
            if D1 < D2 then begin
              if D1 < FDateModif then
                FDateModif := D1;
            end
            else if D2 < FDateModif then
              FDateModif := D2;
          end;
  finally
    aDate.Free;
  end;
end;

{JP 16/09/03 : Lancement de l'écran de saisie des écritures de simulation
{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.PopupSimulation(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {29/09/05 : FQ 10296 : General etait en doublon avec FGeneral}
  TREcritureModified := TRLanceFiche_SaisieFlux('TR', 'TRSAISIEFLUX', '', '', FGeneral + ';');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.PrepareEcriture(NumTransac, NumPiece : string);
{---------------------------------------------------------------------------------------}
var
  tEcr : TOB;
  S    : string;
begin
  {Création de la requête contenant les écritures à intégrer}
  S := 'SELECT * FROM TRECRITURE WHERE TE_NUMTRANSAC = "' + NumTransac + '" AND TE_NUMEROPIECE = "' + NumPiece +
       '" AND (TE_USERCOMPTABLE = "" OR TE_USERCOMPTABLE IS NULL) ORDER BY TE_NUMLIGNE';

  tEcr := TOB.Create('ùùù', nil, -1);
  try
    tEcr.LoadDetailFromSQL(S);
    {S'il n'y a pas d'écritures pour la transaction !!?}
    if tEcr.Detail.Count = 0 then Exit;
    {18/09/06 : Nouvelle intégration des écritures en comptabilité (gestion du multi-sociétés)}
    TRGenererPieceCompta(TobGen, tEcr);
  finally
    if Assigned(tEcr) then FreeAndNil(tEcr);
  end;
end;

{JP 28/08/03 : Intégration en compta d'une transaction dénouée lors de la réalisation
               d'une écriture
{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.IntegreCompta;
{---------------------------------------------------------------------------------------}
begin
  {Par précaution, théoriquement arrivé ici, la tob ne peut être vide}
  if TobGen.Detail.Count = 0 then begin
    {06/07/05 : Il arrive que, sur Bug (!), la tob soit vide : mise en place d'un message plus neutre}
    HShowMessage('4;' + Ecran.Caption + ';Il est impossible de générer les écritures comptables.'#13 +
                 'Veuillez vérifier votre paramétrage.;W;O;O;O;', '', '');
    Exit;
  end;

  {18/09/06 : Nouvelle fonction de traitement Multi-sociétés}
  TRIntegrationPieces(TobGen, True);

  TobGen.ClearDetailPC;
end;

{Imprime la grille rajoutée au lieu de la DB grille
{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.BImprimerClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  {25/06/04 : pour gérer le cas de l'eAgl (pas de PrintDBGrid) et le titre sur deux lignes}
  if TobGrid.Detail.Count = 0 then Exit;
  T := TOB.Create('$TRECRITURE', nil, -1);
  try
    T.Dupliquer(TobGrid, True, True);
//    T.Detail[0].AddChampSupValeur('TITRE', Ecran.Caption, True);
  //  T.Detail[0].AddChampSupValeur('TE_DATELIB', Grid.Titres[COL_DATEAFF], True);
    //T.Detail[0].AddChampSupValeur('TE_GENELIB', Grid.Titres[COL_COMPTE], True);
    //LanceEtatTOB('E', 'ECT', 'DFX', T, True, False, False, nil, '', 'titre', False);
    TObjEtats.GenereEtatGrille(Grid, Ecran.Caption);
  finally
    if Assigned(T) then FreeAndNil(T);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.ListeOnBeforeFlip(Sender : TObject ; ARow : Integer ; var Cancel : Boolean);
{---------------------------------------------------------------------------------------}
var
  Ok : Boolean;
begin
  {18/09/06 : on ne permet la sélection que des écritures "saisies"}
  Ok := Copy(Grid.Cells[COL_TRANSAC, ARow], 7, 4) <> TRANSACSAISIE;
  {On n'autorise pas la sélection d'une écriture réalisée, car on ne peut la réaliser de
   nouveau ou lui changer de date de valeur}
  Cancel := (Grid.Cells[COL_NATURE, ARow] = na_Realise) or Ok;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.ToutSelectionner(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  for n := 1 to Grid.RowCount - 1 do
    if not Grid.IsSelected(n) then Grid.FlipSelection(n);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.ToutDeSelectionner(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  Grid.ClearSelected
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.ListeOnDbleClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  s : string;
  e : string;
begin
  {14/09/05 : FQ 10282 : Ajout de Trim et de FloatToStr(Valeur()), car il semble qu'il y ait
                         maintenant un formatage automatique des zones}
//  e := GetExoFromDate(VarToDateTime(Grid.Cells[COL_DATEAFF, Grid.Row]));
  {23/03/06 : Mise ne place du NODOSSIER}
  e := Trim(Grid.CellValues[Grid.ColCount - 1, Grid.Row]);

  s := e + ';' + Trim(Grid.Cells[COL_TRANSAC, Grid.Row]) + ';' + FloatToStr(ValeurI(Grid.Cells[COL_NUMPIECE, Grid.Row])) + ';' +
       FloatToStr(ValeurI(Grid.Cells[COL_NUMLIGNE, Grid.Row])) + ';';
  AGLLanceFiche('TR', 'TRFICECRITURE', '', s, Trim(Grid.Cells[COL_NATURE, Grid.Row]) + ';');
end;

{Affichage d'une écriture comptable
{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.ListeOnDbleClickCP(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  T := TOB.Create('ùùù', nil, -1);
  try
    T.AddChampSupValeur('E_JOURNAL'      , Grid.CellValues[COL_NATURE      , Grid.Row]);
    T.AddChampSupValeur('E_EXERCICE'     , Grid.CellValues[COL_NUMLIGNE + 1, Grid.Row]);
    T.AddChampSupValeur('E_NUMLIGNE'     , Grid.CellValues[COL_TRANSAC     , Grid.Row]);
    T.AddChampSupValeur('E_NUMECHE'      , Grid.CellValues[COL_NUMLIGNE    , Grid.Row]);
    T.AddChampSupValeur('E_GENERAL'      , Grid.CellValues[COL_COMPTE      , Grid.Row]);
    T.AddChampSupValeur('E_NUMEROPIECE'  , Valeur(Grid.CellValues[COL_NUMPIECE, Grid.Row]));
    T.AddChampSupValeur('E_DATECOMPTABLE', LaTob.GetDateTime('DATEDEB'));
    T.AddChampSupValeur('NOMBASE'        , LaTob.GetString('NOMBASE'));
    AffDetailEcr(T, True);
  finally
    FreeAndNil(T);
  end;
end;

{Affichage d'un mouvement bancaire
{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.ListeOnDbleClickBQ(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  lq : string;
begin
  lq := Grid.CellValues[COL_COMPTE  , Grid.Row] + ';' +
        Grid.CellValues[COL_NUMPIECE, Grid.Row] + ';' +
        Grid.CellValues[COL_NUMLIGNE, Grid.Row] + ';' ;
  CPLanceFiche_EEXBQLIG('', lq, ActionToString(taConsult) + ';;');
end;

{JP 08/03/04 : FQ 10010 : suppression des écritures non réalisées
{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.BDeleteClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  I : Integer;
begin
  if Grid.nbSelected = 0 then begin
    HShowMessage('0;' + Ecran.Caption + ';Aucun élément n''est sélectionné.;I;O;O;O', '', '');
    Exit;
  end;

  if HShowMessage('1;' + Ecran.Caption + ';Êtes-vous sûr de vouloir supprimer les éléments sélectionnés.;Q;YNC;N;N;', '', '') <> mrYes then Exit;

  InitMove(Grid.nbSelected, 'Suppression des écritures');
  for I := Grid.RowCount - 1 downto 1 do begin
    Movecur(False);
    {On ne traite que les lignes sélectionnées qui ne sont pas réalisées ni importé de comptabilité.
     JP 15/04/04 : On considère que l'on peut supprimer une écriture de l'échéancier ....
                   Sinon, il faudra gérer TE_QUALIFORIGINE}
    if Grid.IsSelected(I) and (Grid.Cells[COL_NATURE, I] <> na_Realise) and
       {15/04/04 : On exclue les écritures de comptabilité}
       (Pos(TRANSACIMPORT , UpperCase(Grid.Cells[COL_TRANSAC, I])) < 1) then begin

      {On interdit la suppression d'écritures qui appartiennent à une transaction financières
       14/09/05 : FQ 10282 : Ajout de Trim et de FloatToStr(Valeur()), car il semble qu'il y ait
                             maintenant un formatage automatique des zones}
      Q := OpenSQL('SELECT TE_NUMEROPIECE FROM TRECRITURE WHERE TE_NUMTRANSAC = "' + Trim(Grid.Cells[COL_TRANSAC, I]) +
                   '" AND TE_NUMEROPIECE <> "' + FloatToStr(ValeurI(Grid.Cells[COL_NUMPIECE, I])) + '"', True);
      try
        {Supression de l'écriture}
        if Q.RecordCount = 0 then begin
          {20/07/04 : suppression des écritures avec la gestion des commissions
           14/09/05 : FQ 10282 : Ajout de Trim et de FloatToStr(Valeur()), car il semble qu'il y ait
                                 maintenant un formatage automatique des zones}
          if not SupprimeEcriture(Ecran.Caption, Trim(Grid.Cells[COL_TRANSAC, I]),
                                                 FloatToStr(ValeurI(Grid.Cells[COL_NUMPIECE, I])),
                                                 FloatToStr(ValeurI(Grid.Cells[COL_NUMLIGNE, I])),
                                                 Trim(Grid.CellValues[Grid.ColCount - 1, I])) then Break;

          {On supprime la ligne de la grille}
          Grid.DeleteRow(I);
          {Pour demander un recalcul des soldes en retour sur le détail des suivis}
          TREcritureModified := DateToStr(FDateDeb);
        end {Supression}

        else begin
          {Constitution du message d'avertissement}
          HShowMessage('4;' + Ecran.Caption + ';Certaines écritures appartiennent à des transactions financières,'#13 +
                       'et ne peuvent être supprimées. Pour les supprimer, vous devez "dévalider" la transaction.;I;O;O;O;', '', '');
        end;
      finally
        Ferme(Q);
      end;
    end;
  end;
  FiniMove;
  Grid.ClearSelected;
end;

{10/06/07 : FQ 10432 : gestion de l'appel depuis les échelles d'intérêts : dans ce cas, on
            cache tous les boutons d'office. C'est le seul cas où VisibleOk arrive à False
{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.GestionBoutons(VisibleOk : Boolean = True);
{---------------------------------------------------------------------------------------}
var
  Btn : TToolbarButton97;
  Ok1 : Boolean;
  Ok2 : Boolean;
begin
  if VisibleOk then begin
    Ok1 := not (FAppelant in [dsMensuel, dsTreso, dsControle]);
    Ok2 := not (FAppelant = dsBancaire) or (TobArgument.GetString('TYPE') = na_Prevision);
    VisibleOk := Ok1 and Ok2;
  end;
  Btn := TToolbarButton97(GetControl('BREALISE'));
  Btn.Hint := TraduireMemoire('Réaliser les écritures sélectionnées');
  Btn.OnClick := PopupRealiseOnClick;
  Btn.Visible := VisibleOk;
  {20/06/07 : FQ 10480 : Gestion du concept la création / Suppresion des flux}
  CanValidateBO(Btn);

  Btn := TToolbarButton97(GetControl('BCHANGE'));
  Btn.Hint := TraduireMemoire('Modifier la date des écritures sélectionnées');
  Btn.OnClick := PopupChangeDateOnClick;
  {11/01/07 : FQ 10399 : n'est valable que si l'on vient d'un suivi en valeur}
  Btn.Visible := VisibleOk and (FDateOpe = 'TE_DATEVALEUR');

  Btn := TToolbarButton97(GetControl('BSAISIE'));
  Btn.Hint := TraduireMemoire('Saisie d''une nouvelle écriture');
  Btn.OnClick := PopupSimulation;
  {20/06/07 : FQ 10480 : Gestion du concept la création / Suppresion des flux}
  CanCreateEcr(Btn);

  Btn.Visible := VisibleOk;
  Btn := TToolbarButton97(GetControl('BSELECT'));
  Btn.Hint := TraduireMemoire('Sélectionner toutes les lignes');
  Btn.OnClick := ToutSelectionner;
  Btn.Visible := VisibleOk;
  Btn := TToolbarButton97(GetControl('BDESELECT'));
  Btn.Hint := TraduireMemoire('Désélectionner toutes les lignes');
  Btn.OnClick := ToutDeSelectionner;
  Btn.Visible := VisibleOk;
  {JP 08/03/04 : FQ 10010 : mise en place de la suppression}
  Btn := TToolbarButton97(GetControl('BDELETE'));
  Btn.Hint := TraduireMemoire('Suppression des lignes sélectionnées');
  Btn.OnClick := BDeleteClick;
  Btn.Visible := VisibleOk;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.GestionControls;
{---------------------------------------------------------------------------------------}
begin
  PopupMenu := TPopUpMenu(GetControl('POPUPMENU'));
  PopupMenu.OnPopup := OnPopupMenu;
  PopupMenu.Items[0].OnClick := PopupRealiseOnClick; // Passe en réalisé
  PopupMenu.Items[1].OnClick := PopupChangeDateOnClick;
  PopupMenu.Items[2].OnClick := PopupSimulation;
  PopupMenu.Items[3].OnClick := BDeleteClick;{JP 08/03/04 : suppression d'écritures}
  PopupMenu.Items[4].OnClick := ToutSelectionner;
  PopupMenu.Items[5].OnClick := ToutDeSelectionner;

  {20/06/07 : FQ 10480 : Gestion du concept la création / Suppresion des flux}
  CanCreateEcr(PopupMenu.Items[2]);
  CanValidateBO(PopupMenu.Items[0]);

  Grid := THGrid(GetControl('GRID'));
  TToolbarButton97(GetControl('BIMPRIMER')).OnClick := BImprimerClick;

  {On limite les fonctions pour les fiches de budget et de Treso}
  if FAppelant in [dsTreso, dsMensuel, dsControle] then begin
    PopupMenu.Items.Clear;
    Grid.MultiSelect := False;
  end
  else
    {Ajout du popup au menu 27}
    AddMenuPop(PopupMenu, '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.PrepareGrille(Dev : string);
{---------------------------------------------------------------------------------------}
begin
  {1/ PRÉPARATION DE LA REQUÊTE}
  if FAppelant = dsControle then begin
    PrepareGrilleCtrl(Dev);
    Exit;
  end;

  {24/03/06 : on cache, pour le moment le NODOSSIER qui a été ajouté à toutes les requêtes
              comme dernier champ}
  Grid.ColWidths[Grid.ColCount - 1] := 0;

  {Mise à jour des libellés de la grille}
  if FAppelant in [dsDetail, dsMensuel, dsTreso] then Grid.Titres[COL_COMPTE] := 'Compte'
                                                 else Grid.Titres[COL_COMPTE] := 'Code flux';

  {29/06/04 : Ajout de la date complémentaire dans la grille}
  if FDateOpe = 'TE_DATERAPPRO' then begin
   {15/04/05 : FQ 10240 : Gestion de la date d'affichage}
    if TobArgument.GetString('DATEAFF') <> 'TE_DATEVALEUR' then Grid.Titres[COL_DATEAFF] := 'Date Opé.'
                                                           else Grid.Titres[COL_DATEAFF] := 'Date Val.';
  end
  else if FDateOpe = 'TE_DATEVALEUR' then begin
    {31/10/06 : FQ 10338 : affichage de la date de rappro sur le détail du suivi bancaire}
    if TobArgument.GetString('DATEAFF') = 'TE_DATERAPPRO' then Grid.Titres[COL_DATEAFF] := 'Date Rap.'
                                                          else Grid.Titres[COL_DATEAFF] := 'Date Opé.';
    {Si on est en date de valeur, on cache la colonne contenant la CleValeur}
    Grid.ColWidths[Grid.ColCount - 2] := 0;
  end
  else
    Grid.Titres[COL_DATEAFF] := 'Date Val.';

  {Ajustement des colonnes}
  Grid.ColWidths[COL_TRANSAC] := Grid.ColWidths[COL_TRANSAC] + 15;
  Grid.ColWidths[COL_COMPTE]  := Grid.ColWidths[COL_COMPTE]  - 10;

  Grid.UpdateTitres;
  {Gestion des formats des montants}
  if Dev = V_PGI.DevisePivot then Grid.ColFormats[COL_MONTANT] := StrFMask(V_PGI.OkDecV, '', True)
                             else Grid.ColFormats[COL_MONTANT] := StrFMask(CalcDecimaleDevise(Dev), '', True);
  Grid.ColAligns[COL_MONTANT] := taRightJustify;
  Grid.ColFormats[COL_MONTANT + 1] := StrFMask(V_PGI.OkDecV, '', True);
  Grid.ColAligns [COL_MONTANT + 1] := taRightJustify;

  Grid.OnBeforeFlip := ListeOnBeforeFlip;
  Grid.OnDblClick   := ListeOnDbleClick;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.PrepareGrilleCtrl(Dev : string);
{---------------------------------------------------------------------------------------}
begin
  if TobGrid.Detail.Count = 0 then Exit;
  {24/03/06 : on cache, pour le moment le NODOSSIER qui a été ajouté à toutes les requêtes
              comme dernier champ}
  Grid.ColWidths[Grid.ColCount - 1] := 0;

  Grid.ColTypes[COL_DATEAFF] := 'D';
  Grid.ColFormats[COL_DATEAFF] := 'dd/mm/yyyy';

  {Mise à jour des libellés de la grille}
  Grid.Cells[COL_COMPTE, 0] := 'Compte';

  {Sur TRECRITURE}
  if LaTob.GetString('TYPE') = 'T' then begin
    if LaTob.GetString('DATEAFF') = 'TE_DATEVALEUR' then Grid.Cells[COL_DATEAFF, 0] := 'Date Val.'
                                                    else Grid.Cells[COL_DATEAFF, 0] := 'Date Opé.';

    {Ajustement des colonnes}
    Grid.ColWidths[COL_TRANSAC] := Grid.ColWidths[COL_TRANSAC] + 15;
    Grid.ColWidths[COL_COMPTE]  := Grid.ColWidths[COL_COMPTE]  - 10;

    {Gestion des formats des montants}
    if Dev = V_PGI.DevisePivot then Grid.ColFormats[COL_MONTANT] := StrFMask(V_PGI.OkDecV, '', True)
                               else Grid.ColFormats[COL_MONTANT] := StrFMask(CalcDecimaleDevise(Dev), '', True);
    Grid.OnDblClick  := ListeOnDbleClick;
  end

  {Sur ECRITURE}
  else if LaTob.GetString('TYPE') = 'C' then begin
    Grid.Cells[COL_DATEAFF    , 0] := 'Date Val.';
    Grid.Cells[COL_NATURE     , 0] := 'Jal.';
    Grid.Cells[COL_MONTANT    , 0] := 'Débit Dev.';
    Grid.Cells[COL_MONTANT + 1, 0] := 'Crédit Dev.';
    Grid.Cells[COL_TRANSAC    , 0] := 'Num. Lig.';
    Grid.Cells[COL_NUMLIGNE   , 0] := 'Eche.';
    Grid.Cells[COL_CIB        , 0] := 'Paie.';
    Grid.ColWidths[COL_NUMLIGNE + 2] := 0;
    Grid.ColWidths[COL_TRANSAC     ] := 54;
    Grid.ColWidths[COL_LIBELLE     ] := Grid.ColWidths[COL_LIBELLE] + 55;
    Grid.ColWidths[COL_CIB         ] := Grid.ColWidths[COL_CIB]  + 7;

    Grid.ColFormats[COL_MONTANT] := StrFMask(V_PGI.OkDecV, '', True);
    Grid.OnDblClick  := ListeOnDbleClickCP;
  end
  {Sur EEXBQLIG}
  else if LaTob.GetString('TYPE') = 'B' then begin
    Grid.Cells[COL_NATURE     , 0] := 'Réf. Ptg.';
    Grid.Cells[COL_DATEAFF    , 0] := 'Date Opé.';
    Grid.Cells[COL_MONTANT    , 0] := 'Débit Dev.';
    Grid.Cells[COL_MONTANT + 1, 0] := 'Crédit Dev.';
    Grid.Cells[COL_NUMPIECE   , 0] := 'Rel.';
    Grid.Cells[COL_TRANSAC    , 0] := 'Lib. Enrichi.';

    Grid.ColWidths[COL_NATURE     ] := Grid.ColWidths[COL_NATURE] * 2;
    Grid.ColWidths[COL_TRANSAC    ] := Grid.ColWidths[COL_TRANSAC ] - 7;
    Grid.ColWidths[COL_NUMPIECE   ] := Grid.ColWidths[COL_NUMPIECE] - 17;

    Grid.ColFormats[COL_MONTANT   ] := StrFMask(V_PGI.OkDecV, '', True);
    Grid.OnDblClick  := ListeOnDbleClickBQ;
  end;

  Grid.ColWidths[COL_NUMLIGNE + 1] := 0;

  Grid.ColFormats[COL_MONTANT + 1] := StrFMask(V_PGI.OkDecV, '', True);

  Grid.ColAligns[COL_MONTANT + 1] := taRightJustify;
  Grid.ColAligns[COL_MONTANT    ] := taRightJustify;
  Grid.ColAligns[COL_NUMPIECE   ] := taRightJustify;
  Grid.ColAligns[COL_NUMLIGNE   ] := taRightJustify;
  Grid.ColAligns[COL_DEVISE     ] := taCenter;
  Grid.ColAligns[COL_NATURE     ] := taCenter;
  Grid.ColAligns[COL_COMPTE     ] := taCenter;
  Grid.ColAligns[COL_CIB        ] := taCenter;
  Grid.ColAligns[COL_DATEAFF    ] := taCenter;
  if LaTob.GetString('TYPE') = 'C' then Grid.ColAligns[COL_TRANSAC] := taRightJustify;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.MajCaptionForm;
{---------------------------------------------------------------------------------------}
begin
  {Suivi des commissions}
  if FAppelant = dsCommission then begin
    if TobArgument.GetString('LIBNAT') = suc_Commission then
      TForm(Ecran).Caption := 'Commissions ' + TobArgument.GetString('LIBELLE')
    else if TobArgument.GetString('LIBNAT') = suc_Divers then
      TForm(Ecran).Caption := TobArgument.GetString('LIBELLE')
    else
      TForm(Ecran).Caption := 'Base des commissions ' + TobArgument.GetString('LIBELLE');
  end

  else
    TForm(Ecran).Caption := TobArgument.GetString('LIBELLE') + TobArgument.GetString('MOIS');

  UpdateCaption(Ecran);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.GererDetailTreso(Arg : string);
{---------------------------------------------------------------------------------------}
var
  SDateAff : string;
  Code     : string;
  Where    : string;
  Clef     : string;
  s1       : string;
  sDate    : string;
begin
  FDateOpe := ReadTokenSt(Arg);
  Code := ReadTokenSt(Arg);
  FGeneral := ReadTokenSt(Arg);
  sDate := ReadTokenSt(Arg);
  TForm(Ecran).Caption := ReadTokenSt(Arg);
  UpdateCaption(Ecran);
  FNature := ReadTokenSt(Arg);

  Where := 'TE_QUALIFORIGINE = "' + Code + '" AND ' + FGeneral + ' ' + FNature;
  {15/11/06 : Gestion du filtre sur les dossiers pour le suivi Tréso et les budgets}
  Where := Where + GetFiltreDossier;

  if FDateOpe = 'TE_DATEVALEUR' then begin
    Clef := ', TE_CLEVALEUR';
    SDateAff := 'TE_DATECOMPTABLE AS TE_DATEAFF, '; {29/06/04 : Ajout de la date complémentaire dans la grille}
  end
  else begin
    SDateAff := 'TE_DATEVALEUR AS TE_DATEAFF, '; {29/06/04 : Ajout de la date complémentaire dans la grille}
  end;

  if LaTob.GetDateTime('DATEDEB') = iDate1900 then
    {23/03/06 : Ajout du NO_DOSSIER dans la requête}
    RequeteSQL := 'SELECT ' + SDateAff + 'TE_NATURE, TE_GENERAL AS TE_GENE, TE_NUMEROPIECE, TE_NUMTRANSAC, TE_CODECIB, TE_LIBELLE,' +
                  'TE_DEVISE, TE_MONTANTDEV, TE_MONTANT, TE_NUMLIGNE' + Clef + ', TE_NODOSSIER FROM TRECRITURE '+
                  'LEFT JOIN BANQUECP on BQ_CODE = TE_GENERAL ' +
                  'WHERE ' + FDateOpe + '= "' + UsDateTime(StrToDate(sDate)) + '" AND ' + Where + WhereConf +
                  ' ORDER BY TE_NUMEROPIECE, TE_GENERAL'
  else
    {14/11/06 : On est en affichage mensuel}
    RequeteSQL := 'SELECT ' + SDateAff + 'TE_NATURE, TE_GENERAL AS TE_GENE, TE_NUMEROPIECE, TE_NUMTRANSAC, TE_CODECIB, TE_LIBELLE,' +
                  'TE_DEVISE, TE_MONTANTDEV, TE_MONTANT, TE_NUMLIGNE' + Clef + ', TE_NODOSSIER FROM TRECRITURE '+
                  'LEFT JOIN BANQUECP on BQ_CODE = TE_GENERAL ' +
                  'WHERE ' + FDateOpe + ' BETWEEN "' + UsDateTime(LaTob.GetDateTime('DATEDEB')) +
                  '" AND "' + UsDateTime(LaTob.GetDateTime('DATEFIN')) + '" AND ' + Where + WhereConf +
                  ' ORDER BY TE_NUMEROPIECE, TE_GENERAL';

  {Théoriquement le résultat de l requête ne renvoie qu'une même devise}
  TobGrid.LoadDetailDBFromSQL('1', RequeteSQL);
  if TobGrid.Detail.Count > 0 then
    S1 := TobGrid.Detail[0].GetString('TE_DEVISE');

  TobGrid.PutGridDetail(Grid, False, False, '', True);
  {20/04/05 : Dessin de la grille}
  PrepareGrille(S1);
  {20/04/05 : Gestion des boutons}
  GestionBoutons();

  ADDMenuPop(PopupMenu, '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.SetAppelant(Value : string);
{---------------------------------------------------------------------------------------}
var
  ap : TDescendant;
begin
  if Value = '' then Value := TobArgument.GetString('APPEL');

  ap := dsDetail;
  repeat
    if Value = IntToStr(Ord(ap)) then begin
      FAppelant := ap;
      Exit;
    end;
    ap := Succ(ap);
  until ap = dsOPCVM;
end;

{---------------------------------------------------------------------------------------}
function TOF_DETAILFLUX.SetFVariables : Boolean;
{---------------------------------------------------------------------------------------}
begin
  {Récupération de la date sur laquelle doit se faire la requête}
  FDateOpe := TobArgument.GetString('DATEOPE');
  {31/10/06 : FQ 10338 : affichage de la date de rappro pour le suivi bancaire "prévisionnel"}
//  if (FDateOpe = 'TE_DATERAPPRO') and (FDateOpe = TobArgument.GetString('DATEAFF')) then
    {Remarque : le suivi bancaire est toujours en date de valeur}
  //  FDateOpe := 'TE_DATEVALEUR';                                 
  {Récupération du compte général sur lequel doit se faire la requête}
  FGeneral := TobArgument.GetString('GENERAL');
  {Récupération de la date de début, qui doit toujours être renseignée}
  FDateDeb := TobArgument.GetDateTime('DATEDEB');
  {Récupération de la date de fin, qui doit être renseignée si on n'est pas en affichage quotidien}
  FDateFin := TobArgument.GetDateTime('DATEFIN');
  {Récupération de la clause where sur TE_NATURE}
  FNature := TobArgument.GetString('NATURE');

  {On teste les Variables}
  Result := (FDateOpe <> '') and (FGeneral <> '') and
            (FDateDeb > iDate1900) and ((FDateFin > iDate1900) or (FPeriode = tp_1));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.SetPeriode(Value : string);
{---------------------------------------------------------------------------------------}
var
  ap : TypePeriod;
begin
  ap := tp_1;
  repeat
    if Value = IntToStr(Ord(ap)) then begin
      FPeriode := ap;
      Exit;
    end;
    ap := Succ(ap);
  until ap = tp_30;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.InitVariable;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRCONF}
  WhereConf := TObjConfidentialite.GetWhereConf(V_PGI.User, tyc_Banque);
  if WhereConf <> '' then WhereConf := ' AND (' + WhereConf + ') ';
  {$ENDIF TRCONF}
  
  TREcritureModified := '';
  FDateModif := iDate2099;
  {Tob d'affichage}
  TobGrid := TOB.Create('', nil, -1);
  {Tob pour le processus d'intégration en comptabilité}
  TobGen := TobPieceCompta.Create('§ECRIT', nil, -1);
  {Tob contenant les paramêtres pour la requête et l'affichage}
  TobArgument := TFVierge(Ecran).LaTOF.LaTOB;
end;

{23/03/06 : Ajout du champ NODOSSIER comme dernier champ des requêtes
{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.PrepareRequete;
{---------------------------------------------------------------------------------------}
var
  Where  : string;
  Prefix : string;
  Select : string;
  Clef   : string;
  SDateAff : string;
begin
  {1/ PRÉPARATION DE LA REQUÊTE}
  if FAppelant = dsControle then begin
    PrepareRequeteCtrl;
    Exit;
  end;

  if FAppelant = dsCommission then Prefix := 'T2.';

  Where  := Prefix + FDateOpe + ' >= "' + UsDateTime(TobArgument.GetDateTime('DATEDEB')) + '" AND ' +
            Prefix + FDateOpe + ' <= "' + UsDateTime(TobArgument.GetDateTime('DATEFIN')) + '" ';

  {15/11/06 : Gestion du filtre sur les dossiers pour le suivi Tréso et les budgets}
  Where := Where + GetFiltreDossier;

  {30/07/04 : Suivi des commissions}
  if FAppelant = dsCommission then begin
    {Pour les commissions de mouvement, la requête est indépendante des codes flux}
    if TobArgument.GetString('LIBNAT') <> suc_Divers then
      Where  := Where + 'AND T2.TE_CODEFLUX = "' + TobArgument.GetString('CODE') + '" ';

    if FGeneral <> '' then
      Where  := Where + 'AND T2.TE_GENERAL = "' + TobArgument.GetString('GENERAL') + '" ';

    {Pour la base, la requête est plus spécifique !!!}
    if TobArgument.GetString('LIBNAT') = suc_Commission then
      Where  := Where + 'AND T2.TE_COMMISSION = "' + TobArgument.GetString('LIBNAT') + '"';

    Where := Where + FNature;
    Select := 'T2.TE_GENERAL';
  end

  {Budget mensuel ou périodique}
  else if FAppelant = dsMensuel then begin
    Select := 'TE_GENERAL';
    if TobArgument.GetString('CODE') <> '' then
      Where  := Where + 'AND TE_CODEFLUX = "' + TobArgument.GetString('CODE') + '" ';
    {29/09/05 : FQ 10296 : General etait en doublon avec FGeneral}
    if FGeneral <> '' then
      Where  := Where + 'AND TE_GENERAL = "' + TobArgument.GetString('GENERAL') + '" ';
    Where := Where + FNature;
  end

  {Suivi par flux}
  else if FAppelant = dsDetail then begin
    Select := 'TE_GENERAL';
    Where  := Where + ' AND TE_CODEFLUX = "' + TobArgument.GetString('CODE') +
             '" AND TE_GENERAL = "' + TobArgument.GetString('GENERAL') + '"' + FNature;
  end

  {Suivi bancaire}
  else if FAppelant = dsBancaire then begin
    Select := 'TE_CODEFLUX';
    {05/04/05 : FQ 10240 : Si on demande les écritures pointées, on affiche le rappro du jour
     31/10/06 : FQ 10338 : modification des écritures affichées}
    if TobArgument.GetString('TYPE') = na_Prevision then begin
      Where := 'TE_GENERAL = "' + TobArgument.GetString('GENERAL') + '" ' + FNature;
      Where := Where + ' AND ((TE_DATEVALEUR > "' + UsDateTime(TobArgument.GetDateTime('DATEFIN')) + '" AND ' +
                             'TE_DATERAPPRO BETWEEN "' + UsDateTime(iDate1900 + 1) + '" AND "' + UsDateTime(TobArgument.GetDateTime('DATEDEB')) + '")';
      Where := Where + ' OR (TE_DATEVALEUR <= "' + UsDateTime(TobArgument.GetDateTime('DATEFIN')) + '" AND ' +
                            'TE_DATERAPPRO = "' + UsDateTime(iDate1900) + '")';
      Where := Where + ' OR (TE_DATEVALEUR <= "' + UsDateTime(TobArgument.GetDateTime('DATEDEB')) + '" AND ' +
                            'TE_DATERAPPRO > "' + UsDateTime(TobArgument.GetDateTime('DATEFIN')) + '"))';
    end
    else
      Where  := Where + ' AND TE_GENERAL = "' + TobArgument.GetString('GENERAL') + '" ' + FNature;
//    else if TobArgument.GetString('TYPE') = na_Pointe then
  //    Where := Where + ' AND TE_DATERAPPRO > "' + UsDateTime(TobArgument.GetDateTime('DATEDEB')) + '" AND TE_DATERAPPRO <= "' +
    //           UsDateTime(TobArgument.GetDateTime('DATEFIN')) + '"';
  end

  {Suivi par solde}
  else if FAppelant = dsSolde then begin
    Select := 'TE_CODEFLUX';
    {10/06/07 : FQ 10432 : gestion des échelles d'intérêts}
    if EchelleOk then Where := Where + ' ' + TobArgument.GetString('WHERE')
                 else Where := Where + ' AND TE_GENERAL = "' + TobArgument.GetString('GENERAL') + '"' + FNature;
    Where := Where + FNature;
  end;

  {15/04/05 : FQ 10240 : Gestion de la date d'affichage}
  if FDateOpe = 'TE_DATERAPPRO' then begin
    {Dans ce cas on inverse : la date d'affichage est la date d'opération de la fiche appelante}
    if TobArgument.GetString('DATEAFF') = 'TE_DATECOMPTABLE' then SDateAff := 'TE_DATECOMPTABLE AS TE_DATEAFF, '
                                                             else SDateAff := 'TE_DATEVALEUR AS TE_DATEAFF, ';
  end
  else if FDateOpe = 'TE_DATEVALEUR' then begin
    {31/10/06 : FQ 10338 : affichage de la date de rappro pour le "prévisionnel"}
    if TobArgument.GetString('DATEAFF') = 'TE_DATERAPPRO' then SDateAff := 'TE_DATERAPPRO AS TE_DATEAFF, '
                                                          {29/06/04 : Ajout de la date complémentaire dans la grille}
                                                          else SDateAff := 'TE_DATECOMPTABLE AS TE_DATEAFF, ';
    Clef := ', TE_CLEVALEUR';
  end
  else
    SDateAff := 'TE_DATEVALEUR AS TE_DATEAFF, '; {29/06/04 : Ajout de la date complémentaire dans la grille}

  {2/ REQUÊTE PROPREMENT DITE}
  if FAppelant = dsCommission then begin
    {Commissions : récupération des écritures de commission (TE_COMMISSION = "C") selon les critères passés en paramètre
                   dans le OnArgument de la fiche}
    if TobArgument.GetString('LIBNAT') = suc_Commission then
      {03/10/05 : FQ 10277 : Pour les commissions, on affiche les codes flux au lieu de T2.TE_GENERAL}
      RequeteSQL := 'SELECT T2.TE_DATEVALEUR, T2.TE_NATURE, T2.TE_CODEFLUX AS TE_GENE, T2.TE_NUMEROPIECE, T2.TE_NUMTRANSAC, T2.TE_CODECIB, T2.TE_LIBELLE,' +
             'T2.TE_DEVISE, T2.TE_MONTANTDEV, T2.TE_MONTANT, T2.TE_NUMLIGNE, T2.TE_CLEVALEUR, T2.TE_NODOSSIER FROM TRECRITURE T2 '+
             'LEFT JOIN BANQUECP on BQ_CODE = TE_GENERAL ' +
             ' WHERE ' + Where + WhereConf +
             ' ORDER BY T2.TE_NUMEROPIECE,' + Select

    {Cas particulier des commissions de mouvement}
    else if TobArgument.GetString('LIBNAT') = suc_Divers then begin
      {03/10/05 : FQ 10277 : Pour les commissions, on affiche les codes flux au lieu de T2.TE_GENERAL}
      RequeteSQL := 'SELECT T2.TE_DATEVALEUR, T2.TE_NATURE, T2.TE_CODEFLUX AS TE_GENE, T2.TE_NUMEROPIECE, T2.TE_NUMTRANSAC AS NUM, T2.TE_CODECIB, T2.TE_LIBELLE,' +
             'T2.TE_DEVISE, T2.TE_MONTANTDEV, T2.TE_MONTANT, T2.TE_NUMLIGNE, T2.TE_CLEVALEUR, T2.TE_NODOSSIER FROM TRECRITURE T2, FLUXTRESO, BANQUECP WHERE ' +
             'TE_CODEFLUX = TFT_FLUX AND TFT_COMMOUVEMENT = "X" AND TE_GENERAL = BQ_CODE AND ' + Where + WhereConf;
      RequeteSQL := RequeteSQL + ' UNION ALL ' +
      {03/10/05 : FQ 10277 : Pour les commissions, on affiche les codes flux au lieu de T2.TE_GENERAL}
             'SELECT T2.TE_DATEVALEUR, T2.TE_NATURE, T2.TE_CODEFLUX AS TE_GENE, T2.TE_NUMEROPIECE, T2.TE_NUMTRANSAC AS NUM, T2.TE_CODECIB, T2.TE_LIBELLE,' +
             'T2.TE_DEVISE, T2.TE_MONTANTDEV, T2.TE_MONTANT, T2.TE_NUMLIGNE, T2.TE_CLEVALEUR, T2.TE_NODOSSIER FROM TRECRITURE T2, RUBRIQUE, BANQUECP WHERE ' +
             'TE_CODEFLUX = RB_RUBRIQUE AND RB_CATEGORIE = "X" AND TE_GENERAL = BQ_CODE AND ' + Where + WhereConf +
             ' ORDER BY NUM, TE_GENE' {Remarque : sous oracle, le order by ne supporte que des champs aliassés sur les union}
    end
    else
    {Base : Les critères passés en paramètre dans le OnArgument de la fiche sont les mêmes que ci-dessus sauf TE_COMMISSION
            qui est à "A". On fait une jointure sur la table écriture elle-même, où T1 représente la base et T2 les commissions.
            On récupère donc les ecritures de "Base" qui correpondent aux commissions concernées par les critères de sélection.
            Rappel : dans la fiche de suivi des commissions, le test sur les dates se fait sur les commissions et on ramène la base
                     qui correspond aux dites commissions}
      {10/10/05 : FQ 10277 : Pour les commissions, on affiche les codes flux au lieu de T2.TE_GENERAL.

        !!!!!!!!!!!!!!!! A VOIR A L'OCCASION D'UN BUG !!!!!!!!!!!!!!!!

        REMARQUE !!! Il y a peut-être un problème avec la jointure sur le numéro de pièce : les com et les bases
                     ont-elles le même numéro de pièce ??

        !!!!!!!!!!!!!!!! A VOIR A L'OCCASION D'UN BUG !!!!!!!!!!!!!!!!
        30/11/05 : FQ 10303 : Effectivement la jointure sur TE_NUMEROPIECE est de trop !!
                   AND T1.TE_NUMEROPIECE = T2.TE_NUMEROPIECE}

      RequeteSQL := 'SELECT T1.TE_DATEVALEUR, T1.TE_NATURE, T1.TE_CODEFLUX AS TE_GENE, T1.TE_NUMEROPIECE, T1.TE_NUMTRANSAC, T1.TE_CODECIB, T1.TE_LIBELLE,' +
             'T1.TE_DEVISE, T1.TE_MONTANTDEV, T1.TE_MONTANT, T1.TE_NUMLIGNE, T1.TE_CLEVALEUR, T1.TE_NODOSSIER FROM TRECRITURE T1, TRECRITURE T2, BANQUECP ' +
             'WHERE T1.TE_GENERAL = BQ_CODE AND T1.TE_NUMTRANSAC = T2.TE_NUMTRANSAC AND T1.TE_COMMISSION = "' +
             suc_AvecCom + '" AND T2.TE_COMMISSION = "' + suc_Commission + '" AND ' + Where + WhereConf +
             ' ORDER BY T1.TE_NUMEROPIECE, T1.TE_GENERAL';
  end
  else if FAppelant = dsMensuel then
    RequeteSQL := 'SELECT ' + SDateAff + 'TE_NATURE,' + Select + ' AS TE_GENE, TE_NUMEROPIECE, TE_NUMTRANSAC, TE_CODECIB, TE_LIBELLE,' +
           'TE_DEVISE, TE_MONTANTDEV, TE_MONTANT, TE_NUMLIGNE' + Clef + ', TE_NODOSSIER FROM TRECRITURE ' +
           'LEFT JOIN BANQUECP on BQ_CODE = TE_GENERAL ' +
           'WHERE ' + Where + WhereConf + ' ORDER BY TE_NUMEROPIECE,' + Select
  {15/04/05 : FQ 10240 : Pour le détail des écritures non pointées du suivi bancaire on affiche
             les écritures qui ne sont pas encore pointées en "sDate" : d'où le <=}
  else if (FAppelant = dsBancaire) and (TobArgument.GetString('TYPE') = na_Prevision) then begin
    RequeteSQL:= 'SELECT ' + SDateAff + 'TE_NATURE,' + Select + ' AS TE_GENE, TE_NUMEROPIECE, TE_NUMTRANSAC, TE_CODECIB, TE_LIBELLE,' +
           'TE_DEVISE, TE_MONTANTDEV, TE_MONTANT, TE_NUMLIGNE' + Clef + ', TE_NODOSSIER FROM TRECRITURE '+
           'WHERE ' + Where +
           ' ORDER BY TE_NUMEROPIECE,' + Select;
  end
  else
    RequeteSQL := 'SELECT ' + SDateAff + 'TE_NATURE,' + Select + ' AS TE_GENE, TE_NUMEROPIECE, TE_NUMTRANSAC, TE_CODECIB, TE_LIBELLE,' +
           'TE_DEVISE, TE_MONTANTDEV, TE_MONTANT, TE_NUMLIGNE' + Clef + ', TE_NODOSSIER FROM TRECRITURE ' +
           TobArgument.GetString('LEFTJOIN') + ' WHERE ' + Where +
           ' ORDER BY TE_NUMEROPIECE,' + Select;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_DETAILFLUX.PrepareRequeteCtrl;
{---------------------------------------------------------------------------------------}
var
  DW : string;
begin
  if LaTob.GetString('TYPE') = 'C' then begin
    RequeteSQL := 'SELECT E_DATEVALEUR, E_JOURNAL, E_GENERAL, E_NUMEROPIECE, E_NUMLIGNE, E_MODEPAIE, E_LIBELLE, ' +
                  'E_DEVISE, E_DEBITDEV, E_CREDITDEV, E_NUMECHE, E_EXERCICE, "" FROM ' +
                  GetTableDossier(LaTOB.GetString('NOMBASE') ,'ECRITURE') + ' WHERE ';
    RequeteSQL := RequeteSQL + 'E_DATECOMPTABLE = "' + UsDateTime(LaTOB.GetDateTime('DATEDEB')) + '" AND ' +
                  'E_GENERAL = "' + LaTOB.GetString('GENERAL') + '" AND E_QUALIFPIECE = "N"';
  end

  else if LaTob.GetString('TYPE') = 'B' then begin
    {09/10/07 : FQ 10522 : on travaille sur eexblig et non plus sur eexbq}
    RequeteSQL := 'SELECT CEL_DATEOPERATION, CEL_REFPOINTAGE, CEL_GENERAL, CEL_NUMRELEVE, CEL_LIBELLE1, CEL_CODEAFB, CEL_LIBELLE, ' +
                  'IIF(CEL_DEBITDEV-CEL_CREDITDEV = 0, "' + V_PGI.DevisePivot + '", "DEV"), ' +
                  'IIF(CEL_DEBITDEV-CEL_CREDITDEV = 0, CEL_DEBITEURO, CEL_DEBITDEV), ';
    RequeteSQL := RequeteSQL + 'IIF(CEL_DEBITDEV-CEL_CREDITDEV = 0, CEL_CREDITEURO, CEL_CREDITDEV), ' +
                  'CEL_NUMLIGNE, "", "" FROM ' +
                  GetTableDossier(LaTOB.GetString('NOMBASE') ,'EEXBQLIG') +
                  ' WHERE CEL_DATEVALEUR = "' + UsDateTime(LaTOB.GetDateTime('DATEDEB')) + '" AND ';
    if EstPointageSurTreso then
      RequeteSQL := RequeteSQL + 'CEL_GENERAL = "' + LaTOB.GetString('CODE') + '" '
    else
      RequeteSQL := RequeteSQL + 'CEL_GENERAL = "' + LaTOB.GetString('GENERAL') + '" ';
  end
  else if LaTob.GetString('TYPE') = 'T' then begin
    if LaTob.GetString('DATEAFF') = 'TE_DATECOMPTABLE' then DW := 'TE_DATEVALEUR'
                                                       else DW := 'TE_DATECOMPTABLE';
    RequeteSQL := 'SELECT ' + LaTob.GetString('DATEAFF') + ', TE_NATURE, TE_GENERAL, TE_NUMEROPIECE, TE_NUMTRANSAC, TE_CODECIB, TE_LIBELLE, TE_DEVISE, ' +
                  'TE_MONTANTDEV, TE_MONTANT, TE_NUMLIGNE, "", TE_NODOSSIER FROM TRECRITURE WHERE ';
    RequeteSQL := RequeteSQL + DW + ' = "' + UsDateTime(LaTOB.GetDateTime('DATEDEB')) + '" AND ' +
                  'TE_GENERAL = "' + LaTOB.GetString('CODE') + '" AND TE_NATURE = "' + na_Realise + '"';
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_DETAILFLUX.GetFiltreDossier : string;
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  Result := '';
  s := LaTob.GetString('REGROUP');
  if (s <> '') and (Pos('<<', s) = 0 ) then begin
    if s[Length(s)] <> ';' then s := s + ';';
    Result := ' AND TE_NODOSSIER IN (' + GetClauseIn(s) + ') ';
  end;
end;

initialization
  RegisterClasses([TOF_DETAILFLUX]);

end.
