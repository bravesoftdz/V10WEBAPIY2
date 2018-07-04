{-------------------------------------------------------------------------------------
    Version   |   Date | Qui | Commentaires
--------------------------------------------------------------------------------------
 6.20.001.001  03/01/04  JP   Création de l'unité
 6.50.001.003  01/06/05  JP   FQ 10267 : La requête dans le OnArgument était incompatible avec Oracle
 7.00.001.001  12/01/06  JP   FQ 10323 : Correction de la gestion de la TVA
 7.05.001.001  18/09/06  JP   Mise en place de l'intégration multi-sociétés
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
 8.10.001.012  28/09/07  JP   FQ 10527 : problème de traduction de la requête UNION ALL
--------------------------------------------------------------------------------------}
unit TRDETAILOPCVM_TOF;

interface
    
uses
  Controls, Classes, Commun, Constantes, HMsgBox, 
  {$IFDEF EAGLCLIENT}
  MaineAGL, UtileAgl,
  {$ELSE}
  FE_Main, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} EdtREtat,
  {$ENDIF}
  Forms, SysUtils, HCtrls, HEnt1, UTOF, Grids, UTOB, Menus,
  UObjGen, HStatus, ULibPieceCompta;

type
  TOF_TRDETAILOPCVM = class (TOF)
    procedure OnArgument (S : string); override;
    procedure OnClose                ; override;
  protected
    PopupMenu : TPopUpMenu;
    Grid      : THGrid;
    Nature    : string; {Flux Réalisés, Prévisionnels ou tous}
    sDate     : string;
    FinMois   : string;
    sDateOpe  : string;

    General   : string; {Compte général sur lequel on est}
    Regle     : TReglesInteg;
    TobGrid   : TOB;
    TobGen    : TobPieceCompta;
    {Liste pour le recalcul des soldes}
    lSoldes   : TStringList;
    {Pour le calcul des contrevaleurs}
    ObjDevise : TObjDevise;
    {Pour l'intégration en compta}
    ObjTva    : TObjTVA;

    procedure GestionBtnPopup;
    procedure GestionGrille;

    function  TesteSelection                  : Boolean;
    function  IsLigneVente    (Row : Integer) : Boolean;
    procedure FinirTraitements;

    procedure MenuOnPopup           (Sender : TObject);
    procedure ValideBOClick         (Sender : TObject);
    procedure DeValideBOClick       (Sender : TObject);
    procedure IntegreClick          (Sender : TObject);
    procedure BDeleteClick          (Sender : TObject);
    procedure ToutSelectionner      (Sender : TObject);
    procedure ToutDeSelectionner    (Sender : TObject);
    procedure GrilleDbleClick       (Sender : TObject);
    procedure BImprimerClick        (Sender : TObject);
    {JP 20/11/03 : pour interdir la sélection des écritures réalisées}
    procedure ListeOnBeforeFlip     (Sender : TObject ; ARow : Integer ; Var Cancel : Boolean);

    {JP 28/08/03 : Intégration en comptabilité}
    procedure MajApresInteg;
    procedure Integration(TE_NUMTRANSAC : string);
  end ;

function TRLanceFiche_TRDETAILOPCVM(Dom, Fiche, Range, Lequel, Arguments: string): string;

implementation

uses
  {$IFDEF TRCONF}
  ULIbConfidentialite,
  {$ENDIF TRCONF}
  HTB97, UProcCommission, UProcEcriture, Messages,
  TROPCVM_TOM, TRVENTEOPCVM_TOM, AGLInit, UProcGen, UProcSolde, UObjOPCVM, Windows;

const
  COL_DATE     = 0;
  COL_NATURE   = 1;
  COL_PORTEF   = 2;
  COL_OPCVM    = 3;
  COL_COMPTE   = 4;
  COL_TRANSAC  = 5;
  COL_LIBELLE  = 6;
  COL_ACHAT    = 7;
  COL_VENTE    = 8;
  COL_NUMVENTE = 9;
  COL_COMPTA   = 10;
  COL_VALBO    = 11;

{---------------------------------------------------------------------------------------}
function TRLanceFiche_TRDETAILOPCVM(Dom, Fiche, Range, Lequel, Arguments: string): string;
{---------------------------------------------------------------------------------------}
begin
  Result := AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRDETAILOPCVM.OnArgument (S : string ) ;
{---------------------------------------------------------------------------------------}
var
  SQL,
  DateDeb,
  DateFin,
  Portef,
  Opcvm,
  DateA,
  DateV,
  NatureA,
  whConf,
  NatureV : string;
begin
  inherited ;
  Ecran.HelpContext := 150;
  {Gestion des popups et des Boutons}
  GestionBtnPopup;

  {DateDeb ; DateFin ; Portefeuille ; OPCVM ; NatureA ; NatureV ; Libellé}
  DateDeb := ReadTokenSt(S);
  DateFin := ReadTokenSt(S);
  Portef  := ReadTokenSt(S);
  Opcvm   := ReadTokenSt(S);
  NatureA := ReadTokenSt(S);
  NatureV := ReadTokenSt(S);
  SQL     := ReadTokenSt(S);
  Ecran.Caption := 'Liste des opérations sur le OPCVM du ' + SQL;
  UpdateCaption(Ecran);

  TobGrid := TOB.Create('', nil, -1);

  {On vient d'un affichage quotidien}
  if DateDeb = DateFin then begin
    DateA := ' AND TOP_DATEACHAT = "' + UsDateTime(StrToDate(DateDeb)) + '"';
    DateV := ' AND TVE_DATEVENTE = "' + UsDateTime(StrToDate(DateDeb)) + '"';
  end
  {On vient d'un affichage mensuel}
  else begin
    DateA := ' AND TOP_DATEACHAT BETWEEN "' + UsDateTime(StrToDate(DateDeb)) + '" AND "' + UsDateTime(StrToDate(DateFin)) + '"';
    DateV := ' AND TVE_DATEVENTE BETWEEN "' + UsDateTime(StrToDate(DateDeb)) + '" AND "' + UsDateTime(StrToDate(DateFin)) + '"';
  end;
  {$IFDEF TRCONF}
  whConf := TObjConfidentialite.GetWhereConf(V_PGI.User, tyc_Banque);
  if whConf <> '' then whConf := ' AND (' + whConf + ') ';
  {$ENDIF TRCONF}
  DateA := DateA + whConf;
  DateV := DateV + whConf;

  {JP 01/06/05 : FQ 10267 : Correction de la requête, incompatible avec Oracle}
  SQL := 'SELECT TOP_DATEACHAT AS DATEOPE, "A" AS NATURE, TOP_PORTEFEUILLE AS PORTEFEUILLE, TOP_CODEOPCVM AS OPCVM,';
  SQL := SQL + ' TOP_GENERAL AS GENERAL, TOP_NUMOPCVM AS NUMERO, TOP_LIBELLE AS LIBELLE, TOP_MONTANTACH AS ACHAT,';
  SQL := SQL + ' 0 AS VENTE, -1 AS NUMVENTE, TOP_COMPTA AS COMPTA, TOP_VALBO AS VALBO FROM TROPCVM';
  SQL := SQL + ' LEFT JOIN BANQUECP ON BQ_CODE = TOP_GENERAL ';
  SQL := SQL + ' WHERE TOP_PORTEFEUILLE = "' + Portef + '" AND TOP_CODEOPCVM = "' + Opcvm + '" ' + NatureA + DateA;

  {28/09/07 : FQ 10527 : problème de traduction : UNION ALL doit mis sous la forme ##UNION##
  SQL := SQL + ' UNION ALL SELECT TVE_DATEVENTE AS DATEOPE, "A" AS NATURE, TVE_PORTEFEUILLE AS PORTEFEUILLE,';}
  SQL := SQL + ' ##UNION## SELECT TVE_DATEVENTE AS DATEOPE, "A" AS NATURE, TVE_PORTEFEUILLE AS PORTEFEUILLE,';
  SQL := SQL + ' TVE_CODEOPCVM AS OPCVM , TVE_GENERAL AS GENERAL, "VENTE "||CAST(TVE_NUMVENTE AS VARCHAR(5)) AS NUMERO,';
  SQL := SQL + ' TVE_LIBELLE AS LIBELLE, 0 AS ACHAT, (TVE_NBVENDUE * TVE_COURSEUR) AS VENTE, ';
  SQL := SQL + ' TVE_NUMVENTE AS NUMVENTE, TVE_COMPTA AS COMPTA, TVE_VALBO AS VALBO FROM TRVENTEOPCVM';
  SQL := SQL + ' LEFT JOIN BANQUECP ON BQ_CODE = TVE_GENERAL ';
  SQL := SQL + ' WHERE TVE_PORTEFEUILLE = "' + Portef + '" AND TVE_CODEOPCVM = "' + Opcvm + '" ' + NatureV + DateV;

  TobGrid.LoadDetailFromSQL(SQL);

  {Ne devrait pas arrivé et cependant si cerains bug trainent encore dans TRSUIVIOPCM_TOF,
   cela peut arriver : on cherche à limiter les dégats}
  if TobGrid.Detail.Count = 0 then begin
    PostMessage(Ecran.Handle, WM_CLOSE, 0, 0);
    Exit;
  end;

  Grid := THGrid(GetControl('GRID'));
  {Présentation de la grille}
  GestionGrille;

  {Tob pour le processus d'intégration en comptabilité}
  TobGen := TobPieceCompta.Create('§ECRIT', nil, -1);
  {Création de l'objet devise pour la gestion des taux de change}
  ObjDevise := TObjDevise.Create(V_PGI.DateEntree);
  lSoldes   := TStringList.Create;
  lSoldes.Duplicates := dupIgnore;
  {Objet qui va permettre une éventuelle gestion de la TVA lors de l'intégration en comptabilité}
  ObjTva := TObjTVA.Create;
  {07/06/05 : Réinitialisation du conteneur d'erreurs}
  InitGestionErreur;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRDETAILOPCVM.GestionBtnPopup;
{---------------------------------------------------------------------------------------}
var
  Btn : TToolbarButton97;
begin
  {Gestion des popups}
  PopupMenu := TPopUpMenu(GetControl('POPUPMENU'));
  PopupMenu.OnPopup := MenuOnPopup;
  PopupMenu.Items[0].OnClick := ValideBOClick;
  PopupMenu.Items[1].OnClick := DeValideBOClick;
  PopupMenu.Items[3].OnClick := IntegreClick;
  PopupMenu.Items[5].OnClick := BDeleteClick;
  PopupMenu.Items[7].OnClick := ToutSelectionner;
  PopupMenu.Items[8].OnClick := ToutDeSelectionner;
  ADDMenuPop(PopupMenu, '', '');

  {Gestion des boutons}
  Btn := TToolbarButton97(GetControl('BVBO'));
  Btn.Hint := TraduireMemoire('Valider Back Office les lignes sélectionnées');
  Btn.OnClick := ValideBOClick;

  Btn := TToolbarButton97(GetControl('BDVBO'));
  Btn.Hint := TraduireMemoire('Dévalider Back Office les lignes sélectionnées');
  Btn.OnClick := DeValideBOClick;

  Btn := TToolbarButton97(GetControl('BREALISE'));
  Btn.Hint := TraduireMemoire('Intégrer en comptabilité les lignes sélectionnées');
  Btn.OnClick := IntegreClick;
  Btn.Enabled := AutoriseFonction(dac_Integration);

  Btn := TToolbarButton97(GetControl('BSELECT'));
  Btn.Hint := TraduireMemoire('Sélectionner toutes les lignes');
  Btn.OnClick := ToutSelectionner;

  Btn := TToolbarButton97(GetControl('BDESELECT'));
  Btn.Hint := TraduireMemoire('Désélectionner toutes les lignes');
  Btn.OnClick := ToutDeSelectionner;

  Btn := TToolbarButton97(GetControl('BDELETE'));
  Btn.Hint := TraduireMemoire('Suppression des lignes sélectionnées');
  Btn.OnClick := BDeleteClick;

  TToolbarButton97(GetControl('BIMPRIMER')).OnClick := BImprimerClick;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRDETAILOPCVM.GestionGrille;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  Grid.OnDblClick := GrilleDbleClick;
  Grid.OnBeforeFlip := ListeOnBeforeFlip;

  {Mise à jour du champ Nature : Réalisé si intégré en compta, Prévisionnel si Validé BO, sinon simulation}
  for n := 0 to TobGrid.Detail.Count - 1 do begin
    {Le cast a des effets délétaires}
    TobGrid.Detail[n].SetString('NUMERO', Trim(TobGrid.Detail[n].GetString('NUMERO')));
    if TobGrid.Detail[n].GetString('COMPTA') = 'X' then
      TobGrid.Detail[n].SetString('NATURE', 'R')
    else if TobGrid.Detail[n].GetString('VALBO') = 'X' then
      TobGrid.Detail[n].SetString('NATURE', 'P')
    else
      TobGrid.Detail[n].SetString('NATURE', 'S');
  end;

  TobGrid.PutGridDetail(Grid, False, False, '', True);

  Grid.ColFormats[COL_VENTE]   := StrFMask(V_PGI.OkDecV, '', True);
  Grid.ColFormats[COL_ACHAT]   := StrFMask(V_PGI.OkDecV, '', True);
  Grid.ColFormats[COL_NUMVENTE]:= '#';
  Grid.ColAligns[COL_DATE]     := taCenter;
  Grid.ColAligns[COL_ACHAT]    := taRightJustify;
  Grid.ColAligns[COL_VENTE]    := taRightJustify;
  Grid.ColAligns[COL_TRANSAC]  := taCenter;
  Grid.ColAligns[COL_NATURE]   := taCenter;
  Grid.ColAligns[COL_COMPTE]   := taCenter;
  Grid.ColAligns[COL_OPCVM]    := taCenter;
  Grid.ColAligns[COL_PORTEF]   := taCenter;
  Grid.ColWidths[COL_VALBO]    := -1;
  Grid.ColWidths[COL_COMPTA]   := -1;
  Grid.ColWidths[COL_NUMVENTE] := -1;
  Grid.ColLengths[COL_VALBO]   := -1;
  Grid.ColLengths[COL_COMPTA]  := -1;
  Grid.ColLengths[COL_NUMVENTE]:= -1;
//  Grid.UpdateTitres;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRDETAILOPCVM.OnClose ;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(TobGrid)   then FreeAndNil(TobGrid);
  if Assigned(lSoldes)   then LibereListe(lSoldes, True);
  if Assigned(ObjTva)    then FreeAndNil(ObjTva);
  if Assigned(ObjDevise) then FreeAndNil(ObjDevise);
  if Assigned(TobGen )   then FreeAndNil(TobGen);
  inherited;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRDETAILOPCVM.ValideBOClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  n     : Integer;
  T     : TOB;
  Valid : Boolean;
begin
  {On vérifie si il y a une sélection}
  if not TesteSelection then Exit;
  {Pour éviter un conseil du compilateur}
  Valid := True;

  if HShowMessage('1;' + Ecran.Caption + ';Êtes-vous sûr de vouloir valider les transactions sélectionnées.;Q;YNC;N;N;', '', '') = mrYes then begin
    InitMove(Grid.nbSelected, 'Validation de transactions');
    T := TOB.Create('TRECRITURE', nil, -1);
    try

      BeginTrans;
      try
        for n := Grid.RowCount - 1 downto 1 do begin
          Movecur(False);
          {On traite les lignes sélectionnées non intégrées en compta et non validée BO}
          if Grid.IsSelected(n) and (Grid.Cells[COL_COMPTA, n] <> 'X') and
                                    (Grid.Cells[COL_VALBO, n] <> 'X') then begin
            if IsLigneVente(n) then
              {Création des écritures et des commissionspour les lignes de la vente en cours.
               Maj des tables TRVENTEOPCVM, TROPCVM et de la liste des soldes}
              Valid := CreerEcritureVenteOPCVM(lSoldes, IntToStr(ValeurI(Grid.Cells[COL_NUMVENTE, n])), ObjDevise)

            else begin
              {Initialisation de l'écriture}
              InitNlleEcritureTob(T, Grid.Cells[COL_COMPTE, n], V_PGI.CodeSociete);
              {Reprise des champs de la transaction}
              T.SetDateTime('TE_DATECOMPTABLE', StrToDate(Grid.Cells[COL_DATE, n]));
              T.SetString('TE_LIBELLE'   , Grid.Cells[COL_LIBELLE, n]);
              T.SetString('TE_REFINTERNE', 'Valide par détail ' + Grid.Cells[COL_TRANSAC, n]);
              T.SetString('TE_NUMTRANSAC', Grid.Cells[COL_TRANSAC, n]);

              {Récupération des informations depuis la table FluxTreso : Le flux est celui paramétré comme
               étant celui de versement dans la table Transac}
              ChargeChpFluxTreso(T, Grid.Cells[COL_TRANSAC, n], VERSEMENT);

              {Les montants sont négatifs puisqu'il s'agit d'un achat}
              T.SetDouble('TE_MONTANT', -1 * Abs(Valeur(Grid.Cells[COL_ACHAT, n])));
              {Termine l'écriture, gère les commissions et l'insère dans la rable}
              Valid := TermineEcritureTob(T, ObjDevise, True);

              if Valid then begin
                ExecuteSQL('UPDATE TROPCVM SET TOP_VALBO = "X", TOP_USERVBO = "' + V_PGI.User + '", TOP_DATEVBO = "' +
                           UsDateTime(V_PGI.DateEntree) + '", TOP_USERMODIF = "' + V_PGI.User + '", TOP_DATEMODIF = "' +
                           UsDateTime(Now) + '" WHERE TOP_NUMOPCVM = "' + Grid.Cells[COL_TRANSAC, n] + '"');

                {Mise à jour de la liste pour le recalcul des soldes de la table TRECRITURE}
                AddGestionSoldes(lSoldes, Grid.Cells[COL_COMPTE, n], StrToDateTime(Grid.Cells[COL_Date, n]));
              end;
            end;
            Grid.Cells[COL_VALBO , n] := 'X';
            Grid.Cells[COL_NATURE, n] := na_Prevision;
          end;

          if not Valid then begin
            {22/03/05 : FQ 10223 : Nouvelle gestion des erreurs}
            AfficheMessageErreur;
            RollBackDiscret;
            Exit;
          end;
        end;
        CommitTrans;
      except
        on E : Exception do begin
          RollBack;
          PGIError(E.Message);
          {On sort pour ne pas exécuter FinirTraitements;}
          Exit;
        end;
      end;
      FinirTraitements;
    finally
      if Assigned(T) then FreeAndNil(T);
      FiniMove;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRDETAILOPCVM.IntegreClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  {On vérifie si il y a une sélection}
  if not TesteSelection then Exit;

  if HShowMessage('1;' + Ecran.Caption + ';Êtes-vous sûr de vouloir intégrer les transactions sélectionnées.;Q;YNC;N;N;', '', '') = mrYes then begin
    InitMove(Grid.nbSelected, 'Validation de transactions');
    BeginTrans;
    try
      {Préparation de la Tob "comptable"}
      for n := Grid.RowCount - 1 downto 1 do begin
        MoveCur(False);
        {On traite les lignes sélectionnées non intégrées en compta et validée BO}
        if Grid.IsSelected(n) and (Grid.Cells[COL_COMPTA, n] <> 'X') and
                                  (Grid.Cells[COL_VALBO, n] = 'X') then begin
          if IsLigneVente(n) then Integration(GetNumTransacVente(IntToStr(ValeurI(Grid.Cells[COL_NUMVENTE, n])), True))
                             else Integration(Grid.Cells[COL_TRANSAC, n]);
        end;
      end;{for n}

      {Si des lignes d'écritures comptables ont été générées ...}
      if TobGen.Detail.Count > 0 then begin
        {22/03/05 : FQ 10223 : Nouvelle gestion des erreurs}
        AfficheMessageErreur;

        {18/09/06 : Nouvelle fonction de traitement Multi-sociétés}
        if TRIntegrationPieces(TobGen, False) then
          MajApresInteg
      end;

      FiniMove;
      CommitTrans;
      FinirTraitements;
    except
      on E : Exception do begin
        RollBack;
        PGIError(E.Message);
      end;
    end;
  end;
end;

{Préparation de la TobCompta
{---------------------------------------------------------------------------------------}
procedure TOF_TRDETAILOPCVM.Integration(TE_NUMTRANSAC : string);
{---------------------------------------------------------------------------------------}
var
  tEcr  : TOB;
begin
  if TE_NUMTRANSAC = '' then Exit;

  tEcr := TOB.Create('ùùù', nil, -1);
  try
    tEcr.LoadDetailFromSQL('SELECT * FROM TRECRITURE WHERE TE_NUMTRANSAC = "' + TE_NUMTRANSAC + '"');
    {S'il n'y a pas d'écritures pour la transaction !!?}
    if tEcr.Detail.Count = 0 then Exit;
    {18/09/06 : Nouvelle intégration des écritures en comptabilité (gestion du multi-sociétés)}
    TRGenererPieceCompta(TobGen, tEcr);
  finally
    if Assigned(tEcr) then FreeAndNil(tEcr);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRDETAILOPCVM.DeValideBOClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  n     : Integer;
  MsgVa : string;
begin
  {On vérifie si il y a une sélection}
  if not TesteSelection then Exit;

  if HShowMessage('1;' + Ecran.Caption + ';Êtes-vous sûr de vouloir dévalider les transactions sélectionnées.;Q;YNC;N;N;', '', '') = mrYes then begin
    InitMove(Grid.nbSelected, 'Dévalidation de transactions');

    BeginTrans;
    try
      for n := Grid.RowCount - 1 downto 1 do begin
        Movecur(False);
        {On traite les lignes sélectionnées non intégrées en compta et validée BO}
        if Grid.IsSelected(n) and (Grid.Cells[COL_COMPTA, n] <> 'X') and
                                  (Grid.Cells[COL_VALBO, n] = 'X') then begin
          if IsLigneVente(n) then
            SupprEcritureVenteOPCVM(IntToStr(ValeurI(Grid.Cells[COL_NUMVENTE, n])))

          else begin
            {Suppression des écritures liées à la transaction : opcvm, frais, commissions}
            if SupprimePiece('', Grid.Cells[COL_TRANSAC, n], '', '') then
              {Mise à jour de la table TROPCVM}
              ExecuteSQL('UPDATE TROPCVM SET TOP_VALBO = "", TOP_USERVBO = "", TOP_DATEVBO = "' + UsDateTime(iDate1900) +
                         '", TOP_USERMODIF = "' + V_PGI.User + '", TOP_DATEMODIF = "' + UsDateTime(Now) +
                         '" WHERE TOP_NUMOPCVM = "' + Grid.Cells[COL_TRANSAC, n] + '"');
          end;
          {Mise à jour de la grille}
          Grid.Cells[COL_VALBO , n] := '-';
          Grid.Cells[COL_NATURE, n] := na_Simulation;
          {Mise à jour de la liste de recalcul des soldes}
          AddGestionSoldes(lSoldes, Grid.Cells[COL_COMPTE, n], StrToDateTime(Grid.Cells[COL_DATE, n]));
        end;
        CommitTrans;
      end;{for n := }
    except
      on E : Exception do begin
        RollBack;
        if MsgVa = '' then PGIError(E.Message);
        {On sort pour ne pas exécuter FinirTraitements;}
        Exit;
      end;
    end;
    FinirTraitements;
    FiniMove;
  end; {If HShowMessage}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRDETAILOPCVM.GrilleDbleClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
 if IsLigneVente(Grid.Row) then
   TRLanceFiche_TRVenteOPCVM('TR', 'TRFICVENTEOPCVM', IntToStr(ValeurI(Grid.Cells[COL_NUMVENTE, Grid.Row])), '', ActionToString(taConsult))
 else
   TRLanceFiche_OPCVM('TR', 'TROPCVM', '', Grid.Cells[COL_TRANSAC, Grid.Row], ActionToString(taConsult));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRDETAILOPCVM.MenuOnPopup(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  PopupMenu.Items[3].Enabled := AutoriseFonction(dac_Integration);
end;

{Imprime la grille rajoutée au lieu de la DB grille
{---------------------------------------------------------------------------------------}
procedure TOF_TRDETAILOPCVM.BImprimerClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  n : Integer;
begin
  {25/06/04 : pour gérer le cas de l'eAgl (pas de PrintDBGrid) et le titre sur deux lignes}
  if TobGrid.Detail.Count = 0 then Exit;
  T := TOB.Create('$DETAILOPCVM', nil, -1);
  try
    T.Dupliquer(TobGrid, True, True);
    T.Detail[0].AddChampSupValeur('TITRE', Ecran.Caption, True);
    for n := 0 to T.Detail.Count - 1 do begin
           if T.Detail[n].GetString('NATURE') = na_Prevision  then T.Detail[n].SetString('NATURE', 'Prévision')
      else if T.Detail[n].GetString('NATURE') = na_Simulation then T.Detail[n].SetString('NATURE', 'Simulation')
      else if T.Detail[n].GetString('NATURE') = na_Realise    then T.Detail[n].SetString('NATURE', 'Réalisé')
    end;
    LanceEtatTOB('E', 'ECT', 'OPC', T, True, False, False, nil, '', 'Détail des OPCVM', False);
  finally
    if Assigned(T) then FreeAndNil(T);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRDETAILOPCVM.ListeOnBeforeFlip(Sender : TObject ; ARow : Integer ; Var Cancel : Boolean);
{---------------------------------------------------------------------------------------}
begin
  {On n'autorise pas la sélection d'une écriture réalisée, car on ne peut la réaliser de
   nouveau ou lui changer de date de valeur}
  Cancel := (Grid.Cells[COL_NATURE, ARow] = na_Realise);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRDETAILOPCVM.ToutSelectionner(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  for n := 1 to Grid.RowCount - 1 do
    if not Grid.IsSelected(n) then Grid.FlipSelection(n);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRDETAILOPCVM.ToutDeSelectionner(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  Grid.ClearSelected;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRDETAILOPCVM.BDeleteClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  {On vérifie si il y a une sélection}
  if not TesteSelection then Exit;

  if HShowMessage('1;' + Ecran.Caption + ';Êtes-vous sûr de vouloir supprimer les transactions sélectionnées.;Q;YNC;N;N;', '', '') = mrYes then begin
    InitMove(Grid.nbSelected, 'Suppression de transactions');

    BeginTrans;
    try
      for n := Grid.RowCount - 1 downto 1 do begin
        Movecur(False);
        {On ne traite que les lignes sélectionnées qui ne sont pas réalisées ni importé de comptabilité.}
        if Grid.IsSelected(n) and (Grid.Cells[COL_NATURE, n] <> na_Realise) then begin
          if IsLigneVente(n) then
            {Mise à jour des OPCVM et suppression de la vente et des éventuelles écritures}
            SupprimerVente(IntToStr(ValeurI(Grid.Cells[COL_NUMVENTE, n])))
          else
            {Suppression de l'OPCVM et des écritures attachés à la transaction. On ne s'occupe pas des
             éventuelles ventes, car on ne peut pas vendre un OPCVM non intégré en Comptabilité}
            SupprimerOPCVM(Grid.Cells[COL_TRANSAC, n]);

          if (Grid.Cells[COL_NATURE, n] = na_Prevision) then
            {Mise à jour de la liste pour le recalcul des soldes de la table TRECRITURE}
            AddGestionSoldes(lSoldes, Grid.Cells[COL_COMPTE, n], StrToDateTime(Grid.Cells[COL_Date, n]));
          {Suppression de la ligne dans la grille}
          Grid.DeleteRow(n);
        end;
      end;
      CommitTrans;
    except
      on E : Exception do begin
        RollBack;
        PGIError(E.Message);
        {On sort pour ne pas exécuter FinirTraitements;}
        Exit;
      end;
    end;
    FiniMove;
    FinirTraitements;
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_TRDETAILOPCVM.IsLigneVente(Row : Integer) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := ValeurI(Grid.Cells[COL_NUMVENTE, Row]) <> -1;
end;

{Vérifie si au moins une ligne est sélectionnée
{---------------------------------------------------------------------------------------}
function TOF_TRDETAILOPCVM.TesteSelection : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := True;
  {Aucune sélection, on sort}
  if Grid.NbSelected = 0 then begin
    HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner une ligne.;W;O;O;O;', '', '');
    Result := False;
  end;
  {Vide la liste des comptes pour le recalcul des soldes}
  LibereListe(lSoldes, False);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRDETAILOPCVM.FinirTraitements;
{---------------------------------------------------------------------------------------}
begin
  MultiRecalculSolde(lSoldes);
  Grid.ClearSelected;
  TobGen.ClearDetailPC;
  AfficheMessageErreur;
end;

{Mise à jour des tables TRECRITURE et TROPCVM après l'intégration
{---------------------------------------------------------------------------------------}
procedure TOF_TRDETAILOPCVM.MajApresInteg;
{---------------------------------------------------------------------------------------}
var
  T   : TOB;
  n   : Integer;
  SQL : string;
  Cle : string;
  Ven : Boolean;
begin
  try
    {On Stocke le numéro de transaction dans la référence interne}
    TobGen.Detail.Sort('E_REFINTERNE');
    {On boucle sur les écritures comptables}
    for n := 0 to TobGen.Detail.Count - 1 do begin
      T := TobGen.Detail[n];
      {Si l'intégration s'est bien passée ...}
      if (T.GetString('RESULTAT') = 'OK') and (T.Detail.Count > 0) then begin
        {... on regarde si on a changé de transaction, sachant que pour une OPCVM on peut avoir plusieurs
         écritures de Trésorerie et que pour chacune de ces dernières on a deux ou trois écritures comptables}
        Cle := T.Detail[0].GetString('E_REFINTERNE');
        Cle := ReadTokenPipe(Cle, '-');
        Ven := Copy(Cle, 7, 3) = TRANSACVENTE;

        {Mise à jour des écritures intégrées en comptabilité}
        UpdatePieceStr('', Cle, '', '', 'TE_NATURE', na_Realise);
        if Ven then
          {Mise à jour de la table TROPCVM}
          SQL := 'UPDATE TRVENTEOPCVM SET TVE_USERCOMPTA = "' +  V_PGI.User + '", TVE_COMPTA = "X", ' +
                 'TVE_USERMODIF = "' + V_PGI.User + '", TVE_DATEMODIF = "' + UsDateTime(Now) +
                 '", TVE_DATECOMPTA = "' + USDateTime(v_PGI.DateEntree) + '" WHERE TVE_NUMVENTE = ' + GetNumTransacVente(Cle, False)
        else
          {Mise à jour de la table TROPCVM}
          SQL := 'UPDATE TROPCVM SET TOP_USERCOMPTA = "' +  V_PGI.User + '", TOP_COMPTA = "X", ' +
                 'TOP_USERMODIF = "' + V_PGI.User + '", TOP_DATEMODIF = "' + UsDateTime(Now) +
                 '", TOP_DATECOMPTA = "' + USDateTime(v_PGI.DateEntree) + '" WHERE TOP_NUMOPCVM = "' + Cle + '"';
          ExecuteSQL(SQL);
      end;
    end;

    {Mise à jour de la grille}
    for n := 1 to Grid.RowCount - 1 do
      if Grid.IsSelected(n) then begin
        Grid.Cells[COL_COMPTA, n] := 'X';
        Grid.Cells[COL_NATURE, n] := na_Realise;
      end;
  except
    raise;
  end;
end;



initialization
  RegisterClasses([TOF_TRDETAILOPCVM]);

end.
