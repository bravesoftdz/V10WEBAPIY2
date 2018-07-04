{ Unité : Source TOF de la FICHE : TRMULSOLDE
--------------------------------------------------------------------------------------
    Version    |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
    0.91        09/11/01   BT   Création de l'unité
                15/04/04   JP   Nouvelle gestion des natures (P, R, S)
                14/09/05   JP   FQ 10287 : Indice hors limite
 7.09.001.001   03/08/06   JP   Gestion du mutli sociétés
 7.09.001.001   19/10/06   JP   Filtre sur les comptes et les dossiers
 7.09.001.003   26/12/06   JP   FQ 10387 : Paramétrage de la date d'opération
 7.09.001.007   02/03/07   JP   FQ 10414 : le filtre sur les banques fonctionnent mal
 8.00.001.020   15/06/07   JP   FQ 10406 : Refonte de l'affichage des soldes car ils toujours pris
                                dans la base et ne tiennent pas compte du filtre sur la nature
 8.00.002.002   01/08/07   JP   Nouvelle gestion des soldes : on les recalcule à la volée et on ne les stocke plus
--------------------------------------------------------------------------------------}

unit TofEquilibrage ;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  {$IFDEF EAGLCLIENT}
  eMul, MaineAGL, UtileAgl,
  {$ELSE}
  FE_Main, Mul, DBGrids,
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, EdtREtat,
  {$ENDIF}
  Windows, StdCtrls, Controls, Classes, forms, sysutils, Graphics,
  Grids, HCtrls, HEnt1, UTOF, UTob, Menus, ExtCtrls, HTB97, HQry, UObjGen,
  UProcGen;

type
  PColor = class
    Drapeau : string ;
  end ;

  TOF_EQUILIBRAGE = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnDisplay                ; override ;
  protected
    Chargement: Boolean;
    PopupMenu : TPopUpMenu;
    QListe    : THQuery;
    {$IFDEF EAGLCLIENT}
    FListe    : THGrid;
    {$ELSE}
    FListe    : THDBGrid;
    {$ENDIF}
    Dossier : THMultiValComboBox;

    Indicator : array [0..2] of TBitmap;
    TobListe  : TOB;
    TobConditions: TOB;
    BAuto     : TToolBarButton97;
    BManuel   : TToolBarButton97;
    BanqueCpt : THValComboBox;
    ArrondiN  : THEdit;
    DeviseAff : string;
    TobVir    : TOB;
    {Compte de l'enregistrement pivot}
    FPivot    : string;
    {Objet contenant les taux de conversion}
    ObjTaux   : TObjDevise;
    {Taux de la devise d'affichage par rapport à l'euro}
    TauxAff   : Double;
    {Contient la sélection des natures en SQL}
    ClauseNat : string;
    lSoldes   : TStringList;
    {26/12/06 : FQ 10387 : l'utilisateur peut maintenant choisir la date d'équilibrage}
    FDateEqui : TDateTime;

    procedure OnPopUpMenu                (Sender : TObject);
    procedure PopupAjouterRetirerOnClick (Sender : TObject);
    procedure FListeOnDblClick           (Sender : TObject);
    procedure FListeOnFlipSelection      (Sender : TObject);
    procedure FListeOnColumnWidthsChanged(Sender : TObject);

    {$IFDEF EAGLCLIENT}
    procedure FListeKeyDown   (Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure CocheDecoche;
    procedure FListeOnDrawCell(Sender: TObject; ACol, ARow: Longint; ARect: TRect; AState: TGridDrawState);
    {$ELSE}
    procedure FListeOnDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
    {$ENDIF}

    procedure bImprimerClick       (Sender : TObject);
    procedure BAutoOnClick         (Sender : TObject);
    procedure BManuelOnClick       (Sender : TObject);
    procedure BVirOnClick          (Sender : TObject);
    procedure BAffConditionsOnClick(Sender : TObject);
    procedure NatureOnChange       (Sender : TObject);
    procedure DeviseOnChange       (Sender : TObject);
    procedure BanqueCptOnChange    (Sender : TObject);
    procedure DossierOnChange      (Sender : TObject);
    procedure DateEquiOnChange     (Sender : TObject); {26/12/06 : FQ 10387}

    {Charge les soldes pour pouvoir les afficher dans la grille}
    procedure ChargerListeSolde;
    procedure MajBVir;
    procedure CalculerVir;
    procedure RemplirTobListe;
    procedure SetPivot  (P : string);
    function  MajXXWhere : Boolean; {FQ 10414 : plus de paramètre à la fonction}
    function  EquilibrerCompte(Compte : string; Solde : Double): Double;
    procedure ChargeConditions;

  public
    {            A ENVISAGER DANS UNE VERSION ULTÉRIEURE DANS MAJSOLDE

       lorsque le traitement est fini, si le compte pivot est positif et que des comptes équilibrés sont dans la
       partie négative de la fouchette, il faudrait les rapprocher de zero tant que le compte pivot est positif}
    class procedure MajSolde(TobVir : TOB; Compte : string; Solde : Double);

    property  Pivot: string read FPivot write SetPivot;
  end ;

procedure TRLanceFiche_Equilibrage(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);


implementation

uses
  {$IFDEF TRCONF}
  ULibConfidentialite,
  {$ENDIF TRCONF}
  TofResultEqui, TofVirement, TomConditionEqui, Commun, AGLInit, Constantes, UtilPgi,
  ParamSoc, UTobDebug;

const
  {Voir liste TREQUITRECRITURE
   Remarque : lz demande d'afficher le libellé complet rajoute une colonne en eAGL :
   on a TE_NODOSSIER (invisible) en plus de la colonne affichant le libellé}
{$IFDEF EAGLCLIENT}
  COL_INDICATOR = 0;
  COL_LIBDOS    = 1;
  COL_DOSSIER   = 2;
  COL_DEVISE    = 3;
  COL_GENERAL   = 4;
  COL_LIBELLE   = 5;
  COL_DATEC     = 6;
  COL_SOLDEO    = 7;
  COL_DATEV     = 8;
  COL_SOLDEV    = 9;
{$ELSE}
  COL_INDICATOR = 0;
  COL_DOSSIER   = 1;
  COL_DEVISE    = 2;
  COL_GENERAL   = 3;
  COL_LIBELLE   = 4;
  COL_DATEC     = 5;
  COL_SOLDEO    = 6;
  COL_DATEV     = 7;
  COL_SOLDEV    = 8;
{$ENDIF EAGLCLIENT}


type
  TObjSolde = class
    dtVal : string;
    dtOpe : string;
    soVal : string;
    soOpe : string;
  end;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_Equilibrage(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{Affectation du solde à un compte dans la Tob et constitution du Champ titre SOURCE
{---------------------------------------------------------------------------------------}
procedure PutSource(aTob: TOB; Solde : Double);
{---------------------------------------------------------------------------------------}
begin
  {JP 07/05/04 : Gestion de trois champs de concaténation pour rendre le TobViewer (TofResultEqui) plus lisible}
  aTob.PutValue('SOURCE', Format('Émetteur : %s (solde : %.2n %s)',
                [aTob.GetValue('SLIBELLE'), Solde, aTob.GetValue('SDEVISE')]));
  aTob.PutValue('DESTINATION', Format('Destinataire : %s (solde : %.2n %s)',
                [aTob.GetValue('DLIBELLE'), aTob.GetDouble('DSOLDE'), aTob.GetValue('DDEVISE')]));
  aTob.PutValue('VIREMENT', Format('Virement d''un montant de  %.2n %s',
                [aTob.GetDouble('MONTANT'), aTob.GetValue('SDEVISE')]));
end;

{---------------------------------------------------------------------------------------}
class procedure TOF_EQUILIBRAGE.MajSolde(TobVir : TOB; Compte : string; Solde : Double);
{---------------------------------------------------------------------------------------}
var
  I    : Integer;
  aTob : TOB;
begin
  for I := 0 to TobVir.Detail.Count-1 do begin
    aTob := TobVir.Detail[I];
    if aTob.GetValue('SCOMPTE') = Compte then
    begin
      aTob.PutValue('SSOLDE', Solde); {Pas visible mais utilisé en modif de TobViewer}
      PutSource(aTob, Solde);
    end;

    if aTob.GetValue('DCOMPTE') = Compte then
      aTob.PutValue('DSOLDE', Solde);
  end;
end;

{Rend le montant à verser sur un compte pour l'équilibrer, c'est à dire le rendre
 compatible avec les règles d'équilibrage. Ce montant est négatif s'il faut faire un retrait.
{---------------------------------------------------------------------------------------}
function TOF_EQUILIBRAGE.EquilibrerCompte(Compte: string; Solde: Double): Double;
{---------------------------------------------------------------------------------------}
var
  aTob     : TOB;
  SoldeDeb,
  SoldeCred,
  CreditMin,
  DebitMin : Double;
  DevTmp   : string;
  Arrondi  : string;

   {Convertit les montants dans la devises d'affichage
   {------------------------------------------------------------}
   function Convertir(Name : string): Double;
   {------------------------------------------------------------}
   var
     ME : Double;
     MD : Double;
   begin
      {Récupération du montant à convertir}
      ME := Valeur(aTob.GetValue('TCE_' + Name));
      {On convertit ME en euros}
      ObjTaux.ConvertitMnt(ME, MD, DevTmp, Compte);
      {Puis on convertit ME dans la devise d'affichage}
      Result := ME / TauxAff;
   end;
begin
  {On commence par récupérer la devise du compte}
  aTob := TobListe.FindFirst(['TE_GENERAL'], [Compte], False);
  if aTob <> nil then DevTmp := aTob.GetValue('TE_DEVISE');

  {On recherche les conditions d'équilibrage du Compte en cours : c'est à dire la somme qu'il faut déposer
   ou retirer du compte en cours pour que son solde corresponde aux conditions d'équilibrage.

   Remarque : le résultat retourné est négatif s'il faut retirer de l'argent du compte et positif s'il faut
              déposer de l'argent sur le compte}

  aTob := TobConditions.FindFirst(['TCE_GENERAL'], [Compte], False);
  if aTob <> nil then begin
    Arrondi := aTob.GetValue('TCE_ARRONDI');
    CreditMin := Convertir('CREDITMIN');
    DebitMin := Convertir('DEBITMIN');

    {Sur ce compte, il y a une tolérance}
    if aTob.GetValue('TCE_TYPESOLDE') = 'X' then begin
      {Récupération des valeurs de la fourchette de tolérance}
      SoldeCred := Convertir('SOLDECRED');
      SoldeDeb := Convertir('SOLDEDEB');

      {La somme sur le compte est inférieure à la fourchette de tolérance}
      if Solde < SoldeDeb then begin
        {On va calculer le montant qu'il faut déposer sur ce compte}
        Result := ArrondirTreso(Arrondi, SoldeDeb - Solde);
        {Si le dépôt est inférieur au montant minimum d'un virement créditeur}
        if Result < CreditMin then begin
          {On déposera la somme minimum}
          Result := ArrondirTreso(Arrondi, CreditMin);
          {On s'assure cependant que le futur solde n'est pas supérieur à la fourchette ...}
          if Solde + Result > SoldeCred then
            {... le dépôt est impossible}
            Result := 0;
        end;
      end

      {La somme sur le compte est supérieure à la fourchette de tolérance}
      else if Solde > SoldeCred then  begin
        {On va calculer le montant qu'il faut retirer sur ce compte}
        Result := ArrondirTreso(Arrondi, SoldeCred - Solde);
        {Si le retrait est inférieur au montant minimum d'un virement débiteur}
        if Abs(Result) < Abs(DebitMin) then begin
          {On retirera la somme maximum}
          Result := ArrondirTreso(Arrondi, DebitMin);
          {Le risque est fort que l'on ait saisi un montant positif comme DebitMin}
          if Result > 0 then Result := - Result;
          {On s'assure cependant que le futur solde n'est pas inférieur à la fourchette ...}
          if Solde + Result < SoldeDeb then
            {... le dépôt est impossible}
            Result := 0;
        end;
      end

      {On est dans la fourchette de tolérance, il n'y a pas besoin d'équilibrage}
      else
        Result := 0;
    end

    {Il n'y a pas de fourchette mais un solde à conserver}
    else begin
      {On récupère le solde de référence}
      SoldeCred := Convertir('SOLDECONS');
      {Le solde correspond au solde de référence ...}
      if Solde = SoldeCred then
        {... pas d'équilibrage}
        Result := 0
      else begin
        Result := ArrondirTreso(Arrondi, SoldeCred - Solde);
        {On est en dessous du solde de référence}
        if Result > 0 then begin
          {On est en dessous du minimum d'un virement créditeur ...}
          if Result < CreditMin then
            Result := 0;
        end
        {On est au dessus du solde de référence}
        else
          {On est au dessus du maximum d'un virement débiteur ...}
          if Abs(Result) < DebitMin then {25/06/04 : Mise de la valeur absolue et inversion de l'inégalité FQ 10107}
            Result := 0;
      end;
    end;
  end

  {Si la TOB est vide !!!}
  else
    Result := 0;
end;

{Grise/active les boutons.
 Impossible de faire l'équilibrage auto si moins de deux enregistrement.
 Impossible de faire de nouveaux virements s'il en reste non validés.
{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.MajBVir;
{---------------------------------------------------------------------------------------}
var
  B : Boolean;
begin
  {Le bouton est actif dès qu'il y a des virements non générés (simulation ou non)}
  B := ExisteSql('SELECT TEQ_NUMEQUI FROM EQUILIBRAGE WHERE TEQ_FICEXPORT = "-"');
  SetControlEnabled('BVIR', B);

  B := (TobListe.Detail.Count > 1) and not B;
  BManuel.Enabled := B;
  BAuto.Enabled := B and (Pivot <> ''); // Auto que si existe un pivot
end;

{ Met à jour le bouton Equilibrage auto si le pivot change
{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.SetPivot(P: String);
{---------------------------------------------------------------------------------------}
begin
  FPivot := P;
  {L'équilibrage automatique n'est autorisé que s'il existe un pivot}
  BAuto.Enabled := BManuel.Enabled and (FPivot <> '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
var
  I  : Integer;
begin
  Chargement := True;
  Inherited ;
  Ecran.HelpContext := 50000123;
  // Verifier si les soldes sont à jour ?
  TobListe := TOB.Create('', Nil, -1);
  TobVir := TOB.Create('0', Nil, -1);
  TobConditions := TOB.Create('', Nil, -1);
  QListe := TFMul(Ecran).Q;
  lSoldes := TStringList.Create;

  PopupMenu := TPopUpMenu(GetControl('POPUPMENU'));
  PopupMenu.OnPopup := OnPopupMenu;
  PopupMenu.Items[0].OnClick := FListeOnDblClick; // Pivot
  PopupMenu.Items[1].OnClick := PopupAjouterRetirerOnClick;
  PopupMenu.Items[3].OnClick := BManuelOnClick;
  PopupMenu.Items[4].OnClick := BAutoOnClick;
  PopupMenu.Items[5].OnClick := BAffConditionsOnClick;

  for I := 0 to High(Indicator) do
    Indicator[I] := TImage(GetControl('INDICATOR'+Chr(Ord('0')+I))).Picture.Bitmap;

  {$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FLISTE'));
  {$ELSE}
  FListe := THDBGrid(GetControl('FLISTE'));
  {$ENDIF}
  FListe.OnDblClick := FListeOnDblClick;
  FListe.OnFlipSelection := FListeOnFlipSelection;
  FListe.OnColumnWidthsChanged := FListeOnColumnWidthsChanged;
  {$IFDEF EAGLCLIENT}
  FListe.OnDrawCell := FListeOnDrawCell;
  FListe.OnKeyDown  := FListeKeyDown;
  {$ELSE}
  FListe.OnDrawColumnCell := FListeOnDrawColumnCell;
  {$ENDIF}

  BManuel := TToolBarButton97(GetControl('BMANUEL'));
  BManuel.OnClick := BManuelOnClick;
  BAuto := TToolBarButton97(GetControl('BAUTO'));
  BAuto.OnClick := BAutoOnClick;

  TToolBarButton97(GetControl('BVIR')).OnClick := BVirOnClick;
  TToolBarButton97(GetControl('BAFFCONDITIONS')).OnClick := BAffConditionsOnClick;
  {15/06/04 : gestion de l'impression de la grille : il faut en effet le faire à la main, même en 2/3,
              car les montants sont convertis dans le DrawCell, il ne faut pas travailler sur la source
              de données du mul}
  TToolbarButton97(GetControl('BIMPRIMER')).OnClick := BImprimerClick;

  ArrondiN := THEdit(GetControl('ARRONDIN'));
  BanqueCpt := THValComboBox(GetControl('BANQUECPT'));
  BanqueCpt.OnChange := BanqueCptOnChange;

  Dossier := THMultiValComboBox(GetControl('DOSSIER'));
  Dossier.Visible := IsTresoMultiSoc;
  SetControlVisible('TREGROUPEMENT', IsTresoMultiSoc);

  {Gestion des filtres multi sociétés sur banquecp et dossier}
  Dossier.Plus := 'DOS_NODOSSIER ' + FiltreNodossier;
  Dossier.OnChange := DossierOnChange;

  THValComboBox(GetControl('DEVISEAFF')).OnChange := DeviseOnChange;
  THMultiValComboBox(GetControl('NATURE')).OnChange := NatureOnChange;

  {26/12/06 : FQ 10387 : Saisie de la date d'équilibrage}
  FDateEqui := V_PGI.DateEntree;
  SetControlText('DATEEQUI', DateToStr(FDateEqui));
  (GetControl('DATEEQUI') as THEdit).OnChange := DateEquiOnChange;

  {Création de l'objet contenant les taux de conversion}
  ObjTaux := TObjDevise.Create(FDateEqui);

  MajXXWhere;
  ChargeConditions;
  ADDMenuPop(PopupMenu, '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.OnClose ;
{---------------------------------------------------------------------------------------}
begin
  TobConditions.Free;
  TobVir.Free;
  TobListe.Free;
  LibereListe(lSoldes, True);
  {$IFDEF EAGLCLIENT}
  FListe.VidePile(True);
  {$ENDIF}
  FreeAndNil(ObjTaux);
  inherited;
end;

{Initialistion du PopupMenu
{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.OnPopupMenu(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  S, Compte : string;
begin
  {$IFDEF EAGLCLIENT}
  Compte := FListe.CellValues[COL_GENERAL, FListe.Row];
  {$ELSE}
  Compte := FListe.Fields[COL_GENERAL].AsString;
  {$ENDIF}
  PopupMenu.Items[0].Enabled := Compte <> Pivot; // Sur Pivot ?

  {$IFDEF EAGLCLIENT}
  if FListe.IsSelected(FListe.Row) then
  {$ELSE}
  if FListe.IsCurrentSelected then // Si rétiré ou rien à faire
  {$ENDIF}
    S := 'A&jouter'
  else
    S := '&Retirer';
  PopupMenu.Items[1].Caption := S;

  PopupMenu.Items[3].Enabled := BManuel.Enabled;
  PopupMenu.Items[4].Enabled := BAuto.Enabled;
end;

{Inverse sélection
{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.PopupAjouterRetirerOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF EAGLCLIENT}
  FListe.FlipSelection(FListe.Row);
  CocheDecoche;
  {$ELSE}
  FListe.FlipSelection;
  FListe.Invalidate;
  {$ENDIF}
end;

{Marque le compte pivot
{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.FListeOnDblClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF EAGLCLIENT}
  TFMul(Ecran).Q.TQ.Seek(FListe.Row - 1);
  if FListe.IsSelected(FListe.Row) then
    FListe.FlipSelection(FListe.Row);
  {$ELSE}
  if FListe.IsCurrentSelected then // Pivot non retirable
    FListe.FlipSelection;
  {$ENDIF}
  Pivot := QListe.FindField('TE_GENERAL').AsString;
  FListe.Invalidate;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.FListeOnFlipSelection(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if Chargement then Exit;
  if Assigned(QListe.FindField('TE_GENERAL')) and (QListe.FindField('TE_GENERAL').AsString = Pivot) then
    Pivot := '';
end;

{Fixe la taille de la colonne indicateur
 14/09/05 : FQ 10287 : Manifestement, il y a du y avoir un changement dans l'appel de cet
            évènement, ce qui qui créait en CWas un indice Hors limite => ajout du test
            if Fliste.ColCount = COL_SOLDEV + 1. J'en profite pour "nettoyer" en 2/3
{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.FListeOnColumnWidthsChanged(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF EAGLCLIENT}
  if Fliste.ColCount = COL_SOLDEV + 1 then begin
    FListe.ColWidths[COL_INDICATOR] := 18;
    if IsTresoMultiSoc then FListe.ColWidths[COL_LIBDOS] := 85
                       else FListe.ColWidths[COL_LIBDOS] := -1;
    FListe.ColWidths[COL_GENERAL]   := -1;
  end;
  {$ELSE}
       if Fliste.Columns.Count = COL_INDICATOR + 1 then FListe.Columns[COL_INDICATOR].Width := 18
  else if Fliste.Columns.Count = COL_DOSSIER   + 1 then FListe.Columns.Items[COL_DOSSIER].Visible := IsTresoMultiSoc
  else if Fliste.Columns.Count = COL_GENERAL   + 1 then FListe.Columns.Items[COL_GENERAL].Visible := False;
  {$ENDIF}
end;

{$IFDEF EAGLCLIENT}
{JP 03/08/03 : Remplacement de FListeOnDrawColumnCell par FlisteDrawCell pour rendre le code
               compatible eAGL
{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.FListeOnDrawCell(Sender: TObject;ACol, ARow: Longint; ARect: TRect; AState: TGridDrawState);
{---------------------------------------------------------------------------------------}
var
  Compte : String;
  aTob   : TOB;
  N      : Integer;
  Rc     : TRect;
  X      : PColor;
  Text   : array[0..255] of Char;
  F      : TAlignment ;
  Grille : THgrid ;
  ME, MD : Double;
  NewMnt : string;
  Ok     : Boolean;
begin
  Grille := THGrid(Sender);
  F := taCenter;

  {Conversion des montants dans la devise d'affichage}
  if (ARow > 0) and (ACol in [COL_SOLDEV, COL_SOLDEO]) then begin
    {S'il y a un filtre sur les natures, on récupère le solde dans la liste
     15/06/07 : FQ 10406 : quoiqu'il arrive, il faut reprendre les soldes dans la liste des soldes
    if (Trim(ClauseNat) <> '') then begin}
    n := lSoldes.IndexOf(Grille.CellValues[COL_GENERAL, ARow]);
    if n > -1 then begin
      if ACol = COL_SOLDEV then ME := Valeur(TObjSolde(lSoldes.Objects[n]).soVal)
                           else ME := Valeur(TObjSolde(lSoldes.Objects[n]).soOpe);
    end
    else begin
      ME := Valeur(Grille.Cells[ACol, ARow]);
      {Conversion de ME en Euros}
      ObjTaux.ConvertitMnt(ME, MD, Grille.CellValues[COL_DEVISE, ARow], Grille.CellValues[COL_GENERAL, ARow]);
    end;
    (*
    end
    {... sinon on prend le solde dans la table}
    else begin
      ME := Valeur(Grille.Cells[ACol, ARow]);
      {Conversion de ME en Euros}
      ObjTaux.ConvertitMnt(ME, MD, Grille.CellValues[COL_DEVISE, ARow], Grille.CellValues[COL_GENERAL, ARow]);
    end;
    *)

    {Conversion de ME dans la devise d'affichage}
    StrPCopy(Text, FormatFloat('#,##0.00', ME / TauxAff));
    NewMnt := FormatFloat('#,##0.00', ME / TauxAff);
  end
  else

  {Éventuelle réaffectation des dates}
  if (ARow > 0) and (ACol in [COL_DATEV, COL_DATEC]) then begin
    {S'il y a un filtre sur les natures, on récupère la date dans la liste
     15/06/07 : FQ 10406 : quoiqu'il arrive, il faut reprendre les Dates dans la liste des soldes
     if (Trim(ClauseNat) <> '') then begin}
    n := lSoldes.IndexOf(Grille.CellValues[COL_GENERAL, ARow]);
    if n > -1 then begin
      {JP 02/03/04 : COL_DATEV et non COL_SOLDEV}
      if ACol = COL_DATEV then Compte := TObjSolde(lSoldes.Objects[n]).dtVal
                          else Compte := TObjSolde(lSoldes.Objects[n]).dtOpe;
      StrPCopy(Text, Compte);
    end
    else
      StrPCopy(Text, Grille.Cells[ACol, ARow]);
    (*
    end
    else
      StrPCopy(Text, Grille.Cells[ACol, ARow]);
    *)
  end
  else

  {On précise la devise dans les titres des colonnes}
  if (ARow = 0) and (ACol in [COL_SOLDEV, COL_SOLDEO]) then
    StrPCopy(Text, Grille.Cells[ACol, ARow] + '( ' + GetControlText('DEVISEAFF') + ' )')

  else
    StrPCopy(Text, Grille.Cells[ACol, ARow]);

  Grille.Canvas.Font.Style := Grille.Canvas.Font.Style - [fsItalic] ;

  N := -1;
  {Récupération des conditions d'équilibrage}
  Compte := Grille.Cells[COL_GENERAL, ARow];
  aTob   := TobListe.FindFirst(['TE_GENERAL'], [Compte], False);
  Ok     := (aTob <> nil) and (VarToStr(aTob.GetValue('MONTANT')) <> '') and (VarToStr(aTob.GetValue('MONTANT')) <> '0');

  if Grille.Objects[0, ARow] = nil then begin
    X := PColor.Create;
    X.Drapeau := '';
    Grille.Objects[0, ARow] := X;
  end;

  if (ARow = 0) and (ACol <> 0) then begin
    Grille.Canvas.Brush.Color := clBtnFace;
    Grille.Canvas.Font .Color := clBlack;
    F := taCenter
  end

  else with Grille.Canvas do begin
    if TobListe.Detail.Count > 0 then begin
      if ACol in [COL_DATEC, COL_DATEV, COL_DEVISE] then
        F := taCenter
      else if ACol in [COL_LIBDOS, COL_GENERAL, COL_LIBELLE] then
        F := taLeftJustify
      else
        F := taRightJustify;

      {Lignes fixes}
      if (gdFixed in AState) then
        Font.Style := [fsBold]

      {Lignes sélectionnées}
      else if (gdSelected in AState) then begin
        if PColor(Grille.Objects[0,ARow]).Drapeau='*' then begin {Lignes à retirer}
          Brush.Color:= clHighlight ;
          Font.Color := clHighlightText ;
          Font.Style := [fsItalic] ;
          N := 1;
        end
        else begin
          Font.Style := [];
          Brush.Color:= clHighlight ;
          Font.Color := clHighlightText ;
          if Grille.Cells[COL_GENERAL, ARow] = Pivot then N := 0 {Ligne Pivot}
                                                     else if Ok then N := 2; {Lignes nécessitant un équilibrage}
        end;
      end

      {Lignes ni fixes, ni sélectionnées}
      else begin
        if PColor(Grille.Objects[0,ARow]).Drapeau = '*' then begin {Lignes à ne pas traitée}
         Font.Color := clBtnText;
         Brush.Color := FListe.FixedColor;
         Font.Style := [fsItalic] ;
         N := 1;
        end
        else if Grille.Cells[COL_GENERAL, ARow] = Pivot then begin {Ligne du compte Pivot}
          Font.Style := [fsBold];
          Font.Color := clInfoText;
          Brush.Color := clInfoBk;
          N := 0;
        end
        else begin {Autres lignes ...}
          Font.Style := [];
          Font.Color := clBtnText;
          if Ok then begin {... nécessitant un équilibrage}
            N := 2;
            Brush.Color := AltColors[V_PGI.NumAltCol];
          end
          else {... ne nécessitant pas d'équilibrage}
            Brush.Color := clHighlightText;

        end;
      end;
    end;

    if ACol = 0 then begin
      Rc := ARect;
      DrawEdge(Handle, Rc, BDR_RAISEDINNER, BF_RECT or BF_ADJUST);
      Brush.Color := clBtnFace;
      FillRect(Rc);
      if N >= 0 then
        Grille.Canvas.Draw((ARect.Left+ARect.Right) div 2 -6, (ARect.Top+ARect.Bottom) div 2 -5, Indicator[N]); // Centré
    end;
  end;

  if (ACol > 0) then begin
    case F of
      taRightJustify : ExtTextOut(Grille.Canvas.Handle, ARect.Right - Grille.Canvas.TextWidth(NewMnt{Grille.Cells[ACol,ARow]})-3,
                                 ARect.Top + 2,ETO_OPAQUE or ETO_CLIPPED, @ARect, Text, StrLen(Text), nil) ;
      taCenter       : ExtTextOut(Grille.Canvas.Handle, ARect.Left + ((ARect.Right-ARect.Left-Grille.canvas.TextWidth(Text{Grille.Cells[ACol,ARow]})) div 2),
                           ARect.Top + 2, ETO_OPAQUE or ETO_CLIPPED, @ARect, Text, StrLen(Text), nil) ;
      else
        ExtTextOut(Grille.Canvas.Handle, ARect.Left + 2, ARect.Top + 2, ETO_OPAQUE or ETO_CLIPPED,
                   @ARect, Text, StrLen(Text), nil) ;
    end;

    if ((gdfixed in AState) and Grille.Ctl3D) then begin
      DrawEdge(Grille.Canvas.Handle, ARect, BDR_RAISEDINNER, BF_BOTTOMRIGHT);
      DrawEdge(Grille.Canvas.Handle, ARect, BDR_RAISEDINNER, BF_TOPLEFT);
    end;
  end;
end;

{JP 11/08/03 : Pour dire qu'une ligne est sélectionnée}
{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.CocheDecoche;
{---------------------------------------------------------------------------------------}
begin
  if PColor(FListe.Objects[0, FListe.Row]).Drapeau = '*' then
    PColor(Fliste.Objects[0, FListe.Row]).Drapeau := ''
  else
    PColor(Fliste.Objects[0, FListe.Row]).Drapeau := '*';

  Fliste.Invalidate ;
  Fliste.Refresh;
  Application.ProcessMessages;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
{---------------------------------------------------------------------------------------}
begin
  if (Key = VK_SPACE) then
    CocheDecoche;
end;

{$ELSE}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 28/11/2001
Modifié le ... :   /  /
Description .. : Marque les lignes selon leur type (compte pivot, sélection...)
Mots clefs ... : SELECTION
*****************************************************************}
procedure TOF_EQUILIBRAGE.FListeOnDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  Compte : string;
  aTob   : TOB;
  N      : Integer;
  Rc     : TRect;
  ME, MD : Double;
begin
  try
    N := -1;
    with FListe.Canvas do begin
      if TobListe.Detail.Count > 0 then // Pas vide
        if FListe.Fields[COL_GENERAL].AsString = Pivot then begin	// Pivot
          Font.Style := [fsBold];
          Font.Color := clInfoText;
          Brush.Color := clInfoBk;
          N := 0;
        end
        {= ? FListe.IsCurrentSelected}
        else if Font.Style = [fsItalic] then begin {Sélectionné}
          Font.Color := clBtnText;
          Brush.Color := FListe.FixedColor;
          N := 1;
        end
        else begin // Normal
          //Font.Color := FListe.Font.Color;
          Compte := FListe.Fields[COL_GENERAL].AsString;
          aTob   := TobListe.FindFirst(['TE_GENERAL'], [Compte], False);
          if Assigned(aTob) and (VarToStr(aTob.GetValue('MONTANT')) <> '') and
             (VarToStr(aTob.GetValue('MONTANT')) <> '0') then begin // Il y a qq chose à faire !
            if V_PGI.NumAltCol=0 then
              Brush.Color := clInfoBk
            else
              Brush.Color := AltColors[V_PGI.NumAltCol];
            N := 2;
          end;
        end;

      {Colonne indicateurs}
      if DataCol = 0 then begin
        Rc := Rect;
        DrawEdge(Handle, Rc, BDR_RAISEDINNER, BF_RECT or BF_ADJUST);
        Brush.Color := FListe.FixedColor;
        FillRect(Rc);
        // ou DrawFrameControl(Handle, Rect, DFC_BUTTON, DFCS_BUTTONPUSH);  plus dur
        if N >= 0 then
          Draw((Rect.Left+Rect.Right) div 2 -6, (Rect.Top+Rect.Bottom) div 2 -5, Indicator[N]); // Centré
      end

      else begin
        if (gdSelected in State) then begin // Ligne courante
          Font.Color := clHighlightText;
          Brush.Color := clHighlight;
        end;

        FListe.DefaultDrawColumnCell(Rect, DataCol, Column as HColumn, State); // DefaultDrawing False (THDBGrid)
      end;
    end;

    {Affectation éventuelle des montants et conversion des montants dans la devise d'affichage}
    if (FListe.Row > 0) and (DataCol in [COL_SOLDEV, COL_SOLDEO]) then begin
      Rc := Rect;
      {S'il y a un filtre sur les natures, on récupère le solde dans la liste
       15/06/07 : FQ 10406 : quoiqu'il arrive, il faut reprendre les soldes dans la liste des soldes
      if (Trim(ClauseNat) <> '') then begin}
      n := lSoldes.IndexOf(FListe.Fields[COL_GENERAL].AsString);
      if n > -1 then begin
        if DataCol = COL_SOLDEV then ME := Valeur(TObjSolde(lSoldes.Objects[n]).soVal)
                                else ME := Valeur(TObjSolde(lSoldes.Objects[n]).soOpe);
        ObjTaux.ConvertitMnt(ME, MD, FListe.Fields[COL_DEVISE].AsString, FListe.Fields[COL_GENERAL].AsString);
      end
      else begin
        ME := Valeur(FListe.Fields[DataCol].AsString);
        {Conversion de ME en Euros}
        ObjTaux.ConvertitMnt(ME, MD, FListe.Fields[COL_DEVISE].AsString, FListe.Fields[COL_GENERAL].AsString);
      end;
      (*
      end
      {... sinon on prend le solde dans la table}
      else begin
        ME := Valeur(FListe.Fields[DataCol].AsString);
        {Conversion de ME en Euros}
        ObjTaux.ConvertitMnt(ME, MD, FListe.Fields[COL_DEVISE].AsString, FListe.Fields[COL_GENERAL].AsString);
      end;
      *)

      {Conversion de ME dans la devise d'affichage}
      FListe.Canvas.FillRect(Rc);
      FListe.Canvas.TextRect(Rc, Rect.Right - FListe.Canvas.TextWidth(FormatFloat('#,##0.00', ME / TauxAff))-3, (Rect.Top+Rect.Bottom) div 2 -5 , FormatFloat('#,##0.00', ME / TauxAff));
    end;

    {Éventuelle réaffectation des dates}
    if (FListe.Row > 0) and (DataCol in [COL_DATEV, COL_DATEC]) then begin
      Rc := Rect;
      {S'il y a un filtre sur les natures, on récupère la date dans la liste
       15/06/07 : FQ 10406 : quoiqu'il arrive, il faut reprendre les Dates dans la liste des soldes
      if (Trim(ClauseNat) <> '') then begin}
      n := lSoldes.IndexOf(FListe.Fields[COL_GENERAL].AsString);
      if n > -1 then begin
        if DataCol = COL_DATEV then Compte := TObjSolde(lSoldes.Objects[n]).dtVal
                               else Compte := TObjSolde(lSoldes.Objects[n]).dtOpe;

        FListe.Canvas.FillRect(Rc);
        {05/02/07 : Utilisation du DrawText plutôt que TextRect car il permet de centrer le texte
        FListe.Canvas.TextRect(Rc, (Rect.Left+3), (Rect.Top+Rect.Bottom) div 2 -5 , Compte);}
        DrawText(FListe.Canvas.Handle, PChar(Compte), Length(Compte), Rc, DT_CENTER	or DT_VCENTER);
      end;
      //end;
    end;

  except; // Parfois FListe vide !!
  end;
end;
{$ENDIF}

{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.RemplirTobListe;
{---------------------------------------------------------------------------------------}
var
  T  : TOB;
  n  : Integer;
  Q  : TQuery;
  MD : Double;
  ME : Double;
begin
  {On vide la liste}
  TobListe.ClearDetail;
  {ON récupère la requête du Mul ...}
  if IsTresoMultiSoc then
    {14/12/06 : Remplacement de FROM TRECRITURE, BANQUECP, DOSSIER par
                LEFT JOIN DOSSIER et LEFT JOIN BANQUECP car si BANQUECP et DOSSIER sont
                partagé sur une autre base que la base courante on passe de 25 minutes à 2 secondes !!!!}
    Q := OpenSQL('SELECT TE_GENERAL, TE_DEVISE, TE_SOLDEDEVVALEUR, BQ_LIBELLE, DOS_LIBELLE, TE_NODOSSIER ' +
                 'FROM TRECRITURE LEFT JOIN BANQUECP ON BQ_CODE = TE_GENERAL ' +
                 'LEFT JOIN DOSSIER ON TE_NODOSSIER = DOS_NODOSSIER WHERE ' + GetControlText('XX_WHERE'), True)
  else
    Q := OpenSQL('SELECT TE_GENERAL, TE_DEVISE, TE_SOLDEDEVVALEUR, BQ_LIBELLE, TE_NODOSSIER ' +
                 'FROM TRECRITURE LEFT JOIN BANQUECP ON BQ_CODE = TE_GENERAL ' +
                 'WHERE ' + GetControlText('XX_WHERE'), True);

  try
    {... Puis remplissage de la TobListe}
    while not Q.Eof do begin
      T := TOB.Create('***', TobListe, 0);
      T.AddChampSupValeur('TE_GENERAL'  , Q.FindField('TE_GENERAL'  ).AsString);
      T.AddChampSupValeur('TE_DEVISE'   , Q.FindField('TE_DEVISE'   ).AsString);
      T.AddChampSupValeur('BQ_LIBELLE'  , Q.FindField('BQ_LIBELLE'  ).AsString);
      T.AddChampSupValeur('TE_NODOSSIER', Q.FindField('TE_NODOSSIER').AsString);
      if IsTresoMultiSoc then
        T.AddChampSupValeur('DOS_LIBELLE' , Q.FindField('DOS_LIBELLE' ).AsString)
      else
        T.AddChampSupValeur('DOS_LIBELLE' , GetParamSocSecur('SO_LIBELLE', ''));

      {S'il y a un filtre sur les natures, on récupère le solde dans la liste
       01/08/07 : Nlle gestion des soldes : on ne cherche plus les soldes sur les écritures}
      if (Trim(ClauseNat) <> '') or IsNewSoldes then begin
        n := lSoldes.IndexOf(Q.FindField('TE_GENERAL').AsString);
        {01/08/07 : espérons que cela sera toujours le cas !!!}
        if n > -1 then begin
          ME := Valeur(TObjSolde(lSoldes.Objects[n]).soVal);
        end
        else begin
          ME := Q.FindField('TE_SOLDEDEVVALEUR').AsFloat;
          {Conversion de ME en Euros}
          ObjTaux.ConvertitMnt(ME, MD, Q.FindField('TE_DEVISE').AsString, Q.FindField('TE_GENERAL').AsString);
        end;
      end
      {... sinon on prend le solde dans la table}
      else begin
        ME := Q.FindField('TE_SOLDEDEVVALEUR').AsFloat;
        {Conversion de ME en Euros}
        ObjTaux.ConvertitMnt(ME, MD, Q.FindField('TE_DEVISE').AsString, Q.FindField('TE_GENERAL').AsString);
      end;

      T.AddChampSupValeur('TE_SOLDEDEVVALEUR', ME / TauxAff);
      Q.Next;
    end;
  finally
    Ferme(Q);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 28/11/2001
Modifié le ... :   /  /
Description .. : Recalcule le montant des virements possibles pour toute la
Suite ........ : grille.
Mots clefs ... : VIREMENT
*****************************************************************}
procedure TOF_EQUILIBRAGE.CalculerVir;
var
  I   : Integer;
  Mnt : Double;
begin
  {Empèche le rafraichissement de la grille avant le rechargement de TobListe}
  ValidateRect(FListe.Handle, nil);
  RemplirTobListe;

  if TobListe.Detail.Count > 0 then begin
    TobListe.Detail[0].AddChampSupValeur('MONTANT', '', True);
    TobListe.Detail[0].AddChampSup('SELECTED', True);
  end;

  for I := 0 to TobListe.Detail.Count - 1 do
    with TobListe.Detail[I] do begin
      Mnt := EquilibrerCompte(GetValue('TE_GENERAL'), GetValue('TE_SOLDEDEVVALEUR'));

      {On récupère le montant dans la devise du compte}
      if Mnt <> 0 then
        PutValue('MONTANT', FloatToStr(Mnt));
    end;

  {Maintenant, on peut redessiner la grille}
  FListe.Invalidate;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 21/11/2001
Modifié le ... :   /  /
Description .. : Constitution de la Tob des virements à partir des comptes
Suite ........ : sélectionnés et dont le solde ne satisfait pas aux conditions.
Mots clefs ... : VIREMENT
*****************************************************************}
procedure TOF_EQUILIBRAGE.BAutoOnClick(Sender : TObject);
var
  I         : Integer;
  aTob,
  Tob1Vir   : TOB;
  PLibelle,
  SLibelle,
  DLibelle,
  S, SCompte,
  DCompte,
  SDevise,
  DDevise   : string;
  SSolde,
  DSolde,
  MntDev,
  Mnt,
  PSolde,
  Solde     : Double;
  PDossier  : string;
  DDossier  : string;
  SDossier  : string;
  PDossLib  : string;
  DDossLib  : string;
  SDossLib  : string;
begin
  if Pivot = '' then begin
    TrShowMessage(Ecran.Caption, 8, '', '');
    Exit;
  end;

  { Marque les lignes à prendre}
  for I := 0 to TobListe.Detail.Count - 1 do
    TobListe.Detail[I].PutValue('SELECTED', 'X');

  {Retire les lignes sélectionnées (les lignes sélectionnées sont celles qu'il faut exclure du traitemen !)}
  for I := 0 to FListe.NbSelected - 1 do begin
    FListe.GotoLeBookmark(I);
    SCompte := VarToStr(GetField('TE_GENERAL'));
    TobListe.FindFirst(['TE_GENERAL'], [SCompte], False).SetString('SELECTED', '-');
  end;

  {Place le compte pivot au début}
  aTob := TobListe.FindFirst(['TE_GENERAL'], [Pivot], False);
  aTob.ChangeParent(TobListe, 0);
  {Solde initial du pivot}
  PSolde := aTob.GetDouble('TE_SOLDEDEVVALEUR');
  PLibelle := aTob.GetString('BQ_LIBELLE');
  PDossier := aTob.GetString('TE_NODOSSIER');
  PDossLib := aTob.GetString('DOS_LIBELLE');

  {On ne prend pas le pivot, on part de 1}
  for I := 1 to TobListe.Detail.Count-1 do begin
    aTob := TobListe.Detail[I];
    S := aTob.GetValue('MONTANT');

    if (aTob.GetValue('SELECTED') = 'X') and (S <> '') then begin
      {Par défaut Source = compte, Destination = Pivot, Montant négatif}
      SCompte := aTob.GetValue('TE_GENERAL');
      Solde   := aTob.GetValue('TE_SOLDEDEVVALEUR');
      SDevise := aTob.GetValue('TE_DEVISE');
      SLibelle := aTob.GetString('BQ_LIBELLE');
      SDossier := aTob.GetString('TE_NODOSSIER');
      SDossLib := aTob.GetString('DOS_LIBELLE');
      {DSolde : le solde du pivot sera affecté après cumul}
      DCompte := Pivot;
      DLibelle := PLibelle;
      DDossier := PDossier;
      DDossLib := PDossLib;
      DDevise := TobListe.Detail[0].GetValue('TE_DEVISE');
      Mnt     := Valeur(S);
      Solde   := Solde  + Mnt;
      PSolde  := PSolde - Mnt;

      {Le montant est positif, le compte a besoin d'argent : le pivot devient donc Source du virement}
      if Mnt > 0 then begin
        DLibelle := SLibelle;
        SLibelle := PLibelle;
        DDossier := SDossier;
        SDossier := PDossier;
        DDossLib := SDossLib;
        SDossLib := PDossLib;
        DCompte := SCompte;
        SCompte := Pivot;
        S := SDevise;
        SDevise := DDevise;
        DDevise := S;
        SSolde  := PSolde;
        DSolde  := Solde;
      end
      {Le montant est négatif, le compte a trop d'argent : le pivot devient donc Destination du virement}
      else begin
        Mnt := -Mnt;
        SSolde  := Solde;
        DSolde  := PSolde;
        MntDev  := Mnt;
      end;

      Tob1Vir := TOB.Create('1', TobVir, -1);
      Tob1Vir.AddChampSupValeur('SCOMPTE', SCompte);
      Tob1Vir.AddChampSupValeur('SDEVISE', SDevise);
      Tob1Vir.AddChampSupValeur('DCOMPTE', DCompte);
      Tob1Vir.AddChampSupValeur('DDEVISE', DDevise);
      {Les montants sont calculés dans la devise d'affichage, il faut donc les convertir ...}
      ObjTaux.ConvertitMnt(Mnt, MntDev, DeviseAff, SCompte);
      Tob1Vir.AddChampSupValeur('MONTANT', MntDev); {Devise du compte source ?}
      ObjTaux.ConvertitMnt(SSolde, MntDev, DeviseAff, SCompte);
      {On mémorise le solde dans la devise du compte source}
      Mnt := MntDev;
      Tob1Vir.AddChampSupValeur('SSOLDE', MntDev);
      ObjTaux.ConvertitMnt(DSolde, MntDev, DeviseAff, DCompte);
      Tob1Vir.AddChampSupValeur('DSOLDE', MntDev); {Soldes + virement}
      Tob1Vir.AddChampSup('SOURCE', False);
      {07/05/04 : Pour améliorer la lisibilité du TobViewer, ajout de 4 nouveaux champs et
                  modification de PutSource}
      Tob1Vir.AddChampSup('DESTINATION', False);
      Tob1Vir.AddChampSup('VIREMENT', False);
      Tob1Vir.AddChampSupValeur('SLIBELLE', SLibelle);
      Tob1Vir.AddChampSupValeur('DLIBELLE', DLibelle);
      Tob1Vir.AddChampSupValeur('DDOSSIER', DDossier);
      Tob1Vir.AddChampSupValeur('SDOSSIER', SDossier);
      Tob1Vir.AddChampSupValeur('DDOSSLIB', DDossLib);
      Tob1Vir.AddChampSupValeur('SDOSSLIB', SDossLib);
      {26/12/06 : FQ 10387 : Paramétrage de la date du virement}
      Tob1Vir.AddChampSupValeur('DATEEQUI', FDateEqui);
      PutSource(Tob1Vir, Mnt);
    end;
  end;

  {Solde final du pivot :
   On commence par le convertir de la devise d'affichage dans la devise du compte...}
  ObjTaux.ConvertitMnt(PSolde, MntDev, DeviseAff, Pivot);
  {... puis on le met à jour}
  MajSolde(TobVir, Pivot, PSolde);

  BManuelOnClick(BAuto);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.BManuelOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Ok : Boolean;
  Ch : string;
begin
  try
    {Vide si équilibrage manuel}
    TheTob := TobVir;
    if UpperCase(TComponent(Sender).Name) <> 'BAUTO' then Ch := 'MANUEL;';
    Ch := Ch + GetControlText('DATEEQUI') + ';' + ClauseNat + ';';
    OK := TRLanceFiche_ResultEqui('TR','TRRESULTEQUI', '', '', Ch) <> '';
  finally
    TobVir.ClearDetail;
  end;

  {On regarde s'il y a des ordres de virements en attente}
  MajBVir;
  {On lance l'écran de génération des virements, s'il y a des virements et qu'on n'a pas abandonné}
  if Ok and GetControlEnabled('BVIR') then
    BVirOnClick(Sender);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.BVirOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if TRLanceFiche_Virement('TR','TRVIREMENT', '', '', DeviseAff) <> '' then
    TFMul(Ecran).ChercheClick; {Rafraichit si modif dans TREcriture}
  MajBVir;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.BAffConditionsOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Lequel, Action: String;
begin
  Lequel := QListe.FindField('TE_GENERAL').AsString;
  {On regarde si des conditions existent pour la ligne courante}
  if TobConditions.FindFirst(['TCE_GENERAL'], [Lequel], False) <> Nil then
    Action := 'ACTION=MODIFICATION;' + Lequel
  else
  begin
    Action := 'ACTION=CREATION;' + Lequel;
    LeQuel := '';
  end;
  TRLanceFiche_ConditionEqui('TR','TRCONDITIONEQUI', '', LeQuel, Action);
  {Rafraichissment de la grille : le OnUpdate relance EquilibrerCompte}
  TFMul(Ecran).ChercheClick;
end;

{Charge les conditions d'équilibrage et les devises
{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.ChargeConditions;
{---------------------------------------------------------------------------------------}
var
  C : string;
  Q : TQuery;
begin
  C := 'SELECT TCE_GENERAL,TCE_TYPESOLDE,TCE_SOLDEDEB,TCE_SOLDECRED,TCE_SOLDECONS,' +
			'TCE_DEBITMIN,TCE_CREDITMIN,TCE_ARRONDI FROM CONDITIONEQUI';
  {Par défaut, pas de sélection}
  if Chargement then BanqueCpt.ItemIndex := 0;

  {L'affectation de ItemIndex n'exécute pas le OnChange}
  if Chargement then BanqueCpt.OnChange(BanqueCpt);

  {Charge les conditions d'équilibrage des comptes de la société en cours}
  Q := OpenSQL(C, True);
  {Empèche le rafraichissement de la grille avant le rechargement de TobListe}
  ValidateRect(FListe.Handle, nil);
  TobConditions.ClearDetail;
  TobConditions.LoadDetailDB('1', '', '', Q, False, True);
  Ferme(Q);
  FListe.Invalidate;
end;

{Affichage de la devise
{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.DeviseOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  DeviseAff := THValComboBox(GetControl('DEVISEAFF')).Value;
  {Si TRECRITURE est vide !!!}
  if DeviseAff = '' then DeviseAff := V_PGI.DevisePivot;
  AssignDrapeau(TImage(GetControl('IDEV')), DeviseAff);
end;

{---------------------------------------------------------------------------------------}
function TOF_EQUILIBRAGE.MajXXWhere : Boolean;
{---------------------------------------------------------------------------------------}
var
  SQL : string;
  Dt  : string;
  Nat : string;
  Reg : string;
  S   : string;
  wh  : string;
  Banque : string;
begin
  wh := '';

  {02/03/07 : FQ 10414 : Le code ci-dssous est déplacé de BanqueCptOnChange à MajXXWhere}
  Banque := GetControlText('BANQUECPT');
  if Banque = '' then S := ''
                 else S := 'BQ_BANQUE = "' + Banque + '"';

  Dt := UsDateTime(FDateEqui); {FQ 10387}
  {Clause : On ne prend que les soldes inférieurs ou égal à la date d'entrée.}
  {Gestion des natures des écritures}
  Nat := GetControlText('NATURE');
  if (Pos('<<' , Nat) = 0) and (Trim(Nat) <> '') then {<<Tous>> ou vide}
    {JP 15/04/04 : Nouvelle gestion des natures}
    ClauseNat := ' AND TE_NATURE IN (' + GetClauseIn(Nat) + ')'
  else
    ClauseNat := '';

  if Dossier.Tous then Reg := FiltreNodossier
                  else Reg := FiltreNodossier(Dossier.Value);
  if Reg <> '' then Reg := ' AND TE_NODOSSIER ' + Reg;

  {31/10/06 : Si besoin d'optimisation, remplacer le Group By par un Where T1.TE_GENERAL = TE_GENERAL
   13/08/07 : A étudier, mais il me semblerait mieux de supprimer le IN SELECT et de le remplacer par
              SELECT DISTINCT dans la DBLISTE}
  SQL := 'TE_CLEOPERATION IN (SELECT MAX(TE_CLEOPERATION) FROM TRECRITURE T1 ' +
         'WHERE TE_DATECOMPTABLE <= "' + Dt + '" ' + ClauseNat + Reg +
         ' AND NOT (TE_CODECIB IN ("' + CODECIBCOURANT + '", "' + CODECIBTITRE + '"))' +
         ' AND NOT (TE_NUMTRANSAC LIKE "' + CODEMODULECOU + '%") GROUP BY TE_GENERAL)';

  SQL := SQL + ClauseNat;
  {10/08/06 : Je ne pense pas que l'on traite les comptes courants !}
  SQL := SQL + ' AND ' + BQCLAUSEWHERE;

  {$IFDEF TRCONF}
  wh := TObjConfidentialite.GetWhereConf(V_PGI.User, tyc_Banque);
  if wh <> '' then wh := ' AND (' + wh + ') ';
  {$ENDIF TRCONF}
  SQL := SQL + Wh;
  
  {On va charger en mémoire les soldes recalculés en fonction des natures sélectionnées, pour les affecter dans
   la grille aum moment du dessin des cellules}
//  ChargerListeSolde;

  if S <> '' then
    SQL := SQL + ' AND ' + S;

  Result := GetControlText('XX_WHERE') <> SQL;
  if Result then
    SetControlText('XX_WHERE', SQL);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.BanqueCptOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {02/03/07 : FQ 10414 : déplacé dans le MajXXWhere
  Banque := GetControlText('BANQUECPT');
  if Banque = '' then
    S := ''
  else begin
    S := 'TE_GENERAL IN (SELECT DISTINCT BQ_CODE FROM BANQUECP WHERE BQ_BANQUE = "' +
		 Banque + '")';
  end;}

  if MajXXWhere then
    ;//TFMul(Ecran).ChercheClick;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.DateEquiOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if IsValidDate(GetControlText('DATEEQUI')) then begin
    FDateEqui := StrToDate(GetControlText('DATEEQUI'));
    FreeAndNil(ObjTaux);
    ObjTaux := TObjDevise.Create(FDateEqui);
    MajXXWhere;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.DossierOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  MajXXWhere;
end;

{Si la nature change on met à jour la clause where de la fiche
{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.NatureOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if MajXXWhere then
    ;//TFMul(Ecran).ChercheClick;
end;

{On charge les soldes à la date d'entrée, puis dans le dessin des cellules on affectera
 les valeurs. L'autre possibilité aurait été de faire un faux mul car les soldes qui
 figurent dans les tables sont sur toutes les natures
{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.ChargerListeSolde;
{---------------------------------------------------------------------------------------}
var
  Obj : TObjSolde;
  Q   : TQuery;
  n   : Integer;
  d   : TDateTime;
  f   : Double;
begin
  d := DebutAnnee(FDateEqui); {FQ 10387}
  {Libération de la liste}
  LibereListe(lSoldes, False);
  {Chargement des soldes.
   11/07/07 : On ne travaille plus sur les soldes de TRECRITURE, on les recalcule systématiquement à la volée}
  if (Trim(ClauseNat) <> '') or IsNewSoldes then begin
    {On commence par récupérer les solde forcés en début d'année}
    Q := OpenSQL('SELECT TE_SOLDEDEV, TE_SOLDEDEVVALEUR, TE_COTATION, TE_GENERAL FROM TRECRITURE WHERE TE_DATECOMPTABLE = "' +
                 UsDateTime(d) + '" AND TE_QUALIFORIGINE = "' + CODEREINIT + '"', True);
    while not Q.EOF do begin
      f := Q.Fields[2].AsFloat;
      if f = 0 then f := 1;
      Obj := TObjSolde.Create;
      Obj.dtVal := DateToStr(d);
      Obj.soVal := FloatToStr(Q.Fields[1].AsFloat / f);
      Obj.dtOpe := DateToStr(d);
      Obj.soOpe := FloatToStr(Q.Fields[0].AsFloat / f);
      lSoldes.AddObject(Q.Fields[3].AsString, Obj);
      Q.Next;
    end;
    Ferme(Q);

    {Puis on cumule avec les écritures en valeur depuis le premier janvier de l'année en cours}
    Q := OpenSQL('SELECT SUM(TE_MONTANT), MAX(TE_DATEVALEUR), TE_GENERAL FROM TRECRITURE WHERE TE_DATEVALEUR <= "' +
                 UsDateTime(FDateEqui) + '" AND TE_DATEVALEUR > "' + UsDateTime(d) + '" ' + ClauseNat + ' GROUP BY TE_GENERAL', True);
    while not Q.EOF do begin
      n := lSoldes.IndexOf(Q.Fields[2].AsString);
      if n = -1 then begin
        Obj := TObjSolde.Create;
        Obj.dtVal := Q.Fields[1].AsString;
        Obj.soVal := Q.Fields[0].AsString;
        Obj.dtOpe := DateToStr(Date);
        Obj.soOpe := '0';
        lSoldes.AddObject(Q.Fields[2].AsString, Obj);
      end else begin
        Obj := TObjSolde(lSoldes.Objects[n]);
        Obj.dtVal := Q.Fields[1].AsString;
        Obj.soVal := FloatToStr(Valeur(Obj.soVal) + Q.Fields[0].AsFloat);
      end;

      Q.Next;
    end;
    Ferme(Q);

    {Puis on cumule avec les écritures en opération depuis le premier janvier de l'année en cours}
    Q := OpenSQL('SELECT SUM(TE_MONTANT), MAX(TE_DATECOMPTABLE), TE_GENERAL FROM TRECRITURE WHERE TE_DATECOMPTABLE <= "' +
                 UsDateTime(FDateEqui) + '" AND TE_DATECOMPTABLE > "' + UsDateTime(d) + '" ' + ClauseNat + ' GROUP BY TE_GENERAL', True);
    while not Q.EOF do begin
      n := lSoldes.IndexOf(Q.Fields[2].AsString);
      if n = -1 then begin
        Obj := TObjSolde.Create;
        Obj.dtOpe := Q.Fields[1].AsString;
        Obj.soOpe := Q.Fields[0].AsString;
        Obj.dtVal := DateToStr(Date);
        Obj.soVal := '0';
        lSoldes.AddObject(Q.Fields[2].AsString, Obj);
      end else begin
        Obj := TObjSolde(lSoldes.Objects[n]);
        Obj.dtOpe := Q.Fields[1].AsString;
        Obj.soOpe := FloatToStr(Valeur(Obj.soOpe) + Q.Fields[0].AsFloat);
      end;
      Q.Next;
    end;
    Ferme(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.OnUpdate ;
{---------------------------------------------------------------------------------------}
begin
 { Pour le moment, tous les comptes doivent être de la même devise (= DeviseAff) !!!!!}
  inherited;

  {Recharge les conditions d'équilibrage}
  ChargeConditions;
  {On recharge la TobListe ????}
  CalculerVir;
  MajBVir;

  {On retire le pivot si le compte est absent}
  if (Pivot<>'') and (TobListe.FindFirst(['TE_GENERAL'], [Pivot], False) = nil) then
    Pivot := '';
end ;

{Lorsque l'on change de ligne, on affiche les conditions du compte en cours
{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.OnDisplay;
{---------------------------------------------------------------------------------------}
var
  Compte: String;
  aTob: TOB;
  ME, MD : Double;

    {----------------------------------------------------------}
    procedure Raz(Name: String);
    {----------------------------------------------------------}
    var
      Edit: THNumEdit;
    begin
      Edit := THNumEdit(GetControl(Name));
      {Change la valeur pour provoquer la mise à jour à l'écran !}
      Edit.Value := Edit.Value -1;
      Edit.Text := '';
    end;

    {----------------------------------------------------------}
    procedure Hide(Name: String);
    {----------------------------------------------------------}
    begin
      THLabel(GetControl('T' + Name)).Visible := False;
      THNumEdit(GetControl(Name)).Visible := False;
    end;

    {----------------------------------------------------------}
    procedure Affecte(Name: String);
    {----------------------------------------------------------}
    var
      Edit: THNumEdit;
    begin
      THLabel(GetControl('T' + Name)).Visible := True;
      Edit := THNumEdit(GetControl(Name));
      Edit.Visible := True;
      {Conversion du montant dans la devise d'affichage}
      ME := Valeur(aTob.GetValue('TCE_' + Name));
      {On convertit ME en euros}
      ObjTaux.ConvertitMnt(ME, MD, VarToStr(GetField('TE_DEVISE')), Compte);
      {Puis on convertit ME dans la devise d'affichage}
      Edit.Value := Valeur(FormatFloat('#,##0.00', ME / TauxAff));
    end;

begin
  inherited;
  Compte := VarToStr(GetField('TE_GENERAL'));
  aTob   := TobConditions.FindFirst(['TCE_GENERAL'], [Compte], False);
  if aTob <> nil then begin
    if aTob.GetValue('TCE_TYPESOLDE') = 'X' then begin
      Affecte('SOLDEDEB');
      Affecte('SOLDECRED');
      Hide('SOLDECONS');
    end
    else begin
      Hide('SOLDEDEB');
      Hide('SOLDECRED');
      Affecte('SOLDECONS');
    end;

    Affecte('CREDITMIN');
    Affecte('DEBITMIN');
//    THNumEdit(GetControl('CREDITMIN')).Value := aTob.GetValue('TCE_CREDITMIN');
  //  THNumEdit(GetControl('DEBITMIN')).Value := aTob.GetValue('TCE_DEBITMIN');
    ArrondiN.Text := RechDom('TRARRONDIN', aTob.GetValue('TCE_ARRONDI'), False);
  end
  else begin
    Raz('SOLDEDEB');
    Raz('SOLDECRED');
    Raz('SOLDECONS');
    Raz('CREDITMIN');
    Raz('DEBITMIN');
    ArrondiN.Text := '';
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.OnLoad ;
{---------------------------------------------------------------------------------------}
begin
  inherited ;
  if Chargement then
    THValComboBox(GetControl('DEVISEAFF')).ItemIndex := THValComboBox(GetControl('DEVISEAFF')).Values.IndexOf(V_PGI.DevisePivot);
  DeviseOnChange(GetControl('DEVISEAFF'));
  Chargement := False;
  {Récupération du taux d'affichage par rapport à l'euro}
  TauxAff := RetPariteEuro(THValComboBox(GetControl('DEVISEAFF')).Value, FDateEqui);
  if TauxAff = 0 then TauxAff := 1;

  {On va charger en mémoire les soldes recalculés en fonction des natures sélectionnées, pour les affecter dans
   la grille aum moment du dessin des cellules
   15/06/07 : Déplacé depuis MajXXWhere}
  ChargerListeSolde;

  {$IFDEF EAGLCLIENT}
  FListe.ColFormats[COL_SOLDEO] := StrFMask(V_PGI.OkDecV, '', True); //'#.##0.00';
  FListe.ColFormats[COL_SOLDEV] := StrFMask(V_PGI.OkDecV, '', True); //'#.##0.00';
  {$ENDIF}
end;

{15/06/04 : Il faut gérer les états manuellement car les montants sont si besoin convertis
            dans la devise d'affichage dans le DrawCell => on ne peut dons s'appuyer sur la
            source de données dont les montants sont en devises
{---------------------------------------------------------------------------------------}
procedure TOF_EQUILIBRAGE.BImprimerClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  F : TOB;
  k : Integer;
  s : string;

    {$IFDEF EAGLCLIENT}
    procedure ChargerTob(Q : TOB);
    {$ELSE}
    procedure ChargerTob(Q : THQuery);
    {$ENDIF}
    var
      n   : Integer;
      Obj : TObjSolde;
    begin
      n := lSoldes.IndexOf(Q.FindField('TE_GENERAL').AsString);
      {15/06/07 : FQ 10406 : Théoriquement, cela ne devrait pas arriver que n = -1.
                  Rappel : la liste des soldes contient les soldes calculés à la volée en fonction
                  des critères de sélection. Inversement, Q contient les soldes issus de la requête
                  du mul (donc stockés dans la base) qui seront faux dès qu'il y aura un filtre sur
                  la nature.}
      if n > -1 then begin
        Obj := TObjSolde(lSoldes.Objects[n]);
        F.AddChampSupValeur('DATECOMPTABLE' , Obj.dtOpe);
        F.AddChampSupValeur('DATEVALEUR'    , Obj.dtVal);
        F.AddChampSupValeur('SOLDEDEV'      , Obj.soOpe);
        F.AddChampSupValeur('SOLDEDEVVALEUR', Obj.soVal);
      end
      else begin
        F.AddChampSupValeur('DATECOMPTABLE' , Q.FindField('TE_DATECOMPTABLE').AsString);
        F.AddChampSupValeur('DATEVALEUR'    , Q.FindField('TE_DATEVALEUR'   ).AsString);
        F.AddChampSupValeur('SOLDEDEV'      , Q.FindField('TE_SOLDEDEV'      ).AsFloat);
        F.AddChampSupValeur('SOLDEDEVVALEUR', Q.FindField('TE_SOLDEDEVVALEUR').AsFloat);
      end;

      F.AddChampSupValeur('SOCIETE'       , Q.FindField('TE_NODOSSIER').AsString);
      F.AddChampSupValeur('DEVISE'        , Q.FindField('TE_DEVISE'   ).AsString);
      F.AddChampSupValeur('GENERAL'       , Q.FindField('TE_GENERAL'  ).AsString);
      F.AddChampSupValeur('LIBELLE'       , Q.FindField('BQ_LIBELLE'  ).AsString);
      F.AddChampSupValeur('CHANGE'        , ObjTaux.GetParite(Q.FindField('TE_DEVISE').AsString, DeviseAff));
      F.AddChampSupValeur('AFFICHAGE'     , AnsiUpperCase(s));
    end;

begin
  {$IFDEF EAGLCLIENT}
  k := FListe.RowCount;
  if k < 2 then Exit;
  {$ELSE}
  k := TFMul(Ecran).Q.RecordCount;
  if k = 0 then Exit;
  {$ENDIF}
  T := TOB.Create('', nil, -1);
  try

    s := RechDom('TTDEVISE', DeviseAff, False);
    {$IFDEF EAGLCLIENT}
    while not TFMul(Ecran).Q.TQ.EOF do begin
      F := TOB.Create('', T, -1);
      ChargerTob(TFMul(Ecran).Q.TQ);
      TFMul(Ecran).Q.TQ.Next;
    end;
    {$ELSE}
    while not TFMul(Ecran).Q.EOF do begin
      F := TOB.Create('', T, -1);
      ChargerTob(TFMul(Ecran).Q);
      TFMul(Ecran).Q.Next;
    end;
    {$ENDIF}
    s := Ecran.Caption + ' (en : ' + s + ')';
    LanceEtatTOB('E', 'ECT', 'EQU', T, True, False, False, nil, '', s, False);
  finally
    FreeAndNil(T);
  end;
end;

initialization
  registerclasses ( [ TOF_EQUILIBRAGE ] ) ;
end.



