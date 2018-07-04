{***********UNITE*************************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 06/12/2001
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : FRAIS (FRAIS)
Mots clefs ... : TOM;FRAIS
*****************************************************************}
{-------------------------------------------------------------------------------------
  Version     |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
06.30.001.002  01/03/05   JP   FQ 10206 : Gestion de l'ordre des zones
06.50.001.019  19/09/05   JP   FQ 10292 : Problème d'ergonomie dans la grille des flux associés
--------------------------------------------------------------------------------------}

unit TomFrais ;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL, eFiche, UTob,
  {$ELSE}
  HDB, db, FE_Main, Fiche,
  {$ENDIF}
  StdCtrls, Controls, Classes, Forms, HCtrls, HEnt1, UTOM,
  Menus;

type
  TOM_FRAIS = class(TOM)
    procedure OnNewRecord               ; override;
    procedure OnUpdateRecord            ; override;
    procedure OnLoadRecord              ; override;
    procedure OnChangeField (F : TField); override;
    procedure OnArgument    (S : string); override;
  protected
    GridFlux  : THgrid;
    PopupMenu : TPopUpMenu;

    procedure ViderGrille     ;
    procedure AjouterLigne    ;
    procedure SupprimerLaLigne(ARow : Integer);
    procedure InsererUneLigne (ARow : Integer);
    procedure GererFormat     (Dev : string);
    procedure MajFluxAssocies ;

    procedure DeviseChange    (Sender : TObject);
    procedure GrilleOnBtnClick(Sender : TObject);
    procedure ModeOnChange    (Sender : TObject);
    procedure ppGilleOnClick  (Sender : TObject);
    procedure OnCellSelect    (Sender : TObject; ACol, ARow: Longint; var CanSelect: Boolean);
    procedure GrilleOnKeyDown (Sender : TObject; var Key : Word; Shift : TShiftState);
    procedure GrillOnCellExit (Sender : TObject; var ACol, ARow : Integer; var Cancel : Boolean);
    procedure GrillOnCellEnter(Sender : TObject; var ACol, ARow : Integer; var Cancel : Boolean);
    {19/09/05 : FQ 10292 : Pour exécuter le CellEnter en entrée de la Grille}
    procedure GrilleOnEnter   (Sender : TObject);

    procedure AfterOnShow;{01/03/05 : FQ 10206}
  end ;

procedure TRLanceFiche_Frais(Dom, Fiche, Range, Lequel, Arguments: string);

implementation

uses
  Commun, Windows, HMsgBox, LookUp, Constantes, Messages,
  ExtCtrls{TImage};

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_Frais(Dom, Fiche, Range, Lequel, Arguments: string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
var
  n : Byte;
begin
  inherited;
  Ecran.HelpContext := 150;
  GridFlux := THGrid(GetControl('GRIDFLUX'));
  GridFlux.OnSelectCell    := OnCellSelect;
  GridFlux.OnKeyDown       := GrilleOnKeyDown;
  GridFlux.OnElipsisClick  := GrilleOnBtnClick;
  GridFlux.ColEditables[1] := False;
  GridFlux.OnCellEnter     := GrillOnCellEnter;
  GridFlux.OnCellExit      := GrillOnCellExit;
  {19/09/05 : FQ 10292 : Pour exécuter le CellEnter en entrée de la Grille}
  GridFlux.OnEnter         := GrilleOnEnter;

  {$IFDEF EAGLCLIENT}
  THValComboBox(GetControl('TFR_MODEFRAIS')).OnChange := ModeOnChange;
  THValComboBox(GetControl('TFR_DEVISE')   ).OnChange := DeviseChange;
  {$ELSE}
  THDBValComboBox(GetControl('TFR_MODEFRAIS')).OnChange := ModeOnChange;
  THDBValComboBox(GetControl('TFR_DEVISE')   ).OnChange := DeviseChange;
  {$ENDIF}

  PopupMenu := TPopupMenu(GetControl('PPGRILLE'));
  for n := 0 to PopupMenu.Items.Count - 1 do
    PopupMenu.Items[n].OnClick := ppGilleOnClick;

  TFFiche(Ecran).OnAfterFormShow := AfterOnShow; {01/03/05 : FQ 10206}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.AfterOnShow;
{---------------------------------------------------------------------------------------}
begin
  SetFocusControl('TFR_CODEFRAIS');{01/03/05 : FQ 10206}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.OnNewRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  SetField('TFR_DEVISE', V_PGI.DevisePivot);
  SetFocusControl('TFR_CODEFRAIS');{01/03/05 : FQ 10206}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {Récupération du contenu de la grille des flux associés et remplissage du champ Memo
   stockant les valeurs}
  MajFluxAssocies;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.OnLoadRecord;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  s, c : string;
begin
  inherited;
  {Remplissage de la grille des flux associés à partir des valeurs stockées dans le
   champ Memo TFR_FLUXASSOCIE}
  if not Inserting then begin
    s := GetField('TFR_FLUXASSOCIE');
    c := ReadTokenSt(s);

    {S'il n'y a pas de flux associés, on vide la grille et on sort}
    if c = '' then begin
      GridFlux.VidePile(True);
      Exit;
    end;

    n := 1;
    while c <> '' do begin
      GridFlux.RowCount := n + 1;
      GridFlux.Cells[0, n] := c;
      GridFlux.Cells[1, n] := GetLibFluxRub(c);
      c := ReadTokenSt(s);
      Inc(n);
    end;
  end
  else
    GridFlux.VidePile(True);

  {Mise à jour du drapeau}
  DeviseChange(GetControl('TFR_DEVISE'));
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.OnChangeField(F : TField);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {Selon le mode de frais, certaines zones sont désactivées}
  if F.FieldName = 'TFR_MODEFRAIS' then
    ModeOnChange(GetControl('TFR_MODEFRAIS'))
  {Adaptation du format des zones de saisie des montants en fonction de la devise}
  else if F.FieldName = 'TFR_DEVISE' then
    GererFormat(GetControlText('TFR_DEVISE'));
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.GrilleOnKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin
  if (Key = VK_DOWN) and (Shift = []) and (GridFlux.Row = GridFlux.RowCount - 1) then begin
    AjouterLigne;
  end
  else if (Key = VK_INSERT) and (Shift = [ssCtrl]) then begin
    Key := 0;
    InsererUneLigne(GridFlux.Row + 1);
  end
  else if (Key = VK_DELETE) then begin
    Key := 0;
    if Shift = [ssCtrl, ssShift] then ViderGrille
                                 else SupprimerLaLigne(GridFlux.Row);
  end
  else if Key = VK_F5 then
    GrilleOnBtnClick(Sender);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.ModeOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if GetControlText('TFR_MODEFRAIS') = 'COM' then begin
    SetControlEnabled('TFR_COUTFORFAIT' , False);
    SetControlEnabled('TFR_COUTUNITAIRE', False);
    SetControlEnabled('TFR_COUTMIN'     , True);
    SetControlEnabled('TFR_COUTMAX'     , True);
    SetControlEnabled('TFR_COMMISSION'  , True);
    SetControlEnabled('TFR_TRANCHEFRAIS', False);
  end
  else if GetControlText('TFR_MODEFRAIS') = 'FOR' then begin
    SetControlEnabled('TFR_COUTFORFAIT' , True);
    SetControlEnabled('TFR_COUTUNITAIRE', True);
    SetControlEnabled('TFR_COUTMIN'     , False);
    SetControlEnabled('TFR_COUTMAX'     , False);
    SetControlEnabled('TFR_COMMISSION'  , False);
    SetControlEnabled('TFR_TRANCHEFRAIS', False);
  end
  else if GetControlText('TFR_MODEFRAIS') = 'TRA' then begin
    SetControlEnabled('TFR_COUTFORFAIT' , False);
    SetControlEnabled('TFR_COUTUNITAIRE', False);
    SetControlEnabled('TFR_COUTMIN'     , False);
    SetControlEnabled('TFR_COUTMAX'     , False);
    SetControlEnabled('TFR_COMMISSION'  , False);
    SetControlEnabled('TFR_TRANCHEFRAIS', True);
  end;
    SetControlEnabled('TTFR_COUTFORFAIT',  GetControlEnabled('TFR_COUTFORFAIT' ));
    SetControlEnabled('TTFR_COUTUNITAIRE', GetControlEnabled('TFR_COUTUNITAIRE'));
    SetControlEnabled('TTFR_COUTMIN',      GetControlEnabled('TFR_COUTMIN'     ));
    SetControlEnabled('TTFR_COUTMAX',      GetControlEnabled('TFR_COUTMAX'     ));
    SetControlEnabled('TTFR_COMMISSION',   GetControlEnabled('TFR_COMMISSION'   ));
    SetControlEnabled('TTFR_TRANCHEFRAIS', GetControlEnabled('TFR_TRANCHEFRAIS'));
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.ViderGrille;
{---------------------------------------------------------------------------------------}
begin
  if HShowMessage('1;' + Ecran.Caption + ';Êtes-vous sûr de vouloir vider la grille ?;Q;YNC;N;C;', '', '') = mrYes then begin
    GridFlux.VidePile(True);
    {19/09/05 : FQ 10292 : Si on n'est pas en modification ou insertion, le champ TFR_FLUXASSOCIE
                n'est pas mis à jour}
    if not (DS.State in [dsInsert, dsEdit]) then begin
      DS.Edit;
      MajFluxAssocies;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.AjouterLigne;
{---------------------------------------------------------------------------------------}
begin
  GridFlux.RowCount := GridFlux.RowCount + 1;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.SupprimerLaLigne(ARow : Integer);
{---------------------------------------------------------------------------------------}
begin
  if HShowMessage('1;' + Ecran.Caption + ';Êtes-vous sûr de vouloir supprimer la ligne en cours ?;Q;YNC;N;C;', '', '') = mrYes then begin
    GridFlux.DeleteRow(ARow);
    {19/09/05 : FQ 10292 : Si on n'est pas en modification ou insertion, le champ TFR_FLUXASSOCIE
                n'est pas mis à jour}
    if not (DS.State in [dsInsert, dsEdit]) then begin
      DS.Edit;
      MajFluxAssocies;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.InsererUneLigne(ARow : Integer);
{---------------------------------------------------------------------------------------}
begin
  {Si on est sur la ligne des titres}
  if ARow = 0 then Exit;
  GridFlux.InsertRow(ARow);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.ppGilleOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  case TMenuItem(Sender).MenuIndex of
    0 : InsererUneLigne(GridFlux.Row + 1);
    1 : AjouterLigne;
    3 : SupprimerLaLigne(GridFlux.Row);
    4 : ViderGrille;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.GererFormat(Dev : string);
{---------------------------------------------------------------------------------------}
var
  fmt : string;
begin
  fmt := StrFMask(CalcDecimaleDevise(Dev), '', True);
  {$IFDEF EAGLCLIENT}
  THEdit(GetControl('TFR_COUTFORFAIT' )).DisplayFormat := fmt;
  THEdit(GetControl('TFR_COUTUNITAIRE')).DisplayFormat := fmt;
  THEdit(GetControl('TFR_COUTMIN'     )).DisplayFormat := fmt;
  THEdit(GetControl('TFR_COUTMAx'     )).DisplayFormat := fmt;
  {$ELSE}
  THDBEdit(GetControl('TFR_COUTFORFAIT' )).DisplayFormat := fmt;
  THDBEdit(GetControl('TFR_COUTUNITAIRE')).DisplayFormat := fmt;
  THDBEdit(GetControl('TFR_COUTMIN'     )).DisplayFormat := fmt;
  THDBEdit(GetControl('TFR_COUTMAX'     )).DisplayFormat := fmt;
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.GrillOnCellExit (Sender : TObject; var ACol, ARow : Integer; var Cancel : Boolean);
{---------------------------------------------------------------------------------------}
begin
  if ACol = 0 then
    GridFlux.Cells[1, ARow] := GetLibFluxRub(GridFlux.Cells[ACol, ARow]);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.GrillOnCellEnter(Sender : TObject; var ACol, ARow : Integer; var Cancel : Boolean);
{---------------------------------------------------------------------------------------}
begin
  {Affichage de l'elipsis si on est sur la colonne 0}
//  GridFlux.ElipsisButton := ACol = 0;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.OnCellSelect(Sender: TObject; ACol, ARow: Longint; var CanSelect: Boolean);
{---------------------------------------------------------------------------------------}
begin
  {Affichage de l'elipsis si on est sur la colonne 0}
  GridFlux.ElipsisButton := ACol = 0;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.GrilleOnBtnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Chargement de la liste des flux de trésorerie
   TRLISTEFLUX est une vue faisant une jointure sur les tables RUBRIQUE et FLUXTRESO}
  LookUpList(GridFlux, 'Flux de trésorerie', 'TRLISTEFLUX', 'TFT_FLUX', 'TFT_LIBELLE',
             'TFT_CLASSEFLUX <> "' + cla_Commission + '"', 'TFT_FLUX', True, 0);

  if not (DS.State in [dsEdit, dsInsert]) then begin
    DS.Edit;
    {Mise à jour du champ mémo}
    MajFluxAssocies;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.MajFluxAssocies;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  s : string;
begin
  {Récupération du contenu de la grille des flux associés et remplissage du champ Memo
   stockant les valeurs}
  for n := 1 to GridFlux.RowCount - 1 do begin
    {Si le flux est renseigné ...}
    if GridFlux.Cells[0, n] <> '' then
      {... on le mémorise}
      s := s + GridFlux.Cells[0, n] + ';';
  end;

  SetField('TFR_FLUXASSOCIE', s);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.DeviseChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  AssignDrapeau(TImage(GetControl('IDEV')), GetControlText('TFR_DEVISE'));
end;

{---------------------------------------------------------------------------------------}
procedure TOM_FRAIS.GrilleOnEnter(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {19/09/05 : FQ 10292 : Pour exécuter le CellEnter en entrée de la Grille}
  PostMessage(GridFlux.Handle, WM_KEYDOWN, VK_TAB, 0);
  PostMessage(GridFlux.Handle, WM_KEYDOWN, VK_LEFT, 0);
end;

initialization
  RegisterClasses([TOM_FRAIS]);

end.

