{***********UNITE*************************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 06/12/2001
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : TRANCHEFRAIS (TRANCHEFRAIS)
Mots clefs ... : TOM;TRANCHEFRAIS
*****************************************************************}
unit TomTrancheFrais ;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL, UtileAgl,
  {$ELSE}
  db, FE_Main, EdtREtat,  
  {$ENDIF}
  Controls, Classes, Forms, SysUtils,
  HCtrls, HEnt1, UTOM, UTob, Menus;

type
  TOM_TRANCHEFRAIS = class (TOM)
    procedure OnAfterUpdateRecord      ; override;
    procedure OnUpdateRecord           ; override;
    procedure OnLoadRecord             ; override;
    procedure OnDeleteRecord           ; override;
    procedure OnChangeField(F : TField); override;
    procedure OnArgument   (S : string); override;
    procedure OnClose                  ; override;  
  protected
    TobDetail : TOB;
    GrdDetail : THGrid;
    PopupMenu : TPopUpMenu;
    PasTouche : Boolean;

    {Gestion de la grille détail}
    procedure ViderGrille      ;
    procedure AjouterLigne     ;
    procedure SupprimerLaLigne(ARow : Integer);
    procedure InsererUneLigne (ARow : Integer);

    procedure ImprimerOnClick (Sender : TObject);
    procedure ppGilleOnClick  (Sender : TObject);
    procedure GrilleOnKeyDown (Sender : TObject; var Key : Word; Shift : TShiftState);
    procedure GrilleOnCellExit(Sender : TObject; var ACol, ARow : Integer; var Cancel : Boolean);
  end;

procedure TRLanceFiche_TrancheFrais(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  Windows, HMsgBox, UProcGen, HTB97;

const
  COL_DE  = 0;
  COL_A   = 1;
  COL_MNT = 2;
  COL_PCT = 3;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_TrancheFrais(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRANCHEFRAIS.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
var
  n : Byte;
begin
  inherited;
  Ecran.HelpContext := 150;

  TobDetail := TOB.Create('$DETAIL', nil, -1);
  GrdDetail := THGrid(GetControl('GRIDDETAIL'));
  GrdDetail.OnKeyDown   := GrilleOnKeyDown;
  GrdDetail.OnCellExit  := GrilleOnCellExit;

  {Format des cellules de la grille}
  for n := 0 to COL_PCT do begin
    GrdDetail.ColTypes [n] := 'R' ;
    GrdDetail.ColAligns[n] := taRightJustify ;
  end;
  GrdDetail.ColFormats[COL_DE ] := StrFMask(2, '', True);
  GrdDetail.ColFormats[COL_A  ] := StrFMask(2, '', True);
  GrdDetail.ColFormats[COL_MNT] := StrFMask(2, '', True);
  GrdDetail.ColFormats[COL_PCT] := StrFMask(4, '', True);

  PopupMenu := TPopupMenu(GetControl('PPGRILLE'));
  for n := 0 to PopupMenu.Items.Count - 1 do
    PopupMenu.Items[n].OnClick := ppGilleOnClick;

  TToolbarButton97(GetControl('BIMPRIMER')).OnClick := ImprimerOnClick;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRANCHEFRAIS.OnAfterUpdateRecord;
{---------------------------------------------------------------------------------------}
var
  n, p : Integer;
  SQL  : string;
begin
  inherited;
  BeginTrans;
  try
    ExecuteSQL('DELETE FROM TRTRANCHEDETAIL WHERE TTD_TRANCHEFRAIS = "' + GetField('TTF_CODETRANCHE') + '"');
    p := 1;
    for n := 1 to GrdDetail.RowCount - 1 do begin
      if (GrdDetail.Cells[COL_DE, n] = '') and (GrdDetail.Cells[COL_DE, n] = '') and
         (GrdDetail.Cells[COL_MNT, n] = '') and (GrdDetail.Cells[COL_PCT, n] = '') then Break;
      SQL := 'INSERT INTO TRTRANCHEDETAIL (TTD_BORNEINF, TTD_BORNESUP, TTD_MONTANT, TTD_POURCENTAGE, ' +
             'TTD_NUMTRANCHE, TTD_TRANCHEFRAIS) VALUES (' + VarStrToPoint(GrdDetail.Cells[COL_DE, n]) + ', ' +
             VarStrToPoint(GrdDetail.Cells[COL_A, n]) + ', ' + VarStrToPoint(GrdDetail.Cells[COL_MNT, n]) + ', ' +
             VarStrToPoint(GrdDetail.Cells[COL_PCT, n]) + ', ' + IntToStr(p) + ',"' + GetField('TTF_CODETRANCHE') + '")';
      ExecuteSQL(SQL);
      Inc(p);
    end;
    CommitTrans;
  except
    on E : Exception do begin
      RollBack;
      HShowMessage('0;' + Ecran.Caption + ';L''écriture du détail des tranches ne s''est pas correctement effectuée;E;O;O;O;', '', '');
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRANCHEFRAIS.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if GetField('TTF_TYPECALCFRAIS') = '' then begin
    LastError    := 1;
    LastErrorMsg := TraduireMemoire('Veuillez renseigner le type de calcul.');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRANCHEFRAIS.OnLoadRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  GrdDetail.VidePile(True);
  GrdDetail.RowCount := 2;
  {Si on n'est pas en insertion, on charge le détail des tranches}
  if not AfterInserting then begin
    TobDetail.LoadDetailFromSQL('SELECT TTD_BORNEINF, TTD_BORNESUP, TTD_MONTANT, TTD_POURCENTAGE FROM ' +
                                'TRTRANCHEDETAIL WHERE TTD_TRANCHEFRAIS = "' + GetField('TTF_CODETRANCHE') + '"');
    TobDetail.PutGridDetail(GrdDetail, False, False, '');
  end;
  PasTouche := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRANCHEFRAIS.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(TobDetail) then FreeAndNil(TobDetail);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRANCHEFRAIS.OnDeleteRecord;
{---------------------------------------------------------------------------------------}
begin
  {On s'assure que la tranche en cours n'est pas utilisée par un frais}
  if ExisteSQL('SELECT TFR_TRANCHEFRAIS FROM FRAIS WHERE TFR_TRANCHEFRAIS = "' + GetField('TTF_CODETRANCHE') + '"') then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('La tranche est utilisée dans un tarif : sa suppression est impossible.');
  end;
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRANCHEFRAIS.OnChangeField(F : TField);
{---------------------------------------------------------------------------------------}
begin
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRANCHEFRAIS.GrilleOnKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin
  if (Key = VK_DOWN) and (Shift = []) and (GrdDetail.Row = GrdDetail.RowCount - 1) then begin
    AjouterLigne;
  end
  else if (Key = VK_INSERT) and (ssCtrl in Shift) and (ssShift in Shift)  then begin
    Key := 0;
    InsererUneLigne(GrdDetail.Row);
  end
  else if (Key = VK_DELETE) then begin
    Key := 0;
    if (ssCtrl in Shift) and (ssShift in Shift) then
      ViderGrille
    else if ssShift in Shift then
      SupprimerLaLigne(GrdDetail.Row);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRANCHEFRAIS.GrilleOnCellExit(Sender : TObject; var ACol, ARow : Integer; var Cancel : Boolean);
{---------------------------------------------------------------------------------------}
begin
  if not (DS.State in [dsInsert, dsEdit]) then DS.Edit;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRANCHEFRAIS.ViderGrille;
{---------------------------------------------------------------------------------------}
begin
  if HShowMessage('1;' + Ecran.Caption + ';Êtes-vous sûr de vouloir vider la grille ?;Q;YNC;N;C;', '', '') = mrYes then begin
    GrdDetail.VidePile(True);
    GrdDetail.RowCount := 2;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRANCHEFRAIS.AjouterLigne;
{---------------------------------------------------------------------------------------}
begin
  GrdDetail.RowCount := GrdDetail.RowCount + 1;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRANCHEFRAIS.SupprimerLaLigne(ARow : Integer);
{---------------------------------------------------------------------------------------}
begin
  if HShowMessage('1;' + Ecran.Caption + ';Êtes-vous sûr de vouloir supprimer la ligne en cours ?;Q;YNC;N;C;', '', '') = mrYes then
    GrdDetail.DeleteRow(ARow);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRANCHEFRAIS.InsererUneLigne(ARow : Integer);
{---------------------------------------------------------------------------------------}
begin
  {Si on est sur la ligne des titres}
  if ARow = 0 then Exit;
  GrdDetail.InsertRow(ARow);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRANCHEFRAIS.ppGilleOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  case TMenuItem(Sender).MenuIndex of
    0 : InsererUneLigne(GrdDetail.Row);
    1 : AjouterLigne;
    3 : SupprimerLaLigne(GrdDetail.Row);
    4 : ViderGrille;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TRANCHEFRAIS.ImprimerOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  LanceEtat('E', 'ECT', 'TRF', True, False, False, nil, '', 'TRANCHE DE FRAIS', False);
end;

initialization
  RegisterClasses([TOM_TRANCHEFRAIS]);

end.

