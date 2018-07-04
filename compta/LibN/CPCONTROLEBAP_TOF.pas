{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
08.10.001.011  18/10/07   JP   Création de l'unité : contrôle de la cohérence entre la table
                               UTILISAT et les Tables CPCIRCUIT et CPBONSAPAYER
--------------------------------------------------------------------------------------}
unit CPCONTROLEBAP_TOF;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_MAIN,
  {$ENDIF}
  HCtrls, Classes, UTOF, uTob, Grids, Graphics;

type

  TOF_CPCONTROLEBAP = class (TOF)
    procedure OnArgument (S : string); override;
    procedure OnUpdate               ; override;
    procedure OnClose                ; override;
  private
    lUtilisat : TStringList;

    procedure InitGCircuit;
    procedure InitGBap;
    procedure ChargeUser;
    procedure MajBap;
    procedure MajCircuit;
    procedure RepareBap;
    function  TesteUser(C, R : Byte) : Boolean;
    function  RechUser (Code : string) : string;
  public
    GCircuit : THGrid;
    GBap     : THGrid;

    procedure BRepareClick(Sender : TObject);
    procedure ImprimeClick(Sender : TObject);
    procedure CirDbleClick(Sender : TObject);
    procedure BapDbleClick(Sender : TObject);
    procedure GCircuOnExit(Sender : TObject);
    procedure ViseCellExit(Sender : TObject; var ACol, ARow : Integer; var Cancel : Boolean);
    procedure VisCellEnter(Sender : TObject; var ACol, ARow : Integer; var Cancel : Boolean);
    procedure ViseuKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
    procedure ViseDrawCell(ACol, ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
    procedure GBapDrawCell(ACol, ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
  end;

procedure CpLanceFiche_ControlBap(var tCircuit, tBap : TOB);

implementation

uses
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  StdCtrls, SysUtils, HTB97, UObjEtats, Controls, HEnt1, CPCIRCUITBAP_TOF,
  CPBONSAPAYER_TOF, HMsgBox, ParamSoc, Windows, uLibBonAPayer, UProcGen;

type
  TObjUser = class
    Cod : string;
    lib : string;
  end;

var
  TobCircuit : TOB;
  TobBap     : TOB;

{---------------------------------------------------------------------------------------}
procedure CpLanceFiche_ControlBap(var tCircuit, tBap : TOB);
{---------------------------------------------------------------------------------------}
begin
  TobCircuit := tCircuit;
  TobBap     := tBap;
  AglLanceFiche('CP', 'CPCONTROLEBAP', '', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLEBAP.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  GCircuit := (GetControl('GCIRCUIT') as THGrid);
  GBap     := (GetControl('GBAP'    ) as THGrid);
  (GetControl('BREPARE'  ) as TToolbarButton97).OnClick := BRepareClick;
  (GetControl('BIMPRIMER') as TToolbarButton97).OnClick := ImprimeClick;
  GCircuit.OnCellExit   := ViseCellExit;
  GCircuit.OnDblClick   := CirDbleClick;
  GCircuit.OnCellEnter  := VisCellEnter;
  GCircuit.OnExit       := GCircuOnExit;
  GCircuit.OnKeyDown    := ViseuKeyDown;
  GCircuit.PostDrawCell := ViseDrawCell;
  GBap    .PostDrawCell := GBapDrawCell;
  GBap    .OnDblClick   := BapDbleClick;
  ChargeUser;
  InitGCircuit;
  InitGBap;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLEBAP.OnUpdate;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if GCircuit.Visible then begin
    MajCircuit;
    if GBap.Visible then MajBap;
  end
  else
    MajBap;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLEBAP.OnClose;
{---------------------------------------------------------------------------------------}
begin
  LibereListe(lUtilisat, True);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLEBAP.BRepareClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  RepareBap;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLEBAP.ImprimeClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  TObjEtats.GenereEtatGrille(GBap, TraduireMemoire('Liste des BAP avec viseurs erronés'));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLEBAP.InitGCircuit;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  F : TOB;
begin
  if (TobCircuit.Detail.Count = 0) then begin
    GCircuit.Visible := False;
    SetControlVisible('PNCIRCUIT', False);
    SetControlVisible('BREPARE'  , False);
  end
  else begin
    for n := 0 to TobCircuit.Detail.Count - 1 do begin
      F := TobCircuit.Detail[n];
      if F.GetString('CCI_LIBVISEUR1') = '' then F.SetString('CCI_VISEUR1', '');
      if F.GetString('CCI_LIBVISEUR2') = '' then F.SetString('CCI_VISEUR2', '');
    end;

    TobCircuit.PutGridDetail(GCircuit, False, False, '');
    GCircuit.ColEditables[0] :=  False;
    GCircuit.ColEditables[1] :=  False;
    GCircuit.ColEditables[2] :=  False;
    GCircuit.ColWidths [5] := -1;
    GCircuit.ColWidths [6] := -1;
    GCircuit.ColLengths[5] := -1;
    GCircuit.ColLengths[6] := -1;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLEBAP.InitGBap;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  T : TOB;
begin
  if (TobBap.Detail.Count = 0) then begin
    GBap.Visible := False;
    SetControlVisible('PNBAP'    , False);
    SetControlVisible('BIMPRIMER', False);
    SetControlVisible('BREPARE'  , False);
    GCircuit.Align := alClient;
  end
  else begin
    for n := 1 to TobBap.Detail.Count do 
      GBap.DeleteRow(n);

    GBap.RowCount := TobBap.Detail.Count + 1;
    for n := 0 to TobBap.Detail.Count - 1 do begin
      T := TobBap.Detail[n];
      GBap.Cells[0, n + 1] := T.GetString('BAP_JOURNAL');
      GBap.Cells[1, n + 1] := T.GetString('BAP_DATECOMPTABLE');
      GBap.Cells[2, n + 1] := T.GetString('BAP_NUMEROPIECE');
      GBap.Cells[3, n + 1] := T.GetString('BAP_ECHEANCEBAP');
      GBap.Cells[4, n + 1] := T.GetString('BAP_CIRCUITBAP');
      if T.GetString('BAP_LIBVISEUR1') <> '' then
        GBap.Cells[5, n + 1] := T.GetString('BAP_VISEUR1');
      if T.GetString('BAP_LIBVISEUR2') <> '' then
        GBap.Cells[6, n + 1] := T.GetString('BAP_VISEUR2');
      GBap.Objects[0, n + 1] := T;
    end;
  end;
end;


{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLEBAP.BapDbleClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  T := TOB(GBap.Objects[0, GBap.Row]);
  if not Assigned(T) or (T.GetString('BAP_EXERCICE') = '') then
    PgiInfo(TraduireMemoire('Enregistrement inaccessible'), Ecran.Caption)
  else
    CpLanceFiche_FicheBap(GBap.Cells[0, GBap.Row] + ',' + T.GetString('BAP_EXERCICE') + ',' +
                          GBap.Cells[1, GBap.Row] + ',' + GBap.Cells[2, GBap.Row] + ',;');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLEBAP.CirDbleClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  CpLanceFiche_CircuitFiche(GCircuit.Cells[0, GCircuit.Row], 'Z');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLEBAP.GCircuOnExit(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if not TesteUser(GCircuit.Col, GCircuit.Row) then
    GCircuit.SetFocus;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLEBAP.ViseCellExit(Sender : TObject; var ACol, ARow : Integer; var Cancel : Boolean);
{---------------------------------------------------------------------------------------}
begin
  if ACol in [3..4] then begin
    Cancel := not TesteUser(ACol, ARow);
    //if not Cancel then
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLEBAP.VisCellEnter(Sender : TObject; var ACol, ARow : Integer; var Cancel : Boolean);
{---------------------------------------------------------------------------------------}
begin
  if GCircuit.Col in [3..4] then
    GCircuit.ColEditables[GCircuit.Col] := GCircuit.Cells[GCircuit.Col, GCircuit.Row] = '';
  GCircuit.ElipsisButton := (GCircuit.Col in [3..4]);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLEBAP.ViseuKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin
  case Key of
    VK_F5 : LookUpUtilisateur(Sender);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLEBAP.ViseDrawCell(ACol, ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
{---------------------------------------------------------------------------------------}
var
  s : string;
  R : TRect;
begin
  if (ACol in [3..4]) and (GCircuit.CellValues[ACol, ARow] <> '') then begin
    s := RechUser(GCircuit.CellValues[ACol, ARow]);
    if s <> '' then begin
      R := GCircuit.CellRect(ACol, ARow);
      Canvas.FillRect(R);
      Canvas.TextRect(R, R.Left + 2, R.Top + 2, s);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLEBAP.GBapDrawCell(ACol, ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
{---------------------------------------------------------------------------------------}
var
  s : string;
  R : TRect;
begin
  if (ACol in [5..6]) and (GBap.CellValues[ACol, ARow] <> '') then begin
    s := RechUser(GBap.CellValues[ACol, ARow]);
    if s <> '' then begin
      R := GBap.CellRect(ACol, ARow);
      Canvas.FillRect(R);
      Canvas.TextRect(R, R.Left + 2, R.Top + 2, s);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPCONTROLEBAP.TesteUser(C, R : Byte) : Boolean;
{---------------------------------------------------------------------------------------}
var
  u : string;
begin
  u := GCircuit.CellValues[C, R];

  if (C in [3..4]) and (u <> '') then begin
    if lUtilisat.IndexOf(u) = -1 then begin
      PGIBox(TraduireMemoire('Ce viseur n''est pas autorisé.'), Ecran.Caption);
      Result := False;
    end
    else
      Result := True;
  end
  else
    Result := True;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPCONTROLEBAP.RechUser(Code : string) : string;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  Result := '';
  if Code = '' then Exit;
  n := lUtilisat.IndexOf(Code);
  if n > - 1 then Result := TObjUser(lUtilisat.Objects[n]).lib;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLEBAP.ChargeUser;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  O : TObjUser;
begin
  lUtilisat := TStringList.Create;
  Q := OpenSQL('SELECT US_UTILISATEUR, US_LIBELLE FROM UTILISAT WHERE (US_GRPSDELEGUES LIKE "%' + GetParamSocSecur('SO_CPGROUPEVISEUR', '') +
               ';%" OR US_GRPSDELEGUES = "" OR US_GRPSDELEGUES IS NULL OR US_GRPSDELEGUES LIKE "<<%>>") AND (US_EMAIL <> "" AND US_EMAIL IS NOT NULL)', True);
  try
    while not Q.EOF do begin
      O := TObjUser.Create;
      O.Cod := Q.FindField('US_UTILISATEUR').AsString;
      O.Lib := Q.FindField('US_LIBELLE').AsString;
      lUtilisat.AddObject(Q.FindField('US_UTILISATEUR').AsString, O);
      Q.Next;
    end;
  finally
    Ferme(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLEBAP.MajBap;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  p : Integer;
  S : string;
  T : TOB;
begin
  if GBap.Visible then begin
    p := 0;
    for n := 0 to TobBap.Detail.Count - 1 do begin
      S := '';
      T := TobBap.Detail[n];

      if (T.GetString('BAP_MODIF') = '-') then Continue;
      
      if (T.GetString('BAP_VISEUR1') = '') and (T.GetString('BAP_VISEUR2') = '') then Continue;

      if (T.GetString('BAP_VISEUR1') <> '') then
        S := 'BAP_VISEUR1 = "' + T.GetString('BAP_VISEUR1') + '" ';
      if (T.GetString('BAP_VISEUR2') <> '') then begin
        if S <> '' then S := S + ', ';
        S := S + 'BAP_VISEUR2 = "' + T.GetString('BAP_VISEUR2') + '" ';
      end;

      S := 'UPDATE CPBONSAPAYER set ' + S;
      S := S + 'WHERE BAP_JOURNAL = "' + T.GetString('BAP_JOURNAL') +
               '" AND BAP_EXERCICE = "' + T.GetString('BAP_EXERCICE') +
               '" AND BAP_DATECOMPTABLE = "' + UsDateTime(T.GetDateTime('BAP_DATECOMPTABLE')) +
               '" AND BAP_NUMEROPIECE = ' + T.GetString('BAP_NUMEROPIECE') +
               '  AND BAP_NUMEROORDRE = ' + T.GetString('BAP_NUMEROORDRE');
      ExecuteSQL(S);
      Inc(p);
    end;

    if p = 0 then PGIInfo(TraduireMemoire('Aucun bon à payer n''a été mis à jour'), Ecran.Caption)
    else if p = 1 then PGIInfo(TraduireMemoire('Un bon à payer a été mis à jour'), Ecran.Caption)
    else if p > 1 then PGIInfo(IntToStr(p) + TraduireMemoire(' bons à payer ont été mis à jour'), Ecran.Caption);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLEBAP.MajCircuit;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  p : Integer;
  S : string;
begin
  inherited;
  if GCircuit.Visible then begin
    p := 0;
    for n := 1 to GCircuit.RowCount - 1 do begin
      S := '';
      if (GCircuit.CellValues[3, n] = '') and (GCircuit.CellValues[4, n] = '') then Continue;
      if (GCircuit.CellValues[3, n] <> '') then
        S := 'CCI_VISEUR1 = "' + GCircuit.CellValues[3, n] + '" ';
      if (GCircuit.Cells[4, n] <> '') then begin
        if S <> '' then S := S + ', ';
        S := S + 'CCI_VISEUR2 = "' + GCircuit.CellValues[4, n] + '" ';
      end;

      S := 'UPDATE CPCIRCUIT SET ' + S;
      S := S + 'WHERE CCI_CIRCUITBAP = "' + GCircuit.Cells[0, n] + '" AND CCI_NUMEROORDRE = ' + GCircuit.Cells[2, n];
      ExecuteSQL(S);
      Inc(p);
    end;

    if p = 0 then PGIInfo(TraduireMemoire('Aucun circuit n''a été mis à jour'), Ecran.Caption)
    else if p = 1 then PGIInfo(TraduireMemoire('Un circuit a été mis à jour'), Ecran.Caption)
    else if p > 1 then PGIInfo(IntToStr(p) + TraduireMemoire(' circuits ont été mis à jour'), Ecran.Caption);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPCONTROLEBAP.RepareBap;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  p : Integer;
  C : TOB;
  T : TOB;
  F : TOB;
begin
  T := TOB.Create('ffff', nil, -1);
  try
    T.LoadDetailFromSQL('SELECT CCI_CIRCUITBAP, CCI_NUMEROORDRE, CCI_VISEUR1, CCI_VISEUR2 FROM CPCIRCUIT');
    {Mise à jour des circuits}
    for n := 0 to GCircuit.RowCount - 1 do begin
      C := T.FindFirst(['CCI_CIRCUITBAP', 'CCI_NUMEROORDRE'], [GCircuit.Cells[0, n], GCircuit.Cells[2, n]], True);

      if not Assigned(C) then Continue;

      C.SetString('CCI_VISEUR1', GCircuit.CellValues[3, n]);
      C.SetString('CCI_VISEUR2', GCircuit.CellValues[4, n]);
    end;

    p := 0;
    {Mise à jour des BAP}
    for n := 0 to TobBap.Detail.Count - 1 do begin
      F := TobBap.Detail[n];
      C := T.FindFirst(['CCI_CIRCUITBAP', 'CCI_NUMEROORDRE'], [F.GetString('BAP_CIRCUITBAP'), F.GetString('BAP_NUMEROORDRE')], True);
      if not Assigned(C) then Continue;

      if C.GetString('CCI_VISEUR1') <> '' then begin
        F.SetString('BAP_VISEUR1', C.GetString('CCI_VISEUR1'));
        {Uniquement pour l'affichage}
        F.SetString('BAP_LIBVISEUR1', C.GetString('CCI_VISEUR1'));
      end;

      if C.GetString('CCI_VISEUR2') <> '' then begin
        F.SetString('BAP_VISEUR2', C.GetString('CCI_VISEUR2'));
        {Uniquement pour l'affichage}
        F.SetString('BAP_LIBVISEUR2', C.GetString('CCI_VISEUR2'));
      end;

      if (C.GetString('CCI_VISEUR1') <> '') and (C.GetString('CCI_VISEUR2') <> '') then begin
        Inc(p);
        F.SetString('BAP_MODIF', 'X');
      end;
    end;

    if p > 0 then InitGBap;

    if p = 0 then PGIInfo(TraduireMemoire('Aucun bon à payer n''a été modifié'), Ecran.Caption)
    else if p = 1 then PGIInfo(TraduireMemoire('Un bon à payer a été modifié'), Ecran.Caption)
    else if p > 1 then PGIInfo(IntToStr(p) + TraduireMemoire(' bons à payer ont été modifiés'), Ecran.Caption);
  finally
    FreeAndNil(T);
  end;
end;

initialization
  RegisterClasses([TOF_CPCONTROLEBAP]);


end.
