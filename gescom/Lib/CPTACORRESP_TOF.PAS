{***********UNITE*************************************************
Auteur  ...... : JT
Créé le ...... : 28/07/2003
Modifié le ... :   /  /
Description .. : Saisie correspondance entre GC et Compta pour MODEPAIE
Suite ........ : et MODEREGL
Mots clefs ... : TOF;CPTACORRESP
*****************************************************************}
unit CPTACORRESP_TOF;

interface

uses StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
  db,
  dbtables,
  mul,
  Fe_Main,
  {$ELSE}
  MaineAGL,
  eMul,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  UTOB,
  Vierge,
  HTB97,
  VoirTob;

function GCLanceFiche_CorrespCompta(Nat, Cod: string; Range, Lequel, Argument: string): string;

type
  TOF_CPTACORRESP = class(TOF)
  private
    TableATraiter, CodeATraiter, ChampOrigine, LibOrigine, Titre: string;
    TobTableGC, TobCorresp: TOB;
    QryCorresp: TQuery;
    GLstCorresp: ThGrid;
    NColCpta: integer;
    procedure GLstCorrespOnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    function MajTob(ACol, ARow: Integer): boolean;
    procedure TousIdentiqueClick(Sender: TObject);

  public
    procedure OnArgument(S: string); override;
    procedure OnUpdate; override;
    procedure OnClose; override;
  end;

implementation

function GCLanceFiche_CorrespCompta(Nat, Cod: string; Range, Lequel, Argument: string): string;
begin
  Result := '';
  if Nat = '' then exit;
  if Cod = '' then exit;
  Result := AGLLanceFiche(Nat, Cod, Range, Lequel, Argument);
end;

{==============================================================================================}
{================================== Procédure de la TOF =======================================}
{==============================================================================================}

procedure TOF_CPTACORRESP.OnArgument(S: string);
var Cpt, Cpt1: integer;
  TobTmp, TobTmp1: TOB;
begin
  inherited;
  NColCpta := 2;
  TableATraiter := ReadTokenSt(S);
  CodeATraiter := ReadTokenSt(S);
  ChampOrigine := ReadTokenSt(S);
  LibOrigine := ReadTokenSt(S);
  if CodeAtraiter = 'MDR' then
    Titre := 'Modes de règlement'
  else if CodeAtraiter = 'MDP' then
    Titre := 'Modes de paiement';
  Ecran.Caption := TraduireMemoire('Correspondances des ') + Titre;
  UpdateCaption(Ecran);
  GLstCorresp := THGrid(GetControl('GLSTCORRESP'));
  {Tob des données GC}
  TobTableGC := TOB.Create(TableATraiter, nil, -1);
  TobTableGC.LoadDetailDB(TableATraiter, '', '', nil, False);
  {Tob des correspondances déjà saisies}
  TobCorresp := TOB.Create('CORRESP', nil, -1);
  QryCorresp := OpenSQL('SELECT * FROM CORRESP WHERE CR_TYPE="' + CodeAtraiter + '"', True);
  TobCorresp.LoadDetailDB('CORRESP', '', '', QryCorresp, False);
  Ferme(QryCorresp);
  {Ajout des nouveaux enreg des tables GC}
  for Cpt := 0 to TobTableGC.detail.count - 1 do
  begin
    TobTmp := TobTableGC.Detail[Cpt];
    TobTmp1 := TobCorresp.FindFirst(['CR_CORRESP'], [TobTmp.Getvalue(ChampOrigine)], True);
    if TobTmp1 = nil then
    begin
      TobTmp1 := TOB.Create('CORRESP', TobCorresp, -1);
      TobTmp1.AddChampSupValeur('CR_TYPE', CodeAtraiter, false);
      TobTmp1.AddChampSupValeur('CR_CORRESP', TobTmp.Getvalue(ChampOrigine), false);
      TobTmp1.AddChampSupValeur('CR_LIBELLE', '', false);
      TobTmp1.AddChampSupValeur('CR_ABREGE', TobTmp.Getvalue(LibOrigine), false);
      for Cpt1 := 1 to 5 do
        TobTmp1.AddChampSupValeur('CR_LIBRETEXTE' + IntTostr(Cpt1), '', false);
    end;
  end;
  FreeAndNil(TobTableGC);
  TToolBarButton97(GetControl('BINDENT')).OnClick := TousIdentiqueClick;
  {Format du Grid}
  GLstCorresp.OnCellExit := GLstCorrespOnCellExit;
  GLstCorresp.FixedCols := 1;
  GLstCorresp.FixedRows := 1;
  GLstCorresp.ColCount := 4;
  GLstCorresp.RowCount := TobCorresp.detail.count + 1;
  GLstCorresp.ColWidths[0] := 10;
  GLstCorresp.ColWidths[1] := 64;
  GLstCorresp.ColWidths[2] := 64;
  GLstCorresp.ColWidths[3] := 250;
  GLstCorresp.ColEditables[0] := False;
  GLstCorresp.ColEditables[1] := True;
  GLstCorresp.ColEditables[2] := False;
  GLstCorresp.ColEditables[3] := False;
  TFVierge(Ecran).HMTrad.ResizeGridColumns(GLstCorresp);
  for Cpt := 0 to TobCorresp.detail.count - 1 do
  begin
    TobTmp := TobCorresp.detail[Cpt];
    TobTmp.PutLigneGrid(GLstCorresp, Cpt + 1, False, False, ';CR_LIBELLE;CR_CORRESP;CR_ABREGE');
  end;
end;

procedure TOF_CPTACORRESP.OnUpdate;
begin
  inherited;
  if not MajTob(1, GLstCorresp.Row) then exit;
  TobCorresp.InsertOrUpdateDB(false);
end;

procedure TOF_CPTACORRESP.OnClose;
begin
  inherited;
  if TobCorresp <> nil then
    FreeAndNil(TobCorresp);
end;

{==============================================================================================}
{================================== Fin procédure de la TOF ===================================}
{==============================================================================================}

procedure TOF_CPTACORRESP.GLstCorrespOnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if Acol = 1 then
    MajTob(Acol, Arow);
end;

function TOF_CPTACORRESP.MajTob(ACol, ARow: Integer): boolean;
var ValeurCpt, ValeurGc: string;
    TobTmp: TOB;

  procedure Erreur(Msg: string);
  begin
    PGIBox(Msg, TraduireMemoire('Erreur'));
    GLstCorresp.Cellvalues[1, Arow] := TobCorresp.detail[Arow - 1].GetValue('CR_LIBELLE');
    GLstCorresp.Row := Arow;
    GLstCorresp.Col := 1;
  end;

begin
  Result := True;
  ValeurCpt := uppercase(GLstCorresp.Cellvalues[Acol, Arow]);
  ValeurGc := uppercase(GLstCorresp.Cellvalues[2, Arow]);
  if length(ValeurCpt) > 3 then
  begin
    Erreur(TraduireMemoire('La longueur ne doit pas dépasser 3 alaphanumériques'));
    Result := False;
    exit;
  end;
  if ValeurCpt <> '' then
  begin
    GLstCorresp.Cellvalues[1, Arow] := ValeurCpt;
    TobTmp := TobCorresp.FindFirst(['CR_CORRESP','CR_LIBELLE'], [ValeurCpt,ValeurGc], True);
    if (TobTmp <> nil) and (TobTmp.GetIndex <> Arow - 1) then
    begin
      Erreur(TraduireMemoire('Ce code existe déjà'));
      Result := False;
      exit;
    end;
  end;
  TobCorresp.detail[Arow - 1].PutValue('CR_LIBELLE', GLstCorresp.Cellvalues[1, Arow]);
end;

procedure TOF_CPTACORRESP.TousIdentiqueClick(Sender: TObject);
var Cpt: integer;
  TobTmp: TOB;
begin
  for Cpt := 0 to TobCorresp.detail.count - 1 do
  begin
    TobTmp := TobCorresp.detail[Cpt];
    TobTmp.PutValue('CR_LIBELLE', TobTmp.GetValue('CR_CORRESP'));
    GLstCorresp.Cellvalues[1, Cpt + 1] := TobTmp.GetValue('CR_CORRESP');
  end;
end;

initialization
  registerclasses([TOF_CPTACORRESP]);
end.
