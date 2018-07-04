{--------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
09.00.000.001  08/01/08   JP   FQ 22146 : Création de l'unité pour la gestion de RB_NODOSSIER 

--------------------------------------------------------------------------------------}
unit CPVERIFRUBDOSSIER_TOF;

interface

uses
  Controls, Classes, UTOF, HCtrls, UTob;

type
  TOF_CPVERIFRUBDOSSIER  = class(TOF)
    FListe : THGrid;
    FRub   : TOB;
    
    procedure OnArgument(S : string); override;
    procedure OnClose               ; override;
    procedure OnUpdate              ; override;
    procedure OnDelete              ; override;
    procedure FListeOnDblClick(Sender : TObject);
    procedure FicheCloseQuery (Sender : TObject; var CanClose : Boolean);
  private
    CanExit     : Boolean;
    HasDoublons : Boolean;
    procedure InitFiche;
  end;

procedure CPLanceFiche_CPVERIFRUBDOSSIER(Arg : string);

implementation

uses
  {$IFDEF EAGLCLIENT}
   MaineAGL,
  {$ELSE}
   FE_Main,
//   {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  RUBRIQUE_TOM, HEnt1, HMsgBox, SysUtils, Vierge, UtilPgi;

const
  cl_Rubrique  = 0;
  cl_Libelle   = 1;
  cl_Nodossier = 2;
  cl_DateVal   = 3;
  cl_Predefini = 4;
  cl_NatRub    = 5;



{---------------------------------------------------------------------------------------}
procedure CPLanceFiche_CPVERIFRUBDOSSIER(Arg : string);
{---------------------------------------------------------------------------------------}
var
  SQL : string;
begin
  SQL := 'SELECT RB_RUBRIQUE FROM RUBRIQUE WHERE RB_NODOSSIER <> "' + V_PGI.NoDossier + '"';
  if CtxTreso in V_PGI.PGIContexte then SQL := SQL + ' AND RB_NATRUB = "TRE"'
                                   else SQL := SQL + ' AND RB_NATRUB <> "TRE"';

  {S'il y a des rb_nodossier qui ne sont pas à jour, on ouvre la fiche}
  if ExisteSQL(SQL) then
    AglLanceFiche('CP', 'CPVERIFRUBDOSSIER', '', '', Arg)
  else
    PgiInfo(TraduireMemoire('La table des rubriques ne nécessite aucune réparation.'));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVERIFRUBDOSSIER.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  FListe  := (GetControl('FLISTE') as THGrid);
  FRub    := TOB.Create('_RUB', nil, -1);
  CanExit := True;
  HasDoublons := False;
  TFVierge(Ecran).OnCloseQuery := FicheCloseQuery;
  FListe.OnDblClick := FListeOnDblClick;

  {Si on n'est pas en Tréso, Gestion de la colonne Nature de Rubrique}
  if not (CtxTreso in V_PGI.PGIContexte) then begin
    FListe.ColCount := cl_NatRub + 1;
    FListe.Cells[cl_NatRub, 0] := 'Nature';
    FListe.ColWidths[cl_Rubrique ] := FListe.ColWidths[cl_Rubrique ] - 15;
    FListe.ColWidths[cl_Libelle  ] := FListe.ColWidths[cl_Libelle  ] - 22;
    FListe.ColWidths[cl_Nodossier] := FListe.ColWidths[cl_Nodossier] - 3;
    FListe.ColWidths[cl_DateVal  ] := FListe.ColWidths[cl_DateVal  ] - 1;
    FListe.ColWidths[cl_Predefini] := FListe.ColWidths[cl_Predefini] - 1;
    FListe.ColWidths[cl_NatRub   ] := 42;
  end;

  {Vérification des corrections à faires}
  InitFiche;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVERIFRUBDOSSIER.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(FRub) then FreeAndNil(FRub);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVERIFRUBDOSSIER.OnDelete;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  S : string;
begin
  inherited;
  if FListe.NbSelected = 0 then
    PGIBox(TraduireMemoire('Veuillez sélectionner les rubriques à supprimer.'), Ecran.Caption)
  else begin
    for n := 0 to FListe.NbSelected - 1 do begin
      FListe.GotoLeBookmark(n);
      if (FListe.Cells[cl_Rubrique, FListe.Row] = '') and not isMssql then
        S := 'RB_NODOSSIER IS NULL AND '
      else
        S := 'RB_NODOSSIER = "' + FListe.Cells[cl_Nodossier, FListe.Row] + '" AND ';

      S := 'DELETE FROM RUBRIQUE WHERE RB_RUBRIQUE = "' + FListe.Cells[cl_Rubrique, FListe.Row] + '" AND ' + S;
      S := S + 'RB_DATEVALIDITE = "' + UsDateTime(StrToDateTime(FListe.Cells[cl_DateVal, FListe.Row])) + '" AND ';
      S := S + 'RB_PREDEFINI = "' + FListe.Cells[cl_Predefini, FListe.Row] + '" AND ';
      ExecuteSQL(S);
    end;

    InitFiche;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVERIFRUBDOSSIER.OnUpdate;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if HasDoublons then begin
    PGIBox(TraduireMemoire('Veuillez supprimer les rubriques en doublon.'), Ecran.Caption);
    CanExit := False;
  end
  else begin
    ExecuteSQL('UPDATE RUBRIQUE SET RB_NODOSSIER = "' + V_PGI.NoDossier + '"');
    CanExit := True;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVERIFRUBDOSSIER.InitFiche;
{---------------------------------------------------------------------------------------}
var
  SQL : string;
begin
  if not (CtxTreso in V_PGI.PGIContexte) then SQL := ', R1.RB_NATRUB';

  SQL := 'SELECT R1.RB_RUBRIQUE, R1.RB_LIBELLE, RB_NODOSSIER, R1.RB_DATEVALIDITE, R1.RB_PREDEFINI' + SQL + ' FROM RUBRIQUE R1 WHERE ';
  SQL := SQL + 'EXISTS(SELECT RB_RUBRIQUE FROM RUBRIQUE R2 WHERE R2.RB_PREDEFINI = R1.RB_PREDEFINI AND R2.RB_RUBRIQUE = ';
  SQL := SQL + 'R1.RB_RUBRIQUE  AND R2.RB_DATEVALIDITE = R1.RB_DATEVALIDITE AND R2.RB_NODOSSIER <> R1.RB_NODOSSIER)';
  if CtxTreso in V_PGI.PGIContexte then SQL := SQL + ' AND RB_NATRUB = "TRE"'
                                   else SQL := SQL + ' AND RB_NATRUB <> "TRE"';

  FListe.VidePile(False);
  FListe.ClearSelected;
  FRub.ClearDetail;
  FRub.LoadDetailFromSQL(SQL);
  if FRub.Detail.Count > 0 then begin
    SetControlCaption('LBLIB', TraduireMemoire('Avant de mettre à jour les numéros de dossier, veuillez supprimer les doublons.'));
    HasDoublons := True;
  end
  else begin
    HasDoublons := False;
    SQL := '';
    if not (CtxTreso in V_PGI.PGIContexte) then SQL := ', RB_NATRUB';
    SQL := 'SELECT RB_RUBRIQUE, RB_LIBELLE, RB_NODOSSIER, RB_DATEVALIDITE, RB_PREDEFINI' + SQL + ' FROM RUBRIQUE WHERE RB_NODOSSIER <> "' + V_PGI.NoDossier + '"';
    if CtxTreso in V_PGI.PGIContexte then SQL := SQL + ' AND RB_NATRUB = "TRE"'
                                     else SQL := SQL + ' AND RB_NATRUB <> "TRE"';
    FRub.LoadDetailFromSQL(SQL);
    SetControlVisible('BDELETE', False);
    SetControlCaption('LBLIB', TraduireMemoire('Il n''y a pas de doublon. Vous pouvez mettre à jour les numéros de dossier.'));
  end;

  if FRub.Detail.Count > 0 then begin
    FRub.PutGridDetail(FListe, False, True, '');
    FListe.ColAligns[cl_Rubrique ] := taLeftJustify; {RUBRIQUE}
    FListe.ColAligns[cl_Libelle  ] := taLeftJustify; {LIBELLE}
    FListe.ColAligns[cl_Nodossier] := taLeftJustify; {NODOSSIER}
    FListe.ColAligns[cl_DateVal  ] := taCenter; {DATEVALIDITE}
    FListe.ColAligns[cl_Predefini] := taCenter; {PREDEFINI}
    if not (CtxTreso in V_PGI.PGIContexte) then FListe.ColAligns[cl_NatRub] := taCenter; {NATRUB}
  end
  else begin
    FListe.RowCount := 1;
    FListe.Cells[cl_Rubrique , 1] := '';
    FListe.Cells[cl_Libelle  , 1] := '';
    FListe.Cells[cl_Nodossier, 1] := '';
    FListe.Cells[cl_DateVal  , 1] := '';
    FListe.Cells[cl_Predefini, 1] := '';
    if not (CtxTreso in V_PGI.PGIContexte) then FListe.Cells[cl_NatRub, 1] := '';
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVERIFRUBDOSSIER.FicheCloseQuery(Sender: TObject; var CanClose: Boolean);
{---------------------------------------------------------------------------------------}
begin
  CanClose := CanExit;
  CanExit  := True;
  if CanClose then TFVierge(Ecran).FormCloseQuery(Sender, CanClose);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPVERIFRUBDOSSIER.FListeOnDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  FCtxSurRubrique: TCtxSurRubrique;
begin
  if FListe.Cells[cl_NatRub, FListe.Row] = 'BUD' then FCtxSurRubrique := CtxBudget
                                                 else FCtxSurRubrique := CtxRubrique;
  ParametrageRubrique(FListe.Cells[cl_Rubrique , FListe.Row] + ';' +
                      FListe.Cells[cl_Nodossier, FListe.Row] + ';' +
                      FListe.Cells[cl_Predefini, FListe.Row] + ';' +
                      FListe.Cells[cl_DateVal  , FListe.Row], taConsult, FCtxSurRubrique);
end;


initialization
  RegisterClasses([TOF_CPVERIFRUBDOSSIER]);

end.
