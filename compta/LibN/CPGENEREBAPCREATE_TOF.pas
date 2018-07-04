{ Unité : Source TOF de la fiche CPGENEREBAPCOLL
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 7.01.001.001   08/02/06    JP     Création de l'unité
 7.01.001.011   08/06/06    JP     FQ 18320 Changement de la présentation de la grille des erreurs
--------------------------------------------------------------------------------------}
unit CPGENEREBAPCREATE_TOF;

interface
uses
  Controls, Classes, Vierge, 
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main, {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF} PrintDBG,
  {$ENDIF}
  Grids, SysUtils, HCtrls, UTob, HEnt1, UTOF;

type
  TOF_CPGENEREBAPCOLL = class(TOF)
    FListe   : THGrid;
    FListe1  : THGrid;
    LaListe  : TList;
    TobMasse : TOB;

    procedure OnArgument(S : string); override;
    procedure OnClose               ; override;
  private
    procedure ChargerGrille(Code : string);
  public
    procedure ImprimerClick(Sender : TObject);
    procedure ListeDblClick(Sender : TObject);
  end;

procedure CpLanceFiche_GenereBapColl;

implementation

uses
  HTB97, ULibBonAPayer, AglInit, HMsgBox, CPBONSAPAYER_TOF, Forms;

const
  ERR_JOURNAL      = 0;
  ERR_NUMEROPIECE  = 1;
  ERR_DATEECHEANCE = 2;
  ERR_ERREUR       = 3;

  COL_CODEVISA      = 0;
  COL_CIRCUITBAP    = 1;
  COL_ETABLISSEMENT = 2;
  COL_JOURNAL       = 3;
  COL_NUMEROPIECE   = 4;
  COL_DATEECHEANCE  = 5;
  COL_ECHEANCEBAP   = 6;
  COL_VISEUR1       = 7;
  COL_VISEUR2       = 8;


{---------------------------------------------------------------------------------------}
procedure CpLanceFiche_GenereBapColl;
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('CP', 'CPGENEREBAPCOLL', '', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREBAPCOLL.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 7509060;
  TToolbarButton97(GetControl('BIMPRIMER')).OnClick := ImprimerClick;
  {$IFDEF EAGLCLIENT}
  SetControlVisible('BIMPRIMER', False);
  {$ENDIF EAGLCLIENT}
  TobMasse := TFVierge(Ecran).LaTOF.LaTOB;
  LaListe  := TList(TheData);
  FListe   := THGrid(GetControl('FLISTE'));
  FListe.OnDblClick := ListeDblClick;
  FListe1  := THGrid(GetControl('FLISTE1'));
  ChargerGrille(S);
  if not Assigned(TobMasse) then begin
    Ecran.Caption := TraduireMemoire('Liste des erreurs');
    UpdateCaption(Ecran);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREBAPCOLL.OnClose;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREBAPCOLL.ChargerGrille(Code : string);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  r : Integer;
  p : PClefPiece;
begin
  r := 0;
  {Affichage des BAP créés}
  if Assigned(TobMasse) and (TobMasse.Detail.Count > 0) then begin
    FListe.ColFormats[COL_DATEECHEANCE ] := 'dd/mm/yyyy';
    FListe.ColFormats[COL_ECHEANCEBAP  ] := 'dd/mm/yyyy';
    
    FListe.ColAligns [COL_CODEVISA     ]  := taCenter;
    FListe.ColAligns [COL_CIRCUITBAP   ]  := taCenter;
    FListe.ColAligns [COL_ETABLISSEMENT]  := taCenter;
    FListe.ColAligns [COL_JOURNAL      ]  := taCenter;
    FListe.ColAligns [COL_NUMEROPIECE  ]  := taCenter;
    FListe.ColAligns [COL_DATEECHEANCE ]  := taCenter;
    FListe.ColAligns [COL_ECHEANCEBAP  ]  := taCenter;

    TobMasse.PutGridDetail(FListe, False, False, 'BAP_CODEVISA;BAP_CIRCUITBAP;BAP_ETABLISSEMENT;BAP_JOURNAL;BAP_NUMEROPIECE;BAP_DATEECHEANCE;BAP_ECHEANCEBAP;BAP_VISEUR1;BAP_VISEUR2;');
    for n := 1 to FListe.RowCount - 1 do begin
      FListe.Cells[COL_VISEUR1, n] := RechDom('TTUTILISATEUR', FListe.Cells[COL_VISEUR1, n], False);
      FListe.Cells[COL_VISEUR2, n] := RechDom('TTUTILISATEUR', FListe.Cells[COL_VISEUR2, n], False);
    end;
  end
  else begin
    FListe.Visible := False;
    FListe1.Align := alClient;
  end;

  if (LaListe.Count = 0) and (TobMasse.Detail.Count > 0) then
    FListe1.Visible := False
  else if (LaListe.Count = 0) then begin
    Fliste1.Cells[ERR_JOURNAL     , r] := '###';
    Fliste1.Cells[ERR_NUMEROPIECE , r] := '###';
    Fliste1.Cells[ERR_DATEECHEANCE, r] := '###';
    Fliste1.Cells[ERR_ERREUR      , r] := TraduireMemoire('Aucune écriture à traiter');
  end
  else begin
    {08/06/06 : FQ 18320 : Changement de libellé car c'est la date comptable qui est affichée et
                non la date d'échéance. Il est maintenant trop tard pour modifier la DBListe pour
                la V700 et rendre la date d'échéance obligatoire}
    Fliste1.Cells[ERR_DATEECHEANCE, 0 ] := 'Dt Comptable'; 
    FListe1.ColAligns[ERR_JOURNAL     ] := taCenter;
    FListe1.ColAligns[ERR_NUMEROPIECE ] := taCenter;
    FListe1.ColAligns[ERR_DATEECHEANCE] := taCenter;
    r := 1;
    {Affichage des erreurs}
    for n := 0 to LaListe.Count - 1 do begin
      p := LaListe.Items[n];
      Fliste1.Cells[ERR_JOURNAL     , r] := p^.E_JOURNAL;
      Fliste1.Cells[ERR_NUMEROPIECE , r] := IntToStr(p^.E_NUMEROPIECE);
      Fliste1.Cells[ERR_DATEECHEANCE, r] := DateToText(p^.E_DATECOMPTABLE);
      Fliste1.Cells[ERR_ERREUR      , r] := p^.E_QUALIFPIECE;
      r := Fliste1.RowCount;
      Fliste1.RowCount := r + 1;
    end;
    Fliste1.RowCount := Fliste1.RowCount - 1;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREBAPCOLL.ImprimerClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
{$IFNDEF EAGLCLIENT}
var
  R : TModalResult;
{$ENDIF EAGLCLIENT}
begin
  {$IFDEF EAGLCLIENT}

  {$ELSE}
  {Seule la liste des erreurs est visible}
  if not Fliste.Visible then begin
    if HShowMessage('0;' + Ecran.Caption + ';Voulez-vous imprimer la liste des erreurs générées ?;Q;YN;Y;N;', '', '') = mrYes then
      R := mrNo
    else
      R := mrCancel;
  end

  {Seule la liste des BAP est visible}
  else if not Fliste1.Visible then begin
    if HShowMessage('0;' + Ecran.Caption + ';Voulez-vous imprimer la liste des bons à payer générés ?;Q;YN;Y;N;', '', '') = mrYes then
      R := mrYes
    else
      R := mrCancel;
  end

  {Les deux lsites sont visibles}
  else
    R :=  HShowMessage('0;' + Ecran.Caption + ';Voulez-vous imprimer la liste des bons à payer générés (Oui)'#13 +
                                              'ou celle des erreurs (Non) ?;Q;YNC;Y;C;', '', '');
  case R of
    mrYes : PrintDBGrid(FListe , nil, 'Liste des bons à payer générés', '');
    mrNo  : PrintDBGrid(FListe1, nil, 'Liste des erreurs générées', '');
  end;
  {$ENDIF EAGLCLIENT}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREBAPCOLL.ListeDblClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  T := TobMasse.Detail[FListe.Row - 1];
  if Assigned(T) then
    CpLanceFiche_FicheBap(T.GetString('BAP_JOURNAL') + ',' + T.GetString('BAP_EXERCICE') + ',' +
                          T.GetString('BAP_DATECOMPTABLE') + ',' + T.GetString('BAP_NUMEROPIECE') + ',;'  );
end;

initialization
  RegisterClasses([TOF_CPGENEREBAPCOLL]);

end.
