unit eTempsDetail;

interface

uses Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  uTob,
  SaisiePL,
  Grids,
  Hctrls,
  HTB97,
  HMsgBox,
  ExtCtrls,
  HSysMenu, StdCtrls;

procedure ET_TempsDetail(const VisuResp: Boolean; const PDateDuJour: TDateTime; const PTypePlanning: TTypePlanning; const PT, Aff, Pres, Sal: TOB);

type
  TFicheTempDetail = class(TForm)
    Panel1: TPanel;
    BAide: TToolbarButton97;
    Bannuler: TToolbarButton97;
    G: THGrid;
    HMTrad: THSystemMenu;
    WithLibelle: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure WithLibelleClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  PRIVATE
    procedure Affiche;
    { Déclarations privées }
  PUBLIC
    { Déclarations publiques }
    DateDuJour: TDateTime;
    T: TOB;
    TypePlanning: TTypePLanning;
    LesAffaires: TOB;
    LesPrestations: TOB;
    LesSalaries: TOB;
    Responsable: Boolean;
  end;

implementation

{$R *.DFM}

procedure ET_TempsDetail(const VisuResp: Boolean; const PDateDuJour: TDateTime; const PTypePlanning: TTypePlanning; const PT, Aff, Pres, Sal: TOB);
begin
  with TFicheTempDetail.Create(Application) do
  begin
    LesAffaires := Aff;
    LesPrestations := Pres;
    LesSalaries := Sal;
    DateDuJour := PDateDuJour;
    T := PT;
    Responsable := VisuResp;
    TypePlanning := PTypePlanning;
    ShowModal;
    Free;
  end;
end;

procedure TFicheTempDetail.FormShow(Sender: TObject);
var
  C: Integer;
begin
  Caption := 'Détail de la journée du ' + FormatDateTime('dddd dd mmmm yyyy', DateDuJour);

  if Responsable then
  begin
    G.ColCount := 5;
    G.ColAligns[1] := taRightJustify;
    G.ColAligns[3] := taRightJustify;
  end
  else
  begin
    G.ColCount := 4;
    G.ColAligns[2] := taRightJustify;
  end;

  HMTrad.ResizeGridColumns(G);

  Affiche() ;

  if Responsable then
  begin
    C := G.ClientWidth div 10;
    G.ColWidths[0] := 2 * C;
    G.ColWidths[1] := 2 * C;
    G.ColWidths[2] := 2 * C;
    G.ColWidths[3] := C;
    G.ColWidths[4] := (3 * C) - 5;
  end
  else
  begin
    C := G.ClientWidth div G.ColCount;
    G.ColWidths[0] := C;
    G.ColWidths[1] := C;
    G.ColWidths[2] := C;
    G.ColWidths[3] := C - 4;
  end;
end;

procedure TFicheTempDetail.Affiche;
var
  i: Integer;
  TT: TOB;
  S: string;
begin
  if T.Detail.Count = 0 then G.RowCount := 2 else G.RowCount := 1 + T.Detail.Count;
  for i := 0 to T.Detail.Count - 1 do
  begin
    G.Objects[0, I + G.FixedRows] := T.Detail[i];
    if Responsable then
    begin
      G.Cells[0, 0] := 'Affaire';
      G.Cells[1, 0] := 'Ressource';
      G.Cells[2, 0] := 'Prestation';
      G.Cells[3, 0] := 'Qté';
      G.Cells[4, 0] := 'Libellé';

      if WithLibelle.Checked then
      begin
        G.ColAligns[1] := taLeftJustify;

        S := '';
        TT := LesAffaires.FindFirst(['AFF_AFFAIRE'], [T.Detail[i].GetValue('ACT_AFFAIRE')], False);
        if TT <> nil then S := TT.GetValue('AFF_LIBELLE');
        G.Cells[0, I + G.FixedRows] := S;

        S := '';
        if LesSalaries <> nil then
        begin
          TT := LesSalaries.FindFirst(['SALARIE'], [T.Detail[i].GetValue('ACT_RESSOURCE')], False);
          if TT <> nil then S := TT.GetValue('NOMPRENOM');
        end;
        G.Cells[1, I + G.FixedRows] := S;

        S := '';
        TT := LesPrestations.FindFirst(['GA_CODEARTICLE'], [T.Detail[i].GetValue('ACT_CODEARTICLE')], False);
        if TT <> nil then S := Copy(TT.GetValue('GA_CODEARTICLE'), 5, 200) + ' : ' + TT.GetValue('GA_LIBELLE');
        G.Cells[2, I + G.FixedRows] := S;
      end
      else
      begin
        G.Cells[0, I + G.FixedRows] := T.Detail[i].GetValue('ACT_AFFAIRE');
        G.Cells[1, I + G.FixedRows] := T.Detail[i].GetValue('ACT_RESSOURCE');
        G.Cells[2, I + G.FixedRows] := T.Detail[i].GetValue('ACT_CODEARTICLE');
      end;

      G.Cells[3, I + G.FixedRows] := '#COL#' + ColorToString(clGray) + ';' + FormatFloat('0.00', T.Detail[i].GetValue('Q'));
      G.Cells[4, I + G.FixedRows] := T.Detail[i].GetValue('L');
    end
    else
    begin
      G.Cells[0, 0] := 'Affaire';
      G.Cells[1, 0] := 'Prestation';
      G.Cells[2, 0] := 'Qté';
      G.Cells[3, 0] := 'Libellé';

      if WithLibelle.Checked then
      begin
        S := '';
        TT := LesAffaires.FindFirst(['AFF_AFFAIRE'], [T.Detail[i].GetValue('ACT_AFFAIRE')], False);
        if TT <> nil then S := TT.GetValue('AFF_LIBELLE');
        G.Cells[0, I + G.FixedRows] := S;

        S := '';
        TT := LesPrestations.FindFirst(['GA_CODEARTICLE'], [T.Detail[i].GetValue('ACT_CODEARTICLE')], False);
        if TT <> nil then S := Copy(TT.GetValue('GA_CODEARTICLE'), 5, 200) + ' : ' + TT.GetValue('GA_LIBELLE');
        G.Cells[1, I + G.FixedRows] := S;
      end
      else
      begin
        G.Cells[0, I + G.FixedRows] := T.Detail[i].GetValue('ACT_AFFAIRE');
        G.Cells[1, I + G.FixedRows] := T.Detail[i].GetValue('ACT_CODEARTICLE');
      end;

      G.Cells[0, I + G.FixedRows] := T.Detail[i].GetValue('ACT_AFFAIRE');
      G.Cells[1, I + G.FixedRows] := T.Detail[i].GetValue('ACT_CODEARTICLE');
      G.Cells[2, I + G.FixedRows] := '#COL#' + ColorToString(clGray) + ';' + FormatFloat('0.00', T.Detail[i].GetValue('Q'));
      G.Cells[3, I + G.FixedRows] := T.Detail[i].GetValue('L');
    end;
  end;
end;

procedure TFicheTempDetail.WithLibelleClick(Sender: TObject);
begin
  Affiche;
end;

procedure TFicheTempDetail.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = vk_escape then Bannuler.Click;
end;

end.

