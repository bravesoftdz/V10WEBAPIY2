unit Reassort;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Facture, Menus, hmsgbox, HSysMenu, StdCtrls, ComCtrls, HRichEdt,
  HRichOLE, HTB97, Buttons, Grids, Hctrls, HFLabel, Mask, ExtCtrls, HPanel,
  FactUtil, HEnt1, UIUtil, EntGC, UTOB, Ent1, FactComm, M3FP, ImgList, utilpgi;

function CreerCommandeReassort(NaturePiece: string): boolean;
procedure SaisieCommandeReassort(CleDoc: R_CleDoc; Action: TActionFiche);

type
  TFReassort = class(TFFacture)
    PEnteteReassort: THPanel;
    FTitreReassort: THLabel;
    HGP_NUMEROPIECE_: THLabel;
    GP_NUMEROPIECE_: THPanel;
    HGP_DATEPIECE_: THLabel;
    GP_DATEPIECE_: THCritMaskEdit;
    HGP_REFINTERNE_: THLabel;
    GP_REFINTERNE_: TEdit;
    GP_BLOCNOTE: THRichEditOLE;
    HLabel1: THLabel;
    HLabel2: THLabel;
    GP_DATELIVRAISON_: THCritMaskEdit;
    HGP_DATELIVRAISON_: THLabel;

    procedure FormShow(Sender: TObject);
    procedure GSEnter(Sender: TObject);
    procedure GP_DATELIVRAISONExit(Sender: TObject);
    procedure GereTiersEnabled; override;

  protected
    { Déclarations privées }
    procedure InitPieceCreation; override;
    procedure ReInitPiece; override;

  public
    { Déclarations publiques }
    procedure ClickValide(EnregSeul: Boolean = False); override;

  end;

var
  FReassort: TFReassort;

implementation
uses ParamSoc, FactAdresse, FactTOB;

function CreerCommandeReassort(NaturePiece: string): boolean;
var CleDoc: R_CleDoc;
  T_Auxi: string;
  IsTiersValide: Boolean;
begin
  Result := False;
  T_Auxi := GetParamsoc('SO_GCTIERSREASSORT');
  IsTiersValide := not (T_Auxi = '');
  if IsTiersValide then IsTiersValide := ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="' + T_Auxi + '"');
  if not IsTiersValide then
  begin
    PGIError('Vous devez renseigner un tiers valide dans les paramètres sociétés.', FReassort.Caption);
    exit;
  end;
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  CleDoc.NaturePiece := NaturePiece;
  CleDoc.DatePiece := V_PGI.DateEntree;
  CleDoc.Souche := '';
  CleDoc.NumeroPiece := 0;
  CleDoc.Indice := 0;
  SaisieCommandeReassort(CleDoc, taCreat);
  Result := True;
end;

procedure SaisieCommandeReassort(CleDoc: R_CleDoc; Action: TActionFiche);
var X: TFReassort;
  PP: THPanel;
begin
  SourisSablier;
  PP := FindInsidePanel;
  X := TFReassort.Create(Application);
  X.CleDoc := CleDoc;
  X.Action := Action;
  X.NewNature := X.CleDoc.NaturePiece;
  X.TransfoPiece := False;
  X.DuplicPiece := False;
  if PP = nil then
  begin
    try
      X.ShowModal;
    finally
      X.Free;
    end;
    SourisNormale;
  end
  else
  begin
    InitInside(X, PP);
    X.Show;
  end;
end;

procedure TFReassort.InitPieceCreation;
begin
  inherited;
  GP_DATEPIECE_.Text := DateToStr(CleDoc.DatePiece);
  GP_DATELIVRAISON_.Text := DateToStr(CleDoc.DatePiece);
  GP_NUMEROPIECE_.Caption := GP_NUMEROPIECE.Caption;
  if ctxFO in V_PGI.PGIContexte then GP_DATEPIECE_.Enabled := False;
  GP_TIERS.Text := GetParamsoc('SO_GCTIERSREASSORT');
  ChargeTiers;
  TiersVersAdresses(TOBTiers, TOBAdresses, TOBPiece);
  EtudieColsListe;
  LIBELLETIERS.Visible := False;
  if Action <> taCreat then
  begin
    GP_REFINTERNE_.Text := TOBPiece.GetValue('GP_REFINTERNE');
    StringToRich(GP_BLOCNOTE, TOBPiece.GetValue('GP_BLOCNOTE'));
  end;
end;

procedure TFReassort.ReInitPiece;
begin
  inherited;
  if Action = taCreat then
  begin
    GP_BLOCNOTE.Clear;
    GP_REFINTERNE_.Text := '';
  end;
end;

procedure TFReassort.GereTiersEnabled;
begin
  BZoomTiers.Enabled := False;
end;

procedure TFReassort.FormShow(Sender: TObject);
begin
  inherited;
  GP_NUMEROPIECE_.Caption := GP_NUMEROPIECE.Caption;
  if Action <> taCreat then
  begin
    GP_REFINTERNE_.Text := TOBPiece.GetValue('GP_REFINTERNE');
    StringToRich(GP_BLOCNOTE, TOBPiece.GetValue('GP_BLOCNOTE'));
  end else GP_BLOCNOTE.Clear;
  BZoomTiers.Enabled := False;
end;

{$R *.DFM}

procedure TFReassort.GSEnter(Sender: TObject);
begin
  inherited;
  TOBPiece.PutValue('GP_DATEPIECE', StrToDate(GP_DATEPIECE_.Text));
  TOBPiece.PutValue('GP_DATELIVRAISON', StrToDate(GP_DATELIVRAISON_.Text));
  TOBPiece.PutValue('GP_REFINTERNE', GP_REFINTERNE_.Text);
end;

procedure TFReassort.ClickValide(EnregSeul: Boolean = False);
begin
  GP_REFINTERNE.Text := GP_REFINTERNE_.Text;
  TOBPiece.PutValue('GP_BLOCNOTE', RichToString(GP_BLOCNOTE));
  GP_NUMEROPIECE_.Caption := GP_NUMEROPIECE.Caption;
  inherited;
end;

procedure TFReassort.GP_DATELIVRAISONExit(Sender: TObject);
var DD: TDateTime;
  i, iLiv, iArt: integer;
  TOBL: TOB;
begin
  //  inherited;
  if csDestroying in ComponentState then Exit;
  if Action = taConsult then Exit;
  GP_DATELIVRAISON.Text := GP_DATELIVRAISON_.Text;
  if GeneCharge then Exit;
  if TOBPiece.Detail.Count <= 0 then Exit;
  DD := StrToDate(GP_DATELIVRAISON_.Text);
  if DD = TOBPiece.GetValue('GP_DATELIVRAISON') then Exit;
  TOBPiece.PutValue('GP_DATELIVRAISON', DD);
  iLiv := -1;
  iArt := -1;
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[i];
    if i = 0 then
    begin
      iLiv := TOBL.GetNumChamp('GL_DATELIVRAISON');
      iArt := TOBL.GetNumChamp('GL_ARTICLE');
    end;
    if TOBL.GetValeur(iArt) <> '' then
    begin
      TOBL.PutValeur(iLiv, DD);
      if SG_DateLiv > 0 then AfficheLaLigne(i + 1);
    end;
  end;
end;

end.
