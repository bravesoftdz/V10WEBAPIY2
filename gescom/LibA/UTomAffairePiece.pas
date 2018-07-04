{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/07/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : AFFAIREPIECE (AFFAIREPIECE)
Mots clefs ... : TOM;AFFAIREPIECE
*****************************************************************}
unit UTomAffairePiece;

interface

uses StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOM,
  UTob,
  AffaireUtil,
  EntGC,
  ParamSoc;

type
  TOM_AFFAIREPIECE = class(TOM)
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;
  end;

implementation

procedure TOM_AFFAIREPIECE.OnNewRecord;
begin
  inherited;
end;

procedure TOM_AFFAIREPIECE.OnDeleteRecord;
begin
  inherited;
end;

procedure TOM_AFFAIREPIECE.OnUpdateRecord;
begin
  inherited;
end;

procedure TOM_AFFAIREPIECE.OnAfterUpdateRecord;
begin
  inherited;
end;

procedure TOM_AFFAIREPIECE.OnLoadRecord;
begin
  inherited;
end;

procedure TOM_AFFAIREPIECE.OnChangeField(F: TField);
begin
  inherited;
  if (F.FieldName = 'API_NATUREPIECEG') and (F.Value <> '') then
    if GetField('API_LIBELLE') = '' then
    begin
      SetField('API_LIBELLE', RechDom('GCNATUREPIECEG', GetField('API_NATUREPIECEG'), FALSE));
    end;
end;

procedure TOM_AFFAIREPIECE.OnArgument(S: string);
var i, PosVenteAchat, PosNaturePieceg: integer;
  PlusCombo: string;
begin
  inherited;
  PlusCombo := PlusCombo + ' AND (GPP_NATUREPIECEG IN ("FAC","AVC","FPR","APR","FSI","FRE"';
  if (ctxGCAff in v_PGI.PGIContexte) or (ctxtempo in v_PGI.PGIContexte) then
  begin
    PosVenteAchat := VH_GC.TobParPiece.Detail[0].GetNumChamp('GPP_VENTEACHAT');
    PosNaturePieceg := VH_GC.TobParPiece.Detail[0].GetNumChamp('GPP_NATUREPIECEG');
    for i := 0 to VH_GC.TobParPiece.detail.Count - 1 do
    begin
      if VH_GC.TobParPiece.detail[i].Getvaleur(PosVenteAchat) = 'ACH' then
        PlusCombo := PlusCombo + ',"' + VH_GC.TobParPiece.detail[i].Getvaleur(PosNaturePieceg) + '"';
    end;
  end
  else if (ctxscot in v_PGI.PGIContexte) then
    PlusCombo := PlusCombo + ',"AF","FF"';
  PlusCombo := PlusCombo + '))';
  SetControlProperty('API_NATUREPIECEG', 'Plus', PlusCombo);
  SetControlVisible('API_IMPFORMULE', GetParamSoc('SO_AFVARIABLES'));
  SetControlVisible('API_IMPREVISION', GetParamSoc('SO_AFREVISIONPRIX'));
end;

procedure TOM_AFFAIREPIECE.OnClose;
begin
  inherited;
end;

procedure TOM_AFFAIREPIECE.OnCancelRecord;
begin
  inherited;
end;

initialization
  registerclasses([TOM_AFFAIREPIECE]);
end.
