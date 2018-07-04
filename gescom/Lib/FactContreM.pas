{***********UNITE*************************************************
Auteur  ...... : Dominique BROSSET
Créé le ...... : 14/08/2003
Modifié le ... :   /  /    
Description .. : Fonction utilitaires pour la gestion de la contremarque dans 
Suite ........ : les pièces
Mots clefs ... : PIECES;CONTREMARQUE;UTIL
*****************************************************************}
unit FactContreM ;

interface

uses HEnt1, UTOB, HCtrls, LookUp, SysUtils,
{$IFDEF EAGLCLIENT}
     Utileagl, MaineAgl,
{$ELSE}
     Fe_Main, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     AGLInitGC, EntGC, UtilPGI,
     FactTob, FactArticle;

procedure TraiteLaContreM(ARow: integer; TobPiece, TobCatalogu : Tob);
procedure ChargeTOBDispoContreM(ARow: integer; TobPiece, TobCatalogu : TOB);
procedure ReserveCliContreM(TOBL, TOBCata: TOB; Plus: Boolean);
function CreerTOBDispoContreM(TOBL, TOBCata: TOB): TOB;
function TrouverTobDispoContreM(TOBL, TOBCata: TOB; AvecClient: boolean): TOB;
function GetCodeClientDCM(TOBL: TOB): string;
function GetCodeFourDCM(TOBL: TOB): string;
function TrouverContreMSQL(NaturePiece, RefSais: string; DatePiece: TDateTime; TOBCata: TOB): T_RechArt;

implementation

uses DepotUtil, FactUtil;

procedure TraiteLaContreM(ARow: integer; TobPiece, TobCatalogu : Tob);
var RefCata, RefFour: string;
  TOBL, TOBCata: TOB;
begin
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if TOBL.GetValue('GL_TYPEREF') = 'CAT' then
  begin
    GetCodeCataUnique(TOBPiece, ARow, RefCata, RefFour);
    if (RefCata = '') or (RefFour = '') then Exit;
    TOBCata := FindTOBCataRow(TOBPiece, TOBCataLogu, ARow);
    if TOBCata = nil then Exit;
    if TOBCata <> nil then
      ReserveCliContreM(TOBL, TOBCata, True);
  end;
end;

procedure ChargeTOBDispoContreM(ARow: integer; TobPiece, TobCatalogu : TOB);
var TOBL, TOBC: TOB;
begin
  TOBC := FindTOBCataRow(TOBPiece, TOBCatalogu, ARow);
  if TOBC = nil then Exit;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  LoadTOBDispoContreM(TOBL, TOBC, False, CreerQuelDepot(TOBPiece));
end;

procedure ReserveCliContreM(TOBL, TOBCata: TOB; Plus: Boolean);
var TOBDCM: TOB;
  Qte, Dispo: double;
begin
  if (TOBCata = nil) or (TOBL = nil) then Exit;
  Qte := TOBL.GetValue('GL_QTEFACT');
  if Qte = 0 then Exit;

  if TrouverTobDispoContreM(TOBL, TOBCata, True) = nil then CreerTobDispoContreM(TOBL, TOBCata);

  if Plus then
  begin
    TOBDCM := TrouverTobDispoContreM(TOBL, TOBCata, False);
    if TOBDCM <> nil then
    begin
      Dispo := TOBDCM.GetValue('GQC_PHYSIQUE') - TOBDCM.GetValue('RESERVECLI');
      if Qte <= Dispo then
      begin
        TOBDCM.PutValue('RESERVECLI', TOBDCM.GetValue('RESERVECLI') + Qte);
        Qte := 0;
      end else
      begin
        TOBDCM.PutValue('RESERVECLI', TOBDCM.GetValue('RESERVECLI') + Dispo);
        Qte := Qte - Dispo;
      end;
    end;
    if Qte <> 0 then
    begin
      TOBDCM := TrouverTobDispoContreM(TOBL, TOBCata, true);
      TOBDCM.PutValue('RESERVECLI', TOBDCM.GetValue('RESERVECLI') + Qte);
    end;
  end else
  begin
    TOBDCM := TrouverTobDispoContreM(TOBL, TOBCata, true);
    Dispo := TOBDCM.GetValue('RESERVECLI');
    if (Dispo <> 0) then
    begin
      if (Qte <= Dispo) then
      begin
        TOBDCM.PutValue('RESERVECLI', TOBDCM.GetValue('RESERVECLI') - Qte);
        Qte := 0;
      end else
      begin
        TOBDCM.PutValue('RESERVECLI', TOBDCM.GetValue('RESERVECLI') - Dispo);
        Qte := Qte - Dispo;
      end;
    end;
    if Qte <> 0 then
    begin
      TOBDCM := TrouverTobDispoContreM(TOBL, TOBCata, False);
      if TOBDCM <> nil then
        TOBDCM.PutValue('RESERVECLI', TOBDCM.GetValue('RESERVECLI') - Qte)
      else
      begin
        TOBDCM := TrouverTobDispoContreM(TOBL, TOBCata, true);
        TOBDCM.PutValue('RESERVECLI', TOBDCM.GetValue('RESERVECLI') - Qte);
      end;
    end;
  end;
end;

function CreerTOBDispoContreM(TOBL, TOBCata: TOB): TOB;
var TOBDCM: TOB;
begin
  result := nil;
  if (TOBCata = nil) or (TOBL = nil) then Exit;
  TOBDCM := TOB.Create('DISPOCONTREM', TOBCata, -1);
  TOBDCM.PutValue('GQC_REFERENCE', TOBCata.GetValue('GCA_REFERENCE'));
  TOBDCM.PutValue('GQC_FOURNISSEUR', TOBCata.GetValue('GCA_TIERS'));
  TOBDCM.PutValue('GQC_CLIENT', GetCodeClientDCM(TOBL));
  TOBDCM.PutValue('GQC_DEPOT', TOBL.GetValue('GL_DEPOT'));
  TOBDCM.PutValue('GQC_ARTICLE', TOBL.GetValue('GL_ARTICLE'));
  TOBDCM.AddChampSupValeur('RESERVECLI', 0, False);
  Result := TOBDCM;
end;

function TrouverTobDispoContreM(TOBL, TOBCata: TOB; AvecClient: boolean): TOB;
var CodeClient: string;
begin
  Result := nil;
  if (TOBCata = nil) or (TOBL = nil) then Exit;
  if (TOBCata.Detail.Count = 0) then Exit;
  CodeClient := '';
  if AvecClient then CodeClient := GetCodeClientDCM(TOBL);

  Result := TOBCata.FindFirst(['GQC_REFERENCE', 'GQC_FOURNISSEUR', 'GQC_CLIENT', 'GQC_DEPOT'],
    [TOBCata.GetValue('GCA_REFERENCE'), TOBCata.GetValue('GCA_TIERS'), CodeClient, TOBL.GetValue('GL_DEPOT')], False);
end;

function GetCodeClientDCM(TOBL: TOB): string;
var VenteAchat: string;
  MouvExContrem: boolean;
begin
  VenteAchat := GetInfoParPiece(TOBL.GetValue('GL_NATUREPIECEG'), 'GPP_VENTEACHAT');
  MouvExContrem := ((TOBL.GetValue('GL_NATUREPIECEG') = 'EEX') or (TOBL.GetValue('GL_NATUREPIECEG') = 'SEX'))
    and (TOBL.GetValue('GL_ENCONTREMARQUE') = 'X');
  if (VenteAchat = 'ACH') or (MouvExContrem) then
    result := TOBL.GetValue('GL_FOURNISSEUR')
  else
    result := TOBL.GetValue('GL_TIERS')
end;

function GetCodeFourDCM(TOBL: TOB): string;
var VenteAchat: string;
  MouvExContrem: boolean;
begin
  VenteAchat := GetInfoParPiece(TOBL.GetValue('GL_NATUREPIECEG'), 'GPP_VENTEACHAT');
  MouvExContrem := ((TOBL.GetValue('GL_NATUREPIECEG') = 'EEX') or (TOBL.GetValue('GL_NATUREPIECEG') = 'SEX'))
    and (TOBL.GetValue('GL_ENCONTREMARQUE') = 'X');
  if (VenteAchat = 'ACH') or (MouvExContrem) then
    result := TOBL.GetValue('GL_TIERS')
  else
    result := TOBL.GetValue('GL_FOURNISSEUR');
end;

function TrouverContreMSQL(NaturePiece, RefSais: string; DatePiece: TDateTime; TOBCata: TOB): T_RechArt;
var QQ: TQuery;
  TOBT: TOB;
begin
  Result := traAucun;
  if ((RefSais = '') or (NaturePiece = '') or (TOBCata = nil)) then Exit;
  QQ := OpenSQL('Select * from CATALOGU Where GCA_REFERENCE="' + RefSais + '" AND GCA_DATESUP>="' + USDateTime(DatePiece) + '"', True,-1, '', True);
  if not QQ.EOF then
  begin
    TOBT := TOB.Create('', nil, -1);
    TOBT.LoadDetailDB('CATALOGU', '', '', QQ, False);
    if TOBT.Detail.Count = 1 then
    begin
      TOBCata.Dupliquer(TOBT.Detail[0], True, True);
      Result := traOk;
    end;
    TOBT.Free;
  end else
  begin
    Ferme(QQ);
    QQ := OpenSQL('Select * from CATALOGU Where GCA_ARTICLE like "' + RefSais + ' %" AND GCA_DATESUP>="' + USDateTime(DatePiece) + '"', True,-1, '', True);
    if not QQ.EOF then
    begin
      TOBT := TOB.Create('', nil, -1);
      TOBT.LoadDetailDB('CATALOGU', '', '', QQ, False);
      if TOBT.Detail.Count = 1 then
      begin
        TOBCata.Dupliquer(TOBT.Detail[0], True, True);
        Result := traOk;
      end;
      TOBT.Free;
    end;
  end;
  Ferme(QQ);
  if Result = traOk then
  begin
    if not TOBCata.FieldExists('UTILISE') then TOBCata.AddChampSup('UTILISE', False);
    TOBCata.PutValue('UTILISE', '-');
  end;
end;

end.
