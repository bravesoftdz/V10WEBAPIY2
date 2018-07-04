unit Udefexport;

interface
uses SysUtils,HCtrls;
type
	TmodeAction = (XmaAffaire,XmaDocument, XmaDemandePrix);
  TTypeExport = (TteUndefined,TteDevis,TteCotrait,TteSousTrait, TteDdePrix);
  TModeExport = (TmeUndefined,TmeExcel,TmeXML);

  R_CLEXLS = RECORD
    NaturePiece : string;
    Souche      : String ;
    NumeroPiece : Integer;
    NumLigne    : Integer;
    NumOrdre    : Integer ;
    Indice      : Integer;
    UniqueBlo   : integer;
  END ;


procedure DecodeClefXls (texte : string; var ClefXls : R_CLEXLS) ;
implementation

procedure DecodeClefXls (texte : string; var ClefXls : R_CLEXLS) ;
var TheZones,TheZone : string;
begin
  FillChar(ClefXls,SizeOf(ClefXls),#0);
  TheZones := texte;
  TheZone := READTOKENST(TheZones);
  if TheZone = '' then exit;
  ClefXls.NaturePiece := TheZone;
  TheZone := READTOKENST(TheZones);
  if TheZone = '' then exit;
  ClefXls.Souche  := TheZone;
  TheZone := READTOKENST(TheZones);
  if TheZone = '' then exit;
  ClefXls.NumeroPiece   := StrToInt(TheZone);
  TheZone := READTOKENST(TheZones);
  if TheZone = '' then exit;
  ClefXls.Indice    := StrToInt(TheZone);
  TheZone := READTOKENST(TheZones);
  if TheZone = '' then exit;
  ClefXls.NumOrdre    := StrToInt(TheZone);
  TheZone := READTOKENST(TheZones);
  if TheZone = '' then exit;
  ClefXls.UniqueBlo := StrToInt(TheZone);
end;

end.
 