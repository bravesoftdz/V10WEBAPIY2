unit FichComm;

interface

uses Classes, UTOM, Forms, Hctrls, HEnt1, UtilPGI, AGLInit,
{$IFDEF EAGLCLIENT}
  MaineAGL,
{$ELSE}
{$IFNDEF DBXPRESS}dbtables {BDE}, {$ELSE}uDbxDataSet, {$ENDIF}Fe_Main,
{$ENDIF}
  HmsgBox, SysUtils;


procedure FicheModePaie_AGL(Quel: string);
function FicheModePaieZoom_AGL(Quel: string): string;
procedure FicheEtablissement_AGL(Mode: TActionFiche);
procedure FicheContact_AGL(Auxiliaire: string; Mode: TActionFiche);
function FicheRIB_AGL(Auxiliaire: string; Mode: TActionFiche; FromSaisie: boolean; StRIB: string; IsAux: Boolean): integer;
function FicheRegle_AGL(Quel: string; OkGuide: Boolean; Comment: TActionFiche): string;
procedure FicheBanque_AGL(CodeBanque: string; Comment: TActionFiche; QuellePage: Integer);

implementation

procedure FicheModePaie(Quel: string);
begin
  if _Blocage(['nrCloture', 'nrBatch'], True, 'nrAucun') then Exit;
  AGLLanceFiche('YY', 'YYMODEPAIE', '', Quel, '');
end;

function FicheModePaieZoom(Quel: string): string;
begin
  if _Blocage(['nrCloture', 'nrBatch'], True, 'nrAucun') then Exit;
  Result := AGLLanceFiche('YY', 'YYMODEPAIE', '', Quel, 'ZOOM');
end;

procedure FicheEtablissement(Mode: TActionFiche);
var Arguments: string;
begin
  if _Blocage(['nrCloture', 'nrBatch'], True, 'nrAucun') then Exit;
  Arguments := ActionToString(Mode);
  AGLLanceFiche('YY', 'YYETABLISS', '', '', Arguments);
end;

procedure FicheContact(Auxiliaire: string; Mode: TActionFiche);
begin
  if _Blocage(['nrCloture'], True, 'nrAucun') then Exit;
  AGLLanceFiche('YY', 'YYCONTACT', 'T;' + Auxiliaire, '', ActionToString(Mode));
end;

function FicheRIB(Auxiliaire: string; Mode: TActionFiche; FromSaisie: boolean; StRIB: string; IsAux: Boolean): integer;
var Arguments: string;
  vv: variant;
begin
  Result := -1;
  if Auxiliaire <> '' then
  begin
    Arguments := 'NumAux=' + Auxiliaire + ';' + ActionToString(Mode) + ';FromSaisie=' + booltostr_(FromSaisie) + ';StRib=' + StRIB + ';IsAux=' + BoolToStr_(IsAux);
    vv := AGLLanceFiche('YY', 'YYRIB', Auxiliaire, '', Arguments);
    if vv = '' then Result := -1 else Result := StrToInt(VV);
  end;
end;

function FicheRegle(Quel: string; OkGuide: Boolean; Comment: TActionFiche): string;
var stGuide: string;
begin
  if _Blocage(['nrCloture', 'nrBatch'], True, 'nrAucun') then Exit;
  if OkGuide then stGuide := 'GUIDE' else stGuide := '';
  AGLLanceFiche('YY', 'YYMODEREGL', '', Quel, ActionToString(Comment) + ';' + stGuide);
end;

procedure FicheBanque(CodeBanque: string; Comment: TActionFiche; QuellePage: Integer);
begin
  if _Blocage(['nrCloture'], True, 'nrAucun') then Exit;
  AGLLanceFiche('YY', 'YYBANQUES', '', CodeBanque, ActionToString(Comment));
end;

procedure FicheModePaie_AGL(Quel: string);
begin FicheModePaie(Quel); end;

function FicheModePaieZoom_AGL(Quel: string): string;
begin Result := FicheModePaieZoom(Quel); end;

procedure FicheEtablissement_AGL(Mode: TActionFiche);
begin FicheEtablissement(Mode); end;

procedure FicheContact_AGL(Auxiliaire: string; Mode: TActionFiche);
begin FicheContact(Auxiliaire, Mode); end;

function FicheRIB_AGL(Auxiliaire: string; Mode: TActionFiche; FromSaisie: boolean; StRIB: string; IsAux: Boolean): integer;
begin Result := FicheRIB(Auxiliaire, Mode, FromSaisie, StRIB, IsAux); end;

function FicheRegle_AGL(Quel: string; OkGuide: Boolean; Comment: TActionFiche): string;
begin Result := FicheRegle(Quel, OkGuide, Comment); end;

procedure FicheBanque_AGL(CodeBanque: string; Comment: TActionFiche; QuellePage: Integer);
begin FicheBanque(CodeBanque, Comment, QuellePage); end;

end.
