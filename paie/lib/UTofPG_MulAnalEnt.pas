{***********UNITE*************************************************
Auteur  ...... : Philippe Dumet
Créé le ...... : 27/02/2004
Modifié le ... : 27/02/2004
Description .. : Multi critère analyse des entretiens individuels
Suite ........ :
Mots clefs ... : PAIE
*****************************************************************}
{
}

unit UTofPG_MulAnalEnt;

interface
uses StdCtrls,
  Controls,
  Classes,
  Graphics,
  forms,
  sysutils,
  ComCtrls,
  HTB97,
  HStatus,
  {$IFNDEF EAGLCLIENT}
  HDB, DBGrids, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, mul,
  {$ELSE}
  MaineAGL, eMul,
  {$ENDIF}
  Grids, HCtrls, HEnt1, vierge, EntPaie, HMsgBox, Hqry, UTOF, UTOB, UTOM,
  AGLInit;
type
  TOF_PGMulAnalEnt = class(TOF)
  private
    Q_Mul: THQuery;
    procedure ActiveWhere();
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  end;

implementation


procedure TOF_PGMulAnalEnt.ActiveWhere();
var
  St: string;
  DFin, DDebut: TDateTime;
begin
  st := '';
  DDebut := StrToDate(GetControlText('DATEENTRETIEN'));
  DFin := StrToDate(GetControlText('DATEENTRETIEN_'));

  if GetControlText('SANSENTRETIEN') <> 'X' then
  begin
    if (Q_Mul <> nil) and (Q_Mul.Liste <> 'PGSUIVIPERSONNEL') then TFMul(Ecran).SetDBListe('PGSUIVIPERSONNEL');
    st := 'PSV_PGTYPSUIVI="EAN" AND PSV_TYPEINTERIM="SAL"';
  end
  else
  begin
    if (Q_Mul <> nil) and (Q_Mul.Liste <> 'PGMULSALARIE') then TFMul(Ecran).SetDBListe('PGMULSALARIE');
    St := ' (PSA_SUSPENSIONPAIE <> "X") AND ' +
      '(NOT EXISTS (Select PSV_INTERIMAIRE FROM SUIVIPERSONNEL WHERE PSV_INTERIMAIRE = PSA_SALARIE and PSV_PGTYPSUIVI = "EAN" AND PSV_TYPEINTERIM = "SAL"' +
      ' AND PSV_DATEENTRETIEN >="' + UsDateTime(DDebut) + '" AND PSV_DATEENTRETIEN <="' + UsDateTime(DFin) + '"))';
  end;
  SetControlText('XX_WHERE', St);
end;

procedure TOF_PGMulAnalEnt.OnArgument(Arguments: string);
var
  Num: Integer;
begin
  inherited;
   Q_Mul := THQuery(Ecran.FindComponent('Q'));
 {   for Num := 1 to 4 do
      VisibiliteChampSalarie(IntToStr(Num), GetControl('TRAVAILN' + IntToStr(Num)), GetControl('TPSA_TRAVAILN' + IntToStr(Num)));

    VisibiliteStat(GetControl('CODESTAT'), GetControl('TPSA_CODESTAT'));}

end;

procedure TOF_PGMulAnalEnt.OnLoad;
var
  TT: TFMul;
begin

  if not (Ecran is TFMul) then exit;
  if GetControlText('SANSENTRETIEN') = 'X' then TFMul(Ecran).Caption := 'Liste des salariés n''ayant pas eu d''entretien individuel'
  else TFMul(Ecran).Caption := 'Liste des entretiens individuels';
  TT := TFMul(Ecran);
  if TT <> nil then UpdateCaption(TT);

  ActiveWhere();
  inherited;
end;

initialization
  registerclasses([TOF_PGMulAnalEnt]);
end.

