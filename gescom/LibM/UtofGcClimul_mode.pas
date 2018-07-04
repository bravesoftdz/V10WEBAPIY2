unit UtofGcClimul_mode;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
  HCtrls, HEnt1, HMsgBox, UTOF, HDimension, HDB, UTOM,
  AglInit, UTOB, Dialogs, Menus, M3FP, EntGC, grids, LookUp,
  {$IFDEF EAGLCLIENT}
  eFichList, eFiche, emul, MaineAgl,
  {$ELSE}
  FichList, Fiche, dbTables, db, DBGrids, mul, Fe_Main,
  {$ENDIF}
  AglInitGC, utilPGI, UtilGC, ParamSoc, Ent1;

type
  TOF_GCCLIMUL_MODE = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  end;

implementation

procedure TOF_GCCLIMUL_MODE.OnArgument(Arguments: string);
var Nbr, x: integer;
  stArg, critere, ChampMul: string;
begin
  inherited;
  Nbr := 0;
  if (GCMAJChampLibre(Tform(ecran), False, 'COMBO', 'YTC_TABLELIBRETIERS', 10, '') = 0) then SetControlVisible('PTABLESLIBRES', False);
  if (GCMAJChampLibre(Tform(ecran), False, 'EDIT', 'YTC_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL', False) else inc(Nbr);
  if (GCMAJChampLibre(Tform(ecran), False, 'EDIT', 'YTC_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE', False) else inc(Nbr);
  if (GCMAJChampLibre(Tform(ecran), False, 'BOOL', 'YTC_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL', False) else inc(Nbr);
  if (Nbr = 0) then SetControlVisible('PZONESLIBRES', False);

  // initialisation de la variable Where
  THEdit(GetControl('XX_WHERE')).Text := 'T_NATUREAUXI="CLI"';

  // Initialisation car StringToAction('') retourne par défaut taConsult.
  SetControlText('TYPEACTION', 'ACTION=MODIFICATION');

  stArg := Arguments;
  repeat
    Critere := UpperCase(Trim(ReadTokenSt(stArg)));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        //ValMul:=copy(Critere,x+1,length(Critere));
      end;
      if ChampMul = 'ACTION' then SetControlText('TYPEACTION', Critere);
    end;
  until Critere = '';

  if GetControl('bInsert') <> nil then
    SetControlVisible('bInsert', StringToAction(GetControlText('TYPEACTION')) <> taConsult);

  if not ExJaiLeDroitConcept(TConcept(gcCLICreat), False) then
  begin
    SetControlVisible('BINSERT', False);
  end;
end;

procedure TOF_GCCLIMUL_MODE.OnLoad;
var xx_where: string;
begin
  inherited;
  if not (Ecran is TFMul) then exit;
  xx_where := '';

  //Ajoute l'affichage des prospects dans le mul
  {$IFDEF NOMADE}
  xx_where := 'OR T_NATUREAUXI="PRO"';
  {$ENDIF}

  // Gestion des checkBox : booléens libres
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'BOOL', 'YTC_BOOLLIBRE', 3, '');

  // Gestion des dates libres
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'DATE', 'YTC_DATELIBRE', 3, '_');

  // Gestion des valeurs libres
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'EDIT', 'YTC_VALLIBRE', 3, '_');

  SetControlText('XX_WHERE', xx_where);

  //CB:=TCheckBox(TFMul(F).FindComponent('STATUT_DIM')) ;
  //bAffiche:=boolean((CB<>nil) and (CB.State=cbChecked)) ;

end;

function TOF_GCCLIMUL_MODE_AGLNatureTiers(Parms: array of variant; nb: integer): variant;
begin
  {$IFDEF NOMADE}
  if ctxMode in V_PGI.PGIContexte then Result := 'CLI'
  else
    if GetParamSoc('SO_GCPREFIXETIERS') <> '' then Result := 'PRO' else Result := '0';
  {$ELSE}
  Result := 'CLI';
  {$ENDIF}
end;

initialization
  registerclasses([TOF_GCCLIMUL_MODE]);
  RegisterAglFunc('TiersNatureMode', True, 1, TOF_GCCLIMUL_MODE_AGLNatureTiers);
end.
