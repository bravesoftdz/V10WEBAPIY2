{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 29/12/2000
Modifié le ... : 23/07/2001
Description .. : Source TOF de la FICHE : FOCLIMUL_MODE
Suite ........ : Recherce des clients "Mode" pour le Front Office
Mots clefs ... : TOF;UTOFFOCLIMUL_MODE;FO
*****************************************************************}
unit UTofFOCliMul_Mode;

interface
uses
  Classes, Controls, SysUtils, forms,
  {$IFDEF EAGLCLIENT}
  eMUL,
  {$ELSE}
  MUL, HIdxTabs,
  {$ENDIF}
  UTOF, HMsgBox, Hent1, HCtrls, AglInit;

type
  TOF_FOCLIMUL_MODE = class(TOF)
  private
    {$IFDEF EAGLCLIENT}
    {$ELSE}
    IndexTabs: THIdxTabs;
    procedure IndexTabsButtonClick(Sender: TObject; Index: Integer; ButtonCaption: string);
    {$ENDIF}
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  end;

implementation
uses
  UtilGC, Ent1, FOUtil;

///////////////////////////////////////////////////////////////////////////////////////
//  IdxTabsToWhere : Constitue une clause Where à partir de la caption du bouton choisi
///////////////////////////////////////////////////////////////////////////////////////

function IdxTabsToWhere(Index: Integer; ButtonCaption: string): string;
var Stg: string;
  Ind: Integer;
begin
  Stg := Uppercase(ButtonCaption);
  Result := '';
  if Index > 0 then for Ind := Ord(Stg[1]) to Ord(Stg[Length(Stg)]) do
    begin
      if Result = '' then Result := '(' else Result := Result + ' OR ';
      Result := Result + 'T_LIBELLE LIKE "' + Chr(Ind) + '%"';
    end;
  if Result <> '' then Result := Result + ')';
end;

///////////////////////////////////////////////////////////////////////////////////////
//  IndexTabsButtonClick
///////////////////////////////////////////////////////////////////////////////////////
{$IFNDEF EAGLCLIENT}

procedure TOF_FOCLIMUL_MODE.IndexTabsButtonClick(Sender: TObject; Index: Integer; ButtonCaption: string);
begin
  SetControlText('XX_WHERE_', IdxTabsToWhere(Index, ButtonCaption));
  SetControlText('T_LIBELLE', '');
  TFMul(Ecran).BChercheClick(nil);
  if IndexTabs <> nil then IndexTabs.Refresh;
end;
{$ENDIF}

///////////////////////////////////////////////////////////////////////////////////////
//  OnArgument
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_FOCLIMUL_MODE.OnArgument(Arguments: string);
var Nbr, x: integer;
  stArg, critere, ChampMul: string;
  {$IFNDEF EAGLCLIENT}
  Composant: TComponent;
  {$ENDIF}
begin
  inherited;
  Nbr := 0;
  if (GCMAJChampLibre(Tform(ecran), False, 'COMBO', 'YTC_TABLELIBRETIERS', 10, '') = 0) then SetControlVisible('PTABLESLIBRES', False);
  if (GCMAJChampLibre(Tform(ecran), False, 'EDIT', 'YTC_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL', False) else Nbr := Nbr + 1;
  if (GCMAJChampLibre(Tform(ecran), False, 'EDIT', 'YTC_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE', False) else Nbr := Nbr + 1;
  if (GCMAJChampLibre(Tform(ecran), False, 'BOOL', 'YTC_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL', False) else Nbr := Nbr + 1;
  if (Nbr = 0) then SetControlVisible('PZONESLIBRES', False);

  // initialisation de la variable Where
  SetControlText('XX_WHERE', 'T_NATUREAUXI="CLI"');

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

  SetControlVisible('bInsert', StringToAction(GetControlText('TYPEACTION')) <> taConsult);

  if not ExJaiLeDroitConcept(TConcept(gcCLICreat), False) then
  begin
    SetControlVisible('BINSERT', False);
  end;

  if not FOJaiLeDroit(13, False, False) then
    SetControlVisible('BMODIF', False);

  if not FOJaiLeDroit(35, False, False) then
    SetControlText('TYPEACTION', 'ACTION=CONSULTATION');

  // Ajout de l'index dans le panel PINDEXTABS prévu à cet effet
  {$IFDEF EAGLCLIENT}
  SetControlVisible('PINDEXTABS', False);
  {$ELSE}
  if Ecran <> nil then
  begin
    Composant := Ecran.FindComponent('PINDEXTABS');
    if Composant <> nil then
    begin
      IndexTabs := THIdxTabs.Create(Composant);
      IndexTabs.Name := 'IIndexTabs';
      IndexTabs.Parent := TWinControl(Composant);
      IndexTabs.Align := alClient;
      IndexTabs.Visible := True;
      IndexTabs.UpperCase := True;
      IndexTabs.OnButtonClick := IndexTabsButtonClick;
    end;
  end;
  {$ENDIF}
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnLoad
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_FOCLIMUL_MODE.OnLoad;
var xx_where: string;
begin
  inherited;
  if not (Ecran is TFMul) then exit;
  xx_where := '';

  // Gestion des checkBox : booléens libres
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'BOOL', 'YTC_BOOLLIBRE', 3, '');

  // Gestion des dates libres
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'DATE', 'YTC_DATELIBRE', 3, '_');

  // Gestion des valeurs libres
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'EDIT', 'YTC_VALLIBRE', 3, '_');

  SetControlText('XX_WHERE', xx_where);

  //CB:=TCheckBox(TFMul(F).FindComponent('STATUT_DIM')) ;
  //bAffiche:=boolean((CB<>nil) and (CB.State=cbChecked)) ;

  // Positionnement des boutons dans la barre de boutons
  SetControlProperty('BAgrandir', 'Left', 0);
  SetControlProperty('BRechercher', 'Left', 60);
  SetControlProperty('BParamListe', 'Left', 120);
  SetControlProperty('BExport', 'Left', 180);
  SetControlProperty('BSelectAll', 'Left', 240);
  SetControlProperty('BInsert', 'Left', 300);
  SetControlProperty('BModif', 'Left', 360);
  SetControlProperty('BBlocNote', 'Left', 420);
  SetControlProperty('BImprimer', 'Left', 480);
  SetControlProperty('BOuvrir', 'Left', 540);
  SetControlProperty('BAnnuler', 'Left', 600);
  SetControlProperty('BAide', 'Left', 660);

  // A voir
  //SetControlProperty('PINDEXTABS', 'Visible', False) ;
  //SetControlProperty('PINDEXTABS', 'Enabled', False) ;

end;

initialization
  RegisterClasses([TOF_FOCLIMUL_MODE]);
end.
