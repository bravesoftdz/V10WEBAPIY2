{***********UNITE*************************************************
Auteur  ...... : JP
Créé le ...... : 25/08/2004
Modifié le ... : 18/08/2005
Description .. : Reprise de l'unité et remplacement des THEdit par des
Suite ........ : THValComboBox FQ 13401, 13452, 13671
Suite ........ : - FQ 16267 - CA - 18/08/2005 : Pour ouvrir les choix sur la 
Suite ........ : table RESSOURCE, gestion de THCritMaskEdit pour les 3 
Suite ........ : premières tables libres dans le cas de l'assistant client.
Mots clefs ... : 
*****************************************************************}
unit TofTLTiersRupt;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  uTob,
  {$ELSE}
  FE_Main,
  {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
  Classes, StdCtrls, ComCtrls, UTof, HCtrls, SysUtils, Vierge;

type
  TOF_TLTIERSRUPT = class(TOF)
  private
    ArgRetour : string;
    fTypeTableLibre : string;
    procedure T00OnClick(Sender : TObject);
    procedure TDeOnChange(Sender : TObject);
    procedure TaOnChange(Sender : TObject);
    procedure InitTablesLibres;
  public
    procedure OnArgument(Arguments : string); override ;
    procedure OnCancel; override ;
    procedure OnUpdate; override ;
  end ;

function CPLanceFiche_TLTIERSRUPT(Arg : string) : string;
Procedure GetLibelleTableLibreTiersCompl (stPref : String ; LesLib : HTStringList) ;

implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier, 
  {$ENDIF MODENT1}
  Ent1;


{---------------------------------------------------------------------------------------}
function CPLanceFiche_TLTIERSRUPT(Arg : string) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := AglLanceFiche('CP', 'TLTIERSRUPT', '', '', Arg);
end;

{En cas d'annulation, on renvoie le paramétrage passé en argument
{---------------------------------------------------------------------------------------}
procedure TOF_TLTIERSRUPT.OnCancel;
{---------------------------------------------------------------------------------------}
begin
  TFVierge(Ecran).Retour := ArgRetour;
end;

{Constitution de la chaine de retour de la fiche
{---------------------------------------------------------------------------------------}
procedure TOF_TLTIERSRUPT.OnUpdate;
{---------------------------------------------------------------------------------------}
var
  St,
  StDe,
  StA : string;
  i   : Integer;
  n   : string;
  CB  : TCheckBox;
begin
  for i := 0 to 9 do begin
    n := IntToStr(i);
    CB := TCheckBox(GetControl('T0' + n)) ;

    {La table libre est active ...}
    if CB.Enabled then begin
      {... et la table libre est paramétrée ...}
      if CB.Checked then begin
        {..., mais on ne fait pas de sélection (<<Tous>>)}
        if ((fTypeTableLibre <> 'SCL') and (GetControlText('DETABLET0' + n) = ''))
              or  (((fTypeTableLibre = 'SCL') and (i<=2)) and (GetControlText('DETABLER0' + n) = '')) then
        begin
          StDe := StDe + '*;';
          StA  := StA  + '*;';
        end
        {..., on filtre les valeurs}
        else begin
          if (fTypeTableLibre = 'SCL') then
          begin
            StDe := StDe + GetControlText('DETABLER0'+ n) + ';';
            StA  := StA  + GetControlText('ATABLER0' + n) + ';';
          end else
          begin
            StDe := StDe + GetControlText('DETABLET0'+ n) + ';';
            StA  := StA  + GetControlText('ATABLET0' + n) + ';';
          end;
        end
      end
      {La table libre est active, mais non paramétrée}
      else begin
        StDe := StDe + '-;';
        StA  := StA  + '-;';
      end;{if CB.Checked}
    end {if CB.Enabled}

    {La table libre n'est pas active}
    else begin
      StDe := StDe + '#;';
      StA  := StA  + '#;';
    end;
  end; {for i := 0}

  {Concaténation des fourchettes de début et de fin}
  St := StDe + '|' + StA;
  {La fiche retourne la chaine ainsi constituée}
  TFVierge(Ecran).Retour := St;
END;

{---------------------------------------------------------------------------------------}
procedure TOF_TLTIERSRUPT.OnArgument(Arguments : string);
{---------------------------------------------------------------------------------------}
var
  i   : Integer;
  n   : string;
  CB  : TCheckBox;
  cDe,
  cA  : THValComboBox;
  LA  : TLabel;
  LesLib : HTStringList;
  St , St1,
  St2, St3,
  St4, St5 : string;
  eDe, eA : THcritMaskEdit;
begin
  inherited ;
  {Par défaut, on renvoie la séléction d'origine amputée du type de table libre}
  ArgRetour := Copy (Arguments, 1, Length(Arguments)-4);
  TFVierge(Ecran).Retour := ArgRetour;

  St2 := Arguments;
  {'|' est le séparateur entre les valeurs de début et celles de fin :
   l'argument est de la forme T0deb;T1deb;...;T9deb;|T0fin;T1fin;...;T9fin; cf OnUpdate}
  St3 := ReadTokenPipe(St2, '|');

  { Lecture du type de table libre sur lequel on travaille : Auxiliaire, clients ou assistants }
  fTypeTableLibre := St2;
  St2 := ReadTokenPipe(fTypeTableLibre, '|');
  InitTablesLibres;

  {Récupération du paramétrage des tables libres}
  LesLib:=HTStringList.Create ;
  if fTypeTableLibre = 'AUX' then
    GetLibelleTableLibre('T', LesLib)
  else if fTypeTableLibre = 'CLI' then GetLibelleTableLibreTiersCompl ('CT', LesLib)
  else if fTypeTableLibre = 'SCL' then GetLibelleTableLibreTiersCompl ('CR', LesLib);

  {Activation des composants de l'écran en fonction du paramétrage des tables libles
   et des paramètres passés en argument}
  for i := 0 to 9 do
  begin
    n   := IntToStr(i);
    CB  := TCheckBox(GetControl('T0' + n));
    cDe := THValComboBox(GetControl('DETABLET0' + n));
    cA  := THValComboBox(GetControl('ATABLET0'  + n));
    if (i<=2) then
    begin
      eDe := THCritMaskEdit(GetControl('DETABLER0' + n));
      eA :=  THCritMaskEdit(GetControl('ATABLER0' + n));
    end else
    begin
      eDE := nil;
      eA := nil;
    end;
    LA  := TLabel(GetControl('TaT0' + n));
    if Assigned(CB) and Assigned(cDe) and Assigned(cA) and Assigned(LA) then
    begin
      CB.Visible := not ((fTypeTableLibre = 'SCL') and (i>2));
      cDe.Visible := not (fTypeTableLibre = 'SCL');
      cA.Visible := not (fTypeTableLibre = 'SCL');
      LA.Visible := not ((fTypeTableLibre = 'SCL') and (i>2));
      if (i<=2) then 
      begin
        eDe.Visible := (fTypeTableLibre = 'SCL');
        eA.Visible := (fTypeTableLibre = 'SCL');
      end;
      SetControlVisible('TLIBT0' + n, not ((fTypeTableLibre = 'SCL') and (i>2)));
      SetControlVisible('TDET0' + n, not ((fTypeTableLibre = 'SCL') and (i>2)));
      if ((fTypeTableLibre = 'SCL') and (i>2)) then continue;

      {Récupération du paramétrage de la table libre en cours}
      St  := LesLib[i];
      St1 := ReadTokenSt(St) ;
      CB.Enabled := St = 'X';
      SetControlCaption('TLIBT0' + n, St1);
      CB.OnClick := T00OnClick;
      {Récupération de l'éventuelle fourchette passée en argument pour la table en cours}
      St4 := ReadTokenSt(St2);
      St5 := ReadTokenSt(St3);
      if (St5 <> '-') and (St5 <> '#') and (St5 <> '') then begin
        CB .Checked := True;
        if (fTypeTableLibre = 'SCL') then eDe.Enabled := True
        else cDe.Enabled := True;
        SetControlEnabled('TDET0' + n, True);
        if (St5 <> '*') then
        begin
          if (fTypeTableLibre = 'SCL') then eA.Enabled := True
          else cA.Enabled := True;
          LA .Enabled := True;
          SetControlEnabled('TAT0' + n, True);
          if (fTypeTableLibre = 'SCL') then
          begin
            eDe.Text   := St5;
            eA .Text   := St4;
          end else
          begin
            cDe.Value   := St5;
            cA .Value   := St4;
          end;
        end;
      end;
      if (fTypeTableLibre <> 'SCL') then
      begin
        cDe.OnChange := TDeOnChange;
        cA .OnChange := TaOnChange;
      end;
    end ;
  end ;
  LesLib.Free ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TLTIERSRUPT.TDeOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Cb : THValComboBox;
  n  : string;
  Ok : Boolean;
begin
  cb := THValComboBox(Sender) ;
  n  := IntToStr(cb.Tag);
  {Si on pose un filtre sur la table libre en cours}
  Ok := GetControlText('DETABLET0' + n) <> '';

  {Si la valeur de fin est inférieure à celle de début}
  if Ok and (GetControlText('ATABLET0' + n) < Cb.Value) then
    SetControlText('ATABLET0' + n, Cb.Value)
  {On ne filtre pas, on ne renseigne pas la valeur de fin}
  else if not Ok then
    SetControlText('ATABLET0' + n, '');

  {Les zones de fin ne sont actives que si la valeur de départ est différente de <<Tous>>}
  SetControlEnabled('TAT0'     + n, Ok);
  SetControlEnabled('ATABLET0' + n, Ok);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TLTIERSRUPT.TaOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Cb : THValComboBox;
  n  : string;
begin
  cb := THValComboBox(Sender) ;
  n  := IntToStr(cb.Tag);
  {Si la valeur de début est supérieure à celle de fin}
  if GetControlText('DETABLET0' + n) > Cb.Value then
    SetControlText('DETABLET0' + n, Cb.Value);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TLTIERSRUPT.T00OnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  CB : TCheckBox;
  n  : string;
begin
  CB := TCheckBox(Sender);
  n  := IntToStr(CB.Tag);

  if not CB.Checked then
  begin
    if (fTypeTableLibre = 'SCL') then
    begin
      SetControlText('DETABLER0' + n, '');
      SetControlText('ATABLER0'  + n, '');
    end else
    begin
      SetControlText('DETABLET0' + n, '');
      SetControlText('ATABLET0'  + n, '');
    end;
  end;
  {Les zones de début sont active que si l'on filtre sur la table libre en cours}
  if (fTypeTableLibre = 'SCL') then
    SetControlEnabled('DETABLER0' + n, CB.Checked)
  else SetControlEnabled('DETABLET0' + n, CB.Checked);
  SetControlEnabled('TLIBT0'    + n, CB.Checked);
  SetControlEnabled('TDET0'     + n, CB.Checked);
  {Par contre, les zones de fin sont inactives : elles seront au besoin activées sur
   le OnChange de la combo de début en fonction de sa valeur}
  if (fTypeTableLibre = 'SCL') then
    SetControlEnabled('TAT0'      + n, True)
  else
    SetControlEnabled('TAT0'      + n, False) ;
  if (fTypeTableLibre = 'SCL') then
    SetControlEnabled('ATABLER0'  + n, True)
  else SetControlEnabled('ATABLET0'  + n, False);
end;

procedure TOF_TLTIERSRUPT.InitTablesLibres;
var i : integer;
    stIndiceIni, stIndiceDest : string;
    stTablette : string;
begin
  if fTypeTableLibre = 'SCL' then stTablette := 'AFLRESSOURCE'
  else if fTypeTableLibre = 'CLI' then stTablette :=  'GCLIBRETIERS'
  else stTablette :=  'TZNATTIERS';
  for i := 0 to 9 do
  begin
    stIndiceIni := Format('%2.2d',[i]);
    if fTypeTableLibre = 'SCL' then stIndiceDest := '' else
    begin
      if fTypeTableLibre = 'AUX' then stIndiceDest := IntToStr(i)
      else stIndiceDest := IntToHex(i+1,1);
    end;
    SetControlProperty('DETABLET'+stIndiceIni,'DataType',stTablette+stIndiceDest);
    SetControlProperty('ATABLET'+stIndiceIni,'DataType',stTablette+stIndiceDest);
  end;
end;

Procedure GetLibelleTableLibreTiersCompl (stPref : String ; LesLib : HTStringList ) ;
Var QLoc : TQuery ;
    St : string;
BEGIN
  QLoc:=OpenSql('Select CC_LIBELLE,CC_ABREGE,CC_CODE From CHOIXCOD Where CC_TYPE="ZLT" And CC_CODE Like "'+stPref+'%" Order by CC_CODE',True) ;
  LesLib.Clear ;
  While Not QLoc.Eof do
  begin
    if QLoc.Fields[0].AsString='.-' then St := '-' else St := 'X';
    LesLib.Add(QLoc.Fields[0].AsString+';'+St) ;
    QLoc.Next ;
  end;
  Ferme(QLoc) ;
END ;


initialization
  RegisterClasses([TOF_TLTIERSRUPT]) ;

end.

