{-------------------------------------------------------------------------------------
  Version    |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
6.50.001.004  08/06/05   JP    Création de l'unité en remplacement de CLEREPAR.PAS
--------------------------------------------------------------------------------------}
unit REPARCLE_TOM ;

interface

uses
  {$IFDEF VER150}
    Variants,
  {$ENDIF}
  {$IFDEF EAGLCLIENT}
    MaineAGL,
    UtileAGL,
    eFichList,
  {$ELSE}
    db,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    FE_Main,
    FichList,
  {$ENDIF}
  StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, HCtrls, HEnt1, UTOM, UTob,
  Windows;

type
  TOM_CLEREPAR = class (TOM)
    procedure OnArgument   (S : string); override;
    procedure OnDeleteRecord           ; override;
    procedure OnChangeField(F : TField); override;
    procedure OnNewRecord              ; override;
    procedure OnUpdateRecord           ; override;
    procedure OnClose                  ; override;
  private
    FAvertir : Boolean;

    function GetAxe : string;
    function IsInserting : Boolean;
    function IsModifying : Boolean;
    function IsAllEmpty  : Boolean;
  protected
    {$IFDEF EAGL}
    cbMode : THValComboBox;
    cbQual : THValComboBox;
    {$ELSE}
    cbMode : THDBValComboBox;
    cbQual : THDBValComboBox;
    {$ENDIF EAGL}
    pcAxes : TPageControl;
    procedure BVentilClick  (Sender : TObject);
    procedure pcAxesChange  (Sender : TObject);
    procedure pcAxesChanging(Sender : TObject; var AllowChange: Boolean);
  public
    property Axe : string read GetAxe;
  end;

procedure ParamRepartCle;

Implementation

uses
  HMsgBox, HTB97, Ventil;

{---------------------------------------------------------------------------------------}
procedure ParamRepartCle;
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('CP', 'CPREPARCLE', 'A1', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CLEREPAR.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 1455100;
  {$IFDEF EAGL}
  cbMode := THValComboBox(GetControl('RE_MODEREPARTITION'));
  cbQual := THValComboBox(GetControl('RE_QUALIFQTE'));
  {$ELSE}
  cbMode := THDBValComboBox(GetControl('RE_MODEREPARTITION'));
  cbQual := THDBValComboBox(GetControl('RE_QUALIFQTE'));
  {$ENDIF EAGL}
  FAvertir := False;

  pcAxes := TPageControl(GetControl('PCAXES'));
  pcAxes.OnChange   := pcAxesChange;
  pcAxes.OnChanging := pcAxesChanging;
  {En s3, il n'ya qu'un seul axe}
  if EstSerie(S3) then
    pcAxes.Visible := False;

  TToolbarButton97(GetControl('BVENTIL')).OnClick := BVentilClick;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CLEREPAR.OnChangeField(F: TField);
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOM_CLEREPAR.OnNewRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  SetField('RE_CLE'    , '');
  SetField('RE_LIBELLE', '');
  SetField('RE_ABREGE' , '');
  SetField('RE_MODEREPARTITION', cbMode.Values[0]);
  SetField('RE_QUALIFQTE'      , cbQual.Values[0]);
  SetField('RE_AXE'            , Axe);
  SetField('RE_SOCIETE'        , V_PGI.CodeSociete);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CLEREPAR.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if VarToStr(GetField('RE_CLE')) = '' then begin
    LastError := 2;
    LastErrorMsg := 'Vous devez renseigner un code';
    SetFocusControl('RE_CLE');
  end
  else if VarToStr(GetField('RE_LIBELLE')) = '' then begin
    LastError := 2;
    LastErrorMsg := 'Vous devez renseigner un libellé';
    SetFocusControl('RE_LIBELLE');
  end
  else if VarToStr(GetField('RE_AXE')) = '' then begin
    LastError := 2;
    LastErrorMsg := 'L''axe n''a pu être récupéré';
  end
  else if VarToStr(GetField('RE_MODEREPARTITION')) = '' then begin
    LastError := 2;
    LastErrorMsg := 'Vous devez renseigner un mode de répartition';
    SetFocusControl('RE_MODEREPARTITION');
  end;
  FAvertir := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CLEREPAR.OnDeleteRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if HShowMessage('1;' + Ecran.Caption + ';Confirmez-vous la suppression de l''enregistrement ?;Q;YNC;N;C;', '', '') = mrYes then
  ExecuteSql('DELETE FROM VENTIL WHERE V_NATURE = "CL' + IntToStr(pcAxes.ActivePageIndex + 1) +
             '" AND V_COMPTE = "' + VarToStr(GetField('RE_CLE')) + '"');
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CLEREPAR.OnClose;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if FAvertir then begin
    AvertirTable('TTCLEREPART1');
    AvertirTable('TTCLEREPART2');
    AvertirTable('TTCLEREPART3');
    AvertirTable('TTCLEREPART4');
    AvertirTable('TTCLEREPART5');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CLEREPAR.BVentilClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  St : string;
begin
  if GetField('RE_CLE') <> '' then begin
    St := IntToStr(pcAxes.ActivePageIndex + 1);
    {Appel aux ventilations types}
    ParamVentil('CL', VarToStr(GetField('RE_CLE')), St, taModif, True);
  end;
end;

{---------------------------------------------------------------------------------------}
function TOM_CLEREPAR.GetAxe: string;
{---------------------------------------------------------------------------------------}
begin
  {L'axe courant est fonction de l'onglet courant}
  Result := 'A' + IntToStr(pcAxes.ActivePageIndex + 1);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CLEREPAR.pcAxesChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {On filtre la liste sur l'axe courant qui dépend de l'onglet courant}
  TFFicheListe(Ecran).SetNewRange(Axe);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_CLEREPAR.pcAxesChanging(Sender : TObject; var AllowChange : Boolean);
{---------------------------------------------------------------------------------------}
begin
  {On autorise le changement si l'on n'est pas en modifiction ou bien si l'on a validé ou
   annulé les modifications}
  AllowChange := not IsInserting;
end;

{---------------------------------------------------------------------------------------}
function TOM_CLEREPAR.IsInserting : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := DS.Modified and (DS.State in [dsInsert]);
  {Si l'on est en train d'insérer un nouvel enregistrement ...}
  if Result then begin
    {S'il n'y a pas d'enregistrements sur l'axe courant, la fiche est passé automatiquement
     en mode insertion. Si tous les champs sont vides, c'est que l'on n'a rien saisi. Dans
     ce cas on bidouille pour rendre la chose invisible à l'utilisateur ...}
    if IsAllEmpty then begin
      Result := False;
      DS.State := dsBrowse;
    end
    else
      {... Sinon on propose d'enregistrer les modifications}
      Result := IsModifying;
  end

  {... Si l'on est en train de faire des modifications...}
  else if DS.Modified and (DS.State in [dsEdit]) then
    {... Sinon on propose d'enregistrer les modifications}
    Result := IsModifying;
end;

{---------------------------------------------------------------------------------------}
function TOM_CLEREPAR.IsModifying : Boolean;
{---------------------------------------------------------------------------------------}
var
  mRes : TModalResult;
begin
  Result := DS.Modified;
  {Si on est en train de faire des modification ...}
  if Result then begin
    mRes := HShowMessage('1;' + Ecran.Caption + ';Voulez-vous enregistré vos modifications ?;Q;YNC;N;C;', '', '');
    Result := False;
    {Demande d'enregistrement des modifications}
    if mRes = mrYes then
      TFFicheListe(Ecran).BValider.Click
    {Demande d'annulation des modifications}
    else if mRes = mrNo then
      TFFicheListe(Ecran).bDefaire.Click
    {Abandon du changement d'onglet ...}
    else if mRes = mrCancel then
      {... on renvoie True pour ne pas permettre le changement d'onglet}
      Result := True;
  end;
end;

{---------------------------------------------------------------------------------------}
function TOM_CLEREPAR.IsAllEmpty: Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := (VarToStr(GetField('RE_CLE')) = '') and
            (VarToStr(GetField('RE_LIBELLE')) = '') and
            (VarToStr(GetField('RE_LIBELLE')) = '') and
            (VarToStr(GetField('RE_COMPTES')) = '');
end;

initialization
  RegisterClasses([TOM_CLEREPAR]);

end.

