{-------------------------------------------------------------------------------------
    Version    |   Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 6.0X.xxx.xxx    02/11/04    JP   Création de l'unité : Tom de saisie des différentes OPCVM
--------------------------------------------------------------------------------------}
unit TROPCVMREF_TOM;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL, UTob, 
  {$ELSE}
   db, FE_Main,
  {$ENDIF}
  Controls, Classes, HCtrls, HEnt1, UTOM;

type
  TOM_TROPCVMREF = class(TOM)
    procedure OnUpdateRecord           ; override;
    procedure OnArgument   (S : string); override;
    procedure OnNewRecord              ; override ;
    procedure OnDeleteRecord           ; override;
    procedure OnChangeField(F : TField); override;
  end ;

procedure TRLanceFiche_OPCVMREF(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_OPCVMREF(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVMREF.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVMREF.OnDeleteRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if ExisteSQL('SELECT TOP_CODEOPCVM FROM TROPCVM WHERE TOP_CODEOPCVM = "' + GetField('TOF_CODEOPCVM') + '"') then begin
    LastError := 2;
    LastErrorMsg := TraduireMemoire('Cet OPCVM est utilisé en gestion financière et ne peut être supprimé.');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVMREF.OnChangeField(F : TField);
{---------------------------------------------------------------------------------------}
begin
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVMREF.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  LastErrorMsg := '';
  if GetField('TOF_PORTEFEUILLE') = '' then
    LastErrorMsg := TraduireMemoire('Il est obligatoire de rattacher l''OPCVM à un portefeuille.')
  else if GetField('TOF_DEVISE') = '' then
    LastErrorMsg := TraduireMemoire('Veuillez choisir une devise.')
  else if GetField('TOF_LIBELLE') = '' then
    LastErrorMsg := TraduireMemoire('Veuillez renseigner le libellé.');
  if LastErrorMsg <> '' then LastError := 2;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVMREF.OnNewRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {27/04/05 : FQ 10244 : s'il y a un filtre sur la tablette TROPCVM au moment du DispatchTT,
              il est repris dans le champ de l'index}
  SetField('TOF_CODEOPCVM', '');
end;

initialization
  RegisterClasses([TOM_TROPCVMREF]);

end.

