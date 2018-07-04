{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 27/05/2005
Modifié le ... :   /  /
Description .. : Unité de gestion des Execptions
Mots clefs ... :
*****************************************************************}
{
PT1  07/10/2005 PH V_60 FQ 12515 Tri de la tob pour suppression des exceptions STD si DOS existantes
}
unit uPaieExecpt;

interface

uses CLasses,
  SysUtils,
{$IFNDEF eaglclient}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  hctrls,
  uTob;

function initTOB_Execpt(): TOB;
function TOB_Execpt(): TOB;
procedure Nettoyage_Execpt();

implementation

var
  LocaleTob: TOB;
  DateModif: TDateTime;

////////////////////////////////////////////////////////////////////////////////

function initTOB_Execpt(): TOB;
var q: tquery;
  t: tob;
  tfind: tob;
  MNat, MRub, Mpre, st : string;
  i : Integer;
  Saute : Boolean;
begin

  t := nil;
     // Si premier appel
  if not assigned(LocaleTob) then
  begin
    LocaleTob := Tob.Create('les_exceptions', nil, -1);
    DateModif := Now;
{Flux optimisé
    q := opensql('SELECT * FROM PGEXCEPTIONS WHERE ##PEN_PREDEFINI## ORDER BY PEN_NATURERUB,PEN_PREDEFINI,PEN_RUBRIQUE', True);
    if not q.eof then LocaleTob.LoadDetailDb('PGEXCEPTIONS', '', '', q, False);
    Ferme(q);
}
    st :='SELECT * FROM PGEXCEPTIONS WHERE ##PEN_PREDEFINI## ORDER BY PEN_NATURERUB,PEN_PREDEFINI,PEN_RUBRIQUE';
    LocaleTob.LoadDetailDbFromSQL('PGEXCEPTIONS', st);
  end
  else
  begin
    q := opensql('SELECT * FROM PGEXCEPTIONS WHERE ##PEN_PREDEFINI## PEN_DATEMODIF>="' + UsTime(DateModif) + '" order by PEN_NATURERUB,PEN_PREDEFINI,PEN_RUBRIQUE', True);
    DateModif := Now;
    if not q.eof then
    begin
      t := tob.create('', nil, -1);
      T.LoadDetailDB('PGEXCEPTIONS', '', '', q, false);

           // des variables ont été modififées ??
      if assigned(t) then
      begin
        while t.detail.count > 0 do
        begin
          tfind := localetob.FindFirst(['PEN_NATURERUB', 'PEN_RUBRIQUE'], [t.detail[0].getvalue('PEN_NATURERUB'), t.detail[0].getvalue('PEN_RUBRIQUE')], false);
          if assigned(tfind) then FreeAndNil(tfind);
          t.Detail[0].ChangeParent(LocaleTob, -1);
        end;

              // Tri !!
        LocaleTob.Detail.Sort('PEN_NATURERUB;PEN_RUBRIQUE;PEN_PREDEFINI');
      end;
      FreeAndNil(t);
    end;
    Ferme(q);
  end;
  if Assigned(LocaleTob) then
  begin // Nettoyage de la tob des lignes STD si elles existent en DOS
    LocaleTob.Detail.Sort ('PEN_NATURERUB;PEN_RUBRIQUE;PEN_PREDEFINI');// PT1
    T := localetob.FindFirst([''], [''], False);
    i := 0;
    if T <> nil then
    begin
      MNat := T.GetValue('PEN_NATURERUB');
      MRub := T.GetValue('PEN_RUBRIQUE');
      Mpre := T.GetValue('PEN_PREDEFINI');
    end;
    while T <> nil do
    begin
      Saute := FALSE;
      if (i > 0) and (MNat = T.GetValue('PEN_NATURERUB')) and (MRub = T.GetValue('PEN_RUBRIQUE')) then
      begin
        if (Mpre <> T.GetValue('PEN_PREDEFINI')) and (T.GetValue('PEN_PREDEFINI') = 'STD') then
        begin
          Saute := TRUE;
          T.free;
        end;
      end;
      if not Saute then
      begin
        MNat := T.GetValue('PEN_NATURERUB');
        MRub := T.GetValue('PEN_RUBRIQUE');
        Mpre := T.GetValue('PEN_PREDEFINI');
      end;
      i := i + 1;
      T := localetob.FindNext([''], [''], False);
    end;
  end;
  Result := LocaleTob;
end;

function TOB_Execpt(): TOB;
begin
  Result := LocaleTob;
end;

////////////////////////////////////////////////////////////////////////////////

procedure Init();
begin
  LocaleTob := nil;
end;

procedure Nettoyage_Execpt();
begin
  if Assigned(LocaleTob) then FreeAndNil(LocaleTob);
end;

////////////////////////////////////////////////////////////////////////////////

initialization
  init();

finalization
  Nettoyage_Execpt();
end.

