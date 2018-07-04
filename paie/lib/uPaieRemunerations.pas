{***********UNITE*************************************************
Auteur  ...... : D.T.
Créé le ...... : 22/02/2003
Modifié le ... :   /  /
Description .. : Unité de gestion des rubriques de rémunération.
Mots clefs ... :
*****************************************************************}
unit uPaieRemunerations;

interface

uses CLasses,
  SysUtils,
{$IFNDEF eaglclient}
  db,
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
  hctrls,
  uTob;

function initTOB_Rem(): TOB;
function TOB_Rem(): TOB;
procedure Nettoyage_Rem();

implementation

var
  LocaleTob: TOB;
  DateModif: TDateTime;

////////////////////////////////////////////////////////////////////////////////

function initTOB_Rem(): TOB;
var q: tquery;
  t: tob;
  tfind: tob;
  st: string;
begin

  t := nil;
     // Si premier appel
  if not assigned(LocaleTob) then
  begin
    LocaleTob := Tob.Create('rubrique_de_remuneration', nil, -1);
    DateModif := Now;
{Flux optimisé
        q := opensql ('SELECT * FROM REMUNERATION WHERE ##PRM_PREDEFINI## ORDER BY PRM_RUBRIQUE',True) ;
        if not q.eof then LocaleTob.LoadDetailDb ('REMUNERATION','','',q,False) ;
        Ferme(q) ;
}
    st := 'SELECT * FROM REMUNERATION WHERE ##PRM_PREDEFINI## ORDER BY PRM_RUBRIQUE';
    LocaleTob.LoadDetailDbFromSQL('REMUNERATION', St);
  end
  else
  begin
    q := opensql('SELECT * FROM REMUNERATION WHERE ##PRM_PREDEFINI## AND PRM_DATEMODIF>="' + UsTime(DateModif) + '" ORDER BY PRM_RUBRIQUE', True);
    DateModif := Now;
    if not q.eof then
    begin
      t := tob.create('', nil, -1);
      T.LoadDetailDB('REMUNERATION', '', '', q, false);

           // des variables ont été modififées ??
      while t.detail.count > 0 do
      begin
        tfind := localetob.FindFirst(['PRM_RUBRIQUE'], [t.detail[0].getvalue('PRM_RUBRIQUE')], false);
        if assigned(tfind) then FreeAndNil(tfind);
        t.Detail[0].ChangeParent(LocaleTob, -1);
      end;

           // Tri !!
      LocaleTob.Detail.Sort('PRM_RUBRIQUE');
      FreeAndNil(t);
    end;
    Ferme(q);
  end;

  Result := LocaleTob;
end;

function TOB_Rem(): TOB;
begin
  Result := LocaleTob;
end;

////////////////////////////////////////////////////////////////////////////////

procedure Init();
begin
  LocaleTob := nil;
end;

procedure Nettoyage_Rem();
begin
  if Assigned(LocaleTob) then FreeAndNil(LocaleTob);
end;

////////////////////////////////////////////////////////////////////////////////

initialization
  init();

finalization
  Nettoyage_Rem();


end.

