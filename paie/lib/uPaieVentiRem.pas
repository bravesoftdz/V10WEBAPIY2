{***********UNITE*************************************************
Auteur  ...... : D.T.
Créé le ...... : 22/02/2003
Modifié le ... :   /  /
Description .. : Unité de gestion des ventilations des rémunérations de paie.
Mots clefs ... :
*****************************************************************}
unit uPaieVentiRem;

interface

uses CLasses,
  SysUtils,
{$IFNDEF eaglclient}
  db,
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
  hctrls,
  uTob;

function initTOB_VentiRemPaie(): TOB;
function TOB_VentilRem(): TOB;
procedure Nettoyage_VentilRem();

implementation

var
  LocaleTob: TOB;
  DateModif: TDateTime;

////////////////////////////////////////////////////////////////////////////////

function initTOB_VentiRemPaie(): TOB;
var q: tquery;
  t: tob;
  tfind: tob;
  st: string;
begin

  t := nil;
     // Si premier appel
  if not assigned(LocaleTob) then
  begin
    LocaleTob := Tob.Create('VentilRem_de_la_paie', nil, -1);
    DateModif := Now;
{Flux optimisé
        q := opensql ('SELECT * from VENTIREMPAIE WHERE ##PVS_PREDEFINI## ORDER BY PVS_RUBRIQUE',True) ;
        if not q.eof then LocaleTob.LoadDetailDb ('VENTIREMPAIE','','',q,False) ;
        Ferme(q) ;
}
    st := 'SELECT * from VENTIREMPAIE WHERE ##PVS_PREDEFINI## ORDER BY PVS_RUBRIQUE';
    LocaleTob.LoadDetailDbFromSQL('VENTIREMPAIE', st);
  end
  else
  begin
    q := opensql('select * from VENTIREMPAIE where ##PVS_PREDEFINI## AND PVS_DATEMODIF>="' + UsTime(DateModif) +
      '" ORDER BY PVS_RUBRIQUE', True);
    DateModif := Now;
    if not q.eof then
    begin
      t := tob.create('', nil, -1);
      T.LoadDetailDB('VENTIREMPAIE', '', '', q, false);

           // des variables ont été modififées ??
      while t.detail.count > 0 do
      begin
        tfind := localetob.FindFirst(['PVS_RUBRIQUE'], [t.detail[0].getvalue('PVS_RUBRIQUE')], false);
        if assigned(tfind) then FreeAndNil(tfind);
        t.Detail[0].ChangeParent(LocaleTob, -1);
      end;

           // Tri !!
      LocaleTob.Detail.Sort('PVS_RUBRIQUE');

      FreeAndNil(t);
    end;
    Ferme(q);
  end;

  Result := LocaleTob;
end;

function TOB_VentilRem(): TOB;
begin
  Result := LocaleTob;
end;

////////////////////////////////////////////////////////////////////////////////

procedure Init();
begin
  LocaleTob := nil;
end;

procedure Nettoyage_VentilRem();
begin
  if Assigned(LocaleTob) then FreeAndNil(LocaleTob);
end;

////////////////////////////////////////////////////////////////////////////////

initialization
  init();

finalization
  Nettoyage_VentilRem();


end.

