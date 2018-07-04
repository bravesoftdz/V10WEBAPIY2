{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. : Unité de gestion des tables minimum conventionnel de paie.
Mots clefs ... :
*****************************************************************}
unit uPaieMinimumConv;

interface

uses CLasses,
  SysUtils,
{$IFNDEF eaglclient}
  db,
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
  hctrls,
  uTob;

function initTOB_MinimumConv(): TOB;
function TOB_Minimum(): TOB;
procedure Nettoyage_Minimum();

implementation

var
  LocaleTob: TOB;
  DateModif: TDateTime;

////////////////////////////////////////////////////////////////////////////////


function initTOB_MinimumConv(): TOB;
var q: tquery;
  t: tob;
  tfind: tob;
  st: string;
begin

  t := nil;
     // Si premier appel
  if not assigned(LocaleTob) then
  begin
    LocaleTob := Tob.Create('MinimumConv_de_paie', nil, -1);
    DateModif := Now;
{Flux optimisé
        q := opensql('SELECT * FROM MINIMUMCONVENT WHERE ##PMI_PREDEFINI## AND PMI_TYPENATURE="INT" ORDER BY PMI_PREDEFINI,PMI_NODOSSIER,PMI_NATURE,PMI_CONVENTION,PMI_TYPENATURE,PMI_CODE',True) ;
        if not q.eof then LocaleTob.LoadDetailDb('MINIMUMCONVENT','','',q,False) ;
        Ferme(q) ;
}
    st := 'SELECT * FROM MINIMUMCONVENT WHERE ##PMI_PREDEFINI## AND PMI_TYPENATURE="INT" ORDER BY PMI_PREDEFINI,PMI_NODOSSIER,PMI_NATURE,PMI_CONVENTION,PMI_TYPENATURE,PMI_CODE';
    LocaleTob.LoadDetailDbFromSQL('MINIMUMCONVENT', st);
  end
  else
  begin
    q := opensql('SELECT * FROM MINIMUMCONVENT WHERE ##PMI_PREDEFINI##  AND PMI_TYPENATURE="INT" AND PMI_DATEMODIF>="' + UsTime(DateModif) + '" ORDER BY PMI_PREDEFINI,PMI_NODOSSIER,PMI_NATURE,PMI_CONVENTION,PMI_TYPENATURE,PMI_CODE', True);
    DateModif := Now;
    if not q.eof then
    begin
      t := tob.create('', nil, -1);
      T.LoadDetailDB('MINIMUMCONVENT', '', '', q, false);

           // des variables ont été modififées ??
      while t.detail.count > 0 do
      begin
        tfind := localetob.FindFirst(['PMI_PREDEFINI', 'PMI_NODOSSIER', 'PMI_NATURE', 'PMI_CONVENTION', 'PMI_TYPENATURE', 'PMI_CODE'],
          [t.detail[0].getvalue('PMI_PREDEFINI'), t.detail[0].getvalue('PMI_NODOSSIER'), t.detail[0].getvalue('PMI_NATURE'),
          t.detail[0].getvalue('PMI_CONVENTION'), t.detail[0].getvalue('PMI_TYPENATURE'), t.detail[0].getvalue('PMI_CODE')], false);
        if assigned(tfind) then FreeAndNil(tfind);
        t.Detail[0].ChangeParent(LocaleTob, -1);
      end;

           // Tri !!
      LocaleTob.Detail.Sort('PMI_PREDEFINI;PMI_NODOSSIER;PMI_NATURE;PMI_CONVENTION;PMI_TYPENATURE;PMI_CODE');
      FreeAndNil(t);
    end;
    Ferme(q);
  end;
  Result := LocaleTob;
end;

function TOB_Minimum(): TOB;
begin
  Result := LocaleTob;
end;

////////////////////////////////////////////////////////////////////////////////

procedure Init();
begin
  LocaleTob := nil;
end;

procedure Nettoyage_Minimum();
begin
  if Assigned(LocaleTob) then FreeAndNil(LocaleTob);
end;

////////////////////////////////////////////////////////////////////////////////

initialization
  init();

finalization
  Nettoyage_Minimum();
end.

