{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. : Unité de gestion des éléments nationaux
Mots clefs ... :
*****************************************************************}
{
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
}


unit uPaieEltNatDOS;

interface

uses CLasses,
  SysUtils,
  {$IFNDEF eaglclient}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  hctrls,
  uTob;

function initTOB_EltNatDOS(): TOB;
function TOB_EltNationauxDOS(): TOB;
procedure Nettoyage_EltNationauxDOS();

implementation
uses
  P5Def,PgOutils2;
var
  LocaleTob: TOB;
  DateModif: TDateTime;

  ////////////////////////////////////////////////////////////////////////////////

function initTOB_EltNatDOS(): TOB;
var
  q: tquery;
  t: tob;
  tfind: tob;
  st : String;
begin

  t := nil;
  // Si premier appel
  if not assigned(LocaleTob) then
  begin
    LocaleTob := Tob.Create('EltNat_DOS', nil, -1);
    DateModif := Now;
{Flux optimisé
    // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
    q := opensql('SELECT * FROM ELTNATIONAUX WHERE PEL_PREDEFINI = "DOS" AND PEL_NODOSSIER = "' + PgRendNoDossier() + '" ORDER BY PEL_CODEELT,PEL_DATEVALIDITE', True);
    DateModif := Now;
    if not q.eof then LocaleTob.LoadDetailDb('ELTNATIONAUX', '', '', q, False);
    Ferme(q);
}
    st := 'SELECT * FROM ELTNATIONAUX WHERE PEL_PREDEFINI = "DOS" AND PEL_NODOSSIER = "' + PgRendNoDossier() + '" ORDER BY PEL_CODEELT,PEL_DATEVALIDITE';
    LocaleTob.LoadDetailDbFromSQL('ELTNATIONAUX', st);
  end
  else
  begin
    q := opensql('SELECT * FROM ELTNATIONAUX WHERE PEL_PREDEFINI = "DOS" AND PEL_NODOSSIER = "' + PgRendNoDossier() + '" AND PEL_DATEMODIF>="' + UsTime(DateModif) +
      '" ORDER BY PEL_CODEELT,PEL_DATEVALIDITE', True);
    DateModif := Now;
    if not q.eof then
    begin
      t := tob.create('', nil, -1);
      T.LoadDetailDB('ELTNATIONAUX', '', '', q, false);

      // des variables ont été modififées ??
      while t.detail.count > 0 do
      begin
        tfind := localetob.FindFirst(['PEL_CODEELT', 'PEL_DATEVALIDITE'], [t.detail[0].getvalue('PEL_CODEELT'), t.detail[0].getvalue('PEL_DATEVALIDITE')], false);
        if assigned(tfind) then FreeAndNil(tfind);
        t.Detail[0].ChangeParent(LocaleTob, -1);
      end;

      // Tri !!
      LocaleTob.Detail.Sort('PEL_CODEELT;PEL_DATEVALIDITE');
      FreeAndNil(t);
    end;
    Ferme(q);
  end;
  Result := LocaleTob;
end;

function TOB_EltNationauxDOS(): TOB;
begin
  Result := LocaleTob;
end;

////////////////////////////////////////////////////////////////////////////////

procedure Init();
begin
  LocaleTob := nil;
end;

procedure Nettoyage_EltNationauxDOS();
begin
  if Assigned(LocaleTob) then FreeAndNil(LocaleTob);
end;

////////////////////////////////////////////////////////////////////////////////

initialization
  init();

finalization
  Nettoyage_EltNationauxDOS();
end.

