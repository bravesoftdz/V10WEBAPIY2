{***********UNITE*************************************************
Auteur  ...... : D.T.
Créé le ...... : 22/02/2003
Modifié le ... :   /  /
Description .. : Unité de gestion des profils de paie.
Mots clefs ... :
*****************************************************************}
{
PT1   : 18/06/2007 VG V_72 Mauvaise libération mémoire
}
unit uPaieProfilsPaie;

interface

uses CLasses,
  SysUtils,
{$IFNDEF eaglclient}
  db,
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
  hctrls,
  uTob;

function initTOB_ProfilPaies(): TOB;
function TOB_ProfilPaies(): TOB;
procedure Nettoyage_ProfilPaies();

implementation

var
  LocaleTob: TOB;
  DateModif: TDateTime;
  DateModifP: TDateTime;
  ////////////////////////////////////////////////////////////////////////////////

procedure Loc_ProfilRubrique(const all: boolean = false);
var
  st: string;
  // XP 10.10.2007 q: tquery;
  tobtmp: tob;
  t, tt: tob;
  profil, nature, rubrique: string;
  iPPM_NATURERUB, iPPM_RUBRIQUE, iPPM_PROFIL: Integer;
begin
  try
    // Les profils des rubriques
    if all then
      st := 'SELECT * FROM PROFILRUB WHERE ##PPM_PREDEFINI## ORDER BY PPM_PROFIL,PPM_NATURERUB,PPM_RUBRIQUE'
    else
      st := 'SELECT * FROM PROFILRUB WHERE ##PPM_PREDEFINI## AND PPM_DATEMODIF>="' + UsTime(DateModifP) + '" ORDER BY PPM_PROFIL,PPM_NATURERUB,PPM_RUBRIQUE';

    DateModifP := Now();
{Flux optimisé
    Q := OpenSql(st, TRUE);
    if not q.eof then
    begin
      tobtmp.LoadDetailDB('PROFILRUB', '', '', Q, FALSE, False);
      ...
      Ferme(Q);
}
    tobtmp := tob.create('tob_temporaire', nil, -1);
    tobtmp.LoadDetailDBfromSQL('PROFILRUB', St);
    if tobtmp.detail.count > 0 then
    begin
      iPPM_NATURERUB := tobtmp.detail[0].GetNumChamp('PPM_NATURERUB');
      iPPM_RUBRIQUE := tobtmp.detail[0].GetNumChamp('PPM_RUBRIQUE');
      iPPM_PROFIL := tobtmp.detail[0].GetNumChamp('PPM_PROFIL');
      while tobtmp.detail.count > 0 do
      begin
        Profil := Tobtmp.detail[0].GetValeur(iPPM_PROFIL);
        T := LocaleTob.FindFirst(['PPI_PROFIL'], [profil], FALSE);
        if assigned(t) then
        begin
          Nature := Tobtmp.detail[0].GetValeur(iPPM_NATURERUB);
          Rubrique := Tobtmp.detail[0].GetValeur(iPPM_RUBRIQUE);
            // est-ce que le profil rubrique existe déjà ?
          tt := t.FindFirst(['PPM_PROFIL', 'PPM_NATURERUB', 'PPM_RUBRIQUE'], [profil, Nature, Rubrique], False);
          if assigned(tt) then FreeAndNil(tt);
          Tobtmp.detail[0].ChangeParent(T, -1);
        end else Tobtmp.detail[0].free;
      end;
{    end
    else Ferme(q);
}
{PT1
      FreeAndNil(tobtmp);
}
    end;
  FreeAndNil(tobtmp);	//PT1
  except
    on E: Exception do
    begin
    end;
  end;
end;

function initTOB_ProfilPaies(): TOB;
var
  q: tquery;
  t: tob;
  tfind: tob;
  st: string;
begin

  t := nil;
  // Si premier appel
  if not assigned(LocaleTob) then
  begin
    LocaleTob := Tob.Create('profil_de_paie', nil, -1);
    DateModif := Now;
{Flux optimisé
    q := opensql('SELECT * FROM PROFILPAIE WHERE ##PPI_PREDEFINI## ORDER BY PPI_PROFIL', True);
    if not q.eof then
    begin
    Ferme(q);

}
    st := 'SELECT * FROM PROFILPAIE WHERE ##PPI_PREDEFINI## ORDER BY PPI_PROFIL';
    LocaleTob.LoadDetailDbFromSQL('PROFILPAIE', st);
    if LocaleTob.Detail.Count > 0 then
      Loc_ProfilRubrique(True);
  end
  else
  begin
    q := opensql('SELECT * FROM PROFILPAIE WHERE ##PPI_PREDEFINI## AND PPI_DATEMODIF>="' + UsTime(DateModif) + '" order by PPI_PROFIL', True);
    DateModif := Now;
    if not q.eof then
    begin
      t := tob.create('', nil, -1);
      T.LoadDetailDB('PROFILPAIE', '', '', q, false);

      // des profils de paie ont été modififés ?
      while t.detail.count > 0 do
      begin
        tfind := localetob.FindFirst(['PPI_PROFIL'], [t.detail[0].getvalue('PPI_PROFIL')], false);
        if assigned(tfind) then FreeAndNil(tfind);
        t.Detail[0].ChangeParent(LocaleTob, -1);
      end;

      // Tri !!
      LocaleTob.Detail.Sort('PPI_PROFIL');

      FreeAndNil(t);
    end;
    // chargement des profils de rubrique
// A faire systématiquement car on peut modifier les rubriques associées au profil
    Loc_ProfilRubrique();
    Ferme(q);
  end;

  Result := LocaleTob;
end;

function TOB_ProfilPaies(): TOB;
begin
  Result := LocaleTob;
end;

////////////////////////////////////////////////////////////////////////////////

procedure Init();
begin
  LocaleTob := nil;
end;

procedure Nettoyage_ProfilPaies();
begin
  if Assigned(LocaleTob) then FreeAndNil(LocaleTob);
end;

////////////////////////////////////////////////////////////////////////////////

initialization
  init();

finalization
  Nettoyage_ProfilPaies();


end.

