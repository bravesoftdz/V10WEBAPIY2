{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. : Unité de gestion des cumuls de paie.
Mots clefs ... :
*****************************************************************}
{
PT1    : 13/03/2007 VG V_72 Gestion spécifique des cumuls
}
unit uPaieCumuls;

interface

uses CLasses,
  SysUtils,
{$IFNDEF eaglclient}
  db,
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
  hctrls,
  uTob;

function initTOB_Cumuls(): TOB;
function TOB_Cumuls(): TOB;
procedure Nettoyage_Cumuls();
procedure CreeFilleCumul(var T: TOB; TobTmp: TOB);


implementation
uses P5Util;

var
  LocaleTob: TOB;
  DateModif: TDateTime;
  DateModifC: TDateTime;
  DateModifCK: TDateTime;
  ////////////////////////////////////////////////////////////////////////////////

procedure Loc_CumulRubrique(const all: boolean = false);
var
  nat, rub, st: string;
//  q: tquery;
  t, t1, tobtmp: tob;
  iPCR_NATURERUB, iPCR_RUBRIQUE: Integer;
begin
  try
// Les cumuls des rubriques
    if all then
      st := 'SELECT *' +
        ' FROM CUMULRUBRIQUE WHERE' +
        ' ##PCR_PREDEFINI##' +
        ' ORDER BY PCR_NATURERUB, PCR_RUBRIQUE, PCR_CUMULPAIE'
    else
      st := 'SELECT *' +
        ' FROM CUMULRUBRIQUE WHERE' +
        ' ##PCR_PREDEFINI## AND' +
        ' PCR_DATEMODIF>="' + UsTime(DateModifC) + '"' +
        ' ORDER BY PCR_NATURERUB, PCR_RUBRIQUE, PCR_CUMULPAIE';
{Flux optimisé
q:= opensql (St, True);
DateModifC:= Now ();
if (not q.eof) then
   begin
   tobtmp:= tob.create ('cumul_rubrique', nil, -1);
   tobtmp.LoadDetailDB ('CUMULRUBRIQUE', '', '', Q, FALSE, False);
   Ferme(Q);
}
    tobtmp := tob.create('cumul_rubrique', nil, -1);
    tobtmp.LoadDetailDBFromSQL('CUMULRUBRIQUE', st);
    if (tobtmp.detail.count > 0) then
    begin
      iPCR_NATURERUB := tobtmp.detail[0].GetNumChamp('PCR_NATURERUB');
      iPCR_RUBRIQUE := tobtmp.detail[0].GetNumChamp('PCR_RUBRIQUE');
      while tobtmp.detail.count > 0 do
      begin
        Nat := tobtmp.detail[0].GetValeur(iPCR_NATURERUB);
        Rub := tobtmp.detail[0].GetValeur(iPCR_RUBRIQUE);
// rémunération ?
        t := Paie_RechercheRubrique(Nat, Rub);
        if assigned(t) then
        begin
// est-ce que le cumul est déjà rattaché ?
          t1 := t.FindFirst(['PCR_NATURERUB', 'PCR_RUBRIQUE', 'PCR_CUMULPAIE'],
            [nat, rub, tobtmp.detail[0].getvalue('PCR_CUMULPAIE')],
            false);
          if assigned(t1) then
            FreeAndNil(t1);
          tobtmp.detail[0].ChangeParent(T, -1);
        end
        else
          tobtmp.detail[0].free;
      end;
    end;
    FreeAndNil(tobtmp);

{
  end
  else
    Ferme(q);
}
//PT1
    St := 'SELECT *' +
      ' FROM CUMULRUBDOSSIER WHERE' +
      ' PKC_DATEMODIF>="' + UsTime(DateModifCK) + '"' +
      ' ORDER BY PKC_NATURERUB, PKC_RUBRIQUE, PKC_CUMULPAIE';
{Flux optimisé
    Q := OpenSql(St, True);
    DateModifCK := Now();
    if (not Q.eof) then
    begin
    TobTmp := Tob.Create('cumul_rubrique_dossier', nil, -1);
    TobTmp.LoadDetailDB('CUMULRUBDOSSIER', '', '', Q, FALSE, False);
    Ferme (Q);
}
    TobTmp := Tob.Create('cumul_rubrique_dossier', nil, -1);
    TobTmp.LoadDetailDBFromSQL('CUMULRUBDOSSIER', St);
    if (TobTmp.Detail.Count > 0) then
    begin
      iPCR_NATURERUB := TobTmp.Detail[0].GetNumChamp('PKC_NATURERUB');
      iPCR_RUBRIQUE := TobTmp.Detail[0].GetNumChamp('PKC_RUBRIQUE');
      while TobTmp.Detail.Count > 0 do
      begin
        Nat := TobTmp.Detail[0].GetValeur(iPCR_NATURERUB);
        Rub := TobTmp.Detail[0].GetValeur(iPCR_RUBRIQUE);
// rémunération ?
        T := Paie_RechercheRubrique(Nat, Rub);
        if Assigned(T) then
        begin
// est-ce que le cumul est déjà rattaché ?
          T1 := T.FindFirst(['PCR_NATURERUB', 'PCR_RUBRIQUE', 'PCR_CUMULPAIE'],
            [Nat, Rub, TobTmp.Detail[0].GetValue('PKC_CUMULPAIE')],
            False);
          if Assigned(T1) then
          begin
            if (TobTmp.Detail[0].GetValue('PKC_SENS') <> T1.GetValue('PCR_SENS')) then
            begin
              FreeAndNil(T1);
              CreeFilleCumul(T, TobTmp);
            end
            else
              FreeAndNil(T1);
          end
          else
            CreeFilleCumul(T, TobTmp);
        end;
        TobTmp.Detail[0].Free;
      end;
    end;
    FreeAndNil(tobtmp);
{
    end
    else
      Ferme(q);
//FIN PT1
}
  except
    on E: Exception do
    begin
    end;
  end;
end;

function initTOB_Cumuls(): TOB;
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
    LocaleTob := Tob.Create('cumuls_de_paie', nil, -1);
    DateModif := Now;
{Flux optimisé

    q := opensql('SELECT * FROM CUMULPAIE WHERE ##PCL_PREDEFINI## ORDER BY PCL_CUMULPAIE', True);
    if not q.eof then
    begin
      LocaleTob.LoadDetailDb('CUMULPAIE', '', '', q, False);
}
    st := 'SELECT * FROM CUMULPAIE WHERE ##PCL_PREDEFINI## ORDER BY PCL_CUMULPAIE';
    LocaleTob.LoadDetailDbFromSQL('CUMULPAIE', St);
    if LocaleTob.Detail.Count > 0 then
      Loc_CumulRubrique(True);
//    Ferme(q);
  end
  else
  begin
    q := opensql('SELECT * FROM CUMULPAIE WHERE ##PCL_PREDEFINI## AND PCL_DATEMODIF>="' + UsTime(DateModif) + '" ORDER BY PCL_CUMULPAIE', True);
    DateModif := Now;
    if not q.eof then
    begin
      t := tob.create('', nil, -1);
      T.LoadDetailDB('CUMULPAIE', '', '', q, false);

      // des variables ont été modififées ??
      while t.detail.count > 0 do
      begin
        tfind := localetob.FindFirst(['PCL_CUMULPAIE'], [t.detail[0].getvalue('PCL_CUMULPAIE')], false);
        if assigned(tfind) then FreeAndNil(tfind);
        t.Detail[0].ChangeParent(LocaleTob, -1);
      end;

      // Tri !!
      LocaleTob.Detail.Sort('PCL_CUMULPAIE');
      FreeAndNil(t);
    end;
    Ferme(q);
    // A faire systématiquement car on peut modifier les rubriques associées au cumul
    Loc_CumulRubrique();
  end;

  Result := LocaleTob;
end;

function TOB_Cumuls(): TOB;
begin
  Result := LocaleTob;
end;

////////////////////////////////////////////////////////////////////////////////

procedure Init();
begin
  LocaleTob := nil;
end;

procedure Nettoyage_Cumuls();
begin
  if Assigned(LocaleTob) then FreeAndNil(LocaleTob);
end;

//PT1

procedure CreeFilleCumul(var T: TOB; TobTmp: TOB);
var
  T1: TOB;
begin
  T1 := Tob.Create('CUMULRUBRIQUE', T, -1);
  T1.AddChampSupValeur('PCR_NATURERUB', TobTmp.Detail[0].GetValue('PKC_NATURERUB'));
  T1.AddChampSupValeur('PCR_RUBRIQUE', TobTmp.Detail[0].GetValue('PKC_RUBRIQUE'));
  T1.AddChampSupValeur('PCR_CUMULPAIE', TobTmp.Detail[0].GetValue('PKC_CUMULPAIE'));
  T1.AddChampSupValeur('PCR_PREDEFINI', 'DOS');
  T1.AddChampSupValeur('PCR_NODOSSIER', '');
  T1.AddChampSupValeur('PCR_LIBELLE', TobTmp.Detail[0].GetValue('PKC_LIBELLE'));
  T1.AddChampSupValeur('PCR_SENS', TobTmp.Detail[0].GetValue('PKC_SENS'));
  T1.AddChampSupValeur('PCR_DATECREATION', TobTmp.Detail[0].GetValue('PKC_DATECREATION'));
  T1.AddChampSupValeur('PCR_DATEMODIF', TobTmp.Detail[0].GetValue('PKC_DATEMODIF'));
end;
//FIN PT1

////////////////////////////////////////////////////////////////////////////////

initialization
  init();

finalization
  Nettoyage_Cumuls();
end.

