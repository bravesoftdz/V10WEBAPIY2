{***********UNITE*************************************************
Auteur  ...... : FC
Créé le ...... : 27/02/2007
Modifié le ... :   /  /
Description .. : Unité qui permet de récupérer le niveau requis pour un élément national
Mots clefs ... :
*****************************************************************}
unit uPaieEltNiveauRequis;

interface

uses CLasses,
  SysUtils,
{$IFNDEF eaglclient}
  db,
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
  hctrls,
  UTob;

function initTOB_EltNiveauRequis(): Tob;
function TOB_EltNiveauRequis(): Tob;
procedure Nettoyage_EltNiveauRequis();

implementation

var
  LocaleTob: Tob;
  DateModif: TDateTime;

// Fonction qui renvoie quel est le niveau requis pour un élément national

function initTOB_EltNiveauRequis(): Tob;
var q: tquery;
  t: tob;
  tfind: tob;
  St: string;
begin
  t := nil;
     // Si premier appel
  if not assigned(LocaleTob) then
  begin
    LocaleTob := Tob.Create('EltNiveauRequis', nil, -1);
    DateModif := Now;
{Flux optimisé
        Q := OpenSql('SELECT * FROM ELTNIVEAUREQUIS '
          + ' WHERE ##PNR_PREDEFINI##'
          + ' ORDER BY PNR_CODEELT',True) ;
        if not q.eof then LocaleTob.LoadDetailDb('ELTNIVEAUREQUIS','','',q,False) ;
        Ferme(q) ;
}
    St := 'SELECT * FROM ELTNIVEAUREQUIS '
      + ' WHERE ##PNR_PREDEFINI##'
      + ' ORDER BY PNR_CODEELT,PNR_PREDEFINI DESC';
    LocaleTob.LoadDetailDbfromSQL('ELTNIVEAUREQUIS', st);
  end
  else
  begin
    Q := OpenSql('SELECT * FROM ELTNIVEAUREQUIS '
      + ' WHERE ##PNR_PREDEFINI##'
      + ' AND PNR_DATEMODIF>="' + UsTime(DateModif) + '"'
      + ' ORDER BY PNR_CODEELT,PNR_PREDEFINI DESC', True);
    DateModif := Now;
    if not q.eof then
    begin
      t := tob.create('', nil, -1);
      T.LoadDetailDB('ELTNIVEAUREQUIS', '', '', q, false);

           // des variables ont été modifiées ??
      while t.detail.count > 0 do
      begin
        tfind := LocaleTob.FindFirst(['PNR_CODEELT'], [t.detail[0].getvalue('PNR_CODEELT'), t.detail[0].getvalue('PED_DATEVALIDITE')], false);
        if assigned(tfind) then FreeAndNil(tfind);
        t.Detail[0].ChangeParent(LocaleTob, -1);
      end;

           // Tri !!
      LocaleTob.Detail.Sort('PNR_CODEELT'); //PT1
      FreeAndNil(t);
    end;
    Ferme(q);
  end;
  Result := LocaleTob;
end;

function TOB_EltNiveauRequis(): Tob;
begin
  Result := LocaleTob;
end;

procedure Init();
begin
  LocaleTob := nil;
end;

procedure Nettoyage_EltNiveauRequis();
begin
  if Assigned(LocaleTob) then FreeAndNil(LocaleTob);
end;


initialization
  init();

finalization
  Nettoyage_EltNiveauRequis();
end.

