{***********UNITE*************************************************
Auteur  ...... : FC
Créé le ...... : 28/02/2007
Modifié le ... :   /  /
Description .. : Unité qui permet de récupérer la tob contenant les éléments dynamiques population
Mots clefs ... :
*****************************************************************}
unit uPaieEltDynPop ;

interface

uses CLasses,
     SysUtils,
{$ifndef eaglclient}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$endif}
     hctrls,
{$IFNDEF CPS1}
{$IFNDEF EAGLSERVER}
     PGPOPULOUTILS,
{$ENDIF}
{$ENDIF}
     UTob ;

    function initTOB_EltDynPop(CodeSal:String;DateFin:TDateTime) : Tob;
    function TOB_EltDynPop() : Tob;
    procedure Nettoyage_EltDynPop();

implementation

var
  LocaleTob : Tob;
  DateModif : TDateTime ;

function initTOB_EltDynPop(CodeSal:String;DateFin:TDateTime) : Tob;
var q: tquery ;
    t: tob ;
    tfind: tob ;
    Population:string;
begin

     t         := nil ;
{$IFNDEF CPS1}
{$IFNDEF EAGLSERVER}
    Population := '';
    Q := OpenSQL('SELECT PNA_POPULATION FROM SALARIEPOPUL '
      + ' WHERE PNA_SALARIE = "' + CodeSal + '"'
      + ' AND PNA_TYPEPOP = "PAI"', True);  //Anciennement SAL
    if not Q.Eof then
      Population := Q.FindField('PNA_POPULATION').AsString;
    Ferme(Q);
{$ENDIF}
{$ENDIF}
     // Si premier appel
     if not assigned(LocaleTob) then
        begin
        LocaleTob := Tob.Create('EltDynPop',Nil,-1) ;
        Q := OpenSql('SELECT * FROM ELTNATIONDOS '
          + ' WHERE PED_TYPENIVEAU="POP"'
          + ' AND PED_VALEURNIVEAU="' + Population + '"'
          + ' AND PED_DATEVALIDITE<="' + UsTime(DateFin) + '"'
          + ' ORDER BY PED_CODEELT,PED_DATEVALIDITE',True) ;
        DateModif := Now ;
        if not q.eof then LocaleTob.LoadDetailDb('ELTNATIONDOS','','',q,False) ;
        Ferme(q) ;
        end
     else
        begin
        q := opensql('SELECT * FROM ELTNATIONDOS'
          + ' WHERE PED_TYPENIVEAU = "POP" '
          + ' AND PED_VALEURNIVEAU="' + Population + '"'
          + ' AND PED_DATEVALIDITE<="' + UsTime(DateFin) + '"'
          + ' AND PED_DATEMODIF>="'+UsTime(DateModif)+'"'
          + ' ORDER BY PED_CODEELT,PED_DATEVALIDITE',True) ;
        DateModif := Now ;
        if not q.eof then
           begin
           t := tob.create('',nil,-1) ;
           T.LoadDetailDB('ELTNATIONDOS','','',q,false) ;

           // des variables ont été modifiées ??
           while t.detail.count>0 do
                 begin
                 tfind := LocaleTob.FindFirst(['PED_CODEELT','PED_DATEVALIDITE'],[t.detail[0].getvalue('PED_CODEELT'),t.detail[0].getvalue('PED_DATEVALIDITE')],false);
                 if assigned(tfind) then FreeAndNil(tfind) ;
                 t.Detail[0].ChangeParent(LocaleTob,-1) ;
                 end ;

           // Tri !!
           LocaleTob.Detail.Sort('PED_CODEELT;PED_DATEVALIDITE'); //PT1
           FreeAndNil(t) ;
           end ;
        Ferme(q) ;
        end ;
     Result := LocaleTob ;
end;

// Fonction qui renvoie dans une TOB, la liste des éléments dynamiques des population
function TOB_EltDynPop() : Tob;
begin
     Result := LocaleTob ;
end;
                             
procedure Init() ;
begin
   LocaleTob:=Nil ;
end ;

procedure Nettoyage_EltDynPop();
begin
     if Assigned(LocaleTob) then FreeAndNil(LocaleTob) ;
end;

initialization
  init() ;

finalization
  Nettoyage_EltDynPop() ;
end.
