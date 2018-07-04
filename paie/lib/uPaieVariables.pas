{***********UNITE*************************************************
Auteur  ...... : D.T.
Créé le ...... : 22/02/2003
Modifié le ... :   /  /
Description .. : Unité de gestion des variables de paie.
Mots clefs ... :
*****************************************************************}
unit uPaieVariables;

interface

uses CLasses,
     SysUtils,
{$ifndef eaglclient}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$endif}
     hctrls,
     uTob ;

     function initTOB_Variables() : TOB ;
     function TOB_Variables() : TOB ;
     procedure Nettoyage_Variables() ;
     
implementation

var
   LocaleTob: TOB ;
   DateModif: TDateTime ;

////////////////////////////////////////////////////////////////////////////////

function initTOB_Variables() : TOB ;
var q: tquery ;
    t: tob ;
    tfind: tob ;
    st : String;
begin

     t         := nil ;
     // Si premier appel
     if not assigned(LocaleTob) then
        begin
        LocaleTob := Tob.Create('variable_de_la_paie',Nil,-1) ;
        DateModif := Now ;
{Flux optimisé
        q := opensql('SELECT * FROM VARIABLEPAIE WHERE ##PVA_PREDEFINI## ORDER BY PVA_VARIABLE',True) ;
        if not q.eof then LocaleTob.LoadDetailDb('VARIABLEPAIE','','',q,False) ;
        Ferme(q) ;
}
          st := 'SELECT * FROM VARIABLEPAIE WHERE ##PVA_PREDEFINI## ORDER BY PVA_VARIABLE';
          LocaleTob.LoadDetailDbFromSQL('VARIABLEPAIE',st) ;
        end
     else
        begin
        q := opensql('SELECT * FROM VARIABLEPAIE WHERE ##PVA_PREDEFINI## AND PVA_DATEMODIF>="'+UsTime(DateModif)+'" ORDER BY PVA_VARIABLE',True) ;
        DateModif := Now ;
        if not q.eof then
           begin
           t := tob.create('',nil,-1) ;
           T.LoadDetailDB('VARIABLEPAIE','','',q,false) ;

           // des variables ont été modififées ??
           while t.detail.count>0 do
                    begin
                    tfind := localetob.FindFirst(['PVA_VARIABLE'],[t.detail[0].getvalue('PVA_VARIABLE')],false);
                    if assigned(tfind) then FreeAndNil(tfind) ;
                    t.Detail[0].ChangeParent(LocaleTob,-1) ;
                    end ;

           // Tri !!
           LocaleTob.Detail.Sort('PVA_VARIABLE');

           FreeAndNil(t) ;
           end ;
        Ferme(q) ;
        end ;

     Result := LocaleTob ;
end ;

function TOB_Variables() : TOB ;
begin
     Result := LocaleTob ;
end ;

////////////////////////////////////////////////////////////////////////////////

procedure Init() ;
begin
   LocaleTob:=Nil ;
end ;

procedure Nettoyage_Variables() ;
begin
     if Assigned(LocaleTob) then FreeAndNil(LocaleTob) ;
end ;

////////////////////////////////////////////////////////////////////////////////

initialization
  init() ;

finalization
  Nettoyage_Variables() ;


end.
