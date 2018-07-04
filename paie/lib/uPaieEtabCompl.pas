{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. : Unité de gestion des établissements de paie.
Mots clefs ... :
*****************************************************************}
unit uPaieEtabCompl ;

interface

uses CLasses,
     SysUtils,
{$ifndef eaglclient}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$endif}
     hctrls,
     uTob ;

     function initTOB_EtabCompl () : TOB ;
     function TOB_Etablissement () : TOB ;
     procedure Nettoyage_Etablissement() ;
     
implementation

var
   LocaleTob: TOB ;
   DateModif: TDateTime ;

////////////////////////////////////////////////////////////////////////////////

function initTOB_EtabCompl() : TOB ;
var q: tquery ;
    t: tob ;
    tfind: tob ;
    st : String;
begin

     t         := nil ;
     // Si premier appel
     if not assigned(LocaleTob) then
        begin
        LocaleTob := Tob.Create('etablissements_de_paie',Nil,-1) ;
        DateModif := Now ;

{Flux optimisé
        q := opensql('SELECT * FROM ETABCOMPL ORDER BY ETB_ETABLISSEMENT',True) ;
        if not q.eof then LocaleTob.LoadDetailDb('ETABCOMPL','','',q,False) ;
        Ferme(q) ;
}
          st := 'SELECT * FROM ETABCOMPL ORDER BY ETB_ETABLISSEMENT';
          LocaleTob.LoadDetailDbFromSQL('ETABCOMPL',st) ;
        end
     else
        begin
        q := opensql('SELECT * FROM ETABCOMPL WHERE ETB_DATEMODIF>="'+UsTime(DateModif)+'" ORDER BY ETB_ETABLISSEMENT',True) ;
        DateModif := Now ;
        if not q.eof then
           begin
           t := tob.create('',nil,-1) ;
           T.LoadDetailDB('ETABCOMPL','','',q,false) ;

           // des variables ont été modififées ??
           while t.detail.count>0 do
                 begin
                 tfind := localetob.FindFirst(['ETB_ETABLISSEMENT'],[t.detail[0].getvalue('ETB_ETABLISSEMENT')],false);
                 if assigned(tfind) then FreeAndNil(tfind) ;
                 t.Detail[0].ChangeParent(LocaleTob,-1) ;
                 end ;

           // Tri !!
           LocaleTob.Detail.Sort('ETB_ETABLISSEMENT');
           FreeAndNil(t) ;
           end ;
        Ferme(q) ;
        end ;
     Result := LocaleTob ;
end ;

function TOB_Etablissement () : TOB ;
begin
     Result := LocaleTob ;
end ;

////////////////////////////////////////////////////////////////////////////////

procedure Init() ;
begin
   LocaleTob:=Nil ;
end ;

procedure Nettoyage_Etablissement() ;
begin
     if Assigned(LocaleTob) then FreeAndNil(LocaleTob) ;
end ;

////////////////////////////////////////////////////////////////////////////////

initialization
  init() ;

finalization
  Nettoyage_Etablissement() ;
end.
