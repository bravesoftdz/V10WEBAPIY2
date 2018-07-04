{***********UNITE*************************************************
Auteur  ...... : D.T.
Créé le ...... : 22/02/2003
Modifié le ... :   /  /
Description .. : Unité de gestion des rubriques de cotisations de type bases.
Mots clefs ... :
*****************************************************************}
unit uPaieBases ;

interface

uses CLasses,
     SysUtils,
{$ifndef eaglclient}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$endif}
     hctrls,
     uTob ;

     function initTOB_Bases() : TOB ;
     function TOB_Bases() : TOB ;
     procedure Nettoyage_Bases() ;
     
implementation

var
   LocaleTob: TOB ;
   DateModif: TDateTime ;

////////////////////////////////////////////////////////////////////////////////

function initTOB_Bases() : TOB ;
var q: tquery ;
    t: tob ;
    tfind: tob ;
    st : String ;
begin

     t         := nil ;
     // Si premier appel
     if not assigned(LocaleTob) then
        begin
        LocaleTob := Tob.Create('rubrique_de_base',Nil,-1) ;
        DateModif := Now ;
{Flux optimisé
        q := opensql('SELECT * FROM COTISATION WHERE ##PCT_PREDEFINI## PCT_NATURERUB="BAS" ORDER BY PCT_RUBRIQUE',True) ;
        if not q.eof then LocaleTob.LoadDetailDb('COTISATION','','',q,False) ;
        Ferme(q) ;
}
        st :=' SELECT * FROM COTISATION WHERE ##PCT_PREDEFINI## PCT_NATURERUB="BAS" ORDER BY PCT_RUBRIQUE';
        LocaleTob.LoadDetailDbFromSQL('COTISATION',st);
        end
     else
        begin
        q := opensql('SELECT * FROM COTISATION WHERE ##PCT_PREDEFINI## PCT_NATURERUB="BAS" AND PCT_DATEMODIF>="'+UsTime(DateModif)+'" order by PCT_RUBRIQUE',True) ;
        DateModif := Now ;
        if not q.eof then
           begin
           t := tob.create('',nil,-1) ;
           T.LoadDetailDB('COTISATION','','',q,false) ;

           // des variables ont été modififées ??
           if assigned(t) then
              begin
              while t.detail.count>0 do
                    begin
                    tfind := localetob.FindFirst(['PCT_RUBRIQUE'],[t.detail[0].getvalue('PCT_RUBRIQUE')],false);
                    if assigned(tfind) then FreeAndNil(tfind) ;
                    t.Detail[0].ChangeParent(LocaleTob,-1) ;
                    end ;

              // Tri !!
              LocaleTob.Detail.Sort('PCT_RUBRIQUE');
              end ;
           FreeAndNil(t) ;
           end ;
        Ferme(q) ;
        end ;
     Result := LocaleTob ;
end ;

function TOB_Bases() : TOB ;
begin
     Result := LocaleTob ;
end ;

////////////////////////////////////////////////////////////////////////////////

procedure Init() ;
begin
   LocaleTob:=Nil ;
end ;

procedure Nettoyage_Bases() ;
begin
     if Assigned(LocaleTob) then FreeAndNil(LocaleTob) ;
end ;

////////////////////////////////////////////////////////////////////////////////

initialization
  init() ;

finalization
  Nettoyage_Bases() ;


end.
