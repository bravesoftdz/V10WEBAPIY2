{***********UNITE*************************************************
Auteur  ...... : D.T.
Créé le ...... : 22/02/2003
Modifié le ... :   /  /
Description .. : Unité de gestion des ventilations des Cotisations de paie.
Mots clefs ... :
*****************************************************************}
unit uPaieVentiCot;

interface

uses CLasses,
     SysUtils,
{$ifndef eaglclient}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$endif}
     hctrls,
     uTob ;

     function initTOB_VentiCotPaie () : TOB ;
     function TOB_VentilCot () : TOB ;
     procedure Nettoyage_VentilCot() ;

implementation

var
   LocaleTob: TOB ;
   DateModif: TDateTime ;

////////////////////////////////////////////////////////////////////////////////

function initTOB_VentiCotPaie () : TOB ;
var q: tquery ;
    t: tob ;
    tfind: tob ;
    st : string;
begin

     t         := nil ;
     // Si premier appel
     if not assigned (LocaleTob) then
        begin
        LocaleTob := Tob.Create ('VentilCot_de_la_paie',Nil,-1) ;
        DateModif := Now ;
{Flux optimisé
        q := opensql ('SELECT * from VENTICOTPAIE WHERE ##PVT_PREDEFINI## ORDER BY PVT_RUBRIQUE',True) ;
        if not q.eof then LocaleTob.LoadDetailDb ('VENTICOTPAIE','','',q,False) ;
        Ferme(q) ;
}
        st := 'SELECT * from VENTICOTPAIE WHERE ##PVT_PREDEFINI## ORDER BY PVT_RUBRIQUE' ;
        LocaleTob.LoadDetailDbFromSQL ('VENTICOTPAIE',st) ;
        end
     else
        begin
        q := opensql('select * from VENTICOTPAIE where ##PVT_PREDEFINI## AND PVT_DATEMODIF>="'+UsTime(DateModif)+
            '" ORDER BY PVT_RUBRIQUE',True) ;
        DateModif := Now ;
        if not q.eof then
           begin
           t := tob.create('',nil,-1) ;
           T.LoadDetailDB ('VENTICOTPAIE','','',q,false) ;

           // des variables ont été modififées ??
           while t.detail.count>0 do
                    begin
                    tfind := localetob.FindFirst(['PVT_RUBRIQUE'],[t.detail[0].getvalue('PVT_RUBRIQUE')],false);
                    if assigned(tfind) then FreeAndNil(tfind) ;
                    t.Detail[0].ChangeParent(LocaleTob,-1) ;
                    end ;

           // Tri !!
           LocaleTob.Detail.Sort('PVT_RUBRIQUE');

           FreeAndNil(t) ;
           end ;
        Ferme(q) ;
        end ;

     Result := LocaleTob ;
end ;

function TOB_VentilCot () : TOB ;
begin
     Result := LocaleTob ;
end ;

////////////////////////////////////////////////////////////////////////////////

procedure Init() ;
begin
   LocaleTob:=Nil ;
end ;

procedure Nettoyage_VentilCot() ;
begin
     if Assigned(LocaleTob) then FreeAndNil(LocaleTob) ;
end ;

////////////////////////////////////////////////////////////////////////////////

initialization
  init() ;

finalization
  Nettoyage_VentilCot () ;

end.
