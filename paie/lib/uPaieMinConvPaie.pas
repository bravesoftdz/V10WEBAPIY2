{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. : Unité de gestion des tables détail minimum conventionnel de paie.
Mots clefs ... :
*****************************************************************}
unit uPaieMinConvPaie ;

interface

uses CLasses,
     SysUtils,
{$ifndef eaglclient}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$endif}
     hctrls,
     uTob ;

     function initTOB_MinConvPaie () : TOB ;
     function TOB_DetailMin () : TOB ;
     procedure Nettoyage_DetailMin() ;
     
implementation

var
   LocaleTob: TOB ;
   DateModif: TDateTime ;

////////////////////////////////////////////////////////////////////////////////


function initTOB_MinConvPaie() : TOB ;
var q: tquery ;
    t: tob ;
    tfind: tob ;
    st : string;
begin

     t         := nil ;
     // Si premier appel
     if not assigned(LocaleTob) then
        begin
        LocaleTob := Tob.Create('MinimumConv_de_paie',Nil,-1) ;
        DateModif := Now ;
{Flux optimisé
        q := opensql('SELECT * FROM MINCONVPAIE WHERE ##PCP_PREDEFINI## ORDER BY PCP_PREDEFINI,PCP_NODOSSIER,PCP_NATURE,PCP_CODE,PCP_CONVENTION,PCP_NBRE',True) ;
        if not q.eof then LocaleTob.LoadDetailDb('MINCONVPAIE','','',q,False) ;
        Ferme(q) ;
}
        st := 'SELECT * FROM MINCONVPAIE WHERE ##PCP_PREDEFINI## ORDER BY PCP_PREDEFINI,PCP_NODOSSIER,PCP_NATURE,PCP_CODE,PCP_CONVENTION,PCP_NBRE';
        LocaleTob.LoadDetailDbFromSQL('MINCONVPAIE',st) ;
        end
     else
        begin
        q := opensql('SELECT * FROM MINCONVPAIE WHERE ##PCP_PREDEFINI## AND PCP_DATEMODIF>="'+UsTime(DateModif)+'" ORDER BY PCP_PREDEFINI,PCP_NODOSSIER,PCP_NATURE,PCP_CODE,PCP_CONVENTION,PCP_NBRE',True) ;
        DateModif := Now ;
        if not q.eof then
           begin
           t := tob.create('',nil,-1) ;
           T.LoadDetailDB('MINCONVPAIE','','',q,false) ;

           // des variables ont été modififées ??
           while t.detail.count>0 do
                 begin
                 tfind := localetob.FindFirst(['PCP_PREDEFINI','PCP_NODOSSIER','PCP_NATURE','PCP_CODE','PCP_CONVENTION','PCP_NBRE'],
                   [t.detail[0].getvalue('PCP_PREDEFINI'),t.detail[0].getvalue('PCP_NODOSSIER'),t.detail[0].getvalue('PCP_NATURE'),
                    t.detail[0].getvalue('PCP_CODE'),t.detail[0].getvalue('PCP_CONVENTION'),t.detail[0].getvalue('PCP_NBRE')],false);
                 if assigned(tfind) then FreeAndNil(tfind) ;
                 t.Detail[0].ChangeParent(LocaleTob,-1) ;
                 end ;

           // Tri !!
           LocaleTob.Detail.Sort('PCP_PREDEFINI;PCP_NODOSSIER;PCP_NATURE;PCP_CODE;PCP_CONVENTION;PCP_NBRE');
           FreeAndNil(t) ;
           end ;
        Ferme(q) ;
        end ;
     Result := LocaleTob ;
end ;

function TOB_DetailMin () : TOB ;
begin
     Result := LocaleTob ;
end ;

////////////////////////////////////////////////////////////////////////////////

procedure Init() ;
begin
   LocaleTob:=Nil ;
end ;

procedure Nettoyage_DetailMin() ;
begin
     if Assigned(LocaleTob) then FreeAndNil(LocaleTob) ;
end ;

////////////////////////////////////////////////////////////////////////////////

initialization
  init() ;

finalization
  Nettoyage_DetailMin() ;
end.
