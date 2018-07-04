{***********UNITE*************************************************
Auteur  ...... : FC
Créé le ...... : 10/04/2007
Modifié le ... :   /  /
Description .. : Unité de gestion des éléments nationaux pour les conventions du dossier
Mots clefs ... :
*****************************************************************
}
unit uPaieEltNatSTDCNV ;

interface

uses CLasses,
     SysUtils,
{$ifndef eaglclient}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$endif}
     hctrls,
     uTob, paramsoc ;

     function initTOB_EltNatSTDCNV() : TOB ;
     function TOB_EltNationauxSTDCNV() : TOB ;
     procedure Nettoyage_EltNationauxSTDCNV() ;

implementation

var
   LocaleTob : TOB ;
   DateModif : TDateTime ;

////////////////////////////////////////////////////////////////////////////////

function initTOB_EltNatSTDCNV() : TOB ;
var q: tquery ;
    t: tob ;
    tfind: tob ;
  Qry : TQuery;
  Requete : String;
  ListeConv : String;
begin

    Requete := 'select ET_ETABLISSEMENT, ET_NODOSSIER, ETB_CONVENTION, ETB_CONVENTION1, ETB_CONVENTION2 from ETABLISS '
             + ' left join ETABCOMPL on ET_ETABLISSEMENT = ETB_ETABLISSEMENT '
             + 'where ET_NODOSSIER = "000000" or ET_NODOSSIER = "'+GetParamSocSecur('SO_NODOSSIER', '')+'"';
    Qry := opensql(Requete, true);
    while not Qry.EOF do
    begin
//      if pos(Qry.findfield('ET_ETABLISSEMENT').asstring,ListeEtab) = 0
//        then ListeEtab := ListeEtab + '"'+Qry.findfield('ET_ETABLISSEMENT').asstring+'", ';
      if pos(Qry.findfield('ETB_CONVENTION').asstring,ListeConv) = 0
        then ListeConv := ListeConv + '"'+Qry.findfield('ETB_CONVENTION').asstring +'", ';
      if pos(Qry.findfield('ETB_CONVENTION1').asstring,ListeConv) = 0
        then ListeConv := ListeConv + '"'+Qry.findfield('ETB_CONVENTION1').asstring+'", ';
      if pos(Qry.findfield('ETB_CONVENTION2').asstring,ListeConv) = 0
        then ListeConv := ListeConv + '"'+Qry.findfield('ETB_CONVENTION2').asstring+'", ';
      Qry.next;
    end;
    ferme(Qry);

     t         := nil ;
     // Si premier appel
     if not assigned(LocaleTob) then
        begin
        LocaleTob := Tob.Create('EltNat_STD',Nil,-1) ;
        q := opensql('SELECT * FROM ELTNATIONAUX WHERE PEL_PREDEFINI = "STD" '
        + ' AND PEL_CONVENTION IN (' + ListeConv + ') '
        + ' ORDER BY PEL_CODEELT,PEL_DATEVALIDITE',True) ;
        DateModif := Now ;                
        if not q.eof then LocaleTob.LoadDetailDb('ELTNATIONAUX','','',q,False) ;
        Ferme(q) ;
        end
     else
        begin
        q := opensql('SELECT * FROM ELTNATIONAUX WHERE PEL_PREDEFINI = "STD" '
        + ' AND PEL_CONVENTION IN (' + ListeConv + ') '
        + ' AND PEL_DATEMODIF>="'+UsTime(DateModif)+'"'
        + ' ORDER BY PEL_CODEELT,PEL_DATEVALIDITE',True) ;
        DateModif := Now ;
        if not q.eof then
           begin
           t := tob.create('',nil,-1) ;
           T.LoadDetailDB('ELTNATIONAUX','','',q,false) ;

           // des variables ont été modifiées ??
           while t.detail.count>0 do
                 begin
                 tfind := localetob.FindFirst(['PEL_CODEELT','PEL_DATEVALIDITE'],[t.detail[0].getvalue('PEL_CODEELT'),t.detail[0].getvalue('PEL_DATEVALIDITE')],false);
                 if assigned(tfind) then FreeAndNil(tfind) ;
                 t.Detail[0].ChangeParent(LocaleTob,-1) ;
                 end ;

           // Tri !!
           LocaleTob.Detail.Sort('PEL_CONVENTION;PEL_CODEELT;PEL_DATEVALIDITE');
           FreeAndNil(t) ;
           end ;
        Ferme(q) ;
        end ;

     Result := LocaleTob ;
end ;

function TOB_EltNationauxSTDCNV() : TOB ;
begin
     Result := LocaleTob ;
end ;

////////////////////////////////////////////////////////////////////////////////

procedure Init() ;
begin
   LocaleTob:=Nil ;
end ;

procedure Nettoyage_EltNationauxSTDCNV() ;
begin
     if Assigned(LocaleTob) then FreeAndNil(LocaleTob) ;
end ;

////////////////////////////////////////////////////////////////////////////////

initialization
  init() ;

finalization
  Nettoyage_EltNationauxSTDCNV() ;
end.
