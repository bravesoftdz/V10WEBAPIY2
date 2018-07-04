{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. : Unité de gestion des éléments nationaux
Mots clefs ... :
*****************************************************************
PT1   : 03/09/2007 FC V_80 FQ 14721 Prise en compte de la convention collective pour le prédéfini CEG
}
unit uPaieEltNatCEG;

interface

uses CLasses,
  SysUtils,
{$IFNDEF eaglclient}
  db,
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
  hctrls,
  uTob,paramsoc;

function initTOB_EltNatCEG(): TOB;
function TOB_EltNationauxCEG(): TOB;
procedure Nettoyage_EltNationauxCEG();

implementation

var
  LocaleTob: TOB;
  DateModif: TDateTime;

////////////////////////////////////////////////////////////////////////////////

function initTOB_EltNatCEG() : TOB ;
var q,Qry: tquery ;
    t: tob ;
    tfind: tob ;
    ListeConv,Requete,st : String;
begin
    //DEB PT1 Récupérer toutes les conventions collectives utilisées dans le dossier en cours
    Requete:= 'SELECT ET_ETABLISSEMENT, ET_NODOSSIER, ETB_CONVENTION,'+
              ' ETB_CONVENTION1, ETB_CONVENTION2'+
              ' FROM ETABLISS'+
              ' LEFT JOIN ETABCOMPL ON'+
              ' ET_ETABLISSEMENT=ETB_ETABLISSEMENT WHERE'+
              ' (ET_NODOSSIER="000000" OR'+
              ' ET_NODOSSIER="'+GetParamSocSecur ('SO_NODOSSIER', '')+'")';

    Qry := opensql(Requete, true);
    while not Qry.EOF do
          begin
          if (Qry.findfield ('ETB_CONVENTION').asstring<>'') and
             (pos (Qry.findfield ('ETB_CONVENTION').asstring,ListeConv)=0) then
             begin
             if (ListeConv <> '') then
                ListeConv:= ListeConv+',"'+Qry.findfield ('ETB_CONVENTION').asstring +'"'
             else
                ListeConv:= '"'+Qry.findfield ('ETB_CONVENTION').asstring +'"';
             end;
          if (Qry.findfield ('ETB_CONVENTION1').asstring<>'') and
             (pos (Qry.findfield ('ETB_CONVENTION1').asstring,ListeConv)=0) then
             begin
             if (ListeConv <> '') then
                ListeConv:= ListeConv+',"'+Qry.findfield ('ETB_CONVENTION1').asstring+'"'
             else
                ListeConv:= '"'+Qry.findfield ('ETB_CONVENTION1').asstring+'"';
             end;
          if (Qry.findfield ('ETB_CONVENTION2').asstring<>'') and
             (pos (Qry.findfield ('ETB_CONVENTION2').asstring,ListeConv)=0) then
             begin
             if (ListeConv <> '') then
                ListeConv:= ListeConv+',"'+Qry.findfield ('ETB_CONVENTION2').asstring+'"'
             else
                ListeConv:= '"'+Qry.findfield ('ETB_CONVENTION2').asstring+'"';
             end;
          Qry.next;
          end;
    ferme(Qry);
    //FIN PT1

    if (ListeConv <> '') then
      ListeConv := ',' + ListeConv;

     t         := nil ;
     // Si premier appel
     if not assigned(LocaleTob) then
        begin
        LocaleTob := Tob.Create('EltNat_CEG',Nil,-1) ;
        DateModif := Now ;
{Flux optimisé
        q := opensql('SELECT * FROM ELTNATIONAUX WHERE PEL_PREDEFINI = "CEG" '
        + ' AND PEL_CONVENTION IN ("000"' + ListeConv + ') ' //PT1
        + ' ORDER BY PEL_CODEELT,PEL_DATEVALIDITE',True) ;
        if not q.eof then LocaleTob.LoadDetailDb('ELTNATIONAUX','','',q,False) ;
        Ferme(q) ;
}
        st := 'SELECT * FROM ELTNATIONAUX WHERE PEL_PREDEFINI = "CEG" '
        + ' AND PEL_CONVENTION IN ("000"' + ListeConv + ') ' //PT1
        + ' ORDER BY PEL_CODEELT,PEL_DATEVALIDITE';
        LocaleTob.LoadDetailDbFromSQL('ELTNATIONAUX',st) ;
        end
     else
        begin
        q := opensql('SELECT * FROM ELTNATIONAUX WHERE PEL_PREDEFINI = "CEG" '
        + ' AND PEL_CONVENTION IN ("000"' + ListeConv + ') ' //PT1
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
           LocaleTob.Detail.Sort('PEL_CONVENTION;PEL_CODEELT;PEL_DATEVALIDITE'); //PT1
           FreeAndNil(t) ;
           end ;
        Ferme(q) ;
        end ;

     Result := LocaleTob ;
end ;

function TOB_EltNationauxCEG(): TOB;
begin
  Result := LocaleTob;
end;

////////////////////////////////////////////////////////////////////////////////

procedure Init();
begin
  LocaleTob := nil;
end;

procedure Nettoyage_EltNationauxCEG();
begin
  if Assigned(LocaleTob) then FreeAndNil(LocaleTob);
end;

////////////////////////////////////////////////////////////////////////////////

initialization
  init();

finalization
  Nettoyage_EltNationauxCEG();
end.

