{***********UNITE*************************************************
Auteur  ...... : Joël SIDOS
Créé le ...... : 16/02/2006
Modifié le ... :   /  /    
Description .. : fournit la liste des users ayant acces à un tag
               : de menu (paramètres : Tag:integer, TobResultat : Tob)
Mots clefs ... : TOF;CONF_VUE;CONFIDENTIALITE;GROUPE;UTILISATEUR;MENU
*****************************************************************}

unit ListUsersAutorises;

interface

Uses StdCtrls,
     Controls,Dialogs,
     Classes,Windows,
     Sysutils,
{$IFDEF EAGLCLIENT}
     MaineAGL, UtileAgl,
{$ELSE}
     db,
  {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
     FE_main,
{$ENDIF}
     HCtrls,
     UTOB;

procedure ListUsers(mytag:integer ; myTOBResult : TOB);
procedure ListUsersVignettes(mytag:String ; myTOBResult : TOB);

implementation
//js1:130206: fonction de recherche des users ayant acces a la fonctionnalite
// passée en parametre (mytag) ==> renvoie la tob chargée avec les users/groupes
procedure ListUsers(mytag:integer ; myTOBResult : TOB);
var iCol,iInd:integer;
    TobMonRow, TobMesUser,TobMesUserGrp,myTOBResultFilles:TOB;
    QMn:TQuery;
begin

  TobMonRow:=TOB.Create('',nil,-1);
  TobMesUser:=TOB.Create('',nil,-1);
  TobMesUserGrp:=TOB.Create('',nil,-1);
//on retrouve le menu correspondant au tag
  QMn := OpenSQL('SELECT * FROM MENU WHERE MN_TAG = '+ IntToStr(mytag),True,-1, '', True);
  try
    if not QMn.Eof then TobMonRow.LoadDetailDB('MENU','','',QMn,true);
  finally
    Ferme(QMn);
  end;

  TobMesUserGrp.LoadDetailFromSQL('SELECT UG_GROUPE,UG_NUMERO,UG_LIBELLE FROM USERGRP');
  TobMesUser.LoadDetailFromSQL('SELECT US_UTILISATEUR,US_LIBELLE,UG_LIBELLE,UG_GROUPE,US_GRPSDELEGUES FROM UTILISAT ' +
           'LEFT JOIN USERGRP ON US_GROUPE=UG_GROUPE ORDER BY US_LIBELLE,UG_LIBELLE');

//on boucle sur tous les groupes
  for iCol := 0 to TobMesUserGrp.Detail.Count-1 do
  begin
//on teste si le groupe a acces au menu
    if (copy(TobMonRow.Detail[0].GetString('MN_ACCESGRP'),TobMesUserGrp.Detail[iCol].GetInteger('UG_NUMERO'),1) = '0')
    then begin
//on va chercher tous les utilisateurs du groupe et on charge la tob
      for iInd := 0 to TobMesUser.Detail.Count - 1 do
      begin
        if TobMesUser.Detail[iInd].GetString('UG_GROUPE') =  TobMesUserGrp.Detail[iCol].GetString('UG_GROUPE')
          then begin
                myTOBResultFilles:=TOB.Create('',myTOBResult,-1);
                myTOBResultFilles.AddChampSupValeur('US_LIBELLE',TobMesUser.Detail[iInd].GetString('US_LIBELLE'));
                myTOBResultFilles.AddChampSupValeur('UG_LIBELLE',TobMesUser.Detail[iInd].GetString('UG_LIBELLE'));
          end;
      end;

    end;
  end;

  TobMesUserGrp.free;
  TobMesUser.free;
  TobMonRow.free;

end;
procedure ListUsersVignettes(mytag:String ; myTOBResult : TOB);
var iCol,iInd:integer;
    TobMonRow, TobMesUser,TobMesUserGrp,myTOBResultFilles:TOB;
    QMn:TQuery;
begin

  TobMonRow:=TOB.Create('',nil,-1);
  TobMesUser:=TOB.Create('',nil,-1);
  TobMesUserGrp:=TOB.Create('',nil,-1);
//on retrouve le menu correspondant au tag

  QMn := OpenSQL('SELECT * FROM SWVIGNETTES WHERE SWV_TYPEVIGNETTE = "'+ ReadTokenSt(mytag)
                + '" AND SWV_CODEVIGNETTE = "' + ReadTokenSt(mytag) + '"',True,-1, '', True);
  try
    if not QMn.Eof then TobMonRow.LoadDetailDB('SWVIGNETTES','','',QMn,true);
  finally
    Ferme(QMn);
  end;

  TobMesUserGrp.LoadDetailFromSQL('SELECT UG_GROUPE,UG_NUMERO,UG_LIBELLE FROM USERGRP');
  TobMesUser.LoadDetailFromSQL('SELECT US_UTILISATEUR,US_LIBELLE,UG_LIBELLE,UG_GROUPE,US_GRPSDELEGUES FROM UTILISAT ' +
           'LEFT JOIN USERGRP ON US_GROUPE=UG_GROUPE ORDER BY US_LIBELLE,UG_LIBELLE');

//on boucle sur tous les groupes
  for iCol := 0 to TobMesUserGrp.Detail.Count-1 do
  begin
//on teste si le groupe a acces au menu
    if (copy(TobMonRow.Detail[0].GetString('SWV_ACCESGRP'),TobMesUserGrp.Detail[iCol].GetInteger('UG_NUMERO'),1) = '0')
    then begin
//on va chercher tous les utilisateurs du groupe et on charge la tob
      for iInd := 0 to TobMesUser.Detail.Count - 1 do
      begin
        if TobMesUser.Detail[iInd].GetString('UG_GROUPE') =  TobMesUserGrp.Detail[iCol].GetString('UG_GROUPE')
          then begin
                myTOBResultFilles:=TOB.Create('',myTOBResult,-1);
                myTOBResultFilles.AddChampSupValeur('US_LIBELLE',TobMesUser.Detail[iInd].GetString('US_LIBELLE'));
                myTOBResultFilles.AddChampSupValeur('UG_LIBELLE',TobMesUser.Detail[iInd].GetString('UG_LIBELLE'));
          end;
      end;

    end;
  end;

  TobMesUserGrp.free;
  TobMesUser.free;
  TobMonRow.free;

end;
end.
 