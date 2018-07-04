unit UtilPaieAffaire;

interface

Uses
{$IFDEF EAGLCLIENT}
      UTOB,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
      HCtrls,SysUtils,HEnt1;

Function RechercheSalarieRessource (Code: String; TypeSalarie : boolean): String;
Function NombreJoursOuvresOuvrablesMois(datedebut,datefin : tdatetime;Ouvres :boolean): integer;
Function  SupTablesLiees (NomTable, NomChamp, Valeur, CompWhere : String ; Sup : Boolean) : Boolean;

implementation

Function  SupTablesLiees (NomTable, NomChamp, Valeur, CompWhere : String ; Sup : Boolean) : Boolean;
Begin
Result := True;
If Sup Then ExecuteSQL('DELETE FROM '+ NomTable +' WHERE '+ NomChamp + ' ="' + Valeur + '" '+ CompWhere)
       Else Result := ExisteSQL('SELECT '+ NomChamp + ' FROM '+ NomTable +' WHERE '+ NomChamp + ' ="' + Valeur + '" '+ CompWhere);
End;


{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 02/03/2000
Modifié le ... : 02/03/2000
Description .. : Retourne le code salarié ou ressource du code inverse (ressource ou salarié) passé
Mots clefs ... : RESSOURCE,SALARIE
*****************************************************************}
Function RechercheSalarieRessource (Code: String; TypeSalarie : boolean): String;
Var Q : TQuery;
    NomChamp,ChampRetourne : String;
Begin
Result:='';
If TypeSalarie Then Begin NomChamp := 'ARS_SALARIE'; ChampRetourne := 'ARS_RESSOURCE'; End
Else Begin NomChamp := 'ARS_RESSOURCE'; ChampRetourne :='ARS_SALARIE' End;
Q:=OPENSQL('SELECT '+ ChampRetourne +' From RESSOURCE WHERE '+ NomChamp +'="'+Code+'"',True);
If Not Q.EOF Then result:=Q.Fields[0].AsString;
Ferme(Q);

End;


{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 02/03/2000
Modifié le ... : 02/03/2000
Description .. : Retourne le nombre de jours ouvrés / ouvrable de la période concernée
Mots clefs ... : DATE,OUVRABLE
*****************************************************************}
Function NombreJoursOuvresOuvrablesMois(datedebut,datefin : tdatetime;Ouvres :boolean): integer;
var
{Jour_fin,}Jour_j:TDatetime;
Numero : integer;
Nb_Lu,Nb_Ma,Nb_Me,Nb_Je,Nb_Ve,Nb_Sa{,Nb_Di}: integer;
begin
     Nb_Lu := 0;Nb_Ma := 0;Nb_Me := 0;Nb_Je := 0;
     Nb_Ve := 0;Nb_Sa := 0;//Nb_Di := 0;
     Jour_j  := datedebut;

     while Jour_j <= datefin do
        begin
        Numero  := DayOfWeek(Jour_j);
        Case Numero of
             1 : Nb_Lu := Nb_Lu+1;
             2 : Nb_Ma := Nb_Ma+1;
             3 : Nb_Me := Nb_Me+1;
             4 : Nb_Je := Nb_Je+1;
             5 : Nb_Ve := Nb_Ve+1;
             6 : Nb_Sa := Nb_Sa+1;
             //7 : Nb_Di := Nb_Di+1;
         end;
         Jour_j := Jour_j +1 ;
     end;
     if Ouvres then result := (Nb_Lu+Nb_Ma+Nb_Me+Nb_Je+Nb_Ve+Nb_Sa)
          else result := (Nb_Lu+Nb_Ma+Nb_Me+Nb_Je+Nb_Ve);

end;

end.
