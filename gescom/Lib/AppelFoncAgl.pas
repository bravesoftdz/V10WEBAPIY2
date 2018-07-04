{***********UNITE*************************************************
Auteur  ...... : LUQUET Jean-Marc
Créé le ...... : 09/08/2001
Modifié le ... : 27/09/2001
Description .. : Unité appelant les fonctions AGL pour les scripts d'une
Suite ........ : fiche.
Suite ........ : Pour le moment :
Suite ........ : - RTEPlusMois           Date Française
Suite ........ : - RTEDebutDeMois   Date Française
Suite ........ : - RTEFinDeMois        Date Française
Suite ........ : - RTEDebutAnnee       Date Française
Suite ........ : - RTEUSDateTime    Date US
Suite ........ :
Suite ........ : Exemple d'appel par le script :
Suite ........ : Ma_Date = RTEPlusMois(XX_DATEDEPART.text,3);
Suite ........ : Ma_Date = RTEDebutDeMois(XX_DATEDEPART.text);
Suite ........ : Ma_Date = RTEFinDeMois(XX_DATEDEPART.text);
Mots clefs ... : PLUSMOIS, DEBUTDEMOIS, FINDEMOIS, DEBUTANNEE, SCRIPTS, MOIS
*****************************************************************}
unit AppelFoncAgl;

interface

uses Sysutils, M3FP, Hent1,
     Hctrls;

implementation

Function AGLPlusMois ( parms: array of variant; nb: integer ): variant ;
var  Date_Deb : TDateTime;
begin
 // Date Française
 Date_deb := PlusMois ( StrToDateTime( parms[0] ), strtoInt ( parms[1] ) );
 result   := DateToStr ( Date_Deb );
end;

Function AGLDebutDeMois ( parms: array of variant; nb: integer ): variant ;
var  Date_Deb : TDateTime;
begin
 // Date Française
 Date_Deb := DebutDeMois ( StrToDateTime( parms[0] ) );
 result   := DateToStr ( Date_Deb );
end;

Function AGLFinDeMois ( parms: array of variant; nb: integer ): variant ;
var  Date_Deb : TDateTime;
begin
 // Date Française
 Date_Deb := FinDeMois ( StrToDateTime( parms[0] ) );
 result   := DateToStr ( Date_Deb );
end;

Function AGLDebutAnnee ( parms: array of variant; nb: integer ): variant ;
var  Date_Deb : TDateTime;
begin
 // Date Française
 Date_Deb := DebutAnnee ( StrToDateTime( parms[0] ) );
 result   := DateToStr ( Date_Deb );
end;

Function AGLUSDateTime ( parms: array of variant; nb: integer ): variant ;
var  Date_Deb : TDateTime;
begin
 // Date US pour requête SQL
 Date_Deb := StrToDateTime( parms[0] );
 result   := USDateTime ( Date_Deb );
end;

Initialization
RegisterAglFunc( 'RTEPlusMois'   , FALSE , 2, AGLPlusMois);
RegisterAglFunc( 'RTEDebutDeMois', FALSE , 2, AGLDebutDeMois);
RegisterAglFunc( 'RTEFinDeMois'  , FALSE , 2, AGLFinDeMois);
RegisterAglFunc( 'RTEDebutAnnee' , FALSE , 2, AGLDebutAnnee);
RegisterAglFunc( 'RTEUSDateTime' , FALSE , 2, AGLUSDateTime);
end.
