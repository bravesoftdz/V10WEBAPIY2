
{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 11/09/2002
Modifié le ... :
Description .. :
Mots clefs ... : PAIE;HIERARCHIE
*****************************************************************}
{
PT1   05/12/2005 SB V65 FQ 12737 Ajout fonction clause hierarchie
PT2   28/02/2006 SB V70 FQ 12926 Optimisation + Refonte tri du planning par nom resp., puis nom sal.
PT3   02/04/2008 FL V80 Ajout des fonctions pour adaptation des services au partage
}


unit PGHierarchie;

interface

Uses  Sysutils,
{$IFDEF EAGLCLIENT}
  UtileAGL,
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
   HCtrls,UTOB;


Const SEP_CHAINE_SERVICE = '_';

function RenvoiChampResponsable(TypeResp:String):String; // renvoi libellé du champ responsable
Function EstResponsableHierarchique ( Salarie, TypeResp : String) : Boolean;   // Test si le salarié est un responsable
function RecupServiceRespHierarchie (Salarie,TypeResp : String;MultiNiveau:Boolean) : TOB; //Renvoi tous les services d'un responsable
function RecupSalariesHierarchie (Responsable,TypeResp : String;MultiNiveau:Boolean) : TOB;// Renvoi tous les salariés d'un responsabe
function RecupResponsablesHierarchie(Salarie,TypeResp:String;MultiNiveau:Boolean):TOB;//Renvoi tous les responsables d'un salarié
Function RecupServiceHierarchieDESC(Service:String):TOB;//Renvoi tous les services en dessous du service
Function RecupServiceHierarchieASC(Service:String):TOB;//Renvoi tous les services au dessus du service
Function RendClauseHierarchie(StValeur : string) : string;    {PT1}
Function GetDossierService (NomService : String) : String;    //PT3
Function GetDossierServiceNum (NumService : String) : String; //PT3
Function ServicesDesDossiers : String; //PT3

implementation

Uses StrUtils, PGOutils;


{***********A.G.L.***********************************************
Auteur  ...... : JL
Créé le ...... : 11/09/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE
*****************************************************************}

{
Type de responsable :
 S : Service
 A : absence
 N : Notes de frais
 V : Elements variables
 F : Formation
}
function RenvoiChampResponsable(TypeResp:String):String;
var Champ:String;
begin
If TypeResp='S' Then Champ:='PGS_RESPSERVICE';
If TypeResp='A' Then Champ:='PGS_RESPONSABS';
If TypeResp='N' Then Champ:='PGS_RESPONSNDF';
If TypeResp='V' Then Champ:='PGS_RESPONSVAR';
If TypeResp='F' Then Champ:='PGS_RESPONSFOR';
If Champ='' Then Champ:='PGS_RESPSERVICE'; // Par défaut on prend le responsable du service
Result:=Champ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : JL
Créé le ...... : 11/09/2002
Modifié le ... :   /  /
Description .. : Renvoi Vrai si le salarié est responsable du type choisi
Mots clefs ... : PAIE,Hiérarchie
*****************************************************************}
{La fonction renvoie true si le salarié est responsable du type concerné
}

Function  EstResponsableHierarchique ( Salarie, TypeResp : String) : Boolean;
var champ:String;
begin
Champ:=RenvoiChampResponsable(TypeResp);
If Champ<>'' Then
   begin
   If ExisteSQL('SELECT PGS_CODESERVICE FROM SERVICES WHERE '+Champ+'="'+Salarie+'"') Then Result:=True
   Else Result:=False;
   end
Else Result:=False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : JL
Créé le ...... : 11/09/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE,Hiérarchie
*****************************************************************}
{
La fonction renvoi une tob qui contient tous les services qui appartiennent au responsable
Si Multiniveau=True alors les services sous dépendances indirect sont également pris en compte.
}

function  RecupServiceRespHierarchie (Salarie,TypeResp : String;MultiNiveau:Boolean) : TOB;
var QServices:TQuery;
    TobServices:Tob;
    Champ,Requete:String;
begin
Champ:=RenvoiChampResponsable(TypeResp);
If MultiNiveau=True Then
   begin
   Requete:='SELECT PGS_CODESERVICE,PGS_NOMSERVICE,PGS_CODESERVICE,PGS_SECRETAIRE,PGS_SERVICESUP,'+
   'PGS_ETABLISSEMENT,PGS_RESPONSABS,PGS_RESPONSNDF,PGS_RESPONSVAR,PGS_RESPONSFOR,PGS_SECRETAIREABS'+
   ',PGS_HIERARCHIE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
   ' WHERE ('+Champ+'="'+Salarie+'" AND (PSO_NIVEAUSUP=0 OR PSO_NIVEAUSUP=1))'+
   ' OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE '+Champ+'="'+Salarie+'")';
   end
Else
   begin
   Requete:='SELECT PGS_CODESERVICE,PGS_NOMSERVICE,PGS_CODESERVICE,PGS_SECRETAIRE,PGS_SERVICESUP,'+
   'PGS_ETABLISSEMENT,PGS_RESPONSABS,PGS_RESPONSNDF,PGS_RESPONSVAR,PGS_RESPONSFOR,PGS_SECRETAIREABS'+
   ',PGS_HIERARCHIE FROM SERVICES WHERE '+Champ+'="'+Salarie+'"';
   end;
QServices:=OpenSQL(Requete,True);
TobServices:=Tob.Create('Les services',Nil,-1);
TobServices.LoadDetailDB('SERVICES','','',QServices,False);
Ferme(QServices);
Result:=TobServices;
end;


{***********A.G.L.***********************************************
Auteur  ...... : JL
Créé le ...... : 11/09/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE,Hiérarchie
*****************************************************************}
{
La fonction renvoi tous les codes salariés reliés à un responsable
Si multiniveau=true alors tous les salariés qui dépendent indirectement du responsable sont pris en compte.
}

function  RecupSalariesHierarchie (Responsable,TypeResp : String;MultiNiveau:Boolean) : TOB;
var TobSalaries:Tob;
    QSalaries:TQuery;
    Requete,Champ:String;
begin
Champ:=RenvoiChampResponsable(TypeResp);
If MultiNiveau=True Then
   begin
   Requete:='SELECT PSE_SALARIE,PSE_CODESERVICE FROM DEPORTSAL WHERE PSE_CODESERVICE IN '+
            '(SELECT PGS_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
            ' WHERE ('+Champ+'="'+Responsable+'" AND (PSO_NIVEAUSUP=0 OR PSO_NIVEAUSUP=1))'+
            ' OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE '+Champ+'="'+Responsable+'"))';
   end
Else
   begin
   Requete:='SELECT PSE_SALARIE,PSE_CODESERVICE FROM DEPORTSAL WHERE PSE_CODESERVICE IN '+
   '(SELECT PGS_CODESERVICE FROM SERVICES WHERE '+Champ+'="'+Responsable+'")';
   end;
QSalaries:=OpenSQL(Requete,True);
tobSalaries:=Tob.Create('Les salaries',Nil,-1);
TobSalaries.LoadDetailDB('DEPORTSAL','','',QSalaries,False);
Ferme(QSalaries);
Result:=TobSalaries;
end;
{***********A.G.L.***********************************************
Auteur  ...... : JL
Créé le ...... : 11/09/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE,Hiérarchie
*****************************************************************}
{
La fonction récupère le responsable d'un salarié
Si Multiniveau = true alors la tob contient tous les responsable hiérarchiques.
}

function  RecupResponsablesHierarchie(Salarie,TypeResp:String;MultiNiveau:Boolean):TOB;
var QResponsables,QService:TQuery;
    TobResponsables:Tob;
    Champ,Service,Requete:String;
begin
Champ:=RenvoiChampResponsable(TypeResp);
QService:=OpenSQL('SELECT PSE_CODESERVICE FROM DEPORTSAL WHERE PSE_SALARIE="'+Salarie+'"',True);
Service:=QService.FindField('PSE_CODESERVICE').AsString;
Ferme(QService);
If MultiNiveau=True Then
   begin
   Requete:='SELECT PGS_RESPONSABS,PGS_RESPSERVICE,PGS_RESPONSVAR,PGS_RESPONSNDF,PGS_RESPONSFOR'+
            ' FROM SERVICES WHERE (PGS_CODESERVICE="'+Service+'") OR '+
            '(PGS_CODESERVICE IN (SELECT PSO_SERVICESUP FROM SERVICEORDRE WHERE PSO_CODESERVICE="'+Service+'"'+
            ' AND PSO_NIVEAUSUP>0))';
   end
Else
   begin
   Requete:='SELECT PGS_RESPONSABS,PGS_RESPSERVICE,PGS_RESPONSVAR,PGS_RESPONSNDF,PGS_RESPONSFOR'+
   ' FROM SERVICES WHERE PGS_CODESERVICE="'+Service+'"';
   end;
QResponsables:=OpenSQL(Requete,True);
TobResponsables:=Tob.Create('Les responsables',Nil,-1);
TobResponsables.LoaddetailDB('SERVICES','','',QResponsables,False);
Ferme(QResponsables);
Result:=TobResponsables;
end;

{***********A.G.L.***********************************************
Auteur  ...... : JL
Créé le ...... : 12/09/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE,Hiérarchie
*****************************************************************}

Function RecupServiceHierarchieDESC(Service:String):TOB;
var QServices:TQUery;
    TobServices:Tob;
begin
QServices:=OpenSQL('SELECT PGS_CODESERVICE,PGS_NOMSERVICE,PGS_CODESERVICE,PGS_SECRETAIRE,PGS_SERVICESUP,'+
                  'PGS_ETABLISSEMENT,PGS_RESPONSABS,PGS_RESPONSNDF,PGS_RESPONSVAR,PGS_RESPONSFOR,PGS_SECRETAIREABS'+
                  ',PGS_HIERARCHIE FROM SERVICES WHERE PGS_CODESERVICE IN (SELECT PSO_CODESERVICE FROM SERVICEORDRE WHERE PSO_SERVICESUP="'+Service+'")',True);
TobServices:=Tob.Create('Les services DESC',Nil,-1);
TobServices.LoadDetailDB('SERVICES','','',QServices,False);
Ferme(QServices);
Result:=TobServices;
end;


{***********A.G.L.***********************************************
Auteur  ...... : JL
Créé le ...... : 12/09/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE,Hiérarchie
*****************************************************************}
Function RecupServiceHierarchieASC(Service:String):TOB;
var QServices:TQUery;
    TobServices:Tob;
begin
QServices:=OpenSQL('SELECT PGS_CODESERVICE,PGS_NOMSERVICE,PGS_CODESERVICE,PGS_SECRETAIRE,PGS_SERVICESUP,'+
                  'PGS_ETABLISSEMENT,PGS_RESPONSABS,PGS_RESPONSNDF,PGS_RESPONSVAR,PGS_RESPONSFOR,PGS_SECRETAIREABS'+
                  ',PGS_HIERARCHIE FROM SERVICES WHERE PGS_CODESERVICE IN (SELECT PSO_SERVICESUP FROM SERVICEORDRE WHERE PSO_CODESERVICE="'+Service+'" AND PSO_NIVEAUSUP>0)',True);
TobServices:=Tob.Create('Les services DESC',Nil,-1);
TobServices.LoadDetailDB('SERVICES','','',QServices,False);
Ferme(QServices);
Result:=TobServices;
end;

{ DEB PT1 }
Function RendClauseHierarchie(StValeur : string) : string;
var St,StSql : string;
    Q        : Tquery;
Begin
result := '';
If Trim(StValeur)= '' then exit;
if Pos('<<',StValeur) > 0 then exit;
{ DEB PT2 Refonte pour optimisation
 Si selection d'un niveau hirarchique, recherche des service attaché }
StSql := 'SELECT DISTINCT PGS_CODESERVICE FROM SERVICES WHERE ';
While (StValeur<>'') do
  Begin
  St := ReadTokenSt(StValeur);
  StSql := StSql + 'PGS_HIERARCHIE="'+St+'" OR '
  End;
StSql := Copy(StSql,1 ,length(StSql)-4);
Q := OpenSql(StSql,True);
if not Q.eof then Result := '';
While Not Q.Eof do
  Begin
  Result := Result + '"'+Q.FindField('PGS_CODESERVICE').AsString+'",';
  Q.Next;
  End;
Ferme(Q);
If Result<>'' then Result := 'AND PSE_CODESERVICE IN ('+Copy(Result,1,Length(Result)-1)+') ';
{ FIN PT2 }
End;
{ FIN PT1 }

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 02/04/2008 / PT3
Modifié le ... :   /  /
Description .. : Retourne le n° de dossier d'un service (libellé)
Mots clefs ... :
*****************************************************************}
Function GetDossierService (NomService : String) : String;
Var PosSep : Integer;
begin
	Result := '';
	
	PosSep := Pos (SEP_CHAINE_SERVICE, NomService); 
	// Eclatement du n° dossier et nom
	If (PosSep > 0) Then Result := LeftStr(NomService, PosSep-1);
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 02/04/2008 / PT3
Modifié le ... :   /  /
Description .. : Retourne le n° de dossier d'un service (code)
Mots clefs ... :
*****************************************************************}
Function GetDossierServiceNum (NumService : String) : String;
Var Q      : TQuery;
    Nom    : String;
begin
	Q := OpenSQL('SELECT PGS_NOMSERVICE FROM SERVICES WHERE PGS_CODESERVICE="'+NumService+'"', True);
	If Not Q.EOF Then Nom := Q.FindField('PGS_NOMSERVICE').AsString;
	Ferme(Q);
	
	Result := GetDossierService (Nom);
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 02/04/2008 / PT3
Modifié le ... :   /  /
Description .. : Retourne la liste des services des dossiers accessibles suivant les groupes de travail
Mots clefs ... :
*****************************************************************}
Function ServicesDesDossiers : String;
var Q      : TQuery;
    Chaine : String;
Begin
	Chaine := '';
	Q := OpenSQL ('SELECT DOS_NODOSSIER FROM DOSSIER WHERE '+GererCritereGroupeConfTous, True);
	While Not Q.EOF Do
	Begin
		Chaine := Chaine + 'PGS_NOMSERVICE LIKE "'+Q.FindField('DOS_NODOSSIER').AsString+'%" OR ';
		Q.Next;
	End;
	Ferme(Q);

    // Cas particulier de l'accès au menu des services qui autorise la gestion des services non rattachés à un dossier
    If JaiLeDroitTag(46530) Then
    Begin
        Chaine := Chaine + 'PGS_NOMSERVICE NOT LIKE "%#_%" ESCAPE "#" OR ';
    End;

	Chaine := Copy(Chaine,1,Length(Chaine)-3);
	Result := Chaine;
End;

end.
