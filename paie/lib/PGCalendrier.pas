{***********UNITE*************************************************
Auteur  ...... : Paie PGI
Créé le ...... : 19/05/2003
Modifié le ... :   /  /
Description .. : Méthode utilisée par la paie pour sa gestion des calendriers
Mots clefs ... : PAIE;CALENDRIER
*****************************************************************}
{
PT29REP    : 29/04/2002 SB V571 Fiche de bug n°306,367,10081 : Calcul du décompte des jours eronnés
                           Modification de la fn IfJourNonTravaille intégrant
                           des contrôles de jours spécifiques
                           Ajout d'arrondi pour l'affichage des heures
PTPG30REP  : 30/05/2002 SB V582 Controle de vraisemblance, Jour spécifique
                           réinitialise systématique le jour comme travaillé
PTPG30-2REP: 07/06/2002 SB V852 Controle de vraisemblance, Calendrier bimensuel
                           personnalisé non récupérer
PT35REP    : 25/07/2002 SB V585 Dû à l'intégration des mvts d'annulation d'absences
                           Contrôle des requêtes si typemvt et sensabs en
                           critere
PT48-1REP  : 20/12/2002 SB V591 FQ 10391 Calendrier pris en compte de la date entrée
                           et sortie pour calcul tps de travail
POINT REPRIS DE PGEdtEtat : fonctions et procedures importées
----------------------------------------------------
PT1    : 19/05/2003 SB V_42 FQ 10693 Intégration décompte des jours féries
                            travaillés
PT2    : 28/10/2003 SB V_42 Chargement des jours féries incomplet
PT3    : 17/05/2004 SB V_50 FQ 11298 Intégration de la notion multi sur jours
                            fériés
PT4    : 21/10/2004 MF V_60 POURV6CORRECTIVE - Correction chargement des jours
                            fériés (il manquait des jours fériés qd la période
                            de recherche s'étalait sur 3 mois ou +). Le calcul
                            des jours d'absence pouvait donc être faux.
PT5    : 14/02/2005 SB V_60 FQ 11967 Chargement des jours spécifiques calendaire
                            supérieur à la date début de calcul
PT6    : 29/03/2005 SB V_60 FQ 11990 Ajout option pour rech. etab de
                            l'historique
PT7    : 23/01/2006 SB V_65 FQ 10866 Ajout clause predefini motif d'absence
PT8    : 07/04/2006 SB V_65 FQ 12941 Traitement nouveau mode de décompte motif
                            d'absence
PT9    : 13/04/2006 SB V_65 Traitement gestion maximum
PT10   : 14/04/2006 SB V_65 FQ 12696 Correction proc. recherche jour repos
                            hebdomadaire
PT11   : 21/04/2006 SB V_65 FQ 13126 Ajout méthode pour contrôle affaire
                            existante
PT12   : 19/06/2006 SB V_65 FQ 13231 Retrait des mvt absences annulées
PT13   : 03/10/2006 SB V_65 FQ 13558 Correction anomalie chargement jour férié
PT14   : 17/10/2006 SB V_70 Refonte Fn pour utilisation portail
PT15   : 25/01/2007 SB V_70 Compatibilité format date internationalisation
PT16   : 02/02/2007 SB V_70 Dplct de la function PGEncodeDateBissextile
PT17   : 18/04/2007 PH V_70 Pas de tests planninification dans le cas de dossier
                            optimisé PCL car les tables ne sont pas présentes.
PT18   : 10/05/2007 NA V-70 FQ 14093 Décompte des jours erronés en cas de
                            calendrier salarié
PT19   : 15/05/2007 MF V_72 Pour mise au point PortailPaie.dll
PT20   : 28/08/2007 FC V_72 Extraction de fonctions dans ULibEditionPaie
PT21   : 14/09/2007 VG V_80 Plantage en saisie d'absence
PT22   : 28/09/2007 NA V_80 FQ 14811 Modification du libellé d'absence afin de supprimer le am et pm
PT24   : 10/10/2007 MF V_80 Mise au point portailpaie.dll
                            pb sur vignette de saisie des absences
                            le remplaçant sur site ne faisait pas partie de la Tob T_Abs
PT25   : 30/10/2007 GGU V_80 Gestion des absences en horaires
PT28   : 18/03/2008 FLO V_80 Ajout de la fonction ChargeTobInterimaire, duplication de ChargeTobSalarie
}

unit PGCalendrier;

interface

uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     SysUtils,
{$IFDEF EAGLCLIENT}

{$ELSE}
     DB,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     Classes,Controls,HCtrls,HEnt1,UTOB,paramsoc,hmsgbox,Hdebug
     {$IFDEF MODESMTP}
     {$ELSE}
     ,MailOL
     {$ENDIF}
     ,ULibEditionPaie //PT20
     ;



          { Charge le calendrier salarié, établissement, les jours fériés et absences sous forme de Tob }
Procedure ChargeTobCalendrier(DateDeb,DateFin : TDateTime; Salarie : String; AvecAbsence,AvecFerie,AvecHistoEtab : Boolean; Var  Tob_Semaine,Tob_Standard,Tob_JourFerie,Tob_absence : Tob;var Etab,Calend,StandCalend,EtabStandCalend,NbJrTravEtab,Repos1,Repos2,JourHeure : string; var DateEntree,DateSortie : TDateTime );
          { Renvoie la requête pour chargement calendrier défini dans la table salarié }
function  RendQueryCalendrier(DateDeb : TDateTime ;Salarie : string ;var Etab,Calend,StandCalend,EtabStandCalend,NbJrTravEtab,Repos1,Repos2,JourHeure : string ; var DateEntree,DateSortie : TDateTime; RechHisto : Boolean = False) : String;
          { Revoie le calendrier standard établissement }
Function  RendTobCalStandard(StandCalend,EtabStandCalend : string; DateDeb : TDateTime ) : Tob;
          { Renvoie True si jour non travaillé au niveau salarié }
function  IfJourNonTravaille(Tob_Calendrier,Tob_Standard,Tob_JourFerie: Tob; DateAbsence,DateEntree,DateSortie:TDateTime;NbJrTravEtab,Repos1,Repos2 : string;DecompteCP,DecompteTps : Boolean; var nb_j,nb_h : double; Var AM,PM : Boolean) : Boolean;
          { Renvoie True si jour non travaillé au niveau établissement }
Function  IfJourNonTravailStd(Tob_Standard,Tob_JourFerie :Tob; DateAbsence : TDateTime; DecompteCP: Boolean; Var nb_j,nb_h : double; Var AM,PM : Boolean) : Boolean;
          { Renvoie True si jour Férié }
Function  IfJourFerie(Tob_JourFerie:Tob;DateAbsence : TDateTime): Boolean ;
          { Renvoie True si jour Férié - Prends en compte l'année pour les jours fériés dont la date n'est pas fixe }
Function  IsJourFerie(Tob_JourFerie:Tob;DateAbsence : TDateTime): Boolean ;
          { Renvoie True si jour repos hebdomadaire }
Procedure IfJourReposHebdo ( DOWJour : integer; NbJrTravEtab,Repos1,Repos2 : String; var JourRepos : Boolean) ;
{ DEB PT8 }
          { Charge les jours fériés dans la périodicité }
Function  ChargeTobFerie(DateDeb,DateFin : TDateTime) : Tob;
          { Décompte des jours calendaires (éphéméride) }
Procedure DecompteCalendrierCivil(DateDeb,DateFin : TDateTime; Tob_Ferie : Tob; Var nb_j,nb_h : double; Nodjm: integer = 0; Nodjp: integer = 1);
          { Décompte des jours en ouvrés (sam..dim..exclus) }
Procedure DecompteOuvre(DateDeb,DateFin : TDateTime; Tob_Ferie : Tob; Var nb_j,nb_h : double; Nodjm: integer = 0; Nodjp: integer = 1);
          { Décompte des jours en ouvrables (dim..exclus) }
Procedure DecompteOuvrable(DateDeb,DateFin : TDateTime; Tob_Ferie : Tob; Var nb_j,nb_h : double; Nodjm: integer = 0; Nodjp: integer = 1);
          { Décompte des jours sur la base du calendrier salarié }
procedure DecompteCalendrierSalarie(DateDebut, DateFin: TDatetime; Salarie, Etablissement, TypeMvt: string; Tob_JourFerie : Tob; var Nb_J, Nb_H: double; Nodjm: integer = 0; Nodjp: integer = 1);
          { Calcul le nombre de jours d'absence en fonction du paramétrage du motif d'absence }
Procedure CalculNbJourAbsence(DateDebut, DateFin: TDatetime; Salarie, Etablissement, TypeMvt: string; T_MotifAbs : Tob; var Nb_J, Nb_H: double; Nodjm: integer = 0; Nodjp: integer = 1);
{ FIN PT8 }
          { Test paramétrage gestion maximum définit dans les motifs d'absence }
Function  ControleGestionMaximum(Salarie,Typeconge : String; T_MotifAbs : Tob; DateAbs : TDateTime; Nb_j, Nb_H : Double ) : Boolean; { PT9 }
          { Affectation du codage demi-journée }
procedure AffecteNodemj(valzone: string; var Nodj: integer);
          { Test présence dans la gestion d'affaire }
Function  EstLibre(pDtDebut, pDtFin : TDateTime; pStRes, pStDebut, pStFin : String; var pStAffaire : String): Boolean; { PT11 }
{ DEB PT14 }
          { Formate PCN_LIBELLE pour harmonisation affichage sur bulletin }
Function  RendLibAbsence(TypeConge,DebDj,FinDj : string; DD,DF : TDateTime) : string;
          { Formate enregistrement tout mvt d'absence standard }
procedure InitialiseTobAbsenceSalarie(TCP: tob);
          { Formate enregistrement absence PRI ou ABS }
Procedure InitialiseTobAbsMotif ( T_Abs, T_Param,  T_Sal  :  Tob; ValResp,Exp,SaisiePar,SaisieVia : String);
          { Définit PCN_ORDRE pour création d'absence }
function  IncrementeSeqNoOrdre(cle_type, cle_sal: string): Integer;
          { Test compteur récapitulatif pour alerte prise d'absence }
function  ControleAssezRestant(T : Tob; {Salarie,} TYPECONGE, Action: string; Ejours, EJoursPris : double): boolean;
          { Formate SQL pour cotrôle absence existante }
function  PgGetClauseAbsAControler(T_MotifAbs,LcTob_MotifAbs : tob; TypeConge : String) : string;
Function  ChargeTobSalarie (Sal : String) : Tob;
Function  ChargeTobInterimaire (Sal : String) : Tob; //PT28
          { Recup info resp & sal pour paramétrage et envoie mail }
procedure ChargeInfoRespSalMail(Sal, FicPrec, TypeSal : String; Var Resp,NomResp,MailResp,MailSal : string; Var OKMailEnvoie : Boolean; Var MailDate : TDateTime);
Function  ChargeTob_Recapitulatif( Sal : String) : tob;
procedure PGExeCalculRecapAbsEnCours(Salarie: string);
Procedure PgSupprRecapitulatif(Tob_motifAbs : Tob; Sal, TypeConge : String; EJoursPris : Double; Var Error : integer; Var Retour : string);
          { function d'envoie de mail }
function  PGEnvoieMailCollaborateur(Titre, Mail: string; Texte: TStringList): Boolean;
          { Ensemble des contrôles sur saisie d'absence PRI ou ABS }
Procedure PgControleTobAbsFirst (T_Abs, TSal : TOB;  {Origine, Action, ficPrec,TypeSal : String;} Var Error : Integer; Var NomControl, ComplMessage : String)  ;
Procedure PgControleTobAbsSecond (T_Abs, TSal : TOB; Origine, Action, ficPrec,TypeSal : String; Var Error : Integer; Var NomControl, ComplMessage : String)  ;
Function  RecupMessageErrorAbsence(Error : integer) : string;
Function  PgMajAbsEtatValidSal(Sal,Mode,Typmvt,ValidPar,ValidResp : String ; ordre : integer) : Integer;
Procedure PgMajIndicMailSalRAZ(ValidPar : String) ;
Procedure PgMajIndicMailSalENV(ValidPar : String) ;
Function  PgRecupBodyMailCollaborateur(T_Abs : Tob; Resp : String) : TStringList;
{ FIN PT14 }
{ DEB PT15 }
Function  PgCalculDateAnciennete(DateArret : TDateTime; DateCpAnc : String ) : TDateTime;
Function  PgOkFormatDateJJMM ( St : String)  : Boolean;
{ FIN  PT15 }
//PT20 function  PGEncodeDateBissextile(AA, MM, JJ: WORD): TDateTime;  { PT16 }

implementation
{$IFNDEF PLUGIN}   // PT19
uses UtilPgi;
{$ENDIF PLUGIN}    //P T19

Procedure ChargeTobCalendrier(DateDeb,DateFin : TDateTime; Salarie : String; AvecAbsence,AvecFerie,AvecHistoEtab : Boolean; Var  Tob_Semaine,Tob_Standard,Tob_JourFerie,Tob_absence : Tob;var Etab,Calend,StandCalend,EtabStandCalend,NbJrTravEtab,Repos1,Repos2,JourHeure : string; var DateEntree,DateSortie : TDateTime );
Var
Q : TQuery;
StQuery : string;
Begin
//Init des jours fériés
if AvecFerie then  Tob_JourFerie := ChargeTobFerie(DateDeb,DateFin); { PT8 } 
//Init du calendrier
//PT- 1REP Ajout StQuery
StQuery:=RendQueryCalendrier(DateDeb,Salarie,Etab,Calend,StandCalend,EtabStandCalend,NbJrTravEtab,Repos1,Repos2,JourHeure,DateEntree,DateSortie,AvecHistoEtab); { PT6 }
if StQuery<>'' then Q:=OpenSql(StQuery,True) else  begin Tob_Semaine:=nil; exit; end;
Tob_Semaine:=ToB.create('Les jours de la semaine',nil,-1);
Tob_Semaine.LoadDetailDB('CAL_ENDRIER','','',Q,False);
Ferme(Q);
//Chargement du standard Etablissement si Stand = Personnalisé ou autre standard
Tob_Standard:=RendTobCalStandard(StandCalend,EtabStandCalend,DateDeb);  { PT5 }
if not (Assigned(Tob_Standard)) then
  Begin
  Tob_Standard:=Tob.create('Le standard calendrier',nil,-1);
  Tob_Standard.Dupliquer(Tob_Semaine,True,True);
  End;

//Init des congés payés pris et des absences
if AvecAbsence then
  Begin
  Q:=OpenSql('SELECT PCN_SALARIE,PCN_TYPECONGE,PCN_DATEDEBUTABS,PCN_DATEFINABS,'+
  'PCN_TYPEMVT,PCN_ORDRE,PCN_JOURS,PCN_HEURES,'+
  'PMA_LIBELLE,PMA_ABREGE,PMA_EDITION,PMA_PRISTOTAL FROM ABSENCESALARIE '+
  'LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI## PMA_MOTIFABSENCE=PCN_TYPECONGE '+    { PT7 }
  'WHERE PCN_SALARIE="'+salarie+'" AND ((PCN_DATEDEBUTABS>="'+USDateTime(DateDeb-1)+'" '+
  'AND PCN_DATEDEBUTABS<="'+USDateTime(DateFin)+'") '+ //PT29REP DateDeb-1 : Cas abs fin de mois anterieur, et 1er jour mois en cours repos salarié
  'OR (PCN_DATEFINABS>="'+USDateTime(DateDeb-1)+'" '+
  'AND PCN_DATEFINABS<="'+USDateTime(DateFin)+'" )) '+
  'AND ((PCN_TYPEMVT="ABS" AND PCN_SENSABS="-") OR (PCN_TYPECONGE="PRI" AND PCN_TYPEMVT="CPA" AND PCN_MVTDUPLIQUE="-" )) '+ //PT35REP Ajout AND PCN_SENSABS="-"
  'AND PCN_ETATPOSTPAIE <> "NAN" AND PMA_EDITION="X" '+  { PT12 }
  'ORDER BY PCN_TYPECONGE,PCN_DATEDEBUTABS,PCN_DATEFINABS ',True);
  Tob_absence:=Tob.Create('Les absences',nil,-1);
  Tob_Absence.LoadDetailDB('ABSENCESALARIE','','',Q,False);
  Ferme(Q);
  End;

End;

function RendQueryCalendrier(DateDeb : TDateTime ; Salarie : string ;var Etab,Calend,StandCalend,EtabStandCalend,
NbJrTravEtab,Repos1,Repos2,JourHeure : string;var DateEntree,DateSortie : TDateTime; RechHisto : Boolean = False) : String;
var
Q : TQuery;
StQ : String;
Begin
result:='';
Calend:='';  StandCalend:='';   EtabStandCalend:='';
NbJrTravEtab:=''; Repos1:='';  Repos2:='';        JourHeure:='';
DateEntree:=idate1900;  DateSortie:=idate1900;
  //PT48-1REP Ajout date entrée
  //Voir si salarie n'a pas changé d'établissement
  If RechHisto then { PT6 }
    Begin
    Q:=OpenSql('SELECT PSA_SALARIE,PSA_ETABLISSEMENT,PPU_ETABLISSEMENT '+
    'FROM SALARIES LEFT JOIN PAIEENCOURS ON PPU_SALARIE=PSA_SALARIE '+
    'AND PPU_DATEDEBUT>="'+UsDateTime(DebutdeMois(DateDeb))+'" '+
    'AND PPU_DATEFIN<="'+UsDateTime(FindeMois(DateDeb))+'" '+
    'WHERE PSA_SALARIE="'+salarie+'" ',True);
    if not Q.eof then
      Etab:=Q.FindField('PPU_ETABLISSEMENT').asstring;  //Etab table paieencours
    if (Trim(Etab)='') and (not Q.eof) then
      Etab:=Q.FindField('PSA_ETABLISSEMENT').asstring;  //Etab table salaries
    Ferme(Q);
    End;
//init du paramétrage calendrier du salarié
Q := OpenSql('SELECT PSA_SALARIE,PSA_ETABLISSEMENT,PSA_CALENDRIER,PSA_STANDCALEND,'+
     'PSA_DATEENTREE,PSA_DATESORTIE,PSA_JOURHEURE,'+
     'ETB_STANDCALEND,ETB_NBJOUTRAV,ETB_1ERREPOSH,ETB_2EMEREPOSH '+
     'FROM SALARIES '+
     'LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT="'+Etab+'" '+
     'WHERE PSA_SALARIE="'+Salarie+'"',True);
  if not Q.eof then
    Begin
    If not RechHisto then Etab := Q.FindField('PSA_ETABLISSEMENT').asstring; { PT6 }
    Calend:=Q.FindField('PSA_CALENDRIER').asstring;
    StandCalend:=Q.FindField('PSA_STANDCALEND').asstring;
    EtabStandCalend:=Q.FindField('ETB_STANDCALEND').asstring;
    NbJrTravEtab:=IntToStr(Q.FindField('ETB_NBJOUTRAV').AsInteger);      // DB2
    Repos1:=Q.FindField('ETB_1ERREPOSH').asstring; if Repos1<>'' then repos1:=Copy(repos1,1,1);
    Repos2:=Q.FindField('ETB_2EMEREPOSH').asstring; if Repos2<>'' then repos2:=Copy(repos2,1,1);
    DateEntree:=Q.FindField('PSA_DATEENTREE').AsDateTime;
    DateSortie:=Q.FindField('PSA_DATESORTIE').AsDateTime;
    JourHeure:=Q.FindField('PSA_JOURHEURE').AsString;
    End;
  Ferme(Q);

 //1 - on cherche au niveau calendrier spécifique salarie
 //PSA_SALARIE=CAL1.ACA_SALARIE
if  StandCalend='PER' then
    Begin
    Q:=Opensql('SELECT * FROM CALENDRIER WHERE ACA_SALARIE="'+Salarie+'" AND (ACA_DATE="'+UsDateTime(Idate1900)+'" OR ACA_DATE>="'+UsDateTime(DateDeb)+'") ORDER BY ACA_JOUR',True); { PT5 }
    if (Q.eof) and (calend<>'') then
       Begin
       //2 - on cherche au niveau standard de calendrier salarie
       StQ:='SELECT * FROM CALENDRIER WHERE ACA_STANDCALEN="'+Calend+'" AND ACA_SALARIE="***" AND (ACA_DATE="'+UsDateTime(Idate1900)+'" OR ACA_DATE>="'+UsDateTime(DateDeb)+'") ORDER BY ACA_JOUR'; { PT5 }
       End
     else    //DEB PT30-2REP
        if (not Q.eof) then
          StQ:='SELECT * FROM CALENDRIER WHERE ACA_SALARIE="'+Salarie+'" AND (ACA_DATE="'+UsDateTime(Idate1900)+'" OR ACA_DATE>="'+UsDateTime(DateDeb)+'") ORDER BY ACA_JOUR'; { PT5 }
    Ferme(Q); //FIN PT30-2REP
    End;
//2 - on cherche au niveau standard de calendrier salarie
//ON PSA_CALENDRIER=CAL2.ACA_STANDCALEN AND PSA_STANDCALEND="ETS"
if  StandCalend='ETS' then
    Begin
    if Calend='' then exit;
    StQ:='SELECT * FROM CALENDRIER WHERE ACA_STANDCALEN="'+Calend+'" AND ACA_SALARIE="***" AND (ACA_DATE="'+UsDateTime(Idate1900)+'" OR ACA_DATE>="'+UsDateTime(DateDeb)+'") ORDER BY ACA_JOUR'; { PT5 }
    End;
//3 - on cherche au niveau calendrier standard établissement
//ACA_STANDCALEN=ETB_STANDCALEND
if  StandCalend='ETB' then
    Begin
    if EtabStandCalend='' then exit;
    StQ:='SELECT * FROM CALENDRIER WHERE ACA_STANDCALEN="'+EtabStandCalend+'" AND ACA_SALARIE="***" AND (ACA_DATE="'+UsDateTime(Idate1900)+'" OR ACA_DATE>="'+UsDateTime(DateDeb)+'") ORDER BY ACA_JOUR'; { PT5 }
    End;
result:=StQ;
End;

 {***********A.G.L.***********************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 02/05/2002
Modifié le ... :   /  /
Description .. : Chargement du standard Etablissement si Stand =
Suite ........ : Personnalisé ou autre standard
Mots clefs ... : PAIE,CP,CALEND
*****************************************************************}
//DEB PT29REP Chargement du standard Etablissement si Stand = Personnalisé ou autre standard
Function  RendTobCalStandard(StandCalend,EtabStandCalend : string; DateDeb : TDateTime ) : Tob;
Var
Q  : TQuery;
Begin
if (StandCalend<>'ETB') AND (EtabStandCalend<>'') then
  Begin
  Q:=OpenSql('SELECT * FROM CALENDRIER WHERE ACA_STANDCALEN="'+EtabStandCalend+'" AND ACA_SALARIE="***" '+
             'AND (ACA_DATE="'+UsDateTime(Idate1900)+'" OR ACA_DATE>="'+UsDateTime(DateDeb)+'") '+ { PT5 }
             'ORDER BY ACA_JOUR',True);
  Result:=ToB.create('Le calendrier standard',nil,-1);
  Result.LoadDetailDB('CALENDRIER_STANDARD','','',Q,False);
  Ferme(Q);
  End
else
    Result:=nil;
end;
//FIN PT29REP



{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 18/09/2002
Modifié le ... : 18/09/2002
Description .. : Fonction indiquant si un jour calendaire est travaillé ou non
Suite ........ : pour un salarié en fonction :
Suite ........ : - de son calendrier
Suite ........ : - du standard calendrier
Suite ........ : - des jours spécifiques travaillés ou chômés
Suite ........ : - des jours feriés
Suite ........ : - des jours repos établissement
Mots clefs ... : PAIE;CALENDRIER;CP
*****************************************************************}
function  IfJourNonTravaille(Tob_Calendrier,Tob_Standard,Tob_JourFerie: Tob; DateAbsence,DateEntree,DateSortie:TDateTime;NbJrTravEtab,Repos1,Repos2 : string;DecompteCP,DecompteTps : Boolean; var nb_j,nb_h : double; Var AM,PM : Boolean) : Boolean;
var
YY,MM,JJ : WORD;
DOWJour : Integer;
//Duree : String;
T_Semaine,T_Jr,T_JrSpec : Tob;
Begin
Result:=False;   nb_h:=0; nb_j:=0; AM:=False; PM:=False;
if Tob_Calendrier=nil then Exit; //PT30REP
T_Semaine:=TOB.Create('La calendrier dupliquée',nil,-1);
T_Semaine.Dupliquer(Tob_Calendrier,TRUE,TRUE);
DecodeDate(DateAbsence,YY,MM,JJ);
//Test si jour>Date Entrée et <DateSortie
   if (DateAbsence<DateEntree)  then begin result:=True; exit; End; //PT29REP Arret des controles si salarie non entrée
if (DateSortie>2) and (DateSortie<>null) then
   if (DateAbsence>DateSortie) then begin result:=True; exit; End;  //PT29REP Arret des controles si salarie sortie
//Test si jour non férié
If IfJourFerie(Tob_JourFerie,DateAbsence)=True then result:=True;
//Renvoie jour de la semaine (DOWJour) qui correspond au jour d'absence
DOWJour:=DayOfWeek(DateAbsence);
//Test si DOWJour correspond pas à un jour de repos
IfJourReposHebdo(DOWJour,NbJrTravEtab,Repos1,Repos2,result);
//Correspondance des jours
  if DOWJour=1 then DOWJour:=7 else DOWJour:=DOWJour-1;
//Test si DOWJour correspond à un jour non travaillé du calendrier salarié
if (DateAbsence=DebutDeMois(DateAbsence)) OR  (DOWJour=1) then
   T_Jr:=T_Semaine.FindFirst(['ACA_JOUR'],[DOWJour],False)
else T_Jr:=T_Semaine.FindNext(['ACA_JOUR'],[DOWJour],False);
//Test si jour ne correspond pas à un jour spécifique du calendrier salarié ou du standard
if DecompteTps then
  Begin
  T_JrSpec:=T_Semaine.FindFirst(['ACA_DATE'],[DateAbsence],False); //'ACA_FERIETRAVAIL' ,'-'
  If T_JrSpec<>nil then //La spécificité salarié passe avant celle de l'étab
    Begin
    T_Jr.Dupliquer(T_JrSpec,True,True);
    result:=False;  //PT30REP mise en commentaire
    End
  else
    Begin
    if Tob_Standard<>nil then   //DEB PT29REP Modification du code ajout de test
      T_JrSpec:=Tob_Standard.FindFirst(['ACA_DATE'],[DateAbsence],False);
    If T_JrSpec<>nil then T_Jr.Dupliquer(T_JrSpec,True,True);
    End;
  End;
if T_Jr<>nil then
  Begin
   // result:=False; PT30REP Mise en commentaire //PT29REP Recupération des jours et heures calendrier
  if (IsNumeric(T_Jr.GetValue('ACA_HEUREDEB1'))) and (IsNumeric(T_Jr.GetValue('ACA_HEUREFIN1'))) then
    if (T_Jr.GetValue('ACA_HEUREDEB1')<>0) and (T_Jr.GetValue('ACA_HEUREFIN1')<>0) then
      begin nb_j:=0.5;   AM:=True; End;
  if (IsNumeric(T_Jr.GetValue('ACA_HEUREDEB2'))) and (IsNumeric(T_Jr.GetValue('ACA_HEUREFIN2'))) then
    if (T_Jr.GetValue('ACA_HEUREDEB2')<>0) and (T_Jr.GetValue('ACA_HEUREFIN2')<>0) then
      Begin nb_j:=nb_j+0.5;  PM:=True; End;
  if IsNumeric(T_Jr.GetValue('ACA_DUREE')) then nb_h:=Arrondi(T_Jr.GetValue('ACA_DUREE'),2);
  if Nb_h=0 then Result:=True;
  T_Jr.free; T_Jr:=nil;
  end;
//Pour le calcul d'une duree d'absence des tests spécifiques viennent s'ajouter :
if (DecompteCP) and (result=True)  then
   Begin
{Pour un calendrier spécifique salarié on se calque sur le standard étab pour tester si le jour n'est pas
travaillé dans l'établissement : dans ce cas on décompte ce jour}
   Result:=IfJourNonTravailStd(Tob_Standard,Tob_JourFerie,DateAbsence,DecompteCP,Nb_J,Nb_H,AM,PM);
   End;                     //FIN PT29REP Modification du code ajout de test
if T_Jr<>nil  then T_Jr.free;
if T_Semaine<>nil then T_Semaine.free;
End;

Function  IfJourNonTravailStd(Tob_Standard,Tob_JourFerie :Tob; DateAbsence : TDateTime; DecompteCP: Boolean; Var nb_j,nb_h : double; Var AM,PM : Boolean) : Boolean;
Var DOWJour : integer; T_Jr : Tob;
    JourFerie               : Boolean;
Begin
Result:=True;
nb_j:=0; nb_h:=0;
AM:=False; PM:=False; JourFerie:=False;
if (DecompteCP) then
  Begin
  //Test si jour non férié
  If IfJourFerie(Tob_JourFerie,DateAbsence)=True then
    begin
    result:=True; JourFerie:=True;
    nb_j:=-1;
    End;
  End;
//DEB PT1 On vérifie que le jour n'est pas un jour spécifique travaillé
if Assigned(Tob_Standard) then
  Begin
  T_Jr:=Tob_Standard.FindFirst(['ACA_DATE'],[DateAbsence],False);
  if Assigned(T_jr) then
    Begin
    //Si jour férié, alors travaille autorisé coché
    if (result=True) and (T_jr.GetValue('ACA_FERIETRAVAIL')='X' ) then begin result:=False; nb_j:=0; end;
    if (T_Jr.GetValue('ACA_HEUREDEB1')<>0) and (T_Jr.GetValue('ACA_HEUREFIN1')<>0) then Begin nb_j:=0.5;      AM:=True; End;
    if (T_Jr.GetValue('ACA_HEUREDEB2')<>0) and (T_Jr.GetValue('ACA_HEUREFIN2')<>0) then Begin nb_j:=nb_j+0.5; PM:=True; End;
    nb_h:=Arrondi(T_Jr.GetValue('ACA_DUREE'),2);
    if Nb_h=0 then Result:=True;
    End;
  If (nb_j=0) and (JourFerie) then Nb_J:=-1; 
  if (nb_j<>0) then exit;
  End;
//FIN PT1
DOWJour:=DayOfWeek(DateAbsence);
if DOWJour=1 then DOWJour:=7 else DOWJour:=DOWJour-1;
if Assigned(Tob_Standard) then
  Begin
  T_Jr:=Tob_Standard.FindFirst(['ACA_JOUR'],[DOWJour],False);
  if T_jr<>nil then
    Begin
    nb_h:=0; nb_j:=0; AM:=False; PM:=False;
    result:=False;
    if (T_Jr.GetValue('ACA_HEUREDEB1')<>0) and (T_Jr.GetValue('ACA_HEUREFIN1')<>0) then Begin nb_j:=0.5;      AM:=True; End;
    if (T_Jr.GetValue('ACA_HEUREDEB2')<>0) and (T_Jr.GetValue('ACA_HEUREFIN2')<>0) then Begin nb_j:=nb_j+0.5; PM:=True; End;
    nb_h:=Arrondi(T_Jr.GetValue('ACA_DUREE'),2);
    if Nb_h=0 then Result:=True;
    End
  End;
End;
Function  IfJourFerie(Tob_JourFerie:Tob;DateAbsence : TDateTime): Boolean ;
Var
YY,MM,JJ : Word;
Ijf : integer;
Begin
Result:=False;
if Tob_JourFerie<>nil then
  Begin
  DecodeDate(DateAbsence,YY,MM,JJ);
  For Ijf:=0 to Tob_JourFerie.detail.count-1 do
    If MM=Tob_JourFerie.detail[Ijf].GetValue('AJF_MOIS') then
      if JJ=Tob_JourFerie.detail[Ijf].GetValue('AJF_JOUR') then Result:=True;
  End;
End;

Function  IsJourFerie(Tob_JourFerie:Tob;DateAbsence : TDateTime): Boolean ;
Var
YY,MM,JJ : Word;
Ijf : integer;
Begin
Result:=False;
if Tob_JourFerie<>nil then
  Begin
  DecodeDate(DateAbsence,YY,MM,JJ);
  For Ijf:=0 to Tob_JourFerie.detail.count-1 do
    If MM=Tob_JourFerie.detail[Ijf].GetValue('AJF_MOIS') then
      if JJ=Tob_JourFerie.detail[Ijf].GetValue('AJF_JOUR') then
        if  (YY = Tob_JourFerie.detail[Ijf].GetValue('AJF_ANNEE'))
         or (Tob_JourFerie.detail[Ijf].GetValue('AJF_JOURFIXE') = 'X') then
        Result:=True;
  End;
End;

Procedure IfJourReposHebdo ( DOWJour : integer; NbJrTravEtab,Repos1,Repos2 : String; var JourRepos : Boolean) ;
Begin
//1er repos
if (NbJrTravEtab='5')  then  { DEB PT10 }
   Begin
   if (repos1<>'') AND (DOWJour=StrToInt(Repos1)) then JourRepos:=True;
   if (repos2<>'') AND (DOWJour=StrToInt(Repos2)) then JourRepos:=True;
   End
else
   //2eme repos
   if (NbJrTravEtab='6') and (repos1<>'')  then
     if DOWJour=StrToInt(Repos1) then JourRepos:=True;
				{ FIN PT10 }
End;

{ PT8 : Nouvelle Fonction}
Function  ChargeTobFerie(DateDeb,DateFin : TDateTime) : Tob;
Var YYD,YYF,MMD,MMF,JJ : Word;
    StWhere : String;
    Q : TQuery;
Begin
  StWhere := 'WHERE ##AJF_PREDEFINI##  AJF_CODEFERIE <> "" ';
  if DateDeb <> Idate1900 then
     Begin
     DecodeDate(DateDeb,YYD,MMD,JJ);
     DecodeDate(DateFin,YYF,MMF,JJ);
    (* PT13 Mise en commentaire, refonte de la clause pour prise en compte d'un max de jour
    { DEB PT2 }
     StWhere := StWhere + 'AND ((AJF_MOIS>='+IntToStr(MM)+
             ' AND AJF_ANNEE='+IntToStr(YY)+') '+ { PT3 }
     // d PT4
     //  'OR (AJF_MOIS='+IntToStr(MM)+'  AND AJF_JOURFIXE="X")';       // DB2
             'OR (AJF_MOIS>='+IntToStr(MM)+'  AND AJF_JOURFIXE="X"';
     // f PT4
     End;
  if DateFin <> Idate1900 then
    Begin
    DecodeDate(DateFin,YY,MM,JJ);
    // d PT4
    StWhere:=StWhere +' AND AJF_MOIS<='+IntToStr(MM)+')'+
                    ' OR (AJF_MOIS<='+IntToStr(MM)+
                    ' AND AJF_ANNEE='+IntToStr(YY)+'))';

    {  StWhere:=StWhere+' OR (AJF_MOIS>='+IntToStr(MM)+' AND AJF_ANNEE='+IntToStr(YY)+') '+
    'OR (AJF_MOIS='+IntToStr(MM)+'  AND AJF_JOURFIXE="X"))';       // DB2}
// f PT4 *)
    StWhere := StWhere + 'AND (AJF_JOURFIXE="X" OR (AJF_ANNEE>='+IntToStr(YYD)+' AND AJF_JOURFIXE="-"))';
    if (YYF = YYD) AND (MMD=MMF) then StWhere := StWhere + ' AND AJF_MOIS='+IntToStr(MMD)
    else
      if (YYF = YYD) AND (MMD<MMF) then StWhere := StWhere + ' AND AJF_MOIS>='+IntToStr(MMD)+' AND AJF_MOIS<='+IntToStr(MMF);
    End;
  Q:=OpenSql('SELECT * FROM JOURFERIE '+StWhere,True);
  { FIN PT2 }
  Result:=Tob.create('Les jours feriés',nil,-1);
  Result.LoadDetailDB('JOURFERIE','','',Q,False);
  Ferme(Q);
End;

{ PT8 : Nouvelle procedure }
Procedure DecompteCalendrierCivil(DateDeb,DateFin : TDateTime; Tob_Ferie : Tob; Var nb_j,nb_h : double; Nodjm: integer = 0; Nodjp: integer = 1);
Var i : integer;
Begin
  i := 0; Nb_J := 0;
  While DateDeb <= DateFin do
    Begin
      If (IfJourFerie(Tob_Ferie,DateDeb)) then
        Begin
        Inc(i);
        DateDeb := DateDeb + 1;
        Continue;
        End;
    Nb_J := Nb_J + 1;
    If (i = 0) And (NoDjm = 1) then Nb_J := Nb_J - 0.5;
    if (DateDeb=DateFin) And  (NoDjp = 0) then Nb_J := Nb_J - 0.5;
    Inc(i);
    DateDeb := DateDeb + 1;
    End;
  Nb_H := Nb_J * 7;
End;


{ PT8 : Nouvelle procédure }
Procedure DecompteOuvre(DateDeb,DateFin : TDateTime; Tob_Ferie : Tob; Var nb_j,nb_h : double; Nodjm: integer = 0; Nodjp: integer = 1);
Var i : integer;
Begin
  i := 0; Nb_J := 0;
  While DateDeb <= DateFin do
    Begin
    if (IfJourFerie(Tob_Ferie,DateDeb)) OR (DayOfWeek(DateDeb) in [1,7]) then
        Begin
        Inc(i);
        DateDeb := DateDeb + 1;
        Continue;
        End;
    Nb_J := Nb_J + 1;
    If (i = 0) And (NoDjm = 1) then Nb_J := Nb_J - 0.5;
    if (DateDeb=DateFin) And  (NoDjp = 0) then Nb_J := Nb_J - 0.5;
    Inc(i);
    DateDeb := DateDeb + 1;
    End;
  Nb_H := Nb_J * 7;
End;

{ PT8 : Nouvelle Procedure }
Procedure DecompteOuvrable(DateDeb,DateFin : TDateTime; Tob_Ferie : Tob; Var nb_j,nb_h : double; Nodjm: integer = 0; Nodjp: integer = 1);
Var i : integer;
Begin
  i := 0; Nb_J := 0;
  While DateDeb <= DateFin do
    Begin
    if (IfJourFerie(Tob_Ferie,DateDeb)) OR (DayOfWeek(DateDeb)= 1 ) then
        Begin
        Inc(i);
        DateDeb := DateDeb + 1;
        Continue;
        End;
    Nb_J := Nb_J + 1;
    If (i = 0) And (NoDjm = 1) then Nb_J := Nb_J - 0.5;
    if (DateDeb=DateFin) And  (NoDjp = 0) then Nb_J := Nb_J - 0.5;
    Inc(i);
    DateDeb := DateDeb + 1;
    End;
  Nb_H := Nb_J * 7;
End;


{ PT8 : Nouvelle procédure }
procedure DecompteCalendrierSalarie(DateDebut, DateFin: TDatetime; Salarie, Etablissement, TypeMvt: string; Tob_JourFerie : Tob;  var Nb_J, Nb_H: double; Nodjm: integer = 0; Nodjp: integer = 1);
var
  { PT60 Mise en commentaire puis supprimé }
  calend, StandCalend, EtabStandCalend, NbJrTravEtab, Repos1, Repos2,repos1etb, repos2etb, JourHeure: string;   // pt18
  Tob_Semaine, Tob_Standard, {Tob_JourFerie,} Tob_absence,Tob_Temp : Tob;
  DateEntree, DateSortie, DateDebAbs, DateFinAbs : TDateTime;
  CptJour, CptHeure, Nb_JSal: Double;
  SSDecompte, JourNonTrav, JourNonTravStd, AM, PM, AMDeb, AMFin, PMDeb, PMFin, AMStd, PMStd: boolean;
begin
  { DEB PT60 NOUVELLE FONCTION DE CALCUL DE TEMPS
     Principe de calcul du decompte CP :
  1- On positionne la date de fin à la dernière demi-journée travaillé par le salarié
     On positionne la date de debut à la première demi journée travaillé par l'entreprise
  2- On recherche la demi journée de reprise de travail pour l'entreprise et le salarié
  3- On comptabilise le temps de travail sur le calendrier entreprise
   }
  Nb_J := 0;  { PT93 }
  Nb_H := 0;
  SSDecompte := False;
  //Init Des Tobs
  ChargeTobCalendrier(DateDebut, DateFin, Salarie, FALSE, False, False, Tob_Semaine, Tob_Standard, Tob_Temp, Tob_absence, Etablissement, Calend, StandCalend, EtabStandCalend, NbJrTravEtab, Repos1, Repos2, JourHeure, DateEntree, DateSortie); { PT97 }
  if Tob_Semaine = nil then exit;
  if Tob_Semaine.Somme('ACA_DUREE',[''],[''],False) = 0 then Exit; { PT93 }
  DateDebAbs := DateDebut;
  DateFinAbs := DateFin;
  CptJour := 0;
  CptHeure := 0;
  JourNonTrav := True;
  {Pour calcul congés payés pris on positionne la date de fin sur le dernier jour
  travaillé afin de tester le lendemain du dernier jour travaillé par la suite}
  if TypeMvt = 'PRI' then
  begin
    SSDecompte := (RechDom('PGSALDECOMPTE', salarie, false) = 'X');
    //1- Repositionnement des dates de debut et fin
    while (JourNonTrav = True) and (DateDebAbs <= DateFinAbs) do
    begin
      JourNonTrav := IfJourNonTravaille(Tob_Semaine, Tob_Standard, Tob_JourFerie, DateDebAbs, DateEntree, DateSortie, NbJrTravEtab, Repos1, Repos2, False, True, Nb_jSal, Nb_h, AMDeb, PMDeb);
      if (JourNonTrav = True) then DateDebAbs := DateDebAbs + 1;
    end;
    if DateDebAbs > DateFinAbs then DateDebAbs := DateDebut;
    JourNonTrav := True;
    while (JourNonTrav = True) and (DateFinAbs >= DateDebAbs) do
    begin
      JourNonTrav := IfJourNonTravaille(Tob_Semaine, Tob_Standard, Tob_JourFerie, DateFinAbs, DateEntree, DateSortie, NbJrTravEtab, Repos1, Repos2, False, True, Nb_jSal, Nb_h, AMFin, PMFin);
      if (JourNonTrav = True) { AND (Nodjp=1) PT75-1 } then DateFinAbs := DateFinAbs - 1;
    end;
    if DateFinAbs < DateDebAbs then DateFinAbs := DateFin;
    //2- Recherche de la demi journée de reprise du standard
    JourNonTrav := True;
    while (JourNonTrav = True) and (Nodjp = 1) and (Assigned(Tob_Standard)) do
    begin
      if (DateSortie < DateFinAbs) and (DateSortie > idate1900) then Break; //PT69 PT100
      JourNonTrav := IfJourNonTravaille(Tob_Semaine, Tob_Standard, Tob_JourFerie, DateFinAbs + 1, DateEntree, DateSortie, NbJrTravEtab, Repos1, Repos2, False, True, Nb_j, Nb_h, AMFin, PMFin);
      JourNonTravStd := IfJourNonTravailStd(Tob_Standard, Tob_JourFerie, DateFinAbs + 1, True, Nb_j, Nb_h, AMStd, PMStd);
      {Un jour non travaillé par l'entreprise est soit un jour de repos hebdo, soit un jour férié, soit un jour ouvrable
      Dans le cas d'un jour de repos hebdo , on saute la journée
      Dans le cas d'un jour férié, on vérifie si c'est un jour de travail du calendrier,( Nb_j = -1 pour un jour ferié)
      Tant Que Jour pas travaillé pour salarié et jour travaillé pour l'entreprise}
      if ((JourNonTrav = True) and (JourNonTravStd = False))
        or ((JourNonTrav = True) and (JourNonTravStd = True) and (Nb_j <> -1)) then
      begin
        DateFinAbs := DateFinAbs + 1;
        Nodjp := 1;
      end
      else
        //Cas d'un jour férié correspondant à un jour travaillé dans l'entreprise
        if (JourNonTrav = True) and (JourNonTravStd = True) and (Nb_j = -1) then
      begin
        JourNonTravStd := IfJourNonTravailStd(Tob_Standard, Tob_JourFerie, DateFinAbs + 1, False, Nb_j, Nb_h, AMStd, PMStd);
        if (JourNonTravStd = False) then
        begin
          DateFinAbs := DateFinAbs + 1;
          Nodjp := 1;
        end
        else
          JourNonTrav := False;
      end
      else
        if (JourNonTravStd = True) then JourNonTrav := False;
    end;
  end;
  //PT69 Repositionnement de la date de début et de fin en fonction de la date d'entrée et de sortie
  if DateEntree > DateDebAbs then DateDebAbs := DateEntree;
  if DateEntree > DateFinAbs then
  begin
    Nb_j := 0;
    Nb_h := 0;
    exit;
  end; { PT69 16/04/04 }
  if (DateSortie < DateFinAbs) and (DateSortie > idate1900) then DateFinAbs := DateSortie; { PT100 }
  //3- On comptabilise le temps de travail sur le calendrier entreprise si type Cp acquis = IdemEtab
  repeat
    if (TypeMvt = 'PRI') and (Nb_JSal > 0) then
    begin
      JourNonTrav := IfJourNonTravaille(Tob_Semaine, Tob_Standard, Tob_JourFerie, DateDebAbs, DateEntree, DateSortie, NbJrTravEtab, Repos1, Repos2, False, True, Nb_j, Nb_h, AM, PM);
      //Jour non travaillé pas le salarié
      if ((JourNonTrav = True) or (NB_J < 1)) and (SSDecompte = False) and (Assigned(Tob_Standard)) then //si travail demi journée salarié { PT96 }
      begin
        JourNonTrav := IfJourNonTravailStd(Tob_Standard, Tob_JourFerie, DateDebAbs, True, Nb_j, Nb_h, AM, PM);
        //Dans le cas d'un jour non travaillé mais ouvrable on comptabilise
        if (JourNonTrav = True) and (Nb_j = 0) then
          if (IntToStr(DayOfWeek(DateDebAbs)) <> Repos1) and (IntToStr(DayOfWeek(DateDebAbs)) <> Repos2) then
            CptJour := CptJour + 1;
        //Dans le cas d'un jour travaillé par l'entreprise mais pas par le salarié, on décompte si case non cochée
        if (JourNonTrav = False) and (SSDecompte = True) then JourNonTrav := True;
      end;
    end
    else //Ou on comptabilise sur le calendrier salarié pour les autres motifs d'absence, le calcul du temps de travail..
    // debut pt18
    begin
      if (typeMvt <> 'PRI') then
      begin
      repos1etb := '0'; repos2etb :='0';
      JourNonTrav := IfJourNonTravaille(Tob_Semaine, Tob_Standard, Tob_JourFerie, DateDebAbs, DateEntree, DateSortie, NbJrTravEtab, Repos1etb, Repos2etb, False, True, Nb_j, Nb_h, AM, PM); { PT77 }
      end  // fin pt18
      else
      JourNonTrav := IfJourNonTravaille(Tob_Semaine, Tob_Standard, Tob_JourFerie, DateDebAbs, DateEntree, DateSortie, NbJrTravEtab, Repos1, Repos2, False, True, Nb_j, Nb_h, AM, PM); { PT77 }
    end;
    //une fois la période de calcul des jours calculé on peut décompter le temps consommé
    if (JourNonTrav = False) then
    begin
      //Dans le cas d'une demi-journée sur une journée complète travaillé
      if (DateDebAbs = DateDebut) and (Nodjm = 1) and (AM = True) and (PM = True) then
      begin
        Nb_j := Nb_j / 2;
        Nb_H := Nb_H / 2
      end;
      if (DateDebAbs = DateFin) and (Nodjp = 0) and (AM = True) and (PM = True) then
      begin
        Nb_j := Nb_j / 2;
        Nb_H := Nb_H / 2
      end;
      //Dans le cas d'une demi-journée sur une journée ouvré ou ouvrable travaillé à demi temps
      if (SSDecompte = False) and (Nb_j = 0.5) and (((AM = True) and (PM = FALSE)) or ((AM = False) and (PM = True))) then Nb_j := 1; //PT68-1 PT96
      CptJour := CptJour + Nb_j;
      CptHeure := CptHeure + Nb_H;
    end;
    DateDebAbs := DateDebAbs + 1;
  until (DateDebAbs > DateFinAbs);


  FreeAndNil(Tob_Standard);
  FreeAndNil(Tob_Semaine);
  //FreeAndNil(Tob_JourFerie);
  {if Tob_Standard<>nil then  Begin Tob_Standard.free; Tob_Standard:=nil; end;
  if Tob_Semaine<>nil then   Begin Tob_Semaine.free; Tob_semaine:=nil; end;
  if Tob_JourFerie<>nil then Begin Tob_JourFerie.free; Tob_JourFerie:=nil; end; }
  Nb_J := CptJour;
  Nb_H := CptHeure;

  { FIN PT60 }

  (** PT60 Mise en commentaire de l'ancien mode de calcul puis supprimé  **)
end;

{ PT8 : Nouvelle procédure }
Procedure CalculNbJourAbsence(DateDebut, DateFin: TDatetime; Salarie, Etablissement, TypeMvt: string; T_MotifAbs : Tob; var Nb_J, Nb_H: double; Nodjm: integer = 0; Nodjp: integer = 1);
Var Tob_Ferie : Tob;
    CalCivil, AvecFerie, DcptOuvre, DcptOuvrable : Boolean;
Begin
  Nb_J := 0;
  Nb_H := 0;
  CalCivil     := False;
  AvecFerie    := True;
  DcptOuvre    := False;
  DcptOuvrable := False;
  if Assigned(T_MotifAbs) then
     Begin
     CalCivil     := (T_MotifAbs.GetValue('PMA_CALENDCIVIL') = 'X');
     AvecFerie    := (T_MotifAbs.GetValue('PMA_SSJOURFERIE') = '-');
     DcptOuvre    := (T_MotifAbs.GetValue('PMA_OUVRES') = 'X');
     DcptOuvrable := (T_MotifAbs.GetValue('PMA_OUVRABLE') = 'X');
     end;
  if AvecFerie then
     Tob_Ferie := ChargeTobFerie(DateDebut,DateFin);
  if CalCivil then
     DecompteCalendrierCivil(DateDebut, DateFin, Tob_Ferie, nb_j, nb_h, NoDjm, NoDjp)
  else if DcptOuvre then
     DecompteOuvre(DateDebut, DateFin, Tob_Ferie, nb_j, nb_h, NoDjm, NoDjp)
  else if DcptOuvrable then
     DecompteOuvrable(DateDebut, DateFin, Tob_Ferie, nb_j, nb_h, NoDjm, NoDjp)
  else
     DecompteCalendrierSalarie(DateDebut,DateFin, Salarie, Etablissement, TypeMvt, Tob_Ferie, Nb_J, Nb_H, Nodjm, Nodjp);

 if Assigned(Tob_Ferie) then FreeAndNil(Tob_Ferie);
End;


{ DEB PT9 }
Function ControleGestionMaximum(Salarie,Typeconge : String; T_MotifAbs : Tob; DateAbs : TDateTime; Nb_j, Nb_H : Double ) : Boolean;
var
  Per : Integer;
  DateDebPer : TDateTime;
  StSql, StType : String;
  Q : TQuery;
  YY,MM,JJ : WORD;
  Duree : Double ;
Begin
  Result := False;
  Duree := 0;
  DateDebPer := Idate1900;
  
  if T_MotifAbs.GetValue('PMA_GESTIONMAXI') = 'X' then
    begin
        if T_MotifAbs.GetValue('PMA_JOURHEURE') = 'JOU' then
            Begin
            StType := 'PCN_JOURS';
            Duree := Nb_j;
            End
        else
         if ((T_MotifAbs.GetValue('PMA_JOURHEURE') = 'HEU') or (T_MotifAbs.GetValue('PMA_JOURHEURE') = 'HOR')) then //PT25
            Begin
            StType := 'PCN_HEURES';
            Duree := Nb_H;
            End;

        if T_MotifAbs.GetValue('PMA_JRSMAXI') < Duree then
          Begin
          Result := True;
          End
        else
        if (T_MotifAbs.GetValue('PMA_TYPEPERMAXI')<>'') AND (T_MotifAbs.GetValue('PMA_PERMAXI') > 0) then
          Begin
          Per := T_MotifAbs.GetValue('PMA_PERMAXI');
          DecodeDate(DateAbs,YY,MM,JJ);
          if T_MotifAbs.GetValue('PMA_TYPEPERMAXI')='PER' then
            Begin
            if MM < Per then YY := YY - 1;
            DateDebPer := EncodeDate(YY,Per,1);
            End
          else
            if T_MotifAbs.GetValue('PMA_TYPEPERMAXI')='GLI' then
               Begin
               DateDebPer := PlusDate(DateAbs,-Per,'M');
               End;
          StSql := 'SELECT SUM('+StType+') DUREE FROM ABSENCESALARIE '+
                   'WHERE PCN_TYPEMVT="ABS" AND PCN_SALARIE="'+Salarie+'" AND PCN_TYPECONGE="'+Typeconge+'" '+
                   'AND PCN_DATEDEBUTABS>="'+USDateTime(DateDebPer)+'" '+
                   'AND PCN_DATEFINABS<"'+USDateTime(DateAbs)+'" AND PCN_ETATPOSTPAIE <> "NAN" '; { PT12 }
          Q := OpenSql(StSql,True);
          If Not Q.Eof then
            Begin
            if T_MotifAbs.GetValue('PMA_JRSMAXI') < (Q.FindField('DUREE').AsFloat + Duree) then
              result := True;
            End;
          Ferme(Q);
          End;
    End;
End;
{ FIN PT9 }


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : rapport entre tablette demi journee et valeur numérique
Suite ........ : matin => 0; Après midi => 1
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure AffecteNodemj(valzone: string; var Nodj: integer);
begin
  if valzone = 'MAT' then Nodj := 0 else Nodj := 1;
end;

{ DEB PT11 }
Function  EstLibre(pDtDebut, pDtFin : TDateTime; pStRes, pStDebut, pStFin : String; var pStAffaire : String): Boolean;
var
  vSt    : String;
  vTob   : Tob;
  vDtDeb : TDateTime;
  vDtFin : TDateTime;
begin
  result := True;
//  if (V_PGI.ModePCL = '1') AND
{$IFNDEF PLUGIN}   // PT19
  if (EstBasePclOptimisee) then exit; // PT17 Cas dossier optimisé PCL
{$ENDIF PLUGIN}   // PT19

  if pStDebut = 'MAT' then
     vDtDeb := pDtDebut + GetParamSocSecur('SO_AFAMDEBUT', '00:00:00')
  else
     vDtDeb := pDtDebut + GetParamSocSecur('SO_AFPMDEBUT', '00:00:00');

  if pStFin = 'MAT' then
     vDtFin := pDtFin + GetParamSocSecur('SO_AFAMDEBUT', '00:00:00')
  else
     vDtFin := pDtFin + GetParamSocSecur('SO_AFPMDEBUT', '00:00:00');

  vSt := 'SELECT APL_AFFAIRE, APL_TYPEPLA, APL_NUMEROLIGNE FROM AFPLANNING,RESSOURCE  ';
  vSt := vSt + ' WHERE ARS_RESSOURCE = APL_RESSOURCE AND ARS_SALARIE = "' + pStRes + '"';
  vSt := vSt + ' AND ((APL_HEUREDEB_PLA <= "' + UsTime(vDtFin) + '")';
  vSt := vSt + ' AND (APL_HEUREFIN_PLA >= "' + UsTime(vDtDeb) + '"))';

  vTob := TOB.Create('AFPLANNING', nil, -1);
  try
    vTob.LoadDetailFromSQL(vSt);
    if vTob.detail.count > 0 then
     begin
     result := False;
     pStAffaire := vTob.detail[0].GetString('APL_AFFAIRE');
     end;
  finally
    FreeAndNil(vTob);
  end;
End;
{ FIN PT11 }


{ DEB PT14 }
Function RendLibAbsence(TypeConge,DebDj,FinDj : string; DD,DF : TDateTime) : string;
var MonFormat : String;
Begin
// PT22  if TypeConge = 'PRI' then Result := 'CP'
  if typeConge = 'PRI' then Result := 'Congés payés'  // PT22 
  else if TypeConge <> '' then Result := RechDom('PGMOTIFABSENCE', Typeconge, False)
  else Result := '';
// Result := RechDom('PGMOTIFABSENCE', TypeConge , False);

 if (DD <> idate1900) and (DF <> idate1900) and (Result <> '') then
 begin
    if DebDj = 'MAT' then DebDj := ' am' else DebDj := ' pm';
    if FinDj = 'MAT' then FinDj := ' am' else FinDj := ' pm';
   // PT22 if Length(Result) > 9 then Result := Copy(Result, 1, 9); //Tronquage du libellé : recup des 9 premiers caractères
   if Length(Result) > 14 then Result := Copy(Result, 1, 14); //   PT22
    If V_PGI.LanguePrinc = 'UK' then MonFormat := 'mmm dd yyyy'
      else MonFormat := 'dd/mm/yy';
    // deb PT22
    //   Result := TraduireMemoire(Result) + ' ' + FormatDateTime(MonFormat, DD) + DebDj + ' '+TraduireMemoire('au')+' '+ FormatDateTime(MonFormat, DF) + FinDj;

    Result := TraduireMemoire(Result) + ' ' + FormatDateTime(MonFormat, DD) + ' '+TraduireMemoire('au')+' '+ FormatDateTime(MonFormat, DF);
    // fin PT22
    Result := Trim(Result);
 end;
End;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : fonction d'initialisation d'un mouvement congés payés
Suite ........ : A appeler avant tout renseignement de nouveau mvt
Mots clefs ... : PAIE;CP
*****************************************************************}
//PT6
procedure InitialiseTobAbsenceSalarie(TCP: tob);
begin
  TCP.Putvalue('PCN_TYPEMVT'             , 'CPA' );
  TCP.Putvalue('PCN_SALARIE'             , ''    );
  TCP.Putvalue('PCN_DATEDEBUT'           , Idate1900 );
  TCP.Putvalue('PCN_DATEFIN'             , Idate1900 );
  TCP.Putvalue('PCN_ORDRE'               , 0     );
  TCP.Putvalue('PCN_MVTPRIS'             , ''    );
  TCP.PutValue('PCN_PERIODECP'           , -1    );
  TCP.PutValue('PCN_PERIODEPY'           , -1    );
  TCP.Putvalue('PCN_TYPECONGE'           , ''    );
  TCP.Putvalue('PCN_SENSABS'             , ''    );
  TCP.Putvalue('PCN_LIBELLE'             , ''    );
  TCP.Putvalue('PCN_GENERECLOTURE'       ,'-'    );
  TCP.Putvalue('PCN_DATESOLDE'           , Idate1900 );
  TCP.Putvalue('PCN_DATEVALIDITE'        , Idate1900 );
  TCP.Putvalue('PCN_DATEDEBUTABS'        , Idate1900 );
  TCP.Putvalue('PCN_DEBUTDJ'             , ''    );
  TCP.Putvalue('PCN_DATEFINABS'          , Idate1900 );
  TCP.Putvalue('PCN_FINDJ'               , ''    );
  TCP.Putvalue('PCN_DATEPAIEMENT'        , Idate1900 );
  TCP.Putvalue('PCN_CODETAPE'            , '...' );
  TCP.Putvalue('PCN_JOURS'               , 0     );
  TCP.Putvalue('PCN_HEURES'              , 0     );
  TCP.Putvalue('PCN_BASE'                , 0     );
  TCP.Putvalue('PCN_NBREMOIS'            , 0     );
  TCP.Putvalue('PCN_CODERGRPT'           , 0     );
  TCP.PutValue('PCN_MVTDUPLIQUE'         , '-'   );
  TCP.Putvalue('PCN_ABSENCE'             , 0     );
  TCP.Putvalue('PCN_MODIFABSENCE'        , '-'   );
  TCP.Putvalue('PCN_APAYES'              , 0     );
  TCP.Putvalue('PCN_VALOX'               , 0     );
  TCP.Putvalue('PCN_VALOMS'              , 0     );
  TCP.Putvalue('PCN_VALORETENUE'         , 0     );
  TCP.Putvalue('PCN_VALOMANUELLE'        , 0     );
  TCP.Putvalue('PCN_MODIFVALO'           , '-'   );
  TCP.Putvalue('PCN_PERIODEPAIE'         ,''     );
  TCP.Putvalue('PCN_ETABLISSEMENT'       , ''    );
  TCP.Putvalue('PCN_TRAVAILN1'           , ''    );
  TCP.Putvalue('PCN_TRAVAILN2'           , ''    );
  TCP.Putvalue('PCN_TRAVAILN3'           , ''    );
  TCP.Putvalue('PCN_TRAVAILN4'           , ''    );
  TCP.Putvalue('PCN_CODESTAT'            , ''    );
  TCP.Putvalue('PCN_SAISIEDEPORTEE'      , '-'   );
  TCP.Putvalue('PCN_VALIDSALARIE'        , ''    );
  TCP.Putvalue('PCN_VALIDRESP'           , ''    );
  TCP.Putvalue('PCN_EXPORTOK'            , ''    );
  TCP.Putvalue('PCN_LIBCOMPL1'           , ''    );
  TCP.Putvalue('PCN_LIBCOMPL2'           , ''    );
  TCP.Putvalue('PCN_VALIDABSENCE'        , ''    );
  TCP.Putvalue('PCN_OKFRACTION'          , ''    );
  TCP.Putvalue('PCN_NBJCARENCE'          , 0     );
  TCP.Putvalue('PCN_NBJCALEND'           , 0     );
  TCP.Putvalue('PCN_NBJIJSS'             , 0     );
  TCP.Putvalue('PCN_IJSSSOLDEE'          , '-'   );
  TCP.Putvalue('PCN_GESTIONIJSS'         , '-'   );
  TCP.Putvalue('PCN_ETATPOSTPAIE'        , 'VAL' );
end;

Procedure InitialiseTobAbsMotif ( T_Abs, T_Param,  T_Sal  :  Tob; ValResp,Exp,SaisiePar,SaisieVia : String);
Var No_Ordre : integer;
Begin
if not Assigned(T_Sal) then exit;
if not Assigned(T_Param) then exit;
if not Assigned(T_Abs) then exit; 
 //T_Pris:=TOB.Create('ABSENCESALARIE',Tob_CpPris,-1);
  InitialiseTobAbsenceSalarie(T_Abs);
  if T_Param.GetValue('PCN_TYPECONGE')='PRI' then
     Begin
     T_Abs.PutValue('PCN_TYPEMVT'       , 'CPA');
     T_Abs.PutValue('PCN_MVTPRIS'       , 'PRI');
     No_ordre := IncrementeSeqNoOrdre('CPA',T_Sal.GetValue('PSA_SALARIE'));
     End
   else
     Begin
     T_Abs.PutValue('PCN_TYPEMVT'       , 'ABS');
     T_Abs.PutValue('PCN_MVTPRIS'       , '');
     No_ordre := IncrementeSeqNoOrdre('',T_Sal.GetValue('PSA_SALARIE'));
     End;
  T_Abs.PutValue('PCN_SALARIE'       , T_Sal.GetValue('PSA_SALARIE'));
  T_Abs.PutValue('PCN_ORDRE'         , No_Ordre);
  T_Abs.PutValue('PCN_DATEDEBUTABS'  , T_Param.GetValue('PCN_DATEDEBUTABS'));
  T_Abs.PutValue('PCN_DEBUTDJ'       , T_Param.GetValue('PCN_DEBUTDJ'));
  T_Abs.PutValue('PCN_DATEFINABS'    , T_Param.GetValue('PCN_DATEFINABS'));
  T_Abs.PutValue('PCN_FINDJ'         , T_Param.GetValue('PCN_FINDJ'));
  T_Abs.PutValue('PCN_TYPECONGE'     , T_Param.GetValue('PCN_TYPECONGE'));
  T_Abs.PutValue('PCN_SENSABS'       , '-');
  T_Abs.PutValue('PCN_JOURS'         , T_Param.GetValue('PCN_JOURS'));
  T_Abs.PutValue('PCN_HEURES'        , T_Param.GetValue('PCN_HEURES'));
  T_Abs.PutValue('PCN_LIBELLE'       , T_Param.getvalue('PCN_LIBELLE'));
  T_Abs.PutValue('PCN_LIBCOMPL1'     , T_Param.getvalue('PCN_LIBCOMPL1'));  //PT24
  T_Abs.PutValue('PCN_CODERGRPT'     , No_Ordre);
  T_Abs.PutValue('PCN_DATEVALIDITE'  , StrToDate(T_Param.GetValue('PCN_DATEVALIDITE')));
  T_Abs.PutValue('PCN_PERIODECP'     , 0);
  T_Abs.PutValue('PCN_ETABLISSEMENT' , T_Sal.GetValue('PSA_ETABLISSEMENT'));
  T_Abs.PutValue('PCN_TRAVAILN1'     , T_Sal.GetValue('PSA_TRAVAILN1'));
  T_Abs.PutValue('PCN_TRAVAILN2'     , T_Sal.GetValue('PSA_TRAVAILN2'));
  T_Abs.PutValue('PCN_TRAVAILN3'     , T_Sal.GetValue('PSA_TRAVAILN3'));
  T_Abs.PutValue('PCN_TRAVAILN4'     , T_Sal.GetValue('PSA_TRAVAILN4'));
  T_Abs.PutValue('PCN_CODESTAT'      , T_Sal.GetValue('PSA_CODESTAT'));
  T_Abs.PutValue('PCN_CONFIDENTIEL'  , T_Sal.GetValue('PSA_CONFIDENTIEL'));




  T_Abs.PutValue('PCN_VALIDRESP'     , ValResp);
  T_Abs.PutValue('PCN_EXPORTOK'      , Exp);
  if GetParamSocSecur('SO_PGECABBASEDEPORTE',False) then T_Abs.PutValue('PCN_SAISIEDEPORTEE', 'X');
  T_Abs.PutValue('PCN_VALIDSALARIE'  , SaisiePar);
  T_Abs.PutValue('PCN_ETATPOSTPAIE'  , 'VAL');
  T_Abs.PutValue('PCN_MVTORIGINE'    , SaisieVia);

End;



//------------------------------------------------------------------------------
// fonction d'incrément du numéro d'ordre pour la table
// ABSENCESALARIE
//------------------------------------------------------------------------------

function IncrementeSeqNoOrdre(cle_type, cle_sal: string): Integer;
var
  q: TQuery;
  StWhere: string;
begin
  result := 1;
  if cle_type <> '' then StWhere := 'AND PCN_TYPEMVT = "' + cle_type + '" ' else StWhere := '';
  q := OpenSQL('select max (PCN_ORDRE) as maxno from ABSENCESALARIE ' +
    ' where PCN_SALARIE = "' + cle_sal + '" ' + StWhere, TRUE);
  //cle :  PCN_TYPEMVT,PCN_SALARIE,PCN_DATEDEBUT,PCN_DATEFIN,PCN_ORDRE
  if not q.eof then
{$IFDEF EAGLCLIENT}
    if q.Fields[0].AsInteger <> 0 then
{$ELSE}
    if not q.Fields[0].IsNull then
{$ENDIF}
      result := ((q.Fields[0]).AsInteger + 1);

  Ferme(q);
end;



function ControleAssezRestant(T : Tob; {Salarie,} TYPECONGE, Action: string; Ejours, EJoursPris : double): boolean;
var
  SumPoses, SumAcquis: double;
 // T: Tob;
begin
//  T := nil;
  result := false;
{  if Tob_Recapitulatif <> nil then
    T := Tob_Recapitulatif.FindFirst(['PRS_SALARIE'], [Salarie], False);}
  if T <> nil then
  begin
    if TypeConge = 'PRI' then
    begin
      SumPoses := T.GetValue('PRS_PRIVALIDE') + T.GetValue('PRS_PRIATTENTE');
      SumAcquis := T.GetValue('PRS_RESTN1') + T.GetValue('PRS_RESTN');
    end
    else
    begin
      SumPoses := T.GetValue('PRS_RTTVALIDE') + T.GetValue('PRS_RTTATTENTE');
      SumAcquis := T.GetValue('PRS_CUMRTTREST');
    end;
    if Action = 'MODIFICATION' then SumPoses := SumPoses - EJoursPris;
    if (SumPoses + Ejours) > SumAcquis then
    begin
      if TypeConge = 'PRI' then
        PGIInfo('Attention! Le solde des congés acquis est dépassé.', 'Saisie d''absence')
      else
        PGIInfo('Attention! Le solde des RTT acquis est dépassé.', 'Saisie d''absence');
      result := true;
    end;
  end;
end; 


function  PgGetClauseAbsAControler(T_MotifAbs,LcTob_MotifAbs : tob; TypeConge : String) : string;
var i : integer;
    T : Tob;
    ListeAbs : String;
Begin
   Result := '';
   { Contrôle présence salarié si option non coché }
   if assigned(T_MotifAbs) AND (T_MotifAbs.GetValue('PMA_CONTROLMOTIF') = '-') then
        Begin
        ListeAbs := '';
        if assigned(LcTob_MotifAbs) then
          Begin
          For i:=0 to LcTob_MotifAbs.detail.count-1 do
            Begin
            T := LcTob_MotifAbs.detail[i];
            if (T.GetValue('PMA_CONTROLMOTIF') = '-') and (T.GetValue('PMA_MOTIFABSENCE')<>'PRI') then
               ListeAbs := ListeAbs + '"'+T.GetValue('PMA_MOTIFABSENCE')+'",';
            End;
            ListeAbs := Copy(ListeAbs,1,length(ListeAbs)-1);
          End;
        if ListeAbs = '' then
           Result :='AND (PCN_TYPECONGE = "PRI" AND PCN_TYPEMVT="CPA") '
        else
        Result :='AND ((PCN_TYPECONGE = "PRI" AND PCN_TYPEMVT="CPA") '+
           'OR (PCN_TYPEMVT = "ABS" AND PCN_SENSABS="-" AND PCN_TYPECONGE IN ('+ListeAbs+'))) ';
        end
  else
         { Si pas de contrôle présence salarié alors contrôle au moins pas de doublon d'absence identique }
        Begin
        if (TypeConge='PRI') then Result := 'AND PCN_TYPEMVT="CPA" '
                                                else Result := 'AND PCN_TYPEMVT="ABS" ';
        Result := Result + 'AND PCN_TYPECONGE="'+TypeConge+'" AND PCN_SENSABS="-" ';
        End;

End;

Function  ChargeTobSalarie (Sal : String) : Tob;
Var Q : TQuery;
    St: String;
    Temp : Tob;
Begin
Temp := TOB.Create('Le salarié',nil,-1);
St := 'SELECT * FROM SALARIES, ETABCOMPL '+
'WHERE PSA_ETABLISSEMENT = ETB_ETABLISSEMENT AND PSA_SALARIE="'+Sal+'"';
Q := OpenSql(St,True);
if not Q.eof then
   Begin
   Temp.LoadDetailDB('SAL_ARIES','','',Q,False);
   Result := Temp.detail[0];
   End
else
   Result := nil;
Ferme(Q);
//FreeAndNil(Temp);
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 18/03/2008 / PT28
Modifié le ... :   /  /
Description .. : Idem ChargeTobSalarie en Interimaire
Mots clefs ... :
*****************************************************************}
Function  ChargeTobInterimaire (Sal : String) : Tob;
Var Q : TQuery;
    St: String;
    Temp : Tob;
Begin
	Temp := TOB.Create('Le salarié',nil,-1);
	St := 'SELECT * FROM INTERIMAIRES WHERE PSI_INTERIMAIRE="'+Sal+'"';
	Q := OpenSql(St,True);
	if not Q.eof then
	   Begin
	   Temp.LoadDetailDB('SAL_ARIES','','',Q,False);
	   Result := Temp.detail[0];
	   End
	else
	   Result := nil;
	Ferme(Q);
End;

procedure ChargeInfoRespSalMail(Sal, FicPrec, TypeSal : String; Var Resp,NomResp,MailResp,MailSal : string; Var OKMailEnvoie : Boolean; Var MailDate : TDateTime);
Var Q,Q1 : TQuery;
St : String;
Begin
    Resp := '';
    NomResp := '';
    MailResp := '';
    OkMailEnvoie := False;
    MailDate := Idate1900;
    MailSal := '';
    //**SetControltext('NOMRESP', '');
    if (FicPrec = 'SAL') then
    begin
      Q := OpenSql('SELECT PSE_SALARIE,PSE_EMAILENVOYE,PSE_EMAILPROF,PSE_EMAILDATE,' +
        'PSA_LIBELLE,PSA_PRENOM FROM DEPORTSAL ' +
        'LEFT JOIN SALARIES ON PSE_SALARIE=PSA_SALARIE ' + 
        'WHERE PSE_SALARIE IN (SELECT RESP.PSE_RESPONSABS FROM DEPORTSAL RESP ' +
        'WHERE RESP.PSE_SALARIE="' + Sal + '")', True);
      if not Q.eof then
      begin
        Resp := Q.FindField('PSE_SALARIE').AsString;
        NomResp := Q.FindField('PSA_LIBELLE').AsString + ' ' + Q.FindField('PSA_PRENOM').AsString;
        MailResp := Q.FindField('PSE_EMAILPROF').AsString;
        OkMailEnvoie := (Q.FindField('PSE_EMAILENVOYE').AsString = 'X');
        MailDate := Q.FindField('PSE_EMAILDATE').AsDateTime; 
        if NomResp = ' ' then
        begin
          St := 'SELECT PSI_LIBELLE,PSI_PRENOM FROM INTERIMAIRES WHERE PSI_INTERIMAIRE ="' + Resp + '"';
          Q1 := OpenSql(St, True);
          if not Q1.eof then
            NomResp := Q1.FindField('PSI_LIBELLE').AsString + ' ' + Q1.FindField('PSI_PRENOM').AsString;
          Ferme(Q1);
        end;
        //***SetControltext('NOMRESP', 'Par ' + NomResp);
      end;
      Ferme(Q);
    end
    else
      if (FicPrec = 'RESP') or (FicPrec = 'ADM') then
    begin
      if (TypeSal = 'PSA') OR (TypeSal = 'SAL') OR (TypeSal = '')then
        St := 'SELECT PSE_SALARIE,PSE_RESPONSABS,PSE_EMAILENVOYE,PSE_EMAILPROF,PSE_EMAILDATE,' +
        'PSA_LIBELLE,PSA_PRENOM FROM DEPORTSAL ' +
        'LEFT JOIN SALARIES ON PSE_RESPONSABS=PSA_SALARIE ' +
        'WHERE PSE_SALARIE="' + sal + '"'
      else
      if TypeSal = 'INT' then
        St := 'SELECT PSE_SALARIE,PSE_RESPONSABS,PSE_EMAILENVOYE,PSE_EMAILPROF,PSE_EMAILDATE,' +
        'PSI_LIBELLE,PSI_PRENOM FROM DEPORTSAL ' +
        'LEFT JOIN INTERIMAIRES ON PSE_RESPONSABS=PSI_INTERIMAIRE ' +
        'WHERE PSE_SALARIE="' + Sal + '"';
      Q := OpenSql(St, True);
      if not Q.eof then
      begin
        Resp := Q.FindField('PSE_RESPONSABS').AsString;
        if (TypeSal = 'PSA') OR (TypeSal = 'SAL')then
        NomResp := Q.FindField('PSA_LIBELLE').AsString + ' ' + Q.FindField('PSA_PRENOM').AsString
        else
        if TypeSal = 'INT' then
          NomResp := Q.FindField('PSI_LIBELLE').AsString + ' ' + Q.FindField('PSI_PRENOM').AsString;
        if Trim(NomResp) = '' then
        begin
          St := 'SELECT PSI_LIBELLE,PSI_PRENOM FROM INTERIMAIRES WHERE PSI_INTERIMAIRE ="' + Resp + '"';
          Q1 := OpenSql(St, True);
          if not Q1.eof then
            NomResp := Q1.FindField('PSI_LIBELLE').AsString + ' ' + Q1.FindField('PSI_PRENOM').AsString;
          Ferme(Q1);
        end;
        MailSal := Q.FindField('PSE_EMAILPROF').AsString;
        //*** SetControltext('NOMRESP', 'Par ' + NomResp);
      end;
      Ferme(Q);
    end;
End;


Function  ChargeTob_Recapitulatif( Sal : String) : tob;
var
  Q: TQuery;
  Tob_Recapitulatif : Tob;
begin
  Q := Opensql('SELECT * FROM RECAPSALARIES WHERE PRS_SALARIE = "' + Sal + '"', true);
  if not Q.Eof then
  begin
    Tob_Recapitulatif := Tob.Create('Récapitulatif Salarié', nil, -1);
    Tob_Recapitulatif.LoadDetailDB('Récapitulatif Salarié', '', '', Q, False);
    Result := Tob_Recapitulatif.FindFirst(['PRS_SALARIE'], [Sal], False);
  end
  else
  begin
//    Tob_Recapitulatif := nil;
{PT21
    FreeAndNil(Tob_Recapitulatif);
}
    Result := nil;
  end;
  Ferme(Q);
End;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/10/2002
Modifié le ... :   /  /
Description .. : Calcul des compteur en cours de la table RECAPSALARIES
Mots clefs ... : PAIE;ABSENCE;ECONGES
*****************************************************************}
procedure PGExeCalculRecapAbsEnCours(Salarie: string);
var
//  St,StCumul, StDate : string;
  St,StCumul : string;
  Q: TQuery;
  Tob_MvtEncours, Tob_RecapSalarie, Tmvt, TRecap: TOB;
  i: integer;
  DTemp : TDateTime;
begin
  Tob_RecapSalarie := nil;
  Tob_MvtEncours := nil;
  try
    beginTrans;
    //InitMoveProgressForm(nil, 'Calcul des compteurs en cours', 'Veuillez patienter SVP ...', 1, FALSE, TRUE);
    //Remise à zero des compteurs salariés
    if salarie <> '' then St := ' WHERE PRS_SALARIE="' + salarie + '"' else St := '';
    ExecuteSql('UPDATE RECAPSALARIES SET PRS_PRIVALIDE=0,PRS_PRIATTENTE=0,' +
      'PRS_RTTVALIDE=0,PRS_RTTATTENTE=0 ' + St);

    //Charge la tob des recapsalaries
      Tob_RecapSalarie := TOb.Create('RECAPSALARIES', nil, -1);
      if Salarie <> '' then
        Tob_RecapSalarie.LoadDetailDB('RECAPSALARIES', '"' + Salarie + '"', '', nil, False)
      else
        Tob_RecapSalarie.LoadDetailDB('RECAPSALARIES', '', '', nil, False);

    //Réinitialisation des données
    For i :=0 to Tob_RecapSalarie.detail.count-1 do
      Begin
      Tob_RecapSalarie.detail[i].PutValue('PRS_PRIVALIDE'      , 0 );
      Tob_RecapSalarie.detail[i].PutValue('PRS_PRIATTENTE'     , 0 );
      Tob_RecapSalarie.detail[i].PutValue('PRS_RTTVALIDE'      , 0 );
      Tob_RecapSalarie.detail[i].PutValue('PRS_RTTATTENTE'     , 0 );
      Tob_RecapSalarie.detail[i].PutValue('PRS_ACQUISCPSUIV'   , 0 );
      Tob_RecapSalarie.detail[i].PutValue('PRS_ACQUISRTTSUIV'  , 0 );
      End;

    { Calcul des compteurs récap }
    if salarie <> '' then St := ' AND PCN_SALARIE="' + salarie + '"' else St := '';
    St := 'SELECT PCN_TYPEMVT,PCN_SALARIE,PCN_TYPECONGE,PCN_JOURS,PCN_VALIDRESP FROM ABSENCESALARIE ' +
      'LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI## PMA_MOTIFABSENCE=PCN_TYPECONGE ' +
      'WHERE ((PCN_TYPEMVT="ABS" AND PMA_TYPEABS="RTT") OR PCN_TYPECONGE="PRI") ' +
      'AND PCN_ETATPOSTPAIE <> "NAN" '+
      'AND PMA_MOTIFEAGL="X" AND (PCN_EXPORTOK<>"X" OR PCN_CODETAPE="...") ' + St +
      'ORDER BY PCN_SALARIE,PCN_TYPEMVT';
    Q := OpenSql(st, True);
    if not Q.Eof then
    begin
      //Charger la tob des mvts en cours
      Tob_MvtEncours := TOb.Create('Les mvts en cours', nil, -1);
      Tob_MvtEncours.LoadDetailDB('Les mvts en cours', '', '', Q, False);
      Ferme(Q);

      //Boucle sur mvts en cours et alimente la table RECAPSALARIE en fonction de l'état de validation
      for i := 0 to Tob_MvtEncours.detail.count - 1 do
      begin
        Tmvt := Tob_MvtEncours.detail[i];
        //MoveCurProgressForm('Salarié : ' + Tmvt.GetValue('PCN_SALARIE'));
        TRecap := Tob_RecapSalarie.FindFirst(['PRS_SALARIE'], [Tmvt.GetValue('PCN_SALARIE')], False);
        if TRecap <> nil then
        begin
          if (Tmvt.GetValue('PCN_TYPECONGE') = 'PRI') and (Tmvt.GetValue('PCN_VALIDRESP') = 'VAL') then
            TRecap.PutValue('PRS_PRIVALIDE', TRecap.GetValue('PRS_PRIVALIDE') + Tmvt.GetValue('PCN_JOURS'))
          else
            if (Tmvt.GetValue('PCN_TYPECONGE') = 'PRI') and (Tmvt.GetValue('PCN_VALIDRESP') = 'ATT') then
            TRecap.PutValue('PRS_PRIATTENTE', TRecap.GetValue('PRS_PRIATTENTE') + Tmvt.GetValue('PCN_JOURS'))
          else
            if (Tmvt.GetValue('PCN_TYPECONGE') <> 'PRI') and (Tmvt.GetValue('PCN_VALIDRESP') = 'VAL') then
            TRecap.PutValue('PRS_RTTVALIDE', TRecap.GetValue('PRS_RTTVALIDE') + Tmvt.GetValue('PCN_JOURS'))
          else
            if (Tmvt.GetValue('PCN_TYPECONGE') <> 'PRI') and (Tmvt.GetValue('PCN_VALIDRESP') = 'ATT') then
            TRecap.PutValue('PRS_RTTATTENTE', TRecap.GetValue('PRS_RTTATTENTE') + Tmvt.GetValue('PCN_JOURS'));
        end;
      end;
    end
    else Ferme(Q);
    if Tob_MvtEncours <> nil then Tob_MvtEncours.free;


    { Calcul des acquis CP à venir }
    if GetParamSocSecur('SO_PGCONGES',False) then
    Begin
    if salarie <> '' then St := ' AND PSA_SALARIE="' + salarie + '"' else St := '';
    St := 'SELECT PSA_SALARIE,PSA_NBREACQUISCP FROM SALARIES '+
          'WHERE PSA_CONGESPAYES="X" '+St;
      Q := OpenSql(st, True);
      if not Q.eof then
        Begin
        //Charger la tob des mvts en cours
        Tob_MvtEncours := TOb.Create('Les mvts en cours', nil, -1);
        Tob_MvtEncours.LoadDetailDB('Les mvts en cours', '', '', Q, False);
        Ferme(Q);
        //Boucle sur mvts en cours et alimente la table RECAPSALARIE
        for i := 0 to Tob_MvtEncours.detail.count - 1 do
          begin
          Tmvt := Tob_MvtEncours.detail[i];
          TRecap := Tob_RecapSalarie.FindFirst(['PRS_SALARIE'], [Tmvt.GetValue('PSA_SALARIE')], False);
          if TRecap <> nil then
             TRecap.PutValue('PRS_ACQUISCPSUIV', Tmvt.GetValue('PSA_NBREACQUISCP'));
          End;
        End
      else Ferme(Q);
    End;       

    { Calcul des jours rtt à venir }
    if GetParamSocSecur('SO_PGECABDATEINTEGRATION',idate1900) <> idate1900 then
      Begin
      DTemp := PlusDate(GetParamSocSecur('SO_PGECABDATEINTEGRATION',idate1900),-1,'M') + 1 ;
      StCumul := '';
      St := GetParamSocSecur('SO_PGRTTACQUIS','');
      while St <> '' do
      StCumul := StCumul + ' PHC_CUMULPAIE="' + ReadTokenSt(St) + '" OR';
      if StCumul <> '' then StCumul := ' AND (' + Copy(StCumul, 1, Length(StCumul) - 2) + ') '
      else StCumul := ' AND PHC_CUMULPAIE="46" ';
      if salarie <> '' then St := ' AND PHC_SALARIE="' + salarie + '"' else St := '';

      St := 'SELECT PHC_SALARIE,PHC_MONTANT FROM HISTOCUMSAL '+
            'WHERE PHC_DATEFIN>="'+USDateTime(DTemp)+'" '+
            'AND PHC_DATEFIN<="'+USDateTime(GetParamSocSecur('SO_PGECABDATEINTEGRATION',idate1900))+'" '+
            St+StCumul;
      Q := OpenSql(st, True);
      if not Q.eof then
        Begin
        //Charger la tob des mvts en cours
        Tob_MvtEncours := TOb.Create('Les mvts en cours', nil, -1);
        Tob_MvtEncours.LoadDetailDB('Les mvts en cours', '', '', Q, False);
        Ferme(Q);
        //Boucle sur mvts en cours et alimente la table RECAPSALARIE
        for i := 0 to Tob_MvtEncours.detail.count - 1 do
          begin
          Tmvt := Tob_MvtEncours.detail[i];
          TRecap := Tob_RecapSalarie.FindFirst(['PRS_SALARIE'], [Tmvt.GetValue('PHC_SALARIE')], False);
          if TRecap <> nil then
             TRecap.PutValue('PRS_ACQUISRTTSUIV', Tmvt.GetValue('PHC_MONTANT'));
          End;
        End
      else Ferme(Q);
      End;  

    { MAJ De la Tob }
    if Tob_RecapSalarie <> nil then
      if Tob_RecapSalarie.IsOneModifie then
        for i := 0 to Tob_RecapSalarie.detail.count - 1 do
        begin
          if Tob_RecapSalarie.Detail[i].Modifie then
            Tob_RecapSalarie.Detail[i].UpdateDB;
        end;
    CommitTrans;
  except
    Rollback;
  end;
  //FiniMoveProgressForm;


  if Tob_RecapSalarie <> nil then Tob_RecapSalarie.free;
end;




Procedure PgSupprRecapitulatif(Tob_motifAbs : Tob; Sal, TypeConge : String; EJoursPris : Double; Var Error : integer; Var Retour : string);
Var
  T_MotifAbs : Tob;
Begin
  Retour := '';
  if (TypeConge = 'PRI') then
    begin
      try
        BeginTrans;
        ExecuteSql('UPDATE RECAPSALARIES SET PRS_PRIATTENTE=PRS_PRIATTENTE-' + StrfPoint(EJoursPris) +
          ' WHERE PRS_SALARIE="' + Sal + '"');
        CommitTrans;
      except
        Rollback;
        Error := 32;
        //PGIBox('Echec lors du calcul du récapitulatif.Suppression annulée', Ecran.caption);
      end;
    end
    else
    begin
      if assigned(Tob_motifAbs) then
        Begin
        T_MotifAbs := Tob_MotifAbs.FindFirst(['PMA_MOTIFABSENCE'], [TypeConge], False);
        if assigned(T_MotifAbs) then
          if (T_MotifAbs.GetValue('PMA_TYPEABS') = 'RTT') then
        begin
          try
            BeginTrans;
            ExecuteSql('UPDATE RECAPSALARIES SET PRS_RTTATTENTE=PRS_RTTATTENTE-' + StrfPoint(EJoursPris) +
              ' WHERE PRS_SALARIE="' + Sal + '"');
            CommitTrans;
            Retour := 'SUPPR';
          except
            Rollback;
            Error := 32;
            //PGIBox('Echec lors du calcul du récapitulatif.Suppression annulée.', Ecran.caption);
          end;
        end;
        End;
    end;

End;

function PGEnvoieMailCollaborateur(Titre, Mail: string; Texte: TStringList): Boolean; 
var
  Body: HTStringList;
  I: integer;
begin
  Result := True;
  Body := HTStringList.Create;
  Body.Add('');
  Body.Add('');
  for i := 0 to Texte.Count - 1 do
  begin
    Body.Add(Texte.Strings[i]);
    Body.Add('');
  end;
  Body.Add('CEGID PAIE PGI - Module ECongés');
  try
     {$IFDEF MODESMTP}
     {$ELSE}
     MailOl.SendMail(Titre, Mail, '', Body, '', TRUE);
     {$ENDIF}
  except
    on E: Exception do
    begin
      PGiBox(E.Message, 'Econgés');
      Result := False;
    end;
  end;
  Body.Free;
end;


Procedure PgControleTobAbsFirst (T_Abs, TSal : TOB;  {Origine, Action, ficPrec,TypeSal : String;} Var Error : Integer; Var NomControl, ComplMessage : String)  ;
Begin
Error := 0;
if not assigned(T_Abs) then exit;
if not assigned(TSal) then exit;
     { Testons l'affectation du salarié }
     if (T_Abs.GetValue('PCN_SALARIE') = '') or (TSal.GetValue('PSA_LIBELLE') = '') then
        begin
        Error := 26;
        NomControl := 'PCN_SALARIE';
        exit;
        end;
     { Testons l'affectation du type de mvt }
     if T_Abs.GetValue('PCN_TYPEMVT') = '' then
        begin
        If T_Abs.GetValue('PCN_TYPECONGE') = 'PRI' then
           T_Abs.PutValue('PCN_TYPEMVT','CPA')
        else
           T_Abs.PutValue('PCN_TYPEMVT','ABS');
        end;
     { Testons la saisie du type de congé }
     if T_Abs.GetValue('PCN_TYPECONGE') = '' then
        begin
        Error := 4;
        NomControl := 'PCN_TYPECONGE';
        exit;
        end;
     { Testons la saisie du début de demi-journée d'absence }
     if T_Abs.GetValue('PCN_DEBUTDJ') = '' then
        begin
        Error := 11;
        NomControl := 'PCN_DEBUTDJ';
        exit;
        end;
     { Testons la saisie de fin de demi-journée d'absence }
     if T_Abs.GetValue('PCN_FINDJ') = '' then
        begin
        Error := 12;
        NomControl := 'PCN_FINDJ';
        exit;
        end;
     { Testons la saisie des dates d'absence }
{$IFDEF EAGLSERVER}
ddwriteln('1');
{$ENDIF EAGLSERVER}
     if (T_Abs.GetValue('PCN_DATEDEBUTABS') = idate1900) then
        begin
        Error := 11;
        NomControl := 'PCN_DATEDEBUTABS';
        exit;
        end;
{$IFDEF EAGLSERVER}
ddwriteln('2');
{$ENDIF EAGLSERVER}
     if (IsValiddate(DateTimeToStr(T_Abs.GetValue('PCN_DATEDEBUTABS')))= False) then
        begin
        Error := 11;
        NomControl := 'PCN_DATEDEBUTABS';
        exit;
        end;
{$IFDEF EAGLSERVER}
ddwriteln('3');
{$ENDIF EAGLSERVER}
     if (T_Abs.GetValue('PCN_DATEFINABS') = idate1900) OR (isValiddate(T_Abs.GetValue('PCN_DATEFINABS'))= False) then
        begin
        Error := 12;
        NomControl := 'PCN_DATEFINABS';
        exit;
        end;
     if (T_Abs.GetValue('PCN_DATEVALIDITE') = idate1900) OR (isValiddate(T_Abs.GetValue('PCN_DATEVALIDITE'))= False) then
        begin
        Error := 13;
        NomControl := 'PCN_DATEVALIDITE';
        exit;
        end;
     { Testons la cohérence de la saisie des périodes }
     if T_Abs.GetValue('PCN_DATEDEBUTABS') > T_Abs.GetValue('PCN_DATEFINABS') then
        Begin
        Error := 1;
        NomControl := 'PCN_DATEFINABS';
        exit;
        End;
     If (T_Abs.GetValue('PCN_DEBUTDJ') = 'PAM') AND (T_Abs.GetValue('PCN_FINDJ') = 'MAT') AND (T_Abs.GetValue('PCN_DATEFINABS') = T_Abs.GetValue('PCN_DATEDEBUTABS')) then
        Begin
        Error := 31;
        NomControl := 'PCN_FINDJ';
        exit;
        End;
     { Testons si le nombre de jour est renseigné }
     if (T_Abs.GetValue('PCN_TYPEMVT') <> 'CPA') AND (T_Abs.GetValue('PCN_TYPECONGE') <> 'AJU') then
       If (IsNumeric(T_Abs.GetValue('PCN_JOURS')) AND (T_Abs.GetValue('PCN_JOURS')<> '...') )then
         Begin
         If (T_Abs.GetValue('PCN_JOURS') <= 0) then
           Begin
           Error := 3; //Le nombre de jours doit être positif' ;
           NomControl := 'PCN_JOURS';
           exit;
           End;
         End
       else
           Begin
           Error := 3; //Le nombre de jours doit être positif' ;
           NomControl := 'PCN_JOURS';
           exit;
           End;
     { Testons la saisie du libellé }
     if (T_Abs.GetValue('PCN_LIBELLE') = '') then
       T_Abs.PutValue('PCN_LIBELLE',RendLibAbsence(T_Abs.GetValue('PCN_TYPECONGE'),T_Abs.GetValue('PCN_DEBUTDJ'),T_Abs.GetValue('PCN_FINDJ'),T_Abs.GetValue('PCN_DATEDEBUTABS'),T_Abs.GetValue('PCN_DATEFINABS')));
     { Testons la gestion des CP dans le cas d'une saisie CP }
     if T_Abs.GetValue('PCN_TYPEMVT') = 'CPA' then
        begin
        if GetParamSocSecur('SO_PGCONGES',False)  = False then
           begin
           Error := 24;
           NomControl := 'PCN_TYPECONGE';
           exit;
           end
        else
            if (TSal.GetValue('ETB_CONGESPAYES') = '-')  then
              begin
              Error := 25;
              NomControl := 'PCN_TYPECONGE';
              exit;
              end
            else
              if (TSal.GetValue('PSA_CONGESPAYES')='-') then
              begin
              Error := 29;
              NomControl := 'PCN_TYPECONGE';
              exit;
              end;
        end;
End;



Procedure PgControleTobAbsSecond (T_Abs, TSal : TOB; Origine, Action, ficPrec,TypeSal : String; Var Error : Integer; Var NomControl, ComplMessage : String)  ;
Var
LcTob_motifabs,LcT_MotifAbs                                       : Tob;
DateD, DateF, MailDate                                            : TDateTime;
FinDj, DebutDj, StDate, St, pStAffaire, StTerme                   : String;
YYD,MMD,YYF,MMF,JJ                                                : Word;
Q                                                                 : TQuery;
Body                                                              : TStringList;
Resp,NomResp,MailResp,MailSal                                     : String;
TypeRtt,OkMailEnvoie, IsMailRefOrAnu                              : Boolean;
Begin
 LcTob_motifabs := tob.create('tob_virtuelle', nil, -1);
 LcTob_motifabs.loaddetaildb('MOTIFABSENCE', '', 'PMA_MOTIFABSENCE', nil, False);
 LcT_MotifAbs := LcTob_motifabs.FindFirst(['PMA_MOTIFABSENCE'], [T_Abs.GetValue('PCN_TYPECONGE')], False);

 { Testons la gestion maximun sur le nombre de jour ou d'heure }
   if (T_Abs.GetValue('PCN_TYPEMVT') = 'ABS') then
     begin
     If ControleGestionMaximum(T_Abs.GetValue('PCN_SALARIE'),T_Abs.GetValue('PCN_TYPECONGE'),LcT_MotifAbs,
                T_Abs.GetValue('PCN_DATEDEBUTABS'),T_Abs.GetValue('PCN_JOURS'),T_Abs.GetValue('PCN_HEURES')) then
        Begin
        if LcT_MotifAbs.GetValue('PMA_JOURHEURE') = 'JOU' then
           begin
           Error := 17;
           NomControl := 'PCN_JOURS';
           ComplMessage := floattostr(LcT_MotifAbs.GetValue('PMA_JRSMAXI'));
           exit;
          end
        else
        if ((LcT_MotifAbs.GetValue('PMA_JOURHEURE') = 'HEU') OR (LcT_MotifAbs.GetValue('PMA_JOURHEURE') = 'HOR')) then //PT25
           begin
           Error := 18;
           NomControl := 'PCN_HEURES';
           ComplMessage := floattostr(LcT_MotifAbs.GetValue('PMA_JRSMAXI'));
           exit;
           end;
        End;
     end;

  { Testons l'antériorité de la saisie, elle ne doit pas dépasser un mois }
    IF (Origine = 'E') then
      Begin
      if Assigned(LcT_MotifAbs) then TypeRtt := (LcT_MotifAbs.GetValue('PMA_TYPEABS') = 'RTT')
      else TypeRtt := False;
      if ((T_Abs.GetValue('PCN_TYPECONGE') = 'PRI') or (TypeRtt)) then
         begin
         if (PlusMois(GetParamSocSecur('SO_PGECABDATEINTEGRATION',Idate1900), -1) > T_Abs.GetValue('PCN_DATEDEBUTABS'))
             and (T_Abs.GetValue('PCN_VALIDSALARIE') <> 'ADM') then
           begin
           Error := 20;
           NomControl := 'PCN_DATEDEBUTABS';
           Exit;
           end;
         ControleAssezRestant(Nil,{T_Abs.GetValue('PCN_SALARIE'),} T_Abs.getValue('PCN_TYPECONGE'), Action,T_Abs.getValue('PCN_JOURS'),0 );
         end;
      End;

  { Testons la date de validité de l'absence, elle ne doit pas être postérieur à la date de sortie }
  { ni antérieur à la date d'entrée du salarié }
    if ((T_Abs.GetValue('PCN_TYPECONGE') = 'PRI') and (T_Abs.GetValue('PCN_TYPEMVT') = 'CPA')) or (T_Abs.GetValue('PCN_TYPEMVT') = 'ABS') then
    begin
      if ((T_Abs.GetValue('PCN_DATEFINABS') <> idate1900) and (TSal.GetValue('PSA_DATESORTIE') > idate1900) and (TSal.GetValue('PSA_DATESORTIE') < T_Abs.GetValue('PCN_DATEFINABS'))) then
      begin
        Error := 16;
        NomControl := 'PCN_DATEFINABS';
        exit;
      end;
      if ((T_Abs.GetValue('PCN_DATEDEBUTABS') <> idate1900) and (TSal.GetValue('PSA_DATEENTREE') <> idate1900) and (TSal.GetValue('PSA_DATEENTREE') > T_Abs.GetValue('PCN_DATEDEBUTABS'))) then
      begin
        Error := 23;
        NomControl := 'PCN_DATEDEBUTABS';
        exit;
      end;
    end;


  { Testons le non chevauchement de l'absence sur pls mois }
  If not (GetParamSocSecur ('SO_PGABSENCECHEVAL', False)) then
    if (T_Abs.GetValue('PCN_TYPEMVT') = 'ABS') OR (T_Abs.GetValue('PCN_TYPECONGE') = 'PRI') Then
      Begin
      DecodeDate(T_Abs.GetValue('PCN_DATEDEBUTABS'),YYD,MMD,JJ);
      DecodeDate(T_Abs.GetValue('PCN_DATEFINABS'),YYF,MMF,JJ);
      IF (MMD <> MMF) OR (YYD <> YYF) then
        Begin
        Error := 30;
        NomControl := 'PCN_DATEFINABS';
        exit;
        End;
      End;


  { Testons le non chevauchement des périodes sur d'autres absences }
  if ((T_Abs.GetValue('PCN_TYPECONGE')= 'PRI') or (T_Abs.GetValue('PCN_TYPEMVT') = 'ABS')) then
      Begin
      DateD := T_Abs.GetValue('PCN_DATEDEBUTABS');
      DateF := T_Abs.GetValue('PCN_DATEFINABS');
      FinDj := T_Abs.GetValue('PCN_FINDJ');
      DebutDj := T_Abs.GetValue('PCN_DEBUTDJ');
      if DateD <> DateF then
          StDate := 'OR (PCN_DATEDEBUTABS ="' + usdatetime(Datef) + '" AND PCN_DEBUTDJ = "MAT") '+
                    'OR (PCN_DATEFINABS ="' + usdatetime(Dated) + '" AND PCN_FINDJ = "PAM") ';

      St := PgGetClauseAbsAControler(LcT_MotifAbs,LcTob_MotifAbs,T_Abs.GetValue('PCN_TYPECONGE'));

      St := 'SELECT PCN_DATEDEBUTABS,PCN_DATEFINABS,PCN_TYPECONGE,PCN_TYPEMVT,PCN_ORDRE ' +
         'FROM ABSENCESALARIE ' +
         'WHERE PCN_SALARIE = "' + T_Abs.GetValue('PCN_SALARIE') + '" ' + St +
         'AND PCN_ETATPOSTPAIE <> "NAN" '+
         'AND (((PCN_DATEDEBUTABS >"' + usdatetime(DateD) + '" AND PCN_DATEDEBUTABS < "' + usdatetime(DateF) + '") ' +
         'OR (PCN_DATEFINABS >"' + usdatetime(Datef) + '" AND PCN_DATEDEBUTABS < "' + usdatetime(Dated) + '") ' +
         'OR (PCN_DATEFINABS <"' + usdatetime(Datef) + '" AND PCN_DATEFINABS > "' + usdatetime(DateD) + '"))' +
         'OR (PCN_DATEFINABS ="' + usdatetime(Datef) + '" AND PCN_FINDJ = "' + (FinDj) + '") ' +
         'OR (PCN_DATEDEBUTABS ="' + usdatetime(Dated) + '" AND PCN_DEBUTDJ = "' + (Debutdj) + '") ' +
         'OR (PCN_DATEFINABS ="' + usdatetime(Dated) + '" AND PCN_FINDJ = "' + Debutdj + '") ' +
         'OR (PCN_DATEDEBUTABS ="' + usdatetime(Datef) + '" AND PCN_DEBUTDJ = "' + Findj + '") '+
         StDate+')';

       if Origine = 'E' then
           st := st + ' AND (PCN_VALIDRESP <> "NAN" AND PCN_VALIDRESP <> "REF")';
         Q := opensql(st, True);
         if not Q.eof then
           begin
           if ((Action = 'CREATION') and (Q.RecordCount = 1))
           or ((Action = 'MODIFICATION') and (Q.RecordCount > 1))
           or not ((Action = 'MODIFICATION') and (Q.RecordCount = 1)
           and (T_Abs.GetValue('PCN_ORDRE') = Q.FindField('PCN_ORDRE').AsInteger)
           and (T_Abs.GetValue('PCN_TYPEMVT') = Q.FindField('PCN_TYPEMVT').AsString)) then
             begin
             Error := 15;
             NomControl := 'PCN_TYPECONGE';
             Ferme(Q);
             exit;
             end;
           end;
         ferme(Q);
      end;
   (*  ****
  if ((TypeConge <> nil) and (Ejours <> nil) and (DateDebutcp <> nil)
    and (DateFincp <> nil) and (DateValidite <> nil)) then
  begin


    if  (GetField('PCN_DATEVALIDITE') <> null) and (GetField('PCN_TYPEMVT') = 'CPA') then   { PT63 }
      if GetField('PCN_DATEVALIDITE') > idate1900 then
      begin
        Periode := CalculPeriode(DTClot, getfield('PCN_DATEVALIDITE'));
        if Periode > 1 then
        begin
          Init := HShowMessage('1;Mouvements Congés payés;Attention, La date de validité de ce mouvement le rendra désormais inaccessible.#13#10Etes vous sûr de vouloir continuer ?;Q;YNC;N;N;', '', '');
          if Init = mrYes then
          else Error := 14; //PT- 7-11
        end;
        if periode = -9 then
        begin
          Error := 7;
          //PT- 7-11 LastErrorMsg:='Il est interdit de saisir des mouvements avec une date postérieure à la date de clôture + 1 an' ;
          exit;
        end;
      end;
  end;


   *)



  { Dans le cas d'une saisie déportée }
  { Testons l'affectation de saisie de l'utilisateur }
  if (ACTION = 'CREATION') and (Origine = 'E') then
     begin
     {$IFDEF ETEMPS}
     if not IfMotifabsenceSaisissable(T_Abs.GetValue('PCN_TYPECONGE'), Ficprec) then
        begin
        Error := 27;
        NomControl := 'PCN_TYPECONGE';
        exit;
        end;
    (*else //PT39
      if (FicPrec = 'ADM') and ((T_Abs.GetValue('PCN_TYPECONGE') = 'MAL') or (T_Abs.GetValue('PCN_TYPECONGE') = 'MAT')
      or (T_Abs.GetValue('PCN_TYPECONGE') = 'PAT') or (T_Abs.GetValue('PCN_TYPECONGE') = 'ACT')
      or (T_Abs.GetValue('PCN_TYPECONGE') = 'AM1') or (T_Abs.GetValue('PCN_TYPECONGE') = 'AMS')) then
    begin
      if T_Abs.GetValue('PCN_VALIDRESP') <> 'VAL' then SetField('PCN_VALIDRESP', 'VAL');
      if T_Abs.GetValue('PCN_EXPORTOK') <> 'X' then SetField('PCN_EXPORTOK', 'X');
    end; *)
    {$ENDIF}

    { Testons la présence du salarié dans la gestion d'affaire }
      If not EstLibre(T_Abs.GetValue('PCN_DATEDEBUTABS'),T_Abs.GetValue('PCN_DATEFINABS'),
      T_Abs.GetValue('PCN_SALARIE'),T_Abs.GetValue('PCN_DEBUTDJ'),T_Abs.GetValue('PCN_FINDJ'),pStAffaire) then
         Begin
         PgiBox('Vous avez une intervention planifiée sur cette période pour l''affaire : '+
         pStAffaire+'#13#10 Merci de régulariser le planning avant de saisir cette absence.','Saisie d''absence');  
         Error := 1;
         NomControl := 'PCN_DATEDEBUTABS';
         exit;
         End;

    { Testons la saisie du remplaçant sur site obligatoire }
    if (GetParamSocSecur('SO_PGECABSAISIEREMP',0)) and (isnumeric(T_Abs.GetValue('PCN_JOURS'))) then
      if (Trim(T_Abs.GetValue('PCN_LIBCOMPL1')) = '') and (T_Abs.GetValue('PCN_JOURS') >= GetParamSocSecur('SO_PGECABSAISIEREMP',0)) then
      begin
        Error := 22;
        NomControl := 'PCN_LIBCOMPL1';
        exit;
      end;

    { Avertissement législatif sur le fractionnement }
    if GetParamSocSecur('SO_PGECABFRACTIONNEMENT',False) then
      if (T_Abs.GetValue('PCN_TYPECONGE') = 'PRI') and (ficPrec = 'SAL') then
        if PgiAsk(RecupMessageErrorAbsence(21), 'Saisie d''absence') = Mrno then
        begin
          Error := 1;
          exit;
        end;

    if GetParamSocSecur('SO_PGECABGESTIONMAIL',False) then
      begin
      ChargeInfoRespSalMail(T_Abs.GetValue('PCN_SALARIE'),FicPrec,TypeSal,Resp,NomResp,MailResp,MailSal,OkMailEnvoie,MailDate);


      Body := TStringList.Create;
      if (ficPrec = 'SAL')  and (MailResp <> '') then
      begin
        //Gestion de l'envoi de mail lors de la validation salarié
        //Ajout des paramètres sociétés : gestion de la durée de validité du mail
        //Test si raz ou validité du mail pour envoie
        if (OkMailEnvoie = False) or (GetParamSocSecur('SO_PGECABVALIDITEMAIL',0) = 0)
          or ((MailDate <> idate1900) and ((Date - MailDate) > GetParamSocSecur('SO_PGECABVALIDITEMAIL',0))) then
        begin
          Body := PgRecupBodyMailCollaborateur(T_Abs,Resp);
          if PGEnvoieMailCollaborateur('ECongés : Absence à valider', MailResp, Body) then
          begin
            PgMajIndicMailSalENV(Resp);
            //ExecuteSql('UPDATE DEPORTSAL SET PSE_EMAILENVOYE="X",PSE_EMAILDATE="' + UsDateTime(date) + '" WHERE PSE_SALARIE="' + Resp + '"');
            PGIInfo('Un message a été envoyé sur la messagerie de votre responsable ' + NomResp,'Saisie d''absence');
          end;
        end;
      end
      else
        { si création absence pour salarié alors envoie mail }
        if ((FicPrec = 'RESP') or (FicPrec = 'ADM')) and (MailSal <> '') then
      begin
        St := GetParamSocSecur('SO_LIBELLE','');     
        if St <> '' then St := ' de la société '+St;
        if FicPrec = 'ADM' then St := 'L''administrateur ' + RechDom('PGSALARIE', V_Pgi.UserSalarie, False)+ St else St := 'Le responsable ' + NomResp + St;
        St := St + ' vient de saisir une absence ' + RechDom('PGMOTIFABSENCED', T_Abs.GetValue('PCN_TYPECONGE'), False);
        St := St + ' du ' + DateToStr(T_Abs.GetValue('PCN_DATEDEBUTABS')) + ' ' + RechDom('PGDEMIJOURNEE', T_Abs.GetValue('PCN_DEBUTDJ'), False);
        St := St + ' au ' + DateToStr(T_Abs.GetValue('PCN_DATEFINABS')) + ' ' + RechDom('PGDEMIJOURNEE', T_Abs.GetValue('PCN_FINDJ'), False);
        St := St + ' soit une durée de ' + FloatToStr(T_Abs.GetValue('PCN_JOURS')) + ' jours.';
        Body.Add(st);
        if PGEnvoieMailCollaborateur('ECongés : Absence saisie', MailSal, Body) then
          PGIInfo('Un message a été envoyé sur la messagerie du salarié ' + RechDom('PGSALARIE', T_Abs.GetValue('PCN_SALARIE'), False) + '.', 'Saisie d''absence');
      end;
      Body.Free;
    end;
  end;
  //RAZ du boolean de l'envoi de mail du responsable lors de la validation de l'abs salarié
  if (Error = 0) and (Origine = 'E') then
  begin
    if (ACTION = 'MODIFICATION') then
    begin
      IsMailRefOrAnu := (T_Abs.GetValue('PCN_VALIDRESP') = 'REF') or (T_Abs.GetValue('PCN_VALIDRESP') = 'NAN');


      if (FicPrec <> 'SAL') and (MailResp <> '') and (OkMailEnvoie = True) and (T_Abs.GetValue('PCN_VALIDRESP') <> 'ATT') then
         PgMajIndicMailSalRAZ(Resp);
        //ExecuteSql('UPDATE DEPORTSAL SET PSE_EMAILENVOYE="-" WHERE PSE_SALARIE="' + Resp + '"');
      { en modification si absence refusée ou annulée }
      if ((FicPrec = 'RESP') or (FicPrec = 'ADM')) and (IsMailRefOrAnu) and (MailSal <> '') then
      begin
        if T_Abs.GetValue('PCN_VALIDRESP') = 'REF' then StTerme := 'de refuser' else StTerme := 'd''annuler';
        Body := TStringList.Create;
        St := GetParamSocSecur('SO_LIBELLE','');
        if St <> '' then St := ' de la société '+St; 
        if FicPrec = 'ADM' then St := 'L''administrateur ' + RechDom('PGSALARIE', V_Pgi.UserSalarie, False) + St else St := 'Le responsable ' + NomResp + St; 
        St := St + ' vient ' + StTerme + ' l''absence ' + RechDom('PGMOTIFABSENCED', T_Abs.GetValue('PCN_TYPECONGE'), False);
        St := St + ' du ' + DateToStr(T_Abs.GetValue('PCN_DATEDEBUTABS')) + ' ' + RechDom('PGDEMIJOURNEE', T_Abs.GetValue('PCN_DEBUTDJ'), False);
        St := St + ' au ' + DateToStr(T_Abs.GetValue('PCN_DATEFINABS')) + ' ' + RechDom('PGDEMIJOURNEE', T_Abs.GetValue('PCN_FINDJ'), False);
        St := St + ' soit une durée de ' + FloatToStr(T_Abs.GetValue('PCN_JOURS')) + ' jours.';
        Body.Add(st);
        if T_Abs.GetValue('PCN_VALIDRESP') = 'REF' then StTerme := 'refusée' else StTerme := 'annulée';
        if PGEnvoieMailCollaborateur('ECongés : Absence ' + StTerme, MailSal, Body) then
          PGIInfo('Un message a été envoyé sur la messagerie du salarié ' + RechDom('PGSALARIE', T_Abs.GetValue('PCN_SALARIE'), False) + '.', 'Saisie d''absence');
        Body.Free;
      end;
    end;

  end;

FreeAndNil(LcTob_motifabs);
End;


Function  RecupMessageErrorAbsence(Error : integer) : string;
Begin
  Case Error of     //  TexteMessage: array[1..30] of string = (
    1  : Result := 'La date de fin ne peut être inférieure à la date de début d''absence.';
    2  : Result := 'Le nombre de jours est invalide.';
    3  : Result := 'Le nombre de jours doit être positif.';
    4  : Result := 'Vous devez renseigner : le type d''absence.';
    5  : Result := 'Le nombre de jours pris repris ne peut être supérieur au nombre de jours acquis repris : ' {ComplMessage};
    6  : Result := 'Vous ne pouvez saisir deux mouvements de type reprise acquis sur une même période.';
    7  : Result := 'Vous ne pouvez saisir de mouvement correspondant à une période non ouverte.';
    8  : Result := 'Vous ne pouvez saisir deux mouvements de type reprise pris sur une même période.';
    9  : Result := 'Vous devez renseigner : le type de mouvement.';
    10 : Result := 'Un mouvement d''absence ne peut être à cheval sur la date de clôture.';
    11 : Result := 'Vous devez renseigner : la date de début d''absence.';
    12 : Result := 'Vous devez renseigner : la date de fin d''absence.';
    13 : Result := 'Vous devez renseigner : la date de validité.';
    14 : Result := ''; {Juste pour bloquer la validation}
    15 : Result := 'Vous devez modifier votre période d''absence. Une absence existe déjà.';
    16 : Result := 'La date de fin d''absence doit être antérieure à la date de sortie du salarié.';
    17 : Result := 'Le nombre de jours maximum octroyés pour ce motif est dépassé : ' {ComplMessage};
    18 : Result := 'Le nombre d''heures maximum octroyés pour ce motif est dépassé : ' {ComplMessage};
    19 : Result := 'Vous devez renseigner une valeur pour le nombre de jours, de mois ou la base.';
    20 : Result := 'Vous ne pouvez pas saisir d''absence antérieur à un mois avant la date d''intégration.';
    21 : Result := 'Conformément aux dispositions de l''article L 223-8 du Code du Travail,#13#10' + //PT20 Ajout nouveau message
                   'je déclare expressément renoncer aux jours de congés supplémentaires pour fractionnement,#13#10' +
                   'ce dernier étant réalisé à ma seule initiative pour raisons personnelles.';
    22 : Result := 'Vous devez renseigner le nom du remplaçant sur site.';
    23 : Result := 'La date de début d''absence doit être postérieure à la date d''entrée du salarié.';
    24 : Result := 'Saisie impossible : Vous ne gérez pas les congés payés au niveau dossier.';
    25 : Result := 'Saisie impossible : Vous ne gérez pas les congés payés au niveau établissement.';
    26 : Result := 'Vous devez renseigner : le matricule du salarié.';
    27 : Result := 'Vous n''êtes pas autorisé à saisir ce type d''absence.';
    28 : Result := 'La date de validité ne peut être antérieur à celle du dernier solde de tout compte : ';
    29 : Result := 'Saisie impossible : Vous ne gérez pas les congés payés au niveau salarié.';
    30 : Result := 'Vous ne pouvez saisir une absence à cheval sur plusieurs mois.';
    31 : Result := 'L''absence ne peut s''effectuer de l''après-midi au matin sur la même journée.';
    32 : Result := 'Echec lors du calcul du récapitulatif.Suppression annulée.';
  end;
End;



Function PgMajAbsEtatValidSal(Sal,Mode,Typmvt,ValidPar,ValidResp : String ; ordre : integer) : Integer;
var st    : String;
begin
if ValidResp <>'ATT' then exit;
//TypMvt := TFmul(Ecran).Q.findfield('PCN_TYPEMVT').asstring;
//Sal    := TFmul(Ecran).Q.findfield('PCN_SALARIE').asstring;
//Ordre  :=TFmul(Ecran).Q.findfield('PCN_ORDRE').asinteger;
st := 'UPDATE ABSENCESALARIE SET PCN_VALIDRESP = "'+Mode+'",PCN_VALIDABSENCE="'+ValidPar+'", '+
      'PCN_DATEMODIF = "'+UsDateTime(Date)+'" '+
      'WHERE PCN_TYPEMVT = "'+TypMvt+'" AND PCN_SALARIE ="'+Sal+'" '+
      'AND PCN_ORDRE ='+IntToStr(Ordre);
Result := ExecuteSql (st);
PGExeCalculRecapAbsEnCours(Sal);
End;

//RAZ du boolean de l'envoi de mail du responsable lors de la validation de l'abs salarié
Procedure PgMajIndicMailSalRAZ(ValidPar : String) ;
Begin
ExecuteSql('UPDATE DEPORTSAL SET PSE_EMAILENVOYE="-" WHERE PSE_SALARIE="'+ValidPar+'"');
End;

Procedure PgMajIndicMailSalENV(ValidPar : String) ;
Begin
ExecuteSql('UPDATE DEPORTSAL SET PSE_EMAILENVOYE="X",PSE_EMAILDATE="' + UsDateTime(date) + '" WHERE PSE_SALARIE="' + ValidPar + '"');
End;



Function PgRecupBodyMailCollaborateur(T_Abs : Tob; Resp : String) : TStringList;
Var St : string;
    Q  : TQuery;
Begin
  St := GetParamSocSecur('SO_LIBELLE','');
  if St <> '' then St := ' de la société '+St;
  St := 'Le salarié ' + RechDom('PGSALARIE', T_Abs.GetValue('PCN_SALARIE'), False) + St;
  St := St + ' vient de saisir une absence ' + RechDom('PGMOTIFABSENCED', T_Abs.GetValue('PCN_TYPECONGE'), False);
  St := St + ' du ' + DateToStr(T_Abs.GetValue('PCN_DATEDEBUTABS')) + ' ' + RechDom('PGDEMIJOURNEE', T_Abs.GetValue('PCN_DEBUTDJ'), False);
  St := St + ' au ' + DateToStr(T_Abs.GetValue('PCN_DATEFINABS')) + ' ' + RechDom('PGDEMIJOURNEE', T_Abs.GetValue('PCN_FINDJ'), False);
  St := St + ' soit une durée de ' + FloatToStr(T_Abs.GetValue('PCN_JOURS')) + ' jours.';
  Result.Add(st);
  Q := OpenSql('SELECT COUNT (*) AS NB FROM ABSENCESALARIE ' +
  'LEFT JOIN DEPORTSAL ON  PSE_SALARIE=PCN_SALARIE ' +
  'LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI## PCN_TYPECONGE=PMA_MOTIFABSENCE ' +
  'WHERE PSE_RESPONSABS="' + Resp + '" AND PMA_MOTIFEAGL="X" ' +
  'AND PCN_VALIDRESP="ATT"', True);
  if not Q.Eof then
     if Q.FindField('NB').AsInteger > 0 then
        begin
        St := 'Par ailleurs, vous avez ' + IntToStr(Q.FindField('NB').AsInteger) + ' autre(s) absence(s) à valider.';
        Result.Add(St);
        end;
   Ferme(Q);
End;
{ FIN PT14 }

{ DEB PT15 }
Function  PgCalculDateAnciennete(DateArret : TDateTime; DateCpAnc : String ) : TDateTime;
Var
   IntJ, IntM : Integer;
//   Ok         : Boolean;
   YY, MM, JJ,aaj, mmj, jjj : Word;
//   DateTemp : TdateTime;
Begin
  Result := idate1900;
  decodedate(DateArret, aaj, mmj, jjj);
  if (Length(DateCpAnc) = 4)  and (ISNumeric(DateCpAnc)) then
     Begin
     IntJ := StrToInt(Copy(DateCpAnc, 1, 2));
     IntM := StrToInt(Copy(DateCpAnc, 3, 2));
     if PgOkFormatDateJJMM(DateCpAnc) then
        DateCpAnc:= DateToStr(EncodeDate(2000,IntM, IntJ));
     End;

  if isValidDate(DateCpAnc) then
     Begin
     DecodeDate(StrToDate(DateCpAnc), YY, MM, JJ);
     Result := PGEncodeDateBissextile(aaj, mm, jj);
     End;

End;


Function  PgOkFormatDateJJMM ( St : String)  : Boolean;
Var IntJ, IntM : Integer;
Begin
  result := False;
  if (Length(St) = 4)  and (ISNumeric(St)) then
     Begin
     IntJ := StrToInt(Copy(St, 1, 2));
     IntM := StrToInt(Copy(St, 3, 2));
     if ((IntM=1) OR (IntM=3) OR (IntM=5) OR (IntM=7) OR (IntM=8) OR (IntM=10) OR (IntM=12)) AND (IntJ<32) then result := True;
     if ((IntM=4) OR (IntM=6) OR (IntM=9) OR (IntM=11)) AND (IntJ<31) then result := True;
     if (IntM=2) AND (IntJ<30) then result := True;
     End;
End;
{ FIN PT15 }

{ DEB PT16 }
(*PT20
function PGEncodeDateBissextile(AA, MM, JJ: WORD): TDateTime;
begin
  if (MM = 2) and (JJ = 29) then //Année bissextile
  begin
    Result := encodedate(AA, MM, 1);
    Result := FindeMois(Result);
  end
else
    Result := encodedate(AA, MM, JJ);   
end; *)
{ FIN PT16 }


end.
