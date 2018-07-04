{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/09/2004
Modifié le ... :
Description .. : Unité qui contient l'ensemble des fonctions et procédures
Suite ........ : utilisées pour la DADS et la DADS bilatérale
Mots clefs ... : PAIE;PGDADSU;PGDADSB
*****************************************************************}
{
PT1   : 13/10/2006 VG V_70 Suppression du fichier de contrôle - mise en table
                           des erreurs
PT2   : 20/07/2007 FC V_72 FQ 14423 Gestion des différents régimes
}
unit PgDADSCommun;

interface

uses SysUtils, UTob,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
{$ENDIF}
     hctrls,
     uTobDebug,
     ParamSoc;   

Type TControle = record
     Salarie : string;           //Code salarié
     Libelle : string;           //Nom et prénom
     TypeD : string;             //Type de DADS
     Num : integer;              //Numéro de période
     DateDeb : TDateTime;        //Date de début de période
     DateFin : TDateTime;        //Date de fin de période
     Segment : string;           //Segment comportant une anomalie
     Exercice : string;          //Exercice
     Explication : string;       //Explication de l'erreur
     CtrlBloquant : boolean;     //Anomalie entraînant un rejet
end;


procedure ChTOBSal(Salarie : string);
procedure LibTOB();
procedure DeleteErreur (Salarie : string=''; OrigineCrit : string='';
                        Periode : integer=100);
procedure EcrireErreur(Controle : TControle; TMere : TOB=nil; DadsNom : string='');
procedure EcrireErreurKO ();
procedure AdresseNormalisee(origine : string; var numero, nom : string);

procedure TiersToDads2 ( TobTiers : TOB ; var TobDads2 : TOB ; var Controle : TControle );
procedure Dads2toTiers ( TobDads2 : TOB ; var TobTiers : TOB );

var
TBase, TContrat, TCot, TCot2, TDADSUD, THistoCumSal, THistoSal, TPaie : TOB;
TRem, TSal, TTauxAT : TOB;

DebExer, FinExer : TDateTime;

Erreur : Boolean;

PGAnnee, PGExercice, PGExercice2, TypeD : String;

implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/09/2002
Modifié le ... :   /  /
Description .. : Chargement des TOB necessaires au calcul des éléments
Suite ........ : de la DADS-U à partir de la fiche salarié
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure ChTOBSal(Salarie : string);
var
StSal : String;
QRechSal : TQuery;
begin
//Chargement de la TOB SALARIES
StSal:='SELECT PSA_SALARIE, PSA_ETABLISSEMENT, ET_LIBELLE, ET_APE,'+
       ' PSA_DATEENTREE, PSA_MOTIFENTREE, PSA_DATESORTIE, PSA_MOTIFSORTIE,'+
       ' PSA_DATEENTREEPREC, PSA_DATESORTIEPREC, PSA_NUMEROSS, PSA_SEXE,'+
       ' PSA_LIBELLE, PSA_PRENOM, PSA_NOMJF, PSA_SURNOM, PSA_CIVILITE,'+
       ' PSA_ADRESSE1, PSA_ADRESSE2, PSA_ADRESSE3, PSA_CODEPOSTAL, PSA_VILLE,'+
       ' PSA_PAYS, PSA_DATENAISSANCE, PSA_COMMUNENAISS, PSA_DEPTNAISSANCE,'+
       ' PSA_PAYSNAISSANCE, PSA_NATIONALITE, PSA_MULTIEMPLOY,'+
       ' PSA_LIBELLEEMPLOI, PSA_CODEEMPLOI, PSA_CONDEMPLOI, PSA_DADSPROF,'+
       ' PSA_DADSCAT, PSA_CONVENTION, PSA_COEFFICIENT, PSA_INDICE, PSA_NIVEAU,'+
       ' PSA_REGIMESS,PSA_TAUXPARTIEL, PSA_ORDREAT, PSA_TAUXPARTSS,'+
       ' PSA_REGIMEMAL, PSA_REGIMEAT, PSA_REGIMEVIP, PSA_REGIMEVIS, PSA_TYPEREGIME,'+ //PT2
       ' PSA_TRAVETRANG, PSA_REMPOURBOIRE, PSA_ALLOCFORFAIT, PSA_REMBJUSTIF,'+
       ' PSA_PRISECHARGE, PSA_AUTREDEPENSE, PSA_PRUDHCOLL, PSA_PRUDHSECT,'+
       ' PSA_PRUDHVOTE, PSA_DADSDATE, PSA_DADSFRACTION, PSA_SITUATIONFAMIL,'+
       ' PSA_PERSACHARGE, PSA_UNITETRAVAIL, PSA_DATEANCIENNETE, PSA_HORHEBDO,'+
       ' PSA_HORAIREMOIS, PSA_SORTIEDEFINIT, PSA_PCTFRAISPROF'+
       ' FROM SALARIES'+
       ' LEFT JOIN ETABLISS ON'+
       ' PSA_ETABLISSEMENT=ET_ETABLISSEMENT WHERE'+
       ' PSA_SALARIE = "'+Salarie+'"';

QRechSal:=OpenSql(StSal,TRUE);
TSal := TOB.Create('Les salaries', NIL, -1);
TSal.LoadDetailDB('SAL','','',QRechSal,False);
Ferme(QRechSal);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/09/2002
Modifié le ... :   /  /
Description .. : Libération de la TOB salarié
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure LibTOB();
begin
FreeAndNil(TSal);
end;


//PT1
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 31/05/2006
Modifié le ... :   /  /
Description .. : Suppression des erreurs de la table DADSCONTROLE
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure DeleteErreur (Salarie : string=''; OrigineCrit : string='';
                        Periode : integer=100);
var
StDelete : string;
begin
if (Salarie = '') then
   StDelete:= 'DELETE FROM DADSCONTROLE WHERE'+
              ' PSU_TYPE = "'+TypeD+'" AND'+
              ' PSU_EXERCICEDADS = "'+PGExercice+'"'
else
   StDelete:= 'DELETE FROM DADSCONTROLE WHERE'+
              ' PSU_SALARIE = "'+Salarie+'" AND'+
              ' PSU_TYPE = "'+TypeD+'" AND'+
              ' PSU_EXERCICEDADS = "'+PGExercice+'"';
if ((Salarie = '') and (OrigineCrit = '')) then
   StDelete:= StDelete+' AND PSU_ORIGINECRIT LIKE "S%"'
else
if (OrigineCrit <> '') then
   StDelete:= StDelete+' AND PSU_ORIGINECRIT LIKE "'+OrigineCrit+'%"';

if (Periode <> 100) then
   StDelete:= StDelete+' AND PSU_ORDRE='+IntToStr (Periode);
ExecuteSQL(StDelete) ;
end;
//FIN PT1

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/02/2002
Modifié le ... :   /  /
Description .. : Ecriture dans le fichier de contrôle
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure EcrireErreur(Controle : TControle; TMere : TOB=nil; DadsNom : string='');
var
TDADSCONTROLE : TOB;
begin
{PT1
if (Erreur = False) then
   begin
   Writeln(FRapport, 'Attention, vous devez vérifier :');
   Erreur := True;
   end;
Writeln(FRapport, Chaine);
}
TDADSCONTROLE:= TOB.Create ('DADSCONTROLE', TMere, -1);
TDADSCONTROLE.PutValue('PSU_SALARIE', Controle.Salarie);
if (Controle.Libelle<>'') then
   TDADSCONTROLE.AddChampSupValeur ('PSU_LIBELLE', Controle.Libelle);
TDADSCONTROLE.PutValue('PSU_ORIGINECRIT', Copy (Controle.Segment, 1, 3));
TDADSCONTROLE.PutValue('PSU_TYPE', Controle.TypeD);
TDADSCONTROLE.PutValue('PSU_ORDRE', Controle.Num);
TDADSCONTROLE.PutValue('PSU_DATEDEBUT', Controle.DateDeb);
TDADSCONTROLE.PutValue('PSU_DATEFIN', Controle.DateFin);
if (Controle.Segment <> 'SOK') then
   TDADSCONTROLE.PutValue('PSU_SEGMENT', Controle.Segment);
if (DadsNom <> '') then
   TDADSCONTROLE.AddChampSupValeur ('PDL_DADSNOM', DadsNom);
TDADSCONTROLE.PutValue('PSU_EXERCICEDADS', Controle.Exercice);
TDADSCONTROLE.PutValue('PSU_EXPLICATION', Controle.Explication);
if (Controle.CtrlBloquant = True) then
   TDADSCONTROLE.PutValue('PSU_CTRLBLOQUANT', 'X')
else
   TDADSCONTROLE.PutValue('PSU_CTRLBLOQUANT', '-');
if (TMere = nil) then
   begin
   TDADSCONTROLE.SetAllModifie (TRUE);
   TDADSCONTROLE.InsertOrUpdateDB (FALSE);
   FreeAndNil(TDADSCONTROLE);
   end;
//FIN PT1
end;

//PT1
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 28/06/2006
Modifié le ... :   /  /
Description .. : Ecriture d'un enregistrement KO
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure EcrireErreurKO ();
var
ErreurDADSU : TControle;
begin
ErreurDADSU.Salarie:= GetParamSoc ('SO_LIBELLE');
ErreurDADSU.TypeD:= TypeD;
ErreurDADSU.Num:= 0;
ErreurDADSU.DateDeb:= DebExer;
ErreurDADSU.DateFin:= FinExer;
ErreurDADSU.Segment:= 'SKO';
ErreurDADSU.Exercice:= PGExercice;
ErreurDADSU.Explication:= 'Le contrôle n''a pas été effectué';
ErreurDADSU.CtrlBloquant:= False;

EcrireErreur (ErreurDADSU);
end;
//FIN PT1

{***********A.G.L.***********************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/07/2001
Modifié le ... : 27/08/2001
Description .. : Récupération à partir d'un champ ou d'une chaine du
Suite ........ : numéro de voie et du nom de la voie
Mots clefs ... : PAIE,PGDADSU,ADRESSE
*****************************************************************}
procedure AdresseNormalisee(origine : string; var numero, nom : string);
var
Adresse : string;
i, Long : integer;
begin
numero := '     ';
nom := '                                   ';
Adresse := origine;

i := 1;
Long := Length(Adresse);

While ((Adresse[i]>='0') and (Adresse[i]<='9')) do
      begin
      numero[i] := Adresse[i];
      i := i+1;
      end;

While ((Adresse[i] = ',') or (Adresse[i] = ' ')) do
      i:= i+1;

nom := Copy(Adresse, i, Long-i+1);
numero := Trim(numero);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 12/09/2007
Modifié le ... : 12/09/2007
Description .. : Permet de copier les enregistrements de la table tiers 
Suite ........ : vers la table dads2. FQ 20733
Mots clefs ... : 
*****************************************************************}
procedure TiersToDads2 ( TobTiers : TOB ; var TobDads2 : TOB ; var Controle : TControle );
var
  Buf, numero, nom : String;     
  Longueur: integer;
begin
  TobDads2.PutValue ('PDH_SIRETBEN', TobTiers.GetValue('T_SIRET'));

  if (TobTiers.GetValue ('T_PRENOM') <> '') then
  begin
     TobDads2.PutValue ('PDH_NOMBEN', TobTiers.GetValue('T_LIBELLE'));
     TobDads2.PutValue ('PDH_PRENOMBEN', TobTiers.GetValue ('T_PRENOM'));
  end
  else
     TobDads2.PutValue ('PDH_RAISONSOCBEN', TobTiers.GetValue('T_LIBELLE'));


  {Concaténation de adresse2 et adresse3 afin de remplir le complément d'adresse}
  Buf:= '';
  if TobTiers.FieldExists('T_ADRESSE2') and (TobTiers.GetValue('T_ADRESSE2') <> '') then
  begin
     if TobTiers.FieldExists('T_ADRESSE3') and (TobTiers.GetValue('T_ADRESSE3') <> '') then
        Buf := TobTiers.GetValue('T_ADRESSE2') + ' ' + TobTiers.GetValue('T_ADRESSE3')
     else
        Buf := TobTiers.GetValue('T_ADRESSE2');
  end
  else
     if TobTiers.FieldExists('T_ADRESSE3') and (TobTiers.GetValue('T_ADRESSE3') <> '') then
        Buf := TobTiers.GetValue('T_ADRESSE3');
  if (Buf <> '') then
  begin
     {vérification que la longueur ne dépasse pas 32, sinon, tronquage}
     Longueur := Length(Buf);
     if (Longueur > 32) then
     begin
        Buf:= Copy(Buf, 1, 32);
        Controle.Segment:= 'PDH_ADRCOMPL';
        Controle.Explication:= 'L''adresse a été tronquée';
        Controle.CtrlBloquant:= False;
        EcrireErreur(Controle);
     end;
     TobDads2.PutValue('PDH_ADRCOMPL', Buf);
  end;
  {Adresse1 est coupée en 2 : d'un côté, le numéro, de l'autre côté, le nom}
  numero:= '';
  nom:= '';
  if TobTiers.FieldExists('T_ADRESSE1') and (TobTiers.GetValue('T_ADRESSE1') <> '') then
     AdresseNormalisee(TobTiers.GetValue('T_ADRESSE1'), numero, nom);

  if (numero <> '') then
  begin
     TobDads2.PutValue('PDH_ADRNUM', numero);
  end;

  if (nom <> '') then
  begin
     {vérification que la longueur ne dépasse pas 26, sinon, tronquage}
     Longueur:= Length(nom);
     if (Longueur > 26) then
     begin
        nom:= Copy(nom, 1, 26);
        Controle.Segment:= 'PDH_ADRNOM';
        Controle.Explication:= 'L''adresse a été tronquée';
        Controle.CtrlBloquant:= False;
        EcrireErreur(Controle);
     end;
     TobDads2.PutValue('PDH_ADRNOM', nom);
  end;

  {On ne récupère le code postal que s'il est renseigné et qu'il est numérique sur
5 chiffres}
  if (not(TobTiers.FieldExists('T_CODEPOSTAL')) or
      (TobTiers.GetValue('T_CODEPOSTAL') = '') or
      (TobTiers.GetValue('T_CODEPOSTAL') < '00000') or
      (TobTiers.GetValue('T_CODEPOSTAL') > '99999')) then
  begin
     Controle.Segment:= 'PDH_CODEPOSTAL';
     Controle.Explication:= 'Le code postal n''est pas renseigné';
     Controle.CtrlBloquant:= True;
     EcrireErreur(Controle);
  end
  else
  begin
     TobDads2.PutValue ('PDH_CODEPOSTAL', TobTiers.GetValue('T_CODEPOSTAL'));
  end;

{On ne récupère la ville que si elle est renseignée, sinon, message d'erreur}
  if not(TobTiers.FieldExists('T_VILLE')) or (TobTiers.GetValue('T_VILLE') = '') then
  begin
     Controle.Segment:= 'PDH_BUREAUDISTRIB';
     Controle.Explication:= 'Le bureau distributeur n''est pas renseigné';
     Controle.CtrlBloquant:= True;
     EcrireErreur(Controle);
  end
  else
  begin
     TobDads2.PutValue ('PDH_BUREAUDISTRIB', TobTiers.GetValue('T_VILLE'));
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 12/09/2007
Modifié le ... : 12/09/2007
Description .. : Permet de copier les champs de la table Dads2 dans la 
Suite ........ : table tiers FQ 20728
Suite ........ :  
Mots clefs ... : 
*****************************************************************}
procedure Dads2toTiers ( TobDads2 : TOB ; var TobTiers : TOB );
begin
  TobTiers.PutValue ('T_SIRET', TobDads2.GetValue('PDH_SIRETBEN'));   

  if (TobDads2.GetValue ('PDH_NOMBEN') <> '') then
  begin
     TobTiers.PutValue ('T_LIBELLE', TobDads2.GetValue('PDH_NOMBEN'));
     TobTiers.PutValue ('T_PRENOM', TobDads2.GetValue ('PDH_PRENOMBEN'));
  end
  else
     TobTiers.PutValue ('T_LIBELLE', TobDads2.GetValue('PDH_RAISONSOCBEN'));


(*  {Concaténation de adresse2 et adresse3 afin de remplir le complément d'adresse}
    BVE : Difficilement portable dans ce sens.

  Buf:= '';
  if TobTiers.FieldExists('T_ADRESSE2') and (TobTiers.GetValue('T_ADRESSE2') <> '') then
  begin
     if TobTiers.FieldExists('T_ADRESSE3') and (TobTiers.GetValue('T_ADRESSE3') <> '') then
        Buf := TobTiers.GetValue('T_ADRESSE2') + ' ' + TobTiers.GetValue('T_ADRESSE3')
     else
        Buf := TobTiers.GetValue('T_ADRESSE2');
  end
  else
     if TobTiers.FieldExists('T_ADRESSE3') and (TobTiers.GetValue('T_ADRESSE3') <> '') then
        Buf := TobTiers.GetValue('T_ADRESSE3');
  if (Buf <> '') then
  begin
     {vérification que la longueur ne dépasse pas 32, sinon, tronquage}
     Longueur := Length(Buf);
     if (Longueur > 32) then
     begin
        Buf:= Copy(Buf, 1, 32);
        Controle.Segment:= 'PDH_ADRCOMPL';
        Controle.Explication:= 'L''adresse a été tronquée';
        Controle.CtrlBloquant:= False;
        EcrireErreur(Controle);
     end;
     TobDads2.PutValue('PDH_ADRCOMPL', Buf);
     Debug('Paie PGI/Calcul TD Bilatéral : PDH_ADRCOMPL');
  end;
  *)
  
  {Adresse1 est coupée en 2 : d'un côté, le numéro, de l'autre côté, le nom}
  if TobDads2.GetValue('PDH_ADRNUM') <> '' then
     TobTiers.PutValue ('T_ADRESSE1',TobDads2.GetValue('PDH_ADRNUM') + ' ' + TobDads2.GetValue('PDH_ADRNOM'))
  else
     TobTiers.PutValue ('T_ADRESSE1',TobDads2.GetValue('PDH_ADRNOM'));

  TobTiers.PutValue ('T_CODEPOSTAL', TobDads2.GetValue('PDH_CODEPOSTAL'));

  TobTiers.PutValue ('T_VILLE', TobDads2.GetValue('PDH_BUREAUDISTRIB'));
end;

end.
