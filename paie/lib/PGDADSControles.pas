{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/08/2002
Modifié le ... :   /  /
Description .. : Regroupe tous les contrôles effectués sur le fichier DADSU
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
{
PT1   : 26/11/2002 VG V591 Contrôle des champs obligatoires : ils ne doivent pas
                           être vides
PT2   : 14/01/2003 VG V591 Pour une meilleure compréhension des anomalies
                           rencontrées lors de la génération du pré-fichier, on
                           affiche dorénavant le libellé du segment
PT3   : 12/03/2003 VG V_42 Tous les champs (obligatoires ou non) doivent être
                           différents de ''
PT4   : 31/03/2003 VG V_42 Contrôle des champs BTP
PT5-1 : 17/04/2003 VG V_42 Les contrôles sont désormais conditionnés au type de
                           déclaration effectué (Exemple: on ne contrôle pas les
                           champs concernant la prévoyance pour un fichier de
                           type TDS seul
PT5-2 : 17/04/2003 VG V_42 Ajout du contrôle de la CSG spécifique conditionnée
                           au type de fichier
PT6-1 : 24/10/2003 VG V_42 Ajout de contrôles lors de la préparation du fichier
PT6-2 : 24/10/2003 VG V_42 La raison sociale n'est plus obligatoire
PT7-1 : 05/07/2004 VG V_50 En raison du stockage de nouveaux champs dans la
                           table DADSLEXIQUE, il n'est plus necessaire de
                           vérifier le type de message en cours de construction
PT7-2 : 05/07/2004 VG V_50 Adaptation cahier des charges V8R00
PT7-3 : 05/07/2004 VG V_50 Ajout du contrôle de la prévoyance
PT8   : 27/01/2005 VG V_60 Lorsqu'il existe une erreur au niveau d'un honoraire,
                           affichage dans le .log de son numéro
PT9   : 31/01/2005 VG V_60 Modification de la table des caractères autorisés
                           FQ N°11949
PT10  : 07/10/2005 VG V_60 Adaptation cahier des charges V8R02
PT11  : 10/11/2005 VG V_65 Correction du message pour les honoraires
                           FQ N°12683
PT12-1: 13/10/2006 VG V_70 Adaptation cahier des charges V8R04
PT12-2: 13/10/2006 VG V_70 Suppression du fichier de contrôle - mise en table
                           des erreurs
PT13  : 16/11/2006 VG V_70 Adaptation cahier des charges V8R04 rectificatif du
                           15/11/2006
PT14  : 15/12/2006 VG V_70 Contrôle du numéro de rattachement IRCANTEC
                           FQ N°13768
PT15  : 24/04/2007 VG V_72 Mauvais contrôle lorsqu'un segment obligatoire
                           n'existe pas
PT16-1: 17/08/2007 VG V_80 Adaptation cahier des charges V8R05 - FQ N°14424
PT16-2: 17/08/2007 VG V_80 Adaptation cahier des charges V8R05 - FQ N°14425
PT16-3: 28/08/2007 VG V_80 Adaptation cahier des charges V8R05
PT17-1: 05/12/2007 VG V_80 Affichage du caractère interdit - FQ N°14961
PT17-2: 05/12/2007 VG V_80 Adaptation cahier des charges V8R06 - FQ N°15016
PT18  : 11/12/2007 VG V_80 Gestion de la nature "Honoraires seuls"
}
unit PGDADSControles;

interface

uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     HEnt1,
     sysutils,
     UTob,
     HCtrls,
     hdebug,
     PgDADSCommun,
     ed_tools,
     HStatus,
     PgOutils2;

function ControleCar (var Donnee : string; Nature : string; Modifie : boolean;
                      var CInterdit : string) : integer;
function ControleLong (var Donnee : string; LongMax : integer;
                       Fixe, Modifie : Boolean; Nature : string;
                       var CInterdit : string) : integer;
function ControleValeur (Donnee : string; TLexiqueD : TOB) : integer;
function ControleForme (Segment : string;var Donnee : string; TLexiqueD : TOB;
                        Modifie : boolean; var CInterdit : string) : integer;
procedure ControleSegment (Erreur : TControle; Donnee : string; TLexiqueD : TOB;
                           Modifie : boolean; NoErr : integer=0);
function RenvoieErreur (NoErr : integer; CInterdit : string) : string;
procedure ControleTOB (TDAD : TOB; Suppr, Modifie : boolean);
function GetDADSUNature (TD2:string):string;

var
ErrorDADSU : integer;

TLexique : TOB;
TErreur : TOB;

implementation


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/08/2002
Modifié le ... : 10/10/2006
Description .. : Vérification de la donnée en fonction de la nature issue du
Suite ........ : champ PDL_DADSNATURE (Table DADSLEXIQUE)
Suite ........ : X : AlphaNumerique
Suite ........ : N : Numérique
Suite ........ : D : Date
Suite ........ : I : Identités
Suite ........ : @ : Adresse mail
Suite ........ : 1 : Complément d'adresse et Nature et nom de la voie
Suite ........ : 2 : Numéro dans la voie
Suite ........ : 3 : Bis ou Ter
Suite ........ : 4 : Code INSEE de la commune
Suite ........ : 5 : Nom de la commune
Suite ........ : 6 : Code postal
Suite ........ : 7 : Bureau distributeur ou commune
Suite ........ : 8 : Nom pays en clair
Suite ........ : 9 : Commune ou localité de naissance
Suite ........ : 0 : Code risque A.T.
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
function ControleCar (var Donnee : string; Nature : string; Modifie : boolean;
                      var CInterdit : string) : integer;
var
Ch: Char;
apostrophe, arobase, espace, Interdit, L, Point, tiret : Integer;
Source : PChar;
DateVal, kepabon : boolean;
begin
result:= 0;
kepabon:= True;
apostrophe:= 0;
espace:= 0;
Point:= 0;
tiret:= 0;
Interdit:= 0;  //PT16-3
CInterdit:= '';        //PT17-1

if (Nature = 'D') then
   begin
   DateVal:= IsValidDate (Copy(Donnee,1,2)+DateSeparator+Copy(Donnee,3,2)+
                          DateSeparator+Copy(Donnee,5,4));
   if (DateVal = FALSE) then
      result:= 6;
   end
else
   begin
   if (Nature='1') then
      begin
      Interdit:= Pos (',', Donnee);
      while (Interdit <> 0) do
            begin
            Delete (Donnee, Interdit, 1);
            Interdit:= Pos (',', Donnee);
            end;
      end;

   if ((Nature='7') or (Nature='8') or (Nature='9')) then
      Donnee:= PGUpperCase (Donnee);

   L:= Length(Donnee);
   Source:= Pointer (Donnee);
   while L <> 0 do
         begin
         Ch:= Source^;

         if ((Nature = 'X') or (Nature = 'I') or (Nature = '@') or
            (Nature = '1') or (Nature = '2') or (Nature = '3') or
            (Nature = '4') or (Nature = '5') or (Nature = '6') or
            (Nature = '7') or (Nature = '8') or (Nature = '9') or
            (Nature = '0') or (Nature = 'A')) then
            begin
            if (Modifie = True) then
               begin
               if (Ch in ['*',';']) then
                  begin
                  Ch:= ' ';
                  Source^:= Ch;
                  end;
               if (Ch in ['á','æ','Á','À','Â','Ã','Ä','Å','Æ','Ç','É','È','Ê',
                          'Ë','í','ì','Í','Ì','Î','Ï','ó','ò','õ','ö','Ó','Ò',
                          'Ô','Õ','Ö','ú','Ú','Ù','Û','Ü','ÿ','Ÿ']) then
                  begin
                  Ch:= PGSupprimeaccent (Ch);
                  Source^:= Ch;
                  end;
               end;
            if (Nature = 'X') then
               begin
               if not(Ch in [' ','"','&','''','(',')','+',',','-','.','/',
                             '0'..'9',':','=','@','A'..'Z','_','`','a'..'z','"',
                             '°','"','Ñ','à','â','ç','è','é','ê','ë','î','ï',
                             'ñ','ô','ù','û','ü']) then
                  result:= 1;
               end
            else
            if (Nature = 'I') then
               begin
               if not(Ch in [' ','''','-','A'..'Z','a'..'z','à','â','ç','è','é',
{PT16-2
                             'ê','ë','î','ï','ô']) then
}
                             'ê','ë','î','ï','ô','ù','û','ü']) then
//FIN PT16-2
                  result:= 1;
               if (Ch in [' ','''','-']) then
                  begin
                  if (tiret=2) then
                     result:= 1
                  else
                     begin
{PT16-3
                     if (Ch in ['-']) then
                        begin
                        tiret:= tiret+1;
                        end
                     else
                        begin
                        tiret:= 0;

                        if (Ch in [' ']) then
                           begin
                           if (espace=1) then
                              result:= 1
                           else
                              espace:= 1;
                           end
                        else
                           espace:= 0;

                        if (Ch in ['''']) then
                           begin
                           if (apostrophe=1) then
                              result:= 1
                           else
                              apostrophe:= 1;
                           end
                        else
                           apostrophe:= 0;
                        end;
}
                     if (Ch in ['-']) then
                        begin
                        if ((Interdit<>0) or ((tiret<>0) and (tiret<>1))) then
                           result:=1;
                        end
                     else
                     if (Ch in [' ', '''']) then
                        begin
                        if ((Interdit<>0) or (tiret<>0)) then
                           result:=1;
                        end;

                     if (result<>1) then
                        begin
                        if (Ch in [' ', '''']) then
                           Interdit:= Interdit+1
                        else
                        if (Ch in ['-']) then
                           tiret:= tiret+1;
                        end;
//FIN PT16-3
                     end;

{PT16-2
                  if ((L=Length(Donnee)) or (L=1)) then
}
                  if (L=1) then
//FIN PT16-2
                     result:= 1;
//PT16-2
                  if ((Ch in [' ','-']) and (L=Length (Donnee))) then
                     result:= 1;
//FIN PT16-2
                  end
               else
                  begin
{PT16-2
                  apostrophe:= 0;
                  espace:= 0;
}
                  Interdit:= 0;
//FIN PT16-2
                  tiret:= 0;
                  end;
               end
            else
            if (Nature = '@') then
               begin
               if not(Ch in ['-','.','0'..'9','@','A'..'Z','_','a'..'z']) then
                  result:= 1;
               if (L=Length(Donnee)) then
                  begin
                  arobase:= Pos ('@', Donnee);
                  Point:= Pos('.', Copy (Donnee, arobase, L-arobase));
                  if (Point = 0) then
                     result:= 13;
                  end;
               end
            else
            if (Nature = '1') then
               begin
{PT16-1
               if not(Ch in [' ','''','-','.','0'..'9','A'..'Z','a'..'z','à',
                             'â','ç','è','é','ê','ë','î','ï','ô']) then
}
               if not(Ch in [' ','''','-','.','0'..'9','A'..'Z','a'..'z','à',
                             'â','ç','è','é','ê','ë','î','ï','ô','ù','û',
                             'ü']) then
//FIN PT16-1
                  result:= 1;

//PT16-3
               if (Ch in [' ', '''','-','.']) then
                  begin
                  if (Interdit=0) then
                     begin
                     Interdit:= Interdit+1;
                     if (Ch in ['.']) then
                        Point:= Point+1
                     else
                        Point:= 0;
                     end
                  else
                  if ((Point=1) and (Ch in [' '])) then
                     Interdit:= Interdit+1
                  else
                     result:=1;

                  if ((L=Length (Donnee)) or (L=1)) then
                     result:= 1;
                  end
               else
                  begin
                  Interdit:= 0;
                  Point:= 0;
                  end;
//FIN PT16-3
               end
            else
            if (Nature = '2') then
               begin
{PT16-1
               if not(Ch in ['-','0'..'9','A'..'Z','a'..'z','à']) then
}
               if not(Ch in ['0'..'9']) then
//FIN PT16-1
                  result:= 1;
               end
            else
            if (Nature = '3') then
               begin
{PT16-1
               if not(Ch in ['B','C','Q','T','b','c','q','t']) then
}
               if not(Ch in ['A'..'Z','a'..'z']) then
//FIN PT16-1
                  result:= 1;
               end
            else
            if (Nature = '4') then
               begin
               if not(Ch in ['0'..'9','A','B']) then
                  result:= 1;
               end
            else
            if (Nature = '5') then
               begin
{PT16-1
               if not(Ch in [' ','''','-','0'..'9','A'..'Z','a'..'z','à','â',
                             'ç','è','é','ê','î','ï','ô']) then
}
               if not(Ch in [' ','0'..'9','A'..'Z','a'..'z']) then
//FIN PT16-1
                  result:= 1;
               end
            else
            if (Nature = '6') then
               begin
               if not(Ch in ['0'..'9','A'..'Z','a'..'z']) then
                  result:= 1;
               end
            else
{PT16-1
            if ((Nature = '7') or (Nature = '9')) then
               begin
}
            if (Nature = '7') then
               begin
//FIN PT16-1
{PT16-1
               if not(Ch in [' ','''','-','0'..'9','A'..'Z']) then
}
               if not(Ch in [' ','A'..'Z']) then
//FIN PT16-1
                  result:= 1;

//PT16-3
               if (Ch in [' ']) then
                  begin
                  if (Interdit<>0) then
                     result:=1
                  else
                     Interdit:= Interdit+1;

                  if ((L=Length (Donnee)) or (L=1)) then
                     result:= 1;
                  end
               else
                  Interdit:= 0;
//FIN PT16-3
               end
            else
            if (Nature = '8') then
               begin
               if not(Ch in [' ','''','(',')','-','A'..'Z']) then
                  result:= 1;

//PT16-3
               if (Ch in [' ', '''','-','(',')']) then
                  begin
                  if (Interdit<>0) then
                     result:=1
                  else
                     Interdit:= Interdit+1;

                  if ((L=Length(Donnee)) or (L=1)) then
                     result:= 1;
                  end
               else
                  Interdit:= 0;
//FIN PT16-3
               end
//PT16-1
            else
            if (Nature = '9') then
               begin
{PT16-3
               if not(Ch in [' ','''','-','0'..'9','A'..'Z']) then
}
               if not(Ch in [' ','0'..'9','A'..'Z','a'..'z']) then
                  result:= 1;
//PT16-3
               if (Ch in [' ']) then
                  begin
                  if (Interdit<>0) then
                     result:=1
                  else
                     Interdit:= Interdit+1;

                  if ((L=Length (Donnee)) or (L=1)) then
                     result:= 1;
                  end
               else
                  Interdit:= 0;
//FIN PT16-3
               end
//FIN PT16-1
            else
            if (Nature = '0') then
               begin
               if not(Ch in ['0'..'9','A'..'Z']) then
                  result:= 1;
               end
//PT14
            else
            if (Nature = 'A') then
               begin
               if not(Ch in ['A'..'Z']) then
                  result:= 1;
               end;
//FIN PT14
            end;
         if (Nature = 'N') then
            begin
            if not(Ch in ['0'..'9']) then
               result:= 2;
            if ((L=Length(Donnee)) and (L<>1) and (Ch in ['0'])) then
               result:= 1;
            end;

{PT16-3
         if not(Ch in ['.','(','&',')','-',',','@','=','"','°',' ']) then
}
         if not(Ch in ['.','(','&',')','-',',','@','=','"','°',' ','''','`','+',
                       '/',':','_','"','"']) then
//FIN PT16-3
            kepabon:= False;

         if ((Ch in [' ']) and ((L=Length(Donnee)) or (L=1))) then
            result:= 1;

         if (result <> 0) then
            break;
         Inc (Source);
         Dec(L);
         end;
   if (kepabon = True) then
      result:= 1;

//PT17-1
   if (result=1) then
      begin
      CInterdit:= Ch;
      if (CInterdit=' ') then
         CInterdit:= 'espace';
      end;
//FIN PT17-1
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/08/2002
Modifié le ... :   /  /
Description .. : Vérification de la longueur de la donnée
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
function ControleLong (var Donnee : string; LongMax : integer;
                       Fixe, Modifie : boolean; Nature : string;
                       var CInterdit : string) : integer;
var
Compar, L : Integer;
begin
L:= Length(Donnee);
if (L = 0) then
   begin
   result:= 7;
   exit;
   end;
{PT17-1
result:= ControleCar(Donnee, Nature, Modifie);
}
result:= ControleCar(Donnee, Nature, Modifie, CInterdit);
//FIN PT17-1
if (result = 0) then
   begin
   Compar:= LongMax-L;
   if (Compar<0) then
      result:= 3
   else
      begin
      if ((Compar>0) and (Fixe=TRUE)) then
         result:= 4
      else
         result:= 0;
      end;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/08/2002
Modifié le ... :   /  /
Description .. : Contrôle de la donnée par rapport à la valeur comprise dans
Suite ........ : PDL_DADSVALEUR
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
function ControleValeur(Donnee : string; TLexiqueD : TOB) : integer;
var
Abrege, Chaine, ChampLexique, DonneeLexique, IntervD : string;
IntervF, NomTablette, St, TypeC : string;
begin
result:= 0;
ChampLexique:= 'PDL_DADSVALEUR';

if ((TLexiqueD.GetValue('PDL_DADSUSAGE') = 'O') and (Donnee='')) then
   result:= 11;

if (TLexiqueD.GetValue(ChampLexique) = null) then
   exit;
DonneeLexique:= TLexiqueD.GetValue(ChampLexique);

TypeC:= Trim(ReadTokenPipe(DonneeLexique,'-'));

if (TypeC = 'T') then
   begin
   NomTablette:= Trim(ReadTokenPipe(DonneeLexique,';'));
   Abrege:= Trim(ReadTokenPipe(DonneeLexique,'.'));
   if (Abrege = 'TRUE') then
      St:= RechDom(NomTablette, Donnee, TRUE)
   else
      St:= RechDom(NomTablette, Donnee, FALSE);

   if ((St = '') or (St = 'Error')) then
      result:= 8;
   end;

if (TypeC = 'I') then
   begin
   IntervD:= Trim(ReadTokenPipe(DonneeLexique,';'));
   IntervF:= Trim(ReadTokenPipe(DonneeLexique,'.'));
   if ((Donnee < IntervD) or (Donnee > IntervF)) then
      result:= 9;
   end;

if (TypeC = 'C') then
   begin
   Chaine:= Trim(ReadTokenPipe(DonneeLexique,'.'));
   if (Donnee <> Chaine) then
      result:= 10;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/08/2002
Modifié le ... :   /  /    
Description .. : Récupération du lexique pour ce segment
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
function ControleForme (Segment : string;var Donnee : string; TLexiqueD : TOB;
                        Modifie : boolean; var CInterdit : string) : integer;
var
Buffer, Nature : string;
LongMax : integer;
Fixe : boolean;
begin
result:= 5;

if (TLexiqueD = nil) then
   exit
else
   begin
   LongMax:= TLexiqueD.GetValue('PDL_DADSLONGUEUR');
   Buffer:= TLexiqueD.GetValue('PDL_DADSFIXE');
   if (Buffer = 'X') then
      Fixe:= True
   else
      Fixe:= False;
   Nature:= TLexiqueD.GetValue('PDL_DADSNATURE');
{PT17-1
   result:= ControleLong(Donnee, LongMax, Fixe, Modifie, Nature);
}
   result:= ControleLong(Donnee, LongMax, Fixe, Modifie, Nature, CInterdit);
//FIN PT17-1
   if (result=0) then
      result:= ControleValeur (Donnee, TLexiqueD);
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 18/04/2006
Modifié le ... :   /  /
Description .. : Fonction permettant de contrôler les données issues de la base
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure ControleSegment (Erreur : TControle; Donnee : string; TLexiqueD : TOB;
                           Modifie : boolean; NoErr : integer=0);
var
reponse : integer;
CInterdit, DadsNom, Libelle : string;
TErreurDetail : TOB;
begin
Libelle:= '';
Debug('Paie PGI/Contrôle DADS-U : '+Erreur.Segment+' | '+Donnee);

TErreurDetail:= TErreur.FindFirst (['PSU_SALARIE','PSU_TYPE','PSU_ORDRE',
                                   'PSU_DATEDEBUT','PSU_DATEFIN','PSU_SEGMENT',
                                   'PSU_EXERCICEDADS'],
                                   [Erreur.Salarie,Erreur.TypeD,Erreur.Num,
                                   Erreur.DateDeb,Erreur.DateFin,Erreur.Segment,
                                   Erreur.Exercice],True);

if (NoErr<>0) then
   reponse:= NoErr
else
{PT17-1
   reponse:= ControleForme (Erreur.Segment, Donnee, TLexiqueD, Modifie);
}
   reponse:= ControleForme (Erreur.Segment, Donnee, TLexiqueD, Modifie,
                            CInterdit);
//FIN PT17-1

if ((reponse <> 0) and (TErreurDetail=nil)) then
   begin
   ErrorDADSU:= ErrorDADSU+1;
   DadsNom:= TLexiqueD.GetValue ('PDL_DADSNOM');
{PT17-1
   Erreur.Explication:= RenvoieErreur (reponse);
}
   Erreur.Explication:= RenvoieErreur (reponse, CInterdit);
//FIN PT17-1
   if (reponse <> 12) then
      Erreur.CtrlBloquant:= True
   else
      Erreur.CtrlBloquant:= False;
   EcrireErreur(Erreur, TErreur, DadsNom);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/07/2006
Modifié le ... :   /  /
Description .. : Fonction qui renvoie le libellé de l'erreur
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
function RenvoieErreur (NoErr : integer; CInterdit : string) : string;
begin
case NoErr of
{PT17-1
     1 : result:= 'Caractère interdit';
}
     1 : result:= 'Caractère '+CInterdit+' interdit';
//FIN PT17-1
     2 : result:= 'Caractère non numérique';
     3 : result:= 'Chaine trop longue';
     4 : result:= 'Chaine de longueur fixe non respectée';
     5 : result:= 'Rubrique inconnue';
     6 : result:= 'Date invalide';
     7 : result:= 'Donnée vide';
     8 : result:= 'La valeur n''est pas dans la tablette';
     9 : result:= 'La valeur n''est pas dans l''intervalle autorisé';
     10 : result:= 'La valeur n''est pas égale à la valeur autorisée';
     11 : result:= 'La valeur est obligatoire mais n''est pas renseignée';
     12 : result:= 'Rubrique obligatoire pour un salarié né en France';
     13 : result:= 'Adresse mail incorrecte';
     end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/08/2002
Modifié le ... :   /  /
Description .. : Fonction principale de contrôle et d'écriture des données
Suite ........ : issues de DADSDETAIL vers le fichier temporaire
Mots clefs ... : PAYE,PGDADSU
*****************************************************************}
procedure ControleTOB(TDAD : TOB; Suppr, Modifie : boolean);
var
TDADD, TLexiqueD : TOB;
Buffer, LibSegLex, NomSegDAD, NomSegLex, NomTable : string;
Salarietraite : string;
ErreurUsage, i : integer;
AControler : boolean;
ErreurDADSU : TControle;
ValeurSegment : Array[0..33] of string;
begin
ErreurUsage:= 0;
AControler:= True;
NomSegDAD:= 'PDS_SEGMENT';
NomSegLex:= 'PDL_DADSSEGMENT';
LibSegLex:= 'PDL_DADSNOM';
ValeurSegment[0]:= 'S41.G01.00.001';
ValeurSegment[1]:= 'S41.G01.01.001';
ValeurSegment[2]:= 'S30.G01.00.001';
ValeurSegment[3]:= 'S70.G01.01.001';
ValeurSegment[4]:= 'S80.G01.00.001.001';
//PT16-3
ValeurSegment[5]:= 'S80.G62.05.001';
ValeurSegment[6]:= 'S80.G62.10.001';
//FIN PT16-3
ValeurSegment[7]:= 'S20.G01.00.001';
ValeurSegment[8]:= 'S41.G01.02.001';
ValeurSegment[9]:= 'S41.G01.03.001';
ValeurSegment[10]:= 'S41.G01.04.001';
ValeurSegment[11]:= 'S41.G01.05.001';
ValeurSegment[12]:= 'S41.G01.06.001';
ValeurSegment[13]:= 'S41.G02.00.008';
ValeurSegment[14]:= 'S44.G01.00.001';
ValeurSegment[15]:= 'S45.G01.00.001';
ValeurSegment[16]:= 'S45.G01.01.001';
ValeurSegment[17]:= 'S46.G01.00.001';
ValeurSegment[18]:= 'S46.G01.01.001';
ValeurSegment[19]:= 'S46.G01.02.001.001';
ValeurSegment[20]:= 'S42.G01.00.001';
ValeurSegment[21]:= 'S42.G02.00.001';
ValeurSegment[22]:= 'S41.G30.35.001.001';         //PT17-2
//PT16-3
ValeurSegment[23]:= 'S41.G30.10.001';
ValeurSegment[24]:= 'S41.G30.11.001.001';
ValeurSegment[25]:= 'S41.G30.15.001';
ValeurSegment[26]:= 'S41.G30.20.001';
ValeurSegment[27]:= 'S41.G30.25.001';
ValeurSegment[28]:= 'S41.G30.30.001.001';
//FIN PT16-3
ValeurSegment[29]:= 'S51.G01.00.001';
ValeurSegment[30]:= 'S66.G01.00.001';
ValeurSegment[31]:= 'S10.G01.01.001.001';
ValeurSegment[32]:= 'S10.G01.00.001.001';
ValeurSegment[33]:= 'S90.G01.00.001';

ErreurDADSU.TypeD:= TypeD;
ErreurDADSU.Exercice:= PGExercice;

TDADD:= TDAD.FindFirst([''], [''], True);
if (TDADD <> nil) then
   NomTable:= TDADD.NomTable;

TLexiqueD:= TLexique.FindFirst([''], [''], False);

While (TDADD <> nil) do
      begin
      while ((TDADD <> nil) and (TDADD.FieldExists (NomSegDAD)=False)) do
            TDADD:= TDAD.FindNext ([''], [''], True);
{PT16-3
      for i := 0 to 24 do
}
{PT17-2
      for i := 0 to 32 do
}
      for i := 0 to 33 do
          begin
          if (TDADD.GetValue(NomSegDAD) = ValeurSegment[i]) then
             begin
             TLexiqueD:= TLexique.FindFirst([''],[''],FALSE);
             while (TLexiqueD.GetValue(NomSegLex) <> ValeurSegment[i]) do
                   TLexiqueD:= TLexique.FindNext([''],[''],FALSE);
             if (TDADD.GetValue(NomSegDAD) = 'S30.G01.00.001') then
                Salarietraite:= TDADD.GetValue('PDS_SALARIE');
             AControler:= True;
             break;
             end;
          end;

      Buffer:= Copy(TDADD.GetValue(NomSegDAD),1,10);
      if (Buffer = 'S70.G01.00') then
         begin
         TLexiqueD:= TLexique.FindFirst([''],[''],FALSE);
         while (TLexiqueD.GetValue(NomSegLex) <> TDADD.GetValue(NomSegDAD)) do
               TLexiqueD:= TLexique.FindNext([''],[''],FALSE);
         Salarietraite:= TDADD.GetValue('PDS_ORDRE');
         AControler:= True;
         end;

      if (AControler = True) then
         begin
         if (TLexiqueD.GetValue('PDL_DADSUSAGE') = 'O') then
            begin
            if (TDADD.GetValue(NomSegDAD) <> TLexiqueD.GetValue(NomSegLex)) then
               begin
               ErreurUsage:= -1;
               ErreurDADSU.Salarie:= TDADD.GetValue('PDS_SALARIE');
               if (TDADD.FieldExists ('PDS_LIBELLE')=True) then
                  ErreurDADSU.Libelle:= TDADD.GetValue('PDS_LIBELLE');
               ErreurDADSU.Num:= TDADD.GetValue('PDS_ORDRE');
               ErreurDADSU.DateDeb:= TDADD.GetValue('PDS_DATEDEBUT');
               ErreurDADSU.DateFin:= TDADD.GetValue('PDS_DATEFIN');
{PT15
               ErreurDADSU.Segment:= TDADD.GetValue('PDS_SEGMENT');
               ControleSegment (ErreurDADSU, TDADD.GetValue('PDS_DONNEE'),
                                TLexiqueD, Modifie);
}
{Retour en arrière pour diffusion du 20/12/2007
               ErreurDADSU.Segment:= TLexiqueD.GetValue(NomSegLex);
               ControleSegment (ErreurDADSU, '', TLexiqueD, Modifie, 11);
}
               ErreurDADSU.Segment:= TDADD.GetValue('PDS_SEGMENT');
               ControleSegment (ErreurDADSU, TDADD.GetValue('PDS_DONNEE'),
                                TLexiqueD, Modifie);
//FIN retour en arrière
//FIN PT15
               end
            else
               begin
               ErreurUsage:= 0;
               ErreurDADSU.Salarie:= TDADD.GetValue('PDS_SALARIE');
               if (TDADD.FieldExists ('PDS_LIBELLE')=True) then
                  ErreurDADSU.Libelle:= TDADD.GetValue('PDS_LIBELLE');
               ErreurDADSU.Num:= TDADD.GetValue('PDS_ORDRE');
               ErreurDADSU.DateDeb:= TDADD.GetValue('PDS_DATEDEBUT');
               ErreurDADSU.DateFin:= TDADD.GetValue('PDS_DATEFIN');
               ErreurDADSU.Segment:= TDADD.GetValue('PDS_SEGMENT');
               ControleSegment (ErreurDADSU, TDADD.GetValue('PDS_DONNEE'),
                                TLexiqueD, Modifie);
               end;
            end
         else
            begin
            if (TLexiqueD.GetValue('PDL_DADSUSAGE') = 'P') then
               begin
               if (TDADD.GetValue(NomSegDAD) <> TLexiqueD.GetValue(NomSegLex)) then
                  begin
                  ErreurDADSU.Salarie:= TDADD.GetValue('PDS_SALARIE');
                  if (TDADD.FieldExists ('PDS_LIBELLE')=True) then
                     ErreurDADSU.Libelle:= TDADD.GetValue('PDS_LIBELLE');
                  ErreurDADSU.Num:= TDADD.GetValue('PDS_ORDRE');
                  ErreurDADSU.DateDeb:= TDADD.GetValue('PDS_DATEDEBUT');
                  ErreurDADSU.DateFin:= TDADD.GetValue('PDS_DATEFIN');
                  ErreurDADSU.Segment:= TDADD.GetValue('PDS_SEGMENT');
                  ControleSegment (ErreurDADSU, TDADD.GetValue('PDS_DONNEE'),
                                   TLexiqueD, Modifie, 12);
                  end;
               end;
            if (TDADD.GetValue(NomSegDAD) <> TLexiqueD.GetValue(NomSegLex)) then
               ErreurUsage:= -1
            else
               begin
               ErreurUsage:= 0;

               ErreurDADSU.Salarie:= TDADD.GetValue('PDS_SALARIE');
               if (TDADD.FieldExists ('PDS_LIBELLE')=True) then
                  ErreurDADSU.Libelle:= TDADD.GetValue('PDS_LIBELLE');
               ErreurDADSU.Num:= TDADD.GetValue('PDS_ORDRE');
               ErreurDADSU.DateDeb:= TDADD.GetValue('PDS_DATEDEBUT');
               ErreurDADSU.DateFin:= TDADD.GetValue('PDS_DATEFIN');
               ErreurDADSU.Segment:= TDADD.GetValue('PDS_SEGMENT');
               ControleSegment (ErreurDADSU, TDADD.GetValue('PDS_DONNEE'),
                                TLexiqueD, Modifie);
               end;
            end;
         end;
      if (AControler = True) then
         begin
         if (ErreurUsage = 0) then
            begin
            if (Suppr=True) then
               begin
               MoveCurProgressForm ('Contrôle '+
                                    RechDom ('PGORIGINECRIT',
                                    Copy (TDADD.GetValue('PDS_SEGMENT'), 1, 3),
                                    False)+' '+TDADD.GetValue ('PDS_LIBELLE'));
               MoveCur(False);
               FreeAndNil (TDADD);
               end;
            TDADD:= TDAD.FindNext ([''], [''], True);
            end;
         TLexiqueD:= TLexique.FindNext([''], [''], False);
         end;
      AControler:= True;

      if (TLexiqueD = nil) then
         begin
         if (TDADD <> nil) then
            Debug ('Paie PGI/Contrôle DADS-U : '+TDADD.GetValue(NomSegDAD)+
                   ' : Segment non trouvé');
         break;
         end;
      end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 26/04/2004
Modifié le ... :   /  /
Description .. : Recherche des éléments en fonction de la nature
Suite ........ : DADS-U
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
function GetDADSUNature(TD2:string) : string;
begin
if (TD2='01') then
   result:=' AND PDL_DADSUCOMPLETE="P"'
else
{PT18
if (TD2='02') then
}
if ((TD2='02') or (TD2='12')) then
//FIN PT18
   result:=' AND PDL_DADSUTDS="P"'
else
if (TD2='03') then
   result:=' AND PDL_DADSUIRCIP="P"'
else
if (TD2='04') then
   result:=' AND PDL_DADSUCCPBTP="P"'
else
if (TD2='05') then
   result:=' AND PDL_DADSUNEANT="P"'
else
if (TD2='07') then
   result:=' AND PDL_DADSUIRC="P"'
else
if (TD2='08') then
   result:=' AND PDL_DADSUIP="P"'
else
if (TD2='09') then
   result:=' AND PDL_DADSUASSUR="P"';
end;

end.
