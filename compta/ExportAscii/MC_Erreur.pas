unit mc_erreur;

interface


//Gestipon erreurs
Type TMC_Err=record
     code,Extra : integer;
     Libelle    : String;
     end;

function MC_MsgErrDefaut ( Code : Integer ; Extra : integer = -1  ) : string ;

implementation

uses Hent1, HMsgBox, Sysutils ;

function MC_MsgErrDefaut ( Code : Integer ; Extra : integer = -1  ) : string ;
Begin
Result:='' ;
Code:=abs(Code) ;
case Code of
     1 : Result:='Une erreur inconnue est survenue.' ;
     2 : Result:='Le code du dispositif (%s) est inconnu comme '+MC_MsgErrDefaut(951+Extra)+'.' ;
     3 : Result:='Aucun dispositif n''est paramétré comme '+MC_MsgErrDefaut(951+Extra)+'.' ;
     4 : Result:='Il faut utiliser la méthode ''create'' en indiquant le code de terminal.' ;
     5 : Result:='Aucun port de communication n''est paramétré ('+MC_MsgErrDefaut(951+Extra)+').' ;
     6 : Result:='Aucun paramètre pour le port n''a été renseigné.' ;
     7 : Result:='Le materiel demandé est inconnu ('+MC_MsgErrDefaut(951+Extra)+').' ; //XMG
     8 : Result:='Il y a un périphérique déjà ouvert ('+MC_MsgErrDefaut(951+Extra)+').' ; //XMG
  //nom dispositif (950<=Code<960)
   950 : Result:='Inconnue' ;
   951 : result:='imprimante' ;
   952 : Result:='imprimante de chèques' ;
   953 : Result:='carte bancaire' ;
   954 : Result:='afficheur' ;
   955 : Result:='tiroir' ;
   //Caption fenêtres de messages (960<=Code<1000)
   960 : Result:='Code type dispositif trompé' ;
   961 ,
   962 : Result:='' ;
   963 : Result:='' ;
   964 : Result:='Gestion de l'''+MC_MsgErrDefaut(951+3) ;
   965 : Result:='' ;
  // Messages Impression (1000<=code<2000)
  1001 : Result:=#13+' La requête est annulée.' ;
  1002 : Result:='Erreur interne.' ;
  1003 : Result:='Le modèle n''existe pas.' ; //XMG 30/03/04 'L''état n''existe pas.' ;
  1004 : Result:='La mise en page n''est pas définie !' ;
  1005 : Result:='La requête est absente !' ;
  1006 : Result:='La version du fichier à importer est incompatible ou inconnue.'+#13 ;
  1007 : Result:='Erreur à l''importation du modèle. fichier %s' ;
  1008 : Result:='Erreur à l''exportation du modèle. fichier %s' ;
  1009 : Result:='Vous devez paramétrer un port de sortie.' ;
  1010 : Result:='Vous devez utiliser une imprimante texte.' ;
  1011 : Result:='Erreur interne du paramétrage de l''imprimante texte.' ;
  1012 : Result:='Une erreur est survenue lors de l''ouverture du fichier.' ;
  1013 : Result:='Impossible d''éteindre l''imprimante texte.' ;
  1014 : Result:='Une erreur est survenue lors de la fermeture du fichier.' ;
  1015 : Result:='L''imprimante est en cours de réinitialisation.' ;
  1016 : Result:='L''imprimante est hors service.#13 Appuyez sur ''ON-LINE''.' ;
  1017 : Result:='Le capot de l''imprimante est ouvert.' ;
  1018 : Result:='Erreur fatale.#13 Veuillez éteindre l''imprimante.' ;
  1019 : Result:='L''imprimante est en cours de réinitialisation.#13 Veuillez patienter s''il vous plaît.' ;
  1020 : Result:='Veuillez vérifier l''état de l''imprimante.' ;
  1021 : Result:='Il n''y a plus de papier dans l''imprimante.' ;
  1022 : Result:='Pensez à changer le papier de l''imprimante !' ;
  1023 : Result:='Veuillez vérifier si l''imprimante est activée.' ;
  1024 : Result:='Erreur d''écriture  (code %d).'+#13+' %s'+#13+' Souhaitez-vous relancer l''impression ?' ;
  1025 : Result:='Souhaitez-vous relancer l''impression ?' ;
  1026 : Result:='Préparation de l''aperçu…' ;
  1027 : Result:='Processus en cours...' ; //XMG 30/03/04 'Impression en cours…' ;
  1028 : Result:='Erreur interne inconnue.' ;
  1029 : Result:='La requête est incorrecte.' ;
  1030 : Result:='Aucun enregistrement à imprimer.' ;
  1031 : Result:='Impression annulée par l''utilisateur.' ;
  1032 : Result:='Nouveau modèle' ; //XMG 30/03/04 'Nouvel état' ;
  1033 : Result:='Entête' ;
  1034 : Result:='Détail' ;
  1035 : Result:='Pied' ;
  1036 : Result:='Barre de menu' ;
  1037 : Result:='Sous-détail ' ;
  1038 : Result:='Fichiers texte|*.txt|Tous les fichiers|*.*' ;
  1039 : Result:='Nom du fichier à exporter' ;
  1040 : Result:='Annuler' ;
  1041 : Result:='Variables système' ;
  1042 : Result:='Variables spéciales' ;
  1043 : Result:='Couper partiellement le ticket' ;
  1044 : Result:='Couper totalement le ticket' ;
  1045 : Result:='Ouvrir le tiroir' ;
  1046 : Result:='Copyright' ;
  1047 : Result:='Date du jour' ;
  1048 : Result:='Date d''entrée' ;
  1049 : Result:='Date de version logiciel' ;
  1050 : Result:='Nom du logiciel' ;
  1051 : Result:='Heure' ;
  1052 : Result:='Nom utilisateur' ;
  1053 : Result:='Version programme' ;
  1054 : Result:='Titre programme' ;
  1055 : Result:='Code utilisateur' ;
  1056 : Result:='Nature' ;
  1057 : Result:='Etat' ;
  1058 : Result:=#13+' Désirez-vous suspendre l''impression ?' ;
  1059 : Result:='Vous devez paramétrer les bandes.' ;
  1060 : Result:='Ce nom de champ existe déjà.' ;
  1061 : Result:='Le code de l''imprimante (%s) est inconnu.' ;
  1062 : Result:='Il faut utiliser la méthode ''create'' en indiquant le code d''imprimante.' ;
  1063 : Result:='Démarre l''impression en mode "facturette"' ;
  1064 : Result:='Termine l''impression en mode "facturette"' ;
  1065 : Result:='Démarre l''impression en mode "validation"' ;
  1066 : Result:='Termine l''impression en mode "validation"' ;
  1067 : Result:='Placez le document pour l''impression.' ;
  1068 : Result:='Démarre l''impression en mode chèque' ;
  1069 : Result:='Termine l''impression en mode chèque' ;
  1070 : Result:='Sous-entête ' ;
  1071 : Result:='Sous-Pied ' ;
  1072 : Result:='Voulez-vous créer un état à partir de ce modèle ?' ;
  1073 : Result:='<Aucune>' ;
  1074 : Result:='Entête rupture sur ' ;
  1075 : Result:='Pied rupture sur ' ;
  1076 : Result:='Le modèle a été modifié, désirez-vous l''enregistrer ?' ;
  1077 : Result:='   '+TitreHalley+traduireMemoire(' - Version de démonstration -')+Titrehalley ;
  1078 : Result:='Imprime l''image chargée dans l''imprimante' ;
  1079 : Result:='Augmente l''interligne' ;
  1080 : Result:='Revient à l''interligne standard' ;
  1081 : Result:='Confirmez-vous l''annulation  de l''impression en cours ?' ;
  1082 : Result:='Le modèle que vous êtes en train d''enregistrer a été adapté depuis une version antérieure.#13'
                +' N''oubliez pas de faire une sauvegarde du modèle (à l''ancien format) si vous desirez l''utiliser sur une version antérieure du moteur d''impression.#13'
                +' Désirez-vous continuer l''enregistrement?' ;
  1083 : Result:='Enregistrement annulé par l''utilisateur.' ;
  1084 : Result:='N''oubliez pas que ce modèle ne peut être importé que sur une version égal ou supérieure du produit.#13'
                +' Désirez-vous continuer l''enregistrement?' ;
  // Messages Afficheur (2000<=Code<3000)
  2000 : Result:=MC_MsgErrDefaut(1) ;
  2001 : Result:=MC_MsgErrDefaut(2) ;
  2002 : Result:=MC_MsgErrDefaut(3,3) ;//XMG
  2003 : Result:=MC_MsgErrDefaut(4) ;
  2004 : Result:=MC_MsgErrDefaut(5,3) ;//XMG
  2005 : Result:='Connexion en cours...' ;
  2006 : Result:='Caisse fermée.' ;
  2007 : Result:='' ;
  2008 : Result:='Ce modèle d''afficheur n''est pas géré par le logiciel.' ;
  2009 : Result:='Erreur interne de paramétrage de l''afficheur externe.' ;
  2010 : Result:='Une erreur est survenue lors de l''ouverture du fichier.' ;
  2011 : Result:='Impossible d''éteindre l''afficheur externe.' ;
  2012 : Result:='Une erreur est survenue lors de la fermeture du fichier.' ;
  2013 : Result:='Erreur d''écriture vers l''afficheur externe (code %d).'+#13+' %s'+#13+' Souhaitez-vous réessayer ?' ;
  2014 : Result:=#13+' Souhaitez-vous annuler l''affichage ?' ;
  //EPSON DM-D105
  // Messages Carte Bancaire (3000<=code<4000)
  3000 : Result:=MC_MsgErrDefaut(1) ;
  3001 : Result:=MC_MsgErrDefaut(2) ;
  3002 : Result:=MC_MsgErrDefaut(3,2) ; //XMG
  3003 : Result:=MC_MsgErrDefaut(4) ;
  3004 : Result:=MC_MsgErrDefaut(5,2) ; //XMG
  3005 : Result:='Le TPE est occupé ou éteint.' ;
  3006 : Result:='Aucun paramètre n''a été modifié.' ;
  3007 : Result:='La transaction a échoué.' ;
  3008 : Result:='Ce modèle de TPE n''est pas géré par le logiciel.' ;
  //CKD+INGENICO
  3100 : Result:='Le TPE est occupé.' ;
  3101 : Result:='La devise est inconnue.' ;
  3102 : Result:='La transaction est acceptée.' ;     //'0'
  3103 : Result:='La transaction est autorisée.' ;    //'1'
  3104 : Result:='Appel phonique.' ;                  //'2'
  3105 : Result:='La transaction est forcée.' ;       //'3'
  3106 : Result:='La transaction est refusée.' ;      //'4'
  3107 : Result:='La carte est interdite.' ;          //'5'
  3108 : Result:='Abandon de l''opérateur.' ;         //'6'
  3109 : Result:='La transaction n''a pas aboutie.' ; //'7'
  3110 : Result:='Opération impossible.' ;            //'8'
  // Messages Tiroir (4000<=Code<5000)
  4000 : Result:=MC_MsgErrDefaut(1) ;
  4001 : Result:=MC_MsgErrDefaut(2) ;
  4002 : Result:=MC_MsgErrDefaut(3,4) ; //XMG
  4003 : Result:=MC_MsgErrDefaut(4) ;
  4004 : Result:=MC_MsgErrDefaut(5,4) ; //XMG
  4005 : Result:='' ;
  4006 : Result:='' ;
  4007 : Result:='' ;
  4008 : Result:='Ce modèle de tiroir n''est pas géré par le logiciel.' ;
  4009 : Result:='Erreur interne due au paramétrage du tiroir.' ;
  4010 : Result:='Une erreur est survenue lors de l''ouverture du fichier.' ;
  4011 : Result:='Impossible de fermer le tiroir.' ;
  4012 : Result:='Une erreur est survenue lors de la fermeture du fichier.' ;
  4013 : Result:='Erreur à l''ouverture du tiroir (code %d).'+#13+' %s'+#13+' Souhaitez-vous réessayer ?' ;
  4014 : Result:=#13+' Souhaitez-vous annuler l''ouverture du tiroir ?' ;

  End ;
if (Result='') and (code<>0) and (V_PGI.VersionDev or V_PGI.SAV) then PGIBox('Le code spécifié '+inttostr(Code)+' n''a aucun texte associé.','MC_Erreur') ;
End ;

end.

