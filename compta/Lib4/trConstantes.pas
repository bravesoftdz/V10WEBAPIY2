unit trConstantes;

interface

const
	// Codes des applications
	CodeModule = 'TRE'; // Trésorerie
	CodeModuleCpt = 'CPT'; // Compta

    
   //BASE DE CALCUL
   B30_365 		= 0 ;
   B30_360 		= 1 ;
   B30E_360 	= 2 ;
   BJUSTE_365 	= 3 ;
   BJUSTE_360 	= 4 ;
   BJUSTE_JUSTE	= 5 ;
	//--------------
   //CALENDRIER
   CalendarWeekBound = 18264; // Valeur de date valide (01/01/1950) ;
		// un petit nombre en plus (<10) indique un jour de la semaine (1=lundi)
   CalendarEasterYear = 1960; // Année pour les dates indiquant un décalage par rapport à Pâques
   CalendarEasterBound = 22068; // Valeur de date valide (01/06/1960) ;
		// un nombre de jours (<6 mois) en plus (ou en moins mais pas utilisé) indique une date relative à Pâques
   CalendarClosedYear = 1970; // Année pour stocker les jours fermés (même que tablette TTJOURFERIE pour les jours qui sont aussi fériés)
	//--------------
   //MESSAGES
   TrMessage: array[0..22] of String = (
{0}   'Un CIB est déjà associé à cette règle.'#13' Confirmez vous l''accès à la modification des critères ?;Q;YN;N;N',
{1}	'Voulez-vous sauvegarder les modifications effectuées ?;Q;YN;N;N',
{2}	'Vous allez supprimer définitivement les informations. Confirmez vous l''opération ?;Q;YN;N;N',
{3}	'Vous devez renseigner le code %%;E;O;O;O',
{4}	'La date de départ est postérieure à la date de fin;E;O;O;O',
{5}	'Le compte général %% n''a aucun RIB. Opération annulée;E;O;O;O',
{6}	'Traitement effectué avec succes;E;O;O;O',
{7}	'Traitement non effectué;E;O;O;O',
{8}	'Vous devez sélectionner un compte pivot pour pouvoir faire un équilibrage automatique;E;O;O;O',
{9}	'Vous devez avoir au moins un compte général;E;O;O;O',
{10}  'Le montant d''un virement doit être une valeur positive;E;O;O;O',
{11}  'Le compte d''origine et celui de destination sont identiques;E;O;O;O',
{12}  'Un virement entre ces deux comptes existe déja, modifiez son montant au lieu de créer un autre virement;E;O;O;O',
{13}  'Le fichier %% existe déja, voulez-vous le remplacer ?;Q;YN;N;N',
{14}  'Vous allez annuler des transactions. Confirmez vous l''opération ?;Q;YN;N;N',
{15}  'Vous allez éditez des relevés. Confirmez vous l''opération ?;Q;YN;N;N',
{16}  'Récupération des conditions paramétrées ?;Q;YN;N;N',
{17}  'Vous allez validez des transactions. Confirmez vous l''opération ?;Q;YN;N;N',
{18}  'Vous allez supprimer le virement de %% du compte $$. Confirmez vous l''opération ?;Q;YN;N;N',
{19}  'Erreur sur la clé RIB.'#13' Récupération de la clé ?;Q;YN;N;N',
{20}  'Confirmez-vous le traitement ?;Q;YN;N;N',
{21}  'Vous allez dénouer des transactions. Confirmez vous l''opération ?;Q;YN;N;N',
{22}  'Vous allez annuler les dénouement des transactions. Confirmez vous l''opération ?;Q;YN;N;N');


   //--------------
   //TABLES
   //--------------
   NoDate			= '01/01/1900'; // Valeur de champ date non saisi
   //TrEcritures
   TE_NUMTRANSAC	=	0	;
    TE_NUMLIGNE		=	1	;
    TE_JOURNAL		=	2	;
    TE_NUMEROPIECE	=	3	;
    TE_EXERCICE		=	4	;
    TE_GENERAL		=	5	;
    TE_CONTREPARTIETR = 55 ;
    TE_CODECIB		=	6	;
    TE_CODEFLUX		=	7	;
    TE_REFINTERNE	=	8	;
    TE_SOCIETE		=	9	;
    TE_QUALIFORIGINE=	10	;
    TE_LIBELLE		=	11	;
    TE_USERCREATION	=	12	;
    TE_USERMODIF	=	13	;
    TE_USERVALID	=	14	;
    TE_USERCOMPTABLE=	15	;
    TE_DATECREATION	=	16	;
    TE_DATEMODIF	=	17	;
    TE_DATEVALID	=	18	;
    TE_DATERAPPRO	=	19	;
    TE_DEVISE		=	20	;
    TE_MONTANTDEV	=	21	;
    TE_MONTANT		=	22	;
    TE_COTATION		=	23	;
    TE_DATECOMPTABLE=	24	;
    TE_SOLDEDEV		=	25	;
    TE_DATEVALEUR	=	26	;
    TE_SOLDEDEVVALEUR=	27	;
    TE_CODEBANQUE	=	28	;
    TE_CODEGUICHET	=	29	;
    TE_NUMCOMPTE	=	30	;
    TE_CLERIB	=	31	;
    TE_IBAN	=	32	;
   {	TE_JOURNAL		  = 0 ;
    TE_NUMEROPIECE 	  = 1 ;
    TE_NUMLIGNE 	  = 2 ;
    TE_EXERCICE 	  = 3 ;
    TE_NUMECHE 		  = 4 ;
    TE_GENERAL 		  = 5 ;
    TE_CODECIB 		  = 6 ;
    TE_CODEFLUX 	  = 7 ;
    TE_REFINTERNE 	  = 8 ;
    TE_SOCIETE 		  = 9 ;
    TE_QUALIFORIGINE  = 10 ;
    TE_LIBELLE 		  = 11 ;
    TE_USERCREATION   = 12 ;
    TE_USERMODIF 	  = 13 ;
    TE_USERVALID 	  = 14 ;
    TE_USERCOMPTABLE  = 15 ;
    TE_DATECREATION   = 16 ;
    TE_DATEMODIF 	  = 17 ;
    TE_DATEVALID 	  = 18 ;
    TE_DATECOMPTABLE  = 19 ;
    TE_DATERAPPRO 	  = 20 ;
    TE_DEVISE 		  = 21 ;
    TE_MONTANTDEV 	  = 22 ;
    TE_MONTANT 		  = 23 ;
    TE_COTATION 	  = 24 ;
    TE_DATEVALEUR 	  = 25 ;
	TE_SOLDEDEVVALEUR = 26 ;
    TE_SOLDEDEV 	  = 27 ;
    TE_CODEBANQUE 	  = 28 ;
    TE_CODEGUICHET 	  = 29 ;
    TE_NUMCOMPTE 	  = 30 ;
    TE_CLERIB 		  = 31 ;
    TE_IBAN			  = 32 ; }
   //--------------
   //BORNES
   //--------------
	MAXNBJAMBES 	= 5 ;
	VERSEMENT		= 'VERSEMENT';
    AGIOSINTERETS 	= 'AGIOSINTERETS';
    TOMBE			= 'REMBANTICIPE';
    
   //--------------
   //RECORD
   //--------------
    type
    	TREcritures= record
            TE_JOURNAL          : string ;
            TE_NUMEROPIECE      : integer ;
            TE_NUMLIGNE         : integer ;
            TE_EXERCICE         : string ;
            TE_NUMECHE 		: integer ;
            TE_GENERAL          : string ;
            TE_CONTREPARTIETR   : string ;
            TE_CODECIB 		  : string ;
            TE_CODEFLUX 	  : string ;
            TE_REFINTERNE 	  : string ;
            TE_SOCIETE 		  : string ;
            TE_QUALIFORIGINE  : string ;
            TE_NUMTRANSAC	  : string ;
            TE_LIBELLE 		  : string ;
            TE_USERCREATION   : string ;
            TE_USERMODIF 	  : string ;
            TE_USERVALID 	  : string ;
            TE_USERCOMPTABLE  : string ;
            TE_DATECREATION   : string ;
            TE_DATEMODIF 	  : string ;
            TE_DATEVALID 	  : string ;
            TE_DATECOMPTABLE  : string ;
            TE_DATERAPPRO 	  : string ;
            TE_DEVISE 		  : string ;
            TE_MONTANTDEV 	  : double ;
            TE_MONTANT 		  : double ;
            TE_COTATION 	  : double ;
            TE_DATEVALEUR 	  : string ;
            TE_SOLDEDEVVALEUR : double ;
            TE_SOLDEDEV 	  : double ;
            TE_CODEBANQUE 	  : string ;
			   TE_CODEGUICHET 	  : string ;
            TE_NUMCOMPTE 	  : string ;
            TE_CLERIB 		  : string ;
            TE_IBAN			  : string ;
            TE_NATURE         : string ;
    end ;
    Type
    	TRCONDITIONFINPLAC = record
            TCF_AGIOSPRECOMPTE	: string ;
            TCF_AGIOSDEDUIT		: string ;
            TCF_BASECALCUL		: integer ;
            TCF_TAUXFIXE		: string ;
            TCF_TAUXPRECOMPTE	: string ;
            TCF_VALTAUX			: double ;
            TCF_TAUXVAR			: string ;
            TCF_VALMAJORATION	: double ;
            TCF_VALMULTIPLE		: double ;
            TCF_NBJOURBANQUE	: double ;
            TCF_NBJOURENCAISS	: double ;
            TCF_NBJOURMINAGIOS	: double ;
    end ;



   implementation

end.
