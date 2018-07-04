unit Recordcomsis;

interface
const
  Formatcol =
  '08%-10.10s%-10.10s%-10.10s%1.1s%-25.25s';

  Formatpar3 =
  '03%-14.14s%-4.4s%20.20s%-2.2s%1.1s%10.10s%1.1s%1.1s%1.1s%1.1s%1.1s%-2.2s%-20.20s'+
  '%1.1s%1.1s%1.1s%-2.2s%-2.2s%-1.1s%1.1s';

  Formatmvtsis =
  'E%-2.2s%-10.10s%-20.20s%-10.10s%1.1s%1.1s%-13.13s%-13.13s%-4.4s%-6.6s%-2.2s%-8.8s%-6.6s%1.1s%-10.10s%-6.6s%1.1s%-3.3s%-13.13s%-6.6s%1.1s%1.1s%-13.13s%-8.8s%-3.3s';

  Formatpar4 = '04%-6.6s%-6.6s%-2.2s%-8.8s%1s%-15.15s%1s%1s%1s';

  Formatjrl = '05%-2.2s%-20.20s%-10.10s%1.1s%1.1s%1.1s%1.1s%1.1s%5.5s%5.5s%11.11s%1.1s%1.1s%1.1s';


type

  ar_dossier = record
	  code,		// code '00'
	  dossier, 	        // n° dossier
	  date_debut,        // date de début
	  date_fin,	        // date de fin
	  archivage,         // type d'archivage (E, J, B, R)
          lg_cpt,            // longueur n° compte
          pre_mtp,         	// précision des montants (2 si centimes)
          vers_file,         // version fichier : 05
	  vers_exe, 	        // version programme
          ecriture,          // taille du fichier
          format,		// format édition : P ou G
          origine,           //origine du fichier pl = P, pme = M, Edificas =E
          monnaie : string;	//monnaie de tenue FRF ou EURO
        end;
  ar_par1 = record
	   code,              // code '01'
	   nom,
	   activite : string;
        end;
  ar_par2 = record
	   code,		// code '02'
	   adr1,
	   adr2,
	   adr3 : string;
        end;
  ar_par3 = record
	   code,		// code '03'
	   siret,
	   ape,
	   tel,
	   std,		        // n° du plan comptable standard
	   cre_cpt,	        // Controle création nouveau compte (O/N)
	   cpt_resul,	        // n° compte de résultat
	   ana,		        // gestion analytique (O/N)
	   stat,		// gestion statistique (O/N)
	   fournisseur,	        // 1er caractère compte fournisseur
	   client,		// 1er caractère compte client
	   format,		// format édition : P ou G
	   lg_piece,	        // longueur n° piece
	   info,		// info. complémentaire fiche dossier
	   anouveau,            // indicateur d'anouveau
	   tva,
	   devise,              // GDEV gestion des devises
           v_numc,              // numero d'environnement pour fichier client associé
           v_numf,              // numero d'environnement pour fichier fournisseur associé
	   reper_ecr,
	   monnedit : string;	// édition mention monnaie
        end;
  ar_par4 = record
	   code,		// code '04'
	   date_encours,
	   date_stop,
	   dex_prev,	        // durée exercice précédent
	   password,	        // mot de passe
	   typ,		// type du dossier
	   res_prev,	        // résultat de l'exercice précédent
	   reopen,		// indicateur de réouverture
	   cloture,	        // cloture
	   typ_reouv : string;	// type de réouverture : 0 pas d'opération
					//		 1 en solde
					//		 2 en détail
        end;
  ar_jour = record
	   code,		// code = 05
	   journal,	        // code journal
	   lib,		// libelle
	   cpt_aut,	        // numéro de compte automatique
	   sens,		// sens deb ou cre
	   typ,		// type du journal
	   quan,		// traitement des quantites
	   echeance,	        //échéances
           typconf,
	   ecart,		// ecart toléré sur la TVA
           banque,            // Numéro de banque
           numcpte,           // Numéro de compte bancaire
           dateimpu,          // Date d'imputation: (O)pération ou (V)aleur
	                        //'P'= par pièce 'N'= non 'O'= au niveau du folio
	   ecart_euro,
	   monnaie :string;
        end;

  ar_compte = record
	   code,		// code = 'C'
	   compte,	        // n° du compte
	   lib,		// libellé du compte
	   lettre,	        // lettrage auto ou non
	   sold_pro,           // solde progressif ou non
	   typ,		// type du compte
	   mtpdeb,	        // mtp deb exercice précédent
	   mtpcre,	        // mtp cre
	   taux,
	   tva,
	   monnaie : string
        end;


  ar_cjour = record
	   code,		// code = 'J'
	   journal,	        // code journal
	   lib,		        // libelle du journal
	   cpt_aut,	        // numéro de compte automatique
	   sens,		// sens deb ou cre
	   typ,		        // type du journal
	   quan,		// traitement des quantites
	   echeance : string;	// échéances
        end;
  ar_colreg = record
	code,		        // code = 08
	cpte,		        // n° compte collectif ou regroupement
	inf,		        // borne inférieure
	sup,		        // borne supérieure
	typ,		        // type 'C', 'F', 'R'
	libelle : string	// libellé du compte
        end;

  ar_ecr = record
	   code,		// code = 'E'
	   jour,		// jour de l'écriture
	   cpte,		// n° de compte de l'écriture
	   lib,		        // libellé
	   piece,		// n° de piece
	   lettre,		// indicateur de lettrage
	   extra,		// ecr extra-comptable
	   debit,		// mtp debit
	   credit,	        // credit
	   stat,		// code statistique
	   echeance,	        // date d'échéance
	   mode,		// mode de règlement
	   quan,		// quantité
	   date_cre,	        // date de création de l'écriture
           topee,
	   Nocheque,
           date_rel,
           niv_rel,
           devise,
           mtdev,
           date_valeur,
	   suitelettre,       //2ème caractère de lettrage
	   m_euro,	        //identifiant euro F ou E
	   mteuro,		//contrevaleur euro
	   identecr,		//identifiant pointant sur ecriture
	   autre_monn : string		//indicateur autre monnaie
				//			EUR=EURO, FRF=Franc pour une précision
				//			de la nature debit, credit à blanc correspond
				//			à la monnaie tenue de l'exercice*/
        end;



implementation

end.
