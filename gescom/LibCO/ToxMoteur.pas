unit ToxMoteur;

interface

uses {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Hctrls, SysUtils, uTob, StdCtrls, Dialogs, FactComm,
     FactUtil, FactCpta, SaisUtil, HEnt1, FactNomen, ToxIntegre, ToxMessage,
{$IFNDEF FOS5}
     ToxCash2000,
{$ENDIF}
     UtoxClasses, ParamSoc, Windows;

    
/////////////////////////////////////////////////////////////////////////////////
// Variables Globales
/////////////////////////////////////////////////////////////////////////////////
var NbEnreg    : integer ;
    NbEnregTop : integer ;
    NbMessErr  : integer ;
    IndiceTob  : integer ;
    TobMessage : TOB     ;

////////////////////////////////////////////////////////////////////////////////
// Statuts de la TOX
////////////////////////////////////////////////////////////////////////////////
const   ToxRejeter     = -1 ;
        ToxIntegrer    =  0 ;
        ToxARetraiter  =  1 ;

        BadFormatEvents = -1 ;
        BadFormatCondit = -2 ;
        BadFormatTable  = -3 ;

//////////////////////////////////////////////////////////////////////////////////////////////
// Prototyting des fonctions génériques de la TOX
//
//             TraitementDeLaTox      : fonction générique de traitement d'uneTOX
//             DetopageDeLaTox        : détopage d'une TOX
//             ConsulteMessageDeLaTox : consultation des messages d'intégration dans un TMemo
//             ChargeMessageDeLaTox   : charge les messages d'intégration et d'erreur
//                                      dans une TOB hierarchisée des messages
/////////////////////////////////////////////////////////////////////////////////////////////////
function  TraitementDeLaTox       (var Tox : TOB; CpteRendu : Tmemo; AfficheMessage : boolean; AfficheDonnee : boolean ; AfficheErreur : boolean ; IntegrationRapide : boolean ; MajStock : boolean; sDevise : string; sBasculeEuro : boolean) : integer ;
function  DetopageDeLaTox         (var Tox : TOB; CpteRendu : Tmemo) : integer ;
procedure ConsulteMessageDeLaTox  (var Tox : TOB; CpteRendu : Tmemo; AffDonneeIntegree, AffDonneeRejetee : boolean ) ;
procedure ChargeMessageDeLaTox    (var Tox, TobMessAff : TOB; Integre, Rejet : boolean; Var Nbenr, NbErr : integer );
procedure AvantArchivageTox       (Tox : TOB; var ArchivageTox : boolean);

implementation


///////////////////////////////////////////////////////////////////////////////////////
// Topage/Détopage d'une unité transactionnelle TOX
//
// Rappel : 1 - Unité Transactionnelle = Plusieurs enregistrements TOX rattachés
//              Exemple : Entête PIECE + LIGNE
//          2 - Le topage se fait sur le champ supplémentaire TOX_RESULT :
//                        -> ' ' ou '' pas traité
//                        -> 'T' en cours de traitement
//                        -> 'R' rejeté
//                        -> 'I' intégré
//          3 - Le numéro d'erreur est stocké dans le champ supplémentaire TOX_ERREUR
//
// !!!!!!!!!!!!!!!!!!!!!! Attention : Procédure récursive !!!!!!!!!!!!!!!!!!!!!!
//
///////////////////////////////////////////////////////////////////////////////////////
procedure TopageUniteTransactionnelle (var ToxInfo : TOB; Top : string );
var ToxTableNext    : TOB	   ;
    ToxInfoNext     : TOB    ;
    NomTableNext    : string ;
    i, j            : integer;
begin

  if ToxInfo.FieldExists ('TOX_RESULT') then
  begin
    Inc ( NbEnregTop ) ;
    //
    // Topage de l'enregistrement Entête
    //
    ToxInfo.PutValue ('TOX_RESULT', Top);

    //
    // Topage des enregistrements rattachés = Tob filles
    //
    for i:=0 to ToxInfo.Detail.Count-1 do
    begin
      ToxTableNext:=ToxInfo.Detail[i];
      if (ToxTableNext.FieldExists('TOX_TABLE')) then
      begin
        NomTableNext := ToxTableNext.GetValue ('TOX_TABLE');
        for j:=0 to ToxTableNext.Detail.Count-1 do
	   	  begin
		      ToxInfoNext := ToxTableNext.Detail[j];
    	    TopageUniteTransactionnelle (ToxInfoNext, Top);
    	  end;
      end;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
// Topage d'un enregistrement TOX
///////////////////////////////////////////////////////////////////////////////////////
procedure TopageToxEnreg (var ToxInfo : TOB; Integre : boolean ; Erreur : integer) ;
begin

  ToxInfo.PutValue ('TOX_ERREUR', Erreur);

  if not Integre then
  begin
    if Erreur = IntegOK then ToxInfo.PutValue ('TOX_RESULT', 'T')
    else                     ToxInfo.PutValue ('TOX_RESULT', 'R');
  end
  else begin
    if Erreur = IntegOK then ToxInfo.PutValue ('TOX_RESULT', 'I')
    else                     ToxInfo.PutValue ('TOX_RESULT', 'R');
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
// Détopage de la TOX
///////////////////////////////////////////////////////////////////////////////////////
function DetopageDeLaTox (var Tox : TOB; CpteRendu : Tmemo) : integer ;
var ToxCondition : TOB 		;
		ToxTable0		 : TOB		;
    ToxInfo0		 : TOB		;
    i,j,k				 : integer;
begin

  result := ToxIntegrer ;
  //
	// TOX MERE CONTENANT LE NOM DE L'EVENEMENT
  //
  if (Tox.FieldExists('EVENEMENT')) then
  begin
    CpteRendu.lines.add('   Début du détopage le ' + DateToStr(Date) + ' à ' + TimeToStr(Time));
    NbEnregTop := 0 ;

    for i:=0 to Tox.Detail.Count-1 do
    begin
  	  //
      // TOX CONDITION CONTENANT LA CONDITION
      //
      ToxCondition:=Tox.Detail[i];

      if (ToxCondition.FieldExists('CONDITION')) then
      begin
        for j:=0 to ToxCondition.Detail.Count-1 do
        begin
          //
    	    // TOX TABLE Niveau 0
    	    //
          ToxTable0:=ToxCondition.Detail[j];

          if (ToxTable0.FieldExists('TOX_TABLE')) then
  		    begin
            for k:=0 to ToxTable0.Detail.Count-1 do
            begin
              ToxInfo0 := ToxTable0.Detail[k];
              TopageUniteTransactionnelle (ToxInfo0, 'T');
            end;
  		    end else
          begin
            CpteRendu.lines.add('   Le format de la TOX est incorrect : Nom de la table de niveau 0 n''est pas renseigné');
            result := ToxRejeter ;
            exit ;
          end;
        end;
      end else
      begin
        CpteRendu.lines.add('   Le format de la TOX est incorrect : Pas de nom de condition');
        result := ToxRejeter ;
        exit;
      end;
    end;

    CpteRendu.lines.add('     -> ' + IntToStr (NbEnregTop)  + ' enregistrements traités ');
    CpteRendu.lines.add('   Fin du détopage le ' + DateToStr(Date) + ' à ' + TimeToStr(Time));

  end else
  begin
    CpteRendu.lines.add('   Le format de la TOX est incorrect : Pas de nom d''évènement');
    result := ToxRejeter ;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
// Descriptif des informations intégrées
// Paramètres     1 - NomTable          Nom de la table en cours de traitement
//                2 - ToxInfo           Information en cours d'intégration = TOB
//                3 - MessErreur        Numéro d'erreur.
//                                      Erreur > 0      Erreur d'intégration
//                                      Erreur = 0      Enregistrement intégré
//                                      Erreur = -1     Message d'intégration
//                4 - CpteRendu         TMemo servant d'affichage écran
///////////////////////////////////////////////////////////////////////////////////////
procedure EditeInfoMessage (NomTable : string; ToxInfo : TOB; MessErreur : integer; CpteRendu : TMemo);
var MessageInfo : string;
begin
  MessageInfo := ToxEditeMessage (NomTable, ToxInfo, MessErreur);
  if (MessageInfo<>'') then
  begin
    if MessErreur = -1 then CpteRendu.lines.add('       ' + MessageInfo)
    else CpteRendu.lines.add('           -> ' + MessageInfo);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
// Fonction appelant les fonctions de contrôles ou d'intégration des éléments TOX
// Valeur retournée :	 	0 Pas d'erreur
//											1 Erreur lors des contrôles ou lors des MAJ
//
// Remarque : Seuls les enregistrements non traités ou en erreur sont pris en compte.
///////////////////////////////////////////////////////////////////////////////////////
function ControleInfo (NomTable : string; var ToxInfo : TOB; Integre : boolean ; CpteRendu : TMemo ; AfficheDonnee : boolean; AfficheErreur : Boolean ; IntegrationRapide : boolean ; MajStock : boolean ; sDevise : string; sBasculeEuro : boolean; TobDev : TOB; CodeSite : string) : integer ;
var EtatEnreg : string ;
    Erreur    : integer ;
begin
  Result := 0 ;

  EtatEnreg := Trim(ToxInfo.GetValue ('TOX_RESULT'));

  if (EtatEnreg = '') or (EtatEnreg = 'R') or (EtatEnreg = 'T') then
  begin
    if Integre = False then inc (NbEnreg);
    //
    // Contrôle et import de l'enregistrement courant
    //
    Erreur := ToxControleInfo (NomTable, ToxInfo, Integre, IntegrationRapide, MajStock, sDevise, sBasculeEuro, TobDev, CodeSite);
    //
    // Topage de l'enregistrement traité
    //
    TopageToxEnreg (ToxInfo, Integre, Erreur);
    //
    // Affichage de l'information intégrée
    //
    if (AfficheDonnee) and (Not Integre) then	EditeInfoMessage (Nomtable, ToxInfo, -1, CpteRendu);
    //
    // Affichage de l'erreur
    //
     if Erreur <> 0	then
    begin
      if (AfficheErreur) then EditeInfoMessage (Nomtable, ToxInfo, Erreur, CpteRendu);
      result := 1;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
// Fonction de balayage d'une TOX
//
//    -> Création s'il y a lieu des champs supplémentaires de topage des TOX
//		-> Balayage de la TOX et appel des fonctions de contrôles et d'intégration d'une unité fonctionnelle
//		-> Retourne le nombre total d'erreurs rencontrées pour l'unité transactionnelle
//
// !!!!!!!!!!!!!!!!!!!!!!!!! Attention : Fonction récursive !!!!!!!!!!!!!!!!!!!!!!!!!!
//
///////////////////////////////////////////////////////////////////////////////////////
function ControleTox (var NomTable : string; var ToxInfo : TOB; Integre : boolean ; CpteRendu : TMemo; AfficheDonnee : boolean ; AfficheErreur : boolean ; IntegrationRapide : boolean ; MajStock : boolean; sDevise : string; sBasculeEuro : boolean; TobDev : TOB; CodeSite : string) : integer ;
var ToxTableNext    : TOB	   ;
    ToxInfoNext     : TOB    ;
    NomTableNext    : string ;
    i, j            : integer;
    Erreur          : integer;
    ErreurMAJDoc    : integer;
begin
  ErreurMAJDoc := 0;
  //
  // Ajout des champs spécifiques "TOX_RESULT" et "TOX_ERREUR"
  //
  if not ToxInfo.FieldExists ('TOX_RESULT') then
  begin
       ToxInfo.AddChampSup ('TOX_RESULT', False);
       ToxInfo.AddChampSup ('TOX_ERREUR', False);
       ToxInfo.PutValue    ('TOX_RESULT', ' ');
       ToxInfo.putValue    ('TOX_ERREUR',  0 );
  end;
  //
  // Traitement de l'enregistrement entête
  //
  Erreur := ControleInfo (NomTable, ToxInfo, Integre, CpteRendu, AfficheDonnee, AfficheErreur, IntegrationRapide, MajStock, sDevise, sBasculeEuro, TobDev, CodeSite);
  //
  // Traitement des niveaux inférieurs
  //
  for i:=0 to ToxInfo.Detail.Count-1 do
  begin
    //
 	  // TOX TABLE SUIVANTE
    //
    ToxTableNext:=ToxInfo.Detail[i];

    if (ToxTableNext.FieldExists('TOX_TABLE')) then
    begin
      NomTableNext := ToxTableNext.GetValue ('TOX_TABLE');
      for j:=0 to ToxTableNext.Detail.Count-1 do
	   	begin
   	      ToxInfoNext := ToxTableNext.Detail[j];
    	  Erreur      := Erreur + ControleTox (NomTableNext, ToxInfoNext, Integre, CpteRendu, AfficheDonnee, AfficheErreur, IntegrationRapide, MajStock, sDevise, sBasculeEuro, TobDev, CodeSite);
    	end;
    end
    else begin
      if AfficheErreur then CpteRendu.lines.add('   Le format de la TOX est incorrect : Nom de la table n''est pas renseigné');
      Inc (Erreur);
      break;
    end;
  end;
  //
  // Intégration spécifique des pièces : Tout doit être chargé en TOB pour pouvoir intégrer une pièce
  //                                     avec MAJ des stocks .....
  if (NomTable = 'PIECE') AND Integre AND (Erreur = 0) then
  begin
    if AfficheDonnee then CpteRendu.lines.add('       -> Import du document dans la base');

    try
      BeginTrans;
      ErreurMAJDoc := ImportDocument(MajStock, sDevise, sBasculeEuro, TobDev) ;
      CommitTrans ;
    except
      RollBack;
    end;

    if (AfficheErreur) and (ErreurMAJDoc <> 0)	then
    begin
      EditeInfoMessage (Nomtable, ToxInfo, ErreurMAJDoc, CpteRendu);
      Inc (Erreur);
    end;
  end;

  result:=Erreur;
end;

 ///////////////////////////////////////////////////////////////////////////////////////
// Fonctions génériques de traitement d'une TOX
///////////////////////////////////////////////////////////////////////////////////////
function TraitementDeLaTox (var Tox : TOB; CpteRendu : Tmemo; AfficheMessage : boolean; AfficheDonnee : boolean ; AfficheErreur : boolean ; IntegrationRapide : boolean ;  MajStock : boolean; sDevise : string; sBasculeEuro : boolean) : integer ;
var ToxCondition : TOB 		;
	ToxTable0	 : TOB		;
    ToxInfo0	 : TOB		;
    TobDevise    : TOB    ;
    Q            : TQUERY ;
    NomTable0	 : string	;
    CodeSite     : string ;
    NbErreur	 : integer;
    NbErreurDB   : integer;
    NbErr        : integer;
    NbEnrTotal   : integer;
    NbErrTotal   : integer;
    NbErrDBTotal : integer;
    i,j,k				 : integer;
begin
  //
  // Chargement de la TOB des devises permettant de convertir les montants en devise
  //
  TobDevise := TOB.Create('LES DEVISES', Nil, -1) ;
  Q         := OpenSQL('SELECT * FROM DEVISE',True) ;
  TobDevise.LoadDetailDB('DEVISE','','',Q,False) ;
  Ferme(Q) ;

  result := ToxIntegrer ;
  //
	// TOX MERE CONTENANT LE NOM DE L'EVENEMENT
  //
  if (Tox.FieldExists('EVENEMENT')) then
  begin
    if Tox.FieldExists('SITE-EMETTEUR') then CodeSite := Tox.Getvalue('SITE-EMETTEUR')
    else CodeSite := '';

    NbEnrTotal   := 0;
    NbErrTotal   := 0;
    NbErrDBTotal := 0;

    if AfficheMessage then
    begin
      CpteRendu.lines.add('   Début de l''intégration le ' + DateToStr(Date) + ' à ' + TimeToStr(Time));
      if CodeSite <> '' then CpteRendu.lines.add('   Code du site émetteur : ' + CodeSite);
      CpteRendu.lines.add ('   Evènement '+Tox.GetValue ('EVENEMENT'));
    end;

    for i:=0 to Tox.Detail.Count-1 do
    begin
  	  //
      // TOX CONDITION CONTENANT LA CONDITION
      //
      ToxCondition:=Tox.Detail[i];
      if (ToxCondition.FieldExists('CONDITION')) then
      begin
        if AfficheMessage then
        begin
          CpteRendu.lines.add(' ');
          CpteRendu.lines.add('   Requête '+Toxcondition.GetValue ('CONDITION'));
        end;

        for j:=0 to ToxCondition.Detail.Count-1 do
        begin
          //
    	    // TOX TABLE Niveau 0
    	    //
          ToxTable0:=ToxCondition.Detail[j];

          if (ToxTable0.FieldExists('TOX_TABLE')) then
  		    begin
            NomTable0  := ToxTable0.GetValue ('TOX_TABLE');			// Nom de la table principale
						NbErreur   := 0; NbErreurDB := 0; NBEnreg := 0; NbEnregTop := 0 ;

            if AfficheDonnee then
            begin
              CpteRendu.lines.add('   '); CpteRendu.lines.add('   Table principale : ' + NomTable0);
            end;

            for k:=0 to ToxTable0.Detail.Count-1 do
  					begin
              //
              // Contrôle l'unité transactionnelle (integration = FALSE)
              //
              ToxInfo0 := ToxTable0.Detail[k];
              NbErr    := ControleTox (NomTable0, ToxInfo0, FALSE, CpteRendu, AfficheDonnee, AfficheErreur, IntegrationRapide, MajStock, sDevise, sBasculeEuro, TobDevise, CodeSite);
              //
              // Si pas d'erreur, intégration de l'unité transactionnelle (Intégration = TRUE)
              // Sinon, topage de l'unité transactionnelle à 'R'
              //
        	  	if NbErr = 0 then
              begin
                NbErreurDB := NbErreurDB + ControleTox (NomTable0, ToxInfo0, TRUE, Cpterendu, AfficheDonnee, AfficheErreur, IntegrationRapide, MajStock, sDevise, sBasculeEuro, TobDevise, CodeSite);
                // MODIF LM 25/02/03 LACOSTE - Gère le plantage lors de la MAJ DB
                if NbErreurDB <>0 then TopageUniteTransactionnelle (ToxInfo0, 'R');
              end else
              begin
                TopageUniteTransactionnelle (ToxInfo0, 'R');
                NbErreur := NbErreurDB + Nberr;
              end;
            end;
            NbEnrTotal   := NbEnrTotal   + NbEnreg;
            NbErrTotal   := NbErrTotal   + NbErreur;
            NbErrDBTotal := NbErrDBTotal + NbErreurDB;
  		    end else
          begin
            if AfficheMessage then CpteRendu.lines.add('   Le format de la TOX est incorrect : Nom de la table de niveau 0 n''est pas renseigné');
            result := ToxRejeter ;
            exit ;
          end;
        end;
      end else
      begin
        if AfficheMessage then CpteRendu.lines.add('   Le format de la TOX est incorrect : Pas de nom de condition');
        result := ToxRejeter ;
        exit;
      end;
    end;

    if AfficheMessage then
    begin
      CpteRendu.lines.add('   ');
      CpteRendu.lines.add('   Intégration terminée le ' + DateToStr(Date) + ' à ' + TimeToStr(Time));
      CpteRendu.lines.add('   -> ' + IntToStr (NbEnrTotal)  + ' enregistrements traités ');
      CpteRendu.lines.add('   -> ' + IntToStr (NbErrTotal) + ' erreur(s) lors des contrôles');
      if (NbErrDBTotal <> 0) then CpteRendu.lines.add('   Attention : ' + IntToStr(NbErrDBTotal) + ' erreur(s) lors des mises à jour de la base');
    end;
    //
    // S'il y a des erreurs, il faudra la retraiter
    //
    if (NbErrDBTotal <> 0) or (NbErrTotal <> 0) then result := ToxARetraiter;
  end else
  begin
    if AfficheMessage then CpteRendu.lines.add('   Le format de la TOX est incorrect : Pas de nom d''évènement');
    result := ToxRejeter ;
  end;

  // Libération de la TOB des devises
  if TobDevise <> nil then
  begin
    TobDevise.free ;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
// Fonction permettant d'afficher :
//
//          1 - Un message décrivant l'enregistrement à traiter
//          2 - Un message décrivant l'erreur d'intégration
//
// Paramètres     1 - NomTable                 Nom de la table en cours de traitement
//                2 - ToxInfo                  Information en cours d'intégration = TOB
//                3 - MessErreur               Numéro d'erreur.
//                4 - CpteRendu                TMemo servant d'affichage écran
//                5 - AffDonneeIntegree        Affichage des données intégrées
//                6 - AffDonneeRejetee         Affiche des données rejétées
//
// 3 possibilités d'affichage :
//
//          1 - Données intégrées sans erreur  : AffDonnneeIntegree = TRUE et AffDonneeRejetee = FALSE
//          2 - Données intégrées avec erreur  : AffDonnneeIntegree = FALSE et AffDonneeRejetee = TRUE
//          3 - Toutes les données             : AffDonnneeIntegree = TRUE et AffDonneeRejetee = TRUE
///////////////////////////////////////////////////////////////////////////////////////
procedure AfficheMessage (NomTable : string; var ToxInfo : TOB; CpteRendu : TMemo ; AffDonneeIntegree : boolean; AffDonneeRejetee : Boolean ) ;
var Erreur        : integer ;
    AfficheErreur : boolean;
begin
    //
    // Récupération de l'erreur
    //
    if (ToxInfo.FieldExists ('TOX_RESULT')) and (ToxInfo.FieldExists ('TOX_ERREUR')) then
    begin
         Erreur := ToxInfo.GetValue ('TOX_ERREUR');
         AfficheErreur := True ;
    end
    else begin
     AfficheErreur := False;
     Erreur := -100;
    end;

    Inc (NbEnreg);
    if (AfficheErreur) and (Erreur <> 0) then Inc (NbMessErr);

    if AffDonneeIntegree then EditeInfoMessage (Nomtable, ToxInfo, -1, CpteRendu);
    if ((AffDonneeRejetee) and (Erreur <> 0)) then
    begin
      if not AffDonneeIntegree then EditeInfoMessage (Nomtable, ToxInfo, -1, CpteRendu);

      if AfficheErreur then EditeInfoMessage (Nomtable, ToxInfo, Erreur, CpteRendu)
      else CpteRendu.lines.add('         -> Attention : ce fichier n''a pas encore été traité');
    end;
end;

///////////////////////////////////////////////////////////////////////////////////////
// Affichage des informations contenues dans la TOX
//
//		-> Balayage de l'unité transactionnelle et appel de la fonction AfficheMessage
//
// !!!!!!!!!!!!!!!!!!!!!!!!! Attention : Fonction récursive !!!!!!!!!!!!!!!!!!!!!!!!!!
//
///////////////////////////////////////////////////////////////////////////////////////
procedure ConsulteMessage (var NomTable : string; var ToxInfo : TOB; CpteRendu : TMemo; AffDonneeIntegree, AffDonneeRejetee : boolean) ;
var ToxTableNext    : TOB	   ;
    ToxInfoNext     : TOB    ;
    NomTableNext    : string ;
    i, j            : integer;
begin

  AfficheMessage (NomTable, ToxInfo, CpteRendu, AffDonneeIntegree, AffDonneeRejetee);
  //
  // Balayage des niveaux inférieurs
  //
  for i:=0 to ToxInfo.Detail.Count-1 do
  begin
    ToxTableNext:=ToxInfo.Detail[i];

    if (ToxTableNext.FieldExists('TOX_TABLE')) then
    begin
      NomTableNext := ToxTableNext.GetValue ('TOX_TABLE');
      for j:=0 to ToxTableNext.Detail.Count-1 do
  	  begin
        ToxInfoNext := ToxTableNext.Detail[j];
        AfficheMessage (NomTableNext, ToxInfoNext, CpteRendu, AffDonneeIntegree, AffDonneeRejetee);
    	end;
    end
    else begin
      CpteRendu.lines.add('   Le format de la TOX est incorrect : Nom de la table n''est pas renseigné');
      break;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
// Consultation des messages d'intégration de la TOX
///////////////////////////////////////////////////////////////////////////////////////
procedure ConsulteMessageDeLaTox (var Tox : TOB; CpteRendu : Tmemo; AffDonneeIntegree, AffDonneeRejetee : boolean ) ;
var ToxCondition : TOB 		;
		ToxTable0		 : TOB		;
    ToxInfo0		 : TOB		;
    NomTable0		 : string	;
    i,j,k				 : integer;
begin
  //
	// TOX MERE CONTENANT LE NOM DE L'EVENEMENT
  //
  if (Tox.FieldExists('EVENEMENT')) then
  begin
    for i:=0 to Tox.Detail.Count-1 do
    begin
      NbEnreg   := 0 ;
      NbMessErr := 0 ;
  	  //
      // TOX CONDITION CONTENANT LA CONDITION
      //
      ToxCondition:=Tox.Detail[i];

      if (ToxCondition.FieldExists('CONDITION')) then
      begin
        for j:=0 to ToxCondition.Detail.Count-1 do
        begin
          //
    	    // TOX TABLE Niveau 0
    	    //
          ToxTable0:=ToxCondition.Detail[j];
          if (ToxTable0.FieldExists('TOX_TABLE')) then
  		    begin

            NomTable0  := ToxTable0.GetValue ('TOX_TABLE');			// Nom de la table principale

            for k:=0 to ToxTable0.Detail.Count-1 do
            begin
              ToxInfo0 := ToxTable0.Detail[k];
              ConsulteMessage (NomTable0, ToxInfo0, CpteRendu, AffDonneeIntegree, AffDonneeRejetee);
            end;

            CpteRendu.lines.add('   ');
            if NbEnreg   > 1 then  CpteRendu.lines.add('   -> ' + IntToStr (NbEnreg)   + ' enregistrements traités ')
            else                 CpteRendu.lines.add('   -> ' + IntToStr (NbEnreg)     + ' enregistrement traité ') ;
            if NbMessErr > 1 then CpteRendu.lines.add('   -> ' + IntToStr (NbMessErr)  + ' enregistrements rejetés')
            else                 CpteRendu.lines.add('   -> ' + IntToStr (NbMessErr)   + ' enregistrement rejeté')
            
  		    end else
          begin
            CpteRendu.lines.add('   Le format de la TOX est incorrect : Nom de la table de niveau 0 n''est pas renseigné');
            exit ;
          end;
        end;
      end else
      begin
        CpteRendu.lines.add('   Le format de la TOX est incorrect : Pas de nom de condition');
        exit;
      end;
    end;
  end else
  begin
    CpteRendu.lines.add('   Le format de la TOX est incorrect : Pas de nom d''évènement');
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
// ChargeMessage
// Ajout des champs supplémentaires MESS_DONNEE et MESS_ERREUR et initialisation de ces
// deux messages
///////////////////////////////////////////////////////////////////////////////////////
procedure ChargeMessage (NomTable : string; var ToxInfo, Mess : TOB; Integre, Rejet : boolean ) ;
var Erreur : integer ;
begin
  //
  // Récupération du numéro d'erreur
  //
  if (ToxInfo.FieldExists ('TOX_RESULT')) and (ToxInfo.FieldExists ('TOX_ERREUR')) then Erreur := ToxInfo.GetValue ('TOX_ERREUR')
  else Erreur := -100;
  if (Erreur = 0) and (ToxInfo.FieldExists ('TOX_RESULT')) AND (ToxInfo.GetValue ('TOX_RESULT')='R') then
  begin
   Erreur := 1000;
  end;
  //
  // Indice l'enregistrement pour nous permettre de l'afficher + tard
  //
  if not ToxInfo.FieldExists ('TOX_INDICE') then ToxInfo.AddChampSup ('TOX_INDICE', False);
  Inc (IndiceTob);
  ToxInfo.PutValue ('TOX_INDICE', IndiceTob);
  //
  // Chargement des messages d'intégration et des messages d'erreur
  //
  if ( (Integre = True) and ((Erreur = 0) or (Erreur = -100))) or ( (Rejet = True) and (Erreur > 0) ) then
  begin
    Mess := Tob.Create ('', TobMessage, -1);
    Mess.AddChampSup ('TABLE'      , False); Mess.PutValue ('TABLE'      , NomTable);
    Mess.AddChampSup ('MESS_DONNEE', False); Mess.PutValue ('MESS_DONNEE', ToxEditeMessage (NomTable, ToxInfo, -1));
    Mess.AddChampSup ('MESS_ERREUR', False); Mess.PutValue ('MESS_ERREUR', ToxEditeMessage (NomTable, ToxInfo, Erreur));
    Mess.AddChampSup ('ERREUR'     , False); Mess.PutValue ('ERREUR'     , Erreur);
    Mess.AddChampSup ('INDICE'     , False); Mess.PutValue ('INDICE'     , IndiceTob);
  end;
  //
  // Maj des compteurs d'enregistrements et d'erreurs
  //
  Inc (NbEnreg);
  if  Erreur > 0  then Inc (NbMessErr);
end;

///////////////////////////////////////////////////////////////////////////////////////
// Chargement dans la TOX des messages d'intégration et des messages d'erreur
//
//		-> Balayage de l'unité transactionnelle
//
// !!!!!!!!!!!!!!!!!!!!!!!!! Attention : Fonction récursive !!!!!!!!!!!!!!!!!!!!!!!!!!
//
///////////////////////////////////////////////////////////////////////////////////////
procedure ChargeMessageTox (NomTable : string; var ToxInfo : TOB; Integre, Rejet : Boolean) ;
var ToxTableNext    : TOB	   ;
    ToxInfoNext     : TOB    ;
    TobMess         : TOB    ;
    NomTableNext    : string ;
    i, j            : integer;
begin
  //
  // Initialisation de la TOB fille des messages
  //
  TobMess := nil ;
  ChargeMessage (NomTable, ToxInfo, TobMess, Integre, Rejet);

  if TobMess <> nil then
  begin
    //
    // Balayage des niveaux inférieurs
    //
    for i:=0 to ToxInfo.Detail.Count-1 do
    begin
      TobMessage := TobMess;

      ToxTableNext := ToxInfo.Detail[i];

      if (ToxTableNext.FieldExists('TOX_TABLE')) then
      begin
        NomTableNext := ToxTableNext.GetValue ('TOX_TABLE');
        for j:=0 to ToxTableNext.Detail.Count-1 do
  	    begin
          ToxInfoNext := ToxTableNext.Detail[j];
          ChargeMessageTox (NomTableNext, ToxInfoNext, Integre, Rejet);
    	  end;
      end
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
// TOX ayant un mauvais format
// On ne peut pas charger les messages associés !!!
///////////////////////////////////////////////////////////////////////////////////////
procedure ChargeMessageBadTox (ErreurFormat : integer) ;
var TobMess         : TOB    ;
begin
  TobMess := Tob.Create ('', TobMessage, -1);
  TobMess.AddChampSup ('TABLE'      , False);
  TobMess.PutValue    ('TABLE'      , 'Non Définie');
  TobMess.AddChampSup ('MESS_DONNEE', False);
  TobMess.PutValue    ('MESS_DONNEE', 'Non Définie');
  TobMess.AddChampSup ('MESS_ERREUR', False);

  if ErreurFormat = BadFormatEvents then
  begin
    TobMess.PutValue ('MESS_ERREUR', 'Le Format du fichier est incorrect ! Le nom de l''évènement n''est pas défini');
  end else
  if ErreurFormat = BadFormatCondit then
  begin
    TobMess.PutValue ('MESS_ERREUR', 'Le Format du fichier est incorrect ! Le nom de la condition n''est pas défini');
  end else
  if ErreurFormat = BadFormatTable then
  begin
    TobMess.PutValue ('MESS_ERREUR', 'Le Format du fichier est incorrect ! Le nom de la table n''est pas défini');
  end;
end;

//////////////////////////////////////////////////////////////////////////////////////
// Consultation des messages d'intégration de la TOX
///////////////////////////////////////////////////////////////////////////////////////
procedure ChargeMessageDeLaTox (var Tox, TobMessAff : TOB; Integre, Rejet : boolean; Var Nbenr, NbErr : integer );
var ToxCondition : TOB 		;
    ToxTable0	 : TOB		;
    ToxInfo0	 : TOB		;
    NomTable0	 : string	;
    i,j,k	 : integer;
begin
  Nbenr := 0 ; NbErr := 0;
  //
	// TOX MERE CONTENANT LE NOM DE L'EVENEMENT
  //
  if (Tox.FieldExists('EVENEMENT')) then
  begin
    NbEnreg   := 0 ;               // Initialise le nombre d'enregistrements
    NbMessErr := 0 ;               // Initialise le nombre d'erreurs
    IndiceTob := 0 ;               // Initialise l'indice de TOB

    for i:=0 to Tox.Detail.Count-1 do
    begin
  	  //
      // TOX CONDITION CONTENANT LA CONDITION
      //
      ToxCondition:=Tox.Detail[i];

      if (ToxCondition.FieldExists('CONDITION')) then
      begin
        for j:=0 to ToxCondition.Detail.Count-1 do
        begin
          //
    	  // TOX TABLE Niveau 0
    	  //
          ToxTable0:=ToxCondition.Detail[j];
          if (ToxTable0.FieldExists('TOX_TABLE')) then
  	    begin
            NomTable0  := ToxTable0.GetValue ('TOX_TABLE');			// Nom de la table principale
            for k:=0 to ToxTable0.Detail.Count-1 do
            begin
              ToxInfo0   := ToxTable0.Detail[k];
              TobMessage := TobMessAff ;
              ChargeMessageTox (NomTable0, ToxInfo0, Integre, Rejet);
            end;
  	    end else
            begin
            TobMessage := TobMessAff ;
            ChargeMessageBadTox (BadFormatTable);
            exit ;
            end;
        end;
      end else
      begin
        TobMessage := TobMessAff ;
        ChargeMessageBadTox (BadFormatCondit);
        exit;
      end;
    end;
    NbEnr := NbEnreg ;
    NbErr := NbMessErr ;
  end else
  begin
    TobMessage := TobMessAff ;
    ChargeMessageBadTox (BadFormatEvents);
  end;
end;

//////////////////////////////////////////////////////////////////////////////////////
// Traitement de la TOX avant l'archivage
// Utilité :
//
//         1 - Si site en CASH 2000, les informations seront transformées et un fichier ASCII sera généré
//             Le code retour est False (on ne crée pas de fichier TOX
//         2 - On peut toper les pièces extraites pour interdire les modifications dès l'envoi TOX
//
///////////////////////////////////////////////////////////////////////////////////////
procedure AvantArchivageTox (Tox : TOB; var ArchivageTox : boolean);
var LeSite                 : TCollectionSite  ;
    CodeSite               : string           ;
    Cash2000               : boolean          ;
    NumerosCaisse          : string           ;
    Variable               : TCollectionVariable ;
begin

  ArchivageTox := True ;
  Cash2000     := False ;
  //
  // Le site destinataire est-il un site CASH 2000 ?
  //
  if Tox.FieldExists('SITE') then
  begin
    CodeSite := Tox.Getvalue('SITE');
    //
    // Chargement des sites
    //
    LeSite:=TX_LesSites.Find(CodeSite) ;
    if LeSite<>Nil then
    begin
      Variable:=LeSite.LesVariables.Find('$CASH2000') ;
      if Variable<>Nil then Cash2000:=(Variable.SVA_VALUE = 'X') ;
      if Cash2000 then
      begin
        Variable:=LeSite.LesVariables.Find('$CAISSE') ;
        if Variable<>Nil then NumerosCaisse:=Variable.SVA_VALUE  ;
      end;
    end ;
  end;

  if Cash2000 then
  begin
{$IFNDEF FOS5}
    //
    // Archivage au format CASH 2000
    //
    AvantArchivageToxPourCASH (Tox, CodeSite, NumerosCaisse) ;
    ArchivageTox := False ;                                    // Pas d'archivage TOX : on s'en occupe
{$ENDIF}
  end else ArchivageTox := True ;

end;

end.
