unit PaieToxMoteur;

interface

uses DBTables,
     Hctrls,
     SysUtils,
     uTob,
     StdCtrls,
     Dialogs,
     PaieToxIntegre,
     PaieToxMessage,
     UtoxClasses,
     Windows;


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
function  PaieTraitementDeLaTox       (var Tox : TOB; CpteRendu : Tmemo; AfficheMessage : boolean; AfficheDonnee : boolean ; AfficheErreur : boolean ; IntegrationRapide : boolean ) : integer ;
function  PaieDetopageDeLaTox         (var Tox : TOB; CpteRendu : Tmemo) : integer ;
procedure PaieConsulteMessageDeLaTox  (var Tox : TOB; CpteRendu : Tmemo; AffDonneeIntegree, AffDonneeRejetee : boolean ) ;
procedure PaieChargeMessageDeLaTox    (var Tox, TobMessAff : TOB; Integre, Rejet : boolean; Var Nbenr, NbErr : integer );
procedure PaieAvantArchivageTox       (Tox : TOB; var ArchivageTox : boolean);

implementation


///////////////////////////////////////////////////////////////////////////////////////
// Topage/Détopage d'une unité transactionnelle TOX
//
// Rappel : 1 - Unité Transactionnelle = Plusieurs enregistrements TOX rattachés
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
procedure PaieTopageUniteTransactionnelle (var ToxInfo : TOB; Top : string );
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
    	    PaieTopageUniteTransactionnelle (ToxInfoNext, Top);
    	  end;
      end;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
// Topage d'un enregistrement TOX
///////////////////////////////////////////////////////////////////////////////////////
procedure PaieTopageToxEnreg (var ToxInfo : TOB; Integre : boolean ; Erreur : integer) ;
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
function PaieDetopageDeLaTox (var Tox : TOB; CpteRendu : Tmemo) : integer ;
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
              PaieTopageUniteTransactionnelle (ToxInfo0, 'T');
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
procedure PaieEditeInfoMessage (NomTable : string; ToxInfo : TOB; MessErreur : integer; CpteRendu : TMemo);
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
function PaieControleInfo (NomTable : string; var ToxInfo : TOB; Integre : boolean ; CpteRendu : TMemo ; AfficheDonnee : boolean; AfficheErreur : Boolean ; IntegrationRapide : boolean) : integer ;
var EtatEnreg : string ;
    Erreur    : integer ;
begin
  Result := 0 ;

  EtatEnreg := ToxInfo.GetValue ('TOX_RESULT');

  if (EtatEnreg = ' ') or (EtatEnreg = 'R') or (EtatEnreg = 'T') then
  begin
    if Integre = False then inc (NbEnreg);
    //
    // Contrôle et import de l'enregistrement courant
    //
    Erreur := PaieToxControleInfoPaie (NomTable, ToxInfo, Integre, IntegrationRapide);
    //
    // Topage de l'enregistrement traité
    //
    PaieTopageToxEnreg (ToxInfo, Integre, Erreur);
    //
    // Affichage de l'information intégrée
    //
    if (AfficheDonnee) and (Not Integre) then	PaieEditeInfoMessage (Nomtable, ToxInfo, -1, CpteRendu);
    //
    // Affichage de l'erreur
    //
     if Erreur <> 0	then
    begin
      if (AfficheErreur) then PaieEditeInfoMessage (Nomtable, ToxInfo, Erreur, CpteRendu);
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
function PaieControleTox (var NomTable : string; var ToxInfo : TOB; Integre : boolean ; CpteRendu : TMemo; AfficheDonnee : boolean ; AfficheErreur : boolean ; IntegrationRapide : boolean ) : integer ;
var ToxTableNext    : TOB	   ;
    ToxInfoNext     : TOB    ;
    NomTableNext    : string ;
    i, j            : integer;
    Erreur          : integer;
begin
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
  Erreur := PaieControleInfo (NomTable, ToxInfo, Integre, CpteRendu, AfficheDonnee, AfficheErreur, IntegrationRapide);
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
    	  Erreur      := Erreur + PaieControleTox (NomTableNext, ToxInfoNext, Integre, CpteRendu, AfficheDonnee, AfficheErreur, IntegrationRapide);
    	end;
    end
    else begin
      if AfficheErreur then CpteRendu.lines.add('   Le format de la TOX est incorrect : Nom de la table n''est pas renseigné');
      Inc (Erreur);
      break;
    end;
  end;
   result:=Erreur;
end;

///////////////////////////////////////////////////////////////////////////////////////
// Fonctions génériques de traitement d'une TOX
///////////////////////////////////////////////////////////////////////////////////////
function PaieTraitementDeLaTox (var Tox : TOB; CpteRendu : Tmemo; AfficheMessage : boolean; AfficheDonnee : boolean ; AfficheErreur : boolean ; IntegrationRapide : boolean ) : integer ;
var ToxCondition : TOB 		;
		ToxTable0		 : TOB		;
    ToxInfo0		 : TOB		;
    NomTable0		 : string	;
    NbErreur		 : integer;
    NbErreurDB   : integer;
    NbErr        : integer;
    NbEnrTotal   : integer;
    NbErrTotal   : integer;
    NbErrDBTotal : integer;
    i,j,k				 : integer;
begin
  result := ToxIntegrer ;
  //
	// TOX MERE CONTENANT LE NOM DE L'EVENEMENT
  //
  if (Tox.FieldExists('EVENEMENT')) then
  begin
    NbEnrTotal   := 0;
    NbErrTotal   := 0;
    NbErrDBTotal := 0;

    if AfficheMessage then
    begin
      CpteRendu.lines.add('   Début de l''intégration le ' + DateToStr(Date) + ' à ' + TimeToStr(Time));
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
              NbErr    := PaieControleTox (NomTable0, ToxInfo0, FALSE, CpteRendu, AfficheDonnee, AfficheErreur, IntegrationRapide);
              //
              // Si pas d'erreur, intégration de l'unité transactionnelle (Intégration = TRUE)
              // Sinon, topage de l'unité transactionnelle à 'R'
              //
        	  if NbErr = 0 then NbErreurDB := NbErreurDB + PaieControleTox (NomTable0, ToxInfo0, TRUE, Cpterendu, AfficheDonnee, AfficheErreur, IntegrationRapide)
              else begin
                PaieTopageUniteTransactionnelle (ToxInfo0, 'R');
                NbErreur := NbErreur + Nberr;
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
procedure PaieAfficheMessage (NomTable : string; var ToxInfo : TOB; CpteRendu : TMemo ; AffDonneeIntegree : boolean; AffDonneeRejetee : Boolean ) ;
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

    if AffDonneeIntegree then PaieEditeInfoMessage (Nomtable, ToxInfo, -1, CpteRendu);
    if ((AffDonneeRejetee) and (Erreur <> 0)) then
    begin
      if not AffDonneeIntegree then PaieEditeInfoMessage (Nomtable, ToxInfo, -1, CpteRendu);

      if AfficheErreur then PaieEditeInfoMessage (Nomtable, ToxInfo, Erreur, CpteRendu)
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
procedure PaieConsulteMessage (var NomTable : string; var ToxInfo : TOB; CpteRendu : TMemo; AffDonneeIntegree, AffDonneeRejetee : boolean) ;
var ToxTableNext    : TOB	   ;
    ToxInfoNext     : TOB    ;
    NomTableNext    : string ;
    i, j            : integer;
begin

  PaieAfficheMessage (NomTable, ToxInfo, CpteRendu, AffDonneeIntegree, AffDonneeRejetee);
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
        PaieAfficheMessage (NomTableNext, ToxInfoNext, CpteRendu, AffDonneeIntegree, AffDonneeRejetee);
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
procedure PaieConsulteMessageDeLaTox (var Tox : TOB; CpteRendu : Tmemo; AffDonneeIntegree, AffDonneeRejetee : boolean ) ;
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
              PaieConsulteMessage (NomTable0, ToxInfo0, CpteRendu, AffDonneeIntegree, AffDonneeRejetee);
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
procedure PaieChargeMessage (NomTable : string; var ToxInfo, Mess : TOB; Integre, Rejet : boolean ) ;
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
procedure PaieChargeMessageTox (NomTable : string; var ToxInfo : TOB; Integre, Rejet : Boolean) ;
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
  PaieChargeMessage (NomTable, ToxInfo, TobMess, Integre, Rejet);

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
          PaieChargeMessageTox (NomTableNext, ToxInfoNext, Integre, Rejet);
    	  end;
      end
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
// TOX ayant un mauvais format
// On ne peut pas charger les messages associés !!!
///////////////////////////////////////////////////////////////////////////////////////
procedure PaieChargeMessageBadTox (ErreurFormat : integer) ;
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
procedure PaieChargeMessageDeLaTox (var Tox, TobMessAff : TOB; Integre, Rejet : boolean; Var Nbenr, NbErr : integer );
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
              PaieChargeMessageTox (NomTable0, ToxInfo0, Integre, Rejet);
            end;
  	    end else
            begin
            TobMessage := TobMessAff ;
            PaieChargeMessageBadTox (BadFormatTable);
            exit ;
            end;
        end;
      end else
      begin
        TobMessage := TobMessAff ;
        PaieChargeMessageBadTox (BadFormatCondit);
        exit;
      end;
    end;
    NbEnr := NbEnreg ;
    NbErr := NbMessErr ;
  end else
  begin
    TobMessage := TobMessAff ;
    PaieChargeMessageBadTox (BadFormatEvents);
  end;
end;

//////////////////////////////////////////////////////////////////////////////////////
// Traitement de la TOX avant l'archivage
// Utilité : ????
//
///////////////////////////////////////////////////////////////////////////////////////
procedure PaieAvantArchivageTox (Tox : TOB; var ArchivageTox : boolean);
begin
  ArchivageTox := True ;
end;

end.
