{***********UNITE*************************************************
Auteur  ...... : PL
Créé le ...... : 01/03/2001
Modifié le ... :
Description .. : Charge un planning suivant les paramètres d'un modèle
Suite ........ : de planning (copie modifiée de l'HRPlanningConfig)
Mots clefs ... : PLANNING
*****************************************************************}

unit AFPlanningConfig;

interface
uses Classes, Controls, Windows,StdCtrls,forms,sysutils,ComCtrls,graphics,
HCtrls,HEnt1,HMsgBox,UTOM,UTob,Dialogs,M3fp,Hplanning,HPanel,UiUtil,HStatus,
{$IFDEF EAGLCLIENT}
		eFiche, eFichList,
{$ELSE}
		dbTables, db,   Fiche, FichList, FE_Main,
{$ENDIF}
AffaireUtil;

procedure ChargeParamPlanning(Planning:THPlanning;TobModelePlanning:TOB;indice:integer;Critere:string;DateEnCours:TDateTime;var TobEtats, TobACT, TOBRes, TobEvenements:TOB);

implementation

procedure ChargeParamPlanning(Planning:THPlanning;TobModelePlanning:TOB;indice:integer;Critere:string;DateEnCours:TDateTime;var TobEtats, TobACT, TOBRes, TobEvenements:TOB);
var MaValeur,ColField1,ColField2,ColField3,formegraphique,cadencement:string;
    TobDetEtat, TobDetACT:Tob;
    QRecupTypeChamp,Qressource,QformatDate,Q:TQuery;
    boucle,compteur,i:integer;
    NomDuChamp:array[1..3] of String;
    TypeDuChamp:array[1..3] of String;
    stSQL, stSQLArt, sqry, Champ:string;
    sNomTablette,sNomTable,sListeChamps,sChampCritere : String;
    DateDebut,DateFin:TDateTime;
begin
     // Barre de patiente
     InitMove(10,'Chargement du planning en cours');
     Planning.Activate :=False;
     Planning.SurBooking:=False;
     Planning.ActiveSaturday:=True;
     Planning.ActiveSunday:=True;

     // Cadencement
     cadencement := TobModelePlanning.detail[indice].Getvalue('HPP_CADENCEMENT');
     if cadencement = '003' then
     begin
          if Planning.interval<> piJour then Planning.interval:=piJour;
     end;
(*     if cadencement = '004' then
     begin
          Planning.interval:=piDemiJour;
          Planning.JourneeDebut :=StrToDateTime('00:00');
          Planning.JourneeFin := StrToDateTime('23:00');
     end;
     if cadencement = '005'then
     begin
          Planning.interval:=piHeure;
          Planning.JourneeDebut :=StrToDateTime('00:00');
          Planning.JourneeFin := StrToDateTime('23:00');
    end;     *)

     // Changement date d'un séjour sur planning possible O/N
     if TobModelePlanning.Detail[indice].getValue('HPP_MOVEHORIZONTAL')='-' then planning.MoveHorizontal:=False
     else Planning.MoveHorizontal:=True;

     // Chargement de la forme graphique
     formegraphique:= TobModelePlanning.detail[indice].Getvalue('HPP_FORMEGRAPHIQUE');
     if (formegraphique = 'PGF') then Planning.FormeGraphique :=  pgFleche
     else if (formegraphique = 'PGB') then Planning.FormeGraphique :=  pgBerceau
     else if (formegraphique = 'PGL') then Planning.FormeGraphique :=  pgLosange
     else if (formegraphique = 'PGE') then Planning.FormeGraphique :=  pgEtoile
// supprimé en 5.2.2 agl ....     else if (formegraphique = 'PGA') then Planning.FormeGraphique :=  pgRoundRect
     else if (formegraphique = 'PGI') then Planning.FormeGraphique :=  pgBisot
     else Planning.FormeGraphique :=  pgFleche;

     //Mise à jour du format de la date
     QFormatDate := nil;
     try
       QFormatDate:=OpenSQL('SELECT CC_ABREGE FROM CHOIXCOD WHERE CC_CODE="'+TobModelePlanning.detail[indice].Getvalue('HPP_FORMATDATECOL0')+'" AND CC_TYPE="HFD"',True);
       Planning.DateFormat := QFormatDate.Fields[0].AsString
     finally
        Ferme(QFormatDate);
     end;
     MoveCur(FALSE);

     // MAJ Taille des colonnes
     Planning.ColSizeData :=TobModelePlanning.detail[indice].Getvalue('HPP_TAILLECOLONNE');
     Planning.ColSizeEntete:=TobModelePlanning.detail[indice].Getvalue('HPP_TAILLECOLENTET');

     // MAJ Nombre de lignes & colonnes divers
     // GISE 2 lignes mises en commentaires plante avec AGL 5.2.2 d  
//     Planning.NombreColonneDivers := TobModelePlanning.detail[indice].Getvalue('HPP_NBCOLDIVERS')-1;
//     Planning.NombreLignesDivers  := TobModelePlanning.detail[indice].Getvalue('HPP_NBLIGDIVERS')-1;

     // MAJ couleur
     Planning.ColorSelection := StrToInt64(TobModelePlanning.detail[indice].Getvalue('HPP_COULSELECTION'));
     Planning.ColorBackground  := StrToInt64(TobModelePlanning.detail[indice].Getvalue('HPP_COULEURFOND'));
     Planning.ColorOfSaturday := StrToInt64(TobModelePlanning.detail[indice].Getvalue('HPP_COULEURSAMEDI'));
     Planning.ColorOfSunday := StrToInt64(TobModelePlanning.detail[indice].Getvalue('HPP_COULDIMANCHE'));
     // MAJ de la police des colonnes
     Planning.FontNameColRowFixed:=TobModelePlanning.detail[indice].Getvalue('HPP_FONTCOLONNE');

     // Encadrement
     Planning.FrameOn := True;

     // MAJ regroupement de date
     If TobModelePlanning.detail[indice].Getvalue('HPP_AFFDATEGROUP')='X' then Planning.ActiveLigneGroupeDate:=TRUE;
     If TobModelePlanning.detail[indice].Getvalue('HPP_AFFDATEGROUP')='-' then Planning.ActiveLigneGroupeDate:=FALSE;
     Planning.Align:=AlClient;

     // Champ des items
     Planning.ChampdateDebut := 'ATA_DATEDEBUT';
     Planning.ChampDateFin   := 'ATA_DATEFIN';
//     Planning.ChampLibelle   := 'HDR_NOMCLIENT1';
//     Planning.ChampLineID    := 'HDR_RESSOURCE';
     Planning.ChampEtat      := 'ATA_FAMILLETACHE' ;
//     Planning.ChampHint      := 'HINT';



// Ordre de tri des colonnes.
if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='GLA') or
    (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='AFF') then
    begin
    sNomTablette := 'AFLIBCOLPLANNINGRES';
    sNomTable := 'RESSOURCE';
    sListeChamps := 'ARS_RESSOURCE,';
    end
else
if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='GLR') or
    (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='RES') then
    begin
    sNomTablette := 'AFLIBCOLPLANNINGAFF';
    sNomTable := 'AFFAIRE';
    sListeChamps := 'AFF_AFFAIRE,';
    end;




     sChampCritere := '';
     if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='AFF') or
        (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='RES') then
        begin
        Planning.ChampLibelle:='ATA_LIBELLE';
        Planning.ChampHint :='ATA_LIBELLE';
        if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='AFF') then
            begin
            sChampCritere := 'ATA_AFFAIRE';
            Planning.ChampLineID:='ATR_RESSOURCE';
            end
        else
            begin
            sChampCritere := 'ATR_RESSOURCE';
            Planning.ChampLineID:='ATA_AFFAIRE';
            end;

        stSQL := 'SELECT * FROM tache LEFT OUTER JOIN tacheressource ON (ata_tiers=atr_tiers) AND (ata_numerotache=atr_numerotache)';
        end
    else
        begin
        if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='GLR') then
            begin
            Planning.ChampLibelle:='ARS_LIBELLE';
            Planning.ChampHint :='ARS_LIBELLE';
            Planning.ChampLineID:='ATA_AFFAIRE';
            stSQL := 'SELECT * FROM tache LEFT OUTER JOIN tacheressource ON (ata_tiers=atr_tiers) AND (ata_numerotache=atr_numerotache) AND LEFT OUTER JOIN ressource ON (atr_ressource=ars_ressource)';
            end
        else
        if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='GLA') then
            begin
            Planning.ChampLibelle:='LIBELLEAFFAIRE';
            Planning.ChampHint :='LIBELLEAFFAIRE';
            Planning.ChampLineID:='ATR_RESSOURCE';
            stSQL := 'SELECT * FROM tache LEFT OUTER JOIN tacheressource ON(ata_tiers=atr_tiers) AND (ata_numerotache=atr_numerotache)';
            end;
        end;


     //Lignes du planning
     Planning.RowFieldID := 'R_RESSOURCE';
     Planning.RowFieldReadOnly := 'R_RO';
     Planning.RowFieldColor := 'R_COLOR';
     MoveCur(FALSE);

     // Champ des états
     Planning.EtatChampCode  := 'CC_CODE';
     Planning.EtatChampLibelle := 'CC_LIBELLE';
     Planning.EtatChampBackGroundColor := 'BGC';
     Planning.EtatChampFontColor := 'COLOR';
     Planning.EtatChampFontName := 'FONTE';
     MoveCur(FALSE);

     // Gestion des évenements
     Planning.ChampCodeEvenement:='HEV_EVENEMENT';
     Planning.ChampLibelleEvenement:='HEV_LIBELLE';
     Planning.ChampDateDebutEvenement:='HEV_DATEDEBUT';
     Planning.ChampDateDeFinEvenement:='HEV_DATEFIN';
     Planning.ChampCouleurEvenement:= 'HEV_COULEUR' ;
     Planning.ChampStyleEvenement:='HEV_STYLE';

     // Tobs
     {TobRows := TOB.Create ('les lignes', Nil, -1) ;
     TobRows.AddChampSup ('R_RESSOURCE', False) ;
     TobRows.AddChampSup ('R_COLOR', False) ;
     TobRows.AddChampSup ('R_RO', False) ;}

     MoveCur(FALSE);
     // Activation bulles d'aide
     Planning.ShowHint:=True;
     Planning.MouseAlready:=True;
     // Ordre de tri des colonnes.
    Champ := RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL3'), true);
    if (Champ<>'') and (Champ<>'Error') then
        begin
        sListeChamps := sListeChamps + RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL1'), true)+','
                                + RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL2'), true)+','
                                + Champ;
        sqry := 'SELECT '+ sListeChamps +' FROM '+ sNomTable +' ORDER BY '
            + RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL1'), true)+','
            + RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL2'), true)+','
            + Champ;
        end
    else
        begin
        Champ := RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL2'), true);
        if (Champ<>'') and (Champ<>'Error') then
            begin
            sListeChamps := sListeChamps + RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL1'), true)+','
                                    + Champ;
            sqry := 'SELECT '+ sListeChamps +' FROM '+ sNomTable +' ORDER BY '
                + RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL1'), true)+','
                + Champ;
            end
        else
            begin
            Champ := RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL1'), true);
            if (Champ<>'') and (Champ<>'Error') then
                begin
                sListeChamps := sListeChamps + Champ;
                sqry := 'SELECT '+ sListeChamps +' FROM '+ sNomTable +' ORDER BY ' + Champ;
                end;
            end;
        end;

    if (sqry<>'') then
    begin
      QRessource := nil;
      try
        QRessource:=OpenSQL(sqry, True);
        TOBRes:=Tob.Create('les ressources', Nil, -1);
        TOBRes.LoadDetailDB(sNomTable,'','',Qressource, false);  // HRTYPRES
      finally
        ferme(Qressource);
      end;

     // Récupératon des libéllés des champs pour les colonnes d'entete
    for boucle:=1 to 3 do
        begin
        NomDuChamp[boucle]:=RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL'+IntToStr(boucle)), true);
        QRecupTypeChamp := nil;
        try
          QRecupTypeChamp:=OpenSQL('SELECT DH_TYPECHAMP FROM DECHAMPS WHERE DH_NOMCHAMP="'+NomduChamp[boucle]+'"',True);
          TypeduChamp[boucle]:=QRecupTypeChamp.Fields[0].AsString
        finally
          ferme(QRecupTypeChamp);
        end;
        end;

    for compteur:=0 to TOBRes.Detail.Count-1 do
     begin
     for boucle:=1 to 3 do
          begin
          if (TypeDuchamp[boucle]='COMBO')then
               begin
               MaValeur:=RechDom(ChampToTT(NomDuChamp[Boucle]), TOBRes.detail[compteur].Getvalue(NomDuChamp[Boucle]),True) ;
               TOBRes.detail[compteur].PutValue(NomDuChamp[boucle], MaValeur);
               end;
          end;
     end;
     MoveCur(FALSE);
  
     // Ressources non attribuées
{     QNbTypres:=OpenSQL('SELECT COUNT(*) as NOMBRE from HRTYPRES where HTR_NATURERES="TRE" and HTR_FAMRES="'+TobModelePlanning.detail[indice].getvalue('HPP_FAMRES')+'"',True);
     Nbtypres:=Qnbtypres.FindField('NOMBRE').AsInteger;
     Ferme(QNbTypres);
     SetLength(TobressourceFille,Nbtypres);
     QNbTypres:=OpenSQL('Select * from HRTYPRES where HTR_NATURERES="TRE" and HTR_FAMRES="'+TobModelePlanning.detail[indice].getvalue('HPP_FAMRES')+'"',True);
     QNbTypres.FindFirst;
     try
     for i:=0 to Nbtypres-1 do
     begin
          TobRessourceFille[i]:=Tob.Create('HRTYPRES',TobRessource,-1);
          TobRessourceFille[i].putvalue('HTR_RESSOURCE','');
          TobRessourceFille[i].putvalue('HTR_TYPRES',QNbTypres.findfield('HTR_RESSOURCE').AsString);
          QNbTypres.findNext;
     end;
     finally ferme(QNbTypres)end;
     if TobRessource.Detail.Count=0 then
     begin
          ShowMessage('Aucun type de ressources n''est défini');
          Exit;
     end;
     MoveCur(FALSE);
 }

//Planning.TobRes := TobRes;

if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='GLA') or
    (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='AFF') then
    begin
    Planning.ResChampID := 'ARS_RESSOURCE';
    end
else
if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='GLR') or
    (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='RES') then
    begin
    Planning.ResChampID := 'AFF_AFFAIRE';
    end;


     // Initalisation des colonnes d'entete
    ColField1:=RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL1'), true);
    ColField2:=RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL2'), true);
    ColField3:=RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL3'), true);
    Planning.TokenFieldColFixed:=ColField1+';'+ColField2+';'+ColField3;
    Planning.TokenSizeColFixed:=TobModelePlanning.detail[indice].Getvalue('HPP_TAILLECOLENTET');
    Planning.TokenAlignColFixed:='C;C;C';
     Planning.TobRes:=TobRes;
     MoveCur(FALSE);

     // Icones
(*     if TobModelePlanning.detail[indice].Getvalue('HPP_ICONE')='X' then
     begin
          Planning.ChampIcone:='ICONETYPDOS';
        // Planning.EtatChampIcone :='HES_ICONE';
     end;*)
// ****************** Chargement des Etats **************************
TobEtats := Tob.Create ('liste des etats', Nil,-1);
stSQLArt := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="AF1"';
Q := OPENSQL(stSQLArt,True);
TOBEtats.LoadDetailDB('CHOIXCOD','','',Q,False);
ferme(Q);
if TOBEtats.Detail.Count > 0 then
    BEGIN
    TobDetEtat :=TOBEtats.Detail[0];
    TobDetEtat.AddchampSup ('BGC',True);
    TobDetEtat.AddchampSup ('COLOR',True);
    TobDetEtat.AddchampSup ('FONTE',True);
    for i := 0 to  TOBEtats.Detail.Count-1 do
        BEGIN
        TobDetEtat :=TOBEtats.Detail[i];
        TobDetEtat.putvalue ('BGC','clBlue');
        TobDetEtat.putvalue ('COLOR','$0080FFFF');
        TobDetEtat.putvalue ('FONTE','MS Sans Serif');
        END;
    END;
Planning.TobEtats := TobEtats ;

     Planning.IntervalDebut:=DateEnCours-(TobModelePlanning.Detail[indice].GetValue('HPP_INTERVALLEDEB'));
     Planning.IntervalFin:=DateEnCours+(TobModelePlanning.Detail[indice].GetValue('HPP_INTERVALLEFIN'));
     Planning.DateOfStart := DateEnCours;

     //Chargement des données (items) comprises dans l'intervalle choisi
     // Calcul des la date de debut et de fin
     DateDebut:=(DateEncours-TobModelePlanning.Detail[indice].getvalue('HPP_INTERVALLEDEB'));
     DateFin:=(TobModelePlanning.Detail[indice].getvalue('HPP_INTERVALLEFIN')+DateEncours);
stSQL := stSQL + ' WHERE (ATA_DATEDEBUT>="'+UsDateTime(Datedebut)+'" AND ATA_DATEFIN<="'+USDateTime(DateFin)+'" )';
// Les items
Q := OPENSQL(stSQL,True);
TOBACT := TOB.Create('liste des items', nil, -1);
TOBACT.LoadDetailDB('la tob items', '', '', Q, False);
if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='GLA') then
    // Gestion du code affaire à afficher
    begin
    if TOBACT.Detail.Count > 0 then
        BEGIN
        TobDetACT :=TOBACT.Detail[0];
        TobDetACT.AddchampSup ('LIBELLEAFFAIRE',True);
        for i := 0 to  TOBACT.Detail.Count-1 do
            BEGIN
            TobDetACT :=TOBACT.Detail[i];
            TobDetACT.putvalue('LIBELLEAFFAIRE',CodeAffaireAffiche(TobDetACT.GetValue('ATA_AFFAIRE')));
            END;
        END;
    end;
ferme(Q);

Planning.TOBItems := TOBACT;

     // Gestion evenements
     TobEvenements := TOB.Create ( '', Nil, -1 ) ;
     TobEvenements.LoadDetailDB ('HREVENEMENT', '', '', Nil, False, True ) ;
     Planning.GestionEvenements:=True;
     Planning.TobEvenements:=TobEvenements;
     // Activation Planning
//     Planning.GestionNuitee :=true;
//     Planning.DisplayRatio:=25;
//     Planning.Activate :=True;

end;

(*
begin
     // Barre de patience
     InitMove(10,'Chargement du planning en cours');

     Planning.Activate :=False;
     Planning.SurBooking:=True;
     Planning.ActiveSaturday:=True;
     Planning.ActiveSunday:=True;
     Planning.ShowHint:=True;


     // Cadencement
     cadencement := TobModelePlanning.detail[indice].Getvalue('HPP_CADENCEMENT');
     if cadencement = '003' then
     begin
          Planning.interval:=piJour;
     end;
     if cadencement = '004' then
     begin
          Planning.interval:=piDemiJour;
          Planning.JourneeDebut :=StrToDateTime('08:00');
          Planning.JourneeFin := StrToDateTime('12:00');
     end;
     if cadencement = '005'then
     begin
          Planning.interval:=piHeure;
          Planning.JourneeDebut :=StrToDateTime('08:00');
          Planning.JourneeFin := StrToDateTime('18:00');
     end;

     // Changement date d'un séjour sur planning possible O/N
     if TobModelePlanning.Detail[indice].getValue('HPP_MOVEHORIZONTAL')='-' then planning.MoveHorizontal:=False
     else Planning.MoveHorizontal:=True;

     // Chargement de la forme graphique
     formegraphique:= TobModelePlanning.detail[indice].Getvalue('HPP_FORMEGRAPHIQUE');
     if (formegraphique = 'PGF') then Planning.FormeGraphique :=  pgFleche
     else if (formegraphique = 'PGB') then Planning.FormeGraphique :=  pgBerceau
     else if (formegraphique = 'PGL') then Planning.FormeGraphique :=  pgLosange
     else if (formegraphique = 'PGE') then Planning.FormeGraphique :=  pgEtoile
     else if (formegraphique = 'PGA') then Planning.FormeGraphique :=  pgRoundRect
     else if (formegraphique = 'PGI') then Planning.FormeGraphique :=  pgBisot;

     //Mise à jour du format de la date
     QFormatDate:=OpenSQL('SELECT CC_ABREGE from CHOIXCOD where CC_CODE="'+TobModelePlanning.detail[indice].Getvalue('HPP_FORMATDATECOL0')+'" and CC_TYPE="HFD"',True);
     try Planning.DateFormat := QFormatDate.Fields[0].AsString finally Ferme(QFormatDate); end;
     MoveCur(FALSE);

     // MAJ Taille des colonnes
     Planning.ColSizeData :=TobModelePlanning.detail[indice].Getvalue('HPP_TAILLECOLONNE');
     Planning.ColSizeEntete:=TobModelePlanning.detail[indice].Getvalue('HPP_TAILLECOLENTET');

     // MAJ Nombre de lignes & colonnes divers
     Planning.NombreColonneDivers := TobModelePlanning.detail[indice].Getvalue('HPP_NBCOLDIVERS')-1;
     Planning.NombreLignesDivers  := TobModelePlanning.detail[indice].Getvalue('HPP_NBLIGDIVERS')-1;

     // MAJ couleur
     Planning.ColorSelection := StrToInt64(TobModelePlanning.detail[indice].Getvalue('HPP_COULSELECTION'));
     Planning.ColorBackground  := StrToInt64(TobModelePlanning.detail[indice].Getvalue('HPP_COULEURFOND'));
     Planning.ColorOfSunday := StrToInt64(TobModelePlanning.detail[indice].Getvalue('HPP_COULDIMANCHE'));
     Planning.ColorOfSaturday := StrToInt64(TobModelePlanning.detail[indice].Getvalue('HPP_COULEURSAMEDI'));

     // MAJ de la police des colonnes
     Planning.FontNameColRowFixed:=TobModelePlanning.detail[indice].Getvalue('HPP_FONTCOLONNE');

     // Encadrement
     Planning.FrameOn := True;

     // MAJ regroupement de date
     If TobModelePlanning.detail[indice].Getvalue('HPP_AFFDATEGROUP')='X' then Planning.ActiveLigneGroupeDate:=TRUE;
     If TobModelePlanning.detail[indice].Getvalue('HPP_AFFDATEGROUP')='-' then Planning.ActiveLigneGroupeDate:=FALSE;
     Planning.Align:=AlClient;


// Recup param fiche AGL
//CodeRes := GetControlText('ACT_RESSOURCE');
//DateDeb := StrToDate(GetControlText('ACT_DATEACTIVITE'));

// Paramètres généraux
//Pl.DateOfStart := DateDeb;

(*if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='AFF') or
    (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='RES') then
    begin
    stSQL := 'SELECT ACT_DATEACTIVITE,ACT_RESSOURCE,ACT_AFFAIRE,ACT_TYPEACTIVITE,ACT_LIBELLE,ACT_QTE,ACT_CODEARTICLE' +
          ' FROM ACTIVITE' ;
    end
else
    begin
    if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='GLA' then
        stSQL := 'SELECT * FROM RESSOURCE'
    else
    if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='GLR' then
        stSQL := 'SELECT * FROM AFFAIRE' ;
    end;
//////////*******

// **************** Chargement TOB Ressource  **********************

// Ordre de tri des colonnes.
if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='GLA') or
    (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='AFF') then
    begin
    sNomTablette := 'AFLIBCOLPLANNINGRES';
    sNomTable := 'RESSOURCE';
    sListeChamps := 'ARS_RESSOURCE,';
    end
else
if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='GLR') or
    (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='RES') then
    begin
    sNomTablette := 'AFLIBCOLPLANNINGAFF';
    sNomTable := 'AFFAIRE';
    sListeChamps := 'AFF_AFFAIRE,';
    end;

Champ := RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL3'), true);
if (Champ<>'') and (Champ<>'Error') then
    begin
    sListeChamps := sListeChamps + RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL1'), true)+','
                                + RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL2'), true)+','
                                + Champ;
    sqry := 'SELECT '+ sListeChamps +' from '+ sNomTable +' ORDER BY '
        + RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL1'), true)+','
        + RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL2'), true)+','
        + Champ;
    end
else
    begin
    Champ := RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL2'), true);
    if (Champ<>'') and (Champ<>'Error') then
        begin
        sListeChamps := sListeChamps + RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL1'), true)+','
                                    + Champ;
        sqry := 'SELECT '+ sListeChamps +' from '+ sNomTable +' ORDER BY '
            + RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL1'), true)+','
            + Champ;
        end
    else
        begin
        Champ := RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL1'), true);
        if (Champ<>'') and (Champ<>'Error') then
            begin
            sListeChamps := sListeChamps + Champ;
            sqry := 'SELECT '+ sListeChamps +' from '+ sNomTable +' ORDER BY ' + Champ;
            end;
        end;
    end;

if (sqry<>'') then
try
 begin
      QRessource:=OpenSQL(sqry, True);
      TOBRes:=Tob.Create('les ressources', Nil, -1);
      TOBRes.LoadDetailDB(sNomTable,'','',Qressource, false);  // HRTYPRES
 end
finally ferme(Qressource); end;


     // Récupératon des libéllés des champs pour les colonnes d'entete
for boucle:=1 to 3 do
    begin
    NomDuChamp[boucle]:=RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL'+IntToStr(boucle)), true);
    QRecupTypeChamp:=OpenSQL('SELECT DH_TYPECHAMP from DECHAMPS where DH_NOMCHAMP="'+NomduChamp[boucle]+'"',True);
    try TypeduChamp[boucle]:=QRecupTypeChamp.Fields[0].AsString finally ferme(QRecupTypeChamp); end;
    end;

for compteur:=0 to TOBRes.Detail.Count-1 do
     begin
     for boucle:=1 to 3 do
          begin
          if (TypeDuchamp[boucle]='COMBO')then
               begin
               MaValeur:=RechDom(ChampToTT(NomDuChamp[Boucle]), TOBRes.detail[compteur].Getvalue(NomDuChamp[Boucle]),True) ;
               TOBRes.detail[compteur].PutValue(NomDuChamp[boucle], MaValeur);
               end;
          end;
     end;

Planning.TobRes := TobRes;
if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='GLA') or
    (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='AFF') then
    begin
    Planning.ResChampID := 'ARS_RESSOURCE';
    end
else
if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='GLR') or
    (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='RES') then
    begin
    Planning.ResChampID := 'AFF_AFFAIRE';
    end;

// Initalisation des colonnes d'entete
ColField1:=RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL1'), true);
ColField2:=RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL2'), true);
ColField3:=RechDom(sNomTablette, TobModelePlanning.detail[indice].getvalue('HPP_LIBCOL3'), true);
Planning.TokenFieldColFixed:=ColField1+';'+ColField2+';'+ColField3;
Planning.TokenSizeColFixed:=TobModelePlanning.detail[indice].Getvalue('HPP_TAILLECOLENTET');
Planning.TokenAlignColFixed:='C;C;C';

// ****************** Chargement des Etats **************************
TobEtats := Tob.Create ('liste des etats', Nil,-1);
stSQL := 'SELECT GA_ARTICLE, GA_LIBELLE FROM ARTICLE';
Q := OPENSQL(stSQL,True);
TOBEtats.LoadDetailDB('ARTICLE','','',Q,False);
ferme(Q);
if TOBEtats.Detail.Count > 0 then
    BEGIN
    TobDetEtat :=TOBEtats.Detail[0];
    TobDetEtat.AddchampSup ('BGC',True);
    TobDetEtat.AddchampSup ('COLOR',True);
    TobDetEtat.AddchampSup ('FONTE',True);
    for i := 0 to  TOBEtats.Detail.Count-1 do
        BEGIN
        TobDetEtat :=TOBEtats.Detail[i];
        TobDetEtat.putvalue ('BGC','clBlue');
        TobDetEtat.putvalue ('COLOR','$0080FFFF');
        TobDetEtat.putvalue ('FONTE','MS Sans Serif');
        END;
    END;

Planning.EtatChampCode  := 'GA_ARTICLE' ;
Planning.EtatChampLibelle := 'GA_LIBELLE' ;
Planning.EtatChampBackGroundColor := 'BGC' ;
Planning.EtatChampFontColor := 'COLOR' ;
Planning.EtatChampFontName := 'FONTE' ;
Planning.TobEtats := TobEtats ;

Planning.IntervalDebut := DateEnCours; //StrToDateTime ('01/11/2000') ;
Planning.IntervalFin := plusdate(DateEnCours,3,'M');    //StrToDateTime ('01/04/2001') ;

// Gestion des évenements
     Planning.ChampLibelleEvenement:='HEV_LIBELLE';
     Planning.ChampDateDebutEvenement:='HEV_DATEDEBUT';
     Planning.ChampDateDeFinEvenement:='HEV_DATEFIN';
     Planning.ChampStyleEvenement:='HEV_FORMEGRAPHIQUE';
     Planning.ChampCodeEvenement:='HEV_EVENEMENT';
     Planning.TobEvenements:=TobEvenements;

// **************  Chargement Tob items  *********************
(*
// Correspondance des champs
Planning.ChampdateDebut:='ACT_DATEACTIVITE';
Planning.ChampDateFin:= 'ACT_DATEACTIVITE';
Planning.ChampEtat:='ACT_CODEARTICLE';
sChampCritere := '';
if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='AFF') or
    (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='RES') then
    begin
    Planning.ChampLibelle:='ACT_LIBELLE';
    Planning.ChampHint :='ACT_LIBELLE';
    if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='AFF') then
        begin
        sChampCritere := 'ACT_AFFAIRE';
        Planning.ChampLineID:='ACT_RESSOURCE';
        end
    else
        begin
        sChampCritere := 'ACT_RESSOURCE';
        Planning.ChampLineID:='ACT_AFFAIRE';
        end;

//    stSQL := 'SELECT ACT_TYPEACTIVITE,ACT_RESSOURCE,ACT_DATEACTIVITE,ACT_AFFAIRE,ACT_LIBELLE,ACT_CODEARTICLE,ACT_QTE'
//             + ' FROM ACTIVITE';
    stSQL := 'SELECT * FROM ACTIVITE';
    end
else
    begin
    if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='GLR') then
        begin
        Planning.ChampLibelle:='ARS_LIBELLE';
        Planning.ChampHint :='ARS_LIBELLE';
        Planning.ChampLineID:='ACT_AFFAIRE';
//        stSQL := 'SELECT ARS_LIBELLE,ACT_DATEACTIVITE,ACT_RESSOURCE,ACT_AFFAIRE,ACT_TYPEACTIVITE,ACT_LIBELLE,ACT_QTE,ACT_CODEARTICLE'
//             + ' FROM ACTIVITE LEFT JOIN RESSOURCE ON ACT_RESSOURCE=ARS_RESSOURCE';
        stSQL := 'SELECT * FROM ACTIVITE LEFT JOIN RESSOURCE ON ACT_RESSOURCE=ARS_RESSOURCE';
        end
    else
    if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='GLA') then
        begin
        Planning.ChampLibelle:='LIBELLEAFFAIRE';
        Planning.ChampHint :='LIBELLEAFFAIRE';
        Planning.ChampLineID:='ACT_RESSOURCE';
//        stSQL := 'SELECT ACT_DATEACTIVITE,ACT_RESSOURCE,ACT_AFFAIRE,ACT_TYPEACTIVITE,ACT_LIBELLE,ACT_QTE,ACT_CODEARTICLE'
//             + ' FROM ACTIVITE';
        stSQL := 'SELECT * FROM ACTIVITE';
        end;
    end;

if (Critere<>'') then
    stSQL := stSQL + ' WHERE '+ sChampCritere +'="'+  Critere + '" AND ACT_DATEACTIVITE>="'+usdateTime(DateEnCours)+ '"'
else
    stSQL := stSQL + ' WHERE ACT_DATEACTIVITE>="'+usdateTime(DateEnCours)+ '"' ;

Q := OPENSQL(stSQL,True);
TOBACT := TOB.Create('liste des items', nil, -1);
TOBACT.LoadDetailDB('ACTIVITE', '', '', Q, False);
if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='GLA') then
    // Gestion du code affaire à afficher
    begin
    if TOBACT.Detail.Count > 0 then
        BEGIN
        TobDetACT :=TOBACT.Detail[0];
        TobDetACT.AddchampSup ('LIBELLEAFFAIRE',True);
        for i := 0 to  TOBACT.Detail.Count-1 do
            BEGIN
            TobDetACT :=TOBACT.Detail[i];
            TobDetACT.putvalue('LIBELLEAFFAIRE',CodeAffaireAffiche(TobDetACT.GetValue('ACT_AFFAIRE')));
            END;
        END;
    end;
ferme(Q);
///////*********
Planning.ChampdateDebut:='ATA_DATEDEBUT';
Planning.ChampDateFin:= 'ATA_DATEFIN';
Planning.ChampEtat:='ATA_ARTICLE';
sChampCritere := '';
if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='AFF') or
    (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='RES') then
    begin
    Planning.ChampLibelle:='ATA_LIBELLE';
    Planning.ChampHint :='ATA_LIBELLE';
    if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='AFF') then
        begin
        sChampCritere := 'ATA_AFFAIRE';
        Planning.ChampLineID:='ATR_RESSOURCE';
        end
    else
        begin
        sChampCritere := 'ATR_RESSOURCE';
        Planning.ChampLineID:='ATA_AFFAIRE';
        end;

    stSQL := 'SELECT * from tache left outer join tacheressource on (ata_tiers=atr_tiers) and (ata_numerotache=atr_numerotache)';
    end
else
    begin
    if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='GLR') then
        begin
        Planning.ChampLibelle:='ARS_LIBELLE';
        Planning.ChampHint :='ARS_LIBELLE';
        Planning.ChampLineID:='ATA_AFFAIRE';
        stSQL := 'SELECT * from tache left outer join tacheressource on (ata_tiers=atr_tiers) and (ata_numerotache=atr_numerotache) and left outer join ressource on (atr_ressource=ars_ressource)';
        end
    else
    if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='GLA') then
        begin
        Planning.ChampLibelle:='LIBELLEAFFAIRE';
        Planning.ChampHint :='LIBELLEAFFAIRE';
        Planning.ChampLineID:='ATR_RESSOURCE';
        stSQL := 'SELECT * from tache left outer join tacheressource on(ata_tiers=atr_tiers) and (ata_numerotache=atr_numerotache)';
        end;
    end;

if (Critere<>'') then
    stSQL := stSQL + ' WHERE '+ sChampCritere +'="'+  Critere + '" AND ATA_DATEDEBUT>="'+usdateTime(DateEnCours)+ '"'
else
    stSQL := stSQL + ' WHERE ATA_DATEDEBUT>="'+usdateTime(DateEnCours)+ '"' ;

Q := OPENSQL(stSQL,True);
TOBACT := TOB.Create('liste des items', nil, -1);
TOBACT.LoadDetailDB('la tob items', '', '', Q, False);
if (TobModelePlanning.Detail[indice].getValue('HPP_MODEGESTION')='GLA') then
    // Gestion du code affaire à afficher
    begin
    if TOBACT.Detail.Count > 0 then
        BEGIN
        TobDetACT :=TOBACT.Detail[0];
        TobDetACT.AddchampSup ('LIBELLEAFFAIRE',True);
        for i := 0 to  TOBACT.Detail.Count-1 do
            BEGIN
            TobDetACT :=TOBACT.Detail[i];
            TobDetACT.putvalue('LIBELLEAFFAIRE',CodeAffaireAffiche(TobDetACT.GetValue('ATA_AFFAIRE')));
            END;
        END;
    end;
ferme(Q);

Planning.TOBItems := TOBACT;


Planning.Activate := True;
end;
*)
//$$$jp
end;

initialization
end.

