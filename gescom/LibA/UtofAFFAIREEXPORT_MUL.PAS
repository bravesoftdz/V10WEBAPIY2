{***********UNITE*************************************************
Auteur  ...... : PL - JP
Créé le ...... : 03/09/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : AFFAIREEXPORT_MUL ()
Mots clefs ... : TOF;AFFAIREEXPORT_MUL
*****************************************************************}
Unit UtofAFFAIREEXPORT_MUL ;

Interface

Uses StdCtrls, EntGC
{$IFDEF EAGLCLIENT}
     ,MaineAGL,eMul
{$ELSE}
     ,DB, DBTables,FE_Main,mul
{$ENDIF}

 ,Controls, Classes, forms, sysutils, ComCtrls, M3FP, utob, Hstatus, dicoaf,
     HCtrls, HEnt1, HMsgBox, UTOF, UTofAfBaseCodeAffaire,TraducAffaire,
     utilmultrt, paramsoc, utilxls, utilword, Windows, ConfidentAffaire;

Type
  TOF_AFFAIREEXPORT_MUL = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;

    private
           PeriodeDateDeb, PeriodeDateFin :TDateTime;
           bSuivi                         :boolean;
           strWhereIn                     :string;

           procedure ExportAffairesTiers;
  end ;

procedure AFLanceFiche_Mul_ExportAff (strArgument:string='');

Implementation

uses PgiEnv, AffaireOle, FileCtrl, utofafarticle_select, utilGa;


procedure TOF_AFFAIREEXPORT_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFFAIREEXPORT_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFFAIREEXPORT_MUL.OnUpdate ;
begin
     inherited ;
{$IFDEF EAGLCLIENT}
        TraduitAFLibGridSt(TFMul(Ecran).FListe);
{$ELSE}
        TraduitAFLibGridDB(TFMul(Ecran).FListe);
{$ENDIF}
end ;

procedure TOF_AFFAIREEXPORT_MUL.OnLoad ;
begin
     inherited ;
end ;

procedure TOF_AFFAIREEXPORT_MUL.OnArgument (S : String);
var
   Critere     :string;
   ExportPath  :string;
   iPos        :integer;
begin
     inherited ;

     // Paramètres: date période
     PeriodeDateDeb := 0;
     PeriodeDateFin := 0;
     bSuivi         := FALSE;
     Critere := Trim (ReadTokenSt (S));
     while Critere <> '' do
     begin
          // $$$JP 12/05/03: mode suivi SDA ou mode libre (GA, normalement doit être obsolète bientôt)
          if Copy (Critere, 1, 9) = 'MODESUIVI' then
             bSuivi := TRUE;

          // Date début
          if Copy (Critere,1,7) = 'DATEDEB' then
             PeriodeDateDeb := StrToDate (Copy (Critere, 9, 10));

          // Date fin
          if Copy (Critere,1,7) = 'DATEFIN' then
             PeriodeDateFin := StrToDate (Copy (Critere, 9, 10));

          if Copy (Critere,1,6) = 'ASSIST' then
          begin
               Critere := Copy (Critere, 8, Length (Critere)-7);
               iPos := Pos (')', Critere);
               if iPos > 0 then
                  Critere := Copy (Critere, 1, iPos);
               strWhereIn := Critere;
          end;

          // Paramètre suivant
          Critere := Trim (ReadTokenSt (S));
     end;

     // Filtre sur les propositions ou pas
     if VH_GC.AFProposAct = FALSE then
         THEdit(GetControl('XX_WHERE')).Text := 'AFF_AFFAIRE0="A"'
     else
         THEdit(GetControl('XX_WHERE')).Text := '';

     // Par défaut, on prend le concept "affichage facturable" de l'utilisateur générant les fichiers excel
     // (même si ce n'est pas forcément le cas des utilisateurs qui saisissent dans les classeurs)
     TCHeckBox (GetControl ('CBFACT')).Checked := AffichageFNF;

     // Initialisation du répertoire d'export par défaut (si vide, répertoire dossier)
     ExportPath := GetParamSoc ('SO_AFEACTPATH');
     if ExportPath = '' then
     begin
       // V_PGI.PathDos et V_PGI_ENV.DosPath sont parfois valide, parfois non: on cherche celui qui est renseigné!
       ExportPath := V_PGI.DosPath;
       if ExportPath = '' then
          if V_PGI_ENV <> nil then
             if V_PGI_ENV.PathDos <> '' then
                ExportPath := V_PGI_ENV.PathDos;
     end;
     SetControlText ('REPERTOIREEXPORT', ExportPath);

     // Traduction du titre de la fiche
     Ecran.Caption := TraduitGA ('Saisie Décentralisée - Sélection des affaires');

{$IFDEF EAGLCLIENT}
     TraduitAFLibGridSt(TFMul(Ecran).FListe);
{$ELSE}
     TraduitAFLibGridDB(TFMul(Ecran).FListe);
{$ENDIF}
end ;

procedure TOF_AFFAIREEXPORT_MUL.OnClose;
begin
     inherited;
end ;

procedure TOF_AFFAIREEXPORT_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
      Aff:=THEdit(GetControl('AFF_AFFAIRE'));
      Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
      Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
      Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
      Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
      Aff4:=THEdit(GetControl('AFF_AVENANT'));
      Tiers:=THEdit(GetControl('AFF_TIERS'));
end;



//------------------------------------------------
//              ExportAffairesTiers
//------------------------------------------------
procedure TOF_AFFAIREEXPORT_MUL.ExportAffairesTiers;
Var
    F                                                            :TFMul;
    TobMAff, TobAffExp, TobAffFille, TobTiersExp, TobTiersFille  :TOB;
    TobRessExp, TobPrestExp, {TobPrestFille,} TobFraisExp        :TOB;
    TOBArticle, TOBUnArticle                                     :TOB;     //$$$jp 04/2003
    strTypeArticle                                               :string;  //$$$jp 04/2003
    {TobFraisFille, }TobFourExp, {TobFourFille, }TobUniteExp     :TOB;
    wi,iNbChampAff                                               :integer;
    Q                                                            :TQuery;
    sRepExp                                                      :string;
    tCleAff                                                      :array of string;
    IndTabAff                                                    :integer;
    {bAsked,} bIsNew{, bExportNonEnc}                            :boolean;
    i                                                            :integer;
    NomFichier                                                   :string;
    SDAExcel                                                     :OleVariant;
    SDABook, SDASheet                                            :variant;
    strReq                                                       :string;
    bCopy                                                        :boolean;
    bAffFact                                                     :boolean;
    bUnPourTous                                                  :boolean;
    CreeListe                                                    :TStringList;
// $$$ JP 07/07/03   AppliPathStd                                                 :string;
    bCanceled                                                    :boolean;
    strFormeJuridique                                            :string;
    iNumLigne                                                    :integer;
begin
     // $$$ JP 07/07/03: amélioration identification STD ET DAT
//{$IFNDEF EAGLCLIENT}
//{$IFDEF GIGI}
  //     AppliPathStd := 'C:\PGI00\STD\GIS5';
//{$ELSE}
  //     AppliPathStd := 'C:\PGI00\STD\GAS5';
//{$ENDIF}
//{$ENDIF}
     bCanceled := FALSE;

     F := TFMul (Ecran);
     if (F.FListe.NbSelected=0) and (not F.FListe.AllSelected) then
     begin
          MessageAlerte('Aucun élément sélectionné');
          exit;
     end;

     // Nombre de parties du code affaire: prise en compte du nb de parties gérées et du nb de parties visibles
     iNbChampAff := VH_GC.CleAffaire.NbPartie;
     if iNbChampAff=3 then
     begin
          if Not VH_GC.CleAffaire.Co2Visible then
             iNbChampAff := iNbChampAff-1;
          if Not VH_GC.CleAffaire.Co3Visible then
             iNbChampAff := iNbChampAff-1;
     end
     else
         if iNbChampAff=2 then
            if not VH_GC.CleAffaire.Co2Visible then
               iNbChampAff := iNbChampAff-1;

     // jp 18/12/2002: controle d'existance du répertoire d'export
     sRepExp := GetControlText('REPERTOIREEXPORT');
     if (sRepExp = '') then
     begin
          MessageAlerte ('Veuillez renseigner le répertoire d''export des fichiers' + #10 + ' Vous pouvez également saisir un répertoire par défaut dans les paramètres société');
          SetFocusControl('REPERTOIREEXPORT');
          exit;
     end
     else
     begin
          if not DirectoryExists (sRepExp) then
          begin
               PGIInfoAf ('Le répertoire: "' + sRepExp + '" n''existe pas', TFMUL(Ecran).caption);
               exit;
          end;
     end;

     // Création des tob
     TobMAff     := Tob.Create ('les affaires',NIL,-1);
     TobAffExp   := Tob.Create ('Les Missions',NIL,-1); //mcd 28/01/03 chgmt nom tof
     TobTiersExp := Tob.Create ('Les Clients',NIL,-1);  //mcd 28/01/03 chgmt nom tof
     TobRessExp  := Tob.Create('Les Ressources',NIL,-1);
     TobPrestExp := Tob.Create('Les Prestations',NIL,-1);
     TobFraisExp := Tob.Create('Les Frais',NIL,-1);
     TobFourExp  := Tob.Create('Les Fournitures',NIL,-1);
     TobUniteExp := Tob.Create('Les unités temps',NIL,-1);

     TraiteEnregMulListe (F, 'AFF_AFFAIRE','AFFAIRE', TobMAff, True);

     // constitution des tob d'export des affaires et des clients associés
     try
        InitMove(TobMaff.Detail.count*2, ' ');
        wi := 0;
        while wi < TobMAff.Detail.count  do
        begin
             MoveCur(False);
             IndTabAff:=0;
             TobAffFille:=Tob.Create('une mission',TobAffExp,-1);
             TobAffFille.AddChampSup('AFF_TIERS',false);
             SetLength(tCleAff, IndTabAff+1);
             tCleAff[IndTabAff]:= 'AFF_TIERS';
             TobAffFille.AddChampSup('AFF_AFFAIRE0',false);
             Inc(IndTabAff); SetLength(tCleAff, IndTabAff+1);
             tCleAff[IndTabAff]:= 'AFF_AFFAIRE0';
             TobAffFille.AddChampSup('AFF_AFFAIRE1',false);
             Inc(IndTabAff); SetLength(tCleAff, IndTabAff+1);
             tCleAff[IndTabAff]:= 'AFF_AFFAIRE1';
             TobAffFille.AddChampSup('AFF_AFFAIRE2',false);
             Inc(IndTabAff); SetLength(tCleAff, IndTabAff+1);
             tCleAff[IndTabAff]:= 'AFF_AFFAIRE2';
             TobAffFille.AddChampSup('AFF_AFFAIRE3',false);
             Inc(IndTabAff); SetLength(tCleAff, IndTabAff+1);
             tCleAff[IndTabAff]:= 'AFF_AFFAIRE3';
             TobAffFille.AddChampSup('AFF_AVENANT',false);
             Inc(IndTabAff); SetLength(tCleAff, IndTabAff+1);
             tCleAff[IndTabAff]:= 'AFF_AVENANT';
             TobAffFille.AddChampSup('AFF_RESPONSABLE',false);
             Inc(IndTabAff); SetLength(tCleAff, IndTabAff+1);
             tCleAff[IndTabAff]:= 'AFF_RESPONSABLE';
             TobAffFille.AddChampSup('AFF_LIBELLE',false);
             Inc(IndTabAff); SetLength(tCleAff, IndTabAff+1);
             tCleAff[IndTabAff]:= 'AFF_LIBELLE';
             TobAffFille.AddChampSup('AFF_DATEDEBUT',false);
             Inc(IndTabAff); SetLength(tCleAff, IndTabAff+1);
             tCleAff[IndTabAff]:= 'AFF_DATEDEBUT';
             TobAffFille.AddChampSup('AFF_DATEFIN',false);
             Inc(IndTabAff); SetLength(tCleAff, IndTabAff+1);
             tCleAff[IndTabAff]:= 'AFF_DATEFIN';
             TobAffFille.AddChampSup('AFF_ETATAFFAIRE',false);
             Inc(IndTabAff); SetLength(tCleAff, IndTabAff+1);
             tCleAff[IndTabAff]:= 'AFF_ETATAFFAIRE';
             TobAffFille.InitValeurs ;
             TOBCopyFieldValues(TobMAff.Detail[wi],TobAffFille,tCleAff);

             // Génération clients si pas déjà existant
             if (TobAffFille <> nil) and (TobTiersExp.FindFirst(['T_TIERS'],[TobMaff.Detail[wi].GetValue('AFF_TIERS')],true)=nil) then
             begin
                 TobTiersFille:=Tob.Create('un client',TobTiersExp,-1);
                 TobTiersFille.AddChampSup('T_TIERS',false);
                 TobTiersFille.AddChampSup('T_JURIDIQUE',false); // $$$jp 18/04/02
                 TobTiersFille.AddChampSup('T_LIBELLE',false);
                 TobTiersFille.AddChampSup('T_CODEPOSTAL',false);
                 TobTiersFille.AddChampSup('T_VILLE',false);
                 TobTiersFille.AddChampSup('T_AUXILIAIRE',false);
                 TobTiersFille.InitValeurs ;
                 Q := nil;
                 try
                    Q:=OpenSQL('SELECT T_TIERS,T_JURIDIQUE,T_LIBELLE,T_PRENOM,T_CODEPOSTAL,T_VILLE,T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+TobMaff.Detail[wi].GetValue('AFF_TIERS')+'"',True) ;
                    If (not Q.EOF) then
                        TobTiersFille.SelectDB('',Q)
                    else
                        TobTiersFille.Free;
                 finally
                        Ferme(Q);
                 end;
             end;
             inc (wi);
        end;
     finally
            TobMaff.free;TobMaff:=Nil;
            FiniMove;
     end;

     // Tri des affaires par code client - code mission
     TobAffExp.Detail.Sort ('AFF_TIERS;AFF_AFFAIRE0;AFF_AFFAIRE1;-AFF_AFFAIRE2;AFF_AFFAIRE3;AFF_AVENANT');

     // $$$JP 12/05/03: si pas mode suivi SDA, un seul classeur multi-assistants
     if bSuivi = FALSE then
     begin
          // Requête de sélection des ressources à exporter
          strReq      := 'SELECT ARS_RESSOURCE,ARS_LIBELLE2,ARS_LIBELLE FROM RESSOURCE';
          bUnPourTous := TRUE;
     end
     else
     begin
          // Un classeur par assitant, ou multi-assistant
          bUnPourTous := GetParamSoc ('SO_AFEACTUNPOURTOUS');

          // Requête pour TOB des ressources
          try
             if strWhereIn <> '*' then
                 strReq := 'SELECT ASA_RESSOURCE FROM SUIVIEACT WHERE ASA_DATEDEB = "' + UsDateTime (PeriodeDateDeb) + '" AND ASA_DATEFIN = "' + UsDateTime (PeriodeDateFin) + '" AND ASA_ETATEACT<>"ES0" AND ASA_RESSOURCE IN ' + strWhereIn
             else
                 strReq := 'SELECT ASA_RESSOURCE FROM SUIVIEACT WHERE ASA_DATEDEB = "' + UsDateTime (PeriodeDateDeb) + '" AND ASA_DATEFIN = "' + UsDateTime (PeriodeDateFin) + '" AND ASA_ETATEACT<>"ES0"';
             if (ExisteSQL (strReq) = TRUE) and (PgiAsk ('Certains assistants ont déjà été exportés pour cette période. Désirez-vous tout de même continuer ?', Ecran.Caption) = mrYes) then
             begin
                  if strWhereIn <> '*' then
                      strReq := 'SELECT ARS_RESSOURCE,ARS_LIBELLE2,ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE IN ' + strWhereIn
                  else
                      strReq := 'SELECT ARS_RESSOURCE,ARS_LIBELLE2,ARS_LIBELLE FROM RESSOURCE';
             end
             else
             begin
                  if strWhereIn <> '*' then
                      strReq := 'SELECT ARS_RESSOURCE,ARS_LIBELLE2,ARS_LIBELLE FROM RESSOURCE,SUIVIEACT WHERE ASA_RESSOURCE=ARS_RESSOURCE AND ASA_DATEDEB = "' + UsDateTime (PeriodeDateDeb) + '" AND ASA_DATEFIN = "' + UsDateTime (PeriodeDateFin) + '" AND ASA_ETATEACT="ES0" AND ARS_RESSOURCE IN ' + strWhereIn
                  else
                      strReq := 'SELECT ARS_RESSOURCE,ARS_LIBELLE2,ARS_LIBELLE FROM RESSOURCE,SUIVIEACT WHERE ASA_RESSOURCE=ARS_RESSOURCE AND ASA_DATEDEB = "' + UsDateTime (PeriodeDateDeb) + '" AND ASA_DATEFIN = "' + UsDateTime (PeriodeDateFin) + '" AND ASA_ETATEACT="ES0"';
             end;
          except
                strReq := 'SELECT ARS_RESSOURCE,ARS_LIBELLE2,ARS_LIBELLE FROM RESSOURCE';
          end;
     end;

     // Constitution tob des ressources (assistants)
     Q := nil;
     try
        Q := OpenSQL (strReq,True);
        if not Q.EOF then
             TobRessExp.LoadDetailDB ('','','', Q, true)
        else
        begin
             PgiInfo ('Aucun classeur à générer', 'Saisie décentralisée GI PGI');
             TobMAff.Free;
             TobAffExp.Free;
             TobTiersExp.Free;
             TobRessExp.Free;
             TobPrestExp.Free;
             TobFraisExp.Free;
             TobFourExp.Free;
             TobUniteExp.Free;
             exit;
        end;
     finally
            Ferme(Q);
     end;

     // $$$jp 04/2003: sélection manuelle des articles si demandé
     if GetParamSoc ('SO_AFEACTSELART') = TRUE then
     begin
          TOBArticle := TOB (strtoint(AFLanceFiche_Mul_Article ('NATURE=SDA')));
          if TOBArticle <> nil then
          begin
               while TOBArticle.Detail.Count > 0 do
               begin
                    TOBUnArticle := TOBArticle.Detail [0];
                    strTypeArticle := TOBUnArticle.GetValue ('GA_TYPEARTICLE');
                    if strTypeArticle = 'PRE' then
                        TOBUnArticle.ChangeParent (TOBPrestExp, -1)
                    else
                        if strTypeArticle = 'FRA' then
                            TOBUnArticle.ChangeParent (TOBFraisExp, -1)
                        else
                            if strTypeArticle = 'MAR' then
                                TOBUnArticle.ChangeParent (TOBFourExp, -1)
                            else
                                TOBUnArticle.Free;
               end;
               TOBArticle.Free;
          end
          else
          begin
               MessageAlerte ('Aucun article sélectionné. Génération annulée.');
               TobMAff.Free;
               TobAffExp.Free;
               TobTiersExp.Free;
               TobRessExp.Free;
               TobPrestExp.Free;
               TobFraisExp.Free;
               TobFourExp.Free;
               TobUniteExp.Free;
               exit;
          end;
     end
     else
     begin
          // constitution tob d'export des prestations
          Q := nil;
          try
             // $$$JP 23/05/03: qté d'activité, et non qté de vente
             Q := OpenSQL('SELECT GA_CODEARTICLE,GA_LIBELLE,GA_QUALIFUNITEACT,GA_FAMILLENIV1,GA_ACTIVITEREPRISE FROM ARTICLE WHERE GA_TYPEARTICLE="PRE"',True) ;
             //Q:=OpenSQL('SELECT GA_CODEARTICLE,GA_LIBELLE,GA_QUALIFUNITEVTE,GA_FAMILLENIV1,GA_ACTIVITEREPRISE FROM ARTICLE WHERE GA_TYPEARTICLE="PRE"',True) ;
             If (Not Q.EOF) then
                TobPrestExp.LoadDetailDB('','','', Q, true);
          finally
                 Ferme(Q);
          end;

          // constitution tob d'export des frais
          Q := nil;
          try
             // $$$JP 23/05/03: qté d'activité, et non qté de vente
             Q := OpenSQL('SELECT GA_CODEARTICLE,GA_LIBELLE,GA_QUALIFUNITEACT,GA_FAMILLENIV1,GA_PVHT,GA_ACTIVITEREPRISE FROM ARTICLE WHERE GA_TYPEARTICLE="FRA"',True) ;
             //Q := OpenSQL('SELECT GA_CODEARTICLE,GA_LIBELLE,GA_QUALIFUNITEVTE,GA_FAMILLENIV1,GA_PVHT,GA_ACTIVITEREPRISE FROM ARTICLE WHERE GA_TYPEARTICLE="FRA"',True) ;
             If (Not Q.EOF) then
                TobFraisExp.LoadDetailDB('','','', Q, true);
          finally
                 Ferme(Q);
          end;

          // constitution tob d'export des fournitures
          Q := nil;
          try
             // $$$JP 23/05/03: qté d'activité, et non qté de vente
             Q := OpenSQL('SELECT GA_CODEARTICLE,GA_LIBELLE,GA_QUALIFUNITEACT,GA_FAMILLENIV1,GA_PVHT,GA_ACTIVITEREPRISE FROM ARTICLE WHERE GA_TYPEARTICLE="MAR"',True) ;
             //Q := OpenSQL('SELECT GA_CODEARTICLE,GA_LIBELLE,GA_QUALIFUNITEVTE,GA_FAMILLENIV1,GA_PVHT,GA_ACTIVITEREPRISE FROM ARTICLE WHERE GA_TYPEARTICLE="MAR"',True) ;
             If (Not Q.EOF) then
                TobFourExp.LoadDetailDB('','','', Q, true);
          finally
                 Ferme(Q);
          end;
     end;

     // constitution tob d'export des unités de temps (pour permettre calcul précis des temps passés dans unité de réf.)
     Q := nil;
     try
        // $$$JP 10/04/2003: on ne veut que la quotité jour et heure, pour connaitre le nb d'heure dans un jour
        Q := OpenSQL('SELECT GME_MESURE, GME_LIBELLE, GME_QUOTITE FROM MEA WHERE GME_QUALIFMESURE="TEM" AND (GME_MESURE="H" OR GME_MESURE="J") ORDER BY GME_MESURE',True) ;
//        Q := OpenSQL('SELECT GME_MESURE, GME_LIBELLE, GME_QUOTITE FROM MEA WHERE GME_QUALIFMESURE="TEM"',True) ;
        If (not Q.EOF) then
           TobUniteExp.LoadDetailDB ('','','', Q, true);
     finally
            Ferme(Q);
     end;

    // Affichage colonne facturable ou non
    bAffFact := TCHeckBox (GetControl ('CBFACT')).Checked;

    // Lien OLE Excel pour traitement des fichiers .xls à générer
    try
          // Fiche inutilisable car traitement WORD en "thread de fond" apparement
          Ecran.Enabled := FALSE;

          // Ouverture lien ole vers Excel
          OpenExcel (FALSE, SDAExcel, bIsNew);
          if VarIsEmpty (SDAExcel) = TRUE then
             PgiInfo ('Export des classeurs annulé' + #10 + ' Connexion à Microsoft Excel impossible', Ecran.Caption)
          else
          begin
               CreeListe := nil;
               try
                  // Traitement de tous les enregistrements des 5/6 tobs, + ouverture excel, + màj titre colonnes
                  if bUnPourTous = TRUE then
//                  if GetParamSoc ('SO_AFEACTUNPOURTOUS') = TRUE then
                      InitMove (2 + TOBRessExp.Detail.Count + TobTiersExp.Detail.Count + TobAffExp.Detail.Count + TobPrestExp.Detail.Count + TobFraisExp.Detail.Count + TobFourExp.Detail.Count + TobUniteExp.Detail.Count, '')
                  else
                      InitMove (2 + TobTiersExp.Detail.Count + TobAffExp.Detail.Count + TobPrestExp.Detail.Count + TobFraisExp.Detail.Count + TobFourExp.Detail.Count + TobUniteExp.Detail.Count, '');

                  // Ouverture du fichier modèle
                  SDAExcel.WorkBooks.Open (GetAppliPathStd + '\SDA.XLS');  // $$$ JP 07/07/03: amélioration identification STD ET DAT
                  SDABook  := SDAExcel.ActiveWorkBook;
                  MoveCur (False);

                  // Libellés des parties code affaire
                  SDABook.Sheets ['Affaires'].Cells [1, 3].Value := string (VH_GC.CleAffaire.Co1Lib);
                  SDABook.Sheets ['Affaires'].Cells [1, 4].Value := string (VH_GC.CleAffaire.Co2Lib);
                  SDABook.Sheets ['Affaires'].Cells [1, 5].Value := string (VH_GC.CleAffaire.Co3Lib);
                  SDABook.Sheets ['Saisie activité'].Cells [1, 7].Value := UpperCase (string (VH_GC.CleAffaire.Co1Lib));
                  SDABook.Sheets ['Saisie activité'].Cells [1, 8].Value := UpperCase (string (VH_GC.CleAffaire.Co2Lib));
                  SDABook.Sheets ['Saisie activité'].Cells [1, 9].Value := UpperCase (string (VH_GC.CleAffaire.Co3Lib));

                  // $$$JP 10/04/2003 - feuilles d'impression inutiles désormais
{                  SDABook.Sheets ['Imp_Prest'].Cells [2, 7].Value := string (VH_GC.CleAffaire.Co1Lib);
                  SDABook.Sheets ['Imp_Prest'].Cells [2, 8].Value := string (VH_GC.CleAffaire.Co2Lib);
                  SDABook.Sheets ['Imp_Prest'].Cells [2, 9].Value := string (VH_GC.CleAffaire.Co3Lib);
                  SDABook.Sheets ['Imp_Frais'].Cells [2, 7].Value := string (VH_GC.CleAffaire.Co1Lib);
                  SDABook.Sheets ['Imp_Frais'].Cells [2, 8].Value := string (VH_GC.CleAffaire.Co2Lib);
                  SDABook.Sheets ['Imp_Frais'].Cells [2, 9].Value := string (VH_GC.CleAffaire.Co3Lib);
                  SDABook.Sheets ['Imp_Fourn'].Cells [2, 7].Value := string (VH_GC.CleAffaire.Co1Lib);
                  SDABook.Sheets ['Imp_Fourn'].Cells [2, 8].Value := string (VH_GC.CleAffaire.Co2Lib);
                  SDABook.Sheets ['Imp_Fourn'].Cells [2, 9].Value := string (VH_GC.CleAffaire.Co3Lib);
                  // $$$jp: 07/01/03 - nouvelle feuille d'impression (globale)
                  SDABook.Sheets ['Imp_Global'].Cells [2, 7].Value := string (VH_GC.CleAffaire.Co1Lib);
                  SDABook.Sheets ['Imp_Global'].Cells [2, 8].Value := string (VH_GC.CleAffaire.Co2Lib);
                  SDABook.Sheets ['Imp_Global'].Cells [2, 9].Value := string (VH_GC.CleAffaire.Co3Lib);}

                  MoveCur (False);

                  // Mise à jour feuille des clients (avec code NOUVEAU si saisie sur nouvelle mission possible)
                  SDASheet := SelectSheet (SDABook, 'Clients'); //SDABook.ActiveSheet;
                  for i := 0 to TobTiersExp.Detail.Count-1 do
                  begin
                       SDASheet.Cells [i+2, 1].Value := string (TobTiersExp.Detail [i].GetValue ('T_TIERS'));

                       // $$$JP 10/04/2003: libellé forme juridique, et non le code
                       strFormeJuridique := TobTiersExp.Detail [i].GetValue ('T_JURIDIQUE');
                       if strFormeJuridique <> '' then
                           strFormeJuridique := Trim (RechDom('TTFORMEJURIDIQUE', strFormeJuridique, FALSE));
                       SDASheet.Cells [i+2, 2].Value := strFormeJuridique;
//                       SDASheet.Cells [i+2, 2].Value := string (TobTiersExp.Detail [i].GetValue ('T_JURIDIQUE'));

                       if TobTiersExp.Detail [i].GetValue ('T_PRENOM') <> '' then
                           SDASheet.Cells [i+2, 3].Value := string (TobTiersExp.Detail [i].GetValue ('T_LIBELLE')) + ' ' + string (TobTiersExp.Detail [i].GetValue ('T_PRENOM'))
                       else
                           SDASheet.Cells [i+2, 3].Value := string (TobTiersExp.Detail [i].GetValue ('T_LIBELLE'));
                       SDASheet.Cells [i+2, 4].Value := string (TobTiersExp.Detail [i].GetValue ('T_CODEPOSTAL'));
                       SDASheet.Cells [i+2, 5].Value := string (TobTiersExp.Detail [i].GetValue ('T_VILLE'));
                       SDASheet.Cells [i+2, 6].Value := string (TobTiersExp.Detail [i].GetValue ('T_AUXILIAIRE'));
                       MoveCur (False);
                  end;
                  if GetParamSoc ('SO_AFEACTNEWAFF') = TRUE then
                       SDASheet.Cells [TobTiersExp.Detail.Count+2, 1].Value := 'NOUVEAU';

                  // Mise à jour feuille des missions
                  SDASheet := SelectSheet (SDABook, 'Affaires'); //SDABook.ActiveSheet;
                  for i := 0 to TobAffExp.Detail.Count-1 do
                  begin
                       iNumLigne := i+2;
                       SDASheet.Cells [iNumLigne, 1].Value := string (TobAffExp.Detail [i].GetValue ('AFF_TIERS'));
                       SDASheet.Cells [iNumLigne, 2].Value := string (TobAffExp.Detail [i].GetValue ('AFF_AFFAIRE0'));
                       SDASheet.Cells [iNumLigne, 3].Value := string (TobAffExp.Detail [i].GetValue ('AFF_AFFAIRE1'));
                       SDASheet.Cells [iNumLigne, 4].Value := string (TobAffExp.Detail [i].GetValue ('AFF_AFFAIRE2'));
                       SDASheet.Cells [iNumLigne, 5].Value := string (TobAffExp.Detail [i].GetValue ('AFF_AFFAIRE3'));
                       SDASheet.Cells [iNumLigne, 6].Value := string (TobAffExp.Detail [i].GetValue ('AFF_AVENANT'));
                       SDASheet.Cells [iNumLigne, 7].Value := string (TobAffExp.Detail [i].GetValue ('AFF_RESPONSABLE'));
                       SDASheet.Cells [iNumLigne, 8].Value := string (TobAffExp.Detail [i].GetValue ('AFF_LIBELLE'));
                       SDASheet.Cells [iNumLigne, 9].Value := string (TobAffExp.Detail [i].GetValue ('AFF_DATEDEBUT'));
                       SDASheet.Cells [iNumLigne,10].Value := string (TobAffExp.Detail [i].GetValue ('AFF_DATEFIN'));
                       SDASheet.Cells [iNumLigne,11].Value := string (TobAffExp.Detail [i].GetValue ('AFF_ETATAFFAIRE'));
                       SDASheet.Cells [iNumLigne,12].Formula := Format('=A%d&B%d&C%d&D%d&E%d&F%d', [iNumLigne,iNumLigne,iNumLigne,iNumLigne,iNumLigne,iNumLigne]);
                       MoveCur (False);
                  end;

                  // Mise à jour feuille des prestations
                  SDASheet := SelectSheet (SDABook, 'Prest'); //SDABook.ActiveSheet;
                  for i := 0 to TobPrestExp.Detail.Count-1 do
                  begin
                       SDASheet.Cells [i+2, 1].Value := string (TobPrestExp.Detail [i].GetValue ('GA_CODEARTICLE'));
                       SDASheet.Cells [i+2, 2].Value := string (TobPrestExp.Detail [i].GetValue ('GA_LIBELLE'));
                       // $$$JP 15/05/03: unité ACT et non pas unité VTE
                       SDASheet.Cells [i+2, 3].Value := string (TobPrestExp.Detail [i].GetValue ('GA_QUALIFUNITEACT'));
                       // $$$JP fin
                       SDASheet.Cells [i+2, 4].Value := string (TobPrestExp.Detail [i].GetValue ('GA_FAMILLENIV1'));
                       SDASheet.Cells [i+2, 5].Value := string (TobPrestExp.Detail [i].GetValue ('GA_ACTIVITEREPRISE'));
                       MoveCur (False);
                  end;

                  // Mise à jour feuille des frais
                  SDASheet := SelectSheet (SDABook, 'Frais'); //SDABook.ActiveSheet;
                  for i := 0 to TobFraisExp.Detail.Count-1 do
                  begin
                       SDASheet.Cells [i+2, 1].Value := string (TobFraisExp.Detail [i].GetValue ('GA_CODEARTICLE'));
                       SDASheet.Cells [i+2, 2].Value := string (TobFraisExp.Detail [i].GetValue ('GA_LIBELLE'));
                       // $$$JP 15/05/03: unité ACT et non pas unité VTE
                       SDASheet.Cells [i+2, 3].Value := string (TobFraisExp.Detail [i].GetValue ('GA_QUALIFUNITEACT'));
                       // $$$JP fin
                       SDASheet.Cells [i+2, 4].Value := string (TobFraisExp.Detail [i].GetValue ('GA_FAMILLENIV1'));
                       SDASheet.Cells [i+2, 5].Value := string (TobFraisExp.Detail [i].GetValue ('GA_PVHT'));
                       SDASheet.Cells [i+2, 6].Value := string (TobFraisExp.Detail [i].GetValue ('GA_ACTIVITEREPRISE'));
                       MoveCur (False);
                  end;

                  // Mise à jour feuille des fournitures
                  SDASheet := SelectSheet (SDABook, 'Fourn'); //SDABook.ActiveSheet;
                  for i := 0 to TobFourExp.Detail.Count-1 do
                  begin
                       SDASheet.Cells [i+2, 1].Value := string (TobFourExp.Detail [i].GetValue ('GA_CODEARTICLE'));
                       SDASheet.Cells [i+2, 2].Value := string (TobFourExp.Detail [i].GetValue ('GA_LIBELLE'));
                       // $$$JP 15/05/03: unité ACT et non pas unité VTE
                       SDASheet.Cells [i+2, 3].Value := string (TobFourExp.Detail [i].GetValue ('GA_QUALIFUNITEACT'));
                       // $$$JP fin
                       SDASheet.Cells [i+2, 4].Value := string (TobFourExp.Detail [i].GetValue ('GA_FAMILLENIV1'));
                       SDASheet.Cells [i+2, 5].Value := string (TobFourExp.Detail [i].GetValue ('GA_PVHT'));
                       SDASheet.Cells [i+2, 6].Value := string (TobFourExp.Detail [i].GetValue ('GA_ACTIVITEREPRISE'));
                       MoveCur (False);
                  end;

                  // Mise à jour feuille des unités de temps (la 1ère ligne est l'unité de référence)

                  // $$$JP 10/04/2003 - nb d'heure par jour
                  SDASheet := SelectSheet (SDABook, 'Param');
                  if (TOBUniteExp.Detail.Count = 2) and (TOBUniteExp.Detail [0].GetValue ('GME_QUOTITE') <> 0) then
                      SDASheet.Cells [20, 2].Value := TOBUniteExp.Detail [1].GetValue ('GME_QUOTITE') / TOBUniteExp.Detail [0].GetValue ('GME_QUOTITE')
                  else
                      SDASheet.Cells [20, 2].Value := 7;
                  MoveCur (False);
{                  SDASheet := SelectSheet (SDABook, 'Temps');
                  for i := 0 to TobUniteExp.Detail.Count-1 do
                  begin
                       SDASheet.Cells [i+2, 1].Value := string (TobUniteExp.Detail [i].GetValue ('GME_MESURE'));
                       SDASheet.Cells [i+2, 2].Value := string (TobUniteExp.Detail [i].GetValue ('GME_LIBELLE'));
                       SDASheet.Cells [i+2, 3].Value := string (TobUniteExp.Detail [i].GetValue ('GME_QUOTITE'));
                       MoveCur (False);
                  end;}

                  // Si un classeur pour tous les assistants, on crée la feuille Assistants
                  if bUnPourTous = TRUE then
//                  if GetParamSoc ('SO_AFEACTUNPOURTOUS') = TRUE then
                  begin
                       SDASheet := SelectSheet (SDABook, 'Assistants');
                       for i := 0 to TobRessExp.Detail.Count-1 do
                       begin
                            SDASheet.Cells [i+2, 1].Value := string (TobRessExp.Detail [i].GetValue ('ARS_RESSOURCE'));
                            SDASheet.Cells [i+2, 2].Value := string (TobRessExp.Detail [i].GetValue ('ARS_LIBELLE2'));
                            SDASheet.Cells [i+2, 3].Value := string (TobRessExp.Detail [i].GetValue ('ARS_LIBELLE'));
                            MoveCur (False);
                       end;
                  end;

                  // On cache les parties code affaire en fonction du paramètrage GI
                  // $$$JP 14/04/2003: légère optimisation
                  SDABook.Sheets ['Saisie activité'].Columns ['H:H'].EntireColumn.Hidden := iNbChampAff < 2;
                  SDABook.Sheets ['Saisie activité'].Columns ['I:I'].EntireColumn.Hidden := iNbChampAff < 3;
{                  if iNbChampAff = 3 then
                  begin
                       SDABook.Sheets ['Saisie activité'].Columns ['H:H'].EntireColumn.Hidden := FALSE;
                       SDABook.Sheets ['Saisie activité'].Columns ['I:I'].EntireColumn.Hidden := FALSE;
                       // $$$JP 10/04/2003 - feuilles d'impression inutiles désormais
//                       SDABook.Sheets ['Imp_Prest'].Columns ['H:H'].EntireColumn.Hidden := FALSE;
  //                     SDABook.Sheets ['Imp_Prest'].Columns ['I:I'].EntireColumn.Hidden := FALSE;
    //                   SDABook.Sheets ['Imp_Frais'].Columns ['H:H'].EntireColumn.Hidden := FALSE;
      //                 SDABook.Sheets ['Imp_Frais'].Columns ['I:I'].EntireColumn.Hidden := FALSE;
        //               SDABook.Sheets ['Imp_Fourn'].Columns ['H:H'].EntireColumn.Hidden := FALSE;
          //             SDABook.Sheets ['Imp_Fourn'].Columns ['I:I'].EntireColumn.Hidden := FALSE;
                       // $$$jp: 07/01/03 - nouvelle feuille d'impression (globale)
            //           SDABook.Sheets ['Imp_Global'].Columns ['H:H'].EntireColumn.Hidden := FALSE;
              //         SDABook.Sheets ['Imp_Global'].Columns ['I:I'].EntireColumn.Hidden := FALSE;
                  end
                  else
                      if iNbChampAff = 2 then
                      begin
                           SDABook.Sheets ['Saisie activité'].Columns ['H:H'].EntireColumn.Hidden := FALSE;
                           SDABook.Sheets ['Saisie activité'].Columns ['I:I'].EntireColumn.Hidden := TRUE;
                           // $$$JP 10/04/2003 - feuilles d'impression inutiles désormais
//                           SDABook.Sheets ['Imp_Prest'].Columns ['H:H'].EntireColumn.Hidden := FALSE;
  //                         SDABook.Sheets ['Imp_Prest'].Columns ['I:I'].EntireColumn.Hidden := TRUE;
    //                       SDABook.Sheets ['Imp_Frais'].Columns ['H:H'].EntireColumn.Hidden := FALSE;
      //                     SDABook.Sheets ['Imp_Frais'].Columns ['I:I'].EntireColumn.Hidden := TRUE;
        //                   SDABook.Sheets ['Imp_Fourn'].Columns ['H:H'].EntireColumn.Hidden := FALSE;
          //                 SDABook.Sheets ['Imp_Fourn'].Columns ['I:I'].EntireColumn.Hidden := TRUE;
                           // $$$jp: 07/01/03 - nouvelle feuille d'impression (globale)
            //               SDABook.Sheets ['Imp_Global'].Columns ['H:H'].EntireColumn.Hidden := FALSE;
              //             SDABook.Sheets ['Imp_Global'].Columns ['I:I'].EntireColumn.Hidden := TRUE;
                      end
                      else
                      begin
                           SDABook.Sheets ['Saisie activité'].Columns ['H:H'].EntireColumn.Hidden := TRUE;
                           SDABook.Sheets ['Saisie activité'].Columns ['I:I'].EntireColumn.Hidden := TRUE;
                           // $$$JP 10/04/2003 - feuilles d'impression inutiles désormais
//                           SDABook.Sheets ['Imp_Prest'].Columns ['H:H'].EntireColumn.Hidden := TRUE;
  //                         SDABook.Sheets ['Imp_Prest'].Columns ['I:I'].EntireColumn.Hidden := TRUE;
    //                       SDABook.Sheets ['Imp_Frais'].Columns ['H:H'].EntireColumn.Hidden := TRUE;
      //                     SDABook.Sheets ['Imp_Frais'].Columns ['I:I'].EntireColumn.Hidden := TRUE;
        //                   SDABook.Sheets ['Imp_Fourn'].Columns ['H:H'].EntireColumn.Hidden := TRUE;
          //                 SDABook.Sheets ['Imp_Fourn'].Columns ['I:I'].EntireColumn.Hidden := TRUE;
                           // $$$jp: 07/01/03 - nouvelle feuille d'impression (globale)
            //               SDABook.Sheets ['Imp_Global'].Columns ['H:H'].EntireColumn.Hidden := TRUE;
              //             SDABook.Sheets ['Imp_Global'].Columns ['I:I'].EntireColumn.Hidden := TRUE;
                      end;
}

                  SDABook.Sheets ['Saisie activité'].Columns ['J:J'].EntireColumn.Hidden := not VH_GC.CleAffaire.GestionAvenant;
                  // $$$JP 10/04/2003 - feuilles d'impression inutiles désormais
//                  SDABook.Sheets ['Imp_Prest'].Columns ['J:J'].EntireColumn.Hidden       := not VH_GC.CleAffaire.GestionAvenant;
  //                SDABook.Sheets ['Imp_Frais'].Columns ['J:J'].EntireColumn.Hidden       := not VH_GC.CleAffaire.GestionAvenant;
    //              SDABook.Sheets ['Imp_Fourn'].Columns ['J:J'].EntireColumn.Hidden       := not VH_GC.CleAffaire.GestionAvenant;
                  // $$$jp: 07/01/03 - nouvelle feuille d'impression (globale)
      //            SDABook.Sheets ['Imp_Global'].Columns ['J:J'].EntireColumn.Hidden       := not VH_GC.CleAffaire.GestionAvenant;}

                  // On cache colonne facturable si affichage F/NF est décoché (concept)
                  SDABook.Sheets ['Saisie activité'].Columns ['Q:Q'].EntireColumn.Hidden := not bAffFact;
                  // $$$JP 10/04/2003 - feuilles d'impression inutiles désormais
//                  SDABook.Sheets ['Imp_Frais'].Columns ['O:O'].EntireColumn.Hidden       := not bAffFact;
  //                SDABook.Sheets ['Imp_Fourn'].Columns ['O:O'].EntireColumn.Hidden       := not bAffFact;
                  // $$$jp: 07/01/03 - nouvelle feuille d'impression (globale)
    //              SDABook.Sheets ['Imp_Global'].Columns ['O:O'].EntireColumn.Hidden      := not bAffFact;}

                  // Enregistrement dans fichier de base (dans rép. dossier), qui sera copié pour tous les assistants
                  try
                     DeleteFile (pchar (sRepExp + '\SDA_GEN.xls'));
                     SDABook.SaveAs (sRepExp + '\SDA_GEN.xls');
                  except
                        PgiInfo ('Impossible de générer le classeur générique SDA_GEN.XLS dans ' + sRepExp, 'Saisie décentralisée d''activité');
                        SDAExcel.ActiveWorkBook.Close (FALSE);
                        SDAExcel.Quit;
                        SDAExcel := unAssigned;
                        FiniMove;
                        raise;
                  end;
               finally
                      SDAExcel.ActiveWorkBook.Close (FALSE);
                      SDAExcel.Quit;
                      SDAExcel := unAssigned;
                      FiniMove;
               end;

               // Copie du fichier pour tout les assistants demandés (ou un seul si un classeur pour tous)
//               bUnPourTous := GetParamSoc ('SO_AFEACTUNPOURTOUS');
               if bUnPourTous = FALSE then
               begin
                    InitMove (TobRessExp.Detail.Count, '');
                    CreeListe := TStringList.Create;
                    try
                       for i := 0 to TobRessExp.Detail.Count-1 do
                       begin
                            MoveCur(False);

                            // Nom du fichier à créer
                            NomFichier := sRepExp + '\SDA_' + StringReplace (DateToStr (PeriodeDateDeb), '/', '', [rfReplaceAll, rfIgnoreCase]) + '_' + StringReplace (DateToStr (PeriodeDateFin), '/', '', [rfReplaceAll, rfIgnoreCase]) + '_' + TobRessExp.Detail [i].GetValue ('ARS_RESSOURCE') + '§' + TobRessExp.Detail [i].GetValue ('ARS_LIBELLE') + '.xls';
                            if CopyFile (pchar (sRepExp + '\SDA_GEN.xls'), pchar (NomFichier), FALSE) = FALSE then
                            begin
                                 if PgiAsk ('Génération du classeur pour ' + TobRessExp.Detail [i].GetValue ('ARS_RESSOURCE') + ' - ' + TobRessExp.Detail [i].GetValue ('ARS_LIBELLE') + ' impossible' + #10 + ' Désirez-vous arrêter la génération ?', Ecran.Caption) = mrYes then
                                 begin
                                      bCanceled := TRUE;
                                      PgiInfo ('Génération des classeurs annulée', Ecran.Caption);
                                      break;
                                 end;
                            end
                            else
                                CreeListe.Add (TOBRessExp.Detail [i].GetValue ('ARS_RESSOURCE'));
                       end;
                    finally
                           FiniMove;
                    end;
               end
               else
               begin
                    // Nom du fichier à créer
                    // $$$JP 12/05/03: gestion nom fichier si pas en mode suivi SDA
                    bCopy := TRUE;
                    if bSuivi = TRUE then
                        NomFichier := sRepExp + '\SDA_' + StringReplace (DateToStr (PeriodeDateDeb), '/', '', [rfReplaceAll, rfIgnoreCase]) + '_' + StringReplace (DateToStr (PeriodeDateFin), '/', '', [rfReplaceAll, rfIgnoreCase]) // + '_SaisieActivité§SAct.xls'
                    else
                        NomFichier := sRepExp + '\SDA_01011900_31122099'; //_SaisieActivité§SAct.xls';
                    NomFichier := NomFichier +  '_SaisieActivité§SAct.xls';
                    if FileExists (NomFichier) = TRUE then
                       bCopy := PgiAsk ('Le classeur multi-assistant existe déjà. Désirez-vous le remplacer?', Ecran.Caption) = mrYes;
                    if bCopy = TRUE then
                    begin
                         if CopyFile (pchar (sRepExp + '\SDA_GEN.xls'), pchar (NomFichier), FALSE) = FALSE then
                         begin
                              PgiInfo ('Génération du classeur unique impossible', Ecran.Caption);
                              exit;
                         end;
                    end;
               end;

               // Mise à jour du suivi saisie décentralisé pour les assistants dont le classeur a été créé
               if bSuivi = TRUE then
                  if bUnPourTous = FALSE then
                  begin
                       try
                          for i := 0 to CreeListe.Count - 1 do //TobRessExp.Detail.Count-1 do
                              ExecuteSQL ('UPDATE SUIVIEACT SET ASA_ETATEACT = "ES1" WHERE ASA_RESSOURCE="' + CreeListe [i] + '" AND ASA_DATEDEB="' + UsDateTime (PeriodeDateDeb) + '" AND ASA_DATEFIN="' + UsDateTime (PeriodeDateFin) + '"')
                       finally
                              CreeListe.Free;
                       end;
                  end
                  else
                      for i := 0 to TOBRessExp.Detail.Count-1 do
                          ExecuteSQL ('UPDATE SUIVIEACT SET ASA_ETATEACT = "ES1" WHERE ASA_RESSOURCE="' + TOBRessExp.Detail [i].GetValue ('ARS_RESSOURCE') + '" AND ASA_DATEDEB="' + UsDateTime (PeriodeDateDeb) + '" AND ASA_DATEFIN="' + UsDateTime (PeriodeDateFin) + '"');
          end;
    finally
           Ecran.Enabled := TRUE;
           TOBTiersExp.Free;
           TOBAffExp.Free;
           TOBPrestExp.Free;
           TOBFraisExp.Free;
           TOBFourExp.Free;
           TOBRessExp.Free;
           TobUniteExp.Free;
   end;

    // Fin de l'export
    if bCanceled = FALSE then
        PgiInfoAf ('La génération des classeurs est terminée', Ecran.Caption)
    else
        PgiInfoAf ('Les classeurs non pas été générés', Ecran.Caption);
    if bSuivi = TRUE then
       F.Close;
end;

//$$$jp 04/2003: vire majtitrecolonnes (déjà en commentaire)


procedure AGLExportAffairesTiers(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
      F:=TForm(Longint(Parms[0])) ;
      if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
      if (TOTOF is TOF_AFFAIREEXPORT_MUL) then TOF_AFFAIREEXPORT_MUL(TOTOF).ExportAffairesTiers else exit;
end;

procedure AFLanceFiche_Mul_ExportAff (strArgument:string);
begin
     // $$$JP: pour gérer le passage de paramètre et compatible avec l'ancienne façon (sans param => 'ASSIST=*')
     if strArgument = '' then
         AGLLanceFiche ('AFF', 'AFFAIREEXPORT_MUL', '', '', 'ASSIST=*')
     else
         AGLLanceFiche ('AFF', 'AFFAIREEXPORT_MUL', '', '', strArgument);
end;




Initialization
  registerclasses ( [ TOF_AFFAIREEXPORT_MUL ] ) ;
  RegisterAglProc('ExportAffairesTiers',TRUE,0,AGLExportAffairesTiers);
end.
