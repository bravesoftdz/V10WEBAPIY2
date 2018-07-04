{***********UNITE*************************************************
Auteur  ...... : JP
Créé le ...... : 03/06/2003
Modifié le ... : 03/06/2003
Description .. : Gestion fiche lancement planning pour suivi activité
Mots clefs ... : TOF;ACTIVITE;PLANNING;SUIVI
*****************************************************************}
unit utofafactivite_suivi;

Interface

Uses
    StdCtrls, Controls, Classes, forms, sysutils,ComCtrls,Messages,Windows,
{$IFDEF EAGLCLIENT}
   MaineAGL,
{$ELSE}
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db,FE_Main,
{$ENDIF}

{$IFDEF BTP}
	 CalcOleGenericBTP,
{$ENDIF}

     HCtrls, HEnt1, HMsgBox, HPanel, UTOF , vierge, HTB97,
     UTOB,Affaireutil,AglInitGC,M3FP,Saisutil,Grids,
     Dicobtp,graphics,EntGC, Ent1,
     utilressource,AfUtilArticle,Utilarticle, UTofAfBaseCodeAffaire,
     ParamSoc, HPlanning;

Type
  TOF_AFACTIVITE_SUIVI = Class (TOF_AFBASECODEAFFAIRE)
  public
        procedure OnArgument (strArgument:String); override ;
        procedure OnClose;                         override ;
        procedure NomsChampsAffaire   (var Aff, Aff0,Aff1, Aff2, Aff3,Aff4, Aff_, Aff0_,Aff1_, Aff2_, Aff3_,Aff4_, Tiers, Tiers_:THEdit);  override;

  protected
           SuiviPlanning                           :THPlanning;  // L'objet planning
           PanelPlanning                           :THPanel;     // Le panneau d'affichage du planning + quelques conseils

           TOBRessPlanning                         :TOB;
           TOBItemPlanning                         :TOB;         // Les 3 TOBS utiles pour l'objet planning
           TOBEtatPlanning                         :TOB;

           strRegroup                              :string;      // L'ordre de regroupement (un caractère == une table)
           bRessource                              :boolean;
           bTiers                                  :boolean;
           bAffaire                                :boolean;
           bTypeArticle                            :boolean;
           bArticle                                :boolean;
           DateDebut, DateFin                      :TDateTime;   // Fourchette de date d'analyse
           strAgregat                              :string;      // Le champ à cumuler de la table ACTIVITE
           strFolio                                :string;
           bImmediat                               :boolean;     // Affichage immédiat du suivi d'activité

           procedure AfterShow          ();
           procedure OnClickValider     (Sender: TObject);
           procedure OnClickMaximiser   (Sender: TObject);
           procedure OnClickImprimer    (Sender: TObject);
           procedure OnClickExporter    (Sender: TObject);
           procedure OnDblClickPlanning (Sender: TObject);

           procedure OnKeyMaxWindow   (Sender: TObject; var Key: Word; Shift: TShiftState);

           function  GetSQLFiltre (He:THEdit; strPrefixe:string='')                :string;       overload;
           function  GetSQLFiltre (Hvcb:THValComboBox; strPrefixe:string='')       :string;       overload;
           function  GetSQLFiltre (Hmvcb:THMultiValComboBox; strPrefixe:string='') :string;       overload;
           function  GetSQLFiltre (cb:TCheckBox; strPrefixe:string='')             :string;       overload;

           function  GetRessourceFiltre  (strPrefixe:string)  :string;
           function  GetTiersFiltre      (strPrefixe:string)  :string;
           function  GetTiersComplFiltre (strPrefixe:string)  :string;
           function  GetAffaireFiltre    (strPrefixe:string)  :string;
           function  GetArticleFiltre    (strPrefixe:string)  :string;
           function  GetActiviteFiltre                        :string;

           procedure AffichePlanning (TOBPlanning:TOB);
           procedure MaximisePlanning;
  end;

procedure AFLanceFiche_Suivi_Activite (strArgument:string);


Implementation

uses dialogs, afactivite,utofafactivitemul, confidentaffaire;


procedure TOF_AFACTIVITE_SUIVI.AfterShow;
begin
     if IsInside (TFVierge (Ecran)) then
     begin
          THPanel (Ecran.Parent).CloseInside;
          THPanel (Ecran.Parent).VideToolBar;
     end;

     PostMessage (Ecran.Handle, WM_CLOSE, 0, 0);
end;

procedure TOF_AFACTIVITE_SUIVI.OnArgument (strArgument:string);
var
   strRess, strTypeArt    :string;
   strAffaire, strDates   :string;
   Critere                :string;
   TOBAff                 :TOB;
begin
     inherited;

     // Le panneau où s'affichera le planning
     PanelPlanning := THPanel (GetControl ('PPLANNING'));

     // Bouton Valider: lancement de la génération du planning
     TToolbarButton97 (GetControl ('BLANCESUIVI')).OnClick := OnClickValider;

     // Bouton Imprimer: lancement de l'impression du planning en cours
     TToolbarButton97 (GetControl ('BIMPRIMER')).Visible := TRUE;
     TToolbarButton97 (GetControl ('BIMPRIMER')).OnClick := OnClickImprimer;

     // Export vers classeur Excel
     TToolbarButton97 (GetControl ('BEXPORTEXCEL')).OnClick := OnClickExporter;

     // Bouton maximiser: pour afficher en plein écran le planning en cours
     TToolBarButton97 (GetControl ('BMAXPLANNING')).OnClick := OnClickMaximiser;

     // Restriction des types d'article sur seulement PRE;MAR;FOU
     THMultiValComboBox (GetControl ('GA_TYPEARTICLE')).Plus := PlusTypeArticle (TRUE);

     // Par défaut (avant prise en compte critères de lancement)
     TRadioButton (GetControl ('RBRESS')).Checked       := TRUE;
     TCheckBox (GetControl ('CBRESSOURCE')).Checked     := FALSE;
     TCheckBox (GetControl ('CBTOUTESLESRESS')).Checked := TRUE;

     // Si pas manager, alors restrictions
     if SaisieActiviteManager = FALSE then
     begin
          if VH_GC.RessourceUser <> '' then
          begin
               THEdit (GetControl ('ARS_RESSOURCE')).Text    := VH_GC.RessourceUser;
               THEdit (GetControl ('ARS_RESSOURCE')).Enabled := FALSE;
          end
          else
          begin
               PgiInfo ('Vous n''avez pas accès à cette commande.Le User courant n''a pas de droits ou pas de code associé');
//               TToolBarButton97 (GetControl ('BLANCESUIVI')).Enabled := FALSE;
               TFVierge (Ecran).OnAfterFormShow := AfterShow;
               Ecran.Close;
          end;
     end;

     // Récupère les critères de lancement: ressource, type article, dates
     bImmediat := FALSE;
     strFolio  := '';
     Critere := Trim (ReadTokenSt (strArgument));
     while Critere <>'' do
     begin
          if Copy (Critere, 1, 9) = 'RESSOURCE' then
          begin
               strRess := Copy (Critere, 11, Length (Critere)-10);
               THEdit (GetControl ('ARS_RESSOURCE')).Text := strRess;
               if strRess <> '' then
                  THEdit (GetControl ('ARS_RESSOURCE')).Enabled      := FALSE;
          end
          else if Copy (Critere, 1, 7) = 'AFFAIRE' then
          begin
               strAffaire := Copy (Critere, 9, Length (Critere)-8);
               THEdit (GetControl ('AFF_AFFAIRE')).Text := strAffaire;
               if strAffaire <> '' then
               begin
                    // Affichage code affaire et code tiers
                    TOBAff := TOB.Create ('une affaire', nil, -1);
                    try
                       TOBAff.LoadDetailFromSQL('SELECT AFF_AFFAIRE1, AFF_AFFAIRE2, AFF_AFFAIRE3, AFF_AVENANT, AFF_TIERS FROM AFFAIRE WHERE AFF_AFFAIRE="' + strAffaire + '"');
                       if TOBAff.Detail.Count = 1 then
                       begin
                            THEdit (GetControl ('AFF_AFFAIRE1')).Text := TOBAff.Detail [0].GetValue ('AFF_AFFAIRE1');
                            THEdit (GetControl ('AFF_AFFAIRE2')).Text := TOBAff.Detail [0].GetValue ('AFF_AFFAIRE2');
                            THEdit (GetControl ('AFF_AFFAIRE3')).Text := TOBAff.Detail [0].GetValue ('AFF_AFFAIRE3');
                            THEdit (GetControl ('AFF_AVENANT')).Text  := TOBAff.Detail [0].GetValue ('AFF_AVENANT');
                            THEdit (GetControl ('T_TIERS')).Text      := TOBAff.Detail [0].GetValue ('AFF_TIERS');
                       end;
                    finally
                           TOBAff.Free;
                    end;

                    // On interdit la sélection des affaires et des tiers
                    THEdit (GetControl ('AFF_AFFAIRE1')).Enabled           := FALSE;
                    THEdit (GetControl ('AFF_AFFAIRE2')).Enabled           := FALSE;
                    THEdit (GetControl ('AFF_AFFAIRE3')).Enabled           := FALSE;
                    THEdit (GetControl ('AFF_AVENANT')).Enabled            := FALSE;
                    TToolBarButton97 (GetControl ('BSELECTAFF1')).Enabled  := FALSE;
                    TToolBarButton97 (GetControl ('BEFFACEAFF1')).Enabled  := FALSE;
                    THEdit (GetControl ('T_TIERS')).Enabled                := FALSE;
               end;
               TRadioButton (GetControl ('RBAFFAIRE')).Checked := TRUE;
               TCheckBox (GetControl ('CBARTICLE')).Checked    := TRUE;
          end
          else if Copy (Critere, 1, 11) = 'TYPEARTICLE' then
          begin
               strTypeArt := Copy (Critere, 13, Length (Critere)-12);
               THMultiValComboBox (GetControl ('GA_TYPEARTICLE')).Value := strTypeArt + ';';
               if strTypeArt = 'PRE' then
                  TRadioButton (GetControl ('RBQTE')).Checked := TRUE;
          end
          else if Copy (Critere, 1, 12) = 'DATEACTIVITE' then
          begin
               strDates := Copy (Critere, 14, Length (Critere)-13);
               THEdit (GetControl ('ACT_DATEACTIVITE')).Text := Copy (strDates, 1, 10);
               THEdit (GetControl ('ACT_DATEACTIVITE_')).Text := Copy (strDates, 11, 10);
          end
          else if Copy (Critere, 1, 5) = 'FOLIO' then
          begin
               strFolio := Copy (Critere, 7, Length (Critere)-6);
          end
          else if Copy (Critere, 1, 8) = 'IMMEDIAT' then
          begin
               bImmediat := TRUE;
          end;

          // Critères de lancement suivant
          Critere := Trim (ReadTokenSt (strArgument));
     end;

     // Si visualisation valorisation interdite, on inhibe les agrégats en PR ou PV
     if AffichageValorisation = FALSE then
     begin
          TRadioButton (GetControl ('RBPR')).Enabled  := FALSE;
          TRadioButton (GetControl ('RBPV')).Enabled  := FALSE;
          TRadioButton (GetControl ('RBQTE')).Checked := TRUE;
     end;

     // Le planning en cours n'existe pas encore
     TOBRessPlanning := nil;
     TOBEtatPlanning := nil;
     TOBItemPlanning := nil;
     SuiviPlanning   := nil;

   If Not GetParamSoc ('So_AfAppPoint') then
     begin
     SetControlVisible ('ACT_ETATVISAFAC',False);
     SetControlVisible ('TACT_ETATVISAFAC',False);
     end;

     // Si lancement instantané demandé, on le fait
     if bImmediat = TRUE then
        OnClickValider (nil);
end;

procedure TOF_AFACTIVITE_SUIVI.OnClose ;
begin
     inherited ;

     TOBRessPlanning.Free;
     TOBItemPlanning.Free;
     TOBEtatPlanning.Free;
     SuiviPlanning.Free;
end;

function TOF_AFACTIVITE_SUIVI.GetSQLFiltre (He:THEdit; strPrefixe:string):string;
var
   strChamp    :string;
   strType     :string;
   strValue    :string;
begin
     // Par défaut, filtre SQL inconnu
     Result := '';
     if He = nil then
        exit;

     // Si une valeur existe dans l'edit, on créer le SQL équivalent
     strValue := He.Text;
     if strValue <> '' then
     begin
          // Nom et type du champ associé à l'edit
          strChamp := He.Name;
          if Copy (strChamp, Length (strChamp), 1) = '_' then
             strChamp := Copy (strChamp, 1, Length (strChamp)-1);
          strType  := ChampToType (strChamp);

          // SQL de la forme: CHAMP="VALEUR" (avec ou sans guillemet, selon le type)
          Result := ' AND ' + strPrefixe + strChamp;
          case He.Operateur of
               Superieur:
                         Result := Result + '>=';
               Inferieur:
                         Result := Result + '<=';
               Egal:
                         Result := Result + '=';
          end;
          if (Copy (strType, 1, 7) = 'VARCHAR') OR (strType = 'CHAR') OR (strType = 'COMBO') then
             if HE.Operateur = Commence then
                 Result := Result + ' LIKE "' + strValue + '%"'
             else
                 Result := Result + '"' + strValue + '"'
          else
              if (strType = 'DATE') then
                  Result := Result + '"' + USDATETIME (strtodate(strValue)) + '"'
              else
                  if (strType = 'INTEGER') OR (strType = 'DOUBLE') then
                      Result := Result + strValue
                  else
                      Result := '';
     end;
end;

function TOF_AFACTIVITE_SUIVI.GetSQLFiltre (Hvcb:THValComboBox; strPrefixe:string):string;
var
   strChamp    :string;
   strType     :string;
   strValue    :string;
begin
     // Par défaut, filtre SQL inconnu
     Result := '';
     if Hvcb = nil then
        exit;

     // Si une valeur existe dans le ValComboBox, on créer le SQL équivalent
     strValue := Hvcb.Value;
     if strValue <> '' then
     begin
          // Nom et type du champ associé au ValComboBox
          strChamp := Hvcb.Name;
          strType  := ChampToType (strChamp);

          // SQL de la forme: CHAMP="VALEUR" (avec ou sans guillemet, selon le type)
          Result := ' AND ' + strPrefixe + strChamp + '=';
          if (Copy (strType, 1, 7) = 'VARCHAR') OR (strType = 'CHAR') OR (strType = 'COMBO') OR (strType = 'DATE') then
              Result := Result + '"' + strValue + '"'
          else
              if (strType = 'INTEGER') OR (strType = 'DOUBLE') then
                  Result := Result + '"' + strValue
              else
                  Result := '';
     end;
end;

function TOF_AFACTIVITE_SUIVI.GetSQLFiltre (Hmvcb:THMultiValComboBox; strPrefixe:string):string;
var
   strChamp    :string;
   strType     :string;
   strValues   :string;
   strValue    :string;
   bFirstValue :boolean;
begin
     // Par défaut, filtre SQL inconnu
     Result := '';
     if hmvcb = nil then
        exit;

     // Si une valeur existe dans le MultiValComboBox, on créer le SQL équivalent
     strValues := Trim (Hmvcb.Value);
     if (strValues <> '') AND (UpperCase (strValues) <> '<<TOUS>>') then
     begin
          // Nom et type du champ associé au MultiValComboBox
          strChamp := Hmvcb.Name;
          strType  := ChampToType (strChamp);

          // SQL de la forme: CHAMP IN ("VALEUR1","VALEUR2...)
          bFirstValue := TRUE;
          while strValues <> '' do
          begin
               // Identification d'une valeur parmi la chaine du multival
               strValue := ReadTokenSt (strValues);
               if strValue <> '' then
               begin
                    // Si première valeur trouvée, il faut construire le début de la requête
                    if bFirstValue = TRUE then
                        Result := ' AND ' + strPrefixe + strChamp + ' IN ('
                    else
                        Result := Result + ',';
                    Result := Result + '"' + strValue + '"';
                    bFirstValue := FALSE;
               end;
          end;

          // Si au moins une valeur identifiée, il faut fermer le "IN"
          if bFirstValue = FALSE then
             Result := Result + ')';
     end;
end;

function TOF_AFACTIVITE_SUIVI.GetSQLFiltre (cb:TCheckBox; strPrefixe:string):string;
var
   strChamp    :string;
   cbValue     :TCheckBoxState;
begin
     // Par défaut, filtre SQL inconnu
     Result := '';
     if cb = nil then
        exit;

     // Si une valeur existe dans le MultiValComboBox, on créer le SQL équivalent
     cbValue := cb.State;
     if cbValue <> cbGrayed then
     begin
          // Nom et type du champ associé au MultiValComboBox
          strChamp := cb.Name;

          // SQL de la forme: CHAMP IN ("VALEUR1","VALEUR2...)
          Result   := ' AND ' + strPrefixe + strChamp + '=';
          if cbValue = cbChecked then
              Result := Result + '"X"'
          else
              Result := Result + '"-"';
     end;
end;

function TOF_AFACTIVITE_SUIVI.GetRessourceFiltre (strPrefixe:string):string;
var
   i   :integer;
begin
     Result := GetSQLFiltre (THEdit (GetControl ('ARS_RESSOURCE')), strPrefixe);
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('ARS_TYPERESSOURCE')), strPrefixe);
     Result := Result + GetSQLFiltre (THEdit (GetControl ('ARS_FONCTION1')), strPrefixe);
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('ARS_ETABLISSEMENT')), strPrefixe);
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('ARS_DEPARTEMENT')), strPrefixe);
     Result := Result + GetSQLFiltre (TCheckBox (GetControl ('ARS_FERME')), strPrefixe);
     for i := 1 to 9 do
         Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('ARS_LIBRERES'+IntToStr(i))), strPrefixe);
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('ARS_LIBRERESA')), strPrefixe);
end;

function TOF_AFACTIVITE_SUIVI.GetTiersFiltre (strPrefixe:string):string;
begin
     Result := GetSQLFiltre (THEdit (GetControl ('T_TIERS')), strPrefixe);
     Result := Result + GetSQLFiltre (THEdit (GetControl ('T_SOCIETEGROUPE')), strPrefixe);
end;

function TOF_AFACTIVITE_SUIVI.GetTiersComplFiltre (strPrefixe:string):string;
var
   i   :integer;
begin
     Result := '';
     for i := 1 to 9 do
         Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('YTC_TABLELIBRETIERS'+IntToStr(i))), strPrefixe);
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('YTC_TABLELIBRETIERSA')), strPrefixe);
     for i := 1 to 3 do
         Result := Result + GetSQLFiltre (THEdit (GetControl ('YTC_RESSOURCE'+IntToStr (i))), strPrefixe);
end;

function TOF_AFACTIVITE_SUIVI.GetAffaireFiltre (strPrefixe:string):string;
var
   i   :integer;
begin
     Result := GetSQLFiltre (THEdit (GetControl ('AFF_AFFAIRE')), strPrefixe);
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('AFF_ETABLISSEMENT')), strPrefixe);
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('AFF_DEPARTEMENT')), strPrefixe);
     Result := Result + GetSQLFiltre (THEdit (GetControl ('AFF_RESPONSABLE')), strPrefixe);
     Result := Result + GetSQLFiltre (TCheckBox (GetControl ('AFF_ADMINISTRATIF')), strPrefixe);
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('AFF_GROUPECONF')), strPrefixe);
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('AFF_ETATAFFAIRE')), strPrefixe);
     for i := 1 to 9 do
         Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('AFF_LIBREAFF'+IntToStr (i))), strPrefixe);
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('AFF_LIBREAFFA')), strPrefixe);
     for i := 1 to 3 do
         Result := Result + GetSQLFiltre (THEdit (GetControl ('AFF_RESSOURCE'+IntToStr (i))), strPrefixe);
end;

function TOF_AFACTIVITE_SUIVI.GetArticleFiltre (strPrefixe:string):string;
var
   i   :integer;
begin
     Result := GetSQLFiltre (THEdit (GetControl ('GA_CODEARTICLE')));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('GA_TYPEARTICLE')), strPrefixe);
     for i := 1 to 9 do
         Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('GA_LIBREART'+IntToStr (i))), strPrefixe);
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('GA_LIBREARTA')), strPrefixe);
     for i := 1 to 3 do
         Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('GA_FAMILLENIV'+IntToStr (i))), strPrefixe);
end;

function TOF_AFACTIVITE_SUIVI.GetActiviteFiltre:string;
begin
     Result := GetSQLFiltre (THEdit (GetControl ('ACT_DATEACTIVITE')));
     Result := Result + GetSQLFiltre (THEdit (GetControl ('ACT_DATEACTIVITE_')));
//     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('ACT_ACTORIGINE')));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('ACT_ACTIVITEREPRIS')));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('ACT_ETATVISA')));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('ACT_ETATVISAFAC')));
     Result := Result + GetSQLFiltre (TCheckBox (GetControl ('ACT_ACTIVITEEFFECT')));
     if strFolio <> '' then
        Result := Result + ' AND ACT_FOLIO="' + strFolio + '"';
     Result := Result + ' AND ACT_AFFAIRE0="A"';
end;

procedure TOF_AFACTIVITE_SUIVI.AffichePlanning (TOBPlanning:TOB);
var
   TOBFille            :TOB;
   TOBLigneRessource   :TOB;
   strColEntetePlanning:string;
   strColNamePlanning  :string;
   strColSizePlanning  :string;
   strColAlignPlanning :string;
   strClePlanning      :string;
   strCurClePlanning   :string;
   strCleRess, strCleTiers, strCleAff, strCleTypeArt, strCleArt                 :string;
   strCurCleRess, strCurCleTiers, strCurCleAff, strCurCleTypeArt, strCurCleArt  :string;
   strCleAff1, strCleAff2, strCleAff3, strCleAffAv                              :string;
   bRupture            :boolean;
   dQte, dCumul        :double;
   i, j                :integer;
   iNumChamp           :integer;
   DateAct             :TDateTime;
begin
     // Création et affichage de l'objet THPlanning
     SuiviPlanning := THPlanning.Create (Ecran);
     if SuiviPlanning = nil then
         PgiInfo ('Initialisation de la fenêtre planning impossible')
     else
     begin
          SuiviPlanning.Parent     := PanelPlanning; //THPanel (GetControl ('PPLANNING'));
          SuiviPlanning.Activate   := FALSE;
          SuiviPlanning.Align      := alClient;
          SuiviPlanning.Title      := 'Suivi d''activité';
          SuiviPlanning.OnDblClick := OnDblClickPlanning;

          // Configuration des états (les etats possible pour les items)
          SuiviPlanning.EtatChampBackGroundColor := 'E_COULEURFOND';
          SuiviPlanning.EtatChampCode            := 'E_CODE';
          SuiviPlanning.EtatChampFontColor       := 'E_COULEURFONTE';
          SuiviPlanning.EtatChampFontName        := 'E_NOMFONTE';
          SuiviPlanning.EtatChampFontSize        := 'E_TAILLEFONTE';
          SuiviPlanning.EtatChampFontStyle       := 'E_STYLEFONTE';
          SuiviPlanning.EtatChampIcone           := 'E_ICONE';
          SuiviPlanning.EtatChampLibelle         := 'E_LIBELLE';

          // Configuration des items (les éléments du corps du planning)
          SuiviPlanning.ChampLineID     := 'I_CODE';
          SuiviPlanning.ChampLibelle    := 'I_LIBELLE';
          SuiviPlanning.ChampdateDebut  := 'I_DATEDEBUT';
          SuiviPlanning.ChampDateFin    := 'I_DATEFIN';
          SuiviPlanning.ChampEtat       := 'I_ETAT';
          SuiviPlanning.ChampColor      := 'I_COLOR';
          SuiviPlanning.ChampHint       := 'I_HINT';
          SuiviPlanning.ChampIcone      := 'I_ICONE';

          // Configuration des ressources (les entêtes de ligne)
          SuiviPlanning.ResChampColor       := '';
          SuiviPlanning.ResChampFixedColor  := '';
          SuiviPlanning.ResChampID          := 'R_CODE';
          SuiviPlanning.ResChampReadOnly    := '';

          // Couleurs WE/jours fériés et case en sélection
          SuiviPlanning.GestionJoursFeriesActive := TRUE;
          SuiviPlanning.ActiveSaturday           := TRUE;
          SuiviPlanning.ActiveSunday             := TRUE;
          SuiviPlanning.ColorJoursFeries         := clOlive;
          SuiviPlanning.ColorOfSaturday          := clOlive;
          SuiviPlanning.ColorOfSunday            := clOlive;
          SuiviPlanning.ColorSelection           := clWhite;

          // Taille des lignes/colonnes
          SuiviPlanning.RowSizeData := 16;
          if strAgregat <> 'ACT_QTEUNITEREF' then
              SuiviPlanning.ColSizeData := 45
          else
              SuiviPlanning.ColSizeData := 35;

          // Création de la liste des états possible pour les items (pour l'instant un seul)
          TOBEtatPlanning.Free;
          TOBEtatPlanning := TOB.Create ('les etats', nil, -1);
          with TOB.Create('un etat', TOBEtatPlanning, -1) do
          begin
               AddChampSupValeur ('E_CODE', '001');
               AddChampSupValeur ('E_LIBELLE', '');
               AddChampSupValeur ('E_COULEURFOND', ColorToString (RGB(120,205,205)));
               AddChampSupValeur ('E_COULEURFONTE', 'clBlack');
               AddChampSupValeur ('E_NOMFONTE', 'Arial');
               AddChampSupValeur ('E_TAILLEFONTE', 8);
               AddChampSupValeur ('E_STYLEFONTE', '');
               AddChampSupValeur ('E_ICONE', -1);
          end;
          with TOB.Create('un etat', TOBEtatPlanning, -1) do
          begin
               AddChampSupValeur ('E_CODE', '002');
               AddChampSupValeur ('E_LIBELLE', '');
               AddChampSupValeur ('E_COULEURFOND', ColorToString (RGB(255,255,255)));
               AddChampSupValeur ('E_COULEURFONTE', 'clBlack');
               AddChampSupValeur ('E_NOMFONTE', 'Arial');
               AddChampSupValeur ('E_TAILLEFONTE', 8);
               AddChampSupValeur ('E_STYLEFONTE', '');
               AddChampSupValeur ('E_ICONE', -1);
          end;

          // Taille, alignement des ressources planning (=entête de ligne du planning): en fonction de l'ordre défini
          for i := 1 to Length (strRegroup) do
          begin
               case strRegroup [i] of
                    'R':
                        begin
                             strColEntetePlanning := strColEntetePlanning + TraduitGA(';Ressource');
                             strColNamePlanning   := strColNamePlanning + ';R_LIBRESSOURCE';
                             strColSizePlanning   := strColSizePlanning + ';140';
                             strColAlignPlanning  := strColAlignPlanning + ';L';
                        end;

                    'C':
                        begin
                             strColEntetePlanning := strColEntetePlanning + TraduitGA(';Tiers');
                             strColNamePlanning  := strColNamePlanning + ';R_LIBTIERS';
                             strColSizePlanning  := strColSizePlanning + ';140';
                             strColAlignPlanning := strColAlignPlanning + ';L';
                        end;

                    'M':
                        begin
                             strColEntetePlanning := strColEntetePlanning + TraduitGA(';Affaire');
                             strColNamePlanning  := strColNamePlanning + ';R_LIBAFFAIRE';
                             if bTiers = FALSE then
                                 strColSizePlanning  := strColSizePlanning + ';160'
                             else
                                 strColSizePlanning  := strColSizePlanning + ';100';
                             strColAlignPlanning := strColAlignPlanning + ';L';
                        end;

                    'T':
                        begin
                             strColEntetePlanning := strColEntetePlanning + TraduitGA(';Type article');
                             strColNamePlanning  := strColNamePlanning + ';R_LIBTYPEARTICLE';
                             strColSizePlanning  := strColSizePlanning + ';70';
                             strColAlignPlanning := strColAlignPlanning + ';C';
                        end;

                    'A':
                        begin
                             strColEntetePlanning := strColEntetePlanning + TraduitGA(';Article');
                             strColNamePlanning  := strColNamePlanning + ';R_LIBARTICLE';
                             strColSizePlanning  := strColSizePlanning + ';150';
                             strColAlignPlanning := strColAlignPlanning + ';L';
                        end;
               end;
          end;
          SuiviPlanning.TokenFieldColEntete := Copy (strColEntetePlanning, 2, Length (strColEntetePlanning)-1) + ';Cumul';
          SuiviPlanning.TokenFieldColFixed  := Copy (strColNamePlanning, 2, Length (strColNamePlanning)-1) + ';R_LIBCUMUL';
          SuiviPlanning.TokenSizeColFixed   := Copy (strColSizePlanning, 2, Length (strColSizePlanning)-1) + ';50';
          SuiviPlanning.TokenAlignColFixed  := Copy (strColAlignPlanning, 2, Length (strColAlignPlanning)-1) + ';C';

          // Date périmètres du planning
          SuiviPlanning.IntervalDebut      := DateDebut;
          SuiviPlanning.IntervalFin        := DateFin;
          SuiviPlanning.DateOfStart        := DateDebut;

          // Forme des items (rectangle)
          SuiviPlanning.FormeGraphique     := pgRectangle;
          SuiviPlanning.Autorisation       := [];
          SuiviPlanning.MouseAlready       := TRUE;

          // Entête de colonne
          SuiviPlanning.ActiveLigneGroupeDate := TRUE;
          SuiviPlanning.ActiveLigneDate       := TRUE;
          SuiviPlanning.Interval              := piJour;
          if TRadioButton (GetControl ('RBJOURSEMAINE')).Checked = TRUE then
          begin
               SuiviPlanning.CumulInterval := pciSemaine;
               SuiviPlanning.DateFormat    := 'dd/mm';
          end
          else
          begin
               SuiviPlanning.CumulInterval := pciMois;
               SuiviPlanning.DateFormat    := 'dd';
          end;

          // Génération de la TOB ressources-planning (pas ressources au sens métier)
          strCurClePlanning := '';
          strCurCleRess     := '';
          strCurCleTiers    := '';
          strCurCleAff      := '';
          strCurCleTypeArt  := '';
          strCurCleArt      := '';
          dCumul            := 0.00;
          TOBLigneRessource := nil;

          // Création TOB pour le planning
          TOBRessPlanning.Free;
          TOBRessPlanning := TOB.Create ('les ressources', nil, -1);
          TOBItemPlanning.Free;
          TOBItemPlanning := TOB.Create ('les items', nil, -1);
          for i := 0 to TOBPlanning.Detail.Count-1 do
          begin
               // Agrégat de ligne d'activité à traiter
               TOBFille := TOBPlanning.Detail [i];

               // Valeur "clé" de l'enregistrement: composée des codes de tri, et vérification si rupture
               strClePlanning := '';
               bRupture       := FALSE;
               for j := 1 to Length (strRegroup) do
               begin
                    case strRegroup [j] of
                         'R':
                             begin
                                  strCleRess     := TOBFille.GetValue ('ARS_RESSOURCE');
                                  strClePlanning := strClePlanning + '§' + strCleRess;
                                  if bRupture = TRUE then
                                      strCurCleRess := ''
                                  else
                                      if strCleRess <> strCurCleRess then
                                         bRupture := TRUE;
                             end;

                         'C':
                             begin
                                  strCleTiers    := TOBFille.GetValue ('T_TIERS');
                                  strClePlanning := strClePlanning + '§' + strCleTiers;
                                  if bRupture = TRUE then
                                      strCurCleTiers := ''
                                  else
                                      if strCleTiers <> strCurCleTiers then
                                         bRupture := TRUE;
                             end;

                         'M':
                             begin
                                  // Si mission seule, il faut quand même le code tiers dans la rupture
                                  if bTiers = FALSE then
                                  begin
                                       strCleTiers    := TOBFille.GetValue ('T_TIERS');
                                       strCleAff      := strCleTiers;
                                  end;

                                  // Clé affaire: seulement sur les parties utilisables (en fonction paramètres société)
                                  strCleAff1 := TOBFille.GetValue ('AFF_AFFAIRE1');
                                  iNumChamp := TOBFille.GetNumChamp ('AFF_AFFAIRE2');
                                  if iNumChamp > 0 then
                                     strCleAff2 := TOBFille.GetValeur (iNumChamp);
                                  iNumChamp := TOBFille.GetNumChamp ('AFF_AFFAIRE3');
                                  if iNumChamp > 0 then
                                     strCleAff3 := TOBFille.GetValeur (iNumChamp);
                                  iNumChamp := TOBFille.GetNumChamp ('AFF_AVENANT');
                                  if iNumChamp > 0 then
                                     strCleAffAv := TOBFille.GetValeur (iNumChamp);
                                  strCleAff := strCleAff + strCleAff1 + strCleAff2 + strCleAff3 + strCleAffAv;
                                  strClePlanning := strClePlanning + '§' + strCleAff;

                                  if bRupture = TRUE then
                                      strCurCleAff   := ''
                                  else
                                      if strCleAff <> strCurCleAff then
                                         bRupture := TRUE;
                             end;

                         'T':
                             begin
                                  strCleTypeArt  := TOBFille.GetValue ('GA_TYPEARTICLE');
                                  strClePlanning := strClePlanning + '§' + strCleTypeArt;
                                  if bRupture = TRUE then
                                      strCurCleTypeArt := ''
                                  else
                                      if strCleTypeArt <> strCurCleTypeArt then
                                         bRupture := TRUE;
                             end;

                         'A':
                             begin
                                  strCleArt      := TOBFille.GetValue ('GA_CODEARTICLE');
                                  strClePlanning := strClePlanning + '§' + strCleArt;
                                  if bRupture = TRUE then
                                      strCurCleArt := ''
                                  else
                                      if strCleArt <> strCurCleArt then
                                         bRupture := TRUE;
                             end;
                    end;
               end;
               if bRupture = TRUE then
               begin
                    // Mise à jour des cumuls sur la ligne terminée
                    if TOBLigneRessource <> nil then
                       TOBLigneRessource.PutValue ('R_LIBCUMUL', FormatFloat ('0,.##', dCumul) + ' ');

                    // Nouvelle clé courante et cumul à 0 car nouvelle ligne
                    strCurClePlanning := strClePlanning;
                    dCumul            := 0.00;

                    // Création d'une nouvelle ressource planning avec comme code unique le n° de l'enregistrement
                    TOBLigneRessource := TOB.Create ('une ressource', TOBRessPlanning, -1);
                    with TOBLigneRessource do
                    begin
                         AddChampSupValeur ('R_CODE', strCurClePlanning);
                         if bRessource = TRUE then
                            if strCleRess <> strCurCleRess then
                                AddChampSupValeur ('R_LIBRESSOURCE', Trim (strCleRess) + ' - ' + Trim (string (TOBFille.GetValue('LIBRESSOURCE'))))
                            else
                                AddChampSupValeur ('R_LIBRESSOURCE', '');
                         if bTiers = TRUE then
                            if strCleTiers <> strCurCleTiers then
                                AddChampSupValeur ('R_LIBTIERS', Trim (strCleTiers) + ' - ' + Trim (string (TOBFille.GetValue ('LIBTIERS'))))
                            else
                                AddChampSupValeur ('R_LIBTIERS', '');
                         if bAffaire then
                            if strCleAff <> strCurCleAff then
                               if bTiers = FALSE then
                                   AddChampSupValeur ('R_LIBAFFAIRE', Trim (strCleTiers) + ' - ' + Trim (strCleAff1) + ' ' + Trim (strCleAff2) + ' ' + Trim (strCleAff3) + ' ' + Trim (strCleAffAv))
                               else
                                   AddChampSupValeur ('R_LIBAFFAIRE', Trim (strCleAff1) + ' ' + Trim (strCleAff2) + ' ' + Trim (strCleAff3) + ' ' + Trim (strCleAffAv))
                            else
                                AddChampSupValeur ('R_LIBAFFAIRE', '');
                         if bTypeArticle = TRUE then
                            if strCleTypeArt <> strCurCleTypeArt then
                                AddChampSupValeur ('R_LIBTYPEARTICLE', Trim (strCleTypeArt))
                            else
                                AddChampSupValeur ('R_LIBTYPEARTICLE', '');
                         if bArticle = TRUE then
                            if strCleArt <> strCurCleArt then
                                AddChampSupValeur ('R_LIBARTICLE', Trim (strCleArt) + ' - ' + Trim (string (TOBFille.GetValue ('LIBARTICLE'))))
                            else
                                AddChampSupValeur ('R_LIBARTICLE', '');
                         AddChampSupValeur ('R_LIBCUMUL', '-');
                    end;

                    // Nouvelle sous-clé
                    strCurCleRess      := strCleRess;
                    strCurCleTiers     := strCleTiers;
                    strCurCleAff       := strCleAff;
                    strCurCleTypeArt   := strCleTypeArt;
                    strCurCleArt       := strCleArt;
               end;

               // Création de l'item associé (si date existante)
               DateAct := TOBFille.GetValue ('ACT_DATEACTIVITE');
               with TOB.Create ('un item', TOBItemPlanning, -1) do
               begin
                    AddChampSupValeur ('I_CODE', strCurClePlanning);
                    if DateAct > 0 then
                    begin
                         dQte   := TOBFille.GetValue ('LETOTAL');
                         dCumul := dCumul + dQte;
                         AddChampSupValeur ('I_LIBELLE', FormatFloat ('0,.##', dQte));
                         AddChampSupValeur ('I_DATEDEBUT', DateAct);
                         AddChampSupValeur ('I_DATEFIN', DateAct);
                         if dQte > 0.00 then
                             AddChampSupValeur ('I_ETAT', '001')
                         else
                             AddChampSupValeur ('I_ETAT', '002');
                         AddChampSupValeur ('I_DATA', dQte);
                         AddChampSupValeur ('I_CODERESS', strCleRess);
                         AddChampSupValeur ('I_CODETIERS', strCleTiers);
                         AddChampSupValeur ('I_CODEAFF1', strCleAff1);
                         if strCleAff2 <> '' then
                            AddChampSupValeur ('I_CODEAFF2', strCleAff2);
                         if strCleAff3 <> '' then
                            AddChampSupValeur ('I_CODEAFF3', strCleAff3);
                         if strCleAffAv <> '' then
                            AddChampSupValeur ('I_CODEAFFAV', strCleAffAv);
                         AddChampSupValeur ('I_TYPEARTICLE', strCleTypeArt);
                         AddChampSupValeur ('I_CODEARTICLE', strCleArt);
                         AddChampSupValeur ('I_HINT', IntToStr (TOBFille.GetValue ('LENOMBRE')) + ' ligne(s) d''activité');
                    end
                    else
                    begin
                         AddChampSupValeur ('I_LIBELLE', '');
                         AddChampSupValeur ('I_DATEDEBUT', DateDebut);
                         AddChampSupValeur ('I_DATEFIN', DateFin);
                         AddChampSupValeur ('I_ETAT', '002');
                         AddChampSupValeur ('I_DATA', 0.0);
                         AddChampSupValeur ('I_CODERESS', strCleRess);
                         AddChampSupValeur ('I_HINT', 'Aucune ligne d''activité répondant aux critères indiqués');
                    end;
                    AddChampSupValeur ('I_COLOR', '');
                    AddChampSupValeur ('I_ICONE', -1);
               end;
          end;

          // Mise à jour cumuls dernière ligne
          if TOBLigneRessource <> nil then
             TOBLigneRessource.PutValue ('R_LIBCUMUL', FormatFloat ('0,.##', dCumul) + ' ');

          // Si aucune données, on affiche pas le planning (seulement un message d'avertissement)
          if TOBItemPlanning.Detail.Count > 0 then
          begin
               // Affichage planning
               SuiviPlanning.TobRes   := TOBRessPlanning;
               SuiviPlanning.TobEtats := TOBEtatPlanning;
               SuiviPlanning.TobItems := TOBItemPlanning;
               SuiviPlanning.Activate := TRUE;

               // Dans une fenêtre maximisée de manière automatique
               // $$$JP 28/08/2003: ne plus maximiser automatiquement (suite remarque AF et CR)
               //MaximisePlanning;
               // $$$JP FIN
          end
          else
          begin
               SuiviPlanning.Visible := FALSE;
               PanelPlanning.Caption := 'Aucune ligne d''activité ne correspond aux critères spécifiés';
               PanelPlanning.Font.Style := [fsBold];
               PanelPlanning.Refresh;
               SuiviPlanning.Free;
               SuiviPlanning := nil;
          end;
     end;
end;

procedure TOF_AFACTIVITE_SUIVI.OnClickValider (Sender:TObject);
var
   i                   :integer;
   bShowAllRessource   :boolean;
   strSelectActivite   :string;
   strJoinActivite     :string;
   strGroupActivite    :string;
   strOrderActivite    :string;
   strWhereActivite    :string;
   strWhereRess        :string;
   strWhereTiers       :string;
   strWhereTiersCompl  :string;
   strWhereAffaire     :string;
   strWhereArticle     :string;
   strRequete          :string;
   strLibRessource     :string;
   TOBActivite         :TOB;
   TOBRessource        :TOB;
   TOBPlanning         :TOB;
begin
     // Dernier planning construit doit être supprimé avant nouveaux traitements
     SuiviPlanning.Free;
     SuiviPlanning := nil;

     // On revient sur l'écran de sélection présentation, et on affiche un message d'attente
     TPageControl (GetControl ('PAGES')).ActivePageIndex := 0;
     PanelPlanning.Caption := 'Veuillez patienter pendant la génération du suivi d''activité...';
     PanelPlanning.Font.Style := [];
     PanelPlanning.Refresh;

     // Initialisations
     strRegroup        := '';
     strSelectActivite := '';
     strJoinActivite   := '';
     strWhereActivite  := '';
     strGroupActivite  := '';
     strOrderActivite  := '';
     bRessource        := FALSE;
     bTiers            := FALSE;
     bAffaire          := FALSE;
     bTypeArticle      := FALSE;
     bArticle          := FALSE;
     bShowAllRessource := FALSE;

     // Regroupement principal: ressource, affaire ou tiers/affaire
     if TRadioButton (GetControl ('RBRESS')).Checked = TRUE then
     begin
          strRegroup          := 'R';
          bRessource          := TRUE;
          bShowAllRessource   := TCheckBox (GetControl ('CBTOUTESLESRESS')).Checked;
     end
     else
         if TRadioButton (GetControl ('RBAFFAIRE')).Checked = TRUE then
         begin
              strRegroup := 'M';
              bAffaire   := TRUE;
         end
         else
             if TRadioButton (GetControl ('RBTIERS')).Checked = TRUE then
             begin
                  strRegroup := 'C';
                  bTiers     := TRUE;
             end;

     // Regroupements complémentaires
     if TCheckBox (GetControl ('CBRESSOURCE')).Checked = TRUE then
     begin
          strRegroup := strRegroup + 'R';
          bRessource := TRUE;
     end;
     if TCheckBox (GetControl ('CBTIERS')).Checked = TRUE then
     begin
          strRegroup := strRegroup + 'C';
          bTiers     := TRUE;
     end;
     if TCheckBox (GetControl ('CBAFFAIRE')).Checked = TRUE then
     begin
          strRegroup := strRegroup + 'M';
          bAffaire   := TRUE;
     end;
     if TCheckBox (GetControl ('CBTYPEARTICLE')).Checked = TRUE then
     begin
          strRegroup   := strRegroup + 'T';
          bTypeArticle := TRUE;
     end;
     if TCheckBox (GetControl ('CBARTICLE')).Checked = TRUE then
     begin
          strRegroup := strRegroup + 'A';
          bArticle   := TRUE;
     end;

     // Type de présentation: par ressource, puis par ... tiers/affaire/article, au choix
     if strRegroup = '' then
     begin
          PgiInfo ('Veuillez spécifié l''axe d''analyse (au moins le principal)');
          PanelPlanning.Caption := '';
          PanelPlanning.Font.Style := [];
          PanelPlanning.Refresh;
          exit;
     end;

     // Champs d'agrégat
     if TRadioButton (GetControl ('RBQTE')).Checked = TRUE then
         strAgregat := 'ACT_QTEUNITEREF'
     else
         if TRadioButton (GetControl ('RBPR')).Checked = TRUE then
             strAgregat := 'ACT_TOTPRCHARGE'
         else
             strAgregat := 'ACT_TOTVENTE';

     // Fourchette de dates
     DateDebut := strToDate (THEdit (GetControl ('ACT_DATEACTIVITE')).Text);
     DateFin   := strToDate (THedit (GetControl ('ACT_DATEACTIVITE_')).Text);

     // Critères de filtre sur les ressources, tiers, affaires, articles
     strWhereRess       := GetRessourceFiltre ('');
     strWhereTiers      := GetTiersFiltre ('');
     strWhereTiersCompl := GetTiersComplFiltre ('');
     strWhereAffaire    := GetAffaireFiltre ('');
     strWhereArticle    := GetArticleFiltre ('');

     // Traitement des niveaux de regroupement: clauses SELECT, GROUP BY et ORDER
     for i := 1 to Length (strRegroup) do
     begin
          case strRegroup [i] of
               'R':
                   begin
                        strSelectActivite := strSelectActivite + ',ARS_RESSOURCE';
                        strGroupActivite  := strGroupActivite + ',ARS_RESSOURCE';
                        strOrderActivite  := strOrderActivite + ',ARS_RESSOURCE';
                        if bShowAllRessource = FALSE then
                           strSelectActivite := strSelectActivite + ',MAX(ARS_LIBELLE) AS LIBRESSOURCE';
                   end;

               'C':
                   begin
                        strSelectActivite := strSelectActivite + ',T_TIERS,MAX(T_LIBELLE) AS LIBTIERS';
                        strGroupActivite  := strGroupActivite + ',T_TIERS';
                        strOrderActivite  := strOrderActivite + ',T_TIERS';
                   end;

               'M':
                   begin
                        // Sur mission, le tiers doit être dans la rupture
                        if bTiers = FALSE then
                        begin
                             strSelectActivite := strSelectActivite + ',T_TIERS';
                             strGroupActivite  := strGroupActivite + ',T_TIERS';
                             strOrderActivite  := strOrderActivite + ',T_TIERS';
                        end;

                        // Rupture sur les parties visibles
                        strSelectActivite := strSelectActivite + ',AFF_AFFAIRE1';
                        strGroupActivite  := strGroupActivite + ',AFF_AFFAIRE1';
                        strOrderActivite  := strOrderActivite + ',AFF_AFFAIRE1';
                        if VH_GC.CleAffaire.NbPartie > 1 then
                        begin
                             if VH_GC.CleAffaire.Co2Visible = TRUE then
                             begin
                                  strSelectActivite := strSelectActivite + ',AFF_AFFAIRE2';
                                  strGroupActivite  := strGroupActivite + ',AFF_AFFAIRE2';
                                  strOrderActivite  := strOrderActivite + ',AFF_AFFAIRE2';
                                  if (VH_GC.CleAffaire.NbPartie = 3) AND (VH_GC.CleAffaire.Co3Visible = TRUE) then
                                  begin
                                       strSelectActivite := strSelectActivite + ',AFF_AFFAIRE3';
                                       strGroupActivite  := strGroupActivite + ',AFF_AFFAIRE3';
                                       strOrderActivite  := strOrderActivite + ',AFF_AFFAIRE3';
                                  end;
                             end;
                        end;
                        if VH_GC.CleAffaire.GestionAvenant = TRUE then
                        begin
                             strSelectActivite := strSelectActivite + ',AFF_AVENANT';
                             strGroupActivite  := strGroupActivite + ',AFF_AVENANT';
                             strOrderActivite  := strOrderActivite + ',AFF_AVENANT';
                        end;
                   end;

               'T':
                   begin
                        strSelectActivite := strSelectActivite + ',GA_TYPEARTICLE';
                        strGroupActivite  := strGroupActivite + ',GA_TYPEARTICLE';
                        strOrderActivite  := strOrderActivite + ',GA_TYPEARTICLE';
                   end;

               'A':
                   begin
                        strSelectActivite := strSelectActivite + ',GA_CODEARTICLE,MAX(GA_LIBELLE) AS LIBARTICLE';
                        strGroupActivite  := strGroupActivite + ',GA_CODEARTICLE';
                        strOrderActivite  := strOrderActivite + ',GA_CODEARTICLE';
                   end;
          end;
     end;
     strSelectActivite := Copy (strSelectActivite, 2, Length (strSelectActivite)-1) + ',ACT_DATEACTIVITE,SUM(' + strAgregat + ') AS LETOTAL,COUNT(' + strAgregat + ') AS LENOMBRE FROM ACTIVITE';
     strGroupActivite  := Copy (strGroupActivite, 2, Length (strGroupActivite)-1)   + ',ACT_DATEACTIVITE';
     strOrderActivite  := Copy (strOrderActivite, 2, Length (strOrderActivite)-1)   + ',ACT_DATEACTIVITE';

     // Clause JOIN sur la table activité
     if (bRessource = TRUE) OR (strWhereRess <> '') then
        strJoinActivite := ' JOIN RESSOURCE ON ACT_RESSOURCE=ARS_RESSOURCE';
     if (bTiers = TRUE) OR (bAffaire = TRUE) OR (strWhereTiers <> '') then
        strJoinActivite := strJoinActivite + ' JOIN TIERS ON ACT_TIERS=T_TIERS';
     if strWhereTiersCompl <> '' then
        strJoinActivite := strJoinActivite + ' JOIN TIERSCOMPL ON ACT_TIERS=YTC_TIERS';
     if (bAffaire = TRUE) OR (strWhereAffaire <> '') then
        strJoinActivite := strJoinActivite + ' JOIN AFFAIRE ON ACT_AFFAIRE=AFF_AFFAIRE';
     if (bTypeArticle = TRUE) OR (bArticle = TRUE) OR (strWhereArticle <> '') then
        strJoinActivite := strJoinActivite + ' JOIN ARTICLE ON ACT_ARTICLE=GA_ARTICLE';

     // Clause WHERE sur la table activité
     if (bRessource = FALSE) OR (bShowAllRessource = FALSE) then
        strWhereActivite := strWhereActivite + strWhereRess;
     if (bTiers = TRUE) OR (strWhereTiers <> '') then
        strWhereActivite := strWhereActivite + strWhereTiers;
     if strWhereTiersCompl <> '' then
        strWhereActivite := strWhereActivite + strWhereTiersCompl;
     if (bAffaire = TRUE) OR (strWhereAffaire <> '') then
        strWhereActivite := strWhereActivite + strWhereAffaire;
     if (bTypeArticle = TRUE) OR (bArticle = TRUE) OR (strWhereArticle <> '') then
        strWhereActivite := strWhereActivite + strWhereArticle;
     strWhereActivite := strWhereActivite + GetActiviteFiltre;
     if TCheckBox (GetControl ('CBIGNORENUL')).Checked = TRUE then
        strWhereActivite := strWhereActivite + ' AND ' + strAgregat + '<>0';
     strWhereActivite := Copy (strWhereActivite, 5, Length (strWhereActivite)-4);

     // Construction de la TOB des lignes activité, soit en une seule requête (si pas complétude des ressources) ou en deux
     TOBPlanning  := TOB.Create ('le planning', nil, -1);
     if (bRessource = FALSE) OR (bShowAllRessource = FALSE) then
     begin
          // Une seule requête, car pas de complétude sur les ressources (même si regroupement principal sur ressources)
          strRequete := 'SELECT ' + strSelectActivite + strJoinActivite;
          strRequete := strRequete + ' WHERE ' + strWhereActivite + ' GROUP BY ' + strGroupActivite;
          strRequete := strRequete + ' ORDER BY ' + strOrderActivite;
          TOBPlanning.LoadDetailFromSQL (strRequete);
     end
     else
     begin
          // 2 requêtes, 2 TOB intermédiaires, car complétude sur ressources demandée
          TOBRessource := TOB.Create ('les ressources', nil, -1);
          TOBActivite  := TOB.Create ('les activites', nil, -1);

          // Requête de sélection principale: sur les ressources, avec les filtres spécifiés
          strRequete := 'SELECT ARS_RESSOURCE,ARS_LIBELLE FROM RESSOURCE';
          if strWhereRess <> '' then
             strRequete := strRequete + ' WHERE ' + Copy (strWhereRess, 5, Length (strWhereRess)-4);
          strRequete := strRequete + ' ORDER BY ARS_RESSOURCE';
          TOBRessource.LoadDetailFromSQL (strRequete);
          for i := 0 to TOBRessource.Detail.Count-1 do
          begin
               // Libellé de la ressource
               strLibRessource := TOBRessource.Detail [i].GetValue ('ARS_LIBELLE');

               // Requête sur activité pour LA ressource
               strRequete := 'SELECT ' + strSelectActivite + strJoinActivite;
               strRequete := strRequete + ' WHERE ACT_RESSOURCE="' + TOBRessource.Detail [i].GetValue ('ARS_RESSOURCE') + '" ';
               strRequete := strRequete + ' AND ' + strWhereActivite + ' GROUP BY ' + strGroupActivite;
               strRequete := strRequete + ' ORDER BY ' + strOrderActivite;

               // Chargement des lignes pour LA ressource
               TOBActivite.LoadDetailFromSQL (strRequete);
               if TOBActivite.Detail.Count > 0 then
               begin
                    while TOBActivite.Detail.Count > 0 do
                    begin
                         TOBActivite.Detail [0].AddChampSupValeur ('LIBRESSOURCE', strLibRessource);
                         if TOBActivite.Detail [0].ChangeParent (TOBPlanning, -1) = FALSE then
                         begin
                              PgiInfo ('Erreur lors de la préparation des données (ressource ' + TOBActivite.Detail [0].GetValue('ACT_RESSOURCE') + ')', 'Suivi activité par planning');
                              exit;
                         end;
                    end
               end
               else
               begin
                    // Au moins une ligne par ressource, vide
                    with TOB.Create ('', TOBPlanning, -1) do
                    begin
                         AddChampSupValeur ('ARS_RESSOURCE', TOBRessource.Detail [i].GetValue ('ARS_RESSOURCE'));
                         AddChampSupValeur ('LIBRESSOURCE', strLibRessource);
                         AddChampSupValeur ('T_TIERS', '');
                         AddChampSupValeur ('LIBTIERS', '');
                         AddChampSupValeur ('AFF_AFFAIRE1', '');
                         AddChampSupValeur ('LIBAFFAIRE', '');
                         AddChampSupValeur ('GA_TYPEARTICLE', '');
                         AddChampSupValeur ('GA_CODEARTICLE', '');
                         AddChampSupValeur ('LIBARTICLE', '');

                         AddChampSupValeur ('ACT_DATEACTIVITE', 0);
                    end;
               end;

               // Ressource traitée, lignes d'activité pour ressource suivante
               TOBActivite.ClearDetail;
          end;

          TOBRessource.Free;
          TOBActivite.Free;
     end;

     // Affichage du planning sur la TOB des lignes d'activité construite
     AffichePlanning (TOBPlanning);
     TOBPlanning.Free;
end;

procedure TOF_AFACTIVITE_SUIVI.OnClickImprimer (Sender:TObject);
begin
     if (SuiviPlanning <> nil) AND (SuiviPlanning.TobItems.Detail.Count > 0) then
     begin
          SuiviPlanning.TypeEtat    := 'E';
          SuiviPlanning.NatureEtat  := 'APL';
          SuiviPlanning.CodeEtat    := 'APL';
          SuiviPlanning.Print ('Suivi d''activité du ' + FormatDateTime ('dd/mm/yyyy', DateDebut) + ' au ' + FormatDateTime ('dd/mm/yyyy', DateFin));
     end;
end;

procedure TOF_AFACTIVITE_SUIVI.OnClickExporter (Sender: TObject);
var
   C      :TSaveDialog;
begin
     if (SuiviPlanning <> nil) AND (SuiviPlanning.TobItems.Detail.Count > 0) then
     begin
          C := TSaveDialog.Create (Ecran);
          try
             C.Options := [ofOverwritePrompt];
             C.Filter  := 'Microsoft Excel (*.xls)|*.XLS';
             if C.Execute then
             begin
                  SourisSablier;
                  try
                     SuiviPlanning.Visible := FALSE;
                     PanelPlanning.Caption := 'Export vers classeur Microsoft Excel "' + C.FileName + '" en cours...';
                     PanelPlanning.Refresh;
                     SuiviPlanning.ExportToExcel (TRUE, C.FileName);
                  finally
                         SuiviPlanning.Visible := TRUE;
                         PanelPlanning.Caption := '';
                         PanelPlanning.Refresh;
                         SourisNormale;
                  end;
             end;
          finally
                 C.Free;
          end;
     end;
end;

procedure TOF_AFACTIVITE_SUIVI.OnClickMaximiser (Sender: TObject);
begin
     MaximisePlanning;
end;

procedure TOF_AFACTIVITE_SUIVI.NomsChampsAffaire (var Aff, Aff0,Aff1, Aff2, Aff3,Aff4, Aff_, Aff0_,Aff1_, Aff2_, Aff3_,Aff4_, Tiers, Tiers_:THEdit);
begin
     Aff   := THEdit (GetControl('AFF_AFFAIRE'));
     Aff1  := THEdit (GetControl('AFF_AFFAIRE1'));
     Aff2  := THEdit (GetControl('AFF_AFFAIRE2'));
     Aff3  := THEdit (GetControl('AFF_AFFAIRE3'));
     Aff4  := THEdit (GetControl('AFF_AVENANT'));
     Tiers := THEdit (GetControl('T_TIERS'));
end;

procedure TOF_AFACTIVITE_SUIVI.OnKeyMaxWindow (Sender: TObject; var Key: Word; Shift: TShiftState);
begin
     if (Key = VK_ESCAPE) and (Sender is TForm) then
        TForm (Sender).Close;
end;

procedure TOF_AFACTIVITE_SUIVI.MaximisePlanning;
var
   ZoomSuivi     :TForm;
begin
     if (SuiviPlanning <> nil) AND (SuiviPlanning.TobItems.Detail.Count > 0) then
     begin
          // Affichage dans une fenêtre maximisée
          ZoomSuivi             := TForm.Create (Ecran);
          ZoomSuivi.Caption     := 'Suivi de l''activité';
          ZoomSuivi.BorderIcons := [biSystemMenu, biMinimize, biMaximize];
          ZoomSuivi.BorderStyle := bsSizeable;
          ZoomSuivi.WindowState := wsMaximized;
          ZoomSuivi.KeyPreview  := TRUE;
          ZoomSuivi.OnKeyDown   := OnKeyMaxWindow;
          SuiviPlanning.Parent  := ZoomSuivi;
          PanelPlanning.Caption := 'Suivi d''activité affiché dans une fenêtre agrandie';
          ZoomSuivi.ShowModal;

          // Fin de la fenêtre maximiser: le planning "revient" dans la fenêtre de lancement
          SuiviPlanning.Parent := PanelPlanning;
          ZoomSuivi.Free;
     end;
end;

procedure TOF_AFACTIVITE_SUIVI.OnDblClickPlanning (Sender: TObject);
var
   T               :TOB;
   strArgument     :string;
   iNumChamp       :integer;
begin
     // Pour l'instant, on considère que si lancement immédiat, c'est qu'on connait déjà les critères, et que le zoom est peu utile
     // En fait, pour éviter de boucler sur la saisie activité
     if bImmediat = TRUE then
        exit;

     T := SuiviPlanning.GetCurItem;
     if (T <> Nil) AND (T.GetValue ('I_ETAT') <> '002') then
     begin
          // Détermination des critères à afficher dans la consultation ligne d'activité
          strArgument := 'XX_WHERE=ACT_DATEACTIVITE="' + USDATETIME (T.GetValue ('I_DATEDEBUT')) + '"';
          if bRessource = TRUE then
             strArgument := strArgument + ' AND AFACTAFFTIERS.ARS_RESSOURCE="' + T.GetValue ('I_CODERESS') + '"';
          if bTiers = TRUE then
             strArgument := strArgument + ' AND AFACTAFFTIERS.T_TIERS="' + T.GetValue ('I_CODETIERS') + '"';
          if bAffaire = TRUE then
          begin
               if bTiers = FALSE then
                  strArgument := strArgument + ' AND AFACTAFFTIERS.T_TIERS="' + T.GetValue ('I_CODETIERS') + '"';
               iNumChamp := T.GetNumChamp ('I_CODEAFF1');
               if iNumChamp > 0 then
                  strArgument := strArgument + ' AND AFACTAFFTIERS.AFF_AFFAIRE1="' + T.GetValeur (iNumChamp) + '"';
               iNumChamp := T.GetNumChamp ('I_CODEAFF2');
               if iNumChamp > 0 then
                  strArgument := strArgument + ' AND AFACTAFFTIERS.AFF_AFFAIRE2="' + T.GetValeur (iNumChamp) + '"';
               iNumChamp := T.GetNumChamp ('I_CODEAFF3');
               if iNumChamp > 0 then
                  strArgument := strArgument + ' AND AFACTAFFTIERS.AFF_AFFAIRE3="' + T.GetValeur (iNumChamp) + '"';
               iNumChamp := T.GetNumChamp ('I_CODEAFFAV');
               if iNumChamp > 0 then
                  strArgument := strArgument + ' AND AFACTAFFTIERS.AFF_AVENANT="' + T.GetValeur (iNumChamp) + '"';
          end;
          if bTypeArticle = TRUE then
             strArgument := strArgument + ' AND AFACTAFFTIERS.GA_TYPEARTICLE="' + T.GetValue ('I_TYPEARTICLE') + '"';
          if bArticle = TRUE then
             strArgument := strArgument + ' AND AFACTAFFTIERS.GA_CODEARTICLE="' + T.GetValue ('I_CODEARTICLE') + '"';

          // Argument complémentaires
          strArgument := strArgument + GetRessourceFiltre ('AFACTAFFTIERS.') + GetTiersFiltre ('AFACTAFFTIERS.') + GetTiersComplFiltre ('AFACTAFFTIERS.') + GetAffaireFiltre ('AFACTAFFTIERS.') + GetArticleFiltre ('AFACTAFFTIERS.') + GetActiviteFiltre;

          // Consultation avec critères déjà positionnés sur les éléments choisi dans le suivi
          AFLanceFiche_Consult_Activite (strArgument);
     end;
end;

procedure AFLanceFiche_Suivi_Activite (strArgument:string);
begin
     AGLLanceFiche ('AFF','AFACTIVITE_SUIVI','','', strArgument);
end;


initialization
              registerclasses ([TOF_AFACTIVITE_SUIVI]);
end.

