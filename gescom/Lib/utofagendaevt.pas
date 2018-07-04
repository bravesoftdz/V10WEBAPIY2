{***********UNITE*************************************************
Auteur  ...... : JP
Créé le ...... : 03/06/2003
Modifié le ... : 03/06/2003
Description .. : Gestion fiche lancement planning pour suivi activité
Mots clefs ... : TOF;ACTIVITE;PLANNING;SUIVI
*****************************************************************}
unit utofagendaevt;

Interface

Uses
    StdCtrls, Controls, Classes, forms, sysutils,ComCtrls,Messages,Windows,
{$IFDEF EAGLCLIENT}
   MaineAGL,
{$ELSE}
   dbTables, db,FE_Main,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, HPanel, UTOF , vierge, HTB97,
     UTOB,Affaireutil,AglInitGC,M3FP,Saisutil,Grids,
     DicoAF,graphics,EntGC, Ent1,CalcOleGenericAff,
     utilressource,AfUtilArticle,Utilarticle, UTofAfBaseCodeAffaire,
     HPlanning;

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

           procedure OnClickValider     (Sender: TObject);
           procedure OnClickMaximiser   (Sender: TObject);
           procedure OnClickImprimer    (Sender: TObject);
           procedure OnClickExporter    (Sender: TObject);
           procedure OnDblClickPlanning (Sender: TObject);

           procedure OnKeyMaxWindow   (Sender: TObject; var Key: Word; Shift: TShiftState);

           function  GetSQLFiltre (He:THEdit)                 :string;       overload;
           function  GetSQLFiltre (Hvcb:THValComboBox)        :string;       overload;
           function  GetSQLFiltre (Hmvcb:THMultiValComboBox)  :string;       overload;
           function  GetSQLFiltre (cb:TCheckBox)              :string;       overload;

           function  GetRessourceFiltre                       :string;
           function  GetTiersFiltre                           :string;
           function  GetTiersComplFiltre                      :string;
           function  GetAffaireFiltre                         :string;
           function  GetArticleFiltre                         :string;
           function  GetActiviteFiltre                        :string;

           procedure AffichePlanning (TOBPlanning:TOB);
           procedure MaximisePlanning;
  end;

procedure AFLanceFiche_Suivi_Activite (strArgument:string);


Implementation

uses dialogs, afactivite;

procedure TOF_AFACTIVITE_SUIVI.OnArgument (strArgument:string);
begin
     inherited;

     // Le panneau où s'affichera le planning
     PanelPlanning := THPanel (GetControl ('PPLANNING'));
     
     // Bouton Valider: lancement de la génération du planning
     TToolbarButton97 (GetControl ('BVALIDER')).OnClick := OnClickValider;

     // Bouton Imprimer: lancement de l'impression du planning en cours
     TToolbarButton97 (GetControl ('BIMPRIMER')).Visible := TRUE;
     TToolbarButton97 (GetControl ('BIMPRIMER')).OnClick := OnClickImprimer;

     // Export vers classeur Excel
     TToolbarButton97 (GetControl ('BEXPORTEXCEL')).OnClick := OnClickExporter;

     // Bouton maximiser: pour afficher en plein écran le planning en cours
     TToolBarButton97 (GetControl ('BMAXPLANNING')).OnClick := OnClickMaximiser;

     // Restriction des types d'article sur seulement PRE;MAR;FOU
     THMultiValComboBox (GetControl ('GA_TYPEARTICLE')).Plus := PlusTypeArticle (TRUE);

     // Selon le mode de lancement, initialisation des critères par défaut
     if strArgument = 'RESSOURCE' then
     begin
          TRadioButton (GetControl ('RBRESS')).Checked             := TRUE;
          TCheckBox (GetControl ('CBRESSOURCE')).Enabled           := FALSE;
          TCheckBox (GetControl ('CBTOUTESLESRESS')).Checked                := TRUE;
          TRadioButton (GetControl ('RBQTE')).Checked              := TRUE;
          TRadioButton (GetControl ('RBJOURMOIS')).Checked         := TRUE;
          THMultiValComboBox (GetControl ('GA_TYPEARTICLE')).Value := 'PRE;';
     end;

     // Le planning en cours n'existe pas encore
     TOBRessPlanning := nil;
     TOBEtatPlanning := nil;
     TOBItemPlanning := nil;
     SuiviPlanning   := nil;
end;

procedure TOF_AFACTIVITE_SUIVI.OnClose ;
begin
     inherited ;

     TOBRessPlanning.Free;
     TOBItemPlanning.Free;
     TOBEtatPlanning.Free;
     SuiviPlanning.Free;
end;

function TOF_AFACTIVITE_SUIVI.GetSQLFiltre (He:THEdit):string;
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
          Result := ' AND ' + strChamp;
          case He.Operateur of
               Superieur:
                         Result := Result + '>=';
               Inferieur:
                         Result := Result + '<=';
               Egal:
                         Result := Result + '=';
          end;
          if (Copy (strType, 1, 7) = 'VARCHAR') OR (strType = 'CHAR') OR (strType = 'COMBO') then
              Result := Result + '"' + strValue + '"'
          else
              if (strType = 'DATE') then
                  Result := Result + '"' + USDATETIME (strtodate(strValue)) + '"'
              else
                  if (strType = 'INTEGER') OR (strType = 'DOUBLE') then
                      Result := Result + '"' + strValue
                  else
                      Result := '';
     end;
end;

function TOF_AFACTIVITE_SUIVI.GetSQLFiltre (Hvcb:THValComboBox):string;
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
          Result := ' AND ' + strChamp + '=';
          if (Copy (strType, 1, 7) = 'VARCHAR') OR (strType = 'CHAR') OR (strType = 'COMBO') OR (strType = 'DATE') then
              Result := Result + '"' + strValue + '"'
          else
              if (strType = 'INTEGER') OR (strType = 'DOUBLE') then
                  Result := Result + '"' + strValue
              else
                  Result := '';
     end;
end;

function TOF_AFACTIVITE_SUIVI.GetSQLFiltre (Hmvcb:THMultiValComboBox):string;
var
   strChamp    :string;
   strType     :string;
   strValue    :string;
begin
     // Par défaut, filtre SQL inconnu
     Result := '';
     if hmvcb = nil then
        exit;

     // Si une valeur existe dans le MultiValComboBox, on créer le SQL équivalent
     strValue := Hmvcb.Value;
     if (strValue <> '') AND (UpperCase (strValue) <> '<<TOUS>>') then
     begin
          // Nom et type du champ associé au MultiValComboBox
          strChamp := Hmvcb.Name;
          strType  := ChampToType (strChamp);

          // SQL de la forme: CHAMP IN ("VALEUR1","VALEUR2...)
          Result   := ' AND ' + strChamp + ' IN ("';
          strValue := StringReplace (strValue, ';', '","', [rfReplaceAll]);
          strValue := Copy (strValue, 1, Length (strValue)-2);
          Result   := Result + strValue + ')'
     end;
end;

function TOF_AFACTIVITE_SUIVI.GetSQLFiltre (cb:TCheckBox):string;
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
          Result   := ' AND ' + strChamp + '=';
          if cbValue = cbChecked then
              Result := Result + '"X"'
          else
              Result := Result + '"-"';
     end;
end;

function TOF_AFACTIVITE_SUIVI.GetRessourceFiltre:string;
var
   i   :integer;
begin
     Result := GetSQLFiltre (THEdit (GetControl ('ARS_RESSOURCE')));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('ARS_TYPERESSOURCE')));
     Result := Result + GetSQLFiltre (THEdit (GetControl ('ARS_FONCTION1')));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('ARS_ETABLISSEMENT')));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('ARS_DEPARTEMENT')));
     for i := 1 to 9 do
         Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('ARS_LIBRERES'+IntToStr(i))));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('ARS_LIBRERESA')));
end;

function TOF_AFACTIVITE_SUIVI.GetTiersFiltre:string;
begin
     Result := GetSQLFiltre (THEdit (GetControl ('T_TIERS')));
     Result := Result + GetSQLFiltre (THEdit (GetControl ('T_SOCIETEGROUPE')));
end;

function TOF_AFACTIVITE_SUIVI.GetTiersComplFiltre:string;
var
   i   :integer;
begin
     Result := '';
     for i := 1 to 9 do
         Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('YTC_TABLELIBRETIERS'+IntToStr(i))));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('YTC_TABLELIBRETIERSA')));
     for i := 1 to 3 do
         Result := Result + GetSQLFiltre (THEdit (GetControl ('YTC_RESSOURCE'+IntToStr (i))));
end;

function TOF_AFACTIVITE_SUIVI.GetAffaireFiltre:string;
var
   i   :integer;
begin
     Result := GetSQLFiltre (THEdit (GetControl ('AFF_AFFAIRE')));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('AFF_ETABLISSEMENT')));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('AFF_DEPARTEMENT')));
     Result := Result + GetSQLFiltre (THEdit (GetControl ('AFF_RESPONSABLE')));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('AFF_GROUPECONF')));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('AFF_ETATAFFAIRE')));
     for i := 1 to 9 do
         Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('AFF_LIBREAFF'+IntToStr (i))));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('AFF_LIBREAFFA')));
     for i := 1 to 3 do
         Result := Result + GetSQLFiltre (THEdit (GetControl ('AFF_RESSOURCE'+IntToStr (i))));
end;

function TOF_AFACTIVITE_SUIVI.GetArticleFiltre:string;
var
   i   :integer;
begin
     Result := GetSQLFiltre (THEdit (GetControl ('GA_CODEARTICLE')));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('GA_TYPEARTICLE')));
     for i := 1 to 9 do
         Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('GA_LIBREART'+IntToStr (i))));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('GA_LIBREARTA')));
     for i := 1 to 3 do
         Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('GA_FAMILLENIV'+IntToStr (i))));
end;

function TOF_AFACTIVITE_SUIVI.GetActiviteFiltre:string;
begin
     Result := GetSQLFiltre (THEdit (GetControl ('ACT_DATEACTIVITE')));
     Result := Result + GetSQLFiltre (THEdit (GetControl ('ACT_DATEACTIVITE_')));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('ACT_ACTORIGINE')));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('ACT_ACTIVITEREPRIS')));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('ACT_ETATVISA')));
     Result := Result + GetSQLFiltre (THMultiValComboBox (GetControl ('ACT_ETATVISAFAC')));
     Result := Result + GetSQLFiltre (TCheckBox (GetControl ('ACT_ACTIVITEEFFECT')));
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
   strCleAff0, strCleAff1, strCleAff2, strCleAff3, strCleAffAv                  :string;
   bRupture            :boolean;
   dQte, dCumul        :double;
   i, j                :integer; //, iNumChamp        :integer;
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
{          if bRessource = TRUE then
          begin
               strColNamePlanning  := ';R_LIBRESSOURCE';
               strColSizePlanning  := ';200';
               strColAlignPlanning := ';L';
          end;
          if bTiers = TRUE then
          begin
               strColNamePlanning  := strColNamePlanning + ';R_LIBTIERS';
               strColSizePlanning  := strColSizePlanning + ';200';
               strColAlignPlanning := strColAlignPlanning + ';L';
          end;

          if bAffaireTri = TRUE then
          begin
               strColNamePlanning  := strColNamePlanning + ';R_LIBAFFAIRE';
               strColSizePlanning  := strColSizePlanning + ';200';
               strColAlignPlanning := strColAlignPlanning + ';L';
          end;
          if bTypeArtTri = TRUE then
          begin
               strColNamePlanning  := strColNamePlanning + ';R_LIBTYPEARTICLE';
               strColSizePlanning  := strColSizePlanning + ';50';
               strColAlignPlanning := strColAlignPlanning + ';C';
          end;
          if bArticleTri = TRUE then
          begin
               strColNamePlanning  := strColNamePlanning + ';R_LIBARTICLE';
               strColSizePlanning  := strColSizePlanning + ';150';
               strColAlignPlanning := strColAlignPlanning + ';L';
          end;}
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
                                  // Clé affaire: seulement sur les parties utilisables (en fonction paramètres société)
                                  CodeAffaireDecoupe (TOBFille.GetValue ('AFF_AFFAIRE'), strCleAff0, strCleAff1, strCleAff2, strCleAff3, strCleAffAv, taModif, FALSE);
                                  strCleAff := strCleAff1;
                                  if VH_GC.CleAffaire.NbPartie > 1 then
                                  begin
                                       if VH_GC.CleAffaire.Co2Visible = TRUE then
                                          strCleAff := strCleAff + ' ' + strCleAff2;
                                       if (VH_GC.CleAffaire.NbPartie = 3) AND (VH_GC.CleAffaire.Co3Visible = TRUE) then
                                          strCleAff := strCleAff + ' ' + strCleAff3;
                                  end;
                                  if VH_GC.CleAffaire.GestionAvenant = TRUE then
                                     strCleAff := strCleAff + ' ' + strCleAffAv;

                                  // On doit ajouter la clé tiers si pas sélectionné dans l'analyse
                                  if bTiers = FALSE then
                                      strCleTiers := TOBFille.GetValue ('T_TIERS');
                                  strClePlanning := strClePlanning + '§' + strCleAff;
                                  if bRupture = TRUE then
                                      strCurCleAff := ''
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
{               strClePlanning := '';
               if bRessTri = TRUE then
                  strClePlanning := TOBFille.GetValue ('ARS_RESSOURCE');
               if bTiersTri = TRUE then
                  strClePlanning := strClePlanning + '§' + TOBFille.GetValue ('T_TIERS');
               if bAffaireTri = TRUE then
                  strClePlanning := strClePlanning + '§' + TOBFille.GetValue ('AFF_AFFAIRE');
               if bTypeArtTri = TRUE then
                  strClePlanning := strClePlanning + '§' + TOBFille.GetValue ('GA_TYPEARTICLE');
               if bArticleTri = TRUE then
                  strClePlanning := strClePlanning + '§' + TOBFille.GetValue ('GA_CODEARTICLE');}
               if bRupture = TRUE then //strClePlanning <> strCurClePlanning then
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
                                   AddChampSupValeur ('R_LIBAFFAIRE', Trim (strCleTiers) + ' - ' + Trim (strCleAff))
                               else
                                   AddChampSupValeur ('R_LIBAFFAIRE', Trim (strCleAff))
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
                    if bRessource = TRUE then
                       strCurCleRess  := strCleRess;
                    if bTiers = TRUE then
                       strCurCleTiers := strCleTiers;
                    if bAffaire = TRUE then
                       strCurCleAff := strCleAff;
                    if bTypeArticle = TRUE then
                       strCurCleTypeArt := strCleTypeArt;
                    if bArticle = TRUE then
                       strCurCleArt := strCleArt;
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
                    end
                    else
                    begin
                         AddChampSupValeur ('I_LIBELLE', '');
                         AddChampSupValeur ('I_DATEDEBUT', DateDebut);
                         AddChampSupValeur ('I_DATEFIN', DateFin);
                         AddChampSupValeur ('I_ETAT', '002');
                         AddChampSupValeur ('I_DATA', 0.0);
                    end;
                    AddChampSupValeur ('I_RESSOURCE', '');
                    AddChampSupValeur ('I_TYPE', '');
                    AddChampSupValeur ('I_COLOR', '');
                    AddChampSupValeur ('I_HINT', '');
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
               MaximisePlanning;
          end
          else
          begin
               SuiviPlanning.Visible := FALSE;
               PanelPlanning.Caption := 'Aucune ligne d''activité ne correspond aux critères spécifiés';
               PanelPlanning.Font.Style := [fsBold];
               PanelPlanning.Refresh;
          end;
     end;
end;

procedure TOF_AFACTIVITE_SUIVI.OnClickValider (Sender:TObject);
var
   i                   :integer;
   bShowAllRessource   :boolean;
   strWhereRess        :string;
   strWhereTiers       :string;
   strWhereTiersCompl  :string;
   strWhereAffaire     :string;
   strWhereArticle     :string;
   strSelectActivite   :string;
   strJoinActivite     :string;
   strWhereActivite    :string;
   strGroupActivite    :string;
   strOrderActivite    :string;
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
     strWhereRess       := GetRessourceFiltre;
     strWhereTiers      := GetTiersFiltre;
     strWhereTiersCompl := GetTiersComplFiltre;
     strWhereAffaire    := GetAffaireFiltre;
     strWhereArticle    := GetArticleFiltre;

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
                        strSelectActivite := strSelectActivite + ',AFF_AFFAIRE';
                        strGroupActivite  := strGroupActivite + ',AFF_AFFAIRE';
                        strOrderActivite  := strOrderActivite + ',AFF_AFFAIRE';
                        if bTiers = FALSE then
                        begin
                             strSelectActivite := strSelectActivite + ',T_TIERS';
                             strGroupActivite  := strGroupActivite + ',T_TIERS';
                             strOrderActivite  := strOrderActivite + ',T_TIERS';
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
     strSelectActivite := Copy (strSelectActivite, 2, Length (strSelectActivite)-1) + ',ACT_DATEACTIVITE,SUM(' + strAgregat + ') AS LETOTAL FROM ACTIVITE';
     strGroupActivite  := Copy (strGroupActivite, 2, Length (strGroupActivite)-1) + ',ACT_DATEACTIVITE';
     strOrderActivite  := Copy (strOrderActivite, 2, Length (strOrderActivite)-1) + ',ACT_DATEACTIVITE';

     // Clause SELECT sur la table activité
{     strSelectActivite := '';
     if (bRessTri = TRUE) then
     begin
          strSelectActivite := ',ARS_RESSOURCE';
          if bShowAll = FALSE then
             strSelectActivite := strSelectActivite + ',MAX(ARS_LIBELLE) AS LIBRESSOURCE';
     end;
     if bTiersTri = TRUE then
        strSelectActivite := strSelectActivite + ',T_TIERS,MAX(T_LIBELLE) AS LIBTIERS';
     if bAffaireTri = TRUE then
     begin
          strSelectActivite := strSelectActivite + ',AFF_AFFAIRE,MAX(AFF_LIBELLE) AS LIBAFFAIRE';
          if bTiersTri = FALSE then
             strSelectActivite := strSelectActivite + ',T_TIERS,MAX(T_LIBELLE) AS LIBTIERS';
     end;
     if bTypeArtTri = TRUE then
        strSelectActivite := strSelectActivite + ',GA_TYPEARTICLE,GA_TYPEARTICLE AS LIBTYPEARTICLE';
     if bArticleTri = TRUE then
        strSelectActivite := strSelectActivite + ',GA_CODEARTICLE,MAX(GA_LIBELLE) AS LIBARTICLE';
     strSelectActivite := Copy (strSelectActivite, 2, Length (strSelectActivite)-1) + ',ACT_DATEACTIVITE,SUM(' + strAgregat + ') AS LETOTAL FROM ACTIVITE';
 }

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
        strJoinActivite := strJoinActivite + ' JOIN ARTICLE ON ACT_CODEARTICLE=GA_CODEARTICLE';

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
     strWhereActivite := Copy (strWhereActivite, 5, Length (strWhereActivite)-4);

     // Clause GROUP BY de la table activité
{     if bRessource = TRUE then
        strGroupActivite := ',ARS_RESSOURCE';
     if bTiers = TRUE then
        strGroupActivite := strGroupActivite + ',T_TIERS';
     if bAffaire = TRUE then
     begin
          strGroupActivite := strGroupActivite + ',AFF_AFFAIRE';
          if bTiers = FALSE then
             strGroupActivite := strGroupActivite + ',T_TIERS';
     end;
     if bTypeArticle = TRUE then
        strGroupActivite := strGroupActivite + ',GA_TYPEARTICLE';
     if bArticle = TRUE then
        strGroupActivite := strGroupActivite + ',GA_CODEARTICLE';
     strGroupActivite := Copy (strGroupActivite, 2, Length (strGroupActivite)-1) + ',ACT_DATEACTIVITE';}

     // Clause ORDER de la table activité
{     if bRessource = TRUE then
        strOrderActivite := ',ARS_RESSOURCE';
     if bTiers = TRUE then
        strOrderActivite := strOrderActivite + ',T_TIERS';
     if bAffaire = TRUE then
     begin
          strOrderActivite := strOrderActivite + ',AFF_AFFAIRE';
          if bTiers = FALSE then
             strOrderActivite := strOrderActivite + ',T_TIERS';
     end;
     if bTypeArticle = TRUE then
        strOrderActivite := strOrderActivite + ',GA_TYPEARTICLE';
     if bArticle = TRUE then
        strOrderActivite := strOrderActivite + ',GA_CODEARTICLE';
     strOrderActivite := Copy (strOrderActivite, 2, Length (strOrderActivite)-1) + ',ACT_DATEACTIVITE';
 }

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
          ZoomSuivi.BorderIcons := [biSystemMenu];
          ZoomSuivi.BorderStyle := bsSingle;
          ZoomSuivi.WindowState := wsMaximized;
          ZoomSuivi.KeyPreview  := TRUE;
          ZoomSuivi.OnKeyDown   := OnKeyMaxWindow;
          SuiviPlanning.Parent  := ZoomSuivi;
          ZoomSuivi.ShowModal;

          // Fin de la fenêtre maximiser: le planning "revient" dans la fenêtre de lancement
          SuiviPlanning.Parent := PanelPlanning;
          ZoomSuivi.Free;
     end;
end;

procedure AFLanceFiche_Suivi_Activite (strArgument:string);
begin
     AGLLanceFiche ('AFF','AFACTIVITE_SUIVI','','', strArgument);
end;

procedure TOF_AFACTIVITE_SUIVI.OnDblClickPlanning (Sender: TObject);
var
   T      :TOB;
begin
{     T := LePlanning.GetCurItem;
     if (T<>Nil) AND (T.Detail.Count>0) then
     begin
          //mcd 17/01/2003 si saisie manager on interdit la saisie sur un utilisatuer # du sien
          if SaisieActivitemanager = FALSE then
          begin
               if T.GetValue ('I_CODERESS') <> VH_GC.ressourceUser then
               begin
                    PgiInfo ('Vous n''avez pas le droit d''accès');
                    exit;
               end;
          end
          else
          begin
               AFCreerActiviteModale (tsaRess, tacGlobal, 'REA', T.GetValue ('I_CODERESS'), T.GetValue ('I_CODEAFF'), T.GetValue ('I_CODETIERS'), T.GetValue ('I_DATEDEBUT'));
               exit;
          end;

          AFCreerActiviteModale(tsaClient, tacGlobal, 'REA', T.GetValue ('I_CODERESS'), T.GetValue ('I_CODEAFF'), T.GetValue ('I_CODETIERS'), T.GetValue ('I_DATEDEBUT'));
     end;}
end;

initialization
              registerclasses ([TOF_AFACTIVITE_SUIVI]);
end.

