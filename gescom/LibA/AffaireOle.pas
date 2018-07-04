{***********A.G.L.***********************************************
Auteur  ...... : JP
Créé le ...... : 04/07/2001
Modifié le ... :   /  /
Description .. : Unit de définition des liens OLE de la GI/GA
Suite ........ : Contient les fonctions OLE utilisable dans WORD (parser), EXCEL...
Mots clefs ... : GI;GA;LIENOLE;EXCEL;WORD
*****************************************************************}

unit AffaireOle;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  Windows, Messages,SysUtils,Classes,Graphics,Controls,Forms,Dialogs,
{$IFNDEF EAGLCLIENT}
  MajTable,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  OLEAuto,OleDicoPGI,UTob,HEnt1,HCtrls ,aftableaubord;


const
     COleTiersChamps     = 'T_ABREGE,T_ADRESSE1,T_ADRESSE2,T_ADRESSE3,T_APE,T_CODEPOSTAL,T_CORRESP1,T_CORRESP2,T_DEVISE,T_EMAIL,T_FAX,T_FERME,T_FORMEJURIDIQUE,T_JURIDIQUE,T_LIBELLE,T_MODEREGLE,T_NATIONALITE,T_PAYS,';
     COleTiers2Champs    = 'T_PRENOM,T_REGIMETVA,T_RVA,T_SIRET,T_TELEPHONE,T_TELEPHONE2,T_TELEX,T_TIERS,T_VILLE';
     COleAnnuaireChamps  = 'ANN_ALCP,ANN_ALRUE1,ANN_ALRUE2,ANN_ALRUE3,ANN_ALVILLE,ANN_CAPDEV,ANN_CAPITAL,ANN_CAPLIB,ANN_CAPNBTITRE,ANN_CLESIRET,ANN_CODENAF,ANN_CODEPER,ANN_EMAIL,ANN_ENSEIGNE,ANN_FAMPER,ANN_FAX,ANN_FORME,ANN_FORMEGEN,ANN_NATIONALITE,';
     COleAnnuaire2Champs = 'ANN_NOM1,ANN_NOM2,ANN_NOM3,ANN_NOMPER,ANN_PAYS,ANN_PROFESSION,ANN_RCS,ANN_RCSDATE,ANN_RCSGEST,ANN_RCSNOREF,ANN_RCSVILLE,ANN_RM,ANN_RMANNEE,ANN_RMDEP,ANN_RMNOREF,ANN_SIREN,ANN_TEL1,ANN_TIERS,ANN_TYPEPER,ANN_VOIENO,ANN_VOIENOCOMPL,ANN_VOIENOM,ANN_VOIETYPE';
     COleDPFisChamps     = 'DFI_ADHEREOGA,DFI_ANNEECIVILE,DFI_CA,DFI_CTRLFISC,DFI_DATEADHOGA,DFI_DEMATERIATDFC,DFI_INTEGRAFISC,DFI_NOADHOGA,DFI_NODP,DFI_NOFRP,DFI_NOINTRACOMM,DFI_NOOGADP,DFI_OLDCA,DFI_RESULTFISC';
     COleDPSocChamps     = 'DSO_CE,DSO_DATEDADS,DSO_DATEEFFECTIF,DSO_EFFAPPRENTIS,DSO_EFFECTIF,DSO_EFFHANDICAP,DSO_EFFMOYEN,DSO_MTTDADS,DSO_NODP';
     COleDPEcoChamps     = 'DEC_CA,DEC_DATECA,DEC_DATERESULTAT,DEC_NODP,DEC_RESULTAT';
     COleContactChamps   = 'C_CIVILITE,C_FAX,C_FONCTION,C_NOM,C_PRENOM,C_PRINCIPAL,C_RVA,C_SERVICE,C_TELEPHONE,C_TELEX,C_TEXTELIBRE1';
     COlePieceChamps     = 'GP_DATEPIECE,GP_DEVISE,GP_MODEREGLE,GP_NATUREPIECEG,GP_NUMPIECE,GP_REFCOMPTABLE,GP_REFEXTERNE,GP_REFINTERNE,GP_REGIMETAXE,GP_RIB,GP_TARIFTIERS,GP_TIERSFACTURE,GP_TIERSPAYEUR,GP_TOTALESC,GP_TOTALESCDEV,GP_TOTALHT,GP_TOTALHTDEV,';
     COlePiece2Champs    = 'GP_TOTALREMISE,GP_TOTALREMISEDEV,GP_TOTALTTC,GP_TOTALTTCDEV,GP_TOTESCTTC,GP_TOTESCTTCDEV,GP_TOTREMISETTC,GP_TOTREMISETTCDEV,GP_TVAENCAISSEMENT';
     COleLigneChamps     = 'GL_ARTICLE,GL_CODEARTICLE,GL_DATEPIECE,GL_LIBCOMPL,GL_LIBELLE,GL_MONTANTHTDEV,GL_MONTANTTTCDEV,GL_NATUREPIECEG,GL_PUHT,GL_PUHTDEV,GL_PUHTNET,GL_PUHTNETDEV,GL_PUTTC,GL_PUTTCDEV,GL_PUTTCNET,GL_PUTTCNETDEV,GL_QTEFACT,GL_TOTALHT,';
     COleLigne2Champs    = 'GL_TOTALHTDEV,GL_TYPEARTICLE,GL_TYPELIGNE,GL_BLOCNOTE';
     COleEcheanceChamps  = 'AFA_DATEECHE,AFA_ECHEFACT,AFA_LIBELLEECHE,AFA_LIQUIDATIVE,AFA_MONTANTECHE,AFA_MONTANTECHEDEV,AFA_NUMECHE,AFA_POURCENTAGE,AFA_REPRISEACTIV,AFA_TYPECHE';
     COleEtabChamps      = 'ET_ABREGE,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,ET_APE,ET_CODEPOSTAL,ET_EMAIL,ET_FAX,ET_JURIDIQUE,ET_LIBELLE,ET_PAYS,ET_SIRET,ET_TELEPHONE,ET_TELEX,ET_VILLE';

Type
    // Automation Server
    TAffaireOle = class(TAutoObject)
    automated
             // Pour Lien OLE parser Word
             function New                  :variant;
             function Init                 :variant;
             function Free                 :variant;
             function AfficheDictionnaire  :variant;

             // Démarrage et paramètres des cumuls
//             function RestartAllCumuls     :variant;

             // Récupération des listes des tables "standards": tiers, ...
             function GetTiers             (strNature:string)          :variant;
             function GetArticles          (strTypeArticle:string)     :variant;
             function GetRessources        (strTypeRessource:string)   :variant;
             function GetAffaires          (strStatut, strEtat:string) :variant;

             // Récupération de données d'un champ d'une table ou d'une tablette, avec la clé (unique) spécifiée
             function GetTableValeur       (strChamp, strValeurCle:string)    :variant;
             function GetTabletteValeur    (strType:string)                   :variant;
             function GetLibelleTablette   (strTablette, strValeurCle:string) :variant;

             // Champ d'une table
             function GetChamps            (strPrefixes:string; strTypes:string):variant;

             // Fonctions OLE pour paramétrer et lancer les cumuls sur une table prédéfinie
             // $$$jp 30/04/2003: remplace les 8 fonctions suivantes
             function StartCumul           (strTable:string; strTiersDeb, strTiersFin:string):variant;
             function GetCumul             (strChamp:string):variant; //iValorisation:integer):variant;
             function ZoomCumul            (strChamp:string):variant; //iValorisation:integer):variant;
             function EndCumul             ():variant;
             function StartCumulParam      :variant;
             function SetCumulParam        ({strPrefixeTableCumul:string;} strCodeParam:string; vInf, vSup:variant):variant; //string):variant;

             // Cumul sur table activité
//             function StartCumulActivite   (strTiersDeb, strTiersFin:string):variant;
  //           function GetCumulActivite     (strChamp:string):variant; //iValorisation:integer):variant;
    //         function ZoomCumulActivite    (strChamp:string):variant; //iValorisation:integer):variant;
      //       function EndCumulActivite     ():variant;

             // Cumul sur tableau de bord
  //           function GetTableauBord       (strChamp:string):variant; //strMission, strStatutMission, strArticle, strTypeArticle, strAssistant, strChamp, strVisa, strDateDeb, strDateFin:string):variant;   // function GetTableauBord       (strClient, strMission, strPrestation, strAssistant, strType, strRegroup, strVisa:string):variant;
    //         function ZoomTableauBord      (strChamp:string):variant; //strMission, strStatutMission, strArticle, strTypeArticle, strAssistant, strChamp, strVisa, strDateDeb, strDateFin:string):variant;   // function GetTableauBord       (strClient, strMission, strPrestation, strAssistant, strType, strRegroup, strVisa:string):variant;
      //       function EndTableauBord       ():variant;

    public
             constructor    Create;        override;
             destructor     Destroy;       override;

    protected
             ParamTB                       :TParamTBAffaire;
             m_bModeZoom                   :boolean;
             m_strTableCumul               :string;
             m_iIndexTable                 :integer;
             m_strTiersDeb                 :array [0..3] of string;
             m_strTiersFin                 :array [0..3] of string;
             m_bAsked                      :boolean;

             // Paramètres (filtres) du cumul (activité ou tableau de bord)
             m_iNbParams                   :integer;
             m_strJoin                     :string;
             m_strWhere                    :string;
             m_strArgZoom                  :string;

             function       LienOleEnabled  ():boolean;
             procedure      OnTheLast       (var shutdown:boolean);

             function       GetIndexCode    (strSuiteCode:string; strMax:string):integer;
             function       GetPrefixeTable :string;
//             function       GetTableCumul   (strChamp, strTable, strPrefixe:string):double;

             // Construction de la table AFTABLEAUBORD
             function BuildTableauBord:boolean;
    end;

procedure RegisterAffairePGI;
procedure AFOLE_GetTobFunc (theTobFunc:TOB);
procedure AFOLE_GetTobDico (theTobDico:TOB);


implementation

uses dicoBTP, affaireutil, utofafactivitemul, utofaftabbordole, hmsgbox,
     Paramsoc, utilGa,UFonctionsCBP
     ,CbpMCD
     ,CbpEnumerator
     ;



procedure RegisterAffairePGI;
const
  AutoClassInfo: TAutoClassInfo = (
    AutoClass:    TAffaireOle;
    ProgID:       'CegidPGI.Affaire';
    ClassID:      '{252E1A31-EBAE-4A58-9050-3486DC0A1755}';
    Description:  'Cegid Gestion PGI';
    Instancing:   acMultiInstance);
begin
     Automation.RegisterClass (AutoClassInfo);
end;

{Function PGOLEErreur(i : Integer) : String ;
begin
     Result:='#ERREUR:' ;
     Case i Of
          1  : Result:=Result+'Cumul non renseigné'  ;
          2  : Result:=Result+'Rubrique non renseignée' ;
          3  : Result:=Result+'Mois début non renseigné' ;
          4  : Result:=Result+'Année début non renseignée' ;
          5  : Result:=Result+'Champ à cumuler non renseigné' ;
          6  : Result:=Result+'Nature rubrique non renseignée' ;
          7  : Result:=Result+'tablette non renseignée' ;
          10 : Result:=Result+'générale SQL !' ;
     else
         Result:=Result+'Non définie '+IntToStr(i) ;
  END ;
end;}

function TAffaireOle.LienOleEnabled ():boolean;
begin
     Result := ExJaiLeDroitConcept (ccOLE, FALSE);
     if (Result = FALSE) AND (m_bAsked = FALSE) then
     begin
          PgiInfo('Vous n''avez pas les droits d''utilisation du lien OLE', 'Lien OLE');
          m_bAsked := TRUE;
     end;
end;

procedure TAffaireOle.OnTheLast (var shutdown:boolean);
begin
     ShutDown := FALSE ;  // pour ne pas fermer l'application automatiquement lorsqu'il n'y a plus de réference à l'objet
end;

constructor TAffaireOle.Create;
begin
     inherited;

     // Propriétés du serveur OLE
     Automation.OnLastRelease := OnTheLast;
     m_bAsked                 := FALSE;

     // Initialisations pour les cumuls
     ParamTB         := nil;
     m_iIndexTable   := -1;
     m_strTableCumul := '';
//     RestartAllCumuls;
end;

destructor TAffaireOle.Destroy;
begin
     ParamTB.Free;
     inherited;
end;


                                    //---------------------------------------------------
                                    // LIEN OLE POUR DICTIONNAIRE DOCPARSER DANS WORD
                                    //---------------------------------------------------
function TAffaireOle.AfficheDictionnaire:variant;
begin
     AfficherDictionnairePGI;
     result := unassigned;
end;

function TAffaireOle.Free:variant;
begin
     FreeDictionnairePGI;
     result := unassigned;
end;

function TAffaireOle.New:variant;
begin
     NewDictionnairePGI;
     result := unassigned;
end;

function TAffaireOle.Init:variant;
//var
  // $$$ JP 07/07/03 AppliPathStd  :string;
begin
     // $$$ JP 07/07/03: amélioration identification STD ET DAT
//{$IFNDEF EAGLCLIENT}
  //       if V_PGI_ENV <> nil then
    //     begin
//{$IFDEF GIGI}
  //            AppliPathStd := V_PGI_ENV.PathStd + '\GIS5';
//{$ELSE}
  //            AppliPathStd := V_PGI_ENV.PathStd + '\GAS5';
//{$ENDIF}
//         end
  //       else
    //     begin
      //        // V_PGI_ENV inexistant, pas normal!
        //      AppliPathStd := 'C:\PGI00\STD';
//         end;
//{$ENDIF}
     InitDictionnairePGI (AFOLE_GetTobDico, AFOLE_GetTobFunc, nil, GetAppliPathStd + '\ParseDic.txt'); // $$$ JP 07/07/03: amélioration identification STD ET DAT
end;

procedure AFOLE_GetTobFunc (theTobFunc:TOB);
var
   T, T1     :TOB;
begin
     // Gestion du nb d'éléments dans les tables utilisables (affaire, tiers, ...)
     T  := TOB.Create ('Nombre d éléments :', theTobFunc, -1); //$$$JP 28/01/03 modif nom tob

     T1 := TOB.Create ('contacts / tiers', T, -1);
     T1.AddChampSupValeur ('NOM', 'contact / tiers');
     T1.AddChampSupValeur ('PROTO', '@NBELEMENT("CONTACT")');
     T1.AddChampSupValeur ('HELP', 'nombre de contact d''un tiers');

     T1 := TOB.Create ('prestations / affaire', T, -1);
     T1.AddChampSupValeur ('NOM', 'prestations / affaire / tiers');
     T1.AddChampSupValeur ('PROTO', '@NBELEMENT("PRESTATION")');
     T1.AddChampSupValeur ('HELP', 'nombre de prestations d''une affaire');

     T1 := TOB.Create ('tâches / affaire', T, -1);
     T1.AddChampSupValeur ('NOM', 'tâches / affaire');
     T1.AddChampSupValeur ('PROTO', '@NBELEMENT("TACHE")');
     T1.AddChampSupValeur ('HELP', 'nombre de tâches d''une affaire');

     T1 := TOB.Create ('échéances / affaire / tiers', T, -1);
     T1.AddChampSupValeur ('NOM', 'échéances / affaire');
     T1.AddChampSupValeur ('PROTO', '@NBELEMENT("ECHEANCE")');
     T1.AddChampSupValeur ('HELP', 'nombre d''échéance de paiement d''une affaire');
end;

procedure AFOLE_GetTobDico (theTobDico:TOB);
var
   T1, T2         :TOB;
   i, iTableIndex :integer;
   strNom         :string;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
     theTobDico.ClearDetail;

     // Positionnement sur la table Tiers
     T1          := TOB.Create ('Table ' + TraduitGA('Tiers'), theTobDico, -1); //$$$JP 28/01/03 modif nom tob
    table := Mcd.getTable('TIERS');
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
     begin
          strNom := (FieldList.Current as IFieldCOM).name;
          if (Pos (strNom, COleTiersChamps) > 0) OR (Pos (strNom, COleTiers2Champs) > 0) then
          begin
               T2 := TOB.Create ('nom champ', T1, -1);
               T2.AddChampSupValeur ('DH_NOMCHAMP', strNom);
               T2.AddChampSupValeur ('DH_LIBELLE', (FieldList.Current as IFieldCOM).libelle);
          end;
     end;

     // Table affaire
     T1          := TOB.Create ('Table ' + TraduitGA('Affaire'), theTobDico, -1); //$$$JP 28/01/03 modif nom tob
    table := Mcd.getTable('AFFAIRE');
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
     begin
          strNom := (FieldList.Current as IFieldCOM).name;
          T2 := TOB.Create ('nom champ', T1, -1);
          T2.AddChampSupValeur ('DH_NOMCHAMP', strNom);
          T2.AddChampSupValeur ('DH_LIBELLE', (FieldList.Current as IFieldCOM).libelle);
     end;

     // Table contact
     T1          := TOB.Create ('Table ' + TraduitGA('Contact'), theTobDico, -1); //$$$JP 28/01/03 modif nom tob
    table := Mcd.getTable('CONTACT');
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
     begin
          strNom := (FieldList.Current as IFieldCOM).name;
          if Pos (strNom, COleContactChamps) > 0 then
          begin
               T2 := TOB.Create ('nom champ', T1, -1);
               T2.AddChampSupValeur ('DH_NOMCHAMP', (FieldList.Current as IFieldCOM).name);
               T2.AddChampSupValeur ('DH_LIBELLE', (FieldList.Current as IFieldCOM).libelle);
          end;
     end;

     // Table pièce
     T1          := TOB.Create ('Table ' + TraduitGA('Pièce'), theTobDico, -1); //$$$JP 28/01/03 modif nom tob
    table := Mcd.getTable('PIECE');
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
     begin
          strNom := (FieldList.Current as IFieldCOM).name;
          if (Pos (strNom, COlePieceChamps) > 0) OR (Pos (strNom, COlePiece2Champs) > 0) then
          begin
               T2 := TOB.Create ('nom champ', T1, -1);
               T2.AddChampSupValeur ('DH_NOMCHAMP', strNom);
               T2.AddChampSupValeur ('DH_LIBELLE', (FieldList.Current as IFieldCOM).libelle);
          end;
     end;

     // Table ligne de pièces
     T1          := TOB.Create ('Table ' + TraduitGA('Ligne de pièce'), theTobDico, -1); //$$$JP 28/01/03 modif nom tob
    table := Mcd.getTable('LIGNE');
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
     begin
          strNom := (FieldList.Current as IFieldCOM).name;
          if (Pos (strNom, COleLigneChamps) > 0) OR (Pos (strNom, COleLigne2Champs) > 0) then
          begin
               T2 := TOB.Create ('nom champ', T1, -1);
               T2.AddChampSupValeur ('DH_NOMCHAMP', strNom);
               T2.AddChampSupValeur ('DH_LIBELLE', (FieldList.Current as IFieldCOM).libelle);
          end;
     end;

     // Table Echeances
     T1          := TOB.Create ('Table ' + TraduitGA('Echéance'), theTobDico, -1); //$$$JP 28/01/03 modif nom tob
    table := Mcd.getTable('FACTAFF');
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
     begin
          strNom := (FieldList.Current as IFieldCOM).name;
          if Pos (strNom, COleEcheanceChamps) > 0 then
          begin
               T2 := TOB.Create ('nom champ', T1, -1);
               T2.AddChampSupValeur ('DH_NOMCHAMP', strNom);
               T2.AddChampSupValeur ('DH_LIBELLE', (FieldList.Current as IFieldCOM).libelle);
          end;
     end;

     // Table Annuaire
     T1          := TOB.Create ('Table ' + TraduitGA('Annuaire'), theTobDico, -1); //$$$JP 28/01/03 modif nom tob
    table := Mcd.getTable('ANNUAIRE');
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
     begin
          strNom := (FieldList.Current as IFieldCOM).name;
          if (Pos (strNom, COleAnnuaireChamps) > 0) OR (Pos (strNom, COleAnnuaire2Champs) > 0) then
          begin
               T2 := TOB.Create ('nom champ', T1, -1);
               T2.AddChampSupValeur ('DH_NOMCHAMP', strNom);
               T2.AddChampSupValeur ('DH_LIBELLE', (FieldList.Current as IFieldCOM).libelle);
          end;
     end;

     // Table DP Fiscal
     T1          := TOB.Create ('Table ' + TraduitGA('DPFiscal'), theTobDico, -1); //$$$JP 28/01/03 modif nom tob
    table := Mcd.getTable('DPFISCAL');
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
     begin
          strNom := (FieldList.Current as IFieldCOM).name;
          if Pos (strNom, COleDPFisChamps) > 0 then
          begin
               T2 := TOB.Create ('nom champ', T1, -1);
               T2.AddChampSupValeur ('DH_NOMCHAMP', strNom);
               T2.AddChampSupValeur ('DH_LIBELLE',(FieldList.Current as IFieldCOM).Libelle);
          end;
     end;

     // Table DP Social
     T1          := TOB.Create ('Table ' + TraduitGA('DPSocial'), theTobDico, -1); //$$$JP 28/01/03 modif nom tob
    table := Mcd.getTable('DPSOCIAL');
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
     begin
          strNom := (FieldList.Current as IFieldCOM).name;
          if Pos (strNom, COleDPSocChamps) > 0 then
          begin
               T2 := TOB.Create ('nom champ', T1, -1);
               T2.AddChampSupValeur ('DH_NOMCHAMP', strNom);
               T2.AddChampSupValeur ('DH_LIBELLE', (FieldList.Current as IFieldCOM).Libelle);
          end;
     end;

     // Table DP Economique
     T1          := TOB.Create ('Table ' + TraduitGA('DPEco'), theTobDico, -1); //$$$JP 28/01/03 modif nom tob
    table := Mcd.getTable('DPECO');
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
     begin
          strNom := (FieldList.Current as IFieldCOM).name;
          if Pos (strNom, COleDPEcoChamps) > 0 then
          begin
               T2 := TOB.Create ('nom champ', T1, -1);
               T2.AddChampSupValeur ('DH_NOMCHAMP', strNom);
               T2.AddChampSupValeur ('DH_LIBELLE', (FieldList.Current as IFieldCOM).libelle);
          end;
     end;
end;


                                    //---------------------------------------------------
                                    // LIEN OLE POUR FEUILLE DE SYNTHESE (DANS EXCEL)
                                    //---------------------------------------------------
{function TAffaireOle.RestartAllCumuls ():variant;
begin
     Result := '';

     // Fin activité et tableau de bord
     EndCumulActivite;
     EndTableauBord;

     // Supprime les paramètres courants
     ResetCumulParam;
end;}

function TAffaireOle.GetArticles (strTypeArticle:string):variant;
var
   strRequete     :string;
   TOBArt         :TOB;
   i              :integer;
begin
     Result     := '';
     if LienOleEnabled = FALSE then
        exit;

     // Requête d'extraction de la liste des articles du dossier
     strRequete := 'SELECT GA_CODEARTICLE,GA_LIBELLE,GA_TYPEARTICLE FROM ARTICLE';
     if strTypeArticle <> '' then
        strRequete := strRequete + ' WHERE GA_TYPEARTICLE="' + strTypeArticle + '"';
     strRequete := strRequete + ' ORDER BY GA_CODEARTICLE';

     TOBArt := TOB.Create ('les articles', nil, -1);
     try
        TOBArt.LoadDetailFromSQL (strRequete);
        for i := 0 to TOBArt.Detail.Count - 1 do
            Result := Result + TOBArt.Detail [i].GetValue ('GA_CODEARTICLE') + '§' + TOBArt.Detail [i].GetValue ('GA_LIBELLE') + '§' + TOBArt.Detail [i].GetValue ('GA_TYPEARTICLE')+ '§';
     finally
            TOBArt.Free;
     end;
end;

function TAffaireOle.GetRessources (strTypeRessource:string):variant;
var
   strRequete     :string;
   TOBRess        :TOB;
   i              :integer;
begin
     Result := '';
     if LienOleEnabled = FALSE then
        exit;

     // Requête d'extraction de la liste des ressources du dossier
     strRequete := 'SELECT ARS_RESSOURCE,ARS_LIBELLE,ARS_LIBELLE2,ARS_TYPERESSOURCE FROM RESSOURCE';
     if strTypeRessource <> '' then
        strRequete := strRequete + ' WHERE ARS_TYPERESSOURCE="' + strTypeRessource + '"';
     strRequete := strRequete + ' ORDER BY ARS_RESSOURCE';

     // Boucle sur la liste des ressources
     TOBRess := TOB.Create ('les ressources', nil, -1);
     try
        TOBRess.LoadDetailFromSQL (strRequete);
        for i := 0 to TOBRess.Detail.Count - 1 do
        begin
             Result := Result + TOBRess.Detail [i].GetValue ('ARS_RESSOURCE') + '§' + TOBRess.Detail [i].GetValue ('ARS_LIBELLE');
             if TOBRess.Detail [i].GetValue ('ARS_LIBELLE2') <> '' then
                Result := Result + ' ' + TOBRess.Detail [i].GetValue ('ARS_LIBELLE2');
             Result := Result + '§' + TOBRess.Detail [i].GetValue ('ARS_TYPERESSOURCE') + '§';
        end;
     finally
            TOBRess.Free;
     end;
end;

function TAffaireOle.GetTiers (strNature:string):variant;
var
   strRequete     :string;
   TOBTiers       :TOB;
   i              :integer;
begin
     Result     := '';
     if LienOleEnabled = FALSE then
        exit;

     strRequete := 'SELECT T_TIERS, T_LIBELLE FROM TIERS';
     if strNature <> '' then
        strRequete := strRequete + ' WHERE T_NATUREAUXI="' + strNature + '"';
     strRequete := strRequete + ' ORDER BY T_TIERS';

     TOBTiers := TOB.Create ('les tiers', nil, -1);
     try
        TOBTiers.LoadDetailFromSQL (strRequete);
        for i := 0 to TOBTiers.Detail.Count - 1 do
            Result := Result + TOBTiers.Detail [i].GetValue ('T_TIERS') + '§' + TOBTiers.Detail [i].GetValue ('T_LIBELLE') + '§';
     finally
            TOBTiers.Free;
     end;
end;

function TAffaireOle.GetAffaires(strStatut, strEtat: string): variant;
var
   strRequete     :string;
   TOBAffaire     :TOB;
   i              :integer;
begin
     Result     := '';
     if LienOleEnabled = FALSE then
        exit;

     strRequete := 'SELECT AFF_AFFAIRE, AFF_TIERS, AFF_AFFAIRE0, AFF_AFFAIRE1, AFF_AFFAIRE2, AFF_AFFAIRE3, AFF_AVENANT, AFF_LIBELLE, AFF_STATUTAFFAIRE, AFF_ETATAFFAIRE FROM AFFAIRE';
     if strStatut <> '' then
        strRequete := strRequete + ' WHERE AFF_STATUTAFFAIRE="' + strStatut + '"';
     if strEtat <> '' then
        if strStatut <> '' then
            strRequete := strRequete + ' AND AFF_ETATAFFAIRE="' + strEtat + '"'
        else
            strRequete := strRequete + ' WHERE AFF_ETATAFFAIRE="' + strEtat + '"';
     strRequete := strRequete + ' ORDER BY AFF_TIERS, AFF_AFFAIRE';

     TOBAffaire := TOB.Create ('les affaires', nil, -1); //$$$JP 28/01/03 modif nom tob
     try
        TOBAffaire.LoadDetailFromSQL (strRequete);
        for i := 0 to TOBAffaire.Detail.Count - 1 do
        begin
             Result := Result + TOBAffaire.Detail [i].GetValue ('AFF_AFFAIRE') + '§';
             Result := Result + TOBAffaire.Detail [i].GetValue ('AFF_TIERS') + '§';
             Result := Result + TOBAffaire.Detail [i].GetValue ('AFF_AFFAIRE0') + '§';
             Result := Result + TOBAffaire.Detail [i].GetValue ('AFF_AFFAIRE1') + '§';
             Result := Result + TOBAffaire.Detail [i].GetValue ('AFF_AFFAIRE2') + '§';
             Result := Result + TOBAffaire.Detail [i].GetValue ('AFF_AFFAIRE3') + '§';
             Result := Result + TOBAffaire.Detail [i].GetValue ('AFF_AVENANT') + '§';
             Result := Result + TOBAffaire.Detail [i].GetValue ('AFF_LIBELLE') + '§';
             Result := Result + TOBAffaire.Detail [i].GetValue ('AFF_STATUTAFFAIRE') + '§';
             Result := Result + TOBAffaire.Detail [i].GetValue ('AFF_ETATAFFAIRE') + '§';
        end;
     finally
            TOBAffaire.Free;
     end;
end;

function TAffaireOle.GetTabletteValeur (strType:string):variant;
var
   strRequete     :string;
   TOBChoix       :TOB;
   i              :integer;
begin
     Result     := '';
     if LienOleEnabled = FALSE then
        exit;

     if strType = '' then
        exit;

     if Copy (strType, 1, 2) = 'FN' then
         strRequete := 'SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="' + strType + '" ORDER BY CC_CODE'
     else
         strRequete := 'SELECT YX_CODE,YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="' + strType + '" ORDER BY YX_CODE';

     TOBChoix := TOB.Create ('les choix', nil, -1);
     try
        TOBChoix.LoadDetailFromSQL (strRequete);
        for i := 0 to TOBChoix.Detail.Count - 1 do
        begin
             if Copy (strType, 1, 2) = 'FN' then
                 Result := Result + TOBChoix.Detail [i].GetValue ('CC_CODE') + '§' + TOBChoix.Detail [i].GetValue ('CC_LIBELLE') + '§'
             else
                 Result := Result + TOBChoix.Detail [i].GetValue ('YX_CODE') + '§' + TOBChoix.Detail [i].GetValue ('YX_LIBELLE') + '§';
        end;
     finally
            TOBChoix.Free;
     end;
end;

function TAffaireOle.GetLibelleTablette (strTablette, strValeurCle:string):variant;
begin
     Result := RechDom (strTablette, strValeurCle, FALSE);
end;

function TAffaireOle.GetTableValeur (strChamp, strValeurCle: string):variant;
var
   Q   :TQuery;
begin
     // Requête adaptée à la table contenant le champ
     Result := '<???>';
     if Copy (strChamp, 1, 4) = 'AFF_' then
         Q := OpenSQL ('SELECT ' + strChamp + ' FROM AFFAIRE WHERE AFF_AFFAIRE="' + strValeurCle + '"', TRUE,-1,'',true)
     else
         if Copy (strChamp, 1, 4) = 'ARS_' then
             Q := OpenSQL ('SELECT ' + strChamp + ' FROM RESSOURCE WHERE ARS_RESSOURCE="' + strValeurCle + '"', TRUE,-1,'',true)
         else
             if Copy (strChamp, 1, 4) = 'YTC_' then
                 Q := OpenSQL ('SELECT ' + strChamp + ' FROM TIERSCOMPL WHERE YTC_TIERS="' + strValeurCle + '"', TRUE,-1,'',true)
             else
                 if Copy (strChamp, 1, 2) = 'T_' then
                     Q := OpenSQL ('SELECT ' + strChamp + ' FROM TIERS WHERE T_TIERS="' + strValeurCle + '"', TRUE,-1,'',true)
                 else
                     if Copy (strChamp, 1, 3) = 'GA_' then
                         Q := OpenSQL ('SELECT ' + strChamp + ' FROM ARTICLE WHERE GA_CODEARTICLE="' + strValeurCle + '"', TRUE,-1,'',true)
                     else
                         Q := nil;

     // Renvoie la valeur du champ demandé, pour la clé spécifiée
     if Q <> nil then
     begin
          if not Q.eof then
             Result := Q.FindField (strChamp).AsVariant;
          Ferme (Q);
     end;
end;

function TAffaireOle.GetIndexCode (strSuiteCode:string; strMax:string):integer;
begin
     Result := 0;
     if strSuiteCode = 'A' then
     begin
          if strMax = '9' then
             Result := 10
     end
     else
         if (strSuiteCode >= '1') and (strSuiteCode <= strMax) then
             Result := StrToInt (strSuiteCode);
end;

function TAffaireOle.GetPrefixeTable:string;
begin
     // Préfixe des champs de la table en cours de cumul
     case m_iIndexTable of
          0:
            Result := 'ACT';

          1:
            Result := 'ATB';

          2:
            Result := 'GL';

          3:
            Result := 'AFA';
     else
         Result := '';
     end;
end;

function TAffaireOle.StartCumulParam:variant;
begin
     m_iNbParams     := 0;
     m_strJoin       := '';
     m_strWhere      := '';
     m_strArgZoom    := '';
end;

function TAffaireOle.SetCumulParam ({strPrefixeTableCumul:string;} strCodeParam:string; vInf, vSup:variant):variant; //string):variant;
var
   strType         :string;
   strSubString    :string;
   strChampFiltre  :string;
   strPrefixe      :string;
   strPrefixeParam :string;
   iType           :integer;
   iPos            :integer;
   bFiltre         :boolean;
begin
     // Par défaut, champ non reconnu
     Result := '#erreur: paramètre de cumul incorrect';

     // Test du préfixe de table de cumul: soit ATB, soit ACT (pour l'instant)
//     if (strPrefixeTableCumul <> 'ACT') and (strPrefixeTableCumul <> 'ATB') then
  //      exit;

     // Extrait le nom du champ (sans les éventuels paramètres de sous-chaine)
     iPos := Pos (',', strCodeParam);
     if iPos > 0 then
     begin
          strSubString := Copy (strCodeParam, iPos, Length (strCodeParam)-iPos+1);
          strCodeParam := Copy (strCodeParam, 1, iPos-1);
     end;

     // Le code doit être un champ, sauf quelques cas particuliers
     if m_iNbParams < 16 then
     begin
          strType := ChampToType (strCodeParam);
          if strType <> '??' then
          begin
               strPrefixe      := GetPrefixeTable;
               strPrefixeParam := ExtractPrefixe (strCodeParam);

               // Détermination des borne inférieure et supérieur du filtre
               bFiltre := FALSE;
               if (VarIsEmpty (vInf) = TRUE) OR (VarType (vInf) = varError) then
                  vInf := UnAssigned
               else
                   bFiltre := TRUE;
               if (VarIsEmpty (vSup) = TRUE) OR (VarType (vSup) = varError) then
                  vSup := UnAssigned
               else
                   bFiltre := TRUE;

               // Type du champ pour les requêtes: soit numérique, soit alphanumérique, soit date (à inverser pour la requête)
               if (strType = 'DOUBLE') OR (strType = 'INTEGER') then
                  iType := 0
               else
                   // $$$JP 26/03/2003 - gestion d'un type supplémentaire: COMBO
                   if (Copy (strType, 1, 7) = 'VARCHAR') OR (Copy (strType, 1, 4) = 'CHAR') OR (strType = 'BOOLEAN') OR (strType = 'DATE') OR (strType = 'COMBO') then
                   begin
                        iType := 1;

                        // Traitement des filtres sur date (les bornes doivent exister)
                        if strType = 'DATE' then
                        begin
                             bFiltre := TRUE;
                             if VarIsEmpty (vInf) then
                                vInf := '01/01/1900';
                             if VarIsEmpty (vSup) then
                                vSup := '31/12/2099';

                             // Conversion jj/mm/aaaa en mm/jj/aaaa pour la requête
                             vInf := Copy (vInf, 4, 3) + Copy (vInf, 1, 3) + Copy (vInf, 7, 4);
                             vSup := Copy (vSup, 4, 3) + Copy (vSup, 1, 3) + Copy (vSup, 7, 4);
                        end;
                   end
                   else
                       bFiltre := FALSE;

               // si le champ filtre provient d'une table annexe, on doit déclarer les jointures (sauf si filtre "vide")
               if bFiltre = TRUE then
               begin
                    if strPrefixeParam = 'AFF' then
                    begin
                         if Pos ('AFF_AFFAIRE', m_strJoin) = 0 then
                            m_strJoin := m_strJoin + ' JOIN AFFAIRE ON ' + strPrefixe + '_AFFAIRE=AFF_AFFAIRE';
                    end
                    else if strPrefixeParam = 'ARS' then
                    begin
                         if Pos ('ARS_RESSOURCE', m_strJoin) = 0 then
                            m_strJoin := m_strJoin + ' JOIN RESSOURCE ON ' + strPrefixe + '_RESSOURCE=ARS_RESSOURCE';
                    end
                    else if strPrefixeParam = 'YTC' then
                    begin
                         if Pos ('YTC_TIERS', m_strJoin) = 0 then
                            m_strJoin := m_strJoin + ' JOIN TIERSCOMPL ON ' + strPrefixe + '_TIERS=YTC_TIERS';
                    end
                    else if strPrefixeParam = 'GA' then
                    begin
                         if Pos ('GA_CODEARTICLE', m_strJoin) = 0 then
                            m_strJoin := m_strJoin + ' JOIN ARTICLE ON ' + strPrefixe + '_CODEARTICLE=GA_CODEARTICLE';
                    end
                    else if strPrefixeParam = 'T' then
                    begin
                         if Pos ('T_TIERS', m_strJoin) = 0 then
                            m_strJoin := m_strJoin + ' JOIN TIERS ON ' + strPrefixe + '_TIERS=T_TIERS';
                    end
                    else if strPrefixeParam <> strPrefixe then
                         exit;

                    // Construction champ filtre (peut posséder une clause de sous-chaine)
                    if strSubString <> '' then
                        strChampFiltre := 'SUBSTRING(' + strCodeParam + strSubString + ')'
                    else
                        strChampFiltre := strCodeParam;

                    // Construction complément requête pour filtre
                    if VarIsEmpty (vInf) = FALSE then
                    begin
                         if VarIsEmpty (vSup) = FALSE then
                         begin
                              if m_strWhere <> '' then
                                 m_strWhere := m_strWhere + ' AND ';
                              m_strWhere := m_strWhere + '(' + strChampFiltre + '>=';
                              if iType = 1 then
                                  m_strWhere := m_strWhere + '"' + vInf + '"'
                              else
                                  m_strWhere := m_strWhere + vInf;
                              m_strWhere := m_strWhere + ' AND ' + strChampFiltre + '<=';
                              if iType = 1 then
                                  m_strWhere := m_strWhere + '"' + vSup + '"'
                              else
                                  m_strWhere := m_strWhere + vSup;
                              m_strWhere := m_strWhere + ')';
                         end
                         else
                         begin
                              if m_strWhere <> '' then
                                 m_strWhere := m_strWhere + ' AND ';
                              m_strWhere := m_strWhere + '(' + strChampFiltre + '=';
                              if iType = 1 then
                                  m_strWhere := m_strWhere + '"' + vInf + '"'
                              else
                                  m_strWhere := m_strWhere + vInf;
                              m_strWhere := m_strWhere + ')';
                         end;
                    end
                    else
                    begin
                         if VarIsEmpty (vSup) = FALSE then
                         begin
                              if m_strWhere <> '' then
                                 m_strWhere := m_strWhere + ' AND ';
                              m_strWhere := m_strWhere + '(' + strChampFiltre + '=';
                              if iType = 1 then
                                  m_strWhere := m_strWhere + '"' + vSup + '"'
                              else
                                  m_strWhere := m_strWhere + vSup;
                              m_strWhere := m_strWhere + ')';
                         end;
                    end;

                    // Argument à passer lors d'un zoom sur cumul
                    if VarIsEmpty (vInf) = FALSE then
                       m_strArgZoom := m_strArgZoom + 'OLE' + strCodeParam + '=' + VarAsType (vInf, varString) + ';';
                    if VarIsEmpty (vSup) = FALSE then
                       m_strArgZoom := m_strArgZoom + 'OLE' + strCodeParam + '_=' + VarAsType (vSup, varString) + ';';
               end;

               // Pas d'erreur "bloquante"
               Result := '';
          end;
     end;
end;

//function TAffaireOle.GetTableCumul (strChamp, strTable, strPrefixe:string):double;
//end;

function TAffaireOle.StartCumul (strTable:string; strTiersDeb, strTiersFin:string):variant;
var
   iIndexTable    :integer;
begin
     Result := '';
     if LienOleEnabled = FALSE then
        exit;

     // Identification de la table sujette au cumul
     if strTable = 'ACTIVITE' then
         m_iIndexTable := 0
     else
         if strTable = 'AFTABLEAUBORD' then
             m_iIndexTable := 1
         else
             if strTable = 'LIGNE' then
                 m_iIndexTable := 2
             else
                 if strTable = 'FACTAFF' then
                     m_iIndexTable := 3
                 else
                 begin
                      Result := '#erreur: table ' + strTable + ' inconnue';
                      m_iIndexTable   := -1;
                      m_strTableCumul := '';
                      exit;
                 end;

     // Si client change (sur une même table), on force le recalcul
     if (m_strTiersDeb [m_iIndexTable] <> strTiersDeb) OR (m_strTiersFin [m_iIndexTable] <> strTiersFin) then
        EndCumul; //RestartAllCumuls;

     // Table à cumuler et bornes sur code tiers
     m_strTableCumul               := strTable;
     m_strTiersDeb [m_iIndexTable] := strTiersDeb;
     m_strTiersFin [m_iIndexTable] := strTiersFin;

     // Paramètre de cumul remis à zéro
     StartCumulParam;

     // Initialisations spécifique à chaque table
     case m_iIndexTable of
          1:
            // Si tableau de bord non construit, on le fait
            if ParamTB = nil then
               if BuildTableauBord = FALSE then
                  Result := '#erreur: impossible de générer AFTABLEAUBORD';
     end;
end;

function TAffaireOle.GetCumul (strChamp:string):variant;
var
     Q               :TQuery;
     strRequete      :string;
     strType         :string;
     strPrefixe      :string;
begin
     // Par défaut, résultat nul
     Result := 0.0;

     // Préfixe des champs de la table en cours de cumul: si non défini, cumul impossible
     strPrefixe := GetPrefixeTable;
     if strPrefixe = '' then
        exit;

     // Génération de la TOB Tableau de bord avec les filtres définis
     strType := ChampToType (strChamp);
     if (strType = 'DOUBLE') OR (strType = 'INTEGER') then
     begin
          // Agrégat
          strRequete := 'SELECT SUM(' + strChamp + ') AS AGREG FROM ' + m_strTableCumul;

          // Jointures nécessaires
          strRequete := strRequete + m_strJoin;

          // Filtre sur l'agrégat
          if m_strWhere <> '' then
             strRequete := strRequete + ' WHERE ' + m_strWhere;

          // AJout du filtre sur le code tiers si existant (sinon, on prend TOUS les tiers)
          if m_strTiersDeb [m_iIndexTable] <> '' then
          begin
               if m_strTiersFin [m_iIndexTable] <> '' then
               begin
                    if m_strWhere <> '' then
                        strRequete := strRequete + ' AND '
                    else
                        strRequete := strRequete + ' WHERE ';
                    strRequete := strRequete + '(' + strPrefixe + '_TIERS>="' + m_strTiersDeb [m_iIndexTable] + '" ';
                    strRequete := strRequete + 'AND ' + strPrefixe + '_TIERS<="' + m_strTiersFin [m_iIndexTable] + '")';
               end
               else
               begin
                    if m_strWhere <> '' then
                        strRequete := strRequete + ' AND '
                    else
                        strRequete := strRequete + ' WHERE ';
                    strRequete := strRequete + '(' + strPrefixe + '_TIERS="' + m_strTiersDeb [m_iIndexTable] + '")';
               end
          end
          else
          begin
               if m_strTiersFin [m_iIndexTable] <> '' then
               begin
                    if m_strWhere <> '' then
                        strRequete := strRequete + ' AND '
                    else
                        strRequete := strRequete + ' WHERE ';
                    strRequete := strRequete + '(' + strPrefixe + '_TIERS="' + m_strTiersFin [m_iIndexTable] + '")';
               end;
          end;
          Q := OpenSQL (strRequete, TRUE, 1,'',true);
          if not Q.eof then
             Result := Q.FindField ('AGREG').AsFloat;
          Ferme (Q);
     end;
end;

function TAffaireOle.ZoomCumul (strChamp:string):variant;
var
   strArg     :string;
   TOBZoom    :TOB;
begin
     if LienOleEnabled = FALSE then
        exit;

     Result := GetCumul (strChamp);
     if Result <> 0 then
     begin
          m_bModeZoom := TRUE;
          case m_iIndexTable of
               1:
               begin
                    TOBZoom := TOB.Create ('Tableau de bord', nil, -1);
                    try
                       TOBZoom.LoadDetailFromSQL ('SELECT ATB_DATE,ATB_AFFAIRE1,ATB_AFFAIRE2,ATB_AFFAIRE3,ATB_TIERS,ATB_RESSOURCE,ATB_CODEARTICLE,ATB_LIBELLE,' + strChamp + ' FROM AFTABLEAUBORD' + m_strJoin + ' WHERE ' + m_strWhere + ' ORDER BY ATB_DATE,ATB_TIERS,ATB_AFFAIRE,ATB_NUMERO');
                       strArg := 'TOBOLE=' + IntToStr (integer (TOBZoom));
                       strArg := strArg + ';CHAMPOLE=' + strChamp;

                       // Lancement mul de consultation activité
                       AFLanceFiche_TabBordOLE (strArg);
                    finally
                           TOBZoom.Free;
                    end;
               end;
          end;
          m_bModeZoom := FALSE;
     end;
end;

function TAffaireOle.EndCumul ():variant;
begin
     case m_iIndexTable of
          1:
          begin
               ParamTB.Free;
               ParamTB := nil;
          end;
     end;

     m_bModeZoom := FALSE;
end;
{function TAffaireOle.StartCumulActivite (strTiersDeb, strTiersFin:string):variant;
begin
     if LienOleEnabled = FALSE then
        exit;

     // Si client change, on force le recalcul
     if (m_strTiersDeb <> strTiersDeb) OR (m_strTiersFin <> strTiersFin) then
        RestartAllCumuls;
     m_strTiersDeb := strTiersDeb;
     m_strTiersFin := strTiersFin;
end;

function TAffaireOle.EndCumulActivite ():variant;
begin
end;

function TAffaireOle.GetCumulActivite (strChamp:string):variant; //iValorisation:integer):variant;
begin
     // Par défaut, cumul vide
     Result := GetTableCumul (strChamp, 'ACTIVITE', 'ACT'); //0.0;
end;}

{function TAffaireOle.ZoomCumulActivite (strChamp:string):variant; //iValorisation:integer):variant;
begin
     Result := GetCumulActivite (strChamp); //iValorisation);
     if LienOleEnabled = FALSE then
        exit;
//     Result := GetCumulActivite (strClient, strLibCliNum, strLibCliVal, strRessCliNum, strRessCliVal, strMission, strLibMissNum, strLibMissVal, strRessMissNum, strRessMissVal, strAssistant, strPrestation, strLibArtNum, strLibArtVal, strFamArtNum, strFamArtVal, strFact, strValo, strDateDeb, strDateFin); //strClient, strMission, strAssistant, strPrestation, strFact, strValo, strDateDeb, strDateFin);
     if Result <> 0 then
     begin
          // $$$todo: adapter le zoom aux filtres "bornés" du lien OLE cumul activité
          // Si pas de filtre sur date, on fixe les dates "générales"
          if Pos ('OLEACT_DATEACTIVITE', m_strArgZoom) = 0 then
             m_strArgZoom := m_strArgZoom + 'OLEACT_DATEACTIVITE=01/01/1900;OLEACT_DATEACTIVITE_=31/12/2099;';

          // Filtre éventuel sur code tiers
          if m_strTiersDeb <> '' then
             m_strArgZoom := m_strArgZoom + 'OLEACT_TIERS=' + m_strTiersDeb
          else
              if m_strTiersFin <> '' then
                 m_strArgZoom := m_strArgZoom + 'OLEACT_TIERS=' + m_strTiersFin;

          // Lancement mul de consultation activité
          AFLanceFiche_Consult_OLE_Activite (m_strArgZoom);
     end;
end;}

function TAffaireOle.BuildTableauBord:boolean;
var
     ListeChamps                  :string;
     strSQL                       :string;
     strNaturePieceAchat          :string;
     strNaturePieceAchatEngage    :string;
begin
     Result := FALSE;

     // Création du tableau de bord
     try
        // Paramètres de sélection des enregistrements du tableau de bord
        ParamTB := TParamTBAffaire.Create (ttbAffaire, tdaDetail, tniArt, 'H', FALSE);
        ParamTB.AlimPrevu        := TRUE;
        ParamTB.AlimFacture      := TRUE;
        ParamTB.AlimReal         := TRUE;
        ParamTB.AlimAchat        := TRUE;
        ParamTB.AlimValoRess_Art := TRUE;
        ParamTB.AlimAN           := FALSE;
        ParamTB.AlimVisa         := FALSE;
        ParamTB.AlimMois         := FALSE;

        // $$$jp 29/04/2003: alimentation cutoff si autorisé + alimentation boni/mali + gestion facturable
        if GetParamSoc ('SO_AFAPPCUTOFF') = TRUE then
            ParamTB.AlimCutOff       := TRUE
        else
            ParamTB.AlimCutOff       := FALSE;
        ParamTB.AlimBoni         := TRUE;
        ParamTB.GereFacturable   := TRUE;

        ParamTB.CentrFam         := FALSE;
        ParamTB.CentrStatRes     := FALSE;
        ParamTB.CentrStatArt     := FALSE;
        ParamTB.MtActivite       := ActPRPV;
        ParamTB.MajMultiTiers (m_strTiersDeb [m_iIndexTable], m_strTiersFin [m_iIndexTable], '01/01/1900', '31/12/2099');
        ParamTB.DetailParAffaire := TRUE;

        // Génération de la TOB tableau de bord
        ParamTB.AlimTableauBordAffaire (ListeChamps, strSQL, strNaturePieceAchat, strNaturePieceAchatEngage, FALSE, FALSE, FALSE, FALSE, FALSE, FALSE);
        Result := TRUE;
     except
           Result := FALSE;
           ParamTB.Free;
           ParamTB := nil;
     end;
end;

{function TAffaireOle.EndTableauBord ():variant;
begin
     ParamTB.Free;
     ParamTB           := nil;
     m_bModeZoom       := FALSE;
end;

function TAffaireOle.GetTableauBord (strChamp:string):variant;
begin
     // $$$JP 26/03/2003 - il faut filtrer UNIQUEMENT les enregistrements concernant l'utilisateur courant
     if m_strWhere <> '' then
         m_strWhere := m_strWhere + ' AND ';
     m_strWhere := m_strWhere + '(ATB_UTILISATEUR="' + V_PGI.User + '")';

     // Par défaut, cumul vide
     Result := GetTableCumul (strChamp, 'AFTABLEAUBORD', 'ATB'); //0.0;
end;

function TAffaireOle.ZoomTableauBord (strChamp:string):variant;
var
   strArg     :string;
   TOBZoom    :TOB;
begin
     m_bModeZoom := TRUE;
     Result := GetTableauBord (strChamp);
     if LienOleEnabled = FALSE then
        exit;

     if Result <> 0 then
     begin
          TOBZoom := TOB.Create ('Tableau de bord', nil, -1);
          try
             TOBZoom.LoadDetailFromSQL ('SELECT ATB_DATE,ATB_AFFAIRE1,ATB_AFFAIRE2,ATB_AFFAIRE3,ATB_TIERS,ATB_RESSOURCE,ATB_CODEARTICLE,ATB_LIBELLE,' + strChamp + ' FROM AFTABLEAUBORD' + m_strJoin + ' WHERE ' + m_strWhere + ' ORDER BY ATB_DATE,ATB_TIERS,ATB_AFFAIRE,ATB_NUMERO');
             strArg := 'TOBOLE=' + IntToStr (integer (TOBZoom));
             strArg := strArg + ';CHAMPOLE=' + strChamp;

             // Lancement mul de consultation activité
             AFLanceFiche_TabBordOLE (strArg);
          finally
                 TOBZoom.Free;
          end;
     end;
     m_bModeZoom := FALSE;
end;}


function TAffaireOle.GetChamps (strPrefixes, strTypes:string):variant;
var
   sSQL       :string;
   TOBChamps  :TOB;
   strRequete :string;
   i          :integer;
begin
     Result := '';
     if (strPrefixes = '') or (strTypes = '') then
        exit;

     TOBChamps := TOB.Create ('Les champs', nil, -1);
     try
        strRequete := 'SELECT DH_NOMCHAMP,DH_TYPECHAMP,DH_LIBELLE FROM DECHAMPS WHERE ';
        if strTypes <> '*' then
           strRequete := strRequete + 'DH_TYPECHAMP IN (' + strTypes + ') AND ';
        strRequete := strRequete + 'DH_PREFIXE IN (' + strPrefixes + ') ';
        strRequete := strRequete + 'ORDER BY DH_NOMCHAMP';
        TOBChamps.LoadDetailFromSQL (strRequete);
        for i := 0 to TOBChamps.Detail.Count-1 do
            Result := Result + TOBChamps.Detail [i].GetValue ('DH_NOMCHAMP') + '§' + TOBChamps.Detail [i].GetValue ('DH_LIBELLE') + '§' + TOBChamps.Detail [i].GetValue ('DH_TYPECHAMP') + '§';
     finally
            TOBChamps.Free;
     end;
end;

end.

