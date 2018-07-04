{***********UNITE*************************************************
Auteur  ...... : JP
Créé le ...... : 21/06/2002
Modifié le ... :   /  /
Description .. : Source TOM de la FICHE : LETTMISS ()
Mots clefs ... : TOM;LETTMISS
*****************************************************************}
Unit utofLanceParser ;

Interface

uses {$IFDEF VER150} variants,{$ENDIF} StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HDB,
     paramsoc,
     vierge,
     UTOF,
     UTOB;

Type
    TTIERSTOB = class (TOB)
//           TOBTiersComp   :TOB;
           TOBAnnuaire    :TOB;
           TOBDPFiscal    :TOB;
           TOBDPSocial    :TOB;
           TOBDPEco       :TOB;
           TOBSalarie     :TOB;
           TOBContact     :TOB;

           m_strActivite  :string;
           m_strAbrevJu   :string;
           m_strJuridique :string;

           constructor    Create (LeNomTable:string; LeParent:TOB; IndiceFils:integer); override;
           destructor     Destroy;                                                      override;
           procedure      LoadDetailTiers (strTiers:string);
    end;

    TAFFTOB = class (TOB)
           TOBPrestation  :TOB;
           TOBEcheance    :TOB;
           TOBPiece       :TOB;
//           TOBTache       :TOB;
           TOBEtab        :TOB;
           TOBTiers       :TTIERSTOB;

           m_strModeRegl    :string;
           m_strLibelleAff1 :string;

           constructor    Create (LeNomTable:string; LeParent:TOB; IndiceFils:integer); override;
           destructor     Destroy;                                                      override;
           procedure      LoadDetailAffaire (strAffaire:string;strAffaire2:string;strTiers:string;strCodeGroupe:string;strEtab:string);
    end;

  TOF_GENPARSER = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnImprime  ();
    procedure OnClose                  ; override ;

  private
         TOBDoc         :TOB;
         TOBAff         :TOB;
//         TOBEtab        :TOB;
//         TOBRubrique    :TOB;

         // En cours de génération, TOB pointant sur l'affaire en cours
         TOBDataPrincipale  :TOB;
         TOBDataGroupee     :TOB;
         TOBDataAffaire     :TOB;       // Celle en cours
         TOBPrincipale      :TAFFTOB;
         TOBGroupee         :TAFFTOB;
         TOBAffaire         :TAFFTOB;   // Celle en cours

         // Clé de l'enregistrement maitre (affaire, tiers ou assistant)
//         strAnnTiers        :string;
         strAssistant       :string;
         bFromFicheAff      :boolean;

         strNature          :string;
         strEtat            :string;
         iNoModele          :integer;

         procedure SetCurrentAffaire (iIndex1, iIndex2:integer);

         procedure ParseGetVar     (Sender: TObject; VarName: String; VarIndx: Integer; var Value: Variant );
         procedure ParseSetVar     (Sender: TObject; VarName: String; VarIndx: Integer; Value: Variant );
         procedure ParseGetList    (Sender: TObject; IdentListe: String; var Liste: TStringList );
         procedure ParseFunction   (Sender: TObject; FuncName: String; Params: Array of variant ; var Result: Variant );

         procedure ParseInitialize (Sender: TObject );
         procedure ParseFinalize   (Sender: TObject );
  end ;

procedure AFLanceFiche_LanceParser (Range, Argument:string);

Implementation

uses
{$IFNDEF EAGLCLIENT}
     DOC_Parser, fe_main, AffaireOle,
{$ELSE}
     MainEagl,
{$ENDIF}
     AGLUtilOLE, M3FP, FileCtrl,extctrls,
     windows, ent1, dicobtp, utilword, hstatus, dialogs, entGC, utilGa;

//var
  // strNaturePiece   :string;


function GetBlobFusion (Texte, TypeCh:string):variant;
var
      Lignes            : HTStrings;
      Retour,Chaine     : string;
      RichEdit          : TRichEdit;
      Panel             : TPanel;
      sTempo            : string;
      i                 : integer;
begin
    Panel := TPanel.Create( nil );
    Panel.Visible := False;
    Panel.ParentWindow := GetDesktopWindow;
    RichEdit := TRichEdit.Create(Panel);
    RichEdit.Parent := Panel;
    Lignes := HTStringList.Create;
    Lignes.Text := Texte;
    StringsToRich(RichEdit,Lignes);
    Chaine := RichEdit.Text;
    sTempo := '';
    For i :=1 to Length(Chaine) do
    begin
         if TypeCh = 'S' then
         begin
              if Copy(Chaine,i,1) = #$A
                 then sTempo := Stempo + #10//#10
                 else sTempo := Stempo + Copy(Chaine,i,1);
         end
         else
         begin
              if Copy(Chaine,i,1) = #$D
                 then sTempo := Stempo + ' '
                 else if Copy(Chaine,i,1) = #$A
                      then sTempo := Stempo + ' '
                      else sTempo := Stempo + Copy(Chaine,i,1);
         end;
    end;
    Chaine := Trim (sTempo);

    // Retour à la ligne forcée
    if ( Length (Chaine)>0 ) and ((Chaine [Length (Chaine)]<>#10) or (Chaine [Length (Chaine)-1]<>#13)) then
       Chaine := Chaine + #13 + #10;
    Retour := Copy(Chaine,0,Length(Chaine)); //-pos);
    Lignes.Free;
    RichEdit.Free;
    Panel.Free;
    Result := Retour;
end;

constructor TAFFTOB.Create (LeNomTable:string; LeParent:TOB; IndiceFils:integer);
begin
     inherited;

     TOBPrestation  := TOB.Create ('Les prestations', nil, -1);
//     TOBTache       := TOB.Create ('les taches', nil, -1);
     TOBEcheance    := TOB.Create ('Les echeances', nil, -1);
     TOBPiece       := TOB.Create ('La piece', nil, -1);
     TOBTiers       := TTIERSTOB.Create ('le tiers', nil, -1);
     TOBEtab        := TOB.Create ('L etablissement', nil, -1); //$$$JP 28/01/03 modif nom tob
     m_strModeRegl  := '';
end;

destructor TAFFTOB.Destroy;
begin
     TOBPrestation.Free;
//     TOBTache.Free;
     TOBEcheance.Free;
     TOBPiece.Free;
     TOBTiers.Free;
     TOBEtab.Free;

     inherited;
end;

procedure TAFFTOB.LoadDetailAffaire (strAffaire:string; strAffaire2:string; strTiers:string;strCodeGroupe:string;strEtab:string);
var
   Q                :TQuery;
   strReq           :string;
   i                :integer;
   TOBAffReg        :TAFFTOB;
   strNaturePiece   :string;
begin
     Q := nil;
     try
        // Nature des pièces liées à l'affaire ou la proposition
        if strAffaire [1] = 'P' then
            strNaturePiece  := GetParamSoc ('SO_AFNATPROPOSITION')
        else
            strNaturePiece  := GetParamSoc ('SO_AFNATAFFAIRE');

        // Chargement des affaires regroupées (s'il le code regroupement existe)
        if strCodeGroupe <> '' then
        begin
             // Potentiellement, TOUS les champs de l'affaire sont utilisables dans les documents word parser
             Q := OpenSQL ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE<>"' + strAffaire + '" AND AFF_REGROUPEFACT="' + strCodeGroupe + '" AND AFF_TIERS="' + strTiers + '" AND AFF_AFFAIRE2="' + strAffaire2 + '" ORDER BY AFF_AFFAIRE', TRUE);
             if not Q.eof then
                LoadDetailDB ('', '', '', Q, TRUE);
             Ferme (Q);

             for i := 0 to Detail.Count-1 do
             begin
                  TOBAffReg := TAFFTOB.Create ('une affaire', Detail [i], -1);
                  TOBAffReg.LoadDetailAffaire (Detail [i].GetValue ('AFF_AFFAIRE'), Detail [i].GetValue ('AFF_AFFAIRE2'), Detail [i].GetValue ('AFF_TIERS'), '', Detail [i].GetValue ('AFF_ETABLISSEMENT'));
             end;
        end;

        // Lecture de la pièce associé à l'affaire
        strReq := 'SELECT ' + COlePieceChamps + COlePiece2Champs + ' FROM PIECE WHERE GP_AFFAIRE="' + strAffaire + '" AND GP_NATUREPIECEG="' + strNaturePiece + '"';
        Q := OpenSQL (strReq, True);
        if not Q.EOF then
           TOBPiece.LoadDetailDB ('', '', '', Q, TRUE);
        Ferme (Q);

        // Lecture des prestations de la piece de l'affaire
        strReq := 'SELECT ' + COleLigneChamps + COleLigne2Champs + ' FROM LIGNE WHERE GL_AFFAIRE="' + strAffaire + '" AND GL_NATUREPIECEG="' + strNaturePiece + '" ORDER BY GL_NUMLIGNE';
        Q := OpenSQL (strReq, True);
        if not Q.EOF then
           TOBPrestation.LoadDetailDB ('','','', Q, true);
        Ferme (Q);

        // Lecture des échéances de l'affaire (seulement échéances "normales")
        strReq := 'SELECT ' + COleEcheanceChamps + ' FROM FACTAFF WHERE AFA_AFFAIRE="' + strAffaire + '" AND AFA_TYPECHE="NOR" ORDER BY AFA_NUMECHE';
        Q := OpenSQL (strReq, TRUE);
        if not Q.eof then
           TOBEcheance.LoadDetailDB ('', '', '', Q, TRUE);
        Ferme (Q);

        // Lecture des tâches associées à l'affaire
        //$$$todo jp: activer pour version 01-2003, avec vérif' séria
{        if VH_GC.GAPlanningSeria = TRUE then
        begin
             strReq := 'SELECT ATA_FONCTION,ATA_TYPEARTICLE,ATA_ARTICLE,ATA_CODEARTICLE,ATA_PERIODICITE,ATA_INTERVAL,';
             strReq := strReq + 'ATA_DATEDEBPERIOD,ATA_DATEFINPERIOD,ATA_FAMILLETACHE,ATA_NBINTERVENTION,ATA_LIBELLETACHE1,ATA_LIBELLETACHE2,';
             strReq := strReq + 'ATA_MODECALC,ATA_NUMEROTACHE,ATA_POURCENTAGE,ATA_PUVENTEDEVHT,ATA_PUVENTEHT,';
             strReq := strReq + 'ATA_QTEAFFECTE,ATA_QTEAPLANIFIER,ATA_QTEINITIALE,ATA_QTEREALISE,ATA_UNITETEMPS ';
             strReq := strReq + 'FROM TACHE WHERE ATA_AFFAIRE="' + strAffaire + '" ORDER BY ATA_NUMEROTACHE';
             Q := OpenSQL (strReq, TRUE);
             if not Q.eof then
                TOBTache.LoadDetailDB ('', '', '', Q, TRUE);
             Ferme (Q);
        end;}

        // Lecture de l'établissement associé à la mission
        strReq := 'SELECT ' + COleEtabChamps + ' FROM ETABLISS WHERE ET_ETABLISSEMENT="' + strEtab + '"';
        TOBEtab.LoadDetailFromSQL (strReq);

        // Mode de réglement de la pièce
        if TOBPiece.Detail.Count > 0 then
        begin
             Q := OpenSQL ('SELECT MR_LIBELLE FROM MODEREGL WHERE MR_MODEREGLE="' + TOBPiece.Detail [0].GetValue ('GP_MODEREGLE') + '"', TRUE);
             if not Q.eof then
                m_strModeRegl := Trim (Q.FindField ('MR_LIBELLE').AsString);
             Ferme (Q);
        end;

        // Libellé du code 1 du code affaire
        m_strLibelleAff1 := Trim (RechDom('AFFAIREPART1', Copy (strAffaire, 2, 3), FALSE));

        // Tiers concerné
        TOBTiers.LoadDetailTiers (strTiers);
     except
           PgiInfoAf ('Erreur lors du chargement des données', 'Génération document');
           Ferme (Q);
     end;
end;

constructor TTIERSTOB.Create (LeNomTable:string; LeParent:TOB; IndiceFils:integer);
begin
     inherited;

//     TOBTiersComp          := TOB.Create ('complement tiers', nil, -1);
     TOBAnnuaire           := TOB.Create ('Annuaire tiers', nil, -1); //$$$JP 28/01/03 modif nom tob
     TOBDPFiscal           := TOB.Create ('Dp fiscal tiers', nil, -1); //$$$JP 28/01/03 modif nom tob
     TOBDPSocial           := TOB.Create ('Dp social tiers', nil, -1); //$$$JP 28/01/03 modif nom tob
     TOBDPEco              := TOB.Create ('Dp économique tiers', nil, -1); //$$$JP 28/01/03 modif nom tob
     TOBSalarie            := TOB.Create ('Salarie tiers', nil, -1); //$$$JP 28/01/03 modif nom tob
     TOBContact            := TOB.Create ('Les contacts', nil, -1); //$$$JP 28/01/03 modif nom tob
end;

destructor TTIERSTOB.Destroy;
begin
//     TOBTiersComp.Free;
     TOBAnnuaire.Free;
     TOBDPFiscal.Free;
     TOBDPSocial.Free;
     TOBDPEco.Free;
     TOBSalarie.Free;
     TOBContact.Free;

     inherited;
end;

procedure TTIERSTOB.LoadDetailTiers (strTiers:string);
var
   Q          :TQuery;
   strReq     :string;
begin
     Q := nil;
     try
        // Informations tiers
        strReq := 'SELECT ' + COleTiersChamps + COleTiers2Champs + ' FROM TIERS WHERE T_TIERS="' + strTiers + '"';
        Q := OpenSQL (strReq, True);
        if not Q.EOF then
           LoadDetailDB ('','','', Q, true);
        Ferme (Q);

        // Lecture complément tiers concerné
{        Q := OpenSQL ('SELECT * FROM TIERSCOMPL WHERE YTC_TIERS="' + strTiers + '"', True);
        if not Q.eof then
           TOBTiersComp.LoadDetailDB ('','','', Q, true);
        Ferme (Q);}

        // Lecture fiche annuaire associée au tiers
        strReq := 'SELECT ' + COleAnnuaireChamps + COleAnnuaire2Champs + ' FROM ANNUAIRE WHERE ANN_TIERS = "' + strTiers + '"';
        Q := OpenSQL (strReq, TRUE);
        if not Q.eof then
        begin
             TOBAnnuaire.LoadDetailDB ('', '', '', Q, TRUE);
             Ferme (Q);

             // Lecture DP Fiscal associé au tiers (si base commune)
             strReq := 'SELECT ' + COleDPFisChamps + '  FROM DPFISCAL WHERE DFI_NODP = ' + VarAsType (TOBAnnuaire.Detail [0].GetValue ('ANN_CODEPER'), varString);
             Q := OpenSQL (strReq, TRUE);
             if not Q.eof then
                TOBDPFiscal.LoadDetailDB ('', '', '', Q, TRUE);
             Ferme (Q);

             // Lecture DP social associé au tiers (si base commune)
             strReq := 'SELECT ' + COleDPSocChamps + ' FROM DPSOCIAL WHERE DSO_NODP = ' + VarAsType (TOBAnnuaire.Detail [0].GetValue ('ANN_CODEPER'), varString);
             Q := OpenSQL (strReq, TRUE);
             if not Q.eof then
                TOBDPSocial.LoadDetailDB ('', '', '', Q, TRUE);
             Ferme (Q);

             // Lecture DP Eco associé au tiers (si base commune)
             strReq := 'SELECT ' + COleDPEcoChamps + ' FROM DPECO WHERE DEC_NODP = ' + VarAsType (TOBAnnuaire.Detail [0].GetValue ('ANN_CODEPER'), varString);
             Q := OpenSQL (strReq, TRUE);
             if not Q.eof then
                TOBDPEco.LoadDetailDB ('', '', '', Q, TRUE);
             Ferme (Q);
        end
        else
            Ferme (Q);

        // Lecture des contacts du tiers concerné
        strReq := 'SELECT ' + COleContactChamps + ' FROM CONTACT JOIN TIERS ON T_TIERS="' + strTiers + '" AND T_AUXILIAIRE=C_AUXILIAIRE AND C_PRINCIPAL="X" ORDER BY C_NOM';
        Q := OpenSQL (strReq, True);
        if not Q.EOF then
           TOBContact.LoadDetailDB ('','','', Q, true);
        Ferme (Q);

        // Libellé activité
        m_strActivite := Trim (RechDom ('YYCODENAF', Detail [0].GetValue ('T_APE'), FALSE));
        if LowerCase (m_strActivite) = 'error' then
           m_strActivite := '';

        // Libellé de l'abréviation de la forme juridique (pour les courriers)
        m_strAbrevJu := Trim (RechDom('TTFORMEJURIDIQUE', Detail [0].GetValue ('T_JURIDIQUE'), FALSE));
        if LowerCase (m_strAbrevJu) = 'error' then
           m_strAbrevJu := '';

        // Libellé de la (vraie) forme juridique
        Q := OpenSQL ('SELECT CO_LIBRE FROM COMMUN WHERE CO_TYPE="YFJ" AND CO_CODE="' + Detail [0].GetValue ('T_FORMEJURIDIQUE') + '"', FALSE);
        if not Q.eof then
            m_strJuridique := Trim (Q.FindField('CO_LIBRE').AsString)
        else
            m_strJuridique := '';
        Ferme (Q);
     except
            Ferme (Q);
     end;
end;

procedure TOF_GENPARSER.SetCurrentAffaire (iIndex1, iIndex2:integer);
begin
     // Sur la mission principale
     if iIndex1 > -1 then
        if iIndex1 < TOBAff.Detail.Count then
            TOBDataPrincipale := TOBAff.Detail [iIndex1]
        else
            TOBDataPrincipale := TOBAff.Detail [0];
     TOBPrincipale := TAFFTOB (TOBDataPrincipale.Detail [0]);

     // Sur la mission groupée (n'existe pas forcément)
     if iIndex2 > -1 then
     begin
          if iIndex2 < TOBPrincipale.Detail.Count then
              TOBDataGroupee := TOBPrincipale.Detail [iIndex2]
          else
              TOBDataGroupee := TOBPrincipale.Detail [0];
          TOBGroupee := TAFFTOB (TOBDataGroupee.Detail [0]);
     end
     else
     begin
          TOBDataGroupee := nil;
          TOBGroupee     := nil;
     end;

     // Sélection de la tob: celle principale ou celle groupée si existante
     if TOBGroupee = nil then
     begin
          TOBDataAffaire := TOBDataPrincipale;
          TOBAffaire     := TOBPrincipale;
     end
     else
     begin
          TOBDataAffaire := TOBDataGroupee;
          TOBAffaire     := TOBGroupee;
     end;
end;

procedure TOF_GENPARSER.OnArgument (S:string);
var
   Critere     :string;
   i           :integer;
   Q           :TQuery;
   strReq      :string;
   strRegroup  :string;
   CBModele    :THValComboBox;
   TOBUneAff   :TAFFTOB;
   strAffaire  :string;
begin
     inherited;

     // Clé enregistrement maitre inconnu pour l'instant
     bFromFicheAff := FALSE;
     strAffaire    := '';
     strAssistant  := '';
     strNature     := '';
     strEtat       := 'UTI';
     iNoModele     := -1;
     TOBDoc        := TOB.Create ('Documents externes', nil, -1); //$$$JP 28/01/03 modif nom tob
     TOBAff        := nil;

     // Récupération critères
     Critere := (Trim (ReadTokenSt (S)));
     while (Critere <>'') do
     begin
          // TOB des affaires (à charger)
          if (copy (Critere,1,10) = 'TOBAFFAIRE') then
             TOBAff := TOB (strtoint (Copy (Critere, 12, Length (Critere)-11)));

          // Enregistrement affaire maitre
          if (copy (Critere,1,7) = 'AFFAIRE') then
          begin
               strAffaire := Copy (Critere, 9, Length (Critere)-8);
               bFromFicheAff := TRUE;
          end;

          // Etat des modèles à proposer: utilisable ou bien en cours d'élaboration
          if (copy (Critere,1,4) = 'ETAT') then
             strEtat := Copy (Critere, 6, Length (Critere)-5);

          // Nature de document parser (lettre mission, mailing client, mailing assistant)
          if (copy (Critere, 1, 6) = 'NATURE') then
             strNature := Copy (Critere, 8, Length (Critere)-7);

          // enregistrement assistant maitre
          if (copy (Critere,1,9) = 'ASSISTANT') then
             strAssistant := Copy (Critere, 11, Length (Critere)-10);

          // Si modèle déjà choisi
          if (copy (Critere, 1, 6) = 'MODELE') then
             if Length (Critere) > 7 then
                iNoModele := StrToInt (Copy (Critere, 8, Length (Critere)-7));

          // Titre de la fenêtre
          if (copy (Critere, 1, 5) = 'TITRE') then
             Ecran.Caption := Copy (Critere, 7, Length (Critere)-6);

          // Paramètre suivant
          Critere := (Trim (ReadTokenSt (S)));
     end;

     // Liste des modèle utilisables (table AFDOCEXTERNE)
     CBModele := THValComboBox (GetControl ('CBMODELE'));
     strReq := 'SELECT ADE_DOCEXCODE, ADE_LIBELLE, ADE_FICHIER, ADE_DOSSIER FROM AFDOCEXTERNE WHERE ADE_DOCEXTYPE="WOR" ';
     if strEtat = 'UTI' then
         strReq := strReq + 'AND ADE_DOCEXETAT="UTI" ';
     strReq := strReq + 'AND ADE_DOCEXNATURE="' + strNature + '" ';
        //mcd 27/04/04 on prend tout même si base uniqueif V_PGI_ENV.ModeFonc <> 'MONO' then
     strReq := strReq + 'AND (ADE_DOSSIER="$STD" OR ADE_DOSSIER="$DAT" OR ADE_DOSSIER="' + V_PGI.NoDossier + '") ';
     strReq := strReq + 'ORDER BY ADE_DOCEXCODE';
     Q := OpenSQL (strReq, TRUE);
     if not Q.eof then
        TOBDoc.LoadDetailDB ('', '', '', Q, TRUE);
     Ferme (Q);
     for i := 0 to TOBDoc.Detail.Count-1 do
     begin
          CBModele.Items.AddObject (TOBDoc.Detail [i].GetValue ('ADE_LIBELLE'), TOBDoc.Detail [i]);
          if (iNoModele <> -1) and (iNoModele = TOBDoc.Detail [i].GetValue ('ADE_DOCEXCODE')) then
          begin
               CBModele.ItemIndex := i;
               THEdit (GetControl ('EPATHRESULT')).SetFocus;
          end;
     end;

     // Si nature "en cours d'élaboration", c'est le mode test: pas de rép à demander, ni d'option "écraser", ...
     if strEtat = 'ELA' then
     begin
          THEdit (GetControl ('EPATHRESULT')).Text    := 'Fichier temporaire';
          THEdit (GetControl ('EPATHRESULT')).Enabled := FALSE;
          TCheckBox(GetControl ('CBECRASE')).Visible  := FALSE;
     end
     else
     begin
          THEdit (GetControl ('EPATHRESULT')).Text    := '';
          TCheckBox(GetControl ('CBECRASE')).Checked  := TRUE;
     end;

     // Si une seule affaire, on créer la tob (sinon, elle existe déjà
     if TOBAff = nil then
     begin
          // Potentiellement, TOUS les champs de l'affaire sont utilisables dans les documents word parser
          TOBAff := TOB.Create ('Les affaires', nil, -1); //$$$JP 28/01/03 modif nom tob
          TOBAff.LoadDetailFromSQL ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="' + strAffaire + '"');
     end;

     // Chargement des données liées à chaque affaire principale (affaires liées, tiers, prestations...)
     for i := 0 to TOBAff.Detail.Count-1 do
     begin
          // Détermination si on doit charger les affaires groupées
          strRegroup := TOBAff.Detail [i].GetValue ('AFF_REGROUPEFACT');
          if (TOBAff.Detail [i].GetValue ('AFF_PRINCIPALE') <> 'X') or (strRegroup = 'AUC') or (strRegroup = '') then
             strRegroup := '';

          // Création tobaff (contient tiers, ...)
          TOBUneAff := TAFFTOB.Create ('une affaire', TOBAff.Detail [i], -1);
          TOBUneAff.LoadDetailAffaire (TOBAff.Detail [i].GetValue ('AFF_AFFAIRE'), TOBAff.Detail [i].GetValue ('AFF_AFFAIRE2'), TOBAff.Detail [i].GetValue ('AFF_TIERS'), strRegroup, TOBAff.Detail [i].GetValue ('AFF_ETABLISSEMENT'));
     end;
end;

procedure TOF_GENPARSER.OnImprime ();
var
   sNomFic     :string;
   sOrigine    :string;
   strPathName :string;
   strFullName :string;
   bExist      :boolean;
   i           :integer;
   strAffaire  :string;
   DocListe    :TStringList;
   MonWord     :OleVariant;
   bIsNew      :boolean;
   bOuvreWord  :boolean;
// $$$ JP 07/07/03   AppliPathStd, AppliPathDat   :string;
begin
     // Récupération du modèle de document à générer
     i := THValComboBox (GetControl ('CBMODELE')).ItemIndex;
     if i = -1 then
     begin
          PgiInfoAf ('Veuillez sélectionner un modèle', Ecran.Caption);
          exit;
     end;

     // Nom du fichier modèle
     sNomFic  := TOBDoc.Detail [i].GetValue ('ADE_FICHIER');
     if sNomFic = '' then
     begin
          PgiInfoAf ('Aucun fichier associé au document "' + TOBDoc.Detail [i].GetValue ('ADE_LIBELLE') + '"', Ecran.Caption);
          exit;
     end;

     // Sélection du répertoire où chercher le fichier modèle (std, dat ou dossier)
     sOrigine := TOBDoc.Detail [i].GetValue ('ADE_DOSSIER');
     if sOrigine = V_PGI.NoDossier then
         sNomFic := ChangeStdDatPath ('$DOS' + '\' + sNomFic)
     else
     begin
          // $$$ JP 07/07/03: amélioration identification STD ET DAT
          if sOrigine = '$STD' then
              sNomFic := GetAppliPathStd + '\' + sNomFic
          else
              if sOrigine = '$DAT' then
                 sNomFic := GetAppliPathDat + '\' + sNomFic;
          // $$$ JP FIN
     end;


     // Test existence du fichier modèle
     if FileExists (sNomFic) = FALSE then
     begin
          PgiInfoAf ('Le modèle "' + sNomFic + '" n''existe pas ou n''est pas accessible', Ecran.Caption);
          exit;
     end;

     // En mode élaboration, on visualise et génère un fichier temporaire
     if strEtat = 'ELA' then
     begin
          strFullName := FileTemp ('.doc');
          SetCurrentAffaire (0, -1);
          ConvertDocFile (sNomFic, strFullName, ParseInitialize, ParseFinalize, ParseFunction, ParseGetVar, ParseSetVar, ParseGetList);
          OpenDoc (strFullName, wdWindowStateMaximize, wdPrintView);
     end
     else
     begin
          strPathName := THEdit (GetControl ('EPATHRESULT')).Text;
          if strPathName <> '' then
          begin
               if DirectoryExists (strPathName) = FALSE then
               begin
                    PgiInfoAf ('Le répertoire: "' + strPathName + '" n''existe pas ou n''est pas accessible.', Ecran.Caption);
                    exit;
               end;
          end
          else
          begin
               strPathName := ExtractFilePath (sNomFic);
               PgiInfoAf ('Vous n''avez pas renseigné le répertoire destination.' + #10 + ' Le document sera constitué dans ' + strPathName, Ecran.Caption);
          end;
          if strPathName [Length (strPathName)] <> '\' then
             strPathName := strPathName + '\';

          // Boucle sur les affaires, sinon génération pour l'unique affaire spécifiée
          if TOBAff.Detail.Count > 0 then
          begin
               DocListe := TStringList.Create;
               for i := 0 to TOBAff.Detail.Count-1 do
               begin
                    // Code affaire à traiter
                    strAffaire := TOBAff.Detail [i].GetValue ('AFF_AFFAIRE');

                    // Nom du fichier à générer: indexé par code affaire
                    strFullName := strPathName + strNature + '_' + TOBAff.Detail [i].GetValue ('AFF_TIERS') + '_' + strAffaire; // + '_' + ExtractFileName (sNomFic);
                    bExist := FileExists (strFullName);
                    if (TCheckBox(GetControl ('CBECRASE')).Checked = TRUE) or (bExist = FALSE) then
                    begin
                         SetCurrentAffaire (i, -1);
                         ConvertDocFile (sNomFic, strFullName, ParseInitialize, ParseFinalize, ParseFunction, ParseGetVar, ParseSetVar, ParseGetList);
                         DocListe.Add (strFullName);

                         // Mise à jour des événements (si codeper associé)
                         //$$$todo: à activer dès que possible, voir avec Jury/DP
{                         if TAFFTOB (TOBAff.Detail [i].Detail [0]).TOBTiers.TOBAnnuaire.Detail.Count > 0 then
                         begin
                              iMaxNoEvt := -1;
                              Q := OpenSQL ('SELECT MAX(JEV_NOEVT) AS MAXNOEVT FROM JUEVENEMENT', TRUE);
                              if not Q.eof then
                                 iMaxNoEvt := Q.FindField ('MAXNOEVT').AsInteger + 1;
                              Ferme (Q);
                              strReq := 'INSERT INTO JUEVENEMENT (JEV_DATECREATION,JEV_DATEMODIF,JEV_UTILISATEUR,JEV_NOEVT,JEV_CODEEVT,JEV_USER1,JEV_EVTLIBELLE,JEV_EVTLIBABREGE,JEV_FAMEVT,JEV_CODEPER,JEV_DOCNOM,JEV_DATE) ';
                              strReq := strReq + 'VALUES ("' + UsDateTime (V_PGI.DateEntree) + '","' + UsDateTime (V_PGI.DateEntree) + '","' + V_PGI.User + '",' + IntToStr (iMaxNoEvt) + ',"LETTMISS","' + V_PGI.User + '",';
                              strReq := strReq + '"' + ExtractFileName (strFullName) + '","Document externe",';
                              strReq := strReq + '"DOC",' + IntToStr (TAFFTOB (TOBAff.Detail [i].Detail [0]).TOBTiers.TOBAnnuaire.Detail [0].GetValue ('ANN_CODEPER')) + ',"' + strFullName + '","' + UsDateTime (V_PGI.DateEntree) + '")';
                              ExecuteSQL (strReq);
                         end;}
                    end
                    else
                        if bExist = TRUE then
                           PgiInfoAf ('Le document ' + strFullName  + ' est déjà présent dans ' + strPathName, Ecran.Caption);
               end;

               // Si un seul fichier généré, on l'ouvre, sinon on lance l'impression (si utilisateur ok)
               if TOBAff.Detail.Count = 1 then
                   OpenDoc (DocListe [0], wdWindowStateMaximize, wdPrintView)
               else
               begin
                    bOuvreWord := TRUE;
                    if DocListe.Count > 8 then
                       if PgiAskAF (IntToStr (DocListe.Count) + ' lettres ont été constituées. Désirez-vous TOUTES les ouvrir dans WORD?', Ecran.Caption) <> mrYes then
                          bOuvreWord := FALSE;
                    if bOuvreWord = TRUE then
                    begin
                         OpenWord (FALSE, MonWord, bIsNew);
                         if VarIsEmpty (MonWord) = TRUE then
                            PgiInfo ('Connexion à Microsoft Word impossible', Ecran.Caption)
                         else
                         begin
                              try
                                // InitMove (DocListe.Count, '');
                                 for i := 0 to DocListe.Count-1 do
                                 begin
                                      MonWord.Documents.Open (DocListe [i]);
                                      MonWord.ActiveDocument.ActiveWindow.View.Type := 3;
                                  //    MoveCur (FALSE);
                                 end;
                              finally
                                     MonWord := unAssigned;
                                    // FiniMove;
                              end;
                         end;
                    end;
               end;
               DocListe.Free;
          end;
     end;
end;

procedure TOF_GENPARSER.ParseGetVar (Sender: TObject; VarName: String; VarIndx: Integer; var Value: Variant);
var
   iIndex      :integer;
begin
     // Valeur vide par défaut
     Value := '';

     // Si pas d'index, index 0 par défaut
     if VarIndx = -1 then
        VarIndx := 0;

     // Identification du champ (table, code...) et renvoi la valeur
     try
        // Mode de description des missions
        if VarName = 'MODEDESCRIPTION' then
        begin
             iIndex := THRadioGroup (GetControl ('RGMODEDESCR')).ItemIndex;
             case iIndex of
                  0:
                    Value := 'DES';
                  1:
                    Value := 'PRE';
                  2:
                    Value := 'TAC';
             else
                    Value := 'DES';
             end;
        end

        // Paramètres société
        else if (Copy (VarName, 1, 4) = 'SCO_') or (Copy (VarName, 1, 3) = 'SO_') then
             Value := GetParamSoc (VarName)
        else if Copy (VarName, 1, 3) = 'ET_' then
             Value := TOBAffaire.TOBEtab.Detail [0].GetValue (VarName)
        else if Copy (VarName, 1, 2) = 'C_' then
        begin
             if TOBAffaire.TOBTiers.TOBContact.Detail.Count > VarIndx then
                Value := TOBAffaire.TOBTiers.TOBContact.Detail [VarIndx].GetValue (VarName);
             if VarName = 'C_CIVILITE' then
             begin
                  Value := Trim (RechDom ('TTCIVILITE', Value, FALSE));
                  if LowerCase (Value) = 'error' then
                     Value := '';
             end;
        end
        else if Copy (VarName, 1, 2) = 'T_' then
             begin
                  if VarName = 'T_APE' then
                      Value := TOBAffaire.TOBTiers.m_strActivite
                  else if VarName = 'T_JURIDIQUE' then
                      Value := TOBAffaire.TOBTiers.m_strAbrevJu
                  else if VarName = 'T_FORMEJURIDIQUE' then
                      Value := TOBAffaire.TOBTiers.m_strJuridique
                  else if TOBAffaire.TOBTiers.Detail.Count > VarIndx then
                      Value := TOBAffaire.TOBTiers.Detail [VarIndx].GetValue (VarName);
             end
//        else if Copy (VarName, 1, 4) = 'YTC_' then
  //           begin
    //              if TOBAffaire.TOBTiers.TOBTiersComp.Detail.Count > VarIndx then
      //               Value := TOBAffaire.TOBTiers.TOBTiersComp.Detail [VarIndx].GetValue (VarName);
        //     end
        else if Copy (VarName, 1, 4) = 'AFF_' then
             begin
                  if VarName = 'AFF_AFFAIRE1' then
                      Value := TOBAffaire.m_strLibelleAff1
                  else
                      if VarName = 'AFF_DESCRIPTIF' then
                          Value :=  GetBlobFusion (TOBDataAffaire.GetValue (VarName), 'S')
                      else
                          Value := TOBDataAffaire.GetValue (VarName)
             end
        else if Copy (VarName, 1, 4) = 'ANN_' then
             begin
                  if TOBAffaire.TOBTiers.TOBAnnuaire.Detail.Count > VarIndx then
                     Value := TOBAffaire.TOBTiers.TOBAnnuaire.Detail [VarIndx].GetValue (VarName);
             end
        else if Copy (VarName, 1, 3) = 'GP_' then
             begin
                  if VarName = 'GP_MODEREGLE' then
                      Value := TOBAffaire.m_strModeRegl
                  else
                      if TOBAffaire.TOBPiece.Detail.Count > VarIndx then
                         Value := TOBAffaire.TOBPiece.Detail [VarIndx].GetValue (VarName);
             end
        else if Copy (VarName, 1, 4) = 'AFA_' then
             begin
                  if TOBAffaire.TOBEcheance.Detail.Count > VarIndx then
                    Value := TOBAffaire.TOBPiece.Detail [VarIndx].GetValue (VarName);
              end
        else if Copy (VarName, 1, 3) = 'GL_' then
             begin
                  if TOBAffaire.TOBPrestation.Detail.Count > VarIndx then
                  begin
                      if VarName = 'GL_BLOCNOTE' then   //AB-20031107
                         Value :=  GetBlobFusion (TOBAffaire.TOBPrestation.Detail [VarIndx].GetValue (VarName), 'S')
                      else
                         Value := TOBAffaire.TOBPrestation.Detail [VarIndx].GetValue (VarName);
                  end;
             end
        //$$$todo jp: activer pour version 01-2003
//        else if Copy (VarName, 1, 4) = 'ATA_' then
  //           begin
    //              if TOBAffaire.TOBTache.Detail.Count > VarIndx then
      //               Value := TOBAffaire.TOBTache.Detail [VarIndx].GetValue (VarName);
        //     end
        else if Copy (VarName, 1, 4) = 'DSO_' then
             begin
                  if TOBAffaire.TOBTiers.TOBDPSocial.Detail.Count > VarIndx then
                     Value := TOBAffaire.TOBTiers.TOBDPSocial.Detail [VarIndx].GetValue (VarName);
             end
        else if Copy (VarName, 1, 4) = 'DFI_' then
             begin
                  if TOBAffaire.TOBTiers.TOBDPFiscal.Detail.Count > VarIndx then
                     Value := TOBAffaire.TOBTiers.TOBDPFiscal.Detail [VarIndx].GetValue (VarName);
             end
        else if Copy (VarName, 1, 4) = 'DEC_' then
             begin
                  if TOBAffaire.TOBTiers.TOBDPEco.Detail.Count > VarIndx then
                     Value := TOBAffaire.TOBTiers.TOBDPEco.Detail [VarIndx].GetValue (VarName);
             end;
{         else if Copy (VarName, 1, 4) = 'RUB_' then
         begin
              TOBRubFille := TOBRubrique.FindFirst (['ARL_RUBRIQUE'],[Copy (varName, 5, Length (VarName)-4)], TRUE);
              if TOBRubFille <> Nil then
                 Value := TOBRubFille.GetValue ('ARL_VALEUR');
         end;}
     finally
            if VarIsNull (Value) = TRUE then
               Value := '';
            if ((VarType (Value) and varTypeMask) = varString) or ((VarType (Value) and varTypeMask) = varOleStr) then
               if Value = '' then
                  Value := ' ';
     end;
end;

procedure TOF_GENPARSER.ParseSetVar (Sender: TObject; VarName: String; VarIndx: Integer; Value: Variant);
//var
  // TOBRubFille   :TOB;
begin
     try
         // Mise à jour possible uniquement sur les rubriques de lettre de mission
{         if Copy (VarName, 1, 4) = 'RUB_' then
         begin
              TOBRubFille := TOBRubrique.FindFirst (['ARL_RUBRIQUE'],[Copy (VarName, 5, Length (VarName)-4)], TRUE);
              if TOBRubFille = Nil then
              begin
                   TOBRubFille := TOB.Create ('AFRUBLETTREMISS', TOBRubrique, -1);
                   TOBRubFille.PutValue ('ARL_AFFAIRE', strAffaire);
                   TOBRubFille.PutValue ('ARL_TIERS', strTiers);
                   TOBRubFille.PutValue ('ARL_RUBRIQUE', Copy (VarName, 5, Length (VarName)-4));
                   TOBRubFille.PutValue ('ARL_VOLATILE', '-');
              end;
              TOBRubFille.PutValue ('ARL_VALEUR', VarAsType(Value, varString));
         end;}
     finally
     end;
end;

procedure TOF_GENPARSER.ParseGetList (Sender: TObject; IdentListe: String; var Liste: TStringList);
begin
end;

procedure TOF_GENPARSER.ParseFunction (Sender: TObject; FuncName: String; Params: Array of variant ; var Result: Variant);
var
   strTable    :string;
// $$$ JP 07/07/03  AppliPathStd, AppliPathDat   :string;
begin
     // $$$ JP 07/07/03: amélioration identification STD ET DAT
//{$IFNDEF EAGLCLIENT}
  //       if V_PGI_ENV <> nil then
    //     begin
//{$IFDEF GIGI}
  //            AppliPathStd := V_PGI_ENV.PathStd + '\GIS5';
    //          AppliPathDat := V_PGI_ENV.PathDat + '\GIS5';
//{$ELSE}
  //            AppliPathStd := V_PGI_ENV.PathStd + '\GAS5';
    //          AppliPathDat := V_PGI_ENV.PathDat + '\GAS5';
//{$ENDIF}
//         end
  //       else
    //     begin
//              // V_PGI_ENV inexistant, pas normal!
  //            AppliPathStd := 'C:\PGI00\STD';
    //          AppliPathDat := 'C:\PGI01\STD';
      //   end;
//{$ENDIF}

     Result := '';
//     if FuncName = '@AFFAIREPRINCIPALE' then
  //   begin
    //      if VarType (Params [1]) = varInteger then
      //       SetCurrentAffaire (Params [1], -1);
//     end
     if FuncName = '@AFFAIREGROUPEE' then
     begin
          if VarType (Params [1]) = varInteger then
             SetCurrentAffaire (-1, Params [1]);
     end
     else if FuncName = '@DATEENTREE' then
          Result := V_PGI.DateEntree
     else if FuncName = '@PATHSTD' then
          Result := GetAppliPathStd + '\'          // $$$ JP 07/07/03: amélioration identification STD ET DAT
     else if FuncName = '@PATHDAT' then
          Result := GetAppliPathDat + '\'          // $$$ JP 07/07/03: amélioration identification STD ET DAT
     else if FuncName = '@PATHDOSS' then
          Result := ChangeStdDatPath('$DOS\')
     else if FuncName = '@NOMEXPERT' then
          begin
               Result := THEdit (GetControl ('NOMEXPERT')).Text;
               if Result = '' then
                  Result := ' ';
          end
     else if FuncName = '@NBELEMENT' then
     begin
          // Table spécifiée
          if VarType (Params [1]) = varString then
             strTable := Params [1];

          // Renvoi le nb d'élément de la table demandée, pour l'affaire et l'affaire fille spécifiée
          if strTable = 'AFFAIREPRINCIPALE' then
             Result := TOBDataPrincipale.Detail.Count
          else if strTable = 'AFFAIREGROUPEE' then
                  if TCheckBox (GetControl ('CBMULTI')).Checked = FALSE then
                      Result := 0
                  else
                      Result := TOBPrincipale.Detail.Count
          else if strTable = 'PRESTATION' then
               Result := TOBAffaire.TOBPrestation.Detail.Count
          // $$$todo: activer pour version 01-2003
//          else if strTable = 'TACHE' then
  //             Result := TOBAffaire.TOBTache.Detail.Count
          else if strTable = 'ECHEANCE' then
               Result := TOBAffaire.TOBEcheance.Detail.Count
          else if strTable = 'CONTACT' then
               Result := TOBAffaire.TOBTiers.TOBContact.Detail.Count
          else if strTable = 'ANNUAIRE' then
               Result := TOBAffaire.TOBTiers.TOBAnnuaire.Detail.Count
          else if strTable = 'DPFISCAL' then
               Result := TOBAffaire.TOBTiers.TOBDPFiscal.Detail.Count
          else if strTable = 'DPSOCIAL' then
               Result := TOBAffaire.TOBTiers.TOBDPSocial.Detail.Count
          else if strTable = 'DPECO' then
               Result := TOBAffaire.TOBTiers.TOBDPEco.Detail.Count;
        end;
end;

procedure TOF_GENPARSER.ParseInitialize (Sender: TObject);
begin
end;

procedure TOF_GENPARSER.ParseFinalize (Sender: TObject);
begin
end;


procedure GenereDocParser (params: array of variant; nb:integer);
var
   F         :TForm;
   maTOF     :TOF;
begin
     F := TForm (Longint(Params[0]));
     if (F is TFVierge) then maTOF := TFVierge(F).laTOF else exit;
     if (maTOF is TOF_GENPARSER) then TOF_GENPARSER (maTOF).OnImprime () else exit;
end;


procedure TOF_GENPARSER.OnClose;
begin
     TOBDoc.Free;
     if bFromFicheAff = TRUE then
        TOBAff.Free;

     inherited;
end;


procedure AFLanceFiche_LanceParser (Range, Argument:string);
begin
     AGLLanceFiche ('AFF', 'AFLANCEPARSER', Range, '', Argument);
end;



Initialization
  registerclasses ( [ TOF_GENPARSER ] ) ;
  RegisterAglProc ('GenereLettreMission', TRUE, 0, GenereDocParser);
end.

