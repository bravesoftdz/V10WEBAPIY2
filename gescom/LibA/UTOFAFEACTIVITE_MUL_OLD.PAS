{***********UNITE*************************************************
Auteur  ...... : PL
Créé le ...... : 31/10/2000
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : AFEACTIVITE_MUL ()
Mots clefs ... : TOF;AFEACTIVITE_MUL
*****************************************************************}
Unit UTOFAFEACTIVITE_MUL ;

Interface

Uses StdCtrls, Controls, Classes,  forms, sysutils,  ComCtrls, HTB97, utob, mul,  windows,
     HCtrls, Ent1, HEnt1, HMsgBox, UTOF, UTofAfBaseCodeAffaire, uafo_ressource,
     AffaireUtil,AfUtilArticle, ActiviteUtil,
     FactUtil, UtilArticle, DicoAF,Entgc, heureutil, HQry, Hxlspas, FileCtrl, paramsoc
{$IFDEF EAGLCLIENT}
     ,MaineAGL
{$ELSE}
     ,DB, DBTables,FE_Main,HDB ,dbgrids ,Shellapi
{$ENDIF}
     ,UtobXLS, UtilRessource,traducaffaire,UtofAfEActiviteModif
     ;

Type
  TOF_AFEACTIVITE_MUL = Class (TOF_AFBASECODEAFFAIRE)
    procedure Onupdate; override;
    procedure OnArgument (stArgument : String ) ;override;

    procedure BOuvrirOnClick(Sender: TObject);
    procedure BImportOnClick(Sender: TObject);
    procedure BEffaceOnClick(Sender: TObject);
    procedure BEffaceFicLogOnClick(Sender: TObject);
    procedure BFicLogOnClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);

//    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;

    private
           TobImport                  :TOB;
           FichierLog                 :TextFile;
           EditDateDebut              :THEdit;
           EditDateFin                :THEdit;
           gDateDebutAct, gDateFinAct :TDateTime;
           bSuivi                     :boolean;

           function ImportEActivite(TobEActivite:TOB;bAll:boolean):integer;
           function PutClipboardGrille(laTOB : TOB ) : variant;
           function CompteEActivite:integer;
           function NomFichierLog:string;
    end;

const
	// libellés des messages
	TexteMessage: array[1..2] of string 	= (
          {1}        ' lignes d''activité ont été importées et '
          {2}       ,' ont été refusées.'
                       );
Procedure AFLanceFiche_Mul_EActivite;

implementation


procedure TOF_AFEACTIVITE_MUL.OnArgument(stArgument : String ) ;
var
   ComboTypeArticle :THMultiValComboBox;
   DateDeb          :string;
   DateFin          :string;
   Critere          :string;
   Assistant        :string;
   uneEdit          :THEdit;
   iPos             :integer;
//   tt               :TToolBarButton97;
begin
     Inherited ;

     // Contrôles de la fiche
     TToolBarButton97 (GetControl('BOuvrir')).OnClick       := BOuvrirOnClick;
     TToolBarButton97 (GetControl('BIMPORTASC')).OnClick    := BImportOnClick;
     TToolBarButton97 (GetControl('BIMPORTEXCEL')).OnClick  := BImportOnClick;
     TToolBarButton97 (GetControl('BEFFACETABLE')).OnClick  := BEffaceOnClick;
     TToolBarButton97 (GetControl('BFICLOG')).OnClick       := BFicLogOnClick;
     TToolBarButton97 (GetControl('BEFFACEFICLOG')).OnClick := BEffaceFicLogOnClick;
     THDBGrid (TFMUL(Ecran).FListe).OnDblClick              := FListeDblClick;

     // Paramètres: date période
     DateDeb   := '';
     DateFin   := '';
     bSuivi    := FALSE;
     Critere := (Trim (ReadTokenSt (stArgument)));
     while (Critere <>'') do
     begin
          // Mode suivi eactivité débrayé ou non
          if (copy (Critere,1,9) = 'MODESUIVI') then
             bSuivi := TRUE;

          // Date début
          if (copy (Critere,1,7) = 'DATEDEB') then
             DateDeb := Copy (Critere, 9, 10);

          // Date fin
          if (copy (Critere,1,7) = 'DATEFIN') then
             DateFin := Copy (Critere, 9, 10);

          // Code assistant (si existant, sinon TOUS les assistants)
          if (copy (Critere, 1, 9) = 'ASSISTANT') then
             Assistant := Copy (Critere, 11, Length (Critere)-10);

          // Paramètre suivant
          Critere := (Trim (ReadTokenSt (stArgument)));
     end;

     // Gestion des dates d'activité (si paramètres passés, ceux-ci priment sur les dates activité)
     EditDateDebut := THEdit (GetControl('EAC_DATEACTIVITE'));
     EditDateFin   := THEdit (GetControl('EAC_DATEACTIVITE_'));
     IntervalleDatesActivite (gDateDebutAct, gDateFinAct);
     if EditDateDebut <> nil then
        if DateDeb <> '' then
            EditDateDebut.Text := DateDeb
        else
            EditDateDebut.Text := datetostr (gDateDebutAct);
     if EditDateFin <> nil then
        if DateFin <> '' then
            EditDateFin.Text := DateFin
        else
            EditDateFin.Text := datetostr (gDateFinAct);

     // Filtre sur l'assistant si spécifié
     if (Assistant <> '') and (Assistant <> '*') then
     begin
          uneEdit := THEdit (GetControl('EAC_RESSOURCE'));
          uneEdit.Text := Assistant;
          uneEdit.ElipsisButton := False;
          uneEdit.ReadOnly := TRUE;
          uneEdit := THEdit (GetControl('EAC_RESSOURCE_'));
          uneEdit.Text := Assistant;
          uneEdit.ReadOnly := TRUE;
          uneEdit.ElipsisButton := False;
     end;

     // JP - gestion suivi ou pas: voir les lignes SDE ou non 
     TToolBarButton97 (GetControl('BIMPORTASC')).Visible   := not bSuivi;
     TToolBarButton97 (GetControl('BIMPORTEXCEL')).Visible := not bSuivi;
     if bSuivi = FALSE then
          THValComboBox (GetControl ('EAC_ACTORIGINE')).Plus := 'AND CC_CODE <> "SDE"'
     else
     begin
          THValComboBox (GetControl ('EAC_ACTORIGINE')).Plus := 'AND CC_CODE = "SDE"';
          THValComboBox (GetControl ('EAC_ACTORIGINE')).Value := 'SDE';
          THValComboBox (GetControl ('EAC_ACTORIGINE')).Enabled := FALSE;
     end;


     // mcd 05/03/02
     ComboTypeArticle:=THMultiValComboBox(GetControl('EAC_TYPEARTICLE'));
     ComboTypeArticle.plus := PlusTypeArticle;
     if ComboTypeArticle.Text='' then
        ComboTypeArticle.Text := PlusTypeArticleText;
end ;

procedure TOF_AFEACTIVITE_MUL.Onupdate;
Var
   bDateDebOK, bDateFinOK,bCherche:boolean;
begin
Inherited;
bCherche := false;
{$IFDEF EAGLCLIENT}
TraduitAFLibGridSt(TFMul(Ecran).FListe);
{$ELSE}
TraduitAFLibGridDB(TFMul(Ecran).FListe);
{$ENDIF}

bDateDebOK:=true; bDateFinOK:=true;
if (EditDateDebut.Text='') then bDateDebOK:=false;
if (bDateDebOK and (strtodate(EditDateDebut.Text)<gDateDebutAct) ) then bDateDebOK:=false;
if (bDateDebOK=false) then
    begin
    PGIInfoAf('La date de début est antérieure à la date de début d''activité', Ecran.Caption);
    SetFocusControl('EAC_DATEACTIVITE');
    EditDateDebut.Text:=datetostr(gDateDebutAct);
    bCherche := true;
    end;

if (EditDateFin.Text='') then bDateFinOK:=false;
if (bDateFinOK and ( (strtodate(EditDateFin.Text)>gDateFinAct))) then bDateFinOK:=false;
if (bDateFinOK=false) then
    begin
    PGIInfoAf('La date de fin est postérieure à la date de fin d''activité', Ecran.Caption);
    SetFocusControl('EAC_DATEACTIVITE_');
    EditDateFin.Text:=datetostr(gDateFinAct);
    bCherche := true;
    end;
if bCherche then TFMul(Ecran).ChercheClick ;
end;


procedure TOF_AFEACTIVITE_MUL.FListeDblClick(Sender: TObject);
var
   TobRecup, TobEnreg, TOBL, TOBArt :TOB;
   CleEntree, CleSortie, Concat:string;
   eDate,eRessource,eAffaire,eAff0,eAff1,eAff2,eAff3,eAff4,eArticle,eTiers:string;
   csTypAct, csAffaire, csRessource, csTypArt:string;
   cdDate:TDateTime;
   ciNumLig:integer;
   sDate,sRessource,sAffaire,sArticle,sTiers:string;
   Reponse, Champ, valeur,sTypeArticle  : String;
   x : integer;
   Q:TQuery;
   rechart : T_RechArt;
begin
     sTypeArticle   := '';
     CleEntree      := '';
     CleSortie      := '';

     csTypArt:=THDBGrid(TFMUL(Ecran).FListe).datasource.dataset.FindField('EAC_TYPEARTICLE').Asstring;
     csRessource:=THDBGrid(TFMUL(Ecran).FListe).datasource.dataset.FindField('EAC_RESSOURCE').Asstring;
     csAffaire:=THDBGrid(TFMUL(Ecran).FListe).datasource.dataset.FindField('EAC_AFFAIRE').Asstring;
     csTypAct:=THDBGrid(TFMUL(Ecran).FListe).datasource.dataset.FindField('EAC_TYPEACTIVITE').Asstring;
     ciNumLig:=THDBGrid(TFMUL(Ecran).FListe).datasource.dataset.FindField('EAC_NUMLIGNE').AsInteger;
     cdDate:=THDBGrid(TFMUL(Ecran).FListe).datasource.dataset.FindField('EAC_DATEACTIVITE').AsDateTime;
     eDate:=THDBGrid(TFMUL(Ecran).FListe).datasource.dataset.FindField('EAC_DATEACTIVITE').AsString;
     eRessource:=THDBGrid(TFMUL(Ecran).FListe).datasource.dataset.FindField('EAC_RESSOURCE').AsString;
     eAffaire:=THDBGrid(TFMUL(Ecran).FListe).datasource.dataset.FindField('EAC_AFFAIRE').AsString;
     eAff0:=THDBGrid(TFMUL(Ecran).FListe).datasource.dataset.FindField('EAC_AFFAIRE0').AsString;
     eAff1:=THDBGrid(TFMUL(Ecran).FListe).datasource.dataset.FindField('EAC_AFFAIRE1').AsString;
     eAff2:=THDBGrid(TFMUL(Ecran).FListe).datasource.dataset.FindField('EAC_AFFAIRE2').AsString;
     eAff3:=THDBGrid(TFMUL(Ecran).FListe).datasource.dataset.FindField('EAC_AFFAIRE3').AsString;
     eAff4:=THDBGrid(TFMUL(Ecran).FListe).datasource.dataset.FindField('EAC_AVENANT').AsString;
     eArticle:=THDBGrid(TFMUL(Ecran).FListe).datasource.dataset.FindField('EAC_CODEARTICLE').AsString;
     eTiers:=THDBGrid(TFMUL(Ecran).FListe).datasource.dataset.FindField('EAC_TIERS').AsString;

     // JP: si ligne sélectionné vide, on fait rien
     if eDate = '' then
        exit;

     CleEntree :='DATE:'+eDate+';RESS:'+eRessource+';AFF:'+eAffaire+';AFF0:'+eAff0+';AFF1:'+eAff1+';AFF2:'+eAff2+';AFF3:'+eAff3+';AFF4:'+eAff4+';ART:'+eArticle+';TIE:'+eTiers;
     CleSortie := AFLanceFiche_EActiviteModif(CleEntree);

     Concat := CleSortie;
     Reponse:=(Trim(ReadTokenSt(Concat)));
     while (Reponse <>'') do
     BEGIN
        if Reponse<>'' then
        BEGIN
        X:=pos(':',Reponse);
        if x<>0 then
           begin
           Champ:=copy(Reponse,1,X-1);
           Valeur:=Copy(Reponse,X+1,length(Reponse)-X);
           end;
        if Champ = 'DATE' then sDate := Valeur;
        if Champ = 'RESS' then sRessource := Valeur;
        if Champ = 'AFF' then sAffaire := Valeur;
        if Champ = 'ART' then sArticle := Valeur;
        if Champ = 'TIE' then sTiers := Valeur;
        END;
        Reponse:=(Trim(ReadTokenSt(Concat)));
     END;

     if (CleSortie<>'') then
     if (eDate<>sDate) or (eRessource<>sRessource) or (eAffaire<>sAffaire) or (eArticle<>sArticle) or (eTiers<>sTiers) then
     begin
          // PL le 24/06/02 : correction du type article pour qu'il soit compatible avec l'article
          if (sArticle<>'') then
          begin
               try
                  TOBArt:=TOB.Create('ARTICLE',nil,-1) ;
                  rechart := TrouverArticleSQL_GI(false, sArticle, TOBArt, '');
                  if (rechart <> traAucun) then
                     sTypeArticle := TOBArt.GetValue('GA_TYPEARTICLE');
               finally
                      TOBArt.Free; TOBArt:=nil;
               end;
          end;
          //////////////////////////////////

    try
    // On récupère l'enregistrement à modifier
    TobRecup := Tob.Create('EACTIVITE',nil,-1);

    try
      Q:=OpenSQL('SELECT * FROM EACTIVITE WHERE EAC_TYPEACTIVITE="'+csTypAct+'" AND EAC_AFFAIRE="'+csAffaire+'" AND EAC_RESSOURCE="'+csRessource+'" AND EAC_DATEACTIVITE="'+UsDateTime(cdDate)+'" AND EAC_TYPEARTICLE="'+csTypArt+'" AND EAC_NUMLIGNE='+inttostr(ciNumLig),True) ;
      if Not TobRecup.SelectDB('', Q) then exit;
    finally
    ferme(Q);
    end;

    if (TobRecup<>nil) then
    begin
        // On modifie l'enregistrement
        TobRecup.PutValue('EAC_DATEACTIVITE', strtodate(sDate));
        TobRecup.PutValue('EAC_AFFAIRE', sAffaire);
        TobRecup.PutValue('EAC_TIERS', sTiers);
        TobRecup.PutValue('EAC_RESSOURCE', sRessource);
        TobRecup.PutValue('EAC_CODEARTICLE', sArticle);

        // PL le 24/06/02 : correction du type article pour qu'il soit compatible avec l'article
        TobRecup.PutValue('EAC_TYPEARTICLE', sTypeArticle);
        ///////////////////

        // On supprime dans la table l'enregistrement à modifier
        ExecuteSQL('DELETE FROM EACTIVITE WHERE EAC_TYPEACTIVITE="'+csTypAct+'" AND EAC_AFFAIRE="'+csAffaire+'" AND EAC_RESSOURCE="'+csRessource+'" AND EAC_DATEACTIVITE="'+UsDateTime(cdDate)+'" AND EAC_TYPEARTICLE="'+csTypArt+'" AND EAC_NUMLIGNE='+inttostr(ciNumLig)) ;

        // On sauvegarde l'enregistrement modifié en prenant soin de ne pas avoir deux fois le meme numéro de ligne
        try
           Q:=OpenSQL('SELECT MAX(EAC_NUMLIGNE) FROM EACTIVITE WHERE EAC_TYPEACTIVITE="'+csTypAct+'" AND EAC_AFFAIRE="'+sAffaire+'" AND EAC_RESSOURCE="'+sRessource+'" AND EAC_DATEACTIVITE="'+UsDateTime(strtodate(sDate))+'" AND EAC_TYPEARTICLE="'+csTypArt+'" AND EAC_NUMLIGNE='+inttostr(ciNumLig),True);
           if not Q.EOF then
               TobRecup.PutValue('EAC_NUMLIGNE', inttostr(Q.Fields[0].AsInteger+1))
           else
               TobRecup.PutValue('EAC_NUMLIGNE', inttostr(ciNumLig));
        finally
               ferme(Q);
        end;

        // Mise à jour dans la base
        TobRecup.SetAllModifie(true);
        TobRecup.InsertorupdateDB;

        // JP: si suivi saisie déc., on màj le suivi pour la ressource dont une ligne a été modifiée dans l'e-activité
        if bSuivi = TRUE then
           ExecuteSQL ('UPDATE SUIVIEACT SET ASA_ETATEACT = "ES3" WHERE ASA_RESSOURCE="' + sRessource + '" AND ASA_ETATEACT = "ES2" AND ASA_DATEDEB="' + UsDate (EditDateDebut) + '" AND ASA_DATEFIN="' + UsDate (EditDateFin) + '"');
    end;

    finally
    TobRecup.Free;
    end;

    TFMUL(Ecran).ChercheClick;
    end;

end;

procedure TOF_AFEACTIVITE_MUL.BEffaceFicLogOnClick(Sender: TObject);
var i:integer;
    stDocWord : string ;
begin
{$IFDEF EAGLCLIENT}
//AFAIREEAGL
{$ELSE}
stDocWord:=THEdit(GetControl('FICHIERLOG')).Text;
if (stDocWord='') then exit;
if (PGIAskAF ('Confirmez-vous la suppression du fichier ' + stDocWord + ' ?', Ecran.Caption)=mrYes) then
    if not DeleteFile(pchar(stDocWord)) then
        PGIInfoAF('Le fichier ne peut-être supprimé ou n''existe pas', TFMUL(Ecran).caption);
{$ENDIF}
end;

procedure TOF_AFEACTIVITE_MUL.BFicLogOnClick(Sender: TObject);
var i:integer;
    stDocWord : string ;
begin
{$IFDEF EAGLCLIENT}
//AFAIREEAGL
{$ELSE}
stDocWord:=THEdit(GetControl('FICHIERLOG')).Text;
if (stDocWord='') then exit;
if Not FileExists(stDocWord) then
    begin
    PGIInfoAF('Le fichier ' + stDocWord +' n''existe pas', TFMUL(Ecran).caption);
    exit;
    end;
i:=ShellExecute( 0, PCHAR('open'),PChar(stDocWord), Nil,Nil,SW_RESTORE);
{$ENDIF}
end;


// Traitement de la tob
// 1ere fille : contient les formats
function TOF_AFEACTIVITE_MUL.PutClipboardGrille(laTOB : TOB ) : variant;
var TT,TTF,TOBEActivite, TOBDetEActivite : TOB;
    i, j, k, l, NbChampsSup : integer;
    ColName,ColN, St, STC : string;
begin
//=================> T.Detail[0] = tob des formats <=================\\
  if laTOB<>nil then
    begin
    //laTOB.Detail[0].Free ;
    //laTOB.Detail[1].Free ;
//on initialise les colonnes de la grille : format / alignement / type
    (*TT := laTOB.Detail[0];
    for i:=0 to TT.Detail.Count-1 do
      begin
      TTF := TT.Detail[i];
      laGrille.ColFormats[i] := TTF.GetValue('FORMAT');
      laGrille.ColAligns[i] := TTF.GetValue('ALIGN');
      St := TTF.GetValue('TYPE');
      if St<>'' then laGrille.ColTypes[i] := St[1];
      end;
    TT.Free;*)
//si la tob est vide
    if laTOB.Detail.Count=0 then exit;
//Remplissage de la grille
//    laTOB.PutGridDetail(laGrille,true,true,'',true);
//    For k:=0 to laTOB.Detail[2].ChampsSup.Count-1 do ColName:=ColName+TCS(laTOB.Detail[i].ChampsSup[k]).Nom+';' ;
    NbChampsSup := laTOB.Detail[2].ChampsSup.Count;

    TOBEActivite := TOB.Create('la eactivite', Nil, -1);
    for i:=2 to laTOB.Detail.Count-1 do
        begin
        TTF := laTOB.Detail[i];
        if (TTF<>nil) then
            begin
            TOBDetEActivite := TOB.Create('EACTIVITE', TOBEActivite, -1);
            for j:=1 to NbChampsSup do
                begin
                TOBDetEActivite.PutValeur(j, TCS(TTF.ChampsSup[j-1]).Valeur);
                end;
            end;

(*        ColN:=ColName;
        While ColN<>'' do
            BEGIN
            stC:=ReadTokenSt(ColN) ; l:=TTF.GetNumChamp(StC) ;
            if (l>=1000) then
                BEGIN
                St:=TCS(TTF.ChampsSup[l-1000]).Nom ;
                END;
            END;
        (*if (TTF<>nil) then
            begin
            TOBDetEActivite := TOB.Create('EACTIVITE', TOBEActivite, -1);
            for j:=1 to TOBDetEActivite.NbChamps do
                begin
                TOBDetEActivite.PutValeur(j, TTF.GetValeur(j));
                end;
            end;*)
        end;


    TOBEActivite.InsertDB(nil, true);
    TOBEActivite.Free;
    TFMUL(Ecran).ChercheClick;
    end;
end;


procedure TOF_AFEACTIVITE_MUL.BEffaceOnClick(Sender: TObject);
var
   sDelete, sClauseWhere  :string;
   TypeActivite           :string;
   Affaire                :string;
   Ressource              :string;
   DateActivite           :TDateTime;
   TypeArticle            :string;
   NumLigne               :string;
   FListe                 :THDBGrid;
   Pages                  :TPageControl;
   i, iIndex              :integer;
   ListeAss               :TStringList;
   TOBEAct                :TOB;
begin
     FListe  :=TFMul(Ecran).FListe;
     Pages   :=TFMul(Ecran).Pages;

     //mcd 18/07/01 if ExisteSQL('SELECT * FROM EACTIVITE') then
     //if ExisteSQL('SELECT EAC_TYPEACTIVITE FROM EACTIVITE') then
     if (FListe.NbSelected = 0) and (FListe.AllSelected = FALSE) then
     begin
          PgiInfoAf ('Veuillez sélectionner les lignes à supprimer', Ecran.Caption);
          exit;
//          Fliste.AllSelected:=true;
  //        if (FListe.NbSelected <> 0) then
     //        TToolBarButton97 (GetControl('bSelectAll')).Down := True;
     end;

     if PGIAskAF ('Confirmez-vous la suppression de toutes les lignes sélectionnées de la table e-activité ?', Ecran.Caption) = mrYes then
     begin
          ListeAss := TStringList.Create;
          try
             // Distinction quand tout sélectionné et quelques uns sélectionnés
             if Fliste.AllSelected = True then
             begin
                  // Filtre actif sur le mul
                  sClauseWhere := RecupWhereCritere (Pages);

                  // JP: si suivi saisie déc., on mémorise les assistants dont on supprime les lignes
                  if bSuivi = TRUE then
                  begin
                       TOBEAct := TOB.Create ('lignes eactivité', nil, -1);
                       try
                          TOBEAct.LoadDetailFromSQL ('SELECT DISTINCT EAC_RESSOURCE FROM EACTIVITE ' + sClauseWhere + ' ORDER BY EAC_RESSOURCE');
                          for i := 0 to TOBEAct.Detail.Count - 1 do
                              ListeAss.Add (TOBEAct.Detail [i].GetValue ('EAC_RESSOURCE'));
                       finally
                              TOBEAct.Free;
                       end;
                  end;

                  // Suppression de toutes les lignes visibles dans le mul
                  ExecuteSQL ('DELETE FROM EACTIVITE ' + sClauseWhere);
             end
             else
                 if (FListe.NbSelected <> 0) then
                 begin
                      // suppression des lignes sélectionnées
                      for i:=0 to Fliste.NbSelected-1 do
                      begin
                           // Positionnement sur l'élément en sélection
                           Fliste.GotoLeBookmark (i);

                           // Propriétés de la ligne d'e-activité en sélection
                           TypeActivite  := Fliste.datasource.dataset.FindField('EAC_TYPEACTIVITE').AsString;
                           Affaire       := Fliste.datasource.dataset.FindField('EAC_AFFAIRE').AsString;
                           Ressource     := Fliste.datasource.dataset.FindField('EAC_RESSOURCE').AsString;
                           DateActivite  := Fliste.datasource.dataset.FindField('EAC_DATEACTIVITE').AsDateTime;
                           TypeArticle   := Fliste.datasource.dataset.FindField('EAC_TYPEARTICLE').AsString;
                           NumLigne      := Fliste.datasource.dataset.FindField('EAC_NUMLIGNE').AsString;

                           // Suppression de la table eactivité
                           ExecuteSQL ('DELETE FROM EACTIVITE WHERE EAC_TYPEACTIVITE="' + TypeActivite + '" AND EAC_AFFAIRE="' + Affaire + '" AND EAC_RESSOURCE="' + Ressource + '" AND EAC_DATEACTIVITE="' + UsDateTime(DateActivite) + '" AND EAC_TYPEARTICLE="' + TypeArticle + '" AND EAC_NUMLIGNE=' + NumLigne);

                           // JP: si suivi saisie déc., on mémorise les assistants pour mettre à jour l'état du suivi de ceux-ci
                           if bSuivi = TRUE then
                              if ListeAss.Find (Ressource, iIndex) = FALSE then
                                 ListeAss.Add (Ressource);
                      end;
                 end;

             // JP: si suivi saisie déc., on màj le suivi pour chaque ressource dont au - une ligne a été supprimé
             if bSuivi = TRUE then
                for i := 0 to ListeAss.Count-1 do
                    ExecuteSQL ('UPDATE SUIVIEACT SET ASA_ETATEACT = "ES3" WHERE ASA_RESSOURCE="' + ListeAss [i] + '" AND ASA_ETATEACT = "ES2" AND ASA_DATEDEB="' + UsDate (EditDateDebut) + '" AND ASA_DATEFIN="' + UsDate (EditDateFin) + '"');
          finally
                 ListeAss.Free;
          end;

          TToolBarButton97 (GetControl ('bSelectAll')).Down := False;
          TFMUL (Ecran).ChercheClick;
     end;
end;

procedure TOF_AFEACTIVITE_MUL.BImportOnClick(Sender: TObject);
var
   T : TOB;
   i, dernieri, NbEactiviteAvant, NbEActiviteApres, NbAffaires : integer;
   sAff0,sAff1,sAff2,sAff3,sAff4,sAffaire,sTiers,sNomRep:string;
   TobTempo:TOB;
   SQL,NomCompletFic:string;
   Q:TQuery;
   SearchRec: TSearchRec;
begin
NomCompletFic:='';

V_PGI.ZoomOLE := True;   //pour passer la fiche en modal
sNomRep := AglLanceFiche ('AFF', 'AFEACTIMPORT', '','',  GetParamSoc ('SO_AFEACTPATH'));
V_PGI.ZoomOLE := False;
//rep:= AglLanceFiche( 'AFF', 'AFEACTIVITEREPIMP', '','','') ; //AglLanceFiche( 'AFF', 'AFEACTIVITEREPIMP', '','','') ;

if (sNomRep='') then
   exit;

NbEactiviteAvant:=CompteEActivite;
//sNomRep := rep, 9, length(Rep)-8);
if not DirectoryExists(sNomRep) then begin PGIInfoAf('Le répertoire: "' + sNomRep + '" n''existe pas', TFMUL(Ecran).caption); exit; end;

dernieri:=0;
if TControl(Sender).Name = 'BIMPORTEXCEL' then // pos('EXCEL', rep)<>0 then
    begin
    FindFirst(sNomRep+'\IMPACT_*.xls', faAnyFile, SearchRec);
    repeat
    if (SearchRec.Name<>'') then
    if FileExists (sNomRep+'\'+SearchRec.Name) then
        begin
        NomCompletFic := sNomRep+'\'+SearchRec.Name;

        // récupération de la tob
        T := TOB.Create('', Nil, -1);

        try
        // barre de progression
        // titres de la table dans les titres des colonnes
        ImportTOBFromXLS (T, NomCompletFic, true, true);

        if (T<>nil) then
        for i:=0 to T.detail.count-1 do
        begin
            if (T.Detail[i].GetValue('EAC_TYPEACTIVITE')='') then
               T.Detail[i].PutValue('EAC_TYPEACTIVITE', 'REA');
            if (T.Detail[i].GetValue('EAC_AFFAIRE')='') then
            begin
                NbAffaires:=0;
                sAff0:=T.Detail[i].GetValue('EAC_AFFAIRE0');
                sAff1:=T.Detail[i].GetValue('EAC_AFFAIRE1');
                sAff2:=T.Detail[i].GetValue('EAC_AFFAIRE2');
                sAff3:=T.Detail[i].GetValue('EAC_AFFAIRE3');
                sAff4:=T.Detail[i].GetValue('EAC_AVENANT');
                sTiers:=T.Detail[i].GetValue('EAC_TIERS');
                sAffaire:=RegroupePartiesAffaire (sAff0, sAff1, sAff2, sAff3, sAff4);
                NbAffaires:=TeststCleAffaire( sAffaire,sAff0,sAff1,sAff2,sAff3,sAff4,sTiers, false, false, false, true);
                if (NbAffaires=1) then
                begin
                    T.Detail[i].PutValue('EAC_AFFAIRE',sAffaire);
                    T.Detail[i].PutValue('EAC_AFFAIRE0',sAff0);
                    T.Detail[i].PutValue('EAC_AFFAIRE1',sAff1);
                    T.Detail[i].PutValue('EAC_AFFAIRE2',sAff2);
                    T.Detail[i].PutValue('EAC_AFFAIRE3',sAff3);
                    T.Detail[i].PutValue('EAC_AVENANT',sAff4);
                    T.Detail[i].PutValue('EAC_TIERS',sTiers);
                end;
//                T.Detail[i].PutValue('EAC_ACTORIGINE', '');
            end;
            T.Detail[i].PutValue('EAC_NUMLIGNE', NbEactiviteAvant+1+i+dernieri);
            end;

        T.insertDB(nil,false);
        finally
               dernieri:=dernieri+i;
               T.Free;
        end;
        end;
    until FindNext(SearchRec)<>0 ;
    sysutils.FindClose(SearchRec);
    end
else
//if pos('ASCII', rep)<>0 then
begin
    FindFirst(sNomRep+'\*.asc', faAnyFile, SearchRec);
    repeat
    if (SearchRec.Name<>'') then
    if FileExists (sNomRep+'\'+SearchRec.Name) then
        try
        TobImport := TOB.Create('', nil, -1);
        NomCompletFic := sNomRep+'\'+SearchRec.Name;
        if FileExists (NomCompletFic) then
            begin
            TobLoadFromFile( NomCompletFic, Nil, TobImport ) ;
            end;
        finally
            if (TobImport<>nil) then
                if (TobImport.Detail.count <>0) then
                    begin
                    if (TobImport.Detail[0].NomTable<>'EACTIVITE') then
                        PGIInfoAf('Les lignes importées ne sont pas des lignes de l''e-activité. Table : '+TobImport.Detail[0].NomTable, TFMUL(Ecran).caption)
                    else
                        TobImport.InsertDB(nil, true);
                    end;
            TobImport.Free;
        end;
    until FindNext(SearchRec)<>0 ;
    sysutils.FindClose(SearchRec);
end;

TFMUL(Ecran).ChercheClick;
NbEActiviteApres:=CompteEActivite;

if (NbEActiviteApres>NbEActiviteAvant) then // N lignes ont été importées dans l'e-activité
    PGIInfoAf(inttostr(NbEActiviteApres-NbEActiviteAvant) + ' lignes ont été importées dans l''e-activité', TFMUL(Ecran).caption);

// JP: à priori nombre ligne après import >= nombre ligne avant
//if (NbEActiviteApres<NbEActiviteAvant) then // Tout ou partie des lignes de l'e-activité ont été perdues
  //  PGIInfoAf('Tout ou partie des lignes de l''e-activité ont été supprimées', TFMUL(Ecran).caption);

TToolBarButton97(GetControl('bSelectAll')).Down:=false;
TFMul(Ecran).FListe.AllSelected:=False;
end;


function TOF_AFEACTIVITE_MUL.CompteEActivite:integer;
var
Q:TQuery;
SQL:string;
begin
Result:=0;
SQL:='SELECT COUNT(EAC_DATEACTIVITE) FROM EACTIVITE' ;
try
   Q:=OpenSQL(SQL,True) ;
   if Not Q.EOF then
        begin
        Result:=Q.Fields[0].AsInteger;
        end;
finally
Ferme(Q) ;
end;
end;

procedure TOF_AFEACTIVITE_MUL.BOuvrirOnClick(Sender: TObject);
var
    TobEA, TOBEASelect, TOBLigne, TobTemp :TOB;
    Q, QQ :TQuery;
    i,iNbEactivite,RepImport:integer;
    FListe:THDBGrid;
    TypeArticle, TypeActivite, Affaire, Ressource, st:string;
    DateActivite:TDateTime;
    NumLigne:Integer;
    bFinOK:boolean;
begin
      iNbEactivite:=0; RepImport:=0; bFinOK:=false;

      FListe:=TFMul(Ecran).FListe;
      if (FListe=nil) then exit;
      Q:=TFMul(Ecran).Q;
      if Q.EOF then exit;
      QQ:=PrepareSQL ('SELECT * FROM EACTIVITE', True);
      RecupWhereSQL (Q, QQ);
      QQ.Open;

      TobEA:=TOB.Create('',Nil,-1);
      TOBEASelect:=TOB.Create('',Nil,-1);
      try
      SourisSablier;

      // On remplit la tob avec les lignes de l'e-activité
      TobEA.LoadDetailDB('EACTIVITE','','',QQ,true) ;
      Ferme(QQ);

      if Fliste.AllSelected then
         BEGIN
         iNbEactivite:=TobEA.Detail.count;

        // Importation dans ACTIVITE, suppression dans EACTIVITE
         if (iNbEactivite<>0) then
              RepImport := ImportEActivite(TobEA, TRUE);

         bFinOK:=true;
         Fliste.AllSelected:=False;
         TToolBarButton97(GetControl('bSelectAll')).Down:=false;
         END
      Else
      Begin
        // Remplissage de la TOB avec les lignes sélectionnées
        for i:=0 to Fliste.NbSelected-1 do
        begin
              Fliste.GotoLeBookmark(i);
              TypeActivite:=Fliste.datasource.dataset.FindField('EAC_TYPEACTIVITE').AsString;
              Affaire:=Fliste.datasource.dataset.FindField('EAC_AFFAIRE').AsString;
              Ressource:=Fliste.datasource.dataset.FindField('EAC_RESSOURCE').AsString;
              DateActivite:=Fliste.datasource.dataset.FindField('EAC_DATEACTIVITE').AsDateTime;
              TypeArticle:=Fliste.datasource.dataset.FindField('EAC_TYPEARTICLE').AsString;
              NumLigne:=Fliste.datasource.dataset.FindField('EAC_NUMLIGNE').AsInteger;

              TOBLigne := TobEA.FindFirst(['EAC_TYPEACTIVITE','EAC_AFFAIRE','EAC_RESSOURCE','EAC_DATEACTIVITE','EAC_TYPEARTICLE','EAC_NUMLIGNE']
                                          ,[TypeActivite, Affaire, Ressource, DateActivite, TypeArticle, NumLigne], true);
              if (TOBLigne<>nil) then
                  TOBLigne.ChangeParent(TobEASelect,-1) ;
        end;

        iNbEactivite:=TobEASelect.Detail.count;

        // Importation dans ACTIVITE, suppression dans EACTIVITE
        if (iNbEactivite<>0) then
           RepImport := ImportEActivite (TobEASelect, (TOBEASelect.Detail.Count = FListe.DataSource.DataSet.RecordCount));

        bFinOK:=true;
        Fliste.ClearSelected;
        End;

      // Rafraichit la liste
      TFMul(Ecran).ChercheClick;

      finally
      //st := THEdit(GetControl('FICHIER')).Text;
      //if (st<>'') then
      //    TobEA.SaveToFile(st + '2', false, false, false);
      TOBEASelect.Free;
      TobEA.Free;
      SourisNormale;

      // Message compte rendu de l'importation
      if not bFinOK then
          PGIInfoAf('Importation impossible. Une erreur bloquante est survenue.', Ecran.Caption)
      else
      if (iNbEactivite<>0) then
          begin
          if RepImport<>0 then
              PGIInfoAf('Certaines lignes présentent des erreurs. Veuillez consulter le fichier : '+ NomFichierLog, Ecran.Caption);

          PGIInfoAf(inttostr(iNbEactivite-RepImport) + textemessage[1]+ inttostr(RepImport) + textemessage[2], Ecran.Caption);
          end;
      end;
end;


function TOF_AFEACTIVITE_MUL.NomFichierLog:string;
var
NomLog,NomFicLog,NomFic1,NomFic2,StPath2,StPath1:string;
begin

Result:='';
NomLog:='';
NomFicLog:='';
StPath2:='';StPath1:='';NomFic1:='';NomFic2:='';
NomLog := THEdit(GetControl('FICHIERLOG')).Text;
if (NomLog<>'') then
    begin
    StPath1:=ExtractFilePath(NomLog) ;
    NomFic1:=ChangeFileExt(ExtractFileName(NomLog),'.log') ;
    end;

if (StPath1<>'') then
    begin
    if (NomFic1<>'') then
        NomFicLog:=StPath1+NomFic1
    else
        NomFicLog:=StPath1+'ImportEActivite.log';
    end;

if (NomFicLog='') then
    begin
    NomFicLog:='c:\ImportEActivite.log';
    end;

THEdit(GetControl('FICHIERLOG')).Text := NomFicLog;
Result:=NomFicLog;
end;

function TOF_AFEACTIVITE_MUL.ImportEActivite (TobEActivite:TOB; bAll:boolean):integer;
var
    TobActivite, TobLigneAct, TOBArtRef:TOB;
    CleCourante, ChampSansPref, UniteValo, Article:string;
    i,j,indLibChamp,MaxNumLigneCourant:integer;
    AFOAssistants:TAFO_Ressources;
    TOBArticles, TOBValo, TOBArt:TOB;
    IndexRess, IndexCle, iNbEactivite:integer;
    ListeFonctionsRess, ListeDesCles:TStringList;
    Part0,Part1,Part2,Part3,Avenant,Tiers,Affaire : string;
    dCoeffConvert:double;
    bDelete : Boolean;
    NomFicLog      :string;
    CodeAssistant  :string;
    AssistListe    :TStringList;
//    ValideListe    :TStringList;
//    NonValideListe :TStringList;
begin
      NomFicLog:='';
      Result:=0;

      // gestion du fichier log
      NomFicLog:=NomFichierLog;

      if (NomFicLog<>'') then
          begin
          if (THEdit(GetControl('FICHIERLOG')).Text='') then
              THEdit(GetControl('FICHIERLOG')).Text := NomFicLog;

          AssignFile(FichierLog, NomFicLog);
          // on cree ou réouvre le fichier de log
          if (FileExists(NomFicLog) = TRUE) then  Append (FichierLog)
            else  Rewrite (FichierLog);

          writeln (FichierLog, '********************************************************************************************************' );
          writeln (FichierLog, '************ Début de l''importation des données de l''eactivité ' + FormatDateTime('dd/mm/yyyy ttttt',CurrentDate));
          writeln (FichierLog, '' );
          end;

      TobActivite:=TOB.Create('liste activite',Nil,-1);
      TOBArticles:=TOB.Create('ARTICLES',Nil,-1) ;
      AFOAssistants:=TAFO_Ressources.Create;
      ListeDesCles:=TStringList.Create;
      try
         if (TobEActivite.detail.Count>=0) then
            TobEActivite.Detail[0].AddChampSup ('VALIDE', true);

      // Initialisation liste des assistants valides et non valides pour la période donnée
      AssistListe := TStringList.Create;
      AssistListe.Sorted := TRUE;
//      ValideListe     := TStringList.Create;
  //    ValideListe.Sorted := TRUE;
    //  NonValideListe  := TStringList.Create;
      //NonValideListe.Sorted := TRUE;

      // Sur toutes les lignes de l'eactivite, on genere une ligne d'activite en la complétant
      for i:=0 to TobEActivite.detail.Count-1 do
      begin
          // On créé la ligne d'activite
          TobLigneAct:=TOB.Create('ACTIVITE',TobActivite,-1);
          TobLigneAct.InitValeurs;

          // La ligne n'est pour l'instant pas importée, donc "non valide" (terme peu approprié)
          TobEActivite.Detail [i].PutValue ('VALIDE', '-');

          // On stocke le code assistant (si pas déjà présent dans la liste), pour pouvoir màj le suivi à la fin
          CodeAssistant := TobEActivite.Detail [i].GetValue ('EAC_RESSOURCE');
          if AssistListe.Find (CodeAssistant, j) = FALSE then
             AssistListe.Add (CodeAssistant);
             
          // On remplit tous les champs avec les données stockées dans l'eactivite
          if (NomFicLog<>'') then
          begin
              writeln (FichierLog, '' );
              writeln (FichierLog, '************ Importation des données eactivité brutes ligne : '+inttostr(i) + ' ' + FormatDateTime('dd/mm/yyyy ttttt',CurrentDate));
          end;

          for j:=1 to TobEActivite.detail[i].NbChamps do
          begin
              indLibChamp:=pos('_', TobEActivite.detail[i].GetNomChamp(j))+1;
              if (indLibChamp=0) then continue;
              ChampSansPref := copy(TobEActivite.detail[i].GetNomChamp(j), indLibChamp, Length(TobEActivite.detail[i].GetNomChamp(j))-indLibChamp+1);
              TobLigneAct.PutValue('ACT_'+ ChampSansPref, TobEActivite.detail[i].GetValeur(j));
          end;


          // Complément d'initialisation
          if (TobLigneAct.GetValue('ACT_TYPEACTIVITE')='') then
              begin
              if (NomFicLog<>'') then
                  writeln (FichierLog, TraduitGA('Le type d''activité n''était pas saisi => Complément du type d''activité'));
              TobLigneAct.PutValue('ACT_TYPEACTIVITE','REA');
              end;

          IndexRess := AFOAssistants.AddRessource(TobLigneAct.GetValue('ACT_RESSOURCE'));
          // pas d'articles, on reprend par défaut celui de la ressource
          if (TobLigneAct.GetValue('ACT_CODEARTICLE') = '') and (IndexRess<>-1) and (IndexRess<>-2) then
            BEGIN
            if (NomFicLog<>'') then
                  writeln (FichierLog, TraduitGA('L''article n''était pas saisi => Complément de l''article par celui de la ressource'));
            TobLigneAct.PutValue('ACT_CODEARTICLE', Copy(TAFO_Ressource(AFOAssistants.Objects[IndexRess]).tob_Champs.GetValue('ARS_ARTICLE'),1,18));
            TobLigneAct.PutValue('ACT_TYPEARTICLE', 'PRE');
            END;

          if (TobLigneAct.GetValue('ACT_TYPERESSOURCE') = '') and (IndexRess<>-1) and (IndexRess<>-2) then
              begin
              if (NomFicLog<>'') then
                  writeln (FichierLog, TraduitGA('Le type de la ressource n''était pas saisi => Complément du type de la ressource'));
              TobLigneAct.PutValue('ACT_TYPERESSOURCE', TAFO_Ressource(AFOAssistants.Objects[IndexRess]).tob_Champs.GetValue('ARS_TYPERESSOURCE'));
              end;

          if (IndexRess=-1) or (IndexRess=-2) then
              begin
              if (NomFicLog<>'') then
                  writeln (FichierLog, TraduitGA('La ressource '+TobLigneAct.GetValue('ACT_RESSOURCE')+' est inconnue'));
              TobLigneAct.PutValue('ACT_RESSOURCE','');
              end;

          // Contrôle Affaire  si invalide raz on sort ...
          Part0 := ''; Part1:=''; Part2:=''; Part3 := ''; Avenant:=''; Tiers:='';
          Affaire := TobLigneAct.GetValue('ACT_AFFAIRE');
          if TeststCleAffaire(Affaire,Part0,Part1,Part2,Part3,Avenant,Tiers,false, false,false,True) = 1 then
            BEGIN
            TobLigneAct.PutValue('ACT_AFFAIRE', Affaire); TobLigneAct.PutValue('ACT_TIERS', Tiers);
            END
          else
              begin
              TobLigneAct.PutValue('ACT_AFFAIRE','');
              if (NomFicLog<>'') then
                  writeln (FichierLog, TraduitGA('L''affaire ' + Affaire + ' est inconnue'));
              end;

          TOBArt:=TrouverCodeArticle(TobLigneAct.GetValue('ACT_CODEARTICLE'), TobLigneAct.GetValue('ACT_TYPEARTICLE'), TOBArticles);
          if ((TobLigneAct.GetValue('ACT_TYPEARTICLE')='') and (TOBArt<>nil)) then
              begin
              TobLigneAct.PutValue('ACT_TYPEARTICLE', TOBArt.GetValue('GA_TYPEARTICLE'));
              if (NomFicLog<>'') then
                  writeln (FichierLog, TraduitGA('Le type de l''article n''était pas saisi => Complément du type d''article'));
              end;

          // on controle la validite de la clé de la ligne d'activite
          if  (TOBArt=nil) or (TobLigneAct.GetValue('ACT_TYPEACTIVITE')='') or (TobLigneAct.GetValue('ACT_AFFAIRE')='')
              or ((TobLigneAct.GetValue('ACT_RESSOURCE')='') and (TobLigneAct.GetValue('ACT_TYPEARTICLE')<>'MAR'))
              or (TobLigneAct.GetValue('ACT_DATEACTIVITE')=0)
              or (TobLigneAct.GetValue('ACT_DATEACTIVITE')<gDateDebutAct) or (TobLigneAct.GetValue('ACT_DATEACTIVITE')>gDateFinAct)
              or (TobLigneAct.GetValue('ACT_TYPEARTICLE')='') then
                  begin   // ligne non reprise
//                  TobEActivite.detail[i].PutValue('VALIDE', '-');
                  Result:= Result+1;

                  // remplissage du fichier log
                  if (NomFicLog<>'') then
                      begin
                      writeln (FichierLog, '');
                      writeln (FichierLog, TraduitGA('La clé n''est pas complète pour les raisons suivantes :'));

                      if (TOBArt=nil) then
                          writeln (FichierLog, TraduitGA('L''article est vide ou inconnu'));

                      if (TobLigneAct.GetValue('ACT_TYPEARTICLE')='') then
                          writeln (FichierLog, TraduitGA('Le type d''article est vide ou inconnue'));

                      if (TobLigneAct.GetValue('ACT_TYPEACTIVITE')='') then
                          writeln (FichierLog, TraduitGA('Le type d''activité est vide ou inconnu'));

                      if (TobLigneAct.GetValue('ACT_AFFAIRE')='') then
                          writeln (FichierLog, TraduitGA('L''affaire est vide ou inconnue'));

                      if ((TobLigneAct.GetValue('ACT_RESSOURCE')='') and (TobLigneAct.GetValue('ACT_TYPEARTICLE')<>'MAR')) then
                          writeln (FichierLog, TraduitGA('La ressource est vide ou inconnu'));

                      if  (TobLigneAct.GetValue('ACT_DATEACTIVITE')=0) or (TobLigneAct.GetValue('ACT_DATEACTIVITE')<gDateDebutAct) or (TobLigneAct.GetValue('ACT_DATEACTIVITE')>=gDateFinAct) then
                          writeln (FichierLog, TraduitGA('La date d''activité est incorrecte ou n''est pas dans l''intervalle de saisie d''activité'));

                      writeln (FichierLog, '');
                      writeln (FichierLog, TraduitGA('La ligne n''a pu être importée dans l''activité'));
                      end;

                  TobLigneAct.Free;
                  TobLigneAct:=nil;
                  continue;
          end;

          // clé primaire = ACT_TYPEACTIVITE,ACT_AFFAIRE,ACT_RESSOURCE, ACT_DATEACTIVITE,ACT_FOLIO, ACT_TYPEARTICLE,ACT_NUMLIGNE
          TobEActivite.detail[i].PutValue('VALIDE', 'X');
          CleCourante:=TobLigneAct.GetValue('ACT_TYPEACTIVITE')+';'+TobLigneAct.GetValue('ACT_AFFAIRE')+';'
                      +TobLigneAct.GetValue('ACT_RESSOURCE')+';'+DateToStr(TobLigneAct.GetValue('ACT_DATEACTIVITE'))+';'
                      +TobLigneAct.GetValue('ACT_TYPEARTICLE');
          IndexCle:=ListeDesCles.IndexOf(CleCourante);
          MaxNumLigneCourant:=1;
          if (IndexCle=-1) then
              begin
              // On recherche le max pour cette clé
              MaxNumLigneCourant := MaxNumLigneActivite(TobLigneAct) + 1;
              ListeDesCles.AddObject(CleCourante, TObject(MaxNumLigneCourant));
              end
          else
              begin
              MaxNumLigneCourant:=integer(ListeDesCles.Objects[IndexCle])+1;
              ListeDesCles.Objects[IndexCle]:=TObject(MaxNumLigneCourant);
              end;

          // on complete la ligne d'activite avec ce qu'on connait
          TobLigneAct.PutValue('ACT_DEVISE', V_PGI.DevisePivot);
          TobLigneAct.PutValue('ACT_NUMLIGNE', MaxNumLigneCourant);
          TobLigneAct.PutValue('ACT_FOLIO', 1);
          TobLigneAct.PutValue('ACT_QTEFAC', TobLigneAct.GetValue('ACT_QTE'));
          TobLigneAct.PutValue('ACT_UNITEFAC', TobLigneAct.GetValue('ACT_UNITE'));
          TobLigneAct.PutValue('ACT_SEMAINE', NumSemaine(TobLigneAct.GetValue('ACT_DATEACTIVITE')));
          TobLigneAct.PutValue('ACT_PERIODE', GetPeriode(TobLigneAct.GetValue('ACT_DATEACTIVITE')));

          // idem pour les donnees liees à l'affaire
          CodeAffaireDecoupe(TobLigneAct.GetValue('ACT_AFFAIRE'), Part0, Part1, Part2, Part3, Avenant, taModif, false);
          TobLigneAct.PutValue('ACT_AFFAIRE0', Part0); TobLigneAct.PutValue('ACT_AFFAIRE1', Part1);
          TobLigneAct.PutValue('ACT_AFFAIRE2', Part2); TobLigneAct.PutValue('ACT_AFFAIRE3', Part3);
          TobLigneAct.PutValue('ACT_AVENANT', Avenant);

          //
          // On valorise
          //
              // On vide la tob valo
              if (TOBValo<>nil) then begin TOBValo.Free; end;

              // Valorisation
              TOBValo:=MajTOBValo(TobLigneAct.GetValue('ACT_DATEACTIVITE'), tacGlobal, TobLigneAct.GetValue('ACT_AFFAIRE'), TobLigneAct.GetValue('ACT_RESSOURCE'),
                                  TobLigneAct.GetValue('ACT_CODEARTICLE'), TobLigneAct.GetValue('ACT_TYPEHEURE'), TOBArticles, nil, nil, AFOAssistants, false, VH_GC.AFValoActPR, VH_GC.AFValoActPV);

              if (TOBValo<>nil) then
                  begin
                  // L'unité de valorisation est l'unité dans laquelle on a valorisé les prix dans la fonction MajTOBValo
                  UniteValo:=TOBValo.GetValue('GA_QUALIFUNITEACT');
                  if (TobLigneAct.GetValue('ACT_UNITE')='') then TobLigneAct.PutValue('ACT_UNITE', UniteValo);
                  // Si ce n'est pas convertible, on ne touche pas aux données saisies
                  if IsConvertible(TobLigneAct.GetValue('ACT_UNITE'), UniteValo) or (UniteValo='') then
                      begin
                      dCoeffConvert:=1;
                      // On calcule le coefficient de conversion pour passer de l'unité de saisie en unité de valorisation
                      if (UniteValo<>TobLigneAct.GetValue('ACT_UNITE')) and (UniteValo<>'') then
                          begin
                          // Les prix fournis sont en unité de saisie, il faut les convertir en unité de valorisation
                          dCoeffConvert := ConversionUnite( UniteValo, TobLigneAct.GetValue('ACT_UNITE'), 1);
                          if dCoeffConvert=0 then dCoeffConvert:=1;
                          end;
                      // On réajuste la quantité pour qu'elle soit exprimée en unité de valorisation
                      TobLigneAct.PutValue('ACT_QTE', TobLigneAct.GetValue('ACT_QTE') / dCoeffConvert);
                      // On stock l'unite de valorisation
                      TobLigneAct.PutValue('ACT_UNITE',UniteValo) ;
                      TobLigneAct.PutValue('ACT_UNITEFAC',UniteValo) ;

                      if (TobLigneAct.GetValue('ACT_PUPRCHARGE')=0) then
                          // Si le prix unitaire de revient n'a pas été saisi,
                          // on le valorise et on remplit tous les champs lies
                          begin
                          // PL le 23/01/02 pour gérer le cas où le PU est nul mais pas le Total
                          if (TobLigneAct.GetValue('ACT_TOTPRCHARGE')=0) then
                              begin
                              TobLigneAct.PutValue('ACT_PUPRCHARGE',TOBValo.GetValue('GA_PMRP')) ;
                              TobLigneAct.PutValue('ACT_PUPR',TOBValo.GetValue('GA_PMRP')) ;
                              TobLigneAct.PutValue('ACT_PUPRCHINDIRECT',TOBValo.GetValue('GA_PMRP')) ;
                              TobLigneAct.PutValue('ACT_TOTPRCHARGE', TobLigneAct.GetValue('ACT_QTE')*TobLigneAct.GetValue('ACT_PUPRCHARGE'));
                              TobLigneAct.PutValue('ACT_TOTPR',TobLigneAct.GetValue('ACT_QTE')*TobLigneAct.GetValue('ACT_PUPR'));
                              TobLigneAct.PutValue('ACT_TOTPRCHINDI',TobLigneAct.GetValue('ACT_QTE')*TobLigneAct.GetValue('ACT_PUPRCHINDIRECT'));
                              end
                          else
                              begin
                              // TOTPRCHARGE
                              TobLigneAct.PutValue('ACT_TOTPRCHARGE',TobLigneAct.GetValue('ACT_TOTPRCHARGE')/ dCoeffConvert) ;
                              // TOTPR
                              TobLigneAct.PutValue('ACT_TOTPR',TobLigneAct.GetValue('ACT_TOTPR')/ dCoeffConvert);
                              if (TobLigneAct.GetValue('ACT_TOTPR')=0) then
                                  TobLigneAct.PutValue('ACT_TOTPR',TobLigneAct.GetValue('ACT_TOTPRCHARGE')) ;
                              // TOTPRCHINDI
                              TobLigneAct.PutValue('ACT_TOTPRCHINDI',TobLigneAct.GetValue('ACT_TOTPRCHINDI')/ dCoeffConvert);
                              if (TobLigneAct.GetValue('ACT_TOTPRCHINDI')=0) then
                                  TobLigneAct.PutValue('ACT_TOTPRCHINDI',TobLigneAct.GetValue('ACT_TOTPRCHARGE')) ;
                              end;
                          // Fin PL le 23/01/02
                          end
                      else
                          // Sinon, on convertit toutes les données saisies en unité de valorisation
                          begin
                          // PUPRCHARGE
                          TobLigneAct.PutValue('ACT_PUPRCHARGE', TobLigneAct.GetValue('ACT_PUPRCHARGE')/ dCoeffConvert) ;
                          // PUPR
                          TobLigneAct.PutValue('ACT_PUPR',TobLigneAct.GetValue('ACT_PUPR')/ dCoeffConvert) ;
                          if (TobLigneAct.GetValue('ACT_PUPR')=0) then
                              TobLigneAct.PutValue('ACT_PUPR',TobLigneAct.GetValue('ACT_PUPRCHARGE')) ;
                          // PUPRCHINDIRECT
                          TobLigneAct.PutValue('ACT_PUPRCHINDIRECT',TobLigneAct.GetValue('ACT_PUPRCHINDIRECT')/ dCoeffConvert) ;
                          if (TobLigneAct.GetValue('ACT_PUPRCHINDIRECT')=0) then
                              TobLigneAct.PutValue('ACT_PUPRCHINDIRECT',TobLigneAct.GetValue('ACT_PUPRCHARGE')) ;
                          // TOTPRCHARGE
                          TobLigneAct.PutValue('ACT_TOTPRCHARGE',TobLigneAct.GetValue('ACT_TOTPRCHARGE')/ dCoeffConvert) ;
                          if (TobLigneAct.GetValue('ACT_TOTPRCHARGE')=0) then
                              TobLigneAct.PutValue('ACT_TOTPRCHARGE', TobLigneAct.GetValue('ACT_QTE')*TobLigneAct.GetValue('ACT_PUPRCHARGE')) ;
                          // TOTPR
                          TobLigneAct.PutValue('ACT_TOTPR',TobLigneAct.GetValue('ACT_TOTPR')/ dCoeffConvert);
                          if (TobLigneAct.GetValue('ACT_TOTPR')=0) then
                              TobLigneAct.PutValue('ACT_TOTPR',TobLigneAct.GetValue('ACT_QTE')*TobLigneAct.GetValue('ACT_PUPR')) ;
                          // TOTPRCHINDI
                          TobLigneAct.PutValue('ACT_TOTPRCHINDI',TobLigneAct.GetValue('ACT_TOTPRCHINDI')/ dCoeffConvert);
                          if (TobLigneAct.GetValue('ACT_TOTPRCHINDI')=0) then
                              TobLigneAct.PutValue('ACT_TOTPRCHINDI',TobLigneAct.GetValue('ACT_QTE')*TobLigneAct.GetValue('ACT_PUPRCHINDIRECT')) ;
                          end;

                      if (TobLigneAct.GetValue('ACT_PUVENTE')=0) then
                          // Si le prix unitaire de vente n'a pas été saisi,
                          // on le valorise et on remplit tous les champs lies
                          begin
                          // PL le 23/01/02 pour gérer le cas où le PU est nul mais pas le Total
                          if (TobLigneAct.GetValue('ACT_TOTVENTE')=0) then
                              // Si le total de vente n'a pas été saisi,
                              // on le valorise et on remplit tous les champs lies
                              begin
                              TobLigneAct.PutValue('ACT_PUVENTE',TOBValo.GetValue('GA_PVHT')) ;
                              TobLigneAct.PutValue('ACT_PUVENTEDEV',TOBValo.GetValue('GA_PVHT')) ;
                              TobLigneAct.PutValue('ACT_PUVENTECON',ConvertSaisieEF( TOBValo.GetValue('GA_PVHT'), false)) ;
                              TobLigneAct.PutValue('ACT_TOTVENTE',TobLigneAct.GetValue('ACT_QTE')*TobLigneAct.GetValue('ACT_PUVENTE')) ;
                              TobLigneAct.PutValue('ACT_TOTVENTEDEV',TobLigneAct.GetValue('ACT_QTE')*TobLigneAct.GetValue('ACT_PUVENTEDEV')) ;
                              TobLigneAct.PutValue('ACT_TOTVENTECON',TobLigneAct.GetValue('ACT_QTE')*TobLigneAct.GetValue('ACT_PUVENTECON')) ;
                              end
                          else
                              // Sinon, on convertit toutes les données saisies en unité de valorisation
                              begin
                              // TOTVENTE
                              TobLigneAct.PutValue('ACT_TOTVENTE',TobLigneAct.GetValue('ACT_TOTVENTE')/ dCoeffConvert) ;
                              // TOTVENTEDEV
                              TobLigneAct.PutValue('ACT_TOTVENTEDEV',TobLigneAct.GetValue('ACT_TOTVENTEDEV')/ dCoeffConvert) ;
                              if (TobLigneAct.GetValue('ACT_TOTVENTEDEV')=0) then
                                  TobLigneAct.PutValue('ACT_TOTVENTEDEV',TobLigneAct.GetValue('ACT_TOTVENTE')) ;
                              // TOTVENTECON
                              TobLigneAct.PutValue('ACT_TOTVENTECON',TobLigneAct.GetValue('ACT_TOTVENTECON')/ dCoeffConvert) ;
                              if (TobLigneAct.GetValue('ACT_TOTVENTECON')=0) then
                                  TobLigneAct.PutValue('ACT_TOTVENTECON',TobLigneAct.GetValue('ACT_TOTVENTE')) ;
                              end;
                          // Fin PL le 23/01/02
                          end
                      else
                          // Sinon, on convertit toutes les données saisies en unité de valorisation
                          begin
                          // PUVENTE
                          TobLigneAct.PutValue('ACT_PUVENTE',TobLigneAct.GetValue('ACT_PUVENTE')/ dCoeffConvert) ;
                          // PUVENTEDEV
                          TobLigneAct.PutValue('ACT_PUVENTEDEV',TobLigneAct.GetValue('ACT_PUVENTEDEV')/ dCoeffConvert) ;
                          if (TobLigneAct.GetValue('ACT_PUVENTEDEV')=0) then
                              TobLigneAct.PutValue('ACT_PUVENTEDEV',TobLigneAct.GetValue('ACT_PUVENTE')) ;
                          // PUVENTECON
                          TobLigneAct.PutValue('ACT_PUVENTECON',TobLigneAct.GetValue('ACT_PUVENTECON')/ dCoeffConvert) ;
                          if (TobLigneAct.GetValue('ACT_PUVENTECON')=0) then
                              TobLigneAct.PutValue('ACT_PUVENTECON',ConvertSaisieEF( TobLigneAct.GetValue('ACT_PUVENTE'), false)) ;
                          // TOTVENTE
                          TobLigneAct.PutValue('ACT_TOTVENTE',TobLigneAct.GetValue('ACT_TOTVENTE')/ dCoeffConvert) ;
                          if (TobLigneAct.GetValue('ACT_TOTVENTE')=0) then
                              TobLigneAct.PutValue('ACT_TOTVENTE',TobLigneAct.GetValue('ACT_QTE')*TobLigneAct.GetValue('ACT_PUVENTE')) ;
                          // TOTVENTEDEV
                          TobLigneAct.PutValue('ACT_TOTVENTEDEV',TobLigneAct.GetValue('ACT_TOTVENTEDEV')/ dCoeffConvert) ;
                          if (TobLigneAct.GetValue('ACT_TOTVENTEDEV')=0) then
                              TobLigneAct.PutValue('ACT_TOTVENTEDEV',TobLigneAct.GetValue('ACT_QTE')*TobLigneAct.GetValue('ACT_PUVENTEDEV')) ;
                          // TOTVENTECON
                          TobLigneAct.PutValue('ACT_TOTVENTECON',TobLigneAct.GetValue('ACT_TOTVENTECON')/ dCoeffConvert) ;
                          if (TobLigneAct.GetValue('ACT_TOTVENTECON')=0) then
                              TobLigneAct.PutValue('ACT_TOTVENTECON',TobLigneAct.GetValue('ACT_QTE')*TobLigneAct.GetValue('ACT_PUVENTECON')) ;
                          end;
                      end;
                  end;

          // Mise a jour des donnees liees à la ressource
          // C'est la fonction de mise à jour de la valorisation qui gere l'AFOAssistants
          if (IndexRess<>-1) and (IndexRess<>-2) then
              begin
              TobLigneAct.PutValue('ACT_TYPERESSOURCE', TAFO_Ressource(AFOAssistants.Objects[IndexRess]).tob_Champs.GetValue('ARS_TYPERESSOURCE'));
              try
              ListeFonctionsRess:= TAFO_Ressource(AFOAssistants.Objects[IndexRess]).FonctionDeLaRessource(TobLigneAct.GetValue('ACT_DATEACTIVITE'));
              if (ListeFonctionsRess<>nil) and (ListeFonctionsRess.Count<>0) then
                  TobLigneAct.PutValue('ACT_FONCTIONRES', ListeFonctionsRess[0]);
              finally
              ListeFonctionsRess.Free;
              end;
              end;

          // Mise a jour des donnees liees à l'article
          if (TOBValo<>nil) then TOBArtRef := TOBValo
          else if (TOBArt<>nil) then TOBArtRef := TOBArt;

          if (TOBArtRef<>nil) then
              begin
              TobLigneAct.PutValue('ACT_TYPEARTICLE',TOBArtRef.GetValue('GA_TYPEARTICLE')) ;
              TobLigneAct.PutValue('ACT_ARTICLE',TOBArtRef.GetValue('GA_ARTICLE')) ;
              TobLigneAct.PutValue('ACT_CODEARTICLE',TOBArtRef.GetValue('GA_CODEARTICLE')) ;
              if (TOBArtRef.GetValue('GA_TYPEARTICLE')='PRE') then
                  begin
                  if (TOBArtRef.GetValue('GA_ACTIVITEEFFECT')='') then
                      TobLigneAct.PutValue('ACT_ACTIVITEEFFECT','X')
                  else
                      TobLigneAct.PutValue('ACT_ACTIVITEEFFECT',TOBArtRef.GetValue('GA_ACTIVITEEFFECT')) ;
                  end
              else
                  TobLigneAct.PutValue('ACT_ACTIVITEEFFECT','-');

              if (TobLigneAct.GetValue('ACT_LIBELLE')='') then
                  TobLigneAct.PutValue('ACT_LIBELLE',TOBArtRef.GetValue('GA_LIBELLE')) ;
              if (TobLigneAct.GetValue('ACT_ACTIVITEREPRIS')='') then
                  TobLigneAct.PutValue('ACT_ACTIVITEREPRIS',TOBArtRef.GetValue('GA_ACTIVITEREPRISE')) ;
              end;

          // contrôles complémentaires
          if (TobLigneAct.GetValue('ACT_TYPEACTIVITE')<>'REA') And (TobLigneAct.GetValue('ACT_TYPEACTIVITE')<>'PRE') And (TobLigneAct.GetValue('ACT_TYPEACTIVITE')<>'SIM') then
            TobLigneAct.PutValue('ACT_TYPEACTIVITE','REA');
          if (TobLigneAct.GetValue('ACT_ACTIVITEREPRIS')<>'F') And (TobLigneAct.GetValue('ACT_ACTIVITEREPRIS')<>'FAC') And (TobLigneAct.GetValue('ACT_ACTIVITEREPRIS')<>'N') And (TobLigneAct.GetValue('ACT_ACTIVITEREPRIS')<>'A')then
            TobLigneAct.PutValue('ACT_ACTIVITEREPRIS','F');

          // Gestion automatique du visa sur l'activité
          if (GetParamSoc('SO_AFVISAACTIVITE')=false) then
                  begin
                  TobLigneAct.PutValue('ACT_ETATVISA', 'VIS');
                  TobLigneAct.PutValue('ACT_VISEUR', V_PGI.User);
                  TobLigneAct.PutValue('ACT_DATEVISA', NowH);
                  end
          else
                  begin
                  if TobLigneAct.GetValue('ACT_ETATVISA') = '' then
                     TobLigneAct.PutValue('ACT_ETATVISA', 'ATT');
                  end;

          // Gestion automatique du visa de facturation sur l'activité
          if (GetParamSoc('SO_AFAPPPOINT')=false) then
                  begin
                  TobLigneAct.PutValue('ACT_ETATVISAFAC', 'VIS');
                  TobLigneAct.PutValue('ACT_VISEURFAC', V_PGI.User);
                  TobLigneAct.PutValue('ACT_DATEVISAFAC', NowH);
                  end
          else
                  begin
                  if TobLigneAct.GetValue('ACT_ETATVISAFAC') = '' then
                     TobLigneAct.PutValue('ACT_ETATVISAFAC', 'ATT');
                  end;
      end;  // Boucle sur les lignes d'E-Activité

      // On enregistre les lignes d'activité récupérées et complétées dans la table activite
      TOBActivite.InsertOrUpdateDB(true);

      // Suppression des lignes non importées de la TOBEactivite (pour ne pas les supprimer de la base par DeleteDB)
      i:=0;
      iNbEactivite:=TobEActivite.detail.Count;
      while i <= (iNbEactivite-1) do
      begin
           CodeAssistant := TobEActivite.Detail[i].GetValue ('EAC_RESSOURCE');
           if (TobEActivite.detail[i].GetValue ('VALIDE') = '-') then
           begin
               // Si validité pas encore établie pour cet assistant, on le déplace dans les non valides
{               if NonValideListe.Find (CodeAssistant, j) = FALSE then
               begin
                    if ValideListe.Find (CodeAssistant, j) = TRUE then
                       ValideListe.Delete (j);
                    NonValideListe.Add (CodeAssistant);
               end;}
               TobEActivite.Detail[i].Free;
               iNbEactivite := iNbEactivite-1;
          end
          else
{          begin
               // Si validité pas encore établie pour cet assistant, on le met dans les valides
               if (ValideListe.Find (CodeAssistant, j) = FALSE) AND (NonValideListe.Find (CodeAssistant, j) = FALSE) then
                  ValideListe.Add (CodeAssistant);
}
               i:=i+1;
//          end;
      end;

      // On supprime dans la table EACTIVITE les lignes importées dans la table ACTIVITE
      bDelete := TobEActivite.DeleteDB(True);
      if not bDelete then
      begin
           PGIBoxAF ('Suppression des lignes E-Activités impossible','Import E-Activité');
           AssistListe.Free;
           exit;
      end;
      finally
            TOBValo.Free;
            TOBActivite.Free;
            AFOAssistants.Free;
            TOBArticles.Free;
            ListeDesCles.Free;
//            AssistListe.Free;
//            ValideListe.Free;
  //          NonValideListe.Free;
            if (NomFicLog<>'') then
                begin
                writeln (FichierLog, '');
                writeln (FichierLog, '************ Fin de l''importation des données de l''eactivité ' + FormatDateTime('dd/mm/yyyy ttttt',CurrentDate));
                writeln (FichierLog, '********************************************************************************************************' );
                writeln (FichierLog, '');
                writeln (FichierLog, '');
                writeln (FichierLog, '');
                writeln (FichierLog, '');
                writeln (FichierLog, '');
                writeln (FichierLog, '');
                writeln (FichierLog, '');
                writeln (FichierLog, '');
                writeln (FichierLog, '');
                writeln (FichierLog, '');
                CloseFile (FichierLog);
            end;
      end;

      // En mode suivi de SDA, on màj la table de suivi
      try
      try
          if bSuivi = TRUE then
          begin
               for i := 0 to AssistListe.Count-1 do
               begin
                    if ExisteSql ('SELECT EAC_RESSOURCE FROM EACTIVITE WHERE EAC_RESSOURCE="' + AssistListe [i] + '" AND EAC_DATEACTIVITE>="' + UsDate (EditDateDebut) + '" AND EAC_DATEACTIVITE<="' + UsDate (EditDateFin) + '"') = TRUE then
                        ExecuteSQL ('UPDATE SUIVIEACT SET ASA_ETATEACT = "ES5" WHERE ASA_RESSOURCE="' + AssistListe [i] + '" AND ASA_DATEDEB="' + USDATE(EditDateDebut) + '" AND ASA_DATEFIN="' + USDATE (EditDateFin) + '"')
                    else
                        ExecuteSQL ('UPDATE SUIVIEACT SET ASA_ETATEACT = "ES4" WHERE ASA_RESSOURCE="' + AssistListe [i] + '" AND ASA_DATEDEB="' + USDATE(EditDateDebut) + '" AND ASA_DATEFIN="' + USDATE (EditDateFin) + '"');
               end;
          end;
      except
            PgiInfoAf ('Importation effectuée, mais le suivi de la saisie décentralisée n''a pas pu être mis à jour', Ecran.Caption);
      end;
      finally
            AssistListe.Free;
      end;
end;

{procedure TOF_AFEACTIVITE_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
     Aff    :=THEdit (GetControl ('EAC_AFFAIRE'));
     Aff0   :=THEdit (GetControl ('EAC_AFFAIRE0'));
     Aff1   :=THEdit (GetControl ('EAC_AFFAIRE1'));
     Aff2   :=THEdit (GetControl ('EAC_AFFAIRE2'));
     Aff3   :=THEdit (GetControl ('EAC_AFFAIRE3'));
     Aff4   :=THEdit (GetControl ('EAC_AVENANT'));
     Aff_   :=THEdit (GetControl ('EAC_AFFAIRE_'));
     Aff0_  :=THEdit (GetControl ('EAC_AFFAIRE0_'));
     Aff1_  :=THEdit (GetControl ('EAC_AFFAIRE1_'));
     Aff2_  :=THEdit (GetControl ('EAC_AFFAIRE2_'));
     Aff3_  :=THEdit (GetControl ('EAC_AFFAIRE3_'));
     Aff4_  :=THEdit (GetControl ('EAC_AVENANT_'));
     Tiers  :=THEdit (GetControl ('EAC_TIERS'));
     Tiers_ :=THEdit (GetControl ('EAC_TIERS_'));
end;}

procedure AFLanceFiche_Mul_EActivite;
begin
     AGLLanceFiche ('AFF','AFEACTIVITE_MUL', '', '', '');
end;




Initialization
  registerclasses ( [ TOF_AFEACTIVITE_MUL ] ) ;
end.
