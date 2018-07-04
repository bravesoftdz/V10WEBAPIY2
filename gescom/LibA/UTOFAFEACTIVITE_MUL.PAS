{***********UNITE*************************************************
Auteur  ...... : PL
Créé le ...... : 31/10/2000
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : AFEACTIVITE_MUL ()
Mots clefs ... : TOF;AFEACTIVITE_MUL
*****************************************************************}
Unit UTOFAFEACTIVITE_MUL ;

Interface

Uses StdCtrls, Controls, Classes,  forms, sysutils,  ComCtrls, HTB97, utob,  windows,
     HCtrls, Ent1, HEnt1, HMsgBox, UTOF, UTofAfBaseCodeAffaire, uafo_ressource,
     AffaireUtil,AfUtilArticle, ActiviteUtil, CBPPath,
     FactUtil, UtilArticle, Dicobtp,Entgc, heureutil, HQry, Hxlspas, FileCtrl, paramsoc
{$IFDEF EAGLCLIENT}
     ,MaineAGL ,  emul
{$ELSE}
     ,DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,HDB ,dbgrids ,Shellapi  ,  mul,
{$ENDIF}

{$IFDEF BTP}
	CalcOleGenericBTP,
{$ENDIF}

    UtobXLS,
    UtilRessource,
    traducaffaire,
    UtofAfEActiviteModif,
    UtofAFEACTCHOIXINTEGR;

Type
  TOF_AFEACTIVITE_MUL = Class (TOF_AFBASECODEAFFAIRE)
    procedure Onupdate; override;
    procedure OnArgument (stArgument:string); override;
    procedure OnClose;                        override;
    procedure BOuvrirOnClick(Sender: TObject);
    procedure BImportOnClick(Sender: TObject);
    procedure BEffaceOnClick(Sender: TObject);
    procedure BEffaceFicLogOnClick(Sender: TObject);
    procedure BFicLogOnClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
    private
    TobImport:TOB;
    FichierLog : textfile;
    EditDateDebut       : THEdit;
    EditDateFin         : THEdit;
    gDateDebutAct, gDateFinAct:TDateTime;
    giNbLigneIntegrees : longint;
    gsQuerySelectIntegr :  string;
    DATEDEBINT, DATEFININT, DATEDEBSUPP, DATEFINSUPP, ORIGINE : string;
    ACTIVITEREPRISE : TStringList;
    bSuivi : boolean;
    function ImportEActivite (TobEActivite:TOB; bAll:boolean):integer;
    //function PutClipboardGrille(laTOB : TOB ) : variant;
    function CompteEActivite:integer;
    function NomFichierLog:string;
    function SupprimeLignesActivite:boolean;
    function TraiteLaQueryImportSelect(var sQ:string; sNewDateDeb, sNewDateFin : string):boolean;
    function IntegreLignesEActivite:integer;

    End;

const
	// libellés des messages
	TexteMessage: array[1..2] of string 	= (
          {1}        ' lignes d''activité ont été importées et '
          {2}       ,' ont été refusées.'
                       );
Procedure AFLanceFiche_Mul_EActivite (starg:string);

implementation


procedure TOF_AFEACTIVITE_MUL.OnArgument(stArgument : String ) ;
var
   ComboTypeArticle :THValComboBox;
   DateDeb          :string;
   DateFin          :string;
   Critere          :string;
   Assistant        :string;
   uneEdit          :THEdit;
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
          uneEdit                := THEdit (GetControl('EAC_RESSOURCE'));
          uneEdit.Text           := Assistant;
          uneEdit.ReadOnly       := TRUE;
          uneEdit.ElipsisButton  := FALSE;
          uneEdit.Enabled        := TRUE;

          uneEdit                := THEdit (GetControl('EAC_RESSOURCE_'));
          uneEdit.Text           := Assistant;
          uneEdit.ReadOnly       := TRUE;
          uneEdit.ElipsisButton  := FALSE;
          uneEdit.Enabled        := TRUE;
     end;

     // JP - gestion suivi ou pas: voir les lignes SDA ou non
     TToolBarButton97 (GetControl('BIMPORTASC')).Visible   := not bSuivi;
     TToolBarButton97 (GetControl('BIMPORTEXCEL')).Visible := not bSuivi;

     // Si suivi SDA, filtres interdit
     TFMul(Ecran).FiltreDisabled := bSuivi;

     // Origine d'e-activité
     if bSuivi = FALSE then
          THValComboBox (GetControl ('EAC_ACTORIGINE')).Plus := 'AND CC_CODE <> "SDE"'
     else
     begin
          THValComboBox (GetControl ('EAC_ACTORIGINE')).Plus := 'AND CC_CODE = "SDE"';
          THValComboBox (GetControl ('EAC_ACTORIGINE')).Value := 'SDE';
          THValComboBox (GetControl ('EAC_ACTORIGINE')).Enabled := FALSE;
          EditDateDebut.Enabled := FALSE;
          EditDateFin.Enabled   := FALSE;
     end;


     // mcd 05/03/02
//     ComboTypeArticle:=THMultiValComboBox(GetControl('EAC_TYPEARTICLE'));
     ComboTypeArticle:=THValComboBox(GetControl('EAC_TYPEARTICLE'));
     ComboTypeArticle.plus := PlusTypeArticle;
 //    if ComboTypeArticle.Text='' then   // gm 20/09/02
 //       ComboTypeArticle.Text := PlusTypeArticleText;

ACTIVITEREPRISE:= TStringList.Create;
end ;

procedure TOF_AFEACTIVITE_MUL.Onupdate;
Var
   bDateDebOK, bDateFinOK,bCherche:boolean;
begin
     inherited;

     bCherche := false;
     {$IFDEF EAGLCLIENT}
     TraduitAFLibGridSt(TFMul(Ecran).FListe);
     {$ELSE}
     TraduitAFLibGridDB(TFMul(Ecran).FListe);
     {$ENDIF}

     bDateDebOK:=true; bDateFinOK:=true;
     if (EditDateDebut.Text='') then
        bDateDebOK:=false;
     if (bDateDebOK and (strtodate(EditDateDebut.Text)<gDateDebutAct)) then
        bDateDebOK:=false;
     if (bDateDebOK=false) then
     begin
          PGIInfoAf('La date de début est antérieure à la date de début d''activité', Ecran.Caption);
          SetFocusControl('EAC_DATEACTIVITE');
          EditDateDebut.Text:=datetostr(gDateDebutAct);
          bCherche := true;
     end;

     if (EditDateFin.Text='') then
        bDateFinOK:=false;
     if (bDateFinOK and ( (strtodate(EditDateFin.Text)>gDateFinAct))) then
        bDateFinOK:=false;
     if (bDateFinOK=false) then
     begin
          PGIInfoAf('La date de fin est postérieure à la date de fin d''activité', Ecran.Caption);
          SetFocusControl('EAC_DATEACTIVITE_');
          EditDateFin.Text:=datetostr(gDateFinAct);
          bCherche := true;
     end;

     if bCherche then
        TFMul(Ecran).ChercheClick ;
end;


procedure TOF_AFEACTIVITE_MUL.FListeDblClick(Sender: TObject);
var
   TobRecup,   TOBArt :TOB;
   CleEntree, CleSortie, Concat:string;
   eDate,eRessource,eAffaire,eAff0,eAff1,eAff2,eAff3,eAff4,eArticle,eTiers:string;
   csTypAct, csAffaire, csRessource, csTypArt:string;
   cdDate:TDateTime;
   ciNumLig:integer;
   sDate,sRessource,sAffaire,sAff0,sAff1,sAff2,sAff3,sAff4,sArticle,sTiers:string;
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
               TOBArt := nil;
               try
                  TOBArt:=TOB.Create('ARTICLE',nil,-1) ;
                  rechart := TrouverArticleSQL_GI(false, sArticle, TOBArt, '');
                  if (rechart <> traAucun) then
                     sTypeArticle := TOBArt.GetValue('GA_TYPEARTICLE');
               finally
                      TOBArt.Free;
               end;
          end;
          // PL fin

          TOBRecup := nil;
          try
             // On récupère l'enregistrement à modifier
             TobRecup := Tob.Create('EACTIVITE',nil,-1);
             Q        := nil;
             try // SELECT * normal de tout prendre ....
                Q := OpenSQL('SELECT * FROM EACTIVITE WHERE EAC_TYPEACTIVITE="'+csTypAct+'" AND EAC_AFFAIRE="'+csAffaire+'" AND EAC_RESSOURCE="'+csRessource+'" AND EAC_DATEACTIVITE="'+UsDateTime(cdDate)+'" AND EAC_TYPEARTICLE="'+csTypArt+'" AND EAC_NUMLIGNE='+inttostr(ciNumLig),True) ;
                if Not TobRecup.SelectDB('', Q) then
                   exit;
             finally
                    Ferme(Q);
             end;

             if (TobRecup<>nil) then
             begin
                  // On modifie l'enregistrement
                  TobRecup.PutValue('EAC_DATEACTIVITE', strtodate(sDate));
                  TobRecup.PutValue('EAC_AFFAIRE', sAffaire);
                  TobRecup.PutValue('EAC_TIERS', sTiers);
                  TobRecup.PutValue('EAC_RESSOURCE', sRessource);
                  TobRecup.PutValue('EAC_CODEARTICLE', sArticle);

                  // JP 15/01/2003 - modification partie 0,1,2,3 code affaire et avenant à partir du code affaire complet
                  {$IFDEF BTP}
                  BTPCodeAffaireDecoupe (sAffaire, sAff0, sAff1, sAff2, sAff3, sAff4, taCreat, FALSE);
                  {$ELSE}
                  CodeAffaireDecoupe (sAffaire, sAff0, sAff1, sAff2, sAff3, sAff4, taCreat, FALSE);
                  {$ENDIF}
                  TobRecup.PutValue('EAC_AFFAIRE0', sAff0);
                  TobRecup.PutValue('EAC_AFFAIRE1', sAff1);
                  TobRecup.PutValue('EAC_AFFAIRE2', sAff2);
                  TobRecup.PutValue('EAC_AFFAIRE3', sAff3);
                  TobRecup.PutValue('EAC_AVENANT',  sAff4);
                  // JP fin

                  // PL le 24/06/02 : correction du type article pour qu'il soit compatible avec l'article
                  TobRecup.PutValue('EAC_TYPEARTICLE', sTypeArticle);
                  // PL fin

                  // On supprime dans la table l'enregistrement à modifier
                  ExecuteSQL ('DELETE FROM EACTIVITE WHERE EAC_TYPEACTIVITE="'+csTypAct+'" AND EAC_AFFAIRE="'+csAffaire+'" AND EAC_RESSOURCE="'+csRessource+'" AND EAC_DATEACTIVITE="'+UsDateTime(cdDate)+'" AND EAC_TYPEARTICLE="'+csTypArt+'" AND EAC_NUMLIGNE='+inttostr(ciNumLig)) ;

                  // On sauvegarde l'enregistrement modifié en prenant soin de ne pas avoir deux fois le meme numéro de ligne
                  Q := nil;
                  try
                     // JP 12/02/2003 - pb de calcul du n° de ligne (parfois supprime un enregistrement existant)
                       //Q := OpenSQL('SELECT MAX(EAC_NUMLIGNE) FROM EACTIVITE WHERE EAC_TYPEACTIVITE="'+csTypAct+'" AND EAC_AFFAIRE="'+sAffaire+'" AND EAC_RESSOURCE="'+sRessource+'" AND EAC_DATEACTIVITE="'+UsDateTime(strtodate(sDate))+'" AND EAC_TYPEARTICLE="'+csTypArt+'" AND EAC_NUMLIGNE='+inttostr(ciNumLig),True);
                     Q := OpenSQL('SELECT MAX(EAC_NUMLIGNE) FROM EACTIVITE WHERE EAC_TYPEACTIVITE="' + csTypAct + '" AND EAC_AFFAIRE="' + sAffaire + '" AND EAC_RESSOURCE="' + sRessource + '" AND EAC_DATEACTIVITE="' + UsDateTime(strtodate(sDate)) + '" AND EAC_TYPEARTICLE="' + csTypArt + '"', True);
                     if not Q.EOF then
                         TobRecup.PutValue('EAC_NUMLIGNE', IntToStr (Q.Fields[0].AsInteger+1))
                     else
                         TobRecup.PutValue('EAC_NUMLIGNE', IntToStr (ciNumLig));
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

          TFMUL (Ecran).ChercheClick;
     end;
end;

procedure TOF_AFEACTIVITE_MUL.BEffaceFicLogOnClick(Sender: TObject);
var     stDocWord : string ;
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
var
    //i:integer;
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
    {i:=}ShellExecute (0, PCHAR('open'),PChar(stDocWord), Nil,Nil,SW_RESTORE);
{$ENDIF}
end;


// Traitement de la tob
// 1ere fille : contient les formats
{function TOF_AFEACTIVITE_MUL.PutClipboardGrille(laTOB : TOB ) : variant;
var TTF,TOBEActivite, TOBDetEActivite : TOB;
    i, j,   NbChampsSup : integer;
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
}

procedure TOF_AFEACTIVITE_MUL.BEffaceOnClick(Sender: TObject);
var
   sClauseWhere  :string;
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

          {$IFDEF EAGLCLIENT}
          if TFMul(Ecran).bSelectAll.Down then
            if not TFMul(Ecran).Fetchlestous then
              begin
              TFMul(Ecran).bSelectAllClick(Nil);
              TFMul(Ecran).bSelectAll.Down := False;
              exit;
              end;
          {$ENDIF}

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
   T                                 :TOB;
   ii, dernieri, NbEactiviteAvant    :integer;
   NbEActiviteApres, NbAffaires      :integer;
   sAff0,sAff1,sAff2,sAff3,sAff4     :string;
   sAffaire,sTiers,sNomRep           :string;
   NomCompletFic                     :string;
   sTitre                            :string;
   SearchRec                         :TSearchRec;
begin
     NomCompletFic:='';

     V_PGI.ZoomOLE := True;   //pour passer la fiche en modal

     // jp 07/12/02: pour titre fenêtre import
     if TControl(Sender).Name = 'BIMPORTEXCEL' then
         sTitre := 'Répertoire d''import fichier EXCEL'
     else
         sTitre := 'Répertoire d''import fichier ASCII';

     // $$$JP: pas de tof pour AFEACTIMPORT
     sNomRep := AGLLanceFiche ('AFF', 'AFEACTIMPORT', '','', sTitre + ';' + GetParamSoc ('SO_AFEACTPATH'));     //rep:= AglLanceFiche( 'AFF', 'AFEACTIVITEREPIMP', '','','') ; //AglLanceFiche( 'AFF', 'AFEACTIVITEREPIMP', '','','') ;
     V_PGI.ZoomOLE := False;

     // Si pas de nom de fichier sélectionné
     if (sNomRep='') then
        exit;

     NbEactiviteAvant:=CompteEActivite;
     //sNomRep := rep, 9, length(Rep)-8);
     if not DirectoryExists(sNomRep) then
     begin
          PGIInfoAf ('Le répertoire: "' + sNomRep + '" n''existe pas', TFMUL(Ecran).caption);
          exit;
     end;

     dernieri := 0;
     if TControl(Sender).Name = 'BIMPORTEXCEL' then // pos('EXCEL', rep)<>0 then
     begin
          FindFirst (sNomRep+'\IMPACT_*.xls', faAnyFile, SearchRec);
          repeat
                if (SearchRec.Name<>'') then
                   if FileExists (sNomRep+'\'+SearchRec.Name) then
                   begin
                        NomCompletFic := sNomRep+'\'+SearchRec.Name;

                        // récupération de la tob
                        T := TOB.Create('', Nil, -1);

                        try
                           // titres de la table dans les titres des colonnes
                           ImportTOBFromXLS (T, NomCompletFic, true, true);

                           if T <> nil then
                           for ii := 0 to T.detail.count-1 do
                           begin
                                if (T.Detail[ii].GetValue('EAC_TYPEACTIVITE')='') then
                                   T.Detail[ii].PutValue('EAC_TYPEACTIVITE', 'REA');
                                if (T.Detail[ii].GetValue('EAC_AFFAIRE')='') then
                                begin
                                     //$$$jp   NbAffaires := 0;
                                     sAff0        := T.Detail[ii].GetValue('EAC_AFFAIRE0');
                                     sAff1        := T.Detail[ii].GetValue('EAC_AFFAIRE1');
                                     sAff2        := T.Detail[ii].GetValue('EAC_AFFAIRE2');
                                     sAff3        := T.Detail[ii].GetValue('EAC_AFFAIRE3');
                                     sAff4        := T.Detail[ii].GetValue('EAC_AVENANT');
                                     sTiers       := T.Detail[ii].GetValue('EAC_TIERS');
                                     sAffaire     := RegroupePartiesAffaire (sAff0, sAff1, sAff2, sAff3, sAff4);
                                     NbAffaires   := TeststCleAffaire (sAffaire, sAff0, sAff1, sAff2, sAff3, sAff4, sTiers, false, false, false, true);

                                     if (NbAffaires=1) then
                                     begin
                                          T.Detail[ii].PutValue ('EAC_AFFAIRE', sAffaire);
                                          T.Detail[ii].PutValue ('EAC_AFFAIRE0', sAff0);
                                          T.Detail[ii].PutValue ('EAC_AFFAIRE1', sAff1);
                                          T.Detail[ii].PutValue ('EAC_AFFAIRE2', sAff2);
                                          T.Detail[ii].PutValue ('EAC_AFFAIRE3', sAff3);
                                          T.Detail[ii].PutValue ('EAC_AVENANT', sAff4);
                                          T.Detail[ii].PutValue ('EAC_TIERS', sTiers);
                                     end;
                                end;

                                Inc (dernieri);
                                T.Detail[ii].PutValue('EAC_NUMLIGNE', NbEactiviteAvant + {ii +} dernieri);

                                // JP: si l'origine manque, il faut la spécifier: e-activité
                                if (T.Detail[ii].GetValue ('EAC_ACTORIGINE') = '') then
                                   T.Detail[ii].PutValue ('EAC_ACTORIGINE', 'EAC');
                           end;

                           T.insertDB (nil, false);
                        finally
                               T.Free;
                        end;
                   end;
          until FindNext(SearchRec)<>0 ;

          sysutils.FindClose (SearchRec);
     end
     else   //if pos('ASCII', rep)<>0 then
     begin
          FindFirst (sNomRep+'\*.asc', faAnyFile, SearchRec);
          repeat
                if (SearchRec.Name<>'') then
                   if FileExists (sNomRep+'\'+SearchRec.Name) then
                      try
                         TobImport := TOB.Create('', nil, -1);
                         NomCompletFic := sNomRep+'\'+SearchRec.Name;
                         if FileExists (NomCompletFic) then
                            TobLoadFromFile( NomCompletFic, Nil, TobImport ) ;
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

          sysutils.FindClose (SearchRec);
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
Q := nil;
try
   Q:=OpenSQL(SQL,True) ;
   if Not Q.EOF then
        begin
        Result:=Q.Fields[0].AsInteger;
        end;
finally
       Ferme(Q);
end;
end;

procedure TOF_AFEACTIVITE_MUL.BOuvrirOnClick(Sender: TObject);
var
Arg, rep: string;
Critere, Champ, valeur  : String;
RepInteg,x : integer;
QQ:TQuery;
begin
RepInteg:=0;
if (TFMul(Ecran).Q.RecordCount=0) then
        begin
//$$$jp        RepInteg:=-2;
        exit;
        end;
DATEDEBINT:=''; DATEFININT:=''; DATEDEBSUPP:=''; DATEFINSUPP:=''; ORIGINE:='';
ACTIVITEREPRISE.Clear;
try            // normal de tout prendre
QQ:=PrepareSQL( 'SELECT * FROM EACTIVITE', True );
RecupWhereSQL( TFMul(Ecran).Q, QQ );
QQ.Open;
gsQuerySelectIntegr:=QQ.SQL.Text;
QQ.Close;

// Si l'import sélectif est choisi dans les paramètres
// et si on n'est pas en suivi de saisie décentralisée : ils sont incompatibles
if (GetParamSoc('SO_AFEACTIMPSELEC')= true) and (bSuivi=false) then
    begin
// PL le 24/06/02 : initialisé directment dans la form
//    Arg := 'DATEDEB:'+ GetControlText('EAC_DATEACTIVITE') + ';DATEFIN:' + GetControlText('EAC_DATEACTIVITE_') + ';';
///////////////////
    rep:= AFLanceFiche_AFEACTCHOIXINTEGR(Arg);

    if (rep='') then
        begin
        RepInteg:=-3;
        exit;
        end;

    Critere:=(Trim(ReadTokenSt(rep)));
    While (Critere <>'') do
        BEGIN
        if Critere<>'' then
            BEGIN
            X:=pos(':',Critere);
            if x<>0 then
               begin
               Champ:=copy(Critere,1,X-1);
               Valeur:=Copy(Critere,X+1,length(Critere)-X);
               end;
            if Champ = 'DATEDEBINT' then DATEDEBINT:=Valeur;
            if Champ = 'DATEFININT' then DATEFININT:=Valeur;
            if Champ = 'DATEDEBSUPP' then  DATEDEBSUPP:=Valeur;
            if Champ = 'DATEFINSUPP' then  DATEFINSUPP:=Valeur;
            if copy(Champ,1,15) = 'ACTIVITEREPRISE' then ACTIVITEREPRISE.Add(Valeur);
            if Champ = 'ORIGINE' then  ORIGINE:=Valeur;
            END;
        Critere:=(Trim(ReadTokenSt(rep)));
        END;

    if Not TraiteLaQueryImportSelect(gsQuerySelectIntegr,DATEDEBINT,DATEFININT) then
        begin
        RepInteg:=-6;
        exit;
        end;

    // Suppression des lignes dans l'activite au prealable suivant les critères
    if (ORIGINE<>'') and (ACTIVITEREPRISE.Count<>0) then
        if Not SupprimeLignesActivite then
            begin
            RepInteg:=-4;
            exit;
            end;
    end
else
    begin
    if (TFMUL(Ecran).FListe.nbSelected=0) and (TFMUL(Ecran).FListe.AllSelected = FALSE) then
    begin
            RepInteg:=-5;
            exit;
    end;
end;

{$IFDEF EAGLCLIENT}
if TFMul(Ecran).bSelectAll.Down then
  if not TFMul(Ecran).Fetchlestous then
    begin
    TFMul(Ecran).bSelectAllClick(Nil);
    TFMul(Ecran).bSelectAll.Down := False;
    exit;
    end;
{$ENDIF}


// Integration des lignes dans l'activité suivant les critères
if (RepInteg=0) then
    RepInteg:=IntegreLignesEActivite;

finally
// Message compte rendu de l'importation
if (RepInteg<0)  then
    begin
    if (RepInteg=-2) then PGIInfoAf('Aucune ligne à intégrer : l''e-activité est vide.', Ecran.Caption)
    else
    if (RepInteg=-3) then PGIInfoAf('Annulation de l''intégration.', Ecran.Caption)
    else
    if (RepInteg=-4) then PGIBoxAf('Une erreur est survenue : les lignes d''activité n''ont pas pu être supprimées au préalable.' + chr(13) + 'Les lignes d''e-activité sélectionnées ne peuvent être intégrées.', Ecran.caption)
    else
    if (RepInteg=-5) then PGIBoxAf('Veuillez sélectionner au-moins une ligne d''e-activité à intégrer.', Ecran.caption)
    else
        PGIInfoAf('Importation impossible. Une erreur bloquante est survenue.', Ecran.Caption);
    end
else
    begin
    if (RepInteg<>0) then
        begin
        PGIInfoAf('Certaines lignes présentent des erreurs. Veuillez consulter le fichier : '+ NomFichierLog, Ecran.Caption);
        end;

    PGIInfoAf(inttostr(giNbLigneIntegrees) + textemessage[1]+ inttostr(RepInteg) + textemessage[2], Ecran.Caption);
    end;
end;

end;


function TOF_AFEACTIVITE_MUL.SupprimeLignesActivite:boolean;
var
sWhere, sReqDel, sActRepMot :string;
i, Rep, iPosWhere, iPosOrderBy, iPosActRepDeb, iPosActRepFin, iNbCar:integer;
begin
Result := false;
iPosWhere := pos ('WHERE', gsQuerySelectIntegr);
sWhere := copy(gsQuerySelectIntegr, iPosWhere+5, Length(gsQuerySelectIntegr));
// Suppression du ORDER BY éventuel
iPosOrderBy:= pos('ORDER', sWhere);
sWhere := copy(sWhere, 1, iPosOrderBy-1);

// Traiter les dates
if Not TraiteLaQueryImportSelect (sWhere, DATEDEBSUPP, DATEFINSUPP) then exit;

// Changement des préfixes et des cotes
sWhere := StringReplace (sWhere, 'EAC_', 'ACT_', [rfReplaceAll,rfIgnoreCase]);
sWhere := StringReplace (sWhere, '''', '"', [rfReplaceAll,rfIgnoreCase]);

// suppression des anciens ACT_ACTORIGINE
// on repere le premier
sActRepMot := 'AND (ACT_ACTORIGINE';
iNbCar := 19;
iPosActRepDeb := pos (sActRepMot, sWhere);
if (iPosActRepDeb = 0) then
  begin
  sActRepMot := 'AND ACT_ACTORIGINE';
  iNbCar := 18;
  iPosActRepDeb := pos ('ACT_ACTORIGINE', sWhere);
  end;
// on repere le dernier
if (iPosActRepDeb <> 0) then
  begin
  iPosActRepFin := iPosActRepDeb + iNbCar;
  while (pos ('ACT_ACTORIGINE', copy (sWhere, iPosActRepFin, length(sWhere)))<>0) do
    iPosActRepFin := iPosActRepFin + pos ('ACT_ACTORIGINE', copy (sWhere, iPosActRepFin, length(sWhere))) + 14;

  if (iNbCar = 19) then
    iPosActRepFin := iPosActRepFin + pos (')', copy (sWhere, iPosActRepFin, length(sWhere)))
  else
    begin
    iPosActRepFin := iPosActRepFin + pos ('"', copy (sWhere, iPosActRepFin, length(sWhere))) + 1;
    iPosActRepFin := iPosActRepFin + pos ('"', copy (sWhere, iPosActRepFin, length(sWhere))) + 1;
    end;

  // On constitue le where
  sWhere := copy (sWhere, 1, iPosActRepDeb - 1) + copy (sWhere, iPosActRepFin, Length(sWhere));
  end;

if (ORIGINE <> '') then
  sWhere := sWhere + ' AND ACT_ACTORIGINE="' + ORIGINE + '"';


if (ACTIVITEREPRISE.Count <>0) then
    begin
    // Ajout des nouveaux ACT_ACTIVITEREPRIS
    sWhere := sWhere + ' AND (';
    i:=0;
    while (i <> ACTIVITEREPRISE.count-1) do
        begin
        sWhere := sWhere + 'ACT_ACTIVITEREPRIS="' + ACTIVITEREPRISE.Strings[i] + '" OR ';
        Inc(i);
        end;

    sWhere := sWhere + 'ACT_ACTIVITEREPRIS="' + ACTIVITEREPRISE.Strings[i] + '")';
    end;


sReqDel := 'DELETE FROM ACTIVITE WHERE' + sWhere;


// Message pour prévenir du danger de cette opération
Rep := PGIAskAF ('Attention, cette opération va supprimer des lignes de l''activité de façon irréversible !' + chr(13)+ 'Confirmez-vous cette suppression ?', Ecran.Caption);
if (Rep = mrYes) then
    begin
    // Suppression dans la table
    EXECUTESQL (sReqDel);
    Result := true;
    end;

end;

function TOF_AFEACTIVITE_MUL.TraiteLaQueryImportSelect(var sQ:string; sNewDateDeb, sNewDateFin : string):boolean;
var
iPosSigneSup, iPosSigneInf, iPosDateDeb, iPosDateFin, iPosFinChaine : integer;
sReste,sReq:string;
sDateDeb, sDateFin : string;
begin
Result:=false;
iPosDateDeb:= pos('EAC_DATEACTIVITE', sQ);
sDateDeb:=sNewDateDeb;
sDateFin:=sNewDateFin;

if (iPosDateDeb <> 0) then
    begin
    sReq := copy(sQ, 1, iPosDateDeb+16);
    sReste := copy(sQ, iPosDateDeb+16, length(sQ));
    iPosSigneSup:= pos('>', sReste);
    iPosDateFin:= pos('EAC_DATEACTIVITE', sReste);
    iPosSigneInf:= pos('<', sReste);
    if (iPosDateFin=0) or (iPosSigneInf=0) or (iPosSigneSup=0) then
        // Erreur
        Exit;
    // test de la cohérence des positions des dates
    if (iPosSigneInf<iPosSigneSup) then
        // Les dates sont inversées, ça ne devrait pas arriver...
        begin
        sDateDeb:=DATEFININT;
        sDateFin:=DATEDEBINT;
        end;

    sReq := sReq + '>=''' + UsDateTime(strtodate(sDateDeb)) + ''' AND EAC_DATEACTIVITE<='''+ UsDateTime(strtodate(sDateFin)) + ''' ';
    sReste := copy(sReste, iPosDateFin+16, length(sReste));
    iPosFinChaine := pos('AND',sReste);
    if (iPosFinChaine=0) then
        iPosFinChaine := pos('OR',sReste);
    if (iPosFinChaine=0) then
        iPosFinChaine := pos(')',sReste);

    if (iPosFinChaine<>0) then
        sReq := sReq + copy(sReste, iPosFinChaine, length(sReste));

    sQ:=sReq;
    Result:=true;
    end
else
    begin
    // Normalement c'est pas possible !!!
    end;
end;

function TOF_AFEACTIVITE_MUL.IntegreLignesEActivite:integer;
var
TobEA, TOBEASelect, TOBLigne :TOB;
i,iNbEactivite:integer;
TypeArticle, TypeActivite, Affaire, Ressource{,sq}:string;
DateActivite:TDateTime;
NumLigne:Integer;
QQ:Tquery;
FListe:THDBGrid;
begin
Result:=-1;
iNbEactivite:=0;
giNbLigneIntegrees:=0;
FListe:=TFMUL(Ecran).FListe;
    // normal de tout prendre
QQ:=PrepareSQL( 'SELECT * FROM EACTIVITE', True );
QQ.SQL.Text:=gsQuerySelectIntegr;
QQ.Open;

TobEA:=TOB.Create('',Nil,-1);
TOBEASelect:=TOB.Create('',Nil,-1);
try
SourisSablier;
// On remplit la tob avec les lignes de l'e-activité
TobEA.LoadDetailDB('EACTIVITE','','',QQ,true) ;
Ferme(QQ);

//if TFMul(Ecran).SD.Execute then ExportGrid(Nil,FListe,TFMul(Ecran).SD.FileName,TFMul(Ecran).SD.FilterIndex,TRUE) ;

if (Fliste.AllSelected) or ((GetParamSoc('SO_AFEACTIMPSELEC')= true) and (bSuivi=false)) then
   BEGIN
   iNbEactivite:=TobEA.Detail.count;

  // Importation dans ACTIVITE, suppression dans EACTIVITE
   Result:=0;
   if (iNbEactivite<>0) then
        Result := ImportEActivite(TobEA, TRUE);

   Fliste.AllSelected:=False;
   TToolBarButton97(GetControl('bSelectAll')).Down:=false;
   END
Else
  Begin
  // Remplissage de la TOB avec les lignes sélectionnées
  for i:=0 to Fliste.NbSelected-1 do
        BEGIN
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

        END;

  iNbEactivite:=TobEASelect.Detail.count;

  // Importation dans ACTIVITE, suppression dans EACTIVITE
  Result:=0;
  if (iNbEactivite<>0) then
        Result := ImportEActivite(TobEASelect, (TOBEASelect.Detail.Count = FListe.DataSource.DataSet.RecordCount));

  Fliste.ClearSelected;
  End;

// Rafraichit la liste
TFMul(Ecran).ChercheClick;

finally
//st := THEdit(GetControl('FICHIER')).Text;
//if (st<>'') then
//    TobEA.SaveToFile(st + '2', false, false, false);
if (Result>=0) then
    giNbLigneIntegrees:= iNbEactivite- Result;
TOBEASelect.Free;
TobEA.Free;
SourisNormale;

end;
end;


function TOF_AFEACTIVITE_MUL.NomFichierLog:string;
var
NomLog,NomFicLog,NomFic1,NomFic2,StPath2,StPath1:string;
begin

Result:='';
NomLog:='';
NomFicLog:='';
StPath2:='';
StPath1:='';
NomFic1:='';
NomFic2:='';
NomLog := THEdit(GetControl('FICHIERLOG')).Text;

if (NomLog<>'') then
    begin
    StPath1:=IncludeTrailingBackSlash(ExtractFilePath(NomLog));
    NomFic1:=ChangeFileExt(ExtractFileName(NomLog),'.log') ;
    end;

if (StPath1<>'') then
    begin
    if (NomFic1<>'') then
        NomFicLog:=StPath1 + NomFic1
    else
        NomFicLog:=StPath1 + 'ImportEActivite.log';
    end;

if (NomFicLog='') then
    begin
    //NomFicLog:='c:\ImportEActivite.log';
    NomFicLog := IncludeTrailingBackSlash(TCBPPath.GetCegidUserTempFileName) + 'ImportEActivite.log';
    end;

THEdit(GetControl('FICHIERLOG')).Text := NomFicLog;
Result:=NomFicLog;
end;

function TOF_AFEACTIVITE_MUL.ImportEActivite (TobEActivite:TOB; bAll:boolean):integer;
  var
    TobActivite, TobLigneAct, TOBArtRef : TOB;
    ChampSansPref, UniteValo, Article : string;
    i, j, indLibChamp, MaxNumLigneCourant : integer;
    AFOAssistants : TAFO_Ressources;
    TOBArticles, TOBValo, TOBArt : TOB;
    IndexRess, iNbEactivite : integer;
//    IndexCle : integer;
    ListeFonctionsRess, ListeDesCles : TStringList;
    Part0, Part1, Part2, Part3, Avenant, Tiers, Affaire : string;
    dCoeffConvert : double;
    bDelete : Boolean;
    NomFicLog : string;
    CodeAssistant : string;
    AssistListe : TStringList;
    bPROK, bPVOK : boolean;
    dQteURef     :double;         // $$$JP 13/05/03
    UniteRef, UniteLigne :string; // $$$JP 13/05/03
begin
  NomFicLog := '';
  Result := 0;

  // gestion du fichier log
  NomFicLog := NomFichierLog;

  if (NomFicLog <> '') then
    begin
      if (THEdit (GetControl ('FICHIERLOG')).Text = '') then
        THEdit (GetControl ('FICHIERLOG')).Text := NomFicLog;

      AssignFile (FichierLog, NomFicLog);
      // on cree ou réouvre le fichier de log
      if (FileExists (NomFicLog) = TRUE) then  Append (FichierLog)
        else  Rewrite (FichierLog);

      writeln (FichierLog, '********************************************************************************************************' );
      writeln (FichierLog, '************ Début de l''importation des données de l''eactivité ' + FormatDateTime('dd/mm/yyyy ttttt',CurrentDate));
      writeln (FichierLog, '' );
    end;

  TobActivite := TOB.Create ('liste activite', Nil, -1);
  TOBArticles := TOB.Create ('Les Articles', Nil, -1);
  AFOAssistants := TAFO_Ressources.Create;
  ListeDesCles := TStringList.Create;
  TobValo := Nil;
  try
    if (TobEActivite.detail.Count >= 0) then
      TobEActivite.Detail[0].AddChampSup ('VALIDE', true);

    // Initialisation liste des assistants valides et non valides pour la période donnée
    AssistListe := TStringList.Create;
    AssistListe.Sorted := TRUE;
    //  ValideListe     := TStringList.Create;
    //  ValideListe.Sorted := TRUE;
    //  NonValideListe  := TStringList.Create;
    //  NonValideListe.Sorted := TRUE;

    // Sur toutes les lignes de l'eactivite, on genere une ligne d'activite en la complétant
    TobValo := nil;
    for i := 0 to TobEActivite.detail.Count - 1 do
      begin
        // On créé la ligne d'activite
        TobLigneAct := TOB.Create ('ACTIVITE', TobActivite, -1);
        TobLigneAct.InitValeurs;

        // La ligne n'est pour l'instant pas importée, donc "non valide" (terme peu approprié)
        TobEActivite.Detail[i].PutValue ('VALIDE', '-');

        // On stocke le code assistant (si pas déjà présent dans la liste), pour pouvoir màj le suivi à la fin
        CodeAssistant := TobEActivite.Detail[i].GetValue ('EAC_RESSOURCE');
        if AssistListe.Find (CodeAssistant, j) = FALSE then
          AssistListe.Add (CodeAssistant);

        // On remplit tous les champs avec les données stockées dans l'eactivite
        if (NomFicLog <> '') then
          begin
            writeln (FichierLog, '' );
            writeln (FichierLog, '************ Importation des données eactivité brutes ligne : '
                    + inttostr (i + 1) + ' ' + FormatDateTime('dd/mm/yyyy ttttt', CurrentDate));
          end;

        for j := 1 to TobEActivite.detail[i].NbChamps do
          begin
            indLibChamp := pos ('_', TobEActivite.detail[i].GetNomChamp (j)) + 1;
            if (indLibChamp = 0) then continue;

            ChampSansPref := copy (TobEActivite.detail[i].GetNomChamp (j), indLibChamp, Length (TobEActivite.detail[i].GetNomChamp (j)) - indLibChamp + 1);
            TobLigneAct.PutValue ('ACT_'+ ChampSansPref, TobEActivite.detail[i].GetValeur (j));
          end;


        // Complément d'initialisation
        if (TobLigneAct.GetValue('ACT_TYPEACTIVITE')='') then
          begin
            if (NomFicLog <> '') then
              writeln (FichierLog, TraduitGA ('Le type d''activité n''était pas saisi => Complément du type d''activité'));
            TobLigneAct.PutValue ('ACT_TYPEACTIVITE', 'REA');
          end;

        IndexRess := AFOAssistants.AddRessource (TobLigneAct.GetValue ('ACT_RESSOURCE'));

        // pas d'articles, on reprend par défaut celui de la ressource
        if (TobLigneAct.GetValue ('ACT_CODEARTICLE') = '') and (IndexRess <> -1) and (IndexRess <> -2) then
          begin
            if (NomFicLog <> '') then
              writeln (FichierLog, TraduitGA ('L''article n''était pas saisi => Complément de l''article par celui de la ressource'));
            TobLigneAct.PutValue ('ACT_CODEARTICLE', Copy (TAFO_Ressource (AFOAssistants.Objects[IndexRess]).tob_Champs.GetValue ('ARS_ARTICLE'), 1, 18));
            TobLigneAct.PutValue ('ACT_TYPEARTICLE', 'PRE');
          end;

        if (TobLigneAct.GetValue('ACT_TYPERESSOURCE') = '') and (IndexRess <> -1) and (IndexRess <> -2) then
          begin
            if (NomFicLog <> '') then
              writeln (FichierLog, TraduitGA ('Le type de la ressource n''était pas saisi => Complément du type de la ressource'));
            TobLigneAct.PutValue ('ACT_TYPERESSOURCE', TAFO_Ressource (AFOAssistants.Objects[IndexRess]).tob_Champs.GetValue ('ARS_TYPERESSOURCE'));
          end;

        if (IndexRess = -1) or (IndexRess = -2) then
          begin
            if (NomFicLog <> '') then
              writeln (FichierLog, TraduitGA ('La ressource ' + TobLigneAct.GetValue('ACT_RESSOURCE') + ' est inconnue'));
            TobLigneAct.PutValue ('ACT_RESSOURCE', '');
          end;

        // Contrôle Affaire  si invalide raz on sort ...
        Part0 := ''; Part1:=''; Part2:=''; Part3 := ''; Avenant:=''; Tiers:='';
        Affaire := TobLigneAct.GetValue ('ACT_AFFAIRE');
        if TeststCleAffaire (Affaire, Part0, Part1, Part2, Part3, Avenant, Tiers, false, false, false, True) = 1 then
          begin
            TobLigneAct.PutValue ('ACT_AFFAIRE', Affaire);
            TobLigneAct.PutValue ('ACT_TIERS', Tiers);
          end
        else
          begin
            TobLigneAct.PutValue ('ACT_AFFAIRE','');
            if (NomFicLog <> '') then
              writeln (FichierLog, TraduitGA ('L''affaire ' + Affaire + ' est inconnue'));
          end;

        TOBArt := TrouverCodeArticle (TobLigneAct.GetValue ('ACT_CODEARTICLE'), TobLigneAct.GetValue ('ACT_TYPEARTICLE'), TOBArticles);
        if ((TobLigneAct.GetValue ('ACT_TYPEARTICLE') = '') and (TOBArt <> nil)) then
          begin
            TobLigneAct.PutValue ('ACT_TYPEARTICLE', TOBArt.GetValue ('GA_TYPEARTICLE'));
            if (NomFicLog <> '') then
              writeln (FichierLog, TraduitGA('Le type de l''article n''était pas saisi => Complément du type d''article'));
          end;

        // on controle la validite des éléments importants de la ligne d'activite
        if  (TOBArt = nil) or (TobLigneAct.GetValue('ACT_TYPEACTIVITE') = '') or (TobLigneAct.GetValue('ACT_AFFAIRE') = '')
              or ((TobLigneAct.GetValue('ACT_RESSOURCE') = '') and (TobLigneAct.GetValue('ACT_TYPEARTICLE') <> 'MAR'))
              or (TobLigneAct.GetValue('ACT_DATEACTIVITE') = 0)
              or (TobLigneAct.GetValue('ACT_DATEACTIVITE') < gDateDebutAct) or (TobLigneAct.GetValue('ACT_DATEACTIVITE') > gDateFinAct)
              or (TobLigneAct.GetValue('ACT_TYPEARTICLE') = '') then
          begin   // ligne non reprise
            // TobEActivite.detail[i].PutValue('VALIDE', '-');
            Result := Result + 1;

            // remplissage du fichier log
            if (NomFicLog <> '') then
              begin
                writeln (FichierLog, '');
                writeln (FichierLog, TraduitGA ('La clé n''est pas complète pour les raisons suivantes :'));

                if (TOBArt = nil) then
                  writeln (FichierLog, TraduitGA ('L''article est vide ou inconnu'));

                if (TobLigneAct.GetValue('ACT_TYPEARTICLE') = '') then
                  writeln (FichierLog, TraduitGA ('Le type d''article est vide ou inconnue'));

                if (TobLigneAct.GetValue ('ACT_TYPEACTIVITE') = '') then
                  writeln (FichierLog, TraduitGA ('Le type d''activité est vide ou inconnu'));

                if (TobLigneAct.GetValue ('ACT_AFFAIRE') = '') then
                  writeln (FichierLog, TraduitGA ('L''affaire est vide ou inconnue'));

                if ((TobLigneAct.GetValue ('ACT_RESSOURCE') = '') and (TobLigneAct.GetValue ('ACT_TYPEARTICLE') <> 'MAR')) then
                  writeln (FichierLog, TraduitGA ('La ressource est vide ou inconnu'));

                if  (TobLigneAct.GetValue ('ACT_DATEACTIVITE') = 0) or (TobLigneAct.GetValue ('ACT_DATEACTIVITE') < gDateDebutAct)
                    or (TobLigneAct.GetValue ('ACT_DATEACTIVITE') >= gDateFinAct) then
                  writeln (FichierLog, TraduitGA('La date d''activité est incorrecte ou n''est pas dans l''intervalle de saisie d''activité'));

                writeln (FichierLog, '');
                writeln (FichierLog, TraduitGA('La ligne n''a pu être importée dans l''activité'));
              end;

            TobLigneAct.Free;
            //$$$jp                  TobLigneAct:=nil;
            continue;
          end;

        TobEActivite.detail[i].PutValue ('VALIDE', 'X');

/////////////////////PL le 14/04/03 : modifs changement de clé dans la table ACTIVITE
        // clé primaire = ACT_TYPEACTIVITE,ACT_AFFAIRE,ACT_NUMLIGNEUNIQUE
        //CleCourante := TobLigneAct.GetValue ('ACT_TYPEACTIVITE') + ';' + TobLigneAct.GetValue ('ACT_AFFAIRE') + ';'
        //               + TobLigneAct.GetValue ('ACT_RESSOURCE') + ';' + DateToStr (TobLigneAct.GetValue ('ACT_DATEACTIVITE')) + ';'
        //               + TobLigneAct.GetValue ('ACT_TYPEARTICLE');
        //CleCourante := TobLigneAct.GetValue ('ACT_TYPEACTIVITE') + ';' + TobLigneAct.GetValue ('ACT_AFFAIRE');
        // IndexCle := ListeDesCles.IndexOf (CleCourante);    PL le 05/05/03 : suppression de la gestion manuelle de la liste des clés

        MaxNumLigneCourant := ProchainIndiceAffaires ( TobLigneAct.GetValue ('ACT_TYPEACTIVITE'), TobLigneAct.GetValue ('ACT_AFFAIRE'),
                                                        ListeDesCles, true);
        //$$$jp          MaxNumLigneCourant:=1;
        (*if (IndexCle = -1) then
          begin
            // On recherche le max pour cette clé
            //MaxNumLigneCourant := MaxNumLigneActivite (TobLigneAct) + 1;
            MaxNumLigneCourant := ProchainMoinsNumLigneUniqueActivite (TobLigneAct.GetValue ('ACT_TYPEACTIVITE'),
                                                                        TobLigneAct.GetValue ('ACT_AFFAIRE'));
            ListeDesCles.AddObject (CleCourante, TObject (MaxNumLigneCourant));
          end
        else
          begin
            //MaxNumLigneCourant := integer (ListeDesCles.Objects[IndexCle]) + 1;
            MaxNumLigneCourant := integer (ListeDesCles.Objects[IndexCle]) - 1;  // on travaille dans les négatifs en eactivite
            ListeDesCles.Objects[IndexCle] := TObject (MaxNumLigneCourant);
          end;*) // PL le 05/05/03 : suppression de la gestion manuelle de la liste des clés
/////////////////////PL le 14/04/03 : fin modifs changement de clé dans la table ACTIVITE

        // on complete la ligne d'activite avec ce qu'on connait
        TobLigneAct.PutValue ('ACT_DEVISE', V_PGI.DevisePivot);
        TobLigneAct.PutValue ('ACT_NUMLIGNE', 20000 + i); // // PL le 15/04/03 modifs changement de clé : Dans ce cas le NUMLIGNE n'a pas trop de sens
        TobLigneAct.PutValue ('ACT_NUMLIGNEUNIQUE', MaxNumLigneCourant); // PL le 15/04/03 modifs changement de clé dans la table ACTIVITE
        TobLigneAct.PutValue ('ACT_FOLIO', 1);
        TobLigneAct.PutValue ('ACT_QTEFAC', TobLigneAct.GetValue ('ACT_QTE'));
        TobLigneAct.PutValue ('ACT_UNITEFAC', TobLigneAct.GetValue ('ACT_UNITE'));
        TobLigneAct.PutValue ('ACT_SEMAINE', NumSemaine (TobLigneAct.GetValue ('ACT_DATEACTIVITE')));
        TobLigneAct.PutValue ('ACT_PERIODE', GetPeriode (TobLigneAct.GetValue ('ACT_DATEACTIVITE')));

        // idem pour les donnees liees à l'affaire
        {$IFDEF BTP}
        BTPCodeAffaireDecoupe (TobLigneAct.GetValue ('ACT_AFFAIRE'), Part0, Part1, Part2, Part3, Avenant, taModif, false);
        {$ELSE}
        CodeAffaireDecoupe (TobLigneAct.GetValue ('ACT_AFFAIRE'), Part0, Part1, Part2, Part3, Avenant, taModif, false);
        {$ENDIF}
        TobLigneAct.PutValue ('ACT_AFFAIRE0', Part0);
        TobLigneAct.PutValue ('ACT_AFFAIRE1', Part1);
        TobLigneAct.PutValue ('ACT_AFFAIRE2', Part2);
        TobLigneAct.PutValue ('ACT_AFFAIRE3', Part3);
        TobLigneAct.PutValue ('ACT_AVENANT', Avenant);

        //
        // On valorise
        //
           // On vide la tob valo
        if (TOBValo <> nil) then
          begin
            TOBValo.Free;
          end;

        bPROK := true;
        bPVOK := true;

        // Valorisation
        TOBValo := MajTOBValo (TobLigneAct.GetValue('ACT_DATEACTIVITE'), tacGlobal, TobLigneAct.GetValue('ACT_AFFAIRE'),
                              TobLigneAct.GetValue('ACT_RESSOURCE'), TobLigneAct.GetValue('ACT_CODEARTICLE'),
                              TobLigneAct.GetValue('ACT_TYPEHEURE'), TOBArticles, nil, nil, AFOAssistants,
                              false, VH_GC.AFValoActPR, VH_GC.AFValoActPV, bPROK, bPVOK);

        if (TOBValo <> nil) then
          begin
            // L'unité de valorisation est l'unité dans laquelle on a valorisé les prix dans la fonction MajTOBValo
            UniteValo := TOBValo.GetValue('GA_QUALIFUNITEACT');
            if (TobLigneAct.GetValue('ACT_UNITE') = '') then TobLigneAct.PutValue ('ACT_UNITE', UniteValo);
            // Si ce n'est pas convertible, on ne touche pas aux données saisies
            if IsConvertible (TobLigneAct.GetValue ('ACT_UNITE'), UniteValo) or (UniteValo = '') then
            begin
                 dCoeffConvert := 1;

                 // On calcule le coefficient de conversion pour passer de l'unité de saisie en unité de valorisation
                 if (UniteValo <> TobLigneAct.GetValue('ACT_UNITE')) and (UniteValo <> '') then
                 begin
                      // Les prix fournis sont en unité de saisie, il faut les convertir en unité de valorisation
                      dCoeffConvert := ConversionUnite ( UniteValo, TobLigneAct.GetValue ('ACT_UNITE'), 1);
                      if dCoeffConvert = 0 then
                         dCoeffConvert := 1;
                 end;

                 // On réajuste la quantité pour qu'elle soit exprimée en unité de valorisation
                 TobLigneAct.PutValue ('ACT_QTE', TobLigneAct.GetValue('ACT_QTE') / dCoeffConvert);

                 // $$$JP 13/05/03: QTEFAC doit être identique à QTE à ce stade
                 TobLigneAct.PutValue ('ACT_QTEFAC', TobLigneAct.GetValue ('ACT_QTE'));
                 // $$$JP fin
                 
                 // On stock l'unite de valorisation
                 TobLigneAct.PutValue ('ACT_UNITE', UniteValo) ;
                 TobLigneAct.PutValue ('ACT_UNITEFAC', UniteValo) ;

                 if (TobLigneAct.GetValue ('ACT_PUPRCHARGE') = 0) then
                 // Si le prix unitaire de revient n'a pas été saisi,
                 // on le valorise et on remplit tous les champs lies
                 begin
                      // PL le 23/01/02 pour gérer le cas où le PU est nul mais pas le Total (else)
                      if (TobLigneAct.GetValue ('ACT_TOTPRCHARGE') = 0) then
                      begin
                           if bPROK then
                           begin
                                TobLigneAct.PutValue ('ACT_PUPRCHARGE', TOBValo.GetValue ('GA_PMRP')) ;
                                TobLigneAct.PutValue ('ACT_PUPR', TOBValo.GetValue ('GA_PMRP')) ;
                                TobLigneAct.PutValue ('ACT_PUPRCHINDIRECT', TOBValo.GetValue ('GA_PMRP')) ;
                                TobLigneAct.PutValue ('ACT_TOTPRCHARGE', TobLigneAct.GetValue ('ACT_QTE') * TobLigneAct.GetValue ('ACT_PUPRCHARGE'));
                                TobLigneAct.PutValue ('ACT_TOTPR', TobLigneAct.GetValue ('ACT_QTE') * TobLigneAct.GetValue ('ACT_PUPR'));
                                TobLigneAct.PutValue ('ACT_TOTPRCHINDI', TobLigneAct.GetValue ('ACT_QTE') * TobLigneAct.GetValue ('ACT_PUPRCHINDIRECT'));
                           end;
                      end
                      else
                      begin  // Si le total n'est pas nul
                           // TOTPRCHARGE
                           TobLigneAct.PutValue ('ACT_TOTPRCHARGE', TobLigneAct.GetValue ('ACT_TOTPRCHARGE') / dCoeffConvert);

                           // TOTPR
                           TobLigneAct.PutValue ('ACT_TOTPR', TobLigneAct.GetValue ('ACT_TOTPR') / dCoeffConvert);
                           if (TobLigneAct.GetValue ('ACT_TOTPR') = 0) then
                              TobLigneAct.PutValue ('ACT_TOTPR', TobLigneAct.GetValue ('ACT_TOTPRCHARGE'));

                           // TOTPRCHINDI
                           TobLigneAct.PutValue ('ACT_TOTPRCHINDI', TobLigneAct.GetValue ('ACT_TOTPRCHINDI') / dCoeffConvert);
                           if (TobLigneAct.GetValue ('ACT_TOTPRCHINDI') = 0) then
                              TobLigneAct.PutValue ('ACT_TOTPRCHINDI', TobLigneAct.GetValue ('ACT_TOTPRCHARGE'));
                      end;
                      // Fin PL le 23/01/02
                 end
                 else  // (TobLigneAct.GetValue('ACT_PUPRCHARGE')=0)
                      // Sinon, on convertit toutes les données saisies en unité de valorisation
                 begin
                    // PUPRCHARGE
                    TobLigneAct.PutValue ('ACT_PUPRCHARGE', TobLigneAct.GetValue ('ACT_PUPRCHARGE') / dCoeffConvert);
                    // PUPR
                    TobLigneAct.PutValue ('ACT_PUPR', TobLigneAct.GetValue ('ACT_PUPR') / dCoeffConvert);
                    if (TobLigneAct.GetValue ('ACT_PUPR') = 0) then
                      TobLigneAct.PutValue ('ACT_PUPR', TobLigneAct.GetValue ('ACT_PUPRCHARGE'));
                    // PUPRCHINDIRECT
                    TobLigneAct.PutValue ('ACT_PUPRCHINDIRECT', TobLigneAct.GetValue ('ACT_PUPRCHINDIRECT') / dCoeffConvert);
                    if (TobLigneAct.GetValue ('ACT_PUPRCHINDIRECT') = 0) then
                      TobLigneAct.PutValue ('ACT_PUPRCHINDIRECT', TobLigneAct.GetValue ('ACT_PUPRCHARGE'));
                    // TOTPRCHARGE
                    TobLigneAct.PutValue ('ACT_TOTPRCHARGE', TobLigneAct.GetValue ('ACT_TOTPRCHARGE') / dCoeffConvert);
                    if (TobLigneAct.GetValue ('ACT_TOTPRCHARGE') = 0) then
                      TobLigneAct.PutValue ('ACT_TOTPRCHARGE', TobLigneAct.GetValue ('ACT_QTE') * TobLigneAct.GetValue ('ACT_PUPRCHARGE'));
                    // TOTPR
                    TobLigneAct.PutValue ('ACT_TOTPR', TobLigneAct.GetValue ('ACT_TOTPR') / dCoeffConvert);
                    if (TobLigneAct.GetValue ('ACT_TOTPR') = 0) then
                      TobLigneAct.PutValue ('ACT_TOTPR', TobLigneAct.GetValue ('ACT_QTE') * TobLigneAct.GetValue ('ACT_PUPR'));
                    // TOTPRCHINDI
                    TobLigneAct.PutValue ('ACT_TOTPRCHINDI', TobLigneAct.GetValue ('ACT_TOTPRCHINDI') / dCoeffConvert);
                    if (TobLigneAct.GetValue ('ACT_TOTPRCHINDI') = 0) then
                      TobLigneAct.PutValue ('ACT_TOTPRCHINDI', TobLigneAct.GetValue ('ACT_QTE') * TobLigneAct.GetValue ('ACT_PUPRCHINDIRECT'));
                 end; // (TobLigneAct.GetValue('ACT_PUPRCHARGE')=0)

                 if (TobLigneAct.GetValue ('ACT_PUVENTE') = 0) then
                  // Si le prix unitaire de vente n'a pas été saisi,
                  // on le valorise et on remplit tous les champs lies
                  begin
                    // PL le 23/01/02 pour gérer le cas où le PU est nul mais pas le Total
                    if (TobLigneAct.GetValue ('ACT_TOTVENTE') = 0) then
                      // Si le total de vente n'a pas été saisi,
                      // on le valorise et on remplit tous les champs lies
                      begin
                        if bPVOK then
                          begin
                            TobLigneAct.PutValue ('ACT_PUVENTE', TOBValo.GetValue ('GA_PVHT'));
                            TobLigneAct.PutValue ('ACT_PUVENTEDEV', TOBValo.GetValue ('GA_PVHT'));
                            TobLigneAct.PutValue ('ACT_TOTVENTE', TobLigneAct.GetValue ('ACT_QTE') * TobLigneAct.GetValue ('ACT_PUVENTE'));
                            TobLigneAct.PutValue ('ACT_TOTVENTEDEV', TobLigneAct.GetValue ('ACT_QTE') * TobLigneAct.GetValue ('ACT_PUVENTEDEV'));
                          end;
                      end
                    else
                      // Sinon, on convertit toutes les données saisies en unité de valorisation
                      begin
                        // TOTVENTE
                        TobLigneAct.PutValue ('ACT_TOTVENTE', TobLigneAct.GetValue ('ACT_TOTVENTE') / dCoeffConvert);
                        // TOTVENTEDEV
                        TobLigneAct.PutValue ('ACT_TOTVENTEDEV', TobLigneAct.GetValue ('ACT_TOTVENTEDEV') / dCoeffConvert);
                        if (TobLigneAct.GetValue ('ACT_TOTVENTEDEV') = 0) then
                          TobLigneAct.PutValue ('ACT_TOTVENTEDEV', TobLigneAct.GetValue ('ACT_TOTVENTE'));
                      end;
                  // Fin PL le 23/01/02
                  end
                else // (TobLigneAct.GetValue('ACT_PUVENTE')=0)
                     // Sinon, on convertit toutes les données saisies en unité de valorisation
                  begin
                    // PUVENTE
                    TobLigneAct.PutValue ('ACT_PUVENTE', TobLigneAct.GetValue ('ACT_PUVENTE') / dCoeffConvert);
                    // PUVENTEDEV
                    TobLigneAct.PutValue ('ACT_PUVENTEDEV', TobLigneAct.GetValue ('ACT_PUVENTEDEV') / dCoeffConvert);
                    if (TobLigneAct.GetValue ('ACT_PUVENTEDEV') = 0) then
                      TobLigneAct.PutValue ('ACT_PUVENTEDEV', TobLigneAct.GetValue ('ACT_PUVENTE'));
                    // TOTVENTE
                    TobLigneAct.PutValue ('ACT_TOTVENTE', TobLigneAct.GetValue ('ACT_TOTVENTE') / dCoeffConvert);
                    if (TobLigneAct.GetValue ('ACT_TOTVENTE') = 0) then
                      TobLigneAct.PutValue ('ACT_TOTVENTE', TobLigneAct.GetValue ('ACT_QTE') * TobLigneAct.GetValue ('ACT_PUVENTE'));
                    // TOTVENTEDEV
                    TobLigneAct.PutValue ('ACT_TOTVENTEDEV', TobLigneAct.GetValue ('ACT_TOTVENTEDEV') / dCoeffConvert);
                    if (TobLigneAct.GetValue ('ACT_TOTVENTEDEV') = 0) then
                      TobLigneAct.PutValue ('ACT_TOTVENTEDEV', TobLigneAct.GetValue ('ACT_QTE') * TobLigneAct.GetValue ('ACT_PUVENTEDEV'));
                  end; //(TobLigneAct.GetValue('ACT_PUVENTE')=0)
              end; // IsConvertible(TobLigneAct.GetValue('ACT_UNITE'), UniteValo) or (UniteValo='')

              // $$$JP 13/05/03: pour une prestation, màj du champ QTEUNITEREF pour avoir la quantité exprimée dans l'unité de référence (paramsoc)
              if TobLigneAct.GetValue ('ACT_TYPEARTICLE') = 'PRE' then
              begin
                   UniteRef   := string (GetParamSoc ('SO_AFMESUREACTIVITE'));
                   UniteLigne := string (TobLigneAct.GetValue('ACT_UNITE'));
                   if UniteRef <> UniteLigne then
                       dQteURef := Arrondi (ConversionUnite (UniteLigne, UniteRef, TobLigneAct.GetValue('ACT_QTE')), GetParamSoc ('SO_DECQTE'))
                   else
                       dQteURef := TobLigneAct.GetValue('ACT_QTE');
                   TobLigneAct.PutValue ('ACT_QTEUNITEREF', dQteURef);
              end;
              // $$$JP fin
          end; // (TOBValo<>nil)

        // Mise a jour des donnees liees à la ressource
        // C'est la fonction de mise à jour de la valorisation qui gere l'AFOAssistants
        if (IndexRess <> -1) and (IndexRess <> -2) then
          begin
            TobLigneAct.PutValue ('ACT_TYPERESSOURCE', TAFO_Ressource (AFOAssistants.Objects[IndexRess]).tob_Champs.GetValue ('ARS_TYPERESSOURCE'));
            ListeFonctionsRess := TAFO_Ressource (AFOAssistants.Objects[IndexRess]).FonctionDeLaRessource (TobLigneAct.GetValue ('ACT_DATEACTIVITE'));
            try
              if (ListeFonctionsRess <> nil) and (ListeFonctionsRess.Count <> 0) then
                TobLigneAct.PutValue ('ACT_FONCTIONRES', ListeFonctionsRess[0]);
            finally
              ListeFonctionsRess.Free;
            end;
          end;

        // Mise a jour des donnees liees à l'article
        if (TOBValo <> nil) then TOBArtRef := TOBValo
        else
          if (TOBArt <> nil) then TOBArtRef := TOBArt
          else TOBArtRef := nil;

        if (TOBArtRef <> nil) then
          begin
            TobLigneAct.PutValue ('ACT_TYPEARTICLE', TOBArtRef.GetValue ('GA_TYPEARTICLE'));
            TobLigneAct.PutValue ('ACT_ARTICLE', TOBArtRef.GetValue ('GA_ARTICLE'));
            TobLigneAct.PutValue ('ACT_CODEARTICLE', TOBArtRef.GetValue ('GA_CODEARTICLE'));
            if (TOBArtRef.GetValue ('GA_TYPEARTICLE') = 'PRE') then
              begin
                if (TOBArtRef.GetValue ('GA_ACTIVITEEFFECT') = '') then
                  TobLigneAct.PutValue ('ACT_ACTIVITEEFFECT', 'X')
                else
                  TobLigneAct.PutValue ('ACT_ACTIVITEEFFECT', TOBArtRef.GetValue ('GA_ACTIVITEEFFECT'));
              end
            else
              TobLigneAct.PutValue ('ACT_ACTIVITEEFFECT', '-');

            if (TobLigneAct.GetValue ('ACT_LIBELLE') = '') then
              TobLigneAct.PutValue ('ACT_LIBELLE', TOBArtRef.GetValue ('GA_LIBELLE'));
            if (TobLigneAct.GetValue ('ACT_ACTIVITEREPRIS') = '') then
              TobLigneAct.PutValue ('ACT_ACTIVITEREPRIS', TOBArtRef.GetValue ('GA_ACTIVITEREPRISE'));
          end;

        // contrôles complémentaires
        if (TobLigneAct.GetValue ('ACT_TYPEACTIVITE') <> 'REA')
            And (TobLigneAct.GetValue ('ACT_TYPEACTIVITE') <> 'PRE')
            And (TobLigneAct.GetValue ('ACT_TYPEACTIVITE') <> 'SIM') then
          TobLigneAct.PutValue ('ACT_TYPEACTIVITE', 'REA');

        if (TobLigneAct.GetValue ('ACT_ACTIVITEREPRIS') <> 'F')
            And (TobLigneAct.GetValue ('ACT_ACTIVITEREPRIS') <> 'FAC')
            And (TobLigneAct.GetValue ('ACT_ACTIVITEREPRIS') <> 'N')
            And (TobLigneAct.GetValue ('ACT_ACTIVITEREPRIS') <> 'A')then
          TobLigneAct.PutValue ('ACT_ACTIVITEREPRIS', 'F');

        // Gestion automatique du visa sur l'activité
        if (GetParamSoc ('SO_AFVISAACTIVITE') = false) then
          begin
            TobLigneAct.PutValue ('ACT_ETATVISA', 'VIS');
            TobLigneAct.PutValue ('ACT_VISEUR', V_PGI.User);
            TobLigneAct.PutValue ('ACT_DATEVISA', NowH);
          end
        else
          begin
            if TobLigneAct.GetValue ('ACT_ETATVISA') = '' then
              TobLigneAct.PutValue ('ACT_ETATVISA', 'ATT');
          end;

        // Gestion automatique du visa de facturation sur l'activité
        if (GetParamSoc ('SO_AFAPPPOINT') = false) then
          begin
            TobLigneAct.PutValue ('ACT_ETATVISAFAC', 'VIS');
            TobLigneAct.PutValue ('ACT_VISEURFAC', V_PGI.User);
            TobLigneAct.PutValue ('ACT_DATEVISAFAC', NowH);
          end
        else
          begin
            if TobLigneAct.GetValue ('ACT_ETATVISAFAC') = '' then
              TobLigneAct.PutValue ('ACT_ETATVISAFAC', 'ATT');
          end;
      end;  // Boucle sur les lignes d'E-Activité

    // On enregistre les lignes d'activité récupérées et complétées dans la table activite
    TOBActivite.InsertDB (nil, true);

    // Suppression des lignes non importées de la TOBEactivite (pour ne pas les supprimer de la base par DeleteDB)
    i := 0;
    iNbEactivite := TobEActivite.detail.Count;
    while i <= (iNbEactivite - 1) do
      begin
        CodeAssistant := TobEActivite.Detail[i].GetValue ('EAC_RESSOURCE');
        if (TobEActivite.detail[i].GetValue ('VALIDE') = '-') then
          begin
            // Si validité pas encore établie pour cet assistant, on le déplace dans les non valides
            {if NonValideListe.Find (CodeAssistant, j) = FALSE then
               begin
                    if ValideListe.Find (CodeAssistant, j) = TRUE then
                       ValideListe.Delete (j);
                    NonValideListe.Add (CodeAssistant);
               end;}
            TobEActivite.Detail[i].Free;
            iNbEactivite := iNbEactivite - 1;
          end
        else
          {begin
            // Si validité pas encore établie pour cet assistant, on le met dans les valides
            if (ValideListe.Find (CodeAssistant, j) = FALSE) AND (NonValideListe.Find (CodeAssistant, j) = FALSE) then
              ValideListe.Add (CodeAssistant);
          }
          i := i + 1;
          //end;
      end;

    // On supprime dans la table EACTIVITE les lignes importées dans la table ACTIVITE
    bDelete := TobEActivite.DeleteDB (True);
    if not bDelete then
      begin
        PGIBoxAF ('Suppression des lignes E-Activités impossible', 'Import E-Activité');
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
    if (NomFicLog <> '') then
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
          for i := 0 to AssistListe.Count - 1 do
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

procedure TOF_AFEACTIVITE_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
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
end;

procedure TOF_AFEACTIVITE_MUL.OnClose;
begin
     inherited;

     ACTIVITEREPRISE.Free;
end;


procedure AFLanceFiche_Mul_EActivite (starg:string);
begin
     AGLLanceFiche ('AFF','AFEACTIVITE_MUL', '', '', stArg);
end;


Initialization
  registerclasses ( [ TOF_AFEACTIVITE_MUL ] ) ;
end.
