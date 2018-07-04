unit UTofGCCATALOGUE_EPUR;

interface
uses  StdCtrls,Controls,Classes,db,forms,sysutils,dbTables,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, mul, DBGrids, HDimension,UTOM,AGLInit,
      Utob,HDB,Messages,HStatus,Fe_Main,M3VM , M3FP,Fiche,Formule,Hqry,
      EdtEtat, AglInitGC, Graphics;

Type

    TOF_GCCATALOGUE_EPUR = Class (TOF)
        Mouchard: TListBox;
        TOBCata, TOBEtat : TOB;
        procedure OnArgument (Arguments : string) ; override ;
        procedure OnLoad ; override ;
        procedure EpureCatalogue ;
        procedure ImprimEpur;
        procedure EffaceCatalog;
        procedure EffaceEtat;
        procedure ChargeEtat;
        procedure PrepareMouchard;
        procedure CacheDimensions;
        procedure MontreDimensions (Masque : string);
  end;
  var MemArticle : string;
        
procedure TOF_CATALOGUE_EPUR_CatalogueEpure(parms:array of variant; nb: integer ) ;
procedure TOF_CATALOGUE_EPUR_RechCatalogueEpure(parms:array of variant; nb: integer ) ;

const
// libellés des messages
TexteMessage: array[1..5] of string 	= (
          {1}  'Aucun élément sélectionné'
          {2} ,'ATTENTION : Epuration annulée !'
          {3} ,'ATTENTION : Epuration impossible : un autre utilisateur utilise ces données !'
          {4} ,'ERREUR impression stoppée...'
          {5} ,''
              );

implementation
//================================================================================
//---------------------- initialisations
//================================================================================
procedure TOF_GCCATALOGUE_EPUR.OnArgument(Arguments : String) ;
Begin
Inherited ;
    //Cachedimensions;
    // ça ne marche pas à ce niveau donc on oblige le premier OnLoad à le faire
    MemArticle:='z' ;
End ;
//=================================================================================
//---------- le OnLoad est appelé suivant le timer ceci est INDISPENSABLE pour
//---------- la gestion du deuxième Onglet mais trop rapide pour la saisie de la
//---------- date... c'est pourquoi on utilise une date intermédiaire.
//=================================================================================
Procedure TOF_GCCATALOGUE_EPUR.OnLoad  ;
var  F : TFMul ;
     CC :TCheckBox ;
     DD : string;
     i : integer;
     TCombo: THValComboBox;
     QQ : TQuery;
     Statut, Masque : string;
BEGIN
    inherited ;
    if not (Ecran is TFMul) then exit;
    F:=TFMul(Ecran) ;

    // si les articles non références sont choisis seuls pas de sélection d'article
    // + mise en place de la clause where correspondante
    CC:= TCheckBox(TFMul(F).FindComponent('_ARTICLENONREF')) ;
    If CC.State=cbChecked then
    Begin
        SetControlText('GA_CODEARTICLE','') ;
        SetControlEnabled('GA_CODEARTICLE',False) ;
        SetControlProperty('GA_CODEARTICLE','Color',clbtnFace);
        SetControlText('XX_WHERE','GCA_ARTICLE="" or GCA_ARTICLE is null') ;
    end else
    Begin
        SetControlEnabled('GA_CODEARTICLE',True) ;
        SetControlProperty('GA_CODEARTICLE','Color',clWindow);
        If CC.State=cbUnChecked then
        Begin
             SetControlText('XX_WHERE','GCA_ARTICLE<>""');
        End ;
        If CC.State=cbGrayed then
        Begin
             SetControlText('XX_WHERE','');
        End ;
    End ;

    // on ne peut choisir un écart de références que pour un seul fournisseur
    If GetControlText('GCA_TIERS') <> '' then
    begin
        SetControlEnabled('GCA_REFERENCE',True) ;
        SetControlProperty('GCA_REFERENCE','Color',clWindow);
        SetControlEnabled('GCA_REFERENCE_',True) ;
        SetControlProperty('GCA_REFERENCE_','Color',clWindow);
    end else
    begin
        SetControlText('GCA_REFERENCE','') ;
        SetControlEnabled('GCA_REFERENCE',False) ;
        SetControlProperty('GCA_REFERENCE','Color',clbtnFace);
        SetControlText('GCA_REFERENCE_','') ;
        SetControlEnabled('GCA_REFERENCE_',False) ;
        SetControlProperty('GCA_REFERENCE_','Color',clbtnFace);
    end;

    // si on change d'article, on active (ou non) ses dimensions
    if GetControlText('GA_CODEARTICLE') <> MemArticle then
    begin
         If GetControlText('GA_CODEARTICLE') = '' then
         begin
              CacheDimensions;
         end else
         begin
              Statut := '';
              Masque := '';
              QQ := OpenSQL('SELECT GA_ARTICLE, GA_STATUTART, GA_DIMMASQUE FROM ARTICLE WHERE GA_CODEARTICLE="'
                  +GetControlText('GA_CODEARTICLE')+'" AND GA_STATUTART <> "DIM"',TRUE);
              if not QQ.EOF then
              begin
                   Statut := QQ.FindField('GA_STATUTART').AsString;
                   Masque := QQ.FindField('GA_DIMMASQUE').AsString;
              end;
              if Statut <> 'GEN' then CacheDimensions
              else
              begin
                   If Masque = '' then CacheDimensions
                   else MontreDimensions(Masque);
              end;
              Ferme(QQ);
         end;
         MemArticle := GetControlText('GA_CODEARTICLE');
    end;

    // si la date saisie est valide, elle fait partie de la sélection
    DD := GetControlText('_GCA_DATESUP');
    if IsValidDate(DD) then SetControlText('GCA_DATESUP',GetControlText('_GCA_DATESUP'));

    // S'il y en a, on renseigne les dimensions pour le SQL
    for i := 1 to MaxDimension do
    begin
         if GetControlVisible('CODEDIM'+InttoStr(i)) then
         begin
              TCombo := THValComboBox(GetControl('CODEDIM'+InttoStr(i)));
              SetControlText('GA_CODEDIM'+Copy(TCombo.plus,16,1),GetControlText('CODEDIM'+InttoStr(i)));
         end;
    end;
END;

//=========================================================================
//------------------ normalement, on ne voit pas les dimensions -----------
//=========================================================================
procedure TOF_GCCATALOGUE_EPUR.CacheDimensions;
var F : TFMul ;
    i : integer ;
begin
    F := TFMul(Ecran);
    for i := 0 to F.FListe.Columns.Count - 1 do
    begin
         if copy(F.FListe.Columns[i].Title.caption,1,3)='dim' then
            F.FListe.Columns[i].Visible := false;
    end;
    for i := 1 to MaxDimension do
    begin
         SetControlText('CODEDIM'+InttoStr(i),'');
         SetControlText('GA_CODEDIM'+InttoStr(i),'');
         SetControlText('TCODEDIM'+InttoStr(i),'');
         SetControlProperty('CODEDIM'+InttoStr(i),'PLUS','');
         SetControlProperty('CODEDIM'+InttoStr(i),'VISIBLE',False);
    end;
end;
//============================================================================
//--------------------- Montrer les dimensions de l'article choisi -----------
//============================================================================
procedure TOF_GCCATALOGUE_EPUR.MontreDimensions (Masque : string);
var F : TFMul ;
    d, z : integer ;
    QQ : TQuery;
    TCombo: THValComboBox;
begin
    z := 1;
    QQ := OpenSQL('SELECT GDM_TYPE1,GDM_TYPE2,GDM_TYPE3,GDM_TYPE4,GDM_TYPE5 FROM DIMMASQUE WHERE GDM_MASQUE="'+Masque+'"',TRUE) ;
    if QQ.Eof then CacheDimensions
    else
    begin
         for d := 1 to MaxDimension do
         begin
              SetControlText('CODEDIM'+InttoStr(d),'');
              SetControlText('GA_CODEDIM'+InttoStr(d),'');
              SetControlText('TCODEDIM'+InttoStr(d),'');
              SetControlProperty('CODEDIM'+InttoStr(d),'PLUS','');
              SetControlProperty('CODEDIM'+InttoStr(d),'VISIBLE',False);
              if QQ.FindField('GDM_TYPE'+InttoStr(d)).AsString <> '' then
              begin
                   SetControlText('TCODEDIM'+InttoStr(z),RechDom('GCGRILLEDIM'+InttoStr(d),
                       QQ.FindField('GDM_TYPE'+InttoStr(d)).AsString, False));
                   SetControlText('GA_CODEDIM'+InttoStr(d),'');
                   SetControlProperty('CODEDIM'+InttoStr(z),'PLUS','GDI_TYPEDIM="DI'+InttoStr(d)
                       +'" AND GDI_GRILLEDIM="'+QQ.FindField('GDM_TYPE'+InttoStr(d)).AsString
                       +'" ORDER BY GDI_RANG');
                   SetControlProperty('CODEDIM'+InttoStr(z),'VISIBLE',True);
                   TCombo := THValComboBox(GetControl('CODEDIM'+InttoStr(z)));
                   TCombo.ItemIndex := 0;
                   Inc (z);
              end;
          end;
    end;
    ferme(QQ);
end;
//=========================================================================
//---------------- Epuration du catalogue
//=========================================================================
procedure TOF_GCCATALOGUE_EPUR.EpureCatalogue;
var F : TFMul ;
    TOBEcran, TOBED, TOBCD, TOBJnl, TOBTmp : TOB ;
    i_ind, Numjnl, Nbenreg : integer ;
    Where : string ;
    QQ : TQuery ;
    ioerr : TIOErr;

BEGIN
     F:=TFMul(Ecran);
     // si rien de sélectionné on le dit et on s'arrête
     if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
     BEGIN
          //MessageAlerte('Aucun élément sélectionné');
          if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
          HShowMessage('0;'+F.Caption+';'+TexteMessage[1]+';W;O;O;O;','','') ;
          exit;
     END;
     // on mémorise les choix écran pour le journal des évènements
     PrepareMouchard;

     // on charge la même requete qu'à l'écran dans une TOB
     Where := RecupWhereCritere(F.Pages) ;
     TobTmp := TOB.Create ('', Nil, -1) ;
     QQ := OpenSQL('SELECT * FROM EPUCATALOGU ' + Where, True) ;
     if not QQ.EOF then TOBTmp.LoadDetailDB('EPUCATALOGU', '', '', QQ, FAlse)
     else TOBTmp := nil;
     Ferme (QQ) ;
     if TOBTmp = nil then exit ;

     // si on a pas sélectionné tout, on élimine de la TOB tout ce qui n'est pas sélectionné
     if Not F.FListe.AllSelected then
     BEGIN
          TOBEcran := TOB.Create ('EpureCatalogue', Nil, -1) ;
          for i_ind:=0 to F.FListe.NbSelected-1 do
          BEGIN
               TOBED := TOB.Create ('', TOBEcran, -1);
               TOBED.AddChampSup('GCA_TIERS', True) ;
               TOBED.AddChampSup('GCA_REFERENCE', True) ;
               F.FListe.GotoLeBOOKMARK(i_ind);
               TOBED.PutValue ('GCA_TIERS', TFMul(Ecran).Q.FindField('GCA_TIERS').asstring) ;
               TOBED.PutValue ('GCA_REFERENCE', TFMul(Ecran).Q.FindField('GCA_REFERENCE').asstring) ;
          END;
          TOBEcran.Detail.Sort('GCA_TIERS;GCA_REFERENCE') ;
          For i_ind := TOBTmp.Detail.Count-1 downto 0 do
          BEGIN
               TOBCD := TOBTmp.Detail[i_ind] ;
               TOBED := TOBEcran.FindFirst(['GCA_TIERS','GCA_REFERENCE'], [TOBCD.GetValue('GCA_TIERS'),TOBCD.GetValue('GCA_REFERENCE')], False);
               if TOBED=Nil then TOBCD.Free;
          END;
          TOBEcran.free ;
     END;

     // et on crée la TOB définitive
     TOBCata := TOB.Create ('', Nil, -1) ;
     For i_ind := 0 to TOBTmp.Detail.Count-1 do
     BEGIN
        TOBED := TOBTmp.Detail[i_ind] ;
        TOBCD := TOB.Create('CATALOGU', TOBCata, -1) ;
        TOBCD.PutValue('GCA_REFERENCE',TOBED.GetValue('GCA_REFERENCE'));
        TOBCD.PutValue('GCA_TIERS',TOBED.GetValue('GCA_TIERS'));
        TOBCD.PutValue('GCA_ARTICLE',TOBED.GetValue('GCA_ARTICLE'));
        TOBCD.PutValue('GCA_LIBELLE',TOBED.GetValue('GCA_LIBELLE'));
        TOBCD.PutValue('GCA_DATESUP',TOBED.GetValue('GCA_DATESUP'));
     END;
     TOBTmp.free ;

     // on vérifie que le fichier temporaire d'impression est vide pour l'utilisateur
     TobEtat := TOB.Create ('', Nil, -1) ;
     QQ := OpenSQL('SELECT * FROM GCTMPEPUFIC WHERE GZG_UTILISATEUR="' + V_PGI.User +'"', True) ;
     if not QQ.EOF then TOBEtat.LoadDetailDB('GCTMPEPUFIC', '', '', QQ, FAlse)
     else TOBEtat := nil;
     Ferme (QQ) ;

     // première transaction : épurer
     Nbenreg:=TOBCata.detail.count ;
     Numjnl:=0 ;
     TOBJnl:=TOB.Create('JNALEVENT', Nil, -1) ;
     TOBJnl.PutValue('GEV_TYPEEVENT', 'EPU');
     TOBJnl.PutValue('GEV_LIBELLE', F.Caption);
     TOBJnl.PutValue('GEV_DATEEVENT', Date);
     TOBJnl.PutValue('GEV_UTILISATEUR', V_PGI.User);
     QQ:=OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT',True) ;
     if Not QQ.EOF then Numjnl:=QQ.Fields[0].AsInteger ;
     Ferme(QQ) ;
     Inc(Numjnl) ;
     TOBJnl.PutValue('GEV_NUMEVENT', Numjnl);
     ioerr := Transactions (EffaceCatalog,2);
     Case ioerr of
          oeOK :
          BEGIN
               TOBJnl.PutValue('GEV_ETATEVENT', 'OK');
               Mouchard.Items.Add('');
               if Nbenreg = 1 then
                   Mouchard.Items.Add (IntToStr (NbEnreg) + ' Article supprimé')
               else
                   Mouchard.Items.Add (IntToStr (NbEnreg) + ' Articles supprimés');
               TOBJnl.PutValue('GEV_BLOCNOTE',Mouchard.Items.Text );
               ImprimEpur;
          END;
          oeUnknown :
          BEGIN
               if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
               HShowMessage('0;'+F.Caption+';'+TexteMessage[2]+';W;O;O;O;','','') ;
               TOBJnl.PutValue('GEV_ETATEVENT', 'ERR');
          END;
          oeSaisie :
          BEGIN
               if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
               HShowMessage('0;'+F.Caption+';'+TexteMessage[3]+';W;O;O;O;','','') ;
               TOBJnl.PutValue('GEV_ETATEVENT', 'ERR');
          END;
     END;
     TOBJnl.InsertDB(Nil) ;
     TOBJnl.Free ;
     if F.FListe.AllSelected then F.FListe.AllSelected:=False else F.FListe.ClearSelected;
     F.bSelectAll.Down := False ;
     TOBCata.free ;
     TOBEtat.free ;
END;
//==========================================================================
// effacement des lignes sélectionnées : procédure appelée par transactions
//==========================================================================
procedure TOF_GCCATALOGUE_EPUR.EffaceCatalog;
begin
     TOBCata.DeleteDB(false);
end;
//===========================================================================
// impression des lignes effacées
//===========================================================================
procedure TOF_GCCATALOGUE_EPUR.ImprimEpur;
var  F : TFMul ;
     CC:TCheckBox ;
     PP: TPageControl;
     ioerr : TIOErr;
     TOBED, TOBCD : TOB;
     i_ind : integer ;
begin
     // si la case "imprimer" n'est pas cochée, on n'imprime pas
     F:=TFMul(Ecran) ;
     CC:= TCheckBox(TFMul(F).FindComponent('_IMPRIMDET')) ;
     If CC.State=cbUnChecked then exit;

     // Vider, au cas ou, la table temporaire
     if TOBEtat<>nil then
     begin
          ioerr := Transactions (EffaceEtat,2);
          Case ioerr of
               oeOK : ;
               oeUnknown :
               BEGIN
                    if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
                    HShowMessage('0;'+F.Caption+';'+TexteMessage[4]+';W;O;O;O;','','') ;
                    exit;
               END;
               oeSaisie :
               BEGIN
                    if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
                    HShowMessage('0;'+F.Caption+';'+TexteMessage[4]+';W;O;O;O;','','') ;
                    exit;
               END;
          END;
          TOBEtat.free;
     end;
     // charger la Tob Etat avec la Tob Catalogue
     TobEtat := TOB.Create ('', Nil, -1) ;
     for i_ind:=0 to TOBCata.Detail.Count-1 do
     BEGIN
          TOBCD := TOBCata.detail[i_ind];
          TOBED := TOB.Create ('GCTMPEPUFIC', TOBEtat, -1);
          TOBED.PutValue('GZG_UTILISATEUR',V_PGI.User);
          TOBED.PutValue('GZG_COMPTEUR',i_ind+1);
          TOBED.PutValue('GZG_CODE3',TOBCD.GetValue('GCA_REFERENCE'));
          TOBED.PutValue('GZG_CODE1',TOBCD.GetValue('GCA_TIERS'));
          TOBED.PutValue('GZG_CODE2',TOBCD.GetValue('GCA_ARTICLE'));
          TOBED.PutValue('GZG_LIBELLE1',TOBCD.GetValue('GCA_LIBELLE'));
          TOBED.PutValue('GZG_DATE1',TOBCD.GetValue('GCA_DATESUP'));
     end;
     // Ecrire la table temporaire
     ioerr := Transactions (ChargeEtat,2);
     Case ioerr of
          oeOK : ;
          oeUnknown :
          BEGIN
               if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
               HShowMessage('0;'+F.Caption+';'+TexteMessage[4]+';W;O;O;O;','','') ;
               exit;
          END;
          oeSaisie :
          BEGIN
               if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
               HShowMessage('0;'+F.Caption+';'+TexteMessage[4]+';W;O;O;O;','','') ;
               exit;
          END;
     END;
     // Imprimer ...
     PP:= TPageControl(TFMul(F).FindComponent('Pages')) ;
     LanceEtat('E','GEP','CAT',True,False,False,PP,'','',False);

     // Vider la table temporaire
     ioerr := Transactions (EffaceEtat,2);
     Case ioerr of
          oeOK : ;
          oeUnknown :
          BEGIN
               if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
               HShowMessage('0;'+F.Caption+';'+TexteMessage[4]+';W;O;O;O;','','') ;
               exit;
          END;
          oeSaisie :
          BEGIN
               if VAlerte<>Nil then VAlerte.Visible:=FALSE ;
               HShowMessage('0;'+F.Caption+';'+TexteMessage[4]+';W;O;O;O;','','') ;
               exit;
          END;
     END;
end;
//==========================================================================
// effacement de la table temporaire : appelée par transactions
//==========================================================================
procedure TOF_GCCATALOGUE_EPUR.EffaceEtat;
begin
     TOBEtat.DeleteDB(false);
end;
//==========================================================================
// Chargement de la table temporaire : appelée par transactions
//==========================================================================
procedure TOF_GCCATALOGUE_EPUR.ChargeEtat;
begin
     TOBEtat.InsertOrUpdateDB(false);
end;

//===========================================================================
// Chargement des bornes et choix écran dans le mouchard
//===========================================================================
procedure TOF_GCCATALOGUE_EPUR.PrepareMouchard;
var  F : TFMul ;
     CC:TCheckBox ;
     CodeDim : THValComboBox;
     i : integer;
begin
     F:=TFMul(Ecran);
     Mouchard:= TListBox(TFMul(F).FindComponent('MOUCHARD')) ;
     Mouchard.Items.Clear;
     Mouchard.Items.Add (F.Caption);
     Mouchard.Items.Add ('');
     Mouchard.Items.Add ('Date de supression antérieure à ' + GetControlText('GCA_DATESUP'));
     CC:= TCheckBox(TFMul(F).FindComponent('_ARTICLENONREF')) ;
     If CC.State=cbChecked then
     Begin
          Mouchard.Items.Add('Articles non référencés');
     end;
     If CC.State=cbUnChecked then
     Begin
          Mouchard.Items.Add('Articles référencés');
     End ;
     if GetControlText('GCA_TIERS') <> '' then
     begin
        Mouchard.Items.Add ('Fournisseur : ' + GetControlText('GCA_TIERS'));
        if (GetControlText('GCA_REFERENCE') <> '') OR (GetControlText('GCA_REFERENCE_') <> '') then
             Mouchard.Items.Add ('Référence de ' + GetControlText('GCA_REFERENCE') + ' à ' + GetControlText('GCA_REFERENCE_'));
     end;
     if GetControlText('GA_CODEARTICLE') <> '' then
     begin
          Mouchard.Items.Add ('Code article ' + GetControlText('GA_CODEARTICLE'));
          for i := 1 to MaxDimension do
          begin
               if GetControlText('TCODEDIM'+ InttoStr(i)) <> '' then
               begin
                    Codedim := THValComboBox(TFMul(F).FindComponent('CODEDIM'+ Inttostr(i))) ;
                    Mouchard.Items.Add (GetControlText('TCODEDIM'+InttoStr(i))+ ' '+ Codedim.text);
               end;
          end;
     end;
     if F.FListe.AllSelected then Mouchard.Items.Add ('Sélection de la liste entière')
     else
     begin
          if F.FListe.nbSelected = 1 then
              Mouchard.Items.Add (InttoStr(F.FListe.NbSelected) + ' article sélectionné dans la liste')
          else
              Mouchard.Items.Add (InttoStr(F.FListe.NbSelected) + ' articles sélectionnés dans la liste');
     end;
end;
//==========================================================================
/////////////// Procedure appellée par le bouton Validation //////////////
//==========================================================================
procedure TOF_CATALOGUE_EPUR_CatalogueEpure(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
begin
     F:=TForm(Longint(Parms[0])) ;
     if (F is TFmul) then MaTOF:=TFMul(F).LaTOF else exit;
     if (MaTOF is TOF_GCCATALOGUE_EPUR) then TOF_GCCATALOGUE_EPUR(MaTOF).EpureCatalogue
     else exit;
end;
//==========================================================================
/////////////// Procedure appellée par l'elipsis bouton article ////////////
//==========================================================================
procedure TOF_CATALOGUE_EPUR_RechCatalogueEpure(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     G_Article : THCritMaskEdit;
begin
    F:=TForm(Longint(Parms[0])) ;
    if F.Name <> 'GCCATALOGUE_EPUR' then exit;
    G_Article := THCritMaskEdit(F.FindComponent(string(Parms[1])));
    DispatchRecherche(G_Article,1,'','GA_CODEARTICLE='+Trim(Copy(G_Article.Text,1,18)),'');
   if G_Article.Text <> '' then
   begin
       THEdit(F.FindComponent('GA_CODEARTICLE')).Text := trim(Copy(G_Article.Text,1,18));
   end;
end;

//==================================================================================
Initialization
registerclasses([TOF_GCCATALOGUE_EPUR]);
RegisterAglProc('CatalogueEpure',TRUE,1,TOF_CATALOGUE_EPUR_CatalogueEpure);
RegisterAglProc('RechCatalogueEpure',TRUE,1,TOF_CATALOGUE_EPUR_RechCatalogueEpure);
end.
