unit UtofMBOParamDim;

interface
uses  M3FP,StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}DBGrids, db,
{$ENDIF}
{$IFDEF HDIM}
      utilPgi,
{$ENDIF}
      forms,sysutils,ComCtrls,
      HDB,HCtrls,HEnt1,HMsgBox,UTOF, vierge, HDimension,UTOB, UtilArticle,UDimArticle,
      AglInit,Math,EntGC,ParamSoc;


Const MaxDimChamp = 10;

Type
     TOF_MBOPARAMDIM = Class (TOF)
     private
        insertUserPref : boolean;
        sauveMemorise, detailMemorise : string;
        natureDoc : string ; // ACH, VEN, ..
        naturePiece : string ; // Cde client, Cde fourn., ..
{$IFDEF HDIM}
        AffDimDet, AffDimTotFin : array [1..MaxDimension] of boolean;
{$ENDIF}
        procedure SauveUserPref(ToEnreg : boolean) ;
     public
        initChamp : array[1..MaxDimChamp] of String;
        TobListe,TobAffiche : TOB ;
        GLISTE,GAFFICHE : THGrid ;
        procedure OnLoad ; override ;
        procedure OnClose ; override ;
        procedure OnArgument (Arguments : String ) ; override ;
        procedure OnUpdate ; override ;
        procedure ChangeParamChampDim (numChamp : integer) ;
        procedure ClickFlecheDroite ;
        procedure ClickFlecheGauche ;
        procedure ClickFlecheTous ;
        procedure ClickFlecheAucun ;
        procedure ClickFlecheHaut ;
        procedure ClickFlecheBas ;
        procedure RefreshGrid(posListe,posAffiche : integer) ;
        procedure RefreshBouton ;
     END ;

Const
    BTN_DROIT   = 'DROIT' ;
    BTN_GAUCHE  = 'GAUCHE' ;
    BTN_HAUT    = 'HAUT' ;
    BTN_BAS     = 'BAS' ;
    BTN_TOUS    = 'TOUS' ;
    BTN_AUCUN   = 'AUCUN' ;
    GRD_LISTE   = 'GLISTE' ;
    GRD_AFFICHE = 'GAFFICHE' ;

implementation

procedure TOF_MBOPARAMDIM.OnLoad ;
begin
inherited ;
SetControlProperty('bMemoriser','Enabled',False) ;
end;

procedure TOF_MBOPARAMDIM.OnClose ;
begin
inherited ;
TobListe.Free ;
TobAffiche.Free ;
end;

{Initialise les zones de la fiche de paramétrage de la grille}
Procedure TOF_MBOPARAMDIM.OnArgument(Arguments : String ) ;
var Critere,NomChamp,ComboPlus : string ;
    cpt,ichamp,iListe,iEtab : integer;
    bTrouve,bTypeMasqueUsed : boolean ;
    QQ : TQuery ;
    TobEtab : TOB ;
{$IFDEF HDIM}
    QM : TQuery ;
    iDim,x : integer ;
    bAffDim : boolean ;
    PosDim,stDim,Masque,ChampMul,ValMul : String ;
{$ENDIF}
begin
inherited ;
{$IFNDEF HDIM}
cpt:=0 ;
{$ENDIF}
insertUserPref:=False ;
while Arguments<>'' do
    BEGIN
{$IFDEF HDIM}
    Critere:=Trim(ReadTokenSt(Arguments)) ;
    if Critere<>'' then
        BEGIN
        x:=pos('=',Critere) ;
        if x<>0 then ChampMul:=copy(Critere,1,x-1) ;
        ValMul:=copy(Critere,x+1,length(Critere)) ;
        if ChampMul='NATDOC' then
            BEGIN
            natureDoc:=ValMul ;
            ComboPlus:=natureDoc ;

            // Mise en forme des contrôles de l'écran
            if natureDoc<>NAT_ARTICLE then
                BEGIN
                SetControlVisible('CBDETAIL',False) ;
                if natureDoc<>NAT_STOCK then
                    BEGIN
                    SetControlVisible('P_ETAB',False) ;
                    SetControlVisible('GBOptions',False) ;
                    END ;
                if natureDoc=NAT_STOVTE then
                    BEGIN
                    SetControlVisible('LTYPEMASQUE',False) ;
                    SetControlVisible('TYPEMASQUE',False) ;
                    SetControlVisible('DIMDEPOT',False) ;
                    END ;
                END ;
            if (natureDoc=NAT_ARTICLE) or (natureDoc=NAT_STOCK) or
               (natureDoc=NAT_PROPSAI) then
                BEGIN
                bTypeMasqueUsed:=ExisteSQL('select GMQ_TYPEMASQUE from TYPEMASQUE where GMQ_TYPEMASQUE<>"'+VH_GC.BOTypeMasque_Defaut+'"') ;
                if (natureDoc=NAT_ARTICLE) or (natureDoc=NAT_STOCK) then
                  begin
                  SetControlVisible('P_ETAB',bTypeMasqueUsed);
                  // modif du 12/08/02 par CT pour afficher le bon libellè suivant si on est en multi-dépôts ou non
                  if (VH_GC.GCMultiDepots) then
                     begin
                     SetControlCaption('P_ETAB','Dépôts');
                     SetControlCaption('TGLISTE','&Liste des Dépôts');
                     SetControlCaption('TGAFFICHE','Dépôts &affichés');
                     end;
                  end
                    else SetControlVisible('P_ETAB',False) ;
                SetControlVisible('TYPEMASQUE',bTypeMasqueUsed) ;
                SetControlVisible('LTYPEMASQUE',bTypeMasqueUsed) ;
                SetControlVisible('DIMDEPOT',bTypeMasqueUsed) ;
      //            if not bTypeMasqueUsed then SetControlVisible('GBOptions',False) ;    JCF
                END ;

            // Sélection des valeurs disponibles dans les combos
            if ComboPlus<>'' then
                BEGIN
                for iChamp:=1 to MaxDimChamp do
                    BEGIN
                    NomChamp:='CHAMP'+IntToStr(ichamp) ;
            {$IFDEF EAGLCLIENT}
                    THValComboBox(GetControl(NomChamp)).plus:=ComboPlus;
            {$ELSE}
                    THDBValComboBox(GetControl(NomChamp)).plus:=ComboPlus;
            {$ENDIF}
                    END ;
                END ;
            END
        else if ChampMul='NATPIE' then NaturePiece:=ValMul // nature pièce
        else if ChampMul='SAUVER' then  // Définir par défaut : sauvegarder oui/non
            BEGIN
            if ValMul='?' then
                BEGIN
                insertUserPref:=True ;
                ValMul:='X' ;
                END ;
            SetControlText('CBPARDEFAUT',ValMul) ;
            END
        else if ChampMul='DETAIL'  // Détail : fiche article, afficher toutes les tailles
             then SetControlText('CBDETAIL',ValMul)
        else if ChampMul='TYPMAS'  // Type de masque de présentation
             then SetControlText('TYPEMASQUE',ValMul)
        else if ChampMul='DIMDEP'  // Champ affiché dans la dimension dépôt
             then SetControlText('DIMDEPOT',ValMul)
        else if copy(ChampMul,1,5)='CHAMP' then
            BEGIN
            SetControlText(ChampMul,ValMul) ;
            initChamp[StrToInt(copy(ChampMul,6,length(ChampMul)-5))]:=ValMul ; // Sauvegarde paramétrage initial des champs.
            END
        else if ChampMul='MASQUE' then Masque:=ValMul
        else if ChampMul='AFFDIMTOTFIN' then
            BEGIN
            for iDim:=1 to MaxDimension do AffDimTotFin[iDim]:=StringToCheck(copy(ValMul,iDim,1)) ;
            END
        else if ChampMul='AFFDIMDET' then
            BEGIN
            for iDim:=1 to MaxDimension do AffDimDet[iDim]:=StringToCheck(copy(ValMul,iDim,1)) ;
            END
        END ;
{$ELSE}
    Critere:=Trim(ReadTokenSt(Arguments)) ;
    if cpt=0 then  // nature document
        BEGIN
        natureDoc:=Critere ;
        ComboPlus:=natureDoc ;

        // Mise en forme des contrôles de l'écran
        if natureDoc<>NAT_ARTICLE then
            BEGIN
            SetControlVisible('CBDETAIL',False) ;
            if (natureDoc<>NAT_STOCK) and (natureDoc<>NAT_PROPSAI) then
                BEGIN
                SetControlVisible('P_ETAB',False) ;
                SetControlVisible('GBOptions',False) ;          //JCF
                SetControlVisible('LTYPEMASQUE',False) ;
                SetControlVisible('TYPEMASQUE',False) ;
                SetControlVisible('DIMDEPOT',False) ;
                END ;
            END ;
        if (natureDoc=NAT_ARTICLE) or (natureDoc=NAT_STOCK) or
           (natureDoc=NAT_PROPSAI) then
            BEGIN
            //zzzz bTypeMasqueUsed:=ExisteSQL('select CC_CODE from CHOIXCOD where CC_TYPE="GTQ" and CC_CODE<>"'+VH_GC.BOTypeMasque_Defaut+'"') ;
            bTypeMasqueUsed:=ExisteSQL('select GMQ_TYPEMASQUE from TYPEMASQUE where GMQ_TYPEMASQUE<>"'+VH_GC.BOTypeMasque_Defaut+'"') ;
            if (natureDoc=NAT_ARTICLE) or (natureDoc=NAT_STOCK) then
               begin
               SetControlVisible('P_ETAB',bTypeMasqueUsed);
               // modif du 12/08/02 par CT pour afficher le bon libellè suivant si on est en multi-dépôts ou non
               if (VH_GC.GCMultiDepots) then
                  begin
                  SetControlCaption('P_ETAB','Dépôts');
                  SetControlCaption('TGLISTE','&Liste des Dépôts');
                  SetControlCaption('TGAFFICHE','Dépôts &affichés');
                  end;
               end
            else SetControlVisible('P_ETAB',False) ;
            SetControlVisible('TYPEMASQUE',bTypeMasqueUsed) ;
            SetControlVisible('LTYPEMASQUE',bTypeMasqueUsed) ;
            SetControlVisible('DIMDEPOT',bTypeMasqueUsed) ;
            if (not bTypeMasqueUsed) and (natureDoc<>NAT_ARTICLE) then SetControlVisible('GBOptions',False) ;
            END ;

        // Sélection des valeurs disponibles dans les combos
        if ComboPlus<>'' then
            BEGIN
            // FFO : Les champs ne sont sélectionnés que si FFO est présent dans le champ CO_LIBRE (tablette GCDIMCHAMP)
            if ctxFO in V_PGI.PGIContexte then ComboPlus:=ComboPlus + '%" and CO_LIBRE like "%FFO' ;
            for iChamp:=1 to MaxDimChamp do
                BEGIN
                NomChamp:='CHAMP'+IntToStr(ichamp) ;
{$IFDEF EAGLCLIENT}
                THValComboBox(GetControl(NomChamp)).plus:=ComboPlus;
{$ELSE}
                THDBValComboBox(GetControl(NomChamp)).plus:=ComboPlus;
{$ENDIF}
                END ;
            END ;
        END else if cpt=1 then NaturePiece:=Critere // nature pièce
        else if cpt=2 then  // Définir par défaut : sauvegarder oui/non
        BEGIN
        if Critere='?' then
            BEGIN
            insertUserPref:=True ;
            Critere:='X' ;
            END ;
        SetControlText('CBPARDEFAUT',Critere) ;
        END else if cpt=3 then  // Détail : fiche article, afficher toutes les tailles
        BEGIN
        SetControlText('CBDETAIL',Critere) ;
        END else if cpt=4 then  // Type de masque de présentation
        BEGIN
        SetControlText('TYPEMASQUE',Critere) ;
        END else if cpt=5 then  // Champ affiché dans la dimension dépôt
        BEGIN
        SetControlText('DIMDEPOT',Critere) ;
        END else if cpt<16 then
        BEGIN
        SetControlText('CHAMP'+inttostr(cpt-5),Critere) ;
        initChamp[cpt-5]:=Critere ; // Sauvegarde paramétrage initial des champs.
        END ;
    inc(cpt) ;
{$ENDIF}
    END ;


{$IFDEF HDIM}
// Lecture du masque de l'article pour mise en forme de l'écran de paramétrage
QM:=OpenSQL('select GDM_TYPE1,GDM_POSITION1,GDM_TYPE2,GDM_POSITION2,GDM_TYPE3,GDM_POSITION3,'+
                   'GDM_TYPE4,GDM_POSITION4,GDM_TYPE5,GDM_POSITION5 '+
            'from DIMMASQUE '+
            'where GDM_MASQUE="'+Masque+
            '" and GDM_TYPEMASQUE="'+VH_GC.BOTypeMasque_Defaut+'"', True) ;
if not QM.Eof then
    BEGIN
    for iDim:=1 to MaxDimension do
         BEGIN
         stDim:=IntToStr(iDim) ;
         PosDim:=QM.FindField('GDM_POSITION'+stDim).AsString ;
         bAffDim:=(PosDim<>'') ;
         if bAffDim then SetControlText('CBTOTDI'+stDim,RechDom('GCCATEGORIEDIM',PosDim,True)) ;
         SetControlVisible('CBTOTDI'+stDim,bAffDim) ;
         SetControlText('CBTOTDI'+stDim,CheckToString(AffDimTotFin[iDim])) ;
         SetControlVisible('CBDETDI'+stDim,bAffDim) ;
         SetControlText('CBDETDI'+stDim,CheckToString(AffDimDet[iDim])) ;
         END ;
    END ;
Ferme(QM) ;
{$ENDIF}


if NaturePiece='' then NaturePiece:=NatureDoc ;
// Paramétrage des établissements uniquement en contexte MODE pour fiche article et stock.
if (natureDoc<>NAT_ARTICLE) and (natureDoc<>NAT_STOCK) then exit ;
// modif du 12/08/02 par CT ici c'est la liste des dépôts que l'on veut et non celle des établissements
// Chargement de la liste des établissements et du paramétrage existant.
TobListe:=TOB.CREATE('Liste établissements',NIL,-1) ;
TobAffiche:=TOB.CREATE('Etablissements affichés',NIL,-1) ;
QQ:=OpenSQL('select GDE_DEPOT,GDE_LIBELLE,GDE_ABREGE from DEPOTS order by GDE_DEPOT',True) ;
// Chargement paramétrage sauvegardé
TobEtab:=LaTOB ;
if TobEtab<>nil then
    BEGIN
    // Chargement de la liste des établissements dans TobListe
    // et bascule des établissements de TobEtab dans TobAffiche trié suivant l'ordre définit dans TobEtab
    TobListe.LoadDetailDB('DEPOTS','','',QQ,False) ;
    iEtab:=0 ;
    while iEtab<TobEtab.Detail.Count do
        BEGIN
        iListe:=0 ; bTrouve:=False ;
        while (not bTrouve) and (iListe<TobListe.Detail.Count) do
            BEGIN
            bTrouve:=(TobListe.Detail[iListe].GetValue('GDE_DEPOT')
                      =TobEtab.Detail[iEtab].GetValue('GUD_ETABLISSEMENT')) ;
            inc(iListe) ;
            END ;
        if bTrouve then TobListe.Detail[iListe-1].ChangeParent(TobAffiche,-1) ;
        inc(iEtab) ;
        END ;
    END else
    BEGIN
    // Paramétrage non prédéfini, chargement dans TobAffiche trié par code établissement
    TobAffiche.LoadDetailDB('DEPOTS','','',QQ,False) ;
    END ;
Ferme(QQ) ;
// Affichage des tobs
GLISTE:=THGrid(GetControl('GLISTE')) ;
GAFFICHE:=THGrid(GetControl('GAFFICHE')) ;
TobAffiche.PutGridDetail(GAFFICHE,False,False,'GDE_DEPOT;GDE_LIBELLE',True) ;
TobListe.PutGridDetail(GLISTE,False,False,'GDE_DEPOT;GDE_LIBELLE',True) ;
GLISTE.ColWidths[0]:=30 ; GLISTE.ColAligns[0]:=taCenter ;
GLISTE.ColWidths[1]:=208 ; GLISTE.ColAligns[1]:=taLeftJustify ;
GAFFICHE.ColWidths[0]:=30 ; GAFFICHE.ColAligns[0]:=taCenter ;
GAFFICHE.ColWidths[1]:=208 ; GAFFICHE.ColAligns[1]:=taLeftJustify ;
end ;

procedure TOF_MBOPARAMDIM.RefreshGrid(posListe,posAffiche : integer) ;
begin
TobAffiche.PutGridDetail(GAFFICHE,False,False,'GDE_DEPOT;GDE_LIBELLE',True) ;
TobListe.PutGridDetail(GLISTE,False,False,'GDE_DEPOT;GDE_LIBELLE',True) ;
GAFFICHE.Row:=Min(posAffiche,GAFFICHE.RowCount-1) ;
GLISTE.Row:=Min(posListe,GLISTE.RowCount-1) ;
RefreshBouton ;
end ;

// Boutons enable / disable
procedure TOF_MBOPARAMDIM.RefreshBouton ;
begin
SetControlEnabled('BFLECHEDROITE',TobListe.Detail.Count>0) ;
SetControlEnabled('BFLECHEGAUCHE',TobAffiche.Detail.Count>0) ;
SetControlEnabled('BFLECHEHAUT',GAFFICHE.Row>0) ;
SetControlEnabled('BFLECHEBAS',GAFFICHE.Row<GAFFICHE.RowCount-1) ;
SetControlEnabled('BFLECHETOUS',TobListe.Detail.Count>0) ;
SetControlEnabled('BFLECHEAUCUN',TobAffiche.Detail.Count>0) ;
end ;

procedure TOF_MBOPARAMDIM.ClickFlecheDroite ;
var indiceFille : integer ;
begin
// Y a t il quelque chose de sélectionné ?
if GLISTE.Row<0 then exit ;
// Changement du parent de l'élément de la liste des établissements
if TobAffiche.Detail.Count>0 then indiceFille:=GAFFICHE.Row+1 else indiceFille:=0 ;
TobListe.detail[GLISTE.Row].ChangeParent(TobAffiche,indiceFille) ;
RefreshGrid(GLISTE.Row,GAFFICHE.Row+1) ;
end ;

procedure TOF_MBOPARAMDIM.ClickFlecheGauche ;
var indiceFille : integer ;
begin
// Y a t il quelque chose de sélectionné ?
if GAFFICHE.Row<0 then exit ;
// Changement du parent de l'élément des établissements affichés
if TobListe.Detail.Count>0 then indiceFille:=GLISTE.Row+1 else indiceFille:=0 ;
TobAffiche.detail[GAFFICHE.Row].ChangeParent(TobListe,indiceFille) ;
RefreshGrid(GLISTE.Row+1,GAFFICHE.Row) ;
end ;

procedure TOF_MBOPARAMDIM.ClickFlecheTous ;
var indiceFille,iGrd,posListe : integer ;
begin
if GLISTE.RowCount<1 then exit ;
// Changement du parent de l'élément de la liste des établissements
if GAFFICHE.RowCount>1 then indiceFille:=GAFFICHE.Row+1 else indiceFille:=0 ;
posListe:=TobListe.detail.count-1 ;
for iGrd:=0 to posListe do TobListe.detail[0].ChangeParent(TobAffiche,indiceFille+iGrd) ;
RefreshGrid(0,indiceFille+posListe) ;
end ;

procedure TOF_MBOPARAMDIM.ClickFlecheAucun ;
var indiceFille,iGrd,posAffiche : integer ;
begin
if GAFFICHE.RowCount<1 then exit ;
// Changement du parent de l'élément de la liste des établissements
if GLISTE.RowCount>1 then indiceFille:=GLISTE.Row+1 else indiceFille:=0 ;
posAffiche:=TobAffiche.detail.count-1 ;
for iGrd:=0 to posAffiche do TobAffiche.detail[0].ChangeParent(TobListe,indiceFille+iGrd) ;
RefreshGrid(indiceFille+posAffiche,0) ;
end ;

procedure TOF_MBOPARAMDIM.ClickFlecheHaut ;
begin
if GAFFICHE.Row<1 then exit ;
// Changement de l'indice dans la Tob parent
TobAffiche.detail[GAFFICHE.Row].ChangeParent(TobAffiche,GAFFICHE.Row-1) ;
RefreshGrid(GLISTE.Row,GAFFICHE.Row-1) ;
end ;

procedure TOF_MBOPARAMDIM.ClickFlecheBas ;
begin
if GAFFICHE.Row>GAFFICHE.RowCount-2 then exit ;
// Changement de l'indice dans la Tob parent
TobAffiche.detail[GAFFICHE.Row].ChangeParent(TobAffiche,GAFFICHE.Row+1) ;
RefreshGrid(GLISTE.Row,GAFFICHE.Row+1) ;
end ;

{Active le bouton "Memoriser" si une zone a été modifiée}
Procedure AGLOnChangeParamChampDim (Parms: array of variant; nb: integer) ;
var F : TForm;
    TOTOF : TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge)
    then TOTOF := TFVierge(F).LaTOF
    else exit;
  if (TOTOF is TOF_MBOPARAMDIM) then TOF_MBOPARAMDIM(TOTOF).ChangeParamChampDim(integer(Parms[1])) else exit;
end;

Procedure TOF_MBOPARAMDIM.ChangeParamChampDim (numChamp : integer);
var F : TFVierge;
    bMemoriser : TButton;
    st : String;
begin
  F:=TFVierge(Ecran);
  bMemoriser := TButton(F.FindComponent('bMemoriser'));
  if numChamp < 11 then
  begin
    st := TComboBox(F.FindComponent('CHAMP'+inttostr(numChamp))).Text;
    if (not bMemoriser.Enabled) and (initChamp[numChamp] <> st)
      then bMemoriser.Enabled := True;
  end else if numChamp = 11 then
  begin
    if TCheckBox(F.FindComponent('CBPARDEFAUT')).State <> cbChecked then
    begin
      bMemoriser.Visible := True;
      if not bMemoriser.Enabled then bMemoriser.Enabled := True;
    end else
    begin
      bMemoriser.Visible := False;
    end;
  end else if numChamp = 12 then
  begin
    if bMemoriser.Visible then bMemoriser.Enabled := True;
  end;
end;

{Sauvegarde des champs sélectionnés dans les préférences utilisateurs}
procedure TOF_MBOPARAMDIM.SauveUserPref(ToEnreg : Boolean) ;
var stReq,stSql,ValMul,stTypeMasque,stDimDepot : string ;
    iChamp,iAffiche,iListe : integer ;
    TobFille : TOB ;
    bTrouve : boolean ;
{$IFDEF HDIM}
    iDim : integer ;
{$ENDIF}
begin
if ToEnreg then
    BEGIN
    if not insertUserPref then // Delete anciennes préférences
        BEGIN
        stSql:='delete from USERPREFDIM where GUD_NATUREPIECE="'+naturePiece+'" and GUD_UTILISATEUR="'+V_PGI.User+'"' ;
        ExecuteSQL(stSql) ;
        END ;

    // Svgde des zones détail,défaut et type masque : position = 0
    sauveMemorise:=GetControlText('CBPARDEFAUT') ;
    if natureDoc=NAT_ARTICLE then detailMemorise:=GetControlText('CBDETAIL') ;
    if GetParamSoc('SO_GCARTTYPEMASQUE') then
        BEGIN
        stTypeMasque:=GetControlText('TYPEMASQUE') ;
        stDimDepot:=GetControlText('DIMDEPOT')
        END else
        BEGIN
        stTypeMasque:='' ;
        stDimDepot:=''
        END ;
    stSql:='insert into USERPREFDIM (GUD_NATUREPIECE,GUD_UTILISATEUR,GUD_POSITION,'+
        'GUD_CHAMP,GUD_DETAIL,GUD_DEFAUT,GUD_TYPEMASQUE) values ("'+
        naturePiece+'","'+V_PGI.User+'",0,"'+stDimDepot+'","'+
        detailMemorise+'","'+sauveMemorise+'","'+stTypeMasque+'")' ;
    ExecuteSQL(stSql) ;

    // Svgde des 1..10 champs affichés dans le THDim : position = 1..10
    stReq:='insert into USERPREFDIM (GUD_NATUREPIECE,GUD_UTILISATEUR,GUD_POSITION,GUD_CHAMP,'+
        'GUD_LIBELLE) values ("'+naturePiece+'","'+V_PGI.User+'",' ;
    iChamp:=1 ;
    ValMul:=GetControlText('CHAMP'+IntToStr(iChamp)) ;
    // DCA - FQ MODE 11039 - Sauvegarde du 1er au 10ème champ
    while (ValMul<>'') and (iChamp<11) do
        BEGIN
        stSql:=stReq+IntToStr(iChamp)+',"'+ValMul+'","'+RechDom('GCDIMCHAMP',ValMul,False)+'")' ;
        ExecuteSQL(stSql) ;
        inc(iChamp) ;
        if iChamp < 11 then ValMul := GetControlText('CHAMP' + IntToStr (iChamp)) ;
        END ;

{$IFDEF HDIM}
    // Svgde des indicateurs de totaux
    stReq:='insert into USERPREFDIM (GUD_NATUREPIECE,GUD_UTILISATEUR,GUD_POSITION,GUD_DETAIL)'+
           ' values ("'+naturePiece+'","'+V_PGI.User+'",' ;
    for iDim:=1 to MaxDimension do
        BEGIN
        stSql:=stReq+IntToStr(iDim+15)+',"'+GetControlText('CBTOTDI'+IntToStr(iDim))+'")' ;
        ExecuteSQL(stSql) ;
        stSql:=stReq+IntToStr(iDim+20)+',"'+GetControlText('CBDETDI'+IntToStr(iDim))+'")' ;
        ExecuteSQL(stSql) ;
        END ;
{$ENDIF}

    END ; // if ToEnreg then

if (natureDoc=NAT_ARTICLE) or (natureDoc=NAT_STOCK) then
    BEGIN
    LaTob.free ; LaTob:=nil ;
    // Svgde des établissements affichés en cconsultation multi-dépôts : position > 100
    stReq:='insert into USERPREFDIM (GUD_NATUREPIECE,GUD_UTILISATEUR,GUD_POSITION,'+
        'GUD_ETABLISSEMENT) values ("'+
        naturePiece+'","'+V_PGI.User+'",' ;
    iChamp:=101 ;
    if (TobAffiche=nil) or (TobAffiche.Detail.count<1) then
        BEGIN // Ajout du dépôt par défaut
        if TobAffiche=nil then TobAffiche:=TOB.CREATE('Etablissements affichés',NIL,-1) ;
        iListe:=0 ; bTrouve:=False ;
        while (not bTrouve) and (iListe<TobListe.Detail.Count) do
            BEGIN
            bTrouve:=(TobListe.Detail[iListe].GetValue('GDE_DEPOT')=VH_GC.GCDepotDefaut) ;
            inc(iListe) ;
            END ;
        if bTrouve then TobListe.Detail[iListe-1].ChangeParent(TobAffiche,-1) ;
        END ;
    if TobAffiche<>nil then for iAffiche:=0 to TobAffiche.Detail.count-1 do
        BEGIN
        if ToEnreg then
            BEGIN
            stSql:=stReq+IntToStr(iChamp)+',"'
               +TobAffiche.Detail[iAffiche].GetValue('GDE_DEPOT')+'")' ;
            inc(iChamp) ;
            ExecuteSQL(stSql) ;
            END ;
        if LaTob=nil then LaTob:=TOB.Create('Liste établissements',nil,-1) ;
        TobFille:=TOB.Create('',LaTob,-1) ;
        TobFille.AddChampSup('GUD_ETABLISSEMENT',False) ;
        TobFille.AddChampSup('GDE_DEPOT',False) ;
        TobFille.AddChampSup('GDE_LIBELLE',False) ;
        TobFille.AddChampSup('GDE_ABREGE',False) ;
        TobFille.PutValue('GUD_ETABLISSEMENT',TobAffiche.Detail[iAffiche].GetValue('GDE_DEPOT')) ;
        TobFille.PutValue('GDE_DEPOT',TobAffiche.Detail[iAffiche].GetValue('GDE_DEPOT')) ;
        TobFille.PutValue('GDE_LIBELLE',TobAffiche.Detail[iAffiche].GetValue('GDE_LIBELLE')) ;
        TobFille.PutValue('GDE_ABREGE',TobAffiche.Detail[iAffiche].GetValue('GDE_ABREGE')) ;
        END ;
    TheTob:=LaTob ; // TheTob maj pour retour appel fiche ARTDIM_PARAM
    END ;
end ;

procedure TOF_MBOPARAMDIM.OnUpdate ;
begin
SauveUserPref(Boolean(GetControlText('CBPARDEFAUT')='X')) ;
end;

procedure AGLOnMemoriser (Parms: array of variant; nb: integer) ;
var F : TForm;
    TOTOF : TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge)
    then TOTOF := TFVierge(F).LaTOF
    else exit;
  if (TOTOF is TOF_MBOPARAMDIM) then TOF_MBOPARAMDIM(TOTOF).SauveUserPref(True) else exit;
end;

procedure AGLOnClickBouton (Parms: array of variant; nb: integer) ;
var F : TForm ;
    TOTOF : TOF ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then TOTOF:=TFVierge(F).LaTOF else exit ;
if (TOTOF is TOF_MBOPARAMDIM) then
    BEGIN
    if Parms[1]=BTN_DROIT then TOF_MBOPARAMDIM(TOTOF).ClickFlecheDroite
    else if Parms[1]=BTN_GAUCHE then TOF_MBOPARAMDIM(TOTOF).ClickFlecheGauche
    else if Parms[1]=BTN_HAUT then TOF_MBOPARAMDIM(TOTOF).ClickFlecheHaut
    else if Parms[1]=BTN_BAS then TOF_MBOPARAMDIM(TOTOF).ClickFlecheBas
    else if Parms[1]=BTN_TOUS then TOF_MBOPARAMDIM(TOTOF).ClickFlecheTous
    else if Parms[1]=BTN_AUCUN then TOF_MBOPARAMDIM(TOTOF).ClickFlecheAucun
    else if (Parms[1]=GRD_LISTE) or (Parms[1]=GRD_AFFICHE) then TOF_MBOPARAMDIM(TOTOF).RefreshBouton ;
    END ;
end ;

Initialization
registerclasses([TOF_MBOPARAMDIM]) ;
RegisterAglProc('OnChangeParamChampDim', True , 1, AGLOnChangeParamChampDim) ;
RegisterAglProc('OnMemoriser', True , 1, AGLOnMemoriser) ;
RegisterAglProc('OnClickBouton', True , 1, AGLOnClickBouton) ;
end.

