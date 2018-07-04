unit UtofDispoDim_Mode;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,UTOB,
      HCtrls,HEnt1,HMsgBox,UTOF,HDimension,ComCtrls,
      Dialogs,M3FP,EntGC,
{$IFDEF EAGLCLIENT}
      emul, MaineAGL,
{$ELSE}
      Fiche, mul, dbTables,Fe_Main,
{$ENDIF}
      LookUp,AglInitGC,UtilGC,UtilArticle,utilDimArticle ;

Type
     TOF_DISPODIM_MODE = Class (TOF)
       Private
       Public
        procedure OnArgument(Arguments : String) ; override ;
        Procedure OnUpdate ; override ;
        procedure OnLoad ; override ;
     end ;

procedure ChangeField(F:TForm ; FieldName:String) ;
procedure TOFDispoDim_Mode_ChangeField(Parms : Array of variant; nb: integer) ;
procedure TOFDispoDim_Mode_EnterField(Parms : Array of variant; nb: integer) ;

var
  bModifCritAuto : boolean ;
  bOldStatutUni, bOldStatutDim : boolean ; // Anciennes valeurs des checkBox
  bOldRBGroup : boolean ; // Ancienne valeur des radioButton
  Focus_GA_CODEBARRE : boolean ; // Focus sur GA_CODEBARRE lors du OnLoad ?

implementation


{ TOF_DISPODIM_MODE }

procedure TOF_DISPODIM_MODE.OnArgument(Arguments: String);
var THLIB : THLabel;
    iCol,Nbr : integer ;
    FF : TFMUL;
    BVisible : Boolean;
begin
Inherited ;
bModifCritAuto:=False ;
Focus_GA_CODEBARRE:=False ;
FF:=TFMul(Ecran) ;

// Paramétrage des libellés des familles, collection, stat. article et dimensions
ChangeLibre2('TGA_COLLECTION',Ecran);
for iCol:=1 to 3 do ChangeLibre2('TGA_FAMILLENIV'+InttoStr(iCol),Ecran);
if (ctxMode in V_PGI.PGIContexte) and (GetPresentation=ART_ORLI) then
    begin
    for iCol:=4 to 8 do ChangeLibre2('TGA2_FAMILLENIV'+InttoStr(iCol),Ecran);
    for iCol:=1 to 2 do ChangeLibre2('TGA2_STATART'+InttoStr(iCol),Ecran);
    end ;

// Mise en forme des libellés des dimensions    JD : 20/09/2002
for iCol:=1 to MaxDimension do
    BEGIN
    THLIB:=THLabel(GetControl('DIMENSION'+IntToStr(iCol))) ;
    THLIB.Caption:=RechDom('GCCATEGORIEDIM','DI'+InttoStr(iCol),False);
    BVisible := Not ((Copy(THLIB.Caption,1,2)='.-') or (THLIB.Caption='??'));
    THLIB.Visible := BVisible; TControl(GetControl(THLIB.FocusControl.Name)).Visible := BVisible;
    END;

// Paramètrage des libellés des onglets des zones et tables libres du dépôt
if not VH_GC.GCMultiDepots then
   begin
   SetControlCaption('PTABLESLIBRESDEP','Tables Libres Etab.') ;
   SetControlCaption('PZONESLIBRESDEP','Zones Libres Etab.') ;
   end;
                                
// Paramétrage des libellés des tables libres articles et dépôts
Nbr := 0;
if (GCMAJChampLibre (FF, False, 'EDIT', 'GA_LIBREART', 10, '') = 0) then SetControlVisible('PTABLESLIBRES', False) ;
if (GCMAJChampLibre (FF, False, 'EDIT', 'GDE_LIBREDEP', 10, '') = 0) then SetControlVisible('PTABLESLIBRESDEP', False) ;
// Mise en forme des libellés des dates, booléans libres et montants libres
if (GCMAJChampLibre (FF, False, 'EDIT', 'GA_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL', False) else inc(Nbr) ;
if (GCMAJChampLibre (FF, False, 'EDIT', 'GA_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE', False) else inc(Nbr) ;
if (GCMAJChampLibre (FF, False, 'BOOL', 'GA_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL', False) else inc(Nbr) ;
if (Nbr = 0) then SetControlVisible('PZONESLIBRES', False) ;
Nbr := 0;
if (GCMAJChampLibre (FF, False, 'EDIT', 'GDE_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VALDEP', False) else inc(Nbr) ;
if (GCMAJChampLibre (FF, False, 'EDIT', 'GDE_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATEDEP', False) else inc(Nbr) ;
if (GCMAJChampLibre (FF, False, 'BOOL', 'GDE_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOLDEP', False) else inc(Nbr) ;
if (Nbr = 0) then SetControlVisible('PZONESLIBRESDEP', False) ;

// Spécifique Front Office
if CtxFO in V_PGI.PGIContexte then
   BEGIN
   SetControlVisible('BInsert', False) ;
   SetControlVisible('BINSERTNOMENC', False) ;
   SetControlVisible('B_DUPLICATION', False) ;
   END ;
end;

procedure TOF_DISPODIM_MODE.OnLoad;
var CB : TCheckBox ;
    StockState : TCheckBoxState ;
    bAfficheDetail : boolean ;
    iControl : integer ;
    ctrl,xx_where : string ;
    F : TFMUL ;
begin
if not (Ecran is TFMul) then exit ;
F:=TFMul(Ecran) ;

// Force exécution si le focus est resté sur GA_CODEBARRE
if Focus_GA_CODEBARRE then ChangeField(ECRAN,'GA_CODEBARRE') ;

// Sélection de la liste des articles génériques ou dimensionnés
F.Q.Manuel:=True ; // Evite l'exécution de la requête lors de la maj de Q.Liste
if TRadioButton(F.FindComponent('RBGROUP')).Checked then
    BEGIN
    if GetPresentation=ART_ORLI then F.Q.Liste:='GCMULDISPART_MOS5'
                                else F.Q.Liste:='GCMULDISPART_MODE' ;
    END else
    BEGIN
    if GetPresentation=ART_ORLI then F.Q.Liste:='GCMULDISPDIM_MOS5'
                                else F.Q.Liste:='GCMULDISPDIM_MODE' ;
    END ;

if (ctxMode in V_PGI.PGIContexte) then
   SetControlProperty('GQ_DEPOT','Plus','GDE_SURSITE="X"');

xx_where:='' ;
// Gestion des checkBox du statut article
for iControl:=1 to 2 do
    BEGIN
    if iControl=1 then ctrl:='UNI'
    else if iControl=2 then ctrl:='DIM' ;
    CB:=TCheckBox(F.FindComponent('STATUT_'+ctrl)) ;
    if (CB<>nil) and (CB.State=cbChecked) then
        BEGIN
        if xx_where<>'' then xx_where:=xx_where+' or ' else xx_where:='(' ;
        xx_where:=xx_where+'GA_STATUTART="'+ctrl+'"' ;
        END ;
    END ;
if xx_where<>'' then xx_where:=xx_where+')' ;

// Gestion des checkBox : booléens libres article
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'BOOL', 'GA_BOOLLIBRE', 3, '');
// Gestion des dates libres article
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'DATE', 'GA_DATELIBRE', 3, '_');
// Gestion des montants libres article
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'EDIT', 'GA_VALLIBRE', 3, '_');
// Gestion des checkBox : booléens libres  dépôts
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'BOOL', 'GDE_BOOLLIBRE', 3, '');
// Gestion des dates libres dépôts
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'DATE', 'GDE_DATELIBRE', 3, '_');
// Gestion des montants libres dépôts
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'EDIT', 'GDE_VALLIBRE', 3, '_');
SetControlText('XX_WHERE',xx_where) ;

// Critères "Clôture" et "Article en stock"
xx_where:='GQ_CLOTURE<>"X"' ;
StockState:=TCheckBox(F.FindComponent('_STOCK')).State ;
if StockState=cbChecked then xx_where:=xx_where+' and GQ_PHYSIQUE>0'
else if StockState=cbUnchecked then xx_where:=xx_where+' and GQ_PHYSIQUE<=0' ;
SetControlText('XX_WHERE_STOCK',xx_where) ;

CB:=TCheckBox(F.FindComponent('STATUT_DIM')) ;
bAfficheDetail:=boolean((CB<>nil) and (CB.State=cbChecked)) ;
// Annule la sélection de critères dimension si les articles dimensionnés ne sont pas sélectionnés.
if not bAfficheDetail then
    BEGIN
    reinit_combo(THValComboBox(F.FindComponent('GDI_LIBELLE'))) ;
    reinit_combo(THValComboBox(F.FindComponent('GDI_LIBELLE_1'))) ;
    reinit_combo(THValComboBox(F.FindComponent('GDI_LIBELLE_2'))) ;
    reinit_combo(THValComboBox(F.FindComponent('GDI_LIBELLE_3'))) ;
    reinit_combo(THValComboBox(F.FindComponent('GDI_LIBELLE_4'))) ;
    END ;
F.Q.Manuel:=False ;
end;

procedure TOF_DISPODIM_MODE.OnUpdate;
var iCol : integer ;
    F : TFMUL ;
    TLIB : THLabel ;
    stIndice,stLabel : string ;
Begin
inherited ;
F:=TFMul(Ecran) ;

{$IFDEF EAGLCLIENT}
// Mise en place des libellés dans les colonnes
for iCol:=0 to F.FListe.ColCount-1 do
    BEGIN
    if copy(F.FListe.Cells[iCol,0],1,7)='Famille' then
        BEGIN  // Mise en place des libellés des familles
        stIndice:=copy(F.FListe.Cells[iCol,0],length(F.FListe.Cells[iCol,0]),1) ;
        if (stIndice>'3') and (stIndice<'9') then stLabel:='TGA2_FAMILLENIV' else stLabel:='TGA_FAMILLENIV' ;
        TLIB:=THLabel(F.FindComponent(stLabel+stIndice)) ;
        if TLIB<>nil then
            BEGIN
            if TLIB.Caption='.-' then F.Fliste.colwidths[iCol]:=0
                                 else F.Fliste.cells[iCol,0]:=TLIB.Caption ;
            END ;
        END
    else if (copy(F.FListe.Cells[iCol,0],1,3)='DIM') then
        BEGIN  // Mise en place des libellés des dimensions
        stIndice:=copy(F.FListe.Cells[iCol,0],length(F.FListe.Cells[iCol,0]),1) ;
        TLIB:=THLabel(F.FindComponent('DIMENSION'+stIndice)) ;
        if TLIB<>nil then
            BEGIN
            if TLIB.Caption='.-' then F.Fliste.colwidths[iCol]:=0
                                 else F.Fliste.cells[iCol,0]:=TLIB.Caption ;
            END ;
        END
    else if (copy(F.FListe.Cells[iCol,0],1,19)='Statistique article') then
        BEGIN  // Mise en place des libellés des statistiques articles
        stIndice:=copy(F.FListe.Cells[iCol,0],length(F.FListe.Cells[iCol,0]),1) ;
        TLIB:=THLabel(F.FindComponent('TGA2_STATART'+stIndice)) ;
        if TLIB<>nil then
            BEGIN
            if TLIB.Caption='.-' then F.Fliste.colwidths[iCol]:=0
                                 else F.Fliste.cells[iCol,0]:=TLIB.Caption ;
            END ;
        END ;
    END ;
{$ELSE}
// Mise en place des libellés dans les colonnes
for iCol:=0 to F.FListe.Columns.Count-1 do
    BEGIN
    if copy(F.FListe.Columns[iCol].Title.caption,1,7)='Famille' then
        BEGIN  // Mise en place des libellés des familles
        stIndice:=copy(F.FListe.Columns[iCol].Title.caption,length(F.FListe.Columns[iCol].Title.caption),1) ;
        if (stIndice>'3') and (stIndice<'9') then stLabel:='TGA2_FAMILLENIV' else stLabel:='TGA_FAMILLENIV' ;
        TLIB:=THLabel(F.FindComponent(stLabel+stIndice)) ;
        if TLIB<>nil then
            BEGIN
            if TLIB.Caption='.-' then F.Fliste.columns[iCol].visible:=False
                                 else F.Fliste.columns[iCol].Field.DisplayLabel:=TLIB.Caption ;
            END ;
        END
    else if (copy(F.FListe.Columns[iCol].Title.caption,1,3)='DIM') then
        BEGIN  // Mise en place des libellés des dimensions
        stIndice:=copy(F.FListe.Columns[iCol].Title.caption,length(F.FListe.Columns[iCol].Title.caption),1) ;
        TLIB:=THLabel(F.FindComponent('DIMENSION'+stIndice)) ;
        if TLIB<>nil then
            BEGIN
            if TLIB.Caption='.-' then F.Fliste.columns[iCol].visible:=False
                                 else F.Fliste.columns[iCol].Field.DisplayLabel:=TLIB.Caption ;
            END ;
        END
    else if (copy(F.FListe.Columns[iCol].Title.caption,1,19)='Statistique article') then
        BEGIN  // Mise en place des libellés des statistiques articles
        stIndice:=copy(F.FListe.Columns[iCol].Title.caption,length(F.FListe.Columns[iCol].Title.caption),1) ;
        TLIB:=THLabel(F.FindComponent('TGA2_STATART'+stIndice)) ;
        if TLIB<>nil then
            BEGIN
            if TLIB.Caption='.-' then F.Fliste.columns[iCol].visible:=False
                                 else F.Fliste.columns[iCol].Field.DisplayLabel:=TLIB.Caption ;
            END ;
        END ;
    END ;
{$ENDIF}

end;

procedure TOFDispoDim_Mode_ChangeField(Parms : Array of variant; nb: integer) ;
begin
ChangeField(TForm(Longint(Parms[0])),Parms[1]) ;
end ;

procedure ChangeField(F:TForm ; FieldName:String) ;
var stName : string ;
    iName : integer ;
    bTest : boolean ;
    CodeDim : THValComboBox ;
begin
CodeDim := Nil ;
// Modif automatique des critères : statut article et critères dimension/CAB
if (copy(FieldName,1,11)='GDI_LIBELLE') or (FieldName='GA_CODEBARRE') then
    BEGIN
    if (FieldName<>'GA_CODEBARRE') then CodeDim:=THValComboBox(F.FindComponent(FieldName))
                                   else Focus_GA_CODEBARRE:=False ; // Traitt effectué
    if (FieldName='GA_CODEBARRE') or ((CodeDim.Text<>'') and (CodeDim.Text<> TraduireMemoire ('<<Tous>>'))) then
        BEGIN
        GetSetStatut(F,'STATUT_UNI',True,bModifCritAuto,bOldStatutUni) ;
        GetSetStatut(F,'STATUT_DIM',True,bModifCritAuto,bOldStatutDim) ;
        if TRadioButton(F.FindComponent('RBGROUP')).Checked then
            BEGIN
            bOldRBGroup:=TRadioButton(F.FindComponent('RBGROUP')).Checked ;
            TRadioButton(F.FindComponent('RBDETAIL')).Checked:=True ;
            END ;
        bModifCritAuto:=True ;
        END
    else if (bModifCritAuto) then  // Modif critères automatique
        BEGIN
        bTest:=(THEdit(F.FindComponent('GA_CODEBARRE')).Text='') ;
        if bTest then for iName:=1 to 5 do
            BEGIN
            if iName=1 then stName:='GDI_LIBELLE'
            else if iName=2 then stName:='GDI_LIBELLE_1'
            else if iName=3 then stName:='GDI_LIBELLE_2'
            else if iName=4 then stName:='GDI_LIBELLE_3'
            else stName:='GDI_LIBELLE_4' ;
            CodeDim:=THValComboBox(F.FindComponent(stName)) ;
            bTest:=bTest and ((CodeDim.Text='') or (CodeDim.Text= TraduireMemoire ('<<Tous>>')) ) ;
            END ;
        if bTest then  // Toutes les combos "vides" -> réinit statut
            BEGIN
            bModifCritAuto:=False ;
            GetSetStatut(F,'STATUT_UNI',False,bModifCritAuto,bOldStatutUni) ;
            GetSetStatut(F,'STATUT_DIM',False,bModifCritAuto,bOldStatutDim) ;
            TRadioButton(F.FindComponent('RBGROUP')).Checked:=bOldRBGroup ;
            TRadioButton(F.FindComponent('RBDETAIL')).Checked:=not bOldRBGroup ;
            END else
            BEGIN
            TRadioButton(F.FindComponent('RBGROUP')).Checked:=not bOldRBGroup ;
            TRadioButton(F.FindComponent('RBDETAIL')).Checked:=bOldRBGroup ;
            END ;
        END ;
    END ;

if (copy(FieldName,1,7)='STATUT_') or (copy(FieldName,1,7)='RBGROUP') then
    BEGIN
    bModifCritAuto:=False ;
    if ((FieldName='STATUT_UNI') and (TCheckBox(F.FindComponent(FieldName)).checked)) or
       ((FieldName='RBGROUP') and (TRadioButton(F.FindComponent(FieldName)).checked)) then
        BEGIN
        // réinit combos dimension
        reinit_combo(THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE'))) ;
        reinit_combo(THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_1'))) ;
        reinit_combo(THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_2'))) ;
        reinit_combo(THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_3'))) ;
        reinit_combo(THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_4'))) ;
        // réinit CAB
        THEdit(F.FindComponent('GA_CODEBARRE')).Text:='' ;
        END ;
    if (FieldName='STATUT_UNI') and (TCheckBox(F.FindComponent(FieldName)).checked) and
       (not TRadioButton(F.FindComponent('RBGROUP')).Checked)
        then TRadioButton(F.FindComponent('RBGROUP')).Checked:=True ;
    END ;

if ((copy(FieldName,1,8)='RBDETAIL') and (TRadioButton(F.FindComponent(FieldName)).checked)) then
    BEGIN
    bOldStatutUni:=False ; GetSetStatut(F,'STATUT_UNI',False,False,bOldStatutUni) ;
    bOldStatutDim:=True ; GetSetStatut(F,'STATUT_DIM',False,False,bOldStatutDim) ;
    END ;
end ;

procedure TOFDispoDim_Mode_EnterField(Parms : Array of variant; nb: integer) ;
var FieldName : string ;
begin
FieldName:=Parms[1] ;
// Permet de savoir s'il faut ou non exécuté le ChangeField de GA_CODEBARRE au OnLoad.
if (FieldName='GA_CODEBARRE') then Focus_GA_CODEBARRE:=True ;
end ;

(*------------------------------------RECHERCHE ARTICLES----------------------------------
-----------------------------------------------------------------------------------------*)

Procedure TOFDispoDim_Mode_RechArt (Parms : Array of variant; nb: integer);
var F : TForm;
    G_Article : THCritMaskEdit;
    QArt : TQuery ;
    ChampArticle,NomDuChamp : String;
BEGIN
F := TForm (Longint (Parms[0]));
NomDuChamp := Parms[2];
ChampArticle := Parms[1];
G_Article := THCritMaskEdit (F.FindComponent (ChampArticle));
DispatchRecherche (G_Article, 1, '',
                   NomDuChamp + '=' + Trim (Copy (G_Article.Text, 1, 18))+';GA_TYPEARTICLE=MAR', '');
if G_Article.Text <> '' then
    Begin
    THEdit(F.FindComponent ('ARTICLE')).Text:= G_Article.Text;
    QArt:=OpenSQL('SELECT GA_CODEARTICLE FROM ARTICLE WHERE GA_ARTICLE="'
                                    +G_Article.Text+ '"',True) ;
    If not QArt.EOF then
      THEdit(F.FindComponent(ChampArticle)).Text:=QArt.FindField('GA_CODEARTICLE').AsString ;
    Ferme(QArt) ;
    end ;
END;

///Position des THEdit Libellés des Grilles de Dimension
Procedure TOFDispoDim_Mode_InitForm (Parms : Array of variant; nb: integer) ;
BEGIN
END;


(*-------------------------------AFFICHAGE DE LA FICHE ARTICLE----------------------------------
-----------------------------------------------------------------------------------------*)

Procedure TOFDispoDim_Mode_AfficArt (Parms : Array of variant; nb: integer);
var
    st1 : string;
    st2 : string;
begin
st1 := Parms [1];
st2 := CodeArticleUnique2(st1, '');
AGLLanceFiche('GC','GCARTICLE', '', st2, 'ACTION=CONSULTATION');
end;


// Parms[1] = GetChamp('GQ_ARTICLE') -> réf. article dimensionnée
// Parms[2] = STOCK=X;ARTDIM=...;DEPOT=...;CLOTURE=...
Procedure TOFAfficheDispoDim (Parms : Array of variant; nb: integer);
var st1, st2 : string;
begin
st1 := Parms [1];
st2 := CodeArticleUnique2(st1, '');
AGLLanceFiche('MBO','ARTICLE', '', st2, 'ACTION=CONSULTATION;'+Parms[2]);
end;

Procedure InitTOFDispoDim_Mode ;
begin
RegisterAglProc('RechArtDispoDim_Mode', True , 1, TOFDispoDim_Mode_RechArt);
RegisterAglProc('InitDispoDim_Mode', True , 0, TOFDispoDim_Mode_InitForm);
RegisterAglProc('AfficArtDispoDim_Mode', True , 1, TOFDispoDim_Mode_AfficArt);
RegisterAglProc('AfficheDispoDim', True , 1, TOFAfficheDispoDim);
RegisterAglProc('DispoDimModeChangeField', True , 1, TOFDispoDim_Mode_ChangeField) ;
RegisterAglProc('DispoDimModeEnterField', True , 1, TOFDispoDim_Mode_EnterField) ;
end ;

Initialization
RegisterClasses([TOF_DispoDim_Mode]) ;
InitTOFDispoDim_Mode ;
end.
