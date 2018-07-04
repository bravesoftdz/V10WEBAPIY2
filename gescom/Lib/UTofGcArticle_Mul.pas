unit UTofGcArticle_Mul;

interface                                                          
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF,Hdb,M3FP,UtilGC,HTB97,UTOB,CalcOLEGescom,
{$IFDEF EAGLCLIENT}
      emul,MaineAgl,UtileAGL,  
{$ELSE}
{$IFNDEF V530}
      EdtREtat,
{$ENDIF}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}db,DBGrids, mul,Fe_Main,EdtEtat,
{$IFDEF CCS3}
{$ELSE}
      AssistGenereCAB,
{$ENDIF}

{$ENDIF}
      HDimension,Entgc,utilPGI,utilArticle,utilDimArticle,AglInit,HStatus,
      M3VM,Hqry,FactComm,FactUtil,Ent1, AGLInitGC ;
Type
     TOF_GCARTICLE_MUL = Class (TOF)
        private
        AppelDepuis : string ; // Identifie d'où est appelé la fiche et génère les traitements adéquats
        StMulWhere : string; //stock la clause where du mul articles
        public
        procedure OnArgument(st:String) ; override ;
        procedure OnLoad ; override ;
        procedure OnUpdate ; override ;
        procedure OnClose ; override ; //AC Edition etiquette
        procedure AffMessage(NumMes,NumTit:integer) ;
        function  SaisieAffectationTrf : boolean ;
        function  GenereCAB : boolean ;
        function LanceEditEtiq : boolean ;
        //procedure  MajChampsLibresArticle ;
     END ;

procedure AGLRecupEtat2(parms:array of variant; nb: integer );

var
  bModifCritAuto : boolean ;
  bOldStatutGen, bOldStatutUni, bOldStatutDim : boolean ; // Anciennes valeurs des checkBox
  Tarif: boolean ;// Utilisé dans la saisie des TarifMode pour le choix de l'article

const
  // libellés des messages
  TexteMessage: array[1..2] of string = (
    {1}   'Les articles dimensionnés n''étant pas sélectionnés, les critères dimensions sont ignorés'
    {2}  ,'Attention'
         );

implementation
uses UFonctionsCBP;

procedure TOF_GCARTICLE_MUL.AffMessage (NumMes,NumTit : integer);
begin
PGIBox(TexteMessage[NumMes],TexteMessage[NumTit]) ;
end ;

{Les critères de type "GA_FOURNPRINC=001" passé en parametre après le mot clef EXCLUSIF
 deviennent disable}
procedure TOF_GCARTICLE_MUL.OnArgument(st:String) ;
var i_ind, iCol, PosEgal : integer ;
    THLIB : THLabel;
    BOUT : TToolbarButton97 ;
    FF : TFMUL ;
    Nbr : integer ;
    Critere, ChampMul, ValMul, stArg : string;
    CC : TComponent;
    bExclusif,DroitCreat : boolean ;
    ONGL : TTabSheet ;
    TypeArticle : THValComboBox ;
begin
Inherited ;

bModifCritAuto:=False ;
bExclusif:=False ;
AppelDepuis:='';

// Paramétrage des libellés des familles, stat. article et dimensions
for iCol:=1 to 3 do ChangeLibre2('TGA_FAMILLENIV'+InttoStr(iCol),Ecran);
ChangeLibre2('TGA_COLLECTION',Ecran);
if GetPresentation=ART_ORLI then
    begin
    for iCol:=4 to 8 do ChangeLibre2('TGA2_FAMILLENIV'+InttoStr(iCol),Ecran);
    for iCol:=1 to 2 do ChangeLibre2('TGA2_STATART'+InttoStr(iCol),Ecran);
    end ;

// Paramétrage des libellés des tables libres
Nbr := 0;
//if (GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GA_LIBREART', 10, '') = 0) then SetControlVisible('PTABLESLIBRES', False) ;

MajChampsLibresArticle(TForm(Ecran));

stArg:=st; // Pour traitement de consultation

if ctxMode in V_PGI.PGIContexte then
  begin
  FF:=TFMul(Ecran) ;
  Repeat
   Critere:=uppercase(Trim(ReadTokenSt(St))) ;
   TypeArticle:=THValComboBox(GetControl('GA_TYPEARTICLE')) ;
   if (Critere='GCARTICLE') or (Critere='ARTICLE') or (Critere='TARIF') then
      begin
      TypeArticle.Plus:='AND ((CO_CODE <>"PRE") AND (CO_CODE<>"FI"))';
      TypeArticle.Value:='MAR' ; // Initialisation par défaut
      if critere='TARIF' then Tarif:=True ;
      End else
   if (Critere='PROPTRANS') or (Critere='STOCKMIN') or (Critere='GENERECAB') or (Critere='ETIQARTCAT') then
     begin
     AppelDepuis:=Critere ;
     TypeArticle.Value:='MAR' ; // Initialisation par défaut
     BOUT:=TToolBarButton97 (FF.FindComponent ('bselectAll'));
     BOUT.Visible := True ;
     BOUT:=TToolBarButton97 (FF.FindComponent ('bOuvrir1'));
     BOUT.Visible := True ;
     BOUT:=TToolBarButton97 (FF.FindComponent ('bOuvrir'));
     BOUT.Visible := False ;
     BOUT:=TToolBarButton97 (FF.FindComponent ('binsert'));
     BOUT.Visible := False ;
     BOUT:=TToolBarButton97 (FF.FindComponent ('BINSERTNOMENC'));
     BOUT.Visible := False ;
     BOUT:=TToolBarButton97 (FF.FindComponent ('B_DUPLICATION'));
     BOUT.Visible := False ;
     BOUT:=TToolBarButton97 (FF.FindComponent ('BPROPRIETE'));
     BOUT.Visible := False ;
     BOUT:=TToolBarButton97 (FF.FindComponent ('BExport'));
     BOUT.Visible := False ;
{$IFDEF EAGLCLIENT}
     FF.FListe.MultiSelect := True;
{$ELSE}
     FF.FListe.MultiSelection := True;
{$ENDIF}
     end;
   if (Critere='ETIQARTCAT') then
   begin
        Ecran.Caption:=TraduireMemoire('Etiquettes sur catalogue');
        updatecaption(Ecran);
        ONGL:=TTabSheet(FF.FindComponent ('PMISEENPAGE'));
        ONGL.TabVisible:=True ;
   end;
   if Critere='EXCLUSIF' then bExclusif:=True else
   if Critere<>'' then
      begin
      i_ind:=pos('=',Critere);
      if i_ind>0 then
        begin
        ChampMul:=copy(Critere,1,i_ind-1);
        ValMul:=copy(Critere,i_ind+1,length(Critere));
        if ChampMul='TYPEARTICLE' then
          begin
          TypeArticle:=THValComboBox(GetControl('GA_TYPEARTICLE')) ;
          TypeArticle.Plus:='AND ('+ValMul+')' ;
          end
        else if ChampMul<>'ACTION' then
          begin
          if (ctxMode in V_PGI.PGIContexte) then
            begin
            if ChampMul='GA_CODEARTICLE' then ChampMul:='GA_ARTICLE';
            if (Pos('_ARTICLE',ChampMul)>0) and (ValMul<>'') then FF.FiltreDisabled := True;
            end;
          CC:=FF.FindComponent(ChampMul);
          if CC=nil then PGIError('Impossible de trouver le champ : "'+ChampMul+'"',Ecran.Caption)
          else
            begin
            If (CC is THEdit) then THEdit(CC).Text:=ValMul;
            If (CC is THValComboBox) then THValComboBox(CC).Value:=ValMul;
            if (CC is TCheckBox) then
              begin
              if (ValMul='X') then TCheckBox(CC).state:=cbChecked;
              if (ValMul='-') then TCheckBox(CC).state:=cbUnChecked;
              if (ValMul=' ') then TCheckBox(CC).state:=cbGrayed;
              end;
            if (CC is THLabel) then THLabel(CC).caption:=ValMul;
            SetControlEnabled(CC.Name,not bExclusif) ;
            if bExclusif=True then SetControlProperty (CC.Name,'Tag',-9988) // Non écrasé par le filtre
            else SetControlProperty (CC.Name,'Tag',0) ;
            end;
          end;
        end ;
      end ;
  until  Critere='';

  // Mise en forme des libellés des dimensions
  for iCol:=1 to MaxDimension do
      begin
      THLIB:=THLabel(GetControl('DIMENSION'+InttoStr(iCol)));
      THLIB.Caption:=RechDom('GCCATEGORIEDIM','DI'+InttoStr(iCol),False);
      if THLIB.Caption='.-'
      then BEGIN THLIB.Visible:=False ; THLIB.FocusControl.Visible:=False ; END
      // Annule l'alignement automatique des champs effectués par l'AGL
      else if (THLIB.FocusControl<>nil) and (THLIB.Left=(THLIB.FocusControl).Left) then THLIB.Top:=(THLIB.FocusControl).Top-17 ;
      end ;

  // Mise en forme des libellés des dates, booléans libres et montants libres
  if (GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL', False) else inc(Nbr) ;
  if (GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE', False) else inc(Nbr) ;
  if (GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'GA_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL', False) else inc(Nbr) ;
  if (Nbr = 0) then SetControlVisible('PZONESLIBRES', False) ;

  // Personalisation du Caption de la fiche
  if AppelDepuis='STOCKMIN' then
     begin
     FF.Caption:=TraduireMemoire('Sélection d''articles pour génération du stock mini et maxi') ;
     UpdateCaption(Ecran);
     end;
  end ; // if (ctxMode in V_PGI.PGIContexte) then

SetControlText('TYPEACTION','ACTION=MODIFICATION'); //Initialisation à modification
//Cache les boutons ajout et dupliquer si consultation
Repeat
    Critere := UpperCase(Trim(ReadTokenSt(stArg)));
    if Critere<>'' then
        BEGIN
        PosEgal := pos('=',Critere);
        if PosEgal<>0 then ChampMul := copy(Critere,1,PosEgal-1);
           if ChampMul='ACTION' then SetControlText('TYPEACTION',Critere);
        END;
until Critere='';
SetControlVisible('BINSERT',StringToAction(GetControlText('TYPEACTION'))<>taConsult);
SetControlVisible('BINSERTNOMENC',StringToAction(GetControlText('TYPEACTION'))<>taConsult);
SetControlVisible('B_DUPLICATION',StringToAction(GetControlText('TYPEACTION'))<>taConsult);

DroitCreat:=ExJaiLeDroitConcept(TConcept(gcArtCreat),False) ;
if (ctxFO in V_PGI.PGIContexte) or (Not DroitCreat) then
  begin
  // Suppression des boutons d'insertion et de duplication
  SetControlVisible('BINSERT', False) ;
  SetControlVisible('BINSERTNOMENC', False) ;
  SetControlVisible('B_DUPLICATION', False) ;
  end ;

end ;

{$IFDEF SUPPRIME} //remplacé par procedure homonyme de UtilArticle
Procedure TOF_GCARTICLE_MUL.MajChampsLibresArticle ;
var i: integer;
    AuMoinsUn : boolean;
    FF:TForm ;
begin
FF:=Tform(Ecran);
AuMoinsUn:=False;
For i:=1 to 10 do if i<10 then AuMoinsUn:=(ChangeLibre2('TGA_LIBREART'+intToStr(i),FF) or AuMoinsUn)
                          else AuMoinsUn:=(ChangeLibre2('TGA_LIBREARTA',FF)or AuMoinsUn);
SetControlVisible('PTABLESLIBRES', AuMoinsUn) ; {Invisible si aucun controle}

AuMoinsUn:=False;
For i:=1 to 3 do AuMoinsUn:=(ChangeBoolLibre('GA_BOOLLIBRE'+intToStr(i),FF)or AuMoinsUn);
For i:=1 to 3 do AuMoinsUn:=(ChangeLibre2('TGA_VALLIBRE'+intToStr(i),FF)or AuMoinsUn);
For i:=1 to 3 do AuMoinsUn:=(ChangeLibre2('TGA_DATELIBRE'+intToStr(i),FF)or AuMoinsUn);
{$IFDEF CCS3}
SetControlVisible('PZONESLIBRES',False) ;
{$ELSE}
SetControlVisible('PZONESLIBRES',AuMoinsUn) ;
{$ENDIF}
end;
{$ENDIF}


procedure TOF_GCARTICLE_MUL.OnLoad  ;
var F : TFMul ;
    CC : TCheckBox ;
    CB : TCheckBox ;
    iControl : integer ;
    bAffiche : boolean ;
    ctrl, xx_where : string ;
begin
inherited ;
if not (Ecran is TFMul) then exit ;
if not (ctxMode in V_PGI.PGIContexte) then exit ;

xx_where:='' ;
F:=TFMul(Ecran) ;

F.Q.Manuel:=True ; // Evite l'exécution de la requête lors de la maj de Q.Liste
// Si critère code barre renseigné, sélection forcée sur les articles dimensionnés
if ( GetControlText ( 'GA_CODEBARRE' ) <> '' ) and
   ( GetControlText ( 'STATUT_DIM' ) <> 'X' ) then SetControlText ('STATUT_DIM', 'X' ) ;
if GetControlText ( 'STATUT_DIM' ) = 'X'
    then F.Q.Liste:='GCMULARTDIM_MODE'
    else F.Q.Liste:='GCMULARTICLE_MODE' ;

/// Gestion des checkBox du statut article
for iControl:=1 to 3 do
  begin
  if iControl=1 then ctrl:='GEN'
  else if iControl=2 then ctrl:='UNI'
  else ctrl:='DIM' ;
  CC:=TCheckBox(TFMul(F).FindComponent('STATUT_'+ctrl)) ;
  if (CC<>nil) and (CC.State=cbChecked) then
    begin
    if xx_where<>'' then xx_where:=xx_where+' or ' else xx_where:='(' ;
    xx_where:=xx_where+'GA_STATUTART="'+ctrl+'"' ;
    end ;
  end ;
if xx_where<>'' then xx_where:=xx_where+')' ;

// Gestion des checkBox : tenu en stock, commissionnable, remise ligne et remise pied
for iControl:=1 to 4 do
  begin
  if iControl=1 then ctrl:='GA_TENUESTOCK'
  else if iControl=2 then ctrl:='GA_COMMISSIONNABLE'
  else if iControl=3 then ctrl:='GA_REMISELIGNE'
  else ctrl:='GA_REMISEPIED' ;
  CC:=TCheckBox(TFMul(F).FindComponent(ctrl)) ;
  if (CC<>nil) and (CC.State=cbChecked) then
    begin
    if xx_where<>'' then xx_where:=xx_where+' AND ' ;
    xx_where:=xx_where+ctrl+'="X"' ;
    end
  else if (CC<>nil) and (CC.State=cbUnChecked) then
    begin
    if xx_where<>'' then xx_where:=xx_where+' AND ' ;
    xx_where:=xx_where+ctrl+'="-"' ;
    end ;
  end ;

// Gestion des checkBox : booléens libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'BOOL', 'GA_BOOLLIBRE', 3, '');

// Gestion des dates libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'DATE', 'GA_DATELIBRE', 3, '_');

// Gestion des montants libres
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'EDIT', 'GA_VALLIBRE', 3, '_');

SetControlText('XX_WHERE',xx_where) ;

CB:=TCheckBox(TFMul(F).FindComponent('STATUT_DIM')) ;
bAffiche:=boolean((CB<>nil) and (CB.State=cbChecked)) ;

// Annule la sélection de critères dimension si les articles dimensionnés ne sont pas sélectionnés.
if not bAffiche then
  begin
  reinit_combo (THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE'))) ;
  reinit_combo (THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_1'))) ;
  reinit_combo (THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_2'))) ;
  reinit_combo (THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_3'))) ;
  reinit_combo (THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_4'))) ;
  end ;

StMulWhere := RecupWhereCritere(TPageControl(TFMul(Ecran).Pages));

F.Q.Manuel:=False ;
end ;

procedure TOF_GCARTICLE_MUL.OnUpdate;
var iCol : integer ;
    bAffiche : boolean ;
    F : TFMul ;
    TLIB : THLabel ;
    CB : TCheckBox ;
    stIndice,stLabel : string ;
begin
inherited ;
if not (Ecran is TFMul) then exit;
F:=TFMul(Ecran) ;

CB:=TCheckBox(TFMul(F).FindComponent('STATUT_DIM')) ;
bAffiche:=boolean((CB<>nil) and (CB.State=cbChecked)) ;

// Maj de la liste : affichage ou non des colonnes dimensions et modif des titres

{$IFDEF EAGLCLIENT}
// Mise en place des libellés dans les colonnes
for iCol:=0 to F.FListe.ColCount-1 do
    BEGIN
    if copy(F.FListe.Cells[iCol,0],1,7)=TraduireMemoire('Famille') then
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
    else if (bAffiche) and (copy(F.FListe.Cells[iCol,0],1,3)='DIM') then
        BEGIN  // Mise en place des libellés des dimensions
        stIndice:=copy(F.FListe.Cells[iCol,0],length(F.FListe.Cells[iCol,0]),1) ;
        TLIB:=THLabel(F.FindComponent('DIMENSION'+stIndice)) ;
        if TLIB<>nil then
            BEGIN
            if TLIB.Caption='.-' then F.Fliste.colwidths[iCol]:=0
                                 else F.Fliste.cells[iCol,0]:=TLIB.Caption ;
            END ;
        END
    else if (copy(F.FListe.Cells[iCol,0],1,19)=TraduireMemoire('Statistique article')) then
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
// END ;
{$ELSE}
// Mise en place des libellés dans les colonnes
for iCol:=0 to F.FListe.Columns.Count-1 do
    BEGIN
    if copy(F.FListe.Columns[iCol].Title.caption,1,7)=TraduireMemoire('Famille') then
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
    else if (bAffiche) and (copy(F.FListe.Columns[iCol].Title.caption,1,3)='DIM') then
        BEGIN  // Mise en place des libellés des dimensions
        stIndice:=copy(F.FListe.Columns[iCol].Title.caption,length(F.FListe.Columns[iCol].Title.caption),1) ;
        TLIB:=THLabel(F.FindComponent('DIMENSION'+stIndice)) ;
        if TLIB<>nil then
            BEGIN
            if TLIB.Caption='.-' then F.Fliste.columns[iCol].visible:=False
                                 else F.Fliste.columns[iCol].Field.DisplayLabel:=TLIB.Caption ;
            END ;
        END
    else if (copy(F.FListe.Columns[iCol].Title.caption,1,19)=TraduireMemoire('Statistique article')) then
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

procedure TOF_GCARTICLE_MUL.OnClose ;
begin
if AppelDepuis='ETIQARTCAT' then
  begin
  VH_GC.TOBEdt.ClearDetail ;
  initvariable;
  end ;
end ;

procedure TOFArticleMul_ChangeField(Parms : Array of variant; nb: integer) ;
var F : TForm ;
    FieldName, stName : string ;
    iName : integer ;
    bTest : boolean ;
    CodeDim : THValComboBox ;
begin
F := TForm (Longint (Parms[0])) ;
FieldName:=Parms[1] ;

// Modif automatique des critères : statut article et critères dimension
if (copy(FieldName,1,11)='GDI_LIBELLE') then
   begin
   CodeDim := THValComboBox(F.FindComponent(FieldName)) ;
   if (CodeDim.Text<>'') and (CodeDim.Text<>TraduireMemoire('<<Tous>>')) then
     begin
     GetSetStatut(F,'STATUT_GEN',True,bModifCritAuto,bOldStatutGen) ;
     GetSetStatut(F,'STATUT_UNI',True,bModifCritAuto,bOldStatutUni) ;
     GetSetStatut(F,'STATUT_DIM',True,bModifCritAuto,bOldStatutDim) ;
     bModifCritAuto:=True ;
     end
   else if (bModifCritAuto) then  // Modif critères automatique
     begin
     bTest:=True ;
     for iName:=1 to 5 do
       begin
       if iName=1 then stName:='GDI_LIBELLE'
       else if iName=2 then stName:='GDI_LIBELLE_1'
       else if iName=3 then stName:='GDI_LIBELLE_2'
       else if iName=4 then stName:='GDI_LIBELLE_3'
       else stName:='GDI_LIBELLE_4' ;
       CodeDim:=THValComboBox(F.FindComponent(stName)) ;
       bTest:=bTest and ((CodeDim.Text='') or (CodeDim.Text=TraduireMemoire('<<Tous>>'))) ;
       end ;
     if bTest then  // Toutes les combos "vides" -> réinit statut
       begin
       bModifCritAuto:=False ;
       GetSetStatut(F,'STATUT_GEN',False,bModifCritAuto,bOldStatutGen) ;
       GetSetStatut(F,'STATUT_UNI',False,bModifCritAuto,bOldStatutUni) ;
       GetSetStatut(F,'STATUT_DIM',False,bModifCritAuto,bOldStatutDim) ;
       end ;
     end ;
   end ;

if (copy(FieldName,1,7)='STATUT_') then
  begin
  bModifCritAuto:=False ;
  if ((FieldName='STATUT_GEN') or (FieldName='STATUT_UNI')) and
     (TCheckBox(F.FindComponent(FieldName)).checked) then
    begin
    // réinit combos dimension
    reinit_combo (THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE'))) ;
    reinit_combo (THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_1'))) ;
    reinit_combo (THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_2'))) ;
    reinit_combo (THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_3'))) ;
    reinit_combo (THValComboBox(TFMul(F).FindComponent('GDI_LIBELLE_4'))) ;
    end ;
  end ;

end ;
//
// Procedure de saisie d'une nouvelle affectation de transfert
//
Function TOF_GCARTICLE_MUL.SaisieAffectationTrf : boolean;
var F : TFMUL;
{$IFDEF EAGLCLIENT}
    L : THGrid;
{$ELSE}
    L : THDBGrid;
{$ENDIF}
    Q : TQuery;
    i : integer;
    StSQL : string;
    TraitOk : String ;
    TobMulArticle,TobArticle : TOB ;
begin
  F := TFMul(Ecran);
  L := F.FListe;
{$IFNDEF EAGLCLIENT}
  Q := F.Q;
{$ENDIF}
  Result:=False;
  TraitOk:='';
  // Suite à la validation du multi-critères Article, la liste des articles sélectionnés
  // est chargée en TOB et récupérée dans la fiche "PR_AFFECTATION".
  if (L.NbSelected=0) and (not L.AllSelected) then
    MessageAlerte('Aucun article n''a été sélectionné')
  else
    begin
    TobMulArticle:=TOB.Create('MulArticles', Nil, -1);
    if L.AllSelected then
      begin
      StSQL:='SELECT GA_ARTICLE,GA_CODEARTICLE,GA_LIBELLE,GA_STATUTART FROM ARTICLE_MODE '+StMulWhere;
      TobMulArticle.LoadDetailFromSQL(StSQL,False,True);
      end
    else
      begin
      InitMove(L.NbSelected,'');
      for i:=0 to L.NbSelected-1 do
        BEGIN
        L.GotoLeBOOKMARK(i);
        TobArticle:=TOB.Create('Article', TobMulArticle, -1);
        TobArticle.PutValue    ('GA_ARTICLE', Q.FindField('GA_ARTICLE').AsString);
        TobArticle.PutValue    ('GA_CODEARTICLE', Q.FindField('GA_CODEARTICLE').AsString);
        TobArticle.PutValue    ('GA_LIBELLE', Q.FindField('GA_LIBELLE').AsString);
        TobArticle.PutValue    ('GA_STATUTART', Q.FindField('GA_STATUTART').AsString);
        MoveCur(False);
        END;
      FiniMove;
      end;
    TheTOB := TobMulArticle;
    if AppelDepuis='PROPTRANS' then
      TraitOk:=AGLLanceFiche('MBO', 'PR_AFFECTATION', '','', GetControltext('CODETRAIT'))
    else TraitOk:=AGLLanceFiche('MBO', 'PR_AFFMINMAX', '','', '');
    if L.AllSelected then L.AllSelected:=False else L.ClearSelected;
    if TraitOk='OK' then Result:=True ;
    end;
end;

//
// Procedure de génération des codes barres
//
Function TOF_GCARTICLE_MUL.GenereCAB : boolean ;
var F : TFMUL;
{$IFDEF EAGLCLIENT}
    L : THGrid;
{$ELSE}
    L : THDBGrid;
//    Q : TQuery;
{$ENDIF}
    i : integer;
    Article,CodeArticle,Libelle,Statut,FournPrinc : string ;
    TraitOk : String ;
    TobMulArticle,TobArticle : TOB ;
begin
  F := TFMul(Ecran);
  L := F.FListe;
// INUTILE ET LOURD !!!???
//{$IFNDEF EAGLCLIENT}
//  Q := F.Q;
//{$ENDIF}
  Result:=True;
  TraitOk:='';
  // Suite à la validation du multi-critères Article, la liste des articles sélectionnés
  // est chargée en TOB et récupérée dans la fiche "PR_AFFECTATION".
  if (L.NbSelected=0) and (not L.AllSelected) then
     BEGIN
     MessageAlerte('Aucun article n''a été sélectionné');
     END else
     BEGIN
     TobMulArticle:=TOB.Create('MulArticles', Nil, -1);
     if L.AllSelected then
        BEGIN
//{$IFDEF EAGLCLIENT}
        InitMove(F.Q.RecordCount,'');
        F.Q.First;
        while Not F.Q.EOF do
//{$ELSE}
//        InitMove(Q.RecordCount,'');
//        Q.First;
//        while Not Q.EOF do
//{$ENDIF}
          BEGIN
          Article:=    F.Q.FindField('GA_ARTICLE').AsString ;
          CodeArticle:=F.Q.FindField('GA_CODEARTICLE').AsString ;
          Libelle:=    F.Q.FindField('GA_LIBELLE').AsString ;
          Statut:=     F.Q.FindField('GA_STATUTART').AsString ;
          FournPrinc:= F.Q.FindField('GA_FOURNPRINC').AsString ;
          TobArticle:=TOB.Create('Les Articles', TobMulArticle, -1);
          TobArticle.AddChampSup ('GA_ARTICLE', False);
          TobArticle.AddChampSup ('GA_CODEARTICLE', False);
          TobArticle.AddChampSup ('GA_LIBELLE', False);
          TobArticle.AddChampSup ('GA_STATUTART', False);
          TobArticle.AddChampSup ('GA_FOURNPRINC', False);
          TobArticle.PutValue    ('GA_ARTICLE', Article);
          TobArticle.PutValue    ('GA_CODEARTICLE', CodeArticle);
          TobArticle.PutValue    ('GA_LIBELLE', Libelle);
          TobArticle.PutValue    ('GA_STATUTART', Statut);
          TobArticle.PutValue    ('GA_FOURNPRINC', FournPrinc);
          MoveCur(False);
//{$IFDEF EAGLCLIENT}
          F.Q.NEXT;
//{$ELSE}
//          Q.NEXT;
//{$ENDIF}
          END;
        FiniMove;
        L.AllSelected:=False;
        END ELSE
        BEGIN
        InitMove(L.NbSelected,'');
        for i:=0 to L.NbSelected-1 do
          BEGIN
          L.GotoLeBOOKMARK(i);
          Article:=    F.Q.FindField('GA_ARTICLE').AsString ;
          CodeArticle:=F.Q.FindField('GA_CODEARTICLE').AsString ;
          Libelle:=    F.Q.FindField('GA_LIBELLE').AsString ;
          Statut:=     F.Q.FindField('GA_STATUTART').AsString ;
          FournPrinc:= F.Q.FindField('GA_FOURNPRINC').AsString ;
          TobArticle:=TOB.Create('Les Articles', TobMulArticle, -1);
          TobArticle.AddChampSup ('GA_ARTICLE', False);
          TobArticle.AddChampSup ('GA_CODEARTICLE', False);
          TobArticle.AddChampSup ('GA_LIBELLE', False);
          TobArticle.AddChampSup ('GA_STATUTART', False);
          TobArticle.AddChampSup ('GA_FOURNPRINC', False);
          TobArticle.PutValue    ('GA_ARTICLE', Article);
          TobArticle.PutValue    ('GA_CODEARTICLE', CodeArticle);
          TobArticle.PutValue    ('GA_LIBELLE', Libelle);
          TobArticle.PutValue    ('GA_STATUTART', Statut);
          TobArticle.PutValue    ('GA_FOURNPRINC', FournPrinc);
          MoveCur(False);
          END ;
        FiniMove;
        L.ClearSelected;
        END ;
{$IFDEF EAGLCLIENT}

{$ELSE}
{$IFDEF CCS3}
{$ELSE}
     Assist_GenereCAB(TobMulArticle) ;
{$ENDIF}
{$ENDIF}
     TobMulArticle.Free ;
     END ;
end;

// Procedure d'édition des étiquettes sur catalogue
Function TOF_GCARTICLE_Mul.LanceEditEtiq : Boolean;
var F : TFMUL;
{$IFDEF EAGLCLIENT}
    L : THGrid;
{$ELSE}
    L : THDBGrid;
//    Q : TQuery;
{$ENDIF}
    i,i_ind1,nbEtiq : integer;
    Article,CodeArticle,ArticleGenerique,Libelle,Statut,RegimePrix : string ;
    CodeEtat,LibEtat,stInsert,stWhere,compteur,CodeArticleGen : string ;
    BApercu : Boolean;     // ,statut_gen,statut_dim
begin
  F := TFMul (Ecran) ;
  L := F.FListe ;
  // INUTILE ET LOURD !!!???
  //{$IFNDEF EAGLCLIENT}
  //  Q := F.Q;
  //{$ENDIF}
  Result:=True;
  // Suppression des enregistrements
  ExecuteSQL('DELETE FROM GCTMPETQ WHERE GZD_UTILISATEUR = "'+V_PGI.USer+'"');
  // Récupération de bon RegimePrix
  if (ctxMode in V_PGI.PGIContexte) then RegimePrix := 'TTC' else
  if GetControlText('RHT') = 'X' then RegimePrix := 'HT' else RegimePrix := 'TTC';
  // Récupération du nombre d'exemplaire d'étiquette demandé
  nbEtiq := StrToInt(GetControlText('NBEXEMPLAIRE'));
  // Récupération du modele d'état choisi
  CodeEtat := GetControlText ('FETAT');
  LibEtat := RechDom('TTMODELETIQART',CodeEtat,FALSE);
  // Suite à la validation du multi-critères Article, la liste des articles sélectionnés
  // est chargée en TOB
  if (L.NbSelected=0) and (not L.AllSelected) then
  begin
    MessageAlerte('Aucun article n''a été sélectionné');
  end else
  begin
    if L.AllSelected then
    begin
      //statut_gen := TCheckBox(TForm(Ecran).FindComponent('STATUT_GEN')).Checked ;
      //statut_dim := TCheckBox(TForm(Ecran).FindComponent('STATUT_DIM')).Checked ;
      // ATTENTION article unique
      //if statut_gen and statut_dim then inc(nbEtiq);   POURQUOI ??  OT 14/02/2003
      //{$IFDEF EAGLCLIENT}
      InitMove(F.Q.RecordCount,'');
      F.Q.First;
      //if stWhere<>'' then stWhere:=stWhere+' AND (';
      while Not F.Q.EOF do
      //{$ELSE}
      //        InitMove(Q.RecordCount,'');
      //        Q.First;
      //        while Not Q.EOF do
      //{$ENDIF}
      begin
        Article:=    F.Q.FindField('GA_ARTICLE').AsString ;
        CodeArticle:=F.Q.FindField('GA_CODEARTICLE').AsString ;
        Libelle:=    F.Q.FindField('GA_LIBELLE').AsString ;
        Statut:=     F.Q.FindField('GA_STATUTART').AsString ;
        // Construction de la clause where du générateur d'état afin d'avoir les
        // étiquettes correspondant aux articles sélectionnés
        if (Statut='GEN') or (Statut='UNI') then
        begin
          if stWhere<>'' then
            stWhere:=stWhere+' OR GA_CODEARTICLE="'+CodeArticle+'"' else
            stWhere:=' AND (GA_CODEARTICLE="'+CodeArticle+'"';
        end else
        begin
        // Comme on a tout sélectionné il suffit de récupérer le code article une seul fois pour
        // tous les article génériques
             if ArticleGenerique<>CodeArticle then
             begin
                  if stWhere<>'' then stWhere:=stWhere+' OR GA_CODEARTICLE="'+CodeArticle+'"'
                  else stWhere:=' AND (GA_CODEARTICLE="'+CodeArticle+'"';
                  ArticleGenerique:=CodeArticle
             end;
        end;
        MoveCur(False);
        //{$IFDEF EAGLCLIENT}
        F.Q.NEXT;
        //{$ELSE}
        // Q.NEXT;
        //{$ENDIF}
      end ;
      if stWhere <> '' then stWhere:=stWhere+')';
      FiniMove ;
      L.AllSelected := False ;
    end else
    // Sélection multiple
    begin
      InitMove(L.NbSelected, '') ;
      for i := 0 to L.NbSelected - 1 do
      begin
        L.GotoLeBOOKMARK(i);
        {$IFDEF EAGLCLIENT}
        F.Q.TQ.Seek (F.Fliste.row-1) ;
        {$ENDIF}
        Article:=    F.Q.FindField('GA_ARTICLE').AsString ;
        CodeArticle:=F.Q.FindField('GA_CODEARTICLE').AsString ;
        Libelle:=    F.Q.FindField('GA_LIBELLE').AsString ;
        Statut:=     F.Q.FindField('GA_STATUTART').AsString ;

        // Construction de la clause where du générateur d'état afin d'avoir les
        // étiquettes correspondant aux articles sélectionnés
        if (Statut = 'GEN') then
        begin
          // Mémorisation du code de l'article générique
          ArticleGenerique := CodeArticle;
          if stWhere = '' then
            stWhere := stWhere + ' AND (GA_CODEARTICLE="' + CodeArticle + '"' else
            stWhere := stWhere + ' OR GA_CODEARTICLE="' + CodeArticle + '"' ;
        end else
        begin
          if stWhere = '' then
            stWhere := stWhere + ' AND (GA_ARTICLE="' + Article + '"' else
            stWhere := stWhere + ' OR GA_ARTICLE="' + Article + '"' ;
           // Ceci me permet d'insérer l'article dimentionné voulu en plus de l'article générique
           // En effet lorsque l'utilisateur sélectionne un article générique, tous les articles
           // dimentionnés correspondant au générique seront imprimés
           if CodeArticle=ArticleGenerique then
           begin
             if EtatMonarchFactorise(LibEtat) then
             begin
               CodeArticleGen:=GCDimToGen(CodeArticle);
               stInsert := 'Insert into GCTMPETQ (GZD_UTILISATEUR, GZD_ARTICLE, GZD_DEPOT, GZD_REGIMEPRIX, GZD_NUMERO, GZD_SOUCHE, GZD_INDICEG, GZD_NUMLIGNE,GZD_NBETIQDIM,GZD_CODEARTICLEGEN)'
                         + 'values ("' + V_PGI.User + '", "' + Article + '","' + VH^.EtablisDefaut + '","' + RegimePrix + '",1,0,0,0,' + IntToStr(nbEtiq) + ',"' + CodeArticleGen + '")' ;
               ExecuteSQL(stInsert);
             end else
             begin
               for i_ind1 := 0 to nbEtiq - 1 do
               begin
                 compteur := IntToStr(i_ind1+1) ;
                 CodeArticleGen := GCDimToGen(CodeArticle) ;
                 stInsert := 'Insert into GCTMPETQ (GZD_UTILISATEUR, GZD_COMPTEUR, GZD_ARTICLE,GZD_DEPOT, GZD_REGIMEPRIX, GZD_NUMERO, GZD_SOUCHE, GZD_INDICEG, GZD_NUMLIGNE,GZD_CODEARTICLEGEN)'
                           + 'values ("' + V_PGI.User + '", ' + compteur + ', "' + Article + '","' + VH^.EtablisDefaut + '","' + RegimePrix + '",1,0,0,0,"' + CodeArticleGen + '")' ;
                 ExecuteSQL(stInsert);
               end ;
             end ;
           end ;
        end ;
        MoveCur(False) ;
      end ;
      if stWhere<>'' then stWhere := stWhere + ')' ;
      FiniMove;
      L.ClearSelected;
    end ;

//    {$IFDEF EAGLCLIENT}
//    {$ELSE}
    if stwhere<>'' then
    // On insere les enregistrements que si on a quelquechose à éditer
    begin
      if not (CtxMode in V_PGI.PGIContexte) then
      // On insere tous les articles dans la table temporaire autant de fois que d'exemplaire demandé
      stInsert := 'Insert into GCTMPETQ (GZD_UTILISATEUR, GZD_ARTICLE, GZD_DEPOT, GZD_CODEARTICLE, GZD_REGIMEPRIX,GZD_NUMERO,GZD_SOUCHE, GZD_INDICEG, GZD_NUMLIGNE, GZD_NBETIQDIM, GZD_CODEARTICLEGEN) ' +
                  'Select "' + V_PGI.User + '" as GZD_UTILISATEUR, ' +
                  'A1.GA_ARTICLE as GZD_ARTICLE, ' +
                  '"'+VH^.EtablisDefaut+'", '+
                  'A1.GA_CODEARTICLE as GZD_CODEARTICLE, ' +
                  '"' + RegimePrix + '" as GZD_REGIMEPRIX, '+
                  '0 as GZD_NUMERO,'+
                  '0 as GZD_SOUCHE,'+
                  '0 as GZD_INDICEG,'+
                  '0 as GZD_NUMLIGNE,'+
                  '"'+IntToStr(nbEtiq)+'" as GZD_NBETIQDIM,'+
                  'A2.GA_ARTICLE as GZD_CODEARTICLEGEN from ARTICLE A1, ARTICLE A2 '+
                  'WHERE A2.GA_STATUTART in ("GEN","UNI") and A1.GA_CODEARTICLE=A2.GA_CODEARTICLE'  else
      stInsert := 'Insert into GCTMPETQ (GZD_UTILISATEUR, GZD_ARTICLE, GZD_DEPOT, GZD_CODEARTICLE, GZD_REGIMEPRIX,GZD_NUMERO,GZD_SOUCHE, GZD_INDICEG, GZD_NUMLIGNE, GZD_NBETIQDIM, GZD_CODEARTICLEGEN) ' +
                  'Select "' + V_PGI.User + '" as GZD_UTILISATEUR, ' +
                  'GA_ARTICLE as GZD_ARTICLE, ' +
                  '"'+VH^.EtablisDefaut+'", '+
                  'GA_CODEARTICLE as GZD_CODEARTICLE, ' +
                  '"' + RegimePrix + '" as GZD_REGIMEPRIX, '+
                  '0 as GZD_NUMERO,'+
                  '0 as GZD_SOUCHE,'+
                  '0 as GZD_INDICEG,'+
                  '0 as GZD_NUMLIGNE,'+
                  '"'+IntToStr(nbEtiq)+'" as GZD_NBETIQDIM,'+
                  'GA_ARTICLE as GZD_CODEARTICLEGEN from ARTICLE_MODE '+
                  'WHERE GA_STATUTART in ("DIM","UNI")' + StWhere ;        
      if not EtatMonarchFactorise(LibEtat) then for i_ind1:=0 to nbEtiq - 1 do ExecuteSQL(stInsert)
      else ExecuteSQL(stInsert);
    end;
    // On ne veut que les enregistrements crée par l'utilisateur
    stWhere:='GZD_UTILISATEUR="'+V_PGI.USer+'" AND '+stWhere;

    // Récupération de la coche "apercu avant impression"
    BApercu := TCheckBox(TForm(Ecran).FindComponent('FApercu')).Checked;

    VH_GC.TOBEdt.ClearDetail ;
    initvariable;

    // Lancement de l'état
    EditMonarchSiEtat (LibEtat);
    LanceEtat('E','GED',CodeEtat,BApercu,False,False,Nil,stWhere,'',False);
    EditMonarch ('');
    // Pour que le bouton AllSelected ne soit plus appuyé
    F.bSelectAll.Down := False ;
    // Suppresion de tous les enregistrements de l'utilisateur
    ExecuteSQL('DELETE FROM GCTMPETQ WHERE GZD_UTILISATEUR = "'+V_PGI.USer+'"');
//    {$ENDIF}
  end ;
end;

Function TOF_GCARTICLE_MUL_AGLTraitementOk (Parms: array of variant; nb: integer) : variant ;
var F : TFMUL ;
    TOTOF : TOF ;
begin
Result:=False ;
F:=TFMul(Longint(Parms[0])) ;
if (F is TFMul) then TOTOF:=TFMul(F).LaTof else exit ;
if (TOTOF is TOF_GCARTICLE_MUL) then
  BEGIN
  with TOF_GCARTICLE_MUL(TOTOF) do
    BEGIN
    if (AppelDepuis='PROPTRANS') or (AppelDepuis='STOCKMIN') then Result:=SaisieAffectationTrf
    else if AppelDepuis='GENERECAB' then Result:=GenereCAB
         else if AppelDepuis='ETIQARTCAT' then Result:=LanceEditEtiq ;
    END ;
  END ;
end ;

Function TOF_GCARTICLE_MUL_AGLExtractArgument (Parms: array of variant; nb: integer) : variant ;
begin
if pos('GCARTICLE',string(Parms[1]))=1 then Result := 'GCARTICLE'
else if pos('GCDISPO',string(Parms[1]))=1 then Result := 'GCDISPO'
else if pos('BTDISPO',string(Parms[1]))=1 then Result := 'BTDISPO'
else if pos('SELECTION',string(Parms[1]))=1 then Result := 'SELECTION'
else if pos('GENERECAB',string(Parms[1]))=1 then Result := 'GENERECAB'
else if pos('TARIF',string(Parms[1]))=1 then Result := 'TARIF'
else if pos('ARTICLE',string(Parms[1]))=1 then Result := 'ARTICLE'
else Result := '';
end;

Procedure RechArt_Mode (Parms : Array of variant; nb: integer);
var F : TForm;
    G_Article : THCritMaskEdit;
    StRange : String;
BEGIN
F := TForm (Longint (Parms[0]));
G_Article := THCritMaskEdit (F.FindComponent (Parms[1]));
stRange := 'GA_ARTICLE=' + Trim (Copy (G_Article.Text, 1, 18))+';GA_TYPEARTICLE=MAR';
if Parms[2]='CODEARTICLE' then stRange := stRange + ';RETOUR_CODEARTICLE=X';
DispatchRecherche(G_Article, 1, '',StRange, '');
END;

procedure InitTOFArticleMul ;
begin
RegisterAglProc('ArticleModeChangeField', True , 1, TOFArticleMul_ChangeField) ;
RegisterAglFunc('TraitementOk', True , 1, TOF_GCARTICLE_MUL_AGLTraitementOk) ;
RegisterAglFunc('ExtractArgument',True,1,TOF_GCARTICLE_MUL_AGLExtractArgument);
RegisterAglProc('RechArt_Mode', True , 1, RechArt_Mode);
End ;

// Procédure permettant l'ouverture de l'état
procedure AGLRecupEtat2(parms:array of variant; nb: integer ) ;
var
     modele: string;
begin
  modele := string(Parms[1]);
  {$IFNDEF EAGLCLIENT}
  EditEtat('E','GED',modele,TRUE,nil,'','') ;
  {$ENDIF}
end ;


Initialization
registerclasses([TOF_GCARTICLE_MUL]) ;
RegisterAglProc('RecupEtat2',TRUE,1,AGLRecupEtat2);
InitTOFArticleMul ;
end.

