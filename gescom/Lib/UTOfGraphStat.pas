unit UTofGraphStat;

interface

uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls, graphics,
      HCtrls,HEnt1,HMsgBox,UTOF, UTOM,AglInit,Dialogs,HTB97, M3FP, EntGC,
{$IFDEF EAGLCLIENT}
      MaineAGL,eQRS1,
{$ELSE}
      Fiche, mul, DBGrids,Fe_Main,db,dbTables,QRS1,
{$IFDEF V530}
      EdtEtat,
{$ELSE}
      EdtREtat,
{$ENDIF}      
{$ENDIF}
{$IFDEF NOMADE}
      UtilPOP,
{$ENDIF} //NOMADE
      grids, UtilGc,UTOB, TeEngine, Chart, Series, GraphUtil,AGLInitGC,GRS1,   //gm 20/03/03 GRS1 ok en eagl;
      UTofEditStat;

procedure AffecteDataType (F : TForm; indice : integer; Etat : Boolean; Couleur : TColor);

Const
    nbValeurs = 8;
	TexteMessage: array[1..3] of string 	= (
          {1}         'Vous devez sélectionner une nature de pièce.',
                      'Vous devez renseigner un champ d''application',
                      'Vous devez choisir une série'
                 );

Type
     Tof_GraphStat = Class (TOF)
        public
            TobTmpSta : TOB;
            ValApplication, stperiode, stNatureauxi : string;
             procedure OnUpdate ; override;
             procedure OnArgument (stArgument : string) ; override;
            // ajout
            procedure ChargeLignes (stWhere : string; var inbper : integer);
            procedure PrepareGraphe (var stTitresCol1, stTitresCol2, stColaff1, stColAff2 : string;
                                     stTitre, stChamp : string);
            procedure RegroupLignes (TobL : TOB);
            procedure SelectionLignes (var stWhere : string);
            procedure VideTob (TAVider : TOB);
     end ;

var stTypeValeur, stChampLigne : string;

implementation

procedure Tof_GraphStat.OnArgument (stArgument : string);
var iInd : integer;
    THLIB : THLabel ;
    stTitre, stInd, stPlus, stNatDef : string ;
begin
stNatureAuxi := stArgument;
if stArgument = 'VEN' then
    begin
    stPlus := ' AND GPP_VENTEACHAT="VEN"'  ;
    SetControlProperty ( 'GP_NATUREPIECEG', 'PLUS', stPlus ) ;
{$IFDEF NOMADE}
    stPlus := stPlus + ' AND (GPP_NATUREPIECEG IN ("CC","DE") OR GPP_TRSFVENTE="X")' ;
    stNatDef := 'CC' ;
{$ELSE}
    FactureAvoir_AddItem(GetControl('GP_NATUREPIECEG'),'VEN');
    stNatDef := 'BLC' ;
{$ENDIF}
    SetControlText ('GP_NATUREPIECEG', stNatDef ) ;
    for iInd:=1 to 9 do SetControlProperty('LIBRETIERS'+IntToStr(iInd),'DATATYPE','GCLIBRETIERS'+IntToStr(iInd));
    SetControlProperty('LIBRETIERSA','DATATYPE','GCLIBRETIERSA');
      //*mcd 27/03/02 ** Blocages des natures de pièces autorisées en affaires
      if (ctxAffaire in V_PGI.PGIContexte)  then
      begin
      	if (ctxScot in V_PGI.PGIContexte)  then
        	THValComboBox (Ecran.FindComponent('GP_NATUREPIECEG')).Plus :=  'AND ((GPP_NATUREPIECEG="FAC") or (GPP_NATUREPIECEG="AVC") or (GPP_NATUREPIECEG="FPR") or (GPP_NATUREPIECEG="APR") or (GPP_NATUREPIECEG="FRE"))';
        THValComboBox(GetControl ('GP_NATUREPIECEG')).Value := 'FAC';
                 // mcd 10/10/02 obligation de mettre après car condition plus changée
        Iind:=THValComboBox(GetControl ('GP_NATUREPIECEG')).Values.IndexOf('FAC');
        if Iind<0 then Iind:=0;
        THValComboBox(GetControl ('GP_NATUREPIECEG')).Items.Insert(Iind,'Facture + Avoir clients + Fac. Reprise');
        THValComboBox(GetControl ('GP_NATUREPIECEG')).Values.Insert(Iind,'ZZ1');
        SetControlText('GP_NATUREPIECEG','ZZ1');

        Iind:=THValComboBox(GetControl ('GP_NATUREPIECEG')).Values.IndexOf('FRE');
        if Iind<0 then Iind:=0;
        THValComboBox(GetControl ('GP_NATUREPIECEG')).Items.Insert(Iind,'Facture + Fac. Reprise');
        THValComboBox(GetControl ('GP_NATUREPIECEG')).Values.Insert(Iind,'ZZ2');
        SetControlText('GP_NATUREPIECEG','ZZ2');
       end;
      // **** Fin affaire ***
    // initialisation pour le regroupement, possibilité d'avoir regroupement fournisseur seulement dans les ventes
    if (VH_GC.GCMultiDepots) then
    THValComboBox(GetControl ('APPLICATION')).Plus := 'AND CO_CODE<>"ETA"'
    else THValComboBox(GetControl ('APPLICATION')).Plus := 'AND CO_CODE<>"DEP"';
    end
else
    begin
    stPlus := ' AND GPP_VENTEACHAT="ACH"'  ;
    SetControlProperty ('GP_NATUREPIECEG', 'Plus', stPlus ) ;
{$IFDEF NOMADE}
    stPlus := stPlus + ' AND (GPP_NATUREPIECEG="CF" OR GPP_TRSFACHAT="X")' ;
    stNatDef := 'CF' ;
{$ELSE}
    FactureAvoir_AddItem(GetControl('GP_NATUREPIECEG'),'ACH');
    stNatDef := 'BLF' ;
{$ENDIF}
    SetControlText ( 'GP_NATUREPIECEG', stNatDef ) ;
    THValComboBox(GetControl ('APPLICATION')).Plus := ' AND CO_CODE<>"TTI"';
    TCheckBox(GetControl ('RMARGE')).Enabled := False;
    for iInd:=1 to 3 do SetControlProperty('LIBRETIERS'+IntToStr(iInd),'DATATYPE','GCLIBREFOU'+IntToStr(iInd));
      //*** Blocages des natures de pièces autorisées en affaires
      if (ctxAffaire in V_PGI.PGIContexte)  then
         begin
         THValComboBox (Ecran.FindComponent('GP_NATUREPIECEG')).Plus := 'AND (GPP_VENTEACHAT="ACH") '+ AfPlusNAtureAchat;
         SetControlText ('GP_NATUREPIECEG','FF');
         end;
      // **** Fin affaire ***
    // initialisation pour le regroupement, possibilité d'avoir regroupement fournisseur seulement dans les ventes
    if (VH_GC.GCMultiDepots) then
    THValComboBox(GetControl ('APPLICATION')).Plus := ' AND CO_CODE <>"FOU" AND CO_CODE<>"ETA"'
    else THValComboBox(GetControl ('APPLICATION')).Plus := ' AND CO_CODE <>"FOU" AND CO_CODE<>"DEP"';
    end;
inherited;
SetControlText('DATE1PERIODE',DateToStr(PlusMois(DebutdeMois(Date),-1)));
SetControlText('DATE1PERIODE_',DateToStr(PlusMois(FindeMois(Date),-1)));

{Libres articles}
for iInd := 1 to 9 do
    begin
    stInd:=IntToStr(iInd) ;
    GCTitreZoneLibre('GL_LIBREART'+stInd,stTitre) ;
    if stTitre<>'' then SetControlCaption('TLIBREART'+stInd,stTitre)
    else
        BEGIN
        THLIB:=THLabel(TForm(Ecran).FindComponent('TLIBREART'+stInd)) ;
        THLIB.Visible:=False ;
        if (THLIB.FocusControl<>nil) then (THLIB.FocusControl).Visible:=False ;
        END ;
    {$IFDEF CCS3}
    if iInd>4 then
       BEGIN
       SetControlVisible('LIBREART'+stInd,False) ;
       SetControlVisible('TLIBREART'+stInd,False) ;
       END ;
    {$ENDIF}
    end;
{$IFDEF CCS3}
SetControlVisible('LIBREARTA',False) ;
SetControlVisible('TLIBREARTA',False) ;
{$ELSE}
GCTitreZoneLibre ('GL_LIBREARTA', stTitre);
if stTitre<>'' then SetControlCaption('TLIBREARTA',stTitre)
else
    BEGIN
    THLIB:=THLabel(TForm(Ecran).FindComponent('TLIBREARTA')) ;
    THLIB.Visible:=False ;
    if (THLIB.FocusControl<>nil) then (THLIB.FocusControl).Visible:=False ;
    END ;
{$ENDIF}

{Libres tiers}
for iInd := 1 to 9 do
    begin
    stInd:=IntToStr(iInd) ;
    if (stArgument='ACH') and (iInd>3) then
        BEGIN
        SetControlVisible('TLIBRETIERS'+stInd,False) ;
        SetControlVisible('LIBRETIERS'+stInd,False) ;
        END else
        BEGIN
        if stArgument='ACH'
        then GCTitreZoneLibre('YTC_TABLELIBREFOU'+stInd,stTitre)
        else GCTitreZoneLibre('GP_LIBRETIERS'+stInd,stTitre) ;
        if stTitre<>'' then SetControlCaption('TLIBRETIERS'+stInd,stTitre)
        else
            BEGIN
            THLIB:=THLabel(TForm(Ecran).FindComponent('TLIBRETIERS'+stInd)) ;
            THLIB.Visible:=False ;
            if (THLIB.FocusControl<>nil) then (THLIB.FocusControl).Visible:=False ;
            END ;
        {$IFDEF CCS3}
        if iInd>4 then
           BEGIN
           SetControlVisible('LIBRETIERS'+stInd,False) ;
           SetControlVisible('TLIBRETIERS'+stInd,False) ;
           END ;
        {$ENDIF}
        END ;
    end;
if stArgument='ACH' then
    BEGIN
    SetControlVisible('TLIBRETIERSA',False) ;
    SetControlVisible('LIBRETIERSA',False) ;
    END else
    BEGIN
    {$IFDEF CCS3}
    SetControlVisible('LIBRETIERSA',False) ;
    SetControlVisible('TLIBRETIERSA',False) ;
    {$ELSE}
    GCTitreZoneLibre('GP_LIBRETIERSA'+stInd,stTitre) ;
    if stTitre<>'' then SetControlCaption('TLIBRETIERSA',stTitre)
    else
        BEGIN
        THLIB:=THLabel(TForm(Ecran).FindComponent('TLIBRETIERSA')) ;
        THLIB.Visible:=False ;
        if (THLIB.FocusControl<>nil) then (THLIB.FocusControl).Visible:=False ;
        END ;
    {$ENDIF}
    END ;

TToolbarButton97(GetControl ('bAffGraph')).Down:=True;
THGrid(GetControl ('FListe')).Visible:=False;
end;

procedure Tof_GraphStat.OnUpdate;
var F : TFGRS1;
    stWhere, stColAff1, stColAff2, stTitreCol1, stTitreCol2 : string;
    tstTitre1, tstTitre2 : Tstrings;
    iInd, inbper1, inbper2 : integer;
begin
Inherited;
F := TFGRS1 (Ecran);
if GetControlText ('GP_NATUREPIECEG') = '' then
    begin
    LastError:=1;
    LastErrorMsg:=TexteMessage[LastError] ;
    exit;
    end;
if GetControlText ('APPLICATION') = '' then
    begin
    LastError:=2;
    LastErrorMsg:=TexteMessage[LastError] ;
    exit;
    end;
if (TRadioButton(GetControl ('RCA')).Checked = False) and
   (TRadioButton(GetControl ('RQUANTITE')).Checked = False) and
   (TRadioButton(GetControl ('RMARGE')).Checked = False) then
    begin
    LastError:=3;
    LastErrorMsg:=TexteMessage[LastError] ;
    exit;
    end;
if (ValApplication <> '') and (ValApplication <> GetControlText('APPLICATION')) then
    begin
    ValApplication := GetControlText('APPLICATION');
    for iInd := 1 to nbValeurs do
        begin
        AffecteDataType (F, iInd, False, clBtnFace);
        end;
    end;

TobTmpSta := Tob.Create ('', nil, -1);
inbper1 := 0;
inbper2 := 0;
stPeriode := '1';
SelectionLignes (stWhere);
ChargeLignes (stWhere, inbper1);
if (GetControlText ('DATE1PERIODE') <> GetControlText ('DATE2PERIODE')) or
   (GetControlText ('DATE1PERIODE_') <> GetControlText ('DATE2PERIODE_')) then
    begin
    stPeriode := '2';
    SelectionLignes (stWhere);
    ChargeLignes (stWhere, inbper2);
    end;

tstTitre1 := TStringList.Create;
tstTitre2 := TStringList.Create;
stTitreCol1 := '';
stTitreCol2 := '';
stColAff1 := '';
stColAff2 := '';
if TRadioButton(GetControl ('RCA')).Checked then
    begin
    PrepareGraphe (stTitreCol1, stTitreCol2, stColAff1, stColAff2,
                   'Chiffre d''Affaires', 'GZB_CAGLOBAL');
    end else
    begin
    if TRadioButton(GetControl ('RQUANTITE')).Checked then
        begin
        PrepareGraphe (stTitreCol1, stTitreCol2, stColAff1, stColAff2, 'Quantité',
                       'GZB_COUTQTE');
        end else
        begin
        if TRadioButton(GetControl ('RMARGE')).Checked then
            PrepareGraphe (stTitreCol1, stTitreCol2, stColAff1, stColAff2, 'Marge',
                           'GZB_MARGE');
        end;
     end;

if stNatureauxi = 'VEN' then
    begin
    tstTitre1.Add ('Statistiques de vente');
    tstTitre1.Add ('du ' + GetControlText ('DATE1PERIODE') + ' au ' + GetControlText ('DATE1PERIODE_'));
    tstTitre2.Add ('Statistiques de vente');
    tstTitre2.Add ('du ' + GetControlText ('DATE2PERIODE') + ' au ' + GetControlText ('DATE2PERIODE_'));
    end else
    begin
    tstTitre1.Add ('Statistiques d''achat');
    tstTitre1.Add ('du ' + GetControlText ('DATE1PERIODE') + ' au ' + GetControlText ('DATE1PERIODE_'));
    tstTitre2.Add ('Statistiques d''achat');
    tstTitre2.Add ('du ' + GetControlText ('DATE2PERIODE') + ' au ' + GetControlText ('DATE2PERIODE_'));
    end;

if TobTmpSta.Detail.Count > 0 then
    begin
    LanceGraph (F, TobTmpSta, '', 'GZB_VALRUPT1;' + stColAff1 + ';' + stColAff2, '',
                RechDom('GCGROUPSTA', THValComboBox(GetControl ('APPLICATION')).Value,
                        False) + ';' + stTitreCol1 + ';' + stTitreCol2,
                stColAff1, stColAff2, tstTitre1, tstTitre2,
                TBarSeries, 'GZB_VALRUPT1', False);
    if inbper1 > 0 then F.FChart1.Visible := True
    else F.FChart1.Visible := False;
    if inbper2 > 0 then F.FChart2.Visible := True
    else F.FChart2.Visible := False;

    if TCheckBox(GetControl ('RPOURCENTAGE')).Checked = True then
        begin
        TRadioButton(GetControl ('RVAL')).Checked := True;
        F.FChart1.Legend.LegendStyle := lsValues;
        F.FChart2.Legend.LegendStyle := lsValues;
        F.FChart1.Legend.TextStyle := ltsLeftPercent;
        F.FChart2.Legend.TextStyle := ltsLeftPercent;
        end else
        begin
        if TRadioButton(GetControl ('RVAL')).Checked = True then
            begin
            F.FChart1.Legend.LegendStyle := lsValues;
            F.FChart2.Legend.LegendStyle := lsValues;
            end else
            begin
            F.FChart1.Legend.LegendStyle := lsSeries;
            F.FChart2.Legend.LegendStyle := lsSeries;
            end;
        F.FChart1.Legend.TextStyle := ltsLeftValue;
        F.FChart2.Legend.TextStyle := ltsLeftPercent;
        end;
    end;
tstTitre1.Free;
tstTitre2.Free;
VideTob (TobTmpSta);
end;

// ****************************** Ajout **************************************************
procedure Tof_GraphStat.ChargeLignes (stWhere : string; var inbper : integer);
var TSql : TQuery;
    TobL : TOB;
    stSelect : string;
begin
stSelect := 'SELECT GL_NATUREPIECEG, GL_DATEPIECE, GL_ARTICLE, GL_CODEARTICLE, GL_ETABLISSEMENT, GL_DEPOT, ' +
            'GL_FOURNISSEUR, GL_REPRESENTANT, GL_TARIFARTICLE, GL_TIERSFACTURE, ' +
            'GL_TIERS, GL_TIERSLIVRE, GL_TIERSPAYEUR, GL_TARIFTIERS, GL_TYPEARTICLE, ' +
            'GL_TOTALHT, GL_QTEFACT, GL_PMAP, GL_DPA, GL_PMRP, GL_DPR, GL_LIBREART1, ' +
            'GL_LIBREART2, GL_LIBREART3, GL_LIBREART4, GL_LIBREART5, GL_LIBREART6 ' +
            'GL_LIBREART7, GL_LIBREART8, GL_LIBREART9, GL_LIBREARTA ';
if (pos ('GP_LIBRETIERS', stWhere) <> 0) or
   (copy (GetControlText ('APPLICATION'), 1, 2) = 'LT') then
    begin
    stSelect := stSelect + ', GP_LIBRETIERS1, ' +
            'GP_LIBRETIERS2, GP_LIBRETIERS3, GP_LIBRETIERS4, GP_LIBRETIERS5, GP_LIBRETIERS6, ' +
            'GP_LIBRETIERS7, GP_LIBRETIERS8, GP_LIBRETIERS9, GP_LIBRETIERSA ' +
            'FROM LIGNE, PIECE ';
    stWhere := stWhere + ' AND GP_NATUREPIECEG=GL_NATUREPIECEG AND ' +
            'GP_SOUCHE=GL_SOUCHE AND GP_NUMERO=GL_NUMERO AND GP_INDICEG=GL_INDICEG ';
    end else
    begin
    stSelect := stSelect + ' FROM LIGNE ';
    end;
stSelect :=  stSelect + 'WHERE ' + stWhere;
TSql := OpenSql (stSelect, True);

Tobl := Tob.Create ('', nil, -1);
inbPer := 0;
TSql.First;
while not TSql.Eof do
    begin
    if Trim (TSql.FindField ('GL_CODEARTICLE').AsString) <> '' then
        begin
        TobL.SelectDb ('', TSql);
        RegroupLignes (Tobl);
        inbPer := inbPer + 1;
        end;
    TSql.Next;
    end;
TobL.Free;
Ferme (TSql);
end;

procedure Tof_GraphStat.PrepareGraphe (var stTitresCol1, stTitresCol2, stColaff1, stColAff2 : string;
                                       stTitre, stChamp : string);
begin
stTitresCol1 := stTitre;
stColAff1 := stChamp + '1';
stTitresCol2 := stTitre;
stColAff2 := stChamp + '2';
end;

procedure Tof_GraphStat.RegroupLignes (TobL : TOB);
var stSelect, stValChamp : string;
    iIndex : integer;
    TobTS : TOB;
    dPrixAchat : double;
    TSql : TQuery;
begin
if pos ('GA_', stChampLigne) > 0 then
    begin
    stSelect := 'SELECT ' + stChampLigne + ' FROM ARTICLE WHERE GA_ARTICLE="' +
                TobL.GetValue ('GL_ARTICLE') + '"';
    Tsql := OpenSql (stSelect, True);
    if not Tsql.Eof then stValChamp := Tsql.FindField(stChampLigne).AsString
    else stValChamp := '';
    Ferme (TSql);
    end else
    begin
    if stChampLigne = 'GL_ARTICLE' then stValChamp := TobL.GetValue ('GL_CODEARTICLE')
    else stValChamp := TobL.GetValue(stChampLigne);
    end;

iIndex := 1;
while iIndex <= nbValeurs do
    begin
    if stTypeValeur = 'E' then
        begin
        if (THEdit(GetControl ('EVALEUR'+IntToStr(iIndex))).Text <> '') and
           (THEdit(GetControl ('EVALEUR'+IntToStr(iIndex))).Text = stValChamp) then break;
        end else
        begin
        if (THValComboBox(GetControl ('CVALEUR'+IntToStr(iIndex))).Value <> '') and
           (THValComboBox(GetControl ('CVALEUR'+IntToStr(iIndex))).Value = stValChamp) then break;
        end;
    Inc(iIndex);
    end;
if iIndex > nbValeurs then stValChamp := 'Autres';

TobTS := TobTmpSta.FindFirst (['GZB_VALRUPT1'], [stValChamp], False);
if TobTS = Nil then
    begin
    TobTS := Tob.Create ('GCTMPSTA', TobTmpSta, -1);
    TobTS.InitValeurs;
    TobTS.PutValue ('GZB_VALRUPT1', stValChamp);
    end;

if VH_GC.GCMargeArticle[1] = 'P' then dPrixAchat := TobL.GetValue ('GL_' + VH_GC.GCMargeArticle + 'P')
else dPrixAchat := TobL.GetValue ('GL_' + VH_GC.GCMargeArticle);

TobTS.PutValue ('GZB_CAGLOBAL' + stPeriode,
                TobTS.GetValue ('GZB_CAGLOBAL' + stPeriode) + TobL.GetValue ('GL_TOTALHT'));

TobTS.PutValue ('GZB_COUTQTE' + stPeriode,
                TobTS.GetValue ('GZB_COUTQTE' + stPeriode) + TobL.GetValue ('GL_QTEFACT'));
if dPrixAchat <> 0.0 then
    begin
    TobTS.PutValue ('GZB_CAMARGE' + stPeriode,
                    TobTS.GetValue ('GZB_CAMARGE' + stPeriode) + TobL.GetValue ('GL_TOTALHT'));
    TobTS.PutValue ('GZB_MARGE' + stPeriode,
                   TobTS.GetValue ('GZB_MARGE' + stPeriode) +
                        TobL.GetValue ('GL_TOTALHT') - dPrixAchat);
    end;
end;

procedure Tof_GraphStat.SelectionLignes (var stWhere : string);
var iInd : integer;
    NatureP : string;
begin
stWhere := 'GL_DATEPIECE>="' + USDateTime (StrToDate (GetControlText ('DATE' + stPeriode + 'PERIODE'))) + '" AND ' +
           'GL_DATEPIECE<="' + USDateTime (StrToDate (GetControlText ('DATE' + stPeriode + 'PERIODE_'))) + '"';
NatureP := GetControlText ('GP_NATUREPIECEG');

//gm 06/01/03
if not(ctxAffaire in V_PGI.PGIContexte)  then
Begin
  stWhere := stWhere + ' AND ' + FactureAvoir_RecupSQL(NatureP, 'GL');
end else
Begin
  // gm 05/09/02  Prise en compte Facture reprise
  if NatureP ='ZZ1' then
    stWhere := stWhere + ' AND (GL_NATUREPIECEG="FAC" OR GL_NATUREPIECEG="AVC" OR GL_NATUREPIECEG="AVS" OR GL_NATUREPIECEG="FRE")'
  else
    if NatureP ='ZZ2' then
      stWhere := stWhere + ' AND (GL_NATUREPIECEG="FAC" OR GL_NATUREPIECEG="FRE")'
    else
      stWhere := stWhere + ' AND GL_NATUREPIECEG="' + NatureP + '"';
ENd;

for iInd := 1 to 9 do
    begin
    if GetControlText ('LIBREART' + IntToStr (iInd)) <> '' then
        begin
        stWhere := stWhere + ' AND GL_LIBREART' + IntToStr (iInd) + '="' +
                    GetControlText ('LIBREART' + IntToStr (iInd)) + '"';
        end;
    end;
if GetControlText ('LIBREARTA') <> '' then
    begin
    stWhere := stWhere + ' AND GL_LIBREARTA="' + GetControlText ('LIBREARTA') + '"';
    end;
for iInd := 1 to 9 do
    begin
    if GetControlText ('LIBRETIERS' + IntToStr (iInd)) <> '' then
        begin
        stWhere := stWhere + ' AND GP_LIBRETIERS' + IntToStr (iInd) + '="' +
                    GetControlText ('LIBRETIERS' + IntToStr (iInd)) + '"';
        end;
    end;
if GetControlText ('LIBRETIERSA') <> '' then
    begin
    stWhere := stWhere + ' AND GP_LIBRETIERSA="' + GetControlText ('LIBRETIERSA') + '"';
    end;
if stWhere <> '' then stWhere := ' (' + stWhere + ')';
end;

procedure Tof_GraphStat.VideTob (TAVider : TOB);
var iIndex : integer;
begin
for iIndex := TAVider.Detail.Count - 1 Downto 0 do
    begin
    TAVider.Detail[iIndex].Free ;
    end;
TAVider.Free;
end;

// ****************************** register *****************************************
procedure AffecteDataType (F : TForm; indice : integer; Etat : Boolean; Couleur : TColor);
var stSelect, stCoLibre : string;
    TSql : TQuery;
begin
stChampLigne := RechDom ('GCGROUPSTA',
                         THValComboBox(F.FindComponent ('APPLICATION')).Value, True);
stSelect := 'SELECT CO_LIBRE FROM COMMUN WHERE CO_TYPE="GST" AND CO_CODE="' +
            THValComboBox(F.FindComponent('APPLICATION')).Value + '"';
TSql := OpenSql (stSelect, True);
if not TSql.Eof then
    begin
    stCoLibre := TSql.FindField('CO_LIBRE').AsString;
    stTypeValeur := ReadTokenSt (stCoLibre);
    if stTypeValeur = 'E' then
        begin
        with THEdit(F.FindComponent('EVALEUR'+IntToStr(Indice))) do
            begin
            DataType := ReadTokenSt(stCoLibre);
            if (Etat = True) or (Indice <> 1) then
                begin
                Visible := True;
                TabStop := True;
                Enabled := Etat;
                Color := Couleur;
                end;
            if THValComboBox(F.FindComponent('APPLICATION')).Value = 'TPA' then
                begin
                Plus := 'T_ISPAYEUR="X"';
                end;
            if THValComboBox(F.FindComponent('APPLICATION')).Value = 'TIE' then
                begin
                if Tof_GraphStat(TFGRS1(F).LaTOF).stNatureauxi = 'VEN' then
                    begin
                    DataType := 'GCTIERSCLI';
                    end else
                    begin
                    DataType := 'GCTIERSFOURN';
                    end;
                end;
            end;
        with THValComboBox(F.FindComponent('CVALEUR'+IntToStr(Indice))) do
            begin
            Visible := False;
            TabStop := False;
            end;
        end else
        begin
        with THValComboBox(F.FindComponent('CVALEUR'+IntToStr(Indice))) do
            begin
            DataType := ReadTokenSt(stCoLibre);
            if (Etat = True) or (Indice <> 1) then
                begin
                Visible := True;
                TabStop := True;
                Enabled := Etat;
                Color := Couleur;
                end;
            end;
        with THEdit(F.FindComponent('EVALEUR'+IntToStr(Indice))) do
            begin
            Visible := False;
            TabStop := False;
            end;
        end;
    if Etat = False then
        begin
        THValComboBox(F.FindComponent('CVALEUR'+IntToStr(Indice))).Value := '';
        THEdit(F.FindComponent('EVALEUR'+IntToStr(Indice))).Text := '';
        end;
    end;
Ferme (TSql);
end;

Procedure Tof_GraphStat_AffectValeur (parms:array of variant; nb: integer ) ;
var F : TForm;
    Indice : integer;
begin
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCGRAPHSTAT') then exit;
Indice := Integer (Parms[1]);
AffecteDataType (F, Indice, True, clWindow);
end;

Procedure TOF_GraphStat_ChangeAffichage (parms:array of variant; nb: integer );
var F : TFGRS1;
begin
F := TFGRS1 (Longint (Parms[0]));
if (F.Name <> 'GCGRAPHSTAT') then exit;
if TCheckBox(F.FindComponent('RPOURCENTAGE')).Checked = True then
    begin
    TRadioButton(F.FindComponent('RVAL')).Checked := True;
    F.FChart1.Legend.LegendStyle := lsValues;
    F.FChart2.Legend.LegendStyle := lsValues;
    F.FChart1.Legend.TextStyle := ltsLeftPercent;
    F.FChart2.Legend.TextStyle := ltsLeftPercent;
    end else
    begin
    if TRadioButton(F.FindComponent('RVAL')).Checked = True then
        begin
        F.FChart1.Legend.LegendStyle := lsValues;
        F.FChart2.Legend.LegendStyle := lsValues;
        end else
        begin
        F.FChart1.Legend.LegendStyle := lsSeries;
        F.FChart2.Legend.LegendStyle := lsSeries;
        end;
    F.FChart1.Legend.TextStyle := ltsLeftValue;
    F.FChart2.Legend.TextStyle := ltsLeftValue;
    end;
end;

Procedure Tof_GraphStat_ControleApplication (parms:array of variant; nb: integer ) ;
var F : TFGRS1;
    iInd : integer;
begin
F := TFGRS1 (Longint (Parms[0]));
if (F.Name <> 'GCGRAPHSTAT') then exit;
if THValComboBox(F.FindComponent('APPLICATION')).Value <>
        Tof_GraphStat(F.LaTOF).ValApplication then
    begin
    Tof_GraphStat(F.LaTOF).ValApplication := THValComboBox(F.FindComponent('APPLICATION')).Value;
    for iInd := 1 to nbValeurs do
        begin
        AffecteDataType (F, iInd, False, clBtnFace);
        end;
    end;
end;

procedure Tof_GraphStat_RecupereApplication (parms:array of variant; nb: integer ) ;
var F : TFGRS1;
begin
F := TFGRS1 (Longint (Parms[0]));
if (F.Name <> 'GCGRAPHSTAT') then exit;
Tof_GraphStat(F.LaTof).ValApplication := THValComboBox(F.FindComponent('APPLICATION')).Value;
end;

procedure InitTOFGraphStat ();
begin
RegisterAglProc('GraphStat_AffectValeur', True , 1, Tof_GraphStat_AffectValeur);
RegisterAglProc('GraphStat_ChangeAffichage', True , 0, TOF_GraphStat_ChangeAffichage);
RegisterAglProc('GraphStat_ControleApplication', True , 0, Tof_GraphStat_ControleApplication);
RegisterAglProc('GraphStat_RecupereApplication', True , 0, Tof_GraphStat_RecupereApplication);
end;

Initialization
registerclasses ([Tof_GraphStat]);
InitTOFGraphStat ();

end.
