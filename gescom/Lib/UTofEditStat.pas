unit UTofEditStat;

interface

uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls, graphics,
      HCtrls,HEnt1,HMsgBox,UTOF,UTOM,AglInit,Dialogs,
      M3FP, EntGC, UtilGC, HDimension,
{$IFDEF EAGLCLIENT}
      MaineAGL,eMul,
{$ELSE}
       Fiche, DBGrids,db,mul, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$IFDEF V530}
      EdtEtat,
{$ELSE}
      EdtREtat,
{$ENDIF}
{$ENDIF}
      grids, UTOB;

Const
    nbregroup = 6;
	TexteMessage: array[1..1] of string 	= (
          {1}         'Vous devez sélectionner une nature de pièce.'
                 );

Type
     Tof_EditStat = Class (TOF)
        public
            TobTmpSta : TOB;
            stChampRup : Array [1..nbregroup] of string;
            iCompteur : integer;

            procedure OnArgument (stArgument : string); override;
            procedure OnClose; override;
            procedure OnNew; override;
            procedure OnUpdate; override;
            // ajout
            procedure ChargeLignes (stPeriode, stSelect, stWhere : string);
            procedure InitialiseValeur;
            procedure RegroupLignes (TLigne : TOB; stPeriode : string);
            procedure SelectionLignes (stPeriode : string; var stSelect, stWhere : string);
     end ;

     function  FactureAvoir_RecupSQL(NatureP, PrefixeSQL : string) : string;
     procedure FactureAvoir_AddItem(CC : TControl ; VenteAchat : string);

implementation

procedure TOF_EditStat.OnArgument (stArgument : string);
var iInd : integer;
    THLIB : THLabel ;
    stTitre, stInd, stPlus, stNatDef : string;
{$IFDEF CCS3}
    Ch      : String ;
{$ENDIF}
begin
{$IFDEF NOMADE}
stPlus := 'AND GPP_VENTEACHAT="' + stArgument ;
if stArgument = 'VEN'
  then stPlus := stPlus + '" AND (GPP_NATUREPIECEG IN ("CC","DE") OR GPP_TRSFVENTE="X")'
  else stPlus := stPlus + '" AND (GPP_NATUREPIECEG="CF" OR GPP_TRSFACHAT="X")' ;
{$ELSE}
stPlus := 'AND (GPP_VENTEACHAT="' + stArgument +
          '" OR GPP_NATUREPIECEG="EEX" OR GPP_NATUREPIECEG="SEX")' ;
{$ENDIF}
SetControlProperty ( 'NATUREPIECEG', 'PLUS', stPlus ) ;
SetControlText('DATE1PERIODE',DateToStr(PlusMois(DebutdeMois(Date),-1)));
SetControlText('DATE1PERIODE_',DateToStr(PlusMois(FindeMois(Date),-1)));
// Quand on est en mono-dépôt on doit voir la notion d'établissement et non la notion de dépôt
if not VH_GC.GCMultiDepots then SetControlCaption ( 'TDEPOT', TraduireMemoire ( 'Etablissement' ) ) ;
// Paramétrage des libellés des familles
for iInd:=1 to 3 do
    begin
    THLIB:=THLabel(GetControl('TFAMILLENIV'+IntToStr(iInd)));
    THLIB.Caption:=RechDom('GCLIBFAMILLE','LF'+IntToStr(iInd),False);
    if THLIB.Caption='.-'
    then BEGIN THLIB.Visible:=False ; THLIB.FocusControl.Visible:=False ; END ;
    end;
if stArgument = 'VEN' then
    begin
{$IFDEF NOMADE}
    stNatDef := 'CC' ;
{$ELSE}
    if not (ctxAffaire in V_PGI.PGIContexte)  then
    begin
      FactureAvoir_AddItem(GetControl('NATUREPIECEG'),'VEN');
      stNatDef := 'FAC' ;
    end;
{$ENDIF}

    SetControlText ( 'NATUREPIECEG', stNatDef ) ;
    SetControlText ( 'XX_VARIABLE8', 'de vente' ) ;
    SetControlProperty ( 'TIERS', 'DataType', 'GCTIERSCLI' ) ;
    SetControlProperty ( 'TIERS_', 'DataType', 'GCTIERSCLI' ) ;
    for iInd:=1 to 9 do SetControlProperty('LIBRETIERS'+IntToStr(iInd),'DATATYPE','GCLIBRETIERS'+IntToStr(iInd));
    SetControlProperty('LIBRETIERSA','DATATYPE','GCLIBRETIERSA');
      //mcd 27/03/02 *** Blocages des natures de pièces autorisées en affaires
    if (ctxAffaire in V_PGI.PGIContexte)  then
    begin
    	if (ctxScot in V_PGI.PGIContexte)  then
    		SetControlProperty ( 'NATUREPIECEG', 'Plus', 'AND ((GPP_NATUREPIECEG="FAC") or (GPP_NATUREPIECEG="FPR") or (GPP_NATUREPIECEG="AVC") or (GPP_NATUREPIECEG="APR") or (GPP_NATUREPIECEG="FRE"))' ) ;
                // mcd 10/10/02 obligation de mettre après car condition plus changée
      Iind:=THValComboBox(GetControl ('NATUREPIECEG')).Values.IndexOf('FAC');
      if Iind<0 then Iind:=0;
      THValComboBox(GetControl ('NATUREPIECEG')).Items.Insert(Iind,'Facture + Avoir clients + Fac. Reprise');
      THValComboBox(GetControl ('NATUREPIECEG')).Values.Insert(Iind,'ZZ1');
      SetControlText('NATUREPIECEG','ZZ1');

      Iind:=THValComboBox(GetControl ('NATUREPIECEG')).Values.IndexOf('FRE');
      if Iind<0 then Iind:=0;
      THValComboBox(GetControl ('NATUREPIECEG')).Items.Insert(Iind,'Facture + Fac. Reprise');
      THValComboBox(GetControl ('NATUREPIECEG')).Values.Insert(Iind,'ZZ2');
      SetControlText('NATUREPIECEG','ZZ2');
    end;
      // **** Fin affaire ***
    end else
    begin
    THEdit(GetControl ('XX_VARIABLE8')).Text := TraduireMemoire ( 'd''achat' ) ;
    TRadioButton(GetControl ('RQUANTITE')).Checked := True;
    TGroupBox(GetControl ('PAFFICHAGE')).Visible := False;
{$IFDEF NOMADE}
    stNatDef := 'CF' ;
{$ELSE}
    FactureAvoir_AddItem(GetControl('NATUREPIECEG'),'ACH');
    stNatDef := 'BLF' ;
{$ENDIF}
    SetControlText ( 'NATUREPIECEG', stNatDef ) ;
    THEdit(GetControl ('TIERS')).DataType := 'GCTIERSFOURN';
    THEdit(GetControl ('TIERS_')).DataType := 'GCTIERSFOURN';
    SetControlVisible('REPRESENTANT',false);
    SetControlVisible('TREPRESENTANT',false);
    SetControlVisible('REPRESENTANT_',false);
    SetControlVisible('TREPRESENTANT_',false);
    for iInd:=1 to 3 do SetControlProperty('LIBRETIERS'+IntToStr(iInd),'DATATYPE','GCLIBREFOU'+IntToStr(iInd));
    for iInd:=4 to 9 do
        begin
        SetControlVisible('TLIBRETIERS'+IntToStr(iInd),False);
        SetControlVisible('LIBRETIERS'+IntToStr(iInd),False);
        end ;
      SetControlVisible('TLIBRETIERSA', False);
      SetControlVisible('LIBRETIERSA', False);
      //*** Blocages des natures de pièces autorisées en affaires
      if (ctxAffaire in V_PGI.PGIContexte)  then
         begin
         SetControlProperty ( 'NATUREPIECEG', 'Plus','AND (GPP_VENTEACHAT="ACH") '+AfPlusNatureAchat ) ;
         SetControlText ( 'NATUREPIECEG', 'FF' ) ;
         end;
      // **** Fin affaire ***
    end;

{Libres articles}
for iInd := 1 to 9 do
    begin
    GCTitreZoneLibre ('GL_LIBREART' + IntToStr (iInd), stTitre);
    SetControlCaption ('TLIBREART' + IntToStr (iInd), stTitre);
    if (stTitre='') then begin
       SetControlVisible('TLIBREART'+ IntToStr (iInd),False);
       SetControlVisible('LIBREART'+ IntToStr (iInd),False);
       end;
    end;
GCTitreZoneLibre ('GL_LIBREARTA', stTitre);
SetControlCaption ('TLIBREARTA', stTitre);
if (stTitre='') then begin
   SetControlVisible('TLIBREARTA',False);
   SetControlVisible('LIBREARTA',False);
   end;
{$IFDEF CCS3}
for iInd:=5 to 10 do
   BEGIN
   if iInd<10 then Ch:=IntToStr(iInd) else Ch:='A' ;
   SetControlVisible('LIBREART'+Ch,False) ;
   SetControlVisible('TLIBREART'+Ch,False) ;
   END ;
{$ENDIF}
for iInd := 1 to 3 do
    begin
    SetControlCaption ('TFAMILLENIV' + IntToStr (iInd), RechDom('GCLIBFAMILLE','LF'+ IntToStr (iInd),false));
    end;

{Libres tiers}
if stArgument = 'VEN' then
    begin
    for iInd := 1 to 9 do
        begin
        GCTitreZoneLibre ('GP_LIBRETIERS' + IntToStr (iInd), stTitre);
        SetControlCaption ('TLIBRETIERS' + IntToStr (iInd), stTitre);
      if (stTitre='') then begin
         SetControlVisible('TLIBRETIERS'+ IntToStr (iInd),False);
         SetControlVisible('LIBRETIERS'+ IntToStr (iInd),False);
         end;
        end;
    GCTitreZoneLibre ('GP_LIBRETIERSA', stTitre);
    SetControlCaption ('TLIBRETIERSA', stTitre);
      if (stTitre='') then begin
         SetControlVisible('TLIBRETIERSA',False);
         SetControlVisible('LIBRETIERSA',False);
         end;
    end else
    begin
    for iInd := 1 to 3 do
        begin
        GCTitreZoneLibre ('YTC_TABLELIBREFOU' + IntToStr (iInd), stTitre);
        SetControlCaption ('TLIBRETIERS' + IntToStr (iInd), stTitre);
       if (stTitre='') then begin
         SetControlVisible('TLIBRETIERS'+ IntToStr (iInd),False);
         SetControlVisible('LIBRETIERS'+ IntToStr (iInd),False);
         end;
        end;
    end;
{$IFDEF CCS3}
for iInd:=5 to 10 do
   BEGIN
   if iInd<10 then Ch:=IntToStr(iInd) else Ch:='A' ;
   SetControlVisible('LIBRETIERS'+Ch,False) ;
   SetControlVisible('TLIBRETIERS'+Ch,False) ;
   END ;
{$ENDIF}

{Libres articles}
for iInd := 1 to 9 do
    BEGIN
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
    END ;
{$IFDEF CCS3}
SetControlVisible('LIBREARTA',False) ;
SetControlVisible('TLIBREARTA',False) ;
{$ELSE}
GCTitreZoneLibre ('GL_LIBREARTA', stTitre);
SetControlCaption ('TLIBREARTA', stTitre);
{$ENDIF}

if ctxMode in V_PGI.PGIContexte then
  begin
   SetControlVisible('DEPOT',false);
   SetControlVisible('DEPOTMODE',True);
  end else
  begin
  if (ctxAffaire in V_PGI.PGIContexte)  then
         begin
         SetControlVisible('DEPOTMODE',False);
         SetControlVisible('DEPOT',False);
         SetControlVisible('TDEPOT',False);
         end
  else   begin
   SetControlVisible('DEPOT',True);
   SetControlVisible('DEPOTMODE',false);
   end;
  end;
end;

procedure Tof_EditStat.OnClose;
begin
Inherited;
ExecuteSql ('DELETE FROM GCTMPSTA WHERE GZB_UTILISATEUR="' + V_PGI.user + '"');
end;

procedure Tof_EditStat.OnNew;
var iIndex : integer;
begin
Inherited;
for iIndex := 1 to 6 do
    begin
    if GetControlText('REGROUP' + IntToStr (iIndex)) <> '' then
        begin
        SetControlEnabled ('REGROUP' + IntToStr (iIndex), True);
        SetControlProperty ('REGROUP' + IntToStr(iIndex), 'Color',
                            THLabel(GetControl ('NATUREPIECEG')).Color);
        if iIndex < 6 then
            begin
            SetControlEnabled ('REGROUP' + IntToStr (iIndex + 1), True);
            SetControlProperty ('REGROUP' + IntToStr(iIndex + 1), 'Color',
                                THLabel(GetControl ('NATUREPIECEG')).Color);
            end;
        SetControlEnabled ('SAUTPAGE' + IntToStr (iIndex), True);
        end;
    end;
end;

procedure Tof_EditStat.OnUpdate;
var stWhere, stSelect : string;
begin
Inherited;
ExecuteSql ('DELETE FROM GCTMPSTA WHERE GZB_UTILISATEUR="' + V_PGI.user + '"');
if GetControlText ('NATUREPIECEG') = '' then
    begin
    LastError:=1;
    LastErrorMsg:=TexteMessage[LastError] ;
    exit;
    end;
TobTmpSta := Tob.Create ('', nil, -1);
iCompteur := 1;
InitialiseValeur;
SelectionLignes ('1', stSelect, stWhere);
ChargeLignes ('1', stSelect, stWhere);
if (GetControlText ('DATE1PERIODE') <> GetControlText ('DATE2PERIODE')) or
   (GetControlText ('DATE1PERIODE_') <> GetControlText ('DATE2PERIODE_')) then
    begin
    SelectionLignes ('2', stSelect, stWhere);
    ChargeLignes ('2', stSelect, stWhere);
    end;

if TRadioButton(GetControl ('RPRIXACHAT')).checked = True then
    SetControlText ('XX_VARIABLE7', VH_GC.GCMargeArticle)
else SetControlText ('XX_VARIABLE7', 'Quantité');
TobTmpSta.InsertDB(Nil);
TobTmpSta.Free;
end;

// ****************************** Ajout **************************************************

procedure Tof_EditStat.ChargeLignes (stPeriode, stSelect, stWhere : string);
var TSQl : TQuery;
    TobL : Tob;
begin
if stWhere<>'' then stWhere:=stWhere+' AND GL_ARTICLE<>"" AND GL_TYPELIGNE="ART" '
               else stWhere:='GL_ARTICLE<>"" AND GL_TYPELIGNE="ART" ' ;
TSql := OpenSql (stSelect + ' WHERE ' + stWhere, True);

Tobl := Tob.Create ('', nil, -1);
TSql.First;
while not TSql.Eof do
    begin
    if Trim (TSql.findfield ('GL_CODEARTICLE').AsString) <> '' then
        begin
        TobL.SelectDb ('', TSql);
        RegroupLignes (Tobl, stPeriode);
        end;
    TSql.Next;
    end;
Tobl.Free;
Ferme (TSql);
end;

procedure Tof_EditStat.InitialiseValeur;
var iSource, iCible : integer;
    stLibreTitre : string;
begin
for iCible := 1 to nbregroup do
    begin
    stChampRup[iCible] := '';
    SetControlText ('XX_VARIABLE'+IntToStr (iCible), '');
    end;

iCible := nbregroup;
for iSource := nbregroup downto 1 do
    begin
    if GetControlText ('REGROUP' + IntToStr (iSource)) <> '' then
        begin
        if copy (GetControlText ('REGROUP' + IntToStr (iSource)), 1, 2) = 'LF' then
            begin
            SetControlText ('XX_VARIABLE' + IntToStr (iCible),
                            RechDom ('GCLIBFAMILLE',
                                     GetControlText ('REGROUP' + IntToStr (iSource)), False));
            end else if (copy (GetControlText ('REGROUP' + IntToStr (iSource)), 1, 2) = 'LT') then
                begin
                GCTitreZoneLibre ('GP_LIBRETIERS' +
                                    copy (GetControlText ('REGROUP' + IntToStr(iSource)), 3, 1),
                                  stLibreTitre);
                SetControlText ('XX_VARIABLE' + IntToStr (iCible), stLibreTitre);
                end else
                begin
                SetControlText ('XX_VARIABLE' + IntToStr (iCible),
                                RechDom ('GCGROUPSTA',
                                         GetControlText ('REGROUP'+IntToStr (iSource)), False));
                end;
        stChampRup[iCible] := RechDom ('GCGROUPSTA',
                                       GetControlText ('REGROUP' + IntToStr (iSource)),
                                       True);
        if TCheckBox(GetControl ('SAUTPAGE' + IntTostr(iSource))).Checked = False then
            begin
            THEdit(GetControl ('VARSP' + IntToStr(iCible))).Text := '-';
            end else
            begin
            THedit(GetControl ('VARSP' + IntTostr(iCible))).Text := 'X';
            end;
        iCible := iCible - 1;
        end;
    end;
end;

procedure Tof_EditStat.RegroupLignes (TLigne : TOB; stPeriode : string);
var iIndex{, iInd} : integer;
    TobL : TOB;
    dPrixAchat, dQte, dTotalHT, dTotalTTC : double;
    stValRup : array [1..nbregroup] of variant;
    stLibRup : array [1..nbregroup] of string;
    TSql : TQuery;
    {stLibDim,}codedim : string;
begin
for iIndex := 1 to nbregroup do
    begin
    if stChampRup[iIndex] <> '' then
        begin
        if stChampRup[iIndex][2] = 'A' then
            begin
            stValRup[iIndex] := trim(TLigne.GetValue ('GL_'+ Copy(stChampRup[iIndex], 4, length (stChampRup[iIndex]) - 3)));
            if string(stValRup[iIndex]) <> '' then
                    stLibRup[iIndex] := RechDom ('GC' + Copy (stChampRup[iIndex], 4, length (stChampRup[iIndex]) - 3),
                                                 stValRup[iIndex], False);
            end else
            begin
            if stChampRup[iIndex] = 'GL_ARTICLE' then stValRup[iIndex] := TLigne.GetValue ('GL_CODEARTICLE')
            else stValRup[iIndex] := trim(TLigne.GetValue (stChampRup[iIndex]));
            if string(stValRup[iIndex]) <> '' then
                begin
                if copy (stChampRup[iIndex], 4, 5) = 'TIERS' then
                    stLibRup[iIndex] := TLigne.GetValue ('T_LIBELLE')
                else if stChampRup[iIndex] = 'GL_FOURNISSEUR' then
                    stLibRup[iIndex] := RechDom ('GCTIERSFOURN', stValRup[iIndex], False)
                else if stChampRup[iIndex] = 'GL_TARIFTIERS' then
                    stLibRup[iIndex] := RechDom ('TTTARIFCLIENT', stValRup[iIndex], False)
                else if stChampRup[iIndex] = 'GL_ETABLISSEMENT' then
                    stLibRup[iIndex] := RechDom ('TTETABLISSEMENT', stValRup[iIndex], False)

                else if stChampRup[iIndex] = 'GL_ARTICLE' then
                    begin
                    codedim:= trim(copy(TLigne.GetValue ('GL_ARTICLE'),19,15));
                    if codedim='' then
                       BEGIN
                       stLibRup[iIndex] := TLigne.GetValue ('GL_LIBELLE');
                       END else
                       BEGIN
                       TSql := OpenSql ('SELECT GA_LIBELLE,GA_LIBCOMPL,GA_CODEDIM1, GA_CODEDIM2, GA_CODEDIM3, ' +
                                        'GA_CODEDIM4, GA_CODEDIM5, GA_GRILLEDIM4, ' +
                                        'GA_GRILLEDIM1, GA_GRILLEDIM2, GA_GRILLEDIM3, GA_GRILLEDIM5 ' +
                                        'FROM ARTICLE WHERE GA_ARTICLE="' + TLigne.GetValue ('GL_ARTICLE') + '" AND ' +
                                        'GA_STATUTART="DIM"',
                                        True);
                       if not TSql.Eof then
                          begin
                          stLibRup[iIndex] := TSql.FindField ('GA_LIBELLE').AsString + ' ' + TSql.FindField ('GA_LIBCOMPL').AsString ;
                          {for iInd := 1 to MaxDimension do
                            begin
                            stLibDim := GCGetCodeDim (TSql.FindField ('GA_GRILLEDIM' + IntToStr (iInd)).AsString,
                                                      TSql.FindField ('GA_CODEDIM' + IntToStr (iInd)).AsString,
                                                      iInd);
                            if stLibDim <> '' then
                                begin
                                if (iInd > 1) and (stLibRup[iIndex] <> '') then stLibRup[iIndex] := stLibRup[iIndex] + ' - ';
                                stLibRup[iIndex] := stLibRup[iIndex] + stLibDim;
                                end;
                            end; }
                          end;
                       Ferme (TSql);
                       END;
                    end else stLibRup[iIndex] := Rechdom ('GC' + copy (stChampRup[iIndex], 4, length (stChampRup[iIndex]) - 3),
                                                          stValRup[iIndex], False);
                end else stLibRup[iIndex] := '';
            end;
        end else
        begin
        stValRup[iIndex] := '';
        stLibRup[iIndex] := '';
        end;
    end;
TobL := TobTmpSta.FindFirst (['GZB_VALRUPT1', 'GZB_VALRUPT2', 'GZB_VALRUPT3', 'GZB_VALRUPT4',
                              'GZB_VALRUPT5', 'GZB_VALRUPT6'], stValRup, False);
if TobL = Nil then
    begin
    TobL := Tob.Create ('GCTMPSTA', TobTmpSta, -1);
    TobL.InitValeurs;
    TobL.PutValue ('GZB_UTILISATEUR', V_PGI.User);
    for iIndex := 1 to nbregroup do
        begin
        TobL.PutValue ('GZB_VALRUPT' + IntToStr (iIndex), stValRup[iIndex]);
        TobL.PutValue ('GZB_VALRUPT' + IntToStr (iIndex),Copy(TobL.GetValue('GZB_VALRUPT'+IntToStr(iIndex)),1,50));
        Tobl.PutValue ('GZB_LIBRUPT' + IntToStr (iIndex), Copy(stLibRup[iIndex],1,50));
        end;
    TobL.PutValue ('GZB_CAMARGE' + stPeriode, 0.0);
    TobL.PutValue ('GZB_QTECOMPOSANT' + stPeriode, 0.0);
    TobL.PutValue ('GZB_COMPTEUR', iCompteur);
    iCompteur := iCompteur + 1;
    end;

if VH_GC.GCMargeArticle[1] = 'P' then dPrixAchat := TLigne.GetValue ('GL_' + VH_GC.GCMargeArticle + 'P')
else dPrixAchat := TLigne.GetValue ('GL_' + VH_GC.GCMargeArticle);
// BBI : dQte := TLigne.GetValue ('GL_QTEFACT') - TLigne.GetValue ('GL_QTERESTE');
dQte := TLigne.GetValue('GL_QTEFACT');
if TLigne.GetValue ('GL_QTEFACT') <> 0 then
    BEGIN
    dTotalHT := TLigne.GetValue ('GL_TOTALHT') * (dQte/TLigne.GetValue ('GL_QTEFACT'));
    dTotalTTC := TLigne.GetValue ('GL_TOTALTTC') * (dQte/TLigne.GetValue ('GL_QTEFACT'));
    dPrixAchat := dPrixAchat * dQte;    // ajout JCF

    TobL.PutValue ('GZB_CAGLOBAL' + stPeriode,
                   TobL.GetValue ('GZB_CAGLOBAL' + stPeriode) + dTotalHt);
    TobL.PutValue ('GZB_CAGLOBALTTC' + stPeriode,
                   TobL.GetValue ('GZB_CAGLOBALTTC' + stPeriode) + dTotalTTC);

    if TRadioButton(GetControl ('RPRIXACHAT')).checked = True then
        TobL.PutValue ('GZB_COUTQTE' + stPeriode, TobL.GetValue ('GZB_COUTQTE' + stPeriode) + dPrixAchat)
    else TobL.PutValue ('GZB_COUTQTE' + stPeriode,
                        TobL.GetValue ('GZB_COUTQTE' + stPeriode) + dQte);
    if dPrixAchat <> 0.0 then
        begin
        TobL.PutValue ('GZB_CAMARGE' + stPeriode,
                       TobL.GetValue ('GZB_CAMARGE' + stPeriode) + dTotalHt);
        TobL.PutValue ('GZB_MARGE' + stPeriode,
                       TobL.GetValue ('GZB_MARGE' + stPeriode) + dTotalHt - dPrixAchat);
        end;
    END;
(* else
    BEGIN
    dTotalHT := 0;
    dTotalTTC := 0;
    dPrixAchat := 0;
    END;

TobL.PutValue ('GZB_CAGLOBAL' + stPeriode,
               TobL.GetValue ('GZB_CAGLOBAL' + stPeriode) + dTotalHt);
TobL.PutValue ('GZB_CAGLOBALTTC' + stPeriode,
               TobL.GetValue ('GZB_CAGLOBALTTC' + stPeriode) + dTotalTTC);

if TRadioButton(GetControl ('RPRIXACHAT')).checked = True then
    TobL.PutValue ('GZB_COUTQTE' + stPeriode, TobL.GetValue ('GZB_COUTQTE' + stPeriode) + dPrixAchat)
else TobL.PutValue ('GZB_COUTQTE' + stPeriode,
                    TobL.GetValue ('GZB_COUTQTE' + stPeriode) + dQte);
if dPrixAchat <> 0.0 then
    begin
    TobL.PutValue ('GZB_CAMARGE' + stPeriode,
                   TobL.GetValue ('GZB_CAMARGE' + stPeriode) + dTotalHt);
    TobL.PutValue ('GZB_MARGE' + stPeriode,
                   TobL.GetValue ('GZB_MARGE' + stPeriode) + dTotalHt - dPrixAchat);
    end; *)
end;

procedure Tof_EditStat.SelectionLignes (stPeriode : string;
                                        var stSelect, stWhere : string);
var stdepot,sttmp,stStatutArt, stStatut, stWhereStatut,stChampsTiers,stTable,stWhTiers : string;
    iInd,ILT : integer;
    YaTiers : boolean;
begin
stSelect := 'SELECT GL_CODEARTICLE, GL_ARTICLE, GL_LIBELLE,GL_QTEFACT, GL_TOTALHT, GL_TOTALTTC, GL_QTERESTE, ' +
            'GL_DPA, GL_PMAP, GL_PMRP, GL_DPR, GL_DEPOT, GL_ETABLISSEMENT, GL_FOURNISSEUR, GL_FAMILLENIV1, ' +
            'GL_FAMILLENIV2, GL_FAMILLENIV3, GL_REPRESENTANT, GL_TARIFARTICLE, ' +
            'GL_TIERSFACTURE, GL_TIERS, GL_TIERSLIVRE, GL_TIERSPAYEUR, GL_TARIFTIERS, ' +
            'GL_TYPEARTICLE, GL_LIBREART1, GL_LIBREART2, GL_LIBREART3, GL_LIBREART4, ' +
            'GL_LIBREART5, GL_LIBREART6, GL_LIBREART7, GL_LIBREART8, GL_LIBREART9, ' +
            'GL_LIBREARTA,GL_MTRESTE ';
stWhere := 'GL_INDICEG=0 AND GL_TYPEARTICLE <> "FI" AND ' +
           'GL_DATEPIECE>="' + USDateTime (StrToDate (GetControlText ('DATE' + stPeriode + 'PERIODE'))) + '" AND ' +
           'GL_DATEPIECE<="' + USDateTime (StrToDate (GetControlText ('DATE' + stPeriode + 'PERIODE_'))) + '"';

if GetControlText ('NATUREPIECEG') <> '' then
begin
 	if not(ctxAffaire in V_PGI.PGIContexte)  then
  Begin
    stWhere := stWhere + ' AND ' + FactureAvoir_RecupSQL(GetControlText('NATUREPIECEG'),'GL');
 end
	else
  Begin
  // gm 05/09/02  Prise en compte Facture reprise
    if GetControlText('NATUREPIECEG')='ZZ1' then
      stWhere := stWhere +  ' AND (GL_NATUREPIECEG="FAC" OR GL_NATUREPIECEG="AVC" OR GL_NATUREPIECEG="AVS" OR GL_NATUREPIECEG="FRE")'
    else
      if GetControlText('NATUREPIECEG')='ZZ2' then
        stWhere := stWhere +  ' AND (GL_NATUREPIECEG="FAC" OR GL_NATUREPIECEG="FRE")'
      else
        stWhere := stWhere +  ' AND GL_NATUREPIECEG="' + GetControlText ('NATUREPIECEG') + '"';
	ENd;
ENd;

if GetControlText ('ARTICLE') <> '' then
    begin
    stWhere := stWhere + ' AND GL_ARTICLE>="' + GetControlText ('ARTICLE') + '"';
    end;
if GetControlText ('ARTICLE_') <> '' then
    begin
    stWhere := stWhere + ' AND GL_ARTICLE<="' + format('%-18.18s%16.16s',[GetControlText ('ARTICLE_'),StringOfChar('Z',16)]) + '"';
    end;
if GetControlText ('TIERS') <> '' then
    begin
    stWhere := stWhere + ' AND GL_TIERS>="' + GetControlText ('TIERS') + '"';
    end;
if GetControlText ('TIERS_') <> '' then
    begin
    stWhere := stWhere + ' AND GL_TIERS<="' + GetControlText ('TIERS_') + '"';
    end;
if GetControlText ('REPRESENTANT') <> '' then
    begin
    stWhere := stWhere + ' AND GL_REPRESENTANT>="' + GetControlText ('REPRESENTANT') + '"';
    end;
if GetControlText ('REPRESENTANT_') <> '' then
    begin
    stWhere := stWhere + ' AND GL_REPRESENTANT<="' + GetControlText ('REPRESENTANT_') + '"';
    end;
if ctxMode in V_PGI.PGIContexte then
    BEGIN
    stdepot := GetControlText ('DEPOTMODE');
    if (stdepot<>'') and (stdepot<>'<<Tous>>') then
        begin
        iInd := 0;
        repeat
          sttmp := ReadTokenSt (stdepot);
          if iInd = 0 then stWhere := stWhere + ' AND (GL_DEPOT="'+sttmp+'"'
          else stWhere := stWhere + ' OR GL_DEPOT="'+sttmp+'"';
          inc (iInd);
        until stdepot = '';
        stWhere := stWhere + ')';
        end;
    END ELSE
    BEGIN
    if GetControlText ('DEPOT') <> '' then
        begin
        stWhere := stWhere + ' AND GL_DEPOT="' + GetControlText ('DEPOT') + '"';
        end;
    END;

if GetControlText('FAMILLENIV1')<>'' then stWhere:=stWhere+' AND GL_FAMILLENIV1="'+GetControlText('FAMILLENIV1')+'"';
if GetControlText('FAMILLENIV2')<>'' then stWhere:=stWhere+' AND GL_FAMILLENIV2="'+GetControlText('FAMILLENIV2')+'"';
if GetControlText('FAMILLENIV3')<>'' then stWhere:=stWhere+' AND GL_FAMILLENIV3="'+GetControlText('FAMILLENIV3')+'"';

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
stStatutArt := GetControlText('STATUTART');
if (pos ('GEN', stStatutArt)  <> 0) or (pos ('UNI', stStatutart) <> 0) or
   (pos ('DIM', stStatutArt) <> 0) then
    begin
//    stSelect := stSelect + ', ARTICLE ';  OT : 04/06/2002
    stTable := ',ARTICLE' ;
    stWhereStatut := '';
    repeat
        stStatut := ReadTokenSt (stStatutArt);
        if stWhereStatut <> '' then stWhereStatut := stWhereStatut + ' OR ';
        stWhereStatut := stWhereStatut + 'GA_STATUTART="' + stStatut + '"';
    until stStatutArt = '';
    stWhere := stWhere + ' AND GA_ARTICLE=GL_ARTICLE AND (' + stWhereStatut + ')';
    end;

ILT:=nbregroup+1; YaTiers:=False;
for iInd := 1 to nbregroup do
    begin
    if (stChampRup[iInd] <> '') and (copy (stChampRup[iInd], 4, 5) = 'TIERS') then  YaTiers:=True;
    if copy (GetControlText ('REGROUP' + IntToStr (iInd)), 1, 2) = 'LT' then begin ILT:=iInd; continue; end;
    end;
iInd:=ILT;
if YaTiers then
   begin
   stChampsTiers:=',T_LIBELLE';
   stTable := stTable + ',TIERS ';
   stWhTiers:='T_TIERS=GL_TIERS ';
   end else
   begin
   stChampsTiers:='';
//   stTable:='';   OT : 04/06/2002
   end;
if (pos ('GP_LIBRETIERS', stWhere) <> 0) or (iInd <= nbregroup) then
    begin
    stSelect := stSelect + stChampsTiers + ', GP_LIBRETIERS1, ' +
            'GP_LIBRETIERS2, GP_LIBRETIERS3, GP_LIBRETIERS4, GP_LIBRETIERS5, GP_LIBRETIERS6, ' +
            'GP_LIBRETIERS7, GP_LIBRETIERS8, GP_LIBRETIERS9, GP_LIBRETIERSA ' +
            'FROM LIGNE, PIECE ' + stTable;
    stWhere := stWhere + ' AND GP_NATUREPIECEG=GL_NATUREPIECEG AND ' +
            'GP_SOUCHE=GL_SOUCHE AND GP_NUMERO=GL_NUMERO AND GP_INDICEG=GL_INDICEG ';
    end else
    begin
    stSelect := stSelect + stChampsTiers + ' FROM LIGNE '+ stTable;
    end;
if stWhTiers <> '' then
   begin
   if stWhere <> '' then stWhere := stWhere + ' AND ' + stWhTiers
   else stWhere := stWhere + stWhTiers;
   end;
if stWhere <> '' then stWhere := ' (' + stWhere + ')';
end;

// ****************************** register *****************************************

Procedure Tof_EditStat_AffectRegroup (parms:array of variant; nb: integer ) ;
var F : TForm;
    stPlus, stValue, St_Text, St : string;
    Indice, iInd : integer;
begin
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCSTAT') then exit;
Indice := Integer (Parms[1]);
stPlus := THValComboBox (F.FindComponent ('REGROUP1')).Plus;
stValue := string (THValComboBox (F.FindComponent ('REGROUP'+InttoStr (Indice))).Value);
St_Text := string (THValComboBox (F.FindComponent ('REGROUP'+InttoStr (Indice))).Text);
For iInd := 1 to nbregroup do
    begin
    if iInd = Indice then continue;
    St := string (THValComboBox (F.FindComponent ('REGROUP'+InttoStr (iInd))).Value);
    If St <> '' then stPlus := stPlus + ' AND CO_CODE <>"' + St + '"';
    end;
THValComboBox (F.FindComponent ('REGROUP'+InttoStr (Indice))).Plus := stPlus;
THValComboBox (F.FindComponent ('REGROUP'+InttoStr (Indice))).Value := stValue;
THValComboBox (F.FindComponent ('REGROUP'+InttoStr (Indice))).Text := St_Text ;
end;

Procedure Tof_EditStat_AffectStatutArt (parms:array of variant; nb: integer ) ;
var F : TForm;
    stPlus, stStatut, stStatutArt : string;
begin
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCSTAT') then exit;
stPlus := '';
stStatutArt := string(THMultiValComboBox(F.FindComponent('STATUTART')).Text);
if stStatutArt <> '' then
    begin
    stStatut := '';
    Repeat 
        stStatut := ReadTokenSt (stStatutArt);
        if stPlus <> '' then stPlus := stPlus + ' OR ';
        stPlus := stPlus + 'GA_STATUTART="' + stStatut + '"';
    until stStatutArt = '';
   end;
THEdit(F.FindComponent('ARTICLE')).Plus := stPlus;
THEdit(F.FindComponent('ARTICLE_')).Plus := stPlus;
end;


{***********A.G.L.***********************************************
Auteur  ...... : JS
Créé le ...... : 30/07/2003
Modifié le ... :   /  /    
Description .. : Ajoute l'item Facture+Avoir à un combo de sélection de 
Suite ........ : natures de pièces sous le code ZZA(achat) ou ZZV (vente).
Mots clefs ... : FACTURE;AVOIR
*****************************************************************}
procedure FactureAvoir_AddItem(CC : TControl ; VenteAchat : string);
var CBNaturePiece : THValComboBox;
    Iind : integer;
begin
  if not (CC is THValComboBox) then exit;
  CBNaturePiece := THValComboBox(CC);
  if VenteAchat = 'VEN' then
  begin
    Iind:=CBNaturePiece.Values.IndexOf('FAC');
    if Iind<0 then Iind:=0;
    CBNaturePiece.Items.Insert(Iind, TraduireMemoire('Facture + Avoir clients'));
    CBNaturePiece.Values.Insert(Iind,'ZZV');
  end else if VenteAchat = 'ACH' then
           begin
             Iind:=CBNaturePiece.Values.IndexOf('FF');
             if Iind<0 then Iind:=0;
             CBNaturePiece.Items.Insert(Iind, TraduireMemoire('Facture + Avoir fournisseurs'));
             CBNaturePiece.Values.Insert(Iind,'ZZA');
           end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : JS
Créé le ...... : 30/07/2003
Modifié le ... :   /  /    
Description .. : Associée à la procédure FactureAvoir_AddItem cette 
Suite ........ : fonction renvoie une clause where sur les natures de pièces 
Suite ........ : factures + avoirs clients ou fournisseurs.
Mots clefs ... : FACTURE;AVOIR
*****************************************************************}
function FactureAvoir_RecupSQL(NatureP, PrefixeSQL : string) : string;
var Champ,VenteAchat : string;
    TOBP : TOB;
begin
  Champ := PrefixeSQL + '_NATUREPIECEG';
  if NatureP = 'ZZA' then
  begin
    VenteAchat := 'ACH';
    Result := Champ + '="FF"';
  end else if NatureP = 'ZZV' then
           begin
             VenteAchat := 'VEN';
             Result := Champ + '="FAC"';
           end else begin
                      Result := Champ + '="' + NatureP + '"';
                      exit;
                    end;
  TOBP := VH_GC.TOBParPiece.FindFirst(['GPP_VENTEACHAT','GPP_ESTAVOIR'],[VenteAchat,'X'],False);
  while TOBP <> nil do
  begin
    Result :=  Result + ' OR ' + Champ + '="' + TOBP.GetValue('GPP_NATUREPIECEG') + '"';
    TOBP := VH_GC.TOBParPiece.FindNext(['GPP_VENTEACHAT','GPP_ESTAVOIR'],[VenteAchat,'X'],False);
  end;
  Result := '(' + Result + ')';
end;


procedure InitTOFEditStat ();
begin
RegisterAglProc ('EditStat_AffectRegroup', True , 1, Tof_EditStat_AffectRegroup);
RegisterAglProc ('EditStat_AffectStatutArt', True , 1, Tof_EditStat_AffectStatutArt);
end;

Initialization
registerclasses ([Tof_EditStat]);
InitTOFEditStat ();

end.
