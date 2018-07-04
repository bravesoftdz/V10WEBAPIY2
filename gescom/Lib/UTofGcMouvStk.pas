{***********UNITE*************************************************
Auteur  ...... : B. BICHERAY
Créé le ...... : 24/01/2001
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : GCMOUVSTK ()
Mots clefs ... : TOF;GCMOUVSTK
*****************************************************************}
Unit UTOFGCMOUVSTK ;

Interface

Uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls,
     Windows, HCtrls, HEnt1, HMsgBox, UTOB, UTOF,
     Messages, Dialogs,HSysMenu, Hqry, HTB97,
     ExtCtrls, HPanel, Grids,   UIUtil, Mask,HXlspas,
     Math, SaisUtil, LookUp, UTofConsultStock,
{$IFDEF EAGLCLIENT}
     eMul,eFichList,Maineagl,UtileAgl,
{$ELSE}
     Mul,FichList,FE_Main,DBCtrls,DBGrids,HDB,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}EdtREtat,
{$ENDIF}
     Menus, HRichOLE,HRichEdt, ColMemo, EntGC, AglInitGC, TarifUtil, Facture, FactComm, FactUtil,
     UTofOptionEdit, StockUtil, HFLabel, InvUtil,M3FP, UtilGC, UtilArticle, UtilDispGC,LicUtil,uEntCommun,UtilTOBPiece;

Var PEtat : RPARETAT;

Const
     MOU_Depot          : integer = 0;
     MOU_date           : integer = 0;
     MOU_QualifMvt      : integer = 0;
     MOU_QteStock       : integer = 0;
     NbRowsInit = 50;
     NbRowsPlus = 20;

Type
  TOF_GCMOUVSTK = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure BAgrandirClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure FormClose (Sender: TObject; var Action: TCloseAction);
    procedure G_ConsultMouvColumnWidthsChanged(Sender: TObject);
    procedure G_ConsultMouvDblClick(Sender: TObject);
    procedure MyExportClick(Sender: TObject);
  private
    { Déclarations privées }
    LesColConsultMouv : String;
    LesChampsConsultMouv : String;
    DEV : RDEVISE;
    TobLigne, TobLigneAff : TOB;
    iTableLigne : integer;
    stWhereNat,stWhereNatNom,stWhereArt : string;
    bFormCreate : boolean;
{$IFDEF EAGLCLIENT}
    FListe : THGrid;
{$ELSE}
    FListe : THDBGrid;
{$ENDIF}
    HMTrad : THSystemMenu;
    G_ConsultMouv : THGrid;
    PSAICUMUL : THPanel;
    CUMQTESTO : THNumEdit;
    BAgrandir, BReduire, BOuvrir, BImprimer : TToolBarButton97;
    procedure RecupControl;
    procedure CreeTotal (ColLig : integer; Cumul : THNumEdit);
    procedure MiseEnForme;
    function  PrepareImpression : integer ;
    procedure Imprime ;
//  public
    { Déclarations publiques }
// Actions liées au Grid
    procedure EtudieColsListe;
// Gestion des données
    procedure AjoutSelect (var stSelect : string; stChamp : string);
    procedure ChargeMouvement;
    function CalcPrixRatio (dPrix, dUniteStock, dUnite, dPrixPourQte : Double) : Double;
    Function CalcRatioMesure (stCat, stMesure : String) : Double;
    procedure CopieLaLigne (index : integer);
    procedure DecodeRefLignePiece (stReferenceLignePiece : String; Var CleDoc : R_CleDoc );
    procedure DetermineNatureAutorisee;
    function Evaluedate (St : String) : TDateTime;
    procedure LoadLesTobLigne;
    procedure LoadLesTobNomen;
    procedure TraiteLesLignes(stWhereNomen : string);
    procedure TraiteLigneNomen(TobLigne : TOB; iInd : integer; stWhereNomen : string);
//    function TypePieceStockPhy (stNaturePiece : string) : boolean;
    Procedure VoirDispo ;
    Procedure VoirArticle ;
  end;

const
// libellés des messages
TexteMessage: array[1..2] of string 	= (
          {1}  'Vous devez saisir un code article'
          {2} ,''
              );

Implementation

procedure TOF_GCMOUVSTK.OnUpdate ;
begin
  Inherited ;
MiseEnForme;
EtudieColsListe;
AffecteGrid (G_ConsultMouv, taConsult);
if (GetControlText('_ARTICLE') = '') and (not bFormCreate) then
   begin
   //MessageAlerte('Aucun élément sélectionné');
   HShowMessage('0;'+Ecran.Caption+';'+TexteMessage[1]+';W;O;O;O;','','') ;
   exit;
   end ;
bFormCreate := false;
DetermineNatureAutorisee;
ChargeMouvement;
//SetControlText ('XX_WHERE', '');
end ;

procedure TOF_GCMOUVSTK.OnLoad ;
var Ctl : TControl ;
begin
  Inherited ;
// Libellé Dépôt ou Etablissement
if not VH_GC.GCMultiDepots then
   begin
    Ctl := GetControl('TGL_DEPOT') ;
    if (Ctl <> Nil) and (Ctl is THLabel) then THLabel(Ctl).Caption := 'Etablissement' ;
   end ;
if (ctxMode in V_PGI.PGIContexte) then
   SetControlProperty('GQ_DEPOT','Plus','GDE_SURSITE="X"');
end ;

procedure TOF_GCMOUVSTK.OnArgument (S : String ) ;
var iCol : integer ;
    stIndice : string ;
begin
// Paramétrage des libellés des familles, collection, stat. article et dimensions
SetControlCaption('TGL_COLLECTION',RechDom('GCZONELIBRE','AS0',False));
for iCol:=1 to 3 do
    begin
    stIndice:=IntToStr(iCol) ;
    ChangeLibre2('TGL_FAMILLENIV'+stIndice,Ecran);
    end;
if (ctxMode in V_PGI.PGIContexte) and (GetPresentation=ART_ORLI) then
   begin
   for iCol:=4 to 8 do
       begin
       stIndice:=IntToStr(iCol) ;
       ChangeLibre2('TGA2_FAMILLENIV'+stIndice,Ecran);
       end;
   for iCol:=1 to 2 do
       begin
       stIndice:=IntToStr(iCol) ;
       ChangeLibre2('TGA2_STATART'+stIndice,Ecran);
       end;
   end ;
// Paramètrage du libellé de l'onglet des tables libres du dépôt
if not VH_GC.GCMultiDepots then
   begin
   SetControlCaption('PTABLESLIBRESDEP','Tables Libres Etab.') ;
   end;

// Paramétrage des libellés des tables libres
if (GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GL_LIBREART', 10, '') = 0) then SetControlVisible('PTABLESLIBRES', False) ;
if (GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GDE_LIBREDEP', 10, '') = 0) then SetControlVisible('PTABLESLIBRESDEP', False) ;

bFormCreate := true;
iTableLigne := PrefixeToNum ('GL');
TobLigne := TOB.Create ('', Nil, -1);
TobLigneAff := TOB.Create ('', Nil, -1);
DEV.Code := V_PGI.DevisePivot;
GetInfosDevise (DEV);
PEtat.Tip:='E'; PEtat.Nat:=TableToPrefixe('GCTMPMOUV');
PEtat.Modele:=TableToPrefixe('GCTMPMOUV'); PEtat.Titre:=TFmul(Ecran).Caption;
PEtat.Apercu:=True; PEtat.DeuxPages:=False; PEtat.First:=True;
PEtat.stSQL:='';

if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
    begin
    TToolbarButton97(GetControl('BParamListe')).Visible:=True ;
    end else
    begin
    TToolbarButton97(GetControl('BParamListe')).Visible:=False;
    end;

  Inherited ;
SetControlVisible('Fliste', False);
SetControlText('_COLSTOCK', 'PHY');
RecupControl;
MajChampsLibresArticle(TForm(Ecran),'GL');
TFMul(Ecran).bExport.OnClick:=MyExportClick;
TFMul(Ecran).bExport.Visible:=True;
end ;

procedure TOF_GCMOUVSTK.OnClose ;
begin
G_ConsultMouv.VidePile(True);
TobLigne.Free;
TobLigne := Nil;
TobLigneAff.Free;
TobLigneAff := Nil;
  Inherited ;
end ;

procedure TOF_GCMOUVSTK.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
G_ConsultMouv.VidePile(True);
TobLigne.Free;
TobLigne := Nil;
TobLigneAff.Free;
TobLigneAff := Nil;
  inherited;
end;

{========================================================================================}
{========================= Actions liées au Grid ========================================}
{========================================================================================}
procedure TOF_GCMOUVSTK.EtudieColsListe;
Var NomCol, LesCols,ColChSupp : String;
    Col, ichamp, Larg,Dec,iPos : Integer;
    Perso, NomList,Typ,StA,St,FF, STitre  : string ;
    FRecordSource,FLien,FFieldList,FSortBy,FLargeur,FAlignement,FParams : string ;
    OkTri,OkNumCol,Obli, OkLib, Sep, OkVisu,OkNulle,OkCumul, ChSupp : boolean ;
    FTitre,tt,NC : Hstring;
    tal : TAlignment ;
begin
LesCols := ''; LesColConsultMouv := ''; LesChampsConsultMouv:='';
G_ConsultMouv.ColCount:=60;
Col:=1; STitre:='';
With TFMul(Ecran) do
    BEGIN
    NomList:=G_ConsultMouv.ListeParam;
    ChargeHListe(NomList,FRecordSource,FLien,FSortBy,FFieldList,FTitre,FLargeur,FAlignement,FParams,tt,NC,Perso,OkTri,OkNumCol);
    G_ConsultMouv.Titres.Add(FFieldList) ;
    While Ftitre<> '' do
        BEGIN
        StA:=ReadTokenSt(FAlignement);
        St:=ReadTokenSt(Ftitre);
        Nomcol:=ReadTokenSt(FFieldList); ChSupp:=False;
        if NomCol<>'' Then
            begin
            ichamp := ChampToNum(NomCol);
            if ichamp>=0 Then
                begin
                if Nomcol='GL_DEPOT' then MOU_Depot := col else
                if NomCol='GL_DATEPIECE' Then MOU_Date := col else
                if NomCol='GL_QUALIFMVT' Then MOU_QualifMvt := col else
                if NomCol='GL_QTESTOCK' Then MOU_QteStock := col;
                end else
                begin
                iPos:=Pos('AS', UpperCase(NomCol));
                if ipos>0 then ColChSupp:=Trim(Copy(NomCol,iPos+3,18))
                          else ColChSupp:='COLUMN'+IntToStr(Col);
                ChSupp:=True;
                end;
            end;
        Larg:=ReadTokenI(FLargeur);
        tal:=TransAlign(StA,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
        if OkVisu then
            BEGIN
            G_ConsultMouv.Cells[Col,0]:=St ;
            G_ConsultMouv.ColAligns[Col]:=tal;
            G_ConsultMouv.ColWidths[Col]:=Larg;//*G_ConsultMouv.Canvas.TextWidth('W') ;
            Typ:=ChampToType(Nomcol) ;
            if (Typ='COMBO') then
               if OkLib then G_ConsultMouv.ColFormats[Col]:='CB=' + Get_Join(Nomcol)
               else G_ConsultMouv.ColFormats[Col]:='';
            if (Typ='INTEGER') or (Typ='SMALLINT') or (Typ='DOUBLE') then G_ConsultMouv.ColFormats[Col]:=FF ;
            if STitre = '' then STitre := St
            else STitre := STitre+';'+St;
            if ChSupp then
               begin
               if LesColConsultMouv = '' then LesColConsultMouv := ColChSupp
               else LesColConsultMouv := LesColConsultMouv+';'+ColChSupp;
               end else
               begin
               if LesColConsultMouv = '' then LesColConsultMouv := Nomcol
               else LesColConsultMouv := LesColConsultMouv+';'+Nomcol;
               end;
            if LesChampsConsultMouv = '' then LesChampsConsultMouv := Nomcol
            else LesChampsConsultMouv := LesChampsConsultMouv+';'+Nomcol;
            inc (Col);
            END;
        END;

    G_ConsultMouv.ColCount:=Col ;
    if MOU_Date <> 0 then G_ConsultMouv.ColFormats[Mou_Date] := ShortdateFormat ;
    HMTrad.ResizeGridColumns (G_ConsultMouv);
    CreeTotal (MOU_QteStock, CUMQTESTO);
    end;
end;

{========================================================================================}
{================================= Gestion des Données ==================================}
{========================================================================================}
procedure TOF_GCMOUVSTK.AjoutSelect (var stSelect : string; stChamp : string);
begin
if Pos (stChamp, stSelect) = 0 then stSelect := stSelect + ', ' + stChamp;
end;

function TOF_GCMOUVSTK.CalcPrixRatio (dPrix, dUniteStock, dUnite, dPrixPourQte : Double) : Double;
begin
if dPrixPourQte = 0.0 then Result := dPrix
else result := (dPrix * dUniteStock)/(dPrixPourQte * dUnite);
end;

Function TOF_GCMOUVSTK.CalcRatioMesure (stCat, stMesure : String) : Double;
var TobM : TOB;
    dRatio : Double;
begin
TobM := VH_GC.MTOBMEA.FindFirst(['GME_QUALIFMESURE','GME_MESURE'], [stCat, stMesure], False);
dRatio := 0;
if TobM <> nil then dRatio := TobM.GetValue('GME_QUOTITE');
if dRatio = 0 then dRatio := 1.0;
result := dRatio;
end;

procedure TOF_GCMOUVSTK.ChargeMouvement;
begin
TobLigne.ClearDetail;
TobLigneAff.ClearDetail;
G_ConsultMouv.VidePile (False);
LoadLesTobLigne;
if TobLigneAff.Detail.Count > 0 then
    begin
    G_ConsultMouv.RowCount := TobLigneAff.Detail.Count + 1;
    TobLigneAff.PutGridDetail (G_ConsultMouv, false, false, LesColConsultMouv, False);
    CUMQTESTO.Value := TobLigneAff.Somme ('GL_QTESTOCK', [''], [''], False);
    if G_ConsultMouv.CanFocus then G_ConsultMouv.SetFocus;
    end
else CUMQTESTO.Value := 0;
if MOU_QualifMvt <> 0 then G_ConsultMouv.Cells [MOU_QualifMvt, 0] := 'Mouvement';
if MOU_QteStock <> 0 then G_ConsultMouv.Cells [MOU_QteStock, 0] := 'Quantité';
end;

procedure TOF_GCMOUVSTK.CopieLaLigne (index : integer);
var Tobl : TOB;
    dCoef : double;
begin
TobL := TobLigneAff.FindFirst (['GL_DATEPIECE', 'GL_NATUREPIECEG', 'GL_NUMERO', 'GL_DEPOT',
                                'GL_QUALIFMVT', 'GL_ARTICLE'],
                               [TobLigne.Detail[index].GetValue('GL_DATEPIECE'),
                                TobLigne.Detail[index].GetValue('GL_NATUREPIECEG'),
                                TobLigne.Detail[index].GetValue('GL_NUMERO'),
                                TobLigne.Detail[index].GetValue('GL_DEPOT'),
                                TobLigne.Detail[index].GetValue('GL_QUALIFMVT'),
                                TobLigne.Detail[index].GetValue('GL_ARTICLE')], True);
if Tobl = Nil then
    begin
    Tobl := Tob.Create ('LIGNE', TobLigneAff, -1);
    Tobl.Dupliquer (TobLigne.Detail[index], true, true);
    TobL.AddChampSup ('RATIO', False);
    TobL.AddChampSup ('VA', False);
    Tobl.PutValue('VA', GetInfoParPiece (TOBL.GetValue('GL_NATUREPIECEG'),'GPP_VENTEACHAT'));
    TobL.PutValue ('RATIO', GetRatio (Tobl, nil, trsStock));
    Tobl.PutValue('GL_QTESTOCK', Tobl.GetValue ('GL_QTESTOCK')/TobL.GetValue('RATIO'));
    if Tobl.GetValue ('VA') = 'ACH' then
        dCoef := CalcRatioMesure('PIE', Tobl.GetValue('GL_QUALIFQTEACH'))
    else dCoef := CalcRatioMesure('PIE', Tobl.GetValue('GL_QUALIFQTEVTE'));
    if pos ('GL_PUHT', LesColConsultMouv) > 0 then
        begin
        Tobl.PutValue('GL_PUHT',
                      CalcPrixRatio(Tobl.GetValue('GL_PUHT'),
                                    CalcRatioMesure('PIE', Tobl.GetValue('GL_QUALIFQTESTO')),
                                    dCoef, TobL.GetValue('GL_PRIXPOURQTE')));
        end;
    end else
    begin
    if TobL.GetValue ('VA') = 'ACH' then
        dCoef := CalcRatioMesure('PIE', TobLigne.Detail[index].GetValue('GL_QUALIFQTEACH'))
    else dCoef := CalcRatioMesure('PIE', TobLigne.Detail[index].GetValue('GL_QUALIFQTEVTE'));
    Tobl.PutValue ('GL_QTESTOCK',
                   TobL.GetValue ('GL_QTESTOCK') +
                       (TobLigne.Detail[index].GetValue('GL_QTESTOCK')/
                            Tobl.GetValue('RATIO')));
    if pos ('GL_PUHT', LesColConsultMouv) > 0 then
        begin
        Tobl.PutValue('GL_PUHT',
                      Tobl.GetValue ('GL_PUHT') +
                      CalcPrixRatio(TobLigne.Detail[index].GetValue('GL_PUHT'),
                                    CalcRatioMesure('PIE',
                                                    TobLigne.Detail[index].GetValue('GL_QUALIFQTESTO')),
                                    dCoef, TobLigne.Detail[index].GetValue('GL_PRIXPOURQTE')));
        end;
    end;
    Tobl.PutValue('GL_QTERESTE', Tobl.GetValue('GL_QTESTOCK')); { NEWPIECE }
end;

Procedure TOF_GCMOUVSTK.DecodeRefLignePiece (stReferenceLignePiece : String; Var CleDoc : R_CleDoc);
Var stCleDoc : String;
begin
FillChar (CleDoc, Sizeof(CleDoc), #0);
stCleDoc := stReferenceLignePiece;
CleDoc.DatePiece := EvalueDate(ReadTokenSt(stCleDoc));
CleDoc.NaturePiece := ReadTokenSt(stCleDoc);
CleDoc.Souche := ReadTokenSt(stCleDoc);
CleDoc.NumeroPiece := StrToInt(ReadTokenSt(stCleDoc));
CleDoc.Indice := StrToInt(ReadTokenSt(stCleDoc));
CleDoc.NumLigne := StrToInt(ReadTokenSt(stCleDoc));
end;

procedure TOF_GCMOUVSTK.DetermineNatureAutorisee ;
var Index : integer;
    QteAffect : string;
begin
stWhereNat:='';  stWhereNatNom:='';
QteAffect:=GetControlText('_COLSTOCK')+';';
for Index := 0 to VH_GC.TobParPiece.Detail.Count - 1 do
    begin
    if ((pos (QteAffect, VH_GC.TobParPiece.Detail[Index].GetValue ('GPP_QTEPLUS')) > 0) or
       (pos (QteAffect, VH_GC.TobParPiece.Detail[Index].GetValue ('GPP_QTEMOINS')) > 0)) and
       ((GetControlText ('SENSPIECE') = '') or
        (VH_GC.TobParPiece.Detail[Index].GetValue ('GPP_SENSPIECE') = GetControlText ('SENSPIECE'))) then
        begin
        if stWhereNat<>'' then stWhereNat:=stWhereNat+') OR ' else stWhereNat:='(';
        if stWhereNatNom<>'' then stWhereNatNom:=stWhereNatNom+') OR ' else stWhereNatNom:='(';
        stWhereNat:=stWhereNat+'(GL_NATUREPIECEG="'+VH_GC.TobParPiece.Detail[Index].GetValue ('GPP_NATUREPIECEG') + '"';
        stWhereNatNom:=stWhereNatNom+'(GLN_NATUREPIECEG="'+VH_GC.TobParPiece.Detail[Index].GetValue ('GPP_NATUREPIECEG') + '"';
        end;
    end;
if stWhereNat<>'' then stWhereNat:=stWhereNat + '))';
if stWhereNatNom<>'' then stWhereNatNom:=stWhereNatNom + '))';
end;

Function TOF_GCMOUVSTK.Evaluedate (St : String) : TDateTime;
Var dd,mm,yy : Word;
begin
Result := 0; if St='' then Exit;
dd := StrToInt(Copy(St,1,2)); mm := StrToInt(Copy(St,3,2)); yy := StrToInt(Copy(St,5,4));
Result := Encodedate(yy,mm,dd);
end;

procedure TOF_GCMOUVSTK.LoadLesTobLigne;
var stWhere, stWhereNomen, stSelect, stSelectCount, stLesCols, stChamp, stArt : string;
    TSql : TQuery;
    i_ind,lg,Nbenreg : integer;
    QteAffect, Perso, NomList : string;
    FRecordSource, FLien, FFieldList,  FSortBy, FLargeur, FAlignement, FParams: string;
    OkTri, OkNumCol : boolean;
    FTitre, tt,Nc : Hstring;
begin
stWhereArt:='';
LoadLesTobNomen;
stWhere := RecupWhereCritere(TPageControl(GetControl('Pages'))) ;
stArt := format('%-18.18s',[GetControlText('_ARTICLE')]);
stWhereNomen := stArt;
if stWhereArt='' then stWhereArt:='GL_ARTICLE Like "' + stArt + '%"'
                 else stWhereArt:='('+stWhereArt+'OR GL_ARTICLE Like "' + stArt +'%")';
if GetControlText('_ARTICLE')='' then
   begin
   stWhereArt:='';
   end else
   begin
   i_ind:=pos('AND ('+GetControlText('XX_WHERE')+')',stWhere);
   lg:=Length('AND ('+GetControlText('XX_WHERE')+')');
   if i_ind=0 then
      begin
      i_ind:=pos('('+GetControlText('XX_WHERE')+')',stWhere);
      lg:=Length('('+GetControlText('XX_WHERE')+')');
      end ;
   if i_ind>0 then stWhere:=Copy(stWhere,1,i_ind-1)+Copy(stWhere,i_ind+lg,Length(stWhere));
   end;

stSelect := '';
//stLesCols := LesColConsultMouv;  LesChampsConsultMouv
stLesCols := LesChampsConsultMouv;
stChamp := ReadTokenSt(stLesCols);
while stChamp <> '' do
    begin
    if stSelect <> '' then stSelect := stSelect + ', ';
    stSelect := stSelect + stChamp;
    stChamp := ReadTokenSt (stLesCols);
    end;
stSelect := 'SELECT ' + stSelect;
AjoutSelect (stSelect, 'GL_QTERELIQUAT');
AjoutSelect (stSelect, 'GL_QTERESTE');
AjoutSelect (stSelect, 'GL_PIECEPRECEDENTE');
AjoutSelect (stSelect, 'GL_INDICEG');
AjoutSelect (stSelect, 'GL_NATUREPIECEG');
AjoutSelect (stSelect, 'GL_ARTICLE');
AjoutSelect (stSelect, 'GL_SOUCHE');
AjoutSelect (stSelect, 'GL_NUMERO');
AjoutSelect (stSelect, 'GL_NUMLIGNE');
AjoutSelect (stSelect, 'GL_DEPOT');
AjoutSelect (stSelect, 'GL_PRIXPOURQTE');
AjoutSelect (stSelect, 'GL_QUALIFQTEVTE');
AjoutSelect (stSelect, 'GL_QUALIFQTESTO');
AjoutSelect (stSelect, 'GL_QUALIFQTEACH');
AjoutSelect (stSelect, 'GL_TYPEARTICLE');
stSelect := stSelect + ' FROM LIGNE LEFT JOIN DEPOTS ON GDE_DEPOT=GL_DEPOT '
            + 'LEFT OUTER JOIN PARPIECE ON GL_NATUREPIECEG=GPP_NATUREPIECEG '
            + 'LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE=GL_ARTICLE ';
QteAffect:=GetControlText('_COLSTOCK');
if (QteAffect='RC') or (QteAffect='PRE') or (QteAffect='LC') or (QteAffect='RF') or (QteAffect='LF') then
   begin
   if stWhere='' then stWhere := 'Where GL_VIVANTE="X" ' else stWhere:=stWhere+' AND GL_VIVANTE="X" ';
   stWhere := stWhere + ' AND GL_QTERESTE > 0'; {DBR NEWPIECE}
   end;
if stWhere='' then
   begin
   stWhere := 'Where ';
   if stWhereNat<>'' then stWhere:=stWhere+stWhereNat ;
   end else
   begin
   if stWhereNat<>'' then stWhere:=stWhere+' AND '+stWhereNat ;
   end;
stWhere:=stWhere+' AND '+stWhereArt ;
NomList := TFMul(Ecran).Q.Liste;
ChargeHListe(NomList, FRecordSource, FLien, FSortBy, FFieldList, FTitre, FLargeur, FAlignement, FParams, tt, NC, Perso, OkTri, OkNumCol);
if Trim(FSortBy) <> '' then FSortBy := ' ORDER BY ' + FSortBy;
stSelect := stSelect+stWhere+
            ' AND (GL_TENUESTOCK="X" OR GL_TYPEARTICLE="NOM") AND '+
            ' GL_TYPELIGNE="ART" AND GL_QUALIFMVT<>"ANN"' + FSortBy;
            //' ORDER BY GL_DATEPIECE, GL_NATUREPIECEG, GL_SOUCHE, GL_NUMERO';
stSelectCount := 'SELECT count(GL_ARTICLE) as nombre FROM LIGNE LEFT JOIN DEPOTS ON GDE_DEPOT=GL_DEPOT ' +
            'LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE=GL_ARTICLE ' + stWhere +
            ' AND (GL_TENUESTOCK="X" OR GL_TYPEARTICLE="NOM") AND ' +
            ' GL_TYPELIGNE="ART" AND GL_QUALIFMVT<>"ANN"';


TSql := OpenSql (stSelectCount, True);
if not TSql.Eof then
    begin
    NbEnreg:=TSql.FindField ('nombre').AsInteger;
    if NbEnreg > 2500 then
        begin
        HShowMessage ('0;?caption?;Selection trop importante ! veuillez la restreindre;W;O;O;O;','','');
        Ferme (TSql);
        end else
        begin
        Ferme (TSql);
        if NbEnreg>0 then
           begin
           TSql := OpenSql (stSelect, True);
           if not TSql.Eof then
              begin
              TobLigne.LoadDetailDB ('LIGNE', '', '', TSql, False);
              TraiteLesLignes(stWhereNomen);
              end;// else TOB.Create ('LIGNE', TobLigneAff, 0);
           Ferme (TSql);
           end;// else TOB.Create ('LIGNE', TobLigneAff, 0);
        end;
    end;
THSQLMemo(GetControl('Z_SQL')).Text := stSelect;
end;

procedure TOF_GCMOUVSTK.LoadLesTobNomen;
var TSql : TQuery;
    stSelect,RefUnique,Depot,QteAffect : string;
    TOBNomen, TOBLN : TOB;
    i_ind : integer ;
    DateDeb, DateFin : TDateTime;
begin
if GetControlText('_ARTICLE')='' then exit;
if stWhereNatNom='' then exit;
DateDeb:=StrToDate(GetControlText('GL_DATEPIECE'));
DateFin:=StrToDate(GetControlText('GL_DATEPIECE_'));
Depot:=GetControlText('GL_DEPOT');
TOBNomen:=TOB.Create('',nil,-1); if TOBNomen=nil then exit;
stSelect := 'Select GL_ARTICLE from LIGNENOMEN,LIGNE WHERE '+
            'GL_NATUREPIECEG=GLN_NATUREPIECEG AND GL_SOUCHE=GLN_SOUCHE AND GL_NUMERO=GLN_NUMERO AND '+
            'GL_INDICEG=GLN_INDICEG AND GL_NUMLIGNE=GLN_NUMLIGNE AND'+stWhereNatNom+
            'AND GL_DATEPIECE>="'+UsDateTime(DateDeb)+'" '+'AND GL_DATEPIECE<="'+UsDateTime(DateFin)+'" '+
            'AND GLN_ARTICLE Like "'+format('%-18.18s',[GetControlText('_ARTICLE')])+'%" AND GLN_TENUESTOCK="X" AND GL_QUALIFMVT<>"ANN"';
if Depot<>'' then stSelect := stSelect+' AND GL_DEPOT="'+Depot+'"';
QteAffect:=GetControlText('_COLSTOCK');
if (QteAffect='RC') or (QteAffect='PRE') or (QteAffect='LC') or (QteAffect='RF') or (QteAffect='LF') then
   begin
   stSelect:=stSelect+' AND GL_VIVANTE="X" ';
   stSelect := stSelect + ' AND GL_QTERESTE > 0'; {DBR NEWPIECE}
   end;
TSql := OpenSql (stSelect, True);
if not TSql.Eof then
   begin
   TOBNomen.LoadDetailDB ('', '', '', TSql, False);
   end;
Ferme (TSql);
for i_ind:=0 to TOBNomen.Detail.Count-1 do
    begin
    TOBLN:=TOBNomen.Detail[i_ind];
    RefUnique:=TOBLN.GetValue('GL_ARTICLE') ;
    if (RefUnique='') or (Pos(RefUnique,stWhereArt)>0) then continue;
    if stWhereArt='' then stWhereArt:='GL_ARTICLE="'+RefUnique+'"'
                     else stWhereArt:=stWhereArt+' OR GL_ARTICLE="'+RefUnique+'"';
    end;
if stWhereArt<>'' then stWhereArt:=' ('+stWhereArt+') ';
TOBNomen.Free;
end;

procedure TOF_GCMOUVSTK.TraiteLesLignes(stWhereNomen : string);
Var stPiecePrecedente, stWhere : string;
    TSql : TQuery;
    iIndex : integer;
    QteStockPre, QteMouvement : double;
    CleDocPrec : R_CleDoc;
    TobPP, TobPiecePrec : Tob; {DBR NEWPIECE}
begin
TobPiecePrec := Tob.Create ('', nil, -1); {DBR NEWPIECE}
for iIndex := 0 to TobLigne.Detail.Count - 1 do
    begin
    QteStockPre := 0.0;
    QteMouvement := 0.0; {DBR NEWPIECE}
    if trim (TobLigne.Detail[iIndex].GetValue ('GL_ARTICLE')) <> '' then
        begin
        if TobLigne.Detail[iIndex].GetValue ('GL_INDICEG') = 0 then
            begin
            stPiecePrecedente := TobLigne.Detail[iIndex].GetValue ('GL_PIECEPRECEDENTE');
            if stPiecePrecedente <> '' then
                begin
                DecodeRefLignePiece (stPiecePrecedente, CleDocPrec);
//                if TypePieceStockPhy (CleDocPrec.NaturePiece) then
                if GetPhyPlusMoins (CleDocPrec.NaturePiece) <> 0 then
                   begin
{DBR NEWPIECE DEBUG}
                   TobPP := TobPiecePrec.FindFirst (['GL_NATUREPIECEG', 'GL_SOUCHE', 'GL_NUMERO', 'GL_INDICEG', 'GL_NUMORDRE'],
                                                    [CleDocPrec.NaturePiece, CleDocPrec.Souche, CleDocPrec.NumeroPiece, CleDocPrec.Indice, CleDocPrec.NumLigne],
                                                    True);
                    if TobPP = nil then
                    begin
//                    stWhere := 'SELECT GL_QTESTOCK, GL_QTERELIQUAT, GL_QTERESTE FROM LIGNE WHERE ' +
                      stWhere := 'SELECT GL_NATUREPIECEG, GL_SOUCHE, GL_NUMERO, GL_INDICEG, GL_NUMORDRE, ' +
                                'GL_QTESTOCK, GL_QTERELIQUAT, GL_QTERESTE FROM LIGNE WHERE ' +
{DBR NEWPIECE FIN}
                                'GL_NATUREPIECEG="' + CleDocPrec.NaturePiece +
                                '" AND GL_DATEPIECE="' + UsDateTime (CleDocPrec.DatePiece) +
                                '" AND GL_SOUCHE="' + CleDocPrec.Souche +
                                '" AND GL_NUMERO=' + inttostr (CleDocPrec.NumeroPiece) +
                                ' AND GL_INDICEG=' + inttostr (CleDocPrec.Indice) +
                                ' AND GL_NUMORDRE=' + inttostr (CleDocPrec.NumLigne);
//                                ' AND GL_NUMLIGNE=' + inttostr (CleDocPrec.NumLigne);  //DB 04112003
                      TSql := OpenSql (stWhere, True);
                      if not TSql.Eof then
                      begin
{DBR NEWPIECE DEBUT}
                        TobPP := Tob.Create ('', TobPiecePrec, -1);
                        TobPP.SelectDB ('', TSql);
                        TobPP.AddChampSupValeur ('QTETRANS', 0);
//                        QteStockPre := StrToFloat (TSql.FindField('GL_QTESTOCK').AsString);
{DBR NEWPIECE FIN}
                      end;
                      ferme (TSql);
{DBR NEWPIECE DEBUT}
                    end;
                    if TobLigne.Detail[iIndex].GetValue ('GL_QTESTOCK') >= TobPP.GetValue ('GL_QTESTOCK') - TobPP.GetValue ('QTETRANS') then
                      begin
//                      QteMouvement := TobLigne.Detail[iIndex].GetValue ('GL_QTESTOCK') +
//                                      TobLigne.Detail[iIndex].GetValue ('GL_QTERELIQUAT') - QteStockPre;
                      QteMouvement := TobLigne.Detail[iIndex].GetValue ('GL_QTESTOCK') -
                                      (TobPP.GetValue ('GL_QTESTOCK') - TobPP.GetValue ('QTETRANS'));
                      end else
                      begin
                      QteMouvement := 0;
                      end;
                   TobPP.PutValue ('QTETRANS', TobPP.GetValue ('QTETRANS') + TobLigne.Detail[iIndex].GetValue ('GL_QTESTOCK'));
{DBR NEWPIECE FIN}
                   end else
                       QteMouvement := TobLigne.Detail[iIndex].GetValue ('GL_QTESTOCK'); // - QteStockPre;
                end else
                begin
                QteMouvement := TobLigne.Detail[iIndex].GetValue ('GL_QTESTOCK'); // - TobLigne.Detail[iIndex].GetValue ('GL_QTERESTE');
                end;
            TobLigne.Detail [iIndex].PutValue ('GL_QTESTOCK', QteMouvement);
            end;
        if (TobLigne.Detail[iIndex].GetValue ('GL_QTESTOCK') <> 0) or
           (stPiecePrecedente = '') then
            begin
            CopieLaLigne (iIndex);
            if TobLigne.Detail[iIndex].GetValue ('GL_TYPEARTICLE') = 'NOM' then
                TraiteLigneNomen(TobLigneAff, TobLigneAff.Detail.Count - 1, stWhereNomen);
            end;
        end;
    end;
if (MOU_Depot <> 0) or (MOU_QualifMvt <> 0) then
    begin
    for iIndex := TobLigneAff.Detail.Count - 1 downto 0 do
        begin
        if TobLigneAff.Detail[iIndex].GetValue('GL_ARTICLE') = '' then
            begin
            TobLigneAff.Detail[iIndex].Free;
            Continue;
            end;
        if MOU_QualifMvt <> 0 then
            begin
            if TobLigneAff.Detail[iIndex].GetValue ('GL_QUALIFMVT') <> '' then
               TobLigneAff.Detail[iIndex].PutValue (
                   'GL_QUALIFMVT',
                   RechDom ('GCQUALIFMVT',
                            TobLigneAff.Detail[iIndex].GetValue ('GL_QUALIFMVT'), False));
            end;
        end;
    end;
TobPiecePrec.Free; {DBR NEWPIECE}
end;

procedure TOF_GCMOUVSTK.TraiteLigneNomen(TobLigne : TOB; iInd : integer; stWhereNomen : string);
var
    TobTemp, TobLigNom : TOB;
    stSQL, stSuf : string;
    TSql : TQuery;
    iIndex1, iIndex2 : integer;
    dTemp : double;

begin
stSQL := 'Select * from LIGNENOMEN where ';
stSQL := stSQL + 'GLN_NATUREPIECEG="' + TobLigne.Detail[iInd].GetValue('GL_NATUREPIECEG') + '" and ';
stSQL := stSQL + 'GLN_SOUCHE="' + TobLigne.Detail[iInd].GetValue('GL_SOUCHE') + '" and ';
stSQL := stSQL + 'GLN_NUMERO=' + IntToStr(TobLigne.Detail[iInd].GetValue('GL_NUMERO')) + ' and ';
stSQL := stSQL + 'GLN_NUMLIGNE=' + IntToStr(TobLigne.Detail[iInd].GetValue('GL_NUMLIGNE')) + ' and ';
stSQL := stSQL + 'GLN_INDICEG=' + IntToStr(TobLigne.Detail[iInd].GetValue('GL_INDICEG')) + ' and ';
stSQL := stSQL + 'GLN_TENUESTOCK="X"';
if stWhereNomen <> '' then
    begin
    stSQL := stSQL + ' and GLN_ARTICLE like "' + stWhereNomen + '%"';
    end;

TSql := OpenSql (stSQL, True);
if not TSql.Eof then
    begin
    TobLigNom := TOB.Create('LIGNENOMEN', nil, -1);
    TobLigNom.LoadDetailDB('LIGNENOMEN', '', '', TSql, False);
    for iIndex1 := 0 to TobLigNom.Detail.Count - 1 do
        begin
//  creation d'une ligne supplementaire pour chaque composant de la nomenclature
        TobTemp := TOB.Create('LIGNE', TobLigne, iInd + 1);
        TobTemp.Dupliquer(TobLigne.Detail[iInd], False, True);
//  on multiplie toutes les qtes par la qte de la ligne composant
        for iIndex2 := 1 to TobTemp.NbChamps - 2 do
            begin
            if Copy(TobTemp.GetNomChamp(iIndex2), 0, 6) = 'GL_QTE' then
                begin
                dTemp := TobTemp.GetValeur(iIndex2);
                dTemp := dTemp * TobLigNom.Detail[iIndex1].GetValue('GLN_QTE');
                TobTemp.PutValeur(iIndex2, dTemp);
                end;
            end;
//  on recupere les autres champs communs entre les deux tables
        for iIndex2 := 1 to TobLigNom.Detail[iIndex1].NbChamps - 2 do
            begin
            stSuf := TobLigNom.Detail[iIndex1].GetNomChamp(iIndex2);
            stSuf := Copy(stSuf, Pos('_', stSuf) + 1, 255);
            if TobTemp.FieldExists('GL_' + stSuf) then
                TobTemp.PutValue('GL_' + stSuf, TobLigNom.Detail[iIndex1].GetValue('GLN_' + stSuf));
            end;
        end;
    TobLigNom.Free;
    end;
TobLigne.Detail[iInd].PutValue('GL_ARTICLE', '');
Ferme(TSql);

end;

(* function TOF_GCMOUVSTK.TypePieceStockPhy (stNaturePiece : string) : boolean;
 var TSql : TQuery;
    stSelect : string;
begin
Result:=False;
if VH_GC.TobParPiece.FindFirst(['GPP_NATUREPIECEG'],[stNaturePiece],false)=Nil then Result:=True;

stSelect := 'SELECT GPP_QTEMOINS FROM PARPIECE ' +
            'WHERE GPP_NATUREPIECEG="' + stNaturePiece + '"' +
            ' AND (GPP_QTEMOINS LIKE "%PHY%" OR GPP_QTEPLUS LIKE "%PHY%")';
TSql := OpenSql (stSelect, True);
if not TSql.Eof then
    Result := True
    else Result := False;
Ferme (TSql);
end; *)

{========================================================================================}
{======================================= Evenements de la Grid ==========================}
{========================================================================================}

procedure TOF_GCMOUVSTK.G_ConsultMouvColumnWidthsChanged(Sender: TObject);
var Coord : TRect;
begin
if PSAICUMUL.ControlCount<=0 then exit ;
Coord:=G_ConsultMouv.CellRect(MOU_QteStock,0);
CUMQTESTO.Left:=Coord.Left + 1;
CUMQTESTO.Width:=G_ConsultMouv.ColWidths[MOU_QteStock] + 1;
end;

procedure TOF_GCMOUVSTK.G_ConsultMouvDblClick(Sender: TObject);
var CleDoc : R_CleDoc;
    stNature, stDate, stSouche, stNumero, stIndiceg : string;
begin
inherited;
stNature := Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_NATUREPIECEG');
stDate := Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_DATEPIECE');
stSouche := Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_SOUCHE');
stNumero := IntToStr(Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_NUMERO'));
stIndiceg := IntToStr(Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_INDICEG'));
if (stNature <> '') and (stDate <> '') and (stSouche <> '') and (stNumero <> '') and
   (stIndiceg <> '') then
    begin
    StringToCleDoc (
        stNature + ';' + stDate + ';' + stSouche + ';' + stNumero + ';' + stIndiceg + ';',
        CleDoc);
    SaisiePiece (CleDoc, taConsult) ;
    end;
end;

procedure TOF_GCMOUVSTK.BAgrandirClick(Sender: TObject);
begin
inherited;
ChangeListeCrit(Ecran,True);
G_ConsultMouv.Top := FListe.Top;
G_ConsultMouv.Left := FListe.Left;
G_ConsultMouv.Height := FListe.Height;
G_ConsultMouv.Width := FListe.Width;
HMTrad.ResizeGridColumns (G_ConsultMouv);
end;

procedure TOF_GCMOUVSTK.BReduireClick(Sender: TObject);
begin
inherited;
ChangeListeCrit(Ecran,False);
G_ConsultMouv.Top := FListe.Top;
G_ConsultMouv.Left := FListe.Left;
G_ConsultMouv.Height := FListe.Height;
G_ConsultMouv.Width := FListe.Width;
HMTrad.ResizeGridColumns (G_ConsultMouv);
end;

procedure TOF_GCMOUVSTK.BOuvrirClick(Sender: TObject);
var CleDoc : R_CleDoc;
    stNature, stDate, stSouche, stNumero, stIndiceg : string;
begin
inherited;
stNature := Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_NATUREPIECEG');
stDate := Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_DATEPIECE');
stSouche := Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_SOUCHE');
stNumero := IntToStr(Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_NUMERO'));
stIndiceg := IntToStr(Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_INDICEG'));
StringToCleDoc (
    stNature + ';' + stDate + ';' + stSouche + ';' + stNumero + ';' + stIndiceg + ';',
    CleDoc);
SaisiePiece (CleDoc, taConsult) ;
end;

procedure TOF_GCMOUVSTK.BImprimerClick(Sender: TObject);
begin
Imprime ;
end;

function TOF_GCMOUVSTK.PrepareImpression : integer ;
var i_ind : integer;
    TOBGZF, TOBD, TOBL : TOB;
    PrefixeTT : string;
BEGIN
PrefixeTT := TableToPrefixe('GCTMPMOUV');
Result:=0;
ExecuteSQL('DELETE FROM GCTMPMOUV WHERE '+PrefixeTT+'_UTILISATEUR = "'+V_PGI.USer+'"');
TOBGZF:=TOB.Create('',Nil,-1);
for i_ind:=0 to TobLigneAff.Detail.Count-1 do
    BEGIN
    TOBL:=TobLigneAff.Detail[i_ind] ;
    TOBD:=TOB.Create('GCTMPMOUV',TOBGZF,-1);
    TOBD.PutValue(PrefixeTT+'_UTILISATEUR',V_PGI.USer);
    TOBD.PutValue(PrefixeTT+'_COMPTEUR',i_ind);
    TOBD.PutValue(PrefixeTT+'_NATUREPIECEG',TOBL.GetValue('GL_NATUREPIECEG'));
    TOBD.PutValue(PrefixeTT+'_SOUCHE',TOBL.GetValue('GL_SOUCHE'));
    TOBD.PutValue(PrefixeTT+'_NUMERO',TOBL.GetValue('GL_NUMERO'));
    TOBD.PutValue(PrefixeTT+'_INDICEG',TOBL.GetValue('GL_INDICEG'));
    TOBD.PutValue(PrefixeTT+'_DATEPIECE',TOBL.GetValue('GL_DATEPIECE'));
    TOBD.PutValue(PrefixeTT+'_TIERS',TOBL.GetValue('GL_TIERS'));
    TOBD.PutValue(PrefixeTT+'_CODEARTICLE',TOBL.GetValue('GL_ARTICLE'));
//    TOBD.PutValue(PrefixeTT+'_QUALIFMVT',TOBL.GetValue('GL_QUALIFMVT'));
//    TOBD.PutValue(PrefixeTT+'_DEPOT',TOBL.GetValue('GL_DEPOT'));
    TOBD.PutValue(PrefixeTT+'_QUALIFMVT',G_ConsultMouv.Cells[MOU_QualifMvt, i_ind + 1]);
    TOBD.PutValue(PrefixeTT+'_DEPOT',G_ConsultMouv.Cells[MOU_Depot, i_ind + 1]);
    TOBD.PutValue(PrefixeTT+'_QTESTOCK',TOBL.GetValue('GL_QTESTOCK'));
    TOBD.PutValue(PrefixeTT+'_DPA',TOBL.GetValue('GL_DPA'));
    END;
TOBGZF.InsertDB (Nil);
TOBGZF.Free;
END;

procedure TOF_GCMOUVSTK.Imprime ;
BEGIN
LanceEtatTob('E','GZH','GZM',TobLigneAff,True,False,False,TPageControl(GetControl('Pages')),'',Ecran.Caption,False);
END;

procedure TOF_GCMOUVSTK.RecupControl;
begin
BAgrandir:=TToolBarButton97(GetControl('BAGRANDIR'));
BAgrandir.OnClick:=BAgrandirClick;
BReduire:=TToolBarButton97(GetControl('BREDUIRE'));
BReduire.OnClick:=BReduireClick;
BOuvrir:=TToolBarButton97(GetControl('BOUVRIR'));
BOuvrir.OnClick:=BOuvrirClick;
BImprimer:=TToolBarButton97(GetControl('BIMPRIMER'));
BImprimer.OnClick:=BImprimerClick;
PSAICUMUL := THPanel(GetControl('PSAICUMUL'));
CUMQTESTO := THNumEdit(GetControl('CUMQTESTO'));
HMTrad:=THSystemMenu(GetControl('HMTrad'));
{$IFDEF EAGLCLIENT}
FListe:=THGrid(GetControl('FListe'));
{$ELSE}
FListe:=THDBGrid(GetControl('FListe'));
{$ENDIF}
G_ConsultMouv:=THGrid(GetControl('G_CONSULTMOUV'));
G_ConsultMouv.OnDblClick:=G_ConsultMouvDblClick;
G_ConsultMouv.OnColumnWidthsChanged := G_ConsultMouvColumnWidthsChanged;
end;

procedure TOF_GCMOUVSTK.CreeTotal (ColLig : integer; Cumul : THNumEdit);
var Coord : TRect;
begin
Cumul.ParentColor := True;
Cumul.Font.Style := PSAICUMUL.Font.Style;
Cumul.Font.Size := PSAICUMUL.Font.Size;
Cumul.Masks.PositiveMask := G_ConsultMouv.ColFormats[ColLig];
Cumul.Ctl3D := False;
Cumul.Top := 0;
Coord := G_ConsultMouv.CellRect (ColLig, 0);
Cumul.Left := Coord.Left;
Cumul.Width := G_ConsultMouv.ColWidths[ColLig] + 1;
Cumul.Height := PSAICUMUL.Height;
Cumul.Enabled := False;
end;

procedure TOF_GCMOUVSTK.MiseEnForme;
begin
G_ConsultMouv.Top := FListe.Top;
G_ConsultMouv.Left := FListe.Left;
G_ConsultMouv.Height := FListe.Height;
G_ConsultMouv.Width := FListe.Width;
G_ConsultMouv.ListeParam := TFMul(Ecran).DBListe;
end;

Procedure TOF_GCMOUVSTK.MyExportClick(Sender: TObject);
Var SavDia: THSaveDialog;
begin
SavDia:=TFMUL(Ecran).SD;
if SavDia.Execute then
   begin
   ExportGrid(G_ConsultMouv,Nil,SavDia.FileName,SavDia.FilterIndex,TRUE) ;
   end;
end;

Procedure TOF_GCMOUVSTK.VoirDispo ;
Var RefUnique, Depot : string ;
begin
RefUnique := Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_ARTICLE');
Depot := Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_DEPOT');
if (RefUnique<>'') and (depot<>'') then AGLLanceFiche('BTP','BTDISPO','',RefUnique+';'+Depot,'ACTION=CONSULTATION;') ;
end;

Procedure TOF_GCMOUVSTK.VoirArticle ;
Var RefUnique : string ;
begin
RefUnique := Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_ARTICLE');
{$IFNDEF GPAO}
  if RefUnique <> '' then
    DispatchTTArticle(TaConsult,RefUnique,'ACTION=CONSULTATION;TARIF=0;MONOFICHE','');
{$ELSE}
  if RefUnique <> '' then
    V_PGI.DispatchTT(7, taConsult, RefUnique, 'TARIF=0;MONOFICHE', '');
{$ENDIF GPAO}
end;

Procedure TOF_GCMOUVSTK_VoirDispo (parms:array of variant; nb: integer ) ;
var F: TFMul;
    MaTOF : TOF ;
BEGIN
F := TFMul(Longint (Parms[0]));
if (F.Name <> 'GCMOUVSTK') and (F.Name <> 'MOUVSTK') then exit;
MaTOF:=TFMul(F).LaTOF;
if (MaTOF is TOF_GCMOUVSTK) then TOF_GCMOUVSTK(MaTOF).VoirDispo;
end;

Procedure TOF_GCMOUVSTK_VoirArticle (parms:array of variant; nb: integer ) ;
var F: TFMul;
    MaTOF : TOF ;
BEGIN
F := TFMul(Longint (Parms[0]));
if (F.Name <> 'GCMOUVSTK') and (F.Name <> 'MOUVSTK') then exit;
MaTOF:=TFMul(F).LaTOF;
if (MaTOF is TOF_GCMOUVSTK) then TOF_GCMOUVSTK(MaTOF).VoirArticle;
end;

procedure InitTOF_GCMOUVSTK() ;
begin
RegisterAglProc('GCMOUVSTK_VoirDispo', True , 0, TOF_GCMOUVSTK_VoirDispo);
RegisterAglProc('GCMOUVSTK_VoirArticle', True , 0, TOF_GCMOUVSTK_VoirArticle);
end;

Initialization
registerclasses ( [ TOF_GCMOUVSTK ] ) ;
InitTOF_GCMOUVSTK();
end.
