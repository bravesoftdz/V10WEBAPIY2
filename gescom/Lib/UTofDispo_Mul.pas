unit UTofDispo_Mul;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, HDimension,HDB,UTOM,
      AglInit,UTOB,Dialogs,Menus, M3FP, EntGC, Ent1,
{$IFDEF EAGLCLIENT}
      emul, MaineAGL,
{$ELSE}
      Fiche, FichList, mul, db,DBGrids, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ENDIF}
      grids,LookUp,AglInitGC;

Type
     TOF_DISPO_MUL = Class (TOF)
        procedure OnArgument(Arguments : String) ; override ;
        Procedure OnUpdate ; override ;
        Procedure AffiGrillCodedimension(F : TFMUL ; Masque : string);
        Procedure CacheDim ;
        procedure CodeArticleRech (Sender : Tobject);
     END ;

    var CoordDim  : Array of integer ;//Tableau contenant les positions des
                //Controls liés aux dimensions de l'article

    var ArtOK : boolean ; //True si l'article a été recherché avec le Mul Article
    var CodArtPreced : string ; //Article précédent

Implementation

procedure TOF_Dispo_Mul.OnArgument(Arguments : String) ;
Var CC      : THValComboBox;
    CC1 : Thedit;
Begin
  Inherited ;
  ArtOK:=False ;
  CodArtPreced:='' ;
  {$IFDEF CCS3}
  SetControlVisible ('GQ_EMPLACEMENT', False);
  SetControlVisible ('TGQ_EMPLACEMENT', False);
  {$ENDIF}

  //Gestion Restriction Domaine et Etablissements
  CC:=THValComboBox(GetControl('GA_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

  THEdit(GetControl('GA_CODEARTICLE')).onElipsisClick := CodeArticleRech;

End ;

procedure TOF_Dispo_Mul.OnUpdate ;
var   DimMasque, StatutArt : string;
      QArt, QStatut, QDim : TQuery ;
      FF : TFMUL;
Begin
inherited ;
FF:=TFMul(Ecran) ;

//Recherche de l'identifiant article si le code n'a pas été recherché dans le mul article
If ((not ArtOK) and (GetControlText('GA_CODEARTICLE')<>CodArtPreced)) then
    begin
    QArt:=OpenSQL('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="'
                        +GetControlText('GA_CODEARTICLE')+'" AND GA_STATUTART<>"DIM"',TRUE) ;
    if not QArt.EOF then  SetControlText('ARTICLE',QArt.FindField('GA_ARTICLE').AsString)
       else SetControlText('ARTICLE','') ;
    ferme(QArt) ;
    end ;

StatutArt:='' ;
if GetControlText('GA_CODEARTICLE')<>'' then
   begin
   QStatut:=OpenSQL('SELECT GA_STATUTART FROM ARTICLE WHERE GA_ARTICLE="'
                        +GetControlText('ARTICLE')+'"',TRUE) ;
   If not QStatut.Eof then StatutArt:=QStatut.FindField('GA_STATUTART').AsString ;
   Ferme(QStatut) ;
   end;
If StatutArt<>'GEN' then  CacheDim      //Cache les colonnes Dimensions dans FListe
   else
   begin
   if CodArtPreced=GetControlText('GA_CODEARTICLE') then exit;
   QDim:=OpenSQL('SELECT GA_DIMMASQUE FROM ARTICLE WHERE GA_ARTICLE="'
          +GetControlText('ARTICLE')+'"',TRUE) ;
   if not QDim.Eof then DimMasque:=QDim.FindField('GA_DIMMASQUE').AsString ;
   Ferme(QDim) ;
   If (DimMasque<>'') then  AffiGrillCodedimension(FF , DimMasque)
      else CacheDim ;
   end ;
ArtOK:=False ;
CodArtPreced:=GetControlText('GA_CODEARTICLE') ;
end;

//Active les grilles de dimension si l'article selectionné est générique
Procedure TOF_DISPO_MUL.AffiGrillCodedimension(F: TFMUL ; Masque : string);
var QMasque : TQuery;
    TLIB, TLIB2 : THLabel;
    CHPS, CHPS2 : THValComboBox;
    CodeDim : THEdit;
    i_ind, i_Dim : integer;
BEGIN
i_Dim := 1;
if ctxMode in V_PGI.PGIContexte
then QMasque:=OpenSQL('SELECT GDM_TYPE1,GDM_TYPE2,GDM_TYPE3,GDM_TYPE4,GDM_TYPE5 FROM DIMMASQUE WHERE GDM_MASQUE="'+Masque+'" and GDM_TYPEMASQUE="'+VH_GC.BOTypeMasque_Defaut+'"',TRUE)
else QMasque:=OpenSQL('SELECT GDM_TYPE1,GDM_TYPE2,GDM_TYPE3,GDM_TYPE4,GDM_TYPE5 FROM DIMMASQUE WHERE GDM_MASQUE="'+Masque+'"',TRUE) ;
if not QMasque.EOF then
  begin
   for i_ind := 1 to MaxDimension do
    BEGIN
    TLIB := THLabel (F.FindComponent ('GRILLIBEL' + IntToStr (i_Dim)));
    CHPS := THValComboBox (F.FindComponent ('CODEDIM' + IntToStr (i_Dim)));
    TLIB2 := THLabel (F.FindComponent ('GRILLIBEL' + IntToStr (i_ind)));
    CHPS2 := THValComboBox (F.FindComponent ('CODEDIM' + IntToStr (i_ind)));
    CodeDim := THEdit (F.FindComponent ('GA_CODEDIM' + IntToStr (i_ind)));
    CodeDim.Text:='';
    TLIB2.Caption := '';
    CHPS2.Text := '';
    CHPS2.Visible := False;
    if (QMasque.Findfield('GDM_TYPE'+IntToStr(i_ind)).AsString<>'') then
        BEGIN
        CHPS.DataType:='GCDIMENSION' ;
        CHPS.Plus := 'GDI_TYPEDIM="DI'+InttoStr(i_ind)+'" AND GDI_GRILLEDIM="'
                +QMasque.FindField('GDM_TYPE'+IntToStr(i_ind)).AsString+'" ORDER BY GDI_RANG' ;
        CHPS.Visible := True;
        TLIB.Caption:= RechDom('GCGRILLEDIM'+IntToStr(i_ind),
                        QMasque.Findfield('GDM_TYPE'+IntToStr(i_ind)).AsString,FALSE) ;
        Inc (i_Dim);
        END else
            BEGIN
            TLIB.Caption := '';
            CHPS.Value := '';
            CHPS.Visible := False;
            END;
    END ;
  END;
Ferme(QMasque) ;
END;

//Cache les controles associés aux dimensions de l'article
Procedure TOF_DISPO_MUL.CacheDim ;
Var   i:integer;
      F : TFMUL;
Begin
F := TFMul(Ecran) ;
{$IFDEF EAGLCLIENT}
for i:=0 to F.FListe.ColCount-1 do
    begin
    if copy(F.FListe.Cells[i,0],1,3)='DIM' then
       F.Fliste.ColWidths[i]:=-1 ;
    end;
{$ELSE}
for i:=0 to F.FListe.Columns.Count-1 do
    begin
    if copy(F.FListe.Columns[i].Title.caption,1,3)='DIM' then
       F.Fliste.columns[i].visible:=false ;
    end;
{$ENDIF}
for i:=1 to MaxDimension do
    begin
    SetControlText('CODEDIM'+InttoStr(i),'');
    SetControlText('GA_CODEDIM'+InttoStr(i),'');
    SetControlProperty('CODEDIM'+InttoStr(i),'PLUS', '');
    SetControlProperty('CODEDIM'+InttoStr(i),'VISIBLE',FALSE);
    SetControlText('GRILLIBEL'+InttoStr(i),'');
    end ;
end ;

(*------------------------------------RECHERCHE ARTICLES----------------------------------
-----------------------------------------------------------------------------------------*)

Procedure TOFDispoMUL_RechercheArticle (Parms : Array of variant; nb: integer);
var F           : TForm;
    G_Article : THCritMaskEdit;
    QArt : TQuery ;
BEGIN
F := TForm (Longint (Parms[0]));
{$IFDEF GPAO} 
if (F.Name <> 'WDISPO_MUL') and (F.Name <> 'WDISPOART_MUL') then exit;
{$ELSE}
if (F.Name <> 'BTDISPO_MUL') and (F.Name <> 'GCDISPOART_MUL') then exit;
{$ENDIF GPAO}
G_Article := THCritMaskEdit (F.FindComponent (string (Parms [1])));
DispatchRecherche (G_Article, 1, '',
                   'GA_CODEARTICLE=' + Trim (Copy (G_Article.Text, 1, 18)), '');
if G_Article.Text <> '' then
    Begin
    ArtOK:=True;
    THEdit(F.FindComponent ('ARTICLE')).Text:= G_Article.Text;
    QArt:=OpenSQL('SELECT GA_CODEARTICLE FROM ARTICLE WHERE GA_ARTICLE="'
                                    +G_Article.Text+ '"',True) ;
    If not QArt.EOF then
         THEdit(F.FindComponent('GA_CODEARTICLE')).Text:=
                                    QArt.FindField('GA_CODEARTICLE').AsString ;
    Ferme(QArt) ;
    end ;
END;

///Position des THEdit Libellés des Grilles de Dimension
Procedure TOFDispoMUL_InitForm (Parms : Array of variant; nb: integer) ;
var F     : TForm ;
    i_ind : integer ;
    TLIB : THLabel;
    TCombo : THValComboBox;
BEGIN
F := TForm (Longint (Parms[0]));
{$IFDEF GPAO} 
if (F.Name <> 'WDISPO_MUL') and (F.Name <> 'WDISPOART_MUL') then exit;
{$ELSE}
if (F.Name <> 'BTDISPO_MUL') and (F.Name <> 'GCDISPOART_MUL') then exit;
{$ENDIF GPAO}
SetLength (CoordDim, MaxDimension) ;
For i_ind := Low (CoordDim) to High (CoordDim) do
    BEGIN
    CoordDim[i_ind]:= THValComboBox(F.FindComponent('CODEDIM' + IntToStr(i_ind + 1))).Left ;
    TLIB := THLabel(F.FindComponent('GRILLIBEL'+ IntTostr(i_ind + 1)));
    TCombo := THValComboBox(F.FindComponent('CODEDIM'+IntTostr(i_ind + 1)));
    TLIB.Caption := '';
    TCombo.Text := '';
    TCombo.Visible := Not (TCombo.Text = '');
    END;
END;

procedure TOFDispoMUL_DispoMulRechCodeDim(Parms : Array of variant; nb: integer) ;
var F     : TForm ;
    i_ind, IndiceVal  : integer ;
    CodePlus : string;
    CodeDim : THValComboBox;
begin
F := TForm (Longint (Parms[0]));
{$IFDEF GPAO} 
if (F.Name <> 'WDISPO_MUL') and (F.Name <> 'WDISPOART_MUL') then exit;
{$ELSE}
if (F.Name <> 'GCDISPO_MUL') and (F.Name <> 'GCDISPOART_MUL') then exit
Else if (F.Name <> 'BTDISPO_MUL') and (F.Name <> 'GCDISPOART_MUL') then exit;
{$ENDIF GPAO}
IndiceVal:= StrToInt(Parms[1]);
CodeDim := THValComboBox(F.FindComponent('CODEDIM'+IntToStr(IndiceVal)));
CodePlus := CodeDim.Plus;
i_ind:=strToInt(copy(CodePlus,16,1));
THEdit(F.FindComponent('GA_CODEDIM'+IntToStr(i_ind))).Text:= CodeDim.Value;
end;

Procedure InitTOFDispo ;
begin
    RegisterAglProc('RechercheArticleDispo', True , 1, TOFDispoMUL_RechercheArticle);
    RegisterAglProc('InitDispoMul', True , 0, TOFDispoMUL_InitForm);
    RegisterAglProc('DispoMulRechCodeDim', True , 1, TOFDispoMUL_DispoMulRechCodeDim);
End ;

procedure TOF_DISPO_MUL.CodeArticleRech(Sender: Tobject);
Var CodeArt : String;
    StChamps: String;
begin

  CodeArt := THEdit(GetControl('GA_CODEARTICLE')).text;
  StChamps:= '';

  if CodeArt <> '' then StChamps := 'GA_CODEARTICLE=' + Trim(Copy(CodeArt, 1, 18)) + ';';

  StChamps:= StChamps + 'XX_WHERE=GA_TYPEARTICLE="MAR" OR GA_TYPEARTICLE="ARP" OR GA_TYPEARTICLE="OUV"';

  CodeArt := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', StChamps + ';GA_TENUESTOCK=X;ACTION=CONSULTATION;STATUTART=UNI,DIM');

  if codeArt <> '' then
  begin
    THEdit(sender).text := CodeArt;
  end;

end;

Initialization
registerclasses([TOF_Dispo_Mul]);
InitTOFDispo ;
end.
