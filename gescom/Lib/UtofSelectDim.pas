unit UtofSelectDim;

interface
uses  StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}DBGrids, db,
{$ENDIF}
      forms, sysutils, ComCtrls, EntGC, ParamSoc,
      HCtrls,HEnt1,HMsgBox,UTOF, vierge, HDimension,UTOB, UtilArticle,UdimArticle, AglInit;
Type
     TOF_SELECTDIM = Class (TOF)
        DimensionsArticle : TODimArticle ;
     private
        dimAction : string ;  // SELECT, SAISIE, CONSULT, MULTI
        dimChamp : string ;
        dimMasque : string ;
        CodeArticle : string ;
        PrixUnique : boolean ;
        ChoixAvecCombo: boolean;
        DimError: boolean;
        procedure FormKeyDown ( Sender : TObject ; var Key : Word ;  Shift : TShiftState ) ;
        procedure OnDoubleClick(Sender: TObject);
        procedure InitDimensionsArticle ;
        function GetComboValue(var Indice: integer; var Valeur: string): boolean;
        procedure ChargeCombo(Combo: THValComboBox; Libelle: string; var Indice: integer);
        procedure InitSaisieDimensions;
     public
        procedure OnArgument (Arguments : String ) ; override ;
        procedure OnUpdate; override ;
        procedure OnClose; override ;
        Destructor Destroy ; override ;
     END ;

{ GP Mode Début}
   function uGetArticleDimSelect :string;
{ GP Mode Fin }

implementation

{ GP Mode Début}
var
   ArticleDim : string ;
{ GP Mode Fin }

destructor TOF_SELECTDIM.destroy ;
begin
if DimensionsArticle<>nil then DimensionsArticle.Free ;
inherited ;
end;

procedure TOF_SELECTDIM.OnArgument(Arguments : String ) ;
var Lequel : string ;
    Critere : string ;
    ChampMul,ValMul : string;
    x : integer ;
    SQL : string ;
    Q : TQuery ;
begin
inherited ;
dimAction:='SELECT';
dimChamp:='' ;
DimError := False;

Repeat
    Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
    if Critere<>'' then
        begin
        x:=pos('=',Critere);
        if x<>0 then
           begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));

           if ChampMul='GA_ARTICLE' then Lequel := ValMul;
           if ChampMul='ACTION' then dimAction := ValMul;
           if ChampMul='CHAMP' then dimChamp := ValMul;
           end;
        end;
until  Critere='';

SQL := 'SELECT GA_ARTICLE, GA_CODEARTICLE, GA_DIMMASQUE, GA_STATUTART, GA_PRIXUNIQUE ' ;
if dimChamp<>'' then SQl:=SQL+','+dimChamp ;
SQL := SQL+' from Article where GA_ARTICLE="'+Lequel+'"' ;
Q:=OpenSQl(SQL,True) ;
if Q.EOF  then begin Ferme(Q) ; TFVierge(Ecran).Close ; Exit ; end;
if Q.FindField('GA_STATUTART').AsString <>'GEN' then
 begin
 HShowMessage( '0;'+'Article'+';'+TraduireMemoire('Cet article n''a pas de dimension')+';E;O;O;O','','') ;
 Ferme(Q); TFVierge(Ecran).close;
 Exit;
 end;
//SetControlText('GA_ARTICLE',Lequel);
SetControlText('DIM_CHAMP_AFFICHE',dimChamp);
CodeArticle:=Q.FindField('GA_CODEARTICLE').AsString ;
PrixUnique:=Boolean(Q.FindField('GA_PRIXUNIQUE').AsString='X') ;
SetControlText('GA_CODEARTICLE',CodeArticle) ;
dimMasque:=Q.FindField('GA_DIMMASQUE').AsString ;
SetControlText('GA_STATUTART',Q.FindField('GA_STATUTART').AsString) ;
ChoixAvecCombo := not GetParamSoc('SO_GCMATRICEDIM');
InitDimensionsArticle ;
Ferme(Q) ;
Ecran.OnKeyDown := FormKeyDown ;
if CodeArticle <> '' then Ecran.Caption := Ecran.Caption + ' : ' + CodeArticle;
end;

procedure TOF_SELECTDIM.FormKeyDown ( Sender : TObject ; var Key : Word ;  Shift : TShiftState ) ;
begin
if Key=vk_valide then BEGIN Key:=0 ; if (Ecran<>Nil) and (Ecran is TFVierge) then TFVierge(Ecran).BValider.Click ; END ;
end;

procedure TOF_SELECTDIM.InitDimensionsArticle ;
var NatureDoc: String ;
begin
    if DimensionsArticle<>nil then
        begin
        DimensionsArticle.free;
        end;
    if (ctxFo in V_PGI.PGIContexte) then NatureDoc := NAT_VENTESFFO else NatureDoc := NAT_ARTICLE ;
    DimensionsArticle:=TODimArticle.Create(THDimension(GetControl('FDIM'))
                         ,CodeArticle
                         ,dimMasque
                         ,dimChamp,'GCDIMCHAMP',NatureDoc,'', '', '', PrixUnique);
    DimensionsArticle.Dim.OnDblClick:=OnDoubleClick ;
    DimensionsArticle.Action := taConsult ;
    DimensionsArticle.RefreshGrid(True) ;
    // DCA - FQ MODE 10681
    // Rend invisible l'option "Rendre existant" du "collage spéciale" de l'objet dimension (Option disponible en saisie de pièce / saisie de transferts)
    DimensionsArticle.Dim.Options.RendreExistanteVisible := False;
    // Rend invisible l'option "Rendre existant" du "Aller à" de l'objet dimension (Option disponible en saisie de transferts)
    DimensionsArticle.Dim.Options.AllerARendreExistanteVisible := False;

    if ChoixAvecCombo then InitSaisieDimensions;
end;

procedure TOF_SELECTDIM.OnDoubleClick(Sender: TObject);
var QMasque : TQuery ;
{ GP Mode Début}
//    ArticleDim : string ;
{ GP Mode Fin }
    ItemDim :THDimensionItem ;
    TobSelect : TOB ;
begin
ItemDim:=THDimensionItem(Sender);
if ItemDim = nil then exit;
if ctxMode in V_PGI.PGIContexte
then QMasque:=OpenSQL('SELECT * FROM DIMMASQUE WHERE GDM_MASQUE="'+DimMasque+'" and GDM_TYPEMASQUE="'+VH_GC.BOTypeMasque_Defaut+'"',TRUE)
else QMasque:=OpenSQL('SELECT * FROM DIMMASQUE WHERE GDM_MASQUE="'+DimMasque+'"',TRUE) ;
if not QMasque.EOF then
    begin
    ArticleDim:= CodeArticleUnique(CodeArticle,AffecteDim(QMasque,1,ItemDim),AffecteDim(QMasque,2,ItemDim),AffecteDim(QMasque,3,ItemDim),AffecteDim(QMasque,4,ItemDim),AffecteDim(QMasque,5,ItemDim));
    if LaTob<>Nil then
        begin
        TobSelect:=tob.create('articles select',LaTob, -1 ) ;
        TobSelect.AddChampSup( 'GA_ARTICLE',false);
        TobSelect.PutValue('GA_ARTICLE',ArticleDim) ;
        TobSelect.AddChampSup( 'VALEUR',false);
        TobSelect.PutValue('VALEUR',ItemDim.Valeur[1]) ;

        TheTob:=LaTob ;
        end;
    end;
ferme(Qmasque);
TFVierge(Ecran).close ;
end;

procedure TOF_SELECTDIM.OnUpdate;
var Ind: integer;
  Ctrl: TControl;
  StOng, StLig1, StLig2, StCol1, StCol2: string;
begin
  inherited ;
  DimError := False;
  if ChoixAvecCombo then
  begin
    Ctrl := GetControl('FDIM');
    if (Ctrl = nil) or not (Ctrl is THDimension) then Exit;
    Ind := 0;
    if THDimension(Ctrl).DimOngl <> nil then
      if not GetComboValue(Ind, StOng) then Exit;
    if THDimension(Ctrl).DimLig1 <> nil then
      if not GetComboValue(Ind, StLig1) then Exit;
    if THDimension(Ctrl).DimLig2 <> nil then
      if not GetComboValue(Ind, StLig2) then Exit;
    if THDimension(Ctrl).DimCol1 <> nil then
      if not GetComboValue(Ind, StCol1) then Exit;
    if THDimension(Ctrl).DimCol2 <> nil then
      if not GetComboValue(Ind, StCol2) then Exit;
    if not THDimension(Ctrl).GoToValeurDimension(StOng, StLig1, StLig2, StCol1, StCol2, False) then
    begin
      PgiBox('Les dimensions recherchées n''existent pas.');
      DimError := True;
      Exit;
    end;
  end;
  OnDoubleClick(DimensionsArticle.Dim.CurrentItem);
end;

procedure TOF_SELECTDIM.OnClose;
begin
  inherited ;
  if DimError then
    LastError := 1
  else
    LastError := 0;
  DimError := False;
end;

function TOF_SELECTDIM.GetComboValue(var Indice: integer; var Valeur: string): boolean;
var Ctrl: TControl;
  Txt, Stg: string;
  Ind, jj: integer;
begin
  Result := False;
  Valeur := '';
  Inc(Indice);
  Ctrl := GetControl('GDI_LIBELLE_' + IntToStr(Indice));
  if (Ctrl <> nil) and (Ctrl is THValComboBox) then
  begin
    Txt := TrimRight(THValComboBox(Ctrl).Text);
    if Trim(Txt) = '' then
      Ind := 0
    else
      Ind := THValComboBox(Ctrl).Items.IndexOf(Txt);
    if (Ind < 0) and (Trim(Txt) <> '') then
    begin
      // recherche sans dictinction de case
      for jj := 0 to THValComboBox(Ctrl).Items.Count - 1 do
        if AnsiCompareText(Txt, THValComboBox(Ctrl).Items[jj]) = 0 then
        begin
          Ind := jj;
          Break;
        end;
    end;
    if (Ind >= 0) and (Ind < THValComboBox(Ctrl).Values.Count) then
    begin
      Valeur := THValComboBox(Ctrl).Values[Ind];
      Result := True;
    end else
    begin
      Stg := TraduireMemoire('La dimension') +' "'+ Txt +'" '+ TraduireMemoire('n''existe pas.');
      PgiBox(Stg);
      DimError := True ;
    end;
  end;
end;

procedure TOF_SELECTDIM.ChargeCombo(Combo: THValComboBox; Libelle: string; var Indice: integer);
var Ctrl: TControl;
begin
  if (Combo = nil) or (Combo.Values.Count <= 0) then Exit;
  Inc(Indice);
  Ctrl := GetControl('GDI_LIBELLE_' + IntToStr(Indice));
  if (Ctrl <> nil) and (Ctrl is THValComboBox) then
  begin
    THValComboBox(Ctrl).Visible := True;
    THValComboBox(Ctrl).Values.Clear;
    THValComboBox(Ctrl).Values.Assign(Combo.Values);
    THValComboBox(Ctrl).Items.Clear;
    THValComboBox(Ctrl).Items.Assign(Combo.Items);
  end;
  Ctrl := GetControl('TGDI_LIBELLE_' + IntToStr(Indice));
  if (Ctrl <> nil) and (Ctrl is TLabel) then
  begin
    TLabel(Ctrl).Visible := True;
    TLabel(Ctrl).Caption := Libelle;
  end;
end;

procedure TOF_SELECTDIM.InitSaisieDimensions;
var Ind: integer;
  Ctrl: TControl;
begin
  Ctrl := GetControl('FDIM');
  if (Ctrl = nil) or not (Ctrl is THDimension) then Exit;
  Ctrl.Visible := False;
  Ind := 0;
  ChargeCombo(THDimension(Ctrl).DimOngl, THDimension(Ctrl).TitreOngl, Ind);
  ChargeCombo(THDimension(Ctrl).DimLig1, THDimension(Ctrl).TitreLig1, Ind);
  ChargeCombo(THDimension(Ctrl).DimLig2, THDimension(Ctrl).TitreLig2, Ind);
  ChargeCombo(THDimension(Ctrl).DimCol1, THDimension(Ctrl).TitreCol1, Ind);
  ChargeCombo(THDimension(Ctrl).DimCol2, THDimension(Ctrl).TitreCol2, Ind);
  // retaillage de la fiche
  Ctrl := GetControl('GDI_LIBELLE_' + IntToStr(Ind));
  if Ctrl <> nil then
  begin
    TFVierge(Ecran).FormResize := False;
    Ecran.Width := Ctrl.Left + Ctrl.Width + 60;
    Ecran.Height := Ctrl.Top + Ctrl.Height + 120;
    Ctrl := GetControl('HelpBtn');
    if Ctrl <> nil then Ctrl.Left := Ecran.Width - (Ctrl.Width + 4) - 12;
    Ctrl := GetControl('BFerme');
    if Ctrl <> nil then Ctrl.Left := Ecran.Width - ((Ctrl.Width + 4) * 2) - 12;
    Ctrl := GetControl('BValider');
    if Ctrl <> nil then Ctrl.Left := Ecran.Width - ((Ctrl.Width + 4) * 3) - 12;
  end;
end;

{ GP Mode Début}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Florence DURANTET
Créé le ...... : 20/11/2002
Modifié le ... :   /  /
Description .. : Retourne l'article dimensionné sélectionné dans la grille
Mots clefs ... : ARTICLE DIMENSIONNE
*****************************************************************}
function uGetArticleDimSelect :string;
begin
   Result:=ArticleDim;
end;
{ GP Mode Fin }

Initialization
registerclasses([TOF_SelectDim]);
end.
