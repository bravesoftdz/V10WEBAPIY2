unit UTofMBOPrMinMax;

interface
uses  M3FP,StdCtrls,Controls,Classes,Dialogs,UDimArticle,saisutil,
      UtilDimArticle,HStatus,
      HCtrls,HEnt1,HMsgBox, Hdimension, UTOB, UTOF, AGLInit,EntGC,
{$IFDEF EAGLCLIENT}
      Maineagl,
{$ELSE}
      dbTables,DBGrids, db,FE_Main,
{$ENDIF}
      forms,sysutils,ComCtrls,
      HDB, vierge, Math ;


Const MaxDimChamp = 10;

Type
     TOF_MBOPRMINMAX = Class (TOF)
        DimensionsArticle : TODimArticle ;
     private
        OK_MAJ : boolean ;
        CodeMasque,CodeSerie:String;
        TobArtSerie : TOB;
        MsgEncours : boolean;
        procedure InitDimensionsArticle ;
     public
        TobListe,TobMethode : TOB ;
        GMASQUE, GMETHODE : THGrid ;
        procedure OnClose ; override ;
        Procedure OnLoad  ; override ;
        Procedure MajDim  ;
        procedure OnArgument (Arguments : String ) ; override ;
        procedure ChargeDimArt ;
        procedure AfficheModele ;
        procedure OnChangeItemdim(Sender: TObject); 
     END ;

implementation

Procedure TOF_MBOPRMINMAX.OnArgument(Arguments : String ) ;
begin
inherited ;
GMASQUE:=THGRID(GetControl('GMASQUE'));
CodeSerie := '';
CodeMasque:= '';
end ;

Procedure TOF_MBOPRMINMAX.OnLoad  ;
var Q : TQuery;
BEGIN
inherited ;
OK_MAJ := True;
MsgEncours := False;
TobMethode:=TOB.Create('Methodes',Nil,-1) ;
TobListe:=TOB.CREATE('Masques',NIL,-1) ;
TobArtSerie:=TOB.CREATE('Article Serie',NIL,-1) ;

Q:=OpenSQL('SELECT GDM_MASQUE, GDM_LIBELLE FROM DIMMASQUE where GDM_TYPEMASQUE="'+VH_GC.BOTypeMasque_Defaut+'"',False) ;
if Not Q.EOF then TobListe.LoadDetailDB('MASQUES','','',Q,false);
GMASQUE:=THGrid(GetControl('GMASQUE')) ;
TobListe.PutGridDetail(GMASQUE,False,False,'GDM_MASQUE;GDM_LIBELLE',True) ;
Ferme(Q) ;
if TobListe.detail.count > 0 then AfficheModele;
END;

procedure TOF_MBOPRMINMAX.OnClose ;
begin
inherited ;
if DimensionsArticle<>nil then
  begin
  MajDim ;
  DimensionsArticle.free;
  end;
TobMethode.Free ;
TobMethode:=Nil ;
TobListe.free;
TobListe:=nil;
TobArtSerie.free;
TobArtSerie:=nil;
end;

procedure TOF_MBOPRMINMAX.MajDim ;
var TobT, TobA : TOB;
    i, y : integer;
    NbDim_modifie: integer;
    CodeDim1,CodeDim2,CodeDim3,CodeDim4,CodeDim5:string;
    GrilleDim1,GrilleDim2,GrilleDim3,GrilleDim4,GrilleDim5:string;
    Q : TQuery;
    QteMax, QteMin :Double ;
begin
//DimensionsArticle.ChangeChampDimMul (false);
if DimensionsArticle.TOBArticleDim <> nil then
  begin
  if DimensionsArticle.isModified=False then exit;

  for i:=0 to DimensionsArticle.TOBArticleDim.detail.count-1 do
    begin
    TobT:=DimensionsArticle.TOBArticleDim.detail[i] ;
    QteMin := TobT.GetValue('GAM_QTEDISPOMINI');
    QteMax := TobT.GetValue('GAM_QTEDISPOMAXI');
    if (QteMin <> 0) and (QteMax <>0) then
      begin
      if (QteMin > QteMax) then
        begin
        if MsgEncours = False then MessageAlerte('Présence de quantités MINI > quantités MAXI');
        OK_MAJ := False;
        for y :=0 to GMASQUE.RowCount-1 do
          begin
          if GMASQUE.CellValues [0,y] = TobT.GetValue('GAM_DIMMASQUE') then
            begin
            MsgEncours := True;
            GMASQUE.Row := y;
            MsgEncours := False;
            break;
            end;
          end;
        for y :=0 to GMETHODE.RowCount-1 do
          begin
          if GMETHODE.CellValues [0,y] = TobT.GetValue('GAM_CODESERIE') then
            begin
            MsgEncours := True;
            GMETHODE.Row := y;
            MsgEncours := False;
            break;
            end;
          end;
        exit;
        end;
      end;
    end;
  OK_MAJ := True;

  // Chargement de la table ARTICLESERIE en TOB
  Q:=OpenSQL('SELECT * from ARTICLESERIE where GAM_DIMMASQUE = "'+CodeMasque+'" and '+
             'GAM_CODESERIE ="'+CodeSerie+'"', False) ;
  if Not Q.EOF then TobArtSerie.LoadDetailDB('ARTICLESERIE','','',Q,True);
  Ferme(Q);

  NbDim_modifie:=0;
  for i:=0 to DimensionsArticle.TOBArticleDim.detail.count-1 do
    begin
    TobT:=DimensionsArticle.TOBArticleDim.detail[i] ;
    QteMin := TobT.GetValue('GAM_QTEDISPOMINI');
    QteMax := TobT.GetValue('GAM_QTEDISPOMAXI');
    if TobT.GetValue('GAM_CODEDIM1') = Null then
         CodeDim1:=''
    else CodeDim1:=TobT.GetValue('GAM_CODEDIM1');
    if TobT.GetValue('GAM_CODEDIM2') = Null then
         CodeDim2:=''
    else CodeDim2:=TobT.GetValue('GAM_CODEDIM2');
    if TobT.GetValue('GAM_CODEDIM3') = Null then
         CodeDim3:=''
    else CodeDim3:=TobT.GetValue('GAM_CODEDIM3');
    if TobT.GetValue('GAM_CODEDIM4') = Null then
         CodeDim4:=''
    else CodeDim4:=TobT.GetValue('GAM_CODEDIM4');
    if TobT.GetValue('GAM_CODEDIM5') = Null then
         CodeDim5:=''
    else CodeDim5:=TobT.GetValue('GAM_CODEDIM5');
    if TobT.GetValue('GAM_GRILLEDIM1') = Null then
         GrilleDim1:=''
    else GrilleDim1:=TobT.GetValue('GAM_GRILLEDIM1');
    if TobT.GetValue('GAM_GRILLEDIM2') = Null then
         GrilleDim2:=''
    else GrilleDim2:=TobT.GetValue('GAM_GRILLEDIM2');
    if TobT.GetValue('GAM_GRILLEDIM3') = Null then
         GrilleDim3:=''
    else GrilleDim3:=TobT.GetValue('GAM_GRILLEDIM3');
    if TobT.GetValue('GAM_GRILLEDIM4') = Null then
         GrilleDim4:=''
    else GrilleDim4:=TobT.GetValue('GAM_GRILLEDIM4');
    if TobT.GetValue('GAM_GRILLEDIM5') = Null then
         GrilleDim5:=''
    else GrilleDim5:=TobT.GetValue('GAM_GRILLEDIM5');

    TobA:=TobArtSerie.FindFirst(['GAM_CODEDIM1','GAM_CODEDIM2','GAM_CODEDIM3','GAM_CODEDIM4','GAM_CODEDIM5'],
                                [CodeDim1,CodeDim2,CodeDim3,CodeDim4,CodeDim5],false);
    if TobA=Nil then
      begin
      if (QteMin <> 0) or (QteMax <>0) then
        begin
        TobA:=TOB.Create('ARTICLESERIE',TobArtSerie,-1);
        TobA.PutValue('GAM_DIMMASQUE', CodeMasque);
        TobA.PutValue('GAM_CODESERIE', CodeSerie);
        TobA.PutValue('GAM_CODEDIM1', CodeDim1);
        TobA.PutValue('GAM_CODEDIM2', CodeDim2);
        TobA.PutValue('GAM_CODEDIM3', CodeDim3);
        TobA.PutValue('GAM_CODEDIM4', CodeDim4);
        TobA.PutValue('GAM_CODEDIM5', CodeDim5);
        TobA.PutValue('GAM_GRILLEDIM1', GrilleDim1);
        TobA.PutValue('GAM_GRILLEDIM2', GrilleDim2);
        TobA.PutValue('GAM_GRILLEDIM3', GrilleDim3);
        TobA.PutValue('GAM_GRILLEDIM4', GrilleDim4);
        TobA.PutValue('GAM_GRILLEDIM5', GrilleDim5);
        TobA.PutValue('GAM_QTEDISPOMINI', QteMin);
        TobA.PutValue('GAM_QTEDISPOMAXI', QteMax);
        inc (NbDim_modifie);
        end;
      end else
      begin
      if (QteMin<>TobA.GetValue('GAM_QTEDISPOMINI')) or (QteMax<>TobA.GetValue('GAM_QTEDISPOMAXI')) then
        begin
        TobA.PutValue('GAM_QTEDISPOMINI', QteMin);
        TobA.PutValue('GAM_QTEDISPOMAXI', QteMax);
        inc (NbDim_modifie);
        end;
      end;
    end;
  // Mise à jour de la table ARTICLESERIE
  if NbDim_modifie>0 then TobArtSerie.InsertOrUpdateDB(True);
  TobArtSerie.ClearDetail;
  end;
end;

procedure TOF_MBOPRMINMAX.AfficheModele;
var Q : TQuery;
begin
setcontroltext('LIBELLE_MASQUE',TobListe.Detail[GMASQUE.Row-1].GetValue('GDM_LIBELLE')) ;
Q:=OpenSQL('SELECT CC_CODE, CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE = "GAS"',False) ;
if Not Q.EOF then TobMethode.LoadDetailDB('MODELES','','',Q,false);
GMETHODE:=THGrid(GetControl('GMETHODE')) ;
TobMethode.PutGridDetail(GMETHODE,False,False,'CC_CODE;CC_LIBELLE',True) ;
Ferme(Q) ;
if TobMethode.detail.count > 0 then
  begin
  setcontroltext('LIBELLE_MODELE',TobMethode.Detail[GMETHODE.Row-1].GetValue('CC_LIBELLE')) ;
  ChargeDimArt;
  end;
end;


Procedure TOF_MBOPRMINMAX.ChargeDimArt;
begin
inherited ;
if DimensionsArticle<>nil then MajDim ;
if (TobMethode.detail.count > 0) and (OK_MAJ = True) then
   begin
   setcontroltext('LIBELLE_MODELE',TobMethode.Detail[GMETHODE.Row-1].GetValue('CC_LIBELLE')) ;
   setcontroltext('LIBELLE_MASQUE',TobListe.Detail[GMASQUE.Row-1].GetValue('GDM_LIBELLE')) ;
   CodeSerie :=TobMethode.Detail[GMETHODE.Row-1].GetValue('CC_CODE');
   CodeMasque:=TobListe.Detail[GMASQUE.Row-1].GetValue('GDM_MASQUE');
   InitDimensionsArticle ;
   end;
end;


procedure TOF_MBOPRMINMAX.InitDimensionsArticle ;
var iItem: integer ;
begin
if DimensionsArticle<>nil then DimensionsArticle.free;
DimensionsArticle:=TODimArticle.Create(THDimension(GetControl('DIMART'))
                     ,CodeSerie,CodeMasque
                     ,'','GCDIMCHAMP','PTR','', '', '-', False
                     ,'','','',True) ;
If DimensionsArticle.TOBArticleDim=nil then exit ;

for iItem:=0 to 3 do DimensionsArticle.Dim.PopUp.Items[iItem].Visible:=False ; // Aucun menu visible

DimensionsArticle.Dim.OnChange := OnChangeItemdim;

SetControlVisible('BPARAMDIM',False) ;
SetControlVisible('CBDETAIL',False) ;
OnChangeAfficheChampDimMul(DimensionsArticle,'PTR') ;
DimensionsArticle.TOBArticleDim.SetAllModifie(False);
end;

procedure TOF_MBOPRMINMAX.OnChangeItemdim(Sender: TObject);
var ItemDim :THDimensionItem ;
    QteMax, QteMin :Double ;
begin
ItemDim:=THDimensionItem(Sender) ;
if ItemDim = nil then exit;
QteMin:=valeur(ItemDim.Valeur[1]) ;
QteMax:=valeur(ItemDim.Valeur[2]) ;
if (QteMin <> 0) and (QteMax <> 0) then
  begin
  if (QteMax < QteMin) then
    begin
    MessageAlerte('La quantité maxi doit être supérieure à la quantité mini');
    DimensionsArticle.Dim.FocusDim(ItemDim,2);
    end;
  end;
end ;


procedure AGLOnAfficheModele (Parms: array of variant; nb: integer) ;
var F : TForm ;
     TOTOF : TOF ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then TOTOF:=TFVierge(F).LaTOF else exit ;
if (TOTOF is TOF_MBOPRMINMAX) then TOF_MBOPRMINMAX(TOTOF).ChargeDimArt;
end ;

procedure AGLOnMajDim (Parms: array of variant; nb: integer) ;
var F : TForm ;
     TOTOF : TOF ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then TOTOF:=TFVierge(F).LaTOF else exit ;
if (TOTOF is TOF_MBOPRMINMAX) then TOF_MBOPRMINMAX(TOTOF).MajDim;
end ;

Initialization
registerclasses([TOF_MBOPRMINMAX]) ;
RegisterAglProc('OnAfficheModele', True , 1, AGLOnAfficheModele) ;
RegisterAglProc('OnMajDim', True , 1, AGLOnMajDim) ;
end.

