unit IntegreInv;

interface            

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, hmsgbox, StdCtrls, UTOB, ImgList;

const
  WM_APPSTARTUP = WM_USER + 1;  

type
  TIntegreForm = class(TForm)
    ProgressBar: TProgressBar;
    lv_selection: TListView;
    bCancel: TButton;
    Label1: TLabel;
    HMsg: THMsgBox;
    iml_progress: TImageList;
    procedure lv_selectionCustomDrawItem(Sender: TCustomListView; Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure bCancelClick(Sender: TObject);
    procedure FormShow(Sender: TObject);

  private
    FTOBSelection : TOB;
    FBeingIntegreDepot,FBeingIntegreCodeTrans : String; // Ca c'est à cause du Transactions
    FCodeListeDefaut : String;
    FDateInv : TDateTime;
    FStockClos : Boolean;
    FValidating, FCanceled : boolean;

    procedure SetStatus(TL : TOB; ImageIdx : integer; Msg : String);
    procedure IntegreLaTrans;
    procedure AjoutArticleInv(TOBT, TOBL : TOB);
    procedure WMAppStartup(var msg: TMessage); message WM_APPSTARTUP;
  public
    procedure Init(TOBSelection : TOB);
    procedure ValidAll;
  end;

procedure IntegreLesInventaires(TOBSelection : TOB);

implementation

uses HCtrls, HEnt1,
{$IFDEF EAGLCLIENT}
{$ELSE}
     dbtables,
{$ENDIF}
     UTOFListeInv, InvUtil;


{$R *.DFM}

procedure IntegreLesInventaires(TOBSelection : TOB);
begin
with TIntegreForm.Create(Application) do
  begin
  Init(TOBSelection);
  Try
    ShowModal ;
  Finally
    Free ;
    end;
  end;
end;

procedure TIntegreForm.WMAppStartup(var msg: TMessage);
begin
  // Traitement
  ValidAll;
end ;

procedure TIntegreForm.FormShow(Sender: TObject);
begin
  PosTMessage(Handle, WM_APPSTARTUP, 0, 0);
end;

procedure TIntegreForm.Init(TOBSelection : TOB);
var i : integer;
begin
FTobSelection := TOBSelection;

FValidating := false;
FCanceled := false;

Label1.Caption := HMsg.Mess[0];
bCancel.Caption := HMsg.Mess[2];

lv_selection.Items.BeginUpdate;
for i := 0 to FTOBSelection.Detail.Count-1 do
  begin
  with lv_selection.Items.Add do
    begin
    Data := FTOBSelection.Detail[i];
    ImageIndex := 0;
    with FTOBSelection.Detail[i] do
      begin
      Caption := GetValue('GIT_CODETRANS');
      SubItems.Add(GetValue('GIT_LIBELLE'));
      SubItems.Add('');
      end;
    end;
  end;
lv_selection.Items.EndUpdate;
end;

procedure TIntegreForm.ValidAll;
var i : integer;
    TL : TOB;
    EventsLog : TStringList;
    JNalCode : String;
    TOBJnal : TOB ;
    QQ : TQuery;
    NumEvt : integer;
begin
EventsLog := TStringList.Create;

JNalCode := 'OK';
FValidating := true;
ProgressBar.Max := FTOBSelection.Detail.Count;
for i := 0 to FTOBSelection.Detail.Count-1 do
   begin
   TL := FTOBSelection.Detail[i];

   if FCanceled then
      begin
      SetStatus(TL, 0, HMsg.Mess[5]);
      PGIInfo(HMsg.Mess[6], Caption);
      EventsLog.Add(HMsg.Mess[7]);
      JNalCode := 'INT';
      break;
      end;

   SetStatus(TL, 1, HMsg.Mess[8]);
   Application.ProcessMessages;

   FBeingIntegreDepot := TL.GetValue('GIT_DEPOT');
   FBeingIntegreCodeTrans := TL.GetValue('GIT_CODETRANS');

   // Test si une liste d'inventaire a été générée pour ce dépôt
   QQ := OpenSQL('SELECT GIE_CODELISTE,GIE_DATEINVENTAIRE,GIE_STOCKCLOS FROM LISTEINVENT '+
                 'WHERE GIE_DEPOT="'+FBeingIntegreDepot+'" and GIE_VALIDATION<>"X"', true);
   if QQ.EOF then
      begin
      SetStatus(TL, 3, HMsg.Mess[11]);
      EventsLog.Add(FBeingIntegreCodeTrans+' : '+HMsg.Mess[9]+' '+HMsg.Mess[11]);
      JNalCode := 'ERR';
      Ferme(QQ);
      end else
      begin
      FCodeListeDefaut:=QQ.FindField('GIE_CODELISTE').AsString;
      FDateInv := QQ.FindField('GIE_DATEINVENTAIRE').AsDateTime;
      FStockClos:=((QQ.FindField('GIE_STOCKCLOS').AsString)='X');
      Ferme(QQ);

      case Transactions(IntegreLaTrans, 2) of
        oeSaisie :  begin SetStatus(TL, 3, HMsg.Mess[13]);
                          EventsLog.Add(FBeingIntegreCodeTrans+' : '+HMsg.Mess[9]+' '+HMsg.Mess[14]);
                          JNalCode := 'ERR'; end;
        oeUnknown : begin SetStatus(TL, 3, HMsg.Mess[15]);
                          EventsLog.Add(FBeingIntegreCodeTrans+' : '+HMsg.Mess[9]);
                          JNalCode := 'ERR'; end;
        oeOK :      begin SetStatus(TL, 2, HMsg.Mess[10]);
                          EventsLog.Add(FBeingIntegreCodeTrans+' : '+HMsg.Mess[10]); end;
        end;
      end;

   ProgressBar.StepIt;
   end;

// On force l'indicateur 'Article inventorié' à True sur l'ensemble des lignes de la liste.
// De ce fait, tous les articles non transmis auront une quantité inventoriée à zéro.
ExecuteSQL('UPDATE LISTEINVLIG SET GIL_SAISIINV="X" '+
           'WHERE EXISTS(SELECT GIE_CODELISTE FROM LISTEINVENT WHERE GIE_CODELISTE=GIL_CODELISTE '+
           'AND GIL_SAISIINV="-" AND GIL_DEPOT="'+FBeingIntegreDepot+'" AND GIE_VALIDATION<>"X")');

// Journal des evenements --------

NumEvt := 0;
TOBJnal := TOB.Create('JNALEVENT', nil, -1);
TOBJnal.PutValue('GEV_TYPEEVENT', 'INT');
TOBJnal.PutValue('GEV_LIBELLE', Caption);
TOBJnal.PutValue('GEV_DATEEVENT', Date);
TOBJnal.PutValue('GEV_UTILISATEUR', V_PGI.User);
QQ := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', true);
if not QQ.EOF then NumEvt := QQ.Fields[0].AsInteger;
Ferme(QQ);
inc(NumEvt);
TOBJnal.PutValue('GEV_NUMEVENT', NumEvt);
TOBJnal.PutValue('GEV_ETATEVENT', JNalCode);
TOBJnal.PutValue('GEV_BLOCNOTE', EventsLog.Text);
TOBJnal.InsertDB(nil);
TOBJnal.Free;

// ---------
FValidating := false;
if JNalCode='OK'
then Label1.Caption := HMsg.Mess[1]
else Label1.Caption := HMsg.Mess[4];
bCancel.Caption := HMsg.Mess[3];
bCancel.Cancel := false;
bCancel.Default := true;

EventsLog.Free;
end;


procedure TIntegreForm.SetStatus(TL : TOB; ImageIdx : integer; Msg : String);
begin
with lv_selection.FindData(0, TL, true, true) do
  begin
  ImageIndex := ImageIdx;
  SubItems[1] := Msg;
  MakeVisible(false);
  end;
end;

procedure TIntegreForm.IntegreLaTrans;
var Q : TQuery;
    TOBListeInv,TOBTransInv : TOB;
    TOBListeInvInsert,TOBListeInvUpdate : TOB;
    TOBListeInvInsertFille,TOBListeInvUpdateFille : TOB;
    TOBT, TOBL : TOB;
    i : integer;
    RequeteTransInvLig, RequeteTransInvLig2 : string;
begin
// Chargement de la liste des inventaires du dépôt en TOB
{ Modif 11/02/2003 permettant de regrouper les code barres
RequeteTransInvLig := 'SELECT * FROM TRANSINVLIG WHERE GIN_DEPOT="'+FBeingIntegreDepot+
                      '" and GIN_CODETRANS="'+FBeingIntegreCodeTrans+'"';}
// Ne gère pas les emplacements
// GIN_EMPLACEMENT,GIN_NUMLIGNE,GIN_CODETRANS
RequeteTransInvLig := 'SELECT GIN_ARTICLE,GIN_CODEARTICLE,GIN_CODEBARRE,GIN_DEPOT,'+
                      'SUM(GIN_QTEINV) AS GIN_QTEINV FROM TRANSINVLIG '+
                      'WHERE GIN_DEPOT="'+FBeingIntegreDepot+
                      '" and GIN_CODETRANS="'+FBeingIntegreCodeTrans+'" '+
                      'GROUP BY GIN_ARTICLE,GIN_CODEARTICLE,GIN_CODEBARRE,GIN_DEPOT';
RequeteTransInvLig2 :='SELECT DISTINCT GIN_ARTICLE FROM TRANSINVLIG '+
                      'WHERE GIN_DEPOT="'+FBeingIntegreDepot+
                      '" and GIN_CODETRANS="'+FBeingIntegreCodeTrans+'"';

TOBListeInv := TOB.Create('Liste Inv', nil, -1);
Q := OpenSQL('SELECT LISTEINVLIG.* FROM LISTEINVLIG '+
             'LEFT JOIN LISTEINVENT ON GIL_CODELISTE=GIE_CODELISTE '+
             'WHERE GIL_DEPOT="'+FBeingIntegreDepot+'" and GIE_VALIDATION<>"X" '+
             'AND EXISTS ('+RequeteTransInvLig2+' AND GIN_ARTICLE=GIL_ARTICLE)', true);
TOBListeInv.LoadDetailDB('LISTEINVLIG','','', Q, False);
Ferme(Q);

TOBListeInvUpdate := TOB.Create('Liste Inv', nil, -1);
TOBListeInvInsert := TOB.Create('Liste Inv', nil, -1);

// Chargement de l'inventaire transmis en TOB
TOBTransInv := TOB.Create('Trans Inv', nil, -1);
Q := OpenSQL(RequeteTransInvLig, true);
TOBTransInv.LoadDetailDB('TRANSINVLIG','','', Q, False);
Ferme(Q);

// Recherche du Code article, si celui-ci n'est pas renseigné sur certaines lignes de l'inventaire transmis
{ Cette recherche est désactivée car elle doit s'effectuer en amont
  c-à-d que les codes barres inconnus ne seront pas intégrés
TOBT:=TOBTransInv.FindFirst(['GIN_ARTICLE'],[''],False);
While TOBT<>Nil do
   begin
   Q:=OpenSQL('Select GA_ARTICLE, GA_CODEARTICLE from ARTICLE '+
              'where GA_CODEBARRE="'+TOBT.GetValue('GIN_CODEBARRE')+'"',True) ;
   if Not Q.EOF then
      begin
      TOBT.PutValue('GIN_ARTICLE',Q.FindField('GA_ARTICLE').AsString);
      TOBT.PutValue('GIN_CODEARTICLE',Q.FindField('GA_CODEARTICLE').AsString);
      end;
   Ferme(Q);
   TOBT:=TOBTransInv.FindNext(['GIN_ARTICLE'],[''],False);
   end; }

try
  for i := 0 to TOBTransInv.Detail.Count-1 do
    begin
    TOBT := TOBTransInv.Detail[i];
    if TOBT.GetValue('GIN_ARTICLE')<>'' then
       begin
       TOBL := TOBListeInv.FindFirst(['GIL_ARTICLE'],[TOBT.GetValue('GIN_ARTICLE')],False);
       if TOBL<>nil then
          begin
          TOBL.PutValue('GIL_INVENTAIRE', TOBL.GetValue('GIL_INVENTAIRE')+TOBT.GetValue('GIN_QTEINV'));
          //if TOBT.GetValue('GIN_EMPLACEMENT')<>'' then
          //   TOBL.PutValue('GIL_EMPLACEMENT', TOBT.GetValue('GIN_EMPLACEMENT'));
          TOBL.PutValue('GIL_SAISIINV', 'X');
          TOBListeInvUpdateFille := TOB.Create('LISTEINVLIG', TOBListeInvUpdate, -1);
          TOBListeInvUpdateFille.Dupliquer(TOBL,False,True,True); //MODIF 29/01/2003
          end else
          begin
          TOBL := TOB.Create('LISTEINVLIG', TOBListeInv, -1);
          AjoutArticleInv(TOBT, TOBL);
          TOBListeInvInsertFille := TOB.Create('LISTEINVLIG', TOBListeInvInsert, -1);
          TOBListeInvInsertFille.Dupliquer(TOBL,False,True); //MODIF 29/01/2003
          end;
       end;
    end;
  TOBListeInvUpdate.UpdateDB(True);
  TOBListeInvInsert.InsertDB(Nil,True);
  // Flaggage de l'inventaire intégré
  ExecuteSQL('UPDATE TRANSINVENT SET GIT_INTEGRATION="X" WHERE GIT_DEPOT="'+FBeingIntegreDepot+'" and GIT_CODETRANS="'+FBeingIntegreCodeTrans+'"');
  Finally
   TOBListeInv.Free;
   TOBTransInv.Free;
   TOBListeInvUpdate.Free;
   TOBListeInvInsert.Free;
   end;
end;

procedure TIntegreForm.AjoutArticleInv(TOBT, TOBL : TOB);
var CArticle,Cdepot : string;
    CEmplacement,StSelectDepot : string;
    QA, QD : TQuery;
    FUS,FUV,FPPQ : Double;
    GereParLot : Boolean;
    StocPhy, DPAD, PMAPD, DPRD, PMRPD, DPAA, PMAPA, DPRA, PMRPA : Double;
    QPR : TQtePrixRec;
begin
CArticle := TOBT.GetValue('GIN_ARTICLE');
CDepot := TOBT.GetValue('GIN_DEPOT');
QA := OpenSQL('SELECT GA_ARTICLE, GA_LOT, GA_DPA, GA_PMAP, GA_DPR, GA_PMRP, '+
              'GA_QUALIFUNITESTO, GA_QUALIFUNITEVTE, GA_PRIXPOURQTE '+
              'FROM ARTICLE WHERE GA_ARTICLE="'+CArticle+'"',true);
StSelectDepot:='SELECT GQ_ARTICLE,GQ_DEPOT,GQ_CLOTURE,GQ_DATECLOTURE,GQ_DPA,GQ_DPR,GQ_PMAP,'+
               'GQ_PMRP,GQ_PHYSIQUE,GQ_TRANSFERT,'+
               'GQ_RESERVECLI,GQ_PREPACLI,GQ_EMPLACEMENT FROM DISPO '+
               'WHERE GQ_ARTICLE="'+CArticle+'" AND GQ_DEPOT="'+CDepot+'" AND GQ_CLOTURE="-"';
QD := OpenSQL(StSelectDepot,true);
GereParLot := (QA.FindField('GA_LOT').AsString = 'X');

If FStockClos=True then
   begin
   QPR := GetQtePrixDateListe(CArticle, CDepot, FDateInv);
   if not QPR.SomethingReturned then StocPhy := QD.FindField('GQ_PHYSIQUE').AsFloat
                                else StocPhy := QPR.Qte;
   end else
   begin
   if QD.EOF then StocPhy := 0
             else StocPhy := QD.FindField('GQ_PHYSIQUE').AsFloat;
   end;

FUS := RatioMesure('PIE', QA.FindField('GA_QUALIFUNITESTO').AsString);
FUV := RatioMesure('PIE', QA.FindField('GA_QUALIFUNITEVTE').AsString);
FPPQ := QA.FindField('GA_PRIXPOURQTE').AsFloat; if FPPQ = 0 then FPPQ := 1.0;

DPAA := Ratioize(QA.FindField('GA_DPA').AsFloat, FUS, FUV, FPPQ);
PMAPA := Ratioize(QA.FindField('GA_PMAP').AsFloat, FUS, FUV, FPPQ);
DPRA := Ratioize(QA.FindField('GA_DPR').AsFloat, FUS, FUV, FPPQ);
PMRPA := Ratioize(QA.FindField('GA_PMRP').AsFloat, FUS, FUV, FPPQ);
if QD.EOF then
  begin
  DPAD := DPAA;
  PMAPD := PMAPA;
  DPRD := DPRA;
  PMRPD := PMRPA;
  CEmplacement := '';
  end else
  begin
  DPAD := Ratioize(QD.FindField('GQ_DPA').AsFloat, FUS, FUV, FPPQ);
  PMAPD := Ratioize(QD.FindField('GQ_PMAP').AsFloat, FUS, FUV, FPPQ);
  DPRD := Ratioize(QD.FindField('GQ_DPR').AsFloat, FUS, FUV, FPPQ);
  PMRPD := Ratioize(QD.FindField('GQ_PMRP').AsFloat, FUS, FUV, FPPQ);
  CEmplacement := QD.FindField('GQ_EMPLACEMENT').AsString;
  end;

TOBL.PutValue('GIL_CODELISTE', FCodeListeDefaut);
TOBL.PutValue('GIL_ARTICLE', CArticle);
TOBL.PutValue('GIL_CODEARTICLE', TOBT.GetValue('GIN_CODEARTICLE'));
if GereParLot then TOBL.PutValue('GIL_LOT', 'X') else TOBL.PutValue('GIL_LOT', '-');
TOBL.PutValue('GIL_DEPOT', CDepot);
TOBL.PutValue('GIL_EMPLACEMENT', CEmplacement);
TOBL.PutValue('GIL_SAISIINV', 'X');
TOBL.PutValue('GIL_INVENTAIRE', TOBT.GetValue('GIN_QTEINV'));
TOBL.PutValue('GIL_QTEPHOTOINV', StocPhy);
TOBL.PutValue('GIL_DPA', DPAD);
TOBL.PutValue('GIL_PMAP', PMAPD);
TOBL.PutValue('GIL_DPR', DPRD);
TOBL.PutValue('GIL_PMRP', PMRPD);
TOBL.PutValue('GIL_DPAART', DPAA);
TOBL.PutValue('GIL_PMAPART', PMAPA);
TOBL.PutValue('GIL_DPRART', DPRA);
TOBL.PutValue('GIL_PMRPART', PMRPA);
TOBL.PutValue('GIL_DPASAIS', 0.0);
TOBL.PutValue('GIL_PMAPSAIS', 0.0);
TOBL.PutValue('GIL_DPRSAIS', 0.0);
TOBL.PutValue('GIL_PMRPSAIS', 0.0);

Ferme(QA);
Ferme(QD);
end;

procedure TIntegreForm.lv_selectionCustomDrawItem(Sender: TCustomListView; Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
if Item.ImageIndex = 1 then Sender.Canvas.Font.Style := [fsBold]
                       else Sender.Canvas.Font.Style := [];
end;

procedure TIntegreForm.bCancelClick(Sender: TObject);
begin
if FValidating then FCanceled := true
               else Close;
end;


end.
