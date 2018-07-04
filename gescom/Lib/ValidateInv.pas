{***********UNITE*************************************************
Auteur  ...... : Julien D
Créé le ...... : 16/05/2002
Modifié le ... : 16/05/2002
Description .. : Validation d'un inventaire
Mots clefs ... : INVENTAIRE
*****************************************************************}
unit ValidateInv;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, HStatus,
  ComCtrls, hmsgbox, StdCtrls, UTOB, ImgList, ent1, entgc, SaisUtil, Paramsoc, FactUtil,
  UtilArticle, StockUtil, FactPiece, FactTOB, FactArticle,
  FactLotSerie, Grids, Hctrls, HImgList,uEntCommun, TntGrids ;

type
  TValidateForm = class(TForm)
    ProgressBar: TProgressBar;
    bCancel: TButton;
    Label1: TLabel;
    HMsg: THMsgBox;
    iml_progress: THImageList;
    GridSelection: THGrid;

//    procedure lv_selectionCustomDrawItem(Sender: TCustomListView; Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure bCancelClick(Sender: TObject);

  private
  	fNomTable : string;
    FTOBSelection : TOB;
    Tob_Liste         : TOB ;        // TOB de la liste d'inventaire
    Tob_ListeLot      : TOB ;        // TOB de la liste d'inventaire des lots
    TOBDesLots        : TOB ;        // TOB noeud des lots
    Tob_Tiers         : TOB ;        // TOB du tiers
    Tob_Piece         : TOB ;        // TOB entête de pièce
    Tob_Ligne         : TOB ;        // TOB lignes
    Tob_Article       : TOB ;        // TOB des articles de la pièce
    Tob_PiedBase      : TOB ;        // TOB PiedBase
    Tob_LigneBase      : TOB ;        // TOB PiedBase
    Tob_Acomptes      : TOB ;        // TOB Acomptes
    Tob_Catalogue     : TOB ;        // TOB Catalogues
    Tob_Noment        : TOB ;        // TOB Nomenclatures
    CleDoc : R_CleDoc ;
    DEV    : RDEVISE ;
    CodeTiers : string;
    QualifMvt  : string ;            // qualifiant mouvement de la nature de pièce
    EnHT : boolean;
    ListeValide : boolean;
    FPrix : array[1..3] of String;
    FBeingValidatedCodeListe : String; // Ca c'est à cause du Transactions
    FValidating, FCanceled : boolean;
    FRemiseStockZero : boolean;
    GereLot : boolean; //Le parpièce des inventaires
    fRecalcPmap : Boolean;
    Depot : String;
    procedure SetStatus(TL : TOB; ImageIdx : integer; Msg : String);
    function  ChoozePrice(BaseChamp : String; TOBInvLig : TOB) : Double;
    procedure ValideLaListe;
    function  ConstruireDocInv: boolean;
    function  ConstruirePiece(NaturePiece : string; DateInv : TDate): boolean;
    function  ChargeTOBClient : Boolean;
    function  ConstruireLignePiece (CodeArtGen,CodeArtDim:String; TOBListe:TOB; QteEcart:double; IndiceArticle:integer; IsGenerique:boolean) : boolean;
    procedure ConstruireLigneLot;
    procedure BuildTob_TobArticle(Tob_Liste : TOB);
    procedure MAJQtePrixLignesGEN();
    function  ValideLesPieces : boolean;
    function  RetourneLibelleAvecDimensions(RefArtDim:String;TOBArt:TOB): string;
    Function  QuelPrixBaseInv () : double ;
    procedure SupprimeEcartZero ();
    procedure ValideLesLots(TOBDesLots, TOBArticles, TOBPiece : TOB; VenteAchat : String) ;
  public
    procedure Init(TOBSelection: TOB; Prix1, Prix2, Prix3: String; RemiseStockZero, RecalcPmap: Boolean);
    procedure ValidAll;

  end;

const
// libellés des messages
TexteMessage: array[1..6] of string 	= (
          {1}  'Impossible de déterminer le client par défaut.'
          {2} ,'Erreur à la création de la pièce'
          {3} ,'Erreur à la création d''une ligne'
          {4} ,'Impossible de trouver cet article :'
          {5} ,'Erreur à la validation de l''inventaire'
          {6} ,'Impossible de créer une pièce sans dépôt'
              );
NbArtParRequete : integer = 1000;

{ Images état de la validation }
IMG_Error  = 2;
IMG_Ok     = 12;
IMG_Exec   = 49;
IMG_Cancel = 5;

GRID_Cols = 'IMG;GIE_CODELISTE;GIE_LIBELLE;MSG';


procedure ValideLesListes(Sender : TWinControl; TOBSelection : TOB; Prix1, Prix2, Prix3 : String; RemiseStockZero,RecalcPmap : Boolean);

implementation

uses HEnt1,
{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     UtilPMAPCalcul,
     FactCalc,
     FactComm;

Const NaturePiece = 'INV';
      { StSelectArticle : String = 'SELECT * FROM ARTICLE'; }
      StSelectArticle : String = 'SELECT GA_ARTICLE,GA_CODEARTICLE,GA_LIBELLE,GA_STATUTART,'+
        'GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,GA_GRILLEDIM4,GA_GRILLEDIM5,GA_CODEDIM1,'+
        'GA_CODEDIM2,GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5,GA_PRIXPOURQTE,GA_ESCOMPTABLE,'+
        'GA_REMISEPIED,GA_REMISELIGNE,GA_TENUESTOCK,GA_TARIFARTICLE,GA_QUALIFSURFACE,'+
        'GA_QUALIFVOLUME,GA_QUALIFPOIDS,GA_QUALIFLINEAIRE,GA_QUALIFHEURE,GA_SURFACE,'+
        'GA_VOLUME,GA_POIDSBRUT,GA_POIDSNET,GA_POIDSDOUA,GA_LINEAIRE,GA_HEURE,'+
        'GA_QUALIFUNITESTO,GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3,GA_LIBREART1,'+
        'GA_LIBREART2,GA_LIBREART3,GA_LIBREART4,GA_LIBREART5,GA_LIBREART6,GA_LIBREART7,'+
        'GA_LIBREART8,GA_LIBREART9,GA_LIBREARTA,GA_FAMILLETAXE1,GA_FAMILLETAXE2,'+
        'GA_FAMILLETAXE3,GA_FAMILLETAXE4,GA_FAMILLETAXE5,GA_FOURNPRINC,GA_LOT FROM ARTICLE';

{$R *.DFM}

procedure ValideLesListes(Sender : TWinControl; TOBSelection : TOB; Prix1, Prix2, Prix3 : String; RemiseStockZero,RecalcPmap : Boolean);
begin
  with TValidateForm.Create(Sender) do
  begin
  Init(TOBSelection, Prix1, Prix2, Prix3, RemiseStockZero,RecalcPmap);
  Sender.Enabled := false;
  Show; Update;
  ValidAll;
    while Visible do Application.ProcessMessages;
  Sender.Enabled := true;
  Release;
  end;
end;


(*procedure TValidateForm.Init(TOBSelection : TOB; Prix1, Prix2, Prix3 : String; RemiseStockZero : Boolean);
var i : integer;
begin
FTobSelection := TOBSelection;
FPrix[1] := Prix1;
FPrix[2] := Prix2;
FPrix[3] := Prix3;
FRemiseStockZero := RemiseStockZero;

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
      Caption := GetValue('GIE_CODELISTE');
      SubItems.Add(GetValue('GIE_LIBELLE'));
      SubItems.Add('');
      end;
    end;
  end;
lv_selection.Items.EndUpdate;
end;*)

procedure TValidateForm.Init(TOBSelection : TOB; Prix1, Prix2, Prix3 : String; RemiseStockZero,RecalcPmap : Boolean);
begin
  FTobSelection := TOBSelection;
  FPrix[1] := Prix1;
  FPrix[2] := Prix2;
  FPrix[3] := Prix3;
  FRemiseStockZero := RemiseStockZero;
  fRecalcPmap := RecalcPmap;

  FValidating := false;
  FCanceled := false;

//  PnTop.Caption := HMsg.Mess[0];
//  bCancel.Caption := HMsg.Mess[2];

  {$IFDEF STK }
    Emplacement   := StkGereEmplacement;
    Lot           := StkGereLot;
    Serie         := StkGereSerie;
  {$ENDIF STK }

  if FTOBSelection.Detail.Count > 0 then
  begin
    FTOBSelection.Detail[0].AddChampSupValeur('IMG', '#ICO#' + IntToStr(IMG_Cancel), True);
    FTOBSelection.Detail[0].AddChampSupValeur('MSG', ''       , True);
  end;
  FTOBSelection.PutGridDetail(GridSelection, False, False, GRID_Cols);
  GridSelection.Row := GridSelection.FixedRows
end;


procedure TValidateForm.ValidAll;
var i : integer;
    TL : TOB;
    EventsLog : TStringList;
    JNalCode : String;
    TOBJnal : TOB ;
    QQ : TQuery;
    NumEvt : integer;
begin
fNomTable := ConstitueNomTemp;

EventsLog := TStringList.Create;
EventsLog.Add(HMsg.Mess[4]);
for i := 1 to 3 do EventsLog.Add('  '+inttostr(i)+'. '+RechDom('GCTYPEPRIX', FPrix[i], false));
EventsLog.Add('');

JNalCode := 'OK';
FValidating := true;
ProgressBar.Max := FTOBSelection.Detail.Count;
GereLot := (GetInfoParPiece(NaturePiece,'GPP_LOT')='X') ;

for i := 0 to FTOBSelection.Detail.Count-1 do
  begin
  TL := FTOBSelection.Detail[i];

  if FCanceled then
    begin
      SetStatus(TL, IMG_Cancel, HMsg.Mess[5]);
    PGIInfo(HMsg.Mess[6], Caption);
    EventsLog.Add(HMsg.Mess[7]);
    JNalCode := 'INT';
    break;
    end;

    SetStatus(TL, IMG_Exec, HMsg.Mess[8]);
  Application.ProcessMessages;

  FBeingValidatedCodeListe := TL.GetValue('GIE_CODELISTE');

  if ExisteSQL('SELECT GIL_CODELISTE FROM LISTEINVLIG WHERE GIL_CODELISTE="'+FBeingValidatedCodeListe+'" AND GIL_SAISIINV<>"X"') then
    begin
      SetStatus(TL, IMG_Error, HMsg.Mess[11]);
    EventsLog.Add(FBeingValidatedCodeListe+' : '+HMsg.Mess[9]+' '+HMsg.Mess[12]);
    JNalCode := 'ERR';
    end
   else if GetParamsoc('SO_GCTIERSINV')='' then
    begin
      SetStatus(TL, IMG_Error, HMsg.Mess[17]);
    EventsLog.Add(FBeingValidatedCodeListe+' : '+HMsg.Mess[9]+' '+HMsg.Mess[17]);
    JNalCode := 'ERR';
    end
   else
    case Transactions(ValideLaListe, 0) of
      oeSaisie :
        begin
          SetStatus(TL, IMG_Error, HMsg.Mess[13]);
          EventsLog.Add(FBeingValidatedCodeListe+' : '+HMsg.Mess[9]+' '+HMsg.Mess[14]);
          JNalCode := 'ERR'; end;
      oeUnknown :
        begin
          SetStatus(TL, IMG_Error, HMsg.Mess[15]);
          EventsLog.Add(FBeingValidatedCodeListe+' : '+HMsg.Mess[9]);
          JNalCode := 'ERR'; end;
      oeOK :
        begin if ListeValide=True then
          begin
            SetStatus(TL, IMG_Ok, HMsg.Mess[10]);
            EventsLog.Add(FBeingValidatedCodeListe+' : '+HMsg.Mess[10]);
          end else
          begin
            SetStatus(TL, IMG_Error, HMsg.Mess[15]);
            EventsLog.Add(FBeingValidatedCodeListe+' : '+HMsg.Mess[9]);
            JNalCode := 'ERR';
          end;
        end;
    end;

  ProgressBar.StepIt;
  end;

// Journal des evenements --------

NumEvt := 0;
TOBJnal := TOB.Create('JNALEVENT', nil, -1);
TOBJnal.PutValue('GEV_TYPEEVENT', 'INV');
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
else Label1.Caption := HMsg.Mess[16];
bCancel.Caption := HMsg.Mess[3];
bCancel.Cancel := false;
bCancel.Default := true;

EventsLog.Free;
end;


(*procedure TValidateForm.SetStatus(TL : TOB; ImageIdx : integer; Msg : String);
begin
with lv_selection.FindData(0, TL, true, true) do
  begin
  ImageIndex := ImageIdx;
  SubItems[1] := Msg;
  MakeVisible(false);
  end;
end;*)

procedure TValidateForm.SetStatus(TL : TOB; ImageIdx : integer; Msg : String);
begin
  TL.SetString('MSG', Msg);
  TL.SetString('IMG', '#ICO#' + IntToStr(ImageIdx));
  TL.PutLigneGrid(GridSelection, TL.GetIndex + GridSelection.FixedRows, False, False, GRID_Cols);
  GridSelection.SetFocus;
  GridSelection.Row := TL.GetIndex + GridSelection.FixedRows;
  if GridSelection.Row > GridSelection.VisibleRowCount then
    GridSelection.ScrollBy(0, GridSelection.DefaultRowHeight);
  GridSelection.Repaint;
  Application.ProcessMessages
end;


function TValidateForm.ChoozePrice(BaseChamp : String; TOBInvLig : TOB) : Double;
var i : integer;
begin
result := 0;
for i := 1 to 3 do
 if result = 0 then
   result := TOBInvLig.GetValue(BaseChamp + RechDom('GCTYPEPRIX', FPrix[i], true));
end;


procedure TValidateForm.ValideLaListe;
begin

  ListeValide:=False;

  try
    Tob_Liste := TOB.Create('Liste inv', nil, -1);
    //
    if GereLot then
    begin
      Tob_ListeLot := TOB.Create('Liste inv lot', nil, -1);
      TOBDesLots := TOB.Create('Noeud des Lots',Nil,-1) ;
    end;
    //
    Tob_Piece := TOB.Create('PIECE', nil, -1);
    Tob_Tiers := TOB.CREATE ('TIERS', nil, -1);
    /////////////////////////////////////////////////////
    // Création des TOB Virtuelles rattachées à la pièce
    /////////////////////////////////////////////////////
    Tob_Article   := TOB.CREATE ('Les Articles', nil, -1);
    Tob_PiedBase  := TOB.CREATE ('Les Taxes' , nil, -1);
    Tob_LigneBase := TOB.Create ('LES TAXES L',nil,-1);
    Tob_Acomptes  := TOB.CREATE ('Les Acomptes', nil, -1);
    Tob_Catalogue := TOB.CREATE ('Les Catalogues', nil, -1);
    Tob_noment    := TOB.CREATE ('Les Noments', nil, -1);

    if ConstruireDocInv=True then
      begin
      // Flaggage de la liste validée
      if ExecuteSQL('UPDATE LISTEINVENT SET GIE_VALIDATION="X" WHERE GIE_CODELISTE="'+FBeingValidatedCodeListe+'"') = 0 then
      V_PGI.ioerror := oeUnknown;

      //
      ListeValide:=True;
      //
      // Cloture les stocks à la date de la validation de l'inventaire seulement
      // pour les articles contenus dans la liste d'inventaire
      {ExecuteSQL('INSERT DISPO '+
         'select GQ_ARTICLE,GQ_DEPOT,GQ_STOCKMIN,GQ_STOCKMAX,"'+UsTime(NowH)+'",'+
         'GQ_PHYSIQUE,GQ_RESERVECLI,GQ_RESERVEFOU,GQ_PREPACLI,GQ_LIVRECLIENT,'+
         'GQ_LIVREFOU,GQ_TRANSFERT,GQ_QTE1,GQ_QTE2,GQ_QTE3,GQ_DPA,GQ_PMAP,'+
         'GQ_DPR,GQ_PMRP,GQ_DPACON,GQ_PMAPCON,GQ_DPRCON,GQ_PMRPCON,'+
         'GQ_DATECREATION,GQ_DATEMODIF,GQ_DATEINTEGR,GQ_UTILISATEUR,'+
         'GQ_STOCKINITIAL,"X","'+UsTime(NowH)+
         '",GQ_EMPLACEMENT,GQ_CUMULSORTIES,GQ_CUMULENTREES,'+
         'GQ_VENTEFFO,GQ_ENTREESORTIES,GQ_ECARTINV from DISPO '+
         'where gq_depot="'+Depot+'" and gq_cloture="-" AND Exists'+
         '(Select GIL_ARTICLE From LISTEINVLIG Where '+
         'GIL_CODELISTE="'+FBeingValidatedCodeListe+'" '+
         'AND GIL_ARTICLE=GQ_ARTICLE)');}
      end;
  finally
    Tob_Tiers.Free;
    Tob_Piece.Free;
    Tob_Liste.Free;
    if GereLot then
    begin
      Tob_ListeLot.Free;
      TOBDesLots.Free;
    end;
    Tob_Article.Free;
    Tob_PiedBase.Free;
    Tob_LigneBase.free;
    Tob_Acomptes.Free;
    Tob_Catalogue.Free;
    Tob_noment.Free;
    end;

end;

procedure TValidateForm.SupprimeEcartZero ();
var Index : integer;
    EcartStock , QteInvStock : Double;
    EcartLot , QteInvLot : Double;
    OkSuppLigne : Boolean;
    TobInvLot : TOB;
begin
  For Index := Tob_Liste.Detail.Count-1 downto 0 do
  begin
    QteInvStock := Tob_Liste.Detail[Index].GetValue('GIL_INVENTAIRE');
    EcartStock := QteInvStock - Tob_Liste.Detail[Index].GetValue('GIL_QTEPHOTOINV');

    if EcartStock = 0 then //Suppresion de la ligne si pas d'ecart dans les lots
    begin
      OkSuppLigne := True;
      if GereLot then
      begin
        TobInvLot := Tob_ListeLot.FindFirst(['GLI_ARTICLE'],[Tob_Liste.Detail[Index].GetValue('GIL_ARTICLE')],False) ;
        while TobInvLot <> Nil do
        begin
          QteInvLot := TobInvLot.GetValue('GLI_INVENTAIRE');
          EcartLot := QteInvLot - TobInvLot.GetValue('GLI_QTEPHOTOINV');
          if EcartLot = 0 then
          begin
            TobInvLot.Free; //Suppression de la ligne correspondant au lot
            //TobInvLot := Tob_ListeLot.FindFirst(['GLI_ARTICLE'],[Tob_Liste.Detail[Index].GetValue('GIL_ARTICLE')],False) ;
          end
          else
          begin
            OkSuppLigne := False;
          end;
          TobInvLot := Tob_ListeLot.FindNext(['GLI_ARTICLE'],[Tob_Liste.Detail[Index].GetValue('GIL_ARTICLE')],False) ;
        end;
      end;
      if OkSuppLigne then Tob_Liste.Detail[Index].Free;
    end;
  end;
end;

function TValidateForm.ConstruireDocInv: boolean;
var Q : TQuery;
    TL : TOB;
    CodeArtGen,CodeArtDim,CodeArtGPrec : String;
    DateInv : TDate;
    //Ecart;
    QteInv : Double;
    i,IndiceArticle : integer;
begin
  result:=false; IndiceArticle:=-1;

  Q := OpenSQL('SELECT GIE_DEPOT, GIE_DATEINVENTAIRE FROM LISTEINVENT WHERE GIE_CODELISTE="'+FBeingValidatedCodeListe+'"', true);
  if Q.EOF then
  begin
  	Ferme(Q);
    exit;
  end;
  Depot := Q.FindField('GIE_DEPOT').AsString;
  if Depot = '' then
  begin
  	PGIInfo(TraduireMemoire(TexteMessage[6]), Caption);
    exit;
  end;
  DateInv := Q.FindField('GIE_DATEINVENTAIRE').AsDateTime;
  Ferme(Q);

  if ConstruirePiece(NaturePiece, DateInv)=False then
  begin
  	PGIInfo(TraduireMemoire(TexteMessage[2]), Caption);
    exit;
  end;

  QualifMvt:=GetInfoParPiece(NaturePiece,'GPP_QUALIFMVT');
  EnHT:=(Tob_Piece.GetValue('GP_FACTUREHT')='X');

  // Chargement des lignes la liste d'inventaire en TOB
  Q := OpenSQL('SELECT * FROM LISTEINVLIG '+
  'WHERE GIL_CODELISTE="'+FBeingValidatedCodeListe+'" '+
  'ORDER BY GIL_ARTICLE', true);
  Tob_Liste.LoadDetailDB('LISTEINVLIG','','', Q, false,True);
  Ferme(Q);

  if GereLot then
  begin
    // Chargement des lignes la liste d'inventaire en TOB
    Q := OpenSQL('SELECT * FROM LISTEINVLOT '+
                  'WHERE GLI_CODELISTE="'+FBeingValidatedCodeListe+'" '+
                  'ORDER BY GLI_ARTICLE', true);
    Tob_ListeLot.LoadDetailDB('LISTEINVLOT','','', Q, false,True);
    Ferme(Q);
  end;

  // ---------------------------------------------------------------------
  // Génération d'une ligne sur la pièce, à chaque fois qu'il y a un écart
  // entre le stock théorique et le stock inventorié.
  // ---------------------------------------------------------------------
  CodeArtGPrec := '';

  //  SupprimeEcartZero; Modifie LS le 12/12/2009 pour mise a jour de la valo du stock

  // Chargement des différents Articles dans la liste Tob_Article
  if Tob_Liste.Detail.count>0 then BuildTob_TobArticle(Tob_Liste);

  INITMOVE(Tob_Liste.Detail.Count,'');
  for i := 0 to Tob_Liste.Detail.Count-1 do
  begin
    MoveCur(False);
    TL := Tob_Liste.Detail[i];
    QteInv := TL.GetValue('GIL_INVENTAIRE');
    //Ecart := QteInv - TL.GetValue('GIL_QTEPHOTOINV');
    //if (Ecart=0) or ((QteInv=0) and (FRemiseStockZero=False)) then continue;
    if (QteInv=0) and (FRemiseStockZero=False) then continue;

    IndiceArticle:=IndiceArticle+1;
    CodeArtGen := TL.GetValue('GIL_CODEARTICLE');

    if CodeArtGen<>CodeArtGPrec then
    begin
      // Ajout d'une ligne Commentaire sur l'article Générique, pouvant regrouper un ou plusieurs articles dimensionnés
      if Tob_Article.Detail[IndiceArticle].GetValue('GA_STATUTART')='DIM' then
      begin
        CodeArtDim := CodeArticleUnique(CodeArtGen,'','','','','');
        if Not ConstruireLignePiece(CodeArtGen, CodeArtDim, TL, 0,IndiceArticle,True) then
        begin
          PGIInfo(TraduireMemoire(TexteMessage[3]),Caption);
          exit;
        end;
      end;
      CodeArtGPrec:=CodeArtGen;
    end;

    CodeArtDim := TL.GetValue('GIL_ARTICLE');
    if Not ConstruireLignePiece(CodeArtGen, CodeArtDim, TL, QteInv, IndiceArticle,False) then
    begin
      PGIInfo(TraduireMemoire(TexteMessage[3]),Caption);
      exit;
    end;
    if GereLot then ConstruireLigneLot;
    //
    //
    //FV1 - 28/04/2014 - Mise à jour des Dispos
    TL.AddChampSupValeur('DATEINVENTAIRE', DateToStr(DateInv));
    MiseAjourDispoInv(TL,FPrix,fRecalcPmap);
    //
  end;

  FINIMOVE;

  Tob_Liste.ClearDetail;

  if Tob_Piece.Detail.Count>0 then
  begin
    NumeroteLignesGC(Nil,TOB_Piece) ;
    PutValueDetail(Tob_Piece,'GP_RECALCULER','X');
    MAJQtePrixLignesGEN();
    IncNumSoucheG(CleDoc.Souche,DateInv);
    CalculFacture(nil, Tob_Piece, nil,nil,nil,nil,Tob_PiedBase,Tob_LigneBase, Tob_Tiers, Tob_Article, nil,nil,nil,nil, DEV) ;
    if not ValideLesPieces() then PGIInfo(TraduireMemoire(TexteMessage[5]),Caption)
                             else result:=true;
  end
  else
    result:=true;

end;

function TValidateForm.ConstruirePiece(NaturePiece : string; DateInv : TDate): boolean;
var DD : TDateTime;
    Per,Sem : integer;
begin
  result:=False;
  Tob_Piece.InitValeurs;
  AddLesSupEntete(Tob_Piece);
  InitTOBPiece(Tob_Piece);    // Init date de création, modif, utilisateur .....
  Tob_Piece.PutValue('GP_NATUREPIECEG', NaturePiece);
  Tob_Piece.PutValue('GP_DATEPIECE', DateInv);
  /////////////////////////
  // Souche, Numéro, .....
  ////////////////////////
  CleDoc.Souche      := GetSoucheG(NaturePiece, Depot,'') ;
  CleDoc.NumeroPiece := GetNumSoucheG(CleDoc.Souche,DateInv) ;
  CleDoc.DatePiece   := Tob_Piece.GetValue('GP_DATEPIECE');

  PutValueDetail(Tob_Piece, 'GP_SOUCHE', CleDoc.Souche);
  PutValueDetail(Tob_Piece, 'GP_NUMERO', CleDoc.NumeroPiece);
  PutValueDetail(Tob_Piece, 'GP_INDICEG', 0);
  /////////////////////////////////////////////////////////////////////////
  // Etablissement / Dépôt
  /////////////////////////////////////////////////////////////////////////
  Tob_Piece.PutValue('GP_ETABLISSEMENT', Depot);
  Tob_Piece.PutValue('GP_DEPOT', Depot);
  ///////////////////////////////////////////////////////////////
  // Chargement en TOB du client
  ///////////////////////////////////////////////////////////////
  if Not ChargeTOBClient then exit;
  Tob_Piece.PutValue ('GP_TIERS', CodeTiers);
  Tob_Piece.PutValue ('GP_TIERSLIVRE', CodeTiers) ;
  Tob_Piece.PutValue ('GP_TIERSFACTURE', CodeTiers) ;
  Tob_Piece.PutValue ('GP_TIERSPAYEUR', CodeTiers) ;
  ////////////////////////////////////////////////////////////////////////////
  // Initialise les zones de l'entête en fonction du client
  ////////////////////////////////////////////////////////////////////////////
  Tob_Piece.PutValue ('GP_ESCOMPTE', Tob_Tiers.GetValue ('T_ESCOMPTE')) ;
  Tob_Piece.PutValue ('GP_MODEREGLE', Tob_Tiers.GetValue ('T_MODEREGLE')) ;
  Tob_Piece.PutValue ('GP_REGIMETAXE', Tob_Tiers.GetValue ('T_REGIMETVA')) ;
  Tob_Piece.PutValue ('GP_TVAENCAISSEMENT', Tob_Tiers.GetValue ('T_TVAENCAISSEMENT')) ;
  Tob_Piece.PutValue ('GP_QUALIFESCOMPTE', Tob_Tiers.GetValue ('T_QUALIFESCOMPTE')) ;
  Tob_Piece.PutValue ('GP_TARIFTIERS', Tob_Tiers.GetValue ('T_TARIFTIERS')) ;
  /////////////////////////////////////////////////////////////////////////////
  // Facturation HT ou TTC
  /////////////////////////////////////////////////////////////////////////////
  Tob_Piece.PutValue ('GP_FACTUREHT', Tob_Tiers.GetValue ('T_FACTUREHT'));
  ///////////////////////////////////////////////
  // Champs par  défaut
  //////////////////////////////////////////////
  Tob_Piece.PutValue ('GP_VENTEACHAT', GetInfoParPiece(NaturePiece, 'GPP_VENTEACHAT'));
  Tob_Piece.PutValue ('GP_NUMADRESSELIVR', -1);
  Tob_Piece.PutValue ('GP_NUMADRESSEFACT', -1);
  Tob_Piece.PutValue ('GP_CREATEUR', V_PGI.User) ;
  Tob_Piece.PutValue ('GP_SOCIETE', V_PGI.CodeSociete);
  Tob_Piece.PutValue('GP_CREEPAR', 'GEN');
  if GetInfoParPiece(NaturePiece, 'GPP_ACTIONFINI') = 'ENR' then Tob_Piece.PutValue('GP_VIVANTE', '-')
                                                            else Tob_Piece.PutValue('GP_VIVANTE', 'X');
  if GetInfoParPiece(NaturePiece, 'GPP_VISA') = 'X' then Tob_Piece.PutValue('GP_ETATVISA', 'ATT')
                                                    else Tob_Piece.PutValue('GP_ETATVISA', 'NON');
  /////////////////////////
  // Devise du document
  ////////////////////////
  Tob_Piece.PutValue ('GP_DEVISE', V_PGI.DevisePivot);
  DEV.Code:=Tob_Piece.GetValue ('GP_DEVISE');
  GETINFOSDEVISE(DEV) ;
  DEV.Taux:=GETTAUX(DEV.Code, DEV.DateTaux, CleDoc.DatePiece) ;
  PutValueDetail (Tob_Piece, 'GP_TAUXDEV', DEV.Taux) ;
  PutValueDetail (Tob_Piece, 'GP_COTATION', DEV.Taux) ;
  PutValueDetail (Tob_Piece, 'GP_DATETAUXDEV', DEV.DateTaux) ;

  DD:=Tob_Piece.GetValue('GP_DATEPIECE') ;
  Per:=GetPeriode(DD) ; Sem:=NumSemaine(DD) ;
  Tob_Piece.PutValue ('GP_PERIODE', Per) ;
  Tob_Piece.PutValue ('GP_SEMAINE', Sem) ;
  result:=True;
end;

function TValidateForm.ChargeTOBClient : Boolean;
var Q : TQuery;
begin
  result:=true;
  CodeTiers := GetParamsoc('SO_GCTIERSINV') ;
  Q := OpenSQL('Select * from Tiers Where T_TIERS="'+CodeTiers+'"' ,False) ;
  if not Q.EOF then
    begin
    Tob_Tiers.initValeurs;
    Tob_Tiers.SelectDB('', Q);
    end
  else
    begin
    PGIInfo(TraduireMemoire(TexteMessage[1]), Caption);
    result := False ;
    end;
  Ferme(Q) ;
end;

function TValidateForm.ConstruireLignePiece (CodeArtGen,CodeArtDim:String; TOBListe:TOB; QteEcart:double; IndiceArticle:integer; IsGenerique:boolean) : boolean;
var TobArt : TOB;
    SQL, TypeDim, Designation : string ;
    Q : TQuery ;
    Prix : double ;
    i : integer ;
begin
  result := True;
  Tob_Ligne:= TOB.CREATE ('LIGNE', Tob_Piece, -1);
//  Tob_Ligne.initValeurs;
  ////////////////////////////////////////////////////////////////
  // Initialise les champs de la ligne à partir de l'entête
  ////////////////////////////////////////////////////////////////
  PieceVersLigne (Tob_Piece, Tob_Ligne);
  Tob_Ligne.PutValue('GL_PERIODE', Tob_Piece.GetValue('GP_PERIODE')) ;
  Tob_Ligne.PutValue('GL_SEMAINE', Tob_Piece.GetValue('GP_SEMAINE')) ;
  ////////////////////////////////////////////////////////////////
  // Ajoute des champs supplémentaires pour le calcul des cumuls
  ////////////////////////////////////////////////////////////////
  AddLesSupLigne (Tob_Ligne, False) ; // Ajout Champs spécifiques pour calcul de la pièce
  //JD False à la place de True

  ///////////////////////////////////////////////////////////////
  // Détermination de l'identifiant article
  //////////////////////////////////////////////////////////////
  //TobArt:=Tob_Article.Detail[IndiceArticle];
  TobArt:=Tob_Article.FindFirst(['GA_ARTICLE'],[CodeArtDim],False);
  if TobArt=nil then     //Ne doit plus passer ici, la Tob est pré-chargée
    begin
    SQL:=StSelectArticle+' WHERE GA_ARTICLE="'+CodeArtDim+'"';
    Q:=OpenSQL(SQL,True) ;
    if Not Q.EOF then
      begin
      ////////////////////////////////////////////////////////////////////////////////////////////////
      // Chargement de l'article dans la TOB des articles de la piece
      ////////////////////////////////////////////////////////////////////////////////////////////////
      TobArt := CreerTOBArt(Tob_Article);
      TobArt.SelectDB('',Q);
      ///////////////////////////////////////////////////////////////////////////////////////////////
      // Chargement éventuel des fiches stocks pour MAJ du stock
      ///////////////////////////////////////////////////////////////////////////////////////////////
      if TobArt.GetValue('GA_STATUTART') <> 'GEN' then LoadTOBDispo(TobArt, True) ;
      end
    else
      begin
      PGIInfo(TraduireMemoire(TexteMessage[4])+' '+CodeArtDim,Caption);
      result := False ;
      end;
    Ferme(Q);
    end;
  if TobArt<>nil then
    begin
    /////////////////////////////////////////////////////////////////////////////
    // Initialisation des champs de la ligne à partir de l'article dimensionné
    /////////////////////////////////////////////////////////////////////////////
    Designation := TobArt.GetValue('GA_LIBELLE');
    EntreCote(Designation,true);
    Designation:=copy(Designation,1,ChampToLength('GA_LIBELLE'));  // ENtreCote ajoute une quote si présence d'une quote dans le libellé ....
    //TypeDim := TobArt.GetValue('GA_STATUTART');
    if IsGenerique then TypeDim:='GEN' else TypeDim := TobArt.GetValue('GA_STATUTART');  //JD
    if TypeDim <> 'GEN' then
      begin
      Tob_Ligne.PutValue ('GL_PRIXPOURQTE'   , TobArt.GetValue('GA_PRIXPOURQTE'));
      Tob_Ligne.PutValue ('GL_ESCOMPTABLE'   , TobArt.GetValue('GA_ESCOMPTABLE'));
      Tob_Ligne.PutValue ('GL_REMISABLEPIED' , TobArt.GetValue('GA_REMISEPIED'));
      Tob_Ligne.PutValue ('GL_REMISABLELIGNE', TobArt.GetValue('GA_REMISELIGNE'));
      Tob_Ligne.PutValue ('GL_TENUESTOCK'    , TobArt.GetValue('GA_TENUESTOCK'));
      Tob_Ligne.PutValue ('GL_TARIFARTICLE'  , TobArt.GetValue('GA_TARIFARTICLE'));
      Tob_Ligne.PutValue ('GL_QUALIFSURFACE' , TobArt.GetValue('GA_QUALIFSURFACE'));
      Tob_Ligne.PutValue ('GL_QUALIFVOLUME'  , TobArt.GetValue('GA_QUALIFVOLUME'));
      Tob_Ligne.PutValue ('GL_QUALIFPOIDS'   , TobArt.GetValue('GA_QUALIFPOIDS'));
      Tob_Ligne.PutValue ('GL_QUALIFLINEAIRE', TobArt.GetValue('GA_QUALIFLINEAIRE'));
      Tob_Ligne.PutValue ('GL_QUALIFHEURE'   , TobArt.GetValue('GA_QUALIFHEURE'));
      Tob_Ligne.PutValue ('GL_SURFACE'       , TobArt.GetValue('GA_SURFACE'));
      Tob_Ligne.PutValue ('GL_VOLUME'        , TobArt.GetValue('GA_VOLUME'));
      Tob_Ligne.PutValue ('GL_POIDSBRUT'     , TobArt.GetValue('GA_POIDSBRUT'));
      Tob_Ligne.PutValue ('GL_POIDSNET'      , TobArt.GetValue('GA_POIDSNET'));
      Tob_Ligne.PutValue ('GL_POIDSDOUA'     , TobArt.GetValue('GA_POIDSDOUA'));
      Tob_Ligne.PutValue ('GL_LINEAIRE'      , TobArt.GetValue('GA_LINEAIRE'));
      Tob_Ligne.PutValue ('GL_HEURE'         , TobArt.GetValue('GA_HEURE'));
      Tob_Ligne.PutValue ('GL_QUALIFQTESTO'  , TobArt.GetValue('GA_QUALIFUNITESTO'));
      Tob_Ligne.PutValue ('GL_FAMILLENIV1'   , TobArt.GetValue('GA_FAMILLENIV1')) ;
      Tob_Ligne.PutValue ('GL_FAMILLENIV2'   , TobArt.GetValue('GA_FAMILLENIV2')) ;
      Tob_Ligne.PutValue( 'GL_FAMILLENIV3'   , TobArt.GetValue('GA_FAMILLENIV3')) ;
      Tob_Ligne.PutValue( 'GL_FOURNISSEUR'   , TobArt.GetValue('GA_FOURNPRINC')) ;
      for i:=1 to 9 do Tob_Ligne.PutValue('GL_LIBREART'+IntToStr(i),TobArt.GetValue('GA_LIBREART'+IntToStr(i))) ;
      Tob_Ligne.PutValue ('GL_LIBREARTA'  , TobArt.GetValue('GA_LIBREARTA'));
      end;
    end;
  for i:=1 to 5 do Tob_Ligne.PutValue('GL_FAMILLETAXE'+IntToStr(i),TobArt.GetValue('GA_FAMILLETAXE'+IntToStr(i))) ;
  if result then
    begin
    if TypeDim='GEN' then
      begin
      Tob_Ligne.PutValue ('GL_LIBELLE', Designation);
      Tob_Ligne.PutValue ('GL_TYPELIGNE', 'COM');   // Ligne commentaire
      Tob_Ligne.PutValue ('GL_CODESDIM', CodeArtGen);
      end
    else
      begin
      Tob_Ligne.PutValue ('GL_ARTICLE', CodeArtDim);
      Tob_Ligne.PutValue ('GL_CODEARTICLE', CodeArtGen);
      if TypeDim='DIM' then Tob_Ligne.PutValue('GL_LIBELLE',Copy(Designation+RetourneLibelleAvecDimensions(CodeArtDim,TOBArt),1,70))
      else Tob_Ligne.PutValue('GL_LIBELLE',Designation);
      Tob_Ligne.PutValue ('GL_TARIF', 0);
      Tob_Ligne.PutValue ('GL_TYPELIGNE', 'ART');
      end;
    Tob_Ligne.PutValue ('GL_REFARTSAISIE', CodeArtGen);
    Tob_Ligne.PutValue ('GL_TYPEARTICLE' , 'MAR');
    Tob_Ligne.PutValue ('GL_VALIDECOM'   , 'AFF');
    if TypeDim='UNI' then  Tob_Ligne.PutValue('GL_TYPEDIM', 'NOR')
    else Tob_Ligne.PutValue('GL_TYPEDIM', TypeDim);
    Tob_Ligne.PutValue('GL_PRIXPOURQTE', 1);
    Tob_Ligne.PutValue('GL_QUALIFMVT', QualifMvt) ;
    Tob_Ligne.PutValue('GL_CREERPAR', 'GEN');
    ////////////////////////////////////////////////////////////////////////////
    // Traite les quantités
    ////////////////////////////////////////////////////////////////////////////
    Tob_Ligne.PutValue('GL_QTESTOCK', QteEcart);
    Tob_Ligne.PutValue('GL_QTEFACT' , QteEcart);
    Tob_Ligne.PutValue('GL_QTERESTE', QteEcart); { NEWPIECE }
    /////////////////////////////////////////////////////////
    // Récupération des Prix
    //////////////////////////////////////////////////////////
    if (ctxMode in V_PGI.PGIContexte) then
       begin
       if TypeDim<>'GEN' then
         begin
         Tob_Ligne.PutValue('GL_DPA', ChoozePrice('GIL_DPA', TOBListe));
         Tob_Ligne.PutValue('GL_DPR', ChoozePrice('GIL_DPR', TOBListe));

         Prix:=ChoozePrice('GIL_PMAP', TOBListe);
         Tob_Ligne.PutValue('GL_PMAP', Prix);
         Tob_Ligne.PutValue('GL_PMAPACTU', Prix);

         Prix:=ChoozePrice('GIL_PMRP', TOBListe);
         Tob_Ligne.PutValue('GL_PMRP', Prix);
         Tob_Ligne.PutValue('GL_PMRPACTU', Prix);

         Prix:=QuelPrixBaseInv();
         if EnHT then
           begin
           Tob_Ligne.PutValue('GL_PUHT',Prix); Tob_Ligne.PutValue('GL_PUHTDEV',Prix);
           end
           else begin
           Tob_Ligne.PutValue('GL_PUTTC',Prix); Tob_Ligne.PutValue('GL_PUTTCDEV',Prix);
           end;
         end;
       end else
       begin
         if TypeDim<>'GEN' then
           begin
           Tob_Ligne.PutValue('GL_DPA', ChoozePrice('GIL_DPA', TOBListe));
           Tob_Ligne.PutValue('GL_DPR', ChoozePrice('GIL_DPR', TOBListe));
           Tob_Ligne.PutValue('GL_PMAP', ChoozePrice('GIL_PMAP', TOBListe));
           Tob_Ligne.PutValue('GL_PMRP', ChoozePrice('GIL_PMRP', TOBListe));
           prix := ChoozePrice('GIL_DPA', TOBListe);
           if EnHT then
             begin
             Tob_Ligne.PutValue('GL_PUHT',Prix); Tob_Ligne.PutValue('GL_PUHTDEV',Prix);
           end else
             begin
             Tob_Ligne.PutValue('GL_PUTTC',Prix); Tob_Ligne.PutValue('GL_PUTTCDEV',Prix);
             end;
//           AffectePrixValo(Tob_Ligne,TobArt);
         end;
       end;
    end
  else
    Tob_Ligne.free;
end;

procedure TValidateForm.ConstruireLigneLot;
var TOBNoeud, TobArt, TobInvLot, TOBLigneLot : TOB;
    IndiceLot : integer;
    Ecart, QteInv : Double;
begin
  TOBNoeud := CreerNoeudLot(Tob_Piece, TOBDesLots) ;
  IndiceLot := TOBDesLots.Detail.Count ;
  Tob_Ligne.PutValue('GL_INDICELOT',IndiceLot) ;

  TobArt := Tob_Article.FindFirst(['GA_ARTICLE'],[Tob_Ligne.GetValue('GL_ARTICLE')],False);
  if TobArt.GetValue('GA_LOT') = 'X' then
  begin
    TobInvLot := Tob_ListeLot.FindFirst(['GLI_ARTICLE'],[TobArt.GetValue('GA_ARTICLE')],False) ;
    while TobInvLot <> Nil do
    begin
      QteInv := TobInvLot.GetValue('GLI_INVENTAIRE');
      Ecart := QteInv - TobInvLot.GetValue('GLI_QTEPHOTOINV');
      if Ecart <> 0 then
      begin
        TOBLigneLot := TOB.Create('LIGNELOT',TOBNoeud,-1);
        With TOBLigneLot do
        begin
          PutValue('GLL_NATUREPIECEG', Tob_Ligne.GetValue('GL_NATUREPIECEG'));
          PutValue('GLL_SOUCHE',       Tob_Ligne.GetValue('GL_SOUCHE'));
          PutValue('GLL_NUMERO',       Tob_Ligne.GetValue('GL_NUMERO'));
          PutValue('GLL_INDICEG',      Tob_Ligne.GetValue('GL_INDICEG'));
          PutValue('GLL_NUMLIGNE',     Tob_Ligne.GetValue('GL_NUMLIGNE'));
          PutValue('GLL_RANG',         0);
          PutValue('GLL_ARTICLE',      Tob_Ligne.GetValue('GL_ARTICLE'));
          PutValue('GLL_NUMEROLOT',    TobInvLot.GetValue('GLI_NUMEROLOT'));
          PutValue('GLL_QUANTITE',     Ecart);
          //AddChampSupValeur('DATELOT',TOBDesLots.Detail[i_ind].GetValue('GQL_DATELOT'));
        end;
      end;
      TobInvLot := Tob_ListeLot.FindNext(['GLI_ARTICLE'],[TobArt.GetValue('GA_ARTICLE')],False) ;
    end;
  end;
end;

procedure TValidateForm.BuildTob_TobArticle(Tob_Liste : TOB);
var TOBL          : TOB;
    TObLArt       : TOB;
    TobDispo      : TOB;
    TobLDispo     : TOB;
    TobDispoLot   : TOB;
    TobLDispoLot  : TOB;
    //
    i             : Integer;
    //
    QteInv        : Double;
    //QtePhoto      : Double;
    //Ecart         : Double;
    //
    CodeArt       : String;
    //
    QArticle      : TQuery;
    QDispo        : TQuery;
    QDispoLot     : TQuery;
    //
begin

  //FV1 : 15/02/2016 - FS#1857 - VEUVE CHATAIN - problème de cohérence des quantités en stocks

  //J'essaye de comprendre ce qui a pu passer par la tête de l'esprit malade de l'individu qui a progammé ça !!!!
  //on lit les lignes d'inventaire, on sauvegarde les articles et ensuite on charge la tob des articles et des dispos....
  //Pourquoi on ne fait pas tout en même temps... On est sur un article on charge ce dernier, son dispo, ...
  //Cela me semble plus sûr et plus juste que la gestion via Tableau !!!!
  //Si l'on désire avoir la gestion des dispos par lot (couleur, taille, ...) il faut se reporter au lignes en commentaire

  INITMOVE(Tob_Liste.Detail.Count,'');

  MoveCur(False);

  //Création de la TobDispo et de la TobDispoLot
  TOBDispo  := TOB.CREATE ('Les Dispos', nil, -1);
  if GereLot then   TOBDispoLot  := TOB.CREATE ('Les Dispos Lot', nil, -1);

  //On garde la boucle de lesture des ligne de la liste d'inventaire
  For i := 0 to Tob_Liste.Detail.count-1 do
  begin
    TOBL      := Tob_Liste.detail[i]; //On charge une ligne à la fois
    QteInv    := TOBL.GetDouble('GIL_INVENTAIRE');  //On charge la quantité d'inventaire saisie
    //
    if (QteInv=0) and (FRemiseStockZero=False) then continue;   //Si la Qte d'inventaire est = 0 on passe à la ligne suivante
    //
    //QtePhoto  := TOBL.GetDouble('GIL_QTEPHOTOINV');           //Pas utilisée même pas un tout petit peu...
    //Ecart     := QteInv - QtePhoto;                           //Sinon on calcul l'écart... ce quine sert à rien vu qu'on ne l'utilise pas pas plus que la quantité d'inventaire calculée au dessus...
    //
    CodeArt   := TOBL.GetString('GIL_ARTICLE');
    //Chargement de la tobArticle
    QArticle  := OpenSQL(StSelectArticle+' WHERE GA_ARTICLE = "' + CodeArt + '"',True,1,'',True);
    if Not QArticle.EOF then
    begin
      TObLArt := TOB.Create('Article',Tob_Article, -1);
      TobLArt.SelectDB('ARTICLE',QArticle,False);
      TobLArt.AddChampSupValeur('UTILISE', '-');
      TobLArt.Modifie := False;
    end;
    Ferme(QArticle);
    //Chargement de la TObDispo
    QDispo    := OpenSQL('SELECT * FROM DISPO WHERE GQ_ARTICLE ="' + CodeArt + '" AND GQ_DEPOT="'+ Depot + '"',True,1,'',True);
    if Not QDispo.EOF   then
    begin
      TOBLDispo := TOB.CREATE ('Dispo', TOBDispo, -1);
      TOBLDispo.SelectDB('DISPO',QDispo,False);
    end;
    Ferme(QDispo);
    //Charge DispoLot
    if GereLot then
    begin
      QDispoLot := OpenSQL('SELECT * FROM DISPOLOT WHERE GQL_ARTICLE ="' + CodeArt + '" AND GQL_DEPOT="'+ Depot + '"',True,1,'',True);
      if Not QDispoLot.EOF then
        TOBLDispoLot := TOB.CREATE ('Dispo Lot', TOBDispoLot, -1);
        TOBLDispoLot.SELECTDB('DISPOLOT',QDispoLot,false);
      end;
     Ferme(QDispoLot);
  end;

  FINIMOVE;

  //Affectation du stock aux articles sélectionnés
  For i:=0 to Tob_Article.detail.count-1 do
  begin
    TOBLArt   := Tob_Article.detail[i];
    CodeArt   := TOBLArt.GetString('GA_ARTICLE');
    TobLDispo := TOBDispo.FindFirst(['GQ_ARTICLE'],[CodeArt],False) ;
    if TobLDispo <> nil then TobLDispo.Changeparent(TOBLArt,-1);
    if GereLot then
    begin
      TobLDispoLot := TOBDispoLot.FindFirst(['GQL_ARTICLE'],[CodeArt],False) ;
      if TobLDispoLot <> nil then TobLDispoLot.Changeparent(TobLDispo,-1);
    end;
  end;

  FreeandNil(TOBDispo);
  if GereLot then FreeandNil(TOBDispoLot);
  DispoChampSupp(Tob_Article);

  //Tout ça devrait disparaître mais on le garde par nostalgie
  //ou pour les générations à venir pour montrer ce qu'il ne faut pas
  //faire... Garder une trace de ce génie méconnu
  {*
  CountStArt:=0; NbRequete:=0;
  SetLength(TabWhereArticle, 1); SetLength(TabWhereDispo, 1);
  INITMOVE(Tob_Liste.Detail.Count,'');
  For i:=0 to Tob_Liste.Detail.count-1 do
  begin
    MoveCur(False);
    TL:=Tob_Liste.Detail[i];
    QteInv := TL.GetValue('GIL_INVENTAIRE');
    Ecart := QteInv - TL.GetValue('GIL_QTEPHOTOINV');
    //if (Ecart=0) or ((QteInv=0) and (FRemiseStockZero=False)) then continue;
    if (QteInv=0) and (FRemiseStockZero=False) then continue;

    StArticle:=TL.GetValue('GIL_ARTICLE');
    if CountStArt>=NbArtParRequete then
    begin
      NbRequete:=NbRequete+1;
      SetLength(TabWhereArticle, NbRequete+1); SetLength(TabWhereDispo, NbRequete+1);
      CountStArt:=0;
    end;
    CountStArt:=CountStArt+1;
    if TabWhereArticle[NbRequete]='' then TabWhereArticle[NbRequete] := '"'+StArticle+'"'
    else TabWhereArticle[NbRequete] := TabWhereArticle[NbRequete]+',"'+StArticle+'"';

    if TabWhereDispo[NbRequete]='' then TabWhereDispo[NbRequete] := '"'+StArticle+'"'
    else TabWhereDispo[NbRequete] := TabWhereDispo[NbRequete]+',"'+StArticle+'"';
  end;
  FINIMOVE;

  if TabWhereArticle[0]<>'' then
  begin
    TOBDispo := TOB.CREATE ('Les Dispos', nil, -1);
    if GereLot then TOBDispoLot := TOB.CREATE ('Les Dispolots', nil, -1);

    for y:=Low(TabWhereArticle) to High(TabWhereArticle) do
    begin
      if TabWhereArticle[y]<>'' then
      begin
        QArticle:=OpenSQL(StSelectArticle+' WHERE GA_ARTICLE IN ('+TabWhereArticle[y]+')',True);
        if Not QArticle.EOF then Tob_Article.LoadDetailDB('ARTICLE','','',QArticle,True,True);
        Ferme(QArticle);
      end;
      if TabWhereDispo[y]<>'' then
      begin
        //Charge Dispo
        QDispo :=OpenSQL('SELECT * FROM DISPO WHERE GQ_ARTICLE IN ('+TabWhereDispo[y]+') AND GQ_DEPOT="'+Depot+'"',True);
        if Not QDispo.EOF then TOBDispo.LoadDetailDB('DISPO','','',QDispo,True,True);
        Ferme(QDispo);
        //Charge DispoLot
        if GereLot then
        begin
          QDispoLot :=OpenSQL('SELECT * FROM DISPOLOT WHERE GQL_ARTICLE IN ('+TabWhereDispo[y]+') AND GQL_DEPOT="'+Depot+'"',True);
          if Not QDispoLot.EOF then TOBDispoLot.LoadDetailDB('DISPOLOT','','',QDispoLot,True,True);
          Ferme(QDispoLot);
        end;
      end;
    end;
    For x:=0 to Tob_Article.detail.count-1 do
    begin
      Tob_Article.detail[x].AddChampSup('UTILISE',False); Tob_Article.detail[x].PutValue('UTILISE','-');
      Tob_Article.detail[x].Modifie := False;
    end;
    //Affecte les stocks aux articles sélectionnés
    for NbArt:=0 to Tob_Article.detail.Count-1 do
    begin
      TOBArt := Tob_Article.detail[NbArt];
      TobDispoArt := TOBDispo.FindFirst(['GQ_ARTICLE'],[TOBArt.GetValue('GA_ARTICLE')],False) ;
      if TobDispoArt<>nil then
      begin
        if GereLot then
        begin
          TobDispoLotArt := TOBDispoLot.FindFirst(['GQL_ARTICLE','GQL_DEPOT'],[TobDispoArt.GetValue('GQ_ARTICLE'),TobDispoArt.GetValue('GQ_DEPOT')],False) ;
          while TobDispoLotArt <> nil do
            begin
            TobDispoLotArt.Changeparent(TobDispoArt,-1);
            TobDispoLotArt := TOBDispoLot.FindNext(['GQL_ARTICLE','GQL_DEPOT'],[TobDispoArt.GetValue('GQ_ARTICLE'),TobDispoArt.GetValue('GQ_DEPOT')],False) ;
            end;
        end;
        TobDispoArt.Changeparent(TOBArt,-1);
      end;
    end;
    TOBDispo.Free;
    if GereLot then TOBDispoLot.Free;
    DispoChampSupp(Tob_Article);
  end;
  *}

end;

procedure TValidateForm.MAJQtePrixLignesGEN();
var NbLignes,i,NbDec : integer;
    CodeArtGEN,TypeDim,CodeArt : String;
    TOBLGEN :TOB;
    MoyDPA,MoyPMAP,MoyDPR,MoyPMRP : double ;
    TotQteFact,TotQteStock : double;
begin
i:=0; TOBLGEN:=Nil;
NbDec:=DEV.Decimale ;
while i<Tob_Piece.Detail.Count do
  begin
  TypeDim:=Tob_Piece.Detail[i].GetValue('GL_TYPEDIM');
  if TypeDim='GEN' then
    begin
    TOBLGEN:=Tob_Piece.Detail[i]; if TOBLGEN=nil then exit;
    CodeArtGEN:=Tob_Piece.Detail[i].GetValue('GL_REFARTSAISIE');
    i:=i+1;
    end
  else if TypeDim='DIM' then
    begin
    CodeArt:=Tob_Piece.Detail[i].GetValue('GL_CODEARTICLE');
    //les prix du ligne générique correspondent à la moyenne
    if (CodeArtGEN=CodeArt) then
      begin
      MoyDPA:=0; MoyPMAP:=0; MoyDPR:=0; MoyPMRP:=0; NbLignes:=0;
      TotQteFact:=0; TotQteStock:=0;
      while (TypeDim='DIM') and (CodeArtGEN=CodeArt) and (i<Tob_Piece.Detail.Count) do
        begin
        MoyDPA:=MoyDPA+Tob_Piece.Detail[i].GetValue('GL_DPA');
        MoyPMAP:=MoyPMAP+Tob_Piece.Detail[i].GetValue('GL_PMAP');
        MoyDPR:=MoyDPR+Tob_Piece.Detail[i].GetValue('GL_DPR');
        MoyPMRP:=MoyPMRP+Tob_Piece.Detail[i].GetValue('GL_PMRP');
        TotQteFact:=TotQteFact+Tob_Piece.Detail[i].GetValue('GL_QTEFACT');
        TotQteStock:=TotQteStock+Tob_Piece.Detail[i].GetValue('GL_QTESTOCK');
        NbLignes:=NbLignes+1; i:=i+1;
        if i<Tob_Piece.Detail.Count then
          begin
          CodeArt:=Tob_Piece.Detail[i].GetValue('GL_CODEARTICLE');
          TypeDim:=Tob_Piece.Detail[i].GetValue('GL_TYPEDIM');
          end;
        end;
      if NbLignes<>0 then
        begin
        MoyDPA:=Arrondi(MoyDPA/NbLignes,NbDec); MoyPMAP:=Arrondi(MoyPMAP/NbLignes,NbDec);
        MoyDPR:=Arrondi(MoyDPR/NbLignes,NbDec); MoyPMRP:=Arrondi(MoyPMRP/NbLignes,NbDec);
        end;
      TOBLGEN.PutValue('GL_DPA',MoyDPA); TOBLGEN.PutValue('GL_PMAP',MoyPMAP);
      TOBLGEN.PutValue('GL_DPR',MoyDPR); TOBLGEN.PutValue('GL_PMRP',MoyPMRP);
      TOBLGEN.PutValue('GL_QTEFACT',TotQteFact);
      TOBLGEN.PutValue('GL_QTESTOCK',TotQteStock);
      TOBLGEN.PutValue('GL_QTERESTE', TotQteStock);  { NEWPIECE }
      if EnHT then
        begin
        TOBLGEN.PutValue('GL_PUHTDEV',MoyDPA);
        CalculeLigneHT(TOBLGEN,Tob_LigneBase,Tob_Piece,DEV, NbDec, True);
        end
      else
        begin
        TOBLGEN.PutValue('GL_PUTTCDEV',MoyDPA);
        CalculeLigneTTC (TOBLGEN,Tob_LigneBase,Tob_Piece,DEV, NbDec,True) ;
        end;
      end
      else i:=i+1;
    end
  else i:=i+1;
  end;
end;

procedure TValidateForm.ValideLesLots(TOBDesLots, TOBArticles, TOBPiece : TOB; VenteAchat : String) ;
Var i,IndiceLot : integer ;
    TOBDL,TOBL : TOB ;
begin
  if Not GereLot then Exit ;
  if TOBDesLots.Detail.Count<=0 then Exit ;
  TOBDesLots.Detail[0].AddChampSup('UTILISE',True) ;
  for i:=0 to TOBPiece.Detail.Count-1 do
      BEGIN
      IndiceLot:=GetChampLigne(TOBPiece,'GL_INDICELOT',i+1) ;
      if IndiceLot>0 then
         BEGIN
         TOBL:=TOBPiece.Detail[i] ; TOBDL:=TOBDesLots.Detail[IndiceLot-1] ;
         TOBDL.PutValue('UTILISE','X') ;
         UpdateNoeudLot(TOBL,TOBDL) ;
         MAJDispoLot(TOBArticles,TOBL,TOBDL,VenteAchat,TOBPiece.GetValue('GP_DEPOT'));
         END ;
      END ;
  for i:=TOBDesLots.Detail.Count-1 downto 0 do
      BEGIN
      TOBDL:=TOBDesLots.Detail[i] ;
      if TOBDL.GetValue('UTILISE')<>'X' then TOBDL.Free ;
      END ;
  if TOBDesLots.Detail.Count>0 then TOBPiece.PutValue('GP_ARTICLESLOT','X') ;
  //if Not TOBDesLots.InsertDB(Nil) then V_PGI.IoError:=oeUnknown ;
  TOBDesLots.InsertDB(Nil);
end ;

function TValidateForm.ValideLesPieces : boolean;
begin
BEGINTRANS;
result:=False;
Try
  Tob_PiedBase.SetAllModifie(True) ;
  Tob_LigneBase.setAllModifie (true);
  if Tob_Piece.Detail.Count>0 then
    begin
//    Tob_Piece.SetAllModifie(True);      //Ne pas remettre JD
    ////////////////////////////////////////////////////////////////////
    // MAJ du stock
    ////////////////////////////////////////////////////////////////////
    ValideLesLignes (Tob_Piece, Tob_Article, Tob_Catalogue, Tob_Noment,nil,nil,nil, False,False) ;
    ValideLesLots (TOBDesLots, Tob_Article, Tob_Piece, 'ACH') ; //Les inventaires fonctionnent comme les achats concernant les stocks
    ValideLesArticles (Tob_Piece, Tob_Article) ;

    ValideLeTiers(Tob_Piece, Tob_Tiers);

    Tob_PiedBase.InsertDB(nil);

    ////////////////////////////////////////////////////////////////////
    // MAJ des TOB -> DB de la piece
    ////////////////////////////////////////////////////////////////////
    Tob_Piece.InsertDB(nil,True);
    end;
COMMITTRANS;
result:=true;
Except
 ROLLBACK;
 end;
end;

function TValidateForm.RetourneLibelleAvecDimensions(RefArtDim:String;TOBArt:TOB): string;
var k : integer;
    Grille,CodeDim,LibDim: String ;
begin
LibDim:='';
if TOBArt<>Nil then
  begin
  for k:=1 to 5 do
    begin
    Grille:=TOBArt.GetValue('GA_GRILLEDIM'+IntToStr(k)) ;
    CodeDim:=TOBArt.GetValue('GA_CODEDIM'+IntToStr(k)) ;
    if ((Grille<>'') and (CodeDim<>'')) then
      LibDim:=LibDim+'  '+RechDom('GCGRILLEDIM'+IntToStr(k),Grille,True)+' '+GCGetCodeDim(Grille,CodeDim,k);
    end;
  end;
Result:=LibDim;
end;

Function TValidateForm.QuelPrixBaseInv () : double ;
Var CodeP : String ;
    Prix  : Double ;
begin
Prix:=Tob_Ligne.GetValue('GL_DPA');
CodeP:=GetInfoParPiece(NaturePiece,'GPP_APPELPRIX') ;
if CodeP='DPR' then Prix:=Tob_Ligne.GetValue('GL_DPR');
if CodeP='PPA' then Prix:=Tob_Ligne.GetValue('GL_PMAP');
if CodeP='PPR' then Prix:=Tob_Ligne.GetValue('GL_PMRP');
Result:=Prix ;
end;

(*procedure TValidateForm.lv_selectionCustomDrawItem(Sender: TCustomListView; Item: TListItem; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
if Item.ImageIndex = 1 then Sender.Canvas.Font.Style := [fsBold]
                       else Sender.Canvas.Font.Style := [];
end;*)

procedure TValidateForm.bCancelClick(Sender: TObject);
begin
if FValidating then FCanceled := true
               else Close;
end;

end.
