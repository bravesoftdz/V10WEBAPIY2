unit AssistStockAjust;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
   Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
   assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls, HEnt1,
   HPanel, UTob,  HStatus, Grids, EntGC,
{$IFDEF EAGLCLIENT}
{$ELSE}
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
   HDimension, Mask, TntStdCtrls, TntComCtrls, TntExtCtrls;

procedure EntreeStockAjust;
function  EntreeStockAjustUnArticle(RefArt,RefDep : string) : Boolean;

type
  TFStockAjust = class(TFAssist)
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    PTITRE: THPanel;
    HLabel1: THLabel;
    PanelFin: TPanel;
    TTextFin1: THLabel;
    TTextFin2: THLabel;
    TWarning: THLabel;
    LBX_Param: TListBox;
    HLabel2: THLabel;
    GBParam: TGroupBox;
    TDepot: THLabel;
    Depot: THValComboBox;
    TRecap: THLabel;
    ProgressBar: TProgressBar;
    TProgressBar: THLabel;
    TArticle: THLabel;
    CodeArticle: THCritMaskEdit;
    TCodeArticle: THLabel;
    procedure FormShow(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure bAnnulerClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);

  private
    { Déclarations privées }
    StopTT, TTdebut : boolean;
    i_cpte : integer;
    TobDispoMAJ, TobRecap : TOB;
    RefArticle,RefDepot : string;
    ModifDispo : Boolean;

    //Traitement de MAJ // retourne false si annulation par l'utilisateur
    function TraitementOK : boolean;
    function RecupDepot(Pref,Depot : string) : string;

    //fonctions de traitement des données, recherche des <>
    function  AjustQteStock(TobLigne : TOB ; TypeQte : string) : double; //Rech des pièces dans les lignes de doc
    function  GetQteDetail(TOBLigneEtNomen : TOB; Prefixe : String) : Double;
    function  MAJQuantite (TobDispo : TOB) : boolean; //MAJ des qtés dans la tob
    function  GetPlusMoins(TypeQte,NaturePiece : string) : integer;
    procedure ChargeDispo(TobDispo : TOB ; stArticle,stDepot : string);
    function  ChargeLigne(stArticle, stDepot : string; DateInv : Tdatetime) : TOB;

    //MAJ des quantités dans la TOBDispo
    procedure DebutTransaction;
    procedure MAJTableDispo;

    //Enreg du résultat dans le jal d'événements
    Procedure NoteEvenement(Evenement, Etat : string);

    //Gestion de la tob récap
    procedure ChargeTobRecap(TOBTT : TOB ; stQte : string ; OldQte,NewQte : double);
    function  RechDim(TobDispoArt : TOB) : string;
    procedure RechercheLot; //Recherche dans tob recap tous les articles gérés par lot
    procedure AjouteTobLot(TOBR : TOB);
    function ConstitueWhereFather(TOBLigneEtNomen: TOB): string; //Ajout de tob lot à toutes les tob recap dont l'article est géré par lot
                                        //si les quantités diffèrent


  public
    { Déclarations publiques }
  end;


implementation

uses  InvUtil,
      StockUtil,
      UTofAjustStock,
      UTofConsultStock,
      DateUtils,
      Ulog,
      uEntCommun,
      factComm,
      BTPUtil;

{$R *.DFM}
procedure EntreeStockAjust;
var Fo_Ajust : TFStockAjust;
begin
     Fo_Ajust := TFStockAjust.Create (Application);
     Fo_Ajust.RefArticle:='';
     Fo_Ajust.RefDepot:='';
     Try
         Fo_Ajust.ShowModal;
     Finally
         Fo_Ajust.free;
     End;
end;

function  EntreeStockAjustUnArticle(RefArt,RefDep : string) : Boolean;
var Fo_Ajust : TFStockAjust;
begin
Result:=False;
if (RefArt<>'') and (RefDep<>'')  then
   begin
   Fo_Ajust := TFStockAjust.Create (Application);
   Fo_Ajust.RefArticle:=RefArt;
   Fo_Ajust.RefDepot:=RefDep;
   Try
       Fo_Ajust.ShowModal;
   Finally
       Result:=Fo_Ajust.ModifDispo;
       Fo_Ajust.free;
   End;
   end;
end;

procedure TFStockAjust.FormShow(Sender: TObject);
begin
  inherited;
StopTT:=false; ModifDispo:=False;
Depot.ItemIndex:=0;
if RefDepot<>'' then
   begin
   Depot.Value:=RefDepot;
   Depot.Enabled:=False;
   end;
if RefArticle<>'' then
   begin
   CodeArticle.Visible:=True;
   CodeArticle.text:=Copy(RefArticle,1,18);
   TCodeArticle.Visible:=True;
   end;
end;


////////////////////////////////////////////////////////////////////////////////
//*************************Evénements de la form******************************//
////////////////////////////////////////////////////////////////////////////////

procedure TFStockAjust.bSuivantClick(Sender: TObject);
begin
  inherited;
if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
LBX_Param.Clear;
LBX_Param.Items.Add('');
LBX_Param.Items.Add('Dépôt : ' + Depot.Text);
if RefArticle<>'' then
   begin
   LBX_Param.Items.Add('Article : ' + CodeArticle.Text);
   end;
end;

procedure TFStockAjust.bPrecedentClick(Sender: TObject);
begin
  inherited;
if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
end;

procedure TFStockAjust.bFinClick(Sender: TObject);
begin
inherited;
bfin.Enabled:=false;
TobDispoMAJ:=TOB.Create('',nil,-1);
TobRecap := TOB.Create('GCTMPAJUSTSTOCK',nil,-1);
i_cpte:=0;

if  not TraitementOK then
    begin
    NoteEvenement('Traitement interrompu par l''utilisateur','INT');
    end else
    begin
    if TobRecap.Detail.Count > 0 then
       begin
       RechercheLot;
       if ValideStkAjust(TobRecap, Depot.Text) then DebutTransaction
       else NoteEvenement('Traitement annulé','INT');
       end else
       begin    //si pas de <> de stock
       Msg.Execute (3,Caption,'');
       NoteEvenement('Vérification OK','OK');
       end;
    end;
TProgressBar.Caption := ''; TArticle.Caption := '';
ProgressBar.Visible := false;
ProgressBar.Max := 0;
TTDebut:=false;
TobDispoMAJ.Free ; TobRecap.Free;

RestorePage ;
P.ActivePage := PreviousPage;
PChange(nil) ;
bPrecedent.Enabled := false;
end;

procedure TFStockAjust.bAnnulerClick(Sender: TObject);
begin
  inherited;
if not TTdebut then exit;
if Msg.Execute (1,Caption,'')=mryes then StopTT:=True
   else ModalResult:=0;
end;


////////////////////////////////////////////////////////////////////////////////
//*************************Chargement de la tob dispo*************************//
////////////////////////////////////////////////////////////////////////////////

////////Traitement d'ajustement du stock
function TFStockAjust.TraitementOK : boolean;
var i_dep,i_art : integer;
    TobDispo,TobAMvt : TOB;
    stSQL, stArticle, stDepot,WhereArt,FinSql : string;
begin
//Initialisation
Result := true ; TTdebut:=true;
stDepot := ''; WhereArt:='';
FinSql:=' ORDER BY GA_ARTICLE';
ProgressBar.Visible := true;
TobAMvt := TOB.Create('',nil,-1);

for i_dep := 1 to Depot.Items.Count -1 do
    begin
    if (Depot.Value <> '') and (Depot.Value <> Depot.Values[i_dep]) then continue;
    stDepot := Depot.Values[i_dep];
    if RefArticle<>'' then
       begin
       WhereArt:=' GA_ARTICLE="'+RefArticle+'" AND (';
       FinSql:=')';
       end;
     //récup des identifiants article mouvementés ou présent dans dispo
    stSQL := 'SELECT DISTINCT GA_ARTICLE, GA_LIBELLE FROM ARTICLE WHERE '+WhereArt+' EXISTS '+
        '(SELECT GQ_ARTICLE FROM DISPO WHERE GQ_ARTICLE=ARTICLE.GA_ARTICLE AND GQ_CLOTURE="-"'+RecupDepot('GQ',stDepot)+') OR EXISTS '+
        '(SELECT GL_ARTICLE FROM LIGNE WHERE GL_ARTICLE=ARTICLE.GA_ARTICLE AND GL_ARTICLE<>"" AND GL_TENUESTOCK="X" '
         +RecupDepot('GL',stDepot)+' AND GL_QUALIFMVT<>"ANN" '+') OR EXISTS '+
        '(SELECT BLO_ARTICLE FROM LIGNEOUV LEFT JOIN LIGNE ON '+
        '(GL_NATUREPIECEG=BLO_NATUREPIECEG AND GL_SOUCHE=BLO_SOUCHE AND GL_NUMERO=BLO_NUMERO AND GL_INDICEG=BLO_INDICEG AND GL_NUMLIGNE=BLO_NUMLIGNE) '+
        'WHERE BLO_ARTICLE=ARTICLE.GA_ARTICLE AND BLO_ARTICLE<>"" AND BLO_TENUESTOCK="X" '
         +RecupDepot('GL',stDepot)+' AND GL_QUALIFMVT<>"ANN" )'+FinSql;

    TobAMvt.LoadDetailFromSQL(stSQL);
    ProgressBar.Max := TobAMvt.Detail.Count;
    TProgressBar.Caption := Msg.Mess[5] + Depot.items[i_dep];

    for i_art := 0 to TobAMvt.Detail.Count -1 do
        begin
        stArticle := TobAMvt.Detail[i_art].GetValue('GA_ARTICLE');
        TArticle.Caption := Msg.Mess[6] + Trim(Copy(TobAMvt.Detail[i_art].GetValue('GA_ARTICLE'),1,18)) + '  ' +
            TobAMvt.Detail[i_art].GetValue('GA_LIBELLE');
        if StopTT then
           begin
           Result := false;
           TobAMvt.Free;
           exit;
           end;
        //Chargement des infos
        TobDispo := TOB.Create('DISPO',TobDispoMAJ,-1);
        ChargeDispo(TobDispo,stArticle,stDepot);
        ProgressBar.StepIt;
        if not MAJQuantite(TobDispo) then TobDispo.Free;//Dispo correct donc pas de MAJ
        end;
    TobAMvt.ClearDetail;
    ProgressBar.Max := 0;
    end;
TobAMvt.Free;
end;

function TFStockAjust.RecupDepot(Pref,Depot : string) : string;
begin
if Depot <> '' then Result := ' AND '+ Pref +'_DEPOT="' + Depot + '"'
    else Result := '';
end;


////////MAJ des quantités dans la TOB Dispo
function TFStockAjust.MAJQuantite (TobDispo : TOB) : boolean;
var Change : boolean;
     TobLigne : TOB;
     QtePhys,LivreCli, LivreFou, ReserveCli, ReserveFou, PrepaCli : double;
     PrevLivCli,PrevLivFou,PrevResCli,PrevResFou,PrevPrepCli,PrevQtePhys,QteMvt : Double;
     StArt,stDep : string;
    yy,mm,dd,hh,mn,ss,ms : word;
    DateDebutTrait : TDateTime;
begin

  Change := false;
  stArt := TobDispo.GetValue('GQ_ARTICLE');
  stDep := TobDispo.GetValue('GQ_DEPOT');

  PrevLivCli := TobDispo.GetValue('GQ_LIVRECLIENT');
  PrevLivFou := TobDispo.GetValue('GQ_LIVREFOU');
  PrevResCli := TobDispo.GetValue('GQ_RESERVECLI');
  PrevResFou :=  TobDispo.GetValue('GQ_RESERVEFOU');
  PrevPrepCli := TobDispo.GetValue('GQ_PREPACLI');
  PrevQtePhys := TobDispo.GetValue('GQ_PHYSIQUE');

  TobDispo.PutValue('GQ_PHYSIQUE',0);
  TobDispo.PutValue('GQ_LIVRECLIENT',0);
  TobDispo.PutValue('GQ_LIVREFOU',0);
  TobDispo.PutValue('GQ_RESERVECLI',0);
  TobDispo.PutValue('GQ_RESERVEFOU',0);
  TobDispo.PutValue('GQ_PREPACLI',0);

  LivreCli := 0;
  LivreFou := 0;
  ReserveCli := 0;
  ReserveFou := 0;
  PrepaCli := 0;
  QtePhys := TobDispo.GetValue('GQ_STOCKINV');

  DecodeDateTime (TOBDispo.GetDateTime('GQ_DATEINV'),yy,mm,dd,hh,mn,ss,ms);
  DateDebutTrait := EncodeDateTime(yy,mm,dd,0,0,0,0);

  TobLigne := ChargeLigne(stArt,stDep,DateDebutTrait );
  if TobLigne.Detail.Count>0 then
   begin
   LivreCli:=AjustQteStock(TobLigne,'LC');
    LivreFou:=AjustQteStock(TobLigne,'LF');
    ReserveCli:=AjustQteStock(TobLigne,'RC');
    ReserveFou:=AjustQteStock(TobLigne,'RF');
    PrepaCli:=AjustQteStock(TobLigne,'PRE');
    QteMvt := AjustQteStock(TobLigne,'PHY');
    QtePhys := TobDispo.GetValue('GQ_STOCKINV') + QteMvt;
  end;

   if LivreCli <> PrevLivCli then
      begin
       ChargeTobRecap(TobDispo,'Livré client',TobDispo.GetValue('GQ_LIVRECLIENT'),LivreCli);
       TobDispo.PutValue('GQ_LIVRECLIENT',LivreCli);
       Change := true;
    end else TobDispo.PutValue('GQ_LIVRECLIENT',PrevLivCli);
   if LivreFou <> PrevLivFou then
      begin
      ChargeTobRecap(TobDispo,'Livré fournisseur',TobDispo.GetValue('GQ_LIVREFOU'),LivreFou);
      TobDispo.PutValue('GQ_LIVREFOU',LivreFou);
      Change := true;
    end else TobDispo.PutValue('GQ_LIVREFOU',PrevLivFou);
   if ReserveCli <> PrevResCli then
      begin
      ChargeTobRecap(TobDispo,'Reservé client',TobDispo.GetValue('GQ_RESERVECLI'),ReserveCli);
      TobDispo.PutValue('GQ_RESERVECLI',ReserveCli);
      Change := true;
    end else TobDispo.PutValue('GQ_RESERVECLI',PrevResCli);
   if ReserveFou <> PrevResFou then
      begin
      ChargeTobRecap(TobDispo,'Reservé fournisseur',TobDispo.GetValue('GQ_RESERVEFOU'),ReserveFou);
      TobDispo.PutValue('GQ_RESERVEFOU',ReserveFou);
      Change := true;
    end else TobDispo.PutValue('GQ_RESERVEFOU',PrevResFou);
    if PrepaCli <> PrevPrepCli then
      begin
      ChargeTobRecap(TobDispo,'Preparé client',TobDispo.GetValue('GQ_PREPACLI'),PrepaCli);
      TobDispo.PutValue('GQ_PREPACLI',PrepaCli);
      Change := true;
    end else TobDispo.PutValue('GQ_PREPACLI',PrevPrepCli);
    if QtePhys <> PrevQtePhys then
    begin
      ChargeTobRecap(TobDispo,'Qté Physique',PrevQtePhys,QtePhys);
      TobDispo.PutValue('GQ_PHYSIQUE',QtePhys);
      Change := true;
    end else TobDispo.PutValue('GQ_PHYSIQUE',PrevQtePhys);

  TobLigne.Free;
  Result := Change;
end;

procedure TFStockAjust.ChargeDispo(TobDispo : TOB ; stArticle,stDepot : string);
var stSelect : string;
    QDispo : TQuery;
begin
  stSelect := ' GA_LOT, GA_LIBELLE, GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,'+
         'GA_GRILLEDIM4,GA_GRILLEDIM5,GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3,'+
         'GA_CODEDIM4,GA_CODEDIM5,GA_STATUTART ';
  QDispo:=OpenSQL('SELECT DISPO.*, ' + stSelect +
        'FROM DISPO LEFT JOIN ARTICLE ON GQ_ARTICLE=GA_ARTICLE '+
        'WHERE GQ_ARTICLE="'+stArticle+'"'+ RecupDepot('GQ',stDepot) + ' AND GQ_CLOTURE="-"',true,-1,'',true);
  if not QDispo.Eof then
     begin
     TobDispo.SelectDB('',QDispo);
     Ferme(QDispo);
     end else
     begin //pb : article mouvementé non présent dans dispo
     Ferme(QDispo);
     TobDispo.InitValeurs;
     QDispo:=OpenSQL('SELECT' + stSelect + 'FROM ARTICLE WHERE GA_ARTICLE="'+stArticle+'"',true,-1,'',true);
     if not QDispo.Eof then TobDispo.SelectDB('',QDispo);
     Ferme(QDispo);
     TobDispo.PutValue('GQ_ARTICLE',stArticle);
     TobDispo.PutValue('GQ_DEPOT',stDepot);
     TobDispo.PutValue('GQ_CLOTURE','-');
     end;
end;

function TFStockAjust.ChargeLigne(stArticle,stDepot : string; DateInv : Tdatetime) : TOB;

  procedure AddChampsSup ( TOBD,TOBV : TOB);
  begin
    TOBD.AddChampSupValeur('GL_ARTICLE',TOBV.GetString('GL_ARTICLE'));
    TOBD.AddChampSupValeur('GL_DEPOT',TOBV.GetString('GL_DEPOT'));
    TOBD.AddChampSupValeur('GL_QUALIFQTESTO',TOBV.GetString('GL_QUALIFQTESTO'));
    TOBD.AddChampSupValeur('GL_COEFCONVQTE',TOBV.GetDouble('GL_COEFCONVQTE'));
    TOBD.AddChampSupValeur('GL_COEFCONVQTEVTE',TOBV.GetDouble('GL_COEFCONVQTEVTE'));
    TOBD.AddChampSupValeur('GL_QUALIFQTEACH',TOBV.GetString('GL_QUALIFQTEACH'));
    TOBD.AddChampSupValeur('GL_QUALIFQTEVTE',TOBV.GetString('GL_QUALIFQTEVTE'));
    TOBD.AddChampSupValeur('GL_NATUREPIECEG',TOBV.GetString('GL_NATUREPIECEG'));
    TOBD.AddChampSupValeur('GL_DATEPIECE',TOBV.GetDateTime('GL_DATEPIECE'));
    TOBD.AddChampSupValeur('GL_SOUCHE',TOBV.GetString('GL_SOUCHE'));
    TOBD.AddChampSupValeur('GL_NUMERO',TOBV.GetInteger('GL_NUMERO'));
    TOBD.AddChampSupValeur('GL_INDICEG',TOBV.GetInteger('GL_INDICEG'));
    TOBD.AddChampSupValeur('GL_NUMLIGNE',TOBV.GetInteger('GL_NUMLIGNE'));
    TOBD.AddChampSupValeur('GL_PIECEPRECEDENTE',TOBV.GetString('GL_PIECEPRECEDENTE'));
    TOBD.AddChampSupValeur('GL_PIECEORIGINE',TOBV.GetString('GL_PIECEORIGINE'));
    TOBD.AddChampSupValeur('GL_QTESTOCK',0.0);
    TOBD.AddChampSupValeur('GL_QTERESTE',0.0);
    //--- GUINIER ---
    TOBD.AddChampSupValeur('GL_MTRESTE',0.0);
    TOBD.AddChampSupValeur('GL_VIVANTE',TOBV.GetString('GL_VIVANTE'));
    TOBD.AddChampSupValeur('GL_DPA',TOBV.GetDouble('GL_DPA'));
    TOBD.AddChampSupValeur('GL_DPR',TOBV.GetDouble('GL_DPR'));
    TOBD.AddChampSupValeur('GL_IDENTIFIANTWOL',TOBV.GetInteger('GL_IDENTIFIANTWOL'));
    TOBD.AddChampSupValeur('GL_AFFAIRE',TOBV.GetString('GL_AFFAIRE'));
  end;

var stSQL, stChampLig,stChampOUV, stWhere : string;
    TobLigne, TobLigneNomen : TOB;
    TOBNew, TOBLN : TOB;
    i_ind : integer;
    ParentQte : double;
    OK_MtReliquat : Boolean;
begin
  TobLigne := TOB.Create('',nil,-1);
  stWhere := RecupWhereNaturePiece('LC;LF;RC;RF;PRE;PHY','GL') + ' AND GL_NATUREPIECEG <> "INV" ';
   if stWhere <> '' then stWhere := ' AND ' + stWhere;

  stChampLig := 'GL_ARTICLE,GL_DEPOT,GL_QUALIFQTESTO,GL_COEFCONVQTE,GL_COEFCONVQTEVTE,'+
            'GL_QUALIFQTEACH,GL_QUALIFQTEVTE,GL_NATUREPIECEG,GL_DATEPIECE,GL_SOUCHE,GL_NUMERO,'+
            'GL_INDICEG,GL_NUMLIGNE,GL_PIECEPRECEDENTE,GL_PIECEORIGINE,GL_QTESTOCK,GL_QTERESTE,'+
            'GL_MTRESTE, ' +//--- GUINIER ---
            'GL_VIVANTE,GL_DPA,GL_DPR,GL_IDENTIFIANTWOL,GL_AFFAIRE,GL_DATECREATION';

  stChampOUV := 'GL_ARTICLE,GL_DEPOT,GL_QUALIFQTESTO,BLO_COEFCONVQTE AS GL_COEFCONVQTE,BLO_COEFCONVQTEVTE AS GL_COEFCONVQTEVTE,'+
            'BLO_QUALIFQTEACH AS GL_QUALIFQTEACH,BLO_QUALIFQTEVTE AS GL_QUALIFQTEVTE,'+
            'GL_NATUREPIECEG,GL_DATEPIECE,GL_SOUCHE,GL_NUMERO,'+
            'GL_INDICEG,GL_NUMLIGNE,GL_PIECEPRECEDENTE,GL_PIECEORIGINE,GL_QTESTOCK,GL_QTERESTE,'+
             'GL_MTRESTE, ' +//--- GUINIER ---
            'GL_VIVANTE,BLO_DPA AS GL_DPA,BLO_DPR AS GL_DPR,GL_IDENTIFIANTWOL,GL_AFFAIRE,'+
            'BLO_QTEFACT,BLO_N1,BLO_N2,BLO_N3,BLO_N4,BLO_N5,GL_DATECREATION';

  stSQL := 'SELECT ' + stChampLig + ' FROM LIGNE WHERE GL_ARTICLE="'+stArticle+'" AND '+
            'GL_TYPELIGNE<>"CEN" AND '+
            'GL_DEPOT="'+stDepot+'" AND '+
            'GL_TENUESTOCK="X" AND GL_QUALIFMVT<>"ANN" '+ stWhere +' AND '+
            'GL_DATEPIECE >= "'+USDATETIME(DateInv) +'" '+
         'ORDER BY GL_DATEPIECE';
  TobLigne.LoadDetailFromSQL(stSQL);

  ///Nomenclatures..
  TobLigneNomen := TOB.Create('',nil,-1);
  stSQL := 'SELECT '+stChampOuv+' '+
          'FROM LIGNEOUV '+
          'LEFT JOIN LIGNE ON (GL_NATUREPIECEG=BLO_NATUREPIECEG AND GL_SOUCHE=BLO_SOUCHE AND GL_NUMERO=BLO_NUMERO AND '+
          'GL_INDICEG=BLO_INDICEG AND GL_NUMLIGNE=BLO_NUMLIGNE) '+
          'WHERE '+
          'BLO_ARTICLE="'+stArticle+'" AND BLO_TENUESTOCK="X" AND GL_QUALIFMVT<>"ANN" AND '+
          'GL_DEPOT="'+stDepot+'" '+ stWhere +' AND '+
          'GL_DATEPIECE >= "'+USDATETIME(DateInv) +'" AND '+
          'GL_NATUREPIECEG IN (SELECT GPP_NATUREPIECEG FROM PARPIECE WHERE GPP_STOCKSSDETAIL="X") '+
         'ORDER BY GL_DATEPIECE';
  TobLigneNomen.LoadDetailFromSQL(stSQL);

  /////Concaténation de tobLigne et tobLignenomen
  for i_ind := 0 to TobLigneNomen.Detail.Count -1 do
  begin
  TOBLN := TobLigneNomen.Detail[i_ind];
    //--- GUINIER ---
    Ok_MtReliquat := CtrlOkReliquat(TOBLN, 'GL');
    if Ok_MtReliquat then
      if TOBLN.GetDouble('GL_MTRESTE')=0  then continue
    else
    if TOBLN.GetDouble('GL_QTERESTE')=0 then continue;
    //
  with TOBLN do
      TOBNew := TOBLigne.FindFirst(['GL_NATUREPIECEG', 'GL_DATEPIECE', 'GL_SOUCHE', 'GL_NUMERO', 'GL_INDICEG','GL_NUMLIGNE'],
                                   [GetValue('NATUREPIECEG'), GetValue('DATEPIECE'), GetValue('SOUCHE'),
                                    GetValue('NUMERO'), GetValue('INDICEG'), GetValue('NUMLIGNE')],false);
  if TOBNew = nil then
    begin
      TOBNew := TOB.Create('', TOBLigne, -1);
      AddChampsSup (TOBNew,TOBLN);
    end;

    ParentQte := GetQteDetail(TOBLN, 'GL');
    TOBNew.PutValue('GL_QTESTOCK', TOBNew.GetValue('GL_QTESTOCK') + (TOBLN.GetDouble('GL_QTESTOCK') * ParentQte) );
    TOBNew.PutValue('GL_QTERESTE', TOBNew.GetValue('GL_QTERESTE') + (TOBLN.GetDouble('GL_QTERESTE') * ParentQte) );

    if Ok_MtReliquat then TOBNew.PutValue('GL_MtRESTE', TOBNew.GetValue('GL_QTERESTE') * TOBNew.GetValue('GL_PUHTDEV'));


  end;
  TobLigneNomen.Free;
  Result := TobLigne;
end;

//Rech des natures de pièces affectantes la quantité TypeQte
function TFStockAjust.GetPlusMoins(TypeQte,NaturePiece : string) : integer;
begin
  result := 0;
  if Pos(TypeQte,GetInfoParpiece(NaturePiece,'GPP_QTEPLUS'))> 0 then result := 1 else
  if Pos(TypeQte,GetInfoParpiece(NaturePiece,'GPP_QTEMOINS'))> 0 then result := -1;
end;

function TFStockAjust.ConstitueWhereFather(TOBLigneEtNomen : TOB) : string;
var level : Integer;
begin
  Result := '';
  // Recherche su niveau actuel
  For level := 5 downto 1 do
  begin
    if Level = 5 then
    begin
      if TOBLigneEtNomen.GetInteger('BLO_N5') <> 0 then break;
    end else if Level = 4 then
    begin
      if TOBLigneEtNomen.GetInteger('BLO_N4') <> 0 then break;
    end else if Level = 3 then
    begin
      if TOBLigneEtNomen.GetInteger('BLO_N3') <> 0 then break;
    end else if Level = 2 then
    begin
      if TOBLigneEtNomen.GetInteger('BLO_N2') <> 0 then break;
    end else if Level = 1 then exit;
  end;

  if level > 1 then
  begin
    if level = 5 then
    begin
      Result := ' AND BLO_N5=0'+
                ' AND BLO_N4='+TOBLigneEtNomen.GetString('BLO_N4')+
                ' AND BLO_N3='+TOBLigneEtNomen.GetString('BLO_N3')+
                ' AND BLO_N2='+TOBLigneEtNomen.GetString('BLO_N2')+
                ' AND BLO_N1='+TOBLigneEtNomen.GetString('BLO_N1');
    end else if level = 4 then
    begin
      Result := ' AND BLO_N4=0'+
                ' AND BLO_N3='+TOBLigneEtNomen.GetString('BLO_N3')+
                ' AND BLO_N2='+TOBLigneEtNomen.GetString('BLO_N2')+
                ' AND BLO_N1='+TOBLigneEtNomen.GetString('BLO_N1');
    end else if level = 3 then
    begin
      Result := ' AND BLO_N3=0'+
                ' AND BLO_N2='+TOBLigneEtNomen.GetString('BLO_N2')+
                ' AND BLO_N1='+TOBLigneEtNomen.GetString('BLO_N1');
    end else if level = 2 then
    begin
      Result := ' AND BLO_N2=0 AND BLO_N1='+TOBLigneEtNomen.GetString('BLO_N1');
    end;
  end;
end;


function TFStockAjust.GetQteDetail(TOBLigneEtNomen : TOB; Prefixe : String) : Double;
var Q : TQuery;
    TOBParent : TOB;
    StSqlNivPere : string;
    Ok_MtReliquat : Boolean;
begin

  Ok_MtReliquat := CtrlOkReliquat(TOBLigneEtNomen, Prefixe);

  if Ok_MtReliquat then
    result := TOBLigneEtNomen.getDouble('BLO_MONTANTHTDEV')
  else
  result := TOBLigneEtNomen.getDouble('BLO_QTEFACT');

  StSqlNivPere := ConstitueWhereFather(TOBLigneEtNomen);
  if StSqlNivPere = '' then exit;
  Q := OpenSQL('SELECT BLO_NATUREPIECEG AS GL_NATUREPIECEG,BLO_SOUCHE AS GL_SOUCHE,BLO_NUMERO AS GL_NUMERO,'+
              'BLO_INDICEG AS GL_INDICEG ,BLO_QTEFACT,BLO_NUMLIGNE AS GL_NUMLIGNE,'+
              'BLO_N1,BLO_N2,BLO_N3,BLO_N4,BLO_N5 '+
              'FROM LIGNEOUV '+
              'WHERE '+
              'BLO_NATUREPIECEG="'+TOBLigneEtNomen.GetValue('GL_NATUREPIECEG')+'" AND '+
              'BLO_SOUCHE="'+TOBLigneEtNomen.GetValue('GL_SOUCHE')+'" AND '+
              'BLO_NUMERO='+IntToStr(TOBLigneEtNomen.GetValue('GL_NUMERO'))+' AND '+
              'BLO_INDICEG='+IntToStr(TOBLigneEtNomen.GetValue('GL_INDICEG'))+' AND '+
              'BLO_NUMLIGNE='+IntToStr(TOBLigneEtNomen.GetValue('GL_NUMLIGNE'))+
               StSqlNivPere, true,-1,'',true);
  if Not Q.eof then
  begin
    TOBParent := TOB.Create('LIGNE', nil, -1);
    TOBParent.SelectDB('', Q);
    if Ok_MtReliquat then
      Result := result + GetQteDetail(TOBParent, 'BLO')
    else
      Result := result * GetQteDetail(TOBParent, 'BLO');
    TOBParent.Free;
  end;
  Ferme(Q);
end;

function TFStockAjust.AjustQteStock(TobLigne : TOB ; TypeQte : string) : double;

  procedure EcritOnLog(OneTob : TOB; Qte : double);
  var Amess : string;
      II : integer;
  begin
    (*
    Amess := format ('Date Mvt : %s Date Creation : %s NaturePiece : %3s Numero %5d Qte : %8.3f Qte mouv : %8.3f piece Origine : %s',
                    [
                    DateToStr(OneTob.GetDateTime('GL_DATEPIECE')),
                    DateToStr(OneTob.GetDateTime('GL_DATECREATION')),
                    OneTob.GetString('GL_NATUREPIECEG'),
                    OneTob.Getinteger('GL_NUMERO'),
                    OneTob.GetDouble('GL_QTERESTE'),
                    Qte,
                    OneTOB.getString('GL_PIECEORIGINE')
                    ]);
    ecritLog('C:\pgi01\',Amess,'MVTSTOCK.Txt');
    *)
  end;

var itob, SensMvt : integer;
    Qte,Ratio : double;
    TOBL  : TOB;
    NaturePiece,NumeroP,PieceOrigine,PiecePrecedente : string;
begin
    Result := 0;
    for itob := 0 to TobLigne.Detail.Count -1 do
    begin
        TOBL := TobLigne.Detail[itob];
      NaturePiece := TOBL.GetString('GL_NATUREPIECEG');
      NumeroP := TOBL.getString('GL_NUMERO');
      SensMvt := GetPlusMoins(TypeQte,TOBL.GetValue('GL_NATUREPIECEG'));
      PieceOrigine := TOBL.GetString('GL_PIECEORIGINE');
      PiecePrecedente := TOBL.GetString('GL_PIECEPRECEDENTE');
      if SensMvt = 0 then Continue;
      //
        if Pos (TOBL.GetString('GL_NATUREPIECEG'),'LBT;FBC;BLC')>0 then
        begin
        if IsLivChantier(PiecePrecedente,Pieceorigine) then continue;
        end;
      if Pos (TOBL.GetString('GL_NATUREPIECEG'),'BLF;FF;LFR;')>0 then
        begin
          // On ne prend pas en compte les elments définis pour etre livré directement sur chantier
          if (TOBL.getInteger('GL_IDENTIFIANTWOL')=-1) AND (TOBL.GetString('GL_AFFAIRE')<>'') then continue;
        end;
      if Pos (TOBL.GetString('GL_NATUREPIECEG'),'CF;CFR;CBT')>0 then
      begin
        // On ne prend pas en compte les elments définis pour etre livré directement sur chantier
        if (TOBL.getInteger('GL_IDENTIFIANTWOL')=-1) AND (TOBL.GetString('GL_AFFAIRE')<>'') then continue;
      end;
        Qte := 0;
      Qte := (SensMvt * TOBL.GetValue('GL_QTERESTE'));
        Ratio:=GetRatio (TOBL,nil, trsStock);
        Qte:=Arrondi(Qte/Ratio,V_PGI.OkDecQ) ;
      if (TypeQTe = 'PHY') then EcritOnLog (TOBL,Qte);
        Result := Result + Qte;
    end;
end;

//Gestion de la tob recap
procedure TFStockAjust.ChargeTobRecap(TOBTT : TOB ; stQte : string ; OldQte,NewQte : double);
var TobR : TOB;
begin
TOBR:=TOB.Create('GCTMPAJUSTSTOCK',TobRecap,-1);
TOBR.InitValeurs;
TOBR.PutValue('GZE_UTILISATEUR',V_PGI.USer);
TOBR.PutValue('GZE_COMPTEUR',i_cpte);
TOBR.PutValue('GZE_ARTICLE',TOBTT.GetValue('GQ_ARTICLE'));
TOBR.PutValue('GZE_DEPOT',TOBTT.GetValue('GQ_DEPOT'));
TOBR.PutValue('GZE_RUBRIQUE',stQte);
TOBR.PutValue('GZE_OLDQTE',OldQte);
TOBR.PutValue('GZE_NEWQTE',NewQte);
TOBR.PutValue('GZE_NUMEROLOT','');
if TOBTT.GetValue('GA_STATUTART') = 'DIM' then
   TOBR.PutValue('GZE_DIM',RechDim(TOBTT));
TOBR.AddChampSup('LIBELLE',false);
TOBR.AddChampSup('LOT',false);
TOBR.PutValue('LOT',TOBTT.GetValue('GA_LOT'));
TOBR.PutValue('LIBELLE',TOBTT.GetValue('GA_LIBELLE'));

i_cpte:=i_cpte+1;
end;

function TFStockAjust.RechDim(TobDispoArt : TOB) : string;
var i_indDim : integer;
    GrilleDim,CodeDim,LibDim,StDim : string;
begin
StDim:='';
for i_indDim := 1 to MaxDimension do
    begin
    GrilleDim := TobDispoArt.GetValue ('GA_GRILLEDIM' + IntToStr (i_indDim));
    CodeDim := TobDispoArt.GetValue ('GA_CODEDIM' + IntToStr (i_indDim));
    LibDim := GCGetCodeDim (GrilleDim, CodeDim, i_indDim);
    if LibDim <> '' then
       if StDim='' then StDim:=StDim + LibDim
       else StDim := StDim + ' - ' + LibDim;
    end;
Result := stDim;
end;

procedure TFStockAjust.RechercheLot;
var TobR : TOB;
    stArt, stDep : string;
begin
stArt:=''; stDep:='';
TobR := TobRecap.FindFirst(['LOT'],['X'],false);
While TobR <> nil do
    begin
    if (TobR.GetValue('GZE_ARTICLE')<>stArt) or (TobR.GetValue('GZE_DEPOT')<>stDep) then
       AjouteTobLot(TobR);
    stArt:=TobR.GetValue('GZE_ARTICLE');
    stDep:=TobR.GetValue('GZE_DEPOT');
    TobR := TobRecap.FindNext(['LOT'],['X'],false);
    end;
end;


procedure TFStockAjust.AjouteTobLot(TOBR : TOB);
var QLot : TQuery;
begin
QLot := OpenSQL('SELECT * FROM DISPOLOT WHERE GQL_ARTICLE="'
                +TOBR.GetValue('GZE_ARTICLE')+'" AND GQL_DEPOT="'
                +TOBR.GetValue('GZE_DEPOT')  +'"',true,-1,'',true);
if not QLot.Eof then TOBR.LoadDetailDB('DISPOLOT','','',QLot,false);
Ferme(QLot);
if TOBR.Somme('GQL_PHYSIQUE',[''],[''],true) = TOBR.GetValue('GZE_NEWQTE') then
   begin  //Qté dans dispolot correcte donc pas de réajustement des lots
   TOBR.ClearDetail;
   TOBR.PutValue('LOT','-');
   end;
end;

////////////////////////////////////////////////////////////////////////////////
////MAJ des tables Dispo, Dispolot
procedure TFStockAjust.DebutTransaction;
var ioerr : TIOErr ;
begin
ioerr := Transactions(MAJTableDispo,2);
if ioerr <> oeOk then
   begin
   ModifDispo:=False;
   Msg.Execute (2,Caption,'');
   NoteEvenement('Erreur - Opération annulée','ERR');
   end else
   begin
   ModifDispo:=True;
   NoteEvenement('Le traitement s''est correctement effectué','OK');
   Msg.Execute (4,Caption,'');
   end;
end;

procedure TFStockAjust.MAJTableDispo; //Maj Dispo, DispoLot
var i_lot : integer;
    TOBR, TobLot : TOB;
begin
TOBDispoMAJ.InsertOrUpdateDB(true);
if TobRecap.Detail.Count<=0 then exit;
TobR := TobRecap.FindFirst(['LOT'],['X'],false);
if TobR = nil then exit;
TobLot:=TOB.Create('',nil,-1);
While TobR <> nil do
    begin
    ExecuteSQL('DELETE FROM DISPOLOT WHERE GQL_ARTICLE="'+ TOBR.GetValue('GZE_ARTICLE') +
                              '" AND GQL_DEPOT="' + TOBR.GetValue('GZE_DEPOT') + '"');
    for i_lot := TOBR.Detail.Count -1 downto 0 do
           TOBR.Detail[i_lot].ChangeParent(TobLot,-1);
    TobR := TobRecap.FindNext(['LOT'],['X'],false);
    end;
TobLot.SetAllModifie(True);
TobLot.InsertDB(nil);
TobLot.Free;
end;

////////////////////////////////////////////////////////////////////////////////
////Enreg du résultat dans le jal d'événements
Procedure TFStockAjust.NoteEvenement(Evenement, Etat : string);
var TobJNL : TOB;
    Mess : TStringList;
    Indice : integer;
    QIndice : TQuery;
begin
//exit;
Indice := 0;
Mess := TStringList.Create;
Mess.Clear;
Mess.Add(PTITRE.Caption);
Mess.Add('Traitement effectué pour le dépôt  : '+Depot.Text);
If RefArticle <>'' then Mess.Add('Vérification unitaire pour l''article référence : '+RefArticle);
Mess.Add(Evenement);
QIndice := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT',true,-1,'',true);
if not QIndice.Eof then
     Indice := QIndice.Fields[0].AsInteger + 1;
Ferme(QIndice);

TobJNL := TOB.Create('JNALEVENT',nil,-1);

TobJNL.PutValue('GEV_NUMEVENT',Indice);
TobJNL.PutValue('GEV_TYPEEVENT','STK');
TobJNL.PutValue('GEV_LIBELLE',PTITRE.Caption);
TobJNL.PutValue('GEV_DATEEVENT',V_PGI.DateEntree);
TobJNL.PutValue('GEV_UTILISATEUR',V_PGI.USer);
TobJNL.PutValue('GEV_ETATEVENT',Etat);
TobJNL.PutValue('GEV_BLOCNOTE', Mess.Text);
TobJNL.InsertDB(nil);
TobJNL.free;
Mess.Free;
end;

end.
