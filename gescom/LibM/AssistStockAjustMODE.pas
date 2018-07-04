unit AssistStockAjustMODE;

interface

uses
   Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
   assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls, HEnt1,
   HPanel, UTob, dbTables, HStatus, InvUtil, StockUtil, Grids, UTofAjustStock, EntGC;

procedure EntreeStockAjustMODE;

type
  TFStockAjustMODE = class(TFAssist)
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
    GestionNomen: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure bAnnulerClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);

  private
    { Déclarations privées }

    //Traitement de MAJ // retourne false si annulation par l'utilisateur
    function TraitementOK : boolean;

    //fonctions de traitement des données, recherche des <>
    function  AjustQteStock(TypeQte : string) : double; //Rech de ces pièces dans les lignes de doc
    function  GetParentQte(TOBLigneEtNomen : TOB) : Double;
    function  RechDerniereCloture (Depot : string) : TDateTime;
    procedure MAJQuantite (TOBTT : TOB); //MAJ des qtés dans la tob
    function RecupValStockDerniereCloture (Depot : string; Article : string; Champ : string) : double ;
    function  GetPlusMoins(TypeQte,NaturePiece : string) : integer;
    procedure OuvreQLigne(DatClot,DatDerniereClot : TDateTime ; Depot : string);
    procedure ChargeLignes(Article : string);

    //MAJ des quantités dans la TOBDispo
    procedure DebutTransaction;
    procedure MAJTableDispo;

    //Enreg du résultat dans le jal d'événements
    Procedure NoteEvenement(Evenement, Etat : string);

    //Gestion de la tob récap
    procedure ChargeTobRecap(TOBTT : TOB ; stQte : string ; OldQte,NewQte : double);
    procedure ConstruitTobRecap;
    procedure RechercheLot; //Recherche dans tob recap tous les articles gérés par lot
    procedure AjouteTobLot(TOBR : TOB); //Ajout de tob lot à toutes les tob recap dont l'article est géré par lot
                                        //si les quantités diffèrent


  public
    { Déclarations publiques }
  StopTT, TTdebut : boolean;
  i_cpte : integer;
  TobVerif : TOB;
  TobDispoMAJ, TobRecap : TOB;
  TobLigne,TobLigneNomen : TOB;
  QLigne, QligneNomen : TQuery;
  end;

var
  FStockAjustMODE: TFStockAjustMODE;


implementation

{$R *.DFM}
procedure EntreeStockAjustMODE;
var Fo_Ajust : TFStockAjustMODE;
begin
     Fo_Ajust := TFStockAjustMODE.Create (Application);
     Try
         Fo_Ajust.ShowModal;
     Finally
         Fo_Ajust.free;
     End;
end;

procedure TFStockAjustMODE.FormShow(Sender: TObject);
begin
  inherited;
StopTT:=false;
Depot.ItemIndex:=0;
end;


////////////////////////////////////////////////////////////////////////////////
// Récupération de la valeur d'une zone de la table DISPO à la clôture
////////////////////////////////////////////////////////////////////////////////
function TFStockAjustMODE.RecupValStockDerniereCloture (Depot : string; Article : string; Champ : string) : double ;
var DateCloture : TDateTime;
    Q: TQUERY;
begin
  //
  // Récupération de la date de dernière clôture
  //
  result      := 0 ;
  DateCloture := RechDerniereCloture (Depot);

  if ExisteSQL('SELECT GQ_CLOTURE FROM DISPO WHERE GQ_ARTICLE="'+Article+'" AND GQ_DEPOT="'+Depot+'" AND GQ_CLOTURE="X"') then
  begin
    Q := OpenSQL('SELECT GQ_DATECLOTURE, ' + Champ + ' AS QTE FROM DISPO '+
                 'WHERE GQ_ARTICLE="'+Article+'" AND GQ_DEPOT="'+Depot+'" AND GQ_CLOTURE="X" AND GQ_DATECLOTURE = "'+USDateTime(DateCloture)+'" '+
                 'ORDER BY GQ_DATECLOTURE DESC', true); // La 1ere réponse est la date la + récente
    if not Q.EOF then
    begin
      result := Q.FindField('QTE').AsFloat;
    end;
    Ferme(Q);
  end;
end;


////////////////////////////////////////////////////////////////////////////////
//*************************Evénements de la form******************************//
////////////////////////////////////////////////////////////////////////////////

procedure TFStockAjustMODE.bSuivantClick(Sender: TObject);
begin
  inherited;
if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
LBX_Param.Clear;
LBX_Param.Items.Add('');
LBX_Param.Items.Add('Dépôt : ' + Depot.Text);
end;

procedure TFStockAjustMODE.bPrecedentClick(Sender: TObject);
begin
  inherited;
if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
end;

procedure TFStockAjustMODE.bFinClick(Sender: TObject);
begin
inherited;
bfin.Enabled:=false;
TobDispoMAJ:=TOB.Create('DISPO',nil,-1);
TobVerif:=TOB.Create('',nil,-1);
TobLigne  := TOB.Create('', nil, -1);
TobLigneNomen  := TOB.Create('', nil, -1);

ConstruitTobRecap;

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
TTDebut:=false;
TobDispoMAJ.Free; TobVerif.Free;
TobRecap.Free;
TobLigne.Free; TobLigneNomen.Free;
bfin.Enabled:=true;
end;

procedure TFStockAjustMODE.bAnnulerClick(Sender: TObject);
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
function TFStockAjustMODE.TraitementOK : boolean;
var max : integer;
    QSelect, QMax, QVerif : TQuery;
    TobDispo : TOB;
    DerniereCloture : TDateTime;
    stDepot,stWhere : string;
begin
  Result  := true;
  stDepot := '';
  Max     := 0;

  //
  // Traitement lancé poiur l'instant sur un seul dépot
  //
  if Depot.value <> '' then
  begin
    stWhere := ' AND GQ_DEPOT="' + Depot.Value + '"';
  end else
  begin
    stWhere :='';
    stDepot := TraduireMemoire ('Tous') ;
    PGIBox ('Ce traitement doit être lancé sur un seul dépôt', '');
    exit;
  end;
  //
  // Recherche des fiches stocks à contrôler
  //
  QSelect := OpenSQL('SELECT * FROM DISPO WHERE GQ_CLOTURE="-" ' + stWhere + ' ORDER BY GQ_DEPOT,GQ_ARTICLE', true);

  if QSelect.Eof then
  begin
    Ferme(QSelect);
    exit;
  end;
  //
  // Calcul du nombre total d'articles à contrôler
  //
  QMax := OpenSQL('SELECT COUNT(GQ_ARTICLE) AS NBART FROM DISPO WHERE GQ_CLOTURE="-" ' + StWhere , true);
  if not QMax.Eof then Max := QMax.FindField('NBART').AsInteger;
  Ferme(QMax);

  TobDispo := TOB.Create('DISPO',nil,-1);
  TTdebut:=true;
  DerniereCloture := RechDerniereCloture (Depot.Value);
  InitMove(Max,'');
  QSelect.FindFirst;

  while not QSelect.Eof do
  begin
    //
    // Traitement par dépôt : chargement des lignes de document devant impacter le stock
    //
    if stDepot <> QSelect.FindField('GQ_DEPOT').AsString then
    begin
      stDepot := QSelect.FindField('GQ_DEPOT').AsString;
      OuvreQLigne(iDate2099,DerniereCloture,stDepot);
      TobVerif.ClearDetail;
      QVerif := OpenSQL('SELECT GQ_ARTICLE FROM DISPO WHERE GQ_CLOTURE="-"' + StWhere + ' ORDER BY GQ_ARTICLE', true);
      if not QVerif.Eof then TobVerif.LoadDetailDB('','','',QVerif,false);
      Ferme(QVerif);
    end;
    //
    // Arrêt du traitement ?
    if StopTT then
    begin
      Result := false;
      break;
    end;
    //
    // Traitement d'un article
    //
    TobDispo.SelectDB('',QSelect);
    ChargeLignes(TobDispo.GetValue('GQ_ARTICLE'));
    MoveCur(False);
    MAJQuantite (TobDispo);
    //
    // Article suivant
    //
    QSelect.Next;
    //
    // On sort sui on change de dépôt
    //
    if stDepot <> QSelect.FindField('GQ_DEPOT').AsString then
    begin
      Ferme(QLigne);
      if GestionNomen.checked then Ferme(QLigneNomen);
    end;
  end;
  FiniMove;
  TobDispo.Free;
  Ferme(QSelect);
  Ferme(QLigne);
  if GestionNomen.checked then Ferme(QLigneNomen);
end;


////////Dernière clôture
function TFStockAjustMODE.RechDerniereCloture (Depot : string) : TDateTime;
var QLastDate : TQuery;
begin
  result := 0;
  QLastDate := OpenSQL('SELECT MAX(GQ_DATECLOTURE) AS DATECLOTURE FROM DISPO WHERE GQ_DEPOT="'+Depot+'" AND GQ_CLOTURE="X"', true);
  if not QLastDate.Eof then Result := QLastDate.FindField('DATECLOTURE').AsDateTime;
  Ferme(QLastDate);
  if result <= 0 then result := 02;
end;

////////MAJ des quantités dans la TOB Dispo
procedure TFStockAjustMODE.MAJQuantite (TOBTT : TOB);
var  TQte : TQtePrixRec;
     i_ind : integer;
     Change : boolean;
     TobTemp, TobNew, TobFilleDispo : TOB;
     TobLPhys,TobLNPhys,TobL,TobLN : TOB;
     QtePhys,ParentQte,LivreCli, LivreFou, ReserveCli, ReserveFou : double;
     QteVFO, QteTRA, QteES , QteINV : double;
     StArt,stDep : string;
     QtePlus,QteMoins : string ;
begin
  Change := false;
  //
  // Dupliqcation de la TOB DISPO pour traitement
  //
  TobTemp := TOB.Create('DISPO',nil,-1);
  TobTemp.Dupliquer(TOBTT,false,true);

  QtePhys := TOBTT.GetValue('GQ_PHYSIQUE');
  stArt   := TOBTT.GetValue('GQ_ARTICLE');
  stDep   := TOBTT.GetValue('GQ_DEPOT');


  /////Récup de tobligne et lignenomen toutes lignes affectant le stock physiq
  TobLPhys  := TOB.Create('',nil,-1);
  TobLNPhys := TOB.Create('',nil,-1);
  for i_ind := 0 to TobLigne.Detail.Count -1 do
  begin
    QtePlus  := VarToStr(TobLigne.Detail[i_ind].GetValue('GPP_QTEPLUS'));     // MODIF LM ORACLE
    QteMoins := VarToStr(TobLigne.Detail[i_ind].GetValue('GPP_QTEMOINS'));    // MODIF LM ORACLE

    if (pos('PHY', QtePlus) > 0) or (pos('PHY',QteMoins) > 0) then
    begin
      TobL := TOB.Create('',TobLPhys,-1);
      TobL.Dupliquer(TobLigne.Detail[i_ind],true,true);
    end ;
  end;

  for i_ind := 0 to TobLigneNomen.Detail.Count -1 do
  begin
    QtePlus  := VarToStr(TobLigne.Detail[i_ind].GetValue('GPP_QTEPLUS'));
    QteMoins := VarToStr(TobLigne.Detail[i_ind].GetValue('GPP_QTEMOINS'));
    if (pos('PHY',QtePlus) > 0) or (pos('PHY',QteMoins) > 0) then
    begin
      TobLN := TOB.Create('',TobLNPhys,-1);
      TobLN.Dupliquer(TobLigneNomen.Detail[i_ind],true,true);
    end;
  end;

  TQte := TQtePrixRec(GetQtePrixDateListe(stArt,stDep,StrToDate('31/12/2099'),TobLPhys,TobLNPhys));
  TobLPhys.Free;
  TobLNPhys.Free;
  if TQte.SomethingReturned then
  begin
    if Arrondi(TQte.Qte,V_PGI.OkDecQ) <> QtePhys then
    begin
      ChargeTobRecap(TOBTT,'Qté Physique',QtePhys,TQte.Qte);
      TobTemp.PutValue('GQ_PHYSIQUE',TQte.Qte);
      TobTemp.PutValue('GQ_DPA',TQte.DPA);
      TobTemp.PutValue('GQ_DPR',TQte.DPR);
      TobTemp.PutValue('GQ_PMAP',TQte.PMAP);
      TobTemp.PutValue('GQ_PMRP',TQte.PMRP);
      Change := true;
    end;
  end else TobTemp.PutValue('GQ_PHYSIQUE',0);

  /////Concaténation de tobLigne et tobLignenomen
  for i_ind := 0 to TobLigneNomen.Detail.Count -1 do
  begin
    TOBLN := TobLigneNomen.Detail[i_ind];
    with TOBLN do
      TOBNew := TOBLigne.FindFirst(['GL_NATUREPIECEG', 'GL_DATEPIECE', 'GL_SOUCHE', 'GL_NUMERO', 'GL_INDICEG', 'GL_NUMLIGNE'],
                                   [GetValue('GL_NATUREPIECEG'), GetValue('GL_DATEPIECE'), GetValue('GL_SOUCHE'), GetValue('GL_NUMERO'), GetValue('GL_INDICEG'), GetValue('GL_NUMLIGNE')],
                                   false);
    if TOBNew = nil then
    begin
      TOBNew := TOB.Create('LIGNE', TOBLigne, -1);
      TOBNew.Dupliquer(TOBLN, true, true);
      TOBNew.PutValue('GL_QTESTOCK', 0);
      TOBNew.PutValue('GL_QTERESTE', 0);
    end;

    ParentQte := GetParentQte(TOBLN);
    TOBNew.PutValue('GL_QTESTOCK', TOBNew.GetValue('GL_QTESTOCK') + (TOBLN.GetValue('GL_QTESTOCK') * ParentQte) );
    TOBNew.PutValue('GL_QTERESTE', TOBNew.GetValue('GL_QTERESTE') + (TOBLN.GetValue('GL_QTERESTE') * ParentQte) );
  end;

  if TobLigne.Detail.Count>0 then
  begin
    LivreCli:=AjustQteStock('LC');
    // Ajout de la quantité initiale à la date de démarrage
    LivreCli := LivreCli + RecupValStockDerniereCloture (stDep, start, 'GQ_LIVRECLIENT') ;

    if LivreCli <> TOBTT.GetValue('GQ_LIVRECLIENT') then
    begin
      ChargeTobRecap(TOBTT,'Livré client',TOBTT.GetValue('GQ_LIVRECLIENT'),LivreCli);
      TobTemp.PutValue('GQ_LIVRECLIENT',LivreCli);
      Change := true;
    end;

    LivreFou:=AjustQteStock('LF');
    // Ajout de la quantité initiale à la date de démarrage
    LivreFou := LivreFou + RecupValStockDerniereCloture (stDep, start, 'GQ_LIVREFOU') ;
    if LivreFou <> TOBTT.GetValue('GQ_LIVREFOU') then
    begin
      ChargeTobRecap(TOBTT,'Livré fournisseur',TOBTT.GetValue('GQ_LIVREFOU'),LivreFou);
      TobTemp.PutValue('GQ_LIVREFOU',LivreFou);
      Change := true;
    end;

    ReserveCli:=AjustQteStock('RC');
    // Ajout de la quantité initiale à la date de démarrage
    ReserveCli := ReserveCli + RecupValStockDerniereCloture (stDep, start, 'GQ_RESERVECLI') ;
    if ReserveCli <> TOBTT.GetValue('GQ_RESERVECLI') then
    begin
      ChargeTobRecap(TOBTT,'Reservé client',TOBTT.GetValue('GQ_RESERVECLI'),ReserveCli);
      TobTemp.PutValue('GQ_RESERVECLI',ReserveCli);
      Change := true;
    end;

    ReserveFou:=AjustQteStock('RF');
    // Ajout de la quantité initiale à la date de démarrage
    ReserveFou := ReserveFou + RecupValStockDerniereCloture (stDep, start, 'GQ_RESERVEFOU') ;
    if ReserveFou <> TOBTT.GetValue('GQ_RESERVEFOU') then
    begin
      ChargeTobRecap(TOBTT,'Reservé fournisseur',TOBTT.GetValue('GQ_RESERVEFOU'),ReserveFou);
      TobTemp.PutValue('GQ_RESERVEFOU',ReserveFou);
      Change := true;
    end;

    QteVFO:=AjustQteStock('VFO');
    // Ajout de la quantité initiale à la date de démarrage
    QteVFO := QteVFO + RecupValStockDerniereCloture (stDep, start, 'GQ_VENTEFFO') ;
    if QteVFO <> TOBTT.GetValue('GQ_VENTEFFO') then
    begin
      ChargeTobRecap(TOBTT,'Ventes détail',TOBTT.GetValue('GQ_VENTEFFO'),QteVFO);
      TobTemp.PutValue('GQ_VENTEFFO',QteVFO);
      Change := true;
    end;

    QteES:=AjustQteStock('ESE');
    // Ajout de la quantité initiale à la date de démarrage
    QteES := QteES + RecupValStockDerniereCloture (stDep, start, 'GQ_ENTREESORTIES') ;
    if QteES <> TOBTT.GetValue('GQ_ENTREESORTIES') then
    begin
      ChargeTobRecap(TOBTT,'E/S exceptionnelles',TOBTT.GetValue('GQ_ENTREESORTIES'),QteES);
      TobTemp.PutValue('GQ_ENTREESORTIES',QteES);
      Change := true;
    end;

    QteTRA:=AjustQteStock('TRA');
    // Ajout de la quantité initiale à la date de démarrage
    QteTRA := QteTRA + RecupValStockDerniereCloture (stDep, start, 'GQ_TRANSFERT') ;
    if QteTRA <> TOBTT.GetValue('GQ_TRANSFERT') then
    begin
      ChargeTobRecap(TOBTT,'Transferts',TOBTT.GetValue('GQ_TRANSFERT'),QteTRA);
      TobTemp.PutValue('GQ_TRANSFERT',QteTRA);
      Change := true;
    end;

    QteINV:=AjustQteStock('INV');
    // Ajout de la quantité initiale à la date de démarrage
    QteINV := QteINV + RecupValStockDerniereCloture (stDep, start, 'GQ_ECARTINV') ;
    if QteINV <> TOBTT.GetValue('GQ_ECARTINV') then
    begin
      ChargeTobRecap(TOBTT,'Inventaires',TOBTT.GetValue('GQ_ECARTINV'),QteTRA);
      TobTemp.PutValue('GQ_ECARTINV',QteINV);
      Change := true;
    end;

  end;

  if Change then
  begin
    TobFilleDispo:=TOB.Create('DISPO',TobDispoMAJ,-1);
    TobFilleDispo.Dupliquer(TobTemp,false,true);
  end;

  TobTemp.Free;
end;

procedure TFStockAjustMODE.OuvreQLigne(DatClot,DatDerniereClot : TDateTime ; Depot : string);
var stChamp, stWhere : string;
begin
if Depot <> '' then stWhere := 'AND GL_DEPOT="'+Depot+'" '
else stWhere := '';
stChamp := 'GL_ARTICLE,GL_DEPOT,GL_QUALIFQTESTO,GL_QUALIFQTEACH,GL_QUALIFQTEVTE,GL_NATUREPIECEG,GL_DATEPIECE,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE,GL_PIECEPRECEDENTE,GL_QTESTOCK,GL_QTERESTE,GL_VIVANTE,GL_DPA,GL_DPR,GPP_VENTEACHAT,GPP_QTEPLUS,GPP_QTEMOINS';
QLigne := OpenSQL('SELECT '+stChamp+
             ' FROM LIGNE LEFT JOIN PARPIECE ON GL_NATUREPIECEG=GPP_NATUREPIECEG '+
                   'WHERE (GPP_QTEMOINS <> "" OR GPP_QTEPLUS <> "") '+
                   'AND GL_ARTICLE<>"" '+
                   'AND GL_TENUESTOCK="X" '+
                   'AND GL_QUALIFMVT<>"ANN" '+ stWhere +
                   'AND GL_DATEPIECE>"'+USDateTime(DatDerniereClot)+'" '+
                   'AND GL_DATEPIECE<="'+USDateTime(DatClot)+'" ORDER BY GL_ARTICLE, GL_DATEPIECE', true);


///Nomenclatures..
if GestionNomen.checked then
begin
  QLigneNomen := OpenSQL('SELECT '+stChamp+',GLN_ARTICLE,GLN_QUALIFQTESTO , GLN_QUALIFQTEACH , GLN_QUALIFQTEVTE , GLN_COMPOSE, GLN_NATUREPIECEG, GLN_SOUCHE,GLN_NUMERO,GLN_INDICEG,GLN_NUMLIGNE,GLN_ORDRECOMPO, GLN_QTE '+
             ' FROM LIGNENOMEN LEFT JOIN LIGNE ON (GL_NATUREPIECEG=GLN_NATUREPIECEG AND GL_SOUCHE=GLN_SOUCHE AND GL_NUMERO=GLN_NUMERO AND GL_INDICEG=GLN_INDICEG AND GL_NUMLIGNE=GLN_NUMLIGNE) '+
             'LEFT JOIN PARPIECE ON GL_NATUREPIECEG=GPP_NATUREPIECEG '+
                   'WHERE (GPP_QTEMOINS <> "" OR GPP_QTEPLUS <> "") '+
                   'AND GLN_ARTICLE<>"" '+
                   'AND GLN_TENUESTOCK="X" '+
                   'AND GL_QUALIFMVT<>"ANN" '+ stWhere +
                   'AND GL_DATEPIECE>"'+USDateTime(DatDerniereClot)+'" '+
                   'AND GL_DATEPIECE<="'+USDateTime(DatClot)+'" ORDER BY GLN_ARTICLE,GL_DATEPIECE', true);
end;
end;

procedure TFStockAjustMODE.ChargeLignes(Article : string);
var TobL,TobLN : TOB;
begin
TobLigne.ClearDetail;
if QLigne.Bof then QLigne.First;
while (QLigne.FindField('GL_ARTICLE').AsString = Article) and
  (not QLigne.Eof) do
   begin
   TobL := TOB.Create('',TobLigne,-1);
   TobL.SelectDB('',QLigne);
   repeat
       QLigne.Next;
   Until (TobVerif.FindFirst(['GQ_ARTICLE'],[QLigne.FindField('GL_ARTICLE').AsString],false)<>nil)
         or (QLigne.Eof);
   end;


TobLigneNomen.ClearDetail;
if GestionNomen.checked then
begin
  if QLigneNomen.Bof then QLigneNomen.First;
  while (QLigneNomen.FindField('GLN_ARTICLE').AsString = Article) and
        (not QLigneNomen.Eof) do
  begin
    TobLN := TOB.Create('',TobLigneNomen,-1);
    TobLN.SelectDB('',QLigneNomen);
    repeat
        QLigneNomen.Next;
    Until (TobVerif.FindFirst(['GQ_ARTICLE'],[QLigneNomen.FindField('GLN_ARTICLE').AsString],false)<>nil)
         or (QLigneNomen.Eof);
  end;
end;
end;

//Rech des natures de pièces affectantes la quantité TypeQte
function TFStockAjustMODE.GetPlusMoins(TypeQte,NaturePiece : string) : integer;
var TOBNat : TOB;
begin
result := 0;
TOBNat := VH_GC.TOBParPiece.FindFirst(['GPP_NATUREPIECEG'], [NaturePiece], false);
if TOBNat = nil then exit;
if Pos(TypeQte, TOBNat.GetValue('GPP_QTEPLUS')) > 0 then result := 1 else
if Pos(TypeQte, TOBNat.GetValue('GPP_QTEMOINS')) > 0 then result := -1;
end;

function TFStockAjustMODE.GetParentQte(TOBLigneEtNomen : TOB) : Double;
var Q : TQuery;
    TOBParent : TOB;
begin
result := TOBLigneEtNomen.GetValue('GLN_QTE');
if (TOBLigneEtNomen.GetValue('GLN_COMPOSE') = null) or
   (TOBLigneEtNomen.GetValue('GLN_COMPOSE') = '') then exit;
Q := OpenSQL('SELECT GLN_NATUREPIECEG,GLN_SOUCHE,GLN_NUMERO,GLN_INDICEG,GLN_NUMLIGNE,GLN_COMPOSE,GLN_ORDRECOMPO,GLN_QTE '+
             'FROM LIGNE LEFT JOIN PARPIECE ON GL_NATUREPIECEG=GPP_NATUREPIECEG '+
                        'LEFT JOIN LIGNENOMEN ON (GL_NATUREPIECEG=GLN_NATUREPIECEG AND GL_SOUCHE=GLN_SOUCHE AND GL_NUMERO=GLN_NUMERO AND GL_INDICEG=GLN_INDICEG AND GL_NUMLIGNE=GLN_NUMLIGNE) '+
             'WHERE GLN_ARTICLE="'+TOBLigneEtNomen.GetValue('GLN_COMPOSE')+'" '+
                   'AND GLN_NUMORDRE='+IntToStr(TOBLigneEtNomen.GetValue('GLN_ORDRECOMPO'))+' '+
                   'AND GLN_NATUREPIECEG="'+TOBLigneEtNomen.GetValue('GLN_NATUREPIECEG')+'" '+
                   'AND GLN_SOUCHE="'+TOBLigneEtNomen.GetValue('GLN_SOUCHE')+'" '+
                   'AND GLN_NUMERO='+IntToStr(TOBLigneEtNomen.GetValue('GLN_NUMERO'))+' '+
                   'AND GLN_INDICEG='+IntToStr(TOBLigneEtNomen.GetValue('GLN_INDICEG'))+' '+
                   'AND GLN_NUMLIGNE='+IntToStr(TOBLigneEtNomen.GetValue('GLN_NUMLIGNE')), true);
TOBParent := TOB.Create('LIGNE', nil, -1);
if not TOBParent.SelectDB('', Q) then begin TOBParent.Free; TOBParent := nil; end;
Ferme(Q);
if TOBParent <> nil then begin result := result * GetParentQte(TOBParent);
                               TOBParent.Free; end;
end;

function TFStockAjustMODE.AjustQteStock(TypeQte : string) : double;
var Cle, S : String;
    itob, fac1, fac2 : integer;
    Qte,Ratio : double;
    TOBL, TOBR, TobRelicat : TOB;
begin
Result := 0;
TobRelicat := TOB.Create('',nil,-1);
for itob := 0 to TobLigne.Detail.Count -1 do
  begin
  TOBL := TobLigne.Detail[itob];

  ///////////////////////////////////////////////////////////////////////////////
  // MODIF LM pour traitement MODE
  ///////////////////////////////////////////////////////////////////////////////
  //
  // Cas  : les commandes fournisseurs, réceptionnées n'impacte plus l'attendu fournisseur
  //
  if (TOBL.GetValue('GL_VIVANTE') <> 'X') AND (TOBL.GetValue ('GL_NATUREPIECEG') = 'CF') then continue;
  //
  // Cas 2 : les réceptions fournisseurs, facturées n'impacte plus le reçu fournisseur
  //         la quantité en stock sera calculée à partir de la facture fournisseur
  //
  if (TOBL.GetValue('GL_VIVANTE') <> 'X') AND (TOBL.GetValue ('GL_NATUREPIECEG') = 'BLF') then continue;
  //
  // Cas 3 : les commandes clients, réceptionnées n'impacte plus le réservé client
  //
  if (TOBL.GetValue('GL_VIVANTE') <> 'X') AND (TOBL.GetValue ('GL_NATUREPIECEG') = 'CC') then continue;
  //
  // Cas 4 : les préparations de livraisons clients, réceptionnées n'impacte plus le réservé client
  //
  if (TOBL.GetValue('GL_VIVANTE') <> 'X') AND (TOBL.GetValue ('GL_NATUREPIECEG') = 'PRE') then continue;
  //
  // Cas 5 : les livraisons clients, facturées n'impacte plus le livré client
  //         la quantité en stock sera calculée à partir de la facture client
  //
  if (TOBL.GetValue('GL_VIVANTE') <> 'X') AND (TOBL.GetValue ('GL_NATUREPIECEG') = 'BLC') then continue;

  fac2 := 1;
  Qte := 0;
  repeat
    fac1 := GetPlusMoins(TypeQte,TOBL.GetValue('GL_NATUREPIECEG'));
    if fac1 = 0 then break;
    Qte := Qte + (fac2 * (fac1 * (TOBL.GetValue('GL_QTESTOCK')-TOBL.GetValue('GL_QTERESTE'))));

    Cle := TOBL.GetValue('GL_PIECEPRECEDENTE');
    S := ReadTokenSt(Cle);
    S := Cle;
    if (TOBL.GetValue('GL_INDICEG') > 0) or (Cle = '')
      then TOBL := nil
      else begin
           TOBL := TobLigne.FindFirst(['GL_NATUREPIECEG', 'GL_SOUCHE', 'GL_NUMERO', 'GL_INDICEG', 'GL_NUMLIGNE'],
                                           [ReadTokenSt(S), ReadTokenSt(S), ReadTokenSt(S), ReadTokenSt(S), ReadTokenSt(S)], false);
           if TOBL = nil then
             begin
             Cle := '"'+ReadTokenSt(Cle)+'";"'+ReadTokenSt(Cle)+'";'+ Cle;
             TobR := TOB.Create('LIGNE',TobRelicat,-1);
             TOBR.SelectDB(Cle, nil);
             TOBL := TOBR;
             end;
           end;
    fac2 := -1;
  until (TOBL = nil);
  TOBL := TobLigne.Detail[itob];
  if TOBL.FieldExists('GLN_QUALIFQTESTO') then
         Ratio:=GetRatio (TOBL,TOBL, trsStock)    //nomenclature
  else   Ratio:=GetRatio (TOBL,nil, trsStock);
  Qte:=Arrondi(Qte/Ratio,V_PGI.OkDecQ) ;
  Result := Result + Qte;
  end;
TobRelicat.Free;
end;

//Gestion de la tob recap
procedure TFStockAjustMODE.ChargeTobRecap(TOBTT : TOB ; stQte : string ; OldQte,NewQte : double);
var TobR : TOB;
begin
TOBR:=TOB.Create('GCTMPAJUSTSTOCK',TobRecap,-1);
TOBR.PutValue('GZE_UTILISATEUR',V_PGI.USer);
TOBR.PutValue('GZE_COMPTEUR',i_cpte);
TOBR.PutValue('GZE_ARTICLE',TOBTT.GetValue('GQ_ARTICLE'));
TOBR.PutValue('GZE_DEPOT',TOBTT.GetValue('GQ_DEPOT'));
TOBR.PutValue('GZE_RUBRIQUE',stQte);
TOBR.PutValue('GZE_OLDQTE',OldQte);
TOBR.PutValue('GZE_NEWQTE',NewQte);
TOBR.PutValue('GZE_NUMEROLOT','');
i_cpte:=i_cpte+1;
end;

procedure TFStockAjustMODE.ConstruitTobRecap;
begin
TobRecap := TOB.Create('GCTMPAJUSTSTOCK',nil,-1);
i_cpte:=0;
end;

procedure TFStockAjustMODE.RechercheLot;
var QSelect : TQuery;
    TobArt, TobTemp : TOB;
    i_ind : integer;
    stArt, stDep : string;
begin
if TobRecap.Detail.Count>0 then TobRecap.Detail[0].AddChampSup('LOT',true);
QSelect:=OpenSQL('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_LOT="X"', false);
if QSelect.Eof then
   begin
   Ferme(QSelect);
   for i_ind:=0 to TobRecap.Detail.Count-1 do
       TobRecap.Detail[i_ind].PutValue('LOT','');
   exit;
   end;
TobArt := TOB.Create('',nil,-1);
TobArt.LoadDetailDB('','','',QSelect,False);
Ferme(QSelect);

stArt:='';
stDep:='';
for i_ind:=0 to TobRecap.Detail.Count-1 do
    begin
    TobRecap.Detail[i_ind].PutValue('LOT','');
    if (TobRecap.Detail[i_ind].GetValue('GZE_ARTICLE')<>stArt) or
       (TobRecap.Detail[i_ind].GetValue('GZE_DEPOT')<>stDep) then
       begin
       TobTemp:=TobArt.FindFirst(['GA_ARTICLE'],[TobRecap.Detail[i_ind].GetValue('GZE_ARTICLE')],true);
       if TobTemp<>nil then AjouteTobLot(TobRecap.Detail[i_ind]);
       end;
    stArt:=TobRecap.Detail[i_ind].GetValue('GZE_ARTICLE');
    stDep:=TobRecap.Detail[i_ind].GetValue('GZE_DEPOT');
    end;
TobArt.Free;
end;


procedure TFStockAjustMODE.AjouteTobLot(TOBR : TOB);
var TobLot : TOB;
    QSelect : TQuery;
    bAjout : boolean;
begin
bAjout:=true;
TobLot := TOB.Create('DISPOLOT',nil,-1);
QSelect:=OpenSQL('SELECT * FROM DISPOLOT WHERE GQL_ARTICLE="'
                +TOBR.GetValue('GZE_ARTICLE')+'" AND GQL_DEPOT="'
                +TOBR.GetValue('GZE_DEPOT')  +'"', false);
if not QSelect.Eof then
   begin
   TobLot.LoadDetailDB('DISPOLOT','','',QSelect,False);
   if TobLot.Somme('GQL_PHYSIQUE',[''],[''],true) = TOBR.GetValue('GZE_NEWQTE') then
      bAjout:=false;
   end;
Ferme(QSelect);
if bAjout then
   begin
   TobLot.ChangeParent(TOBR,-1);
   TOBR.PutValue('LOT','X');
   end else TobLot.free;
end;

////////////////////////////////////////////////////////////////////////////////
////MAJ des tables Dispo, Dispolot
procedure TFStockAjustMODE.DebutTransaction;
var ioerr : TIOErr ;
begin
ioerr := Transactions(MAJTableDispo,2);
if ioerr <> oeOk then
   begin
   Msg.Execute (2,Caption,'');
   NoteEvenement('Erreur - Opération annulée','ERR');
   end else NoteEvenement('Le traitement s''est correctement effectué','OK');
end;

procedure TFStockAjustMODE.MAJTableDispo;
var i_ind1, i_ind2 : integer;
    TOBR, TOBL, TobLot : TOB;
begin
TOBDispoMAJ.InsertOrUpdateDB(true);
if TobRecap.Detail.Count<=0 then exit;
TobLot:=TOB.Create('DISPOLOT',nil,-1);
for i_ind1:=0 to TobRecap.detail.Count -1 do
    begin
    TOBR:=TobRecap.Detail[i_ind1];
    if TOBR.Detail.Count>0 then
       begin
       TOBL:=TOBR.Detail[0];
       ExecuteSQL('DELETE FROM DISPOLOT WHERE GQL_ARTICLE="'+TOBR.GetValue('GZE_ARTICLE')
                              +'" AND GQL_DEPOT="'+TOBR.GetValue('GZE_DEPOT')+'"');
       for i_ind2:=TOBL.Detail.Count -1 downto 0 do
           TOBL.Detail[i_ind2].ChangeParent(TobLot,-1);
       end;
    end;
TobLot.InsertOrUpdateDB(true);
TobLot.Free;
end;

////////////////////////////////////////////////////////////////////////////////
////Enreg du résultat dans le jal d'événements
Procedure TFStockAjustMODE.NoteEvenement(Evenement, Etat : string);
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
Mess.Add(Evenement);
QIndice := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT',true);
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
