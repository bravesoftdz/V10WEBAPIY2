unit MulBds;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, ExtCtrls, HSysMenu, Menus, Db,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Hqry, ComCtrls, HRichEdt,
  Grids, DBGrids, HDB, HTB97, StdCtrls, ColMemo, HEnt1, Hctrls, Mask, Ent1,
  HPanel,UiUtil, HMsgBox,
  HStatus, SaisUtil,
  ZTypes, ZBalance, HRichOLE, ADODB ;

procedure MultiCritereBds(Comment : TActionFiche) ;

type
  TFMulBds = class(TFMul)
    HB_EXERCICE: THValComboBox;
    HLEXERCICE: THLabel;
    HB_DATE1: THCritMaskEdit;
    HLDATE1: THLabel;
    HLFIN: THLabel;
    HB_DATE2: THCritMaskEdit;
    XX_WHERE: TEdit;
    DATE1: THValComboBox;
    DATE2: THValComboBox;
    BDelete: TToolbarButton97;
    HMulBds: THMsgBox;
    procedure HB_EXERCICEChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure DATE1Change(Sender: TObject);
    procedure DATE2Change(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BNouvRechClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
  private
    ObjBalance : TZBalance ;
    memOptions : ROpt ;
    memSaisie  : RDevise ;
    // Fonctions ressource
    function  GetMessageRC(MessageID : Integer) : string ;
    function  PrintMessageRC(MessageID : Integer; StAfter : string='') : Integer ;
    // Fonctions utilitaires
    procedure CreateBal ;
    procedure FreeBal ;
    procedure DelBalance ;
    procedure EnableButtons ;
  public
    { Déclarations publiques }
  end;

implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  SaisBal, UtilPgi;

{$R *.DFM}

const RC_NOBALANCE        =  0 ;
      RC_CONFIRMDEL       =  1 ;
      RC_BADDEL           =  2 ;

//=======================================================
//============= Point d'entrée dans le MUL ==============
//=======================================================
procedure MultiCritereBds(Comment : TActionFiche) ;
var FMulBds : TFMulBds ; PP : THPanel ;
begin
if Comment<>taConsult then if _Blocage(['nrCloture'], FALSE, 'nrAucun') then Exit ;
FMulBds:=TFMulBds.Create(Application) ;
FMulBds.TypeAction:=Comment ;
case Comment of
  taConsult,
  taModif   : begin
              FMulBds.FNomFiltre:='CPMULBDS' ;
              FMulBds.Q.Liste:='CPMULBDSZ' ;
              //FMulGene.HelpContext := 7112000 ;
              end ;
  end ;
PP:=FindInsidePanel ;
if PP=nil then
  begin
   try
    FMulBds.ShowModal ;
   finally
    FMulBds.Free ;
   end ;
  Screen.Cursor:=SyncrDefault ;
  end else
  begin
  InitInside(FMulBds, PP) ;
  FMulBds.Show ;
  end ;
end ;

//=======================================================
//================= Fonctions Ressource =================
//=======================================================
function TFMulBds.GetMessageRC(MessageID : Integer) : string ;
begin
Result:=HMulBds.Mess[MessageID] ;
end ;

function TFMulBds.PrintMessageRC(MessageID : Integer; StAfter : string) : Integer ;
begin
Result:=HMulBds.Execute(MessageID, Caption, StAfter) ;
end ;

//=======================================================
//================ Fonctions Utilitaires ================
//=======================================================
procedure TFMulBds.CreateBal ;
begin
memOptions.TypeBal:='BDS' ;
memOptions.ExoBal:=Q.FindField('HB_EXERCICE').AsString ;
DecodeDate(Q.FindField('HB_DATE1').AsDateTime, memOptions.DebYear, memOptions.DebMonth, memOptions.DebJour) ;
DecodeDate(Q.FindField('HB_DATE2').AsDateTime, memOptions.Year,    memOptions.Month,    memOptions.MaxJour) ;
ObjBalance:=TZBalance.Create(memSaisie, memOptions) ;
end ;

procedure TFMulBds.FreeBal ;
begin
ObjBalance.Free ; ObjBalance:=nil ;
end ;

procedure TFMulBds.DelBalance ;
begin
ObjBalance.Del ;
end ;

procedure TFMulBds.EnableButtons ;
var Bal : TZBalance ; Dev : RDevise ; Opt : ROpt ;
begin
if (HB_EXERCICE.Vide) and (HB_EXERCICE.ItemIndex=0) then
  begin Binsert.Enabled:=FALSE ; Exit ; end ;
Opt.TypeBal:='BDS' ;
Opt.ExoBal:=HB_EXERCICE.Value ;
DecodeDate(StrToDate(HB_DATE1.Text), Opt.DebYear, Opt.DebMonth, Opt.DebJour) ;
DecodeDate(StrToDate(HB_DATE2.Text), Opt.Year,    Opt.Month,    Opt.MaxJour) ;
Bal:=TZBalance.Create(Dev, Opt) ;
if Bal.ExistBds(FALSE) then Binsert.Enabled:=FALSE else Binsert.Enabled:=TRUE ;
Bal.Free ;
end ;

//=======================================================
//================ Evénements de la Form ================
//=======================================================
procedure TFMulBds.FormShow(Sender: TObject);
begin
if VH^.CPExoRef.Code<>'' then
   BEGIN
   HB_EXERCICE.Value:=VH^.CPExoRef.Code ;
   HB_DATE1.Text:=DateToStr(VH^.CPExoRef.Deb) ;
   HB_DATE2.Text:=DateToStr(VH^.CPExoRef.Fin) ;
   END else
   BEGIN
   HB_EXERCICE.Value:=VH^.Entree.Code ;
   HB_DATE1.Text:=DateToStr(VH^.Entree.Deb) ;
   HB_DATE2.Text:=DateToStr(VH^.Entree.Fin) ;
   END ;
inherited ;
EnableButtons ;
end;

procedure TFMulBds.FormClose(Sender: TObject; var Action: TCloseAction);
begin
inherited ;
if Parent is THPanel then Action:=caFree ;
end;

//=======================================================
//================ Evénements des contrôles =============
//=======================================================
procedure TFMulBds.HB_EXERCICEChange(Sender: TObject);
begin
inherited ;
HB_DATE1.InitDate ; HB_DATE2.InitDate ;
if (HB_EXERCICE.Vide) and (HB_EXERCICE.ItemIndex=0) then
  begin (*DATE1.Enabled:=FALSE ;*) DATE2.Enabled:=FALSE ; Exit ; end ;
(*DATE1.Enabled:=TRUE ;*) DATE2.Enabled:=TRUE ;
ListePeriode(HB_EXERCICE.Value, DATE1.Items, DATE1.Values, TRUE) ;
ListePeriode(HB_EXERCICE.Value, DATE2.Items, DATE2.Values, FALSE) ;
ExoToDates(HB_EXERCICE.Value, HB_DATE1, HB_DATE2) ;
DATE1.ItemIndex:=0 ;
DATE2.ItemIndex:=DATE2.Items.Count-1 ;
end;

procedure TFMulBds.DATE1Change(Sender: TObject);
begin
inherited ;
HB_DATE1.Text:=DATE1.Values[DATE1.ItemIndex] ;
end;

procedure TFMulBds.DATE2Change(Sender: TObject);
begin
inherited ;
HB_DATE2.Text:=DATE2.Values[DATE2.ItemIndex] ;
end;

// Modification de la balance sélectionnée
procedure TFMulBds.FListeDblClick(Sender: TObject);
var Params : RParBal ;
begin
if (Q.Eof) and (Q.Bof) then begin PrintMessageRC(RC_NOBALANCE) ; Exit ; end ;
inherited ;
FillChar(Params, Sizeof(Params), #0) ;
Params.ParExercice:=Q.FindField('HB_EXERCICE').AsString ;
Params.ParDebPeriode:=DateToStr(Q.FindField('HB_DATE1').AsDateTime) ;
Params.ParFinPeriode:=DateToStr(Q.FindField('HB_DATE2').AsDateTime) ;
//if Parent is THPanel then CloseInsidePanel(Self) ;
ChargeSaisieBalance(Params, taModif) ;
end;

procedure TFMulBds.BDeleteClick(Sender: TObject);
var (*SaveBookmark : TBookmark ;*) i : Integer ;
begin
if (FListe.NbSelected=0) and (not FListe.AllSelected) then
  begin PrintMessageRC(RC_NOBALANCE) ; Exit ; end ;
if PrintMessageRC(RC_CONFIRMDEL)<>mrYes then Exit ;
inherited ;
//SaveBookmark:=Q.GetBookmark ;
if FListe.AllSelected then
  begin
  InitMove(Q.RecordCount, '') ; Q.First ;
  while not Q.EOF do
    begin
    MoveCur(FALSE) ;
    CreateBal ;
    if Transactions(DelBalance, 3)<>oeOk then
      begin FreeBal ; PrintMessageRC(RC_BADDEL) ; Break ; end ;
    Q.Next;
    end ;
  FListe.AllSelected:=FALSE ;
  end else
  begin
  InitMove(FListe.NbSelected, '') ;
  for i:=0 to FListe.NbSelected-1 do
    begin
    MoveCur(FALSE) ;
    FListe.GotoLeBookmark(i) ;
    CreateBal ;
    if Transactions(DelBalance, 3)<>oeOk then
      begin FreeBal ; PrintMessageRC(RC_BADDEL) ; Break ; end ;
    end ;
  FListe.ClearSelected ;
  end ;
FiniMove ; BChercheClick(Sender) ;
//Q.GotoBookmark(SaveBookmark) ; Q.FreeBookmark(SaveBookmark) ;
end;

procedure TFMulBds.BChercheClick(Sender: TObject);
begin
inherited ;
EnableButtons ;
end;

procedure TFMulBds.BinsertClick(Sender: TObject);
var Params : RParBal ;
begin
inherited ;
FillChar(Params, Sizeof(Params), #0) ;
Params.ParExercice:=HB_EXERCICE.Value ;
Params.ParDebPeriode:=HB_DATE1.Text ;
Params.ParFinPeriode:=HB_DATE2.Text ;
//if Parent is THPanel then CloseInsidePanel(Self) ;
ChargeSaisieBalance(Params, taCreat) ;
end;

procedure TFMulBds.BNouvRechClick(Sender: TObject);
begin
DATE1.ItemIndex:=0 ;
DATE2.ItemIndex:=DATE2.Items.Count-1 ;
end;

procedure TFMulBds.BOuvrirClick(Sender: TObject);
var Params : RParBal ;
begin
inherited ;
if (Q.Eof) and (Q.Bof) then begin PrintMessageRC(RC_NOBALANCE) ; Exit ; end ;
FillChar(Params, Sizeof(Params), #0) ;
Params.ParExercice:=Q.FindField('HB_EXERCICE').AsString ;
Params.ParDebPeriode:=DateToStr(Q.FindField('HB_DATE1').AsDateTime) ;
Params.ParFinPeriode:=DateToStr(Q.FindField('HB_DATE2').AsDateTime) ;
//if Parent is THPanel then CloseInsidePanel(Self) ;
ChargeSaisieBalance(Params, taModif) ;
end;

end.
