unit TVAEdite;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, Hctrls, Mask, Buttons, ParamDat, Ent1, HEnt1,
  hmsgbox, HSysMenu,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DB;

procedure TVAEditeEtat ;

type
  TFTVAEdite = class(TForm)
    HPB: TPanel;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    Label7: TLabel;
    FDateCompta2: TMaskEdit;
    FDateCompta1: TMaskEdit;
    FExercice: THValComboBox;
    HLabel4: THLabel;
    HLabel6: THLabel;
    TypeTVA: TRadioGroup;
    HMTrad: THSystemMenu;
    MsgBox: THMsgBox;
    procedure TypeTVAClick(Sender: TObject);
    procedure FDateCompta1KeyPress(Sender: TObject; var Key: Char);
    procedure FDateCompta2KeyPress(Sender: TObject; var Key: Char);
    procedure FExerciceChange(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    procedure LanceEditeEtatTva ;
    function  SQLDefaut : String ;
    function  SQLFactureDirecte : String ;
    function  SQLFactureIndirecte : String ;
    procedure MajLignesTTC ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  HStatus;


const
  UpdateTmp = 'L' ;


procedure TVAEditeEtat ;
var X : TFTVAEdite ;
BEGIN
X:= TFTVAEdite.Create(Application) ;
try
 X.ShowModal ;
 finally
 X.Free ;
 end ;
SourisNormale ;
END;

procedure TFTVAEdite.TypeTVAClick(Sender: TObject);
begin
Caption:=MsgBox.Mess[TypeTva.ItemIndex] ;
UpdateCaption(Self) ;
end;

procedure TFTVAEdite.FDateCompta1KeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

procedure TFTVAEdite.FDateCompta2KeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

procedure TFTVAEdite.FExerciceChange(Sender: TObject);
begin ExoToDates(FExercice.Value,FDateCompta1,FDateCompta2) ; end;

procedure TFTVAEdite.BValiderClick(Sender: TObject);
begin LanceEditeEtatTva ;end;

function TFTVAEdite.SQLDefaut : String ;
var SQL : String ;
BEGIN
SQL:='UPDATE ECRITURE Q SET E_EDITEETATTVA="'+UpdateTmp+'" WHERE ' ;
if FExercice.ItemIndex>=0 then SQL:=SQL+' E_EXERCICE="'+FExercice.Value+'" AND ' ;
if TypeTva.ItemIndex=0 then SQL:=SQL+' E_DATECOMPTABLE>="'+USDateTime(StrToDate(FDateCompta1.Text))+'"'
                                    +' AND E_DATECOMPTABLE<="'+USDateTime(StrToDate(FDateCompta2.Text))+'"'
                       else SQL:=SQL+' E_DATEECHEANCE>="'+USDateTime(StrToDate(FDateCompta1.Text))+'"'
                                    +' AND E_DATEECHEANCE<="'+USDateTime(StrToDate(FDateCompta2.Text))+'"' ;
case TypeTva.ItemIndex of
  0 : SQL:=SQL+' AND E_GENERAL LIKE "'+VH^.CollCliEnc+'%"' ;
  1 : SQL:=SQL+' AND E_GENERAL LIKE "'+VH^.CollFouEnc+'%"' ;
  END ;
SQL:=SQL+' AND E_AUXILIAIRE IN(SELECT T_AUXILIAIRE FROM TIERS WHERE T_TVAENCAISSEMENT="TE")' ;
Result:=SQL ;
END ;

function SQLEcheEnc : String ;
var SQL : String ;
BEGIN
SQL:=SQL+' AND (E_ECHEENC1<>0 OR E_ECHEENC2<>0 '
        +'OR E_ECHEENC3<>0 OR E_ECHEENC4<>0) AND E_EDITEETATTVA="-" ' ;
Result:=SQL ;
END ;

function SQLMP : String ;
var SQL : String ;
BEGIN
SQL:=SQL+' AND E_MODEPAIE IN(SELECT MP_MODEPAIE FROM MODEPAIE WHERE (MP_CATEGORIE="CB") OR (MP_CATEGORIE="ESP") OR (MP_CATEGORIE="CHQ")'
        +' OR (MP_CATEGORIE="VIR") OR (MP_CATEGORIE="PRE"))' ;
Result:=SQL ;
END ;

function TFTVAEdite.SQLFactureDirecte : String ;
var Nat1,Nat2,NatR,SQL : String ;
BEGIN
if TypeTva.ItemIndex=0 then BEGIN Nat1:='FC' ; Nat2:='AC' ; NatR:='RC' ; END
                       else BEGIN Nat1:='FF' ; Nat2:='AF' ; NatR:='RF' ; END ;
SQL:=' AND E_LETTRAGE<>"" AND (((E_NATUREPIECE="'+NatR+'" OR E_NATUREPIECE="OD") AND (E_CONTREPARTIEGEN LIKE "512%"'
    +' OR E_CONTREPARTIEGEN LIKE "514%"))'
    +' OR (((E_NATUREPIECE="'+Nat1+'") OR (E_NATUREPIECE="'+Nat2+'")) '
    +' AND EXISTS(SELECT E_AUXILIAIRE FROM ECRITURE WHERE E_AUXILIAIRE=Q.E_AUXILIAIRE'
    +' AND (E_LETTRAGE=Q.E_LETTRAGE)'
    +' AND (E_NUMEROPIECE<>Q.E_NUMEROPIECE) AND ((E_NATUREPIECE="'+NatR+'")'
    +' OR (E_NATUREPIECE="OD")) AND (E_CONTREPARTIEGEN LIKE "512%" OR E_CONTREPARTIEGEN LIKE "514%"))))' ;
Result:=SQL ;
END ;

function TFTVAEdite.SQLFactureIndirecte : String ;
var Nat1,Nat2,NatR,SQL : String ;
BEGIN
if TypeTva.ItemIndex=0 then BEGIN Nat1:='FC' ; Nat2:='AC' ; NatR:='RC' ; END
                       else BEGIN Nat1:='FF' ; Nat2:='AF' ; NatR:='RF' ; END ;
SQL:=' AND E_LETTRAGE<>"" AND (((E_NATUREPIECE="'+NatR+'" OR E_NATUREPIECE="OD")'
    +' AND (E_CONTREPARTIEGEN NOT LIKE "512%" AND E_CONTREPARTIEGEN NOT LIKE "514%"))'
    +' OR ((E_NATUREPIECE<>"'+Nat1+'") AND (E_NATUREPIECE<>"'+Nat2+'") AND (E_NATUREPIECE<>"'+NatR+'")))'
    +' AND EXISTS(SELECT E_AUXILIAIRE FROM ECRITURE WHERE E_AUXILIAIRE=Q.E_AUXILIAIRE'
    +' AND (E_LETTRAGE=Q.E_LETTRAGE)'
    +' AND (E_NUMEROPIECE<>Q.E_NUMEROPIECE) AND (E_NATUREPIECE<>"'+Nat1+'")'
    +' AND (E_NATUREPIECE<>"'+Nat2+'") AND (E_NATUREPIECE<>"'+NatR+'"))' ;
Result:=SQL ;
END ;

procedure TFTVAEdite.MajLignesTTC ;
var SQL : String ;
BEGIN
BEGINTRANS ;
InitMove(3,'') ;
MoveCur(False) ;
//Défaut : Acomptes
SQL:=SQLDefaut ;
ExecuteSQL(SQL+SQLEcheEnc) ;
MoveCur(False) ;
//Factures directes
ExecuteSQL(SQL+SQLFactureDirecte+SQLMP) ;
MoveCur(False) ;
//Factures indirectes
ExecuteSQL(SQL+SQLFactureIndirecte+SQLMP) ;
FiniMove ;
COMMITTRANS ;
END ;

procedure MajPieces ;
var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE,'+
           ' E_QUALIFPIECE FROM ECRITURE'+
           ' WHERE E_EDITEETATTVA="'+UpdateTmp+'"',True) ;
BEGINTRANS ;
InitMove(QCount(Q),'') ;
While not Q.Eof do
  BEGIN
  ExecuteSQL('UPDATE ECRITURE SET E_EDITEETATTVA="X" WHERE E_JOURNAL="'+Q.Fields[0].AsString+'"'+
             ' AND E_EXERCICE="'+Q.Fields[1].AsString+'" AND E_DATECOMPTABLE="'+USDateTime(Q.Fields[2].AsDateTime)+'"'+
             ' AND E_NUMEROPIECE='+Q.Fields[3].AsString+' AND E_QUALIFPIECE="'+Q.Fields[4].AsString+'"') ;
  Q.Next ;
  MoveCur(False) ;
  END ;
Ferme(Q) ;
FiniMove ;
COMMITTRANS ;
END ;

procedure TFTVAEdite.LanceEditeEtatTva ;
BEGIN
case MsgBox.Execute(2,Caption,'') of
  mrNo     : Exit ;
  mrCancel : BEGIN ModalResult:=mrNone ; Exit ; END ;
  END ;
MajLignesTTC ;
MajPieces ;
MsgBox.Execute(3,Caption,'') ;
END ;

procedure TFTVAEdite.FormShow(Sender: TObject);
begin
FExercice.Value:=EXRF(VH^.Entree.Code) ;
end;

procedure TFTVAEdite.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
end;

end.
