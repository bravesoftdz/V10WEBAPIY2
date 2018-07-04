unit VerLettr;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls,HCtrls, Ent1, Hent1, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  SaisUtil, LettUtil, Hcompte,HStatus,
  Buttons, ExtCtrls, hmsgbox, HSysMenu, ed_tools,RappType ;

type
  TFVerLettr = class(TForm)
    QLett: TQuery;
    E_GENERAL: THCpteEdit;
    HLabel5: THLabel;
    E_GENERAL_: THCpteEdit;
    E_AUXILIAIRE_: THCpteEdit;
    HLabel2: THLabel;
    E_AUXILIAIRE: THCpteEdit;
    HLabel1: THLabel;
    HLabel4: THLabel;
    T_NATUREAUXI: THValComboBox;
    Label14: THLabel;
    HPB: TPanel;
    BAide: THBitBtn;
    BValider: THBitBtn;
    BFerme: THBitBtn;
    HMTrad: THSystemMenu;
    MsgBox: THMsgBox;
    procedure FormCreate(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    TDev : TList ;
    Procedure RepareLettrage ;
    function  ChargeLettrage(Gene,Auxi : String17 ; Lettrage : String) : TList ;
    function  ChargeTLRappro : TL_Rappro ;
    function  ChargeDevises : TList ;

  public
    { Déclarations publiques }
  end;

procedure RepareLeLettrage ;

implementation

uses LetBatch ;

{$R *.DFM}

procedure RepareLeLettrage ;
var FVerLettr : TFVerLettr ;
BEGIN
FVerLettr:=TFVerLettr.Create(Application) ;
try
 FVerLettr.ShowModal ;
 finally
 FVerLettr.Free ;
 end ;
SourisNormale ;
END ;

Procedure TFVerLettr.RepareLettrage ;
var Q : TQuery ;
    TLett:TList ;
    St : String ;
BEGIN ;
TDev:=ChargeDevises ;
St:='SELECT DISTINCT E_GENERAL,E_AUXILIAIRE,E_LETTRAGE FROM ECRITURE LEFT JOIN TIERS ON T_AUXILIAIRE=E_AUXILIAIRE WHERE E_ECHE="X" AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU="N"' ;
if E_GENERAL.Text<>'' then St:=St+' AND E_GENERAL>="'+E_GENERAL.Text+'"' ;
if E_GENERAL_.Text<>'' then St:=St+' AND E_GENERAL<="'+E_GENERAL_.Text+'"' ;
if E_AUXILIAIRE.Text<>'' then St:=St+' AND E_AUXILIAIRE>="'+E_AUXILIAIRE.Text+'"' ;
if E_AUXILIAIRE_.Text<>'' then St:=St+' AND E_AUXILIAIRE<="'+E_AUXILIAIRE_.Text+'"' ;
if T_NATUREAUXI.Value<>'' then St:=St+' AND T_NATUREAUXI="'+T_NATUREAUXI.Value+'"' ;
Q:=OpenSQL(St,True) ;
InitMove(RecordsCount(Q),'') ;
TLett:=TList.Create ;
While (not Q.Eof) do
  BEGIN
  VideListe(TLett) ;
  TLett:=ChargeLettrage(Q.FindField('E_GENERAL').AsString,Q.FindField('E_AUXILIAIRE').AsString,Q.FindField('E_LETTRAGE').AsString) ;
  LettrerUnPaquet(TLett,True,False) ;
  Q.Next ;
  MoveCur(False) ;
  END ;
VideListe(TLett) ; TLett.Free ; Ferme(Q) ;
FiniMove ;
END ;

function TFVerLettr.ChargeLettrage(Gene,Auxi : String17 ; Lettrage : String) : TList ;
var T : TList ;
    St,StV8 : String ;
BEGIN
St:='UPDATE ECRITURE SET E_COUVERTURE=0,E_COUVERTUREDEV=0,E_COUVERTUREEURO=0 WHERE'
   +' E_GENERAL="'+Gene+'" AND E_AUXILIAIRE="'+Auxi+'" AND E_LETTRAGE="'+Lettrage+'"' ;
StV8:=LWhereV8 ; if StV8<>'' then St:=St+' AND '+StV8 ;
ExecuteSQL(St) ;
QLett.Params[0].AsString:=Auxi ;
QLett.Params[1].AsString:=Gene ;
QLett.Params[2].AsString:=Lettrage ;
QLett.Open ;
T:=TList.Create ;
BEGINTRANS ;
While not QLett.Eof do
  BEGIN
  T.Add(ChargeTLRappro) ;
  QLett.Next ;
  END ;
COMMITTRANS ;
Result:=T ;
QLett.Close ;
END ;

function TFVerLettr.ChargeTLRappro : TL_Rappro ;
var L      : TL_Rappro ;
    DecDev : Integer ;
    Quotite : Double ;
BEGIN
DecDev:=V_PGI.OkDecV ;
L:=TL_Rappro.Create ;
RecupDevise(V_PGI.DevisePivot,DecDev,Quotite,TDev) ;
L.Jal:=QLett.FindField('E_JOURNAL').AsString ;
L.Exo:=QLett.FindField('E_EXERCICE').AsString ;
L.DateC:=QLett.FindField('E_DATECOMPTABLE').AsDateTime ;
L.Numero:=QLett.FindField('E_NUMEROPIECE').AsInteger;
L.Nature:=QLett.FindField('E_NATUREPIECE').AsString ;
L.NumLigne:=QLett.FindField('E_NUMLIGNE').AsInteger ;
L.NumEche:=QLett.FindField('E_NUMECHE').AsInteger ;
L.General:=QLett.FindField('E_GENERAL').AsString ;
L.Auxiliaire:=QLett.FindField('E_AUXILIAIRE').AsString ;
L.DateR:=QLett.FindField('E_DATEREFEXTERNE').AsDateTime ;
L.RefI:=QLett.FindField('E_REFINTERNE').AsString ; L.RefL:=QLett.FindField('E_REFLIBRE').AsString ;
L.RefE:=QLett.FindField('E_REFEXTERNE').AsString ; L.Lib:=QLett.FindField('E_LIBELLE').AsString ;
L.TauxDEV:=QLett.FindField('E_TAUXDEV').AsFloat ;
L.Debit:=QLett.FindField('E_DEBIT').AsFloat ; L.Credit:=QLett.FindField('E_CREDIT').AsFloat ;
L.DebDev:=QLett.FindField('E_DEBITDEV').AsFloat ; L.CredDev:=QLett.FindField('E_CREDITDEV').AsFloat ;
L.CodeL:=QLett.FindField('E_LETTRAGE').AsString ;
L.DateE:=QLett.FindField('E_DATEECHEANCE').AsDateTime ;
L.CodeD:=QLett.FindField('E_DEVISE').AsString ;
L.Decim:=DecDev ;
L.Facture:=((L.Nature='FC') or (L.Nature='FF') or (L.Nature='AC') or (L.Nature='AF')) ;
L.Client:=((L.Nature='FC') or (L.Nature='AC') or (L.Nature='RC') or (L.Nature='OC')) ;
L.Solution:=0 ;
L.EditeEtatTva:=(QLett.FindField('E_EDITEETATTVA').AsString='X') ;
Result:=L ;
END ;

Function TFVerLettr.ChargeDevises : TList ;
var TDev : TList ;
    Q,Q1 : TQuery ;
    TDevise : TFDevise ;
BEGIN
Q:=OpenSQL('SELECT D_DEVISE,D_QUOTITE,D_DECIMALE FROM DEVISE ORDER BY D_DEVISE',True) ;
Q1:=TQuery.Create(Application) ; Q1.DatabaseName:='SOC' ;
Q1.SQL.Add('SELECT H_TAUXREEL FROM CHANCELL WHERE H_DEVISE=:DEV'
          +' AND H_DATECOURS=(SELECT MAX(H_DATECOURS) FROM CHANCELL WHERE H_DEVISE=:DEV)') ;
ChangeSQL(Q1) ;//Q1.Prepare ;
PrepareSQLODBC(Q1) ;
TDev:=TList.Create ;
While not Q.Eof do
  BEGIN
  TDevise:=TFDevise.Create ;
  TDevise.Code:=Q.Fields[0].AsString ;
  TDevise.Quotite:=Q.Fields[1].AsFloat ;
  TDevise.Decimale:=Q.Fields[2].AsInteger ;
  Q1.Params[0].AsString:=Q.Fields[0].AsString ; ;
  Q1.Open ; TDevise.TauxDev:=Q1.Fields[0].AsFloat ; Q1.Close ;
  TDev.Add(TDevise) ;
  Q.Next ;
  END ;
Ferme(Q) ; Q1.Free ;
Result:=TDev ;
END ;

procedure TFVerLettr.FormCreate(Sender: TObject);
var StV8 : String ;
begin
StV8:=LWhereV8 ; if StV8<>'' then QLett.SQL.Add(' AND '+StV8) ;
ChangeSQL(QLett) ; //QLett.Prepare ;
PrepareSQLODBC(QLett) ;
end;


procedure TFVerLettr.BValiderClick(Sender: TObject);
begin
if MsgBox.Execute(0,Caption,'')<>mrYes then Exit ;
SourisSablier ;
RepareLettrage ;
MsgBox.Execute(1,Caption,'') ;
SourisNormale ;
Close ;
end;

procedure TFVerLettr.FormShow(Sender: TObject);
begin T_NATUREAUXI.ItemIndex:=0 ; end;

end.
