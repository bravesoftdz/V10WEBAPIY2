{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 24/01/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : MVTVISU ()
Mots clefs ... : TOF;MVTVISU
*****************************************************************}
Unit TOFMVTVISU ;

Interface

Uses Windows, StdCtrls, Controls, Classes, db, forms, sysutils,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Messages, ComCtrls, HCtrls, ZFolio, ZFolioU, HEnt1, HMsgBox, SaisBor,
     UTOF, HStatus, Vierge, PrintDbg;

Type
  TOF_MVTVISU = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    Jal1, Jal2, Exo, Date1, Date2, Etabl, Piece1, Piece2 : string ;
    FListe : THGrid ;
    BPrint, BZoom : TToolButton;
    function  Reverse(p: Integer):TDateTime ;
    procedure BPrintClick(Sender: TObject);
    procedure BZoomClick(Sender: TObject);
    procedure Go(Cas : Integer) ;
  end ;

Implementation

procedure TOF_MVTVISU.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_MVTVISU.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_MVTVISU.OnUpdate ;
begin
  Inherited ;
end ;

function TOF_MVTVISU.Reverse(p: Integer):TDateTime ;
var z, zz: Integer ;
begin
z:=Trunc(p/100) ;
zz:=p-(z*100) ;
Result:=EncodeDate(z, zz, 1) ;
end ;

procedure TOF_MVTVISU.Go(Cas : Integer) ;
var Q, QEcr: TQuery ; Folio: TFolioU ; sQ, sQEcr: string ; bToWrite : Boolean ;
begin
sQ:='SELECT J_JOURNAL, J_LIBELLE, J_MODESAISIE FROM JOURNAL WHERE J_MODESAISIE<>"-"' ;
if Jal1<>'' then sQ:=sQ+' AND J_JOURNAL>="'+Jal1+'"' ;
if Jal2<>'' then sQ:=sQ+' AND J_JOURNAL<="'+Jal2+'"' ;
sQ:=sQ+' ORDER BY J_JOURNAL' ;
Q:=OpenSQL(sQ, TRUE) ;
while not Q.EOF do
  begin
  MoveCur(FALSE) ;
  sQEcr:='SELECT DISTINCT E_PERIODE, E_NUMEROPIECE FROM ECRITURE WHERE E_JOURNAL="'+Q.FindField('J_JOURNAL').AsString+'"' ;
  sQEcr:=sQEcr+' AND E_EXERCICE="'+Exo+'"' ;
  sQEcr:=sQEcr+' AND E_DATECOMPTABLE>="'+UsDateTime(StrToDate(Date1))+'"' ;
  sQEcr:=sQEcr+' AND E_DATECOMPTABLE<="'+UsDateTime(StrToDate(Date2))+'"' ;
  if Etabl<>'' then
    sQEcr:=sQEcr+' AND E_ETABLISSEMENT="'+Etabl+'"' ;
  if StrToInt(Piece1)>0 then
    sQEcr:=sQEcr+' AND E_NUMEROPIECE>='+Piece1 ;
  if StrToInt(Piece2)<999999999 then
    sQEcr:=sQEcr+' AND E_NUMEROPIECE<='+Piece2 ;
  sQEcr:=sQEcr+' ORDER BY E_PERIODE, E_NUMEROPIECE' ;
  QEcr:=OpenSQL(sQEcr, TRUE) ;
  while not QEcr.EOF do
    begin
    bToWrite:=FALSE ;
    Folio:=TFolioU.Create(IntToStr(QEcr.FindField('E_NUMEROPIECE').AsInteger),
                          Q.FindField('J_JOURNAL').AsString,
                          Exo,
                          FindeMois(Reverse(QEcr.FindField('E_PERIODE').AsInteger)),
                          Q.FindField('J_MODESAISIE').AsString='LIB') ;
    Folio.Read ;
    if (Cas=3) and (Folio.VerifNumLigne) then
      begin
      FListe.Cells[5, Fliste.RowCount-1]:=TraduireMemoire('Correction numéro de ligne') ;
      bToWrite:=TRUE ;
      end ;
  {  if (Cas=1) and (Folio.VerifOppose) then
      begin
      FListe.Cells[5, Fliste.RowCount-1]:=TraduireMemoire('Correction contre valeur') ;
      bToWrite:=TRUE ;
      end ;   }
  {  if (Cas=2) and (Folio.VerifFolio)  then
      begin
      FListe.Cells[5, Fliste.RowCount-1]:=TraduireMemoire('Correction équilibre') ;
      bToWrite:=TRUE ;
      end ;    }
    if bToWrite then
      begin
      FListe.Cells[0, Fliste.RowCount-1]:=Q.FindField('J_JOURNAL').AsString ;
      FListe.Cells[1, Fliste.RowCount-1]:=QEcr.FindField('E_PERIODE').AsString ;
      FListe.Cells[2, Fliste.RowCount-1]:=Q.FindField('J_LIBELLE').AsString ;
      FListe.Cells[3, Fliste.RowCount-1]:=QEcr.FindField('E_NUMEROPIECE').AsString ;
      FListe.Cells[4, Fliste.RowCount-1]:=FormatDateTime('mmmm yyyy', FindeMois(Reverse(QEcr.FindField('E_PERIODE').AsInteger))) ;
      FListe.RowCount:=FListe.RowCount+1 ;
      Folio.Write ;
      end ;
    Folio.Free ;
    QEcr.Next ;
    end ;
  Ferme(QEcr) ;
  Q.Next ;
  end ;
Ferme(Q) ;
FListe.RowCount:=FListe.RowCount-1 ;
end ;

procedure TOF_MVTVISU.OnLoad ;
const MAXFUNC = 3 ;
const ORDFUNC : array[1..MAXFUNC] of integer = (3, 1, 2) ;
var Q : TQuery ; sQ : string ; i : integer ;
begin
inherited ;
sQ:='SELECT COUNT(*) FROM JOURNAL WHERE J_MODESAISIE<>"-"' ;
if Jal1<>'' then sQ:=sQ+' AND J_JOURNAL>="'+Jal1+'"' ;
if Jal2<>'' then sQ:=sQ+' AND J_JOURNAL<="'+Jal2+'"' ;
sQ:=sQ+' ORDER BY J_JOURNAL' ;
Q:=OpenSQL(sQ, TRUE) ;
InitMove(Q.Fields[0].AsInteger*3, '') ;
Ferme(Q) ;
for i:=0 to MAXFUNC do Go(ORDFUNC[i]) ;
FiniMove ;
//PgiBox('Le traitement est terminé.', TFVierge(Ecran).Caption);
end ;

procedure TOF_MVTVISU.OnArgument (S : String ) ;
begin
inherited ;
FListe:=THGrid(GetControl('FLISTE')) ;
FListe.ColWidths[0]:=0 ;
FListe.ColWidths[1]:=0 ;
FListe.ColAligns[3]:=taCenter ;
FListe.ColAligns[4]:=taCenter ;
BPrint:=TToolButton(GetControl('BIMPRIMER'));
BZoom:=TToolButton(GetControl('BZOOM'));
if not Assigned(FListe.OnDblClick) then FListe.OnDblClick:=BZoomClick ;
if BPrint<>nil then BPrint.OnClick:=BPrintClick ;
if BZoom<>nil  then  BZoom.OnClick:=BZoomClick ;
Jal1:=ReadTokenSt(s) ;
Jal2:=ReadTokenSt(s) ;
Exo:=ReadTokenSt(s) ;
Date1:=ReadTokenSt(s) ;
Date2:=ReadTokenSt(s) ;
Etabl:=ReadTokenSt(s) ;
Piece1:=ReadTokenSt(s) ;
Piece2:=ReadTokenSt(s) ;
end ;

procedure TOF_MVTVISU.BPrintClick(Sender: TObject);
begin
PrintDBGrid(FListe, nil, TFVierge(Ecran).Caption, '') ;
end ;

procedure TOF_MVTVISU.BZoomClick(Sender: TObject);
var P : RParFolio ;
begin
FillChar(P, Sizeof(P), #0) ;
P.ParPeriode:=DateToStr(Reverse(StrToInt(FListe.Cells[1, FListe.Row])));
P.ParCodeJal:=FListe.Cells[0, FListe.Row] ;
P.ParNumFolio:=FListe.Cells[3, FListe.Row] ;
P.ParNumLigne:=1 ;
ChargeSaisieFolio(P, taConsult) ;
end ;

procedure TOF_MVTVISU.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_MVTVISU ] ) ; 
end.
