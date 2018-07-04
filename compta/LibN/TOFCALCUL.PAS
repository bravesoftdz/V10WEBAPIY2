{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 23/01/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : RECALCUL ()
Mots clefs ... : TOF;RECALCUL
*****************************************************************}
Unit TOFCALCUL ;

Interface

Uses Windows, StdCtrls, Controls, Classes, db, forms, sysutils,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, TofMeth, Ent1, Vierge, FE_Main;

Type
  TOF_CALCUL = Class (TOF_Meth)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    Jal1, Jal2, Date1, Date2, Piece1, Piece2 : THEdit;
    Exo : THValComboBox ;
    DateD, DateF : TDatetime;
    procedure ExoOnChange(Sender: TObject) ;
    procedure DateOnExit(Sender: TObject) ;
   // function  Reverse(p: Integer):TDateTime ;
    procedure OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  end ;

Implementation

procedure TOF_CALCUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CALCUL.OnDelete ;
begin
  Inherited ;
end ;
(*
function TOF_CALCUL.Reverse(p: Integer):TDateTime ;
var z, zz: Integer ;
begin
z:=Trunc(p/100) ;
zz:=p-(z*100) ;
Result:=EncodeDate(z, zz, 1) ;
end ;
*)

procedure TOF_CALCUL.OnUpdate ;
//var Q, QEcr: TQuery ; Folio: TFolioU ; sQ, sQEcr: string ;
var sParams : string ;
begin
inherited ;
sParams:=Jal1.Text+';'+Jal2.Text+';'+Exo.Values[Exo.ItemIndex]+';' ;
sParams:=sParams+Date1.Text+';'+Date2.Text+';' ;
if ComboEtab.ItemIndex>0 then sParams:=sParams+ComboEtab.Values[ComboEtab.ItemIndex] ;
sParams:=sParams+';' ;
sParams:=sParams+Piece1.Text+';'+Piece2.Text+';' ;
AGLLanceFiche('CP', 'ECMVTVISU', '', '', sParams) ;
(*
sQ:='SELECT J_JOURNAL, J_MODESAISIE FROM JOURNAL WHERE J_MODESAISIE<>"-"' ;
if Jal1.Text<>'' then sQ:=sQ+' AND J_JOURNAL>="'+Jal1.Text+'"' ;
if Jal2.Text<>'' then sQ:=sQ+' AND J_JOURNAL<="'+Jal2.Text+'"' ;
sQ:=sQ+' ORDER BY J_JOURNAL' ;
Q:=OpenSQL(sQ, TRUE) ;
InitMove(20, '');
while not Q.EOF do
  begin
  MoveCur(FALSE) ;
  sQEcr:='SELECT DISTINCT E_PERIODE, E_NUMEROPIECE FROM ECRITURE WHERE E_JOURNAL="'+Q.FindField('J_JOURNAL').AsString+'"' ;
  sQEcr:=sQEcr+' AND E_EXERCICE="'+Exo.Values[Exo.ItemIndex]+'"' ;
  sQEcr:=sQEcr+' AND E_DATECOMPTABLE>="'+UsDateTime(StrToDate(Date1.Text))+'"' ;
  sQEcr:=sQEcr+' AND E_DATECOMPTABLE<="'+UsDateTime(StrToDate(Date2.Text))+'"' ;
  if Etabl.ItemIndex>0 then
    sQEcr:=sQEcr+' AND E_ETABLISSEMENT="'+Etabl.Values[Etabl.ItemIndex]+'"' ;
  if StrToInt(Piece1.Text)>0 then
    sQEcr:=sQEcr+' AND E_NUMEROPIECE>='+Piece1.Text ;
  if StrToInt(Piece2.Text)<999999999 then
    sQEcr:=sQEcr+' AND E_NUMEROPIECE<='+Piece2.Text ;
  sQEcr:=sQEcr+' ORDER BY E_PERIODE, E_NUMEROPIECE' ;
  QEcr:=OpenSQL(sQEcr, TRUE) ;
  while not QEcr.EOF do
    begin
    Folio:=TFolioU.Create(IntToStr(QEcr.FindField('E_NUMEROPIECE').AsInteger),
                          Q.FindField('J_JOURNAL').AsString,
                          Exo.Values[Exo.ItemIndex],
                          FindeMois(Reverse(QEcr.FindField('E_PERIODE').AsInteger)),
                          Q.FindField('J_MODESAISIE').AsString='LIB') ;
    Folio.Read ;
    Folio.VerifFolio ;
  //  Folio.Write ;
    Folio.Free ;
    QEcr.Next ;
    end ;
  Ferme(QEcr) ;
  Q.Next ;
  end ;
Ferme(Q) ;
FiniMove ;
PgiBox('Le traitement est terminé.', TFVierge(Ecran).Caption);
*)
end ;

procedure TOF_CALCUL.OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if key=VK_F10 then OnUpdate ;
end ;

procedure TOF_CALCUL.OnLoad ;
begin
inherited ;
end ;

procedure TOF_CALCUL.OnArgument (S : String ) ;
begin
inherited ;
Jal1:=THEdit(GetControl('E_JOURNAL')) ;
Jal2:=THEdit(GetControl('E_JOURNAL_')) ;
Exo:=THValComboBox(GetControl('E_EXERCICE')) ;
Date1:=THEdit(GetControl('E_DATECOMPTABLE')) ;
Date2:=THEdit(GetControl('E_DATECOMPTABLE_')) ;
Piece1:=THEdit(GetControl('E_NUMEROPIECE')) ;
Piece2:=THEdit(GetControl('E_NUMEROPIECE_')) ;
if Exo<>nil      then begin Exo.OnChange:=ExoOnChange ; Exo.Value:=VH^.Entree.Code ; end ;
if Date1<>nil    then Date1.OnExit:=DateOnExit ;
if Date2<>nil    then Date2.OnExit:=DateOnExit ;
if Ecran<>nil    then TFVierge(Ecran).OnKeyDown:=OnKeyDown ;
end ;

procedure TOF_CALCUL.ExoOnChange(Sender: TObject) ;
begin
DoExoToDateOnChange(Exo, Date1, Date2) ;
DateD:=StrToDate(Date1.Text) ;
DateF:=StrToDate(Date2.Text) ;
end ;

procedure TOF_CALCUL.DateOnExit(Sender: TObject) ;
begin
DoDateOnExit(THEdit(Sender), Date1, Date2, DateD, DateF) ;
end ;

procedure TOF_CALCUL.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_CALCUL ] ) ;
end.
