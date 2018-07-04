{***********UNITE*************************************************
Auteur  ...... : Paie PGI
Créé le ...... : 15/06/2000
Modifié le ... :   /  /
Description .. : fonction de suppression d'une ligne dans un multicritère
Mots clefs ... : PAIE;MULTICRITERE
*****************************************************************}
{
PT1 20/06/2005 PH V_60 FQ 11813 Compatibilité CWAS
}
unit PGOperation;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,HTB97,Dialogs
{$IFNDEF EAGLCLIENT}
     ,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB, mul
{$ELSE}
      ,UTOB, emul
{$ENDIF}
      ,HStatus,HEnt1,HMsgBox,HCtrls;

Procedure  AppelSuppOpe (L,Q : TComponent;F : TForm);
var PGSuppression : string;
Type
      COperation = Class
      private
      TheQ : TQUERY;
      LaForme : TForm; // PT1
      procedure SuppressionOpe;
      public
{$IFNDEF EAGLCLIENT}
      procedure SupprimeOperation(L : THDBGrid; Q : TQuery ; F : TForm);
{$ELSE}
      procedure SupprimeOperation(L : THGrid; Q : TQuery ; F : TForm);
{$ENDIF}
      end;

implementation
{$IFNDEF EAGLCLIENT}
procedure COperation.SupprimeOperation(L : THDBGrid; Q : TQuery; F : TForm);
{$ELSE}
procedure COperation.SupprimeOperation(L : THGrid; Q : TQuery; F : TForm);
{$ENDIF}
var i : integer;
Btn : TToolBarButton97;
begin
TheQ := Q;
if (L.NbSelected=0) and (not L.AllSelected) then
   begin
   MessageAlerte('Aucun élément sélectionné');
   exit;
   end;
if MessageDlg('Vous allez supprimer définitivement les informations. Confirmez vous l''opération ?',mtConfirmation,[mbYes,mbNo],0)<>mrYes then exit ;
if L.AllSelected then
   BEGIN
   InitMove(Q.RecordCount,'');
   Q.First;
   while Not Q.EOF do
      BEGIN
      MoveCur(False);
      if Transactions(SuppressionOpe, 3)<>oeOk then
         BEGIN
         MessageAlerte('Suppression impossible') ;
         Break ;
         END ;
      Q.Next;
      END;
   L.AllSelected:=False;
   END ELSE
   BEGIN
   InitMove(L.NbSelected,'');
   for i:=0 to L.NbSelected-1 do
      BEGIN
      MoveCur(False);
      L.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
      TFmul(F).Q.TQ.Seek(L.Row - 1); // PT1
       // Q.Seek(L.Row - 1);
{$ENDIF}
      LaForme := TFmul(F);
      if Transactions(SuppressionOpe, 3)<>oeOk then
         BEGIN
         MessageAlerte('Suppression impossible') ;
         Break ;
         END ;
      END;
   L.ClearSelected;
   END;
FiniMove;
Btn:=TToolBarButton97(F.FindComponent('Bcherche')) ;
if Btn<>nil then Btn.click;
end;

procedure COperation.SuppressionOpe ();
VAR
DateDeb,DateFin : TDateTime;
begin
if PGSuppression='VIR' then
  Begin
  {
  DateDeb:=StrToDate(TheQ.FindField('PVI_DATEDEBUT').asstring);
  DateFin:=StrToDate(TheQ.FindField('PVI_DATEFIN').asstring);
  }
  // DEB PT1
  DateDeb:= TFmul(LaForme).Q.FindField('PVI_DATEDEBUT').AsDateTime ;
  DateFin:=TFmul(LaForme).Q.FindField('PVI_DATEFIN').AsDateTime;

  if (Datedeb>0) and (DateFin>0) then
     ExecuteSQL('DELETE FROM VIREMENTS '+
     'WHERE PVI_SALARIE="'+TFmul(LaForme).Q.FindField('PVI_SALARIE').asstring+'" '+
     'AND PVI_DATEDEBUT="'+UsDateTime(DateDeb)+'" '+
     'AND PVI_DATEFIN="'+UsDateTime(DateFin)+'"') ;
  End;
if PGSuppression='ACP' then
  Begin
  DateDeb:=StrToDate(TFmul(LaForme).Q.FindField('PSD_DATEDEBUT').asstring);
  DateFin:=StrToDate(TFmul(LaForme).Q.FindField('PSD_DATEFIN').asstring);
  if (Datedeb>0) and (DateFin>0) then
     ExecuteSQL('DELETE FROM HISTOSAISRUB '+
     'WHERE PSD_SALARIE="'+TFmul(LaForme).Q.FindField('PSD_SALARIE').asstring+'" '+
     'AND PSD_DATEDEBUT="'+UsDateTime(DateDeb)+'" '+
     'AND PSD_DATEFIN="'+UsDateTime(DateFin)+'" AND PSD_ORIGINEMVT="ACP"') ;
  End;
  // FIN PT1
end;

Procedure  AppelSuppOpe (L,Q : TComponent; F : TForm);
var X : COperation;
begin
X := COperation.create;
{$IFNDEF EAGLCLIENT}
X.SupprimeOperation (THDBGrid (L),TQuery (Q),F);
{$ELSE}
X.SupprimeOperation (THGrid (L),TQuery (Q),F);
{$ENDIF}
X.Free;
end;

end.
